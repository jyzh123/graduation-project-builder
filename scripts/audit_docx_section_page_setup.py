from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def attrs(node: ET.Element | None) -> dict[str, str]:
    if node is None:
        return {}
    return {key.rsplit("}", 1)[-1]: value for key, value in node.attrib.items()}


def load_document_root(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def section_layout(sect_pr: ET.Element) -> dict[str, dict[str, str]]:
    return {
        name: attrs(sect_pr.find(f"w:{name}", NS))
        for name in ("pgSz", "pgMar", "cols", "docGrid")
        if sect_pr.find(f"w:{name}", NS) is not None
    }


def first_reference_layout(reference_docx: Path) -> tuple[int, dict[str, dict[str, str]]]:
    root = load_document_root(reference_docx)
    for index, sect_pr in enumerate(root.findall(".//w:sectPr", NS), start=1):
        layout = section_layout(sect_pr)
        if "pgSz" in layout and "pgMar" in layout:
            return index, layout
    raise ValueError(f"reference DOCX has no section with pgSz and pgMar: {reference_docx}")


def compare_attrs(
    *,
    section_index: int,
    node_name: str,
    expected: dict[str, str],
    actual: dict[str, str],
) -> list[str]:
    issues: list[str] = []
    for key, expected_value in expected.items():
        actual_value = actual.get(key)
        if actual_value != expected_value:
            issues.append(
                f"section {section_index} {node_name}.{key} drift: expected {expected_value}, actual {actual_value}"
            )
    return issues


def parse_pdf_media_boxes(pdf_path: Path) -> list[dict[str, float]]:
    text = pdf_path.read_bytes().decode("latin-1", errors="ignore")
    pattern = re.compile(
        r"/MediaBox\s*\[\s*"
        r"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+"
        r"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*\]"
    )
    boxes: list[dict[str, float]] = []
    for match in pattern.finditer(text):
        x0, y0, x1, y1 = [float(item) for item in match.groups()]
        boxes.append(
            {
                "x0": x0,
                "y0": y0,
                "x1": x1,
                "y1": y1,
                "width": abs(x1 - x0),
                "height": abs(y1 - y0),
            }
        )
    return boxes


def audit_docx_sections(
    *,
    final_docx: Path,
    reference_docx: Path | None,
    require_cols_docgrid: bool,
    compare_reference: bool,
) -> tuple[dict[str, object], list[str], dict[str, dict[str, str]] | None]:
    reference_layout: dict[str, dict[str, str]] | None = None
    reference_section_index: int | None = None
    if reference_docx is not None:
        reference_section_index, reference_layout = first_reference_layout(reference_docx)

    root = load_document_root(final_docx)
    sections = root.findall(".//w:sectPr", NS)
    issues: list[str] = []
    records: list[dict[str, object]] = []
    required_nodes = ["pgSz", "pgMar"]
    if require_cols_docgrid:
        required_nodes.extend(["cols", "docGrid"])

    if not sections:
        issues.append("final DOCX has no sectPr nodes")

    for index, sect_pr in enumerate(sections, start=1):
        layout = section_layout(sect_pr)
        missing = [name for name in required_nodes if name not in layout]
        if missing:
            issues.append(f"section {index} missing required page setup: {', '.join(missing)}")
        if compare_reference and reference_layout is not None:
            for name, expected in reference_layout.items():
                if name not in required_nodes and name not in ("pgSz", "pgMar"):
                    continue
                actual = layout.get(name, {})
                if not actual:
                    continue
                issues.extend(compare_attrs(section_index=index, node_name=name, expected=expected, actual=actual))
        records.append({"section_index": index, "layout": layout, "missing_required_nodes": missing})

    report = {
        "schema": "graduation-project-builder.docx-section-page-setup-audit.v1",
        "final_docx_path": str(final_docx),
        "final_docx_sha256": sha256(final_docx),
        "reference_docx_path": str(reference_docx) if reference_docx is not None else None,
        "reference_docx_sha256": sha256(reference_docx) if reference_docx is not None else None,
        "reference_section_index": reference_section_index,
        "require_cols_docgrid": require_cols_docgrid,
        "compare_reference": compare_reference,
        "section_count": len(sections),
        "sections": records,
    }
    return report, issues, reference_layout


def audit_pdf_page_boxes(
    *,
    final_pdf: Path,
    reference_layout: dict[str, dict[str, str]] | None,
    tolerance_points: float,
) -> tuple[dict[str, object], list[str]]:
    boxes = parse_pdf_media_boxes(final_pdf)
    issues: list[str] = []
    expected_width = None
    expected_height = None
    if reference_layout is not None and "pgSz" in reference_layout:
        pg_sz = reference_layout["pgSz"]
        if "w" in pg_sz and "h" in pg_sz:
            expected_width = int(pg_sz["w"]) / 20.0
            expected_height = int(pg_sz["h"]) / 20.0
    if not boxes:
        issues.append(f"PDF MediaBox not found: {final_pdf}")
    elif expected_width is not None and expected_height is not None:
        for index, box in enumerate(boxes, start=1):
            width_delta = abs(float(box["width"]) - expected_width)
            height_delta = abs(float(box["height"]) - expected_height)
            if width_delta > tolerance_points or height_delta > tolerance_points:
                issues.append(
                    "PDF page box drift: "
                    f"page {index} expected {expected_width:.2f}x{expected_height:.2f} pt, "
                    f"actual {box['width']:.2f}x{box['height']:.2f} pt"
                )
                break
    return {
        "final_pdf_path": str(final_pdf),
        "final_pdf_sha256": sha256(final_pdf),
        "expected_width_points": expected_width,
        "expected_height_points": expected_height,
        "tolerance_points": tolerance_points,
        "media_box_count": len(boxes),
        "media_boxes": boxes[:20],
    }, issues


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit DOCX section page setup and optional PDF page boxes.")
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--reference-docx")
    parser.add_argument("--final-pdf")
    parser.add_argument("--report-json")
    parser.add_argument("--require-cols-docgrid", action="store_true")
    parser.add_argument("--compare-reference", action="store_true")
    parser.add_argument("--pdf-page-box-tolerance-points", type=float, default=2.5)
    args = parser.parse_args(argv)

    final_docx = Path(args.final_docx)
    reference_docx = Path(args.reference_docx) if args.reference_docx else None
    report, issues, reference_layout = audit_docx_sections(
        final_docx=final_docx,
        reference_docx=reference_docx,
        require_cols_docgrid=args.require_cols_docgrid,
        compare_reference=args.compare_reference,
    )
    if args.final_pdf:
        pdf_report, pdf_issues = audit_pdf_page_boxes(
            final_pdf=Path(args.final_pdf),
            reference_layout=reference_layout,
            tolerance_points=args.pdf_page_box_tolerance_points,
        )
        report["pdf_page_box_audit"] = pdf_report
        issues.extend(pdf_issues)

    report["issue_count"] = len(issues)
    report["issues"] = issues
    report["verdict"] = "pass" if not issues else "fail"

    text = json.dumps(report, ensure_ascii=True, indent=2) + "\n"
    if args.report_json:
        Path(args.report_json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report_json).write_text(text, encoding="utf-8")
    print("section page setup audit: " + ("PASS" if not issues else "FAIL"))
    if issues:
        for issue in issues[:20]:
            print(issue)
    return 0 if not issues else 1


if __name__ == "__main__":
    raise SystemExit(main())
