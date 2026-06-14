from __future__ import annotations

import argparse
import hashlib
import json
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


FORMULA_LABEL_RE = re.compile(r"式\s*[（(]\s*(\d+)\s*[-－.．]\s*(\d+)\s*[)）]")
SOURCE_NUMBER_RE = re.compile(r"(?:式\s*)?[（(]?\s*(\d+)\s*[-－.．]\s*(\d+)\s*[)）]?")


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{" + NS["w"] + "}"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def normalize_number(value: object) -> str | None:
    if value is None:
        return None
    match = SOURCE_NUMBER_RE.search(str(value).strip())
    if not match:
        return None
    return f"{int(match.group(1))}-{int(match.group(2))}"


def load_source_numbers(paths: list[Path], source_docxes: list[Path] | None = None) -> dict[str, list[dict[str, object]]]:
    formulas: dict[str, list[dict[str, object]]] = {}
    for path in paths:
        data = json.loads(path.read_text(encoding="utf-8"))
        for item in data.get("formulas", []):
            number = normalize_number(item.get("number"))
            if number is None:
                continue
            formulas.setdefault(number, []).append(
                {
                    "source_map": str(path),
                    "source_text": item.get("source_text"),
                    "number": item.get("number"),
                }
            )
    for path in source_docxes or []:
        for number in sorted(set(extract_docx_formula_object_labels(path)), key=label_sort_key):
            formulas.setdefault(number, []).append(
                {
                    "source_docx": str(path),
                    "source_text": "docx-formula-object-label",
                    "number": f"({number})",
                }
            )
    return formulas


def extract_docx_formula_labels(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    return [f"{int(chapter)}-{int(seq)}" for chapter, seq in FORMULA_LABEL_RE.findall(xml)]


def element_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def element_has_formula_object(element: ET.Element) -> bool:
    return (
        element.find(".//m:oMath", NS) is not None
        or element.find(".//m:oMathPara", NS) is not None
        or element.find(".//w:object", NS) is not None
    )


def extract_docx_formula_object_labels(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return []
    labels: list[str] = []
    for child in list(body):
        if child.tag not in {W + "p", W + "tbl"}:
            continue
        if not element_has_formula_object(child):
            continue
        labels.extend(f"{int(chapter)}-{int(seq)}" for chapter, seq in FORMULA_LABEL_RE.findall(element_text(child)))
    return labels


def label_sort_key(value: str) -> tuple[int, int]:
    chapter, seq = value.split("-", 1)
    return int(chapter), int(seq)


def audit(
    docx_path: Path,
    source_maps: list[Path],
    source_docxes: list[Path],
    min_source_coverage_ratio: float,
    min_unique_visible_formula_labels: int | None,
) -> dict[str, object]:
    source_numbers = load_source_numbers(source_maps, source_docxes)
    visible_labels = extract_docx_formula_labels(docx_path)
    labels = extract_docx_formula_object_labels(docx_path)
    unique_labels = sorted(set(labels), key=label_sort_key)
    unique_visible_labels = sorted(set(visible_labels), key=label_sort_key)
    label_only = sorted(set(visible_labels) - set(labels), key=label_sort_key)
    source_set = set(source_numbers)
    final_set = set(unique_labels)
    matched = sorted(source_set & final_set, key=label_sort_key)
    missing = sorted(source_set - final_set, key=label_sort_key)
    extra = sorted(final_set - source_set, key=label_sort_key)
    coverage_ratio = (len(matched) / len(source_set)) if source_set else 1.0

    issues: list[dict[str, object]] = []
    if source_set and coverage_ratio < min_source_coverage_ratio:
        issues.append(
            {
                "issue": "formula-source-coverage-below-threshold",
                "source_formula_count": len(source_set),
                "matched_source_formula_count": len(matched),
                "coverage_ratio": coverage_ratio,
                "required_ratio": min_source_coverage_ratio,
                "missing_source_formula_numbers": missing,
            }
        )
    if min_unique_visible_formula_labels is not None and len(unique_labels) < min_unique_visible_formula_labels:
        issues.append(
            {
                "issue": "unique-visible-formula-label-count-below-threshold",
                "unique_visible_formula_label_count": len(unique_labels),
                "required_unique_visible_formula_label_count": min_unique_visible_formula_labels,
            }
        )

    per_chapter: dict[str, int] = {}
    for label in unique_labels:
        chapter = label.split("-", 1)[0]
        per_chapter[chapter] = per_chapter.get(chapter, 0) + 1

    return {
        "schema": "graduation-project-builder.formula-source-coverage-audit.v1",
        "result": "pass" if not issues else "fail",
        "docx_path": str(docx_path),
        "docx_sha256": sha256_file(docx_path),
        "source_map_paths": [str(path) for path in source_maps],
        "source_docx_paths": [str(path) for path in source_docxes],
        "source_docx_formula_label_counts": {
            str(path): len(set(extract_docx_formula_object_labels(path))) for path in source_docxes
        },
        "source_formula_count": len(source_set),
        "matched_source_formula_count": len(matched),
        "missing_source_formula_count": len(missing),
        "coverage_ratio": coverage_ratio,
        "min_source_coverage_ratio": min_source_coverage_ratio,
        "strict_visible_formula_label_count": len(visible_labels),
        "unique_visible_formula_label_count": len(unique_labels),
        "raw_unique_visible_formula_label_count": len(unique_visible_labels),
        "formula_object_visible_label_count": len(labels),
        "unique_formula_object_visible_label_count": len(unique_labels),
        "label_only_formula_label_count": len(label_only),
        "label_only_formula_numbers": label_only,
        "min_unique_visible_formula_labels": min_unique_visible_formula_labels,
        "visible_formula_label_count_by_chapter": per_chapter,
        "matched_source_formula_numbers": matched,
        "missing_source_formula_numbers": missing,
        "extra_visible_formula_numbers_not_in_source": extra,
        "issues": issues,
    }


def write_report(report: dict[str, object], path: Path | None) -> None:
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if path is not None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text + "\n", encoding="utf-8")
    print(text)


def make_minimal_docx(path: Path, labels: list[str]) -> None:
    body = "".join(f"<w:p><w:r><w:t>式({label})</w:t></w:r></w:p>" for label in labels)
    body = "".join(
        "<w:tbl><w:tr>"
        "<w:tc><w:p/></w:tc>"
        "<w:tc><w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p></w:tc>"
        f"<w:tc><w:p><w:r><w:t>&#x5F0F;({label})</w:t></w:r></w:p></w:tc>"
        "</w:tr></w:tbl>"
        for label in labels
    )
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        f"<w:body>{body}</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)


def self_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        docx = root / "sample.docx"
        make_minimal_docx(docx, ["1-1", "1-2"])
        ok_map = root / "ok.json"
        ok_map.write_text(
            json.dumps(
                {
                    "formulas": [
                        {"source_text": "A", "number": "(1-1)"},
                        {"source_text": "B", "number": "(1-2)"},
                    ]
                }
            ),
            encoding="utf-8",
        )
        bad_map = root / "bad.json"
        bad_map.write_text(
            json.dumps(
                {
                    "formulas": [
                        {"source_text": "A", "number": "(1-1)"},
                        {"source_text": "B", "number": "(1-2)"},
                        {"source_text": "C", "number": "(1-3)"},
                    ]
                }
            ),
            encoding="utf-8",
        )
        ok = audit(docx, [ok_map], [], 1.0, 2)
        docx_source = root / "source.docx"
        make_minimal_docx(docx_source, ["1-1", "1-2"])
        ok_docx_source = audit(docx, [], [docx_source], 1.0, 2)
        bad = audit(docx, [bad_map], [], 1.0, 3)
        if ok["result"] != "pass":
            raise AssertionError(ok)
        if ok_docx_source["result"] != "pass":
            raise AssertionError(ok_docx_source)
        if bad["result"] != "fail":
            raise AssertionError(bad)
    print("self-test pass")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", nargs="?")
    parser.add_argument("--source-map", action="append", default=[])
    parser.add_argument("--source-docx", action="append", default=[])
    parser.add_argument("--report")
    parser.add_argument("--min-source-coverage-ratio", type=float, default=1.0)
    parser.add_argument("--min-unique-visible-formula-labels", type=int)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)

    if args.self_test:
        return self_test()
    if not args.docx:
        parser.error("docx is required unless --self-test is used")
    if not args.source_map and not args.source_docx:
        parser.error("--source-map or --source-docx is required unless --self-test is used")

    report = audit(
        Path(args.docx),
        [Path(item) for item in args.source_map],
        [Path(item) for item in args.source_docx],
        args.min_source_coverage_ratio,
        args.min_unique_visible_formula_labels,
    )
    write_report(report, Path(args.report) if args.report else None)
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
