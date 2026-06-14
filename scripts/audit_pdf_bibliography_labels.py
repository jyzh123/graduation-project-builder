from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path


REFERENCE_HEADING_RE = re.compile(r"(\u53c2\u8003\u6587\u732e|References)\s*$", re.I)
TAIL_HEADING_RE = re.compile(r"(\u81f4\u8c22|Acknowledgements?|Appendix|Appendices|\u9644\u5f55(?:[A-Z\uff21-\uff3a])?)\s*$", re.I)
BRACKET_LABEL_RE = re.compile(r"^\s*\[(\d{1,3})\]")
DOT_LABEL_RE = re.compile(r"^\s*(\d{1,3})\.(?!\d)")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def line_text(line: dict[str, object]) -> str:
    spans = line.get("spans", [])
    if not isinstance(spans, list):
        return ""
    return "".join(str(span.get("text", "")) for span in spans if isinstance(span, dict))


def line_bbox(line: dict[str, object]) -> tuple[float, float, float, float]:
    bbox = line.get("bbox", [0, 0, 0, 0])
    if not isinstance(bbox, (list, tuple)) or len(bbox) != 4:
        return (0.0, 0.0, 0.0, 0.0)
    try:
        return tuple(float(value) for value in bbox)  # type: ignore[return-value]
    except (TypeError, ValueError):
        return (0.0, 0.0, 0.0, 0.0)


def collect_reference_lines(pdf_path: Path) -> tuple[int, list[dict[str, object]]]:
    import fitz  # type: ignore

    pages: list[dict[str, object]] = []
    reference_start_page = -1
    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc, start=1):
            text = page.get_text("text")
            if reference_start_page < 0 and any(
                REFERENCE_HEADING_RE.search(part.strip()) for part in text.splitlines()
            ):
                reference_start_page = page_index
            text_dict = page.get_text("dict")
            raw_lines: list[dict[str, object]] = []
            for block in text_dict.get("blocks", []):
                if not isinstance(block, dict):
                    continue
                for line in block.get("lines", []):
                    if isinstance(line, dict):
                        raw_lines.append(line)
            raw_lines.sort(key=lambda item: (line_bbox(item)[1], line_bbox(item)[0]))
            for raw_line in raw_lines:
                text_line = re.sub(r"\s+", " ", line_text(raw_line)).strip()
                if not text_line:
                    continue
                pages.append(
                    {
                        "page": page_index,
                        "text": text_line,
                        "bbox": [round(value, 2) for value in line_bbox(raw_line)],
                    }
                )
    if reference_start_page < 0:
        reference_start_page = max(1, pages[-1]["page"] - 5) if pages else 1
    reference_lines: list[dict[str, object]] = []
    for row in pages:
        if int(row["page"]) < reference_start_page:
            continue
        text = str(row["text"]).strip()
        if TAIL_HEADING_RE.search(text):
            break
        reference_lines.append(row)
    return reference_start_page, reference_lines


def audit_pdf_labels(
    pdf_path: Path,
    expected_style: str,
    min_count: int,
    expected_label_content_spacing: str = "any",
) -> dict[str, object]:
    issues: list[str] = []
    reference_start_page, lines = collect_reference_lines(pdf_path)
    bracket_labels: list[dict[str, object]] = []
    dot_labels: list[dict[str, object]] = []
    for row in lines:
        text = str(row["text"])
        bracket_match = BRACKET_LABEL_RE.match(text)
        dot_match = DOT_LABEL_RE.match(text)
        if bracket_match:
            bracket_labels.append({**row, "number": int(bracket_match.group(1)), "style": "bracket"})
        if dot_match:
            dot_labels.append({**row, "number": int(dot_match.group(1)), "style": "dot"})

    selected = bracket_labels if expected_style == "bracket" else dot_labels
    unexpected = dot_labels if expected_style == "bracket" else bracket_labels
    selected_numbers = [int(row["number"]) for row in selected]
    core_selected = [row for row in selected if 1 <= int(row["number"]) <= min_count]
    core_selected_numbers = [int(row["number"]) for row in core_selected]
    expected_numbers = list(range(1, min_count + 1))
    first_numbers = core_selected_numbers[:min_count]
    if len(core_selected) < min_count:
        issues.append(
            f"expected at least {min_count} {expected_style} bibliography labels, found {len(core_selected)}"
        )
    if first_numbers != expected_numbers:
        issues.append(
            f"{expected_style} bibliography labels are not sequential 1..{min_count}: first_numbers={first_numbers}"
        )
    unexpected_core = [row for row in unexpected if 1 <= int(row["number"]) <= min_count]
    if unexpected_core:
        issues.append(
            f"unexpected bibliography label family present: {unexpected_core[:10]}"
        )
    spacing_issues: list[dict[str, object]] = []
    if expected_label_content_spacing != "any":
        for row in core_selected[:min_count]:
            text = str(row["text"])
            number = int(row["number"])
            if expected_style == "bracket":
                pattern = rf"^\s*\[{number}\](\s+)?\S"
            else:
                pattern = rf"^\s*{number}\.(\s+)?\S"
            match = re.match(pattern, text)
            has_space = bool(match and match.group(1))
            bad = (expected_label_content_spacing == "none" and has_space) or (
                expected_label_content_spacing == "space" and not has_space
            )
            if bad:
                spacing_issues.append(
                    {
                        "page": row.get("page"),
                        "number": number,
                        "text": text,
                        "expected_spacing": expected_label_content_spacing,
                    }
                )
        if spacing_issues:
            issues.append(
                f"bibliography label/content spacing mismatch: expected {expected_label_content_spacing}, "
                f"found {len(spacing_issues)} mismatches"
            )

    x_positions = [float(row["bbox"][0]) for row in core_selected[:min_count] if row.get("bbox")]
    geometry_verdict = "not-applicable"
    geometry_spread = 0.0
    if x_positions:
        geometry_spread = max(x_positions) - min(x_positions)
        geometry_verdict = "pass" if geometry_spread <= 18.0 else "fail"
        if geometry_verdict != "pass":
            issues.append(f"bibliography label x-position spread too large: {geometry_spread:.2f} pt")

    result = "pass" if not issues else "fail"
    return {
        "schema": "graduation-project-builder.pdf-bibliography-label-audit.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "result": result,
        "pdf_path": str(pdf_path.resolve()),
        "pdf_sha256": sha256_file(pdf_path),
        "reference_start_page": reference_start_page,
        "expected_label_family": expected_style,
        "chosen_label_family": expected_style,
        "expected_label_content_spacing": expected_label_content_spacing,
        "min_count": min_count,
        "bracket_label_count": len(bracket_labels),
        "dot_label_count": len(dot_labels),
        "selected_label_numbers": selected_numbers,
        "core_selected_label_numbers": core_selected_numbers,
        "unexpected_label_numbers": [int(row["number"]) for row in unexpected],
        "unexpected_core_label_count": len(unexpected_core),
        "unexpected_core_label_numbers": [int(row["number"]) for row in unexpected_core],
        "label_geometry_verdict": geometry_verdict,
        "label_x_position_spread_pt": round(geometry_spread, 2),
        "label_rows_sample": core_selected[: min(80, len(core_selected))],
        "non_core_selected_label_rows_sample": [
            row for row in selected if not (1 <= int(row["number"]) <= min_count)
        ][:40],
        "unexpected_label_rows_sample": unexpected[: min(40, len(unexpected))],
        "label_content_spacing_issue_count": len(spacing_issues),
        "label_content_spacing_issues_sample": spacing_issues[:40],
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit rendered PDF bibliography label family and geometry.")
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--expected-style", choices=("bracket", "dot"), required=True)
    parser.add_argument("--expected-label-content-spacing", choices=("any", "space", "none"), default="any")
    parser.add_argument("--min-count", type=int, default=60)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    report = audit_pdf_labels(
        args.pdf,
        args.expected_style,
        max(1, args.min_count),
        args.expected_label_content_spacing,
    )
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
