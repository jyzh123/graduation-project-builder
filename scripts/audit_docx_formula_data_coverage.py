from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def text_from_docx(docx: Path) -> str:
    with zipfile.ZipFile(docx, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return "".join(node.text or "" for node in root.findall(".//w:t", NS))


def validate_entry(entry: dict[str, object]) -> list[str]:
    issues: list[str] = []
    if not str(entry.get("normalized_number", "")).strip():
        issues.append("missing normalized_number")
    if not str(entry.get("formula_text", "")).strip():
        issues.append("missing formula_text")
    values = entry.get("input_values")
    if not isinstance(values, list) or not values:
        issues.append("missing input_values")
    else:
        for idx, item in enumerate(values):
            if not isinstance(item, dict):
                issues.append(f"input_values[{idx}] not object")
                continue
            for key in ("symbol", "value", "unit", "source"):
                if not str(item.get(key, "")).strip():
                    issues.append(f"input_values[{idx}] missing {key}")
            if not re.search(r"\d|按|由|图纸|公式|台账|许用|样本", str(item.get("value", ""))):
                issues.append(f"input_values[{idx}] value lacks data marker")
    if not str(entry.get("data_source", "")).strip():
        issues.append("missing data_source")
    if not str(entry.get("substitution_note", "")).strip():
        issues.append("missing substitution_note")
    return issues


def audit(docx: Path, data_map_path: Path, min_formula_count: int) -> dict[str, object]:
    data = json.loads(data_map_path.read_text(encoding="utf-8-sig"))
    entries = data.get("formulas")
    if not isinstance(entries, list):
        entries = []
    docx_text = text_from_docx(docx)
    data_map_docx_sha = str(data.get("docx_sha256", "")).upper()
    current_sha = sha256_file(docx)
    issue_rows: list[dict[str, object]] = []
    seen_numbers: set[str] = set()
    visible_note_hits = 0
    for entry in entries:
        if not isinstance(entry, dict):
            issue_rows.append({"issue": "entry-not-object"})
            continue
        number = str(entry.get("normalized_number", ""))
        seen_numbers.add(number)
        entry_issues = validate_entry(entry)
        visible_tokens = [
            f"公式数据台账 {number}：",
            f"公式数据说明 {number}：",
        ]
        if any(token in docx_text for token in visible_tokens):
            visible_note_hits += 1
        else:
            entry_issues.append("visible data note missing in DOCX")
        if entry_issues:
            issue_rows.append({"number": number, "issues": entry_issues})
    missing_count = max(0, min_formula_count - len(seen_numbers))
    if data_map_docx_sha and data_map_docx_sha != current_sha:
        issue_rows.append(
            {
                "issue": "data-map-source-hash-note",
                "detail": "data map was generated from a source DOCX before data-note insertion; current hash is reported separately",
                "data_map_docx_sha256": data_map_docx_sha,
                "current_docx_sha256": current_sha,
            }
        )
    blocking_issues = [
        row
        for row in issue_rows
        if row.get("issue") != "data-map-source-hash-note"
    ]
    result = (
        "pass"
        if len(seen_numbers) >= min_formula_count
        and visible_note_hits >= min_formula_count
        and not blocking_issues
        else "fail"
    )
    return {
        "schema": "graduation-project-builder.formula-data-coverage-audit.v1",
        "result": result,
        "docx_path": str(docx.resolve()),
        "docx_sha256": current_sha,
        "data_map_path": str(data_map_path.resolve()),
        "data_map_docx_sha256": data_map_docx_sha,
        "formula_data_entry_count": len(entries),
        "unique_formula_data_number_count": len(seen_numbers),
        "visible_formula_data_note_count": visible_note_hits,
        "min_formula_data_count": min_formula_count,
        "missing_formula_data_count": missing_count,
        "issue_count": len(blocking_issues),
        "issues": issue_rows,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--data-map", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--min-formula-data-count", type=int, default=200)
    args = parser.parse_args()
    report = audit(args.docx, args.data_map, args.min_formula_data_count)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
