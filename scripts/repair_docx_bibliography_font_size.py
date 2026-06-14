#!/usr/bin/env python3
"""Repair bibliography entry font size in a DOCX without rewriting content."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
REFERENCE_HEADINGS = {"\u53c2\u8003\u6587\u732e", "references", "bibliography"}
TAIL_HEADINGS = {"\u81f4\u8c22", "\u9644\u5f55", "acknowledgement", "acknowledgements", "appendix"}
BIBLIOGRAPHY_MARKERS = (
    "[J]",
    "[M]",
    "[C]",
    "[D]",
    "[N]",
    "[R]",
    "[S]",
    "[P]",
    "[EB/OL]",
    "[DB/OL]",
    "[CP/OL]",
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").lower()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def is_reference_heading(text: str) -> bool:
    return normalize(text) in {normalize(item) for item in REFERENCE_HEADINGS}


def is_tail_heading(text: str) -> bool:
    key = normalize(text)
    return key in {normalize(item) for item in TAIL_HEADINGS}


def paragraph_num_id(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:numPr/w:numId", NS)
    return node.get(W + "val", "") if node is not None else ""


def is_bibliography_entry(paragraph: ET.Element, text: str) -> bool:
    if paragraph_num_id(paragraph):
        return True
    stripped = (text or "").strip()
    if re.match(r"^\s*(?:[\[\uff3b]\d{1,3}[\]\uff3d]|\d{1,3}[\.．、])\s*\S+", stripped):
        return True
    upper_text = stripped.upper()
    if any(marker in upper_text for marker in BIBLIOGRAPHY_MARKERS):
        return True
    if re.search(r"https?://|doi[:：]", stripped, re.IGNORECASE):
        return True
    return False


def ensure_child(parent: ET.Element, tag: str, *, first: bool = False) -> ET.Element:
    child = parent.find(f"./w:{tag.removeprefix(W)}", NS)
    if child is not None:
        return child
    child = ET.Element(tag)
    if first:
        parent.insert(0, child)
    else:
        parent.append(child)
    return child


def set_size(rpr: ET.Element, half_points: str, size_cs_half_points: str | None = None) -> bool:
    changed = False
    slots = {
        "sz": half_points,
        "szCs": size_cs_half_points if size_cs_half_points is not None else half_points,
    }
    for tag, value in slots.items():
        element = rpr.find(f"./w:{tag}", NS)
        if element is None:
            element = ET.Element(W + tag)
            rpr.append(element)
            changed = True
        if element.attrib.get(W + "val") != value:
            element.set(W + "val", value)
            changed = True
    return changed


def repair_document_xml(
    xml_bytes: bytes,
    half_points: str,
    size_cs_half_points: str | None = None,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        return xml_bytes, {
            "bibliography_entry_count": 0,
            "changed_entry_count": 0,
            "changed_run_count": 0,
            "changed_paragraph_mark_count": 0,
            "issues": ["word/document.xml has no w:body"],
        }

    in_references = False
    entry_count = 0
    changed_entry_count = 0
    changed_run_count = 0
    changed_paragraph_mark_count = 0
    changed_entries: list[dict[str, object]] = []

    for body_index, child in enumerate(list(body)):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        if not text:
            continue
        if not in_references:
            if is_reference_heading(text):
                in_references = True
            continue
        if is_tail_heading(text):
            break
        if not is_bibliography_entry(child, text):
            continue

        entry_count += 1
        paragraph_changed = False
        ppr = ensure_child(child, W + "pPr", first=True)
        paragraph_mark_rpr = ensure_child(ppr, W + "rPr")
        if set_size(paragraph_mark_rpr, half_points, size_cs_half_points):
            paragraph_changed = True
            changed_paragraph_mark_count += 1

        run_changes = 0
        for run in child.findall("w:r", NS):
            if not paragraph_text(run).strip():
                continue
            rpr = run.find("./w:rPr", NS)
            if rpr is None:
                rpr = ET.Element(W + "rPr")
                run.insert(0, rpr)
            if set_size(rpr, half_points, size_cs_half_points):
                paragraph_changed = True
                run_changes += 1

        changed_run_count += run_changes
        if paragraph_changed:
            changed_entry_count += 1
            changed_entries.append(
                {
                    "body_child_index": body_index,
                    "entry_number": entry_count,
                    "run_changes": run_changes,
                    "text_prefix": text[:100],
                }
            )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), {
        "bibliography_entry_count": entry_count,
        "changed_entry_count": changed_entry_count,
        "changed_run_count": changed_run_count,
        "changed_paragraph_mark_count": changed_paragraph_mark_count,
        "changed_entries": changed_entries,
        "issues": [] if entry_count else ["no bibliography entries found"],
    }


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    half_points: str,
    size_cs_half_points: str | None = None,
) -> dict[str, object]:
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()
    if input_docx == output_docx:
        raise RuntimeError("output DOCX must be a new review-copy path")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    changed_parts: list[str] = []
    repair_summary: dict[str, object] = {}
    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data, repair_summary = repair_document_xml(data, half_points, size_cs_half_points)
                if int(repair_summary.get("changed_entry_count") or 0) > 0:
                    changed_parts.append(info.filename)
            zout.writestr(info, data)

    issues = list(repair_summary.get("issues") or [])
    return {
        "schema": "graduation-project-builder.bibliography-font-size-repair.v1",
        "input_docx_path": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx_path": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "target_size_half_points": half_points,
        "target_size_cs_half_points": size_cs_half_points if size_cs_half_points is not None else half_points,
        "target_size_points": int(half_points) / 2,
        "target_size_cs_points": int(size_cs_half_points if size_cs_half_points is not None else half_points) / 2,
        "changed_parts": changed_parts,
        **repair_summary,
        "verdict": "pass" if not issues and int(repair_summary.get("bibliography_entry_count") or 0) > 0 else "fail",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--size-half-points", default="21")
    parser.add_argument(
        "--size-cs-half-points",
        default=None,
        help="Optional w:szCs half-point value. Defaults to --size-half-points for backwards compatibility.",
    )
    parser.add_argument("--report-json", required=True, type=Path)
    args = parser.parse_args()

    if not re.fullmatch(r"\d{1,3}", str(args.size_half_points)):
        raise SystemExit("--size-half-points must be an integer half-point value")
    if args.size_cs_half_points is not None and not re.fullmatch(r"\d{1,3}", str(args.size_cs_half_points)):
        raise SystemExit("--size-cs-half-points must be an integer half-point value")
    report = repair_docx(
        args.input,
        args.output,
        str(args.size_half_points),
        str(args.size_cs_half_points) if args.size_cs_half_points is not None else None,
    )
    args.report_json.resolve().parent.mkdir(parents=True, exist_ok=True)
    args.report_json.resolve().write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"verdict": report["verdict"], "changed_entry_count": report["changed_entry_count"]}, ensure_ascii=False))
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
