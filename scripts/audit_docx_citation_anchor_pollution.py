#!/usr/bin/env python3
"""Audit visible citation-anchor pollution in DOCX and rendered PDF outputs."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


SCHEMA = "graduation-project-builder.citation-anchor-pollution-audit.v1"
GENERATOR = "scripts/audit_docx_citation_anchor_pollution.py"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

ANCHOR_TOKEN_RE = re.compile(
    r"\b(?:cite_ref|ref_anchor|bookmark)_(?:[A-Za-z0-9][A-Za-z0-9_.:-]*)?\b"
    r"|\b(?:cite_ref|ref_anchor|bookmark)_",
    re.IGNORECASE,
)
POLLUTED_MARKER_RE = re.compile(
    r"\[\d{1,4}\]\s*(?:cite_ref|ref_anchor|bookmark)_",
    re.IGNORECASE,
)
FIELD_HYPERLINK_ANCHOR_RE = re.compile(
    r'HYPERLINK\s+\\l\s+"([^"]*(?:cite_ref|ref_anchor|bookmark)_[^"]*)"',
    re.IGNORECASE,
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _path_payload(path: Path) -> dict[str, str]:
    return {"path": str(path.resolve()), "sha256": sha256_file(path)}


def _matches(text: str) -> list[str]:
    return [match.group(0) for match in ANCHOR_TOKEN_RE.finditer(text or "")]


def _visible_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(f"{W}t"))


def _paragraph_records(root: ET.Element, part_name: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for index, paragraph in enumerate(root.iter(f"{W}p"), start=1):
        text = _visible_text(paragraph)
        hits = sorted(set(_matches(text)))
        if hits:
            records.append(
                {
                    "part": part_name,
                    "paragraph_index": index,
                    "hits": hits,
                    "text_excerpt": text[:300],
                }
            )
    return records


def _field_records(root: ET.Element, part_name: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    instruction_records: list[dict[str, Any]] = []
    result_records: list[dict[str, Any]] = []
    unclosed_field_count = 0

    for paragraph_index, paragraph in enumerate(root.iter(f"{W}p"), start=1):
        stack: list[dict[str, Any]] = []
        for element in paragraph.iter():
            if element.tag == f"{W}fldChar":
                fld_type = element.attrib.get(f"{W}fldCharType", "")
                if fld_type == "begin":
                    stack.append({"instruction": [], "in_result": False, "result_hits": []})
                elif fld_type == "separate" and stack:
                    stack[-1]["in_result"] = True
                elif fld_type == "end" and stack:
                    current = stack.pop()
                    instruction = "".join(current["instruction"])
                    for anchor in FIELD_HYPERLINK_ANCHOR_RE.findall(instruction):
                        instruction_records.append(
                            {
                                "part": part_name,
                                "paragraph_index": paragraph_index,
                                "anchor": anchor,
                                "instruction_excerpt": instruction[:300],
                            }
                        )
                    for hit in current["result_hits"]:
                        result_records.append(hit)
                continue

            if element.tag == f"{W}instrText" and stack:
                stack[-1]["instruction"].append(element.text or "")
                continue

            if element.tag == f"{W}t" and stack:
                visible_hits = sorted(set(_matches(element.text or "")))
                if visible_hits:
                    for current in stack:
                        if current.get("in_result"):
                            current["result_hits"].append(
                                {
                                    "part": part_name,
                                    "paragraph_index": paragraph_index,
                                    "hits": visible_hits,
                                    "text_excerpt": (element.text or "")[:300],
                                }
                            )
        unclosed_field_count += len(stack)
        for current in stack:
            for hit in current.get("result_hits", []):
                result_records.append(hit)

    return instruction_records, result_records, unclosed_field_count


def _bookmark_records(root: ET.Element, part_name: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for bookmark in root.iter(f"{W}bookmarkStart"):
        name = bookmark.attrib.get(f"{W}name", "")
        if _matches(name):
            records.append({"part": part_name, "name": name})
    return records


def audit_docx(docx_path: Path, *, fail_on_duplicate_bookmarks: bool = False) -> dict[str, Any]:
    docx_path = docx_path.resolve()
    report: dict[str, Any] = {
        **_path_payload(docx_path),
        "word_xml_part_count": 0,
        "visible_text_node_count": 0,
        "visible_anchor_hit_count": 0,
        "visible_anchor_polluted_text_node_count": 0,
        "visible_anchor_polluted_paragraph_count": 0,
        "visible_anchor_polluted_paragraphs": [],
        "field_instruction_anchor_count": 0,
        "field_instruction_anchor_records": [],
        "field_result_anchor_hit_count": 0,
        "field_result_anchor_records": [],
        "unclosed_field_count": 0,
        "bookmark_anchor_name_count": 0,
        "duplicate_anchor_bookmark_count": 0,
        "duplicate_anchor_bookmarks": [],
    }

    bookmark_records: list[dict[str, Any]] = []
    with zipfile.ZipFile(docx_path) as zf:
        for part_name in sorted(name for name in zf.namelist() if name.startswith("word/") and name.endswith(".xml")):
            try:
                root = ET.fromstring(zf.read(part_name))
            except ET.ParseError:
                continue
            report["word_xml_part_count"] += 1
            text_nodes = list(root.iter(f"{W}t"))
            report["visible_text_node_count"] += len(text_nodes)
            for node in text_nodes:
                hits = _matches(node.text or "")
                if hits:
                    report["visible_anchor_hit_count"] += len(hits)
                    report["visible_anchor_polluted_text_node_count"] += 1
            paragraph_hits = _paragraph_records(root, part_name)
            report["visible_anchor_polluted_paragraphs"].extend(paragraph_hits)
            instructions, results, unclosed = _field_records(root, part_name)
            report["field_instruction_anchor_records"].extend(instructions)
            report["field_result_anchor_records"].extend(results)
            report["unclosed_field_count"] += unclosed
            bookmark_records.extend(_bookmark_records(root, part_name))

    report["visible_anchor_polluted_paragraph_count"] = len(report["visible_anchor_polluted_paragraphs"])
    report["field_instruction_anchor_count"] = len(report["field_instruction_anchor_records"])
    report["field_result_anchor_hit_count"] = sum(len(item.get("hits", [])) for item in report["field_result_anchor_records"])
    report["bookmark_anchor_name_count"] = len(bookmark_records)

    counts = Counter(record["name"] for record in bookmark_records)
    duplicate_names = sorted(name for name, count in counts.items() if count > 1)
    report["duplicate_anchor_bookmark_count"] = len(duplicate_names)
    locations_by_name: dict[str, list[str]] = defaultdict(list)
    for record in bookmark_records:
        if record["name"] in duplicate_names:
            locations_by_name[record["name"]].append(record["part"])
    report["duplicate_anchor_bookmarks"] = [
        {"name": name, "count": counts[name], "parts": sorted(set(locations_by_name[name]))}
        for name in duplicate_names
    ]
    report["fail_on_duplicate_bookmarks"] = fail_on_duplicate_bookmarks
    return report


def audit_pdf(pdf_path: Path) -> dict[str, Any]:
    pdf_path = pdf_path.resolve()
    report: dict[str, Any] = {
        **_path_payload(pdf_path),
        "page_count": 0,
        "text_length": 0,
        "visible_anchor_hit_count": 0,
        "polluted_marker_pattern_count": 0,
        "visible_anchor_records": [],
    }
    try:
        import fitz  # type: ignore
    except Exception as exc:  # pragma: no cover - environment dependent
        report["extraction_error"] = f"PyMuPDF/fitz unavailable: {exc}"
        return report

    try:
        with fitz.open(str(pdf_path)) as document:
            report["page_count"] = document.page_count
            for page_index, page in enumerate(document, start=1):
                text = page.get_text() or ""
                report["text_length"] += len(text)
                hits = _matches(text)
                marker_hits = POLLUTED_MARKER_RE.findall(text)
                report["visible_anchor_hit_count"] += len(hits)
                report["polluted_marker_pattern_count"] += len(marker_hits)
                if hits or marker_hits:
                    report["visible_anchor_records"].append(
                        {
                            "page": page_index,
                            "hits": sorted(set(hits)),
                            "polluted_marker_pattern_count": len(marker_hits),
                            "text_excerpt": text[:500],
                        }
                    )
    except Exception as exc:
        report["extraction_error"] = f"PDF text extraction failed: {exc}"
    return report


def audit_docx_citation_anchor_pollution(
    docx_path: Path,
    *,
    rendered_pdf_path: Path | None = None,
    fail_on_duplicate_bookmarks: bool = False,
) -> dict[str, Any]:
    docx_report = audit_docx(docx_path, fail_on_duplicate_bookmarks=fail_on_duplicate_bookmarks)
    pdf_report = audit_pdf(rendered_pdf_path) if rendered_pdf_path is not None else None

    error_codes: list[str] = []
    if docx_report["visible_anchor_hit_count"] > 0:
        error_codes.append("DOCX_VISIBLE_CITATION_ANCHOR_POLLUTION")
    if docx_report["field_result_anchor_hit_count"] > 0:
        error_codes.append("DOCX_FIELD_RESULT_CITATION_ANCHOR_POLLUTION")
    if fail_on_duplicate_bookmarks and docx_report["duplicate_anchor_bookmark_count"] > 0:
        error_codes.append("DOCX_DUPLICATE_CITATION_ANCHOR_BOOKMARKS")
    if pdf_report is not None:
        if pdf_report.get("extraction_error"):
            error_codes.append("PDF_TEXT_EXTRACTION_UNAVAILABLE")
        if pdf_report.get("visible_anchor_hit_count", 0) > 0:
            error_codes.append("PDF_VISIBLE_CITATION_ANCHOR_POLLUTION")
        if pdf_report.get("polluted_marker_pattern_count", 0) > 0:
            error_codes.append("PDF_MARKER_CITATION_ANCHOR_POLLUTION")

    return {
        "schema": SCHEMA,
        "generator": GENERATOR,
        "docx": docx_report,
        "pdf": pdf_report,
        "error_codes": error_codes,
        "verdict": "pass" if not error_codes else "fail",
        "notes": [
            "Visible w:t text is scanned across all word/*.xml story parts.",
            "HYPERLINK field instructions that name cite_ref/ref_anchor/bookmark targets are allowed only when their field result does not expose the anchor.",
            "Rendered PDF text is scanned when --rendered-pdf is supplied.",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--rendered-pdf", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--fail-on-duplicate-bookmarks", action="store_true")
    args = parser.parse_args(argv)

    report = audit_docx_citation_anchor_pollution(
        args.docx,
        rendered_pdf_path=args.rendered_pdf,
        fail_on_duplicate_bookmarks=args.fail_on_duplicate_bookmarks,
    )
    text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(text, encoding="utf-8")
    sys.stdout.write(text)
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
