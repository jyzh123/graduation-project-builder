#!/usr/bin/env python3
"""Repair thesis front-matter/body page-number sections without touching body text."""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


SECTION_CHILD_ORDER = {
    "headerReference": 0,
    "footerReference": 1,
    "footnotePr": 2,
    "endnotePr": 3,
    "type": 4,
    "pgSz": 5,
    "pgMar": 6,
    "paperSrc": 7,
    "pgBorders": 8,
    "lnNumType": 9,
    "pgNumType": 10,
    "cols": 11,
    "formProt": 12,
    "vAlign": 13,
    "noEndnote": 14,
    "titlePg": 15,
    "textDirection": 16,
    "bidi": 17,
    "rtlGutter": 18,
    "docGrid": 19,
    "printerSettings": 20,
}


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", str(text or ""))


def is_zh_abstract_title(text: str) -> bool:
    compact = compact_text(text)
    if compact in {"\u6458\u8981", "\u4e2d\u6587\u6458\u8981"}:
        return True
    return compact.startswith(("\u6458\u8981\uff1a", "\u6458\u8981:", "\u4e2d\u6587\u6458\u8981\uff1a", "\u4e2d\u6587\u6458\u8981:"))


def is_body_start(text: str) -> bool:
    compact = compact_text(text)
    stripped = str(text or "").strip()
    return bool(
        compact.startswith("1\u5f15\u8a00")
        or compact.startswith("1\u7eea\u8bba")
        or compact.startswith("1\u7dd2\u8ad6")
        or re.match(r"^\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u7ae0", stripped)
        or re.match(r"^1[\s\u3000]+", stripped)
    )


def paragraph_has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(qn("val"), "") if style is not None else ""


def is_toc_style_id(style_id: str) -> bool:
    normalized = (style_id or "").strip().upper()
    return normalized.startswith("TOC") or normalized in {"13", "16", "17"}


def paragraph_is_body_start(paragraph: ET.Element) -> bool:
    # Static TOC cache rows can look like "1 绪论1" and may not contain a
    # literal tab after other repair passes. Never treat TOC-styled paragraphs
    # as the body opener for section/page-number repair.
    return is_body_start(paragraph_text(paragraph)) and not paragraph_has_tab(paragraph) and not is_toc_style_id(paragraph_style_id(paragraph))


def ensure_pg_num_type(sect_pr: ET.Element, *, fmt: str, start: str | None) -> bool:
    pg = sect_pr.find("./w:pgNumType", NS)
    changed = False
    if pg is None:
        pg = ET.Element(qn("pgNumType"))
        sect_pr.insert(0, pg)
        changed = True
    if pg.get(qn("fmt")) != fmt:
        pg.set(qn("fmt"), fmt)
        changed = True
    if start is None:
        if qn("start") in pg.attrib:
            del pg.attrib[qn("start")]
            changed = True
    elif pg.get(qn("start")) != start:
        pg.set(qn("start"), start)
        changed = True
    return changed


def child_local_name(node: ET.Element) -> str:
    return node.tag.rsplit("}", 1)[-1] if "}" in node.tag else node.tag


def normalize_sect_pr_child_order(sect_pr: ET.Element) -> bool:
    children = list(sect_pr)
    ordered = sorted(
        enumerate(children),
        key=lambda item: (SECTION_CHILD_ORDER.get(child_local_name(item[1]), 10_000), item[0]),
    )
    new_children = [child for _idx, child in ordered]
    if new_children == children:
        return False
    for child in children:
        sect_pr.remove(child)
    for child in new_children:
        sect_pr.append(child)
    return True


def first_child(parent: ET.Element, local: str) -> ET.Element | None:
    return parent.find(f"./w:{local}", NS)


def replace_or_append_child(parent: ET.Element, replacement: ET.Element) -> None:
    local = child_local_name(replacement)
    for idx, child in enumerate(list(parent)):
        if child_local_name(child) == local:
            parent.remove(child)
            parent.insert(idx, deepcopy(replacement))
            return
    parent.append(deepcopy(replacement))


def xml_signature(node: ET.Element | None) -> bytes:
    return b"" if node is None else ET.tostring(node, encoding="utf-8")


def sync_frontmatter_section_layout(sect_pr: ET.Element, reference: ET.Element) -> list[str]:
    """Copy page geometry and visible footer/header owners from the body section."""
    changed: list[str] = []
    for local in ("pgSz", "pgMar", "cols", "docGrid"):
        ref_child = first_child(reference, local)
        if ref_child is None:
            continue
        current = first_child(sect_pr, local)
        if xml_signature(current) != xml_signature(ref_child):
            replace_or_append_child(sect_pr, ref_child)
            changed.append(local)

    for local in ("headerReference", "footerReference"):
        existing = sect_pr.findall(f"./w:{local}", NS)
        reference_nodes = reference.findall(f"./w:{local}", NS)
        if existing or not reference_nodes:
            continue
        insert_at = 0
        for ref_child in reference_nodes:
            sect_pr.insert(insert_at, deepcopy(ref_child))
            insert_at += 1
        changed.append(local)
    return changed


def ensure_child(parent: ET.Element, local: str, *, first: bool = False) -> ET.Element:
    child = parent.find(f"./w:{local}", NS)
    if child is not None:
        return child
    child = ET.Element(qn(local))
    if first:
        parent.insert(0, child)
    else:
        parent.append(child)
    return child


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def remove_page_break_before(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return False
    node = ppr.find("./w:pageBreakBefore", NS)
    if node is None:
        return False
    ppr.remove(node)
    return True


def body_paragraphs(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return [node for node in list(body) if node.tag == qn("p")]


def final_body_sect_pr(root: ET.Element) -> ET.Element:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    sect_pr = body.find("./w:sectPr", NS)
    if sect_pr is None:
        raise RuntimeError("body final sectPr missing; refusing page-number repair")
    return sect_pr


def all_sect_pr(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        return []
    sections: list[ET.Element] = []
    for paragraph in body_paragraphs(root):
        sect_pr = paragraph.find("./w:pPr/w:sectPr", NS)
        if sect_pr is not None:
            sections.append(sect_pr)
    body_sect = body.find("./w:sectPr", NS)
    if body_sect is not None:
        sections.append(body_sect)
    return sections


def repair_document_xml(document_xml: bytes, *, roman_fmt: str = "lowerRoman") -> tuple[bytes, list[dict[str, object]]]:
    root = ET.fromstring(document_xml)
    paragraphs = body_paragraphs(root)
    zh_abs_idx = next((idx for idx, p in enumerate(paragraphs) if is_zh_abstract_title(paragraph_text(p))), None)
    body_idx = next((idx for idx, p in enumerate(paragraphs) if paragraph_is_body_start(p)), None)
    if zh_abs_idx is None:
        raise RuntimeError("Chinese abstract title not found; refusing page-number repair")
    if body_idx is None:
        raise RuntimeError("body start heading not found; refusing page-number repair")
    if body_idx <= zh_abs_idx:
        raise RuntimeError("body start precedes abstract; refusing page-number repair")

    body_layout_reference = deepcopy(final_body_sect_pr(root))
    changed: list[dict[str, object]] = []
    roman_section_seen = False
    for idx, paragraph in enumerate(paragraphs):
        ppr = paragraph.find("./w:pPr", NS)
        sect_pr = ppr.find("./w:sectPr", NS) if ppr is not None else None
        if sect_pr is None:
            continue
        if zh_abs_idx <= idx < body_idx:
            start = "1" if not roman_section_seen else None
            layout_restored = sync_frontmatter_section_layout(sect_pr, body_layout_reference)
            pg_num_changed = ensure_pg_num_type(sect_pr, fmt=roman_fmt, start=start)
            if pg_num_changed or layout_restored:
                changed.append(
                    {
                        "paragraph_index": idx,
                        "surface": "front_matter_page_numbers",
                        "fmt": roman_fmt,
                        "start": start or "continue",
                        "layout_source": "body_final_sectPr",
                        "layout_restored": layout_restored,
                        "text": paragraph_text(paragraph)[:80],
                    }
                )
            roman_section_seen = True

    if not roman_section_seen:
        frontmatter_end_idx = body_idx - 1
        frontmatter_end = paragraphs[frontmatter_end_idx]
        ppr = ensure_ppr(frontmatter_end)
        sect_pr = ppr.find("./w:sectPr", NS)
        if sect_pr is None:
            sect_pr = deepcopy(body_layout_reference)
            ppr.append(sect_pr)
        else:
            sync_frontmatter_section_layout(sect_pr, body_layout_reference)
        type_node = ensure_child(sect_pr, "type", first=True)
        type_node.set(qn("val"), "nextPage")
        ensure_pg_num_type(sect_pr, fmt=roman_fmt, start="1")
        removed_body_page_break = remove_page_break_before(paragraphs[body_idx])
        changed.append(
            {
                "paragraph_index": frontmatter_end_idx,
                "surface": "front_matter_page_numbers",
                "fmt": roman_fmt,
                "start": "1",
                "text": paragraph_text(frontmatter_end)[:80],
                "inserted_section_break_before_body": True,
                "removed_body_heading_page_break_before": removed_body_page_break,
                "layout_source": "body_final_sectPr",
                "layout_restored": ["pgSz", "pgMar", "cols", "docGrid"],
            }
        )

    body_decimal_section_seen = False
    for idx, paragraph in enumerate(paragraphs):
        if idx < body_idx:
            continue
        ppr = paragraph.find("./w:pPr", NS)
        sect_pr = ppr.find("./w:sectPr", NS) if ppr is not None else None
        if sect_pr is None:
            continue
        start = "1" if not body_decimal_section_seen else None
        if ensure_pg_num_type(sect_pr, fmt="decimal", start=start):
            changed.append(
                {
                    "paragraph_index": idx,
                    "surface": "body_page_numbers",
                    "fmt": "decimal",
                    "start": start or "continue",
                    "text": paragraph_text(paragraph)[:80],
                }
            )
        body_decimal_section_seen = True

    body_sect = final_body_sect_pr(root)
    final_start = "1" if not body_decimal_section_seen else None
    if ensure_pg_num_type(body_sect, fmt="decimal", start=final_start):
        changed.append(
            {
                "paragraph_index": "body-sectPr",
                "surface": "body_page_numbers",
                "fmt": "decimal",
                "start": final_start or "continue",
                "text": "body final section",
            }
        )
    for idx, sect_pr in enumerate(all_sect_pr(root), start=1):
        if normalize_sect_pr_child_order(sect_pr):
            changed.append(
                {
                    "paragraph_index": f"sectPr-order-{idx}",
                    "surface": "section_property_order",
                    "fmt": "not-applicable",
                    "start": "not-applicable",
                    "text": "normalized w:sectPr child order for OpenXML schema compatibility",
                }
            )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def write_docx_with_document_xml(input_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    with tempfile.TemporaryDirectory() as tmp_name:
        temp_output = Path(tmp_name) / output_docx.name
        with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
        if output_docx.exists():
            output_docx.unlink()
        shutil.move(str(temp_output), output_docx)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument(
        "--roman-fmt",
        default="lowerRoman",
        choices=("lowerRoman", "upperRoman"),
        help="Front-matter page-number format; use upperRoman when the locked template renders I, II, III.",
    )
    args = parser.parse_args()

    input_docx = Path(args.input_docx).resolve()
    output_docx = Path(args.output_docx).resolve()
    report = Path(args.report).resolve()
    if input_docx == output_docx:
        raise RuntimeError("output DOCX must be a new review-copy path")
    with zipfile.ZipFile(input_docx, "r") as zf:
        repaired_xml, changed = repair_document_xml(zf.read("word/document.xml"), roman_fmt=args.roman_fmt)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    report.parent.mkdir(parents=True, exist_ok=True)
    write_docx_with_document_xml(input_docx, output_docx, repaired_xml)
    payload = {
        "schema": "graduation-project-builder.front-matter-page-numbering-repair.v2",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "roman_fmt": args.roman_fmt,
        "changed_sections": changed,
        "repair_verdict": "pass" if changed else "pass-no-change-needed",
    }
    report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
