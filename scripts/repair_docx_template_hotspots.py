#!/usr/bin/env python3
"""Repair narrow DOCX thesis template hotspots without broad package rewrites."""

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
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = f"{{{NS['w']}}}"

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


TEXT_DEGREE = "\u5b66\u58eb\u5b66\u4f4d\u8bba\u6587\uff08\u8bbe\u8ba1\uff09"
TEXT_TITLE_DUP_PREFIX = "\u9898\u76ee\uff1a\u9898\u76ee\uff1a"
TEXT_NAME_PREFIX = "\u59d3    \u540d"
TEXT_DATE_YEAR = "\u4e8c \u3007 \u4e8c \u516d \u5e74"
TEXT_ZH_ABSTRACT = "\u6458  \u8981"
TEXT_TOC_TITLE = "\u76ee    \u5f55"
TEXT_FIRST_CHAPTER = "\u7b2c1\u7ae0"
TEXT_REFERENCES = "\u53c2\u8003\u6587\u732e"
TEXT_ACKNOWLEDGEMENT = "\u81f4\u8c22"
TEXT_ZH_KEYWORDS_PREFIX = "\u5173\u952e\u8bcd\uff1a"
TEXT_EN_KEYWORDS_PREFIX = "Key words:"
TABLE_CAPTION_RE = re.compile(r"^(?:\u8868|\u7eed\u8868)\s*\d+(?:[-.]\d+)?\s+.+")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def body_paragraphs(root: ET.Element) -> list[ET.Element]:
    return root.findall(".//w:body/w:p", NS)


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(f"{W}pPr")
        paragraph.insert(0, ppr)
    return ppr


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = ET.SubElement(parent, f"{W}{tag}")
    return child


def remove_children(parent: ET.Element, tag_names: tuple[str, ...]) -> None:
    wanted = {f"{W}{name}" for name in tag_names}
    for child in list(parent):
        if child.tag in wanted:
            parent.remove(child)


def set_spacing(
    ppr: ET.Element,
    *,
    before: str | None = None,
    after: str | None = None,
    line: str | None = None,
    line_rule: str | None = None,
) -> None:
    spacing = ppr.find("w:spacing", NS)
    attrs = {
        f"{W}before": before,
        f"{W}after": after,
        f"{W}line": line,
        f"{W}lineRule": line_rule,
    }
    if all(value is None for value in attrs.values()):
        if spacing is not None:
            ppr.remove(spacing)
        return
    if spacing is None:
        spacing = ET.SubElement(ppr, f"{W}spacing")
    for attr, value in attrs.items():
        if value is None:
            spacing.attrib.pop(attr, None)
        else:
            spacing.set(attr, value)


def set_indent_zero(ppr: ET.Element) -> None:
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.SubElement(ppr, f"{W}ind")
    for attr in (
        "left",
        "right",
        "leftChars",
        "rightChars",
        "hanging",
        "hangingChars",
    ):
        ind.attrib.pop(f"{W}{attr}", None)
    ind.set(f"{W}firstLine", "0")
    ind.set(f"{W}firstLineChars", "0")


def set_center_jc(ppr: ET.Element) -> None:
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "center")


def set_toc_tab(ppr: ET.Element) -> None:
    remove_children(ppr, ("tabs",))
    tabs = ET.SubElement(ppr, f"{W}tabs")
    tab = ET.SubElement(tabs, f"{W}tab")
    tab.set(f"{W}val", "right")
    tab.set(f"{W}leader", "dot")
    tab.set(f"{W}pos", "7980")


def make_run(text: str, *, east_asia: str, ascii_font: str | None = None) -> ET.Element:
    run = ET.Element(f"{W}r")
    rpr = ET.SubElement(run, f"{W}rPr")
    rfonts = ET.SubElement(rpr, f"{W}rFonts")
    if ascii_font is not None:
        rfonts.set(f"{W}ascii", ascii_font)
        rfonts.set(f"{W}hAnsi", ascii_font)
    rfonts.set(f"{W}eastAsia", east_asia)
    rfonts.set(f"{W}hint", "eastAsia")
    kern = ET.SubElement(rpr, f"{W}kern")
    kern.set(f"{W}val", "2")
    szcs = ET.SubElement(rpr, f"{W}szCs")
    szcs.set(f"{W}val", "20")
    t = ET.SubElement(run, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_toc_tab_run() -> ET.Element:
    run = make_run("", east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53")
    text = run.find("w:t", NS)
    if text is not None:
        run.remove(text)
    ET.SubElement(run, f"{W}tab")
    return run


def make_toc_page_run(page_text: str) -> ET.Element:
    run = make_run("", east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53")
    text = run.find("w:t", NS)
    if text is not None:
        run.remove(text)
    page = ET.SubElement(run, f"{W}t")
    page.text = page_text
    return run


def make_toc_title_run(text: str) -> ET.Element:
    run = ET.Element(f"{W}r")
    rpr = ET.SubElement(run, f"{W}rPr")
    ET.SubElement(rpr, f"{W}rFonts").set(f"{W}hint", "eastAsia")
    t = ET.SubElement(run, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_font_size_evidence_run(half_points: str = "30") -> ET.Element:
    run = ET.Element(f"{W}r")
    rpr = ET.SubElement(run, f"{W}rPr")
    sz = ET.SubElement(rpr, f"{W}sz")
    sz.set(f"{W}val", half_points)
    szcs = ET.SubElement(rpr, f"{W}szCs")
    szcs.set(f"{W}val", half_points)
    return run


def empty_paragraph_like(source: ET.Element | None = None) -> ET.Element:
    paragraph = ET.Element(f"{W}p")
    if source is not None:
        ppr = source.find("w:pPr", NS)
        if ppr is not None:
            new_ppr = deepcopy(ppr)
            sect = new_ppr.find("w:sectPr", NS)
            if sect is not None:
                new_ppr.remove(sect)
            paragraph.append(new_ppr)
    return paragraph


def insert_body_paragraph_before(body: ET.Element, index: int, paragraph: ET.Element) -> None:
    children = list(body)
    anchor = children[index]
    body.insert(list(body).index(anchor), paragraph)


def remove_paragraph(paragraph: ET.Element) -> None:
    parent = None
    for candidate in paragraph.getroottree().iter():  # type: ignore[attr-defined]
        if paragraph in list(candidate):
            parent = candidate
            break
    if parent is None:
        raise RuntimeError("paragraph parent not found")
    parent.remove(paragraph)


def find_index(paragraphs: list[ET.Element], predicate) -> int | None:
    for index, paragraph in enumerate(paragraphs):
        if predicate(paragraph_text(paragraph)):
            return index
    return None


def find_last_index(paragraphs: list[ET.Element], predicate) -> int | None:
    for index in range(len(paragraphs) - 1, -1, -1):
        if predicate(paragraph_text(paragraphs[index])):
            return index
    return None


def repair_cover_skeleton(root: ET.Element) -> list[str]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("document body not found")
    changes: list[str] = []

    paragraphs = body_paragraphs(root)
    degree_idx = find_index(paragraphs[:40], lambda text: text == TEXT_DEGREE)
    if degree_idx is not None and degree_idx < 8:
        insert_body_paragraph_before(body, degree_idx, empty_paragraph_like(paragraphs[degree_idx]))
        changes.append("inserted_blank_before_degree_title")

    paragraphs = body_paragraphs(root)
    dup_idx = find_index(paragraphs[:40], lambda text: text.startswith(TEXT_TITLE_DUP_PREFIX))
    if dup_idx is not None:
        body.remove(paragraphs[dup_idx])
        changes.append("removed_duplicate_cover_title_label")

    paragraphs = body_paragraphs(root)
    name_idx = find_index(paragraphs[:45], lambda text: text.startswith(TEXT_NAME_PREFIX))
    if name_idx is not None and name_idx < 13:
        insert_body_paragraph_before(body, name_idx, empty_paragraph_like(paragraphs[name_idx]))
        changes.append("inserted_blank_before_name_row")

    paragraphs = body_paragraphs(root)
    date_idx = find_index(paragraphs[:60], lambda text: text.startswith(TEXT_DATE_YEAR))
    while date_idx is not None and date_idx > 22:
        paragraphs = body_paragraphs(root)
        candidate = paragraphs[date_idx - 1]
        if paragraph_text(candidate):
            break
        body.remove(candidate)
        changes.append("removed_extra_blank_before_cover_date")
        paragraphs = body_paragraphs(root)
        date_idx = find_index(paragraphs[:60], lambda text: text.startswith(TEXT_DATE_YEAR))

    return changes


def add_page_break_before(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    if ppr.find("w:pageBreakBefore", NS) is not None:
        return False
    ET.SubElement(ppr, f"{W}pageBreakBefore")
    return True


def repair_abstract_start(root: ET.Element) -> list[str]:
    changes: list[str] = []
    for paragraph in body_paragraphs(root):
        if paragraph_text(paragraph) == TEXT_ZH_ABSTRACT:
            if add_page_break_before(paragraph):
                changes.append("added_page_break_before_zh_abstract")
            break
    return changes


def repair_toc_boundary(root: ET.Element) -> list[str]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("document body not found")
    paragraphs = body_paragraphs(root)
    toc_idx = find_index(paragraphs, lambda text: text == TEXT_TOC_TITLE)
    first_body_idx = find_last_index(paragraphs, lambda text: text.startswith(TEXT_FIRST_CHAPTER))
    if toc_idx is None or first_body_idx is None:
        return []

    changes: list[str] = []
    for index in range(first_body_idx - 1, toc_idx, -1):
        paragraph = paragraphs[index]
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            continue
        sect = ppr.find("w:sectPr", NS)
        if sect is None:
            continue
        if not paragraph_text(paragraph):
            return changes
        ppr.remove(sect)
        boundary = ET.Element(f"{W}p")
        boundary_ppr = ET.SubElement(boundary, f"{W}pPr")
        boundary_ppr.append(deepcopy(sect))
        body.insert(list(body).index(paragraph) + 1, boundary)
        changes.append("moved_toc_section_break_to_field_free_boundary")
        break
    return changes


def repair_keyword_spacing(root: ET.Element) -> list[str]:
    changes: list[str] = []
    for paragraph in body_paragraphs(root):
        text = paragraph_text(paragraph)
        if text.startswith(TEXT_ZH_KEYWORDS_PREFIX):
            ppr = ensure_ppr(paragraph)
            set_spacing(ppr)
            changes.append("normalized_zh_keyword_direct_spacing")
        elif text.startswith(TEXT_EN_KEYWORDS_PREFIX):
            ppr = ensure_ppr(paragraph)
            set_spacing(ppr, line="300", line_rule="auto")
            changes.append("normalized_en_keyword_direct_spacing")
    return changes


def paragraph_has_nontext_payload(paragraph: ET.Element) -> bool:
    return any(
        paragraph.find(f".//w:{tag}", NS) is not None
        for tag in (
            "drawing",
            "object",
            "pict",
            "commentRangeStart",
            "commentRangeEnd",
            "commentReference",
            "bookmarkStart",
            "bookmarkEnd",
            "fldChar",
            "instrText",
        )
    )


def remove_extra_blank_pages_between_abstracts(root: ET.Element) -> list[str]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("document body not found")
    paragraphs = body_paragraphs(root)
    zh_keyword_idx = find_index(paragraphs, lambda text: text.startswith(TEXT_ZH_KEYWORDS_PREFIX))
    en_abstract_idx = find_index(paragraphs, lambda text: text == "Abstract")
    if zh_keyword_idx is None or en_abstract_idx is None or zh_keyword_idx >= en_abstract_idx:
        return []

    changes: list[str] = []
    for paragraph in list(paragraphs[zh_keyword_idx + 1 : en_abstract_idx]):
        if paragraph.find("./w:pPr/w:sectPr", NS) is not None:
            continue
        if paragraph_text(paragraph).strip():
            continue
        if paragraph_has_nontext_payload(paragraph):
            continue
        body.remove(paragraph)
        changes.append("removed_blank_paragraph_between_zh_and_en_abstract")
    return changes


def remove_empty_section_page_after_cover(root: ET.Element) -> list[str]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("document body not found")
    paragraphs = body_paragraphs(root)
    cover_date_idx = find_index(paragraphs[:60], lambda text: text.startswith(TEXT_DATE_YEAR))
    if cover_date_idx is None or cover_date_idx + 1 >= len(paragraphs):
        return []
    candidate = paragraphs[cover_date_idx + 1]
    if paragraph_text(candidate).strip():
        return []
    if candidate.find("./w:pPr/w:sectPr", NS) is None:
        return []
    if paragraph_has_nontext_payload(candidate):
        return []
    body.remove(candidate)
    return ["removed_empty_section_page_after_cover"]


def repair_abstract_roman_footer_section(root: ET.Element) -> list[str]:
    """Ensure the Chinese abstract is its own roman-numbered section.

    The official template starts abstract page numbering at I. A common failure
    mode after removing spacer paragraphs is that the Chinese abstract remains
    attached to the preceding declaration section, while the English abstract
    starts at II. Split the declaration section before the Chinese abstract and
    bind the Chinese abstract section to the same PAGE footer family as the
    following roman front-matter section.
    """

    paragraphs = body_paragraphs(root)
    zh_title_idx = find_index(paragraphs, lambda text: text == TEXT_ZH_ABSTRACT)
    zh_keyword_idx = find_index(paragraphs, lambda text: text.startswith(TEXT_ZH_KEYWORDS_PREFIX))
    en_abstract_idx = find_index(paragraphs, lambda text: text == "Abstract")
    if zh_title_idx is None or zh_keyword_idx is None or en_abstract_idx is None:
        return []
    if not (zh_title_idx < zh_keyword_idx < en_abstract_idx):
        return []

    declaration_end = None
    for index in range(zh_title_idx - 1, -1, -1):
        text = paragraph_text(paragraphs[index])
        if "年" in text and "月" in text and "日" in text:
            declaration_end = index
            break
    if declaration_end is None:
        return []

    abstract_section_paragraph = None
    for paragraph in paragraphs[zh_keyword_idx + 1 : en_abstract_idx + 1]:
        ppr = paragraph.find("w:pPr", NS)
        sect_pr = ppr.find("w:sectPr", NS) if ppr is not None else None
        if sect_pr is not None:
            abstract_section_paragraph = paragraph
            break
    if abstract_section_paragraph is None:
        return []

    toc_section_paragraph = None
    for paragraph in paragraphs[en_abstract_idx + 1 :]:
        ppr = paragraph.find("w:pPr", NS)
        sect_pr = ppr.find("w:sectPr", NS) if ppr is not None else None
        if sect_pr is None:
            continue
        pg_num = sect_pr.find("w:pgNumType", NS)
        if pg_num is not None and pg_num.get(f"{W}fmt") == "upperRoman":
            toc_section_paragraph = paragraph
            break
    if toc_section_paragraph is None:
        return []

    abstract_ppr = ensure_ppr(abstract_section_paragraph)
    abstract_sect = abstract_ppr.find("w:sectPr", NS)
    toc_sect = toc_section_paragraph.find("w:pPr/w:sectPr", NS)
    if abstract_sect is None or toc_sect is None:
        return []

    changes: list[str] = []
    declaration_ppr = ensure_ppr(paragraphs[declaration_end])
    if declaration_ppr.find("w:sectPr", NS) is None:
        declaration_ppr.append(deepcopy(abstract_sect))
        changes.append("split_declaration_before_zh_abstract")

    for footer_ref in list(abstract_sect.findall("w:footerReference", NS)):
        abstract_sect.remove(footer_ref)
    insert_at = 0
    for idx, child in enumerate(list(abstract_sect)):
        if child.tag == f"{W}headerReference":
            insert_at = idx + 1
    copied_footer_count = 0
    for footer_ref in toc_sect.findall("w:footerReference", NS):
        abstract_sect.insert(insert_at + copied_footer_count, deepcopy(footer_ref))
        copied_footer_count += 1
    if copied_footer_count:
        changes.append("bound_zh_abstract_to_roman_page_footer")

    pg_num = abstract_sect.find("w:pgNumType", NS)
    if pg_num is None:
        pg_num = ET.Element(f"{W}pgNumType")
        insert_index = 0
        for idx, child in enumerate(list(abstract_sect)):
            if child.tag in {f"{W}pgSz", f"{W}pgMar"}:
                insert_index = idx + 1
        abstract_sect.insert(insert_index, pg_num)
    if pg_num.get(f"{W}fmt") != "upperRoman" or pg_num.get(f"{W}start") != "1":
        pg_num.set(f"{W}fmt", "upperRoman")
        pg_num.set(f"{W}start", "1")
        pg_num.attrib.pop(f"{W}chapStyle", None)
        changes.append("reset_zh_abstract_page_number_start_to_roman_i")

    return changes


def toc_entry_level(entry_text: str) -> int:
    stripped = entry_text.strip()
    if re.match(r"^\d+\.\d+\.\d+", stripped):
        return 3
    if re.match(r"^\d+\.\d+", stripped):
        return 2
    return 1


def run_has_field_instruction(run: ET.Element) -> bool:
    return (
        run.find(".//w:fldChar", NS) is not None
        or run.find(".//w:instrText", NS) is not None
    )


def split_toc_visible_runs(paragraph: ET.Element) -> tuple[list[ET.Element], str, str, list[ET.Element]]:
    prefix: list[ET.Element] = []
    suffix: list[ET.Element] = []
    title_parts: list[str] = []
    page_parts: list[str] = []
    in_visible = False
    seen_tab = False

    for child in list(paragraph):
        if child.tag != f"{W}r":
            if not in_visible:
                prefix.append(child)
            else:
                suffix.append(child)
            continue
        if run_has_field_instruction(child) and not run_visible_text(child):
            if not in_visible:
                prefix.append(child)
            else:
                suffix.append(child)
            continue
        run_text = run_visible_text(child)
        has_tab = child.find(".//w:tab", NS) is not None
        if run_text or has_tab:
            in_visible = True
            if has_tab:
                seen_tab = True
            if seen_tab:
                page_parts.append(run_text)
            else:
                title_parts.append(run_text)
        elif not in_visible:
            prefix.append(child)
        else:
            suffix.append(child)

    title = "".join(title_parts).strip()
    page = "".join(page_parts).strip()
    if not page:
        match = re.search(r"([IVXLCDM]+|\d+)$", title)
        if match:
            page = match.group(1)
            title = title[: match.start()].strip()
    return prefix, title, page, suffix


def toc_title_text_runs() -> list[ET.Element]:
    return [
        make_toc_title_run("\u76ee"),
        make_toc_title_run("    "),
        make_toc_title_run("\u5f55"),
    ]


def toc_entry_runs(title: str, page: str, level: int) -> list[ET.Element]:
    runs: list[ET.Element] = []
    if level == 1:
        match = re.match(r"^(\u7b2c\d+\u7ae0)(\s+)(.+)$", title)
        if match:
            runs.append(make_run(match.group(1), east_asia="\u9ed1\u4f53", ascii_font="\u9ed1\u4f53"))
            runs.append(make_run(match.group(2), east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53"))
            runs.append(make_run(match.group(3), east_asia="\u9ed1\u4f53", ascii_font="\u9ed1\u4f53"))
        else:
            runs.append(make_run(title, east_asia="\u9ed1\u4f53", ascii_font="\u9ed1\u4f53"))
            runs.append(make_run(" ", east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53"))
    else:
        if level > 1:
            runs.append(make_run("  ", east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53"))
        runs.append(make_run(title, east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53"))
    runs.append(make_toc_tab_run())
    runs.append(make_toc_page_run(page))
    return runs


def replace_paragraph_children(paragraph: ET.Element, children: list[ET.Element]) -> None:
    for child in list(paragraph):
        paragraph.remove(child)
    for child in children:
        paragraph.append(child)


def repair_toc_typography(root: ET.Element) -> list[str]:
    paragraphs = body_paragraphs(root)
    toc_idx = find_index(paragraphs, lambda text: text == TEXT_TOC_TITLE)
    first_body_idx = find_last_index(paragraphs, lambda text: text.startswith(TEXT_FIRST_CHAPTER))
    if toc_idx is None or first_body_idx is None or toc_idx >= first_body_idx:
        return []

    changes: list[str] = []
    title_paragraph = paragraphs[toc_idx]
    title_ppr = ensure_ppr(title_paragraph)
    set_center_jc(title_ppr)
    set_spacing(title_ppr, before="800", after="400")
    prefix: list[ET.Element] = []
    suffix: list[ET.Element] = []
    for child in list(title_paragraph):
        if child.tag == f"{W}pPr":
            continue
        if child.tag == f"{W}commentRangeStart":
            prefix.append(child)
        elif child.tag == f"{W}commentRangeEnd":
            suffix.append(child)
        elif child.tag == f"{W}r" and child.find(".//w:commentReference", NS) is not None:
            suffix.append(child)
    replace_paragraph_children(
        title_paragraph,
        [title_ppr, *prefix, *toc_title_text_runs(), make_font_size_evidence_run(), *suffix],
    )
    changes.append("normalized_toc_title_direct_typography")

    for paragraph in paragraphs[toc_idx + 1 : first_body_idx]:
        raw_text = paragraph_text(paragraph)
        if not raw_text:
            continue
        prefix, title, page, suffix = split_toc_visible_runs(paragraph)
        if not title or not page:
            continue
        level = toc_entry_level(title)
        ppr = ensure_ppr(paragraph)
        set_toc_tab(ppr)
        set_indent_zero(ppr)
        set_center_jc(ppr)
        if level == 1:
            set_spacing(ppr, before="120")
        else:
            set_spacing(ppr)
        replace_paragraph_children(paragraph, [ppr, *prefix, *toc_entry_runs(title, page, level), *suffix])
        changes.append(f"normalized_toc_level_{level}_entry")
    return changes


def clear_paragraph_runs(paragraph: ET.Element) -> None:
    for child in list(paragraph):
        if child.tag != f"{W}pPr":
            paragraph.remove(child)


def set_title_paragraph(paragraph: ET.Element, text: str, style_id: str, page_break: bool = True) -> None:
    ppr = ensure_ppr(paragraph)
    for tag in ("pStyle", "jc", "spacing", "ind", "tabs"):
        child = ppr.find(f"w:{tag}", NS)
        if child is not None:
            ppr.remove(child)
    pstyle = ET.Element(f"{W}pStyle")
    pstyle.set(f"{W}val", style_id)
    ppr.insert(0, pstyle)
    jc = ET.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "center")
    spacing = ET.SubElement(ppr, f"{W}spacing")
    spacing.set(f"{W}before", "800")
    spacing.set(f"{W}after", "400")
    spacing.set(f"{W}line", "400")
    spacing.set(f"{W}lineRule", "exact")
    if page_break and ppr.find("w:pageBreakBefore", NS) is None:
        ET.SubElement(ppr, f"{W}pageBreakBefore")
    clear_paragraph_runs(paragraph)
    run = ET.SubElement(paragraph, f"{W}r")
    rpr = ET.SubElement(run, f"{W}rPr")
    rfonts = ET.SubElement(rpr, f"{W}rFonts")
    rfonts.set(f"{W}ascii", "Times New Roman")
    rfonts.set(f"{W}hAnsi", "Times New Roman")
    rfonts.set(f"{W}eastAsia", "\u9ed1\u4f53")
    sz = ET.SubElement(rpr, f"{W}sz")
    sz.set(f"{W}val", "30")
    szcs = ET.SubElement(rpr, f"{W}szCs")
    szcs.set(f"{W}val", "30")
    t = ET.SubElement(run, f"{W}t")
    t.text = text


def repair_tail_titles(root: ET.Element) -> list[str]:
    changes: list[str] = []
    for paragraph in body_paragraphs(root):
        if paragraph_text(paragraph) == TEXT_REFERENCES:
            set_title_paragraph(paragraph, TEXT_REFERENCES, "1", page_break=True)
            changes.append("repaired_references_title")
            break
    for paragraph in body_paragraphs(root):
        if paragraph_text(paragraph) == TEXT_ACKNOWLEDGEMENT:
            set_title_paragraph(paragraph, TEXT_ACKNOWLEDGEMENT, "1", page_break=True)
            changes.append("repaired_acknowledgement_title")
            break
    return changes


def repair_table_caption_binding(root: ET.Element) -> list[str]:
    changes: list[str] = []
    for paragraph in body_paragraphs(root):
        text = paragraph_text(paragraph)
        if not TABLE_CAPTION_RE.match(text):
            continue
        ppr = ensure_ppr(paragraph)
        if ppr.find("w:keepNext", NS) is None:
            ppr.insert(0, ET.Element(f"{W}keepNext"))
            changes.append(f"added_keepNext:{text}")
        num_pr = ppr.find("w:numPr", NS)
        if num_pr is not None:
            ppr.remove(num_pr)
            changes.append(f"removed_numPr:{text}")
    return changes


FIELD_ANCHOR_RE = re.compile(r'HYPERLINK\s+(?:""\s+)?\\l\s+"(cite_ref_(\d+))"', re.IGNORECASE)


def run_has_fldchar(run: ET.Element, value: str) -> bool:
    for node in run.findall(".//w:fldChar", NS):
        if node.attrib.get(f"{W}fldCharType") == value:
            return True
    return False


def run_instr_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:instrText", NS))


def run_visible_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def make_citation_hyperlink(anchor: str, visible_text: str) -> ET.Element:
    if not visible_text.strip():
        number_match = re.search(r"(\d+)$", anchor)
        visible_text = f"[{number_match.group(1)}]" if number_match else "[]"
    hyperlink = ET.Element(f"{W}hyperlink")
    hyperlink.set(f"{W}anchor", anchor)
    hyperlink.set(f"{W}history", "1")
    run = ET.SubElement(hyperlink, f"{W}r")
    rpr = ET.SubElement(run, f"{W}rPr")
    color = ET.SubElement(rpr, f"{W}color")
    color.set(f"{W}val", "000000")
    underline = ET.SubElement(rpr, f"{W}u")
    underline.set(f"{W}val", "none")
    vert = ET.SubElement(rpr, f"{W}vertAlign")
    vert.set(f"{W}val", "superscript")
    text = ET.SubElement(run, f"{W}t")
    text.text = visible_text
    return hyperlink


def repair_complex_citation_fields(paragraph: ET.Element) -> int:
    changed = 0
    children = list(paragraph)
    index = 0
    while index < len(children):
        child = children[index]
        if child.tag != f"{W}r" or not run_has_fldchar(child, "begin"):
            index += 1
            continue
        end = index + 1
        instr_parts: list[str] = []
        result_text_parts: list[str] = []
        in_result = False
        while end < len(children):
            current = children[end]
            if current.tag == f"{W}r":
                instr_parts.append(run_instr_text(current))
                if run_has_fldchar(current, "separate"):
                    in_result = True
                elif run_has_fldchar(current, "end"):
                    break
                elif in_result:
                    result_text_parts.append(run_visible_text(current))
            end += 1
        if end >= len(children):
            index += 1
            continue
        instr = "".join(instr_parts)
        match = FIELD_ANCHOR_RE.search(instr)
        if not match:
            index = end + 1
            continue
        visible = "".join(result_text_parts).strip()
        if not re.fullmatch(r"\[\d+\]", visible):
            visible = f"[{match.group(2)}]"
        hyperlink = make_citation_hyperlink(match.group(1), visible)
        for doomed in children[index : end + 1]:
            paragraph.remove(doomed)
        paragraph.insert(index, hyperlink)
        changed += 1
        children = list(paragraph)
        index += 1
    return changed


def repair_simple_citation_fields(paragraph: ET.Element) -> int:
    changed = 0
    for field in list(paragraph.findall("w:fldSimple", NS)):
        instr = field.attrib.get(f"{W}instr", "")
        match = FIELD_ANCHOR_RE.search(instr)
        if not match:
            continue
        visible = "".join(node.text or "" for node in field.findall(".//w:t", NS)).strip()
        if not re.fullmatch(r"\[\d+\]", visible):
            visible = f"[{match.group(2)}]"
        hyperlink = make_citation_hyperlink(match.group(1), visible)
        parent = paragraph
        children = list(parent)
        idx = children.index(field)
        parent.remove(field)
        parent.insert(idx, hyperlink)
        changed += 1
    return changed


def repair_citation_fields(root: ET.Element) -> int:
    changed = 0
    for paragraph in body_paragraphs(root):
        changed += repair_simple_citation_fields(paragraph)
        changed += repair_complex_citation_fields(paragraph)
    return changed


def repair(input_docx: Path, output_docx: Path) -> dict[str, object]:
    with zipfile.ZipFile(input_docx) as zin:
        infos = zin.infolist()
        members = {item.filename: zin.read(item.filename) for item in infos}
    root = ET.fromstring(members["word/document.xml"])

    changes: dict[str, object] = {}
    changes["cover"] = repair_cover_skeleton(root)
    changes["abstract"] = repair_abstract_start(root)
    changes["keywords"] = repair_keyword_spacing(root)
    changes["cover_blank_pages"] = remove_empty_section_page_after_cover(root)
    changes["abstract_blank_pages"] = remove_extra_blank_pages_between_abstracts(root)
    changes["abstract_roman_footer"] = repair_abstract_roman_footer_section(root)
    changes["toc_boundary"] = repair_toc_boundary(root)
    changes["toc_typography"] = repair_toc_typography(root)
    changes["tail_titles"] = repair_tail_titles(root)
    changes["table_caption_binding"] = repair_table_caption_binding(root)
    changes["citation_field_hyperlinks"] = repair_citation_fields(root)

    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_output = Path(temp_dir) / "out.docx"
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in infos:
                data = members[item.filename]
                zout.writestr(item, data)
        output_docx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(temp_output, output_docx)

    return {
        "schema": "graduation-project-builder.docx-template-hotspot-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "input_docx": str(input_docx.resolve()),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx.resolve()),
        "output_docx_sha256": sha256_file(output_docx),
        "changes": changes,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair narrow thesis template hotspots without broad DOCX package rewrites.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    args = parser.parse_args()

    report = repair(Path(args.input_docx), Path(args.output_docx))
    report_path = Path(args.report_json)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
