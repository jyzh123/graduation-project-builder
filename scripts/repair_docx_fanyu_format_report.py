#!/usr/bin/env python3
"""Repair DOCX surfaces reported by a Fanyu thesis format report.

The helper is intentionally bounded to recurring report-owned surfaces:
front-matter/TOC order, paragraph spacing, captions, table cells, reference
label spacing, running header rules, footer PAGE-field typography, and report
comment cleanup. It preserves package media, formulas, fields, bookmarks, and
body text outside the explicitly owned format surfaces.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"

W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
R = f"{{{R_NS}}}"
REL = f"{{{REL_NS}}}"
CT = f"{{{CT_NS}}}"
NS = {"w": W_NS, "m": M_NS, "r": R_NS, "rel": REL_NS, "ct": CT_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("m", M_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("mc", MC_NS)
ET.register_namespace("w14", W14_NS)
ET.register_namespace("w15", W15_NS)
ET.register_namespace("wp14", WP14_NS)
ET.register_namespace("", REL_NS)


ZH_ABSTRACT = "\u6458\u8981"
EN_ABSTRACT = "Abstract"
TOC_TITLE = "\u76ee\u5f55"
REFERENCES_TITLE = "\u53c2\u8003\u6587\u732e"
ACK_TITLE = "\u81f4\u8c22"
CONCLUSION_TITLE = "\u7ed3\u8bba"
SONGTI = "\u5b8b\u4f53"
HEITI = "\u9ed1\u4f53"
KAITI = "\u6977\u4f53"
HEADER_LEFT = "\u6c88\u9633\u79d1\u6280\u5b66\u9662\u5b66\u58eb\u5b66\u4f4d\u8bba\u6587"
FOOTER_SMALL_FIVE = "18"


def qn(local: str) -> str:
    return f"{W}{local}"


def rqn(local: str) -> str:
    return f"{R}{local}"


def mqn(local: str) -> str:
    return f"{M}{local}"


def ctqn(local: str) -> str:
    return f"{CT}{local}"


def relqn(local: str) -> str:
    return f"{REL}{local}"


def local_name(node: ET.Element) -> str:
    return node.tag.rsplit("}", 1)[-1] if "}" in node.tag else node.tag


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def text_of(element: ET.Element) -> str:
    chunks: list[str] = []
    def walk(node: ET.Element, in_run: bool = False) -> None:
        current_in_run = in_run or node.tag == qn("r")
        if node.tag == qn("t"):
            chunks.append(node.text or "")
        elif node.tag == qn("tab") and in_run:
            chunks.append("\t")
        for child_node in list(node):
            walk(child_node, current_in_run)

    walk(element)
    return "".join(chunks)


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def ensure_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        run.insert(0, rpr)
    return rpr


def child(parent: ET.Element, tag: str) -> ET.Element | None:
    return parent.find(f"./w:{tag}", NS)


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    existing = child(parent, tag)
    if existing is not None:
        return existing
    created = ET.Element(qn(tag))
    parent.append(created)
    return created


def remove_children_by_local(parent: ET.Element, names: set[str]) -> None:
    for item in list(parent):
        if local_name(item) in names:
            parent.remove(item)


def set_spacing(
    paragraph: ET.Element,
    *,
    before: str | None = None,
    after: str | None = None,
    before_lines: str | None = None,
    after_lines: str | None = None,
    line: str | None = None,
    line_rule: str | None = None,
) -> None:
    ppr = ensure_ppr(paragraph)
    spacing = child(ppr, "spacing")
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    for attr in ("before", "after", "beforeLines", "afterLines", "line", "lineRule"):
        spacing.attrib.pop(qn(attr), None)
    if before is not None:
        spacing.set(qn("before"), before)
    if after is not None:
        spacing.set(qn("after"), after)
    if before_lines is not None:
        spacing.set(qn("beforeLines"), before_lines)
    if after_lines is not None:
        spacing.set(qn("afterLines"), after_lines)
    if line is not None:
        spacing.set(qn("line"), line)
    if line_rule is not None:
        spacing.set(qn("lineRule"), line_rule)


def set_jc(paragraph: ET.Element, value: str) -> None:
    ppr = ensure_ppr(paragraph)
    jc = child(ppr, "jc")
    if jc is None:
        jc = ET.Element(qn("jc"))
        ppr.append(jc)
    jc.set(qn("val"), value)


def set_keep_next(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    if child(ppr, "keepNext") is None:
        ppr.append(ET.Element(qn("keepNext")))


def set_run_font(run: ET.Element, *, east_asia: str | None = None, latin: str | None = None, size: str | None = None) -> None:
    rpr = ensure_rpr(run)
    if east_asia or latin:
        fonts = child(rpr, "rFonts")
        if fonts is None:
            fonts = ET.Element(qn("rFonts"))
            rpr.insert(0, fonts)
        if east_asia:
            fonts.set(qn("eastAsia"), east_asia)
            fonts.set(qn("cs"), east_asia)
        if latin:
            fonts.set(qn("ascii"), latin)
            fonts.set(qn("hAnsi"), latin)
    if size:
        for tag in ("sz", "szCs"):
            node = child(rpr, tag)
            if node is None:
                node = ET.Element(qn(tag))
                rpr.append(node)
            node.set(qn("val"), size)


def format_all_runs(paragraph: ET.Element, *, east_asia: str | None = None, latin: str | None = None, size: str | None = None) -> None:
    for run in paragraph.findall(".//w:r", NS):
        set_run_font(run, east_asia=east_asia, latin=latin, size=size)


def set_paragraph_text(paragraph: ET.Element, value: str, *, east_asia: str | None = None, latin: str | None = None, size: str | None = None) -> None:
    ppr = paragraph.find("./w:pPr", NS)
    first_rpr = paragraph.find(".//w:rPr", NS)
    for item in list(paragraph):
        if item is not ppr:
            paragraph.remove(item)
    run = ET.SubElement(paragraph, qn("r"))
    if first_rpr is not None:
        run.append(deepcopy(first_rpr))
    set_run_font(run, east_asia=east_asia, latin=latin, size=size)
    text = ET.SubElement(run, qn("t"))
    if value.startswith(" ") or value.endswith(" "):
        text.set(f"{{{XML_NS}}}space", "preserve")
    text.text = value


def set_paragraph_runs(paragraph: ET.Element, values: list[str], *, east_asia: str | None = None, latin: str | None = None, size: str | None = None) -> None:
    ppr = paragraph.find("./w:pPr", NS)
    first_rpr = paragraph.find(".//w:rPr", NS)
    for item in list(paragraph):
        if item is not ppr:
            paragraph.remove(item)
    for value in values:
        if value == "\t":
            run = ET.SubElement(paragraph, qn("r"))
            if first_rpr is not None:
                run.append(deepcopy(first_rpr))
            set_run_font(run, east_asia=east_asia, latin=latin, size=size)
            ET.SubElement(run, qn("tab"))
            continue
        run = ET.SubElement(paragraph, qn("r"))
        if first_rpr is not None:
            run.append(deepcopy(first_rpr))
        set_run_font(run, east_asia=east_asia, latin=latin, size=size)
        text = ET.SubElement(run, qn("t"))
        if value.startswith(" ") or value.endswith(" "):
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = value


def visible_text_nodes(paragraph: ET.Element) -> list[ET.Element]:
    return paragraph.findall(".//w:t", NS)


def set_text_node_space(text_node: ET.Element) -> None:
    value = text_node.text or ""
    if value.startswith(" ") or value.endswith(" "):
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    else:
        text_node.attrib.pop(f"{{{XML_NS}}}space", None)


def replace_text_span_in_runs(paragraph: ET.Element, start: int, end: int, replacement: str) -> bool:
    """Replace visible text without removing bookmarks, fields, or run anchors."""
    nodes = visible_text_nodes(paragraph)
    offset = 0
    start_node: tuple[ET.Element, int] | None = None
    end_node: tuple[ET.Element, int] | None = None
    for node in nodes:
        value = node.text or ""
        next_offset = offset + len(value)
        if start_node is None and offset <= start <= next_offset:
            start_node = (node, start - offset)
        if end_node is None and offset <= end <= next_offset:
            end_node = (node, end - offset)
            break
        offset = next_offset
    if start_node is None or end_node is None:
        return False
    first, first_pos = start_node
    last, last_pos = end_node
    if first is last:
        value = first.text or ""
        first.text = value[:first_pos] + replacement + value[last_pos:]
        set_text_node_space(first)
        return True
    first.text = (first.text or "")[:first_pos] + replacement
    set_text_node_space(first)
    clearing = False
    for node in nodes:
        if node is first:
            clearing = True
            continue
        if not clearing:
            continue
        if node is last:
            node.text = (node.text or "")[last_pos:]
            set_text_node_space(node)
            return True
        node.text = ""
        set_text_node_space(node)
    return False


def replace_first_regex_preserving_runs(paragraph: ET.Element, pattern: str, repl: str) -> bool:
    value = text_of(paragraph)
    match = re.search(pattern, value)
    if not match:
        return False
    replacement = match.expand(repl)
    return replace_text_span_in_runs(paragraph, match.start(), match.end(), replacement)


def empty_paragraph() -> ET.Element:
    return ET.Element(qn("p"))


def is_page_break_paragraph(paragraph: ET.Element) -> bool:
    return (
        paragraph.tag == qn("p")
        and not compact(text_of(paragraph))
        and paragraph.find(".//w:br[@w:type='page']", NS) is not None
    )


def collapse_empty_section_break_before_body(root: ET.Element) -> dict[str, object]:
    body = find_body(root)
    children = list(body)
    first_body_index = find_direct_child_index(
        root,
        lambda n: n.tag == qn("p") and re.match(r"^第[一二三四五六七八九十0-9]+章\s+", text_of(n).strip()),
    )
    if first_body_index is None:
        return {"first_body_found": False, "moved_section_breaks": 0, "removed_empty_paragraphs": 0}

    moved = 0
    removed = 0
    for index in range(first_body_index - 1, -1, -1):
        node = children[index]
        if node.tag != qn("p"):
            continue
        sect_pr = node.find("./w:pPr/w:sectPr", NS)
        if sect_pr is None:
            continue
        if compact(text_of(node)) or is_page_break_paragraph(node):
            break
        target = None
        fallback = None
        for scan in range(index - 1, -1, -1):
            candidate = children[scan]
            if candidate.tag != qn("p"):
                continue
            if not compact(text_of(candidate)) and candidate.find("./w:pPr/w:sectPr", NS) is None:
                target = candidate
                break
            if compact(text_of(candidate)):
                fallback = candidate
                break
        if target is None:
            target = fallback
        if target is None:
            break
        ensure_ppr(node).remove(sect_pr)
        target_ppr = ensure_ppr(target)
        existing = target_ppr.find("./w:sectPr", NS)
        if existing is not None:
            target_ppr.remove(existing)
        target_ppr.append(sect_pr)
        moved += 1
        break
    return {"first_body_found": True, "moved_section_breaks": moved, "removed_empty_paragraphs": removed}


def make_page_break_paragraph() -> ET.Element:
    p = ET.Element(qn("p"))
    r = ET.SubElement(p, qn("r"))
    br = ET.SubElement(r, qn("br"))
    br.set(qn("type"), "page")
    return p


def direct_body_children(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("document.xml has no body")
    return list(body)


def direct_paragraphs(root: ET.Element) -> list[ET.Element]:
    return [node for node in direct_body_children(root) if node.tag == qn("p")]


def find_direct_child_index(root: ET.Element, predicate) -> int | None:
    for index, node in enumerate(direct_body_children(root)):
        if predicate(node):
            return index
    return None


def find_body(root: ET.Element) -> ET.Element:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("document.xml has no body")
    return body


def is_toc_sdt(node: ET.Element) -> bool:
    if node.tag != qn("sdt"):
        return False
    value = compact(text_of(node))
    if value.startswith(TOC_TITLE):
        return True
    instr = "".join(item.text or "" for item in node.findall(".//w:instrText", NS))
    return bool(re.search(r"(^|\s)TOC(\s|$)", instr, flags=re.IGNORECASE))


def toc_entry_level(text: str) -> int | None:
    value = text.strip()
    if not value or compact(value).startswith(TOC_TITLE):
        return None
    if re.match(r"^(摘要|Abstract|参考文献|致谢)(\s|\t|\d|[IVX])", value):
        return 1
    if re.match(r"^第[一二三四五六七八九十0-9]+章", value):
        return 1
    if re.match(r"^[1-9]\d*\.[1-9]\d*", value):
        return 2
    return None


def ensure_blank_before_toc(root: ET.Element) -> bool:
    for node in direct_body_children(root):
        if not is_toc_sdt(node):
            continue
        content = node.find(".//w:sdtContent", NS)
        if content is None:
            return False
        children = list(content)
        title_index = next(
            (
                index
                for index, child_node in enumerate(children)
                if child_node.tag == qn("p") and compact(text_of(child_node)).startswith(TOC_TITLE)
            ),
            None,
        )
        if title_index is None:
            return False
        previous = children[title_index - 1] if title_index else None
        if previous is not None and previous.tag == qn("p") and not compact(text_of(previous)) and not is_page_break_paragraph(previous):
            return False
        content.insert(title_index, empty_paragraph())
        return True
    return False


def move_toc_before_abstract(root: ET.Element) -> dict[str, object]:
    body = find_body(root)
    children = list(body)
    toc_index = find_direct_child_index(root, is_toc_sdt)
    if toc_index is None:
        return {"toc_found": False, "moved": False}
    zh_index = find_direct_child_index(root, lambda n: n.tag == qn("p") and compact(text_of(n)) == ZH_ABSTRACT)
    cover_end = None
    for index, node in enumerate(children):
        if zh_index is not None and index >= zh_index:
            break
        if node.tag == qn("p") and node.find("./w:pPr/w:sectPr", NS) is not None:
            cover_end = index
    if cover_end is None:
        return {"toc_found": True, "moved": False, "reason": "cover section boundary not found"}
    toc_node = children[toc_index]
    moved_nodes = [toc_node]
    following = children[toc_index + 1] if toc_index + 1 < len(children) else None
    if following is not None and is_page_break_paragraph(following):
        moved_nodes.append(following)
    else:
        moved_nodes.append(make_page_break_paragraph())
    for node in moved_nodes:
        if node in list(body):
            body.remove(node)
    insert_at = list(body).index(children[cover_end]) + 1 if children[cover_end] in list(body) else cover_end + 1
    for offset, node in enumerate(moved_nodes):
        body.insert(insert_at + offset, node)
    return {"toc_found": True, "moved": toc_index > cover_end, "old_index": toc_index, "new_index": insert_at}


def ensure_blank_after(root: ET.Element, match_text: str) -> bool:
    body = find_body(root)
    children = list(body)
    for index, node in enumerate(children):
        if node.tag == qn("p") and compact(text_of(node)) == compact(match_text):
            next_node = children[index + 1] if index + 1 < len(children) else None
            if next_node is not None and next_node.tag == qn("p") and not compact(text_of(next_node)) and not is_page_break_paragraph(next_node):
                return False
            body.insert(index + 1, empty_paragraph())
            return True
    return False


def ensure_blank_before(root: ET.Element, match_text: str) -> bool:
    body = find_body(root)
    children = list(body)
    for index, node in enumerate(children):
        if node.tag == qn("p") and compact(text_of(node)) == compact(match_text):
            previous = children[index - 1] if index else None
            if previous is not None and previous.tag == qn("p") and not compact(text_of(previous)) and not is_page_break_paragraph(previous):
                return False
            body.insert(index, empty_paragraph())
            return True
    return False


def format_front_matter(root: ET.Element) -> dict[str, int]:
    counts = {"titles": 0, "blank_insertions": 0}
    for title in (ZH_ABSTRACT, EN_ABSTRACT):
        if ensure_blank_before(root, title):
            counts["blank_insertions"] += 1
        if ensure_blank_after(root, title):
            counts["blank_insertions"] += 1
    for child_node in direct_body_children(root):
        if child_node.tag != qn("p"):
            continue
        value = compact(text_of(child_node))
        if value in {ZH_ABSTRACT, compact(EN_ABSTRACT)}:
            set_jc(child_node, "center")
            set_spacing(child_node, line="360", line_rule="auto")
            counts["titles"] += 1
        if text_of(child_node).strip().startswith("Key words"):
            set_jc(child_node, "left")
            set_spacing(child_node, before="0", after="0", line="360", line_rule="auto")
            counts["titles"] += 1
    for title in (ZH_ABSTRACT, EN_ABSTRACT):
        # Recompute after the title blank lines are inserted.
        children = direct_body_children(root)
        for index, node in enumerate(children):
            if node.tag == qn("p") and compact(text_of(node)) == compact(title):
                # Insert the content-after blank before the keyword line when it is absent.
                for scan in range(index + 1, min(index + 8, len(children))):
                    text = text_of(children[scan]).strip()
                    if text.startswith("\u5173\u952e\u8bcd") or text.startswith("Key words"):
                        prev = children[scan - 1]
                        if compact(text_of(prev)):
                            find_body(root).insert(scan, empty_paragraph())
                            counts["blank_insertions"] += 1
                        break
                break
    return counts


def format_toc(root: ET.Element) -> dict[str, int]:
    formatted = 0
    blank_inserted = 1 if ensure_blank_before_toc(root) else 0
    for node in direct_body_children(root):
        if not is_toc_sdt(node):
            continue
        for paragraph in node.findall(".//w:p", NS):
            text = text_of(paragraph)
            if not compact(text):
                continue
            set_spacing(paragraph, line="360", line_rule="auto")
            if compact(text).startswith(TOC_TITLE):
                set_jc(paragraph, "center")
            else:
                level = toc_entry_level(text)
                if level == 1:
                    format_all_runs(paragraph, east_asia=KAITI, latin="Times New Roman", size="28")
                elif level == 2:
                    format_all_runs(paragraph, size="28")
            formatted += 1
    return {"toc_paragraphs": formatted, "blank_before_toc_inserted": blank_inserted}


def is_body_heading(text: str) -> bool:
    stripped = text.strip()
    return bool(re.match(r"^第[一二三四五六七八九十0-9]+章\s+", stripped) or re.match(r"^[1-9]\d*\.[1-9]\d*\s+", stripped))


def format_body_headings(root: ET.Element) -> dict[str, int]:
    count = 0
    for paragraph in direct_paragraphs(root):
        value = text_of(paragraph).strip()
        if not is_body_heading(value):
            continue
        set_spacing(paragraph, before_lines="100", after_lines="100")
        if re.match(r"^第[一二三四五六七八九十0-9]+章\s+", value):
            set_jc(paragraph, "center")
        count += 1
    return {"body_headings": count}


def is_caption_text(text: str) -> bool:
    stripped = text.strip()
    if "\u7528\u4e8e\u8bf4\u660e" in stripped:
        return False
    return bool(
        re.match(r"^\u56fe\s*\d+[-.]\d+(?![\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u56fe\s*\d+(?![-.\d\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u8868\s*\d+[-.]\d+(?![\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u8868\s*\d+(?![-.\d\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u7eed\u8868\s*\d", stripped)
    )


def format_captions(root: ET.Element) -> dict[str, int]:
    count = 0
    for paragraph in direct_paragraphs(root):
        if is_caption_text(text_of(paragraph)):
            set_jc(paragraph, "center")
            set_keep_next(paragraph)
            set_spacing(paragraph, before_lines="50", after="0", line="360", line_rule="auto")
            format_all_runs(paragraph, east_asia=HEITI, latin="Times New Roman", size="21")
            count += 1
    return {"captions": count}


def format_tables(root: ET.Element) -> dict[str, int]:
    count = 0
    for paragraph in root.findall(".//w:tbl//w:tc//w:p", NS):
        set_jc(paragraph, "center")
        set_spacing(paragraph, before="0", after="0", line="240", line_rule="atLeast")
        count += 1
    return {"table_cell_paragraphs": count}


def format_tail_titles_and_references(root: ET.Element) -> dict[str, int]:
    counts = {"tail_titles": 0, "reference_label_spaces_removed": 0}
    for paragraph in direct_paragraphs(root):
        value = text_of(paragraph).strip()
        compact_value = compact(value)
        if compact_value in {REFERENCES_TITLE, ACK_TITLE, CONCLUSION_TITLE}:
            format_all_runs(paragraph, east_asia=HEITI, latin="Times New Roman", size="28")
            set_jc(paragraph, "center")
            set_spacing(paragraph, before="0", after="0", line="360", line_rule="auto")
            counts["tail_titles"] += 1
            continue
        if re.match(r"^\[\d+\]\s+", value):
            replace_first_regex_preserving_runs(paragraph, r"^(\[\d+\])\s+", r"\1")
            format_all_runs(paragraph, east_asia=SONGTI, latin="Times New Roman", size="21")
            set_spacing(paragraph, before="0", after="0", line="360", line_rule="auto")
            counts["reference_label_spaces_removed"] += 1
    return counts


def load_issue_ledger(path: Path | None) -> dict[str, object]:
    if path is None:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def comment_issues_from_ledger(ledger: dict[str, object]) -> list[dict[str, object]]:
    payload = ledger.get("comment_docx")
    if not isinstance(payload, dict):
        return []
    rows = payload.get("issues")
    return rows if isinstance(rows, list) else []


def find_paragraph_by_text(root: ET.Element, target: str) -> ET.Element | None:
    compact_target = compact(target)
    if not compact_target:
        return None
    for paragraph in direct_paragraphs(root):
        if compact_target in compact(text_of(paragraph)):
            return paragraph
    return None


def format_targeted_report_body_text(root: ET.Element, ledger: dict[str, object]) -> dict[str, int]:
    counts = {"font_size_repairs": 0, "punctuation_repairs": 0, "missing_targets": 0}
    for row in comment_issues_from_ledger(ledger):
        if not str(row.get("surface") or "").startswith("正文文本内容"):
            continue
        targets = row.get("target_texts") if isinstance(row.get("target_texts"), list) else []
        if not targets:
            continue
        paragraph = find_paragraph_by_text(root, str(targets[0]))
        if paragraph is None:
            counts["missing_targets"] += 1
            continue
        name = str(row.get("name") or "")
        if name == "标点符号问题":
            if replace_first_regex_preserving_runs(paragraph, r"(?<=[A-Za-z0-9]),(?=[A-Za-z])", "，"):
                format_all_runs(paragraph, east_asia=SONGTI, latin="Times New Roman", size="24")
                set_spacing(paragraph, before="0", after="0", line="360", line_rule="auto")
                counts["punctuation_repairs"] += 1
            continue
        if name in {"字体问题", "字号问题"}:
            set_spacing(paragraph, before="0", after="0", line="360", line_rule="auto")
            format_all_runs(paragraph, east_asia=SONGTI, latin="Times New Roman", size="24")
            counts["font_size_repairs"] += 1
        elif name == "标点符号问题":
            value = text_of(paragraph)
            new_value = re.sub(r"(?<=[A-Za-z0-9]),(?=[A-Za-z])", "，", value)
            if new_value != value:
                set_paragraph_text(paragraph, new_value, east_asia=SONGTI, latin="Times New Roman", size="24")
                set_spacing(paragraph, before="0", after="0", line="360", line_rule="auto")
                counts["punctuation_repairs"] += 1
    return counts


def body_heading_titles(root: ET.Element) -> list[str]:
    titles: list[str] = []
    for paragraph in direct_paragraphs(root):
        value = text_of(paragraph).strip()
        if is_header_right_title(value):
            titles.append(value)
    return titles


def is_header_right_title(value: str) -> bool:
    stripped = value.strip()
    return bool(re.match(r"^第[一二三四五六七八九十0-9]+章\s+\S+", stripped)) or compact(stripped) in {
        REFERENCES_TITLE,
        ACK_TITLE,
        CONCLUSION_TITLE,
    }


def normalize_word_target(target: str) -> str:
    target = (target or "").replace("\\", "/")
    if not target:
        return ""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join("word", target))


def document_header_targets(parts: dict[str, bytes]) -> dict[str, str]:
    rels_payload = parts.get("word/_rels/document.xml.rels")
    if not rels_payload:
        return {}
    rels_root = ET.fromstring(rels_payload)
    rows: dict[str, str] = {}
    for rel in rels_root:
        if not str(rel.get("Type") or "").endswith("/header"):
            continue
        rid = str(rel.get("Id") or "")
        target = normalize_word_target(str(rel.get("Target") or ""))
        if rid and target:
            rows[rid] = target
    return rows


def sect_header_parts(sect_pr: ET.Element, rel_targets: dict[str, str]) -> list[str]:
    rows: list[str] = []
    for ref in sect_pr.findall("./w:headerReference", NS):
        rid = ref.get(rqn("id")) or ""
        part = rel_targets.get(rid)
        if part and part not in rows:
            rows.append(part)
    return rows


def section_header_title_map(root: ET.Element, parts: dict[str, bytes]) -> dict[str, str]:
    rel_targets = document_header_targets(parts)
    body = find_body(root)
    current_section_titles: list[str] = []
    rows: dict[str, str] = {}

    def record_section(sect_pr: ET.Element) -> None:
        nonlocal current_section_titles
        if not current_section_titles:
            return
        for part in sect_header_parts(sect_pr, rel_targets):
            rows.setdefault(part, current_section_titles[0])
        current_section_titles = []

    for node in list(body):
        if node.tag == qn("p"):
            value = text_of(node).strip()
            if is_header_right_title(value):
                current_section_titles.append(value)
            sect_pr = node.find("./w:pPr/w:sectPr", NS)
            if sect_pr is not None:
                record_section(sect_pr)
        elif node.tag == qn("sectPr"):
            record_section(node)
    return rows


def set_frontmatter_page_numbering(root: ET.Element) -> dict[str, object]:
    body = find_body(root)
    children = list(body)
    first_body_index = find_direct_child_index(root, lambda n: n.tag == qn("p") and re.match(r"^第[一二三四五六七八九十0-9]+章\s+", text_of(n).strip()))
    if first_body_index is None:
        return {"frontmatter_section_found": False}
    frontmatter_sect = None
    for node in reversed(children[:first_body_index]):
        if node.tag == qn("p"):
            frontmatter_sect = node.find("./w:pPr/w:sectPr", NS)
            if frontmatter_sect is not None:
                break
    if frontmatter_sect is None:
        return {"frontmatter_section_found": False}
    pg = child(frontmatter_sect, "pgNumType")
    if pg is None:
        pg = ET.Element(qn("pgNumType"))
        frontmatter_sect.append(pg)
    pg.set(qn("fmt"), "upperRoman")
    pg.set(qn("start"), "1")
    return {"frontmatter_section_found": True, "fmt": "upperRoman", "start": "1"}


def header_right_title_from_text(header_text: str, fallback_title: str) -> str:
    header_text = header_text.strip("\t ")
    if header_text.startswith(HEADER_LEFT):
        right = header_text[len(HEADER_LEFT):].strip()
        if right:
            return right
    stripped = header_text.strip()
    if stripped and stripped != HEADER_LEFT:
        return stripped
    return fallback_title


def set_header_rule(header_root: ET.Element, *, fallback_title: str = "") -> dict[str, int]:
    changed = 0
    content_repaired = 0
    for paragraph in header_root.findall(".//w:p", NS):
        right_title = header_right_title_from_text(text_of(paragraph), fallback_title)
        ppr = ensure_ppr(paragraph)
        tabs = child(ppr, "tabs")
        if tabs is None:
            tabs = ET.Element(qn("tabs"))
            ppr.append(tabs)
        if not tabs.findall("./w:tab", NS):
            tab = ET.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "right")
            tab.set(qn("pos"), "9000")
        set_jc(paragraph, "left")
        pbdr = child(ppr, "pBdr")
        if pbdr is None:
            pbdr = ET.Element(qn("pBdr"))
            ppr.append(pbdr)
        bottom = child(pbdr, "bottom")
        if bottom is None:
            bottom = ET.Element(qn("bottom"))
            pbdr.append(bottom)
        bottom.set(qn("val"), "single")
        bottom.set(qn("sz"), "4")
        bottom.set(qn("space"), "1")
        bottom.set(qn("color"), "auto")
        if right_title:
            set_paragraph_runs(paragraph, [HEADER_LEFT, "\t", right_title], east_asia=SONGTI, latin="Times New Roman", size="18")
            content_repaired += 1
        else:
            for run in paragraph.findall(".//w:r", NS):
                set_run_font(run, east_asia=SONGTI, latin="Times New Roman", size="18")
        changed += 1
    return {"paragraphs": changed, "content_repaired": content_repaired}


def set_footer_page_number_font(footer_root: ET.Element) -> int:
    changed = 0
    for paragraph in footer_root.findall(".//w:p", NS):
        set_jc(paragraph, "center")
        for run in paragraph.findall(".//w:r", NS):
            set_run_font(run, east_asia=SONGTI, latin=SONGTI, size=FOOTER_SMALL_FIVE)
            changed += 1
    return changed


def set_update_fields(settings_root: ET.Element) -> bool:
    node = settings_root.find("./w:updateFields", NS)
    if node is None:
        node = ET.Element(qn("updateFields"))
        settings_root.append(node)
    node.set(qn("val"), "true")
    return True


PPR_ORDER = {
    "pStyle": 10,
    "keepNext": 20,
    "keepLines": 30,
    "pageBreakBefore": 40,
    "framePr": 50,
    "widowControl": 60,
    "numPr": 70,
    "suppressLineNumbers": 80,
    "pBdr": 90,
    "shd": 100,
    "tabs": 110,
    "suppressAutoHyphens": 120,
    "kinsoku": 130,
    "wordWrap": 140,
    "overflowPunct": 150,
    "topLinePunct": 160,
    "autoSpaceDE": 170,
    "autoSpaceDN": 180,
    "bidi": 190,
    "adjustRightInd": 200,
    "snapToGrid": 210,
    "spacing": 220,
    "ind": 230,
    "contextualSpacing": 240,
    "mirrorIndents": 250,
    "suppressOverlap": 260,
    "jc": 270,
    "textDirection": 280,
    "textAlignment": 290,
    "textboxTightWrap": 300,
    "outlineLvl": 310,
    "divId": 320,
    "cnfStyle": 330,
    "rPr": 800,
    "sectPr": 900,
    "pPrChange": 910,
}


STYLE_ORDER = {
    "name": 10,
    "aliases": 20,
    "basedOn": 30,
    "next": 40,
    "link": 50,
    "autoRedefine": 60,
    "hidden": 70,
    "uiPriority": 80,
    "semiHidden": 90,
    "unhideWhenUsed": 100,
    "qFormat": 110,
    "locked": 120,
    "personal": 130,
    "personalCompose": 140,
    "personalReply": 150,
    "rsid": 160,
    "pPr": 300,
    "rPr": 310,
    "tblPr": 320,
    "trPr": 330,
    "tcPr": 340,
    "tblStylePr": 500,
}


def normalize_ppr_order(root: ET.Element) -> int:
    changed = 0
    for ppr in root.findall(".//w:pPr", NS):
        children = list(ppr)
        ordered = sorted(enumerate(children), key=lambda item: (PPR_ORDER.get(local_name(item[1]), 500), item[0]))
        if [node for _, node in ordered] != children:
            for node in children:
                ppr.remove(node)
            for _, node in ordered:
                ppr.append(node)
            changed += 1
    return changed


def normalize_style_order(styles_root: ET.Element) -> int:
    changed = 0
    for style in styles_root.findall("./w:style", NS):
        children = list(style)
        ordered = sorted(enumerate(children), key=lambda item: (STYLE_ORDER.get(local_name(item[1]), 250), item[0]))
        if [node for _, node in ordered] != children:
            for node in children:
                style.remove(node)
            for _, node in ordered:
                style.append(node)
            changed += 1
    return changed


def remove_false_nowrap(root: ET.Element) -> int:
    removed = 0
    for parent in root.findall(".//w:tcPr", NS):
        for node in list(parent):
            if node.tag == qn("noWrap") and node.get(qn("val"), "") in {"0", "false", "off"}:
                parent.remove(node)
                removed += 1
    return removed


def remove_conflicting_math_run_style(root: ET.Element) -> int:
    removed = 0
    for rpr in root.findall(".//m:rPr", NS):
        if rpr.find("./m:nor", NS) is None or rpr.find("./m:sty", NS) is None:
            continue
        for node in list(rpr):
            if node.tag == mqn("sty"):
                rpr.remove(node)
                removed += 1
    return removed


def strip_comment_markup(root: ET.Element) -> int:
    removed = 0
    for parent in list(root.iter()):
        for node in list(parent):
            lname = local_name(node)
            if lname in {"commentRangeStart", "commentRangeEnd"}:
                parent.remove(node)
                removed += 1
            elif node.tag == qn("r") and node.find("./w:commentReference", NS) is not None and not text_of(node):
                parent.remove(node)
                removed += 1
            elif lname == "commentReference":
                parent.remove(node)
                removed += 1
    return removed


def serialize_xml(root: ET.Element) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    ignorable = root.get(f"{{{MC_NS}}}Ignorable", "")
    if not ignorable:
        return payload
    text = payload.decode("utf-8")
    namespace_uris = {
        "mc": MC_NS,
        "w14": W14_NS,
        "w15": W15_NS,
        "wp14": WP14_NS,
    }
    additions = []
    for prefix in ["mc", *ignorable.split()]:
        uri = namespace_uris.get(prefix)
        if uri and f"xmlns:{prefix}=" not in text:
            additions.append(f'xmlns:{prefix}="{uri}"')
    if additions:
        text = re.sub(r"(<[A-Za-z0-9_:.-]+)(\s|>)", lambda m: f"{m.group(1)} {' '.join(additions)}{m.group(2)}", text, count=1)
    return text.encode("utf-8")


def serialize_content_types(root: ET.Element) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    text = payload.decode("utf-8")
    text = re.sub(r"<(/?)ns\d+:", r"<\1", text)
    text = re.sub(r" xmlns:ns\d+=\"%s\"" % re.escape(CT_NS), "", text)
    text = text.replace(f"<Types>", f'<Types xmlns="{CT_NS}">', 1)
    return text.encode("utf-8")


def remove_report_comment_parts(parts: dict[str, bytes]) -> dict[str, object]:
    removed_parts = []
    for name in list(parts):
        if name in {"word/comments.xml", "word/commentsExtended.xml", "word/people.xml"}:
            removed_parts.append(name)
            del parts[name]
    if "[Content_Types].xml" in parts:
        ct_root = ET.fromstring(parts["[Content_Types].xml"])
        for override in list(ct_root):
            part_name = override.get("PartName", "")
            if part_name in {"/word/comments.xml", "/word/commentsExtended.xml", "/word/people.xml"}:
                ct_root.remove(override)
        parts["[Content_Types].xml"] = serialize_content_types(ct_root)
    if "word/_rels/document.xml.rels" in parts:
        rels = ET.fromstring(parts["word/_rels/document.xml.rels"])
        for rel in list(rels):
            target = rel.get("Target", "")
            rel_type = rel.get("Type", "")
            if target in {"comments.xml", "commentsExtended.xml", "people.xml"} or any(token in rel_type for token in ("comments", "commentsExtended", "people")):
                rels.remove(rel)
        parts["word/_rels/document.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    return {"removed_comment_parts": removed_parts}


def read_parts(docx: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    with zipfile.ZipFile(docx, "r") as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    return infos, parts


def write_parts(infos: list[zipfile.ZipInfo], parts: dict[str, bytes], output_docx: Path) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    written: set[str] = set()
    with zipfile.ZipFile(output_docx, "w") as archive:
        for info in infos:
            if info.filename not in parts:
                continue
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, parts[info.filename])
            written.add(info.filename)
        for name, payload in parts.items():
            if name not in written:
                archive.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)


def repair(input_docx: Path, output_docx: Path, *, strip_comments: bool, issue_ledger: Path | None = None) -> dict[str, object]:
    infos, parts = read_parts(input_docx)
    ledger = load_issue_ledger(issue_ledger)
    document_root = ET.fromstring(parts["word/document.xml"])
    header_titles = body_heading_titles(document_root) + [CONCLUSION_TITLE, REFERENCES_TITLE, ACK_TITLE]
    changes: dict[str, object] = {}
    changes["toc_move"] = move_toc_before_abstract(document_root)
    changes["front_matter"] = format_front_matter(document_root)
    changes["frontmatter_blank_page_collapse"] = collapse_empty_section_break_before_body(document_root)
    changes["toc_format"] = format_toc(document_root)
    changes["body_headings"] = format_body_headings(document_root)
    changes["captions"] = format_captions(document_root)
    changes["tables"] = format_tables(document_root)
    changes["tail_and_references"] = format_tail_titles_and_references(document_root)
    changes["targeted_body_text"] = format_targeted_report_body_text(document_root, ledger)
    changes["frontmatter_page_numbering"] = set_frontmatter_page_numbering(document_root)
    changes["false_nowrap_removed"] = remove_false_nowrap(document_root)
    changes["conflicting_math_run_style_removed"] = remove_conflicting_math_run_style(document_root)
    comment_markup_removed = 0
    if strip_comments:
        for part_name in list(parts):
            if not re.match(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$", part_name):
                continue
            root = document_root if part_name == "word/document.xml" else ET.fromstring(parts[part_name])
            comment_markup_removed += strip_comment_markup(root)
            parts[part_name] = serialize_xml(root)
        changes["comment_markup_removed"] = comment_markup_removed
        changes["comment_parts"] = remove_report_comment_parts(parts)
    header_paragraphs = 0
    header_content_repairs = 0
    footer_runs = 0
    ppr_order_repairs = normalize_ppr_order(document_root)
    parts["word/document.xml"] = serialize_xml(document_root)
    header_title_map = section_header_title_map(document_root, parts)
    header_index = 0
    for part_name in list(parts):
        if re.match(r"word/header\d+\.xml$", part_name):
            root = ET.fromstring(parts[part_name])
            fallback_title = header_title_map.get(
                part_name,
                header_titles[min(header_index, len(header_titles) - 1)] if header_titles else "",
            )
            header_changes = set_header_rule(root, fallback_title=fallback_title)
            header_paragraphs += header_changes["paragraphs"]
            header_content_repairs += header_changes["content_repaired"]
            ppr_order_repairs += normalize_ppr_order(root)
            parts[part_name] = serialize_xml(root)
            header_index += 1
        elif re.match(r"word/footer\d+\.xml$", part_name):
            root = ET.fromstring(parts[part_name])
            footer_runs += set_footer_page_number_font(root)
            ppr_order_repairs += normalize_ppr_order(root)
            parts[part_name] = serialize_xml(root)
    changes["header_rule_paragraphs"] = header_paragraphs
    changes["header_content_repairs"] = header_content_repairs
    changes["header_section_title_map"] = header_title_map
    changes["footer_runs_formatted"] = footer_runs
    changes["footer_page_number_size_half_points"] = FOOTER_SMALL_FIVE
    changes["ppr_order_repairs"] = ppr_order_repairs
    if "word/styles.xml" in parts:
        styles_root = ET.fromstring(parts["word/styles.xml"])
        changes["style_order_repairs"] = normalize_style_order(styles_root)
        parts["word/styles.xml"] = serialize_xml(styles_root)
    changes["update_fields"] = "not-mutated"
    write_parts(infos, parts, output_docx)
    return {
        "schema": "graduation-project-builder.docx-fanyu-format-report-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "generator": "scripts/repair_docx_fanyu_format_report.py",
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "issue_ledger": str(issue_ledger) if issue_ledger else "",
        "strip_report_comments": strip_comments,
        "changes": changes,
        "repair_verdict": "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument("--issue-ledger")
    parser.add_argument("--strip-report-comments", action="store_true")
    args = parser.parse_args()
    report = repair(
        Path(args.input_docx).resolve(),
        Path(args.output_docx).resolve(),
        strip_comments=args.strip_report_comments,
        issue_ledger=Path(args.issue_ledger).resolve() if args.issue_ledger else None,
    )
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"repair_verdict": report["repair_verdict"], "output_docx_sha256": report["output_docx_sha256"]}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
