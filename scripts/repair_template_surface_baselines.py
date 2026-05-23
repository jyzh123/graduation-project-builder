#!/usr/bin/env python3
"""Replay locked-template surface baselines onto a DOCX review copy.

This repair is intentionally bounded to template-owned format surfaces that
previous recovery passes frequently damage: styles, cover rows, static TOC
entry paragraph metrics, body heading levels 1-3, acknowledgement body,
reference title/entries, header/footer paragraph typography, image-holder
safety, and table-cell typography. It preserves visible project text and does
not rebuild the manuscript.
"""

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
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W = f"{{{W_NS}}}"
NS = {"w": W_NS, "r": R_NS}
EXTRA_NAMESPACES = {
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
for _prefix, _uri in EXTRA_NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)


def qn(local: str) -> str:
    return f"{W}{local}"


def rn(local: str) -> str:
    return f"{{{R_NS}}}{local}"


def prn(local: str) -> str:
    return f"{{{PR_NS}}}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def report_path(path: Path) -> str:
    return path.as_posix()


def ensure_root_namespace_declarations(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode("utf-8")
    start = text.find("<w:document")
    if start < 0:
        return xml_bytes
    end = text.find(">", start)
    if end < 0:
        return xml_bytes
    root_tag = text[start:end]
    additions = []
    for prefix, uri in EXTRA_NAMESPACES.items():
        if f"xmlns:{prefix}=" not in root_tag:
            additions.append(f' xmlns:{prefix}="{uri}"')
    if not additions:
        return xml_bytes
    return (text[:end] + "".join(additions) + text[end:]).encode("utf-8")


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "")


def body_paragraphs(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return [node for node in list(body) if node.tag == qn("p")]


def body_tables(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return [node for node in list(body) if node.tag == qn("tbl")]


def table_rows(table: ET.Element) -> list[ET.Element]:
    return table.findall("./w:tr", NS)


def row_cells(row: ET.Element) -> list[ET.Element]:
    return row.findall("./w:tc", NS)


def cell_text(cell: ET.Element) -> str:
    return "".join(node.text or "" for node in cell.findall(".//w:t", NS))


def paragraph_ppr(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find("./w:pPr", NS)


def first_text_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run).strip():
            rpr = run.find("./w:rPr", NS)
            return deepcopy(rpr) if rpr is not None else None
    return None


def replace_ppr(paragraph: ET.Element, donor_ppr: ET.Element | None, *, keep_page_break: bool = False) -> None:
    saved_page_break = None
    if keep_page_break:
        old = paragraph_ppr(paragraph)
        if old is not None:
            page_break = old.find("./w:pageBreakBefore", NS)
            if page_break is not None:
                saved_page_break = deepcopy(page_break)
    old = paragraph_ppr(paragraph)
    if old is not None:
        paragraph.remove(old)
    if donor_ppr is not None:
        paragraph.insert(0, deepcopy(donor_ppr))
    if saved_page_break is not None:
        ppr = paragraph_ppr(paragraph)
        if ppr is None:
            ppr = ET.Element(qn("pPr"))
            paragraph.insert(0, ppr)
        if ppr.find("./w:pageBreakBefore", NS) is None:
            ppr.append(saved_page_break)


def replace_text_run_rprs(paragraph: ET.Element, donor_rpr: ET.Element | None) -> int:
    changed = 0
    for run in paragraph.findall("./w:r", NS):
        if not paragraph_text(run):
            continue
        old = run.find("./w:rPr", NS)
        if old is not None:
            run.remove(old)
        if donor_rpr is not None:
            run.insert(0, deepcopy(donor_rpr))
        changed += 1
    return changed


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph_ppr(paragraph)
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


def ensure_rfonts(rpr: ET.Element) -> ET.Element:
    rfonts = rpr.find("./w:rFonts", NS)
    if rfonts is None:
        rfonts = ET.Element(qn("rFonts"))
        rpr.insert(0, rfonts)
    return rfonts


def ensure_child(parent: ET.Element, child_name: str) -> ET.Element:
    child = parent.find(f"./w:{child_name}", NS)
    if child is None:
        child = ET.Element(qn(child_name))
        parent.append(child)
    return child


def visible_run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def contains_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text or "")


def contains_latin_or_digit(text: str) -> bool:
    return any(("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch.isdigit() for ch in text or "")


def xml_paragraph_has_real_image(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:drawing", NS) is not None
        or paragraph.find(".//w:pict", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
    )


def remove_ppr_children(ppr: ET.Element, local_names: set[str]) -> int:
    removed = 0
    for child in list(ppr):
        if child.tag.rsplit("}", 1)[-1] in local_names:
            ppr.remove(child)
            removed += 1
    return removed


def set_or_clear_page_break(paragraph: ET.Element, saved_page_break: ET.Element | None) -> None:
    ppr = ensure_ppr(paragraph)
    for node in list(ppr.findall("./w:pageBreakBefore", NS)):
        ppr.remove(node)
    if saved_page_break is not None:
        ppr.append(deepcopy(saved_page_break))


def style_ids_and_names(styles_xml: bytes) -> set[str]:
    root = ET.fromstring(styles_xml)
    names: set[str] = set()
    for style in root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        name = style.find("./w:name", NS)
        if style_id:
            names.add(style_id.lower())
        if name is not None and name.get(qn("val")):
            names.add((name.get(qn("val")) or "").lower())
    return names


def set_text_node_value(text_node: ET.Element, text: str) -> None:
    text_node.text = text
    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(xml_space, "preserve")
    elif xml_space in text_node.attrib:
        del text_node.attrib[xml_space]


def set_cell_single_text(cell: ET.Element, text: str) -> None:
    text_nodes = cell.findall(".//w:t", NS)
    if text_nodes:
        set_text_node_value(text_nodes[0], text)
        for extra in text_nodes[1:]:
            set_text_node_value(extra, "")
        return

    paragraphs = cell.findall("./w:p", NS)
    if paragraphs:
        paragraph = paragraphs[0]
    else:
        paragraph = ET.Element(qn("p"))
        cell.append(paragraph)
    run = ET.Element(qn("r"))
    text_node = ET.Element(qn("t"))
    set_text_node_value(text_node, text)
    run.append(text_node)
    paragraph.append(run)


def set_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        style = ET.Element(qn("pStyle"))
        ppr.insert(0, style)
    style.set(qn("val"), style_id)


def heading_level(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    if "\u2026" in stripped:
        return None
    sep = r"[\s\u25a1]+"
    normalized_numbering = re.sub(r"(?<=\d)[\s\u25a1]*[\.．][\s\u25a1]*(?=\d)", ".", stripped)
    if re.match(r"^\u7b2c\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\s*\u7ae0(?:\s*\S.*)?$", stripped):
        return 1
    if re.match(r"^\d{1,2}\.\d{1,2}$", normalized_numbering):
        return 2
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}$", normalized_numbering):
        return 3
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}$", normalized_numbering):
        return 4
    if re.match(rf"^\d{{1,2}}{sep}\S", stripped):
        return 1
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}{sep}\S", normalized_numbering):
        return 2
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}\.\d{{1,2}}{sep}\S", normalized_numbering):
        return 3
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}\.\d{{1,2}}\.\d{{1,2}}{sep}\S", normalized_numbering):
        return 4
    return None


def is_toc_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    return normalized in {compact_text("目录").lower(), "contents", "tableofcontents"}


def strip_toc_page_number(text: str) -> str:
    value = str(text or "").strip()
    if "\t" in value:
        return value.split("\t", 1)[0].strip()
    if "\u2026" in value:
        return value.split("\u2026", 1)[0].strip()
    return re.sub(r"\s*\d+\s*$", "", value).strip()


def looks_like_toc_entry(text: str) -> bool:
    value = str(text or "").strip()
    if not value:
        return False
    if "\t" in value or "\u2026" in value:
        return heading_level(strip_toc_page_number(value)) is not None
    return bool(re.search(r"\d+\s*$", value)) and heading_level(strip_toc_page_number(value)) is not None


def is_reference_heading(text: str) -> bool:
    return compact_text(text).lower() in {compact_text("参考文献").lower(), "references", "bibliography"}


def is_ack_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    return normalized in {compact_text("致谢").lower(), compact_text("谢辞").lower(), "acknowledgements", "acknowledgments"}


def is_bibliography_entry(text: str) -> bool:
    return re.match(r"^\s*\[\d+\]", text or "") is not None


def is_toc_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    toc_label = compact_text("\u76ee\u5f55").lower()
    if normalized in {toc_label, "contents", "tableofcontents"}:
        return True
    if normalized.startswith(toc_label) and (
        len(normalized) == len(toc_label)
        or normalized[len(toc_label)] in {"(", "\uff08", "\uff1a", ":"}
    ):
        return True
    return False


def is_reference_heading(text: str) -> bool:
    return compact_text(text).lower() in {compact_text("\u53c2\u8003\u6587\u732e").lower(), "references", "bibliography"}


def is_ack_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    return normalized in {compact_text("\u81f4\u8c22").lower(), compact_text("\u8c22\u8f9e").lower(), "acknowledgements", "acknowledgments"}


def is_zh_abstract_title(text: str) -> bool:
    return compact_text(text).lower() in {compact_text("\u6458\u8981").lower(), compact_text("\u6458  \u8981").lower()}


def is_en_abstract_title(text: str) -> bool:
    return compact_text(text).lower() in {"abstract"}


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(qn("val")) if style is not None else ""


def has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None or "\t" in paragraph_text(paragraph)


def paragraph_style_id_to_name(styles_xml: bytes) -> dict[str, str]:
    root = ET.fromstring(styles_xml)
    mapping: dict[str, str] = {}
    for style in root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        name = style.find("./w:name", NS)
        if style_id and name is not None:
            mapping[style_id] = (name.get(qn("val")) or "").lower()
    return mapping


def paragraph_style_name_to_id(styles_xml: bytes) -> dict[str, str]:
    return {name: style_id for style_id, name in paragraph_style_id_to_name(styles_xml).items() if name}


def remap_style_id(style_id: str, source_style_names: dict[str, str], target_style_ids: dict[str, str]) -> str:
    style_name = source_style_names.get(style_id, "")
    return target_style_ids.get(style_name, style_id)


def remap_ppr_style_id(
    ppr: ET.Element | None,
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> ET.Element | None:
    if ppr is None:
        return None
    copied = deepcopy(ppr)
    style = copied.find("./w:pStyle", NS)
    if style is not None:
        style.set(qn("val"), remap_style_id(style.get(qn("val")) or "", source_style_names, target_style_ids))
    return copied


STYLE_REFERENCE_TAGS = {qn("basedOn"), qn("next"), qn("link")}


def merge_paragraph_style_definitions_by_name(
    template_styles_xml: bytes,
    final_styles_xml: bytes,
    style_names: set[str],
) -> tuple[bytes, list[dict[str, str]]]:
    template_root = ET.fromstring(template_styles_xml)
    final_root = ET.fromstring(final_styles_xml)
    source_style_names = paragraph_style_id_to_name(template_styles_xml)
    target_style_ids = paragraph_style_name_to_id(final_styles_xml)
    template_styles_by_name: dict[str, ET.Element] = {}
    final_styles_by_name: dict[str, ET.Element] = {}
    for style in template_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        name = source_style_names.get(style_id, "")
        if name:
            template_styles_by_name[name] = style
    for style in final_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        name_node = style.find("./w:name", NS)
        name = (name_node.get(qn("val")) or "").lower() if name_node is not None else ""
        if name:
            final_styles_by_name[name] = style

    changed: list[dict[str, str]] = []
    for name in sorted(style_names):
        template_style = template_styles_by_name.get(name)
        final_style = final_styles_by_name.get(name)
        target_style_id = target_style_ids.get(name, "")
        if template_style is None or final_style is None or not target_style_id:
            continue
        replacement = deepcopy(template_style)
        replacement.set(qn("styleId"), target_style_id)
        name_node = replacement.find("./w:name", NS)
        if name_node is not None:
            name_node.set(qn("val"), name_node.get(qn("val")) or name)
        for child in replacement:
            if child.tag in STYLE_REFERENCE_TAGS:
                child.set(qn("val"), remap_style_id(child.get(qn("val")) or "", source_style_names, target_style_ids))
        parent = final_root
        children = list(parent)
        index = children.index(final_style)
        parent.remove(final_style)
        parent.insert(index, replacement)
        changed.append({"style_name": name, "target_style_id": target_style_id})
    return ET.tostring(final_root, encoding="utf-8", xml_declaration=True), changed


def find_first(paragraphs: list[ET.Element], predicate, label: str) -> tuple[int, ET.Element]:
    for index, paragraph in enumerate(paragraphs):
        if predicate(index, paragraph):
            return index, paragraph
    raise RuntimeError(f"required donor/target paragraph not found: {label}")


def collect_template_heading_donors(paragraphs: list[ET.Element]) -> dict[int, tuple[str, ET.Element | None, ET.Element | None]]:
    donors: dict[int, tuple[str, ET.Element | None, ET.Element | None]] = {}
    in_toc = False
    body_started = False
    for paragraph in paragraphs:
        text = paragraph_text(paragraph).strip()
        current_style_id = paragraph_style_id(paragraph)
        if is_toc_heading(text):
            in_toc = True
            continue
        if in_toc:
            if looks_like_toc_entry(text):
                continue
            if heading_level(text) == 1 or current_style_id in {"1", "2", "3"}:
                in_toc = False
            else:
                continue
        if current_style_id.upper().startswith("TOC") or looks_like_toc_entry(text):
            continue
        level = heading_level(text)
        if level is None and current_style_id in {"1", "2", "3"}:
            level = int(current_style_id)
        if level == 1:
            body_started = True
        if not body_started or level not in {1, 2, 3}:
            continue
        donors.setdefault(level, (current_style_id, deepcopy(paragraph_ppr(paragraph)), first_text_rpr(paragraph)))
        if len(donors) == 3:
            break
    if not donors:
        raise RuntimeError("template heading donors missing all levels")
    return donors


def collect_template_toc_donors(paragraphs: list[ET.Element]) -> dict[int, tuple[str, ET.Element | None, ET.Element | None]]:
    donors: dict[int, tuple[str, ET.Element | None, ET.Element | None]] = {}
    in_toc = False
    for paragraph in paragraphs:
        text = paragraph_text(paragraph).strip()
        if is_toc_heading(text):
            in_toc = True
            continue
        if not in_toc:
            continue
        if not looks_like_toc_entry(text):
            if heading_level(text) == 1 and donors:
                break
            continue
        level = heading_level(strip_toc_page_number(text))
        if level in {1, 2, 3}:
            donors.setdefault(level, (paragraph_style_id(paragraph), deepcopy(paragraph_ppr(paragraph)), first_text_rpr(paragraph)))
            continue
    if 1 not in donors:
        raise RuntimeError("template TOC level-1 donor not found")
    return donors


def build_cover_donor_map(template_paragraphs: list[ET.Element]) -> dict[str, tuple[ET.Element | None, ET.Element | None]]:
    rows: dict[str, tuple[ET.Element | None, ET.Element | None]] = {}
    for index, paragraph in enumerate(template_paragraphs[:80]):
        text = paragraph_text(paragraph).strip()
        compact = compact_text(text)
        if not text:
            continue
        key = ""
        if index == 12 or "中文标题" in compact:
            key = "cover_zh_title"
        elif index == 14 or "外文标题" in compact:
            key = "cover_en_title"
        elif compact.startswith("承诺书") or "诚信承诺书" in compact:
            key = "commitment_title"
        elif compact.startswith("本人郑重承诺") or compact.startswith("我承诺"):
            key = "student_statement"
        elif compact.startswith("承诺人") or compact.startswith("学生（签名"):
            key = "student_signature"
        elif compact.startswith("在指导学生"):
            key = "teacher_statement"
        elif compact.startswith("指导教师（签名"):
            key = "teacher_signature"
        elif compact.startswith("日期") or re.fullmatch(r"年\s*月\s*日", text):
            key = "date_line"
        if key and key not in rows:
            rows[key] = (deepcopy(paragraph_ppr(paragraph)), first_text_rpr(paragraph))
    return rows


def cover_table_field_key(text: str) -> str:
    compact = compact_text(text).replace("\uff1a", "").replace(":", "")
    if "\u8bba\u6587\u9898\u76ee" in compact or "\u9898\u76ee" in compact:
        return "title"
    if "\u5b66\u9662" in compact or "\u9662\u7cfb" in compact:
        return "college"
    if "\u4e13\u4e1a" in compact:
        return "major"
    if "\u73ed\u7ea7" in compact:
        return "class"
    if "\u5b66\u53f7" in compact:
        return "student_id"
    if "\u59d3\u540d" in compact:
        return "student_name"
    if "\u6307\u5bfc\u6559\u5e08" in compact:
        return "advisor"
    return ""


def extract_cover_table_values(table: ET.Element) -> dict[str, str]:
    values: dict[str, str] = {}
    for row in table_rows(table):
        cells = row_cells(row)
        if len(cells) < 2:
            continue
        for index, label_cell in enumerate(cells[:-1]):
            key = cover_table_field_key(cell_text(label_cell))
            if not key:
                continue
            value = cell_text(cells[index + 1]).strip()
            if key not in values or value:
                values[key] = value
            break
    return values


def fill_cover_table_values(table: ET.Element, values: dict[str, str]) -> list[dict[str, str]]:
    changed: list[dict[str, str]] = []
    for row_index, row in enumerate(table_rows(table)):
        cells = row_cells(row)
        if len(cells) < 2:
            continue
        for cell_index, label_cell in enumerate(cells[:-1]):
            key = cover_table_field_key(cell_text(label_cell))
            if not key or key not in values:
                continue
            set_cell_single_text(cells[cell_index + 1], values.get(key, ""))
            changed.append({"row_index": str(row_index), "surface": f"cover_table_{key}"})
            break
    return changed


def replay_cover_table(
    template_document: ET.Element,
    final_document: ET.Element,
) -> list[dict[str, object]]:
    template_tables = body_tables(template_document)
    final_tables = body_tables(final_document)
    if not template_tables:
        raise RuntimeError("template cover table donor not found")
    if not final_tables:
        raise RuntimeError("target cover table not found")

    donor = deepcopy(template_tables[0])
    values = extract_cover_table_values(final_tables[0])
    filled = fill_cover_table_values(donor, values)

    body = final_document.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    children = list(body)
    target_index = children.index(final_tables[0])
    body.remove(final_tables[0])
    body.insert(target_index, donor)
    return [
        {
            "table_index": 0,
            "surface": "cover_identity_table",
            "source": "template table 0",
            "template_rows": len(table_rows(template_tables[0])),
            "target_rows_before": len(table_rows(final_tables[0])),
            "target_rows_after": len(table_rows(donor)),
            "filled_value_keys": sorted(key for key, value in values.items() if value),
            "field_updates": filled,
        }
    ]


def cover_key(text: str, index: int | None = None) -> str:
    compact = compact_text(text)
    if index == 12:
        return "cover_zh_title"
    if index == 14:
        return "cover_en_title"
    if compact.startswith("承诺书") or "诚信承诺书" in compact:
        return "commitment_title"
    if compact.startswith("本人郑重承诺") or compact.startswith("我承诺"):
        return "student_statement"
    if compact.startswith("承诺人") or compact.startswith("学生（签名"):
        return "student_signature"
    if compact.startswith("在指导学生"):
        return "teacher_statement"
    if compact.startswith("指导教师（签名"):
        return "teacher_signature"
    if compact.startswith("日期") or re.fullmatch(r"年\s*月\s*日", text.strip()):
        return "date_line"
    return ""


def replay_cover(
    template_paragraphs: list[ET.Element],
    final_paragraphs: list[ET.Element],
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> list[dict[str, object]]:
    donors = build_cover_donor_map(template_paragraphs)
    changed: list[dict[str, object]] = []
    for index, paragraph in enumerate(final_paragraphs[:90]):
        text = paragraph_text(paragraph).strip()
        key = cover_key(text, index)
        if not key or key not in donors:
            continue
        donor_ppr, donor_rpr = donors[key]
        replace_ppr(paragraph, remap_ppr_style_id(donor_ppr, source_style_names, target_style_ids))
        replace_text_run_rprs(paragraph, donor_rpr)
        changed.append({"paragraph_index": index, "surface": key, "text": text[:80]})
    return changed


def replay_headings(
    template_paragraphs: list[ET.Element],
    final_paragraphs: list[ET.Element],
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> list[dict[str, object]]:
    donors = collect_template_heading_donors(template_paragraphs)
    changed: list[dict[str, object]] = []
    in_toc = False
    body_started = False
    for index, paragraph in enumerate(final_paragraphs):
        text = paragraph_text(paragraph).strip()
        if is_toc_heading(text):
            in_toc = True
            continue
        if in_toc:
            if looks_like_toc_entry(text):
                continue
            if heading_level(text) == 1:
                in_toc = False
            else:
                continue
        if is_reference_heading(text) or is_ack_heading(text):
            continue
        level = heading_level(text)
        if level == 1:
            body_started = True
        if not body_started or level not in donors:
            continue
        style_id, donor_ppr, donor_rpr = donors[level]
        target_style_id = remap_style_id(style_id, source_style_names, target_style_ids)
        replace_ppr(paragraph, remap_ppr_style_id(donor_ppr, source_style_names, target_style_ids), keep_page_break=True)
        set_paragraph_style(paragraph, target_style_id)
        replace_text_run_rprs(paragraph, donor_rpr)
        changed.append({"paragraph_index": index, "level": level, "style_id": target_style_id, "text": text[:80]})
    return changed


def replay_toc(
    template_paragraphs: list[ET.Element],
    final_paragraphs: list[ET.Element],
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> list[dict[str, object]]:
    donors = collect_template_toc_donors(template_paragraphs)
    changed: list[dict[str, object]] = []
    in_toc = False
    for index, paragraph in enumerate(final_paragraphs):
        text = paragraph_text(paragraph).strip()
        if is_toc_heading(text):
            in_toc = True
            continue
        if not in_toc:
            continue
        if not looks_like_toc_entry(text):
            if heading_level(text) == 1:
                break
            continue
        level = heading_level(strip_toc_page_number(text))
        if level in donors:
            style_id, donor_ppr, donor_rpr = donors[level]
            target_style_id = remap_style_id(style_id, source_style_names, target_style_ids)
            replace_ppr(paragraph, remap_ppr_style_id(donor_ppr, source_style_names, target_style_ids))
            set_paragraph_style(paragraph, target_style_id)
            replace_text_run_rprs(paragraph, donor_rpr)
            changed.append({"paragraph_index": index, "level": level, "style_id": target_style_id, "text": text[:80]})
            continue
    return changed


def int_or_none(value: str | None) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(value)
    except ValueError:
        return None


def paragraph_indent_score(paragraph: ET.Element) -> int:
    ppr = paragraph_ppr(paragraph)
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    if ind is None:
        return 0
    score = 0
    first_line = int_or_none(ind.get(qn("firstLine")))
    left = int_or_none(ind.get(qn("left")))
    right = int_or_none(ind.get(qn("right")))
    if first_line is not None and first_line > 0:
        score += 4
    if left is None or abs(left) < 1000:
        score += 2
    if right is None or abs(right) < 1000:
        score += 2
    if paragraph_style_id(paragraph) != "1":
        score += 1
    return score


def find_abstract_range(
    paragraphs: list[ET.Element],
    title_predicate,
    keyword_predicate,
    label: str,
) -> tuple[int, int]:
    title_index, _title = find_first(
        paragraphs,
        lambda _i, p: title_predicate(paragraph_text(p).strip()) and not looks_like_toc_entry(paragraph_text(p).strip()),
        f"{label} title",
    )
    keyword_index, _keyword = find_first(
        paragraphs,
        lambda i, p: i > title_index
        and keyword_predicate(paragraph_text(p).strip())
        and not looks_like_toc_entry(paragraph_text(p).strip()),
        f"{label} keyword",
    )
    return title_index, keyword_index


def abstract_body_donor(
    paragraphs: list[ET.Element],
    title_index: int,
    keyword_index: int,
    title_predicate,
) -> ET.Element:
    candidates: list[ET.Element] = []
    for paragraph in paragraphs[title_index + 1 : keyword_index]:
        text = paragraph_text(paragraph).strip()
        if not text or title_predicate(text):
            continue
        candidates.append(paragraph)
    if not candidates:
        raise RuntimeError("abstract body donor not found")
    return sorted(candidates, key=paragraph_indent_score, reverse=True)[0]


def replay_abstracts(template_paragraphs: list[ET.Element], final_paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    surface_specs = [
        ("zh_abstract", is_zh_abstract_title, is_zh_keyword_text),
        ("en_abstract", is_en_abstract_title, is_en_keyword_text),
    ]
    for surface, title_predicate, keyword_predicate in surface_specs:
        template_title_index, template_keyword_index = find_abstract_range(
            template_paragraphs,
            title_predicate,
            keyword_predicate,
            f"template {surface}",
        )
        final_title_index, final_keyword_index = find_abstract_range(
            final_paragraphs,
            title_predicate,
            keyword_predicate,
            f"target {surface}",
        )
        donor_body = abstract_body_donor(
            template_paragraphs,
            template_title_index,
            template_keyword_index,
            title_predicate,
        )
        donor_keyword = template_paragraphs[template_keyword_index]
        for index in range(final_title_index + 1, final_keyword_index):
            paragraph = final_paragraphs[index]
            text = paragraph_text(paragraph).strip()
            if not text or title_predicate(text):
                continue
            replace_ppr(paragraph, paragraph_ppr(donor_body))
            set_paragraph_style(paragraph, paragraph_style_id(donor_body))
            replace_text_run_rprs(paragraph, first_text_rpr(donor_body))
            changes.append({"paragraph_index": index, "surface": f"{surface}_body", "text": text[:80]})
        final_keyword = final_paragraphs[final_keyword_index]
        replace_ppr(final_keyword, paragraph_ppr(donor_keyword))
        set_paragraph_style(final_keyword, paragraph_style_id(donor_keyword))
        replace_text_run_rprs(final_keyword, first_text_rpr(donor_keyword))
        changes.append({"paragraph_index": final_keyword_index, "surface": f"{surface}_keyword"})
    return changes


def replay_acknowledgement(template_paragraphs: list[ET.Element], final_paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    _donor_heading_index, donor_heading = find_first(template_paragraphs, lambda _i, p: is_ack_heading(paragraph_text(p).strip()), "template acknowledgement heading")
    donor_body_index, donor_body = find_first(
        template_paragraphs,
        lambda i, p: i > _donor_heading_index
        and bool(paragraph_text(p).strip())
        and not is_ack_heading(paragraph_text(p).strip())
        and paragraph_ppr(p) is not None
        and paragraph_ppr(p).find("./w:ind", NS) is not None,
        "template acknowledgement body",
    )
    del donor_body_index
    target_heading_index, target_heading = find_first(final_paragraphs, lambda _i, p: is_ack_heading(paragraph_text(p).strip()), "target acknowledgement heading")
    changed: list[dict[str, object]] = []
    replace_ppr(target_heading, paragraph_ppr(donor_heading), keep_page_break=True)
    set_paragraph_style(target_heading, paragraph_style_id(donor_heading))
    replace_text_run_rprs(target_heading, first_text_rpr(donor_heading))
    changed.append({"paragraph_index": target_heading_index, "surface": "acknowledgement_title"})
    for index in range(target_heading_index + 1, len(final_paragraphs)):
        paragraph = final_paragraphs[index]
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        if is_reference_heading(text) or is_bibliography_entry(text) or heading_level(text) == 1:
            break
        replace_ppr(paragraph, paragraph_ppr(donor_body))
        set_paragraph_style(paragraph, paragraph_style_id(donor_body))
        replace_text_run_rprs(paragraph, first_text_rpr(donor_body))
        changed.append({"paragraph_index": index, "surface": "acknowledgement_body", "text": text[:80]})
    return changed


def dominant_reference_entry_donor(template_paragraphs: list[ET.Element]) -> tuple[ET.Element | None, ET.Element | None]:
    candidates = [p for p in template_paragraphs if is_bibliography_entry(paragraph_text(p).strip())]
    if not candidates:
        raise RuntimeError("template reference entry donor not found")
    def score(paragraph: ET.Element) -> int:
        ppr = paragraph_ppr(paragraph)
        ind = ppr.find("./w:ind", NS) if ppr is not None else None
        return 2 if ind is not None and ind.get(qn("hangingChars")) == "177" else 1
    donor = sorted(candidates, key=score, reverse=True)[0]
    return deepcopy(paragraph_ppr(donor)), first_text_rpr(donor)


def replay_references(template_paragraphs: list[ET.Element], final_paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    _title_index, donor_title = find_first(template_paragraphs, lambda _i, p: is_reference_heading(paragraph_text(p).strip()), "template references title")
    entry_ppr, entry_rpr = dominant_reference_entry_donor(template_paragraphs)
    target_title_index, target_title = find_first(final_paragraphs, lambda _i, p: is_reference_heading(paragraph_text(p).strip()), "target references title")
    changed: list[dict[str, object]] = []
    replace_ppr(target_title, paragraph_ppr(donor_title), keep_page_break=True)
    set_paragraph_style(target_title, paragraph_style_id(donor_title))
    replace_text_run_rprs(target_title, first_text_rpr(donor_title))
    changed.append({"paragraph_index": target_title_index, "surface": "references_title"})
    for index in range(target_title_index + 1, len(final_paragraphs)):
        paragraph = final_paragraphs[index]
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        if is_ack_heading(text) or heading_level(text) == 1:
            break
        if not is_bibliography_entry(text):
            continue
        replace_ppr(paragraph, entry_ppr)
        set_paragraph_style(paragraph, "Normal")
        replace_text_run_rprs(paragraph, entry_rpr)
        changed.append({"paragraph_index": index, "surface": "references_entry", "text": text[:80]})
    return changed


def replay_footer_indent(
    template_zip: zipfile.ZipFile,
    final_parts: dict[str, bytes],
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> list[dict[str, object]]:
    changed: list[dict[str, object]] = []
    template_names = {name for name in template_zip.namelist() if name.startswith("word/footer") and name.endswith(".xml")}
    final_names = {name for name in final_parts if name.startswith("word/footer") and name.endswith(".xml")}
    donor_footer_name = sorted(template_names)[-1] if template_names else ""
    for name in sorted(final_names):
        template_name = name if name in template_names else donor_footer_name
        if not template_name:
            continue
        template_root = ET.fromstring(template_zip.read(template_name))
        final_root = ET.fromstring(final_parts[name])
        template_paragraphs = template_root.findall(".//w:p", NS)
        final_paragraphs = final_root.findall(".//w:p", NS)
        page_result_rpr = footer_page_result_rpr(template_paragraphs)
        touched = 0
        for idx, paragraph in enumerate(final_paragraphs):
            donor_ppr = paragraph_ppr(template_paragraphs[idx]) if idx < len(template_paragraphs) else None
            if donor_ppr is not None:
                replace_ppr(paragraph, remap_ppr_style_id(donor_ppr, source_style_names, target_style_ids))
            donor_rpr = page_result_rpr if page_result_rpr is not None else (first_text_rpr(template_paragraphs[idx]) if idx < len(template_paragraphs) else None)
            replace_text_run_rprs(paragraph, donor_rpr)
            for run in paragraph.findall(".//w:r", NS):
                if not visible_run_text(run).strip() and run.find(".//w:fldChar", NS) is None:
                    continue
                rpr = ensure_rpr(run)
                for tag in ("sz", "szCs"):
                    node = rpr.find(f"./w:{tag}", NS)
                    if node is None:
                        node = ET.Element(qn(tag))
                        rpr.append(node)
                    node.set(qn("val"), "18")
            touched += 1
        if touched:
            final_parts[name] = ET.tostring(final_root, encoding="utf-8", xml_declaration=True)
            changed.append({"part": name, "paragraphs": touched})
    return changed


def footer_page_result_rpr(template_paragraphs: list[ET.Element]) -> ET.Element | None:
    """Use the footer run whose visible PAGE result matches the template size."""
    for paragraph in template_paragraphs:
        for run in paragraph.findall("./w:r", NS):
            if not visible_run_text(run).strip():
                continue
            rpr = run.find("./w:rPr", NS)
            if rpr is None:
                continue
            size = rpr.find("./w:sz", NS)
            if size is not None and size.get(qn("val")) == "18":
                return deepcopy(rpr)
    for paragraph in template_paragraphs:
        for run in paragraph.findall("./w:r", NS):
            if not visible_run_text(run).strip():
                continue
            rpr = run.find("./w:rPr", NS)
            if rpr is None:
                continue
            if rpr.find("./w:sz", NS) is None and rpr.find("./w:szCs", NS) is None:
                return deepcopy(rpr)
    return None


def replay_header_indent(
    template_zip: zipfile.ZipFile,
    final_parts: dict[str, bytes],
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> list[dict[str, object]]:
    """Replay template header paragraph metrics on existing final header parts.

    Missing header part names are not synthesized here because section references
    in a real target manuscript may intentionally use fewer sections than the
    template. The final self-check accepts fewer parts when the active parts
    match a template header baseline.
    """
    changed: list[dict[str, object]] = []
    template_names = {name for name in template_zip.namelist() if name.startswith("word/header") and name.endswith(".xml")}
    final_names = {name for name in final_parts if name.startswith("word/header") and name.endswith(".xml")}
    for name in sorted(final_names & template_names):
        template_root = ET.fromstring(template_zip.read(name))
        final_root = ET.fromstring(final_parts[name])
        template_paragraphs = template_root.findall(".//w:p", NS)
        final_paragraphs = final_root.findall(".//w:p", NS)
        touched = 0
        for idx, paragraph in enumerate(final_paragraphs):
            donor_ppr = paragraph_ppr(template_paragraphs[idx]) if idx < len(template_paragraphs) else None
            if donor_ppr is not None:
                replace_ppr(paragraph, remap_ppr_style_id(donor_ppr, source_style_names, target_style_ids))
            donor_rpr = first_text_rpr(template_paragraphs[idx]) if idx < len(template_paragraphs) else None
            replace_text_run_rprs(paragraph, donor_rpr)
            touched += 1
        if touched:
            final_parts[name] = ET.tostring(final_root, encoding="utf-8", xml_declaration=True)
            changed.append({"part": name, "paragraphs": touched})
    return changed


def ensure_image_holder_style(final_parts: dict[str, bytes]) -> tuple[bytes, str, bool]:
    """Create or reuse a dedicated non-body image-holder paragraph style."""
    styles_root = ET.fromstring(final_parts["word/styles.xml"])
    desired_id = "ThesisImageHolder"
    existing = None
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        if (style.get(qn("styleId")) or "").lower() == desired_id.lower():
            existing = style
            break
        name = style.find("./w:name", NS)
        if name is not None and (name.get(qn("val")) or "").lower() == "thesis image holder":
            existing = style
            desired_id = style.get(qn("styleId")) or desired_id
            break
    if existing is not None:
        return ET.tostring(styles_root, encoding="utf-8", xml_declaration=True), desired_id, False

    style = ET.Element(qn("style"), {qn("type"): "paragraph", qn("styleId"): desired_id})
    ET.SubElement(style, qn("name"), {qn("val"): "Thesis Image Holder"})
    ET.SubElement(style, qn("uiPriority"), {qn("val"): "99"})
    ET.SubElement(style, qn("unhideWhenUsed"))
    ppr = ET.SubElement(style, qn("pPr"))
    ET.SubElement(ppr, qn("keepNext"))
    ET.SubElement(ppr, qn("jc"), {qn("val"): "center"})
    ET.SubElement(
        ppr,
        qn("spacing"),
        {qn("before"): "0", qn("after"): "0", qn("line"): "240", qn("lineRule"): "auto"},
    )
    ET.SubElement(
        ppr,
        qn("ind"),
        {qn("left"): "0", qn("right"): "0", qn("firstLine"): "0", qn("firstLineChars"): "0"},
    )
    styles_root.append(style)
    return ET.tostring(styles_root, encoding="utf-8", xml_declaration=True), desired_id, True


def repair_image_holder_safety(final_document: ET.Element, final_parts: dict[str, bytes]) -> list[dict[str, object]]:
    final_parts["word/styles.xml"], style_id, style_created = ensure_image_holder_style(final_parts)
    body = final_document.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    changes: list[dict[str, object]] = []
    if style_created:
        changes.append({"kind": "style_created", "style_id": style_id})
    toc_seen = False
    body_started = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if not xml_paragraph_has_real_image(child) and heading_level(text) is not None:
            body_started = True
            continue
        if not body_started or not xml_paragraph_has_real_image(child):
            continue
        ppr = ensure_ppr(child)
        style = ppr.find("./w:pStyle", NS)
        if style is None:
            style = ET.Element(qn("pStyle"))
            ppr.insert(0, style)
        old_style = style.get(qn("val")) or ""
        style.set(qn("val"), style_id)
        remove_ppr_children(ppr, {"outlineLvl", "numPr"})
        if ppr.find("./w:keepNext", NS) is None:
            ppr.append(ET.Element(qn("keepNext")))
        spacing = ppr.find("./w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(qn("spacing"))
            ppr.append(spacing)
        spacing.set(qn("line"), "240")
        spacing.set(qn("lineRule"), "auto")
        spacing.set(qn("before"), "120")
        spacing.set(qn("after"), "0")
        jc = ppr.find("./w:jc", NS)
        if jc is None:
            jc = ET.Element(qn("jc"))
            ppr.append(jc)
        jc.set(qn("val"), "center")
        ind = ppr.find("./w:ind", NS)
        if ind is None:
            ind = ET.Element(qn("ind"))
            ppr.append(ind)
        for key in ("left", "right", "firstLine", "firstLineChars", "leftChars", "rightChars", "hanging", "hangingChars"):
            if key in {"hanging", "hangingChars"}:
                ind.attrib.pop(qn(key), None)
            else:
                ind.set(qn(key), "0")
        changes.append({"kind": "image_holder", "paragraph_index": index, "old_style_id": old_style, "new_style_id": style_id})
    if len([c for c in changes if c.get("kind") == "image_holder"]) == 0:
        raise RuntimeError("image-holder repair did not locate body image-holder paragraphs")
    return changes


def body_table_cell_donors(template_document: ET.Element) -> dict[str, tuple[ET.Element | None, ET.Element | None]]:
    body = template_document.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    donors: dict[str, tuple[ET.Element | None, ET.Element | None]] = {}
    body_started = False
    for child in list(body):
        if child.tag == qn("p") and heading_level(paragraph_text(child).strip()) == 1:
            body_started = True
            continue
        if not body_started or child.tag != qn("tbl"):
            continue
        rows = child.findall("./w:tr", NS)
        for row_index, row in enumerate(rows[:2]):
            role = "header" if row_index == 0 else "body"
            for cell in row.findall("./w:tc", NS):
                para = next((p for p in cell.findall("./w:p", NS) if paragraph_text(p).strip()), None)
                if para is None:
                    continue
                cell_value = paragraph_text(para)
                if contains_cjk(cell_value):
                    donors.setdefault(f"{role}_cjk", (deepcopy(paragraph_ppr(para)), first_text_rpr(para)))
                if contains_latin_or_digit(cell_value):
                    donors.setdefault(f"{role}_latin", (deepcopy(paragraph_ppr(para)), first_text_rpr(para)))
                donors.setdefault(role, (deepcopy(paragraph_ppr(para)), first_text_rpr(para)))
        if donors:
            break
    if "header" not in donors or "body" not in donors:
        raise RuntimeError("template table-cell donor baseline not found")
    return donors


def split_runs_for_script(paragraph: ET.Element, donor_rprs: dict[str, ET.Element | None]) -> int:
    original_runs = [run for run in paragraph.findall("./w:r", NS)]
    if not original_runs:
        return 0
    changed = 0
    for run in original_runs:
        text_nodes = run.findall(".//w:t", NS)
        if len(text_nodes) != 1:
            continue
        value = text_nodes[0].text or ""
        if not value or (contains_cjk(value) and contains_latin_or_digit(value)):
            pieces: list[tuple[str, str]] = []
            current_kind = ""
            current = ""
            for ch in value:
                kind = "cjk" if contains_cjk(ch) else "latin" if contains_latin_or_digit(ch) else current_kind or "latin"
                if current and kind != current_kind:
                    pieces.append((current_kind, current))
                    current = ch
                else:
                    current += ch
                current_kind = kind
            if current:
                pieces.append((current_kind or "latin", current))
            if len(pieces) <= 1:
                continue
            insert_at = list(paragraph).index(run)
            paragraph.remove(run)
            for offset, (kind, part) in enumerate(pieces):
                new_run = ET.Element(qn("r"))
                donor_rpr = donor_rprs.get(kind)
                if donor_rpr is not None:
                    new_run.append(deepcopy(donor_rpr))
                text_node = ET.Element(qn("t"))
                set_text_node_value(text_node, part)
                new_run.append(text_node)
                paragraph.insert(insert_at + offset, new_run)
            changed += 1
    return changed


def repair_table_cell_baselines(template_document: ET.Element, final_document: ET.Element) -> list[dict[str, object]]:
    donors = body_table_cell_donors(template_document)
    body = final_document.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    changes: list[dict[str, object]] = []
    body_started = False
    table_index = 0
    for child in list(body):
        if child.tag == qn("p") and heading_level(paragraph_text(child).strip()) == 1:
            body_started = True
            continue
        if not body_started or child.tag != qn("tbl"):
            continue
        table_index += 1
        rows = child.findall("./w:tr", NS)
        for row_index, row in enumerate(rows):
            role = "header" if row_index == 0 else "body"
            for cell_index, cell in enumerate(row.findall("./w:tc", NS), start=1):
                for paragraph in cell.findall("./w:p", NS):
                    cell_value = paragraph_text(paragraph).strip()
                    if not cell_value:
                        continue
                    donor_key = f"{role}_cjk" if contains_cjk(cell_value) else f"{role}_latin" if contains_latin_or_digit(cell_value) else role
                    donor_ppr, donor_rpr = donors.get(donor_key) or donors[role]
                    replace_ppr(paragraph, donor_ppr)
                    replace_text_run_rprs(paragraph, donor_rpr)
                    donor_rprs = {
                        "cjk": (donors.get(f"{role}_cjk") or donors[role])[1],
                        "latin": (donors.get(f"{role}_latin") or donors[role])[1],
                    }
                    split_runs_for_script(paragraph, donor_rprs)
                    changes.append(
                        {
                            "table_index": table_index,
                            "row_index": row_index + 1,
                            "cell_index": cell_index,
                            "role": role,
                            "donor_key": donor_key,
                            "text": cell_value[:80],
                        }
                    )
    if not changes:
        raise RuntimeError("table-cell repair did not locate body table cells")
    return changes


def paragraph_section_properties(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find("./w:pPr/w:sectPr", NS)


def ensure_section_properties(paragraph: ET.Element) -> ET.Element:
    ppr = ensure_ppr(paragraph)
    sect = ppr.find("./w:sectPr", NS)
    if sect is None:
        sect = ET.Element(qn("sectPr"))
        ppr.append(sect)
    return sect


def replace_section_properties(paragraph: ET.Element, donor_section: ET.Element) -> ET.Element:
    ppr = ensure_ppr(paragraph)
    for node in list(ppr.findall("./w:sectPr", NS)):
        ppr.remove(node)
    cloned = deepcopy(donor_section)
    ppr.append(cloned)
    return cloned


def section_page_number_summary(section: ET.Element, *, required: bool) -> dict[str, str | bool | None]:
    pg_num = section.find("./w:pgNumType", NS)
    summary: dict[str, str | bool | None] = {"present": pg_num is not None, "fmt": None, "start": None}
    if pg_num is not None:
        summary["fmt"] = pg_num.get(qn("fmt"))
        summary["start"] = pg_num.get(qn("start"))
    if required and pg_num is None:
        raise RuntimeError("frontmatter-section-headers donor section missing required pgNumType")
    return summary


def relationship_targets(final_parts: dict[str, bytes]) -> tuple[ET.Element, dict[str, str]]:
    rels_root = ET.fromstring(final_parts["word/_rels/document.xml.rels"])
    target_to_rid: dict[str, str] = {}
    for rel in rels_root:
        rid = rel.get("Id") or rel.get(prn("Id")) or ""
        target = rel.get("Target") or rel.get(prn("Target")) or ""
        if rid and target:
            target_to_rid[target] = rid
    return rels_root, target_to_rid


def rel_id_for_target(final_parts: dict[str, bytes], target: str) -> str:
    _rels_root, target_to_rid = relationship_targets(final_parts)
    target = target.lstrip("/")
    if target.startswith("word/"):
        target = target.removeprefix("word/")
    rid = target_to_rid.get(target)
    if not rid:
        raise RuntimeError(f"missing document relationship for {target}")
    return rid


def set_section_references(
    section: ET.Element,
    *,
    header_targets: dict[str, str],
    footer_targets: dict[str, str],
    final_parts: dict[str, bytes],
) -> dict[str, object]:
    for tag in ("headerReference", "footerReference"):
        for node in list(section.findall(f"./w:{tag}", NS)):
            section.remove(node)
    changes: dict[str, object] = {"headers": {}, "footers": {}}
    insert_at = 0
    for ref_type, target in header_targets.items():
        node = ET.Element(qn("headerReference"), {qn("type"): ref_type, rn("id"): rel_id_for_target(final_parts, target)})
        section.insert(insert_at, node)
        insert_at += 1
        changes["headers"][ref_type] = target  # type: ignore[index]
    for ref_type, target in footer_targets.items():
        node = ET.Element(qn("footerReference"), {qn("type"): ref_type, rn("id"): rel_id_for_target(final_parts, target)})
        section.insert(insert_at, node)
        insert_at += 1
        changes["footers"][ref_type] = target  # type: ignore[index]
    return changes


def find_frontmatter_paragraph_index(final_paragraphs: list[ET.Element], predicate) -> int:
    for index, paragraph in enumerate(final_paragraphs):
        if predicate(paragraph_text(paragraph).strip()):
            return index
    raise RuntimeError("frontmatter-section-headers could not locate required paragraph")


def is_zh_keyword_text(text: str) -> bool:
    return compact_text(text).startswith(compact_text("\u5173\u952e\u8bcd"))


def is_en_keyword_text(text: str) -> bool:
    normalized = compact_text(text).lower()
    return normalized.startswith("keywords") or normalized.startswith("keywords:") or normalized.startswith("keywords\uff1a") or normalized.startswith("keywords")


def is_toc_title_text(text: str) -> bool:
    return is_toc_heading(text)


def first_body_heading_index(final_paragraphs: list[ET.Element]) -> int:
    toc_seen = False
    for index, paragraph in enumerate(final_paragraphs):
        text = paragraph_text(paragraph).strip()
        if is_toc_title_text(text):
            toc_seen = True
            continue
        if toc_seen and looks_like_toc_entry(text):
            continue
        if heading_level(text) == 1:
            return index
        if toc_seen and text and not has_tab(paragraph):
            style_id = paragraph_style_id(paragraph)
            if style_id and not style_id.upper().startswith("TOC") and style_id not in {"16", "17", "13"}:
                return index
    raise RuntimeError("frontmatter-section-headers could not locate first body heading")


def template_section_donors(template_document: ET.Element) -> dict[str, ET.Element]:
    paragraphs = body_paragraphs(template_document)
    zh_keyword = find_frontmatter_paragraph_index(paragraphs, is_zh_keyword_text)
    en_keyword = find_frontmatter_paragraph_index(paragraphs, is_en_keyword_text)
    toc_title = find_frontmatter_paragraph_index(paragraphs, is_toc_title_text)
    body_start = first_body_heading_index(paragraphs)

    def first_section_between(start: int, end: int, label: str) -> ET.Element:
        for paragraph in paragraphs[start : end + 1]:
            section = paragraph_section_properties(paragraph)
            if section is not None:
                return section
        raise RuntimeError(f"template section donor missing for {label}")

    toc_tail = body_start - 1
    donors: dict[str, ET.Element] = {}
    donors["zh_abstract"] = first_section_between(zh_keyword, en_keyword - 1, "zh_abstract")
    donors["en_abstract"] = first_section_between(en_keyword, toc_title - 1, "en_abstract")
    donors["toc"] = first_section_between(toc_title, toc_tail, "toc")
    for label, section in donors.items():
        section_page_number_summary(section, required=True)
    return donors


def rewrite_header_part_text(final_parts: dict[str, bytes], part: str, text: str, donor_part: str | None = None) -> dict[str, object]:
    if part not in final_parts:
        raise RuntimeError(f"missing header part to rewrite: {part}")
    root = ET.fromstring(final_parts[part])
    first_text = root.find(".//w:t", NS)
    before = "".join(node.text or "" for node in root.findall(".//w:t", NS))
    if donor_part and donor_part in final_parts:
        donor_root = ET.fromstring(final_parts[donor_part])
        donor_para = donor_root.find(".//w:p", NS)
        target_para = root.find(".//w:p", NS)
        if donor_para is not None and target_para is not None:
            replace_ppr(target_para, paragraph_ppr(donor_para))
            replace_text_run_rprs(target_para, first_text_rpr(donor_para))
    first_text = root.find(".//w:t", NS)
    if first_text is None:
        paragraph = root.find(".//w:p", NS)
        if paragraph is None:
            paragraph = ET.Element(qn("p"))
            root.append(paragraph)
        run = ET.Element(qn("r"))
        first_text = ET.Element(qn("t"))
        run.append(first_text)
        paragraph.append(run)
    set_text_node_value(first_text, text)
    for extra in root.findall(".//w:t", NS)[1:]:
        set_text_node_value(extra, "")
    final_parts[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return {"part": part, "text_before": before[:120], "text_after": text}


def repair_frontmatter_section_headers(
    template_document: ET.Element,
    final_document: ET.Element,
    final_parts: dict[str, bytes],
) -> list[dict[str, object]]:
    """Attach template-like front-matter header/footer references to target sections.

    The target manuscript intentionally has fewer sections than the school
    template, but Chinese abstract, English abstract, and TOC still require
    independent header identities in rendered output.
    """
    final_paragraphs = body_paragraphs(final_document)
    donors = template_section_donors(template_document)
    zh_keyword = find_frontmatter_paragraph_index(final_paragraphs, is_zh_keyword_text)
    en_keyword = find_frontmatter_paragraph_index(final_paragraphs, is_en_keyword_text)
    toc_title = find_frontmatter_paragraph_index(final_paragraphs, is_toc_title_text)
    body_start = first_body_heading_index(final_paragraphs)
    if not (zh_keyword < en_keyword < toc_title < body_start):
        raise RuntimeError("frontmatter-section-headers located inconsistent front-matter order")

    changes: list[dict[str, object]] = []
    zh_section = replace_section_properties(final_paragraphs[zh_keyword], donors["zh_abstract"])
    changes.append(
        {
            "surface": "zh_abstract",
            "paragraph_index": zh_keyword,
            "pg_num_type": section_page_number_summary(zh_section, required=True),
            **set_section_references(
                zh_section,
                header_targets={"default": "header3.xml"},
                footer_targets={"default": "footer3.xml", "even": "footer4.xml"},
                final_parts=final_parts,
            ),
        }
    )

    en_section = replace_section_properties(final_paragraphs[en_keyword], donors["en_abstract"])
    changes.append(
        {
            "surface": "en_abstract",
            "paragraph_index": en_keyword,
            "pg_num_type": section_page_number_summary(en_section, required=True),
            **set_section_references(
                en_section,
                header_targets={"default": "header4.xml", "even": "header5.xml"},
                footer_targets={"default": "footer5.xml", "even": "footer6.xml"},
                final_parts=final_parts,
            ),
        }
    )

    toc_tail_index = body_start - 1
    toc_section = replace_section_properties(final_paragraphs[toc_tail_index], donors["toc"])
    changes.append(
        {
            "surface": "toc",
            "paragraph_index": toc_tail_index,
            "pg_num_type": section_page_number_summary(toc_section, required=True),
            **set_section_references(
                toc_section,
                header_targets={"default": "header6.xml", "even": "header6.xml"},
                footer_targets={"default": "footer7.xml", "even": "footer8.xml"},
                final_parts=final_parts,
            ),
        }
    )
    changes.append({"surface": "header_text", **rewrite_header_part_text(final_parts, "word/header3.xml", "\u6458\u8981", "word/header2.xml")})
    changes.append({"surface": "header_text", **rewrite_header_part_text(final_parts, "word/header4.xml", "Abstract", "word/header4.xml")})
    changes.append({"surface": "header_text", **rewrite_header_part_text(final_parts, "word/header5.xml", "Abstract", "word/header4.xml")})
    changes.append({"surface": "header_text", **rewrite_header_part_text(final_parts, "word/header6.xml", "\u76ee\u5f55", "word/header6.xml")})
    return changes


def changed_zip_parts(before: Path, after: Path) -> list[str]:
    with zipfile.ZipFile(before) as zf:
        before_map = {name: zf.read(name) for name in zf.namelist()}
    with zipfile.ZipFile(after) as zf:
        after_map = {name: zf.read(name) for name in zf.namelist()}
    return [name for name in sorted(set(before_map) | set(after_map)) if before_map.get(name) != after_map.get(name)]


SURFACE_CHOICES = {
    "abstracts",
    "cover",
    "toc",
    "headings",
    "acknowledgement",
    "references",
    "header",
    "footer",
    "frontmatter-section-headers",
    "image-holder",
    "table-cells",
    "styles",
}


def parse_surface_scope(raw: str) -> set[str]:
    surfaces = {item.strip().lower() for item in raw.split(",") if item.strip()}
    unknown = sorted(surfaces - SURFACE_CHOICES)
    if unknown:
        raise RuntimeError(f"unknown surface scope(s): {', '.join(unknown)}")
    if not surfaces:
        raise RuntimeError("explicit --surfaces is required; do not run a broad template replay by default")
    return surfaces


def repair(input_docx: Path, template_docx: Path, output_docx: Path, *, surfaces: set[str]) -> dict[str, object]:
    if input_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy")
    with zipfile.ZipFile(template_docx) as template_zip, zipfile.ZipFile(input_docx) as zin:
        template_document = ET.fromstring(template_zip.read("word/document.xml"))
        final_document = ET.fromstring(zin.read("word/document.xml"))
        template_styles_xml = template_zip.read("word/styles.xml")
        final_styles_xml = zin.read("word/styles.xml")
        template_style_names = paragraph_style_id_to_name(template_styles_xml)
        final_style_ids = paragraph_style_name_to_id(final_styles_xml)
        template_paragraphs = body_paragraphs(template_document)
        final_paragraphs = body_paragraphs(final_document)
        final_parts = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        style_definition_changes: list[dict[str, str]] = []
        if "styles" in surfaces:
            final_parts["word/styles.xml"], style_definition_changes = merge_paragraph_style_definitions_by_name(
                template_styles_xml,
                final_styles_xml,
                {"normal", "heading 1", "heading 2", "heading 3", "toc 1", "toc 2", "toc 3"},
            )

        changes: dict[str, object] = {
            "requested_surfaces": sorted(surfaces),
            "styles_xml_merged_from_template_by_name": "styles" in surfaces,
            "style_definition_changes": style_definition_changes,
            "abstracts": replay_abstracts(template_paragraphs, final_paragraphs) if "abstracts" in surfaces else [],
            "cover": replay_cover(template_paragraphs, final_paragraphs, template_style_names, final_style_ids) if "cover" in surfaces else [],
            "cover_table": replay_cover_table(template_document, final_document) if "cover" in surfaces else [],
            "toc_entries": replay_toc(template_paragraphs, final_paragraphs, template_style_names, final_style_ids) if "toc" in surfaces else [],
            "body_headings": replay_headings(template_paragraphs, final_paragraphs, template_style_names, final_style_ids) if "headings" in surfaces else [],
            "acknowledgement": replay_acknowledgement(template_paragraphs, final_paragraphs) if "acknowledgement" in surfaces else [],
            "references": replay_references(template_paragraphs, final_paragraphs) if "references" in surfaces else [],
            "header_indent": replay_header_indent(template_zip, final_parts, template_style_names, final_style_ids) if "header" in surfaces else [],
            "footer_indent": replay_footer_indent(template_zip, final_parts, template_style_names, final_style_ids) if "footer" in surfaces else [],
            "frontmatter_section_headers": repair_frontmatter_section_headers(template_document, final_document, final_parts) if "frontmatter-section-headers" in surfaces else [],
            "image_holder_safety": repair_image_holder_safety(final_document, final_parts) if "image-holder" in surfaces else [],
            "table_cells": repair_table_cell_baselines(template_document, final_document) if "table-cells" in surfaces else [],
        }

        final_parts["word/document.xml"] = ensure_root_namespace_declarations(
            ET.tostring(final_document, encoding="utf-8", xml_declaration=True)
        )

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td) / "out.docx"
            with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    zout.writestr(item, final_parts[item.filename])
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(tmp, output_docx)

    return {
        "schema": "graduation-project-builder.template-surface-baseline-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "generator": "scripts/repair_template_surface_baselines.py",
        "input_docx": report_path(input_docx),
        "input_sha256": sha256_file(input_docx),
        "template_docx": report_path(template_docx),
        "template_sha256": sha256_file(template_docx),
        "output_docx": report_path(output_docx),
        "output_sha256": sha256_file(output_docx),
        "changed_zip_parts": changed_zip_parts(input_docx, output_docx),
        "changes": changes,
        "verdict": "needs-protected-surface-review" if "styles" in surfaces and len(surfaces) > 1 else "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Replay locked template surface baselines onto a DOCX review copy.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument(
        "--surfaces",
        required=True,
        help=(
            "Comma-separated explicit scope: abstracts,cover,toc,headings,acknowledgement,references,header,footer,"
            "frontmatter-section-headers,image-holder,table-cells,styles. "
            "Including styles replaces word/styles.xml and must be justified by the caller."
        ),
    )
    args = parser.parse_args()

    surfaces = parse_surface_scope(args.surfaces)
    report = repair(
        Path(args.input_docx).resolve(),
        Path(args.template_docx).resolve(),
        Path(args.output_docx).resolve(),
        surfaces=surfaces,
    )
    report_path = Path(args.report).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    verdict = "needs-protected-surface-review" if "styles" in surfaces and len(surfaces) > 1 else "pass"
    print(json.dumps({"output_docx": report["output_docx"], "changed_zip_parts": report["changed_zip_parts"], "verdict": verdict}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
