#!/usr/bin/env python3
"""Repair generated thesis DOCX structures that fail the whole-format gate.

This helper is intentionally bounded to structural release blockers:
section topology, live TOC field/cache, footer PAGE fields, page-number chain,
school-spec page geometry, and builder-owned paragraph-style contamination.
It preserves media, tables, comments, citations, and body text.
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
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
V_NS = "urn:schemas-microsoft-com:vml"

W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
REL = f"{{{REL_NS}}}"
CT = f"{{{CT_NS}}}"
NS = {"w": W_NS, "r": R_NS, "rel": REL_NS, "ct": CT_NS, "v": V_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", REL_NS)
ET.register_namespace("v", V_NS)


FALLBACK_HEADER_TEXT = "学士学位论文"
HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def serialize_content_types_root(root: ET.Element) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    text = payload.decode("utf-8")
    match = re.search(
        rf"<(?P<prefix>ns\d+):Types xmlns:(?P=prefix)=\"{re.escape(CT_NS)}\">",
        text,
    )
    if not match:
        return payload
    prefix = match.group("prefix")
    text = text.replace(
        f'<{prefix}:Types xmlns:{prefix}="{CT_NS}">',
        f'<Types xmlns="{CT_NS}">',
        1,
    )
    text = re.sub(fr"<(/?){prefix}:", r"<\1", text)
    return text.encode("utf-8")


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def child_local_name(node: ET.Element) -> str:
    return node.tag.rsplit("}", 1)[-1] if "}" in node.tag else node.tag


def body_element(root: ET.Element) -> ET.Element:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return body


def direct_paragraphs(body: ET.Element) -> list[ET.Element]:
    return [child for child in list(body) if child.tag == qn("p")]


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(qn("val"), "") if style is not None else ""


def style_metadata_map(styles_root: ET.Element | None) -> dict[str, dict[str, str]]:
    if styles_root is None:
        return {}
    result: dict[str, dict[str, str]] = {}
    for style in styles_root.findall("./w:style", NS):
        style_id = style.get(qn("styleId"), "")
        if not style_id:
            continue
        name_node = style.find("./w:name", NS)
        outline_node = style.find("./w:pPr/w:outlineLvl", NS)
        result[style_id] = {
            "name": name_node.get(qn("val"), "") if name_node is not None else "",
            "outline_level": outline_node.get(qn("val"), "") if outline_node is not None else "",
        }
    return result


def is_heading1_style(style_id: str, styles: dict[str, dict[str, str]]) -> bool:
    if style_id in {"Heading1", "1"}:
        return True
    meta = styles.get(style_id, {})
    name = compact(meta.get("name", "")).lower()
    return (
        meta.get("outline_level") == "0"
        or name in {"heading1", "标题1", "一级标题"}
        or "heading1" in name
        or "一级标题" in name
    )


def set_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = ensure_ppr(paragraph)
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        style = ET.Element(qn("pStyle"))
        ppr.insert(0, style)
    style.set(qn("val"), style_id)


def remove_children(parent: ET.Element, local_names: set[str]) -> None:
    for child in list(parent):
        if child_local_name(child) in local_names:
            parent.remove(child)


def clear_paragraph_runs(paragraph: ET.Element) -> None:
    for child in list(paragraph):
        if child.tag == qn("r") or child.tag == qn("hyperlink"):
            paragraph.remove(child)


def ensure_run_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        run.insert(0, rpr)
    return rpr


def set_run_size(run: ET.Element, size: str | None) -> None:
    if not size:
        return
    rpr = ensure_run_rpr(run)
    for local in ("sz", "szCs"):
        node = rpr.find(f"./w:{local}", NS)
        if node is None:
            node = ET.SubElement(rpr, qn(local))
        node.set(qn("val"), size)


def ensure_run_text(parent: ET.Element, text: str, *, size: str | None = None) -> ET.Element:
    run = ET.SubElement(parent, qn("r"))
    set_run_size(run, size)
    text_node = ET.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    text_node.text = text
    return run


def add_tab_run(parent: ET.Element) -> ET.Element:
    run = ET.SubElement(parent, qn("r"))
    ET.SubElement(run, qn("tab"))
    return run


def add_field_run(parent: ET.Element, field_char_type: str, *, locked: bool = False, size: str | None = None) -> ET.Element:
    run = ET.SubElement(parent, qn("r"))
    set_run_size(run, size)
    field = ET.SubElement(run, qn("fldChar"))
    field.set(qn("fldCharType"), field_char_type)
    if locked and field_char_type == "begin":
        field.set(qn("fldLock"), "true")
    return run


def add_instr_run(parent: ET.Element, instr: str, *, size: str | None = None) -> ET.Element:
    run = ET.SubElement(parent, qn("r"))
    set_run_size(run, size)
    instr_text = ET.SubElement(run, qn("instrText"))
    instr_text.set(f"{{{XML_NS}}}space", "preserve")
    instr_text.text = instr
    return run


def run_has_toc_field_surface(run: ET.Element) -> bool:
    return (
        run.find(".//w:fldChar", NS) is not None
        or run.find(".//w:instrText", NS) is not None
        or run.find(".//w:fldSimple", NS) is not None
    )


def strip_toc_field_surfaces(paragraph: ET.Element) -> int:
    removed = 0
    for child in list(paragraph):
        if child.tag == qn("hyperlink") and child.find(".//w:fldSimple", NS) is not None:
            paragraph.remove(child)
            removed += 1
            continue
        if child.tag != qn("r"):
            continue
        if not run_has_toc_field_surface(child):
            continue
        for node in list(child):
            if child_local_name(node) in {"fldChar", "instrText", "fldSimple"}:
                child.remove(node)
                removed += 1
        if not text_of(child).strip() and child.find(".//w:tab", NS) is None:
            paragraph.remove(child)
    return removed


def normalize_toc_field_boundary(paragraphs: list[ET.Element], toc_title: int, body_start: int) -> dict[str, object]:
    rows = [index for index in range(toc_title + 1, body_start) if text_of(paragraphs[index]).strip()]
    if not rows:
        return {"toc_field_rows": 0, "removed_toc_field_surfaces": 0, "toc_field_boundary_repaired": False}
    removed = 0
    for index in range(toc_title + 1, body_start):
        removed += strip_toc_field_surfaces(paragraphs[index])
    first_entry = paragraphs[rows[0]]
    last_entry = paragraphs[rows[-1]]
    insert_at = 1 if first_entry.find("./w:pPr", NS) is not None else 0
    first_entry.insert(insert_at, add_field_run(ET.Element("tmp"), "begin", locked=True))
    first_entry.insert(insert_at + 1, add_instr_run(ET.Element("tmp"), ' TOC \\o "1-3" \\h \\z \\u '))
    first_entry.insert(insert_at + 2, add_field_run(ET.Element("tmp"), "separate"))
    last_entry.append(add_field_run(ET.Element("tmp"), "end"))
    return {
        "toc_field_rows": len(rows),
        "toc_field_first_row_index": rows[0],
        "toc_field_last_row_index": rows[-1],
        "removed_toc_field_surfaces": removed,
        "toc_field_boundary_repaired": True,
    }


def find_index(paragraphs: list[ET.Element], predicate, start: int = 0) -> int | None:
    for index in range(start, len(paragraphs)):
        if predicate(paragraphs[index]):
            return index
    return None


def is_zh_abstract(paragraph: ET.Element) -> bool:
    value = compact(text_of(paragraph))
    return value in {"摘要", "中文摘要"} or value.startswith(("摘要：", "摘要:", "中文摘要：", "中文摘要:"))


def is_toc_title(paragraph: ET.Element) -> bool:
    return compact(text_of(paragraph)) == "目录"


def is_real_body_start(paragraph: ET.Element, styles: dict[str, dict[str, str]] | None = None) -> bool:
    text = text_of(paragraph).strip()
    style_id = paragraph_style_id(paragraph)
    return is_heading1_style(style_id, styles or {}) and bool(re.match(r"^第\s*[0-9一二三四五六七八九十]+章", text))


def is_real_body_start_text(text: str) -> bool:
    return bool(re.match(r"^第\s*[0-9一二三四五六七八九十]+章", (text or "").strip()))


def is_tail_heading(text: str) -> bool:
    return compact(text) in {"参考文献", "致谢"} or compact(text).startswith("附录")


def toc_level(label: str) -> int | None:
    stripped = label.strip()
    if not stripped:
        return None
    if re.match(r"^第[0-9一二三四五六七八九十]+章", stripped) or is_tail_heading(stripped):
        return 1
    if re.match(r"^\d+\.\d+\.\d+\s+", stripped):
        return 3
    if re.match(r"^\d+\.\d+\s+", stripped):
        return 2
    return None


def existing_final_sect_pr(body: ET.Element) -> ET.Element:
    sect = body.find("./w:sectPr", NS)
    if sect is not None:
        return sect
    sect = ET.SubElement(body, qn("sectPr"))
    return sect


def remove_paragraph_section_breaks(paragraphs: list[ET.Element]) -> int:
    removed = 0
    for paragraph in paragraphs:
        ppr = paragraph.find("./w:pPr", NS)
        if ppr is None:
            continue
        for sect in list(ppr.findall("./w:sectPr", NS)):
            ppr.remove(sect)
            removed += 1
    return removed


def element_has_toc_field_surface(element: ET.Element) -> bool:
    instr = " ".join(
        [
            *(node.text or "" for node in element.findall(".//w:instrText", NS)),
            *(node.attrib.get(qn("instr"), "") for node in element.findall(".//w:fldSimple", NS)),
        ]
    )
    return bool(re.search(r"(^|\s)TOC(\s|$)", instr, flags=re.IGNORECASE))


def remove_legacy_toc_content_controls(body: ET.Element) -> int:
    removed = 0
    for child in list(body):
        if child_local_name(child) != "sdt":
            continue
        if element_has_toc_field_surface(child) or compact(text_of(child)).startswith("\u76ee\u5f55"):
            body.remove(child)
            removed += 1
    return removed


def get_refs(sect_pr: ET.Element, local: str) -> list[ET.Element]:
    return [deepcopy(node) for node in sect_pr.findall(f"./w:{local}", NS)]


def make_section(
    *,
    header_refs: list[ET.Element] | None = None,
    footer_refs: list[ET.Element] | None = None,
    layout_source: ET.Element | None = None,
    page_fmt: str | None = None,
    page_start: str | None = None,
    section_type: str | None = "nextPage",
) -> ET.Element:
    sect = ET.Element(qn("sectPr"))
    if header_refs:
        for ref in header_refs:
            sect.append(deepcopy(ref))
    if footer_refs:
        for ref in footer_refs:
            sect.append(deepcopy(ref))
    if section_type:
        node = ET.SubElement(sect, qn("type"))
        node.set(qn("val"), section_type)
    if layout_source is not None:
        for local_name in ("pgSz", "pgMar"):
            source_node = layout_source.find(f"./w:{local_name}", NS)
            if source_node is not None:
                sect.append(deepcopy(source_node))
    if sect.find("./w:pgSz", NS) is None:
        pg_sz = ET.SubElement(sect, qn("pgSz"))
        pg_sz.set(qn("w"), "11906")
        pg_sz.set(qn("h"), "16838")
    if sect.find("./w:pgMar", NS) is None:
        pg_mar = ET.SubElement(sect, qn("pgMar"))
        pg_mar.set(qn("top"), "1417")
        pg_mar.set(qn("right"), "1134")
        pg_mar.set(qn("bottom"), "1417")
        pg_mar.set(qn("left"), "1701")
        pg_mar.set(qn("header"), "850")
        pg_mar.set(qn("footer"), "992")
        pg_mar.set(qn("gutter"), "0")
    if page_fmt is not None:
        pg_num = ET.SubElement(sect, qn("pgNumType"))
        pg_num.set(qn("fmt"), page_fmt)
        if page_start is not None:
            pg_num.set(qn("start"), page_start)
    if layout_source is not None:
        for local_name in ("cols", "docGrid"):
            source_node = layout_source.find(f"./w:{local_name}", NS)
            if source_node is not None and sect.find(f"./w:{local_name}", NS) is None:
                sect.append(deepcopy(source_node))
    if sect.find("./w:cols", NS) is None:
        cols = ET.SubElement(sect, qn("cols"))
        cols.set(qn("space"), "425")
    if sect.find("./w:docGrid", NS) is None:
        doc_grid = ET.SubElement(sect, qn("docGrid"))
        doc_grid.set(qn("linePitch"), "312")
    return sect


def replace_paragraph_section(paragraph: ET.Element, sect_pr: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    existing = ppr.find("./w:sectPr", NS)
    if existing is not None:
        ppr.remove(existing)
    ppr.append(sect_pr)


def insert_empty_section_break_before(body: ET.Element, anchor: ET.Element, sect_pr: ET.Element) -> ET.Element:
    boundary = ET.Element(qn("p"))
    ppr = ET.SubElement(boundary, qn("pPr"))
    ppr.append(sect_pr)
    children = list(body)
    try:
        insert_at = children.index(anchor)
    except ValueError:
        insert_at = len(children)
    body.insert(insert_at, boundary)
    return boundary


def ensure_header_part(
    parts: dict[str, bytes],
    rels_root: ET.Element,
    content_types: ET.Element,
    *,
    target: str,
    text: str,
    rule_surface: bool = False,
) -> str:
    rid = ensure_relationship(rels_root, rel_type=HEADER_REL_TYPE, target=target)
    ensure_content_type_override(
        content_types,
        f"/word/{target}",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    )
    parts[f"word/{target}"] = make_header_xml(text, rule_surface=rule_surface)
    return rid


def next_numbered_part_target(
    parts: dict[str, bytes],
    rels_root: ET.Element,
    *,
    prefix: str,
    suffix: str,
) -> str:
    used: set[str] = set()
    pattern = re.compile(rf"{re.escape(prefix)}(\d+){re.escape(suffix)}$")
    max_index = 0
    for name in parts:
        candidate = name.removeprefix("word/")
        match = pattern.fullmatch(candidate)
        if match:
            used.add(candidate)
            max_index = max(max_index, int(match.group(1)))
    for rel in rels_root.findall("./rel:Relationship", NS):
        target = rel.get("Target") or ""
        match = pattern.fullmatch(target)
        if match:
            used.add(target)
            max_index = max(max_index, int(match.group(1)))
    next_index = max_index + 1
    while f"{prefix}{next_index}{suffix}" in used:
        next_index += 1
    return f"{prefix}{next_index}{suffix}"


def set_hyperlink_style_black(styles_root: ET.Element) -> int:
    changed = 0
    for style in styles_root.findall("./w:style", NS):
        style_id = style.get(qn("styleId"), "")
        name_node = style.find("./w:name", NS)
        style_name = name_node.get(qn("val"), "") if name_node is not None else ""
        if style_id != "Hyperlink" and style_name.lower() != "hyperlink":
            continue
        rpr = style.find("./w:rPr", NS)
        if rpr is None:
            rpr = ET.SubElement(style, qn("rPr"))
            changed += 1
        color = rpr.find("./w:color", NS)
        if color is None:
            color = ET.SubElement(rpr, qn("color"))
            changed += 1
        if color.get(qn("val")) != "000000":
            color.set(qn("val"), "000000")
            changed += 1
        for local in ("themeColor", "themeTint", "themeShade"):
            if qn(local) in color.attrib:
                del color.attrib[qn(local)]
                changed += 1
    return changed


def relation_id_int(rid: str) -> int:
    match = re.search(r"(\d+)$", rid or "")
    return int(match.group(1)) if match else 0


def ensure_content_type_override(types_root: ET.Element, part_name: str, content_type: str) -> None:
    for override in types_root.findall("./ct:Override", NS):
        if override.get("PartName") == part_name:
            override.set("ContentType", content_type)
            return
    node = ET.SubElement(types_root, f"{CT}Override")
    node.set("PartName", part_name)
    node.set("ContentType", content_type)


def ensure_relationship(
    rels_root: ET.Element,
    *,
    rel_type: str,
    target: str,
) -> str:
    for rel in rels_root.findall("./rel:Relationship", NS):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return rel.get("Id", "")
    max_id = max((relation_id_int(rel.get("Id", "")) for rel in rels_root.findall("./rel:Relationship", NS)), default=0)
    rid = f"rId{max_id + 1}"
    node = ET.SubElement(rels_root, f"{REL}Relationship")
    node.set("Id", rid)
    node.set("Type", rel_type)
    node.set("Target", target)
    return rid


def target_for_rid(rels_root: ET.Element, rid: str) -> str:
    for rel in rels_root.findall("./rel:Relationship", NS):
        if rel.get("Id") == rid:
            return rel.get("Target") or ""
    return ""


def make_ref(local: str, rid: str) -> ET.Element:
    node = ET.Element(qn(local))
    node.set(qn("type"), "default")
    node.set(f"{R}id", rid)
    return node


def set_default_header_ref(sect_pr: ET.Element, rid: str) -> None:
    for ref in list(sect_pr.findall("./w:headerReference", NS)):
        sect_pr.remove(ref)
    insert_at = 0
    for index, child in enumerate(list(sect_pr)):
        if child_local_name(child) == "footerReference":
            insert_at = index + 1
    sect_pr.insert(insert_at, make_ref("headerReference", rid))


def ensure_header_footer_refs(
    parts: dict[str, bytes],
    document_rels: ET.Element,
    content_types: ET.Element,
    final_sect: ET.Element,
    header_text: str,
) -> tuple[list[ET.Element], list[ET.Element], list[str], list[str]]:
    header_refs = get_refs(final_sect, "headerReference")
    footer_refs = get_refs(final_sect, "footerReference")
    created: list[str] = []
    updated: list[str] = []
    if not header_refs:
        rid = ensure_relationship(document_rels, rel_type=HEADER_REL_TYPE, target="header1.xml")
        header_refs = [make_ref("headerReference", rid)]
        created.append("word/header1.xml relationship")
    if not footer_refs:
        rid = ensure_relationship(document_rels, rel_type=FOOTER_REL_TYPE, target="footer1.xml")
        footer_refs = [make_ref("footerReference", rid)]
        created.append("word/footer1.xml relationship")
    ensure_content_type_override(
        content_types,
        "/word/header1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    )
    ensure_content_type_override(
        content_types,
        "/word/footer1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    )
    parts["word/header1.xml"] = make_header_xml(header_text)
    parts["word/footer1.xml"] = make_footer_xml()
    updated.extend(["word/header1.xml", "word/footer1.xml"])
    return header_refs, footer_refs, created, updated


def make_header_xml(text_value: str = FALLBACK_HEADER_TEXT, *, rule_surface: bool = False) -> bytes:
    root = ET.Element(qn("hdr"))
    paragraph = ET.SubElement(root, qn("p"))
    ppr = ET.SubElement(paragraph, qn("pPr"))
    jc = ET.SubElement(ppr, qn("jc"))
    jc.set(qn("val"), "center")
    run = ET.SubElement(paragraph, qn("r"))
    rpr = ET.SubElement(run, qn("rPr"))
    fonts = ET.SubElement(rpr, qn("rFonts"))
    fonts.set(qn("ascii"), "Times New Roman")
    fonts.set(qn("hAnsi"), "Times New Roman")
    fonts.set(qn("eastAsia"), "SimSun")
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), "21")
    text = ET.SubElement(run, qn("t"))
    text.text = text_value
    if rule_surface:
        rule_paragraph = ET.SubElement(root, qn("p"))
        rule_run = ET.SubElement(rule_paragraph, qn("r"))
        pict = ET.SubElement(rule_run, qn("pict"))
        shape = ET.SubElement(pict, f"{{{V_NS}}}shape")
        shape.set("id", "BodyHeaderRule")
        shape.set("type", "#_x0000_t32")
        shape.set("style", "width:450pt;height:0;position:absolute")
        stroke = ET.SubElement(shape, f"{{{V_NS}}}stroke")
        stroke.set("weight", "0.5pt")
        stroke.set("color", "#000000")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def make_footer_xml() -> bytes:
    root = ET.Element(qn("ftr"))
    paragraph = ET.SubElement(root, qn("p"))
    ppr = ET.SubElement(paragraph, qn("pPr"))
    jc = ET.SubElement(ppr, qn("jc"))
    jc.set(qn("val"), "center")
    add_field_run(paragraph, "begin", size="18")
    add_instr_run(paragraph, " PAGE ", size="18")
    add_field_run(paragraph, "separate", size="18")
    ensure_run_text(paragraph, "1", size="18")
    add_field_run(paragraph, "end", size="18")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_style(styles_root: ET.Element, style_id: str, name: str) -> ET.Element:
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("styleId")) == style_id:
            return style
    style = ET.SubElement(styles_root, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), style_id)
    name_node = ET.SubElement(style, qn("name"))
    name_node.set(qn("val"), name)
    return style


def ensure_style_child(style: ET.Element, local: str) -> ET.Element:
    child = style.find(f"./w:{local}", NS)
    if child is None:
        child = ET.SubElement(style, qn(local))
    return child


def replace_style_ppr(style: ET.Element, *, jc_val: str | None, first_line_chars: str | None, left: str | None, line: str | None, tab_level: int | None = None) -> None:
    remove_children(style, {"pPr"})
    ppr = ET.SubElement(style, qn("pPr"))
    if jc_val:
        jc = ET.SubElement(ppr, qn("jc"))
        jc.set(qn("val"), jc_val)
    if first_line_chars or left:
        ind = ET.SubElement(ppr, qn("ind"))
        if first_line_chars:
            ind.set(qn("firstLineChars"), first_line_chars)
            ind.set(qn("firstLine"), "567")
        if left:
            ind.set(qn("left"), left)
    if line:
        spacing = ET.SubElement(ppr, qn("spacing"))
        spacing.set(qn("line"), line)
        spacing.set(qn("lineRule"), "auto")
    if tab_level is not None:
        tabs = ET.SubElement(ppr, qn("tabs"))
        tab = ET.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "right")
        tab.set(qn("leader"), "dot")
        tab.set(qn("pos"), "9350")


def replace_style_rpr(style: ET.Element, *, size: str, bold: bool = False) -> None:
    remove_children(style, {"rPr"})
    rpr = ET.SubElement(style, qn("rPr"))
    fonts = ET.SubElement(rpr, qn("rFonts"))
    fonts.set(qn("ascii"), "Times New Roman")
    fonts.set(qn("hAnsi"), "Times New Roman")
    fonts.set(qn("cs"), "Times New Roman")
    fonts.set(qn("eastAsia"), "SimSun")
    if bold:
        ET.SubElement(rpr, qn("b"))
        ET.SubElement(rpr, qn("bCs"))
    color = ET.SubElement(rpr, qn("color"))
    color.set(qn("val"), "000000")
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), size)
    szcs = ET.SubElement(rpr, qn("szCs"))
    szcs.set(qn("val"), size)


def normalize_styles(styles_root: ET.Element) -> dict[str, object]:
    changes: dict[str, object] = {"styles_touched": []}
    body = ensure_style(styles_root, "BodyText", "Body Text")
    replace_style_ppr(body, jc_val="both", first_line_chars="200", left=None, line="360")
    replace_style_rpr(body, size="24")
    changes["styles_touched"].append("BodyText")
    heading_specs = [
        ("Heading1", "heading 1", "center", "30", True),
        ("Heading2", "heading 2", "left", "28", True),
        ("Heading3", "heading 3", "left", "24", True),
    ]
    for style_id, name, jc, size, bold in heading_specs:
        style = ensure_style(styles_root, style_id, name)
        replace_style_ppr(style, jc_val=jc, first_line_chars=None, left=None, line="360")
        replace_style_rpr(style, size=size, bold=bold)
        changes["styles_touched"].append(style_id)
    caption = ensure_style(styles_root, "Caption", "caption")
    replace_style_ppr(caption, jc_val="center", first_line_chars=None, left=None, line="360")
    replace_style_rpr(caption, size="21")
    changes["styles_touched"].append("Caption")
    for level, left in [(1, "0"), (2, "420"), (3, "840")]:
        style = ensure_style(styles_root, f"TOC{level}", f"toc {level}")
        replace_style_ppr(style, jc_val=None, first_line_chars=None, left=left, line="360", tab_level=level)
        replace_style_rpr(style, size="24")
        changes["styles_touched"].append(f"TOC{level}")
    hyperlink_changes = set_hyperlink_style_black(styles_root)
    if hyperlink_changes:
        changes["hyperlink_style_black_normalizations"] = hyperlink_changes
    return changes


def header_text_from_refs(parts: dict[str, bytes], rels_root: ET.Element, final_sect: ET.Element) -> str:
    rels = {rel.get("Id", ""): rel.get("Target", "") for rel in rels_root.findall("./rel:Relationship", NS)}
    candidates: list[str] = []
    for ref in final_sect.findall("./w:headerReference", NS):
        target = rels.get(ref.get(f"{R}id", ""))
        if not target:
            continue
        data = parts.get(f"word/{target}")
        if not data:
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        text = text_of(root).strip()
        if text:
            candidates.append(text)
    if candidates:
        return candidates[0]
    for name, data in sorted(parts.items()):
        if not re.fullmatch(r"word/header\d+\.xml", name):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        text = text_of(root).strip()
        if text:
            return text
    return FALLBACK_HEADER_TEXT


def header_text_from_sections(
    parts: dict[str, bytes],
    rels_root: ET.Element,
    section_properties: list[ET.Element],
    preferred_indices: list[int],
) -> str:
    rels = {rel.get("Id", ""): rel.get("Target", "") for rel in rels_root.findall("./rel:Relationship", NS)}
    ordered_sections: list[ET.Element] = []
    for index in preferred_indices:
        if 0 <= index < len(section_properties):
            ordered_sections.append(section_properties[index])
    for section in section_properties:
        if section not in ordered_sections:
            ordered_sections.append(section)
    for section in ordered_sections:
        for ref in section.findall("./w:headerReference", NS):
            target = rels.get(ref.get(f"{R}id", ""))
            if not target:
                continue
            data = parts.get(f"word/{target}")
            if not data:
                continue
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                continue
            text = text_of(root).strip()
            if text:
                return text
    for name, data in sorted(parts.items()):
        if not re.fullmatch(r"word/header\d+\.xml", name):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        text = text_of(root).strip()
        if text:
            return text
    return FALLBACK_HEADER_TEXT


def header_text_from_section(
    parts: dict[str, bytes],
    rels_root: ET.Element,
    section_properties: list[ET.Element],
    section_index: int,
) -> str:
    if section_index < 0 or section_index >= len(section_properties):
        return ""
    rels = {rel.get("Id", ""): rel.get("Target", "") for rel in rels_root.findall("./rel:Relationship", NS)}
    for ref in section_properties[section_index].findall("./w:headerReference", NS):
        target = rels.get(ref.get(f"{R}id", ""))
        if not target:
            continue
        data = parts.get(f"word/{target}")
        if not data:
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        text = text_of(root).strip()
        if text:
            return text
    return ""


def apply_toc_tabs(paragraph: ET.Element, level: int) -> None:
    ppr = ensure_ppr(paragraph)
    for tabs in list(ppr.findall("./w:tabs", NS)):
        ppr.remove(tabs)
    tabs = ET.Element(qn("tabs"))
    tab = ET.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    tab.set(qn("pos"), "9350")
    insert_at = len(list(ppr))
    for index, child in enumerate(list(ppr)):
        if child_local_name(child) in {"rPr", "sectPr"}:
            insert_at = index
            break
    ppr.insert(insert_at, tabs)


def body_heading_page_map(paragraphs: list[ET.Element], body_start: int) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for index, paragraph in enumerate(paragraphs[body_start:], start=body_start):
        text = text_of(paragraph).strip()
        style_id = paragraph_style_id(paragraph)
        if style_id in {"Heading1", "Heading2", "Heading3"} or is_tail_heading(text):
            # This is a cache estimate for compatibility. Word/WPS can update
            # the live TOC because the field instruction and headings are real.
            page = str(max(1, round((index - body_start) / 7.0) + 1))
            mapping.setdefault(text, page)
    return mapping


def rewrite_toc_cache(paragraphs: list[ET.Element], toc_title: int, body_start: int) -> dict[str, object]:
    page_map = body_heading_page_map(paragraphs, body_start)
    toc_entries = []
    first_entry: ET.Element | None = None
    last_entry: ET.Element | None = None
    for index in range(toc_title + 1, body_start):
        paragraph = paragraphs[index]
        raw_label = text_of(paragraph).strip()
        label = raw_label
        cached_page = None
        if paragraph.find(".//w:tab", NS) is not None:
            page_match = re.search(r"(\d+)\s*$", raw_label)
            if page_match:
                label = raw_label[: page_match.start()].rstrip()
                cached_page = page_match.group(1)
        level = toc_level(label)
        if level is None:
            continue
        first_entry = first_entry or paragraph
        last_entry = paragraph
        page = page_map.get(label, cached_page or "1")
        set_paragraph_style(paragraph, f"TOC{level}")
        apply_toc_tabs(paragraph, level)
        clear_paragraph_runs(paragraph)
        ensure_run_text(paragraph, label)
        add_tab_run(paragraph)
        ensure_run_text(paragraph, page)
        toc_entries.append({"index": index, "level": level, "label": label, "page_cache": page})
    if first_entry is not None and last_entry is not None:
        insert_at = 1 if first_entry.find("./w:pPr", NS) is not None else 0
        first_entry.insert(insert_at, add_field_run(ET.Element("tmp"), "begin", locked=True))
        first_entry.insert(insert_at + 1, add_instr_run(ET.Element("tmp"), ' TOC \\o "1-3" \\h \\z \\u '))
        first_entry.insert(insert_at + 2, add_field_run(ET.Element("tmp"), "separate"))
        last_entry.append(add_field_run(ET.Element("tmp"), "end"))
    return {"toc_cache_entry_count": len(toc_entries), "toc_entries_sample": toc_entries[:10]}


def normalize_paragraph_style_usage(paragraphs: list[ET.Element], body_start: int) -> dict[str, int]:
    changed = {"sgb_to_bodytext": 0, "body_no_style_to_bodytext": 0}
    for index, paragraph in enumerate(paragraphs):
        text = text_of(paragraph).strip()
        style_id = paragraph_style_id(paragraph)
        if style_id.lower().startswith("sgb"):
            set_paragraph_style(paragraph, "BodyText")
            changed["sgb_to_bodytext"] += 1
        elif index >= body_start and text and not style_id:
            set_paragraph_style(paragraph, "BodyText")
            changed["body_no_style_to_bodytext"] += 1
    return changed


def paragraph_section_index_by_zero_based_paragraph(paragraphs: list[ET.Element]) -> dict[int, int]:
    result: dict[int, int] = {}
    running_section_index = -1
    for index, paragraph in enumerate(paragraphs):
        result[index] = running_section_index + 1
        if paragraph.find("./w:pPr/w:sectPr", NS) is not None:
            running_section_index += 1
    return result


def body_direct_section_properties(body: ET.Element) -> list[ET.Element]:
    sections: list[ET.Element] = []
    for child in list(body):
        if child.tag == qn("p"):
            sect = child.find("./w:pPr/w:sectPr", NS)
            if sect is not None:
                sections.append(sect)
        elif child.tag == qn("sectPr"):
            sections.append(child)
    return sections


def default_header_rid(sect_pr: ET.Element) -> str:
    for ref in sect_pr.findall("./w:headerReference", NS):
        if (ref.get(qn("type")) or "default") == "default":
            return ref.get(f"{R}id", "")
    return ""


def repair_toc_body_section_headers_parts(parts: dict[str, bytes]) -> dict[str, object]:
    document_root = ET.fromstring(parts["word/document.xml"])
    styles_root = ET.fromstring(parts["word/styles.xml"])
    rels_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    content_types = ET.fromstring(parts["[Content_Types].xml"])
    body = body_element(document_root)
    removed_legacy_toc_content_controls = remove_legacy_toc_content_controls(body)
    paragraphs = direct_paragraphs(body)
    styles = style_metadata_map(styles_root)
    toc_title = find_index(paragraphs, is_toc_title)
    body_start = find_index(paragraphs, lambda paragraph: is_real_body_start(paragraph, styles), start=(toc_title or 0) + 1)
    if toc_title is None:
        raise RuntimeError("TOC title not found")
    if body_start is None:
        raise RuntimeError("real body start heading not found")
    section_properties = body_direct_section_properties(body)
    section_map = paragraph_section_index_by_zero_based_paragraph(paragraphs)
    toc_section_index = section_map.get(toc_title)
    body_section_index = section_map.get(body_start)
    if toc_section_index is None or body_section_index is None:
        raise RuntimeError("could not map TOC/body paragraphs to sections")
    if toc_section_index < 0 or toc_section_index >= len(section_properties):
        raise RuntimeError(f"TOC section index out of range: {toc_section_index}")
    if body_section_index < 0 or body_section_index >= len(section_properties):
        raise RuntimeError(f"body section index out of range: {body_section_index}")
    if toc_section_index == body_section_index:
        full_changes = repair_docx_parts(parts)
        return {
            "operation_escalated_to_full_repair": True,
            "reason": "TOC and first body heading were in the same section",
            "removed_legacy_toc_content_controls_before_escalation": removed_legacy_toc_content_controls,
            **full_changes,
        }

    fallback_header_text = header_text_from_sections(
        parts,
        rels_root,
        section_properties,
        [body_section_index, toc_section_index],
    )
    toc_title_text = text_of(paragraphs[toc_title]).strip() or "目 录"
    body_title_text = text_of(paragraphs[body_start]).strip()
    toc_header_text = header_text_from_section(parts, rels_root, section_properties, toc_section_index) or toc_title_text
    body_header_text = header_text_from_section(parts, rels_root, section_properties, body_section_index) or fallback_header_text
    if compact(toc_header_text) == compact(body_title_text) or is_real_body_start_text(toc_header_text):
        toc_header_text = toc_title_text
    if not body_header_text:
        body_header_text = fallback_header_text
    toc_field_changes = normalize_toc_field_boundary(paragraphs, toc_title, body_start)
    toc_entry_section_indices = sorted(
        {
            section_map.get(index)
            for index in range(toc_title + 1, body_start)
            if text_of(paragraphs[index]).strip()
        }
        - {None, toc_section_index, body_section_index}
    )
    toc_header_target = next_numbered_part_target(parts, rels_root, prefix="header", suffix=".xml")
    toc_entry_header_target = next_numbered_part_target(
        {**parts, f"word/{toc_header_target}": b""},
        rels_root,
        prefix="header",
        suffix=".xml",
    )
    body_header_target = next_numbered_part_target(
        {
            **parts,
            f"word/{toc_header_target}": b"",
            f"word/{toc_entry_header_target}": b"",
        },
        rels_root,
        prefix="header",
        suffix=".xml",
    )
    toc_header_rid = ensure_relationship(rels_root, rel_type=HEADER_REL_TYPE, target=toc_header_target)
    toc_entry_header_rid = ensure_relationship(rels_root, rel_type=HEADER_REL_TYPE, target=toc_entry_header_target)
    body_header_rid = ensure_relationship(rels_root, rel_type=HEADER_REL_TYPE, target=body_header_target)
    parts[f"word/{toc_header_target}"] = make_header_xml(toc_header_text)
    parts[f"word/{toc_entry_header_target}"] = make_header_xml("")
    parts[f"word/{body_header_target}"] = make_header_xml(body_header_text, rule_surface=True)
    for target in (toc_header_target, toc_entry_header_target, body_header_target):
        ensure_content_type_override(
            content_types,
            f"/word/{target}",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        )
    set_default_header_ref(section_properties[toc_section_index], toc_header_rid)
    for entry_section_index in toc_entry_section_indices:
        if entry_section_index is not None and 0 <= entry_section_index < len(section_properties):
            set_default_header_ref(section_properties[entry_section_index], toc_entry_header_rid)
    set_default_header_ref(section_properties[body_section_index], body_header_rid)
    postcondition = {
        "toc_default_header_rid": default_header_rid(section_properties[toc_section_index]),
        "toc_entry_default_header_rids": [
            default_header_rid(section_properties[index])
            for index in toc_entry_section_indices
            if index is not None and 0 <= index < len(section_properties)
        ],
        "body_default_header_rid": default_header_rid(section_properties[body_section_index]),
        "toc_body_rids_distinct": toc_header_rid != body_header_rid,
        "toc_entry_body_rids_distinct": toc_entry_header_rid != body_header_rid,
    }
    if postcondition["toc_default_header_rid"] != toc_header_rid:
        raise RuntimeError("TOC section default header relationship was not set")
    if any(rid != toc_entry_header_rid for rid in postcondition["toc_entry_default_header_rids"]):
        raise RuntimeError("TOC entry section default header relationship was not set")
    if postcondition["body_default_header_rid"] != body_header_rid:
        raise RuntimeError("body section default header relationship was not set")
    if not postcondition["toc_body_rids_distinct"] or not postcondition["toc_entry_body_rids_distinct"]:
        raise RuntimeError("TOC/body header relationships are not distinct")

    parts["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    parts["word/_rels/document.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    parts["[Content_Types].xml"] = serialize_content_types_root(content_types)
    return {
        "toc_title_paragraph_index": toc_title,
        "body_start_paragraph_index": body_start,
        "toc_section_index": toc_section_index,
        "body_section_index": body_section_index,
        "toc_header_target": toc_header_target,
        "toc_header_rid": toc_header_rid,
        "toc_header_text": toc_header_text,
        "toc_entry_section_indices": toc_entry_section_indices,
        "toc_entry_header_target": toc_entry_header_target,
        "toc_entry_header_rid": toc_entry_header_rid,
        "body_header_target": body_header_target,
        "body_header_rid": body_header_rid,
        "body_header_text": body_header_text,
        "body_header_rule_surface": "v:shape/v:stroke",
        "postcondition": postcondition,
        **toc_field_changes,
        "updated_parts": [
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
            f"word/{toc_header_target}",
            f"word/{toc_entry_header_target}",
            f"word/{body_header_target}",
        ],
    }


def repair_docx_parts(parts: dict[str, bytes]) -> dict[str, object]:
    document_root = ET.fromstring(parts["word/document.xml"])
    styles_root = ET.fromstring(parts["word/styles.xml"])
    rels_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    content_types = ET.fromstring(parts["[Content_Types].xml"])
    styles = style_metadata_map(styles_root)
    body = body_element(document_root)
    removed_legacy_toc_content_controls = remove_legacy_toc_content_controls(body)
    paragraphs = direct_paragraphs(body)
    zh_abs = find_index(paragraphs, is_zh_abstract)
    toc_title = find_index(paragraphs, is_toc_title)
    body_start = find_index(paragraphs, lambda paragraph: is_real_body_start(paragraph, styles), start=(toc_title or 0) + 1)
    if zh_abs is None:
        raise RuntimeError("Chinese abstract title not found")
    if toc_title is None:
        raise RuntimeError("TOC title not found")
    if body_start is None:
        raise RuntimeError("real body start heading not found")
    if not (zh_abs < toc_title < body_start):
        raise RuntimeError(f"front matter order invalid: zh_abs={zh_abs}, toc={toc_title}, body_start={body_start}")

    section_properties = document_root.findall(".//w:sectPr", NS)
    final_sect = existing_final_sect_pr(body)
    cover_layout = section_properties[0] if len(section_properties) >= 1 else final_sect
    toc_layout = section_properties[1] if len(section_properties) >= 2 else cover_layout
    body_layout = final_sect
    header_text = header_text_from_refs(parts, rels_root, final_sect)
    header_refs, footer_refs, created_refs, updated_parts = ensure_header_footer_refs(
        parts, rels_root, content_types, final_sect, header_text
    )
    footer_rid = ensure_relationship(rels_root, rel_type=FOOTER_REL_TYPE, target="footer1.xml")
    ensure_content_type_override(
        content_types,
        "/word/footer1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    )
    parts["word/footer1.xml"] = make_footer_xml()
    footer_refs = [make_ref("footerReference", footer_rid)]
    toc_header_rid = ensure_header_part(
        parts,
        rels_root,
        content_types,
        target="header2.xml",
        text="",
    )
    body_header_rid = ensure_header_part(
        parts,
        rels_root,
        content_types,
        target="header3.xml",
        text=header_text,
        rule_surface=True,
    )
    toc_header_refs = [make_ref("headerReference", toc_header_rid)]
    body_header_refs = [make_ref("headerReference", body_header_rid)]
    updated_parts.extend(["word/header2.xml", "word/header3.xml"])
    removed_existing_section_breaks = remove_paragraph_section_breaks(paragraphs)
    cover_end = max(0, zh_abs - 1)
    replace_paragraph_section(
        paragraphs[cover_end],
        make_section(section_type="nextPage", layout_source=cover_layout),
    )
    toc_body_boundary = insert_empty_section_break_before(
        body,
        paragraphs[body_start],
        make_section(
            header_refs=toc_header_refs,
            footer_refs=footer_refs,
            layout_source=toc_layout,
            page_fmt="lowerRoman",
            page_start="1",
            section_type="nextPage",
        ),
    )
    body.remove(final_sect)
    body.append(
        make_section(
            header_refs=body_header_refs,
            footer_refs=footer_refs,
            layout_source=body_layout,
            page_fmt="decimal",
            page_start="1",
            section_type=None,
        )
    )

    hyperlink_changes = set_hyperlink_style_black(styles_root)
    style_changes = {"hyperlink_style_black_normalizations": hyperlink_changes}
    paragraph_style_changes = normalize_paragraph_style_usage(paragraphs, body_start)
    toc_changes = normalize_toc_field_boundary(paragraphs, toc_title, body_start)

    parts["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    parts["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
    parts["word/_rels/document.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    parts["[Content_Types].xml"] = serialize_content_types_root(content_types)
    return {
        "surfaces": {
            "cover_end_paragraph_index": cover_end,
            "zh_abstract_paragraph_index": zh_abs,
            "toc_title_paragraph_index": toc_title,
            "frontmatter_end_paragraph_index": body_start,
            "toc_body_boundary_field_free": not bool(text_of(toc_body_boundary).strip() or toc_body_boundary.findall(".//w:fldChar", NS) or toc_body_boundary.findall(".//w:instrText", NS) or toc_body_boundary.findall(".//w:fldSimple", NS)),
            "body_start_paragraph_index": body_start,
        },
        "removed_legacy_toc_content_controls": removed_legacy_toc_content_controls,
        "removed_existing_paragraph_section_breaks": removed_existing_section_breaks,
        "created_relationships": created_refs,
        "updated_parts": sorted({"word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels", "[Content_Types].xml", *updated_parts}),
        "style_changes": style_changes,
        "paragraph_style_changes": paragraph_style_changes,
        "toc_changes": toc_changes,
    }


def load_parts(path: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo]]:
    with zipfile.ZipFile(path, "r") as zf:
        infos = zf.infolist()
        parts = {info.filename: zf.read(info.filename) for info in infos}
    required = ["word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"]
    missing = [name for name in required if name not in parts]
    if missing:
        raise RuntimeError(f"DOCX missing required parts: {missing}")
    return parts, infos


def write_parts(input_infos: list[zipfile.ZipInfo], parts: dict[str, bytes], output_docx: Path) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        temp_output = Path(td) / output_docx.name
        written = set()
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in input_infos:
                data = parts[info.filename]
                zout.writestr(info, data)
                written.add(info.filename)
            for name, data in sorted(parts.items()):
                if name not in written:
                    zout.writestr(name, data)
        if output_docx.exists():
            output_docx.unlink()
        shutil.move(str(temp_output), output_docx)


def repair(input_docx: Path, output_docx: Path, *, operation: str = "full") -> dict[str, object]:
    if input_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a new path")
    parts, infos = load_parts(input_docx)
    if operation == "full":
        changes = repair_docx_parts(parts)
    elif operation == "toc-body-section-headers":
        changes = repair_toc_body_section_headers_parts(parts)
    else:
        raise RuntimeError(f"unknown repair operation: {operation}")
    write_parts(infos, parts, output_docx)
    return {
        "schema": "graduation-project-builder.docx-whole-format-structure-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "generator": "scripts/repair_docx_whole_format_structure.py",
        "operation": operation,
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "changes": changes,
        "repair_verdict": "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument(
        "--operation",
        choices=["full", "toc-body-section-headers"],
        default="full",
        help="Bounded structure repair operation. Default full preserves legacy behavior.",
    )
    args = parser.parse_args()
    report = repair(Path(args.input_docx).resolve(), Path(args.output_docx).resolve(), operation=args.operation)
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
