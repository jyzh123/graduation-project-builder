#!/usr/bin/env python3
"""Repair bibliography entry run structure from a locked DOCX template."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable

from audit_docx_font_encoding import (
    NS,
    WESTERN_REFERENCE_CHARS,
    bibliography_entry_paragraph_elements,
    contains_cjk,
    effective_run_signature,
    extract_reference_font_policy,
    font_size_name_for_half_points,
    load_style_tables,
    paragraph_text,
    paragraph_style_id,
    resolve_font_size_half_points,
    script_class,
    w,
)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


ET.register_namespace("w", NS["w"])
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_document_xml(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def load_xml_part(docx_path: Path, part_name: str) -> ET.Element | None:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            return ET.fromstring(zf.read(part_name))
    except KeyError:
        return None


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").lower()


def paragraph_num_id(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:numPr/w:numId", NS)
    return node.get(w("val"), "") if node is not None else ""


def local_bibliography_entry_paragraphs(root: ET.Element) -> list[tuple[int, ET.Element]]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    entries: list[tuple[int, ET.Element]] = []
    unnumbered_entries: list[tuple[int, ET.Element]] = []
    in_references = False
    for body_index, child in enumerate(list(body)):
        if child.tag != w("p"):
            continue
        text = paragraph_text(child).strip()
        compact = compact_text(text)
        if compact in {compact_text("\u53c2\u8003\u6587\u732e"), "references", "bibliography"}:
            in_references = True
            continue
        if not in_references:
            continue
        if compact in {
            compact_text("\u81f4\u8c22"),
            compact_text("\u9644\u5f55"),
            "acknowledgements",
            "acknowledgments",
            "appendix",
        }:
            break
        if not text:
            continue
        if paragraph_num_id(child) or re.match(r"^\s*\[\d{1,3}\]", text):
            entries.append((body_index, child))
            continue
        if looks_like_reference_entry(text):
            unnumbered_entries.append((body_index, child))
    return entries or unnumbered_entries


def looks_like_reference_entry(text: str) -> bool:
    """Detect real GB/T-style bibliography entries whose visible number was lost."""
    stripped = (text or "").strip()
    if not stripped:
        return False
    compact = compact_text(stripped)
    if compact in {
        compact_text("\u53c2\u8003\u6587\u732e"),
        compact_text("\u81f4\u8c22"),
        compact_text("\u9644\u5f55"),
        "references",
        "bibliography",
        "acknowledgements",
        "acknowledgments",
        "appendix",
    }:
        return False
    return bool(
        re.search(r"\[[JMCDEB/OL]{1,8}\]", stripped)
        or re.search(r"\[EB/OL\]", stripped)
        or re.search(r"\[C\]//", stripped)
        or re.search(r"https?://", stripped, re.IGNORECASE)
    )


def bibliography_entry_paragraphs(root: ET.Element) -> list[tuple[int, ET.Element]]:
    return bibliography_entry_paragraph_elements(root) or local_bibliography_entry_paragraphs(root)


def references_title_paragraph(root: ET.Element) -> tuple[int, ET.Element]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    matches: list[tuple[int, ET.Element]] = []
    instruction_matches: list[tuple[int, ET.Element]] = []
    for body_index, child in enumerate(list(body)):
        if child.tag != w("p"):
            continue
        compact = compact_text(paragraph_text(child))
        if compact in {compact_text("\u53c2\u8003\u6587\u732e"), "references", "bibliography"}:
            matches.append((body_index, child))
        elif compact.startswith(compact_text("\u53c2\u8003\u6587\u732e")):
            instruction_matches.append((body_index, child))
    if not matches:
        if instruction_matches:
            return instruction_matches[-1]
        raise RuntimeError("no references title paragraph found")
    return matches[-1]


def remove_child(parent: ET.Element | None, tag: str) -> None:
    if parent is None:
        return
    for child in list(parent):
        if child.tag == w(tag):
            parent.remove(child)


def remove_direct_emphasis(rpr: ET.Element | None) -> None:
    if rpr is None:
        return
    for tag in ("b", "bCs", "i", "iCs", "u", "highlight"):
        remove_child(rpr, tag)
    for color in list(rpr.findall("./w:color", NS)):
        value = (color.get(w("val")) or "").upper()
        if value in {"FF0000", "C00000", "0000FF", "0000CC", "BLUE", "RED"}:
            rpr.remove(color)


def force_cjk_slots_to_east_asia(rpr: ET.Element | None) -> None:
    if rpr is None:
        return
    fonts = rpr.find("./w:rFonts", NS)
    if fonts is None:
        return
    east_asia = fonts.get(w("eastAsia"))
    if not east_asia:
        return
    for key in ("ascii", "hAnsi", "cs"):
        fonts.set(w(key), east_asia)
        fonts.attrib.pop(w(f"{key}Theme"), None)


def has_page_break(paragraph: ET.Element) -> bool:
    return any(br.get(w("type"), "textWrapping") == "page" for br in paragraph.findall(".//w:br", NS))


def has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def has_nontext_payload(paragraph: ET.Element) -> bool:
    return any(paragraph.find(f".//w:{tag}", NS) is not None for tag in ("drawing", "pict", "object", "fldChar", "instrText"))


def repair_references_title(root: ET.Element, template_docx: Path) -> dict[str, object]:
    template_root = load_document_xml(template_docx)
    template_index, template_title = references_title_paragraph(template_root)
    actual_index, actual_title = references_title_paragraph(root)
    actual_text = paragraph_text(actual_title)
    template_ppr = copy.deepcopy(template_title.find("./w:pPr", NS))
    if template_ppr is not None:
        existing_ppr = actual_title.find("./w:pPr", NS)
        if existing_ppr is not None:
            actual_title.remove(existing_ppr)
        actual_title.insert(0, template_ppr)
    template_rpr = None
    for run in template_title.findall("./w:r", NS):
        if paragraph_text(run).strip():
            template_rpr = copy.deepcopy(run.find("./w:rPr", NS)) if run.find("./w:rPr", NS) is not None else ET.Element(w("rPr"))
            break
    changed_runs = 0
    if template_rpr is not None:
        for run in actual_title.findall("./w:r", NS):
            if not paragraph_text(run):
                continue
            existing = run.find("./w:rPr", NS)
            if existing is not None:
                run.remove(existing)
            run.insert(0, copy.deepcopy(template_rpr))
            changed_runs += 1
    body = root.find(".//w:body", NS)
    removed_preceding_break = False
    if body is not None:
        children = list(body)
        if actual_index > 0:
            previous = children[actual_index - 1]
            if (
                previous.tag == w("p")
                and not paragraph_text(previous).strip()
                and has_page_break(previous)
                and not has_section_break(previous)
                and not has_nontext_payload(previous)
            ):
                body.remove(previous)
                removed_preceding_break = True
    if paragraph_text(actual_title) != actual_text:
        raise RuntimeError("references title text changed during title format repair")
    return {
        "template_title_body_child_index": template_index,
        "actual_title_body_child_index_before": actual_index,
        "title_run_count_reformatted": changed_runs,
        "preceding_explicit_page_break_removed": removed_preceding_break,
    }


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    node = parent.find(f"./w:{tag}", NS)
    if node is None:
        node = ET.SubElement(parent, w(tag))
    return node


def set_ppr_numbering(ppr: ET.Element, num_id: str) -> None:
    num_pr = ensure_child(ppr, "numPr")
    ilvl = num_pr.find("./w:ilvl", NS)
    if ilvl is None:
        ilvl = ET.Element(w("ilvl"))
        num_pr.insert(0, ilvl)
    ilvl.set(w("val"), "0")
    num_id_node = num_pr.find("./w:numId", NS)
    if num_id_node is None:
        num_id_node = ET.SubElement(num_pr, w("numId"))
    num_id_node.set(w("val"), str(num_id))


def set_reference_entry_outdent(ppr: ET.Element, left_twips: str, hanging_twips: str) -> None:
    """Adjust the list paragraph geometry while preserving the reference run model."""
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.SubElement(ppr, w("ind"))
    ind.set(w("left"), left_twips)
    ind.set(w("leftChars"), "0")
    ind.set(w("hanging"), hanging_twips)
    ind.set(w("firstLineChars"), "0")


def first_template_numbered_bibliography_ppr(template_docx: Path) -> tuple[ET.Element, str]:
    root = load_document_xml(template_docx)
    for _, paragraph in local_bibliography_entry_paragraphs(root):
        num_id = paragraph_num_id(paragraph)
        ppr = paragraph.find("./w:pPr", NS)
        if num_id and ppr is not None:
            return copy.deepcopy(ppr), num_id
    raise RuntimeError(f"no automatic bibliography numbering model found in template: {template_docx}")


def load_numbering_root(docx_path: Path) -> ET.Element:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            return ET.fromstring(zf.read("word/numbering.xml"))
    except KeyError:
        return ET.Element(w("numbering"))


def child_by_attr(root: ET.Element, tag: str, attr: str, value: str) -> ET.Element | None:
    for child in root.findall(f"./w:{tag}", NS):
        if child.get(w(attr)) == value:
            return child
    return None


def next_numeric_id(nodes: list[ET.Element], attr: str) -> str:
    max_seen = 0
    for node in nodes:
        value = node.get(w(attr), "")
        if value.isdigit():
            max_seen = max(max_seen, int(value))
    return str(max_seen + 1)


def abstract_num_id_for_num(numbering_root: ET.Element, num_id: str) -> str:
    num = child_by_attr(numbering_root, "num", "numId", num_id)
    if num is None:
        raise RuntimeError(f"template numbering numId {num_id} not found")
    abstract = num.find("./w:abstractNumId", NS)
    if abstract is None or not abstract.get(w("val")):
        raise RuntimeError(f"template numbering numId {num_id} lacks abstractNumId")
    return abstract.get(w("val")) or ""


def numbering_lvl_text(numbering_root: ET.Element, num_id: str) -> str:
    try:
        abstract_id = abstract_num_id_for_num(numbering_root, num_id)
    except RuntimeError:
        return ""
    abstract = child_by_attr(numbering_root, "abstractNum", "abstractNumId", abstract_id)
    if abstract is None:
        return ""
    lvl_text = abstract.find(".//w:lvlText", NS)
    return lvl_text.get(w("val"), "") if lvl_text is not None else ""


def set_numbering_level_outdent(numbering_root: ET.Element, num_id: str, left_twips: str, hanging_twips: str) -> bool:
    abstract_id = abstract_num_id_for_num(numbering_root, num_id)
    abstract = child_by_attr(numbering_root, "abstractNum", "abstractNumId", abstract_id)
    if abstract is None:
        return False
    lvl = abstract.find("./w:lvl", NS)
    if lvl is None:
        return False
    ppr = lvl.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.SubElement(lvl, w("pPr"))
    set_reference_entry_outdent(ppr, left_twips, hanging_twips)
    return True


def merge_template_bibliography_numbering(
    source_docx: Path,
    template_docx: Path,
    template_num_id: str,
) -> tuple[bytes, str, dict[str, object]]:
    target_root = load_numbering_root(source_docx)
    template_root = load_numbering_root(template_docx)
    existing = child_by_attr(target_root, "num", "numId", template_num_id)
    if existing is not None and numbering_lvl_text(target_root, template_num_id) == "[%1]":
        return (
            ET.tostring(target_root, encoding="utf-8", xml_declaration=True),
            template_num_id,
            {"mode": "reuse-existing", "numId": template_num_id, "lvlText": "[%1]"},
        )

    template_abs_id = abstract_num_id_for_num(template_root, template_num_id)
    template_num = child_by_attr(template_root, "num", "numId", template_num_id)
    template_abs = child_by_attr(template_root, "abstractNum", "abstractNumId", template_abs_id)
    if template_num is None or template_abs is None:
        raise RuntimeError("template bibliography numbering model is incomplete")

    target_abs_id = template_abs_id
    if child_by_attr(target_root, "abstractNum", "abstractNumId", target_abs_id) is not None:
        target_abs_id = next_numeric_id(target_root.findall("./w:abstractNum", NS), "abstractNumId")
    target_num_id = template_num_id
    if child_by_attr(target_root, "num", "numId", target_num_id) is not None:
        target_num_id = next_numeric_id(target_root.findall("./w:num", NS), "numId")

    abs_clone = copy.deepcopy(template_abs)
    abs_clone.set(w("abstractNumId"), target_abs_id)
    num_clone = copy.deepcopy(template_num)
    num_clone.set(w("numId"), target_num_id)
    abstract_ref = num_clone.find("./w:abstractNumId", NS)
    if abstract_ref is None:
        abstract_ref = ET.SubElement(num_clone, w("abstractNumId"))
    abstract_ref.set(w("val"), target_abs_id)
    target_root.append(abs_clone)
    target_root.append(num_clone)
    return (
        ET.tostring(target_root, encoding="utf-8", xml_declaration=True),
        target_num_id,
        {
            "mode": "copied-from-template",
            "template_numId": template_num_id,
            "template_abstractNumId": template_abs_id,
            "numId": target_num_id,
            "abstractNumId": target_abs_id,
            "lvlText": numbering_lvl_text(target_root, target_num_id),
        },
    )


def ensure_rpr_child(rpr: ET.Element, tag: str) -> ET.Element:
    child = rpr.find(f"./w:{tag}", NS)
    if child is not None:
        return child
    child = ET.Element(w(tag))
    rpr.insert(0, child)
    return child


def ensure_direct_font_policy(rpr: ET.Element | None, policy: dict[str, str]) -> None:
    if rpr is None:
        return
    east_asia = policy.get("eastAsia", "")
    latin = policy.get("latin", "")
    if not east_asia and not latin:
        return
    fonts = ensure_rpr_child(rpr, "rFonts")
    if east_asia and not fonts.get(w("eastAsia")):
        fonts.set(w("eastAsia"), east_asia)
        fonts.attrib.pop(w("eastAsiaTheme"), None)
    if latin:
        for key in ("ascii", "hAnsi", "cs"):
            if not fonts.get(w(key)):
                fonts.set(w(key), latin)
            fonts.attrib.pop(w(f"{key}Theme"), None)


def has_required_reference_fonts(latin_rpr: ET.Element, cjk_rpr: ET.Element) -> bool:
    latin_fonts = latin_rpr.find("./w:rFonts", NS)
    cjk_fonts = cjk_rpr.find("./w:rFonts", NS)
    if latin_fonts is None or cjk_fonts is None:
        return False
    return bool(
        cjk_fonts.get(w("eastAsia"))
        and latin_fonts.get(w("ascii"))
        and latin_fonts.get(w("hAnsi"))
        and latin_fonts.get(w("cs"))
    )


def latin_font_score(rpr: ET.Element) -> int:
    fonts = rpr.find("./w:rFonts", NS)
    if fonts is None:
        return 0
    return sum(1 for key in ("ascii", "hAnsi", "cs") if fonts.get(w(key)))


def cjk_font_score(rpr: ET.Element) -> int:
    fonts = rpr.find("./w:rFonts", NS)
    if fonts is None:
        return 0
    return 1 if fonts.get(w("eastAsia")) else 0


def materialize_effective_run_signature(rpr: ET.Element, effective: dict[str, str]) -> None:
    """Write style-inherited donor font/size slots into a direct rPr model."""
    fonts = ensure_rpr_child(rpr, "rFonts")
    for key in ("eastAsia", "ascii", "hAnsi", "cs"):
        value = effective.get(key, "")
        if value and not fonts.get(w(key)):
            fonts.set(w(key), value)
            fonts.attrib.pop(w(f"{key}Theme"), None)
    for source_key, target_tag in (("size", "sz"), ("sizeCs", "szCs")):
        value = effective.get(source_key, "")
        if value and rpr.find(f"./w:{target_tag}", NS) is None:
            node = ET.SubElement(rpr, w(target_tag))
            node.set(w("val"), value)


def first_nonempty_entry_template(
    template_docx: Path,
    *,
    expected_size_half_points: str | None = None,
) -> tuple[ET.Element | None, ET.Element, ET.Element]:
    root = load_document_xml(template_docx)
    font_policy = extract_reference_font_policy(template_docx)
    policy_size_half_points = expected_size_half_points or font_policy.get("size") or font_policy.get("sizeCs")
    style_tables = load_style_tables(template_docx)
    last_error = ""
    for _, paragraph in bibliography_entry_paragraphs(root):
        ppr = paragraph.find("./w:pPr", NS)
        latin_rpr: ET.Element | None = None
        cjk_rpr: ET.Element | None = None
        fallback_rpr: ET.Element | None = None
        latin_score = -1
        cjk_score = -1
        p_style = paragraph_style_id(paragraph)
        for run in paragraph.findall("./w:r", NS):
            run_text = paragraph_text(run)
            if not run_text.strip():
                continue
            effective_signature, _effective_owners = effective_run_signature(run, p_style, style_tables)
            rpr = copy.deepcopy(run.find("./w:rPr", NS)) if run.find("./w:rPr", NS) is not None else ET.Element(w("rPr"))
            materialize_effective_run_signature(rpr, effective_signature)
            if fallback_rpr is None:
                fallback_rpr = rpr
            run_class = script_class(run_text)
            if run_class in {"latin", "mixed"} and latin_font_score(rpr) > latin_score:
                latin_rpr = rpr
                latin_score = latin_font_score(rpr)
            if run_class in {"cjk", "mixed"} and cjk_font_score(rpr) > cjk_score:
                cjk_rpr = rpr
                cjk_score = cjk_font_score(rpr)
            if latin_score >= 3 and cjk_score >= 1:
                break
        if fallback_rpr is None:
            continue
        ppr_model = copy.deepcopy(ppr) if ppr is not None else ET.Element(w("pPr"))
        latin_model = copy.deepcopy(latin_rpr if latin_rpr is not None else fallback_rpr)
        cjk_model = copy.deepcopy(cjk_rpr if cjk_rpr is not None else fallback_rpr)
        normalize_visible_size_pair(ppr_model.find(".//w:rPr", NS))
        normalize_visible_size_pair(latin_model)
        normalize_visible_size_pair(cjk_model)
        force_size_pair(ppr_model.find(".//w:rPr", NS), policy_size_half_points)
        force_size_pair(latin_model, policy_size_half_points)
        force_size_pair(cjk_model, policy_size_half_points)
        ensure_direct_font_policy(ppr_model.find(".//w:rPr", NS), font_policy)
        ensure_direct_font_policy(latin_model, font_policy)
        ensure_direct_font_policy(cjk_model, font_policy)
        if not has_required_reference_fonts(latin_model, cjk_model):
            last_error = f"bibliography donor lacks locked direct font families and no usable instruction policy was found in template: {template_docx}"
            continue
        return (ppr_model, latin_model, cjk_model)
    if font_policy.get("eastAsia") and font_policy.get("latin"):
        latin_model = ET.Element(w("rPr"))
        cjk_model = ET.Element(w("rPr"))
        force_size_pair(latin_model, policy_size_half_points)
        force_size_pair(cjk_model, policy_size_half_points)
        ensure_direct_font_policy(latin_model, font_policy)
        ensure_direct_font_policy(cjk_model, font_policy)
        if not has_required_reference_fonts(latin_model, cjk_model):
            raise RuntimeError(
                f"bibliography instruction policy lacks usable direct font families: {template_docx}"
            )
        return (None, latin_model, cjk_model)
    raise RuntimeError(last_error or f"no usable bibliography entry run donor found in template: {template_docx}")


def normalize_visible_size_pair(rpr: ET.Element | None) -> None:
    if rpr is None:
        return
    size = rpr.find("./w:sz", NS)
    size_cs = rpr.find("./w:szCs", NS)
    if size is None and size_cs is not None and size_cs.get(w("val")):
        size = ET.Element(w("sz"))
        size.set(w("val"), size_cs.get(w("val")) or "")
        rpr.insert(list(rpr).index(size_cs), size)
    elif size_cs is None and size is not None and size.get(w("val")):
        size_cs = ET.SubElement(rpr, w("szCs"))
        size_cs.set(w("val"), size.get(w("val")) or "")


def force_size_pair(rpr: ET.Element | None, size_half_points: str | None) -> None:
    if rpr is None or not size_half_points:
        return
    size = rpr.find("./w:sz", NS)
    size_cs = rpr.find("./w:szCs", NS)
    if size is None:
        size = ET.SubElement(rpr, w("sz"))
    if size_cs is None:
        size_cs = ET.SubElement(rpr, w("szCs"))
    size.set(w("val"), size_half_points)
    size_cs.set(w("val"), size_half_points)


def character_class(char: str) -> str:
    if contains_cjk(char):
        return "cjk"
    if char.isascii() and (char.isalnum() or char in WESTERN_REFERENCE_CHARS or char.isspace()):
        return "latin"
    if ord(char) < 128:
        return "latin"
    return "neutral"


def next_non_neutral(chars: list[str], start: int) -> str | None:
    for char in chars[start:]:
        cls = character_class(char)
        if cls != "neutral":
            return cls
    return None


def split_entry_text(text: str, *, strip_visible_number: bool = False) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    number_match = re.match(r"^\[[0-9]+\]", text)
    offset = 0
    if number_match and not strip_visible_number:
        segments.append(("latin", number_match.group(0)))
        offset = number_match.end()
    elif number_match and strip_visible_number:
        offset = number_match.end()
        if offset < len(text) and text[offset].isspace():
            offset += 1

    chars = list(text[offset:])
    current_class: str | None = None
    current_chars: list[str] = []
    for pos, char in enumerate(chars):
        cls = character_class(char)
        if cls == "neutral":
            cls = current_class or next_non_neutral(chars, pos + 1) or "latin"
        if current_class is None:
            current_class = cls
        if cls != current_class:
            if current_chars:
                segments.append((current_class, "".join(current_chars)))
            current_class = cls
            current_chars = [char]
        else:
            current_chars.append(char)
    if current_chars and current_class:
        segments.append((current_class, "".join(current_chars)))
    return [(cls, payload) for cls, payload in segments if payload]


def make_text_run(text: str, rpr_model: ET.Element) -> ET.Element:
    run = ET.Element(w("r"))
    run.append(copy.deepcopy(rpr_model))
    text_node = ET.SubElement(run, w("t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    return run


def direct_bookmarks(paragraph: ET.Element) -> tuple[list[ET.Element], list[ET.Element]]:
    starts: list[ET.Element] = []
    ends: list[ET.Element] = []
    for child in list(paragraph):
        if child.tag == w("bookmarkStart"):
            starts.append(copy.deepcopy(child))
        elif child.tag == w("bookmarkEnd"):
            ends.append(copy.deepcopy(child))
    return starts, ends


def replace_entry_runs(
    paragraph: ET.Element,
    *,
    ppr_model: ET.Element | None,
    latin_rpr: ET.Element,
    cjk_rpr: ET.Element,
    strip_visible_number: bool = False,
    manual_visible_number: int | None = None,
) -> int:
    original_text = paragraph_text(paragraph)
    text_for_runs = original_text
    if manual_visible_number is not None and not re.match(r"^\s*\[[0-9]+\]", original_text):
        text_for_runs = f"[{manual_visible_number}] {original_text}"
    expected_text = re.sub(r"^\[[0-9]+\]\s*", "", text_for_runs, count=1) if strip_visible_number else text_for_runs
    starts, ends = direct_bookmarks(paragraph)
    existing_ppr = paragraph.find("./w:pPr", NS)
    paragraph_ppr = copy.deepcopy(ppr_model if ppr_model is not None else existing_ppr)
    if paragraph_ppr is None:
        paragraph_ppr = ET.Element(w("pPr"))
    segments = split_entry_text(text_for_runs, strip_visible_number=strip_visible_number)
    if not segments and text_for_runs and not strip_visible_number:
        raise RuntimeError(f"entry could not be split into multiple bibliography runs: {text_for_runs[:80]}")

    runs = [
        make_text_run(payload, cjk_rpr if cls == "cjk" else latin_rpr)
        for cls, payload in segments
    ]
    paragraph[:] = []
    paragraph.append(paragraph_ppr)
    paragraph.extend(starts)
    if runs:
        paragraph.append(runs[0])
        paragraph.extend(ends)
        paragraph.extend(runs[1:])

    updated_text = paragraph_text(paragraph)
    if updated_text != expected_text:
        raise RuntimeError(f"entry text changed during bibliography format repair: {original_text!r} -> {updated_text!r}")
    return len(runs)


def zip_content_map(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def changed_zip_parts(before: Path, after: Path) -> list[str]:
    before_map = zip_content_map(before)
    after_map = zip_content_map(after)
    names = sorted(set(before_map) | set(after_map))
    return [name for name in names if before_map.get(name) != after_map.get(name)]


def write_docx_with_document_xml(source_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    write_docx_with_parts(source_docx, output_docx, {"word/document.xml": document_xml})


def write_docx_with_parts(source_docx: Path, output_docx: Path, replacements: dict[str, bytes]) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as zin, zipfile.ZipFile(output_docx, "w") as zout:
        seen: set[str] = set()
        for item in zin.infolist():
            payload = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, payload)
            seen.add(item.filename)
        for name, payload in replacements.items():
            if name not in seen:
                zout.writestr(name, payload)


STYLE_REFERENCE_TAGS = {"pStyle", "rStyle", "tblStyle"}
STYLE_DEPENDENCY_TAGS = {"basedOn", "next", "link", "numStyleLink", "styleLink"}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def style_ids_in_document(root: ET.Element) -> set[str]:
    style_ids: set[str] = set()
    for node in root.iter():
        if local_name(node.tag) in STYLE_REFERENCE_TAGS:
            value = node.get(w("val"))
            if value:
                style_ids.add(value)
    return style_ids


def style_dependencies(style: ET.Element) -> set[str]:
    deps: set[str] = set()
    for node in style.iter():
        if local_name(node.tag) in STYLE_DEPENDENCY_TAGS:
            value = node.get(w("val"))
            if value:
                deps.add(value)
    return deps


def copy_missing_referenced_styles(
    source_docx: Path,
    template_docx: Path,
    document_root: ET.Element,
) -> tuple[bytes | None, dict[str, object]]:
    target_styles = load_xml_part(source_docx, "word/styles.xml")
    template_styles = load_xml_part(template_docx, "word/styles.xml")
    referenced = set(style_ids_in_document(document_root))
    if target_styles is None or template_styles is None:
        return None, {
            "status": "skipped",
            "reason": "source or template styles.xml missing",
            "referenced_style_ids": sorted(referenced),
            "copied_style_ids": [],
            "missing_in_template": [],
        }
    target_by_id = {
        style.get(w("styleId")): style
        for style in target_styles.findall("./w:style", NS)
        if style.get(w("styleId"))
    }
    template_by_id = {
        style.get(w("styleId")): style
        for style in template_styles.findall("./w:style", NS)
        if style.get(w("styleId"))
    }
    copied: list[str] = []
    missing_in_template: list[str] = []
    queued = list(sorted(referenced))
    seen: set[str] = set()
    while queued:
        style_id = queued.pop(0)
        if style_id in seen:
            continue
        seen.add(style_id)
        if style_id in target_by_id:
            continue
        template_style = template_by_id.get(style_id)
        if template_style is None:
            missing_in_template.append(style_id)
            continue
        cloned = copy.deepcopy(template_style)
        target_styles.append(cloned)
        target_by_id[style_id] = cloned
        copied.append(style_id)
        for dep in sorted(style_dependencies(cloned)):
            if dep not in target_by_id and dep not in seen:
                queued.append(dep)
    if not copied:
        return None, {
            "status": "unchanged",
            "referenced_style_ids": sorted(referenced),
            "copied_style_ids": [],
            "missing_in_template": sorted(set(missing_in_template)),
        }
    return (
        ET.tostring(target_styles, encoding="utf-8", xml_declaration=True, short_empty_elements=True),
        {
            "status": "changed",
            "referenced_style_ids": sorted(referenced),
            "copied_style_ids": copied,
            "missing_in_template": sorted(set(missing_in_template)),
            "source": str(template_docx),
        },
    )


def repair_bibliography(
    source_docx: Path,
    template_docx: Path,
    output_docx: Path,
    *,
    expected_size_half_points: str | None = None,
    expected_size_name: str | None = None,
    use_template_numbering: bool = False,
    add_visible_numbers: bool = False,
    rendered_number_left_twips: str | None = None,
    rendered_number_hanging_twips: str | None = None,
) -> dict[str, object]:
    if source_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy, not the source path")
    ppr_model, latin_rpr, cjk_rpr = first_nonempty_entry_template(
        template_docx,
        expected_size_half_points=expected_size_half_points,
    )
    remove_direct_emphasis(ppr_model.find(".//w:rPr", NS) if ppr_model is not None else None)
    remove_direct_emphasis(latin_rpr)
    remove_direct_emphasis(cjk_rpr)
    force_cjk_slots_to_east_asia(cjk_rpr)
    numbering_report: dict[str, object] = {"mode": "manual-visible-numbering"}
    numbering_xml: bytes | None = None
    rendered_outdent_report: dict[str, object] = {"applied": False}
    if use_template_numbering:
        numbered_ppr, template_num_id = first_template_numbered_bibliography_ppr(template_docx)
        ppr_model = numbered_ppr
        remove_direct_emphasis(ppr_model.find(".//w:rPr", NS))
        numbering_xml, final_num_id, numbering_report = merge_template_bibliography_numbering(
            source_docx,
            template_docx,
            template_num_id,
        )
        set_ppr_numbering(ppr_model, final_num_id)
        if rendered_number_left_twips is not None:
            hanging_twips = str(rendered_number_hanging_twips) if rendered_number_hanging_twips is not None else str(
                ppr_model.find("./w:ind", NS).get(w("hanging"), "425")
                if ppr_model.find("./w:ind", NS) is not None
                else "425"
            )
            numbering_root = ET.fromstring(numbering_xml)
            numbering_level_changed = set_numbering_level_outdent(
                numbering_root,
                final_num_id,
                str(rendered_number_left_twips),
                hanging_twips,
            )
            numbering_xml = ET.tostring(numbering_root, encoding="utf-8", xml_declaration=True)
            rendered_outdent_report = {
                "applied": True,
                "numId": final_num_id,
                "left_twips": str(rendered_number_left_twips),
                "hanging_twips": hanging_twips,
                "paragraph_ppr_preserved": True,
                "numbering_level_changed": numbering_level_changed,
            }
    root = load_document_xml(source_docx)
    title_report = repair_references_title(root, template_docx)
    entries = bibliography_entry_paragraphs(root)
    if not entries:
        raise RuntimeError(f"no bibliography entries found in source DOCX: {source_docx}")

    touched: list[dict[str, object]] = []
    visible_number_removed = 0
    visible_number_added = 0
    for entry_ordinal, (paragraph_index, paragraph) in enumerate(entries, start=1):
        before = paragraph_text(paragraph)
        manual_visible_number = entry_ordinal if add_visible_numbers and not use_template_numbering else None
        run_count = replace_entry_runs(
            paragraph,
            ppr_model=ppr_model,
            latin_rpr=latin_rpr,
            cjk_rpr=cjk_rpr,
            strip_visible_number=use_template_numbering,
            manual_visible_number=manual_visible_number,
        )
        if use_template_numbering and re.match(r"^\[[0-9]+\]", before):
            visible_number_removed += 1
        if manual_visible_number is not None and not re.match(r"^\s*\[[0-9]+\]", before):
            visible_number_added += 1
        touched.append(
            {
                "body_child_index": paragraph_index,
                "visible_prefix": before[:100],
                "run_count_after": run_count,
                "bookmark_names": [
                    start.get(w("name"), "")
                    for start in paragraph.findall("./w:bookmarkStart", NS)
                    if start.get(w("name"), "")
                ],
            }
        )

    document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)
    replacements = {"word/document.xml": document_xml}
    styles_xml, missing_style_report = copy_missing_referenced_styles(source_docx, template_docx, root)
    if styles_xml is not None:
        replacements["word/styles.xml"] = styles_xml
    if numbering_xml is not None:
        replacements["word/numbering.xml"] = numbering_xml
    write_docx_with_parts(source_docx, output_docx, replacements)
    changed_parts = changed_zip_parts(source_docx, output_docx)
    expected_parts = ["word/document.xml"]
    if styles_xml is not None:
        expected_parts.append("word/styles.xml")
    if numbering_xml is not None:
        expected_parts.append("word/numbering.xml")
    if changed_parts != sorted(expected_parts):
        raise RuntimeError(f"unexpected DOCX package drift: {changed_parts}")

    return {
        "schema": "graduation-project-builder.bibliography-entry-format-repair.v1",
        "source_docx": str(source_docx),
        "source_sha256": sha256_file(source_docx),
        "template_docx": str(template_docx),
        "template_sha256": sha256_file(template_docx),
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "changed_zip_parts": changed_parts,
        "entry_count": len(entries),
        "template_numbering_replay": use_template_numbering,
        "template_numbering_model": numbering_report,
        "rendered_number_outdent": rendered_outdent_report,
        "visible_manual_numbers_removed": visible_number_removed,
        "visible_manual_numbers_added": visible_number_added,
        "references_title_repair": title_report,
        "missing_referenced_style_repair": missing_style_report,
        "expected_size_name": expected_size_name or font_size_name_for_half_points(expected_size_half_points) or "template-derived",
        "expected_size_half_points": expected_size_half_points or "template-derived",
        "touched_scope": "word/document.xml bibliography entry paragraphs only",
        "template_replay_scope": [
            "reference entry paragraph properties",
            "reference entry leading bracket number run separation or template automatic-numbering marker",
            "reference entry latin run rPr",
            "reference entry cjk run rPr",
            "existing citation bookmark start/end anchors",
        ],
        "entries": touched,
        "verdict": "pass",
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Replay template bibliography entry run formatting onto a DOCX review copy.")
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--bibliography-size-name", help="Exact Chinese size name for bibliography entries, for example 五号")
    parser.add_argument("--bibliography-size-half-points", help="Exact bibliography entry size in DOCX half-points, for example 21 for 五号")
    parser.add_argument(
        "--use-template-numbering",
        action="store_true",
        help="Copy the template bibliography numbering model and remove visible [n] prefixes.",
    )
    parser.add_argument(
        "--add-visible-numbers",
        action="store_true",
        help="Add manual [n] prefixes when bibliography entries lost visible numbering and no template automatic-numbering model is available.",
    )
    parser.add_argument(
        "--rendered-number-left-twips",
        help="When using template numbering, force only the numbering-level left indent in twips to match rendered template geometry while preserving entry paragraph pPr.",
    )
    parser.add_argument(
        "--rendered-number-hanging-twips",
        help="When using template numbering, force only the numbering-level hanging indent in twips to match rendered template geometry while preserving entry paragraph pPr.",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    source_docx = Path(args.source_docx).resolve()
    template_docx = Path(args.template_docx).resolve()
    output_docx = Path(args.output_docx).resolve()
    report_path = Path(args.report).resolve()
    try:
        expected_size_half_points = resolve_font_size_half_points(
            args.bibliography_size_name,
            args.bibliography_size_half_points,
        )
    except ValueError as exc:
        parser.error(str(exc))
    template_policy = extract_reference_font_policy(template_docx)
    if expected_size_half_points is None and template_policy:
        expected_size_half_points = template_policy.get("size") or template_policy.get("sizeCs")

    report = repair_bibliography(
        source_docx,
        template_docx,
        output_docx,
        expected_size_half_points=expected_size_half_points,
        expected_size_name=args.bibliography_size_name or font_size_name_for_half_points(expected_size_half_points),
        use_template_numbering=args.use_template_numbering,
        add_visible_numbers=args.add_visible_numbers,
        rendered_number_left_twips=args.rendered_number_left_twips,
        rendered_number_hanging_twips=args.rendered_number_hanging_twips,
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": str(output_docx), "entry_count": report["entry_count"], "verdict": "pass"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
