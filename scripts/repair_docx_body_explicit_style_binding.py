#!/usr/bin/env python3
"""Bind audited thesis body paragraphs to an explicit paragraph style.

This helper is deliberately narrow: it reuses ``audit_docx_body_style`` body
paragraph discovery, then rewrites only ``word/document.xml`` so body prose that
currently falls back to the default style receives an explicit ``w:pStyle``.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from audit_docx_body_style import (
    NS,
    W,
    body_paragraphs,
    choose_reference_body_style,
    declared_body_size_half_points,
    find_matching_final_style_id,
    paragraph_style_id,
    resolved_style_metrics,
    sha256_file,
    style_node_by_type,
    style_matches_target,
    style_name_by_id,
    normalize_space,
)
from validate_skill_gate_record_evidence import _body_text_paragraphs


W_NS = NS["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is not None:
        return ppr
    ppr = ET.Element(W + "pPr")
    paragraph.insert(0, ppr)
    return ppr


PPR_CHILD_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]
PPR_CHILD_ORDER_INDEX = {W + name: index for index, name in enumerate(PPR_CHILD_ORDER)}


def reorder_ppr_children(ppr: ET.Element) -> bool:
    """Keep w:pPr children in schema order after narrow XML edits."""

    children = list(ppr)
    if len(children) < 2:
        return False
    indexed_children = list(enumerate(children))
    sorted_children = sorted(
        indexed_children,
        key=lambda item: (
            PPR_CHILD_ORDER_INDEX.get(item[1].tag, len(PPR_CHILD_ORDER_INDEX)),
            item[0],
        ),
    )
    reordered = [child for _index, child in sorted_children]
    if reordered == children:
        return False
    ppr[:] = reordered
    return True


def ensure_pstyle(ppr: ET.Element, style_id: str) -> tuple[bool, str | None]:
    pstyle = ppr.find("./w:pStyle", NS)
    old_value = pstyle.attrib.get(W + "val") if pstyle is not None else None
    if pstyle is None:
        pstyle = ET.Element(W + "pStyle")
        ppr.insert(0, pstyle)
    elif list(ppr).index(pstyle) != 0:
        ppr.remove(pstyle)
        ppr.insert(0, pstyle)
    if old_value == style_id:
        return False, old_value
    pstyle.set(W + "val", style_id)
    return True, old_value


def ensure_spacing_and_indent(
    ppr: ET.Element,
    *,
    before: str = "0",
    after: str = "0",
    line: str = "360",
    line_rule: str = "auto",
    first_line: str = "480",
    first_line_chars: str = "200",
) -> bool:
    changed = False
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(W + "spacing")
        ppr.append(spacing)
        changed = True
    for attr, value in (
        (W + "before", before),
        (W + "after", after),
        (W + "line", line),
        (W + "lineRule", line_rule),
    ):
        if spacing.attrib.get(attr) != value:
            spacing.set(attr, value)
            changed = True
    for stale_attr in (W + "beforeLines", W + "afterLines"):
        if stale_attr in spacing.attrib:
            spacing.attrib.pop(stale_attr, None)
            changed = True

    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(W + "ind")
        ppr.append(ind)
        changed = True
    for attr in (W + "leftChars", W + "rightChars", W + "hanging", W + "hangingChars"):
        if attr in ind.attrib:
            ind.attrib.pop(attr, None)
            changed = True
    for attr, value in ((W + "firstLine", first_line), (W + "firstLineChars", first_line_chars)):
        if ind.attrib.get(attr) != value:
            ind.set(attr, value)
            changed = True
    return changed


def ensure_spacing_and_indent_from_reference(ppr: ET.Element, reference_metrics: dict[str, str]) -> bool:
    """Add explicit body prose metrics while preserving the template line rule."""

    return ensure_spacing_and_indent(
        ppr,
        before=reference_metrics.get("before") or "0",
        after=reference_metrics.get("after") or "0",
        line=reference_metrics.get("line") or "360",
        line_rule=reference_metrics.get("lineRule") or "auto",
        first_line=reference_metrics.get("firstLine") or "480",
        first_line_chars=reference_metrics.get("firstLineChars") or "200",
    )


def ensure_line_rule_from_reference(ppr: ET.Element, reference_metrics: dict[str, str]) -> bool:
    """Repair only the direct lineRule override that masks the body style baseline."""

    expected = reference_metrics.get("lineRule") or ""
    if not expected:
        return False
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(W + "spacing")
        ppr.append(spacing)
        changed = True
    else:
        changed = False
    reference_line = reference_metrics.get("line") or ""
    if reference_line and spacing.attrib.get(W + "line") != reference_line:
        spacing.set(W + "line", reference_line)
        changed = True
    if spacing.attrib.get(W + "lineRule") == expected:
        return changed
    spacing.set(W + "lineRule", expected)
    return True


def _set_or_remove_attr(node: ET.Element, attr: str, value: str) -> bool:
    current = node.attrib.get(attr, "")
    if value:
        if current == value:
            return False
        node.set(attr, value)
        return True
    if attr in node.attrib:
        node.attrib.pop(attr, None)
        return True
    return False


def ensure_full_body_metrics_from_reference(ppr: ET.Element, reference_metrics: dict[str, str]) -> bool:
    """Make direct paragraph metrics match the locked template body instance."""

    changed = False
    jc_value = reference_metrics.get("jc", "")
    jc = ppr.find("./w:jc", NS)
    if jc_value:
        if jc is None:
            jc = ET.Element(W + "jc")
            ppr.append(jc)
            changed = True
        changed = _set_or_remove_attr(jc, W + "val", jc_value) or changed
    elif jc is not None:
        ppr.remove(jc)
        changed = True

    spacing_keys = ("before", "after", "line", "lineRule")
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None and any(reference_metrics.get(key, "") for key in spacing_keys):
        spacing = ET.Element(W + "spacing")
        ppr.append(spacing)
        changed = True
    if spacing is not None:
        for key in spacing_keys:
            changed = _set_or_remove_attr(spacing, W + key, reference_metrics.get(key, "")) or changed
        if not spacing.attrib:
            ppr.remove(spacing)
            changed = True

    ind_keys = ("firstLine", "left", "right", "firstLineChars", "leftChars", "rightChars", "hanging", "hangingChars")
    ind = ppr.find("./w:ind", NS)
    if ind is None and any(reference_metrics.get(key, "") for key in ind_keys):
        ind = ET.Element(W + "ind")
        ppr.append(ind)
        changed = True
    if ind is not None:
        for key in ind_keys:
            changed = _set_or_remove_attr(ind, W + key, reference_metrics.get(key, "")) or changed
        if not ind.attrib:
            ppr.remove(ind)
            changed = True
    return changed


ASCII_ALNUM_REPAIR_RE = re.compile(r"[A-Za-z0-9]")
CJK_REPAIR_RE = re.compile(r"[\u4e00-\u9fff]")
ASCII_ALPHA_REPAIR_RE = re.compile(r"[A-Za-z]")
BODY_CJK_FONT = "\u5b8b\u4f53"
BODY_WESTERN_FONT = "Times New Roman"
FONT_THEME_ATTRS = ("eastAsiaTheme", "asciiTheme", "hAnsiTheme", "csTheme", "cstheme")


def _ensure_rpr(parent: ET.Element) -> ET.Element:
    rpr = parent.find("w:rPr", NS)
    if rpr is not None:
        return rpr
    rpr = ET.Element(W + "rPr")
    parent.insert(0, rpr)
    return rpr


def _ensure_rfonts(rpr: ET.Element) -> ET.Element:
    rfonts = rpr.find("w:rFonts", NS)
    if rfonts is not None:
        return rfonts
    rfonts = ET.Element(W + "rFonts")
    rpr.insert(0, rfonts)
    return rfonts


def _set_child_val(parent: ET.Element, tag: str, value: str) -> bool:
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = ET.Element(W + tag)
        parent.append(child)
        changed = True
    else:
        changed = False
    if child.attrib.get(W + "val", "") != value:
        child.set(W + "val", value)
        changed = True
    return changed


def _remove_font_theme_attrs(rfonts: ET.Element) -> bool:
    changed = False
    for attr in FONT_THEME_ATTRS:
        key = W + attr
        if key in rfonts.attrib:
            rfonts.attrib.pop(key, None)
            changed = True
    return changed


def _set_font_attr(rfonts: ET.Element, attr: str, value: str) -> bool:
    key = W + attr
    if rfonts.attrib.get(key, "") == value:
        return False
    rfonts.set(key, value)
    return True


def materialize_body_style_baseline(style: ET.Element, reference_metrics: dict[str, str]) -> bool:
    """Make the target body style match the locked template baseline.

    If the template body style leaves font slots unset, keep them unset in the
    final body style instead of materializing fallback fonts. This prevents a
    name-equivalent style such as reference ``styleId=1/name=Normal`` from
    being repaired into a visually acceptable but audit-drifting ``Normal``
    style with extra rFonts.
    """

    changed = False
    based_on = style.find("w:basedOn", NS)
    if based_on is not None:
        style.remove(based_on)
        changed = True
    ppr = style.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        style.append(ppr)
        changed = True
    if reference_metrics.get("jc"):
        jc = ppr.find("w:jc", NS)
        if jc is None:
            jc = ET.Element(W + "jc")
            ppr.append(jc)
            changed = True
        changed = _set_or_remove_attr(jc, W + "val", reference_metrics.get("jc", "")) or changed
    spacing_values = {key: reference_metrics.get(key, "") for key in ("before", "after", "line", "lineRule")}
    if any(spacing_values.values()):
        spacing = ppr.find("w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(W + "spacing")
            ppr.append(spacing)
            changed = True
        for key, value in spacing_values.items():
            changed = _set_or_remove_attr(spacing, W + key, value) or changed
    indent_values = {
        key: reference_metrics.get(key, "")
        for key in ("firstLine", "left", "right", "firstLineChars", "leftChars", "rightChars", "hanging", "hangingChars")
    }
    if any(indent_values.values()):
        ind = ppr.find("w:ind", NS)
        if ind is None:
            ind = ET.Element(W + "ind")
            ppr.append(ind)
            changed = True
        for key, value in indent_values.items():
            changed = _set_or_remove_attr(ind, W + key, value) or changed
    changed = reorder_ppr_children(ppr) or changed
    rpr = _ensure_rpr(style)
    font_values = {
        "eastAsia": reference_metrics.get("eastAsia", ""),
        "ascii": reference_metrics.get("ascii", ""),
        "hAnsi": reference_metrics.get("hAnsi", ""),
    }
    rfonts = rpr.find("w:rFonts", NS)
    if any(font_values.values()):
        rfonts = _ensure_rfonts(rpr)
        changed = _remove_font_theme_attrs(rfonts) or changed
        for attr, value in font_values.items():
            changed = _set_or_remove_attr(rfonts, W + attr, value) or changed
        cs_value = reference_metrics.get("ascii") or reference_metrics.get("hAnsi") or ""
        changed = _set_or_remove_attr(rfonts, W + "cs", cs_value) or changed
        if not rfonts.attrib:
            rpr.remove(rfonts)
            changed = True
    elif rfonts is not None:
        rpr.remove(rfonts)
        changed = True
    size = reference_metrics.get("size") or "24"
    changed = _set_child_val(rpr, "sz", size) or changed
    return changed


def split_text_by_script(text: str) -> list[tuple[str, str]]:
    if not (CJK_REPAIR_RE.search(text) and ASCII_ALNUM_REPAIR_RE.search(text)):
        return [("mixed", text)]
    parts: list[tuple[str, str]] = []
    current_kind = ""
    buffer: list[str] = []
    for char in text:
        if CJK_REPAIR_RE.match(char):
            kind = "cjk"
        elif char.isascii() and (
            char.isalnum() or char in " ./\\-+_:%()[]<>="
        ):
            kind = "latin"
        else:
            kind = current_kind or "cjk"
        if buffer and kind != current_kind:
            parts.append((current_kind, "".join(buffer)))
            buffer = []
        current_kind = kind
        buffer.append(char)
    if buffer:
        parts.append((current_kind, "".join(buffer)))
    return parts


def _run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", NS))


def _is_simple_visible_text_run(run: ET.Element) -> bool:
    if run.find(".//w:txbxContent", NS) is not None:
        return False
    for child in list(run):
        if child.tag not in {W + "rPr", W + "t"}:
            return False
    return bool(_run_text(run))


def _set_text_preserve_space(text_node: ET.Element, text: str) -> None:
    text_node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(XML_SPACE, "preserve")


def _apply_body_run_font(rpr: ET.Element, kind: str) -> bool:
    changed = False
    rfonts = _ensure_rfonts(rpr)
    changed = _remove_font_theme_attrs(rfonts) or changed
    if kind == "latin":
        for attr in ("ascii", "hAnsi", "cs"):
            changed = _set_font_attr(rfonts, attr, BODY_WESTERN_FONT) or changed
    elif kind == "cjk":
        changed = _set_font_attr(rfonts, "eastAsia", BODY_CJK_FONT) or changed
        if not (rfonts.attrib.get(W + "ascii") or rfonts.attrib.get(W + "hAnsi")):
            changed = _set_font_attr(rfonts, "ascii", BODY_WESTERN_FONT) or changed
            changed = _set_font_attr(rfonts, "hAnsi", BODY_WESTERN_FONT) or changed
            changed = _set_font_attr(rfonts, "cs", BODY_WESTERN_FONT) or changed
    return changed


def split_mixed_script_runs(paragraph: ET.Element) -> bool:
    changed = False
    children = list(paragraph)
    index = 0
    while index < len(children):
        run = children[index]
        if run.tag != W + "r" or not _is_simple_visible_text_run(run):
            index += 1
            continue
        text = _run_text(run)
        if not (CJK_REPAIR_RE.search(text) and ASCII_ALNUM_REPAIR_RE.search(text)):
            index += 1
            continue
        parts = split_text_by_script(text)
        if len(parts) <= 1:
            index += 1
            continue
        old_rpr = run.find("w:rPr", NS)
        new_runs: list[ET.Element] = []
        for kind, value in parts:
            if not value:
                continue
            new_run = ET.Element(W + "r")
            if old_rpr is not None:
                new_run.append(copy.deepcopy(old_rpr))
            rpr = _ensure_rpr(new_run)
            _apply_body_run_font(rpr, "latin" if kind == "latin" else "cjk")
            text_node = ET.Element(W + "t")
            _set_text_preserve_space(text_node, value)
            new_run.append(text_node)
            new_runs.append(new_run)
        if not new_runs:
            index += 1
            continue
        paragraph.remove(run)
        for offset, new_run in enumerate(new_runs):
            paragraph.insert(index + offset, new_run)
        children = list(paragraph)
        index += len(new_runs)
        changed = True
    return changed


def repair_body_run_metrics(paragraph: ET.Element, western_font: str = "Times New Roman") -> bool:
    """Remove heading-size direct overrides and add Western font slots on ASCII runs."""

    changed = split_mixed_script_runs(paragraph)
    for run in paragraph.findall("w:r", NS):
        if run.find(".//w:txbxContent", NS) is not None:
            continue
        text = "".join(node.text or "" for node in run.findall(".//w:t", NS))
        if not text:
            continue
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            if not ASCII_ALNUM_REPAIR_RE.search(text):
                continue
            rpr = ET.Element(W + "rPr")
            run.insert(0, rpr)
            changed = True
        for tag in ("sz", "szCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is not None:
                rpr.remove(node)
                changed = True
        if ASCII_ALNUM_REPAIR_RE.search(text):
            rfonts = rpr.find("w:rFonts", NS)
            if rfonts is None:
                rfonts = ET.Element(W + "rFonts")
                rpr.insert(0, rfonts)
                changed = True
            if _remove_font_theme_attrs(rfonts):
                changed = True
            if not (
                rfonts.attrib.get(W + "ascii")
                or rfonts.attrib.get(W + "hAnsi")
                or rfonts.attrib.get(W + "asciiTheme")
                or rfonts.attrib.get(W + "hAnsiTheme")
                or rfonts.attrib.get(W + "cs")
                or rfonts.attrib.get(W + "csTheme")
            ):
                rfonts.set(W + "ascii", western_font)
                rfonts.set(W + "hAnsi", western_font)
                changed = True
            if ASCII_ALPHA_REPAIR_RE.search(text) and not rfonts.attrib.get(W + "cs"):
                rfonts.set(W + "cs", western_font)
                changed = True
        if CJK_REPAIR_RE.search(text):
            rfonts = rpr.find("w:rFonts", NS)
            if rfonts is None:
                rfonts = ET.Element(W + "rFonts")
                rpr.insert(0, rfonts)
                changed = True
            if _remove_font_theme_attrs(rfonts):
                changed = True
            if not rfonts.attrib.get(W + "eastAsia"):
                rfonts.set(W + "eastAsia", BODY_CJK_FONT)
                changed = True
    return changed


def remove_body_keep_next(ppr: ET.Element) -> bool:
    """Remove table/caption keep-with-next leakage from audited body prose."""

    changed = False
    for node in list(ppr.findall("w:keepNext", NS)):
        ppr.remove(node)
        changed = True
    return changed


def _style_by_id(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for style in styles_root.findall("./w:style", NS):
        if style.attrib.get(W + "styleId") == style_id:
            return style
    return None


def _style_name(style: ET.Element | None) -> str:
    if style is None:
        return ""
    name = style.find("./w:name", NS)
    return name.attrib.get(W + "val", "") if name is not None else ""


def _style_id_by_name(styles_root: ET.Element, style_name: str) -> str | None:
    normalized = normalize_space(style_name).lower()
    if not normalized:
        return None
    for style in styles_root.findall("./w:style", NS):
        if style.attrib.get(W + "type") != "paragraph":
            continue
        if normalize_space(_style_name(style)).lower() == normalized:
            return style.attrib.get(W + "styleId")
    return None


def _style_based_on(style: ET.Element | None) -> str | None:
    if style is None:
        return None
    node = style.find("./w:basedOn", NS)
    return node.attrib.get(W + "val") if node is not None else None


def _default_paragraph_style_id(styles_root: ET.Element) -> str | None:
    for style in styles_root.findall("./w:style", NS):
        if style.attrib.get(W + "type") == "paragraph" and style.attrib.get(W + "default") == "1":
            return style.attrib.get(W + "styleId")
    return None


def _style_ids_to_copy(reference_root: ET.Element, target_style_id: str) -> list[str]:
    """Copy the target body style plus non-default ancestors it inherits from."""

    default_style_id = _default_paragraph_style_id(reference_root)
    style_ids: list[str] = []
    seen: set[str] = set()
    current: str | None = target_style_id
    while current and current not in seen:
        seen.add(current)
        style = _style_by_id(reference_root, current)
        if style is None:
            break
        style_ids.append(current)
        parent = _style_based_on(style)
        if not parent or parent == default_style_id or parent.lower() == "normal":
            break
        current = parent
    return style_ids


def repair_styles_xml_from_reference(
    xml_bytes: bytes,
    *,
    reference_docx: Path,
    target_style_id: str,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    try:
        with zipfile.ZipFile(reference_docx, "r") as zf:
            reference_root = ET.fromstring(zf.read("word/styles.xml"))
    except KeyError:
        return xml_bytes, {
            "target_style_definition_copied": False,
            "target_style_definition_issue": "reference DOCX has no word/styles.xml",
        }
    target_style = _style_by_id(root, target_style_id)
    target_style_name = _style_name(target_style) or target_style_id
    reference_target_style_id = target_style_id
    if _style_by_id(reference_root, reference_target_style_id) is None:
        reference_target_style_id = (
            _style_id_by_name(reference_root, target_style_name)
            or (
                _style_id_by_name(reference_root, "Normal")
                if normalize_space(target_style_name).lower() == "normal"
                or target_style_id.lower() == "normal"
                else None
            )
            or target_style_id
        )
    style_ids = _style_ids_to_copy(reference_root, reference_target_style_id)
    if not style_ids:
        return xml_bytes, {
            "target_style_definition_copied": False,
            "target_style_definition_issue": (
                f"reference style not found: {target_style_id}"
                + (f" / {target_style_name}" if target_style_name != target_style_id else "")
            ),
        }
    reference_styles, _default_reference_style_id = style_node_by_type(reference_root)
    reference_metrics = resolved_style_metrics(reference_styles, reference_target_style_id)
    declared_size = declared_body_size_half_points(reference_docx)
    if declared_size is not None:
        reference_metrics["size"] = str(declared_size)
    copied_style_ids: list[str] = []
    copied_style_mappings: list[dict[str, str]] = []
    for style_id in style_ids:
        reference_style = _style_by_id(reference_root, style_id)
        if reference_style is None:
            continue
        replacement = copy.deepcopy(reference_style)
        final_style_id = target_style_id if style_id == reference_target_style_id else style_id
        if final_style_id != style_id:
            replacement.set(W + "styleId", final_style_id)
        existing = _style_by_id(root, final_style_id)
        if existing is None:
            root.append(replacement)
            copied_style_ids.append(final_style_id)
            copied_style_mappings.append({"reference_style_id": style_id, "final_style_id": final_style_id})
        elif ET.tostring(existing, encoding="utf-8") != ET.tostring(replacement, encoding="utf-8"):
            children = list(root)
            index = children.index(existing)
            root.remove(existing)
            root.insert(index, replacement)
            copied_style_ids.append(final_style_id)
            copied_style_mappings.append({"reference_style_id": style_id, "final_style_id": final_style_id})
    target_style = _style_by_id(root, target_style_id)
    body_baseline_materialized = False
    if target_style is not None:
        body_baseline_materialized = materialize_body_style_baseline(target_style, reference_metrics)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), {
        "target_style_definition_copied": bool(copied_style_ids),
        "target_style_definition_copy_scope": style_ids,
        "target_style_reference_style_id": reference_target_style_id,
        "target_style_definition_copied_style_ids": copied_style_ids,
        "target_style_definition_copied_style_mappings": copied_style_mappings,
        "target_style_body_baseline_materialized": body_baseline_materialized,
        "target_style_body_reference_metrics": reference_metrics,
        "target_style_definition_issue": "",
    }


def select_body_targets(
    reference_docx: Path | None,
    input_docx: Path,
    explicit_style_id: str | None,
) -> tuple[str, str, list[dict[str, object]], dict[str, object]]:
    final_body, final_styles, default_final_style_id = body_paragraphs(input_docx)
    if explicit_style_id:
        target_style_id = explicit_style_id
        target_style_name = explicit_style_id
    else:
        if reference_docx is None:
            raise RuntimeError("either --reference-docx or --style-id is required")
        reference_body, reference_styles, default_reference_style_id = body_paragraphs(reference_docx)
        reference_body_style_id, reference_body_style_name = choose_reference_body_style(
            reference_body,
            reference_styles,
            default_reference_style_id,
        )
        target_style_id = find_matching_final_style_id(
            final_styles,
            reference_body_style_id,
            reference_body_style_name,
            default_final_style_id,
        )
        target_style_name = reference_body_style_name or reference_body_style_id or ""
        if not target_style_id:
            raise RuntimeError("could not find matching final body style")

    normalized_target_name = normalize_space(target_style_name).lower()
    targets = []
    for paragraph in final_body:
        paragraph_style = paragraph.get("style_id")
        if not paragraph_style:
            targets.append(paragraph)
            continue
        paragraph_style_name = style_name_by_id(final_styles, str(paragraph_style))
        if str(paragraph_style) == target_style_id:
            continue
        if normalized_target_name and normalize_space(paragraph_style_name).lower() == normalized_target_name:
            continue
        targets.append(paragraph)
    return target_style_id, target_style_name, targets, {
        "audited_body_paragraph_count": len(final_body),
        "default_final_style_id": default_final_style_id,
        "target_style_id": target_style_id,
        "target_style_name": target_style_name,
    }


def reference_body_instance_metrics(reference_docx: Path) -> dict[str, str]:
    reference_body, reference_styles, default_reference_style_id = body_paragraphs(reference_docx)
    reference_body_style_id, reference_body_style_name = choose_reference_body_style(
        reference_body,
        reference_styles,
        default_reference_style_id,
    )
    reference_instance_candidates = [
        paragraph for paragraph in reference_body
        if style_matches_target(
            paragraph,
            target_style_id=reference_body_style_id,
            target_style_name=reference_body_style_name,
            default_style_id=default_reference_style_id,
        )
    ] or reference_body
    counter = Counter(
        tuple(sorted(dict(paragraph.get("effective_metrics") or paragraph["instance_metrics"]).items()))
        for paragraph in reference_instance_candidates
    )
    return dict(counter.most_common(1)[0][0]) if counter else {}


def repair_document_xml(
    xml_bytes: bytes,
    *,
    target_style_id: str,
    target_paragraph_indices: set[int],
    repair_direct_body_metrics: bool = False,
    direct_metric_paragraph_indices: set[int] | None = None,
    repair_body_line_rule_from_reference: bool = False,
    body_line_rule_paragraph_indices: set[int] | None = None,
    repair_full_body_metrics_from_reference: bool = False,
    full_metric_paragraph_indices: set[int] | None = None,
    repair_body_run_metrics_from_reference: bool = False,
    run_metric_paragraph_indices: set[int] | None = None,
    repair_body_keep_next_from_reference: bool = False,
    keep_next_paragraph_indices: set[int] | None = None,
    reference_body_metrics: dict[str, str] | None = None,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        return xml_bytes, {
            "changed_paragraph_count": 0,
            "issues": ["word/document.xml has no w:body"],
            "changed_paragraphs": [],
        }

    direct_metric_paragraph_ids: set[int] = set()
    if repair_direct_body_metrics:
        if direct_metric_paragraph_indices is None:
            direct_metric_paragraph_ids = {id(paragraph) for paragraph in _body_text_paragraphs(root)}
        else:
            direct_metric_paragraph_ids = set()

    changed_paragraphs: list[dict[str, object]] = []
    metric_changed_paragraphs: list[dict[str, object]] = []
    line_rule_changed_paragraphs: list[dict[str, object]] = []
    full_metric_changed_paragraphs: list[dict[str, object]] = []
    run_metric_changed_paragraphs: list[dict[str, object]] = []
    keep_next_changed_paragraphs: list[dict[str, object]] = []
    for body_child_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        direct_metric_target = (
            body_child_index in direct_metric_paragraph_indices
            if direct_metric_paragraph_indices is not None
            else id(child) in direct_metric_paragraph_ids
        )
        line_rule_target = (
            repair_body_line_rule_from_reference
            and body_line_rule_paragraph_indices is not None
            and body_child_index in body_line_rule_paragraph_indices
        )
        full_metric_target = (
            repair_full_body_metrics_from_reference
            and full_metric_paragraph_indices is not None
            and body_child_index in full_metric_paragraph_indices
        )
        run_metric_target = (
            repair_body_run_metrics_from_reference
            and run_metric_paragraph_indices is not None
            and body_child_index in run_metric_paragraph_indices
        )
        keep_next_target = (
            repair_body_keep_next_from_reference
            and keep_next_paragraph_indices is not None
            and body_child_index in keep_next_paragraph_indices
        )
        if (
            body_child_index not in target_paragraph_indices
            and not direct_metric_target
            and not line_rule_target
            and not full_metric_target
            and not run_metric_target
            and not keep_next_target
        ):
            continue
        old_style = paragraph_style_id(child)
        ppr = ensure_ppr(child)
        text = "".join(node.text or "" for node in child.findall(".//w:t", NS)).strip()
        if body_child_index in target_paragraph_indices:
            changed, old_value = ensure_pstyle(ppr, target_style_id)
            if changed:
                changed_paragraphs.append(
                    {
                        "paragraph_index": body_child_index,
                        "old_style_id": old_style or old_value or "",
                        "new_style_id": target_style_id,
                        "text_prefix": text[:100],
                    }
                )
        if direct_metric_target and ensure_spacing_and_indent(ppr):
            changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "old_style_id": old_style or "",
                    "new_style_id": old_style or target_style_id,
                    "text_prefix": text[:100],
                }
            )
            metric_changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "text_prefix": text[:100],
                }
            )
        if line_rule_target and ensure_line_rule_from_reference(ppr, reference_body_metrics or {}):
            changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "old_style_id": old_style or "",
                    "new_style_id": old_style or target_style_id,
                    "text_prefix": text[:100],
                }
            )
            line_rule_changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "text_prefix": text[:100],
                }
            )
        if full_metric_target and ensure_full_body_metrics_from_reference(ppr, reference_body_metrics or {}):
            changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "old_style_id": old_style or "",
                    "new_style_id": old_style or target_style_id,
                    "text_prefix": text[:100],
                }
            )
            full_metric_changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "text_prefix": text[:100],
                }
            )
        if run_metric_target and repair_body_run_metrics(child):
            changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "old_style_id": old_style or "",
                    "new_style_id": old_style or target_style_id,
                    "text_prefix": text[:100],
                }
            )
            run_metric_changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "text_prefix": text[:100],
                }
            )
        if keep_next_target and remove_body_keep_next(ppr):
            changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "old_style_id": old_style or "",
                    "new_style_id": old_style or target_style_id,
                    "text_prefix": text[:100],
                }
            )
            keep_next_changed_paragraphs.append(
                {
                    "paragraph_index": body_child_index,
                    "text_prefix": text[:100],
                }
            )
        reorder_ppr_children(ppr)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), {
        "changed_paragraph_count": len(changed_paragraphs),
        "direct_metric_target_paragraph_count": (
            len(direct_metric_paragraph_indices)
            if direct_metric_paragraph_indices is not None
            else len(direct_metric_paragraph_ids)
        ),
        "direct_metric_changed_paragraph_count": len(metric_changed_paragraphs),
        "body_line_rule_target_paragraph_count": (
            len(body_line_rule_paragraph_indices or set())
            if repair_body_line_rule_from_reference
            else 0
        ),
        "body_line_rule_changed_paragraph_count": len(line_rule_changed_paragraphs),
        "full_body_metric_target_paragraph_count": (
            len(full_metric_paragraph_indices or set())
            if repair_full_body_metrics_from_reference
            else 0
        ),
        "full_body_metric_changed_paragraph_count": len(full_metric_changed_paragraphs),
        "body_run_metric_target_paragraph_count": (
            len(run_metric_paragraph_indices or set())
            if repair_body_run_metrics_from_reference
            else 0
        ),
        "body_run_metric_changed_paragraph_count": len(run_metric_changed_paragraphs),
        "body_keep_next_target_paragraph_count": (
            len(keep_next_paragraph_indices or set())
            if repair_body_keep_next_from_reference
            else 0
        ),
        "body_keep_next_changed_paragraph_count": len(keep_next_changed_paragraphs),
        "reference_body_line_rule": (reference_body_metrics or {}).get("lineRule", ""),
        "changed_paragraphs": changed_paragraphs,
        "direct_metric_changed_paragraphs": metric_changed_paragraphs,
        "body_line_rule_changed_paragraphs": line_rule_changed_paragraphs,
        "full_body_metric_changed_paragraphs": full_metric_changed_paragraphs,
        "body_run_metric_changed_paragraphs": run_metric_changed_paragraphs,
        "body_keep_next_changed_paragraphs": keep_next_changed_paragraphs,
        "issues": [] if target_paragraph_indices else ["no body paragraphs required explicit style repair"],
    }


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    reference_docx: Path | None = None,
    style_id: str | None = None,
    repair_direct_body_metrics: bool = False,
    repair_body_line_rule_from_reference: bool = False,
    repair_full_body_metrics_from_reference: bool = False,
    repair_body_run_metrics_from_reference: bool = False,
    repair_body_keep_next_from_reference: bool = False,
    copy_target_style_from_reference: bool = False,
) -> dict[str, object]:
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()
    reference_docx = reference_docx.resolve() if reference_docx else None
    if input_docx == output_docx:
        raise RuntimeError("output DOCX must be a new review-copy path")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    target_style_id, target_style_name, targets, selection_summary = select_body_targets(
        reference_docx,
        input_docx,
        style_id,
    )
    target_indices = {int(paragraph["paragraph_index"]) for paragraph in targets}
    direct_metric_indices: set[int] | None = None
    reference_metrics: dict[str, str] = {}
    if repair_direct_body_metrics:
        if reference_docx is not None:
            reference_metrics = reference_body_instance_metrics(reference_docx)
        audited_body, _audited_styles, _audited_default_style = body_paragraphs(input_docx)
        direct_metric_indices = {int(paragraph["paragraph_index"]) for paragraph in audited_body}
    body_line_rule_indices: set[int] | None = None
    if repair_body_line_rule_from_reference:
        if reference_docx is None:
            raise RuntimeError("--repair-body-line-rule-from-reference requires --reference-docx")
        audited_body, _audited_styles, _audited_default_style = body_paragraphs(input_docx)
        body_line_rule_indices = {int(paragraph["paragraph_index"]) for paragraph in audited_body}
        reference_metrics = reference_body_instance_metrics(reference_docx)
    full_metric_indices: set[int] | None = None
    if repair_full_body_metrics_from_reference:
        if reference_docx is None:
            raise RuntimeError("--repair-full-body-metrics-from-reference requires --reference-docx")
        audited_body, _audited_styles, _audited_default_style = body_paragraphs(input_docx)
        full_metric_indices = {int(paragraph["paragraph_index"]) for paragraph in audited_body}
        reference_metrics = reference_body_instance_metrics(reference_docx)
    run_metric_indices: set[int] | None = None
    if repair_body_run_metrics_from_reference:
        audited_body, _audited_styles, _audited_default_style = body_paragraphs(input_docx)
        run_metric_indices = {int(paragraph["paragraph_index"]) for paragraph in audited_body}
    keep_next_indices: set[int] | None = None
    if repair_body_keep_next_from_reference:
        audited_body, _audited_styles, _audited_default_style = body_paragraphs(input_docx)
        keep_next_indices = {int(paragraph["paragraph_index"]) for paragraph in audited_body}

    changed_parts: list[str] = []
    repair_summary: dict[str, object] = {}
    style_repair_summary: dict[str, object] = {}
    original_document_sha256 = ""
    repaired_document_sha256 = ""
    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                original_document_sha256 = _sha256_bytes(data)
                data, repair_summary = repair_document_xml(
                    data,
                    target_style_id=target_style_id,
                    target_paragraph_indices=target_indices,
                    repair_direct_body_metrics=repair_direct_body_metrics,
                    direct_metric_paragraph_indices=direct_metric_indices,
                    repair_body_line_rule_from_reference=repair_body_line_rule_from_reference,
                    body_line_rule_paragraph_indices=body_line_rule_indices,
                    repair_full_body_metrics_from_reference=repair_full_body_metrics_from_reference,
                    full_metric_paragraph_indices=full_metric_indices,
                    repair_body_run_metrics_from_reference=repair_body_run_metrics_from_reference,
                    run_metric_paragraph_indices=run_metric_indices,
                    repair_body_keep_next_from_reference=repair_body_keep_next_from_reference,
                    keep_next_paragraph_indices=keep_next_indices,
                    reference_body_metrics=reference_metrics,
                )
                repaired_document_sha256 = _sha256_bytes(data)
                if original_document_sha256 != repaired_document_sha256:
                    changed_parts.append(info.filename)
            elif (
                info.filename == "word/styles.xml"
                and copy_target_style_from_reference
                and reference_docx is not None
            ):
                original_styles_sha256 = _sha256_bytes(data)
                data, style_repair_summary = repair_styles_xml_from_reference(
                    data,
                    reference_docx=reference_docx,
                    target_style_id=target_style_id,
                )
                if original_styles_sha256 != _sha256_bytes(data):
                    changed_parts.append(info.filename)
            zout.writestr(info, data)

    changed_count = int(repair_summary.get("changed_paragraph_count") or 0)
    issues = list(repair_summary.get("issues") or [])
    if repair_direct_body_metrics and issues == ["no body paragraphs required explicit style repair"]:
        issues = []
    if repair_body_line_rule_from_reference and issues == ["no body paragraphs required explicit style repair"]:
        issues = []
    if repair_full_body_metrics_from_reference and issues == ["no body paragraphs required explicit style repair"]:
        issues = []
    if repair_body_run_metrics_from_reference and issues == ["no body paragraphs required explicit style repair"]:
        issues = []
    if repair_body_keep_next_from_reference and issues == ["no body paragraphs required explicit style repair"]:
        issues = []
    if copy_target_style_from_reference and issues == ["no body paragraphs required explicit style repair"]:
        if (
            style_repair_summary.get("target_style_definition_copied")
            or style_repair_summary.get("target_style_body_baseline_materialized")
        ):
            issues = []
    return {
        "schema": "graduation-project-builder.body-explicit-style-binding-repair.v1",
        "input_docx_path": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "reference_docx_path": str(reference_docx) if reference_docx else "",
        "reference_docx_sha256": sha256_file(reference_docx) if reference_docx else "",
        "output_docx_path": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "changed_parts": changed_parts,
        "original_document_xml_sha256": original_document_sha256,
        "repaired_document_xml_sha256": repaired_document_sha256,
        **style_repair_summary,
        **selection_summary,
        **repair_summary,
        "issues": issues,
        "verdict": (
            "pass"
            if (
                not issues
                and (
                    repair_direct_body_metrics
                    or repair_body_line_rule_from_reference
                    or repair_full_body_metrics_from_reference
                    or repair_body_run_metrics_from_reference
                    or repair_body_keep_next_from_reference
                    or (
                        copy_target_style_from_reference
                        and (
                            style_repair_summary.get("target_style_definition_copied")
                            or style_repair_summary.get("target_style_body_baseline_materialized")
                        )
                    )
                    or changed_count == len(target_indices)
                )
            )
            else "fail"
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--reference-docx", type=Path)
    parser.add_argument("--style-id", help="Explicit target paragraph style id. Defaults to the reference body style match.")
    parser.add_argument("--repair-direct-body-metrics", action="store_true")
    parser.add_argument(
        "--repair-body-line-rule-from-reference",
        action="store_true",
        help="Repair only direct body paragraph lineRule values using the locked reference body baseline.",
    )
    parser.add_argument(
        "--repair-full-body-metrics-from-reference",
        action="store_true",
        help="Repair direct paragraph metrics on audited body paragraphs using the locked reference body instance baseline.",
    )
    parser.add_argument(
        "--repair-body-run-metrics-from-reference",
        action="store_true",
        help="Remove direct body run sizes and add missing Western font slots on ASCII body runs.",
    )
    parser.add_argument(
        "--repair-body-keep-next-from-reference",
        action="store_true",
        help="Remove keepNext leakage from audited body prose paragraphs while leaving captions, headings, figures, formulas, and references untouched.",
    )
    parser.add_argument("--copy-target-style-from-reference", action="store_true")
    parser.add_argument("--report-json", required=True, type=Path)
    args = parser.parse_args()

    report = repair_docx(
        args.input_docx,
        args.output_docx,
        reference_docx=args.reference_docx,
        style_id=args.style_id,
        repair_direct_body_metrics=args.repair_direct_body_metrics,
        repair_body_line_rule_from_reference=args.repair_body_line_rule_from_reference,
        repair_full_body_metrics_from_reference=args.repair_full_body_metrics_from_reference,
        repair_body_run_metrics_from_reference=args.repair_body_run_metrics_from_reference,
        repair_body_keep_next_from_reference=args.repair_body_keep_next_from_reference,
        copy_target_style_from_reference=args.copy_target_style_from_reference,
    )
    args.report_json.resolve().parent.mkdir(parents=True, exist_ok=True)
    args.report_json.resolve().write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "verdict": report["verdict"],
                "target_style_id": report["target_style_id"],
                "changed_paragraph_count": report["changed_paragraph_count"],
            },
            ensure_ascii=False,
        )
    )
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
