#!/usr/bin/env python3
"""Apply narrow front-matter, TOC, and heading structure repairs to a DOCX.

The helper is intentionally package-preserving. It rewrites only
``word/document.xml`` and, when needed, ``word/styles.xml`` or bounded
``word/header*.xml`` parts. It does not rebuild the manuscript, touch media
parts, regenerate citations, or generate new figures.
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
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ASCII_ALNUM_RE = re.compile(r"[A-Za-z0-9]")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", text or "")


def paragraph_style_id(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", NS)
    return node.get(qn("val")) if node is not None else ""


def paragraph_ppr(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find("./w:pPr", NS)


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def set_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = ensure_ppr(paragraph)
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        style = ET.Element(qn("pStyle"))
        ppr.insert(0, style)
    style.set(qn("val"), style_id)


def clear_paragraph_style(paragraph: ET.Element) -> bool:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return False
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        return False
    ppr.remove(style)
    return True


def set_frontmatter_label_metrics(paragraph: ET.Element) -> None:
    """Keep abstract/keyword label paragraphs on the template 1.5-line, 2-char rhythm."""
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    for attr in ("beforeLines", "afterLines"):
        spacing.attrib.pop(qn(attr), None)
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(qn("ind"))
        ppr.append(ind)
    ind.set(qn("firstLine"), "480")
    ind.set(qn("firstLineChars"), "200")
    for attr in ("left", "right", "hanging", "leftChars", "rightChars", "hangingChars"):
        ind.attrib.pop(qn(attr), None)


def body_element(root: ET.Element) -> ET.Element:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return body


def body_paragraphs(body: ET.Element) -> list[ET.Element]:
    return [child for child in list(body) if child.tag == qn("p")]


def style_ids(styles_root: ET.Element) -> set[str]:
    return {
        style.get(qn("styleId")) or ""
        for style in styles_root.findall("./w:style", NS)
    }


def style_name(style: ET.Element) -> str:
    name = style.find("./w:name", NS)
    return name.get(qn("val")) if name is not None else ""


def normal_style_id(styles_root: ET.Element) -> str:
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("type")) == "paragraph" and style.get(qn("default")) == "1":
            return style.get(qn("styleId")) or "Normal"
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        if style_id.lower() == "normal" or style_name(style).lower() == "normal":
            return style_id or "Normal"
    return "Normal"


def ensure_image_holder_style(styles_root: ET.Element) -> tuple[str, bool]:
    """Create or normalize the dedicated non-body image-holder paragraph style."""
    desired_id = "ThesisImageHolder"
    existing = None
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id = style.get(qn("styleId")) or ""
        if style_id.lower() == desired_id.lower():
            existing = style
            desired_id = style_id or desired_id
            break
        name = style.find("./w:name", NS)
        if name is not None and (name.get(qn("val")) or "").lower() == "thesis image holder":
            existing = style
            desired_id = style_id or desired_id
            break

    changed = False
    if existing is None:
        existing = ET.Element(qn("style"), {qn("type"): "paragraph", qn("styleId"): desired_id})
        ET.SubElement(existing, qn("name"), {qn("val"): "Thesis Image Holder"})
        ET.SubElement(existing, qn("uiPriority"), {qn("val"): "99"})
        ET.SubElement(existing, qn("unhideWhenUsed"))
        styles_root.append(existing)
        changed = True

    ppr = existing.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.SubElement(existing, qn("pPr"))
        changed = True
    if ppr.find("./w:keepNext", NS) is None:
        ppr.append(ET.Element(qn("keepNext")))
        changed = True
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.SubElement(ppr, qn("spacing"))
        changed = True
    for attr, value in {"before": "120", "after": "0", "line": "360", "lineRule": "auto"}.items():
        if spacing.get(qn(attr)) != value:
            spacing.set(qn(attr), value)
            changed = True
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.SubElement(ppr, qn("ind"))
        changed = True
    for attr in ("left", "right", "firstLine", "firstLineChars", "leftChars", "rightChars"):
        if ind.get(qn(attr)) != "0":
            ind.set(qn(attr), "0")
            changed = True
    for attr in ("hanging", "hangingChars"):
        if qn(attr) in ind.attrib:
            ind.attrib.pop(qn(attr), None)
            changed = True
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.SubElement(ppr, qn("jc"))
        changed = True
    if jc.get(qn("val")) != "center":
        jc.set(qn("val"), "center")
        changed = True
    return desired_id, changed


def copy_style_definition(styles_root: ET.Element, donor_styles_root: ET.Element, style_id: str) -> bool:
    if style_id in style_ids(styles_root):
        return False
    donor = style_by_id(donor_styles_root, style_id)
    if donor is None:
        raise RuntimeError(f"template DOCX does not define required style `{style_id}`")
    styles_root.append(deepcopy(donor))
    return True


def template_toc_style_ids(template_styles_root: ET.Element) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for style in template_styles_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        name = style_name(style).strip().lower()
        match = re.match(r"toc\s+([1-9])$", name)
        if not match:
            continue
        mapping[int(match.group(1))] = style.get(qn("styleId")) or ""
    return {level: style_id for level, style_id in mapping.items() if style_id}


def toc_level_from_style(style_id: str, styles_root: ET.Element, template_level_map: dict[int, str]) -> int | None:
    normalized = (style_id or "").strip().upper()
    match = re.match(r"TOC([1-9])$", normalized)
    if match:
        return int(match.group(1))
    for level, template_style_id in template_level_map.items():
        if style_id == template_style_id:
            return level
    style = style_by_id(styles_root, style_id)
    if style is not None:
        match = re.match(r"toc\s+([1-9])$", style_name(style).strip().lower())
        if match:
            return int(match.group(1))
    return None


def toc_placeholder_level(text: str) -> int | None:
    normalized = re.sub(r"[\u25a1\u3000]+", " ", (text or "").strip())
    normalized = re.sub(r"\s+", " ", normalized)
    if re.match(r"^\u7b2c\s*\d{1,2}\s*\u7ae0(?:\s+|$)", normalized):
        return 1
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}(?:\s+|$)", normalized):
        return 3
    if re.match(r"^\d{1,2}\.\d{1,2}(?:\s+|$)", normalized):
        return 2
    if re.match(r"^\d{1,2}(?:\s+|$)", normalized):
        return 1
    return None


def ppr_without_style(ppr: ET.Element | None) -> ET.Element | None:
    if ppr is None:
        return None
    clone = deepcopy(ppr)
    style = clone.find("./w:pStyle", NS)
    if style is not None:
        clone.remove(style)
    return clone


def replace_ppr(paragraph: ET.Element, donor_ppr: ET.Element | None, style_id: str | None = None) -> None:
    old_ppr = paragraph_ppr(paragraph)
    if old_ppr is not None:
        paragraph.remove(old_ppr)
    if donor_ppr is not None:
        paragraph.insert(0, deepcopy(donor_ppr))
    if style_id:
        set_paragraph_style(paragraph, style_id)


def direct_child_index(body: ET.Element, paragraph: ET.Element) -> int | None:
    children = list(body)
    for index, child in enumerate(children):
        if child is paragraph:
            return index
        if paragraph in child.findall(".//w:p", NS):
            return index
    return None


def iter_paragraphs_in_body_child(child: ET.Element) -> list[ET.Element]:
    if child.tag == qn("p"):
        return [child]
    return child.findall(".//w:p", NS)


def toc_entry_paragraphs_in_range(
    body: ET.Element,
    toc_index: int,
    body_start: int,
    styles_root: ET.Element,
    template_level_map: dict[int, str],
) -> list[tuple[int, ET.Element, int | None]]:
    entries: list[tuple[int, ET.Element, int | None]] = []
    children = list(body)
    range_start = toc_index if children[toc_index].tag == qn("sdt") else toc_index + 1
    for body_index in range(range_start, body_start):
        for paragraph in iter_paragraphs_in_body_child(children[body_index]):
            text = paragraph_text(paragraph).strip()
            if not text:
                continue
            if is_toc_heading(text):
                continue
            style_id = paragraph_style_id(paragraph)
            level = toc_level_from_style(style_id, styles_root, template_level_map)
            if level is None:
                inferred_heading_level = heading_level(text)
                if inferred_heading_level is not None:
                    level = max(1, min(3, inferred_heading_level))
            if level is None and is_front_matter_toc_label(text):
                level = 1
            if level is None and not is_toc_entry_paragraph(paragraph):
                continue
            entries.append((body_index, paragraph, level))
    return entries


def fallback_toc_donor_for_level(donors: dict[int, ET.Element], level: int) -> ET.Element | None:
    """Choose the nearest available TOC donor instead of leaving a level unrepaired."""
    if level in donors:
        return donors[level]
    lower_or_equal = [candidate for candidate in donors if candidate <= level]
    if lower_or_equal:
        return donors[max(lower_or_equal)]
    higher = [candidate for candidate in donors if candidate > level]
    if higher:
        return donors[min(higher)]
    return None


def collect_template_toc_donors(
    template_document_root: ET.Element | None,
    template_styles_root: ET.Element | None,
) -> dict[int, ET.Element]:
    if template_document_root is None or template_styles_root is None:
        raise RuntimeError("toc-template-styles requires --template-docx")
    template_body = template_document_root.find("./w:body", NS)
    if template_body is None:
        raise RuntimeError("template DOCX has no word/document.xml body")
    template_level_map = template_toc_style_ids(template_styles_root)
    donors: dict[int, ET.Element] = {}
    try:
        toc_index, body_start = find_toc_and_body_start(template_body)
        entries = toc_entry_paragraphs_in_range(
            template_body,
            toc_index,
            body_start,
            template_styles_root,
            template_level_map,
        )
    except RuntimeError:
        children = list(template_body)
        toc_index = None
        body_start = len(children)
        for index, child in enumerate(children):
            if child.tag != qn("p"):
                continue
            text = paragraph_text(child)
            compact = compact_text(text)
            if "\u76ee\u5f55" in compact or "\u76ee\u5f55\u4ece" in compact or "\u76ee" in compact and "\u7ae0" in compact:
                toc_index = index
                break
        if toc_index is None:
            raise RuntimeError("template DOCX has no measurable TOC entry donor paragraphs")
        entries = []
        for index in range(toc_index + 1, len(children)):
            child = children[index]
            if child.tag != qn("p"):
                continue
            text = paragraph_text(child).strip()
            compact = compact_text(text)
            if "\u6ce8:\u76ee\u5f55\u5355\u72ec\u6210\u9875" in compact or "\u6ce8\uff1a\u76ee\u5f55\u5355\u72ec\u6210\u9875" in compact:
                body_start = index
                break
            level = toc_level_from_style(paragraph_style_id(child), template_styles_root, template_level_map)
            if level is None:
                level = toc_placeholder_level(text)
            if level is not None:
                entries.append((index, child, level))
    for _body_index, paragraph, level in entries:
        if level is None:
            continue
        donors.setdefault(level, paragraph)
    if not donors:
        raise RuntimeError("template DOCX has no measurable TOC entry donor paragraphs")
    return donors


def toc_child_content(child: ET.Element) -> ET.Element | None:
    if child.tag == qn("sdt"):
        return child.find("./w:sdtContent", NS)
    return None


def toc_entry_like(paragraph: ET.Element, styles_root: ET.Element, template_level_map: dict[int, str]) -> bool:
    text = paragraph_text(paragraph).strip()
    if not text or is_toc_heading(text):
        return False
    if toc_level_from_style(paragraph_style_id(paragraph), styles_root, template_level_map) is not None:
        return True
    if heading_level(text) is not None:
        return True
    return is_toc_entry_paragraph(paragraph)


def first_toc_entry_child_index(
    container: ET.Element,
    styles_root: ET.Element,
    template_level_map: dict[int, str],
) -> int | None:
    for index, child in enumerate(list(container)):
        if child.tag != qn("p"):
            continue
        if toc_entry_like(child, styles_root, template_level_map):
            return index
    return None


def collect_template_toc_prefix(
    template_document_root: ET.Element | None,
    template_styles_root: ET.Element | None,
) -> list[ET.Element]:
    if template_document_root is None or template_styles_root is None:
        raise RuntimeError("toc-template-styles requires --template-docx")
    template_body = template_document_root.find("./w:body", NS)
    if template_body is None:
        raise RuntimeError("template DOCX has no word/document.xml body")
    template_level_map = template_toc_style_ids(template_styles_root)
    try:
        toc_index, _body_start = find_toc_and_body_start(template_body)
    except RuntimeError:
        return []
    toc_child = list(template_body)[toc_index]
    content = toc_child_content(toc_child)
    if content is None:
        return []
    first_entry = first_toc_entry_child_index(content, template_styles_root, template_level_map)
    if first_entry is None:
        raise RuntimeError("template TOC content-control has no first entry paragraph")
    prefix = [child for child in list(content)[:first_entry] if child.tag == qn("p")]
    if not any(is_toc_heading(paragraph_text(paragraph).strip()) for paragraph in prefix):
        raise RuntimeError("template TOC content-control has no title paragraph before entries")
    return [deepcopy(paragraph) for paragraph in prefix]


def replay_template_toc_prefix(
    body: ET.Element,
    toc_index: int,
    styles_root: ET.Element,
    template_styles_root: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    children = list(body)
    toc_child = children[toc_index]
    content = toc_child_content(toc_child)
    if content is None:
        return []
    preceding_removed: list[str] = []
    probe_index = toc_index - 1
    while probe_index >= 0:
        candidate = list(body)[probe_index]
        if candidate.tag != qn("p"):
            break
        candidate_text = paragraph_text(candidate).strip()
        if candidate_text and not is_toc_heading(candidate_text):
            break
        preceding_removed.append(candidate_text)
        body.remove(candidate)
        probe_index -= 1
        continue
    template_level_map = template_toc_style_ids(template_styles_root)
    first_entry = first_toc_entry_child_index(content, styles_root, template_level_map)
    if first_entry is None:
        raise RuntimeError("target TOC content-control has no first entry paragraph")
    prefix = collect_template_toc_prefix(template_document_root, template_styles_root)
    if not prefix:
        return []
    before_prefix_text = [paragraph_text(child) for child in list(content)[:first_entry] if child.tag == qn("p")]
    for child in list(content)[:first_entry]:
        content.remove(child)
    for offset, clone in enumerate(prefix):
        content.insert(offset, deepcopy(clone))
    return [
        {
            "kind": "toc_prefix_replayed",
            "body_child_index": toc_index,
            "removed_prefix_paragraph_count": first_entry,
            "removed_preceding_standalone_toc_prefix": list(reversed(preceding_removed)),
            "inserted_prefix_paragraph_count": len(prefix),
            "before_prefix_text": before_prefix_text,
            "inserted_prefix_text": [paragraph_text(paragraph) for paragraph in prefix],
        }
    ]


def style_xml_equal(left: ET.Element | None, right: ET.Element | None) -> bool:
    if left is None or right is None:
        return left is right
    return ET.tostring(left, encoding="utf-8") == ET.tostring(right, encoding="utf-8")


def paragraphs_using_style(body: ET.Element, style_id: str) -> list[ET.Element]:
    return [paragraph for paragraph in body.findall(".//w:p", NS) if paragraph_style_id(paragraph) == style_id]


def unique_style_id(styles_root: ET.Element, base: str) -> str:
    existing = style_ids(styles_root)
    candidate = base
    counter = 1
    while candidate in existing:
        counter += 1
        candidate = f"{base}{counter}"
    return candidate


def clone_style_with_id(donor_style: ET.Element, new_style_id: str) -> ET.Element:
    clone = deepcopy(donor_style)
    clone.set(qn("styleId"), new_style_id)
    name = clone.find("./w:name", NS)
    if name is not None:
        name.set(qn("val"), f"{name.get(qn('val')) or new_style_id} donor")
    return clone


def ensure_template_toc_style_mapping(
    body: ET.Element,
    styles_root: ET.Element,
    template_styles_root: ET.Element,
    target_toc_paragraphs: set[int],
) -> tuple[dict[int, str], list[dict[str, object]]]:
    template_level_map = template_toc_style_ids(template_styles_root)
    changes: list[dict[str, object]] = []
    selected: dict[int, str] = {}
    for level, template_style_id in sorted(template_level_map.items()):
        donor_style = style_by_id(template_styles_root, template_style_id)
        if donor_style is None:
            raise RuntimeError(f"template DOCX does not define required style `{template_style_id}`")
        existing = style_by_id(styles_root, template_style_id)
        if existing is None:
            styles_root.append(deepcopy(donor_style))
            selected[level] = template_style_id
            changes.append({"kind": "style_copied", "level": level, "style_id": template_style_id})
            continue
        if style_xml_equal(existing, donor_style):
            selected[level] = template_style_id
            continue
        users = paragraphs_using_style(body, template_style_id)
        non_toc_users = [paragraph for paragraph in users if id(paragraph) not in target_toc_paragraphs]
        if not non_toc_users:
            parent_styles = list(styles_root)
            replace_at = parent_styles.index(existing)
            styles_root.remove(existing)
            styles_root.insert(replace_at, deepcopy(donor_style))
            selected[level] = template_style_id
            changes.append(
                {
                    "kind": "style_replaced_no_non_toc_usage",
                    "level": level,
                    "style_id": template_style_id,
                    "prior_user_count": len(users),
                }
            )
            continue
        safe_style_id = unique_style_id(styles_root, f"GPBToc{level}")
        styles_root.append(clone_style_with_id(donor_style, safe_style_id))
        selected[level] = safe_style_id
        changes.append(
            {
                "kind": "style_cloned_due_to_conflict",
                "level": level,
                "template_style_id": template_style_id,
                "new_style_id": safe_style_id,
                "non_toc_user_count": len(non_toc_users),
            }
        )
    return selected, changes


def visible_run_role(run: ET.Element, *, seen_tab: bool) -> str | None:
    if run.find(".//w:tab", NS) is not None:
        return "tab"
    text = run_text(run).strip()
    if not text:
        return None
    if seen_tab or re.fullmatch(r"(?:\d+|[ivxlcdmIVXLCDM\u2160-\u2188]+)", re.sub(r"\s+", "", text)):
        return "page_number"
    return "text"


def run_role_rprs(paragraph: ET.Element) -> dict[str, ET.Element | None]:
    roles: dict[str, ET.Element | None] = {}
    seen_tab = False
    for run in paragraph.findall(".//w:r", NS):
        role = visible_run_role(run, seen_tab=seen_tab)
        if role == "tab":
            seen_tab = True
        if role is None:
            continue
        roles.setdefault(role, deepcopy(run.find("./w:rPr", NS)) if run.find("./w:rPr", NS) is not None else None)
    return roles


def apply_visible_run_rprs(paragraph: ET.Element, role_rprs: dict[str, ET.Element | None]) -> dict[str, int]:
    counts: dict[str, int] = {}
    seen_tab = False
    for run in paragraph.findall(".//w:r", NS):
        role = visible_run_role(run, seen_tab=seen_tab)
        if role == "tab":
            seen_tab = True
        if role is None or role not in role_rprs:
            continue
        old_rpr = run.find("./w:rPr", NS)
        if old_rpr is not None:
            run.remove(old_rpr)
        donor_rpr = role_rprs.get(role)
        if donor_rpr is not None:
            run.insert(0, deepcopy(donor_rpr))
        counts[role] = counts.get(role, 0) + 1
    return counts


def direct_text_children(paragraph: ET.Element) -> list[str]:
    return [node.text or "" for node in paragraph.findall(".//w:t", NS)]


def toc_entry_label_and_page_from_paragraph(paragraph: ET.Element) -> tuple[str, str] | None:
    texts = direct_text_children(paragraph)
    if not texts:
        return None
    page_pattern = re.compile(r"^\s*(?:\d+|[ivxlcdmIVXLCDM\u2160-\u2188]+)\s*$", re.IGNORECASE)
    for index in range(len(texts) - 1, -1, -1):
        value = texts[index].strip()
        if page_pattern.fullmatch(value):
            label = "".join(texts[:index]).strip()
            return (toc_entry_visible_label(label), value) if label else None
    joined = "".join(texts).strip()
    static_entry = static_toc_label_and_page(joined)
    if static_entry is not None:
        return static_entry
    return None


def append_toc_run(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> ET.Element:
    run = ET.SubElement(paragraph, qn("r"))
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    node = ET.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    return run


def append_toc_tab_run(paragraph: ET.Element, donor_rpr: ET.Element | None) -> ET.Element:
    run = ET.SubElement(paragraph, qn("r"))
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    ET.SubElement(run, qn("tab"))
    return run


def field_char_run(kind: str) -> ET.Element:
    run = ET.Element(qn("r"))
    fld = ET.SubElement(run, qn("fldChar"))
    fld.set(qn("fldCharType"), kind)
    return run


def toc_instruction_run() -> ET.Element:
    run = ET.Element(qn("r"))
    instr = ET.SubElement(run, qn("instrText"))
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    return run


def live_toc_instruction_count(root: ET.Element) -> int:
    count = 0
    for node in root.findall(".//w:instrText", NS):
        if re.search(r"\bTOC\b", node.text or "", re.IGNORECASE):
            count += 1
    for node in root.findall(".//w:fldSimple", NS):
        if re.search(r"\bTOC\b", node.get(qn("instr")) or "", re.IGNORECASE):
            count += 1
    return count


def insert_toc_field_start(paragraph: ET.Element) -> None:
    insert_at = 0
    ppr = paragraph_ppr(paragraph)
    if ppr is not None:
        insert_at = list(paragraph).index(ppr) + 1
    paragraph.insert(insert_at, field_char_run("begin"))
    paragraph.insert(insert_at + 1, toc_instruction_run())
    paragraph.insert(insert_at + 2, field_char_run("separate"))


def wrap_static_toc_cache_in_live_field(
    body: ET.Element,
    root: ET.Element,
    styles_root: ET.Element,
) -> dict[str, object]:
    existing_count = live_toc_instruction_count(root)
    if existing_count:
        return {
            "status": "already_present",
            "live_toc_field_count": existing_count,
            "entries_wrapped": 0,
        }
    toc_index, body_start = find_toc_and_body_start(body)
    entries = [
        (body_index, paragraph, level)
        for body_index, paragraph, level in toc_entry_paragraphs_in_range(body, toc_index, body_start, styles_root, {})
        if paragraph_text(paragraph).strip() and is_toc_entry_paragraph(paragraph)
    ]
    if not entries:
        raise RuntimeError("live-toc-field requires cached TOC entry paragraphs to wrap")
    first_index, first_paragraph, _first_level = entries[0]
    last_index, last_paragraph, _last_level = entries[-1]
    insert_toc_field_start(first_paragraph)
    last_paragraph.append(field_char_run("end"))
    return {
        "status": "wrapped_static_cache",
        "live_toc_field_count": live_toc_instruction_count(root),
        "entries_wrapped": len(entries),
        "first_toc_entry_body_child_index": first_index,
        "last_toc_entry_body_child_index": last_index,
        "field_instruction": 'TOC \\o "1-3" \\h \\z \\u',
        "preserved_cached_result": True,
    }


def rewrite_toc_entry_runs(paragraph: ET.Element, role_rprs: dict[str, ET.Element | None]) -> dict[str, int]:
    label_page = toc_entry_label_and_page_from_paragraph(paragraph)
    if label_page is None:
        text = paragraph_text(paragraph).strip()
        style_id = paragraph_style_id(paragraph).lower()
        if not text or (not style_id.startswith("toc") and toc_placeholder_level(text) is None):
            return {}
        label = toc_entry_visible_label(text)
        if not label:
            return {}
        label_page = (label, "0")
    label, page = label_page
    starts, ends = review_anchor_children(paragraph)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    append_toc_run(paragraph, label, role_rprs.get("text"))
    append_toc_tab_run(paragraph, role_rprs.get("tab"))
    append_toc_run(paragraph, page, role_rprs.get("page_number") or role_rprs.get("text"))
    for child in ends:
        paragraph.append(child)
    return {"text": 1, "tab": 1, "page_number": 1}


def repair_toc_template_styles(
    body: ET.Element,
    styles_root: ET.Element,
    template_styles_root: ET.Element | None,
    template_document_root: ET.Element | None = None,
) -> list[dict[str, object]]:
    if template_styles_root is None:
        raise RuntimeError("toc-template-styles requires --template-docx")
    template_level_map = template_toc_style_ids(template_styles_root)
    donors = collect_template_toc_donors(template_document_root, template_styles_root)
    changes: list[dict[str, object]] = []

    toc_index, body_start = find_toc_and_body_start(body)
    changes.extend(
        replay_template_toc_prefix(
            body,
            toc_index,
            styles_root,
            template_styles_root,
            template_document_root,
        )
    )
    toc_index, body_start = find_toc_and_body_start(body)
    entries = toc_entry_paragraphs_in_range(body, toc_index, body_start, styles_root, template_level_map)
    target_toc_paragraphs = {id(paragraph) for _index, paragraph, _level in entries}
    selected_styles, style_changes = ensure_template_toc_style_mapping(
        body,
        styles_root,
        template_styles_root,
        target_toc_paragraphs,
    )
    changes.extend(style_changes)
    for body_index, paragraph, level in entries:
        if level is None:
            level = 1
        donor = fallback_toc_donor_for_level(donors, level)
        if donor is None:
            continue
        old_style_id = paragraph_style_id(paragraph)
        selected_style_id = selected_styles.get(level)
        role_rprs = run_role_rprs(donor)
        replace_ppr(paragraph, ppr_without_style(paragraph_ppr(donor)), selected_style_id)
        role_counts = rewrite_toc_entry_runs(paragraph, role_rprs)
        if not role_counts:
            role_counts = apply_visible_run_rprs(paragraph, role_rprs)
        changes.append(
            {
                "kind": "toc_paragraph_template_replayed",
                "body_child_index": body_index,
                "level": level,
                "old_style_id": old_style_id,
                "new_style_id": selected_style_id,
                "template_style_id": template_level_map.get(level),
                "visible_run_role_counts": role_counts,
                "text": paragraph_text(paragraph).strip()[:120],
            }
        )
    return changes


def repair_toc_level_indents(
    body: ET.Element,
    styles_root: ET.Element,
    *,
    level_indent_twips: int = 420,
) -> list[dict[str, object]]:
    """Apply visible hierarchical indents to cached TOC entries.

    Some official school templates are instruction guides rather than clean
    finished samples, so they do not always carry a reliable TOC level style.
    This narrow operation preserves the cached TOC runs and dotted right tab,
    while giving level-2 and deeper entries a visible left indent.
    """
    toc_index, body_start = find_toc_and_body_start(body)
    entries = toc_entry_paragraphs_in_range(body, toc_index, body_start, styles_root, {})
    changes: list[dict[str, object]] = []
    for body_index, paragraph, level in entries:
        if level is None:
            continue
        ppr = ensure_ppr(paragraph)
        ind = ppr.find("./w:ind", NS)
        if ind is None:
            ind = ET.Element(qn("ind"))
            ppr.append(ind)
        before = dict(ind.attrib)
        left = max(0, (level - 1) * level_indent_twips)
        ind.set(qn("left"), str(left))
        ind.set(qn("firstLine"), "0")
        for attr_name in ("leftChars", "rightChars", "firstLineChars", "hanging", "hangingChars"):
            ind.attrib.pop(qn(attr_name), None)
        if before != dict(ind.attrib):
            changes.append(
                {
                    "kind": "toc_level_indent_repaired",
                    "body_child_index": body_index,
                    "level": level,
                    "left_twips": left,
                    "first_line_twips": 0,
                    "text": paragraph_text(paragraph).strip()[:120],
                }
            )
    return changes


def heading_level(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    if "\u2026" in stripped:
        return None
    normalized = re.sub(r"(?<=\d)[\s\u25a1]*[\.\uff0e．-][\s\u25a1]*(?=\d)", ".", stripped)
    if re.match(r"^\u7b2c\s*\d{1,2}\s*\u7ae0(?:[\s\u3000]+|$)\S*", stripped):
        return 1
    if re.match(r"^\d{1,2}\s+\S", stripped):
        return 1
    if re.match(r"^\d{1,2}\.\d{1,2}\s+\S", normalized):
        return 2
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\s+\S", normalized):
        return 3
    return None


def int_attr(node: ET.Element | None, attr_name: str) -> int | None:
    if node is None:
        return None
    raw = node.get(qn(attr_name))
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def paragraph_heading_level(paragraph: ET.Element) -> int | None:
    text_level = heading_level(paragraph_text(paragraph).strip())
    if text_level is not None:
        return text_level
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return None
    outline = int_attr(ppr.find("./w:outlineLvl", NS), "val")
    ilvl = int_attr(ppr.find("./w:numPr/w:ilvl", NS), "val")
    num_id = int_attr(ppr.find("./w:numPr/w:numId", NS), "val")
    if outline is None or outline < 0 or outline > 2:
        return None
    if ilvl is not None and ilvl != outline:
        return None
    if outline == 0 and num_id == 0:
        return None
    return outline + 1


def paragraph_style_name(paragraph: ET.Element, styles_root: ET.Element | None) -> str:
    if styles_root is None:
        return ""
    style_id = paragraph_style_id(paragraph)
    if not style_id:
        return ""
    style = style_by_id(styles_root, style_id)
    return style_name(style).lower() if style is not None else ""


def paragraph_has_explicit_outline(paragraph: ET.Element) -> bool:
    ppr = paragraph_ppr(paragraph)
    return ppr is not None and ppr.find("./w:outlineLvl", NS) is not None


def paragraph_heading_level_for_template_donor(paragraph: ET.Element, styles_root: ET.Element | None) -> int | None:
    text = paragraph_text(paragraph).strip()
    if re.match(r"^\u7b2c\s*\d{1,2}\s*\u7ae0(?:[\s\u3000]+|$)\S*", text):
        return 1
    level = paragraph_heading_level(paragraph)
    if level is None:
        return None
    if level == 1 and not paragraph_has_explicit_outline(paragraph):
        style_id = paragraph_style_id(paragraph) or ""
        style_name_value = paragraph_style_name(paragraph, styles_root)
        if style_id.lower() not in {"heading1", "heading 1"} and "heading 1" not in style_name_value:
            return None
    return level


def is_toc_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    toc = compact_text("\u76ee\u5f55").lower()
    return normalized in {toc, "contents", "tableofcontents"} or (
        normalized.endswith(toc) and len(normalized) <= 120
    )


def is_front_matter_toc_label(text: str) -> bool:
    if re.search(r"(?:\d+|[IVXLCDM]+)\s*$", text or "", flags=re.IGNORECASE) is None:
        return False
    label = compact_text(re.sub(r"(?:\d+|[IVXLCDM]+)\s*$", "", text or "", flags=re.IGNORECASE)).lower()
    return label in {compact_text("\u6458\u8981").lower(), "abstract"}


def static_toc_label_and_page(text: str) -> tuple[str, str] | None:
    stripped = (text or "").strip()
    if not stripped or "\t" in stripped:
        return None
    match = re.match(r"^(?P<label>.+?)(?P<page>\d+|[ivxlcdmIVXLCDM]+)\s*$", stripped)
    if not match:
        return None
    label = match.group("label").strip()
    page = match.group("page").strip()
    if not label or not page:
        return None
    return label, page


def is_static_toc_entry_label(label: str) -> bool:
    normalized = compact_text(label).lower()
    if normalized in {
        compact_text("\u6458\u8981").lower(),
        "abstract",
        compact_text(REFERENCES_LABEL),
        compact_text(ACKNOWLEDGEMENT_LABEL),
        compact_text(APPENDIX_LABEL),
        compact_text("\u7ed3\u8bba"),
    }:
        return True
    return heading_level(label) is not None


def has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None


def is_toc_entry_paragraph(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph).strip()
    style_id = paragraph_style_id(paragraph).lower()
    if style_id.startswith("toc"):
        return True
    if has_tab(paragraph) or "\u2026" in text:
        return True
    static_entry = static_toc_label_and_page(text)
    if static_entry is not None and is_static_toc_entry_label(static_entry[0]):
        return True
    return is_front_matter_toc_label(text)


CAPTION_RE = re.compile(r"^\s*[\u56fe\u8868]\s*\d+(?:[-.\uff0d]\d+)?")


def is_zh_abstract_label(text: str) -> bool:
    compact = compact_text(re.sub(r"[:：].*$", "", text or ""))
    marker = "\u6458\u8981"
    return compact == marker or (compact.endswith(marker) and len(compact) <= 120)


def is_en_abstract_label(text: str) -> bool:
    stripped = (text or "").strip()
    compact = re.sub(r"[\s:：]+", "", stripped).lower()
    return compact.startswith("abstract") or (compact.endswith("abstract") and len(compact) <= 160)


def is_zh_keyword_line(text: str) -> bool:
    return compact_text(text).startswith("\u5173\u952e\u8bcd")


def is_en_keyword_line(text: str) -> bool:
    lowered = (text or "").strip().lower()
    return lowered.startswith(("key words", "keywords", "keyword"))


def find_first_paragraph(body: ET.Element, predicate) -> tuple[int, ET.Element] | None:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and predicate(child):
            return index, child
    return None


def next_nonempty_paragraph(body: ET.Element, start_index: int, stop_predicate) -> tuple[int, ET.Element] | None:
    children = list(body)
    for index in range(start_index + 1, len(children)):
        child = children[index]
        if child.tag != qn("p"):
            continue
        if stop_predicate(child):
            return None
        if paragraph_text(child).strip():
            return index, child
    return None


CAPTION_RE = re.compile(
    r"^(?:\u56fe|\u8868|\u7eed\u8868)\s*"
    r"[0-9A-Za-z\u4e00-\u9fff]+(?:[-\.－][0-9A-Za-z\u4e00-\u9fff]+)*"
    r"(?=$|[\s\u3000:：、])"
)


CAPTION_NUMBER_RE = re.compile(r"^(?:\u56fe|\u8868|\u7eed\u8868)\s*(?P<major>\d+)[\-\.．](?P<minor>\d+)")
REFERENCES_LABEL = "\u53c2\u8003\u6587\u732e"
ACKNOWLEDGEMENT_LABEL = "\u81f4\u8c22"
APPENDIX_LABEL = "\u9644\u5f55"
TAIL_LABELS = {REFERENCES_LABEL, ACKNOWLEDGEMENT_LABEL, APPENDIX_LABEL}


def remove_page_break_runs(paragraph: ET.Element) -> int:
    removed = 0
    for run in list(paragraph.findall("./w:r", NS)):
        page_breaks = [
            node for node in run.findall("./w:br", NS)
            if (node.get(qn("type")) or "page") == "page"
        ]
        if not page_breaks:
            continue
        if not paragraph_text(run).strip() and len(page_breaks) == len(run.findall("./w:br", NS)):
            paragraph.remove(run)
            removed += len(page_breaks)
            continue
        for br in page_breaks:
            run.remove(br)
            removed += 1
    return removed


def suppress_inherited_page_break_before(paragraph: ET.Element) -> bool:
    """Add a direct off switch when an inherited heading style owns pagination."""
    ppr = ensure_ppr(paragraph)
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        page_break = ET.Element(qn("pageBreakBefore"))
        page_break.set(qn("val"), "0")
        ppr.insert(0, page_break)
        return True
    current = page_break.get(qn("val"))
    if current in {"0", "false", "False", "off", "OFF"}:
        return False
    page_break.set(qn("val"), "0")
    return True


def remove_section_owned_opener_page_break_before(paragraph: ET.Element) -> bool:
    """Remove direct pageBreakBefore when the previous section break owns the start page.

    WPS may still paginate on a present ``w:pageBreakBefore`` element even when
    ``w:val="0"`` is set. For a body opener immediately after a section-break
    paragraph, the section break is the durable page owner, so the opener must
    not retain a direct pageBreakBefore node.
    """
    return remove_page_break_before_property(paragraph)


def normalize_duplicate_page_breaks(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    children = list(body)
    for index, child in enumerate(children):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        ppr = paragraph_ppr(child)
        page_before = ppr is not None and ppr.find("./w:pageBreakBefore", NS) is not None
        section_break = ppr is not None and ppr.find("./w:sectPr", NS) is not None
        previous_section_break = False
        if index > 0 and children[index - 1].tag == qn("p"):
            previous_ppr = paragraph_ppr(children[index - 1])
            previous_section_break = previous_ppr is not None and previous_ppr.find("./w:sectPr", NS) is not None
        should_remove = False
        reason = ""
        if not text and section_break:
            should_remove = True
            reason = "empty section paragraph keeps section break; remove duplicate hard page break"
        elif is_toc_heading(text) and page_before:
            should_remove = True
            reason = "TOC title owns pageBreakBefore; remove duplicate hard page-break runs only"
        elif paragraph_heading_level(child) == 1 and previous_section_break:
            should_remove = True
            reason = "previous paragraph owns the section break; remove duplicate hard page break on body opener"
        elif paragraph_heading_level(child) == 1 and page_before:
            should_remove = True
            reason = "level-1 body heading already has pageBreakBefore"
        if not should_remove:
            continue
        removed_page_break_before = False
        removed = remove_page_break_runs(child)
        suppressed_inherited = False
        removed_section_owned_page_break_before = False
        if paragraph_heading_level(child) == 1 and previous_section_break:
            removed_section_owned_page_break_before = remove_section_owned_opener_page_break_before(child)
        if (
            removed
            or removed_page_break_before
            or suppressed_inherited
            or removed_section_owned_page_break_before
        ):
            changes.append({"index": index, "text": text, "removed_page_break_runs": removed, "reason": reason})
            if removed_page_break_before:
                changes[-1]["removed_pageBreakBefore"] = True
            if removed_section_owned_page_break_before:
                changes[-1]["removed_section_owned_pageBreakBefore"] = True
            if suppressed_inherited:
                changes[-1]["suppressed_inherited_pageBreakBefore"] = True
    return changes


def remove_toc_contamination(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    children = list(body)
    toc_exists = False
    for child in children:
        if child.tag != qn("p"):
            continue
        if is_toc_heading(paragraph_text(child).strip()):
            toc_exists = True
            break
    if not toc_exists:
        return changes

    toc_started = False
    for child in list(body):
        if child.tag == qn("p") and is_toc_heading(paragraph_text(child).strip()):
            toc_started = True
            continue
        if not toc_started:
            continue
        if child.tag == qn("p"):
            text = paragraph_text(child).strip()
            if paragraph_heading_level(child) == 1 and not is_toc_entry_paragraph(child):
                break
            if is_toc_entry_paragraph(child) or not text:
                continue
            if CAPTION_RE.match(text) or text:
                body.remove(child)
                changes.append({"kind": "paragraph", "text": text[:160], "reason": "non-TOC paragraph inside TOC range"})
            continue
        if child.tag == qn("tbl"):
            text = "".join(node.text or "" for node in child.findall(".//w:t", NS)).strip()
            body.remove(child)
            changes.append({"kind": "table", "text": text[:160], "reason": "table inside TOC range"})
    return changes


def remove_empty_toc_entries(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_index, body_start = find_toc_and_body_start(body)
    for child in list(body)[toc_index + 1 : body_start]:
        if child.tag != qn("p"):
            continue
        if paragraph_text(child).strip():
            continue
        style_id = paragraph_style_id(child)
        if style_id.upper().startswith("TOC"):
            body.remove(child)
            changes.append({"kind": "paragraph", "style_id": style_id, "reason": "empty TOC entry inside TOC range"})
    return changes


def find_toc_and_body_start(body: ET.Element) -> tuple[int, int]:
    children = list(body)
    toc_index = None
    for index, child in enumerate(children):
        if child.tag == qn("sdt") and re.search(
            r"\bTOC\b",
            " ".join(node.text or "" for node in child.findall(".//w:instrText", NS)),
            re.IGNORECASE,
        ):
            toc_index = index
            break
    for index, child in enumerate(children):
        if toc_index is not None:
            break
        if child.tag == qn("p") and is_toc_heading(paragraph_text(child).strip()):
            toc_index = index
            break
        if child.tag == qn("sdt") and (
            any(is_toc_heading(paragraph_text(paragraph).strip()) for paragraph in child.findall(".//w:p", NS))
            or re.search(r"\bTOC\b", " ".join(node.text or "" for node in child.findall(".//w:instrText", NS)))
        ):
            toc_index = index
            break
    if toc_index is None:
        raise RuntimeError("TOC title not found; refusing TOC structure repair")
    for index in range(toc_index + 1, len(children)):
        child = children[index]
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if paragraph_heading_level(child) == 1 and not is_toc_entry_paragraph(child):
            return toc_index, index
    raise RuntimeError("first body level-1 heading not found after TOC; refusing TOC structure repair")


def heading_major_minor(text: str) -> tuple[int, int | None] | None:
    normalized = re.sub(r"^[\s\u25a1]+", "", (text or "").strip()).replace("\uff0e", ".").replace("\uff61", ".").replace("．", ".")
    match = re.match(r"^(?P<major>\d+)(?:\.(?P<minor>\d+))?(?:[\s\u3000]+|$)", normalized)
    if not match:
        return None
    return int(match.group("major")), int(match.group("minor")) if match.group("minor") else None


def find_body_relocation_anchor(body: ET.Element, *, major: int) -> ET.Element | None:
    _toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    for child in children[body_start:]:
        if child.tag == qn("sectPr"):
            return child
        if child.tag != qn("p"):
            continue
        compact = compact_text(paragraph_text(child))
        if compact in {compact_text(label) for label in TAIL_LABELS}:
            return child
        major_minor = heading_major_minor(paragraph_text(child))
        if major_minor is None:
            continue
        heading_major, heading_minor = major_minor
        if heading_major == major and heading_minor is not None and heading_minor >= 3:
            return child
        if heading_major > major and paragraph_heading_level(child) == 1:
            return child
    return None


def relocate_toc_contamination(body: ET.Element) -> list[dict[str, object]]:
    toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    groups: list[dict[str, object]] = []
    index = toc_index + 1
    while index < body_start:
        child = children[index]
        if child.tag == qn("p"):
            text = paragraph_text(child).strip()
            if not text or is_toc_entry_paragraph(child):
                index += 1
                continue
            caption_match = CAPTION_NUMBER_RE.match(text)
            if caption_match and index + 1 < body_start and children[index + 1].tag == qn("tbl"):
                groups.append(
                    {
                        "nodes": [child, children[index + 1]],
                        "major": int(caption_match.group("major")),
                        "text": text,
                    }
                )
                index += 2
                continue
            raise RuntimeError(f"non-relocatable paragraph inside TOC range: {text[:80]}")
        if child.tag == qn("tbl"):
            raise RuntimeError("standalone table inside TOC range has no adjacent caption; refusing relocation")
        if "".join(node.text or "" for node in child.findall(".//w:t", NS)).strip():
            raise RuntimeError(f"non-relocatable content inside TOC range: {child.tag.rsplit('}', 1)[-1]}")
        index += 1

    changes: list[dict[str, object]] = []
    for group in groups:
        nodes = list(group["nodes"])
        major = int(group["major"])
        for node in nodes:
            body.remove(node)
        anchor = find_body_relocation_anchor(body, major=major)
        current_children = list(body)
        insert_at = current_children.index(anchor) if anchor is not None else len(current_children)
        for offset, node in enumerate(nodes):
            body.insert(insert_at + offset, node)
        changes.append(
            {
                "kind": "caption_table_group",
                "text": str(group["text"])[:160],
                "major": major,
                "inserted_before": paragraph_text(anchor).strip()[:120] if anchor is not None and anchor.tag == qn("p") else "body-final-section",
                "reason": "caption/table group relocated from protected TOC range back to body",
            }
        )
    return changes


def exact_tail_label(text: str) -> str | None:
    normalized = compact_text(text)
    for label in TAIL_LABELS:
        if normalized == compact_text(label):
            return label
    return None


def find_tail_heading(body: ET.Element, label: str) -> tuple[int, ET.Element] | None:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and compact_text(paragraph_text(child)) == compact_text(label):
            return index, child
    return None


def tail_block_end(children: list[ET.Element], start_index: int) -> int:
    for index in range(start_index + 1, len(children)):
        child = children[index]
        if child.tag == qn("sectPr"):
            return index
        if child.tag == qn("p") and exact_tail_label(paragraph_text(child)):
            return index
    return len(children)


def ensure_page_break_before(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    if ppr.find("./w:pageBreakBefore", NS) is not None:
        return False
    page_break = ET.Element(qn("pageBreakBefore"))
    style = ppr.find("./w:pStyle", NS)
    insert_at = list(ppr).index(style) + 1 if style is not None else 0
    ppr.insert(insert_at, page_break)
    return True


def ensure_outline_level(paragraph: ET.Element, level: str = "0") -> bool:
    ppr = ensure_ppr(paragraph)
    outline = ppr.find("./w:outlineLvl", NS)
    if outline is None:
        outline = ET.Element(qn("outlineLvl"))
        ppr.append(outline)
    if outline.get(qn("val")) == level:
        return False
    outline.set(qn("val"), level)
    return True


def max_bookmark_id(root: ET.Element) -> int:
    values = []
    for node in root.iter(qn("bookmarkStart")):
        raw = node.get(qn("id"))
        if raw and raw.isdigit():
            values.append(int(raw))
    return max(values, default=0)


def ensure_bookmark(paragraph: ET.Element, root: ET.Element, name: str) -> bool:
    for node in paragraph.findall("./w:bookmarkStart", NS):
        if node.get(qn("name")) == name:
            return False
    bookmark_id = str(max_bookmark_id(root) + 1)
    start = ET.Element(qn("bookmarkStart"))
    start.set(qn("id"), bookmark_id)
    start.set(qn("name"), name)
    end = ET.Element(qn("bookmarkEnd"))
    end.set(qn("id"), bookmark_id)
    ppr = paragraph_ppr(paragraph)
    insert_at = 1 if ppr is not None and list(paragraph) and list(paragraph)[0] is ppr else 0
    paragraph.insert(insert_at, start)
    paragraph.append(end)
    return True


def repair_tail_block_order_and_openers(body: ET.Element, root: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    references = find_tail_heading(body, REFERENCES_LABEL)
    acknowledgement = find_tail_heading(body, ACKNOWLEDGEMENT_LABEL)
    if references is None or acknowledgement is None:
        raise RuntimeError("tail-block-order requires both references and acknowledgement headings")
    references_index, references_para = references
    acknowledgement_index, acknowledgement_para = acknowledgement
    if references_index > acknowledgement_index:
        children = list(body)
        block_end = tail_block_end(children, references_index)
        references_block = children[references_index:block_end]
        for node in references_block:
            body.remove(node)
        current_children = list(body)
        insert_at = current_children.index(acknowledgement_para)
        for offset, node in enumerate(references_block):
            body.insert(insert_at + offset, node)
        changes.append(
            {
                "kind": "tail_block_order",
                "moved": REFERENCES_LABEL,
                "from_index": references_index,
                "before": ACKNOWLEDGEMENT_LABEL,
                "node_count": len(references_block),
            }
        )

    for label, bookmark_name in (
        (REFERENCES_LABEL, "_TocTailReferences"),
        (ACKNOWLEDGEMENT_LABEL, "_TocTailAcknowledgement"),
    ):
        match = find_tail_heading(body, label)
        if match is None:
            continue
        index, paragraph = match
        removed = remove_page_break_runs(paragraph)
        page_break_added = ensure_page_break_before(paragraph)
        outline_changed = ensure_outline_level(paragraph, "0")
        bookmark_added = ensure_bookmark(paragraph, root, bookmark_name)
        hard_break_added = False
        if label == ACKNOWLEDGEMENT_LABEL:
            hard_break_added = ensure_leading_page_break_run(paragraph)
        changes.append(
            {
                "kind": "tail_block_opener",
                "label": label,
                "index": index,
                "removed_hard_page_breaks": removed,
                "page_break_before_added": page_break_added,
                "hard_page_break_paragraph_added": hard_break_added,
                "outline_level_0_changed": outline_changed,
                "bookmark_added": bookmark_added,
                "bookmark": bookmark_name,
            }
        )
    sectpr_move = move_body_level_sectpr_to_end(body)
    if sectpr_move:
        changes.append(sectpr_move)
    return changes


def move_body_level_sectpr_to_end(body: ET.Element) -> dict[str, object] | None:
    children = list(body)
    sects = [(index, child) for index, child in enumerate(children) if child.tag == qn("sectPr")]
    if not sects:
        return None
    index, sect = sects[-1]
    if index == len(children) - 1:
        return None
    body.remove(sect)
    body.append(sect)
    return {
        "kind": "body_level_sectpr_moved_to_end",
        "from_index": index,
        "to_index": len(children) - 1,
    }


def ensure_leading_page_break_run(paragraph: ET.Element) -> bool:
    first_non_ppr = next((child for child in list(paragraph) if child.tag != qn("pPr")), None)
    if first_non_ppr is not None and first_non_ppr.tag == qn("r") and first_non_ppr.find("./w:br", NS) is not None:
        return False
    run = ET.Element(qn("r"))
    br = ET.SubElement(run, qn("br"))
    br.set(qn("type"), "page")
    ppr = paragraph_ppr(paragraph)
    insert_at = 1 if ppr is not None and list(paragraph) and list(paragraph)[0] is ppr else 0
    paragraph.insert(insert_at, run)
    return True


def ensure_hard_page_break_paragraph_before(body: ET.Element, paragraph: ET.Element) -> bool:
    children = list(body)
    index = children.index(paragraph)
    if index > 0:
        previous = children[index - 1]
        if paragraph_has_only_page_break_payload(previous):
            return False
    page_break_paragraph = ET.Element(qn("p"))
    run = ET.SubElement(page_break_paragraph, qn("r"))
    br = ET.SubElement(run, qn("br"))
    br.set(qn("type"), "page")
    body.insert(index, page_break_paragraph)
    return True


def paragraph_has_only_page_break_payload(paragraph: ET.Element) -> bool:
    if paragraph.tag != qn("p"):
        return False
    if paragraph_text(paragraph).strip():
        return False
    if paragraph.find("./w:pPr/w:sectPr", NS) is not None:
        return False
    page_break_count = 0
    for child in list(paragraph):
        if child.tag == qn("pPr"):
            continue
        if child.tag != qn("r"):
            return False
        for run_child in list(child):
            if run_child.tag == qn("rPr"):
                continue
            if run_child.tag == qn("br") and (run_child.get(qn("type")) or "page") == "page":
                page_break_count += 1
                continue
            return False
    return page_break_count > 0


def paragraph_has_true_page_break_before(paragraph: ET.Element) -> bool:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return False
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        return False
    return (page_break.get(qn("val")) or "true") not in {"0", "false", "False", "off", "OFF"}


def remove_redundant_tail_pagebreaks(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    changes.extend(merge_duplicate_tail_visual_titles(body))
    changes.extend(move_tail_blank_section_owners_to_previous_content(body))
    for label in TAIL_LABELS:
        match = find_tail_heading(body, label)
        if match is None:
            continue
        heading_index, heading = match
        removed_heading_breaks = 0
        previous_section_owns_open = False
        if heading_index > 0:
            previous = list(body)[heading_index - 1]
            previous_section_owns_open = previous.tag == qn("p") and has_section_break(previous)
        if paragraph_has_true_page_break_before(heading) or previous_section_owns_open:
            removed_heading_breaks = remove_page_break_runs(heading)
            removed_indices: list[int] = []
            scan = heading_index - 1
            while scan >= 0:
                children = list(body)
                candidate = children[scan]
                if not paragraph_has_only_page_break_payload(candidate):
                    break
                body.remove(candidate)
                removed_indices.append(scan)
                scan -= 1
            if removed_indices or removed_heading_breaks:
                changes.append(
                    {
                        "kind": "tail_redundant_pagebreaks",
                        "label": label,
                        "heading_index_before": heading_index,
                        "removed_preceding_blank_pagebreak_paragraph_indices": removed_indices,
                        "removed_page_break_runs_on_heading": removed_heading_breaks,
                        "previous_section_owns_open": previous_section_owns_open,
                        "reason": (
                            "tail heading already owns pagination through pageBreakBefore"
                            if paragraph_has_true_page_break_before(heading)
                            else "previous section break already owns tail heading opener pagination"
                        ),
                    }
                )
    return changes


def direct_bookmark_nodes(paragraph: ET.Element, local: str) -> list[ET.Element]:
    return [child for child in list(paragraph) if child.tag == qn(local)]


def move_direct_bookmark_ranges(source: ET.Element, target: ET.Element) -> int:
    starts = direct_bookmark_nodes(source, "bookmarkStart")
    ends = direct_bookmark_nodes(source, "bookmarkEnd")
    if not starts and not ends:
        return 0
    for node in starts + ends:
        source.remove(node)
    ppr = paragraph_ppr(target)
    insert_at = 1 if ppr is not None and list(target) and list(target)[0] is ppr else 0
    for node in starts:
        target.insert(insert_at, node)
        insert_at += 1
    for node in ends:
        target.append(node)
    return len(starts) + len(ends)


def move_page_break_before(source: ET.Element, target: ET.Element) -> bool:
    source_ppr = paragraph_ppr(source)
    if source_ppr is None:
        return False
    page_break = source_ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        return False
    target_ppr = ensure_ppr(target)
    source_ppr.remove(page_break)
    if target_ppr.find("./w:pageBreakBefore", NS) is not None:
        return True
    style = target_ppr.find("./w:pStyle", NS)
    insert_at = list(target_ppr).index(style) + 1 if style is not None else 0
    target_ppr.insert(insert_at, page_break)
    return True


def merge_duplicate_tail_visual_titles(body: ET.Element) -> list[dict[str, object]]:
    """Move TOC/bookmark anchors from duplicate tail headings to visual titles."""
    changes: list[dict[str, object]] = []
    index = 0
    while index < len(list(body)) - 1:
        children = list(body)
        first = children[index]
        second = children[index + 1]
        if first.tag != qn("p") or second.tag != qn("p"):
            index += 1
            continue
        first_label = exact_tail_label(paragraph_text(first))
        second_label = exact_tail_label(paragraph_text(second))
        if first_label is None or first_label != second_label:
            index += 1
            continue
        if not direct_bookmark_nodes(first, "bookmarkStart") and paragraph_style_id(first) != "1":
            index += 1
            continue
        moved_bookmarks = move_direct_bookmark_ranges(first, second)
        moved_page_break_before = move_page_break_before(first, second)
        outline_changed = ensure_outline_level(second, "0")
        body.remove(first)
        changes.append(
            {
                "kind": "tail_duplicate_visual_title_merged",
                "label": first_label,
                "removed_logical_title_index": index,
                "kept_visual_title_index_after": index,
                "moved_bookmark_nodes": moved_bookmarks,
                "moved_page_break_before": moved_page_break_before,
                "outline_level_0_changed": outline_changed,
                "reason": "tail TOC/bookmark owner duplicated the visible template title",
            }
        )
    return changes


def previous_nonempty_paragraph_index(children: list[ET.Element], before_index: int) -> int | None:
    for index in range(before_index - 1, -1, -1):
        child = children[index]
        if child.tag == qn("p") and paragraph_text(child).strip():
            return index
    return None


def move_tail_blank_section_owners_to_previous_content(body: ET.Element) -> list[dict[str, object]]:
    """Move a blank section-break owner before a tail heading onto real content."""
    changes: list[dict[str, object]] = []
    for label in TAIL_LABELS:
        match = find_tail_heading(body, label)
        if match is None:
            continue
        heading_index, _heading = match
        children = list(body)
        if heading_index <= 0:
            continue
        candidate_index: int | None = None
        for scan in range(heading_index - 1, -1, -1):
            candidate = children[scan]
            if candidate.tag != qn("p"):
                continue
            if paragraph_text(candidate).strip():
                break
            if is_blank_paragraph_without_payload(candidate) and has_section_break(candidate):
                candidate_index = scan
                break
        if candidate_index is None:
            continue
        candidate = children[candidate_index]
        previous_index = previous_nonempty_paragraph_index(children, candidate_index)
        if previous_index is None:
            changes.append(
                {
                    "kind": "tail_blank_section_owner_migration",
                    "label": label,
                    "changed": False,
                    "blank_section_index": candidate_index,
                    "reason": "no previous non-empty paragraph found",
                }
            )
            continue
        previous = children[previous_index]
        if has_section_break(previous):
            changes.append(
                {
                    "kind": "tail_blank_section_owner_migration",
                    "label": label,
                    "changed": False,
                    "blank_section_index": candidate_index,
                    "previous_content_index": previous_index,
                    "reason": "previous content paragraph already owns a section break",
                }
            )
            continue
        sect = take_paragraph_section_break(candidate)
        if sect is None:
            continue
        replace_paragraph_section_break(previous, sect)
        removed_blank_count = 0
        for remove_index in range(heading_index - 1, previous_index, -1):
            current_children = list(body)
            if remove_index >= len(current_children):
                continue
            blank = current_children[remove_index]
            if blank.tag == qn("p") and is_blank_paragraph_without_payload(blank) and not has_section_break(blank):
                body.remove(blank)
                removed_blank_count += 1
        changes.append(
            {
                "kind": "tail_blank_section_owner_migration",
                "label": label,
                "changed": True,
                "from_blank_index": candidate_index,
                "to_previous_content_index": previous_index,
                "removed_blank_count": removed_blank_count,
                "reason": "blank section-break owner moved to previous real tail content",
            }
        )
    return changes


def toc_entry_visible_label(text: str) -> str:
    value = re.sub(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", "", text or "", flags=re.IGNORECASE)
    value = re.sub(r"[\t\u2026·•]+", " ", value)
    value = re.sub(r"[.\uff0e．]{2,}", " ", value)
    return value.strip()


def toc_entry_labels(body: ET.Element, toc_index: int, body_start: int) -> set[str]:
    labels: set[str] = set()
    for child in list(body)[toc_index + 1 : body_start]:
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        labels.add(compact_text(toc_entry_visible_label(paragraph_text(child))))
    return labels


def toc_level1_entries(body: ET.Element, toc_index: int, body_start: int) -> list[ET.Element]:
    return [
        child
        for child in list(body)[toc_index + 1 : body_start]
        if child.tag == qn("p") and paragraph_style_id(child).upper() in {"TOC1", "1"} and is_toc_entry_paragraph(child)
    ]


def find_toc_entry_by_label(body: ET.Element, toc_index: int, body_start: int, label: str) -> ET.Element | None:
    target = compact_text(label)
    for child in list(body)[toc_index + 1 : body_start]:
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        if compact_text(toc_entry_visible_label(paragraph_text(child))) == target:
            return child
    return None


def remove_last_field_end(paragraph: ET.Element) -> bool:
    matches: list[tuple[ET.Element, ET.Element]] = []
    for run in paragraph.findall("./w:r", NS):
        for fld in run.findall("./w:fldChar", NS):
            if fld.get(qn("fldCharType")) == "end":
                matches.append((run, fld))
    if not matches:
        return False
    run, fld = matches[-1]
    run.remove(fld)
    if len(list(run)) == 0:
        paragraph.remove(run)
    return True


def append_main_toc_field_end(paragraph: ET.Element) -> None:
    run = ET.SubElement(paragraph, qn("r"))
    fld = ET.SubElement(run, qn("fldChar"))
    fld.set(qn("fldCharType"), "end")


def replace_toc_bookmark_refs(paragraph: ET.Element, bookmark_name: str) -> None:
    for node in paragraph.findall(".//w:instrText", NS):
        node.text = re.sub(r"_Toc[A-Za-z0-9_]+", bookmark_name, node.text or "")


def set_toc_entry_text_and_page(paragraph: ET.Element, label: str, page: str) -> None:
    title_written = False
    page_written = False
    seen_tab = False
    for run in paragraph.findall("./w:r", NS):
        for child in list(run):
            if child.tag == qn("tab"):
                seen_tab = True
                continue
            if child.tag != qn("t"):
                continue
            value = child.text or ""
            if not value.strip():
                continue
            if not seen_tab:
                child.text = label if not title_written else ""
                title_written = True
            else:
                child.text = page if not page_written else ""
                page_written = True
    if not title_written:
        run = ET.SubElement(paragraph, qn("r"))
        text_node = ET.SubElement(run, qn("t"))
        text_node.text = label
    if not page_written:
        run = ET.SubElement(paragraph, qn("r"))
        text_node = ET.SubElement(run, qn("t"))
        text_node.text = page


def write_plain_toc_entry_text(paragraph: ET.Element, label: str, page: str) -> None:
    target = f"{label}{page}"
    starts, ends = review_anchor_children(paragraph)
    first_visible_rpr = None
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run).strip():
            rpr = run.find("./w:rPr", NS)
            first_visible_rpr = deepcopy(rpr) if rpr is not None else None
            break
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    run = ET.SubElement(paragraph, qn("r"))
    if first_visible_rpr is not None:
        run.append(first_visible_rpr)
    text_node = ET.SubElement(run, qn("t"))
    text_node.text = target
    for child in ends:
        paragraph.append(child)


def set_toc_entry_label_preserving_form(paragraph: ET.Element, label: str, page: str) -> None:
    if has_tab(paragraph):
        set_toc_entry_text_and_page(paragraph, label, page)
    else:
        write_plain_toc_entry_text(paragraph, label, page)
    ensure_toc_right_dotted_tab(paragraph)


def ensure_toc_right_dotted_tab(paragraph: ET.Element, *, pos: str = "8500") -> bool:
    ppr = ensure_ppr(paragraph)
    tabs = ppr.find("./w:tabs", NS)
    if tabs is None:
        tabs = ET.Element(qn("tabs"))
        ppr.append(tabs)
    for tab in tabs.findall("./w:tab", NS):
        if tab.get(qn("val")) in {"right", "end"}:
            changed = False
            if not tab.get(qn("pos")):
                tab.set(qn("pos"), pos)
                changed = True
            if tab.get(qn("leader")) != "dot":
                tab.set(qn("leader"), "dot")
                changed = True
            return changed
    tab = ET.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    tab.set(qn("pos"), pos)
    return True


def ensure_toc_entry_dotted_tabs(body: ET.Element) -> list[dict[str, object]]:
    toc_index, body_start = find_toc_and_body_start(body)
    changes: list[dict[str, object]] = []
    for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1):
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        if not paragraph_text(child).strip():
            continue
        if ensure_toc_right_dotted_tab(child):
            changes.append({"kind": "toc_right_dotted_tab", "body_child_index": index, "text": paragraph_text(child)[:100]})
    return changes


def toc_entry_cached_page(paragraph: ET.Element) -> str:
    visible = paragraph_text(paragraph).strip()
    parsed = static_toc_label_and_page(visible)
    if parsed is not None:
        return parsed[1]
    runs = paragraph.findall("./w:r", NS)
    seen_tab = False
    for run in runs:
        if run.find("./w:tab", NS) is not None:
            seen_tab = True
        if not seen_tab:
            continue
        for text_node in run.findall("./w:t", NS):
            value = (text_node.text or "").strip()
            if value:
                return value
    return "0"


def append_missing_tail_toc_entries(body: ET.Element, root: ET.Element) -> list[dict[str, object]]:
    toc_index, body_start = find_toc_and_body_start(body)
    existing_labels = toc_entry_labels(body, toc_index, body_start)
    required: list[tuple[str, str]] = []
    if find_tail_heading(body, REFERENCES_LABEL) is not None:
        required.append((REFERENCES_LABEL, "_TocTailReferences"))
    if find_tail_heading(body, ACKNOWLEDGEMENT_LABEL) is not None:
        required.append((ACKNOWLEDGEMENT_LABEL, "_TocTailAcknowledgement"))
    missing = [
        (label, bookmark)
        for label, bookmark in required
        if compact_text(label) not in existing_labels
    ]
    if not missing:
        return []
    donors = toc_level1_entries(body, toc_index, body_start)
    if not donors:
        raise RuntimeError("toc-tail-entries requires a level-1 TOC donor paragraph")
    donor = donors[-1]
    insertion_anchor = donor
    if any(label == ACKNOWLEDGEMENT_LABEL for label, _bookmark in missing):
        references_entry = find_toc_entry_by_label(body, toc_index, body_start, REFERENCES_LABEL)
        if references_entry is not None:
            donor = references_entry
            insertion_anchor = references_entry
    main_end_removed = remove_last_field_end(donor)
    current_children = list(body)
    insert_at = current_children.index(insertion_anchor) + 1
    changes: list[dict[str, object]] = []
    appended: list[ET.Element] = []
    for offset, (label, bookmark_name) in enumerate(missing):
        heading_match = find_tail_heading(body, label)
        if heading_match is None:
            continue
        ensure_bookmark(heading_match[1], root, bookmark_name)
        clone = deepcopy(donor)
        replace_toc_bookmark_refs(clone, bookmark_name)
        set_toc_entry_text_and_page(clone, label, "0")
        body.insert(insert_at + offset, clone)
        appended.append(clone)
        changes.append(
            {
                "kind": "toc_tail_entry",
                "label": label,
                "bookmark": bookmark_name,
                "placeholder_page": "0",
                "inserted_after": paragraph_text(donor)[:120],
                "field_end_transferred": main_end_removed,
            }
        )
    if appended and main_end_removed:
        append_main_toc_field_end(appended[-1])
    return changes


FINAL_CHAPTER_TITLE = "\u7b2c6\u7ae0  \u7ed3\u8bba\u4e0e\u5c55\u671b"
FINAL_CHAPTER_SHORT_LABELS = {
    compact_text("\u7ed3\u8bba"),
    compact_text("\u7ed3\u8bba\u4e0e\u5c55\u671b"),
    compact_text("\u603b\u7ed3\u4e0e\u5c55\u671b"),
}


def is_final_chapter_label(text: str) -> bool:
    compact = compact_text(text)
    if compact in {compact_text(FINAL_CHAPTER_TITLE), *FINAL_CHAPTER_SHORT_LABELS}:
        return True
    if not re.match(r"^\u7b2c\s*6\s*\u7ae0", text or ""):
        return False
    return any(token in compact for token in ("\u7ed3\u8bba", "\u5c55\u671b", "\u6539\u8fdb\u65b9\u5411", "\u8bbe\u8ba1\u7ed3\u8bba"))


def clear_text_nodes_after_first(paragraph: ET.Element) -> None:
    seen = False
    for text_node in paragraph.findall(".//w:t", NS):
        if not seen:
            seen = True
            continue
        text_node.text = ""


def set_paragraph_visible_text_preserving_first_run(paragraph: ET.Element, value: str) -> None:
    for text_node in paragraph.findall(".//w:t", NS):
        if (text_node.text or "").strip():
            text_node.text = value
            clear_text_nodes_after_first(paragraph)
            return
    run = ET.SubElement(paragraph, qn("r"))
    text_node = ET.SubElement(run, qn("t"))
    text_node.text = value


def find_final_chapter_heading_candidate(body: ET.Element) -> tuple[int, ET.Element] | None:
    children = list(body)
    reference_index = next(
        (
            index
            for index, child in enumerate(children)
            if child.tag == qn("p") and compact_text(paragraph_text(child).strip()) == compact_text(REFERENCES_LABEL)
        ),
        len(children),
    )
    for index in range(reference_index - 1, -1, -1):
        child = children[index]
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if not text:
            continue
        compact = compact_text(text)
        if is_final_chapter_label(text):
            return index, child
        if re.fullmatch(r"6\s*\.\s*1\s+.+", text):
            previous = previous_nonempty_body_paragraph(body, index)
            if previous is not None:
                prev_index, prev_para = previous
                prev_text = paragraph_text(prev_para).strip()
                if is_final_chapter_label(prev_text):
                    return prev_index, prev_para
    return None


def previous_nonempty_body_paragraph(body: ET.Element, before_index: int) -> tuple[int, ET.Element] | None:
    children = list(body)
    for index in range(before_index - 1, -1, -1):
        child = children[index]
        if child.tag == qn("p") and paragraph_text(child).strip():
            return index, child
    return None


def first_existing_chapter_heading_donor(body: ET.Element) -> tuple[ET.Element | None, ET.Element | None, str]:
    try:
        _toc_index, body_start = find_toc_and_body_start(body)
    except RuntimeError:
        body_start = 0
    for child in list(body)[body_start:]:
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if (
            re.match(r"^\u7b2c\s*\d+\s*\u7ae0", text)
            and compact_text(text) != compact_text(FINAL_CHAPTER_TITLE)
            and not is_toc_entry_paragraph(child)
        ):
            donor_rpr = None
            for run in child.findall("./w:r", NS):
                if paragraph_text(run).strip():
                    donor_rpr = run.find("./w:rPr", NS)
                    break
            return paragraph_ppr(child), donor_rpr, paragraph_style_id(child)
    raise RuntimeError("final-chapter-heading could not locate an existing numbered chapter heading donor")


def replace_visible_run_rpr_from_donor(paragraph: ET.Element, donor_rpr: ET.Element | None) -> None:
    if donor_rpr is None:
        return
    for run in paragraph.findall("./w:r", NS):
        if not paragraph_text(run).strip():
            continue
        old_rpr = run.find("./w:rPr", NS)
        if old_rpr is not None:
            run.remove(old_rpr)
        run.insert(0, deepcopy(donor_rpr))


def repair_final_chapter_heading(body: ET.Element) -> list[dict[str, object]]:
    match = find_final_chapter_heading_candidate(body)
    if match is None:
        raise RuntimeError("final-chapter-heading could not locate the final chapter heading candidate before references")
    index, paragraph = match
    old_text = paragraph_text(paragraph).strip()
    old_style = paragraph_style_id(paragraph)
    donor_ppr, donor_rpr, donor_style = first_existing_chapter_heading_donor(body)
    replace_ppr(paragraph, ppr_without_style(donor_ppr), donor_style or None)
    replace_visible_run_rpr_from_donor(paragraph, donor_rpr)
    set_paragraph_visible_text_preserving_first_run(paragraph, FINAL_CHAPTER_TITLE)
    ppr = ensure_ppr(paragraph)
    outline = ppr.find("./w:outlineLvl", NS)
    if outline is None:
        outline = ET.Element(qn("outlineLvl"))
        ppr.append(outline)
    outline.set(qn("val"), "0")
    return [
        {
            "kind": "final_chapter_heading",
            "body_child_index": index,
            "old_text": old_text,
            "new_text": FINAL_CHAPTER_TITLE,
            "old_style_id": old_style,
            "new_style_id": paragraph_style_id(paragraph),
            "donor_style_id": donor_style,
        }
    ]


def first_toc_body_entry(body: ET.Element, toc_index: int, body_start: int) -> tuple[int, ET.Element] | None:
    for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1):
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        text = paragraph_text(child).strip()
        label = toc_entry_visible_label(text)
        if heading_level(label) == 1:
            return index, child
    return None


def repair_static_toc_frontmatter_and_final_chapter(body: ET.Element) -> list[dict[str, object]]:
    toc_index, body_start = find_toc_and_body_start(body)
    entries = [
        (index, child)
        for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1)
        if child.tag == qn("p") and is_toc_entry_paragraph(child)
    ]
    if not entries:
        raise RuntimeError("static-toc-frontmatter-final-chapter could not locate TOC entries")
    changes: list[dict[str, object]] = []
    self_entry_indices = [
        index
        for index, child in entries
        if is_toc_heading(toc_entry_visible_label(paragraph_text(child).strip()))
    ]
    if self_entry_indices:
        donor = next((row for row in entries if row[0] not in self_entry_indices), None)
        if donor is None:
            raise RuntimeError("static-toc-frontmatter-final-chapter could not locate donor after TOC self entry")
        host_index = self_entry_indices[0]
        host = list(body)[host_index]
        donor_index, donor_paragraph = donor
        moved_prefix_count = rewrite_toc_host_from_donor_preserving_field_prefix(host, donor_paragraph)
        remove_indices = sorted(set(self_entry_indices[1:] + [donor_index]), reverse=True)
        removed_rows: list[dict[str, object]] = []
        children = list(body)
        for index in remove_indices:
            paragraph = children[index]
            removed_rows.append({"body_child_index": index, "text": paragraph_text(paragraph).strip()})
            body.remove(paragraph)
        changes.append(
            {
                "kind": "toc_self_entry_exclusion",
                "host_body_child_index": host_index,
                "donor_body_child_index": donor_index,
                "field_prefix_run_count_preserved": moved_prefix_count,
                "removed_rows": list(reversed(removed_rows)),
                "reason": "cached TOC must start from the Chinese abstract, not include the TOC title itself",
            }
        )
        toc_index, body_start = find_toc_and_body_start(body)
        entries = [
            (index, child)
            for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1)
            if child.tag == qn("p") and is_toc_entry_paragraph(child)
        ]
    existing_labels = toc_entry_labels(body, toc_index, body_start)
    first_body_entry = first_toc_body_entry(body, toc_index, body_start)
    if first_body_entry is None:
        raise RuntimeError("static-toc-frontmatter-final-chapter could not locate first body TOC donor")
    first_body_index, first_body_paragraph = first_body_entry
    front_required = [
        ("\u6458\u8981", "I"),
        ("ABSTRACT", "II"),
    ]
    insert_at = first_body_index
    inserted_count = 0
    for label, page in front_required:
        if compact_text(label) in existing_labels:
            continue
        clone = deepcopy(first_body_paragraph)
        set_toc_entry_label_preserving_form(clone, label, page)
        body.insert(insert_at + inserted_count, clone)
        inserted_count += 1
        changes.append(
            {
                "kind": "toc_frontmatter_entry",
                "label": label,
                "placeholder_page": page,
                "inserted_before": paragraph_text(first_body_paragraph)[:80],
            }
        )
    toc_index, body_start = find_toc_and_body_start(body)
    for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1):
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        visible_text = paragraph_text(child).strip()
        label = toc_entry_visible_label(visible_text)
        if is_final_chapter_label(label):
            page = toc_entry_cached_page(child)
            old_text = visible_text
            set_toc_entry_label_preserving_form(child, FINAL_CHAPTER_TITLE, page)
            changes.append(
                {
                    "kind": "toc_final_chapter_entry",
                    "body_child_index": index,
                    "old_text": old_text,
                    "new_label": FINAL_CHAPTER_TITLE,
                    "page": page,
                }
            )
            break
    return changes


def repair_frontmatter_title_style_bindings(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    changes = ensure_abstract_title_style(styles_root)
    toc_match = find_first_paragraph(body, lambda p: is_toc_heading(paragraph_text(p).strip()))
    if toc_match is None:
        raise RuntimeError("frontmatter-title-style-bindings could not locate TOC boundary")
    toc_index, _toc_paragraph = toc_match
    frontmatter_children = list(body)[:toc_index]
    for label, matcher, english in (
        ("zh_abstract_title", is_zh_abstract_label, False),
        ("en_abstract_title", is_en_abstract_label, True),
    ):
        match = next(
            (
                (index, child)
                for index, child in enumerate(frontmatter_children)
                if child.tag == qn("p") and is_frontmatter_title_text(paragraph_text(child), english=english)
            ),
            None,
        )
        if match is None:
            raise RuntimeError(f"frontmatter-title-style-bindings could not locate {label}")
        index, paragraph = match
        old_style = paragraph_style_id(paragraph)
        set_frontmatter_title_metrics(paragraph)
        for run in paragraph.findall("./w:r", NS):
            if not paragraph_text(run).strip():
                continue
            old_rpr = run.find("./w:rPr", NS)
            if old_rpr is not None:
                run.remove(old_rpr)
            run.insert(0, frontmatter_title_run_rpr(english=english))
        changes.append(
            {
                "kind": "frontmatter_title_style_binding",
                "surface": label,
                "body_child_index": index,
                "old_style_id": old_style,
                "new_style_id": paragraph_style_id(paragraph),
            }
        )
    return changes


def is_frontmatter_title_text(text: str, *, english: bool) -> bool:
    compact = compact_text(text).lower()
    if english:
        return compact == "abstract"
    return compact == "\u6458\u8981"


def frontmatter_title_run_rpr(*, english: bool) -> ET.Element:
    rpr = ET.Element(qn("rPr"))
    fonts = ET.SubElement(rpr, qn("rFonts"))
    if english:
        fonts.set(qn("eastAsia"), "Arial")
        fonts.set(qn("ascii"), "Arial")
        fonts.set(qn("hAnsi"), "Arial")
        fonts.set(qn("cs"), "Arial")
    else:
        fonts.set(qn("eastAsia"), "\u9ed1\u4f53")
        fonts.set(qn("ascii"), "SimHei")
        fonts.set(qn("hAnsi"), "SimHei")
        fonts.set(qn("cs"), "SimHei")
    ET.SubElement(rpr, qn("b"))
    ET.SubElement(rpr, qn("bCs"))
    size = ET.SubElement(rpr, qn("sz"))
    size.set(qn("val"), "30")
    size_cs = ET.SubElement(rpr, qn("szCs"))
    size_cs.set(qn("val"), "30")
    return rpr


def child_index(body: ET.Element, element: ET.Element) -> int:
    for index, child in enumerate(list(body)):
        if child is element:
            return index
    raise RuntimeError("element is not a child of the supplied body")


def has_section_break(element: ET.Element) -> bool:
    return element.tag == qn("p") and element.find("./w:pPr/w:sectPr", NS) is not None


def has_page_break(element: ET.Element) -> bool:
    return element.tag == qn("p") and any(
        (br.get(qn("type")) or "page") == "page"
        for br in element.findall(".//w:br", NS)
    )


def first_frontmatter_child_index(body: ET.Element, predicate) -> int:
    match = find_first_paragraph(body, predicate)
    if match is None:
        raise RuntimeError("frontmatter-template-sections could not locate a required front-matter paragraph")
    return match[0]


def previous_nonempty_child_index(body: ET.Element, *, start_exclusive: int, stop_exclusive: int) -> int:
    children = list(body)
    for index in range(start_exclusive - 1, stop_exclusive, -1):
        child = children[index]
        if child.tag == qn("p") and paragraph_text(child).strip():
            return index
    raise RuntimeError("frontmatter-template-sections could not locate the previous nonempty paragraph")


def last_nonempty_child_index_before(body: ET.Element, stop_exclusive: int) -> int:
    children = list(body)
    for index in range(stop_exclusive - 1, -1, -1):
        child = children[index]
        if child.tag == qn("p") and paragraph_text(child).strip():
            return index
    raise RuntimeError("frontmatter-template-sections could not locate a prior nonempty paragraph")


def frontmatter_topology_points(body: ET.Element) -> dict[str, int]:
    zh_keyword = first_frontmatter_child_index(
        body,
        lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    en_abstract = first_frontmatter_child_index(
        body,
        lambda p: is_en_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    try:
        en_title = previous_nonempty_child_index(body, start_exclusive=en_abstract, stop_exclusive=zh_keyword)
    except RuntimeError:
        # Some approved samples use the Abstract paragraph itself as the first
        # English-front-matter marker, without a separate English title line.
        en_title = en_abstract
    en_keyword = first_frontmatter_child_index(
        body,
        lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    try:
        toc_index, body_start = find_toc_and_body_start(body)
    except RuntimeError:
        toc_match = find_first_paragraph(body, lambda p: is_toc_heading(paragraph_text(p).strip()))
        if toc_match is None:
            raise
        toc_index = toc_match[0]
        body_start = None
        for index, child in enumerate(list(body)[toc_index + 1 :], start=toc_index + 1):
            if child.tag != qn("p"):
                continue
            if paragraph_text(child).strip() and not is_toc_entry_paragraph(child):
                body_start = index
                break
        if body_start is None:
            raise RuntimeError("frontmatter-template-sections could not locate relaxed body start after TOC")
    last_toc_nonempty = last_nonempty_child_index_before(body, body_start)
    return {
        "zh_keyword": zh_keyword,
        "en_title": en_title,
        "en_abstract": en_abstract,
        "en_keyword": en_keyword,
        "toc": toc_index,
        "body_start": body_start,
        "last_toc_nonempty": last_toc_nonempty,
    }


def section_break_on_child(body: ET.Element, index: int) -> ET.Element | None:
    children = list(body)
    if index < 0 or index >= len(children):
        return None
    child = children[index]
    if child.tag != qn("p"):
        return None
    return child.find("./w:pPr/w:sectPr", NS)


def ensure_trailing_page_break_run(paragraph: ET.Element) -> bool:
    for br in paragraph.findall(".//w:br", NS):
        if br.get(qn("type")) == "page":
            return False
    run = ET.Element(qn("r"))
    br = ET.SubElement(run, qn("br"))
    br.set(qn("type"), "page")
    paragraph.append(run)
    return True


def normalize_frontmatter_text_section_owner(
    body: ET.Element,
    *,
    text_paragraph_index: int,
    next_content_index: int,
    location: str,
) -> list[dict[str, object]]:
    """Keep a next-page section break on the previous visible line, not a blank page owner."""
    changes: list[dict[str, object]] = []
    text_para = list(body)[text_paragraph_index]
    next_para = list(body)[next_content_index]
    sect = section_break_on_child(body, text_paragraph_index)
    if sect is None:
        changes.append(
            migrate_blank_section_to_previous_content(
                body,
                previous_content_index=text_paragraph_index,
                next_content_index=next_content_index,
                location=location,
            )
        )
    else:
        changes.append(
            {
                "kind": "frontmatter_section_owner_on_text",
                "location": location,
                "changed": False,
                "text_paragraph_index": text_paragraph_index,
                "reason": "visible previous line already owns the section break",
            }
        )

    current_text_index = child_index(body, text_para)
    current_next_index = child_index(body, next_para)
    next_para_index = child_index(body, next_para)
    next_pagebreak_changed = False
    next_pagebreak_reason = "next content is not a paragraph"
    if location == "en_keyword_to_toc":
        removed_section_break = take_paragraph_section_break(text_para) is not None
        if next_para.tag == qn("p"):
            next_pagebreak_changed, next_pagebreak_reason = ensure_true_page_break_before(next_para)
        changes.append(
            {
                "kind": "frontmatter_intermediate_section_removed",
                "location": location,
                "changed": removed_section_break,
                "section_owner_index": current_text_index if removed_section_break else None,
                "next_content_index": next_para_index,
                "next_content_pageBreakBefore_changed": next_pagebreak_changed,
                "next_content_pageBreakBefore_reason": next_pagebreak_reason,
            }
        )
        removed = remove_blank_range_between(body, current_text_index, child_index(body, next_para))
        changes.append(
            {
                "kind": "frontmatter_separator_blank_cleanup",
                "location": location,
                "removed_blank_count": len(removed),
                "removed": removed,
            }
        )
        return changes

    sect = section_break_on_child(body, current_text_index)
    if next_para.tag == qn("p"):
        next_pagebreak_changed = remove_page_break_before_property(next_para)
        next_pagebreak_reason = "removed pageBreakBefore; previous keyword paragraph owns the hard page break"
    trailing_break_changed = ensure_trailing_page_break_run(text_para)
    changes.append(
        {
            "kind": "frontmatter_section_start_type",
            "location": location,
            "changed": set_section_start_type(sect, "continuous") if sect is not None else False,
            "value": "continuous",
            "section_owner_index": current_text_index if sect is not None else None,
            "next_content_index": next_para_index,
            "next_content_pageBreakBefore_changed": next_pagebreak_changed,
            "next_content_pageBreakBefore_reason": next_pagebreak_reason,
            "text_paragraph_page_break_run_changed": trailing_break_changed,
        }
    )
    removed = remove_blank_range_between(body, current_text_index, current_next_index)
    changes.append(
        {
            "kind": "frontmatter_separator_blank_cleanup",
            "location": location,
            "removed_blank_count": len(removed),
            "removed": removed,
        }
    )
    return changes


def repair_frontmatter_section_owner_migration(body: ET.Element) -> list[dict[str, object]]:
    """Normalize abstract/TOC section ownership without cloning template text."""
    changes: list[dict[str, object]] = []
    zh_title = find_first_paragraph(
        body,
        lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    if zh_title is not None:
        title_index, title_para = zh_title
        changed, reason = ensure_true_page_break_before(title_para)
        changes.append(
            {
                "kind": "frontmatter_first_abstract_pagebreak",
                "body_child_index": title_index,
                "changed": changed,
                "reason": reason,
            }
        )
    points = frontmatter_topology_points(body)
    changes.extend(
        normalize_frontmatter_text_section_owner(
            body,
            text_paragraph_index=points["zh_keyword"],
            next_content_index=points["en_abstract"],
            location="zh_keyword_to_en_abstract",
        )
    )

    points = frontmatter_topology_points(body)
    changes.extend(
        normalize_frontmatter_text_section_owner(
            body,
            text_paragraph_index=points["en_keyword"],
            next_content_index=points["toc"],
            location="en_keyword_to_toc",
        )
    )

    points = frontmatter_topology_points(body)
    toc_tail = list(body)[points["last_toc_nonempty"]]
    if toc_tail.tag == qn("p") and has_section_break(toc_tail):
        changes.append(
            migrate_text_section_to_following_blank(
                body,
                text_paragraph_index=points["last_toc_nonempty"],
                stop_exclusive=points["body_start"],
                location="toc_tail_to_body_start",
            )
        )
    else:
        changes.append(
            migrate_blank_section_to_previous_content(
                body,
                previous_content_index=points["last_toc_nonempty"],
                next_content_index=points["body_start"],
                location="toc_tail_to_body_start",
            )
        )
    return changes


def clone_template_range(template_body: ET.Element, start: int, end: int) -> list[ET.Element]:
    return [deepcopy(child) for child in list(template_body)[start:end]]


def strip_section_header_footer_references(element: ET.Element) -> int:
    removed = 0
    for sect in element.findall(".//w:sectPr", NS):
        for tag in ("headerReference", "footerReference"):
            for ref in list(sect.findall(f"./w:{tag}", NS)):
                sect.remove(ref)
                removed += 1
    return removed


def clone_template_separator_range(template_body: ET.Element, start: int, end: int) -> tuple[list[ET.Element], int]:
    clones = [
        clone
        for clone in clone_template_range(template_body, start, end)
        if clone.tag == qn("p") and is_blank_paragraph_without_payload(clone)
    ]
    stripped_reference_count = sum(strip_section_header_footer_references(clone) for clone in clones)
    return clones, stripped_reference_count


def insert_clones(body: ET.Element, insert_at: int, clones: list[ET.Element]) -> None:
    for offset, clone in enumerate(clones):
        body.insert(insert_at + offset, clone)


def remove_paragraph_section_break(paragraph: ET.Element) -> bool:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return False
    sect = ppr.find("./w:sectPr", NS)
    if sect is None:
        return False
    ppr.remove(sect)
    return True


def take_paragraph_section_break(paragraph: ET.Element) -> ET.Element | None:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return None
    sect = ppr.find("./w:sectPr", NS)
    if sect is None:
        return None
    ppr.remove(sect)
    return sect


def replace_paragraph_section_break(paragraph: ET.Element, sect: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    existing = ppr.find("./w:sectPr", NS)
    if existing is not None:
        ppr.remove(existing)
    ppr.append(sect)
    return True


def element_attrs(element: ET.Element | None, names: tuple[str, ...]) -> dict[str, str]:
    if element is None:
        return {}
    return {name: element.get(qn(name), "") for name in names if element.get(qn(name)) is not None}


def ensure_section_pg_num_type(
    sect: ET.Element,
    *,
    fmt: str | None = None,
    start: str | None = None,
    remove: bool = False,
) -> tuple[bool, dict[str, str]]:
    """Set or remove section page-number restart in schema-safe child order."""
    existing = sect.find("./w:pgNumType", NS)
    before = element_attrs(existing, ("fmt", "start", "chapStyle", "chapSep")) if existing is not None else {}
    if remove:
        if existing is None:
            return False, before
        sect.remove(existing)
        return True, before

    if existing is None:
        existing = ET.Element(qn("pgNumType"))
        desired_order = SECTION_CHILD_ORDER_INDEX["pgNumType"]
        insert_at = len(list(sect))
        for index, child in enumerate(list(sect)):
            local = child.tag.split("}", 1)[-1]
            if SECTION_CHILD_ORDER_INDEX.get(local, desired_order) > desired_order:
                insert_at = index
                break
        sect.insert(insert_at, existing)
    changed = False
    desired = {"fmt": fmt, "start": start}
    for attr, value in desired.items():
        qattr = qn(attr)
        if value is None:
            if qattr in existing.attrib:
                existing.attrib.pop(qattr, None)
                changed = True
        elif existing.get(qattr) != value:
            existing.set(qattr, value)
            changed = True
    return changed or existing is not None and not before, before


def repair_body_page_number_restart(body: ET.Element) -> list[dict[str, object]]:
    """Move decimal page restart to the body section owner and remove tail restarts.

    In WordprocessingML, a paragraph-level ``sectPr`` stores properties for the
    section that *ends* at that paragraph. Therefore the body decimal restart
    belongs on the section owner after the body/references block, not on the TOC
    tail paragraph immediately before the first body heading.
    """
    changes: list[dict[str, object]] = []
    points = frontmatter_topology_points(body)
    toc_tail = list(body)[points["last_toc_nonempty"]]
    if toc_tail.tag == qn("p") and has_section_break(toc_tail):
        sect = toc_tail.find("./w:pPr/w:sectPr", NS)
        pg_num = sect.find("./w:pgNumType", NS) if sect is not None else None
        attrs = element_attrs(pg_num, ("fmt", "start", "chapStyle", "chapSep")) if pg_num is not None else {}
        if attrs.get("fmt") == "decimal" and attrs.get("start") == "1":
            changed, before = ensure_section_pg_num_type(sect, fmt="lowerRoman", start="1")
            changes.append(
                {
                    "kind": "toc_tail_frontmatter_page_number_restored",
                    "changed": changed,
                    "body_child_index": points["last_toc_nonempty"],
                    "text": paragraph_text(toc_tail)[:120],
                    "before_pgNumType": before,
                    "after_pgNumType": element_attrs(
                        sect.find("./w:pgNumType", NS),
                        ("fmt", "start", "chapStyle", "chapSep"),
                    ),
                    "reason": "TOC/front-matter section must not own the body decimal restart",
                }
            )

    body_start = points["body_start"]
    acknowledgement_match = find_tail_heading(body, ACKNOWLEDGEMENT_LABEL)
    tail_start = acknowledgement_match[0] if acknowledgement_match is not None else len(list(body))
    body_section_owner: tuple[int, ET.Element] | None = None
    for index, child in enumerate(list(body)):
        if index <= body_start or index >= tail_start or child.tag != qn("p") or not has_section_break(child):
            continue
        body_section_owner = (index, child)
    if body_section_owner is None:
        raise RuntimeError("body-page-number-restart could not locate body/references section owner before acknowledgement")
    owner_index, owner = body_section_owner
    sect = owner.find("./w:pPr/w:sectPr", NS)
    if sect is None:
        raise RuntimeError("body-page-number-restart body section owner has no sectPr")
    changed, before = ensure_section_pg_num_type(sect, fmt="decimal", start="1")
    changes.append(
        {
            "kind": "body_page_number_restart_owner",
            "changed": changed,
            "body_child_index": owner_index,
            "text": paragraph_text(owner)[:120],
            "before_pgNumType": before,
            "after_pgNumType": element_attrs(
                sect.find("./w:pgNumType", NS),
                ("fmt", "start", "chapStyle", "chapSep"),
            ),
            "reason": "the body/references section owner must restart visible decimal numbering at 1",
        }
    )

    for index, child in enumerate(list(body)):
        if index <= owner_index or child.tag != qn("p") or not has_section_break(child):
            continue
        text = paragraph_text(child).strip()
        sect = child.find("./w:pPr/w:sectPr", NS)
        if sect is None:
            continue
        pg_num = sect.find("./w:pgNumType", NS)
        if pg_num is None:
            continue
        attrs = element_attrs(pg_num, ("fmt", "start", "chapStyle", "chapSep"))
        if attrs.get("fmt") == "decimal" and attrs.get("start") == "1":
            removed, before_tail = ensure_section_pg_num_type(sect, remove=True)
            changes.append(
                {
                    "kind": "tail_page_number_restart_removed",
                    "changed": removed,
                    "body_child_index": index,
                    "text": text[:120],
                    "before_pgNumType": before_tail,
                    "reason": "tail sections must continue body decimal numbering instead of restarting at 1",
                }
            )

    body_sect = body.find("./w:sectPr", NS)
    if body_sect is not None:
        pg_num = body_sect.find("./w:pgNumType", NS)
        attrs = element_attrs(pg_num, ("fmt", "start", "chapStyle", "chapSep")) if pg_num is not None else {}
        if attrs.get("fmt") == "decimal" and attrs.get("start") == "1":
            removed, before_body = ensure_section_pg_num_type(body_sect, remove=True)
            changes.append(
                {
                    "kind": "body_level_tail_page_number_restart_removed",
                    "changed": removed,
                    "before_pgNumType": before_body,
                    "reason": "final body-level section must not restart acknowledgement/tail page numbering",
                }
            )
    return changes


def set_section_start_type(sect: ET.Element, value: str) -> bool:
    sect_type = sect.find("./w:type", NS)
    if sect_type is None:
        sect_type = ET.Element(qn("type"))
        insert_at = 0
        for index, child in enumerate(list(sect)):
            if child.tag in {qn("pgSz"), qn("pgMar"), qn("pgBorders"), qn("lnNumType")}:
                insert_at = index
                break
            insert_at = index + 1
        sect.insert(insert_at, sect_type)
    if sect_type.get(qn("val")) == value:
        return False
    sect_type.set(qn("val"), value)
    return True


SECTION_GEOMETRY_TAGS = ("pgSz", "pgMar", "cols")
SECTION_CHILD_ORDER = (
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
)
SECTION_CHILD_ORDER_INDEX = {tag: index for index, tag in enumerate(SECTION_CHILD_ORDER)}


def section_geometry_nodes(document_root: ET.Element) -> list[tuple[str, int, ET.Element]]:
    body = body_element(document_root)
    nodes: list[tuple[str, int, ET.Element]] = []
    paragraphs = body_paragraphs(body)
    for index, paragraph in enumerate(paragraphs, start=1):
        sect = paragraph.find("./w:pPr/w:sectPr", NS)
        if sect is not None:
            nodes.append(("paragraph", index, sect))
    body_sect = body.find("./w:sectPr", NS)
    if body_sect is not None:
        nodes.append(("body", len(paragraphs) + 1, body_sect))
    return nodes


def section_geometry_signature(sect: ET.Element) -> tuple[bytes, bytes, bytes]:
    return tuple(
        ET.tostring(sect.find(f"./w:{tag}", NS), encoding="utf-8")
        if sect.find(f"./w:{tag}", NS) is not None
        else b""
        for tag in SECTION_GEOMETRY_TAGS
    )


def dominant_template_section_geometry(template_document_root: ET.Element | None) -> dict[str, ET.Element]:
    if template_document_root is None:
        raise RuntimeError("template-page-geometry requires --template-docx")
    counts: dict[tuple[bytes, bytes, bytes], int] = {}
    donors: dict[tuple[bytes, bytes, bytes], ET.Element] = {}
    for _owner, _index, sect in section_geometry_nodes(template_document_root):
        signature = section_geometry_signature(sect)
        if not any(signature):
            continue
        counts[signature] = counts.get(signature, 0) + 1
        donors.setdefault(signature, sect)
    if not counts:
        raise RuntimeError("template-page-geometry could not find template section geometry")
    selected_signature = max(counts, key=lambda item: counts[item])
    selected = donors[selected_signature]
    result: dict[str, ET.Element] = {}
    for tag in SECTION_GEOMETRY_TAGS:
        child = selected.find(f"./w:{tag}", NS)
        if child is not None:
            result[tag] = deepcopy(child)
    if not {"pgSz", "pgMar"}.issubset(result):
        raise RuntimeError("template-page-geometry template donor lacks pgSz or pgMar")
    return result


def replace_section_child_preserving_order(sect: ET.Element, tag: str, donor: ET.Element) -> bool:
    existing = sect.find(f"./w:{tag}", NS)
    donor_clone = deepcopy(donor)
    if existing is not None:
        old_xml = ET.tostring(existing, encoding="utf-8")
        new_xml = ET.tostring(donor_clone, encoding="utf-8")
        if old_xml == new_xml:
            return False
        index = list(sect).index(existing)
        sect.remove(existing)
        sect.insert(index, donor_clone)
        return True

    donor_order = SECTION_CHILD_ORDER_INDEX.get(tag, len(SECTION_CHILD_ORDER))
    insert_at = len(list(sect))
    for index, child in enumerate(list(sect)):
        local_name = child.tag.split("}", 1)[-1]
        if SECTION_CHILD_ORDER_INDEX.get(local_name, len(SECTION_CHILD_ORDER)) > donor_order:
            insert_at = index
            break
    sect.insert(insert_at, donor_clone)
    return True


def repair_template_page_geometry(
    document_root: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    """Replay template page geometry without touching numbering or header/footer refs."""
    donor_geometry = dominant_template_section_geometry(template_document_root)
    changes: list[dict[str, object]] = []
    for owner, index, sect in section_geometry_nodes(document_root):
        changed_tags: list[str] = []
        before = {
            tag: ET.tostring(sect.find(f"./w:{tag}", NS), encoding="unicode")
            if sect.find(f"./w:{tag}", NS) is not None
            else ""
            for tag in SECTION_GEOMETRY_TAGS
        }
        for tag, donor in donor_geometry.items():
            if replace_section_child_preserving_order(sect, tag, donor):
                changed_tags.append(tag)
        if changed_tags:
            after = {
                tag: ET.tostring(sect.find(f"./w:{tag}", NS), encoding="unicode")
                if sect.find(f"./w:{tag}", NS) is not None
                else ""
                for tag in SECTION_GEOMETRY_TAGS
            }
            changes.append(
                {
                    "kind": "template_page_geometry",
                    "owner": owner,
                    "section_index": index,
                    "changed_tags": changed_tags,
                    "before": before,
                    "after": after,
                }
            )
    if not changes:
        changes.append({"kind": "template_page_geometry", "status": "already_matches_template_donor"})
    return changes


def style_id_by_name(styles_root: ET.Element, target_name: str) -> str | None:
    target = target_name.strip().lower()
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        if style_name(style).strip().lower() == target:
            return style.get(qn("styleId")) or None
    return None


def replace_or_append_style_definition(styles_root: ET.Element, donor_style: ET.Element) -> str:
    style_id = donor_style.get(qn("styleId")) or ""
    if not style_id:
        raise RuntimeError("template style has no styleId")
    existing = style_by_id(styles_root, style_id)
    donor_clone = deepcopy(donor_style)
    if existing is None:
        styles_root.append(donor_clone)
        return "copied"
    if style_xml_equal(existing, donor_style):
        return "already_matches"
    parent_styles = list(styles_root)
    replace_at = parent_styles.index(existing)
    styles_root.remove(existing)
    styles_root.insert(replace_at, donor_clone)
    return "replaced"


def first_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is not None:
            return deepcopy(rpr)
    return None


def first_paragraph_with_style(root: ET.Element, style_id: str) -> ET.Element | None:
    for paragraph in root.findall(".//w:p", NS):
        if paragraph_style_id(paragraph) == style_id:
            return paragraph
    return None


def set_run_rpr(run: ET.Element, donor_rpr: ET.Element | None) -> bool:
    existing = run.find("./w:rPr", NS)
    if donor_rpr is None:
        return False
    donor_clone = deepcopy(donor_rpr)
    if existing is not None and ET.tostring(existing, encoding="utf-8") == ET.tostring(donor_clone, encoding="utf-8"):
        return False
    if existing is not None:
        run.remove(existing)
    run.insert(0, donor_clone)
    return True


def repair_header_footer_template_styles(
    parts: dict[str, bytes],
    styles_root: ET.Element,
    template_styles_root: ET.Element | None,
    template_parts: dict[str, bytes] | None,
) -> tuple[list[dict[str, object]], dict[str, bytes]]:
    """Bind header/footer paragraphs and PAGE runs to template-owned styles.

    This is intentionally narrower than section-topology repair: it does not
    alter section relationships or body text. It closes the recurring failure
    where visible header/footer text looks close, but paragraph style ids,
    footer style, and PAGE field typography drift away from the locked template.
    Existing field instructions are protected review artifacts; this pass may
    restyle PAGE runs but must not rewrite the PAGE field code itself.
    """
    if template_styles_root is None or template_parts is None:
        raise RuntimeError("template-header-footer-styles requires --template-docx")

    header_style_id = style_id_by_name(template_styles_root, "header")
    footer_style_id = style_id_by_name(template_styles_root, "footer")
    if not header_style_id or not footer_style_id:
        raise RuntimeError("template DOCX must define paragraph styles named header and footer")

    style_changes: list[dict[str, object]] = []
    for surface, style_id in (("header", header_style_id), ("footer", footer_style_id)):
        donor = style_by_id(template_styles_root, style_id)
        if donor is None:
            raise RuntimeError(f"template DOCX does not define {surface} style `{style_id}`")
        status = replace_or_append_style_definition(styles_root, donor)
        style_changes.append({"kind": f"{surface}_style_definition", "style_id": style_id, "status": status})

    donor_header_rpr: ET.Element | None = None
    donor_footer_rpr: ET.Element | None = None
    for name in sorted(template_parts):
        if donor_header_rpr is None and re.fullmatch(r"word/header\d+\.xml", name):
            root = ET.fromstring(template_parts[name])
            donor_para = first_paragraph_with_style(root, header_style_id)
            if donor_para is not None:
                donor_header_rpr = first_run_rpr(donor_para)
        if donor_footer_rpr is None and re.fullmatch(r"word/footer\d+\.xml", name):
            root = ET.fromstring(template_parts[name])
            donor_para = first_paragraph_with_style(root, footer_style_id)
            if donor_para is not None:
                donor_footer_rpr = first_run_rpr(donor_para)
        if donor_header_rpr is not None and donor_footer_rpr is not None:
            break

    updates: dict[str, bytes] = {}
    part_changes: list[dict[str, object]] = []
    for name in sorted(parts):
        if re.fullmatch(r"word/header\d+\.xml", name):
            root = ET.fromstring(parts[name])
            paragraph_count = 0
            run_rpr_changes = 0
            before_text = paragraph_text(root)
            for paragraph in root.findall(".//w:p", NS):
                if paragraph_text(paragraph).strip() or paragraph.find(".//w:tab", NS) is not None:
                    previous_style = paragraph_style_id(paragraph)
                    set_paragraph_style(paragraph, header_style_id)
                    paragraph_count += 1 if previous_style != header_style_id else 0
                    for run in paragraph.findall("./w:r", NS):
                        if set_run_rpr(run, donor_header_rpr):
                            run_rpr_changes += 1
            if paragraph_count or run_rpr_changes:
                updates[name] = serialize_xml(root)
                part_changes.append(
                    {
                        "part": name,
                        "surface": "header",
                        "style_id": header_style_id,
                        "paragraph_style_changes": paragraph_count,
                        "run_rpr_changes": run_rpr_changes,
                        "text": before_text,
                    }
                )
        elif re.fullmatch(r"word/footer\d+\.xml", name):
            root = ET.fromstring(parts[name])
            paragraph_count = 0
            run_rpr_changes = 0
            for paragraph in root.findall(".//w:p", NS):
                previous_style = paragraph_style_id(paragraph)
                set_paragraph_style(paragraph, footer_style_id)
                paragraph_count += 1 if previous_style != footer_style_id else 0
                for run in paragraph.findall("./w:r", NS):
                    if set_run_rpr(run, donor_footer_rpr):
                        run_rpr_changes += 1
            if paragraph_count or run_rpr_changes:
                updates[name] = serialize_xml(root)
                part_changes.append(
                    {
                        "part": name,
                        "surface": "footer",
                        "style_id": footer_style_id,
                        "paragraph_style_changes": paragraph_count,
                        "page_instr_preservation": "preserved-existing-instructions",
                        "run_rpr_changes": run_rpr_changes,
                    }
                )

    return style_changes + part_changes, updates


def remove_page_break_before_property(paragraph: ET.Element) -> bool:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return False
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        return False
    ppr.remove(page_break)
    return True


def template_toc_page_owner_spacers(template_document_root: ET.Element | None) -> list[ET.Element]:
    if template_document_root is None:
        return []
    template_body = template_document_root.find("./w:body", NS)
    if template_body is None:
        return []
    try:
        toc_index, _body_start = find_toc_and_body_start(template_body)
    except RuntimeError:
        return []
    children = list(template_body)
    spacers: list[ET.Element] = []
    probe_index = toc_index - 1
    while probe_index >= 0:
        candidate = children[probe_index]
        if candidate.tag != qn("p"):
            break
        if paragraph_text(candidate).strip():
            break
        ppr = paragraph_ppr(candidate)
        if ppr is not None and ppr.find("./w:sectPr", NS) is not None:
            break
        spacers.append(deepcopy(candidate))
        probe_index -= 1
    return list(reversed(spacers))


def repair_toc_title_pagebreak(
    body: ET.Element,
    template_document_root: ET.Element | None = None,
) -> list[dict[str, object]]:
    toc_index, _body_start = find_toc_and_body_start(body)
    toc_child = list(body)[toc_index]
    title_paragraph: ET.Element | None = None
    if toc_child.tag == qn("p") and is_toc_heading(paragraph_text(toc_child).strip()):
        title_paragraph = toc_child
    else:
        content = toc_child_content(toc_child)
        if content is not None:
            for paragraph in content.findall("./w:p", NS):
                if is_toc_heading(paragraph_text(paragraph).strip()):
                    title_paragraph = paragraph
                    break
    if title_paragraph is None:
        raise RuntimeError("toc-title-pagebreak could not locate TOC title paragraph")
    ppr = ensure_ppr(title_paragraph)
    removed_toc_title_outline_children: list[str] = []
    for tag in ("outlineLvl", "numPr"):
        for node in list(ppr.findall(f"./w:{tag}", NS)):
            ppr.remove(node)
            removed_toc_title_outline_children.append(tag)
    template_spacers = template_toc_page_owner_spacers(template_document_root)
    if template_spacers and toc_child.tag == qn("sdt"):
        removed_title_pagebreak = remove_page_break_before_property(title_paragraph)
        inserted: list[str] = []
        for offset, spacer in enumerate(template_spacers):
            clone = deepcopy(spacer)
            if offset == 0:
                ensure_true_page_break_before(clone)
            body.insert(toc_index + offset, clone)
            inserted.append(paragraph_text(clone).strip())
        return [
            {
                "kind": "toc_title_page_owner_from_template_spacers",
                "body_child_index": toc_index,
                "text": paragraph_text(title_paragraph).strip(),
                "inserted_spacer_count": len(template_spacers),
                "inserted_spacer_text": inserted,
                "removed_title_pageBreakBefore": removed_title_pagebreak,
                "removed_toc_title_outline_children": removed_toc_title_outline_children,
                "reason": "template has a blank paragraph before the TOC content-control; pagination belongs to that spacer so title vertical position matches the template",
            }
        ]
    changed, reason = ensure_true_page_break_before(title_paragraph)
    return [
        {
            "kind": "toc_title_pagebreak",
            "body_child_index": toc_index,
            "text": paragraph_text(title_paragraph).strip(),
            "changed": changed,
            "removed_toc_title_outline_children": removed_toc_title_outline_children,
            "reason": reason,
        }
    ]


def repair_frontmatter_title_outline_exclusion(body: ET.Element) -> list[dict[str, object]]:
    """Exclude abstract/front-matter titles from live TOC outline collection.

    The live TOC field uses the ``\\u`` switch so body headings with local
    outline levels can appear even when their style ids are template-specific.
    Front-matter titles before the TOC must therefore not carry direct
    outline/numbering properties, otherwise WPS refreshes the live TOC with
    stale abstract rows and can push the TOC/body section break onto a blank
    page.
    """
    toc_index, _body_start = find_toc_and_body_start(body)
    changes: list[dict[str, object]] = []
    frontmatter_title_labels = {
        compact_text("\u6458\u8981").lower(),
        "abstract",
    }
    for index, child in enumerate(list(body)[:toc_index]):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if compact_text(text).lower() not in frontmatter_title_labels:
            continue
        ppr = paragraph_ppr(child)
        if ppr is None:
            continue
        removed: list[str] = []
        for tag in ("outlineLvl", "numPr"):
            for node in list(ppr.findall(f"./w:{tag}", NS)):
                ppr.remove(node)
                removed.append(tag)
        if removed:
            changes.append(
                {
                    "kind": "frontmatter_title_outline_exclusion",
                    "body_child_index": index,
                    "text": text,
                    "removed_children": removed,
                    "reason": "front-matter title must not be collected by live TOC outline switch",
                }
            )
    return changes


def toc_entry_is_frontmatter_cache(paragraph: ET.Element) -> bool:
    parsed = toc_entry_label_and_page_from_paragraph(paragraph)
    if parsed is None:
        return False
    label, page = parsed
    return is_front_matter_toc_label(f"{label}{page}")


def paragraph_field_prefix_children(paragraph: ET.Element) -> list[ET.Element]:
    prefix: list[ET.Element] = []
    for child in list(paragraph):
        if child.tag == qn("pPr"):
            continue
        has_field_marker = (
            child.find(".//w:fldChar", NS) is not None
            or child.find(".//w:instrText", NS) is not None
        )
        if has_field_marker:
            prefix.append(deepcopy(child))
            continue
        if prefix:
            break
    return prefix


def paragraph_visible_nonfield_children(paragraph: ET.Element) -> list[ET.Element]:
    visible: list[ET.Element] = []
    for child in list(paragraph):
        if child.tag == qn("pPr"):
            continue
        if child.find(".//w:fldChar", NS) is not None or child.find(".//w:instrText", NS) is not None:
            continue
        visible.append(deepcopy(child))
    return visible


def rewrite_toc_host_from_donor_preserving_field_prefix(host: ET.Element, donor: ET.Element) -> int:
    prefix = paragraph_field_prefix_children(host)
    visible = paragraph_visible_nonfield_children(donor)
    ppr = host.find("./w:pPr", NS)
    preserved_ppr = deepcopy(ppr) if ppr is not None else None
    for child in list(host):
        host.remove(child)
    if preserved_ppr is not None:
        host.append(preserved_ppr)
    for child in prefix + visible:
        host.append(child)
    return len(prefix)


def repair_toc_frontmatter_cache_exclusion(body: ET.Element) -> list[dict[str, object]]:
    """Remove stale abstract rows from the cached live-TOC result.

    WPS refreshes the live TOC from body outline levels after
    ``frontmatter-title-outline-exclusion``. The cached DOCX result should
    mirror that body-only TOC so opening the document before a field refresh
    does not show abstract rows that the rendered PDF no longer contains.
    """
    toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    entry_indices = [
        index
        for index in range(toc_index + 1, body_start)
        if children[index].tag == qn("p") and toc_entry_label_and_page_from_paragraph(children[index]) is not None
    ]
    frontmatter_indices: list[int] = []
    for index in entry_indices:
        paragraph = children[index]
        if toc_entry_is_frontmatter_cache(paragraph):
            frontmatter_indices.append(index)
            continue
        break
    if not frontmatter_indices:
        return []
    donor_index = next((index for index in entry_indices if index not in frontmatter_indices), None)
    if donor_index is None:
        raise RuntimeError("toc-frontmatter-cache-exclusion could not locate first body TOC donor")
    host_index = frontmatter_indices[0]
    host = children[host_index]
    donor = children[donor_index]
    moved_prefix_count = rewrite_toc_host_from_donor_preserving_field_prefix(host, donor)
    remove_indices = sorted(set(frontmatter_indices[1:] + [donor_index]), reverse=True)
    removed_rows: list[dict[str, object]] = []
    for index in remove_indices:
        paragraph = children[index]
        removed_rows.append({"body_child_index": index, "text": paragraph_text(paragraph).strip()})
        body.remove(paragraph)
    return [
        {
            "kind": "toc_frontmatter_cache_exclusion",
            "host_body_child_index": host_index,
            "donor_body_child_index": donor_index,
            "removed_rows": list(reversed(removed_rows)),
            "field_prefix_run_count_preserved": moved_prefix_count,
            "reason": "cached live-TOC result follows body-only outline collection",
        }
    ]


def repair_toc_field_lock(body: ET.Element) -> list[dict[str, object]]:
    """Lock reviewed TOC field caches so WPS export does not refresh them."""
    changes: list[dict[str, object]] = []
    for index, paragraph in enumerate(list(body)):
        if paragraph.tag != qn("p"):
            continue
        instr = "".join(node.text or "" for node in paragraph.findall(".//w:instrText", NS))
        if "TOC" not in instr.upper():
            continue
        locked = 0
        for fld_char in paragraph.findall(".//w:fldChar", NS):
            if fld_char.get(qn("fldCharType")) != "begin":
                continue
            if fld_char.get(qn("fldLock")) != "true":
                fld_char.set(qn("fldLock"), "true")
                locked += 1
        if locked:
            changes.append(
                {
                    "kind": "toc_field_lock",
                    "body_child_index": index,
                    "locked_begin_field_count": locked,
                    "reason": "preserve reviewed TOC cache during WPS PDF export",
                }
            )
    return changes


def paragraph_field_char_types(paragraph: ET.Element) -> list[str]:
    return [
        value
        for value in (field.get(qn("fldCharType")) for field in paragraph.findall(".//w:fldChar", NS))
        if value
    ]


def paragraph_has_field_surface(paragraph: ET.Element) -> bool:
    return bool(paragraph.findall(".//w:fldChar", NS) or paragraph.findall(".//w:instrText", NS))


def make_section_boundary_paragraph(sect: ET.Element) -> ET.Element:
    paragraph = ET.Element(qn("p"))
    ppr = ET.SubElement(paragraph, qn("pPr"))
    ppr.append(sect)
    return paragraph


def repair_toc_field_boundary(body: ET.Element) -> list[dict[str, object]]:
    """Move the TOC/body section break out of the live TOC field result."""
    toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    existing_boundary_index: int | None = None
    for index in range(max(toc_index, body_start - 4), body_start):
        child = children[index]
        if (
            child.tag == qn("p")
            and has_section_break(child)
            and not paragraph_text(child).strip()
            and not paragraph_has_field_surface(child)
        ):
            existing_boundary_index = index
            break
    if existing_boundary_index is not None:
        return [
            {
                "kind": "toc_field_boundary",
                "changed": False,
                "reason": "independent field-free section boundary already present",
                "body_child_index": existing_boundary_index,
            }
        ]

    section_owner_index: int | None = None
    for index in range(body_start - 1, toc_index, -1):
        child = children[index]
        if child.tag == qn("p") and has_section_break(child):
            section_owner_index = index
            break
    if section_owner_index is None:
        raise RuntimeError("toc-field-boundary could not locate a TOC/body section break before the first body heading")

    section_owner = children[section_owner_index]
    owner_text = paragraph_text(section_owner).strip()
    owner_fields = paragraph_field_char_types(section_owner)
    sect = take_paragraph_section_break(section_owner)
    if sect is None:
        raise RuntimeError("toc-field-boundary could not detach the TOC/body section break")

    boundary = make_section_boundary_paragraph(sect)
    body.insert(section_owner_index + 1, boundary)
    return [
        {
            "kind": "toc_field_boundary",
            "changed": True,
            "from_body_child_index": section_owner_index,
            "to_body_child_index": section_owner_index + 1,
            "from_text": owner_text[:120],
            "from_field_char_types": owner_fields,
            "reason": "moved section break from TOC cache paragraph to independent field-free empty paragraph",
        }
    ]


def repair_toc_tail_section_collapse(body: ET.Element) -> list[dict[str, object]]:
    """Remove redundant blank section owners between terminal TOC rows.

    A reviewed TOC may contain both a blank field-free section boundary and a
    section break on the final tail TOC row. LibreOffice/WPS can then render the
    final row, usually ``致谢``, as a near-empty standalone page even when the
    prior TOC page has room. Keep the final row's section break so the first
    body heading still starts after the TOC, and remove only redundant blank
    section-owner paragraphs between the ``参考文献`` and ``致谢`` TOC rows.
    """
    toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    terminal_entries: dict[str, tuple[int, ET.Element, str]] = {}
    for index in range(toc_index + 1, body_start):
        child = children[index]
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        parsed = toc_entry_label_and_page_from_paragraph(child)
        if parsed is None:
            continue
        label, page = parsed
        compact = compact_text(label)
        if compact == compact_text(REFERENCES_LABEL):
            terminal_entries[REFERENCES_LABEL] = (index, child, page)
        elif compact == compact_text(ACKNOWLEDGEMENT_LABEL):
            terminal_entries[ACKNOWLEDGEMENT_LABEL] = (index, child, page)

    references = terminal_entries.get(REFERENCES_LABEL)
    acknowledgement = terminal_entries.get(ACKNOWLEDGEMENT_LABEL)
    if references is None or acknowledgement is None:
        return [
            {
                "kind": "toc_tail_section_collapse",
                "changed": False,
                "reason": "terminal TOC rows not both present",
                "found_labels": sorted(terminal_entries),
            }
        ]

    references_index, references_entry, references_page = references
    acknowledgement_index, acknowledgement_entry, acknowledgement_page = acknowledgement
    if acknowledgement_index <= references_index:
        return [
            {
                "kind": "toc_tail_section_collapse",
                "changed": False,
                "reason": "acknowledgement TOC row is not after references TOC row",
                "references_index": references_index,
                "acknowledgement_index": acknowledgement_index,
            }
        ]

    removed: list[dict[str, object]] = []
    for index in range(acknowledgement_index - 1, references_index, -1):
        current_children = list(body)
        child = current_children[index]
        if child.tag != qn("p") or paragraph_text(child).strip():
            continue
        section_removed = remove_paragraph_section_break(child)
        if not section_removed:
            continue
        body.remove(child)
        removed.append(
            {
                "body_child_index": index,
                "reason": "blank section owner between terminal TOC rows removed",
            }
        )

    return [
        {
            "kind": "toc_tail_section_collapse",
            "changed": bool(removed),
            "references_index": references_index,
            "acknowledgement_index_before": acknowledgement_index,
            "references_text": paragraph_text(references_entry).strip()[:120],
            "acknowledgement_text": paragraph_text(acknowledgement_entry).strip()[:120],
            "references_page_cache": references_page,
            "acknowledgement_page_cache": acknowledgement_page,
            "removed_blank_section_owner_count": len(removed),
            "removed": list(reversed(removed)),
        }
    ]


def default_header_rid(sect: ET.Element | None) -> str:
    if sect is None:
        return ""
    for ref in sect.findall("./w:headerReference", NS):
        if (ref.get(qn("type")) or "default") == "default":
            return ref.get(R + "id", "")
    return ""


def default_footer_rid(sect: ET.Element | None) -> str:
    if sect is None:
        return ""
    for ref in sect.findall("./w:footerReference", NS):
        if (ref.get(qn("type")) or "default") == "default":
            return ref.get(R + "id", "")
    return ""


def set_default_header_rid(sect: ET.Element, rid: str) -> bool:
    if not rid:
        return False
    changed = False
    default_refs = [
        ref
        for ref in sect.findall("./w:headerReference", NS)
        if (ref.get(qn("type")) or "default") == "default"
    ]
    if not default_refs:
        ref = ET.Element(qn("headerReference"))
        ref.set(qn("type"), "default")
        ref.set(R + "id", rid)
        insert_at = 0
        while insert_at < len(list(sect)) and list(sect)[insert_at].tag in {
            qn("headerReference"),
            qn("footerReference"),
        }:
            insert_at += 1
        sect.insert(insert_at, ref)
        return True
    for ref in default_refs:
        if ref.get(R + "id", "") != rid:
            ref.set(R + "id", rid)
            changed = True
    return changed


def set_default_footer_rid(sect: ET.Element, rid: str) -> bool:
    if not rid:
        return False
    changed = False
    default_refs = [
        ref
        for ref in sect.findall("./w:footerReference", NS)
        if (ref.get(qn("type")) or "default") == "default"
    ]
    if not default_refs:
        ref = ET.Element(qn("footerReference"))
        ref.set(qn("type"), "default")
        ref.set(R + "id", rid)
        insert_at = 0
        while insert_at < len(list(sect)) and list(sect)[insert_at].tag in {
            qn("headerReference"),
            qn("footerReference"),
        }:
            insert_at += 1
        sect.insert(insert_at, ref)
        return True
    for ref in default_refs:
        if ref.get(R + "id", "") != rid:
            ref.set(R + "id", rid)
            changed = True
    return changed


def find_default_section_reference_donors(body: ET.Element) -> tuple[str, str]:
    section_properties = body.findall(".//w:sectPr", NS)
    body_level_section = body.find("./w:sectPr", NS)
    candidates: list[ET.Element] = []
    if body_level_section is not None:
        candidates.append(body_level_section)
    candidates.extend(reversed(section_properties))
    header_rid = ""
    footer_rid = ""
    for sect in candidates:
        if not header_rid:
            header_rid = default_header_rid(sect)
        if not footer_rid:
            footer_rid = default_footer_rid(sect)
        if header_rid and footer_rid:
            break
    return header_rid, footer_rid


def repair_body_section_header_footer(body: ET.Element) -> list[dict[str, object]]:
    """Ensure body, references, and appendix sections carry explicit header/footer refs."""
    _toc_index, body_start = find_toc_and_body_start(body)
    header_rid, footer_rid = find_default_section_reference_donors(body)
    changes: list[dict[str, object]] = []
    if not header_rid and not footer_rid:
        return [
            {
                "kind": "body_section_header_footer",
                "changed": False,
                "reason": "no default header/footer relationship donor found",
            }
        ]
    for index, child in enumerate(list(body)):
        if index < body_start or child.tag != qn("p"):
            continue
        sect = child.find("./w:pPr/w:sectPr", NS)
        if sect is None:
            continue
        header_changed = set_default_header_rid(sect, header_rid)
        footer_changed = set_default_footer_rid(sect, footer_rid)
        if header_changed or footer_changed:
            changes.append(
                {
                    "kind": "body_section_header_footer",
                    "body_child_index": index,
                    "text": paragraph_text(child).strip()[:120],
                    "header_relationship_id": header_rid,
                    "footer_relationship_id": footer_rid,
                    "header_changed": header_changed,
                    "footer_changed": footer_changed,
                    "reason": "body/tail section must explicitly carry the template running header and footer",
                }
            )
    return changes


def repair_toc_last_entry_header_section(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    """Split the final static TOC tail row into the template's running-header section.

    Some school templates place the final tail TOC row under the institutional
    running header while the earlier TOC rows stay header-free. The split must
    stay outside the live TOC field, otherwise Word/WPS can expose a field
    boundary failure.
    """
    toc_index, body_start = find_toc_and_body_start(body)
    children = list(body)
    entries = [
        (body_index, paragraph, level)
        for body_index, paragraph, level in toc_entry_paragraphs_in_range(body, toc_index, body_start, styles_root, {})
        if paragraph_text(paragraph).strip() and is_toc_entry_paragraph(paragraph)
    ]
    if len(entries) < 2:
        return [
            {
                "kind": "toc_last_entry_header_section",
                "changed": False,
                "reason": "fewer than two TOC entries",
            }
        ]

    last_index, last_entry, _last_level = entries[-1]
    previous_index, previous_entry, _previous_level = entries[-2]
    existing_split = any(
        child.tag == qn("p") and has_section_break(child)
        for child in children[previous_index + 1 : last_index]
    )
    boundary_index: int | None = None
    for index in range(last_index + 1, body_start):
        child = children[index]
        if child.tag == qn("p") and has_section_break(child) and not paragraph_text(child).strip():
            boundary_index = index
            break
    if boundary_index is None:
        return [
            {
                "kind": "toc_last_entry_header_section",
                "changed": False,
                "reason": "no field-free TOC/body boundary found after final TOC entry",
            }
        ]

    boundary = children[boundary_index]
    boundary_sect = boundary.find("./w:pPr/w:sectPr", NS)
    final_body_sect = body.find("./w:sectPr", NS)
    body_header_rid = default_header_rid(final_body_sect)
    if not body_header_rid:
        return [
            {
                "kind": "toc_last_entry_header_section",
                "changed": False,
                "reason": "final body section has no default header relationship",
            }
        ]

    moved_field_end = False
    if paragraph_field_char_types(last_entry).count("end"):
        moved_field_end = remove_last_field_end(last_entry)
        if moved_field_end:
            append_main_toc_field_end(previous_entry)

    inserted_split = False
    if not existing_split and boundary_sect is not None:
        split_sect = deepcopy(boundary_sect)
        set_section_start_type(split_sect, "continuous")
        body.insert(last_index, make_section_boundary_paragraph(split_sect))
        inserted_split = True
        children = list(body)
        boundary = children[boundary_index + 1]
        boundary_sect = boundary.find("./w:pPr/w:sectPr", NS)

    boundary_header_changed = False
    if boundary_sect is not None:
        boundary_header_changed = set_default_header_rid(boundary_sect, body_header_rid)

    return [
        {
            "kind": "toc_last_entry_header_section",
            "changed": inserted_split or boundary_header_changed or moved_field_end,
            "inserted_split_before_body_child_index": last_index if inserted_split else None,
            "last_entry_text": paragraph_text(last_entry).strip()[:120],
            "last_entry_original_body_child_index": last_index,
            "previous_entry_body_child_index": previous_index,
            "toc_body_boundary_original_index": boundary_index,
            "body_header_relationship_id": body_header_rid,
            "inserted_continuous_section": inserted_split,
            "boundary_header_changed": boundary_header_changed,
            "field_end_moved_to_previous_entry": moved_field_end,
        }
    ]


def repair_toc_frontmatter_cache_compact(body: ET.Element) -> list[dict[str, object]]:
    """Compact required abstract TOC cache rows so the TOC/body break fits."""
    toc_index, body_start = find_toc_and_body_start(body)
    changes: list[dict[str, object]] = []
    for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1):
        if child.tag != qn("p") or not toc_entry_is_frontmatter_cache(child):
            continue
        ppr = ensure_ppr(child)
        spacing = ppr.find("./w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(qn("spacing"))
            ppr.append(spacing)
        before = spacing.get(qn("before"))
        after = spacing.get(qn("after"))
        line = spacing.get(qn("line"))
        line_rule = spacing.get(qn("lineRule"))
        spacing.set(qn("before"), "0")
        spacing.set(qn("after"), "0")
        spacing.set(qn("line"), "200")
        spacing.set(qn("lineRule"), "exact")
        if (before, after, line, line_rule) != ("0", "0", "200", "exact"):
            changes.append(
                {
                    "kind": "toc_frontmatter_cache_compact",
                    "body_child_index": index,
                    "text": paragraph_text(child).strip(),
                    "old_spacing": {
                        "before": before,
                        "after": after,
                        "line": line,
                        "lineRule": line_rule,
                    },
                    "new_spacing": {
                        "before": "0",
                        "after": "0",
                        "line": "200",
                        "lineRule": "exact",
                    },
                    "reason": "required abstract TOC rows remain visible without pushing the final TOC row to a near-empty continuation page",
                }
            )
    return changes


def paragraph_has_nontext_payload(paragraph: ET.Element) -> bool:
    return any(paragraph.find(f".//w:{tag}", NS) is not None for tag in ("drawing", "pict", "object", "fldChar", "instrText"))


def is_blank_paragraph_without_payload(paragraph: ET.Element) -> bool:
    return paragraph.tag == qn("p") and not paragraph_text(paragraph).strip() and not paragraph_has_nontext_payload(paragraph)


def remove_blank_range_between(body: ET.Element, start_exclusive: int, end_exclusive: int) -> list[dict[str, object]]:
    removed: list[dict[str, object]] = []
    children = list(body)
    for index in range(end_exclusive - 1, start_exclusive, -1):
        child = children[index]
        if child.tag != qn("p") or not is_blank_paragraph_without_payload(child):
            continue
        body.remove(child)
        removed.append({"body_child_index": index, "reason": "blank separator removed after section ownership migration"})
    return list(reversed(removed))


def paragraph_has_page_break(paragraph: ET.Element) -> bool:
    return any(node.get(qn("type"), "textWrapping") == "page" for node in paragraph.findall(".//w:br", NS))


def paragraph_has_real_image(paragraph: ET.Element) -> bool:
    """Mirror sample_self_check's real-image predicate for blank-paragraph repair."""
    return (
        paragraph.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip") is not None
        or paragraph.find(".//{urn:schemas-microsoft-com:vml}imagedata") is not None
    )


def is_removable_blank_pre_submission_paragraph(paragraph: ET.Element) -> bool:
    if paragraph.tag != qn("p") or paragraph_text(paragraph).strip():
        return False
    if paragraph_has_page_break(paragraph) or has_section_break(paragraph):
        return False
    if not paragraph_has_nontext_payload(paragraph):
        return True
    children = [node for node in list(paragraph) if node.tag != qn("pPr")]
    if not children:
        return True
    for child in children:
        if child.tag != qn("r"):
            return False
        meaningful = [node for node in list(child) if node.tag != qn("rPr")]
        if len(meaningful) != 1 or meaningful[0].tag != qn("fldChar"):
            return False
        if meaningful[0].get(qn("fldCharType")) != "end":
            return False
    return True


def remove_pre_submission_blank_paragraphs(body: ET.Element) -> list[dict[str, object]]:
    """Remove blank paragraphs caught by sample_self_check's pre-submission gate."""
    children = list(body)
    first_body_idx = next(
        (
            index
            for index, child in enumerate(children)
            if child.tag == qn("p")
            and heading_level(paragraph_text(child).strip()) == 1
        ),
        0,
    )
    references_idx = bibliography_heading_index(children)
    removable: list[tuple[int, ET.Element, str, str]] = []
    for index, child in enumerate(children):
        if index <= first_body_idx or index >= references_idx:
            continue
        if not is_removable_blank_pre_submission_paragraph(child):
            continue
        previous = next(
            (
                item
                for item in reversed(children[:index])
                if item.tag == qn("p") and paragraph_text(item).strip()
            ),
            None,
        )
        following = next(
            (
                item
                for item in children[index + 1 :]
                if item.tag == qn("p") and paragraph_text(item).strip()
            ),
            None,
        )
        if previous is None or following is None:
            continue
        prev_text = paragraph_text(previous).strip()
        next_text = paragraph_text(following).strip()
        if paragraph_has_real_image(previous) or paragraph_has_real_image(following):
            continue
        if CAPTION_RE.match(prev_text) or CAPTION_RE.match(next_text):
            continue
        if heading_level(prev_text) is not None or heading_level(next_text) is not None:
            continue
        removable.append((index, child, prev_text, next_text))
    changes: list[dict[str, object]] = []
    for index, child, prev_text, next_text in reversed(removable):
        body.remove(child)
        changes.append(
            {
                "body_child_index": index,
                "previous_text": prev_text[:120],
                "next_text": next_text[:120],
                "reason": "no-payload blank paragraph removed for common pre-submission gate",
            }
        )

    records: list[dict[str, object]] = []

    def walk(parent: ET.Element, *, in_table: bool = False) -> None:
        child_in_table = in_table or parent.tag == qn("tbl")
        for child in list(parent):
            if child.tag == qn("p"):
                records.append(
                    {
                        "parent": parent,
                        "paragraph": child,
                        "text": paragraph_text(child).strip(),
                        "in_table": child_in_table,
                    }
                )
            else:
                walk(child, in_table=child_in_table)

    walk(body)
    first_body_record = next((idx for idx, item in enumerate(records) if heading_level(str(item["text"])) == 1), 0)
    references_record = next(
        (
            idx
            for idx, item in enumerate(records)
            if compact_text(str(item["text"])) == compact_text("参考文献")
        ),
        len(records),
    )
    removed_ids = {id(item[1]) for item in removable}
    for index, item in reversed(list(enumerate(records))):
        if index <= first_body_record or index >= references_record:
            continue
        paragraph = item["paragraph"]
        if not isinstance(paragraph, ET.Element) or id(paragraph) in removed_ids:
            continue
        if item.get("in_table") or not is_removable_blank_pre_submission_paragraph(paragraph):
            continue
        previous = records[index - 1] if index > 0 else None
        following = records[index + 1] if index + 1 < len(records) else None
        if previous is None or following is None:
            continue
        prev_para = previous["paragraph"]
        next_para = following["paragraph"]
        if not isinstance(prev_para, ET.Element) or not isinstance(next_para, ET.Element):
            continue
        prev_text = str(previous.get("text") or "").strip()
        next_text = str(following.get("text") or "").strip()
        if not prev_text or not next_text:
            continue
        if paragraph_has_real_image(prev_para) or paragraph_has_real_image(next_para):
            continue
        if CAPTION_RE.match(prev_text) or CAPTION_RE.match(next_text):
            continue
        if heading_level(prev_text) is not None or heading_level(next_text) is not None:
            continue
        parent = item["parent"]
        if isinstance(parent, ET.Element):
            parent.remove(paragraph)
            removed_ids.add(id(paragraph))
            changes.append(
                {
                    "paragraph_record_index": index,
                    "previous_text": prev_text[:120],
                    "next_text": next_text[:120],
                    "reason": "recursive no-payload blank paragraph removed for common pre-submission gate",
                }
            )
    return list(reversed(changes))


def remove_adjacent_duplicate_long_body_paragraphs(body: ET.Element) -> list[dict[str, object]]:
    """Remove later copies of duplicated long body paragraphs reported by sample_self_check."""
    children = list(body)
    first_body_idx = next(
        (
            index
            for index, child in enumerate(children)
            if child.tag == qn("p")
            and heading_level(paragraph_text(child).strip()) == 1
        ),
        0,
    )
    references_idx = bibliography_heading_index(children)
    removable: list[tuple[int, ET.Element, str]] = []
    seen_keys: set[str] = set()
    for index, child in enumerate(children):
        if index <= first_body_idx or index >= references_idx:
            continue
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        key = compact_text(text)
        is_body_text = bool(key) and len(key) >= 60 and heading_level(text) is None
        if (
            is_body_text
            and key in seen_keys
            and child.find("./w:pPr/w:sectPr", NS) is None
            and child.find(".//w:drawing", NS) is None
            and child.find(".//w:pict", NS) is None
            and child.find(".//w:object", NS) is None
        ):
            removable.append((index, child, text))
            continue
        if is_body_text:
            seen_keys.add(key)
    changes: list[dict[str, object]] = []
    for index, child, text in reversed(removable):
        body.remove(child)
        changes.append(
            {
                "body_child_index": index,
                "text": text[:120],
                "reason": "later duplicate long body paragraph removed for full-thesis content gate",
            }
        )
    return list(reversed(changes))


def migrate_blank_section_to_previous_content(
    body: ET.Element,
    *,
    previous_content_index: int,
    next_content_index: int,
    location: str,
) -> dict[str, object]:
    children = list(body)
    section_owner_index: int | None = None
    for index in range(previous_content_index + 1, next_content_index):
        child = children[index]
        if child.tag == qn("p") and is_blank_paragraph_without_payload(child) and has_section_break(child):
            section_owner_index = index
            break
    if section_owner_index is None:
        return {
            "kind": "frontmatter_section_owner_migration",
            "location": location,
            "changed": False,
            "reason": "no blank section owner found in separator range",
        }
    section_owner = children[section_owner_index]
    sect = take_paragraph_section_break(section_owner)
    if sect is None:
        return {
            "kind": "frontmatter_section_owner_migration",
            "location": location,
            "changed": False,
            "reason": "blank section owner had no movable sectPr",
        }
    replace_paragraph_section_break(children[previous_content_index], sect)
    removed = remove_blank_range_between(body, previous_content_index, next_content_index)
    return {
        "kind": "frontmatter_section_owner_migration",
        "location": location,
        "changed": True,
        "from_blank_index": section_owner_index,
        "to_previous_content_index": previous_content_index,
        "next_content_index": next_content_index,
        "removed_blank_count": len(removed),
        "removed": removed,
    }


def migrate_text_section_to_following_blank(
    body: ET.Element,
    *,
    text_paragraph_index: int,
    stop_exclusive: int,
    location: str,
) -> dict[str, object]:
    children = list(body)
    source = children[text_paragraph_index]
    if source.tag != qn("p"):
        return {
            "kind": "frontmatter_text_section_migration",
            "location": location,
            "changed": False,
            "reason": "source is not a paragraph",
        }
    sect = take_paragraph_section_break(source)
    if sect is None:
        return {
            "kind": "frontmatter_text_section_migration",
            "location": location,
            "changed": False,
            "reason": "text paragraph has no sectPr",
        }
    target_index: int | None = None
    for index in range(text_paragraph_index + 1, stop_exclusive):
        child = children[index]
        if child.tag == qn("p") and is_blank_paragraph_without_payload(child):
            target_index = index
            break
    inserted_blank = False
    if target_index is None:
        blank = ET.Element(qn("p"))
        blank.append(ET.Element(qn("pPr")))
        body.insert(text_paragraph_index + 1, blank)
        target_index = text_paragraph_index + 1
        inserted_blank = True
        children = list(body)
    replace_paragraph_section_break(children[target_index], sect)
    duplicate_removed: list[dict[str, object]] = []
    children_after = list(body)
    adjusted_stop = stop_exclusive + (1 if inserted_blank else 0)
    for index in range(adjusted_stop - 1, target_index, -1):
        child = children_after[index]
        if child.tag != qn("p") or not is_blank_paragraph_without_payload(child) or not has_section_break(child):
            continue
        remove_paragraph_section_break(child)
        duplicate_removed.append({"body_child_index": index, "reason": "duplicate blank section owner removed"})
    return {
        "kind": "frontmatter_text_section_migration",
        "location": location,
        "changed": True,
        "from_text_index": text_paragraph_index,
        "to_blank_index": target_index,
        "inserted_blank": inserted_blank,
        "duplicate_section_breaks_removed": list(reversed(duplicate_removed)),
    }


def collapse_declaration_to_abstract_separator(body: ET.Element) -> list[dict[str, object]]:
    zh_abstract_index = first_frontmatter_child_index(
        body,
        lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    children = list(body)
    section_owner_index: int | None = None
    for index in range(zh_abstract_index - 1, -1, -1):
        child = children[index]
        if child.tag == qn("p") and has_section_break(child):
            section_owner_index = index
            break
    if section_owner_index is None:
        return [
            {
                "kind": "frontmatter_declaration_separator",
                "changed": False,
                "reason": "no section owner found before Chinese abstract",
            }
        ]
    section_owner = children[section_owner_index]
    if not is_blank_paragraph_without_payload(section_owner):
        return [
            {
                "kind": "frontmatter_declaration_separator",
                "changed": False,
                "section_owner_index": section_owner_index,
                "reason": "section owner before Chinese abstract is not a removable blank paragraph",
            }
        ]
    zh_title_index = previous_nonempty_child_index(body, start_exclusive=zh_abstract_index, stop_exclusive=section_owner_index)
    zh_title = children[zh_title_index]
    zh_keyword_index = first_frontmatter_child_index(
        body,
        lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p),
    )
    zh_section_type_changed = False
    for index in range(zh_abstract_index, zh_keyword_index + 1):
        child = children[index]
        sect = paragraph_ppr(child).find("./w:sectPr", NS) if child.tag == qn("p") and paragraph_ppr(child) is not None else None
        if sect is None:
            continue
        zh_section_type_changed = set_section_start_type(sect, "continuous")
        break
    zh_title_page_break_changed, zh_title_page_break_reason = ensure_true_page_break_before(zh_title)
    previous_content_index = previous_nonempty_child_index(body, start_exclusive=section_owner_index, stop_exclusive=0)
    sect = take_paragraph_section_break(section_owner)
    migrated = False
    if sect is not None:
        replace_paragraph_section_break(children[previous_content_index], sect)
        migrated = True
    removed = remove_blank_range_between(body, previous_content_index, zh_abstract_index)
    return [
        {
            "kind": "frontmatter_declaration_separator",
            "changed": migrated or bool(removed),
            "section_owner_index": section_owner_index,
            "migrated_section_to_previous_content": migrated,
            "previous_content_index": previous_content_index,
            "zh_section_type_continuous": zh_section_type_changed,
            "zh_title_index": zh_title_index,
            "zh_title_page_break_changed": zh_title_page_break_changed,
            "zh_title_page_break_reason": zh_title_page_break_reason,
            "removed_blank_count": len(removed),
            "removed": removed,
        }
    ]


def collapse_frontmatter_empty_run_before_toc(body: ET.Element) -> list[dict[str, object]]:
    """Remove surplus blank paragraphs between English keywords and the TOC.

    Keep only the first empty section-break paragraph as the TOC page owner.
    Earlier blank paragraphs or duplicate section breaks can render as
    header/footer-only pages between the English abstract and the TOC.
    """
    points = frontmatter_topology_points(body)
    children = list(body)
    start = points["en_keyword"] + 1
    end = points["toc"]
    blank_paragraphs: list[tuple[int, ET.Element]] = []
    section_paragraphs: list[tuple[int, ET.Element]] = []
    blocked_payload_indices: list[int] = []
    for index in range(start, end):
        child = children[index]
        if child.tag != qn("p"):
            continue
        if paragraph_text(child).strip() or paragraph_has_nontext_payload(child):
            blocked_payload_indices.append(index)
            continue
        if has_section_break(child):
            section_paragraphs.append((index, child))
        else:
            blank_paragraphs.append((index, child))
    preserved_section_index = section_paragraphs[0][0] if section_paragraphs else None
    candidates: list[tuple[int, ET.Element, str]] = [
        (index, paragraph, "surplus blank before TOC; first section owner preserved")
        for index, paragraph in blank_paragraphs
    ]
    candidates.extend(
        (index, paragraph, "duplicate section break before TOC; first section owner preserved")
        for index, paragraph in section_paragraphs[1:]
    )
    removed: list[dict[str, object]] = []
    for index, paragraph, reason in sorted(candidates, key=lambda item: item[0], reverse=True):
        body.remove(paragraph)
        removed.append({"body_child_index": index, "reason": reason})
    return [
        {
            "removed_count": len(removed),
            "removed": list(reversed(removed)),
            "preserved_section_break_index": preserved_section_index,
            "removed_section_break_count": max(0, len(section_paragraphs) - 1),
            "blocked_payload_indices": blocked_payload_indices,
            "start_after_en_keyword_index": start,
            "toc_index_before": end,
        }
    ]


def normalize_front_toc_label(text: str) -> str | None:
    compact = compact_text(toc_entry_visible_label(text)).lower()
    if compact in {compact_text("摘要").lower(), compact_text("摘 要").lower()}:
        return "zh_abstract"
    if compact.startswith(compact_text("关键词").lower()):
        return "zh_keyword"
    if compact == "abstract":
        return "en_abstract"
    if compact in {"keywords", "keyword", "key words".replace(" ", "")}:
        return "en_keyword"
    return None


def first_toc_bookmark_ref(paragraph: ET.Element) -> str | None:
    for node in paragraph.findall(".//w:instrText", NS):
        match = re.search(r"_Toc[A-Za-z0-9_]+", node.text or "")
        if match:
            return match.group(0)
    return None


def frontmatter_anchor_paragraphs(body: ET.Element) -> dict[str, ET.Element]:
    return {
        "zh_abstract": find_first_paragraph(
            body,
            lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p),
        )[1],
        "zh_keyword": find_first_paragraph(
            body,
            lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p),
        )[1],
        "en_abstract": find_first_paragraph(
            body,
            lambda p: is_en_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p),
        )[1],
        "en_keyword": find_first_paragraph(
            body,
            lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p),
        )[1],
    }


def repair_frontmatter_toc_bookmarks(body: ET.Element, root: ET.Element) -> list[dict[str, object]]:
    toc_index, body_start = find_toc_and_body_start(body)
    anchors = frontmatter_anchor_paragraphs(body)
    changes: list[dict[str, object]] = []
    for child in list(body)[toc_index + 1 : body_start]:
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        surface_id = normalize_front_toc_label(paragraph_text(child))
        if surface_id is None:
            continue
        bookmark_name = first_toc_bookmark_ref(child)
        if not bookmark_name:
            changes.append(
                {
                    "kind": "frontmatter_toc_bookmark",
                    "surface_id": surface_id,
                    "toc_entry": paragraph_text(child)[:80],
                    "changed": False,
                    "reason": "toc entry has no bookmark ref",
                }
            )
            continue
        changed = ensure_bookmark(anchors[surface_id], root, bookmark_name)
        changes.append(
            {
                "kind": "frontmatter_toc_bookmark",
                "surface_id": surface_id,
                "toc_entry": paragraph_text(child)[:80],
                "bookmark": bookmark_name,
                "changed": changed,
            }
        )
    return changes


def normalize_toc_page_map_label(text: str) -> str:
    return compact_text(toc_entry_visible_label(text)).lower()


def load_toc_page_cache_map(path: Path | None) -> dict[str, str]:
    if path is None:
        raise RuntimeError("toc-page-cache requires --toc-page-map-json")
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    pairs: list[tuple[str, str]] = []
    if isinstance(payload, dict):
        for key, value in payload.items():
            if isinstance(value, dict):
                page = value.get("page_number") or value.get("toc_page_number") or value.get("page")
            else:
                page = value
            pairs.append((str(key), str(page)))
    elif isinstance(payload, list):
        for row in payload:
            if not isinstance(row, dict):
                continue
            label = row.get("text") or row.get("label") or row.get("toc_label")
            page = row.get("page_number") or row.get("toc_page_number") or row.get("logical_page")
            if label is not None and page is not None:
                pairs.append((str(label), str(page)))
    else:
        raise RuntimeError("toc-page-cache map must be a JSON object or list")
    result: dict[str, str] = {}
    for label, page in pairs:
        normalized = normalize_toc_page_map_label(label)
        if normalized and page.strip():
            result[normalized] = page.strip()
    if not result:
        raise RuntimeError("toc-page-cache map did not contain any usable label/page pairs")
    return result


def set_toc_entry_cached_page(paragraph: ET.Element, page_number: str) -> tuple[bool, str | None, str]:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        return False, None, "no text nodes"
    page_pattern = re.compile(r"^\s*(?:\d+|[ivxlcdmIVXLCDM]+)\s*$")
    for node in reversed(text_nodes):
        old_text = node.text or ""
        if page_pattern.match(old_text):
            if old_text == page_number:
                return False, old_text, "already-current"
            node.text = page_number
            return True, old_text, "replaced-page-text-node"
    trailing_pattern = re.compile(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", flags=re.IGNORECASE)
    last = text_nodes[-1]
    old_text = last.text or ""
    match = trailing_pattern.search(old_text)
    if not match:
        return False, None, "no trailing page number"
    new_text = old_text[: match.start(1)] + page_number + old_text[match.end(1) :]
    if new_text == old_text:
        return False, match.group(1), "already-current"
    last.text = new_text
    return True, match.group(1), "replaced-trailing-page-number"


def repair_toc_page_cache(body: ET.Element, toc_page_map_json: Path | None) -> list[dict[str, object]]:
    page_map = load_toc_page_cache_map(toc_page_map_json)
    toc_index, body_start = find_toc_and_body_start(body)
    changes: list[dict[str, object]] = []
    matched_labels: set[str] = set()
    for index, child in enumerate(list(body)[toc_index + 1 : body_start], start=toc_index + 1):
        if child.tag != qn("p") or not is_toc_entry_paragraph(child):
            continue
        text_before = paragraph_text(child)
        label = normalize_toc_page_map_label(text_before)
        desired_page = page_map.get(label)
        if desired_page is None:
            continue
        matched_labels.add(label)
        changed, old_page, reason = set_toc_entry_cached_page(child, desired_page)
        changes.append(
            {
                "paragraph_index": index,
                "label": label,
                "old_page": old_page,
                "new_page": desired_page,
                "changed": changed,
                "reason": reason,
                "text_before": text_before[:120],
                "text_after": paragraph_text(child)[:120],
            }
        )
    missing = sorted(set(page_map) - matched_labels)
    if missing:
        changes.append(
            {
                "kind": "unmatched_toc_page_map_labels",
                "count": len(missing),
                "labels": missing[:40],
            }
        )
    if not any(change.get("changed") for change in changes):
        raise RuntimeError("toc-page-cache did not update any TOC page cache text")
    return changes


def repair_frontmatter_template_sections(
    body: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    if template_document_root is None:
        raise RuntimeError("frontmatter-template-sections requires --template-docx")
    template_body = body_element(template_document_root)
    template_points = frontmatter_topology_points(template_body)
    changes: list[dict[str, object]] = []

    changes.extend(collapse_declaration_to_abstract_separator(body))

    points = frontmatter_topology_points(body)
    zh_to_en_count = points["en_title"] - points["zh_keyword"] - 1
    template_zh_to_en, zh_to_en_stripped_refs = clone_template_separator_range(
        template_body,
        template_points["zh_keyword"] + 1,
        template_points["en_title"],
    )
    if zh_to_en_count == 0:
        insert_clones(body, points["zh_keyword"] + 1, template_zh_to_en)
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "zh_keyword_to_en_title",
                "inserted_node_count": len(template_zh_to_en),
                "stripped_header_footer_reference_count": zh_to_en_stripped_refs,
            }
        )
    else:
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "zh_keyword_to_en_title",
                "inserted_node_count": 0,
                "reason": "separator nodes already present",
                "stripped_header_footer_reference_count": 0,
            }
        )

    points = frontmatter_topology_points(body)
    changes.append(
        migrate_text_section_to_following_blank(
            body,
            text_paragraph_index=points["zh_keyword"],
            stop_exclusive=points["en_title"],
            location="zh_keyword_to_en_title",
        )
    )

    points = frontmatter_topology_points(body)
    en_to_toc_count = points["toc"] - points["en_keyword"] - 1
    template_en_to_toc, en_to_toc_stripped_refs = clone_template_separator_range(
        template_body,
        template_points["en_keyword"] + 1,
        template_points["toc"],
    )
    if en_to_toc_count < max(3, len(template_en_to_toc) // 2):
        insert_clones(body, points["en_keyword"] + 1, template_en_to_toc)
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "en_keyword_to_toc",
                "inserted_node_count": len(template_en_to_toc),
                "prior_separator_node_count": en_to_toc_count,
                "stripped_header_footer_reference_count": en_to_toc_stripped_refs,
            }
        )
    else:
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "en_keyword_to_toc",
                "inserted_node_count": 0,
                "reason": "separator nodes already present",
                "prior_separator_node_count": en_to_toc_count,
                "stripped_header_footer_reference_count": 0,
            }
        )

    points = frontmatter_topology_points(body)
    changes.append(
        migrate_text_section_to_following_blank(
            body,
            text_paragraph_index=points["en_keyword"],
            stop_exclusive=points["toc"],
            location="en_keyword_to_toc",
        )
    )

    points = frontmatter_topology_points(body)
    template_trailing_start = template_points["last_toc_nonempty"] + 1
    template_toc_to_body, toc_to_body_stripped_refs = clone_template_separator_range(
        template_body,
        template_trailing_start,
        template_points["body_start"],
    )
    trailing_nodes = list(body)[points["last_toc_nonempty"] + 1 : points["body_start"]]
    trailing_has_section = any(has_section_break(node) for node in trailing_nodes)
    current_tail = list(body)[points["last_toc_nonempty"]]
    removed_tail_section = remove_paragraph_section_break(current_tail)
    if not trailing_has_section:
        insert_clones(body, points["last_toc_nonempty"] + 1, template_toc_to_body)
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "toc_tail_to_body_start",
                "inserted_node_count": len(template_toc_to_body),
                "removed_section_from_toc_tail": removed_tail_section,
                "stripped_header_footer_reference_count": toc_to_body_stripped_refs,
            }
        )
    else:
        changes.append(
            {
                "kind": "frontmatter_template_separator",
                "location": "toc_tail_to_body_start",
                "inserted_node_count": 0,
                "removed_section_from_toc_tail": removed_tail_section,
                "reason": "toc-to-body separator section already present",
                "stripped_header_footer_reference_count": 0,
            }
        )
    points = frontmatter_topology_points(body)
    changes.append(
        migrate_blank_section_to_previous_content(
            body,
            previous_content_index=points["last_toc_nonempty"],
            next_content_index=points["body_start"],
            location="toc_tail_to_body_start",
        )
    )
    return changes


def preserve_heading_page_break_owner(old_ppr: ET.Element | None, new_ppr: ET.Element | None) -> None:
    if old_ppr is None or new_ppr is None:
        return
    if new_ppr.find("./w:pageBreakBefore", NS) is not None:
        return
    page_break_before = old_ppr.find("./w:pageBreakBefore", NS)
    if page_break_before is not None:
        new_ppr.append(deepcopy(page_break_before))


def rewrite_first_text(paragraph: ET.Element, old: str, new: str) -> bool:
    for text_node in paragraph.findall(".//w:t", NS):
        if text_node.text and text_node.text.startswith(old):
            text_node.text = new + text_node.text[len(old) :]
            return True
    return False


def citation_numbers(text: str) -> set[int]:
    return {int(match.group(1)) for match in re.finditer(r"\[(\d+)\]", text or "")}


def bibliography_heading_index(children: list[ET.Element]) -> int:
    for index, child in enumerate(children):
        if child.tag == qn("p") and compact_text(paragraph_text(child)) == compact_text("\u53c2\u8003\u6587\u732e"):
            return index
    return len(children)


def repair_reference_supplement_residue(body: ET.Element) -> list[dict[str, object]]:
    prefix = "\u4e3a\u8865\u8db3\u53c2\u8003\u6587\u732e\u6765\u6e90\u5e76\u4fdd\u6301\u6b63\u6587\u5f15\u7528\u987a\u5e8f"
    replacement_prefix = "\u7ed3\u5408\u672c\u6587\u91c7\u7528\u7684\u76ee\u6807\u68c0\u6d4b\u3001\u59ff\u6001\u4f30\u8ba1\u548c\u9aa8\u67b6\u884c\u4e3a\u8bc6\u522b\u8def\u7ebf\uff0c\u76f8\u5173\u7814\u7a76\u53ef\u4ee5\u8fdb\u4e00\u6b65\u5f52\u7eb3\u4e3a\u4ee5\u4e0b\u51e0\u7c7b"
    changes: list[dict[str, object]] = []
    children = list(body)
    bibliography_index = bibliography_heading_index(children)
    insertion_anchor: ET.Element | None = None
    residue_paragraph: ET.Element | None = None
    residue_old_index: int | None = None
    for index, child in enumerate(children):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if text.startswith(prefix):
            residue_paragraph = child
            residue_old_index = index
            continue
        if index < bibliography_index and 11 in citation_numbers(text):
            insertion_anchor = child
    if residue_paragraph is not None and residue_old_index is not None:
        original_text = paragraph_text(residue_paragraph).strip()
        rewrite_first_text(residue_paragraph, prefix, replacement_prefix)
        body.remove(residue_paragraph)
        current_children = list(body)
        if insertion_anchor is not None and insertion_anchor in current_children:
            insert_at = current_children.index(insertion_anchor) + 1
            reason = "repair-note wording replaced and citation paragraph moved after final pre-existing body citation"
        else:
            bibliography = next(
                (
                    item
                    for item in current_children
                    if item.tag == qn("p") and compact_text(paragraph_text(item)) == compact_text("\u53c2\u8003\u6587\u732e")
                ),
                None,
            )
            insert_at = current_children.index(bibliography) if bibliography is not None else len(current_children)
            reason = "repair-note wording replaced and citation paragraph kept before bibliography"
        body.insert(insert_at, residue_paragraph)
        changes.append(
            {
                "old_index": residue_old_index,
                "new_index": insert_at,
                "text_before": original_text[:200],
                "text_after": paragraph_text(residue_paragraph).strip()[:200],
                "citation_numbers_preserved": sorted(citation_numbers(paragraph_text(residue_paragraph))),
                "reason": reason,
            }
        )
    return changes


def remove_unanchored_empty_frontmatter_paragraphs(body: ET.Element) -> list[dict[str, object]]:
    """Remove blank spacer paragraphs that can push declaration signatures onto a residue page."""
    changes: list[dict[str, object]] = []
    in_declaration_block = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        compact = compact_text(text)
        if compact in {compact_text("\u72ec\u521b\u6027\u58f0\u660e"), compact_text("\u5173\u4e8e\u8bba\u6587\u4f7f\u7528\u6388\u6743\u7684\u8bf4\u660e")}:
            in_declaration_block = True
            continue
        if compact in {compact_text("\u6458\u8981"), compact_text("ABSTRACT")}:
            break
        if not in_declaration_block or text:
            continue
        ppr = paragraph_ppr(child)
        has_section_break = ppr is not None and ppr.find("./w:sectPr", NS) is not None
        has_protected_anchor = any(
            child.find(f".//w:{name}", NS) is not None
            for name in (
                "bookmarkStart",
                "bookmarkEnd",
                "commentRangeStart",
                "commentRangeEnd",
                "commentReference",
                "fldChar",
                "instrText",
                "drawing",
                "object",
            )
        )
        if has_section_break or has_protected_anchor:
            continue
        body.remove(child)
        changes.append(
            {
                "old_index": index,
                "reason": "blank declaration spacer removed to keep signature block with declaration page",
            }
        )
    return changes


def tighten_declaration_title_spacing(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    target_titles = {
        compact_text("\u72ec\u521b\u6027\u58f0\u660e"),
        compact_text("\u5173\u4e8e\u8bba\u6587\u4f7f\u7528\u6388\u6743\u7684\u8bf4\u660e"),
    }
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if compact_text(text) not in target_titles:
            continue
        ppr = ensure_ppr(child)
        spacing = ppr.find("./w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(qn("spacing"))
            ppr.append(spacing)
        before = dict(spacing.attrib)
        changed = False
        desired = {
            qn("after"): "180",
            qn("line"): "360",
            qn("lineRule"): "auto",
        }
        for key, value in desired.items():
            if spacing.get(key) != value:
                spacing.set(key, value)
                changed = True
        for key in (qn("afterLines"), qn("beforeLines")):
            if key in spacing.attrib:
                del spacing.attrib[key]
                changed = True
        if changed:
            changes.append(
                {
                    "index": index,
                    "text": text,
                    "spacing_before": {k.split('}')[-1]: v for k, v in before.items()},
                    "spacing_after": {k.split('}')[-1]: v for k, v in spacing.attrib.items()},
                    "reason": "declaration title spacing tightened to prevent a signature-only residue page",
                }
            )
    return changes


def repair_frontmatter_signature_residual(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    changes.extend(remove_unanchored_empty_frontmatter_paragraphs(body))
    changes.extend(tighten_declaration_title_spacing(body))
    return changes


DECLARATION_TITLE_COMPACTS = {
    compact_text("独创性声明"),
    compact_text("学位论文原创性声明"),
}


def is_declaration_title_text(text: str) -> bool:
    return compact_text(text) in DECLARATION_TITLE_COMPACTS


def is_cover_date_line(text: str) -> bool:
    compact = compact_text(text)
    return bool("年" in compact and "月" in compact and "日" in compact and len(compact) <= 24)


def extract_value_after_label(text: str, labels: tuple[str, ...]) -> str:
    compact = compact_text(text)
    for label in labels:
        label_compact = compact_text(label)
        if compact.startswith(label_compact):
            return compact[len(label_compact) :].strip()
    return ""


def likely_cover_title_text(text: str) -> bool:
    compact = compact_text(text)
    if len(compact) < 8:
        return False
    if any(token in compact for token in ("学士学位论文", "河北北方学院", "原创性声明", "版权使用授权书")):
        return False
    if any(compact.startswith(compact_text(label)) for label in ("作者姓名", "姓名", "学号", "院系", "专业", "指导教师")):
        return False
    return bool("系统" in compact or "研究" in compact or "设计" in compact or "分析" in compact)


def extract_cover_metadata(body: ET.Element, declaration_index: int) -> dict[str, str]:
    children = list(body)
    metadata = {
        "title": "",
        "name": "",
        "student_id": "",
        "department": "",
        "major": "",
        "advisor": "",
        "date": "",
    }
    for child in children[:declaration_index]:
        text = paragraph_text(child).strip()
        if not text:
            continue
        if child.tag == qn("tbl") and not metadata["title"] and likely_cover_title_text(text):
            metadata["title"] = text
        if not metadata["title"] and likely_cover_title_text(text):
            metadata["title"] = text
        metadata["name"] = metadata["name"] or extract_value_after_label(text, ("作者姓名", "姓    名", "姓名"))
        metadata["student_id"] = metadata["student_id"] or extract_value_after_label(text, ("学    号", "学号"))
        metadata["department"] = metadata["department"] or extract_value_after_label(text, ("院    系", "院系", "学院"))
        metadata["major"] = metadata["major"] or extract_value_after_label(text, ("专    业", "专业"))
        metadata["advisor"] = metadata["advisor"] or extract_value_after_label(text, ("指导教师", "导师"))
        if not metadata["date"] and is_cover_date_line(text):
            metadata["date"] = text
    if not metadata["title"]:
        raise RuntimeError("attachment8-cover-cleanup could not locate a cover thesis title before declaration")
    return metadata


def clone_template_cover_title_paragraph(template_document_root: ET.Element | None, title: str) -> ET.Element:
    if template_document_root is None:
        raise RuntimeError("attachment8-cover-cleanup requires --template-docx")
    template_body = body_element(template_document_root)
    donor: ET.Element | None = None
    for child in list(template_body):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child)
        if compact_text(text).startswith(compact_text("题目：")):
            donor = deepcopy(child)
            break
    if donor is None:
        raise RuntimeError("attachment8-cover-cleanup could not locate `题目：` donor paragraph in template")
    set_paragraph_visible_text_preserving_first_run(donor, f"题目：{title}")
    return donor


def paragraph_has_only_page_break_or_blank(paragraph: ET.Element) -> bool:
    if paragraph_text(paragraph).strip():
        return False
    return True


def make_page_break_paragraph() -> ET.Element:
    paragraph = ET.Element(qn("p"))
    run = ET.SubElement(paragraph, qn("r"))
    br = ET.SubElement(run, qn("br"))
    br.set(qn("type"), "page")
    return paragraph


def clear_paragraph_payload(paragraph: ET.Element) -> None:
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)


def underline_rpr(rpr: ET.Element | None) -> ET.Element | None:
    if rpr is None:
        return None
    cloned = deepcopy(rpr)
    underline = cloned.find("./w:u", NS)
    if underline is None:
        underline = ET.Element(qn("u"))
        cloned.append(underline)
    underline.set(qn("val"), "single")
    return cloned


def first_text_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if not run_text(run).strip():
            continue
        rpr = run.find("./w:rPr", NS)
        return deepcopy(rpr) if rpr is not None else None
    return None


def rebuild_cover_field_paragraph(paragraph: ET.Element, label: str, value: str, *, pad_left: int = 7, pad_right: int = 7) -> None:
    label_rpr = first_text_run_properties(paragraph)
    value_rpr = underline_rpr(label_rpr)
    clear_paragraph_payload(paragraph)
    append_text_run_with_rpr(paragraph, label, label_rpr)
    append_text_run_with_rpr(paragraph, " " * pad_left, value_rpr)
    append_text_run_with_rpr(paragraph, value, value_rpr)
    append_text_run_with_rpr(paragraph, " " * pad_right, value_rpr)


def cover_child_matches_text(child: ET.Element, expected: str) -> bool:
    return child.tag == qn("p") and compact_text(paragraph_text(child)) == compact_text(expected)


def first_cover_drawing_index(body: ET.Element) -> int | None:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and (
            child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None
        ):
            return index
    return None


def first_cover_degree_index(body: ET.Element) -> int | None:
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        compact = compact_text(paragraph_text(child))
        if compact in {compact_text("学士学位论文"), compact_text("学士学位论文（设计）")}:
            return index
    return None


def first_cover_title_label_index(body: ET.Element) -> int | None:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and compact_text(paragraph_text(child)).startswith(compact_text("题目：")):
            return index
    return None


def first_cover_label_index(body: ET.Element, label: str) -> int | None:
    label_compact = compact_text(label)
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and compact_text(paragraph_text(child)).startswith(label_compact):
            return index
    return None


def remove_surplus_blank_paragraphs(
    body: ET.Element,
    *,
    start_exclusive: int,
    end_exclusive: int,
    keep_count: int,
    reason: str,
) -> list[dict[str, object]]:
    blank_indices = [
        index
        for index in range(start_exclusive + 1, end_exclusive)
        if is_blank_paragraph_without_payload(list(body)[index])
    ]
    remove_count = max(0, len(blank_indices) - keep_count)
    removed: list[dict[str, object]] = []
    for index in reversed(blank_indices[-remove_count:]):
        child = list(body)[index]
        body.remove(child)
        removed.append({"body_child_index": index, "reason": reason})
    return list(reversed(removed))


def clone_attachment8_cover_prefix(template_document_root: ET.Element | None) -> list[ET.Element]:
    if template_document_root is None:
        raise RuntimeError("attachment8-cover-template-layout requires --template-docx")
    template_body = body_element(template_document_root)
    clones: list[ET.Element] = []
    for child in list(template_body):
        if child.tag == qn("p") and (
            child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None
        ):
            break
        clones.append(deepcopy(child))
    if len(clones) < 2:
        raise RuntimeError("attachment8-cover-template-layout could not locate Attachment 8 cover prefix")
    return clones


def clone_attachment8_cover_blank(template_document_root: ET.Element | None, template_index: int) -> ET.Element:
    if template_document_root is None:
        raise RuntimeError("attachment8-cover-template-layout requires --template-docx")
    template_body = body_element(template_document_root)
    children = list(template_body)
    if template_index >= len(children):
        raise RuntimeError("attachment8-cover-template-layout template blank index is out of range")
    clone = deepcopy(children[template_index])
    if clone.tag != qn("p"):
        clone = ET.Element(qn("p"))
    clear_paragraph_payload(clone)
    return clone


def clone_attachment8_cover_section_owner(template_document_root: ET.Element | None) -> ET.Element:
    if template_document_root is None:
        raise RuntimeError("attachment8-cover-template-layout requires --template-docx")
    template_body = body_element(template_document_root)
    template_end = cover_section_end_index(template_body)
    owner = deepcopy(list(template_body)[template_end])
    clear_paragraph_payload(owner)
    if not has_section_break(owner):
        raise RuntimeError("attachment8-cover-template-layout template cover section owner has no section break")
    strip_section_header_footer_references(owner)
    return owner


def repair_attachment8_cover_template_layout(
    body: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    """Align the cleaned Attachment 8 cover to the template skeleton without copying callout prose."""
    declaration_index = first_declaration_index(body)
    metadata = extract_cover_metadata(body, declaration_index)
    changes: list[dict[str, object]] = []

    if not (
        len(list(body)) >= 2
        and cover_child_matches_text(list(body)[0], "附件8")
        and "河北北方学院学士学位论文模板" in paragraph_text(list(body)[1])
    ):
        prefix = clone_attachment8_cover_prefix(template_document_root)
        insert_clones(body, 0, prefix)
        changes.append({"kind": "attachment8_cover_prefix_inserted", "inserted_node_count": len(prefix)})
    else:
        changes.append({"kind": "attachment8_cover_prefix_inserted", "inserted_node_count": 0, "reason": "already present"})

    drawing_index = first_cover_drawing_index(body)
    degree_index = first_cover_degree_index(body)
    if drawing_index is None or degree_index is None:
        raise RuntimeError("attachment8-cover-template-layout could not locate cover logo or degree title")
    removed = remove_surplus_blank_paragraphs(
        body,
        start_exclusive=drawing_index,
        end_exclusive=degree_index,
        keep_count=1,
        reason="trim cover logo-to-degree spacer to Attachment 8 skeleton",
    )
    if removed:
        changes.append({"kind": "cover_logo_to_degree_spacers_trimmed", "removed": removed})

    title_index = first_cover_title_label_index(body)
    name_index = first_cover_label_index(body, "姓    名")
    if title_index is None or name_index is None:
        raise RuntimeError("attachment8-cover-template-layout could not locate cover title/name block")
    removed = remove_surplus_blank_paragraphs(
        body,
        start_exclusive=title_index,
        end_exclusive=name_index,
        keep_count=2,
        reason="trim title-to-name spacer to Attachment 8 skeleton",
    )
    if removed:
        changes.append({"kind": "cover_title_to_name_spacers_trimmed", "removed": removed})

    field_specs = [
        ("姓    名", metadata["name"], 7, 7),
        ("学    号", metadata["student_id"], 7, 7),
        ("院    系", metadata["department"], 3, 2),
        ("专    业", metadata["major"], 2, 2),
        ("指导教师", metadata["advisor"], 5, 8),
    ]
    for label, value, pad_left, pad_right in field_specs:
        index = first_cover_label_index(body, label)
        if index is None:
            continue
        before = paragraph_text(list(body)[index])
        rebuild_cover_field_paragraph(list(body)[index], label, value, pad_left=pad_left, pad_right=pad_right)
        changes.append(
            {
                "kind": "cover_field_line_rebuilt",
                "label": label,
                "body_child_index": index,
                "old_text": before,
                "new_text": paragraph_text(list(body)[index]),
            }
        )

    date_index = None
    for index, child in enumerate(list(body)[: first_declaration_index(body)]):
        if child.tag == qn("p") and is_cover_date_line(paragraph_text(child)):
            date_index = index
    if date_index is None:
        raise RuntimeError("attachment8-cover-template-layout could not locate cover date")
    while date_index < 22:
        body.insert(date_index, clone_attachment8_cover_blank(template_document_root, 18))
        changes.append({"kind": "cover_date_vertical_spacer_inserted", "body_child_index": date_index})
        date_index += 1
    while date_index > 22 and is_blank_paragraph_without_payload(list(body)[date_index - 1]):
        body.remove(list(body)[date_index - 1])
        date_index -= 1
        changes.append({"kind": "cover_date_vertical_spacer_removed", "body_child_index": date_index})

    declaration_index = first_declaration_index(body)
    section_owner_index = declaration_index - 1
    section_owner = clone_attachment8_cover_section_owner(template_document_root)
    existing_owner = list(body)[section_owner_index]
    if existing_owner.tag == qn("p") and (has_page_break(existing_owner) or is_blank_paragraph_without_payload(existing_owner)):
        body.remove(existing_owner)
        body.insert(section_owner_index, section_owner)
        changes.append({"kind": "cover_section_owner_replaced", "body_child_index": section_owner_index})
    else:
        body.insert(declaration_index, section_owner)
        changes.append({"kind": "cover_section_owner_inserted", "body_child_index": declaration_index})
    return changes


def replace_author_label_text(paragraph: ET.Element) -> bool:
    changed = False
    for text_node in paragraph.findall(".//w:t", NS):
        current = text_node.text or ""
        if "作者姓名" in current:
            text_node.text = current.replace("作者姓名", "姓    名", 1)
            changed = True
    return changed


def repair_attachment8_cover_cleanup(
    body: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    declaration_index = first_declaration_index(body)
    metadata = extract_cover_metadata(body, declaration_index)
    children = list(body)
    changes: list[dict[str, object]] = []

    degree_index: int | None = None
    title_table_index: int | None = None
    date_index: int | None = None
    for index, child in enumerate(children[:declaration_index]):
        text = paragraph_text(child).strip()
        compact = compact_text(text)
        if child.tag == qn("p") and compact in {compact_text("学士学位论文"), compact_text("学士学位论文（设计）")}:
            degree_index = index
            if text != "学士学位论文（设计）":
                set_paragraph_visible_text_preserving_first_run(child, "学士学位论文（设计）")
                changes.append({"kind": "cover_degree_title", "body_child_index": index, "new_text": "学士学位论文（设计）"})
        if child.tag == qn("tbl") and title_table_index is None and compact_text(metadata["title"]) in compact:
            title_table_index = index
        if child.tag == qn("p") and replace_author_label_text(child):
            changes.append({"kind": "cover_author_label", "body_child_index": index, "new_label": "姓    名"})
        if child.tag == qn("p") and is_cover_date_line(text):
            date_index = index

    insert_at = title_table_index if title_table_index is not None else (degree_index + 1 if degree_index is not None else 0)
    title_paragraph = clone_template_cover_title_paragraph(template_document_root, metadata["title"])
    if title_table_index is not None:
        body.remove(children[title_table_index])
        changes.append({"kind": "cover_old_title_table_removed", "body_child_index": title_table_index})
    body.insert(insert_at, title_paragraph)
    changes.append({"kind": "cover_title_label_inserted", "body_child_index": insert_at, "text": f"题目：{metadata['title']}"})

    children = list(body)
    declaration_index = first_declaration_index(body)
    if date_index is None:
        date_index = max(0, declaration_index - 1)
    # Remove the legacy standalone Chinese/English title page that the new Attachment 8 template deleted.
    removed: list[dict[str, object]] = []
    metadata_title_compact = compact_text(metadata["title"])
    for index in range(declaration_index - 1, date_index, -1):
        child = list(body)[index]
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        text_compact = compact_text(text)
        is_duplicate_title_line = bool(
            metadata_title_compact
            and metadata_title_compact in text_compact
            and ("\u9898\u76ee" in text_compact or text_compact == metadata_title_compact)
        )
        if text and not is_duplicate_title_line and not re.search(r"[A-Za-z]{4,}", text):
            continue
        body.remove(child)
        removed.append(
            {
                "body_child_index": index,
                "text": text[:120],
                "duplicate_title_line": is_duplicate_title_line,
            }
        )
    removed.reverse()
    if removed:
        changes.append({"kind": "legacy_standalone_title_page_removed", "removed": removed})

    declaration_index = first_declaration_index(body)
    previous = list(body)[declaration_index - 1] if declaration_index > 0 else None
    if previous is None or previous.tag != qn("p") or not has_page_break(previous):
        body.insert(declaration_index, make_page_break_paragraph())
        changes.append({"kind": "cover_to_declaration_page_break_inserted", "body_child_index": declaration_index})
    return changes


def is_cover_metadata_line(text: str) -> bool:
    compact = compact_text(text)
    return any(
        compact.startswith(prefix)
        for prefix in (
            "\u4e13\u4e1a",
            "\u59d3\u540d",
            "\u5b66\u53f7",
            "\u6307\u5bfc\u6559\u5e08",
            "\u5b8c\u6210\u65f6\u95f4",
        )
    )


def disable_true_page_break_before(paragraph: ET.Element) -> tuple[bool, str]:
    ppr = paragraph_ppr(paragraph)
    if ppr is None:
        return False, "no-paragraph-properties"
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        return False, "no-pageBreakBefore"
    current = page_break.get(qn("val"))
    if current in {"0", "false", "False", "off", "OFF"}:
        return False, "already-disabled"
    page_break.set(qn("val"), "0")
    return True, "disabled-true-pageBreakBefore"


def ensure_true_page_break_before(paragraph: ET.Element) -> tuple[bool, str]:
    ppr = ensure_ppr(paragraph)
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is None:
        page_break = ET.Element(qn("pageBreakBefore"))
        style = ppr.find("./w:pStyle", NS)
        insert_at = list(ppr).index(style) + 1 if style is not None else 0
        ppr.insert(insert_at, page_break)
        return True, "added-true-pageBreakBefore"
    current = page_break.get(qn("val"))
    if current in {None, "1", "true", "True", "on", "ON"}:
        return False, "already-enabled"
    page_break.attrib.pop(qn("val"), None)
    return True, "enabled-existing-pageBreakBefore"


def repair_cover_metadata_pagebreaks(body: ET.Element) -> list[dict[str, object]]:
    """Prevent individual school cover metadata lines from creating stray pages."""
    changes: list[dict[str, object]] = []
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if compact_text(text) == compact_text("\u72ec\u521b\u6027\u58f0\u660e"):
            break
        if not is_cover_metadata_line(text):
            continue
        changed, reason = disable_true_page_break_before(child)
        changes.append(
            {
                "paragraph_index": index,
                "text": text[:120],
                "changed": changed,
                "reason": reason,
            }
        )
    if not changes:
        raise RuntimeError("cover-metadata-pagebreaks could not locate cover metadata lines before the declaration page")
    return changes


def repair_cover_metadata_page_owner(body: ET.Element) -> list[dict[str, object]]:
    """Make the first metadata line own the template's second cover page."""
    changes: list[dict[str, object]] = []
    metadata_seen = 0
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if compact_text(text) == compact_text("\u72ec\u521b\u6027\u58f0\u660e"):
            break
        if not is_cover_metadata_line(text):
            continue
        metadata_seen += 1
        if metadata_seen == 1:
            changed, reason = ensure_true_page_break_before(child)
            owner = "metadata-page-start"
        else:
            changed, reason = disable_true_page_break_before(child)
            owner = "same-metadata-page"
        changes.append(
            {
                "paragraph_index": index,
                "text": text[:120],
                "changed": changed,
                "reason": reason,
                "owner": owner,
            }
        )
    if metadata_seen != 5:
        raise RuntimeError(f"cover-metadata-page-owner expected 5 metadata lines before declaration, found {metadata_seen}")
    return changes


def is_confidential_authorization_placeholder(text: str) -> bool:
    compact = compact_text(text)
    return (
        "\u4fdd\u5bc6\u7684\u5b66\u4f4d\u8bba\u6587\u5728" in compact
        and "\u89e3\u5bc6\u540e\u9002\u7528\u672c\u6388\u6743\u4e66" in compact
        and ("_" in text or "\uff3f" in text)
    )


def repair_cover_confidential_placeholder(body: ET.Element) -> list[dict[str, object]]:
    """Replace the unresolved confidentiality year placeholder on the authorization page."""
    replacement = "\u672c\u8bba\u6587\u4e0d\u6d89\u53ca\u4fdd\u5bc6\u5185\u5bb9\u3002"
    changes: list[dict[str, object]] = []
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_zh_abstract_label(text) or is_toc_heading(text):
            break
        if not is_confidential_authorization_placeholder(text):
            continue
        set_paragraph_visible_text_preserving_first_run(child, replacement)
        changes.append({"paragraph_index": index, "old_text": text, "new_text": replacement})
    if not changes:
        raise RuntimeError("cover-confidential-placeholder could not locate unresolved confidentiality placeholder")
    return changes


def first_declaration_index(body: ET.Element) -> int:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and is_declaration_title_text(paragraph_text(child)):
            return index
    raise RuntimeError("could not locate declaration page after cover")


def cover_section_end_index(body: ET.Element) -> int:
    declaration_index = first_declaration_index(body)
    for index in range(declaration_index - 1, -1, -1):
        child = list(body)[index]
        if has_section_break(child):
            return index
    raise RuntimeError("could not locate cover section break before declaration page")


def clone_ppr_preserving_section(template_paragraph: ET.Element, target_paragraph: ET.Element) -> ET.Element | None:
    template_ppr = paragraph_ppr(template_paragraph)
    if template_ppr is None:
        return None
    target_ppr = paragraph_ppr(target_paragraph)
    target_sect = target_ppr.find("./w:sectPr", NS) if target_ppr is not None else None
    cloned = deepcopy(template_ppr)
    cloned_sect = cloned.find("./w:sectPr", NS)
    if cloned_sect is not None:
        cloned.remove(cloned_sect)
    if target_sect is not None:
        cloned.append(deepcopy(target_sect))
    return cloned


def replace_paragraph_ppr(paragraph: ET.Element, new_ppr: ET.Element | None) -> None:
    old_ppr = paragraph_ppr(paragraph)
    if old_ppr is not None:
        paragraph.remove(old_ppr)
    if new_ppr is not None:
        paragraph.insert(0, new_ppr)


def repair_cover_template_paragraph_properties(
    body: ET.Element,
    template_document_root: ET.Element | None,
) -> list[dict[str, object]]:
    """Copy only cover paragraph properties from the locked template."""
    if template_document_root is None:
        raise RuntimeError("cover-template-ppr requires --template-docx")
    template_body = body_element(template_document_root)
    target_end = cover_section_end_index(body)
    template_end = cover_section_end_index(template_body)
    target_children = list(body)
    template_children = list(template_body)
    if target_end != template_end:
        raise RuntimeError(
            f"cover-template-ppr expected matching cover paragraph counts, got target={target_end + 1}, template={template_end + 1}"
        )
    changes: list[dict[str, object]] = []
    for index in range(target_end + 1):
        target = target_children[index]
        template = template_children[index]
        if target.tag != qn("p") or template.tag != qn("p"):
            continue
        before = ET.tostring(paragraph_ppr(target), encoding="unicode") if paragraph_ppr(target) is not None else ""
        new_ppr = clone_ppr_preserving_section(template, target)
        after = ET.tostring(new_ppr, encoding="unicode") if new_ppr is not None else ""
        changed = before != after
        if changed:
            replace_paragraph_ppr(target, new_ppr)
        changes.append(
            {
                "paragraph_index": index,
                "text": paragraph_text(target)[:120],
                "changed": changed,
                "template_text": paragraph_text(template)[:120],
                "preserved_target_section_break": has_section_break(target),
            }
        )
    return changes


def set_heading_line_spacing(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    before = dict(spacing.attrib)
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    for attr in ("beforeLines", "afterLines"):
        spacing.attrib.pop(qn(attr), None)
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(qn("ind"))
        ppr.append(ind)
    ind_before = dict(ind.attrib)
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    for attr in ("hanging", "firstLineChars", "leftChars", "rightChars", "hangingChars"):
        ind.attrib.pop(qn(attr), None)
    return before != dict(spacing.attrib) or ind_before != dict(ind.attrib)


def repair_body_heading_line_spacing(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_seen = False
    main_body_started = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if not text:
            continue
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen or is_toc_entry_paragraph(child):
            continue
        compact = compact_text(text)
        if compact in {compact_text("\u53c2\u8003\u6587\u732e"), compact_text("\u81f4\u8c22"), compact_text("\u81f4  \u8c22")}:
            break
        level = paragraph_heading_level(child)
        if level is None:
            continue
        main_body_started = True
        if set_heading_line_spacing(child):
            changes.append({"paragraph_index": index, "level": level, "text": text[:120]})
    if not main_body_started:
        raise RuntimeError("body-heading-line-spacing could not locate body headings after TOC")
    return changes


def clone_heading_ppr_preserving_layout_owners(donor: ET.Element, target: ET.Element) -> ET.Element | None:
    donor_ppr = paragraph_ppr(donor)
    target_ppr = paragraph_ppr(target)
    preserved: list[ET.Element] = []
    if target_ppr is not None:
        for tag in ("pageBreakBefore", "sectPr"):
            node = target_ppr.find(f"./w:{tag}", NS)
            if node is not None:
                preserved.append(deepcopy(node))
    if donor_ppr is not None:
        cloned = deepcopy(donor_ppr)
    elif target_ppr is not None and target_ppr.find("./w:pStyle", NS) is not None:
        cloned = ET.Element(qn("pPr"))
        cloned.append(deepcopy(target_ppr.find("./w:pStyle", NS)))
    elif preserved:
        cloned = ET.Element(qn("pPr"))
    else:
        return None
    for tag in ("pageBreakBefore", "sectPr"):
        old = cloned.find(f"./w:{tag}", NS)
        if old is not None:
            cloned.remove(old)
    num_pr = cloned.find("./w:numPr", NS)
    if num_pr is not None:
        cloned.remove(num_pr)
    for node in preserved:
        cloned.append(node)
    return cloned


def apply_visible_run_rpr_from_donor_or_clear(paragraph: ET.Element, donor_rpr: ET.Element | None) -> int:
    touched = 0
    for run in paragraph.findall("./w:r", NS):
        if not run_text(run).strip():
            continue
        old_rpr = run.find("./w:rPr", NS)
        before = ET.tostring(old_rpr, encoding="unicode") if old_rpr is not None else ""
        if old_rpr is not None:
            run.remove(old_rpr)
        if donor_rpr is not None:
            run.insert(0, deepcopy(donor_rpr))
        after_rpr = run.find("./w:rPr", NS)
        after = ET.tostring(after_rpr, encoding="unicode") if after_rpr is not None else ""
        if before != after:
            touched += 1
    return touched


def collect_body_heading_donors(body: ET.Element, styles_root: ET.Element | None = None) -> dict[int, ET.Element]:
    children = list(body)
    try:
        _toc_index, body_start = find_toc_and_body_start(body)
    except RuntimeError:
        body_start = 0
    donors: dict[int, ET.Element] = {}
    for child in children[body_start:]:
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if not text:
            continue
        if compact_text(text) in {compact_text(REFERENCES_LABEL), compact_text(ACKNOWLEDGEMENT_LABEL)}:
            break
        if is_toc_entry_paragraph(child):
            continue
        level = paragraph_heading_level_for_template_donor(child, styles_root)
        if level is None or level not in (1, 2, 3):
            continue
        donors.setdefault(level, child)
    return donors


def copy_template_heading_styles(
    styles_root: ET.Element,
    template_styles_root: ET.Element,
    donors: dict[int, ET.Element],
) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    for level, donor in sorted(donors.items()):
        style_id = paragraph_style_id(donor)
        if not style_id:
            continue
        donor_style = style_by_id(template_styles_root, style_id)
        if donor_style is None:
            continue
        action = replace_or_append_style_definition(styles_root, donor_style)
        changes.append(
            {
                "kind": "body_heading_template_style",
                "level": level,
                "style_id": style_id,
                "action": action,
            }
        )
    return changes


def repair_body_heading_template_baseline(
    body: ET.Element,
    styles_root: ET.Element,
    template_document_root: ET.Element | None,
    template_styles_root: ET.Element | None,
) -> list[dict[str, object]]:
    if template_document_root is None or template_styles_root is None:
        raise RuntimeError("body-heading-template-baseline requires --template-docx")
    template_body = body_element(template_document_root)
    donors = collect_body_heading_donors(template_body, template_styles_root)
    if not donors:
        raise RuntimeError("body-heading-template-baseline could not locate template body heading donors")
    changes = copy_template_heading_styles(styles_root, template_styles_root, donors)
    toc_seen = False
    main_body_started = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen or is_toc_entry_paragraph(child):
            continue
        compact = compact_text(text)
        if compact in {compact_text(REFERENCES_LABEL), compact_text(ACKNOWLEDGEMENT_LABEL)}:
            break
        level = paragraph_heading_level(child)
        if level is None or level not in donors:
            continue
        main_body_started = True
        donor = donors[level]
        before_ppr = ET.tostring(paragraph_ppr(child), encoding="unicode") if paragraph_ppr(child) is not None else ""
        new_ppr = clone_heading_ppr_preserving_layout_owners(donor, child)
        replace_paragraph_ppr(child, new_ppr)
        after_ppr = ET.tostring(paragraph_ppr(child), encoding="unicode") if paragraph_ppr(child) is not None else ""
        donor_rpr = donor_run_rpr_or_none(donor)
        touched_runs = apply_visible_run_rpr_from_donor_or_clear(child, donor_rpr)
        if before_ppr != after_ppr or touched_runs:
            changes.append(
                {
                    "kind": "body_heading_template_baseline",
                    "body_child_index": index,
                    "level": level,
                    "text": text[:120],
                    "template_text": paragraph_text(donor).strip()[:120],
                    "paragraph_properties_changed": before_ppr != after_ppr,
                    "visible_runs_replayed": touched_runs,
                    "donor_run_rpr_applied": donor_rpr is not None,
                    "preserved_pagebreak_or_section_owner": (
                        (paragraph_ppr(child) is not None)
                        and (
                            paragraph_ppr(child).find("./w:pageBreakBefore", NS) is not None
                            or paragraph_ppr(child).find("./w:sectPr", NS) is not None
                        )
                    ),
                }
            )
    if not main_body_started:
        raise RuntimeError("body-heading-template-baseline could not locate body headings after TOC")
    return changes


def set_body_image_holder_metrics(paragraph: ET.Element, *, style_id: str) -> bool:
    ppr = ensure_ppr(paragraph)
    before = ET.tostring(ppr, encoding="unicode")
    set_paragraph_style(paragraph, style_id)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("before"), "120")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    for attr in ("beforeLines", "afterLines"):
        spacing.attrib.pop(qn(attr), None)
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(qn("ind"))
        ppr.append(ind)
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    ind.set(qn("firstLineChars"), "0")
    ind.set(qn("leftChars"), "0")
    ind.set(qn("rightChars"), "0")
    for attr in ("hanging", "hangingChars"):
        ind.attrib.pop(qn(attr), None)
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.Element(qn("jc"))
        ppr.append(jc)
    jc.set(qn("val"), "center")
    if ppr.find("./w:keepNext", NS) is None:
        ppr.append(ET.Element(qn("keepNext")))
    return before != ET.tostring(ppr, encoding="unicode")


def is_body_terminal_heading(text: str) -> bool:
    compact = compact_text(text)
    return compact in {
        compact_text("\u53c2\u8003\u6587\u732e"),
        compact_text("\u81f4\u8c22"),
        compact_text("\u81f4  \u8c22"),
        "References",
        "Acknowledgements",
        "Acknowledgments",
    }


def repair_body_image_holder_spacing(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_seen = False
    main_body_started = False
    style_id, style_changed = ensure_image_holder_style(styles_root)
    if style_changed:
        changes.append({"kind": "style_created_or_normalized", "style_id": style_id})
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen or is_toc_entry_paragraph(child):
            continue
        if text and is_body_terminal_heading(text):
            break
        if text and paragraph_heading_level(child) is not None:
            main_body_started = True
        if not main_body_started:
            continue
        has_picture = child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None
        if not has_picture:
            continue
        if set_body_image_holder_metrics(child, style_id=style_id):
            changes.append({"kind": "image_holder", "paragraph_index": index, "text": text[:120], "style_id": style_id})
    if not main_body_started:
        raise RuntimeError("body-image-holder-spacing could not locate the main body after TOC")
    return changes


def is_formal_caption_text(text: str) -> bool:
    stripped = (text or "").strip()
    return bool(re.match(r"^(?:\u56fe|\u8868)\s*\d+(?:[-.\uff0d\uff0e]\d+)*\s+\S", stripped))


def set_caption_direct_format(paragraph: ET.Element) -> bool:
    before = ET.tostring(paragraph, encoding="unicode")
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("before"), "120")
    spacing.set(qn("after"), "120")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(qn("ind"))
        ppr.append(ind)
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    for attr in ("hanging", "firstLineChars", "leftChars", "rightChars", "hangingChars"):
        ind.attrib.pop(qn(attr), None)
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.Element(qn("jc"))
        ppr.append(jc)
    jc.set(qn("val"), "center")
    for run in paragraph.findall("./w:r", NS):
        if not paragraph_text(run).strip():
            continue
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(qn("rPr"))
            run.insert(0, rpr)
        size = rpr.find("./w:sz", NS)
        if size is None:
            size = ET.Element(qn("sz"))
            rpr.append(size)
        size.set(qn("val"), "21")
        size_cs = rpr.find("./w:szCs", NS)
        if size_cs is None:
            size_cs = ET.Element(qn("szCs"))
            rpr.append(size_cs)
        size_cs.set(qn("val"), "21")
    return before != ET.tostring(paragraph, encoding="unicode")


def repair_caption_direct_format(body: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_seen = False
    main_body_started = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen or is_toc_entry_paragraph(child):
            continue
        if text and is_body_terminal_heading(text):
            break
        if text and paragraph_heading_level(child) is not None:
            main_body_started = True
        if not main_body_started or not is_formal_caption_text(text):
            continue
        if set_caption_direct_format(child):
            changes.append({"paragraph_index": index, "text": text[:120]})
    return changes


def style_run_rpr(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    style = style_by_id(styles_root, style_id)
    if style is None:
        return None
    rpr = style.find("./w:rPr", NS)
    return deepcopy(rpr) if rpr is not None else None


def repair_title_run_format(
    body: ET.Element,
    styles_root: ET.Element,
    *,
    target_texts: set[str] | None = None,
) -> list[dict[str, object]]:
    """Materialize template-owned title run formatting on key title paragraphs."""
    if target_texts is None:
        target_texts = {
            compact_text("\u76ee\u5f55"),
            compact_text("\u53c2\u8003\u6587\u732e"),
            compact_text("\u81f4\u8c22"),
        }
    changes: list[dict[str, object]] = []
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if compact_text(text) not in target_texts:
            continue
        style_id = paragraph_style_id(child) or "Heading1"
        donor_rpr = style_run_rpr(styles_root, style_id)
        if donor_rpr is None and style_id != "Heading1":
            donor_rpr = style_run_rpr(styles_root, "Heading1")
        if donor_rpr is None:
            raise RuntimeError(f"title-run-format cannot locate run formatting for style `{style_id}`")
        touched = 0
        for run in child.findall("./w:r", NS):
            if not paragraph_text(run).strip():
                continue
            if set_run_rpr(run, donor_rpr):
                touched += 1
        if touched:
            changes.append(
                {
                    "body_child_index": index,
                    "text": text,
                    "style_id": style_id,
                    "visible_runs_touched": touched,
                }
            )
    if not changes:
        raise RuntimeError("title-run-format did not update any title runs")
    return changes


def collapse_font_alias_value(value: str) -> str:
    if ";" not in value:
        return value
    return next((part.strip() for part in value.split(";") if part.strip()), value)


def repair_font_alias_lists(*roots: ET.Element) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    for root_index, root in enumerate(roots):
        for fonts in root.findall(".//w:rFonts", NS):
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                key = qn(attr)
                value = fonts.get(key)
                if value and ";" in value:
                    collapsed = collapse_font_alias_value(value)
                    fonts.set(key, collapsed)
                    changes.append(
                        {
                            "root_index": root_index,
                            "slot": attr,
                            "before": value,
                            "after": collapsed,
                        }
                    )
    return changes


def style_by_id(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("styleId")) == style_id:
            return style
    return None


def ensure_abstract_title_style(styles_root: ET.Element) -> list[dict[str, object]]:
    if style_by_id(styles_root, "Style17") is not None:
        return []
    style = ET.Element(qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), "Style17")
    name = ET.SubElement(style, qn("name"))
    name.set(qn("val"), "Abstract Title")
    based_on = ET.SubElement(style, qn("basedOn"))
    based_on.set(qn("val"), "Normal")
    ppr = ET.SubElement(style, qn("pPr"))
    spacing = ET.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "240")
    spacing.set(qn("line"), "500")
    spacing.set(qn("lineRule"), "exact")
    jc = ET.SubElement(ppr, qn("jc"))
    jc.set(qn("val"), "center")
    rpr = ET.SubElement(style, qn("rPr"))
    fonts = ET.SubElement(rpr, qn("rFonts"))
    fonts.set(qn("ascii"), "SimHei")
    fonts.set(qn("hAnsi"), "SimHei")
    fonts.set(qn("eastAsia"), "\u9ed1\u4f53")
    fonts.set(qn("cs"), "SimHei")
    ET.SubElement(rpr, qn("b"))
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), "32")
    sz_cs = ET.SubElement(rpr, qn("szCs"))
    sz_cs.set(qn("val"), "32")
    styles_root.append(style)
    return [{"style_id": "Style17", "name": "Abstract Title"}]


def paragraph_style_name(paragraph: ET.Element, styles_root: ET.Element) -> str:
    style_id = paragraph_style_id(paragraph)
    style = style_by_id(styles_root, style_id) if style_id else None
    if style is None:
        return ""
    name = style.find("./w:name", NS)
    return name.get(qn("val")) if name is not None else ""


def paragraph_uses_title_style(paragraph: ET.Element, styles_root: ET.Element) -> bool:
    style_id = paragraph_style_id(paragraph)
    style_name = paragraph_style_name(paragraph, styles_root)
    style_key = compact_text(f"{style_id} {style_name}").lower()
    title_tokens = ("heading", "title", "toctitle", "tocheading", "caption", "abstracttitle", "abstitle")
    keyword_tokens = ("keyword", "keywords", "key words")
    return any(token in style_key for token in title_tokens) and not any(token in style_key for token in keyword_tokens)


def clear_abstract_non_title_style_bindings(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    """Remove title-style paragraph bindings from abstract bodies and keyword rows."""
    targets: list[tuple[str, tuple[int, ET.Element] | None]] = [
        ("zh_abstract_body", abstract_body_paragraph_between_title_and_keyword(body, english=False)),
        (
            "zh_keyword_line",
            find_first_paragraph(body, lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p)),
        ),
        ("en_abstract_body", abstract_body_paragraph_between_title_and_keyword(body, english=True)),
        (
            "en_keyword_line",
            find_first_paragraph(body, lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p)),
        ),
    ]
    changes: list[dict[str, object]] = []
    for surface_id, match in targets:
        if match is None:
            continue
        target_index, paragraph = match
        before_style = paragraph_style_id(paragraph)
        before_name = paragraph_style_name(paragraph, styles_root)
        if not before_style or not paragraph_uses_title_style(paragraph, styles_root):
            continue
        changed = clear_paragraph_style(paragraph)
        changes.append(
            {
                "kind": "abstract_non_title_style_binding",
                "surface_id": surface_id,
                "target_index": target_index,
                "before_style": before_style,
                "before_style_name": before_name,
                "after_style": paragraph_style_id(paragraph),
                "changed": changed,
            }
        )
    return changes


def ensure_heading1_style_from_existing_chapters(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    if style_by_id(styles_root, "Heading1") is not None:
        return []
    toc_seen = False
    for child in list(body):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen or paragraph_heading_level(child) != 1 or is_toc_entry_paragraph(child):
            continue
        donor_style_id = paragraph_style_id(child)
        donor = style_by_id(styles_root, donor_style_id)
        if donor is None:
            continue
        cloned = deepcopy(donor)
        cloned.set(qn("styleId"), "Heading1")
        name = cloned.find("./w:name", NS)
        if name is None:
            name = ET.Element(qn("name"))
            cloned.insert(0, name)
        name.set(qn("val"), "Heading 1")
        styles_root.append(cloned)
        return [{"style_id": "Heading1", "source_style_id": donor_style_id, "source": "existing_chapter_heading"}]
    raise RuntimeError("Heading1 style is missing and no existing chapter heading style can be cloned")


def apply_heading1_baseline(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    changes = ensure_heading1_style_from_existing_chapters(body, styles_root)
    heading_style = style_by_id(styles_root, "Heading1")
    if heading_style is None:
        raise RuntimeError("Heading1 style is missing")
    donor_ppr = heading_style.find("./w:pPr", NS)
    donor_rpr = heading_style.find("./w:rPr", NS)
    toc_seen = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if is_toc_heading(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if paragraph_heading_level(child) != 1 or is_toc_entry_paragraph(child):
            continue
        old_ppr = paragraph_ppr(child)
        if old_ppr is not None:
            child.remove(old_ppr)
        if donor_ppr is not None:
            child.insert(0, deepcopy(donor_ppr))
        set_paragraph_style(child, "Heading1")
        preserve_heading_page_break_owner(old_ppr, paragraph_ppr(child))
        new_ppr = paragraph_ppr(child)
        if new_ppr is not None:
            num_pr = new_ppr.find("./w:numPr", NS)
            if num_pr is not None:
                new_ppr.remove(num_pr)
        for run in child.findall("./w:r", NS):
            if not paragraph_text(run).strip():
                continue
            old_rpr = run.find("./w:rPr", NS)
            if old_rpr is not None:
                run.remove(old_rpr)
            if donor_rpr is not None:
                run.insert(0, deepcopy(donor_rpr))
        changes.append({"index": index, "text": text, "style_id": "Heading1"})
    return changes


def normalize_heading1_pagebreak_ownership(body: ET.Element, styles_root: ET.Element) -> list[dict[str, object]]:
    """Move Heading1 page starts from style inheritance to explicit chapter owners."""
    changes: list[dict[str, object]] = []
    heading_style = style_by_id(styles_root, "Heading1")
    if heading_style is not None:
        ppr = heading_style.find("./w:pPr", NS)
        page_break = ppr.find("./w:pageBreakBefore", NS) if ppr is not None else None
        if ppr is not None and page_break is not None:
            ppr.remove(page_break)
            changes.append({"kind": "style_pageBreakBefore_removed", "style_id": "Heading1"})

    first_numbered_chapter_seen = False
    for index, child in enumerate(list(body)):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if paragraph_style_id(child) != "Heading1" or paragraph_heading_level(child) != 1:
            continue
        previous = next((node for node in reversed(list(body)[:index]) if node.tag == qn("p")), None)
        previous_section_break = previous is not None and has_section_break(previous)
        if not first_numbered_chapter_seen:
            first_numbered_chapter_seen = True
            if previous is not None and not previous_section_break:
                children_now = list(body)
                previous_index = children_now.index(previous)
                previous_text = paragraph_text(previous).strip()
                previous_has_embedded_object = any(
                    previous.find(f".//w:{tag}", NS) is not None
                    for tag in ("drawing", "pict", "object")
                )
                section_owner_index = next(
                    (
                        candidate_index
                        for candidate_index in range(previous_index - 1, -1, -1)
                        if children_now[candidate_index].tag == qn("p")
                        and has_section_break(children_now[candidate_index])
                    ),
                    None,
                )
                if (
                    section_owner_index is not None
                    and not previous_text
                    and not previous_has_embedded_object
                ):
                    section_owner = children_now[section_owner_index]
                    moved_child_count = 0
                    for tail_child in list(previous):
                        if tail_child.tag == qn("pPr"):
                            continue
                        previous.remove(tail_child)
                        section_owner.append(tail_child)
                        moved_child_count += 1
                    body.remove(previous)
                    previous = section_owner
                    previous_section_break = True
                    changes.append(
                        {
                            "kind": "field_only_tail_paragraph_merged_into_section_owner",
                            "from_index": previous_index,
                            "to_index": section_owner_index,
                            "moved_child_count": moved_child_count,
                        }
                    )
            if previous_section_break:
                children_now = list(body)
                previous_index = children_now.index(previous)
                previous_text = paragraph_text(previous).strip()
                previous_has_embedded_object = any(
                    previous.find(f".//w:{tag}", NS) is not None
                    for tag in ("drawing", "pict", "object")
                )
                if not previous_text and not previous_has_embedded_object:
                    section_ppr = paragraph_ppr(previous)
                    section_break = section_ppr.find("./w:sectPr", NS) if section_ppr is not None else None
                    target_index = next(
                        (
                            candidate_index
                            for candidate_index in range(previous_index - 1, -1, -1)
                            if children_now[candidate_index].tag == qn("p")
                            and paragraph_text(children_now[candidate_index]).strip()
                        ),
                        None,
                    )
                    if section_break is not None and target_index is not None:
                        section_ppr.remove(section_break)
                        target_ppr = ensure_ppr(children_now[target_index])
                        existing_target_section = target_ppr.find("./w:sectPr", NS)
                        if existing_target_section is not None:
                            target_ppr.remove(existing_target_section)
                        target_ppr.append(section_break)
                        removed_indices: list[int] = []
                        for remove_index in range(index - 1, target_index, -1):
                            candidate = children_now[remove_index]
                            if (
                                candidate.tag == qn("p")
                                and not paragraph_text(candidate).strip()
                                and not paragraph_has_nontext_payload(candidate)
                            ):
                                body.remove(candidate)
                                removed_indices.append(remove_index)
                        changes.append(
                            {
                                "kind": "empty_section_break_moved_to_previous_visible_toc_paragraph",
                                "from_index": previous_index,
                                "to_index": target_index,
                                "removed_blank_indices": list(reversed(removed_indices)),
                            }
                        )
                ppr = paragraph_ppr(child)
                page_break = ppr.find("./w:pageBreakBefore", NS) if ppr is not None else None
                if ppr is not None and page_break is not None:
                    ppr.remove(page_break)
                changes.append(
                    {
                        "kind": "first_body_chapter_uses_previous_section_break",
                        "index": index,
                        "text": text,
                        "direct_pageBreakBefore_removed": page_break is not None,
                    }
                )
            continue
        if ensure_page_break_before(child):
            changes.append({"kind": "direct_chapter_pageBreakBefore_added", "index": index, "text": text})
    return changes


def review_anchor_children(paragraph: ET.Element) -> tuple[list[ET.Element], list[ET.Element]]:
    starts: list[ET.Element] = []
    ends: list[ET.Element] = []
    for child in list(paragraph):
        if child.tag in {qn("commentRangeStart"), qn("bookmarkStart")}:
            starts.append(deepcopy(child))
        elif child.tag in {qn("commentRangeEnd"), qn("bookmarkEnd")}:
            ends.append(deepcopy(child))
        elif child.tag == qn("r") and child.find("./w:commentReference", NS) is not None:
            ends.append(deepcopy(child))
    return starts, ends


def make_run(text: str, *, english: bool, bold: bool) -> ET.Element:
    run = ET.Element(qn("r"))
    rpr = ET.SubElement(run, qn("rPr"))
    fonts = ET.SubElement(rpr, qn("rFonts"))
    if english:
        fonts.set(qn("ascii"), "Times New Roman")
        fonts.set(qn("hAnsi"), "Times New Roman")
        fonts.set(qn("cs"), "Times New Roman")
        fonts.set(qn("eastAsia"), "\u5b8b\u4f53")
    else:
        fonts.set(qn("ascii"), "Times New Roman")
        fonts.set(qn("hAnsi"), "Times New Roman")
        fonts.set(qn("cs"), "Times New Roman")
        fonts.set(qn("eastAsia"), "\u5b8b\u4f53")
    if bold:
        ET.SubElement(rpr, qn("b"))
        ET.SubElement(rpr, qn("bCs"))
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), "24")
    sz_cs = ET.SubElement(rpr, qn("szCs"))
    sz_cs.set(qn("val"), "24")
    text_node = ET.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def explicit_font_size_rpr(
    primary: ET.Element | None,
    *fallbacks: ET.Element | None,
    default_half_points: str = "24",
) -> ET.Element:
    """Return an rPr that keeps template formatting but materializes sz/szCs."""
    result = deepcopy(primary) if primary is not None else ET.Element(qn("rPr"))
    for tag in ("sz", "szCs"):
        if result.find(f"./w:{tag}", NS) is not None:
            continue
        donor = next(
            (
                fallback.find(f"./w:{tag}", NS)
                for fallback in fallbacks
                if fallback is not None and fallback.find(f"./w:{tag}", NS) is not None
            ),
            None,
        )
        if donor is not None:
            result.append(deepcopy(donor))
            continue
        if default_half_points:
            size = ET.Element(qn(tag))
            size.set(qn("val"), default_half_points)
            result.append(size)
    return result


def rewrite_label_content_paragraph(
    paragraph: ET.Element,
    *,
    label: str,
    content: str,
    english: bool,
    style_id: str = "Style17",
    content_anchor_source: ET.Element | None = None,
) -> None:
    ppr = ensure_ppr(paragraph)
    set_paragraph_style(paragraph, style_id)
    set_frontmatter_label_metrics(paragraph)
    starts, ends = review_anchor_children(paragraph)
    content_starts, content_ends = (
        review_anchor_children(content_anchor_source)
        if content_anchor_source is not None and content_anchor_source is not paragraph
        else ([], [])
    )
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    paragraph.append(make_run(label, english=english, bold=True))
    for child in content_starts:
        paragraph.append(child)
    content_text = content.strip()
    if english:
        content_text = " " + content_text
    paragraph.append(make_run(content_text, english=english, bold=False))
    for child in content_ends:
        paragraph.append(child)
    for child in ends:
        paragraph.append(child)


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def first_run_rpr_by_label_boundary(
    paragraph: ET.Element,
    *,
    label_len: int,
) -> tuple[ET.Element | None, ET.Element | None]:
    label_rpr: ET.Element | None = None
    content_rpr: ET.Element | None = None
    cursor = 0
    for run in paragraph.findall(".//w:r", NS):
        text = run_text(run)
        if not text:
            continue
        start = cursor
        end = cursor + len(text)
        cursor = end
        rpr = run.find("./w:rPr", NS)
        if start < label_len and label_rpr is None and rpr is not None:
            label_rpr = deepcopy(rpr)
        if end > label_len and content_rpr is None and rpr is not None:
            content_rpr = deepcopy(rpr)
        if label_rpr is not None and content_rpr is not None:
            break
    return label_rpr, content_rpr


def replace_paragraph_properties_from_donor(target: ET.Element, donor: ET.Element) -> None:
    old_ppr = paragraph_ppr(target)
    if old_ppr is not None:
        target.remove(old_ppr)
    donor_ppr = paragraph_ppr(donor)
    if donor_ppr is not None:
        target.insert(0, deepcopy(donor_ppr))


def donor_run_rpr_or_none(donor: ET.Element) -> ET.Element | None:
    """Return the donor's direct visible run model; no donor means inherited."""
    return first_text_run_rpr(donor)


def ensure_gate_spacing_on_donor_paragraph(paragraph: ET.Element) -> None:
    """Keep donor paragraph shape while making zero spacing auditable for gates."""
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    if not spacing.get(qn("line")):
        spacing.set(qn("line"), "360")
    if not spacing.get(qn("lineRule")):
        spacing.set(qn("lineRule"), "auto")


def append_text_run_with_rpr(paragraph: ET.Element, text: str, rpr: ET.Element | None) -> None:
    run = ET.Element(qn("r"))
    if rpr is not None:
        run.append(deepcopy(rpr))
    text_node = ET.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    paragraph.append(run)


def donor_label_run_segments(
    donor: ET.Element,
    *,
    donor_label: str,
    target_label: str,
    fallback_rpr: ET.Element | None,
) -> list[tuple[str, ET.Element | None]]:
    """Replay the template label as separate runs so title marker audits remain stable."""
    if not donor_label or len(donor_label) != len(target_label):
        return [(target_label, fallback_rpr)]
    segments: list[tuple[str, ET.Element | None]] = []
    donor_cursor = 0
    target_cursor = 0
    for run in donor.findall("./w:r", NS):
        if donor_cursor >= len(donor_label):
            break
        text = run_text(run)
        if not text:
            continue
        take = min(len(text), len(donor_label) - donor_cursor)
        target_piece = target_label[target_cursor : target_cursor + take]
        if target_piece:
            segments.append((target_piece, run.find("./w:rPr", NS)))
        donor_cursor += take
        target_cursor += take
    if target_cursor != len(target_label):
        return [(target_label, fallback_rpr)]
    return segments


def rpr_with_gate_bold(rpr: ET.Element | None, *, bold: bool) -> ET.Element:
    result = deepcopy(rpr) if rpr is not None else ET.Element(qn("rPr"))
    for tag in ("b", "bCs"):
        for node in list(result.findall(f"./w:{tag}", NS)):
            result.remove(node)
    if bold:
        ET.SubElement(result, qn("b"))
        ET.SubElement(result, qn("bCs"))
    return result


def first_content_rprs_by_script(paragraph: ET.Element, *, label_len: int) -> tuple[ET.Element | None, ET.Element | None]:
    cjk_rpr: ET.Element | None = None
    latin_rpr: ET.Element | None = None
    cursor = 0
    for run in paragraph.findall(".//w:r", NS):
        text = run_text(run)
        if not text:
            continue
        start = cursor
        end = cursor + len(text)
        cursor = end
        content_text = text[max(0, label_len - start) :] if end > label_len else ""
        if not content_text:
            continue
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            continue
        if cjk_rpr is None and CJK_RE.search(content_text):
            cjk_rpr = deepcopy(rpr)
        if latin_rpr is None and ASCII_ALNUM_RE.search(content_text):
            latin_rpr = deepcopy(rpr)
        if cjk_rpr is not None and latin_rpr is not None:
            break
    return cjk_rpr, latin_rpr


def split_mixed_script_text(text: str) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    current: list[str] = []
    current_mode = ""
    for ch in text:
        if ASCII_ALNUM_RE.match(ch):
            mode = "latin"
        elif CJK_RE.match(ch):
            mode = "cjk"
        elif ord(ch) < 128 and not ch.isspace():
            mode = current_mode or "latin"
        else:
            mode = current_mode or "cjk"
        if current and mode != current_mode:
            segments.append((current_mode, "".join(current)))
            current = []
        current.append(ch)
        current_mode = mode
    if current:
        segments.append((current_mode or "cjk", "".join(current)))
    return segments


def append_content_runs_with_rprs(
    paragraph: ET.Element,
    text: str,
    *,
    cjk_rpr: ET.Element | None,
    latin_rpr: ET.Element | None,
    english: bool,
) -> None:
    if english:
        append_text_run_with_rpr(paragraph, text, first_non_none(latin_rpr, cjk_rpr))
        return
    for mode, segment in split_mixed_script_text(text):
        append_text_run_with_rpr(paragraph, segment, latin_rpr if mode == "latin" else cjk_rpr)


def first_non_none(*items: ET.Element | None) -> ET.Element | None:
    for item in items:
        if item is not None:
            return item
    return None


def rewrite_label_content_from_template_donor(
    paragraph: ET.Element,
    donor: ET.Element,
    *,
    english: bool,
    force_label_bold: bool = True,
) -> dict[str, object]:
    before_style = paragraph_style_id(paragraph)
    label, content = split_label_content(paragraph_text(paragraph).strip(), english=english)
    donor_label, _donor_content = split_label_content(paragraph_text(donor).strip(), english=english)
    if not label or not content:
        raise RuntimeError(f"abstract template replay cannot split target paragraph: {paragraph_text(paragraph)[:80]}")
    if not donor_label:
        raise RuntimeError(f"abstract template replay cannot split donor paragraph: {paragraph_text(donor)[:80]}")

    donor_label_rpr, donor_content_rpr = first_run_rpr_by_label_boundary(donor, label_len=len(donor_label))
    target_label_rpr, target_content_rpr = first_run_rpr_by_label_boundary(paragraph, label_len=len(label))
    donor_cjk_rpr, donor_latin_rpr = first_content_rprs_by_script(donor, label_len=len(donor_label))
    target_cjk_rpr, target_latin_rpr = first_content_rprs_by_script(paragraph, label_len=len(label))
    label_rpr = donor_label_rpr if donor_label_rpr is not None else target_label_rpr
    content_cjk_rpr = first_non_none(donor_cjk_rpr, donor_content_rpr, target_cjk_rpr, target_content_rpr)
    content_latin_rpr = first_non_none(donor_latin_rpr, donor_content_rpr, target_latin_rpr, target_content_rpr)
    label_rpr = explicit_font_size_rpr(label_rpr, target_label_rpr, donor_label_rpr)
    content_cjk_rpr = explicit_font_size_rpr(
        content_cjk_rpr,
        target_cjk_rpr,
        donor_content_rpr,
        target_content_rpr,
    )
    content_latin_rpr = explicit_font_size_rpr(
        content_latin_rpr,
        target_latin_rpr,
        donor_content_rpr,
        target_content_rpr,
    )
    if force_label_bold:
        label_rpr = rpr_with_gate_bold(label_rpr, bold=True)
    content_cjk_rpr = rpr_with_gate_bold(content_cjk_rpr, bold=False)
    content_latin_rpr = rpr_with_gate_bold(content_latin_rpr, bold=False)
    label_segments = donor_label_run_segments(
        donor,
        donor_label=donor_label,
        target_label=label,
        fallback_rpr=label_rpr,
    )
    if force_label_bold:
        label_segments = [
            (
                text,
                rpr_with_gate_bold(
                    explicit_font_size_rpr(rpr, label_rpr, target_label_rpr, donor_label_rpr),
                    bold=True,
                ),
            )
            for text, rpr in label_segments
        ]
    else:
        label_segments = [
            (text, explicit_font_size_rpr(rpr, label_rpr, target_label_rpr, donor_label_rpr))
            for text, rpr in label_segments
        ]

    replace_paragraph_properties_from_donor(paragraph, donor)
    starts, ends = review_anchor_children(paragraph)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    for label_text, label_segment_rpr in label_segments:
        append_text_run_with_rpr(paragraph, label_text, label_segment_rpr)
    content_text = content.strip()
    if english:
        content_text = " " + content_text
    append_content_runs_with_rprs(
        paragraph,
        content_text,
        cjk_rpr=content_cjk_rpr,
        latin_rpr=content_latin_rpr,
        english=english,
    )
    for child in ends:
        paragraph.append(child)
    return {
        "before_style": before_style,
        "after_style": paragraph_style_id(paragraph),
        "target_label": label,
        "donor_label": donor_label,
        "label_rpr_from_template": donor_label_rpr is not None,
        "label_run_segmentation": "template-donor" if len(label_segments) > 1 else "single-run",
        "force_label_bold": force_label_bold,
        "content_rpr_from_template": donor_content_rpr is not None,
    }


def abstract_body_paragraph_between_title_and_keyword(
    body: ET.Element,
    *,
    english: bool,
) -> tuple[int, ET.Element] | None:
    title_match = find_first_paragraph(
        body,
        lambda p: (
            is_en_abstract_label(paragraph_text(p))
            if english
            else is_zh_abstract_label(paragraph_text(p))
        )
        and not is_toc_entry_paragraph(p),
    )
    keyword_match = find_first_paragraph(
        body,
        lambda p: (
            is_en_keyword_line(paragraph_text(p))
            if english
            else is_zh_keyword_line(paragraph_text(p))
        )
        and not is_toc_entry_paragraph(p),
    )
    if title_match is None or keyword_match is None:
        return None
    title_index, _title = title_match
    keyword_index, _keyword = keyword_match
    for index in range(title_index + 1, keyword_index):
        child = list(body)[index]
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child).strip()
        if not text:
            continue
        if is_abstract_title_echo(text, english=english):
            continue
        return index, child
    return None


def rewrite_abstract_body_from_template_donor(
    paragraph: ET.Element,
    donor: ET.Element,
    *,
    english: bool,
) -> dict[str, object]:
    before_style = paragraph_style_id(paragraph)
    content = paragraph_text(paragraph).strip()
    donor_cjk_rpr, donor_latin_rpr = first_content_rprs_by_script(donor, label_len=0)
    target_cjk_rpr, target_latin_rpr = first_content_rprs_by_script(paragraph, label_len=0)
    content_cjk_rpr = rpr_with_gate_bold(first_non_none(donor_cjk_rpr, target_cjk_rpr), bold=False)
    content_latin_rpr = rpr_with_gate_bold(first_non_none(donor_latin_rpr, target_latin_rpr, content_cjk_rpr), bold=False)
    content_cjk_rpr = explicit_font_size_rpr(content_cjk_rpr, target_cjk_rpr, donor_cjk_rpr)
    content_latin_rpr = explicit_font_size_rpr(content_latin_rpr, target_latin_rpr, donor_latin_rpr, content_cjk_rpr)

    replace_paragraph_properties_from_donor(paragraph, donor)
    starts, ends = review_anchor_children(paragraph)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    append_content_runs_with_rprs(
        paragraph,
        content,
        cjk_rpr=content_cjk_rpr,
        latin_rpr=content_latin_rpr,
        english=english,
    )
    for child in ends:
        paragraph.append(child)
    return {
        "before_style": before_style,
        "after_style": paragraph_style_id(paragraph),
        "content_rpr_from_template": donor_cjk_rpr is not None or donor_latin_rpr is not None,
        "standalone_body_replay": True,
    }


def first_text_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if not run_text(run).strip():
            continue
        rpr = run.find("./w:rPr", NS)
        if rpr is not None:
            return deepcopy(rpr)
    return None


def rewrite_single_text_from_template_donor(
    paragraph: ET.Element,
    donor: ET.Element,
    *,
    text: str | None = None,
) -> dict[str, object]:
    before_style = paragraph_style_id(paragraph)
    before_text = paragraph_text(paragraph).strip()
    target_text = text if text is not None else before_text
    donor_rpr = donor_run_rpr_or_none(donor)
    target_rpr = first_text_run_rpr(paragraph)
    run_rpr = explicit_font_size_rpr(donor_rpr, donor_rpr) if donor_rpr is not None else None
    starts, ends = review_anchor_children(paragraph)
    replace_paragraph_properties_from_donor(paragraph, donor)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    append_text_run_with_rpr(paragraph, target_text, run_rpr)
    for child in ends:
        paragraph.append(child)
    return {
        "before_style": before_style,
        "after_style": paragraph_style_id(paragraph),
        "donor_ppr_applied": paragraph_ppr(donor) is not None,
        "donor_rpr_applied": donor_rpr is not None,
        "target_rpr_discarded_when_donor_inherits": donor_rpr is None and target_rpr is not None,
        "text_preserved": before_text == paragraph_text(paragraph).strip(),
    }


def repair_abstract_from_template_donors(body: ET.Element, template_document_root: ET.Element | None) -> list[dict[str, object]]:
    if template_document_root is None:
        raise RuntimeError("abstract-template-donor requires --template-docx")
    template_body = body_element(template_document_root)
    surfaces = [
        ("zh_abstract_title", False, "title", is_zh_abstract_label),
        ("zh_abstract_body", False, "body", is_zh_abstract_label),
        ("zh_keyword_line", False, "keyword", is_zh_keyword_line),
        ("en_abstract_title", True, "title", is_en_abstract_label),
        ("en_abstract_body", True, "body", is_en_abstract_label),
        ("en_keyword_line", True, "keyword", is_en_keyword_line),
    ]
    changes: list[dict[str, object]] = []
    for surface_id, english, surface_kind, matcher in surfaces:
        if surface_kind == "title":
            target_match = find_first_paragraph(
                body,
                lambda p, english=english: is_frontmatter_title_text(paragraph_text(p), english=english)
                and not is_toc_entry_paragraph(p),
            )
            donor_match = find_first_paragraph(
                template_body,
                lambda p, english=english: is_frontmatter_title_text(paragraph_text(p), english=english)
                and not is_toc_entry_paragraph(p),
            )
        elif surface_kind == "body":
            target_match = abstract_body_paragraph_between_title_and_keyword(body, english=english)
            donor_match = abstract_body_paragraph_between_title_and_keyword(template_body, english=english)
        else:
            target_match = find_first_paragraph(
                body,
                lambda p, matcher=matcher: matcher(paragraph_text(p)) and not is_toc_entry_paragraph(p),
            )
            donor_match = find_first_paragraph(
                template_body,
                lambda p, matcher=matcher: matcher(paragraph_text(p)) and not is_toc_entry_paragraph(p),
            )
        if target_match is None or donor_match is None:
            raise RuntimeError(f"abstract-template-donor could not locate {surface_id} target and donor")
        target_index, target_para = target_match
        donor_index, donor_para = donor_match
        if surface_kind == "title":
            detail = rewrite_single_text_from_template_donor(target_para, donor_para)
        elif surface_kind == "body":
            detail = rewrite_abstract_body_from_template_donor(
                target_para,
                donor_para,
                english=english,
            )
        else:
            detail = rewrite_label_content_from_template_donor(
                target_para,
                donor_para,
                english=english,
                force_label_bold=surface_id.endswith("_line"),
            )
        changes.append(
            {
                "kind": "abstract_template_donor",
                "surface_id": surface_id,
                "surface_kind": surface_kind,
                "target_index": target_index,
                "template_index": donor_index,
                "target_text_prefix": paragraph_text(target_para)[:80],
                **detail,
            }
        )
    return changes


def repair_tail_titles_from_template_donors(body: ET.Element, template_document_root: ET.Element | None) -> list[dict[str, object]]:
    if template_document_root is None:
        raise RuntimeError("tail-title-template-baseline requires --template-docx")
    template_body = body_element(template_document_root)
    changes: list[dict[str, object]] = []
    for label in (REFERENCES_LABEL, ACKNOWLEDGEMENT_LABEL):
        target_match = find_tail_heading(body, label)
        donor_match = find_tail_heading(template_body, label)
        if target_match is None:
            raise RuntimeError(f"tail-title-template-baseline could not locate target {label}")
        if donor_match is None:
            raise RuntimeError(f"tail-title-template-baseline could not locate template donor {label}")
        target_index, target_para = target_match
        donor_index, donor_para = donor_match
        detail = rewrite_single_text_from_template_donor(target_para, donor_para)
        changes.append(
            {
                "kind": "tail_title_template_baseline",
                "label": label,
                "target_index": target_index,
                "template_index": donor_index,
                "target_text": paragraph_text(target_para).strip(),
                **detail,
            }
        )
    return changes


def tail_body_paragraphs(body: ET.Element, label: str) -> list[ET.Element]:
    heading = find_tail_heading(body, label)
    if heading is None:
        return []
    start_index, _heading_para = heading
    children = list(body)
    rows: list[ET.Element] = []
    for paragraph in children[start_index + 1 :]:
        if paragraph.tag != qn("p"):
            continue
        text = paragraph_text(paragraph).strip()
        compact = compact_text(text)
        if compact in {compact_text(item) for item in TAIL_LABELS}:
            break
        if not text:
            continue
        rows.append(paragraph)
    return rows


def first_text_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall(".//w:r", NS):
        if any((node.text or "").strip() for node in run.findall("./w:t", NS)):
            return deepcopy(run.find("./w:rPr", NS))
    return None


def apply_run_rpr_to_text_runs(paragraph: ET.Element, donor_rpr: ET.Element | None) -> int:
    changed = 0
    for run in paragraph.findall(".//w:r", NS):
        if not any((node.text or "").strip() for node in run.findall("./w:t", NS)):
            continue
        old = run.find("./w:rPr", NS)
        if old is not None:
            run.remove(old)
            changed += 1
        if donor_rpr is not None:
            run.insert(0, deepcopy(donor_rpr))
            changed += 1
    return changed


def repair_tail_bodies_from_template_donors(body: ET.Element, template_document_root: ET.Element | None) -> list[dict[str, object]]:
    if template_document_root is None:
        raise RuntimeError("tail-body-template-baseline requires --template-docx")
    template_body = body_element(template_document_root)
    changes: list[dict[str, object]] = []
    fallback_donor_rows = tail_body_paragraphs(template_body, REFERENCES_LABEL)
    for label in (ACKNOWLEDGEMENT_LABEL, APPENDIX_LABEL):
        target_rows = tail_body_paragraphs(body, label)
        if not target_rows:
            continue
        donor_rows = tail_body_paragraphs(template_body, label) or fallback_donor_rows
        if not donor_rows:
            raise RuntimeError(f"tail-body-template-baseline could not locate template body donor for {label}")
        donor = donor_rows[0]
        donor_ppr = paragraph_ppr(donor)
        donor_rpr = first_text_run_rpr(donor)
        for paragraph in target_rows:
            text = paragraph_text(paragraph).strip()
            before_style = paragraph_style_id(paragraph)
            replace_ppr(paragraph, donor_ppr)
            run_changes = apply_run_rpr_to_text_runs(paragraph, donor_rpr)
            changes.append(
                {
                    "kind": "tail_body_template_baseline",
                    "label": label,
                    "target_text_prefix": text[:80],
                    "before_style_id": before_style,
                    "after_style_id": paragraph_style_id(paragraph),
                    "donor_text_prefix": paragraph_text(donor).strip()[:80],
                    "run_rpr_changes": run_changes,
                }
            )
    return changes


def split_label_content(text: str, *, english: bool) -> tuple[str, str]:
    if english:
        match = re.match(r"^\s*(key\s+words|keywords|keyword|abstract)\s*[:：]\s*(.*)$", text or "", flags=re.IGNORECASE)
        if not match:
            return "", (text or "").strip()
        raw_label = match.group(1).strip()
        if raw_label.lower().replace(" ", "") == "abstract":
            label = "Abstract:"
        elif raw_label.lower().replace(" ", "") == "keywords":
            label = "Key words:"
        else:
            label = "Key words:"
        return label, match.group(2).strip()
    match = re.match(r"^\s*(\u6458\s*\u8981|\u5173\s*\u952e\s*\u8bcd)\s*[:：]\s*(.*)$", text or "")
    if not match:
        return "", (text or "").strip()
    compact_label = compact_text(match.group(1))
    label = "\u6458  \u8981：" if compact_label == "\u6458\u8981" else "\u5173\u952e\u8bcd："
    return label, match.group(2).strip()


def is_abstract_title_echo(text: str, *, english: bool) -> bool:
    compact = compact_text(text).lower()
    if english:
        return compact in {"abstract", "abstract:"}
    return compact in {"\u6458\u8981", "\u6458\u8981\uff1a"}


def set_frontmatter_title_metrics(paragraph: ET.Element) -> None:
    """Make abstract titles standalone title paragraphs, not body/keyword rows."""
    ppr = ensure_ppr(paragraph)
    set_paragraph_style(paragraph, "Style17")
    page_break = ppr.find("./w:pageBreakBefore", NS)
    if page_break is not None:
        ppr.remove(page_break)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    ind = ppr.find("./w:ind", NS)
    if ind is not None:
        ppr.remove(ind)
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.Element(qn("jc"))
        ppr.append(jc)
    jc.set(qn("val"), "center")


def rewrite_standalone_title_paragraph(paragraph: ET.Element, *, title: str, english: bool) -> None:
    set_frontmatter_title_metrics(paragraph)
    starts, ends = review_anchor_children(paragraph)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    paragraph.append(make_run(title, english=english, bold=True))
    for child in ends:
        paragraph.append(child)


def rewrite_abstract_body_paragraph(
    paragraph: ET.Element,
    *,
    content: str,
    english: bool,
    content_anchor_source: ET.Element | None = None,
) -> None:
    ensure_ppr(paragraph)
    set_paragraph_style(paragraph, "Style17")
    set_frontmatter_label_metrics(paragraph)
    starts, ends = review_anchor_children(paragraph)
    content_starts, content_ends = (
        review_anchor_children(content_anchor_source)
        if content_anchor_source is not None and content_anchor_source is not paragraph
        else ([], [])
    )
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    for child in content_starts:
        paragraph.append(child)
    paragraph.append(make_run(content.strip(), english=english, bold=False))
    for child in content_ends:
        paragraph.append(child)
    for child in ends:
        paragraph.append(child)


def make_frontmatter_body_paragraph_from(title_para: ET.Element) -> ET.Element:
    paragraph = ET.Element(qn("p"))
    donor_ppr = paragraph_ppr(title_para)
    if donor_ppr is not None:
        paragraph.append(deepcopy(donor_ppr))
    return paragraph


def repair_inline_abstract_and_keywords(body: ET.Element, styles_root: ET.Element, template_styles_root: ET.Element | None) -> list[dict[str, object]]:
    if template_styles_root is None:
        raise RuntimeError("abstract-inline-labels requires --template-docx so Style17 comes from the locked school template")
    changes: list[dict[str, object]] = []
    style_source = "template_docx"
    if style_by_id(template_styles_root, "Style17") is not None:
        copied = copy_style_definition(styles_root, template_styles_root, "Style17")
    else:
        copied = bool(ensure_abstract_title_style(styles_root))
        style_source = "generated_template_compatible_fallback"
    if copied:
        changes.append({"kind": "style", "style_id": "Style17", "source": style_source})

    zh_title = find_first_paragraph(body, lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    en_title = find_first_paragraph(body, lambda p: is_en_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    zh_keyword = find_first_paragraph(body, lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    en_keyword = find_first_paragraph(body, lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    if zh_title is None or zh_keyword is None or en_title is None or en_keyword is None:
        raise RuntimeError("abstract-inline-labels could not locate all abstract and keyword surfaces")

    for title_match, keyword_match, label, english in (
        (zh_title, zh_keyword, "\u6458  \u8981：", False),
        (en_title, en_keyword, "Abstract:", True),
    ):
        _title_index, title_para = title_match
        _keyword_index, _keyword_para = keyword_match
        title_index = list(body).index(title_para)
        keyword_index = list(body).index(_keyword_para)
        title_text = paragraph_text(title_para).strip()
        title_label, existing_content = split_label_content(title_text, english=english)
        if existing_content and not is_abstract_title_echo(existing_content, english=english):
            content = existing_content
            removed_body_indices: list[int] = []
            content_anchor_source = None
        else:
            body_matches: list[tuple[int, ET.Element]] = []
            for index in range(title_index + 1, keyword_index):
                candidate = list(body)[index]
                if candidate.tag != qn("p"):
                    continue
                if not paragraph_text(candidate).strip():
                    continue
                body_matches.append((index, candidate))
            if not body_matches:
                raise RuntimeError(f"abstract-inline-labels found no body paragraph for {label}")
            content = " ".join(paragraph_text(body_para).strip() for _body_index, body_para in body_matches)
            content_anchor_source = body_matches[0][1]
            removed_body_indices = []
            for body_index, body_para in reversed(body_matches):
                body.remove(body_para)
                removed_body_indices.append(body_index)
            removed_body_indices.reverse()
        rewrite_label_content_paragraph(
            title_para,
            label=title_label or label,
            content=content,
            english=english,
            content_anchor_source=content_anchor_source,
        )
        changes.append(
            {
                "kind": "abstract_inline_label",
                "label": label,
                "title_index": title_index,
                "removed_body_indices": removed_body_indices,
                "content_prefix": content[:80],
            }
        )

    for keyword_match, english in ((zh_keyword, False), (en_keyword, True)):
        _index, paragraph = keyword_match
        raw_text = paragraph_text(paragraph).strip()
        label, content = split_label_content(raw_text, english=english)
        if not label or not content:
            raise RuntimeError(f"keyword line cannot be split into label/content: {raw_text[:80]}")
        rewrite_label_content_paragraph(paragraph, label=label, content=content, english=english)
        changes.append({"kind": "keyword_label_content", "label": label, "content_prefix": content[:80]})
    return changes


def repair_inline_abstract_and_keywords_v2(body: ET.Element, styles_root: ET.Element, template_styles_root: ET.Element | None) -> list[dict[str, object]]:
    """Split inline abstract labels into title/body surfaces for six-surface audits."""
    if template_styles_root is None:
        raise RuntimeError("abstract-inline-labels requires --template-docx so Style17 comes from the locked school template")
    changes: list[dict[str, object]] = []
    style_source = "template_docx"
    if style_by_id(template_styles_root, "Style17") is not None:
        copied = copy_style_definition(styles_root, template_styles_root, "Style17")
    else:
        copied = bool(ensure_abstract_title_style(styles_root))
        style_source = "generated_template_compatible_fallback"
    if copied:
        changes.append({"kind": "style", "style_id": "Style17", "source": style_source})

    zh_title = find_first_paragraph(body, lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    en_title = find_first_paragraph(body, lambda p: is_en_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    zh_keyword = find_first_paragraph(body, lambda p: is_zh_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    en_keyword = find_first_paragraph(body, lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    if zh_title is None or zh_keyword is None or en_title is None or en_keyword is None:
        raise RuntimeError("abstract-inline-labels could not locate all abstract and keyword surfaces")

    for title_match, keyword_match, standalone_title, inline_label, english in (
        (zh_title, zh_keyword, "摘  要", "摘  要：", False),
        (en_title, en_keyword, "Abstract", "Abstract:", True),
    ):
        _title_index, title_para = title_match
        _keyword_index, keyword_para = keyword_match
        title_index = list(body).index(title_para)
        keyword_index = list(body).index(keyword_para)
        raw_title_text = paragraph_text(title_para).strip()
        _title_label, existing_content = split_label_content(raw_title_text, english=english)
        body_para: ET.Element | None = None
        if existing_content and not is_abstract_title_echo(existing_content, english=english):
            content = existing_content
            removed_body_indices: list[int] = []
            content_anchor_source = None
            body_para = make_frontmatter_body_paragraph_from(title_para)
            body.insert(title_index + 1, body_para)
        else:
            body_matches: list[tuple[int, ET.Element]] = []
            for index in range(title_index + 1, keyword_index):
                candidate = list(body)[index]
                if candidate.tag != qn("p"):
                    continue
                if not paragraph_text(candidate).strip():
                    continue
                body_matches.append((index, candidate))
            if not body_matches:
                raise RuntimeError(f"abstract-inline-labels found no body paragraph for {inline_label}")
            content = " ".join(paragraph_text(match_para).strip() for _body_index, match_para in body_matches)
            content_anchor_source = body_matches[0][1]
            body_para = body_matches[0][1]
            removed_body_indices = []
            for body_index, extra_body_para in reversed(body_matches[1:]):
                body.remove(extra_body_para)
                removed_body_indices.append(body_index)
            removed_body_indices.reverse()
        rewrite_standalone_title_paragraph(title_para, title=standalone_title, english=english)
        rewrite_abstract_body_paragraph(
            body_para,
            content=content,
            english=english,
            content_anchor_source=content_anchor_source,
        )
        changes.append(
            {
                "kind": "abstract_inline_split",
                "title": standalone_title,
                "inline_label": inline_label,
                "title_index": title_index,
                "body_index": list(body).index(body_para),
                "removed_body_indices": removed_body_indices,
                "content_prefix": content[:80],
            }
        )

    for keyword_match, english in ((zh_keyword, False), (en_keyword, True)):
        _index, paragraph = keyword_match
        raw_text = paragraph_text(paragraph).strip()
        label, content = split_label_content(raw_text, english=english)
        if not label or not content:
            raise RuntimeError(f"keyword line cannot be split into label/content: {raw_text[:80]}")
        rewrite_label_content_paragraph(paragraph, label=label, content=content, english=english)
        changes.append({"kind": "keyword_label_content", "label": label, "content_prefix": content[:80]})
    return changes


def move_toc_after_keywords(body: ET.Element) -> list[dict[str, object]]:
    children = list(body)
    toc_match = find_first_paragraph(body, lambda p: is_toc_heading(paragraph_text(p)))
    zh_abs_match = find_first_paragraph(body, lambda p: is_zh_abstract_label(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    en_keyword_match = find_first_paragraph(body, lambda p: is_en_keyword_line(paragraph_text(p)) and not is_toc_entry_paragraph(p))
    if toc_match is None or zh_abs_match is None or en_keyword_match is None:
        raise RuntimeError("frontmatter-order could not locate TOC, Chinese abstract, and English keyword surfaces")
    toc_index, toc_title = toc_match
    zh_abs_index, _zh_abs = zh_abs_match
    _en_keyword_index, en_keyword = en_keyword_match
    if zh_abs_index < toc_index:
        return []
    toc_block = children[toc_index:zh_abs_index]
    if not toc_block or toc_title not in toc_block:
        raise RuntimeError("frontmatter-order failed to isolate the TOC block")
    for node in toc_block:
        body.remove(node)
    insert_at = list(body).index(en_keyword) + 1
    for offset, node in enumerate(toc_block):
        body.insert(insert_at + offset, node)
    return [
        {
            "kind": "frontmatter_order",
            "moved_toc_paragraph_count": len([node for node in toc_block if node.tag == qn("p")]),
            "from_index": toc_index,
            "after": "en_keyword_line",
        }
    ]


def copy_heading_baseline_from_source(body: ET.Element, source_document_root: ET.Element | None) -> list[dict[str, object]]:
    if source_document_root is None:
        raise RuntimeError("source-heading-baseline requires --source-docx")
    source_body = source_document_root.find("./w:body", NS)
    if source_body is None:
        raise RuntimeError("source DOCX has no word/document.xml body")
    donors: dict[str, tuple[ET.Element | None, ET.Element | None]] = {}
    for source_para in source_body.findall("./w:p", NS):
        text = paragraph_text(source_para).strip()
        if paragraph_heading_level(source_para) is None or is_toc_entry_paragraph(source_para):
            continue
        ppr = paragraph_ppr(source_para)
        first_rpr = None
        for run in source_para.findall("./w:r", NS):
            if paragraph_text(run).strip():
                first_rpr = run.find("./w:rPr", NS)
                break
        donors[text] = (deepcopy(ppr) if ppr is not None else None, deepcopy(first_rpr) if first_rpr is not None else None)
    changes: list[dict[str, object]] = []
    for index, para in enumerate(list(body)):
        if para.tag != qn("p"):
            continue
        text = paragraph_text(para).strip()
        if text not in donors or is_toc_entry_paragraph(para):
            continue
        donor_ppr, donor_rpr = donors[text]
        old_ppr = paragraph_ppr(para)
        if old_ppr is not None:
            para.remove(old_ppr)
        if donor_ppr is not None:
            para.insert(0, deepcopy(donor_ppr))
        for run in para.findall("./w:r", NS):
            if not paragraph_text(run).strip():
                continue
            old_rpr = run.find("./w:rPr", NS)
            if old_rpr is not None:
                run.remove(old_rpr)
            if donor_rpr is not None:
                run.insert(0, deepcopy(donor_rpr))
        changes.append({"index": index, "text": text[:120], "source": "source_docx"})
    return changes


CHAPTER_HEADER_FULL_DISPLAY = {
    "\u7eea\u8bba": "\u7b2c\u4e00\u7ae0 \u7eea\u8bba",
    "\u7167\u660e\u7cfb\u7edf\u8bbe\u8ba1": "\u7b2c\u4e8c\u7ae0 \u7167\u660e\u7cfb\u7edf\u8bbe\u8ba1",
    "\u63d2\u5ea7\u7cfb\u7edf\u8bbe\u8ba1": "\u7b2c\u4e09\u7ae0 \u63d2\u5ea7\u7cfb\u7edf\u8bbe\u8ba1",
    "\u4f4e\u538b\u914d\u7535\u7cfb\u7edf\u8bbe\u8ba1": "\u7b2c\u56db\u7ae0 \u4f4e\u538b\u914d\u7535\u7cfb\u7edf\u8bbe\u8ba1",
    "\u8bbe\u5907\u9009\u62e9\u4e0e\u6821\u9a8c": "\u7b2c\u4e94\u7ae0 \u8bbe\u5907\u9009\u62e9\u4e0e\u6821\u9a8c",
    "\u9632\u96f7\u63a5\u5730\u7cfb\u7edf\u8bbe\u8ba1": "\u7b2c\u516d\u7ae0 \u9632\u96f7\u63a5\u5730\u7cfb\u7edf\u8bbe\u8ba1",
    "\u6280\u672f\u7ecf\u6d4e\u5206\u6790": "\u7b2c\u4e03\u7ae0 \u6280\u672f\u7ecf\u6d4e\u5206\u6790",
}


def replace_text_in_text_nodes(root: ET.Element, old: str, new: str) -> int:
    replacements = 0
    if old in {"\u7ed3\u8bba", "\u53c2\u8003\u6587\u732e", "\u81f4\u8c22"}:
        return 0
    full_pattern = re.compile(r"\u7b2c[一二三四五六七八九十0-9]+\u7ae0\s*" + re.escape(old))
    for text_node in root.findall(".//w:t", NS):
        current = text_node.text or ""
        if not current or old not in current:
            continue
        if new in current:
            continue
        updated = full_pattern.sub(new, current)
        if updated == current:
            updated = current.replace(old, new)
        if updated != current:
            text_node.text = updated
            replacements += 1
    return replacements


def repair_header_chapter_number_display(parts: dict[str, bytes]) -> tuple[list[dict[str, object]], dict[str, bytes]]:
    changes: list[dict[str, object]] = []
    updated: dict[str, bytes] = {}
    for name in sorted(part for part in parts if re.fullmatch(r"word/header\d+\.xml", part)):
        root = ET.fromstring(parts[name])
        before = paragraph_text(root)
        replacement_count = 0
        for short_title, full_title in CHAPTER_HEADER_FULL_DISPLAY.items():
            replacement_count += replace_text_in_text_nodes(root, short_title, full_title)
        after = paragraph_text(root)
        if replacement_count:
            updated[name] = serialize_xml(root)
            changes.append(
                {
                    "part": name,
                    "replacement_count": replacement_count,
                    "before": before,
                    "after": after,
                }
            )
        elif any(full_title in before for full_title in CHAPTER_HEADER_FULL_DISPLAY.values()):
            changes.append({"part": name, "replacement_count": 0, "status": "already_full_display", "text": before})
    return changes, updated


def repair_footer_page_number_font_sizes(parts: dict[str, bytes], *, size_half_points: str = "21") -> tuple[list[dict[str, object]], dict[str, bytes]]:
    changes: list[dict[str, object]] = []
    updated: dict[str, bytes] = {}
    for name in sorted(part for part in parts if re.fullmatch(r"word/footer\d+\.xml", part)):
        root = ET.fromstring(parts[name])
        touched = 0
        for paragraph in root.findall(".//w:p", NS):
            has_page = any("PAGE" in (node.text or "").upper() for node in paragraph.findall(".//w:instrText", NS))
            has_page = has_page or any("PAGE" in node.get(qn("instr"), "").upper() for node in paragraph.findall(".//w:fldSimple", NS))
            if not has_page:
                continue
            for run in paragraph.findall(".//w:r", NS):
                rpr = run.find("./w:rPr", NS)
                if rpr is None:
                    rpr = ET.Element(qn("rPr"))
                    run.insert(0, rpr)
                size = rpr.find("./w:sz", NS)
                if size is None:
                    size = ET.SubElement(rpr, qn("sz"))
                size.set(qn("val"), size_half_points)
                size_cs = rpr.find("./w:szCs", NS)
                if size_cs is None:
                    size_cs = ET.SubElement(rpr, qn("szCs"))
                size_cs.set(qn("val"), size_half_points)
                touched += 1
        if touched:
            updated[name] = serialize_xml(root)
            changes.append({"part": name, "page_field_runs_touched": touched, "size_half_points": size_half_points})
    if not changes:
        raise RuntimeError("footer-page-number-font-size could not locate footer PAGE field runs")
    return changes, updated


def repair_blank_header_display_text(parts: dict[str, bytes]) -> tuple[list[dict[str, object]], dict[str, bytes]]:
    """Populate bound but visually empty header parts from the first non-empty header donor."""
    donor_paragraph: ET.Element | None = None
    donor_part = ""
    donor_text = ""
    for name in sorted(part for part in parts if re.fullmatch(r"word/header\d+\.xml", part)):
        root = ET.fromstring(parts[name])
        text = paragraph_text(root).strip()
        if not text:
            continue
        for paragraph in root.findall(".//w:p", NS):
            if paragraph_text(paragraph).strip():
                donor_paragraph = deepcopy(paragraph)
                donor_part = name
                donor_text = text
                break
        if donor_paragraph is not None:
            break
    if donor_paragraph is None:
        raise RuntimeError("header-blank-display-text could not locate a non-empty header donor")

    changes: list[dict[str, object]] = []
    updated: dict[str, bytes] = {}
    for name in sorted(part for part in parts if re.fullmatch(r"word/header\d+\.xml", part)):
        root = ET.fromstring(parts[name])
        before = paragraph_text(root).strip()
        if before:
            continue
        for child in list(root):
            root.remove(child)
        root.append(deepcopy(donor_paragraph))
        updated[name] = serialize_xml(root)
        changes.append({"part": name, "donor_part": donor_part, "new_text": donor_text})
    if not changes:
        raise RuntimeError("header-blank-display-text found no empty header parts to repair")
    return changes, updated


def repair_template_empty_headers(
    parts: dict[str, bytes],
    template_parts: dict[str, bytes] | None,
) -> tuple[list[dict[str, object]], dict[str, bytes]]:
    """Clear visible target headers only when the locked template has none."""
    if template_parts is None:
        raise RuntimeError("template-empty-headers requires --template-docx")
    template_header_parts = sorted(part for part in template_parts if re.fullmatch(r"word/header\d+\.xml", part))
    template_header_texts: list[dict[str, str]] = []
    for name in template_header_parts:
        root = ET.fromstring(template_parts[name])
        text = paragraph_text(root).strip()
        if text:
            template_header_texts.append({"part": name, "text": text})
    if template_header_texts:
        raise RuntimeError(f"template-empty-headers refused because template has visible headers: {template_header_texts}")

    changes: list[dict[str, object]] = []
    updated: dict[str, bytes] = {}
    for name in sorted(part for part in parts if re.fullmatch(r"word/header\d+\.xml", part)):
        root = ET.fromstring(parts[name])
        before_text = paragraph_text(root).strip()
        before_child_count = len(list(root))
        if not before_text and before_child_count == 1 and list(root)[0].tag == qn("p"):
            continue
        for child in list(root):
            root.remove(child)
        root.append(ET.Element(qn("p")))
        updated[name] = serialize_xml(root)
        changes.append(
            {
                "part": name,
                "before_text": before_text,
                "before_child_count": before_child_count,
                "template_header_part_count": len(template_header_parts),
            }
        )
    if not changes:
        raise RuntimeError("template-empty-headers found no visible or non-empty target header parts to clear")
    return changes, updated


def document_relationship_targets(parts: dict[str, bytes]) -> dict[str, str]:
    rel_bytes = parts.get("word/_rels/document.xml.rels")
    if not rel_bytes:
        return {}
    root = ET.fromstring(rel_bytes)
    targets: dict[str, str] = {}
    for rel in root.findall(f".//{{{PKG_REL_NS}}}Relationship"):
        rid = rel.get("Id") or ""
        target = rel.get("Target") or ""
        if not rid or not target:
            continue
        if target.startswith("/"):
            normalized = target.lstrip("/")
        elif target.startswith("word/"):
            normalized = target
        else:
            normalized = f"word/{target.lstrip('/')}"
        targets[rid] = normalized
    return targets


def first_body_child_index(body: ET.Element) -> int:
    for index, child in enumerate(list(body)):
        text = paragraph_text(child).strip() if child.tag == qn("p") else ""
        if re.match(r"^\u7b2c\s*\d+\s*\u7ae0", text):
            return index
    return len(list(body))


def frontmatter_header_targets(
    parts: dict[str, bytes],
    body: ET.Element,
) -> dict[str, list[dict[str, object]]]:
    rels = document_relationship_targets(parts)
    targets: dict[str, list[dict[str, object]]] = {}
    section_properties = body.findall(".//w:sectPr", NS)
    if len(section_properties) <= 1:
        candidate_sections = section_properties
    else:
        candidate_sections = section_properties[:-1]
    for section_index, sect in enumerate(candidate_sections):
        for ref in sect.findall("./w:headerReference", NS):
            rid = ref.get(R + "id", "")
            target = rels.get(rid)
            if not target:
                continue
            targets.setdefault(target, []).append(
                {
                    "section_index": section_index,
                    "reference_type": ref.get(qn("type"), ""),
                    "relationship_id": rid,
                }
            )
    return targets


def repair_template_frontmatter_empty_headers(
    parts: dict[str, bytes],
    body: ET.Element,
    template_parts: dict[str, bytes] | None,
    template_document_root: ET.Element | None,
) -> tuple[list[dict[str, object]], dict[str, bytes]]:
    """Clear cover/front-matter headers only when those template sections are empty."""
    if template_parts is None or template_document_root is None:
        raise RuntimeError("template-frontmatter-empty-headers requires --template-docx")
    template_body = body_element(template_document_root)
    template_targets = frontmatter_header_targets(template_parts, template_body)
    template_visible: list[dict[str, str]] = []
    for target in sorted(template_targets):
        data = template_parts.get(target)
        if not data:
            continue
        root = ET.fromstring(data)
        text = paragraph_text(root).strip()
        if text:
            template_visible.append({"part": target, "text": text})
    target_parts = frontmatter_header_targets(parts, body)
    changes: list[dict[str, object]] = []
    updated: dict[str, bytes] = {}
    for target, refs in sorted(target_parts.items()):
        data = parts.get(target)
        if not data:
            continue
        root = ET.fromstring(data)
        before_text = paragraph_text(root).strip()
        before_child_count = len(list(root))
        if not before_text and before_child_count == 1 and list(root)[0].tag == qn("p"):
            continue
        for child in list(root):
            root.remove(child)
        root.append(ET.Element(qn("p")))
        updated[target] = serialize_xml(root)
        changes.append(
            {
                "part": target,
                "before_text": before_text,
                "before_child_count": before_child_count,
                "frontmatter_references": refs,
                "template_frontmatter_visible_headers_detected": template_visible,
            }
        )
    if not changes:
        raise RuntimeError("template-frontmatter-empty-headers found no visible or non-empty front-matter headers to clear")
    return changes, updated


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def repair(
    input_docx: Path,
    output_docx: Path,
    operations: set[str],
    *,
    source_docx: Path | None = None,
    template_docx: Path | None = None,
    toc_page_map_json: Path | None = None,
    footer_page_size_half_points: str = "21",
) -> dict[str, object]:
    if input_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a new review-copy path")
    with zipfile.ZipFile(input_docx, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    document_root = ET.fromstring(parts["word/document.xml"])
    styles_root = ET.fromstring(parts["word/styles.xml"])
    source_document_root = None
    template_document_root = None
    template_styles_root = None
    template_parts = None
    if source_docx is not None:
        with zipfile.ZipFile(source_docx, "r") as zsource:
            source_document_root = ET.fromstring(zsource.read("word/document.xml"))
    if template_docx is not None:
        with zipfile.ZipFile(template_docx, "r") as ztemplate:
            template_parts = {name: ztemplate.read(name) for name in ztemplate.namelist()}
            template_document_root = ET.fromstring(ztemplate.read("word/document.xml"))
            template_styles_root = ET.fromstring(ztemplate.read("word/styles.xml"))
    body = body_element(document_root)
    changes: dict[str, object] = {}
    if "frontmatter-order" in operations:
        changes["frontmatter_order_repaired"] = move_toc_after_keywords(body)
    if "abstract-inline-labels" in operations:
        changes["abstract_inline_labels_repaired"] = repair_inline_abstract_and_keywords_v2(
            body,
            styles_root,
            template_styles_root,
        )
    if "abstract-template-donor" in operations:
        changes["abstract_template_donor_repaired"] = repair_abstract_from_template_donors(
            body,
            template_document_root,
        )
    if "source-heading-baseline" in operations:
        changes["source_heading_baseline_repaired"] = copy_heading_baseline_from_source(body, source_document_root)
    if "final-chapter-heading" in operations:
        changes["final_chapter_heading_repaired"] = repair_final_chapter_heading(body)
    if "static-toc-frontmatter-final-chapter" in operations:
        changes["static_toc_frontmatter_final_chapter_repaired"] = repair_static_toc_frontmatter_and_final_chapter(body)
    if "toc-dotted-tabs" in operations:
        changes["toc_dotted_tabs_repaired"] = ensure_toc_entry_dotted_tabs(body)
    if "frontmatter-title-style-bindings" in operations:
        changes["frontmatter_title_style_bindings_repaired"] = repair_frontmatter_title_style_bindings(body, styles_root)
    if "toc-contamination-relocate" in operations:
        changes["toc_contamination_relocated"] = relocate_toc_contamination(body)
    if "toc-contamination" in operations:
        changes["toc_contamination_removed"] = remove_toc_contamination(body)
    if "toc-empty-entries" in operations:
        changes["toc_empty_entries_removed"] = remove_empty_toc_entries(body)
    if "duplicate-page-breaks" in operations:
        changes["duplicate_page_breaks_removed"] = normalize_duplicate_page_breaks(body)
    if "tail-block-order" in operations:
        changes["tail_block_order_repaired"] = repair_tail_block_order_and_openers(body, document_root)
    if "tail-redundant-pagebreaks" in operations:
        changes["tail_redundant_pagebreaks_removed"] = remove_redundant_tail_pagebreaks(body)
    if "tail-title-template-baseline" in operations:
        changes["tail_title_template_baseline_repaired"] = repair_tail_titles_from_template_donors(body, template_document_root)
    if "tail-body-template-baseline" in operations:
        changes["tail_body_template_baseline_repaired"] = repair_tail_bodies_from_template_donors(body, template_document_root)
    if "toc-tail-entries" in operations:
        changes["toc_tail_entries_repaired"] = append_missing_tail_toc_entries(body, document_root)
    if "frontmatter-template-sections" in operations:
        changes["frontmatter_template_sections_repaired"] = repair_frontmatter_template_sections(
            body,
            template_document_root,
        )
    if "frontmatter-section-owner-migration" in operations:
        changes["frontmatter_section_owner_migration_repaired"] = repair_frontmatter_section_owner_migration(body)
    if "template-page-geometry" in operations:
        changes["template_page_geometry_repaired"] = repair_template_page_geometry(
            document_root,
            template_document_root,
        )
    if "frontmatter-empty-run-collapse" in operations:
        changes["frontmatter_empty_run_collapsed"] = collapse_frontmatter_empty_run_before_toc(body)
    if "pre-submission-blank-paragraphs" in operations:
        changes["pre_submission_blank_paragraphs_removed"] = remove_pre_submission_blank_paragraphs(body)
    if "duplicate-long-body-paragraphs" in operations:
        changes["duplicate_long_body_paragraphs_removed"] = remove_adjacent_duplicate_long_body_paragraphs(body)
    if "frontmatter-toc-bookmarks" in operations:
        changes["frontmatter_toc_bookmarks_repaired"] = repair_frontmatter_toc_bookmarks(body, document_root)
    if "toc-page-cache" in operations:
        changes["toc_page_cache_repaired"] = repair_toc_page_cache(body, toc_page_map_json)
    if "toc-template-styles" in operations:
        changes["toc_template_styles_repaired"] = repair_toc_template_styles(
            body,
            styles_root,
            template_styles_root,
            template_document_root,
        )
    if "toc-template-styles" in operations and "toc-dotted-tabs" in operations:
        changes["toc_dotted_tabs_post_template_repaired"] = ensure_toc_entry_dotted_tabs(body)
    if "toc-level-indents" in operations:
        changes["toc_level_indents_repaired"] = repair_toc_level_indents(body, styles_root)
    if "live-toc-field" in operations:
        changes["live_toc_field_repaired"] = wrap_static_toc_cache_in_live_field(
            body,
            document_root,
            styles_root,
        )
    if "toc-title-pagebreak" in operations:
        changes["toc_title_pagebreak_repaired"] = repair_toc_title_pagebreak(body, template_document_root)
    if "frontmatter-title-outline-exclusion" in operations:
        changes["frontmatter_title_outline_exclusion_repaired"] = repair_frontmatter_title_outline_exclusion(body)
    if "toc-frontmatter-cache-exclusion" in operations:
        changes["toc_frontmatter_cache_exclusion_repaired"] = repair_toc_frontmatter_cache_exclusion(body)
    if "toc-field-lock" in operations:
        changes["toc_field_lock_repaired"] = repair_toc_field_lock(body)
    if "toc-field-boundary" in operations:
        changes["toc_field_boundary_repaired"] = repair_toc_field_boundary(body)
    if "toc-tail-section-collapse" in operations:
        changes["toc_tail_section_collapsed"] = repair_toc_tail_section_collapse(body)
    if "toc-last-entry-header-section" in operations:
        changes["toc_last_entry_header_section_repaired"] = repair_toc_last_entry_header_section(body, styles_root)
    if "toc-frontmatter-cache-compact" in operations:
        changes["toc_frontmatter_cache_compact_repaired"] = repair_toc_frontmatter_cache_compact(body)
    if "reference-residue" in operations:
        changes["reference_supplement_residue_repaired"] = repair_reference_supplement_residue(body)
    if "frontmatter-signature-residual" in operations:
        changes["frontmatter_signature_residual_repaired"] = repair_frontmatter_signature_residual(body)
    if "cover-metadata-pagebreaks" in operations:
        changes["cover_metadata_pagebreaks_repaired"] = repair_cover_metadata_pagebreaks(body)
    if "cover-template-ppr" in operations:
        changes["cover_template_ppr_repaired"] = repair_cover_template_paragraph_properties(
            body,
            template_document_root,
        )
    if "cover-metadata-page-owner" in operations:
        changes["cover_metadata_page_owner_repaired"] = repair_cover_metadata_page_owner(body)
    if "cover-confidential-placeholder" in operations:
        changes["cover_confidential_placeholder_repaired"] = repair_cover_confidential_placeholder(body)
    if "attachment8-cover-cleanup" in operations:
        changes["attachment8_cover_cleanup_repaired"] = repair_attachment8_cover_cleanup(
            body,
            template_document_root,
        )
    if "attachment8-cover-template-layout" in operations:
        changes["attachment8_cover_template_layout_repaired"] = repair_attachment8_cover_template_layout(
            body,
            template_document_root,
        )
    if "body-heading-line-spacing" in operations:
        changes["body_heading_line_spacing_repaired"] = repair_body_heading_line_spacing(body)
    if "body-heading-template-baseline" in operations:
        changes["body_heading_template_baseline_repaired"] = repair_body_heading_template_baseline(
            body,
            styles_root,
            template_document_root,
            template_styles_root,
        )
    if "body-image-holder-spacing" in operations:
        changes["body_image_holder_spacing_repaired"] = repair_body_image_holder_spacing(body, styles_root)
    if "caption-direct-format" in operations:
        changes["caption_direct_format_repaired"] = repair_caption_direct_format(body)
    if "title-run-format" in operations:
        changes["title_run_format_repaired"] = repair_title_run_format(body, styles_root)
    if "font-alias-list-cleanup" in operations:
        changes["font_alias_lists_cleaned"] = repair_font_alias_lists(document_root, styles_root)
    if "abstract-style" in operations:
        changes["abstract_title_style_repaired"] = {
            "title_style_definition": ensure_abstract_title_style(styles_root),
            "non_title_style_bindings": clear_abstract_non_title_style_bindings(body, styles_root),
        }
    if "heading1-baseline" in operations:
        changes["heading1_baseline_applied"] = apply_heading1_baseline(body, styles_root)
    if "heading1-pagebreak-ownership" in operations:
        changes["heading1_pagebreak_ownership_repaired"] = normalize_heading1_pagebreak_ownership(body, styles_root)
    if "body-page-number-restart" in operations:
        changes["body_page_number_restart_repaired"] = repair_body_page_number_restart(body)
    if "body-section-header-footer" in operations:
        changes["body_section_header_footer_repaired"] = repair_body_section_header_footer(body)

    header_part_updates: dict[str, bytes] = {}
    if "header-chapter-number-display" in operations:
        header_changes, header_part_updates = repair_header_chapter_number_display(parts)
        changes["header_chapter_number_display_repaired"] = header_changes
    footer_page_number_updates: dict[str, bytes] = {}
    if "footer-page-number-font-size" in operations:
        footer_changes, footer_page_number_updates = repair_footer_page_number_font_sizes(
            parts,
            size_half_points=footer_page_size_half_points,
        )
        changes["footer_page_number_font_size_repaired"] = footer_changes
    header_blank_updates: dict[str, bytes] = {}
    if "header-blank-display-text" in operations:
        header_blank_changes, header_blank_updates = repair_blank_header_display_text(parts)
        changes["header_blank_display_text_repaired"] = header_blank_changes
    header_empty_updates: dict[str, bytes] = {}
    if "template-empty-headers" in operations:
        header_empty_changes, header_empty_updates = repair_template_empty_headers(parts, template_parts)
        changes["template_empty_headers_repaired"] = header_empty_changes
    header_frontmatter_empty_updates: dict[str, bytes] = {}
    if "template-frontmatter-empty-headers" in operations:
        header_frontmatter_empty_changes, header_frontmatter_empty_updates = repair_template_frontmatter_empty_headers(
            parts,
            body,
            template_parts,
            template_document_root,
        )
        changes["template_frontmatter_empty_headers_repaired"] = header_frontmatter_empty_changes
    header_footer_part_updates: dict[str, bytes] = {}
    if "template-header-footer-styles" in operations:
        header_footer_changes, header_footer_part_updates = repair_header_footer_template_styles(
            parts,
            styles_root,
            template_styles_root,
            template_parts,
        )
        changes["template_header_footer_styles_repaired"] = header_footer_changes
    footer_page_number_post_template_updates: dict[str, bytes] = {}
    if "footer-page-number-font-size" in operations and "template-header-footer-styles" in operations:
        merged_footer_parts = dict(parts)
        merged_footer_parts.update(header_footer_part_updates)
        post_footer_changes, footer_page_number_post_template_updates = repair_footer_page_number_font_sizes(
            merged_footer_parts,
            size_half_points=footer_page_size_half_points,
        )
        changes["footer_page_number_font_size_post_template_repaired"] = post_footer_changes

    updated_parts = dict(parts)
    updated_parts["word/document.xml"] = serialize_xml(document_root)
    updated_parts.update(header_part_updates)
    updated_parts.update(footer_page_number_updates)
    updated_parts.update(header_blank_updates)
    updated_parts.update(header_empty_updates)
    updated_parts.update(header_frontmatter_empty_updates)
    updated_parts.update(header_footer_part_updates)
    updated_parts.update(footer_page_number_post_template_updates)
    if "abstract-style" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "abstract-inline-labels" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "font-alias-list-cleanup" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "toc-template-styles" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "template-header-footer-styles" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "heading1-pagebreak-ownership" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "body-heading-template-baseline" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "body-image-holder-spacing" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "frontmatter-title-style-bindings" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    if "title-run-format" in operations:
        updated_parts["word/styles.xml"] = serialize_xml(styles_root)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        temp_output = Path(td) / "out.docx"
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in parts:
                zout.writestr(name, updated_parts[name])
        shutil.move(str(temp_output), output_docx)

    changed_zip_parts = [
        name for name in sorted(updated_parts)
        if parts.get(name) != updated_parts.get(name)
    ]
    return {
        "schema": "graduation-project-builder.thesis-frontmatter-toc-structure-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "generator": "scripts/repair_thesis_frontmatter_toc_structure.py",
        "input_docx": str(input_docx),
        "input_sha256": sha256_file(input_docx),
        "source_docx": str(source_docx) if source_docx is not None else None,
        "source_sha256": sha256_file(source_docx) if source_docx is not None else None,
        "template_docx": str(template_docx) if template_docx is not None else None,
        "template_sha256": sha256_file(template_docx) if template_docx is not None else None,
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "operations": sorted(operations),
        "footer_page_size_half_points": footer_page_size_half_points if "footer-page-number-font-size" in operations else None,
        "changed_zip_parts": changed_zip_parts,
        "changes": changes,
    }


def parse_operations(raw: str) -> set[str]:
    allowed = {
        "toc-contamination",
        "toc-empty-entries",
        "duplicate-page-breaks",
        "abstract-style",
        "abstract-inline-labels",
        "abstract-template-donor",
        "frontmatter-order",
        "heading1-baseline",
        "heading1-pagebreak-ownership",
        "source-heading-baseline",
        "final-chapter-heading",
        "static-toc-frontmatter-final-chapter",
        "toc-dotted-tabs",
        "frontmatter-title-style-bindings",
        "toc-contamination-relocate",
        "tail-block-order",
        "tail-redundant-pagebreaks",
        "tail-title-template-baseline",
        "tail-body-template-baseline",
        "toc-tail-entries",
        "toc-title-pagebreak",
        "frontmatter-title-outline-exclusion",
        "toc-frontmatter-cache-exclusion",
        "toc-field-lock",
        "toc-field-boundary",
        "toc-tail-section-collapse",
        "toc-last-entry-header-section",
        "toc-frontmatter-cache-compact",
        "toc-level-indents",
        "toc-template-styles",
        "template-page-geometry",
        "template-header-footer-styles",
        "frontmatter-template-sections",
        "frontmatter-section-owner-migration",
        "frontmatter-empty-run-collapse",
        "pre-submission-blank-paragraphs",
        "duplicate-long-body-paragraphs",
        "frontmatter-toc-bookmarks",
        "toc-page-cache",
        "live-toc-field",
        "header-chapter-number-display",
        "footer-page-number-font-size",
        "reference-residue",
        "frontmatter-signature-residual",
        "cover-metadata-pagebreaks",
        "cover-metadata-page-owner",
        "cover-template-ppr",
        "cover-confidential-placeholder",
        "attachment8-cover-cleanup",
        "attachment8-cover-template-layout",
        "body-heading-line-spacing",
        "body-heading-template-baseline",
        "body-image-holder-spacing",
        "caption-direct-format",
        "title-run-format",
        "body-page-number-restart",
        "body-section-header-footer",
        "font-alias-list-cleanup",
        "header-blank-display-text",
        "template-empty-headers",
        "template-frontmatter-empty-headers",
    }
    operations = {item.strip() for item in raw.split(",") if item.strip()}
    unknown = sorted(operations - allowed)
    if unknown:
        raise RuntimeError(f"unknown operation(s): {', '.join(unknown)}")
    if not operations:
        raise RuntimeError("at least one explicit operation is required")
    if "toc-template-styles" in operations:
        operations.add("toc-title-pagebreak")
    return operations


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair narrow thesis front-matter, TOC, and heading structure issues.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument("--source-docx")
    parser.add_argument("--template-docx")
    parser.add_argument("--toc-page-map-json")
    parser.add_argument(
        "--footer-page-size-half-points",
        default="21",
        help="Direct half-point size for PAGE field runs when footer-page-number-font-size is selected. Default: 21.",
    )
    parser.add_argument(
        "--operations",
        required=True,
        help=(
            "Comma-separated explicit operations: toc-contamination,duplicate-page-breaks,"
            "toc-empty-entries,"
            "abstract-style,abstract-inline-labels,abstract-template-donor,frontmatter-order,heading1-baseline,"
            "heading1-pagebreak-ownership,source-heading-baseline,final-chapter-heading,static-toc-frontmatter-final-chapter,"
            "toc-dotted-tabs,live-toc-field,frontmatter-title-style-bindings,"
            "toc-contamination-relocate,tail-block-order,toc-tail-entries,"
            "tail-redundant-pagebreaks,tail-title-template-baseline,tail-body-template-baseline,"
            "toc-title-pagebreak,frontmatter-title-outline-exclusion,toc-frontmatter-cache-exclusion,toc-field-lock,toc-field-boundary,toc-tail-section-collapse,toc-last-entry-header-section,toc-frontmatter-cache-compact,toc-level-indents,toc-template-styles,template-page-geometry,frontmatter-template-sections,frontmatter-empty-run-collapse,"
            "template-header-footer-styles,"
            "pre-submission-blank-paragraphs,duplicate-long-body-paragraphs,frontmatter-toc-bookmarks,toc-page-cache,"
            "header-chapter-number-display,footer-page-number-font-size,"
            "reference-residue,frontmatter-signature-residual,"
            "cover-metadata-pagebreaks,cover-metadata-page-owner,cover-template-ppr,cover-confidential-placeholder,"
            "attachment8-cover-cleanup,attachment8-cover-template-layout,"
            "body-heading-line-spacing,body-heading-template-baseline,body-image-holder-spacing,caption-direct-format,"
            "tail-body-template-baseline,"
            "body-page-number-restart,body-section-header-footer,font-alias-list-cleanup,header-blank-display-text,template-empty-headers,template-frontmatter-empty-headers"
        ),
    )
    args = parser.parse_args()
    try:
        operations = parse_operations(args.operations)
        report = repair(
            Path(args.input_docx).resolve(),
            Path(args.output_docx).resolve(),
            operations,
            source_docx=Path(args.source_docx).resolve() if args.source_docx else None,
            template_docx=Path(args.template_docx).resolve() if args.template_docx else None,
            toc_page_map_json=Path(args.toc_page_map_json).resolve() if args.toc_page_map_json else None,
            footer_page_size_half_points=args.footer_page_size_half_points,
        )
    except RuntimeError as exc:
        print(str(exc))
        return 2
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": report["output_docx"], "changed_zip_parts": report["changed_zip_parts"]}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
