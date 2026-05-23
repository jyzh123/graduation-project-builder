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
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
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
    if not template_level_map:
        raise RuntimeError("template DOCX does not define TOC paragraph styles")
    toc_index, body_start = find_toc_and_body_start(template_body)
    donors: dict[int, ET.Element] = {}
    for _body_index, paragraph, level in toc_entry_paragraphs_in_range(
        template_body,
        toc_index,
        body_start,
        template_styles_root,
        template_level_map,
    ):
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
    toc_index, _body_start = find_toc_and_body_start(template_body)
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


def repair_toc_template_styles(
    body: ET.Element,
    styles_root: ET.Element,
    template_styles_root: ET.Element | None,
    template_document_root: ET.Element | None = None,
) -> list[dict[str, object]]:
    if template_styles_root is None:
        raise RuntimeError("toc-template-styles requires --template-docx")
    template_level_map = template_toc_style_ids(template_styles_root)
    if not template_level_map:
        raise RuntimeError("template DOCX does not define TOC paragraph styles")
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
            continue
        donor = donors.get(level)
        if donor is None:
            continue
        old_style_id = paragraph_style_id(paragraph)
        selected_style_id = selected_styles.get(level)
        replace_ppr(paragraph, ppr_without_style(paragraph_ppr(donor)), selected_style_id)
        role_counts = apply_visible_run_rprs(paragraph, run_role_rprs(donor))
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


def is_toc_heading(text: str) -> bool:
    normalized = compact_text(text).lower()
    return normalized in {compact_text("\u76ee\u5f55").lower(), "contents", "tableofcontents"}


def is_front_matter_toc_label(text: str) -> bool:
    if re.search(r"(?:\d+|[IVXLCDM]+)\s*$", text or "", flags=re.IGNORECASE) is None:
        return False
    label = compact_text(re.sub(r"(?:\d+|[IVXLCDM]+)\s*$", "", text or "", flags=re.IGNORECASE)).lower()
    return label in {compact_text("\u6458\u8981").lower(), "abstract"}


def has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None


def is_toc_entry_paragraph(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph).strip()
    style_id = paragraph_style_id(paragraph).lower()
    if style_id.startswith("toc"):
        return True
    if has_tab(paragraph) or "\u2026" in text:
        return True
    return is_front_matter_toc_label(text)


CAPTION_RE = re.compile(r"^\s*[\u56fe\u8868]\s*\d+(?:[-.\uff0d]\d+)?")


def is_zh_abstract_label(text: str) -> bool:
    compact = compact_text(re.sub(r"[:：].*$", "", text or ""))
    return compact == "\u6458\u8981"


def is_en_abstract_label(text: str) -> bool:
    stripped = (text or "").strip()
    compact = re.sub(r"[\s:：]+", "", stripped).lower()
    return compact.startswith("abstract")


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
    """Add a direct off switch when a previous section break already owns pagination."""
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
            reason = "TOC title already has pageBreakBefore"
        elif paragraph_heading_level(child) == 1 and page_before:
            should_remove = True
            reason = "level-1 body heading already has pageBreakBefore"
        elif paragraph_heading_level(child) == 1 and previous_section_break:
            should_remove = True
            reason = "previous paragraph owns the section break; remove duplicate hard page break on body opener"
        if not should_remove:
            continue
        removed = remove_page_break_runs(child)
        suppressed_inherited = False
        if paragraph_heading_level(child) == 1 and previous_section_break:
            suppressed_inherited = suppress_inherited_page_break_before(child)
        if removed or suppressed_inherited:
            changes.append({"index": index, "text": text, "removed_page_break_runs": removed, "reason": reason})
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
    for label in TAIL_LABELS:
        match = find_tail_heading(body, label)
        if match is None:
            continue
        heading_index, heading = match
        removed_heading_breaks = 0
        if paragraph_has_true_page_break_before(heading):
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
                        "reason": "tail heading already owns pagination through pageBreakBefore",
                    }
                )
    return changes


def toc_entry_visible_label(text: str) -> str:
    value = re.sub(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", "", text or "", flags=re.IGNORECASE)
    value = re.sub(r"[\t\u2026.·•]+", " ", value)
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
        if run.find("./w:tab", NS) is not None:
            seen_tab = True
        for text_node in run.findall("./w:t", NS):
            value = text_node.text or ""
            if not value.strip():
                continue
            if not seen_tab:
                text_node.text = label if not title_written else ""
                title_written = True
            else:
                text_node.text = page if not page_written else ""
                page_written = True
    if not title_written:
        run = ET.SubElement(paragraph, qn("r"))
        text_node = ET.SubElement(run, qn("t"))
        text_node.text = label
    if not page_written:
        run = ET.SubElement(paragraph, qn("r"))
        text_node = ET.SubElement(run, qn("t"))
        text_node.text = page


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


def child_index(body: ET.Element, element: ET.Element) -> int:
    for index, child in enumerate(list(body)):
        if child is element:
            return index
    raise RuntimeError("element is not a child of the supplied body")


def has_section_break(element: ET.Element) -> bool:
    return element.tag == qn("p") and element.find("./w:pPr/w:sectPr", NS) is not None


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
    en_title = previous_nonempty_child_index(body, start_exclusive=en_abstract, stop_exclusive=zh_keyword)
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
    clones = clone_template_range(template_body, start, end)
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
            page_instr_changes = 0
            run_rpr_changes = 0
            for paragraph in root.findall(".//w:p", NS):
                previous_style = paragraph_style_id(paragraph)
                set_paragraph_style(paragraph, footer_style_id)
                paragraph_count += 1 if previous_style != footer_style_id else 0
                for run in paragraph.findall("./w:r", NS):
                    if set_run_rpr(run, donor_footer_rpr):
                        run_rpr_changes += 1
                for instr in paragraph.findall(".//w:instrText", NS):
                    text = instr.text or ""
                    if re.fullmatch(r"\s*PAGE\s*", text, re.IGNORECASE):
                        instr.text = " PAGE  \\* MERGEFORMAT "
                        page_instr_changes += 1
            if paragraph_count or page_instr_changes or run_rpr_changes:
                updates[name] = serialize_xml(root)
                part_changes.append(
                    {
                        "part": name,
                        "surface": "footer",
                        "style_id": footer_style_id,
                        "paragraph_style_changes": paragraph_count,
                        "page_instr_mergeformat_changes": page_instr_changes,
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
            "reason": reason,
        }
    ]


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


def first_declaration_index(body: ET.Element) -> int:
    for index, child in enumerate(list(body)):
        if child.tag == qn("p") and compact_text(paragraph_text(child)) == compact_text("\u72ec\u521b\u6027\u58f0\u660e"):
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
    spacing.set(qn("line"), "240")
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
    for attr in ("hanging", "firstLineChars", "leftChars", "rightChars", "hangingChars"):
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
    style_id = normal_style_id(styles_root)
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
            changes.append({"paragraph_index": index, "text": text[:120], "style_id": style_id})
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
    spacing.set(qn("line"), "240")
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
        review_anchor_children(content_anchor_source) if content_anchor_source is not None else ([], [])
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
    else:
        for tag in ("b", "bCs"):
            node = ET.SubElement(result, qn(tag))
            node.set(qn("val"), "0")
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
        label_segments = [(text, rpr_with_gate_bold(rpr, bold=True)) for text, rpr in label_segments]

    replace_paragraph_properties_from_donor(paragraph, donor)
    ensure_gate_spacing_on_donor_paragraph(paragraph)
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


def repair_abstract_from_template_donors(body: ET.Element, template_document_root: ET.Element | None) -> list[dict[str, object]]:
    if template_document_root is None:
        raise RuntimeError("abstract-template-donor requires --template-docx")
    template_body = body_element(template_document_root)
    surfaces = [
        ("zh_abstract_body", False, is_zh_abstract_label),
        ("zh_keyword_line", False, is_zh_keyword_line),
        ("en_abstract_body", True, is_en_abstract_label),
        ("en_keyword_line", True, is_en_keyword_line),
    ]
    changes: list[dict[str, object]] = []
    for surface_id, english, matcher in surfaces:
        target_match = find_first_paragraph(body, lambda p, matcher=matcher: matcher(paragraph_text(p)) and not is_toc_entry_paragraph(p))
        donor_match = find_first_paragraph(template_body, lambda p, matcher=matcher: matcher(paragraph_text(p)) and not is_toc_entry_paragraph(p))
        if target_match is None or donor_match is None:
            raise RuntimeError(f"abstract-template-donor could not locate {surface_id} target and donor")
        target_index, target_para = target_match
        donor_index, donor_para = donor_match
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
                "target_index": target_index,
                "template_index": donor_index,
                "target_text_prefix": paragraph_text(target_para)[:80],
                **detail,
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
        changes["abstract_inline_labels_repaired"] = repair_inline_abstract_and_keywords(
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
    if "toc-tail-entries" in operations:
        changes["toc_tail_entries_repaired"] = append_missing_tail_toc_entries(body, document_root)
    if "frontmatter-template-sections" in operations:
        changes["frontmatter_template_sections_repaired"] = repair_frontmatter_template_sections(
            body,
            template_document_root,
        )
    if "template-page-geometry" in operations:
        changes["template_page_geometry_repaired"] = repair_template_page_geometry(
            document_root,
            template_document_root,
        )
    if "frontmatter-empty-run-collapse" in operations:
        changes["frontmatter_empty_run_collapsed"] = collapse_frontmatter_empty_run_before_toc(body)
    if "pre-submission-blank-paragraphs" in operations:
        changes["pre_submission_blank_paragraphs_removed"] = remove_pre_submission_blank_paragraphs(body)
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
    if "toc-title-pagebreak" in operations:
        changes["toc_title_pagebreak_repaired"] = repair_toc_title_pagebreak(body, template_document_root)
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
    if "body-heading-line-spacing" in operations:
        changes["body_heading_line_spacing_repaired"] = repair_body_heading_line_spacing(body)
    if "body-image-holder-spacing" in operations:
        changes["body_image_holder_spacing_repaired"] = repair_body_image_holder_spacing(body, styles_root)
    if "caption-direct-format" in operations:
        changes["caption_direct_format_repaired"] = repair_caption_direct_format(body)
    if "font-alias-list-cleanup" in operations:
        changes["font_alias_lists_cleaned"] = repair_font_alias_lists(document_root, styles_root)
    if "abstract-style" in operations:
        changes["abstract_title_style_repaired"] = ensure_abstract_title_style(styles_root)
    if "heading1-baseline" in operations:
        changes["heading1_baseline_applied"] = apply_heading1_baseline(body, styles_root)
    if "heading1-pagebreak-ownership" in operations:
        changes["heading1_pagebreak_ownership_repaired"] = normalize_heading1_pagebreak_ownership(body, styles_root)

    header_part_updates: dict[str, bytes] = {}
    if "header-chapter-number-display" in operations:
        header_changes, header_part_updates = repair_header_chapter_number_display(parts)
        changes["header_chapter_number_display_repaired"] = header_changes
    header_footer_part_updates: dict[str, bytes] = {}
    if "template-header-footer-styles" in operations:
        header_footer_changes, header_footer_part_updates = repair_header_footer_template_styles(
            parts,
            styles_root,
            template_styles_root,
            template_parts,
        )
        changes["template_header_footer_styles_repaired"] = header_footer_changes

    updated_parts = dict(parts)
    updated_parts["word/document.xml"] = serialize_xml(document_root)
    updated_parts.update(header_part_updates)
    updated_parts.update(header_footer_part_updates)
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
        "toc-contamination-relocate",
        "tail-block-order",
        "tail-redundant-pagebreaks",
        "toc-tail-entries",
        "toc-title-pagebreak",
        "toc-template-styles",
        "template-page-geometry",
        "template-header-footer-styles",
        "frontmatter-template-sections",
        "frontmatter-empty-run-collapse",
        "pre-submission-blank-paragraphs",
        "frontmatter-toc-bookmarks",
        "toc-page-cache",
        "header-chapter-number-display",
        "reference-residue",
        "frontmatter-signature-residual",
        "cover-metadata-pagebreaks",
        "cover-metadata-page-owner",
        "cover-template-ppr",
        "body-heading-line-spacing",
        "body-image-holder-spacing",
        "caption-direct-format",
        "font-alias-list-cleanup",
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
        "--operations",
        required=True,
        help=(
            "Comma-separated explicit operations: toc-contamination,duplicate-page-breaks,"
            "toc-empty-entries,"
            "abstract-style,abstract-inline-labels,abstract-template-donor,frontmatter-order,heading1-baseline,"
            "heading1-pagebreak-ownership,source-heading-baseline,toc-contamination-relocate,tail-block-order,toc-tail-entries,"
            "tail-redundant-pagebreaks,"
            "toc-title-pagebreak,toc-template-styles,template-page-geometry,frontmatter-template-sections,frontmatter-empty-run-collapse,"
            "template-header-footer-styles,"
            "pre-submission-blank-paragraphs,frontmatter-toc-bookmarks,toc-page-cache,"
            "header-chapter-number-display,"
            "reference-residue,frontmatter-signature-residual,"
            "cover-metadata-pagebreaks,cover-metadata-page-owner,cover-template-ppr,"
            "body-heading-line-spacing,body-image-holder-spacing,caption-direct-format,"
            "font-alias-list-cleanup"
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
        )
    except RuntimeError as exc:
        print(str(exc))
        return 2
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": report["output_docx"], "changed_zip_parts": report["changed_zip_parts"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
