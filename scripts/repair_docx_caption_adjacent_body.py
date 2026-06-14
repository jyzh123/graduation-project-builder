#!/usr/bin/env python3
"""Repair caption-adjacent body prose in a DOCX review copy.

This helper is intentionally narrow. It rewrites only ``word/document.xml`` and
only the formal caption paragraphs selected by ``--target-caption`` plus nearby
explanatory body paragraphs that can inherit caption/title formatting after
figure or screenshot insertion.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("w", W_NS)

CJK_NUMERAL_CLASS = "一二三四五六七八九十零〇壹贰叁肆伍陆柒捌玖"
CAPTION_RE = re.compile(
    rf"^\s*(?:\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"(?:\s+|[\u3000:：]\s*)(?P<title>\S.*)$",
    re.I,
)
LABEL_PREFIX_RE = re.compile(
    rf"^\s*(?P<kind>\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"\s*(?P<body>.*)$",
    re.I,
)
REF_TOKEN = (
    rf"(?P<kind>\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
)
SECTION_HEADING_RE = re.compile(
    rf"^\s*(?:"
    rf"\d+(?:\.\d+){{0,4}}"
    rf"|\u7b2c?[{CJK_NUMERAL_CLASS}]+[\u7ae0\u8282]"
    rf"|\u7ed3\u8bba|\u81f4\u8c22|\u53c2\u8003\u6587\u732e|\u9644\u5f55"
    rf")\s+\S+",
    re.I,
)
LEADING_REF_CLUSTER_RE = re.compile(
    rf"^\s*(?P<refs>{REF_TOKEN}(?:\s*(?:\u548c|\u4e0e|\u53ca|\u3001|,|\uff0c)\s*"
    rf"(?:\u56fe|\u8868)\s*(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*)*)\s*(?P<body>.*)$",
    re.I,
)
THIS_PLUS_REF_CLUSTER_RE = re.compile(
    rf"^\s*\u8be5(?P<kind>\u56fe|\u8868)\s*(?:\u548c|\u4e0e|\u53ca|\u3001)\s*"
    rf"(?:\u56fe|\u8868)\s*(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"(?:\s*(?:\u548c|\u4e0e|\u53ca|\u3001|,|\uff0c)\s*(?:\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*)*"
    rf"\s*(?P<body>.*)$",
    re.I,
)
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
    "rPr": 340,
    "sectPr": 350,
    "pPrChange": 360,
}


def qn(local_name: str) -> str:
    return W + local_name


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def is_formal_caption_text(text: str) -> bool:
    match = CAPTION_RE.match(text or "")
    if not match:
        return False
    title = (match.group("title") or "").strip()
    if not title:
        return False
    compact_title = re.sub(r"\s+", "", title)
    if compact_title.startswith(
        (
            "\u5c55\u793a",
            "\u663e\u793a",
            "\u6240\u793a",
            "\u7ed9\u51fa",
            "\u8bf4\u660e",
            "\u53cd\u6620",
            "\u8868\u660e",
            "\u4f53\u73b0",
            "\u63cf\u8ff0",
        )
    ):
        return False
    if re.search(r"[\u3002\uff0c\uff1b,;]", title) and re.search(
        r"(\u7528\u4e8e|\u7528\u4ee5|\u4fbf\u4e8e|\u8bf4\u660e|\u5bf9\u5e94|\u5176|\u4e2d\u7684)",
        title,
    ):
        return False
    if len(title) > 80 and re.search(r"[\u3002\uff0c\uff1b,;]", title):
        return False
    return True


def is_probable_section_heading_text(text: str) -> bool:
    stripped = (text or "").strip()
    if not stripped:
        return False
    return SECTION_HEADING_RE.match(stripped) is not None


def target_matches(text: str, targets: set[str]) -> bool:
    if not targets:
        return is_formal_caption_text(text)
    return normalized(text) in targets


def first_text_run(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run):
            return run
    return paragraph.find("./w:r", NS)


def clone_rpr(paragraph: ET.Element | None) -> ET.Element | None:
    if paragraph is None:
        return None
    run = first_text_run(paragraph)
    if run is None:
        return None
    rpr = run.find("./w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def clone_ppr(paragraph: ET.Element | None) -> ET.Element | None:
    if paragraph is None:
        return None
    ppr = paragraph.find("./w:pPr", NS)
    return copy.deepcopy(ppr) if ppr is not None else None


def replace_ppr(paragraph: ET.Element, donor_ppr: ET.Element | None) -> ET.Element:
    old = paragraph.find("./w:pPr", NS)
    if old is not None:
        paragraph.remove(old)
    if donor_ppr is None:
        donor_ppr = ET.Element(qn("pPr"))
    paragraph.insert(0, donor_ppr)
    return donor_ppr


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(f"./w:{tag}", NS)
    if child is None:
        child = ET.Element(qn(tag))
        parent.append(child)
    return child


def set_attr(node: ET.Element, local_name: str, value: str) -> None:
    node.set(qn(local_name), value)


def remove_child(parent: ET.Element, tag: str) -> None:
    for child in list(parent.findall(f"./w:{tag}", NS)):
        parent.remove(child)


def order_ppr_children(ppr: ET.Element) -> None:
    children = list(ppr)
    indexed = list(enumerate(children))
    indexed.sort(key=lambda item: (PPR_ORDER.get(item[1].tag.removeprefix(W), 1000), item[0]))
    for child in children:
        ppr.remove(child)
    for _old_index, child in indexed:
        ppr.append(child)


def ensure_paragraph_style(ppr: ET.Element, style_id: str, style_ids: set[str] | None) -> None:
    if style_ids is not None and style_id not in style_ids:
        return
    pstyle = ppr.find("./w:pStyle", NS)
    if pstyle is None:
        pstyle = ET.Element(qn("pStyle"))
        ppr.insert(0, pstyle)
    set_attr(pstyle, "val", style_id)


def sanitize_caption_ppr(ppr: ET.Element, caption_text: str, style_ids: set[str] | None = None) -> None:
    style_id = "IMUSTFigureCaption" if (caption_text or "").lstrip().startswith("\u56fe") else "IMUSTTableCaption"
    ensure_paragraph_style(ppr, style_id, style_ids)
    remove_child(ppr, "keepNext")
    if style_id == "IMUSTTableCaption":
        ppr.append(ET.Element(qn("keepNext")))
    spacing = ensure_child(ppr, "spacing")
    spacing_profile = (
        (("before", "240"), ("after", "0"), ("line", "360"), ("lineRule", "auto"))
        if style_id == "IMUSTTableCaption"
        else (("before", "0"), ("after", "240"), ("line", "360"), ("lineRule", "auto"))
    )
    for key, value in spacing_profile:
        set_attr(spacing, key, value)
    ind = ensure_child(ppr, "ind")
    for key in ("left", "right", "firstLine", "leftChars", "rightChars", "firstLineChars", "hanging", "hangingChars"):
        ind.attrib.pop(qn(key), None)
    jc = ensure_child(ppr, "jc")
    set_attr(jc, "val", "center")
    order_ppr_children(ppr)


def sanitize_image_holder_ppr(ppr: ET.Element, style_ids: set[str] | None = None) -> None:
    ensure_paragraph_style(ppr, "ThesisImageHolder", style_ids)
    remove_child(ppr, "outlineLvl")
    if ppr.find("./w:keepNext", NS) is None:
        ppr.append(ET.Element(qn("keepNext")))
    spacing = ensure_child(ppr, "spacing")
    for key, value in (("before", "120"), ("after", "0"), ("line", "240"), ("lineRule", "auto")):
        set_attr(spacing, key, value)
    ind = ensure_child(ppr, "ind")
    for key in ("hanging", "hangingChars"):
        ind.attrib.pop(qn(key), None)
    for key, value in (
        ("left", "0"),
        ("leftChars", "0"),
        ("right", "0"),
        ("rightChars", "0"),
        ("firstLine", "0"),
        ("firstLineChars", "0"),
    ):
        set_attr(ind, key, value)
    jc = ensure_child(ppr, "jc")
    set_attr(jc, "val", "center")
    order_ppr_children(ppr)


def sanitize_body_ppr(ppr: ET.Element) -> None:
    remove_child(ppr, "keepNext")
    remove_child(ppr, "outlineLvl")
    spacing = ensure_child(ppr, "spacing")
    for key, value in (("before", "0"), ("beforeAutospacing", "0"), ("after", "0"), ("afterAutospacing", "0"), ("line", "360"), ("lineRule", "auto")):
        set_attr(spacing, key, value)
    ind = ensure_child(ppr, "ind")
    for key, value in (("left", "0"), ("leftChars", "0"), ("right", "0"), ("rightChars", "0"), ("firstLine", "480"), ("firstLineChars", "200")):
        set_attr(ind, key, value)
    for key in ("hanging", "hangingChars"):
        ind.attrib.pop(qn(key), None)
    jc = ensure_child(ppr, "jc")
    set_attr(jc, "val", "both")
    order_ppr_children(ppr)


def sanitize_body_ppr_preserve_donor(ppr: ET.Element) -> None:
    remove_child(ppr, "keepNext")
    remove_child(ppr, "outlineLvl")
    order_ppr_children(ppr)


def rewrite_label_prefix(text: str) -> tuple[str, bool]:
    match = THIS_PLUS_REF_CLUSTER_RE.match(text or "")
    if match:
        lead = "\u8be5\u7ec4\u56fe" if match.group("kind") == "\u56fe" else "\u8be5\u7ec4\u8868"
        body = (match.group("body") or "").lstrip()
        if not body or body.startswith(("\u3002", "\uff0c", "\uff1a", ":", ",")):
            return text, False
        return lead + body, True

    match = LEADING_REF_CLUSTER_RE.match(text or "")
    if not match:
        return text, False
    ref_text = match.group("refs") or ""
    kind = match.group("kind")
    ref_count = len(re.findall(rf"(?:\u56fe|\u8868)\s*(?:\d+|[{CJK_NUMERAL_CLASS}]+)", ref_text, re.I))
    if ref_count > 1:
        lead = "\u8be5\u7ec4\u56fe" if kind == "\u56fe" else "\u8be5\u7ec4\u8868"
    else:
        lead = "\u8be5\u56fe" if kind == "\u56fe" else "\u8be5\u8868"
    body = (match.group("body") or "").lstrip()
    if not body or body.startswith(("\u3002", "\uff0c", "\uff1a", ":", ",")):
        return text, False
    return lead + body, True


def set_plain_paragraph_text(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> None:
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    run = ET.Element(qn("r"))
    if donor_rpr is not None:
        run.append(copy.deepcopy(donor_rpr))
    node = ET.Element(qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        node.set(XML_SPACE, "preserve")
    node.text = text
    run.append(node)
    paragraph.append(run)


def char_script_kind(char: str) -> str:
    if "\u4e00" <= char <= "\u9fff":
        return "cjk"
    if char.isascii() and char.isalnum():
        return "latin"
    if char.isascii() and char in "-+*/=().,:%[] ":
        return "neutral"
    return "neutral"


def next_script_kind(chars: list[str], start: int) -> str | None:
    for char in chars[start:]:
        kind = char_script_kind(char)
        if kind != "neutral":
            return kind
    return None


def split_text_by_script(text: str) -> list[str]:
    chars = list(text)
    segments: list[str] = []
    current_kind: str | None = None
    current_chars: list[str] = []
    for index, char in enumerate(chars):
        kind = char_script_kind(char)
        if kind == "neutral":
            kind = current_kind or next_script_kind(chars, index + 1) or "cjk"
        if current_kind is None:
            current_kind = kind
        if kind != current_kind and current_chars:
            segments.append("".join(current_chars))
            current_chars = [char]
            current_kind = kind
        else:
            current_chars.append(char)
    if current_chars:
        segments.append("".join(current_chars))
    return segments or [text]


def set_split_paragraph_text(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> None:
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for segment in split_text_by_script(text):
        run = ET.Element(qn("r"))
        if donor_rpr is not None:
            run.append(copy.deepcopy(donor_rpr))
        node = ET.Element(qn("t"))
        if segment.startswith(" ") or segment.endswith(" ") or "  " in segment:
            node.set(XML_SPACE, "preserve")
        node.text = segment
        run.append(node)
        paragraph.append(run)


def ensure_run_size(rpr: ET.Element | None, half_points: str) -> ET.Element:
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
    for tag in ("sz", "szCs"):
        child = ensure_child(rpr, tag)
        set_attr(child, "val", half_points)
    return rpr


def find_caption_donor(paragraphs: list[ET.Element], target_norms: set[str]) -> ET.Element | None:
    for paragraph in paragraphs:
        text = paragraph_text(paragraph).strip()
        if not is_formal_caption_text(text) or normalized(text) in target_norms:
            continue
        ppr = paragraph.find("./w:pPr", NS)
        pstyle = ppr.find("./w:pStyle", NS) if ppr is not None else None
        if pstyle is not None:
            continue
        return paragraph
    return None


def find_body_donor(paragraphs: list[ET.Element], start_index: int) -> ET.Element | None:
    for paragraph in reversed(paragraphs[:start_index]):
        text = paragraph_text(paragraph).strip()
        if len(text) < 80:
            continue
        if is_formal_caption_text(text) or LEADING_REF_CLUSTER_RE.match(text) or THIS_PLUS_REF_CLUSTER_RE.match(text):
            continue
        if paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None:
            continue
        return paragraph
    return None


def next_explanatory_paragraph(paragraphs: list[ET.Element], caption_index: int) -> tuple[int, ET.Element] | None:
    for index in range(caption_index + 1, len(paragraphs)):
        paragraph = paragraphs[index]
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        if is_formal_caption_text(text):
            return None
        return index, paragraph
    return None


def paragraph_has_picture(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None


def paragraph_has_caption_like_body_format(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return False
    pstyle = ppr.find("./w:pStyle", NS)
    style_id = (pstyle.get(qn("val")) if pstyle is not None else "") or ""
    if re.search(r"(caption|title|heading)", style_id, re.I):
        return True
    if ppr.find("./w:keepNext", NS) is not None or ppr.find("./w:outlineLvl", NS) is not None:
        return True
    jc = ppr.find("./w:jc", NS)
    if ((jc.get(qn("val")) if jc is not None else "") or "").lower() in {"center", "right"}:
        return True
    ind = ppr.find("./w:ind", NS)
    first_line = (ind.get(qn("firstLine")) if ind is not None else "") or ""
    spacing = ppr.find("./w:spacing", NS)
    line = (spacing.get(qn("line")) if spacing is not None else "") or ""
    before = (spacing.get(qn("before")) if spacing is not None else "") or ""
    after = (spacing.get(qn("after")) if spacing is not None else "") or ""
    if first_line in {"", "0"} and line in {"240", "300"}:
        return True
    if first_line in {"", "0"} and (before not in {"", "0"} or after not in {"", "0"}):
        return True
    for size in paragraph.findall(".//w:sz", NS):
        try:
            if int(size.get(qn("val")) or "0") >= 28:
                return True
        except ValueError:
            continue
    return False


def nearby_explanatory_paragraphs(
    paragraphs: list[ET.Element],
    caption_index: int,
    max_count: int,
) -> list[tuple[int, ET.Element]]:
    results: list[tuple[int, ET.Element]] = []
    for index in range(caption_index + 1, len(paragraphs)):
        paragraph = paragraphs[index]
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        if is_formal_caption_text(text) or is_probable_section_heading_text(text) or paragraph_has_picture(paragraph):
            break
        results.append((index, paragraph))
        if len(results) >= max_count:
            break
    return results


def repair_document_xml(
    xml_bytes: bytes,
    target_captions: list[str],
    *,
    max_nearby_body_paragraphs: int = 2,
    style_ids: set[str] | None = None,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    paragraphs = [child for child in list(body) if child.tag == qn("p")]
    target_norms = {normalized(item) for item in target_captions if item.strip()}
    caption_donor = find_caption_donor(paragraphs, target_norms)
    caption_ppr = clone_ppr(caption_donor)
    caption_rpr = ensure_run_size(clone_rpr(caption_donor), "21")

    changes: list[dict[str, object]] = []
    touched = 0
    for index, paragraph in enumerate(paragraphs):
        caption_text = paragraph_text(paragraph).strip()
        if not target_matches(caption_text, target_norms):
            continue
        if not is_formal_caption_text(caption_text):
            continue
        touched += 1
        old_caption_ppr = ET.tostring(paragraph.find("./w:pPr", NS), encoding="unicode") if paragraph.find("./w:pPr", NS) is not None else ""
        ppr = replace_ppr(paragraph, copy.deepcopy(caption_ppr))
        sanitize_caption_ppr(ppr, caption_text, style_ids)
        set_plain_paragraph_text(paragraph, caption_text, caption_rpr)
        changes.append(
            {
                "kind": "caption",
                "paragraph_index": index + 1,
                "caption": caption_text,
                "old_ppr_prefix": old_caption_ppr[:160],
            }
        )
        if index > 0 and paragraph_has_picture(paragraphs[index - 1]):
            holder_ppr = paragraphs[index - 1].find("./w:pPr", NS)
            if holder_ppr is None:
                holder_ppr = ET.Element(qn("pPr"))
                paragraphs[index - 1].insert(0, holder_ppr)
            sanitize_image_holder_ppr(holder_ppr, style_ids)
            changes.append(
                {
                    "kind": "image-holder",
                    "paragraph_index": index,
                    "caption": caption_text,
                    "image_holder_repaired": True,
                }
            )

        nearby_paragraphs = nearby_explanatory_paragraphs(paragraphs, index, max_nearby_body_paragraphs)
        if not nearby_paragraphs:
            changes.append(
                {
                    "kind": "body",
                    "paragraph_index": None,
                    "caption": caption_text,
                    "issue": "no explanatory body paragraph before next formal caption",
                }
            )
            continue
        for nearby_number, (body_index, body_paragraph) in enumerate(nearby_paragraphs, start=1):
            old_body_text = paragraph_text(body_paragraph)
            new_body_text, prefix_changed = rewrite_label_prefix(old_body_text)
            caption_like_format = paragraph_has_caption_like_body_format(body_paragraph)
            if nearby_number > 1 and not prefix_changed and not caption_like_format:
                changes.append(
                    {
                        "kind": "body",
                        "paragraph_index": body_index + 1,
                        "caption": caption_text,
                        "old_text_prefix": old_body_text[:120],
                        "skipped": "nearby body paragraph already looks body-like",
                    }
                )
                continue
            body_donor = find_body_donor(paragraphs, index)
            if body_donor is None:
                body_donor = find_body_donor(paragraphs, body_index)
            body_ppr = replace_ppr(body_paragraph, clone_ppr(body_donor))
            if body_donor is not None:
                sanitize_body_ppr_preserve_donor(body_ppr)
            else:
                sanitize_body_ppr(body_ppr)
            set_split_paragraph_text(body_paragraph, new_body_text, clone_rpr(body_donor))
            changes.append(
                {
                    "kind": "body",
                    "paragraph_index": body_index + 1,
                    "caption": caption_text,
                    "nearby_number": nearby_number,
                    "old_text_prefix": old_body_text[:120],
                    "new_text_prefix": new_body_text[:120],
                    "label_prefix_rewritten": prefix_changed,
                    "caption_like_format_repaired": caption_like_format,
                    "body_donor_text_prefix": paragraph_text(body_donor)[:120] if body_donor is not None else "",
                }
            )

    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph).strip()
        if not text or is_formal_caption_text(text) or paragraph_has_picture(paragraph):
            continue
        new_text, prefix_changed = rewrite_label_prefix(text)
        if not prefix_changed:
            continue
        body_donor = find_body_donor(paragraphs, index)
        if body_donor is None:
            body_donor = find_body_donor(paragraphs, len(paragraphs))
        body_ppr = replace_ppr(paragraph, clone_ppr(body_donor))
        if body_donor is not None:
            sanitize_body_ppr_preserve_donor(body_ppr)
        else:
            sanitize_body_ppr(body_ppr)
        set_split_paragraph_text(paragraph, new_text, clone_rpr(body_donor))
        changes.append(
            {
                "kind": "body-label-prefix",
                "paragraph_index": index + 1,
                "old_text_prefix": text[:120],
                "new_text_prefix": new_text[:120],
                "label_prefix_rewritten": True,
                "body_donor_text_prefix": paragraph_text(body_donor)[:120] if body_donor is not None else "",
            }
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), {
        "target_caption_count": len(target_norms) if target_norms else "all",
        "touched_caption_count": touched,
        "nearby_body_paragraph_limit": max_nearby_body_paragraphs,
        "caption_donor_text": paragraph_text(caption_donor).strip() if caption_donor is not None else "",
        "changes": changes,
        "verdict": "pass" if touched and changes else "fail",
    }


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    target_captions: list[str],
    *,
    max_nearby_body_paragraphs: int = 2,
) -> dict[str, object]:
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()
    if input_docx == output_docx:
        raise RuntimeError("output DOCX must be a fresh review-copy path")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    changed_parts: list[str] = []
    document_report: dict[str, object] = {}
    original_document_sha = ""
    repaired_document_sha = ""
    style_ids: set[str] | None = None
    with zipfile.ZipFile(input_docx, "r") as zin:
        if "word/styles.xml" in zin.namelist():
            styles_root = ET.fromstring(zin.read("word/styles.xml"))
            style_ids = {
                style.get(qn("styleId")) or ""
                for style in styles_root.findall(".//w:style", NS)
                if style.get(qn("styleId"))
            }
    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                original_document_sha = sha256_bytes(data)
                data, document_report = repair_document_xml(
                    data,
                    target_captions,
                    max_nearby_body_paragraphs=max_nearby_body_paragraphs,
                    style_ids=style_ids,
                )
                repaired_document_sha = sha256_bytes(data)
                if original_document_sha != repaired_document_sha:
                    changed_parts.append(info.filename)
            zout.writestr(info, data)

    return {
        "schema": "graduation-project-builder.caption-adjacent-body-repair.v1",
        "input_docx_path": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx_path": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "changed_parts": changed_parts,
        "original_document_xml_sha256": original_document_sha,
        "repaired_document_xml_sha256": repaired_document_sha,
        **document_report,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--target-caption", action="append", default=[], help="Visible figure/table caption to repair. Repeatable.")
    parser.add_argument("--max-nearby-body-paragraphs", type=int, default=2, help="Maximum explanatory body paragraphs after each caption to inspect.")
    parser.add_argument("--report-json", required=True, type=Path)
    args = parser.parse_args()
    if args.max_nearby_body_paragraphs < 1:
        parser.error("--max-nearby-body-paragraphs must be at least 1")

    report = repair_docx(
        args.input_docx,
        args.output_docx,
        args.target_caption,
        max_nearby_body_paragraphs=args.max_nearby_body_paragraphs,
    )
    args.report_json.resolve().parent.mkdir(parents=True, exist_ok=True)
    args.report_json.resolve().write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"verdict": report["verdict"], "touched_caption_count": report["touched_caption_count"]}, ensure_ascii=True))
    return 0 if report.get("verdict") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
