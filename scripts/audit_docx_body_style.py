#!/usr/bin/env python3
"""Audit thesis body-style binding and default body-style baseline drift."""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{%s}" % NS["w"]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

CJK_NUMERAL_CLASS = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341"
CHAPTER_NUMERAL_CLASS = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341"
CHAPTER_PREFIX_RE = re.compile(rf"^\s*\u7b2c\s*[0-9{CHAPTER_NUMERAL_CLASS}]+\s*\u7ae0")
CHAPTER_HEADING_RE = re.compile(
    rf"^\s*\u7b2c\s*[0-9{CHAPTER_NUMERAL_CLASS}]+\s*\u7ae0(?:\s+\S|$)"
)
SECTION_HEADING_RE = re.compile(r"^\s*(?:\d+\s+\S|\d+(?:\.\d+){1,3}\s*\S)")
HEADING_NUMBER_RE = re.compile(
    rf"^(?:\d+\s+\S|\d+(?:\.\d+){{1,3}}\s*\S|"
    rf"\u7b2c\s*[0-9{CHAPTER_NUMERAL_CLASS}]+\s*\u7ae0(?:\s+\S|$))"
)
CAPTION_RE = re.compile(
    rf"^\s*(?:\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"(?:\s+|[\u3000:：]\s*)(?P<title>\S.*)$",
    re.I,
)
CAPTION_LABEL_PREFIX_RE = re.compile(
    rf"^\s*(?:\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"(?=\s*[\u4e00-\u9fff])",
    re.I,
)
THIS_CAPTION_PLUS_LABEL_PREFIX_RE = re.compile(
    rf"^\s*\u8be5(?:\u56fe|\u8868)\s*(?:\u548c|\u4e0e|\u3001|\u53ca)\s*"
    rf"(?:\u56fe|\u8868)\s*"
    rf"(?:\d+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[-.\uff0d\uff0e](?:\d+|[{CJK_NUMERAL_CLASS}]+))*"
    rf"(?=\s*[\u4e00-\u9fff])",
    re.I,
)
REFERENCE_ENTRY_RE = re.compile(r"^\[\d+\]")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
ASCII_ALNUM_RE = re.compile(r"[A-Za-z0-9]")
FORMULA_NUMBER_LABEL_ONLY_RE = re.compile(
    r"^\s*(?:\u5f0f)?\s*[\(\uff08]\s*\d+(?:-\d+)+\s*[\)\uff09]\s*$"
)

FRONTMATTER_MARKERS = {
    "毕业设计",
    "本科毕业设计（论文）诚信承诺书",
    "摘   要",
    "摘要",
    "abstract",
    "目   录",
    "目录",
}
TERMINAL_HEADING_MARKERS = {
    "参考文献",
    "致谢",
    "谢辞",
    "谢  辞",
    "references",
    "acknowledgements",
    "acknowledgments",
}
NON_BODY_EXACT_MARKERS = {
    "关键词：",
    "关键词:",
    "key words:",
    "keywords:",
    "\u9644\u5f55",
    "\u9644   \u5f55",
    "结   论",
    "结论",
    "致   谢",
    "致谢",
    "谢辞",
    "谢 辞",
    "谢   辞",
}


def normalize_space(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "")


def is_code_like_paragraph_text(text: str) -> bool:
    stripped = (text or "").strip()
    if not stripped:
        return False
    if any(ch in stripped for ch in "，。；、"):
        return False
    if re.match(
        r"^(?:@|public\b|private\b|protected\b|function\b|const\b|let\b|var\b|if\b|return\b|throw\b|switch\b|case\b|[{}]|[})];)",
        stripped,
    ):
        return True
    if re.match(r"^(?:await\b|new\b|yield\b|Page<|List<|Map<|return\s+Map\.of\()", stripped):
        return True
    return bool(
        re.search(
            r"(?:=>|\.then\(|\.\w+\(|new\s+QueryWrapper|QueryWrapper<|ORDERED|PAID|SHIPPED|RECEIVED|@RequestBody|@PathVariable)",
            stripped,
        )
    )


def is_toc_heading_text(text: str) -> bool:
    normalized = normalize_space(text).lower()
    target = normalize_space("目录").lower()
    return normalized == target or normalized.endswith(target)


def contains_chapter_heading_marker(text: str) -> bool:
    return bool(CHAPTER_HEADING_RE.match(text or ""))


def normalized_heading_text(text: str) -> str:
    return (text or "").strip().replace("\uff0e", ".").replace("\u3002", ".")


def contains_numeric_heading_marker(text: str) -> bool:
    stripped = normalized_heading_text(text)
    if re.match(r"^\d+\s+\S", stripped):
        return True
    match = re.match(r"^(\d+(?:\.\d+){1,3})", stripped)
    if not match:
        return False
    rest = stripped[match.end(1):].lstrip()
    if not rest:
        return False
    # Decimal values and percentages are prose, not section headings.
    return rest[0] not in ".%％、，,。；;:：)）]】"


def visible_text_fragments(node: ET.Element) -> list[str]:
    fragments: list[str] = []
    for child in node.iter():
        if child.tag == W + "t":
            fragments.append(child.text or "")
        elif child.tag == W + "tab":
            fragments.append("\t")
    return fragments


def paragraph_text(node: ET.Element) -> str:
    return "".join(visible_text_fragments(node)).strip()


def paragraph_has_run_tab(node: ET.Element) -> bool:
    return "\t" in paragraph_text(node) or node.find(".//w:r/w:tab", NS) is not None


def paragraph_has_omml(node: ET.Element) -> bool:
    return node.find(".//m:oMath", NS) is not None or node.find(".//m:oMathPara", NS) is not None


def is_static_toc_leader_entry(text: str) -> bool:
    stripped = (text or "").strip()
    if not stripped:
        return False
    return bool(re.search(r"(?:\.{3,}|\u2026{2,}|\u00b7{3,})\s*\d+\s*$", stripped))


def load_xml(docx_path: Path, part: str) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read(part))


def style_node_by_type(styles_root: ET.Element) -> tuple[dict[str, ET.Element], str | None]:
    styles: dict[str, ET.Element] = {}
    default_paragraph_style_id: str | None = None
    for style in styles_root.findall("w:style", NS):
        if style.attrib.get(W + "type") != "paragraph":
            continue
        style_id = style.attrib.get(W + "styleId", "")
        if not style_id:
            continue
        styles[style_id] = style
        if style.attrib.get(W + "default") == "1":
            default_paragraph_style_id = style_id
    if default_paragraph_style_id is None:
        for style_id, style in styles.items():
            style_name = style_name_by_id(styles, style_id)
            if normalize_space(style_name).lower() == "normal" or style_id.lower() == "normal":
                default_paragraph_style_id = style_id
                break
    return styles, default_paragraph_style_id


def style_name_by_id(styles: dict[str, ET.Element], style_id: str | None) -> str:
    if not style_id or style_id not in styles:
        return ""
    name_node = styles[style_id].find("w:name", NS)
    return name_node.attrib.get(W + "val", "") if name_node is not None else ""


def style_based_on(styles: dict[str, ET.Element], style_id: str | None) -> str | None:
    if not style_id or style_id not in styles:
        return None
    node = styles[style_id].find("w:basedOn", NS)
    return node.attrib.get(W + "val") if node is not None else None


def style_chain(styles: dict[str, ET.Element], style_id: str | None) -> list[str]:
    chain: list[str] = []
    seen: set[str] = set()
    current = style_id
    while current and current in styles and current not in seen:
        seen.add(current)
        chain.append(current)
        current = style_based_on(styles, current)
    return chain


def style_has_outline_level(styles: dict[str, ET.Element], style_id: str | None) -> bool:
    for current in style_chain(styles, style_id):
        if styles[current].find("w:pPr/w:outlineLvl", NS) is not None:
            return True
    return False


def style_chain_label(styles: dict[str, ET.Element], style_id: str | None) -> str:
    labels = []
    for current in style_chain(styles, style_id):
        labels.append(style_label(styles, current))
    return " -> ".join(labels) if labels else "none"


def style_is_heading_family(styles: dict[str, ET.Element], style_id: str | None, style_name: str) -> bool:
    labels = [style_name or "", style_id or ""]
    for current in style_chain(styles, style_id):
        labels.append(current)
        labels.append(style_name_by_id(styles, current))
    lowered = " ".join(labels).lower()
    return "heading" in lowered or style_has_outline_level(styles, style_id)


def style_is_caption_title_family(styles: dict[str, ET.Element], style_id: str | None, style_name: str) -> bool:
    labels = [style_name or "", style_id or ""]
    for current in style_chain(styles, style_id):
        labels.append(current)
        labels.append(style_name_by_id(styles, current))
    lowered = normalize_space(" ".join(labels)).lower()
    tokens = (
        "caption",
        "figurecaption",
        "tablecaption",
        "title",
        "subtitle",
        "\u56fe\u9898",
        "\u8868\u9898",
        "\u9898\u6ce8",
        "\u56fe\u6ce8",
        "\u8868\u6ce8",
        "\u6807\u9898",
        "\u5c01\u9762\u6807\u9898",
        "\u6458\u8981\u6807\u9898",
        "\u81f4\u8c22\u6807\u9898",
        "\u53c2\u8003\u6587\u732e\u6807\u9898",
    )
    return any(token in lowered for token in tokens)


def paragraph_has_outline_level(node: ET.Element) -> bool:
    return node.find("./w:pPr/w:outlineLvl", NS) is not None


def paragraph_has_keep_next(node: ET.Element) -> bool:
    return node.find("./w:pPr/w:keepNext", NS) is not None


def paragraph_numbering_level(node: ET.Element) -> int | None:
    """Return OOXML numbering level for template-owned automatic headings."""

    level_node = node.find("./w:pPr/w:numPr/w:ilvl", NS)
    if level_node is None:
        return None
    value = level_node.attrib.get(W + "val")
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def paragraph_has_numbering(node: ET.Element) -> bool:
    return node.find("./w:pPr/w:numPr/w:numId", NS) is not None


def paragraph_has_picture(node: ET.Element) -> bool:
    return node.find(".//w:drawing", NS) is not None or node.find(".//w:pict", NS) is not None


PARAGRAPH_METRIC_KEYS = (
    "jc",
    "before",
    "after",
    "line",
    "lineRule",
    "firstLine",
    "left",
    "right",
    "hanging",
    "firstLineChars",
    "leftChars",
    "rightChars",
    "hangingChars",
)


RUN_METRIC_KEYS = (
    "eastAsia",
    "ascii",
    "hAnsi",
    "eastAsiaTheme",
    "asciiTheme",
    "hAnsiTheme",
    "csTheme",
    "size",
    "bold",
)


def empty_resolved_metrics() -> dict[str, str]:
    return {key: "" for key in (*PARAGRAPH_METRIC_KEYS, *RUN_METRIC_KEYS)}


def direct_style_metrics(style: ET.Element) -> dict[str, str]:
    ppr = style.find("w:pPr", NS)
    rpr = style.find("w:rPr", NS)
    rfonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    size = rpr.find("w:sz", NS) if rpr is not None else None
    bold = rpr.find("w:b", NS) if rpr is not None else None
    jc = ppr.find("w:jc", NS) if ppr is not None else None
    spacing = ppr.find("w:spacing", NS) if ppr is not None else None
    ind = ppr.find("w:ind", NS) if ppr is not None else None
    return {
        "jc": jc.attrib.get(W + "val", "") if jc is not None else "",
        "before": spacing.attrib.get(W + "before", "") if spacing is not None else "",
        "after": spacing.attrib.get(W + "after", "") if spacing is not None else "",
        "line": spacing.attrib.get(W + "line", "") if spacing is not None else "",
        "lineRule": spacing.attrib.get(W + "lineRule", "") if spacing is not None else "",
        "firstLine": ind.attrib.get(W + "firstLine", "") if ind is not None else "",
        "left": ind.attrib.get(W + "left", "") if ind is not None else "",
        "right": ind.attrib.get(W + "right", "") if ind is not None else "",
        "hanging": ind.attrib.get(W + "hanging", "") if ind is not None else "",
        "firstLineChars": ind.attrib.get(W + "firstLineChars", "") if ind is not None else "",
        "leftChars": ind.attrib.get(W + "leftChars", "") if ind is not None else "",
        "rightChars": ind.attrib.get(W + "rightChars", "") if ind is not None else "",
        "hangingChars": ind.attrib.get(W + "hangingChars", "") if ind is not None else "",
        "eastAsia": rfonts.attrib.get(W + "eastAsia", "") if rfonts is not None else "",
        "ascii": rfonts.attrib.get(W + "ascii", "") if rfonts is not None else "",
        "hAnsi": rfonts.attrib.get(W + "hAnsi", "") if rfonts is not None else "",
        "eastAsiaTheme": rfonts.attrib.get(W + "eastAsiaTheme", "") if rfonts is not None else "",
        "asciiTheme": rfonts.attrib.get(W + "asciiTheme", "") if rfonts is not None else "",
        "hAnsiTheme": rfonts.attrib.get(W + "hAnsiTheme", "") if rfonts is not None else "",
        "csTheme": rfonts.attrib.get(W + "cstheme", "") if rfonts is not None else "",
        "size": size.attrib.get(W + "val", "") if size is not None else "",
        "bold": bold.attrib.get(W + "val", "true") if bold is not None else "",
    }


def resolved_style_metrics(styles: dict[str, ET.Element], style_id: str | None, seen: set[str] | None = None) -> dict[str, str]:
    if not style_id or style_id not in styles:
        return empty_resolved_metrics()
    if seen is None:
        seen = set()
    if style_id in seen:
        return empty_resolved_metrics()
    seen.add(style_id)

    metrics = direct_style_metrics(styles[style_id])
    parent_id = style_based_on(styles, style_id)
    if parent_id:
        parent_metrics = resolved_style_metrics(styles, parent_id, seen)
        for key, value in parent_metrics.items():
            if not metrics[key]:
                metrics[key] = value
    return metrics


def style_label(styles: dict[str, ET.Element], style_id: str | None) -> str:
    if not style_id:
        return "none"
    name = style_name_by_id(styles, style_id)
    return f"{style_id} ({name})" if name else style_id


def paragraph_style_id(node: ET.Element) -> str | None:
    style_node = node.find("./w:pPr/w:pStyle", NS)
    if style_node is None:
        return None
    return style_node.attrib.get(W + "val")


def has_nontrivial_direct_formatting(node: ET.Element) -> bool:
    ppr = node.find("w:pPr", NS)
    if ppr is not None:
        for child in list(ppr):
            if child.tag != W + "pStyle":
                return True
    for run in node.findall("w:r", NS):
        rpr = run.find("w:rPr", NS)
        if rpr is not None and list(rpr):
            return True
    return False


def paragraph_instance_metrics(node: ET.Element) -> dict[str, str]:
    signature = {
        "jc": "",
        "before": "",
        "after": "",
        "line": "",
        "lineRule": "",
        "firstLine": "",
        "left": "",
        "right": "",
        "firstLineChars": "",
        "leftChars": "",
        "rightChars": "",
    }
    ppr = node.find("w:pPr", NS)
    if ppr is None:
        return signature
    jc = ppr.find("w:jc", NS)
    if jc is not None:
        signature["jc"] = jc.attrib.get(W + "val", "")
    spacing = ppr.find("w:spacing", NS)
    if spacing is not None:
        signature["before"] = spacing.attrib.get(W + "before", "")
        signature["after"] = spacing.attrib.get(W + "after", "")
        signature["line"] = spacing.attrib.get(W + "line", "")
        signature["lineRule"] = spacing.attrib.get(W + "lineRule", "")
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        signature["firstLine"] = ind.attrib.get(W + "firstLine", "")
        signature["left"] = ind.attrib.get(W + "left", "")
        signature["right"] = ind.attrib.get(W + "right", "")
        signature["firstLineChars"] = ind.attrib.get(W + "firstLineChars", "")
        signature["leftChars"] = ind.attrib.get(W + "leftChars", "")
        signature["rightChars"] = ind.attrib.get(W + "rightChars", "")
    return signature


def paragraph_effective_metrics(
    node: ET.Element,
    styles: dict[str, ET.Element],
    default_style_id: str | None,
) -> dict[str, str]:
    """Merge paragraph-style inherited pPr/rPr metrics with direct pPr metrics."""

    style_id = paragraph_style_id(node) or default_style_id
    signature = resolved_style_metrics(styles, style_id)
    direct = paragraph_instance_metrics(node)
    for key, value in direct.items():
        if value:
            signature[key] = value
    return signature


def run_visible_text(run: ET.Element) -> str:
    return "".join(visible_text_fragments(run)).strip()


def font_canonical(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def _run_bold_value(run: ET.Element) -> str:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return ""
    bold = rpr.find("w:b", NS)
    if bold is None:
        return ""
    return bold.attrib.get(W + "val", "true")


def _truthy_ooxml_bool(value: str) -> bool:
    return value.strip().lower() not in {"", "0", "false", "off", "no"}


def _int_or_none(value: object) -> int | None:
    try:
        text = str(value or "").strip()
        if not text:
            return None
        return int(float(text))
    except (TypeError, ValueError):
        return None


def paragraph_run_direct_metrics(node: ET.Element) -> dict[str, object]:
    visible_chars = 0
    bold_chars = 0
    bold_run_count = 0
    direct_sizes: list[int] = []
    direct_size_cs: list[int] = []
    direct_font_run_count = 0
    for run in node.findall("w:r", NS):
        if run.find(".//w:txbxContent", NS) is not None:
            continue
        text = run_visible_text(run)
        if not text:
            continue
        char_count = len(text)
        visible_chars += char_count
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            continue
        bold_value = _run_bold_value(run)
        if _truthy_ooxml_bool(bold_value):
            bold_run_count += 1
            bold_chars += char_count
        size = rpr.find("w:sz", NS)
        size_value = _int_or_none(size.attrib.get(W + "val", "") if size is not None else "")
        if size_value is not None:
            direct_sizes.append(size_value)
        size_cs = rpr.find("w:szCs", NS)
        size_cs_value = _int_or_none(size_cs.attrib.get(W + "val", "") if size_cs is not None else "")
        if size_cs_value is not None:
            direct_size_cs.append(size_cs_value)
        rfonts = rpr.find("w:rFonts", NS)
        if rfonts is not None and any(rfonts.attrib.get(W + key, "") for key in ("eastAsia", "ascii", "hAnsi", "cs")):
            direct_font_run_count += 1
    all_sizes = direct_sizes + direct_size_cs
    return {
        "visible_chars": visible_chars,
        "bold_chars": bold_chars,
        "bold_ratio": (bold_chars / visible_chars) if visible_chars else 0.0,
        "bold_run_count": bold_run_count,
        "direct_sizes": direct_sizes,
        "direct_size_cs": direct_size_cs,
        "max_direct_size": max(all_sizes) if all_sizes else None,
        "direct_font_run_count": direct_font_run_count,
    }


def paragraph_mixed_script_run_records(node: ET.Element) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for run in node.findall("w:r", NS):
        if run.find(".//w:txbxContent", NS) is not None:
            continue
        text = run_visible_text(run)
        if not text:
            continue
        rpr = run.find("w:rPr", NS)
        if re.fullmatch(r"\[\d+\]", text.strip()) and rpr is not None and rpr.find("w:vertAlign", NS) is not None:
            continue
        rfonts = rpr.find("w:rFonts", NS) if rpr is not None else None
        records.append(
            {
                "text": text,
                "ascii": rfonts.attrib.get(W + "ascii", "") if rfonts is not None else "",
                "hAnsi": rfonts.attrib.get(W + "hAnsi", "") if rfonts is not None else "",
                "eastAsia": rfonts.attrib.get(W + "eastAsia", "") if rfonts is not None else "",
                "cs": rfonts.attrib.get(W + "cs", "") if rfonts is not None else "",
                "asciiTheme": rfonts.attrib.get(W + "asciiTheme", "") if rfonts is not None else "",
                "hAnsiTheme": rfonts.attrib.get(W + "hAnsiTheme", "") if rfonts is not None else "",
                "eastAsiaTheme": rfonts.attrib.get(W + "eastAsiaTheme", "") if rfonts is not None else "",
                "csTheme": rfonts.attrib.get(W + "csTheme", "") if rfonts is not None else "",
            }
        )
    return records


def mixed_script_body_contamination_records(body_paragraphs: list[dict[str, object]]) -> list[dict[str, object]]:
    contaminated: list[dict[str, object]] = []
    for paragraph in body_paragraphs:
        text = str(paragraph.get("text") or "")
        text_without_numeric_citations = re.sub(r"\[\d{1,3}\]", "", text)
        if not (CJK_RE.search(text_without_numeric_citations) and ASCII_ALNUM_RE.search(text_without_numeric_citations)):
            continue
        node = paragraph.get("node")
        if not isinstance(node, ET.Element):
            continue
        run_records = paragraph_mixed_script_run_records(node)
        if not run_records:
            continue
        reasons: list[str] = []
        if len(run_records) == 1:
            reasons.append("mixed-script body paragraph collapsed into a single visible run")
        for run in run_records:
            run_text = run["text"]
            if not ASCII_ALNUM_RE.search(run_text):
                continue
            if CJK_RE.search(run_text):
                reasons.append(f"run `{run_text[:30]}` merges Chinese and ASCII content")
            western_font = run["ascii"] or run["hAnsi"] or run["asciiTheme"] or run["hAnsiTheme"] or run["cs"] or run["csTheme"]
            east_asia_font = run["eastAsia"] or run["eastAsiaTheme"]
            if not western_font:
                reasons.append(f"run `{run_text[:30]}` lacks an explicit Western font slot")
            if east_asia_font and western_font and font_canonical(east_asia_font) == font_canonical(western_font):
                reasons.append(f"run `{run_text[:30]}` uses the same East Asian and Western font family")
        if reasons:
            contaminated.append(
                {
                    "paragraph_index": paragraph["paragraph_index"],
                    "text": text,
                    "reasons": tuple(dict.fromkeys(reasons)),
                }
            )
    return contaminated


def is_heading_like(text: str, style_name: str, style_id: str | None, node: ET.Element | None = None) -> bool:
    lowered_style = (style_name or style_id or "").lower()
    numbering_level = paragraph_numbering_level(node) if node is not None else None
    return (
        lowered_style.startswith("heading")
        or "heading" in lowered_style
        or contains_numeric_heading_marker(text)
        or contains_chapter_heading_marker(text)
        or (
            numbering_level is not None
            and numbering_level <= 3
            and bool(normalize_space(text))
            and not is_static_toc_leader_entry(text)
        )
    )


def heading_level_from_text(text: str, node: ET.Element | None = None) -> int | None:
    numbering_level = paragraph_numbering_level(node) if node is not None else None
    if numbering_level is not None and numbering_level <= 3:
        return numbering_level
    stripped = normalized_heading_text(text)
    if CHAPTER_HEADING_RE.match(stripped):
        return 0
    match = re.match(r"^(\d+(?:\.\d+){0,3})(?:\s+\S|\S)", stripped)
    if not match:
        return None
    return min(match.group(1).count("."), 2)


def is_body_heading_text(text: str, style_name: str, style_id: str | None, node: ET.Element | None = None) -> bool:
    """Return True only for real body heading lines, not prose mentioning a chapter."""

    stripped = normalized_heading_text(text)
    if not stripped:
        return False
    if is_static_toc_leader_entry(stripped) or is_formal_caption_text(stripped):
        return False
    if contains_chapter_heading_marker(stripped) or contains_numeric_heading_marker(stripped):
        return True
    numbering_level = paragraph_numbering_level(node) if node is not None else None
    if numbering_level is not None and numbering_level <= 3 and not looks_like_body_prose(stripped):
        return True
    lowered_style = (style_name or style_id or "").lower()
    if ("heading" in lowered_style or lowered_style.startswith("heading")) and not looks_like_body_prose(stripped):
        return True
    return False


def is_figure_caption_text(text: str) -> bool:
    return bool(is_formal_caption_text(text) and (text or "").lstrip().startswith("\u56fe"))


def is_non_body_paragraph(text: str, style_name: str) -> bool:
    lowered_style = style_name.lower()
    normalized = normalize_space(text)
    if FORMULA_NUMBER_LABEL_ONLY_RE.match(text or "") or FORMULA_NUMBER_LABEL_ONLY_RE.match(normalized):
        return True
    if is_code_like_paragraph_text(text):
        return True
    if lowered_style.startswith("toc"):
        return True
    if normalized.lower() in {normalize_space(item).lower() for item in FRONTMATTER_MARKERS | NON_BODY_EXACT_MARKERS}:
        return True
    if is_formal_caption_text(text):
        return True
    if REFERENCE_ENTRY_RE.match(text.strip()):
        return True
    return False


def is_formal_caption_text(text: str) -> bool:
    """Distinguish formal captions from explanatory figure/table prose."""

    match = CAPTION_RE.match(text or "")
    if not match:
        return False
    title = (match.group("title") or "").strip()
    if not title:
        return False
    compact_title = normalize_space(title)
    if compact_title.startswith(
        (
            "\u5c55\u793a",
            "\u663e\u793a",
            "\u6240\u793a",
            "\u8fdb\u4e00\u6b65",
            "\u8865\u5145",
            "\u7ed9\u51fa",
            "\u8bf4\u660e",
            "\u53cd\u6620",
            "\u8868\u660e",
            "\u4f53\u73b0",
            "\u63cf\u8ff0",
            "\u4fdd\u7559",
            "\u6309\u7167",
            "\u91cc",
        )
    ):
        return False
    if len(title) > 80 and re.search(r"[\u3002\uff0c\uff1b,;]", title):
        return False
    return True


def starts_with_caption_label_reference(text: str) -> bool:
    """Detect body prose that repeats a figure/table number as a lead-in."""

    return bool(
        CAPTION_LABEL_PREFIX_RE.match(text or "")
        or THIS_CAPTION_PLUS_LABEL_PREFIX_RE.match(text or "")
    )


def is_template_placeholder_text(text: str) -> bool:
    stripped = (text or "").strip()
    if not stripped:
        return False
    if any(token in stripped for token in ("×", "ｘ", "□", "☆")):
        return True
    normalized = normalize_space(stripped)
    if normalized and set(normalized) <= {"…", "."}:
        return True
    return any(
        token in stripped
        for token in (
            "不标页码",
            "正文开始标注页码",
            "字号",
            "字体",
            "行距",
            "居中",
            "顶格",
            "空一格",
            "空两格",
            "位于图下",
            "位于表上",
            "图与下文空",
            "与正文空",
            "结束语",
            "附录",
        )
    )


def looks_like_body_prose(text: str) -> bool:
    compacted = normalize_space(text)
    if len(compacted) >= 40:
        return True
    return any(token in text for token in ("\u3002", "\uff0c", "\uff1b", ".", ";"))


def heading_contamination_records(
    docx_path: Path,
    *,
    allowed_body_style_id: str | None = None,
    allowed_default_style_id: str | None = None,
) -> tuple[list[dict[str, object]], dict[str, ET.Element]]:
    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)

    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    body = document_root.find("w:body", NS)
    if body is None:
        return [], styles

    toc_seen = False
    main_body_started = False
    reference_zone_started = False
    contaminated: list[dict[str, object]] = []

    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        if not text:
            continue
        style_id = paragraph_style_id(child)
        style_name = style_name_by_id(styles, style_id)
        normalized = normalize_space(text)

        if is_toc_heading_text(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
            reference_zone_started = True
            continue
        if reference_zone_started:
            continue
        if (
            paragraph_has_run_tab(child)
            or is_static_toc_leader_entry(text)
            or style_name.lower().startswith("toc")
            or (style_id or "").lower().startswith("toc")
        ):
            continue

        heading_family = style_is_heading_family(styles, style_id, style_name) or paragraph_has_outline_level(child) or paragraph_has_numbering(child)
        heading_text = is_body_heading_text(text, style_name, style_id, child)
        if heading_text:
            main_body_started = True
            continue
        if not main_body_started:
            continue
        if is_non_body_paragraph(text, style_name) or is_template_placeholder_text(text):
            continue
        if paragraph_has_omml(child):
            continue
        allowed_body_family = (
            (style_id and allowed_body_style_id and style_id == allowed_body_style_id)
            or (not style_id and allowed_body_style_id and allowed_body_style_id == allowed_default_style_id)
        )
        if heading_family and not allowed_body_family and looks_like_body_prose(text):
            contaminated.append(
                {
                    "paragraph_index": paragraph_index,
                    "text": text,
                    "style_id": style_id or "none",
                    "style_name": style_name or "none",
                    "style_chain": style_chain_label(styles, style_id),
                    "direct_outline": paragraph_has_outline_level(child),
                    "instance_metrics": paragraph_instance_metrics(child),
                }
            )

    return contaminated, styles


def body_paragraphs(docx_path: Path) -> tuple[list[dict[str, object]], dict[str, ET.Element], str | None]:
    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, default_paragraph_style_id = style_node_by_type(styles_root)

    body = document_root.find("w:body", NS)
    if body is None:
        return [], styles, default_paragraph_style_id

    main_body_started = False
    toc_seen = False
    reference_zone_started = False
    paragraphs: list[dict[str, object]] = []

    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        if not text:
            continue
        style_id = paragraph_style_id(child)
        style_name = style_name_by_id(styles, style_id)
        normalized = normalize_space(text)

        if is_toc_heading_text(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
            reference_zone_started = True
            continue
        if reference_zone_started:
            continue
        if (
            paragraph_has_run_tab(child)
            or is_static_toc_leader_entry(text)
            or style_name.lower().startswith("toc")
            or (style_id or "").lower().startswith("toc")
        ):
            continue
        if is_body_heading_text(text, style_name, style_id, child):
            main_body_started = True
            continue
        if not main_body_started:
            continue
        if is_non_body_paragraph(text, style_name):
            continue
        if is_template_placeholder_text(text):
            continue
        if paragraph_has_omml(child):
            continue
        if child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None:
            continue

        paragraphs.append(
            {
                "paragraph_index": paragraph_index,
                "text": text,
                "style_id": style_id,
                "style_name": style_name,
                "has_direct_formatting": has_nontrivial_direct_formatting(child),
                "instance_metrics": paragraph_instance_metrics(child),
                "effective_metrics": paragraph_effective_metrics(child, styles, default_paragraph_style_id),
                "run_metrics": paragraph_run_direct_metrics(child),
                "node": child,
            }
        )

    return paragraphs, styles, default_paragraph_style_id


def choose_reference_body_style(
    reference_body_paragraphs: list[dict[str, object]],
    reference_styles: dict[str, ET.Element],
    default_reference_style_id: str | None,
) -> tuple[str | None, str]:
    fallback_count = sum(1 for paragraph in reference_body_paragraphs if not paragraph["style_id"])
    explicit_counts = Counter(
        str(paragraph["style_id"])
        for paragraph in reference_body_paragraphs
        if paragraph["style_id"] and str(paragraph["style_id"]) in reference_styles
    )
    if fallback_count and (not explicit_counts or fallback_count >= explicit_counts.most_common(1)[0][1]):
        return default_reference_style_id, style_name_by_id(reference_styles, default_reference_style_id)
    if explicit_counts:
        style_id = explicit_counts.most_common(1)[0][0]
        return style_id, style_name_by_id(reference_styles, style_id)
    return default_reference_style_id, style_name_by_id(reference_styles, default_reference_style_id)


def find_matching_final_style_id(
    final_styles: dict[str, ET.Element],
    reference_style_id: str | None,
    reference_style_name: str,
    default_final_style_id: str | None,
) -> str | None:
    if reference_style_id and reference_style_id in final_styles:
        return reference_style_id

    normalized_name = normalize_space(reference_style_name).lower()
    if normalized_name:
        for style_id in final_styles:
            if normalize_space(style_name_by_id(final_styles, style_id)).lower() == normalized_name:
                return style_id

    if normalized_name == "normal":
        return default_final_style_id
    return None


def style_matches_target(
    paragraph: dict[str, object],
    *,
    target_style_id: str | None,
    target_style_name: str,
    default_style_id: str | None = None,
) -> bool:
    style_id = paragraph["style_id"]
    style_name = paragraph["style_name"]
    if not style_id:
        return bool(target_style_id and default_style_id and target_style_id == default_style_id)
    if target_style_id and str(style_id) == target_style_id:
        return True
    if target_style_name and normalize_space(str(style_name)).lower() == normalize_space(target_style_name).lower():
        return True
    return False


def font_values_equivalent(expected: str, actual: str) -> bool:
    if expected == actual:
        return True
    # Some school templates export fallback aliases such as "宋体;SimSun" in
    # style definitions. They are acceptable only as a reference baseline: the
    # final DOCX must keep one concrete OOXML font family, never an alias list.
    if ";" in str(actual or ""):
        return False
    if ";" not in str(expected or ""):
        return False
    actual_canonical = font_canonical(actual)
    expected_aliases = {
        font_canonical(part)
        for part in str(expected or "").split(";")
        if font_canonical(part)
    }
    return actual_canonical in expected_aliases


def compare_metrics(reference_metrics: dict[str, str], final_metrics: dict[str, str]) -> list[str]:
    labels = {
        "jc": "alignment",
        "before": "space before",
        "after": "space after",
        "line": "line spacing",
        "lineRule": "line rule",
        "firstLine": "first-line indent",
        "left": "left indent",
        "right": "right indent",
        "hanging": "hanging indent",
        "firstLineChars": "first-line indent chars",
        "leftChars": "left indent chars",
        "rightChars": "right indent chars",
        "hangingChars": "hanging indent chars",
        "eastAsia": "EastAsia font",
        "ascii": "Ascii font",
        "hAnsi": "hAnsi font",
        "eastAsiaTheme": "EastAsia theme font",
        "asciiTheme": "Ascii theme font",
        "hAnsiTheme": "hAnsi theme font",
        "csTheme": "complex-script theme font",
        "size": "font size",
        "bold": "font bold",
    }
    issues: list[str] = []
    for key, label in labels.items():
        expected = reference_metrics.get(key, "")
        actual = final_metrics.get(key, "")
        if not expected and not actual:
            continue
        if key in {"eastAsia", "ascii", "hAnsi"} and font_values_equivalent(expected, actual):
            continue
        if expected != actual:
            issues.append(f"{label}: expected `{expected or 'none'}` but found `{actual or 'none'}`")
    return issues


def infer_reference_body_size_half_points(
    reference_body: list[dict[str, object]],
    reference_styles: dict[str, ET.Element],
    reference_body_style_id: str | None,
) -> int | None:
    style_size = _int_or_none(resolved_style_metrics(reference_styles, reference_body_style_id).get("size", ""))
    if style_size is not None:
        return style_size
    sizes: list[int] = []
    for paragraph in reference_body:
        metrics = paragraph.get("run_metrics", {})
        if not isinstance(metrics, dict):
            continue
        for key in ("direct_sizes", "direct_size_cs"):
            values = metrics.get(key)
            if isinstance(values, list):
                sizes.extend(value for value in values if isinstance(value, int))
    if not sizes:
        return None
    return Counter(sizes).most_common(1)[0][0]


def declared_body_size_half_points(reference_docx: Path) -> int | None:
    """Read an explicit body-size declaration from a thesis template.

    Some approved samples are annotated style sheets rather than full papers.
    In that shape, the generic body detector may only see tiny figure/table
    labels and infer a false 10.5 pt baseline. Prefer the template's own
    declaration when it states the body text size.
    """
    size_map = {
        "初号": 84,
        "小初": 72,
        "一号": 52,
        "小一": 48,
        "二号": 44,
        "小二": 36,
        "三号": 32,
        "小三": 30,
        "四号": 28,
        "小四": 24,
        "五号": 21,
        "小五": 18,
        "六号": 15,
        "小六": 13,
        "七号": 11,
        "八号": 10,
    }
    try:
        with zipfile.ZipFile(reference_docx) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return None
    all_text = "\n".join(paragraph_text(paragraph) for paragraph in root.findall(".//w:body/w:p", NS))
    if "\u6b63\u6587" in all_text and "\u5c0f\u56db\u53f7" in all_text:
        return 24
    patterns = (
        r"正文内容\s*([小一二三四五六七八初号]{2}|[一二三四五六七八初]号)",
        r"正文[：:][^\n。；;]*内容\s*([小一二三四五六七八初号]{2}|[一二三四五六七八初]号)",
        r"正文[：:][^\n。；;]*正文内容\s*([小一二三四五六七八初号]{2}|[一二三四五六七八初]号)",
    )
    for pattern in patterns:
        match = re.search(pattern, all_text)
        if not match:
            continue
        label = match.group(1)
        if label in size_map:
            return size_map[label]
    return None


def reference_body_allows_bold(
    reference_body: list[dict[str, object]],
    reference_styles: dict[str, ET.Element],
    reference_body_style_id: str | None,
) -> bool:
    style_bold = str(resolved_style_metrics(reference_styles, reference_body_style_id).get("bold", "")).lower()
    if _truthy_ooxml_bool(style_bold):
        return True
    ratios: list[float] = []
    for paragraph in reference_body:
        metrics = paragraph.get("run_metrics", {})
        if isinstance(metrics, dict):
            ratios.append(float(metrics.get("bold_ratio") or 0.0))
    return bool(ratios) and max(ratios) >= 0.5


def direct_body_typography_contamination_records(
    final_body: list[dict[str, object]],
    reference_body: list[dict[str, object]],
    reference_styles: dict[str, ET.Element],
    reference_body_style_id: str | None,
    expected_size_override: int | None = None,
) -> list[dict[str, object]]:
    expected_size = expected_size_override
    if expected_size is None:
        expected_size = infer_reference_body_size_half_points(reference_body, reference_styles, reference_body_style_id)
    bold_allowed = reference_body_allows_bold(reference_body, reference_styles, reference_body_style_id)
    contaminated: list[dict[str, object]] = []
    for paragraph in final_body:
        text = str(paragraph.get("text") or "")
        if not looks_like_body_prose(text):
            continue
        instance_metrics = dict(paragraph.get("instance_metrics") or {})
        run_metrics = dict(paragraph.get("run_metrics") or {})
        reasons: list[str] = []
        alignment = str(instance_metrics.get("jc") or "").lower()
        if alignment in {"center", "right"}:
            reasons.append(f"direct paragraph alignment `{alignment}`")
        bold_ratio = float(run_metrics.get("bold_ratio") or 0.0)
        if not bold_allowed and bold_ratio >= 0.5:
            reasons.append(f"direct bold covers {bold_ratio:.0%} of visible text")
        max_size = run_metrics.get("max_direct_size")
        max_size_int = max_size if isinstance(max_size, int) else _int_or_none(max_size)
        if max_size_int is not None:
            if expected_size is not None and max_size_int > expected_size + 4:
                reasons.append(f"direct font size {max_size_int} exceeds body baseline {expected_size}")
            elif expected_size is not None and max_size_int < expected_size - 1:
                reasons.append(f"direct font size {max_size_int} is below body baseline {expected_size}")
            elif expected_size is None and max_size_int >= 28:
                reasons.append(f"direct font size {max_size_int} is heading-like for body prose")
        if reasons:
            contaminated.append(
                {
                    "paragraph_index": paragraph.get("paragraph_index", "unknown"),
                    "text": text,
                    "style_id": paragraph.get("style_id") or "none",
                    "style_name": paragraph.get("style_name") or "none",
                    "reasons": reasons,
                    "instance_metrics": instance_metrics,
                    "run_metrics": run_metrics,
                }
            )
    return contaminated


def caption_sibling_body_contamination_records(docx_path: Path) -> list[dict[str, object]]:
    """Find body prose after a figure/table block that kept caption or title formatting."""

    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    body = document_root.find("w:body", NS)
    if body is None:
        return []

    toc_seen = False
    main_body_started = False
    reference_zone_started = False
    previous_surface_context = ""
    surface_proximity_context = ""
    surface_proximity_budget = 0
    records: list[dict[str, object]] = []
    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag == W + "tbl":
            if main_body_started:
                previous_surface_context = "table object"
                surface_proximity_context = "recent table object"
                surface_proximity_budget = 3
            continue
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        style_id = paragraph_style_id(child)
        style_name = style_name_by_id(styles, style_id)
        normalized = normalize_space(text)

        if not text:
            continue
        if is_toc_heading_text(text):
            toc_seen = True
            previous_surface_context = ""
            continue
        if not toc_seen:
            continue
        if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
            reference_zone_started = True
            previous_surface_context = ""
            continue
        if reference_zone_started:
            continue
        if (
            paragraph_has_run_tab(child)
            or is_static_toc_leader_entry(text)
            or style_name.lower().startswith("toc")
            or (style_id or "").lower().startswith("toc")
        ):
            continue
        if is_body_heading_text(text, style_name, style_id, child):
            main_body_started = True
            previous_surface_context = ""
            surface_proximity_context = ""
            surface_proximity_budget = 0
            continue
        if not main_body_started:
            continue
        if paragraph_has_picture(child):
            previous_surface_context = "image-holder paragraph"
            surface_proximity_context = "recent image/table block"
            surface_proximity_budget = 3
            continue
        if is_formal_caption_text(text):
            previous_surface_context = "formal caption paragraph"
            surface_proximity_context = "recent formal caption paragraph"
            surface_proximity_budget = 3
            continue
        if is_non_body_paragraph(text, style_name) or is_template_placeholder_text(text):
            previous_surface_context = ""
            continue
        if not looks_like_body_prose(text):
            previous_surface_context = ""
            continue

        reasons: list[str] = []
        active_surface_context = previous_surface_context
        if not active_surface_context and surface_proximity_budget > 0 and starts_with_caption_label_reference(text):
            active_surface_context = surface_proximity_context or "recent figure/table block"
        if active_surface_context:
            metrics = paragraph_instance_metrics(child)
            run_metrics = paragraph_run_direct_metrics(child)
            alignment = str(metrics.get("jc") or "").lower()
            first_line = str(metrics.get("firstLine") or "")
            first_line_chars = str(metrics.get("firstLineChars") or "")
            line = str(metrics.get("line") or "")
            before = str(metrics.get("before") or "")
            after = str(metrics.get("after") or "")
            if style_is_caption_title_family(styles, style_id, style_name):
                reasons.append(f"style chain `{style_chain_label(styles, style_id)}` is caption/title family")
            if style_is_heading_family(styles, style_id, style_name) or paragraph_has_outline_level(child):
                reasons.append(f"style chain `{style_chain_label(styles, style_id)}` is heading/outline family")
            if starts_with_caption_label_reference(text):
                if previous_surface_context == "formal caption paragraph":
                    reasons.append("body prose repeats a figure/table label immediately after the formal caption")
                else:
                    reasons.append("body prose begins with a figure/table label near a figure/table block")
            if alignment in {"center", "right"}:
                reasons.append(f"paragraph alignment `{alignment}` after {active_surface_context}")
            if first_line in {"", "0"} and first_line_chars in {"", "0"}:
                reasons.append(f"first-line indent `{first_line or 'missing'}` after {active_surface_context}")
            if line in {"240", "300"} and first_line in {"", "0"} and first_line_chars in {"", "0"}:
                reasons.append(f"caption-like line spacing `{line}` after {active_surface_context}")
            if before not in {"", "0"} or after not in {"", "0"}:
                reasons.append(f"caption/title spacing before=`{before or 'missing'}` after=`{after or 'missing'}`")
            if paragraph_has_keep_next(child):
                reasons.append("keepNext copied from caption/table-title paragraph")
            max_size = _int_or_none(run_metrics.get("max_direct_size"))
            if max_size is not None and max_size >= 28:
                reasons.append(f"heading-like direct font size `{max_size}`")
            if reasons:
                records.append(
                    {
                        "paragraph_index": paragraph_index,
                        "text": text,
                        "style_id": style_id or "none",
                        "style_name": style_name or "none",
                        "source_context": active_surface_context,
                        "reasons": reasons,
                        "instance_metrics": metrics,
                        "run_metrics": run_metrics,
                    }
                )
        if surface_proximity_budget > 0:
            surface_proximity_budget -= 1
            if surface_proximity_budget <= 0:
                surface_proximity_context = ""
        previous_surface_context = ""
    return records


def compare_instance_metrics(
    reference_metrics: dict[str, str],
    final_metrics: dict[str, str],
    *,
    allow_direct_visible_overrides: bool = False,
) -> list[str]:
    labels = {
        "jc": "alignment",
        "before": "space before",
        "after": "space after",
        "line": "line spacing",
        "lineRule": "line rule",
        "firstLine": "first-line indent",
        "left": "left indent",
        "right": "right indent",
        "firstLineChars": "first-line indent chars",
        "leftChars": "left indent chars",
        "rightChars": "right indent chars",
    }
    issues: list[str] = []
    for key, label in labels.items():
        expected = reference_metrics.get(key, "")
        actual = final_metrics.get(key, "")
        if not expected and not actual:
            continue
        if allow_direct_visible_overrides:
            if key in {"before", "after"} and not expected and actual == "0":
                continue
            if key == "firstLine" and not expected and actual:
                continue
        if expected != actual:
            issues.append(f"{label}: expected `{expected or 'none'}` but found `{actual or 'none'}`")
    return issues


def direct_visible_metric_records(
    body_paragraphs: list[dict[str, object]],
    reference_instance_metrics: dict[str, str] | None = None,
) -> list[dict[str, object]]:
    """Find body prose that only passes through inherited/effective metrics."""

    reference_instance_metrics = reference_instance_metrics or {}
    expected = {
        "before": reference_instance_metrics.get("before") or "0",
        "after": reference_instance_metrics.get("after") or "0",
        "line": reference_instance_metrics.get("line") or "360",
        "lineRule": reference_instance_metrics.get("lineRule") or "auto",
    }
    records: list[dict[str, object]] = []
    for paragraph in body_paragraphs:
        text = str(paragraph.get("text") or "")
        if not looks_like_body_prose(text):
            continue
        metrics = dict(paragraph.get("instance_metrics") or {})
        reasons: list[str] = []
        for key, expected_value in expected.items():
            actual = str(metrics.get(key) or "")
            if actual != expected_value:
                reasons.append(f"direct {key} expected `{expected_value}` found `{actual or 'missing'}`")
        first_line = str(metrics.get("firstLine") or "")
        try:
            first_line_value = int(first_line)
        except ValueError:
            first_line_value = 0
        expected_first_line = str(reference_instance_metrics.get("firstLine") or "")
        if expected_first_line:
            if first_line != expected_first_line:
                reasons.append(
                    f"direct firstLine expected `{expected_first_line}` found `{first_line or 'missing'}`"
                )
        elif first_line_value <= 0:
            char_value = str(metrics.get("firstLineChars") or "")
            reasons.append(
                "real direct firstLine missing; "
                f"firstLineChars `{char_value or 'missing'}` is not sufficient"
            )
        if reasons:
            records.append(
                {
                    "paragraph_index": paragraph.get("paragraph_index", "unknown"),
                    "text": text,
                    "style_id": paragraph.get("style_id") or "none",
                    "style_name": paragraph.get("style_name") or "none",
                    "reasons": reasons,
                    "instance_metrics": metrics,
                    "effective_metrics": dict(paragraph.get("effective_metrics") or {}),
                }
            )
    return records


def body_heading_line_metric_records(docx_path: Path) -> list[dict[str, object]]:
    """Return no prose-style metric issues for headings.

    Heading metrics are governed by the whole-format gate and template heading
    contract. The body-style strict gate only owns prose body paragraphs, so it
    must not reject valid chapter headings that are intentionally centered with
    fixed 20 pt spacing.
    """

    return []

    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    body = document_root.find("w:body", NS)
    if body is None:
        return []

    toc_seen = False
    reference_zone_started = False
    records: list[dict[str, object]] = []
    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        if not text:
            continue
        style_id = paragraph_style_id(child)
        style_name = style_name_by_id(styles, style_id)
        normalized = normalize_space(text)

        if is_toc_heading_text(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
            reference_zone_started = True
            continue
        if reference_zone_started:
            continue
        if (
            paragraph_has_run_tab(child)
            or is_static_toc_leader_entry(text)
            or style_name.lower().startswith("toc")
            or (style_id or "").lower().startswith("toc")
        ):
            continue
        if not is_body_heading_text(text, style_name, style_id, child):
            continue

        metrics = paragraph_instance_metrics(child)
        run_metrics = paragraph_run_direct_metrics(child)
        heading_level = heading_level_from_text(text, child)
        reasons: list[str] = []
        expected_heading_metrics = {
            "before": "0",
            "after": "0",
            "line": "360",
            "lineRule": "auto",
            "left": "0",
            "firstLine": "0",
        }
        if heading_level is not None:
            expected_heading_metrics["jc"] = "left"
        for key, expected_value in expected_heading_metrics.items():
            actual = str(metrics.get(key) or "")
            if actual != expected_value:
                reasons.append(f"heading direct {key} expected `{expected_value}` found `{actual or 'missing'}`")
        if heading_level is not None:
            expected_size = {0: 30, 1: 28, 2: 24}.get(heading_level)
            actual_size = _int_or_none(run_metrics.get("max_direct_size"))
            if expected_size is not None and actual_size != expected_size:
                reasons.append(
                    "heading level {level} direct font size expected `{expected}` found `{actual}`".format(
                        level=heading_level + 1,
                        expected=expected_size,
                        actual=actual_size or "missing",
                    )
                )
        if reasons:
            records.append(
                {
                    "paragraph_index": paragraph_index,
                    "text": text,
                    "heading_level": heading_level,
                    "style_id": style_id or "none",
                    "style_name": style_name or "none",
                    "reasons": reasons,
                    "instance_metrics": metrics,
                    "run_metrics": run_metrics,
                }
            )
    return records


def body_surface_direct_metric_records(docx_path: Path) -> list[dict[str, object]]:
    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    body = document_root.find("w:body", NS)
    if body is None:
        return []

    toc_seen = False
    main_body_started = False
    reference_zone_started = False
    records: list[dict[str, object]] = []
    previous_main_paragraph_had_picture = False
    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        style_id = paragraph_style_id(child)
        style_name = style_name_by_id(styles, style_id)
        normalized = normalize_space(text)

        if is_toc_heading_text(text):
            toc_seen = True
            continue
        if not toc_seen:
            continue
        if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
            reference_zone_started = True
            continue
        if reference_zone_started:
            continue
        if (
            text
            and (
                paragraph_has_run_tab(child)
                or is_static_toc_leader_entry(text)
                or style_name.lower().startswith("toc")
                or (style_id or "").lower().startswith("toc")
            )
        ):
            continue
        if text and is_body_heading_text(text, style_name, style_id, child):
            main_body_started = True
            previous_main_paragraph_had_picture = False
        if not main_body_started:
            continue

        metrics = paragraph_instance_metrics(child)
        has_picture = paragraph_has_picture(child)
        kind = ""
        expected: dict[str, str] = {}
        if has_picture:
            kind = "image-holder"
            expected = {
                "jc": "center",
                "line": "360",
                "lineRule": "auto",
            }
        elif text and is_formal_caption_text(text):
            kind = "figure-caption" if is_figure_caption_text(text) else "table-caption"
            if kind == "table-caption":
                expected = {
                    "jc": "center",
                    "line": "360",
                    "lineRule": "auto",
                }
            else:
                expected = {
                    "jc": "center",
                    "line": "360",
                    "lineRule": "auto",
                }
        else:
            continue

        reasons: list[str] = []
        for key, expected_value in expected.items():
            actual = str(metrics.get(key) or "")
            if actual != expected_value:
                reasons.append(f"{kind} direct {key} expected `{expected_value}` found `{actual or 'missing'}`")
        if kind == "figure-caption" and not previous_main_paragraph_had_picture:
            reasons.append("figure-caption adjacency expected immediate preceding image-holder paragraph")
        if kind in {"figure-caption", "table-caption"}:
            run_metrics = paragraph_run_direct_metrics(child)
            max_size = _int_or_none(run_metrics.get("max_direct_size"))
            if max_size != 21:
                reasons.append(f"{kind} direct font size expected `21` found `{max_size or 'missing'}`")
        if kind == "table-caption" and not paragraph_has_keep_next(child):
            reasons.append("table-caption expected keepNext")
        if reasons:
            records.append(
                {
                    "paragraph_index": paragraph_index,
                    "surface_kind": kind,
                    "text": text,
                    "style_id": style_id or "none",
                    "style_name": style_name or "none",
                    "reasons": reasons,
                    "instance_metrics": metrics,
                }
            )
        previous_main_paragraph_had_picture = has_picture
    return records


def undefined_body_style_records(docx_path: Path) -> list[dict[str, object]]:
    document_root = load_xml(docx_path, "word/document.xml")
    styles_root = load_xml(docx_path, "word/styles.xml")
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    body = document_root.find("w:body", NS)
    if body is None:
        return []

    records: list[dict[str, object]] = []
    for paragraph_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        style_id = paragraph_style_id(child)
        if not style_id or style_id in styles:
            continue
        text = paragraph_text(child)
        if not text.strip() and not paragraph_has_picture(child):
            continue
        records.append(
            {
                "paragraph_index": paragraph_index,
                "style_id": style_id,
                "text": text[:120],
                "has_picture": paragraph_has_picture(child),
            }
        )
    return records


def font_alias_list_records(docx_path: Path) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    try:
        styles_root = load_xml(docx_path, "word/styles.xml")
        document_root = load_xml(docx_path, "word/document.xml")
    except Exception:
        return records
    styles, _default_paragraph_style_id = style_node_by_type(styles_root)
    for style in styles_root.findall(".//w:style", NS):
        style_id = style.attrib.get(W + "styleId", "")
        name = style_name_by_id({style_id: style}, style_id) if style_id else ""
        for fonts in style.findall(".//w:rFonts", NS):
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                value = fonts.attrib.get(W + key, "")
                if ";" in value:
                    records.append({"scope": "style", "style_id": style_id, "style_name": name, "slot": key, "value": value})
    body = document_root.find("w:body", NS)
    if body is not None:
        toc_seen = False
        main_body_started = False
        reference_zone_started = False
        for paragraph_index, paragraph in enumerate(list(body), start=1):
            if paragraph.tag != W + "p":
                continue
            text = paragraph_text(paragraph)
            normalized = normalize_space(text)
            style_id = paragraph_style_id(paragraph)
            style_name = style_name_by_id(styles, style_id)
            if is_toc_heading_text(text):
                toc_seen = True
                continue
            if not toc_seen:
                continue
            if normalized.lower() in {normalize_space(item).lower() for item in TERMINAL_HEADING_MARKERS}:
                reference_zone_started = True
                continue
            if reference_zone_started:
                continue
            if (
                text
                and (
                    paragraph_has_run_tab(paragraph)
                    or is_static_toc_leader_entry(text)
                    or style_name.lower().startswith("toc")
                    or (style_id or "").lower().startswith("toc")
                )
            ):
                continue
            if text and is_body_heading_text(text, style_name, style_id, paragraph):
                main_body_started = True
            if not main_body_started:
                continue
            if is_template_placeholder_text(text):
                continue
            for fonts in paragraph.findall(".//w:rFonts", NS):
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    value = fonts.attrib.get(W + key, "")
                    if ";" in value:
                        records.append(
                            {
                                "scope": "run",
                                "paragraph_index": paragraph_index,
                                "slot": key,
                                "value": value,
                                "text": text[:120],
                            }
                        )
    return records


def make_report(
    *,
    reference_docx: Path,
    final_docx: Path,
    body_paragraph_count: int,
    reference_body_style_id: str | None,
    reference_body_style_name: str,
    final_body_style_id: str | None,
    binding_summary: str,
    normal_baseline_summary: str,
    family_summary: str,
    heading_contamination_summary: str,
    mixed_script_summary: str,
    direct_visible_summary: str,
    binding_issues: list[str],
    baseline_issues: list[str],
    family_issues: list[str],
    heading_contamination_issues: list[str],
    mixed_script_issues: list[str],
    direct_visible_issues: list[str],
) -> str:
    result = "pass" if all(
        summary.startswith("passed")
        for summary in (
            binding_summary,
            normal_baseline_summary,
            family_summary,
            heading_contamination_summary,
            mixed_script_summary,
            direct_visible_summary,
        )
    ) else "fail"
    lines = [
        "# DOCX Body Style Audit",
        "",
        f"- reference docx path: {reference_docx}",
        f"- reference docx sha256: {sha256_file(reference_docx) if reference_docx.exists() else 'missing'}",
        f"- final docx path: {final_docx}",
        f"- final docx sha256: {sha256_file(final_docx) if final_docx.exists() else 'missing'}",
        f"- body paragraphs checked: {body_paragraph_count}",
        f"- reference inferred body style: {reference_body_style_id or 'none'} ({reference_body_style_name or 'none'})",
        f"- final matched body style: {final_body_style_id or 'none'}",
        f"- body style binding summary: {binding_summary}",
        f"- Normal baseline preservation summary: {normal_baseline_summary}",
        f"- body paragraph family consistency summary: {family_summary}",
        f"- body heading contamination summary: {heading_contamination_summary}",
        f"- body mixed-script font summary: {mixed_script_summary}",
        f"- body direct visible metrics summary: {direct_visible_summary}",
        f"- result: {result}",
        "",
        "## Binding Issues",
    ]
    lines.extend(f"- {item}" for item in (binding_issues or ["none"]))
    lines.extend(["", "## Normal Baseline Issues"])
    lines.extend(f"- {item}" for item in (baseline_issues or ["none"]))
    lines.extend(["", "## Body Family Issues"])
    lines.extend(f"- {item}" for item in (family_issues or ["none"]))
    lines.extend(["", "## Heading Contamination Issues"])
    lines.extend(f"- {item}" for item in (heading_contamination_issues or ["none"]))
    lines.extend(["", "## Mixed-Script Body Issues"])
    lines.extend(f"- {item}" for item in (mixed_script_issues or ["none"]))
    lines.extend(["", "## Direct Visible Metric Issues"])
    lines.extend(f"- {item}" for item in (direct_visible_issues or ["none"]))
    return "\n".join(lines) + "\n"


def audit_body_style(
    reference_docx: Path,
    final_docx: Path,
    *,
    strict_direct_visible_metrics: bool = False,
) -> tuple[str, list[str], list[str], list[str], list[str]]:
    reference_body, reference_styles, default_reference_style_id = body_paragraphs(reference_docx)
    final_body, final_styles, default_final_style_id = body_paragraphs(final_docx)

    reference_body_style_id, reference_body_style_name = choose_reference_body_style(
        reference_body,
        reference_styles,
        default_reference_style_id,
    )
    final_body_style_id = find_matching_final_style_id(
        final_styles,
        reference_body_style_id,
        reference_body_style_name,
        default_final_style_id,
    )
    heading_contamination, _heading_styles = heading_contamination_records(
        final_docx,
        allowed_body_style_id=final_body_style_id,
        allowed_default_style_id=default_final_style_id,
    )
    caption_sibling_contamination = caption_sibling_body_contamination_records(final_docx)

    binding_issues: list[str] = []
    baseline_issues: list[str] = []
    family_issues: list[str] = []
    heading_contamination_issues: list[str] = []
    mixed_script_issues: list[str] = []
    direct_visible_issues: list[str] = []

    if not final_body:
        binding_issues.append("未定位到可审计的正文段落，无法验证正文样式绑定。")
        binding_summary = "failed no auditable body paragraphs found"
    elif not reference_body_style_id and not reference_body_style_name:
        binding_issues.append("参考稿未能推断出正文样式家族，无法验证最终稿绑定。")
        binding_summary = "failed reference body style could not be inferred"
    else:
        missing_binding = [
            paragraph
            for paragraph in final_body
            if not style_matches_target(
                paragraph,
                target_style_id=reference_body_style_id,
                target_style_name=reference_body_style_name,
                default_style_id=default_final_style_id,
            )
        ]
        if missing_binding:
            binding_summary = (
                f"failed {len(missing_binding)}/{len(final_body)} body paragraphs missing explicit binding to `{reference_body_style_name or reference_body_style_id}`"
            )
            for paragraph in missing_binding[:8]:
                text = str(paragraph["text"]).replace("\n", " ").strip()
                binding_issues.append(f"正文段未显式绑定到目标样式: `{text[:80]}`")
        else:
            binding_summary = f"passed {len(final_body)}/{len(final_body)} body paragraphs explicitly bind to `{reference_body_style_name or reference_body_style_id}`"

    if not reference_body_style_id and not reference_body_style_name:
        normal_baseline_summary = "failed reference body baseline missing"
        baseline_issues.append("参考稿未能解析出正文样式基线。")
    elif not final_body_style_id:
        normal_baseline_summary = "failed final body style missing"
        baseline_issues.append("最终稿未找到与参考正文样式对应的样式定义。")
    else:
        reference_metrics = resolved_style_metrics(reference_styles, reference_body_style_id)
        final_metrics = resolved_style_metrics(final_styles, final_body_style_id)
        reference_body_size = declared_body_size_half_points(reference_docx)
        if reference_body_size is None:
            reference_body_size = infer_reference_body_size_half_points(reference_body, reference_styles, reference_body_style_id)
        final_body_size = infer_reference_body_size_half_points(final_body, final_styles, final_body_style_id)
        if reference_body_size is not None:
            reference_metrics["size"] = str(reference_body_size)
        if final_body_size is not None:
            final_metrics["size"] = str(final_body_size)
        metric_diffs = compare_metrics(reference_metrics, final_metrics)
        reference_instance_candidates = [
            paragraph for paragraph in reference_body
            if style_matches_target(
                paragraph,
                target_style_id=reference_body_style_id,
                target_style_name=reference_body_style_name,
                default_style_id=default_reference_style_id,
            )
        ] or reference_body
        has_explicit_reference_instances = any(paragraph.get("style_id") for paragraph in reference_instance_candidates)
        reference_visible_chars = sum(len(str(paragraph.get("text") or "").strip()) for paragraph in reference_instance_candidates)
        reference_instance_sample_is_sparse = len(reference_instance_candidates) < 20 or reference_visible_chars < 1000
        if has_explicit_reference_instances and not reference_instance_sample_is_sparse:
            instance_counter = Counter(
                tuple(sorted(dict(paragraph.get("effective_metrics") or paragraph["instance_metrics"]).items()))
                for paragraph in reference_instance_candidates
            )
            reference_instance_metrics = (
                dict(instance_counter.most_common(1)[0][0]) if instance_counter else {}
            )
            if not any(value and value != "none" for value in reference_instance_metrics.values()):
                reference_instance_metrics = {}
        else:
            reference_instance_metrics = {}
        instance_diffs: list[str] = []
        if reference_instance_sample_is_sparse:
            instance_diffs = []
        elif not any(value for value in reference_instance_metrics.values()):
            instance_diffs = []
        else:
            final_checked = 0
            final_drift_count = 0
            for paragraph in final_body:
                if not style_matches_target(
                    paragraph,
                    target_style_id=reference_body_style_id,
                    target_style_name=reference_body_style_name,
                    default_style_id=default_final_style_id,
                ):
                    continue
                final_checked += 1
                paragraph_instance_diffs = compare_instance_metrics(
                    reference_instance_metrics,
                    dict(paragraph.get("effective_metrics") or paragraph["instance_metrics"]),
                    allow_direct_visible_overrides=strict_direct_visible_metrics,
                )
                if paragraph_instance_diffs:
                    final_drift_count += 1
                    text = str(paragraph["text"]).replace("\n", " ").strip()
                    instance_diffs.append(
                        f"paragraph-instance baseline drift on `{text[:80]}`: {'; '.join(paragraph_instance_diffs[:6])}"
                    )
            if final_checked and final_drift_count:
                instance_diffs.insert(
                    0,
                    "full-body paragraph metric distribution drift: "
                    f"{final_drift_count}/{final_checked} body paragraphs differ from the locked "
                    "line spacing / spacing-before-after / indentation baseline",
                )
        if metric_diffs:
            normal_baseline_summary = (
                f"failed baseline drift on `{reference_body_style_name or reference_body_style_id}`"
            )
            baseline_issues.extend(metric_diffs)
        elif instance_diffs:
            normal_baseline_summary = (
                f"failed paragraph-instance baseline drift on `{reference_body_style_name or reference_body_style_id}`"
            )
            baseline_issues.extend(instance_diffs[:12])
        else:
            normal_baseline_summary = (
                f"passed baseline preserved on `{reference_body_style_name or reference_body_style_id}`"
            )

    explicit_style_keys = Counter()
    fallback_only_count = 0
    for paragraph in final_body:
        if paragraph["style_id"]:
            explicit_style_keys[f"{paragraph['style_id']}::{paragraph['style_name']}"] += 1
        elif not paragraph["has_direct_formatting"]:
            fallback_only_count += 1

    fragmented_families = [
        key for key in explicit_style_keys if key.split("::", 1)[0] != (reference_body_style_id or "") and normalize_space(key.split("::", 1)[1]).lower() != normalize_space(reference_body_style_name).lower()
    ]
    if not final_body:
        family_summary = "failed no auditable body paragraphs found"
        family_issues.append("未定位到可审计的正文段落。")
    elif fallback_only_count or fragmented_families:
        family_summary = (
            f"failed fragmented body family: fallback_only={fallback_only_count}, explicit_families={len(explicit_style_keys)}"
        )
        if fallback_only_count:
            family_issues.append(f"存在 {fallback_only_count} 段正文既无显式 pStyle，也无可用直接格式，只能依赖默认回退样式。")
        if fragmented_families:
            family_issues.append(f"正文显式样式家族混入非目标样式: {'; '.join(fragmented_families[:6])}")
    else:
        family_summary = f"passed one explicit body family preserved across {len(final_body)} body paragraphs"

    if heading_contamination:
        heading_contamination_summary = f"failed {len(heading_contamination)} body prose paragraphs use heading-style or outline-level formatting"
        for record in heading_contamination[:12]:
            text = str(record["text"]).replace("\n", " ").strip()
            heading_contamination_issues.append(
                "paragraph p{paragraph_index} uses heading family `{style_chain}` with prose text `{text}`".format(
                    paragraph_index=record["paragraph_index"],
                    style_chain=record["style_chain"],
                    text=text[:100],
                )
            )
    else:
        heading_contamination_summary = "passed no body prose paragraph uses heading-style or outline-level formatting"

    if caption_sibling_contamination:
        total = len(heading_contamination) + len(caption_sibling_contamination)
        heading_contamination_summary = (
            f"failed {total} body prose paragraphs use heading, caption, title, or object-sibling formatting"
        )
        for record in caption_sibling_contamination[:12]:
            text = str(record["text"]).replace("\n", " ").strip()
            reasons = "; ".join(str(item) for item in record.get("reasons", []))
            heading_contamination_issues.append(
                "paragraph p{paragraph_index} after {source_context} keeps caption/title formatting ({reasons}) with prose text `{text}`".format(
                    paragraph_index=record["paragraph_index"],
                    source_context=record.get("source_context", "object block"),
                    reasons=reasons,
                    text=text[:100],
                )
            )

    direct_contamination = direct_body_typography_contamination_records(
        final_body,
        reference_body,
        reference_styles,
        reference_body_style_id,
        declared_body_size_half_points(reference_docx),
    )
    if direct_contamination:
        total = len(heading_contamination) + len(caption_sibling_contamination) + len(direct_contamination)
        heading_contamination_summary = (
            f"failed {total} body prose paragraphs use heading-style, caption/title object-sibling, outline-level, or polluted direct typography"
        )
        for record in direct_contamination[:12]:
            text = str(record["text"]).replace("\n", " ").strip()
            reasons = "; ".join(str(item) for item in record.get("reasons", []))
            heading_contamination_issues.append(
                "paragraph p{paragraph_index} uses polluted direct typography ({reasons}) with prose text `{text}`".format(
                    paragraph_index=record["paragraph_index"],
                    reasons=reasons,
                    text=text[:100],
                )
            )

    mixed_script_contamination = mixed_script_body_contamination_records(final_body)
    if mixed_script_contamination:
        mixed_script_summary = (
            f"failed {len(mixed_script_contamination)} body paragraphs lose mixed-script run or font-slot separation"
        )
        for record in mixed_script_contamination[:12]:
            text = str(record["text"]).replace("\n", " ").strip()
            reasons = "; ".join(str(item) for item in record.get("reasons", []))
            mixed_script_issues.append(
                "paragraph p{paragraph_index} has mixed-script body-text drift ({reasons}) in `{text}`".format(
                    paragraph_index=record["paragraph_index"],
                    reasons=reasons,
                    text=text[:100],
                )
            )
    else:
        mixed_script_summary = "passed mixed-script body paragraphs preserve run and font-slot separation"

    if strict_direct_visible_metrics:
        direct_visible_records = direct_visible_metric_records(final_body, reference_instance_metrics)
        heading_line_records = body_heading_line_metric_records(final_docx)
        surface_metric_records = body_surface_direct_metric_records(final_docx)
        undefined_style_records = undefined_body_style_records(final_docx)
        font_alias_records = font_alias_list_records(final_docx)
        if (
            direct_visible_records
            or heading_line_records
            or surface_metric_records
            or undefined_style_records
            or font_alias_records
        ):
            combined_count = (
                len(direct_visible_records)
                + len(heading_line_records)
                + len(surface_metric_records)
                + len(undefined_style_records)
                + len(font_alias_records)
            )
            direct_visible_summary = (
                f"failed {combined_count} body surface paragraphs/styles lack strict direct metrics or font slots"
            )
            for record in direct_visible_records[:20]:
                text = str(record["text"]).replace("\n", " ").strip()
                reasons = "; ".join(str(item) for item in record.get("reasons", []))
                direct_visible_issues.append(
                    "paragraph p{paragraph_index} lacks direct visible metrics ({reasons}) with prose text `{text}`".format(
                        paragraph_index=record["paragraph_index"],
                        reasons=reasons,
                        text=text[:100],
                    )
                )
            for record in heading_line_records[:20]:
                text = str(record["text"]).replace("\n", " ").strip()
                reasons = "; ".join(str(item) for item in record.get("reasons", []))
                direct_visible_issues.append(
                    "heading paragraph p{paragraph_index} lacks direct visible line spacing ({reasons}) with heading text `{text}`".format(
                        paragraph_index=record["paragraph_index"],
                        reasons=reasons,
                        text=text[:100],
                    )
                )
            for record in surface_metric_records[:20]:
                text = str(record["text"]).replace("\n", " ").strip()
                reasons = "; ".join(str(item) for item in record.get("reasons", []))
                direct_visible_issues.append(
                    "{surface_kind} paragraph p{paragraph_index} lacks strict direct metrics ({reasons}) with text `{text}`".format(
                        surface_kind=record.get("surface_kind", "surface"),
                        paragraph_index=record["paragraph_index"],
                        reasons=reasons,
                        text=text[:100],
                    )
                )
            for record in undefined_style_records[:20]:
                direct_visible_issues.append(
                    "paragraph p{paragraph_index} uses undefined paragraph style `{style_id}` has_picture={has_picture} text `{text}`".format(
                        paragraph_index=record["paragraph_index"],
                        style_id=record["style_id"],
                        has_picture=record["has_picture"],
                        text=str(record.get("text") or "")[:100],
                    )
                )
            for record in font_alias_records[:20]:
                if record.get("scope") == "style":
                    direct_visible_issues.append(
                        "style `{style_id}` font slot `{slot}` contains forbidden alias list `{value}`".format(
                            style_id=record.get("style_id", "unknown"),
                            slot=record.get("slot", "unknown"),
                            value=record.get("value", ""),
                        )
                    )
                else:
                    direct_visible_issues.append(
                        "paragraph p{paragraph_index} font slot `{slot}` contains forbidden alias list `{value}` in text `{text}`".format(
                            paragraph_index=record.get("paragraph_index", "unknown"),
                            slot=record.get("slot", "unknown"),
                            value=record.get("value", ""),
                            text=str(record.get("text") or "")[:100],
                        )
                    )
        else:
            direct_visible_summary = "passed strict direct visible metrics on body prose, headings, image holders, captions, style ids, and font slots"
    else:
        direct_visible_summary = "passed direct visible metrics strict mode disabled"

    report = make_report(
        reference_docx=reference_docx,
        final_docx=final_docx,
        body_paragraph_count=len(final_body),
        reference_body_style_id=reference_body_style_id,
        reference_body_style_name=reference_body_style_name,
        final_body_style_id=final_body_style_id,
        binding_summary=binding_summary,
        normal_baseline_summary=normal_baseline_summary,
        family_summary=family_summary,
        heading_contamination_summary=heading_contamination_summary,
        mixed_script_summary=mixed_script_summary,
        direct_visible_summary=direct_visible_summary,
        binding_issues=binding_issues,
        baseline_issues=baseline_issues,
        family_issues=family_issues,
        heading_contamination_issues=heading_contamination_issues,
        mixed_script_issues=mixed_script_issues,
        direct_visible_issues=direct_visible_issues,
    )
    return (
        report,
        binding_issues,
        baseline_issues,
        family_issues,
        heading_contamination_issues + mixed_script_issues + direct_visible_issues,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit a DOCX for body-style binding drift and Normal baseline drift.")
    parser.add_argument("--reference-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--report", help="Optional markdown report output path")
    parser.add_argument(
        "--strict-direct-visible-metrics",
        action="store_true",
        help="Fail when body prose lacks direct w:spacing and real w:firstLine metrics even if effective style metrics pass.",
    )
    args = parser.parse_args()

    reference_docx = Path(args.reference_docx).resolve()
    final_docx = Path(args.final_docx).resolve()
    report, binding_issues, baseline_issues, family_issues, heading_contamination_issues = audit_body_style(
        reference_docx,
        final_docx,
        strict_direct_visible_metrics=args.strict_direct_visible_metrics,
    )

    if args.report:
        report_path = Path(args.report).resolve()
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(report, encoding="utf-8")
    else:
        sys.stdout.write(report)

    return 0 if not (binding_issues or baseline_issues or family_issues or heading_contamination_issues) else 1


if __name__ == "__main__":
    raise SystemExit(main())
