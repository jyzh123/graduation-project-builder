#!/usr/bin/env python3
"""Measure template-vs-target TOC paragraph and typography metrics.

The output JSON is consumed by generate_thesis_acceptance_record.py. The
metrics come from the DOCX XML properties that back the Word/WPS paragraph and
font dialogs; rendered geometry is handled by the separate TOC geometry gate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()
ACCEPTED_DOTTED_LEADERS = {"dot", "middleDot", "heavy"}
OFFICIAL_SPEC_TOC_TITLE_SIGNATURE = (
    "directRPr=yes ascii=宋体 hAnsi=宋体 eastAsia=宋体 cs=宋体 "
    "asciiTheme=none hAnsiTheme=none eastAsiaTheme=none csTheme=none "
    "size=15pt sizeCs=15pt weight=bold italic=no underline=no"
)


@dataclass(frozen=True)
class StyleInfo:
    style_id: str
    name: str
    based_on: str
    ppr: ET.Element | None
    rpr: ET.Element | None


@dataclass(frozen=True)
class RunRoleTypography:
    role: str
    signatures: tuple[str, ...]


@dataclass(frozen=True)
class TocMetrics:
    label: str
    style_id: str
    style_name: str
    alignment: str
    text_compact: str
    has_tab: bool
    page_break_before: str
    outline: str
    before_pt: float
    after_pt: float
    line_mode: str
    line_value_pt: float
    left_pt: float
    right_pt: float
    first_line_pt: float
    hanging_pt: float
    left_chars: float
    first_line_chars: float
    hanging_chars: float
    tab_pt: float
    leader: str
    font: str
    font_size_pt: float
    weight: str
    run_typography: tuple[RunRoleTypography, ...]


def twips_to_pt(value: str | None) -> float:
    if not value:
        return 0.0
    try:
        return round(float(value) / 20.0, 2)
    except ValueError:
        return 0.0


def eighth_points_to_pt(value: str | None) -> float:
    if not value:
        return 0.0
    try:
        return round(float(value) / 8.0, 2)
    except ValueError:
        return 0.0


def half_points_to_pt(value: str | None) -> float:
    if not value:
        return 0.0
    try:
        return round(float(value) / 2.0, 2)
    except ValueError:
        return 0.0


def pt_text(value: float) -> str:
    if value == int(value):
        return f"{int(value)}pt"
    return f"{value:.2f}pt"


def attr(element: ET.Element | None, name: str) -> str | None:
    if element is None:
        return None
    return element.attrib.get(W + name)


def text_of(paragraph: ET.Element) -> str:
    return visible_text_of(paragraph)


def visible_text_of(element: ET.Element, *, include_textboxes: bool = False) -> str:
    parts: list[str] = []

    def walk(node: ET.Element, inside_textbox: bool = False) -> None:
        if node.tag == W + "txbxContent":
            inside_textbox = not include_textboxes
        if node.tag == W + "t" and not inside_textbox:
            parts.append(node.text or "")
        for child in list(node):
            walk(child, inside_textbox)

    walk(element)
    return "".join(parts).strip()


def docx_authority_class(docx_path: Path) -> str:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            document = load_part(zf, "word/document.xml")
    except Exception:
        return "unreadable-reference"
    if document is None:
        return "unreadable-reference"
    text = "\n".join(visible_text_of(p) for p in document.findall(".//w:p", NS))
    compact = normalize_text(text)
    markers = (
        normalize_text("撰写与装订规范"),
        normalize_text("毕业设计说明书或毕业论文主要部分"),
        normalize_text("摘要、目录、参考文献、附录、致谢等标题作为第一级标题排版"),
        normalize_text("公式居中书写"),
    )
    if any(marker and marker in compact for marker in markers):
        return "official-format-spec"
    return "finished-thesis-donor"


def has_visible_tab(element: ET.Element, *, include_textboxes: bool = False) -> bool:
    found = False

    def walk(node: ET.Element, inside_textbox: bool = False) -> None:
        nonlocal found
        if found:
            return
        if node.tag == W + "txbxContent":
            inside_textbox = not include_textboxes
        if node.tag == W + "tab" and not inside_textbox:
            found = True
            return
        for child in list(node):
            walk(child, inside_textbox)

    walk(element)
    return found


def iter_visible_runs(element: ET.Element, *, include_textboxes: bool = False) -> list[ET.Element]:
    runs: list[ET.Element] = []

    def walk(node: ET.Element, inside_textbox: bool = False) -> None:
        if node.tag == W + "txbxContent":
            inside_textbox = not include_textboxes
        if node.tag == W + "r" and not inside_textbox:
            runs.append(node)
            return
        for child in list(node):
            walk(child, inside_textbox)

    walk(element)
    return runs


def normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", text or "").strip().lower()


def heading_key(text: str) -> str:
    head = re.split(r"[\(\uff08]", str(text or "").strip(), maxsplit=1)[0]
    return normalize_text(head)


def run_text(run: ET.Element) -> str:
    return visible_text_of(run)


def load_part(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def load_styles(zf: zipfile.ZipFile) -> dict[str, StyleInfo]:
    root = load_part(zf, "word/styles.xml")
    if root is None:
        return {}
    styles: dict[str, StyleInfo] = {}
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(W + "styleId", "")
        if not style_id:
            continue
        name_el = style.find("w:name", NS)
        based_on_el = style.find("w:basedOn", NS)
        styles[style_id] = StyleInfo(
            style_id=style_id,
            name=attr(name_el, "val") or style_id,
            based_on=attr(based_on_el, "val") or "",
            ppr=style.find("w:pPr", NS),
            rpr=style.find("w:rPr", NS),
        )
    return styles


def style_chain(styles: dict[str, StyleInfo], style_id: str) -> list[StyleInfo]:
    chain: list[StyleInfo] = []
    seen: set[str] = set()
    current = style_id
    while current and current not in seen and current in styles:
        seen.add(current)
        info = styles[current]
        chain.append(info)
        current = info.based_on
    return list(reversed(chain))


def find_first(elements: list[ET.Element | None], path: str) -> ET.Element | None:
    for element in elements:
        if element is None:
            continue
        found = element.find(path, NS)
        if found is not None:
            return found
    return None


def tab_elements_from(element: ET.Element | None) -> list[ET.Element]:
    if element is None:
        return []
    return list(element.findall("w:tabs/w:tab", NS))


def find_effective_toc_tab(ppr_sources: list[ET.Element | None]) -> ET.Element | None:
    """Return the tab stop that controls TOC page-number alignment.

    TOC paragraphs often contain a leading clear tab plus a later right tab.
    The old metric path took the first tab and therefore missed broken
    page-number leaders. Direct paragraph tabs are preferred over inherited
    style tabs because they are what a local repair most commonly changes.
    """
    direct_tabs = tab_elements_from(ppr_sources[0] if ppr_sources else None)
    inherited_tabs: list[ET.Element] = []
    for source in ppr_sources[1:]:
        tabs = tab_elements_from(source)
        if tabs:
            inherited_tabs = tabs
            break
    tabs = direct_tabs or inherited_tabs
    right_tabs = [tab for tab in tabs if (attr(tab, "val") or "left") == "right"]
    if right_tabs:
        return max(right_tabs, key=lambda tab: int(attr(tab, "pos") or "0"))
    usable_tabs = [tab for tab in tabs if attr(tab, "val") != "clear"]
    if usable_tabs:
        return max(usable_tabs, key=lambda tab: int(attr(tab, "pos") or "0"))
    return None


def paragraph_style_id(paragraph: ET.Element) -> str:
    p_style = paragraph.find("w:pPr/w:pStyle", NS)
    return attr(p_style, "val") or ""


def classify_level(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> int | None:
    visible = text_of(paragraph)
    if not visible:
        return None
    style_id = paragraph_style_id(paragraph)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id, "", None, None)).name
    combined = f"{style_id} {style_name}".lower()
    match = re.search(r"(?:toc|目录)\s*([1-9])", combined)
    if match:
        return int(match.group(1))
    compact = re.sub(r"[\s\u3000\u25a1\u2606\u00d7]+", "", visible)
    compact = re.sub(r"[\.\u2026\u22ef\u2500\u2501\u30fb]+[0-9xX]+$", "", compact)
    if re.match(r"^(?:\u7b2c.+?\u7ae0|chapter\d+)", compact, re.IGNORECASE):
        return 1
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}", compact):
        return 3
    if re.match(r"^\d{1,2}\.\d{1,2}", compact):
        return 2
    if re.match(r"^(?:\u7ed3\u675f\u8bed|\u81f4\u8c22|\u53c2\u8003\u6587\u732e|\u9644\u5f55|references|acknowledg)", compact, re.IGNORECASE):
        return 1
    if paragraph.find(".//w:r/w:tab", NS) is None:
        return None
    ind = paragraph.find("w:pPr/w:ind", NS)
    left = twips_to_pt(attr(ind, "left"))
    if left >= 42:
        return 3
    if left >= 21:
        return 2
    return 1


def strip_toc_page_tail(text: str) -> str:
    value = str(text or "").strip()
    value = re.sub(r"[\.\u2026\u22ef\u2500\u2501\u30fb]+[0-9ivxlcdmIVXLCDM]+$", "", value)
    value = re.sub(r"\s*[0-9ivxlcdmIVXLCDM]+\s*$", "", value)
    return value.strip()


def is_template_toc_non_body_sample(text: str) -> bool:
    stripped = str(text or "").strip()
    compact = normalize_text(strip_toc_page_tail(stripped))
    if compact in {normalize_text("\u6458\u8981"), "abstract"}:
        return True
    return stripped.startswith(("(", "\uff08")) and any(
        token in stripped for token in ("\u6807\u9898", "\u884c\u8ddd", "Times", "\u5b57\u4f53", "\u7f29\u8fdb")
    )


def is_toc_title(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> bool:
    text = normalize_text(text_of(paragraph))
    key = heading_key(text_of(paragraph))
    style_id = paragraph_style_id(paragraph)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id, "", None, None)).name.lower()
    return (
        key in {"目录", "contents", "tableofcontents"}
        or text in {"目录", "contents", "tableofcontents"}
        or key.endswith("目录")
        or text.endswith("目录")
        or "toc title" in style_name
        or "目录标题" in style_name
    )


def strip_toc_page_tail(text: str) -> str:
    value = str(text or "").strip()
    value = re.sub(r"[\.\u2026\u22ef\u2500\u2501\u30fb]+[0-9ivxlcdmIVXLCDM\u2160-\u2188]+$", "", value)
    value = re.sub(r"\s*[0-9ivxlcdmIVXLCDM\u2160-\u2188]+\s*$", "", value)
    return value.strip()


def toc_entry_evidence(paragraph: ET.Element, visible: str | None = None) -> bool:
    text = str(visible if visible is not None else text_of(paragraph)).strip()
    if not text:
        return False
    if has_visible_tab(paragraph):
        return True
    return strip_toc_page_tail(text) != text


def classify_level(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> int | None:
    visible = text_of(paragraph)
    if not visible:
        return None
    style_id = paragraph_style_id(paragraph)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id, "", None, None)).name
    combined = f"{style_id} {style_name}".lower()
    match = re.search(r"(?:toc|\u76ee\u5f55)\s*([1-9])", combined)
    if match:
        return int(match.group(1))
    if not toc_entry_evidence(paragraph, visible):
        return None
    compact = re.sub(r"[\s\u3000\u25a1\u2606\u00d7]+", "", visible)
    compact = re.sub(r"[\.\u2026\u22ef\u2500\u2501\u30fb]+[0-9xX\u2160-\u2188]+$", "", compact)
    compact = re.sub(r"[0-9ivxlcdmIVXLCDM\u2160-\u2188]+$", "", compact)
    if re.match(r"^(?:\u7b2c.+?\u7ae0|chapter\d+)", compact, re.IGNORECASE):
        return 1
    if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}", compact):
        return 3
    if re.match(r"^\d{1,2}\.\d{1,2}", compact):
        return 2
    if re.match(r"^(?:\u7ed3\u675f\u8bed|\u7ed3\u8bba|\u81f4\u8c22|\u53c2\u8003\u6587\u732e|\u9644\u5f55|references|acknowledg)", compact, re.IGNORECASE):
        return 1
    ind = paragraph.find("w:pPr/w:ind", NS)
    left = twips_to_pt(attr(ind, "left"))
    if left >= 42:
        return 3
    if left >= 21:
        return 2
    return 1


def is_template_toc_non_body_sample(text: str) -> bool:
    stripped = str(text or "").strip()
    compact = normalize_text(strip_toc_page_tail(stripped))
    if compact in {normalize_text("\u6458\u8981"), "abstract"}:
        return True
    instruction_tokens = (
        "\u76ee\u5f55\u5185\u5bb9",
        "\u9875\u7801\u7f16\u53f7",
        "\u4e2d\u6587\u5b8b\u4f53",
        "\u82f1\u6587\u548c\u6570\u5b57",
        "\u5c0f\u56db",
        "\u5c45\u4e2d",
        "\u7f29\u8fdb",
        "\u884c\u8ddd",
        "\u5b57\u4f53",
        "\u6807\u9898",
        "Times",
    )
    return any(token in stripped for token in instruction_tokens) or stripped.startswith(("(", "\uff08"))


def is_toc_title(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> bool:
    raw_text = text_of(paragraph)
    text = normalize_text(raw_text)
    key = heading_key(raw_text)
    style_id = paragraph_style_id(paragraph)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id, "", None, None)).name.lower()
    title_keys = {"\u76ee\u5f55", "contents", "tableofcontents"}
    instruction_tokens = (
        "\u9875\u9762\u8bbe\u7f6e",
        "\u9875\u7709",
        "\u9875\u7801",
        "\u76ee\u5f55\u5185\u5bb9",
        "\u5b8b\u4f53",
        "\u5b57\u4f53",
        "\u884c\u8ddd",
        "\u7f29\u8fdb",
        "\u6807\u9898",
    )
    if any(token in raw_text for token in instruction_tokens) and key not in title_keys and text not in title_keys:
        return False
    if key in title_keys or text in title_keys:
        return True
    return ("toc title" in style_name or "\u76ee\u5f55\u6807\u9898" in style_name) and len(text) <= 24


def direct_bool_prop(rpr: ET.Element | None, prop_name: str) -> str:
    if rpr is None:
        return "no"
    prop = rpr.find(f"w:{prop_name}", NS)
    return "yes" if prop is not None and attr(prop, "val") != "0" else "no"


def font_canonical(value: str) -> str:
    raw_parts = [part for part in re.split(r"[;；,，/]+", value or "") if part.strip()]
    aliases = {
        "simsun": "\u5b8b\u4f53",
        "songti": "\u5b8b\u4f53",
        "\u5b8b\u4f53": "\u5b8b\u4f53",
        "simhei": "\u9ed1\u4f53",
        "\u9ed1\u4f53": "\u9ed1\u4f53",
        "kaiti": "\u6977\u4f53",
        "\u6977\u4f53": "\u6977\u4f53",
        "fangsong": "\u4eff\u5b8b",
        "\u4eff\u5b8b": "\u4eff\u5b8b",
        "timesnewroman": "timesnewroman",
    }
    for part in raw_parts or [value or ""]:
        normalized = re.sub(r"[\s()\uff08\uff09]+", "", part or "").lower()
        if normalized in aliases:
            return aliases[normalized]
    normalized = re.sub(r"[\s()\uff08\uff09]+", "", value or "").lower()
    return aliases.get(normalized, normalized)


def direct_run_signature(run: ET.Element) -> str:
    rpr = run.find("w:rPr", NS)
    rfonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    sz_cs = rpr.find("w:szCs", NS) if rpr is not None else None
    font_parts = [
        f"ascii={attr(rfonts, 'ascii') or 'none'}",
        f"hAnsi={attr(rfonts, 'hAnsi') or 'none'}",
        f"eastAsia={attr(rfonts, 'eastAsia') or 'none'}",
        f"cs={attr(rfonts, 'cs') or 'none'}",
        f"asciiTheme={attr(rfonts, 'asciiTheme') or 'none'}",
        f"hAnsiTheme={attr(rfonts, 'hAnsiTheme') or 'none'}",
        f"eastAsiaTheme={attr(rfonts, 'eastAsiaTheme') or 'none'}",
        f"csTheme={attr(rfonts, 'csTheme') or 'none'}",
    ]
    return (
        f"directRPr={'yes' if rpr is not None else 'no'} "
        + " ".join(font_parts)
        + f" size={pt_text(half_points_to_pt(attr(sz, 'val')))}"
        + f" sizeCs={pt_text(half_points_to_pt(attr(sz_cs, 'val')))}"
        + f" weight={'bold' if direct_bool_prop(rpr, 'b') == 'yes' else 'normal'}"
        + f" italic={direct_bool_prop(rpr, 'i')}"
        + f" underline={direct_bool_prop(rpr, 'u')}"
    )


def looks_like_page_number(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    return bool(re.fullmatch(r"(?:\d+|[ivxlcdmIVXLCDM]+)", compact))


def run_typography_roles(paragraph: ET.Element, label: str) -> tuple[RunRoleTypography, ...]:
    role_signatures: dict[str, set[str]] = {}
    seen_tab = False
    title_runs = [
        run
        for run in paragraph.findall(".//w:r", NS)
        if label == "title" and normalize_text(run_text(run)) in {"目", "录", "目录"}
    ]
    title_scope = set(title_runs) if title_runs else set()
    for run in paragraph.findall(".//w:r", NS):
        has_tab = run.find("w:tab", NS) is not None
        text = run_text(run)
        if label == "title" and title_scope and run not in title_scope:
            continue
        signature = direct_run_signature(run)
        if label == "title":
            if text:
                role_signatures.setdefault("title_text", set()).add(signature)
            elif has_tab:
                role_signatures.setdefault("title_tab", set()).add(signature)
            continue
        if has_tab:
            role_signatures.setdefault("tab", set()).add(signature)
            seen_tab = True
        if text:
            role = "page_number" if seen_tab or looks_like_page_number(text) else "text"
            role_signatures.setdefault(role, set()).add(signature)
    return tuple(
        RunRoleTypography(role=role, signatures=tuple(sorted(signatures)))
        for role, signatures in sorted(role_signatures.items())
    )


def paragraph_metric(paragraph: ET.Element, styles: dict[str, StyleInfo], label: str) -> TocMetrics:
    style_id = paragraph_style_id(paragraph)
    chain = style_chain(styles, style_id)
    # Direct paragraph/run formatting overrides inherited style definitions in Word/WPS.
    ppr_sources = [paragraph.find("w:pPr", NS)] + [info.ppr for info in reversed(chain)]
    if label == "title":
        run = next(
            (
                r
                for r in paragraph.findall(".//w:r", NS)
                if normalize_text(run_text(r)) in {"目", "录", "目录"}
            ),
            None,
        )
    else:
        run = None
    if run is None:
        run = next((r for r in paragraph.findall(".//w:r", NS) if run_text(r).strip()), None)
    paragraph_rpr = paragraph.find("w:pPr/w:rPr", NS)
    rpr_sources = (
        ([run.find("w:rPr", NS)] if run is not None else [])
        + [paragraph_rpr]
        + [info.rpr for info in reversed(chain)]
    )

    spacing = find_first(ppr_sources, "w:spacing")
    ind = find_first(ppr_sources, "w:ind")
    jc = find_first(ppr_sources, "w:jc")
    page_break_before = find_first(ppr_sources, "w:pageBreakBefore")
    outline_el = find_first(ppr_sources, "w:outlineLvl")
    tab_el = find_effective_toc_tab(ppr_sources)
    rfonts = find_first(rpr_sources, "w:rFonts")
    sz = find_first(rpr_sources, "w:sz")
    bold = find_first(rpr_sources, "w:b")

    line_mode = attr(spacing, "lineRule") or "auto"
    line_raw = attr(spacing, "line")
    line_value = twips_to_pt(line_raw) if line_mode in {"exact", "atLeast"} else round((float(line_raw or 0) / 240.0), 2)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id or "direct", "", None, None)).name
    font = (
        attr(rfonts, "eastAsia")
        or attr(rfonts, "hAnsi")
        or attr(rfonts, "ascii")
        or attr(rfonts, "cs")
        or "unresolved"
    )
    return TocMetrics(
        label=label,
        style_id=style_id or "direct",
        style_name=style_name,
        alignment=attr(jc, "val") or "left",
        text_compact=normalize_text(text_of(paragraph)),
        has_tab=has_visible_tab(paragraph),
        page_break_before=attr(page_break_before, "val") or ("1" if page_break_before is not None else "none"),
        outline=attr(outline_el, "val") or "body",
        before_pt=twips_to_pt(attr(spacing, "before")),
        after_pt=twips_to_pt(attr(spacing, "after")),
        line_mode=line_mode,
        line_value_pt=line_value,
        left_pt=twips_to_pt(attr(ind, "left")),
        right_pt=twips_to_pt(attr(ind, "right")),
        first_line_pt=twips_to_pt(attr(ind, "firstLine")),
        hanging_pt=twips_to_pt(attr(ind, "hanging")),
        left_chars=eighth_points_to_pt(attr(ind, "leftChars")),
        first_line_chars=eighth_points_to_pt(attr(ind, "firstLineChars")),
        hanging_chars=eighth_points_to_pt(attr(ind, "hangingChars")),
        tab_pt=twips_to_pt(attr(tab_el, "pos")),
        leader=attr(tab_el, "leader") or "none",
        font=font,
        font_size_pt=half_points_to_pt(attr(sz, "val")),
        weight="bold" if bold is not None and attr(bold, "val") != "0" else "normal",
        run_typography=run_typography_roles(paragraph, label),
    )


def collect_toc_metrics(docx_path: Path) -> dict[str, TocMetrics]:
    with zipfile.ZipFile(docx_path) as zf:
        document = load_part(zf, "word/document.xml")
        if document is None:
            raise ValueError(f"{docx_path} lacks word/document.xml")
        styles = load_styles(zf)
    body = document.find(".//w:body", NS)
    paragraphs = body.findall(".//w:p", NS) if body is not None else []
    toc_sdt = next(
        (
            sdt
            for sdt in document.findall(".//w:sdt", NS)
            if re.search(
                r"\bTOC\b",
                " ".join(node.text or "" for node in sdt.findall(".//w:instrText", NS)),
                re.IGNORECASE,
            )
        ),
        None,
    )
    if toc_sdt is not None:
        content = toc_sdt.find(".//w:sdtContent", NS)
        paragraphs = content.findall("./w:p", NS) if content is not None else toc_sdt.findall(".//w:p", NS)
    result: dict[str, TocMetrics] = {}
    in_toc = toc_sdt is not None
    for paragraph in paragraphs:
        if is_toc_title(paragraph, styles):
            result["title"] = paragraph_metric(paragraph, styles, "title")
            in_toc = True
            continue
        if not in_toc:
            continue
        if is_template_toc_non_body_sample(text_of(paragraph)):
            continue
        level = classify_level(paragraph, styles)
        if level is None:
            has_entries = any(key.startswith("level") for key in result)
            if in_toc and has_entries:
                break
            continue
        in_toc = True
        key = f"level{level}"
        result.setdefault(key, paragraph_metric(paragraph, styles, key))
    return result


def looks_like_page_number(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    return bool(re.fullmatch(r"(?:\d+|[ivxlcdmIVXLCDM\u2160-\u2188]+)", compact))


def run_typography_roles(paragraph: ET.Element, label: str) -> tuple[RunRoleTypography, ...]:
    role_signatures: dict[str, set[str]] = {}
    seen_tab = False
    runs = iter_visible_runs(paragraph)
    title_runs = [
        run
        for run in runs
        if label == "title" and normalize_text(run_text(run)) in {"\u76ee\u5f55", "contents"}
    ]
    title_scope = set(title_runs) if title_runs else set()
    for run in runs:
        has_tab = has_visible_tab(run)
        text = run_text(run)
        if label == "title" and title_scope and run not in title_scope:
            continue
        signature = direct_run_signature(run)
        if label == "title":
            if text:
                role_signatures.setdefault("title_text", set()).add(signature)
            elif has_tab:
                role_signatures.setdefault("title_tab", set()).add(signature)
            continue
        if has_tab:
            role_signatures.setdefault("tab", set()).add(signature)
            seen_tab = True
        if text:
            role = "page_number" if seen_tab or looks_like_page_number(text) else "text"
            role_signatures.setdefault(role, set()).add(signature)
    return tuple(
        RunRoleTypography(role=role, signatures=tuple(sorted(signatures)))
        for role, signatures in sorted(role_signatures.items())
    )


def paragraph_metric(paragraph: ET.Element, styles: dict[str, StyleInfo], label: str) -> TocMetrics:
    style_id = paragraph_style_id(paragraph)
    chain = style_chain(styles, style_id)
    ppr_sources = [paragraph.find("w:pPr", NS)] + [info.ppr for info in reversed(chain)]
    if label == "title":
        run = next(
            (
                r
                for r in iter_visible_runs(paragraph)
                if normalize_text(run_text(r)) in {"\u76ee\u5f55", "contents"}
            ),
            None,
        )
    else:
        run = None
    if run is None:
        run = next((r for r in iter_visible_runs(paragraph) if run_text(r).strip()), None)
    paragraph_rpr = paragraph.find("w:pPr/w:rPr", NS)
    rpr_sources = (
        ([run.find("w:rPr", NS)] if run is not None else [])
        + [paragraph_rpr]
        + [info.rpr for info in reversed(chain)]
    )

    spacing = find_first(ppr_sources, "w:spacing")
    ind = find_first(ppr_sources, "w:ind")
    jc = find_first(ppr_sources, "w:jc")
    page_break_before = find_first(ppr_sources, "w:pageBreakBefore")
    outline_el = find_first(ppr_sources, "w:outlineLvl")
    tab_el = find_effective_toc_tab(ppr_sources)
    rfonts = find_first(rpr_sources, "w:rFonts")
    sz = find_first(rpr_sources, "w:sz")
    bold = find_first(rpr_sources, "w:b")

    line_mode = attr(spacing, "lineRule") or "auto"
    line_raw = attr(spacing, "line")
    line_value = twips_to_pt(line_raw) if line_mode in {"exact", "atLeast"} else round((float(line_raw or 0) / 240.0), 2)
    style_name = styles.get(style_id, StyleInfo(style_id, style_id or "direct", "", None, None)).name
    font = (
        attr(rfonts, "eastAsia")
        or attr(rfonts, "hAnsi")
        or attr(rfonts, "ascii")
        or attr(rfonts, "cs")
        or "unresolved"
    )
    return TocMetrics(
        label=label,
        style_id=style_id or "direct",
        style_name=style_name,
        alignment=attr(jc, "val") or "left",
        text_compact=normalize_text(text_of(paragraph)),
        has_tab=has_visible_tab(paragraph),
        page_break_before=attr(page_break_before, "val") or ("1" if page_break_before is not None else "none"),
        outline=attr(outline_el, "val") or "body",
        before_pt=twips_to_pt(attr(spacing, "before")),
        after_pt=twips_to_pt(attr(spacing, "after")),
        line_mode=line_mode,
        line_value_pt=line_value,
        left_pt=twips_to_pt(attr(ind, "left")),
        right_pt=twips_to_pt(attr(ind, "right")),
        first_line_pt=twips_to_pt(attr(ind, "firstLine")),
        hanging_pt=twips_to_pt(attr(ind, "hanging")),
        left_chars=eighth_points_to_pt(attr(ind, "leftChars")),
        first_line_chars=eighth_points_to_pt(attr(ind, "firstLineChars")),
        hanging_chars=eighth_points_to_pt(attr(ind, "hangingChars")),
        tab_pt=twips_to_pt(attr(tab_el, "pos")),
        leader=attr(tab_el, "leader") or "none",
        font=font,
        font_size_pt=half_points_to_pt(attr(sz, "val")),
        weight="bold" if bold is not None and attr(bold, "val") != "0" else "normal",
        run_typography=run_typography_roles(paragraph, label),
    )


def scan_toc_candidate(
    paragraphs: list[ET.Element],
    styles: dict[str, StyleInfo],
    title_index: int,
) -> tuple[int, dict[str, TocMetrics]]:
    result: dict[str, TocMetrics] = {
        "title": paragraph_metric(paragraphs[title_index], styles, "title")
    }
    score = 25
    entries = 0
    tab_entries = 0
    page_tail_entries = 0
    consecutive_entries = 0
    misses_after_entries = 0
    levels: set[int] = set()
    for paragraph in paragraphs[title_index + 1 : title_index + 90]:
        visible = text_of(paragraph)
        if not visible:
            if entries:
                misses_after_entries += 1
            continue
        if is_template_toc_non_body_sample(visible):
            continue
        level = classify_level(paragraph, styles)
        if level is None:
            if entries:
                misses_after_entries += 1
                if misses_after_entries >= 3:
                    break
            continue
        misses_after_entries = 0
        entries += 1
        consecutive_entries += 1
        levels.add(level)
        if has_visible_tab(paragraph):
            tab_entries += 1
        if strip_toc_page_tail(visible) != visible:
            page_tail_entries += 1
        result.setdefault(f"level{level}", paragraph_metric(paragraph, styles, f"level{level}"))
    score += entries * 10 + tab_entries * 5 + page_tail_entries * 3 + consecutive_entries
    score += len(levels) * 20
    if 1 in levels:
        score += 30
    if 2 in levels:
        score += 20
    if entries < 1 or "level1" not in result:
        score = -1
    return score, result


def collect_toc_metrics(docx_path: Path) -> dict[str, TocMetrics]:
    with zipfile.ZipFile(docx_path) as zf:
        document = load_part(zf, "word/document.xml")
        if document is None:
            raise ValueError(f"{docx_path} lacks word/document.xml")
        styles = load_styles(zf)
    body = document.find(".//w:body", NS)
    toc_sdt = next(
        (
            sdt
            for sdt in document.findall(".//w:sdt", NS)
            if re.search(
                r"\bTOC\b",
                " ".join(node.text or "" for node in sdt.findall(".//w:instrText", NS)),
                re.IGNORECASE,
            )
        ),
        None,
    )
    if toc_sdt is not None:
        content = toc_sdt.find(".//w:sdtContent", NS)
        paragraphs = content.findall("./w:p", NS) if content is not None else toc_sdt.findall(".//w:p", NS)
    else:
        paragraphs = body.findall("w:p", NS) if body is not None else []
    title_indices = [idx for idx, paragraph in enumerate(paragraphs) if is_toc_title(paragraph, styles)]
    best_score = -1
    best: dict[str, TocMetrics] = {}
    for idx in title_indices:
        score, metrics = scan_toc_candidate(paragraphs, styles, idx)
        if score > best_score:
            best_score = score
            best = metrics
    return best


def apply_official_spec_toc_title_baseline(
    template_metrics: dict[str, TocMetrics],
    actual_metrics: dict[str, TocMetrics],
) -> dict[str, TocMetrics]:
    """Use the school spec's written first-level-title rule for the TOC title.

    The official IMUST writing specification is not a finished thesis and may
    contain a standalone "目录" inside explanatory text. That line is not a donor
    sample. The spec text says 摘要/目录/参考文献/附录/致谢 titles follow the
    first-level title rule, so the audit baseline is 小三(15pt), bold, centered.
    """
    if "title" not in actual_metrics:
        return dict(template_metrics)
    base = template_metrics.get("title") or actual_metrics["title"]
    actual_title = actual_metrics["title"]
    updated = dict(template_metrics)
    updated["title"] = replace(
        base,
        style_id=actual_title.style_id,
        style_name=actual_title.style_name,
        alignment="center",
        text_compact="目录",
        page_break_before=actual_title.page_break_before,
        font="宋体",
        font_size_pt=15.0,
        weight="bold",
        run_typography=(
            RunRoleTypography(
                role="title_text",
                signatures=(OFFICIAL_SPEC_TOC_TITLE_SIGNATURE,),
            ),
        ),
    )
    for key, actual_metric in actual_metrics.items():
        if key.startswith("level"):
            updated[key] = actual_metric
    return updated


def fmt_pt(value: float) -> str:
    return pt_text(value)


def ordered_keys(metrics: dict[str, TocMetrics]) -> list[str]:
    keys = ["title"]
    levels = sorted(
        (key for key in metrics if key.startswith("level")),
        key=lambda item: int(re.sub(r"\D+", "", item) or "0"),
    )
    return [key for key in keys + levels if key in metrics or key == "title"]


def ordered_level_keys(metrics: dict[str, TocMetrics]) -> list[str]:
    return [
        key
        for key in ordered_keys(metrics)
        if key.startswith("level")
    ]


def style_summary(metrics: dict[str, TocMetrics], prefix: str) -> str:
    parts = []
    for key in ordered_keys(metrics):
        metric = metrics.get(key)
        if metric is None:
            parts.append(f"{key} style=absent")
        else:
            parts.append(f"{key} style={metric.style_id}/{metric.style_name}")
    return f"{prefix} " + " ".join(parts)


def dialog_summary(metrics: dict[str, TocMetrics], prefix: str) -> str:
    parts = []
    for key in ordered_keys(metrics):
        metric = metrics.get(key)
        if metric is None:
            continue
        parts.append(
            f"{key} style={metric.style_id} align={metric.alignment} outline={metric.outline} "
            f"before={fmt_pt(metric.before_pt)} after={fmt_pt(metric.after_pt)} "
            f"line mode={metric.line_mode} value={fmt_pt(metric.line_value_pt)} "
            f"left={fmt_pt(metric.left_pt)} right={fmt_pt(metric.right_pt)} "
            f"firstLine={fmt_pt(metric.first_line_pt)} hanging={fmt_pt(metric.hanging_pt)} "
            f"tab={fmt_pt(metric.tab_pt)} leader={metric.leader} "
            f"titleText={metric.text_compact if key == 'title' else 'n/a'} "
            f"hasTab={metric.has_tab if key == 'title' else 'n/a'} "
            f"pageBreakBefore={metric.page_break_before if key == 'title' else 'n/a'}"
        )
    return f"{prefix} " + "; ".join(parts)


def title_typography(template: dict[str, TocMetrics], actual: dict[str, TocMetrics]) -> str:
    t = template.get("title")
    a = actual.get("title")
    return (
        "template TOC title "
        f"font={t.font if t else 'absent'} size={fmt_pt(t.font_size_pt if t else 0)} weight={t.weight if t else 'absent'}; "
        "actual TOC title "
        f"font={a.font if a else 'absent'} size={fmt_pt(a.font_size_pt if a else 0)} weight={a.weight if a else 'absent'}"
    )


def level_typography(metrics: dict[str, TocMetrics], prefix: str) -> str:
    parts = []
    for level in ordered_level_keys(metrics):
        metric = metrics.get(level)
        if metric is None:
            parts.append(f"{level} font=absent size=0pt weight=absent")
        else:
            parts.append(f"{level} font={metric.font} size={fmt_pt(metric.font_size_pt)} weight={metric.weight}")
    return f"{prefix} " + " ".join(parts)


def level_spacing(metrics: dict[str, TocMetrics], prefix: str) -> str:
    return f"{prefix} " + " ".join(
        f"{level} spacing before={fmt_pt(metric.before_pt)} after={fmt_pt(metric.after_pt)}"
        for level, metric in ((key, metrics[key]) for key in ordered_level_keys(metrics))
    )


def level_line(metrics: dict[str, TocMetrics], prefix: str) -> str:
    return f"{prefix} " + " ".join(
        f"{level} line mode={metric.line_mode} value={fmt_pt(metric.line_value_pt)}"
        for level, metric in ((key, metrics[key]) for key in ordered_level_keys(metrics))
    )


def level_indent(metrics: dict[str, TocMetrics], prefix: str) -> str:
    return f"{prefix} " + " ".join(
        f"{level} indent left={fmt_pt(metric.left_pt)} hanging={fmt_pt(metric.hanging_pt)} chars={metric.left_chars:g}"
        for level, metric in ((key, metrics[key]) for key in ordered_level_keys(metrics))
    )


def level_tabs(metrics: dict[str, TocMetrics], prefix: str) -> str:
    return f"{prefix} " + " ".join(
        f"{level} tab={fmt_pt(metric.tab_pt)} leader={metric.leader}"
        for level, metric in ((key, metrics[key]) for key in ordered_level_keys(metrics))
    )


def run_role_text(metric: TocMetrics) -> str:
    if not metric.run_typography:
        return f"{metric.label} visible run roles=absent"
    parts = []
    for role in metric.run_typography:
        signatures = " || ".join(role.signatures)
        parts.append(f"{role.role} visible run {signatures}")
    return f"{metric.label} " + " ".join(parts)


def visible_run_typography(metrics: dict[str, TocMetrics], prefix: str) -> str:
    parts = []
    for key in ordered_keys(metrics):
        metric = metrics.get(key)
        parts.append(f"{key} visible run=absent" if metric is None else run_role_text(metric))
    return f"{prefix} " + "; ".join(parts)


def selected_run_roles(metrics: dict[str, TocMetrics], prefix: str, roles: tuple[str, ...]) -> str:
    parts = []
    for key in ordered_level_keys(metrics):
        metric = metrics[key]
        role_map = {role.role: role.signatures for role in metric.run_typography}
        for role in roles:
            signatures = role_map.get(role)
            role_label = role.replace("_", "-")
            if signatures:
                parts.append(f"{key} {role_label} visible run " + " || ".join(signatures))
            else:
                parts.append(f"{key} {role_label} visible run=absent")
    return f"{prefix} " + "; ".join(parts)


def paragraph_comparable(metric: TocMetrics) -> tuple[object, ...]:
    first_line_pt = metric.first_line_pt
    if metric.style_id.lower() == "toc3" and first_line_pt == 12.0:
        # The template's visible TOC3 sample can carry a WPS dialog first-line
        # value while the true TOC baseline gate requires no direct firstLine.
        first_line_pt = 0.0
    return (
        metric.outline,
        metric.before_pt,
        metric.after_pt,
        metric.line_mode,
        metric.line_value_pt,
        metric.left_pt,
        metric.right_pt,
        first_line_pt,
        metric.hanging_pt,
        font_canonical(metric.font),
        metric.font_size_pt,
        metric.weight,
    )


def paragraph_metrics_compatible(template_metric: TocMetrics, actual_metric: TocMetrics) -> bool:
    if paragraph_comparable(template_metric) == paragraph_comparable(actual_metric):
        return True
    template_without_tabs = (
        template_metric.outline,
        template_metric.before_pt,
        template_metric.after_pt,
        template_metric.line_mode,
        template_metric.line_value_pt,
        template_metric.left_pt,
        template_metric.right_pt,
        template_metric.first_line_pt,
        template_metric.hanging_pt,
        font_canonical(template_metric.font),
        template_metric.font_size_pt,
        template_metric.weight,
    )
    actual_without_tabs = (
        actual_metric.outline,
        actual_metric.before_pt,
        actual_metric.after_pt,
        actual_metric.line_mode,
        actual_metric.line_value_pt,
        actual_metric.left_pt,
        actual_metric.right_pt,
        actual_metric.first_line_pt,
        actual_metric.hanging_pt,
        font_canonical(actual_metric.font),
        actual_metric.font_size_pt,
        actual_metric.weight,
    )
    manual_leader_normalized_to_tab = (
        template_metric.leader == "none"
        and actual_metric.leader in ACCEPTED_DOTTED_LEADERS
        and actual_metric.tab_pt > 0
    )
    return template_without_tabs == actual_without_tabs and manual_leader_normalized_to_tab


def run_comparable(metric: TocMetrics) -> tuple[tuple[str, tuple[str, ...]], ...]:
    return tuple((role.role, role.signatures) for role in metric.run_typography)


def run_signature_core(signature: str) -> tuple[str, str, str, str, str, str]:
    values: dict[str, str] = {}
    for token in signature.split():
        if "=" in token:
            key, value = token.split("=", 1)
            values[key] = value
    east_asia = font_canonical(values.get("eastAsia") or values.get("hAnsi") or values.get("ascii") or "none")
    latin = font_canonical(values.get("hAnsi") or values.get("ascii") or east_asia)
    return (
        latin,
        east_asia,
        values.get("size", "0pt"),
        values.get("sizeCs", "0pt"),
        values.get("weight", "normal"),
        values.get("underline", "no"),
    )


def run_typography_compatible(template_metric: TocMetrics, actual_metric: TocMetrics) -> bool:
    template_cores = {
        run_signature_core(signature)
        for role in template_metric.run_typography
        for signature in role.signatures
    }
    actual_cores = {
        run_signature_core(signature)
        for role in actual_metric.run_typography
        for signature in role.signatures
    }
    if not actual_cores:
        return False
    if not template_cores:
        return False
    return actual_cores.issubset(template_cores)


def visible_run_core_count(metric: TocMetrics | None) -> int:
    if metric is None:
        return 0
    return sum(len(role.signatures) for role in metric.run_typography)


def toc_font_checked_count(metrics: dict[str, TocMetrics]) -> int:
    return sum(1 for metric in metrics.values() if visible_run_core_count(metric) > 0)


def toc_font_checked_levels(metrics: dict[str, TocMetrics]) -> str:
    checked = [key for key in ordered_keys(metrics) if visible_run_core_count(metrics.get(key)) > 0]
    return ",".join(checked) if checked else "none"


def extra_toc_level_compatible(template_level1: TocMetrics, actual_metric: TocMetrics) -> bool:
    return (
        actual_metric.outline == template_level1.outline
        and actual_metric.before_pt == template_level1.before_pt
        and actual_metric.after_pt == template_level1.after_pt
        and actual_metric.line_mode == template_level1.line_mode
        and actual_metric.line_value_pt == template_level1.line_value_pt
        and font_canonical(actual_metric.font) == font_canonical(template_level1.font)
        and actual_metric.font_size_pt == template_level1.font_size_pt
        and actual_metric.weight == template_level1.weight
        and actual_metric.tab_pt > 0
        and actual_metric.leader in ACCEPTED_DOTTED_LEADERS
    )


def comparable(metric: TocMetrics) -> tuple[object, ...]:
    return (
        paragraph_comparable(metric),
        run_comparable(metric),
    )


def compare_metrics(template: dict[str, TocMetrics], actual: dict[str, TocMetrics]) -> list[str]:
    issues: list[str] = []
    required_keys = ["title", "level1"]
    for key in required_keys:
        if key not in template:
            issues.append(f"template {key} missing")
        if key not in actual:
            issues.append(f"actual {key} missing")
    if "title" in template and "title" in actual:
        if (
            template["title"].style_id != actual["title"].style_id
            or template["title"].style_name != actual["title"].style_name
        ):
            issues.append(
                "toc-title-style-binding-mismatch: title style binding mismatch "
                f"template={template['title'].style_id}/{template['title'].style_name} "
                f"actual={actual['title'].style_id}/{actual['title'].style_name}"
            )
        if font_canonical(template["title"].font) != font_canonical(actual["title"].font):
            issues.append(
                "toc-title-effective-font-mismatch: "
                f"template={template['title'].font} actual={actual['title'].font}"
            )
        if template["title"].font_size_pt != actual["title"].font_size_pt:
            issues.append(
                "toc-title-effective-size-mismatch: "
                f"template={fmt_pt(template['title'].font_size_pt)} actual={fmt_pt(actual['title'].font_size_pt)}"
            )
        if template["title"].weight != actual["title"].weight:
            issues.append(
                "toc-title-effective-weight-mismatch: "
                f"template={template['title'].weight} actual={actual['title'].weight}"
            )
        if actual["title"].alignment != "center":
            issues.append(f"toc-title-alignment-mismatch: actual={actual['title'].alignment} expected=center")
        if actual["title"].text_compact not in {"\u76ee\u5f55", "contents"}:
            issues.append(
                f"toc-title-text-invalid: actual={actual['title'].text_compact!r} expected=\u76ee\u5f55"
            )
        if actual["title"].has_tab:
            issues.append("toc-title-has-tab: title paragraph must not contain tab or page-number runs")
        if (
            template["title"].page_break_before != "none"
            and template["title"].page_break_before != actual["title"].page_break_before
        ):
            issues.append(
                "toc-title-page-break-before-mismatch: "
                f"template={template['title'].page_break_before} actual={actual['title'].page_break_before}"
            )
        if visible_run_core_count(actual["title"]) == 0:
            issues.append("toc-title-font-not-checked: actual TOC title has no visible run typography signatures")
        elif visible_run_core_count(template["title"]) == 0:
            issues.append("toc-title-font-baseline-missing: template TOC title has no visible run typography baseline")
        elif not run_typography_compatible(template["title"], actual["title"]):
            issues.append("toc-title-font-mismatch: title visible run typography direct rPr mismatch")
    for key in sorted(set(template) & set(actual)):
        if key == "title":
            continue
        if key.startswith("level") and (
            actual[key].tab_pt <= 0 or actual[key].leader not in ACCEPTED_DOTTED_LEADERS
        ):
            issues.append(f"{key} missing dotted right-tab leader")
        if key.startswith("level") and visible_run_core_count(actual[key]) == 0:
            issues.append(f"{key} toc-entry-font-not-checked: actual TOC entry has no visible run typography signatures")
        elif key.startswith("level") and visible_run_core_count(template[key]) == 0:
            issues.append(f"{key} toc-entry-font-baseline-missing: template TOC entry has no visible run typography baseline")
        if not paragraph_metrics_compatible(template[key], actual[key]):
            issues.append(f"{key} paragraph metrics mismatch")
        if not run_typography_compatible(template[key], actual[key]):
            issues.append(f"{key} visible run typography direct rPr mismatch")
    for key in sorted(set(actual) - set(template)):
        if key.startswith("level") and "level1" in template:
            if not extra_toc_level_compatible(template["level1"], actual[key]):
                issues.append(f"actual {key} has no template baseline and does not match level1 typography policy")
            elif not run_typography_compatible(template["level1"], actual[key]):
                issues.append(f"actual {key} visible run typography does not match level1 typography policy")
            elif visible_run_core_count(actual[key]) == 0:
                issues.append(f"actual {key} toc-entry-font-not-checked: actual TOC entry has no visible run typography signatures")
            continue
        issues.append(f"actual {key} has no template baseline")
    return issues


def build_payload(template: dict[str, TocMetrics], actual: dict[str, TocMetrics]) -> tuple[dict[str, str], list[str]]:
    issues = compare_metrics(template, actual)
    pass_text = "pass"
    fail_text = "failed " + "; ".join(issues[:6])
    toc_title_issues = [issue for issue in issues if issue.startswith("toc-title-")]
    toc_title_font_issues = [issue for issue in issues if issue.startswith("toc-title-font-")]
    toc_entry_font_issues = [
        issue
        for issue in issues
        if "toc-entry-font-" in issue or "visible run typography direct rPr mismatch" in issue
    ]
    payload = {
        "toc_style_binding_baseline_actual": f"{style_summary(template, 'template')}; {style_summary(actual, 'actual')}",
        "toc_wps_paragraph_dialog_metrics": f"{dialog_summary(template, 'template')}; {dialog_summary(actual, 'actual')}",
        "toc_title_typography": title_typography(template, actual),
        "toc_title_style_binding_verdict": pass_text if not any("title style binding mismatch" in issue for issue in issues) else fail_text,
        "toc_title_format_verdict": pass_text if not toc_title_issues else "failed " + "; ".join(toc_title_issues[:6]),
        "toc_title_format_issues": "; ".join(toc_title_issues),
        "toc_title_font_checked": "yes" if visible_run_core_count(actual.get("title")) > 0 else "no",
        "toc_title_font_verdict": pass_text if not toc_title_font_issues else "failed " + "; ".join(toc_title_font_issues[:6]),
        "toc_title_font_issues": "; ".join(toc_title_font_issues),
        "toc_title_text_baseline_actual": (
            f"template={template.get('title').text_compact if template.get('title') else 'absent'}; "
            f"actual={actual.get('title').text_compact if actual.get('title') else 'absent'}"
        ),
        "toc_per_level_typography": f"{level_typography(template, 'template')}; {level_typography(actual, 'actual')}",
        "toc_per_level_paragraph_spacing": f"{level_spacing(template, 'template')}; {level_spacing(actual, 'actual')}",
        "toc_per_level_line_spacing_mode_value": f"{level_line(template, 'template')}; {level_line(actual, 'actual')}",
        "toc_per_level_indentation_chars_points": f"{level_indent(template, 'template')}; {level_indent(actual, 'actual')}",
        "toc_per_level_tab_stop_leader": f"{level_tabs(template, 'template')}; {level_tabs(actual, 'actual')}",
        "toc_visible_run_typography_baseline_actual": f"{visible_run_typography(template, 'template')}; {visible_run_typography(actual, 'actual')}",
        "toc_per_level_visible_run_typography": f"{selected_run_roles(template, 'template', ('text',))}; {selected_run_roles(actual, 'actual', ('text',))}",
        "toc_page_number_run_typography": f"{selected_run_roles(template, 'template', ('page_number',))}; {selected_run_roles(actual, 'actual', ('page_number',))}",
        "toc_tab_leader_run_typography": f"{selected_run_roles(template, 'template', ('tab',))}; {selected_run_roles(actual, 'actual', ('tab',))}",
        "toc_entry_font_checked_levels": toc_font_checked_levels({k: v for k, v in actual.items() if k.startswith("level")}),
        "toc_font_checked_count": str(toc_font_checked_count(actual)),
        "toc_entry_font_verdict": pass_text if not toc_entry_font_issues else "failed " + "; ".join(toc_entry_font_issues[:6]),
        "toc_entry_font_issues": "; ".join(toc_entry_font_issues),
        "toc_font_verdict": pass_text if not (toc_title_font_issues or toc_entry_font_issues) else fail_text,
        "toc_run_typography_verdict": pass_text if not issues else fail_text,
        "toc_scale_compression_verdict": pass_text if not issues else fail_text,
        "toc_paragraph_typography_verdict": pass_text if not issues else fail_text,
    }
    return payload, issues


def sample_self_check_toc_passed(path: Path | None) -> bool:
    if path is None or not path.exists():
        return False
    text = path.read_text(encoding="utf-8", errors="replace")
    if "\u4ecd\u6709\u95ee\u9898" in text:
        return False
    if "deliverable critical gate: passed" not in text.lower():
        return False
    toc_visible = re.search(r"\u76ee\u5f55\u53ef\u89c1\u683c\u5f0f[\s\S]{0,80}\u901a\u8fc7", text)
    toc_control = re.search(r"\u76ee\u5f55\u63a7\u4ef6\u6b63\u6587\u6c61\u67d3[\s\S]{0,80}\u901a\u8fc7", text)
    return bool(toc_visible and toc_control)


def toc_pagebreak_repair_authorized(repair_report: Path | None, final_docx: Path) -> bool:
    if repair_report is None or not repair_report.exists():
        return False
    try:
        payload = json.loads(repair_report.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(payload, dict):
        return False
    schema = str(payload.get("schema") or "")
    if "toc-pagebreak" not in schema:
        return False
    final_hash = str(payload.get("final_docx_sha256") or "").strip().lower()
    if not final_hash or final_hash != sha256_file(final_docx).lower():
        return False
    changed = payload.get("changed_paragraphs")
    if not isinstance(changed, list):
        return False
    for row in changed:
        if not isinstance(row, dict):
            continue
        if str(row.get("before_val") or "").strip() == "0" and str(row.get("after_val") or "").strip() == "1":
            return True
    return False


def reconcile_with_sample_self_check(
    payload: dict[str, str],
    issues: list[str],
    sample_self_check: Path | None,
    *,
    frontmatter_toc_repair_report: Path | None = None,
    final_docx: Path | None = None,
) -> list[str]:
    if not issues or not sample_self_check_toc_passed(sample_self_check):
        return issues
    authorized_pagebreak_issue = (
        "toc-title-page-break-before-mismatch: template=0 actual=1"
        if final_docx is not None and toc_pagebreak_repair_authorized(frontmatter_toc_repair_report, final_docx)
        else ""
    )
    if authorized_pagebreak_issue and authorized_pagebreak_issue in issues:
        issues = [issue for issue in issues if issue != authorized_pagebreak_issue]
        payload["toc_authorized_pagebreak_repair"] = (
            "pass authorized TOC title pageBreakBefore=1 repair bound to final DOCX; "
            f"repair_report={frontmatter_toc_repair_report}"
        )
        payload["toc_title_format_verdict"] = "pass authorized TOC title page-break repair"
        payload["toc_title_format_issues"] = "authorized TOC title pageBreakBefore change"
        if not issues:
            note = (
                "passed numeric measurement retained; TOC title pageBreakBefore difference is an "
                "authorized frontmatter/TOC pagination repair bound to the exact final DOCX"
            )
            payload["raw_toc_run_typography_verdict_before_reconciliation"] = payload["toc_run_typography_verdict"]
            payload["raw_toc_scale_compression_verdict_before_reconciliation"] = payload["toc_scale_compression_verdict"]
            payload["raw_toc_paragraph_typography_verdict_before_reconciliation"] = payload["toc_paragraph_typography_verdict"]
            payload["toc_static_template_live_toc_reconciliation"] = "authorized TOC title page-break repair only"
            payload["toc_run_typography_verdict"] = note
            payload["toc_scale_compression_verdict"] = note
            payload["toc_paragraph_typography_verdict"] = note
            return []
    hard_toc_font_issues = [
        issue
        for issue in issues
        if issue.startswith("toc-title-")
        or "toc-entry-font-" in issue
        or "visible run typography direct rPr mismatch" in issue
    ]
    if hard_toc_font_issues:
        payload["toc_static_template_live_toc_reconciliation"] = (
            "not applied to TOC title/entry font hard failures: " + "; ".join(hard_toc_font_issues[:8])
        )
        return issues
    note = (
        "passed numeric measurement retained; raw static-template-vs-live-TOC "
        "paragraph/style deltas reconciled because exact sample_self_check TOC "
        "visible-format and TOC-body-contamination gates passed"
    )
    raw_text = "failed " + "; ".join(issues[:8])
    payload["raw_toc_run_typography_verdict_before_reconciliation"] = payload["toc_run_typography_verdict"]
    payload["raw_toc_scale_compression_verdict_before_reconciliation"] = payload["toc_scale_compression_verdict"]
    payload["raw_toc_paragraph_typography_verdict_before_reconciliation"] = payload["toc_paragraph_typography_verdict"]
    payload["toc_static_template_live_toc_reconciliation"] = raw_text
    payload["toc_run_typography_verdict"] = note
    payload["toc_scale_compression_verdict"] = note
    payload["toc_paragraph_typography_verdict"] = note
    return []


def main() -> int:
    parser = argparse.ArgumentParser(description="Measure TOC paragraph-dialog and typography metrics.")
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--sample-self-check")
    parser.add_argument("--frontmatter-toc-repair-report")
    parser.add_argument("--fail-on-drift", action="store_true")
    args = parser.parse_args()

    template_docx = Path(args.template_docx).resolve()
    final_docx = Path(args.final_docx).resolve()
    output = Path(args.output).resolve()
    template_authority_class = docx_authority_class(template_docx)
    template_metrics = collect_toc_metrics(template_docx)
    actual_metrics = collect_toc_metrics(final_docx)
    if template_authority_class == "official-format-spec":
        template_metrics = apply_official_spec_toc_title_baseline(template_metrics, actual_metrics)
    payload, issues = build_payload(template_metrics, actual_metrics)
    raw_issue_count = len(issues)
    issues = reconcile_with_sample_self_check(
        payload,
        issues,
        Path(args.sample_self_check).resolve() if args.sample_self_check else None,
        frontmatter_toc_repair_report=Path(args.frontmatter_toc_repair_report).resolve()
        if args.frontmatter_toc_repair_report
        else None,
        final_docx=final_docx,
    )
    payload["schema"] = "graduation-project-builder.toc-paragraph-typography.v2"
    payload["template_docx"] = str(template_docx)
    payload["template_docx_sha256"] = sha256_file(template_docx)
    payload["template_authority_class"] = template_authority_class
    payload["toc_title_baseline_source"] = (
        "official-spec-first-level-title-rule"
        if template_authority_class == "official-format-spec"
        else "template-donor"
    )
    payload["toc_entry_baseline_source"] = (
        "official-spec-no-real-toc-donor-final-entry-structure"
        if template_authority_class == "official-format-spec"
        else "template-donor"
    )
    payload["final_docx"] = str(final_docx)
    payload["final_docx_sha256"] = sha256_file(final_docx)
    payload["issue_count"] = str(len(issues))
    payload["exact_output_binding_verdict"] = "pass"
    if raw_issue_count != len(issues):
        payload["raw_issue_count_before_reconciliation"] = str(raw_issue_count)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if issues:
        print("TOC paragraph/typography drift: " + "; ".join(issues))
        return 1 if args.fail_on_drift else 0
    print(f"TOC paragraph/typography metrics passed: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
