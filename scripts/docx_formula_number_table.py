from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{%s}" % NS["w"]
W14 = "{%s}" % NS["w14"]
M = "{%s}" % NS["m"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("w", NS["w"])
ET.register_namespace("w14", NS["w14"])
ET.register_namespace("m", NS["m"])

PPR_AFTER_JC_TAGS = {
    W + "textDirection",
    W + "textAlignment",
    W + "textboxTightWrap",
    W + "outlineLvl",
    W + "divId",
    W + "cnfStyle",
    W + "rPr",
    W + "sectPr",
    W + "pPrChange",
}
PPR_AFTER_TABS_TAGS = {
    W + "suppressAutoHyphens",
    W + "kinsoku",
    W + "wordWrap",
    W + "overflowPunct",
    W + "topLinePunct",
    W + "autoSpaceDE",
    W + "autoSpaceDN",
    W + "bidi",
    W + "adjustRightInd",
    W + "snapToGrid",
    W + "spacing",
    W + "ind",
    W + "contextualSpacing",
    W + "mirrorIndents",
    W + "suppressOverlap",
    W + "jc",
    W + "textDirection",
    W + "textAlignment",
    W + "textboxTightWrap",
    W + "outlineLvl",
    W + "divId",
    W + "cnfStyle",
    W + "rPr",
    W + "sectPr",
    W + "pPrChange",
}

FORMULA_NUMBER_CORE = r"\d+(?:[.-]\d+[A-Za-z]?)?"
FORMULA_NUMBER_CELL_RE = re.compile(rf"^(\u5f0f\s*)?[\uff08(]\s*(?:\u5f0f\s*)?({FORMULA_NUMBER_CORE})\s*[\uff09)]$")
FORMULA_NUMBER_LABEL_RE = re.compile(rf"\u5f0f\s*[\uff08(]\s*{FORMULA_NUMBER_CORE}\s*[\uff09)]")
FORMULA_NUMBER_LABEL_COMPACT_RE = re.compile(rf"\u5f0f[\uff08(]{FORMULA_NUMBER_CORE}[\uff09)]")
ANY_FORMULA_NUMBER_LABEL_RE = re.compile(rf"(?:\u5f0f\s*)?[\uff08(]\s*(?:\u5f0f\s*)?{FORMULA_NUMBER_CORE}\s*[\uff09)]")
CHAPTER_HEADING_RE = re.compile(r"^\s*\u7b2c\s*([0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+)\s*\u7ae0")
TAIL_HEADING_RE = re.compile(r"^\s*(?:\u53c2\u8003\u6587\u732e|\u81f4\u8c22|\u9644\u5f55|References|Acknowledgements?|Appendix)\b", re.I)
MATH_SIGNAL_RE = re.compile(r"(=|\uff1d|[<>≤≥]|\+|-|−|\*|/|×|÷|√|π|sin|cos|tan|σ|ε|η|Δ|△)", re.I)
MATH_SYMBOL_RE = re.compile(r"[A-Za-z\u0370-\u03ff]")
RAW_MATH_COMMAND_TOKEN_RE = re.compile(
    r"\b(?:sub|sup|frac|sqrt|over|below|above|nary|lim|eqarr)\b",
    re.I,
)
PROSE_FORMULA_SKIP_RE = re.compile(
    r"^(?:\u67e5\u8868|\u67e5\u56fe|\u56e0\u4e3a|\u6240\u4ee5|\u578b\u53f7|\u957f\u5ea6|\u5f0f\u4e2d|\u5176\u4e2d|"
    r"[A-Za-zΑ-Ωα-ω]\s*[\u2014\u2013-]{2})"
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def paragraph_math_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//m:t", NS)).strip()


def paragraph_contains_omml(paragraph: ET.Element) -> bool:
    return paragraph.find(".//m:oMathPara", NS) is not None or paragraph.find(".//m:oMath", NS) is not None


def paragraph_contains_formula_object(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//m:oMathPara", NS) is not None
        or paragraph.find(".//m:oMath", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
    )


def cjk_char_count(text_value: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text_value or ""))


def chinese_or_arabic_int(value: str, default: int = 1) -> int:
    value = re.sub(r"\s+", "", value or "")
    if value.isdigit():
        return int(value)
    digits = {
        "\u4e00": 1,
        "\u4e8c": 2,
        "\u4e09": 3,
        "\u56db": 4,
        "\u4e94": 5,
        "\u516d": 6,
        "\u4e03": 7,
        "\u516b": 8,
        "\u4e5d": 9,
    }
    if value == "\u5341":
        return 10
    if "\u5341" in value:
        left, _, right = value.partition("\u5341")
        tens = digits.get(left, 1 if not left else 0)
        ones = digits.get(right, 0) if right else 0
        parsed = tens * 10 + ones
        return parsed if parsed > 0 else default
    return digits.get(value, default)


def chapter_number_from_heading_text(text_value: str) -> int | None:
    match = CHAPTER_HEADING_RE.search(text_value or "")
    if not match:
        return None
    return chinese_or_arabic_int(match.group(1), default=1)


def should_number_standalone_formula_paragraph(paragraph: ET.Element) -> bool:
    contains_omml = paragraph_contains_omml(paragraph)
    contains_legacy_object = paragraph.find(".//w:object", NS) is not None
    if not contains_omml and not contains_legacy_object:
        return False
    visible_text = re.sub(r"\s+", "", paragraph_text(paragraph)).strip()
    math_text = re.sub(r"\s+", "", paragraph_math_text(paragraph)).strip()
    combined_text = visible_text + math_text
    if ANY_FORMULA_NUMBER_LABEL_RE.search(visible_text):
        return True
    if PROSE_FORMULA_SKIP_RE.search(visible_text):
        return False
    if contains_legacy_object and not contains_omml:
        if not visible_text or visible_text in {".", "\uff0c", ",", "\u2235", "\u2234"}:
            return True
        if MATH_SIGNAL_RE.search(visible_text) or re.search(r"\d", visible_text):
            return True
        if cjk_char_count(visible_text) <= 2 and len(visible_text) <= 6:
            return True
        return False
    if not math_text and not MATH_SIGNAL_RE.search(visible_text):
        return False
    has_formula_signal = bool(MATH_SIGNAL_RE.search(combined_text) or MATH_SYMBOL_RE.search(math_text))
    if not has_formula_signal:
        return False
    if cjk_char_count(visible_text) > 30 and len(math_text) <= 16:
        return False
    if cjk_char_count(visible_text) <= 18:
        return True
    if len(math_text) >= 18:
        return True
    return False


def root_namespace_declarations(document_xml: bytes) -> dict[str, str]:
    match = re.search(rb"<w:document\b(?P<attrs>[^>]*)>", document_xml[:12000], flags=re.S)
    if not match:
        return {}
    attrs = match.group("attrs")
    return {
        prefix.decode("ascii", errors="ignore"): uri.decode("utf-8", errors="ignore")
        for prefix, uri in re.findall(rb'\sxmlns:([A-Za-z0-9]+)="([^"]+)"', attrs)
    }


def serialize_document(root: ET.Element, original_document_xml: bytes) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    declarations = root_namespace_declarations(original_document_xml)
    root_start = payload.find(b"<w:document")
    if root_start < 0:
        return payload
    root_end = payload.find(b">", root_start)
    if root_end < 0:
        return payload
    root_tag = payload[root_start:root_end]
    additions: list[bytes] = []
    for prefix, uri in declarations.items():
        token = f"xmlns:{prefix}=".encode("ascii")
        if token not in root_tag:
            additions.append(f' xmlns:{prefix}="{uri}"'.encode("utf-8"))
    if not additions:
        return payload
    return payload[:root_end] + b"".join(additions) + payload[root_end:]


def add_paragraph_jc(ppr: ET.Element, value: str) -> None:
    for jc in list(ppr.findall(W + "jc")):
        ppr.remove(jc)
    element = ET.Element(W + "jc", {W + "val": value})
    for index, child in enumerate(list(ppr)):
        if child.tag in PPR_AFTER_JC_TAGS:
            ppr.insert(index, element)
            return
    ppr.append(element)


def add_right_tab(ppr: ET.Element, position: str = "9000") -> None:
    for tabs in list(ppr.findall(W + "tabs")):
        ppr.remove(tabs)
    tabs = ET.Element(W + "tabs")
    ET.SubElement(tabs, W + "tab", {W + "val": "right", W + "pos": position})
    for index, child in enumerate(list(ppr)):
        if child.tag in PPR_AFTER_TABS_TAGS:
            ppr.insert(index, tabs)
            return
    ppr.append(tabs)


def add_formula_tabs(ppr: ET.Element, center_position: str, right_position: str) -> None:
    for tabs in list(ppr.findall(W + "tabs")):
        ppr.remove(tabs)
    tabs = ET.Element(W + "tabs")
    ET.SubElement(tabs, W + "tab", {W + "val": "center", W + "pos": center_position})
    ET.SubElement(tabs, W + "tab", {W + "val": "right", W + "pos": right_position})
    for index, child in enumerate(list(ppr)):
        if child.tag in PPR_AFTER_TABS_TAGS:
            ppr.insert(index, tabs)
            return
    ppr.append(tabs)


def remove_first_child(parent: ET.Element, tag: str) -> None:
    node = parent.find(tag)
    if node is not None:
        parent.remove(node)


def sanitize_generated_formula_ppr(ppr: ET.Element) -> None:
    """Generated formulas must not inherit heading/list/page-break behavior."""

    for tag in (
        W + "pStyle",
        W + "numPr",
        W + "pageBreakBefore",
        W + "keepNext",
        W + "keepLines",
        W + "outlineLvl",
    ):
        remove_first_child(ppr, tag)


def remove_paragraph_indentation(ppr: ET.Element) -> None:
    for ind in list(ppr.findall(W + "ind")):
        ppr.remove(ind)


def attr_int(element: ET.Element | None, name: str, default: int = 0) -> int:
    if element is None:
        return default
    value = element.attrib.get(W + name)
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def content_width_from_sect_pr(sect_pr: ET.Element | None) -> int:
    if sect_pr is None:
        return 9000
    page_size = sect_pr.find(W + "pgSz")
    page_margin = sect_pr.find(W + "pgMar")
    page_width = attr_int(page_size, "w", 11906)
    left_margin = attr_int(page_margin, "left", 1800)
    right_margin = attr_int(page_margin, "right", 1800)
    content_width = page_width - left_margin - right_margin
    return content_width if content_width > 0 else 9000


def section_content_width_for_body_index(body: ET.Element, index: int) -> int:
    children = list(body)
    for child in children[index:]:
        sect_pr = None
        if child.tag == W + "sectPr":
            sect_pr = child
        elif child.tag == W + "p":
            ppr = child.find(W + "pPr")
            if ppr is not None:
                sect_pr = ppr.find(W + "sectPr")
        if sect_pr is not None:
            return content_width_from_sect_pr(sect_pr)
    return content_width_from_sect_pr(body.find(W + "sectPr"))


def paragraph_indent_adjustment(paragraph: ET.Element) -> int:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        return 0
    ind = ppr.find(W + "ind")
    if ind is None:
        return 0
    return max(0, attr_int(ind, "left", 0)) + max(0, attr_int(ind, "right", 0))


def formula_tab_positions(content_width_twips: int, paragraph: ET.Element) -> tuple[str, str]:
    usable_width = max(2000, content_width_twips - paragraph_indent_adjustment(paragraph))
    return str(usable_width // 2), str(usable_width)


def remove_paragraph_jc(ppr: ET.Element) -> None:
    for jc in list(ppr.findall(W + "jc")):
        ppr.remove(jc)


def make_text_run(text_value: str) -> ET.Element:
    run = ET.Element(W + "r")
    text = ET.SubElement(run, W + "t")
    text.text = text_value
    return run


def make_formula_number_run(text_value: str) -> ET.Element:
    run = ET.Element(W + "r")
    rpr = ET.SubElement(run, W + "rPr")
    ET.SubElement(
        rpr,
        W + "rFonts",
        {
            W + "ascii": "Times New Roman",
            W + "hAnsi": "Times New Roman",
            W + "eastAsia": "宋体",
            W + "cs": "Times New Roman",
        },
    )
    ET.SubElement(rpr, W + "sz", {W + "val": "21"})
    ET.SubElement(rpr, W + "szCs", {W + "val": "21"})
    text = ET.SubElement(run, W + "t")
    text.text = text_value
    return run


def make_tab_run() -> ET.Element:
    run = ET.Element(W + "r")
    ET.SubElement(run, W + "tab")
    return run


def run_has_direct_tab(run: ET.Element) -> bool:
    return run.tag == W + "r" and run.find(W + "tab") is not None


def clone_run_properties(run: ET.Element) -> ET.Element | None:
    rpr = run.find(W + "rPr")
    return deepcopy(rpr) if rpr is not None else None


def make_text_run_like(text_value: str, source_run: ET.Element | None = None) -> ET.Element:
    run = ET.Element(W + "r")
    if source_run is not None:
        rpr = clone_run_properties(source_run)
        if rpr is not None:
            run.append(rpr)
    text = ET.SubElement(run, W + "t")
    if text_value[:1].isspace() or text_value[-1:].isspace():
        text.set(XML_SPACE, "preserve")
    text.text = text_value
    return run


def make_break_run() -> ET.Element:
    run = ET.Element(W + "r")
    ET.SubElement(run, W + "br")
    return run


def add_math_run(parent: ET.Element, text: str) -> None:
    run = ET.SubElement(parent, M + "r")
    text_node = ET.SubElement(run, M + "t")
    text_node.text = text


def add_segments(parent: ET.Element, segments: list[object]) -> None:
    for segment in segments:
        if isinstance(segment, str):
            add_math_run(parent, segment)
            continue
        if not isinstance(segment, dict):
            raise ValueError(f"Unsupported formula segment: {segment!r}")
        if "text" in segment:
            add_math_run(parent, str(segment["text"]))
            continue
        if "sub" in segment:
            value = segment["sub"]
            if not isinstance(value, dict):
                raise ValueError(f"Formula sub segment must be an object: {segment!r}")
            sub = ET.SubElement(parent, M + "sSub")
            base = ET.SubElement(sub, M + "e")
            add_segments(base, [str(value.get("base", ""))])
            subscript = ET.SubElement(sub, M + "sub")
            add_segments(subscript, [str(value.get("subscript", ""))])
            continue
        if "frac" in segment:
            value = segment["frac"]
            if not isinstance(value, dict):
                raise ValueError(f"Formula frac segment must be an object: {segment!r}")
            frac = ET.SubElement(parent, M + "f")
            num = ET.SubElement(frac, M + "num")
            add_segments(num, list(value.get("num", [])))
            den = ET.SubElement(frac, M + "den")
            add_segments(den, list(value.get("den", [])))
            continue
        if "abs" in segment:
            add_math_run(parent, "|")
            add_segments(parent, list(segment["abs"]))
            add_math_run(parent, "|")
            continue
        raise ValueError(f"Unsupported formula segment keys: {segment!r}")


def make_formula_paragraph(template_para: ET.Element, segments: list[object], keep_paragraph_attrs: bool = True) -> ET.Element:
    formula_para = ET.Element(W + "p", dict(template_para.attrib) if keep_paragraph_attrs else {})
    ppr = template_para.find(W + "pPr")
    if ppr is not None:
        ppr_copy = deepcopy(ppr)
    else:
        ppr_copy = ET.Element(W + "pPr")
    sanitize_generated_formula_ppr(ppr_copy)
    add_paragraph_jc(ppr_copy, "center")
    formula_para.append(ppr_copy)
    omath_para = ET.SubElement(formula_para, M + "oMathPara")
    omath = ET.SubElement(omath_para, M + "oMath")
    add_segments(omath, segments)
    return formula_para


def make_inline_formula_paragraph(
    template_para: ET.Element,
    lines: list[list[object]],
    number_text: str,
    content_width_twips: int,
) -> ET.Element:
    formula_para = ET.Element(W + "p", dict(template_para.attrib))
    ppr = template_para.find(W + "pPr")
    if ppr is not None:
        ppr_copy = deepcopy(ppr)
    else:
        ppr_copy = ET.Element(W + "pPr")
    sanitize_generated_formula_ppr(ppr_copy)
    center_position, right_position = formula_tab_positions(content_width_twips, template_para)
    add_formula_tabs(ppr_copy, center_position=center_position, right_position=right_position)
    remove_paragraph_jc(ppr_copy)
    formula_para.append(ppr_copy)
    for index, line in enumerate(lines):
        if index > 0:
            formula_para.append(make_break_run())
        formula_para.append(make_tab_run())
        omath = ET.SubElement(formula_para, M + "oMath")
        add_segments(omath, line)
        if index == 0:
            formula_para.append(make_tab_run())
            formula_para.append(make_text_run(number_text))
    return formula_para


def make_borderless_table_from_formula_paragraphs(
    formula_paras: list[ET.Element],
    number_text: str,
    total_width_twips: int = 9000,
) -> ET.Element:
    if not formula_paras:
        raise ValueError("formula_paras must not be empty")
    total_width = max(3600, total_width_twips)
    side_width = min(1800, max(1440, total_width // 6))
    middle_width = total_width - (side_width * 2)
    tbl = ET.Element(W + "tbl")
    tbl_pr = ET.SubElement(tbl, W + "tblPr")
    ET.SubElement(tbl_pr, W + "tblW", {W + "w": str(total_width), W + "type": "dxa"})
    tbl_borders = ET.SubElement(tbl_pr, W + "tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(tbl_borders, W + edge, {W + "val": "nil"})
    ET.SubElement(tbl_pr, W + "tblLayout", {W + "type": "fixed"})

    tbl_grid = ET.SubElement(tbl, W + "tblGrid")
    ET.SubElement(tbl_grid, W + "gridCol", {W + "w": str(side_width)})
    ET.SubElement(tbl_grid, W + "gridCol", {W + "w": str(middle_width)})
    ET.SubElement(tbl_grid, W + "gridCol", {W + "w": str(side_width)})

    tr = ET.SubElement(tbl, W + "tr")
    left_tc = ET.SubElement(tr, W + "tc")
    left_tc_pr = ET.SubElement(left_tc, W + "tcPr")
    ET.SubElement(left_tc_pr, W + "tcW", {W + "w": str(side_width), W + "type": "dxa"})
    ET.SubElement(left_tc, W + "p")

    middle_tc = ET.SubElement(tr, W + "tc")
    middle_tc_pr = ET.SubElement(middle_tc, W + "tcPr")
    ET.SubElement(middle_tc_pr, W + "tcW", {W + "w": str(middle_width), W + "type": "dxa"})
    for formula_para in formula_paras:
        middle_tc.append(formula_para)

    right_tc = ET.SubElement(tr, W + "tc")
    right_tc_pr = ET.SubElement(right_tc, W + "tcPr")
    ET.SubElement(right_tc_pr, W + "tcW", {W + "w": str(side_width), W + "type": "dxa"})
    ET.SubElement(right_tc_pr, W + "vAlign", {W + "val": "center"})
    right_para = ET.SubElement(right_tc, W + "p")
    right_ppr = ET.SubElement(right_para, W + "pPr")
    add_paragraph_jc(right_ppr, "right")
    right_para.append(make_formula_number_run(normalize_formula_number_text(number_text) or number_text))
    return tbl


def cell_visible_text(cell: ET.Element) -> str:
    return "".join(node.text or "" for node in cell.findall(".//w:t", NS)).strip()


def clear_split_formula_label_runs_in_cell(cell: ET.Element) -> bool:
    runs = [run for run in cell.findall(".//w:r", NS) if run.findall("./w:t", NS)]
    pieces: list[tuple[ET.Element, ET.Element, str, int, int]] = []
    joined = ""
    for run in runs:
        text_nodes = run.findall("./w:t", NS)
        if not text_nodes:
            continue
        text = "".join(node.text or "" for node in text_nodes)
        start = len(joined)
        joined += text
        pieces.append((run, text_nodes[0], text, start, len(joined)))
    compact = re.sub(r"\s+", "", joined)
    if not FORMULA_NUMBER_LABEL_COMPACT_RE.search(compact):
        return False
    cleaned = FORMULA_NUMBER_LABEL_RE.sub("", joined)
    cleaned = re.sub(rf"\u5f0f\s*[\uff08(]\s*{FORMULA_NUMBER_CORE}\s*[\uff09)]", "", cleaned)
    cleaned = re.sub(r"[\t ]+$", "", cleaned)
    cleaned = re.sub(r"[\uff0f/]\s*$", "", cleaned)
    if cleaned == joined:
        return False
    for index, (_run, node, _text, _start, _end) in enumerate(pieces):
        if index == 0:
            node.text = cleaned
            if cleaned[:1].isspace() or cleaned[-1:].isspace():
                node.set(XML_SPACE, "preserve")
        else:
            node.text = ""
    return True


def strip_trailing_formula_number_label(text_value: str) -> tuple[str, bool]:
    """Remove a trailing old formula number from a text fragment."""

    matches = list(ANY_FORMULA_NUMBER_LABEL_RE.finditer(text_value or ""))
    if not matches:
        return text_value, False
    match = matches[-1]
    if (text_value[match.end() :] or "").strip():
        return text_value, False
    before = text_value[: match.start()].rstrip(" \t\u3000")
    after = text_value[match.end() :]
    if not before and not after.strip():
        return "", True
    return before + after, True


def strip_formula_number_from_math_text_nodes(paragraph: ET.Element) -> int:
    """Remove old visible formula labels embedded inside OMML math text."""

    removed = 0
    for text_node in paragraph.findall(".//m:t", NS):
        original = text_node.text or ""
        cleaned, changed = strip_trailing_formula_number_label(original)
        if not changed:
            continue
        text_node.text = cleaned
        if cleaned[:1].isspace() or cleaned[-1:].isspace():
            text_node.set(XML_SPACE, "preserve")
        else:
            text_node.attrib.pop(XML_SPACE, None)
        removed += 1
    return removed


def clear_cell_content_keep_properties(cell: ET.Element) -> None:
    for child in list(cell):
        if child.tag != W + "tcPr":
            cell.remove(child)


def remove_invalid_cell_no_wrap(cell: ET.Element) -> int:
    tcpr = cell.find(W + "tcPr")
    if tcpr is None:
        return 0
    removed = 0
    for child in list(tcpr):
        if child.tag == W + "noWrap":
            tcpr.remove(child)
            removed += 1
    return removed


def ensure_cell_width(cell: ET.Element, width_twips: int) -> None:
    tcpr = cell.find(W + "tcPr")
    if tcpr is None:
        tcpr = ET.Element(W + "tcPr")
        cell.insert(0, tcpr)
    tcw = tcpr.find(W + "tcW")
    if tcw is None:
        tcw = ET.Element(W + "tcW", {W + "w": str(width_twips), W + "type": "dxa"})
        tcpr.insert(0, tcw)
    else:
        try:
            current_width = int(tcw.get(W + "w", "0") or "0")
        except ValueError:
            current_width = 0
        tcw.set(W + "w", str(max(width_twips, current_width)))
        tcw.set(W + "type", "dxa")


def ensure_formula_table_widths(table: ET.Element, number_cell_width_twips: int = 1500) -> None:
    tbl_pr = table.find(W + "tblPr")
    total_width = 9000
    if tbl_pr is not None:
        tbl_w = tbl_pr.find(W + "tblW")
        if tbl_w is not None:
            try:
                total_width = max(total_width, int(tbl_w.get(W + "w", "0") or "0"))
            except ValueError:
                pass
    if tbl_pr is None:
        tbl_pr = ET.Element(W + "tblPr")
        table.insert(0, tbl_pr)
    tbl_w = tbl_pr.find(W + "tblW")
    if tbl_w is None:
        tbl_w = ET.Element(W + "tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(W + "w", str(total_width))
    tbl_w.set(W + "type", "dxa")
    side_width = min(1800, max(number_cell_width_twips, total_width // 6))
    middle_width = max(2400, total_width - (side_width * 2))
    grid = table.find(W + "tblGrid")
    if grid is not None:
        cols = grid.findall(W + "gridCol")
        if len(cols) >= 3:
            cols[0].set(W + "w", str(side_width))
            cols[1].set(W + "w", str(middle_width))
            cols[2].set(W + "w", str(side_width))
    for row in table.findall("./w:tr", NS):
        cells = row.findall("./w:tc", NS)
        if len(cells) >= 3:
            ensure_cell_width(cells[0], side_width)
            ensure_cell_width(cells[1], middle_width)
            ensure_cell_width(cells[-1], side_width)


def normalize_formula_number_text(text_value: str, *, number_format: str = "hyphen") -> str | None:
    compact = re.sub(r"\s+", "", text_value or "")
    match = FORMULA_NUMBER_CELL_RE.fullmatch(compact)
    if not match:
        return None
    prefix = "\u5f0f"
    core = match.group(2)
    if number_format == "hyphen":
        core = core.replace(".", "-")
    elif number_format == "dot":
        raise ValueError("formula number format 'dot' is no longer allowed; use 式(6-1) style with a hyphen")
    elif number_format != "preserve":
        raise ValueError(f"unsupported formula number format: {number_format}")
    return f"{prefix}({core})"


def normalize_formula_number_key(text_value: str, *, number_format: str = "hyphen") -> str | None:
    compact = re.sub(r"\s+", "", text_value or "")
    match = FORMULA_NUMBER_CELL_RE.fullmatch(compact)
    if not match:
        return None
    core = match.group(2)
    if number_format == "hyphen":
        core = core.replace(".", "-")
    elif number_format == "dot":
        raise ValueError("formula number format 'dot' is no longer allowed; use formula-word labels with a hyphen")
    elif number_format != "preserve":
        raise ValueError(f"unsupported formula number format: {number_format}")
    return core


def repair_formula_number_cells(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
    number_format: str = "hyphen",
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    repaired: list[dict[str, object]] = []
    skipped_math_rows = 0
    removed_invalid_no_wrap_count = 0

    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        if table.find(".//m:oMath", NS) is None and table.find(".//m:oMathPara", NS) is None:
            continue
        ensure_formula_table_widths(table)
        for row_index, row in enumerate(table.findall("./w:tr", NS), start=1):
            cells = row.findall("./w:tc", NS)
            if not cells:
                continue
            if not any(cell.find(".//m:oMath", NS) is not None or cell.find(".//m:oMathPara", NS) is not None for cell in cells):
                continue
            number_cell = cells[-1]
            before_text = cell_visible_text(number_cell)
            number_text = normalize_formula_number_text(before_text, number_format=number_format)
            if number_text is None:
                skipped_math_rows += 1
                continue
            clear_cell_content_keep_properties(number_cell)
            removed_invalid_no_wrap_count += remove_invalid_cell_no_wrap(number_cell)
            ensure_cell_width(number_cell, 1500)
            paragraph = ET.SubElement(number_cell, W + "p")
            ppr = ET.SubElement(paragraph, W + "pPr")
            add_paragraph_jc(ppr, "right")
            paragraph.append(make_formula_number_run(number_text))
            repaired.append(
                {
                    "table_index": table_index,
                    "row_index": row_index,
                    "before": before_text,
                    "after": number_text,
                }
            )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.formula-number-cell-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "repaired_formula_number_cell_count": len(repaired),
        "skipped_math_row_count": skipped_math_rows,
        "removed_invalid_cell_no_wrap_count": removed_invalid_no_wrap_count,
        "formula_number_format": number_format,
        "normalization": "normalize visible formula numbers to the required 式(6-1) family, optionally normalize chapter separator, remove schema-invalid cell noWrap, widen the number column, and use single-line labels with explicit 10.5pt font model",
        "repairs": repaired,
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    return report


def rewrite_standalone_formula_number_paragraph(paragraph: ET.Element, *, number_format: str) -> tuple[bool, str, str]:
    text_nodes = paragraph.findall(".//w:t", NS)
    before = "".join(node.text or "" for node in text_nodes)
    normalized = normalize_formula_number_text(before.strip(), number_format=number_format)
    if normalized is None:
        return False, before, before
    prefix_match = re.match(r"^\s*", before)
    suffix_match = re.search(r"\s*$", before)
    prefix = prefix_match.group(0) if prefix_match else ""
    suffix = suffix_match.group(0) if suffix_match else ""
    if paragraph_contains_formula_object(paragraph) and len(prefix) > 1:
        prefix = " "
    after = f"{prefix}{normalized}{suffix}"
    if after == before:
        return False, before, after
    first_text_node = next((node for node in text_nodes if (node.text or "").strip()), None)
    if first_text_node is None and text_nodes:
        first_text_node = text_nodes[0]
    if first_text_node is None:
        return False, before, before
    first_text_node.text = after
    if prefix or suffix:
        first_text_node.set(XML_SPACE, "preserve")
        first_text_node.attrib.pop(W + "space", None)
    else:
        first_text_node.attrib.pop(XML_SPACE, None)
        first_text_node.attrib.pop(W + "space", None)
    for node in text_nodes:
        if node is not first_text_node:
            node.text = ""
    return True, before, after


def repair_standalone_formula_number_labels(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
    number_format: str = "hyphen",
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    repaired: list[dict[str, object]] = []
    for body_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p":
            continue
        changed, before, after = rewrite_standalone_formula_number_paragraph(child, number_format=number_format)
        if not changed:
            continue
        repaired.append(
            {
                "body_child_index": body_index,
                "before": before,
                "after": after,
            }
        )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.standalone-formula-number-label-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if repaired else [],
        "repaired_standalone_label_count": len(repaired),
        "formula_number_format": number_format,
        "normalization": "normalize standalone visible formula-number labels to the required formula-word prefix, ASCII parentheses, and hyphen chapter numbering",
        "repairs": repaired,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    return report


def ensure_formula_center_tab_run(paragraph: ET.Element) -> bool:
    """Ensure a standalone formula paragraph has a real tab run before content."""

    children = list(paragraph)
    insert_at = 1 if children and children[0].tag == W + "pPr" else 0
    for child in children[insert_at:]:
        if run_has_direct_tab(child):
            return False
        if child.tag == W + "r" or child.tag.startswith(M) or child.tag == W + "object":
            paragraph.insert(insert_at, make_tab_run())
            return True
    paragraph.insert(insert_at, make_tab_run())
    return True


def normalize_formula_label_tab_runs(paragraph: ET.Element) -> int:
    """Split visible formula labels from text runs and put a real tab before them."""

    children = list(paragraph)
    new_children: list[ET.Element] = []
    inserted_tabs = 0
    changed = False

    for child in children:
        if child.tag != W + "r":
            new_children.append(child)
            continue
        direct_text_nodes = [node for node in list(child) if node.tag == W + "t"]
        if not direct_text_nodes:
            new_children.append(child)
            continue
        if any(node.tag not in {W + "rPr", W + "t"} for node in list(child)):
            new_children.append(child)
            continue
        run_text = "".join(node.text or "" for node in direct_text_nodes)
        matches = list(ANY_FORMULA_NUMBER_LABEL_RE.finditer(run_text))
        if not matches:
            new_children.append(child)
            continue

        match = matches[-1]
        before = run_text[: match.start()].rstrip(" \t\u3000")
        label_raw = match.group(0)
        label = normalize_formula_number_text(label_raw, number_format="hyphen") or re.sub(r"\s+", "", label_raw)
        after = run_text[match.end() :]

        if before:
            new_children.append(make_text_run_like(before, child))
        if not new_children or not run_has_direct_tab(new_children[-1]):
            new_children.append(make_tab_run())
            inserted_tabs += 1
        new_children.append(make_text_run_like(label, child))
        if after:
            new_children.append(make_text_run_like(after, child))
        changed = True

    if changed:
        paragraph[:] = new_children
    return inserted_tabs


def strip_formula_number_from_formula_paragraph(paragraph: ET.Element) -> int:
    """Remove the original visible formula number before moving a formula into a table."""

    removed = strip_formula_number_from_math_text_nodes(paragraph)
    children = list(paragraph)
    new_children: list[ET.Element] = []
    changed = False

    for child in children:
        if child.tag != W + "r":
            new_children.append(child)
            continue
        direct_text_nodes = [node for node in list(child) if node.tag == W + "t"]
        if not direct_text_nodes:
            new_children.append(child)
            continue
        if any(node.tag not in {W + "rPr", W + "t"} for node in list(child)):
            new_children.append(child)
            continue
        run_text = "".join(node.text or "" for node in direct_text_nodes)
        matches = list(ANY_FORMULA_NUMBER_LABEL_RE.finditer(run_text))
        if not matches:
            new_children.append(child)
            continue

        match = matches[-1]
        before = run_text[: match.start()].rstrip(" \t\u3000")
        after = run_text[match.end() :]
        if new_children and run_has_direct_tab(new_children[-1]) and not before:
            new_children.pop()
        if before:
            new_children.append(make_text_run_like(before, child))
        if after:
            new_children.append(make_text_run_like(after, child))
        removed += 1
        changed = True

    if changed:
        paragraph[:] = new_children
    return removed


def formula_labels_in_element(element: ET.Element) -> list[str]:
    labels: list[str] = []
    for match in FORMULA_NUMBER_LABEL_RE.finditer(paragraph_text(element)):
        normalized = normalize_formula_number_text(match.group(0), number_format="hyphen")
        labels.append(normalized or re.sub(r"\s+", "", match.group(0)))
    return labels


def remove_duplicate_standalone_formula_labels(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    formula_object_labels: set[str] = set()
    for child in list(body):
        if child.find(".//m:oMath", NS) is not None or child.find(".//m:oMathPara", NS) is not None:
            formula_object_labels.update(formula_labels_in_element(child))

    seen_formula_labels: set[str] = set()
    removed: list[dict[str, object]] = []
    for body_index, child in list(enumerate(list(body), start=1)):
        if child.tag == W + "tbl":
            if child.find(".//m:oMath", NS) is not None or child.find(".//m:oMathPara", NS) is not None:
                seen_formula_labels.update(formula_labels_in_element(child))
                continue
            compact = re.sub(r"\s+", "", paragraph_text(child))
            normalized = normalize_formula_number_text(compact, number_format="hyphen")
            if normalized and normalized in formula_object_labels:
                body.remove(child)
                removed.append({"body_child_index": body_index, "label": normalized, "kind": "label_only_table"})
            continue
        if child.tag != W + "p":
            continue
        compact = re.sub(r"\s+", "", paragraph_text(child))
        normalized = normalize_formula_number_text(compact, number_format="hyphen")
        has_formula_object = paragraph_contains_formula_object(child)
        if normalized and not has_formula_object and (normalized in seen_formula_labels or normalized in formula_object_labels):
            body.remove(child)
            removed.append({"body_child_index": body_index, "label": normalized, "kind": "label_only_paragraph"})
            continue
        if has_formula_object:
            seen_formula_labels.update(formula_labels_in_element(child))

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.duplicate-standalone-formula-label-removal.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if removed else [],
        "removed_duplicate_label_count": len(removed),
        "removed_duplicate_labels": removed,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def repair_formula_paragraph_tabs(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    repaired: list[dict[str, object]] = []
    for body_index, child in enumerate(list(body), start=1):
        if child.tag != W + "p" or not paragraph_contains_formula_object(child):
            continue
        text = re.sub(r"\s+", "", paragraph_text(child))
        if not re.search(r"\u5f0f\(\d+-\d+\)", text):
            continue
        ppr = child.find(W + "pPr")
        if ppr is None:
            ppr = ET.Element(W + "pPr")
            child.insert(0, ppr)
        content_width = section_content_width_for_body_index(body, body_index - 1)
        center_position = str(int(round(content_width / 2)))
        right_position = str(content_width)
        add_formula_tabs(ppr, center_position=center_position, right_position=right_position)
        add_paragraph_jc(ppr, "left")
        remove_paragraph_indentation(ppr)
        center_tab_inserted = ensure_formula_center_tab_run(child)
        label_tab_inserted = normalize_formula_label_tab_runs(child)
        repaired.append(
            {
                "body_child_index": body_index,
                "text": text,
                "center_tab_twips": int(center_position),
                "right_tab_twips": int(right_position),
                "paragraph_alignment": "left",
                "indent_removed": True,
                "center_tab_run_inserted": center_tab_inserted,
                "label_tab_run_inserted": label_tab_inserted,
            }
        )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.formula-paragraph-tab-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if repaired else [],
        "repaired_formula_paragraph_count": len(repaired),
        "repairs": repaired,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def convert_standalone_formula_paragraphs_to_tables(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
    number_format: str = "hyphen",
) -> dict[str, object]:
    """Move direct body formula paragraphs into borderless formula tables.

    This keeps the existing OMML content but gives the formula number a real
    right-side table cell when paragraph tab rendering does not reach the
    visible numbering surface in exported PDF.
    """

    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    converted: list[dict[str, object]] = []
    for body_index, child in list(enumerate(list(body), start=1)):
        if child.tag != W + "p" or not paragraph_contains_formula_object(child):
            continue
        labels = formula_labels_in_element(child)
        if not labels:
            continue
        number_text = normalize_formula_number_text(labels[-1], number_format=number_format)
        if number_text is None:
            continue
        content_width_twips = max(9000, section_content_width_for_body_index(body, body_index - 1))
        table = make_borderless_table(child, number_text, content_width_twips)
        body.remove(child)
        body.insert(body_index - 1, table)
        converted.append(
            {
                "body_child_index": body_index,
                "number": number_text,
                "table_width_twips": content_width_twips,
            }
        )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.standalone-formula-table-conversion.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if converted else [],
        "converted_formula_paragraph_count": len(converted),
        "formula_number_format": number_format,
        "conversion": "standalone body formula paragraphs were moved into borderless three-cell tables so the visible formula number owns a right-side numbering cell while OMML content is preserved",
        "converted": converted,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def number_standalone_formula_paragraphs_to_tables(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
    number_format: str = "hyphen",
) -> dict[str, object]:
    """Number every standalone body formula paragraph by chapter and table it."""

    if number_format != "hyphen":
        raise ValueError("bulk formula numbering requires hyphen chapter numbering")
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    converted: list[dict[str, object]] = []
    skipped_inline: list[dict[str, object]] = []
    chapter_counters: dict[int, int] = {}
    for child in list(body):
        for label in formula_labels_in_element(child):
            match = re.search(r"\((\d+)-(\d+)", label)
            if not match:
                continue
            chapter_key = int(match.group(1))
            sequence = int(match.group(2))
            chapter_counters[chapter_key] = max(chapter_counters.get(chapter_key, 0), sequence)
    current_chapter = 1
    body_started = False
    tail_reached = False

    for _original_index, child in list(enumerate(list(body), start=1)):
        if child.tag != W + "p":
            continue
        text_value = paragraph_text(child).strip()
        chapter = chapter_number_from_heading_text(text_value)
        if chapter is not None:
            current_chapter = chapter
            body_started = True
            continue
        if body_started and TAIL_HEADING_RE.search(text_value):
            tail_reached = True
        if tail_reached or not body_started:
            continue
        if not paragraph_contains_formula_object(child):
            continue
        if formula_labels_in_element(child):
            continue
        if not should_number_standalone_formula_paragraph(child):
            skipped_inline.append(
                {
                    "text_prefix": text_value[:120],
                    "reason": "inline-math-prose-not-numbered",
                }
            )
            continue
        chapter_counters[current_chapter] = chapter_counters.get(current_chapter, 0) + 1
        number_text = f"\u5f0f({current_chapter}-{chapter_counters[current_chapter]})"
        current_index = list(body).index(child)
        content_width_twips = max(9000, section_content_width_for_body_index(body, current_index))
        table = make_borderless_table(child, number_text, content_width_twips)
        body.remove(child)
        body.insert(current_index, table)
        converted.append(
            {
                "body_child_index_before": current_index + 1,
                "chapter": current_chapter,
                "sequence": chapter_counters[current_chapter],
                "number": number_text,
                "text_prefix": text_value[:120],
                "table_width_twips": content_width_twips,
            }
        )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.bulk-standalone-formula-numbering.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if converted else [],
        "converted_formula_paragraph_count": len(converted),
        "skipped_inline_math_paragraph_count": len(skipped_inline),
        "formula_number_format": number_format,
        "chapter_formula_counts": {str(key): value for key, value in sorted(chapter_counters.items())},
        "conversion": "standalone body formula paragraphs were assigned chapter-sequence 式(章-序) labels and moved into borderless three-cell formula tables while inline prose math was preserved",
        "converted": converted,
        "skipped_inline_math_paragraphs": skipped_inline,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def strip_formula_labels_from_non_number_table_cells(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    """Remove formula-number residue from ordinary table cells.

    A formula number in a real numbering surface must be the whole right-side
    number cell.  When a label such as ``式(3-10)`` is appended to a data-table
    header cell, PDF extraction may see a split formula label even though the
    cell is not the formula-number surface.  This pass removes only those
    mixed-content residues and leaves dedicated formula-number cells untouched.
    """

    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    cleaned: list[dict[str, object]] = []
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        for cell_index, cell in enumerate(table.findall(".//w:tc", NS), start=1):
            cell_text_before = cell_visible_text(cell)
            if not cell_text_before:
                continue
            compact_before = re.sub(r"\s+", "", cell_text_before)
            if FORMULA_NUMBER_CELL_RE.fullmatch(compact_before):
                continue
            if not FORMULA_NUMBER_LABEL_COMPACT_RE.search(compact_before):
                continue
            changed = clear_split_formula_label_runs_in_cell(cell)
            if changed:
                cleaned.append(
                    {
                        "table_index": table_index,
                        "cell_index": cell_index,
                        "before": cell_text_before,
                        "after": cell_visible_text(cell),
                    }
                )

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.non-number-table-formula-label-cleanup.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"] if cleaned else [],
        "cleaned_cell_count": len(cleaned),
        "cleanup": "removed formula-number labels from ordinary mixed-content table cells while preserving dedicated formula-number cells",
        "cleaned": cleaned,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def make_formula_table_from_segments(
    template_para: ET.Element,
    segments: list[object],
    number_text: str,
    content_width_twips: int = 9000,
) -> ET.Element:
    formula_para = make_formula_paragraph(template_para, segments)
    return make_borderless_table_from_formula_paragraphs([formula_para], number_text, content_width_twips)


def make_inline_omml_formula(segments: list[object]) -> ET.Element:
    omath = ET.Element(M + "oMath")
    add_segments(omath, segments)
    return omath


def make_formula_table_from_lines(
    template_para: ET.Element,
    lines: list[list[object]],
    number_text: str,
    content_width_twips: int = 9000,
) -> ET.Element:
    formula_paras = [
        make_formula_paragraph(template_para, line, keep_paragraph_attrs=(index == 0))
        for index, line in enumerate(lines)
    ]
    return make_borderless_table_from_formula_paragraphs(formula_paras, number_text, content_width_twips)


def make_borderless_table(
    formula_para: ET.Element,
    number_text: str,
    content_width_twips: int = 9000,
) -> ET.Element:
    moved_para = ET.Element(W + "p", formula_para.attrib)
    ppr = formula_para.find(W + "pPr")
    if ppr is not None:
        moved_para.append(ppr)
    else:
        moved_para.append(ET.Element(W + "pPr"))
    add_paragraph_jc(moved_para.find(W + "pPr"), "center")
    for child in list(formula_para):
        if child.tag != W + "pPr":
            moved_para.append(child)
    strip_formula_number_from_formula_paragraph(moved_para)
    return make_borderless_table_from_formula_paragraphs([moved_para], number_text, content_width_twips)


def element_word_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def element_math_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//m:t", NS))


def make_formula_table_from_formula_spec(
    template_para: ET.Element,
    formula: dict[str, object],
    number_text: str,
    content_width_twips: int,
) -> ET.Element:
    lines = formula.get("lines")
    segments = formula.get("segments")
    if lines is not None:
        if not isinstance(lines, list):
            raise ValueError(f"formula row lines must be a list: {formula!r}")
        return make_formula_table_from_lines(template_para, list(lines), number_text, content_width_twips)
    if not isinstance(segments, list):
        raise ValueError(f"formula row missing segments or lines: {formula!r}")
    return make_formula_table_from_segments(template_para, list(segments), number_text, content_width_twips)


def first_formula_paragraph_in_element(element: ET.Element) -> ET.Element | None:
    for paragraph in element.findall(".//w:p", NS):
        if paragraph_contains_formula_object(paragraph):
            return paragraph
    return None


def replace_formula_tables_by_number(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    """Replace the OMML body of existing numbered formula tables.

    This is intentionally narrow: the visible formula-number cell is kept as a
    dedicated right-side surface, while the math object content is rebuilt from
    a locked formula spec.
    """
    shutil.copy2(docx_path, output_path)
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    formulas = formula_map.get("formulas")
    if not isinstance(formulas, list) or not formulas:
        raise ValueError("formula map must contain a non-empty formulas list")

    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    replacements: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    for formula in formulas:
        if not isinstance(formula, dict) or "match_text" not in formula:
            continue
        number_text = normalize_formula_number_text(str(formula.get("number", "")).strip())
        number_key = normalize_formula_number_key(str(formula.get("number", "")).strip())
        if number_text is None:
            raise ValueError(f"formula row missing valid number: {formula!r}")
        match_text = str(formula.get("match_text", ""))
        body_children = list(body)
        match_index = None
        match_table = None
        for index, child in enumerate(body_children):
            if child.tag != W + "tbl":
                continue
            if formula_table_number(child) != number_key:
                continue
            combined_text = element_word_text(child) + element_math_text(child)
            if match_text and match_text not in combined_text:
                continue
            match_index = index
            match_table = child
            break
        if match_index is None or match_table is None:
            missing.append({"number": number_text, "match_text": match_text})
            continue
        template_para = first_formula_paragraph_in_element(match_table) or first_formula_template_paragraph(body)
        if template_para is None:
            raise ValueError("no formula template paragraph is available")
        content_width_twips = section_content_width_for_body_index(body, match_index)
        replacement_table = make_formula_table_from_formula_spec(
            template_para,
            formula,
            number_text,
            content_width_twips,
        )
        before_text = element_word_text(match_table)
        before_math = element_math_text(match_table)
        body.remove(match_table)
        body.insert(match_index, replacement_table)
        replacements.append(
            {
                "body_child_index": match_index,
                "number": number_text,
                "before_text": before_text,
                "before_math_text": before_math,
                "after_math_text": element_math_text(replacement_table),
            }
        )

    if missing:
        raise ValueError(f"formula tables not found: {missing}")

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.formula-table-content-replacement.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "input_docx": str(docx_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_sha256": sha256_file(output_path),
        "formula_map": str(formula_map_path.resolve()),
        "replacement_count": len(replacements),
        "replacements": replacements,
    }
    if report_path is not None:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def replace_plain_numbered_formulas(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    """Replace plain text numbered formulas with real OMML formula tables."""
    shutil.copy2(docx_path, output_path)
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    formulas = formula_map.get("formulas")
    if not isinstance(formulas, list) or not formulas:
        raise ValueError("formula map must contain a non-empty formulas list")

    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    replacements: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    for formula in formulas:
        if not isinstance(formula, dict) or "anchor_text" not in formula:
            continue
        number_text = normalize_formula_number_text(str(formula.get("number", "")).strip())
        if number_text is None:
            raise ValueError(f"formula row missing valid number: {formula!r}")
        anchor_text = str(formula.get("anchor_text", ""))
        body_children = list(body)
        match_index = None
        match_para = None
        for index, child in enumerate(body_children):
            if child.tag != W + "p":
                continue
            if paragraph_contains_formula_object(child):
                continue
            if anchor_text and anchor_text not in paragraph_text(child):
                continue
            match_index = index
            match_para = child
            break
        if match_index is None or match_para is None:
            missing.append({"number": number_text, "anchor_text": anchor_text})
            continue

        removed_previous: list[dict[str, object]] = []
        remove_offsets = formula.get("remove_previous_offsets", [])
        if remove_offsets is None:
            remove_offsets = []
        if not isinstance(remove_offsets, list):
            raise ValueError(f"remove_previous_offsets must be a list: {formula!r}")
        for raw_offset in sorted({int(offset) for offset in remove_offsets}, reverse=True):
            remove_index = match_index - raw_offset
            if remove_index < 0:
                continue
            remove_child = body_children[remove_index]
            if remove_child.tag != W + "p" or paragraph_contains_formula_object(remove_child):
                continue
            removed_previous.append(
                {
                    "body_child_index": remove_index,
                    "text": paragraph_text(remove_child),
                }
            )
            body.remove(remove_child)
            if remove_index < match_index:
                match_index -= 1

        template_para = first_formula_template_paragraph(body) or match_para
        content_width_twips = section_content_width_for_body_index(body, match_index)
        replacement_table = make_formula_table_from_formula_spec(
            template_para,
            formula,
            number_text,
            content_width_twips,
        )
        before_text = paragraph_text(match_para)
        body.remove(match_para)
        body.insert(match_index, replacement_table)
        replacements.append(
            {
                "body_child_index": match_index,
                "number": number_text,
                "before_text": before_text,
                "removed_previous": removed_previous,
                "after_math_text": element_math_text(replacement_table),
            }
        )

    if missing:
        raise ValueError(f"plain numbered formula anchors not found: {missing}")

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.plain-numbered-formula-replacement.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "input_docx": str(docx_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_sha256": sha256_file(output_path),
        "formula_map": str(formula_map_path.resolve()),
        "replacement_count": len(replacements),
        "replacements": replacements,
    }
    if report_path is not None:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def patch_formula(docx_path: Path, para_id: str, number_text: str, output_path: Path | None) -> Path:
    out_path = output_path or docx_path
    if output_path:
        shutil.copy2(docx_path, out_path)

    with zipfile.ZipFile(out_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    index = None
    formula_para = None
    for i, child in enumerate(list(body)):
        if child.tag != W + "p":
            continue
        if child.attrib.get(W14 + "paraId") == para_id:
            formula_para = child
            index = i
            break
    if formula_para is None or index is None:
        raise ValueError(f"Formula paragraph {para_id} not found")
    if formula_para.find(".//m:oMathPara", NS) is None and formula_para.find(".//m:oMath", NS) is None:
        raise ValueError(f"Paragraph {para_id} does not contain a math object")

    content_width_twips = section_content_width_for_body_index(body, index)
    tbl = make_borderless_table(formula_para, number_text, content_width_twips)
    body.remove(formula_para)
    body.insert(index, tbl)

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)
    return out_path


def replace_text_formulas(docx_path: Path, formula_map_path: Path, output_path: Path, report_path: Path | None = None) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    formulas = formula_map.get("formulas")
    if not isinstance(formulas, list) or not formulas:
        raise ValueError("formula map must contain a non-empty formulas list")

    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    replacements: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    body_children = list(body)
    for formula in formulas:
        if not isinstance(formula, dict):
            raise ValueError(f"formula map row must be an object: {formula!r}")
        source_text = str(formula.get("source_text", "")).strip()
        number_text = str(formula.get("number", "")).strip()
        normalized_number_text = normalize_formula_number_text(number_text)
        segments = formula.get("segments")
        lines = formula.get("lines")
        if not source_text or not number_text or normalized_number_text is None:
            raise ValueError(f"formula map row missing source_text or number: {formula!r}")
        number_text = normalized_number_text
        if lines is not None and not isinstance(lines, list):
            raise ValueError(f"formula map row lines must be a list: {formula!r}")
        if lines is None and not isinstance(segments, list):
            raise ValueError(f"formula map row missing segments or lines: {formula!r}")
        match_index = None
        match_para = None
        for index, child in enumerate(body_children):
            if child.tag != W + "p":
                continue
            if paragraph_text(child) == source_text:
                match_index = index
                match_para = child
                break
        if match_index is None or match_para is None:
            missing.append({"source_text": source_text, "number": number_text})
            continue
        if lines is not None:
            line_segments = [list(line) for line in lines]
        else:
            line_segments = [list(segments)]
        content_width_twips = section_content_width_for_body_index(body, match_index)
        if str(formula.get("layout", "table")).strip().lower() == "paragraph":
            replacement = make_inline_formula_paragraph(match_para, line_segments, number_text, content_width_twips)
        elif lines is not None:
            replacement = make_formula_table_from_lines(match_para, line_segments, number_text, content_width_twips)
        else:
            replacement = make_formula_table_from_segments(match_para, line_segments[0], number_text, content_width_twips)
        body.remove(match_para)
        body.insert(match_index, replacement)
        body_children = list(body)
        replacements.append(
            {
                "source_text": source_text,
                "number": number_text,
                "body_child_index": match_index,
                "source_para_id": match_para.attrib.get(W14 + "paraId", ""),
            }
        )

    if missing:
        raise ValueError(f"formula text anchors not found: {missing}")

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.text-formula-omml-replacement.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "replacement_count": len(replacements),
        "replacements": replacements,
        "missing": missing,
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def replace_inline_text_formulas(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    formulas = formula_map.get("formulas")
    if not isinstance(formulas, list) or not formulas:
        raise ValueError("formula map must contain a non-empty formulas list")

    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    replacements: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    for formula in formulas:
        if not isinstance(formula, dict):
            raise ValueError(f"formula map row must be an object: {formula!r}")
        source_text = str(formula.get("source_text", ""))
        anchor_text = str(formula.get("anchor_text", ""))
        segments = formula.get("segments")
        if not anchor_text:
            raise ValueError(f"formula map row missing anchor_text: {formula!r}")
        if not isinstance(segments, list):
            raise ValueError(f"formula map row missing segments: {formula!r}")

        match_para = None
        match_text = ""
        for paragraph in body.findall(".//w:p", NS):
            text_value = paragraph_text(paragraph)
            if source_text and text_value != source_text:
                continue
            if anchor_text in text_value:
                match_para = paragraph
                match_text = text_value
                break
        if match_para is None:
            missing.append({"source_text": source_text, "anchor_text": anchor_text})
            continue

        runs = list(match_para.findall(W + "r"))
        run_texts = ["".join(node.text or "" for node in run.findall(".//w:t", NS)) for run in runs]
        joined = "".join(run_texts)
        anchor_start = joined.find(anchor_text)
        if anchor_start < 0:
            missing.append({"source_text": source_text, "anchor_text": anchor_text, "reason": "anchor not in direct run text"})
            continue
        anchor_end = anchor_start + len(anchor_text)
        boundaries: list[tuple[int, int, ET.Element, str]] = []
        cursor = 0
        for run, text_value in zip(runs, run_texts):
            next_cursor = cursor + len(text_value)
            boundaries.append((cursor, next_cursor, run, text_value))
            cursor = next_cursor
        touched = [item for item in boundaries if item[0] < anchor_end and item[1] > anchor_start]
        if not touched:
            missing.append({"source_text": source_text, "anchor_text": anchor_text, "reason": "anchor runs not found"})
            continue

        parent = match_para
        insert_at = list(parent).index(touched[0][2])
        first_run = touched[0][2]
        last_run = touched[-1][2]
        first_start, _first_end, _first_run, first_text = touched[0]
        last_start, _last_end, _last_run, last_text = touched[-1]
        before = first_text[: max(0, anchor_start - first_start)]
        after = last_text[max(0, anchor_end - last_start) :]

        prototype_run = deepcopy(first_run)
        for text_node in prototype_run.findall(".//w:t", NS):
            text_node.text = ""

        new_nodes: list[ET.Element] = []
        if before:
            before_run = deepcopy(prototype_run)
            text_node = before_run.find(".//w:t", NS)
            if text_node is None:
                text_node = ET.SubElement(before_run, W + "t")
            text_node.text = before
            new_nodes.append(before_run)
        new_nodes.append(make_inline_omml_formula(list(segments)))
        if after:
            after_run = deepcopy(prototype_run)
            text_node = after_run.find(".//w:t", NS)
            if text_node is None:
                text_node = ET.SubElement(after_run, W + "t")
            text_node.text = after
            new_nodes.append(after_run)

        for _start, _end, run, _text in touched:
            parent.remove(run)
        for offset, node in enumerate(new_nodes):
            parent.insert(insert_at + offset, node)
        replacements.append(
            {
                "source_text": match_text,
                "anchor_text": anchor_text,
                "source_para_id": match_para.attrib.get(W14 + "paraId", ""),
            }
        )

    if missing:
        raise ValueError(f"inline formula anchors not found: {missing}")

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.inline-text-formula-omml-replacement.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "replacement_count": len(replacements),
        "replacements": replacements,
        "missing": missing,
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def find_body_child_index(
    body: ET.Element,
    *,
    anchor_text: str,
    occurrence: str = "last",
) -> int | None:
    matches = [
        index
        for index, child in enumerate(list(body))
        if child.tag == W + "p" and anchor_text in paragraph_text(child)
    ]
    if not matches:
        return None
    if occurrence == "first":
        return matches[0]
    if occurrence == "last":
        return matches[-1]
    try:
        one_based = int(occurrence)
    except ValueError as exc:
        raise ValueError(f"unsupported anchor occurrence: {occurrence!r}") from exc
    if one_based <= 0 or one_based > len(matches):
        raise ValueError(f"anchor occurrence out of range for {anchor_text!r}: {occurrence}")
    return matches[one_based - 1]


def first_formula_template_paragraph(body: ET.Element) -> ET.Element | None:
    for child in list(body):
        if child.tag == W + "p" and paragraph_contains_formula_object(child):
            return child
    return None


def insert_formula_batch(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    batches = formula_map.get("batches") or formula_map.get("insertions")
    if not isinstance(batches, list) or not batches:
        raise ValueError("formula batch map must contain a non-empty batches or insertions list")

    with zipfile.ZipFile(docx_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml missing w:body")
    formula_template_para = first_formula_template_paragraph(body)

    inserted: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    for batch_index, batch in enumerate(batches, start=1):
        if not isinstance(batch, dict):
            raise ValueError(f"formula batch row must be an object: {batch!r}")
        anchor_text = str(batch.get("anchor_text", "")).strip()
        occurrence = str(batch.get("anchor_occurrence", "last")).strip() or "last"
        formulas = batch.get("formulas")
        if not anchor_text or not isinstance(formulas, list) or not formulas:
            raise ValueError(f"formula batch missing anchor_text or formulas: {batch!r}")
        anchor_index = find_body_child_index(body, anchor_text=anchor_text, occurrence=occurrence)
        if anchor_index is None:
            missing.append(
                {
                    "batch_index": batch_index,
                    "anchor_text": anchor_text,
                    "anchor_occurrence": occurrence,
                }
            )
            continue
        children = list(body)
        anchor_para = children[anchor_index]
        template_para = formula_template_para or anchor_para
        content_width_twips = section_content_width_for_body_index(body, anchor_index)
        insert_at = anchor_index + 1
        inserted_numbers: list[str] = []
        for formula in formulas:
            if not isinstance(formula, dict):
                raise ValueError(f"formula row must be an object: {formula!r}")
            number_text = str(formula.get("number", "")).strip()
            normalized_number_text = normalize_formula_number_text(number_text) or number_text
            segments = formula.get("segments")
            if not normalized_number_text or not isinstance(segments, list) or not segments:
                raise ValueError(f"formula row missing number or segments: {formula!r}")
            paragraph = make_inline_formula_paragraph(
                template_para,
                [list(segments)],
                normalized_number_text,
                content_width_twips,
            )
            body.insert(insert_at, paragraph)
            insert_at += 1
            inserted_numbers.append(normalized_number_text)
        inserted.append(
            {
                "batch_index": batch_index,
                "anchor_text": anchor_text,
                "anchor_occurrence": occurrence,
                "anchor_body_index": anchor_index,
                "inserted_count": len(inserted_numbers),
                "inserted_numbers": inserted_numbers,
            }
        )

    if missing:
        raise ValueError(f"formula batch anchors not found: {missing}")

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.formula-batch-insertion.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "batch_count": len(inserted),
        "inserted_formula_count": sum(int(item["inserted_count"]) for item in inserted),
        "inserted": inserted,
        "missing": missing,
        "verdict": "pass",
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def formula_map_numbers(formula_map_path: Path) -> set[str]:
    formula_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    batches = formula_map.get("batches") or formula_map.get("insertions")
    if not isinstance(batches, list) or not batches:
        raise ValueError("formula batch map must contain a non-empty batches or insertions list")
    numbers: set[str] = set()
    for batch in batches:
        if not isinstance(batch, dict):
            continue
        formulas = batch.get("formulas")
        if not isinstance(formulas, list):
            continue
        for formula in formulas:
            if not isinstance(formula, dict):
                continue
            normalized = normalize_formula_number_text(str(formula.get("number", "")).strip())
            if normalized:
                numbers.add(normalized)
    if not numbers:
        raise ValueError("formula batch map does not contain any formula numbers")
    return numbers


def paragraph_has_raw_formula_token(paragraph: ET.Element) -> bool:
    math_text = paragraph_math_text(paragraph)
    visible_text = paragraph_text(paragraph)
    return RAW_MATH_COMMAND_TOKEN_RE.search(f"{math_text} {visible_text}") is not None


def remove_raw_token_formula_batch(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    numbers = formula_map_numbers(formula_map_path)
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml missing w:body")

    removed: list[dict[str, object]] = []
    for child in list(body):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child)
        labels = {
            normalize_formula_number_text(match.group(0)) or re.sub(r"\s+", "", match.group(0))
            for match in ANY_FORMULA_NUMBER_LABEL_RE.finditer(text)
        }
        matched = sorted(label for label in labels if label in numbers)
        if not matched or not paragraph_has_raw_formula_token(child):
            continue
        removed.append(
            {
                "labels": matched,
                "text": text[:240],
                "math_text": paragraph_math_text(child)[:360],
            }
        )
        body.remove(child)

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.raw-token-formula-batch-removal.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "target_formula_number_count": len(numbers),
        "removed_formula_count": len(removed),
        "removed": removed,
        "changed_zip_parts": ["word/document.xml"],
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def replace_formula_batch(
    docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    temp_path = output_path.with_name(output_path.stem + ".raw-token-cleaned.tmp.docx")
    removal = remove_raw_token_formula_batch(
        docx_path=docx_path,
        formula_map_path=formula_map_path,
        output_path=temp_path,
        report_path=None,
    )
    insertion = insert_formula_batch(
        docx_path=temp_path,
        formula_map_path=formula_map_path,
        output_path=output_path,
        report_path=None,
    )
    try:
        temp_path.unlink()
    except OSError:
        pass
    report = {
        "schema": "graduation-project-builder.formula-batch-replacement.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "removed_formula_count": removal.get("removed_formula_count"),
        "inserted_formula_count": insertion.get("inserted_formula_count"),
        "removal": removal,
        "insertion": insertion,
        "changed_zip_parts": ["word/document.xml"],
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def repair_raw_sqrt_formula_tokens(
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    replacements: list[dict[str, object]] = []
    for index, node in enumerate(root.findall(".//m:t", NS)):
        original = node.text or ""
        fixed = re.sub(r"\bsqrt\s*\(", "√(", original)
        if fixed == original:
            continue
        node.text = fixed
        replacements.append({"math_text_index": index, "before": original, "after": fixed})

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.raw-sqrt-formula-token-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "replacement_count": len(replacements),
        "replacements": replacements,
        "changed_zip_parts": ["word/document.xml"],
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def clone_for_insertion(element: ET.Element, unique_suffix: str = "") -> ET.Element:
    cloned = deepcopy(element)
    shape_type_id_map: dict[str, str] = {}
    if unique_suffix:
        for node in cloned.iter():
            if node.tag.endswith("}shapetype"):
                old_id = node.attrib.get("id")
                if old_id:
                    new_id = f"{old_id}_{unique_suffix}"
                    shape_type_id_map[old_id] = new_id
                    node.attrib["id"] = new_id
        for node in cloned.iter():
            if node.tag.endswith("}shape"):
                old_id = node.attrib.get("id")
                if old_id:
                    node.attrib["id"] = f"{old_id}_{unique_suffix}"
                old_type = node.attrib.get("type")
                if old_type in shape_type_id_map:
                    node.attrib["type"] = shape_type_id_map[old_type]
                elif old_type and old_type.startswith("#") and old_type[1:] in shape_type_id_map:
                    node.attrib["type"] = "#" + shape_type_id_map[old_type[1:]]
    for node in cloned.iter():
        for attr_name in list(node.attrib):
            local_name = attr_name.rsplit("}", 1)[-1]
            if local_name in {
                "paraId",
                "textId",
                "rsidR",
                "rsidRDefault",
                "rsidP",
                "rsidRPr",
                "rsidDel",
            }:
                del node.attrib[attr_name]
    for tag_name in (W + "bookmarkStart", W + "bookmarkEnd", W + "lastRenderedPageBreak"):
        for parent in cloned.iter():
            for child in list(parent):
                if child.tag == tag_name:
                    parent.remove(child)
    if cloned.tag == W + "p":
        ppr = cloned.find(W + "pPr")
        if ppr is not None:
            for child in list(ppr):
                if child.tag == W + "sectPr":
                    ppr.remove(child)
    return cloned


def formula_table_number(table: ET.Element) -> str | None:
    text = "".join(node.text or "" for node in table.findall(".//w:t", NS))
    match = re.search(r"式\s*\(\s*(\d+)\s*-\s*(\d+)\s*\)", text)
    if not match:
        return None
    return f"{int(match.group(1))}-{int(match.group(2))}"


def is_context_paragraph(paragraph: ET.Element) -> bool:
    if paragraph.find(".//m:oMath", NS) is not None or paragraph.find(".//m:oMathPara", NS) is not None:
        return False
    if paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None:
        return False
    if paragraph.find(".//w:object", NS) is not None:
        return False
    text = paragraph_text(paragraph).strip()
    if not text:
        return False
    style = paragraph.find("w:pPr/w:pStyle", NS)
    style_id = style.get(W + "val", "") if style is not None else ""
    if style_id.lower().startswith("heading") or style_id.upper().startswith("TOC"):
        return False
    return True


def is_source_narrative_context_text(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if len(compact) < 18:
        return False
    if re.fullmatch(r"式\(\d+-\d+\)", compact):
        return False
    if re.search(r"[=<>±*/+\-]\s*\d", compact) and len(compact) < 28:
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", compact))


def nearest_context_paragraph(source_children: list[ET.Element], index: int, max_scan: int = 5) -> ET.Element | None:
    for scan_index in range(index - 1, max(-1, index - max_scan - 1), -1):
        candidate = source_children[scan_index]
        if candidate.tag == W + "p" and is_context_paragraph(candidate):
            return candidate
    return None


def make_note_paragraph_like(template_para: ET.Element, text_value: str) -> ET.Element:
    paragraph = ET.Element(W + "p")
    ppr = template_para.find(W + "pPr")
    if ppr is not None:
        paragraph.append(deepcopy(ppr))
    run = ET.SubElement(paragraph, W + "r")
    for template_run in template_para.findall(W + "r"):
        rpr = template_run.find(W + "rPr")
        if rpr is not None:
            run.append(deepcopy(rpr))
            break
    text = ET.SubElement(run, W + "t")
    text.text = text_value
    return paragraph


def collect_source_formula_blocks(
    source_body: ET.Element,
    chapter: str,
    include_context: bool = True,
) -> tuple[list[ET.Element], list[str]]:
    source_children = list(source_body)
    collected: list[ET.Element] = []
    numbers: list[str] = []
    for index, child in enumerate(source_children):
        if child.tag != W + "tbl":
            continue
        number = formula_table_number(child)
        if number is None:
            continue
        if number.split("-", 1)[0] != chapter:
            continue
        unique_suffix = f"src{chapter}_{index}"
        if include_context:
            context = nearest_context_paragraph(source_children, index)
            if context is not None:
                if is_source_narrative_context_text(paragraph_text(context)):
                    collected.append(clone_for_insertion(context, unique_suffix=unique_suffix))
                else:
                    collected.append(
                        make_note_paragraph_like(
                            context,
                            "本项校核承接参考计算稿的计算过程，说明相关结构参数、载荷参数与强度条件的代入和判定依据。",
                        )
                    )
            else:
                template_para = child.find(".//w:p", NS)
                if template_para is not None:
                    collected.append(
                        make_note_paragraph_like(
                            template_para,
                            "本项校核承接参考计算稿的计算过程，说明相关结构参数、载荷参数与强度条件的代入和判定依据。",
                        )
                    )
        collected.append(clone_for_insertion(child, unique_suffix=unique_suffix))
        numbers.append(number)
    return collected, numbers


def insert_source_formula_blocks(
    docx_path: Path,
    source_docx_path: Path,
    formula_map_path: Path,
    output_path: Path,
    report_path: Path | None = None,
) -> dict[str, object]:
    insertion_map = json.loads(formula_map_path.read_text(encoding="utf-8"))
    insertions = insertion_map.get("insertions") or insertion_map.get("batches")
    if not isinstance(insertions, list) or not insertions:
        raise ValueError("source formula insertion map must contain a non-empty insertions list")

    with zipfile.ZipFile(docx_path, "r") as zin:
        target_members = {name: zin.read(name) for name in zin.namelist()}
    target_document_xml = target_members["word/document.xml"]
    target_root = ET.fromstring(target_document_xml)
    target_body = target_root.find("w:body", NS)
    if target_body is None:
        raise ValueError("target word/document.xml missing w:body")

    with zipfile.ZipFile(source_docx_path, "r") as zin:
        source_root = ET.fromstring(zin.read("word/document.xml"))
    source_body = source_root.find("w:body", NS)
    if source_body is None:
        raise ValueError("source word/document.xml missing w:body")

    inserted: list[dict[str, object]] = []
    missing: list[dict[str, object]] = []
    for insertion_index, insertion in enumerate(insertions, start=1):
        if not isinstance(insertion, dict):
            raise ValueError(f"insertion row must be an object: {insertion!r}")
        anchor_text = str(insertion.get("target_anchor_text") or insertion.get("anchor_text") or "").strip()
        occurrence = str(insertion.get("target_anchor_occurrence") or insertion.get("anchor_occurrence") or "last").strip() or "last"
        source_chapter = str(insertion.get("source_chapter") or insertion.get("chapter") or "").strip()
        note_text = str(insertion.get("note_text") or "").strip()
        include_context = bool(insertion.get("include_context", True))
        if not anchor_text or not source_chapter:
            raise ValueError(f"insertion row missing target anchor or source chapter: {insertion!r}")
        anchor_index = find_body_child_index(target_body, anchor_text=anchor_text, occurrence=occurrence)
        if anchor_index is None:
            missing.append(
                {
                    "insertion_index": insertion_index,
                    "anchor_text": anchor_text,
                    "anchor_occurrence": occurrence,
                    "source_chapter": source_chapter,
                }
            )
            continue
        blocks, numbers = collect_source_formula_blocks(source_body, source_chapter, include_context=include_context)
        if not blocks:
            missing.append(
                {
                    "insertion_index": insertion_index,
                    "anchor_text": anchor_text,
                    "source_chapter": source_chapter,
                    "reason": "no source formula tables found for chapter",
                }
            )
            continue
        target_children = list(target_body)
        insert_at = anchor_index + 1
        if note_text:
            target_body.insert(insert_at, make_note_paragraph_like(target_children[anchor_index], note_text))
            insert_at += 1
        for block in blocks:
            target_body.insert(insert_at, block)
            insert_at += 1
        inserted.append(
            {
                "insertion_index": insertion_index,
                "anchor_text": anchor_text,
                "anchor_occurrence": occurrence,
                "anchor_body_index": anchor_index,
                "source_chapter": source_chapter,
                "inserted_formula_count": len(numbers),
                "inserted_block_count": len(blocks) + (1 if note_text else 0),
                "inserted_numbers": numbers,
            }
        )

    if missing:
        raise ValueError(f"source formula insertion anchors or source blocks not found: {missing}")

    target_members["word/document.xml"] = serialize_document(target_root, target_document_xml)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in target_members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.source-formula-block-insertion.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "target_docx": str(docx_path.resolve()),
        "target_docx_sha256": sha256_file(docx_path),
        "source_docx": str(source_docx_path.resolve()),
        "source_docx_sha256": sha256_file(source_docx_path),
        "formula_map": str(formula_map_path.resolve()),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "inserted_chapter_count": len(inserted),
        "inserted_formula_count": sum(int(item["inserted_formula_count"]) for item in inserted),
        "inserted": inserted,
        "verdict": "pass",
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert one formula paragraph into a borderless table with right-side number.")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--para-id")
    parser.add_argument("--number")
    parser.add_argument("--output")
    parser.add_argument("--replace-text-formulas", action="store_true")
    parser.add_argument("--replace-inline-text-formulas", action="store_true")
    parser.add_argument("--replace-formula-tables-by-number", action="store_true")
    parser.add_argument("--replace-plain-numbered-formulas", action="store_true")
    parser.add_argument("--insert-formula-batch", action="store_true")
    parser.add_argument("--replace-formula-batch", action="store_true")
    parser.add_argument("--remove-raw-token-formula-batch", action="store_true")
    parser.add_argument("--repair-raw-sqrt-formula-tokens", action="store_true")
    parser.add_argument("--insert-source-formula-blocks", action="store_true")
    parser.add_argument("--source-docx")
    parser.add_argument("--repair-number-cells", action="store_true")
    parser.add_argument("--repair-standalone-number-labels", action="store_true")
    parser.add_argument("--repair-formula-paragraph-tabs", action="store_true")
    parser.add_argument("--convert-standalone-formulas-to-tables", action="store_true")
    parser.add_argument("--number-standalone-formulas-to-tables", action="store_true")
    parser.add_argument("--strip-non-number-table-formula-labels", action="store_true")
    parser.add_argument("--remove-duplicate-standalone-formula-labels", action="store_true")
    parser.add_argument("--number-format", choices=["preserve", "hyphen"], default="hyphen")
    parser.add_argument("--formula-map")
    parser.add_argument("--report")
    args = parser.parse_args()

    if args.repair_number_cells:
        if not args.output:
            parser.error("--repair-number-cells requires --output")
        report = repair_formula_number_cells(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
            number_format=args.number_format,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.repair_standalone_number_labels:
        if not args.output:
            parser.error("--repair-standalone-number-labels requires --output")
        report = repair_standalone_formula_number_labels(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
            number_format=args.number_format,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.repair_formula_paragraph_tabs:
        if not args.output:
            parser.error("--repair-formula-paragraph-tabs requires --output")
        report = repair_formula_paragraph_tabs(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.convert_standalone_formulas_to_tables:
        if not args.output:
            parser.error("--convert-standalone-formulas-to-tables requires --output")
        report = convert_standalone_formula_paragraphs_to_tables(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
            number_format=args.number_format,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.number_standalone_formulas_to_tables:
        if not args.output:
            parser.error("--number-standalone-formulas-to-tables requires --output")
        report = number_standalone_formula_paragraphs_to_tables(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
            number_format=args.number_format,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.strip_non_number_table_formula_labels:
        if not args.output:
            parser.error("--strip-non-number-table-formula-labels requires --output")
        report = strip_formula_labels_from_non_number_table_cells(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.remove_duplicate_standalone_formula_labels:
        if not args.output:
            parser.error("--remove-duplicate-standalone-formula-labels requires --output")
        report = remove_duplicate_standalone_formula_labels(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.replace_inline_text_formulas:
        if not args.formula_map or not args.output:
            parser.error("--replace-inline-text-formulas requires --formula-map and --output")
        report = replace_inline_text_formulas(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.replace_formula_tables_by_number:
        if not args.formula_map or not args.output:
            parser.error("--replace-formula-tables-by-number requires --formula-map and --output")
        report = replace_formula_tables_by_number(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.replace_plain_numbered_formulas:
        if not args.formula_map or not args.output:
            parser.error("--replace-plain-numbered-formulas requires --formula-map and --output")
        report = replace_plain_numbered_formulas(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.insert_formula_batch:
        if not args.formula_map or not args.output:
            parser.error("--insert-formula-batch requires --formula-map and --output")
        report = insert_formula_batch(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.remove_raw_token_formula_batch:
        if not args.formula_map or not args.output:
            parser.error("--remove-raw-token-formula-batch requires --formula-map and --output")
        report = remove_raw_token_formula_batch(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.replace_formula_batch:
        if not args.formula_map or not args.output:
            parser.error("--replace-formula-batch requires --formula-map and --output")
        report = replace_formula_batch(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.repair_raw_sqrt_formula_tokens:
        if not args.output:
            parser.error("--repair-raw-sqrt-formula-tokens requires --output")
        report = repair_raw_sqrt_formula_tokens(
            docx_path=Path(args.docx),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.insert_source_formula_blocks:
        if not args.source_docx or not args.formula_map or not args.output:
            parser.error("--insert-source-formula-blocks requires --source-docx, --formula-map, and --output")
        report = insert_source_formula_blocks(
            docx_path=Path(args.docx),
            source_docx_path=Path(args.source_docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if args.replace_text_formulas:
        if not args.formula_map or not args.output:
            parser.error("--replace-text-formulas requires --formula-map and --output")
        report = replace_text_formulas(
            docx_path=Path(args.docx),
            formula_map_path=Path(args.formula_map),
            output_path=Path(args.output),
            report_path=Path(args.report) if args.report else None,
        )
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0

    if not args.para_id or not args.number:
        parser.error("--para-id and --number are required unless --replace-text-formulas is used")

    out_path = patch_formula(
        docx_path=Path(args.docx),
        para_id=args.para_id,
        number_text=args.number,
        output_path=Path(args.output) if args.output else None,
    )
    print(f"patched_docx={out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
