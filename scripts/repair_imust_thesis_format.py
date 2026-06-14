#!/usr/bin/env python3
"""Repair Inner Mongolia University of Science and Technology thesis format drift.

The repair is intentionally scoped to visible thesis-format surfaces repeatedly
reported by users for IMUST graduation-design manuscripts:
cover/front matter, abstracts, TOC cache, body heading levels, references,
appendix/acknowledgement titles, and running headers. It does not rewrite body
content.
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

try:
    from audit_docx_list_pollution import audit_docx_list_pollution
except Exception:  # pragma: no cover
    audit_docx_list_pollution = None  # type: ignore[assignment]


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
NS = {"w": W_NS, "r": R_NS}
XML_SPACE = f"{{{XML_NS}}}space"

BULLET_PREFIX_RE = re.compile(r"^[\s\u3000]*[\u2022\u25cf\u25e6\u2219\u2043\uf0b7][\s\u3000]*")
HEADER_TEXT = "内蒙古科技大学毕业设计说明书（毕业论文）"


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def attr(element: ET.Element | None, local: str = "val") -> str:
    if element is None:
        return ""
    return element.get(qn(local), "")


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def visible_text_with_tabs(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for node in paragraph.iter():
        if node.tag == qn("t"):
            parts.append(node.text or "")
        elif node.tag == qn("tab"):
            parts.append("\t")
    return "".join(parts)


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def body_children(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    return list(body) if body is not None else []


def body_paragraphs(root: ET.Element) -> list[ET.Element]:
    return [child for child in body_children(root) if child.tag == qn("p")]


def paragraph_style_id(paragraph: ET.Element) -> str:
    return attr(paragraph.find("./w:pPr/w:pStyle", NS))


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def clear_ppr_children(ppr: ET.Element, names: set[str]) -> int:
    removed = 0
    for child in list(ppr):
        if child.tag.rsplit("}", 1)[-1] in names:
            ppr.remove(child)
            removed += 1
    return removed


def clear_paragraph_numbering(paragraph: ET.Element) -> int:
    return clear_ppr_children(ensure_ppr(paragraph), {"numPr", "outlineLvl"})


def set_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = ensure_ppr(paragraph)
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        style = ET.Element(qn("pStyle"))
        ppr.insert(0, style)
    style.set(qn("val"), style_id)


def set_child_val(parent: ET.Element, tag: str, value: str) -> ET.Element:
    node = parent.find(f"./w:{tag}", NS)
    if node is None:
        node = ET.SubElement(parent, qn(tag))
    node.set(qn("val"), value)
    return node


def set_jc(ppr: ET.Element, value: str | None) -> None:
    for node in list(ppr.findall("./w:jc", NS)):
        ppr.remove(node)
    if value:
        node = ET.SubElement(ppr, qn("jc"))
        node.set(qn("val"), value)


def set_spacing(ppr: ET.Element, *, before: str = "0", after: str = "0", line: str = "360", line_rule: str = "auto") -> None:
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), before)
    spacing.set(qn("after"), after)
    spacing.set(qn("line"), line)
    spacing.set(qn("lineRule"), line_rule)


def set_indent(
    ppr: ET.Element,
    *,
    first_line: str | None = None,
    first_line_chars: str | None = None,
    left: str | None = None,
    hanging: str | None = None,
) -> None:
    for node in list(ppr.findall("./w:ind", NS)):
        ppr.remove(node)
    if first_line is None and first_line_chars is None and left is None and hanging is None:
        return
    ind = ET.SubElement(ppr, qn("ind"))
    if left is not None:
        ind.set(qn("left"), left)
    if hanging is not None:
        ind.set(qn("hanging"), hanging)
    if first_line is not None:
        ind.set(qn("firstLine"), first_line)
    if first_line_chars is not None:
        ind.set(qn("firstLineChars"), first_line_chars)


def set_tabs(ppr: ET.Element, *, right_pos: str = "9350", level: int = 1) -> None:
    for node in list(ppr.findall("./w:tabs", NS)):
        ppr.remove(node)
    tabs = ET.SubElement(ppr, qn("tabs"))
    tab = ET.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    tab.set(qn("pos"), right_pos)
    if level > 1:
        set_indent(ppr, left=str(420 * (level - 1)))


def ensure_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        run.insert(0, rpr)
    return rpr


def set_rpr(
    rpr: ET.Element,
    *,
    size: str,
    east_asia: str = "宋体",
    latin: str = "Times New Roman",
    bold: bool | None = None,
    italic: bool = False,
    underline: bool = False,
) -> None:
    for tag in ("rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "u", "color", "highlight"):
        for node in list(rpr.findall(f"./w:{tag}", NS)):
            rpr.remove(node)
    fonts = ET.SubElement(rpr, qn("rFonts"))
    fonts.set(qn("eastAsia"), east_asia)
    fonts.set(qn("ascii"), latin)
    fonts.set(qn("hAnsi"), latin)
    fonts.set(qn("cs"), latin)
    if bold is True:
        ET.SubElement(rpr, qn("b"))
        ET.SubElement(rpr, qn("bCs"))
    elif bold is False:
        b = ET.SubElement(rpr, qn("b"))
        b.set(qn("val"), "0")
        bcs = ET.SubElement(rpr, qn("bCs"))
        bcs.set(qn("val"), "0")
    if italic:
        ET.SubElement(rpr, qn("i"))
        ET.SubElement(rpr, qn("iCs"))
    if underline:
        u = ET.SubElement(rpr, qn("u"))
        u.set(qn("val"), "single")
    color = ET.SubElement(rpr, qn("color"))
    color.set(qn("val"), "000000")
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), size)
    szcs = ET.SubElement(rpr, qn("szCs"))
    szcs.set(qn("val"), size)


def format_runs(paragraph: ET.Element, *, size: str, bold: bool | None, east_asia: str = "宋体", latin: str = "Times New Roman") -> int:
    changed = 0
    for run in paragraph.findall("./w:r", NS):
        if run.find(".//w:fldChar", NS) is not None or run.find(".//w:instrText", NS) is not None:
            continue
        if not text_of(run) and run.find("./w:tab", NS) is None:
            continue
        set_rpr(ensure_rpr(run), size=size, east_asia=east_asia, latin=latin, bold=bold)
        changed += 1
    return changed


def append_run(paragraph: ET.Element, text: str = "", *, size: str = "24", bold: bool | None = None, east_asia: str = "宋体", latin: str = "Times New Roman", tab: bool = False) -> ET.Element:
    run = ET.Element(qn("r"))
    set_rpr(ensure_rpr(run), size=size, east_asia=east_asia, latin=latin, bold=bold)
    if tab:
        run.append(ET.Element(qn("tab")))
    else:
        node = ET.SubElement(run, qn("t"))
        if text[:1].isspace() or text[-1:].isspace():
            node.set(XML_SPACE, "preserve")
        node.text = text
    paragraph.append(run)
    return run


def set_paragraph_text(paragraph: ET.Element, text: str, *, size: str, bold: bool | None, east_asia: str = "宋体", latin: str = "Times New Roman") -> None:
    ppr = copy.deepcopy(paragraph.find("./w:pPr", NS))
    paragraph[:] = []
    if ppr is not None:
        paragraph.append(ppr)
    append_run(paragraph, text, size=size, bold=bold, east_asia=east_asia, latin=latin)


def strip_visible_bullet(text: str) -> str:
    return BULLET_PREFIX_RE.sub("", text or "", count=1)


def is_zh_abstract(text: str) -> bool:
    return compact(strip_visible_bullet(text)) in {"摘要", "摘要:"}


def is_en_abstract(text: str) -> bool:
    return compact(strip_visible_bullet(text)).startswith("abstract")


def is_toc_title(text: str) -> bool:
    return compact(strip_visible_bullet(text)) in {"目录", "contents", "tableofcontents"}


def is_reference_title(text: str) -> bool:
    return compact(strip_visible_bullet(text)) in {"参考文献", "references", "bibliography"}


def is_ack_title(text: str) -> bool:
    return compact(strip_visible_bullet(text)) in {"致谢", "谢辞", "acknowledgements", "acknowledgments"}


def is_appendix_title(text: str) -> bool:
    return compact(strip_visible_bullet(text)) in {"附录", "appendix"}


def heading_level(text: str) -> int | None:
    value = strip_visible_bullet(text).strip()
    if re.match(r"^第\s*[0-9一二三四五六七八九十]+\s*章(?:\s+\S.*)?$", value):
        return 1
    if re.match(r"^[1-9]\d*\.\d+\.\d+(?:\s+\S.*)?$", value):
        return 3
    if re.match(r"^[1-9]\d*\.\d+(?:\s+\S.*)?$", value):
        return 2
    return None


def reference_entry(text: str) -> bool:
    value = strip_visible_bullet(text).strip()
    return bool(re.match(r"^(?:\[[0-9]{1,3}\]|[0-9]{1,3}[.、])\s*", value))


def paragraph_has_field_end(paragraph: ET.Element) -> bool:
    return any(attr(node, "fldCharType") == "end" for node in paragraph.findall(".//w:fldChar", NS))


def load_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except (KeyError, ET.ParseError):
        return None


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)


def style_by_id(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for style in styles_root.findall("./w:style", NS):
        if attr(style, "styleId") == style_id:
            return style
    return None


def ensure_style(styles_root: ET.Element, style_id: str, name: str) -> ET.Element:
    style = style_by_id(styles_root, style_id)
    if style is None:
        style = ET.SubElement(styles_root, qn("style"))
        style.set(qn("type"), "paragraph")
        style.set(qn("styleId"), style_id)
        name_node = ET.SubElement(style, qn("name"))
        name_node.set(qn("val"), name)
    else:
        name_node = style.find("./w:name", NS)
        if name_node is None:
            name_node = ET.SubElement(style, qn("name"))
        name_node.set(qn("val"), name)
    return style


def define_paragraph_style(
    styles_root: ET.Element,
    style_id: str,
    name: str,
    *,
    size: str,
    bold: bool,
    jc: str | None,
    outline: int | None,
    first_line: bool = False,
    line: str = "360",
    before: str = "0",
    after: str = "0",
) -> None:
    style = ensure_style(styles_root, style_id, name)
    for child_name in ("pPr", "rPr"):
        for child in list(style.findall(f"./w:{child_name}", NS)):
            style.remove(child)
    ppr = ET.SubElement(style, qn("pPr"))
    if outline is not None:
        outline_node = ET.SubElement(ppr, qn("outlineLvl"))
        outline_node.set(qn("val"), str(outline))
    set_jc(ppr, jc)
    set_spacing(ppr, before=before, after=after, line=line, line_rule="auto")
    if first_line:
        set_indent(ppr, first_line="480", first_line_chars="200")
    rpr = ET.SubElement(style, qn("rPr"))
    set_rpr(rpr, size=size, bold=bold)


def numbering_models(numbering_root: ET.Element | None) -> dict[str, dict[str, str]]:
    if numbering_root is None:
        return {}
    abstract_by_id = {
        attr(item, "abstractNumId"): item
        for item in numbering_root.findall("./w:abstractNum", NS)
        if attr(item, "abstractNumId")
    }
    result: dict[str, dict[str, str]] = {}
    for num in numbering_root.findall("./w:num", NS):
        num_id = attr(num, "numId")
        abstract_id = attr(num.find("./w:abstractNumId", NS))
        abstract = abstract_by_id.get(abstract_id)
        if not num_id or abstract is None:
            continue
        lvl = abstract.find("./w:lvl", NS)
        result[num_id] = {
            "numFmt": attr(lvl.find("./w:numFmt", NS)) if lvl is not None else "",
            "lvlText": attr(lvl.find("./w:lvlText", NS)) if lvl is not None else "",
        }
    return result


def model_is_bullet(model: dict[str, str]) -> bool:
    return (model.get("numFmt") or "").lower() == "bullet" or "\uf0b7" in (model.get("lvlText") or "")


def strip_bullet_numbering_from_styles(styles_root: ET.Element, numbering_root: ET.Element | None) -> list[dict[str, object]]:
    models = numbering_models(numbering_root)
    changes: list[dict[str, object]] = []
    for style in styles_root.findall("./w:style", NS):
        style_id = attr(style, "styleId")
        ppr = style.find("./w:pPr", NS)
        if ppr is None:
            continue
        for num_pr in list(ppr.findall("./w:numPr", NS)):
            num_id = attr(num_pr.find("./w:numId", NS))
            model = models.get(num_id, {})
            if style_id in {"Heading1", "Heading2", "Heading3", "TOC1", "TOC2", "TOC3", "2"} or model_is_bullet(model):
                ppr.remove(num_pr)
                changes.append({"style_id": style_id, "removed_numId": num_id, "model": model})
        for outline in list(ppr.findall("./w:outlineLvl", NS)):
            if style_id in {"2"}:
                ppr.remove(outline)
    return changes


def apply_ppr_baseline(
    paragraph: ET.Element,
    *,
    style_id: str,
    jc: str | None,
    outline: int | None,
    first_line: bool = False,
    left: str | None = None,
    hanging: str | None = None,
    line: str = "360",
    before: str = "0",
    after: str = "0",
    keep_page_break_before: bool = False,
) -> None:
    old_ppr = paragraph.find("./w:pPr", NS)
    keep_page_break = old_ppr.find("./w:pageBreakBefore", NS) is not None if old_ppr is not None else False
    sect_pr = copy.deepcopy(old_ppr.find("./w:sectPr", NS)) if old_ppr is not None and old_ppr.find("./w:sectPr", NS) is not None else None
    ppr = ET.Element(qn("pPr"))
    p_style = ET.SubElement(ppr, qn("pStyle"))
    p_style.set(qn("val"), style_id)
    if outline is not None:
        outline_node = ET.SubElement(ppr, qn("outlineLvl"))
        outline_node.set(qn("val"), str(outline))
    set_jc(ppr, jc)
    set_spacing(ppr, before=before, after=after, line=line, line_rule="auto")
    if first_line:
        set_indent(ppr, first_line="480", first_line_chars="200")
    elif left is not None or hanging is not None:
        set_indent(ppr, left=left, hanging=hanging)
    if keep_page_break_before and keep_page_break:
        ET.SubElement(ppr, qn("pageBreakBefore"))
    if sect_pr is not None:
        ppr.append(sect_pr)
    old = paragraph.find("./w:pPr", NS)
    if old is not None:
        paragraph.remove(old)
    paragraph.insert(0, ppr)


def style_front_matter(paragraphs: list[ET.Element], *, cover_values: dict[str, str]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    cover_replacements = {
        "学生姓名": cover_values.get("student_name", "待填写"),
        "学    号": cover_values.get("student_id", "待填写"),
        "学 号": cover_values.get("student_id", "待填写"),
        "班    级": cover_values.get("class_name", "待填写"),
        "班 级": cover_values.get("class_name", "待填写"),
        "指导教师": cover_values.get("advisor", "待填写"),
    }
    for index, paragraph in enumerate(paragraphs[:9], start=1):
        text = text_of(paragraph).strip()
        if not text:
            continue
        apply_ppr_baseline(paragraph, style_id="IMUSTCoverLine", jc="center", outline=None, line="360")
        size = "28"
        bold = False
        if index == 1:
            size = "32"
            bold = True
        elif "题" in text and "目" in text:
            size = "30"
            bold = True
        new_text = text
        for label, value in cover_replacements.items():
            if compact(new_text).startswith(compact(label + "：")) or compact(new_text).startswith(compact(label + ":")):
                normalized_label = label.replace(" ", "")
                if normalized_label == "学号":
                    new_text = f"学    号：{value}"
                elif normalized_label == "班级":
                    new_text = f"班    级：{value}"
                else:
                    new_text = re.sub(r"[:：].*$", f"：{value}", new_text)
        set_paragraph_text(paragraph, new_text, size=size, bold=bold)
        changes.append({"paragraph_index": index, "surface": "cover_style", "text": new_text})
    # Use the existing blank spacer before the Chinese abstract as a standard cover date line.
    if len(paragraphs) >= 9 and not text_of(paragraphs[8]).strip():
        apply_ppr_baseline(paragraphs[8], style_id="IMUSTCoverLine", jc="center", outline=None, line="360")
        set_paragraph_text(paragraphs[8], cover_values.get("date", "2026年5月"), size="28", bold=False)
        changes.append({"paragraph_index": 9, "surface": "cover_date", "text": cover_values.get("date", "2026年5月")})
    return changes


def style_abstracts(paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    for index, paragraph in enumerate(paragraphs, start=1):
        text = strip_visible_bullet(text_of(paragraph).strip())
        if not text:
            continue
        if is_zh_abstract(text):
            apply_ppr_baseline(paragraph, style_id="IMUSTFrontTitle", jc="center", outline=0, line="360", after="0")
            set_paragraph_text(paragraph, "摘  要", size="30", bold=True)
            changes.append({"paragraph_index": index, "surface": "zh_abstract_title"})
        elif is_en_abstract(text):
            apply_ppr_baseline(paragraph, style_id="IMUSTFrontTitle", jc="center", outline=0, line="360", after="0")
            set_paragraph_text(paragraph, "Abstract", size="30", bold=True, east_asia="Times New Roman", latin="Times New Roman")
            changes.append({"paragraph_index": index, "surface": "en_abstract_title"})
        elif compact(text).startswith("关键词"):
            apply_ppr_baseline(paragraph, style_id="IMUSTKeyword", jc=None, outline=None, first_line=True, line="360")
            rewrite_keyword(paragraph, chinese=True)
            changes.append({"paragraph_index": index, "surface": "zh_keyword_line"})
        elif compact(text).startswith("keywords") or compact(text).startswith("key words"):
            apply_ppr_baseline(paragraph, style_id="IMUSTKeyword", jc=None, outline=None, first_line=True, line="360")
            rewrite_keyword(paragraph, chinese=False)
            changes.append({"paragraph_index": index, "surface": "en_keyword_line"})
        elif index > 1 and (is_zh_abstract(text_of(paragraphs[index - 2])) or is_en_abstract(text_of(paragraphs[index - 2]))):
            apply_ppr_baseline(paragraph, style_id="IMUSTAbstractBody", jc="both", outline=None, first_line=True, line="360")
            format_runs(paragraph, size="24", bold=False)
            changes.append({"paragraph_index": index, "surface": "abstract_body"})
    return changes


def rewrite_keyword(paragraph: ET.Element, *, chinese: bool) -> None:
    text = strip_visible_bullet(text_of(paragraph).strip())
    if chinese:
        match = re.match(r"^\s*关键词\s*[:：]\s*(.*)$", text)
        label, rest = "关键词：", match.group(1).strip() if match else text
    else:
        match = re.match(r"^\s*(?:key\s+words|keywords)\s*[:：]\s*(.*)$", text, re.I)
        label, rest = "Key words:", match.group(1).strip() if match else text
    ppr = copy.deepcopy(paragraph.find("./w:pPr", NS))
    paragraph[:] = []
    if ppr is not None:
        paragraph.append(ppr)
    append_run(paragraph, label, size="24", bold=True, east_asia="黑体" if chinese else "Times New Roman")
    append_run(paragraph, (" " if not chinese else "") + rest, size="24", bold=False)


def split_toc_label_page(paragraph: ET.Element) -> tuple[str, str]:
    text = visible_text_with_tabs(paragraph).strip()
    if "\t" in text:
        left, right = text.rsplit("\t", 1)
        return left.strip(), right.strip()
    match = re.match(r"^(.*?)([ivxlcdmIVXLCDM]+|[0-9]+)\s*$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return text, ""


def toc_entry_level(label: str, style_id: str) -> int:
    if style_id == "TOC2" or re.match(r"^[1-9]\d*\.\d+\s+", label):
        return 2
    if style_id == "TOC3" or re.match(r"^[1-9]\d*\.\d+\.\d+\s+", label):
        return 3
    return 1


def rebuild_toc_entry(paragraph: ET.Element, *, label: str, page: str, level: int, field_prefix: list[ET.Element], field_suffix: list[ET.Element]) -> None:
    paragraph[:] = []
    ppr = ET.Element(qn("pPr"))
    p_style = ET.SubElement(ppr, qn("pStyle"))
    p_style.set(qn("val"), f"TOC{level}")
    set_spacing(ppr, before="0", after="0", line="300", line_rule="auto")
    set_tabs(ppr, level=level)
    paragraph.append(ppr)
    paragraph.extend(copy.deepcopy(item) for item in field_prefix)
    append_run(paragraph, label, size="24", bold=False)
    append_run(paragraph, tab=True, size="24", bold=False)
    append_run(paragraph, page, size="24", bold=False)
    paragraph.extend(copy.deepcopy(item) for item in field_suffix)


def style_toc(paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_index = next((i for i, p in enumerate(paragraphs) if is_toc_title(text_of(p))), None)
    if toc_index is None:
        return changes
    toc_end = toc_index
    for scan in range(toc_index + 1, len(paragraphs)):
        if paragraph_has_field_end(paragraphs[scan]):
            toc_end = scan
            break
        if paragraphs[scan].find("./w:pPr/w:sectPr", NS) is not None:
            toc_end = scan - 1
            break
    if toc_end <= toc_index:
        return changes
    title = paragraphs[toc_index]
    apply_ppr_baseline(title, style_id="IMUSTTocTitle", jc="center", outline=None, line="360", after="0")
    set_paragraph_text(title, "目 录", size="30", bold=True)
    changes.append({"paragraph_index": toc_index + 1, "surface": "toc_title"})
    for index in range(toc_index + 1, toc_end + 1):
        paragraph = paragraphs[index]
        text = text_of(paragraph).strip()
        if not text and paragraph.find(".//w:fldChar", NS) is None:
            continue
        label, page = split_toc_label_page(paragraph)
        label = strip_visible_bullet(label)
        if not page and compact(label) in {"abstract", "abstractii"}:
            label = "Abstract"
            page = "II"
        if compact(label) in {"目录", "contents", "tableofcontents"}:
            # Keep user-visible TOC clean if an old field cache contained itself.
            label = ""
            page = ""
        if not label:
            continue
        level = toc_entry_level(label, paragraph_style_id(paragraph))
        field_prefix: list[ET.Element] = []
        field_suffix: list[ET.Element] = []
        for child in paragraph:
            fld = child.find("./w:fldChar", NS)
            if fld is not None and attr(fld, "fldCharType") in {"begin", "separate"}:
                if attr(fld, "fldCharType") == "begin":
                    fld.set(qn("fldLock"), "true")
                    fld.attrib.pop(qn("dirty"), None)
                field_prefix.append(child)
                continue
            if child.find("./w:instrText", NS) is not None:
                field_prefix.append(child)
                continue
            if fld is not None and attr(fld, "fldCharType") == "end":
                field_suffix.append(child)
        rebuild_toc_entry(paragraph, label=label, page=page, level=level, field_prefix=field_prefix, field_suffix=field_suffix)
        changes.append({"paragraph_index": index + 1, "surface": "toc_entries", "level": level, "label": label[:80], "page": page})
    return changes


def style_headings_and_tail(paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    toc_index = next((i for i, p in enumerate(paragraphs) if is_toc_title(text_of(p))), None)
    toc_end = toc_index
    if toc_index is not None:
        for scan in range(toc_index + 1, len(paragraphs)):
            if paragraph_has_field_end(paragraphs[scan]):
                toc_end = scan
                break
            if paragraphs[scan].find("./w:pPr/w:sectPr", NS) is not None:
                toc_end = scan - 1
                break
    for index, paragraph in enumerate(paragraphs, start=1):
        zero_index = index - 1
        if toc_index is not None and toc_end is not None and toc_index <= zero_index <= toc_end:
            continue
        text = strip_visible_bullet(text_of(paragraph).strip())
        if not text:
            continue
        level = heading_level(text)
        if level == 1:
            apply_ppr_baseline(paragraph, style_id="Heading1", jc="center", outline=0, line="360", after="0", keep_page_break_before=True)
            format_runs(paragraph, size="30", bold=True)
            changes.append({"paragraph_index": index, "surface": "body_heading_levels", "level": 1})
        elif level == 2:
            apply_ppr_baseline(paragraph, style_id="Heading2", jc="left", outline=1, line="360", after="0")
            format_runs(paragraph, size="28", bold=True)
            changes.append({"paragraph_index": index, "surface": "body_heading_levels", "level": 2})
        elif level == 3:
            apply_ppr_baseline(paragraph, style_id="Heading3", jc="left", outline=2, line="360", after="0")
            format_runs(paragraph, size="24", bold=True)
            changes.append({"paragraph_index": index, "surface": "body_heading_levels", "level": 3})
        elif is_reference_title(text) or is_appendix_title(text) or is_ack_title(text):
            apply_ppr_baseline(paragraph, style_id="Heading1", jc="center", outline=0, line="360", after="0", keep_page_break_before=True)
            set_paragraph_text(paragraph, text, size="30", bold=True)
            changes.append({"paragraph_index": index, "surface": "tail_title", "text": text})
    return changes


def style_references(paragraphs: list[ET.Element]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    ref_index = next((i for i, p in enumerate(paragraphs) if is_reference_title(text_of(p))), None)
    if ref_index is None:
        return changes
    for index in range(ref_index + 1, len(paragraphs)):
        paragraph = paragraphs[index]
        text = strip_visible_bullet(text_of(paragraph).strip())
        if not text:
            continue
        if is_ack_title(text) or is_appendix_title(text) or heading_level(text) == 1:
            break
        if not reference_entry(text):
            continue
        clear_paragraph_numbering(paragraph)
        apply_ppr_baseline(paragraph, style_id="IMUSTReferencesEntry", jc=None, outline=None, left="420", hanging="420", line="300")
        set_paragraph_text(paragraph, text, size="21", bold=False)
        changes.append({"paragraph_index": index + 1, "surface": "references_entries", "text_prefix": text[:80]})
    return changes


def repair_headers(parts: dict[str, bytes]) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    for name, payload in list(parts.items()):
        if not (name.startswith("word/header") and name.endswith(".xml")):
            continue
        try:
            root = ET.fromstring(payload)
        except ET.ParseError:
            continue
        paragraphs = root.findall(".//w:p", NS)
        if not paragraphs:
            continue
        first = paragraphs[0]
        apply_ppr_baseline(first, style_id="Header", jc="center", outline=None, line="240")
        set_paragraph_text(first, HEADER_TEXT, size="21", bold=False)
        for extra in paragraphs[1:]:
            if text_of(extra).strip():
                set_paragraph_text(extra, "", size="21", bold=False)
        parts[name] = xml_bytes(root)
        changes.append({"part": name, "text": HEADER_TEXT})
    return changes


def repair_settings(parts: dict[str, bytes]) -> dict[str, object]:
    name = "word/settings.xml"
    if name not in parts:
        return {"status": "missing"}
    try:
        root = ET.fromstring(parts[name])
    except ET.ParseError:
        return {"status": "parse-error"}
    removed = 0
    for node in list(root.findall("./w:updateFields", NS)):
        root.remove(node)
        removed += 1
    update = ET.SubElement(root, qn("updateFields"))
    update.set(qn("val"), "false")
    parts[name] = xml_bytes(root)
    return {"status": "updated", "removed_updateFields": removed, "set_updateFields": "false"}


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


def repair_docx(
    source_docx: Path,
    output_docx: Path,
    *,
    cover_values: dict[str, str],
) -> dict[str, object]:
    if source_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy, not the source path")
    with zipfile.ZipFile(source_docx) as zf:
        parts = {item.filename: zf.read(item.filename) for item in zf.infolist()}
    document_root = ET.fromstring(parts["word/document.xml"])
    styles_root = ET.fromstring(parts["word/styles.xml"])
    numbering_root = ET.fromstring(parts["word/numbering.xml"]) if "word/numbering.xml" in parts else None
    style_numbering_changes = strip_bullet_numbering_from_styles(styles_root, numbering_root)
    define_paragraph_style(styles_root, "Heading1", "heading 1", size="30", bold=True, jc="center", outline=0, after="0")
    define_paragraph_style(styles_root, "Heading2", "heading 2", size="28", bold=True, jc="left", outline=1, after="0")
    define_paragraph_style(styles_root, "Heading3", "heading 3", size="24", bold=True, jc="left", outline=2, after="0")
    define_paragraph_style(styles_root, "IMUSTFrontTitle", "IMUST front matter title", size="30", bold=True, jc="center", outline=0, after="0")
    define_paragraph_style(styles_root, "IMUSTTocTitle", "IMUST TOC title", size="30", bold=True, jc="center", outline=None, after="0")
    define_paragraph_style(styles_root, "IMUSTAbstractBody", "IMUST abstract body", size="24", bold=False, jc="both", outline=None, first_line=True)
    define_paragraph_style(styles_root, "IMUSTKeyword", "IMUST keyword line", size="24", bold=False, jc=None, outline=None, first_line=True)
    define_paragraph_style(styles_root, "IMUSTReferencesEntry", "IMUST references entry", size="21", bold=False, jc=None, outline=None, line="300")
    define_paragraph_style(styles_root, "IMUSTCoverLine", "IMUST cover line", size="28", bold=False, jc="center", outline=None)
    for style_id in ("TOC1", "TOC2", "TOC3"):
        define_paragraph_style(styles_root, style_id, style_id.lower(), size="24", bold=False, jc=None, outline=None, line="300")
    paragraphs = body_paragraphs(document_root)
    changes = {
        "cover": style_front_matter(paragraphs, cover_values=cover_values),
        "abstracts": style_abstracts(paragraphs),
        "toc": style_toc(paragraphs),
        "headings_and_tail": style_headings_and_tail(paragraphs),
        "references": style_references(paragraphs),
    }
    header_changes = repair_headers(parts)
    settings_change = repair_settings(parts)
    parts["word/document.xml"] = xml_bytes(document_root)
    parts["word/styles.xml"] = xml_bytes(styles_root)
    write_docx_with_parts(
        source_docx,
        output_docx,
        {name: payload for name, payload in parts.items() if name.startswith("word/header") or name in {"word/document.xml", "word/styles.xml", "word/settings.xml"}},
    )
    list_audit = audit_docx_list_pollution(output_docx) if audit_docx_list_pollution is not None else {"passed": False, "reason": "audit unavailable"}
    return {
        "schema": "graduation-project-builder.imust-thesis-format-repair.v1",
        "source_docx": str(source_docx),
        "source_sha256": sha256_file(source_docx),
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "school_rule_source": "内蒙古科技大学毕业设计（论文）撰写与装订规范: A4 margins; header text; heading sizes; abstract/TOC/reference title as level-1 title; body小四/1.5倍/首行缩进2字符; references Arabic-number order.",
        "style_numbering_changes": style_numbering_changes,
        "surface_changes": changes,
        "header_changes": header_changes,
        "settings_change": settings_change,
        "cover_placeholder_values_used": {
            key: value
            for key, value in cover_values.items()
            if value in {"待填写", ""}
        },
        "post_repair_list_pollution_audit": list_audit,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair IMUST thesis front-matter, heading, TOC, reference, and header format drift.")
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report-json", required=True, type=Path)
    parser.add_argument("--student-name", default="待填写")
    parser.add_argument("--student-id", default="待填写")
    parser.add_argument("--class-name", default="待填写")
    parser.add_argument("--advisor", default="待填写")
    parser.add_argument("--date", default="2026年5月")
    args = parser.parse_args()
    report = repair_docx(
        args.input_docx.resolve(),
        args.output_docx.resolve(),
        cover_values={
            "student_name": args.student_name,
            "student_id": args.student_id,
            "class_name": args.class_name,
            "advisor": args.advisor,
            "date": args.date,
        },
    )
    args.report_json.parent.mkdir(parents=True, exist_ok=True)
    args.report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    passed = report.get("post_repair_list_pollution_audit", {}).get("passed") is True
    print(json.dumps({"passed": passed, "output_docx": str(args.output_docx), "output_sha256": report["output_sha256"]}, ensure_ascii=False))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
