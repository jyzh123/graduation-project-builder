#!/usr/bin/env python3
"""Bounded DOCX surface-format repair for thesis recovery passes.

This helper intentionally owns only narrow, template-donor repairs:
- abstract body replacement from a JSON plan while preserving donor metrics
- body prose paragraph normalization from a real template body donor
- uniquely matched body prose text replacement from a JSON plan for pagination
  orphan cleanup, only when the target paragraph has no fields, drawings, or
  citation markers
- removal of empty paragraphs immediately before explicitly named headings
- image-holder paragraph safety cleanup

It may not independently resize existing image drawings. Display-size changes
are drawing mutations and must be owned by a transaction/figure-manifest wrapper.

It does not rebuild the document, refresh fields, rewrite TOC entries, or change
citations/bibliography numbering.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import html
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from audit_docx_body_style import body_paragraphs as audited_body_paragraphs


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W = f"{{{W_NS}}}"

NS = {"w": W_NS, "r": R_NS, "a": A_NS, "wp": WP_NS}
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
ASCII_ALNUM_RE = re.compile(r"[A-Za-z0-9]")
EMU_PER_CM = 360000

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("wp", WP_NS)


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
    return re.sub(r"\s+", "", text or "")


def has_mixed_cjk_ascii(text: str) -> bool:
    return bool(CJK_RE.search(text or "") and ASCII_ALNUM_RE.search(text or ""))


def is_zh_abstract_title_text(text: str) -> bool:
    compact = compact_text(text).lower()
    return compact == "摘要" or compact.startswith(("摘要(", "摘要（"))


def is_zh_abstract_title_text(text: str) -> bool:
    compact = compact_text(text).lower()
    return compact == "\u6458\u8981" or compact.startswith(("\u6458\u8981(", "\u6458\u8981\uff08", "\u6458\u8981:", "\u6458\u8981\uff1a"))


def is_en_abstract_title_text(text: str) -> bool:
    compact = compact_text(text).lower()
    return compact == "abstract" or compact.startswith(("abstract(", "abstract（", "abstract:", "abstract\uff1a"))


def abstract_inline_body_text(text: str) -> bool:
    compact = compact_text(text)
    lowered = compact.lower()
    return (
        lowered.startswith(("\u6458\u8981:", "\u6458\u8981\uff1a"))
        and len(compact) > len("\u6458\u8981\uff1a")
    ) or (
        lowered.startswith(("abstract:", "abstract\uff1a"))
        and len(compact) > len("abstract:")
    )


def body_paragraphs(document_xml: bytes) -> list[ET.Element]:
    root = ET.fromstring(document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    return [node for node in list(body) if node.tag == qn("p")]


def parse_document(document_xml: bytes) -> tuple[ET.Element, ET.Element, list[ET.Element]]:
    root = ET.fromstring(document_xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    paragraphs = [node for node in list(body) if node.tag == qn("p")]
    return root, body, paragraphs


def paragraph_property(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find("./w:pPr", NS)


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph_property(paragraph)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def replace_ppr(paragraph: ET.Element, donor_ppr: ET.Element) -> None:
    old = paragraph_property(paragraph)
    if old is not None:
        paragraph.remove(old)
    paragraph.insert(0, deepcopy(donor_ppr))


def replace_rpr(run: ET.Element, donor_rpr: ET.Element | None) -> None:
    old = run.find("./w:rPr", NS)
    if old is not None:
        run.remove(old)
    if donor_rpr is not None:
        run.insert(0, deepcopy(donor_rpr))


def remove_annotation_visual_markers(rpr: ET.Element | None) -> int:
    if rpr is None:
        return 0
    removed = 0
    for child in list(rpr):
        if child.tag == qn("highlight"):
            rpr.remove(child)
            removed += 1
            continue
        if child.tag == qn("color"):
            value = (child.get(qn("val")) or "").upper()
            if value in {"FF0000", "C00000", "0000FF", "0000CC", "BLUE", "RED"}:
                rpr.remove(child)
                removed += 1
    return removed


def ensure_run_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        run.insert(0, rpr)
    return rpr


def ensure_rpr_fonts(rpr: ET.Element) -> ET.Element:
    rfonts = rpr.find("./w:rFonts", NS)
    if rfonts is None:
        rfonts = ET.Element(qn("rFonts"))
        rpr.insert(0, rfonts)
    return rfonts


def ensure_rpr_bold(rpr: ET.Element) -> None:
    if rpr.find("./w:b", NS) is None:
        rpr.append(ET.Element(qn("b")))
    if rpr.find("./w:bCs", NS) is None:
        rpr.append(ET.Element(qn("bCs")))


def strong_label_rpr(donor_rpr: ET.Element | None) -> ET.Element:
    rpr = deepcopy(donor_rpr) if donor_rpr is not None else ET.Element(qn("rPr"))
    ensure_rpr_bold(rpr)
    return rpr


def first_present_element(*items: ET.Element | None) -> ET.Element | None:
    for item in items:
        if item is not None:
            return item
    return None


def apply_latin_font_slots(paragraph: ET.Element, *, east_asia_font: str = "\u5b8b\u4f53") -> int:
    changed = 0
    for run in paragraph.findall(".//w:r", NS):
        if has_field_or_drawing(run) or is_citation_run(run) or run_inside_hyperlink(paragraph, run):
            continue
        text = paragraph_text(run)
        if not re.search(r"[A-Za-z0-9]", text or ""):
            continue
        rpr = ensure_run_rpr(run)
        rfonts = ensure_rpr_fonts(rpr)
        rfonts.set(qn("ascii"), "Times New Roman")
        rfonts.set(qn("hAnsi"), "Times New Roman")
        rfonts.set(qn("cs"), "Times New Roman")
        if east_asia_font:
            rfonts.set(qn("eastAsia"), east_asia_font)
        changed += 1
    return changed


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


def clear_paragraph_content_preserving_review_anchors(paragraph: ET.Element) -> tuple[list[ET.Element], list[ET.Element]]:
    starts, ends = review_anchor_children(paragraph)
    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)
    for child in starts:
        paragraph.append(child)
    return starts, ends


def clone_rpr_with_latin_slots(rpr: ET.Element | None, *, east_asia_font: str = "\u5b8b\u4f53") -> ET.Element:
    cloned = deepcopy(rpr) if rpr is not None else ET.Element(qn("rPr"))
    remove_annotation_visual_markers(cloned)
    rfonts = ensure_rpr_fonts(cloned)
    rfonts.set(qn("ascii"), "Times New Roman")
    rfonts.set(qn("hAnsi"), "Times New Roman")
    rfonts.set(qn("cs"), "Times New Roman")
    if east_asia_font:
        rfonts.set(qn("eastAsia"), east_asia_font)
    return cloned


def clone_rpr_with_cjk_slot(rpr: ET.Element | None, *, east_asia_font: str = "\u5b8b\u4f53") -> ET.Element:
    cloned = deepcopy(rpr) if rpr is not None else ET.Element(qn("rPr"))
    remove_annotation_visual_markers(cloned)
    if east_asia_font:
        rfonts = ensure_rpr_fonts(cloned)
        rfonts.set(qn("eastAsia"), east_asia_font)
    return cloned


def split_mixed_script_text(text: str) -> list[str]:
    if not has_mixed_cjk_ascii(text):
        return [text]
    segments: list[str] = []
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
            segments.append("".join(current))
            current = []
        current.append(ch)
        current_mode = mode
    if current:
        segments.append("".join(current))
    return segments


def split_mixed_script_runs(paragraph: ET.Element, *, east_asia_font: str = "\u5b8b\u4f53") -> int:
    changed = 0
    children = list(paragraph)
    rebuilt: list[ET.Element] = []
    for child in children:
        if child.tag != qn("r") or has_field_or_drawing(child) or is_citation_run(child) or run_inside_hyperlink(paragraph, child):
            rebuilt.append(deepcopy(child))
            continue
        text_nodes = child.findall("./w:t", NS)
        text = "".join(node.text or "" for node in text_nodes)
        if not has_mixed_cjk_ascii(text):
            rebuilt.append(deepcopy(child))
            continue
        base_rpr = child.find("./w:rPr", NS)
        for segment in split_mixed_script_text(text):
            run = ET.Element(qn("r"))
            if ASCII_ALNUM_RE.search(segment):
                run.append(clone_rpr_with_latin_slots(base_rpr, east_asia_font=east_asia_font))
            else:
                run.append(clone_rpr_with_cjk_slot(base_rpr, east_asia_font=east_asia_font))
            t = ET.Element(qn("t"))
            if segment.startswith(" ") or segment.endswith(" "):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = segment
            run.append(t)
            rebuilt.append(run)
        changed += 1
    if changed:
        for child in children:
            paragraph.remove(child)
        for child in rebuilt:
            paragraph.append(child)
    return changed


def apply_body_mixed_script_font_slot_repairs(
    input_docx: Path, output_docx: Path, plan: dict[str, object]
) -> dict[str, object]:
    allowed_truthy = {"package_preserving_body_mixed_script_font_slots"}
    allowed_optional = {"east_asia_font", "include_front_matter_abstracts"}
    for key, value in plan.items():
        if key in allowed_optional:
            continue
        if key in allowed_truthy and bool(value):
            continue
        if value not in (False, None, "", [], {}):
            raise ValueError(
                "package-preserving body mixed-script font-slot repair may not be combined with other mutation plan fields: "
                f"{key}"
            )

    east_asia_font = str(plan.get("east_asia_font") or "\u5b8b\u4f53")
    with zipfile.ZipFile(input_docx) as zin:
        document_bytes = zin.read("word/document.xml")
        root, body, paragraphs = parse_document(document_bytes)
        toc_index = next((i for i, p in enumerate(paragraphs) if compact_text(paragraph_text(p)) == "\u76ee\u5f55"), -1)
        body_start, _ = find_first(
            paragraphs,
            lambda i, p: i > toc_index
            and not re.search(r"\d\s*$", paragraph_text(p).strip())
            and (
                is_body_heading_paragraph(p)
                or re.match(r"^\s*1\s+", paragraph_text(p)) is not None
                or re.match(r"^\s*\u7b2c\s*1\s*\u7ae0\s+\S", paragraph_text(p)) is not None
            ),
        )
        references_index = next(
            (
                i
                for i, p in enumerate(paragraphs)
                if compact_text(paragraph_text(p)) in {"\u53c2\u8003\u6587\u732e", "鍙傝€冩枃鐚?"}
            ),
            len(paragraphs),
        )
        acknowledgement_index = next(
            (
                i
                for i, p in enumerate(paragraphs)
                if compact_text(paragraph_text(p)) in {"\u81f4\u8c22", "\u8c22\u8f9e", "鑷磋阿", "璋㈣緸"}
            ),
            references_index,
        )
        body_end = min(references_index, acknowledgement_index)
        audited_targets, _audited_styles, _audited_default = audited_body_paragraphs(input_docx)
        body_children = list(body)
        target_paragraphs: list[tuple[int, ET.Element]] = []
        for record in audited_targets:
            body_child_index = int(record.get("paragraph_index") or 0) - 1
            if 0 <= body_child_index < len(body_children) and body_children[body_child_index].tag == qn("p"):
                target_paragraphs.append((body_child_index, body_children[body_child_index]))
        using_audited_body_targets = bool(target_paragraphs)
        if not target_paragraphs:
            target_paragraphs = [(index, paragraphs[index]) for index in range(body_start, body_end)]
        if bool(plan.get("include_front_matter_abstracts")):
            zh_title_index = next(
                (i for i, p in enumerate(paragraphs) if is_zh_abstract_title_text(paragraph_text(p))),
                None,
            )
            if zh_title_index is not None:
                zh_keyword_index = next(
                    (
                        i
                        for i, p in enumerate(paragraphs[zh_title_index + 1 : body_start], start=zh_title_index + 1)
                        if is_safe_zh_keyword_text(paragraph_text(p))
                    ),
                    body_start,
                )
                target_paragraphs.extend(
                    (index, paragraphs[index]) for index in range(zh_title_index + 1, zh_keyword_index)
                )

        changed: list[dict[str, object]] = []
        seen_target_ids: set[int] = set()
        for index, paragraph in sorted(target_paragraphs, key=lambda item: item[0]):
            if id(paragraph) in seen_target_ids:
                continue
            seen_target_ids.add(id(paragraph))
            text = paragraph_text(paragraph).strip()
            if not text or not has_mixed_cjk_ascii(text):
                continue
            if (not using_audited_body_targets) and (
                is_body_heading_paragraph(paragraph) or is_caption_or_reference(text) or is_code_like_paragraph_text(text)
            ):
                continue
            if paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None:
                continue
            latin_slots_changed = apply_latin_font_slots(paragraph, east_asia_font=east_asia_font)
            runs_split = split_mixed_script_runs(paragraph, east_asia_font=east_asia_font)
            if latin_slots_changed or runs_split:
                changed.append(
                    {
                        "target_paragraph_index": index,
                        "text_prefix": text[:120],
                        "latin_slots_changed": latin_slots_changed,
                        "mixed_script_runs_split": runs_split,
                    }
                )

        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    zout.writestr(item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)

    return {
        "package_preserving_body_mixed_script_font_slots": changed,
        "changed_zip_parts": ["word/document.xml"] if changed else [],
        "paragraphs_changed": len(changed),
    }


def first_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is not None:
            return deepcopy(rpr)
    return None


def first_text_run(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run):
            return run
    return None


def abstract_title_marker_rpr(paragraph: ET.Element, surface_id: str) -> ET.Element | None:
    runs = [run for run in paragraph.findall("./w:r", NS) if paragraph_text(run).strip()]
    if not runs:
        return None
    marker_run: ET.Element | None = None
    if surface_id == "zh_abstract_title":
        for run in runs:
            value = compact_text(paragraph_text(run))
            if value in {"摘", "要", "摘要"}:
                marker_run = run
                break
    elif surface_id == "en_abstract_title":
        for run in runs:
            if compact_text(paragraph_text(run)).lower() == "abstract":
                marker_run = run
                break
    selected = marker_run if marker_run is not None else runs[0]
    rpr = selected.find("./w:rPr", NS)
    return deepcopy(rpr) if rpr is not None else None


def set_paragraph_text(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> None:
    _starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    append_text_run(paragraph, text, donor_rpr)
    for child in ends:
        paragraph.append(child)


def set_paragraph_text_with_mixed_script(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> None:
    _starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    for segment in split_mixed_script_text(text):
        if ASCII_ALNUM_RE.search(segment):
            run_rpr = clone_rpr_with_latin_slots(donor_rpr)
        else:
            run_rpr = deepcopy(donor_rpr) if donor_rpr is not None else None
        append_text_run(paragraph, segment, run_rpr)
    for child in ends:
        paragraph.append(child)


def append_text_run(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> ET.Element:
    run = ET.Element(qn("r"))
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    t = ET.Element(qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    run.append(t)
    paragraph.append(run)
    return run


def text_run_rprs(paragraph: ET.Element) -> list[ET.Element | None]:
    rprs: list[ET.Element | None] = []
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run):
            rpr = run.find("./w:rPr", NS)
            rprs.append(deepcopy(rpr) if rpr is not None else None)
    return rprs


def keyword_template_role_rprs(paragraph: ET.Element) -> tuple[ET.Element | None, ET.Element | None, ET.Element | None]:
    label_rpr: ET.Element | None = None
    separator_rpr: ET.Element | None = None
    content_rpr: ET.Element | None = None
    seen_separator = False
    found_separator = False
    found_content = False
    for run in paragraph.findall("./w:r", NS):
        text = paragraph_text(run)
        if not text:
            continue
        rpr = run.find("./w:rPr", NS)
        copied = deepcopy(rpr) if rpr is not None else None
        if not seen_separator and label_rpr is None and text.strip():
            label_rpr = copied
        if not seen_separator and (":" in text or "\uff1a" in text):
            separator_rpr = copied
            seen_separator = True
            found_separator = True
            continue
        if seen_separator and content_rpr is None and text.strip():
            content_rpr = copied
            found_content = True
            break
    if not found_separator:
        separator_rpr = label_rpr
    if not found_content:
        content_rpr = separator_rpr
    return label_rpr, separator_rpr, content_rpr


def paragraph_style_name(paragraph: ET.Element, style_names: dict[str, str]) -> str:
    sid = style_id(paragraph) or ""
    return style_names.get(sid, "").strip().lower()


def is_toc_style_name(name: str) -> bool:
    lowered = name.strip().lower()
    return lowered.startswith("toc") or "\u76ee\u5f55" in lowered


def is_instruction_like_abstract_donor(paragraph: ET.Element, style_names: dict[str, str]) -> bool:
    text = paragraph_text(paragraph).strip()
    lowered = text.lower()
    instruction_tokens = (
        "\u5b8b\u4f53",
        "\u9ed1\u4f53",
        "\u5c0f\u56db",
        "\u56db\u53f7",
        "\u5b57\u53f7",
        "\u52a0\u7c97",
        "\u7f29\u8fdb",
        "\u884c\u8ddd",
        "\u5185\u5bb9",
        "\u683c\u5f0f",
        "times new roman",
        "font",
    )
    token_hits = sum(1 for token in instruction_tokens if token in lowered)
    return token_hits >= 2 or is_toc_style_name(paragraph_style_name(paragraph, style_names))


def split_keyword_line(text: str) -> tuple[str, str, str]:
    stripped = text.strip()
    if stripped.startswith("\u5173\u952e\u8bcd") and "\uff1a" in stripped:
        label, rest = stripped.split("\uff1a", 1)
        return label, "\uff1a", rest.lstrip()
    if stripped.startswith("关键词") and "：" in stripped:
        label, rest = stripped.split("：", 1)
        return label, "：", rest
    lowered = stripped.lower()
    if (lowered.startswith("keywords") or lowered.startswith("key words")) and ":" in stripped:
        label, rest = stripped.split(":", 1)
        separator = ": " if rest.startswith(" ") else ":"
        return label, separator, rest.lstrip()
    return "", "", stripped


def set_keyword_line_text_with_roles(
    paragraph: ET.Element,
    text: str,
    *,
    label_rpr: ET.Element | None,
    separator_rpr: ET.Element | None,
    content_rpr: ET.Element | None,
) -> None:
    _starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    label, separator, content = split_keyword_line(text)
    if label:
        append_text_run(paragraph, label + separator, strong_label_rpr(first_present_element(label_rpr, separator_rpr)))
    elif separator:
        append_text_run(paragraph, separator, strong_label_rpr(first_present_element(separator_rpr, label_rpr)))
    if content:
        append_text_run(paragraph, content, content_rpr)
    for child in ends:
        paragraph.append(child)


def set_keyword_line_text(paragraph: ET.Element, text: str, template_para: ET.Element) -> None:
    _starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    label, separator, content = split_keyword_line(text)
    label_rpr, separator_rpr, content_rpr = keyword_template_role_rprs(template_para)
    label_rpr = strong_label_rpr(first_present_element(label_rpr, separator_rpr))
    if label:
        append_text_run(paragraph, label + separator, label_rpr)
    elif separator:
        append_text_run(paragraph, separator, strong_label_rpr(separator_rpr))
    if content:
        append_text_run(paragraph, content, content_rpr)
    for child in ends:
        paragraph.append(child)


def remove_children(parent: ET.Element, tags: set[str]) -> None:
    for child in list(parent):
        if child.tag in tags:
            parent.remove(child)


def style_id(paragraph: ET.Element) -> str | None:
    ppr = paragraph_property(paragraph)
    if ppr is None:
        return None
    style = ppr.find("./w:pStyle", NS)
    return style.get(qn("val")) if style is not None else None


def set_style(ppr: ET.Element, value: str) -> None:
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        style = ET.Element(qn("pStyle"))
        ppr.insert(0, style)
    style.set(qn("val"), value)


def normal_style_id(styles_xml: bytes) -> str:
    root = ET.fromstring(styles_xml)
    for style in root.findall("./w:style", NS):
        name = style.find("./w:name", NS)
        if name is not None and name.get(qn("val")) == "Normal":
            return style.get(qn("styleId")) or "Normal"
    return "Normal"


def ensure_image_holder_style_xml(styles_xml: bytes) -> tuple[bytes, str, bool]:
    if not styles_xml:
        return styles_xml, "ThesisImageHolder", False
    root = ET.fromstring(styles_xml)
    desired_id = "ThesisImageHolder"
    existing = None
    for style in root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id_value = style.get(qn("styleId")) or ""
        name = style.find("./w:name", NS)
        style_name = name.get(qn("val")) if name is not None else ""
        if style_id_value.lower() == desired_id.lower() or style_name.lower() == "thesis image holder":
            existing = style
            desired_id = style_id_value or desired_id
            break
    changed = False
    if existing is None:
        existing = ET.Element(qn("style"), {qn("type"): "paragraph", qn("styleId"): desired_id})
        ET.SubElement(existing, qn("name"), {qn("val"): "Thesis Image Holder"})
        ET.SubElement(existing, qn("uiPriority"), {qn("val"): "99"})
        ET.SubElement(existing, qn("unhideWhenUsed"))
        root.append(existing)
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
    if not changed:
        return styles_xml, desired_id, False
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), desired_id, True


def style_element_by_id(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for style in styles_root.findall("./w:style", NS):
        if style.get(qn("styleId")) == style_id:
            return style
    return None


def ensure_style_rpr(style: ET.Element) -> ET.Element:
    rpr = style.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        style.append(rpr)
    return rpr


def sync_normal_style_font_baseline(
    target_styles_xml: bytes,
    template_styles_xml: bytes,
    *,
    target_normal_style_id: str,
    template_normal_style_id: str,
) -> bytes:
    if not target_styles_xml or not template_styles_xml:
        return target_styles_xml
    target_root = ET.fromstring(target_styles_xml)
    template_root = ET.fromstring(template_styles_xml)
    target_style = style_element_by_id(target_root, target_normal_style_id)
    template_style = style_element_by_id(template_root, template_normal_style_id)
    if target_style is None or template_style is None:
        return target_styles_xml
    template_ppr = template_style.find("./w:pPr", NS)
    target_ppr = target_style.find("./w:pPr", NS)
    if template_ppr is not None:
        if target_ppr is not None:
            target_style.remove(target_ppr)
        insert_at = 0
        for index, child in enumerate(list(target_style)):
            if child.tag in {qn("name"), qn("aliases"), qn("basedOn"), qn("next"), qn("link"), qn("uiPriority"), qn("qFormat")}:
                insert_at = index + 1
        target_style.insert(insert_at, deepcopy(template_ppr))
    template_rfonts = template_style.find("./w:rPr/w:rFonts", NS)
    if template_rfonts is None:
        return target_styles_xml
    target_rpr = ensure_style_rpr(target_style)
    target_rfonts = target_rpr.find("./w:rFonts", NS)
    if target_rfonts is None:
        target_rfonts = ET.Element(qn("rFonts"))
        target_rpr.insert(0, target_rfonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "cstheme"):
        value = template_rfonts.get(qn(key))
        if value is not None:
            target_rfonts.set(qn(key), value)
        elif target_rfonts.get(qn(key)) is not None:
            del target_rfonts.attrib[qn(key)]
    for tag in ("sz", "szCs"):
        template_node = template_style.find(f"./w:rPr/w:{tag}", NS)
        target_node = target_rpr.find(f"./w:{tag}", NS)
        if template_node is not None:
            if target_node is None:
                target_node = ET.Element(qn(tag))
                target_rpr.append(target_node)
            target_node.set(qn("val"), template_node.get(qn("val")) or "")
        elif target_node is not None:
            target_rpr.remove(target_node)
    remove_annotation_visual_markers(target_rpr)
    return ET.tostring(target_root, encoding="utf-8", xml_declaration=True)


def paragraph_style_id_to_name(styles_xml: bytes) -> dict[str, str]:
    root = ET.fromstring(styles_xml)
    mapping: dict[str, str] = {}
    for style in root.findall("./w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        style_id_value = style.get(qn("styleId")) or ""
        name = style.find("./w:name", NS)
        if style_id_value and name is not None:
            mapping[style_id_value] = (name.get(qn("val")) or "").lower()
    return mapping


def paragraph_style_name_to_id(styles_xml: bytes) -> dict[str, str]:
    return {name: style_id_value for style_id_value, name in paragraph_style_id_to_name(styles_xml).items() if name}


def remap_ppr_style_id(
    ppr: ET.Element | None,
    source_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> ET.Element | None:
    if ppr is None:
        return None
    copied = deepcopy(ppr)
    style = copied.find("./w:pStyle", NS)
    if style is not None:
        source_style_id = style.get(qn("val")) or ""
        source_style_name = source_style_names.get(source_style_id, "")
        target_style_id = target_style_ids.get(source_style_name, "")
        if target_style_id:
            style.set(qn("val"), target_style_id)
    return copied


def is_heading_text(text: str) -> bool:
    stripped = text.strip()
    if not stripped or len(stripped) > 80:
        return False
    normalized = stripped.replace("\uff0e", ".").replace("\u3002", ".")
    if re.match(r"^\u7b2c\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\s*\u7ae0\s+\S", stripped):
        return True
    if re.match(r"^\d+(?:\.\d+)+\s+", normalized):
        return True
    if re.match(r"^\d+(?:[.]\d+)+\s*\S", normalized):
        return True
    return False


BODY_HEADING_STYLE_LEVELS = {
    "2": 0,
    "21": 0,
    "23": 1,
    "4": 2,
}


def body_heading_level_from_style(paragraph: ET.Element) -> int | None:
    style = style_id(paragraph)
    if not style:
        return None
    return BODY_HEADING_STYLE_LEVELS.get(style)


def is_body_heading_paragraph(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph)
    if not text.strip():
        return False
    style = style_id(paragraph) or ""
    return body_heading_level_from_style(paragraph) is not None or style.lower().startswith("heading") or is_heading_text(text)


def is_caption_or_reference(text: str) -> bool:
    stripped = text.strip()
    compact = compact_text(stripped)
    figure_or_table_label = re.match(r"^[\u56fe\u8868]\s*\d+(?:[-.]\d+)?", stripped)
    if figure_or_table_label is not None:
        after_label = stripped[figure_or_table_label.end():]
        if after_label and (not after_label[0].isspace()) and after_label[0] not in {"\u3000", ":", "\uff1a", "\u25a1"}:
            return False
    return (
        re.match(r"^\u56fe\s*\d+(?:[-.]\d+)?(?:\s|[\u3000:：])", stripped) is not None
        or re.match(r"^\u8868\s*\d+(?:[-.]\d+)?(?:\s|[\u3000:：])", stripped) is not None
        or re.match(r"^\u56fe\s*[0-9\u4e00-\u9fff]+(?:[-.][0-9\u4e00-\u9fff]+)?\s+\S", stripped) is not None
        or re.match(r"^\u8868\s*[0-9\u4e00-\u9fff]+(?:[-.][0-9\u4e00-\u9fff]+)?\s+\S", stripped) is not None
        or compact.startswith("\u5173\u952e\u8bcd")
        or compact in {"\u6458\u8981", "\u76ee\u5f55", "\u53c2\u8003\u6587\u732e", "\u81f4\u8c22", "\u8c22\u8f9e", "\u9644\u5f55", "\u7ed3\u8bed"}
        or re.match(r"^\[\d+\]", compact) is not None
        or compact in {"摘要", "摘要", "目录", "参考文献", "致谢", "谢辞", "谢辞", "附录"}
        or compact in {"摘  要".replace(" ", ""), "目  录".replace(" ", ""), "参  考  文  献".replace(" ", "")}
    )


def is_citation_run(run: ET.Element) -> bool:
    text = paragraph_text(run).strip()
    rpr = run.find("./w:rPr", NS)
    if rpr is not None and rpr.find("./w:vertAlign", NS) is not None:
        return True
    return re.fullmatch(r"\[\d+\]", text) is not None


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


def has_field_or_drawing(run: ET.Element) -> bool:
    return (
        run.find(".//w:fldChar", NS) is not None
        or run.find(".//w:instrText", NS) is not None
        or run.find(".//w:drawing", NS) is not None
        or run.find(".//w:pict", NS) is not None
    )


def run_inside_hyperlink(paragraph: ET.Element, run: ET.Element) -> bool:
    for hyperlink in paragraph.findall(".//w:hyperlink", NS):
        if any(candidate is run for candidate in hyperlink.findall(".//w:r", NS)):
            return True
    return False


def paragraph_has_field_drawing_or_citation(paragraph: ET.Element) -> bool:
    for run in paragraph.findall(".//w:r", NS):
        if has_field_or_drawing(run) or is_citation_run(run) or run_inside_hyperlink(paragraph, run):
            return True
    if paragraph.find(".//w:bookmarkStart", NS) is not None or paragraph.find(".//w:bookmarkEnd", NS) is not None:
        return True
    return False


def find_first(paragraphs: list[ET.Element], predicate) -> tuple[int, ET.Element]:
    for index, paragraph in enumerate(paragraphs):
        if predicate(index, paragraph):
            return index, paragraph
    raise ValueError("required paragraph donor not found")


def find_abstract_surfaces(paragraphs: list[ET.Element]) -> dict[str, tuple[int, ET.Element]]:
    zh_title = find_first(paragraphs, lambda _i, p: is_zh_abstract_title_text(paragraph_text(p)))
    zh_body = find_first(
        paragraphs[zh_title[0] + 1 :],
        lambda _i, p: bool(paragraph_text(p).strip()) and not compact_text(paragraph_text(p)).startswith("关键词"),
    )
    zh_body = (zh_title[0] + 1 + zh_body[0], zh_body[1])
    zh_keyword = find_first(paragraphs[zh_body[0] + 1 :], lambda _i, p: compact_text(paragraph_text(p)).startswith("关键词"))
    zh_keyword = (zh_body[0] + 1 + zh_keyword[0], zh_keyword[1])
    en_title = find_first(
        paragraphs,
        lambda _i, p: is_en_abstract_title_text(paragraph_text(p)),
    )
    en_body = find_first(
        paragraphs[en_title[0] + 1 :],
        lambda _i, p: bool(paragraph_text(p).strip()) and not paragraph_text(p).strip().lower().startswith("keywords"),
    )
    en_body = (en_title[0] + 1 + en_body[0], en_body[1])
    en_keyword = find_first(paragraphs[en_body[0] + 1 :], lambda _i, p: paragraph_text(p).strip().lower().startswith("keywords"))
    en_keyword = (en_body[0] + 1 + en_keyword[0], en_keyword[1])
    return {
        "zh_abstract_title": zh_title,
        "zh_abstract_body": zh_body,
        "zh_keyword_line": zh_keyword,
        "en_abstract_title": en_title,
        "en_abstract_body": en_body,
        "en_keyword_line": en_keyword,
    }


def find_abstract_surfaces(paragraphs: list[ET.Element]) -> dict[str, tuple[int, ET.Element]]:
    """Locate abstract surfaces with Chinese-space and `Key words` variants."""
    zh_title = find_first(
        paragraphs,
        lambda _i, p: is_zh_abstract_title_text(paragraph_text(p)),
    )
    if abstract_inline_body_text(paragraph_text(zh_title[1])):
        zh_body = zh_title
    else:
        zh_body = find_first(
            paragraphs[zh_title[0] + 1 :],
            lambda _i, p: bool(paragraph_text(p).strip())
            and not compact_text(paragraph_text(p)).startswith("\u5173\u952e\u8bcd"),
        )
        zh_body = (zh_title[0] + 1 + zh_body[0], zh_body[1])
    zh_keyword = find_first(
        paragraphs[zh_body[0] + 1 :],
        lambda _i, p: compact_text(paragraph_text(p)).startswith("\u5173\u952e\u8bcd"),
    )
    zh_keyword = (zh_body[0] + 1 + zh_keyword[0], zh_keyword[1])
    en_title = find_first(
        paragraphs,
        lambda _i, p: is_en_abstract_title_text(paragraph_text(p)),
    )
    if abstract_inline_body_text(paragraph_text(en_title[1])):
        en_body = en_title
    else:
        en_body = find_first(
            paragraphs[en_title[0] + 1 :],
            lambda _i, p: bool(paragraph_text(p).strip())
            and not paragraph_text(p).strip().lower().startswith(("keywords", "key words")),
        )
        en_body = (en_title[0] + 1 + en_body[0], en_body[1])
    en_keyword = find_first(
        paragraphs[en_body[0] + 1 :],
        lambda _i, p: paragraph_text(p).strip().lower().startswith(("keywords", "key words")),
    )
    en_keyword = (en_body[0] + 1 + en_keyword[0], en_keyword[1])
    return {
        "zh_abstract_title": zh_title,
        "zh_abstract_body": zh_body,
        "zh_keyword_line": zh_keyword,
        "en_abstract_title": en_title,
        "en_abstract_body": en_body,
        "en_keyword_line": en_keyword,
    }


def first_explicit_run_size_half_points(paragraph: ET.Element) -> int | None:
    for run in paragraph.findall("./w:r", NS):
        if not paragraph_text(run).strip():
            continue
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            continue
        size = rpr.find("./w:sz", NS) or rpr.find("./w:szCs", NS)
        if size is not None and (size.get(qn("val")) or "").isdigit():
            return int(size.get(qn("val")) or "0")
    return None


def find_first_real_body_heading_index(paragraphs: list[ET.Element]) -> int:
    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        level = heading_level_from_paragraph(paragraph, text)
        if level is None:
            continue
        compact = compact_text(text)
        if compact in {"\u6458\u8981", "abstract", "\u76ee\u5f55", "\u53c2\u8003\u6587\u732e", "\u81f4\u8c22", "\u9644\u5f55"}:
            continue
        return index
    raise ValueError("template real body heading not found")


def is_template_body_donor_candidate(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph).strip()
    if len(text) < 40:
        return False
    if is_body_heading_paragraph(paragraph) or is_caption_or_reference(text) or is_code_like_paragraph_text(text):
        return False
    ppr = paragraph_property(paragraph)
    if ppr is None:
        return False
    style = ppr.find("./w:pStyle", NS)
    indent = ppr.find("./w:ind", NS)
    spacing = ppr.find("./w:spacing", NS)
    size = first_explicit_run_size_half_points(paragraph)
    return (
        (style is None or style.get(qn("val")) in {"Normal", "a", "1"})
        and indent is not None
        and indent.get(qn("firstLineChars")) in {"100", "200"}
        and spacing is not None
        and spacing.get(qn("line")) in {"360", "440"}
        and (size is None or 18 <= size <= 24)
    )


def find_template_body_donor(paragraphs: list[ET.Element]) -> tuple[int, ET.Element]:
    try:
        body_heading_index: int | None = find_first_real_body_heading_index(paragraphs)
    except ValueError:
        body_heading_index = None
    if body_heading_index is not None:
        for index, paragraph in enumerate(paragraphs[body_heading_index + 1 :], start=body_heading_index + 1):
            if is_template_body_donor_candidate(paragraph):
                return index, paragraph
    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph).strip()
        if (
            body_heading_index is not None
            and index <= body_heading_index
        ) or compact_text(text) in {"\u6458\u8981", "abstract", "\u76ee\u5f55"}:
            continue
        if is_template_body_donor_candidate(paragraph):
            return index, paragraph
    raise ValueError("template body donor paragraph not found")


def is_safe_zh_keyword_text(text: str) -> bool:
    compact = compact_text(text)
    return compact.startswith("\u5173\u952e\u8bcd") and "." not in compact


def is_safe_en_keyword_text(text: str) -> bool:
    stripped = text.strip()
    lowered = stripped.lower()
    compact = compact_text(stripped)
    return lowered.startswith(("keywords", "key words")) and "." not in compact


KEYWORD_LINE_SIZE_HALF_POINTS = 24


def remove_rfont_slots(rpr: ET.Element | None, *slots: str) -> None:
    if rpr is None:
        return
    rfonts = rpr.find("./w:rFonts", NS)
    if rfonts is None:
        return
    for slot in slots:
        rfonts.attrib.pop(qn(slot), None)


def keyword_line_rpr(*, english: bool, label: bool, segment: str = "") -> ET.Element:
    rpr = ET.Element(qn("rPr"))
    if label:
        ensure_rpr_bold(rpr)
    if english:
        set_font_slots(rpr, latin="Times New Roman")
        if CJK_RE.search(segment):
            set_font_slots(rpr, east_asia="\u5b8b\u4f53")
    elif label or not ASCII_ALNUM_RE.search(segment):
        set_font_slots(rpr, east_asia="\u5b8b\u4f53", latin="\u5b8b\u4f53")
    else:
        set_font_slots(rpr, east_asia="\u5b8b\u4f53", latin="Times New Roman")
    set_run_size(rpr, KEYWORD_LINE_SIZE_HALF_POINTS)
    return rpr


def set_keyword_line_text_without_template(paragraph: ET.Element, text: str, *, english: bool = False) -> None:
    _starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    label, separator, content = split_keyword_line(text)
    if not label and not separator:
        raise ValueError("keyword line must contain a keyword label and separator")
    label_text = label + separator
    append_text_run(paragraph, label_text, keyword_line_rpr(english=english, label=True, segment=label_text))
    if content:
        for segment in split_mixed_script_text(content):
            append_text_run(
                paragraph,
                segment,
                keyword_line_rpr(english=english, label=False, segment=segment),
            )
    for child in ends:
        paragraph.append(child)


def remove_bold_from_rpr(rpr: ET.Element | None) -> None:
    if rpr is None:
        return
    for child in list(rpr):
        if child.tag in {qn("b"), qn("bCs")}:
            rpr.remove(child)


def set_font_slots(rpr: ET.Element | None, *, east_asia: str | None = None, latin: str | None = None) -> None:
    if rpr is None:
        return
    rfonts = ensure_rpr_fonts(rpr)
    if east_asia:
        rfonts.set(qn("eastAsia"), east_asia)
    if latin:
        for attr in ("ascii", "hAnsi", "cs"):
            rfonts.set(qn(attr), latin)


def normalize_keyword_line_role_format(paragraph: ET.Element, *, english: bool) -> None:
    runs = paragraph.findall("./w:r", NS)
    for index, run in enumerate(runs):
        rpr = ensure_run_rpr(run)
        text = paragraph_text(run)
        if index == 0:
            ensure_rpr_bold(rpr)
        else:
            remove_bold_from_rpr(rpr)
        if english:
            set_font_slots(rpr, latin="Times New Roman")
            if index == 0 or not CJK_RE.search(text):
                remove_rfont_slots(rpr, "eastAsia")
            else:
                set_font_slots(rpr, east_asia="\u5b8b\u4f53")
        elif index == 0 or not ASCII_ALNUM_RE.search(text):
            set_font_slots(rpr, east_asia="\u5b8b\u4f53", latin="\u5b8b\u4f53")
        else:
            set_font_slots(rpr, east_asia="\u5b8b\u4f53", latin="Times New Roman")
        set_run_size(rpr, KEYWORD_LINE_SIZE_HALF_POINTS)
    split_mixed_script_runs(paragraph)


def normalize_abstract_body_role_format(paragraph: ET.Element, *, english: bool) -> None:
    ppr = paragraph_property(paragraph)
    if ppr is not None:
        ppr_rpr = ppr.find("./w:rPr", NS)
        remove_bold_from_rpr(ppr_rpr)
        set_font_slots(ppr_rpr, east_asia="\u5b8b\u4f53", latin="Times New Roman")
    for run in paragraph.findall("./w:r", NS):
        if has_field_or_drawing(run) or is_citation_run(run):
            continue
        rpr = ensure_run_rpr(run)
        remove_bold_from_rpr(rpr)
        set_font_slots(rpr, east_asia="\u5b8b\u4f53", latin="Times New Roman")


def split_inline_abstract_text(text: str, *, english: bool) -> tuple[str, str, str, str]:
    stripped = text.strip()
    if english:
        match = re.match(r"^(Abstract)([:\uff1a])\s*(.*)$", stripped, flags=re.IGNORECASE | re.DOTALL)
        if not match:
            raise ValueError("English inline abstract paragraph must start with Abstract: or Abstract\uff1a")
        return match.group(1), "", match.group(2), match.group(3)
    match = re.match(r"^(\u6458)\s*(\u8981)([:\uff1a])\s*(.*)$", stripped, flags=re.DOTALL)
    if not match:
        raise ValueError("Chinese inline abstract paragraph must start with \u6458  \u8981\uff1a")
    return match.group(1), "  ", match.group(2) + match.group(3), match.group(4)


def append_segmented_body_runs(paragraph: ET.Element, text: str, donor_rpr: ET.Element | None) -> None:
    for segment in split_mixed_script_text(text):
        if not segment:
            continue
        if ASCII_ALNUM_RE.search(segment):
            append_text_run(paragraph, segment, clone_rpr_with_latin_slots(donor_rpr))
        else:
            append_text_run(paragraph, segment, deepcopy(donor_rpr) if donor_rpr is not None else None)


def template_inline_title_rprs(template_para: ET.Element, *, english: bool) -> tuple[ET.Element | None, ET.Element | None, ET.Element | None, ET.Element | None]:
    runs = [run for run in template_para.findall("./w:r", NS) if paragraph_text(run)]
    if not runs:
        return None, None, None, None
    if english:
        title = runs[0].find("./w:rPr", NS)
        separator = runs[1].find("./w:rPr", NS) if len(runs) > 1 else title
        body = next((run.find("./w:rPr", NS) for run in runs[2:] if paragraph_text(run).strip()), None)
        return deepcopy(title) if title is not None else None, None, deepcopy(separator) if separator is not None else None, deepcopy(body) if body is not None else None
    title_runs = [run.find("./w:rPr", NS) for run in runs[:4]]
    title = next((rpr for rpr in title_runs if rpr is not None), None)
    separator = title_runs[3] if len(title_runs) >= 4 and title_runs[3] is not None else title
    body = next((run.find("./w:rPr", NS) for run in runs[4:] if paragraph_text(run).strip()), None)
    return deepcopy(title) if title is not None else None, deepcopy(title) if title is not None else None, deepcopy(separator) if separator is not None else None, deepcopy(body) if body is not None else None


def set_inline_abstract_paragraph(
    paragraph: ET.Element,
    text: str,
    template_para: ET.Element,
    *,
    english: bool,
) -> None:
    label_a, label_space, label_b_or_sep, body_text = split_inline_abstract_text(text, english=english)
    title_rpr, title_space_rpr, separator_rpr, body_rpr = template_inline_title_rprs(template_para, english=english)
    starts, ends = clear_paragraph_content_preserving_review_anchors(paragraph)
    for child in starts:
        paragraph.append(child)
    if english:
        append_text_run(paragraph, "Abstract", title_rpr)
        append_text_run(paragraph, "\uff1a", separator_rpr or title_rpr)
    else:
        append_text_run(paragraph, "\u6458", title_rpr)
        append_text_run(paragraph, label_space or "  ", title_space_rpr or title_rpr)
        append_text_run(paragraph, "\u8981", title_rpr)
        append_text_run(paragraph, "\uff1a", separator_rpr or title_rpr)
    append_segmented_body_runs(paragraph, body_text.lstrip(), body_rpr)
    for child in ends:
        paragraph.append(child)


def repair_inline_abstract_surfaces_from_template(
    target_surfaces: dict[str, tuple[int, ET.Element]],
    template_surfaces: dict[str, tuple[int, ET.Element]],
    abstract_text: dict[str, object],
    *,
    template_style_names: dict[str, str],
    target_style_ids: dict[str, str],
) -> dict[str, object]:
    changed: dict[str, object] = {}
    for body_surface, title_surface, english in (
        ("zh_abstract_body", "zh_abstract_title", False),
        ("en_abstract_body", "en_abstract_title", True),
    ):
        target_index, target_para = target_surfaces[body_surface]
        _template_index, template_para = template_surfaces[body_surface]
        donor_ppr = remap_ppr_style_id(paragraph_property(template_para), template_style_names, target_style_ids)
        if donor_ppr is not None:
            replace_ppr(target_para, donor_ppr)
        replacement = str(abstract_text.get(body_surface, paragraph_text(target_para)))
        set_inline_abstract_paragraph(target_para, replacement, template_para, english=english)
        changed[title_surface] = {
            "target_paragraph_index": target_index,
            "inline_title_and_body_shared_paragraph": True,
            "template_text": paragraph_text(template_para)[:80],
            "final_text": paragraph_text(target_para)[:120],
        }
        changed[body_surface] = {
            "target_paragraph_index": target_index,
            "inline_title_and_body_shared_paragraph": True,
            "template_text": paragraph_text(template_para)[:80],
            "final_text": paragraph_text(target_para)[:120],
        }
    for keyword_surface in ("zh_keyword_line", "en_keyword_line"):
        target_index, target_para = target_surfaces[keyword_surface]
        _template_index, template_para = template_surfaces[keyword_surface]
        donor_ppr = remap_ppr_style_id(paragraph_property(template_para), template_style_names, target_style_ids)
        if donor_ppr is not None:
            replace_ppr(target_para, donor_ppr)
        set_keyword_line_text(target_para, str(abstract_text.get(keyword_surface, paragraph_text(target_para))), template_para)
        changed[keyword_surface] = {
            "target_paragraph_index": target_index,
            "template_text": paragraph_text(template_para)[:80],
            "final_text": paragraph_text(target_para)[:120],
        }
    return changed


def repair_keyword_only_surfaces(
    paragraphs: list[ET.Element],
    abstract_text: dict[str, object],
    *,
    allow_missing: bool = False,
) -> dict[str, object]:
    changed: dict[str, object] = {}
    zh_match = [(index, paragraph) for index, paragraph in enumerate(paragraphs) if is_safe_zh_keyword_text(paragraph_text(paragraph))]
    en_match = [(index, paragraph) for index, paragraph in enumerate(paragraphs) if is_safe_en_keyword_text(paragraph_text(paragraph))]
    if len(zh_match) != 1 and not (allow_missing and len(zh_match) == 0):
        raise ValueError(f"keyword-only repair requires exactly one Chinese keyword paragraph, found {len(zh_match)}")
    if len(en_match) != 1 and not (allow_missing and len(en_match) == 0):
        raise ValueError(f"keyword-only repair requires exactly one English keyword paragraph, found {len(en_match)}")
    matches: list[tuple[str, tuple[int, ET.Element]]] = []
    if zh_match:
        matches.append(("zh_keyword_line", zh_match[0]))
    if en_match:
        matches.append(("en_keyword_line", en_match[0]))
    for surface_id, match in matches:
        index, paragraph = match
        before = paragraph_text(paragraph)
        replacement = str(abstract_text.get(surface_id, before))
        english = surface_id == "en_keyword_line"
        set_keyword_line_text_without_template(paragraph, replacement, english=english)
        normalize_keyword_line_role_format(paragraph, english=english)
        changed[surface_id] = {
            "target_paragraph_index": index,
            "before_text": before[:120],
            "final_text": paragraph_text(paragraph)[:120],
            "keyword_only_rule_driven": True,
        }
    return changed


def repair_abstract_role_surfaces(paragraphs: list[ET.Element]) -> dict[str, object]:
    surfaces = find_abstract_surfaces(paragraphs)
    changed: dict[str, object] = {}
    for surface_id, english in (("zh_abstract_body", False), ("en_abstract_body", True)):
        index, paragraph = surfaces[surface_id]
        normalize_abstract_body_role_format(paragraph, english=english)
        changed[surface_id] = {
            "target_paragraph_index": index,
            "abstract_body_content_not_bold": True,
            "role_format_rule_driven": True,
        }
    for surface_id, english in (("zh_keyword_line", False), ("en_keyword_line", True)):
        index, paragraph = surfaces[surface_id]
        normalize_keyword_line_role_format(paragraph, english=english)
        changed[surface_id] = {
            "target_paragraph_index": index,
            "keyword_label_bold_content_not_bold": True,
            "role_format_rule_driven": True,
        }
    return changed


def clear_body_run_formatting(paragraph: ET.Element, donor_rpr: ET.Element | None) -> int:
    changed = 0
    safe_donor_rpr = deepcopy(donor_rpr) if donor_rpr is not None else None
    remove_annotation_visual_markers(safe_donor_rpr)
    for run in paragraph.findall(".//w:r", NS):
        if has_field_or_drawing(run) or is_citation_run(run) or run_inside_hyperlink(paragraph, run):
            continue
        replace_rpr(run, safe_donor_rpr)
        changed += 1
    return changed


def remove_document_annotation_run_colors(paragraphs: list[ET.Element]) -> int:
    changed = 0
    for paragraph in paragraphs:
        for run in paragraph.findall(".//w:r", NS):
            if has_field_or_drawing(run):
                continue
            changed += remove_annotation_visual_markers(run.find("./w:rPr", NS))
    return changed


def normalize_paragraph_to_donor(paragraph: ET.Element, donor_ppr: ET.Element | None, donor_rpr: ET.Element | None) -> int:
    if donor_ppr is not None:
        replace_ppr(paragraph, donor_ppr)
    return clear_body_run_formatting(paragraph, donor_rpr)


def force_visible_spacing_and_firstline(paragraph: ET.Element, *, firstline_twips: int | None) -> None:
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(qn("spacing"))
        ppr.append(spacing)
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(qn("ind"))
        ppr.append(ind)
    ind.set(qn("start"), "0")
    ind.set(qn("startChars"), "0")
    ind.set(qn("end"), "0")
    ind.set(qn("endChars"), "0")
    if firstline_twips is not None:
        ind.set(qn("firstLine"), str(max(0, firstline_twips)))


def heading_level_from_text(text: str) -> int | None:
    stripped = (text or "").strip().replace("\uff0e", ".").replace("\u3002", ".")
    if re.match(r"^\u7b2c\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\s*\u7ae0(?:\s|$)", stripped):
        return 0
    match = re.match(r"^(\d+(?:\.\d+){1,3})\s+\S", stripped)
    if not match:
        return None
    return min(match.group(1).count("."), 2)


def heading_level_from_paragraph(paragraph: ET.Element, text: str) -> int | None:
    text_level = heading_level_from_text(text)
    if text_level is not None:
        return text_level
    style_level = body_heading_level_from_style(paragraph)
    if style_level is not None:
        return style_level
    return None


def set_run_size(rpr: ET.Element, half_points: int) -> None:
    for tag in ("sz", "szCs"):
        node = rpr.find(f"./w:{tag}", NS)
        if node is None:
            node = ET.Element(qn(tag))
            rpr.append(node)
        node.set(qn("val"), str(half_points))


def force_body_heading_format(paragraph: ET.Element, text: str) -> dict[str, object] | None:
    level = heading_level_from_paragraph(paragraph, text)
    if level is None:
        return None
    ppr = ensure_ppr(paragraph)
    remove_children(ppr, {qn("spacing"), qn("ind"), qn("jc"), qn("outlineLvl")})
    spacing = ET.Element(qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    ppr.append(spacing)
    ind = ET.Element(qn("ind"))
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    ppr.append(ind)
    jc = ET.Element(qn("jc"))
    jc.set(qn("val"), "left")
    ppr.append(jc)
    outline = ET.Element(qn("outlineLvl"))
    outline.set(qn("val"), str(level))
    ppr.append(outline)

    expected_size = {0: 30, 1: 28, 2: 24}[level]
    changed_runs = 0
    for run in paragraph.findall("./w:r", NS):
        if has_field_or_drawing(run) or is_citation_run(run) or run_inside_hyperlink(paragraph, run) or not paragraph_text(run).strip():
            continue
        rpr = ensure_run_rpr(run)
        set_font_slots(rpr, east_asia="\u9ed1\u4f53", latin="\u9ed1\u4f53")
        remove_bold_from_rpr(rpr)
        set_run_size(rpr, expected_size)
        changed_runs += 1
    return {
        "heading_level": level + 1,
        "expected_size_half_points": expected_size,
        "changed_run_count": changed_runs,
    }


def apply_image_holder_ppr(paragraph: ET.Element, *, style_id: str = "ThesisImageHolder") -> None:
    ppr = ensure_ppr(paragraph)
    remove_children(
        ppr,
        {
            qn("pStyle"),
            qn("spacing"),
            qn("ind"),
            qn("jc"),
            qn("outlineLvl"),
            qn("numPr"),
            qn("pageBreakBefore"),
            qn("widowControl"),
            qn("suppressLineNumbers"),
        },
    )
    set_style(ppr, style_id)
    spacing = ET.Element(qn("spacing"))
    spacing.set(qn("before"), "120")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    ppr.append(spacing)
    ind = ET.Element(qn("ind"))
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    ind.set(qn("firstLineChars"), "0")
    ind.set(qn("leftChars"), "0")
    ind.set(qn("rightChars"), "0")
    ppr.append(ind)
    if ppr.find("./w:keepNext", NS) is None:
        ppr.append(ET.Element(qn("keepNext")))
    jc = ET.Element(qn("jc"))
    jc.set(qn("val"), "center")
    ppr.append(jc)
    if ppr.find("./w:rPr", NS) is None:
        rpr = ET.Element(qn("rPr"))
        fonts = ET.Element(qn("rFonts"))
        fonts.set(qn("hint"), "eastAsia")
        rpr.append(fonts)
        ppr.append(rpr)


def paragraph_comment_ids(paragraph: ET.Element) -> set[str]:
    ids: set[str] = set()
    for child in paragraph.iter():
        if child.tag in {qn("commentRangeStart"), qn("commentRangeEnd"), qn("commentReference")}:
            comment_id = child.attrib.get(qn("id"), "")
            if comment_id:
                ids.add(comment_id)
    return ids


def resize_drawing_extents(paragraph: ET.Element, *, max_width_cm: float, max_height_cm: float) -> list[dict[str, object]]:
    resized: list[dict[str, object]] = []
    for extent in paragraph.findall(".//wp:extent", NS):
        try:
            cx = int(extent.get("cx") or "0")
            cy = int(extent.get("cy") or "0")
        except ValueError:
            continue
        if cx <= 0 or cy <= 0:
            continue
        max_cx = int(max_width_cm * EMU_PER_CM) if max_width_cm > 0 else cx
        max_cy = int(max_height_cm * EMU_PER_CM) if max_height_cm > 0 else cy
        scale = min(max_cx / cx if max_cx else 1.0, max_cy / cy if max_cy else 1.0, 1.0)
        if scale >= 0.999:
            continue
        new_cx = max(1, int(round(cx * scale)))
        new_cy = max(1, int(round(cy * scale)))
        extent.set("cx", str(new_cx))
        extent.set("cy", str(new_cy))
        resized.append(
            {
                "before_cm": [round(cx / EMU_PER_CM, 2), round(cy / EMU_PER_CM, 2)],
                "after_cm": [round(new_cx / EMU_PER_CM, 2), round(new_cy / EMU_PER_CM, 2)],
            }
        )
        for graphic_extent in paragraph.findall(".//a:xfrm/a:ext", NS):
            if graphic_extent.get("cx") == str(cx) and graphic_extent.get("cy") == str(cy):
                graphic_extent.set("cx", str(new_cx))
                graphic_extent.set("cy", str(new_cy))
    return resized


def apply_image_display_resize(paragraphs: list[ET.Element], resize_plan: object) -> list[dict[str, object]]:
    if not isinstance(resize_plan, dict):
        return []
    raw_comment_ids = resize_plan.get("comment_ids") or resize_plan.get("comment_anchor_ids") or []
    comment_ids = {str(item).strip() for item in raw_comment_ids if str(item).strip()} if isinstance(raw_comment_ids, list) else set()
    max_width_cm = float(resize_plan.get("max_width_cm") or 0)
    max_height_cm = float(resize_plan.get("max_height_cm") or 0)
    start_paragraph_index = int(resize_plan.get("start_paragraph_index") or 0)
    include_uncommented = bool(resize_plan.get("include_uncommented"))
    if max_width_cm <= 0 and max_height_cm <= 0:
        return []
    changed: list[dict[str, object]] = []
    for index, paragraph in enumerate(paragraphs):
        if index < start_paragraph_index:
            continue
        if paragraph.find(".//w:drawing", NS) is None and paragraph.find(".//w:pict", NS) is None:
            continue
        para_comment_ids = paragraph_comment_ids(paragraph)
        if comment_ids and not (comment_ids & para_comment_ids) and not include_uncommented:
            continue
        resized = resize_drawing_extents(paragraph, max_width_cm=max_width_cm, max_height_cm=max_height_cm)
        if resized:
            changed.append(
                {
                    "target_paragraph_index": index,
                    "comment_ids": sorted(para_comment_ids),
                    "resizes": resized,
                }
            )
    return changed


def apply_body_text_replacements(
    paragraphs: list[ET.Element],
    plan: dict[str, object],
    *,
    body_ppr: ET.Element | None = None,
    body_rpr: ET.Element | None = None,
    firstline_twips: int | None = None,
) -> list[dict[str, object]]:
    replacements = plan.get("body_text_replacements", [])
    if replacements in (None, []):
        return []
    if not isinstance(replacements, list):
        raise ValueError("plan.body_text_replacements must be a list")
    changed: list[dict[str, object]] = []
    for item in replacements:
        if not isinstance(item, dict):
            raise ValueError("each body_text_replacements item must be an object")
        prefix = str(item.get("match_text_prefix", ""))
        replacement_text = item.get("replacement_text")
        if not prefix or replacement_text is None:
            raise ValueError("body text replacement requires match_text_prefix and replacement_text")
        matches = [(index, paragraph) for index, paragraph in enumerate(paragraphs) if paragraph_text(paragraph).startswith(prefix)]
        if len(matches) != 1:
            raise ValueError(f"body text replacement prefix matched {len(matches)} paragraphs: {prefix[:40]}")
        index, paragraph = matches[0]
        if paragraph_has_field_drawing_or_citation(paragraph):
            raise ValueError(f"body text replacement target has field, drawing, or citation markers at paragraph {index}")
        before = paragraph_text(paragraph)
        allow_mixed_script = bool(item.get("allow_mixed_script")) or bool(plan.get("allow_mixed_script_body_text_replacements"))
        if not allow_mixed_script and (has_mixed_cjk_ascii(before) or has_mixed_cjk_ascii(str(replacement_text))):
            raise ValueError(
                "body text replacement target contains mixed Chinese/ASCII content and must be repaired through the canonical mixed-script body-text path"
            )
        donor_rpr = body_rpr if body_rpr is not None else first_run_rpr(paragraph)
        if body_ppr is not None:
            replace_ppr(paragraph, body_ppr)
        if allow_mixed_script:
            set_paragraph_text_with_mixed_script(paragraph, str(replacement_text), donor_rpr)
        else:
            set_paragraph_text(paragraph, str(replacement_text), donor_rpr)
        if body_ppr is not None:
            force_visible_spacing_and_firstline(paragraph, firstline_twips=firstline_twips)
            apply_latin_font_slots(paragraph)
            split_mixed_script_runs(paragraph)
        changed.append(
            {
                "target_paragraph_index": index,
                "match_text_prefix": prefix,
                "before_text": before[:120],
                "after_text": paragraph_text(paragraph)[:120],
            }
        )
    return changed


def remove_empty_paragraphs_before_headings(
    body: ET.Element, paragraphs: list[ET.Element], plan: dict[str, object]
) -> list[dict[str, object]]:
    prefixes = plan.get("remove_empty_paragraphs_before_heading_prefixes", [])
    if prefixes in (None, []):
        return []
    if not isinstance(prefixes, list):
        raise ValueError("plan.remove_empty_paragraphs_before_heading_prefixes must be a list")
    changed: list[dict[str, object]] = []
    for raw_prefix in prefixes:
        exact_text = ""
        min_paragraph_index = 0
        if isinstance(raw_prefix, dict):
            prefix = str(raw_prefix.get("heading_prefix", ""))
            exact_text = str(raw_prefix.get("heading_text_exact", ""))
            min_paragraph_index = int(raw_prefix.get("min_paragraph_index", 0))
        else:
            prefix = str(raw_prefix)
        if not prefix and not exact_text:
            raise ValueError("heading prefix or exact heading text is required")
        if exact_text:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph) == exact_text
            ]
            display = exact_text
        else:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph).startswith(prefix)
            ]
            display = prefix
        if len(matches) != 1:
            raise ValueError(f"heading match selected {len(matches)} paragraphs: {display[:40]}")
        heading_index = matches[0]
        removed_indices: list[int] = []
        scan = heading_index - 1
        while scan >= 0 and not paragraph_text(paragraphs[scan]).strip():
            if paragraphs[scan].find(".//w:br", NS) is not None or paragraphs[scan].find("./w:pPr/w:sectPr", NS) is not None:
                break
            body.remove(paragraphs[scan])
            paragraphs.pop(scan)
            removed_indices.append(scan)
            heading_index -= 1
            scan -= 1
        changed.append(
            {
                "heading_prefix": prefix,
                "heading_text_exact": exact_text,
                "removed_empty_paragraph_indices": list(reversed(removed_indices)),
            }
        )
    return changed


def paragraph_has_page_break(paragraph: ET.Element) -> bool:
    for br in paragraph.findall(".//w:br", NS):
        if br.get(qn("type")) == "page":
            return True
    return False


def paragraph_has_page_break_before(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:pageBreakBefore", NS) is not None


def paragraph_has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def remove_redundant_empty_page_breaks_before_headings(
    body: ET.Element, paragraphs: list[ET.Element], plan: dict[str, object]
) -> list[dict[str, object]]:
    targets = plan.get("remove_redundant_empty_page_breaks_before_headings", [])
    if targets in (None, []):
        return []
    if not isinstance(targets, list):
        raise ValueError("plan.remove_redundant_empty_page_breaks_before_headings must be a list")
    changed: list[dict[str, object]] = []
    for raw_target in targets:
        if not isinstance(raw_target, dict):
            raise ValueError("each redundant page-break removal target must be an object")
        prefix = str(raw_target.get("heading_prefix", ""))
        exact_text = str(raw_target.get("heading_text_exact", ""))
        min_paragraph_index = int(raw_target.get("min_paragraph_index", 0))
        if not prefix and not exact_text:
            raise ValueError("redundant page-break removal target requires heading_prefix or heading_text_exact")
        if exact_text:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph) == exact_text
            ]
            display = exact_text
        else:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph).startswith(prefix)
            ]
            display = prefix
        if len(matches) != 1:
            raise ValueError(f"redundant page-break target selected {len(matches)} paragraphs: {display[:40]}")
        heading_index = matches[0]
        heading = paragraphs[heading_index]
        if not paragraph_has_page_break_before(heading):
            raise ValueError(
                "redundant page-break removal is allowed only when the target heading owns pageBreakBefore"
            )
        removed_indices: list[int] = []
        scan = heading_index - 1
        while scan >= 0 and not paragraph_text(paragraphs[scan]).strip():
            candidate = paragraphs[scan]
            if paragraph_has_section_break(candidate):
                break
            if candidate.find(".//w:drawing", NS) is not None or candidate.find(".//w:pict", NS) is not None:
                raise ValueError(f"redundant page-break candidate {scan} contains a drawing or pict object")
            if candidate.find(".//w:fldChar", NS) is not None or candidate.find(".//w:instrText", NS) is not None:
                raise ValueError(f"redundant page-break candidate {scan} contains a field")
            if not paragraph_has_page_break(candidate):
                break
            body.remove(candidate)
            paragraphs.pop(scan)
            removed_indices.append(scan)
            heading_index -= 1
            scan -= 1
        changed.append(
            {
                "heading_prefix": prefix,
                "heading_text_exact": exact_text,
                "removed_empty_page_break_paragraph_indices": list(reversed(removed_indices)),
            }
        )
    return changed


def previous_structural_page_owner(
    paragraphs: list[ET.Element],
    heading_index: int,
) -> tuple[int | None, str]:
    scan = heading_index - 1
    while scan >= 0:
        paragraph = paragraphs[scan]
        if paragraph_has_section_break(paragraph):
            return scan, "section_break"
        if paragraph_has_page_break(paragraph):
            return scan, "manual_page_break"
        if paragraph_text(paragraph).strip():
            return scan, "nonempty_text"
        scan -= 1
    return None, "none"


def remove_redundant_page_break_before_from_headings(
    paragraphs: list[ET.Element], plan: dict[str, object]
) -> list[dict[str, object]]:
    targets = plan.get("remove_redundant_page_break_before_from_headings", [])
    if targets in (None, []):
        return []
    if not isinstance(targets, list):
        raise ValueError("plan.remove_redundant_page_break_before_from_headings must be a list")
    changed: list[dict[str, object]] = []
    for raw_target in targets:
        if not isinstance(raw_target, dict):
            raise ValueError("each pageBreakBefore removal target must be an object")
        prefix = str(raw_target.get("heading_prefix", ""))
        exact_text = str(raw_target.get("heading_text_exact", ""))
        min_paragraph_index = int(raw_target.get("min_paragraph_index", 0))
        if not prefix and not exact_text:
            raise ValueError("pageBreakBefore removal target requires heading_prefix or heading_text_exact")
        if exact_text:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph) == exact_text
            ]
            display = exact_text
        else:
            matches = [
                index
                for index, paragraph in enumerate(paragraphs)
                if index >= min_paragraph_index and paragraph_text(paragraph).startswith(prefix)
            ]
            display = prefix
        if len(matches) != 1:
            raise ValueError(f"pageBreakBefore removal target selected {len(matches)} paragraphs: {display[:40]}")
        heading_index = matches[0]
        heading = paragraphs[heading_index]
        ppr = paragraph_property(heading)
        if ppr is None:
            raise ValueError("pageBreakBefore removal target has no paragraph properties")
        page_break_before = ppr.find("./w:pageBreakBefore", NS)
        if page_break_before is None:
            raise ValueError("pageBreakBefore removal target does not own pageBreakBefore")
        owner_index, owner_kind = previous_structural_page_owner(paragraphs, heading_index)
        if owner_kind not in {"section_break", "manual_page_break"}:
            raise ValueError(
                "pageBreakBefore removal requires a preceding section break or manual page break page owner"
            )
        ppr.remove(page_break_before)
        changed.append(
            {
                "heading_prefix": prefix,
                "heading_text_exact": exact_text,
                "target_paragraph_index": heading_index,
                "removed": "w:pageBreakBefore",
                "preceding_page_owner_index": owner_index,
                "preceding_page_owner_kind": owner_kind,
            }
        )
    return changed


def suppress_paragraph_auto_numbering(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    remove_children(ppr, {qn("numPr")})
    num_pr = ET.Element(qn("numPr"))
    num_id = ET.Element(qn("numId"))
    # Word/WPS use numId=0 as the explicit "no numbering" override.  This
    # keeps the heading style/outline level while blocking corrupt list output.
    num_id.set(qn("val"), "0")
    num_pr.append(num_id)
    ppr.append(num_pr)


def add_explicit_heading_number_text(paragraph: ET.Element, number_text: str) -> str:
    text = paragraph_text(paragraph)
    normalized = text.lstrip()
    prefix = f"{number_text} "
    if normalized.startswith(prefix) or normalized == number_text:
        return text
    for node in paragraph.findall(".//w:t", NS):
        if node.text and node.text.strip():
            leading = node.text[: len(node.text) - len(node.text.lstrip())]
            node.text = f"{leading}{prefix}{node.text.lstrip()}"
            return text
    raise ValueError("heading paragraph has no editable text run")


def apply_explicit_body_heading_numbers(
    paragraphs: list[ET.Element], plan: dict[str, object]
) -> list[dict[str, object]]:
    items = plan.get("explicit_body_heading_numbers", [])
    if items in (None, []):
        return []
    if not isinstance(items, list):
        raise ValueError("plan.explicit_body_heading_numbers must be a list")
    changed: list[dict[str, object]] = []
    used_indices: set[int] = set()
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("each explicit_body_heading_numbers item must be an object")
        exact_text = str(item.get("heading_text_exact") or "").strip()
        number_text = str(item.get("number_text") or "").strip()
        if not exact_text or not number_text:
            raise ValueError("explicit heading number repair requires heading_text_exact and number_text")
        min_paragraph_index = int(item.get("min_paragraph_index") or 0)
        matches = [
            (index, paragraph)
            for index, paragraph in enumerate(paragraphs)
            if index >= min_paragraph_index
            and index not in used_indices
            and paragraph_text(paragraph).strip() == exact_text
        ]
        if len(matches) != 1:
            raise ValueError(f"explicit heading number repair matched {len(matches)} paragraphs: {exact_text[:40]}")
        index, paragraph = matches[0]
        if paragraph_has_field_drawing_or_citation(paragraph):
            raise ValueError(f"explicit heading number target has field, drawing, or citation markers at paragraph {index}")
        before = add_explicit_heading_number_text(paragraph, number_text)
        suppress_paragraph_auto_numbering(paragraph)
        after = paragraph_text(paragraph)
        used_indices.add(index)
        changed.append(
            {
                "target_paragraph_index": index,
                "heading_text_exact": exact_text,
                "number_text": number_text,
                "before_text": before,
                "after_text": after,
                "auto_numbering_suppressed": True,
            }
        )
    return changed


P_XML_RE = re.compile(rb"<w:p(?:\s[^>]*)?>.*?</w:p>", re.DOTALL)
T_XML_RE = re.compile(rb"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.DOTALL)


def paragraph_text_from_xml_bytes(paragraph_xml: bytes) -> str:
    parts: list[str] = []
    for match in T_XML_RE.finditer(paragraph_xml):
        parts.append(html.unescape(match.group(2).decode("utf-8")))
    return "".join(parts)


def xml_text_escape(text: str) -> bytes:
    return html.escape(text, quote=False).encode("utf-8")


def insert_no_numbering_override_in_ppr(paragraph_xml: bytes) -> bytes:
    no_numbering = b'<w:numPr><w:numId w:val="0"/></w:numPr>'
    ppr_match = re.search(rb"<w:pPr\b[^>]*>.*?</w:pPr>", paragraph_xml, flags=re.DOTALL)
    if ppr_match is None:
        start_match = re.match(rb"<w:p(?:\s[^>]*)?>", paragraph_xml)
        if start_match is None:
            raise ValueError("paragraph XML start tag not found")
        insert_at = start_match.end()
        return paragraph_xml[:insert_at] + b"<w:pPr>" + no_numbering + b"</w:pPr>" + paragraph_xml[insert_at:]

    ppr_xml = ppr_match.group(0)
    ppr_xml = re.sub(rb"<w:numPr\b[^>]*>.*?</w:numPr>", b"", ppr_xml, flags=re.DOTALL)
    style_match = re.search(rb"</w:pStyle>", ppr_xml)
    if style_match is not None:
        insert_at = style_match.end()
    else:
        open_match = re.match(rb"<w:pPr\b[^>]*>", ppr_xml)
        if open_match is None:
            raise ValueError("paragraph properties XML start tag not found")
        insert_at = open_match.end()
    ppr_xml = ppr_xml[:insert_at] + no_numbering + ppr_xml[insert_at:]
    return paragraph_xml[: ppr_match.start()] + ppr_xml + paragraph_xml[ppr_match.end() :]


def prepend_heading_number_in_paragraph_xml(paragraph_xml: bytes, number_text: str) -> tuple[bytes, str, str]:
    before_text = paragraph_text_from_xml_bytes(paragraph_xml)
    stripped = before_text.lstrip()
    prefix = f"{number_text} "
    if stripped.startswith(prefix) or stripped == number_text:
        numbered_xml = paragraph_xml
    else:
        def replace_first_nonempty(match: re.Match[bytes]) -> bytes:
            content = html.unescape(match.group(2).decode("utf-8"))
            if not content.strip():
                return match.group(0)
            leading = content[: len(content) - len(content.lstrip())]
            new_content = f"{leading}{prefix}{content.lstrip()}"
            return match.group(1) + xml_text_escape(new_content) + match.group(3)

        numbered_xml, replacements = T_XML_RE.subn(replace_first_nonempty, paragraph_xml, count=1)
        if replacements != 1 or paragraph_text_from_xml_bytes(numbered_xml) == before_text:
            raise ValueError("heading paragraph has no editable non-empty text run")
    patched_xml = insert_no_numbering_override_in_ppr(numbered_xml)
    return patched_xml, before_text, paragraph_text_from_xml_bytes(patched_xml)


def set_paragraph_section_type(paragraph_xml: bytes, section_type: str) -> bytes:
    if section_type not in {"nextPage", "continuous", "evenPage", "oddPage"}:
        raise ValueError(f"unsupported section type: {section_type}")
    sect_match = re.search(rb"<w:sectPr\b[^>]*>.*?</w:sectPr>", paragraph_xml, flags=re.DOTALL)
    if sect_match is None:
        raise ValueError("target paragraph has no w:sectPr")
    sect_xml = sect_match.group(0)
    sect_xml = re.sub(rb"<w:type\b[^>]*/>", b"", sect_xml, flags=re.DOTALL)
    sect_xml = re.sub(rb"<w:type\b[^>]*>.*?</w:type>", b"", sect_xml, flags=re.DOTALL)
    type_xml = f'<w:type w:val="{section_type}"/>'.encode("utf-8")
    header_footer_matches = list(re.finditer(rb"<w:(?:headerReference|footerReference)\b[^>]*/>", sect_xml))
    if header_footer_matches:
        insert_at = header_footer_matches[-1].end()
    else:
        open_match = re.match(rb"<w:sectPr\b[^>]*>", sect_xml)
        if open_match is None:
            raise ValueError("section properties XML start tag not found")
        insert_at = open_match.end()
    sect_xml = sect_xml[:insert_at] + type_xml + sect_xml[insert_at:]
    return paragraph_xml[: sect_match.start()] + sect_xml + paragraph_xml[sect_match.end() :]


def apply_package_preserving_frontmatter_section_type_repairs(
    input_docx: Path, output_docx: Path, plan: dict[str, object]
) -> dict[str, object]:
    items = plan.get("frontmatter_section_type_repairs", [])
    if not isinstance(items, list) or not items:
        raise ValueError("package_preserving_frontmatter_section_types requires frontmatter_section_type_repairs")
    allowed_truthy = {"package_preserving_frontmatter_section_types"}
    allowed_nonempty = {"frontmatter_section_type_repairs"}
    for key, value in plan.items():
        if key in allowed_nonempty:
            continue
        if key in allowed_truthy and bool(value):
            continue
        if value not in (False, None, "", [], {}):
            raise ValueError(
                "package-preserving frontmatter section repair may not be combined with other mutation plan fields: "
                f"{key}"
            )

    with zipfile.ZipFile(input_docx) as zin:
        document_xml = zin.read("word/document.xml")
        paragraph_matches = list(P_XML_RE.finditer(document_xml))
        if not paragraph_matches:
            raise ValueError("word/document.xml has no w:p paragraphs")
        replacements: dict[int, bytes] = {}
        changed: list[dict[str, object]] = []
        used_indices: set[int] = set()
        paragraph_texts = [paragraph_text_from_xml_bytes(match.group(0)).strip() for match in paragraph_matches]
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("each frontmatter_section_type_repairs item must be an object")
            exact_text = str(item.get("paragraph_text_exact") or "").strip()
            startswith = str(item.get("paragraph_text_startswith") or "").strip()
            section_type = str(item.get("section_type") or "").strip()
            min_paragraph_index = int(item.get("min_paragraph_index") or 0)
            max_paragraph_index = int(item.get("max_paragraph_index") or len(paragraph_matches))
            if not section_type or (not exact_text and not startswith):
                raise ValueError("frontmatter section repair requires section_type and exact/startswith text")
            matches = []
            for index, text in enumerate(paragraph_texts):
                if index < min_paragraph_index or index > max_paragraph_index or index in used_indices:
                    continue
                if exact_text and text == exact_text:
                    matches.append(index)
                elif startswith and text.startswith(startswith):
                    matches.append(index)
            if len(matches) != 1:
                display = exact_text or startswith
                raise ValueError(f"frontmatter section repair matched {len(matches)} paragraphs: {display[:40]}")
            index = matches[0]
            before_xml = paragraph_matches[index].group(0)
            patched = set_paragraph_section_type(before_xml, section_type)
            replacements[index] = patched
            used_indices.add(index)
            changed.append(
                {
                    "target_paragraph_index": index,
                    "paragraph_text": paragraph_texts[index][:120],
                    "section_type": section_type,
                    "patch_mode": "package_preserving_document_xml_span",
                }
            )

        rebuilt = bytearray()
        cursor = 0
        for index, match in enumerate(paragraph_matches):
            rebuilt.extend(document_xml[cursor : match.start()])
            rebuilt.extend(replacements.get(index, match.group(0)))
            cursor = match.end()
        rebuilt.extend(document_xml[cursor:])

        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for zip_item in zin.infolist():
                    data = zin.read(zip_item.filename)
                    if zip_item.filename == "word/document.xml":
                        data = bytes(rebuilt)
                    zout.writestr(zip_item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)

    return {
        "package_preserving_frontmatter_section_types": changed,
        "document_xml_bytes_before": len(document_xml),
        "document_xml_bytes_after": len(rebuilt),
    }


def apply_package_preserving_settings_repairs(
    input_docx: Path, output_docx: Path, plan: dict[str, object]
) -> dict[str, object]:
    if not bool(plan.get("disable_even_and_odd_headers", False)):
        raise ValueError("package_preserving_settings_repair currently requires disable_even_and_odd_headers=true")
    allowed_truthy = {"package_preserving_settings_repair", "disable_even_and_odd_headers"}
    for key, value in plan.items():
        if key in allowed_truthy and bool(value):
            continue
        if value not in (False, None, "", [], {}):
            raise ValueError(
                "package-preserving settings repair may not be combined with other mutation plan fields: "
                f"{key}"
            )
    with zipfile.ZipFile(input_docx) as zin:
        if "word/settings.xml" not in zin.namelist():
            raise ValueError("word/settings.xml is missing")
        settings_xml = zin.read("word/settings.xml")
        original_settings = settings_xml
        settings_xml = re.sub(
            rb"<w:evenAndOddHeaders\b[^>]*/>",
            b'<w:evenAndOddHeaders w:val="0" />',
            settings_xml,
        )
        if settings_xml == original_settings:
            raise ValueError("word/settings.xml did not contain w:evenAndOddHeaders to disable")
        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for zip_item in zin.infolist():
                    data = zin.read(zip_item.filename)
                    if zip_item.filename == "word/settings.xml":
                        data = settings_xml
                    zout.writestr(zip_item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)
    return {
        "package_preserving_settings_repair": {
            "changed_part": "word/settings.xml",
            "disable_even_and_odd_headers": True,
            "reason": "prevent WPS from inserting blank parity pages between front-matter sections",
        }
    }


def apply_package_preserving_explicit_heading_number_repairs(
    input_docx: Path, output_docx: Path, plan: dict[str, object]
) -> dict[str, object]:
    items = plan.get("explicit_body_heading_numbers", [])
    if not isinstance(items, list) or not items:
        raise ValueError("package_preserving_explicit_heading_numbers requires explicit_body_heading_numbers")
    allowed_truthy = {"package_preserving_explicit_heading_numbers"}
    allowed_nonempty = {"explicit_body_heading_numbers"}
    for key, value in plan.items():
        if key in allowed_nonempty:
            continue
        if key in allowed_truthy and bool(value):
            continue
        if value not in (False, None, "", [], {}):
            raise ValueError(
                "package-preserving heading repair may not be combined with other mutation plan fields: "
                f"{key}"
            )

    with zipfile.ZipFile(input_docx) as zin:
        document_xml = zin.read("word/document.xml")
        paragraph_matches = list(P_XML_RE.finditer(document_xml))
        if not paragraph_matches:
            raise ValueError("word/document.xml has no w:p paragraphs")
        replacements: dict[int, bytes] = {}
        changed: list[dict[str, object]] = []
        used_indices: set[int] = set()
        paragraph_texts = [paragraph_text_from_xml_bytes(match.group(0)).strip() for match in paragraph_matches]
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("each explicit_body_heading_numbers item must be an object")
            exact_text = str(item.get("heading_text_exact") or "").strip()
            number_text = str(item.get("number_text") or "").strip()
            min_paragraph_index = int(item.get("min_paragraph_index") or 0)
            if not exact_text or not number_text:
                raise ValueError("explicit heading number repair requires heading_text_exact and number_text")
            matches = [
                index
                for index, text in enumerate(paragraph_texts)
                if index >= min_paragraph_index and index not in used_indices and text == exact_text
            ]
            if len(matches) != 1:
                raise ValueError(f"explicit heading number repair matched {len(matches)} paragraphs: {exact_text[:40]}")
            index = matches[0]
            patched, before_text, after_text = prepend_heading_number_in_paragraph_xml(
                paragraph_matches[index].group(0), number_text
            )
            replacements[index] = patched
            used_indices.add(index)
            changed.append(
                {
                    "target_paragraph_index": index,
                    "heading_text_exact": exact_text,
                    "number_text": number_text,
                    "before_text": before_text,
                    "after_text": after_text,
                    "auto_numbering_suppressed": True,
                    "patch_mode": "package_preserving_document_xml_span",
                }
            )

        rebuilt = bytearray()
        cursor = 0
        for index, match in enumerate(paragraph_matches):
            rebuilt.extend(document_xml[cursor : match.start()])
            rebuilt.extend(replacements.get(index, match.group(0)))
            cursor = match.end()
        rebuilt.extend(document_xml[cursor:])

        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for zip_item in zin.infolist():
                    data = zin.read(zip_item.filename)
                    if zip_item.filename == "word/document.xml":
                        data = bytes(rebuilt)
                    zout.writestr(zip_item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)

    return {
        "package_preserving_explicit_heading_numbers": changed,
        "document_xml_bytes_before": len(document_xml),
        "document_xml_bytes_after": len(rebuilt),
    }


def remove_outline_level_from_non_heading_paragraph_xml(paragraph_xml: bytes, paragraph_text_value: str) -> bytes:
    if is_heading_text(paragraph_text_value):
        raise ValueError("outline removal target text is a heading-like paragraph")
    ppr_match = re.search(rb"<w:pPr\b[^>]*>.*?</w:pPr>", paragraph_xml, flags=re.DOTALL)
    if ppr_match is None:
        raise ValueError("outline removal target has no paragraph properties")
    ppr_xml = ppr_match.group(0)
    if not re.search(rb"<w:outlineLvl\b[^>]*/>", ppr_xml) and not re.search(
        rb"<w:outlineLvl\b[^>]*>.*?</w:outlineLvl>", ppr_xml, flags=re.DOTALL
    ):
        raise ValueError("outline removal target has no w:outlineLvl")
    ppr_xml = re.sub(rb"<w:outlineLvl\b[^>]*/>", b"", ppr_xml)
    ppr_xml = re.sub(rb"<w:outlineLvl\b[^>]*>.*?</w:outlineLvl>", b"", ppr_xml, flags=re.DOTALL)
    return paragraph_xml[: ppr_match.start()] + ppr_xml + paragraph_xml[ppr_match.end() :]


def apply_package_preserving_non_heading_outline_repairs(
    input_docx: Path, output_docx: Path, plan: dict[str, object]
) -> dict[str, object]:
    items = plan.get("non_heading_outline_removals", [])
    if not isinstance(items, list) or not items:
        raise ValueError("package_preserving_non_heading_outline_removal requires non_heading_outline_removals")
    allowed_truthy = {"package_preserving_non_heading_outline_removal"}
    allowed_nonempty = {"non_heading_outline_removals"}
    for key, value in plan.items():
        if key in allowed_nonempty:
            continue
        if key in allowed_truthy and bool(value):
            continue
        if value not in (False, None, "", [], {}):
            raise ValueError(
                "package-preserving non-heading outline repair may not be combined with other mutation plan fields: "
                f"{key}"
            )

    with zipfile.ZipFile(input_docx) as zin:
        document_xml = zin.read("word/document.xml")
        paragraph_matches = list(P_XML_RE.finditer(document_xml))
        if not paragraph_matches:
            raise ValueError("word/document.xml has no w:p paragraphs")
        replacements: dict[int, bytes] = {}
        changed: list[dict[str, object]] = []
        used_indices: set[int] = set()
        paragraph_texts = [paragraph_text_from_xml_bytes(match.group(0)).strip() for match in paragraph_matches]
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("each non_heading_outline_removals item must be an object")
            exact_text = str(item.get("paragraph_text_exact") or "").strip()
            startswith = str(item.get("paragraph_text_startswith") or "").strip()
            min_paragraph_index = int(item.get("min_paragraph_index") or 0)
            max_paragraph_index = int(item.get("max_paragraph_index") or len(paragraph_matches))
            if not exact_text and not startswith:
                raise ValueError("non-heading outline repair requires exact or startswith text")
            matches = []
            for index, text in enumerate(paragraph_texts):
                if index < min_paragraph_index or index > max_paragraph_index or index in used_indices:
                    continue
                if exact_text and text == exact_text:
                    matches.append(index)
                elif startswith and text.startswith(startswith):
                    matches.append(index)
            if len(matches) != 1:
                display = exact_text or startswith
                raise ValueError(f"non-heading outline repair matched {len(matches)} paragraphs: {display[:40]}")
            index = matches[0]
            before_xml = paragraph_matches[index].group(0)
            after_xml = remove_outline_level_from_non_heading_paragraph_xml(before_xml, paragraph_texts[index])
            replacements[index] = after_xml
            used_indices.add(index)
            changed.append(
                {
                    "target_paragraph_index": index,
                    "paragraph_text": paragraph_texts[index][:160],
                    "removed": "w:outlineLvl",
                    "patch_mode": "package_preserving_document_xml_span",
                }
            )

        rebuilt = bytearray()
        cursor = 0
        for index, match in enumerate(paragraph_matches):
            rebuilt.extend(document_xml[cursor : match.start()])
            rebuilt.extend(replacements.get(index, match.group(0)))
            cursor = match.end()
        rebuilt.extend(document_xml[cursor:])

        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for zip_item in zin.infolist():
                    data = zin.read(zip_item.filename)
                    if zip_item.filename == "word/document.xml":
                        data = bytes(rebuilt)
                    zout.writestr(zip_item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)

    return {
        "package_preserving_non_heading_outline_removal": changed,
        "document_xml_bytes_before": len(document_xml),
        "document_xml_bytes_after": len(rebuilt),
    }


def remove_blank_paragraphs_between_abstract_body_and_keywords(
    body: ET.Element, paragraphs: list[ET.Element], plan: dict[str, object]
) -> list[dict[str, object]]:
    if not bool(plan.get("remove_blank_paragraphs_between_abstract_body_and_keywords", False)):
        return []
    surfaces = find_abstract_surfaces(paragraphs)
    changed: list[dict[str, object]] = []
    for body_surface, keyword_surface in (
        ("zh_abstract_body", "zh_keyword_line"),
        ("en_abstract_body", "en_keyword_line"),
    ):
        body_index = surfaces[body_surface][0]
        keyword_index = surfaces[keyword_surface][0]
        if keyword_index <= body_index:
            raise ValueError(f"{keyword_surface} appears before {body_surface}")
        removed: list[int] = []
        for index in range(keyword_index - 1, body_index, -1):
            text = paragraph_text(paragraphs[index])
            if text.strip():
                continue
            if paragraphs[index].find(".//w:drawing", NS) is not None or paragraphs[index].find(".//w:pict", NS) is not None:
                raise ValueError(f"blank abstract seam paragraph {index} contains a drawing or pict object")
            if paragraphs[index].find(".//w:fldChar", NS) is not None or paragraphs[index].find(".//w:instrText", NS) is not None:
                raise ValueError(f"blank abstract seam paragraph {index} contains a field")
            body.remove(paragraphs[index])
            paragraphs.pop(index)
            removed.append(index)
        if removed:
            changed.append(
                {
                    "body_surface": body_surface,
                    "keyword_surface": keyword_surface,
                    "removed_blank_paragraph_indices": list(reversed(removed)),
                }
            )
    return changed


def apply_repairs(input_docx: Path, template_docx: Path, output_docx: Path, plan: dict[str, object]) -> dict[str, object]:
    if bool(plan.get("package_preserving_settings_repair", False)):
        return apply_package_preserving_settings_repairs(input_docx, output_docx, plan)
    if bool(plan.get("package_preserving_body_mixed_script_font_slots", False)):
        return apply_body_mixed_script_font_slot_repairs(input_docx, output_docx, plan)
    if bool(plan.get("package_preserving_frontmatter_section_types", False)):
        return apply_package_preserving_frontmatter_section_type_repairs(input_docx, output_docx, plan)
    if bool(plan.get("package_preserving_explicit_heading_numbers", False)):
        return apply_package_preserving_explicit_heading_number_repairs(input_docx, output_docx, plan)
    if bool(plan.get("package_preserving_non_heading_outline_removal", False)):
        return apply_package_preserving_non_heading_outline_repairs(input_docx, output_docx, plan)
    if plan.get("image_display_resize") not in (None, {}, []):
        raise ValueError(
            "plan.image_display_resize requires a transaction-owned figure manifest; "
            "repair_thesis_surface_format.py may not independently mutate drawing extents"
        )
    keyword_only = bool(plan.get("keyword_only", False))
    repair_abstract_surfaces = bool(plan.get("repair_abstract_surfaces", False if keyword_only else True))
    if keyword_only:
        template_paragraphs: list[ET.Element] = []
        template_style_names: dict[str, str] = {}
        template_normal_style_id = "Normal"
        template_surfaces: dict[str, tuple[int, ET.Element]] = {}
    else:
        with zipfile.ZipFile(template_docx) as template_zip:
            template_root, _template_body, template_paragraphs = parse_document(template_zip.read("word/document.xml"))
            try:
                template_styles_xml = template_zip.read("word/styles.xml")
            except KeyError:
                template_styles_xml = b""
                template_style_names = {}
                template_normal_style_id = "Normal"
            else:
                template_style_names = paragraph_style_id_to_name(template_styles_xml)
            if template_styles_xml:
                try:
                    template_normal_style_id = normal_style_id(template_styles_xml)
                except KeyError:
                    template_normal_style_id = "Normal"
        del template_root
        template_surfaces = find_abstract_surfaces(template_paragraphs) if repair_abstract_surfaces else {}
    normalize_body_prose = bool(plan.get("normalize_body_prose", False if keyword_only else True))
    normalize_image_holders = bool(plan.get("normalize_image_holders", False if keyword_only else normalize_body_prose))
    normalize_body_headings = bool(plan.get("normalize_body_headings", False))
    sync_normal_style_baseline = bool(plan.get("sync_normal_style_baseline", True))
    remove_annotation_colors = bool(plan.get("remove_annotation_run_colors", False))
    force_real_firstline_twips = (
        int(plan["force_real_firstline_twips"])
        if str(plan.get("force_real_firstline_twips", "")).strip()
        else None
    )
    keyword_firstline_twips = (
        int(plan["keyword_firstline_twips"])
        if str(plan.get("keyword_firstline_twips", "")).strip()
        else None
    )
    body_donor_index: int | None = None
    body_ppr: ET.Element | None = None
    body_rpr: ET.Element | None = None
    body_style_policy = str(plan.get("body_style_policy") or "normal").strip().lower()
    if body_style_policy not in {"normal", "template-donor"}:
        raise ValueError("plan.body_style_policy must be 'normal' or 'template-donor'")
    body_style_id_override = str(plan.get("body_style_id") or "").strip()
    if normalize_body_prose or normalize_image_holders:
        body_donor_index, body_donor = find_template_body_donor(template_paragraphs)
        body_ppr = deepcopy(paragraph_property(body_donor))
        if body_ppr is None:
            raise ValueError("body donor has no pPr")
        if body_style_policy == "normal":
            set_style(body_ppr, template_normal_style_id)
        body_rpr = first_run_rpr(body_donor)

    with zipfile.ZipFile(input_docx) as zin:
        document_bytes = zin.read("word/document.xml")
        try:
            target_styles_xml = zin.read("word/styles.xml")
            target_style_ids = paragraph_style_name_to_id(target_styles_xml)
        except KeyError:
            target_styles_xml = b""
            target_style_ids = {}
        target_normal_style_id = target_style_ids.get("normal", template_normal_style_id)
        target_image_holder_style_id = target_normal_style_id
        image_holder_style_changed = False
        if normalize_image_holders and target_styles_xml:
            target_styles_xml, target_image_holder_style_id, image_holder_style_changed = ensure_image_holder_style_xml(target_styles_xml)
        if body_ppr is not None:
            if body_style_policy == "template-donor":
                body_ppr = remap_ppr_style_id(body_ppr, template_style_names, target_style_ids)
            else:
                set_style(body_ppr, target_normal_style_id)
            if body_style_id_override:
                if target_styles_xml and style_element_by_id(ET.fromstring(target_styles_xml), body_style_id_override) is None:
                    raise ValueError(f"plan.body_style_id not found in target styles: {body_style_id_override}")
                set_style(body_ppr, body_style_id_override)
        if not keyword_only and sync_normal_style_baseline and target_styles_xml and template_styles_xml:
            target_styles_xml = sync_normal_style_font_baseline(
                target_styles_xml,
                template_styles_xml,
                target_normal_style_id=target_normal_style_id,
                template_normal_style_id=template_normal_style_id,
            )
        root, body, paragraphs = parse_document(document_bytes)
        target_surfaces = (
            {}
            if keyword_only or not repair_abstract_surfaces
            else find_abstract_surfaces(paragraphs)
        )

        changed: dict[str, object] = {
            "abstract_surfaces": {},
            "body_prose_paragraphs_normalized": [],
            "body_headings_normalized": [],
            "body_text_replacements": [],
            "explicit_body_heading_numbers": [],
            "empty_paragraphs_removed_before_headings": [],
            "redundant_empty_page_breaks_removed_before_headings": [],
            "redundant_page_break_before_removed_from_headings": [],
            "blank_paragraphs_removed_between_abstract_body_and_keywords": [],
            "image_holders_normalized": [],
            "body_donor_template_index": body_donor_index,
            "template_normal_style_id": template_normal_style_id,
            "normalize_body_prose": normalize_body_prose,
            "normalize_image_holders": normalize_image_holders,
            "normalize_body_headings": normalize_body_headings,
            "repair_abstract_surfaces": repair_abstract_surfaces,
            "keyword_only": keyword_only,
            "body_style_policy": body_style_policy,
            "body_style_id_override": body_style_id_override,
            "sync_normal_style_baseline": sync_normal_style_baseline,
            "remove_annotation_run_colors": remove_annotation_colors,
            "annotation_run_visual_markers_removed": 0,
            "image_holder_style_definition": {
                "style_id": target_image_holder_style_id,
                "changed": image_holder_style_changed,
            },
        }

        abstract_text = plan.get("abstract_text", {})
        if not isinstance(abstract_text, dict):
            raise ValueError("plan.abstract_text must be an object")
        if keyword_only:
            changed["abstract_surfaces"] = repair_keyword_only_surfaces(
                paragraphs,
                abstract_text,
                allow_missing=bool(plan.get("allow_missing_keyword_surface", False)),
            )
            if bool(plan.get("repair_abstract_role_format", False)):
                changed["abstract_surfaces"].update(repair_abstract_role_surfaces(paragraphs))
            normalize_body_prose = False
            normalize_image_holders = False
            normalize_body_headings = False
            repair_abstract_surfaces = False
            changed["normalize_body_prose"] = normalize_body_prose
            changed["normalize_image_holders"] = normalize_image_holders
            changed["normalize_body_headings"] = normalize_body_headings
            changed["repair_abstract_surfaces"] = repair_abstract_surfaces
        elif bool(plan.get("repair_inline_abstract_surfaces", False)):
            changed["abstract_surfaces"] = repair_inline_abstract_surfaces_from_template(
                target_surfaces,
                template_surfaces,
                abstract_text,
                template_style_names=template_style_names,
                target_style_ids=target_style_ids,
            )
            normalize_body_prose = False
            normalize_image_holders = False
            normalize_body_headings = False
            repair_abstract_surfaces = False
            changed["normalize_body_prose"] = normalize_body_prose
            changed["normalize_image_holders"] = normalize_image_holders
            changed["normalize_body_headings"] = normalize_body_headings
            changed["repair_abstract_surfaces"] = repair_abstract_surfaces
        elif repair_abstract_surfaces:
            for surface_id in (
                "zh_abstract_title",
                "zh_abstract_body",
                "zh_keyword_line",
                "en_abstract_title",
                "en_abstract_body",
                "en_keyword_line",
            ):
                target_index, target_para = target_surfaces[surface_id]
                _template_index, template_para = template_surfaces[surface_id]
                donor_ppr = remap_ppr_style_id(paragraph_property(template_para), template_style_names, target_style_ids)
                donor_rpr = (
                    abstract_title_marker_rpr(template_para, surface_id)
                    if surface_id.endswith("title")
                    else first_run_rpr(template_para)
                )
                donor_source = "template"
                keyword_role_rprs: tuple[ET.Element | None, ET.Element | None, ET.Element | None] | None = None
                if surface_id.endswith("keyword_line") and is_instruction_like_abstract_donor(template_para, template_style_names):
                    body_surface = "zh_abstract_body" if surface_id == "zh_keyword_line" else "en_abstract_body"
                    body_para = target_surfaces[body_surface][1]
                    body_para_ppr = paragraph_property(body_para)
                    body_para_rpr = first_run_rpr(body_para)
                    donor_ppr = deepcopy(body_para_ppr) if body_para_ppr is not None else deepcopy(body_ppr) if body_ppr is not None else None
                    donor_rpr = deepcopy(body_para_rpr) if body_para_rpr is not None else deepcopy(body_rpr) if body_rpr is not None else None
                    keyword_role_rprs = (
                        strong_label_rpr(donor_rpr),
                        strong_label_rpr(donor_rpr),
                        deepcopy(donor_rpr) if donor_rpr is not None else None,
                    )
                    donor_source = "target_abstract_body_fallback"
                if donor_ppr is not None:
                    replace_ppr(target_para, donor_ppr)
                if surface_id in abstract_text:
                    if surface_id.endswith("keyword_line"):
                        if keyword_role_rprs is not None:
                            set_keyword_line_text_with_roles(
                                target_para,
                                str(abstract_text[surface_id]),
                                label_rpr=keyword_role_rprs[0],
                                separator_rpr=keyword_role_rprs[1],
                                content_rpr=keyword_role_rprs[2],
                            )
                        else:
                            set_keyword_line_text(target_para, str(abstract_text[surface_id]), template_para)
                    else:
                        set_paragraph_text(target_para, str(abstract_text[surface_id]), donor_rpr)
                elif surface_id.endswith("keyword_line"):
                    if keyword_role_rprs is not None:
                        set_keyword_line_text_with_roles(
                            target_para,
                            paragraph_text(target_para),
                            label_rpr=keyword_role_rprs[0],
                            separator_rpr=keyword_role_rprs[1],
                            content_rpr=keyword_role_rprs[2],
                        )
                    else:
                        set_keyword_line_text(target_para, paragraph_text(target_para), template_para)
                else:
                    clear_body_run_formatting(target_para, donor_rpr)
                if surface_id in {"zh_abstract_body", "en_abstract_body"}:
                    force_visible_spacing_and_firstline(target_para, firstline_twips=force_real_firstline_twips)
                    apply_latin_font_slots(target_para)
                    split_mixed_script_runs(target_para)
                elif surface_id in {"zh_keyword_line", "en_keyword_line"} and keyword_firstline_twips is not None:
                    force_visible_spacing_and_firstline(target_para, firstline_twips=keyword_firstline_twips)
                changed["abstract_surfaces"][surface_id] = {
                    "target_paragraph_index": target_index,
                    "template_text": paragraph_text(template_para)[:80],
                    "final_text": paragraph_text(target_para)[:120],
                    "run_formatting_preserved_when_text_unchanged": surface_id not in abstract_text,
                    "donor_source": donor_source,
                }

            for body_surface, keyword_surface in (
                ("zh_abstract_body", "zh_keyword_line"),
                ("en_abstract_body", "en_keyword_line"),
            ):
                template_body_para = template_surfaces[body_surface][1]
                template_body_ppr = remap_ppr_style_id(paragraph_property(template_body_para), template_style_names, target_style_ids)
                template_body_rpr = first_run_rpr(template_body_para)
                start = target_surfaces[body_surface][0]
                end = target_surfaces[keyword_surface][0]
                for body_index in range(start + 1, end):
                    body_para = paragraphs[body_index]
                    if paragraph_text(body_para).strip():
                        normalize_paragraph_to_donor(body_para, template_body_ppr, template_body_rpr)
                        force_visible_spacing_and_firstline(body_para, firstline_twips=force_real_firstline_twips)
                        apply_latin_font_slots(body_para)
                        split_mixed_script_runs(body_para)
                        changed["abstract_surfaces"][f"{body_surface}_paragraph_{body_index}"] = {
                            "target_paragraph_index": body_index,
                            "template_text": paragraph_text(template_body_para)[:80],
                            "final_text": paragraph_text(body_para)[:120],
                        }
            for role_key, role_change in repair_abstract_role_surfaces(paragraphs).items():
                changed["abstract_surfaces"][f"{role_key}_role_format"] = role_change

        if normalize_body_prose or normalize_image_holders or normalize_body_headings:
            if "en_keyword_line" in target_surfaces:
                body_start, _ = find_first(
                    paragraphs,
                    lambda i, p: i > target_surfaces["en_keyword_line"][0]
                    and not re.search(r"\d\s*$", paragraph_text(p).strip())
                    and (
                        is_body_heading_paragraph(p)
                        or re.match(r"^\s*1\s+", paragraph_text(p)) is not None
                        or re.match(r"^\s*\u7b2c\s*1\s*\u7ae0\s+\S", paragraph_text(p)) is not None
                    ),
                )
            else:
                toc_index = next(
                    (
                        i
                        for i, p in enumerate(paragraphs)
                        if compact_text(paragraph_text(p)) == "\u76ee\u5f55"
                    ),
                    -1,
                )
                body_start, _ = find_first(
                    paragraphs,
                    lambda i, p: i > toc_index
                    and not re.search(r"\d\s*$", paragraph_text(p).strip())
                    and is_body_heading_paragraph(p),
                )
            references_index = next(
                (i for i, p in enumerate(paragraphs) if compact_text(paragraph_text(p)) == "参考文献"),
                len(paragraphs),
            )
            acknowledgement_index = next(
                (i for i, p in enumerate(paragraphs) if compact_text(paragraph_text(p)) in {"谢辞", "致谢"}),
                references_index,
            )
            if references_index == len(paragraphs):
                references_index = next(
                    (
                        i
                        for i, p in enumerate(paragraphs)
                        if compact_text(paragraph_text(p)) == "\u53c2\u8003\u6587\u732e"
                    ),
                    len(paragraphs),
                )
            if acknowledgement_index == references_index:
                acknowledgement_index = next(
                    (
                        i
                        for i, p in enumerate(paragraphs)
                        if compact_text(paragraph_text(p)) in {"\u81f4\u8c22", "\u8c22\u8f9e"}
                    ),
                    references_index,
                )
            body_end = min(references_index, acknowledgement_index)
            for index, paragraph in enumerate(paragraphs[body_start:body_end], start=body_start):
                text = paragraph_text(paragraph).strip()
                has_drawing = paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None
                if has_drawing:
                    if normalize_image_holders:
                        apply_image_holder_ppr(paragraph, style_id=target_image_holder_style_id)
                        changed["image_holders_normalized"].append(index)
                    continue
                if text and is_body_heading_paragraph(paragraph):
                    if normalize_body_headings:
                        heading_change = force_body_heading_format(paragraph, text)
                        if heading_change is not None:
                            changed["body_headings_normalized"].append(
                                {
                                    "paragraph_index": index,
                                    "text": text[:120],
                                    **heading_change,
                                }
                            )
                    continue
                if not normalize_body_prose or not text:
                    continue
                if is_caption_or_reference(text) or is_code_like_paragraph_text(text):
                    continue
                normalize_paragraph_to_donor(paragraph, body_ppr, body_rpr)
                force_visible_spacing_and_firstline(paragraph, firstline_twips=force_real_firstline_twips)
                apply_latin_font_slots(paragraph)
                split_mixed_script_runs(paragraph)
                changed["body_prose_paragraphs_normalized"].append(index)

            if normalize_body_prose and acknowledgement_index < references_index:
                for index, paragraph in enumerate(paragraphs):
                    if compact_text(paragraph_text(paragraph)) in {"谢辞", "致谢"}:
                        continue
                    text = paragraph_text(paragraph).strip()
                    if acknowledgement_index < index < references_index and text and not is_code_like_paragraph_text(text):
                        normalize_paragraph_to_donor(paragraph, body_ppr, body_rpr)
                        force_visible_spacing_and_firstline(paragraph, firstline_twips=force_real_firstline_twips)
                        changed["body_prose_paragraphs_normalized"].append(index)

        changed["body_text_replacements"] = apply_body_text_replacements(
            paragraphs,
            plan,
            body_ppr=body_ppr,
            body_rpr=body_rpr,
            firstline_twips=force_real_firstline_twips,
        )
        changed["explicit_body_heading_numbers"] = apply_explicit_body_heading_numbers(paragraphs, plan)
        changed["blank_paragraphs_removed_between_abstract_body_and_keywords"] = (
            remove_blank_paragraphs_between_abstract_body_and_keywords(body, paragraphs, plan)
        )
        changed["empty_paragraphs_removed_before_headings"] = remove_empty_paragraphs_before_headings(body, paragraphs, plan)
        changed["redundant_empty_page_breaks_removed_before_headings"] = (
            remove_redundant_empty_page_breaks_before_headings(body, paragraphs, plan)
        )
        changed["redundant_page_break_before_removed_from_headings"] = (
            remove_redundant_page_break_before_from_headings(paragraphs, plan)
        )
        changed["image_display_resizes"] = apply_image_display_resize(paragraphs, plan.get("image_display_resize"))
        if remove_annotation_colors:
            changed["annotation_run_visual_markers_removed"] = remove_document_annotation_run_colors(paragraphs)

        with tempfile.TemporaryDirectory() as td:
            temp_output = Path(td) / "out.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    elif item.filename == "word/styles.xml" and target_styles_xml:
                        data = target_styles_xml
                    zout.writestr(item, data)
            output_docx.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(temp_output, output_docx)

    return changed


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply bounded thesis surface-format repair.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--plan", required=True, help="JSON plan containing abstract_text overrides.")
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    input_docx = Path(args.input_docx).resolve()
    template_docx = Path(args.template_docx).resolve()
    output_docx = Path(args.output_docx).resolve()
    plan_path = Path(args.plan).resolve()
    report_path = Path(args.report).resolve()

    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    before_sha = sha256_file(input_docx)
    template_sha = sha256_file(template_docx)
    changes = apply_repairs(input_docx, template_docx, output_docx, plan)
    after_sha = sha256_file(output_docx)

    report = {
        "schema": "graduation-project-builder.repair-thesis-surface-format.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "generator": "scripts/repair_thesis_surface_format.py",
        "input_docx": str(input_docx),
        "input_sha256": before_sha,
        "template_docx": str(template_docx),
        "template_sha256": template_sha,
        "output_docx": str(output_docx),
        "output_sha256": after_sha,
        "plan": str(plan_path),
        "changes": changes,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raise SystemExit(main())
