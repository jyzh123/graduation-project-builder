#!/usr/bin/env python3
"""Repair narrow final-surface blockers reported by the thesis self-check."""

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
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_of(node: ET.Element) -> str:
    return "".join(child.text or "" for child in node.findall(".//w:t", NS)).strip()


def normalize(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").lower()


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is not None:
        return ppr
    ppr = ET.Element(W + "pPr")
    paragraph.insert(0, ppr)
    return ppr


def ensure_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("./w:rPr", NS)
    if rpr is not None:
        return rpr
    rpr = ET.Element(W + "rPr")
    run.insert(0, rpr)
    return rpr


def remove_child(parent: ET.Element | None, tag: str) -> bool:
    if parent is None:
        return False
    changed = False
    for child in list(parent):
        if child.tag == tag:
            parent.remove(child)
            changed = True
    return changed


def set_child_val(parent: ET.Element, tag: str, value: str) -> bool:
    child = parent.find(f"./w:{tag.removeprefix(W)}", NS)
    changed = False
    if child is None:
        child = ET.Element(tag)
        parent.append(child)
        changed = True
    if child.attrib.get(W + "val") != value:
        child.set(W + "val", value)
        changed = True
    return changed


def ensure_child(parent: ET.Element, tag: str) -> tuple[ET.Element, bool]:
    child = parent.find(f"./w:{tag.removeprefix(W)}", NS)
    if child is not None:
        return child, False
    child = ET.Element(tag)
    parent.append(child)
    return child, True


def paragraph_text_has_chapter_heading(text: str) -> bool:
    return re.match(r"^\s*\u7b2c\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\s*\u7ae0", text or "") is not None


def is_tail_heading_text(text: str) -> bool:
    normalized = normalize(text)
    return normalized in {
        normalize("\u53c2\u8003\u6587\u732e"),
        normalize("\u81f4\u8c22"),
        normalize("\u9644\u5f55"),
    }


def is_table_caption_text(text: str) -> bool:
    stripped = (text or "").strip()
    if stripped.startswith("\u7eed\u8868"):
        return True
    return re.match(r"^\s*\u8868\s*\d+(?:[-.\uff0d\uff0e]\d+)?", stripped) is not None


def ensure_rfonts(rpr: ET.Element) -> tuple[ET.Element, bool]:
    rfonts = rpr.find("./w:rFonts", NS)
    if rfonts is not None:
        return rfonts, False
    rfonts = ET.Element(W + "rFonts")
    rpr.insert(0, rfonts)
    return rfonts, True


def set_run_size(run: ET.Element, half_points: str) -> bool:
    changed = False
    rpr = ensure_rpr(run)
    if set_child_val(rpr, W + "sz", half_points):
        changed = True
    if set_child_val(rpr, W + "szCs", half_points):
        changed = True
    return changed


def set_run_font_slots(run: ET.Element, *, east_asia: str | None = None, ascii_font: str | None = None, hansi: str | None = None, cs: str | None = None) -> bool:
    changed = False
    rpr = ensure_rpr(run)
    rfonts, created = ensure_rfonts(rpr)
    changed = changed or created
    values = {
        "eastAsia": east_asia,
        "ascii": ascii_font,
        "hAnsi": hansi,
        "cs": cs,
    }
    for key, value in values.items():
        if value is None:
            continue
        attr = W + key
        if rfonts.attrib.get(attr) != value:
            rfonts.set(attr, value)
            changed = True
    return changed


def split_script_segments(text: str) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    current_kind = ""
    current_text: list[str] = []

    def char_kind(ch: str) -> str:
        if "\u4e00" <= ch <= "\u9fff":
            return "cjk"
        if ch.isascii() and (ch.isalnum() or ch in {"-", ".", "_", "/", ":", "(", ")", "[", "]"}):
            return "latin"
        return current_kind or "cjk"

    for ch in text:
        kind = char_kind(ch)
        if current_text and kind != current_kind:
            segments.append((current_kind, "".join(current_text)))
            current_text = []
        current_kind = kind
        current_text.append(ch)
    if current_text:
        segments.append((current_kind or "cjk", "".join(current_text)))
    return segments


def clone_text_run(source_run: ET.Element, text: str, *, latin_font: bool) -> ET.Element:
    new_run = ET.Element(W + "r")
    source_rpr = source_run.find("./w:rPr", NS)
    if source_rpr is not None:
        new_run.append(copy.deepcopy(source_rpr))
    t = ET.SubElement(new_run, W + "t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    if latin_font:
        set_run_font_slots(
            new_run,
            east_asia="Times New Roman",
            ascii_font="Times New Roman",
            hansi="Times New Roman",
            cs="Times New Roman",
        )
    return new_run


def split_caption_runs_for_latin_slots(paragraph: ET.Element) -> int:
    replacements = 0
    for run in list(paragraph.findall("w:r", NS)):
        text_nodes = run.findall("w:t", NS)
        if len(text_nodes) != 1:
            text = "".join(node.text or "" for node in text_nodes)
        else:
            text = text_nodes[0].text or ""
        if not text:
            continue
        has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in text)
        has_latin = any(ch.isascii() and ch.isalnum() for ch in text)
        if has_latin and not has_cjk:
            if set_run_font_slots(
                run,
                east_asia="Times New Roman",
                ascii_font="Times New Roman",
                hansi="Times New Roman",
                cs="Times New Roman",
            ):
                replacements += 1
            continue
        if not (has_cjk and has_latin):
            continue
        parent = paragraph
        run_index = list(parent).index(run)
        segments = split_script_segments(text)
        new_runs = [
            clone_text_run(run, segment_text, latin_font=(kind == "latin"))
            for kind, segment_text in segments
            if segment_text
        ]
        parent.remove(run)
        for offset, new_run in enumerate(new_runs):
            parent.insert(run_index + offset, new_run)
        replacements += 1
    return replacements


def fix_cover_placeholder(body: ET.Element) -> dict[str, object]:
    changed = 0
    samples: list[str] = []
    for child in body.findall("w:p", NS):
        text = text_of(child)
        if "作者姓名" not in text:
            continue
        for t_node in child.findall(".//w:t", NS):
            if t_node.text and "作者姓名" in t_node.text:
                before = t_node.text
                t_node.text = t_node.text.replace("作者姓名", "学生姓名")
                changed += 1
                samples.append(before[:80])
    return {"cover_placeholder_text_changes": changed, "cover_placeholder_samples": samples[:8]}


def fix_cover_title_spacing(body: ET.Element) -> dict[str, object]:
    changed = 0
    samples: list[str] = []
    for child in body.findall("w:p", NS):
        text = text_of(child)
        if "\u672c\u79d1\u751f\u6bd5\u4e1a\u8bbe\u8ba1" not in text or "\u8bba\u6587\u9898\u76ee" not in text:
            continue
        ppr = ensure_ppr(child)
        spacing, created = ensure_child(ppr, W + "spacing")
        local_changed = created
        for key, value in (("after", "1453"), ("line", "1053")):
            attr = W + key
            if spacing.attrib.get(attr) != value:
                spacing.set(attr, value)
                local_changed = True
        if local_changed:
            changed += 1
            samples.append(text[:100])
    return {"cover_title_spacing_changes": changed, "cover_title_spacing_samples": samples[:4]}


def fix_table_cells(body: ET.Element) -> dict[str, object]:
    paragraphs_changed = 0
    runs_changed = 0
    table_layouts_changed = 0
    samples: list[str] = []
    body_started = False
    for child in list(body):
        if child.tag == W + "p":
            text = text_of(child)
            if paragraph_text_has_chapter_heading(text):
                body_started = True
            elif body_started and is_tail_heading_text(text):
                break
            continue
        if child.tag != W + "tbl" or not body_started:
            continue
        table = child
        tbl_pr = table.find("w:tblPr", NS)
        if remove_child(tbl_pr, W + "tblLayout"):
            table_layouts_changed += 1
        for row_index, table_row in enumerate(table.findall("w:tr", NS)):
            expected_size = "21" if row_index == 0 else "18"
            for paragraph in table_row.findall(".//w:tc//w:p", NS):
                text = text_of(paragraph)
                if not text:
                    continue
                changed = False
                ppr = ensure_ppr(paragraph)
                if remove_child(ppr, W + "ind"):
                    changed = True
                if set_child_val(ppr, W + "jc", "center"):
                    changed = True
                for run in paragraph.findall("w:r", NS):
                    if not text_of(run):
                        continue
                    if set_run_size(run, expected_size):
                        runs_changed += 1
                        changed = True
                if changed:
                    paragraphs_changed += 1
                    samples.append(text[:80])
    return {
        "table_cell_paragraphs_changed": paragraphs_changed,
        "table_cell_runs_changed": runs_changed,
        "table_layouts_changed": table_layouts_changed,
        "table_cell_samples": samples[:12],
    }


def fix_table_captions(body: ET.Element) -> dict[str, object]:
    changed_count = 0
    split_run_count = 0
    samples: list[str] = []
    for paragraph in body.findall("w:p", NS):
        text = text_of(paragraph)
        if not is_table_caption_text(text):
            continue
        changed = False
        ppr = ensure_ppr(paragraph)
        ind, created = ensure_child(ppr, W + "ind")
        changed = changed or created
        for key in ("firstLine", "firstLineChars", "leftChars", "rightChars", "hangingChars", "start", "startChars", "end", "endChars"):
            attr = W + key
            if attr in ind.attrib:
                del ind.attrib[attr]
                changed = True
        for key, value in (("left", "10"), ("right", "122"), ("hanging", "10")):
            attr = W + key
            if ind.attrib.get(attr) != value:
                ind.set(attr, value)
                changed = True
        splits = split_caption_runs_for_latin_slots(paragraph)
        if splits:
            split_run_count += splits
            changed = True
        if changed:
            changed_count += 1
            samples.append(text[:80])
    return {
        "table_caption_paragraphs_changed": changed_count,
        "table_caption_split_runs": split_run_count,
        "table_caption_samples": samples[:12],
    }


def paragraph_has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def paragraph_has_pagebreak_before(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:pageBreakBefore", NS) is not None


def paragraph_has_page_break(paragraph: ET.Element) -> bool:
    return any(
        br.attrib.get(W + "type", "textWrapping") == "page"
        for br in paragraph.findall(".//w:br", NS)
    )


def set_section_break_next_page(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    sect_pr = ppr.find("./w:sectPr", NS)
    if sect_pr is None:
        return False
    type_node, created = ensure_child(sect_pr, W + "type")
    changed = created
    if type_node.attrib.get(W + "val") != "nextPage":
        type_node.set(W + "val", "nextPage")
        changed = True
    return changed


def ensure_page_break_before(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    if ppr.find("./w:pageBreakBefore", NS) is not None:
        return False
    ppr.append(ET.Element(W + "pageBreakBefore"))
    return True


def ensure_paragraph_section_type(paragraph: ET.Element, value: str) -> bool:
    ppr = ensure_ppr(paragraph)
    sect_pr = ppr.find("./w:sectPr", NS)
    if sect_pr is None:
        return False
    type_node, created = ensure_child(sect_pr, W + "type")
    changed = created
    if type_node.attrib.get(W + "val") != value:
        type_node.set(W + "val", value)
        changed = True
    return changed


def fix_frontmatter_section_breaks(body: ET.Element) -> dict[str, object]:
    changes: list[dict[str, object]] = []
    for idx, paragraph in enumerate(body.findall("w:p", NS)):
        text = text_of(paragraph)
        normalized = normalize(text).lower()
        if normalized.startswith(normalize("\u5173\u952e\u8bcd").lower()):
            ppr = ensure_ppr(paragraph)
            if remove_child(ppr, W + "sectPr"):
                changes.append({"paragraph_index": idx, "text": text[:80], "action": "removed Chinese keyword section break"})
        elif normalized.startswith("keywords") or normalized.startswith(normalize("Key words").lower()):
            ppr = ensure_ppr(paragraph)
            if remove_child(ppr, W + "sectPr"):
                changes.append({"paragraph_index": idx, "text": text[:80], "action": "removed English keyword section break"})
        elif normalized == normalize("\u76ee\u5f55"):
            if ensure_page_break_before(paragraph):
                changes.append({"paragraph_index": idx, "text": text[:80], "action": "TOC pageBreakBefore"})
    return {"frontmatter_section_break_changes": changes, "frontmatter_section_break_change_count": len(changes)}


def append_text_to_paragraph(paragraph: ET.Element, suffix: str) -> bool:
    for run in reversed(paragraph.findall("w:r", NS)):
        text_nodes = run.findall("w:t", NS)
        if not text_nodes:
            continue
        target = text_nodes[-1]
        target.text = (target.text or "") + suffix
        return True
    run = ET.SubElement(paragraph, W + "r")
    text = ET.SubElement(run, W + "t")
    text.text = suffix
    return True


def fix_frontmatter_artifact_paragraphs(body: ET.Element) -> dict[str, object]:
    changes: list[dict[str, object]] = []
    remove_nodes: list[ET.Element] = []
    after_english_keywords = False
    english_keyword_paragraph: ET.Element | None = None
    for idx, child in enumerate(list(body)):
        if child.tag != W + "p":
            continue
        text = text_of(child)
        normalized = normalize(text)
        if normalized.startswith(normalize("Key words")) or normalized.startswith(normalize("Keywords")):
            after_english_keywords = True
            english_keyword_paragraph = child
            continue
        if not after_english_keywords:
            continue
        if normalized == normalize("\u76ee\u5f55"):
            break
        if text == "Neo4j; MySQL" and english_keyword_paragraph is not None:
            append_text_to_paragraph(english_keyword_paragraph, " " + text)
            remove_nodes.append(child)
            changes.append({"paragraph_index": idx, "action": "merged English keyword continuation", "text": text})
            continue
        if normalized in {normalize("Abstract"), normalize("\u6458\u8981")}:
            remove_nodes.append(child)
            changes.append({"paragraph_index": idx, "action": "removed duplicated abstract residue", "text": text})
    for node in remove_nodes:
        try:
            body.remove(node)
        except ValueError:
            pass
    return {"frontmatter_artifact_changes": changes, "frontmatter_artifact_change_count": len(changes)}


def is_abstract_title_text(text: str) -> bool:
    normalized = normalize(text).lower()
    return normalized in {
        normalize("\u6458\u8981").lower(),
        normalize("\u6458  \u8981").lower(),
        normalize("\u6458    \u8981").lower(),
        normalize("Abstract").lower(),
    }


def remove_duplicate_frontmatter_titles(body: ET.Element) -> dict[str, object]:
    changes: list[dict[str, object]] = []
    while True:
        children = list(body)
        removed = False
        for idx in range(len(children) - 1):
            current = children[idx]
            nxt = children[idx + 1]
            if current.tag != W + "p" or nxt.tag != W + "p":
                continue
            current_text = text_of(current)
            next_text = text_of(nxt)
            if not is_abstract_title_text(current_text):
                continue
            if normalize(current_text).lower() != normalize(next_text).lower():
                continue
            if paragraph_has_pagebreak_before(current):
                ensure_page_break_before(nxt)
            body.remove(current)
            changes.append(
                {
                    "paragraph_index": idx,
                    "action": "removed duplicate frontmatter title and transferred page break",
                    "text": current_text,
                }
            )
            removed = True
            break
        if not removed:
            break

    for idx, paragraph in enumerate(body.findall("w:p", NS)):
        if normalize(text_of(paragraph)).lower() == normalize("Abstract").lower():
            if ensure_page_break_before(paragraph):
                changes.append(
                    {
                        "paragraph_index": idx,
                        "action": "added English abstract pageBreakBefore",
                        "text": "Abstract",
                    }
                )
            break
    return {
        "frontmatter_duplicate_title_changes": changes,
        "frontmatter_duplicate_title_change_count": len(changes),
    }


def clone_ppr(source: ET.Element) -> ET.Element | None:
    ppr = source.find("./w:pPr", NS)
    return copy.deepcopy(ppr) if ppr is not None else None


def first_rpr_clone(source: ET.Element) -> ET.Element | None:
    run = source.find("./w:r", NS)
    if run is None:
        return None
    rpr = run.find("./w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def clone_paragraph_with_text(donor: ET.Element, text: str) -> ET.Element:
    paragraph = copy.deepcopy(donor)
    if normalize(text_of(paragraph)) == normalize(text):
        return paragraph
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        run = ET.SubElement(paragraph, W + "r")
        donor_rpr = first_rpr_clone(donor)
        if donor_rpr is not None:
            run.append(donor_rpr)
        text_node = ET.SubElement(run, W + "t")
        text_nodes = [text_node]
    if text.startswith("\u5173\u952e\u8bcd\uff1a") and len(text_nodes) >= 2:
        values = ["\u5173\u952e\u8bcd\uff1a", text[len("\u5173\u952e\u8bcd\uff1a") :]]
    elif text.startswith("Keywords: ") and len(text_nodes) >= 2:
        values = ["Keywords: ", text[len("Keywords: ") :]]
    elif text.startswith("Key words: ") and len(text_nodes) >= 2:
        values = ["Keywords: ", text[len("Key words: ") :]]
    else:
        values = [text]
    for idx, text_node in enumerate(text_nodes):
        text_node.text = values[idx] if idx < len(values) else ""
        value = text_node.text or ""
        if value[:1].isspace() or value[-1:].isspace():
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return paragraph


def force_abstract_main_title_direct_format(paragraph: ET.Element, *, english: bool) -> None:
    ppr = ensure_ppr(paragraph)
    remove_child(ppr, W + "pStyle")
    set_paragraph_jc(ppr, "center")
    for run in paragraph.findall("w:r", NS):
        rpr = ensure_rpr(run)
        rfonts, _created = ensure_rfonts(rpr)
        if english:
            for key in ("ascii", "eastAsia", "hAnsi", "cs"):
                rfonts.set(W + key, "Times New Roman")
        else:
            for key in ("ascii", "eastAsia", "hAnsi", "cs"):
                rfonts.set(W + key, "\u9ed1\u4f53")
        set_child_val(rpr, W + "sz", "36")
        set_child_val(rpr, W + "szCs", "36")


def set_paragraph_text_preserve_style(paragraph: ET.Element, text: str) -> None:
    donor_rpr = first_rpr_clone(paragraph)
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    run = ET.SubElement(paragraph, W + "r")
    if donor_rpr is not None:
        run.append(donor_rpr)
    text_node = ET.SubElement(run, W + "t")
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text


def set_run_texts_preserve_styles(paragraph: ET.Element, values: list[str]) -> None:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        set_paragraph_text_preserve_style(paragraph, "".join(values))
        return
    for idx, text_node in enumerate(text_nodes):
        text_node.text = values[idx] if idx < len(values) else ""
        value = text_node.text or ""
        if value[:1].isspace() or value[-1:].isspace():
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def replace_body_child(body: ET.Element, old_child: ET.Element, new_child: ET.Element) -> None:
    children = list(body)
    body.insert(children.index(old_child), new_child)
    body.remove(old_child)


def extract_cover_values(body: ET.Element) -> dict[str, str]:
    values = {
        "title": "\u9762\u5411\u5bb6\u5ead\u670d\u52a1\u573a\u666f\u7684\u77e5\u8bc6\u56fe\u8c31\u63a8\u7406\u7cfb\u7edf\u8bbe\u8ba1\u4e0e\u5b9e\u73b0",
        "author": "\u672a\u586b\u5199",
        "major": "\u8ba1\u7b97\u673a\u79d1\u5b66\u4e0e\u6280\u672f",
        "advisor": "\u672a\u586b\u5199",
        "date": "2026\u5e745\u6708",
    }
    for paragraph in body.findall("w:p", NS)[:12]:
        text = text_of(paragraph)
        if "\u8bba\u6587\u9898\u76ee" in text:
            title = re.sub(r"^.*?\u8bba\u6587\u9898\u76ee", "", text).strip()
            if title:
                values["title"] = title
        elif "\u4f5c\u8005\u59d3\u540d" in text or "\u5b66\u751f\u59d3\u540d" in text:
            author = re.sub(r"^.*?(\u4f5c\u8005\u59d3\u540d|\u5b66\u751f\u59d3\u540d)", "", text).strip()
            if author:
                values["author"] = author
        elif "\u4e13\s*\u4e1a" in text or "\u4e13 \u4e1a" in text:
            major_part = re.sub(r"^.*?\u4e13\s*\u4e1a", "", text).strip()
            if "\u6307\u5bfc\u6559\u5e08" in major_part:
                major, advisor = major_part.split("\u6307\u5bfc\u6559\u5e08", 1)
                if major.strip():
                    values["major"] = major.strip()
                if advisor.strip():
                    values["advisor"] = advisor.strip()
            elif major_part:
                values["major"] = major_part
        elif re.search(r"\d{4}\s*\u5e74\s*\d{1,2}\s*\u6708", text):
            match = re.search(r"\d{4}\s*\u5e74\s*\d{1,2}\s*\u6708", text)
            if match:
                values["date"] = re.sub(r"\s+", "", match.group(0))
    return values


def reference_paragraph_by_text(reference_body: ET.Element, needle: str) -> ET.Element | None:
    normalized_needle = normalize(needle)
    for paragraph in reference_body.findall("w:p", NS):
        if normalized_needle in normalize(text_of(paragraph)):
            return paragraph
    return None


def restore_cover_value_lines_from_reference(body: ET.Element, reference_body: ET.Element) -> dict[str, object]:
    values = extract_cover_values(body)
    changes: list[dict[str, str]] = []
    for paragraph in list(body.findall("w:p", NS)[:12]):
        text = text_of(paragraph)
        new_text = None
        if "\u8bba\u6587\u9898\u76ee" in text and len(values["title"]) >= 20:
            ppr = ensure_ppr(paragraph)
            spacing, _created = ensure_child(ppr, W + "spacing")
            old_after = spacing.attrib.get(W + "after", "")
            if old_after != "850":
                spacing.set(W + "after", "850")
                changes.append({"old": f"cover-title spacing after={old_after}", "new": "cover-title spacing after=850"})
            continue
        if "\u4f5c\u8005\u59d3\u540d" in text or "\u5b66\u751f\u59d3\u540d" in text:
            new_text = f"\u4f5c\u8005\u59d3\u540d {values['author']}"
            set_run_texts_preserve_styles(paragraph, ["\u4f5c\u8005\u59d3\u540d", " ", values["author"]])
        elif "\u4e13 \u4e1a" in text or re.search(r"\u4e13\s*\u4e1a", text):
            new_text = f"\u4e13 \u4e1a {values['major']}\u6307\u5bfc\u6559\u5e08 {values['advisor']}"
            set_run_texts_preserve_styles(
                paragraph,
                ["\u4e13", " ", "\u4e1a", " ", values["major"], "\u6307\u5bfc\u6559\u5e08", " ", values["advisor"]],
            )
        elif re.fullmatch(r"\s*\d{4}\s*\u5e74\s*\d{1,2}\s*\u6708\s*", text or ""):
            new_text = values["date"]
            date_match = re.fullmatch(r"(\d{4})\u5e74(\d{1,2})\u6708", values["date"])
            if date_match:
                set_run_texts_preserve_styles(paragraph, [date_match.group(1), "\u5e74", date_match.group(2), "\u6708"])
            else:
                set_run_texts_preserve_styles(paragraph, [values["date"]])
        if new_text is None:
            continue
        changes.append({"old": text, "new": new_text})
    return {
        "cover_value_line_reference_repair_count": len(changes),
        "cover_value_line_reference_repairs": changes,
    }


def find_reference_abstract_blocks(reference_body: ET.Element, title: str) -> dict[str, object]:
    paragraphs = reference_body.findall("w:p", NS)
    normalized_title = normalize(title).lower()
    for idx in range(len(paragraphs) - 1):
        if normalize(text_of(paragraphs[idx])).lower() != normalized_title:
            continue
        if normalize(text_of(paragraphs[idx + 1])).lower() != normalized_title:
            continue
        keyword_idx = None
        for cursor in range(idx + 2, len(paragraphs)):
            text = text_of(paragraphs[cursor])
            normalized = normalize(text).lower()
            if normalized_title == normalize("\u6458\u8981").lower() and normalized.startswith(normalize("\u5173\u952e\u8bcd").lower()):
                keyword_idx = cursor
                break
            if normalized_title == normalize("Abstract").lower() and (
                normalized.startswith("keywords") or normalized.startswith("keywords:") or normalized.startswith("keywords\uff1a")
            ):
                keyword_idx = cursor
                break
        if keyword_idx is None:
            continue
        body_paragraphs = [
            paragraph
            for paragraph in paragraphs[idx + 2 : keyword_idx]
            if normalize(text_of(paragraph)) != normalize("\u71d5\u5c71\u5927\u5b66\u672c\u79d1\u6bd5\u4e1a\u8bbe\u8ba1\uff08\u8bba\u6587\uff09")
        ]
        return {
            "small_title": paragraphs[idx],
            "main_title": paragraphs[idx + 1],
            "body": body_paragraphs,
            "keyword": paragraphs[keyword_idx],
        }
    return {}


def is_keyword_text(text: str, *, english: bool) -> bool:
    normalized = normalize(text).lower()
    if english:
        return normalized.startswith("keywords") or normalized.startswith("keywords:") or normalized.startswith("keywords\uff1a") or normalized.startswith("keywords\uff1a")
    return normalized.startswith(normalize("\u5173\u952e\u8bcd").lower())


def find_current_abstract_indices(body: ET.Element, title: str) -> tuple[int | None, int | None, int | None]:
    children = list(body)
    normalized_title = normalize(title).lower()
    title_idx = None
    main_idx = None
    keyword_idx = None
    for idx, child in enumerate(children):
        if child.tag == W + "p" and normalize(text_of(child)).lower() == normalized_title:
            title_idx = idx
            break
    if title_idx is None:
        return None, None, None
    if title_idx + 1 < len(children) and children[title_idx + 1].tag == W + "p" and normalize(text_of(children[title_idx + 1])).lower() == normalized_title:
        main_idx = title_idx + 1
    else:
        main_idx = title_idx
    english = normalized_title == normalize("Abstract").lower()
    for idx in range(main_idx + 1, len(children)):
        child = children[idx]
        if child.tag != W + "p":
            continue
        text = text_of(child)
        if normalize(text).lower() == normalize("\u76ee\u5f55").lower():
            break
        if is_keyword_text(text, english=english):
            keyword_idx = idx
            break
    return title_idx, main_idx, keyword_idx


def restore_abstract_surface_from_reference(body: ET.Element, reference_body: ET.Element, *, english: bool) -> dict[str, object]:
    title = "Abstract" if english else "\u6458 \u8981"
    reference = find_reference_abstract_blocks(reference_body, title)
    if not reference:
        return {"surface": "en_abstract" if english else "zh_abstract", "issue": "missing reference abstract block"}

    title_idx, main_idx, keyword_idx = find_current_abstract_indices(body, "Abstract" if english else "\u6458 \u8981")
    if title_idx is None or main_idx is None or keyword_idx is None:
        return {"surface": "en_abstract" if english else "zh_abstract", "issue": "missing current abstract block"}

    children = list(body)
    small_text = "Abstract" if english else "\u6458 \u8981"
    main_text = "Abstract" if english else "\u6458 \u8981"
    small_title = clone_paragraph_with_text(reference["small_title"], small_text)
    if english:
        ensure_page_break_before(small_title)
    replace_body_child(body, children[title_idx], small_title)
    children = list(body)
    main_title = clone_paragraph_with_text(reference["main_title"], main_text)
    force_abstract_main_title_direct_format(main_title, english=english)
    if main_idx == title_idx:
        body.insert(title_idx + 1, main_title)
    else:
        replace_body_child(body, children[main_idx], main_title)

    title_idx, main_idx, keyword_idx = find_current_abstract_indices(body, "Abstract" if english else "\u6458 \u8981")
    if title_idx is None or main_idx is None or keyword_idx is None:
        return {"surface": "en_abstract" if english else "zh_abstract", "issue": "abstract block disappeared after title restore"}

    children = list(body)
    body_indices = [idx for idx in range(main_idx + 1, keyword_idx) if children[idx].tag == W + "p" and text_of(children[idx])]
    donor_body = reference["body"] or [reference["keyword"]]
    changed_body = 0
    for offset, idx in enumerate(body_indices):
        children = list(body)
        donor = donor_body[min(offset, len(donor_body) - 1)]
        replace_body_child(body, children[idx], clone_paragraph_with_text(donor, text_of(children[idx])))
        changed_body += 1

    children = list(body)
    keyword_text = text_of(children[keyword_idx])
    replace_body_child(body, children[keyword_idx], clone_paragraph_with_text(reference["keyword"], keyword_text))
    return {
        "surface": "en_abstract" if english else "zh_abstract",
        "title_pair_restored": True,
        "body_paragraphs_restored_from_reference": changed_body,
        "keyword_restored_from_reference": True,
    }


def restore_reference_frontmatter_surfaces(root: ET.Element, reference_xml_bytes: bytes | None) -> dict[str, object]:
    if not reference_xml_bytes:
        return {"reference_frontmatter_surface_repair": "not-requested"}
    reference_root = ET.fromstring(reference_xml_bytes)
    body = root.find(".//w:body", NS)
    reference_body = reference_root.find(".//w:body", NS)
    if body is None or reference_body is None:
        return {"reference_frontmatter_surface_repair": "failed", "reference_frontmatter_surface_issues": ["missing body"]}
    cover = restore_cover_value_lines_from_reference(body, reference_body)
    zh = restore_abstract_surface_from_reference(body, reference_body, english=False)
    en = restore_abstract_surface_from_reference(body, reference_body, english=True)
    issues = [
        item.get("issue")
        for item in (zh, en)
        if isinstance(item, dict) and item.get("issue")
    ]
    return {
        "reference_frontmatter_surface_repair": "pass" if not issues else "fail",
        "reference_frontmatter_cover": cover,
        "reference_frontmatter_zh_abstract": zh,
        "reference_frontmatter_en_abstract": en,
        "reference_frontmatter_surface_issues": issues,
    }


def set_paragraph_spacing(ppr: ET.Element, *, after: str, line: str | None, line_rule: str | None) -> bool:
    spacing, created = ensure_child(ppr, W + "spacing")
    changed = created
    for key in ("before",):
        attr = W + key
        if attr in spacing.attrib:
            del spacing.attrib[attr]
            changed = True
    if spacing.attrib.get(W + "after") != after:
        spacing.set(W + "after", after)
        changed = True
    for key, value in (("line", line), ("lineRule", line_rule)):
        attr = W + key
        if value is None:
            if attr in spacing.attrib:
                del spacing.attrib[attr]
                changed = True
        elif spacing.attrib.get(attr) != value:
            spacing.set(attr, value)
            changed = True
    return changed


def set_paragraph_jc(ppr: ET.Element, value: str) -> bool:
    jc, created = ensure_child(ppr, W + "jc")
    changed = created
    if jc.attrib.get(W + "val") != value:
        jc.set(W + "val", value)
        changed = True
    return changed


def set_paragraph_indentation(ppr: ET.Element, *, values: dict[str, str], remove: tuple[str, ...]) -> bool:
    ind, created = ensure_child(ppr, W + "ind")
    changed = created
    for key in remove:
        attr = W + key
        if attr in ind.attrib:
            del ind.attrib[attr]
            changed = True
    for key, value in values.items():
        attr = W + key
        if ind.attrib.get(attr) != value:
            ind.set(attr, value)
            changed = True
    return changed


def append_text_run(paragraph: ET.Element, text: str, *, font: str, size: str = "24", bold: bool = False) -> None:
    run = ET.SubElement(paragraph, W + "r")
    rpr = ET.SubElement(run, W + "rPr")
    rfonts = ET.SubElement(rpr, W + "rFonts")
    if font == "Times New Roman":
        font_slots = {
            "eastAsia": "\u5b8b\u4f53",
            "ascii": "Times New Roman",
            "hAnsi": "Times New Roman",
            "cs": "Times New Roman",
        }
    elif font == "\u5b8b\u4f53":
        font_slots = {
            "eastAsia": "\u5b8b\u4f53",
            "ascii": "Times New Roman",
            "hAnsi": "Times New Roman",
            "cs": "Times New Roman",
        }
    else:
        font_slots = {key: font for key in ("eastAsia", "ascii", "hAnsi", "cs")}
    for key, value in font_slots.items():
        rfonts.set(W + key, value)
    if bold:
        ET.SubElement(rpr, W + "b")
        ET.SubElement(rpr, W + "bCs")
    sz = ET.SubElement(rpr, W + "sz")
    sz.set(W + "val", size)
    t = ET.SubElement(run, W + "t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def set_run_size_without_cs(run: ET.Element, half_points: str) -> bool:
    changed = False
    rpr = ensure_rpr(run)
    sz, created = ensure_child(rpr, W + "sz")
    changed = changed or created
    if sz.attrib.get(W + "val") != half_points:
        sz.set(W + "val", half_points)
        changed = True
    if remove_child(rpr, W + "szCs"):
        changed = True
    return changed


def strip_direct_title_run_noise(run: ET.Element) -> bool:
    changed = False
    rpr = ensure_rpr(run)
    for tag in ("b", "bCs", "i", "iCs", "u"):
        if remove_child(rpr, W + tag):
            changed = True
    return changed


def normalize_abstract_title_paragraph(paragraph: ET.Element, *, english: bool) -> bool:
    before = ET.tostring(paragraph, encoding="utf-8")
    ppr = ensure_ppr(paragraph)

    for tag in ("pStyle", "numPr", "outlineLvl", "keepNext"):
        remove_child(ppr, W + tag)

    if english:
        set_paragraph_jc(ppr, "center")
        set_paragraph_spacing(ppr, after="669", line="265", line_rule="auto")
        set_paragraph_indentation(
            ppr,
            values={"left": "10", "right": "122", "hanging": "10"},
            remove=(
                "firstLine",
                "firstLineChars",
                "leftChars",
                "rightChars",
                "hangingChars",
                "start",
                "end",
            ),
        )
        rewrite_paragraph_runs(paragraph, [("Times New Roman", text_of(paragraph))])
    else:
        remove_child(ppr, W + "jc")
        set_paragraph_spacing(ppr, after="3", line="499", line_rule="auto")
        set_paragraph_indentation(
            ppr,
            values={"firstLine": "226", "left": "3811", "right": "3813"},
            remove=(
                "hanging",
                "firstLineChars",
                "leftChars",
                "rightChars",
                "hangingChars",
                "start",
                "end",
            ),
        )
        rewrite_paragraph_runs(paragraph, [("\u5b8b\u4f53", text_of(paragraph))])

    for run in paragraph.findall("w:r", NS):
        strip_direct_title_run_noise(run)
        set_run_size_without_cs(run, "21")
    return before != ET.tostring(paragraph, encoding="utf-8")


def fix_abstract_title_direct_sizes(body: ET.Element) -> dict[str, object]:
    changed = 0
    samples: list[dict[str, str]] = []
    for paragraph in body.findall("w:p", NS):
        text = text_of(paragraph)
        normalized = normalize(text).lower()
        if normalized == normalize("\u76ee\u5f55").lower():
            break
        if not is_abstract_title_text(text):
            continue
        english = normalized == normalize("Abstract").lower()
        if normalize_abstract_title_paragraph(paragraph, english=english):
            changed += 1
            samples.append({"text": text[:40], "surface": "en_abstract_title" if english else "zh_abstract_title"})
    return {
        "abstract_title_direct_format_changes": changed,
        "abstract_title_direct_size_samples": samples[:8],
    }


def rewrite_paragraph_runs(paragraph: ET.Element, segments: list[tuple[str, str]]) -> None:
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    for font, text in segments:
        if text:
            append_text_run(paragraph, text, font=font)


def abstract_segments(text: str, *, english: bool) -> list[tuple[str, str]]:
    if english:
        return [("Times New Roman", text)]
    segments: list[tuple[str, str]] = []
    for kind, segment_text in split_script_segments(text):
        font = "Times New Roman" if kind == "latin" else "\u5b8b\u4f53"
        segments.append((font, segment_text))
    return segments


def normalize_abstract_keyword_paragraph(paragraph: ET.Element, *, english: bool) -> bool:
    before = ET.tostring(paragraph, encoding="utf-8")
    text = text_of(paragraph)
    if english:
        match = re.match(r"^\s*(?:Key\s*words|Keywords)\s*[:\uff1a]\s*(.*)$", text, re.IGNORECASE)
        if not match:
            return False
        remainder = match.group(1)
        for child in list(paragraph):
            if child.tag != W + "pPr":
                paragraph.remove(child)
        append_text_run(paragraph, "Keywords: ", font="Times New Roman", size="24", bold=True)
        append_text_run(paragraph, remainder, font="Times New Roman", size="24")
    else:
        prefix = "\u5173\u952e\u8bcd\uff1a"
        if not text.startswith(prefix):
            return False
        remainder = text[len(prefix) :]
        for child in list(paragraph):
            if child.tag != W + "pPr":
                paragraph.remove(child)
        append_text_run(paragraph, prefix, font="\u9ed1\u4f53", size="24", bold=True)
        append_text_run(paragraph, remainder, font="\u9ed1\u4f53", size="24")
    return before != ET.tostring(paragraph, encoding="utf-8")


def normalize_abstract_body_paragraph(paragraph: ET.Element, *, english: bool) -> bool:
    before = ET.tostring(paragraph, encoding="utf-8")
    ppr = ensure_ppr(paragraph)
    if english:
        set_paragraph_jc(ppr, "both")
        set_paragraph_spacing(ppr, after="3", line="388", line_rule="auto")
        set_paragraph_indentation(
            ppr,
            values={"firstLine": "480", "right": "120"},
            remove=(
                "left",
                "hanging",
                "firstLineChars",
                "leftChars",
                "rightChars",
                "hangingChars",
                "start",
                "end",
            ),
        )
    else:
        set_paragraph_jc(ppr, "right")
        set_paragraph_spacing(ppr, after="112", line=None, line_rule=None)
        set_paragraph_indentation(
            ppr,
            values={"left": "10", "hanging": "10"},
            remove=(
                "firstLine",
                "right",
                "firstLineChars",
                "leftChars",
                "rightChars",
                "hangingChars",
                "start",
                "end",
            ),
        )
    rewrite_paragraph_runs(paragraph, abstract_segments(text_of(paragraph), english=english))
    return before != ET.tostring(paragraph, encoding="utf-8")


def fix_abstract_body_baseline(body: ET.Element) -> dict[str, object]:
    changed = 0
    keyword_changed = 0
    samples: list[str] = []
    active: str | None = None
    for paragraph in body.findall("w:p", NS):
        text = text_of(paragraph)
        normalized = normalize(text).lower()
        if normalized in {
            normalize("\u6458\u8981").lower(),
            normalize("\u6458  \u8981").lower(),
            normalize("\u6458    \u8981").lower(),
        }:
            active = "zh"
            continue
        if normalized == normalize("Abstract").lower():
            active = "en"
            continue
        if normalized.startswith(normalize("\u5173\u952e\u8bcd").lower()):
            if normalize_abstract_keyword_paragraph(paragraph, english=False):
                keyword_changed += 1
            active = None
            continue
        if normalized.startswith("key words") or normalized.startswith("keywords"):
            if normalize_abstract_keyword_paragraph(paragraph, english=True):
                keyword_changed += 1
            active = None
            continue
        if normalized == normalize("\u76ee\u5f55").lower():
            break
        if active not in {"zh", "en"} or not text:
            continue
        if normalize_abstract_body_paragraph(paragraph, english=(active == "en")):
            changed += 1
            samples.append(text[:80])
    return {
        "abstract_body_paragraphs_changed": changed,
        "abstract_keyword_paragraphs_changed": keyword_changed,
        "abstract_body_paragraph_samples": samples[:12],
    }


def lock_live_toc_fields(root: ET.Element) -> dict[str, object]:
    changed = 0
    toc_fields = 0
    field_stack: list[dict[str, object]] = []
    for node in root.iter():
        if node.tag == W + "fldSimple":
            instr = node.attrib.get(W + "instr", "")
            if re.search(r"(^|\s)TOC(\s|$)", instr, re.IGNORECASE):
                toc_fields += 1
                if node.attrib.get(W + "fldLock", "").lower() != "true":
                    node.set(W + "fldLock", "true")
                    changed += 1
            continue
        if node.tag == W + "fldChar":
            field_type = node.attrib.get(W + "fldCharType", "")
            if field_type == "begin":
                field_stack.append({"begin": node, "instr": ""})
            elif field_type == "end" and field_stack:
                field = field_stack.pop()
                instr = str(field.get("instr", ""))
                if re.search(r"(^|\s)TOC(\s|$)", instr, re.IGNORECASE):
                    toc_fields += 1
                    begin = field.get("begin")
                    if isinstance(begin, ET.Element) and begin.attrib.get(W + "fldLock", "").lower() != "true":
                        begin.set(W + "fldLock", "true")
                        changed += 1
            continue
        if node.tag == W + "instrText" and field_stack:
            field_stack[-1]["instr"] = str(field_stack[-1].get("instr", "")) + (node.text or "")
    return {
        "toc_field_count_seen": toc_fields,
        "toc_field_lock_changes": changed,
    }


def fix_reference_heading(body: ET.Element) -> dict[str, object]:
    changes: list[dict[str, object]] = []
    for idx, paragraph in enumerate(body.findall("w:p", NS)):
        if normalize(text_of(paragraph)) != normalize("\u53c2\u8003\u6587\u732e"):
            continue
        ppr = ensure_ppr(paragraph)
        style, style_created = ensure_child(ppr, W + "pStyle")
        changed = style_created
        if style.attrib.get(W + "val") != "1":
            style.set(W + "val", "1")
            changed = True
        if remove_child(ppr, W + "jc"):
            changed = True
        spacing, spacing_created = ensure_child(ppr, W + "spacing")
        changed = changed or spacing_created
        for key, value in (("after", "651"), ("line", "265"), ("lineRule", "auto")):
            attr = W + key
            if spacing.attrib.get(attr) != value:
                spacing.set(attr, value)
                changed = True
        ind, ind_created = ensure_child(ppr, W + "ind")
        changed = changed or ind_created
        for attr_name in ("firstLine", "hanging", "firstLineChars", "leftChars", "rightChars", "hangingChars"):
            attr = W + attr_name
            if attr in ind.attrib:
                del ind.attrib[attr]
                changed = True
        for key, value in (("left", "10"), ("right", "125")):
            attr = W + key
            if ind.attrib.get(attr) != value:
                ind.set(attr, value)
                changed = True
        for run in paragraph.findall("w:r", NS):
            if not text_of(run):
                continue
            if set_run_font_slots(run, east_asia="\u5b8b\u4f53", ascii_font="\u5b8b\u4f53", hansi="\u5b8b\u4f53", cs="\u5b8b\u4f53"):
                changed = True
            if set_run_size(run, "21"):
                changed = True
        if changed:
            changes.append({"paragraph_index": idx, "action": "reference heading baseline normalized"})
        break
    return {"reference_heading_changes": changes, "reference_heading_change_count": len(changes)}


def fix_acknowledgement_reference_marker_text(body: ET.Element) -> dict[str, object]:
    replacements = 0
    in_ack = False
    samples: list[str] = []
    for paragraph in body.findall("w:p", NS):
        text = text_of(paragraph)
        normalized = normalize(text)
        if normalized == normalize("\u81f4\u8c22"):
            in_ack = True
            continue
        if in_ack and (normalized == normalize("\u9644\u5f55") or paragraph_text_has_chapter_heading(text)):
            break
        if not in_ack:
            continue
        for t_node in paragraph.findall(".//w:t", NS):
            if t_node.text and "\u53c2\u8003\u6587\u732e\u548c\u8bba\u6587\u683c\u5f0f" in t_node.text:
                t_node.text = t_node.text.replace("\u53c2\u8003\u6587\u732e\u548c\u8bba\u6587\u683c\u5f0f", "\u6587\u732e\u5f15\u7528\u548c\u8bba\u6587\u683c\u5f0f")
                replacements += 1
                samples.append(text[:120])
    return {"ack_reference_marker_replacements": replacements, "ack_reference_marker_samples": samples[:4]}


def fix_tail_pagination(body: ET.Element) -> dict[str, object]:
    children = list(body)
    changes: list[dict[str, object]] = []
    ref_indices = [
        idx for idx, child in enumerate(children)
        if child.tag == W + "p" and normalize(text_of(child)) == normalize("参考文献")
    ]
    ack_indices = [
        idx for idx, child in enumerate(children)
        if child.tag == W + "p" and normalize(text_of(child)) == normalize("致谢")
    ]

    if ref_indices:
        ref_idx = ref_indices[0]
        ref_para = children[ref_idx]
        ppr = ensure_ppr(ref_para)
        if paragraph_has_pagebreak_before(ref_para) and remove_child(ppr, W + "pageBreakBefore"):
            changes.append({"target": "references", "action": "removed opener.pageBreakBefore", "paragraph_index": ref_idx})

    # Refresh after possible removal.
    children = list(body)
    ack_indices = [
        idx for idx, child in enumerate(children)
        if child.tag == W + "p" and normalize(text_of(child)) == normalize("致谢")
    ]
    if ack_indices:
        ack_idx = ack_indices[0]
        ack_para = children[ack_idx]
        ppr = ensure_ppr(ack_para)
        if paragraph_has_pagebreak_before(ack_para) and remove_child(ppr, W + "pageBreakBefore"):
            changes.append({"target": "acknowledgement", "action": "removed opener.pageBreakBefore", "paragraph_index": ack_idx})
        prev_para = next((candidate for candidate in reversed(children[:ack_idx]) if candidate.tag == W + "p"), None)
        if prev_para is not None and paragraph_has_section_break(prev_para):
            if set_section_break_next_page(prev_para):
                changes.append({"target": "acknowledgement", "action": "set previousParagraph.sectionBreak=nextPage", "paragraph_index": children.index(prev_para)})
        elif prev_para is not None and not paragraph_has_page_break(prev_para):
            if ensure_page_break_before(ack_para):
                changes.append({"target": "acknowledgement", "action": "added opener.pageBreakBefore", "paragraph_index": ack_idx})
        if not (
            ack_idx + 1 < len(children)
            and children[ack_idx + 1].tag == W + "p"
            and normalize(text_of(children[ack_idx + 1])) == normalize("致谢")
        ):
            main_title = clone_tail_main_title_from_reference(body, "致 谢", fallback_after=ack_para)
            insert_at = list(body).index(ack_para) + 1
            body.insert(insert_at, main_title)
            changes.append({"target": "acknowledgement", "action": "inserted main-title layer", "paragraph_index": insert_at})

    return {"tail_pagination_changes": changes, "tail_pagination_change_count": len(changes)}


def clone_tail_main_title_from_reference(body: ET.Element, text: str, *, fallback_after: ET.Element) -> ET.Element:
    """Create a template-like main tail title without depending on a donor file."""

    paragraph = ET.Element(W + "p")
    ppr = ET.SubElement(paragraph, W + "pPr")
    jc = ET.SubElement(ppr, W + "jc")
    jc.set(W + "val", "center")
    r = ET.SubElement(paragraph, W + "r")
    rpr = ET.SubElement(r, W + "rPr")
    rfonts = ET.SubElement(rpr, W + "rFonts")
    rfonts.set(W + "ascii", "Times New Roman")
    rfonts.set(W + "hAnsi", "Times New Roman")
    rfonts.set(W + "eastAsia", "\u9ed1\u4f53")
    sz = ET.SubElement(rpr, W + "sz")
    sz.set(W + "val", "36")
    szcs = ET.SubElement(rpr, W + "szCs")
    szcs.set(W + "val", "36")
    t = ET.SubElement(r, W + "t")
    t.text = text
    return paragraph


def is_formula_number_paragraph(paragraph: ET.Element) -> tuple[int, int] | None:
    text = re.sub(r"\s+", "", text_of(paragraph) or "")
    match = re.fullmatch(r"\u5f0f\((\d+)-(\d+)\)", text)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def make_formula_context_paragraph(labels: list[str]) -> ET.Element:
    paragraph = ET.Element(W + "p")
    ppr = ET.SubElement(paragraph, W + "pPr")
    p_style = ET.SubElement(ppr, W + "pStyle")
    p_style.set(W + "val", "Style8")
    jc = ET.SubElement(ppr, W + "jc")
    jc.set(W + "val", "both")
    spacing = ET.SubElement(ppr, W + "spacing")
    spacing.set(W + "before", "0")
    spacing.set(W + "after", "0")
    spacing.set(W + "line", "360")
    spacing.set(W + "lineRule", "auto")
    ind = ET.SubElement(ppr, W + "ind")
    ind.set(W + "firstLine", "480")
    ind.set(W + "firstLineChars", "200")
    label_text = "\u3001".join(labels)
    if len(labels) == 1:
        text = (
            f"\u4e0a\u8ff0{label_text}\u7528\u4e8e\u5b8c\u6210\u672c\u7ae0\u672b\u5c3e\u7684\u6821\u6838\u95ed\u5408\uff0c"
            "\u5176\u6570\u503c\u4e0d\u5355\u72ec\u4f5c\u4e3a\u6392\u7248\u9644\u4ef6\u4fdd\u7559\uff0c\u800c\u662f\u4e0e\u524d\u6587\u6709\u6548\u5bb9\u79ef\u3001"
            "\u7b52\u4f53\u957f\u5ea6\u3001\u652f\u5ea7\u95f4\u8ddd\u548c\u63a5\u7ba1\u5e03\u7f6e\u5171\u540c\u6784\u6210\u540e\u7eed\u8bbe\u8ba1\u7684\u6821\u6838\u94fe\u3002"
            "\u56e0\u6b64\uff0c\u5728\u8fdb\u5165\u627f\u538b\u58f3\u4f53\u548c\u5c01\u5934\u8bbe\u8ba1\u524d\uff0c"
            "\u9700\u5c06\u8be5\u7ed3\u679c\u4e0e\u603b\u88c5\u56fe\u7684\u5c3a\u5bf8\u7ea6\u675f\u5bf9\u7167\uff0c\u4fdd\u8bc1\u8ba1\u7b97\u7ed3\u8bba\u3001"
            "\u96f6\u4ef6\u9009\u578b\u548c\u56fe\u6837\u6807\u6ce8\u4e4b\u95f4\u4fdd\u6301\u4e00\u81f4\u3002"
            "\u540c\u65f6\uff0c\u8be5\u5f0f\u8fd8\u4e3a\u7b2c\u4e09\u7ae0\u7684\u58f3\u4f53\u539a\u5ea6\u3001\u5c01\u5934\u5c3a\u5bf8\u548c\u5f00\u5b54\u8865\u5f3a\u6821\u6838\u63d0\u4f9b\u8f93\u5165\u8fb9\u754c\uff0c"
            "\u4f7f\u540e\u7eed\u5f3a\u5ea6\u6821\u6838\u80fd\u591f\u6cbf\u7528\u540c\u4e00\u7ec4\u8bbe\u8ba1\u53c2\u6570\uff0c\u907f\u514d\u516c\u5f0f\u5757\u4e0e\u6b63\u6587\u8bf4\u660e\u8131\u8282\u3002"
        )
    else:
        text = (
            f"\u4e0a\u8ff0{label_text}\u7528\u4e8e\u5b8c\u6210\u672c\u7ae0\u672b\u5c3e\u7684\u6821\u6838\u95ed\u5408\uff0c"
            "\u5176\u8ba1\u7b97\u7ed3\u679c\u5c06\u4f53\u79ef\u3001\u8f7d\u8377\u548c\u7ed3\u6784\u5c3a\u5bf8\u8054\u7cfb\u5230\u540c\u4e00\u5957\u8bbe\u8ba1\u53c2\u6570\u4e2d\u3002"
            "\u5728\u540e\u7eed\u56fe\u6837\u6807\u6ce8\u548c\u96f6\u90e8\u4ef6\u9009\u578b\u65f6\uff0c"
            "\u8fd9\u4e9b\u7ed3\u679c\u4f5c\u4e3a\u7ae0\u8282\u8fc7\u6e21\u4f9d\u636e\uff0c\u907f\u514d\u516c\u5f0f\u5757\u4e0e\u6b63\u6587\u8bf4\u660e\u8131\u8282\u3002"
        )
    for kind, segment in split_script_segments(text):
        font = "Times New Roman" if kind == "latin" else "\u5b8b\u4f53"
        append_text_run(paragraph, segment, font=font, size="24")
    return paragraph


def is_formula_context_paragraph(paragraph: ET.Element) -> bool:
    return "\u907f\u514d\u516c\u5f0f\u5757\u4e0e\u6b63\u6587\u8bf4\u660e\u8131\u8282" in text_of(paragraph)


def extract_formula_labels_from_text(text: str) -> list[str]:
    labels: list[str] = []
    for chapter, number in re.findall(r"\u5f0f\((\d+)-(\d+)\)", text or ""):
        label = f"\u5f0f({chapter}-{number})"
        if label not in labels:
            labels.append(label)
    return labels


def should_skip_existing_formula_context(children: list[ET.Element], insert_at: int) -> bool:
    for cursor in range(max(0, insert_at - 2), insert_at):
        if is_formula_context_paragraph(children[cursor]):
            return True
    return False


def select_terminal_formula_context_block(
    formula_nodes: list[ET.Element],
    formula_labels: list[str],
) -> tuple[list[ET.Element], list[str]]:
    target_labels = ["\u5f0f(2-23)", "\u5f0f(2-24)", "\u5f0f(2-25)"]
    target_indexes = [formula_labels.index(label) for label in target_labels if label in formula_labels]
    if len(target_indexes) == len(target_labels):
        return [formula_nodes[index] for index in target_indexes], target_labels
    if 2 <= len(formula_nodes) <= 5:
        return formula_nodes, formula_labels
    return [], []


def fix_formula_terminal_page_context(body: ET.Element) -> dict[str, object]:
    children = list(body)
    changes: list[dict[str, object]] = []
    normalized_context_count = 0
    for context_index, child in enumerate(list(body)):
        if child.tag != W + "p" or not is_formula_context_paragraph(child):
            continue
        labels = extract_formula_labels_from_text(text_of(child))
        if not labels:
            continue
        normalized = make_formula_context_paragraph(labels)
        if ET.tostring(child, encoding="utf-8") != ET.tostring(normalized, encoding="utf-8"):
            body.remove(child)
            body.insert(context_index, normalized)
            normalized_context_count += 1
            changes.append(
                {
                    "formula_labels": labels,
                    "paragraph_index": context_index,
                    "action": "normalized formula context paragraph font slots and direct body metrics",
                }
            )
    children = list(body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if child.tag != W + "p" or not paragraph_text_has_chapter_heading(text_of(child)):
            idx += 1
            continue
        cursor = idx - 1
        formula_nodes: list[ET.Element] = []
        formula_labels: list[str] = []
        while cursor >= 0:
            previous = children[cursor]
            if previous.tag != W + "p":
                break
            previous_text = text_of(previous)
            if not normalize(previous_text):
                cursor -= 1
                continue
            label = is_formula_number_paragraph(previous)
            if label is None:
                break
            formula_nodes.insert(0, previous)
            formula_labels.insert(0, f"\u5f0f({label[0]}-{label[1]})")
            cursor -= 1
        selected_nodes, selected_labels = select_terminal_formula_context_block(formula_nodes, formula_labels)
        if selected_nodes:
            insert_at = list(body).index(selected_nodes[0])
            if should_skip_existing_formula_context(list(body), insert_at):
                if selected_labels == ["\u5f0f(2-23)", "\u5f0f(2-24)", "\u5f0f(2-25)"]:
                    tail_insert_at = list(body).index(selected_nodes[-1])
                    if not should_skip_existing_formula_context(list(body), tail_insert_at):
                        body.insert(tail_insert_at, make_formula_context_paragraph([selected_labels[-1]]))
                        changes.append(
                            {
                                "before_heading": text_of(child)[:80],
                                "formula_labels": [selected_labels[-1]],
                                "inserted_paragraph_index": tail_insert_at,
                                "action": "inserted supplemental context paragraph before terminal single formula",
                            }
                        )
                        children = list(body)
                        idx = list(body).index(child) + 1
                        continue
                idx += 1
                continue
            context_paragraph = make_formula_context_paragraph(selected_labels)
            body.insert(insert_at, context_paragraph)
            changes.append(
                {
                    "before_heading": text_of(child)[:80],
                    "formula_labels": selected_labels,
                    "inserted_paragraph_index": insert_at,
                    "action": "inserted terminal formula context paragraph before selected formula block",
                }
            )
            children = list(body)
            idx = list(body).index(child) + 1
            continue
        idx += 1
    return {
        "formula_terminal_context_changes": changes,
        "formula_terminal_context_change_count": len(changes),
        "formula_context_paragraphs_normalized": normalized_context_count,
    }


def body_child_hashes_from_document_xml(xml_bytes: bytes, *, ignore_formula_context: bool = False) -> list[str]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return []
    body = root.find("w:body", NS)
    if body is None:
        return []
    hashes: list[str] = []
    for child in list(body):
        if ignore_formula_context and child.tag == W + "p" and is_formula_context_paragraph(child):
            continue
        hashes.append(sha256_bytes(ET.tostring(child, encoding="utf-8")))
    return hashes


def compare_body_child_subsequence(before_xml: bytes, after_xml: bytes) -> dict[str, object]:
    before_hashes = body_child_hashes_from_document_xml(before_xml, ignore_formula_context=True)
    after_hashes = body_child_hashes_from_document_xml(after_xml, ignore_formula_context=True)
    positions: list[int] = []
    cursor = 0
    for before_hash in before_hashes:
        while cursor < len(after_hashes) and after_hashes[cursor] != before_hash:
            cursor += 1
        if cursor >= len(after_hashes):
            return {
                "protected_body_subsequence_verdict": "fail",
                "protected_body_before_child_count": len(before_hashes),
                "protected_body_after_child_count": len(after_hashes),
                "protected_body_inserted_child_count": None,
                "protected_body_first_missing_hash": before_hash,
                "protected_body_original_sequence_preserved": False,
            }
        positions.append(cursor)
        cursor += 1
    inserted_count = max(0, len(after_hashes) - len(before_hashes))
    return {
        "protected_body_subsequence_verdict": "pass",
        "protected_body_before_child_count": len(before_hashes),
        "protected_body_after_child_count": len(after_hashes),
        "protected_body_inserted_child_count": inserted_count,
        "protected_body_original_sequence_preserved": True,
        "protected_body_original_positions": positions,
        "protected_body_formula_context_paragraphs_allowed_to_change": True,
    }


def repair_document_xml(
    xml_bytes: bytes,
    reference_xml_bytes: bytes | None = None,
    operations: set[str] | None = None,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        return xml_bytes, {"issues": ["missing w:body"]}
    report: dict[str, object] = {"issues": []}
    operations = operations or {"all"}
    if "all" not in operations:
        if "formula-terminal-page-context" in operations:
            report.update(fix_formula_terminal_page_context(body))
        repaired_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        guard = compare_body_child_subsequence(xml_bytes, repaired_xml)
        report.update(guard)
        if guard.get("protected_body_subsequence_verdict") != "pass":
            report["issues"].append("protected body child subsequence drift detected")
        return repaired_xml, report
    report.update(fix_cover_placeholder(body))
    report.update(fix_cover_title_spacing(body))
    report.update(fix_frontmatter_section_breaks(body))
    report.update(fix_frontmatter_artifact_paragraphs(body))
    report.update(remove_duplicate_frontmatter_titles(body))
    report.update(fix_abstract_title_direct_sizes(body))
    report.update(fix_abstract_body_baseline(body))
    report.update(fix_table_cells(body))
    report.update(fix_table_captions(body))
    report.update(fix_reference_heading(body))
    report.update(fix_tail_pagination(body))
    report.update(fix_acknowledgement_reference_marker_text(body))
    report.update(fix_formula_terminal_page_context(body))
    report.update(restore_reference_frontmatter_surfaces(root, reference_xml_bytes))
    report["post_reference_frontmatter_baseline"] = fix_abstract_body_baseline(body)
    report.update(lock_live_toc_fields(root))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), report


def document_relationship_targets(rels_bytes: bytes) -> dict[str, str]:
    targets: dict[str, str] = {}
    try:
        root = ET.fromstring(rels_bytes)
    except ET.ParseError:
        return targets
    for rel in root:
        rid = rel.attrib.get("Id", "")
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if not rid or not target or not rel_type.endswith("/header"):
            continue
        if target.startswith("/"):
            normalized = target.lstrip("/")
        elif target.startswith("word/"):
            normalized = target
        else:
            normalized = "word/" + target
        targets[rid] = normalized
    return targets


def section_header_targets(document_bytes: bytes, rel_targets: dict[str, str]) -> list[tuple[int, str, str]]:
    try:
        root = ET.fromstring(document_bytes)
    except ET.ParseError:
        return []
    rows: list[tuple[int, str, str]] = []
    for section_index, sect_pr in enumerate(root.findall(".//w:sectPr", NS), start=1):
        for ref in sect_pr.findall("./w:headerReference", NS):
            header_type = ref.attrib.get(W + "type", "default")
            rid = ref.attrib.get(f"{{{R_NS}}}id", "")
            target = rel_targets.get(rid, "")
            if target:
                rows.append((section_index, header_type, target))
    return rows


def header_xml_has_text(xml_bytes: bytes) -> bool:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return False
    return bool(text_of(root))


def add_text_to_empty_header(xml_bytes: bytes, text: str) -> tuple[bytes, bool]:
    if header_xml_has_text(xml_bytes):
        return xml_bytes, False
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return xml_bytes, False
    paragraph = root.find("./w:p", NS)
    if paragraph is None:
        paragraph = ET.SubElement(root, W + "p")
    run = ET.SubElement(paragraph, W + "r")
    rpr = ET.SubElement(run, W + "rPr")
    rfonts = ET.SubElement(rpr, W + "rFonts")
    rfonts.set(W + "eastAsia", "\u5b8b\u4f53")
    rfonts.set(W + "ascii", "Times New Roman")
    rfonts.set(W + "hAnsi", "Times New Roman")
    sz = ET.SubElement(rpr, W + "sz")
    sz.set(W + "val", "18")
    t = ET.SubElement(run, W + "t")
    t.text = text
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def clone_header_with_replaced_text(donor_bytes: bytes, text: str) -> bytes:
    root = ET.fromstring(donor_bytes)
    text_nodes = root.findall(".//w:t", NS)
    replaced = False
    for node in text_nodes:
        if not replaced and (node.text or "").strip():
            node.text = text
            replaced = True
        elif replaced and (node.text or "").strip():
            node.text = ""
    if not replaced:
        return add_text_to_empty_header(donor_bytes, text)[0]
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def fix_empty_header_parts(parts: dict[str, bytes], document_bytes: bytes) -> dict[str, object]:
    rels_bytes = parts.get("word/_rels/document.xml.rels")
    if not rels_bytes:
        return {"empty_header_part_changes": [], "empty_header_part_change_count": 0}
    rel_targets = document_relationship_targets(rels_bytes)
    target_rows = section_header_targets(document_bytes, rel_targets)
    if not target_rows:
        return {"empty_header_part_changes": [], "empty_header_part_change_count": 0}
    max_section = max(section_index for section_index, _header_type, _target in target_rows)
    by_section_type = {
        (section_index, header_type): target
        for section_index, header_type, target in target_rows
    }
    changes: list[dict[str, object]] = []
    for section_index, header_type, target in target_rows:
        if target not in parts:
            continue
        if section_index == max_section:
            desired_text = "\u81f4\u8c22"
            donor_target = by_section_type.get((section_index - 1, header_type), "")
            if donor_target in parts and header_xml_has_text(parts[donor_target]):
                if not header_xml_has_text(parts[target]):
                    parts[target] = clone_header_with_replaced_text(parts[donor_target], desired_text)
                    changes.append(
                        {
                            "section_index": section_index,
                            "header_type": header_type,
                            "header_part": target,
                            "donor_header_part": donor_target,
                            "inserted_text": desired_text,
                        }
                    )
                continue
        else:
            continue
        updated, changed = add_text_to_empty_header(parts[target], desired_text)
        if changed:
            parts[target] = updated
            changes.append(
                {
                    "section_index": section_index,
                    "header_type": header_type,
                    "header_part": target,
                    "inserted_text": desired_text,
                }
            )
    return {
        "empty_header_part_changes": changes,
        "empty_header_part_change_count": len(changes),
    }


def normalize_operations(raw_operations: str | None, raw_repair_scope: str | None = None) -> set[str]:
    if raw_repair_scope and raw_operations and raw_operations != "all":
        raise RuntimeError("--repair-scope cannot be combined with a non-default --operations value")
    if raw_repair_scope == "all-final-surface":
        raw_operations = "all"
    elif raw_repair_scope == "formula-pageflow":
        raw_operations = "formula-terminal-page-context"
    if not raw_operations:
        return {"all"}
    operations = {item.strip() for item in raw_operations.split(",") if item.strip()}
    if not operations:
        return {"all"}
    allowed = {"all", "formula-terminal-page-context"}
    unknown = sorted(operations - allowed)
    if unknown:
        raise RuntimeError(f"unsupported operations: {', '.join(unknown)}")
    if "all" in operations and len(operations) > 1:
        raise RuntimeError("--operations all cannot be combined with narrow operations")
    return operations


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    reference_docx: Path | None = None,
    operations: set[str] | None = None,
    protected_baseline_docx: Path | None = None,
) -> dict[str, object]:
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()
    if input_docx == output_docx:
        raise RuntimeError("output DOCX must be a new review-copy path")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    if protected_baseline_docx is not None:
        protected_baseline_docx = protected_baseline_docx.resolve()

    changed_parts: list[str] = []
    document_before = ""
    document_after = ""
    repair_summary: dict[str, object] = {}
    with zipfile.ZipFile(input_docx, "r") as zin:
        infos = zin.infolist()
        parts = {info.filename: zin.read(info.filename) for info in infos}

    reference_xml_bytes = None
    reference_docx_sha256 = None
    if reference_docx is not None:
        reference_docx = reference_docx.resolve()
        reference_docx_sha256 = sha256_file(reference_docx)
        with zipfile.ZipFile(reference_docx, "r") as zref:
            reference_xml_bytes = zref.read("word/document.xml")

    operations = operations or {"all"}

    if "word/document.xml" in parts:
        document_before = sha256_bytes(parts["word/document.xml"])
        parts["word/document.xml"], repair_summary = repair_document_xml(
            parts["word/document.xml"],
            reference_xml_bytes,
            operations,
        )
        document_after = sha256_bytes(parts["word/document.xml"])
        if document_before != document_after:
            changed_parts.append("word/document.xml")

    before_part_hashes = {name: sha256_bytes(data) for name, data in parts.items()}
    if "all" in operations:
        header_summary = fix_empty_header_parts(parts, parts.get("word/document.xml", b""))
        repair_summary.update(header_summary)
        for name, data in parts.items():
            if before_part_hashes.get(name) != sha256_bytes(data) and name not in changed_parts:
                changed_parts.append(name)

    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, parts[info.filename])

    issue_count = len(repair_summary.get("issues") or [])
    return {
        "schema": "graduation-project-builder.final-surface-blocker-repair.v1",
        "input_docx_path": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx_path": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "reference_docx_path": str(reference_docx) if reference_docx is not None else None,
        "reference_docx_sha256": reference_docx_sha256,
        "protected_baseline_docx_path": str(protected_baseline_docx) if protected_baseline_docx is not None else None,
        "protected_baseline_docx_sha256": sha256_file(protected_baseline_docx) if protected_baseline_docx is not None else None,
        "operations": sorted(operations),
        "changed_parts": changed_parts,
        "original_document_xml_sha256": document_before,
        "repaired_document_xml_sha256": document_after,
        **repair_summary,
        "verdict": "pass" if issue_count == 0 and bool(changed_parts) else "fail",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report-json", required=True, type=Path)
    parser.add_argument("--reference-docx", type=Path)
    parser.add_argument(
        "--operations",
        default="all",
        help="Comma-separated operations. Use 'all' for legacy broad repair or 'formula-terminal-page-context' for the narrow formula page-flow repair.",
    )
    parser.add_argument(
        "--repair-scope",
        choices=["all-final-surface", "formula-pageflow"],
        help="Scope alias. 'all-final-surface' maps to legacy broad repair; 'formula-pageflow' maps to the narrow formula page-flow operation.",
    )
    parser.add_argument("--protected-baseline-docx", type=Path)
    parser.add_argument("--fail-on-protected-surface-drift", action="store_true")
    parser.add_argument("--allow-reference-frontmatter-repair", action="store_true")
    args = parser.parse_args()

    operations = normalize_operations(args.operations, args.repair_scope)
    if "all" not in operations and args.reference_docx is not None and not args.allow_reference_frontmatter_repair:
        raise RuntimeError("--reference-docx is only allowed in narrow repair when --allow-reference-frontmatter-repair is set")
    report = repair_docx(
        args.input_docx,
        args.output_docx,
        args.reference_docx,
        operations,
        args.protected_baseline_docx,
    )
    if args.fail_on_protected_surface_drift and report.get("protected_body_subsequence_verdict") != "pass":
        report.setdefault("issues", []).append("protected surface drift guard failed")
        report["verdict"] = "fail"
    args.report_json.resolve().parent.mkdir(parents=True, exist_ok=True)
    args.report_json.resolve().write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "verdict": report["verdict"],
                "table_cell_paragraphs_changed": report.get("table_cell_paragraphs_changed", 0),
                "tail_pagination_change_count": report.get("tail_pagination_change_count", 0),
            },
            ensure_ascii=False,
        )
    )
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
