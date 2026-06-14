#!/usr/bin/env python3
"""Normalize thesis body citations without rebuilding bibliography formatting."""

from __future__ import annotations

import argparse
import re
import zipfile
from collections import defaultdict
from copy import deepcopy
from pathlib import Path
from typing import Iterator
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]
ET.register_namespace("w", NS["w"])

BIBLIO_HEADING = "\u53c2\u8003\u6587\u732e"
RAW_CITATION_RE = re.compile(r"\[(\d+)\]")
GROUPED_CITATION_RE = re.compile(r"\[((?:\d+\s*,\s*)+\d+)\]")
BOOKMARK_PREFIX = "cite_ref_"
SUMMARY_PREFIXES = ("\u6574\u4f53\u6765\u770b", "\u7efc\u5408\u6765\u770b", "\u603b\u4f53\u6765\u770b", "\u7efc\u4e0a", "\u603b\u7684\u6765\u770b")
INTRO_PREFIXES = ("\u56fd\u5185\u7814\u7a76", "\u56fd\u5916\u7814\u7a76", "\u73b0\u6709\u7814\u7a76", "\u76f8\u5173\u7814\u7a76")
TAIL_BLOCK_HEADINGS = (
    "\u9644\u5f55",
    "\u81f4\u8c22",
    "\u8c22\u8f9e",
    "\u540e\u8bb0",
    "\u58f0\u660e",
    "\u56fe\u7eb8\u5f52\u6863\u8bf4\u660e",
)
SENTENCE_RE = re.compile(r"[^\u3002\uff01\uff1f\uff1b.!?;]+[\u3002\uff01\uff1f\uff1b.!?;]?")
UNSAFE_PARAGRAPH_CHILD_TAGS = {
    W + "drawing",
    W + "pict",
    W + "object",
    W + "fldSimple",
    W + "hyperlink",
    W + "bookmarkStart",
    W + "bookmarkEnd",
    W + "commentRangeStart",
    W + "commentRangeEnd",
}


def normalize(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "")


def is_bibliography_heading_text(text: str) -> bool:
    normalized = normalize(text)
    heading = normalize(BIBLIO_HEADING)
    return normalized == heading or normalized.endswith(heading)


def heading_level(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    sep = r"[\s\u25a1]+"
    if re.match(rf"^\d{{1,2}}{sep}\S", stripped):
        return 1
    if re.match(r"^\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u7ae0", stripped):
        return 1
    if re.match(rf"^\d{{1,2}}\.\d+{sep}\S", stripped):
        return 2
    if re.match(rf"^\d{{1,2}}\.\d+\.\d+{sep}\S", stripped):
        return 3
    if re.match(rf"^\d{{1,2}}\.\d+\.\d+\.\d+{sep}\S", stripped):
        return 4
    return None


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join((node.text or "") for node in paragraph.findall(".//w:t", NS))


def has_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text or "")


def has_ascii_alnum(text: str) -> bool:
    return any(ch.isascii() and ch.isalnum() for ch in text or "")


def split_by_script_runs(text: str) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    current_kind: str | None = None
    current: list[str] = []
    for char in text:
        if has_cjk(char):
            kind = "cjk"
        elif char.isascii() and (char.isalnum() or char in "[](){}.,;:/\\-+_=&%#@'\"<> "):
            kind = "latin"
        else:
            kind = current_kind or "latin"
        if current_kind is None:
            current_kind = kind
        if kind != current_kind:
            segments.append((current_kind, "".join(current)))
            current = [char]
            current_kind = kind
        else:
            current.append(char)
    if current:
        segments.append((current_kind or "latin", "".join(current)))
    return segments


def paragraph_has_unsafe_payload(paragraph: ET.Element) -> bool:
    if paragraph.find(".//w:instrText", NS) is not None:
        return True
    return any(node.tag in UNSAFE_PARAGRAPH_CHILD_TAGS for node in paragraph.iter())


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.attrib.get(W + "val", "") if style is not None else ""


def paragraph_has_outline_level(paragraph: ET.Element) -> bool:
    return paragraph.find("w:pPr/w:outlineLvl", NS) is not None


def paragraph_is_forbidden_citation_surface(paragraph: ET.Element, text: str) -> bool:
    style_key = paragraph_style_id(paragraph).lower()
    if paragraph_has_outline_level(paragraph):
        return True
    if any(token in style_key for token in ("heading", "caption", "toc")):
        return True
    if heading_level(text) is not None:
        return True
    return bool(re.match(r"^\s*[\u56fe\u8868]\s*\d+(?:[.\-]\d+)*", text or ""))


def nonempty_direct_text_run_count(paragraph: ET.Element) -> int:
    count = 0
    for run in paragraph.findall("w:r", NS):
        text = "".join(node.text or "" for node in run.findall(".//w:t", NS))
        if text.strip():
            count += 1
    return count


def paragraph_requires_lossy_text_run_rebuild(paragraph: ET.Element) -> bool:
    """A multi-run host paragraph may carry protected direct formatting."""
    return nonempty_direct_text_run_count(paragraph) > 1


def split_sentences(text: str) -> list[str]:
    """Split visible prose without treating decimal points as sentence breaks."""
    stripped = text.strip()
    if not stripped:
        return []
    parts: list[str] = []
    start = 0
    for idx, char in enumerate(stripped):
        if char not in "\u3002\uff01\uff1f\uff1b!?;.":
            continue
        if char == ".":
            prev_char = stripped[idx - 1] if idx > 0 else ""
            next_char = stripped[idx + 1] if idx + 1 < len(stripped) else ""
            if prev_char.isdigit() and next_char.isdigit():
                continue
        parts.append(stripped[start : idx + 1].strip())
        start = idx + 1
    if start < len(stripped):
        tail = stripped[start:].strip()
        if tail:
            parts.append(tail)
    return parts if parts else [stripped]


def grouped_citation_numbers(raw: str) -> list[int]:
    return [int(part.strip()) for part in raw.split(",") if part.strip()]


def expand_grouped_citation_markers(text: str) -> str:
    """Rewrite grouped numeric markers into one citation-bearing sentence each."""

    def replacement(match: re.Match[str]) -> str:
        numbers = grouped_citation_numbers(match.group(1))
        if len(numbers) <= 1:
            return match.group(0)
        pieces = [f"[{numbers[0]}]"]
        for number in numbers[1:]:
            pieces.append(f"。同一设计依据见文献[{number}]")
        return "".join(pieces)

    return GROUPED_CITATION_RE.sub(replacement, text)


def is_summary_sentence(text: str) -> bool:
    stripped = text.strip()
    return any(stripped.startswith(prefix) for prefix in SUMMARY_PREFIXES)


def is_intro_sentence(text: str) -> bool:
    stripped = text.strip()
    return any(stripped.startswith(prefix) for prefix in INTRO_PREFIXES)


def attach_markers(sentences: list[str], numbers: list[int]) -> str:
    if not sentences:
        return "".join(f"[{n}]" for n in numbers)

    mapping: dict[int, list[int]] = defaultdict(list)
    if len(sentences) == len(numbers):
        start_idx = 0
        available = len(sentences)
    else:
        start_idx = 0
        end_idx = len(sentences)
        if sentences and is_intro_sentence(sentences[0]):
            start_idx = 1
        if end_idx - start_idx > len(numbers) and sentences and is_summary_sentence(sentences[-1]):
            end_idx -= 1
        available = max(0, end_idx - start_idx)
        if available < len(numbers):
            target_count = min(len(sentences), len(numbers))
            start_idx = max(0, len(sentences) - target_count)
            available = target_count

    target_count = min(available, len(numbers))
    for offset, number in enumerate(numbers[:target_count]):
        mapping[start_idx + offset].append(number)
    for number in numbers[target_count:]:
        mapping[len(sentences) - 1].append(number)

    out: list[str] = []
    for idx, sentence in enumerate(sentences):
        nums = mapping.get(idx, [])
        if not nums:
            out.append(sentence)
            continue
        markers = "".join(f"[{n}]" for n in nums)
        match = re.match(r"^(.*?)([\u3002\uff01\uff1f\uff1b.!?;])?$", sentence)
        if match:
            core = match.group(1) or ""
            punct = match.group(2) or ""
            out.append(f"{core}{markers}{punct}")
        else:
            out.append(sentence + markers)
    return "".join(out)


def normalize_citation_punctuation_in_text(text: str) -> str:
    """Move visible numeric citation markers before sentence punctuation."""
    fixed = str(text or "")
    citation = r"(\[(?:\d{1,3})\])"
    punctuation = r"([\u3002\uff01\uff1f\uff1b\uff0c.!?;,])"
    previous = None
    while fixed != previous:
        previous = fixed
        fixed = re.sub(rf"{punctuation}\s*{citation}", r"\2\1", fixed)
    return fixed


def first_text_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("w:r", NS):
        text = "".join(node.text or "" for node in run.findall(".//w:t", NS)).strip()
        if not text:
            continue
        rpr = run.find("w:rPr", NS)
        return deepcopy(rpr) if rpr is not None else None
    return None


def normalize_latin_run_rpr(rpr: ET.Element) -> None:
    rfonts = rpr.find(W + "rFonts")
    if rfonts is None:
        rfonts = ET.Element(W + "rFonts")
        rpr.insert(0, rfonts)
    for attr_name in (W + "eastAsia", W + "eastAsiaTheme"):
        if attr_name in rfonts.attrib:
            del rfonts.attrib[attr_name]
    for attr_name in (W + "ascii", W + "hAnsi", W + "cs"):
        rfonts.set(attr_name, "Times New Roman")


def make_run(
    text: str,
    *,
    superscript: bool,
    rpr_template: ET.Element | None = None,
    latin_text: bool = False,
) -> ET.Element:
    run = ET.Element(W + "r")
    rpr = deepcopy(rpr_template) if rpr_template is not None else None
    if superscript:
        rpr = ET.Element(W + "rPr")
        ET.SubElement(rpr, W + "color", {W + "val": "000000"})
        ET.SubElement(rpr, W + "u", {W + "val": "none"})
        ET.SubElement(rpr, W + "vertAlign", {W + "val": "superscript"})
    elif latin_text:
        if rpr is None:
            rpr = ET.Element(W + "rPr")
        normalize_latin_run_rpr(rpr)
    if rpr is not None:
        run.append(rpr)
    t = ET.SubElement(run, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_text_runs(text: str, rpr_template: ET.Element | None = None) -> list[ET.Element]:
    runs: list[ET.Element] = []
    for kind, value in split_by_script_runs(text):
        if not value:
            continue
        runs.append(
            make_run(
                value,
                superscript=False,
                rpr_template=rpr_template,
                latin_text=(kind == "latin" and has_ascii_alnum(value)),
            )
        )
    return runs


def make_hyperlink_run(text: str, anchor: str, rpr_template: ET.Element | None = None) -> ET.Element:
    hyperlink = ET.Element(W + "hyperlink")
    hyperlink.set(W + "anchor", anchor)
    hyperlink.set(W + "history", "1")
    hyperlink.append(make_run(text, superscript=True, rpr_template=rpr_template))
    return hyperlink


def make_bookmark_start(bookmark_id: int, name: str) -> ET.Element:
    node = ET.Element(W + "bookmarkStart")
    node.set(W + "id", str(bookmark_id))
    node.set(W + "name", name)
    return node


def make_bookmark_end(bookmark_id: int) -> ET.Element:
    node = ET.Element(W + "bookmarkEnd")
    node.set(W + "id", str(bookmark_id))
    return node


def next_bookmark_id(root: ET.Element) -> int:
    ids: list[int] = []
    for node in root.findall(".//w:bookmarkStart", NS):
        raw = node.attrib.get(W + "id", "")
        if raw.isdigit():
            ids.append(int(raw))
    return max(ids, default=0) + 1


def rebuild_paragraph(paragraph: ET.Element, text: str, anchor_map: dict[int, str]) -> None:
    ppr = paragraph.find(W + "pPr")
    rpr_template = first_text_run_rpr(paragraph)
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    cursor = 0
    for match in RAW_CITATION_RE.finditer(text):
        if match.start() > cursor:
            for run in make_text_runs(text[cursor:match.start()], rpr_template=rpr_template):
                paragraph.append(run)
        number = int(match.group(1))
        anchor = anchor_map[number]
        paragraph.append(make_hyperlink_run(match.group(0), anchor, rpr_template=rpr_template))
        cursor = match.end()
    if cursor < len(text):
        for run in make_text_runs(text[cursor:], rpr_template=rpr_template):
            paragraph.append(run)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        paragraph.insert(0, ppr)


def direct_run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", NS))


def run_rpr(run: ET.Element) -> ET.Element | None:
    return run.find("w:rPr", NS)


def replace_child(parent: ET.Element, old_child: ET.Element, new_children: list[ET.Element]) -> None:
    children = list(parent)
    index = children.index(old_child)
    parent.remove(old_child)
    for offset, child in enumerate(new_children):
        parent.insert(index + offset, child)


def citation_hyperlink_ok(hyperlink: ET.Element, anchor_map: dict[int, str]) -> bool:
    text = "".join(node.text or "" for node in hyperlink.findall(".//w:t", NS))
    match = RAW_CITATION_RE.fullmatch(text)
    if not match:
        return True
    return hyperlink.attrib.get(W + "anchor") == anchor_map.get(int(match.group(1)))


def linkify_split_citation_run_sequences(
    paragraph: ET.Element,
    anchor_map: dict[int, str],
    renumber_map: dict[int, int] | None = None,
) -> None:
    """Recover citation markers split across adjacent runs, e.g. "[", "4", "]"."""
    children = list(paragraph)
    index = 0
    while index < len(children):
        child = children[index]
        if child.tag != W + "r":
            index += 1
            continue
        first_text = direct_run_text(child)
        if not first_text.startswith("["):
            index += 1
            continue

        collected: list[ET.Element] = []
        combined = ""
        scan = index
        while scan < len(children) and children[scan].tag == W + "r" and len(combined) <= 16:
            run_text = direct_run_text(children[scan])
            if not run_text:
                break
            collected.append(children[scan])
            combined += run_text
            if RAW_CITATION_RE.fullmatch(combined):
                break
            if re.match(r"^\[\d*$", combined) is None:
                break
            scan += 1

        match = RAW_CITATION_RE.fullmatch(combined)
        if match is None or len(collected) <= 1:
            index += 1
            continue
        old_number = int(match.group(1))
        new_number = renumber_map[old_number] if renumber_map is not None else old_number
        if new_number not in anchor_map:
            index += 1
            continue
        replacement = make_hyperlink_run(f"[{new_number}]", anchor_map[new_number], rpr_template=run_rpr(collected[0]))
        for old_child in collected:
            paragraph.remove(old_child)
        paragraph.insert(index, replacement)
        children = list(paragraph)
        index += 1


def split_run_preserving_citations(run: ET.Element, anchor_map: dict[int, str]) -> list[ET.Element]:
    text = direct_run_text(run)
    if not RAW_CITATION_RE.search(text):
        return [run]
    rpr_template = run_rpr(run)
    pieces: list[ET.Element] = []
    cursor = 0
    for match in RAW_CITATION_RE.finditer(text):
        if match.start() > cursor:
            pieces.extend(make_text_runs(text[cursor:match.start()], rpr_template=rpr_template))
        number = int(match.group(1))
        pieces.append(make_hyperlink_run(match.group(0), anchor_map[number], rpr_template=rpr_template))
        cursor = match.end()
    if cursor < len(text):
        pieces.extend(make_text_runs(text[cursor:], rpr_template=rpr_template))
    return pieces


def split_run_preserving_renumbered_citations(
    run: ET.Element,
    renumber_map: dict[int, int],
    anchor_map: dict[int, str],
) -> list[ET.Element]:
    text = direct_run_text(run)
    if not RAW_CITATION_RE.search(text):
        return [run]
    rpr_template = run_rpr(run)
    pieces: list[ET.Element] = []
    cursor = 0
    for match in RAW_CITATION_RE.finditer(text):
        if match.start() > cursor:
            pieces.extend(make_text_runs(text[cursor:match.start()], rpr_template=rpr_template))
        old_number = int(match.group(1))
        new_number = renumber_map[old_number]
        pieces.append(make_hyperlink_run(f"[{new_number}]", anchor_map[new_number], rpr_template=rpr_template))
        cursor = match.end()
    if cursor < len(text):
        pieces.extend(make_text_runs(text[cursor:], rpr_template=rpr_template))
    return pieces


def update_hyperlink_citation_text(
    hyperlink: ET.Element,
    renumber_map: dict[int, int],
    anchor_map: dict[int, str],
) -> bool:
    text = "".join(node.text or "" for node in hyperlink.findall(".//w:t", NS))
    match = RAW_CITATION_RE.fullmatch(text)
    if not match:
        return False
    old_number = int(match.group(1))
    new_number = renumber_map[old_number]
    hyperlink.set(W + "anchor", anchor_map[new_number])
    text_nodes = hyperlink.findall(".//w:t", NS)
    if not text_nodes:
        return True
    set_text_node_text(text_nodes[0], f"[{new_number}]")
    for node in text_nodes[1:]:
        set_text_node_text(node, "")
    return True


def renumber_and_linkify_existing_citation_runs(
    paragraph: ET.Element,
    renumber_map: dict[int, int],
    anchor_map: dict[int, str],
) -> None:
    linkify_split_citation_run_sequences(paragraph, anchor_map, renumber_map=renumber_map)
    for child in list(paragraph):
        if child.tag == W + "hyperlink":
            update_hyperlink_citation_text(child, renumber_map, anchor_map)
            continue
        if child.tag != W + "r":
            continue
        new_children = split_run_preserving_renumbered_citations(child, renumber_map, anchor_map)
        if len(new_children) != 1 or new_children[0] is not child:
            replace_child(paragraph, child, new_children)


def linkify_existing_citation_runs(paragraph: ET.Element, anchor_map: dict[int, str]) -> None:
    linkify_split_citation_run_sequences(paragraph, anchor_map)
    for child in list(paragraph):
        if child.tag == W + "hyperlink":
            if not citation_hyperlink_ok(child, anchor_map):
                text = "".join(node.text or "" for node in child.findall(".//w:t", NS))
                match = RAW_CITATION_RE.fullmatch(text)
                if match:
                    child.set(W + "anchor", anchor_map[int(match.group(1))])
            continue
        if child.tag != W + "r":
            continue
        new_children = split_run_preserving_citations(child, anchor_map)
        if len(new_children) != 1 or new_children[0] is not child:
            replace_child(paragraph, child, new_children)


def consume_next_citation_number(number_iter: Iterator[int]) -> int:
    try:
        return next(number_iter)
    except StopIteration as exc:
        raise ValueError("more visible citation markers were found than planned occurrence assignments") from exc


def hyperlink_citation_fullmatch(hyperlink: ET.Element) -> bool:
    text = "".join(node.text or "" for node in hyperlink.findall(".//w:t", NS))
    return RAW_CITATION_RE.fullmatch(text) is not None


def replace_citation_occurrences_preserving_runs(
    paragraph: ET.Element,
    new_numbers: list[int],
    anchor_map: dict[int, str],
) -> None:
    """Replace citation markers by visible occurrence order, not old number value."""
    number_iter = iter(new_numbers)
    children = list(paragraph)
    index = 0
    while index < len(children):
        child = children[index]

        if child.tag == W + "hyperlink":
            if hyperlink_citation_fullmatch(child):
                new_number = consume_next_citation_number(number_iter)
                replacement = make_hyperlink_run(f"[{new_number}]", anchor_map[new_number])
                replace_child(paragraph, child, [replacement])
                children = list(paragraph)
            index += 1
            continue

        if child.tag != W + "r":
            index += 1
            continue

        first_text = direct_run_text(child)
        if first_text.startswith("["):
            collected: list[ET.Element] = []
            combined = ""
            scan = index
            while scan < len(children) and children[scan].tag == W + "r" and len(combined) <= 16:
                run_text = direct_run_text(children[scan])
                if not run_text:
                    break
                collected.append(children[scan])
                combined += run_text
                if RAW_CITATION_RE.fullmatch(combined):
                    break
                if re.match(r"^\[\d*$", combined) is None:
                    break
                scan += 1
            if RAW_CITATION_RE.fullmatch(combined):
                new_number = consume_next_citation_number(number_iter)
                replacement = make_hyperlink_run(
                    f"[{new_number}]",
                    anchor_map[new_number],
                    rpr_template=run_rpr(collected[0]),
                )
                for old_child in collected:
                    paragraph.remove(old_child)
                paragraph.insert(index, replacement)
                children = list(paragraph)
                index += 1
                continue

        text = direct_run_text(child)
        if not RAW_CITATION_RE.search(text):
            index += 1
            continue

        rpr_template = run_rpr(child)
        pieces: list[ET.Element] = []
        cursor = 0
        for match in RAW_CITATION_RE.finditer(text):
            if match.start() > cursor:
                pieces.extend(make_text_runs(text[cursor:match.start()], rpr_template=rpr_template))
            new_number = consume_next_citation_number(number_iter)
            pieces.append(make_hyperlink_run(f"[{new_number}]", anchor_map[new_number], rpr_template=rpr_template))
            cursor = match.end()
        if cursor < len(text):
            pieces.extend(make_text_runs(text[cursor:], rpr_template=rpr_template))
        replace_child(paragraph, child, pieces)
        children = list(paragraph)
        index += len(pieces)

    try:
        extra_number = next(number_iter)
    except StopIteration:
        return
    raise ValueError(f"planned citation assignment [{extra_number}] was not consumed by a visible marker")


def paragraph_has_occurrence_mode_unsupported_payload(paragraph: ET.Element) -> bool:
    if paragraph.find(".//w:instrText", NS) is not None:
        return True
    for node in paragraph.iter():
        if node.tag == W + "hyperlink":
            if hyperlink_citation_fullmatch(node):
                continue
            return True
        if node.tag in {
            W + "drawing",
            W + "pict",
            W + "object",
            W + "fldSimple",
            W + "bookmarkStart",
            W + "bookmarkEnd",
            W + "commentRangeStart",
            W + "commentRangeEnd",
        }:
            return True
    return False


def bibliography_entry_numbers(paragraphs: list[ET.Element], biblio_idx: int) -> list[int]:
    numbers: list[int] = []
    biblio_end = bibliography_end_index(paragraphs, biblio_idx)
    for para in paragraphs[biblio_idx + 1 : biblio_end]:
        text = paragraph_text(para).strip()
        match = re.match(r"^(?:\[(\d+)\]|(\d+)\.)", text)
        if match:
            numbers.append(int(match.group(1) or match.group(2)))
    return numbers


def is_tail_block_heading(text: str) -> bool:
    normalized = normalize(text)
    if not normalized:
        return False
    return any(normalized.startswith(normalize(heading)) for heading in TAIL_BLOCK_HEADINGS)


def bibliography_end_index(paragraphs: list[ET.Element], biblio_idx: int) -> int:
    for idx, para in enumerate(paragraphs[biblio_idx + 1 :], start=biblio_idx + 1):
        text = paragraph_text(para).strip()
        if is_tail_block_heading(text):
            return idx
    return len(paragraphs)


def bibliography_heading_index(paragraphs: list[ET.Element]) -> int | None:
    """Return the real bibliography heading, not a TOC row such as '参考文献39'."""
    candidates: list[tuple[int, int]] = []
    for idx, para in enumerate(paragraphs):
        if not is_bibliography_heading_text(paragraph_text(para)):
            continue
        numbered_followers = 0
        for next_para in paragraphs[idx + 1 : idx + 140]:
            next_text = paragraph_text(next_para).strip()
            if not next_text:
                continue
            if re.match(r"^(?:\[(\d+)\]|(\d+)\.)", next_text):
                numbered_followers += 1
                continue
            if numbered_followers and heading_level(next_text) is not None:
                break
        candidates.append((numbered_followers, idx))
    useful = [(count, idx) for count, idx in candidates if count > 0]
    if useful:
        useful.sort(key=lambda item: (item[0], item[1]), reverse=True)
        return useful[0][1]
    if candidates:
        return candidates[-1][1]
    return None


def bookmark_citation_numbers(root: ET.Element) -> list[int]:
    numbers: list[int] = []
    for bookmark in root.findall(".//w:bookmarkStart", NS):
        name = bookmark.attrib.get(W + "name", "")
        match = re.fullmatch(rf"{BOOKMARK_PREFIX}(\d+)", name or "")
        if match:
            numbers.append(int(match.group(1)))
    return sorted(set(numbers))


def update_bibliography_bookmarks_without_renumber(
    root: ET.Element,
    paragraphs: list[ET.Element],
    biblio_idx: int,
    bookmark_map: dict[int, str],
) -> None:
    bookmark_id = next_bookmark_id(root)
    biblio_end = bibliography_end_index(paragraphs, biblio_idx)
    for para in paragraphs[biblio_idx + 1 : biblio_end]:
        text = paragraph_text(para).strip()
        match = re.match(r"^(?:\[(\d+)\]|(\d+)\.)", text)
        if not match:
            continue
        number = int(match.group(1) or match.group(2))
        if number not in bookmark_map:
            continue
        remove_existing_citation_bookmarks(para)
        add_bibliography_bookmark_preserving_runs(
            para,
            bookmark_name=bookmark_map[number],
            bookmark_id=bookmark_id,
        )
        bookmark_id += 1


def update_bibliography_entries_by_existing_order(
    root: ET.Element,
    paragraphs: list[ET.Element],
    biblio_idx: int,
    bookmark_map: dict[int, str],
    *,
    numbering_style: str = "bracket",
) -> int:
    bookmark_id = next_bookmark_id(root)
    changed = 0
    new_number = 1
    biblio_end = bibliography_end_index(paragraphs, biblio_idx)
    for para in paragraphs[biblio_idx + 1 : biblio_end]:
        text = paragraph_text(para).strip()
        if not re.match(r"^(?:\[(\d+)\]|(\d+)\.)", text):
            continue
        if new_number not in bookmark_map:
            break
        update_bibliography_paragraph_preserving_runs(
            para,
            number=new_number,
            bookmark_name=bookmark_map[new_number],
            bookmark_id=bookmark_id,
            numbering_style=numbering_style,
        )
        bookmark_id += 1
        new_number += 1
        changed += 1
    return changed


def paragraph_has_fast_path_unsupported_payload(paragraph: ET.Element, anchor_map: dict[int, str]) -> bool:
    if paragraph.find(".//w:instrText", NS) is not None:
        return True
    for node in paragraph.iter():
        if node.tag == W + "hyperlink" and citation_hyperlink_ok(node, anchor_map):
            continue
        if node.tag in {
            W + "drawing",
            W + "pict",
            W + "object",
            W + "fldSimple",
            W + "commentRangeStart",
            W + "commentRangeEnd",
        }:
            return True
    return False


def direct_text_nodes(paragraph: ET.Element) -> list[ET.Element]:
    nodes: list[ET.Element] = []
    for run in paragraph.findall("w:r", NS):
        nodes.extend(run.findall("w:t", NS))
    return nodes


def set_text_node_text(node: ET.Element, text: str) -> None:
    if text.startswith(" ") or text.endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif "{http://www.w3.org/XML/1998/namespace}space" in node.attrib:
        del node.attrib["{http://www.w3.org/XML/1998/namespace}space"]
    node.text = text


def replace_visible_prefix(paragraph: ET.Element, old_len: int, new_prefix: str) -> None:
    remaining = old_len
    wrote_prefix = False
    for node in direct_text_nodes(paragraph):
        value = node.text or ""
        if remaining <= 0:
            break
        if len(value) <= remaining:
            set_text_node_text(node, new_prefix if not wrote_prefix else "")
            wrote_prefix = True
            remaining -= len(value)
            continue
        suffix = value[remaining:]
        set_text_node_text(node, (new_prefix if not wrote_prefix else "") + suffix)
        wrote_prefix = True
        remaining = 0
    if remaining > 0 or not wrote_prefix:
        raise ValueError("bibliography entry prefix spans unsupported non-text content; refusing lossy bibliography rebuild")


def replace_bibliography_number_preserving_runs(
    paragraph: ET.Element,
    number: int,
    *,
    numbering_style: str = "bracket",
) -> None:
    text = paragraph_text(paragraph)
    match = re.match(r"^(?:\[(\d+)\]|(\d+)\.)", text)
    if not match:
        raise ValueError(f"bibliography entry does not start with a visible number label: {text[:80]}")
    if numbering_style == "arabic-dot":
        new_prefix = f"{number}."
    elif numbering_style == "preserve":
        new_prefix = f"{number}." if match.group(2) is not None else f"[{number}]"
    elif numbering_style == "bracket":
        new_prefix = f"[{number}]"
    else:
        raise ValueError(f"unsupported bibliography numbering style: {numbering_style}")
    replace_visible_prefix(paragraph, match.end(), new_prefix)


def remove_existing_citation_bookmarks(paragraph: ET.Element) -> None:
    bookmark_ids: set[str] = set()
    for child in list(paragraph):
        if child.tag == W + "bookmarkStart" and child.attrib.get(W + "name", "").startswith(BOOKMARK_PREFIX):
            bookmark_ids.add(child.attrib.get(W + "id", ""))
            paragraph.remove(child)
    for child in list(paragraph):
        if child.tag == W + "bookmarkEnd" and child.attrib.get(W + "id", "") in bookmark_ids:
            paragraph.remove(child)


def add_bibliography_bookmark_preserving_runs(paragraph: ET.Element, *, bookmark_name: str, bookmark_id: int) -> None:
    children = list(paragraph)
    insert_pos = 0
    for idx, child in enumerate(children):
        if child.tag == W + "pPr":
            insert_pos = idx + 1
            continue
        if child.tag == W + "r" and "".join(node.text or "" for node in child.findall(".//w:t", NS)).strip():
            insert_pos = idx
            break
    paragraph.insert(insert_pos, make_bookmark_start(bookmark_id, bookmark_name))
    paragraph.append(make_bookmark_end(bookmark_id))


def update_bibliography_paragraph_preserving_runs(
    paragraph: ET.Element,
    *,
    number: int,
    bookmark_name: str,
    bookmark_id: int,
    numbering_style: str = "bracket",
) -> None:
    remove_existing_citation_bookmarks(paragraph)
    replace_bibliography_number_preserving_runs(paragraph, number, numbering_style=numbering_style)
    add_bibliography_bookmark_preserving_runs(paragraph, bookmark_name=bookmark_name, bookmark_id=bookmark_id)


def delete_paragraph(paragraph: ET.Element, body: ET.Element) -> None:
    body.remove(paragraph)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True)
    parser.add_argument(
        "--assign-all-bibliography-by-occurrence",
        action="store_true",
        help=(
            "Assign the first N visible body citation markers to bibliography entries 1..N by occurrence order, "
            "where N is the bibliography entry count; keep later citation markers on the last bibliography entry."
        ),
    )
    parser.add_argument(
        "--split-grouped-citations",
        action="store_true",
        help="Split grouped numeric markers such as [7,8] into separate sentence-level single-number citation markers.",
    )
    parser.add_argument(
        "--bibliography-numbering-style",
        choices=("bracket", "arabic-dot", "preserve"),
        default="bracket",
        help="Visible bibliography entry label style to write while keeping cite_ref_N bookmarks for body [N] links.",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    with zipfile.ZipFile(docx_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(members["word/document.xml"])
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml missing w:body")

    paragraphs = body.findall(W + "p")
    biblio_idx = bibliography_heading_index(paragraphs)
    if biblio_idx is None:
        raise ValueError(f"Could not find bibliography heading '{BIBLIO_HEADING}'")
    body_start_idx = next((idx for idx, para in enumerate(paragraphs[:biblio_idx]) if heading_level(paragraph_text(para))), None)
    if body_start_idx is None:
        raise ValueError("Could not find a body heading before bibliography; refusing broad citation rewrite")

    body_targets: list[tuple[int, list[int], str]] = []
    appearance_order: list[int] = []
    seen: set[int] = set()

    for idx, para in enumerate(paragraphs[body_start_idx:biblio_idx], start=body_start_idx):
        text = paragraph_text(para).strip()
        numbers = [int(x) for x in RAW_CITATION_RE.findall(text)]
        if not numbers:
            continue
        if paragraph_is_forbidden_citation_surface(para, text):
            raise ValueError(
                f"citation paragraph {idx} appears to be a heading, caption, TOC, or other non-body surface; "
                "refusing to attach citations to protected formatting surfaces"
            )
        clean = RAW_CITATION_RE.sub("", text)
        body_targets.append((idx, numbers, clean))
        for number in numbers:
            if number not in seen:
                seen.add(number)
                appearance_order.append(number)

    if not body_targets and not args.split_grouped_citations:
        print(docx_path)
        return 0

    if args.split_grouped_citations:
        current_paragraphs = body.findall(W + "p")
        biblio_numbers = bibliography_entry_numbers(current_paragraphs, biblio_idx)
        if not biblio_numbers:
            biblio_numbers = bookmark_citation_numbers(root)
        if not biblio_numbers:
            raise ValueError("bibliography has no numbered entries or citation bookmarks; refusing grouped citation split")
        bibliography_count = max(biblio_numbers)
        bookmark_map = {number: f"{BOOKMARK_PREFIX}{number}" for number in range(1, bibliography_count + 1)}
        changed_count = 0
        biblio_end = bibliography_end_index(current_paragraphs, biblio_idx)
        scan_ranges = (
            (body_start_idx, biblio_idx),
            (biblio_idx + 1, biblio_end),
        )
        for scan_start, scan_end in scan_ranges:
            for idx, para in enumerate(current_paragraphs[scan_start:scan_end], start=scan_start):
                text = paragraph_text(para).strip()
                if not GROUPED_CITATION_RE.search(text):
                    continue
                if re.match(r"^(?:\[(\d+)\]|(\d+)\.)", text):
                    continue
                if paragraph_is_forbidden_citation_surface(para, text):
                    raise ValueError(
                        f"grouped citation paragraph {idx} appears to be a heading, caption, TOC, or other non-body surface; "
                        "refusing grouped citation split on protected formatting surfaces"
                    )
                if paragraph_has_unsafe_payload(para):
                    raise ValueError(f"citation paragraph {idx} contains protected payload; refusing grouped citation split")
                new_text = expand_grouped_citation_markers(text)
                used_numbers = [int(x) for x in RAW_CITATION_RE.findall(new_text)]
                out_of_range = [number for number in used_numbers if number not in bookmark_map]
                if out_of_range:
                    raise ValueError(
                        f"citation paragraph {idx} contains markers without bibliography anchors: {out_of_range}"
                    )
                rebuild_paragraph(para, new_text, anchor_map=bookmark_map)
                changed_count += 1
        if changed_count == 0:
            print(docx_path)
            return 0
        update_bibliography_bookmarks_without_renumber(root, current_paragraphs, biblio_idx, bookmark_map)
        members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in members.items():
                zout.writestr(name, data)
        print(docx_path)
        return 0

    if args.assign_all_bibliography_by_occurrence:
        current_paragraphs = body.findall(W + "p")
        biblio_numbers = bibliography_entry_numbers(current_paragraphs, biblio_idx)
        bibliography_count = len(biblio_numbers)
        marker_count = sum(len(numbers) for _, numbers, _ in body_targets)
        if bibliography_count <= 0:
            raise ValueError("bibliography has no numbered entries; refusing occurrence-based citation assignment")
        if marker_count < bibliography_count:
            raise ValueError(
                "not enough body citation markers to cite every bibliography entry: "
                f"{marker_count} markers for {bibliography_count} entries"
            )

        bookmark_map = {number: f"{BOOKMARK_PREFIX}{number}" for number in range(1, bibliography_count + 1)}
        next_number = 1
        overflow_number = 1
        for idx, old_numbers, _ in body_targets:
            para = current_paragraphs[idx]
            if paragraph_has_occurrence_mode_unsupported_payload(para):
                raise ValueError(
                    f"citation paragraph {idx} contains protected content unsupported by occurrence assignment"
                )
            assigned: list[int] = []
            for _old_number in old_numbers:
                if next_number <= bibliography_count:
                    assigned.append(next_number)
                    next_number += 1
                else:
                    assigned.append(overflow_number)
                    overflow_number += 1
                    if overflow_number > bibliography_count:
                        overflow_number = 1
            replace_citation_occurrences_preserving_runs(para, assigned, bookmark_map)

        updated_count = update_bibliography_entries_by_existing_order(
            root,
            current_paragraphs,
            biblio_idx,
            bookmark_map,
            numbering_style=args.bibliography_numbering_style,
        )
        if updated_count != bibliography_count:
            raise ValueError(
                f"bibliography numbered-entry count changed during occurrence assignment: "
                f"expected {bibliography_count}, updated {updated_count}"
            )

        members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in members.items():
                zout.writestr(name, data)
        print(docx_path)
        return 0

    renumber_map = {old: new for new, old in enumerate(appearance_order, start=1)}
    bookmark_map = {new: f"{BOOKMARK_PREFIX}{new}" for new in range(1, len(appearance_order) + 1)}
    biblio_numbers = bibliography_entry_numbers(paragraphs, biblio_idx)
    expected_numbers = list(range(1, len(appearance_order) + 1))
    if appearance_order == expected_numbers and sorted(biblio_numbers) == expected_numbers:
        current_paragraphs = body.findall(W + "p")
        for idx, _, _ in body_targets:
            para = current_paragraphs[idx]
            if paragraph_has_fast_path_unsupported_payload(para, bookmark_map):
                raise ValueError(f"citation paragraph {idx} contains unsupported protected content; refusing citation link repair")
            fixed_text = normalize_citation_punctuation_in_text(paragraph_text(para).strip())
            if fixed_text != paragraph_text(para).strip():
                rebuild_paragraph(para, fixed_text, anchor_map=bookmark_map)
            else:
                linkify_existing_citation_runs(para, bookmark_map)
        update_bibliography_bookmarks_without_renumber(root, current_paragraphs, biblio_idx, bookmark_map)
        members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in members.items():
                zout.writestr(name, data)
        print(docx_path)
        return 0

    current_paragraphs = body.findall(W + "p")
    for idx, _, _ in body_targets:
        para = current_paragraphs[idx]
        if paragraph_has_unsafe_payload(para):
            raise ValueError(f"citation paragraph {idx} contains fields, drawings, bookmarks, comments, or hyperlinks; refusing lossy rebuild")
    for idx, old_numbers, clean in body_targets:
        para = current_paragraphs[idx]
        if paragraph_requires_lossy_text_run_rebuild(para):
            renumber_and_linkify_existing_citation_runs(para, renumber_map, bookmark_map)
        else:
            new_numbers = [renumber_map[number] for number in old_numbers if number in renumber_map]
            rebuilt = attach_markers(split_sentences(clean), new_numbers)
            rebuild_paragraph(para, rebuilt, anchor_map=bookmark_map)

    current_paragraphs = body.findall(W + "p")
    biblio_end = bibliography_end_index(current_paragraphs, biblio_idx)
    bibliography_paragraphs = current_paragraphs[biblio_idx + 1 : biblio_end]
    kept_entries: list[tuple[int, str, ET.Element]] = []
    unused_entries: list[int] = []
    for para in list(bibliography_paragraphs):
        text = paragraph_text(para).strip()
        match = re.match(r"^(?:\[(\d+)\]|(\d+)\.)(.*)$", text)
        if not match:
            continue
        old_number = int(match.group(1) or match.group(2))
        remainder = match.group(3).lstrip()
        if old_number not in renumber_map:
            unused_entries.append(old_number)
            continue
        kept_entries.append((renumber_map[old_number], remainder, para))
    if unused_entries:
        raise ValueError(
            "bibliography contains entries not cited in the authorized body range; refusing to delete entries: "
            + ", ".join(str(number) for number in unused_entries)
        )

    kept_entries.sort(key=lambda item: item[0])
    for _, _, para in kept_entries:
        delete_paragraph(para, body)

    heading_para = body.findall(W + "p")[biblio_idx]
    insert_pos = list(body).index(heading_para) + 1
    bookmark_id = next_bookmark_id(root)
    for new_number, _, para in kept_entries:
        update_bibliography_paragraph_preserving_runs(
            para,
            number=new_number,
            bookmark_name=bookmark_map[new_number],
            bookmark_id=bookmark_id,
            numbering_style=args.bibliography_numbering_style,
        )
        bookmark_id += 1
        body.insert(insert_pos, para)
        insert_pos += 1

    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)
    print(docx_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
