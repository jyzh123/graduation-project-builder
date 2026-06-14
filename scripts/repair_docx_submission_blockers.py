#!/usr/bin/env python3
"""Bounded DOCX repair for submission-blocking thesis surfaces.

This helper is intentionally narrow: it patches only ``word/document.xml`` in a
package-preserving way and reports source/final SHA values. It does not touch
media, relationships, comments, citations, styles, headers, or footers.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)


def qn(local: str) -> str:
    prefix, name = local.split(":", 1)
    if prefix == "w":
        return f"{{{W_NS}}}{name}"
    if prefix == "xml":
        return f"{{{XML_NS}}}{name}"
    raise ValueError(local)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t")))


def paragraph_style_id(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", NS)
    return node.attrib.get(qn("w:val"), "") if node is not None else ""


def has_image(paragraph: ET.Element) -> bool:
    return any(node.tag.endswith("}drawing") or node.tag.endswith("}pict") for node in paragraph.iter())


def has_page_break(paragraph: ET.Element) -> bool:
    return any(node.attrib.get(qn("w:type")) == "page" for node in paragraph.iter(qn("w:br")))


def has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def heading_level(text: str) -> int | None:
    compact = compact_text(text)
    if re.fullmatch(r"(第[一二三四五六七八九十]+章|[1-9]\d*)[^。；;,.，、]{0,30}", compact):
        return 1
    if re.fullmatch(r"[1-9]\d*\.\d+[^。；;,.，、]{0,50}", compact):
        return 2
    if re.fullmatch(r"[1-9]\d*\.\d+\.\d+[^。；;,.，、]{0,50}", compact):
        return 3
    return None


CAPTION_RE = re.compile(r"^\s*[图表]\s*\d")


def remove_blank_paragraphs_between_visible_thesis_paragraphs(body: ET.Element) -> int:
    """Remove only blank prose separators that pre-submission gates reject.

    This intentionally preserves paragraphs carrying page breaks, section
    breaks, pictures, captions, headings, and image-adjacent spacing.
    """
    paragraphs = body_paragraphs(body)
    texts = [paragraph_text(paragraph) for paragraph in paragraphs]
    first_body = next((idx for idx, text in enumerate(texts) if heading_level(text) == 1), 0)
    references = next(
        (
            idx
            for idx, text in enumerate(texts)
            if compact_text(text) in {"参考文献", "references", "bibliography"}
        ),
        len(paragraphs),
    )
    doomed: list[ET.Element] = []
    for idx, paragraph in enumerate(paragraphs):
        if idx <= first_body or idx >= references:
            continue
        if texts[idx].strip():
            continue
        if has_image(paragraph) or has_page_break(paragraph) or has_section_break(paragraph):
            continue
        if idx == 0 or idx + 1 >= len(paragraphs):
            continue
        prev = paragraphs[idx - 1]
        next_paragraph = paragraphs[idx + 1]
        prev_text = texts[idx - 1].strip()
        next_text = texts[idx + 1].strip()
        if not prev_text or not next_text:
            continue
        if has_image(prev) or has_image(next_paragraph):
            continue
        if CAPTION_RE.match(prev_text) or CAPTION_RE.match(next_text):
            continue
        if heading_level(prev_text) is not None or heading_level(next_text) is not None:
            continue
        doomed.append(paragraph)
    for paragraph in doomed:
        body.remove(paragraph)
    return len(doomed)


def ensure_keep_lines_for_paragraphs(body: ET.Element, needles: list[str]) -> dict[str, object]:
    targets = [needle for needle in needles if str(needle or "").strip()]
    if not targets:
        return {"status": "skipped", "requested": 0, "matched": 0, "changed": 0, "unmatched": []}
    matched: set[str] = set()
    changed = 0
    for paragraph in body_paragraphs(body):
        text = paragraph_text(paragraph)
        matched_needles = [needle for needle in targets if needle in text]
        if not matched_needles:
            continue
        matched.update(matched_needles)
        ppr = paragraph.find("./w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)
        if ppr.find("./w:keepLines", NS) is None:
            style = ppr.find("./w:pStyle", NS)
            keep_lines = ET.Element(qn("w:keepLines"))
            if style is not None:
                children = list(ppr)
                ppr.insert(children.index(style) + 1, keep_lines)
            else:
                ppr.insert(0, keep_lines)
            changed += 1
    unmatched = [needle for needle in targets if needle not in matched]
    return {
        "status": "passed" if not unmatched else "failed",
        "requested": len(targets),
        "matched": len(matched),
        "changed": changed,
        "unmatched": unmatched,
    }


def ensure_page_break_before_paragraphs(body: ET.Element, needles: list[str]) -> dict[str, object]:
    targets = [needle for needle in needles if str(needle or "").strip()]
    if not targets:
        return {"status": "skipped", "requested": 0, "matched": 0, "changed": 0, "unmatched": []}
    matched: set[str] = set()
    changed = 0
    for paragraph in body_paragraphs(body):
        text = paragraph_text(paragraph)
        matched_needles = [needle for needle in targets if needle in text]
        if not matched_needles:
            continue
        matched.update(matched_needles)
        ppr = paragraph.find("./w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)
        if ppr.find("./w:pageBreakBefore", NS) is None:
            style = ppr.find("./w:pStyle", NS)
            page_break_before = ET.Element(qn("w:pageBreakBefore"))
            if style is not None:
                children = list(ppr)
                ppr.insert(children.index(style) + 1, page_break_before)
            else:
                ppr.insert(0, page_break_before)
            changed += 1
    unmatched = [needle for needle in targets if needle not in matched]
    return {
        "status": "passed" if not unmatched else "failed",
        "requested": len(targets),
        "matched": len(matched),
        "changed": changed,
        "unmatched": unmatched,
    }


def remove_page_break_after_paragraphs(body: ET.Element, needles: list[str]) -> dict[str, object]:
    targets = [needle for needle in needles if str(needle or "").strip()]
    if not targets:
        return {"status": "skipped", "requested": 0, "matched": 0, "removed": 0, "unmatched": []}
    matched: set[str] = set()
    removed = 0
    children = list(body)
    for child in list(children):
        if child.tag != qn("w:p"):
            continue
        text = paragraph_text(child)
        matched_needles = [needle for needle in targets if needle in text]
        if not matched_needles:
            continue
        matched.update(matched_needles)
        current_children = list(body)
        try:
            index = current_children.index(child)
        except ValueError:
            continue
        if index + 1 >= len(current_children):
            continue
        next_child = current_children[index + 1]
        if next_child.tag != qn("w:p"):
            continue
        if paragraph_text(next_child).strip():
            continue
        if has_image(next_child) or has_section_break(next_child):
            continue
        if not has_page_break(next_child):
            continue
        body.remove(next_child)
        removed += 1
    unmatched = [needle for needle in targets if needle not in matched]
    return {
        "status": "passed" if not unmatched else "failed",
        "requested": len(targets),
        "matched": len(matched),
        "removed": removed,
        "unmatched": unmatched,
    }


def has_cover_title_label(text: str) -> bool:
    return bool(re.search(r"(?:^|[\s\u3000])(?:\u8bba\s*\u6587\s*)?\u9898\s*\u76ee\s*[:\uff1a]", text or ""))


def is_front_matter_boundary(text: str) -> bool:
    compact = compact_text(text)
    return compact in {
        "\u6458\u8981",
        "abstract",
        "\u76ee\u5f55",
        "\u5b66\u4f4d\u8bba\u6587\u539f\u521b\u6027\u58f0\u660e",
        "\u5b66\u4f4d\u8bba\u6587\u7248\u6743\u4f7f\u7528\u6388\u6743\u4e66",
    } or heading_level(text) == 1


def is_likely_cover_title(text: str) -> bool:
    stripped = (text or "").strip()
    compact = compact_text(stripped)
    if len(compact) < 12:
        return False
    if has_cover_title_label(stripped):
        return False
    blocked_fragments = (
        "\u5b66\u58eb\u5b66\u4f4d\u8bba\u6587",
        "\u6bd5\u4e1a\u8bbe\u8ba1",
        "\u4f5c\u8005\u59d3\u540d",
        "\u5b66\u751f\u59d3\u540d",
        "\u5b66\u53f7",
        "\u5b66\u9662",
        "\u9662\u7cfb",
        "\u4e13\u4e1a",
        "\u6307\u5bfc\u6559\u5e08",
    )
    if any(fragment in compact for fragment in blocked_fragments):
        return False
    cjk_count = sum(1 for char in stripped if "\u4e00" <= char <= "\u9fff")
    return cjk_count >= 8


def prefix_cover_title_label(body: ET.Element) -> dict[str, object]:
    paragraphs = body_paragraphs(body)
    for paragraph in paragraphs[:40]:
        text = paragraph_text(paragraph)
        if is_front_matter_boundary(text):
            break
        if has_cover_title_label(text):
            return {"status": "already-present", "matched": 1, "changed": 0, "paragraph_text": text[:120]}
    for paragraph in paragraphs[:25]:
        text = paragraph_text(paragraph)
        if is_front_matter_boundary(text):
            break
        if not is_likely_cover_title(text):
            continue
        text_node = paragraph.find(".//w:t", NS)
        if text_node is None:
            return {"status": "failed", "matched": 1, "changed": 0, "issue": "cover title paragraph has no text node"}
        text_node.text = "\u9898\u76ee\uff1a" + (text_node.text or "")
        return {"status": "passed", "matched": 1, "changed": 1, "paragraph_text_before": text[:120]}
    return {"status": "failed", "matched": 0, "changed": 0, "issue": "cover title paragraph not found before front matter boundary"}


def body_paragraphs(body: ET.Element) -> list[ET.Element]:
    return [child for child in list(body) if child.tag == qn("w:p")]


def clone_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    rpr = paragraph.find(".//w:r/w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def make_text_run(text: str, rpr: ET.Element | None = None) -> ET.Element:
    run = ET.Element(qn("w:r"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = ET.SubElement(run, qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set(qn("xml:space"), "preserve")
    t.text = text
    return run


def first_text_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run).strip():
            rpr = run.find("./w:rPr", NS)
            return copy.deepcopy(rpr) if rpr is not None else None
    for run in paragraph.findall(".//w:r", NS):
        if paragraph_text(run).strip():
            rpr = run.find("./w:rPr", NS)
            return copy.deepcopy(rpr) if rpr is not None else None
    return None


def ensure_clean_citation_rpr(rpr: ET.Element | None = None) -> ET.Element:
    result = copy.deepcopy(rpr) if rpr is not None else ET.Element(qn("w:rPr"))
    for tag in ("w:b", "w:i", "w:rStyle", "w:highlight"):
        for child in list(result.findall(tag, NS)):
            result.remove(child)
    color = result.find("./w:color", NS)
    if color is None:
        color = ET.SubElement(result, qn("w:color"))
    color.set(qn("w:val"), "000000")
    underline = result.find("./w:u", NS)
    if underline is None:
        underline = ET.SubElement(result, qn("w:u"))
    underline.set(qn("w:val"), "none")
    vert = result.find("./w:vertAlign", NS)
    if vert is None:
        vert = ET.SubElement(result, qn("w:vertAlign"))
    vert.set(qn("w:val"), "superscript")
    return result


def existing_citation_rpr(root: ET.Element) -> ET.Element:
    for run in root.findall(".//w:r", NS):
        text = paragraph_text(run)
        if re.fullmatch(r"\[\d+\]", text or ""):
            rpr = run.find("./w:rPr", NS)
            return ensure_clean_citation_rpr(rpr)
    return ensure_clean_citation_rpr()


def make_hyperlink_citation(number: int, rpr: ET.Element) -> ET.Element:
    hyperlink = ET.Element(qn("w:hyperlink"))
    hyperlink.set(qn("w:anchor"), f"cite_ref_{number}")
    hyperlink.set(qn("w:history"), "1")
    hyperlink.append(make_text_run(f"[{number}]", rpr))
    return hyperlink


def text_char_family(char: str) -> str:
    if "\u4e00" <= char <= "\u9fff":
        return "cjk"
    if char.isascii() and (char.isalnum() or char in " -_./:&+()[]'\""):
        return "latin"
    return "neutral"


def next_non_neutral_family(chars: list[str], start: int) -> str | None:
    for char in chars[start:]:
        family = text_char_family(char)
        if family != "neutral":
            return family
    return None


def split_body_text_by_script(text: str) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    chars = list(text or "")
    current_family: str | None = None
    current: list[str] = []
    for index, char in enumerate(chars):
        family = text_char_family(char)
        if family == "neutral":
            family = current_family or next_non_neutral_family(chars, index + 1) or "cjk"
        if current_family is None:
            current_family = family
        if family != current_family:
            if current:
                segments.append((current_family, "".join(current)))
            current_family = family
            current = [char]
        else:
            current.append(char)
    if current and current_family:
        segments.append((current_family, "".join(current)))
    return [(family, value) for family, value in segments if value]


def body_latin_rpr(rpr: ET.Element | None) -> ET.Element:
    result = copy.deepcopy(rpr) if rpr is not None else ET.Element(qn("w:rPr"))
    fonts = result.find("./w:rFonts", NS)
    if fonts is None:
        fonts = ET.Element(qn("w:rFonts"))
        result.insert(0, fonts)
    for key in ("ascii", "hAnsi", "cs"):
        fonts.set(qn(f"w:{key}"), "Times New Roman")
        fonts.attrib.pop(qn(f"w:{key}Theme"), None)
    return result


def append_body_text_runs(paragraph: ET.Element, text: str, cjk_rpr: ET.Element | None, latin_rpr: ET.Element) -> None:
    for family, value in split_body_text_by_script(text):
        paragraph.append(make_text_run(value, latin_rpr if family == "latin" else cjk_rpr))


def next_bookmark_id(root: ET.Element) -> int:
    values = []
    for node in root.findall(".//w:bookmarkStart", NS):
        raw = node.get(qn("w:id"), "")
        if raw.isdigit():
            values.append(int(raw))
    return max(values, default=0) + 1


def make_bookmark_start(bookmark_id: int, name: str) -> ET.Element:
    node = ET.Element(qn("w:bookmarkStart"))
    node.set(qn("w:id"), str(bookmark_id))
    node.set(qn("w:name"), name)
    return node


def make_bookmark_end(bookmark_id: int) -> ET.Element:
    node = ET.Element(qn("w:bookmarkEnd"))
    node.set(qn("w:id"), str(bookmark_id))
    return node


def clone_text_paragraph_from_donor(
    donor: ET.Element,
    text: str,
    *,
    bookmark_name: str | None = None,
    bookmark_id: int | None = None,
) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ppr = donor.find("./w:pPr", NS)
    if ppr is not None:
        paragraph.append(copy.deepcopy(ppr))
    rpr = first_text_rpr(donor)
    if bookmark_name and bookmark_id is not None:
        paragraph.append(make_bookmark_start(bookmark_id, bookmark_name))
    paragraph.append(make_text_run(text, rpr))
    if bookmark_name and bookmark_id is not None:
        paragraph.append(make_bookmark_end(bookmark_id))
    return paragraph


def replace_text_node_fragments(root: ET.Element, replacements: list[dict[str, str]]) -> list[dict[str, object]]:
    results: list[dict[str, object]] = []
    for item in replacements:
        old = str(item.get("old") or "")
        new = str(item.get("new") or "")
        if not old:
            continue
        count = 0
        for node in root.findall(".//w:t", NS):
            if node.text and old in node.text:
                node.text = node.text.replace(old, new)
                if node.text.startswith(" ") or node.text.endswith(" "):
                    node.set(qn("xml:space"), "preserve")
                count += 1
        results.append({"old": old, "new": new, "count": count})
    return results


def append_tail_citation_text(root: ET.Element, body: ET.Element, plan: dict[str, object]) -> dict[str, object]:
    append_plan = plan.get("citation_append")
    if not isinstance(append_plan, dict):
        return {"status": "skipped"}
    host_contains = str(append_plan.get("host_contains") or "")
    pieces = append_plan.get("pieces")
    if not host_contains or not isinstance(pieces, list):
        raise RuntimeError("tail plan citation_append requires host_contains and pieces")
    paragraphs = body_paragraphs(body)
    candidates = [p for p in paragraphs if host_contains in paragraph_text(p)]
    if len(candidates) != 1:
        raise RuntimeError(f"citation_append host match count must be 1, found {len(candidates)}")
    paragraph = candidates[0]
    existing_text = paragraph_text(paragraph)
    already_present_numbers = set(int(n) for n in re.findall(r"\[(\d+)\]", existing_text))
    prose_rpr = first_text_rpr(paragraph)
    latin_rpr = body_latin_rpr(prose_rpr)
    citation_rpr = existing_citation_rpr(root)
    inserted_numbers: list[int] = []
    for piece in pieces:
        if not isinstance(piece, dict):
            continue
        if "text" in piece:
            text = str(piece.get("text") or "")
            if text:
                append_body_text_runs(paragraph, text, prose_rpr, latin_rpr)
        elif "citation" in piece:
            number = int(piece["citation"])
            if number in already_present_numbers:
                continue
            paragraph.append(make_hyperlink_citation(number, citation_rpr))
            inserted_numbers.append(number)
    return {"status": "passed", "paragraph_text_prefix": existing_text[:80], "inserted_numbers": inserted_numbers}


def find_bibliography_heading(paragraphs: list[ET.Element]) -> int | None:
    for index, paragraph in enumerate(paragraphs):
        if compact_text(paragraph_text(paragraph)) in {"参考文献", "references", "bibliography"}:
            return index
    return None


def find_ack_heading(paragraphs: list[ET.Element]) -> int | None:
    for index, paragraph in enumerate(paragraphs):
        compact = compact_text(paragraph_text(paragraph)).lower()
        if compact in {"致谢", "谢辞", "acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments"}:
            return index
    return None


def first_body_text_donor(paragraphs: list[ET.Element], bibliography_index: int) -> ET.Element:
    body_started = False
    fallback: ET.Element | None = None
    for paragraph in paragraphs[:bibliography_index]:
        text = paragraph_text(paragraph).strip()
        if heading_level(text) == 1:
            body_started = True
            continue
        if not body_started or not text:
            continue
        if heading_level(text) is not None or CAPTION_RE.match(text) or has_image(paragraph):
            continue
        fallback = paragraph
    if fallback is None:
        raise RuntimeError("could not find body prose donor before bibliography")
    return fallback


def tail_block_end_index(children: list[ET.Element], paragraphs: list[ET.Element], start_para_index: int) -> int:
    start_child_index = children.index(paragraphs[start_para_index])
    for index in range(start_child_index + 1, len(children)):
        child = children[index]
        if child.tag != qn("w:p"):
            continue
        text = paragraph_text(child).strip()
        if not text:
            continue
        compact = compact_text(text).lower()
        if compact in {"致谢", "谢辞", "附录", "appendix", "acknowledgement", "acknowledgements"}:
            return index
        if heading_level(text) == 1 and compact not in {"参考文献", "references", "bibliography"}:
            return index
    return len(children)


def append_bibliography_entries(root: ET.Element, body: ET.Element, entries: list[str]) -> dict[str, object]:
    if not entries:
        return {"status": "skipped", "inserted": 0}
    paragraphs = body_paragraphs(body)
    bibliography_index = find_bibliography_heading(paragraphs)
    if bibliography_index is None:
        raise RuntimeError("extra references require an existing bibliography heading")
    children = list(body)
    block_end = tail_block_end_index(children, paragraphs, bibliography_index)
    bibliography_children = children[children.index(paragraphs[bibliography_index]) + 1 : block_end]
    donor = None
    insert_at = block_end
    existing_text = {paragraph_text(child).strip() for child in bibliography_children if child.tag == qn("w:p")}
    for offset, child in enumerate(bibliography_children):
        if child.tag != qn("w:p") or not paragraph_text(child).strip():
            continue
        donor = child
        insert_at = children.index(child) + 1
    if donor is None:
        raise RuntimeError("could not find bibliography entry donor")
    bookmark_id = next_bookmark_id(root)
    inserted: list[str] = []
    for offset, entry_text in enumerate(entries):
        if entry_text in existing_text:
            continue
        number = len([c for c in bibliography_children if c.tag == qn("w:p") and paragraph_text(c).strip()]) + len(inserted) + 1
        paragraph = clone_text_paragraph_from_donor(
            donor,
            entry_text,
            bookmark_name=f"cite_ref_{number}",
            bookmark_id=bookmark_id,
        )
        bookmark_id += 1
        body.insert(insert_at + len(inserted), paragraph)
        inserted.append(entry_text)
    return {"status": "passed", "inserted": len(inserted), "inserted_prefixes": [item[:80] for item in inserted]}


def insert_acknowledgement(body: ET.Element, plan: dict[str, object]) -> dict[str, object]:
    ack_plan = plan.get("acknowledgement")
    if not isinstance(ack_plan, dict):
        return {"status": "skipped"}
    paragraphs = body_paragraphs(body)
    if find_ack_heading(paragraphs) is not None:
        return {"status": "already-present"}
    bibliography_index = find_bibliography_heading(paragraphs)
    if bibliography_index is None:
        raise RuntimeError("acknowledgement insertion requires bibliography heading")
    title = str(ack_plan.get("title") or "致谢")
    content = [str(item) for item in (ack_plan.get("paragraphs") or []) if str(item).strip()]
    if not content:
        raise RuntimeError("acknowledgement plan requires non-empty paragraphs")
    children = list(body)
    block_end = tail_block_end_index(children, paragraphs, bibliography_index)
    reference_heading = paragraphs[bibliography_index]
    body_donor = first_body_text_donor(paragraphs, bibliography_index)
    heading = clone_text_paragraph_from_donor(reference_heading, title)
    ppr = heading.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("w:pPr"))
        heading.insert(0, ppr)
    if ppr.find("./w:pageBreakBefore", NS) is None:
        ppr.insert(0, ET.Element(qn("w:pageBreakBefore")))
    insert_nodes = [heading] + [clone_text_paragraph_from_donor(body_donor, paragraph) for paragraph in content]
    for offset, node in enumerate(insert_nodes):
        body.insert(block_end + offset, node)
    return {"status": "passed", "paragraphs_inserted": len(insert_nodes), "title": title}


def apply_tail_content_plan(root: ET.Element, body: ET.Element, tail_plan: dict[str, object] | None) -> dict[str, object]:
    if not tail_plan:
        return {"status": "skipped"}
    entries = [str(item) for item in (tail_plan.get("extra_references") or []) if str(item).strip()]
    return {
        "status": "passed",
        "text_replacements": replace_text_node_fragments(root, tail_plan.get("text_replacements") or []),
        "citation_append": append_tail_citation_text(root, body, tail_plan),
        "extra_references": append_bibliography_entries(root, body, entries),
        "acknowledgement": insert_acknowledgement(body, tail_plan),
    }


def make_tab_run(rpr: ET.Element | None = None) -> ET.Element:
    run = ET.Element(qn("w:r"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    ET.SubElement(run, qn("w:tab"))
    return run


def set_paragraph_runs(paragraph: ET.Element, runs: list[ET.Element]) -> None:
    for child in list(paragraph):
        if child.tag != qn("w:pPr"):
            paragraph.remove(child)
    for run in runs:
        paragraph.append(run)


def make_toc_entry_from(template: ET.Element, label: str, page: str) -> ET.Element:
    paragraph = copy.deepcopy(template)
    rpr = clone_run_rpr(template)
    set_paragraph_runs(
        paragraph,
        [make_text_run(label, rpr), make_tab_run(rpr), make_text_run(page, rpr)],
    )
    return paragraph


def normalize_square_placeholders(root: ET.Element) -> int:
    changed = 0
    for node in root.iter(qn("w:t")):
        if node.text and "\u25a1" in node.text:
            node.text = node.text.replace("\u25a1", " ")
            changed += 1
    return changed


def find_toc_range(paragraphs: list[ET.Element]) -> tuple[int, int] | None:
    start = None
    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph).strip().replace(" ", "")
        if text == "\u76ee\u5f55":
            start = index + 1
            break
    if start is None:
        return None
    end = start
    for index in range(start, len(paragraphs)):
        style_id = paragraph_style_id(paragraphs[index])
        if style_id in {"TOC1", "TOC2", "TOC3"}:
            end = index + 1
            continue
        if end > start:
            break
    return (start, end) if end > start else None


def add_front_matter_toc_entries(body: ET.Element) -> int:
    paragraphs = body_paragraphs(body)
    toc_range = find_toc_range(paragraphs)
    if toc_range is None:
        return 0
    start, _end = toc_range
    toc_texts = [paragraph_text(paragraph) for paragraph in paragraphs[start:]]
    labels = {text.split("\t", 1)[0].strip().lower() for text in toc_texts}
    insertions: list[ET.Element] = []
    template = paragraphs[start]
    if "\u6458\u8981" not in labels:
        insertions.append(make_toc_entry_from(template, "\u6458\u8981", "I"))
    if "abstract" not in labels:
        insertions.append(make_toc_entry_from(template, "ABSTRACT", "II"))
    body_children = list(body)
    anchor = paragraphs[start]
    body_index = body_children.index(anchor)
    for offset, paragraph in enumerate(insertions):
        body.insert(body_index + offset, paragraph)
    return len(insertions)


def add_live_toc_field(body: ET.Element) -> int:
    paragraphs = body_paragraphs(body)
    toc_range = find_toc_range(paragraphs)
    if toc_range is None:
        return 0
    start, end = toc_range
    existing_instr = " ".join(node.text or "" for node in body.iter(qn("w:instrText")))
    if re.search(r"(^|\s)TOC(\s|$)", existing_instr, re.IGNORECASE):
        return 0
    first = paragraphs[start]
    begin_run = ET.Element(qn("w:r"))
    begin_char = ET.SubElement(begin_run, qn("w:fldChar"))
    begin_char.set(qn("w:fldCharType"), "begin")
    instr_run = ET.Element(qn("w:r"))
    instr = ET.SubElement(instr_run, qn("w:instrText"))
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    separate_run = ET.Element(qn("w:r"))
    separate = ET.SubElement(separate_run, qn("w:fldChar"))
    separate.set(qn("w:fldCharType"), "separate")
    insert_pos = 1 if first.find("./w:pPr", NS) is not None else 0
    first.insert(insert_pos, separate_run)
    first.insert(insert_pos, instr_run)
    first.insert(insert_pos, begin_run)
    paragraphs = body_paragraphs(body)
    _start, end = find_toc_range(paragraphs) or (start, end)
    last = paragraphs[end - 1]
    end_run = ET.Element(qn("w:r"))
    end_char = ET.SubElement(end_run, qn("w:fldChar"))
    end_char.set(qn("w:fldCharType"), "end")
    last.append(end_run)
    return 1


def make_paragraph(text: str, *, style_id: str = "Normal", center: bool = False) -> ET.Element:
    p = ET.Element(qn("w:p"))
    ppr = ET.SubElement(p, qn("w:pPr"))
    pstyle = ET.SubElement(ppr, qn("w:pStyle"))
    pstyle.set(qn("w:val"), style_id)
    if center:
        jc = ET.SubElement(ppr, qn("w:jc"))
        jc.set(qn("w:val"), "center")
    p.append(make_text_run(text))
    return p


def make_cell(text: str, width: str, *, header: bool = False) -> ET.Element:
    tc = ET.Element(qn("w:tc"))
    tcpr = ET.SubElement(tc, qn("w:tcPr"))
    tcw = ET.SubElement(tcpr, qn("w:tcW"))
    tcw.set(qn("w:w"), width)
    tcw.set(qn("w:type"), "dxa")
    if header:
        borders = ET.SubElement(tcpr, qn("w:tcBorders"))
        bottom = ET.SubElement(borders, qn("w:bottom"))
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
    tc.append(make_paragraph(text))
    return tc


def make_table() -> ET.Element:
    rows = [
        ["\u6d4b\u8bd5\u9879", "\u7ed3\u679c", "\u8bf4\u660e"],
        ["\u89c6\u9891\u4e0a\u4f20\u4e0e\u89e3\u6790", "\u901a\u8fc7", "\u53ef\u8bfb\u53d6\u89c6\u9891\u57fa\u672c\u4fe1\u606f\u5e76\u5efa\u7acb\u68c0\u6d4b\u4efb\u52a1"],
        ["\u5f02\u5e38\u884c\u4e3a\u5224\u5b9a", "\u901a\u8fc7", "\u80fd\u8f93\u51fa\u4f4e\u5934\u3001\u8d77\u7acb\u3001\u8f6c\u5934\u7b49\u98ce\u9669\u7c7b\u578b"],
        ["\u7ed3\u679c\u7edf\u8ba1\u4e0e\u56de\u653e", "\u901a\u8fc7", "\u7edf\u8ba1\u3001\u5217\u8868\u548c\u5386\u53f2\u56de\u653e\u80fd\u5f62\u6210\u95ed\u73af"],
    ]
    widths = ["2600", "1800", "5000"]
    tbl = ET.Element(qn("w:tbl"))
    tblpr = ET.SubElement(tbl, qn("w:tblPr"))
    tblw = ET.SubElement(tblpr, qn("w:tblW"))
    tblw.set(qn("w:w"), "9400")
    tblw.set(qn("w:type"), "dxa")
    jc = ET.SubElement(tblpr, qn("w:jc"))
    jc.set(qn("w:val"), "center")
    borders = ET.SubElement(tblpr, qn("w:tblBorders"))
    for name, val, size in (
        ("top", "single", "8"),
        ("left", "nil", "0"),
        ("bottom", "single", "8"),
        ("right", "nil", "0"),
        ("insideH", "nil", "0"),
        ("insideV", "nil", "0"),
    ):
        border = ET.SubElement(borders, qn(f"w:{name}"))
        border.set(qn("w:val"), val)
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
    grid = ET.SubElement(tbl, qn("w:tblGrid"))
    for width in widths:
        col = ET.SubElement(grid, qn("w:gridCol"))
        col.set(qn("w:w"), width)
    for row_index, row_values in enumerate(rows):
        tr = ET.SubElement(tbl, qn("w:tr"))
        for text, width in zip(row_values, widths):
            tr.append(make_cell(text, width, header=row_index == 0))
    return tbl


def add_results_table(body: ET.Element) -> int:
    if body.find(".//w:tbl", NS) is not None:
        return 0
    paragraphs = body_paragraphs(body)
    anchor = None
    for paragraph in paragraphs:
        text = paragraph_text(paragraph).strip()
        if text.startswith("5.3") or text.startswith("5\uff0e3") or text.startswith("5\uff0e 3"):
            anchor = paragraph
            break
    if anchor is None:
        for paragraph in paragraphs:
            if paragraph_text(paragraph).strip().startswith("6"):
                anchor = paragraph
                break
    if anchor is None:
        return 0
    caption = make_paragraph("\u88685-1  \u7cfb\u7edf\u6d4b\u8bd5\u7ed3\u679c\u6c47\u603b", center=True)
    table = make_table()
    children = list(body)
    index = children.index(anchor)
    body.insert(index, caption)
    body.insert(index + 1, table)
    return 1


def patch_document_xml(
    xml_bytes: bytes,
    *,
    empty_paragraphs_only: bool = False,
    tail_plan: dict[str, object] | None = None,
    keep_lines_containing: list[str] | None = None,
    page_break_before_containing: list[str] | None = None,
    remove_page_break_after_containing: list[str] | None = None,
    no_blank_removal: bool = False,
    prefix_cover_title: bool = False,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    removed_blank_paragraphs = 0 if no_blank_removal else remove_blank_paragraphs_between_visible_thesis_paragraphs(body)
    keep_lines_report = ensure_keep_lines_for_paragraphs(body, keep_lines_containing or [])
    page_break_before_report = ensure_page_break_before_paragraphs(
        body,
        page_break_before_containing or [],
    )
    removed_page_break_after_report = remove_page_break_after_paragraphs(
        body,
        remove_page_break_after_containing or [],
    )
    cover_title_label_report = (
        prefix_cover_title_label(body)
        if prefix_cover_title
        else {"status": "skipped", "matched": 0, "changed": 0}
    )
    if empty_paragraphs_only:
        report = {
            "square_placeholder_text_nodes_changed": 0,
            "front_matter_toc_entries_added": 0,
            "live_toc_fields_added": 0,
            "results_tables_added": 0,
            "blank_between_visible_paragraphs_removed": removed_blank_paragraphs,
            "keep_lines_containing": keep_lines_report,
            "page_break_before_containing": page_break_before_report,
            "remove_page_break_after_containing": removed_page_break_after_report,
            "cover_title_label": cover_title_label_report,
            "tail_content_plan": apply_tail_content_plan(root, body, tail_plan),
        }
        return ET.tostring(root, encoding="utf-8", xml_declaration=True), report
    report = {
        "square_placeholder_text_nodes_changed": normalize_square_placeholders(root),
        "front_matter_toc_entries_added": add_front_matter_toc_entries(body),
        "live_toc_fields_added": add_live_toc_field(body),
        "results_tables_added": add_results_table(body),
        "blank_between_visible_paragraphs_removed": removed_blank_paragraphs,
        "keep_lines_containing": keep_lines_report,
        "page_break_before_containing": page_break_before_report,
        "remove_page_break_after_containing": removed_page_break_after_report,
        "cover_title_label": cover_title_label_report,
        "tail_content_plan": apply_tail_content_plan(root, body, tail_plan),
    }
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), report


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    empty_paragraphs_only: bool = False,
    tail_plan_json: Path | None = None,
    keep_lines_containing: list[str] | None = None,
    page_break_before_containing: list[str] | None = None,
    remove_page_break_after_containing: list[str] | None = None,
    no_blank_removal: bool = False,
    prefix_cover_title: bool = False,
) -> dict[str, object]:
    if not input_docx.exists():
        raise FileNotFoundError(input_docx)
    tail_plan: dict[str, object] | None = None
    if tail_plan_json is not None:
        with tail_plan_json.open("r", encoding="utf-8-sig") as handle:
            loaded = json.load(handle)
        if not isinstance(loaded, dict):
            raise RuntimeError("tail content plan root must be an object")
        tail_plan = loaded
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    temp = output_docx.with_suffix(output_docx.suffix + ".tmp")
    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        changed_parts: list[str] = []
        patch_report: dict[str, object] = {}
        for item in zin.infolist():
            payload = zin.read(item.filename)
            if item.filename == "word/document.xml":
                payload, patch_report = patch_document_xml(
                    payload,
                    empty_paragraphs_only=empty_paragraphs_only,
                    tail_plan=tail_plan,
                    keep_lines_containing=keep_lines_containing or [],
                    page_break_before_containing=page_break_before_containing or [],
                    remove_page_break_after_containing=remove_page_break_after_containing or [],
                    no_blank_removal=no_blank_removal,
                    prefix_cover_title=prefix_cover_title,
                )
                changed_parts.append(item.filename)
            zout.writestr(item, payload)
    shutil.move(str(temp), str(output_docx))
    return {
        "schema": "graduation-project-builder.docx-submission-blocker-repair.v1",
        "generator": "scripts/repair_docx_submission_blockers.py",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source_docx_path": str(input_docx),
        "source_docx_sha256": sha256(input_docx),
        "final_docx_path": str(output_docx),
        "final_docx_sha256": sha256(output_docx),
        "repair_mode": "empty-paragraphs-only" if empty_paragraphs_only else "default",
        "tail_plan_json": str(tail_plan_json) if tail_plan_json is not None else None,
        "tail_plan_sha256": sha256(tail_plan_json) if tail_plan_json is not None else None,
        "no_blank_removal": no_blank_removal,
        "prefix_cover_title": prefix_cover_title,
        "changed_parts": changed_parts,
        **patch_report,
        "verdict": "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument(
        "--empty-paragraphs-only",
        action="store_true",
        help="Only remove pre-submission blocking blank paragraphs; do not run TOC/table placeholder repairs.",
    )
    parser.add_argument(
        "--tail-plan-json",
        help="Optional controlled plan for acknowledgement, extra references, body citation append, and literal placeholder replacements.",
    )
    parser.add_argument(
        "--keep-lines-containing",
        action="append",
        default=[],
        help="Add w:keepLines to body paragraphs containing this exact text; repeat for multiple anchors.",
    )
    parser.add_argument(
        "--page-break-before-containing",
        action="append",
        default=[],
        help="Add w:pageBreakBefore to paragraphs containing this exact text; repeat for multiple anchors.",
    )
    parser.add_argument(
        "--remove-page-break-after-containing",
        action="append",
        default=[],
        help="Remove an immediately following empty manual page-break paragraph after a paragraph containing this exact text.",
    )
    parser.add_argument(
        "--no-blank-removal",
        action="store_true",
        help="Do not remove blank paragraphs; use when applying only explicit anchor repairs.",
    )
    parser.add_argument(
        "--prefix-cover-title-label",
        action="store_true",
        help="Prefix the first cover title paragraph with the required label `\u9898\u76ee\uff1a` when missing.",
    )
    args = parser.parse_args()
    report = repair_docx(
        Path(args.input_docx).resolve(),
        Path(args.output_docx).resolve(),
        empty_paragraphs_only=args.empty_paragraphs_only,
        tail_plan_json=Path(args.tail_plan_json).resolve() if args.tail_plan_json else None,
        keep_lines_containing=args.keep_lines_containing,
        page_break_before_containing=args.page_break_before_containing,
        remove_page_break_after_containing=args.remove_page_break_after_containing,
        no_blank_removal=args.no_blank_removal,
        prefix_cover_title=args.prefix_cover_title_label,
    )
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
