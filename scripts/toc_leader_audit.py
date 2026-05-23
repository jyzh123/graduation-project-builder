"""Audit TOC entries for real right-tab dotted leaders in DOCX XML."""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

ACCEPTED_DOTTED_LEADERS = {"dot", "middleDot", "heavy"}
TAIL_BLOCK_LABELS = {
    "references": "\u53c2\u8003\u6587\u732e",
    "acknowledgement": "\u81f4\u8c22",
}


@dataclass(frozen=True)
class StyleInfo:
    name: str
    based_on: str
    ppr: ET.Element | None


@dataclass(frozen=True)
class TocEntryAudit:
    paragraph_index: int
    text: str
    tab_count: int
    page_number: str
    right_tab_pos: str
    right_tab_positions: tuple[str, ...]
    right_tab_leader: str
    direct_tab_count: int
    inherited_tab_count: int
    issues: tuple[str, ...]

    @property
    def passed(self) -> bool:
        return not self.issues


def w_attr(element: ET.Element | None, name: str) -> str:
    if element is None:
        return ""
    return element.attrib.get(W + name, "")


def compact_text(value: str) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", value or "").strip().lower()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return w_attr(style, "val")


def paragraph_numbering_level(paragraph: ET.Element) -> int | None:
    level = paragraph.find("w:pPr/w:numPr/w:ilvl", NS)
    value = w_attr(level, "val")
    try:
        return int(value) if value else None
    except ValueError:
        return None


def trailing_page_number_from_text(text: str) -> str:
    match = re.search(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", text or "")
    return match.group(1) if match else ""


def toc_visible_label(text: str) -> str:
    value = text or ""
    value = re.sub(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", "", value).strip()
    value = re.sub(r"[\t\u2026.·•]+", " ", value)
    return value.strip()


def normalized_label(text: str) -> str:
    return compact_text(toc_visible_label(text))


def looks_like_body_level1_heading(text: str) -> bool:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    if not stripped:
        return False
    if re.match(r"^\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u7ae0", stripped):
        return True
    return bool(re.match(r"^[1-9]\d*(?:[\s\u3000]+|$)", stripped))


def is_body_level1_heading(paragraph: ET.Element) -> bool:
    level = paragraph_numbering_level(paragraph)
    if level == 0 and paragraph_text(paragraph):
        return True
    return looks_like_body_level1_heading(paragraph_text(paragraph))


def load_styles(zf: zipfile.ZipFile) -> dict[str, StyleInfo]:
    try:
        root = ET.fromstring(zf.read("word/styles.xml"))
    except (KeyError, ET.ParseError):
        return {}
    styles: dict[str, StyleInfo] = {}
    for style in root.findall("w:style", NS):
        style_id = w_attr(style, "styleId")
        if not style_id:
            continue
        name_el = style.find("w:name", NS)
        based_on_el = style.find("w:basedOn", NS)
        styles[style_id] = StyleInfo(
            name=w_attr(name_el, "val") or style_id,
            based_on=w_attr(based_on_el, "val"),
            ppr=style.find("w:pPr", NS),
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
    return chain


def tabs_from_ppr(ppr: ET.Element | None) -> list[ET.Element]:
    if ppr is None:
        return []
    return list(ppr.findall("w:tabs/w:tab", NS))


def direct_tabs(paragraph: ET.Element) -> list[ET.Element]:
    return tabs_from_ppr(paragraph.find("w:pPr", NS))


def inherited_tabs(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> list[ET.Element]:
    for info in style_chain(styles, paragraph_style_id(paragraph)):
        tabs = tabs_from_ppr(info.ppr)
        if tabs:
            return tabs
    return []


def effective_tabs(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> tuple[list[ET.Element], int, int]:
    direct = direct_tabs(paragraph)
    inherited = inherited_tabs(paragraph, styles)
    if direct:
        return direct, len(direct), len(inherited)
    return inherited, 0, len(inherited)


def choose_right_tab(tabs: list[ET.Element]) -> ET.Element | None:
    right_tabs = [tab for tab in tabs if (w_attr(tab, "val") or "left") == "right"]
    if right_tabs:
        return max(right_tabs, key=lambda tab: int(w_attr(tab, "pos") or "0"))
    usable_tabs = [tab for tab in tabs if w_attr(tab, "val") != "clear"]
    if usable_tabs:
        return max(usable_tabs, key=lambda tab: int(w_attr(tab, "pos") or "0"))
    return None


def run_tab_segments(paragraph: ET.Element) -> tuple[int, list[str]]:
    segments = [""]
    tab_count = 0
    for node in paragraph.iter():
        if node.tag == W + "tab":
            tab_count += 1
            segments.append("")
        elif node.tag == W + "t":
            for index, part in enumerate((node.text or "").split("\t")):
                if index:
                    tab_count += 1
                    segments.append("")
                segments[-1] += part
    return tab_count, segments


def page_number_after_last_tab(paragraph: ET.Element) -> str:
    _tab_count, segments = run_tab_segments(paragraph)
    tail = segments[-1].strip() if segments else ""
    return trailing_page_number_from_text(tail)


def looks_like_toc_entry_without_tab(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> bool:
    text = paragraph_text(paragraph)
    if not trailing_page_number_from_text(text):
        return False
    if re.search(r"(?:\.{3,}|\u2026{2,})\s*(?:\d+|[ivxlcdmIVXLCDM]+)\s*$", text or ""):
        return True
    style_id = paragraph_style_id(paragraph)
    style_name = compact_text(styles.get(style_id, StyleInfo(style_id, "", None)).name)
    if "toc" in style_name or style_id in {"11", "21", "31"}:
        return True
    label = re.sub(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", "", text or "").strip()
    if compact_text(toc_visible_label(text)) in {compact_text(label) for label in TAIL_BLOCK_LABELS.values()}:
        return True
    if compact_text(label) in {"\u6458\u8981", "abstract"}:
        return True
    return bool(
        re.match(r"^\u7b2c\d+\u7ae0\s+\S", label)
        or re.match(r"^\d{1,2}\.\d+(?:\.\d+)?\s+\S", label)
    )


def is_toc_title(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> bool:
    text = compact_text(paragraph_text(paragraph))
    style_id = paragraph_style_id(paragraph)
    style_name = compact_text(styles.get(style_id, StyleInfo(style_id, "", None)).name)
    return (
        text in {"\u76ee\u5f55", "contents", "tableofcontents"}
        or ("toc" in style_name and "title" in style_name)
    )


def body_paragraphs(root: ET.Element) -> list[ET.Element]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    return [child for child in list(body) if child.tag == W + "p"]


def body_children(root: ET.Element) -> list[ET.Element]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    return list(body)


def sdt_content(element: ET.Element) -> ET.Element | None:
    if element.tag != W + "sdt":
        return None
    return element.find("w:sdtContent", NS)


def direct_paragraphs(element: ET.Element) -> list[ET.Element]:
    return [child for child in list(element) if child.tag == W + "p"]


def child_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS)).strip()


def first_body_heading_index(children: list[ET.Element], start_index: int, styles: dict[str, StyleInfo]) -> int | None:
    for index in range(start_index, len(children)):
        child = children[index]
        if child.tag != W + "p":
            continue
        if is_toc_entry_candidate(child, styles):
            continue
        if is_body_level1_heading(child):
            return index
    return None


def is_toc_entry_candidate(paragraph: ET.Element, styles: dict[str, StyleInfo]) -> bool:
    tab_count, _segments = run_tab_segments(paragraph)
    page_number = page_number_after_last_tab(paragraph)
    return (tab_count and bool(page_number)) or (
        not tab_count and looks_like_toc_entry_without_tab(paragraph, styles)
    )


def collect_toc_entry_paragraphs(root: ET.Element, styles: dict[str, StyleInfo]) -> tuple[list[ET.Element], list[str]]:
    children = body_children(root)
    title_index: int | None = None
    title_paragraph: ET.Element | None = None
    toc_container: ET.Element | None = None
    for index, child in enumerate(children):
        if child.tag == W + "p" and is_toc_title(child, styles):
            title_index = index
            title_paragraph = child
            break
        content = sdt_content(child)
        if content is not None:
            for paragraph in direct_paragraphs(content):
                if is_toc_title(paragraph, styles):
                    title_index = index
                    title_paragraph = paragraph
                    toc_container = content
                    break
        if title_index is not None:
            break
    if title_index is None:
        return [], []

    entries: list[ET.Element] = []
    issues: list[str] = []
    if toc_container is not None:
        seen_title = False
        for paragraph in direct_paragraphs(toc_container):
            if paragraph is title_paragraph:
                seen_title = True
                continue
            if not seen_title:
                continue
            text = paragraph_text(paragraph)
            tab_count, _segments = run_tab_segments(paragraph)
            if is_toc_entry_candidate(paragraph, styles):
                entries.append(paragraph)
                continue
            if not text.strip():
                continue
            if tab_count and not page_number_after_last_tab(paragraph):
                issues.append(f"TOC candidate paragraph lacks page number after final tab: {text[:80]}")
                continue
            issues.append(f"TOC content-control range contains non-entry paragraph: {text[:80]}")
        if not entries:
            issues.append("TOC entry block has no tabbed entry paragraphs after TOC title")
        return entries, issues

    body_start_index = first_body_heading_index(children, title_index + 1, styles)
    scan_end = body_start_index if body_start_index is not None else len(children)
    for child in children[title_index + 1 : scan_end]:
        if child.tag == W + "p":
            paragraph = child
            text = paragraph_text(paragraph)
            tab_count, _segments = run_tab_segments(paragraph)
            if is_toc_entry_candidate(paragraph, styles):
                entries.append(paragraph)
                continue
            if not text.strip():
                continue
            if tab_count and not page_number_after_last_tab(paragraph):
                issues.append(f"TOC candidate paragraph lacks page number after final tab: {text[:80]}")
                continue
            issues.append(f"TOC protected range contains non-entry paragraph before body start: {text[:80]}")
            continue
        if child.tag == W + "tbl":
            issues.append(f"TOC protected range contains table before body start: {child_text(child)[:80]}")
            continue
        if child_text(child).strip():
            issues.append(
                "TOC protected range contains non-paragraph/non-table body content before body start: "
                f"{child.tag.rsplit('}', 1)[-1]}"
            )
    if not entries:
        issues.append("TOC entry block has no tabbed entry paragraphs after TOC title")
    return entries, issues


def body_tail_blocks(root: ET.Element) -> dict[str, int]:
    result: dict[str, int] = {}
    for index, paragraph in enumerate(body_paragraphs(root)):
        normalized = compact_text(paragraph_text(paragraph))
        for key, label in TAIL_BLOCK_LABELS.items():
            if normalized == compact_text(label) and key not in result:
                result[key] = index
    return result


def audit_required_tail_entries(root: ET.Element, entries: list[ET.Element]) -> tuple[dict[str, object], list[str]]:
    body_tail = body_tail_blocks(root)
    entry_positions: dict[str, int] = {}
    entry_labels = [normalized_label(paragraph_text(entry)) for entry in entries]
    for key, label in TAIL_BLOCK_LABELS.items():
        target = compact_text(label)
        for index, label_text in enumerate(entry_labels):
            if label_text == target:
                entry_positions[key] = index
                break
    issues: list[str] = []
    for key, label in TAIL_BLOCK_LABELS.items():
        if key in body_tail and key not in entry_positions:
            issues.append(f"TOC is missing required tail-block entry: {label}")
    body_order = sorted(body_tail, key=body_tail.get)
    toc_order = sorted((key for key in body_order if key in entry_positions), key=entry_positions.get)
    if len(toc_order) == len(body_order) and toc_order != body_order:
        expected = " before ".join(TAIL_BLOCK_LABELS[key] for key in body_order)
        found = " before ".join(TAIL_BLOCK_LABELS[key] for key in toc_order)
        issues.append(f"TOC tail-block order does not match body tail order: expected {expected}; found {found}")
    return (
        {
            "body_tail_blocks": body_tail,
            "toc_tail_entries": entry_positions,
            "body_tail_order": body_order,
            "toc_tail_order": toc_order,
            "required_tail_labels": TAIL_BLOCK_LABELS,
        },
        issues,
    )


def audit_docx_toc_dotted_leaders(docx_path: Path) -> tuple[dict[str, object], list[str]]:
    path = Path(docx_path)
    issues: list[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
            styles = load_styles(zf)
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return {}, [f"TOC dotted-leader DOCX audit could not read document XML: {exc}"]

    entries, collection_issues = collect_toc_entry_paragraphs(root, styles)
    tail_payload, tail_issues = audit_required_tail_entries(root, entries)
    audits: list[TocEntryAudit] = []
    paragraphs = root.findall(".//w:p", NS)
    page_column_positions: set[str] = set()
    for paragraph in entries:
        paragraph_index = paragraphs.index(paragraph) + 1 if paragraph in paragraphs else -1
        text = paragraph_text(paragraph)
        tab_count, _segments = run_tab_segments(paragraph)
        page_number = page_number_after_last_tab(paragraph)
        tabs, direct_count, inherited_count = effective_tabs(paragraph, styles)
        right_tabs = [tab for tab in tabs if (w_attr(tab, "val") or "left") == "right"]
        right_tab_positions = tuple(sorted({w_attr(tab, "pos") or "none" for tab in right_tabs}))
        right_tab = choose_right_tab(tabs)
        leader = w_attr(right_tab, "leader") or "none"
        pos = w_attr(right_tab, "pos") or "none"
        entry_issues: list[str] = []
        if tab_count < 1:
            entry_issues.append("missing visible w:tab run before page number")
        if not page_number:
            entry_issues.append("missing page-number text after final w:tab")
        if right_tab is None:
            entry_issues.append("missing right tab stop for page-number column")
        elif (w_attr(right_tab, "val") or "left") != "right":
            entry_issues.append(f"page-number tab stop is not right-aligned: val={w_attr(right_tab, 'val') or 'left'}")
        if leader not in ACCEPTED_DOTTED_LEADERS:
            entry_issues.append(f"right tab leader is not dotted: leader={leader}")
        if len(right_tab_positions) > 1:
            entry_issues.append(f"multiple effective right-tab page columns: {', '.join(right_tab_positions)}")
        if pos != "none":
            page_column_positions.add(pos)
        audits.append(
            TocEntryAudit(
                paragraph_index=paragraph_index,
                text=text,
                tab_count=tab_count,
                page_number=page_number,
                right_tab_pos=pos,
                right_tab_positions=right_tab_positions,
                right_tab_leader=leader,
                direct_tab_count=direct_count,
                inherited_tab_count=inherited_count,
                issues=tuple(entry_issues),
            )
        )

    if len(page_column_positions) > 1:
        issues.append(f"TOC page-number column positions are not unified: {', '.join(sorted(page_column_positions))}")
    for audit in audits:
        if audit.issues:
            issues.append(
                "TOC entry paragraph "
                f"{audit.paragraph_index} failed dotted leader audit ({'; '.join(audit.issues)}): "
                f"{audit.text[:100]}"
            )
    issues = collection_issues + tail_issues + issues
    payload: dict[str, object] = {
        "schema": "graduation-project-builder.toc-dotted-leader-audit.v1",
        "docx_path": str(path),
        "entry_count": len(audits),
        "passed": not issues,
        "accepted_leaders": sorted(ACCEPTED_DOTTED_LEADERS),
        "tail_block_coverage": tail_payload,
        "entries": [
            {
                "paragraph_index": audit.paragraph_index,
                "text": audit.text,
                "tab_count": audit.tab_count,
                "page_number": audit.page_number,
                "right_tab_pos": audit.right_tab_pos,
                "right_tab_positions": list(audit.right_tab_positions),
                "right_tab_leader": audit.right_tab_leader,
                "direct_tab_count": audit.direct_tab_count,
                "inherited_tab_count": audit.inherited_tab_count,
                "issues": list(audit.issues),
            }
            for audit in audits
        ],
        "issues": issues,
    }
    return payload, issues


def audit_summary(payload: dict[str, object]) -> str:
    entries = payload.get("entries")
    if not isinstance(entries, list):
        return "entry_count=0"
    pieces = []
    for entry in entries[:8]:
        if not isinstance(entry, dict):
            continue
        pieces.append(
            "p{paragraph_index}:tabs={tab_count},right={right_tab_pos},leader={right_tab_leader},page={page_number}".format(
                paragraph_index=entry.get("paragraph_index", "?"),
                tab_count=entry.get("tab_count", "?"),
                right_tab_pos=entry.get("right_tab_pos", "?"),
                right_tab_leader=entry.get("right_tab_leader", "?"),
                page_number=entry.get("page_number", "?"),
            )
        )
    suffix = f"; ... total={len(entries)}" if len(entries) > 8 else ""
    return f"entry_count={len(entries)} " + "; ".join(pieces) + suffix


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit DOCX TOC dotted leaders, content contamination, and required tail entries."
    )
    parser.add_argument("--docx", required=True)
    parser.add_argument("--report-json")
    args = parser.parse_args()
    payload, issues = audit_docx_toc_dotted_leaders(Path(args.docx).resolve())
    if args.report_json:
        output = Path(args.report_json).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"passed": not issues, "issue_count": len(issues), "summary": audit_summary(payload)}, ensure_ascii=False))
    return 0 if not issues else 1


if __name__ == "__main__":
    raise SystemExit(main())
