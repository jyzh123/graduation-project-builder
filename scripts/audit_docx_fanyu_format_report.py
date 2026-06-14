#!/usr/bin/env python3
"""Audit a DOCX against the recurring Fanyu format-report issue families."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "r": R_NS}
W = f"{{{W_NS}}}"

ZH_ABSTRACT = "\u6458\u8981"
EN_ABSTRACT = "Abstract"
TOC_TITLE = "\u76ee\u5f55"
REFERENCES_TITLE = "\u53c2\u8003\u6587\u732e"
ACK_TITLE = "\u81f4\u8c22"
CONCLUSION_TITLE = "\u7ed3\u8bba"
SONGTI = "\u5b8b\u4f53"
HEITI = "\u9ed1\u4f53"
KAITI = "\u6977\u4f53"
HEADER_LEFT = "\u6c88\u9633\u79d1\u6280\u5b66\u9662\u5b66\u58eb\u5b66\u4f4d\u8bba\u6587"
FOOTER_SMALL_FIVE = "18"


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def local_name(node: ET.Element) -> str:
    return node.tag.rsplit("}", 1)[-1] if "}" in node.tag else node.tag


def text_of(element: ET.Element) -> str:
    chunks: list[str] = []
    def walk(node: ET.Element, in_run: bool = False) -> None:
        current_in_run = in_run or node.tag == qn("r")
        if node.tag == qn("t"):
            chunks.append(node.text or "")
        elif node.tag == qn("tab") and in_run:
            chunks.append("\t")
        for child in list(node):
            walk(child, current_in_run)

    walk(element)
    return "".join(chunks)


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def spacing(paragraph: ET.Element) -> dict[str, str]:
    node = paragraph.find("./w:pPr/w:spacing", NS)
    if node is None:
        return {}
    return {key: node.get(qn(key), "") for key in ("before", "after", "beforeLines", "afterLines", "line", "lineRule")}


def jc(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:jc", NS)
    return node.get(qn("val"), "") if node is not None else ""


def run_fonts(paragraph: ET.Element) -> list[dict[str, str]]:
    rows = []
    for run in paragraph.findall(".//w:r", NS):
        rpr = run.find("./w:rPr", NS)
        fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
        size = rpr.find("./w:sz", NS) if rpr is not None else None
        rows.append(
            {
                "eastAsia": fonts.get(qn("eastAsia"), "") if fonts is not None else "",
                "ascii": fonts.get(qn("ascii"), "") if fonts is not None else "",
                "hAnsi": fonts.get(qn("hAnsi"), "") if fonts is not None else "",
                "size": size.get(qn("val"), "") if size is not None else "",
            }
        )
    return rows


def visible_runs(paragraph: ET.Element) -> list[ET.Element]:
    rows: list[ET.Element] = []
    for run in paragraph.findall(".//w:r", NS):
        if text_of(run).strip():
            rows.append(run)
    return rows


def run_font_row(run: ET.Element) -> dict[str, str]:
    rpr = run.find("./w:rPr", NS)
    fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
    size = rpr.find("./w:sz", NS) if rpr is not None else None
    return {
        "eastAsia": fonts.get(qn("eastAsia"), "") if fonts is not None else "",
        "ascii": fonts.get(qn("ascii"), "") if fonts is not None else "",
        "hAnsi": fonts.get(qn("hAnsi"), "") if fonts is not None else "",
        "size": size.get(qn("val"), "") if size is not None else "",
    }


def all_visible_runs_match(paragraph: ET.Element, *, east_asia: str | None = None, size: str | None = None) -> bool:
    runs = visible_runs(paragraph)
    if not runs:
        return True
    for run in runs:
        row = run_font_row(run)
        if east_asia is not None and row.get("eastAsia") != east_asia:
            return False
        if size is not None and row.get("size") != size:
            return False
    return True


def load_docx(docx: Path) -> tuple[dict[str, bytes], ET.Element]:
    with zipfile.ZipFile(docx, "r") as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    return parts, ET.fromstring(parts["word/document.xml"])


def body_children(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no body")
    return list(body)


def direct_paragraphs(root: ET.Element) -> list[ET.Element]:
    return [node for node in body_children(root) if node.tag == qn("p")]


def is_toc_sdt(node: ET.Element) -> bool:
    if node.tag != qn("sdt"):
        return False
    value = compact(text_of(node))
    if value.startswith(TOC_TITLE):
        return True
    instr = "".join(item.text or "" for item in node.findall(".//w:instrText", NS))
    return bool(re.search(r"(^|\s)TOC(\s|$)", instr, flags=re.IGNORECASE))


def is_blank_paragraph(node: ET.Element) -> bool:
    return node.tag == qn("p") and not compact(text_of(node))


def has_page_break(node: ET.Element) -> bool:
    return node.tag == qn("p") and node.find(".//w:br[@w:type='page']", NS) is not None


def toc_has_blank_before_title(toc_node: ET.Element) -> bool:
    content = toc_node.find(".//w:sdtContent", NS)
    if content is None:
        return False
    children = list(content)
    for index, child_node in enumerate(children):
        if child_node.tag != qn("p") or not compact(text_of(child_node)).startswith(TOC_TITLE):
            continue
        previous = children[index - 1] if index else None
        return previous is not None and is_blank_paragraph(previous) and not has_page_break(previous)
    return False


def toc_entry_level(text: str) -> int | None:
    value = text.strip()
    if not value or compact(value).startswith(TOC_TITLE):
        return None
    if re.match(r"^(摘要|Abstract|参考文献|致谢)(\s|\t|\d|[IVX])", value):
        return 1
    if re.match(r"^第[一二三四五六七八九十0-9]+章", value):
        return 1
    if re.match(r"^[1-9]\d*\.[1-9]\d*", value):
        return 2
    return None


def load_ledger(issue_ledger: Path | None) -> dict[str, object]:
    if not issue_ledger:
        return {}
    return json.loads(issue_ledger.read_text(encoding="utf-8"))


def comment_issues_from_ledger(ledger: dict[str, object]) -> list[dict[str, object]]:
    payload = ledger.get("comment_docx")
    if not isinstance(payload, dict):
        return []
    rows = payload.get("issues")
    return rows if isinstance(rows, list) else []


def stats_issues_from_ledger(ledger: dict[str, object]) -> list[dict[str, object]]:
    rows = ledger.get("stats_issues")
    return rows if isinstance(rows, list) else []


def supported_comment_issue(row: dict[str, object]) -> bool:
    surface = str(row.get("surface") or "")
    name = str(row.get("name") or "")
    if surface.startswith("目录") and name in {"标题空行问题", "字体问题", "字号问题"}:
        return True
    if surface.startswith("页面设置") and name in {"页码样式问题", "空白页问题", "页眉内容", "字号问题"}:
        return True
    if surface.startswith("正文") and name in {"段前间距", "段后间距", "字体问题", "字号问题", "标点符号问题"}:
        return True
    if surface in {"参考文献标题", "致谢标题", "正文结论标题"} and name in {"字体问题", "字号问题"}:
        return True
    return False


def wants_comment_issue(comment_issues: list[dict[str, object]], *, surface_prefix: str, names: set[str]) -> bool:
    for row in comment_issues:
        surface = str(row.get("surface") or "")
        name = str(row.get("name") or "")
        if surface.startswith(surface_prefix) and name in names and row.get("blocking", True):
            return True
    return False


def wants_stats_issue(stats_issues: list[dict[str, object]], *, surfaces: set[str], names: set[str]) -> bool:
    for row in stats_issues:
        surface = str(row.get("surface") or "")
        name = str(row.get("name") or "")
        if surface in surfaces and name in names and row.get("blocking", True):
            return True
    return False


def wants_front_matter_page_number_style(stats_issues: list[dict[str, object]]) -> bool:
    return wants_stats_issue(
        stats_issues,
        surfaces={"front_matter_footer_page_number"},
        names={"页码样式问题", "页码样式错误"},
    )


def header_part_names(parts: dict[str, bytes]) -> list[str]:
    return sorted(name for name in parts if re.match(r"word/header\d+\.xml$", name))


def footer_part_names(parts: dict[str, bytes]) -> list[str]:
    return sorted(name for name in parts if re.match(r"word/footer\d+\.xml$", name))


def body_heading_titles(root: ET.Element) -> list[str]:
    titles: list[str] = []
    for paragraph in direct_paragraphs(root):
        value = text_of(paragraph).strip()
        if is_header_right_title(value):
            titles.append(value)
    return titles


def is_header_right_title(value: str) -> bool:
    stripped = value.strip()
    return bool(re.match(r"^第[一二三四五六七八九十0-9]+章\s+\S+", stripped)) or compact(stripped) in {
        REFERENCES_TITLE,
        ACK_TITLE,
        CONCLUSION_TITLE,
    }


def header_right_text(header_text: str) -> str:
    value = header_text.strip("\t ")
    if not value.startswith(HEADER_LEFT):
        return ""
    return value[len(HEADER_LEFT):].strip()


def valid_header_text(header_text: str, allowed_right_titles: set[str]) -> bool:
    header_text = header_text.strip("\t ")
    if not header_text.strip():
        return True
    if not header_text.startswith(HEADER_LEFT):
        return False
    right_text = header_right_text(header_text)
    if not right_text:
        return False
    if right_text in allowed_right_titles:
        return True
    return is_header_right_title(right_text)


def normalize_word_target(target: str) -> str:
    target = (target or "").replace("\\", "/")
    if not target:
        return ""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join("word", target))


def document_header_targets(parts: dict[str, bytes]) -> dict[str, str]:
    rels_payload = parts.get("word/_rels/document.xml.rels")
    if not rels_payload:
        return {}
    rels_root = ET.fromstring(rels_payload)
    rows: dict[str, str] = {}
    for rel in rels_root:
        if not str(rel.get("Type") or "").endswith("/header"):
            continue
        rid = str(rel.get("Id") or "")
        target = normalize_word_target(str(rel.get("Target") or ""))
        if rid and target:
            rows[rid] = target
    return rows


def sect_header_parts(sect_pr: ET.Element, rel_targets: dict[str, str]) -> list[str]:
    parts: list[str] = []
    for ref in sect_pr.findall("./w:headerReference", NS):
        rid = ref.get(f"{{{R_NS}}}id") or ""
        part = rel_targets.get(rid)
        if part and part not in parts:
            parts.append(part)
    return parts


def section_header_expectations(root: ET.Element, parts: dict[str, bytes]) -> tuple[dict[str, set[str]], list[dict[str, str]]]:
    rel_targets = document_header_targets(parts)
    expectations: dict[str, set[str]] = {}
    rows: list[dict[str, str]] = []
    body = root.find("./w:body", NS)
    if body is None:
        return expectations, rows
    current_section_titles: list[str] = []

    def record_section(sect_pr: ET.Element) -> None:
        nonlocal current_section_titles
        if not current_section_titles:
            return
        for part in sect_header_parts(sect_pr, rel_targets):
            expectations.setdefault(part, set()).update(current_section_titles)
            for title in current_section_titles:
                rows.append({"header_part": part, "expected_right_title": title})
        current_section_titles = []

    for node in list(body):
        if node.tag == qn("p"):
            value = text_of(node).strip()
            if is_header_right_title(value):
                current_section_titles.append(value)
            sect_pr = node.find("./w:pPr/w:sectPr", NS)
            if sect_pr is not None:
                record_section(sect_pr)
        elif node.tag == qn("sectPr"):
            record_section(node)
    return expectations, rows


def header_uses_report_owned_separator(header_root: ET.Element) -> bool:
    text = text_of(header_root)
    if not text.strip():
        return True
    if "\t" in text:
        return True
    return bool(re.search(re.escape(HEADER_LEFT) + r"\s{10,}\S", text))


def header_right_texts(header_root: ET.Element) -> set[str]:
    rows: set[str] = set()
    for paragraph in header_root.findall(".//w:p", NS):
        right = header_right_text(text_of(paragraph))
        if right:
            rows.add(right)
    return rows


def direct_run_font_size(run: ET.Element) -> str:
    rpr = run.find("./w:rPr", NS)
    size = rpr.find("./w:sz", NS) if rpr is not None else None
    return size.get(qn("val"), "") if size is not None else ""


def direct_run_font_size_cs(run: ET.Element) -> str:
    rpr = run.find("./w:rPr", NS)
    size = rpr.find("./w:szCs", NS) if rpr is not None else None
    return size.get(qn("val"), "") if size is not None else ""


def footer_page_number_runs(footer_root: ET.Element) -> list[ET.Element]:
    runs: list[ET.Element] = []
    for run in footer_root.findall(".//w:r", NS):
        has_field = run.find(".//w:fldChar", NS) is not None or run.find(".//w:instrText", NS) is not None
        value = text_of(run).strip()
        if has_field or re.fullmatch(r"[IVXLCDMivxlcdm\d]+", value or ""):
            runs.append(run)
    return runs


def has_page_setup_comment(comment_issues: list[dict[str, object]], name: str) -> bool:
    return wants_comment_issue(comment_issues, surface_prefix="页面设置", names={name})


def find_paragraph_by_text(root: ET.Element, target: str) -> ET.Element | None:
    compact_target = compact(target)
    if not compact_target:
        return None
    for paragraph in root.findall(".//w:p", NS):
        value = text_of(paragraph)
        if compact_target in compact(value):
            return paragraph
    return None


def find_index(root: ET.Element, predicate) -> int | None:
    for index, node in enumerate(body_children(root)):
        if predicate(node):
            return index
    return None


def is_heading(text: str) -> bool:
    return bool(re.match(r"^第[一二三四五六七八九十0-9]+章\s+", text.strip()) or re.match(r"^[1-9]\d*\.[1-9]\d*\s+", text.strip()))


def is_caption_text(text: str) -> bool:
    stripped = text.strip()
    if "\u7528\u4e8e\u8bf4\u660e" in stripped:
        return False
    return bool(
        re.match(r"^\u56fe\s*\d+[-.]\d+(?![\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u56fe\s*\d+(?![-.\d\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u8868\s*\d+[-.]\d+(?![\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u8868\s*\d+(?![-.\d\u4e2d\u7ed9\u7684])", stripped)
        or re.match(r"^\u7eed\u8868\s*\d", stripped)
    )


def first_font_value(fonts: list[dict[str, str]], key: str) -> str:
    for row in fonts:
        if row.get(key):
            return row[key]
    return ""


def rendered_blank_page_report(pdf: Path | None) -> dict[str, object]:
    if pdf is None:
        return {"provided": False, "verdict": "not-provided", "blank_page_count": None, "blank_pages": []}
    try:
        import fitz  # type: ignore
    except Exception as exc:  # pragma: no cover - environment-specific fallback
        return {"provided": True, "verdict": "blocked", "error": f"PyMuPDF unavailable: {exc}", "blank_page_count": None, "blank_pages": []}
    if not pdf.exists():
        return {"provided": True, "verdict": "blocked", "error": f"PDF not found: {pdf}", "blank_page_count": None, "blank_pages": []}
    blank_pages: list[dict[str, object]] = []
    with fitz.open(str(pdf)) as document:
        for index, page in enumerate(document, start=1):
            text = page.get_text("text") or ""
            content_lines = []
            for line in text.splitlines():
                value = compact(line)
                if not value:
                    continue
                if value.startswith(compact(HEADER_LEFT)):
                    continue
                if re.fullmatch(r"[IVXLCDMivxlcdm0-9]+", value):
                    continue
                content_lines.append(value)
            normalized = "".join(content_lines)
            drawings = page.get_drawings()
            images = page.get_images(full=True)
            if not normalized and not drawings and not images:
                blank_pages.append({"page": index, "reason": "no text/drawings/images"})
            elif not normalized and not images and len(drawings) <= 2:
                blank_pages.append({"page": index, "reason": "only header/footer-like residue"})
    return {
        "provided": True,
        "path": str(pdf),
        "verdict": "pass" if not blank_pages else "fail",
        "blank_page_count": len(blank_pages),
        "blank_pages": blank_pages,
    }


def audit(docx: Path, *, issue_ledger: Path | None = None, require_no_comments: bool = False, rendered_pdf: Path | None = None) -> dict[str, object]:
    parts, root = load_docx(docx)
    issues: list[str] = []
    ledger = load_ledger(issue_ledger)
    comment_issues = comment_issues_from_ledger(ledger)
    stats_issues = stats_issues_from_ledger(ledger)
    unhandled_comment_issues = [row for row in comment_issues if not supported_comment_issue(row)]
    if unhandled_comment_issues:
        issues.append(f"unhandled comment-ledger issue families: {len(unhandled_comment_issues)}")
    wants_header_content = has_page_setup_comment(comment_issues, "页眉内容") or wants_stats_issue(
        stats_issues, surfaces={"running_header_content"}, names={"页眉内容"}
    )
    wants_page_number_style = wants_front_matter_page_number_style(stats_issues)
    wants_footer_small_five = has_page_setup_comment(comment_issues, "字号问题") or wants_page_number_style
    wants_blank_page = has_page_setup_comment(comment_issues, "空白页问题")
    toc = find_index(root, is_toc_sdt)
    zh = find_index(root, lambda n: n.tag == qn("p") and compact(text_of(n)) == ZH_ABSTRACT)
    en = find_index(root, lambda n: n.tag == qn("p") and compact(text_of(n)) == compact(EN_ABSTRACT))
    body = find_index(root, lambda n: n.tag == qn("p") and re.match(r"^第[一二三四五六七八九十0-9]+章\s+", text_of(n).strip()))
    refs = find_index(root, lambda n: n.tag == qn("p") and compact(text_of(n)) == REFERENCES_TITLE)
    ack = find_index(root, lambda n: n.tag == qn("p") and compact(text_of(n)) == ACK_TITLE)
    order_ok = all(value is not None for value in (toc, zh, en, body, refs, ack)) and toc < zh < en < body < refs < ack  # type: ignore[operator]
    if not order_ok:
        issues.append(f"structure order failed: toc={toc}, zh={zh}, en={en}, body={body}, refs={refs}, ack={ack}")

    title_failures = []
    keyword_failures = []
    heading_failures = []
    caption_failures = []
    reference_space_failures = []
    tail_title_failures = []
    for paragraph in direct_paragraphs(root):
        value = text_of(paragraph).strip()
        cvalue = compact(value)
        if cvalue in {ZH_ABSTRACT, compact(EN_ABSTRACT), TOC_TITLE}:
            sp = spacing(paragraph)
            if sp.get("line") != "360" or sp.get("lineRule") != "auto":
                title_failures.append(value)
        if value.startswith("Key words") and jc(paragraph) != "left":
            keyword_failures.append(value)
        if is_heading(value):
            sp = spacing(paragraph)
            if sp.get("beforeLines") != "100" or sp.get("afterLines") != "100":
                heading_failures.append(value[:80])
        if is_caption_text(value):
            fonts = run_fonts(paragraph)
            sp = spacing(paragraph)
            if (
                first_font_value(fonts, "eastAsia") != HEITI
                or first_font_value(fonts, "size") != "21"
                or sp.get("beforeLines") != "50"
            ):
                caption_failures.append(value[:80])
        if re.match(r"^\[\d+\]\s+", value):
            reference_space_failures.append(value[:80])
        if cvalue in {REFERENCES_TITLE, ACK_TITLE, CONCLUSION_TITLE}:
            fonts = run_fonts(paragraph)
            sp = spacing(paragraph)
            if (
                jc(paragraph) != "center"
                or not all_visible_runs_match(paragraph, east_asia=HEITI, size="28")
                or sp.get("before") != "0"
                or sp.get("after") != "0"
                or sp.get("line") != "360"
            ):
                tail_title_failures.append(value)
    if title_failures:
        issues.append(f"front/TOC title line-spacing failures: {title_failures[:5]}")
    if keyword_failures:
        issues.append(f"keyword alignment failures: {len(keyword_failures)}")
    if heading_failures:
        issues.append(f"body heading spacing failures: {heading_failures[:5]}")
    if caption_failures:
        issues.append(f"caption format failures: {caption_failures[:5]}")
    if reference_space_failures:
        issues.append(f"reference label spacing failures: {len(reference_space_failures)}")
    if tail_title_failures:
        issues.append(f"tail/conclusion title failures: {tail_title_failures}")

    toc_line_failures = 0
    toc_typography_failures = []
    toc_title_blank_failure = False
    if toc is not None:
        children = body_children(root)
        previous = children[toc - 1] if toc > 0 else None
        toc_node = body_children(root)[toc]
        if not toc_has_blank_before_title(toc_node) and (
            previous is None or not is_blank_paragraph(previous) or previous.find("./w:pPr/w:sectPr", NS) is not None or has_page_break(previous)
        ):
            toc_title_blank_failure = True
        for paragraph in toc_node.findall(".//w:p", NS):
            text = text_of(paragraph)
            if compact(text):
                sp = spacing(paragraph)
                if sp.get("line") != "360" or sp.get("lineRule") != "auto":
                    toc_line_failures += 1
                level = toc_entry_level(text)
                if level == 1 and not all_visible_runs_match(paragraph, east_asia=KAITI, size="28"):
                    toc_typography_failures.append(text.strip()[:80])
                elif level == 2 and not all_visible_runs_match(paragraph, size="28"):
                    toc_typography_failures.append(text.strip()[:80])
    if toc_title_blank_failure:
        issues.append("TOC title blank-line-before failure")
    if toc_line_failures:
        issues.append(f"TOC paragraph line-spacing failures: {toc_line_failures}")
    if toc_typography_failures:
        issues.append(f"TOC entry typography failures: {toc_typography_failures[:8]}")

    targeted_body_failures = []
    punctuation_failures = []
    for row in comment_issues:
        if not str(row.get("surface") or "").startswith("正文文本内容"):
            continue
        targets = row.get("target_texts") if isinstance(row.get("target_texts"), list) else []
        if not targets:
            continue
        paragraph = find_paragraph_by_text(root, str(targets[0]))
        name = str(row.get("name") or "")
        if paragraph is None:
            if name == "标点符号问题":
                whole_text = text_of(root)
                if re.search(r"(?<=[A-Za-z0-9]),(?=[A-Za-z])", whole_text):
                    punctuation_failures.append(f"missing target for comment {row.get('comment_id')}")
            else:
                targeted_body_failures.append(f"missing target for comment {row.get('comment_id')}")
            continue
        if name in {"字体问题", "字号问题"} and not all_visible_runs_match(paragraph, east_asia=SONGTI, size="24"):
            targeted_body_failures.append(text_of(paragraph)[:80])
        if name == "标点符号问题" and re.search(r"(?<=[A-Za-z0-9]),(?=[A-Za-z])", text_of(paragraph)):
            punctuation_failures.append(text_of(paragraph)[:80])
    if targeted_body_failures:
        issues.append(f"targeted body text font/size failures: {targeted_body_failures[:5]}")
    if punctuation_failures:
        issues.append(f"targeted punctuation failures: {punctuation_failures[:5]}")

    table_cell_failures = 0
    table_cell_count = 0
    for paragraph in root.findall(".//w:tbl//w:tc//w:p", NS):
        table_cell_count += 1
        sp = spacing(paragraph)
        if jc(paragraph) != "center" or sp.get("lineRule") != "atLeast":
            table_cell_failures += 1
    if table_cell_failures:
        issues.append(f"table cell alignment/lineRule failures: {table_cell_failures}")

    header_failures = []
    header_content_failures = []
    header_separator_failures = []
    header_section_mismatch_failures = []
    footer_failures = []
    footer_field_failures = []
    header_count = 0
    footer_count = 0
    footer_checked_runs = 0
    allowed_header_right_titles = set(body_heading_titles(root)) | {REFERENCES_TITLE, ACK_TITLE, CONCLUSION_TITLE}
    header_expectations, header_expectation_rows = section_header_expectations(root, parts)
    for name, payload in parts.items():
        if re.match(r"word/header\d+\.xml$", name):
            header_count += 1
            hroot = ET.fromstring(payload)
            header_text = text_of(hroot)
            if header_text.strip() and not valid_header_text(header_text, allowed_header_right_titles):
                header_content_failures.append(name)
            if wants_header_content and not header_uses_report_owned_separator(hroot):
                header_separator_failures.append(name)
            expected_rights = header_expectations.get(name)
            if wants_header_content and expected_rights:
                actual_rights = header_right_texts(hroot)
                if not actual_rights or not actual_rights.issubset(expected_rights):
                    header_section_mismatch_failures.append(f"{name}: expected={sorted(expected_rights)}, actual={sorted(actual_rights)}")
            bottom = hroot.find(".//w:pPr/w:pBdr/w:bottom", NS)
            if bottom is None or bottom.get(qn("val")) != "single":
                header_failures.append(name)
        if re.match(r"word/footer\d+\.xml$", name):
            footer_count += 1
            froot = ET.fromstring(payload)
            if "PAGE" not in "".join(item.text or "" for item in froot.findall(".//w:instrText", NS)):
                footer_field_failures.append(name)
            runs_to_check = footer_page_number_runs(froot) if wants_footer_small_five else froot.findall(".//w:r", NS)
            for run in runs_to_check:
                footer_checked_runs += 1
                rpr = run.find("./w:rPr", NS)
                rf = rpr.find("./w:rFonts", NS) if rpr is not None else None
                expected_size = FOOTER_SMALL_FIVE if wants_footer_small_five else "21"
                if (
                    rf is None
                    or rf.get(qn("eastAsia")) != SONGTI
                    or direct_run_font_size(run) != expected_size
                    or direct_run_font_size_cs(run) != expected_size
                ):
                    footer_failures.append(name)
                    break
    if wants_header_content and not header_count:
        issues.append("report-ledger header content issue unresolved: no header parts found")
    if wants_footer_small_five and not footer_count:
        issues.append("report-ledger footer page-number size issue unresolved: no footer parts found")
    if header_failures:
        issues.append(f"header bottom-border failures: {header_failures}")
    if header_content_failures:
        issues.append(f"header content failures: {header_content_failures}")
    if header_separator_failures:
        issues.append(f"report-ledger header left/right separator failures: {header_separator_failures}")
    if header_section_mismatch_failures:
        issues.append(f"report-ledger header section-title mismatches: {header_section_mismatch_failures[:8]}")
    if footer_failures:
        issues.append(f"footer page-number font failures: {footer_failures}")
    if footer_field_failures:
        issues.append(f"footer PAGE field failures: {footer_field_failures}")

    comment_parts = [name for name in parts if name in {"word/comments.xml", "word/commentsExtended.xml", "word/people.xml"}]
    comment_refs = len(root.findall(".//w:commentReference", NS)) + len(root.findall(".//w:commentRangeStart", NS)) + len(root.findall(".//w:commentRangeEnd", NS))
    if require_no_comments and (comment_parts or comment_refs):
        issues.append(f"report comments remain: parts={comment_parts}, refs={comment_refs}")

    frontmatter_pg = []
    for sect in root.findall(".//w:sectPr", NS):
        pg = sect.find("./w:pgNumType", NS)
        if pg is not None:
            frontmatter_pg.append({"fmt": pg.get(qn("fmt"), ""), "start": pg.get(qn("start"), "")})
    if not any(row.get("fmt") == "upperRoman" and row.get("start") == "1" for row in frontmatter_pg):
        issues.append("front-matter upperRoman page numbering was not found")

    structural_blank_page_risks = 0
    children = body_children(root)
    for left, right in zip(children, children[1:]):
        if has_page_break(left) and has_page_break(right):
            structural_blank_page_risks += 1
    if structural_blank_page_risks:
        issues.append(f"structural consecutive page-break blank-page risks: {structural_blank_page_risks}")
    rendered_blank_pages = rendered_blank_page_report(rendered_pdf) if wants_blank_page else {"provided": False, "verdict": "not-requested", "blank_page_count": None, "blank_pages": []}
    blank_page_report_closed = not structural_blank_page_risks and rendered_blank_pages.get("verdict") == "pass"
    if wants_blank_page and rendered_blank_pages.get("verdict") == "not-provided":
        issues.append("report-ledger blank-page issue requires rendered blank-page evidence; XML-only audit cannot close it")
    elif wants_blank_page and rendered_blank_pages.get("verdict") == "blocked":
        issues.append(f"report-ledger blank-page rendered audit blocked: {rendered_blank_pages.get('error')}")
    elif wants_blank_page and rendered_blank_pages.get("verdict") == "fail":
        issues.append(f"rendered blank-page failures: {rendered_blank_pages.get('blank_pages')}")

    ledger_summary = None
    if issue_ledger:
        data = ledger
        ledger_summary = {
            "schema": data.get("schema"),
            "issue_count": data.get("issue_count"),
            "comment_issue_count": data.get("comment_issue_count"),
            "official_comment_error_count": (((data.get("comment_docx") or {}).get("summary") or {}).get("official_error_count") if isinstance(data.get("comment_docx"), dict) else None),
            "stats_issue_count": len(data.get("stats_issues") or []),
            "format_error_count": (data.get("overview") or {}).get("format_error_count"),
        }
    report_issue_closure = {
        "header_content_requested": wants_header_content,
        "header_content_verdict": "pass" if wants_header_content and header_count and not header_content_failures and not header_separator_failures and not header_section_mismatch_failures else ("not-requested" if not wants_header_content else "fail"),
        "footer_small_five_requested": wants_footer_small_five,
        "footer_small_five_expected_half_points": FOOTER_SMALL_FIVE if wants_footer_small_five else "",
        "footer_small_five_verdict": "pass" if wants_footer_small_five and footer_count and not footer_failures and not footer_field_failures else ("not-requested" if not wants_footer_small_five else "fail"),
        "front_matter_page_number_style_requested": wants_page_number_style,
        "front_matter_page_number_style_verdict": "pass" if wants_page_number_style and any(row.get("fmt") == "upperRoman" and row.get("start") == "1" for row in frontmatter_pg) else ("not-requested" if not wants_page_number_style else "fail"),
        "blank_page_requested": wants_blank_page,
        "blank_page_verdict": "pass" if wants_blank_page and blank_page_report_closed else ("not-requested" if not wants_blank_page else "fail"),
    }
    return {
        "schema": "graduation-project-builder.docx-fanyu-format-report-audit.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "generator": "scripts/audit_docx_fanyu_format_report.py",
        "docx_path": str(docx),
        "docx_sha256": sha256_file(docx),
        "issue_ledger": str(issue_ledger) if issue_ledger else "",
        "ledger_summary": ledger_summary,
        "surface_indices": {"toc": toc, "zh_abstract": zh, "en_abstract": en, "first_body": body, "references": refs, "acknowledgement": ack},
        "counts": {
            "table_cell_paragraph_count": table_cell_count,
            "table_cell_failures": table_cell_failures,
            "toc_line_failures": toc_line_failures,
            "toc_typography_failures": len(toc_typography_failures),
            "unhandled_comment_issue_count": len(unhandled_comment_issues),
            "targeted_body_failures": len(targeted_body_failures),
            "punctuation_failures": len(punctuation_failures),
            "structural_blank_page_risks": structural_blank_page_risks,
            "header_part_count": header_count,
            "footer_part_count": footer_count,
            "footer_page_number_run_count": footer_checked_runs,
            "comment_part_count": len(comment_parts),
            "comment_anchor_count": comment_refs,
        },
        "report_issue_closure": report_issue_closure,
        "header_section_expectations": header_expectation_rows,
        "rendered_blank_page_report": rendered_blank_pages,
        "unhandled_comment_issues": unhandled_comment_issues[:20],
        "issues": issues,
        "verdict": "pass" if not issues else "fail",
        "passed": not issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--issue-ledger")
    parser.add_argument("--rendered-pdf")
    parser.add_argument("--report-json", required=True)
    parser.add_argument("--require-no-comments", action="store_true")
    args = parser.parse_args()
    report = audit(
        Path(args.docx).resolve(),
        issue_ledger=Path(args.issue_ledger).resolve() if args.issue_ledger else None,
        require_no_comments=args.require_no_comments,
        rendered_pdf=Path(args.rendered_pdf).resolve() if args.rendered_pdf else None,
    )
    report_path = Path(args.report_json).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"verdict": report["verdict"], "issue_count": len(report["issues"])}, ensure_ascii=True))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
