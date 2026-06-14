#!/usr/bin/env python3
"""Parse Fanyu-style thesis format-check HTML reports into a stable issue ledger."""

from __future__ import annotations

import argparse
import collections
import hashlib
import html
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def read_text(path: Path) -> str:
    payload = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            text = payload.decode(encoding)
        except UnicodeDecodeError:
            continue
        if "\ufffd" not in text:
            return text
    return payload.decode("utf-8", errors="replace")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def parse_parag_json(compare_html: Path) -> dict[str, Any]:
    text = read_text(compare_html)
    start_match = re.search(r"var\s+paragJson\s*=\s*\"", text)
    if not start_match:
        raise ValueError(f"paragJson not found in {compare_html}")
    start = start_match.end() - 1
    end = start + 1
    escaped = False
    while end < len(text):
        ch = text[end]
        if escaped:
            escaped = False
        elif ch == "\\":
            escaped = True
        elif ch == '"':
            break
        end += 1
    if end >= len(text):
        raise ValueError(f"paragJson string is not closed in {compare_html}")
    inner_json = json.loads(text[start : end + 1])
    data = json.loads(inner_json)
    if not isinstance(data, dict):
        raise ValueError(f"paragJson in {compare_html} is not an object")
    return data


def parse_overview(compare_html: Path, stats_html: Path | None = None) -> dict[str, Any]:
    text = read_text(compare_html)
    overview: dict[str, Any] = {}
    m = re.search(r"格式错误数：</span><span>(\d+)</span>", text)
    if m:
        overview["format_error_count"] = int(m.group(1))
    m = re.search(r'class="main_code_time">([^<]+)</span>', text)
    if m:
        overview["generated_time"] = html.unescape(m.group(1).strip())
    m = re.search(r'<div title="([^"]+)" class="temp_width">', text)
    if m:
        overview["template_name"] = html.unescape(m.group(1).strip())
    if stats_html and stats_html.exists():
        stats = read_text(stats_html)
        counts = re.findall(
            r'data-link="([^"]+)".*?class="child_color">\((\d+)\)</span>',
            stats,
            flags=re.S,
        )
        if counts:
            overview["section_counts"] = {html.unescape(k): int(v) for k, v in counts}
    return overview


def parse_stats_issues(stats_html: Path | None) -> list[dict[str, Any]]:
    if not stats_html or not stats_html.exists():
        return []
    text = read_text(stats_html)
    issues: list[dict[str, Any]] = []
    if re.search(r"顺序是否准确：\s*<span[^>]*>\s*否\s*</span>", text):
        issues.append(
            {
                "surface": "document_structure_order",
                "name": "结构顺序",
                "actual": "当前论文结构顺序与模板不一致",
                "expected": "标准结构顺序",
                "source": "stats-html",
                "blocking": True,
            }
        )
    stats_patterns = [
        ("front_matter_footer_page_number", "页码样式问题", "页码样式错误"),
        ("running_header_content", "页眉内容", "页眉内容错误"),
        ("running_header_rule", "页眉横线问题", "页眉无横线"),
    ]
    for surface, name, phrase in stats_patterns:
        count = len(re.findall(re.escape(phrase), text))
        if count:
            issues.append(
                {
                    "surface": surface,
                    "name": name,
                    "actual": phrase,
                    "expected": "模板对应页眉页脚设置",
                    "source": "stats-html",
                    "count": count,
                    "blocking": True,
                }
            )
    return issues


def text_of(element: ET.Element) -> str:
    chunks: list[str] = []
    for node in element.iter():
        if node.tag == f"{W}t":
            chunks.append(node.text or "")
        elif node.tag == f"{W}tab":
            chunks.append("\t")
    return "".join(chunks)


def map_comment_targets(document_root: ET.Element) -> dict[str, list[str]]:
    targets: dict[str, list[str]] = collections.defaultdict(list)
    for paragraph in document_root.findall(".//w:p", NS):
        ids = [node.get(f"{W}id") for node in paragraph.findall(".//w:commentRangeStart", NS)]
        ids = [item for item in ids if item is not None]
        if not ids:
            continue
        paragraph_text = text_of(paragraph).strip()
        for comment_id in ids:
            targets[str(comment_id)].append(paragraph_text)
    return dict(targets)


ISSUE_NAME_PATTERN = (
    r"标题空行问题|字体问题|字号问题|页码样式问题|空白页问题|页眉内容|"
    r"段前间距|段后间距|标点符号问题"
)


def split_comment_issue_text(comment_text: str) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    pattern = re.compile(
        rf"(?P<label>.*?)(?P<name>{ISSUE_NAME_PATTERN})："
        rf"(?P<actual>.*?)，规范：(?P<expected>.*?)"
        rf"(?=(?:目录|页面设置|正文|参考文献|致谢)[^：]{{0,90}}(?:{ISSUE_NAME_PATTERN})：|$)"
    )
    for match in pattern.finditer(comment_text):
        label = match.group("label").strip("-")
        name = match.group("name").strip()
        surface = label.strip("-") or name
        rows.append(
            {
                "surface": surface,
                "name": name,
                "actual": match.group("actual").strip(),
                "expected": match.group("expected").strip(),
            }
        )
    if not rows and comment_text.strip():
        rows.append(
            {
                "surface": "unparsed_comment",
                "name": "unparsed_comment",
                "actual": comment_text.strip(),
                "expected": "",
            }
        )
    return rows


def parse_comment_docx(comment_docx: Path | None) -> dict[str, Any] | None:
    if not comment_docx:
        return None
    with zipfile.ZipFile(comment_docx, "r") as archive:
        names = set(archive.namelist())
        if "word/comments.xml" not in names:
            return {
                "path": str(comment_docx),
                "sha256": sha256_file(comment_docx),
                "comment_count": 0,
                "summary": {},
                "issues": [],
            }
        comments_root = ET.fromstring(archive.read("word/comments.xml"))
        document_root = ET.fromstring(archive.read("word/document.xml"))
    targets = map_comment_targets(document_root)
    summary: dict[str, Any] = {}
    issues: list[dict[str, Any]] = []
    comment_count = 0
    for comment in comments_root.findall("./w:comment", NS):
        comment_count += 1
        comment_id = str(comment.get(f"{W}id") or "")
        comment_text = text_of(comment).strip()
        if comment_id == "0":
            for key, pattern in {
                "official_error_count": r"检测错误数：(\d+)",
                "severe_error_count": r"严重错误：(\d+)",
                "normal_error_count": r"一般错误：(\d+)",
                "reminder_count": r"提醒：(\d+)",
            }.items():
                found = re.search(pattern, comment_text)
                if found:
                    summary[key] = int(found.group(1))
            template = re.search(r"检测模板：《([^》]+)》", comment_text)
            if template:
                summary["template_name"] = template.group(1)
            summary["raw_text"] = comment_text[:500]
            continue
        for row in split_comment_issue_text(comment_text):
            row.update(
                {
                    "comment_id": comment_id,
                    "target_texts": targets.get(comment_id, []),
                    "source": "comment-docx",
                    "blocking": row.get("name") != "标点符号问题",
                }
            )
            issues.append(row)
    return {
        "path": str(comment_docx),
        "sha256": sha256_file(comment_docx),
        "comment_count": comment_count,
        "summary": summary,
        "issues": issues,
    }


def normalize_issue_rows(parag_map: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for segment_id, record in parag_map.items():
        if not isinstance(record, dict):
            continue
        errors = record.get("totalError") or []
        parts = record.get("parts") or []
        if not isinstance(errors, list):
            continue
        if not isinstance(parts, list):
            parts = []
        text = "".join(str(part.get("t") or "") for part in parts if isinstance(part, dict)).strip()
        page_index = None
        titles: list[str] = []
        structures: list[str] = []
        paragraph_part_ids: list[Any] = []
        for part in parts:
            if not isinstance(part, dict):
                continue
            if page_index is None:
                page_index = part.get("pageIndex")
            title = part.get("title")
            if title and title not in titles:
                titles.append(str(title))
            structure = part.get("struc")
            if structure and structure not in structures:
                structures.append(str(structure))
            if part.get("id") is not None:
                paragraph_part_ids.append(part.get("id"))
        for err in errors:
            if not isinstance(err, dict):
                continue
            rows.append(
                {
                    "segment_id": str(segment_id),
                    "paragraph_part_ids": paragraph_part_ids,
                    "page_index": page_index,
                    "structures": structures,
                    "titles": titles,
                    "text": text[:240],
                    "category_top": err.get("categoryTop"),
                    "category_second": err.get("categorySecond"),
                    "level": err.get("level"),
                    "name": err.get("name"),
                    "actual": err.get("error"),
                    "expected": err.get("temp"),
                    "desc": err.get("desc"),
                    "part_id": err.get("partId"),
                }
            )
    return rows


def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    summary: dict[str, Any] = {}
    for field in ("category_top", "category_second", "name"):
        summary[field] = collections.Counter(str(row.get(field) or "") for row in rows).most_common()
    by_structure: collections.Counter[str] = collections.Counter()
    by_title: collections.Counter[str] = collections.Counter()
    by_expected: collections.Counter[str] = collections.Counter()
    for row in rows:
        structures = row.get("structures") or ["UNKNOWN"]
        titles = row.get("titles") or ["UNKNOWN"]
        for item in structures:
            by_structure[str(item)] += 1
        for item in titles:
            by_title[str(item)] += 1
        expected = row.get("expected")
        if expected:
            by_expected[f"{row.get('name')} -> {expected}"] += 1
    summary["by_structure"] = by_structure.most_common()
    summary["by_title"] = by_title.most_common()
    summary["by_expected"] = by_expected.most_common()
    return summary


def summarize_comment_issues(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "by_surface": collections.Counter(str(row.get("surface") or "") for row in rows).most_common(),
        "by_name": collections.Counter(str(row.get("name") or "") for row in rows).most_common(),
        "by_expected": collections.Counter(
            f"{row.get('name')} -> {row.get('expected')}" for row in rows if row.get("expected")
        ).most_common(),
    }


def write_summary_markdown(path: Path, ledger: dict[str, Any], limit: int) -> None:
    rows = ledger["issues"]
    summaries = ledger["summaries"]
    stats_issues = ledger.get("stats_issues") or []
    comment_docx = ledger.get("comment_docx") or {}
    comment_issues = comment_docx.get("issues") or []
    lines = [
        "# Fanyu Format Report Issue Summary",
        "",
        f"- compare html: {ledger['compare_html']}",
        f"- stats html: {ledger.get('stats_html') or 'none'}",
        f"- generated time: {ledger.get('overview', {}).get('generated_time', 'unknown')}",
        f"- template name: {ledger.get('overview', {}).get('template_name', 'unknown')}",
        f"- report format error count: {ledger.get('overview', {}).get('format_error_count', 'unknown')}",
        f"- parsed issue rows: {len(rows)}",
        f"- comment docx: {comment_docx.get('path', 'none')}",
        f"- comment count: {comment_docx.get('comment_count', 0)}",
        f"- parsed comment issue rows: {len(comment_issues)}",
        f"- official comment error count: {(comment_docx.get('summary') or {}).get('official_error_count', 'unknown')}",
        "",
        "## Categories",
    ]
    for key, value in summaries.get("category_top", []):
        lines.append(f"- {key}: {value}")
    lines.append("")
    lines.append("## Problem Names")
    for key, value in summaries.get("name", []):
        lines.append(f"- {key}: {value}")
    lines.append("")
    lines.append("## Detected Surface Titles")
    for key, value in summaries.get("by_title", [])[:50]:
        lines.append(f"- {key}: {value}")
    if stats_issues:
        lines.append("")
        lines.append("## Stats/Structure/Page Issues")
        for issue in stats_issues:
            count = issue.get("count")
            suffix = f" count={count}" if count is not None else ""
            lines.append(
                f"- {issue.get('surface')}: {issue.get('name')} "
                f"actual={issue.get('actual')} expected={issue.get('expected')}{suffix}"
            )
    if comment_issues:
        lines.append("")
        lines.append("## Comment DOCX Issue Families")
        for key, value in (ledger.get("comment_summaries") or {}).get("by_surface", []):
            lines.append(f"- {key}: {value}")
        lines.append("")
        lines.append("## First Comment Issue Rows")
        for row in comment_issues[:limit]:
            target = "; ".join(row.get("target_texts") or [])[:100]
            lines.append(
                f"- comment {row.get('comment_id')} {row.get('surface')}/{row.get('name')}: "
                f"actual={row.get('actual')} expected={row.get('expected')} target={target}"
            )
    lines.append("")
    lines.append(f"## First {limit} Issue Rows")
    for row in rows[:limit]:
        structures = ",".join(row.get("structures") or [])
        titles = ",".join(row.get("titles") or [])
        text = str(row.get("text") or "").replace("\n", " ")[:100]
        lines.append(
            f"- seg {row.get('segment_id')} page {row.get('page_index')} "
            f"{structures}/{titles}: {row.get('name')} actual={row.get('actual')} "
            f"expected={row.get('expected')} text={text}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--compare-html", required=True)
    parser.add_argument("--stats-html")
    parser.add_argument("--comment-docx")
    parser.add_argument("--out-json", required=True)
    parser.add_argument("--out-md")
    parser.add_argument("--sample-limit", type=int, default=60)
    args = parser.parse_args(argv)

    compare_html = Path(args.compare_html).resolve()
    stats_html = Path(args.stats_html).resolve() if args.stats_html else None
    comment_docx = Path(args.comment_docx).resolve() if args.comment_docx else None
    out_json = Path(args.out_json).resolve()
    out_md = Path(args.out_md).resolve() if args.out_md else None

    parag_map = parse_parag_json(compare_html)
    rows = normalize_issue_rows(parag_map)
    comment_docx_payload = parse_comment_docx(comment_docx)
    comment_issues = (comment_docx_payload or {}).get("issues") or []
    ledger = {
        "schema": "graduation-project-builder.fanyu-format-report-ledger.v1",
        "generator": "scripts/parse_fanyu_format_report.py",
        "compare_html": str(compare_html),
        "stats_html": str(stats_html) if stats_html else None,
        "comment_docx": comment_docx_payload,
        "overview": parse_overview(compare_html, stats_html),
        "stats_issues": parse_stats_issues(stats_html),
        "segments_with_errors": len(parag_map),
        "issue_count": len(rows),
        "comment_issue_count": len(comment_issues),
        "summaries": summarize(rows),
        "comment_summaries": summarize_comment_issues(comment_issues),
        "issues": rows,
    }
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(ledger, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    if out_md:
        out_md.parent.mkdir(parents=True, exist_ok=True)
        write_summary_markdown(out_md, ledger, args.sample_limit)
    print(json.dumps({k: ledger[k] for k in ("schema", "segments_with_errors", "issue_count")}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
