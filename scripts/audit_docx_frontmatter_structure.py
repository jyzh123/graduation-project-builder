#!/usr/bin/env python3
"""Audit front-matter order, abstract surfaces, keyword runs, and style refs."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def paragraph_style_id(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", NS)
    return node.get(qn("val")) if node is not None else ""


def paragraph_numbering_level(paragraph: ET.Element) -> int | None:
    level_node = paragraph.find("./w:pPr/w:numPr/w:ilvl", NS)
    if level_node is None:
        return None
    value = level_node.get(qn("val"))
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def style_ids(styles_root: ET.Element) -> set[str]:
    return {
        style.get(qn("styleId")) or ""
        for style in styles_root.findall("./w:style", NS)
    }


def has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None


def is_toc_title(text: str) -> bool:
    return compact_text(text).lower() in {"目录", "contents", "tableofcontents"}


def is_toc_entry(paragraph: ET.Element) -> bool:
    text = paragraph_text(paragraph).strip()
    style_id = paragraph_style_id(paragraph).lower()
    if style_id.startswith("toc") or has_tab(paragraph) or "…" in text:
        return True
    if re.search(r"(?:\d+|[ivxlcdm]+)\s*$", text, flags=re.IGNORECASE) is None:
        return False
    stripped = re.sub(r"(?:\d+|[ivxlcdm]+)\s*$", "", compact_text(text), flags=re.IGNORECASE).lower()
    return stripped in {"摘要", "abstract", "关键词", "keywords", "keyword"}


def heading_level(text: str) -> int | None:
    stripped = (text or "").strip()
    dot = r"[.\uFF0E\u3002]"
    if re.match(r"^第[0-9一二三四五六七八九十]+章(?:\s*\S.*)?$", stripped):
        return 1
    if re.match(rf"^\d{{1,2}}{dot}\d{{1,2}}{dot}\d{{1,2}}\s+\S", stripped):
        return 3
    if re.match(rf"^\d{{1,2}}{dot}\d{{1,2}}\s+\S", stripped):
        return 2
    if re.match(r"^\d{1,2}\s+\S", stripped):
        return 1
    return None


def paragraph_heading_level(paragraph: ET.Element) -> int | None:
    numbering_level = paragraph_numbering_level(paragraph)
    if numbering_level is not None and numbering_level <= 3:
        return numbering_level + 1
    return heading_level(paragraph_text(paragraph))


def is_zh_abstract_title(text: str) -> bool:
    return compact_text(text) in {"摘要", "中文摘要"}


def is_en_abstract_title(text: str) -> bool:
    return compact_text(text).lower() == "abstract"


def is_zh_keyword(text: str) -> bool:
    return compact_text(text).startswith("关键词")


def is_en_keyword(text: str) -> bool:
    return compact_text(text).lower().startswith(("keywords", "keyword"))


def find_index(paragraphs: list[ET.Element], predicate) -> int | None:
    for index, paragraph in enumerate(paragraphs, start=1):
        if predicate(paragraph):
            return index
    return None


def find_index_after(paragraphs: list[ET.Element], after_index: int | None, predicate) -> int | None:
    start = after_index or 0
    for index, paragraph in enumerate(paragraphs[start:], start=start + 1):
        if predicate(paragraph):
            return index
    return None


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def run_is_bold(run: ET.Element) -> bool:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        return False
    for tag in ("b", "bCs"):
        node = rpr.find(f"./w:{tag}", NS)
        if node is None:
            continue
        value = (node.get(qn("val")) or "").strip().lower()
        if value not in {"0", "false", "off"}:
            return True
    return False


def normalize_label_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").replace("：", ":")).strip().lower()


def split_label_and_content_runs(
    runs: list[ET.Element],
    labels: tuple[str, ...],
) -> tuple[list[ET.Element], list[ET.Element]]:
    normalized_labels = {normalize_label_text(label) for label in labels}
    for end in range(1, len(runs) + 1):
        candidate = "".join(run_text(run) for run in runs[:end])
        if normalize_label_text(candidate) in normalized_labels:
            return runs[:end], runs[end:]
    return (runs[:1], runs[1:]) if runs else ([], [])


def paragraph_metric_detail(paragraph: ET.Element) -> dict[str, object]:
    ppr = paragraph.find("./w:pPr", NS)
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    jc = ppr.find("./w:jc", NS) if ppr is not None else None
    return {
        "style_id": paragraph_style_id(paragraph),
        "spacing": {
            "before": spacing.get(qn("before")) if spacing is not None else "",
            "after": spacing.get(qn("after")) if spacing is not None else "",
            "line": spacing.get(qn("line")) if spacing is not None else "",
            "lineRule": spacing.get(qn("lineRule")) if spacing is not None else "",
        },
        "indent": {
            "firstLine": ind.get(qn("firstLine")) if ind is not None else "",
            "firstLineChars": ind.get(qn("firstLineChars")) if ind is not None else "",
            "left": ind.get(qn("left")) if ind is not None else "",
            "right": ind.get(qn("right")) if ind is not None else "",
            "hanging": ind.get(qn("hanging")) if ind is not None else "",
        },
        "jc": jc.get(qn("val")) if jc is not None else "",
    }


def label_run_detail(paragraph: ET.Element, labels: tuple[str, ...] = ()) -> dict[str, object]:
    runs = [run for run in paragraph.findall("./w:r", NS) if run_text(run)]
    if labels:
        label_runs, content_run_candidates = split_label_and_content_runs(runs, labels)
    else:
        label_runs, content_run_candidates = (runs[:1], runs[1:]) if runs else ([], [])
    label_text = "".join(run_text(run) for run in label_runs)
    content_runs = [run for run in content_run_candidates if run_text(run).strip()]
    return {
        "run_count": len(runs),
        "label_text": label_text,
        "label_bold": any(run_is_bold(run) for run in label_runs),
        "content_run_count": len(content_runs),
        "content_first_text_prefix": run_text(content_runs[0])[:40] if content_runs else "",
        "content_bold_count": sum(1 for run in content_runs if run_is_bold(run)),
        "label_content_split": bool(runs and content_runs),
    }


def surface_detail(
    paragraphs: list[ET.Element],
    index: int | None,
    labels: tuple[str, ...] = (),
) -> dict[str, object] | None:
    if index is None:
        return None
    paragraph = paragraphs[index - 1]
    return {
        "paragraph_index": index,
        "text_prefix": paragraph_text(paragraph).strip()[:160],
        "metrics": paragraph_metric_detail(paragraph),
        "label_runs": label_run_detail(paragraph, labels),
    }


def label_run_issues(paragraph: ET.Element, labels: tuple[str, ...], surface: str) -> list[str]:
    runs = [run for run in paragraph.findall("./w:r", NS) if run_text(run)]
    if not runs:
        return [f"{surface} has no visible runs"]
    label_runs, content_run_candidates = split_label_and_content_runs(runs, labels)
    label_text = "".join(run_text(run) for run in label_runs)
    if normalize_label_text(label_text) not in {normalize_label_text(label) for label in labels}:
        return [f"{surface} first run sequence is not an isolated label: {label_text!r}"]
    issues: list[str] = []
    if not any(run_is_bold(run) for run in label_runs):
        issues.append(f"{surface} label run is not bold")
    content_runs = [run for run in content_run_candidates if run_text(run).strip()]
    if not content_runs:
        issues.append(f"{surface} has no content run after label")
    elif surface == "en_keyword" and not (
        label_text.endswith(" ") or "".join(run_text(run) for run in content_run_candidates).startswith(" ")
    ):
        issues.append(f"{surface} content run must start with a separator space after the bold label")
    if any(run_is_bold(run) for run in content_runs):
        issues.append(f"{surface} content run is bold")
    return issues


def direct_spacing_issues(paragraph: ET.Element, surface: str) -> list[str]:
    issues: list[str] = []
    ppr = paragraph.find("./w:pPr", NS)
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    expected_spacing = {"line": "360", "lineRule": "auto"}
    for key, expected in expected_spacing.items():
        actual = spacing.get(qn(key)) if spacing is not None else None
        if actual != expected:
            issues.append(f"{surface} paragraph {key} must be {expected}, found {actual or '<missing>'}")
    for key in ("before", "after"):
        actual = spacing.get(qn(key)) if spacing is not None else None
        if actual not in {None, "0"}:
            issues.append(f"{surface} paragraph {key} must be template-compatible 0 or omitted, found {actual}")
    return issues


def indented_body_issues(paragraph: ET.Element, surface: str) -> list[str]:
    issues = direct_spacing_issues(paragraph, surface)
    ppr = paragraph.find("./w:pPr", NS)
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    first_line = ind.get(qn("firstLine")) if ind is not None else None
    if first_line not in {"480", "482"}:
        issues.append(
            f"{surface} first-line indent must be template-compatible 480/482 twips (2 characters), "
            f"found {first_line or '<missing>'}"
        )
    first_line_chars = ind.get(qn("firstLineChars")) if ind is not None else None
    if first_line_chars != "200":
        issues.append(
            f"{surface} first-line indent must also include firstLineChars=200 for WPS/Word 2-character indentation, "
            f"found {first_line_chars or '<missing>'}"
        )
    return issues


def title_paragraph_issues(paragraph: ET.Element, surface: str, expected: tuple[str, ...]) -> list[str]:
    text = paragraph_text(paragraph).strip()
    compact = compact_text(text)
    if compact not in {compact_text(item) for item in expected}:
        return [f"{surface} title text is not an isolated abstract title: {text!r}"]
    if ":" in text or "：" in text:
        if surface == "en_abstract_title":
            return ["English abstract title must be a standalone title paragraph, not an inline `Abstract: body` paragraph"]
        return ["Chinese abstract title must be a standalone title paragraph, not an inline `摘要：正文` paragraph"]
    return []


def abstract_body_issues(paragraph: ET.Element, surface: str, *, english: bool) -> list[str]:
    text = paragraph_text(paragraph).strip()
    compact = compact_text(text).lower()
    if not text:
        return [f"{surface} has no abstract body text"]
    if english:
        if compact in {"abstract", "abstract:"}:
            return [f"{surface} content is only a duplicated abstract title, not the abstract body"]
        if compact.startswith(("keywords", "keyword")):
            return ["English abstract title must be followed by one standalone abstract body paragraph before the keyword line"]
    else:
        if compact in {"摘要", "中文摘要", "摘要:", "摘要："}:
            return [f"{surface} content is only a duplicated abstract title, not the abstract body"]
        if compact.startswith("关键词"):
            return ["Chinese abstract title must be followed by one standalone abstract body paragraph before the keyword line"]
    return []


def locate_body_paragraph(
    paragraphs: list[ET.Element],
    title_index: int | None,
    keyword_index: int | None,
    surface: str,
    issues: list[str],
    *,
    english: bool,
) -> int | None:
    if title_index is None or keyword_index is None or title_index >= keyword_index:
        return None
    nonempty = [
        index
        for index in range(title_index + 1, keyword_index)
        if paragraph_text(paragraphs[index - 1]).strip()
    ]
    if not nonempty:
        if english:
            issues.append("English abstract title must be followed by one standalone abstract body paragraph before the keyword line")
        else:
            issues.append("Chinese abstract title must be followed by one standalone abstract body paragraph before the keyword line")
        return None
    body_index = nonempty[0]
    if body_index != title_index + 1:
        issues.append(f"{surface} body paragraph must immediately follow the abstract title paragraph")
    if len(nonempty) > 1:
        extras = [
            {"paragraph": index, "text": paragraph_text(paragraphs[index - 1]).strip()[:120]}
            for index in nonempty[1:]
        ]
        issues.append(f"{surface} contains unexpected extra non-empty paragraphs before the keyword line: {extras[:3]}")
    return body_index


def audit_frontmatter(final_docx: Path) -> dict[str, object]:
    with zipfile.ZipFile(final_docx, "r") as zf:
        document_root = ET.fromstring(zf.read("word/document.xml"))
        styles_root = ET.fromstring(zf.read("word/styles.xml"))
    styles = style_ids(styles_root)
    paragraphs = document_root.findall("./w:body/w:p", NS)
    issues: list[str] = []
    undefined = []
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        style_id = paragraph_style_id(paragraph)
        if text and style_id and style_id not in styles:
            undefined.append({"paragraph": index, "style_id": style_id, "text": text[:120]})
    if undefined:
        issues.append(f"undefined non-empty paragraph styles: {undefined[:8]}")

    zh_abs = find_index(paragraphs, lambda p: is_zh_abstract_title(paragraph_text(p)) and not is_toc_entry(p))
    zh_key = find_index(paragraphs, lambda p: is_zh_keyword(paragraph_text(p)) and not is_toc_entry(p))
    en_abs = find_index(paragraphs, lambda p: is_en_abstract_title(paragraph_text(p)) and not is_toc_entry(p))
    en_key = find_index(paragraphs, lambda p: is_en_keyword(paragraph_text(p)) and not is_toc_entry(p))
    toc = find_index(paragraphs, lambda p: is_toc_title(paragraph_text(p)))
    first_body = find_index_after(paragraphs, toc, lambda p: paragraph_heading_level(p) == 1 and not is_toc_entry(p))

    zh_body = locate_body_paragraph(paragraphs, zh_abs, zh_key, "zh_abstract", issues, english=False)
    en_body = locate_body_paragraph(paragraphs, en_abs, en_key, "en_abstract", issues, english=True)

    surface_order = {
        "zh_abstract_title": zh_abs,
        "zh_abstract_body": zh_body,
        "zh_keyword": zh_key,
        "en_abstract_title": en_abs,
        "en_abstract_body": en_body,
        "en_keyword": en_key,
        "toc": toc,
        "first_body": first_body,
    }
    if any(value is None for value in surface_order.values()):
        issues.append(f"front-matter required surface missing: {surface_order}")
    elif not (
        zh_abs < zh_body < zh_key < en_abs < en_body < en_key < toc < first_body  # type: ignore[operator]
    ):
        issues.append(
            "front-matter order must be zh abstract title -> zh abstract body -> zh keyword -> "
            f"en abstract title -> en abstract body -> en keyword -> TOC -> body, found {surface_order}"
        )

    if zh_abs is not None:
        issues.extend(title_paragraph_issues(paragraphs[zh_abs - 1], "zh_abstract_title", ("摘要", "中文摘要")))
    if zh_body is not None:
        issues.extend(abstract_body_issues(paragraphs[zh_body - 1], "zh_abstract_body", english=False))
        issues.extend(indented_body_issues(paragraphs[zh_body - 1], "zh_abstract_body"))
    if zh_key is not None:
        issues.extend(label_run_issues(paragraphs[zh_key - 1], ("关键词：", "关键词:"), "zh_keyword"))
        issues.extend(direct_spacing_issues(paragraphs[zh_key - 1], "zh_keyword"))

    if en_abs is not None:
        issues.extend(title_paragraph_issues(paragraphs[en_abs - 1], "en_abstract_title", ("Abstract",)))
    if en_body is not None:
        issues.extend(abstract_body_issues(paragraphs[en_body - 1], "en_abstract_body", english=True))
        issues.extend(indented_body_issues(paragraphs[en_body - 1], "en_abstract_body"))
    if en_key is not None:
        issues.extend(
            label_run_issues(
                paragraphs[en_key - 1],
                ("Key words:", "Key Words:", "Keywords:", "Keyword:"),
                "en_keyword",
            )
        )
        issues.extend(direct_spacing_issues(paragraphs[en_key - 1], "en_keyword"))

    surface_details = {
        "zh_abstract_title": surface_detail(paragraphs, zh_abs),
        "zh_abstract_body": surface_detail(paragraphs, zh_body),
        "zh_keyword": surface_detail(paragraphs, zh_key, ("关键词：", "关键词:")),
        "en_abstract_title": surface_detail(paragraphs, en_abs),
        "en_abstract_body": surface_detail(paragraphs, en_body),
        "en_keyword": surface_detail(
            paragraphs,
            en_key,
            ("Key words:", "Key Words:", "Keywords:", "Keyword:"),
        ),
        "toc": surface_detail(paragraphs, toc),
        "first_body": surface_detail(paragraphs, first_body),
    }

    return {
        "schema": "graduation-project-builder.frontmatter-structure-audit.v1",
        "final_docx_path": str(final_docx),
        "final_docx_sha256": sha256_file(final_docx),
        "surface_order": surface_order,
        "surface_details": surface_details,
        "undefined_nonempty_style_count": len(undefined),
        "undefined_nonempty_styles": undefined,
        "issues": issues,
        "passed": not issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--report-json")
    args = parser.parse_args()
    payload = audit_frontmatter(Path(args.final_docx).resolve())
    if args.report_json:
        report = Path(args.report_json).resolve()
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
