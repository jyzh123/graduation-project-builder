#!/usr/bin/env python3
"""Audit school-style bibliography quantity and citation-use requirements."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@dataclass(frozen=True)
class ParagraphRecord:
    text: str
    has_numbering: bool


@dataclass(frozen=True)
class BibliographyEntry:
    text: str
    numbering_style: str


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def docx_paragraphs(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return [paragraph_text(p).strip() for p in root.findall(".//w:body/w:p", NS)]


def docx_paragraph_records(docx: Path) -> list[ParagraphRecord]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    records: list[ParagraphRecord] = []
    for paragraph in root.findall(".//w:body/w:p", NS):
        records.append(
            ParagraphRecord(
                text=paragraph_text(paragraph).strip(),
                has_numbering=paragraph.find("./w:pPr/w:numPr", NS) is not None,
            )
        )
    return records


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def is_references_heading(text: str) -> bool:
    normalized = compact(text).rstrip(":：")
    if normalized == "参考文献":
        return True
    if not normalized.startswith("参考文献"):
        return False
    suffix = normalized[len("参考文献") :]
    if re.fullmatch(r"[0-9ivxlcdmIVXLCDM]+", suffix):
        return False
    return suffix.lower() in {"n", ""}


def is_tail_heading(text: str) -> bool:
    return compact(text) in {"致谢", "附录", "附录A", "Appendix"}


def bibliography_block(paragraphs: list[str]) -> list[str]:
    start = None
    for index, text in enumerate(paragraphs):
        if is_references_heading(text):
            start = index + 1
            break
    if start is None:
        return []
    result: list[str] = []
    for text in paragraphs[start:]:
        stripped = text.strip()
        if not stripped:
            continue
        if is_tail_heading(stripped):
            break
        if re.match(r"^(?:\[\d+\]|\d+\.)", stripped):
            result.append(stripped)
        elif re.search(r"\[[A-Z]+(?:/[A-Z]+)?\]", stripped) or re.search(r"\b(?:20\d{2}|19\d{2})\b", stripped):
            result.append(stripped)
        elif result:
            result[-1] += " " + stripped
    return result


def bibliography_block_records(records: list[ParagraphRecord]) -> list[BibliographyEntry]:
    start = None
    for index, record in enumerate(records):
        if is_references_heading(record.text):
            start = index + 1
            break
    if start is None:
        return []
    result: list[BibliographyEntry] = []
    for record in records[start:]:
        stripped = record.text.strip()
        if not stripped:
            continue
        if is_tail_heading(stripped):
            break
        manual_style = bibliography_numbering_style(stripped)
        if manual_style != "none":
            result.append(BibliographyEntry(text=stripped, numbering_style=manual_style))
        elif record.has_numbering:
            result.append(BibliographyEntry(text=stripped, numbering_style="automatic-decimal"))
        elif re.search(r"\[[A-Z]+(?:/[A-Z]+)?\]", stripped) or re.search(r"\b(?:20\d{2}|19\d{2})\b", stripped):
            result.append(BibliographyEntry(text=stripped, numbering_style="none"))
        elif result:
            previous = result[-1]
            result[-1] = BibliographyEntry(
                text=previous.text + " " + stripped,
                numbering_style=previous.numbering_style,
            )
    return result


def bibliography_example_block(paragraphs: list[str]) -> list[str]:
    """Return numbered reference examples when a school spec lacks a live bibliography block."""
    start = None
    for index, text in enumerate(paragraphs):
        normalized = compact(text).rstrip(":：")
        if normalized in {"参考文献举例", "参考文献示例"}:
            start = index + 1
            break
    if start is None:
        return []
    result: list[str] = []
    for text in paragraphs[start:]:
        stripped = text.strip()
        if not stripped:
            continue
        if compact(stripped).rstrip(":：") in {"表格举例", "公式举例", "图举例"}:
            break
        if re.match(r"^(?:\[\d+\]|\d+\.)", stripped):
            result.append(stripped)
    return result


def body_text_before_references(paragraphs: list[str]) -> str:
    body: list[str] = []
    for text in paragraphs:
        if is_references_heading(text):
            break
        body.append(text)
    return "\n".join(body)


def ref_number(entry: str) -> int | None:
    match = re.match(r"^(?:\[(\d+)\]|(\d+)\.)", entry.strip())
    if not match:
        return None
    return int(match.group(1) or match.group(2))


def bibliography_entry_content_text(entry: str) -> str:
    content = re.sub(r"^(?:\[\s*\d+\s*\]|\d+[\.\u3001])", "", entry.strip(), count=1)
    content = re.sub(r"^\s*[\.\u3001\u3002\]\)）:：,，;；-]+", "", content)
    return re.sub(r"\s+", "", content)


def bibliography_numbering_style(entry: str) -> str:
    stripped = entry.strip()
    if re.match(r"^\[\d+\]", stripped):
        return "bracket"
    if re.match(r"^\d+\.", stripped):
        return "arabic-dot"
    return "none"


def ref_kind(entry: str) -> str:
    match = re.search(r"\[([A-Z]+(?:/[A-Z]+)?)\]", entry)
    return match.group(1) if match else ""


def ref_year(entry: str) -> int | None:
    years = [int(value) for value in re.findall(r"\b(20\d{2}|19\d{2})\b", entry)]
    return max(years) if years else None


def is_web(entry: str) -> bool:
    kind = ref_kind(entry)
    return kind in {"EB/OL", "OL"} or bool(re.search(r"https?://|www\.|docs\.", entry, re.I))


def is_foreign(entry: str) -> bool:
    body = re.sub(r"^\[\d+\]\s*", "", entry)
    latin = len(re.findall(r"[A-Za-z]", body))
    cjk = len(re.findall(r"[\u4e00-\u9fff]", body))
    return latin >= 20 and latin >= cjk


def citation_numbers(body_text: str) -> list[int]:
    return [int(value) for value in re.findall(r"\[(\d+)\]", body_text)]


def malformed_citation_groups(body_text: str) -> list[str]:
    pattern = re.compile(r"\[(?:\d+\s*[-,，、]\s*)+\d+\]")
    return sorted(set(match.group(0) for match in pattern.finditer(body_text)))


def contiguous_citation_groups(body_text: str) -> list[list[int]]:
    groups: list[list[int]] = []
    pattern = re.compile(r"(?:\[\d+\]\s*){2,}")
    for match in pattern.finditer(body_text):
        groups.append([int(value) for value in re.findall(r"\[(\d+)\]", match.group(0))])
    return groups


def paragraph_citation_warnings(paragraphs: list[str], max_per_paragraph: int) -> list[dict[str, object]]:
    warnings: list[dict[str, object]] = []
    for index, text in enumerate(paragraphs, start=1):
        nums = citation_numbers(text)
        unique = sorted(set(nums))
        if len(unique) > max_per_paragraph:
            warnings.append({"paragraph_index": index, "citation_numbers": unique, "text_prefix": text[:120]})
    return warnings


def dominant_numbering_style(entries: list[str]) -> str:
    numbering_styles = [bibliography_numbering_style(entry) for entry in entries]
    return max(set(numbering_styles), key=numbering_styles.count) if numbering_styles else "none"


def dominant_numbering_style_values(numbering_styles: list[str]) -> str:
    return max(set(numbering_styles), key=numbering_styles.count) if numbering_styles else "none"


def reference_numbering_style(reference_docx: Path | None) -> str:
    if reference_docx is None or not reference_docx.exists():
        return ""
    records = docx_paragraph_records(reference_docx)
    record_entries = bibliography_block_records(records)
    if record_entries:
        return dominant_numbering_style_values([entry.numbering_style for entry in record_entries])
    paragraphs = [record.text for record in records]
    entries = bibliography_example_block(paragraphs)
    return dominant_numbering_style(entries)


def audit(
    docx: Path,
    *,
    recent_start_year: int,
    max_group_size: int,
    min_reference_count: int,
    min_journal_count: int,
    min_foreign_count: int,
    min_recent_count: int,
    reference_docx: Path | None = None,
    expected_numbering_style: str | None = None,
) -> dict[str, object]:
    records = docx_paragraph_records(docx)
    paragraphs = [record.text for record in records]
    entry_records = bibliography_block_records(records)
    entries = [entry.text for entry in entry_records]
    body_text = body_text_before_references(paragraphs)
    body_numbers = citation_numbers(body_text)
    body_unique = sorted(set(body_numbers))
    ref_numbers = [
        number if number is not None else index
        for index, number in enumerate((ref_number(entry) for entry in entries), start=1)
    ]
    kinds = {kind: 0 for kind in ("J", "C", "M", "D", "EB/OL", "OL", "other")}
    for entry in entries:
        kind = ref_kind(entry) or "other"
        kinds[kind if kind in kinds else "other"] += 1
    non_web_entries = [entry for entry in entries if not is_web(entry)]
    journal_entries = [entry for entry in non_web_entries if ref_kind(entry) == "J"]
    foreign_entries = [entry for entry in non_web_entries if is_foreign(entry)]
    recent_entries = [entry for entry in non_web_entries if (ref_year(entry) or 0) >= recent_start_year]
    malformed_groups = malformed_citation_groups(body_text)
    contiguous_groups = contiguous_citation_groups(body_text)
    too_large_groups = [group for group in contiguous_groups if len(group) > max_group_size]
    missing_from_body = sorted(set(ref_numbers) - set(body_unique))
    extra_body_only = sorted(set(body_unique) - set(ref_numbers))
    empty_content_entries = [
        {"number": ref_number(entry) or index, "text": entry}
        for index, entry in enumerate(entries, start=1)
        if len(bibliography_entry_content_text(entry)) < 8
    ]
    ordered = body_unique == list(range(1, max(body_unique) + 1)) if body_unique else False
    docx_numbering_style = dominant_numbering_style_values([entry.numbering_style for entry in entry_records])
    template_numbering_style = reference_numbering_style(reference_docx)
    active_expected_numbering_style = expected_numbering_style or template_numbering_style
    allowed_numbering_styles = (
        {active_expected_numbering_style}
        if active_expected_numbering_style
        else {"bracket", "arabic-dot", "automatic-decimal"}
    )
    issues: list[str] = []
    if len(entries) < min_reference_count:
        issues.append(f"reference count below {min_reference_count}: {len(entries)}")
    if len(journal_entries) < min_journal_count:
        issues.append(f"journal/non-web reference count below {min_journal_count}: {len(journal_entries)}")
    if len(foreign_entries) < min_foreign_count:
        issues.append(f"foreign/non-web reference count below {min_foreign_count}: {len(foreign_entries)}")
    if len(recent_entries) < min_recent_count:
        issues.append(
            f"recent non-web reference count below {min_recent_count} since {recent_start_year}: "
            f"{len(recent_entries)}"
        )
    if missing_from_body:
        issues.append(f"bibliography entries not cited in body: {missing_from_body}")
    if extra_body_only:
        issues.append(f"body citations without bibliography entries: {extra_body_only}")
    if empty_content_entries:
        issues.append(
            "bibliography entries with missing substantive content: "
            + str([item["number"] for item in empty_content_entries])
        )
    if body_unique and not ordered:
        issues.append(f"body first-use citation numbers are not a continuous ordered chain: {body_unique}")
    if malformed_groups:
        issues.append(f"malformed grouped/range citation markers found: {malformed_groups}")
    if too_large_groups:
        issues.append(f"contiguous citation group exceeds {max_group_size}: {too_large_groups}")
    if docx_numbering_style not in allowed_numbering_styles:
        expected = active_expected_numbering_style or "bracket, arabic-dot, or automatic-decimal"
        issues.append(
            f"bibliography entry numbering style must follow template example `{expected}`, "
            f"found dominant style: {docx_numbering_style}"
        )
    style_policy = (
        "explicit-user-or-label-decision"
        if expected_numbering_style
        else ("template-derived" if template_numbering_style else "accepted GB/T variants")
    )
    return {
        "schema": "graduation-project-builder.bibliography-school-requirements.v1",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "docx_path": str(docx),
        "docx_sha256": sha256_file(docx),
        "recent_start_year": recent_start_year,
        "min_reference_count": min_reference_count,
        "min_journal_count": min_journal_count,
        "min_foreign_count": min_foreign_count,
        "min_recent_count": min_recent_count,
        "reference_count": len(entries),
        "non_web_reference_count": len(non_web_entries),
        "journal_non_web_count": len(journal_entries),
        "foreign_non_web_count": len(foreign_entries),
        "recent_non_web_count": len(recent_entries),
        "kind_counts": kinds,
        "body_unique_citations": body_unique,
        "bibliography_numbers": ref_numbers,
        "bibliography_numbering_style": docx_numbering_style,
        "bibliography_expected_numbering_style": active_expected_numbering_style or "bracket-or-arabic-dot",
        "bibliography_template_numbering_style": template_numbering_style or "",
        "bibliography_numbering_style_policy": style_policy,
        "missing_from_body": missing_from_body,
        "extra_body_only": extra_body_only,
        "empty_content_entries": empty_content_entries,
        "malformed_groups": malformed_groups,
        "contiguous_citation_groups": contiguous_groups,
        "too_large_contiguous_groups": too_large_groups,
        "paragraph_multi_citation_warnings": paragraph_citation_warnings(paragraphs, max_group_size),
        "issues": issues,
        "passed": not issues,
    }


def write_markdown(report: dict[str, object], output: Path) -> None:
    lines = [
        "# Bibliography School Requirements Audit",
        "",
        f"- docx path: {report['docx_path']}",
        f"- docx sha256: {report['docx_sha256']}",
        f"- minimum reference count: {report['min_reference_count']}",
        f"- minimum journal/non-web count: {report['min_journal_count']}",
        f"- minimum foreign/non-web count: {report['min_foreign_count']}",
        f"- minimum recent non-web count: {report['min_recent_count']}",
        f"- reference count: {report['reference_count']}",
        f"- non-web reference count: {report['non_web_reference_count']}",
        f"- journal non-web count: {report['journal_non_web_count']}",
        f"- foreign non-web count: {report['foreign_non_web_count']}",
        f"- recent non-web count since {report['recent_start_year']}: {report['recent_non_web_count']}",
        f"- body unique citations: {report['body_unique_citations']}",
        f"- bibliography empty/content-missing entries: {len(report.get('empty_content_entries') or [])}",
        f"- malformed grouped/range markers: {report['malformed_groups']}",
        f"- too-large contiguous citation groups: {report['too_large_contiguous_groups']}",
        f"- paragraph multi-citation warnings: {len(report['paragraph_multi_citation_warnings'])}",
        f"- result: {'pass' if report['passed'] else 'fail'}",
        "",
        "## Issues",
    ]
    issues = report.get("issues") or []
    if issues:
        lines.extend(f"- {issue}" for issue in issues)  # type: ignore[union-attr]
    else:
        lines.append("- none")
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--json-report", type=Path)
    parser.add_argument("--reference-docx", type=Path)
    parser.add_argument(
        "--expected-numbering-style",
        choices=("bracket", "arabic-dot", "automatic-decimal"),
        help="Explicit bibliography label-family decision from current user/template evidence. Overrides template-derived examples.",
    )
    parser.add_argument("--recent-start-year", type=int)
    parser.add_argument("--max-group-size", type=int, default=3)
    parser.add_argument("--min-reference-count", type=int, default=15)
    parser.add_argument("--min-journal-count", type=int, default=8)
    parser.add_argument("--min-foreign-count", type=int, default=3)
    parser.add_argument("--min-recent-count", type=int, default=5)
    args = parser.parse_args()
    recent_start_year = args.recent_start_year if args.recent_start_year is not None else datetime.now().year - 5
    report = audit(
        args.docx,
        recent_start_year=recent_start_year,
        max_group_size=args.max_group_size,
        min_reference_count=args.min_reference_count,
        min_journal_count=args.min_journal_count,
        min_foreign_count=args.min_foreign_count,
        min_recent_count=args.min_recent_count,
        reference_docx=args.reference_docx,
        expected_numbering_style=args.expected_numbering_style,
    )
    args.report.parent.mkdir(parents=True, exist_ok=True)
    write_markdown(report, args.report)
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
