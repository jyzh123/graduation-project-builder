#!/usr/bin/env python3
"""Audit thesis body citations for order, placement, and bibliography coverage."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
import xml.etree.ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]

BIBLIO_HEADING = "\u53c2\u8003\u6587\u732e"
TAIL_HEADINGS = {
    "\u81f4\u8c22",
    "\u9644\u5f55",
    "acknowledgements",
    "acknowledgments",
    "appendix",
}
TAIL_HEADINGS_NORMALIZED = {re.sub(r"\s+", "", item.lower()) for item in TAIL_HEADINGS}
BIBLIO_HEADING_NORMALIZED = re.sub(r"\s+", "", BIBLIO_HEADING)
BODY_EXCLUDE_EXACT = {
    "",
    "\u76ee\u5f55",
    "\u6458\u8981",
    "abstract",
    BIBLIO_HEADING,
}
CAPTION_RE = re.compile(r"^[\u56fe\u8868]\s*\d+-\d+\s+\S")
HEADING_RE = re.compile(
    r"^(?:\d+(?:\.\d+){0,3}\s+\S.*|"
    r"\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\u7ae0\u8282]\s*\S*)$"
)
BIBLIOGRAPHY_VISIBLE_NUMBER_RE = re.compile(r"^\s*\[\s*\d+\s*\]")

ERROR_CODES = {
    "order": "BODY_CITATION_ORDER",
    "forbidden": "BODY_CITATION_FORBIDDEN_SURFACE",
    "superscript": "BODY_CITATION_NOT_SUPERSCRIPT",
    "hyperlink": "BODY_CITATION_NOT_HYPERLINK",
    "target": "BODY_CITATION_WRONG_HYPERLINK_TARGET",
    "visual": "BODY_CITATION_VISUAL_STYLE",
    "multi_marker": "BODY_CITATION_MULTI_MARKER_SENTENCE",
    "multi_number": "BODY_CITATION_MULTI_NUMBER_MARKER",
    "punctuation": "BODY_CITATION_PUNCTUATION",
    "coverage": "UNCITED_BIBLIOGRAPHY_ITEMS",
    "out_of_range": "BODY_CITATION_OUT_OF_RANGE",
    "bibliography_numbering": "BIBLIOGRAPHY_MANUAL_AND_AUTO_NUMBERING",
    "anchor_leak": "BODY_CITATION_ANCHOR_LEAK",
}


@dataclass
class ParagraphRecord:
    index: int
    para_id: str
    text: str
    style_id: str
    style_name: str
    has_numbering: bool
    xml: ET.Element


@dataclass
class CitationRecord:
    paragraph_index: int
    para_id: str
    surface: str
    numbers: list[int]
    marker_count: int
    superscript_ok: bool
    hyperlink_ok: bool
    target_ok: bool
    visual_style_ok: bool
    punctuation_ok: bool
    text: str


@dataclass
class AuditResult:
    passed: bool
    error_codes: list[str]
    body_citation_paragraph_count: int
    unique_citation_count: int
    bibliography_item_count: int
    first_appearance_chain: list[int]
    expected_chain: list[int]
    missing_bibliography_numbers: list[int]
    extra_body_citation_numbers: list[int]
    order_failures: list[str]
    forbidden_surface_hits: list[str]
    non_superscript_hits: list[str]
    non_hyperlink_hits: list[str]
    wrong_hyperlink_target_hits: list[str]
    visual_style_failures: list[str]
    multi_marker_sentence_hits: list[str]
    multi_number_marker_hits: list[str]
    punctuation_failures: list[str]
    bibliography_numbering_conflicts: list[str]
    visible_anchor_leaks: list[str]
    citation_records: list[CitationRecord]


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def split_sentences(text: str) -> list[str]:
    parts = [seg.strip() for seg in re.findall(r"[^\u3002\uff01\uff1f\uff1b.!?;]+[\u3002\uff01\uff1f\uff1b.!?;]?", text) if seg.strip()]
    return parts if parts else ([text.strip()] if text.strip() else [])


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join((node.text or "") for node in paragraph.iterfind(".//w:t", NS))


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    if style is None:
        return ""
    return style.attrib.get(f"{W}val", "")


def paragraph_has_numbering(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:numPr", NS) is not None


def paragraph_id(paragraph: ET.Element) -> str:
    for key, value in paragraph.attrib.items():
        if key.endswith("paraId"):
            return value
    return ""


def collect_bookmark_locations(paragraphs: list[ParagraphRecord]) -> dict[str, int]:
    locations: dict[str, int] = {}
    for record in paragraphs:
        for bookmark in record.xml.findall(".//w:bookmarkStart", NS):
            name = bookmark.attrib.get(f"{W}name", "")
            if name and name not in locations:
                locations[name] = record.index
    return locations


def load_style_map(docx_path: Path) -> dict[str, str]:
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            styles_xml = zf.read("word/styles.xml")
    except KeyError:
        return {}

    root = ET.fromstring(styles_xml)
    result: dict[str, str] = {}
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(f"{W}styleId", "")
        name = style.find("w:name", NS)
        if style_id:
            result[style_id] = name.attrib.get(f"{W}val", "") if name is not None else ""
    return result


def iter_paragraphs(docx_path: Path) -> list[ParagraphRecord]:
    style_map = load_style_map(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    records: list[ParagraphRecord] = []
    for index, paragraph in enumerate(root.iterfind(".//w:body/w:p", NS)):
        style_id = paragraph_style_id(paragraph)
        records.append(
            ParagraphRecord(
                index=index,
                para_id=paragraph_id(paragraph),
                text=normalize_text(paragraph_text(paragraph)),
                style_id=style_id,
                style_name=style_map.get(style_id, ""),
                has_numbering=paragraph_has_numbering(paragraph),
                xml=paragraph,
            )
        )
    return records


def looks_like_heading(text: str) -> bool:
    stripped = normalize_text(text)
    if not stripped:
        return False
    lowered = stripped.lower()
    compact = re.sub(r"\s+", "", lowered)
    if compact in TAIL_HEADINGS_NORMALIZED or re.sub(r"\s+", "", stripped) == BIBLIO_HEADING_NORMALIZED:
        return True
    if len(stripped) > 60:
        return False
    return bool(HEADING_RE.match(stripped))


def find_bibliography_range(paragraphs: list[ParagraphRecord]) -> tuple[int, int]:
    start = next(
        (p.index for p in paragraphs if re.sub(r"\s+", "", p.text) == BIBLIO_HEADING_NORMALIZED),
        -1,
    )
    if start < 0:
        raise ValueError(f"Could not find bibliography heading '{BIBLIO_HEADING}'.")

    end = len(paragraphs)
    for record in paragraphs[start + 1 :]:
        lowered = record.text.lower()
        compact = re.sub(r"\s+", "", lowered)
        if compact in TAIL_HEADINGS_NORMALIZED:
            end = record.index
            break
        if looks_like_heading(record.text) and not record.has_numbering:
            end = record.index
            break
    return start, end


def classify_surface(record: ParagraphRecord) -> str:
    text = normalize_text(record.text)
    style_tokens = f"{record.style_id} {record.style_name}".lower()
    if text.lower() in {"\u76ee\u5f55", "table of contents"} or "toc" in style_tokens:
        return "toc"
    if CAPTION_RE.match(text) or any(
        token in style_tokens for token in ("caption", "\u56fe\u9898", "\u8868\u9898")
    ):
        return "caption"
    style_marks_heading = any(token in style_tokens for token in ("heading", "title", "\u6807\u9898"))
    if looks_like_heading(text) or (style_marks_heading and len(text) <= 60):
        return "heading"
    return "body"


def extract_citation_numbers(text: str) -> list[int]:
    return [int(match) for match in re.findall(r"\[(\d+)\]", text)]


def citation_marker_count(paragraph: ParagraphRecord, number: int) -> int:
    target = f"[{number}]"
    return sum(
        1
        for run in paragraph.xml.findall(".//w:r", NS)
        if "".join((node.text or "") for node in run.iterfind(".//w:t", NS)) == target
    )


def run_has_superscript_citation(paragraph: ParagraphRecord, number: int) -> bool:
    target = f"[{number}]"
    matched = 0
    for run in paragraph.xml.findall(".//w:r", NS):
        text = "".join((node.text or "") for node in run.iterfind(".//w:t", NS))
        if text != target:
            continue
        matched += 1
        vert = run.find("./w:rPr/w:vertAlign", NS)
        if vert is None or vert.attrib.get(f"{W}val") != "superscript":
            return False
    return matched > 0


def run_has_hyperlink_citation(paragraph: ParagraphRecord, number: int) -> bool:
    target = f"[{number}]"
    matched = 0
    for hyperlink in paragraph.xml.findall(".//w:hyperlink", NS):
        anchor = hyperlink.attrib.get(f"{W}anchor", "")
        if not anchor:
            continue
        for run in hyperlink.findall(".//w:r", NS):
            text = "".join((node.text or "") for node in run.iterfind(".//w:t", NS))
            if text == target:
                matched += 1
    return matched == citation_marker_count(paragraph, number) and matched > 0


def run_has_bibliography_target(
    paragraph: ParagraphRecord,
    number: int,
    *,
    bookmark_locations: dict[str, int],
    bibliography_start: int,
    bibliography_end: int,
) -> bool:
    target = f"[{number}]"
    anchor = f"cite_ref_{number}"
    target_count = citation_marker_count(paragraph, number)
    matched = 0
    for hyperlink in paragraph.xml.findall(".//w:hyperlink", NS):
        if hyperlink.attrib.get(f"{W}anchor", "") != anchor:
            continue
        target_index = bookmark_locations.get(anchor)
        if target_index is None or not (bibliography_start < target_index < bibliography_end):
            continue
        for run in hyperlink.findall(".//w:r", NS):
            text = "".join((node.text or "") for node in run.iterfind(".//w:t", NS))
            if text == target:
                matched += 1
    return matched == target_count and matched > 0


def citation_runs(paragraph: ParagraphRecord, number: int) -> list[ET.Element]:
    target = f"[{number}]"
    matched: list[ET.Element] = []
    for run in paragraph.xml.findall(".//w:r", NS):
        text = "".join((node.text or "") for node in run.iterfind(".//w:t", NS))
        if text == target:
            matched.append(run)
    return matched


def run_has_clean_visual_style(run: ET.Element) -> bool:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        return True
    rstyle = rpr.find("w:rStyle", NS)
    if rstyle is not None and "hyperlink" in rstyle.attrib.get(f"{W}val", "").lower():
        return False
    bold = rpr.find("w:b", NS)
    if bold is not None and bold.attrib.get(f"{W}val", "1") != "0":
        return False
    italic = rpr.find("w:i", NS)
    if italic is not None and italic.attrib.get(f"{W}val", "1") != "0":
        return False
    underline = rpr.find("w:u", NS)
    if underline is not None and underline.attrib.get(f"{W}val", "single").lower() not in {"none", "0"}:
        return False
    color = rpr.find("w:color", NS)
    if color is not None:
        value = color.attrib.get(f"{W}val", "").strip().lower()
        if value not in {"", "000000", "auto"}:
            return False
    return True


def citation_visual_style_ok(paragraph: ParagraphRecord, number: int) -> bool:
    runs = citation_runs(paragraph, number)
    if not runs:
        return False
    return all(run_has_clean_visual_style(run) for run in runs)


def citation_is_before_punctuation(text: str, number: int) -> bool:
    return bool(re.search(rf"\[{number}\]\s*[\u3002\uff01\uff1f\uff1b\uff0c\.,;!?]", text))


def sentence_marker_rule_violations(text: str) -> tuple[list[str], list[str]]:
    multi_marker_hits: list[str] = []
    multi_number_hits: list[str] = []
    for sentence in split_sentences(text):
        markers = re.findall(r"\[([^\]]+)\]", sentence)
        if not markers:
            continue
        if len(markers) > 1:
            multi_marker_hits.append(sentence)
        for marker in markers:
            if not re.fullmatch(r"\s*\d+\s*", marker):
                multi_number_hits.append(f"{sentence} :: [{marker}]")
    return multi_marker_hits, multi_number_hits


def extract_bibliography_entries(paragraphs: list[ParagraphRecord]) -> list[ParagraphRecord]:
    start, end = find_bibliography_range(paragraphs)
    return [record for record in paragraphs[start + 1 : end] if record.text]


def find_bibliography_numbering_conflicts(bibliography: list[ParagraphRecord]) -> list[str]:
    conflicts: list[str] = []
    for record in bibliography:
        if record.has_numbering and BIBLIOGRAPHY_VISIBLE_NUMBER_RE.search(record.text):
            conflicts.append(
                f"paragraph {record.index + 1} has both Word numbering and a visible manual bibliography number: {record.text}"
            )
    return conflicts


VISIBLE_ANCHOR_LEAK_RE = re.compile(r"\b(?:cite_ref|ref_anchor|bookmark)_\d+\b", re.IGNORECASE)


def visible_anchor_leak_hits(record: ParagraphRecord) -> list[str]:
    return sorted(set(match.group(0) for match in VISIBLE_ANCHOR_LEAK_RE.finditer(record.text or "")))


def audit_docx(docx_path: Path, *, allow_uncited_bibliography: bool = False) -> AuditResult:
    paragraphs = iter_paragraphs(docx_path)
    start, end = find_bibliography_range(paragraphs)
    bibliography = extract_bibliography_entries(paragraphs)
    body_records = paragraphs[:start]
    bookmark_locations = collect_bookmark_locations(paragraphs)

    seen: set[int] = set()
    unique_chain: list[int] = []
    order_failures: list[str] = []
    forbidden_surface_hits: list[str] = []
    non_superscript_hits: list[str] = []
    non_hyperlink_hits: list[str] = []
    wrong_hyperlink_target_hits: list[str] = []
    visual_style_failures: list[str] = []
    multi_marker_sentence_hits: list[str] = []
    multi_number_marker_hits: list[str] = []
    punctuation_failures: list[str] = []
    citation_records: list[CitationRecord] = []
    bibliography_numbering_conflicts = find_bibliography_numbering_conflicts(bibliography)
    visible_anchor_leaks: list[str] = []

    for record in body_records:
        leaks = visible_anchor_leak_hits(record)
        if leaks:
            visible_anchor_leaks.append(
                f"paragraph {record.index + 1} exposes visible internal citation anchor text {leaks}: {record.text}"
            )

        sentence_multi_marker_hits, sentence_multi_number_hits = sentence_marker_rule_violations(record.text)
        for sentence in sentence_multi_marker_hits:
            multi_marker_sentence_hits.append(
                f"paragraph {record.index + 1} contains more than one citation marker in one sentence: {sentence}"
            )
        for detail in sentence_multi_number_hits:
            multi_number_marker_hits.append(
                f"paragraph {record.index + 1} contains a grouped or non-single-number citation marker: {detail}"
            )

        numbers = extract_citation_numbers(record.text)
        if not numbers:
            continue

        surface = classify_surface(record)
        superscript_ok = all(run_has_superscript_citation(record, number) for number in numbers)
        hyperlink_ok = all(run_has_hyperlink_citation(record, number) for number in numbers)
        target_ok = all(
            run_has_bibliography_target(
                record,
                number,
                bookmark_locations=bookmark_locations,
                bibliography_start=start,
                bibliography_end=end,
            )
            for number in numbers
        )
        visual_style_ok = all(citation_visual_style_ok(record, number) for number in numbers)
        punctuation_ok = all(citation_is_before_punctuation(record.text, number) for number in numbers)
        marker_count = sum(citation_marker_count(record, number) for number in sorted(set(numbers)))

        citation_records.append(
            CitationRecord(
                paragraph_index=record.index + 1,
                para_id=record.para_id,
                surface=surface,
                numbers=numbers,
                marker_count=marker_count,
                superscript_ok=superscript_ok,
                hyperlink_ok=hyperlink_ok,
                target_ok=target_ok,
                visual_style_ok=visual_style_ok,
                punctuation_ok=punctuation_ok,
                text=record.text,
            )
        )

        if surface != "body":
            forbidden_surface_hits.append(
                f"paragraph {record.index + 1} carries citations on forbidden surface '{surface}': {record.text}"
            )

        if not superscript_ok:
            non_superscript_hits.append(
                f"paragraph {record.index + 1} has citation markers that are not preserved as superscript runs"
            )

        if not hyperlink_ok:
            non_hyperlink_hits.append(
                f"paragraph {record.index + 1} has citation markers that are not wrapped as internal hyperlinks"
            )

        if not target_ok:
            wrong_hyperlink_target_hits.append(
                f"paragraph {record.index + 1} has citation hyperlinks that do not resolve to bibliography entries inside the references block"
            )

        if not visual_style_ok:
            visual_style_failures.append(
                f"paragraph {record.index + 1} has citation markers that are not black, non-underlined, and style-clean"
            )

        if not punctuation_ok:
            punctuation_failures.append(
                f"paragraph {record.index + 1} has a citation marker that is not immediately before punctuation"
            )

        if surface != "body":
            continue

        for number in numbers:
            if number in seen:
                continue
            expected = len(seen) + 1
            if number != expected:
                order_failures.append(
                    f"paragraph {record.index + 1} introduces [{number}] as a new citation but expected [{expected}] by first appearance order"
                )
            seen.add(number)
            unique_chain.append(number)

    bibliography_count = len(bibliography)
    unique_numbers = sorted(set(unique_chain))
    missing_bibliography_numbers = (
        [] if allow_uncited_bibliography else sorted(set(range(1, bibliography_count + 1)) - set(unique_numbers))
    )
    extra_body_citation_numbers = sorted(number for number in unique_numbers if number > bibliography_count)
    expected_chain = list(range(1, len(unique_chain) + 1))

    error_codes: set[str] = set()
    if order_failures:
        error_codes.add(ERROR_CODES["order"])
    if forbidden_surface_hits:
        error_codes.add(ERROR_CODES["forbidden"])
    if non_superscript_hits:
        error_codes.add(ERROR_CODES["superscript"])
    if non_hyperlink_hits:
        error_codes.add(ERROR_CODES["hyperlink"])
    if wrong_hyperlink_target_hits:
        error_codes.add(ERROR_CODES["target"])
    if visual_style_failures:
        error_codes.add(ERROR_CODES["visual"])
    if multi_marker_sentence_hits:
        error_codes.add(ERROR_CODES["multi_marker"])
    if multi_number_marker_hits:
        error_codes.add(ERROR_CODES["multi_number"])
    if punctuation_failures:
        error_codes.add(ERROR_CODES["punctuation"])
    if missing_bibliography_numbers:
        error_codes.add(ERROR_CODES["coverage"])
    if extra_body_citation_numbers:
        error_codes.add(ERROR_CODES["out_of_range"])
    if bibliography_numbering_conflicts:
        error_codes.add(ERROR_CODES["bibliography_numbering"])
    if visible_anchor_leaks:
        error_codes.add(ERROR_CODES["anchor_leak"])

    return AuditResult(
        passed=not error_codes,
        error_codes=sorted(error_codes),
        body_citation_paragraph_count=len(citation_records),
        unique_citation_count=len(unique_chain),
        bibliography_item_count=bibliography_count,
        first_appearance_chain=unique_chain,
        expected_chain=expected_chain,
        missing_bibliography_numbers=missing_bibliography_numbers,
        extra_body_citation_numbers=extra_body_citation_numbers,
        order_failures=order_failures,
        forbidden_surface_hits=forbidden_surface_hits,
        non_superscript_hits=non_superscript_hits,
        non_hyperlink_hits=non_hyperlink_hits,
        wrong_hyperlink_target_hits=wrong_hyperlink_target_hits,
        visual_style_failures=visual_style_failures,
        multi_marker_sentence_hits=multi_marker_sentence_hits,
        multi_number_marker_hits=multi_number_marker_hits,
        punctuation_failures=punctuation_failures,
        bibliography_numbering_conflicts=bibliography_numbering_conflicts,
        visible_anchor_leaks=visible_anchor_leaks,
        citation_records=citation_records,
    )


def build_report(result: AuditResult, docx_path: Path) -> str:
    lines = [
        "# Body Citation Audit Report",
        "",
        "## Summary",
        f"- document path: {docx_path}",
        f"- document sha256: {sha256_file(docx_path)}",
        f"- body citation paragraph count: {result.body_citation_paragraph_count}",
        f"- unique citation count: {result.unique_citation_count}",
        f"- bibliography item count: {result.bibliography_item_count}",
        f"- first appearance chain: {result.first_appearance_chain if result.first_appearance_chain else '[]'}",
        f"- expected chain: {result.expected_chain if result.expected_chain else '[]'}",
        f"- missing bibliography numbers: {result.missing_bibliography_numbers if result.missing_bibliography_numbers else 'none'}",
        f"- extra body citation numbers: {result.extra_body_citation_numbers if result.extra_body_citation_numbers else 'none'}",
        f"- bibliography numbering conflicts: {len(result.bibliography_numbering_conflicts)}",
        f"- error codes: {', '.join(result.error_codes) if result.error_codes else 'none'}",
        f"- result: {'pass' if result.passed else 'fail'}",
        "",
        "## Body Citation Records",
    ]
    if result.citation_records:
        for record in result.citation_records:
            lines.extend(
                [
                    f"### Paragraph {record.paragraph_index}",
                    f"- paragraph index: {record.paragraph_index}",
                    f"- paragraph id: {record.para_id or 'none'}",
                    f"- surface: {record.surface}",
                    f"- citation numbers: {record.numbers}",
                    f"- marker count: {record.marker_count}",
                    f"- superscript ok: {'yes' if record.superscript_ok else 'no'}",
                    f"- hyperlink ok: {'yes' if record.hyperlink_ok else 'no'}",
                    f"- bibliography target ok: {'yes' if record.target_ok else 'no'}",
                    f"- visual style ok: {'yes' if record.visual_style_ok else 'no'}",
                    f"- punctuation ok: {'yes' if record.punctuation_ok else 'no'}",
                    f"- text: {record.text}",
                    "",
                ]
            )
    else:
        lines.extend(["- no body citations detected", ""])

    lines.extend(["## Findings"])
    finding_groups = [
        ("order failures", result.order_failures),
        ("forbidden surface hits", result.forbidden_surface_hits),
        ("non-superscript hits", result.non_superscript_hits),
        ("non-hyperlink hits", result.non_hyperlink_hits),
        ("wrong hyperlink target hits", result.wrong_hyperlink_target_hits),
        ("visual-style failures", result.visual_style_failures),
        ("multi-marker sentence hits", result.multi_marker_sentence_hits),
        ("multi-number marker hits", result.multi_number_marker_hits),
        ("punctuation failures", result.punctuation_failures),
        ("bibliography numbering conflicts", result.bibliography_numbering_conflicts),
        ("visible anchor leaks", result.visible_anchor_leaks),
    ]
    any_findings = False
    for label, items in finding_groups:
        if not items:
            continue
        any_findings = True
        lines.append(f"- {label}:")
        for item in items:
            lines.append(f"  - {item}")
    if result.missing_bibliography_numbers:
        any_findings = True
        lines.append(f"- uncited bibliography items: {result.missing_bibliography_numbers}")
    if result.extra_body_citation_numbers:
        any_findings = True
        lines.append(f"- out-of-range body citations: {result.extra_body_citation_numbers}")
    if not any_findings:
        lines.append("- none")
    return "\n".join(lines).rstrip() + "\n"


def run_cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True, help="Thesis DOCX path")
    parser.add_argument("--report", help="Write markdown audit report to this path")
    parser.add_argument(
        "--allow-uncited-bibliography",
        action="store_true",
        help="Do not fail when bibliography entries are present but never cited in the body",
    )
    parser.add_argument("--json", action="store_true", help="Print JSON instead of summary text")
    args = parser.parse_args(argv)

    docx_path = Path(args.docx).resolve()
    result = audit_docx(docx_path, allow_uncited_bibliography=args.allow_uncited_bibliography)

    if args.report:
        report_path = Path(args.report).resolve()
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(build_report(result, docx_path), encoding="utf-8")

    if args.json:
        print(json.dumps(asdict(result), ensure_ascii=False, indent=2))
    else:
        print(f"Body citation audit: {'PASS' if result.passed else 'FAIL'}")
        print(f"- first appearance chain: {result.first_appearance_chain}")
        print(f"- expected chain: {result.expected_chain}")
        print(
            f"- missing bibliography numbers: {result.missing_bibliography_numbers if result.missing_bibliography_numbers else 'none'}"
        )
        print(
            f"- extra body citation numbers: {result.extra_body_citation_numbers if result.extra_body_citation_numbers else 'none'}"
        )
        print(
            f"- bibliography numbering conflicts: {len(result.bibliography_numbering_conflicts)}"
        )
        print(f"- error codes: {', '.join(result.error_codes) if result.error_codes else 'none'}")
        if args.report:
            print(f"- report: {Path(args.report).resolve()}")

    return 0 if result.passed else 1


def main() -> int:
    try:
        return run_cli(sys.argv[1:])
    except Exception as exc:  # pragma: no cover - CLI safety
        print(f"Body citation audit failed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
