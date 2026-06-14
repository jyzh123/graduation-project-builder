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
BIBLIOGRAPHY_VISIBLE_NUMBER_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\d+\.)")
BIBLIOGRAPHY_VISIBLE_LABEL_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\d+[\.\u3001])")
MIN_BIBLIOGRAPHY_ENTRY_CONTENT_CHARS = 8

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
    "bibliography_empty_entry": "BIBLIOGRAPHY_ENTRY_CONTENT_MISSING",
    "anchor_leak": "BODY_CITATION_ANCHOR_LEAK",
    "citation_distribution": "BODY_CITATION_DISTRIBUTION_ABNORMAL",
    "bibliography_visible_label": "BIBLIOGRAPHY_VISIBLE_LABEL_MISSING",
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


@dataclass(frozen=True)
class TextUnit:
    char: str
    run: ET.Element
    hyperlink_anchor: str


@dataclass(frozen=True)
class CitationMarkerOccurrence:
    number: int
    text: str
    units: tuple[TextUnit, ...]

    @property
    def common_hyperlink_anchor(self) -> str:
        anchors = {unit.hyperlink_anchor for unit in self.units}
        if len(anchors) == 1:
            return next(iter(anchors))
        return ""


@dataclass
class AuditResult:
    schema: str
    generator: str
    docx_path: str
    docx_sha256: str
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
    bibliography_empty_entry_hits: list[str]
    citation_distribution_failures: list[str]
    bibliography_visible_label_failures: list[str]
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


def record_compact_text(record: ParagraphRecord) -> str:
    return re.sub(r"\s+", "", record.text or "")


def is_bibliography_heading(record: ParagraphRecord) -> bool:
    return record_compact_text(record) == BIBLIO_HEADING_NORMALIZED


def is_tail_heading(record: ParagraphRecord) -> bool:
    return record_compact_text(record).lower() in TAIL_HEADINGS_NORMALIZED


def paragraph_has_citation_bookmark(paragraph: ET.Element) -> bool:
    for bookmark in paragraph.findall(".//w:bookmarkStart", NS):
        name = bookmark.attrib.get(f"{W}name", "")
        if re.fullmatch(r"cite_ref_\d+", name or ""):
            return True
    return False


def is_bibliography_entry(record: ParagraphRecord) -> bool:
    if not record.text:
        return False
    return bool(
        record.has_numbering
        or BIBLIOGRAPHY_VISIBLE_NUMBER_RE.search(record.text)
        or paragraph_has_citation_bookmark(record.xml)
    )


def bibliography_entry_content_text(record: ParagraphRecord) -> str:
    text = BIBLIOGRAPHY_VISIBLE_LABEL_RE.sub("", record.text or "", count=1)
    text = re.sub(r"^\s*[\.\u3001\u3002\]\)）:：,，;；-]+", "", text)
    return re.sub(r"\s+", "", text)


def is_bibliography_number_only_record(record: ParagraphRecord) -> bool:
    if not record.text:
        return False
    has_entry_signal = bool(
        record.has_numbering
        or BIBLIOGRAPHY_VISIBLE_LABEL_RE.search(record.text)
        or paragraph_has_citation_bookmark(record.xml)
    )
    if not has_entry_signal:
        return False
    return len(bibliography_entry_content_text(record)) < MIN_BIBLIOGRAPHY_ENTRY_CONTENT_CHARS


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
    heading_indexes = [record.index for record in paragraphs if is_bibliography_heading(record)]
    if not heading_indexes:
        raise ValueError(f"Could not find bibliography heading '{BIBLIO_HEADING}'.")

    for heading_index in heading_indexes:
        cursor = heading_index + 1
        active_heading_index = heading_index
        while cursor < len(paragraphs):
            record = paragraphs[cursor]
            if not record.text:
                cursor += 1
                continue
            if is_bibliography_heading(record):
                active_heading_index = record.index
                cursor += 1
                continue
            if is_bibliography_entry(record):
                end = len(paragraphs)
                for tail in paragraphs[cursor + 1 :]:
                    if is_tail_heading(tail):
                        end = tail.index
                        break
                    if looks_like_heading(tail.text) and not tail.has_numbering and not is_bibliography_entry(tail):
                        end = tail.index
                        break
                return active_heading_index, end
            if looks_like_heading(record.text) and not record.has_numbering:
                break
            cursor += 1

    raise ValueError(f"Could not find bibliography entries after heading '{BIBLIO_HEADING}'.")


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


def run_text(run: ET.Element) -> str:
    return "".join((node.text or "") for node in run.iterfind(".//w:t", NS))


def paragraph_text_units(paragraph: ParagraphRecord) -> list[TextUnit]:
    units: list[TextUnit] = []

    def walk(element: ET.Element, hyperlink_anchor: str = "") -> None:
        if element.tag == f"{W}hyperlink":
            hyperlink_anchor = element.attrib.get(f"{W}anchor", "")
        if element.tag == f"{W}r":
            for text_node in element.iterfind(".//w:t", NS):
                for char in text_node.text or "":
                    units.append(TextUnit(char=char, run=element, hyperlink_anchor=hyperlink_anchor))
            return
        for child in list(element):
            walk(child, hyperlink_anchor)

    walk(paragraph.xml)
    return units


def citation_marker_occurrences(paragraph: ParagraphRecord, number: int | None = None) -> list[CitationMarkerOccurrence]:
    units = paragraph_text_units(paragraph)
    visible_text = "".join(unit.char for unit in units)
    occurrences: list[CitationMarkerOccurrence] = []
    for match in re.finditer(r"\[(\d+)\]", visible_text):
        marker_number = int(match.group(1))
        if number is not None and marker_number != number:
            continue
        occurrences.append(
            CitationMarkerOccurrence(
                number=marker_number,
                text=match.group(0),
                units=tuple(units[match.start() : match.end()]),
            )
        )
    return occurrences


def occurrence_runs(occurrence: CitationMarkerOccurrence) -> list[ET.Element]:
    runs: list[ET.Element] = []
    seen: set[int] = set()
    for unit in occurrence.units:
        marker = id(unit.run)
        if marker in seen:
            continue
        seen.add(marker)
        runs.append(unit.run)
    return runs


def run_is_superscript(run: ET.Element) -> bool:
    vert = run.find("./w:rPr/w:vertAlign", NS)
    return vert is not None and vert.attrib.get(f"{W}val") == "superscript"


def occurrence_uses_marker_only_runs(occurrence: CitationMarkerOccurrence) -> bool:
    for run in occurrence_runs(occurrence):
        compact_run_text = re.sub(r"\s+", "", run_text(run))
        if not compact_run_text:
            continue
        if compact_run_text not in occurrence.text:
            return False
    return True


def citation_marker_count(paragraph: ParagraphRecord, number: int) -> int:
    return len(citation_marker_occurrences(paragraph, number))


def run_has_superscript_citation(paragraph: ParagraphRecord, number: int) -> bool:
    occurrences = citation_marker_occurrences(paragraph, number)
    if not occurrences:
        return False
    for occurrence in occurrences:
        if not occurrence_uses_marker_only_runs(occurrence):
            return False
        if any(not run_is_superscript(run) for run in occurrence_runs(occurrence)):
            return False
    return True


def run_has_hyperlink_citation(paragraph: ParagraphRecord, number: int) -> bool:
    occurrences = citation_marker_occurrences(paragraph, number)
    if not occurrences:
        return False
    return all(bool(occurrence.common_hyperlink_anchor) for occurrence in occurrences)


def run_has_bibliography_target(
    paragraph: ParagraphRecord,
    number: int,
    *,
    bookmark_locations: dict[str, int],
    bibliography_start: int,
    bibliography_end: int,
) -> bool:
    anchor = f"cite_ref_{number}"
    target_index = bookmark_locations.get(anchor)
    if target_index is None or not (bibliography_start < target_index < bibliography_end):
        return False
    occurrences = citation_marker_occurrences(paragraph, number)
    if not occurrences:
        return False
    return all(occurrence.common_hyperlink_anchor == anchor for occurrence in occurrences)


def citation_runs(paragraph: ParagraphRecord, number: int) -> list[ET.Element]:
    matched: list[ET.Element] = []
    seen: set[int] = set()
    for occurrence in citation_marker_occurrences(paragraph, number):
        for run in occurrence_runs(occurrence):
            marker = id(run)
            if marker in seen:
                continue
            seen.add(marker)
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


NUMERIC_CITATION_MARKER_RE = re.compile(
    r"\[\s*\d+(?:\s*(?:,|，|;|；|-|–|—|~|～)\s*\d+)*\s*\]"
)


def sentence_marker_rule_violations(text: str) -> tuple[list[str], list[str]]:
    multi_marker_hits: list[str] = []
    multi_number_hits: list[str] = []
    for sentence in split_sentences(text):
        markers = [
            match.group(0)[1:-1]
            for match in NUMERIC_CITATION_MARKER_RE.finditer(sentence)
        ]
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
    return [record for record in paragraphs[start + 1 : end] if is_bibliography_entry(record)]


def find_bibliography_numbering_conflicts(bibliography: list[ParagraphRecord]) -> list[str]:
    conflicts: list[str] = []
    for record in bibliography:
        if record.has_numbering and BIBLIOGRAPHY_VISIBLE_NUMBER_RE.search(record.text):
            conflicts.append(
                f"paragraph {record.index + 1} has both Word numbering and a visible manual bibliography number: {record.text}"
            )
    return conflicts


def find_bibliography_empty_entry_hits(bibliography: list[ParagraphRecord]) -> list[str]:
    hits: list[str] = []
    for record in bibliography:
        if is_bibliography_number_only_record(record):
            content_len = len(bibliography_entry_content_text(record))
            hits.append(
                f"paragraph {record.index + 1} bibliography entry has only a label/bookmark "
                f"or too little substantive content ({content_len} chars): {record.text}"
            )
    return hits


def find_bibliography_visible_label_failures(bibliography: list[ParagraphRecord]) -> list[str]:
    failures: list[str] = []
    for offset, record in enumerate(bibliography, start=1):
        # Word automatic numbering renders the label outside w:t, so a numbered
        # bibliography paragraph does not need a manual [n] text prefix.
        if record.has_numbering:
            continue
        if not BIBLIOGRAPHY_VISIBLE_LABEL_RE.match(record.text):
            failures.append(
                f"bibliography entry {offset} at paragraph {record.index + 1} lacks a visible reference label: {record.text}"
            )
    return failures


VISIBLE_ANCHOR_LEAK_RE = re.compile(
    r"\b(?:cite_ref|ref_anchor|bookmark)_(?:[A-Za-z0-9][A-Za-z0-9_.:-]*)?\b"
    r"|\b(?:cite_ref|ref_anchor|bookmark)_",
    re.IGNORECASE,
)
FIELD_HYPERLINK_ANCHOR_RE = re.compile(r'HYPERLINK\s+\\l\s+"([^"]+)"', re.IGNORECASE)


def field_hyperlink_anchors(paragraph: ParagraphRecord) -> list[str]:
    anchors: list[str] = []
    for instr in paragraph.xml.findall(".//w:instrText", NS):
        match = FIELD_HYPERLINK_ANCHOR_RE.search(instr.text or "")
        if match:
            anchors.append(match.group(1))
    return anchors


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
    bibliography_empty_entry_hits = find_bibliography_empty_entry_hits(bibliography)
    bibliography_visible_label_failures = find_bibliography_visible_label_failures(bibliography)
    citation_number_counts: dict[int, int] = {}
    citation_distribution_failures: list[str] = []
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
            citation_number_counts[number] = citation_number_counts.get(number, 0) + citation_marker_count(record, number)
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
    if bibliography_empty_entry_hits:
        error_codes.add(ERROR_CODES["bibliography_empty_entry"])
    if bibliography_visible_label_failures:
        error_codes.add(ERROR_CODES["bibliography_visible_label"])
    total_citation_markers = sum(citation_number_counts.values())
    if total_citation_markers and citation_number_counts:
        dominant_number, dominant_count = max(citation_number_counts.items(), key=lambda item: item[1])
        dominant_ratio = dominant_count / max(total_citation_markers, 1)
        if bibliography_count >= 20 and dominant_count >= 20 and dominant_ratio > 0.35:
            citation_distribution_failures.append(
                f"citation [{dominant_number}] appears {dominant_count}/{total_citation_markers} times; dominant ratio {dominant_ratio:.2%} indicates filler or stale citation distribution"
            )
    if citation_distribution_failures:
        error_codes.add(ERROR_CODES["citation_distribution"])
    if visible_anchor_leaks:
        error_codes.add(ERROR_CODES["anchor_leak"])

    return AuditResult(
        schema="graduation-project-builder.thesis-citation-audit.v2",
        generator="scripts/audit_thesis_citations.py",
        docx_path=str(docx_path),
        docx_sha256=sha256_file(docx_path),
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
        bibliography_empty_entry_hits=bibliography_empty_entry_hits,
        citation_distribution_failures=citation_distribution_failures,
        bibliography_visible_label_failures=bibliography_visible_label_failures,
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
        f"- bibliography empty/content-missing entries: {len(result.bibliography_empty_entry_hits)}",
        f"- citation distribution failures: {len(result.citation_distribution_failures)}",
        f"- bibliography visible-label failures: {len(result.bibliography_visible_label_failures)}",
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
        ("bibliography empty/content-missing entries", result.bibliography_empty_entry_hits),
        ("citation distribution failures", result.citation_distribution_failures),
        ("bibliography visible-label failures", result.bibliography_visible_label_failures),
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
        print(
            f"- bibliography empty/content-missing entries: {len(result.bibliography_empty_entry_hits)}"
        )
        print(f"- citation distribution failures: {len(result.citation_distribution_failures)}")
        print(f"- bibliography visible-label failures: {len(result.bibliography_visible_label_failures)}")
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
