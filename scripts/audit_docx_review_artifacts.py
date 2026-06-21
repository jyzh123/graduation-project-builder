#!/usr/bin/env python3
"""Audit DOCX review artifacts and body citation run preservation."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from .audit_thesis_citations import (
        audit_docx as audit_body_citations,
        citation_is_before_punctuation,
        citation_marker_count,
        citation_runs,
        citation_visual_style_ok,
        classify_surface,
        extract_citation_numbers,
        find_bibliography_range,
        iter_paragraphs,
        run_has_hyperlink_citation,
        run_has_superscript_citation,
    )
    from .audit_thesis_comment_resolution import validate_comment_resolution_ledger
except ImportError:
    from audit_thesis_citations import (
        audit_docx as audit_body_citations,
        citation_is_before_punctuation,
        citation_marker_count,
        citation_runs,
        citation_visual_style_ok,
        classify_surface,
        extract_citation_numbers,
        find_bibliography_range,
        iter_paragraphs,
        run_has_hyperlink_citation,
        run_has_superscript_citation,
    )
    from audit_thesis_comment_resolution import validate_comment_resolution_ledger


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_text(text: str) -> str:
    return " ".join((text or "").split()).strip()


def digest_text(text: str) -> str:
    return hashlib.sha256(normalize_text(text).encode("utf-8")).hexdigest()


def parse_report_values(text: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("- ") or ":" not in stripped:
            continue
        key, _, value = stripped.partition(":")
        values[key + ":"] = value.strip()
    return values


def read_zip_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    if name not in zf.namelist():
        return None
    return ET.fromstring(zf.read(name))


@dataclass
class PackageCounts:
    comments_xml: int
    comments_extended_xml: int
    people_xml: int
    comment_anchors: int
    tracked_changes: int
    bookmarks: int
    fields: int
    hyperlinks: int


@dataclass
class CommentRecord:
    comment_id: str
    text_digest: str


@dataclass
class CitationRecord:
    paragraph_index: int
    para_id: str
    paragraph_text_digest: str
    number: int
    marker_count: int
    superscript_ok: bool
    hyperlink_ok: bool
    visual_style_ok: bool
    punctuation_ok: bool
    surface: str


@dataclass
class ReviewArtifactSnapshot:
    docx_path: str
    docx_sha256: str
    package_counts: PackageCounts
    comment_part_names: list[str]
    comments: list[CommentRecord]
    comment_anchor_ids: list[str]
    tracked_change_ids: list[str]
    bookmark_names: list[str]
    field_instr_digests: list[str]
    hyperlink_anchors: list[str]


@dataclass
class CitationSnapshot:
    docx_path: str
    docx_sha256: str
    records: list[CitationRecord]
    citation_numbers: list[int]


@dataclass
class BookmarkDisposition:
    path: Path
    schema: str
    source_docx_path: str
    source_docx_sha256: str
    final_docx_path: str
    final_docx_sha256: str
    allowed_missing_bookmarks: set[str]
    allowed_missing_field_instr_digests: set[str]
    allowed_missing_hyperlinks: set[str]
    allow_source_citation_nonpreservation: bool
    scope: str
    verdict: str


EMPTY_PARAGRAPH_BOOKMARK_DISPOSITION_SCHEMA = "graduation-project-builder.empty-paragraph-bookmark-disposition.v1"
NEW_THESIS_SOURCE_ARTIFACT_DISPOSITION_SCHEMA = "graduation-project-builder.new-thesis-source-artifact-disposition.v1"
STRICT_CITATION_PRESERVATION_SCOPE = "local-surface-preservation"
WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE = "whole-rebuild-chain-integrity"
APPROVED_NONPRESERVATION_CITATION_SCOPE = "approved-non-preservation"
ALLOWED_CITATION_PRESERVATION_SCOPES = {
    STRICT_CITATION_PRESERVATION_SCOPE,
    WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE,
    APPROVED_NONPRESERVATION_CITATION_SCOPE,
}


def is_format_template_path(path: str | Path) -> bool:
    name = Path(path).name.lower()
    return any(token in name for token in ("\u683c\u5f0f", "\u6a21\u677f", "format", "template"))


STORY_PART_RE = re.compile(r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$")
TRACKED_CHANGE_TAGS = {
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "cellIns",
    "cellDel",
    "cellMerge",
    "numberingChange",
}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def is_tracked_change_element(tag: str) -> bool:
    name = local_name(tag)
    return name in TRACKED_CHANGE_TAGS or name.endswith("PrChange")


def collect_review_artifacts(docx_path: Path) -> ReviewArtifactSnapshot:
    with zipfile.ZipFile(docx_path, "r") as zf:
        part_names = set(zf.namelist())
        story_part_names = sorted(name for name in part_names if STORY_PART_RE.match(name))
        comment_part_names = sorted(
            name for name in part_names
            if name.startswith("word/comments") and name.endswith(".xml")
        )
        comments = read_zip_xml(zf, "word/comments.xml")
        comments_extended = read_zip_xml(zf, "word/commentsExtended.xml")
        people = read_zip_xml(zf, "word/people.xml")

        comment_records: list[CommentRecord] = []
        if comments is not None:
            for comment in comments.findall(f"{W}comment"):
                comment_id = comment.attrib.get(f"{W}id", "")
                comment_records.append(
                    CommentRecord(
                        comment_id=comment_id,
                        text_digest=digest_text("".join(comment.itertext())),
                    )
                )

        comment_anchor_ids: list[str] = []
        tracked_change_ids: list[str] = []
        bookmark_names: list[str] = []
        field_instr_digests: list[str] = []
        hyperlink_anchors: list[str] = []
        comment_anchor_count = 0
        tracked_change_count = 0
        bookmark_count = 0
        field_count = 0
        hyperlink_count = 0

        for part_name in story_part_names:
            document = read_zip_xml(zf, part_name)
            if document is None:
                continue
            for element in document.iter():
                tag = element.tag
                if tag == f"{W}commentRangeStart" or tag == f"{W}commentRangeEnd" or tag == f"{W}commentReference":
                    comment_anchor_count += 1
                    anchor_id = element.attrib.get(f"{W}id", "")
                    if anchor_id:
                        comment_anchor_ids.append(f"{part_name}:{local_name(tag)}:{anchor_id}")
                elif is_tracked_change_element(tag):
                    tracked_change_count += 1
                    change_id = element.attrib.get(f"{W}id", "")
                    if not change_id:
                        change_id = digest_text(ET.tostring(element, encoding="unicode"))
                    tracked_change_ids.append(f"{part_name}:{local_name(tag)}:{change_id}")
                elif tag == f"{W}bookmarkStart":
                    bookmark_count += 1
                    name = element.attrib.get(f"{W}name", "")
                    if name:
                        bookmark_names.append(f"{part_name}:{name}")
                elif tag == f"{W}bookmarkEnd":
                    bookmark_count += 1
                    bid = element.attrib.get(f"{W}id", "")
                    if bid:
                        bookmark_names.append(f"{part_name}:bookmark-id:{bid}")
                elif tag == f"{W}fldChar":
                    field_count += 1
                elif tag == f"{W}instrText":
                    field_instr_digests.append(f"{part_name}:{digest_text(element.text or '')}")
                elif tag == f"{W}hyperlink":
                    hyperlink_count += 1
                    anchor = element.attrib.get(f"{W}anchor", "")
                    rid = element.attrib.get(f"{R}id", "")
                    text = digest_text("".join(element.itertext()))
                    if anchor:
                        hyperlink_anchors.append(f"{part_name}:anchor:{anchor}:{text}")
                    elif rid:
                        hyperlink_anchors.append(f"{part_name}:rid:{rid}:{text}")
                    else:
                        hyperlink_anchors.append(f"{part_name}:text:{text}")

        parts = PackageCounts(
            comments_xml=1 if comments is not None else 0,
            comments_extended_xml=1 if comments_extended is not None else 0,
            people_xml=1 if people is not None else 0,
            comment_anchors=comment_anchor_count,
            tracked_changes=tracked_change_count,
            bookmarks=bookmark_count,
            fields=field_count,
            hyperlinks=hyperlink_count,
        )

    return ReviewArtifactSnapshot(
        docx_path=str(docx_path),
        docx_sha256=sha256_file(docx_path),
        package_counts=parts,
        comment_part_names=comment_part_names,
        comments=comment_records,
        comment_anchor_ids=sorted(set(comment_anchor_ids)),
        tracked_change_ids=sorted(set(tracked_change_ids)),
        bookmark_names=sorted(set(bookmark_names)),
        field_instr_digests=sorted(set(field_instr_digests)),
        hyperlink_anchors=sorted(set(hyperlink_anchors)),
    )


def collect_citation_snapshot(docx_path: Path) -> CitationSnapshot:
    paragraphs = iter_paragraphs(docx_path)
    records: list[CitationRecord] = []
    try:
        start, _ = find_bibliography_range(paragraphs)
    except ValueError:
        start = len(paragraphs)
    citation_numbers: list[int] = []
    for record in paragraphs[:start]:
        numbers = extract_citation_numbers(record.text)
        if not numbers:
            continue
        surface = classify_surface(record)
        for number in numbers:
            citation_numbers.append(number)
            records.append(
                CitationRecord(
                    paragraph_index=record.index + 1,
                    para_id=record.para_id,
                    paragraph_text_digest=digest_text(record.text),
                    number=number,
                    marker_count=citation_marker_count(record, number),
                    superscript_ok=run_has_superscript_citation(record, number),
                    hyperlink_ok=run_has_hyperlink_citation(record, number),
                    visual_style_ok=citation_visual_style_ok(record, number),
                    punctuation_ok=citation_is_before_punctuation(record.text, number),
                    surface=surface,
                )
            )
    return CitationSnapshot(
        docx_path=str(docx_path),
        docx_sha256=sha256_file(docx_path),
        records=records,
        citation_numbers=citation_numbers,
    )


def _resolve_disposition_path(value: object, base: Path) -> Path | None:
    text = str(value or "").strip()
    if not text or text.lower() in {"none", "n/a", "not-applicable"}:
        return None
    path = Path(text)
    if not path.is_absolute():
        path = (base.parent / path).resolve()
    return path


def load_empty_paragraph_bookmark_disposition(path: Path) -> tuple[BookmarkDisposition | None, list[str]]:
    issues: list[str] = []
    if not path.exists():
        return None, [f"controlled bookmark disposition missing: {path}"]
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return None, [f"controlled bookmark disposition unreadable: {path} ({exc})"]
    if not isinstance(payload, dict):
        return None, [f"controlled bookmark disposition root must be an object: {path}"]

    schema = str(payload.get("schema", "")).strip()
    if schema not in {EMPTY_PARAGRAPH_BOOKMARK_DISPOSITION_SCHEMA, NEW_THESIS_SOURCE_ARTIFACT_DISPOSITION_SCHEMA}:
        issues.append("controlled bookmark disposition schema is not recognized")
    verdict = str(payload.get("verdict", "")).strip().lower()
    if "pass" not in verdict or any(token in verdict for token in ("fail", "blocked", "missing", "pending")):
        issues.append("controlled bookmark disposition verdict is not pass")

    allowed_missing = {
        str(item).strip()
        for item in payload.get("allowed_missing_bookmarks", [])
        if str(item).strip()
    }
    allowed_missing_fields = {
        str(item).strip()
        for item in payload.get("allowed_missing_field_instr_digests", [])
        if str(item).strip()
    }
    allowed_missing_hyperlinks = {
        str(item).strip()
        for item in payload.get("allowed_missing_hyperlinks", [])
        if str(item).strip()
    }
    scope = str(payload.get("scope", "")).strip()
    allow_source_citation_nonpreservation = payload.get("allow_source_citation_nonpreservation") is True

    if schema == EMPTY_PARAGRAPH_BOOKMARK_DISPOSITION_SCHEMA:
        truthy_fields = ("empty_paragraph_verified", "visible_text_absent")
        false_fields = (
            "contains_image_or_table",
            "contains_section_or_page_break",
            "contains_comment_anchor",
            "contains_tracked_change",
            "contains_field_host",
            "contains_hyperlink",
            "contains_citation_marker",
        )
        for key in truthy_fields:
            if payload.get(key) is not True:
                issues.append(f"controlled bookmark disposition {key} must be true")
        for key in false_fields:
            if payload.get(key) is not False:
                issues.append(f"controlled bookmark disposition {key} must be false")
        if allowed_missing_fields:
            issues.append("empty-paragraph bookmark disposition cannot authorize field host loss")
        if allowed_missing_hyperlinks:
            issues.append("empty-paragraph bookmark disposition cannot authorize hyperlink host loss")
        if allow_source_citation_nonpreservation:
            issues.append("empty-paragraph bookmark disposition cannot authorize source citation non-preservation")
        if not scope:
            scope = "empty-paragraph-bookmark"
    elif schema == NEW_THESIS_SOURCE_ARTIFACT_DISPOSITION_SCHEMA:
        if payload.get("selected_thesis_workflow") != "new-thesis-production":
            issues.append("new-thesis source artifact disposition requires selected_thesis_workflow=new-thesis-production")
        if payload.get("rebuild_class") != "new-thesis-production":
            issues.append("new-thesis source artifact disposition requires rebuild_class=new-thesis-production")
        if payload.get("source_subject_replaced") is not True:
            issues.append("new-thesis source artifact disposition must confirm source_subject_replaced")
        if payload.get("final_citation_chain_audit_required") is not True:
            issues.append("new-thesis source artifact disposition must require final citation chain audit")
        if payload.get("reference_entries_full_content_required") is not True:
            issues.append("new-thesis source artifact disposition must require full-content references")
        if not allow_source_citation_nonpreservation:
            issues.append("new-thesis source artifact disposition must allow source citation non-preservation explicitly")
        if not scope:
            scope = "new-thesis-source-artifact-replacement"

    if not allowed_missing and not allowed_missing_fields and not allowed_missing_hyperlinks and not allow_source_citation_nonpreservation:
        issues.append("controlled bookmark disposition allowed_missing_bookmarks is empty")

    source_docx_path = str(payload.get("source_docx_path", "")).strip()
    final_docx_path = str(payload.get("final_docx_path", "")).strip()
    source_docx_sha256 = str(payload.get("source_docx_sha256", "")).strip().lower()
    final_docx_sha256 = str(payload.get("final_docx_sha256", "")).strip().lower()
    for label, value in (
        ("source_docx_path", source_docx_path),
        ("final_docx_path", final_docx_path),
        ("source_docx_sha256", source_docx_sha256),
        ("final_docx_sha256", final_docx_sha256),
    ):
        if not value:
            issues.append(f"controlled bookmark disposition {label} is missing")

    return (
        BookmarkDisposition(
            path=path,
            schema=schema,
            source_docx_path=source_docx_path,
            source_docx_sha256=source_docx_sha256,
            final_docx_path=final_docx_path,
            final_docx_sha256=final_docx_sha256,
            allowed_missing_bookmarks=allowed_missing,
            allowed_missing_field_instr_digests=allowed_missing_fields,
            allowed_missing_hyperlinks=allowed_missing_hyperlinks,
            allow_source_citation_nonpreservation=allow_source_citation_nonpreservation,
            scope=scope,
            verdict=verdict,
        ),
        issues,
    )


def disposition_matches_docs(
    disposition: BookmarkDisposition | None,
    source: ReviewArtifactSnapshot,
    final: ReviewArtifactSnapshot,
) -> list[str]:
    if disposition is None:
        return ["controlled bookmark disposition is missing"]
    issues: list[str] = []
    if Path(disposition.source_docx_path).resolve() != Path(source.docx_path).resolve():
        issues.append("controlled bookmark disposition source_docx_path does not match source DOCX")
    if Path(disposition.final_docx_path).resolve() != Path(final.docx_path).resolve():
        issues.append("controlled bookmark disposition final_docx_path does not match final DOCX")
    if disposition.source_docx_sha256.lower() != source.docx_sha256.lower():
        issues.append("controlled bookmark disposition source_docx_sha256 does not match source DOCX")
    if disposition.final_docx_sha256.lower() != final.docx_sha256.lower():
        issues.append("controlled bookmark disposition final_docx_sha256 does not match final DOCX")
    return issues


def compare_review_artifacts(
    source: ReviewArtifactSnapshot,
    final: ReviewArtifactSnapshot,
    *,
    allow_approved_comment_disposal: bool = False,
    controlled_bookmark_disposition: BookmarkDisposition | None = None,
) -> list[str]:
    issues: list[str] = []
    source_is_format_template = is_format_template_path(source.docx_path)
    if source.docx_path == final.docx_path:
        issues.append("source and final review DOCX paths must differ")
    if source.package_counts.comments_xml > 0 and final.package_counts.comments_xml == 0:
        if not allow_approved_comment_disposal:
            issues.append("source comments.xml was stripped from final DOCX")
    if source.package_counts.comments_xml == 0 and final.package_counts.comments_xml > 0:
        issues.append("final DOCX gained comments.xml without source comments evidence")
    if source.package_counts.comments_extended_xml > 0 and final.package_counts.comments_extended_xml == 0:
        if not allow_approved_comment_disposal:
            issues.append("source commentsExtended.xml was stripped from final DOCX")
    if source.package_counts.people_xml > 0 and final.package_counts.people_xml == 0:
        if not allow_approved_comment_disposal:
            issues.append("source people.xml was stripped from final DOCX")
    missing_comment_parts = sorted(set(source.comment_part_names) - set(final.comment_part_names))
    if missing_comment_parts and not allow_approved_comment_disposal:
        issues.append(f"source comment-related parts were stripped from final DOCX: {missing_comment_parts[:8]}")

    if source.comments:
        source_comment_ids = {item.comment_id for item in source.comments}
        final_comment_map = {item.comment_id: item.text_digest for item in final.comments}
        missing_comment_ids = sorted(source_comment_ids - set(final_comment_map))
        if missing_comment_ids and not allow_approved_comment_disposal:
            issues.append(f"source comment ids missing from final DOCX: {missing_comment_ids}")
        for item in source.comments:
            if final_comment_map.get(item.comment_id, "") != item.text_digest and not allow_approved_comment_disposal:
                issues.append(f"source comment text digest changed for comment id {item.comment_id}")
    if source.package_counts.comment_anchors > 0:
        if final.package_counts.comment_anchors < source.package_counts.comment_anchors and not allow_approved_comment_disposal:
            issues.append(
                "source comment anchor count decreased in final DOCX "
                f"(source={source.package_counts.comment_anchors}, final={final.package_counts.comment_anchors})"
            )
        missing_anchor_ids = sorted(set(source.comment_anchor_ids) - set(final.comment_anchor_ids))
        if missing_anchor_ids and not allow_approved_comment_disposal:
            issues.append(f"source comment anchor ids missing from final DOCX: {missing_anchor_ids[:12]}")

    if source.package_counts.tracked_changes > 0:
        if final.package_counts.tracked_changes < source.package_counts.tracked_changes:
            issues.append("tracked change marks were stripped from final DOCX")
        missing_change_ids = sorted(set(source.tracked_change_ids) - set(final.tracked_change_ids))
        if missing_change_ids:
            issues.append(f"source tracked change ids missing from final DOCX: {missing_change_ids}")

    if source.package_counts.bookmarks > 0 and not source_is_format_template:
        missing_bookmarks = sorted(set(source.bookmark_names) - set(final.bookmark_names))
        if missing_bookmarks:
            disposition_issues = disposition_matches_docs(controlled_bookmark_disposition, source, final)
            if disposition_issues:
                issues.append(f"source bookmark anchors missing from final DOCX: {missing_bookmarks[:8]}")
                issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in disposition_issues)
            else:
                allowed = controlled_bookmark_disposition.allowed_missing_bookmarks if controlled_bookmark_disposition else set()
                unauthorized = sorted(set(missing_bookmarks) - allowed)
                stale_allowed = sorted(allowed - set(missing_bookmarks))
                if unauthorized:
                    issues.append(f"source bookmark anchors missing from final DOCX: {unauthorized[:8]}")
                if stale_allowed:
                    issues.append(f"controlled bookmark disposition names bookmarks not missing from final DOCX: {stale_allowed[:8]}")

    if source.package_counts.fields > 0 and not source_is_format_template:
        missing_fields = sorted(set(source.field_instr_digests) - set(final.field_instr_digests))
        if missing_fields:
            disposition_issues = disposition_matches_docs(controlled_bookmark_disposition, source, final)
            allowed = (
                controlled_bookmark_disposition.allowed_missing_field_instr_digests
                if controlled_bookmark_disposition is not None
                else set()
            )
            unauthorized = sorted(set(missing_fields) - allowed)
            stale_allowed = sorted(allowed - set(missing_fields))
            if disposition_issues or unauthorized:
                issues.append("source field hosts missing from final DOCX")
                issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in disposition_issues)
            if stale_allowed:
                issues.append(f"controlled bookmark disposition names field hosts not missing from final DOCX: {stale_allowed[:8]}")

    if source.package_counts.hyperlinks > 0 and not source_is_format_template:
        missing_hyperlinks = sorted(set(source.hyperlink_anchors) - set(final.hyperlink_anchors))
        if missing_hyperlinks:
            disposition_issues = disposition_matches_docs(controlled_bookmark_disposition, source, final)
            allowed = (
                controlled_bookmark_disposition.allowed_missing_hyperlinks
                if controlled_bookmark_disposition is not None
                else set()
            )
            unauthorized = sorted(set(missing_hyperlinks) - allowed)
            stale_allowed = sorted(allowed - set(missing_hyperlinks))
            if disposition_issues or unauthorized:
                issues.append("source hyperlink hosts missing from final DOCX")
                issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in disposition_issues)
            if stale_allowed:
                issues.append(f"controlled bookmark disposition names hyperlinks not missing from final DOCX: {stale_allowed[:8]}")

    return issues


def _final_citation_record_issues(number: int, item: CitationRecord, *, source_has_hyperlink: bool) -> list[str]:
    issues: list[str] = []
    location = f"paragraph {item.paragraph_index}"
    if item.marker_count <= 0:
        issues.append(f"citation marker [{number}] missing concrete run occurrence in final DOCX at {location}")
    if not item.superscript_ok:
        issues.append(f"citation marker [{number}] lost superscript run state in final DOCX at {location}")
    if source_has_hyperlink and not item.hyperlink_ok:
        issues.append(f"citation marker [{number}] lost hyperlink host in final DOCX at {location}")
    if not item.visual_style_ok:
        issues.append(f"citation marker [{number}] lost clean citation visual style in final DOCX at {location}")
    if not item.punctuation_ok:
        issues.append(f"citation marker [{number}] lost punctuation-side placement in final DOCX at {location}")
    return issues


def _whole_rebuild_chain_integrity_issues(
    source: CitationSnapshot,
    final: CitationSnapshot,
    *,
    citation_preservation_scope: str,
    source_counts: dict[int, int],
    final_counts: dict[int, int],
    source_has_any_hyperlink: bool,
    source_is_format_template: bool = False,
) -> list[str]:
    issues: list[str] = []
    if citation_preservation_scope == WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE and source_has_any_hyperlink:
        issues.append(
            "whole-rebuild citation occurrence preservation scope requires source citation hyperlinks to be absent "
            "unless approved-non-preservation scope is recorded"
        )
    missing_numbers = sorted(set(source_counts) - set(final_counts))
    if (
        missing_numbers
        and not source_is_format_template
        and citation_preservation_scope != APPROVED_NONPRESERVATION_CITATION_SCOPE
    ):
        issues.append(f"whole-rebuild citation chain missing source citation numbers in final DOCX: {missing_numbers}")
    try:
        final_audit = audit_body_citations(Path(final.docx_path))
    except Exception as exc:
        issues.append(f"whole-rebuild final DOCX body citation audit could not run: {exc}")
    else:
        if not final_audit.passed:
            issues.append(
                "whole-rebuild final DOCX body citation audit failed: "
                f"{', '.join(final_audit.error_codes) if final_audit.error_codes else 'unknown'}"
            )
    return issues


def compare_citations(
    source: CitationSnapshot,
    final: CitationSnapshot,
    *,
    citation_preservation_scope: str = STRICT_CITATION_PRESERVATION_SCOPE,
) -> list[str]:
    issues: list[str] = []
    if source.docx_path == final.docx_path:
        issues.append("source and final citation DOCX paths must differ")
    if citation_preservation_scope not in ALLOWED_CITATION_PRESERVATION_SCOPES:
        issues.append(f"unrecognized citation occurrence preservation scope: {citation_preservation_scope}")
        citation_preservation_scope = STRICT_CITATION_PRESERVATION_SCOPE

    source_by_number: dict[int, list[CitationRecord]] = {}
    for record in source.records:
        source_by_number.setdefault(record.number, []).append(record)

    final_by_number: dict[int, list[CitationRecord]] = {}
    for record in final.records:
        final_by_number.setdefault(record.number, []).append(record)

    source_counts: dict[int, int] = {}
    for number in source.citation_numbers:
        source_counts[number] = source_counts.get(number, 0) + 1

    final_counts: dict[int, int] = {}
    for number in final.citation_numbers:
        final_counts[number] = final_counts.get(number, 0) + 1

    source_has_any_hyperlink = any(item.hyperlink_ok for item in source.records)
    if citation_preservation_scope in {
        WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE,
        APPROVED_NONPRESERVATION_CITATION_SCOPE,
    }:
        issues.extend(
            _whole_rebuild_chain_integrity_issues(
                source,
                final,
                citation_preservation_scope=citation_preservation_scope,
                source_counts=source_counts,
                final_counts=final_counts,
                source_has_any_hyperlink=source_has_any_hyperlink,
                source_is_format_template=is_format_template_path(source.docx_path),
            )
        )
        for number, final_records in final_by_number.items():
            source_has_hyperlink = any(item.hyperlink_ok for item in source_by_number.get(number, []))
            if citation_preservation_scope == APPROVED_NONPRESERVATION_CITATION_SCOPE:
                source_has_hyperlink = False
            for item in final_records:
                issues.extend(
                    _final_citation_record_issues(
                        number,
                        item,
                        source_has_hyperlink=source_has_hyperlink,
                    )
                )
        return issues

    for number, source_count in source_counts.items():
        final_count = final_counts.get(number, 0)
        if final_count < source_count:
            issues.append(f"citation marker [{number}] missing from final DOCX")
            continue
        final_records = final_by_number.get(number, [])
        if not final_records:
            issues.append(f"citation marker [{number}] missing from final DOCX")
            continue
        source_records = source_by_number.get(number, [])
        source_has_hyperlink = any(item.hyperlink_ok for item in source_records)
        for source_item in source_records:
            matching_final = [
                item
                for item in final_records
                if (
                    item.para_id
                    and source_item.para_id
                    and item.para_id == source_item.para_id
                )
                or (
                    not item.para_id
                    and not source_item.para_id
                    and item.paragraph_index == source_item.paragraph_index
                    and item.paragraph_text_digest == source_item.paragraph_text_digest
                )
            ]
            source_location = f"paragraph {source_item.paragraph_index}"
            if not matching_final:
                same_index_changed_host = [
                    item
                    for item in final_records
                    if (
                        not item.para_id
                        and not source_item.para_id
                        and item.paragraph_index == source_item.paragraph_index
                        and item.paragraph_text_digest != source_item.paragraph_text_digest
                    )
                ]
                if same_index_changed_host:
                    issues.append(
                        f"citation marker [{number}] source occurrence host text changed or moved in final DOCX at {source_location}"
                    )
                issues.append(
                    f"citation marker [{number}] source occurrence missing from final DOCX at {source_location}"
                )
                continue
            if not any(item.marker_count >= source_item.marker_count for item in matching_final):
                issues.append(
                    f"citation marker [{number}] occurrence count decreased at source {source_location}"
                )
            if source_item.superscript_ok and not all(item.superscript_ok for item in matching_final):
                issues.append(
                    f"citation marker [{number}] source occurrence lost superscript run state at {source_location}"
                )
            if source_item.hyperlink_ok and not all(item.hyperlink_ok for item in matching_final):
                issues.append(
                    f"citation marker [{number}] source occurrence lost hyperlink host at {source_location}"
                )
        if citation_preservation_scope != STRICT_CITATION_PRESERVATION_SCOPE:
            for item in final_records:
                issues.extend(
                    _final_citation_record_issues(
                        number,
                        item,
                        source_has_hyperlink=source_has_hyperlink,
                    )
                )

    return issues


def build_review_inventory_report(snapshot: ReviewArtifactSnapshot) -> str:
    counts = snapshot.package_counts
    return (
        "# DOCX Review Artifact Inventory\n\n"
        "## Summary\n"
        f"- docx path: {snapshot.docx_path}\n"
        f"- docx sha256: {snapshot.docx_sha256}\n"
        f"- comments.xml part count: {counts.comments_xml}\n"
        f"- commentsExtended.xml part count: {counts.comments_extended_xml}\n"
        f"- people.xml part count: {counts.people_xml}\n"
        f"- comment-related part names: {snapshot.comment_part_names if snapshot.comment_part_names else 'none'}\n"
        f"- comment anchor count: {counts.comment_anchors}\n"
        f"- tracked change mark count: {counts.tracked_changes}\n"
        f"- bookmark host count: {counts.bookmarks}\n"
        f"- field host count: {counts.fields}\n"
        f"- hyperlink host count: {counts.hyperlinks}\n"
        f"- result: pass\n\n"
        "## Comment Inventory\n"
        + ("\n".join(f"- comment id {item.comment_id}: {item.text_digest}" for item in snapshot.comments) if snapshot.comments else "- none")
        + "\n\n## Comment Anchors\n"
        + ("\n".join(f"- {item}" for item in snapshot.comment_anchor_ids) if snapshot.comment_anchor_ids else "- none")
        + "\n\n## Tracked Changes\n"
        + ("\n".join(f"- {item}" for item in snapshot.tracked_change_ids) if snapshot.tracked_change_ids else "- none")
        + "\n\n## Bookmarks\n"
        + ("\n".join(f"- {item}" for item in snapshot.bookmark_names) if snapshot.bookmark_names else "- none")
        + "\n\n## Fields\n"
        + ("\n".join(f"- {item}" for item in snapshot.field_instr_digests) if snapshot.field_instr_digests else "- none")
        + "\n\n## Hyperlinks\n"
        + ("\n".join(f"- {item}" for item in snapshot.hyperlink_anchors) if snapshot.hyperlink_anchors else "- none")
        + "\n"
    )


def build_review_diff_report(
    source: ReviewArtifactSnapshot,
    final: ReviewArtifactSnapshot,
    *,
    allow_approved_comment_disposal: bool = False,
    comment_disposal_audit_path: Path | None = None,
    comment_disposal_issues: list[str] | None = None,
    controlled_bookmark_disposition: BookmarkDisposition | None = None,
    controlled_bookmark_disposition_issues: list[str] | None = None,
) -> tuple[str, list[str]]:
    issues = compare_review_artifacts(
        source,
        final,
        allow_approved_comment_disposal=allow_approved_comment_disposal,
        controlled_bookmark_disposition=controlled_bookmark_disposition,
    )
    if comment_disposal_issues:
        issues.extend(f"comment disposal authorization failed: {issue}" for issue in comment_disposal_issues)
    if controlled_bookmark_disposition_issues:
        issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in controlled_bookmark_disposition_issues)
    counts = final.package_counts
    missing_bookmarks = sorted(set(source.bookmark_names) - set(final.bookmark_names))
    controlled_bookmarks = sorted(
        set(missing_bookmarks)
        & (controlled_bookmark_disposition.allowed_missing_bookmarks if controlled_bookmark_disposition else set())
    )
    missing_fields = sorted(set(source.field_instr_digests) - set(final.field_instr_digests))
    controlled_fields = sorted(
        set(missing_fields)
        & (
            controlled_bookmark_disposition.allowed_missing_field_instr_digests
            if controlled_bookmark_disposition
            else set()
        )
    )
    missing_hyperlinks = sorted(set(source.hyperlink_anchors) - set(final.hyperlink_anchors))
    controlled_hyperlinks = sorted(
        set(missing_hyperlinks)
        & (
            controlled_bookmark_disposition.allowed_missing_hyperlinks
            if controlled_bookmark_disposition
            else set()
        )
    )
    controlled_artifacts = bool(
        controlled_bookmarks
        or controlled_fields
        or controlled_hyperlinks
        or (
            controlled_bookmark_disposition is not None
            and controlled_bookmark_disposition.allow_source_citation_nonpreservation
        )
    )
    report = (
        "# DOCX Review Artifact Diff\n\n"
        "## Summary\n"
        f"- source docx path: {source.docx_path}\n"
        f"- source docx sha256: {source.docx_sha256}\n"
        f"- final docx path: {final.docx_path}\n"
        f"- final docx sha256: {final.docx_sha256}\n"
        f"- comments stripped: {'yes' if source.package_counts.comments_xml > 0 and final.package_counts.comments_xml == 0 else 'no'}\n"
        f"- comment-related parts stripped: {'yes' if set(source.comment_part_names) - set(final.comment_part_names) else 'no'}\n"
        f"- approved comment disposal: {'yes' if allow_approved_comment_disposal else 'no'}\n"
        f"- comment disposal audit path: {comment_disposal_audit_path if comment_disposal_audit_path is not None else 'none'}\n"
        f"- controlled bookmark disposition path: {controlled_bookmark_disposition.path if controlled_bookmark_disposition is not None else 'none'}\n"
        f"- controlled bookmark disposition: {'yes' if controlled_artifacts else 'no'}\n"
        f"- controlled missing bookmarks: {controlled_bookmarks if controlled_bookmarks else 'none'}\n"
        f"- controlled missing fields: {controlled_fields if controlled_fields else 'none'}\n"
        f"- controlled missing hyperlinks: {controlled_hyperlinks if controlled_hyperlinks else 'none'}\n"
        f"- tracked changes stripped: {'yes' if source.package_counts.tracked_changes > final.package_counts.tracked_changes else 'no'}\n"
        f"- bookmarks stripped: {'yes' if source.package_counts.bookmarks > 0 and len(set(missing_bookmarks) - set(controlled_bookmarks)) > 0 else 'no'}\n"
        f"- fields stripped: {'yes' if source.package_counts.fields > 0 and len(set(missing_fields) - set(controlled_fields)) > 0 else 'no'}\n"
        f"- hyperlinks stripped: {'yes' if source.package_counts.hyperlinks > 0 and len(set(missing_hyperlinks) - set(controlled_hyperlinks)) > 0 else 'no'}\n"
        f"- error codes: {', '.join(sorted(set(issues))) if issues else 'none'}\n"
        f"- result: {'pass' if not issues else 'fail'}\n\n"
        "## Findings\n"
    )
    if issues:
        report += "\n".join(f"- {item}" for item in issues) + "\n"
    else:
        report += "- none\n"
    return report, issues


def build_citation_inventory_report(snapshot: CitationSnapshot) -> str:
    return (
        "# DOCX Citation Run Inventory\n\n"
        "## Summary\n"
        f"- docx path: {snapshot.docx_path}\n"
        f"- docx sha256: {snapshot.docx_sha256}\n"
        f"- citation marker count: {len(snapshot.citation_numbers)}\n"
        f"- unique citation number count: {len(set(snapshot.citation_numbers))}\n"
        f"- result: pass\n\n"
        "## Citation Records\n"
        + (
            "\n".join(
                f"- paragraph {item.paragraph_index} number={item.number} marker_count={item.marker_count} superscript={'yes' if item.superscript_ok else 'no'} hyperlink={'yes' if item.hyperlink_ok else 'no'} visual={'yes' if item.visual_style_ok else 'no'} punctuation={'yes' if item.punctuation_ok else 'no'} surface={item.surface}"
                f" paragraph_text_digest={item.paragraph_text_digest}"
                for item in snapshot.records
            )
            if snapshot.records
            else "- none"
        )
        + "\n"
    )


def build_citation_diff_report(
    source: CitationSnapshot,
    final: CitationSnapshot,
    *,
    citation_preservation_scope: str = STRICT_CITATION_PRESERVATION_SCOPE,
) -> tuple[str, list[str]]:
    issues = compare_citations(
        source,
        final,
        citation_preservation_scope=citation_preservation_scope,
    )
    source_counts: dict[int, int] = {}
    for number in source.citation_numbers:
        source_counts[number] = source_counts.get(number, 0) + 1
    final_counts: dict[int, int] = {}
    for number in final.citation_numbers:
        final_counts[number] = final_counts.get(number, 0) + 1
    source_hyperlink_numbers = sorted(
        {
            record.number
            for record in source.records
            if record.hyperlink_ok
        }
    )
    final_hyperlink_numbers = sorted(
        {
            record.number
            for record in final.records
            if record.hyperlink_ok
        }
    )
    report = (
        "# DOCX Citation Run Diff\n\n"
        "## Summary\n"
        f"- source docx path: {source.docx_path}\n"
        f"- source docx sha256: {source.docx_sha256}\n"
        f"- final docx path: {final.docx_path}\n"
        f"- final docx sha256: {final.docx_sha256}\n"
        f"- source citation numbers: {sorted(source_counts)}\n"
        f"- final citation numbers: {sorted(final_counts)}\n"
        f"- source hyperlink citation numbers: {source_hyperlink_numbers if source_hyperlink_numbers else 'none'}\n"
        f"- final hyperlink citation numbers: {final_hyperlink_numbers if final_hyperlink_numbers else 'none'}\n"
        f"- citation hyperlink preservation scope: {'source-existing-only' if source_hyperlink_numbers else 'not-applicable-source-has-no-citation-hyperlinks'}\n"
        f"- citation occurrence preservation scope: {citation_preservation_scope}\n"
        f"- missing source citation numbers: {sorted(set(source_counts) - set(final_counts)) if source_counts else 'none'}\n"
        f"- error codes: {', '.join(sorted(set(issues))) if issues else 'none'}\n"
        f"- result: {'pass' if not issues else 'fail'}\n\n"
        "## Findings\n"
    )
    if issues:
        report += "\n".join(f"- {item}" for item in issues) + "\n"
    else:
        report += "- none\n"
    return report, issues


def validate_review_artifact_reports(source_report_path: Path, diff_report_path: Path, *, expected_final_docx: Path | None = None) -> list[str]:
    issues: list[str] = []
    for path in (source_report_path, diff_report_path):
        if not path.exists():
            issues.append(f"review artifact report missing: {path}")
            return issues
    source_text = source_report_path.read_text(encoding="utf-8")
    diff_text = diff_report_path.read_text(encoding="utf-8")
    source_values = parse_report_values(source_text)
    diff_values = parse_report_values(diff_text)
    if source_values.get("- result:") != "pass":
        issues.append(f"source review artifact inventory is not pass in {source_report_path}")
    if diff_values.get("- result:") != "pass":
        issues.append(f"final review artifact diff is not pass in {diff_report_path}")
    source_docx = Path(source_values.get("- docx path:", ""))
    final_docx = Path(diff_values.get("- final docx path:", ""))
    diff_source_docx_value = diff_values.get("- source docx path:", "")
    if not source_values.get("- docx path:", "").strip():
        issues.append(f"source review artifact inventory missing docx path in {source_report_path}")
        return issues
    if not diff_values.get("- final docx path:", "").strip():
        issues.append(f"review artifact diff missing final docx path in {diff_report_path}")
        return issues
    if not source_docx.is_absolute():
        source_docx = (source_report_path.parent / source_docx).resolve()
    else:
        source_docx = source_docx.resolve()
    if not final_docx.is_absolute():
        final_docx = (diff_report_path.parent / final_docx).resolve()
    else:
        final_docx = final_docx.resolve()
    if expected_final_docx is not None and final_docx != expected_final_docx.resolve():
        issues.append(
            f"review artifact diff in {diff_report_path} targets {final_docx} instead of exact final DOCX {expected_final_docx.resolve()}"
        )
    if not source_docx.exists():
        issues.append(f"source review artifact inventory docx path does not exist: {source_docx}")
        return issues
    if not final_docx.exists():
        issues.append(f"review artifact diff final docx path does not exist: {final_docx}")
        return issues
    if diff_source_docx_value.strip():
        diff_source_docx = Path(diff_source_docx_value)
        if not diff_source_docx.is_absolute():
            diff_source_docx = (diff_report_path.parent / diff_source_docx).resolve()
        else:
            diff_source_docx = diff_source_docx.resolve()
        if diff_source_docx != source_docx:
            issues.append(f"review artifact diff source docx path does not match source inventory in {diff_report_path}")
    expected_source_sha = sha256_file(source_docx)
    if source_values.get("- docx sha256:", "").lower() != expected_source_sha.lower():
        issues.append(f"source review artifact inventory docx sha256 does not match source DOCX in {source_report_path}")
    if diff_values.get("- source docx sha256:", "").strip() and diff_values.get("- source docx sha256:", "").lower() != expected_source_sha.lower():
        issues.append(f"review artifact diff source docx sha256 does not match source DOCX in {diff_report_path}")
    expected_final_sha = sha256_file(final_docx)
    if diff_values.get("- final docx sha256:", "").lower() != expected_final_sha.lower():
        issues.append(f"review artifact diff final docx sha256 does not match final DOCX in {diff_report_path}")
    if source_docx == final_docx:
        issues.append("review artifact reports must use distinct source and final DOCX paths")
        return issues
    recomputed_source = collect_review_artifacts(source_docx)
    recomputed_final = collect_review_artifacts(final_docx)
    approved_comment_disposal_claimed = diff_values.get("- approved comment disposal:", "").strip().lower() in {"yes", "true", "pass", "passed"}
    approved_comment_disposal = False
    if approved_comment_disposal_claimed:
        audit_path_value = diff_values.get("- comment disposal audit path:", "").strip()
        audit_path = Path(audit_path_value) if audit_path_value and audit_path_value.lower() != "none" else None
        if audit_path is None:
            issues.append(f"review artifact diff claims approved comment disposal but lacks comment disposal audit path in {diff_report_path}")
        else:
            if not audit_path.is_absolute():
                audit_path = (diff_report_path.parent / audit_path).resolve()
            else:
                audit_path = audit_path.resolve()
            if not audit_path.exists() or (audit_path.is_file() and audit_path.stat().st_size == 0):
                issues.append(f"review artifact diff comment disposal audit path is missing or empty: {audit_path}")
            else:
                comment_disposal_issues = validate_comment_resolution_ledger(
                    audit_path,
                    final_docx=final_docx,
                    source_docx=source_docx,
                    assert_all_resolved=True,
                )
                if comment_disposal_issues:
                    issues.extend(
                        f"comment disposal authorization failed: {issue}"
                        for issue in comment_disposal_issues
                    )
                else:
                    approved_comment_disposal = True
    disposition = None
    disposition_claimed = diff_values.get("- controlled bookmark disposition:", "").strip().lower() in {"yes", "true", "pass", "passed"}
    if disposition_claimed:
        disposition_path = _resolve_disposition_path(
            diff_values.get("- controlled bookmark disposition path:", ""),
            diff_report_path,
        )
        if disposition_path is None:
            issues.append(f"review artifact diff claims controlled bookmark disposition but lacks disposition path in {diff_report_path}")
        else:
            disposition, disposition_issues = load_empty_paragraph_bookmark_disposition(disposition_path)
            issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in disposition_issues)
    issues.extend(
        compare_review_artifacts(
            recomputed_source,
            recomputed_final,
            allow_approved_comment_disposal=approved_comment_disposal,
            controlled_bookmark_disposition=disposition,
        )
    )
    return issues


def validate_citation_run_reports(
    source_report_path: Path,
    diff_report_path: Path,
    *,
    expected_final_docx: Path | None = None,
    allowed_preservation_scopes: set[str] | None = None,
) -> list[str]:
    issues: list[str] = []
    for path in (source_report_path, diff_report_path):
        if not path.exists():
            issues.append(f"citation run report missing: {path}")
            return issues
    source_text = source_report_path.read_text(encoding="utf-8")
    diff_text = diff_report_path.read_text(encoding="utf-8")
    source_values = parse_report_values(source_text)
    diff_values = parse_report_values(diff_text)
    if source_values.get("- result:") != "pass":
        issues.append(f"source citation inventory is not pass in {source_report_path}")
    if diff_values.get("- result:") != "pass":
        issues.append(f"final citation diff is not pass in {diff_report_path}")
    source_docx = Path(source_values.get("- docx path:", ""))
    final_docx = Path(diff_values.get("- final docx path:", ""))
    diff_source_docx_value = diff_values.get("- source docx path:", "")
    if not source_values.get("- docx path:", "").strip():
        issues.append(f"source citation inventory missing docx path in {source_report_path}")
        return issues
    if not diff_values.get("- final docx path:", "").strip():
        issues.append(f"citation diff missing final docx path in {diff_report_path}")
        return issues
    if not source_docx.is_absolute():
        source_docx = (source_report_path.parent / source_docx).resolve()
    else:
        source_docx = source_docx.resolve()
    if not final_docx.is_absolute():
        final_docx = (diff_report_path.parent / final_docx).resolve()
    else:
        final_docx = final_docx.resolve()
    if expected_final_docx is not None and final_docx != expected_final_docx.resolve():
        issues.append(
            f"citation diff in {diff_report_path} targets {final_docx} instead of exact final DOCX {expected_final_docx.resolve()}"
        )
    if not source_docx.exists():
        issues.append(f"source citation inventory docx path does not exist: {source_docx}")
        return issues
    if not final_docx.exists():
        issues.append(f"citation diff final docx path does not exist: {final_docx}")
        return issues
    if diff_source_docx_value.strip():
        diff_source_docx = Path(diff_source_docx_value)
        if not diff_source_docx.is_absolute():
            diff_source_docx = (diff_report_path.parent / diff_source_docx).resolve()
        else:
            diff_source_docx = diff_source_docx.resolve()
        if diff_source_docx != source_docx:
            issues.append(f"citation diff source docx path does not match source inventory in {diff_report_path}")
    expected_source_sha = sha256_file(source_docx)
    if source_values.get("- docx sha256:", "").lower() != expected_source_sha.lower():
        issues.append(f"source citation inventory docx sha256 does not match source DOCX in {source_report_path}")
    if diff_values.get("- source docx sha256:", "").strip() and diff_values.get("- source docx sha256:", "").lower() != expected_source_sha.lower():
        issues.append(f"citation diff source docx sha256 does not match source DOCX in {diff_report_path}")
    expected_final_sha = sha256_file(final_docx)
    if diff_values.get("- final docx sha256:", "").lower() != expected_final_sha.lower():
        issues.append(f"citation diff final docx sha256 does not match final DOCX in {diff_report_path}")
    if source_docx == final_docx:
        issues.append("citation run reports must use distinct source and final DOCX paths")
        return issues
    source_snapshot = collect_citation_snapshot(source_docx)
    final_snapshot = collect_citation_snapshot(final_docx)
    citation_preservation_scope = diff_values.get(
        "- citation occurrence preservation scope:",
        STRICT_CITATION_PRESERVATION_SCOPE,
    )
    if allowed_preservation_scopes is not None and citation_preservation_scope not in allowed_preservation_scopes:
        issues.append(
            "citation occurrence preservation scope "
            f"{citation_preservation_scope} is not allowed for this thesis workflow"
        )
    issues.extend(
        compare_citations(
            source_snapshot,
            final_snapshot,
            citation_preservation_scope=citation_preservation_scope,
        )
    )
    if citation_preservation_scope != STRICT_CITATION_PRESERVATION_SCOPE and (
        source_snapshot.records or final_snapshot.records
    ):
        try:
            final_audit = audit_body_citations(final_docx)
        except Exception as exc:
            issues.append(f"final DOCX body citation audit could not run in {final_docx}: {exc}")
        else:
            if not final_audit.passed:
                issues.append(f"final DOCX body citation audit failed in {final_docx}: {', '.join(final_audit.error_codes)}")
    return issues


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--review-inventory-report")
    parser.add_argument("--review-diff-report")
    parser.add_argument("--citation-inventory-report")
    parser.add_argument("--citation-diff-report")
    parser.add_argument(
        "--citation-preservation-scope",
        choices=sorted(ALLOWED_CITATION_PRESERVATION_SCOPES),
        default=STRICT_CITATION_PRESERVATION_SCOPE,
        help="Use whole-rebuild-chain-integrity only for whole-thesis/new-thesis rebuilds where final citation chain audit is the preservation authority.",
    )
    parser.add_argument(
        "--comment-resolution-ledger",
        help="Allow source comments to be absent from a clean final only when this ledger validates approved disposal.",
    )
    parser.add_argument(
        "--controlled-bookmark-disposition",
        help="Allow only listed empty-paragraph bookmark loss when a SHA-bound disposition record validates it.",
    )
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--json-output", help="Write the JSON audit payload as UTF-8 without relying on shell redirection.")
    args = parser.parse_args(argv)

    source_docx = Path(args.source_docx).resolve()
    final_docx = Path(args.final_docx).resolve()
    review_source = collect_review_artifacts(source_docx)
    review_final = collect_review_artifacts(final_docx)
    citation_source = collect_citation_snapshot(source_docx)
    citation_final = collect_citation_snapshot(final_docx)

    comment_disposal_audit_path = Path(args.comment_resolution_ledger).resolve() if args.comment_resolution_ledger else None
    comment_disposal_issues: list[str] = []
    allow_approved_comment_disposal = False
    if comment_disposal_audit_path is not None:
        comment_disposal_issues = validate_comment_resolution_ledger(
            comment_disposal_audit_path,
            final_docx=final_docx,
            source_docx=source_docx,
            assert_all_resolved=True,
        )
        allow_approved_comment_disposal = not comment_disposal_issues
    controlled_bookmark_disposition = None
    controlled_bookmark_disposition_issues: list[str] = []
    controlled_bookmark_disposition_path = (
        Path(args.controlled_bookmark_disposition).resolve()
        if args.controlled_bookmark_disposition
        else None
    )
    if controlled_bookmark_disposition_path is not None:
        controlled_bookmark_disposition, controlled_bookmark_disposition_issues = load_empty_paragraph_bookmark_disposition(
            controlled_bookmark_disposition_path
        )

    review_inventory_text = build_review_inventory_report(review_source)
    review_diff_text, review_issues = build_review_diff_report(
        review_source,
        review_final,
        allow_approved_comment_disposal=allow_approved_comment_disposal,
        comment_disposal_audit_path=comment_disposal_audit_path,
        comment_disposal_issues=comment_disposal_issues,
        controlled_bookmark_disposition=controlled_bookmark_disposition,
        controlled_bookmark_disposition_issues=controlled_bookmark_disposition_issues,
    )
    citation_inventory_text = build_citation_inventory_report(citation_source)
    citation_diff_text, citation_issues = build_citation_diff_report(
        citation_source,
        citation_final,
        citation_preservation_scope=args.citation_preservation_scope,
    )
    json_payload = {
        "review": asdict(review_source),
        "review_diff_issues": review_issues,
        "comment_disposal_ledger": str(comment_disposal_audit_path) if comment_disposal_audit_path is not None else None,
        "comment_disposal_authorized": allow_approved_comment_disposal,
        "comment_disposal_issues": comment_disposal_issues,
        "controlled_bookmark_disposition": str(controlled_bookmark_disposition_path) if controlled_bookmark_disposition_path is not None else None,
        "controlled_bookmark_disposition_issues": controlled_bookmark_disposition_issues,
        "citation": asdict(citation_source),
        "citation_diff_issues": citation_issues,
    }

    if args.review_inventory_report:
        Path(args.review_inventory_report).resolve().parent.mkdir(parents=True, exist_ok=True)
        Path(args.review_inventory_report).write_text(review_inventory_text, encoding="utf-8")
    if args.review_diff_report:
        Path(args.review_diff_report).resolve().parent.mkdir(parents=True, exist_ok=True)
        Path(args.review_diff_report).write_text(review_diff_text, encoding="utf-8")
    if args.citation_inventory_report:
        Path(args.citation_inventory_report).resolve().parent.mkdir(parents=True, exist_ok=True)
        Path(args.citation_inventory_report).write_text(citation_inventory_text, encoding="utf-8")
    if args.citation_diff_report:
        Path(args.citation_diff_report).resolve().parent.mkdir(parents=True, exist_ok=True)
        Path(args.citation_diff_report).write_text(citation_diff_text, encoding="utf-8")
    if args.json_output:
        json_output = Path(args.json_output).resolve()
        json_output.parent.mkdir(parents=True, exist_ok=True)
        json_output.write_text(
            json.dumps(json_payload, ensure_ascii=True, indent=2, default=str) + "\n",
            encoding="utf-8",
        )

    if args.json:
        print(json.dumps(json_payload, ensure_ascii=True, indent=2, default=str))
    else:
        overall = not review_issues and not citation_issues
        print(f"DOCX preservation audit: {'PASS' if overall else 'FAIL'}")
        if args.review_diff_report:
            print(f"- review diff: {Path(args.review_diff_report).resolve()}")
        if args.citation_diff_report:
            print(f"- citation diff: {Path(args.citation_diff_report).resolve()}")
    return 0 if not review_issues and not citation_issues else 1


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raise SystemExit(main(sys.argv[1:]))
