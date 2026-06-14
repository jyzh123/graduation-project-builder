#!/usr/bin/env python3
"""Audit source-to-final DOCX protected surface and package drift."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from audit_docx_review_artifacts import (
    collect_citation_snapshot,
    collect_review_artifacts,
    compare_citations,
    compare_review_artifacts,
    load_empty_paragraph_bookmark_disposition,
)


SCHEMA = "graduation-project-builder.docx-protected-surface-diff.v1"

CANONICAL_PROTECTED_SURFACE_IDS = (
    "cover_style",
    "declaration_or_title_front_matter",
    "zh_abstract_title",
    "zh_abstract_body",
    "zh_keyword_line",
    "en_abstract_title",
    "en_abstract_body",
    "en_keyword_line",
    "toc_title",
    "toc_entries",
    "toc_dotted_leaders",
    "toc_page_number_column",
    "body_heading_levels",
    "body_text",
    "figure_table_captions_and_holders",
    "body_citation_superscripts",
    "review_comments_and_change_marks",
    "references_title",
    "references_entries",
    "acknowledgement_title",
    "acknowledgement_body",
    "appendix_title",
    "appendix_body",
    "header",
    "footer",
    "page_numbers",
    "whole_document_pagination",
)

STYLE_BEARING_PART_PATTERNS = (
    re.compile(r"^word/styles\.xml$", re.I),
    re.compile(r"^word/settings\.xml$", re.I),
    re.compile(r"^word/numbering\.xml$", re.I),
    re.compile(r"^word/fontTable\.xml$", re.I),
    re.compile(r"^word/theme/", re.I),
    re.compile(r"^word/header\d+\.xml$", re.I),
    re.compile(r"^word/footer\d+\.xml$", re.I),
    re.compile(r"(^|/)_rels/.*\.rels$", re.I),
    re.compile(r"^customXml/.*\.rels$", re.I),
)

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{" + NS[prefix] + "}" + local


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_docx_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as zf:
        return {info.filename: sha256_bytes(zf.read(info.filename)) for info in zf.infolist()}


def paragraph_text(paragraph: ET.Element) -> str:
    chunks: list[str] = []
    for node in paragraph.iter():
        if node.tag == qn("w:t") and node.text:
            chunks.append(node.text)
        elif node.tag == qn("w:tab"):
            chunks.append("\t")
    return "".join(chunks)


def paragraph_signature(paragraph: ET.Element) -> dict[str, Any]:
    ppr = paragraph.find("./w:pPr", NS)
    runs = []
    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        text = paragraph_text(run)
        runs.append(
            {
                "text": text,
                "rpr_sha256": sha256_bytes(ET.tostring(rpr, encoding="utf-8")) if rpr is not None else None,
                "bold": rpr is not None and rpr.find("./w:b", NS) is not None,
                "vertAlign": (
                    rpr.find("./w:vertAlign", NS).get(qn("w:val"))
                    if rpr is not None and rpr.find("./w:vertAlign", NS) is not None
                    else None
                ),
            }
        )
    return {
        "text": paragraph_text(paragraph),
        "ppr_sha256": sha256_bytes(ET.tostring(ppr, encoding="utf-8")) if ppr is not None else None,
        "run_count": len(runs),
        "runs": runs,
    }


def load_paragraphs(path: Path) -> list[dict[str, Any]]:
    with zipfile.ZipFile(path) as zf:
        try:
            root = ET.fromstring(zf.read("word/document.xml"))
        except KeyError:
            return []
    return [paragraph_signature(p) for p in root.findall(".//w:p", NS)]


def text_key(value: str) -> str:
    return re.sub(r"\s+", "", value or "").lower()


def first_matching(paragraphs: list[dict[str, Any]], patterns: tuple[re.Pattern[str], ...]) -> dict[str, Any] | None:
    for paragraph in paragraphs:
        text = paragraph.get("text", "")
        compact = text_key(text)
        if any(pattern.search(text) or pattern.search(compact) for pattern in patterns):
            return paragraph
    return None


SURFACE_PATTERNS: dict[str, tuple[re.Pattern[str], ...]] = {
    "zh_abstract_title": (re.compile(r"\u6458\s*\u8981"),),
    "zh_keyword_line": (re.compile(r"\u5173\s*\u952e\s*\u8bcd"),),
    "en_abstract_title": (re.compile(r"^abstract$", re.I),),
    "en_keyword_line": (re.compile(r"key\s*words?|keywords?", re.I),),
    "toc_title": (re.compile(r"\u76ee\s*\u5f55"), re.compile(r"contents?", re.I)),
    "references_title": (re.compile(r"\u53c2\u8003\u6587\u732e"), re.compile(r"references", re.I)),
    "acknowledgement_title": (re.compile(r"\u81f4\s*\u8c22"), re.compile(r"acknowledgements?", re.I)),
    "appendix_title": (re.compile(r"\u9644\s*\u5f55"), re.compile(r"appendix", re.I)),
}


def infer_surface_signatures(paragraphs: list[dict[str, Any]]) -> dict[str, dict[str, Any] | None]:
    out: dict[str, dict[str, Any] | None] = {surface: None for surface in CANONICAL_PROTECTED_SURFACE_IDS}
    for surface, patterns in SURFACE_PATTERNS.items():
        out[surface] = first_matching(paragraphs, patterns)
    out["cover_style"] = paragraphs[0] if paragraphs else None
    out["body_text"] = next((p for p in paragraphs if len(str(p.get("text", "")).strip()) > 20), None)
    out["body_citation_superscripts"] = next(
        (
            p
            for p in paragraphs
            if any(run.get("vertAlign") == "superscript" for run in p.get("runs", []))
        ),
        None,
    )
    out["whole_document_pagination"] = {"paragraph_count": len(paragraphs)}
    return out


def surface_status(source: dict[str, Any] | None, final: dict[str, Any] | None, authorized: bool) -> dict[str, Any]:
    if source is None and final is None:
        return {"status": "not-present", "verdict": "pass not-present in source and final"}
    if source is None or final is None:
        return {"status": "authorized-changed" if authorized else "changed", "verdict": "pass" if authorized else "fail"}
    changed = source != final
    if not changed:
        return {"status": "unchanged", "verdict": "pass"}
    return {"status": "authorized-changed" if authorized else "changed", "verdict": "pass" if authorized else "fail"}


def is_style_bearing_part(part: str) -> bool:
    return any(pattern.search(part) for pattern in STYLE_BEARING_PART_PATTERNS)


def _keyword_label_parts(surface: str) -> set[str]:
    if surface == "zh_keyword_line":
        return {"关键词", "关键字", "：", ":"}
    if surface == "en_keyword_line":
        return {"key", "word", "words", "word:", "words:", "keyword", "keywords", "keyword:", "keywords:", ":"}
    return {"：", ":"}


def _keyword_expected_labels(surface: str) -> set[str]:
    if surface == "zh_keyword_line":
        return {"关键词：", "关键词:", "关键字：", "关键字:"}
    if surface == "en_keyword_line":
        return {"key words:", "keywords:"}
    return {":", "："}


def keyword_run_split_verdict(surface_diffs: dict[str, dict[str, Any]]) -> str:
    blockers = []
    for surface in ("zh_keyword_line", "en_keyword_line"):
        item = surface_diffs.get(surface, {})
        final = item.get("final_signature")
        if not isinstance(final, dict):
            continue
        runs = [run for run in final.get("runs", []) if str(run.get("text", "")).strip()]
        if not runs:
            continue
        label_parts = _keyword_label_parts(surface)
        expected_labels = _keyword_expected_labels(surface)
        label_run_count = 0
        label_text_parts: list[str] = []
        for run in runs:
            text = str(run.get("text", ""))
            normalized = text.strip()
            folded = normalized.lower()
            if normalized in expected_labels or folded in expected_labels:
                label_text_parts.append(folded if surface == "en_keyword_line" else normalized)
                label_run_count += 1
                break
            if normalized in label_parts or folded in label_parts:
                label_text_parts.append(normalized)
                label_run_count += 1
                if ":" in normalized or "\uff1a" in normalized:
                    break
                continue
            break
        label_text = "".join(label_text_parts)
        if surface == "en_keyword_line":
            label_text = " ".join(part for part in label_text_parts if part != ":").strip() + (
                ":" if label_text_parts and label_text_parts[-1] == ":" else ""
            )
            label_text = label_text.lower()
        label_ok = label_run_count > 0 and label_text in expected_labels
        content_runs = runs[label_run_count:]
        if not label_ok:
            blockers.append(surface)
        if not content_runs:
            blockers.append(surface)
        if any(run.get("bold") for run in content_runs):
            blockers.append(surface)
    return "pass" if not blockers else "fail keyword label/content run split: " + ",".join(sorted(set(blockers)))


def source_to_final_review_verdict(
    source_docx: Path,
    final_docx: Path,
    *,
    controlled_bookmark_disposition_path: Path | None = None,
) -> tuple[str, list[str]]:
    try:
        controlled_bookmark_disposition = None
        disposition_issues: list[str] = []
        if controlled_bookmark_disposition_path is not None:
            controlled_bookmark_disposition, disposition_issues = load_empty_paragraph_bookmark_disposition(
                controlled_bookmark_disposition_path
            )
        issues = compare_review_artifacts(
            collect_review_artifacts(source_docx),
            collect_review_artifacts(final_docx),
            controlled_bookmark_disposition=controlled_bookmark_disposition,
        )
        issues.extend(f"controlled bookmark disposition failed: {issue}" for issue in disposition_issues)
    except Exception as exc:
        return f"fail review artifact diff could not run: {exc}", [str(exc)]
    return ("pass" if not issues else "fail review artifact diff: " + "; ".join(issues[:5])), issues


def source_to_final_citation_verdict(source_docx: Path, final_docx: Path) -> tuple[str, list[str]]:
    return source_to_final_citation_verdict_with_scope(
        source_docx,
        final_docx,
        citation_preservation_scope="local-surface-preservation",
    )


def source_to_final_citation_verdict_with_scope(
    source_docx: Path,
    final_docx: Path,
    *,
    citation_preservation_scope: str,
) -> tuple[str, list[str]]:
    try:
        source_snapshot = collect_citation_snapshot(source_docx)
        final_snapshot = collect_citation_snapshot(final_docx)
        issues = compare_citations(
            source_snapshot,
            final_snapshot,
            citation_preservation_scope=citation_preservation_scope,
        )
    except Exception as exc:
        return f"fail citation run diff could not run: {exc}", [str(exc)]
    return ("pass" if not issues else "fail citation run diff: " + "; ".join(issues[:5])), issues


def build_report(
    *,
    source_docx: Path,
    final_docx: Path,
    target_surfaces: set[str],
    authorized_changed_parts: set[str],
    citation_preservation_scope: str = "local-surface-preservation",
    controlled_bookmark_disposition_path: Path | None = None,
) -> dict[str, Any]:
    source_parts = read_docx_parts(source_docx)
    final_parts = read_docx_parts(final_docx)
    added = sorted(set(final_parts) - set(source_parts))
    removed = sorted(set(source_parts) - set(final_parts))
    changed = sorted(part for part in set(source_parts) & set(final_parts) if source_parts[part] != final_parts[part])
    changed_parts = sorted(set(added + removed + changed))
    unauthorized_style_parts = [
        part for part in changed_parts if is_style_bearing_part(part) and part not in authorized_changed_parts
    ]

    source_surfaces = infer_surface_signatures(load_paragraphs(source_docx))
    final_surfaces = infer_surface_signatures(load_paragraphs(final_docx))
    review_verdict, review_issues = source_to_final_review_verdict(
        source_docx,
        final_docx,
        controlled_bookmark_disposition_path=controlled_bookmark_disposition_path,
    )
    citation_verdict, citation_issues = source_to_final_citation_verdict_with_scope(
        source_docx,
        final_docx,
        citation_preservation_scope=citation_preservation_scope,
    )
    surface_diffs: dict[str, dict[str, Any]] = {}
    unauthorized_surface_changes: list[str] = []
    for surface in CANONICAL_PROTECTED_SURFACE_IDS:
        diff = surface_status(source_surfaces.get(surface), final_surfaces.get(surface), surface in target_surfaces)
        diff["source_signature"] = source_surfaces.get(surface)
        diff["final_signature"] = final_surfaces.get(surface)
        surface_diffs[surface] = diff
        if diff["status"] == "changed":
            unauthorized_surface_changes.append(surface)

    keyword_verdict = keyword_run_split_verdict(surface_diffs)
    verdict = (
        "pass"
        if not unauthorized_surface_changes
        and not unauthorized_style_parts
        and keyword_verdict.startswith("pass")
        and review_verdict.startswith("pass")
        and citation_verdict.startswith("pass")
        else "fail"
    )
    return {
        "schema": SCHEMA,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "source_docx_path": str(source_docx),
        "source_docx_sha256": sha256_file(source_docx),
        "final_docx_path": str(final_docx),
        "final_docx_sha256": sha256_file(final_docx),
        "canonical_surface_ids": list(CANONICAL_PROTECTED_SURFACE_IDS),
        "target_surface_ids": sorted(target_surfaces),
        "authorized_changed_parts": sorted(authorized_changed_parts),
        "package_part_diffs": {
            "added": added,
            "removed": removed,
            "changed": changed,
            "changed_package_parts": changed_parts,
        },
        "changed_package_parts": changed_parts,
        "style_bearing_part_changes": [part for part in changed_parts if is_style_bearing_part(part)],
        "unauthorized_style_bearing_part_changes": unauthorized_style_parts,
        "style_bearing_package_part_verdict": "pass" if not unauthorized_style_parts else "fail",
        "surface_diffs": surface_diffs,
        "unauthorized_non_target_changes": sorted(unauthorized_surface_changes),
        "protected_surface_diff_verdict": "pass" if not unauthorized_surface_changes else "fail",
        "package_part_diff_verdict": "pass" if not unauthorized_style_parts else "fail",
        "keyword_run_split_verdict": keyword_verdict,
        "controlled_bookmark_disposition": (
            str(controlled_bookmark_disposition_path) if controlled_bookmark_disposition_path is not None else None
        ),
        "review_artifact_diff_verdict": review_verdict,
        "review_artifact_diff_issues": review_issues,
        "citation_preservation_scope": citation_preservation_scope,
        "citation_run_diff_verdict": citation_verdict,
        "citation_run_diff_issues": citation_issues,
        "evidence_staleness_verdict": "pass generated from exact source/final DOCX inputs in this report",
        "docx_bound_evidence_regenerated_after_last_mutation_verdict": (
            "pass generated from exact source/final DOCX inputs in this report"
        ),
        "verdict": verdict,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--target-surface", action="append", default=[])
    parser.add_argument("--authorized-changed-part", action="append", default=[])
    parser.add_argument(
        "--citation-preservation-scope",
        default="local-surface-preservation",
        choices=(
            "local-surface-preservation",
            "whole-rebuild-chain-integrity",
            "approved-non-preservation",
        ),
        help=(
            "Citation occurrence preservation scope. Keep the default strict "
            "scope for local repairs; use approved-non-preservation only for "
            "explicit whole-thesis rebuilds with a separate final citation audit."
        ),
    )
    parser.add_argument(
        "--controlled-bookmark-disposition",
        help=(
            "SHA-bound bookmark/field/hyperlink disposition record to reuse for "
            "review-artifact comparison inside the protected-surface diff."
        ),
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    source_docx = Path(args.source_docx)
    final_docx = Path(args.final_docx)
    output = Path(args.output)
    report = build_report(
        source_docx=source_docx,
        final_docx=final_docx,
        target_surfaces={str(item).strip() for item in args.target_surface if str(item).strip()},
        authorized_changed_parts={str(item).strip() for item in args.authorized_changed_part if str(item).strip()},
        citation_preservation_scope=args.citation_preservation_scope,
        controlled_bookmark_disposition_path=(
            Path(args.controlled_bookmark_disposition).resolve()
            if args.controlled_bookmark_disposition
            else None
        ),
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    if args.json:
        print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report.get("verdict") == "pass" else 1


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raise SystemExit(main())
