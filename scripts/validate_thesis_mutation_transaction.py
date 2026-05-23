#!/usr/bin/env python3
"""Validate thesis DOCX mutation transaction records."""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

from thesis_figure_contract import (
    docx_drawing_object_manifest,
    docx_image_relationship_manifest,
    validate_figure_manifest,
)
from audit_thesis_comment_resolution import (
    has_all_comments_claim,
    validate_comment_resolution_ledger,
)


SCHEMA = "graduation-project-builder.thesis-mutation-transaction.v1"
WORKFLOWS = {
    "new-thesis-production",
    "whole-thesis-revision",
    "local-surface-repair",
    "content-only-paragraph-revision",
    "audit-only",
}
TRANSACTION_WORKFLOWS = {
    "whole-thesis-revision",
    "local-surface-repair",
    "audit-only",
}
BASE_REQUIRED_EVIDENCE = (
    "protected_surface_freeze_manifest",
    "post_mutation_surface_diff",
    "target_surface_render_review",
    "blast_radius_render_review",
    "cross_surface_regression_report",
)
CHAPTER_FORMAT_PRESERVATION_EVIDENCE = "chapter_format_preservation_report"
CHAPTER_FORMAT_PRESERVATION_DETECTOR = "chapter.format-preservation-contract"
NON_TARGET_FORMAT_PRESERVATION_FIELD = "non_target_format_preservation_verdict"
BLOCKING_VERDICT_TOKENS = {
    "fail",
    "failed",
    "missing",
    "not checked",
    "pending",
    "blocked",
    "unresolved",
    "drift",
    "stale",
    "not-applicable",
}
WHOLE_SCOPE_TOKENS = {
    "whole-thesis",
    "whole thesis",
    "full-thesis",
    "full thesis",
    "whole-paper",
    "full-paper",
    "template alignment",
    "1:1",
    "submission-ready",
    "ready to submit",
    "ready-to-submit",
}
NEGATED_WHOLE_SCOPE_PATTERNS = (
    re.compile(
        r"\b(?:no|not|without)\s+(?:a\s+)?(?:whole[- ]thesis|full[- ]thesis|whole[- ]paper|full[- ]paper|template alignment|submission[- ]ready)\b"
    ),
    re.compile(
        r"\b(?:does\s+not|do\s+not|cannot|must\s+not)\s+(?:claim|assert|state|report)\s+(?:a\s+)?(?:whole[- ]thesis|full[- ]thesis|whole[- ]paper|full[- ]paper|template alignment|submission[- ]ready)\b"
    ),
)
TOC_PAGE_NUMBER_SURFACES = {
    "toc_page_number_column",
    "toc page-number column",
    "toc page number column",
}
COMMENT_DRIVEN_TOKENS = {
    "comment",
    "comments",
    "teacher comment",
    "review comment",
    "comment-driven",
    "tracked change",
}
FIXED_LOCATOR_TOKENS = {
    "doc.paragraphs[",
    "document.paragraphs[",
    "paragraphs[",
    "paragraph[",
    "paragraph index",
    "fixed paragraph",
    "fixed index",
    "word com paragraph",
    "com paragraph",
    "paragraph number",
    "range.paragraphs",
    "paragraph.next",
    "paragraph.previous",
}
FIXED_LOCATOR_PATTERNS = (
    re.compile(r"\b(?:doc|document)\.paragraphs\s*\[\s*\d+\s*\]"),
    re.compile(r"\bparagraphs?\s*\[\s*\d+\s*\]"),
    re.compile(r"\b(?:range\.)?paragraphs\s*\(\s*\d+\s*\)", re.IGNORECASE),
    re.compile(r"\bparagraph(?:_index| index)\s*(?:==|=|:)\s*\d+\b"),
    re.compile(r"第\s*(?:\d+|[一二三四五六七八九十百零〇两]+)\s*段"),
)
IMAGE_MUTATION_TOKENS = {
    "figure",
    "image",
    "picture",
    "drawing",
    "screenshot",
    "insert image",
    "replace image",
    "figure insertion",
    "图片",
    "图像",
    "截图",
    "插图",
    "绘图",
    "图表",
    "替换图片",
    "替换图像",
    "替换截图",
    "插入图片",
    "插入图像",
    "插入截图",
}
GENERIC_IMAGE_WORD_TOKENS = {
    "image",
    "图像",
}
IMAGE_MUTATION_ACTION_TOKENS = {
    "insert",
    "insertion",
    "replace",
    "replacement",
    "redraw",
    "draw",
    "crop",
    "resize",
    "move",
    "delete",
    "remove",
    "add",
    "generate",
    "rebuild",
    "rewrite",
    "change",
    "modify",
    "update",
    "swap",
    "插入",
    "替换",
    "重绘",
    "绘制",
    "裁剪",
    "缩放",
    "移动",
    "删除",
    "移除",
    "新增",
    "生成",
    "重建",
    "重写",
    "修改",
    "更新",
}
IMAGE_MUTATION_NEGATION_TOKENS = {
    "do not",
    "does not",
    "did not",
    "must not",
    "should not",
    "cannot",
    "without",
    "not-applicable",
    "no image",
    "no media",
    "no drawing",
    "unchanged",
    "preserve",
    "preserved",
    "preservation",
    "skip",
    "skipped",
    "false",
    "不要",
    "不得",
    "不能",
    "未",
    "不",
    "无",
    "保留",
    "不变",
}
PROTECTED_IMAGE_TARGET_TOKENS = {
    "toc",
    "table of contents",
    "front matter",
    "cover",
    "declaration",
    "abstract",
    "keyword",
}
WHOLE_REWRITE_TOKENS = {
    "whole-doc rewrite",
    "whole document rewrite",
    "whole document rebuild",
    "full body rebuild",
    "global format rewrite",
    "regenerate docx",
    "regenerate document",
    "rebuild full document",
    "rebuild full docx",
    "rewrite body",
    "rewrite entire body",
    "rebuild document from plain text",
    "copy visible text into new docx",
    "delete all paragraphs",
    "clear document body",
    "replace document.xml",
    "recreate docx",
    "rewrite all paragraphs",
}
WHOLE_REWRITE_PATTERNS = (
    re.compile(r"\bparagraph\s*\.\s*clear\s*\(", re.IGNORECASE),
)
NEGATED_BODY_SCOPE_PATTERNS = (
    re.compile(
        r"\b(?:do\s+not|does\s+not|did\s+not|must\s+not|should\s+not|cannot|without|no)\b[^.;\n]{0,120}\b(?:body|chapter|body\s+paragraphs?)\b",
        re.IGNORECASE,
    ),
)
BODY_CHAPTER_SURFACE_TOKENS = {
    "body",
    "body_text",
    "body text",
    "body_heading",
    "body heading",
    "body_heading_levels",
    "chapter",
    "chapter_body",
    "chapter body",
}
BODY_CHAPTER_OPERATION_PATTERNS = (
    re.compile(r"\bbody[_ -]?(?:text|paragraphs?|chapters?|headings?)\b", re.IGNORECASE),
    re.compile(r"\bchapter(?:[_ -]?(?:body|text|paragraphs?|headings?))?\b", re.IGNORECASE),
)
BODY_DISABLED_OPTION_PATTERNS = (
    re.compile(r"\bnormalize_body_[a-z0-9_]*\s*=\s*false\b", re.IGNORECASE),
    re.compile(r"\bbody[_ -]?citation(?:s|[_ -]?superscripts?)?\b", re.IGNORECASE),
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize(value: object) -> str:
    return " ".join(str(value or "").strip().lower().split())


def is_pass(value: object) -> bool:
    text = normalize(value)
    if "pass" not in text:
        return False
    return not any(token in text for token in BLOCKING_VERDICT_TOKENS)


def is_not_applicable(value: object) -> bool:
    text = normalize(value)
    if not text:
        return False
    return text in {"none", "n/a", "not-applicable", "not applicable"} or text.startswith(
        ("not-applicable-", "not applicable ")
    )


def states_no_chapter_format_promise(value: object) -> bool:
    text = normalize(value).replace("_", "-")
    return any(
        token in text
        for token in (
            "no-whole-format-claim",
            "no whole format claim",
            "no-format-preservation-promise",
            "no format preservation promise",
            "no-chapter-format-claim",
            "no chapter format claim",
            "local-keyword-only-no-whole-format-claim",
        )
    )


def as_list(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, tuple):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    if not text or text.lower() in {"none", "n/a", "not-applicable"}:
        return []
    parts = []
    for separator in (";", ",", "\n"):
        if separator in text:
            parts = [part.strip() for part in text.split(separator)]
            break
    if not parts:
        parts = [text]
    return [part for part in parts if part]


def has_any_token(text: str, tokens: set[str]) -> bool:
    lowered = text.lower()
    return any(token.lower() in lowered for token in tokens)


def has_specific_image_mutation_token(text: str) -> bool:
    lowered = text.lower()
    return any(
        token.lower() in lowered
        for token in IMAGE_MUTATION_TOKENS
        if token not in GENERIC_IMAGE_WORD_TOKENS
    )


def remove_patterns(text: str, patterns: tuple[re.Pattern[str], ...]) -> str:
    cleaned = text
    for pattern in patterns:
        cleaned = pattern.sub(" ", cleaned)
    return normalize(cleaned)


def has_whole_scope_claim(text: str) -> bool:
    return any(token in remove_patterns(text, NEGATED_WHOLE_SCOPE_PATTERNS) for token in WHOLE_SCOPE_TOKENS)


def has_image_mutation_intent(payload: dict[str, Any]) -> bool:
    fields = (
        "target_surfaces",
        "planned_operations",
        "operation_summary",
        "mutation_intent",
        "target_anchor_surface",
        "target_anchor_caption",
        "target_anchor_chapter",
        "insertion_target_surface",
        "write_scope",
    )
    for key in fields:
        value = payload.get(key)
        items = value if isinstance(value, list) else [value]
        for item in items:
            text = normalize(item)
            if not has_any_token(text, IMAGE_MUTATION_TOKENS):
                continue
            if key in {"target_surfaces", "target_anchor_surface", "insertion_target_surface"}:
                return True
            if has_any_token(text, IMAGE_MUTATION_NEGATION_TOKENS):
                continue
            if has_any_token(text, IMAGE_MUTATION_ACTION_TOKENS):
                return True
            if has_specific_image_mutation_token(text):
                return True
    return False


def has_fixed_locator_signal(text: str) -> bool:
    lowered = text.lower()
    return has_any_token(lowered, FIXED_LOCATOR_TOKENS) or any(
        pattern.search(text) for pattern in FIXED_LOCATOR_PATTERNS
    )


def has_whole_rewrite_signal(text: str) -> bool:
    return has_any_token(text, WHOLE_REWRITE_TOKENS) or any(
        pattern.search(text) for pattern in WHOLE_REWRITE_PATTERNS
    )


def touches_body_chapter_surface(target_surfaces: list[str], operation_text: str) -> bool:
    surface_text = normalize(" ".join(target_surfaces))
    if has_any_token(surface_text, BODY_CHAPTER_SURFACE_TOKENS):
        return True
    positive_operation_text = remove_patterns(operation_text, NEGATED_BODY_SCOPE_PATTERNS)
    positive_operation_text = remove_patterns(positive_operation_text, BODY_DISABLED_OPTION_PATTERNS)
    return any(pattern.search(positive_operation_text) for pattern in BODY_CHAPTER_OPERATION_PATTERNS)


def requires_chapter_format_preservation(payload: dict[str, Any], target_surfaces: list[str], operation_text: str) -> bool:
    if touches_body_chapter_surface(target_surfaces, operation_text):
        return True
    if payload.get(CHAPTER_FORMAT_PRESERVATION_EVIDENCE) is not None:
        return True
    detector_verdicts = payload.get("detector_verdicts")
    if (
        isinstance(detector_verdicts, dict)
        and CHAPTER_FORMAT_PRESERVATION_DETECTOR in detector_verdicts
        and not is_not_applicable(detector_verdicts.get(CHAPTER_FORMAT_PRESERVATION_DETECTOR))
    ):
        return True
    for key in ("format_preservation_promise_verdict", "chapter_format_preservation_detector_verdict"):
        value = normalize(payload.get(key))
        if value and not is_not_applicable(value) and not states_no_chapter_format_promise(value):
            return True
    return False


def payload_text(payload: dict[str, Any], keys: tuple[str, ...]) -> str:
    chunks: list[str] = []
    for key in keys:
        value = payload.get(key)
        if isinstance(value, (dict, list, tuple)):
            chunks.append(json.dumps(value, ensure_ascii=False))
        elif value is not None:
            chunks.append(str(value))
    return normalize(" ".join(chunks))


def load_markdown_record(path: Path) -> dict[str, Any]:
    data: dict[str, Any] = {}
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        stripped = line.strip()
        if not stripped.startswith("- ") or ":" not in stripped:
            continue
        key, value = stripped[2:].split(":", 1)
        data[key.strip().replace("-", "_").replace(" ", "_")] = value.strip()
    return data


def load_record(path: Path) -> dict[str, Any]:
    if path.suffix.lower() == ".json":
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("transaction record JSON root must be an object")
        return payload
    return load_markdown_record(path)


def resolve_path(value: object, base: Path) -> Path | None:
    text = str(value or "").strip()
    if not text or normalize(text) in {"none", "n/a", "not-applicable", "missing"}:
        return None
    path = Path(text)
    if not path.is_absolute():
        path = (base.parent / path).resolve()
    return path


def evidence_object(payload: dict[str, Any], name: str) -> dict[str, Any]:
    value = payload.get(name)
    if isinstance(value, dict):
        return value
    result: dict[str, Any] = {}
    for key in ("path", "sha256", "verdict", "final_docx_path", "final_docx_sha256"):
        flat_key = f"{name}_{key}"
        if flat_key in payload:
            result[key] = payload[flat_key]
    if value is not None and "path" not in result:
        result["path"] = value
    return result


def check_path_sha(
    *,
    issues: list[str],
    field_name: str,
    path_value: object,
    sha_value: object,
    record_path: Path,
    require_exists: bool = True,
    require_sha: bool = False,
) -> Path | None:
    resolved = resolve_path(path_value, record_path)
    if resolved is None:
        issues.append(f"{field_name} path is missing")
        return None
    if require_exists and not resolved.exists():
        issues.append(f"{field_name} path does not exist: {resolved}")
        return resolved
    sha_text = str(sha_value or "").strip().lower()
    if require_sha and not sha_text:
        issues.append(f"{field_name} sha256 is missing")
    if require_exists and resolved.exists() and sha_text:
        if len(sha_text) != 64:
            issues.append(f"{field_name} sha256 must be 64 hex characters")
        else:
            actual = sha256_file(resolved).lower()
            if actual != sha_text:
                issues.append(f"{field_name} sha256 does not match path: {resolved}")
    return resolved


def evidence_has_right_edge_metric(payload: dict[str, Any]) -> bool:
    text = json.dumps(payload, ensure_ascii=False).lower()
    if any(token in text for token in ("leader_x0", "leader start", "leader-start", "dotted-leader start")):
        return False
    return any(
        token in text
        for token in ("page_number_right_edge", "page-number right edge", "right_edge", "right-boundary", "right boundary")
    )


def first_payload_path(payload: dict[str, Any], keys: tuple[str, ...], record_path: Path) -> Path | None:
    for key in keys:
        value = payload.get(key)
        if isinstance(value, dict):
            value = value.get("path")
        resolved = resolve_path(value, record_path)
        if resolved is not None:
            return resolved
    return None


def require_existing_payload_path(
    *,
    issues: list[str],
    payload: dict[str, Any],
    keys: tuple[str, ...],
    label: str,
    record_path: Path,
) -> Path | None:
    path = first_payload_path(payload, keys, record_path)
    if path is None:
        issues.append(f"transaction {label} path is missing")
        return None
    if not path.exists():
        issues.append(f"transaction {label} path does not exist: {path}")
    return path


def load_figure_manifest_for_transaction(path: Path) -> tuple[dict[str, Any] | None, str | None]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return None, f"transaction figure asset manifest is unreadable: {path} ({exc})"
    if not isinstance(payload, dict):
        return None, f"transaction figure asset manifest root must be an object: {path}"
    return payload, None


def manifest_path_matches(
    manifest: dict[str, Any],
    keys: tuple[str, ...],
    expected: Path | None,
    record_path: Path,
) -> bool:
    if expected is None:
        return False
    for key in keys:
        resolved = resolve_path(manifest.get(key), record_path)
        if resolved is None:
            continue
        try:
            if resolved.resolve() == expected.resolve():
                return True
        except OSError:
            continue
    return False


def validate_manifest_transaction_binding(
    manifest: dict[str, Any],
    *,
    final_docx: Path | None,
    source_docx: Path | None,
    record_path: Path,
) -> list[str]:
    issues: list[str] = []
    final_keys = ("final_docx_path", "output_docx_path", "reviewed_output_path", "final_manuscript_path")
    source_keys = ("source_docx_path", "original_docx_path", "source_manuscript_path")
    if not manifest_path_matches(manifest, final_keys, final_docx, record_path):
        issues.append("transaction figure asset manifest final_docx_path must match transaction final DOCX")
    replacement_signal = payload_text(
        manifest,
        (
            "mutation_intent",
            "replacement_authorization_scope",
            "explicit_replacement_authorization_scope",
            "authorized_replacement_scope",
        ),
    )
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if isinstance(collection, dict):
            replacement_signal += " " + payload_text(
                collection,
                (
                    "mutation_intent",
                    "replacement_authorization_scope",
                    "explicit_replacement_authorization_scope",
                    "authorized_replacement_scope",
                    "original_media_sha256",
                    "original_asset_sha256",
                    "original_rid",
                    "original_relationship_id",
                ),
            )
    if has_any_token(replacement_signal, {"replace_existing", "replacement", "replace image", "redraw", "original_media_sha256"}):
        if source_docx is None or not source_docx.exists():
            issues.append("transaction figure replacement manifest requires existing source_docx for source-to-final binding")
        elif not manifest_path_matches(manifest, source_keys, source_docx, record_path):
            issues.append("transaction figure replacement manifest source_docx_path must match transaction source DOCX")
    return issues


def docx_has_review_artifacts(path: Path | None) -> bool:
    if path is None or path.suffix.lower() != ".docx" or not path.exists():
        return False
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if any(name.startswith("word/comments") and name.endswith(".xml") for name in names):
                return True
            story_names = [
                name
                for name in names
                if name.startswith("word/")
                and name.endswith(".xml")
                and (
                    name == "word/document.xml"
                    or name.startswith("word/header")
                    or name.startswith("word/footer")
                    or name in {"word/footnotes.xml", "word/endnotes.xml"}
                )
            ]
            for name in story_names:
                text = zf.read(name).decode("utf-8", errors="ignore")
                if "commentRangeStart" in text or "commentReference" in text or "<w:ins" in text or "<w:del" in text:
                    return True
    except (OSError, zipfile.BadZipFile, KeyError):
        return False
    return False


def docx_media_manifest_changed(source_docx: Path | None, final_docx: Path | None) -> bool:
    if source_docx is None or final_docx is None:
        return False
    if not source_docx.exists() or not final_docx.exists():
        return False
    source_media = docx_image_relationship_manifest(source_docx)
    final_media = docx_image_relationship_manifest(final_docx)
    if set(source_media) != set(final_media):
        return True
    for key, source in source_media.items():
        final = final_media.get(key, {})
        if source.get("target", "") != final.get("target", ""):
            return True
        if source.get("sha256", "").lower() != final.get("sha256", "").lower():
            return True
    return False


def docx_drawing_manifest_changed(source_docx: Path | None, final_docx: Path | None) -> bool:
    if source_docx is None or final_docx is None:
        return False
    if not source_docx.exists() or not final_docx.exists():
        return False
    source_drawings = docx_drawing_object_manifest(source_docx)
    final_drawings = docx_drawing_object_manifest(final_docx)
    if Counter(stable_drawing_signature(row) for row in source_drawings.values()) == Counter(
        stable_drawing_signature(row) for row in final_drawings.values()
    ):
        return False
    if set(source_drawings) != set(final_drawings):
        return True
    compared_fields = (
        "drawing_kind",
        "extent_signature",
        "relationship_ids",
        "media_signature",
    )
    for key, source in source_drawings.items():
        final = final_drawings.get(key, {})
        if any(str(source.get(field, "")) != str(final.get(field, "")) for field in compared_fields):
            return True
    return False


def stable_drawing_signature(row: dict[str, Any]) -> tuple[str, ...]:
    fields = (
        "story_part",
        "drawing_kind",
        "extent_signature",
        "relationship_ids",
        "media_signature",
    )
    return tuple(str(row.get(field, "")) for field in fields)


def validate_transaction_record(record_path: Path, expected_final_docx: Path | None = None) -> list[str]:
    issues: list[str] = []
    if not record_path.exists():
        return [f"thesis mutation transaction record not found: {record_path}"]
    try:
        payload = load_record(record_path)
    except Exception as exc:
        return [f"thesis mutation transaction record is unreadable: {record_path} ({exc})"]

    schema = str(payload.get("schema", "")).strip()
    if schema != SCHEMA:
        issues.append(f"transaction schema must be {SCHEMA}: {schema or 'missing'}")

    selected_workflow = normalize(payload.get("selected_workflow") or payload.get("selected_thesis_workflow"))
    transaction_workflow = normalize(payload.get("transaction_workflow") or selected_workflow)
    if selected_workflow not in WORKFLOWS:
        issues.append(f"selected_workflow is not recognized: {selected_workflow or 'missing'}")
    if transaction_workflow not in TRANSACTION_WORKFLOWS:
        issues.append(f"transaction_workflow is not recognized: {transaction_workflow or 'missing'}")

    target_surfaces = as_list(payload.get("target_surfaces") or payload.get("target_surface_ids"))
    if not target_surfaces and transaction_workflow != "audit-only":
        issues.append("transaction target_surfaces must not be empty")
    write_owner = str(payload.get("write_owner") or payload.get("single_write_owner") or "").strip()
    if not write_owner and transaction_workflow != "audit-only":
        issues.append("transaction write_owner is missing")
    operation_text = payload_text(
        payload,
        (
            "selected_workflow",
            "transaction_workflow",
            "target_surfaces",
            "planned_operations",
            "operation_summary",
            "mutation_intent",
            "target_anchor_surface",
            "target_anchor_caption",
            "target_anchor_chapter",
            "insertion_target_surface",
            "locator_strategy",
            "anchor_strategy",
            "comment_revision_scope",
            "teacher_comment_scope",
            "write_scope",
            "helper_strategy",
            "claimed_acceptance_scope",
            "scope_claim",
            "handoff_claim",
        ),
    )
    chapter_format_required = transaction_workflow != "audit-only" and requires_chapter_format_preservation(
        payload,
        target_surfaces,
        operation_text,
    )

    final_docx = check_path_sha(
        issues=issues,
        field_name="final_docx",
        path_value=payload.get("final_docx_path"),
        sha_value=payload.get("final_docx_sha256"),
        record_path=record_path,
        require_exists=transaction_workflow != "audit-only",
        require_sha=transaction_workflow != "audit-only",
    )
    if expected_final_docx is not None and final_docx is not None:
        try:
            if final_docx.resolve() != expected_final_docx.resolve():
                issues.append("transaction final_docx_path does not match the acceptance record rendered DOCX")
        except OSError:
            issues.append("transaction final_docx_path cannot be resolved")

    checked_docx_paths: dict[str, Path | None] = {"final_docx": final_docx}
    for field in ("source_docx_path", "template_docx_path", "review_copy_path"):
        sha_field = field.replace("_path", "_sha256")
        checked_docx_paths[field.removesuffix("_path")] = check_path_sha(
            issues=issues,
            field_name=field.removesuffix("_path"),
            path_value=payload.get(field),
            sha_value=payload.get(sha_field),
            record_path=record_path,
            require_exists=transaction_workflow != "audit-only",
            require_sha=transaction_workflow != "audit-only",
        )
    if transaction_workflow != "audit-only":
        source_checked = checked_docx_paths.get("source_docx")
        if source_checked is not None and final_docx is not None:
            try:
                if source_checked.resolve() == final_docx.resolve():
                    issues.append("transaction source_docx_path and final_docx_path must differ")
            except OSError:
                issues.append("transaction source_docx_path or final_docx_path cannot be resolved")

    if transaction_workflow != "audit-only":
        required_evidence = list(BASE_REQUIRED_EVIDENCE)
        if chapter_format_required:
            required_evidence.append(CHAPTER_FORMAT_PRESERVATION_EVIDENCE)
        for name in required_evidence:
            evidence = evidence_object(payload, name)
            if not evidence:
                issues.append(f"transaction missing evidence object: {name}")
                continue
            check_path_sha(
                issues=issues,
                field_name=name,
                path_value=evidence.get("path"),
                sha_value=evidence.get("sha256"),
                record_path=record_path,
                require_exists=True,
            )
            if not is_pass(evidence.get("verdict")):
                issues.append(f"transaction evidence verdict is not pass: {name}")
            evidence_docx = evidence.get("final_docx_path")
            evidence_sha = evidence.get("final_docx_sha256")
            if not evidence_docx:
                issues.append(f"{name} final_docx_path is missing")
            else:
                resolved_evidence_docx = resolve_path(evidence_docx, record_path)
                if final_docx is not None and resolved_evidence_docx is not None:
                    try:
                        if resolved_evidence_docx.resolve() != final_docx.resolve():
                            issues.append(f"{name} final_docx_path does not match transaction final DOCX")
                    except OSError:
                        issues.append(f"{name} final_docx_path cannot be resolved")
            if not str(evidence_sha or "").strip():
                issues.append(f"{name} final_docx_sha256 is missing")
            elif final_docx is not None and final_docx.exists():
                if str(evidence_sha).strip().lower() != sha256_file(final_docx).lower():
                    issues.append(f"{name} final_docx_sha256 does not match transaction final DOCX")

    detector_verdicts = payload.get("detector_verdicts")
    if isinstance(detector_verdicts, dict):
        if chapter_format_required:
            if CHAPTER_FORMAT_PRESERVATION_DETECTOR not in detector_verdicts:
                issues.append(
                    f"transaction detector_verdicts missing required detector: {CHAPTER_FORMAT_PRESERVATION_DETECTOR}"
                )
            elif not is_pass(detector_verdicts.get(CHAPTER_FORMAT_PRESERVATION_DETECTOR)):
                issues.append(f"transaction detector verdict is not pass: {CHAPTER_FORMAT_PRESERVATION_DETECTOR}")
        for detector, verdict in detector_verdicts.items():
            if detector == CHAPTER_FORMAT_PRESERVATION_DETECTOR and not chapter_format_required:
                continue
            if not is_pass(verdict):
                issues.append(f"transaction detector verdict is not pass: {detector}")
    elif transaction_workflow != "audit-only":
        issues.append("transaction detector_verdicts must be an object")

    if transaction_workflow != "audit-only":
        if chapter_format_required:
            for field in ("format_preservation_promise_verdict", "chapter_format_preservation_detector_verdict"):
                if not is_pass(payload.get(field)):
                    issues.append(f"transaction {field} is not pass")
        if not is_pass(payload.get(NON_TARGET_FORMAT_PRESERVATION_FIELD)):
            issues.append(f"transaction {NON_TARGET_FORMAT_PRESERVATION_FIELD} is not pass")

    transaction_verdict = payload.get("transaction_verdict") or payload.get("final_verdict")
    if transaction_workflow != "audit-only" and not is_pass(transaction_verdict):
        issues.append("transaction final verdict is not pass")
    non_target_verdict = payload.get("non_target_protected_surface_change_verdict")
    if transaction_workflow != "audit-only" and not is_pass(non_target_verdict):
        issues.append("transaction non-target protected surface change verdict is not pass")
    unauthorized = as_list(payload.get("unauthorized_changes") or payload.get("non_target_protected_surface_changes"))
    if unauthorized:
        issues.append("transaction contains unauthorized non-target protected surface changes")

    claim_text = normalize(
        " ".join(
            str(payload.get(key, ""))
            for key in ("claimed_acceptance_scope", "scope_claim", "handoff_claim")
        )
    )
    if transaction_workflow == "local-surface-repair" and has_whole_scope_claim(claim_text):
        issues.append("local-surface transaction cannot claim whole-thesis/template/submission pass")
    local_claim_verdict = payload.get("local_surface_whole_thesis_claim_verdict")
    if transaction_workflow == "local-surface-repair" and not is_pass(local_claim_verdict):
        issues.append("local-surface transaction whole-thesis claim verdict is not pass")

    target_text = normalize(" ".join(target_surfaces))
    if any(surface in target_text for surface in TOC_PAGE_NUMBER_SURFACES):
        render_evidence = evidence_object(payload, "target_surface_render_review")
        regression_evidence = evidence_object(payload, "cross_surface_regression_report")
        if not (evidence_has_right_edge_metric(render_evidence) or evidence_has_right_edge_metric(regression_evidence)):
            issues.append("TOC page-number column transaction evidence must use rendered page-number right-edge metrics")

    protected_sibling_text = payload_text(payload, ("protected_sibling_surfaces",))
    surface_scope_text = normalize(f"{operation_text} {protected_sibling_text}")
    source_docx = first_payload_path(payload, ("source_docx_path",), record_path)
    comment_scope = normalize(payload.get("comment_revision_scope") or payload.get("teacher_comment_scope"))
    comment_driven = has_any_token(operation_text, COMMENT_DRIVEN_TOKENS) or (
        docx_has_review_artifacts(source_docx)
        and comment_scope not in {"", "none", "n/a", "not-applicable", "not-applicable-with-reason"}
    )
    if transaction_workflow != "audit-only":
        if has_fixed_locator_signal(operation_text):
            issues.append("DOCX mutation transaction cannot use fixed paragraph indexes or Word COM paragraph numbers")
        if has_whole_rewrite_signal(operation_text) and selected_workflow != "new-thesis-production":
            issues.append("DOCX mutation transaction cannot use whole-document rewrite/rebuild without new-thesis-production routing")
    if comment_driven and transaction_workflow != "audit-only":
        require_existing_payload_path(
            issues=issues,
            payload=payload,
            keys=("source_review_artifact_inventory_path", "review_artifact_source_inventory_path"),
            label="source review-artifact inventory",
            record_path=record_path,
        )
        require_existing_payload_path(
            issues=issues,
            payload=payload,
            keys=("final_review_artifact_diff_path", "review_artifact_preservation_report"),
            label="final review-artifact diff",
            record_path=record_path,
        )
        if not is_pass(payload.get("review_artifact_preservation_verdict")):
            issues.append("comment-driven transaction review_artifact_preservation_verdict is not pass")
        comment_resolution_ledger_path = require_existing_payload_path(
            issues=issues,
            payload=payload,
            keys=("comment_resolution_ledger_path", "teacher_comment_resolution_ledger_path"),
            label="comment-resolution ledger",
            record_path=record_path,
        )
        if comment_resolution_ledger_path is not None and comment_resolution_ledger_path.exists() and final_docx is not None:
            issues.extend(
                validate_comment_resolution_ledger(
                    comment_resolution_ledger_path,
                    final_docx=final_docx,
                    source_docx=source_docx,
                    assert_all_resolved=has_all_comments_claim(
                        payload.get("claimed_acceptance_scope"),
                        payload.get("scope_claim"),
                        payload.get("handoff_claim"),
                        payload.get("comment_resolution_scope"),
                    ),
                )
            )
        if has_all_comments_claim(
            payload.get("claimed_acceptance_scope"),
            payload.get("scope_claim"),
            payload.get("handoff_claim"),
            payload.get("comment_resolution_scope"),
        ) and comment_resolution_ledger_path is None:
            issues.append("all-comments-resolved transaction claim requires a comment-resolution ledger")

    media_diff_mutation = docx_media_manifest_changed(source_docx, final_docx)
    drawing_diff_mutation = docx_drawing_manifest_changed(source_docx, final_docx)
    image_intent = has_image_mutation_intent(payload)
    image_mutation = image_intent or media_diff_mutation or drawing_diff_mutation
    protected_image_target = image_mutation and has_any_token(surface_scope_text, PROTECTED_IMAGE_TARGET_TOKENS)
    protected_surface_image_authorized = is_pass(
        payload.get("official_template_protected_image_authorization_verdict")
        or payload.get("protected_surface_image_authorization_verdict")
    )
    if image_mutation and not protected_surface_image_authorized and not is_pass(payload.get("target_anchor_not_protected_surface_verdict")):
        issues.append("image insertion/replacement target must prove it is outside TOC/front matter protected surfaces")
    if image_mutation and transaction_workflow != "audit-only":
        if media_diff_mutation and not image_intent:
            issues.append("source-to-final DOCX media relationships changed; transaction must route as an image mutation")
        if drawing_diff_mutation and not image_intent:
            issues.append("source-to-final DOCX drawing objects changed; transaction must route as an image mutation")
        figure_manifest_path = require_existing_payload_path(
            issues=issues,
            payload=payload,
            keys=("figure_manifest_path", "figure_asset_manifest_path"),
            label="figure asset manifest",
            record_path=record_path,
        )
        if figure_manifest_path is not None and figure_manifest_path.exists():
            manifest, manifest_error = load_figure_manifest_for_transaction(figure_manifest_path)
            if manifest_error:
                issues.append(manifest_error)
            elif manifest is not None:
                issues.extend(
                    validate_manifest_transaction_binding(
                        manifest,
                        final_docx=final_docx,
                        source_docx=source_docx,
                        record_path=record_path,
                    )
                )
                for issue in validate_figure_manifest(
                    manifest,
                    final_docx=final_docx,
                    source_docx=source_docx,
                    manifest_path=figure_manifest_path,
                ):
                    issues.append(f"transaction figure asset manifest contract failed: {issue}")
        if not is_pass(payload.get("figure_contract_verdict") or payload.get("figure_manifest_contract_verdict")):
            issues.append("image mutation transaction figure contract verdict is not pass")
        if not is_pass(payload.get("figure_anchor_location_verdict")):
            issues.append("image mutation transaction figure_anchor_location_verdict is not pass")
        if not is_pass(payload.get("caption_asset_binding_verdict")):
            issues.append("image mutation transaction caption_asset_binding_verdict is not pass")

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a thesis mutation transaction record.")
    parser.add_argument("--record", required=True, help="Transaction record JSON or Markdown path")
    parser.add_argument("--final-docx", help="Expected final DOCX path")
    parser.add_argument("--json", action="store_true", help="Print JSON result")
    parser.add_argument("--json-output", help="Write JSON result as UTF-8 without relying on shell redirection")
    args = parser.parse_args()

    record_path = Path(args.record).resolve()
    final_docx = Path(args.final_docx).resolve() if args.final_docx else None
    issues = validate_transaction_record(record_path, expected_final_docx=final_docx)
    payload = {"ok": not issues, "issues": issues}
    if args.json_output:
        json_output = Path(args.json_output).resolve()
        json_output.parent.mkdir(parents=True, exist_ok=True)
        json_output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    elif issues:
        print("THESIS MUTATION TRANSACTION FAILED")
        for issue in issues:
            print(f"- {issue}")
    else:
        print("THESIS MUTATION TRANSACTION PASSED")
        print(f"- record: {record_path}")
    return 1 if issues else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raise SystemExit(main())
