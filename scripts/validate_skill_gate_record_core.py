"""Core record helpers for validate_skill_gate."""

from __future__ import annotations

__all__ = [
    "check_docx_body_style_audit_report",
    "check_thesis_citation_audit_report",
    "check_humanizer_evidence_record",
    "validate_gate_single_prefix",
    "detect_format_repair_surfaces",
    "detect_review_evidence_surfaces",
    "check_docx_font_audit_report",
    "check_docx_font_color_audit_report",
    "check_docx_whole_format_gate_report",
    "format_repair_task_touches_surface",
    "validate_template_lock_fields",
    "check_project_local_helper_preflight_report",
    "validate_project_local_helper_preflight_fields",
]

import hashlib
import json
import re
import shlex
from pathlib import Path

try:
    from .validate_skill_gate_registry_core import BODY_STYLE_AUDIT_SCHEMA, CITATION_AUDIT_SCHEMA, DOCX_FONT_AUDIT_SCHEMA, EXPLICIT_VALUES, FORMAT_REPAIR_TASK_SCHEMA, HUMANIZER_EVIDENCE_SCHEMA, TEXT_EVIDENCE_EXTENSIONS
    from .validate_skill_gate_registry_records import (
        FORMAT_REPAIR_TASK_SURFACE_DETECTOR_SPECS,
        GATE_SINGLE_PREFIX_POLICY,
        REVIEW_EVIDENCE_SURFACE_DETECTOR_SPECS,
        REVIEW_EVIDENCE_SURFACE_TEXT_PREFIXES,
    )
    from .validate_skill_gate_utils import (
        build_raw_context,
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        is_explicit_or_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )
    from .audit_thesis_citations import audit_docx as audit_body_citations, build_report as build_citation_audit_report
except ImportError:
    from validate_skill_gate_registry_core import BODY_STYLE_AUDIT_SCHEMA, CITATION_AUDIT_SCHEMA, DOCX_FONT_AUDIT_SCHEMA, EXPLICIT_VALUES, FORMAT_REPAIR_TASK_SCHEMA, HUMANIZER_EVIDENCE_SCHEMA, TEXT_EVIDENCE_EXTENSIONS
    from validate_skill_gate_registry_records import (
        FORMAT_REPAIR_TASK_SURFACE_DETECTOR_SPECS,
        GATE_SINGLE_PREFIX_POLICY,
        REVIEW_EVIDENCE_SURFACE_DETECTOR_SPECS,
        REVIEW_EVIDENCE_SURFACE_TEXT_PREFIXES,
    )
    from validate_skill_gate_utils import (
        build_raw_context,
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        is_explicit_or_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )
    from audit_thesis_citations import audit_docx as audit_body_citations, build_report as build_citation_audit_report


ALLOWED_TEMPLATE_SOURCE_TYPES = {
    "project-auto-discovered",
    "user-provided",
    "teacher-approved-sample",
}
TEMPLATE_SELECTED_BEFORE_VALUES = {
    "yes",
    "locked-before-mutation",
    "selected-before-mutation",
    "locked before mutation",
    "selected before mutation",
}
TEMPLATE_PASS_VALUES = {"pass", "passed"}
PROFILE_BEFORE_MUTATION_VALUES = {"yes", "locked-before-mutation", "generated-before-mutation"}
PATH_ENCODING_PASS_VALUES = {"pass", "passed", "clean", "utf-8-safe", "utf8-safe"}
MOJIBAKE_PATH_TOKEN_ESCAPES = (
    "\\u6924\\u572d\\u6d30",
    "\\u9369\\u8f70\\u7c2c",
    "\\u93c8\\u54c4\\u6ad2",
    "\\u7459\\u55da",
    "\\u59e3\\u66da\\u7b1f",
    "\\u7481\\u6350",
    "\\u93b8\\u590b\\u5270",
)
MOJIBAKE_PATH_TOKENS = tuple(
    token.encode("ascii").decode("unicode_escape")
    for token in MOJIBAKE_PATH_TOKEN_ESCAPES
)
THESIS_MODES = {"thesis-only", "format-repair-only", "program-plus-thesis"}
PROJECT_LOCAL_RISK_TOKENS = {"failed", "contaminated", "thick", "risky project-local", "risk detected"}
PROJECT_LOCAL_BLOCKING_DISPOSITIONS = {"audit-only", "clean-source-restart-required"}
PROJECT_LOCAL_COMPLETED_DISPOSITIONS = {"clean-source-restart-completed"}
PROJECT_LOCAL_RESTART_REQUIRED_VALUES = {"yes", "required", "clean-source-restart-required"}
PROJECT_LOCAL_RESTART_COMPLETED_VALUES = {"completed", "clean-source-restart-completed"}


def project_local_risk_count(value: str) -> int | None:
    text = normalize(str(value))
    if not re.fullmatch(r"\d+", text):
        return None
    return int(text)


def _extract_markdown_field(text: str, label: str) -> str:
    prefix = f"- {label.lower()}:"
    for line in text.splitlines():
        if normalize(line).lower().startswith(prefix):
            return raw_line_value(line)
    return ""


def _extract_gate_record_arg(command_text: str) -> str:
    try:
        parts = shlex.split(command_text, posix=False)
    except ValueError:
        parts = command_text.split()
    for index, part in enumerate(parts):
        token = part.strip().strip("`").strip('"').strip("'")
        if token == "--gate-record" and index + 1 < len(parts):
            return parts[index + 1].strip().strip("`").strip('"').strip("'")
        if token.startswith("--gate-record="):
            return token.split("=", 1)[1].strip().strip("`").strip('"').strip("'")
    match = re.search(r"--gate-record(?:=|\s+)(\"[^\"]+\"|'[^']+'|`[^`]+`|\S+)", command_text)
    if match:
        return match.group(1).strip().strip("`").strip('"').strip("'")
    return ""


def _as_int(value: str) -> int | None:
    value = normalize(value)
    if not re.fullmatch(r"-?\d+", value):
        return None
    return int(value)


def check_project_local_helper_preflight_report(
    report_path: Path,
    *,
    expected_risk_count: int | None = None,
    expected_project_root: Path | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues
    try:
        text = report_path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        return [f"project-local helper preflight report is not valid UTF-8: {report_path} ({exc})"]
    lowered = text.lower()
    schema = _extract_markdown_field(text, "report schema")
    if schema != "graduation-project-builder.project-local-helper-preflight.v2":
        issues.append(f"project-local helper preflight report must use schema v2: {report_path}")
    generated_at = _extract_markdown_field(text, "generated at UTC")
    generated_unix = _extract_markdown_field(text, "generated at unix")
    if not generated_at or not re.search(r"\d{4}-\d{2}-\d{2}T", generated_at):
        issues.append(f"project-local helper preflight report lacks generated-at timestamp: {report_path}")
    try:
        if not generated_unix or float(generated_unix) <= 0:
            issues.append(f"project-local helper preflight report lacks positive generated-at unix timestamp: {report_path}")
    except ValueError:
        issues.append(f"project-local helper preflight report generated-at unix timestamp is not numeric: {report_path}")
    report_root_value = _extract_markdown_field(text, "project root")
    if not report_root_value:
        issues.append(f"project-local helper preflight report lacks project root: {report_path}")
    else:
        report_root = resolve_record_path(report_root_value, report_path)
        if expected_project_root is not None:
            try:
                if report_root.resolve() != expected_project_root.resolve():
                    issues.append(
                        "project-local helper preflight report project root does not match the acceptance/template lock "
                        f"root: report={report_root} expected={expected_project_root}"
                    )
            except OSError:
                issues.append(f"project-local helper preflight report project root cannot be resolved: {report_root}")
    if "scan_project_local_thesis_helpers.py" not in lowered and "project-local thesis helper" not in lowered:
        issues.append(f"project-local helper preflight report does not name the scanner: {report_path}")
    scanner_command = _extract_markdown_field(text, "scanner command")
    if "scan_project_local_thesis_helpers.py" not in scanner_command or "--project-root" not in scanner_command:
        issues.append(f"project-local helper preflight report lacks scanner command with --project-root: {report_path}")
    exit_status = _as_int(_extract_markdown_field(text, "scanner exit status"))
    if exit_status is None:
        issues.append(f"project-local helper preflight report lacks numeric scanner exit status: {report_path}")
    match = re.search(r"risky script count\s*:\s*(\d+)", lowered)
    if not match:
        issues.append(f"project-local helper preflight report lacks risky script count: {report_path}")
        return issues
    actual_count = int(match.group(1))
    if expected_risk_count is not None and actual_count != expected_risk_count:
        issues.append(
            "project-local helper preflight report risky script count "
            f"{actual_count} does not match acceptance field {expected_risk_count}: {report_path}"
        )
    if actual_count > 0 and not contains_any(lowered, PROJECT_LOCAL_RISK_TOKENS):
        issues.append(f"project-local helper preflight report has risks but no failure/risk summary: {report_path}")
    if actual_count == 0 and not contains_any(lowered, {"passed", "clean", "no risky", "none"}):
        issues.append(f"project-local helper preflight report has zero risks but no clean/pass summary: {report_path}")
    if exit_status is not None:
        if actual_count > 0 and exit_status == 0:
            issues.append(f"project-local helper preflight report exit status must be non-zero when risks are found: {report_path}")
        if actual_count == 0 and exit_status != 0:
            issues.append(f"project-local helper preflight report exit status must be zero when no risks are found: {report_path}")
    return issues


def validate_project_local_helper_preflight_fields(
    *,
    record_path: Path,
    values: dict[str, str],
    raw_values: dict[str, str],
    record_kind: str,
    task_mode: str,
    risk_summary: str = "",
) -> list[str]:
    issues: list[str] = []
    if task_mode not in THESIS_MODES:
        return issues

    preflight_summary = values.get("- project-local helper script preflight summary:", "")
    preflight_report = values.get("- project-local helper preflight report path:", "")
    risk_count_text = values.get("- project-local helper risk count:", "")
    helper_disposition = values.get("- project-local helper disposition:", "")
    canonical_restart = values.get("- canonical source restart required?:", "")
    clean_source_restart_path = values.get("- clean-source restart source path:", "")
    contaminated_disposition = values.get("- contaminated-baseline disposition:", "")
    contaminated_intermediate_used = values.get("- contaminated intermediate source used?:", "")
    source_retention_manifest_path = values.get("- source retention manifest path:", "")

    if not is_explicit(preflight_summary) or preflight_summary in EXPLICIT_VALUES:
        issues.append(f"{record_kind} must explicitly record a project-local helper script preflight summary")
    if not is_explicit(preflight_report) or is_explicit_none(preflight_report):
        issues.append(f"{record_kind} must name a project-local helper preflight report path")
    if not is_explicit(helper_disposition) or helper_disposition in EXPLICIT_VALUES:
        issues.append(f"{record_kind} must explicitly record a project-local helper disposition")
    if not is_explicit(canonical_restart) or (canonical_restart in EXPLICIT_VALUES and canonical_restart != "no"):
        issues.append(f"{record_kind} must explicitly record whether canonical source restart is required")

    risk_count = project_local_risk_count(risk_count_text)
    if risk_count is None:
        issues.append(f"{record_kind} project-local helper risk count must be a non-negative integer")

    expected_project_root: Path | None = None
    source_type = values.get("- active template source type:", "")
    raw_preflight_root = raw_values.get("- project-local helper preflight project root:", "")
    if raw_preflight_root and not is_explicit_none(normalize(raw_preflight_root)):
        expected_project_root = resolve_record_path(raw_preflight_root, record_path)
    elif source_type == "project-auto-discovered":
        raw_discovery_root = raw_values.get("- project template discovery root:", "")
        if raw_discovery_root and not is_explicit_none(normalize(raw_discovery_root)):
            expected_project_root = resolve_record_path(raw_discovery_root, record_path)

    if is_explicit(preflight_report) and not is_explicit_none(preflight_report):
        raw_report = raw_values.get("- project-local helper preflight report path:", preflight_report)
        for raw_path in split_path_values(raw_report):
            resolved = resolve_record_path(raw_path, record_path)
            issues.extend(
                check_project_local_helper_preflight_report(
                    resolved,
                    expected_risk_count=risk_count,
                    expected_project_root=expected_project_root,
                )
            )

    risk_detected = (
        (risk_count or 0) > 0
        or contains_any(preflight_summary, PROJECT_LOCAL_RISK_TOKENS)
        or contains_any(risk_summary, PROJECT_LOCAL_RISK_TOKENS)
    )
    disposition_value = normalize(helper_disposition).lower()
    restart_value = normalize(canonical_restart).lower()
    completed_restart = (
        disposition_value in PROJECT_LOCAL_COMPLETED_DISPOSITIONS
        or restart_value in PROJECT_LOCAL_RESTART_COMPLETED_VALUES
    )
    if risk_detected:
        if completed_restart:
            if disposition_value not in PROJECT_LOCAL_COMPLETED_DISPOSITIONS:
                issues.append(
                    f"{record_kind} project-local helper disposition must be clean-source-restart-completed when canonical restart is marked completed"
                )
            if restart_value not in PROJECT_LOCAL_RESTART_COMPLETED_VALUES:
                issues.append(
                    f"{record_kind} canonical source restart status must be completed when risky local thesis helpers were bypassed by a completed clean-source restart"
                )
            if not is_explicit(clean_source_restart_path) or is_explicit_none(clean_source_restart_path):
                issues.append(f"{record_kind} must name the clean-source restart source path when restart is completed")
            else:
                clean_source_path = resolve_record_path(
                    raw_values.get("- clean-source restart source path:", clean_source_restart_path),
                    record_path,
                )
                issues.extend(validate_existing_path(clean_source_path, require_nonempty_file=True))
            if not contains_any(contaminated_disposition, {"completed", "clean source", "clean-source", "not used"}):
                issues.append(
                    f"{record_kind} contaminated-baseline disposition must state that clean-source restart completed and contaminated helpers were not used"
                )
            if contaminated_intermediate_used and normalize(contaminated_intermediate_used).lower() not in {"no", "none"}:
                issues.append(f"{record_kind} contaminated intermediate source must not be used after clean-source restart completion")
            if not is_explicit(source_retention_manifest_path) or is_explicit_none(source_retention_manifest_path):
                issues.append(f"{record_kind} must name a source retention manifest when clean-source restart is completed")
            else:
                manifest_path = resolve_record_path(
                    raw_values.get("- source retention manifest path:", source_retention_manifest_path),
                    record_path,
                )
                manifest_issues = validate_existing_path(manifest_path, require_nonempty_file=True)
                issues.extend(manifest_issues)
                if not manifest_issues:
                    try:
                        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                    except Exception as exc:
                        issues.append(f"{record_kind} source retention manifest is not valid JSON: {manifest_path} ({exc})")
                    else:
                        if not isinstance(manifest, dict):
                            issues.append(f"{record_kind} source retention manifest root must be an object: {manifest_path}")
                        else:
                            for field in ("schema", "source", "source_sha256", "source_class", "output", "output_sha256", "verdict", "rebuild_class"):
                                if not str(manifest.get(field) or "").strip():
                                    issues.append(f"{record_kind} source retention manifest lacks {field}: {manifest_path}")
                            if contains_any(str(manifest.get("verdict") or ""), {"failed", "fail", "lost", "missing", "blocked"}):
                                issues.append(f"{record_kind} source retention manifest verdict is blocking: {manifest_path}")
        elif disposition_value not in PROJECT_LOCAL_BLOCKING_DISPOSITIONS:
            issues.append(
                f"{record_kind} project-local helper disposition must be audit-only or clean-source-restart-required when risky local thesis helpers are detected"
            )
        if not completed_restart and restart_value not in PROJECT_LOCAL_RESTART_REQUIRED_VALUES:
            issues.append(
                f"{record_kind} canonical source restart must be required when risky local thesis helpers are detected"
            )
        if not completed_restart:
            issues.append(f"{record_kind} detected risky project-local thesis helper scripts; thesis handoff must stay audit-only or restart from a clean source")
    elif risk_count == 0:
        if disposition_value in PROJECT_LOCAL_BLOCKING_DISPOSITIONS or "blocked" in disposition_value:
            issues.append(f"{record_kind} project-local helper disposition conflicts with zero scanner risks")
        if restart_value in PROJECT_LOCAL_RESTART_REQUIRED_VALUES:
            issues.append(f"{record_kind} canonical source restart requirement conflicts with zero scanner risks")

    return issues


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _resolve_optional_path(raw_values: dict[str, str], prefix: str, anchor_path: Path) -> Path | None:
    value = raw_values.get(prefix, "").strip()
    if not value or is_explicit_none(normalize(value)):
        return None
    paths = split_path_values(value)
    if not paths:
        return None
    return resolve_record_path(paths[0], anchor_path)


def _load_template_profile(profile_path: Path) -> tuple[dict[str, object] | None, str | None]:
    try:
        payload = json.loads(profile_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return None, f"active template profile is not valid JSON: {profile_path} ({exc})"
    if not isinstance(payload, dict):
        return None, f"active template profile root must be an object: {profile_path}"
    return payload, None


def _load_template_discovery_report(report_path: Path) -> tuple[dict[str, object] | None, str | None]:
    try:
        payload = json.loads(report_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return None, f"template discovery report is not valid JSON: {report_path} ({exc})"
    if not isinstance(payload, dict):
        return None, f"template discovery report root must be an object: {report_path}"
    return payload, None


def _profile_template_path(profile: dict[str, object]) -> str:
    for key in ("template_docx", "template_path", "source_template_path", "selected_template_path"):
        value = profile.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _profile_template_fingerprint(profile: dict[str, object]) -> str:
    for key in (
        "template_fingerprint",
        "template_sha256",
        "sha256",
        "fingerprint",
        "selected_template_fingerprint",
    ):
        value = profile.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip().lower()
    return ""


def _validate_lock_report_metadata(
    payload: dict[str, object],
    *,
    path: Path,
    expected_stage: str,
    record_kind: str,
    label: str,
) -> list[str]:
    issues: list[str] = []
    if payload.get("generation_stage") != expected_stage:
        issues.append(
            f"{record_kind} {label} must record generation_stage={expected_stage}: {path}"
        )
    generated_at_utc = str(payload.get("generated_at_utc") or "").strip()
    if not re.search(r"\d{4}-\d{2}-\d{2}T", generated_at_utc):
        issues.append(f"{record_kind} {label} lacks generated_at_utc timestamp: {path}")
    generated_at_unix = payload.get("generated_at_unix")
    try:
        if float(str(generated_at_unix)) <= 0:
            issues.append(f"{record_kind} {label} generated_at_unix must be positive: {path}")
    except (TypeError, ValueError):
        issues.append(f"{record_kind} {label} lacks numeric generated_at_unix timestamp: {path}")
    generator = str(payload.get("generator") or "").strip()
    if "scripts/" not in generator and not generator.endswith(".py"):
        issues.append(f"{record_kind} {label} must name the canonical generator script: {path}")
    return issues


def validate_template_lock_fields(
    *,
    record_path: Path,
    values: dict[str, str],
    raw_values: dict[str, str],
    record_kind: str,
    explicit_source_text: str,
) -> list[str]:
    issues: list[str] = []
    source_type = values.get("- active template source type:", "")
    selected_before = values.get("- active template selected before mutation?:", "")
    alignment_verdict = values.get("- template alignment verdict:", "")
    fingerprint = values.get("- active template fingerprint:", "")
    profile_command = values.get("- template profile generation command:", "")
    profile_before = values.get("- template profile generated before mutation?:", "")
    path_encoding = values.get("- locked path encoding verdict:", "")
    contaminated_used = values.get("- contaminated intermediate source used?:", "")
    contaminated_disposition = values.get("- contaminated intermediate source disposition:", "")
    known_caveats = values.get("- known caveats:", "")

    if source_type not in ALLOWED_TEMPLATE_SOURCE_TYPES:
        issues.append(
            f"{record_kind} active template source type must be one of "
            f"{', '.join(sorted(ALLOWED_TEMPLATE_SOURCE_TYPES))}: {source_type or 'blank'}"
        )

    if selected_before.lower() not in TEMPLATE_SELECTED_BEFORE_VALUES:
        issues.append(
            f"{record_kind} active template selected before mutation must be yes/locked-before-mutation: {selected_before or 'blank'}"
        )

    if not alignment_verdict or is_explicit_none(alignment_verdict):
        issues.append(f"{record_kind} template alignment verdict must be pass")
    elif alignment_verdict.lower() not in TEMPLATE_PASS_VALUES:
        issues.append(f"{record_kind} template alignment verdict must be pass: {alignment_verdict}")

    active_template = _resolve_optional_path(raw_values, "- active template path lock:", record_path)
    profile_path = _resolve_optional_path(raw_values, "- active template profile path:", record_path)
    discovery_report_path = _resolve_optional_path(raw_values, "- template discovery report path:", record_path)
    discovery_root = _resolve_optional_path(raw_values, "- project template discovery root:", record_path)
    candidate_paths = [
        resolve_record_path(raw_path, record_path)
        for raw_path in split_path_values(raw_values.get("- discovered candidate template paths:", ""))
        if not is_explicit_none(normalize(raw_path))
    ]

    if active_template is None:
        issues.append(f"{record_kind} must lock an active template path before mutation")
    else:
        issues.extend(validate_existing_path(active_template, require_nonempty_file=True))
        if active_template.suffix.lower() not in {".docx", ".doc"}:
            issues.append(f"{record_kind} active template path must be a Word template/document file: {active_template}")

    if profile_path is None:
        issues.append(f"{record_kind} must provide an active template profile path")
    else:
        issues.extend(validate_existing_path(profile_path, require_nonempty_file=True))

    if discovery_report_path is None:
        issues.append(f"{record_kind} must provide a template discovery report path")
    else:
        issues.extend(validate_existing_path(discovery_report_path, require_nonempty_file=True))

    if not profile_command or "thesis_template_profile" not in profile_command:
        issues.append(f"{record_kind} must record the thesis_template_profile command/equivalent used before mutation")

    if profile_before.lower() not in PROFILE_BEFORE_MUTATION_VALUES:
        issues.append(
            f"{record_kind} template profile generated before mutation must be yes/generated-before-mutation: {profile_before or 'blank'}"
        )

    if path_encoding.lower() not in PATH_ENCODING_PASS_VALUES:
        issues.append(f"{record_kind} locked path encoding verdict must be pass/utf-8-safe: {path_encoding or 'blank'}")

    for prefix, raw_value in raw_values.items():
        if "path" in prefix.lower() and contains_any(raw_value, MOJIBAKE_PATH_TOKENS):
            issues.append(f"{record_kind} path field appears mojibake-corrupted and is not a valid lock: {prefix}")

    if contaminated_used.lower() not in {"no", "none", "not-applicable", "n/a"}:
        issues.append(f"{record_kind} must not use a contaminated intermediate manuscript as a source: {contaminated_used or 'blank'}")
    if contaminated_used.lower() in {"yes", "true", "used"} and not contains_any(
        contaminated_disposition,
        {"content reference only", "audit-only", "restart", "clean source", "blocked"},
    ):
        issues.append(f"{record_kind} contaminated intermediate source disposition is unsafe or missing")
    if known_caveats and not is_explicit_none(known_caveats):
        issues.append(f"{record_kind} known caveats must be none before pass: {known_caveats}")

    if not fingerprint or is_explicit_none(fingerprint):
        issues.append(f"{record_kind} active template fingerprint must be a sha256 value")
    elif len(fingerprint) != 64 or any(ch not in "0123456789abcdefABCDEF" for ch in fingerprint):
        issues.append(f"{record_kind} active template fingerprint must be a 64-hex sha256 value: {fingerprint}")
    elif active_template is not None and active_template.exists():
        actual = sha256_file(active_template)
        if actual.lower() != fingerprint.lower():
            issues.append(
                f"{record_kind} active template fingerprint does not match active template path: {active_template}"
            )

    if source_type == "project-auto-discovered":
        if discovery_root is None:
            issues.append(f"{record_kind} project-auto-discovered template requires a project template discovery root")
        else:
            issues.extend(validate_existing_path(discovery_root, require_nonempty_file=False))
        if not candidate_paths:
            issues.append(f"{record_kind} project-auto-discovered template requires discovered candidate template paths")
        elif active_template is not None:
            active_resolved = active_template.resolve()
            candidate_resolved = {candidate.resolve() for candidate in candidate_paths}
            if active_resolved not in candidate_resolved:
                issues.append(
                    f"{record_kind} active template path lock must be one of discovered candidate template paths"
                )
        if discovery_root is not None and active_template is not None and discovery_root.exists():
            try:
                active_template.resolve().relative_to(discovery_root.resolve())
            except ValueError:
                issues.append(
                    f"{record_kind} project-auto-discovered active template must stay under discovery root"
                )
    elif source_type in {"user-provided", "teacher-approved-sample"}:
        source_text = normalize(explicit_source_text).lower()
        if not contains_any(source_text, {"user", "explicit", "provided", "approved", "teacher", "sample"}):
            issues.append(
                f"{record_kind} {source_type} template source requires an explicit user/teacher source note"
            )

    if profile_path is not None and profile_path.exists() and active_template is not None:
        profile, profile_error = _load_template_profile(profile_path)
        if profile_error:
            issues.append(f"{record_kind} {profile_error}")
        elif profile is not None:
            issues.extend(
                _validate_lock_report_metadata(
                    profile,
                    path=profile_path,
                    expected_stage="pre-mutation-template-profile-lock",
                    record_kind=record_kind,
                    label="active template profile",
                )
            )
            profile_template = _profile_template_path(profile)
            if not profile_template:
                issues.append(f"{record_kind} active template profile must record template_docx/template_path")
            else:
                profile_template_path = resolve_record_path(profile_template, profile_path)
                if profile_template_path.resolve() != active_template.resolve():
                    issues.append(
                        f"{record_kind} active template profile path does not match active template path lock"
                    )
            profile_fingerprint = _profile_template_fingerprint(profile)
            if not profile_fingerprint:
                issues.append(f"{record_kind} active template profile must record template_fingerprint/sha256")
            elif fingerprint and profile_fingerprint.lower() != fingerprint.lower():
                issues.append(
                    f"{record_kind} active template profile fingerprint does not match active template fingerprint"
                )

    if discovery_report_path is not None and discovery_report_path.exists() and active_template is not None:
        report, report_error = _load_template_discovery_report(discovery_report_path)
        if report_error:
            issues.append(f"{record_kind} {report_error}")
        elif report is not None:
            issues.extend(
                _validate_lock_report_metadata(
                    report,
                    path=discovery_report_path,
                    expected_stage="pre-mutation-template-lock",
                    record_kind=record_kind,
                    label="template discovery report",
                )
            )
            selected_path = report.get("selected_template_path")
            selected_fingerprint = report.get("selected_template_fingerprint")
            if not isinstance(selected_path, str) or not selected_path.strip():
                issues.append(f"{record_kind} template discovery report must record selected_template_path")
            else:
                selected_resolved = resolve_record_path(selected_path, discovery_report_path)
                if selected_resolved.resolve() != active_template.resolve():
                    issues.append(f"{record_kind} template discovery report selected path does not match active template path")
            if not isinstance(selected_fingerprint, str) or not selected_fingerprint.strip():
                issues.append(f"{record_kind} template discovery report must record selected_template_fingerprint")
            elif fingerprint and selected_fingerprint.lower() != fingerprint.lower():
                issues.append(f"{record_kind} template discovery report fingerprint does not match active template fingerprint")
            if source_type == "project-auto-discovered":
                authority_class = str(report.get("selected_template_authority_class") or "").strip().lower()
                if authority_class not in {"official-format-spec", "project-template-candidate"}:
                    issues.append(
                        f"{record_kind} project-auto-discovered template authority class is too weak: {authority_class or 'blank'}"
                    )
                for candidate in report.get("candidates") or []:
                    if not isinstance(candidate, dict):
                        continue
                    candidate_path = str(candidate.get("path") or "")
                    if not candidate_path:
                        continue
                    resolved_candidate = resolve_record_path(candidate_path, discovery_report_path)
                    if resolved_candidate.resolve() != active_template.resolve():
                        continue
                    if bool(candidate.get("under_codex")):
                        issues.append(
                            f"{record_kind} project-auto-discovered active template must not resolve to a .codex candidate: {resolved_candidate}"
                        )
                    break

    return issues


def validate_gate_single_prefix(
    *,
    prefix: str,
    line: str,
    value: str,
    task_mode: str,
    paper_only_required: bool,
    record_lines: list[str],
    record_path: Path | None = None,
) -> list[str]:
    issues: list[str] = []
    policy = GATE_SINGLE_PREFIX_POLICY.get(prefix, "explicit_or_none")

    if policy == "explicit_nonempty":
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"gate record field must not be empty: {normalize(line)}")
        return issues

    if policy == "format_repair_nonempty":
        if task_mode == "format-repair-only":
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record field must not be empty for format-repair-only: {normalize(line)}")
        elif not is_explicit_or_none(value):
            issues.append(f"gate record field is incomplete: {normalize(line)}")
        return issues

    if policy == "explicit_none":
        if not is_explicit_none(value):
            issues.append(f"gate record still has failed reasons: {normalize(line)}")
        return issues

    if policy.startswith("exact_value:"):
        expected = policy.split(":", 1)[1]
        if value != expected:
            label = prefix.removeprefix("- ").removesuffix(":")
            issues.append(f"{label} is not {expected}: {normalize(line)}")
        return issues

    if policy.startswith("command:"):
        required_token = policy.split(":", 1)[1]
        normalized_line = normalize(line)
        if required_token not in normalized_line:
            label = prefix.removeprefix("- ").removesuffix(":")
            issues.append(f"{label} does not name the expected command token '{required_token}': {normalize(line)}")
        if prefix == "- validation command:" and "--gate-record" not in normalized_line:
            issues.append(f"validation command must validate the exact acceptance record with --gate-record: {normalize(line)}")
        if prefix == "- validation command:" and "--gate-record" in normalized_line and record_path is not None:
            gate_record_arg = _extract_gate_record_arg(raw_line_value(line))
            if not gate_record_arg:
                issues.append(f"validation command must name the exact acceptance record after --gate-record: {normalize(line)}")
            else:
                referenced = resolve_record_path(gate_record_arg, record_path)
                try:
                    if referenced.resolve() != record_path.resolve():
                        issues.append(
                            "validation command --gate-record must point to this exact acceptance record: "
                            f"{referenced}"
                        )
                except OSError:
                    issues.append(
                        "validation command --gate-record could not be resolved to this exact acceptance record: "
                        f"{gate_record_arg}"
                    )
        return issues

    if policy == "paper_only_evidence":
        if paper_only_required:
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(
                    f"gate record field must not be empty when paper-only literature is enabled: {normalize(line)}"
                )
        elif not is_explicit_or_none(value):
            issues.append(f"gate record field is incomplete: {normalize(line)}")
        return issues

    if policy == "thesis_evidence":
        if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}:
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record field must not be empty for thesis work: {normalize(line)}")
        elif not is_explicit_or_none(value):
            issues.append(f"gate record field is incomplete: {normalize(line)}")
        return issues

    if policy == "humanizer_evidence":
        if task_mode in {"thesis-only", "program-plus-thesis"}:
            route = parse_line_value(find_lines_with_prefix(record_lines, "- humanizer route decision:")[0])
            if route != "none":
                if not is_explicit(value) or value in EXPLICIT_VALUES:
                    issues.append(
                        f"gate record field must not be empty when a humanizer route is active: {normalize(line)}"
                    )
            elif not is_explicit_or_none(value):
                issues.append(f"gate record field is incomplete: {normalize(line)}")
        elif task_mode == "format-repair-only":
            if value != "none":
                issues.append(f"format-repair-only gate record must keep humanizer evidence path as none: {normalize(line)}")
        elif not is_explicit_or_none(value):
            issues.append(f"gate record field is incomplete: {normalize(line)}")
        return issues

    if not is_explicit_or_none(value):
        issues.append(f"gate record field is incomplete: {normalize(line)}")
    return issues


def _humanizer_language_kind(value: str) -> str:
    if contains_any(value, {"zh", "chinese", "\u4e2d\u6587", "\u6c49\u8bed"}):
        return "zh"
    if contains_any(value, {"en", "english", "\u82f1\u6587", "\u82f1\u8bed"}):
        return "en"
    if contains_any(value, {"zh", "chinese", "中文", "汉语"}):
        return "zh"
    if contains_any(value, {"en", "english", "英文", "英语"}):
        return "en"
    return ""


HUMANIZER_SKILL_MOJIBAKE_TOKENS = {
    "\u951b",
    "\u9225",
    "\u9348",
    "\u93c2",
    "\u95c2",
    "\u6dc7",
    "\u5be4",
    "\u752f",
    "\u8fe9",
    "\u9286",
    "\u20ac",
    "\ufffd",
}


def _validate_loaded_humanizer_skill_path(raw_skill_path: str, skill_name: str, evidence_path: Path) -> list[str]:
    if "SKILL.md" not in raw_skill_path:
        return [f"humanizer evidence record must name the loaded SKILL.md path in {evidence_path}"]
    skill_path = resolve_record_path(raw_skill_path, evidence_path)
    path_issues = validate_existing_path(skill_path, require_nonempty_file=True)
    if path_issues:
        return [f"humanizer evidence skill file path is not readable in {evidence_path}: {skill_path}"]
    try:
        skill_text = skill_path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        return [f"humanizer evidence skill file path is not UTF-8 in {evidence_path}: {skill_path} ({exc})"]
    issues: list[str] = []
    frontmatter = "\n".join(skill_text.splitlines()[:20])
    if not re.search(rf"(?im)^name:\s*{re.escape(skill_name)}\s*$", frontmatter):
        issues.append(
            f"humanizer evidence skill file path does not match skill name {skill_name} in {evidence_path}: {skill_path}"
        )
    if contains_any(skill_text[:4000], HUMANIZER_SKILL_MOJIBAKE_TOKENS):
        issues.append(f"humanizer evidence skill file appears mojibake-corrupted in {evidence_path}: {skill_path}")
    return issues


def check_humanizer_evidence_record(evidence_path: Path) -> tuple[list[str], set[str], set[str]]:
    issues: list[str] = []
    seen_skills: set[str] = set()
    seen_languages: set[str] = set()

    issues.extend(validate_existing_path(evidence_path, require_nonempty_file=True))
    if issues:
        return issues, seen_skills, seen_languages
    if evidence_path.suffix.lower() not in TEXT_EVIDENCE_EXTENSIONS:
        return [f"humanizer evidence record must be a text file (.md or .txt): {evidence_path}"], seen_skills, seen_languages

    try:
        lines = read_lines(evidence_path)
    except UnicodeDecodeError as exc:
        return [f"humanizer evidence record is not valid UTF-8: {evidence_path} ({exc})"], seen_skills, seen_languages

    normalized_lines = {normalize(line) for line in lines if normalize(line)}
    for heading in HUMANIZER_EVIDENCE_SCHEMA["headings"]:
        if heading not in normalized_lines:
            issues.append(f"humanizer evidence record missing heading in {evidence_path}: {heading}")

    values: dict[str, str] = {}
    for prefix in HUMANIZER_EVIDENCE_SCHEMA["single_prefixes"]:
        matches = find_lines_with_prefix(lines, prefix)
        if len(matches) != 1:
            issues.append(f"humanizer evidence record must contain exactly one '{prefix}' line: {evidence_path}")
            continue
        values[prefix] = parse_line_value(matches[0])

    if issues:
        return issues, seen_skills, seen_languages

    for prefix in HUMANIZER_EVIDENCE_SCHEMA["single_prefixes"]:
        value = values[prefix]
        if prefix == "- blocker summary:":
            if not is_explicit_or_none(value):
                issues.append(f"humanizer evidence record blocker field is incomplete in {evidence_path}")
            continue
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"humanizer evidence record field must not be empty in {evidence_path}: {prefix}")

    skill_name = values["- skill name:"].lower()
    target_language = values["- target language:"]
    language_kind = _humanizer_language_kind(target_language)

    if skill_name not in {"humanizer-zh", "humanizer"}:
        issues.append(f"humanizer evidence record has unsupported skill name in {evidence_path}: {values['- skill name:']}")
    else:
        seen_skills.add(skill_name)

    if not language_kind:
        issues.append(
            f"humanizer evidence record target language must be Chinese/zh or English/en in {evidence_path}: {target_language}"
        )
    else:
        seen_languages.add(language_kind)

    if skill_name == "humanizer-zh" and language_kind != "zh":
        issues.append(f"humanizer-zh evidence must target Chinese text in {evidence_path}")
    if skill_name == "humanizer" and language_kind != "en":
        issues.append(f"humanizer evidence must target English text in {evidence_path}")
    raw_skill_path = values["- skill file path:"]
    issues.extend(_validate_loaded_humanizer_skill_path(raw_skill_path, skill_name, evidence_path))
    if normalize(values["- skill loaded before rewrite?:"]).lower() not in {"yes", "true", "pass", "passed"}:
        issues.append(f"humanizer evidence record must prove the skill was loaded before rewrite in {evidence_path}")
    if not is_explicit(values["- route application method:"]) or contains_any(
        values["- route application method:"],
        {
            "acceptance generator",
            "acceptance-record generator",
            "generate_thesis_acceptance_record",
            "validator-only",
            "report synthesizer",
            "smoke checker",
            "path-only",
        },
    ):
        issues.append(f"humanizer evidence record route application method is not a real content-edit route in {evidence_path}")
    before_text = values["- before text:"].strip()
    after_text = values["- after text:"].strip()
    rewrite_changed = normalize(values["- rewrite changed?:"]).lower()
    cleanup_verdict = values["- AI-pattern cleanup verdict:"].lower()
    if before_text == after_text and rewrite_changed in {"yes", "true", "pass", "passed"}:
        issues.append(f"humanizer evidence record claims a rewrite but before/after text is identical in {evidence_path}")
    if before_text == after_text and not contains_any(cleanup_verdict, {"no-change", "no change", "unchanged with reason"}):
        issues.append(f"humanizer evidence record identical before/after text lacks a concrete no-change reason in {evidence_path}")
    if language_kind == "zh" and not contains_any(
        cleanup_verdict,
        {"meta-evaluation", "meta evaluation", "meta-writing", "thesis-native", "ai", "ai-pattern", "ai pattern", "ai腔", "论文腔"},
    ):
        issues.append(
            "humanizer-zh evidence must explicitly cover meta-evaluation voice or thesis-native wording cleanup "
            f"in {evidence_path}"
        )

    return issues, seen_skills, seen_languages


def detect_format_repair_surfaces(values: dict[str, str], raw_values: dict[str, str]) -> dict[str, bool]:
    flags: dict[str, bool] = {}
    for name, spec in FORMAT_REPAIR_TASK_SURFACE_DETECTOR_SPECS.items():
        context = build_raw_context(raw_values, spec["raw_prefixes"])
        touched = contains_any(context, spec["tokens"])
        explicit_flag_prefix = spec.get("explicit_flag_prefix")
        if explicit_flag_prefix:
            touched = values[explicit_flag_prefix] not in EXPLICIT_VALUES or touched
        flags[name] = touched
    flags["header_footer"] = flags.get("header_footer", False) or flags.get("tail_block", False)
    return flags


def detect_review_evidence_surfaces(raw_values: dict[str, str]) -> tuple[dict[str, bool], str]:
    combined_text = build_raw_context(raw_values, REVIEW_EVIDENCE_SURFACE_TEXT_PREFIXES)
    flags = {
        name: contains_any(combined_text, spec["tokens"])
        for name, spec in REVIEW_EVIDENCE_SURFACE_DETECTOR_SPECS.items()
    }
    return flags, combined_text


def check_docx_font_audit_report(report_path: Path, expected_docx_path: Path | None = None) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues

    lines = read_lines(report_path)
    normalized_lines = {normalize(line) for line in lines if normalize(line)}
    for heading in DOCX_FONT_AUDIT_SCHEMA["headings"]:
        if heading not in normalized_lines:
            issues.append(f"docx font audit report missing heading '{heading}' in {report_path}")

    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    for prefix in DOCX_FONT_AUDIT_SCHEMA["single_prefixes"]:
        matches = find_lines_with_prefix(lines, prefix)
        if len(matches) != 1:
            issues.append(f"docx font audit report must contain exactly one '{prefix}' line in {report_path}")
            continue
        values[prefix] = parse_line_value(matches[0])
        raw_values[prefix] = raw_line_value(matches[0])

    if values.get("- result:") != "pass":
        issues.append(f"docx font audit report is not pass in {report_path}")
    if values.get("- bibliography font-slot checks:") != "pass":
        issues.append(f"docx font audit report bibliography font-slot checks are not pass in {report_path}")

    raw_doc_path = raw_values.get("- docx path:", "")
    resolved_doc_path: Path | None = None
    if raw_doc_path:
        resolved_doc_path = Path(raw_doc_path)
        if not resolved_doc_path.is_absolute():
            resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
        else:
            resolved_doc_path = resolved_doc_path.resolve()
        if resolved_doc_path.exists():
            expected_sha = sha256_file(resolved_doc_path)
            report_sha = values.get("- docx sha256:", "")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
                issues.append(f"docx font audit report must record a 64-hex docx sha256 in {report_path}")
            elif report_sha.lower() != expected_sha.lower():
                issues.append(
                    f"docx font audit report sha256 in {report_path} does not match its audited DOCX path"
                )
        else:
            issues.append(f"docx font audit report targets a missing DOCX path in {report_path}: {resolved_doc_path}")

    reference_path_value = raw_values.get("- bibliography reference docx path:", "")
    if reference_path_value and normalize(reference_path_value).lower() not in EXPLICIT_VALUES:
        reference_path = Path(reference_path_value)
        if not reference_path.is_absolute():
            reference_path = (report_path.parent / reference_path).resolve()
        else:
            reference_path = reference_path.resolve()
        if reference_path.exists():
            reference_sha = values.get("- bibliography reference docx sha256:", "")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", reference_sha):
                issues.append(f"docx font audit report must record a 64-hex bibliography reference docx sha256 in {report_path}")
            elif reference_sha.lower() != sha256_file(reference_path).lower():
                issues.append(f"docx font audit report bibliography reference sha256 does not match in {report_path}")
        else:
            issues.append(f"docx font audit report bibliography reference DOCX is missing in {report_path}: {reference_path}")

    def parse_nonnegative_int(prefix: str) -> int | None:
        raw = values.get(prefix, "")
        try:
            parsed = int(raw)
        except ValueError:
            issues.append(f"docx font audit report field must be an integer in {report_path}: {prefix} {raw!r}")
            return None
        if parsed < 0:
            issues.append(f"docx font audit report field must be non-negative in {report_path}: {prefix} {raw!r}")
            return None
        return parsed

    bibliography_entry_count = parse_nonnegative_int("- bibliography entry count:")
    bibliography_checked_run_count = parse_nonnegative_int("- bibliography checked run count:")
    bibliography_content_model_status = values.get("- bibliography content-format model checks:", "")
    bibliography_content_model_source = raw_values.get("- bibliography content-format model source:", "")
    if values.get("- bibliography font-slot checks:") == "pass":
        if bibliography_entry_count is None or bibliography_entry_count <= 0:
            issues.append(f"docx font audit report must record positive bibliography entry coverage in {report_path}")
        if bibliography_checked_run_count is None or bibliography_checked_run_count <= 0:
            issues.append(f"docx font audit report must record positive bibliography run coverage in {report_path}")
        if bibliography_entry_count is not None and bibliography_entry_count > 0:
            if bibliography_content_model_status != "pass":
                issues.append(
                    f"docx font audit report must pass bibliography content-format model checks in {report_path}"
                )
            if not bibliography_content_model_source or normalize(bibliography_content_model_source).lower() in EXPLICIT_VALUES:
                issues.append(
                    f"docx font audit report must bind bibliography content-format model source in {report_path}"
                )

    size_policy_source = values.get("- bibliography size policy source:", "")
    expected_size_name = values.get("- bibliography expected size name:", "")
    expected_size_half_points = values.get("- bibliography expected size half-points:", "")
    explicit_named_size = size_policy_source == "explicit-named-size-cli"
    if expected_size_name and expected_size_name not in {"template-derived", "none", "not-required"}:
        if explicit_named_size:
            wps_path_value = raw_values.get("- bibliography named-size WPS evidence path:", "")
            wps_verdict = values.get("- bibliography named-size WPS evidence verdict:", "")
            wps_path_lower = normalize(wps_path_value).lower()
            if not wps_path_value or wps_path_lower in EXPLICIT_VALUES or wps_path_lower == "not-required":
                issues.append(f"docx font audit report explicit named-size policy lacks WPS named-size evidence path in {report_path}")
            if wps_verdict != "pass":
                issues.append(f"docx font audit report explicit named-size policy lacks WPS named-size pass verdict in {report_path}")
            if wps_path_value and wps_path_lower not in EXPLICIT_VALUES and wps_path_lower != "not-required" and resolved_doc_path is not None:
                wps_path = Path(wps_path_value)
                if not wps_path.is_absolute():
                    wps_path = (report_path.parent / wps_path).resolve()
                else:
                    wps_path = wps_path.resolve()
                issues.extend(validate_existing_path(wps_path, require_nonempty_file=True))
                if wps_path.exists():
                    try:
                        payload = json.loads(wps_path.read_text(encoding="utf-8-sig"))
                    except Exception as exc:
                        issues.append(f"WPS named-size evidence is not valid JSON in {report_path}: {wps_path} ({exc})")
                    else:
                        if payload.get("schema") != "graduation-project-builder.wps-reference-entry-ui-font.v1":
                            issues.append(f"WPS named-size evidence schema mismatch in {report_path}: {wps_path}")
                        if str(payload.get("verdict") or "").lower() != "pass":
                            issues.append(f"WPS named-size evidence verdict is not pass in {report_path}: {wps_path}")
                        docx_sha = str(payload.get("docxSha256") or payload.get("docx_sha256") or "")
                        if docx_sha.lower() != sha256_file(resolved_doc_path).lower():
                            issues.append(f"WPS named-size evidence DOCX sha256 does not match final font audit target in {report_path}: {wps_path}")
                        checked_entry_count = payload.get("checkedEntryCount")
                        if checked_entry_count is None and isinstance(payload.get("entries"), list):
                            checked_entry_count = len(payload["entries"])
                        try:
                            checked_entry_count_int = int(str(checked_entry_count))
                        except (TypeError, ValueError):
                            checked_entry_count_int = -1
                        if bibliography_entry_count is not None and checked_entry_count_int != bibliography_entry_count:
                            issues.append(
                                f"WPS named-size evidence coverage does not match bibliography entry count in {report_path}: {wps_path}"
                            )
                        if str(payload.get("expectedWpsDisplaySizeName") or payload.get("expectedDisplaySizeName") or "") != expected_size_name:
                            issues.append(f"WPS named-size evidence expected display size name mismatch in {report_path}: {wps_path}")
                        if str(payload.get("expectedSizeHalfPoints") or payload.get("expected_size_half_points") or "") != expected_size_half_points:
                            issues.append(f"WPS named-size evidence expected half-points mismatch in {report_path}: {wps_path}")
                        entries = payload.get("entries")
                        if not isinstance(entries, list) or not entries:
                            issues.append(f"WPS named-size evidence must include entries[] all-entry proof in {report_path}: {wps_path}")
                        else:
                            for index, entry in enumerate(entries, start=1):
                                if not isinstance(entry, dict):
                                    issues.append(f"WPS named-size evidence entry {index} is not an object in {report_path}: {wps_path}")
                                    continue
                                display_name = str(
                                    entry.get("inferredWpsDisplaySizeName")
                                    or entry.get("wpsDisplaySizeName")
                                    or entry.get("displaySizeName")
                                    or ""
                                )
                                if str(entry.get("verdict") or "").lower() != "pass":
                                    issues.append(f"WPS named-size evidence entry {index} verdict is not pass in {report_path}: {wps_path}")
                                if display_name != expected_size_name:
                                    issues.append(f"WPS named-size evidence entry {index} display size is not the exact named size in {report_path}: {wps_path}")
        elif size_policy_source not in {"explicit-half-points-cli", "template-derived"}:
            issues.append(f"docx font audit report has unknown bibliography size policy source in {report_path}: {size_policy_source}")
        elif size_policy_source == "explicit-half-points-cli" and reference_path_value and normalize(reference_path_value).lower() not in EXPLICIT_VALUES:
            issues.append(
                f"docx font audit report must not use explicit half-point bibliography size override when a reference DOCX is bound; use template-derived or explicit named-size evidence in {report_path}"
            )
    if expected_docx_path is not None and "- docx path:" in raw_values:
        if resolved_doc_path is None:
            raw_doc_path = raw_values["- docx path:"]
            resolved_doc_path = Path(raw_doc_path)
            if not resolved_doc_path.is_absolute():
                resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
            else:
                resolved_doc_path = resolved_doc_path.resolve()
        if resolved_doc_path != expected_docx_path.resolve():
            issues.append(
                f"docx font audit report in {report_path} targets {resolved_doc_path} instead of the exact rendered deliverable {expected_docx_path.resolve()}"
            )
        else:
            expected_sha = sha256_file(expected_docx_path)
            report_sha = values.get("- docx sha256:", "")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
                issues.append(f"docx font audit report must record a 64-hex docx sha256 in {report_path}")
            elif report_sha.lower() != expected_sha.lower():
                issues.append(
                    f"docx font audit report sha256 in {report_path} does not match the exact rendered deliverable {expected_docx_path.resolve()}"
                )

    return issues


def check_docx_body_style_audit_report(report_path: Path, expected_final_docx_path: Path | None = None) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues

    lines = read_lines(report_path)
    normalized_lines = {normalize(line) for line in lines if normalize(line)}
    for heading in BODY_STYLE_AUDIT_SCHEMA["headings"]:
        if heading not in normalized_lines:
            issues.append(f"docx body style audit report missing heading '{heading}' in {report_path}")

    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    for prefix in BODY_STYLE_AUDIT_SCHEMA["single_prefixes"]:
        matches = find_lines_with_prefix(lines, prefix)
        if len(matches) != 1:
            issues.append(f"docx body style audit report must contain exactly one '{prefix}' line in {report_path}")
            continue
        values[prefix] = parse_line_value(matches[0])
        raw_values[prefix] = raw_line_value(matches[0])

    if values.get("- result:") != "pass":
        issues.append(f"docx body style audit report is not pass in {report_path}")
    heading_contamination_summary = values.get("- body heading contamination summary:", "")
    if heading_contamination_summary and not heading_contamination_summary.startswith("passed"):
        issues.append(f"docx body style audit report indicates body prose heading-style contamination in {report_path}")
    if expected_final_docx_path is not None and "- final docx path:" in raw_values:
        raw_doc_path = raw_values["- final docx path:"]
        resolved_doc_path = Path(raw_doc_path)
        if not resolved_doc_path.is_absolute():
            resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
        else:
            resolved_doc_path = resolved_doc_path.resolve()
        if resolved_doc_path != expected_final_docx_path.resolve():
            issues.append(
                f"docx body style audit report in {report_path} targets {resolved_doc_path} instead of the exact rendered deliverable {expected_final_docx_path.resolve()}"
            )
        expected_sha = sha256_file(expected_final_docx_path)
        report_sha = values.get("- final docx sha256:", "")
        if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
            issues.append(f"docx body style audit report must record a 64-hex final docx sha256 in {report_path}")
        elif report_sha.lower() != expected_sha.lower():
            issues.append(
                f"docx body style audit report sha256 in {report_path} does not match the exact rendered deliverable {expected_final_docx_path.resolve()}"
            )

    return issues


def check_docx_whole_format_gate_report(
    report_path: Path,
    expected_final_docx_path: Path | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues

    try:
        payload = json.loads(report_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        return [f"whole-format DOCX gate report is not valid JSON in {report_path}: {exc}"]
    if not isinstance(payload, dict):
        return [f"whole-format DOCX gate report must be a JSON object in {report_path}"]

    if payload.get("schema") != "graduation-project-builder.docx-whole-format-gate.v1":
        issues.append(f"whole-format DOCX gate report schema mismatch in {report_path}")
    if payload.get("passed") is not True:
        report_issues = payload.get("issues")
        issues.append(f"whole-format DOCX gate report is not pass in {report_path}: {report_issues}")

    raw_doc_path = str(payload.get("docx_path") or "")
    resolved_doc_path: Path | None = None
    if not raw_doc_path:
        issues.append(f"whole-format DOCX gate report lacks docx_path in {report_path}")
    else:
        resolved_doc_path = Path(raw_doc_path)
        if not resolved_doc_path.is_absolute():
            resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
        else:
            resolved_doc_path = resolved_doc_path.resolve()
        if not resolved_doc_path.exists():
            issues.append(f"whole-format DOCX gate report targets a missing DOCX in {report_path}: {resolved_doc_path}")

    report_sha = str(payload.get("docx_sha256") or "")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
        issues.append(f"whole-format DOCX gate report must record a 64-hex docx_sha256 in {report_path}")
    elif resolved_doc_path is not None and resolved_doc_path.exists():
        actual_sha = sha256_file(resolved_doc_path)
        if report_sha.lower() != actual_sha.lower():
            issues.append(f"whole-format DOCX gate report sha256 does not match audited DOCX in {report_path}")

    if expected_final_docx_path is not None and resolved_doc_path is not None:
        if resolved_doc_path != expected_final_docx_path.resolve():
            issues.append(
                f"whole-format DOCX gate report in {report_path} targets {resolved_doc_path} "
                f"instead of the exact final DOCX {expected_final_docx_path.resolve()}"
            )
        elif re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
            expected_sha = sha256_file(expected_final_docx_path)
            if report_sha.lower() != expected_sha.lower():
                issues.append(
                    f"whole-format DOCX gate report sha256 does not match the exact final DOCX "
                    f"{expected_final_docx_path.resolve()}"
                )

    counts = payload.get("counts")
    if isinstance(counts, dict):
        required_minimums = {
            "section_count": 3,
            "live_toc_field_count": 1,
            "footer_page_field_count": 1,
        }
        for key, minimum in required_minimums.items():
            try:
                value = int(counts.get(key))
            except (TypeError, ValueError):
                issues.append(f"whole-format DOCX gate report count '{key}' is missing or non-integer in {report_path}")
                continue
            if value < minimum:
                issues.append(
                    f"whole-format DOCX gate report count '{key}' is below release minimum {minimum} in {report_path}: {value}"
                )
        try:
            builder_style_count = int(counts.get("builder_style_visible_paragraph_count"))
        except (TypeError, ValueError):
            issues.append(
                f"whole-format DOCX gate report count 'builder_style_visible_paragraph_count' is missing or non-integer in {report_path}"
            )
        else:
            if builder_style_count != 0:
                issues.append(
                    f"whole-format DOCX gate report still shows builder-owned visible styles in {report_path}: {builder_style_count}"
                )
    else:
        issues.append(f"whole-format DOCX gate report lacks counts object in {report_path}")

    return issues


def check_docx_font_color_audit_report(
    report_path: Path,
    expected_final_docx_path: Path | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues

    try:
        payload = json.loads(report_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        return [f"font-color DOCX audit report is not valid JSON in {report_path}: {exc}"]
    if not isinstance(payload, dict):
        return [f"font-color DOCX audit report must be a JSON object in {report_path}"]

    if payload.get("schema") != "graduation-project-builder.docx-font-color-audit.v1":
        issues.append(f"font-color DOCX audit report schema mismatch in {report_path}")
    if payload.get("passed") is not True:
        issues.append(
            f"font-color DOCX audit report is not pass in {report_path}: "
            f"nonblack_color_count={payload.get('nonblack_color_count')}"
        )

    raw_doc_path = str(payload.get("docx_path") or "")
    resolved_doc_path: Path | None = None
    if not raw_doc_path:
        issues.append(f"font-color DOCX audit report lacks docx_path in {report_path}")
    else:
        resolved_doc_path = Path(raw_doc_path)
        if not resolved_doc_path.is_absolute():
            resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
        else:
            resolved_doc_path = resolved_doc_path.resolve()
        if not resolved_doc_path.exists():
            issues.append(f"font-color DOCX audit report targets a missing DOCX in {report_path}: {resolved_doc_path}")

    report_sha = str(payload.get("docx_sha256") or "")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
        issues.append(f"font-color DOCX audit report must record a 64-hex docx_sha256 in {report_path}")
    elif resolved_doc_path is not None and resolved_doc_path.exists():
        actual_sha = sha256_file(resolved_doc_path)
        if report_sha.lower() != actual_sha.lower():
            issues.append(f"font-color DOCX audit report sha256 does not match audited DOCX in {report_path}")

    try:
        nonblack_count = int(payload.get("nonblack_color_count"))
    except (TypeError, ValueError):
        issues.append(f"font-color DOCX audit report nonblack_color_count is missing or non-integer in {report_path}")
    else:
        if nonblack_count != 0:
            issues.append(f"font-color DOCX audit report found visible non-black/theme-colored text in {report_path}")

    if expected_final_docx_path is not None and resolved_doc_path is not None:
        if resolved_doc_path != expected_final_docx_path.resolve():
            issues.append(
                f"font-color DOCX audit report in {report_path} targets {resolved_doc_path} "
                f"instead of the exact final DOCX {expected_final_docx_path.resolve()}"
            )
        elif re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
            expected_sha = sha256_file(expected_final_docx_path)
            if report_sha.lower() != expected_sha.lower():
                issues.append(
                    f"font-color DOCX audit report sha256 does not match the exact final DOCX "
                    f"{expected_final_docx_path.resolve()}"
                )

    return issues


def check_thesis_citation_audit_report(report_path: Path, expected_docx_path: Path | None = None) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues

    lines = read_lines(report_path)
    normalized_lines = {normalize(line) for line in lines if normalize(line)}
    for heading in CITATION_AUDIT_SCHEMA["headings"]:
        if heading not in normalized_lines:
            issues.append(f"citation audit report missing heading '{heading}' in {report_path}")

    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    for prefix in CITATION_AUDIT_SCHEMA["single_prefixes"]:
        matches = find_lines_with_prefix(lines, prefix)
        if len(matches) != 1:
            issues.append(f"citation audit report must contain exactly one '{prefix}' line in {report_path}")
            continue
        values[prefix] = parse_line_value(matches[0])
        raw_values[prefix] = raw_line_value(matches[0])

    if issues:
        return issues

    if values.get("- result:") != "pass":
        issues.append(f"citation audit report is not pass in {report_path}")
    if normalize(values.get("- error codes:", "")).lower() not in {"none", ""}:
        issues.append(f"citation audit report still carries error codes in {report_path}: {values['- error codes:']}")

    def parse_int(prefix: str) -> int | None:
        raw_value = values.get(prefix, "")
        try:
            return int(raw_value)
        except ValueError:
            issues.append(f"citation audit report field must be an integer in {report_path}: {prefix} {raw_value!r}")
            return None

    body_citation_paragraph_count = parse_int("- body citation paragraph count:")
    unique_citation_count = parse_int("- unique citation count:")
    bibliography_item_count = parse_int("- bibliography item count:")

    def expected_list_value(items: list[int], *, none_value: str = "[]") -> str:
        return str(items) if items else none_value

    def report_section(source_lines: list[str], start_heading: str, end_heading: str | None = None) -> list[str]:
        normalized_source = [normalize(line) for line in source_lines]
        start_norm = normalize(start_heading)
        end_norm = normalize(end_heading) if end_heading else None
        try:
            start_index = normalized_source.index(start_norm) + 1
        except ValueError:
            return []
        end_index = len(source_lines)
        if end_norm is not None:
            for idx in range(start_index, len(normalized_source)):
                if normalized_source[idx] == end_norm:
                    end_index = idx
                    break
        return [item for item in normalized_source[start_index:end_index] if item]

    if bibliography_item_count and bibliography_item_count > 0:
        if body_citation_paragraph_count == 0:
            issues.append(f"citation audit report claims a non-empty bibliography but zero body citation paragraphs in {report_path}")
        if unique_citation_count == 0:
            issues.append(f"citation audit report claims a non-empty bibliography but zero unique citations in {report_path}")

    if expected_docx_path is not None:
        raw_doc_path = raw_values.get("- document path:", "")
        resolved_doc_path = Path(raw_doc_path)
        if not resolved_doc_path.is_absolute():
            resolved_doc_path = (report_path.parent / resolved_doc_path).resolve()
        else:
            resolved_doc_path = resolved_doc_path.resolve()
        if resolved_doc_path != expected_docx_path.resolve():
            issues.append(
                f"citation audit report in {report_path} targets {resolved_doc_path} instead of the exact rendered deliverable {expected_docx_path.resolve()}"
            )
        expected_sha = sha256_file(expected_docx_path)
        report_sha = values.get("- document sha256:", "")
        if not re.fullmatch(r"[0-9a-fA-F]{64}", report_sha):
            issues.append(f"citation audit report must record a 64-hex document sha256 in {report_path}")
        elif report_sha.lower() != expected_sha.lower():
            issues.append(
                f"citation audit report sha256 in {report_path} does not match the exact rendered deliverable {expected_docx_path.resolve()}"
            )
        try:
            actual_audit = audit_body_citations(expected_docx_path)
        except Exception as exc:
            issues.append(f"citation audit report cannot be verified against exact DOCX {expected_docx_path.resolve()}: {exc}")
        else:
            if not actual_audit.passed:
                issues.append(
                    "citation audit report claims pass but recomputed body citation audit fails for "
                    f"{expected_docx_path.resolve()}: {', '.join(actual_audit.error_codes)}"
                )
            if body_citation_paragraph_count is not None and body_citation_paragraph_count != actual_audit.body_citation_paragraph_count:
                issues.append(
                    f"citation audit report body citation paragraph count is stale in {report_path}: "
                    f"report={body_citation_paragraph_count} actual={actual_audit.body_citation_paragraph_count}"
                )
            if unique_citation_count is not None and unique_citation_count != actual_audit.unique_citation_count:
                issues.append(
                    f"citation audit report unique citation count is stale in {report_path}: "
                    f"report={unique_citation_count} actual={actual_audit.unique_citation_count}"
                )
            if bibliography_item_count is not None and bibliography_item_count != actual_audit.bibliography_item_count:
                issues.append(
                    f"citation audit report bibliography item count is stale in {report_path}: "
                    f"report={bibliography_item_count} actual={actual_audit.bibliography_item_count}"
                )
            expected_summary_values = {
                "- first appearance chain:": expected_list_value(actual_audit.first_appearance_chain),
                "- expected chain:": expected_list_value(actual_audit.expected_chain),
                "- missing bibliography numbers:": expected_list_value(
                    actual_audit.missing_bibliography_numbers,
                    none_value="none",
                ),
                "- extra body citation numbers:": expected_list_value(
                    actual_audit.extra_body_citation_numbers,
                    none_value="none",
                ),
            }
            for prefix, expected_value in expected_summary_values.items():
                if normalize(values.get(prefix, "")) != normalize(expected_value):
                    issues.append(
                        f"citation audit report field is stale in {report_path}: "
                        f"{prefix} report={values.get(prefix, '')!r} actual={expected_value!r}"
                    )
            expected_report_lines = build_citation_audit_report(
                actual_audit,
                expected_docx_path.resolve(),
            ).splitlines()
            if report_section(lines, "## Body Citation Records", "## Findings") != report_section(
                expected_report_lines,
                "## Body Citation Records",
                "## Findings",
            ):
                issues.append(f"citation audit report body citation records are stale in {report_path}")
            if report_section(lines, "## Findings") != report_section(expected_report_lines, "## Findings"):
                issues.append(f"citation audit report findings section is stale in {report_path}")

    return issues


def format_repair_task_touches_surface(task_path: Path, surface_name: str) -> bool:
    try:
        lines = read_lines(task_path)
    except Exception:
        return False
    values = {
        prefix: parse_line_value(found[0]) if (found := find_lines_with_prefix(lines, prefix)) else ""
        for prefix in FORMAT_REPAIR_TASK_SCHEMA["single_prefixes"]
    }
    raw_values = {
        prefix: raw_line_value(found[0]) if (found := find_lines_with_prefix(lines, prefix)) else ""
        for prefix in FORMAT_REPAIR_TASK_SCHEMA["single_prefixes"]
    }
    return detect_format_repair_surfaces(values, raw_values).get(surface_name, False)
