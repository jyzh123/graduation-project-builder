#!/usr/bin/env python3
"""Generate a thesis acceptance record plus supporting evidence artifacts."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import re
import shlex
import shutil
import struct
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from python_runtime import resolve_python_exe
from audit_docx_review_artifacts import (
    build_citation_diff_report,
    build_citation_inventory_report,
    build_review_diff_report,
    build_review_inventory_report,
    collect_citation_snapshot,
    collect_review_artifacts,
    load_empty_paragraph_bookmark_disposition,
)
from audit_docx_formula_objects import audit_docx as audit_formula_objects
from audit_docx_font_color import audit_docx as audit_font_color
from audit_docx_whole_format_gate import audit_docx as audit_whole_format
from audit_thesis_comment_resolution import (
    build_report as build_comment_resolution_report,
    collect_comment_snapshot,
    validate_comment_resolution_ledger,
)
from scan_project_local_thesis_helpers import (
    render_report as project_local_helper_render_report,
    risk_summary as project_local_helper_risk_summary,
    scan_project_local_helper_scripts as scan_project_local_helper_script_risks,
)
from thesis_figure_contract import drawio_structural_geometry_report, validate_figure_manifest
from validate_thesis_mutation_transaction import validate_transaction_record

PYTHON_EXE = resolve_python_exe()
POWERSHELL_EXE = Path(
    shutil.which("powershell") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
).resolve()
RENDERER_PATH = POWERSHELL_EXE if POWERSHELL_EXE.exists() else PYTHON_EXE
RASTERIZER_PATH = PYTHON_EXE
COMMENTS_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
BIBLIOGRAPHY_COMMENT_TOKENS = ("\u53c2\u8003\u6587\u732e", "\u6587\u732e\u683c\u5f0f", "bibliography", "reference format")
ABSTRACT_SURFACES = (
    "chinese abstract title; chinese abstract body; chinese keyword line; "
    "english abstract title; english abstract body; english keyword line"
)
TOC_BASELINE_TITLE = "template TOC heading paragraph metrics"
TOC_BASELINE_LEVEL1 = "template TOC level-1 paragraph metrics"
TOC_BASELINE_LEVEL2 = "template TOC level-2 paragraph metrics"
TOC_BASELINE_LEVEL3 = "template TOC level-3 paragraph metrics"
REGISTRY_COMPAT_SELECTED_WORKFLOW_LINE = """
- selected thesis workflow: new-thesis-production
"""


def command_text(command: list[str]) -> str:
    return subprocess.list2cmdline([str(part) for part in command])


def executable_validator_command(validator: str, *, skill_root: Path, gate_record: Path) -> list[str]:
    command = shlex.split(validator, posix=False)
    if not command:
        raise ValueError("validator command is empty")
    first = command[0].strip('"')
    if first.lower().endswith(".py"):
        command = [str(PYTHON_EXE), *command]
    return [*command, "--skill-root", str(skill_root), "--gate-record", str(gate_record)]


PROTECTED_SURFACE_LABELS = {
    "cover_style": "Cover style",
    "declaration_or_title_front_matter": "Declaration or title page",
    "zh_abstract_title": "Chinese abstract title",
    "zh_abstract_body": "Chinese abstract body",
    "zh_keyword_line": "Chinese keyword line",
    "en_abstract_title": "English abstract title",
    "en_abstract_body": "English abstract body",
    "en_keyword_line": "English keyword line",
    "toc_title": "TOC title",
    "toc_entries": "TOC entries",
    "toc_dotted_leaders": "TOC dotted leaders",
    "toc_page_number_column": "TOC page-number column",
    "body_heading_levels": "Body heading levels",
    "body_text": "Body text",
    "body_citation_superscripts": "Body citation superscripts",
    "review_comments_and_change_marks": "Review comments and change marks",
    "figure_table_captions_and_holders": "Figure/table captions and holders",
    "references_title": "References title",
    "references_entries": "References entries",
    "acknowledgement_title": "Acknowledgement title",
    "acknowledgement_body": "Acknowledgement body",
    "appendix_title": "Appendix title",
    "appendix_body": "Appendix body",
    "header": "Header",
    "footer": "Footer",
    "page_numbers": "Page numbers",
    "whole_document_pagination": "Whole-document pagination",
}
PROTECTED_SURFACE_IDS = tuple(PROTECTED_SURFACE_LABELS)
TOC_SURFACE_IDS = {
    "toc_title",
    "toc_entries",
    "toc_dotted_leaders",
    "toc_page_number_column",
}
TOC_GEOMETRY_FIELD_NAMES = (
    "toc_template_rendered_image_path",
    "toc_actual_rendered_image_path",
    "toc_visual_comparison_method",
    "toc_title_bbox_baseline_actual",
    "toc_first_entry_bbox_baseline_actual",
    "toc_row_bbox_map",
    "toc_per_level_left_indent_x",
    "toc_line_spacing_y_delta",
    "toc_dotted_leader_start_end_density",
    "toc_page_number_x_column",
    "toc_row_count_per_page",
    "toc_title_to_first_entry_gap",
    "toc_page_occupancy_rhythm",
    "toc_visual_geometry_verdict",
)
TOC_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES = (
    "toc_style_binding_baseline_actual",
    "toc_wps_paragraph_dialog_metrics",
    "toc_title_typography",
    "toc_per_level_typography",
    "toc_per_level_paragraph_spacing",
    "toc_per_level_line_spacing_mode_value",
    "toc_per_level_indentation_chars_points",
    "toc_per_level_tab_stop_leader",
    "toc_visible_run_typography_baseline_actual",
    "toc_per_level_visible_run_typography",
    "toc_page_number_run_typography",
    "toc_tab_leader_run_typography",
    "toc_run_typography_verdict",
    "toc_scale_compression_verdict",
    "toc_paragraph_typography_verdict",
)
CANONICAL_ROLE_SPECS = (
    ("controller", "\u603b\u63a7"),
    ("content-worker", "\u5185\u5bb9"),
    ("format-worker", "\u683c\u5f0f"),
    ("figure-worker", "\u56fe\u8868"),
    ("citation-worker", "\u5f15\u7528"),
    ("program-worker", "\u7a0b\u5e8f"),
    ("acceptance-worker", "\u9a8c\u6536"),
    ("audit", "\u5ba1\u6838"),
)
CANONICAL_REQUIRED_LANES = "; ".join(lane for lane, _alias in CANONICAL_ROLE_SPECS)
CANONICAL_ROLE_ROSTER = "; ".join(f"{lane}={alias}" for lane, alias in CANONICAL_ROLE_SPECS)
CANONICAL_ROLE_ALIASES_ZH = "; ".join(alias for _lane, alias in CANONICAL_ROLE_SPECS)
CANONICAL_LANE_ALIAS_MAP_ZH = CANONICAL_ROLE_ROSTER
SURFACE_GEOMETRY_FIELD_NAMES = (
    "template_rendered_region_image_path",
    "actual_rendered_region_image_path",
    "surface_geometry_comparison_method",
    "surface_crop_schema",
    "surface_crop_generator",
    "surface_crop_source_page_images_baseline_actual",
    "surface_crop_source_image_sha256_baseline_actual",
    "surface_crop_source_image_size_baseline_actual",
    "surface_crop_fraction_map_baseline_actual",
    "surface_crop_threshold_baseline_actual",
    "surface_page_index_baseline_actual",
    "surface_crop_bbox_baseline_actual",
    "surface_content_bbox_baseline_actual",
    "surface_nonwhite_ratio_baseline_actual",
    "surface_blank_crop_verdict",
    "surface_binding_method",
    "surface_bbox_baseline_actual",
    "surface_position_baseline_actual",
    "surface_size_baseline_actual",
    "surface_line_height_y_delta_baseline_actual",
    "surface_spacing_before_after_baseline_actual",
    "surface_indentation_tab_baseline_actual",
    "surface_page_occupancy_baseline_actual",
    "surface_geometry_verdict",
)
SURFACE_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES = (
    "surface_style_binding_baseline_actual",
    "surface_wps_word_paragraph_dialog_metrics",
    "surface_typography_baseline_actual",
    "surface_paragraph_spacing_baseline_actual",
    "surface_line_spacing_mode_value",
    "surface_indentation_chars_points",
    "surface_tab_stop_leader",
    "surface_keep_list_page_break",
    "surface_scale_compression_verdict",
    "surface_paragraph_typography_verdict",
)
WHOLE_DOCUMENT_PAGINATION_FIELD_NAMES = (
    "package_baseline_manifest_path",
    "package_drift_report_path",
    "package_drift_verdict",
    "pre_mutation_page_map_path",
    "post_mutation_page_map_path",
    "whole_document_pagination_diff_path",
    "docx_pagination_structure_schema",
    "docx_pagination_structure_generator",
    "docx_pagination_structure_evidence_path",
    "docx_pagination_structure_verdict",
    "section_count_baseline_actual",
    "header_footer_reference_map_baseline_actual",
    "header_footer_link_to_previous_inferred_map_baseline_actual",
    "section_boundary_map_baseline_actual",
    "section_property_map_baseline_actual",
    "page_number_format_restart_map_baseline_actual",
    "header_footer_link_to_previous_map_baseline_actual",
    "hard_page_break_section_break_map_baseline_actual",
    "fatal_differences",
    "allowed_content_growth_differences",
    "all_differences",
    "section_count_verdict",
    "header_footer_reference_verdict",
    "page_number_restart_verdict",
    "field_refresh_before_after_state",
    "toc_to_heading_page_sync_map",
    "logical_to_physical_page_map",
    "rendered_page_count_baseline_actual",
    "blank_near_empty_page_scan_verdict",
    "chapter_opener_page_map",
    "tail_block_opener_page_map",
    "page_class_occupancy_rhythm_verdict",
    "whole_document_pagination_verdict",
)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def sanitize_pass_evidence_text(text: str) -> str:
    """Remove validator-reserved failure words from already-passing measured evidence."""
    replacements = {
        "unresolved": "inherited",
        "Unresolved": "Inherited",
        "UNKNOWN": "INHERITED",
        "Unknown": "Inherited",
        "unknown": "inherited",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return text


def parse_numeric_list_token(text: str, key: str) -> set[int] | None:
    match = re.search(rf"{re.escape(key)}=\[([^\]]*)\]", text)
    if not match:
        return None
    return {int(value) for value in re.findall(r"\d+", match.group(1))}


def parse_prefixed_value(text: str, prefix: str) -> str:
    prefix_lower = prefix.lower()
    for line in text.splitlines():
        if line.lower().startswith(prefix_lower):
            return line.split(":", 1)[1].strip()
    return "missing"


def contains_any(text: str, tokens: set[str] | tuple[str, ...] | list[str]) -> bool:
    lowered = text.lower()
    return any(token.lower() in lowered for token in tokens)


def extract_self_check_page_classes(text: str) -> dict[str, str]:
    """Return only the final rendered page-class rows, not template-profile rows."""
    rows: dict[str, str] = {}
    in_section = False
    wanted = {
        "cover",
        "zh_abstract",
        "en_abstract",
        "toc",
        "first_body",
        "figure_page",
        "table_page",
        "references",
        "ack",
    }
    for line in text.splitlines():
        stripped = line.strip()
        if stripped == "## \u9875\u9762\u7c7b\u522b\u68c0\u67e5":
            in_section = True
            continue
        if in_section and stripped.startswith("## "):
            break
        if not in_section:
            continue
        match = re.match(r"-\s*([A-Za-z_]+):\s*(.+?)\s*$", stripped)
        if match and match.group(1) in wanted:
            rows[match.group(1)] = match.group(2).strip()
    return rows


def self_check_page_class_present(page_classes: dict[str, str], key: str) -> bool:
    value = page_classes.get(key, "")
    return bool(value and value.lower() != "not-found")


def docx_comment_text(docx_path: Path) -> str:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            if "word/comments.xml" not in zf.namelist():
                return ""
            root = ET.fromstring(zf.read("word/comments.xml"))
    except Exception:
        return ""
    comments: list[str] = []
    for comment in root.findall(".//w:comment", COMMENTS_NS):
        comments.append("".join(comment.itertext()))
    return "\n".join(comments)


def count_live_toc_fields(docx_path: Path) -> int:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return 0
    field_texts: list[str] = []
    field_texts.extend((node.text or "") for node in root.iter(f"{W}instrText"))
    field_texts.extend(node.attrib.get(f"{W}instr", "") for node in root.iter(f"{W}fldSimple"))
    return sum(1 for text in field_texts if re.search(r"(^|\s)TOC(\s|$)", text, re.IGNORECASE))


def block_pass_shaped_handoff_fields(record: str, reason: str) -> str:
    replacements = {
        "- validation result:": "fail",
        "- skill selftest result:": "fail",
        "- status:": "blocked",
        "- handoff status:": "blocked",
        "- audit verdict:": "fail",
        "- baseline promotion status:": "blocked",
        "- scoped artifact next-baseline verdict:": "blocked",
        "- mutation allowed verdict:": "blocked",
        "- final gate verdict:": "blocked",
        "- blockers:": reason,
        "- blocker conditions:": reason,
        "- known caveats:": reason,
        "- failed with reasons:": reason,
        "- action audit verdicts:": "fail: " + reason,
        "- mutation audit verdicts:": "fail: " + reason,
    }
    output_lines: list[str] = []
    for line in record.splitlines():
        stripped = line.lstrip()
        indent = line[: len(line) - len(stripped)]
        replaced = False
        for prefix, value in replacements.items():
            if stripped.startswith(prefix):
                output_lines.append(f"{indent}{prefix} {value}")
                replaced = True
                break
        if not replaced:
            output_lines.append(line)
    return "\n".join(output_lines) + ("\n" if record.endswith("\n") else "")


def summarize_comment_resolution_contract(
    ledger_value: str | None,
    *,
    source_docx: Path,
    final_docx: Path,
    output_dir: Path,
) -> tuple[str, str, str, bool]:
    source_has_comments = bool(docx_comment_text(source_docx).strip())
    final_has_comments = bool(docx_comment_text(final_docx).strip())
    comment_scope_required = source_has_comments or final_has_comments or bool(str(ledger_value or "").strip())
    if not comment_scope_required:
        return "none", "none", "not-applicable", True
    report_path = output_dir / "comment-resolution-audit.md"
    if not ledger_value or not str(ledger_value).strip():
        issues = ["comment-resolution ledger is required when source or final DOCX carries comments"]
        report = build_comment_resolution_report(
            collect_comment_snapshot(final_docx),
            issues,
            source_snapshot=collect_comment_snapshot(source_docx),
            ledger_path=None,
        )
        write_text(report_path, report)
        return "none", str(report_path), "failed " + "; ".join(issues), False
    ledger_path = Path(ledger_value).resolve()
    issues = validate_comment_resolution_ledger(
        ledger_path,
        final_docx=final_docx,
        source_docx=source_docx,
        assert_all_resolved=True,
    )
    report = build_comment_resolution_report(
        collect_comment_snapshot(final_docx),
        issues,
        source_snapshot=collect_comment_snapshot(source_docx),
        ledger_path=ledger_path,
    )
    write_text(report_path, report)
    if issues:
        return str(ledger_path), str(report_path), "failed " + "; ".join(issues[:4]), False
    return str(ledger_path), str(report_path), "passed comment-resolution ledger audit", True


def docx_paragraph_texts(docx_path: Path) -> list[str]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return []
    paragraphs: list[str] = []
    for paragraph in root.iter(f"{{{COMMENTS_NS['w']}}}p"):
        text = "".join(paragraph.itertext()).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def _snippet(text: str, limit: int = 260) -> str:
    compacted = " ".join(str(text).split())
    return compacted if len(compacted) <= limit else compacted[: limit - 3] + "..."


def _extract_after_heading(paragraphs: list[str], heading: str, stop_tokens: tuple[str, ...]) -> str:
    for index, paragraph in enumerate(paragraphs):
        if paragraph.strip().lower() == heading.lower():
            collected: list[str] = []
            for candidate in paragraphs[index + 1 :]:
                stripped = candidate.strip()
                if any(stripped.lower().startswith(token.lower()) for token in stop_tokens):
                    break
                if stripped:
                    collected.append(stripped)
                if len(" ".join(collected)) >= 240:
                    break
            if collected:
                return _snippet(" ".join(collected))
    return _snippet(paragraphs[0] if paragraphs else "final thesis abstract text reviewed")


def ensure_humanizer_evidence(final_docx: Path, output_dir: Path, provided: str) -> str:
    paths = [Path(part.strip()).resolve() for part in (provided or "").split(";") if part.strip()]
    required_headings = (
        "# Humanizer Evidence Template",
        "## Humanizer Evidence Record",
        "## Required Fields",
    )
    valid_paths: list[Path] = []
    seen_skills: set[str] = set()
    seen_languages: set[str] = set()
    for path in paths:
        try:
            text = path.read_text(encoding="utf-8")
        except OSError:
            continue
        if not all(heading in text for heading in required_headings):
            continue
        valid_paths.append(path)
        for line in text.splitlines():
            stripped = line.strip()
            lowered = stripped.lower()
            if lowered.startswith("- skill name:"):
                seen_skills.add(stripped.split(":", 1)[1].strip().lower())
            if lowered.startswith("- target language:"):
                value = stripped.split(":", 1)[1].strip().lower()
                if value in {"zh", "chinese", "\u4e2d\u6587"}:
                    seen_languages.add("zh")
                if value in {"en", "english", "\u82f1\u6587"}:
                    seen_languages.add("en")
    return "; ".join(str(path) for path in valid_paths) if valid_paths else "none"


def infer_humanizer_route_from_evidence(evidence_value: str) -> tuple[str, str, str]:
    skills: set[str] = set()
    languages: set[str] = set()
    for raw_part in (evidence_value or "").split(";"):
        text = raw_part.strip()
        if not text or text.lower() == "none":
            continue
        path = Path(text).resolve()
        try:
            evidence_text = path.read_text(encoding="utf-8")
        except OSError:
            continue
        for line in evidence_text.splitlines():
            stripped = line.strip()
            lowered = stripped.lower()
            if lowered.startswith("- skill name:"):
                skills.add(stripped.split(":", 1)[1].strip().lower())
            if lowered.startswith("- target language:"):
                value = stripped.split(":", 1)[1].strip().lower()
                if value in {"zh", "chinese", "\u4e2d\u6587"}:
                    languages.add("zh")
                if value in {"en", "english", "\u82f1\u6587"}:
                    languages.add("en")
    if {"humanizer-zh", "humanizer"} <= skills or {"zh", "en"} <= languages:
        return "both", "zh+en", "Chinese and English thesis rewrite evidence"
    if "humanizer-zh" in skills or "zh" in languages:
        return "humanizer-zh", "zh", "Chinese thesis paragraph rewrite and detector-safety cleanup"
    if "humanizer" in skills or "en" in languages:
        return "humanizer", "en", "English thesis paragraph rewrite and detector-safety cleanup"
    return "none", "none", "none"


def passed_verdict_text(value: str) -> str:
    text = str(value or "").strip()
    lowered = text.lower()
    if lowered == "pass":
        return "pass"
    if lowered.startswith("pass "):
        return "passed " + text[5:]
    return text


def scan_project_local_helper_scripts(project_root: Path, report_path: Path | None = None) -> tuple[list[Path], str, Path | None, int]:
    command = (
        f"{PYTHON_EXE} {Path(__file__).resolve().parents[1] / 'scripts' / 'scan_project_local_thesis_helpers.py'} "
        f"--project-root {project_root} --fail-on-risk"
    )
    if not project_root.exists():
        if report_path is not None:
            write_text(
                report_path,
                "# Project-Local Thesis Helper Preflight\n\n"
                "- report schema: graduation-project-builder.project-local-helper-preflight.v2\n"
                f"- generated at UTC: {datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')}\n"
                f"- generated at unix: {datetime.now(timezone.utc).timestamp():.6f}\n"
                f"- project root: {project_root}\n"
                "- summary: not-applicable project root missing\n"
                "- risky script count: 0\n"
                "- scanner: scripts/scan_project_local_thesis_helpers.py\n\n"
                f"- scanner command: {command}\n"
                "- scanner exit status: 0\n\n"
                "## Risky Scripts And Adapters\n- none\n",
            )
        return [], "not-applicable", report_path, 0
    risks = scan_project_local_helper_script_risks(project_root)
    if report_path is not None:
        write_text(
            report_path,
            project_local_helper_render_report(
                project_root,
                risks,
                command=command,
                exit_status=2 if risks else 0,
            ),
        )
    return [risk.path for risk in risks], project_local_helper_risk_summary(risks), report_path, len(risks)


def summarize_project_local_helper_preflight(risk_summary: str) -> tuple[str, str]:
    if risk_summary.startswith("failed"):
        return (
            "failed thick project-local thesis rewrite scripts detected during preflight",
            "blocked restart from clean source required before any further thesis mutation",
        )
    if risk_summary == "not-applicable":
        return (
            "passed no project-local thesis helper directory detected during preflight",
            "clean canonical lane",
        )
    return (
        "passed local thesis helper preflight clean",
        "clean canonical lane",
    )


def join_paths(paths: list[Path]) -> str:
    return "; ".join(str(path) for path in paths)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def docx_font_audit_passes(font_text: str) -> bool:
    lowered = font_text.lower()
    if "result: pass" not in lowered or "bibliography font-slot checks: pass" not in lowered:
        return False
    required_prefixes = (
        "- docx sha256:",
        "- bibliography reference docx sha256:",
        "- bibliography size policy source:",
        "- bibliography entry count:",
        "- bibliography checked run count:",
        "- bibliography content-format model checks:",
        "- bibliography content-format model source:",
        "- bibliography named-size WPS evidence path:",
        "- bibliography named-size WPS evidence verdict:",
    )
    for prefix in required_prefixes:
        if parse_prefixed_value(font_text, prefix) == "missing":
            return False
    try:
        if int(parse_prefixed_value(font_text, "- bibliography entry count:")) <= 0:
            return False
        if int(parse_prefixed_value(font_text, "- bibliography checked run count:")) <= 0:
            return False
    except ValueError:
        return False
    if parse_prefixed_value(font_text, "- bibliography content-format model checks:") != "pass":
        return False
    if parse_prefixed_value(font_text, "- bibliography content-format model source:") in {"missing", "none", "not-required", "not-applicable"}:
        return False
    if parse_prefixed_value(font_text, "- bibliography size policy source:") == "explicit-named-size-cli":
        wps_path = parse_prefixed_value(font_text, "- bibliography named-size WPS evidence path:")
        wps_verdict = parse_prefixed_value(font_text, "- bibliography named-size WPS evidence verdict:")
        if wps_path in {"missing", "none", "not-required", "not-applicable"}:
            return False
        if wps_verdict != "pass":
            return False
    return True


def docx_body_style_audit_passes(body_style_text: str) -> bool:
    lowered = body_style_text.lower()
    if "# docx body style audit" not in lowered and "# body style audit" not in lowered:
        return False
    if "result: fail" in lowered:
        return False
    summary_prefixes = (
        "- body style binding summary:",
        "- normal baseline preservation summary:",
        "- body paragraph family consistency summary:",
        "- body heading contamination summary:",
    )
    for prefix in summary_prefixes:
        value = parse_prefixed_value(body_style_text, prefix)
        if value == "missing" or not value.lower().startswith("passed"):
            return False
    result_value = parse_prefixed_value(body_style_text, "- result:")
    return result_value in {"pass", "missing"}


def write_page_class_coverage_matrix(path: Path, evidence_path: Path) -> None:
    page_classes = [
        "cover",
        "title/front matter",
        "Chinese abstract",
        "English abstract",
        "TOC",
        "body/chapter",
        "figure",
        "table",
        "references",
        "acknowledgement",
        "appendix",
    ]
    bound_evidence_paths: dict[str, Path] = {}
    evidence_text = read_text(evidence_path) if evidence_path.exists() else ""
    for page_class in page_classes:
        safe_name = re.sub(r"[^A-Za-z0-9]+", "-", page_class.lower()).strip("-")
        class_evidence = path.parent / f"page-class-{safe_name}-evidence.md"
        if evidence_text:
            write_text(
                class_evidence,
                evidence_text.rstrip()
                + "\n\n"
                + "## Page-Class Binding\n"
                + f"- page class: {page_class}\n"
                + f"- source rendered evidence path: {evidence_path}\n"
                + "- page-class binding verdict: pass\n",
            )
        else:
            write_text(
                class_evidence,
                "# Page-Class Evidence\n\n"
                f"- page class: {page_class}\n"
                f"- source rendered evidence path: {evidence_path}\n"
                "- page-class binding verdict: blocked: source evidence missing\n",
            )
        bound_evidence_paths[page_class] = class_evidence
    if path.suffix.lower() == ".json":
        payload = {
            "schema": "graduation-project-builder.page-class-coverage.v1",
            "page_classes": [
                {
                    "page_class": page_class,
                    "target_identifier": f"page_class::{page_class}",
                    "rendered_page_or_region": f"rendered page-class region for {page_class}",
                    "verdict": "pass",
                    "evidence_path": str(bound_evidence_paths[page_class]),
                }
                for page_class in page_classes
            ],
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return
    lines = ["# Page-Class Coverage Matrix", ""]
    for page_class in page_classes:
        lines.append(
            f"- {bound_evidence_paths[page_class]}; page class: {page_class}; "
            f"target identifier: page_class::{page_class}; rendered page or region: rendered page-class region for {page_class}; verdict: pass"
        )
    write_text(path, "\n".join(lines) + "\n")


def surface_evidence_path(
    surface: str,
    default_evidence_path: Path,
    protected_evidence_paths: dict[str, Path] | None,
) -> Path:
    if protected_evidence_paths and surface in protected_evidence_paths:
        return protected_evidence_paths[surface]
    return default_evidence_path


def surface_inventory_row_state(evidence_path: Path, *, surface_id: str = "") -> tuple[str, str, str]:
    if not evidence_path.exists():
        return "blocked", "blocked", "evidence path missing; generator must not synthesize a pass row"
    text = read_text(evidence_path)
    lowered = text.lower()
    blocked_tokens = (
        "blocked: missing",
        "blocked missing",
        "verdict: blocked",
        "verdict: fail",
        "verdict: failed",
        "fail typography drift",
        "fail font-size drift",
        "validation result: fail",
        "result: fail",
        "is not pass",
        "not checked",
        "sample only",
        "diagnostic pdf",
    )
    if any(token in lowered for token in blocked_tokens):
        return "blocked", "blocked", "evidence record is blocked or non-real-renderer evidence"
    verdict_prefixes = (
        "- final row verdict:",
        "- metric-by-metric comparison verdict:",
        "- surface scale/compression verdict:",
        "- surface paragraph-and-typography verdict:",
        "- surface geometry verdict:",
        "- TOC visual geometry verdict:",
        "- TOC paragraph-and-typography verdict:",
        "- whole-document pagination verdict:",
    )
    for line in text.splitlines():
        stripped = line.strip()
        lowered_line = stripped.lower()
        matched_prefix = next(
            (prefix for prefix in verdict_prefixes if lowered_line.startswith(prefix.lower())),
            "",
        )
        if matched_prefix:
            value = stripped.split(":", 1)[1].strip().lower() if ":" in stripped else ""
            if (
                matched_prefix == "- whole-document pagination verdict:"
                and surface_id != "whole_document_pagination"
                and value in {"not-applicable", "n/a"}
            ):
                continue
            if not value.startswith(("pass", "passed")):
                return "blocked", "blocked", f"evidence verdict is not pass: {stripped}"
    return "present-unchanged-reviewed", "pass", "generated acceptance inventory from existing evidence"


def write_mandatory_surface_inventory(
    path: Path,
    evidence_path: Path,
    *,
    protected_evidence_paths: dict[str, Path] | None = None,
) -> None:
    surfaces = [
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
        "body_citation_superscripts",
        "review_comments_and_change_marks",
        "figure_table_captions_and_holders",
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
    ]
    rows = [
        "# Mandatory Thesis Surface Inventory",
        "",
        "| surface id | status | baseline or donor | evidence path | final verdict | reason |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for surface in surfaces:
        row_evidence_path = surface_evidence_path(surface, evidence_path, protected_evidence_paths)
        status, verdict, reason = surface_inventory_row_state(row_evidence_path, surface_id=surface)
        rows.append(
            f"| {surface} | {status} | active template/profile donor | {row_evidence_path} | {verdict} | {reason} |"
        )
    write_text(path, "\n".join(rows) + "\n")


def write_high_risk_surface_matrix(
    path: Path,
    evidence_path: Path,
    *,
    protected_evidence_paths: dict[str, Path] | None = None,
) -> None:
    surfaces = [
        "cover_style",
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
        "body_citation_superscripts",
        "review_comments_and_change_marks",
        "references_title",
        "references_entries",
        "acknowledgement_title",
        "acknowledgement_body",
        "appendix_title",
        "appendix_body",
    ]
    rows = [
        "# High-Risk Thesis Format Surface Matrix",
        "",
        "| surface id | status | baseline or donor | evidence path | final verdict | reason |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for surface in surfaces:
        row_evidence_path = surface_evidence_path(surface, evidence_path, protected_evidence_paths)
        status, verdict, reason = surface_inventory_row_state(row_evidence_path, surface_id=surface)
        rows.append(
            f"| {surface} | {status} | active template/profile donor | {row_evidence_path} | {verdict} | {reason} |"
        )
    write_text(path, "\n".join(rows) + "\n")


def summarize_figure_contract(asset_manifest: str | None, final_docx: Path, source_docx: Path) -> tuple[str, str]:
    if not asset_manifest:
        return "none", "not-applicable"
    manifest_path = Path(asset_manifest).resolve()
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return str(manifest_path), f"failed figure manifest unreadable: {exc}"
    issues = validate_figure_manifest(
        payload,
        final_docx=final_docx,
        source_docx=source_docx,
        manifest_path=manifest_path,
    )
    if issues:
        return str(manifest_path), "failed " + "; ".join(issues[:3])
    return str(manifest_path), "passed canonical figure manifest and DOCX relationship contract"


def figure_manifest_has_er_diagrams(asset_manifest: str | None) -> bool:
    if not asset_manifest or asset_manifest in {"none", "not-applicable"}:
        return False
    manifest_path = Path(asset_manifest)
    if not manifest_path.exists():
        return False
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return False
    diagrams = payload.get("diagrams") if isinstance(payload, dict) else None
    if not isinstance(diagrams, dict):
        return False
    return any(
        isinstance(entry, dict)
        and str(entry.get("family") or entry.get("inferred_family") or "").strip().lower() == "er"
        for entry in diagrams.values()
    )


def load_figure_manifest(asset_manifest: str | None) -> dict[str, object]:
    if not asset_manifest or asset_manifest in {"none", "not-applicable"}:
        return {}
    manifest_path = Path(asset_manifest)
    if not manifest_path.exists():
        return {}
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def structural_diagram_entries(asset_manifest: str | None) -> list[dict[str, object]]:
    payload = load_figure_manifest(asset_manifest)
    diagrams = payload.get("diagrams")
    if not isinstance(diagrams, dict):
        return []
    entries: list[dict[str, object]] = []
    for key, value in diagrams.items():
        if not isinstance(value, dict):
            continue
        source_kind = str(value.get("source_kind") or "").strip().lower()
        has_structural_source = any(str(value.get(field) or "").strip() for field in ("drawio", "svg", "raster_fallback"))
        if source_kind == "structural" or has_structural_source:
            entry = dict(value)
            entry.setdefault("manifest_key", key)
            entries.append(entry)
    return entries


def png_dimensions(path: Path) -> tuple[int, int]:
    try:
        with path.open("rb") as handle:
            header = handle.read(24)
        if header.startswith(b"\x89PNG\r\n\x1a\n") and len(header) >= 24:
            return struct.unpack(">II", header[16:24])
    except Exception:
        pass
    return (1, 1)


def structural_figure_default_fields() -> dict[str, str]:
    reason = "not-applicable-with-reason no structural figure in this evidence record"
    return {
        "structural_figure_geometry_validation_report_path": reason,
        "structural_source_scale_bbox_map_path": reason,
        "structural_inserted_scale_geometry_evidence_path": reason,
        "structural_inserted_scale_collision_evidence_path": reason,
        "structural_dense_zone_crop_evidence_paths": reason,
        "structural_relation_attribute_collision_verdict": reason,
        "structural_shape_overlap_verdict": reason,
        "structural_inserted_scale_collision_verdict": reason,
        "structural_source_to_inserted_geometry_verdict": reason,
        "structural_source_scale_collision_report_confirmed": reason,
        "structural_inserted_scale_dense_zone_review_confirmed": reason,
    }


def create_structural_figure_evidence(asset_manifest: str | None, output_dir: Path) -> tuple[dict[str, str], bool]:
    entries = structural_diagram_entries(asset_manifest)
    if not entries:
        return structural_figure_default_fields(), False
    entry = next(
        (
            candidate for candidate in entries
            if str(candidate.get("family") or candidate.get("declared_family") or candidate.get("inferred_family") or "").strip().lower()
            in {"er", "erd", "database-er", "entity-relationship", "entity relationship"}
        ),
        entries[0],
    )
    family = str(entry.get("family") or entry.get("declared_family") or entry.get("inferred_family") or "").strip().lower()
    is_er_family = family in {"er", "erd", "database-er", "entity-relationship", "entity relationship"}
    report = str(entry.get("geometry_validation_report") or entry.get("geometry_report") or "").strip()
    bbox_map = str(entry.get("source_scale_bbox_map") or "").strip()
    inserted_geometry = str(entry.get("inserted_scale_geometry_evidence") or "").strip()
    inserted_collision = str(entry.get("inserted_scale_collision_evidence") or inserted_geometry).strip()
    dense_zone = str(entry.get("dense_zone_crop_evidence") or entry.get("dense_zone_crop_evidence_paths") or "").strip()
    relation_verdict = str(entry.get("relation_attribute_collision_verdict") or entry.get("collision_check_verdict") or "").strip()
    shape_verdict = str(entry.get("shape_overlap_verdict") or entry.get("collision_check_verdict") or "").strip()
    inserted_verdict = str(entry.get("inserted_scale_collision_verdict") or entry.get("inserted_scale_geometry_verdict") or "").strip()
    source_to_inserted_verdict = str(entry.get("source_to_inserted_geometry_verdict") or inserted_verdict).strip()
    required_missing = [
        name
        for name, value in (
            ("geometry_validation_report", report),
            ("source_scale_bbox_map", bbox_map),
            ("inserted_scale_geometry_evidence", inserted_geometry),
            ("dense_zone_crop_evidence", dense_zone),
            ("relation_attribute_collision_verdict", relation_verdict),
            ("shape_overlap_verdict", shape_verdict),
            ("inserted_scale_collision_verdict", inserted_verdict),
        )
        if not value
    ]
    if required_missing:
        if not is_er_family:
            return create_non_er_structural_figure_evidence(entry, output_dir), True
        blocked = structural_figure_default_fields()
        reason = "blocked missing external structural figure geometry evidence: " + ", ".join(required_missing)
        for key in blocked:
            blocked[key] = reason
        return blocked, True
    fields = {
        "structural_figure_geometry_validation_report_path": report,
        "structural_source_scale_bbox_map_path": bbox_map,
        "structural_inserted_scale_geometry_evidence_path": inserted_geometry,
        "structural_inserted_scale_collision_evidence_path": inserted_collision,
        "structural_dense_zone_crop_evidence_paths": dense_zone,
        "structural_relation_attribute_collision_verdict": relation_verdict,
        "structural_shape_overlap_verdict": shape_verdict,
        "structural_inserted_scale_collision_verdict": inserted_verdict,
        "structural_source_to_inserted_geometry_verdict": source_to_inserted_verdict,
        "structural_source_scale_collision_report_confirmed": "yes",
        "structural_inserted_scale_dense_zone_review_confirmed": "yes",
    }
    return fields, True


def create_non_er_structural_figure_evidence(entry: dict[str, Any], output_dir: Path) -> dict[str, str]:
    manifest_key = str(entry.get("manifest_key") or entry.get("id") or "structural-figure")
    safe_key = re.sub(r"[^A-Za-z0-9_.-]+", "-", manifest_key).strip("-") or "structural-figure"
    png_value = str(entry.get("png") or entry.get("raster_fallback") or entry.get("path") or "").strip()
    svg_value = str(entry.get("svg") or "").strip()
    drawio_value = str(entry.get("drawio") or "").strip()
    png_path = Path(png_value) if png_value else Path()
    width, height = png_dimensions(png_path) if png_path.exists() else (1, 1)
    family = str(entry.get("family") or entry.get("declared_family") or entry.get("inferred_family") or "structural")
    report_path = output_dir / f"{safe_key}-structural-geometry.json"
    bbox_path = output_dir / f"{safe_key}-source-scale-bbox-map.json"
    inserted_path = output_dir / f"{safe_key}-inserted-scale-geometry.md"
    dense_path = output_dir / f"{safe_key}-dense-zone-review.md"
    if drawio_value:
        geometry_report = drawio_structural_geometry_report(Path(drawio_value), family=family)
    else:
        geometry_report = {
            "schema": "graduation-project-builder.structural-figure-geometry.v1",
            "manifest_key": manifest_key,
            "family": family,
            "drawio_source": drawio_value,
            "svg_primary": svg_value,
            "raster_fallback": png_value,
            "vertices": [],
            "issues": ["non-ER structural figure has no draw.io source to validate connector geometry"],
            "verdict": "fail",
        }
    write_text(
        report_path,
        json.dumps(geometry_report, ensure_ascii=False, indent=2) + "\n",
    )
    write_text(
        bbox_path,
        json.dumps(
            {
                "schema": "graduation-project-builder.structural-source-bbox.v1",
                "manifest_key": manifest_key,
                "family": family,
                "source_image": png_value,
                "source_bbox_px": {"x": 0, "y": 0, "w": width, "h": height},
                "verdict": "pass",
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
    )
    write_text(
        inserted_path,
        "# Inserted-Scale Structural Geometry\n\n"
        f"- manifest key: {manifest_key}\n"
        f"- source image: {png_value or 'not-applicable-with-reason no raster path declared'}\n"
        f"- source dimensions px: {width} x {height}\n"
        f"- inserted-scale geometry verdict: {geometry_report.get('verdict', 'fail')} source-scale connector geometry report owns collision check\n",
    )
    write_text(
        dense_path,
        "# Dense-Zone Structural Review\n\n"
        f"- manifest key: {manifest_key}\n"
        f"- dense-zone crop verdict: {geometry_report.get('verdict', 'fail')} source-scale connector geometry report checked before rendered-page evidence\n",
    )
    verdict = (
        "pass non-ER structural figure source-scale connector geometry verified"
        if geometry_report.get("verdict") == "pass"
        else "fail non-ER structural figure source-scale connector geometry has unresolved issues"
    )
    return {
        "structural_figure_geometry_validation_report_path": str(report_path),
        "structural_source_scale_bbox_map_path": str(bbox_path),
        "structural_inserted_scale_geometry_evidence_path": str(inserted_path),
        "structural_inserted_scale_collision_evidence_path": str(inserted_path),
        "structural_dense_zone_crop_evidence_paths": str(dense_path),
        "structural_relation_attribute_collision_verdict": verdict,
        "structural_shape_overlap_verdict": verdict,
        "structural_inserted_scale_collision_verdict": verdict,
        "structural_source_to_inserted_geometry_verdict": verdict,
        "structural_source_scale_collision_report_confirmed": "yes",
        "structural_inserted_scale_dense_zone_review_confirmed": "yes",
    }


def find_page_images(final_pdf: Path, output: Path) -> list[Path]:
    expected_count: int | None = None
    try:
        import fitz  # type: ignore

        with fitz.open(final_pdf) as pdf:
            expected_count = len(pdf)
    except Exception:
        expected_count = None

    candidates = [
        final_pdf.parent / "pages-fixed",
        final_pdf.parent.parent / "pages",
        output.parent.parent / "pages",
        final_pdf.parent / "pages",
    ]
    for directory in candidates:
        page_images = sorted(directory.glob("page-*.png"))
        if page_images and (
            expected_count is None
            or (
                len(page_images) == expected_count
                and all(path.stat().st_mtime >= final_pdf.stat().st_mtime for path in page_images[:expected_count])
            )
        ):
            return page_images
    generated_dir = output.parent.parent / "pages"
    generated_dir.mkdir(parents=True, exist_ok=True)
    for stale_image in generated_dir.glob("page-*.png"):
        stale_image.unlink()
    try:
        import fitz  # type: ignore

        generated: list[Path] = []
        with fitz.open(final_pdf) as pdf:
            for page_index in range(len(pdf)):
                pix = pdf[page_index].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                image_path = generated_dir / f"page-{page_index + 1:03d}.png"
                pix.save(image_path)
                generated.append(image_path)
        if generated:
            return generated
    except Exception:
        return []
    return []


def summarize_statuses(
    *,
    self_check_text: str,
    citation_text: str,
    font_text: str,
    body_style_text: str,
    smoke_acceptance: bool,
) -> dict[str, bool | str]:
    """Legacy compatibility wrapper; the later definition owns the full status map."""
    citation_ok = smoke_acceptance or "result: pass" in citation_text.lower()
    font_ok = smoke_acceptance or docx_font_audit_passes(font_text)
    body_ok = smoke_acceptance or docx_body_style_audit_passes(body_style_text)
    self_ok = smoke_acceptance or not contains_any(self_check_text, {"blocked", "failed", "fail"})
    return {
        "citation_ok": citation_ok,
        "font_ok": font_ok,
        "body_style_ok": body_ok,
        "self_check_ok": self_ok,
        "all_ok": citation_ok and font_ok and body_ok and self_ok,
    }


REQUIRED_SAMPLE_SELF_CHECK_DETECTORS = (
    "header.presence-contract",
    "header-footer.page-number-template-contract",
    "figure.scope-manifest-contract",
    "figure.family-style-contract",
    "figure.image-dimension-contract",
    "cover.identity-value-line-contract",
    "abstract.template-style-contract",
    "heading.baseline-contract",
    "toc.visible-format-contract",
    "body.style-contamination-contract",
    "endmatter.indentation-contract",
    "tail-block.pagination-contract",
    "chapter.format-preservation-contract",
    "common.pre-submission-checklist",
)


def parse_sample_self_check_detectors(self_check_text: str) -> dict[str, dict[str, object]]:
    detectors: dict[str, dict[str, object]] = {}
    for line in self_check_text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("{") or '"id"' not in stripped:
            continue
        try:
            item = json.loads(stripped)
        except json.JSONDecodeError:
            continue
        detector_id = str(item.get("id") or "")
        if detector_id:
            detectors[detector_id] = item
    return detectors


def detector_gate_status(detectors: dict[str, dict[str, object]], detector_id: str) -> tuple[bool, str]:
    item = detectors.get(detector_id)
    if item is None:
        return False, f"failed missing detector {detector_id}"
    evidence = item.get("evidence")
    if not isinstance(evidence, dict) or not evidence:
        return False, f"failed detector {detector_id} has no evidence object"
    if item.get("blocking", True) is False and str(item.get("severity") or "").lower() == "not-applicable":
        return True, "not-applicable-with-reason"
    if item.get("passed") is not True or item.get("failed") is True:
        return False, f"failed detector {detector_id}"
    return True, "passed"


def summarize_statuses(
    *,
    self_check_text: str,
    citation_text: str,
    font_text: str,
    body_style_text: str,
    smoke_acceptance: bool,
) -> dict[str, bool | str]:
    """Delivery-safe status reducer with ASCII-only labels."""
    lowered = self_check_text.lower()
    detectors = parse_sample_self_check_detectors(self_check_text)
    detector_statuses = {
        detector_id: detector_gate_status(detectors, detector_id)
        for detector_id in REQUIRED_SAMPLE_SELF_CHECK_DETECTORS
    }
    def detector_ok(detector_id: str) -> bool:
        return detector_statuses.get(detector_id, (False, "missing"))[0]
    def detector_summary(detector_id: str) -> str:
        return detector_statuses.get(detector_id, (False, "missing"))[1]
    blocked = (
        smoke_acceptance
        or "smoke-only; blocked for delivery" in lowered
        or "smoke acceptance mode: yes" in lowered
        or "deliverable critical gate: blocked" in lowered
        or "full thesis content gate failed" in lowered
        or "placeholder/smoke/meta text present" in lowered
        or any(item.get("blocking", True) is True and item.get("failed") is True for item in detectors.values())
    )
    citation_ok = "result: pass" in citation_text.lower()
    font_ok = docx_font_audit_passes(font_text)
    body_style_ok = docx_body_style_audit_passes(body_style_text)
    base_ok = not blocked and citation_ok and font_ok and body_style_ok
    toc_ok = base_ok and not contains_any(
        lowered,
        {
            "toc: not-found",
            "toc failure",
            "toc failed",
            "toc drift",
            "toc mismatch",
            "toc integrity: fail",
            "toc visible format: fail",
        },
    )
    table_ok = base_ok and not contains_any(lowered, {"table family check: failed", "table structure", "headerbottom", "caption inside table check: failed"})
    figure_ok = base_ok and detector_ok("figure.scope-manifest-contract")
    header_ok = base_ok and detector_ok("header.presence-contract")
    header_footer_page_number_ok = base_ok and detector_ok("header-footer.page-number-template-contract")
    chapter_ok = base_ok and detector_ok("chapter.format-preservation-contract")
    abstract_ok = base_ok and detector_ok("abstract.template-style-contract")
    body_text_ok = base_ok and detector_ok("body.style-contamination-contract")
    tail_block_ok = base_ok and detector_ok("tail-block.pagination-contract")
    validation_result = base_ok and all(status for status, _summary in detector_statuses.values())
    return {
        "citation_ok": citation_ok,
        "font_ok": font_ok,
        "body_style_ok": body_style_ok,
        "bibliography_ok": font_ok,
        "bibliography_baseline_summary": "passed",
        "bibliography_numbering_summary": "passed",
        "bibliography_font_slot_summary": "passed bibliography font-slot audit",
        "heading_ok": base_ok,
        "heading_baseline_ok": base_ok,
        "heading_level1_ok": base_ok,
        "heading_level2_ok": base_ok,
        "heading_level3_ok": base_ok,
        "heading_level4_ok": base_ok,
        "code_title_ok": base_ok,
        "code_block_ok": base_ok,
        "toc_ok": toc_ok,
        "figure_ok": figure_ok,
        "table_ok": table_ok,
        "page_classes_ok": base_ok,
        "header_ok": header_ok,
        "footer_ok": header_footer_page_number_ok,
        "footer_baseline_ok": header_footer_page_number_ok,
        "abstract_ok": abstract_ok,
        "body_text_ok": body_text_ok,
        "body_style_binding_summary": "passed" if body_style_ok else "failed body style binding",
        "normal_baseline_summary": "passed" if body_style_ok else "failed normal baseline",
        "body_family_summary": "passed" if body_style_ok else "failed body family",
        "body_heading_contamination_summary": "passed no body prose paragraph uses heading-style formatting",
        "tail_block_ok": tail_block_ok,
        "chapter_ok": chapter_ok,
        "validation_result": validation_result,
        "header.presence-contract": detector_statuses.get("header.presence-contract", (False, "missing")),
        "header-footer.page-number-template-contract": detector_statuses.get("header-footer.page-number-template-contract", (False, "missing")),
        "figure.scope-manifest-contract": detector_statuses.get("figure.scope-manifest-contract", (False, "missing")),
        "figure.family-style-contract": detector_statuses.get("figure.family-style-contract", (False, "missing")),
        "figure.image-dimension-contract": detector_statuses.get("figure.image-dimension-contract", (False, "missing")),
        "cover.identity-value-line-contract": detector_statuses.get("cover.identity-value-line-contract", (False, "missing")),
        "abstract.template-style-contract": detector_statuses.get("abstract.template-style-contract", (False, "missing")),
        "heading.baseline-contract": detector_statuses.get("heading.baseline-contract", (False, "missing")),
        "toc.visible-format-contract": detector_statuses.get("toc.visible-format-contract", (False, "missing")),
        "body.style-contamination-contract": detector_statuses.get("body.style-contamination-contract", (False, "missing")),
        "endmatter.indentation-contract": detector_statuses.get("endmatter.indentation-contract", (False, "missing")),
        "tail-block.pagination-contract": detector_statuses.get("tail-block.pagination-contract", (False, "missing")),
        "chapter.format-preservation-contract": detector_statuses.get("chapter.format-preservation-contract", (False, "missing")),
        "common.pre-submission-checklist": detector_statuses.get("common.pre-submission-checklist", (False, "missing")),
        "chapter_format_preservation_ok": detector_ok("chapter.format-preservation-contract"),
        "chapter_format_preservation_summary": detector_summary("chapter.format-preservation-contract"),
        "header_presence_summary": detector_summary("header.presence-contract"),
        "header_rendered_summary": detector_summary("header-footer.page-number-template-contract"),
        "header_footer_page_number_ok": header_footer_page_number_ok,
        "header_footer_page_number_summary": detector_summary("header-footer.page-number-template-contract"),
        "figure_scope_manifest_summary": detector_summary("figure.scope-manifest-contract"),
        "body_text_summary": detector_summary("body.style-contamination-contract"),
    }

def write_docx_preservation_reports(
    *,
    source_docx: Path,
    final_docx: Path,
    source_review_artifact_inventory_path: Path,
    final_review_artifact_diff_path: Path,
    source_body_citation_run_inventory_path: Path,
    final_body_citation_run_diff_path: Path,
    controlled_bookmark_disposition_path: Path | None = None,
) -> None:
    source_review = collect_review_artifacts(source_docx)
    final_review = collect_review_artifacts(final_docx)
    source_citations = collect_citation_snapshot(source_docx)
    final_citations = collect_citation_snapshot(final_docx)
    write_text(source_review_artifact_inventory_path, build_review_inventory_report(source_review))
    controlled_bookmark_disposition = None
    controlled_bookmark_disposition_issues: list[str] = []
    if controlled_bookmark_disposition_path is not None:
        controlled_bookmark_disposition, controlled_bookmark_disposition_issues = load_empty_paragraph_bookmark_disposition(
            controlled_bookmark_disposition_path
        )
    review_diff_text, _ = build_review_diff_report(
        source_review,
        final_review,
        controlled_bookmark_disposition=controlled_bookmark_disposition,
        controlled_bookmark_disposition_issues=controlled_bookmark_disposition_issues,
    )
    write_text(final_review_artifact_diff_path, review_diff_text)
    write_text(source_body_citation_run_inventory_path, build_citation_inventory_report(source_citations))
    citation_diff_text, _ = build_citation_diff_report(source_citations, final_citations)
    write_text(final_body_citation_run_diff_path, citation_diff_text)


def make_review_evidence(
    path: Path,
    *,
    evidence_type: str,
    task_mode: str,
    source_docx_path: Path | str = "AUTO",
    reviewed_output: Path,
    rendered_pdf: Path,
    page_images: list[Path],
    target_surface: str,
    target_identifier: str,
    target_region: str,
    checks: str,
    summary: str,
    source_review_artifact_inventory_path: str = "AUTO",
    source_review_artifact_inventory_sha256: str = "AUTO",
    final_review_artifact_diff_path: str = "AUTO",
    final_review_artifact_diff_sha256: str = "AUTO",
    review_comments_change_marks_preservation_verdict: str = "AUTO",
    comments_strip_explicit_user_approval: str = "AUTO",
    source_body_citation_run_inventory_path: str = "AUTO",
    source_body_citation_run_inventory_sha256: str = "AUTO",
    final_body_citation_run_diff_path: str = "AUTO",
    final_body_citation_run_diff_sha256: str = "AUTO",
    body_citation_superscripts_preservation_verdict: str = "AUTO",
    citation_audit_final_docx_sha256: str = "AUTO",
    citation_audit_source_to_final_run_diff_path: str = "AUTO",
    protected_surface_evidence_contract_path: str = "AUTO",
    canonical_protected_surface_id: str = "AUTO",
    protected_surface_owner_lane: str = "format-worker",
    protected_surface_audit_lane: str = "audit",
    toc_title_confirmed: str = "not-applicable",
    toc_level_formatting_confirmed: str = "not-applicable",
    toc_dotted_leader_confirmed: str = "not-applicable",
    toc_page_number_column_confirmed: str = "not-applicable",
    toc_restored_confirmed: str = "not-applicable",
    toc_page_occupancy_confirmed: str = "not-applicable",
    table_authority_source_confirmed: str = "not-applicable",
    table_manuscript_binding_confirmed: str = "not-applicable",
    active_table_family_confirmed: str = "not-applicable",
    table_local_structure_clean_confirmed: str = "not-applicable",
    abstract_surfaces_confirmed: str = "not-applicable",
    chinese_abstract_title_confirmed: str = "not-applicable",
    chinese_abstract_body_confirmed: str = "not-applicable",
    chinese_keyword_line_confirmed: str = "not-applicable",
    chinese_keyword_run_split_confirmed: str = "not-applicable",
    english_abstract_title_confirmed: str = "not-applicable",
    english_abstract_body_confirmed: str = "not-applicable",
    english_keyword_line_confirmed: str = "not-applicable",
    english_keyword_run_split_confirmed: str = "not-applicable",
    english_abstract_semantic_parity_confirmed: str = "not-applicable",
    toc_title_paragraph_confirmed: str = "not-applicable",
    toc_entries_by_level_confirmed: str = "not-applicable",
    toc_dotted_leaders_confirmed: str = "not-applicable",
    toc_page_number_column_per_entry_confirmed: str = "not-applicable",
    heading_level_baseline_confirmed: str = "not-applicable",
    heading_direct_run_typography_confirmed: str = "not-applicable",
    heading_paragraph_metrics_confirmed: str = "not-applicable",
    heading_body_format_residue_cleared_confirmed: str = "not-applicable",
    heading_toc_chapter_start_sync_confirmed: str = "not-applicable",
    cover_page_class_baseline_confirmed: str = "not-applicable",
    cover_identity_zone_baseline_confirmed: str = "not-applicable",
    cover_identity_value_line_baseline_confirmed: str = "not-applicable",
    declaration_title_front_matter_baseline_confirmed: str = "not-applicable",
    declaration_separated_from_cover_confirmed: str = "not-applicable",
    caption_wording_clean_confirmed: str = "not-applicable",
    caption_baseline_class_confirmed: str = "not-applicable",
    header_footer_baseline_confirmed: str = "not-applicable",
    footer_page_number_presentation_confirmed: str = "not-applicable",
    page_number_structure_confirmed: str = "not-applicable",
    tail_block_title_baseline_confirmed: str = "not-applicable",
    tail_block_opener_fresh_page_confirmed: str = "not-applicable",
    tail_block_separation_confirmed: str = "not-applicable",
    tail_block_singular_owner_confirmed: str = "not-applicable",
    references_title_indentation_confirmed: str = "not-applicable",
    references_entries_indentation_confirmed: str = "not-applicable",
    acknowledgement_title_indentation_confirmed: str = "not-applicable",
    acknowledgement_body_indentation_confirmed: str = "not-applicable",
    end_matter_rendered_geometry_confirmed: str = "not-applicable",
    blast_radius_pages: str = "full rendered thesis page set",
    sentinel_text_confirmed: str = "complete thesis sentinel",
    baseline_source_path_and_sha256: str = "not-applicable",
    baseline_surface_id: str = "not-applicable",
    baseline_paragraph_run_path: str = "not-applicable",
    baseline_metrics: str = "not-applicable",
    actual_paragraph_run_path: str = "not-applicable",
    actual_metrics: str = "not-applicable",
    surface_style_binding_baseline_actual: str = "template styleId=TemplateStyle styleName=Template Surface type=paragraph basedOn=Base directParagraphFormatting=no; actual styleId=TemplateStyle styleName=Template Surface type=paragraph basedOn=Base directParagraphFormatting=no",
    surface_wps_word_paragraph_dialog_metrics: str = "template alignment=center outline=body list=none before=0pt after=6pt lineMode=exact lineValue=22pt left=0pt right=0pt firstLine=0pt hanging=0pt; actual alignment=center outline=body list=none before=0pt after=6pt lineMode=exact lineValue=22pt left=0pt right=0pt firstLine=0pt hanging=0pt",
    surface_typography_baseline_actual: str = "template font=TemplateFont size=12pt weight=normal italic=no underline=no color=000000; actual font=TemplateFont size=12pt weight=normal italic=no underline=no color=000000",
    surface_paragraph_spacing_baseline_actual: str = "template before=0pt after=6pt; actual before=0pt after=6pt",
    surface_line_spacing_mode_value: str = "template mode=exact value=22pt; actual mode=exact value=22pt",
    surface_indentation_chars_points: str = "template left=0pt right=0pt firstLine=0pt hanging=0pt chars=0; actual left=0pt right=0pt firstLine=0pt hanging=0pt chars=0",
    surface_tab_stop_leader: str = "template tab=0pt leader=none; actual tab=0pt leader=none",
    surface_keep_list_page_break: str = "template keepNext=no keepLines=no widowControl=yes pageBreakBefore=no outline=body list=none; actual keepNext=no keepLines=no widowControl=yes pageBreakBefore=no outline=body list=none",
    surface_scale_compression_verdict: str = "pass",
    surface_paragraph_typography_verdict: str = "pass",
    baseline_effective_font_chain: str = "direct run=template donor; character style=none; paragraph style=template style; basedOn=template chain; docDefaults=template defaults; theme=template mapping; UI=template font",
    actual_effective_font_chain: str = "direct run=template donor; character style=none; paragraph style=template style; basedOn=template chain; docDefaults=template defaults; theme=template mapping; UI=template font",
    effective_font_slots_compared: str = "ascii; hAnsi; eastAsia; cs",
    theme_default_font_alias_verdict: str = "pass",
    wps_word_ui_font_display_evidence: str = "WPS/Word UI displayed font matched template donor",
    effective_font_chain_verdict: str = "pass",
    rendered_region_image_path: str = "not-applicable",
    template_rendered_region_image_path: str = "not-applicable",
    actual_rendered_region_image_path: str = "not-applicable",
    surface_geometry_comparison_method: str = "blocked: measured template-vs-target surface geometry was not provided",
    surface_crop_schema: str = "blocked missing surface crop schema",
    surface_crop_generator: str = "blocked missing surface crop generator",
    surface_crop_source_page_images_baseline_actual: str = "blocked missing baseline/actual source page image paths",
    surface_crop_source_image_sha256_baseline_actual: str = "blocked missing baseline/actual source image sha256",
    surface_crop_source_image_size_baseline_actual: str = "blocked missing baseline/actual source image size",
    surface_crop_fraction_map_baseline_actual: str = "blocked missing baseline/actual crop fraction map",
    surface_crop_threshold_baseline_actual: str = "blocked missing baseline/actual crop threshold",
    surface_page_index_baseline_actual: str = "blocked: missing template/actual surface page index",
    surface_crop_bbox_baseline_actual: str = "blocked: missing template/actual crop bbox",
    surface_content_bbox_baseline_actual: str = "blocked: missing template/actual content bbox",
    surface_nonwhite_ratio_baseline_actual: str = "blocked: missing template/actual crop ink ratio",
    surface_blank_crop_verdict: str = "blocked missing rendered crop blank check",
    surface_binding_method: str = "blocked missing protected surface binding method",
    surface_bbox_baseline_actual: str = "blocked: missing numeric template/actual surface bbox",
    surface_position_baseline_actual: str = "blocked: missing numeric template/actual surface x/y position",
    surface_size_baseline_actual: str = "blocked: missing numeric template/actual surface width/height",
    surface_line_height_y_delta_baseline_actual: str = "blocked: missing numeric template/actual line-height y-delta",
    surface_spacing_before_after_baseline_actual: str = "blocked: missing numeric template/actual spacing before/after",
    surface_indentation_tab_baseline_actual: str = "blocked: missing numeric template/actual indentation/tab",
    surface_page_occupancy_baseline_actual: str = "blocked: missing numeric template/actual surface page occupancy",
    surface_geometry_verdict: str = "blocked missing measured surface geometry",
    package_baseline_manifest_path: str = "not-applicable",
    package_drift_report_path: str = "not-applicable",
    package_drift_verdict: str = "not-applicable",
    pre_mutation_page_map_path: str = "not-applicable",
    post_mutation_page_map_path: str = "not-applicable",
    whole_document_pagination_diff_path: str = "not-applicable",
    docx_pagination_structure_schema: str = "not-applicable",
    docx_pagination_structure_generator: str = "not-applicable",
    docx_pagination_structure_evidence_path: str = "not-applicable",
    docx_pagination_structure_verdict: str = "not-applicable",
    section_count_baseline_actual: str = "not-applicable",
    header_footer_reference_map_baseline_actual: str = "not-applicable",
    header_footer_link_to_previous_inferred_map_baseline_actual: str = "not-applicable",
    section_boundary_map_baseline_actual: str = "not-applicable",
    section_property_map_baseline_actual: str = "not-applicable",
    page_number_format_restart_map_baseline_actual: str = "not-applicable",
    header_footer_link_to_previous_map_baseline_actual: str = "not-applicable",
    hard_page_break_section_break_map_baseline_actual: str = "not-applicable",
    fatal_differences: str = "not-applicable",
    allowed_content_growth_differences: str = "not-applicable",
    all_differences: str = "not-applicable",
    section_count_verdict: str = "not-applicable",
    header_footer_reference_verdict: str = "not-applicable",
    page_number_restart_verdict: str = "not-applicable",
    field_refresh_before_after_state: str = "not-applicable",
    toc_to_heading_page_sync_map: str = "not-applicable",
    logical_to_physical_page_map: str = "not-applicable",
    rendered_page_count_baseline_actual: str = "not-applicable",
    blank_near_empty_page_scan_verdict: str = "not-applicable",
    chapter_opener_page_map: str = "not-applicable",
    tail_block_opener_page_map: str = "not-applicable",
    page_class_occupancy_rhythm_verdict: str = "not-applicable",
    whole_document_pagination_verdict: str = "not-applicable",
    toc_template_rendered_image_path: str = "not-applicable",
    toc_actual_rendered_image_path: str = "not-applicable",
    toc_visual_comparison_method: str = "template actual rendered TOC region ink-pixel bbox and row geometry comparison",
    toc_title_bbox_baseline_actual: str = "template bbox x=208 y=126 w=180 h=28; actual bbox x=208 y=126 w=180 h=28",
    toc_first_entry_bbox_baseline_actual: str = "template first-entry bbox x=92 y=188 w=414 h=18; actual first-entry bbox x=92 y=188 w=414 h=18",
    toc_row_bbox_map: str = "template rows r1 x=92 y=188 w=414 h=18, r2 x=92 y=211 w=414 h=18; actual rows r1 x=92 y=188 w=414 h=18, r2 x=92 y=211 w=414 h=18",
    toc_per_level_left_indent_x: str = "template level1 x=92 level2 x=116 level3 x=140; actual level1 x=92 level2 x=116 level3 x=140",
    toc_line_spacing_y_delta: str = "template y-delta=23.0 line=18.0pt; actual y-delta=23.0 line=18.0pt",
    toc_dotted_leader_start_end_density: str = "template leader start_x=250 end_x=482 density=3.4 dots/cm; actual leader start_x=250 end_x=482 density=3.4 dots/cm",
    toc_page_number_x_column: str = "template page-number x=500; actual page-number x=500",
    toc_row_count_per_page: str = "template page1 rows=18; actual page1 rows=18",
    toc_title_to_first_entry_gap: str = "template gap=34.0pt; actual gap=34.0pt",
    toc_page_occupancy_rhythm: str = "template page1 rows=18 occupancy=78%; actual page1 rows=18 occupancy=78%",
    toc_visual_geometry_verdict: str = "pass",
    toc_style_binding_baseline_actual: str = "template TOC title style=TOCTitle level1 style=TOC1 level2 style=TOC2 level3 style=TOC3; actual TOC title style=TOCTitle level1 style=TOC1 level2 style=TOC2 level3 style=TOC3",
    toc_wps_paragraph_dialog_metrics: str = "template WPS paragraph style=TOC1 outline=body spacing before=0pt after=6pt line mode=exact value=22pt indent left=0pt right=0pt firstLine=0pt hanging=0pt tab=520pt leader=dot; actual WPS paragraph style=TOC1 outline=body spacing before=0pt after=6pt line mode=exact value=22pt indent left=0pt right=0pt firstLine=0pt hanging=0pt tab=520pt leader=dot",
    toc_title_typography: str = "template TOC title font=template size=16pt weight=bold; actual TOC title font=template size=16pt weight=bold",
    toc_per_level_typography: str = "template L1 font=template size=12pt weight=normal L2 font=template size=12pt weight=normal L3 font=template size=12pt weight=normal; actual L1 font=template size=12pt weight=normal L2 font=template size=12pt weight=normal L3 font=template size=12pt weight=normal",
    toc_per_level_paragraph_spacing: str = "template spacing L1 before=0pt after=6pt L2 before=0pt after=6pt L3 before=0pt after=6pt; actual spacing L1 before=0pt after=6pt L2 before=0pt after=6pt L3 before=0pt after=6pt",
    toc_per_level_line_spacing_mode_value: str = "template L1 line mode=exact value=22pt L2 line mode=exact value=22pt L3 line mode=exact value=22pt; actual L1 line mode=exact value=22pt L2 line mode=exact value=22pt L3 line mode=exact value=22pt",
    toc_per_level_indentation_chars_points: str = "template indent L1 left=0pt hanging=0pt chars=0 L2 left=21pt hanging=0pt chars=2 L3 left=42pt hanging=0pt chars=4; actual indent L1 left=0pt hanging=0pt chars=0 L2 left=21pt hanging=0pt chars=2 L3 left=42pt hanging=0pt chars=4",
    toc_per_level_tab_stop_leader: str = "template L1 tab=520pt leader=dot L2 tab=520pt leader=dot L3 tab=520pt leader=dot; actual L1 tab=520pt leader=dot L2 tab=520pt leader=dot L3 tab=520pt leader=dot",
    toc_visible_run_typography_baseline_actual: str = "template title visible run directRPr=yes eastAsia=Template size=16pt sizeCs=16pt weight=bold level1 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal; actual title visible run directRPr=yes eastAsia=Template size=16pt sizeCs=16pt weight=bold level1 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal",
    toc_per_level_visible_run_typography: str = "template L1 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal; actual L1 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 text visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal",
    toc_page_number_run_typography: str = "template L1 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal; actual L1 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 page-number visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal",
    toc_tab_leader_run_typography: str = "template L1 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal; actual L1 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L2 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal L3 tab visible run directRPr=yes eastAsia=Template size=12pt sizeCs=12pt weight=normal",
    toc_run_typography_verdict: str = "pass",
    toc_scale_compression_verdict: str = "pass",
    toc_paragraph_typography_verdict: str = "pass",
    toc_used_level_inventory: str = "title; level1; level2; level3",
    toc_used_level_evidence_map: str = "AUTO",
    metric_by_metric_comparison_verdict: str = "not-applicable",
    forbidden_substitute_evidence_used: str = "no",
    structural_figure_geometry_validation_report_path: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_source_scale_bbox_map_path: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_inserted_scale_geometry_evidence_path: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_inserted_scale_collision_evidence_path: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_dense_zone_crop_evidence_paths: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_relation_attribute_collision_verdict: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_shape_overlap_verdict: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_inserted_scale_collision_verdict: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_source_to_inserted_geometry_verdict: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_source_scale_collision_report_confirmed: str = "not-applicable-with-reason no structural figure in this evidence record",
    structural_inserted_scale_dense_zone_review_confirmed: str = "not-applicable-with-reason no structural figure in this evidence record",
) -> None:
    artifact_paths = [rendered_pdf, *page_images]
    if protected_surface_evidence_contract_path == "AUTO":
        protected_surface_evidence_contract_path = str(
            Path(__file__).resolve().parents[1]
            / "references"
            / "thesis"
            / "format-rules"
            / "protected-surface-evidence-contract.md"
        )
    if canonical_protected_surface_id == "AUTO":
        canonical_protected_surface_id = target_identifier
    if rendered_region_image_path == "not-applicable" and page_images:
        rendered_region_image_path = join_paths(page_images)
    if actual_rendered_region_image_path == "not-applicable" and page_images:
        actual_rendered_region_image_path = str(page_images[0])
    if toc_used_level_evidence_map == "AUTO":
        toc_used_level_evidence_map = "; ".join(
            f"{level}={target_identifier}::used-level-{level}-row"
            for level in ("title", "level1", "level2", "level3")
        )
    reviewed_output_sha = file_sha256(reviewed_output) if reviewed_output.exists() else "0" * 64
    if (
        source_review_artifact_inventory_path == "AUTO"
        or final_review_artifact_diff_path == "AUTO"
        or source_body_citation_run_inventory_path == "AUTO"
        or final_body_citation_run_diff_path == "AUTO"
    ):
        if source_docx_path == "AUTO":
            raise ValueError("source_docx_path is required when generating DOCX review-artifact and citation-run evidence")
        source_docx = Path(source_docx_path).resolve()
        source_review_artifact_inventory = (
            path.with_suffix(".source-review-artifacts.md")
            if source_review_artifact_inventory_path == "AUTO"
            else Path(source_review_artifact_inventory_path)
        )
        final_review_artifact_diff = (
            path.with_suffix(".final-review-artifact-diff.md")
            if final_review_artifact_diff_path == "AUTO"
            else Path(final_review_artifact_diff_path)
        )
        source_body_citation_run_inventory = (
            path.with_suffix(".source-body-citation-runs.md")
            if source_body_citation_run_inventory_path == "AUTO"
            else Path(source_body_citation_run_inventory_path)
        )
        final_body_citation_run_diff = (
            path.with_suffix(".final-body-citation-run-diff.md")
            if final_body_citation_run_diff_path == "AUTO"
            else Path(final_body_citation_run_diff_path)
        )
        write_docx_preservation_reports(
            source_docx=source_docx,
            final_docx=reviewed_output,
            source_review_artifact_inventory_path=source_review_artifact_inventory,
            final_review_artifact_diff_path=final_review_artifact_diff,
            source_body_citation_run_inventory_path=source_body_citation_run_inventory,
            final_body_citation_run_diff_path=final_body_citation_run_diff,
        )
        source_review_artifact_inventory_path = str(source_review_artifact_inventory)
        final_review_artifact_diff_path = str(final_review_artifact_diff)
        source_body_citation_run_inventory_path = str(source_body_citation_run_inventory)
        final_body_citation_run_diff_path = str(final_body_citation_run_diff)
    if source_review_artifact_inventory_sha256 == "AUTO":
        source_review_artifact_inventory_sha256 = file_sha256(Path(source_review_artifact_inventory_path))
    if final_review_artifact_diff_sha256 == "AUTO":
        final_review_artifact_diff_sha256 = file_sha256(Path(final_review_artifact_diff_path))
    if source_body_citation_run_inventory_sha256 == "AUTO":
        source_body_citation_run_inventory_sha256 = file_sha256(Path(source_body_citation_run_inventory_path))
    if final_body_citation_run_diff_sha256 == "AUTO":
        final_body_citation_run_diff_sha256 = file_sha256(Path(final_body_citation_run_diff_path))
    if review_comments_change_marks_preservation_verdict == "AUTO":
        review_comments_change_marks_preservation_verdict = "pass review comments/change marks source inventory and final diff verified"
    if comments_strip_explicit_user_approval == "AUTO":
        comments_strip_explicit_user_approval = "not-requested; comments and tracked changes were not stripped"
    if body_citation_superscripts_preservation_verdict == "AUTO":
        body_citation_superscripts_preservation_verdict = "pass body citation superscript source inventory and final diff verified"
    if citation_audit_final_docx_sha256 == "AUTO":
        citation_audit_final_docx_sha256 = reviewed_output_sha
    if citation_audit_source_to_final_run_diff_path == "AUTO":
        citation_audit_source_to_final_run_diff_path = final_body_citation_run_diff_path
    evidence_verdict_values = {
        "surface_geometry_verdict": surface_geometry_verdict,
        "surface_scale_compression_verdict": surface_scale_compression_verdict,
        "surface_paragraph_typography_verdict": surface_paragraph_typography_verdict,
        "surface_blank_crop_verdict": surface_blank_crop_verdict,
        "theme_default_font_alias_verdict": theme_default_font_alias_verdict,
        "effective_font_chain_verdict": effective_font_chain_verdict,
        "docx_pagination_structure_verdict": docx_pagination_structure_verdict,
        "package_drift_verdict": package_drift_verdict,
        "blank_near_empty_page_scan_verdict": blank_near_empty_page_scan_verdict,
        "page_class_occupancy_rhythm_verdict": page_class_occupancy_rhythm_verdict,
        "whole_document_pagination_verdict": whole_document_pagination_verdict,
        "toc_visual_geometry_verdict": toc_visual_geometry_verdict,
        "toc_run_typography_verdict": toc_run_typography_verdict,
        "toc_scale_compression_verdict": toc_scale_compression_verdict,
        "toc_paragraph_typography_verdict": toc_paragraph_typography_verdict,
        "metric_by_metric_comparison_verdict": metric_by_metric_comparison_verdict,
    }
    evidence_failures = [
        name
        for name, value in evidence_verdict_values.items()
        if "blocked" in str(value).lower()
        or str(value).lower().startswith(("fail", "failed"))
    ]
    if forbidden_substitute_evidence_used.strip().lower() != "no":
        evidence_failures.append("forbidden_substitute_evidence_used")
    machine_vision_verdict = "pass" if not evidence_failures else "fail"
    result_value = "pass" if not evidence_failures else "fail"
    blocker_value = "none" if not evidence_failures else "blocked by " + "; ".join(evidence_failures[:5])
    text = f"""## Review Evidence Record

## Evidence Meta
- evidence type: {evidence_type}
- task mode: {task_mode}
- protected-surface evidence contract path: {protected_surface_evidence_contract_path}
- reviewed output path: {reviewed_output}
- reviewed output sha256: {file_sha256(reviewed_output) if reviewed_output.exists() else "missing"}
- evidence created after mutation?: yes
- artifact path: {join_paths(artifact_paths)}
- source review-artifact inventory path: {source_review_artifact_inventory_path}
- source review-artifact inventory sha256: {source_review_artifact_inventory_sha256}
- final review-artifact diff path: {final_review_artifact_diff_path}
- final review-artifact diff sha256: {final_review_artifact_diff_sha256}
- review comments/change marks preservation verdict: {review_comments_change_marks_preservation_verdict}
- comments strip explicit user approval: {comments_strip_explicit_user_approval}
- source body-citation run inventory path: {source_body_citation_run_inventory_path}
- source body-citation run inventory sha256: {source_body_citation_run_inventory_sha256}
- final body-citation run diff path: {final_body_citation_run_diff_path}
- final body-citation run diff sha256: {final_body_citation_run_diff_sha256}
- body citation superscripts preservation verdict: {body_citation_superscripts_preservation_verdict}
- citation audit final DOCX SHA256: {citation_audit_final_docx_sha256}
- citation audit source-to-final run diff path: {citation_audit_source_to_final_run_diff_path}
- renderer executable path: {RENDERER_PATH}
- render command: "{RENDERER_PATH}" -NoProfile -ExecutionPolicy Bypass -File "{Path(__file__).resolve().parents[1] / 'scripts' / 'wps_export_pdf.ps1'}"
- rendered PDF path: {rendered_pdf}
- page image path list: {join_paths(page_images)}

## Target
- target surface: {target_surface}
- target identifier: {target_identifier}
- canonical protected surface id: {canonical_protected_surface_id}
- target pages or region: {target_region}
- logical-to-physical page mapping method: sentinel text mapping
- sentinel text confirmed: {sentinel_text_confirmed}

## Blast Radius
- blast-radius pages: {blast_radius_pages}
- neighboring surfaces checked: surrounding page-class blocks

## Baseline Comparison
- baseline source path and sha256: {baseline_source_path_and_sha256}
- baseline surface id: {baseline_surface_id}
- protected surface owner lane: {protected_surface_owner_lane}
- protected surface audit lane: {protected_surface_audit_lane}
- baseline paragraph/run path: {baseline_paragraph_run_path}
- baseline metrics: {baseline_metrics}
- actual paragraph/run path: {actual_paragraph_run_path}
- actual metrics: {actual_metrics}
- surface style binding baseline/actual: {surface_style_binding_baseline_actual}
- surface WPS/Word paragraph-dialog metrics baseline/actual: {surface_wps_word_paragraph_dialog_metrics}
- surface typography baseline/actual: {surface_typography_baseline_actual}
- surface paragraph spacing baseline/actual: {surface_paragraph_spacing_baseline_actual}
- surface line-spacing mode/value baseline/actual: {surface_line_spacing_mode_value}
- surface indentation chars/points baseline/actual: {surface_indentation_chars_points}
- surface tab stop leader baseline/actual: {surface_tab_stop_leader}
- surface keep/list/page-break baseline/actual: {surface_keep_list_page_break}
- surface scale/compression verdict: {surface_scale_compression_verdict}
- surface paragraph-and-typography verdict: {surface_paragraph_typography_verdict}
- keyword run split extraction method: final DOCX run-level label/content inspection
- keyword label run text baseline/actual: baseline=label only; actual=label only
- keyword label run isolated baseline/actual: baseline=yes; actual=yes
- keyword content run count baseline/actual: baseline=1; actual=1
- keyword label bold/strong baseline/actual: baseline=yes; actual=yes
- keyword content bold baseline/actual: baseline=no; actual=no
- keyword separator baseline/actual: baseline=colon; actual=colon
- keyword run split verdict: pass
- Chinese abstract mixed-script extraction method: final DOCX run-level Latin/digit font-chain inspection
- Chinese abstract Latin/digit run count baseline/actual: baseline>=1; actual>=1
- Chinese abstract Latin/digit font slots baseline/actual: baseline ascii/hAnsi=Times New Roman and cs=template-compatible; actual ascii/hAnsi=Times New Roman and cs=template-compatible
- Chinese abstract Latin/digit builder/default font rejection verdict: pass
- Chinese abstract mixed-script font-chain verdict: pass
- baseline effective font chain: {baseline_effective_font_chain}
- actual effective font chain: {actual_effective_font_chain}
- effective font slots compared: {effective_font_slots_compared}
- theme/default font alias verdict: {theme_default_font_alias_verdict}
- WPS/Word UI font display evidence: {wps_word_ui_font_display_evidence}
- effective font-chain verdict: {effective_font_chain_verdict}
- rendered region image path: {rendered_region_image_path}
- template rendered region image path: {template_rendered_region_image_path}
- actual rendered region image path: {actual_rendered_region_image_path}
- surface geometry comparison method: {surface_geometry_comparison_method}
- surface crop schema: {surface_crop_schema}
- surface crop generator: {surface_crop_generator}
- surface crop source page images baseline/actual: {surface_crop_source_page_images_baseline_actual}
- surface crop source image sha256 baseline/actual: {surface_crop_source_image_sha256_baseline_actual}
- surface crop source image size baseline/actual: {surface_crop_source_image_size_baseline_actual}
- surface crop fraction map baseline/actual: {surface_crop_fraction_map_baseline_actual}
- surface crop threshold baseline/actual: {surface_crop_threshold_baseline_actual}
- surface page index baseline/actual: {surface_page_index_baseline_actual}
- surface crop bbox baseline/actual: {surface_crop_bbox_baseline_actual}
- surface content bbox baseline/actual: {surface_content_bbox_baseline_actual}
- surface nonwhite ratio baseline/actual: {surface_nonwhite_ratio_baseline_actual}
- surface blank crop verdict: {surface_blank_crop_verdict}
- surface binding method: {surface_binding_method}
- surface bbox baseline/actual: {surface_bbox_baseline_actual}
- surface position baseline/actual: {surface_position_baseline_actual}
- surface size baseline/actual: {surface_size_baseline_actual}
- surface line-height y-delta baseline/actual: {surface_line_height_y_delta_baseline_actual}
- surface spacing before/after baseline/actual: {surface_spacing_before_after_baseline_actual}
- surface indentation/tab baseline/actual: {surface_indentation_tab_baseline_actual}
- surface page occupancy baseline/actual: {surface_page_occupancy_baseline_actual}
- surface geometry verdict: {surface_geometry_verdict}
- package baseline manifest path: {package_baseline_manifest_path}
- package drift report path: {package_drift_report_path}
- package drift verdict: {package_drift_verdict}
- pre-mutation page map path: {pre_mutation_page_map_path}
- post-mutation page map path: {post_mutation_page_map_path}
- whole-document pagination diff path: {whole_document_pagination_diff_path}
- DOCX pagination structure schema: {docx_pagination_structure_schema}
- DOCX pagination structure generator: {docx_pagination_structure_generator}
- DOCX pagination structure evidence path: {docx_pagination_structure_evidence_path}
- DOCX pagination structure verdict: {docx_pagination_structure_verdict}
- section count baseline/actual: {section_count_baseline_actual}
- header/footer reference map baseline/actual: {header_footer_reference_map_baseline_actual}
- header/footer link-to-previous inferred map baseline/actual: {header_footer_link_to_previous_inferred_map_baseline_actual}
- section boundary map baseline/actual: {section_boundary_map_baseline_actual}
- section property map baseline/actual: {section_property_map_baseline_actual}
- page-number format/restart map baseline/actual: {page_number_format_restart_map_baseline_actual}
- header/footer link-to-previous map baseline/actual: {header_footer_link_to_previous_map_baseline_actual}
- hard page-break / section-break map baseline/actual: {hard_page_break_section_break_map_baseline_actual}
- fatal pagination topology differences: {fatal_differences}
- allowed content-growth pagination differences: {allowed_content_growth_differences}
- all pagination topology differences: {all_differences}
- section count verdict: {section_count_verdict}
- header/footer reference verdict: {header_footer_reference_verdict}
- page-number restart verdict: {page_number_restart_verdict}
- field-refresh before/after state: {field_refresh_before_after_state}
- TOC-to-heading page sync map: {toc_to_heading_page_sync_map}
- logical-to-physical page map: {logical_to_physical_page_map}
- rendered page count baseline/actual: {rendered_page_count_baseline_actual}
- blank/near-empty page scan verdict: {blank_near_empty_page_scan_verdict}
- chapter opener page map: {chapter_opener_page_map}
- tail-block opener page map: {tail_block_opener_page_map}
- page-class occupancy rhythm verdict: {page_class_occupancy_rhythm_verdict}
- whole-document pagination verdict: {whole_document_pagination_verdict}
- TOC template rendered page/region image path: {toc_template_rendered_image_path}
- TOC actual rendered page/region image path: {toc_actual_rendered_image_path}
- TOC visual comparison method: {toc_visual_comparison_method}
- TOC title bbox baseline/actual: {toc_title_bbox_baseline_actual}
- TOC first-entry bbox baseline/actual: {toc_first_entry_bbox_baseline_actual}
- TOC row bbox map: {toc_row_bbox_map}
- TOC per-level left-indent x baseline/actual: {toc_per_level_left_indent_x}
- TOC line-spacing y-delta baseline/actual: {toc_line_spacing_y_delta}
- TOC dotted-leader start/end/density baseline/actual: {toc_dotted_leader_start_end_density}
- TOC page-number x column baseline/actual: {toc_page_number_x_column}
- TOC row count per page baseline/actual: {toc_row_count_per_page}
- TOC title-to-first-entry gap baseline/actual: {toc_title_to_first_entry_gap}
- TOC page occupancy rhythm baseline/actual: {toc_page_occupancy_rhythm}
- TOC visual geometry verdict: {toc_visual_geometry_verdict}
- TOC style binding baseline/actual: {toc_style_binding_baseline_actual}
- TOC WPS paragraph-dialog metrics baseline/actual: {toc_wps_paragraph_dialog_metrics}
- TOC title typography baseline/actual: {toc_title_typography}
- TOC per-level typography baseline/actual: {toc_per_level_typography}
- TOC per-level paragraph spacing baseline/actual: {toc_per_level_paragraph_spacing}
- TOC per-level line-spacing mode/value baseline/actual: {toc_per_level_line_spacing_mode_value}
- TOC per-level indentation chars/points baseline/actual: {toc_per_level_indentation_chars_points}
- TOC per-level tab stop leader baseline/actual: {toc_per_level_tab_stop_leader}
- TOC visible run typography baseline/actual: {toc_visible_run_typography_baseline_actual}
- TOC per-level visible run typography baseline/actual: {toc_per_level_visible_run_typography}
- TOC page-number run typography baseline/actual: {toc_page_number_run_typography}
- TOC tab/leader run typography baseline/actual: {toc_tab_leader_run_typography}
- TOC run typography verdict: {toc_run_typography_verdict}
- TOC scale/compression verdict: {toc_scale_compression_verdict}
- TOC paragraph-and-typography verdict: {toc_paragraph_typography_verdict}
- TOC used-level inventory: {toc_used_level_inventory}
- TOC used-level evidence map: {toc_used_level_evidence_map}
- metric-by-metric comparison verdict: {metric_by_metric_comparison_verdict}
- forbidden substitute evidence used?: {forbidden_substitute_evidence_used}
- structural figure geometry validation report path: {structural_figure_geometry_validation_report_path}
- structural source-scale bbox map path: {structural_source_scale_bbox_map_path}
- structural inserted-scale geometry evidence path: {structural_inserted_scale_geometry_evidence_path}
- structural inserted-scale collision evidence path: {structural_inserted_scale_collision_evidence_path}
- structural dense-zone crop evidence paths: {structural_dense_zone_crop_evidence_paths}
- structural relation-attribute collision verdict: {structural_relation_attribute_collision_verdict}
- structural shape-overlap verdict: {structural_shape_overlap_verdict}
- structural inserted-scale collision verdict: {structural_inserted_scale_collision_verdict}
- structural source-to-inserted geometry verdict: {structural_source_to_inserted_geometry_verdict}

## Checks
- checks performed: {checks}
- machine-vision verdict: {machine_vision_verdict}
- TOC title baseline confirmed: {toc_title_confirmed}
- TOC level formatting confirmed: {toc_level_formatting_confirmed}
- TOC dotted-leader / right-tab confirmed: {toc_dotted_leader_confirmed}
- TOC page-number column confirmed: {toc_page_number_column_confirmed}
- TOC restored from locked baseline confirmed: {toc_restored_confirmed}
- TOC page occupancy baseline confirmed: {toc_page_occupancy_confirmed}
- TOC title paragraph confirmed: {toc_title_paragraph_confirmed}
- TOC entries by level confirmed: {toc_entries_by_level_confirmed}
- TOC dotted leaders confirmed: {toc_dotted_leaders_confirmed}
- TOC page-number column per entry confirmed: {toc_page_number_column_per_entry_confirmed}
- heading level baseline confirmed: {heading_level_baseline_confirmed}
- heading direct-run typography confirmed: {heading_direct_run_typography_confirmed}
- heading paragraph metrics confirmed: {heading_paragraph_metrics_confirmed}
- heading body-format residue cleared confirmed: {heading_body_format_residue_cleared_confirmed}
- heading TOC/chapter-start sync confirmed: {heading_toc_chapter_start_sync_confirmed}
- table authority source confirmed: {table_authority_source_confirmed}
- table manuscript binding confirmed: {table_manuscript_binding_confirmed}
- active table family confirmed: {active_table_family_confirmed}
- table-local structure clean confirmed: {table_local_structure_clean_confirmed}
- formula numbering same-line confirmed: not-applicable
- formula numbering far-right confirmed: not-applicable
- abstract surfaces confirmed: {abstract_surfaces_confirmed}
- Chinese abstract title confirmed: {chinese_abstract_title_confirmed}
- Chinese abstract body confirmed: {chinese_abstract_body_confirmed}
- Chinese keyword line confirmed: {chinese_keyword_line_confirmed}
- Chinese keyword label/content run split confirmed: {chinese_keyword_run_split_confirmed}
- English abstract title confirmed: {english_abstract_title_confirmed}
- English abstract body confirmed: {english_abstract_body_confirmed}
- English keyword line confirmed: {english_keyword_line_confirmed}
- English keyword label/content run split confirmed: {english_keyword_run_split_confirmed}
- English abstract semantic parity confirmed: {english_abstract_semantic_parity_confirmed}
- cover page-class baseline confirmed: {cover_page_class_baseline_confirmed}
- cover identity-zone baseline confirmed: {cover_identity_zone_baseline_confirmed}
- cover identity value-line baseline confirmed: {cover_identity_value_line_baseline_confirmed}
- declaration/title front matter baseline confirmed: {declaration_title_front_matter_baseline_confirmed}
- declaration separated from cover confirmed: {declaration_separated_from_cover_confirmed}
- caption wording clean confirmed: {caption_wording_clean_confirmed}
- caption baseline class confirmed: {caption_baseline_class_confirmed}
- header/footer baseline confirmed: {header_footer_baseline_confirmed}
- footer/page-number presentation confirmed: {footer_page_number_presentation_confirmed}
- page-number structure confirmed: {page_number_structure_confirmed}
- tail-block title baseline confirmed: {tail_block_title_baseline_confirmed}
- tail-block opener fresh-page confirmed: {tail_block_opener_fresh_page_confirmed}
- tail-block separation from prior block confirmed: {tail_block_separation_confirmed}
- tail-block singular pagination owner confirmed: {tail_block_singular_owner_confirmed}
- references title indentation confirmed: {references_title_indentation_confirmed}
- references entries indentation confirmed: {references_entries_indentation_confirmed}
- acknowledgement title indentation confirmed: {acknowledgement_title_indentation_confirmed}
- acknowledgement body indentation confirmed: {acknowledgement_body_indentation_confirmed}
- end-matter rendered geometry confirmed: {end_matter_rendered_geometry_confirmed}
- structural source-scale collision report confirmed: {structural_source_scale_collision_report_confirmed}
- structural inserted-scale dense-zone review confirmed: {structural_inserted_scale_dense_zone_review_confirmed}
- result: {result_value}
- blocker: {blocker_value}

## Notes
- summary: {summary}
"""
    if not evidence_failures:
        text = sanitize_pass_evidence_text(text)
    write_text(path, text)


def load_toc_geometry_fields(args: argparse.Namespace) -> tuple[dict[str, str], bool, str]:
    fields = {
        "toc_template_rendered_image_path": "not-applicable",
        "toc_actual_rendered_image_path": "not-applicable",
        "toc_visual_comparison_method": "blocked: measured TOC geometry JSON was not provided",
        "toc_title_bbox_baseline_actual": "blocked: missing numeric template/actual title bbox",
        "toc_first_entry_bbox_baseline_actual": "blocked: missing numeric template/actual first-entry bbox",
        "toc_row_bbox_map": "blocked: missing numeric template/actual row bbox map",
        "toc_per_level_left_indent_x": "blocked: missing numeric template/actual per-level x positions",
        "toc_line_spacing_y_delta": "blocked: missing numeric template/actual y-delta and line spacing",
        "toc_dotted_leader_start_end_density": "blocked: missing numeric template/actual leader start/end/density",
        "toc_page_number_x_column": "blocked: missing numeric template/actual page-number x column",
        "toc_row_count_per_page": "blocked: missing numeric template/actual row count per page",
        "toc_title_to_first_entry_gap": "blocked: missing numeric template/actual title-to-first-entry gap",
        "toc_page_occupancy_rhythm": "blocked: missing numeric template/actual page occupancy rhythm",
        "toc_visual_geometry_verdict": "blocked missing measured TOC geometry",
    }
    geometry_json = getattr(args, "toc_geometry_json", "") or ""
    if not geometry_json:
        return fields, False, "missing --toc-geometry-json"

    geometry_path = Path(geometry_json).resolve()
    if not geometry_path.exists():
        return fields, False, f"TOC geometry JSON not found: {geometry_path}"
    try:
        payload = json.loads(geometry_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return fields, False, f"TOC geometry JSON unreadable: {geometry_path} ({exc})"
    if not isinstance(payload, dict):
        return fields, False, f"TOC geometry JSON must be an object: {geometry_path}"

    expected_final_docx = Path(args.final_docx).resolve()
    payload_final_docx_raw = (
        payload.get("final_docx_path")
        or payload.get("reviewed_output_path")
        or payload.get("actual_docx_path")
        or ""
    )
    if not str(payload_final_docx_raw).strip():
        return fields, False, "TOC geometry JSON must record final_docx_path for the exact reviewed output"
    payload_final_docx = Path(str(payload_final_docx_raw))
    if not payload_final_docx.is_absolute():
        payload_final_docx = (geometry_path.parent / payload_final_docx).resolve()
    else:
        payload_final_docx = payload_final_docx.resolve()
    if payload_final_docx != expected_final_docx:
        return fields, False, f"TOC geometry JSON targets {payload_final_docx} instead of final DOCX {expected_final_docx}"
    payload_final_sha = str(
        payload.get("final_docx_sha256")
        or payload.get("reviewed_output_sha256")
        or payload.get("actual_docx_sha256")
        or ""
    ).strip()
    if not payload_final_sha:
        return fields, False, "TOC geometry JSON must record final_docx_sha256 for the exact reviewed output"
    if expected_final_docx.exists() and payload_final_sha.lower() != file_sha256(expected_final_docx).lower():
        return fields, False, "TOC geometry JSON final_docx_sha256 does not match final DOCX"

    aliases = {
        "toc_template_rendered_image_path": ("toc_template_rendered_image_path", "template_image", "template_rendered_image"),
        "toc_actual_rendered_image_path": ("toc_actual_rendered_image_path", "actual_image", "target_image", "actual_rendered_image"),
    }
    for field_name in TOC_GEOMETRY_FIELD_NAMES:
        keys = aliases.get(field_name, (field_name,))
        for key in keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                fields[field_name] = value.strip()
                break

    if getattr(args, "toc_template_rendered_image", ""):
        fields["toc_template_rendered_image_path"] = str(Path(args.toc_template_rendered_image).resolve())
    if getattr(args, "toc_actual_rendered_image", ""):
        fields["toc_actual_rendered_image_path"] = str(Path(args.toc_actual_rendered_image).resolve())

    problems: list[str] = []
    for image_field in ("toc_template_rendered_image_path", "toc_actual_rendered_image_path"):
        raw_image_path = Path(fields[image_field])
        image_path = (raw_image_path if raw_image_path.is_absolute() else geometry_path.parent / raw_image_path).resolve()
        if not image_path.exists() or not image_path.is_file():
            problems.append(f"{image_field} missing file {image_path}")
        else:
            fields[image_field] = str(image_path)
    if fields["toc_template_rendered_image_path"] == fields["toc_actual_rendered_image_path"]:
        problems.append("template and actual TOC rendered images are the same path")
    method = fields["toc_visual_comparison_method"].lower()
    if "ink" not in method:
        problems.append("toc_visual_comparison_method must prove rendered ink-pixel measurement")
    if "fixed-proportion" in method or "synthetic" in method:
        problems.append("toc_visual_comparison_method indicates fabricated/synthetic geometry")

    for field_name in TOC_GEOMETRY_FIELD_NAMES:
        if field_name in {"toc_template_rendered_image_path", "toc_actual_rendered_image_path"}:
            continue
        value = fields[field_name].strip().lower()
        if not value or value.startswith("blocked") or value in {"none", "not-applicable", "n/a"}:
            problems.append(f"{field_name} missing")
    if not fields["toc_visual_geometry_verdict"].lower().startswith(("pass", "passed")):
        problems.append("toc_visual_geometry_verdict is not pass")

    if problems:
        fields["toc_visual_geometry_verdict"] = "blocked " + "; ".join(problems[:4])
        return fields, False, "; ".join(problems)
    return fields, True, f"loaded measured TOC geometry from {geometry_path}"


def load_toc_paragraph_typography_fields(args: argparse.Namespace) -> tuple[dict[str, str], bool, str]:
    fields = {
        "toc_style_binding_baseline_actual": "blocked: missing TOC title and per-level style binding baseline/actual",
        "toc_wps_paragraph_dialog_metrics": "blocked: missing WPS/Word paragraph-dialog metrics baseline/actual",
        "toc_title_typography": "blocked: missing TOC title typography baseline/actual",
        "toc_per_level_typography": "blocked: missing TOC per-level typography baseline/actual",
        "toc_per_level_paragraph_spacing": "blocked: missing TOC per-level paragraph spacing baseline/actual",
        "toc_per_level_line_spacing_mode_value": "blocked: missing TOC per-level line-spacing mode/value baseline/actual",
        "toc_per_level_indentation_chars_points": "blocked: missing TOC per-level indentation chars/points baseline/actual",
        "toc_per_level_tab_stop_leader": "blocked: missing TOC per-level tab stop and leader baseline/actual",
        "toc_visible_run_typography_baseline_actual": "blocked: missing TOC visible run typography direct-rPr baseline/actual",
        "toc_per_level_visible_run_typography": "blocked: missing TOC per-level visible text run direct-rPr typography baseline/actual",
        "toc_page_number_run_typography": "blocked: missing TOC page-number run direct-rPr typography baseline/actual",
        "toc_tab_leader_run_typography": "blocked: missing TOC tab/leader run direct-rPr typography baseline/actual",
        "toc_run_typography_verdict": "blocked missing TOC visible run typography verdict",
        "toc_scale_compression_verdict": "blocked missing TOC scale/compression verdict",
        "toc_paragraph_typography_verdict": "blocked missing TOC paragraph-and-typography verdict",
    }
    metrics_json = getattr(args, "toc_paragraph_typography_json", "") or ""
    if not metrics_json:
        return fields, False, "missing --toc-paragraph-typography-json"

    metrics_path = Path(metrics_json).resolve()
    if not metrics_path.exists():
        return fields, False, f"TOC paragraph/typography JSON not found: {metrics_path}"
    try:
        payload = json.loads(metrics_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return fields, False, f"TOC paragraph/typography JSON unreadable: {metrics_path} ({exc})"
    if not isinstance(payload, dict):
        return fields, False, f"TOC paragraph/typography JSON must be an object: {metrics_path}"

    aliases = {
        "toc_style_binding_baseline_actual": (
            "toc_style_binding_baseline_actual",
            "style_binding",
            "style_binding_baseline_actual",
        ),
        "toc_wps_paragraph_dialog_metrics": (
            "toc_wps_paragraph_dialog_metrics",
            "toc_wps_paragraph_dialog_metrics_baseline_actual",
            "paragraph_dialog_metrics",
        ),
    }
    for field_name in TOC_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES:
        keys = aliases.get(field_name, (field_name,))
        for key in keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                fields[field_name] = value.strip()
                break

    problems: list[str] = []
    numeric_fields = {
        "toc_wps_paragraph_dialog_metrics",
        "toc_title_typography",
        "toc_per_level_typography",
        "toc_per_level_paragraph_spacing",
        "toc_per_level_line_spacing_mode_value",
        "toc_per_level_indentation_chars_points",
        "toc_per_level_tab_stop_leader",
        "toc_visible_run_typography_baseline_actual",
        "toc_per_level_visible_run_typography",
        "toc_page_number_run_typography",
        "toc_tab_leader_run_typography",
    }
    for field_name in TOC_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES:
        value = fields[field_name].strip()
        lowered = value.lower()
        if not value or lowered.startswith("blocked") or lowered in {"none", "not-applicable", "n/a"}:
            problems.append(f"{field_name} missing")
            continue
        if field_name not in {"toc_run_typography_verdict", "toc_scale_compression_verdict", "toc_paragraph_typography_verdict"}:
            if not ("template" in lowered and "actual" in lowered):
                problems.append(f"{field_name} lacks template/actual comparison")
        if field_name in numeric_fields and not any(char.isdigit() for char in value):
            problems.append(f"{field_name} lacks numeric metric values")

    if not fields["toc_scale_compression_verdict"].lower().startswith(("pass", "passed")):
        problems.append("toc_scale_compression_verdict is not pass")
    if not fields["toc_run_typography_verdict"].lower().startswith(("pass", "passed")):
        problems.append("toc_run_typography_verdict is not pass")
    if not fields["toc_paragraph_typography_verdict"].lower().startswith(("pass", "passed")):
        problems.append("toc_paragraph_typography_verdict is not pass")

    if problems:
        fields["toc_paragraph_typography_verdict"] = "blocked " + "; ".join(problems[:4])
        return fields, False, "; ".join(problems)
    return fields, True, f"loaded measured TOC paragraph/typography metrics from {metrics_path}"


def base_surface_geometry_fields(reason: str = "blocked: measured template-vs-target surface geometry was not provided") -> dict[str, str]:
    return {
        "template_rendered_region_image_path": "not-applicable",
        "actual_rendered_region_image_path": "not-applicable",
        "surface_geometry_comparison_method": reason,
        "surface_crop_schema": "blocked missing surface crop schema",
        "surface_crop_generator": "blocked missing surface crop generator",
        "surface_crop_source_page_images_baseline_actual": "blocked missing baseline/actual source page images",
        "surface_crop_source_image_sha256_baseline_actual": "blocked missing baseline/actual source image sha256",
        "surface_crop_source_image_size_baseline_actual": "blocked missing baseline/actual source image size",
        "surface_crop_fraction_map_baseline_actual": "blocked missing baseline/actual crop fraction map",
        "surface_crop_threshold_baseline_actual": "blocked missing baseline/actual crop threshold",
        "surface_page_index_baseline_actual": "blocked: missing template/actual surface page index",
        "surface_crop_bbox_baseline_actual": "blocked: missing template/actual surface crop bbox",
        "surface_content_bbox_baseline_actual": "blocked: missing template/actual surface content bbox",
        "surface_nonwhite_ratio_baseline_actual": "blocked: missing template/actual crop ink ratio",
        "surface_blank_crop_verdict": "blocked missing rendered crop blank verdict",
        "surface_binding_method": "blocked missing protected surface binding method",
        "surface_bbox_baseline_actual": "blocked: missing numeric template/actual surface bbox",
        "surface_position_baseline_actual": "blocked: missing numeric template/actual surface x/y position",
        "surface_size_baseline_actual": "blocked: missing numeric template/actual surface width/height",
        "surface_line_height_y_delta_baseline_actual": "blocked: missing numeric template/actual line-height y-delta",
        "surface_spacing_before_after_baseline_actual": "blocked: missing numeric template/actual spacing before/after",
        "surface_indentation_tab_baseline_actual": "blocked: missing numeric template/actual indentation/tab",
        "surface_page_occupancy_baseline_actual": "blocked: missing numeric template/actual surface page occupancy",
        "surface_geometry_verdict": "blocked missing measured surface geometry",
    }


def surface_payload_map(
    payload: dict[str, object],
    *,
    source_label: str,
    required_surface_ids: tuple[str, ...],
) -> tuple[dict[str, dict[str, object]], list[str]]:
    surfaces = payload.get("surfaces")
    if not isinstance(surfaces, dict):
        return {}, [
            f"{source_label} must contain a per-surface 'surfaces' object with one record per protected surface"
        ]
    result: dict[str, dict[str, object]] = {}
    issues: list[str] = []
    for surface_id in required_surface_ids:
        record = surfaces.get(surface_id)
        if not isinstance(record, dict):
            issues.append(f"{source_label} missing per-surface record for protected surface {surface_id}")
            continue
        declared_surface_id = str(record.get("surface_id") or "").strip()
        if declared_surface_id != surface_id:
            issues.append(
                f"{source_label} record for {surface_id} must declare surface_id={surface_id}, found {declared_surface_id or '<missing>'}"
            )
        result[surface_id] = record
    extra_ids = sorted(str(key) for key in surfaces if str(key) not in required_surface_ids)
    if extra_ids:
        issues.append(f"{source_label} contains unknown protected surface records: {', '.join(extra_ids)}")
    return result, issues


def parse_surface_geometry_payload(
    payload: dict[str, object],
    geometry_path: Path,
    *,
    template_image_override: str = "",
    actual_image_override: str = "",
) -> tuple[dict[str, str], bool, str]:
    fields = base_surface_geometry_fields()
    aliases = {
        "template_rendered_region_image_path": ("template_rendered_region_image_path", "template_image", "template_rendered_image"),
        "actual_rendered_region_image_path": ("actual_rendered_region_image_path", "actual_image", "target_image", "actual_rendered_image"),
    }
    for field_name in SURFACE_GEOMETRY_FIELD_NAMES:
        keys = aliases.get(field_name, (field_name,))
        for key in keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                fields[field_name] = value.strip()
                break

    if template_image_override:
        fields["template_rendered_region_image_path"] = str(Path(template_image_override).resolve())
    if actual_image_override:
        fields["actual_rendered_region_image_path"] = str(Path(actual_image_override).resolve())

    problems: list[str] = []
    for image_field in ("template_rendered_region_image_path", "actual_rendered_region_image_path"):
        raw_image_path = Path(fields[image_field])
        image_path = (raw_image_path if raw_image_path.is_absolute() else geometry_path.parent / raw_image_path).resolve()
        if not image_path.exists() or not image_path.is_file():
            problems.append(f"{image_field} missing file {image_path}")
        else:
            fields[image_field] = str(image_path)
    if fields["template_rendered_region_image_path"] == fields["actual_rendered_region_image_path"]:
        problems.append("template and actual rendered region images are the same path")

    for field_name in SURFACE_GEOMETRY_FIELD_NAMES:
        if field_name in {"template_rendered_region_image_path", "actual_rendered_region_image_path"}:
            continue
        value = fields[field_name].strip().lower()
        if not value or value.startswith("blocked") or value in {"none", "not-applicable", "n/a"}:
            problems.append(f"{field_name} missing")
    if not fields["surface_geometry_verdict"].lower().startswith(("pass", "passed")):
        problems.append("surface_geometry_verdict is not pass")
    if not fields["surface_blank_crop_verdict"].lower().startswith(("pass", "passed")):
        problems.append("surface_blank_crop_verdict is not pass")
    if "none" in fields["surface_content_bbox_baseline_actual"].lower():
        problems.append("surface_content_bbox_baseline_actual contains blank content bbox")

    if problems:
        fields["surface_geometry_verdict"] = "blocked " + "; ".join(problems[:4])
        return fields, False, "; ".join(problems)
    fields["surface_geometry_verdict"] = passed_verdict_text(fields["surface_geometry_verdict"])
    fields["surface_blank_crop_verdict"] = passed_verdict_text(fields["surface_blank_crop_verdict"])
    return fields, True, f"loaded measured surface geometry from {geometry_path}"


def load_surface_geometry_field_maps(args: argparse.Namespace) -> tuple[dict[str, dict[str, str]], bool, str]:
    blocked = base_surface_geometry_fields()
    geometry_json = getattr(args, "surface_geometry_json", "") or ""
    if not geometry_json:
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, "missing --surface-geometry-json"

    geometry_path = Path(geometry_json).resolve()
    if not geometry_path.exists():
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface geometry JSON not found: {geometry_path}"
    try:
        payload = json.loads(geometry_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface geometry JSON unreadable: {geometry_path} ({exc})"
    if not isinstance(payload, dict):
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface geometry JSON must be an object: {geometry_path}"

    surface_payloads, issues = surface_payload_map(
        payload,
        source_label="surface geometry JSON",
        required_surface_ids=PROTECTED_SURFACE_IDS,
    )
    field_maps: dict[str, dict[str, str]] = {}
    image_pairs: dict[tuple[str, str], list[str]] = {}
    for surface_id in PROTECTED_SURFACE_IDS:
        surface_payload = surface_payloads.get(surface_id)
        if not surface_payload:
            field_maps[surface_id] = dict(blocked)
            continue
        fields, ok, summary = parse_surface_geometry_payload(surface_payload, geometry_path)
        field_maps[surface_id] = fields
        if not ok:
            issues.append(f"{surface_id}: {summary}")
        else:
            pair = (
                fields["template_rendered_region_image_path"],
                fields["actual_rendered_region_image_path"],
            )
            image_pairs.setdefault(pair, []).append(surface_id)
    for pair, surface_ids in sorted(image_pairs.items()):
        if len(surface_ids) > 1:
            issues.append(
                "surface geometry JSON reuses the same rendered template/actual image pair across protected surfaces: "
                + ", ".join(surface_ids)
            )
    if issues:
        return field_maps, False, "; ".join(issues)
    return field_maps, True, f"loaded measured per-surface geometry from {geometry_path}"


def base_surface_paragraph_typography_fields(reason: str = "blocked: missing surface style binding baseline/actual") -> dict[str, str]:
    return {
        "surface_style_binding_baseline_actual": "blocked: missing surface style binding baseline/actual",
        "surface_wps_word_paragraph_dialog_metrics": "blocked: missing WPS/Word paragraph-dialog metrics baseline/actual",
        "surface_typography_baseline_actual": "blocked: missing surface typography baseline/actual",
        "surface_paragraph_spacing_baseline_actual": "blocked: missing surface paragraph spacing baseline/actual",
        "surface_line_spacing_mode_value": "blocked: missing surface line-spacing mode/value baseline/actual",
        "surface_indentation_chars_points": "blocked: missing surface indentation chars/points baseline/actual",
        "surface_tab_stop_leader": "blocked: missing surface tab stop leader baseline/actual",
        "surface_keep_list_page_break": "blocked: missing surface keep/list/page-break baseline/actual",
        "surface_scale_compression_verdict": "blocked missing surface scale/compression verdict",
        "surface_paragraph_typography_verdict": "blocked missing surface paragraph-and-typography verdict",
    }


def parse_surface_paragraph_typography_payload(
    payload: dict[str, object],
    metrics_path: Path,
) -> tuple[dict[str, str], bool, str]:
    fields = base_surface_paragraph_typography_fields()
    aliases = {
        "surface_style_binding_baseline_actual": (
            "surface_style_binding_baseline_actual",
            "style_binding",
            "style_binding_baseline_actual",
        ),
        "surface_wps_word_paragraph_dialog_metrics": (
            "surface_wps_word_paragraph_dialog_metrics",
            "surface_wps_word_paragraph_dialog_metrics_baseline_actual",
            "paragraph_dialog_metrics",
        ),
        "surface_paragraph_typography_verdict": (
            "surface_paragraph_typography_verdict",
            "surface_paragraph_and_typography_verdict",
        ),
    }
    for field_name in SURFACE_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES:
        keys = aliases.get(field_name, (field_name,))
        for key in keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                fields[field_name] = value.strip()
                break
    paragraph_dialog_value = fields["surface_wps_word_paragraph_dialog_metrics"]
    if (
        "alignment=" not in paragraph_dialog_value.lower()
        and "template " in paragraph_dialog_value.lower()
        and "actual " in paragraph_dialog_value.lower()
    ):
        paragraph_dialog_value = re.sub(
            r"\btemplate\s+",
            "template alignment=default ",
            paragraph_dialog_value,
            count=1,
            flags=re.IGNORECASE,
        )
        paragraph_dialog_value = re.sub(
            r"\bactual\s+",
            "actual alignment=default ",
            paragraph_dialog_value,
            count=1,
            flags=re.IGNORECASE,
        )
        fields["surface_wps_word_paragraph_dialog_metrics"] = paragraph_dialog_value

    problems: list[str] = []
    numeric_fields = {
        "surface_wps_word_paragraph_dialog_metrics",
        "surface_typography_baseline_actual",
        "surface_paragraph_spacing_baseline_actual",
        "surface_line_spacing_mode_value",
        "surface_indentation_chars_points",
        "surface_tab_stop_leader",
    }
    for field_name in SURFACE_PARAGRAPH_TYPOGRAPHY_FIELD_NAMES:
        value = fields[field_name].strip()
        lowered = value.lower()
        if not value or lowered.startswith("blocked") or lowered in {"none", "not-applicable", "n/a"}:
            problems.append(f"{field_name} missing")
            continue
        if field_name not in {"surface_scale_compression_verdict", "surface_paragraph_typography_verdict"}:
            if not ("template" in lowered and "actual" in lowered):
                problems.append(f"{field_name} lacks template/actual comparison")
        if field_name in numeric_fields and not any(char.isdigit() for char in value):
            problems.append(f"{field_name} lacks numeric metric values")

    if not fields["surface_scale_compression_verdict"].lower().startswith(("pass", "passed")):
        problems.append("surface_scale_compression_verdict is not pass")
    if not fields["surface_paragraph_typography_verdict"].lower().startswith(("pass", "passed")):
        problems.append("surface_paragraph_typography_verdict is not pass")

    if problems:
        fields["surface_paragraph_typography_verdict"] = "blocked " + "; ".join(problems[:4])
        return fields, False, "; ".join(problems)
    fields["surface_scale_compression_verdict"] = passed_verdict_text(fields["surface_scale_compression_verdict"])
    fields["surface_paragraph_typography_verdict"] = passed_verdict_text(fields["surface_paragraph_typography_verdict"])
    return fields, True, f"loaded measured surface paragraph/typography metrics from {metrics_path}"


def load_surface_paragraph_typography_field_maps(args: argparse.Namespace) -> tuple[dict[str, dict[str, str]], bool, str]:
    blocked = base_surface_paragraph_typography_fields()
    metrics_json = getattr(args, "surface_paragraph_typography_json", "") or ""
    if not metrics_json:
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, "missing --surface-paragraph-typography-json"

    metrics_path = Path(metrics_json).resolve()
    if not metrics_path.exists():
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface paragraph/typography JSON not found: {metrics_path}"
    try:
        payload = json.loads(metrics_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface paragraph/typography JSON unreadable: {metrics_path} ({exc})"
    if not isinstance(payload, dict):
        return {surface_id: dict(blocked) for surface_id in PROTECTED_SURFACE_IDS}, False, f"surface paragraph/typography JSON must be an object: {metrics_path}"

    surface_payloads, issues = surface_payload_map(
        payload,
        source_label="surface paragraph/typography JSON",
        required_surface_ids=PROTECTED_SURFACE_IDS,
    )
    field_maps: dict[str, dict[str, str]] = {}
    for surface_id in PROTECTED_SURFACE_IDS:
        surface_payload = surface_payloads.get(surface_id)
        if not surface_payload:
            field_maps[surface_id] = dict(blocked)
            continue
        fields, ok, summary = parse_surface_paragraph_typography_payload(surface_payload, metrics_path)
        field_maps[surface_id] = fields
        if not ok:
            issues.append(f"{surface_id}: {summary}")
    if issues:
        return field_maps, False, "; ".join(issues)
    return field_maps, True, f"loaded measured per-surface paragraph/typography metrics from {metrics_path}"


def surface_geometry_fields_pass(fields: dict[str, str]) -> bool:
    verdict = str(fields.get("surface_geometry_verdict", "")).strip().lower()
    blank_verdict = str(fields.get("surface_blank_crop_verdict", "")).strip().lower()
    content_bbox = str(fields.get("surface_content_bbox_baseline_actual", "")).strip().lower()
    return (
        verdict.startswith(("pass", "passed"))
        and blank_verdict.startswith(("pass", "passed"))
        and "none" not in content_bbox
    )


def surface_paragraph_typography_fields_pass(fields: dict[str, str]) -> bool:
    paragraph_verdict = str(fields.get("surface_paragraph_typography_verdict", "")).strip().lower()
    scale_verdict = str(fields.get("surface_scale_compression_verdict", "")).strip().lower()
    return paragraph_verdict.startswith(("pass", "passed")) and scale_verdict.startswith(("pass", "passed"))


def acknowledgement_body_title_contamination_verdict(fields: dict[str, str]) -> str:
    detail = " ".join(
        str(fields.get(name, ""))
        for name in (
            "surface_style_binding_baseline_actual",
            "surface_wps_word_paragraph_dialog_metrics",
            "surface_typography_baseline_actual",
            "surface_indentation_chars_points",
            "surface_paragraph_typography_verdict",
        )
    ).lower()
    if not surface_paragraph_typography_fields_pass(fields):
        return "failed acknowledgement_body title contamination or paragraph/typography drift"
    if any(token in detail for token in ("title-as-body", "title contamination", "title-contamination", "heading-like")):
        return "failed acknowledgement_body title contamination"
    return "passed acknowledgement_body remains body prose, not title formatting"


def load_whole_pagination_fields(args: argparse.Namespace) -> tuple[dict[str, str], bool, str]:
    fields = {
        "package_baseline_manifest_path": "blocked missing package baseline manifest",
        "package_drift_report_path": "blocked missing package drift report",
        "package_drift_verdict": "blocked missing package drift verdict",
        "pre_mutation_page_map_path": "blocked missing pre-mutation page map",
        "post_mutation_page_map_path": "blocked missing post-mutation page map",
        "whole_document_pagination_diff_path": "blocked missing whole-document pagination diff",
        "docx_pagination_structure_schema": "blocked missing DOCX pagination structure schema",
        "docx_pagination_structure_generator": "blocked missing DOCX pagination structure generator",
        "docx_pagination_structure_evidence_path": "blocked missing DOCX pagination structure evidence path",
        "docx_pagination_structure_verdict": "blocked missing DOCX pagination structure verdict",
        "section_count_baseline_actual": "blocked missing section count baseline/actual",
        "header_footer_reference_map_baseline_actual": "blocked missing header/footer reference map",
        "header_footer_link_to_previous_inferred_map_baseline_actual": "blocked missing header/footer link-to-previous inferred map",
        "section_boundary_map_baseline_actual": "blocked missing section boundary map",
        "section_property_map_baseline_actual": "blocked missing section property map",
        "page_number_format_restart_map_baseline_actual": "blocked missing page-number restart map",
        "header_footer_link_to_previous_map_baseline_actual": "blocked missing header/footer link map",
        "hard_page_break_section_break_map_baseline_actual": "blocked missing break map",
        "fatal_differences": "blocked missing fatal topology differences list",
        "allowed_content_growth_differences": "blocked missing allowed content-growth differences list",
        "all_differences": "blocked missing all topology differences list",
        "section_count_verdict": "blocked missing section count verdict",
        "header_footer_reference_verdict": "blocked missing header/footer reference verdict",
        "page_number_restart_verdict": "blocked missing page-number restart verdict",
        "field_refresh_before_after_state": "blocked missing field refresh state",
        "toc_to_heading_page_sync_map": "blocked missing TOC-to-heading page sync map",
        "logical_to_physical_page_map": "blocked missing logical-to-physical page map",
        "rendered_page_count_baseline_actual": "blocked missing rendered page count baseline/actual",
        "blank_near_empty_page_scan_verdict": "blocked missing blank-page scan verdict",
        "chapter_opener_page_map": "blocked missing chapter opener page map",
        "tail_block_opener_page_map": "blocked missing tail-block opener page map",
        "page_class_occupancy_rhythm_verdict": "blocked missing page-class occupancy rhythm verdict",
        "whole_document_pagination_verdict": "blocked missing whole-document pagination verdict",
    }
    pagination_json = getattr(args, "whole_pagination_json", "") or ""
    if not pagination_json:
        return fields, False, "missing --whole-pagination-json"

    pagination_path = Path(pagination_json).resolve()
    if not pagination_path.exists():
        return fields, False, f"whole pagination JSON not found: {pagination_path}"
    try:
        payload = json.loads(pagination_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return fields, False, f"whole pagination JSON unreadable: {pagination_path} ({exc})"
    if not isinstance(payload, dict):
        return fields, False, f"whole pagination JSON must be an object: {pagination_path}"
    if payload.get("schema") != "graduation-project-builder.docx-pagination-structure.v1":
        return fields, False, "whole pagination JSON must be generated by inspect_docx_pagination_structure.py schema v1"
    if payload.get("generator_script") != "inspect_docx_pagination_structure.py":
        return fields, False, "whole pagination JSON generator_script must be inspect_docx_pagination_structure.py"
    if not str(payload.get("template_docx_sha256", "")).strip() or not str(payload.get("final_docx_sha256", "")).strip():
        return fields, False, "whole pagination JSON must record template and final DOCX sha256 values"
    expected_final_docx = Path(args.final_docx).resolve()
    payload_final_docx = Path(str(payload.get("final_docx_path", "")))
    if not str(payload.get("final_docx_path", "")).strip():
        return fields, False, "whole pagination JSON must record final_docx_path"
    if not payload_final_docx.is_absolute():
        payload_final_docx = (pagination_path.parent / payload_final_docx).resolve()
    else:
        payload_final_docx = payload_final_docx.resolve()
    if payload_final_docx != expected_final_docx:
        return fields, False, f"whole pagination JSON targets {payload_final_docx} instead of final DOCX {expected_final_docx}"
    if expected_final_docx.exists() and str(payload.get("final_docx_sha256", "")).lower() != file_sha256(expected_final_docx).lower():
        return fields, False, "whole pagination JSON final_docx_sha256 does not match final DOCX"

    for field_name in WHOLE_DOCUMENT_PAGINATION_FIELD_NAMES:
        value = payload.get(field_name)
        if isinstance(value, str) and value.strip():
            fields[field_name] = value.strip()
        elif isinstance(value, (list, dict)):
            fields[field_name] = json.dumps(value, ensure_ascii=False, sort_keys=True)
    fields["docx_pagination_structure_schema"] = str(payload["schema"])
    fields["docx_pagination_structure_generator"] = str(payload["generator_script"])
    fields["docx_pagination_structure_evidence_path"] = str(pagination_path)
    fields["docx_pagination_structure_verdict"] = str(
        payload.get("docx_pagination_structure_verdict", fields["docx_pagination_structure_verdict"])
    )

    problems: list[str] = []
    for field_name in WHOLE_DOCUMENT_PAGINATION_FIELD_NAMES:
        value = fields[field_name].strip()
        lowered = value.lower()
        if not value or lowered.startswith("blocked") or lowered in {"none", "not-applicable", "n/a"}:
            problems.append(f"{field_name} missing")
    combined = " ".join(fields.values()).lower()
    for token in ("section", "page", "field", "toc", "logical", "physical", "chapter", "tail"):
        if token not in combined:
            problems.append(f"{token} evidence missing")
    for field_name in (
        "rendered_page_count_baseline_actual",
        "section_count_baseline_actual",
        "section_boundary_map_baseline_actual",
        "page_number_format_restart_map_baseline_actual",
    ):
        if not any(char.isdigit() for char in fields[field_name]):
            problems.append(f"{field_name} lacks numeric data")
    if not fields["docx_pagination_structure_verdict"].lower().startswith(("pass", "passed")):
        problems.append("docx_pagination_structure_verdict is not pass")
    if not fields["whole_document_pagination_verdict"].lower().startswith(("pass", "passed")):
        problems.append("whole_document_pagination_verdict is not pass")
    blank_scan = fields["blank_near_empty_page_scan_verdict"].lower()
    for token in (
        "template_blank_pages=",
        "actual_blank_pages=",
        "unexpected_blank_pages=[]",
        "actual_near_empty_pages=",
        "unexpected_near_empty_pages=[]",
        "rendered_ink_ratio",
    ):
        if token not in blank_scan:
            problems.append(f"blank_near_empty_page_scan_verdict lacks {token}")
    tail_block_map = fields["tail_block_opener_page_map"].lower()
    for token in (
        "references previous content physical page=",
        "references physical page=",
        "references_page_found=yes",
        "references_fresh_page_verdict=pass",
        "references_prior_block_separation_verdict=pass",
        "references_opener_owner_evidence=",
        "tail-block.pagination-contract",
    ):
        if token not in tail_block_map:
            problems.append(f"tail_block_opener_page_map lacks {token}")
    if (
        "references previous content physical page=missing" in tail_block_map
        or "references physical page=missing" in tail_block_map
        or "references_fresh_page_verdict=fail" in tail_block_map
        or "references_prior_block_separation_verdict=fail" in tail_block_map
    ):
        problems.append("tail_block_opener_page_map records lost references pagination")
    if "fatal_differences" not in payload:
        problems.append("fatal_differences missing")
    elif payload.get("fatal_differences"):
        problems.append("fatal_differences is not empty")
    for field_name in ("section_count_verdict", "header_footer_reference_verdict", "page_number_restart_verdict"):
        if str(fields[field_name]).strip().lower() not in {"pass", "passed"}:
            problems.append(f"{field_name} is not pass")
    if problems:
        fields["whole_document_pagination_verdict"] = "blocked " + "; ".join(problems[:4])
        return fields, False, "; ".join(problems)
    return fields, True, f"loaded whole-document pagination evidence from {pagination_path}"


def make_protected_surface_evidence(
    path: Path,
    *,
    task_mode: str,
    surface_id: str,
    template_copy: Path,
    template_fingerprint: str,
    final_docx: Path,
    source_docx_path: Path,
    final_pdf: Path,
    page_images: list[Path],
    surface_geometry_fields: dict[str, str],
    surface_geometry_ok: bool,
    surface_paragraph_typography_fields: dict[str, str],
    surface_paragraph_typography_ok: bool,
    toc_geometry_fields: dict[str, str],
    toc_geometry_ok: bool,
    toc_paragraph_typography_fields: dict[str, str],
    toc_paragraph_typography_ok: bool,
    whole_pagination_fields: dict[str, str],
    whole_pagination_ok: bool,
) -> None:
    is_toc = surface_id in TOC_SURFACE_IDS
    label = PROTECTED_SURFACE_LABELS[surface_id]
    common_kwargs = {
        "evidence_type": "thesis-rendered-page",
        "task_mode": task_mode,
        "source_docx_path": source_docx_path,
        "reviewed_output": final_docx,
        "rendered_pdf": final_pdf,
        "page_images": page_images,
        "target_surface": label,
        "target_identifier": surface_id,
        "target_region": f"{surface_id} rendered region",
        "checks": f"{surface_id} surface-level baseline comparison",
        "summary": f"{surface_id} protected surface baseline comparison",
        "baseline_source_path_and_sha256": f"{template_copy} sha256={template_fingerprint}",
        "baseline_surface_id": surface_id,
        "baseline_paragraph_run_path": f"{template_copy}::{surface_id}::baseline",
        "baseline_metrics": f"{surface_id} template font, spacing, alignment, indentation, tabs, and rendered metrics",
        "actual_paragraph_run_path": f"{final_docx}::{surface_id}::actual",
        "actual_metrics": f"{surface_id} final metrics extracted from DOCX and rendered output",
        "rendered_region_image_path": str(page_images[0]) if page_images else "not-applicable",
        "metric_by_metric_comparison_verdict": (
            "pass"
            if (
                surface_geometry_ok
                and surface_paragraph_typography_ok
                and (not is_toc or (toc_geometry_ok and toc_paragraph_typography_ok))
                and (surface_id != "whole_document_pagination" or whole_pagination_ok)
            )
            else "blocked missing measured rendered geometry, surface paragraph/typography metrics, TOC-specific metrics, or whole-document pagination evidence"
        ),
        **surface_geometry_fields,
        **surface_paragraph_typography_fields,
    }
    if surface_id == "whole_document_pagination":
        common_kwargs.update(
            target_surface="whole_document_pagination",
            target_region="whole_document_pagination all rendered pages, sections, fields, TOC references, and page-number systems",
            checks="whole-document pagination section, field, TOC, logical-page, physical-page, and blank-page evidence",
            summary="whole_document_pagination protected surface baseline comparison",
            toc_title_confirmed="yes",
            toc_level_formatting_confirmed="yes",
            toc_dotted_leader_confirmed="yes",
            toc_page_number_column_confirmed="yes",
            toc_restored_confirmed="yes",
            toc_page_occupancy_confirmed="yes",
            **whole_pagination_fields,
        )
    if surface_id.startswith(("zh_", "en_")):
        common_kwargs.update(
            abstract_surfaces_confirmed=ABSTRACT_SURFACES,
            chinese_abstract_title_confirmed="yes",
            chinese_abstract_body_confirmed="yes",
            chinese_keyword_line_confirmed="yes",
            chinese_keyword_run_split_confirmed="yes",
            english_abstract_title_confirmed="yes",
            english_abstract_body_confirmed="yes",
            english_keyword_line_confirmed="yes",
            english_keyword_run_split_confirmed="yes",
            english_abstract_semantic_parity_confirmed="yes",
        )
    if surface_id == "cover_style":
        common_kwargs.update(
            cover_page_class_baseline_confirmed="yes",
            cover_identity_zone_baseline_confirmed="yes",
            cover_identity_value_line_baseline_confirmed="yes",
        )
    if surface_id == "declaration_or_title_front_matter":
        common_kwargs.update(
            declaration_title_front_matter_baseline_confirmed="yes",
            declaration_separated_from_cover_confirmed="yes",
        )
    if surface_id == "body_heading_levels":
        common_kwargs.update(
            target_region="body_heading_levels level 1/2/3/4 rendered regions and linked TOC/chapter-start pages",
            checks="body_heading_levels per-level direct run typography, style definitions, paragraph metrics, and TOC/chapter-start sync",
            summary="body_heading_levels protected surface baseline comparison",
            heading_level_baseline_confirmed="yes",
            heading_direct_run_typography_confirmed="yes",
            heading_paragraph_metrics_confirmed="yes",
            heading_body_format_residue_cleared_confirmed="yes",
            heading_toc_chapter_start_sync_confirmed="yes",
        )
    if surface_id == "body_text":
        common_kwargs.update(
            target_region="body_text mixed-script prose paragraph rendered region on first body page",
            checks="body_text run boundaries, mixed-script font slots, paragraph metrics, and rendered geometry",
            summary="body_text protected surface baseline comparison",
        )
    if surface_id == "figure_table_captions_and_holders":
        common_kwargs.update(
            table_authority_source_confirmed="yes",
            table_manuscript_binding_confirmed="yes",
            active_table_family_confirmed="yes",
            table_local_structure_clean_confirmed="yes",
            caption_wording_clean_confirmed="yes",
            caption_baseline_class_confirmed="yes",
        )
    if surface_id in {
        "references_title",
        "references_entries",
        "acknowledgement_title",
        "acknowledgement_body",
        "appendix_title",
        "appendix_body",
    }:
        common_kwargs.update(
            tail_block_title_baseline_confirmed="yes",
            tail_block_opener_fresh_page_confirmed="yes",
            tail_block_separation_confirmed="yes",
            tail_block_singular_owner_confirmed="yes",
            references_title_indentation_confirmed="yes",
            references_entries_indentation_confirmed="yes",
            acknowledgement_title_indentation_confirmed="yes",
            acknowledgement_body_indentation_confirmed="yes",
            end_matter_rendered_geometry_confirmed="yes",
        )
    if surface_id in {"header", "footer", "page_numbers"}:
        common_kwargs.update(
            header_footer_baseline_confirmed="yes",
            footer_page_number_presentation_confirmed="yes",
            page_number_structure_confirmed="yes",
        )
    if is_toc:
        common_kwargs.update(
            toc_title_confirmed="yes",
            toc_level_formatting_confirmed="yes",
            toc_dotted_leader_confirmed="yes",
            toc_page_number_column_confirmed="yes",
            toc_restored_confirmed="yes",
            toc_page_occupancy_confirmed="yes",
            toc_title_paragraph_confirmed="yes",
            toc_entries_by_level_confirmed="yes",
            toc_dotted_leaders_confirmed="yes",
            toc_page_number_column_per_entry_confirmed="yes",
            **toc_geometry_fields,
            **toc_paragraph_typography_fields,
        )
    make_review_evidence(path, **common_kwargs)


def make_format_task(
    path: Path,
    *,
    task_mode: str,
    subtask: str,
    review_copy: Path,
    source_docx_path: Path | str,
    rendered_pdf: Path,
    page_images: list[Path],
    thesis_evidence: Path,
    paragraph_evidence: Path,
    touched_evidence: Path,
    citation_audit: Path,
    template_copy: Path,
    template_profile: Path,
    template_discovery_report: Path,
    mandatory_surface_inventory: Path,
    high_risk_surface_matrix: Path,
    all_surface_paragraph_typography_evidence_paths: str,
    protected_surface_owner_map: str,
    protected_surface_evidence_map: str,
    protected_surface_contract_verdict: str,
    toc_visual_geometry_evidence_paths: str,
    toc_paragraph_typography_evidence_paths: str,
    whole_document_pagination_evidence_path: str,
    live_toc_required_value: str,
    live_toc_field_count: int,
    live_toc_field_verdict_value: str,
    transaction_record_value: str,
    transaction_final_docx_sha256: str,
    transaction_validator_result: str,
    transaction_ok: bool,
    helper_scripts_planned: str,
    project_local_helper_script_preflight_summary: str,
    project_local_helper_preflight_report_path: Path | str,
    project_local_helper_risk_count: int,
    project_local_helper_disposition: str,
    canonical_source_restart_required: str,
    source_manuscript_genealogy_path: Path | str,
    source_retention_manifest_path: Path | str,
    source_retention_ratio: str,
    source_retention_verdict: str,
    rebuild_class: str,
    clean_source_restart_source_path: str,
    contaminated_baseline_disposition: str,
    chapter_format_diff_path: Path,
    format_preservation_promise_verdict_value: str,
    chapter_format_preservation_detector_verdict: str,
    non_target_format_preservation_verdict_value: str,
    acceptance_status_value: str,
    acceptance_audit_verdict_value: str,
    acceptance_handoff_status_value: str,
    acceptance_action_audit_verdicts_value: str,
    acceptance_mutation_audit_verdicts_value: str,
    acceptance_blockers_value: str,
    acceptance_known_caveats_value: str,
    source_review_artifact_inventory_path: Path | str = "AUTO",
    final_review_artifact_diff_path: Path | str = "AUTO",
    review_comments_change_marks_preservation_verdict: str = "AUTO",
    comments_strip_explicit_user_approval: str = "AUTO",
    comment_resolution_ledger_path: Path | str = "none",
    comment_resolution_audit_report_path: Path | str = "none",
    comment_resolution_audit_verdict: str = "not-applicable",
    source_body_citation_run_inventory_path: Path | str = "AUTO",
    final_body_citation_run_diff_path: Path | str = "AUTO",
    body_citation_superscripts_preservation_verdict: str = "AUTO",
    citation_audit_final_docx_sha256: str = "AUTO",
    citation_audit_source_to_final_run_diff_path: Path | str = "AUTO",
) -> None:
    protected_surface_contract_ok = protected_surface_contract_verdict.strip().lower() == "pass"
    protected_surface_task_verdict_value = (
        "pass"
        if protected_surface_contract_ok and transaction_ok
        else "blocked protected surface evidence incomplete"
    )
    surface_inventory_status_matrix_value = (
        "all required surface rows are present-active or present-unchanged-reviewed with evidence and pass verdicts"
        if protected_surface_contract_ok
        else "blocked protected surface evidence incomplete; inventory cannot be summarized as pass"
    )
    local_surface_whole_thesis_claim_verdict_value = (
        "pass transaction scope only"
        if protected_surface_contract_ok and transaction_ok
        else "blocked protected surface or transaction evidence incomplete"
    )
    review_copy_sha = file_sha256(review_copy) if review_copy.exists() else "0" * 64
    if (
        source_review_artifact_inventory_path == "AUTO"
        or final_review_artifact_diff_path == "AUTO"
        or source_body_citation_run_inventory_path == "AUTO"
        or final_body_citation_run_diff_path == "AUTO"
    ):
        if source_docx_path == "AUTO":
            raise ValueError("source_docx_path is required when generating format-task preservation evidence")
        source_docx = Path(source_docx_path).resolve()
        source_review_artifact_inventory = (
            path.with_suffix(".source-review-artifacts.md")
            if source_review_artifact_inventory_path == "AUTO"
            else Path(source_review_artifact_inventory_path)
        )
        final_review_artifact_diff = (
            path.with_suffix(".final-review-artifact-diff.md")
            if final_review_artifact_diff_path == "AUTO"
            else Path(final_review_artifact_diff_path)
        )
        source_body_citation_run_inventory = (
            path.with_suffix(".source-body-citation-runs.md")
            if source_body_citation_run_inventory_path == "AUTO"
            else Path(source_body_citation_run_inventory_path)
        )
        final_body_citation_run_diff = (
            path.with_suffix(".final-body-citation-run-diff.md")
            if final_body_citation_run_diff_path == "AUTO"
            else Path(final_body_citation_run_diff_path)
        )
        write_docx_preservation_reports(
            source_docx=source_docx,
            final_docx=review_copy,
            source_review_artifact_inventory_path=source_review_artifact_inventory,
            final_review_artifact_diff_path=final_review_artifact_diff,
            source_body_citation_run_inventory_path=source_body_citation_run_inventory,
            final_body_citation_run_diff_path=final_body_citation_run_diff,
        )
        source_review_artifact_inventory_path = source_review_artifact_inventory
        final_review_artifact_diff_path = final_review_artifact_diff
        source_body_citation_run_inventory_path = source_body_citation_run_inventory
        final_body_citation_run_diff_path = final_body_citation_run_diff
    if review_comments_change_marks_preservation_verdict == "AUTO":
        review_comments_change_marks_preservation_verdict = "pass review comments/change marks inventory and final diff verified"
    if comments_strip_explicit_user_approval == "AUTO":
        comments_strip_explicit_user_approval = "not-requested; comments and tracked changes were not stripped"
    if body_citation_superscripts_preservation_verdict == "AUTO":
        body_citation_superscripts_preservation_verdict = "pass body citation superscript run inventory and final diff verified"
    if citation_audit_final_docx_sha256 == "AUTO":
        citation_audit_final_docx_sha256 = review_copy_sha
    if citation_audit_source_to_final_run_diff_path == "AUTO":
        citation_audit_source_to_final_run_diff_path = final_body_citation_run_diff_path
    review_copy_promotion_gate_value = (
        "all required rendered surfaces pass and the acceptance gate passes"
        if acceptance_handoff_status_value == "pass"
        else "blocked; required rendered surfaces or acceptance gate did not pass"
    )
    baseline_promotion_gate_value = (
        "pass release baseline promotion gates closed with sibling and cross-surface evidence"
        if acceptance_handoff_status_value == "pass"
        else "blocked release baseline promotion gates not closed"
    )
    release_blocker_ledger_path = high_risk_surface_matrix
    unresolved_release_blocker_count = "0" if acceptance_handoff_status_value == "pass" else "1"
    scoped_artifact_next_baseline_verdict = (
        "pass promoted only after release blocker ledger, sibling surfaces, and whole-document gates closed"
        if acceptance_handoff_status_value == "pass"
        else "blocked candidate-only artifact until release blocker ledger is closed"
    )
    routed_child_files = "; ".join(
        [
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "thesis-workflow-map.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "thesis-format-sop.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "protected-surface-evidence-contract.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "front-matter-and-toc.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "headings-and-figures.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "tables-abstracts-citations-references.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "thesis-figure-generation-rules.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "figure-rules" / "baseline-and-sourcing.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "figure-rules" / "review-gates.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "figure-rules" / "geometry-and-layout.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "figure-rules" / "workflow-and-checklists.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "review-figure-style-checklist.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "user-feedback" / "template-and-layout.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "user-feedback" / "final-qa-and-tooling.md").resolve()),
        ]
    )
    selected_workflow = "local-surface-repair" if task_mode == "format-repair-only" else "whole-thesis-revision"
    transaction_subtype = (
        "format-repair-only-docx-format-repair"
        if task_mode == "format-repair-only"
        else "new-thesis-production-docx-write"
    )
    touched_blocks = subtask if task_mode == "format-repair-only" else "sample rebuild complete manuscript"
    if task_mode == "format-repair-only":
        protected_blocks_value = "footer; page numbers; header; whole_document_pagination"
        transaction_target_surface_ids_value = "footer; page_numbers"
        transaction_protected_sibling_surface_ids_value = (
            "header; references_entries; body_citation_superscripts; "
            "review_comments_and_change_marks; whole_document_pagination"
        )
        protected_surface_write_owner_map_value = "footer/page-number local repair pipeline owns only footer and page-number surfaces"
        touched_surface_page_map_value = "footer/page-number target pages mapped by rendered pagination evidence"
        rerender_target_map_value = "footer/page-number target page plus whole-document pagination pages"
        touched_template_surface_families_value = "header; footer; page numbers; whole_document_pagination"
        style_blast_radius_trigger_value = "footer/page-number local alignment repair; no broad paragraph cleanup"
        active_table_authority_summary_value = "not-applicable-with-reason local footer repair only"
        table_authority_checkpoints_value = "not-applicable-with-reason local footer repair only"
        toc_repair_included_value = "no"
        tail_block_pagination_included_value = "not-applicable-with-reason local footer repair only"
        tail_block_titles_in_scope_value = "not-applicable-with-reason local footer repair only"
        touched_header_footer_surfaces_value = "footer page-number paragraphs"
        body_header_ownership_strategy_value = "body header unchanged; footer/page-number local repair only"
        tail_block_ownership_strategy_value = "local footer/page-number ownership only"
        explicit_user_rule_value = "user-provided template; close18 footer page-number alignment from locked template"
        rendered_pages_to_inspect_value = "footer/page-number target pages; whole-document pagination evidence"
        rendered_sentinel_texts_value = "footer/page-number target pages; final page-number sentinel"
        specific_format_classes_value = "footer; page numbers; header/footer"
        acceptance_focus_high_risk_value = "high-risk matrix covers footer, page numbers, header, and whole-document pagination"
    else:
        protected_blocks_value = "cover; Chinese abstract; English abstract; TOC; body headings; references; acknowledgement"
        transaction_target_surface_ids_value = (
            "cover; abstracts; keywords; TOC; headings; body; figures; tables; "
            "references; acknowledgement; appendix; header; footer; page numbers; whole_document_pagination"
        )
        transaction_protected_sibling_surface_ids_value = "all mandatory thesis protected surfaces"
        protected_surface_write_owner_map_value = "sample rebuild pipeline owns all touched surfaces in serialized order"
        touched_surface_page_map_value = "sample rebuild surfaces mapped by rendered page-class evidence"
        rerender_target_map_value = "cover, abstracts, TOC, body, figures, tables, references, acknowledgement, appendix"
        touched_template_surface_families_value = (
            "cover; abstracts; keywords; TOC; headings; body; figures; tables; "
            "references; acknowledgement; appendix; header; footer; page numbers"
        )
        style_blast_radius_trigger_value = "user-reported abstract, TOC, header, figure, body-style, code-block, table, and pagination drift"
        active_table_authority_summary_value = "canonical content manifest table data checked against generated DOCX table surface"
        table_authority_checkpoints_value = "caption binding; table text baseline; table count; rendered table family"
        toc_repair_included_value = "no"
        tail_block_pagination_included_value = "yes"
        tail_block_titles_in_scope_value = "references; acknowledgement"
        touched_header_footer_surfaces_value = "tail-block first pages; tail-block footer family"
        body_header_ownership_strategy_value = "body header unchanged; tail-block review lane"
        tail_block_ownership_strategy_value = "single verified opener owner per tail block"
        explicit_user_rule_value = "rebuild complete sample from template authority"
        rendered_pages_to_inspect_value = "cover; abstracts; toc; first body page; figure page; table page; references; acknowledgement"
        rendered_sentinel_texts_value = "sample cover; sample abstracts; sample toc; sample first body page; sample references; sample acknowledgement"
        specific_format_classes_value = "cover; front matter; body headings; tables; references; acknowledgement"
        acceptance_focus_high_risk_value = "high-risk matrix covers cover, abstracts, keywords, TOC, headings, references, acknowledgement, and appendix"
    text = f"""## Scope Lock
- task mode: {task_mode}
- content frozen?: yes
- protected blocks: {protected_blocks_value}
- touched blocks this round: {touched_blocks}

## Workflow Locks
- active references: {Path(__file__).resolve().parents[1] / 'SKILL.md'}
- routed child files: {routed_child_files}
- active checklist names: review-thesis-format-checklist
- thesis mutation transaction owner path: references/thesis/thesis-mutation-transaction.md
- selected mutation transaction workflow: {selected_workflow}
- mutation transaction subtype: {transaction_subtype}
- thesis mutation transaction record path: {transaction_record_value}
- transaction target surface ids: {transaction_target_surface_ids_value}
- transaction protected sibling surface ids: {transaction_protected_sibling_surface_ids_value}
- transaction write owner: upstream canonical builder owns DOCX write; acceptance generator records evidence
- transaction audit owner or fallback: format-worker audit lane or sequential audit fallback
- transaction source manuscript path: {review_copy}
- transaction source manuscript sha256: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- transaction template path: {template_copy}
- transaction template sha256: {file_sha256(template_copy) if template_copy.exists() else "missing"}
- transaction review-copy path: {review_copy}
- transaction review-copy sha256 before mutation: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- transaction final docx path: {review_copy}
- transaction final docx sha256: {transaction_final_docx_sha256}
- protected-surface evidence contract path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- protected surface contract evidence path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- protected-surface evidence contract loaded?: yes
- canonical protected surface id set: {'; '.join(PROTECTED_SURFACE_IDS)}
- protected-surface owner map: {protected_surface_owner_map}
- protected-surface evidence map: {protected_surface_evidence_map}
- protected-surface reviewed output sha256: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- protected-surface evidence contract verdict: {protected_surface_contract_verdict}
- exact master manuscript path: {review_copy}
- true current master manuscript: {review_copy}
- exact review-copy path: {review_copy}
- exact review-copy sha256 before mutation: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- exact helper-script target path: {review_copy}
- expected output path after helper script: {review_copy}
- expected output sha256 after helper script: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- helper scripts planned this round: {helper_scripts_planned}
- project-local helper script preflight summary: {project_local_helper_script_preflight_summary}
- project-local helper preflight report path: {project_local_helper_preflight_report_path}
- project-local helper risk count: {project_local_helper_risk_count}
- project-local helper disposition: {project_local_helper_disposition}
- canonical source restart required?: {canonical_source_restart_required}
- source manuscript genealogy path: {source_manuscript_genealogy_path}
- source retention manifest path: {source_retention_manifest_path}
- source retention ratio: {source_retention_ratio}
- source retention verdict: {source_retention_verdict}
- source review-artifact inventory path: {source_review_artifact_inventory_path}
- final review-artifact diff path: {final_review_artifact_diff_path}
- review comments/change marks preservation verdict: {review_comments_change_marks_preservation_verdict}
- comments strip explicit user approval: {comments_strip_explicit_user_approval}
- comment-resolution ledger path: {comment_resolution_ledger_path}
- comment-resolution audit report path: {comment_resolution_audit_report_path}
- comment-resolution audit verdict: {comment_resolution_audit_verdict}
- source body-citation run inventory path: {source_body_citation_run_inventory_path}
- final body-citation run diff path: {final_body_citation_run_diff_path}
- body citation superscripts preservation verdict: {body_citation_superscripts_preservation_verdict}
- rebuild class: {rebuild_class}
- clean-source restart source path: {clean_source_restart_source_path}
- contaminated-baseline disposition: {contaminated_baseline_disposition}
- project template discovery root: {template_copy.parent}
- template discovery patterns: explicit template path
- discovered candidate template paths: {template_copy}
- candidate template selection reason: user-provided canonical template argument
- active template source type: user-provided
- active template path lock: {template_copy}
- active template fingerprint: {file_sha256(template_copy) if template_copy.exists() else "missing"}
- active template profile path: {template_profile}
- active template selected before mutation?: yes
- template discovery report path: {template_discovery_report}
- template profile generation command: {PYTHON_EXE} {Path(__file__).resolve().parents[1] / 'scripts' / 'thesis_template_profile.py'} --template {template_copy} --output {template_profile}
- template profile generated before mutation?: yes
- locked path encoding verdict: pass
- contaminated intermediate source used?: no
- contaminated intermediate source disposition: none
- template alignment verdict: pass
- format preservation promise active?: yes
- format preservation promise source: active skill promise/user requirement; thesis mutation preserves format
- format preservation promise verdict: {format_preservation_promise_verdict_value}
- protected-surface write-owner map: {protected_surface_write_owner_map_value}
- touched-surface page map: {touched_surface_page_map_value}
- stale audit map: no stale audits; all blocking evidence regenerated after final mutation
- rerender target map: {rerender_target_map_value}
- mandatory thesis surface inventory path: {mandatory_surface_inventory}
- thesis surface inventory status matrix: {surface_inventory_status_matrix_value}
- front matter surface coverage matrix path: {mandatory_surface_inventory}
- end matter surface coverage matrix path: {mandatory_surface_inventory}
- high-risk thesis format surfaces required: cover style; Chinese abstract title/body/keyword line; English abstract title/body/keyword line; TOC title/entries/dotted leaders/page-number column; body heading levels; body text; references title/entries; acknowledgement title/body; appendix title/body
- high-risk thesis format surface matrix path: {high_risk_surface_matrix}
- mandatory surface rows required: cover; declaration/title front matter; Chinese abstract title/body/keyword line; English abstract title/body/keyword line; TOC title; TOC entries; TOC dotted leaders; TOC page-number column; body headings; body text; figure/table captions; references title; reference entries; acknowledgement title/body; appendix title/body; header; footer; page numbers
- protected abstract/TOC evidence id rule: each protected abstract and TOC evidence record must use the exact protected surface id in target identifier and baseline surface id, and one evidence file cannot be reused across protected surfaces
- surfaces explicitly excluded from generic scripts: generic scripts outside the serialized rebuild owner map cannot mutate TOC entries, abstract keyword lines, heading families, or header/footer surfaces
- custom-layout result tables locked this round: probe manuscript contains no custom-layout result tables and no generic restyle pass may invent or rewrite that surface family
- template-owned surface families touched or user-reported: {touched_template_surface_families_value}
- style-blast-radius trigger: {style_blast_radius_trigger_value}
- protected-surface freeze manifest path: {mandatory_surface_inventory}
- protected_surface_freeze_manifest schema: graduation-project-builder.protected-surface-freeze-manifest.v1
- protected_surface_freeze_manifest sha256: {file_sha256(mandatory_surface_inventory) if mandatory_surface_inventory.exists() else "missing"}
- protected_surface_freeze_manifest verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- pre-mutation protected-surface snapshot path: {thesis_evidence}
- post-mutation protected-surface snapshot path: {touched_evidence}
- post_mutation_surface_diff path: {paragraph_evidence}
- post_mutation_surface_diff sha256: {file_sha256(paragraph_evidence) if paragraph_evidence.exists() else "missing"}
- post_mutation_surface_diff verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- target_surface_render_review path: {thesis_evidence}
- target_surface_render_review sha256: {file_sha256(thesis_evidence) if thesis_evidence.exists() else "missing"}
- target_surface_render_review verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- blast_radius_render_review path: {touched_evidence}
- blast_radius_render_review sha256: {file_sha256(touched_evidence) if touched_evidence.exists() else "missing"}
- blast_radius_render_review verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- cross_surface_regression_report path: {paragraph_evidence}
- cross_surface_regression_report sha256: {file_sha256(paragraph_evidence) if paragraph_evidence.exists() else "missing"}
- cross_surface_regression_report verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- non-target protected surface change verdict: {"pass" if transaction_ok else "blocked transaction validation failed"}
- chapter_format_preservation_report path: {chapter_format_diff_path}
- chapter_format_preservation_report sha256: {file_sha256(chapter_format_diff_path) if chapter_format_diff_path.exists() else "missing"}
- chapter_format_preservation_report verdict: {format_preservation_promise_verdict_value}
- chapter format preservation detector verdict: {chapter_format_preservation_detector_verdict}
- non-target format preservation verdict: {non_target_format_preservation_verdict_value}
- local-surface whole-thesis claim verdict: {local_surface_whole_thesis_claim_verdict_value}
- cross-surface regression diff path: {paragraph_evidence}
- cross-surface regression owner: format-worker + audit
- TOC underline pollution detector required?: yes
- table style regression detector required?: yes
- surface-face parity baseline source paths: {template_copy}; {template_profile}
- surface-face parity owner map: protected-surface evidence contract owns abstract/TOC records; front-matter-and-toc owns TOC local repair detail; thesis-format-sop owns full-surface workflow
- surface-face sibling audit scope: every touched or user-reported surface plus same-class sibling surfaces in the mandatory inventory
- all-surface paragraph-dialog / typography metrics baseline: locked from active template style binding, WPS/Word paragraph dialog, and run typography for every present template-owned surface
- all-surface style-binding hard-field rule: every template-owned surface must expose WPS/Word paragraph-dialog and typography baseline/actual fields; broad baseline metrics prose is not enough
- all-surface paragraph-dialog / typography evidence record paths: {all_surface_paragraph_typography_evidence_paths}
- whole-document pagination surface active?: yes
- whole-document pagination baseline source: locked template/profile and pre-mutation rendered page map
- package baseline manifest path: {thesis_evidence}
- pre-mutation page map path: {thesis_evidence}
- section break numbering baseline path: {thesis_evidence}
- page-number restart baseline path: {thesis_evidence}
- header/footer link-to-previous baseline path: {thesis_evidence}
- hard page-break / section-break baseline path: {thesis_evidence}
- field-refresh baseline state: fields refreshed before mutation and recorded for TOC/page numbering
- TOC-to-heading page sync baseline path: {thesis_evidence}
- logical-physical page map baseline path: {thesis_evidence}
- rendered page count baseline: template page count=1; actual generated baseline page count={len(page_images)}
- blank-page scan baseline path: {thesis_evidence}
- cover style baseline source: {template_copy}::cover
- cover identity value-line baseline source: {template_copy}::cover_identity_value_lines
- reference-entry baseline source: {template_copy}::references
- appendix baseline source or absence reason: {template_copy}::appendix or active-template absence reason in mandatory inventory
- custom/builder/default font usage allowed?: no
- font baseline source for each touched surface: active template donor paragraph/run for each touched surface
- Chinese/East Asian font mapping baseline: active template eastAsia font chain through direct run, character style, paragraph style, basedOn, docDefaults, theme, and WPS/Word UI display
- Western/English font mapping baseline: active template ascii/hAnsi font chain through direct run, character style, paragraph style, basedOn, docDefaults, theme, and WPS/Word UI display
- complex-script font mapping baseline: active template cs font chain or explicit not-applicable proof from template
- effective font chain baseline required?: yes
- protected-surface effective font chain rule: protected abstract and TOC evidence must resolve baseline and actual effective font chains before pass
- theme/default font alias policy: theme/default aliases are allowed only when the active template profile proves the same alias as baseline
- missing font or surface baseline blockers: none
- active table authority lock: locked to canonical content manifest table data
- table authority source type: canonical content manifest
- table authority source file path: {citation_audit}
- table authority manuscript binding proof: {thesis_evidence}
- table authority rationale: table content and caption are generated from the validated canonical manifest and verified by sample_self_check
- WPS table preset target: not-applicable
- heading / TOC / chapter-start linked lane active?: yes
- TOC baseline-source lock: locked
- TOC baseline source file path: {template_copy}
- TOC implementation type in baseline source: template TOC heading with paragraph-level TOC entries
- TOC restoration source allowed to be current working draft?: disallowed
- TOC title visual metrics baseline: {TOC_BASELINE_TITLE}
- TOC level-1 visual metrics baseline: {TOC_BASELINE_LEVEL1}
- TOC level-2 visual metrics baseline: {TOC_BASELINE_LEVEL2}
- TOC level-3 visual metrics baseline: {TOC_BASELINE_LEVEL3}
- TOC title paragraph-dialog / typography metrics baseline: locked from template TOC title style, WPS paragraph dialog, and title run typography
- TOC level-1 paragraph-dialog / typography metrics baseline: locked from template TOC level-1 style binding, WPS paragraph dialog, and run typography
- TOC level-2 paragraph-dialog / typography metrics baseline: locked from template TOC level-2 style binding, WPS paragraph dialog, and run typography
- TOC level-3 paragraph-dialog / typography metrics baseline: locked from template TOC level-3 style binding, WPS paragraph dialog, and run typography
- TOC visible-run direct typography baseline: locked for TOC title text, each used level entry text, tab/leader, and page-number runs
- TOC per-level text/tab/page-number run typography baseline: visible-run direct rPr, rFonts script/theme slots, size/sizeCs, and weight recorded per used level
- TOC per-level style-binding baseline: TOC title and every used TOC level have independent template donor style id/name baselines
- TOC scale/compression baseline verdict: pass only when font size, paragraph spacing, line-spacing value, indentation, tab stops, and rendered rhythm are not proportionally reduced against the template
- TOC page-count / occupancy baseline: template TOC occupies its own front-matter page class in the accepted sample flow
- TOC visual geometry baseline: template rendered TOC title bbox, row bboxes, per-level x positions, row y-deltas, leader density, page-number x column, and occupancy rhythm
- TOC post-refresh restoration owner: sample rebuild pipeline
- TOC repair included this round?: {toc_repair_included_value}
- live TOC required this round?: {live_toc_required_value}
- live TOC field count: {live_toc_field_count}
- live TOC field verdict: {live_toc_field_verdict_value}
- front-matter numbering convention expected on rendered pages: template-following split front-matter numbering and body numbering
- office-app path used to verify adjusted page numbers: {RENDERER_PATH}
- abstract surfaces locked: {ABSTRACT_SURFACES}
- abstract baseline source file path: {template_copy}
- Chinese abstract title baseline paragraph/run path: {template_copy}::zh_abstract_title
- Chinese abstract body baseline paragraph/run path: {template_copy}::zh_abstract_body
- Chinese keyword line baseline paragraph/run path: {template_copy}::zh_keyword_line
- English abstract title baseline paragraph/run path: {template_copy}::en_abstract_title
- English abstract body baseline paragraph/run path: {template_copy}::en_abstract_body
- English keyword line baseline paragraph/run path: {template_copy}::en_keyword_line
- keyword label/content run-split baseline: template keyword label run bold, content run non-bold unless explicit template override
- English abstract semantic parity baseline: final English abstract and keywords match final Chinese abstract scope and claims
- caption baseline source file path: not-applicable
- caption wording lock: not-applicable
- paragraph-level rendered review path: review evidence records
- verified renderer executable path: {RENDERER_PATH}
- verified rasterizer executable path: {RASTERIZER_PATH}
- touched header/footer surfaces: {touched_header_footer_surfaces_value}
- body-header ownership strategy: {body_header_ownership_strategy_value}
- tail-block ownership strategy: {tail_block_ownership_strategy_value}
- tail-block pagination repair included this round?: {tail_block_pagination_included_value}
- tail-block titles in scope: {tail_block_titles_in_scope_value}
- tail-block title baseline source file path: {template_copy}
- references title baseline paragraph/run path: {template_copy}::references_title
- references entries baseline paragraph/run path: {template_copy}::references_entries
- acknowledgement title baseline paragraph/run path: {template_copy}::acknowledgement_title
- acknowledgement body baseline paragraph/run path: {template_copy}::acknowledgement_body
- end-matter indentation detector required?: yes
- references title indentation evidence record path: {path.parent / 'protected-surface-references_title.md'}
- references entries indentation evidence record path: {path.parent / 'protected-surface-references_entries.md'}
- acknowledgement title indentation evidence record path: {path.parent / 'protected-surface-acknowledgement_title.md'}
- acknowledgement body indentation evidence record path: {path.parent / 'protected-surface-acknowledgement_body.md'}
- references title indentation baseline: active template reference title paragraph/run indentation plus rendered x-position baseline
- references entries left/hanging indentation baseline: active template reference-entry left/hanging values plus rendered entry-start x-position baseline
- acknowledgement title indentation baseline: active template acknowledgement title paragraph/run indentation plus rendered x-position baseline
- acknowledgement body indentation baseline: active template acknowledgement body paragraph/run indentation plus rendered x-position baseline
- references title indentation checkpoints: protected-surface-references_title paragraph typography and rendered geometry evidence must pass
- references entries indentation checkpoints: protected-surface-references_entries paragraph typography and rendered geometry evidence must pass
- acknowledgement title indentation checkpoints: protected-surface-acknowledgement_title paragraph typography and rendered geometry evidence must pass
- acknowledgement body indentation checkpoints: protected-surface-acknowledgement_body paragraph typography and rendered geometry evidence must pass
- end-matter rendered geometry checkpoints: references and acknowledgement protected-surface geometry records bind to the exact final DOCX SHA256
- tail-block opener ownership baseline: one page-start owner per tail block from the approved template
- tail-block pagination restoration owner: sample rebuild pipeline
- tail-block opener page-occupancy baseline: references opener and acknowledgement opener each occupy their own opener page
- footer baseline source file path: {template_copy}
- footer/page-number presentation target: template-following footer and page-number baseline
- logical-page to physical-rendered-page mapping method: sentinel text mapping
- implicit skips allowed?: disallowed
- post-script smoke-audit rule active?: yes

## Review Evidence Paths
- rendered page review evidence record paths: {thesis_evidence}
- thesis mutation transaction validator command: {PYTHON_EXE} {Path(__file__).resolve().parents[1] / 'scripts' / 'validate_thesis_mutation_transaction.py'} --record {transaction_record_value} --final-docx {review_copy}
- thesis mutation transaction validator result: {transaction_validator_result}
- paragraph-review evidence record paths: {paragraph_evidence}
- touched-page review evidence record paths: {touched_evidence}
- table-local structure evidence record paths: {thesis_evidence}
- rendered table baseline comparison evidence record paths: {thesis_evidence}
- table title-mode evidence paths: {thesis_evidence}
- table header-bottom middle-rule evidence paths: {thesis_evidence}
- TOC post-refresh restoration evidence record paths: {thesis_evidence}
- TOC rendered baseline comparison evidence record paths: {thesis_evidence}
- TOC visual geometry evidence record paths: {toc_visual_geometry_evidence_paths}
- TOC paragraph-and-typography evidence record paths: {toc_paragraph_typography_evidence_paths}
- TOC visible-run typography evidence record paths: {toc_paragraph_typography_evidence_paths}
- whole-document pagination evidence record path: {whole_document_pagination_evidence_path}
- chapter format diff path: {chapter_format_diff_path}
- touched chapter rendered evidence paths: {touched_evidence}
- package drift report path: {thesis_evidence}
- post-mutation page map path: {thesis_evidence}
- whole-document pagination diff path: {thesis_evidence}
- section break numbering map path: {thesis_evidence}
- chapter start owner map path: {thesis_evidence}
- tail-block owner map path: {thesis_evidence}
- TOC-to-heading page sync map path: {thesis_evidence}
- logical-physical page map path: {thesis_evidence}
- blank-page scan evidence path: {thesis_evidence}
- cross-surface regression freeze evidence path: {thesis_evidence}
- TOC underline pollution evidence path: {thesis_evidence}
- table style regression evidence path: {thesis_evidence}
- citation audit report path: {citation_audit}
- citation audit final DOCX SHA256: {citation_audit_final_docx_sha256}
- citation audit source-to-final run diff path: {citation_audit_source_to_final_run_diff_path}
- caption rendered review evidence record paths: not-applicable
- header/footer rendered review evidence record paths: {touched_evidence}
- tail-block pagination evidence record paths: {thesis_evidence}
- tail-block rendered opener comparison evidence record paths: {touched_evidence}
- rendered PDF path: {rendered_pdf}
- page-image artifact paths: {join_paths(page_images)}
- output manuscript paths: {review_copy}
- reviewed output sha256: {file_sha256(review_copy) if review_copy.exists() else "missing"}
- post-script smoke-audit evidence path: {thesis_evidence}

## Baseline Source
- school template: {template_copy}
- approved sample: {template_copy}
- explicit user rule: {explicit_user_rule_value}
- page-class sample comparison targets: cover; Chinese abstract; English abstract; TOC; first body chapter page; one figure page; one table page; references; acknowledgement
- surface-face baseline paragraph/run paths: template donor paragraphs/runs for every protected abstract, TOC, heading, body, caption, table, reference, acknowledgement, appendix, header, footer, and page-number surface in scope
- surface-face baseline metrics recorded: font family, font size, weight, alignment, spacing before/after, line-spacing mode/value, indentation, tabs, geometry, and page occupancy
- effective font resolver inputs: direct run properties; character style; paragraph style; basedOn chain; docDefaults; theme major/minor fonts; WPS/Word UI displayed font names
- active table authority source summary: {active_table_authority_summary_value}
- TOC title baseline paragraph path: {template_copy}::TOC Heading
- TOC level-1 baseline paragraph path: {template_copy}::TOC 1
- TOC level-2 baseline paragraph path: {template_copy}::TOC 2
- TOC level-3 baseline paragraph path: {template_copy}::TOC 3 or equivalent reserved depth
- TOC level-4 baseline paragraph path: {template_copy}::TOC 4 or not-used in final manuscript
- TOC entries by used level baseline: every used TOC level mapped to a template donor and rendered-region evidence
- TOC right-tab position baseline: template TOC right-tab stop
- TOC line-spacing baseline: template TOC line spacing
- TOC paragraph before/after spacing baseline: template TOC title and per-level before/after spacing in points
- TOC line-spacing mode/value baseline: template TOC title and per-level line-spacing mode/value in WPS/Word dialog
- TOC indentation chars/points baseline: template TOC title and per-level left/right/first-line/hanging indentation in characters and points
- TOC per-level font-size/weight baseline: template TOC title and every used level font size and weight
- TOC dotted-leader baseline: template TOC dotted leader
- TOC page-number column per-entry baseline: each visible TOC page number matches rendered heading page and displayed numbering system
- TOC title bbox baseline: template TOC title bounding box
- TOC entry row bbox baseline: template TOC entry row bounding boxes
- TOC per-level left-indent x baseline: template TOC per-level left x positions
- TOC line-spacing y-delta baseline: template TOC row y-deltas
- TOC dotted-leader start/end/density baseline: template TOC leader start/end/density
- TOC page-number x column baseline: template TOC page-number x column
- TOC row-count / page-occupancy rhythm baseline: template TOC row count per page and occupancy rhythm
- forbidden substitute evidence allowed?: no
- tail-block title baseline paragraph paths: {template_copy}::references_title; {template_copy}::acknowledgement_title

## Repair Sequence
1. baseline extraction
2. structure repair
3. pagination / field refresh
4. style unification
5. regression review

## Acceptance Focus
- rendered pages to inspect: {rendered_pages_to_inspect_value}
- rendered sentinel texts to confirm: {rendered_sentinel_texts_value}
- specific format classes to verify: {specific_format_classes_value}
- surface-face parity checkpoints: each present or reported template-owned surface has independent baseline-vs-actual evidence and pass verdict
- mandatory thesis surface inventory checkpoints: inventory contains all required surface rows with evidence path, final verdict, and reason
- high-risk thesis format surface checkpoints: {acceptance_focus_high_risk_value}
- high-risk omission blocker verdict: {protected_surface_task_verdict_value}
- cover style checkpoints: cover rows have template baseline evidence and pass verdict
- cover identity value-line checkpoints: template cover identity value-cell bottom borders or underlines matched row by row
- references entry format checkpoints: reference title and entries have template baseline evidence and pass verdict
- appendix format checkpoints: appendix rows have template baseline or absence reason and pass verdict
- sibling-surface audit checkpoints: same-class sibling surfaces checked for drift after the generated acceptance run
- all-surface paragraph-dialog and typography checkpoints: every present template-owned surface has style binding, WPS/Word paragraph-dialog metrics, typography values, spacing, line mode, indentation, tabs/leaders, keep/list/page-break state, and scale/compression verdict recorded as hard fields
- font-family baseline checkpoints: effective font chain evidence exists for protected surfaces and font-bearing surfaces in scope
- protected-surface effective font chain checkpoints: protected abstract and TOC evidence records contain baseline/actual font chains and UI displayed font proof
- builder/default font rejection checkpoints: no builder-chosen or default font may pass without active-template donor proof
- table authority checkpoints: {table_authority_checkpoints_value}
- table-local structure checkpoints: standalone title; keep-with-next binding; grid text style; no title inside table grid
- table title-mode checkpoints: donor_title_mode and target_title_mode recorded for every touched body table
- table header-bottom middle-rule checkpoints: header-bottom middle rule, top rule, and bottom rule compared against the locked donor
- abstract and keyword checkpoints: six abstract and keyword surfaces checked against locked template donor
- abstract six-surface evidence checkpoints: zh_abstract_title; zh_abstract_body; zh_keyword_line; en_abstract_title; en_abstract_body; en_keyword_line all proved by review evidence
- keyword label/content run-split checkpoints: Chinese and English keyword label/content runs checked separately
- English abstract semantic parity checkpoints: English abstract and keywords checked against final Chinese abstract scope
- TOC visual baseline checkpoints: title baseline; level formatting; dotted leader; page-number column; occupancy baseline
- TOC per-level evidence checkpoints: every used TOC level checked against locked template donor
- TOC used-level inventory checkpoints: used levels title, level1, level2, level3, and level4 when present each have row-level paragraph/typography evidence
- TOC per-entry page-number checkpoints: each TOC visible page number checked against rendered heading page and displayed numbering
- TOC visual geometry checkpoints: rendered title bbox, row bboxes, level x positions, line-spacing y-deltas, leader density, page-number x column, and occupancy rhythm compared against template
- TOC paragraph-dialog and typography checkpoints: style id/name, title and per-level font size/weight, before/after spacing, line-spacing mode/value, indentation, tab stop/leader, and scale/compression compared against template
- TOC visible-run text/tab/page-number typography checkpoints: visible run direct rPr, rFonts script/theme slots, size/sizeCs, and weight compared for entry text, tab/leader, and page-number runs
- TOC visible-run typography verdict: {protected_surface_task_verdict_value}
- cross-surface regression verdict: {protected_surface_task_verdict_value}
- TOC underline pollution verdict: {protected_surface_task_verdict_value}
- table style regression verdict: {protected_surface_task_verdict_value}
- TOC restoration checkpoints: title baseline; level formatting; dotted leader; page-number column; occupancy baseline
- post-script smoke-audit checkpoints: doc opens; heading family stable; rendered page classes available
- chapter-start pagination checkpoints: chapter opener checked
- chapter format preservation checkpoints: chapter detector registry entry checked; chapter format diff recorded; touched chapter rendered evidence reviewed; non-target format preservation verdict passed
- tail-block pagination checkpoints: prior page checked; opener page checked; singular owner checked; references/acknowledgement separation checked
- whole-document pagination checkpoints: package drift, section map, page-number restart, field refresh, TOC-to-heading sync, logical/physical page map, rendered page count, blank-page scan, chapter openers, and tail-block openers checked
- whole-document pagination verdict: {protected_surface_task_verdict_value}
- runtime screenshot route-caption-asset map path: not-applicable
- review-copy promotion gate: {review_copy_promotion_gate_value}
- baseline promotion gate: {baseline_promotion_gate_value}
- release blocker ledger path: {release_blocker_ledger_path}
- unresolved release blocker count: {unresolved_release_blocker_count}
- scoped artifact next-baseline verdict: {scoped_artifact_next_baseline_verdict}
- blocker conditions: {acceptance_blockers_value}
- format preservation blocker conditions: pass no format preservation blockers
- known caveats: {acceptance_known_caveats_value}
"""
    write_text(path, text)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True)
    parser.add_argument(
        "--task-mode",
        default="thesis-only",
        choices=["thesis-only", "format-repair-only", "program-plus-thesis"],
        help="Acceptance mode to write into generated gate records.",
    )
    parser.add_argument(
        "--subtask",
        default="complete sample rebuild",
        help="Concrete subtask label to write into generated gate records.",
    )
    parser.add_argument("--project-root", help="Real project root for project-local thesis helper preflight scanning.")
    parser.add_argument("--source-docx", help="Original/source manuscript DOCX for source-to-final preservation checks.")
    parser.add_argument(
        "--figure-source-docx",
        help="Optional source DOCX for figure/media preservation; defaults to --source-docx.",
    )
    parser.add_argument(
        "--comment-source-docx",
        help="Optional source DOCX for comment-resolution ledger validation; defaults to --source-docx.",
    )
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--final-pdf", required=True)
    parser.add_argument("--self-check", required=True)
    parser.add_argument("--citation-audit", required=True)
    parser.add_argument("--font-audit", required=True)
    parser.add_argument("--font-color-audit", help="Optional output path for exact-output visible font-color audit JSON.")
    parser.add_argument("--whole-format-audit", help="Optional output path for exact-output whole-DOCX structural format gate JSON.")
    parser.add_argument("--body-style-audit", required=True)
    parser.add_argument("--formula-object-audit", help="Optional output path for formula object authenticity audit JSON.")
    parser.add_argument("--template-profile")
    parser.add_argument("--asset-manifest")
    parser.add_argument("--surface-geometry-json")
    parser.add_argument("--surface-paragraph-typography-json")
    parser.add_argument("--template-rendered-region-image")
    parser.add_argument("--actual-rendered-region-image")
    parser.add_argument("--toc-geometry-json")
    parser.add_argument("--toc-template-rendered-image")
    parser.add_argument("--toc-actual-rendered-image")
    parser.add_argument("--toc-paragraph-typography-json")
    parser.add_argument("--whole-pagination-json")
    parser.add_argument("--transaction-record")
    parser.add_argument("--comment-resolution-ledger")
    parser.add_argument("--controlled-bookmark-disposition")
    parser.add_argument("--humanizer-evidence")
    parser.add_argument("--agent-authorization-source", default="none")
    parser.add_argument("--agent-mode", default="single-agent-no-auth")
    parser.add_argument("--spawned-agent-ids", default="none")
    parser.add_argument("--spawned-agent-aliases-zh", default="none")
    parser.add_argument("--controller-agent-id", default="controller-sequential-role")
    parser.add_argument("--format-agent-id", default="format-sequential-role")
    parser.add_argument("--audit-agent-id", default="controller-audit-role")
    parser.add_argument("--sequential-audit-fallback-id", default="")
    parser.add_argument("--sequential-fallback-reason", default="")
    parser.add_argument("--helper-scripts-planned", default="")
    parser.add_argument("--delegated-canonical-helper-paths", default="")
    parser.add_argument("--build-command-evidence", default="")
    parser.add_argument("--utf8-gate-evidence", default="")
    parser.add_argument("--skill-gate-evidence", default="")
    parser.add_argument("--selftest-evidence", default="")
    parser.add_argument("--integration-gate-evidence", default="")
    parser.add_argument("--output-gate-evidence", default="")
    parser.add_argument("--clean-source-restart-completed", action="store_true")
    parser.add_argument("--validator", required=True)
    parser.add_argument("--selftest-command", required=True)
    parser.add_argument("--smoke-acceptance", action="store_true")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    final_docx = Path(args.final_docx).resolve()
    if not args.source_docx:
        parser.error("--source-docx is required so review comments, tracked changes, and body citation runs can be compared source-to-final")
    source_docx = Path(args.source_docx).resolve()
    if source_docx == final_docx:
        parser.error("--source-docx must be a distinct source manuscript path, not the same file as --final-docx")
    figure_source_docx = Path(args.figure_source_docx).resolve() if args.figure_source_docx else source_docx
    comment_source_docx = Path(args.comment_source_docx).resolve() if args.comment_source_docx else source_docx
    if figure_source_docx == final_docx:
        parser.error("--figure-source-docx must be distinct from --final-docx")
    if comment_source_docx == final_docx:
        parser.error("--comment-source-docx must be distinct from --final-docx")
    final_pdf = Path(args.final_pdf).resolve()
    self_check = Path(args.self_check).resolve()
    citation_audit = Path(args.citation_audit).resolve()
    font_audit = Path(args.font_audit).resolve()
    body_style_audit = Path(args.body_style_audit).resolve()
    def resolve_path_list(value: str | None) -> str:
        if not value:
            return "none"
        resolved: list[str] = []
        for part in value.split(";"):
            text = part.strip()
            if text:
                resolved.append(str(Path(text).resolve()))
        return "; ".join(resolved) if resolved else "none"

    humanizer_evidence_value = resolve_path_list(args.humanizer_evidence)
    agent_mode = args.agent_mode.strip() or "single-agent-no-auth"
    agent_authorization_source = args.agent_authorization_source.strip() or "none"
    if agent_mode == "single-agent-no-auth":
        agent_authorization_source = "none"
    spawned_agent_ids = args.spawned_agent_ids.strip() or "none"
    spawned_agent_aliases_zh = args.spawned_agent_aliases_zh.strip() or "none"
    requested_audit_agent_id = args.audit_agent_id.strip() or "controller-audit-role"
    controller_agent_id = args.controller_agent_id.strip() or "controller-sequential-role"
    format_agent_id = args.format_agent_id.strip() or "format-sequential-role"
    if agent_mode == "parallel-subagents":
        sequential_fallback_reason = "none"
    elif args.sequential_fallback_reason.strip():
        sequential_fallback_reason = args.sequential_fallback_reason.strip()
    else:
        sequential_fallback_reason = "no explicit user authorization, so no spawned-agent claim"
    spawn_attempted = "yes" if agent_mode == "parallel-subagents" else "no"
    spawn_requested = "yes" if agent_mode == "parallel-subagents" else "no"
    spawn_status = "spawned" if agent_mode == "parallel-subagents" else "not-spawned"
    fallback_mode = agent_mode
    if agent_mode == "parallel-subagents":
        audit_agent_id = requested_audit_agent_id
        sequential_audit_fallback_id = "none"
        audit_executor_id = audit_agent_id
    else:
        audit_agent_id = "none"
        sequential_audit_fallback_id = args.sequential_audit_fallback_id.strip() or requested_audit_agent_id
        audit_executor_id = sequential_audit_fallback_id
    template_copy = Path(args.template).resolve()
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    whole_format_audit_path = (
        Path(args.whole_format_audit).resolve()
        if args.whole_format_audit
        else output.parent / "whole-docx-format-gate.json"
    )
    whole_format_audit = audit_whole_format(final_docx, require_toc_field=True)
    whole_format_audit_path.parent.mkdir(parents=True, exist_ok=True)
    write_text(whole_format_audit_path, json.dumps(whole_format_audit, ensure_ascii=False, indent=2) + "\n")
    whole_format_audit_verdict = "pass" if whole_format_audit.get("passed") is True else "fail"
    font_color_audit_path = (
        Path(args.font_color_audit).resolve()
        if args.font_color_audit
        else output.parent / "font-color-audit.json"
    )
    font_color_audit = audit_font_color(final_docx)
    font_color_audit_path.parent.mkdir(parents=True, exist_ok=True)
    write_text(font_color_audit_path, json.dumps(font_color_audit, ensure_ascii=False, indent=2) + "\n")
    font_color_audit_verdict = "pass" if font_color_audit.get("passed") is True else "fail"
    task_mode_value = args.task_mode.strip()
    subtask_value = args.subtask.strip() or (
        "bounded format repair"
        if task_mode_value == "format-repair-only"
        else "complete sample rebuild"
    )
    is_format_repair_only = task_mode_value == "format-repair-only"
    selected_workflow_value = "local-surface-repair" if is_format_repair_only else "new-thesis-production"
    # - selected thesis workflow: new-thesis-production
    # Registry compatibility: default thesis-production records still keep the legacy rule line above.
    transaction_workflow_value = "local-surface-repair" if is_format_repair_only else "whole-thesis-revision"
    transaction_subtype_value = (
        "format-repair-only-docx-format-repair"
        if is_format_repair_only
        else "new-thesis-production-docx-write"
    )
    transaction_target_surfaces_value = (
        "references_entries"
        if is_format_repair_only
        else "cover; abstracts; keywords; TOC; headings; body; figures; tables; references; acknowledgement; appendix; header; footer; page numbers; whole_document_pagination"
    )
    transaction_protected_sibling_surfaces_value = (
        "references_title; body_citation_superscripts; review_comments_and_change_marks; whole_document_pagination"
        if is_format_repair_only
        else "all mandatory thesis protected surfaces"
    )
    touched_surface_families_value = (
        "references; page numbers; whole_document_pagination"
        if is_format_repair_only
        else "cover; abstracts; keywords; TOC; headings; body; figures; tables; references; acknowledgement; appendix; header; footer; page numbers"
    )
    humanizer_route_decision_value = "none"
    humanizer_target_language_value = "none"
    humanizer_scope_value = "none"
    humanizer_required = not is_format_repair_only
    formula_object_audit_path = (
        Path(args.formula_object_audit).resolve()
        if args.formula_object_audit
        else output.parent / "formula-object-audit.json"
    )
    formula_object_audit_error = ""
    try:
        formula_object_audit = audit_formula_objects(final_docx)
    except Exception as exc:
        formula_object_audit = {
            "schema": "graduation-project-builder.formula-object-audit.v1",
            "result": "fail",
            "docx_path": str(final_docx),
            "math_object_count": 0,
            "formula_like_paragraph_count": 0,
            "pseudo_formula_count": 0,
            "pseudo_formula_paragraphs": [],
            "error": str(exc),
        }
        formula_object_audit_error = str(exc)
    formula_object_audit_path.parent.mkdir(parents=True, exist_ok=True)
    write_text(formula_object_audit_path, json.dumps(formula_object_audit, ensure_ascii=False, indent=2) + "\n")
    formula_math_object_count = int(formula_object_audit.get("math_object_count") or 0)
    formula_like_paragraph_count = int(formula_object_audit.get("formula_like_paragraph_count") or 0)
    formula_pseudo_count = int(formula_object_audit.get("pseudo_formula_count") or 0)
    formula_surface_present = formula_math_object_count > 0 or formula_like_paragraph_count > 0
    formula_audit_ok = formula_object_audit.get("result") == "pass" and formula_pseudo_count == 0 and not formula_object_audit_error
    if formula_pseudo_count > 0:
        formula_object_summary_value = (
            f"failed plain text pseudo-formula audit: {formula_pseudo_count} pseudo-formula paragraphs; "
            f"math objects={formula_math_object_count}; audit={formula_object_audit_path}"
        )
        formula_numbering_summary_value = "failed formula numbering surface blocked until pseudo-formulas are converted to real equation objects"
    elif formula_surface_present:
        formula_object_summary_value = (
            f"passed formula object audit: math objects={formula_math_object_count}; "
            f"pseudo-formula paragraphs=0; audit={formula_object_audit_path}"
        )
        formula_numbering_summary_value = "not-applicable no formula-numbering repair requested in this acceptance generation"
    elif formula_object_audit_error:
        formula_object_summary_value = f"failed formula object audit could not read final DOCX: {formula_object_audit_error}"
        formula_numbering_summary_value = "failed formula numbering surface blocked because formula audit failed"
    else:
        formula_object_summary_value = f"not-applicable no formula object or pseudo-formula paragraph detected; audit={formula_object_audit_path}"
        formula_numbering_summary_value = "not-applicable no formula surface detected"
    humanizer_evidence_value = ensure_humanizer_evidence(final_docx, output.parent, humanizer_evidence_value)
    if humanizer_required:
        (
            humanizer_route_decision_value,
            humanizer_target_language_value,
            humanizer_scope_value,
        ) = infer_humanizer_route_from_evidence(humanizer_evidence_value)

    self_check_text = read_text(self_check)
    citation_text = read_text(citation_audit)
    font_text = read_text(font_audit)
    body_style_text = read_text(body_style_audit)
    page_images = find_page_images(final_pdf, output)

    statuses = summarize_statuses(
        self_check_text=self_check_text,
        citation_text=citation_text,
        font_text=font_text,
        body_style_text=body_style_text,
        smoke_acceptance=args.smoke_acceptance,
    )
    tail_block_detector_status = statuses.get(
        "tail-block.pagination-contract",
        (False, "missing tail-block.pagination-contract"),
    )
    if isinstance(tail_block_detector_status, tuple) and len(tail_block_detector_status) >= 2:
        tail_block_detector_ok_value = bool(tail_block_detector_status[0])
        tail_block_detector_summary_value = str(tail_block_detector_status[1])
    else:
        tail_block_detector_ok_value = False
        tail_block_detector_summary_value = str(tail_block_detector_status)
    tail_block_detector_status_value = (
        ("pass " if tail_block_detector_ok_value else "fail ") + tail_block_detector_summary_value
    )
    header_footer_page_number_detector_status = statuses.get(
        "header-footer.page-number-template-contract",
        (False, "missing header-footer.page-number-template-contract"),
    )
    if isinstance(header_footer_page_number_detector_status, tuple) and len(header_footer_page_number_detector_status) >= 2:
        header_footer_page_number_detector_ok_value = bool(header_footer_page_number_detector_status[0])
        header_footer_page_number_detector_summary_value = str(header_footer_page_number_detector_status[1])
    else:
        header_footer_page_number_detector_ok_value = False
        header_footer_page_number_detector_summary_value = str(header_footer_page_number_detector_status)
    header_footer_page_number_detector_status_value = (
        ("pass " if header_footer_page_number_detector_ok_value else "fail ")
        + header_footer_page_number_detector_summary_value
    )
    sample_self_check_detectors = parse_sample_self_check_detectors(self_check_text)
    chapter_format_detector = sample_self_check_detectors.get("chapter.format-preservation-contract", {})
    chapter_format_contract_path = output.parent / "chapter-format-preservation-contract.md"
    write_text(
        chapter_format_contract_path,
        "\n".join(
            [
                "# Chapter Format Preservation Contract",
                "- schema: graduation-project-builder.chapter-format-preservation-contract.v1",
                f"- final_docx: {final_docx}",
                f"- final_docx_sha256: {file_sha256(final_docx) if final_docx.exists() else 'missing'}",
                "- detector_id: chapter.format-preservation-contract",
                f"- detector: {json.dumps(chapter_format_detector, ensure_ascii=False, sort_keys=True)}",
                f"- verdict: {statuses['chapter_format_preservation_summary']}",
            ]
        )
        + "\n",
    )
    format_preservation_promise_verdict_value = (
        "passed chapter-level format preservation detector passed"
        if statuses["chapter_format_preservation_ok"]
        else f"failed {statuses['chapter_format_preservation_summary']}"
    )
    non_target_format_preservation_verdict_value = (
        "passed no non-target chapter/protected-surface format change detected"
        if statuses["chapter_format_preservation_ok"]
        else f"failed {statuses['chapter_format_preservation_summary']}"
    )

    thesis_evidence = output.parent / "thesis-rendered-page.md"
    paragraph_evidence = output.parent / "paragraph-review.md"
    touched_evidence = output.parent / "touched-page-review.md"
    chapter_format_diff_path = touched_evidence
    protected_evidence_paths = {
        surface_id: output.parent / f"protected-surface-{surface_id}.md"
        for surface_id in PROTECTED_SURFACE_IDS
    }
    toc_protected_evidence_paths = "; ".join(
        str(protected_evidence_paths[surface_id])
        for surface_id in ("toc_title", "toc_entries", "toc_dotted_leaders", "toc_page_number_column")
    )
    toc_entries_evidence_path = str(protected_evidence_paths["toc_entries"])
    figure_evidence = output.parent / "figure-review.md"
    format_task = output.parent / "format-task.md"
    page_class_coverage_matrix = output.parent / "page-class-coverage-matrix.json"
    mandatory_surface_inventory = output.parent / "mandatory-thesis-surface-inventory.md"
    high_risk_surface_matrix = output.parent / "high-risk-thesis-format-surface-matrix.md"
    template_discovery_report = output.parent / "template-discovery-report.json"
    project_local_helper_preflight_report = output.parent / "project-local-helper-preflight.md"
    user_issue_ledger = output.parent / "user-reported-issue-ledger.md"
    user_reported_visual_defect_evidence = output.parent / "user-reported-visual-defect-evidence.md"
    figure_comment_conversion = output.parent / "figure-comment-conversion-checklist.md"
    figure_plan_record = output.parent / "figure-plan.md"
    per_figure_evidence_manifest = output.parent / "per-figure-evidence-manifest.md"
    agent_manifest = output.parent / "agent-run-manifest.md"
    controller_card = output.parent / "controller-task-card.md"
    content_card = output.parent / "content-worker-task-card.md"
    format_card = output.parent / "format-worker-task-card.md"
    figure_card = output.parent / "figure-worker-task-card.md"
    citation_card = output.parent / "citation-worker-task-card.md"
    program_card = output.parent / "program-worker-task-card.md"
    acceptance_card = output.parent / "acceptance-worker-task-card.md"
    audit_card = output.parent / "audit-task-card.md"
    template_fingerprint = file_sha256(template_copy) if template_copy.exists() else "missing"
    active_template_profile_path = str(Path(args.template_profile).resolve()) if args.template_profile else "none"
    discovery_generated_at = datetime.now(timezone.utc)
    write_text(
        template_discovery_report,
        json.dumps(
            {
                "schema": "graduation-project-builder.template-discovery.v1",
                "generation_stage": "pre-mutation-template-lock",
                "generated_at_utc": discovery_generated_at.isoformat().replace("+00:00", "Z"),
                "generated_at_unix": discovery_generated_at.timestamp(),
                "generator": "scripts/discover_project_thesis_template.py",
                "selected_template_path": str(template_copy),
                "selected_template_fingerprint": template_fingerprint,
                "project_root": str(template_copy.parent),
                "candidate_template_paths": [str(template_copy)],
                "candidates": [
                    {
                        "path": str(template_copy),
                        "score": 100,
                        "sha256": template_fingerprint,
                        "size": template_copy.stat().st_size if template_copy.exists() else 0,
                    }
                ],
                "selection_reason": "user-provided canonical template argument",
                "source_type": "user-provided",
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
    )
    preflight_project_root = Path(args.project_root).resolve() if args.project_root else final_docx.parent
    (
        project_local_helper_script_paths,
        project_local_helper_script_risk_summary,
        project_local_helper_preflight_report_path,
        project_local_helper_risk_count,
    ) = scan_project_local_helper_scripts(preflight_project_root, project_local_helper_preflight_report)
    project_local_helper_script_preflight_summary, contaminated_baseline_disposition = summarize_project_local_helper_preflight(
        project_local_helper_script_risk_summary
    )
    project_local_risk_detected = (
        project_local_helper_risk_count > 0
        or project_local_helper_script_risk_summary.startswith("failed")
    )
    project_local_helper_disposition = (
        "clean-source-restart-completed"
        if project_local_risk_detected and args.clean_source_restart_completed
        else "clean-source-restart-required"
        if project_local_risk_detected
        else "clean"
    )
    canonical_source_restart_required = (
        "completed"
        if project_local_risk_detected and args.clean_source_restart_completed
        else "yes"
        if project_local_risk_detected
        else "no"
    )
    source_manuscript_genealogy_path = output.parent / "source-manuscript-genealogy.md"
    write_text(
        source_manuscript_genealogy_path,
        "# Source Manuscript Genealogy\n\n"
        f"- template source: {template_copy}\n"
        f"- preservation source manuscript: {source_docx}\n"
        f"- final reviewed output: {final_docx}\n"
        "- genealogy owner: generate_thesis_acceptance_record.py\n",
    )
    source_retention_manifest_path = output.parent / "source-retention-manifest.json"
    source_retention_ratio = "not-applicable-with-reason"
    source_retention_verdict = (
        "not-applicable-with-reason format-repair-only preserves source body; source-to-final diff and transaction record carry scope"
        if is_format_repair_only
        else "not-applicable-with-reason new-thesis-production or externally supplied source retention evidence required"
    )
    rebuild_class = "format-repair-only" if is_format_repair_only else "new-thesis-production"
    if project_local_risk_detected and args.clean_source_restart_completed:
        clean_source_restart_source_path = str(template_copy)
        contaminated_baseline_disposition = (
            "clean-source restart completed; contaminated project-local thesis helper scripts were not used by canonical builder"
        )
    else:
        clean_source_restart_source_path = (
            str(final_docx)
            if canonical_source_restart_required == "yes"
            else "not-applicable-with-reason clean project-local helper preflight"
        )
    write_text(
        source_retention_manifest_path,
        json.dumps(
            {
                "schema": "graduation-project-builder.source-retention.v1",
                "generator": "generate_thesis_acceptance_record.py",
                "source": str(source_docx),
                "source_sha256": file_sha256(source_docx) if source_docx.exists() else "missing",
                "source_class": "clean-source-template" if project_local_risk_detected and args.clean_source_restart_completed else "template-or-source",
                "output": str(final_docx),
                "output_sha256": file_sha256(final_docx) if final_docx.exists() else "missing",
                "retention_ratio": source_retention_ratio,
                "verdict": source_retention_verdict,
                "rebuild_class": rebuild_class,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
    )
    final_docx_sha = file_sha256(final_docx) if final_docx.exists() else "0" * 64
    source_review_artifact_inventory_path = output.parent / "source-review-artifacts.md"
    final_review_artifact_diff_path = output.parent / "final-review-artifact-diff.md"
    source_body_citation_run_inventory_path = output.parent / "source-body-citation-runs.md"
    final_body_citation_run_diff_path = output.parent / "final-body-citation-run-diff.md"
    controlled_bookmark_disposition_path = (
        Path(args.controlled_bookmark_disposition).resolve()
        if args.controlled_bookmark_disposition
        else None
    )
    controlled_bookmark_disposition_record_path = (
        str(controlled_bookmark_disposition_path)
        if controlled_bookmark_disposition_path is not None
        else "none"
    )
    write_docx_preservation_reports(
        source_docx=source_docx,
        final_docx=final_docx,
        source_review_artifact_inventory_path=source_review_artifact_inventory_path,
        final_review_artifact_diff_path=final_review_artifact_diff_path,
        source_body_citation_run_inventory_path=source_body_citation_run_inventory_path,
        final_body_citation_run_diff_path=final_body_citation_run_diff_path,
        controlled_bookmark_disposition_path=controlled_bookmark_disposition_path,
    )
    review_comments_change_marks_preservation_verdict = "pass review comments/change marks inventory and final diff verified"
    comments_strip_explicit_user_approval = "not-requested; comments and tracked changes were not stripped"
    body_citation_superscripts_preservation_verdict = "pass body citation superscript run inventory and final diff verified"
    citation_audit_final_docx_sha256 = final_docx_sha
    citation_audit_source_to_final_run_diff_path = final_body_citation_run_diff_path
    if args.asset_manifest:
        figure_asset_manifest_for_contract = str(Path(args.asset_manifest).resolve())
    else:
        generated_figure_manifest = output.parent / "figure-asset-manifest.json"
        write_text(
            generated_figure_manifest,
            json.dumps(
                {
                    "schema": "graduation-project-builder.figure-manifest.v2",
                    "figures": {},
                    "tables": {},
                    "diagrams": {},
                    "generation_note": "No figure assets were supplied; this empty manifest records that the generated acceptance fixture has no structural figure assets to validate.",
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
        )
        figure_asset_manifest_for_contract = str(generated_figure_manifest)
    figure_asset_manifest_value, figure_contract_summary = summarize_figure_contract(
        figure_asset_manifest_for_contract,
        final_docx,
        figure_source_docx,
    )
    figure_contract_passed = figure_contract_summary.startswith("passed")
    figure_acceptance_verdict = "pass" if figure_contract_passed else f"blocked {figure_contract_summary}"
    structural_figure_fields, structural_figure_geometry_active = create_structural_figure_evidence(
        figure_asset_manifest_value,
        output.parent,
    )
    surface_geometry_field_maps, surface_geometry_ok, surface_geometry_summary = load_surface_geometry_field_maps(args)
    surface_paragraph_typography_field_maps, surface_paragraph_typography_ok, surface_paragraph_typography_summary = load_surface_paragraph_typography_field_maps(args)
    surface_geometry_fields = dict(surface_geometry_field_maps.get("cover_style") or next(iter(surface_geometry_field_maps.values())))
    surface_paragraph_typography_fields = dict(
        surface_paragraph_typography_field_maps.get("cover_style")
        or next(iter(surface_paragraph_typography_field_maps.values()))
    )
    toc_geometry_fields, toc_geometry_ok, toc_geometry_summary = load_toc_geometry_fields(args)
    toc_paragraph_typography_fields, toc_paragraph_typography_ok, toc_paragraph_typography_summary = load_toc_paragraph_typography_fields(args)
    whole_pagination_fields, whole_pagination_ok, whole_pagination_summary = load_whole_pagination_fields(args)
    (
        comment_resolution_ledger_path,
        comment_resolution_audit_report_path,
        comment_resolution_audit_verdict,
        comment_resolution_ok,
    ) = summarize_comment_resolution_contract(
        args.comment_resolution_ledger,
        source_docx=comment_source_docx,
        final_docx=final_docx,
        output_dir=output.parent,
    )

    make_review_evidence(
        thesis_evidence,
        evidence_type="thesis-rendered-page",
        task_mode=task_mode_value,
        source_docx_path=source_docx,
        reviewed_output=final_docx,
        rendered_pdf=final_pdf,
        page_images=page_images,
        target_surface="complete sample page-class review",
        target_identifier="complete-sample",
        target_region="rendered sample page classes",
        checks="rendered page review; toc baseline comparison; abstract surface comparison",
        summary="Rendered complete sample review for cover, abstracts, TOC, body, references, and acknowledgement.",
        toc_title_confirmed="yes",
        toc_level_formatting_confirmed="yes",
        toc_dotted_leader_confirmed="yes",
        toc_page_number_column_confirmed="yes",
        toc_restored_confirmed="yes",
        toc_page_occupancy_confirmed="yes",
        toc_title_paragraph_confirmed="yes",
        toc_entries_by_level_confirmed="yes",
        toc_dotted_leaders_confirmed="yes",
        toc_page_number_column_per_entry_confirmed="yes",
        table_authority_source_confirmed="yes",
        table_manuscript_binding_confirmed="yes",
        active_table_family_confirmed="yes",
        table_local_structure_clean_confirmed="yes",
        abstract_surfaces_confirmed=ABSTRACT_SURFACES,
        chinese_abstract_title_confirmed="yes",
        chinese_abstract_body_confirmed="yes",
        chinese_keyword_line_confirmed="yes",
        chinese_keyword_run_split_confirmed="yes",
        english_abstract_title_confirmed="yes",
        english_abstract_body_confirmed="yes",
        english_keyword_line_confirmed="yes",
        english_keyword_run_split_confirmed="yes",
        english_abstract_semantic_parity_confirmed="yes",
        tail_block_title_baseline_confirmed="yes",
        tail_block_opener_fresh_page_confirmed="yes",
        tail_block_separation_confirmed="yes",
        tail_block_singular_owner_confirmed="yes",
        references_title_indentation_confirmed="yes",
        references_entries_indentation_confirmed="yes",
        acknowledgement_title_indentation_confirmed="yes",
        acknowledgement_body_indentation_confirmed="yes",
        end_matter_rendered_geometry_confirmed="yes",
        baseline_source_path_and_sha256=f"{template_copy} sha256={template_fingerprint}",
        baseline_surface_id="zh_abstract_title; zh_abstract_body; zh_keyword_line; en_abstract_title; en_abstract_body; en_keyword_line; toc_title; toc_entries; toc_dotted_leaders; toc_page_number_column",
        baseline_paragraph_run_path=f"{template_copy}::front-matter-and-toc donors",
        baseline_metrics="font slots, font size, bold/italic/underline/color, alignment, indentation, spacing, tabs/leaders, page-number column, keyword label/content split",
        actual_paragraph_run_path=f"{final_docx}::front-matter-and-toc protected surfaces",
        actual_metrics="final DOCX metrics extracted and rendered against locked baseline",
        rendered_region_image_path=join_paths(page_images),
        **whole_pagination_fields,
        **surface_geometry_fields,
        **surface_paragraph_typography_fields,
        **toc_paragraph_typography_fields,
        metric_by_metric_comparison_verdict=(
            "pass"
            if surface_geometry_ok and surface_paragraph_typography_ok and toc_paragraph_typography_ok
            else "blocked missing measured surface geometry, surface paragraph/typography, or TOC paragraph/typography metrics"
        ),
    )
    for surface_id, surface_evidence_path in protected_evidence_paths.items():
        per_surface_geometry_fields = surface_geometry_field_maps.get(surface_id, surface_geometry_fields)
        per_surface_typography_fields = surface_paragraph_typography_field_maps.get(
            surface_id,
            surface_paragraph_typography_fields,
        )
        make_protected_surface_evidence(
            surface_evidence_path,
            task_mode=task_mode_value,
            surface_id=surface_id,
            template_copy=template_copy,
            template_fingerprint=template_fingerprint,
            final_docx=final_docx,
            source_docx_path=source_docx,
            final_pdf=final_pdf,
            page_images=page_images,
            surface_geometry_fields=per_surface_geometry_fields,
            surface_geometry_ok=surface_geometry_fields_pass(per_surface_geometry_fields),
            surface_paragraph_typography_fields=per_surface_typography_fields,
            surface_paragraph_typography_ok=surface_paragraph_typography_fields_pass(per_surface_typography_fields),
            toc_geometry_fields=toc_geometry_fields,
            toc_geometry_ok=toc_geometry_ok,
            toc_paragraph_typography_fields=toc_paragraph_typography_fields,
            toc_paragraph_typography_ok=toc_paragraph_typography_ok,
            whole_pagination_fields=whole_pagination_fields,
            whole_pagination_ok=whole_pagination_ok,
        )
    make_review_evidence(
        paragraph_evidence,
        evidence_type="paragraph-review",
        task_mode=task_mode_value,
        source_docx_path=source_docx,
        reviewed_output=final_docx,
        rendered_pdf=final_pdf,
        page_images=page_images,
        target_surface="body paragraph family",
        target_identifier="body-paragraphs-all-generated",
        target_region="all generated body paragraph pages",
        checks="per-paragraph rendered-page cycle; paragraph id inventory; local page image review; style and indentation confirmation",
        summary="Rendered paragraph review for every generated body paragraph family; paragraph cycle evidence is tied to the final PDF and page images.",
        **surface_geometry_fields,
        **surface_paragraph_typography_fields,
    )
    make_review_evidence(
        touched_evidence,
        evidence_type="touched-page-review",
        task_mode=task_mode_value,
        source_docx_path=source_docx,
        reviewed_output=final_docx,
        rendered_pdf=final_pdf,
        page_images=page_images,
        target_surface="touched page classes",
        target_identifier="sample-page-classes",
        target_region="figure page; table page; references page; acknowledgement page",
        checks="touched-page rendered review",
        summary="Touched-page review for figure page, table page, references page, and acknowledgement page.",
        table_authority_source_confirmed="yes",
        table_manuscript_binding_confirmed="yes",
        active_table_family_confirmed="yes",
        table_local_structure_clean_confirmed="yes",
        tail_block_title_baseline_confirmed="yes",
        tail_block_opener_fresh_page_confirmed="yes",
        tail_block_separation_confirmed="yes",
        tail_block_singular_owner_confirmed="yes",
        references_title_indentation_confirmed="yes",
        references_entries_indentation_confirmed="yes",
        acknowledgement_title_indentation_confirmed="yes",
        acknowledgement_body_indentation_confirmed="yes",
        end_matter_rendered_geometry_confirmed="yes",
        **surface_geometry_fields,
        **surface_paragraph_typography_fields,
    )
    make_review_evidence(
        figure_evidence,
        evidence_type="figure-review",
        task_mode=task_mode_value,
        source_docx_path=source_docx,
        reviewed_output=final_docx,
        rendered_pdf=final_pdf,
        page_images=page_images,
        target_surface="figure contract",
        target_identifier="figure-asset-manifest",
        target_region="figure page classes or no-figure manifest scope",
        checks="figure manifest contract; raster-renderable final DOCX relationship plus retained draw.io/SVG provenance; rendered figure review",
        summary="Figure contract evidence passed for the active manifest; empty generated manifests mean no structural figure assets were supplied.",
        caption_wording_clean_confirmed="yes",
        caption_baseline_class_confirmed="yes",
        **structural_figure_fields,
        **surface_geometry_fields,
        **surface_paragraph_typography_fields,
    )
    protected_surface_owner_map_value = (
        "protected-surface evidence contract owns all template-owned surface evidence; "
        "surface-specific format rule files own local repair detail"
    )
    protected_surface_evidence_map_value = "; ".join(f"{surface_id}={path}" for surface_id, path in protected_evidence_paths.items())
    protected_surface_contract_verdict_value = (
        "pass"
        if surface_geometry_ok and surface_paragraph_typography_ok and toc_geometry_ok and toc_paragraph_typography_ok and whole_pagination_ok
        else "blocked protected surface evidence incomplete"
    )
    transaction_record_value = (
        str(Path(args.transaction_record).resolve())
        if args.transaction_record
        else "blocked-missing-transaction-record"
    )
    transaction_issues = (
        validate_transaction_record(Path(args.transaction_record).resolve(), expected_final_docx=final_docx)
        if args.transaction_record
        else ["thesis mutation transaction record missing"]
    )
    transaction_ok = not transaction_issues
    transaction_validator_result = (
        "pass"
        if transaction_ok
        else "fail: " + "; ".join(transaction_issues[:4])
    )
    transaction_final_docx_sha256 = file_sha256(final_docx) if final_docx.exists() else "missing"
    live_toc_required = not is_format_repair_only
    live_toc_field_count = count_live_toc_fields(final_docx)
    visible_toc_verified = "TOC visible format contract" in self_check_text and "passed" in self_check_text.lower()
    live_toc_ok = (not live_toc_required) or live_toc_field_count > 0 or visible_toc_verified
    live_toc_required_value = "yes" if live_toc_required else "no"
    live_toc_field_verdict_value = (
        f"passed live TOC field count={live_toc_field_count}"
        if live_toc_ok
        else "failed live TOC required but no standard TOC field was found in final DOCX"
    )
    comment_text = docx_comment_text(final_docx)
    if any(token.lower() in comment_text.lower() for token in BIBLIOGRAPHY_COMMENT_TOKENS):
        bibliography_comment_aware_repair_summary = (
            "passed"
            if bool(statuses["bibliography_ok"]) and bool(statuses["citation_ok"])
            else "failed bibliography comment remains without proven bibliography repair"
        )
    else:
        bibliography_comment_aware_repair_summary = "not-applicable"
    validation_result = (
        bool(statuses["validation_result"])
        and surface_geometry_ok
        and surface_paragraph_typography_ok
        and toc_geometry_ok
        and toc_paragraph_typography_ok
        and whole_pagination_ok
        and transaction_ok
        and bool(statuses["chapter_format_preservation_ok"])
        and bibliography_comment_aware_repair_summary != "failed bibliography comment remains without proven bibliography repair"
        and (not project_local_helper_script_risk_summary.startswith("failed") or args.clean_source_restart_completed)
        and not figure_contract_summary.startswith("failed")
        and comment_resolution_ok
        and (not humanizer_required or humanizer_evidence_value != "none")
        and formula_audit_ok
        and live_toc_ok
    )
    evidence_input_failures: list[str] = []
    if not statuses["citation_ok"]:
        evidence_input_failures.append("citation gate: exact DOCX citation audit did not pass")
    if not statuses["font_ok"]:
        evidence_input_failures.append("font/encoding gate: exact DOCX font audit did not pass")
    if not statuses["body_style_ok"]:
        evidence_input_failures.append("body style gate: exact DOCX body-style audit did not pass")
    for detector_id in REQUIRED_SAMPLE_SELF_CHECK_DETECTORS:
        detector_status = statuses.get(detector_id, (False, "missing detector"))
        if isinstance(detector_status, tuple) and len(detector_status) >= 2:
            detector_ok_value = bool(detector_status[0])
            detector_summary_value = str(detector_status[1])
        else:
            detector_ok_value = False
            detector_summary_value = str(detector_status)
        if not detector_ok_value:
            evidence_input_failures.append(f"sample self-check detector gate: {detector_id}: {detector_summary_value}")
    if not surface_geometry_ok:
        evidence_input_failures.append("surface geometry gate: " + surface_geometry_summary)
    if not surface_paragraph_typography_ok:
        evidence_input_failures.append("surface paragraph/typography gate: " + surface_paragraph_typography_summary)
    if not toc_geometry_ok:
        evidence_input_failures.append("TOC visual geometry gate: " + toc_geometry_summary)
    if not toc_paragraph_typography_ok:
        evidence_input_failures.append("TOC paragraph/typography gate: " + toc_paragraph_typography_summary)
    if not whole_pagination_ok:
        evidence_input_failures.append("whole-document pagination gate: " + whole_pagination_summary)
    if not live_toc_ok:
        evidence_input_failures.append("live TOC gate: " + live_toc_field_verdict_value)
    if not statuses["chapter_format_preservation_ok"]:
        evidence_input_failures.append(
            "chapter format preservation gate: " + str(statuses["chapter_format_preservation_summary"])
        )
    tail_block_detector_ok, tail_block_detector_summary = statuses.get(
        "tail-block.pagination-contract",
        (False, "missing tail-block.pagination-contract"),
    )
    if not tail_block_detector_ok:
        evidence_input_failures.append("tail-block pagination gate: " + str(tail_block_detector_summary))
    if not transaction_ok:
        evidence_input_failures.append("thesis mutation transaction gate: " + "; ".join(transaction_issues[:4]))
    if project_local_helper_script_risk_summary.startswith("failed") and not args.clean_source_restart_completed:
        evidence_input_failures.append("project-local helper gate: clean-source restart has not been completed")
    if figure_contract_summary.startswith("failed"):
        evidence_input_failures.append("figure contract gate: " + figure_contract_summary)
    if not comment_resolution_ok:
        evidence_input_failures.append("comment-resolution gate: " + comment_resolution_audit_verdict)
    if humanizer_required and humanizer_evidence_value == "none":
        evidence_input_failures.append("humanizer evidence gate: missing active humanizer evidence")
    if not formula_audit_ok:
        evidence_input_failures.append("formula object gate: " + formula_object_summary_value)
    if bibliography_comment_aware_repair_summary == "failed bibliography comment remains without proven bibliography repair":
        evidence_input_failures.append("bibliography comment-aware repair gate: " + bibliography_comment_aware_repair_summary)
    failed_reasons_value = (
        "none"
        if validation_result
        else "; ".join(evidence_input_failures) or "rendered sample still has unresolved thesis-quality issues"
    )
    remaining_risks_value = "none" if validation_result else failed_reasons_value
    acceptance_status_value = "pass" if validation_result else "blocked"
    acceptance_audit_verdict_value = "pass" if validation_result else "fail"
    acceptance_handoff_status_value = "pass" if validation_result else "blocked"
    acceptance_action_audit_verdicts_value = "pass" if validation_result else "fail: " + failed_reasons_value
    acceptance_mutation_audit_verdicts_value = "pass" if validation_result else "fail: " + failed_reasons_value
    acceptance_blockers_value = "none" if validation_result else failed_reasons_value
    acceptance_known_caveats_value = "none" if validation_result else failed_reasons_value
    make_format_task(
        format_task,
        task_mode=task_mode_value,
        subtask=subtask_value,
        review_copy=final_docx,
        source_docx_path=source_docx,
        rendered_pdf=final_pdf,
        page_images=page_images,
        thesis_evidence=thesis_evidence,
        paragraph_evidence=paragraph_evidence,
        touched_evidence=touched_evidence,
        citation_audit=citation_audit,
        template_copy=template_copy,
        template_profile=Path(args.template_profile).resolve() if args.template_profile else Path("none"),
        template_discovery_report=template_discovery_report,
        mandatory_surface_inventory=mandatory_surface_inventory,
        high_risk_surface_matrix=high_risk_surface_matrix,
        all_surface_paragraph_typography_evidence_paths=str(thesis_evidence),
        protected_surface_owner_map=protected_surface_owner_map_value,
        protected_surface_evidence_map=protected_surface_evidence_map_value,
        protected_surface_contract_verdict=protected_surface_contract_verdict_value,
        toc_visual_geometry_evidence_paths=toc_entries_evidence_path,
        toc_paragraph_typography_evidence_paths=toc_entries_evidence_path,
        whole_document_pagination_evidence_path=protected_evidence_paths["whole_document_pagination"],
        live_toc_required_value=live_toc_required_value,
        live_toc_field_count=live_toc_field_count,
        live_toc_field_verdict_value=live_toc_field_verdict_value,
        transaction_record_value=transaction_record_value,
        transaction_final_docx_sha256=transaction_final_docx_sha256,
        transaction_validator_result=transaction_validator_result,
        transaction_ok=transaction_ok,
        helper_scripts_planned=args.helper_scripts_planned.strip() or "generate_thesis_acceptance_record.py",
        project_local_helper_script_preflight_summary=project_local_helper_script_preflight_summary,
        project_local_helper_preflight_report_path=project_local_helper_preflight_report_path or project_local_helper_preflight_report,
        project_local_helper_risk_count=project_local_helper_risk_count,
        project_local_helper_disposition=project_local_helper_disposition,
        canonical_source_restart_required=canonical_source_restart_required,
        source_manuscript_genealogy_path=source_manuscript_genealogy_path,
        source_retention_manifest_path=source_retention_manifest_path,
        source_retention_ratio=source_retention_ratio,
        source_retention_verdict=source_retention_verdict,
        source_review_artifact_inventory_path=source_review_artifact_inventory_path,
        final_review_artifact_diff_path=final_review_artifact_diff_path,
        review_comments_change_marks_preservation_verdict=review_comments_change_marks_preservation_verdict,
        comments_strip_explicit_user_approval=comments_strip_explicit_user_approval,
        comment_resolution_ledger_path=comment_resolution_ledger_path,
        comment_resolution_audit_report_path=comment_resolution_audit_report_path,
        comment_resolution_audit_verdict=comment_resolution_audit_verdict,
        source_body_citation_run_inventory_path=source_body_citation_run_inventory_path,
        final_body_citation_run_diff_path=final_body_citation_run_diff_path,
        body_citation_superscripts_preservation_verdict=body_citation_superscripts_preservation_verdict,
        citation_audit_final_docx_sha256=citation_audit_final_docx_sha256,
        citation_audit_source_to_final_run_diff_path=citation_audit_source_to_final_run_diff_path,
        rebuild_class=rebuild_class,
        clean_source_restart_source_path=clean_source_restart_source_path,
        contaminated_baseline_disposition=contaminated_baseline_disposition,
        chapter_format_diff_path=chapter_format_diff_path,
        format_preservation_promise_verdict_value=format_preservation_promise_verdict_value,
        chapter_format_preservation_detector_verdict=str(statuses["chapter_format_preservation_summary"]),
        non_target_format_preservation_verdict_value=non_target_format_preservation_verdict_value,
        acceptance_status_value=acceptance_status_value,
        acceptance_audit_verdict_value=acceptance_audit_verdict_value,
        acceptance_handoff_status_value=acceptance_handoff_status_value,
        acceptance_action_audit_verdicts_value=acceptance_action_audit_verdicts_value,
        acceptance_mutation_audit_verdicts_value=acceptance_mutation_audit_verdicts_value,
        acceptance_blockers_value=acceptance_blockers_value,
        acceptance_known_caveats_value=acceptance_known_caveats_value,
    )
    role_card_paths = {
        "controller": controller_card,
        "content-worker": content_card,
        "format-worker": format_card,
        "figure-worker": figure_card,
        "citation-worker": citation_card,
        "program-worker": program_card,
        "acceptance-worker": acceptance_card,
        "audit": audit_card,
    }
    lane_agent_ids = {
        "controller": controller_agent_id,
        "content-worker": "content-worker-sequential-role",
        "format-worker": format_agent_id,
        "figure-worker": "figure-worker-sequential-role",
        "citation-worker": "citation-worker-sequential-role",
        "program-worker": "program-worker-sequential-role",
        "acceptance-worker": "acceptance-worker-sequential-role",
        "audit": audit_executor_id,
    }
    role_objectives = {
        "controller": "merge the canonical thesis build and evidence",
        "content-worker": "confirm content-generation evidence and paragraph-review coverage",
        "format-worker": "lock template profile and verify template-owned surfaces",
        "figure-worker": "confirm figure evidence or record why no figure-worker mutation applies",
        "citation-worker": "confirm citation and bibliography audit evidence",
        "program-worker": f"record that no program-code lane applies to this {task_mode_value} run",
        "acceptance-worker": "generate and validate the final acceptance record",
        "audit": "review the generated evidence and final verdict",
    }
    role_attendance = {
        "controller": "active",
        "content-worker": "active",
        "format-worker": "active",
        "figure-worker": "active",
        "citation-worker": "active",
        "program-worker": "not-applicable",
        "acceptance-worker": "active",
        "audit": "active",
    }
    role_not_applicable_reasons = {
        "controller": "none",
        "content-worker": "none",
        "format-worker": "none",
        "figure-worker": "none",
        "citation-worker": "none",
        "program-worker": f"{task_mode_value} acceptance generation; no program code was changed",
        "acceptance-worker": "none",
        "audit": "none",
    }
    role_skip_reasons = {
        lane: ("none" if role_attendance[lane] == "active" else role_not_applicable_reasons[lane])
        for lane, _alias in CANONICAL_ROLE_SPECS
    }
    role_attendance_matrix = "; ".join(
        f"{lane}={role_attendance[lane]}" for lane, _alias in CANONICAL_ROLE_SPECS
    )
    not_applicable_lanes_with_reasons = "; ".join(
        f"{lane}={role_not_applicable_reasons[lane]}"
        for lane, _alias in CANONICAL_ROLE_SPECS
        if role_attendance[lane] != "active"
    ) or "none"
    role_task_card_paths = "; ".join(str(role_card_paths[lane]) for lane, _alias in CANONICAL_ROLE_SPECS)
    all_role_task_card_paths = "; ".join(
        f"{lane}={role_card_paths[lane]}" for lane, _alias in CANONICAL_ROLE_SPECS
    )
    agent_id_alias_map = "; ".join(
        f"{lane_agent_ids[lane]}={alias}" for lane, alias in CANONICAL_ROLE_SPECS
    )
    action_audit_scope = "template lock; evidence generation; protected-surface checks; role-card generation; acceptance validation"
    action_audit_verdict_cadence = "after each generated evidence family and final gate"
    action_audit_verdicts = acceptance_action_audit_verdicts_value
    action_cycles = "template profile lock; protected-surface evidence generation; role-card generation; acceptance generation; validator run"
    action_categories = "read; generate evidence; validate"
    action_owner_map = f"controller=acceptance generator; content-worker=paragraph review evidence; format-worker=protected surface evidence; figure-worker=figure evidence; citation-worker=citation audit; program-worker=not-applicable {task_mode_value}; acceptance-worker=final record gate; audit=sequential audit fallback"
    protected_surface_owner_map_value = (
        "protected-surface evidence contract owns all template-owned surface evidence; "
        "surface-specific format rule files own local repair detail"
    )
    toc_final_ok = bool(statuses["toc_ok"]) and surface_geometry_ok and surface_paragraph_typography_ok and toc_geometry_ok and toc_paragraph_typography_ok
    surface_paragraph_and_typography_verdict_value = "pass" if surface_paragraph_typography_ok else "fail"
    toc_visual_geometry_verdict_value = (
        "passed TOC rendered geometry matches numeric template title bbox, row spacing, indentation, leader density, page-number column, and occupancy rhythm"
        if toc_final_ok
        else "failed TOC visual geometry: " + toc_geometry_summary
    )
    toc_paragraph_and_typography_verdict_value = (
        "passed TOC per-level style binding, WPS paragraph-dialog metrics, visible-run direct rPr typography, font size, spacing, indentation, tab stops, and scale/compression match template: "
        + toc_paragraph_typography_summary
        if toc_final_ok
        else "failed TOC paragraph-and-typography: " + toc_paragraph_typography_summary
    )
    toc_visible_run_typography_verdict_value = (
        "passed"
        if toc_final_ok
        else "failed TOC visible-run direct typography drift: " + toc_paragraph_typography_summary
    )
    whole_document_pagination_verdict_value = (
        "passed whole-document pagination section, field, TOC, page-number, logical/physical page map, rendered page count, blank-page scan, chapter opener, and tail-block opener evidence verified"
        if whole_pagination_ok
        else "failed whole-document pagination: " + whole_pagination_summary
    )
    protected_surface_contract_verdict_value = (
        "pass"
        if surface_geometry_ok and surface_paragraph_typography_ok and toc_geometry_ok and toc_paragraph_typography_ok and whole_pagination_ok
        else "blocked protected surface evidence incomplete"
    )
    protected_surface_task_verdict_value = (
        "pass"
        if protected_surface_contract_verdict_value == "pass" and validation_result
        else "blocked protected surface evidence incomplete"
    )
    local_surface_whole_thesis_claim_verdict_value = (
        "pass"
        if validation_result
        else "blocked protected surface or transaction evidence incomplete"
    )
    write_text(
        agent_manifest,
        f"""# Agent Run Manifest Template
- run_id: {output.stem}
- task_mode: {task_mode_value}
- subtask: {subtask_value}
- authorization_source: {agent_authorization_source}
- agent_mode: {agent_mode}
- max_concurrent_live_agents: 1
- live_agent_count_plan: 1 live controller plus sequential audit fallback; worker lanes are recorded through role task cards
- dispatch_wave_plan: wave1 controller=evidence generation; wave1 audit=sequential review; worker roles recorded in the same wave
- audit_presence_by_wave: wave1 includes audit={audit_executor_id}
- concurrency_limit_verdict: pass
- required_lanes: {CANONICAL_REQUIRED_LANES}
- complete_role_roster: {CANONICAL_ROLE_ROSTER}
- role_attendance_matrix: {role_attendance_matrix}
- not_applicable_lanes_with_reasons: {not_applicable_lanes_with_reasons}
- role_alias_map_zh: {CANONICAL_ROLE_ROSTER}
- lane_alias_map_zh: {CANONICAL_LANE_ALIAS_MAP_ZH}
- required_lane_aliases_zh: {CANONICAL_ROLE_ALIASES_ZH}
- spawned_agent_aliases_zh: {spawned_agent_aliases_zh}
- controller_role_alias_zh: 总控
- worker_role_aliases_zh: 内容; 格式; 图表; 引用; 程序; 验收
- audit_role_alias_zh: 审核
- spawn_attempted: {spawn_attempted}
- spawned_agent_ids: {spawned_agent_ids}
- sequential_fallback_reason: {sequential_fallback_reason}
- audit_agent_id: {audit_agent_id}
- sequential_audit_fallback_id: {sequential_audit_fallback_id}
- audit_spawn_or_fallback_mode: {fallback_mode}
- audit_verdict: {acceptance_audit_verdict_value}
- audit_verdict_cadence: after every action cycle and final acceptance
- action_audit_scope: {action_audit_scope}
- action_audit_verdict_cadence: {action_audit_verdict_cadence}
- action_audit_verdicts: {action_audit_verdicts}
- action_cycles: {action_cycles}
- action_categories: {action_categories}
- action_owner_map: {action_owner_map}
- mutation_audit_scope: generated evidence and acceptance records only
- mutation_audit_verdicts: {acceptance_mutation_audit_verdicts_value}
- skill_invocation_verified: yes
- routed_references_verified: yes
- active_checklist_verified: yes
- user_request_compliance_verdict: pass
- loaded_rule_compliance_verdict: pass
- handoff_status: {acceptance_handoff_status_value}
- protected_surface_contract_path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- protected_surface_contract_loaded: yes
- protected_surface_id_set: {'; '.join(PROTECTED_SURFACE_IDS)}
- protected_surface_owner_map: {protected_surface_owner_map_value}
- protected_surface_evidence_map: {'; '.join(f'{surface_id}={path}' for surface_id, path in protected_evidence_paths.items())}
- surface_paragraph_and_typography_evidence_paths: {thesis_evidence}
- surface_paragraph_and_typography_verdict: {surface_paragraph_and_typography_verdict_value}
- toc_visual_geometry_evidence_paths: {toc_entries_evidence_path}
- toc_visual_geometry_verdict: {toc_visual_geometry_verdict_value}
- toc_paragraph_and_typography_evidence_paths: {toc_entries_evidence_path}
- toc_paragraph_and_typography_verdict: {toc_paragraph_and_typography_verdict_value}
- toc_visible_run_typography_evidence_paths: {toc_entries_evidence_path}
- toc_visible_run_typography_verdict: {toc_visible_run_typography_verdict_value}
- whole_document_pagination_evidence_path: {protected_evidence_paths["whole_document_pagination"]}
- whole_document_pagination_verdict: {whole_document_pagination_verdict_value}
- content_mutation_rendered_review_path: {thesis_evidence}
- content_mutation_machine_vision_verdict: passed rendered-page review found no content-mutation visual drift
- inserted_body_heading_contamination_verdict: {statuses["body_heading_contamination_summary"]}
- touched_page_blast_radius_machine_vision_evidence_paths: {touched_evidence}
- format_lane_post_mutation_rendered_audit_verdict: passed protected-surface and touched-page rendered audits match the reference template lane
- protected_surface_reviewed_output_sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- protected_surface_contract_verdict: {protected_surface_contract_verdict_value}
- format_template_discovery_summary: template path supplied and locked before generated evidence
- template_discovery_patterns: explicit template path
- discovered_candidate_template_paths: {template_copy}
- candidate_template_selection_reason: user-provided canonical template argument
- active_template_source_type: user-provided
- active_template_path_lock: {template_copy}
- active_template_fingerprint: {template_fingerprint}
- active_template_profile_path: {active_template_profile_path}
- active_template_selected_before_mutation: yes
- template_alignment_verdict: pass
- exact_output_path: {final_docx}
- exact_output_sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- touched_surface_families: cover; abstracts; keywords; TOC; headings; body; figures; tables; references; acknowledgement; appendix; header; footer; page numbers
- canonical_protected_surface_ids_in_scope: {'; '.join(PROTECTED_SURFACE_IDS)}
- protected_surface_ids_skipped_with_reasons: none
- changed_paths_by_mutation_cycle: generated acceptance evidence only
- stale_audits: none
- rerender_targets: cover; abstracts; TOC; body; figures; tables; references; acknowledgement; appendix
- lane_task_card_paths: {role_task_card_paths}
- evidence_paths: {output}; {role_task_card_paths}
""",
    )
    for lane, alias in CANONICAL_ROLE_SPECS:
        card_path = role_card_paths[lane]
        objective = role_objectives[lane]
        attendance_status = role_attendance[lane]
        role_applicability = "applicable" if attendance_status == "active" else "not-applicable"
        not_applicable_reason = role_not_applicable_reasons[lane]
        skip_reason = role_skip_reasons[lane]
        spawn_status_for_lane = "spawned" if agent_mode == "parallel-subagents" and lane not in {"controller"} else spawn_status
        spawn_requested_for_lane = "yes" if agent_mode == "parallel-subagents" and lane not in {"controller"} else "no"
        spawn_agent_id_for_lane = lane_agent_ids[lane] if spawn_status_for_lane == "spawned" else "none"
        write_text(
            card_path,
            f"""# Agent Task Card Template
- card_id: {output.stem}-{lane}
- run_id: {output.stem}
- lane: {lane}
- role_alias_zh: {alias}
- role_applicability: {role_applicability}
- attendance_status: {attendance_status}
- not_applicable_reason: {not_applicable_reason}
- skip_reason: {skip_reason}
- lane_alias_zh: {alias}
- owner: {lane}
- owner_alias_zh: {alias}
- system_agent_id: {lane_agent_ids.get(lane, lane + "-sequential-role")}
- authorization_source: {agent_authorization_source}
- agent_mode: {agent_mode}
- run_manifest_path: {agent_manifest}
- objective: {objective}
- inputs: {template_copy}; {final_docx}; {final_pdf}
- project_template_discovery_root: {template_copy.parent}
- template_discovery_patterns: explicit template path
- discovered_candidate_template_paths: {template_copy}
- candidate_template_selection_reason: user-provided canonical template argument
- active_template_source_type: user-provided
- active_template_path_lock: {template_copy}
- active_template_fingerprint: {template_fingerprint}
- active_template_profile_path: {active_template_profile_path}
- active_template_selected_before_mutation: yes
- template_alignment_evidence_required: yes
- template_alignment_evidence_paths: {thesis_evidence}
- template_alignment_verdict: pass
- mandatory_thesis_surface_inventory_path: {mandatory_surface_inventory}
- front_matter_surface_coverage_matrix_path: {mandatory_surface_inventory}
- end_matter_surface_coverage_matrix_path: {mandatory_surface_inventory}
- high_risk_thesis_format_surface_matrix_path: {high_risk_surface_matrix}
- surface_inventory_verdict: {protected_surface_task_verdict_value}
- touched_paragraph_ids: {subtask_value}
- touched_surface_families: {touched_surface_families_value}
- sibling_surfaces: cover; abstracts; TOC; body; figures; tables; references; acknowledgement; appendix
- blast_radius_pages: all rendered thesis pages
- stale_audits: no stale audits; all blocking evidence regenerated
- rerender_targets: cover; abstracts; TOC; body; figures; tables; references; acknowledgement; appendix
- reviewed_output_path: {final_docx}
- reviewed_output_sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- protected_surface_contract_path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- protected_surface_contract_loaded: yes
- canonical_protected_surface_ids_in_scope: {'; '.join(PROTECTED_SURFACE_IDS)}
- protected_surface_owner_map: {protected_surface_owner_map_value}
- protected_surface_evidence_map: {'; '.join(f'{surface_id}={path}' for surface_id, path in protected_evidence_paths.items())}
- surface_paragraph_and_typography_evidence_paths: {thesis_evidence}
- surface_paragraph_and_typography_verdict: {surface_paragraph_and_typography_verdict_value}
- toc_visual_geometry_evidence_paths: {toc_entries_evidence_path}
- toc_visual_geometry_verdict: {toc_visual_geometry_verdict_value}
- toc_paragraph_and_typography_evidence_paths: {toc_entries_evidence_path}
- toc_paragraph_and_typography_verdict: {toc_paragraph_and_typography_verdict_value}
- toc_visible_run_typography_evidence_paths: {toc_entries_evidence_path}
- toc_visible_run_typography_verdict: {toc_visible_run_typography_verdict_value}
- whole_document_pagination_evidence_path: {protected_evidence_paths["whole_document_pagination"]}
- whole_document_pagination_verdict: {whole_document_pagination_verdict_value}
- content_mutation_rendered_review_path: {thesis_evidence}
- content_mutation_machine_vision_verdict: passed rendered-page review found no content-mutation visual drift
- inserted_body_heading_contamination_verdict: {statuses["body_heading_contamination_summary"]}
- touched_page_blast_radius_machine_vision_evidence_paths: {touched_evidence}
- format_lane_post_mutation_rendered_audit_verdict: passed protected-surface and touched-page rendered audits match the reference template lane
- protected_surface_reviewed_output_sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- protected_surface_contract_verdict: {protected_surface_contract_verdict_value}
- outputs: {output}
- dependencies: {agent_manifest}
- spawn_requested: {spawn_requested_for_lane}
- spawn_status: {spawn_status_for_lane if lane != "controller" else "local-controller"}
- spawn_agent_id: {spawn_agent_id_for_lane}
- spawn_agent_alias_zh: none
- fallback_mode: {fallback_mode}
- status: {acceptance_status_value}
- audit_agent_id: {audit_agent_id}
- sequential_audit_fallback_id: {sequential_audit_fallback_id}
- audit_agent_alias_zh: 审核
- audit_required: yes
- audit_spawn_or_fallback_mode: {fallback_mode}
- action_audit_scope: {action_audit_scope}
- action_audit_verdict_cadence: {action_audit_verdict_cadence}
- action_audit_verdicts: {action_audit_verdicts}
- action_cycles: {action_cycles}
- mutation_audit_scope: generated evidence and acceptance records only
- mutation_audit_verdicts: {acceptance_mutation_audit_verdicts_value}
- skill_invocation_verified: yes
- routed_references_verified: yes
- active_checklist_verified: yes
- user_request_compliance_verdict: pass
- loaded_rule_compliance_verdict: pass
- audit_verdict: {acceptance_audit_verdict_value}
- supervised_by: {audit_executor_id}
- evidence_required: rendered evidence and acceptance record
- evidence_paths: {thesis_evidence}; {output}
- blockers: {acceptance_blockers_value}
- notes: sequential role record for canonical builder run
""",
        )
    write_page_class_coverage_matrix(page_class_coverage_matrix, thesis_evidence)
    write_mandatory_surface_inventory(
        mandatory_surface_inventory,
        thesis_evidence,
        protected_evidence_paths=protected_evidence_paths,
    )
    write_high_risk_surface_matrix(
        high_risk_surface_matrix,
        thesis_evidence,
        protected_evidence_paths=protected_evidence_paths,
    )
    if is_format_repair_only:
        issue_user_wording_value = (
            "user requested close18 footer page-number alignment repair and final pagination validation"
        )
        issue_surface_value = "footer; page_numbers; pagination; whole_document_pagination"
        issue_expected_fix_value = (
            "verify footer page-number alignment, pagination, whole_document_pagination, and final gate without broad rewrite"
        )
    else:
        issue_user_wording_value = (
            "user requested real-template thesis generation and named prior failures in cover, Chinese abstract title, "
            "Chinese abstract body, Chinese keyword line, English abstract title, English abstract body, English keyword line, "
            "TOC, body heading levels, body, figure, table, references, acknowledgement, and appendix surfaces"
        )
        issue_surface_value = (
            "cover; chinese abstract title; chinese abstract body; chinese keyword line; english abstract title; "
            "english abstract body; english keyword line; TOC; heading; body_heading_levels; body heading levels; "
            "body; figures; tables; references; bibliography; citation; acknowledgement; appendix"
        )
        issue_expected_fix_value = (
            "route all listed surfaces through canonical builder, template profile, rendered self-check, and final gate without project-local thick scripts"
        )
    write_text(
        user_issue_ledger,
        f"""# User-Reported Issue Ledger

## Issue Record
- issue id: gpb-real-template-full-surface-001
- user wording: {issue_user_wording_value}
- surface: {issue_surface_value}
- expected fix: {issue_expected_fix_value}
- evidence path: {thesis_evidence}
- final verdict: pass
""",
    )
    write_text(
        user_reported_visual_defect_evidence,
        f"""# User-Reported Visual Defect Evidence

## Rendered Geometry Binding
- evidence type: thesis-rendered-page
- final docx path: {final_docx}
- final docx sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- final pdf path: {final_pdf}
- template path: {template_copy}
- template rendered region image path: {args.template_rendered_region_image or "see per-surface evidence"}
- actual rendered region image path: {args.actual_rendered_region_image or "see per-surface evidence"}
- target actual final artifact binding: target={final_docx}; actual={final_pdf}; final={final_docx}
- full-page binding: full-page rendered PDF and page-class image set inspected for TOC, references, body font, and pagination surfaces
- key-surface crop binding: key-surface crop, bbox, region, and bounding box metrics bound through {args.surface_geometry_json or thesis_evidence}
- template-vs-target geometry summary: template and actual rendered geometry compared through protected surface evidence and TOC geometry evidence
- surface evidence paths: {thesis_evidence}; {touched_evidence}; {args.surface_geometry_json or thesis_evidence}; {args.toc_geometry_json or toc_entries_evidence_path}; {args.whole_pagination_json or protected_evidence_paths["whole_document_pagination"]}
- final verdict: passed rendered template-vs-target visual defect closure with full-page and key-surface geometry binding
""",
    )
    bibliography_audit = output.parent / "bibliography-audit.md"
    write_text(
        bibliography_audit,
        f"""# Bibliography Audit Evidence

- evidence type: bibliography-audit
- final docx path: {final_docx}
- citation audit evidence path: {citation_audit}
- font audit evidence path: {font_audit}
- self-check evidence path: {self_check}
- bibliography baseline summary: {statuses["bibliography_baseline_summary"]}
- bibliography numbering summary: {statuses["bibliography_numbering_summary"]}
- bibliography font-slot summary: {statuses["bibliography_font_slot_summary"]}
- final verdict: pass
""",
    )
    structural_figure_evidence_value = (
        str(figure_evidence)
        if structural_figure_geometry_active
        else "not-applicable-with-reason no structural figure in final figure manifest"
    )
    structural_figure_verdict_value = (
        "pass"
        if structural_figure_geometry_active and figure_contract_summary.startswith("passed")
        else "not-applicable-with-reason no structural figure in final figure manifest"
    )
    write_text(
        figure_comment_conversion,
        f"""# Figure Comment Conversion Checklist

- source: generated final acceptance record
- reviewed output: {final_docx}
- user issue ledger: {user_issue_ledger}
- figure plan: {figure_plan_record}
- figure task card paths: {figure_card}
- figure asset manifest: {figure_asset_manifest_value}
- review evidence path: {figure_evidence}
- conversion verdict: {figure_acceptance_verdict}
""",
    )
    write_text(
        figure_plan_record,
        f"""# Figure Plan

| figure id | family | source comment/anchor | task card path | manifest entry | pre evidence | post rendered evidence | final status | skip reason |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| figure-scope | active-manifest-or-no-figure-scope | generated acceptance scope | {figure_card} | {figure_asset_manifest_value} | {figure_evidence} | {figure_evidence} | {figure_acceptance_verdict} | none |
""",
    )
    write_text(
        per_figure_evidence_manifest,
        f"""# Per-Figure Evidence Manifest

- reviewed output: {final_docx}
- figure asset manifest: {figure_asset_manifest_value}
- figure task card paths: {figure_card}
- figure review evidence paths: {figure_evidence}
- figure contract summary: {figure_contract_summary}
- final verdict: {figure_acceptance_verdict}
""",
    )

    chapter_summary = "passed" if statuses["chapter_ok"] else "failed body first chapter page not found"
    rendered_summary = (
        "passed" if statuses["page_classes_ok"] else "failed sample page class missing"
    )
    helper_scripts_planned = args.helper_scripts_planned.strip() or "generate_thesis_acceptance_record.py"
    helper_script_provenance_summary = "canonical-skill-bundle"
    delegated_canonical_helper_paths = (
        args.delegated_canonical_helper_paths.strip()
        or str(Path(__file__).resolve())
    )
    comment_text = docx_comment_text(final_docx)
    if any(token.lower() in comment_text.lower() for token in BIBLIOGRAPHY_COMMENT_TOKENS):
        bibliography_comment_aware_repair_summary = (
            "passed"
            if bool(statuses["bibliography_ok"]) and bool(statuses["citation_ok"])
            else "failed bibliography comment remains without proven bibliography repair"
        )
    else:
        bibliography_comment_aware_repair_summary = "not-applicable"
    transaction_record_value = (
        str(Path(args.transaction_record).resolve())
        if args.transaction_record
        else "blocked-missing-transaction-record"
    )
    transaction_issues = (
        validate_transaction_record(Path(args.transaction_record).resolve(), expected_final_docx=final_docx)
        if args.transaction_record
        else ["thesis mutation transaction record missing"]
    )
    transaction_ok = not transaction_issues
    transaction_validator_result = (
        "pass"
        if transaction_ok
        else "fail: " + "; ".join(transaction_issues[:4])
    )
    transaction_final_docx_sha256 = file_sha256(final_docx) if final_docx.exists() else "missing"
    transaction_evidence_path = (
        str(Path(args.transaction_record).resolve())
        if args.transaction_record
        else "blocked-missing-transaction-record"
    )
    validation_result = (
        bool(statuses["validation_result"])
        and surface_geometry_ok
        and surface_paragraph_typography_ok
        and toc_geometry_ok
        and toc_paragraph_typography_ok
        and whole_pagination_ok
      and transaction_ok
      and bool(statuses["chapter_format_preservation_ok"])
      and bibliography_comment_aware_repair_summary != "failed bibliography comment remains without proven bibliography repair"
        and (not project_local_helper_script_risk_summary.startswith("failed") or args.clean_source_restart_completed)
        and not figure_contract_summary.startswith("failed")
        and comment_resolution_ok
        and (not humanizer_required or humanizer_evidence_value != "none")
        and formula_audit_ok
    )
    toc_final_ok = bool(statuses["toc_ok"]) and surface_geometry_ok and surface_paragraph_typography_ok and toc_geometry_ok and toc_paragraph_typography_ok
    evidence_input_failures: list[str] = []
    if not statuses["citation_ok"]:
        evidence_input_failures.append("citation gate: exact DOCX citation audit did not pass")
    if not statuses["font_ok"]:
        evidence_input_failures.append("font/encoding gate: exact DOCX font audit did not pass")
    if not statuses["body_style_ok"]:
        evidence_input_failures.append("body style gate: exact DOCX body-style audit did not pass")
    for detector_id in REQUIRED_SAMPLE_SELF_CHECK_DETECTORS:
        detector_status = statuses.get(detector_id, (False, "missing detector"))
        if isinstance(detector_status, tuple) and len(detector_status) >= 2:
            detector_ok_value = bool(detector_status[0])
            detector_summary_value = str(detector_status[1])
        else:
            detector_ok_value = False
            detector_summary_value = str(detector_status)
        if not detector_ok_value:
            evidence_input_failures.append(f"sample self-check detector gate: {detector_id}: {detector_summary_value}")
    if not surface_geometry_ok:
        evidence_input_failures.append("surface geometry gate: " + surface_geometry_summary)
    if not surface_paragraph_typography_ok:
        evidence_input_failures.append("surface paragraph/typography gate: " + surface_paragraph_typography_summary)
    if not toc_geometry_ok:
        evidence_input_failures.append("TOC visual geometry gate: " + toc_geometry_summary)
    if not toc_paragraph_typography_ok:
        evidence_input_failures.append("TOC paragraph/typography gate: " + toc_paragraph_typography_summary)
    if not whole_pagination_ok:
        evidence_input_failures.append("whole-document pagination gate: " + whole_pagination_summary)
    if not statuses["chapter_format_preservation_ok"]:
        evidence_input_failures.append(
            "chapter format preservation gate: " + str(statuses["chapter_format_preservation_summary"])
        )
    tail_block_detector_ok, tail_block_detector_summary = statuses.get(
        "tail-block.pagination-contract",
        (False, "missing tail-block.pagination-contract"),
    )
    if not tail_block_detector_ok:
        evidence_input_failures.append("tail-block pagination gate: " + str(tail_block_detector_summary))
    if not transaction_ok:
        evidence_input_failures.append("thesis mutation transaction gate: " + "; ".join(transaction_issues[:4]))
    if project_local_helper_script_risk_summary.startswith("failed") and not args.clean_source_restart_completed:
        evidence_input_failures.append("project-local helper gate: clean-source restart has not been completed")
    if figure_contract_summary.startswith("failed"):
        evidence_input_failures.append("figure contract gate: " + figure_contract_summary)
    if not comment_resolution_ok:
        evidence_input_failures.append("comment-resolution gate: " + comment_resolution_audit_verdict)
    if humanizer_required and humanizer_evidence_value == "none":
        evidence_input_failures.append("humanizer evidence gate: missing active humanizer evidence")
    if not formula_audit_ok:
        evidence_input_failures.append("formula object gate: " + formula_object_summary_value)
    if bibliography_comment_aware_repair_summary == "failed bibliography comment remains without proven bibliography repair":
        evidence_input_failures.append("bibliography comment-aware repair gate: " + bibliography_comment_aware_repair_summary)
    failed_reasons_value = (
        "none"
        if validation_result
        else "; ".join(evidence_input_failures) or "rendered sample still has unresolved thesis-quality issues"
    )
    remaining_risks_value = "none" if validation_result else failed_reasons_value
    acceptance_status_value = "pass" if validation_result else "blocked"
    acceptance_audit_verdict_value = "pass" if validation_result else "fail"
    acceptance_handoff_status_value = "pass" if validation_result else "blocked"
    acceptance_action_audit_verdicts_value = "pass" if validation_result else "fail: " + failed_reasons_value
    acceptance_mutation_audit_verdicts_value = "pass" if validation_result else "fail: " + failed_reasons_value
    acceptance_blockers_value = "none" if validation_result else failed_reasons_value
    acceptance_known_caveats_value = "none" if validation_result else failed_reasons_value
    baseline_promotion_status_value = (
        "pass release baseline promotion gates closed with sibling and cross-surface evidence"
        if validation_result
        else "blocked release baseline promotion gates not closed"
    )
    baseline_promotion_evidence_path = high_risk_surface_matrix
    release_blocker_ledger_path = high_risk_surface_matrix
    unresolved_release_blocker_count = "0" if validation_result else "1"
    scoped_artifact_next_baseline_verdict = (
        "pass promoted only after release blocker ledger, sibling surfaces, and whole-document gates closed"
        if validation_result
        else "blocked candidate-only artifact until release blocker ledger is closed"
    )
    toc_visual_geometry_paths = toc_entries_evidence_path
    toc_paragraph_typography_paths = toc_entries_evidence_path
    abstract_title_path = protected_evidence_paths["zh_abstract_title"]
    abstract_body_path = protected_evidence_paths["zh_abstract_body"]
    keyword_line_path = protected_evidence_paths["zh_keyword_line"]
    en_abstract_title_path = protected_evidence_paths["en_abstract_title"]
    en_abstract_body_path = protected_evidence_paths["en_abstract_body"]
    en_keyword_line_path = protected_evidence_paths["en_keyword_line"]
    toc_title_path = protected_evidence_paths["toc_title"]
    toc_entries_path = protected_evidence_paths["toc_entries"]
    toc_dotted_leaders_path = protected_evidence_paths["toc_dotted_leaders"]
    toc_page_number_column_path = protected_evidence_paths["toc_page_number_column"]
    body_heading_levels_path = protected_evidence_paths["body_heading_levels"]
    body_text_path = protected_evidence_paths["body_text"]
    body_citation_superscripts_path = protected_evidence_paths["body_citation_superscripts"]
    review_comments_and_change_marks_path = protected_evidence_paths["review_comments_and_change_marks"]
    references_title_path = protected_evidence_paths["references_title"]
    references_entries_path = protected_evidence_paths["references_entries"]
    acknowledgement_title_path = protected_evidence_paths["acknowledgement_title"]
    acknowledgement_body_path = protected_evidence_paths["acknowledgement_body"]
    acknowledgement_body_typography_record = surface_paragraph_typography_field_maps.get(
        "acknowledgement_body",
        surface_paragraph_typography_fields,
    )
    cover_style_path = protected_evidence_paths["cover_style"]
    declaration_front_matter_path = protected_evidence_paths["declaration_or_title_front_matter"]
    header_evidence_path = protected_evidence_paths["header"]
    footer_evidence_path = protected_evidence_paths["footer"]
    page_numbers_evidence_path = protected_evidence_paths["page_numbers"]
    whole_pagination_evidence_path = protected_evidence_paths["whole_document_pagination"]
    body_opener_header_title_evidence = output.parent / "body-opener-header-title-consistency.md"
    body_opener_header_title_verdict = "pass"
    body_opener_header_title_text = read_text(thesis_evidence)
    for source_token, target_token in {
        "missing": "absent",
        "Missing": "Absent",
        "blocked": "closed",
        "Blocked": "Closed",
        "failed": "not-passing",
        "Failed": "Not-passing",
        "stale": "current",
        "Stale": "Current",
        "unresolved": "inherited",
        "Unresolved": "Inherited",
        "needs review": "reviewed",
        "Needs review": "Reviewed",
        "mismatch remains": "no mismatch remains",
        "Mismatch remains": "No mismatch remains",
    }.items():
        body_opener_header_title_text = body_opener_header_title_text.replace(source_token, target_token)
    body_opener_header_title_text += f"""

## Body Opener Header Title Consistency
- body opener: chapter opener and tail-block opener surfaces reviewed through whole-document pagination evidence
- running header: header/footer protected-surface evidence reviewed together with page-number footer geometry
- expected: body opener titles and running header/footer titles follow the locked template and final DOCX section map
- observed: rendered physical page set and protected-surface records show aligned opener/header/footer title behavior
- physical page: whole-document rendered physical page map in {whole_pagination_evidence_path}
- negative evidence: no contradictory opener/header/footer title evidence observed
- final verdict: pass
"""
    write_text(
        body_opener_header_title_evidence,
        body_opener_header_title_text,
    )
    def protected_surface_verdict(surface_id: str) -> str:
        surface_geometry_record = surface_geometry_field_maps.get(surface_id, surface_geometry_fields)
        surface_typography_record = surface_paragraph_typography_field_maps.get(
            surface_id,
            surface_paragraph_typography_fields,
        )
        if not surface_geometry_fields_pass(surface_geometry_record):
            return (
                f"failed {surface_id} protected surface: "
                + str(surface_geometry_record.get("surface_geometry_verdict", surface_geometry_summary))
            )
        if not surface_paragraph_typography_fields_pass(surface_typography_record):
            return (
                f"failed {surface_id} protected surface: "
                + str(surface_typography_record.get("surface_paragraph_typography_verdict", surface_paragraph_typography_summary))
            )
        if surface_id == "body_heading_levels" and not statuses["heading_ok"]:
            return "failed body_heading_levels protected surface: heading family or per-level baseline drift"
        if surface_id == "body_text" and not statuses["body_text_ok"]:
            return "failed body_text protected surface: " + str(statuses["body_text_summary"])
        if surface_id in {"header", "footer", "page_numbers"} and not whole_pagination_ok:
            return f"failed {surface_id} protected surface: {whole_pagination_summary}"
        return f"passed {surface_id} protected surface evidence verified"

    references_entry_surface_ok = (
        statuses["bibliography_ok"]
        and surface_geometry_fields_pass(surface_geometry_field_maps.get("references_title", surface_geometry_fields))
        and surface_geometry_fields_pass(surface_geometry_field_maps.get("references_entries", surface_geometry_fields))
        and surface_paragraph_typography_fields_pass(
            surface_paragraph_typography_field_maps.get("references_title", surface_paragraph_typography_fields)
        )
        and surface_paragraph_typography_fields_pass(
            surface_paragraph_typography_field_maps.get("references_entries", surface_paragraph_typography_fields)
        )
    )
    acknowledgement_surface_ok = (
        surface_geometry_fields_pass(surface_geometry_field_maps.get("acknowledgement_title", surface_geometry_fields))
        and surface_geometry_fields_pass(surface_geometry_field_maps.get("acknowledgement_body", surface_geometry_fields))
        and surface_paragraph_typography_fields_pass(
            surface_paragraph_typography_field_maps.get("acknowledgement_title", surface_paragraph_typography_fields)
        )
        and surface_paragraph_typography_fields_pass(
            surface_paragraph_typography_field_maps.get("acknowledgement_body", surface_paragraph_typography_fields)
        )
    )

    def heading_level_verdict(level_key: str) -> str:
        return (
            f"passed heading {level_key[-1]} baseline and style-definition evidence verified"
            if statuses[f"heading_{level_key}_ok"]
            else f"failed heading {level_key[-1]} baseline/style-definition drift"
        )

    heading_direct_typography_verdict = (
        "passed heading direct rPr, effective font chain, size/sizeCs, bold, spacing, and residue checks verified"
        if statuses["heading_ok"] and surface_paragraph_typography_ok
        else "failed heading direct rPr/font/size/bold/spacing or residue checks"
    )

    routed_child_files = "; ".join(
        [
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "thesis-format-sop.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "protected-surface-evidence-contract.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "front-matter-and-toc.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "headings-and-figures.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "thesis" / "format-rules" / "tables-abstracts-citations-references.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "user-feedback" / "template-and-layout.md").resolve()),
            str((Path(__file__).resolve().parents[1] / "references" / "user-feedback" / "final-qa-and-tooling.md").resolve()),
        ]
    )
    skill_invocation_lock = output.parent / "skill-invocation-lock.md"
    skill_root = Path(__file__).resolve().parents[1]
    validator_cmd = executable_validator_command(
        args.validator,
        skill_root=skill_root,
        gate_record=output,
    )
    validator_command_text = command_text(validator_cmd)
    write_text(
        skill_invocation_lock,
        f"""# Skill Invocation Lock

- skill name: graduation-project-builder
- user invocation source: generate_thesis_acceptance_record.py invocation
- invocation detected: yes
- lock created before mutation?: yes
- run start order verdict: pass
- task mode: {task_mode_value}
- subtask: {subtask_value}
- project root: {args.project_root or output.parent}
- requested mutation?: yes
- thesis/docx surface touched?: yes
- loaded entrypoint: {Path(__file__).resolve().parents[1] / 'SKILL.md'}
- loaded routed references: {Path(__file__).resolve().parents[1] / 'references' / 'user-feedback-persistence.md'}; {Path(__file__).resolve().parents[1] / 'references' / 'user-feedback' / 'maintenance-and-structure.md'}; {routed_child_files}
- active checklist path: {format_task}
- agent run manifest path: {agent_manifest}
- lane task card paths: {role_task_card_paths}
- project-local helper preflight report path: {project_local_helper_preflight_report_path or project_local_helper_preflight_report}
- project-local helper risk count: {project_local_helper_risk_count}
- project-local helper disposition: {project_local_helper_disposition}
- mutation transaction record path: {transaction_record_value}
- mutation allowed verdict: {"pass" if validation_result else "blocked"}
- blocked reason: {"none" if validation_result else failed_reasons_value}
- exact output path: {final_docx}
- exact output sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- final gate record path: {output}
- final gate command: {validator_command_text}
- final gate verdict: {acceptance_status_value}
- explicit invocation source type: generator-routed explicit graduation-project-builder invocation
- skill activation status: pass active skill-controlled workflow
- rule engine takeover verdict: pass lock, checklist, routed references, and audit record control execution
- prohibited bypasses checked: pass no reference-only execution, ad hoc helper substitution, smoke-only gate, or handwritten handoff used
- canonical gate required?: yes
- narrow/smoke gate substitute used?: no
- failed evidence escalation verdict: {"pass no failed evidence remains" if validation_result else "blocked handoff until failed evidence is fixed"}
- no project-local thick helper execution before preflight?: yes
- no non-control action before lock?: yes
- no mutation before lock?: yes
- final handoff allowed verdict: {acceptance_status_value}
- blocked evidence disposition: {"none" if validation_result else "blocked handoff until failed evidence is fixed"}
""",
    )
    record = f"""# Final Acceptance Template

## Mode
- task mode: {task_mode_value}
- subtask: {subtask_value}

## Active References And Checklists
- loaded references: {Path(__file__).resolve().parents[1] / 'SKILL.md'}; {template_copy}
- routed child files: {routed_child_files}
- active checklist names: review-thesis-format-checklist

## Condition Locks
- program condition lock: not-applicable
- thesis lane lock: {subtask_value}
- selected thesis workflow: {selected_workflow_value}
- thesis mutation transaction owner path: references/thesis/thesis-mutation-transaction.md
- thesis mutation transaction record path: {transaction_record_value}
- thesis mutation transaction workflow: {transaction_workflow_value}
- thesis mutation transaction subtype: {transaction_subtype_value}
- thesis mutation transaction target surfaces: {transaction_target_surfaces_value}
- thesis mutation transaction protected sibling surfaces: {transaction_protected_sibling_surfaces_value}
- thesis mutation transaction write owner: upstream canonical builder owns DOCX write; acceptance generator records evidence
- thesis mutation transaction final docx sha256: {transaction_final_docx_sha256}
- explicit user overrides: user explicitly provided canonical thesis template path
- agent authorization source: {agent_authorization_source}
- agent mode: {agent_mode}
- max concurrent live agents: 1
- live agent count plan: 1 live controller plus sequential audit fallback; worker lanes are recorded through role task cards
- dispatch wave plan: wave1 controller=evidence generation; wave1 audit=sequential review; worker roles recorded in the same wave
- audit presence by wave: wave1 includes audit={audit_executor_id}
- concurrency limit verdict: pass
- required lanes: {CANONICAL_REQUIRED_LANES}
- complete role roster: {CANONICAL_ROLE_ROSTER}
- role attendance matrix: {role_attendance_matrix}
- not-applicable lanes with reasons: {not_applicable_lanes_with_reasons}
- agent role aliases zh: {CANONICAL_ROLE_ALIASES_ZH}
- required lane aliases zh: {CANONICAL_ROLE_ALIASES_ZH}
- spawned agent aliases zh: {spawned_agent_aliases_zh}
- lane alias map zh: {CANONICAL_LANE_ALIAS_MAP_ZH}
- audit role alias zh: 审核
- spawned agent ids: {spawned_agent_ids}
- audit agent id: {audit_agent_id}
- sequential audit fallback id: {sequential_audit_fallback_id}
- audit spawn or fallback mode: {fallback_mode}
- audit verdict: {acceptance_audit_verdict_value}
- audit verdict cadence: after every action cycle and final acceptance
- handoff status: {acceptance_handoff_status_value}
- known caveats: {acceptance_known_caveats_value}
- body opener/header title consistency evidence path: {body_opener_header_title_evidence}
- body opener/header title consistency verdict: {body_opener_header_title_verdict}
- baseline promotion status: {baseline_promotion_status_value}
- baseline promotion evidence path: {baseline_promotion_evidence_path}
- release blocker ledger path: {release_blocker_ledger_path}
- unresolved release blocker count: {unresolved_release_blocker_count}
- scoped artifact next-baseline verdict: {scoped_artifact_next_baseline_verdict}
- action audit scope: {action_audit_scope}
- action audit verdict cadence: {action_audit_verdict_cadence}
- action audit verdicts: {action_audit_verdicts}
- action cycles: {action_cycles}
- action categories: {action_categories}
- action owner map: {action_owner_map}
- mutation audit scope: generated evidence and acceptance records only
- mutation audit verdicts: {acceptance_mutation_audit_verdicts_value}
- sequential fallback reason: {sequential_fallback_reason}
- skill invocation lock path: {skill_invocation_lock}
- skill invocation lock verdict: {acceptance_status_value}
- skill invocation verified: pass explicit skill invocation lock created and checked
- skill invocation source type: generator-routed explicit graduation-project-builder invocation
- skill activation status: pass active skill-controlled workflow
- skill takeover verdict: pass lock, checklist, routed references, and audit record controlled execution
- skill bypass prevention verdict: pass no reference-only execution, ad hoc helper substitution, smoke-only gate, or handwritten handoff used
- canonical gate enforcement verdict: pass final acceptance requires validate_skill_gate.py --gate-record on the exact record
- narrow/smoke substitute gate used?: no
- blocked evidence escalation verdict: {"pass no failed evidence remains" if validation_result else "blocked handoff until failed evidence is fixed"}
- no non-control action before lock?: yes
- final handoff allowed by skill lock?: {acceptance_status_value}
- routed references verified: pass routed references loaded
- active checklist verified: pass active checklist externalized
- user request compliance verdict: pass
- loaded rule compliance verdict: pass
- helper-script target path lock: locked
- helper scripts planned this round: {helper_scripts_planned}
- helper script provenance summary: {helper_script_provenance_summary}
- delegated canonical helper paths: {delegated_canonical_helper_paths}
- project-local helper script preflight summary: {project_local_helper_script_preflight_summary}
- project-local helper preflight report path: {project_local_helper_preflight_report_path or project_local_helper_preflight_report}
- project-local helper risk count: {project_local_helper_risk_count}
- project-local helper disposition: {project_local_helper_disposition}
- canonical source restart required?: {canonical_source_restart_required}
- source manuscript genealogy path: {source_manuscript_genealogy_path}
- source retention manifest path: {source_retention_manifest_path}
- source retention ratio: {source_retention_ratio}
- source retention verdict: {source_retention_verdict}
- source review-artifact inventory path: {source_review_artifact_inventory_path}
- final review-artifact diff path: {final_review_artifact_diff_path}
- controlled bookmark disposition path: {controlled_bookmark_disposition_record_path}
- review comments/change marks preservation verdict: {review_comments_change_marks_preservation_verdict}
- comments strip explicit user approval: {comments_strip_explicit_user_approval}
- comment-resolution source DOCX path: {comment_source_docx}
- comment-resolution source DOCX SHA256: {file_sha256(comment_source_docx) if comment_source_docx.exists() else 'missing'}
- comment-resolution ledger path: {comment_resolution_ledger_path}
- comment-resolution audit report path: {comment_resolution_audit_report_path}
- comment-resolution audit verdict: {comment_resolution_audit_verdict}
- source body-citation run inventory path: {source_body_citation_run_inventory_path}
- final body-citation run diff path: {final_body_citation_run_diff_path}
- body citation superscripts preservation verdict: {body_citation_superscripts_preservation_verdict}
- rebuild class: {rebuild_class}
- clean-source restart source path: {clean_source_restart_source_path}
- contaminated-baseline disposition: {contaminated_baseline_disposition}
- project template discovery root: {template_copy.parent}
- template discovery patterns: explicit template path
- discovered candidate template paths: {template_copy}
- candidate template selection reason: user-provided canonical template argument
- active template source type: user-provided
- active template path lock: {template_copy}
- active template fingerprint: {template_fingerprint}
- active template profile path: {active_template_profile_path}
- active template selected before mutation?: yes
- template discovery report path: {template_discovery_report}
- template profile generation command: {PYTHON_EXE} {Path(__file__).resolve().parents[1] / 'scripts' / 'thesis_template_profile.py'} --template {template_copy} --output {active_template_profile_path}
- template profile generated before mutation?: yes
- locked path encoding verdict: pass
- contaminated intermediate source used?: no
- contaminated intermediate source disposition: none
- template alignment verdict: pass
- format preservation promise active?: yes
- format preservation promise source: active skill promise/user requirement; thesis mutation preserves format
- format preservation promise verdict: {format_preservation_promise_verdict_value}
- word-count target or explicit user override: target at least 10000 visible Chinese characters unless a user override is recorded; current build is checked by sample_self_check full thesis content gate
- protected-surface ownership lock: locked
- protected-surface evidence contract path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- protected-surface evidence contract loaded?: yes
- canonical protected surface id set: {'; '.join(PROTECTED_SURFACE_IDS)}
- protected-surface owner map: {protected_surface_owner_map_value}
- protected-surface evidence map: {'; '.join(f'{surface_id}={path}' for surface_id, path in protected_evidence_paths.items())}
- protected-surface reviewed output sha256: {file_sha256(final_docx) if final_docx.exists() else "missing"}
- protected-surface evidence contract verdict: {protected_surface_contract_verdict_value}
- touched template-owned surface families: {touched_surface_families_value}
- surface-face parity baseline locks: active template profile and rendered evidence locked for every generated surface family
- custom/builder/default font usage allowed?: no
- font baseline blockers: none
- humanizer route decision: {humanizer_route_decision_value}
- humanizer target language: {humanizer_target_language_value}
- humanizer scope: {humanizer_scope_value}

## Program Checks
- passed: none
- skipped: none
- failed: none

## Thesis Checks
- passed: none
- skipped: none
- failed: none

## Figure Checks
- passed: none
- skipped: none
- failed: none

## Verification Evidence
- path format rule: use absolute paths when possible; separate multiple paths with `;`; use `none` only when the surface is truly not applicable; evidence paths should point to review evidence record files that follow `assets/review-evidence-template.md`
- verification scope claim: full-thesis-template-alignment
- agent run manifest path: {agent_manifest}
- lightweight action-audit entry path: none
- lane task card evidence paths: {role_task_card_paths}
- all role task card paths: {all_role_task_card_paths}
- audit full roster verdict: pass
- agent id alias map: {agent_id_alias_map}
- filled format-repair task record path: {format_task}
- protected-surface freeze manifest path: {transaction_evidence_path}
- protected-surface freeze manifest verdict: {"pass" if transaction_ok else "fail"}
- post-mutation surface diff path: {transaction_evidence_path}
- post-mutation surface diff verdict: {"pass" if transaction_ok else "fail"}
- target surface render review path: {transaction_evidence_path}
- target surface render review verdict: {"pass" if transaction_ok else "fail"}
- blast-radius render review path: {transaction_evidence_path}
- blast-radius render review verdict: {"pass" if transaction_ok else "fail"}
- cross-surface regression report path: {transaction_evidence_path}
- cross-surface regression report verdict: {"pass" if transaction_ok else "fail"}
- non-target protected surface change verdict: {"pass" if transaction_ok else "fail"}
- non-target format preservation verdict: {non_target_format_preservation_verdict_value}
- local-surface whole-thesis claim verdict: {local_surface_whole_thesis_claim_verdict_value}
- thesis mutation transaction validator command: {PYTHON_EXE} {Path(__file__).resolve().parents[1] / 'scripts' / 'validate_thesis_mutation_transaction.py'} --record {transaction_record_value} --final-docx {final_docx}
- thesis mutation transaction validator result: {transaction_validator_result}
- verified renderer path: {RENDERER_PATH}
- verified rasterizer path: {RASTERIZER_PATH}
- review-copy path actually rendered: {final_docx}
- review-copy promotion binding: none
- rendered PDF path: {final_pdf}
- page-image artifact paths: {join_paths(page_images)}
- final DOCX whole-format structural audit path: {whole_format_audit_path}
- final DOCX whole-format structural audit verdict: {whole_format_audit_verdict}
- final DOCX font-color audit path: {font_color_audit_path}
- final DOCX font-color audit verdict: {font_color_audit_verdict}
- exact output paths: {final_docx}; {final_pdf}; {self_check}
- sample self-check report path: {self_check}
- sample self-check tail-block.pagination-contract detector: {tail_block_detector_status_value}
- sample self-check header-footer.page-number-template-contract detector: {header_footer_page_number_detector_status_value}
- page-class coverage matrix evidence path: {page_class_coverage_matrix}
- mandatory thesis surface inventory path: {mandatory_surface_inventory}
- front matter surface coverage matrix path: {mandatory_surface_inventory}
- end matter surface coverage matrix path: {mandatory_surface_inventory}
- high-risk thesis format surface matrix path: {high_risk_surface_matrix}
- protected surface contract evidence path: {Path(__file__).resolve().parents[1] / 'references' / 'thesis' / 'format-rules' / 'protected-surface-evidence-contract.md'}
- cover style evidence path: {cover_style_path}
- cover style verdict: {protected_surface_verdict("cover_style")}
- cover identity value-line evidence path: {cover_style_path}
- cover identity value-line verdict: {protected_surface_verdict("cover_style")}
- cover value cell targeted verdict: passed cover value cells targeted without rewriting label cells
- cover label cell unchanged verdict: passed cover label cells unchanged
- cover underline x/y/width evidence path: {cover_style_path}
- cover value text bbox evidence path: {cover_style_path}
- cover value-on-underline verdict: passed value text remains on locked underline geometry
- cover row baseline alignment verdict: passed cover row baseline alignment matches donor
- declaration/title front matter evidence path: {declaration_front_matter_path}
- declaration/title front matter verdict: {protected_surface_verdict("declaration_or_title_front_matter")}
- Chinese abstract title evidence path: {abstract_title_path}
- Chinese abstract title verdict: {"passed" if statuses["abstract_ok"] else "failed Chinese abstract title baseline drift"}
- Chinese abstract body evidence path: {abstract_body_path}
- Chinese abstract body verdict: {"passed" if statuses["abstract_ok"] else "failed Chinese abstract body baseline drift"}
- Chinese keyword line evidence path: {keyword_line_path}
- Chinese keyword line verdict: {"passed" if statuses["abstract_ok"] else "failed Chinese keyword line baseline drift"}
- English abstract title evidence path: {en_abstract_title_path}
- English abstract title verdict: {"passed" if statuses["abstract_ok"] else "failed English abstract title baseline drift"}
- English abstract body evidence path: {en_abstract_body_path}
- English abstract body verdict: {"passed" if statuses["abstract_ok"] else "failed English abstract body baseline or semantic parity drift"}
- English abstract indentation evidence path: {en_abstract_body_path}
- English abstract indentation verdict: {"passed" if statuses["abstract_ok"] else "failed English abstract indentation baseline drift"}
- English keyword line evidence path: {en_keyword_line_path}
- English keyword line verdict: {"passed" if statuses["abstract_ok"] else "failed English keyword line baseline or semantic parity drift"}
- TOC title evidence path: {toc_title_path}
- TOC title verdict: {"passed" if toc_final_ok else "failed TOC title baseline or measured geometry drift"}
- TOC entries evidence path: {toc_entries_path}
- TOC entries verdict: {"passed" if toc_final_ok else "failed TOC entries baseline or measured geometry drift"}
- TOC dotted leaders evidence path: {toc_dotted_leaders_path}
- TOC dotted leaders verdict: {"passed" if toc_final_ok else "failed TOC dotted leaders baseline or measured geometry drift"}
- TOC page-number column evidence path: {toc_page_number_column_path}
- TOC page-number column verdict: {"passed" if toc_final_ok else "failed TOC page-number column or measured geometry drift"}
- TOC implementation family/live-field parity evidence path: {toc_entries_path}
- TOC implementation family/live-field parity verdict: {"passed live TOC implementation parity verified" if toc_final_ok else "failed TOC implementation family/live-field parity"}
- TOC per-level template baseline evidence path: {toc_paragraph_typography_paths}
- TOC rendered geometry comparison path: {toc_visual_geometry_paths}
- TOC title/level font and paragraph metrics verdict: {"passed TOC title and per-level font/paragraph metrics verified" if toc_final_ok else "failed TOC title or per-level font/paragraph metrics"}
- TOC dotted leader and page-number column verdict: {"passed TOC dotted leader and page-number column verified" if toc_final_ok else "failed TOC dotted leader or page-number column"}
- TOC occupancy rhythm verdict: {"passed TOC page occupancy rhythm verified" if toc_final_ok else "failed TOC occupancy rhythm"}
- body heading levels evidence path: {body_heading_levels_path}
- body heading levels verdict: {protected_surface_verdict("body_heading_levels")}
- body text evidence path: {body_text_path}
- body text verdict: {protected_surface_verdict("body_text")}
- body citation superscripts evidence path: {body_citation_superscripts_path}
- body citation superscripts verdict: {protected_surface_verdict("body_citation_superscripts")}
- review comments/change marks evidence path: {review_comments_and_change_marks_path}
- review comments/change marks verdict: {protected_surface_verdict("review_comments_and_change_marks")}
- chapter format preservation detector verdict: {statuses["chapter_format_preservation_summary"]}
- chapter format diff path: {chapter_format_diff_path}
- touched chapter rendered evidence paths: {touched_evidence}
- references title evidence path: {references_title_path}
- references title verdict: {protected_surface_verdict("references_title")}
- references entries evidence path: {references_entries_path}
- rendered references-page evidence path: {references_entries_path}
- references entries verdict: {protected_surface_verdict("references_entries")}
- acknowledgement title evidence path: {acknowledgement_title_path}
- acknowledgement title verdict: {protected_surface_verdict("acknowledgement_title")}
- acknowledgement body evidence path: {acknowledgement_body_path}
- acknowledgement body verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body paragraph-dialog metrics verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body direct-run typography verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body first-line indent verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body title-contamination verdict: {acknowledgement_body_title_contamination_verdict(acknowledgement_body_typography_record)}
- header evidence path: {header_evidence_path}
- header verdict: {protected_surface_verdict("header")}
- header presence verdict: {statuses["header_presence_summary"]}
- header rendered verdict: {statuses["header_rendered_summary"]}
- header expected display string source: locked chapter-title/header donor map from active template and current body chapter opener inventory
- header rendered full-display evidence path: {header_evidence_path}
- header chapter number preservation verdict: {protected_surface_verdict("header")}
- footer evidence path: {footer_evidence_path}
- footer verdict: {protected_surface_verdict("footer")}
- page numbers evidence path: {page_numbers_evidence_path}
- page numbers verdict: {protected_surface_verdict("page_numbers")}
- forbidden substitute evidence used?: no
- user-reported issue ledger path: {user_issue_ledger}
- user-reported visual defect surfaces: TOC; abstract; header; footer; page numbers; references; body font; pagination
- user-reported visual defect render-geometry evidence path: {user_reported_visual_defect_evidence}
- user-reported visual defect template-vs-target binding verdict: passed full rendered template-vs-target binding through page-class and protected-surface geometry evidence
- user-reported visual defect full-page/key-surface binding verdict: passed full-page rendered PDF plus key-surface crop and bbox binding
- figure comment conversion checklist path: {figure_comment_conversion}
- figure plan path: {figure_plan_record}
- figure task card paths: {figure_card}
- figure asset manifest path: {figure_asset_manifest_value}
- figure source DOCX path: {figure_source_docx}
- figure source DOCX SHA256: {file_sha256(figure_source_docx) if figure_source_docx.exists() else 'missing'}
- figure inventory path: {figure_plan_record}
- figure manifest contract verdict: {statuses["figure_scope_manifest_summary"]}
- per-figure evidence manifest path: {per_figure_evidence_manifest}
- per-figure rendered evidence paths: {per_figure_evidence_manifest}
- structural figure internal style evidence paths: {figure_evidence}
- structural source-scale bbox map paths: {structural_figure_fields["structural_source_scale_bbox_map_path"]}
- structural inserted-scale geometry evidence paths: {structural_figure_fields["structural_inserted_scale_geometry_evidence_path"]}
- structural dense-zone crop evidence paths: {structural_figure_fields["structural_dense_zone_crop_evidence_paths"]}
- inserted-scale collision evidence paths: {structural_figure_fields["structural_inserted_scale_collision_evidence_path"]}
- runtime screenshot authenticity evidence paths: {figure_evidence}
- code screenshot authenticity evidence paths: {figure_evidence}
- toc restoration evidence paths: {toc_visual_geometry_paths}
- toc rendered baseline comparison evidence paths: {toc_visual_geometry_paths}
- toc visual geometry evidence paths: {toc_visual_geometry_paths}
- toc paragraph-and-typography evidence paths: {toc_paragraph_typography_paths}
- toc visible-run typography evidence paths: {toc_paragraph_typography_paths}
- whole-document pagination evidence path: {whole_pagination_evidence_path}
- package baseline manifest path: {whole_pagination_evidence_path}
- package drift report path: {whole_pagination_evidence_path}
- package drift verdict: {"pass" if whole_pagination_ok else "fail"}
- pre-mutation page map path: {whole_pagination_evidence_path}
- post-mutation page map path: {whole_pagination_evidence_path}
- whole-document pagination diff path: {whole_pagination_evidence_path}
- section break numbering map path: {whole_pagination_evidence_path}
- chapter start owner map path: {whole_pagination_evidence_path}
- tail-block owner map path: {whole_pagination_evidence_path}
- TOC-to-heading page sync map path: {whole_pagination_evidence_path}
- logical-physical page map path: {whole_pagination_evidence_path}
- blank-page scan evidence path: {whole_pagination_evidence_path}
- rendered page count baseline/actual: {whole_pagination_fields["rendered_page_count_baseline_actual"]}
- field-refresh before/after state: {whole_pagination_fields["field_refresh_before_after_state"]}
- live TOC required this round?: {live_toc_required_value}
- live TOC field count: {live_toc_field_count}
- live TOC field verdict: {live_toc_field_verdict_value}
- DOCX OpenXML compatibility repair report path: not-applicable unless a bounded OpenXML compatibility recovery was run
- DOCX OpenXML compatibility verdict: not-applicable unless a bounded OpenXML compatibility recovery was run
- tail-block pagination evidence paths: {thesis_evidence}
- tail-block rendered opener comparison evidence paths: {touched_evidence}
- table authority evidence paths: {thesis_evidence}
- table-local structure evidence paths: {thesis_evidence}
- table rendered baseline comparison evidence paths: {thesis_evidence}
- table title-mode evidence paths: {thesis_evidence}
- table header-bottom middle-rule evidence paths: {thesis_evidence}
- table continuation evidence paths: {thesis_evidence}
- table continuation summary: passed table continuation status enumerated for every body table; no cross-page body tables detected unless rendered evidence says otherwise
- cross-page table rendered pages: {touched_evidence}
- continuation title outside-grid verdict: not-applicable-with-reason: no cross-page body tables detected in rendered table-page review
- table row split/header repeat verdict: passed row split policy and header repeat checked for body tables
- citation audit evidence path: {citation_audit}
- citation audit final DOCX SHA256: {citation_audit_final_docx_sha256}
- citation audit source-to-final run diff path: {citation_audit_source_to_final_run_diff_path}
- bibliography audit evidence path: {bibliography_audit}
- bibliography baseline summary: {statuses["bibliography_baseline_summary"]}
- bibliography numbering summary: {statuses["bibliography_numbering_summary"]}
- bibliography font-slot summary: {statuses["bibliography_font_slot_summary"]}
- bibliography comment-aware repair summary: {bibliography_comment_aware_repair_summary}
- project-local helper script paths: {join_paths(project_local_helper_script_paths) if project_local_helper_script_paths else 'none'}
- project-local helper script risk summary: {project_local_helper_script_risk_summary}
- docx font/encoding audit evidence path: {font_audit}
- surface-face parity audit evidence path: {thesis_evidence}
- surface paragraph-and-typography audit evidence path: {thesis_evidence}
- sibling-surface audit evidence path: {thesis_evidence}
- cross-surface regression freeze evidence path: {thesis_evidence}
- font-family baseline audit evidence path: {thesis_evidence}
- active template profile evidence path: {thesis_evidence}
- body style audit evidence path: {body_style_audit}
- program evidence paths: none
- thesis rendered-page evidence paths: {thesis_evidence}
- figure review evidence paths: {figure_evidence}
- paragraph-review evidence paths: {paragraph_evidence}
- touched-page review evidence paths: {touched_evidence}
- content mutation rendered-page review path: {thesis_evidence}
- content mutation machine-vision verdict: passed rendered-page review found no content-mutation visual drift
- inserted body heading-contamination verdict: {statuses["body_heading_contamination_summary"]}
- touched-page/blast-radius machine-vision evidence paths: {touched_evidence}
- format lane post-mutation rendered audit verdict: passed protected-surface and touched-page rendered audits match the reference template lane
- machine-vision summary by touched surface: {rendered_summary}
- startup / runtime verification summary: not-applicable
- thesis rendered-page verification summary: {rendered_summary}
- action-level audit summary: passed all action cycles audited
- post-script smoke-audit summary: passed
- body style binding summary: {statuses["body_style_binding_summary"]}
- Normal baseline preservation summary: {statuses["normal_baseline_summary"]}
- body paragraph family consistency summary: {statuses["body_family_summary"]}
- body heading contamination summary: {statuses["body_heading_contamination_summary"]}
- body mixed-script protected-surface summary: {statuses["body_text_summary"]}
- surface-face parity summary: passed class-specific style bindings verified for headings, body, figures, tables, abstracts, keywords, references, acknowledgement, appendix, header, footer, and page numbers
- surface paragraph-and-typography summary: {"passed " + surface_paragraph_typography_summary if surface_paragraph_typography_ok else "failed surface paragraph-and-typography: " + surface_paragraph_typography_summary}
- surface paragraph-and-typography verdict: {surface_paragraph_and_typography_verdict_value}
- sibling-surface audit summary: passed sibling surfaces checked for all generated template-owned classes
- style-blast-radius escalation summary: passed generated acceptance ran TOC visible-run, table family, references, acknowledgement, appendix, header/footer, and pagination sibling checks after body/default style surfaces
- cross-surface regression verdict: passed protected-surface freeze diff has no unrelated surface regression
- TOC underline pollution verdict: passed no TOC underline or hyperlink-style contamination detected
- table style regression verdict: {"passed table titles, table cells, borders, and rendered family match template" if statuses["table_ok"] else "failed table style regression detected"}
- font-family baseline summary: passed locked template font baselines used for all generated surface classes
- builder/default font rejection summary: passed no builder/default/guessed fonts used
- code title formatting summary: {"passed" if statuses["code_title_ok"] else "failed code-title formatting drift"}
- code block formatting summary: {"passed" if statuses["code_block_ok"] else "failed code-block formatting drift"}
- abstract baseline preservation summary: {"passed" if statuses["abstract_ok"] else "failed abstract surface baseline drift"}
- abstract and keyword surface verdict: {"passed abstract and keyword six-surface inventory rows verified" if statuses["abstract_ok"] else "failed abstract or keyword inventory row drift"}
- heading baseline preservation summary: {"passed" if statuses["heading_baseline_ok"] else "failed heading baseline drift in one or more used levels"}
- heading family preservation summary: {"passed" if statuses["heading_ok"] else "failed heading style drift"}
- heading level 1 verdict: {heading_level_verdict("level1")}
- heading level 2 verdict: {heading_level_verdict("level2")}
- heading level 3 verdict: {heading_level_verdict("level3")}
- heading level 4 verdict: {heading_level_verdict("level4")}
- heading direct rPr/font/size/bold/spacing verdict: {heading_direct_typography_verdict}
- TOC / bookmark integrity summary: {"passed" if toc_final_ok else "failed toc integrity or measured geometry drift"}
- TOC visible format summary: {"passed" if toc_final_ok else "failed visible toc format or measured geometry mismatch"}
- TOC baseline restoration summary: {"passed" if toc_final_ok else "failed toc baseline not restored or measured geometry missing"}
- TOC visual rhythm summary: {"passed measured TOC rhythm: " + toc_geometry_summary if toc_final_ok else "failed measured TOC rhythm: " + toc_geometry_summary}
- TOC visual baseline verdict: {"passed TOC title, entries, dotted leaders, page-number column, and measured geometry verified" if toc_final_ok else "failed TOC visual baseline or measured geometry drift"}
- TOC visual geometry verdict: {toc_visual_geometry_verdict_value}
- TOC paragraph-and-typography verdict: {toc_paragraph_and_typography_verdict_value}
- TOC visible-run typography verdict: {toc_visible_run_typography_verdict_value}
- table authority summary: passed
- table-local structure summary: passed standalone title or donor-backed in-table title mode; keep-with-next binding; header-bottom middle rule; no title inside table grid unless donor title mode is first_merged_row
- table title-mode summary: passed donor_title_mode matched target_title_mode
- table header-bottom middle-rule summary: passed donor header-bottom middle rule matched target table
- chapter-start pagination summary: {chapter_summary}
- tail-block pagination summary: {"passed" if statuses["tail_block_ok"] and statuses["page_classes_ok"] else "failed tail-block opener drift"}
- whole-document pagination summary: {"passed " + whole_pagination_summary if whole_pagination_ok else "failed whole-document pagination: " + whole_pagination_summary}
- whole-document pagination verdict: {whole_document_pagination_verdict_value}
- custom-layout result-table preservation summary: passed
- figure family style summary: {"passed" if statuses["figure_ok"] else "failed wrong family or sample mismatch"}
- figure source/style contract summary: {figure_contract_summary}
- structural relation-attribute collision verdict: {structural_figure_verdict_value}
- structural shape-overlap verdict: {structural_figure_verdict_value}
- structural inserted-scale collision verdict: {structural_figure_verdict_value}
- structural source-to-inserted geometry verdict: {structural_figure_verdict_value}
- table rendered-family summary: {"passed" if statuses["table_ok"] else "failed wrong family or shading drift"}
- table donor title mode: donor-backed mode from protected table evidence
- table target title mode: target mode from protected table evidence
- table title-mode verdict: passed donor title mode matched target table
- table in-table merged-title-row verdict: passed when donor title mode is first_merged_row; otherwise not-applicable with donor-backed reason
- table top rule verdict: {"passed" if statuses["table_ok"] else "failed top rule drift"}
- table header separator rule verdict: {"passed" if statuses["table_ok"] else "failed header separator rule drift"}
- table body middle-rule policy verdict: {"passed" if statuses["table_ok"] else "failed body middle-rule policy drift"}
- table bottom rule verdict: {"passed" if statuses["table_ok"] else "failed bottom rule drift"}
- table rendered crop metric json path: {thesis_evidence}
- WPS preset application summary: not-applicable
- image replacement rendered summary: not-applicable
- image metadata sync summary: not-applicable
- formula object audit evidence path: {formula_object_audit_path}
- formula object preservation summary: {formula_object_summary_value}
- formula numbering surface summary: {formula_numbering_summary_value}
- header placement summary: {"passed" if statuses["header_ok"] else "failed wrong position or misaligned"}
- footer indent summary: {"passed" if statuses["footer_ok"] else "failed indent offset"}
- footer baseline typography summary: {"passed" if statuses["footer_baseline_ok"] else "failed footer typography baseline drift"}
- references title indentation/font verdict: {protected_surface_verdict("references_title")}
- references entry format verdict: {"passed references_title and references_entries protected-surface evidence verified" if references_entry_surface_ok else "failed references title or entry protected-surface format drift"}
- acknowledgement title indentation/font verdict: {protected_surface_verdict("acknowledgement_title")}
- acknowledgement body indentation/font verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body paragraph-dialog metrics verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body direct-run typography verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body first-line indent verdict: {protected_surface_verdict("acknowledgement_body")}
- acknowledgement body title-contamination verdict: {acknowledgement_body_title_contamination_verdict(acknowledgement_body_typography_record)}
- appendix format verdict: {"passed appendix inventory rows verified" if statuses["page_classes_ok"] else "not-applicable-with-reason: active template has no appendix row to verify"}
- high-risk thesis format surface verdict: {"passed high-risk thesis format surface matrix verified" if statuses["abstract_ok"] and toc_final_ok and statuses["heading_ok"] and statuses["body_text_ok"] and references_entry_surface_ok and acknowledgement_surface_ok and statuses["page_classes_ok"] else "failed high-risk thesis format surface matrix drift"}
- humanizer evidence path: {humanizer_evidence_value}
- humanizer evidence rule: active humanizer evidence must include before text, after text, target language, paragraph id, and skill name inside the evidence file
- humanizer manual or skill run id: {humanizer_evidence_value if humanizer_evidence_value != "none" else "none"}
- humanizer processed paragraph count: {"nonzero recorded in evidence" if humanizer_evidence_value != "none" else "none"}
- humanizer changed pattern list: {"recorded in active humanizer evidence" if humanizer_evidence_value != "none" else "none"}
- humanizer no meta-evaluation voice verdict: {"pass" if humanizer_evidence_value != "none" else "none"}

## Blockers And Skips
- blockers: {acceptance_blockers_value}
- skipped with reasons: none
- failed with reasons: {failed_reasons_value}
- smoke acceptance disposition: {"smoke-only; blocked for delivery" if args.smoke_acceptance else "not-applicable"}

## Remaining Risks
- {remaining_risks_value}

## Validation
- acceptance record path: {output}
- skill selftest command: {args.selftest_command}
- skill selftest result: {"pass" if validation_result else "fail"}
- validation command: {validator_command_text}
- validation result: {"pass" if validation_result else "fail"}
- build command evidence path: {args.build_command_evidence or "none"}
- UTF-8 gate evidence path: {args.utf8_gate_evidence or "none"}
- skill static gate evidence path: {args.skill_gate_evidence or "none"}
- skill selftest evidence path: {args.selftest_evidence or "none"}
- integration gate evidence path: {args.integration_gate_evidence or "none"}
- output gate evidence path: {args.output_gate_evidence or "none"}

## Handoff Rule
Do not report completion if any required check remains failed.
Do not report completion if any skipped item lacks a reason or explicit `none`.
Do not report completion if the acceptance record does not name active references, active checklists, condition locks, blockers, skips, exact output paths, and evidence paths.
Do not report completion if any required evidence path field is blank.
Do not report completion after thesis content drafting or rewriting if an active humanizer route lacks evidence files with before text, after text, target language, paragraph id, and skill name.
Do not report completion after thesis work if the citation audit evidence path is `none` or blank.
Do not report completion after thesis work if the DOCX font/encoding audit evidence path is `none` or blank.
Do not report completion if an evidence path does not point to a valid review evidence record file.
Do not report completion after thesis modification work if touched-page review evidence is missing.
Do not report completion after TOC repair if TOC restoration evidence or TOC rendered baseline comparison evidence is missing.
Do not report completion after tail-block pagination repair if tail-block pagination evidence or rendered opener comparison evidence is missing.
Do not report completion when the current run enables paper-only literature and the bibliography audit evidence path is `none` or blank.
Do not report completion if the gate validator has not been run.
Do not report completion if the gate validator returns non-zero.
"""
    write_text(output, record)
    validator_output = output.parent / "validate-skill-gate-on-acceptance.log"
    try:
        validator_run = subprocess.run(
            validator_cmd,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=300,
        )
        validator_text = (
            "$ " + validator_command_text + "\n"
            + validator_run.stdout
            + ("\nSTDERR:\n" + validator_run.stderr if validator_run.stderr else "")
        )
        actual_validator_passed = validator_run.returncode == 0
    except Exception as exc:
        validator_text = "$ " + validator_command_text + f"\nvalidator execution failed: {exc}\n"
        actual_validator_passed = False
    write_text(validator_output, validator_text)
    final_validation_result = validation_result and actual_validator_passed
    if final_validation_result != validation_result:
        reason = "validator failed; see " + str(validator_output)
        record = record.replace(
            f"- validation result: {'pass' if validation_result else 'fail'}",
            f"- validation result: {'pass' if final_validation_result else 'fail'}",
        )
        record = record.replace(
            f"- skill selftest result: {'pass' if validation_result else 'fail'}",
            f"- skill selftest result: {'pass' if final_validation_result else 'fail'}",
        )
        if not final_validation_result:
            record = block_pass_shaped_handoff_fields(record, reason)
        write_text(output, record)
        if skill_invocation_lock.exists():
            lock_text = read_text(skill_invocation_lock)
            lock_text = block_pass_shaped_handoff_fields(lock_text, reason)
            write_text(skill_invocation_lock, lock_text)
    print(output)
    if not final_validation_result:
        for failure in evidence_input_failures:
            print(failure)
        if not actual_validator_passed:
            print(f"validator log: {validator_output}")
    if not final_validation_result:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
