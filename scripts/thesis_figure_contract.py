#!/usr/bin/env python3
"""Canonical thesis figure manifest and validation helpers."""

from __future__ import annotations

import argparse
import base64
from collections import Counter
import copy
import hashlib
import json
import os
import posixpath
import re
import shutil
import urllib.parse
import zipfile
import zlib
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


ASSET_SCHEMA = "graduation-project-builder.figure-manifest.v2"
SEQUENCE_TOKENS = (
    "\u6d41\u7a0b",
    "\u94fe\u8def",
    "\u6b65\u9aa4",
    "\u8fc7\u7a0b",
    "\u5904\u7406",
    "\u987a\u5e8f",
    "\u9636\u6bb5",
    "pipeline",
    "workflow",
    "flowchart",
    "process",
)
STRUCTURE_TOKENS = (
    "\u7ed3\u6784",
    "\u67b6\u6784",
    "\u6a21\u5757",
    "\u529f\u80fd",
    "\u5c42\u6b21",
    "\u5b9e\u4f53",
    "\u7528\u4f8b",
    "\u65f6\u5e8f",
    "architecture",
    "module",
    "entity",
    "use case",
    "sequence",
)
STRUCTURAL_FAMILIES = {"structure", "flowchart", "architecture", "er", "use-case", "sequence", "module-tree"}
MECHANICAL_CAD_FAMILIES = {
    "mechanical-cad",
    "mechanical_cad",
    "mechanical_cad_sheet",
    "cad-sheet",
    "cad_sheet",
    "cad_sheet_png",
    "cad-png",
    "cad_png",
    "verified-cad-png",
    "verified_cad_png",
}
RUNTIME_SCREENSHOT_FAMILIES = {"runtime", "runtime-screenshot", "screenshot", "ui-screenshot", "page-screenshot"}
STRUCTURAL_FORBIDDEN_SOURCE_PATTERNS = (
    "mermaid",
    "pillow",
    "pil",
    "image draw",
    "imagedraw",
    "manual png",
    "manual_png",
    "manual-png",
    "hand drawn png",
    "hand-drawn png",
    "drawn png",
    "\u624b\u7ed8png",
    "\u624b\u7ed8 png",
    "\u624b\u5de5png",
    "\u624b\u5de5 png",
)
RUNTIME_SCREENSHOT_TOKENS = (
    "screenshot",
    "screen shot",
    "runtime",
    "system page",
    "system UI",
    "login page",
    "dashboard",
    "\u622a\u56fe",
    "\u754c\u9762",
    "\u9875\u9762",
    "\u9996\u9875",
    "\u767b\u5f55",
    "\u6ce8\u518c",
    "\u540e\u53f0",
    "\u524d\u53f0",
    "\u7cfb\u7edf\u8fd0\u884c",
    "\u8fd0\u884c\u6548\u679c",
    "\u7ba1\u7406\u7aef",
    "\u7528\u6237\u7aef",
    "\u4e3b\u754c\u9762",
    "\u7ba1\u7406\u754c\u9762",
    "\u7ed3\u679c\u622a\u56fe",
    "\u6570\u636e\u5b58\u50a8\u622a\u56fe",
    "\u5904\u7406\u7ed3\u679c\u622a\u56fe",
    "\u6e05\u6d17\u7ed3\u679c\u622a\u56fe",
    "\u8bad\u7ec3\u7ed3\u679c\u622a\u56fe",
    "\u6d4b\u8bd5\u7ed3\u679c\u622a\u56fe",
    "\u6a21\u578b\u7ed3\u679c\u622a\u56fe",
)
ALGORITHM_RESULT_FAMILIES = {
    "algorithm_result",
    "algorithm-result",
    "algorithm result",
    "detection-result",
    "recognition-result",
    "ocr-result",
    "model-output",
}
ALGORITHM_RESULT_TOKENS = (
    "algorithm result",
    "algorithm-result",
    "detection result",
    "recognition result",
    "model output",
    "yolov8",
    "yolo",
    "dbnet",
    "crnn",
    "ocr",
    "drug_box",
    "drug_name",
    "expiry_date",
    "\u7b97\u6cd5\u7ed3\u679c",
    "\u68c0\u6d4b\u6548\u679c",
    "\u8bc6\u522b\u7ed3\u679c",
    "\u68c0\u6d4b\u6846",
    "\u6587\u672c\u68c0\u6d4b",
    "\u6587\u5b57\u8bc6\u522b",
    "\u7ed3\u679c\u56fe",
    "\u8fd0\u884c\u7ed3\u679c",
    "\u6d4b\u8bd5\u7ed3\u679c",
    "\u9884\u6d4b\u7ed3\u679c",
    "\u5206\u7c7b\u7ed3\u679c",
    "\u8bc6\u522b\u6548\u679c",
    "\u6a21\u578b\u8bad\u7ec3\u7ed3\u679c",
    "\u8bad\u7ec3\u4e0e\u6d4b\u8bd5\u7ed3\u679c",
    "\u8bad\u7ec3\u7ed3\u679c",
    "\u6570\u636e\u5904\u7406\u7ed3\u679c",
    "\u6570\u636e\u6e05\u6d17\u7ed3\u679c",
    "\u6570\u636e\u5b58\u50a8\u7ed3\u679c",
    "\u9884\u5904\u7406\u7ed3\u679c",
)
SYNTHETIC_RESULT_TOKENS = (
    "schematic",
    "mock",
    "mockup",
    "placeholder",
    "synthetic",
    "fake",
    "sample-only",
    "\u793a\u610f\u56fe",
    "\u6837\u4f8b\u56fe",
    "\u5360\u4f4d",
    "\u624b\u7ed8",
    "\u81ea\u5b9a\u4e49\u7ed8\u5236",
)
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
MIN_RUNTIME_SCREENSHOT_WIDTH = 700
MIN_RUNTIME_SCREENSHOT_HEIGHT = 450
MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE = 0.90
MAX_BLANK_IMAGE_DOMINANT_RATIO = 0.985
MAX_SOLID_BLOCK_DOMINANT_RATIO = 0.92
MIN_IMAGE_VARIANCE = 2.0
PURPLE_PLACEHOLDER_DOMINANT_RATIO = 0.45
ER_TOKENS = (
    "er\u56fe",
    "e-r",
    "er diagram",
    "entity relationship",
    "\u5b9e\u4f53\u5173\u7cfb",
    "\u6570\u636e\u5e93er",
    "\u6570\u636e\u5e93\u5b9e\u4f53",
)
ER_RELATION_ATTRIBUTE_MIN_CLEARANCE = 24.0
ALLOWED_WHITE = {"#ffffff", "#fff", "ffffff", "fff", "white", ""}
ALLOWED_BLACK = {"#000000", "#111111", "#111827", "#222222", "#333333", "000000", "111111", "111827", "222222", "333333", "black", ""}
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
CJK_NUMERAL_CLASS = r"\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341"
FIGURE_CAPTION_RE = re.compile(
    rf"^\s*(?:\u9644\s*\u56fe|\u56fe|figure|fig\.)\s*"
    rf"(?P<number>(?:\d+|[A-Za-z]+|[{CJK_NUMERAL_CLASS}]+)"
    rf"(?:[\-.\u2010-\u2015\uff0d\uff0e](?:\d+|[A-Za-z]+|[{CJK_NUMERAL_CLASS}]+))*)"
    rf"(?!(?:[\-.\u2010-\u2015\uff0d\uff0e]\s*(?:\d+|[A-Za-z]+|[{CJK_NUMERAL_CLASS}]+)))"
    # A real caption has a separator after the figure number. Explanatory
    # prose such as "图3-1展示..." is a figure reference, not a caption.
    rf"[\s\u3000:：、.\-\u2010-\u2015\uff0d\uff0e]+(?P<title>.*)$",
    re.I,
)
FIGURE_CAPTION_TITLE_STRIP_RE = re.compile(r"^[\s\u3000:：.\-\u2010-\u2015\uff0d\uff0e、_]+")
BODY_HEADING_RE = re.compile(r"^\s*(?:\d{1,2}\s+\S|\u7b2c[0-9\u4e00-\u9fff]+\u7ae0)")
FORMULA_LABEL_TEXT_RE = re.compile(
    rf"^\s*(?:\u5f0f\s*)?[\(\uff08]\s*\d+(?:[-.\uff0d\uff0e]\d+)+\s*[\)\uff09]\s*$"
    rf"|^\s*\u5f0f\s*[\(\uff08]\s*\d+(?:[-.\uff0d\uff0e]\d+)+\s*[\)\uff09]\s*$",
    re.I,
)
FORMULA_CONTEXT_TOKENS = (
    "\u5f0f(",
    "\u5f0f\uff08",
    "\u8ba1\u7b97",
    "\u5706\u5468\u529b",
    "\u5f84\u5411\u529b",
    "\u8f74\u5411\u529b",
    "\u5f2f\u77e9",
    "\u6c34\u5e73\u9762",
    "\u5782\u76f4\u9762",
)
STRUCTURAL_DOCX_TOKENS = tuple(dict.fromkeys(SEQUENCE_TOKENS + STRUCTURE_TOKENS + (
    "\u6d41\u7a0b\u56fe",
    "er diagram",
    "er\u56fe",
    "e-r",
    "\u6570\u636e\u5e93",
    "\u5b9e\u4f53\u5173\u7cfb",
    "\u67b6\u6784\u56fe",
    "\u7528\u4f8b\u56fe",
    "\u65f6\u5e8f\u56fe",
    "diagram",
    "draw.io",
    "svg",
)))
REQUIRED_FIGURE_RULES = (
    {
        "id": "database_er_diagram",
        "family": "er",
        "caption_tokens": ("er", "e-r", "\u5b9e\u4f53\u5173\u7cfb", "\u6982\u5ff5\u6a21\u578b", "\u6570\u636e\u5e93"),
        "trigger_any": ("\u6570\u636e\u5e93\u6982\u5ff5\u8bbe\u8ba1", "\u5b9e\u4f53\u95f4\u8054\u7cfb", "\u5b9e\u4f53\u4e4b\u95f4\u7684\u8054\u7cfb"),
        "trigger_all": ("\u6570\u636e\u5e93\u8bbe\u8ba1", "\u4e3b\u8981\u5b9e\u4f53"),
        "reason": "database concept design/entity relationship prose requires an ER diagram",
    },
)


def infer_figure_family(figure: dict[str, Any]) -> str:
    declared = str(figure.get("family") or figure.get("diagram_family") or "").strip().lower()
    combined = " ".join(
        str(figure.get(key) or "")
        for key in ("caption", "description", "followup", "title", "alt", "purpose")
    ).lower()
    source_text = " ".join(
        str(figure.get(key) or "")
        for key in ("source_kind", "final_source_kind", "authoring_tool", "asset_type", "kind")
    ).lower()
    if declared in MECHANICAL_CAD_FAMILIES or any(token in source_text for token in MECHANICAL_CAD_FAMILIES):
        return declared or "mechanical-cad"
    if declared in ALGORITHM_RESULT_FAMILIES or any(token.lower() in combined for token in ALGORITHM_RESULT_TOKENS):
        return "algorithm_result"
    if declared == "er" or any(token.lower() in combined for token in ER_TOKENS):
        return "er"
    if any(token.lower() in combined for token in SEQUENCE_TOKENS):
        return "flowchart"
    if declared:
        return declared
    if any(token.lower() in combined for token in STRUCTURE_TOKENS):
        return "structure"
    if figure.get("drawio_path") or figure.get("drawio") or figure.get("svg_path") or figure.get("svg"):
        return "structure"
    return "raster"


def is_structural_figure(family: str, figure: dict[str, Any]) -> bool:
    if family in STRUCTURAL_FAMILIES:
        return True
    return bool(figure.get("drawio_path") or figure.get("drawio") or figure.get("svg_path") or figure.get("svg"))


def is_runtime_screenshot_figure(family: str, figure: dict[str, Any]) -> bool:
    declared = str(
        figure.get("source_kind")
        or figure.get("kind")
        or figure.get("figure_type")
        or figure.get("asset_type")
        or ""
    ).strip().lower()
    combined = " ".join(
        str(figure.get(key) or "")
        for key in ("caption", "description", "followup", "title", "alt", "purpose")
    ).lower()
    return (
        family in RUNTIME_SCREENSHOT_FAMILIES
        or declared in RUNTIME_SCREENSHOT_FAMILIES
        or "runtime screenshot" in combined
        or "system screenshot" in combined
        or any(token.lower() in combined for token in RUNTIME_SCREENSHOT_TOKENS)
    )


def is_algorithm_result_figure(family: str, figure: dict[str, Any]) -> bool:
    declared = str(
        figure.get("source_kind")
        or figure.get("kind")
        or figure.get("figure_type")
        or figure.get("asset_type")
        or ""
    ).strip().lower()
    combined = " ".join(
        str(figure.get(key) or "")
        for key in ("caption", "description", "followup", "title", "alt", "purpose")
    ).lower()
    return family in ALGORITHM_RESULT_FAMILIES or declared in ALGORITHM_RESULT_FAMILIES or any(
        token.lower() in combined for token in ALGORITHM_RESULT_TOKENS
    )


def path_text(value: object) -> str:
    return str(Path(str(value)).resolve()) if value else ""


def resolve_manifest_path_value(value: object, manifest_path: Path | None = None) -> Path | None:
    text = str(value or "").strip()
    if not text or text.lower() in {"none", "not-applicable", "null"}:
        return None
    path = Path(text)
    if not path.is_absolute() and manifest_path is not None:
        path = manifest_path.resolve().parent / path
    return path.resolve()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


MANIFEST_PATH_FIELDS = {
    "path",
    "image_path",
    "png",
    "raster_fallback",
    "drawio",
    "drawio_path",
    "svg",
    "svg_path",
    "accepted_screenshot_path",
    "accepted_result_image_path",
    "result_image_path",
    "source_image_path",
    "input_image_path",
    "generation_script_path",
    "model_output_log_path",
    "task_card",
    "task_card_path",
    "figure_task_card",
    "post_insertion_rendered_evidence",
    "post_insertion_rendered_evidence_path",
    "rendered_page_evidence",
    "rendered_evidence",
    "final_docx_relationship_evidence",
    "docx_relationship_evidence",
    "relationship_evidence",
    "template_sample_evidence_path",
    "skill_internal_fallback_sample_path",
    "stored_fallback_sample_path",
    "stored_skill_sample_path",
    "sample_lock_evidence_path",
    "geometry_validation_report",
    "geometry_report",
    "source_geometry_report",
    "source_scale_bbox_map",
    "inserted_scale_geometry_evidence",
    "dense_zone_crop_evidence",
    "source_docx_path",
    "original_docx_path",
    "source_manuscript_path",
    "final_docx_path",
    "output_docx_path",
    "reviewed_output_path",
    "final_manuscript_path",
    "table_authority_source_path",
    "authority_source_file_path",
    "active_table_authority_path",
    "table_authority_manuscript_binding_proof",
    "table_authority_binding_evidence",
    "table_to_manuscript_binding_evidence",
    "rendered_table_evidence",
    "rendered_table_evidence_path",
    "rendered_table_page_evidence",
    "table_rendered_baseline_comparison_evidence",
    "table_rendered_baseline_comparison_evidence_path",
    "final_docx_table_evidence",
    "final_docx_table_evidence_path",
    "table_final_docx_binding_evidence",
    "table_audit_evidence_path",
}


def _resolve_manifest_path_string(value: str, manifest_path: Path | None) -> str:
    if manifest_path is None:
        return value
    text = value.strip()
    if not text or text.lower() in {"none", "n/a", "not-applicable", "missing"}:
        return value
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", text):
        return value
    path = Path(text)
    if path.is_absolute():
        return value
    return str((manifest_path.resolve().parent / path).resolve())


def manifest_with_resolved_paths(manifest: dict[str, Any], manifest_path: Path | None) -> dict[str, Any]:
    if manifest_path is None:
        return manifest
    resolved = copy.deepcopy(manifest)

    def visit(node: Any) -> None:
        if isinstance(node, dict):
            for key, value in list(node.items()):
                if key in MANIFEST_PATH_FIELDS and isinstance(value, str):
                    node[key] = _resolve_manifest_path_string(value, manifest_path)
                else:
                    visit(value)
        elif isinstance(node, list):
            for item in node:
                visit(item)

    visit(resolved)
    return resolved


def iter_figures(content: dict[str, Any]) -> list[tuple[dict[str, Any], str]]:
    found: list[tuple[dict[str, Any], str]] = []

    def visit_node(node: Any, chapter_title: str) -> None:
        if not isinstance(node, dict):
            return
        for figure in node.get("figures", []) or []:
            if isinstance(figure, dict):
                found.append((figure, chapter_title))
        for section in node.get("sections", []) or []:
            visit_node(section, chapter_title)

    for chapter in content.get("chapters", []) or []:
        visit_node(chapter, str(chapter.get("title") or ""))
    appendix = content.get("appendix") or content.get("appendices")
    if isinstance(appendix, dict):
        visit_node(appendix, str(appendix.get("title") or "附录"))
    elif isinstance(appendix, list):
        for index, item in enumerate(appendix, start=1):
            if isinstance(item, dict):
                visit_node(item, str(item.get("title") or f"附录{index}"))
    return found


def build_figure_asset_manifest(content: dict[str, Any], run_root: Path) -> dict[str, Any]:
    manifest: dict[str, Any] = {"schema": ASSET_SCHEMA, "figures": {}, "tables": {}, "diagrams": {}}
    fig_idx = 1
    table_idx = 1
    for figure, chapter_title in iter_figures(content):
        inferred = infer_figure_family(figure)
        declared = str(figure.get("family") or figure.get("diagram_family") or "").strip().lower()
        image_path = figure.get("image_path") or figure.get("path") or figure.get("png") or ""
        drawio_path = figure.get("drawio_path") or figure.get("drawio") or ""
        svg_path = figure.get("svg_path") or figure.get("svg") or ""
        geometry_report = figure.get("geometry_validation_report") or figure.get("geometry_report") or ""
        source_bbox_map = figure.get("source_scale_bbox_map") or ""
        inserted_geometry_evidence = figure.get("inserted_scale_geometry_evidence") or ""
        dense_zone_crop_evidence = figure.get("dense_zone_crop_evidence") or figure.get("dense_zone_crop_evidence_paths") or ""
        collision_verdict = str(figure.get("collision_check_verdict") or "pending")
        task_card = figure.get("task_card") or figure.get("task_card_path") or figure.get("figure_task_card") or ""
        post_insertion_rendered_evidence = (
            figure.get("post_insertion_rendered_evidence")
            or figure.get("post_insertion_rendered_evidence_path")
            or figure.get("rendered_page_evidence")
            or figure.get("rendered_evidence")
            or ""
        )
        final_docx_relationship_evidence = (
            figure.get("final_docx_relationship_evidence")
            or figure.get("docx_relationship_evidence")
            or figure.get("relationship_evidence")
            or ""
        )
        rendered_page_status = str(
            figure.get("rendered_page_status")
            or figure.get("post_insertion_rendered_verdict")
            or figure.get("rendered_verdict")
            or "pending"
        )
        insertion_status = str(figure.get("insertion_status") or figure.get("final_insertion_status") or "pending")
        relation_attribute_collision_verdict = str(
            figure.get("relation_attribute_collision_verdict")
            or figure.get("er_relation_attribute_clearance_verdict")
            or collision_verdict
        )
        shape_overlap_verdict = str(
            figure.get("shape_overlap_verdict")
            or figure.get("er_overlap_scan_verdict")
            or collision_verdict
        )
        inserted_scale_collision_verdict = str(
            figure.get("inserted_scale_collision_verdict")
            or figure.get("inserted_scale_geometry_verdict")
            or collision_verdict
        )
        source_to_inserted_geometry_verdict = str(
            figure.get("source_to_inserted_geometry_verdict")
            or figure.get("final_source_to_inserted_geometry_verdict")
            or inserted_scale_collision_verdict
        )
        if is_structural_figure(inferred, figure):
            source_kind = "structural"
        elif is_algorithm_result_figure(declared or inferred, figure):
            source_kind = "algorithm-result"
        elif is_runtime_screenshot_figure(declared or inferred, figure):
            source_kind = "runtime-screenshot"
        else:
            source_kind = "raster"
        runtime_route = (
            figure.get("real_route")
            or figure.get("route")
            or figure.get("page_url")
            or figure.get("pageURL")
            or figure.get("url")
            or ""
        )
        capture_method = figure.get("capture_method") or figure.get("method") or ""
        capture_kind = (
            figure.get("capture_kind")
            or figure.get("capture_source_kind")
            or figure.get("screenshot_capture_kind")
            or ""
        )
        readiness_cue = (
            figure.get("readiness_cue")
            or figure.get("ready_selector")
            or figure.get("wait_for")
            or figure.get("readiness")
            or ""
        )
        accepted_screenshot_path = (
            figure.get("accepted_screenshot_path")
            or figure.get("screenshot_path")
            or image_path
            or ""
        )
        caption_to_asset_mapping = (
            figure.get("caption_to_asset_mapping")
            or figure.get("caption_to_asset")
            or figure.get("caption_asset_map")
            or ""
        )
        accepted_result_image_path = (
            figure.get("accepted_result_image_path")
            or figure.get("result_image_path")
            or figure.get("output_image_path")
            or figure.get("accepted_screenshot_path")
            or figure.get("screenshot_path")
            or image_path
            or ""
        )
        algorithm_provenance_source = (
            figure.get("algorithm_provenance_source")
            or figure.get("result_source")
            or figure.get("existing_result_source")
            or figure.get("user_provided_asset_evidence")
            or figure.get("real_program_run_evidence")
            or ""
        )
        algorithm_authenticity_verdict = (
            figure.get("algorithm_authenticity_verdict")
            or figure.get("authenticity_verdict")
            or figure.get("result_provenance_verdict")
            or figure.get("algorithm_result_verdict")
            or ""
        )
        template_sample_baseline = (
            figure.get("template_sample_baseline")
            or figure.get("template_figure_sample_baseline")
            or figure.get("active_template_figure_sample")
            or figure.get("accepted_sample_figure_baseline")
            or ""
        )
        template_sample_evidence_path = (
            figure.get("template_sample_evidence_path")
            or figure.get("template_figure_sample_evidence")
            or figure.get("template_sample_check_evidence")
            or figure.get("active_template_figure_sample_evidence")
            or ""
        )
        no_template_figure_sample_verdict = (
            figure.get("no_template_figure_sample_verdict")
            or figure.get("template_sample_absence_verdict")
            or figure.get("template_figure_sample_absence_verdict")
            or ""
        )
        skill_internal_fallback_sample_path = (
            figure.get("skill_internal_fallback_sample_path")
            or figure.get("stored_fallback_sample_path")
            or figure.get("stored_skill_sample_path")
            or ""
        )
        chosen_style_source = (
            figure.get("chosen_style_source")
            or figure.get("final_chosen_style_source")
            or figure.get("source_of_truth_for_style")
            or ""
        )
        sample_lock_evidence_path = (
            figure.get("sample_lock_evidence_path")
            or figure.get("style_sample_lock_evidence")
            or figure.get("sample_lock_evidence")
            or ""
        )
        in_figure_language_verdict = (
            figure.get("in_figure_language_verdict")
            or figure.get("figure_language_verdict")
            or figure.get("chinese_label_verdict")
            or ""
        )
        english_label_exception_reason = (
            figure.get("english_label_exception_reason")
            or figure.get("english_identifier_exception_reason")
            or figure.get("literal_identifier_exception_reason")
            or ""
        )
        english_labels_present = (
            figure.get("english_labels_present")
            or figure.get("english_identifier_labels_present")
            or figure.get("has_english_labels")
            or ""
        )
        authoring_tool = figure.get("authoring_tool") or figure.get("final_authoring_tool") or ""
        final_source_kind = figure.get("final_source_kind") or figure.get("final_source_type") or source_kind
        entry = {
            "id": f"figure_{fig_idx}",
            "chapter_title": chapter_title,
            "caption": str(figure.get("caption") or ""),
            "description": str(figure.get("description") or figure.get("followup") or ""),
            "declared_family": declared,
            "inferred_family": inferred,
            "family": inferred,
            "source_kind": source_kind,
            "authoring_tool": str(authoring_tool),
            "final_source_kind": str(final_source_kind),
            "drawio_path": path_text(drawio_path),
            "svg_path": path_text(svg_path),
            "raster_fallback": path_text(image_path),
            "caption_to_asset_mapping": str(caption_to_asset_mapping),
            "geometry_validation_report": path_text(geometry_report),
            "source_scale_bbox_map": str(source_bbox_map or ""),
            "inserted_scale_geometry_evidence": str(inserted_geometry_evidence or ""),
            "dense_zone_crop_evidence": str(dense_zone_crop_evidence or ""),
            "collision_check_verdict": collision_verdict,
            "relation_attribute_collision_verdict": relation_attribute_collision_verdict,
            "shape_overlap_verdict": shape_overlap_verdict,
            "inserted_scale_collision_verdict": inserted_scale_collision_verdict,
            "source_to_inserted_geometry_verdict": source_to_inserted_geometry_verdict,
            "task_card": path_text(task_card),
            "post_insertion_rendered_evidence": path_text(post_insertion_rendered_evidence),
            "final_docx_relationship_evidence": path_text(final_docx_relationship_evidence),
            "rendered_page_status": rendered_page_status,
            "insertion_status": insertion_status,
            "template_sample_baseline": str(template_sample_baseline),
            "template_sample_evidence_path": path_text(template_sample_evidence_path),
            "no_template_figure_sample_verdict": str(no_template_figure_sample_verdict),
            "skill_internal_fallback_sample_path": path_text(skill_internal_fallback_sample_path),
            "chosen_style_source": str(chosen_style_source),
            "sample_lock_evidence_path": path_text(sample_lock_evidence_path),
            "in_figure_language_verdict": str(in_figure_language_verdict),
            "english_labels_present": str(english_labels_present),
            "english_label_exception_reason": str(english_label_exception_reason),
        }
        manifest["figures"][f"figure_{fig_idx}"] = {
            "path": entry["raster_fallback"],
            "caption": entry["caption"],
            "family": inferred,
            "source_kind": source_kind,
            "authoring_tool": entry["authoring_tool"],
            "final_source_kind": entry["final_source_kind"],
            "task_card": entry["task_card"],
            "post_insertion_rendered_evidence": entry["post_insertion_rendered_evidence"],
            "final_docx_relationship_evidence": entry["final_docx_relationship_evidence"],
            "rendered_page_status": entry["rendered_page_status"],
            "insertion_status": entry["insertion_status"],
            "template_sample_baseline": entry["template_sample_baseline"],
            "template_sample_evidence_path": entry["template_sample_evidence_path"],
            "no_template_figure_sample_verdict": entry["no_template_figure_sample_verdict"],
            "skill_internal_fallback_sample_path": entry["skill_internal_fallback_sample_path"],
            "chosen_style_source": entry["chosen_style_source"],
            "sample_lock_evidence_path": entry["sample_lock_evidence_path"],
            "in_figure_language_verdict": entry["in_figure_language_verdict"],
            "english_labels_present": entry["english_labels_present"],
            "english_label_exception_reason": entry["english_label_exception_reason"],
            "real_route": str(runtime_route),
            "page_url": str(runtime_route),
            "capture_method": str(capture_method),
            "capture_kind": str(capture_kind),
            "readiness_cue": str(readiness_cue),
            "accepted_screenshot_path": path_text(accepted_screenshot_path),
            "caption_to_asset_mapping": str(caption_to_asset_mapping),
            "capture_bbox": figure.get("capture_bbox") or figure.get("screenshot_bbox") or "",
            "window_bbox": figure.get("window_bbox") or figure.get("application_window_bbox") or "",
            "client_size": figure.get("client_size") or figure.get("app_client_size") or "",
            "expected_window_size": figure.get("expected_window_size") or figure.get("window_size") or "",
            "viewport_size": figure.get("viewport_size") or "",
            "image_width": figure.get("image_width") or "",
            "image_height": figure.get("image_height") or "",
            "full_window_coverage_ratio": figure.get("full_window_coverage_ratio")
            or figure.get("capture_coverage_ratio")
            or figure.get("coverage_ratio")
            or "",
            "full_window_capture_verdict": str(
                figure.get("full_window_capture_verdict")
                or figure.get("full_page_capture_verdict")
                or figure.get("capture_geometry_verdict")
                or ""
            ),
            "accepted_result_image_path": path_text(accepted_result_image_path),
            "result_image_path": path_text(accepted_result_image_path),
            "source_image_path": path_text(figure.get("source_image_path") or figure.get("input_image_path") or ""),
            "input_image_path": path_text(figure.get("input_image_path") or figure.get("source_image_path") or ""),
            "generation_script_path": path_text(
                figure.get("generation_script_path")
                or figure.get("inference_script_path")
                or figure.get("program_script_path")
                or ""
            ),
            "model_output_log_path": path_text(
                figure.get("model_output_log_path")
                or figure.get("inference_log_path")
                or figure.get("result_log_path")
                or ""
            ),
            "algorithm_provenance_source": str(algorithm_provenance_source),
            "existing_result_source": str(figure.get("existing_result_source") or ""),
            "user_provided_asset_evidence": str(figure.get("user_provided_asset_evidence") or ""),
            "real_program_run_evidence": str(figure.get("real_program_run_evidence") or ""),
            "algorithm_authenticity_verdict": str(algorithm_authenticity_verdict),
        }
        if source_kind == "structural":
            manifest["diagrams"][f"diagram_{fig_idx}"] = {
                "id": entry["id"],
                "chapter_title": chapter_title,
                "caption": entry["caption"],
                "description": entry["description"],
                "declared_family": declared,
                "inferred_family": inferred,
                "family": inferred,
                "source_kind": source_kind,
                "authoring_tool": entry["authoring_tool"],
                "final_source_kind": entry["final_source_kind"],
                "drawio": entry["drawio_path"],
                "svg": entry["svg_path"],
                "png": entry["raster_fallback"],
                "raster_fallback": entry["raster_fallback"],
                "caption_to_asset_mapping": entry["caption_to_asset_mapping"],
                "geometry_validation_report": entry["geometry_validation_report"],
                "source_scale_bbox_map": entry["source_scale_bbox_map"],
                "inserted_scale_geometry_evidence": entry["inserted_scale_geometry_evidence"],
                "dense_zone_crop_evidence": entry["dense_zone_crop_evidence"],
                "collision_check_verdict": entry["collision_check_verdict"],
                "relation_attribute_collision_verdict": entry["relation_attribute_collision_verdict"],
                "shape_overlap_verdict": entry["shape_overlap_verdict"],
                "inserted_scale_collision_verdict": entry["inserted_scale_collision_verdict"],
                "source_to_inserted_geometry_verdict": entry["source_to_inserted_geometry_verdict"],
                "task_card": entry["task_card"],
                "post_insertion_rendered_evidence": entry["post_insertion_rendered_evidence"],
                "final_docx_relationship_evidence": entry["final_docx_relationship_evidence"],
                "rendered_page_status": entry["rendered_page_status"],
                "insertion_status": entry["insertion_status"],
                "template_sample_baseline": entry["template_sample_baseline"],
                "template_sample_evidence_path": entry["template_sample_evidence_path"],
                "no_template_figure_sample_verdict": entry["no_template_figure_sample_verdict"],
                "skill_internal_fallback_sample_path": entry["skill_internal_fallback_sample_path"],
                "chosen_style_source": entry["chosen_style_source"],
                "sample_lock_evidence_path": entry["sample_lock_evidence_path"],
                "in_figure_language_verdict": entry["in_figure_language_verdict"],
                "english_labels_present": entry["english_labels_present"],
                "english_label_exception_reason": entry["english_label_exception_reason"],
            }
        fig_idx += 1

    for chapter in content.get("chapters", []) or []:
        stack = [chapter]
        while stack:
            node = stack.pop(0)
            for table_data in node.get("tables", []) or []:
                manifest["tables"][f"table_{table_idx}"] = {
                    "id": str(table_data.get("id") or f"table_{table_idx}"),
                    "chapter_title": str(node.get("title") or chapter.get("title") or ""),
                    "caption": str(table_data.get("caption") or ""),
                    "rows": len(table_data.get("rows", []) or []),
                    "table_authority_lock": str(
                        table_data.get("table_authority_lock")
                        or table_data.get("active_table_authority")
                        or table_data.get("table_authority")
                        or ""
                    ),
                    "authority_source_type": str(
                        table_data.get("authority_source_type")
                        or table_data.get("table_authority_source_type")
                        or ""
                    ),
                    "authority_source_file_path": str(
                        table_data.get("authority_source_file_path")
                        or table_data.get("table_authority_source_path")
                        or table_data.get("active_table_authority_path")
                        or ""
                    ),
                    "no_template_table_authority_verdict": str(
                        table_data.get("no_template_table_authority_verdict")
                        or table_data.get("table_authority_fallback_verdict")
                        or ""
                    ),
                    "table_authority_manuscript_binding_proof": str(
                        table_data.get("table_authority_manuscript_binding_proof")
                        or table_data.get("table_authority_binding_evidence")
                        or table_data.get("table_to_manuscript_binding_evidence")
                        or ""
                    ),
                    "title_mode": str(
                        table_data.get("title_mode")
                        or table_data.get("caption_title_mode")
                        or table_data.get("table_title_mode")
                        or ""
                    ),
                    "border_family_verdict": str(
                        table_data.get("border_family_verdict")
                        or table_data.get("three_line_border_verdict")
                        or table_data.get("table_border_family_verdict")
                        or ""
                    ),
                    "header_separator_verdict": str(
                        table_data.get("header_separator_verdict")
                        or table_data.get("table_header_separator_verdict")
                        or table_data.get("header_bottom_middle_rule_verdict")
                        or ""
                    ),
                    "vertical_separator_verdict": str(
                        table_data.get("vertical_separator_verdict")
                        or table_data.get("table_vertical_separator_verdict")
                        or ""
                    ),
                    "body_row_separator_verdict": str(
                        table_data.get("body_row_separator_verdict")
                        or table_data.get("table_body_row_separator_verdict")
                        or ""
                    ),
                    "table_local_structure_verdict": str(
                        table_data.get("table_local_structure_verdict")
                        or table_data.get("local_structure_verdict")
                        or ""
                    ),
                    "rendered_table_evidence": str(
                        table_data.get("rendered_table_evidence")
                        or table_data.get("rendered_table_evidence_path")
                        or table_data.get("rendered_table_page_evidence")
                        or table_data.get("table_rendered_baseline_comparison_evidence")
                        or table_data.get("table_rendered_baseline_comparison_evidence_path")
                        or ""
                    ),
                    "table_pagination_verdict": str(
                        table_data.get("table_pagination_verdict")
                        or table_data.get("table_continuation_verdict")
                        or table_data.get("table_row_split_header_repeat_verdict")
                        or ""
                    ),
                    "final_docx_table_evidence": str(
                        table_data.get("final_docx_table_evidence")
                        or table_data.get("final_docx_table_evidence_path")
                        or table_data.get("final_docx_relationship_evidence")
                        or table_data.get("table_final_docx_binding_evidence")
                        or ""
                    ),
                    "insertion_status": str(
                        table_data.get("insertion_status")
                        or table_data.get("table_insertion_status")
                        or table_data.get("final_insertion_status")
                        or ""
                    ),
                    "rendered_page_status": str(
                        table_data.get("rendered_page_status")
                        or table_data.get("table_rendered_page_status")
                        or table_data.get("rendered_table_verdict")
                        or ""
                    ),
                    "source_preserved_verdict": str(
                        table_data.get("source_preserved_verdict")
                        or table_data.get("table_source_preserved_verdict")
                        or table_data.get("table_preservation_verdict")
                        or ""
                    ),
                }
                table_idx += 1
            stack.extend(node.get("sections", []) or [])
    return manifest


def write_manifest(manifest: dict[str, Any], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def style_map(style: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for part in style.split(";"):
        if "=" in part:
            key, value = part.split("=", 1)
            values[key] = value
    return values


def normalize_color(value: object) -> str:
    return str(value or "").strip().lower()


def drawio_flowchart_issues(drawio: Path) -> list[str]:
    try:
        root = ET.fromstring(drawio.read_text(encoding="utf-8", errors="replace"))
    except ET.ParseError as exc:
        return [f"flowchart draw.io XML parse failed: {exc}"]
    vertices: list[dict[str, Any]] = []
    for cell in root.findall(".//mxCell"):
        if cell.attrib.get("vertex") != "1":
            continue
        geom = cell.find("mxGeometry")
        style = cell.attrib.get("style", "")
        if not str(cell.attrib.get("value", "")).strip() and (
            "shape=text" in style.lower() or style.lower().startswith("text;")
        ):
            continue
        values = style_map(style)
        vertices.append(
            {
                "value": str(cell.attrib.get("value", "")),
                "style": style,
                "style_map": values,
                "x": float(geom.attrib.get("x", "0")) if geom is not None else 0.0,
                "y": float(geom.attrib.get("y", "0")) if geom is not None else 0.0,
                "width": float(geom.attrib.get("width", "0")) if geom is not None else 0.0,
                "height": float(geom.attrib.get("height", "0")) if geom is not None else 0.0,
            }
        )
    if not vertices:
        return ["flowchart draw.io source has no usable vertices"]
    issues: list[str] = []
    ellipses = [node for node in vertices if "ellipse" in str(node["style"]).lower() or str(node["style_map"].get("rounded", "")) == "1"]
    processes = [node for node in vertices if node not in ellipses and "rhombus" not in str(node["style"]).lower()]
    if not any("\u5f00\u59cb" in str(node["value"]) or "start" in str(node["value"]).lower() for node in ellipses):
        issues.append("flowchart missing explicit start terminator")
    if not any("\u7ed3\u675f" in str(node["value"]) or "end" in str(node["value"]).lower() for node in ellipses):
        issues.append("flowchart missing explicit end terminator")
    if not processes:
        issues.append("flowchart missing process nodes")
    for node in vertices:
        values = dict(node["style_map"])
        fill = normalize_color(values.get("fillColor", ""))
        stroke = normalize_color(values.get("strokeColor", ""))
        font = normalize_color(values.get("fontColor", ""))
        if fill not in ALLOWED_WHITE:
            issues.append(f"flowchart non-white fill: {node['value']}")
            break
        if stroke not in ALLOWED_BLACK:
            issues.append(f"flowchart non-black stroke: {node['value']}")
            break
        if font not in ALLOWED_BLACK:
            issues.append(f"flowchart non-black font: {node['value']}")
            break
        if re.match(r"^(\u56fe|figure)\s*\d+", str(node["value"]).strip(), flags=re.I):
            issues.append(f"flowchart contains in-image caption/title text: {node['value']}")
            break
    return issues


def _decode_drawio_payload(payload: str) -> list[ET.Element]:
    roots: list[ET.Element] = []
    raw_payload = (payload or "").strip()
    if not raw_payload:
        return roots
    candidates = [raw_payload, urllib.parse.unquote(raw_payload)]
    for candidate in list(candidates):
        try:
            roots.append(ET.fromstring(candidate))
            continue
        except ET.ParseError:
            pass
        compact = "".join(candidate.split())
        try:
            data = base64.b64decode(compact + "=" * (-len(compact) % 4))
        except Exception:
            continue
        for wbits in (-15, zlib.MAX_WBITS):
            try:
                inflated = zlib.decompress(data, wbits)
            except Exception:
                continue
            for decoded in (inflated.decode("utf-8", errors="replace"), urllib.parse.unquote(inflated.decode("utf-8", errors="replace"))):
                try:
                    roots.append(ET.fromstring(decoded))
                except ET.ParseError:
                    pass
    return roots


def drawio_roots(drawio: Path) -> tuple[list[ET.Element], list[str]]:
    try:
        root = ET.fromstring(drawio.read_text(encoding="utf-8", errors="replace"))
    except ET.ParseError as exc:
        return [], [f"draw.io XML parse failed: {exc}"]
    roots = [root]
    if root.findall(".//mxCell"):
        return roots, []
    for diagram in root.findall(".//diagram"):
        roots.extend(_decode_drawio_payload(diagram.text or ""))
    return roots, []


def _float_value(value: object, default: float = 0.0) -> float:
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return default


def _clean_label(value: object) -> str:
    text = re.sub(r"<[^>]+>", "", str(value or ""))
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", text).strip() or "<blank>"


def _drawio_shape_kind(style: str) -> str:
    lowered = (style or "").lower()
    values = style_map(style or "")
    shape = (values.get("shape") or "").lower()
    if shape == "text" or lowered.startswith("text;") or ";text;" in lowered:
        return "annotation"
    if "rhombus" in lowered or shape == "rhombus":
        return "relation"
    if "ellipse" in lowered or shape == "ellipse":
        return "attribute"
    return "entity"


def _split_evidence_paths(value: object) -> list[Path]:
    raw = str(value or "").strip()
    if not raw:
        return []
    if raw.lower().startswith("not-applicable") or raw.lower() in {"none", "n/a"}:
        return []
    paths: list[Path] = []
    for part in re.split(r"[;\n|]+", raw):
        candidate = part.strip()
        if candidate:
            paths.append(Path(candidate))
    return paths


def _validate_evidence_paths(value: object, field_name: str) -> list[str]:
    paths = _split_evidence_paths(value)
    if not paths:
        return [f"ER structural figure missing {field_name}"]
    issues: list[str] = []
    for path in paths:
        if not path.exists():
            issues.append(f"ER structural figure {field_name} missing file: {path}")
        elif path.is_file() and path.stat().st_size <= 0:
            issues.append(f"ER structural figure {field_name} file is empty: {path}")
    return issues


def _verdict_is_pass(value: object) -> bool:
    return str(value or "").strip().lower().startswith("pass")


def _truthy_intent(value: object, expected: str) -> bool:
    return str(value or "").strip().lower() == expected


def _status_is_pass_or_inserted(value: object) -> bool:
    lowered = str(value or "").strip().lower()
    return lowered.startswith("pass") or lowered in {"inserted", "complete", "completed", "verified", "accepted"}


def _validate_existing_file_field(entry: dict[str, Any], key: str, aliases: tuple[str, ...], label: str) -> list[str]:
    raw = _entry_value(entry, aliases)
    if not raw:
        return [f"{key}: figure manifest entry missing {label}"]
    path = Path(raw)
    if not path.exists():
        return [f"{key}: figure manifest {label} file missing: {raw}"]
    if path.is_file() and path.stat().st_size <= 0:
        return [f"{key}: figure manifest {label} file is empty: {raw}"]
    return []


def validate_common_figure_entry_contract(key: str, entry: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    if not str(entry.get("caption") or "").strip():
        issues.append(f"{key}: figure manifest entry missing caption")
    if not _entry_value(entry, ("caption_to_asset_mapping", "caption_to_asset", "caption_asset_map", "caption_asset_mapping")):
        issues.append(f"{key}: figure manifest entry missing caption-to-asset mapping")
    issues.extend(
        _validate_existing_file_field(
            entry,
            key,
            ("task_card", "task_card_path", "figure_task_card"),
            "task card",
        )
    )
    issues.extend(
        _validate_existing_file_field(
            entry,
            key,
            (
                "post_insertion_rendered_evidence",
                "post_insertion_rendered_evidence_path",
                "rendered_page_evidence",
                "rendered_evidence",
            ),
            "post-insertion rendered evidence",
        )
    )
    issues.extend(
        _validate_existing_file_field(
            entry,
            key,
            ("final_docx_relationship_evidence", "docx_relationship_evidence", "relationship_evidence"),
            "final DOCX relationship evidence",
        )
    )
    rendered_status = (
        entry.get("rendered_page_status")
        or entry.get("post_insertion_rendered_verdict")
        or entry.get("rendered_verdict")
    )
    if not _status_is_pass_or_inserted(rendered_status):
        issues.append(f"{key}: rendered page status/verdict must be pass")
    insertion_status = entry.get("insertion_status") or entry.get("final_insertion_status")
    if not _status_is_pass_or_inserted(insertion_status):
        issues.append(f"{key}: insertion status must be pass/inserted")
    mutation_intent = str(
        entry.get("mutation_intent")
        or entry.get("figure_mutation_intent")
        or entry.get("replacement_intent")
        or entry.get("insertion_intent")
        or ""
    ).strip().lower()
    if mutation_intent in {"replace_existing", "replace-existing", "redraw_existing", "redraw-existing"}:
        if not _entry_value(
            entry,
            (
                "explicit_replacement_authorization_source",
                "replacement_authorization_source",
                "user_replacement_authorization",
                "teacher_comment_authorization",
            ),
        ):
            issues.append(f"{key}: existing figure replacement/redraw missing explicit replacement authorization source")
        if not _entry_value(
            entry,
            (
                "explicit_replacement_authorization_scope",
                "replacement_authorization_scope",
                "authorized_replacement_scope",
            ),
        ):
            issues.append(f"{key}: existing figure replacement/redraw missing explicit replacement authorization scope")
        for label, aliases in (
            ("original media relationship id", ("original_rid", "original_relationship_id")),
            ("original media target", ("original_media_target", "original_target")),
            ("original media sha256", ("original_media_sha256", "original_asset_sha256")),
            ("final media relationship id", ("final_rid", "final_relationship_id")),
            ("final media target", ("final_media_target", "final_target")),
            ("final media sha256", ("final_media_sha256", "final_asset_sha256")),
            ("target anchor caption", ("target_anchor_caption", "original_caption")),
        ):
            if not _entry_value(entry, aliases):
                issues.append(f"{key}: existing figure replacement/redraw missing {label}")
        if not _verdict_is_pass(
            entry.get("protected_surface_location_verdict")
            or entry.get("figure_anchor_location_verdict")
            or entry.get("target_anchor_not_protected_surface_verdict")
        ):
            issues.append(f"{key}: existing figure replacement/redraw location verdict must be pass")
    if mutation_intent in {
        "insert_image",
        "insert-image",
        "insert_figure",
        "insert-figure",
        "add_image",
        "add-image",
        "add_figure",
        "add-figure",
        "new_figure",
        "new-figure",
        "new_image",
        "new-image",
    }:
        if not _entry_value(
            entry,
            (
                "explicit_insertion_authorization_source",
                "insertion_authorization_source",
                "approved_figure_task_card",
                "figure_task_card",
                "task_card",
            ),
        ):
            issues.append(f"{key}: figure insertion missing explicit insertion authorization source")
        if not _entry_value(
            entry,
            (
                "explicit_insertion_authorization_scope",
                "insertion_authorization_scope",
                "authorized_insertion_scope",
            ),
        ):
            issues.append(f"{key}: figure insertion missing explicit insertion authorization scope")
        for label, aliases in (
            ("final media relationship id", ("final_rid", "final_relationship_id")),
            ("final media target", ("final_media_target", "final_target")),
            ("final media sha256", ("final_media_sha256", "final_asset_sha256")),
            ("target anchor caption", ("target_anchor_caption", "final_caption", "caption")),
        ):
            if not _entry_value(entry, aliases):
                issues.append(f"{key}: figure insertion missing {label}")
        if not _verdict_is_pass(
            entry.get("protected_surface_location_verdict")
            or entry.get("figure_anchor_location_verdict")
            or entry.get("target_anchor_not_protected_surface_verdict")
        ):
            issues.append(f"{key}: figure insertion location verdict must be pass")

    authoring_tool = str(entry.get("authoring_tool") or entry.get("final_authoring_tool") or "").strip().lower()
    final_source_kind = str(entry.get("final_source_kind") or entry.get("source_kind") or "").strip().lower()
    if _is_structural_manifest_entry(entry):
        if _has_structural_forbidden_source(authoring_tool):
            issues.append(f"{key}: structural figure final source cannot be {authoring_tool}")
        if _has_structural_forbidden_source(final_source_kind):
            issues.append(f"{key}: structural figure final_source_kind cannot be {final_source_kind}")
        if _truthy_intent(entry.get("is_generated_replacement"), "yes") and not _verdict_is_pass(
            entry.get("figure_contract_verdict") or entry.get("source_contract_verdict")
        ):
            issues.append(f"{key}: generated structural replacement requires a pass figure/source contract verdict")
    issues.extend(validate_template_sample_priority_contract(key, entry))
    return issues


def validate_template_sample_priority_contract(key: str, entry: dict[str, Any]) -> list[str]:
    if _is_preserved_existing_figure_entry(entry):
        return []
    issues: list[str] = []
    template_baseline = _entry_value(
        entry,
        (
            "template_sample_baseline",
            "template_figure_sample_baseline",
            "active_template_figure_sample",
            "accepted_sample_figure_baseline",
        ),
    )
    user_sample = _entry_value(entry, ("user_sample_baseline", "user_sample", "user_provided_sample"))
    no_template_verdict = _entry_value(
        entry,
        (
            "no_template_figure_sample_verdict",
            "template_sample_absence_verdict",
            "template_figure_sample_absence_verdict",
        ),
    )
    fallback = _entry_value(
        entry,
        (
            "skill_internal_fallback_sample_path",
            "stored_fallback_sample_path",
            "stored_skill_sample_path",
            "stored_fallback_sample",
        ),
    )
    if fallback:
        if not _verdict_is_pass(no_template_verdict):
            issues.append(f"{key}: skill-internal fallback sample used without pass no-template-figure-sample verdict")
        fallback_path = Path(fallback)
        if not fallback_path.exists():
            issues.append(f"{key}: skill-internal fallback sample path missing file: {fallback}")
    if template_baseline or user_sample:
        return issues
    if not _verdict_is_pass(no_template_verdict):
        issues.append(f"{key}: figure manifest missing template/sample baseline or pass no-template-figure-sample verdict")
    if _verdict_is_pass(no_template_verdict) and not fallback:
        issues.append(f"{key}: no-template-figure-sample verdict requires skill-internal fallback sample path")
    return issues


def validate_in_figure_language_contract(key: str, entry: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    verdict = _entry_value(
        entry,
        (
            "in_figure_language_verdict",
            "figure_language_verdict",
            "chinese_label_verdict",
        ),
    )
    if not _verdict_is_pass(verdict):
        issues.append(f"{key}: figure manifest in_figure_language_verdict must be pass")
    english_exception = _entry_value(
        entry,
        (
            "english_label_exception_reason",
            "english_identifier_exception_reason",
            "literal_identifier_exception_reason",
        ),
    )
    english_present = str(
        entry.get("english_labels_present")
        or entry.get("english_identifier_labels_present")
        or entry.get("has_english_labels")
        or ""
    ).strip().lower()
    if english_present in {"true", "yes", "1", "present"} and not english_exception:
        issues.append(f"{key}: English in-figure labels require a literal-identifier exception reason")
    return issues


def drawio_vertices(drawio: Path) -> tuple[list[dict[str, Any]], list[str]]:
    roots, issues = drawio_roots(drawio)
    vertices: list[dict[str, Any]] = []
    for root in roots:
        for cell in root.findall(".//mxCell"):
            if cell.attrib.get("vertex") != "1":
                continue
            geom = cell.find("mxGeometry")
            if geom is None:
                continue
            width = _float_value(geom.attrib.get("width"))
            height = _float_value(geom.attrib.get("height"))
            if width <= 0 or height <= 0:
                continue
            style = cell.attrib.get("style", "")
            x = _float_value(geom.attrib.get("x"))
            y = _float_value(geom.attrib.get("y"))
            vertices.append(
                {
                    "id": str(cell.attrib.get("id") or ""),
                    "label": _clean_label(cell.attrib.get("value")),
                    "kind": _drawio_shape_kind(style),
                    "style": style,
                    "x": x,
                    "y": y,
                    "width": width,
                    "height": height,
                    "right": x + width,
                    "bottom": y + height,
                }
            )
    return vertices, issues


def _drawio_visible_vertices(vertices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    visible: list[dict[str, Any]] = []
    for node in vertices:
        values = style_map(str(node.get("style") or ""))
        if _float_value(values.get("opacity"), 100.0) <= 0:
            continue
        if (values.get("shape") or "").lower() == "point":
            continue
        if str(node.get("kind") or "") == "annotation" and not str(node.get("label") or "").strip():
            continue
        if float(node.get("width") or 0) < 8 or float(node.get("height") or 0) < 8:
            continue
        visible.append(node)
    return visible


def _drawio_router_vertices(vertices: list[dict[str, Any]]) -> set[str]:
    routers: set[str] = set()
    for node in vertices:
        values = style_map(str(node.get("style") or ""))
        if (
            _float_value(values.get("opacity"), 100.0) <= 0
            or (values.get("shape") or "").lower() == "point"
            or float(node.get("width") or 0) <= 2
            or float(node.get("height") or 0) <= 2
        ):
            node_id = str(node.get("id") or "")
            if node_id:
                routers.add(node_id)
    return routers


def _drawio_edges(drawio: Path) -> tuple[list[dict[str, Any]], list[str]]:
    roots, issues = drawio_roots(drawio)
    edges: list[dict[str, Any]] = []
    for root in roots:
        for cell in root.findall(".//mxCell"):
            if cell.attrib.get("edge") != "1":
                continue
            geom = cell.find("mxGeometry")
            points: list[tuple[float, float]] = []
            if geom is not None:
                for point in geom.findall(".//mxPoint"):
                    points.append((_float_value(point.attrib.get("x")), _float_value(point.attrib.get("y"))))
            edges.append(
                {
                    "id": str(cell.attrib.get("id") or ""),
                    "source": str(cell.attrib.get("source") or ""),
                    "target": str(cell.attrib.get("target") or ""),
                    "style": str(cell.attrib.get("style") or ""),
                    "points": points,
                }
            )
    return edges, issues


def _drawio_line_segments(drawio: Path) -> tuple[list[dict[str, Any]], list[str]]:
    roots, issues = drawio_roots(drawio)
    segments: list[dict[str, Any]] = []
    for root in roots:
        for cell in root.findall(".//mxCell"):
            if cell.attrib.get("vertex") != "1":
                continue
            values = style_map(cell.attrib.get("style", ""))
            if (values.get("shape") or "").lower() != "line":
                continue
            geom = cell.find("mxGeometry")
            if geom is None:
                continue
            x = _float_value(geom.attrib.get("x"))
            y = _float_value(geom.attrib.get("y"))
            width = _float_value(geom.attrib.get("width"))
            height = _float_value(geom.attrib.get("height"))
            segments.append(
                {
                    "id": str(cell.attrib.get("id") or ""),
                    "x1": x,
                    "y1": y,
                    "x2": x + width,
                    "y2": y + height,
                }
            )
    return segments, issues


def _segment_intersects_rect(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    rect: dict[str, Any],
    *,
    margin: float = 3.0,
) -> bool:
    left = float(rect["x"]) + margin
    right = float(rect["right"]) - margin
    top = float(rect["y"]) + margin
    bottom = float(rect["bottom"]) - margin
    if right <= left or bottom <= top:
        return False
    dx = x2 - x1
    dy = y2 - y1
    p = [-dx, dx, -dy, dy]
    q = [x1 - left, right - x1, y1 - top, bottom - y1]
    u1 = 0.0
    u2 = 1.0
    for pi, qi in zip(p, q):
        if abs(pi) < 1e-9:
            if qi < 0:
                return False
            continue
        t = qi / pi
        if pi < 0:
            if t > u2:
                return False
            u1 = max(u1, t)
        else:
            if t < u1:
                return False
            u2 = min(u2, t)
    return u2 - u1 > 0.01


def _structural_edge_boundary_orthogonal_issues(edge: dict[str, Any], visible_by_id: dict[str, dict[str, Any]]) -> list[str]:
    issues: list[str] = []
    edge_id = str(edge.get("id") or "<missing>")
    source = str(edge.get("source") or "")
    target = str(edge.get("target") or "")
    style = str(edge.get("style") or "")
    values = style_map(style)
    edge_style = str(values.get("edgeStyle") or "").strip().lower()
    rounded = str(values.get("rounded") or "").strip().lower()
    curved = str(values.get("curved") or "").strip().lower()
    if source not in visible_by_id or target not in visible_by_id:
        issues.append(
            "structural connector is not boundary-bound to real visible source and target nodes: "
            f"edge `{edge_id}` source=`{source or '<missing>'}` target=`{target or '<missing>'}`"
        )
    if edge_style != "orthogonaledgestyle":
        issues.append(
            "structural connector must use orthogonal right-angle routing: "
            f"edge `{edge_id}` edgeStyle=`{values.get('edgeStyle') or '<missing>'}`"
        )
    if rounded != "0":
        issues.append(
            "structural connector must use square right-angle bends with rounded=0: "
            f"edge `{edge_id}` rounded=`{values.get('rounded') or '<missing>'}`"
        )
    if curved in {"1", "true"}:
        issues.append(
            "structural connector must not use curved routing: "
            f"edge `{edge_id}` curved=`{values.get('curved')}`"
        )
    return issues


def drawio_structural_geometry_report(drawio: Path, *, family: str = "structure") -> dict[str, Any]:
    vertices, vertex_issues = drawio_vertices(drawio)
    edges, edge_issues = _drawio_edges(drawio)
    line_segments, line_issues = _drawio_line_segments(drawio)
    issues = list(vertex_issues) + list(edge_issues) + list(line_issues)
    visible_vertices = _drawio_visible_vertices(vertices)
    vertex_by_id = {str(node["id"]): node for node in vertices if str(node.get("id") or "")}
    visible_by_id = {str(node["id"]): node for node in visible_vertices if str(node.get("id") or "")}
    router_ids = _drawio_router_vertices(vertices)
    container_ids = _drawio_container_vertex_ids(visible_vertices)
    if not visible_vertices:
        issues.append("structural draw.io source has no visible shape vertices")
    if not edges and not line_segments:
        issues.append("structural draw.io source has no connector edges or line segments")
    boundary_orthogonal_issues: list[str] = []
    for edge in edges:
        source = str(edge.get("source") or "")
        target = str(edge.get("target") or "")
        boundary_orthogonal_issues.extend(_structural_edge_boundary_orthogonal_issues(edge, visible_by_id))
        if source in router_ids or target in router_ids:
            boundary_orthogonal_issues.append(
                "structural connector uses invisible point/router vertex, which can create line routing through node boxes: "
                f"edge `{edge.get('id')}` source=`{source or '<missing>'}` target=`{target or '<missing>'}`"
            )
    for segment in line_segments:
        boundary_orthogonal_issues.append(
            "structural connector line is source/targetless; use a boundary-bound orthogonal edge instead: "
            f"line `{segment.get('id')}`"
        )
    issues.extend(boundary_orthogonal_issues)
    for index, left in enumerate(visible_vertices):
        for right in visible_vertices[index + 1 :]:
            if str(left.get("id") or "") in container_ids or str(right.get("id") or "") in container_ids:
                if _bbox_contains(left, right) or _bbox_contains(right, left):
                    continue
            overlap_x, overlap_y = _bbox_overlap(left, right)
            if overlap_x > 1 and overlap_y > 1:
                issues.append(
                    "structural shape overlap: "
                    f"`{left['label']}` intersects `{right['label']}` "
                    f"overlap_width={overlap_x:.2f} overlap_height={overlap_y:.2f}"
                )
    for edge in edges:
        source = str(edge.get("source") or "")
        target = str(edge.get("target") or "")
        path: list[tuple[float, float]] = []
        if source in vertex_by_id:
            node = vertex_by_id[source]
            path.append((float(node["x"]) + float(node["width"]) / 2, float(node["y"]) + float(node["height"]) / 2))
        path.extend(edge.get("points") or [])
        if target in vertex_by_id:
            node = vertex_by_id[target]
            path.append((float(node["x"]) + float(node["width"]) / 2, float(node["y"]) + float(node["height"]) / 2))
        if len(path) < 2:
            continue
        endpoint_ids = {source, target}
        for x1, y1, x2, y2 in zip([p[0] for p in path], [p[1] for p in path], [p[0] for p in path[1:]], [p[1] for p in path[1:]]):
            for node_id, node in visible_by_id.items():
                if node_id in endpoint_ids:
                    continue
                if node_id in container_ids:
                    continue
                if _segment_intersects_rect(x1, y1, x2, y2, node):
                    issues.append(
                        "structural connector crosses non-endpoint shape interior: "
                        f"edge `{edge.get('id')}` crosses `{node['label']}`"
                    )
    for segment in line_segments:
        for node_id, node in visible_by_id.items():
            if node_id in container_ids:
                continue
            if _segment_intersects_rect(
                float(segment["x1"]),
                float(segment["y1"]),
                float(segment["x2"]),
                float(segment["y2"]),
                node,
            ):
                issues.append(
                    "structural connector line crosses shape interior: "
                    f"line `{segment.get('id')}` crosses `{node['label']}`"
                )
    return {
        "schema": "graduation-project-builder.structural-figure-geometry.v1",
        "drawio_path": str(drawio.resolve()),
        "family": family,
        "visible_vertex_count": len(visible_vertices),
        "edge_count": len(edges),
        "line_segment_count": len(line_segments),
        "boundary_orthogonal_connector_verdict": "fail" if boundary_orthogonal_issues else "pass",
        "boundary_orthogonal_connector_issues": boundary_orthogonal_issues,
        "vertices": [
            {
                "id": node["id"],
                "label": node["label"],
                "kind": node["kind"],
                "bbox": {
                    "x": node["x"],
                    "y": node["y"],
                    "width": node["width"],
                    "height": node["height"],
                    "right": node["right"],
                    "bottom": node["bottom"],
                },
            }
            for node in visible_vertices
        ],
        "issues": issues,
        "verdict": "fail" if issues else "pass",
    }


def _bbox_contains(outer: dict[str, Any], inner: dict[str, Any], *, margin: float = 1.0) -> bool:
    return (
        float(outer["x"]) <= float(inner["x"]) + margin
        and float(outer["y"]) <= float(inner["y"]) + margin
        and float(outer["right"]) >= float(inner["right"]) - margin
        and float(outer["bottom"]) >= float(inner["bottom"]) - margin
    )


def _drawio_container_vertex_ids(vertices: list[dict[str, Any]]) -> set[str]:
    """Return visible frame/container ids so containment is not misread as overlap."""
    containers: set[str] = set()
    for outer in vertices:
        outer_id = str(outer.get("id") or "")
        if not outer_id:
            continue
        outer_area = float(outer.get("width") or 0) * float(outer.get("height") or 0)
        if outer_area <= 0:
            continue
        contained = 0
        for inner in vertices:
            if inner is outer:
                continue
            inner_area = float(inner.get("width") or 0) * float(inner.get("height") or 0)
            if inner_area <= 0:
                continue
            if outer_area >= inner_area * 2 and _bbox_contains(outer, inner, margin=2.0):
                contained += 1
        if contained >= 1:
            containers.add(outer_id)
    return containers


def drawio_structural_geometry_issues(drawio: Path, *, family: str = "structure") -> list[str]:
    report = drawio_structural_geometry_report(drawio, family=family)
    return [str(issue) for issue in report.get("issues", [])]


def _bbox_overlap(a: dict[str, Any], b: dict[str, Any]) -> tuple[float, float]:
    overlap_x = min(float(a["right"]), float(b["right"])) - max(float(a["x"]), float(b["x"]))
    overlap_y = min(float(a["bottom"]), float(b["bottom"])) - max(float(a["y"]), float(b["y"]))
    return overlap_x, overlap_y


def _bbox_clearance(a: dict[str, Any], b: dict[str, Any]) -> float:
    overlap_x, overlap_y = _bbox_overlap(a, b)
    if overlap_x > 0 and overlap_y > 0:
        return 0.0
    dx = max(float(a["x"]) - float(b["right"]), float(b["x"]) - float(a["right"]), 0.0)
    dy = max(float(a["y"]) - float(b["bottom"]), float(b["y"]) - float(a["bottom"]), 0.0)
    return (dx * dx + dy * dy) ** 0.5


def drawio_er_geometry_report(drawio: Path) -> dict[str, Any]:
    vertices, parse_issues = drawio_vertices(drawio)
    issues = list(parse_issues)
    er_vertices = [node for node in vertices if node["kind"] in {"entity", "relation", "attribute"}]
    relations = [node for node in er_vertices if node["kind"] == "relation"]
    attributes = [node for node in er_vertices if node["kind"] == "attribute"]
    if not er_vertices:
        issues.append("ER draw.io source has no usable entity/relation/attribute vertices")
    if not relations:
        issues.append("ER draw.io source has no relationship diamonds")
    if not attributes:
        issues.append("ER draw.io source has no attribute ellipses")
    for index, left in enumerate(er_vertices):
        for right in er_vertices[index + 1:]:
            overlap_x, overlap_y = _bbox_overlap(left, right)
            if overlap_x > 0 and overlap_y > 0:
                issues.append(
                    "ER shape overlap: "
                    f"{left['kind']} `{left['label']}` intersects {right['kind']} `{right['label']}` "
                    f"overlap_width={overlap_x:.2f} overlap_height={overlap_y:.2f}"
                )
    for relation in relations:
        for attribute in attributes:
            clearance = _bbox_clearance(relation, attribute)
            if clearance < ER_RELATION_ATTRIBUTE_MIN_CLEARANCE:
                issues.append(
                    "ER relation-attribute clearance failure: "
                    f"relation `{relation['label']}` vs attribute `{attribute['label']}` "
                    f"clearance={clearance:.2f} minimum={ER_RELATION_ATTRIBUTE_MIN_CLEARANCE:.2f}"
                )
    return {
        "schema": "graduation-project-builder.structural-figure-geometry.v1",
        "drawio_path": str(drawio.resolve()),
        "family": "er",
        "minimum_relation_attribute_clearance": ER_RELATION_ATTRIBUTE_MIN_CLEARANCE,
        "vertices": [
            {
                "id": node["id"],
                "label": node["label"],
                "kind": node["kind"],
                "bbox": {
                    "x": node["x"],
                    "y": node["y"],
                    "width": node["width"],
                    "height": node["height"],
                    "right": node["right"],
                    "bottom": node["bottom"],
                },
            }
            for node in er_vertices
        ],
        "issues": issues,
        "verdict": "fail" if issues else "pass",
    }


def drawio_er_geometry_issues(drawio: Path) -> list[str]:
    report = drawio_er_geometry_report(drawio)
    return [str(issue) for issue in report.get("issues", [])]


def validate_structural_geometry_report(report_path: Path, drawio: Path) -> list[str]:
    issues: list[str] = []
    if not report_path.exists():
        return [f"structural figure geometry report missing: {report_path}"]
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"structural figure geometry report is not valid JSON: {report_path} ({exc})"]
    if report.get("schema") != "graduation-project-builder.structural-figure-geometry.v1":
        issues.append(f"structural figure geometry report schema mismatch: {report_path}")
    if str(report.get("family") or "").lower() != "er":
        issues.append(f"structural figure geometry report family must be er: {report_path}")
    reported_drawio = Path(str(report.get("drawio_path") or "")).resolve()
    if reported_drawio != drawio.resolve():
        issues.append(f"structural figure geometry report draw.io path mismatch: {report_path}")
    if report.get("verdict") != "pass":
        issues.append(f"structural figure geometry report verdict is not pass: {report_path}")
    report_issues = report.get("issues")
    if report_issues not in ([], None):
        issues.append(f"structural figure geometry report still lists issues: {report_path}")
    vertices = report.get("vertices")
    if not isinstance(vertices, list) or not vertices:
        issues.append(f"structural figure geometry report has no vertex bbox map: {report_path}")
    else:
        kinds = {str(item.get("kind") or "") for item in vertices if isinstance(item, dict)}
        if not {"entity", "relation", "attribute"}.issubset(kinds):
            issues.append(f"structural figure geometry report must include entity, relation, and attribute bbox rows: {report_path}")
        for item in vertices:
            if not isinstance(item, dict) or not isinstance(item.get("bbox"), dict):
                issues.append(f"structural figure geometry report vertex lacks bbox object: {report_path}")
                break
    return issues


def validate_non_er_structural_geometry_report(report_path: Path, drawio: Path) -> list[str]:
    issues: list[str] = []
    if not report_path.exists():
        return [f"non-ER structural figure geometry report missing: {report_path}"]
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"non-ER structural figure geometry report is not valid JSON: {report_path} ({exc})"]
    if report.get("schema") != "graduation-project-builder.structural-figure-geometry.v1":
        issues.append(f"non-ER structural figure geometry report schema mismatch: {report_path}")
    family = str(report.get("family") or "").lower()
    if family in {"", "er"}:
        issues.append(f"non-ER structural figure geometry report family must be non-ER structural/flowchart: {report_path}")
    reported_drawio = Path(str(report.get("drawio_path") or "")).resolve()
    if reported_drawio != drawio.resolve():
        issues.append(f"non-ER structural figure geometry report draw.io path mismatch: {report_path}")
    if report.get("verdict") != "pass":
        issues.append(f"non-ER structural figure geometry report verdict is not pass: {report_path}")
    if report.get("boundary_orthogonal_connector_verdict") != "pass":
        issues.append(f"non-ER structural figure geometry report boundary/orthogonal connector verdict is not pass: {report_path}")
    report_issues = report.get("issues")
    if report_issues not in ([], None):
        issues.append(f"non-ER structural figure geometry report still lists issues: {report_path}")
    vertices = report.get("vertices")
    if not isinstance(vertices, list) or not vertices:
        issues.append(f"non-ER structural figure geometry report has no visible vertex bbox map: {report_path}")
    else:
        for item in vertices:
            if not isinstance(item, dict) or not isinstance(item.get("bbox"), dict):
                issues.append(f"non-ER structural figure geometry report vertex lacks bbox object: {report_path}")
                break
    edge_count = report.get("edge_count")
    line_segment_count = report.get("line_segment_count")
    if not isinstance(edge_count, int) or not isinstance(line_segment_count, int):
        issues.append(f"non-ER structural figure geometry report lacks connector counts: {report_path}")
    elif edge_count + line_segment_count <= 0:
        issues.append(f"non-ER structural figure geometry report has no connector geometry: {report_path}")
    return issues


def docx_image_targets(docx_path: Path) -> list[str]:
    formula_relationship_keys = docx_formula_image_relationship_keys(docx_path)
    return [
        entry.get("target", "")
        for key, entry in docx_image_relationship_manifest(docx_path).items()
        if key not in formula_relationship_keys
    ]


def docx_image_relationship_manifest(docx_path: Path) -> dict[str, dict[str, str]]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            result: dict[str, dict[str, str]] = {}
            rel_parts = sorted(
                name
                for name in names
                if name.startswith("word/") and "/_rels/" in name and name.endswith(".rels")
            )
            for rels_part in rel_parts:
                try:
                    root = ET.fromstring(zf.read(rels_part))
                except Exception:
                    continue
                source_dir = rels_part.rsplit("/_rels/", 1)[0]
                source_name = rels_part.rsplit("/_rels/", 1)[1][:-5]
                source_part = f"{source_dir}/{source_name}"
                for rel in root.findall(f"{{{PR_NS}}}Relationship"):
                    if "image" not in rel.attrib.get("Type", ""):
                        continue
                    rid = rel.attrib.get("Id", "")
                    target = rel.attrib.get("Target", "")
                    decoded_target = urllib.parse.unquote(target)
                    if decoded_target.startswith("/"):
                        media_name = decoded_target.lstrip("/")
                    else:
                        media_name = posixpath.normpath(posixpath.join(source_dir, decoded_target))
                    target_exists = media_name in names
                    digest = hashlib.sha256(zf.read(media_name)).hexdigest() if target_exists else ""
                    key = f"{source_part}#{rid}"
                    result[key] = {
                        "key": key,
                        "rid": rid,
                        "target": target,
                        "media_name": media_name,
                        "sha256": digest,
                        "target_exists": "true" if target_exists else "false",
                        "rels_part": rels_part,
                        "source_part": source_part,
                    }
            return result
    except Exception:
        return {}


def _docx_word_xml_story_parts(names: set[str]) -> list[str]:
    return sorted(
        name
        for name in names
        if name.startswith("word/")
        and name.endswith(".xml")
        and "/_rels/" not in name
        and not name.startswith("word/media/")
    )


def paragraph_has_omml(paragraph: ET.Element) -> bool:
    return any(
        node.tag in {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}
        or node.tag.startswith(f"{{{M_NS}}}")
        for node in paragraph.iter()
    )


def paragraph_image_relationship_ids(paragraph: ET.Element) -> list[str]:
    ids = [
        value
        for blip in paragraph.findall(f".//{{{A_NS}}}blip")
        for value in (
            blip.attrib.get(f"{{{R_NS}}}embed", ""),
            blip.attrib.get(f"{{{R_NS}}}link", ""),
        )
        if value
    ]
    for descendant in paragraph.iter():
        if descendant.tag.endswith("}imagedata"):
            for attr_name in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed"):
                value = descendant.attrib.get(attr_name, "")
                if value:
                    ids.append(value)
    return ids


def paragraph_has_ole_object(paragraph: ET.Element) -> bool:
    return any(node.tag.endswith("}OLEObject") for node in paragraph.iter())


def _all_media_targets_are_vector_formula_fallback(media_targets: list[str]) -> bool:
    targets = [str(item or "").lower() for item in media_targets if str(item or "").strip()]
    return bool(targets) and all(target.endswith((".wmf", ".emf")) for target in targets)


def paragraph_is_formula_like_drawing(
    paragraph: ET.Element,
    *,
    text: str = "",
    previous_text: str = "",
    next_text: str = "",
    extents: list[dict[str, int]] | None = None,
    media_targets: list[str] | None = None,
) -> bool:
    if paragraph_has_omml(paragraph):
        return True
    if is_docx_figure_caption(text):
        return False
    media_targets = media_targets or []
    vector_formula_media = _all_media_targets_are_vector_formula_fallback(media_targets)
    formula_media_count = sum(1 for item in media_targets if str(item or "").strip())
    adjacent_caption = is_docx_figure_caption(previous_text) or is_docx_figure_caption(next_text)
    if adjacent_caption and not text.strip():
        return False
    compact_context = re.sub(r"\s+", "", f"{previous_text}{text}{next_text}")
    has_vector_formula_media = vector_formula_media
    if paragraph_has_ole_object(paragraph):
        return True
    has_formula_label_text = bool(FORMULA_LABEL_TEXT_RE.match(text or ""))
    has_formula_context = any(token in compact_context for token in FORMULA_CONTEXT_TOKENS)
    max_width = max((int(item.get("cx") or 0) for item in (extents or [])), default=0)
    max_height = max((int(item.get("cy") or 0) for item in (extents or [])), default=0)
    # Context words such as "calculation" often appear near real engineering
    # figures. Treat them as formula-fallback evidence only when the drawing is
    # a vector fallback and the block is not caption-adjacent.
    if (has_formula_label_text or has_formula_context) and has_vector_formula_media and not adjacent_caption:
        return True
    if text.strip():
        # Text-bearing image paragraphs can be legitimate mixed figure defects.
        # Exclude them from the figure lane only when the media and surrounding
        # context identify a vector formula fallback. A multi-image fallback
        # block with body text is especially likely to be a formula/derivation
        # paragraph rather than a thesis figure.
        if not has_vector_formula_media:
            return False
        if has_formula_label_text or has_formula_context:
            return True
        if formula_media_count > 1:
            return True
        if max_height and max_height <= int(3.2 * 360000):
            return True
        if max_width and max_width <= int(2.8 * 360000):
            return True
        return False
    if adjacent_caption:
        return False
    if max_height and max_height <= int(3.2 * 360000) and has_vector_formula_media:
        return True
    return False


def docx_formula_image_relationship_keys(docx_path: Path) -> set[str]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            media_manifest = docx_image_relationship_manifest(docx_path)
            keys: set[str] = set()
            for story_part in _docx_word_xml_story_parts(names):
                try:
                    root = ET.fromstring(zf.read(story_part))
                except Exception:
                    continue
                paragraphs = root.findall(f".//{{{W_NS}}}p")
                for paragraph_index, paragraph in enumerate(paragraphs, start=1):
                    rids = paragraph_image_relationship_ids(paragraph)
                    if not rids:
                        continue
                    current_index = paragraph_index - 1
                    previous_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, -1)
                    next_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, 1)
                    media_targets = [
                        str(media_manifest.get(f"{story_part}#{rid}", {}).get("target", ""))
                        for rid in rids
                    ]
                    if not paragraph_is_formula_like_drawing(
                        paragraph,
                        text=_paragraph_text(paragraph),
                        previous_text=previous_text,
                        next_text=next_text,
                        extents=_paragraph_extents(paragraph),
                        media_targets=media_targets,
                    ):
                        continue
                    for rid in rids:
                        keys.add(f"{story_part}#{rid}")
            return keys
    except Exception:
        return set()


def docx_drawing_object_manifest(docx_path: Path) -> dict[str, dict[str, Any]]:
    """Return a package-wide drawing placement manifest.

    Media relationship hashes alone miss unsafe figure edits such as resizing,
    inline/anchor conversion, and caption adjacency drift. This manifest keeps
    those surfaces visible to transaction and figure-contract gates.
    """

    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            media_manifest = docx_image_relationship_manifest(docx_path)
            result: dict[str, dict[str, Any]] = {}
            for story_part in _docx_word_xml_story_parts(names):
                try:
                    root = ET.fromstring(zf.read(story_part))
                except Exception:
                    continue
                paragraphs = root.findall(f".//{{{W_NS}}}p")
                for paragraph_index, paragraph in enumerate(paragraphs, start=1):
                    has_drawing = (
                        paragraph.find(f".//{{{W_NS}}}drawing") is not None
                        or paragraph.find(f".//{{{W_NS}}}object") is not None
                        or paragraph.find(f".//{{{W_NS}}}pict") is not None
                    )
                    if not has_drawing:
                        continue
                    blip_rids = [
                        value
                        for blip in paragraph.findall(f".//{{{A_NS}}}blip")
                        for value in (
                            blip.attrib.get(f"{{{R_NS}}}embed", ""),
                            blip.attrib.get(f"{{{R_NS}}}link", ""),
                        )
                        if value
                    ]
                    for descendant in paragraph.iter():
                        if descendant.tag.endswith("}imagedata"):
                            for attr_name in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed"):
                                value = descendant.attrib.get(attr_name, "")
                                if value:
                                    blip_rids.append(value)
                    rid_signature = ";".join(blip_rids)
                    media_rows = [
                        media_manifest.get(f"{story_part}#{rid}", {})
                        for rid in blip_rids
                    ]
                    raw_extents = _paragraph_extents(paragraph)
                    current_index = paragraph_index - 1
                    previous_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, -1)
                    next_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, 1)
                    paragraph_text = _paragraph_text(paragraph)
                    if paragraph_is_formula_like_drawing(
                        paragraph,
                        text=paragraph_text,
                        previous_text=previous_text,
                        next_text=next_text,
                        extents=raw_extents,
                        media_targets=[str(row.get("target", "")) for row in media_rows],
                    ):
                        continue
                    media_signature = ";".join(
                        f"{row.get('rid', '')}|{row.get('target', '')}|{row.get('sha256', '')}"
                        for row in media_rows
                    )
                    extents = [
                        f"{extent['cx']}x{extent['cy']}"
                        for extent in raw_extents
                    ]
                    inline_count = len(paragraph.findall(f".//{{{WP_NS}}}inline"))
                    anchor_count = len(paragraph.findall(f".//{{{WP_NS}}}anchor"))
                    object_count = len(paragraph.findall(f".//{{{W_NS}}}object"))
                    pict_count = len(paragraph.findall(f".//{{{W_NS}}}pict"))
                    row: dict[str, Any] = {
                        "story_part": story_part,
                        "paragraph_index": str(paragraph_index),
                        "drawing_kind": f"inline={inline_count};anchor={anchor_count};object={object_count};pict={pict_count}",
                        "extent_signature": ";".join(extents),
                        "relationship_ids": rid_signature,
                        "media_signature": media_signature,
                        "paragraph_text": paragraph_text,
                        "previous_text": previous_text,
                        "next_text": next_text,
                        "next_is_figure_caption": is_docx_figure_caption(next_text),
                    }
                    row_hash = hashlib.sha256(
                        json.dumps(row, ensure_ascii=False, sort_keys=True).encode("utf-8")
                    ).hexdigest()
                    key_tail = rid_signature or row_hash[:12]
                    key = f"{story_part}#p{paragraph_index}#{key_tail}"
                    row["key"] = key
                    row["sha256"] = row_hash
                    result[key] = row
            return result
    except Exception:
        return {}


def docx_drawing_object_manifest_sha256(docx_path: Path) -> str:
    manifest = docx_drawing_object_manifest(docx_path)
    payload = json.dumps(manifest, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _replacement_intent(value: object) -> bool:
    text = str(value or "").strip().lower()
    return text in {"replace_existing", "replace-existing", "redraw_existing", "redraw-existing"}


def _insertion_intent(value: object) -> bool:
    text = str(value or "").strip().lower()
    return text in {
        "insert_image",
        "insert-image",
        "insert_figure",
        "insert-figure",
        "add_image",
        "add-image",
        "add_figure",
        "add-figure",
        "new_figure",
        "new-figure",
        "new_image",
        "new-image",
    }


def _removal_intent(value: object) -> bool:
    text = str(value or "").strip().lower()
    return text in {
        "remove_existing",
        "remove-existing",
        "remove_image",
        "remove-image",
        "delete_image",
        "delete-image",
        "remove_figure",
        "remove-figure",
        "delete_figure",
        "delete-figure",
        "remove_media_relationship",
        "remove-media-relationship",
        "remove_noncompliant_cad_converted_thesis_image",
        "remove_noncompliant_cad_converted_thesis_images",
    }


def _entry_replacement_intent(entry: dict[str, Any]) -> bool:
    return _replacement_intent(
        entry.get("mutation_intent")
        or entry.get("figure_mutation_intent")
        or entry.get("replacement_intent")
    )


def _entry_insertion_intent(entry: dict[str, Any]) -> bool:
    return _insertion_intent(
        entry.get("mutation_intent")
        or entry.get("figure_mutation_intent")
        or entry.get("insertion_intent")
    )


def _entry_removal_intent(entry: dict[str, Any]) -> bool:
    return _removal_intent(
        entry.get("mutation_intent")
        or entry.get("figure_mutation_intent")
        or entry.get("removal_intent")
        or entry.get("media_removal_intent")
    )


def manifest_has_replacement_intent(manifest: dict[str, Any]) -> bool:
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry in collection.values():
            if isinstance(entry, dict) and _entry_replacement_intent(entry):
                return True
    return False


def manifest_has_image_mutation_intent(manifest: dict[str, Any]) -> bool:
    if manifest_has_replacement_intent(manifest):
        return True
    for collection_name in (
        "media_removal_authorizations",
        "media_relationship_removal_authorizations",
        "image_removal_authorizations",
    ):
        collection = manifest.get(collection_name)
        if not isinstance(collection, list):
            continue
        for entry in collection:
            if isinstance(entry, dict) and _entry_removal_intent(entry):
                return True
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry in collection.values():
            if isinstance(entry, dict) and (_entry_insertion_intent(entry) or _entry_removal_intent(entry)):
                return True
    return False


def manifest_has_authorized_display_extent_resize(manifest: dict[str, Any]) -> bool:
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry in collection.values():
            if not isinstance(entry, dict):
                continue
            intent = str(entry.get("display_extent_resize_intent") or "").strip().lower()
            if intent not in {"resize_display_extent", "display_extent_resize"}:
                continue
            if not _verdict_is_pass(entry.get("display_extent_resize_verdict")):
                continue
            if not _entry_value(entry, ("resize_authorization_source", "explicit_resize_authorization_source")):
                continue
            if not _entry_value(entry, ("original_drawing_sha256", "source_drawing_sha256")):
                continue
            if not _entry_value(entry, ("final_drawing_sha256", "final_drawing_object_sha256")):
                continue
            return True
    return False


def _manifest_authorized_image_changes(manifest: dict[str, Any]) -> list[dict[str, str]]:
    authorized: list[dict[str, str]] = []

    def add_authorization(entry_key: str, entry: dict[str, Any]) -> None:
        is_replacement = _entry_replacement_intent(entry)
        is_insertion = _entry_insertion_intent(entry)
        is_removal = _entry_removal_intent(entry)
        if not (is_replacement or is_insertion or is_removal):
            return
        authorization = _entry_value(
            entry,
            (
                "explicit_replacement_authorization_source",
                "replacement_authorization_source",
                "explicit_insertion_authorization_source",
                "insertion_authorization_source",
                "explicit_removal_authorization_source",
                "removal_authorization_source",
                "user_replacement_authorization",
                "teacher_comment_authorization",
                "approved_figure_task_card",
                "figure_task_card",
                "task_card",
            ),
        )
        if not authorization:
            return
        if is_removal:
            change_type = "removal"
        elif is_replacement:
            change_type = "replacement"
        else:
            change_type = "insertion"
        authorized.append(
            {
                "entry_key": str(entry_key),
                "change_type": change_type,
                "original_rid": _entry_value(entry, ("original_rid", "original_relationship_id")),
                "original_target": _entry_value(entry, ("original_media_target", "original_target")),
                "original_sha256": _entry_value(entry, ("original_media_sha256", "original_asset_sha256")).lower(),
                "original_part": _entry_value(
                    entry,
                    ("original_owner_part", "original_rels_part", "original_relationship_part", "original_part"),
                ),
                "final_rid": _entry_value(entry, ("final_rid", "final_relationship_id")),
                "final_target": _entry_value(entry, ("final_media_target", "final_target")),
                "final_sha256": _entry_value(entry, ("final_media_sha256", "final_asset_sha256")).lower(),
                "final_part": _entry_value(
                    entry,
                    ("final_owner_part", "final_rels_part", "final_relationship_part", "final_part"),
                ),
            }
        )

    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry_key, entry in collection.items():
            if not isinstance(entry, dict):
                continue
            add_authorization(str(entry_key), entry)
            svg_rid = _entry_value(entry, ("docx_svg_rid", "svg_rid"))
            svg_target = _entry_value(entry, ("docx_svg_media_target", "svg_docx_media_target"))
            svg_sha256 = _entry_value(entry, ("docx_svg_media_sha256", "svg_docx_media_sha256")).lower()
            if svg_rid and svg_target and svg_sha256:
                authorized.append(
                    {
                        "entry_key": f"{entry_key}:svg-primary",
                        "change_type": "insertion",
                        "original_rid": "",
                        "original_target": "",
                        "original_sha256": "",
                        "original_part": "",
                        "final_rid": svg_rid,
                        "final_target": svg_target,
                        "final_sha256": svg_sha256,
                        "final_part": _entry_value(
                            entry,
                            ("docx_svg_owner_part", "final_owner_part", "final_rels_part", "final_relationship_part", "final_part"),
                        ),
                    }
                )
    for collection_name in (
        "media_removal_authorizations",
        "media_relationship_removal_authorizations",
        "image_removal_authorizations",
    ):
        collection = manifest.get(collection_name)
        if not isinstance(collection, list):
            continue
        for index, entry in enumerate(collection, start=1):
            if isinstance(entry, dict):
                add_authorization(f"{collection_name}[{index}]", entry)
    return authorized


def _media_identity(info: dict[str, str]) -> tuple[str, str, str]:
    return (info.get("rid", ""), info.get("target", ""), info.get("sha256", "").lower())


def _media_sha_set(media: dict[str, dict[str, str]]) -> set[str]:
    return {
        str(info.get("sha256") or "").strip().lower()
        for info in media.values()
        if str(info.get("sha256") or "").strip()
    }


def _manifest_template_docx_paths(manifest: dict[str, Any]) -> list[Path]:
    paths: list[Path] = []
    raw_values: list[object] = [
        manifest.get("template_docx"),
        manifest.get("template_docx_path"),
        manifest.get("reference_docx"),
        manifest.get("reference_docx_path"),
    ]
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry in collection.values():
            if not isinstance(entry, dict):
                continue
            raw_values.extend(
                [
                    entry.get("template_sample_baseline"),
                    entry.get("template_figure_sample_baseline"),
                    entry.get("active_template_figure_sample"),
                ]
            )
    seen: set[Path] = set()
    for raw in raw_values:
        value = str(raw or "").strip()
        if not value:
            continue
        path = Path(value)
        if path.suffix.lower() != ".docx" or not path.exists():
            continue
        try:
            resolved = path.resolve()
        except OSError:
            resolved = path
        if resolved not in seen:
            seen.add(resolved)
            paths.append(resolved)
    return paths


def _manifest_template_media_sha_set(manifest: dict[str, Any]) -> set[str]:
    result: set[str] = set()
    for path in _manifest_template_docx_paths(manifest):
        result.update(_media_sha_set(docx_image_relationship_manifest(path)))
    return result


def _is_template_cover_media(info: dict[str, str]) -> bool:
    target = str(info.get("target") or info.get("media_name") or "").lower()
    return "template-cover" in target or "direct-donor" in target


def _media_hash_allowed_by_source_or_template(
    info: dict[str, str],
    *,
    source_hashes: set[str],
    final_hashes: set[str],
    template_hashes: set[str],
) -> bool:
    digest = str(info.get("sha256") or "").strip().lower()
    if not digest:
        return False
    return (
        digest in source_hashes
        or digest in final_hashes
        or digest in template_hashes
        or _is_template_cover_media(info)
    )


def _media_matches_authorized_original(info: dict[str, str], auth: dict[str, str]) -> bool:
    if auth.get("original_rid") != info.get("rid", ""):
        return False
    if auth.get("original_target") != info.get("target", ""):
        return False
    if auth.get("original_sha256") != info.get("sha256", "").lower():
        return False
    part = auth.get("original_part", "")
    return not part or part in {info.get("source_part", ""), info.get("rels_part", "")}


def _media_matches_authorized_final(info: dict[str, str], auth: dict[str, str]) -> bool:
    if auth.get("final_rid") != info.get("rid", ""):
        return False
    if auth.get("final_target") != info.get("target", ""):
        return False
    if auth.get("final_sha256") != info.get("sha256", "").lower():
        return False
    part = auth.get("final_part", "")
    return not part or part in {info.get("source_part", ""), info.get("rels_part", "")}


def _duplicate_identities(media: dict[str, dict[str, str]]) -> set[tuple[str, str, str]]:
    counts: dict[tuple[str, str, str], int] = {}
    for info in media.values():
        identity = _media_identity(info)
        counts[identity] = counts.get(identity, 0) + 1
    return {identity for identity, count in counts.items() if count > 1}


def _find_authorization_for_source(
    source: dict[str, str],
    authorized: list[dict[str, str]],
    duplicate_source_identities: set[tuple[str, str, str]],
    issues: list[str],
) -> dict[str, str] | None:
    candidates = [auth for auth in authorized if _media_matches_authorized_original(source, auth)]
    if not candidates:
        return None
    if _media_identity(source) in duplicate_source_identities and not any(auth.get("original_part") for auth in candidates):
        issues.append(
            "source-to-final media replacement authorization is ambiguous; original owner part is required: "
            f"rid={source.get('rid', '<missing>')} target={source.get('target', '<missing>')}"
        )
        return None
    return candidates[0]


def _authorization_final_binding_issue(
    auth: dict[str, str],
    final: dict[str, str] | None,
    final_media: dict[str, dict[str, str]],
    duplicate_final_identities: set[tuple[str, str, str]],
) -> str | None:
    missing = [field for field in ("final_rid", "final_target", "final_sha256") if not auth.get(field)]
    if missing:
        return (
            f"{auth.get('entry_key', '<unknown>')}: authorized replacement missing final media binding fields: "
            + ", ".join(missing)
        )
    if (
        (auth.get("final_rid"), auth.get("final_target"), auth.get("final_sha256"))
        in duplicate_final_identities
        and not auth.get("final_part")
    ):
        return (
            f"{auth.get('entry_key', '<unknown>')}: authorized replacement final media binding is ambiguous; "
            "final owner part is required"
        )
    candidates = [info for info in final_media.values() if _media_matches_authorized_final(info, auth)]
    if not candidates:
        return (
            f"{auth.get('entry_key', '<unknown>')}: authorized replacement final media binding does not match final DOCX "
            f"(rid={auth.get('final_rid', '<missing>')} target={auth.get('final_target', '<missing>')})"
        )
    if final is not None and not _media_matches_authorized_final(final, auth):
        return (
            f"{auth.get('entry_key', '<unknown>')}: authorized replacement final media does not match manifest binding "
            f"(actual rid={final.get('rid', '<missing>')} target={final.get('target', '<missing>')})"
        )
    return None


def _find_authorization_for_final(
    final: dict[str, str],
    authorized: list[dict[str, str]],
    duplicate_final_identities: set[tuple[str, str, str]],
    issues: list[str],
) -> dict[str, str] | None:
    candidates = [auth for auth in authorized if _media_matches_authorized_final(final, auth)]
    if not candidates:
        return None
    if _media_identity(final) in duplicate_final_identities and not any(auth.get("final_part") for auth in candidates):
        issues.append(
            "source-to-final media replacement authorization is ambiguous; final owner part is required: "
            f"rid={final.get('rid', '<missing>')} target={final.get('target', '<missing>')}"
        )
        return None
    return candidates[0]


def _manifest_authorized_drawing_changes(manifest: dict[str, Any]) -> list[dict[str, str]]:
    authorized: list[dict[str, str]] = []
    top_level_authorizations = manifest.get("drawing_authorizations")
    if isinstance(top_level_authorizations, list):
        for index, entry in enumerate(top_level_authorizations, start=1):
            if not isinstance(entry, dict):
                continue
            authorization = _entry_value(
                entry,
                (
                    "explicit_drawing_authorization_source",
                    "drawing_authorization_source",
                    "authorization_source",
                    "task_card",
                ),
            )
            if not authorization:
                continue
            authorized.append(
                {
                    "entry_key": f"drawing_authorizations[{index}]",
                    "original_drawing_sha256": _entry_value(
                        entry,
                        ("original_drawing_sha256", "source_drawing_sha256", "original_drawing_object_sha256"),
                    ).lower(),
                    "final_drawing_sha256": _entry_value(
                        entry,
                        ("final_drawing_sha256", "final_drawing_object_sha256"),
                    ).lower(),
                    "original_owner_part": _entry_value(
                        entry,
                        ("original_drawing_owner_part", "original_owner_part", "source_drawing_owner_part"),
                    ),
                    "final_owner_part": _entry_value(
                        entry,
                        ("final_drawing_owner_part", "final_owner_part"),
                    ),
                }
            )
    for collection_name in ("figures", "diagrams"):
        collection = manifest.get(collection_name)
        if not isinstance(collection, dict):
            continue
        for entry_key, entry in collection.items():
            if not isinstance(entry, dict):
                continue
            authorization = _entry_value(
                entry,
                (
                    "explicit_drawing_authorization_source",
                    "drawing_authorization_source",
                    "explicit_resize_authorization_source",
                    "resize_authorization_source",
                    "explicit_insertion_authorization_source",
                    "insertion_authorization_source",
                    "explicit_replacement_authorization_source",
                    "replacement_authorization_source",
                    "approved_figure_task_card",
                    "figure_task_card",
                    "task_card",
                ),
            )
            if not authorization:
                continue
            authorized.append(
                {
                    "entry_key": str(entry_key),
                    "original_drawing_sha256": _entry_value(
                        entry,
                        ("original_drawing_sha256", "source_drawing_sha256", "original_drawing_object_sha256"),
                    ).lower(),
                    "final_drawing_sha256": _entry_value(
                        entry,
                        ("final_drawing_sha256", "final_drawing_object_sha256"),
                    ).lower(),
                    "original_owner_part": _entry_value(
                        entry,
                        ("original_drawing_owner_part", "original_owner_part", "source_drawing_owner_part"),
                    ),
                    "final_owner_part": _entry_value(
                        entry,
                        ("final_drawing_owner_part", "final_owner_part"),
                    ),
                }
            )
    return authorized


def _drawing_non_media_changed(source: dict[str, Any], final: dict[str, Any]) -> bool:
    compared_fields = (
        "drawing_kind",
        "extent_signature",
        "relationship_ids",
    )
    return any(str(source.get(field, "")) != str(final.get(field, "")) for field in compared_fields)


def _drawing_auth_matches(source: dict[str, Any], final: dict[str, Any], auth: dict[str, str]) -> bool:
    if not auth.get("original_drawing_sha256") or not auth.get("final_drawing_sha256"):
        return False
    if auth["original_drawing_sha256"] != str(source.get("sha256", "")).lower():
        return False
    if auth["final_drawing_sha256"] != str(final.get("sha256", "")).lower():
        return False
    if auth.get("original_owner_part") and auth["original_owner_part"] != str(source.get("story_part", "")):
        return False
    if auth.get("final_owner_part") and auth["final_owner_part"] != str(final.get("story_part", "")):
        return False
    return True


def _drawing_removal_auth_matches(source: dict[str, Any], auth: dict[str, str]) -> bool:
    if not auth.get("original_drawing_sha256"):
        return False
    if auth["original_drawing_sha256"] != str(source.get("sha256", "")).lower():
        return False
    if auth.get("original_owner_part") and auth["original_owner_part"] != str(source.get("story_part", "")):
        return False
    return True


def _drawing_addition_auth_matches(final: dict[str, Any], auth: dict[str, str]) -> bool:
    if not auth.get("final_drawing_sha256"):
        return False
    if auth["final_drawing_sha256"] != str(final.get("sha256", "")).lower():
        return False
    if auth.get("final_owner_part") and auth["final_owner_part"] != str(final.get("story_part", "")):
        return False
    return True


def _drawing_relocation_identity(drawing: dict[str, Any]) -> tuple[str, str, str, str, str]:
    return (
        str(drawing.get("story_part", "")),
        str(drawing.get("relationship_ids", "")),
        str(drawing.get("media_signature", "")),
        str(drawing.get("extent_signature", "")),
        str(drawing.get("next_text", "")),
    )


def _drawing_media_hashes(drawing: dict[str, Any]) -> set[str]:
    result: set[str] = set()
    for item in str(drawing.get("media_signature") or "").split(";"):
        parts = item.split("|")
        if len(parts) >= 3 and parts[2].strip():
            result.add(parts[2].strip().lower())
    return result


def _drawing_is_figure_bound(drawing: dict[str, Any]) -> bool:
    return str(drawing.get("next_is_figure_caption") or "").strip().lower() in {"true", "1", "yes"}


def _drawing_hash_allowed_by_source_or_template(
    drawing: dict[str, Any],
    *,
    source_hashes: set[str],
    final_hashes: set[str],
    template_hashes: set[str],
) -> bool:
    hashes = _drawing_media_hashes(drawing)
    if not hashes:
        return False
    if hashes & (source_hashes | final_hashes | template_hashes):
        return True
    media_signature = str(drawing.get("media_signature") or "").lower()
    return "template-cover" in media_signature or "direct-donor" in media_signature


def validate_docx_drawing_object_authorization(
    source_docx: Path,
    final_docx: Path,
    manifest: dict[str, Any],
) -> list[str]:
    source_drawings = docx_drawing_object_manifest(source_docx)
    final_drawings = docx_drawing_object_manifest(final_docx)
    if not source_drawings and not final_drawings:
        return []
    source_role = str(manifest.get("source_docx_role") or manifest.get("source_role") or "").strip().lower()
    source_is_format_template = source_role in {"format_template", "format-template", "template", "school_template"}
    source_preserved_existing = source_role in {"source-preserved-existing-figures", "source_preserved_existing_figures"}
    authorized = _manifest_authorized_drawing_changes(manifest)
    issues: list[str] = []
    source_media_hashes = _media_sha_set(docx_image_relationship_manifest(source_docx))
    final_media_hashes = _media_sha_set(docx_image_relationship_manifest(final_docx))
    template_media_hashes = _manifest_template_media_sha_set(manifest)
    relocated_final_identities = {_drawing_relocation_identity(final) for final in final_drawings.values()}
    relocated_source_identities = {_drawing_relocation_identity(source) for source in source_drawings.values()}
    for key, source in sorted(source_drawings.items()):
        final = final_drawings.get(key)
        if final is None:
            if _is_header_footer_decorative_drawing(source):
                continue
            if source_is_format_template:
                continue
            if _drawing_relocation_identity(source) in relocated_final_identities:
                continue
            if source_preserved_existing and _drawing_hash_allowed_by_source_or_template(
                source,
                source_hashes=source_media_hashes,
                final_hashes=final_media_hashes,
                template_hashes=template_media_hashes,
            ):
                continue
            if any(_drawing_removal_auth_matches(source, auth) for auth in authorized):
                continue
            issues.append(
                "source-to-final drawing object removed without explicit drawing authorization: "
                f"part={source.get('story_part', '<missing>')} paragraph={source.get('paragraph_index', '<missing>')}"
            )
            continue
        if not _drawing_non_media_changed(source, final):
            continue
        if source_preserved_existing and _drawing_hash_allowed_by_source_or_template(
            final,
            source_hashes=source_media_hashes,
            final_hashes=final_media_hashes,
            template_hashes=template_media_hashes,
        ):
            continue
        if any(_drawing_auth_matches(source, final, auth) for auth in authorized):
            continue
        issues.append(
            "source-to-final drawing object changed without explicit drawing authorization: "
            f"part={source.get('story_part', '<missing>')} paragraph={source.get('paragraph_index', '<missing>')} "
            f"source_extent={source.get('extent_signature', '<missing>')} final_extent={final.get('extent_signature', '<missing>')}"
        )
    for key, final in sorted(final_drawings.items()):
        if key in source_drawings:
            continue
        if _is_header_footer_decorative_drawing(final):
            continue
        if source_is_format_template and not str(final.get("relationship_ids") or "").strip():
            continue
        if _drawing_relocation_identity(final) in relocated_source_identities:
            continue
        if source_preserved_existing and _drawing_hash_allowed_by_source_or_template(
            final,
            source_hashes=source_media_hashes,
            final_hashes=final_media_hashes,
            template_hashes=template_media_hashes,
        ) and (
            _drawing_is_figure_bound(final)
            or int(str(final.get("paragraph_index") or "0")) < 20
            or str(final.get("story_part") or "") != "word/document.xml"
        ):
            continue
        if any(_drawing_addition_auth_matches(final, auth) for auth in authorized):
            continue
        issues.append(
            "source-to-final drawing object added without explicit drawing authorization: "
            f"part={final.get('story_part', '<missing>')} paragraph={final.get('paragraph_index', '<missing>')}"
        )
    return issues


def _manifest_docx_binding_issues(
    manifest: dict[str, Any],
    *,
    manifest_path: Path | None,
    final_docx: Path,
    source_docx: Path | None,
    strict_required: bool,
) -> tuple[list[str], Path | None]:
    issues: list[str] = []
    final_path = resolve_manifest_path_value(
        manifest.get("final_docx_path")
        or manifest.get("output_docx_path")
        or manifest.get("reviewed_output_path")
        or manifest.get("final_manuscript_path"),
        manifest_path,
    )
    final_sha = str(manifest.get("final_docx_sha256") or manifest.get("output_docx_sha256") or "").strip().lower()
    if strict_required:
        if final_path is None:
            issues.append("figure manifest final_docx_path is required when validating final DOCX figure surfaces")
        elif final_path != final_docx.resolve():
            issues.append("figure manifest final_docx_path does not match final DOCX")
        if not final_sha:
            issues.append("figure manifest final_docx_sha256 is required when validating final DOCX figure surfaces")
        elif final_docx.exists() and final_sha != file_sha256(final_docx):
            issues.append("figure manifest final_docx_sha256 does not match final DOCX")
    elif final_path is not None and final_path != final_docx.resolve():
        issues.append("figure manifest final_docx_path does not match final DOCX")
    elif final_sha and final_docx.exists() and final_sha != file_sha256(final_docx):
        issues.append("figure manifest final_docx_sha256 does not match final DOCX")

    source_path = source_docx
    if source_path is None:
        source_path = resolve_manifest_path_value(
            manifest.get("source_docx_path")
            or manifest.get("original_docx_path")
            or manifest.get("source_manuscript_path"),
            manifest_path,
        )
    manifest_source_path = resolve_manifest_path_value(
        manifest.get("source_docx_path")
        or manifest.get("original_docx_path")
        or manifest.get("source_manuscript_path"),
        manifest_path,
    )
    source_sha = str(manifest.get("source_docx_sha256") or manifest.get("original_docx_sha256") or "").strip().lower()
    if strict_required:
        if manifest_source_path is None:
            issues.append("figure manifest source_docx_path is required when validating final DOCX figure surfaces")
        elif not manifest_source_path.exists():
            issues.append("figure manifest source_docx_path does not exist")
        if source_path is not None and manifest_source_path is not None and manifest_source_path != source_path.resolve():
            issues.append("figure manifest source_docx_path does not match supplied source DOCX")
        if not source_sha:
            issues.append("figure manifest source_docx_sha256 is required when validating final DOCX figure surfaces")
        elif source_path is not None and source_path.exists() and source_sha != file_sha256(source_path):
            issues.append("figure manifest source_docx_sha256 does not match source DOCX")
    elif manifest_source_path is not None and source_path is not None and manifest_source_path != source_path.resolve():
        issues.append("figure manifest source_docx_path does not match supplied source DOCX")
    elif source_sha and source_path is not None and source_path.exists() and source_sha != file_sha256(source_path):
        issues.append("figure manifest source_docx_sha256 does not match source DOCX")
    return issues, source_path


def validate_docx_media_replacement_authorization(
    source_docx: Path,
    final_docx: Path,
    manifest: dict[str, Any],
) -> list[str]:
    source_media = docx_image_relationship_manifest(source_docx)
    final_media = docx_image_relationship_manifest(final_docx)
    if not source_media and not final_media:
        return []
    source_role = str(manifest.get("source_docx_role") or manifest.get("source_role") or "").strip().lower()
    source_preserved_existing = source_role in {"source-preserved-existing-figures", "source_preserved_existing_figures"}
    source_hashes = _media_sha_set(source_media)
    final_hashes = _media_sha_set(final_media)
    template_hashes = _manifest_template_media_sha_set(manifest)
    authorized = _manifest_authorized_image_changes(manifest)
    duplicate_source_identities = _duplicate_identities(source_media)
    duplicate_final_identities = _duplicate_identities(final_media)
    issues: list[str] = []
    for info in source_media.values():
        if info.get("target_exists") == "false":
            issues.append(
                "source DOCX image relationship target missing from package: "
                f"part={info.get('source_part', '<missing>')} rid={info.get('rid', '<missing>')} "
                f"target={info.get('target', '<missing>')}"
            )
    for info in final_media.values():
        if info.get("target_exists") == "false":
            issues.append(
                "final DOCX image relationship target missing from package: "
                f"part={info.get('source_part', '<missing>')} rid={info.get('rid', '<missing>')} "
                f"target={info.get('target', '<missing>')}"
            )
    for key, source in sorted(source_media.items()):
        final = final_media.get(key)
        if final is None:
            if source_preserved_existing and _media_hash_allowed_by_source_or_template(
                source,
                source_hashes=source_hashes,
                final_hashes=final_hashes,
                template_hashes=template_hashes,
            ):
                continue
            auth = _find_authorization_for_source(source, authorized, duplicate_source_identities, issues)
            if auth is None:
                issues.append(
                    "source-to-final media relationship removed without explicit replacement authorization: "
                    f"part={source.get('source_part', '<missing>')} rid={source.get('rid', '<missing>')} "
                    f"target={source.get('target', '<missing>')}"
                )
            elif auth.get("change_type") == "removal":
                continue
            else:
                final_issue = _authorization_final_binding_issue(auth, None, final_media, duplicate_final_identities)
                if final_issue:
                    issues.append(final_issue)
            continue
        changed = (
            source.get("target", "") != final.get("target", "")
            or source.get("sha256", "").lower() != final.get("sha256", "").lower()
        )
        if not changed:
            continue
        if source_preserved_existing and _media_hash_allowed_by_source_or_template(
            final,
            source_hashes=source_hashes,
            final_hashes=final_hashes,
            template_hashes=template_hashes,
        ):
            continue
        auth = _find_authorization_for_source(source, authorized, duplicate_source_identities, issues)
        if auth is not None:
            final_issue = _authorization_final_binding_issue(auth, final, final_media, duplicate_final_identities)
            if final_issue:
                issues.append(final_issue)
            continue
        issues.append(
            "source-to-final media relationship changed without explicit replacement authorization: "
            f"part={source.get('source_part', '<missing>')} rid={source.get('rid', '<missing>')} "
            f"source_target={source.get('target', '<missing>')} final_target={final.get('target', '<missing>')}"
        )
    for key, final in sorted(final_media.items()):
        if key in source_media:
            continue
        if source_preserved_existing and _media_hash_allowed_by_source_or_template(
            final,
            source_hashes=source_hashes,
            final_hashes=final_hashes,
            template_hashes=template_hashes,
        ):
            continue
        auth = _find_authorization_for_final(final, authorized, duplicate_final_identities, issues)
        if auth is not None:
            final_issue = _authorization_final_binding_issue(auth, final, final_media, duplicate_final_identities)
            if final_issue:
                issues.append(final_issue)
            continue
        issues.append(
            "source-to-final media relationship added without explicit insertion authorization: "
            f"part={final.get('source_part', '<missing>')} rid={final.get('rid', '<missing>')} "
            f"target={final.get('target', '<missing>')}"
        )
    return issues


def _svg_payload_renderer_status(payload: bytes) -> tuple[bool, str]:
    if len(payload) < 512:
        return False, "svg payload is too small to contain a rendered diagram"
    text = payload.decode("utf-8", errors="replace").lower()
    if "text is not svg - cannot display" in text:
        return False, "svg contains draw.io fallback notice"
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        return False, f"svg parse error: {exc}"
    drawable_names = {"path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "image", "g"}
    for child in root.iter():
        if child is root:
            continue
        if child.tag.rsplit("}", 1)[-1].lower() in drawable_names:
            return True, "pass"
    return False, "svg contains no drawable child elements"


def docx_svg_primary_fallback_pairs(docx_path: Path) -> list[dict[str, str]]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            rel_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            doc_root = ET.fromstring(zf.read("word/document.xml"))
            package_names = set(zf.namelist())
            svg_status_by_target: dict[str, dict[str, str]] = {}
            for name in package_names:
                if not name.startswith("word/media/") or not name.lower().endswith(".svg"):
                    continue
                payload = zf.read(name)
                safe, issue = _svg_payload_renderer_status(payload)
                target = "media/" + Path(name).name
                svg_status_by_target[target] = {
                    "svg_package_part": name,
                    "svg_sha256": hashlib.sha256(payload).hexdigest(),
                    "svg_byte_length": str(len(payload)),
                    "svg_renderer_safe": "true" if safe else "false",
                    "svg_renderer_issue": issue,
                }
    except Exception:
        return []
    rel_targets = {
        rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
        for rel in rel_root.findall(f"{{{PR_NS}}}Relationship")
        if "image" in rel.attrib.get("Type", "")
    }
    pairs: list[dict[str, str]] = []
    for blip in doc_root.findall(f".//{{{A_NS}}}blip"):
        embed = blip.attrib.get(f"{{{R_NS}}}embed", "")
        target = rel_targets.get(embed, "")
        svg_blip = blip.find(f".//{{{ASVG_NS}}}svgBlip")
        svg_embed = svg_blip.attrib.get(f"{{{R_NS}}}embed", "") if svg_blip is not None else ""
        svg_target = rel_targets.get(svg_embed, "")
        if svg_embed:
            row = {
                "raster_rid": embed,
                "raster_target": target,
                "svg_rid": svg_embed,
                "svg_target": svg_target,
            }
            row.update(svg_status_by_target.get(svg_target, {}))
            pairs.append(row)
        elif target.lower().endswith(".svg"):
            row = {
                "raster_rid": "",
                "raster_target": "",
                "svg_rid": embed,
                "svg_target": target,
            }
            row.update(svg_status_by_target.get(target, {}))
            pairs.append(row)
    return pairs


def validate_structural_docx_svg_fallback(
    key: str,
    entry: dict[str, Any],
    svg_fallback_pairs: list[dict[str, str]],
) -> list[str]:
    svg_target = _entry_value(entry, ("docx_svg_media_target", "svg_docx_media_target", "docx_media_target"))
    svg_rid = _entry_value(entry, ("docx_svg_rid", "svg_rid", "rid"))
    if svg_target and not svg_target.lower().endswith(".svg"):
        svg_target = _entry_value(entry, ("docx_svg_media_target", "svg_docx_media_target"))
    if not svg_target and not svg_rid:
        return [f"{key}: structural figure missing final DOCX SVG relationship binding"]
    matching_pairs = [
        pair
        for pair in svg_fallback_pairs
        if (svg_rid and pair.get("svg_rid") == svg_rid)
        or (svg_target and pair.get("svg_target", "").lower() == svg_target.lower())
    ]
    if not matching_pairs:
        return [f"{key}: structural figure final DOCX has no SVG-primary/PNG-fallback drawing pair"]
    issues: list[str] = []
    for pair in matching_pairs:
        raster_target = pair.get("raster_target", "")
        if raster_target.lower().endswith((".png", ".jpg", ".jpeg")):
            if pair.get("svg_renderer_safe") == "false":
                issues.append(
                    f"{key}: structural figure SVG-primary relationship points to renderer-unsafe SVG "
                    f"(svg_rid={pair.get('svg_rid') or '<missing>'}, "
                    f"svg_target={pair.get('svg_target') or '<missing>'}, "
                    f"issue={pair.get('svg_renderer_issue') or '<missing>'})"
                )
                continue
            return []
        issues.append(
            f"{key}: structural figure SVG-primary pair lacks raster fallback relationship "
            f"(svg_rid={pair.get('svg_rid') or '<missing>'}, raster_target={raster_target or '<missing>'})"
        )
    return issues


def docx_paragraph_texts(docx_path: Path) -> list[str]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return []
    texts: list[str] = []
    for paragraph in root.findall(f".//{{{W_NS}}}p"):
        text = "".join(node.text or "" for node in paragraph.findall(f".//{{{W_NS}}}t")).strip()
        if text:
            texts.append(text)
    return texts


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(f".//{{{W_NS}}}t")).strip()


def _nearest_nonempty_paragraph_text(paragraphs: list[ET.Element], index: int, step: int, limit: int = 4) -> str:
    cursor = index + step
    checked = 0
    while 0 <= cursor < len(paragraphs) and checked < limit:
        text = _paragraph_text(paragraphs[cursor])
        if text:
            return text
        cursor += step
        checked += 1
    return ""


def _css_length_to_emu(value: str) -> int | None:
    match = re.match(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(cm|mm|in|pt|px)?\s*$", value or "", re.I)
    if not match:
        return None
    amount = float(match.group(1))
    raw_unit = match.group(2)
    if raw_unit is None and amount >= 1000 and float(amount).is_integer():
        # WPS sometimes serializes grouped VML child-shape dimensions as raw
        # EMUs inside CSS-like style attributes. Treating those as points
        # creates false huge header/footer drawing failures.
        return int(amount)
    unit = (raw_unit or "pt").lower()
    if unit == "cm":
        return int(round(amount * 360000))
    if unit == "mm":
        return int(round(amount * 36000))
    if unit == "in":
        return int(round(amount * 914400))
    if unit == "px":
        return int(round(amount * 9525))
    return int(round(amount * 12700))


def _vml_style_dimensions(style: str) -> tuple[int | None, int | None]:
    width: int | None = None
    height: int | None = None
    for part in (style or "").split(";"):
        if ":" not in part:
            continue
        key, raw_value = part.split(":", 1)
        key = key.strip().lower()
        value = raw_value.strip()
        if key == "width":
            width = _css_length_to_emu(value)
        elif key == "height":
            height = _css_length_to_emu(value)
    return width, height


def _paragraph_vml_extents(paragraph: ET.Element) -> list[dict[str, int]]:
    extents: list[dict[str, int]] = []
    for shape in paragraph.iter():
        if not (shape.tag.endswith("}shape") or shape.tag.endswith("}rect")):
            continue
        has_image = any(descendant.tag.endswith("}imagedata") for descendant in shape.iter())
        if not has_image:
            continue
        width, height = _vml_style_dimensions(shape.attrib.get("style", ""))
        width = width or _css_length_to_emu(shape.attrib.get("width", ""))
        height = height or _css_length_to_emu(shape.attrib.get("height", ""))
        if width or height:
            extents.append({"cx": int(width or 0), "cy": int(height or 0)})
    return extents


def _paragraph_extents(paragraph: ET.Element) -> list[dict[str, int]]:
    extents: list[dict[str, int]] = []
    for extent in paragraph.findall(f".//{{{WP_NS}}}extent"):
        try:
            cx = int(extent.attrib.get("cx", "0") or "0")
            cy = int(extent.attrib.get("cy", "0") or "0")
        except ValueError:
            continue
        extents.append({"cx": cx, "cy": cy})
    extents.extend(_paragraph_vml_extents(paragraph))
    return extents


def docx_text_width_emu(docx_path: Path) -> int | None:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return None
    sect_prs = root.findall(f".//{{{W_NS}}}sectPr")
    sect_pr = sect_prs[-1] if sect_prs else None
    if sect_pr is None:
        return 9026 * 635
    pg_sz = sect_pr.find(f"./{{{W_NS}}}pgSz")
    pg_mar = sect_pr.find(f"./{{{W_NS}}}pgMar")
    try:
        page_width = int(pg_sz.attrib.get(f"{{{W_NS}}}w", "11906") if pg_sz is not None else "11906")
    except ValueError:
        page_width = 11906
    try:
        left = int(pg_mar.attrib.get(f"{{{W_NS}}}left", "1440") if pg_mar is not None else "1440")
    except ValueError:
        left = 1440
    try:
        right = int(pg_mar.attrib.get(f"{{{W_NS}}}right", "1440") if pg_mar is not None else "1440")
    except ValueError:
        right = 1440
    return max(page_width - left - right, 1) * 635


def docx_text_height_emu(docx_path: Path) -> int | None:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return None
    sect_prs = root.findall(f".//{{{W_NS}}}sectPr")
    sect_pr = sect_prs[-1] if sect_prs else None
    if sect_pr is None:
        return (16838 - 1440 - 1440) * 635
    pg_sz = sect_pr.find(f"./{{{W_NS}}}pgSz")
    pg_mar = sect_pr.find(f"./{{{W_NS}}}pgMar")
    try:
        page_height = int(pg_sz.attrib.get(f"{{{W_NS}}}h", "16838") if pg_sz is not None else "16838")
    except ValueError:
        page_height = 16838
    try:
        top = int(pg_mar.attrib.get(f"{{{W_NS}}}top", "1440") if pg_mar is not None else "1440")
    except ValueError:
        top = 1440
    try:
        bottom = int(pg_mar.attrib.get(f"{{{W_NS}}}bottom", "1440") if pg_mar is not None else "1440")
    except ValueError:
        bottom = 1440
    return max(page_height - top - bottom, 1) * 635


def _paragraph_has_body_drawing(paragraph: ET.Element) -> bool:
    if paragraph_has_omml(paragraph):
        return False
    for node in paragraph.iter():
        if node.tag == f"{{{W_NS}}}drawing":
            return True
        if node.tag == f"{{{W_NS}}}object":
            return True
        if node.tag != f"{{{W_NS}}}pict":
            continue
        # WPS/Word may serialize TOC leaders and front-matter divider lines as
        # VML pict/line shapes with no image relationship. Those are protected
        # format decoration, not thesis figure holders.
        for descendant in node.iter():
            if descendant.tag.endswith("}imagedata"):
                if descendant.attrib.get(f"{{{R_NS}}}id") or descendant.attrib.get(f"{{{R_NS}}}embed"):
                    return True
    return False


def docx_body_figure_paragraphs(docx_path: Path) -> list[dict[str, Any]]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return []
    body = root.find(f".//{{{W_NS}}}body")
    if body is None:
        return []
    rows: list[dict[str, Any]] = []
    paragraphs = body.findall(f".//{{{W_NS}}}p")
    media_manifest = docx_image_relationship_manifest(docx_path)
    for paragraph_index, child in enumerate(paragraphs, start=1):
        text = _paragraph_text(child)
        has_drawing = _paragraph_has_body_drawing(child)
        is_caption = is_docx_figure_caption(text)
        if not text and not has_drawing:
            continue
        extents = _paragraph_extents(child)
        if has_drawing:
            rids = paragraph_image_relationship_ids(child)
            current_index = paragraph_index - 1
            previous_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, -1)
            next_text = _nearest_nonempty_paragraph_text(paragraphs, current_index, 1)
            media_targets = [
                str(media_manifest.get(f"word/document.xml#{rid}", {}).get("target", ""))
                for rid in rids
            ]
            if paragraph_is_formula_like_drawing(
                child,
                text=text,
                previous_text=previous_text,
                next_text=next_text,
                extents=extents,
                media_targets=media_targets,
            ):
                has_drawing = False
        rows.append(
            {
                "paragraph_index": paragraph_index,
                "text": text,
                "has_drawing": has_drawing,
                "is_caption": is_caption,
                "caption_incomplete": is_incomplete_docx_figure_caption(text) if is_caption else False,
                "extent_cx_emu_values": [item["cx"] for item in extents],
                "extent_cy_emu_values": [item["cy"] for item in extents],
                "max_extent_cx_emu": max((item["cx"] for item in extents), default=0),
                "max_extent_cy_emu": max((item["cy"] for item in extents), default=0),
                "front_matter_drawing": False,
            }
        )
    first_body_index = next(
        (
            int(row["paragraph_index"])
            for row in rows
            if row["text"] and BODY_HEADING_RE.match(str(row["text"]))
        ),
        None,
    )
    if first_body_index is not None:
        for row in rows:
            if row["has_drawing"] and not row["is_caption"] and int(row["paragraph_index"]) < first_body_index:
                row["front_matter_drawing"] = True
    return rows


def is_docx_figure_caption(text: str) -> bool:
    match = FIGURE_CAPTION_RE.match(text or "")
    if not match:
        return False
    title = FIGURE_CAPTION_TITLE_STRIP_RE.sub("", match.group("title") or "").strip()
    compact_title = re.sub(r"\s+", "", title).lower()
    if compact_title.startswith(("展示", "显示", "说明", "给出", "呈现", "反映", "是", "用于")):
        return False
    if len(title) > 80 and re.search(r"[。；;]", title):
        return False
    return True


def is_incomplete_docx_figure_caption(text: str) -> bool:
    match = FIGURE_CAPTION_RE.match(text or "")
    if not match:
        return False
    title = FIGURE_CAPTION_TITLE_STRIP_RE.sub("", match.group("title") or "").strip()
    return not title


def has_structural_signal(text: str) -> bool:
    lowered = (text or "").lower()
    return any(token.lower() in lowered for token in STRUCTURAL_DOCX_TOKENS if token)


def has_docx_figure_reference_signal(text: str) -> bool:
    """Detect figure/diagram context without treating ordinary words like 图像 as a figure reference."""
    value = text or ""
    lowered = value.lower()
    if is_docx_figure_caption(value):
        return True
    if re.search(rf"\u56fe\s*(?:\d+|[{CJK_NUMERAL_CLASS}]+)(?:[\-.\uff0d\uff0e]\s*(?:\d+|[{CJK_NUMERAL_CLASS}]+))?", value):
        return True
    if re.search(r"(?:\u5982|\u89c1|\u53c2\u89c1|\u8be6\u89c1|\u4e0a|\u4e0b)\s*\u56fe", value):
        return True
    if any(token in value for token in ("\u6d41\u7a0b\u56fe", "\u67b6\u6784\u56fe", "\u7ed3\u6784\u56fe", "ER\u56fe", "E-R\u56fe", "\u7528\u4f8b\u56fe", "\u65f6\u5e8f\u56fe")):
        return True
    return any(token in lowered for token in ("figure", "fig.", "diagram", "draw.io", "svg"))


def compact_for_required_figure(text: str) -> str:
    return re.sub(r"[\s\u3000:：,，.;；、\-_\u2010-\u2015]+", "", text or "").lower()


def text_has_any_compact(text: str, tokens: tuple[str, ...]) -> bool:
    compact = compact_for_required_figure(text)
    return any(compact_for_required_figure(token) in compact for token in tokens if token)


def docx_required_figure_rules(docx_path: Path) -> list[dict[str, Any]]:
    if not docx_path.exists():
        return []
    rows = docx_body_figure_paragraphs(docx_path)
    full_text = "\n".join(str(row["text"]) for row in rows if row["text"])
    required: list[dict[str, Any]] = []
    for rule in REQUIRED_FIGURE_RULES:
        trigger_any = tuple(str(item) for item in rule.get("trigger_any", ()))
        trigger_all = tuple(str(item) for item in rule.get("trigger_all", ()))
        any_hit = text_has_any_compact(full_text, trigger_any) if trigger_any else False
        all_hit = bool(trigger_all) and all(text_has_any_compact(full_text, (token,)) for token in trigger_all)
        if any_hit or all_hit:
            required.append(dict(rule))
    return required


def required_figure_caption_present(captions: list[str], rule: dict[str, Any]) -> bool:
    family = str(rule.get("family") or "").strip().lower()
    tokens = tuple(str(item) for item in rule.get("caption_tokens", ()))
    for caption in captions:
        if text_has_any_compact(caption, tokens):
            if family == "er" and not text_has_any_compact(caption, ("er", "e-r", "\u5b9e\u4f53\u5173\u7cfb", "\u6982\u5ff5\u6a21\u578b")):
                continue
            return True
    return False


def required_figure_manifest_entry(manifest: dict[str, Any], rule: dict[str, Any]) -> dict[str, Any] | None:
    required = manifest.get("required_figures")
    rule_id = str(rule.get("id") or "").strip()
    family = str(rule.get("family") or "").strip().lower()
    tokens = tuple(str(item) for item in rule.get("caption_tokens", ()))
    if isinstance(required, dict):
        direct = required.get(rule_id)
        if isinstance(direct, dict):
            return direct
        for entry in required.values():
            if not isinstance(entry, dict):
                continue
            entry_id = str(entry.get("id") or entry.get("requirement_id") or "").strip()
            entry_family = str(entry.get("family") or "").strip().lower()
            caption = str(entry.get("caption") or entry.get("expected_caption") or "")
            if entry_id == rule_id or (entry_family == family and text_has_any_compact(caption, tokens)):
                return entry
    return None


def manifest_has_required_diagram(manifest: dict[str, Any], rule: dict[str, Any]) -> bool:
    family = str(rule.get("family") or "").strip().lower()
    tokens = tuple(str(item) for item in rule.get("caption_tokens", ()))
    diagrams = manifest.get("diagrams") if isinstance(manifest.get("diagrams"), dict) else {}
    for entry in diagrams.values():
        if not isinstance(entry, dict):
            continue
        entry_family = str(entry.get("family") or entry.get("inferred_family") or "").strip().lower()
        caption = str(entry.get("caption") or "")
        if entry_family == family and text_has_any_compact(caption, tokens):
            return True
    return False


def required_figure_coverage_issues(final_docx: Path, manifest: dict[str, Any] | None = None) -> list[str]:
    rules = docx_required_figure_rules(final_docx)
    if not rules:
        return []
    rows = docx_body_figure_paragraphs(final_docx)
    captions = [str(row["text"]) for row in rows if row.get("is_caption")]
    issues: list[str] = []
    for rule in rules:
        rule_id = str(rule.get("id") or "required_figure")
        family = str(rule.get("family") or "").strip().lower()
        if not required_figure_caption_present(captions, rule):
            issues.append(f"required figure missing from final DOCX captions: {rule_id} ({rule.get('reason')})")
        if manifest is None:
            continue
        entry = required_figure_manifest_entry(manifest, rule)
        if entry is None:
            issues.append(f"required figure checklist missing manifest row: {rule_id}")
            continue
        entry_family = str(entry.get("family") or "").strip().lower()
        if entry_family != family:
            issues.append(f"required figure checklist row {rule_id} has wrong family: {entry_family or '<missing>'}")
        status = str(entry.get("status") or entry.get("verdict") or entry.get("final_verdict") or "").strip().lower()
        if not _verdict_is_pass(status):
            issues.append(f"required figure checklist row {rule_id} must have pass status/verdict")
        for field_name in ("task_card", "post_insertion_rendered_evidence", "final_docx_relationship_evidence"):
            value = str(entry.get(field_name) or "").strip()
            if not value:
                issues.append(f"required figure checklist row {rule_id} missing {field_name}")
                continue
            path = Path(value)
            if not path.exists():
                issues.append(f"required figure checklist row {rule_id} {field_name} file missing: {value}")
        if not manifest_has_required_diagram(manifest, rule):
            issues.append(f"required figure checklist row {rule_id} has no matching {family} diagram manifest entry")
    return issues


def docx_figure_surface_summary(docx_path: Path) -> dict[str, Any]:
    image_targets = docx_image_targets(docx_path)
    figure_paragraphs = docx_body_figure_paragraphs(docx_path)
    texts = [str(row["text"]) for row in figure_paragraphs if row["text"]]
    captions = [str(row["text"]) for row in figure_paragraphs if row["is_caption"]]
    incomplete_caption_rows = [row for row in figure_paragraphs if row["caption_incomplete"]]
    drawing_rows = [row for row in figure_paragraphs if row["has_drawing"] and not row.get("front_matter_drawing")]
    front_matter_drawing_rows = [row for row in figure_paragraphs if row.get("front_matter_drawing")]
    drawing_blocks: list[dict[str, Any]] = []
    orphan_drawing_blocks: list[dict[str, Any]] = []
    index = 0
    while index < len(figure_paragraphs):
        row = figure_paragraphs[index]
        if not row["has_drawing"] or row["is_caption"] or row.get("front_matter_drawing"):
            index += 1
            continue
        start_index = index
        end_index = index
        while (
            end_index + 1 < len(figure_paragraphs)
            and figure_paragraphs[end_index + 1]["has_drawing"]
            and not figure_paragraphs[end_index + 1]["is_caption"]
            and not figure_paragraphs[end_index + 1].get("front_matter_drawing")
        ):
            end_index += 1
        next_row = figure_paragraphs[end_index + 1] if end_index + 1 < len(figure_paragraphs) else None
        block = {
            "start_paragraph_index": row["paragraph_index"],
            "end_paragraph_index": figure_paragraphs[end_index]["paragraph_index"],
            "next_text": str(next_row["text"]) if next_row else "",
        }
        drawing_blocks.append(block)
        if next_row is None or not next_row["is_caption"]:
            orphan_drawing_blocks.append(block)
        index = end_index + 1
    orphan_caption_rows: list[dict[str, Any]] = []
    for pos, row in enumerate(figure_paragraphs):
        if not row["is_caption"]:
            continue
        prev_row = figure_paragraphs[pos - 1] if pos > 0 else None
        if prev_row is None or not prev_row["has_drawing"] or prev_row.get("front_matter_drawing"):
            orphan_caption_rows.append(row)
    structural_captions = [text for text in captions if has_structural_signal(text)]
    structural_contexts = [
        text for text in texts
        if has_structural_signal(text) and has_docx_figure_reference_signal(text)
    ][:12]
    return {
        "image_count": len(image_targets),
        "image_relationship_count": len(image_targets),
        "body_drawing_count": len(drawing_rows),
        "front_matter_drawing_count": len(front_matter_drawing_rows),
        "body_drawing_block_count": len(drawing_blocks),
        "figure_caption_count": len(captions),
        "incomplete_caption_count": len(incomplete_caption_rows),
        "orphan_drawing_block_count": len(orphan_drawing_blocks),
        "orphan_caption_count": len(orphan_caption_rows),
        "structural_caption_count": len(structural_captions),
        "structural_context_count": len(structural_contexts),
        "has_figure_surfaces": bool(drawing_blocks or captions),
        "has_structural_signals": bool(structural_captions or structural_contexts),
        "captions": captions[:12],
        "captions_all": captions,
        "incomplete_captions": [str(row["text"]) for row in incomplete_caption_rows[:12]],
        "orphan_drawing_blocks": orphan_drawing_blocks[:12],
        "orphan_captions": [
            {"paragraph_index": row["paragraph_index"], "text": row["text"]}
            for row in orphan_caption_rows[:12]
        ],
        "drawing_blocks": drawing_blocks[:12],
        "front_matter_drawings": [
            {"paragraph_index": row["paragraph_index"], "text": row["text"]}
            for row in front_matter_drawing_rows[:12]
        ],
        "structural_contexts": structural_contexts,
    }


def _front_matter_drawing_preserved_from_source(final_docx: Path, source_docx: Path | None, paragraph_index: int) -> bool:
    if source_docx is None or not source_docx.exists():
        return False
    source_drawings = docx_drawing_object_manifest(source_docx)
    final_drawings = docx_drawing_object_manifest(final_docx)
    prefix = f"word/document.xml#p{paragraph_index}#"
    final_rows = {key: value for key, value in final_drawings.items() if key.startswith(prefix)}
    if not final_rows:
        return False
    compared_fields = ("drawing_kind", "extent_signature", "relationship_ids", "media_signature")
    for key, final_row in final_rows.items():
        source_row = source_drawings.get(key)
        if source_row is None:
            return False
        if any(str(source_row.get(field, "")) != str(final_row.get(field, "")) for field in compared_fields):
            return False
    return True


def _front_matter_drawing_allowed_as_template_cover(
    final_docx: Path,
    paragraph_index: int,
    source_docx_role: str,
) -> bool:
    role = source_docx_role.strip().lower()
    if role not in {"source-preserved-existing-figures", "source_preserved_existing_figures"}:
        return False
    prefix = f"word/document.xml#p{paragraph_index}#"
    for key, drawing in docx_drawing_object_manifest(final_docx).items():
        if not key.startswith(prefix):
            continue
        media_signature = str(drawing.get("media_signature") or "").lower()
        if "template-cover" in media_signature or "direct-donor" in media_signature:
            return True
    return False



def _max_extent_cx_from_signature(extent_signature: str) -> int:
    max_cx = 0
    for item in extent_signature.split(";"):
        if "x" not in item:
            continue
        cx_text, _sep, _cy_text = item.partition("x")
        try:
            max_cx = max(max_cx, int(cx_text))
        except ValueError:
            continue
    return max_cx


def _max_extent_cy_from_signature(extent_signature: str) -> int:
    max_cy = 0
    for item in extent_signature.split(";"):
        if "x" not in item:
            continue
        _cx_text, cy_text = item.split("x", 1)
        try:
            max_cy = max(max_cy, int(cy_text))
        except ValueError:
            continue
    return max_cy


def _drawing_preserved_from_source(
    final_key: str,
    final_drawing: dict[str, Any],
    source_drawings: dict[str, dict[str, Any]],
) -> bool:
    return source_drawings.get(final_key) == final_drawing


def final_docx_figure_surface_issues(
    final_docx: Path,
    *,
    source_docx: Path | None = None,
    source_docx_role: str = "",
) -> list[str]:
    if not final_docx.exists():
        return []
    summary = docx_figure_surface_summary(final_docx)
    issues: list[str] = []
    issues.extend(required_figure_coverage_issues(final_docx))
    text_width_emu = docx_text_width_emu(final_docx)
    text_height_emu = docx_text_height_emu(final_docx)
    if text_width_emu:
        for row in docx_body_figure_paragraphs(final_docx):
            if not row.get("has_drawing"):
                continue
            if row.get("front_matter_drawing"):
                continue
            extent_cx = int(row.get("max_extent_cx_emu") or 0)
            if extent_cx > int(text_width_emu * 1.02):
                issues.append(
                    "final DOCX figure drawing exceeds available text width: "
                    f"paragraph {row['paragraph_index']} cx_emu={extent_cx} text_width_emu={text_width_emu}"
                )
            if text_height_emu:
                extent_cy = int(row.get("max_extent_cy_emu") or 0)
                if extent_cy > int(text_height_emu * 0.82):
                    issues.append(
                        "final DOCX figure drawing exceeds safe page height occupancy: "
                        f"paragraph {row['paragraph_index']} cy_emu={extent_cy} "
                        f"text_height_emu={text_height_emu} threshold=82%"
                    )
        source_drawings = (
            docx_drawing_object_manifest(source_docx)
            if source_docx is not None and source_docx.exists()
            else {}
        )
        for key, drawing in sorted(docx_drawing_object_manifest(final_docx).items()):
            story_part = str(drawing.get("story_part", ""))
            if story_part == "word/document.xml":
                continue
            extent_cx = _max_extent_cx_from_signature(str(drawing.get("extent_signature", "")))
            if extent_cx <= int(text_width_emu * 1.02):
                continue
            if _drawing_preserved_from_source(key, drawing, source_drawings):
                continue
            issues.append(
                "final DOCX non-body story drawing exceeds available text width: "
                f"part={story_part} paragraph={drawing.get('paragraph_index', '<missing>')} "
                f"cx_emu={extent_cx} text_width_emu={text_width_emu}"
            )
    if text_height_emu:
        source_drawings = (
            docx_drawing_object_manifest(source_docx)
            if source_docx is not None and source_docx.exists()
            else {}
        )
        for key, drawing in sorted(docx_drawing_object_manifest(final_docx).items()):
            story_part = str(drawing.get("story_part", ""))
            if story_part == "word/document.xml":
                continue
            extent_cy = _max_extent_cy_from_signature(str(drawing.get("extent_signature", "")))
            if extent_cy <= int(text_height_emu * 0.82):
                continue
            if _drawing_preserved_from_source(key, drawing, source_drawings):
                continue
            issues.append(
                "final DOCX non-body story drawing exceeds safe page height occupancy: "
                f"part={story_part} paragraph={drawing.get('paragraph_index', '<missing>')} "
                f"cy_emu={extent_cy} text_height_emu={text_height_emu} threshold=82%"
            )
    for drawing in summary["front_matter_drawings"]:
        paragraph_index = int(drawing["paragraph_index"])
        if _is_front_matter_signature_line_drawing(drawing):
            continue
        if (
            source_docx_role.strip().lower()
            in {"format_template", "format-template", "template", "school_template"}
            and not str(drawing.get("relationship_ids") or "").strip()
        ):
            continue
        if _front_matter_drawing_allowed_as_template_cover(final_docx, paragraph_index, source_docx_role):
            continue
        if _front_matter_drawing_preserved_from_source(final_docx, source_docx, paragraph_index):
            continue
        issues.append(
            "final DOCX contains a drawing before the first body chapter, which may indicate TOC/front-matter image insertion: "
            f"paragraph {drawing['paragraph_index']} text={drawing['text'] or '<image-only>'}"
        )
    if not summary["has_figure_surfaces"] and not summary["front_matter_drawings"]:
        return issues
    for caption in summary["incomplete_captions"]:
        issues.append(f"final DOCX figure caption has no descriptive name: {caption}")
    for block in summary["orphan_drawing_blocks"]:
        next_text = block.get("next_text") or "<end-of-document>"
        issues.append(
            "final DOCX body drawing lacks adjacent figure caption after "
            f"paragraph {block['start_paragraph_index']}: next={next_text}"
        )
    for caption in summary["orphan_captions"]:
        issues.append(
            "final DOCX figure caption is not immediately preceded by an image holder: "
            f"paragraph {caption['paragraph_index']} text={caption['text']}"
        )
    drawing_block_count = int(summary["body_drawing_block_count"])
    caption_count = int(summary["figure_caption_count"])
    if drawing_block_count != caption_count and (drawing_block_count or caption_count):
        issues.append(
            "final DOCX figure drawing/caption block count mismatch "
            f"(drawing_blocks={drawing_block_count}, figure_captions={caption_count})"
        )
    return issues


def final_docx_manifest_requirement_issues(final_docx: Path) -> list[str]:
    if not final_docx.exists():
        return []
    summary = docx_figure_surface_summary(final_docx)
    if not summary["has_figure_surfaces"] and not summary["image_count"] and not docx_required_figure_rules(final_docx):
        return []
    details = (
        f"images={summary['image_count']}, body_drawing_blocks={summary['body_drawing_block_count']}, "
        f"figure_captions={summary['figure_caption_count']}"
    )
    if summary["has_structural_signals"]:
        details += f", structural_signals={summary['structural_caption_count'] + summary['structural_context_count']}"
    return [f"final DOCX contains figure/image surfaces but no figure asset manifest was provided ({details})"]


def docx_figure_surfaces_preserved(source_docx: Path | None, final_docx: Path | None) -> bool:
    if source_docx is None or final_docx is None:
        return False
    if not source_docx.exists() or not final_docx.exists():
        return False
    source_drawings = docx_drawing_object_manifest(source_docx)
    final_drawings = docx_drawing_object_manifest(final_docx)
    source_signatures = Counter(_drawing_surface_preservation_signature(row) for row in source_drawings.values())
    final_signatures = Counter(_drawing_surface_preservation_signature(row) for row in final_drawings.values())
    return (
        docx_image_relationship_manifest(source_docx) == docx_image_relationship_manifest(final_docx)
        and source_signatures == final_signatures
    )


def manifest_entries_are_source_preserved(figures: dict[str, Any], diagrams: dict[str, Any]) -> bool:
    entries = [
        entry
        for collection in (figures, diagrams)
        for entry in collection.values()
        if isinstance(entry, dict)
    ]
    return bool(entries) and all(_is_preserved_existing_figure_entry(entry) for entry in entries)


def manifest_entries_are_mechanical_cad(figures: dict[str, Any]) -> bool:
    entries = [entry for entry in figures.values() if isinstance(entry, dict)]
    if not entries:
        return False
    for entry in entries:
        text = _entry_search_text(entry)
        if _is_preserved_existing_figure_entry(entry):
            continue
        mechanical_tokens = set(MECHANICAL_CAD_FAMILIES) | {
            "cad png",
            "drawing-package",
            "drawing package",
        }
        if not any(token in text for token in mechanical_tokens):
            return False
    return True


def _drawing_surface_preservation_signature(row: dict[str, Any]) -> tuple[str, ...]:
    return (
        str(row.get("story_part", "")),
        str(row.get("drawing_kind", "")),
        str(row.get("extent_signature", "")),
        str(row.get("relationship_ids", "")),
        str(row.get("media_signature", "")),
    )


def _is_header_footer_decorative_drawing(row: dict[str, Any]) -> bool:
    story_part = str(row.get("story_part", ""))
    if not (story_part.startswith("word/header") or story_part.startswith("word/footer")):
        return False
    if str(row.get("relationship_ids") or "").strip():
        return False
    if str(row.get("media_signature") or "").strip():
        return False
    drawing_kind = str(row.get("drawing_kind") or "")
    return "pict=" in drawing_kind and "pict=0" not in drawing_kind


def _is_front_matter_signature_line_drawing(row: dict[str, Any]) -> bool:
    text = str(row.get("text") or "")
    compact = re.sub(r"\s+", "", text)
    if "签名" not in compact:
        return False
    if not any(token in compact for token in ("论文作者", "指导教师", "教师确认", "作者")):
        return False
    if str(row.get("relationship_ids") or "").strip():
        return False
    if str(row.get("media_signature") or "").strip():
        return False
    drawing_kind = str(row.get("drawing_kind") or "")
    return not drawing_kind or "pict=" in drawing_kind or "shape=" in drawing_kind


def _entry_value(entry: dict[str, Any], aliases: tuple[str, ...]) -> str:
    for alias in aliases:
        value = entry.get(alias)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def _entry_search_text(entry: dict[str, Any]) -> str:
    return " ".join(
        str(entry.get(key) or "")
        for key in (
            "family",
            "source_kind",
            "declared_family",
            "inferred_family",
            "kind",
            "asset_type",
            "caption",
            "description",
            "title",
            "alt",
            "purpose",
        )
    ).lower()


def _is_preserved_existing_figure_entry(entry: dict[str, Any]) -> bool:
    text = " ".join(
        str(entry.get(key) or "")
        for key in (
            "family",
            "source_kind",
            "declared_family",
            "inferred_family",
            "kind",
            "asset_type",
            "mutation_intent",
            "figure_mutation_intent",
            "preservation_status",
        )
    ).strip().lower()
    return any(
        token in text
        for token in (
            "preserved-existing",
            "preserve_existing",
            "preserve-existing",
            "source-preserved",
            "source_preserved",
            "no_image_mutation",
            "no-image-mutation",
        )
    )


def _has_structural_forbidden_source(value: str) -> bool:
    compact = re.sub(r"[\s_\-]+", " ", (value or "").strip().lower())
    compact_no_space = compact.replace(" ", "")
    return any(
        token in compact or token.replace(" ", "") in compact_no_space
        for token in STRUCTURAL_FORBIDDEN_SOURCE_PATTERNS
    )


def _is_structural_manifest_entry(entry: dict[str, Any]) -> bool:
    if _is_preserved_existing_figure_entry(entry):
        return False
    family_text = _entry_search_text(entry)
    declared_fields = {
        str(entry.get(key) or "").strip().lower()
        for key in ("family", "source_kind", "declared_family", "inferred_family", "kind", "asset_type")
        if str(entry.get(key) or "").strip()
    }
    structural_kinds = {"structural", "structure", "diagram", "drawio", "draw.io"}
    if declared_fields & (STRUCTURAL_FAMILIES | structural_kinds):
        return True
    if any(marker in family_text for marker in ("draw.io", "drawio", "structural figure")):
        return True
    if any(token.lower() in family_text for token in RUNTIME_SCREENSHOT_TOKENS):
        return False
    if any(token.lower() in family_text for token in STRUCTURAL_DOCX_TOKENS):
        return True
    return bool(_entry_value(entry, ("drawio", "drawio_path", "svg", "svg_path")))


def _matching_diagram_entry(
    figure_key: str,
    figure: dict[str, Any],
    diagrams: dict[str, Any],
) -> dict[str, Any] | None:
    figure_id = str(figure.get("id") or figure_key or "").strip().lower()
    figure_caption = compact_for_required_figure(str(figure.get("caption") or ""))
    for diagram_key, candidate in diagrams.items():
        if not isinstance(candidate, dict):
            continue
        candidate_id = str(candidate.get("id") or diagram_key or "").strip().lower()
        if figure_id and candidate_id == figure_id:
            return candidate
        candidate_caption = compact_for_required_figure(str(candidate.get("caption") or ""))
        if figure_caption and candidate_caption == figure_caption:
            return candidate
    return None


def _is_runtime_screenshot_entry(entry: dict[str, Any]) -> bool:
    family_text = _entry_search_text(entry)
    runtime_markers = {"runtime-screenshot", "runtime screenshot", "ui-screenshot", "page-screenshot"}
    return (
        any(marker in family_text for marker in runtime_markers)
        or any(token.lower() in family_text for token in RUNTIME_SCREENSHOT_TOKENS)
        or (
            "screenshot" in family_text
            and not any(token in family_text for token in ("code screenshot", "code-screenshot"))
        )
        or any(
            str(entry.get(key) or "").strip()
            for key in ("real_route", "route", "page_url", "screenshot_path")
        )
    )


def _is_algorithm_result_entry(entry: dict[str, Any]) -> bool:
    if _is_preserved_existing_figure_entry(entry):
        return False
    family_text = _entry_search_text(entry)
    return any(marker in family_text for marker in ALGORITHM_RESULT_FAMILIES) or any(
        token.lower() in family_text for token in ALGORITHM_RESULT_TOKENS
    )


def _entry_has_synthetic_result_marker(entry: dict[str, Any]) -> bool:
    text = " ".join(str(value or "") for value in entry.values()).lower()
    return any(token.lower() in text for token in SYNTHETIC_RESULT_TOKENS)


def _coerce_float(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _dimension_pair_from_dict(value: dict[str, Any]) -> tuple[int, int] | None:
    width = None
    height = None
    for key in ("width", "w", "client_width", "viewport_width", "image_width"):
        width = _coerce_float(value.get(key))
        if width is not None:
            break
    for key in ("height", "h", "client_height", "viewport_height", "image_height"):
        height = _coerce_float(value.get(key))
        if height is not None:
            break
    if width is None or height is None:
        x1 = _coerce_float(value.get("x"))
        if x1 is None:
            x1 = _coerce_float(value.get("left"))
        y1 = _coerce_float(value.get("y"))
        if y1 is None:
            y1 = _coerce_float(value.get("top"))
        x2 = _coerce_float(value.get("right"))
        if x2 is None:
            x2 = _coerce_float(value.get("x2"))
        y2 = _coerce_float(value.get("bottom"))
        if y2 is None:
            y2 = _coerce_float(value.get("y2"))
        if x1 is not None and y1 is not None and x2 is not None and y2 is not None:
            width = max(0.0, x2 - x1)
            height = max(0.0, y2 - y1)
    if width is None or height is None or width <= 0 or height <= 0:
        return None
    return int(round(width)), int(round(height))


def _dimension_pair_from_value(value: object) -> tuple[int, int] | None:
    if value is None or value == "":
        return None
    if isinstance(value, dict):
        return _dimension_pair_from_dict(value)
    if isinstance(value, (list, tuple)):
        numbers = [_coerce_float(part) for part in value]
        numbers = [part for part in numbers if part is not None]
        if len(numbers) >= 4:
            width = numbers[2] - numbers[0] if numbers[2] > numbers[0] and numbers[3] > numbers[1] else numbers[2]
            height = numbers[3] - numbers[1] if numbers[2] > numbers[0] and numbers[3] > numbers[1] else numbers[3]
            if width > 0 and height > 0:
                return int(round(width)), int(round(height))
        if len(numbers) >= 2 and numbers[0] > 0 and numbers[1] > 0:
            return int(round(numbers[0])), int(round(numbers[1]))
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        parsed = None
    if isinstance(parsed, (dict, list, tuple)):
        parsed_pair = _dimension_pair_from_value(parsed)
        if parsed_pair:
            return parsed_pair
    width_match = re.search(r"(?:width|client_width|viewport_width|image_width|w)\s*[:=]\s*(\d+(?:\.\d+)?)", text, re.I)
    height_match = re.search(r"(?:height|client_height|viewport_height|image_height|h)\s*[:=]\s*(\d+(?:\.\d+)?)", text, re.I)
    if width_match and height_match:
        return int(round(float(width_match.group(1)))), int(round(float(height_match.group(1))))
    numbers = [float(match) for match in re.findall(r"\d+(?:\.\d+)?", text)]
    if len(numbers) >= 4:
        width = numbers[2] - numbers[0] if numbers[2] > numbers[0] and numbers[3] > numbers[1] else numbers[2]
        height = numbers[3] - numbers[1] if numbers[2] > numbers[0] and numbers[3] > numbers[1] else numbers[3]
        if width > 0 and height > 0:
            return int(round(width)), int(round(height))
    if len(numbers) >= 2 and numbers[0] > 0 and numbers[1] > 0:
        return int(round(numbers[0])), int(round(numbers[1]))
    return None


def _entry_declared_size(entry: dict[str, Any], aliases: tuple[str, ...]) -> tuple[int, int] | None:
    for alias in aliases:
        pair = _dimension_pair_from_value(entry.get(alias))
        if pair:
            return pair
    return None


def _entry_image_size(entry: dict[str, Any]) -> tuple[int, int] | None:
    pair = _entry_declared_size(entry, ("image_size", "captured_image_size", "screenshot_size"))
    if pair:
        return pair
    width = _coerce_float(entry.get("image_width"))
    height = _coerce_float(entry.get("image_height"))
    if width is not None and height is not None and width > 0 and height > 0:
        return int(round(width)), int(round(height))
    return None


def _image_size_from_file(path: Path) -> tuple[int, int] | None:
    try:
        with path.open("rb") as handle:
            header = handle.read(32)
    except OSError:
        return None
    if header.startswith(b"\x89PNG\r\n\x1a\n") and len(header) >= 24:
        return int.from_bytes(header[16:20], "big"), int.from_bytes(header[20:24], "big")
    if header[:2] == b"BM" and len(header) >= 26:
        return int.from_bytes(header[18:22], "little", signed=True), abs(int.from_bytes(header[22:26], "little", signed=True))
    try:
        from PIL import Image  # type: ignore

        with Image.open(path) as image:
            return int(image.width), int(image.height)
    except Exception:
        return None


def _image_visual_quality_issues(path: Path, label: str) -> list[str]:
    try:
        from PIL import Image, ImageStat  # type: ignore

        with Image.open(path) as image:
            rgb = image.convert("RGB")
            sample = rgb.resize((min(64, rgb.width), min(64, rgb.height)))
            pixels = list(sample.getdata())
            if not pixels:
                return [f"{label} rendered image appears blank or unreadable: {path}"]
            color_counts: dict[tuple[int, int, int], int] = {}
            purple_pixels = 0
            white_or_black_pixels = 0
            for red, green, blue in pixels:
                color_counts[(red, green, blue)] = color_counts.get((red, green, blue), 0) + 1
                if red >= 110 and blue >= 110 and green <= 95 and abs(red - blue) <= 80:
                    purple_pixels += 1
                if (red >= 248 and green >= 248 and blue >= 248) or (red <= 7 and green <= 7 and blue <= 7):
                    white_or_black_pixels += 1
            total = len(pixels)
            dominant_rgb, dominant_count = max(color_counts.items(), key=lambda item: item[1])
            dominant_ratio = dominant_count / total
            purple_ratio = purple_pixels / total
            white_black_ratio = white_or_black_pixels / total
            variance = sum(ImageStat.Stat(sample).var) / 3.0
            if purple_ratio >= PURPLE_PLACEHOLDER_DOMINANT_RATIO:
                return [
                    f"{label} rendered image appears to be a purple placeholder/color block "
                    f"(purple_ratio={purple_ratio:.2f}): {path}"
                ]
            if dominant_ratio >= MAX_BLANK_IMAGE_DOMINANT_RATIO and white_black_ratio >= MAX_BLANK_IMAGE_DOMINANT_RATIO:
                return [
                    f"{label} rendered image appears blank/near-empty "
                    f"(dominant_rgb={dominant_rgb}, dominant_ratio={dominant_ratio:.2f}, variance={variance:.2f}): {path}"
                ]
            if dominant_ratio >= MAX_SOLID_BLOCK_DOMINANT_RATIO and variance < MIN_IMAGE_VARIANCE:
                return [
                    f"{label} rendered image appears to be a large solid color block "
                    f"(dominant_rgb={dominant_rgb}, dominant_ratio={dominant_ratio:.2f}, variance={variance:.2f}): {path}"
                ]
    except Exception:
        return []
    return []


def _structural_raster_style_issues(path: Path, key: str) -> list[str]:
    try:
        from PIL import Image  # type: ignore

        with Image.open(path) as image:
            rgb = image.convert("RGB")
            sample = rgb.resize((min(96, rgb.width), min(96, rgb.height)))
            pixels = list(sample.getdata())
    except Exception:
        return []
    if not pixels:
        return [f"{key}: structural figure raster fallback appears unreadable: {path}"]
    total = len(pixels)
    dark_pixels = sum(1 for red, green, blue in pixels if max(red, green, blue) <= 36)
    light_pixels = sum(1 for red, green, blue in pixels if min(red, green, blue) >= 238)
    dark_ratio = dark_pixels / total
    light_ratio = light_pixels / total
    if dark_ratio >= 0.35 and light_ratio < 0.55:
        return [
            f"{key}: structural figure raster fallback has dark/black-dominant background "
            f"(dark_ratio={dark_ratio:.2f}, light_ratio={light_ratio:.2f}): {path}"
        ]
    return []


def validate_structural_raster_style_contract(key: str, entry: dict[str, Any]) -> list[str]:
    raster_value = _entry_value(entry, ("png", "raster_fallback", "path", "image_path"))
    if not raster_value:
        return []
    raster = Path(raster_value)
    if not raster.exists():
        return []
    return _structural_raster_style_issues(raster, key)


FORBIDDEN_RUNTIME_CAPTURE_SUBSTRINGS = (
    "real_pyqt_widget_runtime_capture",
    "widget.render",
    "widget grab",
    "widget.grab",
    "qwidget.grab",
    "qwidget render",
    "qpixmap.grabwidget",
    "offscreen widget",
    "component snapshot",
    "control snapshot",
    "canvas.todataurl",
    "canvas export",
)


def _runtime_capture_method_issues(key: str, entry: dict[str, Any]) -> list[str]:
    method = _entry_value(entry, ("capture_method", "method"))
    kind = _entry_value(entry, ("capture_kind", "capture_source_kind", "screenshot_capture_kind"))
    combined = f"{method} {kind}".lower().replace("_", " ")
    compact = combined.replace(" ", "")
    issues: list[str] = []
    for token in FORBIDDEN_RUNTIME_CAPTURE_SUBSTRINGS:
        token_lower = token.lower()
        token_compact = token_lower.replace("_", " ").replace(" ", "")
        if token_lower in combined or token_compact in compact:
            issues.append(
                f"{key}: runtime screenshot capture method is a widget/component render substitute "
                f"({token})"
            )
            break
    return issues


def validate_runtime_screenshot_entry(key: str, entry: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    route = _entry_value(entry, ("real_route", "route", "page_url", "pageURL", "url"))
    capture_method = _entry_value(entry, ("capture_method", "method"))
    readiness_cue = _entry_value(entry, ("readiness_cue", "ready_selector", "wait_for", "readiness"))
    accepted_path = _entry_value(entry, ("accepted_screenshot_path", "screenshot_path", "path", "png", "raster_fallback"))
    caption_mapping = _entry_value(
        entry,
        ("caption_to_asset_mapping", "caption_to_asset", "caption_asset_map", "caption_asset_mapping"),
    )
    if not route:
        issues.append(f"{key}: runtime screenshot missing real route/page URL")
    if not capture_method:
        issues.append(f"{key}: runtime screenshot missing capture method")
    else:
        issues.extend(_runtime_capture_method_issues(key, entry))
    if not readiness_cue:
        issues.append(f"{key}: runtime screenshot missing readiness cue")
    if not accepted_path:
        issues.append(f"{key}: runtime screenshot missing accepted screenshot path")
    else:
        screenshot_path = Path(accepted_path)
        if not screenshot_path.exists():
            issues.append(f"{key}: runtime screenshot accepted path missing file: {accepted_path}")
        elif screenshot_path.suffix.lower() not in IMAGE_EXTENSIONS:
            issues.append(f"{key}: runtime screenshot accepted path is not an image: {accepted_path}")
        else:
            actual_size = _image_size_from_file(screenshot_path)
            if actual_size is None:
                issues.append(f"{key}: runtime screenshot image dimensions could not be read: {accepted_path}")
            else:
                issues.extend(_image_visual_quality_issues(screenshot_path, f"{key}: runtime screenshot"))
                actual_width, actual_height = actual_size
                declared_image_size = _entry_image_size(entry)
                if declared_image_size and (
                    abs(declared_image_size[0] - actual_width) > 2 or abs(declared_image_size[1] - actual_height) > 2
                ):
                    issues.append(
                        f"{key}: runtime screenshot declared image size does not match file "
                        f"(declared={declared_image_size[0]}x{declared_image_size[1]}, actual={actual_width}x{actual_height})"
                    )
                if actual_width < MIN_RUNTIME_SCREENSHOT_WIDTH or actual_height < MIN_RUNTIME_SCREENSHOT_HEIGHT:
                    issues.append(
                        f"{key}: runtime screenshot image is too small for full-window evidence "
                        f"({actual_width}x{actual_height}, minimum={MIN_RUNTIME_SCREENSHOT_WIDTH}x{MIN_RUNTIME_SCREENSHOT_HEIGHT})"
                    )
                expected_size = _entry_declared_size(
                    entry,
                    (
                        "expected_window_size",
                        "client_size",
                        "viewport_size",
                        "window_size",
                        "capture_bbox",
                        "window_bbox",
                        "application_window_bbox",
                    ),
                )
                if expected_size is None:
                    issues.append(
                        f"{key}: runtime screenshot missing full-window geometry "
                        "(expected_window_size/client_size/viewport_size or capture/window bbox)"
                    )
                else:
                    expected_width, expected_height = expected_size
                    if expected_width < MIN_RUNTIME_SCREENSHOT_WIDTH or expected_height < MIN_RUNTIME_SCREENSHOT_HEIGHT:
                        issues.append(
                            f"{key}: runtime screenshot declared window is too small for thesis full-window evidence "
                            f"({expected_width}x{expected_height})"
                        )
                    min_width = int(round(expected_width * MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE))
                    min_height = int(round(expected_height * MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE))
                    if actual_width < min_width or actual_height < min_height:
                        issues.append(
                            f"{key}: runtime screenshot appears cropped or partial-window "
                            f"(image={actual_width}x{actual_height}, expected_window={expected_width}x{expected_height}, "
                            f"required_coverage={MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE:.0%})"
                        )
    if not caption_mapping:
        issues.append(f"{key}: runtime screenshot missing caption-to-asset mapping")
    if not _verdict_is_pass(
        entry.get("full_window_capture_verdict")
        or entry.get("full_page_capture_verdict")
        or entry.get("capture_geometry_verdict")
    ):
        issues.append(f"{key}: runtime screenshot full-window capture verdict must be pass")
    coverage = _coerce_float(
        entry.get("full_window_coverage_ratio")
        or entry.get("capture_coverage_ratio")
        or entry.get("coverage_ratio")
    )
    if coverage is None:
        issues.append(f"{key}: runtime screenshot missing full-window coverage ratio")
    elif coverage < MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE:
        issues.append(
            f"{key}: runtime screenshot full-window coverage ratio is too low "
            f"({coverage:.2f}, minimum={MIN_RUNTIME_SCREENSHOT_WINDOW_COVERAGE:.2f})"
        )
    return issues


def validate_algorithm_result_entry(key: str, entry: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    accepted_path = _entry_value(
        entry,
        (
            "accepted_result_image_path",
            "result_image_path",
            "output_image_path",
            "accepted_screenshot_path",
            "screenshot_path",
            "path",
            "png",
            "raster_fallback",
        ),
    )
    caption_mapping = _entry_value(
        entry,
        ("caption_to_asset_mapping", "caption_to_asset", "caption_asset_map", "caption_asset_mapping"),
    )
    provenance_fields = (
        "source_image_path",
        "input_image_path",
        "generation_script_path",
        "inference_script_path",
        "program_script_path",
        "model_output_log_path",
        "inference_log_path",
        "result_log_path",
        "algorithm_provenance_source",
        "result_source",
        "existing_result_source",
        "user_provided_asset_evidence",
        "real_program_run_evidence",
    )
    provenance_values = [_entry_value(entry, (field_name,)) for field_name in provenance_fields]
    has_provenance = any(provenance_values)
    if not accepted_path:
        issues.append(f"{key}: algorithm result missing accepted result image path")
    else:
        result_path = Path(accepted_path)
        if not result_path.exists():
            issues.append(f"{key}: algorithm result accepted image path missing file: {accepted_path}")
        elif result_path.suffix.lower() not in IMAGE_EXTENSIONS:
            issues.append(f"{key}: algorithm result accepted path is not an image: {accepted_path}")
        elif _image_size_from_file(result_path) is None:
            issues.append(f"{key}: algorithm result image dimensions could not be read: {accepted_path}")
        else:
            issues.extend(_image_visual_quality_issues(result_path, f"{key}: algorithm result"))
    if not caption_mapping:
        issues.append(f"{key}: algorithm result missing caption-to-asset mapping")
    if not has_provenance:
        issues.append(f"{key}: algorithm result missing source/provenance evidence")
    if not _verdict_is_pass(
        entry.get("algorithm_authenticity_verdict")
        or entry.get("authenticity_verdict")
        or entry.get("result_provenance_verdict")
        or entry.get("algorithm_result_verdict")
    ):
        issues.append(f"{key}: algorithm result authenticity verdict must be pass")
    if _entry_has_synthetic_result_marker(entry) and not (
        has_provenance
        and _verdict_is_pass(
            entry.get("algorithm_authenticity_verdict")
            or entry.get("authenticity_verdict")
            or entry.get("result_provenance_verdict")
            or entry.get("algorithm_result_verdict")
        )
    ):
        issues.append(f"{key}: algorithm result appears to be a schematic/mock/placeholder without real provenance")
    return issues


def _validate_manifest_evidence_paths(value: object, field_name: str) -> list[str]:
    paths = _split_evidence_paths(value)
    if not paths:
        return [f"missing {field_name}"]
    issues: list[str] = []
    for path in paths:
        if not path.exists():
            issues.append(f"{field_name} missing file: {path}")
        elif path.is_file() and path.stat().st_size <= 0:
            issues.append(f"{field_name} file is empty: {path}")
    return issues


def validate_table_manifest_entries(tables: object) -> list[str]:
    if tables in (None, {}):
        return []
    if not isinstance(tables, dict):
        return ["figure manifest tables section must be an object when present"]
    issues: list[str] = []
    required_text_fields: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("table authority lock", ("table_authority_lock", "active_table_authority", "table_authority")),
        ("authority source type", ("authority_source_type", "table_authority_source_type")),
        ("title mode", ("title_mode", "caption_title_mode", "table_title_mode")),
    )
    required_pass_verdicts: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("border_family_verdict", ("border_family_verdict", "three_line_border_verdict", "table_border_family_verdict")),
        ("header_separator_verdict", ("header_separator_verdict", "table_header_separator_verdict", "header_bottom_middle_rule_verdict")),
        ("vertical_separator_verdict", ("vertical_separator_verdict", "table_vertical_separator_verdict")),
        ("body_row_separator_verdict", ("body_row_separator_verdict", "table_body_row_separator_verdict")),
        ("table_local_structure_verdict", ("table_local_structure_verdict", "local_structure_verdict")),
        ("table_pagination_verdict", ("table_pagination_verdict", "table_continuation_verdict", "table_row_split_header_repeat_verdict")),
    )
    required_path_fields: tuple[tuple[str, tuple[str, ...]], ...] = (
        (
            "table_authority_manuscript_binding_proof",
            (
                "table_authority_manuscript_binding_proof",
                "table_authority_binding_evidence",
                "table_to_manuscript_binding_evidence",
            ),
        ),
        (
            "rendered_table_evidence",
            (
                "rendered_table_evidence",
                "rendered_table_evidence_path",
                "rendered_table_page_evidence",
                "table_rendered_baseline_comparison_evidence",
                "table_rendered_baseline_comparison_evidence_path",
            ),
        ),
        (
            "final_docx_table_evidence",
            (
                "final_docx_table_evidence",
                "final_docx_table_evidence_path",
                "final_docx_relationship_evidence",
                "table_final_docx_binding_evidence",
                "table_audit_evidence_path",
            ),
        ),
    )
    for key, entry in tables.items():
        if not isinstance(entry, dict):
            issues.append(f"{key}: table manifest entry is not an object")
            continue
        present_keys = {str(field) for field, value in entry.items() if value not in (None, "", [], {})}
        if present_keys <= {"caption", "rows", "id", "chapter_title"}:
            issues.append(f"{key}: table manifest cannot pass with caption/row-count only; table authority, rendered evidence, and final DOCX binding are required")
        if not _entry_value(entry, ("caption", "title", "table_caption")):
            issues.append(f"{key}: table manifest entry missing caption/title")
        for label, aliases in required_text_fields:
            if not _entry_value(entry, aliases):
                issues.append(f"{key}: table manifest entry missing {label}")
        authority_path = _entry_value(
            entry,
            (
                "authority_source_file_path",
                "table_authority_source_path",
                "active_table_authority_path",
                "template_table_sample_path",
            ),
        )
        if authority_path:
            path = Path(authority_path)
            if not path.exists():
                issues.append(f"{key}: table authority source file missing: {authority_path}")
        elif not _verdict_is_pass(
            _entry_value(entry, ("no_template_table_authority_verdict", "table_authority_fallback_verdict"))
        ):
            issues.append(f"{key}: table manifest entry missing authority source file or pass no-template table authority verdict")
        for label, aliases in required_pass_verdicts:
            if not _verdict_is_pass(_entry_value(entry, aliases)):
                issues.append(f"{key}: table manifest {label} must be pass")
        if not _status_is_pass_or_inserted(_entry_value(entry, ("insertion_status", "table_insertion_status", "final_insertion_status"))):
            issues.append(f"{key}: table manifest insertion_status must be pass/inserted")
        if not _status_is_pass_or_inserted(_entry_value(entry, ("rendered_page_status", "table_rendered_page_status", "rendered_table_verdict"))):
            issues.append(f"{key}: table manifest rendered_page_status must be pass")
        for label, aliases in required_path_fields:
            issues.extend(f"{key}: {issue}" for issue in _validate_manifest_evidence_paths(_entry_value(entry, aliases), label))
    return issues


def validate_figure_manifest(
    manifest: dict[str, Any],
    *,
    final_docx: Path | None = None,
    source_docx: Path | None = None,
    manifest_path: Path | None = None,
) -> list[str]:
    manifest = manifest_with_resolved_paths(manifest, manifest_path)
    issues: list[str] = []
    if manifest.get("schema") != ASSET_SCHEMA:
        issues.append("figure manifest schema mismatch")
    diagrams = manifest.get("diagrams")
    if not isinstance(diagrams, dict):
        issues.append("figure manifest diagrams section missing")
        diagrams = {}
    figures = manifest.get("figures")
    if not isinstance(figures, dict):
        figures = {}
    issues.extend(validate_table_manifest_entries(manifest.get("tables")))
    svg_fallback_pairs = (
        docx_svg_primary_fallback_pairs(final_docx)
        if final_docx is not None and final_docx.exists()
        else []
    )
    for key, entry in figures.items():
        if not isinstance(entry, dict):
            issues.append(f"{key}: figure entry is not an object")
            continue
        issues.extend(validate_common_figure_entry_contract(key, entry))
        is_structural_entry = _is_structural_manifest_entry(entry)
        if is_structural_entry:
            issues.extend(validate_in_figure_language_contract(key, entry))
            issues.extend(validate_structural_raster_style_contract(key, entry))
        if is_structural_entry and _matching_diagram_entry(key, entry, diagrams) is None:
            issues.append(f"{key}: structural figure missing matching diagrams manifest entry")
            drawio_value = _entry_value(entry, ("drawio", "drawio_path"))
            svg_value = _entry_value(entry, ("svg", "svg_path"))
            png_value = _entry_value(entry, ("png", "raster_fallback", "path", "image_path"))
            if not drawio_value or not Path(drawio_value).exists():
                issues.append(f"{key}: structural figure missing draw.io source: {drawio_value or '<missing>'}")
            if not svg_value or not Path(svg_value).exists():
                issues.append(f"{key}: structural figure missing SVG export: {svg_value or '<missing>'}")
            if not png_value or not Path(png_value).exists():
                issues.append(f"{key}: structural figure missing raster fallback: {png_value or '<missing>'}")
        if is_structural_entry and final_docx is not None and final_docx.exists():
            issues.extend(validate_structural_docx_svg_fallback(key, entry, svg_fallback_pairs))
        elif _is_runtime_screenshot_entry(entry) and not _is_preserved_existing_figure_entry(entry):
            issues.extend(validate_runtime_screenshot_entry(key, entry))
        elif _is_algorithm_result_entry(entry):
            issues.extend(validate_algorithm_result_entry(key, entry))
    for key, entry in diagrams.items():
        if not isinstance(entry, dict):
            issues.append(f"{key}: diagram entry is not an object")
            continue
        issues.extend(validate_common_figure_entry_contract(key, entry))
        issues.extend(validate_in_figure_language_contract(key, entry))
        issues.extend(validate_structural_raster_style_contract(key, entry))
        family = str(entry.get("family") or "").strip().lower()
        inferred = str(entry.get("inferred_family") or "").strip().lower()
        declared = str(entry.get("declared_family") or "").strip().lower()
        if inferred == "flowchart" and declared and declared != "flowchart":
            issues.append(f"{key}: sequential figure declared as `{declared}` but inferred as `flowchart`")
        drawio_value = str(entry.get("drawio") or "")
        svg_value = str(entry.get("svg") or "")
        png_value = str(entry.get("png") or entry.get("raster_fallback") or "")
        drawio = Path(drawio_value) if drawio_value else None
        svg = Path(svg_value) if svg_value else None
        png = Path(png_value) if png_value else None
        if drawio is None or not drawio.exists():
            issues.append(f"{key}: structural figure missing draw.io source: {drawio_value or '<missing>'}")
        if svg is None or not svg.exists():
            issues.append(f"{key}: structural figure missing SVG export: {svg_value or '<missing>'}")
        if png is None or not png.exists():
            issues.append(f"{key}: structural figure missing raster fallback: {png_value or '<missing>'}")
        if family == "flowchart" and drawio is not None and drawio.exists():
            issues.extend(f"{key}: {issue}" for issue in drawio_flowchart_issues(drawio))
            issues.extend(f"{key}: {issue}" for issue in drawio_structural_geometry_issues(drawio, family=family))
            geometry_report_value = str(
                entry.get("geometry_validation_report")
                or entry.get("geometry_report")
                or entry.get("source_geometry_report")
                or ""
            )
            geometry_report = Path(geometry_report_value) if geometry_report_value else None
            if geometry_report is None:
                issues.append(f"{key}: flowchart structural figure missing geometry validation report")
            else:
                issues.extend(f"{key}: {issue}" for issue in validate_non_er_structural_geometry_report(geometry_report, drawio))
            for field_name in (
                "source_scale_bbox_map",
                "inserted_scale_geometry_evidence",
                "post_insertion_rendered_evidence",
                "final_docx_relationship_evidence",
            ):
                issues.extend(f"{key}: {issue}" for issue in _validate_manifest_evidence_paths(entry.get(field_name), field_name))
            if not _verdict_is_pass(entry.get("collision_check_verdict")):
                issues.append(f"{key}: flowchart structural figure collision_check_verdict must be pass")
            if not _verdict_is_pass(entry.get("source_to_inserted_geometry_verdict")):
                issues.append(f"{key}: flowchart structural figure source_to_inserted_geometry_verdict must be pass")
            if not _status_is_pass_or_inserted(entry.get("rendered_page_status")):
                issues.append(f"{key}: flowchart structural figure rendered_page_status must be pass")
            if not _status_is_pass_or_inserted(entry.get("insertion_status")):
                issues.append(f"{key}: flowchart structural figure insertion_status must be pass/inserted")
        if family not in {"er", "flowchart"} and drawio is not None and drawio.exists():
            issues.extend(f"{key}: {issue}" for issue in drawio_structural_geometry_issues(drawio, family=family or "structure"))
        if family == "er" and drawio is not None and drawio.exists():
            issues.extend(f"{key}: {issue}" for issue in drawio_er_geometry_issues(drawio))
            geometry_report_value = str(
                entry.get("geometry_validation_report")
                or entry.get("geometry_report")
                or entry.get("source_geometry_report")
                or ""
            )
            geometry_report = Path(geometry_report_value) if geometry_report_value else None
            if geometry_report is None:
                issues.append(f"{key}: ER structural figure missing geometry validation report")
            else:
                issues.extend(f"{key}: {issue}" for issue in validate_structural_geometry_report(geometry_report, drawio))
            for field_name in (
                "source_scale_bbox_map",
                "inserted_scale_geometry_evidence",
                "dense_zone_crop_evidence",
            ):
                issues.extend(f"{key}: {issue}" for issue in _validate_evidence_paths(entry.get(field_name), field_name))
            collision_verdict = str(entry.get("collision_check_verdict") or "").strip().lower()
            if not _verdict_is_pass(collision_verdict):
                issues.append(f"{key}: ER structural figure collision_check_verdict must be pass")
            for field_name, label in (
                ("relation_attribute_collision_verdict", "relation_attribute_collision_verdict"),
                ("shape_overlap_verdict", "shape_overlap_verdict"),
                ("inserted_scale_collision_verdict", "inserted_scale_collision_verdict"),
                ("source_to_inserted_geometry_verdict", "source_to_inserted_geometry_verdict"),
            ):
                if not _verdict_is_pass(entry.get(field_name)):
                    issues.append(f"{key}: ER structural figure {label} must be pass")
        if svg is not None and svg.exists() and "Text is not SVG - cannot display" in svg.read_text(encoding="utf-8", errors="replace"):
            issues.append(f"{key}: SVG still contains draw.io fallback notice: {svg}")
        if final_docx is not None and final_docx.exists():
            issues.extend(validate_structural_docx_svg_fallback(key, entry, svg_fallback_pairs))
    if final_docx is not None and final_docx.exists() and diagrams:
        targets = docx_image_targets(final_docx)
        if not any(target.lower().endswith((".png", ".jpg", ".jpeg")) for target in targets):
            issues.append("final DOCX has structural figures but no raster-renderable image relationship")
    if final_docx is not None and final_docx.exists():
        summary = docx_figure_surface_summary(final_docx)
        strict_docx_binding_required = bool(
            summary["has_figure_surfaces"]
            or summary["image_count"]
            or figures
            or diagrams
            or manifest_has_image_mutation_intent(manifest)
        )
        binding_issues, source_docx = _manifest_docx_binding_issues(
            manifest,
            manifest_path=manifest_path,
            final_docx=final_docx,
            source_docx=source_docx,
            strict_required=strict_docx_binding_required,
        )
        issues.extend(binding_issues)
        source_preserved_figure_surfaces = docx_figure_surfaces_preserved(source_docx, final_docx)
        issues.extend(required_figure_coverage_issues(final_docx, manifest))
        issues.extend(
            final_docx_figure_surface_issues(
                final_docx,
                source_docx=source_docx,
                source_docx_role=str(manifest.get("source_docx_role") or manifest.get("source_role") or ""),
            )
        )
        if source_docx is not None and source_docx.exists():
            issues.extend(validate_docx_media_replacement_authorization(source_docx, final_docx, manifest))
            issues.extend(validate_docx_drawing_object_authorization(source_docx, final_docx, manifest))
        figures = manifest.get("figures") if isinstance(manifest.get("figures"), dict) else {}
        # In v2 manifests structural figures are listed in `figures` and then
        # detailed again in `diagrams`; do not double-count the same figure.
        manifest_figure_count = len(figures) if figures else len(diagrams)
        if summary["has_figure_surfaces"] and not diagrams and not figures:
            issues.append(
                "final DOCX contains figure/image surfaces but manifest has no figure or diagram entries "
                f"(images={summary['image_count']}, body_drawing_blocks={summary['body_drawing_block_count']}, "
                f"figure_captions={summary['figure_caption_count']})"
            )
        if (figures or diagrams) and not summary["has_figure_surfaces"] and not summary["image_count"]:
            issues.append("figure manifest contains entries but final DOCX has no figure/image surfaces")
        if summary["figure_caption_count"] and manifest_figure_count and manifest_figure_count != summary["figure_caption_count"]:
            issues.append(
                "final DOCX figure caption count does not match figure manifest entries "
                f"(manifest_entries={manifest_figure_count}, figure_captions={summary['figure_caption_count']})"
            )
        manifest_captions = [
            str(entry.get("caption") or "")
            for collection in (figures, diagrams)
            for entry in collection.values()
            if isinstance(entry, dict) and str(entry.get("caption") or "").strip()
        ]
        manifest_caption_keys = {compact_for_required_figure(caption) for caption in manifest_captions}
        docx_caption_source = summary.get("captions_all") or summary["captions"]
        docx_caption_keys = {compact_for_required_figure(str(caption)) for caption in docx_caption_source}
        for caption in sorted(docx_caption_keys - manifest_caption_keys):
            if caption:
                issues.append(f"final DOCX figure caption is missing from manifest caption mapping: {caption}")
        for caption in sorted(manifest_caption_keys - docx_caption_keys):
            if caption and docx_caption_keys:
                issues.append(f"figure manifest caption is not present in final DOCX captions: {caption}")
        authorized_display_extent_resize = manifest_has_authorized_display_extent_resize(manifest)
        if (
            summary["has_structural_signals"]
            and not diagrams
            and not source_preserved_figure_surfaces
            and not manifest_entries_are_source_preserved(figures, diagrams)
            and not manifest_entries_are_mechanical_cad(figures)
            and not authorized_display_extent_resize
        ):
            issues.append(
                "final DOCX contains structural figure signals but manifest has no diagram entries "
                f"(structural_signals={summary['structural_caption_count'] + summary['structural_context_count']})"
            )
    return issues


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def apply_svg_primary_to_docx(docx_path: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    diagrams = [entry for entry in (manifest.get("diagrams") or {}).values() if isinstance(entry, dict)]
    candidates = [
        (Path(str(entry.get("png") or entry.get("raster_fallback") or "")), Path(str(entry.get("svg") or "")))
        for entry in diagrams
    ]
    candidates = [(png, svg) for png, svg in candidates if png.exists() and svg.exists()]
    if not candidates:
        return {"docx": str(docx_path), "svg_primary_replacements": 0, "reason": "no structural svg candidates"}

    def next_rel_id(root: ET.Element) -> str:
        used = {rel.attrib.get("Id", "") for rel in root.findall(f"{{{PR_NS}}}Relationship")}
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"

    tmp = docx_path.with_suffix(docx_path.suffix + ".svgprimary.tmp")
    replacements = 0
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        rel_root = ET.fromstring(zin.read("word/_rels/document.xml.rels"))
        doc_root = ET.fromstring(zin.read("word/document.xml"))
        media_hashes: dict[str, str] = {}
        for name in zin.namelist():
            if name.startswith("word/media/") and Path(name).suffix.lower() in {".png", ".jpg", ".jpeg"}:
                media_hashes[name] = hashlib.sha256(zin.read(name)).hexdigest()
        candidate_hashes = {sha256(png): (png, svg) for png, svg in candidates}
        svg_overrides: list[str] = []
        svg_media_to_source: dict[str, Path] = {}
        rel_by_id = {rel.attrib.get("Id", ""): rel for rel in rel_root.findall(f"{{{PR_NS}}}Relationship")}
        rel_by_target = {rel.attrib.get("Target", ""): rel for rel in rel_root.findall(f"{{{PR_NS}}}Relationship")}
        svg_by_media_name: dict[str, str] = {}

        for blip in doc_root.findall(f".//{{{A_NS}}}blip"):
            embed = blip.attrib.get(f"{{{R_NS}}}embed", "")
            rel = rel_by_id.get(embed)
            if rel is None:
                continue
            target = rel.attrib.get("Target", "")
            media_name = "word/" + target if target.startswith("media/") else target
            if media_name not in media_hashes:
                continue
            matched = candidate_hashes.get(media_hashes[media_name])
            if not matched:
                continue

            _png, svg = matched
            svg_name = f"media/{Path(media_name).stem}.svg"
            svg_media_name = "word/" + svg_name
            svg_overrides.append("/word/" + svg_name)
            svg_media_to_source[svg_media_name] = svg

            svg_rel = rel_by_target.get(svg_name)
            if svg_rel is None:
                svg_rel = ET.Element(f"{{{PR_NS}}}Relationship")
                svg_rel.attrib["Id"] = next_rel_id(rel_root)
                svg_rel.attrib["Type"] = rel.attrib.get(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                )
                svg_rel.attrib["Target"] = svg_name
                rel_root.append(svg_rel)
                rel_by_id[svg_rel.attrib["Id"]] = svg_rel
                rel_by_target[svg_name] = svg_rel

            ext_lst = blip.find(f"{{{A_NS}}}extLst")
            if ext_lst is None:
                ext_lst = ET.SubElement(blip, f"{{{A_NS}}}extLst")
            svg_ext = None
            for ext in ext_lst.findall(f"{{{A_NS}}}ext"):
                if ext.find(f".//{{{ASVG_NS}}}svgBlip") is not None:
                    svg_ext = ext
                    break
            if svg_ext is None:
                svg_ext = ET.SubElement(ext_lst, f"{{{A_NS}}}ext")
                svg_ext.attrib["uri"] = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
            svg_blip = svg_ext.find(f"{{{ASVG_NS}}}svgBlip")
            if svg_blip is None:
                svg_blip = ET.SubElement(svg_ext, f"{{{ASVG_NS}}}svgBlip")
            svg_blip.attrib[f"{{{R_NS}}}embed"] = svg_rel.attrib["Id"]
            svg_by_media_name[media_name] = svg_rel.attrib["Id"]
            replacements += 1

        ET.register_namespace("", PR_NS)
        rel_xml = ET.tostring(rel_root, encoding="utf-8", xml_declaration=True)
        ET.register_namespace("a", A_NS)
        ET.register_namespace("asvg", ASVG_NS)
        ET.register_namespace("r", R_NS)
        doc_xml = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)
        content_root = ET.fromstring(zin.read("[Content_Types].xml"))
        existing_overrides = {item.attrib.get("PartName") for item in content_root.findall(f"{{{CT_NS}}}Override")}
        for part_name in sorted(set(svg_overrides)):
            if part_name in existing_overrides:
                continue
            override = ET.Element(f"{{{CT_NS}}}Override")
            override.attrib["PartName"] = part_name
            override.attrib["ContentType"] = "image/svg+xml"
            content_root.append(override)
        ET.register_namespace("", CT_NS)
        content_xml = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)

        written = set()
        for item in zin.infolist():
            if item.filename == "word/_rels/document.xml.rels":
                zout.writestr(item, rel_xml)
            elif item.filename == "word/document.xml":
                zout.writestr(item, doc_xml)
            elif item.filename == "[Content_Types].xml":
                zout.writestr(item, content_xml)
            elif item.filename in svg_media_to_source:
                zout.writestr(item, svg_media_to_source[item.filename].read_bytes())
            else:
                zout.writestr(item, zin.read(item.filename))
            written.add(item.filename)
        for new_media, source_svg in svg_media_to_source.items():
            if new_media in written:
                continue
            zout.writestr(new_media, source_svg.read_bytes())
    os.replace(str(tmp), str(docx_path))
    return {
        "docx": str(docx_path),
        "svg_primary_replacements": replacements,
        "svg_primary_mode": "a:blip PNG fallback with asvg:svgBlip primary",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--source-docx")
    parser.add_argument("--final-docx")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    issues = validate_figure_manifest(
        manifest,
        final_docx=Path(args.final_docx) if args.final_docx else None,
        source_docx=Path(args.source_docx) if args.source_docx else None,
        manifest_path=Path(args.manifest),
    )
    if args.json:
        print(json.dumps({"issues": issues}, ensure_ascii=False, indent=2))
    else:
        for issue in issues:
            print(issue)
    return 1 if issues else 0


if __name__ == "__main__":
    raise SystemExit(main())
