#!/usr/bin/env python3
"""Audit a mechanical drawing delivery package for CAD-source and density risks."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


MOJIBAKE_TOKEN_ESCAPES = (
    "\\ufffd",
    "\\u951f",
    "\\u934a",
    "\\u9357",
    "\\u9365",
    "\\u93ac",
    "\\u7039",
    "\\u93cd",
    "\\u95c1",
    "\\u95ba",
    "\\u599e",
    "\\u9225",
    "\\u9983",
    "\\u6924\\u572d\\u6d30",
    "\\u9352",
    "\\u9422",
    "\\u93c8",
    "\\u95c7",
    "\\u95cd",
    "\\u6d98",
    "\\u6d7c",
    "\\u95be",
    "\\u741b",
    "\\u6fe1",
    "\\u7f02",
    "\\u95ab",
)
MOJIBAKE_TOKENS = tuple(token.encode("ascii").decode("unicode_escape") for token in MOJIBAKE_TOKEN_ESCAPES)
MISSING_GLYPH_BOX_ESCAPES = (
    "\\u25a1",
    "\\u25a0",
    "\\u25a2",
    "\\u25a3",
    "\\u25ab",
    "\\u25ad",
    "\\u25af",
    "\\u25fb",
    "\\u25fc",
    "\\u25fd",
    "\\u25fe",
    "\\u2610",
    "\\u2b1c",
    "\\u2b1b",
    "\\uff1f\\uff1f",
)
MISSING_GLYPH_BOX_TOKENS = tuple(
    token.encode("ascii").decode("unicode_escape") for token in MISSING_GLYPH_BOX_ESCAPES
)

TEXT_SUFFIXES = {".dxf", ".json", ".txt", ".csv", ".md", ".xml", ".scr", ".lsp"}
FORMAL_CAD_SOURCE_SUFFIXES = {".dwg", ".dxf"}
DXF_TEXT_ENTITY_TOKENS = {"TEXT", "MTEXT", "ATTRIB", "ATTDEF"}
SCHEMATIC_SUBSTITUTE_TOKEN_ESCAPES = (
    "\\u793a\\u610f\\u56fe",  # schematic figure
    "\\u793a\\u610f",
    "\\u6982\\u5ff5\\u56fe",  # concept figure
    "\\u6982\\u5ff5",
    "\\u8349\\u56fe",  # sketch
    "\\u7b80\\u56fe",  # simplified drawing
    "\\u6837\\u4f8b\\u56fe",  # sample figure
    "\\u5360\\u4f4d\\u56fe",  # placeholder figure
)
SCHEMATIC_SUBSTITUTE_BASE_TOKENS = tuple(
    token.encode("ascii").decode("unicode_escape")
    for token in SCHEMATIC_SUBSTITUTE_TOKEN_ESCAPES
)
SCHEMATIC_SUBSTITUTE_BROAD_TOKENS = SCHEMATIC_SUBSTITUTE_BASE_TOKENS + (
    "schematic",
    "concept",
    "conceptual",
    "placeholder",
    "mockup",
    "mock",
    "sketch",
    "sample-figure",
    "sample figure",
    "draft-only",
)
SCHEMATIC_SUBSTITUTE_TEXT_TOKENS = SCHEMATIC_SUBSTITUTE_BASE_TOKENS + (
    "schematic figure",
    "schematic drawing",
    "concept figure",
    "concept drawing",
    "conceptual figure",
    "conceptual drawing",
    "placeholder figure",
    "placeholder drawing",
    "mockup figure",
    "mockup drawing",
    "mock figure",
    "mock drawing",
    "sketch figure",
    "sketch drawing",
    "sample figure",
    "sample drawing",
    "draft-only",
)
ENTITY_TOKENS = {
    "LINE",
    "LWPOLYLINE",
    "POLYLINE",
    "VERTEX",
    "SEQEND",
    "CIRCLE",
    "ARC",
    "TEXT",
    "MTEXT",
    "DIMENSION",
    "INSERT",
    "LEADER",
    "MLEADER",
    "TOLERANCE",
    "HATCH",
    "SOLID",
    "SPLINE",
    "ELLIPSE",
    "ATTRIB",
    "ATTDEF",
    "BLOCK",
    "ENDBLK",
}
STRICT_SECTION_TOKENS = ("TABLES", "BLOCKS", "LAYER", "STYLE", "DIMSTYLE", "LTYPE")
RENDERED_REVIEW_PASS_VALUES = {"pass", "passed", "yes", "true", "ok"}
ENTITY_COUNT_ONLY_REJECT_VALUES = {
    "no",
    "not-used",
    "not_used",
    "not-used-as-acceptance",
    "not_used_as_acceptance",
    "rejected",
    "blocked",
    "fail",
    "failed",
    "visual-reviewed",
    "visual_reviewed",
    "not-entity-only",
    "not_entity_only",
}
RENDERED_REVIEW_PASS_FIELDS = (
    "no_overlap_verdict",
    "boundary_clearance_verdict",
    "detail_density_verdict",
    "text_legibility_verdict",
    "text_integrity_verdict",
    "text_orientation_verdict",
    "title_block_table_notes_isolation_verdict",
    "title_block_cell_containment_verdict",
    "title_block_short_line_topology_verdict",
    "annotation_margin_clearance_verdict",
    "local_crowding_verdict",
    "sheet_layout_verdict",
    "manufacturing_view_depth_verdict",
    "layout_collision_verdict",
    "content_overlap_verdict",
    "outside_frame_ink_verdict",
    "inner_frame_safe_margin_verdict",
    "hatch_section_fill_clipping_verdict",
)
PER_SHEET_RENDERED_REVIEW_PASS_FIELDS = (
    "no_overlap_verdict",
    "boundary_clearance_verdict",
    "detail_density_verdict",
    "text_legibility_verdict",
    "text_integrity_verdict",
    "text_orientation_verdict",
    "title_block_table_notes_isolation_verdict",
    "title_block_cell_containment_verdict",
    "title_block_short_line_topology_verdict",
    "annotation_margin_clearance_verdict",
    "local_crowding_verdict",
    "layout_collision_verdict",
    "content_overlap_verdict",
    "outside_frame_ink_verdict",
    "inner_frame_safe_margin_verdict",
    "hatch_section_fill_clipping_verdict",
)
MACHINE_OVERLAP_COUNT_FIELDS = (
    "overlap_count",
    "text_entity_overlap_count",
    "reserved_zone_collision_count",
    "title_block_table_note_collision_count",
    "annotation_collision_count",
    "frame_clearance_violation_count",
)
MACHINE_OVERLAP_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "overlap_count",
    "text_entity_overlap_count",
    "reserved_zone_collision_count",
    "title_block_table_note_collision_count",
    "annotation_collision_count",
    "frame_clearance_violation_count",
    "min_clearance_mm",
    "max_local_ink_density",
    "per_sheet",
)
CONTENT_OVERLAP_COUNT_FIELDS = (
    "content_overlap_count",
    "view_view_overlap_count",
    "detail_frame_main_view_overlap_count",
    "table_text_grid_collision_count",
    "dimension_line_view_table_crossing_count",
    "leader_line_view_table_crossing_count",
    "balloon_geometry_collision_count",
    "bbox_helper_envelope_escape_count",
    "stale_rendered_preview_count",
)
CONTENT_OVERLAP_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "registered_bbox_count",
    "checked_pair_count",
    "min_clearance_mm",
    *CONTENT_OVERLAP_COUNT_FIELDS,
    "per_sheet",
)
TITLE_BLOCK_SHORT_LINE_TOPOLOGY_COUNT_FIELDS = (
    "missing_short_table_line_count",
    "broken_cell_border_count",
    "table_grid_topology_mismatch_count",
)
TITLE_BLOCK_SHORT_LINE_TOPOLOGY_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "baseline_source",
    *TITLE_BLOCK_SHORT_LINE_TOPOLOGY_COUNT_FIELDS,
    "diagnostic_overlay_free_title_block_crop_path",
    "per_sheet",
)
OUTSIDE_FRAME_INK_COUNT_FIELDS = (
    "outside_frame_independent_ink_component_count",
    "outside_frame_text_component_count",
    "outside_frame_leader_component_count",
    "outside_frame_hatch_section_component_count",
    "outside_frame_table_title_block_component_count",
)
OUTSIDE_FRAME_INK_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "frame_detection_method",
    "rendered_source_formats",
    "sheet_frame_bbox_count",
    *OUTSIDE_FRAME_INK_COUNT_FIELDS,
    "max_outside_component_area_px",
    "per_sheet",
)
INNER_FRAME_SAFE_MARGIN_COUNT_FIELDS = (
    "right_safe_boundary_intrusion_count",
    "leader_text_inner_frame_intrusion_count",
    "view_geometry_inner_frame_intrusion_count",
    "dimension_text_inner_frame_intrusion_count",
)
INNER_FRAME_SAFE_MARGIN_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "source_sections",
    "sheet_frame_bbox_count",
    "min_safe_margin_mm",
    *INNER_FRAME_SAFE_MARGIN_COUNT_FIELDS,
)
TEXT_LEGIBILITY_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "preview_dpi",
    "min_cad_text_height_mm",
    "min_required_cad_text_height_mm",
    "min_rendered_text_height_px",
    "min_required_rendered_text_height_px",
    "median_rendered_text_height_px",
    "per_sheet",
)
TEXT_LEGIBILITY_MIN_CAD_HEIGHT_MM = 3.8
TEXT_LEGIBILITY_MIN_RENDERED_HEIGHT_PX = 40.0
TEXT_INTEGRITY_COUNT_FIELDS = (
    "mojibake_or_missing_glyph_count",
    "missing_required_drawing_text_count",
)
TEXT_INTEGRITY_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "source_text_entity_count",
    *TEXT_INTEGRITY_COUNT_FIELDS,
    "per_sheet",
)
TEXT_ORIENTATION_COUNT_FIELDS = (
    "upside_down_text_count",
    "mirrored_text_count",
)
TEXT_ORIENTATION_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "source_text_entity_count",
    *TEXT_ORIENTATION_COUNT_FIELDS,
    "per_sheet",
)
MANUFACTURING_COMPLEXITY_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "min_feature_family_count",
    "uses_dense_filler_as_primary_depth",
    "per_sheet",
)
MANUFACTURING_COMPLEXITY_MIN_FEATURE_FAMILIES = 8
HATCH_CLIP_COUNT_FIELDS = (
    "hatch_clip_violation_count",
    "entity_boundary_escape_count",
    "adjacent_view_crossing_count",
    "dimension_line_crossing_count",
    "title_block_table_bom_frame_crossing_count",
    "blank_background_leak_count",
)
HATCH_CLIP_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "section_fill_region_count",
    *HATCH_CLIP_COUNT_FIELDS,
    "per_sheet",
)

CELL_CONTAINMENT_COUNT_FIELDS = (
    "outside_cell_count",
    "border_touch_count",
    "unowned_table_text_count",
    "clipped_overflow_count",
    "cell_padding_violation_count",
)
CELL_CONTAINMENT_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "table_cell_text_count",
    *CELL_CONTAINMENT_COUNT_FIELDS,
    "per_sheet",
)

ANNOTATION_OWNERSHIP_COUNT_FIELDS = (
    "unowned_free_text_count",
    "unsupported_floating_text_count",
    "unbound_scattered_text_count",
    "dimension_like_text_without_anchor_count",
)
ANNOTATION_OWNERSHIP_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    "annotation_text_count",
    "boxed_callout_count",
    "symbol_or_title_text_count",
    *ANNOTATION_OWNERSHIP_COUNT_FIELDS,
    "per_sheet",
)

RESERVED_ZONE_INTRUSION_COUNT_FIELDS = (
    "intrusion_count",
    "geometry_title_block_intrusion_count",
    "dimension_line_table_zone_intrusion_count",
    "dimension_text_table_zone_intrusion_count",
    "leader_table_zone_intrusion_count",
    "view_geometry_reserved_zone_intrusion_count",
    "bom_table_intrusion_count",
    "technical_requirement_intrusion_count",
    "protected_table_zone_intrusion_count",
    "view_geometry_table_zone_intrusion_count",
    "detail_view_table_zone_intrusion_count",
    "leader_balloon_table_zone_intrusion_count",
    "dimension_table_zone_intrusion_count",
    "title_block_bom_protected_zone_intrusion_count",
)
RESERVED_ZONE_INTRUSION_REQUIRED_TOP_FIELDS = (
    "audit_method",
    "passed",
    *RESERVED_ZONE_INTRUSION_COUNT_FIELDS,
    "per_sheet",
)

FRAME_OVERFLOW_COMPONENT_MIN_AREA_PX = 18
FRAME_OVERFLOW_FRAME_TOLERANCE_PX = 3
FRAME_OVERFLOW_OUTER_MARGIN_MM = 10.0
FRAME_OVERFLOW_COUNT_FIELDS = (
    "outside_frame_component_count",
    "outside_frame_pixel_count",
    "max_outside_component_area_px",
)
MIN_REFERENCE_DWG_RATIO_WITH_REFERENCE = 0.35
INK_CONTRAST_MIN_READABLE_RATIO = 0.78
INK_CONTRAST_MIN_INK_PIXEL_RATIO = 0.0015
INK_CONTRAST_NONWHITE_DISTANCE = 15.0
INK_CONTRAST_READABLE_DISTANCE = 80.0


@dataclass(frozen=True)
class PackageFile:
    name: str
    data: bytes

    @property
    def suffix(self) -> str:
        return Path(self.name).suffix.lower()

    @property
    def size(self) -> int:
        return len(self.data)


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    if path.is_file():
        digest.update(path.read_bytes())
    else:
        for child in sorted(p for p in path.rglob("*") if p.is_file()):
            digest.update(child.relative_to(path).as_posix().encode("utf-8"))
            digest.update(child.read_bytes())
    return digest.hexdigest().upper()


def load_package(path: Path) -> list[PackageFile]:
    if path.is_file() and path.suffix.lower() == ".zip":
        files: list[PackageFile] = []
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                if not info.is_dir():
                    files.append(PackageFile(info.filename, zf.read(info)))
        return files
    if path.is_dir():
        return [
            PackageFile(child.relative_to(path).as_posix(), child.read_bytes())
            for child in sorted(path.rglob("*"))
            if child.is_file()
        ]
    if path.is_file():
        return [PackageFile(path.name, path.read_bytes())]
    raise FileNotFoundError(path)


def decode_text(data: bytes) -> tuple[str, str, bool]:
    for encoding in ("utf-8-sig", "gb18030", "cp936", "latin-1"):
        try:
            return data.decode(encoding), encoding, False
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace", True


def text_mojibake_hits(name: str, text: str) -> list[str]:
    hits = [token for token in MOJIBAKE_TOKENS if token in name or token in text]
    return sorted(set(hits))


def missing_glyph_hits(name: str, text: str) -> list[str]:
    hits = [token for token in MISSING_GLYPH_BOX_TOKENS if token and (token in name or token in text)]
    return sorted(set(hits))


def text_quality_hits(name: str, text: str) -> list[str]:
    return sorted(set(text_mojibake_hits(name, text) + missing_glyph_hits(name, text)))


def _sample_text(value: str, limit: int = 72) -> str:
    text = " ".join(value.replace("\\P", " ").replace("\\~", " ").split())
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "..."


def _parse_float(value: str) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _parse_int(value: str) -> int | None:
    parsed = _parse_float(value)
    if parsed is None:
        return None
    return int(parsed)


def _normalize_degrees(value: float) -> float:
    normalized = value % 360.0
    if normalized < 0:
        normalized += 360.0
    return normalized


def _is_upside_down_angle(value: float | None) -> bool:
    if value is None:
        return False
    angle = _normalize_degrees(value)
    return 135.0 <= angle <= 225.0


def extract_dxf_text_entities(file: PackageFile) -> list[dict[str, object]]:
    text, encoding, decode_loss = decode_text(file.data)
    lines = text.splitlines()
    entities: list[dict[str, object]] = []
    index = 0
    while index + 1 < len(lines):
        code = lines[index].strip()
        value = lines[index + 1].strip()
        if code != "0" or value.upper() not in DXF_TEXT_ENTITY_TOKENS:
            index += 2
            continue

        entity_type = value.upper()
        start_line = index + 1
        index += 2
        pairs: list[tuple[str, str]] = []
        while index + 1 < len(lines):
            group_code = lines[index].strip()
            group_value = lines[index + 1].rstrip("\r\n")
            if group_code == "0":
                break
            pairs.append((group_code, group_value))
            index += 2

        chunks: list[str] = []
        rotation: float | None = None
        x_scale: float | None = None
        generation_flags = 0
        mtext_direction_x: float | None = None
        mtext_direction_y: float | None = None
        for group_code, group_value in pairs:
            if group_code in {"1", "3"}:
                chunks.append(group_value)
            elif group_code == "50":
                rotation = _parse_float(group_value)
            elif group_code == "41":
                x_scale = _parse_float(group_value)
            elif entity_type != "MTEXT" and group_code == "71":
                generation_flags = _parse_int(group_value) or 0
            elif entity_type == "MTEXT" and group_code == "11":
                mtext_direction_x = _parse_float(group_value)
            elif entity_type == "MTEXT" and group_code == "21":
                mtext_direction_y = _parse_float(group_value)

        direction_angle: float | None = None
        if mtext_direction_x is not None and mtext_direction_y is not None:
            direction_angle = math.degrees(math.atan2(mtext_direction_y, mtext_direction_x))

        raw_text = "".join(chunks)
        entities.append(
            {
                "file": file.name,
                "entity_type": entity_type,
                "line": start_line,
                "encoding": encoding,
                "decode_loss": decode_loss,
                "text": raw_text,
                "text_sample": _sample_text(raw_text),
                "rotation_degrees": rotation,
                "direction_angle_degrees": direction_angle,
                "generation_flags": generation_flags,
                "x_scale": x_scale,
            }
        )
    return entities


def cad_text_quality_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    dxf_text_entities: list[dict[str, object]] = []
    text_quality_hits_by_entity: list[dict[str, object]] = []
    upside_down_text: list[dict[str, object]] = []
    mirrored_text: list[dict[str, object]] = []
    text_payload_hits: list[dict[str, object]] = []

    for file in files:
        if file.suffix in TEXT_SUFFIXES:
            text, _, _ = decode_text(file.data)
            hits = text_quality_hits(file.name, text)
            if hits:
                text_payload_hits.append({"file": file.name, "tokens": hits[:12]})
        if file.suffix != ".dxf":
            continue
        for entity in extract_dxf_text_entities(file):
            dxf_text_entities.append(entity)
            hits = text_quality_hits(str(entity.get("file", "")), str(entity.get("text", "")))
            if hits or entity.get("decode_loss") is True:
                text_quality_hits_by_entity.append(
                    {
                        "file": entity.get("file"),
                        "entity_type": entity.get("entity_type"),
                        "line": entity.get("line"),
                        "text_sample": entity.get("text_sample"),
                        "tokens": hits[:12],
                        "decode_loss": entity.get("decode_loss"),
                    }
                )
            flags = int(entity.get("generation_flags") or 0)
            rotation = entity.get("rotation_degrees")
            direction_angle = entity.get("direction_angle_degrees")
            is_upside_down = bool(flags & 4) or _is_upside_down_angle(
                rotation if isinstance(rotation, (int, float)) else None
            ) or _is_upside_down_angle(direction_angle if isinstance(direction_angle, (int, float)) else None)
            x_scale = entity.get("x_scale")
            is_mirrored = bool(flags & 2) or (isinstance(x_scale, (int, float)) and float(x_scale) < 0)
            if is_upside_down:
                upside_down_text.append(
                    {
                        "file": entity.get("file"),
                        "entity_type": entity.get("entity_type"),
                        "line": entity.get("line"),
                        "text_sample": entity.get("text_sample"),
                        "rotation_degrees": rotation,
                        "direction_angle_degrees": direction_angle,
                        "generation_flags": flags,
                    }
                )
            if is_mirrored:
                mirrored_text.append(
                    {
                        "file": entity.get("file"),
                        "entity_type": entity.get("entity_type"),
                        "line": entity.get("line"),
                        "text_sample": entity.get("text_sample"),
                        "x_scale": x_scale,
                        "generation_flags": flags,
                    }
                )

    source_hit_count = len(text_quality_hits_by_entity)
    payload_hit_count = len(text_payload_hits)
    missing_required_count = 0
    upside_down_count = len(upside_down_text)
    mirrored_count = len(mirrored_text)
    return {
        "audit_method": "dxf-text-entity-integrity-and-orientation-machine",
        "passed": source_hit_count == 0
        and payload_hit_count == 0
        and missing_required_count == 0
        and upside_down_count == 0
        and mirrored_count == 0,
        "source_text_entity_count": len(dxf_text_entities),
        "mojibake_or_missing_glyph_count": source_hit_count + payload_hit_count,
        "missing_required_drawing_text_count": missing_required_count,
        "upside_down_text_count": upside_down_count,
        "mirrored_text_count": mirrored_count,
        "entity_text_quality_hits": text_quality_hits_by_entity[:40],
        "text_payload_quality_hits": text_payload_hits[:40],
        "upside_down_text_samples": upside_down_text[:40],
        "mirrored_text_samples": mirrored_text[:40],
    }


def schematic_substitute_hits(name: str, text: str = "", *, broad: bool = True) -> list[str]:
    haystack = f"{name}\n{text}".lower()
    tokens = SCHEMATIC_SUBSTITUTE_BROAD_TOKENS if broad else SCHEMATIC_SUBSTITUTE_TEXT_TOKENS
    hits = [token for token in tokens if token.lower() in haystack]
    return sorted(set(hits))


def _find_dxf_header_value(lines: list[str], key: str) -> str:
    for i, line in enumerate(lines):
        if line == key and i + 2 < len(lines):
            return lines[i + 2]
    return ""


def dxf_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    aggregate_entities: Counter[str] = Counter()
    aggregate_sections: Counter[str] = Counter()
    per_file: list[dict[str, object]] = []
    mojibake_hits: list[dict[str, object]] = []
    text_quality_hits_by_file: list[dict[str, object]] = []

    for file in files:
        if file.suffix != ".dxf":
            continue
        text, encoding, decode_loss = decode_text(file.data)
        normalized_lines = [line.strip().upper() for line in text.splitlines()]
        line_counts = Counter(normalized_lines)
        entities = Counter(token for token in normalized_lines if token in ENTITY_TOKENS)
        sections = Counter(token for token in normalized_lines if token in STRICT_SECTION_TOKENS)
        aggregate_entities.update(entities)
        aggregate_sections.update(sections)
        hits = text_mojibake_hits(file.name, text)
        if hits:
            mojibake_hits.append({"file": file.name, "tokens": hits[:12]})
        quality_hits = text_quality_hits(file.name, text)
        if quality_hits:
            text_quality_hits_by_file.append({"file": file.name, "tokens": quality_hits[:12]})
        per_file.append(
            {
                "file": file.name,
                "encoding": encoding,
                "decode_loss": decode_loss,
                "size": file.size,
                "acadver": _find_dxf_header_value(normalized_lines, "$ACADVER"),
                "codepage": _find_dxf_header_value(normalized_lines, "$DWGCODEPAGE"),
                "section_count": int(line_counts["SECTION"]),
                "entities": dict(sorted(entities.items())),
                "structure_tokens": dict(sorted(sections.items())),
            }
        )

    return {
        "aggregate_entities": dict(sorted(aggregate_entities.items())),
        "aggregate_structure_tokens": dict(sorted(aggregate_sections.items())),
        "per_file": per_file,
        "mojibake_hits": mojibake_hits,
        "text_quality_hits": text_quality_hits_by_file,
    }


def json_manifest_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    manifests: list[dict[str, object]] = []
    failures: list[dict[str, object]] = []
    mojibake_hits: list[dict[str, object]] = []
    for file in files:
        if file.suffix != ".json":
            continue
        try:
            text = file.data.decode("utf-8-sig")
        except UnicodeDecodeError as exc:
            failures.append({"file": file.name, "error": f"not utf-8: {exc}"})
            continue
        hits = text_mojibake_hits(file.name, text)
        if hits:
            mojibake_hits.append({"file": file.name, "tokens": hits[:12]})
        try:
            payload = json.loads(text)
        except json.JSONDecodeError as exc:
            failures.append({"file": file.name, "error": str(exc)})
            continue
        manifests.append(
            {
                "file": file.name,
                "top_level_type": type(payload).__name__,
                "entry_count": len(payload) if isinstance(payload, (list, dict)) else None,
            }
        )
    return {"manifests": manifests, "failures": failures, "mojibake_hits": mojibake_hits}


def normalized_token(value: object) -> str:
    return str(value).strip().lower().replace(" ", "-")


def is_pass_value(value: object) -> bool:
    if isinstance(value, bool):
        return value
    return normalized_token(value) in RENDERED_REVIEW_PASS_VALUES


def is_entity_count_only_rejected(value: object) -> bool:
    if isinstance(value, bool):
        return value is False
    return normalized_token(value) in ENTITY_COUNT_ONLY_REJECT_VALUES


def collect_string_values(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        values: list[str] = []
        for item in value:
            values.extend(collect_string_values(item))
        return values
    return []


def collect_nested_string_values(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        values: list[str] = []
        for item in value:
            values.extend(collect_nested_string_values(item))
        return values
    if isinstance(value, dict):
        values = []
        for item in value.values():
            values.extend(collect_nested_string_values(item))
        return values
    return []


def formal_cad_source_provenance(files: Iterable[PackageFile]) -> dict[str, object]:
    package_files = list(files)
    cad_sources = [file.name for file in package_files if file.suffix in FORMAL_CAD_SOURCE_SUFFIXES]
    dwg_sources = [file.name for file in package_files if file.suffix == ".dwg"]
    dxf_sources = [file.name for file in package_files if file.suffix == ".dxf"]
    pdf_drawings = [file.name for file in package_files if file.suffix == ".pdf"]
    substitute_hits: list[dict[str, object]] = []
    for file in package_files:
        filename_hits = schematic_substitute_hits(file.name)
        if filename_hits:
            substitute_hits.append({"file": file.name, "surface": "filename", "tokens": filename_hits})
        if file.suffix == ".json":
            try:
                payload = json.loads(file.data.decode("utf-8-sig"))
            except Exception:
                continue
            values_text = "\n".join(collect_nested_string_values(payload))
            value_hits = schematic_substitute_hits("", values_text)
            if value_hits:
                substitute_hits.append({"file": file.name, "surface": "json-values", "tokens": value_hits})
        elif file.suffix in TEXT_SUFFIXES:
            text, _, _ = decode_text(file.data)
            text_hits = schematic_substitute_hits("", text, broad=False)
            if text_hits:
                substitute_hits.append({"file": file.name, "surface": "text-payload", "tokens": text_hits})
    return {
        "passed": bool(cad_sources) and not substitute_hits,
        "cad_source_suffixes": sorted(FORMAL_CAD_SOURCE_SUFFIXES),
        "editable_cad_source_count": len(cad_sources),
        "dwg_source_count": len(dwg_sources),
        "dxf_source_count": len(dxf_sources),
        "pdf_drawing_count": len(pdf_drawings),
        "cad_source_files": cad_sources[:200],
        "pdf_drawing_files": pdf_drawings[:200],
        "schematic_substitute_hits": substitute_hits,
        "schematic_substitute_rejection_verdict": "pass" if not substitute_hits else "fail",
    }


def _guess_sheet_mm(width_px: int, height_px: int) -> tuple[str, float, float]:
    long_px = max(width_px, height_px)
    short_px = min(width_px, height_px)
    ratio = long_px / max(short_px, 1)
    # ISO A-series drawings share the same aspect ratio, so this is enough for
    # rendered-frame placement even when the exact sheet size cannot be read.
    if abs(ratio - (2 ** 0.5)) > 0.06:
        return "unknown", 1189.0, 841.0
    if width_px >= height_px:
        return "A-series-landscape", 1189.0, 841.0
    return "A-series-portrait", 841.0, 1189.0


def _component_bboxes(mask: list[bytearray], width: int, height: int) -> list[dict[str, object]]:
    visited = [bytearray(width) for _ in range(height)]
    components: list[dict[str, object]] = []
    for y in range(height):
        row = mask[y]
        for x in range(width):
            if not row[x] or visited[y][x]:
                continue
            stack = [(x, y)]
            visited[y][x] = 1
            area = 0
            min_x = max_x = x
            min_y = max_y = y
            while stack:
                cx, cy = stack.pop()
                area += 1
                if cx < min_x:
                    min_x = cx
                if cx > max_x:
                    max_x = cx
                if cy < min_y:
                    min_y = cy
                if cy > max_y:
                    max_y = cy
                for nx, ny in ((cx - 1, cy), (cx + 1, cy), (cx, cy - 1), (cx, cy + 1)):
                    if 0 <= nx < width and 0 <= ny < height and mask[ny][nx] and not visited[ny][nx]:
                        visited[ny][nx] = 1
                        stack.append((nx, ny))
            if area >= FRAME_OVERFLOW_COMPONENT_MIN_AREA_PX:
                components.append(
                    {
                        "bbox_px": [min_x, min_y, max_x, max_y],
                        "area_px": area,
                        "width_px": max_x - min_x + 1,
                        "height_px": max_y - min_y + 1,
                    }
                )
    components.sort(key=lambda item: int(item["area_px"]), reverse=True)
    return components


def audit_png_frame_overflow(name: str, data: bytes) -> dict[str, object]:
    try:
        from PIL import Image  # type: ignore
        import io

        with Image.open(io.BytesIO(data)) as image:
            rgb = image.convert("RGB")
            width, height = rgb.size
            pixels = rgb.load()
    except Exception as exc:
        return {
            "file": name,
            "passed": False,
            "error": f"PNG open/parse failure: {exc}",
            "outside_frame_component_count": 1,
            "outside_frame_pixel_count": 1,
            "max_outside_component_area_px": 1,
            "outside_components": [],
        }

    sheet_guess, sheet_width_mm, sheet_height_mm = _guess_sheet_mm(width, height)
    left = round(FRAME_OVERFLOW_OUTER_MARGIN_MM / sheet_width_mm * width)
    right = round((sheet_width_mm - FRAME_OVERFLOW_OUTER_MARGIN_MM) / sheet_width_mm * width)
    top = round(FRAME_OVERFLOW_OUTER_MARGIN_MM / sheet_height_mm * height)
    bottom = round((sheet_height_mm - FRAME_OVERFLOW_OUTER_MARGIN_MM) / sheet_height_mm * height)
    tolerance = FRAME_OVERFLOW_FRAME_TOLERANCE_PX
    mask = [bytearray(width) for _ in range(height)]
    outside_pixel_count = 0
    for y in range(height):
        outside_y = y < top - tolerance or y > bottom + tolerance
        for x in range(width):
            if not outside_y and not (x < left - tolerance or x > right + tolerance):
                continue
            r, g, b = pixels[x, y]
            if min(r, g, b) < 245:
                mask[y][x] = 1
                outside_pixel_count += 1
    components = _component_bboxes(mask, width, height)
    return {
        "file": name,
        "passed": len(components) == 0,
        "sheet_guess": sheet_guess,
        "image_size_px": [width, height],
        "outer_frame_bbox_px": [left, top, right, bottom],
        "outer_margin_mm": FRAME_OVERFLOW_OUTER_MARGIN_MM,
        "frame_line_tolerance_px": tolerance,
        "outside_frame_component_count": len(components),
        "outside_frame_pixel_count": outside_pixel_count,
        "max_outside_component_area_px": int(components[0]["area_px"]) if components else 0,
        "outside_components": components[:12],
    }


def rendered_frame_overflow_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    png_files = [file for file in files if file.suffix == ".png"]
    per_sheet = [audit_png_frame_overflow(file.name, file.data) for file in png_files]
    outside_component_count = sum(int(row.get("outside_frame_component_count") or 0) for row in per_sheet)
    outside_pixel_count = sum(int(row.get("outside_frame_pixel_count") or 0) for row in per_sheet)
    max_area = max([int(row.get("max_outside_component_area_px") or 0) for row in per_sheet] + [0])
    failures = [
        {
            "file": row.get("file"),
            "outside_frame_component_count": row.get("outside_frame_component_count"),
            "max_outside_component_area_px": row.get("max_outside_component_area_px"),
            "outside_components": row.get("outside_components"),
        }
        for row in per_sheet
        if row.get("passed") is not True
    ]
    return {
        "audit_method": "png-rendered-outer-frame-connected-component-machine",
        "passed": bool(per_sheet) and not failures,
        "png_sheet_count": len(per_sheet),
        "outside_frame_component_count": outside_component_count,
        "outside_frame_pixel_count": outside_pixel_count,
        "max_outside_component_area_px": max_area,
        "component_min_area_px": FRAME_OVERFLOW_COMPONENT_MIN_AREA_PX,
        "frame_line_tolerance_px": FRAME_OVERFLOW_FRAME_TOLERANCE_PX,
        "outer_margin_mm": FRAME_OVERFLOW_OUTER_MARGIN_MM,
        "per_sheet": per_sheet,
        "failures": failures,
    }


def audit_png_ink_contrast(name: str, data: bytes) -> dict[str, object]:
    """Measure whether rendered CAD ink is readable against a white background."""

    try:
        from PIL import Image  # type: ignore
        import io

        with Image.open(io.BytesIO(data)) as image:
            rgb = image.convert("RGB")
            width, height = rgb.size
            pixels = rgb.load()
    except Exception as exc:
        return {
            "file": name,
            "passed": False,
            "error": f"PNG open/parse failure: {exc}",
            "ink_pixel_ratio": 0.0,
            "readable_ink_ratio": 0.0,
            "min_readable_ink_ratio": INK_CONTRAST_MIN_READABLE_RATIO,
            "min_ink_pixel_ratio": INK_CONTRAST_MIN_INK_PIXEL_RATIO,
        }

    total = max(width * height, 1)
    ink_pixels = 0
    readable_pixels = 0
    luma_sum = 0.0
    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            distance = ((255 - r) ** 2 + (255 - g) ** 2 + (255 - b) ** 2) ** 0.5
            if distance <= INK_CONTRAST_NONWHITE_DISTANCE:
                continue
            ink_pixels += 1
            luma_sum += 0.2126 * r + 0.7152 * g + 0.0722 * b
            if distance >= INK_CONTRAST_READABLE_DISTANCE:
                readable_pixels += 1
    ink_ratio = ink_pixels / total
    readable_ratio = readable_pixels / max(ink_pixels, 1)
    avg_luma = luma_sum / max(ink_pixels, 1)
    return {
        "file": name,
        "passed": ink_ratio >= INK_CONTRAST_MIN_INK_PIXEL_RATIO
        and readable_ratio >= INK_CONTRAST_MIN_READABLE_RATIO,
        "image_size_px": [width, height],
        "ink_pixel_count": ink_pixels,
        "readable_ink_pixel_count": readable_pixels,
        "ink_pixel_ratio": round(ink_ratio, 6),
        "readable_ink_ratio": round(readable_ratio, 4),
        "average_ink_luminance": round(avg_luma, 2),
        "nonwhite_distance_threshold": INK_CONTRAST_NONWHITE_DISTANCE,
        "readable_distance_threshold": INK_CONTRAST_READABLE_DISTANCE,
        "min_readable_ink_ratio": INK_CONTRAST_MIN_READABLE_RATIO,
        "min_ink_pixel_ratio": INK_CONTRAST_MIN_INK_PIXEL_RATIO,
    }


def rendered_ink_contrast_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    png_files = [file for file in files if file.suffix == ".png"]
    per_sheet = [audit_png_ink_contrast(file.name, file.data) for file in png_files]
    failures = [
        {
            "file": row.get("file"),
            "ink_pixel_ratio": row.get("ink_pixel_ratio"),
            "readable_ink_ratio": row.get("readable_ink_ratio"),
            "average_ink_luminance": row.get("average_ink_luminance"),
        }
        for row in per_sheet
        if row.get("passed") is not True
    ]
    readable_ratios = [float(row.get("readable_ink_ratio") or 0.0) for row in per_sheet]
    ink_ratios = [float(row.get("ink_pixel_ratio") or 0.0) for row in per_sheet]
    return {
        "audit_method": "png-rendered-ink-contrast-machine",
        "passed": bool(per_sheet) and not failures,
        "png_sheet_count": len(per_sheet),
        "min_readable_ink_ratio": INK_CONTRAST_MIN_READABLE_RATIO,
        "min_ink_pixel_ratio": INK_CONTRAST_MIN_INK_PIXEL_RATIO,
        "worst_readable_ink_ratio": round(min(readable_ratios), 4) if readable_ratios else None,
        "worst_ink_pixel_ratio": round(min(ink_ratios), 6) if ink_ratios else None,
        "per_sheet": per_sheet,
        "failures": failures,
    }


def outside_frame_ink_audit_from_rendered_frame_overflow(value: dict[str, object]) -> dict[str, object]:
    per_sheet_rows = value.get("per_sheet")
    per_sheet = per_sheet_rows if isinstance(per_sheet_rows, list) else []
    rows: list[dict[str, object]] = []
    for row in per_sheet:
        if not isinstance(row, dict):
            continue
        component_count = int(row.get("outside_frame_component_count") or 0)
        rows.append(
            {
                "sheet": row.get("file"),
                "passed": component_count == 0,
                "frame_bbox_px": row.get("outer_frame_bbox_px"),
                "outside_frame_independent_ink_component_count": component_count,
                "outside_frame_text_component_count": 0,
                "outside_frame_leader_component_count": 0,
                "outside_frame_hatch_section_component_count": 0,
                "outside_frame_table_title_block_component_count": 0,
                "max_outside_component_area_px": int(row.get("max_outside_component_area_px") or 0),
                "outside_components": row.get("outside_components", []),
            }
        )
    component_count = int(value.get("outside_frame_component_count") or 0)
    max_area = int(value.get("max_outside_component_area_px") or 0)
    return {
        "audit_method": "png-rendered-outer-frame-connected-component-machine",
        "frame_detection_method": "iso-a-series-outer-frame-margin-machine",
        "passed": bool(value.get("passed")) and component_count == 0 and max_area == 0,
        "rendered_source_formats": ["png"],
        "sheet_frame_bbox_count": int(value.get("png_sheet_count") or len(rows)),
        "outside_frame_independent_ink_component_count": component_count,
        "outside_frame_text_component_count": 0,
        "outside_frame_leader_component_count": 0,
        "outside_frame_hatch_section_component_count": 0,
        "outside_frame_table_title_block_component_count": 0,
        "max_outside_component_area_px": max_area,
        "per_sheet": rows,
    }


def collect_rendered_review_payloads(payload: object) -> list[tuple[str, dict[str, object]]]:
    reviews: list[tuple[str, dict[str, object]]] = []
    seen: set[int] = set()

    def walk(value: object, location: str) -> None:
        if isinstance(value, dict):
            review = value.get("rendered_review")
            if isinstance(review, dict) and id(review) not in seen:
                seen.add(id(review))
                reviews.append((f"{location}.rendered_review", review))
            is_review_like = (
                "entity_count_only_verdict" in value
                or "rendered_sheet_previews" in value
                or "preview_paths" in value
                or "manufacturing_view_depth_verdict" in value
                or "sheet_layout_verdict" in value
            )
            if is_review_like:
                if id(value) not in seen:
                    seen.add(id(value))
                    reviews.append((location, value))
                return
            for key, child in value.items():
                if key == "rendered_review" and isinstance(child, dict) and id(child) in seen:
                    continue
                walk(child, f"{location}.{key}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{location}[{index}]")

    walk(payload, "$")
    return reviews


def package_path_exists(reference: str, package_path: Path, member_names: set[str]) -> bool:
    if not reference:
        return False
    path = Path(reference)
    if path.is_absolute() and path.exists():
        try:
            if package_path.is_dir():
                path.resolve().relative_to(package_path.resolve())
                return True
            path.resolve().relative_to(package_path.parent.resolve())
            return True
        except ValueError:
            return False
    if package_path.is_dir() and (package_path / reference).exists():
        return True
    normalized = reference.replace("\\", "/").lstrip("./")
    return normalized in member_names


def _as_float(value: object, default: float | None = None) -> float | None:
    if isinstance(value, bool):
        return default
    try:
        return float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default


def _as_int(value: object, default: int | None = None) -> int | None:
    if isinstance(value, bool):
        return default
    try:
        return int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default


def validate_machine_overlap_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine-generated collision evidence for rendered CAD sheets."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["machine_overlap_audit is missing or not an object"]

    missing_fields = [field for field in MACHINE_OVERLAP_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("machine_overlap_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("machine_overlap_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("machine_overlap_audit.passed is not true/pass")

    for field in MACHINE_OVERLAP_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"machine_overlap_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"machine_overlap_audit.{field} must be 0")

    min_clearance = _as_float(value.get("min_clearance_mm"))
    max_density = _as_float(value.get("max_local_ink_density"))
    summary["min_clearance_mm"] = min_clearance
    summary["max_local_ink_density"] = max_density
    if min_clearance is None:
        failures.append("machine_overlap_audit.min_clearance_mm is not numeric")
    elif min_clearance < 2.0:
        failures.append("machine_overlap_audit.min_clearance_mm must be >= 2.0")
    if max_density is None:
        failures.append("machine_overlap_audit.max_local_ink_density is not numeric")
    elif max_density > 0.72:
        failures.append("machine_overlap_audit.max_local_ink_density must be <= 0.72")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"machine_overlap_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("machine_overlap_audit.per_sheet is missing or empty")
    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"machine_overlap_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"machine_overlap_audit.per_sheet[{index}].passed is not true/pass")
        for field in MACHINE_OVERLAP_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"machine_overlap_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"machine_overlap_audit.per_sheet[{index}].{field} must be 0")
        row_clearance = _as_float(row.get("min_clearance_mm", value.get("min_clearance_mm")))
        row_density = _as_float(row.get("max_local_ink_density", value.get("max_local_ink_density")))
        if row_clearance is None or row_clearance < 2.0:
            failures.append(f"machine_overlap_audit.per_sheet[{index}].min_clearance_mm must be >= 2.0")
        if row_density is None or row_density > 0.72:
            failures.append(f"machine_overlap_audit.per_sheet[{index}].max_local_ink_density must be <= 0.72")

    summary["passed"] = not failures
    return summary, failures


def validate_content_overlap_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate that rendered CAD content envelopes do not overlap."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["content_overlap_audit is missing or not an object"]

    missing_fields = [field for field in CONTENT_OVERLAP_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("content_overlap_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("content_overlap_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("content_overlap_audit.passed is not true/pass")

    bbox_count = _as_int(value.get("registered_bbox_count"))
    pair_count = _as_int(value.get("checked_pair_count"))
    min_clearance = _as_float(value.get("min_clearance_mm"))
    summary["registered_bbox_count"] = bbox_count
    summary["checked_pair_count"] = pair_count
    summary["min_clearance_mm"] = min_clearance
    if bbox_count is None or bbox_count <= 0:
        failures.append("content_overlap_audit.registered_bbox_count must be > 0")
    if pair_count is None:
        failures.append("content_overlap_audit.checked_pair_count is not numeric")
    elif bbox_count and bbox_count > 1 and pair_count <= 0:
        failures.append("content_overlap_audit.checked_pair_count must be > 0 when multiple boxes exist")
    if min_clearance is None:
        failures.append("content_overlap_audit.min_clearance_mm is not numeric")
    elif min_clearance < 2.0:
        failures.append("content_overlap_audit.min_clearance_mm must be >= 2.0")

    for field in CONTENT_OVERLAP_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"content_overlap_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"content_overlap_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            "content_overlap_audit.per_sheet has fewer rows than rendered review: "
            f"{len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("content_overlap_audit.per_sheet is missing or empty")
    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"content_overlap_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"content_overlap_audit.per_sheet[{index}].passed is not true/pass")
        row_bbox_count = _as_int(row.get("registered_bbox_count", bbox_count))
        row_pair_count = _as_int(row.get("checked_pair_count", pair_count))
        row_clearance = _as_float(row.get("min_clearance_mm", min_clearance))
        if row_bbox_count is None or row_bbox_count <= 0:
            failures.append(f"content_overlap_audit.per_sheet[{index}].registered_bbox_count must be > 0")
        if row_pair_count is None:
            failures.append(f"content_overlap_audit.per_sheet[{index}].checked_pair_count is not numeric")
        elif row_bbox_count and row_bbox_count > 1 and row_pair_count <= 0:
            failures.append(f"content_overlap_audit.per_sheet[{index}].checked_pair_count must be > 0")
        if row_clearance is None or row_clearance < 2.0:
            failures.append(f"content_overlap_audit.per_sheet[{index}].min_clearance_mm must be >= 2.0")
        for field in CONTENT_OVERLAP_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"content_overlap_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"content_overlap_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_title_block_short_line_topology_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that protected title/BOM/table short-line topology is complete."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["title_block_short_line_topology_audit is missing or not an object"]

    missing_fields = [field for field in TITLE_BLOCK_SHORT_LINE_TOPOLOGY_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("title_block_short_line_topology_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("title_block_short_line_topology_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("title_block_short_line_topology_audit.passed is not true/pass")

    baseline_source = str(value.get("baseline_source", "")).strip()
    summary["baseline_source"] = baseline_source
    if not baseline_source or normalized_token(baseline_source) in {"none", "not-applicable", "n/a"}:
        failures.append("title_block_short_line_topology_audit.baseline_source must name a source PDF/DXF/sample baseline")

    crop_path = str(value.get("diagnostic_overlay_free_title_block_crop_path", "")).strip()
    summary["diagnostic_overlay_free_title_block_crop_path"] = crop_path
    if not crop_path or normalized_token(crop_path) in {"none", "not-applicable", "n/a"}:
        failures.append("title_block_short_line_topology_audit.diagnostic_overlay_free_title_block_crop_path is missing")

    for field in TITLE_BLOCK_SHORT_LINE_TOPOLOGY_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"title_block_short_line_topology_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"title_block_short_line_topology_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            "title_block_short_line_topology_audit.per_sheet has fewer rows than rendered review: "
            f"{len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("title_block_short_line_topology_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"title_block_short_line_topology_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"title_block_short_line_topology_audit.per_sheet[{index}].passed is not true/pass")
        for field in TITLE_BLOCK_SHORT_LINE_TOPOLOGY_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"title_block_short_line_topology_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"title_block_short_line_topology_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_outside_frame_ink_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that no independent ink remains outside the drawing frame."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["outside_frame_ink_audit is missing or not an object"]

    missing_fields = [field for field in OUTSIDE_FRAME_INK_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("outside_frame_ink_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    frame_method = str(value.get("frame_detection_method", "")).strip().lower()
    summary["audit_method"] = method
    summary["frame_detection_method"] = frame_method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("outside_frame_ink_audit.audit_method must be a non-manual machine method")
    if not frame_method or any(token in frame_method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("outside_frame_ink_audit.frame_detection_method must be a non-manual frame detection method")
    if not is_pass_value(value.get("passed")):
        failures.append("outside_frame_ink_audit.passed is not true/pass")

    formats = collect_string_values(value.get("rendered_source_formats"))
    summary["rendered_source_formats"] = formats
    if not formats or not any(fmt.lower().lstrip(".") in {"png", "pdf"} for fmt in formats):
        failures.append("outside_frame_ink_audit.rendered_source_formats must include PNG or PDF")

    frame_count = _as_int(value.get("sheet_frame_bbox_count"))
    max_area = _as_float(value.get("max_outside_component_area_px"))
    summary["sheet_frame_bbox_count"] = frame_count
    summary["max_outside_component_area_px"] = max_area
    if frame_count is None:
        failures.append("outside_frame_ink_audit.sheet_frame_bbox_count is not numeric")
    elif expected_sheet_count and frame_count < expected_sheet_count:
        failures.append(
            f"outside_frame_ink_audit.sheet_frame_bbox_count has fewer frames than rendered review: {frame_count} < {expected_sheet_count}"
        )
    if max_area is None:
        failures.append("outside_frame_ink_audit.max_outside_component_area_px is not numeric")
    elif max_area != 0:
        failures.append("outside_frame_ink_audit.max_outside_component_area_px must be 0")

    for field in OUTSIDE_FRAME_INK_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"outside_frame_ink_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"outside_frame_ink_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            "outside_frame_ink_audit.per_sheet has fewer rows than rendered review: "
            f"{len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("outside_frame_ink_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"outside_frame_ink_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"outside_frame_ink_audit.per_sheet[{index}].passed is not true/pass")
        if not row.get("frame_bbox_px"):
            failures.append(f"outside_frame_ink_audit.per_sheet[{index}].frame_bbox_px is missing")
        row_max_area = _as_float(row.get("max_outside_component_area_px", 0))
        if row_max_area is None or row_max_area != 0:
            failures.append(f"outside_frame_ink_audit.per_sheet[{index}].max_outside_component_area_px must be 0")
        for field in OUTSIDE_FRAME_INK_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"outside_frame_ink_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"outside_frame_ink_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def inner_frame_safe_margin_audit_from_sections(
    *,
    machine_overlap: object,
    content_overlap: object,
    reserved_zone_intrusion: object,
    outside_frame_ink: object,
    expected_sheet_count: int,
) -> dict[str, object]:
    """Derive an inner-frame safe-margin audit from existing machine sections."""

    machine = machine_overlap if isinstance(machine_overlap, dict) else {}
    content = content_overlap if isinstance(content_overlap, dict) else {}
    reserved = reserved_zone_intrusion if isinstance(reserved_zone_intrusion, dict) else {}
    outside = outside_frame_ink if isinstance(outside_frame_ink, dict) else {}
    right_safe_boundary_intrusion_count = (
        (_as_int(machine.get("frame_clearance_violation_count")) or 0)
        + (_as_int(outside.get("outside_frame_independent_ink_component_count")) or 0)
        + (_as_int(outside.get("outside_frame_text_component_count")) or 0)
        + (_as_int(outside.get("outside_frame_leader_component_count")) or 0)
        + (_as_int(outside.get("outside_frame_hatch_section_component_count")) or 0)
        + (_as_int(outside.get("outside_frame_table_title_block_component_count")) or 0)
    )
    leader_text_inner_frame_intrusion_count = (
        (_as_int(reserved.get("leader_table_zone_intrusion_count")) or 0)
        + (_as_int(content.get("leader_line_view_table_crossing_count")) or 0)
    )
    view_geometry_inner_frame_intrusion_count = (
        (_as_int(reserved.get("view_geometry_reserved_zone_intrusion_count")) or 0)
        + (_as_int(content.get("view_view_overlap_count")) or 0)
        + (_as_int(content.get("detail_frame_main_view_overlap_count")) or 0)
    )
    dimension_text_inner_frame_intrusion_count = (
        (_as_int(reserved.get("dimension_text_table_zone_intrusion_count")) or 0)
        + (_as_int(reserved.get("dimension_line_table_zone_intrusion_count")) or 0)
        + (_as_int(content.get("dimension_line_view_table_crossing_count")) or 0)
    )
    sheet_frame_bbox_count = _as_int(outside.get("sheet_frame_bbox_count"))
    if sheet_frame_bbox_count is None:
        sheet_frame_bbox_count = expected_sheet_count
    min_safe_margin = _as_float(machine.get("min_clearance_mm"))
    if min_safe_margin is None:
        min_safe_margin = _as_float(content.get("min_clearance_mm"))
    if min_safe_margin is None:
        min_safe_margin = 0.0
    counts = {
        "right_safe_boundary_intrusion_count": right_safe_boundary_intrusion_count,
        "leader_text_inner_frame_intrusion_count": leader_text_inner_frame_intrusion_count,
        "view_geometry_inner_frame_intrusion_count": view_geometry_inner_frame_intrusion_count,
        "dimension_text_inner_frame_intrusion_count": dimension_text_inner_frame_intrusion_count,
    }
    passed = (
        all(count == 0 for count in counts.values())
        and sheet_frame_bbox_count >= expected_sheet_count
        and min_safe_margin >= 2.0
    )
    return {
        "present": True,
        "passed": passed,
        "audit_method": "derived_static_inner_frame_safe_margin_from_bbox_frame_reserved_zone",
        "source_sections": [
            "machine_overlap_audit",
            "content_overlap_audit",
            "reserved_zone_intrusion_audit",
            "outside_frame_ink_audit",
        ],
        "sheet_frame_bbox_count": sheet_frame_bbox_count,
        "min_safe_margin_mm": min_safe_margin,
        "per_sheet_count": expected_sheet_count,
        **counts,
    }


def validate_inner_frame_safe_margin_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate that content stays within the inner safe boundary of formal CAD sheets."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["inner_frame_safe_margin_audit is missing or not an object"]

    missing_fields = [field for field in INNER_FRAME_SAFE_MARGIN_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("inner_frame_safe_margin_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("inner_frame_safe_margin_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("inner_frame_safe_margin_audit.passed is not true/pass")

    sources = collect_string_values(value.get("source_sections"))
    summary["source_sections"] = sources
    required_sources = {
        "machine_overlap_audit",
        "content_overlap_audit",
        "reserved_zone_intrusion_audit",
        "outside_frame_ink_audit",
    }
    if not required_sources.issubset(set(sources)):
        failures.append("inner_frame_safe_margin_audit.source_sections must cite frame, content, reserved-zone, and outside-frame audits")

    frame_count = _as_int(value.get("sheet_frame_bbox_count"))
    min_margin = _as_float(value.get("min_safe_margin_mm"))
    per_sheet_count = _as_int(value.get("per_sheet_count", frame_count))
    summary["sheet_frame_bbox_count"] = frame_count
    summary["min_safe_margin_mm"] = min_margin
    summary["per_sheet_count"] = per_sheet_count or 0
    if frame_count is None:
        failures.append("inner_frame_safe_margin_audit.sheet_frame_bbox_count is not numeric")
    elif expected_sheet_count and frame_count < expected_sheet_count:
        failures.append(
            f"inner_frame_safe_margin_audit.sheet_frame_bbox_count has fewer frames than rendered review: {frame_count} < {expected_sheet_count}"
        )
    if min_margin is None:
        failures.append("inner_frame_safe_margin_audit.min_safe_margin_mm is not numeric")
    elif min_margin < 2.0:
        failures.append("inner_frame_safe_margin_audit.min_safe_margin_mm must be >= 2.0")

    for field in INNER_FRAME_SAFE_MARGIN_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"inner_frame_safe_margin_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"inner_frame_safe_margin_audit.{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_cell_containment_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine-generated table-cell text containment evidence."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["cell_containment_audit is missing or not an object"]

    missing_fields = [field for field in CELL_CONTAINMENT_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("cell_containment_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("cell_containment_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("cell_containment_audit.passed is not true/pass")

    table_text_count = _as_int(value.get("table_cell_text_count"))
    summary["table_cell_text_count"] = table_text_count
    if table_text_count is None:
        failures.append("cell_containment_audit.table_cell_text_count is not numeric")
    elif table_text_count <= 0:
        failures.append("cell_containment_audit.table_cell_text_count must be > 0")

    min_padding = _as_float(value.get("min_cell_padding_mm", 1.0))
    summary["min_cell_padding_mm"] = min_padding
    if min_padding is None or min_padding < 1.0:
        failures.append("cell_containment_audit.min_cell_padding_mm must be >= 1.0")

    for field in CELL_CONTAINMENT_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"cell_containment_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"cell_containment_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"cell_containment_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("cell_containment_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"cell_containment_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"cell_containment_audit.per_sheet[{index}].passed is not true/pass")
        row_text_count = _as_int(row.get("cell_text_count", 0))
        if row_text_count is None or row_text_count <= 0:
            failures.append(f"cell_containment_audit.per_sheet[{index}].cell_text_count must be > 0")
        row_min_padding = _as_float(row.get("min_cell_padding_mm", value.get("min_cell_padding_mm", 1.0)))
        if row_min_padding is None or row_min_padding < 1.0:
            failures.append(f"cell_containment_audit.per_sheet[{index}].min_cell_padding_mm must be >= 1.0")
        for field in CELL_CONTAINMENT_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"cell_containment_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"cell_containment_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_annotation_ownership_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that free text is owned by cells, symbols, or visible callout boxes."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["annotation_ownership_audit is missing or not an object"]

    missing_fields = [field for field in ANNOTATION_OWNERSHIP_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("annotation_ownership_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("annotation_ownership_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("annotation_ownership_audit.passed is not true/pass")

    text_count = _as_int(value.get("annotation_text_count"))
    summary["annotation_text_count"] = text_count
    if text_count is None:
        failures.append("annotation_ownership_audit.annotation_text_count is not numeric")
    elif text_count <= 0:
        failures.append("annotation_ownership_audit.annotation_text_count must be > 0")

    for field in ("boxed_callout_count", "symbol_or_title_text_count", *ANNOTATION_OWNERSHIP_COUNT_FIELDS):
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"annotation_ownership_audit.{field} is not numeric")
        elif field in ANNOTATION_OWNERSHIP_COUNT_FIELDS and count != 0:
            failures.append(f"annotation_ownership_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            "annotation_ownership_audit.per_sheet has fewer rows than rendered review: "
            f"{len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("annotation_ownership_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"annotation_ownership_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"annotation_ownership_audit.per_sheet[{index}].passed is not true/pass")
        row_text_count = _as_int(row.get("annotation_text_count", 0))
        if row_text_count is None or row_text_count <= 0:
            failures.append(f"annotation_ownership_audit.per_sheet[{index}].annotation_text_count must be > 0")
        for field in ANNOTATION_OWNERSHIP_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"annotation_ownership_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"annotation_ownership_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_reserved_zone_intrusion_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that non-table objects do not intrude into protected table zones."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["reserved_zone_intrusion_audit is missing or not an object"]

    missing_fields = [field for field in RESERVED_ZONE_INTRUSION_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("reserved_zone_intrusion_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("reserved_zone_intrusion_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("reserved_zone_intrusion_audit.passed is not true/pass")

    for field in RESERVED_ZONE_INTRUSION_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"reserved_zone_intrusion_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"reserved_zone_intrusion_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            "reserved_zone_intrusion_audit.per_sheet has fewer rows than rendered review: "
            f"{len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("reserved_zone_intrusion_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"reserved_zone_intrusion_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"reserved_zone_intrusion_audit.per_sheet[{index}].passed is not true/pass")
        for field in RESERVED_ZONE_INTRUSION_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"reserved_zone_intrusion_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"reserved_zone_intrusion_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_text_legibility_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine-readable text-height evidence for rendered CAD sheets."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
    }
    if not isinstance(value, dict):
        return summary, ["text_legibility_audit is missing or not an object"]

    missing_fields = [field for field in TEXT_LEGIBILITY_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("text_legibility_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("text_legibility_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("text_legibility_audit.passed is not true/pass")

    preview_dpi = _as_float(value.get("preview_dpi"))
    min_cad_height = _as_float(value.get("min_cad_text_height_mm"))
    required_cad_height = _as_float(value.get("min_required_cad_text_height_mm"))
    min_rendered_height = _as_float(value.get("min_rendered_text_height_px"))
    required_rendered_height = _as_float(value.get("min_required_rendered_text_height_px"))
    median_rendered_height = _as_float(value.get("median_rendered_text_height_px"))
    summary.update(
        {
            "preview_dpi": preview_dpi,
            "min_cad_text_height_mm": min_cad_height,
            "min_required_cad_text_height_mm": required_cad_height,
            "min_rendered_text_height_px": min_rendered_height,
            "min_required_rendered_text_height_px": required_rendered_height,
            "median_rendered_text_height_px": median_rendered_height,
        }
    )
    cad_threshold = max(TEXT_LEGIBILITY_MIN_CAD_HEIGHT_MM, required_cad_height or 0.0)
    rendered_threshold = max(TEXT_LEGIBILITY_MIN_RENDERED_HEIGHT_PX, required_rendered_height or 0.0)
    if preview_dpi is None or preview_dpi < 180:
        failures.append("text_legibility_audit.preview_dpi must be >= 180")
    if min_cad_height is None or min_cad_height < cad_threshold:
        failures.append(f"text_legibility_audit.min_cad_text_height_mm must be >= {cad_threshold:g}")
    if min_rendered_height is None or min_rendered_height < rendered_threshold:
        failures.append(f"text_legibility_audit.min_rendered_text_height_px must be >= {rendered_threshold:g}")
    if median_rendered_height is None or median_rendered_height < rendered_threshold:
        failures.append(f"text_legibility_audit.median_rendered_text_height_px must be >= {rendered_threshold:g}")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"text_legibility_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("text_legibility_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"text_legibility_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"text_legibility_audit.per_sheet[{index}].passed is not true/pass")
        row_dpi = _as_float(row.get("preview_dpi", preview_dpi))
        row_cad_height = _as_float(row.get("min_cad_text_height_mm", min_cad_height))
        row_rendered_height = _as_float(row.get("min_rendered_text_height_px", min_rendered_height))
        row_required_rendered = _as_float(row.get("min_required_rendered_text_height_px", required_rendered_height))
        row_threshold = max(TEXT_LEGIBILITY_MIN_RENDERED_HEIGHT_PX, row_required_rendered or 0.0)
        if row_dpi is None or row_dpi < 180:
            failures.append(f"text_legibility_audit.per_sheet[{index}].preview_dpi must be >= 180")
        if row_cad_height is None or row_cad_height < cad_threshold:
            failures.append(f"text_legibility_audit.per_sheet[{index}].min_cad_text_height_mm must be >= {cad_threshold:g}")
        if row_rendered_height is None or row_rendered_height < row_threshold:
            failures.append(
                f"text_legibility_audit.per_sheet[{index}].min_rendered_text_height_px must be >= {row_threshold:g}"
            )

    summary["passed"] = not failures
    return summary, failures


def validate_text_integrity_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that CAD drawing text is not mojibake or tofu."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
    }
    if not isinstance(value, dict):
        return summary, ["text_integrity_audit is missing or not an object"]

    missing_fields = [field for field in TEXT_INTEGRITY_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("text_integrity_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("text_integrity_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("text_integrity_audit.passed is not true/pass")

    source_count = _as_int(value.get("source_text_entity_count"))
    summary["source_text_entity_count"] = source_count
    if source_count is None:
        failures.append("text_integrity_audit.source_text_entity_count is not numeric")
    for field in TEXT_INTEGRITY_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"text_integrity_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"text_integrity_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"text_integrity_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("text_integrity_audit.per_sheet is missing or empty")
    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"text_integrity_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"text_integrity_audit.per_sheet[{index}].passed is not true/pass")
        for field in TEXT_INTEGRITY_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"text_integrity_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"text_integrity_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_text_orientation_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate machine evidence that CAD text is upright and not mirrored."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
    }
    if not isinstance(value, dict):
        return summary, ["text_orientation_audit is missing or not an object"]

    missing_fields = [field for field in TEXT_ORIENTATION_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("text_orientation_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("text_orientation_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("text_orientation_audit.passed is not true/pass")

    source_count = _as_int(value.get("source_text_entity_count"))
    summary["source_text_entity_count"] = source_count
    if source_count is None:
        failures.append("text_orientation_audit.source_text_entity_count is not numeric")
    for field in TEXT_ORIENTATION_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"text_orientation_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"text_orientation_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"text_orientation_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("text_orientation_audit.per_sheet is missing or empty")
    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"text_orientation_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"text_orientation_audit.per_sheet[{index}].passed is not true/pass")
        for field in TEXT_ORIENTATION_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"text_orientation_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"text_orientation_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def validate_manufacturing_complexity_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate that rendered CAD depth is not only sparse outlines or dense filler."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
    }
    if not isinstance(value, dict):
        return summary, ["manufacturing_complexity_audit is missing or not an object"]

    missing_fields = [field for field in MANUFACTURING_COMPLEXITY_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("manufacturing_complexity_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("manufacturing_complexity_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("manufacturing_complexity_audit.passed is not true/pass")
    min_family_count = _as_int(value.get("min_feature_family_count"))
    summary["min_feature_family_count"] = min_family_count
    if min_family_count is None or min_family_count < MANUFACTURING_COMPLEXITY_MIN_FEATURE_FAMILIES:
        failures.append(
            f"manufacturing_complexity_audit.min_feature_family_count must be >= {MANUFACTURING_COMPLEXITY_MIN_FEATURE_FAMILIES}"
        )
    dense_filler = value.get("uses_dense_filler_as_primary_depth")
    summary["uses_dense_filler_as_primary_depth"] = bool(dense_filler)
    if dense_filler is not False:
        failures.append("manufacturing_complexity_audit.uses_dense_filler_as_primary_depth must be false")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"manufacturing_complexity_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("manufacturing_complexity_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"manufacturing_complexity_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"manufacturing_complexity_audit.per_sheet[{index}].passed is not true/pass")
        families = row.get("manufacturing_detail_families")
        family_count = len(families) if isinstance(families, list) else 0
        if min_family_count is not None and family_count < min_family_count:
            failures.append(
                f"manufacturing_complexity_audit.per_sheet[{index}].manufacturing_detail_families has fewer than {min_family_count}"
            )
        if row.get("main_view_enhanced") is not True:
            failures.append(f"manufacturing_complexity_audit.per_sheet[{index}].main_view_enhanced must be true")
        if row.get("uses_dense_filler_as_primary_depth") is not False:
            failures.append(
                f"manufacturing_complexity_audit.per_sheet[{index}].uses_dense_filler_as_primary_depth must be false"
            )

    summary["passed"] = not failures
    return summary, failures


def validate_hatch_clip_audit(value: object, expected_sheet_count: int) -> tuple[dict[str, object], list[str]]:
    """Validate that hatch/section-fill evidence is clipped to its owning geometry."""

    failures: list[str] = []
    summary: dict[str, object] = {
        "present": isinstance(value, dict),
        "passed": False,
        "per_sheet_count": 0,
    }
    if not isinstance(value, dict):
        return summary, ["hatch_clip_audit is missing or not an object"]

    missing_fields = [field for field in HATCH_CLIP_REQUIRED_TOP_FIELDS if field not in value]
    if missing_fields:
        failures.append("hatch_clip_audit missing fields: " + ", ".join(missing_fields))

    method = str(value.get("audit_method", "")).strip().lower()
    summary["audit_method"] = method
    if not method or any(token in method for token in ("manual", "visual-only", "human-only", "self-report")):
        failures.append("hatch_clip_audit.audit_method must be a non-manual machine method")
    if not is_pass_value(value.get("passed")):
        failures.append("hatch_clip_audit.passed is not true/pass")

    region_count = _as_int(value.get("section_fill_region_count"))
    summary["section_fill_region_count"] = region_count
    if region_count is None:
        failures.append("hatch_clip_audit.section_fill_region_count is not numeric")
    elif region_count <= 0:
        failures.append("hatch_clip_audit.section_fill_region_count must be > 0")

    for field in HATCH_CLIP_COUNT_FIELDS:
        count = _as_int(value.get(field))
        summary[field] = count
        if count is None:
            failures.append(f"hatch_clip_audit.{field} is not numeric")
        elif count != 0:
            failures.append(f"hatch_clip_audit.{field} must be 0")

    rows = value.get("per_sheet")
    per_sheet_rows = rows if isinstance(rows, list) else []
    summary["per_sheet_count"] = len(per_sheet_rows)
    if expected_sheet_count and len(per_sheet_rows) < expected_sheet_count:
        failures.append(
            f"hatch_clip_audit.per_sheet has fewer rows than rendered review: {len(per_sheet_rows)} < {expected_sheet_count}"
        )
    if not per_sheet_rows:
        failures.append("hatch_clip_audit.per_sheet is missing or empty")

    for index, row in enumerate(per_sheet_rows):
        if not isinstance(row, dict):
            failures.append(f"hatch_clip_audit.per_sheet[{index}] is not an object")
            continue
        if not is_pass_value(row.get("passed")):
            failures.append(f"hatch_clip_audit.per_sheet[{index}].passed is not true/pass")
        row_regions = _as_int(row.get("section_fill_region_count", region_count))
        if row_regions is None or row_regions <= 0:
            failures.append(f"hatch_clip_audit.per_sheet[{index}].section_fill_region_count must be > 0")
        for field in HATCH_CLIP_COUNT_FIELDS:
            count = _as_int(row.get(field, 0))
            if count is None:
                failures.append(f"hatch_clip_audit.per_sheet[{index}].{field} is not numeric")
            elif count != 0:
                failures.append(f"hatch_clip_audit.per_sheet[{index}].{field} must be 0")

    summary["passed"] = not failures
    return summary, failures


def rendered_review_stats(files: Iterable[PackageFile], package_path: Path) -> dict[str, object]:
    package_files = list(files)
    member_names = {file.name.replace("\\", "/").lstrip("./") for file in package_files}
    live_frame_overflow = rendered_frame_overflow_stats(package_files)
    live_outside_frame_ink_audit = outside_frame_ink_audit_from_rendered_frame_overflow(live_frame_overflow)
    reviews: list[dict[str, object]] = []
    failures: list[dict[str, object]] = []

    for file in package_files:
        if file.suffix != ".json":
            continue
        try:
            text = file.data.decode("utf-8-sig")
            payload = json.loads(text)
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            failures.append({"file": file.name, "location": "$", "reason": f"manifest parse failure: {exc}"})
            continue
        for location, original_review in collect_rendered_review_payloads(payload):
            review = dict(original_review)
            if "outside_frame_ink_audit" not in review:
                review["outside_frame_ink_audit"] = live_outside_frame_ink_audit
            if "outside_frame_ink_verdict" not in review:
                review["outside_frame_ink_verdict"] = "pass" if live_outside_frame_ink_audit.get("passed") else "fail"
            if "inner_frame_safe_margin_audit" not in review:
                review["inner_frame_safe_margin_audit"] = inner_frame_safe_margin_audit_from_sections(
                    machine_overlap=review.get("machine_overlap_audit"),
                    content_overlap=review.get("content_overlap_audit"),
                    reserved_zone_intrusion=review.get("reserved_zone_intrusion_audit"),
                    outside_frame_ink=review.get("outside_frame_ink_audit"),
                    expected_sheet_count=len(review.get("per_sheet") if isinstance(review.get("per_sheet"), list) else []),
                )
            if "inner_frame_safe_margin_verdict" not in review:
                inner_frame_value = review.get("inner_frame_safe_margin_audit")
                review["inner_frame_safe_margin_verdict"] = (
                    "pass" if isinstance(inner_frame_value, dict) and inner_frame_value.get("passed") is True else "fail"
                )
            per_sheet_value = review.get("per_sheet")
            if isinstance(per_sheet_value, list):
                injected_rows: list[object] = []
                live_rows = live_outside_frame_ink_audit.get("per_sheet")
                live_row_list = live_rows if isinstance(live_rows, list) else []
                for row_index, row in enumerate(per_sheet_value):
                    if isinstance(row, dict) and "outside_frame_ink_verdict" not in row:
                        live_row = live_row_list[row_index] if row_index < len(live_row_list) else {}
                        row = dict(row)
                        row["outside_frame_ink_verdict"] = (
                            "pass" if isinstance(live_row, dict) and live_row.get("passed") is True else "fail"
                        )
                    if isinstance(row, dict) and "inner_frame_safe_margin_verdict" not in row:
                        row = dict(row)
                        inner_frame_value = review.get("inner_frame_safe_margin_audit")
                        row["inner_frame_safe_margin_verdict"] = (
                            "pass"
                            if isinstance(inner_frame_value, dict) and inner_frame_value.get("passed") is True
                            else "fail"
                        )
                    injected_rows.append(row)
                review["per_sheet"] = injected_rows
            review_failures: list[str] = []
            for field in RENDERED_REVIEW_PASS_FIELDS:
                if not is_pass_value(review.get(field)):
                    review_failures.append(f"{field} is not pass")
            if not is_entity_count_only_rejected(review.get("entity_count_only_verdict")):
                review_failures.append("entity_count_only_verdict does not reject entity-count-only acceptance")

            preview_paths = collect_string_values(review.get("preview_paths"))
            preview_paths.extend(collect_string_values(review.get("rendered_sheet_previews")))
            per_sheet = review.get("per_sheet")
            if per_sheet is None:
                per_sheet = review.get("sheets")
            per_sheet_rows = per_sheet if isinstance(per_sheet, list) else []
            if not preview_paths:
                review_failures.append("rendered preview paths are missing")
            if not per_sheet_rows:
                review_failures.append("per-sheet rendered review rows are missing")

            for row_index, row in enumerate(per_sheet_rows):
                if not isinstance(row, dict):
                    review_failures.append(f"per_sheet[{row_index}] is not an object")
                    continue
                for field in PER_SHEET_RENDERED_REVIEW_PASS_FIELDS:
                    if not is_pass_value(row.get(field)):
                        review_failures.append(f"per_sheet[{row_index}].{field} is not pass")
                preview_paths.extend(collect_string_values(row.get("preview_path")))
                preview_paths.extend(collect_string_values(row.get("preview_paths")))

            machine_overlap_summary, machine_overlap_failures = validate_machine_overlap_audit(
                review.get("machine_overlap_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(machine_overlap_failures)
            content_overlap_summary, content_overlap_failures = validate_content_overlap_audit(
                review.get("content_overlap_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(content_overlap_failures)
            short_line_topology_summary, short_line_topology_failures = validate_title_block_short_line_topology_audit(
                review.get("title_block_short_line_topology_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(short_line_topology_failures)
            outside_frame_ink_summary, outside_frame_ink_failures = validate_outside_frame_ink_audit(
                review.get("outside_frame_ink_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(outside_frame_ink_failures)
            cell_containment_summary, cell_containment_failures = validate_cell_containment_audit(
                review.get("cell_containment_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(cell_containment_failures)
            annotation_ownership_summary, annotation_ownership_failures = validate_annotation_ownership_audit(
                review.get("annotation_ownership_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(annotation_ownership_failures)
            reserved_zone_intrusion_summary, reserved_zone_intrusion_failures = validate_reserved_zone_intrusion_audit(
                review.get("reserved_zone_intrusion_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(reserved_zone_intrusion_failures)
            inner_frame_safe_margin_summary, inner_frame_safe_margin_failures = validate_inner_frame_safe_margin_audit(
                review.get("inner_frame_safe_margin_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(inner_frame_safe_margin_failures)
            text_legibility_summary, text_legibility_failures = validate_text_legibility_audit(
                review.get("text_legibility_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(text_legibility_failures)
            text_integrity_summary, text_integrity_failures = validate_text_integrity_audit(
                review.get("text_integrity_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(text_integrity_failures)
            text_orientation_summary, text_orientation_failures = validate_text_orientation_audit(
                review.get("text_orientation_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(text_orientation_failures)
            manufacturing_complexity_summary, manufacturing_complexity_failures = validate_manufacturing_complexity_audit(
                review.get("manufacturing_complexity_audit"),
                len(per_sheet_rows),
            )
            review_failures.extend(manufacturing_complexity_failures)
            hatch_clip_payload = review.get("hatch_clip_audit")
            if hatch_clip_payload is None:
                hatch_clip_payload = review.get("section_fill_clip_audit")
            hatch_clip_summary, hatch_clip_failures = validate_hatch_clip_audit(
                hatch_clip_payload,
                len(per_sheet_rows),
            )
            review_failures.extend(hatch_clip_failures)

            unique_preview_paths = sorted(set(preview_paths))
            preview_paths.extend(
                collect_string_values(short_line_topology_summary.get("diagnostic_overlay_free_title_block_crop_path"))
            )
            unique_preview_paths = sorted(set(preview_paths))
            missing_preview_paths = [
                path for path in unique_preview_paths if not package_path_exists(path, package_path, member_names)
            ]
            if missing_preview_paths:
                review_failures.append(
                    "rendered preview paths are not present in the package or filesystem: "
                    + ", ".join(missing_preview_paths[:8])
                )

            review_record = {
                "file": file.name,
                "location": location,
                "passed": not review_failures,
                "preview_paths": unique_preview_paths,
                "per_sheet_count": len(per_sheet_rows),
                "machine_overlap_audit": machine_overlap_summary,
                "content_overlap_audit": content_overlap_summary,
                "title_block_short_line_topology_audit": short_line_topology_summary,
                "outside_frame_ink_audit": outside_frame_ink_summary,
                "inner_frame_safe_margin_audit": inner_frame_safe_margin_summary,
                "cell_containment_audit": cell_containment_summary,
                "annotation_ownership_audit": annotation_ownership_summary,
                "reserved_zone_intrusion_audit": reserved_zone_intrusion_summary,
                "text_legibility_audit": text_legibility_summary,
                "text_integrity_audit": text_integrity_summary,
                "text_orientation_audit": text_orientation_summary,
                "manufacturing_complexity_audit": manufacturing_complexity_summary,
                "hatch_clip_audit": hatch_clip_summary,
                "failures": review_failures,
            }
            for field in (*RENDERED_REVIEW_PASS_FIELDS, "entity_count_only_verdict"):
                if field in review:
                    review_record[field] = review.get(field)
            reviews.append(review_record)
            if review_failures:
                failures.append(
                    {
                        "file": file.name,
                        "location": location,
                        "reason": "; ".join(review_failures[:8]),
                    }
                )

    accepted = [review for review in reviews if review.get("passed")]
    summary: dict[str, object] = {
        "required_schema": "rendered_review",
        "passed": bool(accepted),
        "review_count": len(reviews),
        "accepted_review_count": len(accepted),
        "reviews": reviews,
        "failures": failures,
    }
    if accepted:
        primary = accepted[0]
        for field in (
            *RENDERED_REVIEW_PASS_FIELDS,
            "entity_count_only_verdict",
            "preview_paths",
            "per_sheet_count",
            "machine_overlap_audit",
            "content_overlap_audit",
            "title_block_short_line_topology_audit",
            "outside_frame_ink_audit",
            "inner_frame_safe_margin_audit",
            "cell_containment_audit",
            "annotation_ownership_audit",
            "reserved_zone_intrusion_audit",
            "text_legibility_audit",
            "text_integrity_audit",
            "text_orientation_audit",
            "manufacturing_complexity_audit",
            "hatch_clip_audit",
        ):
            if field in primary:
                summary[field] = primary[field]
    return summary


def classify_sheet(width_mm: float, height_mm: float) -> tuple[str, float]:
    long_side = max(width_mm, height_mm)
    short_side = min(width_mm, height_mm)
    specs = (
        ("A0", 1189.0, 841.0, 1.0),
        ("A1", 841.0, 594.0, 0.5),
        ("A2", 594.0, 420.0, 0.25),
        ("A3", 420.0, 297.0, 0.125),
    )
    for name, spec_long, spec_short, equivalent in specs:
        if abs(long_side - spec_long) <= 12 and abs(short_side - spec_short) <= 12:
            return name, equivalent
    return "unknown", 0.0


def pdf_stats(files: Iterable[PackageFile]) -> dict[str, object]:
    pdf_files = [file for file in files if file.suffix == ".pdf"]
    header_failures = [file.name for file in pdf_files if not file.data.startswith(b"%PDF-")]
    per_file: list[dict[str, object]] = []
    workload_candidates: list[float] = []
    single_page_workload = 0.0
    fitz_available = False

    try:
        import fitz  # type: ignore

        fitz_available = True
    except Exception:
        fitz = None  # type: ignore

    if fitz_available:
        for file in pdf_files:
            item: dict[str, object] = {"file": file.name, "size": file.size}
            try:
                doc = fitz.open(stream=file.data, filetype="pdf")  # type: ignore[union-attr]
                page_workload = 0.0
                drawing_objects = 0
                pages: list[dict[str, object]] = []
                for index, page in enumerate(doc):
                    width_mm = page.rect.width * 25.4 / 72.0
                    height_mm = page.rect.height * 25.4 / 72.0
                    sheet_class, a0_equivalent = classify_sheet(width_mm, height_mm)
                    page_workload += a0_equivalent
                    try:
                        drawing_objects += len(page.get_drawings())
                    except Exception:
                        pass
                    pages.append(
                        {
                            "index": index + 1,
                            "width_mm": round(width_mm, 1),
                            "height_mm": round(height_mm, 1),
                            "sheet_class": sheet_class,
                            "a0_equivalent": a0_equivalent,
                        }
                    )
                item.update(
                    {
                        "page_count": doc.page_count,
                        "a0_equivalent": round(page_workload, 3),
                        "drawing_objects": drawing_objects,
                        "pages": pages,
                    }
                )
                if doc.page_count == 1:
                    single_page_workload += page_workload
                workload_candidates.append(page_workload)
            except Exception as exc:
                item["parse_error"] = str(exc)
            per_file.append(item)

    estimated_workload = max(workload_candidates + [single_page_workload, 0.0])
    return {
        "count": len(pdf_files),
        "total_bytes": sum(file.size for file in pdf_files),
        "header_failures": header_failures,
        "fitz_available": fitz_available,
        "estimated_a0_equivalent_workload": round(estimated_workload, 3),
        "per_file": per_file,
    }


def summarize_package(path: Path) -> dict[str, object]:
    files = load_package(path)
    dwg_files = [file for file in files if file.suffix == ".dwg"]
    dxf_files = [file for file in files if file.suffix == ".dxf"]
    filename_mojibake = [
        {"file": file.name, "tokens": text_mojibake_hits(file.name, "")}
        for file in files
        if text_mojibake_hits(file.name, "")
    ]
    text_mojibake: list[dict[str, object]] = []
    text_quality: list[dict[str, object]] = []
    for file in files:
        if file.suffix not in TEXT_SUFFIXES:
            continue
        text, _, _ = decode_text(file.data)
        hits = text_mojibake_hits(file.name, text)
        if hits:
            text_mojibake.append({"file": file.name, "tokens": hits[:12]})
        quality_hits = text_quality_hits(file.name, text)
        if quality_hits:
            text_quality.append({"file": file.name, "tokens": quality_hits[:12]})
    cad_text_quality = cad_text_quality_stats(files)

    return {
        "path": str(path),
        "path_type": "zip" if path.is_file() and path.suffix.lower() == ".zip" else "directory" if path.is_dir() else "file",
        "sha256": sha256_path(path),
        "file_count": len(files),
        "total_bytes": sum(file.size for file in files),
        "extension_counts": dict(sorted(Counter(file.suffix or "<none>" for file in files).items())),
        "dwg_count": len(dwg_files),
        "dwg_total_bytes": sum(file.size for file in dwg_files),
        "dwg_header_failures": [file.name for file in dwg_files if not file.data.startswith(b"AC")],
        "dxf_count": len(dxf_files),
        "dxf": dxf_stats(files),
        "pdf": pdf_stats(files),
        "json_manifest": json_manifest_stats(files),
        "formal_cad_source_provenance": formal_cad_source_provenance(files),
        "rendered_review": rendered_review_stats(files, path),
        "rendered_frame_overflow": rendered_frame_overflow_stats(files),
        "rendered_ink_contrast": rendered_ink_contrast_stats(files),
        "cad_text_quality": cad_text_quality,
        "filename_mojibake_hits": filename_mojibake,
        "text_mojibake_hits": text_mojibake,
        "text_quality_hits": text_quality,
    }


def text_corpus(files: Iterable[PackageFile]) -> str:
    chunks: list[str] = []
    for file in files:
        chunks.append(file.name)
        if file.suffix in TEXT_SUFFIXES:
            text, _, _ = decode_text(file.data)
            chunks.append(text)
    return "\n".join(chunks)


def audit_package(
    package_path: Path,
    *,
    reference_paths: list[Path],
    require_dwg: bool,
    require_pdf: bool,
    strict_cad_structure: bool,
    min_dwg_count: int,
    min_dxf_count: int,
    min_pdf_count: int,
    min_total_a0: float,
    min_true_dimensions: int,
    min_reference_dwg_ratio: float,
    min_geometry_entities: int,
    min_insert_count: int,
    min_arc_count: int,
    min_hatch_count: int,
    min_text_count: int,
    min_dxf_entities_per_sheet: int,
    min_dimensions_per_dxf: int,
    min_dimensioned_dxf_files: int,
    min_pdf_drawing_objects_per_a0: int,
    required_text_tokens: list[str],
    require_rendered_review: bool,
) -> dict[str, object]:
    candidate_files = load_package(package_path)
    candidate = summarize_package(package_path)
    references = [summarize_package(path) for path in reference_paths]
    issues: list[str] = []
    warnings: list[str] = []

    dwg_count = int(candidate["dwg_count"])
    pdf_count = int(candidate["pdf"]["count"])  # type: ignore[index]
    dxf_payload = candidate["dxf"]  # type: ignore[assignment]
    dxf_entities = dxf_payload["aggregate_entities"]  # type: ignore[index]
    dxf_structure = dxf_payload["aggregate_structure_tokens"]  # type: ignore[index]
    dxf_per_file = dxf_payload["per_file"]  # type: ignore[index]
    dxf_file_count = len(dxf_per_file) if isinstance(dxf_per_file, list) else int(candidate["dxf_count"])
    effective_min_dimensioned_dxf_files = min_dimensioned_dxf_files
    if min_dimensions_per_dxf and min_dimensioned_dxf_files and dxf_file_count:
        max_dimensioned_requirement = dxf_file_count
        if min_dxf_count:
            max_dimensioned_requirement = min(max_dimensioned_requirement, min_dxf_count)
        if effective_min_dimensioned_dxf_files > max_dimensioned_requirement:
            effective_min_dimensioned_dxf_files = max_dimensioned_requirement
            warnings.append(
                "dimensioned DXF distribution requirement clamped to available required sheet count: "
                f"{min_dimensioned_dxf_files} -> {effective_min_dimensioned_dxf_files}"
            )
    dimension_count = int(dxf_entities.get("DIMENSION", 0)) if isinstance(dxf_entities, dict) else 0
    workload = float(candidate["pdf"]["estimated_a0_equivalent_workload"])  # type: ignore[index]
    density_failures: list[dict[str, object]] = []
    geometry_entity_count = 0
    if isinstance(dxf_entities, dict):
        geometry_entity_count = sum(
            int(dxf_entities.get(token, 0))
            for token in ("LINE", "LWPOLYLINE", "POLYLINE", "ARC", "CIRCLE", "SPLINE", "ELLIPSE")
        )
        if min_geometry_entities and geometry_entity_count < min_geometry_entities:
            issues.append(
                f"geometry entity count below requirement: {geometry_entity_count} < {min_geometry_entities}"
            )
        for token, minimum, label in (
            ("INSERT", min_insert_count, "block INSERT"),
            ("ARC", min_arc_count, "ARC"),
            ("HATCH", min_hatch_count, "HATCH/section"),
            ("TEXT", min_text_count, "TEXT annotation"),
        ):
            actual = int(dxf_entities.get(token, 0))
            if minimum and actual < minimum:
                issues.append(f"{label} entity count below requirement: {actual} < {minimum}")

    if require_dwg and dwg_count < max(1, min_dwg_count):
        issues.append(f"DWG count below requirement: {dwg_count} < {max(1, min_dwg_count)}")
    if min_dxf_count and int(candidate["dxf_count"]) < min_dxf_count:
        issues.append(f"DXF count below requirement: {candidate['dxf_count']} < {min_dxf_count}")
    if require_pdf and pdf_count < max(1, min_pdf_count):
        issues.append(f"PDF count below requirement: {pdf_count} < {max(1, min_pdf_count)}")
    if candidate["dwg_header_failures"]:
        issues.append("one or more DWG files do not have an AutoCAD DWG header")
    if candidate["pdf"]["header_failures"]:  # type: ignore[index]
        issues.append("one or more PDF files do not have a PDF header")
    if candidate["filename_mojibake_hits"] or candidate["text_mojibake_hits"]:
        issues.append("package filenames or text payloads contain mojibake-like tokens")
    formal_cad_source = candidate["formal_cad_source_provenance"]  # type: ignore[assignment]
    if not formal_cad_source.get("editable_cad_source_count"):  # type: ignore[union-attr]
        issues.append("formal CAD source provenance missing: package must contain real DWG or DXF source files")
    if formal_cad_source.get("schematic_substitute_hits"):  # type: ignore[union-attr]
        issues.append(
            "schematic/concept/sketch placeholder files or metadata cannot substitute for required formal CAD drawings"
        )
    json_manifest = candidate["json_manifest"]  # type: ignore[assignment]
    if json_manifest["failures"]:  # type: ignore[index]
        issues.append("one or more JSON manifest files are invalid or not UTF-8")
    if json_manifest["mojibake_hits"]:  # type: ignore[index]
        issues.append("one or more JSON manifest files contain mojibake-like tokens")
    rendered_review = candidate["rendered_review"]  # type: ignore[assignment]
    effective_require_rendered_review = bool(require_rendered_review or strict_cad_structure or reference_paths)
    if effective_require_rendered_review and not rendered_review["passed"]:  # type: ignore[index]
        issues.append("rendered-sheet visual review is missing or failing")
    rendered_frame_overflow = candidate["rendered_frame_overflow"]  # type: ignore[assignment]
    if not isinstance(rendered_frame_overflow, dict) or rendered_frame_overflow.get("passed") is not True:
        issues.append("rendered-sheet frame overflow audit is missing or failing")
    rendered_ink_contrast = candidate["rendered_ink_contrast"]  # type: ignore[assignment]
    if not isinstance(rendered_ink_contrast, dict) or rendered_ink_contrast.get("passed") is not True:
        issues.append("rendered-sheet ink contrast/readability audit is missing or failing")
    cad_text_quality = candidate["cad_text_quality"]  # type: ignore[assignment]
    if not isinstance(cad_text_quality, dict) or cad_text_quality.get("passed") is not True:
        issues.append("CAD drawing text integrity/orientation audit is missing or failing")

    if strict_cad_structure and int(candidate["dxf_count"]) > 0:
        missing_structure = [token for token in STRICT_SECTION_TOKENS if not dxf_structure.get(token)]  # type: ignore[union-attr]
        if missing_structure:
            issues.append(f"DXF package lacks standard CAD table/style tokens: {', '.join(missing_structure)}")
    if min_true_dimensions and dimension_count < min_true_dimensions:
        issues.append(f"true DXF DIMENSION entity count below requirement: {dimension_count} < {min_true_dimensions}")
    if min_dxf_entities_per_sheet:
        for item in dxf_per_file if isinstance(dxf_per_file, list) else []:
            entities = item.get("entities", {}) if isinstance(item, dict) else {}
            entity_total = sum(int(v) for v in entities.values()) if isinstance(entities, dict) else 0
            if entity_total < min_dxf_entities_per_sheet:
                density_failures.append(
                    {
                        "file": item.get("file"),
                        "kind": "dxf_entity_total",
                        "actual": entity_total,
                        "required": min_dxf_entities_per_sheet,
                    }
                )
    if min_dimensions_per_dxf and effective_min_dimensioned_dxf_files:
        dimensioned_files = 0
        low_dimension_files: list[dict[str, object]] = []
        for item in dxf_per_file if isinstance(dxf_per_file, list) else []:
            entities = item.get("entities", {}) if isinstance(item, dict) else {}
            dims = int(entities.get("DIMENSION", 0)) if isinstance(entities, dict) else 0
            if dims >= min_dimensions_per_dxf:
                dimensioned_files += 1
            else:
                low_dimension_files.append(
                    {
                        "file": item.get("file"),
                        "actual": dims,
                        "required": min_dimensions_per_dxf,
                    }
                )
        if dimensioned_files < effective_min_dimensioned_dxf_files:
            density_failures.append(
                {
                    "kind": "dimension_distribution",
                    "actual_dimensioned_files": dimensioned_files,
                    "required_dimensioned_files": effective_min_dimensioned_dxf_files,
                    "configured_required_dimensioned_files": min_dimensioned_dxf_files,
                    "min_dimensions_per_dxf": min_dimensions_per_dxf,
                    "low_dimension_files": low_dimension_files[:12],
                }
            )
    if min_total_a0 and candidate["pdf"]["fitz_available"] and workload < min_total_a0:  # type: ignore[index]
        issues.append(f"estimated PDF sheet workload below requirement: {workload} A0 < {min_total_a0} A0")
    elif min_total_a0 and not candidate["pdf"]["fitz_available"]:  # type: ignore[index]
        warnings.append("PyMuPDF unavailable; sheet workload could not be measured")
    if min_pdf_drawing_objects_per_a0 and candidate["pdf"]["fitz_available"]:  # type: ignore[index]
        for item in candidate["pdf"]["per_file"]:  # type: ignore[index]
            if not isinstance(item, dict) or item.get("page_count") != 1:
                continue
            sheet_weight = float(item.get("a0_equivalent") or 0.0)
            if sheet_weight <= 0:
                continue
            required = int(min_pdf_drawing_objects_per_a0 * sheet_weight)
            actual = int(item.get("drawing_objects") or 0)
            if actual < required:
                density_failures.append(
                    {
                        "file": item.get("file"),
                        "kind": "pdf_drawing_objects",
                        "actual": actual,
                        "required": required,
                        "a0_equivalent": sheet_weight,
                    }
                )
    if required_text_tokens:
        corpus = text_corpus(candidate_files)
        missing = [token for token in required_text_tokens if token not in corpus]
        if missing:
            issues.append(f"required mechanical drawing text tokens missing: {', '.join(missing)}")
    if density_failures:
        issues.append("one or more sheets are below mechanical drawing density/distributed-dimension thresholds")

    reference_dwg_bytes = max([int(ref["dwg_total_bytes"]) for ref in references] + [0])
    candidate_dwg_bytes = int(candidate["dwg_total_bytes"])
    effective_min_reference_dwg_ratio = min_reference_dwg_ratio
    if reference_dwg_bytes:
        effective_min_reference_dwg_ratio = max(
            min_reference_dwg_ratio,
            MIN_REFERENCE_DWG_RATIO_WITH_REFERENCE,
        )
    if reference_dwg_bytes and effective_min_reference_dwg_ratio:
        ratio = candidate_dwg_bytes / reference_dwg_bytes if reference_dwg_bytes else 0.0
        candidate["reference_dwg_byte_ratio"] = round(ratio, 4)
        if ratio < effective_min_reference_dwg_ratio:
            issues.append(
                "DWG byte density ratio below reference threshold: "
                f"{ratio:.4f} < {effective_min_reference_dwg_ratio:.4f}"
            )
    density_verdict = {
        "passed": not density_failures
        and (not min_true_dimensions or dimension_count >= min_true_dimensions)
        and (not min_geometry_entities or geometry_entity_count >= min_geometry_entities),
        "dimension_count": dimension_count,
        "min_true_dimensions": min_true_dimensions,
        "geometry_entity_count": geometry_entity_count,
        "min_geometry_entities": min_geometry_entities,
        "reference_dwg_byte_ratio": candidate.get("reference_dwg_byte_ratio"),
        "min_reference_dwg_ratio": min_reference_dwg_ratio,
        "effective_min_reference_dwg_ratio": effective_min_reference_dwg_ratio,
    }
    manufacturing_depth = {
        "insert_count": int(dxf_entities.get("INSERT", 0)) if isinstance(dxf_entities, dict) else 0,
        "min_insert_count": min_insert_count,
        "arc_count": int(dxf_entities.get("ARC", 0)) if isinstance(dxf_entities, dict) else 0,
        "min_arc_count": min_arc_count,
        "hatch_count": int(dxf_entities.get("HATCH", 0)) if isinstance(dxf_entities, dict) else 0,
        "min_hatch_count": min_hatch_count,
        "text_count": int(dxf_entities.get("TEXT", 0)) if isinstance(dxf_entities, dict) else 0,
        "min_text_count": min_text_count,
    }
    manufacturing_depth["passed"] = (
        (not min_insert_count or int(manufacturing_depth["insert_count"]) >= min_insert_count)
        and (not min_arc_count or int(manufacturing_depth["arc_count"]) >= min_arc_count)
        and (not min_hatch_count or int(manufacturing_depth["hatch_count"]) >= min_hatch_count)
        and (not min_text_count or int(manufacturing_depth["text_count"]) >= min_text_count)
    )
    formal_cad_source_verdict = {
        "passed": bool(formal_cad_source.get("passed")) if isinstance(formal_cad_source, dict) else False,
        "editable_cad_source_count": formal_cad_source.get("editable_cad_source_count") if isinstance(formal_cad_source, dict) else 0,
        "dwg_source_count": formal_cad_source.get("dwg_source_count") if isinstance(formal_cad_source, dict) else 0,
        "dxf_source_count": formal_cad_source.get("dxf_source_count") if isinstance(formal_cad_source, dict) else 0,
        "pdf_drawing_count": formal_cad_source.get("pdf_drawing_count") if isinstance(formal_cad_source, dict) else 0,
        "schematic_substitute_rejection_verdict": (
            formal_cad_source.get("schematic_substitute_rejection_verdict")
            if isinstance(formal_cad_source, dict)
            else "fail"
        ),
        "schematic_substitute_hits": (
            formal_cad_source.get("schematic_substitute_hits") if isinstance(formal_cad_source, dict) else []
        ),
    }

    return {
        "schema": "graduation-project-builder.mechanical-drawing-package-audit.v6",
        "package_path": str(package_path),
        "package_sha256": candidate["sha256"],
        "reference_paths": [str(path) for path in reference_paths],
        "passed": not issues,
        "issues": issues,
        "warnings": warnings,
        "candidate": candidate,
        "references": references,
        "requirements": {
            "require_dwg": require_dwg,
            "require_pdf": require_pdf,
            "strict_cad_structure": strict_cad_structure,
            "min_dwg_count": min_dwg_count,
            "min_dxf_count": min_dxf_count,
            "min_pdf_count": min_pdf_count,
            "min_total_a0": min_total_a0,
            "min_true_dimensions": min_true_dimensions,
            "min_reference_dwg_ratio": min_reference_dwg_ratio,
            "effective_min_reference_dwg_ratio": effective_min_reference_dwg_ratio,
            "min_geometry_entities": min_geometry_entities,
            "min_insert_count": min_insert_count,
            "min_arc_count": min_arc_count,
            "min_hatch_count": min_hatch_count,
            "min_text_count": min_text_count,
            "min_dxf_entities_per_sheet": min_dxf_entities_per_sheet,
            "min_dimensions_per_dxf": min_dimensions_per_dxf,
            "min_dimensioned_dxf_files": min_dimensioned_dxf_files,
            "effective_min_dimensioned_dxf_files": effective_min_dimensioned_dxf_files,
            "min_pdf_drawing_objects_per_a0": min_pdf_drawing_objects_per_a0,
            "required_text_tokens": required_text_tokens,
            "require_rendered_review": require_rendered_review,
            "effective_require_rendered_review": effective_require_rendered_review,
        },
        "drawing_density_checks": {
            "failures": density_failures,
        },
        "formal_cad_source_verdict": formal_cad_source_verdict,
        "density_verdict": density_verdict,
        "manufacturing_depth": manufacturing_depth,
        "rendered_frame_overflow_verdict": {
            "passed": bool(rendered_frame_overflow.get("passed")) if isinstance(rendered_frame_overflow, dict) else False,
            "png_sheet_count": rendered_frame_overflow.get("png_sheet_count") if isinstance(rendered_frame_overflow, dict) else 0,
            "outside_frame_component_count": (
                rendered_frame_overflow.get("outside_frame_component_count")
                if isinstance(rendered_frame_overflow, dict)
                else None
            ),
            "outside_frame_pixel_count": (
                rendered_frame_overflow.get("outside_frame_pixel_count")
                if isinstance(rendered_frame_overflow, dict)
                else None
            ),
            "max_outside_component_area_px": (
                rendered_frame_overflow.get("max_outside_component_area_px")
                if isinstance(rendered_frame_overflow, dict)
                else None
            ),
            "audit_method": rendered_frame_overflow.get("audit_method") if isinstance(rendered_frame_overflow, dict) else None,
        },
        "rendered_ink_contrast_verdict": {
            "passed": bool(rendered_ink_contrast.get("passed")) if isinstance(rendered_ink_contrast, dict) else False,
            "png_sheet_count": rendered_ink_contrast.get("png_sheet_count") if isinstance(rendered_ink_contrast, dict) else 0,
            "worst_readable_ink_ratio": (
                rendered_ink_contrast.get("worst_readable_ink_ratio")
                if isinstance(rendered_ink_contrast, dict)
                else None
            ),
            "worst_ink_pixel_ratio": (
                rendered_ink_contrast.get("worst_ink_pixel_ratio")
                if isinstance(rendered_ink_contrast, dict)
                else None
            ),
            "min_readable_ink_ratio": (
                rendered_ink_contrast.get("min_readable_ink_ratio")
                if isinstance(rendered_ink_contrast, dict)
                else None
            ),
            "audit_method": rendered_ink_contrast.get("audit_method") if isinstance(rendered_ink_contrast, dict) else None,
        },
        "cad_text_quality_verdict": {
            "passed": bool(cad_text_quality.get("passed")) if isinstance(cad_text_quality, dict) else False,
            "audit_method": cad_text_quality.get("audit_method") if isinstance(cad_text_quality, dict) else None,
            "source_text_entity_count": (
                cad_text_quality.get("source_text_entity_count") if isinstance(cad_text_quality, dict) else None
            ),
            "mojibake_or_missing_glyph_count": (
                cad_text_quality.get("mojibake_or_missing_glyph_count") if isinstance(cad_text_quality, dict) else None
            ),
            "missing_required_drawing_text_count": (
                cad_text_quality.get("missing_required_drawing_text_count") if isinstance(cad_text_quality, dict) else None
            ),
            "upside_down_text_count": (
                cad_text_quality.get("upside_down_text_count") if isinstance(cad_text_quality, dict) else None
            ),
            "mirrored_text_count": (
                cad_text_quality.get("mirrored_text_count") if isinstance(cad_text_quality, dict) else None
            ),
            "entity_text_quality_hits": (
                cad_text_quality.get("entity_text_quality_hits") if isinstance(cad_text_quality, dict) else []
            ),
            "text_payload_quality_hits": (
                cad_text_quality.get("text_payload_quality_hits") if isinstance(cad_text_quality, dict) else []
            ),
            "upside_down_text_samples": (
                cad_text_quality.get("upside_down_text_samples") if isinstance(cad_text_quality, dict) else []
            ),
            "mirrored_text_samples": (
                cad_text_quality.get("mirrored_text_samples") if isinstance(cad_text_quality, dict) else []
            ),
        },
        "rendered_review_verdict": {
            "passed": bool(rendered_review["passed"]),  # type: ignore[index]
            "review_count": int(rendered_review["review_count"]),  # type: ignore[index]
            "accepted_review_count": int(rendered_review["accepted_review_count"]),  # type: ignore[index]
            "machine_overlap_audit": rendered_review.get("machine_overlap_audit") if isinstance(rendered_review, dict) else None,
            "content_overlap_audit": rendered_review.get("content_overlap_audit") if isinstance(rendered_review, dict) else None,
            "title_block_short_line_topology_audit": rendered_review.get("title_block_short_line_topology_audit") if isinstance(rendered_review, dict) else None,
            "outside_frame_ink_audit": rendered_review.get("outside_frame_ink_audit") if isinstance(rendered_review, dict) else None,
            "cell_containment_audit": rendered_review.get("cell_containment_audit") if isinstance(rendered_review, dict) else None,
            "annotation_ownership_audit": rendered_review.get("annotation_ownership_audit") if isinstance(rendered_review, dict) else None,
            "reserved_zone_intrusion_audit": rendered_review.get("reserved_zone_intrusion_audit") if isinstance(rendered_review, dict) else None,
            "inner_frame_safe_margin_audit": rendered_review.get("inner_frame_safe_margin_audit") if isinstance(rendered_review, dict) else None,
            "text_legibility_audit": rendered_review.get("text_legibility_audit") if isinstance(rendered_review, dict) else None,
            "text_integrity_audit": rendered_review.get("text_integrity_audit") if isinstance(rendered_review, dict) else None,
            "text_orientation_audit": rendered_review.get("text_orientation_audit") if isinstance(rendered_review, dict) else None,
            "manufacturing_complexity_audit": rendered_review.get("manufacturing_complexity_audit") if isinstance(rendered_review, dict) else None,
            "hatch_clip_audit": rendered_review.get("hatch_clip_audit") if isinstance(rendered_review, dict) else None,
        },
    }


def run_self_test() -> int:
    with tempfile.TemporaryDirectory() as tmpdir:
        root = Path(tmpdir)
        good = root / "good"
        bad = root / "bad"
        good.mkdir()
        bad.mkdir()
        entity_only = root / "entity_only"
        overlap = root / "overlap"
        machine_missing = root / "machine_missing"
        schematic_substitute = root / "schematic_substitute"
        frame_overflow = root / "frame_overflow"
        tofu_text = root / "tofu_text"
        inverted_text = root / "inverted_text"
        mtext_attachment = root / "mtext_attachment"
        entity_only.mkdir()
        overlap.mkdir()
        machine_missing.mkdir()
        schematic_substitute.mkdir()
        frame_overflow.mkdir()
        tofu_text.mkdir()
        inverted_text.mkdir()
        mtext_attachment.mkdir()
        try:
            from PIL import Image, ImageDraw  # type: ignore
        except Exception as exc:
            print(f"PIL unavailable for frame-overflow self-test: {exc}")
            return 1

        def write_test_sheet_png(path: Path, *, overflow: bool = False) -> None:
            width, height = 1190, 842
            image = Image.new("RGB", (width, height), "white")
            draw = ImageDraw.Draw(image)
            outer = (10, 10, width - 10, height - 10)
            inner = (20, 20, width - 20, height - 20)
            draw.rectangle(outer, outline="black", width=1)
            draw.rectangle(inner, outline="black", width=1)
            draw.rectangle((80, 80, 360, 180), outline="black", width=1)
            if overflow:
                draw.rectangle((width - 4, 140, width - 1, 220), fill="black")
            image.save(path)

        (good / "sheet.dwg").write_bytes(b"AC1032\x00" + b"x" * 4096)
        (good / "sheet.pdf").write_bytes(b"%PDF-1.4\n%%EOF\n")
        write_test_sheet_png(good / "sheet-preview.png")
        write_test_sheet_png(good / "sheet-title-crop.png")
        (good / "sheet.dxf").write_text(
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n0\nSTYLE\n0\nDIMSTYLE\n0\nLTYPE\n0\nENDSEC\n"
            "0\nSECTION\n2\nBLOCKS\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n"
            "0\nTEXT\n1\nTITLE\n50\n0\n71\n0\n0\nDIMENSION\n0\nENDSEC\n0\nEOF\n",
            encoding="utf-8",
        )
        (good / "manifest.json").write_text(
            json.dumps(
                {
                    "files": ["sheet.dwg", "sheet.pdf", "sheet.dxf"],
                    "rendered_review": {
                        "preview_paths": ["sheet-preview.png"],
                        "no_overlap_verdict": "pass",
                        "boundary_clearance_verdict": "pass",
                        "detail_density_verdict": "pass",
                        "text_legibility_verdict": "pass",
                        "text_integrity_verdict": "pass",
                        "text_orientation_verdict": "pass",
                        "title_block_table_notes_isolation_verdict": "pass",
                        "title_block_cell_containment_verdict": "pass",
                        "title_block_short_line_topology_verdict": "pass",
                        "annotation_margin_clearance_verdict": "pass",
                        "local_crowding_verdict": "pass",
                        "sheet_layout_verdict": "pass",
                        "manufacturing_view_depth_verdict": "pass",
                        "layout_collision_verdict": "pass",
                        "content_overlap_verdict": "pass",
                        "outside_frame_ink_verdict": "pass",
                        "hatch_section_fill_clipping_verdict": "pass",
                        "entity_count_only_verdict": "not-used",
                        "machine_overlap_audit": {
                            "audit_method": "bbox-and-rendered-zone-machine",
                            "passed": True,
                            "overlap_count": 0,
                            "text_entity_overlap_count": 0,
                            "reserved_zone_collision_count": 0,
                            "title_block_table_note_collision_count": 0,
                            "annotation_collision_count": 0,
                            "frame_clearance_violation_count": 0,
                            "min_clearance_mm": 4.0,
                            "max_local_ink_density": 0.38,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "overlap_count": 0,
                                    "text_entity_overlap_count": 0,
                                    "reserved_zone_collision_count": 0,
                                    "title_block_table_note_collision_count": 0,
                                    "annotation_collision_count": 0,
                                    "frame_clearance_violation_count": 0,
                                    "min_clearance_mm": 4.0,
                                    "max_local_ink_density": 0.38,
                                }
                            ],
                        },
                        "content_overlap_audit": {
                            "audit_method": "content-bbox-pairwise-machine",
                            "passed": True,
                            "registered_bbox_count": 8,
                            "checked_pair_count": 28,
                            "min_clearance_mm": 4.0,
                            "content_overlap_count": 0,
                            "view_view_overlap_count": 0,
                            "detail_frame_main_view_overlap_count": 0,
                            "table_text_grid_collision_count": 0,
                            "dimension_line_view_table_crossing_count": 0,
                            "leader_line_view_table_crossing_count": 0,
                            "balloon_geometry_collision_count": 0,
                            "bbox_helper_envelope_escape_count": 0,
                            "stale_rendered_preview_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "registered_bbox_count": 8,
                                    "checked_pair_count": 28,
                                    "min_clearance_mm": 4.0,
                                    "content_overlap_count": 0,
                                    "view_view_overlap_count": 0,
                                    "detail_frame_main_view_overlap_count": 0,
                                    "table_text_grid_collision_count": 0,
                                    "dimension_line_view_table_crossing_count": 0,
                                    "leader_line_view_table_crossing_count": 0,
                                    "balloon_geometry_collision_count": 0,
                                    "bbox_helper_envelope_escape_count": 0,
                                    "stale_rendered_preview_count": 0,
                                }
                            ],
                        },
                        "title_block_short_line_topology_audit": {
                            "audit_method": "source-pdf-dxf-short-line-topology-machine",
                            "passed": True,
                            "baseline_source": "sheet.pdf; sheet.dxf",
                            "missing_short_table_line_count": 0,
                            "broken_cell_border_count": 0,
                            "table_grid_topology_mismatch_count": 0,
                            "diagnostic_overlay_free_title_block_crop_path": "sheet-title-crop.png",
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "missing_short_table_line_count": 0,
                                    "broken_cell_border_count": 0,
                                    "table_grid_topology_mismatch_count": 0,
                                }
                            ],
                        },
                        "outside_frame_ink_audit": {
                            "audit_method": "rendered-component-frame-mask-machine",
                            "passed": True,
                            "frame_detection_method": "outer-rectangle-line-detection-machine",
                            "rendered_source_formats": ["png", "pdf"],
                            "sheet_frame_bbox_count": 1,
                            "outside_frame_independent_ink_component_count": 0,
                            "outside_frame_text_component_count": 0,
                            "outside_frame_leader_component_count": 0,
                            "outside_frame_hatch_section_component_count": 0,
                            "outside_frame_table_title_block_component_count": 0,
                            "max_outside_component_area_px": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "frame_bbox_px": [12, 12, 1180, 830],
                                    "outside_frame_independent_ink_component_count": 0,
                                    "outside_frame_text_component_count": 0,
                                    "outside_frame_leader_component_count": 0,
                                    "outside_frame_hatch_section_component_count": 0,
                                    "outside_frame_table_title_block_component_count": 0,
                                    "max_outside_component_area_px": 0,
                                }
                            ],
                        },
                        "cell_containment_audit": {
                            "audit_method": "cell-bbox-text-placement-machine",
                            "passed": True,
                            "table_cell_text_count": 6,
                            "min_cell_padding_mm": 1.2,
                            "outside_cell_count": 0,
                            "border_touch_count": 0,
                            "unowned_table_text_count": 0,
                            "clipped_overflow_count": 0,
                            "cell_padding_violation_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "cell_text_count": 6,
                                    "min_cell_padding_mm": 1.2,
                                    "outside_cell_count": 0,
                                    "border_touch_count": 0,
                                    "unowned_table_text_count": 0,
                                    "clipped_overflow_count": 0,
                                    "cell_padding_violation_count": 0,
                                }
                            ],
                        },
                        "annotation_ownership_audit": {
                            "audit_method": "cad-text-owner-and-callout-box-machine",
                            "passed": True,
                            "annotation_text_count": 3,
                            "boxed_callout_count": 2,
                            "symbol_or_title_text_count": 1,
                            "unowned_free_text_count": 0,
                            "unsupported_floating_text_count": 0,
                            "unbound_scattered_text_count": 0,
                            "dimension_like_text_without_anchor_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "annotation_text_count": 3,
                                    "boxed_callout_count": 2,
                                    "symbol_or_title_text_count": 1,
                                    "unowned_free_text_count": 0,
                                    "unsupported_floating_text_count": 0,
                                    "unbound_scattered_text_count": 0,
                                    "dimension_like_text_without_anchor_count": 0,
                                }
                            ],
                        },
                        "reserved_zone_intrusion_audit": {
                            "audit_method": "cad-object-bbox-protected-table-zone-machine",
                            "passed": True,
                            "intrusion_count": 0,
                            "geometry_title_block_intrusion_count": 0,
                            "dimension_line_table_zone_intrusion_count": 0,
                            "dimension_text_table_zone_intrusion_count": 0,
                            "leader_table_zone_intrusion_count": 0,
                            "view_geometry_reserved_zone_intrusion_count": 0,
                            "bom_table_intrusion_count": 0,
                            "technical_requirement_intrusion_count": 0,
                            "protected_table_zone_intrusion_count": 0,
                            "view_geometry_table_zone_intrusion_count": 0,
                            "detail_view_table_zone_intrusion_count": 0,
                            "leader_balloon_table_zone_intrusion_count": 0,
                            "dimension_table_zone_intrusion_count": 0,
                            "title_block_bom_protected_zone_intrusion_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "intrusion_count": 0,
                                    "geometry_title_block_intrusion_count": 0,
                                    "dimension_line_table_zone_intrusion_count": 0,
                                    "dimension_text_table_zone_intrusion_count": 0,
                                    "leader_table_zone_intrusion_count": 0,
                                    "view_geometry_reserved_zone_intrusion_count": 0,
                                    "bom_table_intrusion_count": 0,
                                    "technical_requirement_intrusion_count": 0,
                                    "protected_table_zone_intrusion_count": 0,
                                    "view_geometry_table_zone_intrusion_count": 0,
                                    "detail_view_table_zone_intrusion_count": 0,
                                    "leader_balloon_table_zone_intrusion_count": 0,
                                    "dimension_table_zone_intrusion_count": 0,
                                    "title_block_bom_protected_zone_intrusion_count": 0,
                                }
                            ],
                        },
                        "text_legibility_audit": {
                            "audit_method": "cad-text-height-and-rendered-pixel-machine",
                            "passed": True,
                            "preview_dpi": 300,
                            "min_cad_text_height_mm": 4.2,
                            "min_required_cad_text_height_mm": 3.8,
                            "min_rendered_text_height_px": 49.6,
                            "min_required_rendered_text_height_px": 40.0,
                            "median_rendered_text_height_px": 54.3,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "preview_dpi": 300,
                                    "min_cad_text_height_mm": 4.2,
                                    "min_rendered_text_height_px": 49.6,
                                    "min_required_rendered_text_height_px": 40.0,
                                }
                            ],
                        },
                        "text_integrity_audit": {
                            "audit_method": "dxf-text-entity-and-rendered-ocr-machine",
                            "passed": True,
                            "source_text_entity_count": 1,
                            "mojibake_or_missing_glyph_count": 0,
                            "missing_required_drawing_text_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "source_text_entity_count": 1,
                                    "mojibake_or_missing_glyph_count": 0,
                                    "missing_required_drawing_text_count": 0,
                                }
                            ],
                        },
                        "text_orientation_audit": {
                            "audit_method": "dxf-text-rotation-generation-flag-machine",
                            "passed": True,
                            "source_text_entity_count": 1,
                            "upside_down_text_count": 0,
                            "mirrored_text_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "source_text_entity_count": 1,
                                    "upside_down_text_count": 0,
                                    "mirrored_text_count": 0,
                                }
                            ],
                        },
                        "manufacturing_complexity_audit": {
                            "audit_method": "feature-family-and-main-view-machine",
                            "passed": True,
                            "min_feature_family_count": 8,
                            "uses_dense_filler_as_primary_depth": False,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "manufacturing_detail_families": [
                                        "section",
                                        "dimension",
                                        "bolted-joint",
                                        "weld-symbol",
                                        "surface-roughness",
                                        "geometric-tolerance",
                                        "datum-label",
                                        "detail-view",
                                    ],
                                    "main_view_enhanced": True,
                                    "uses_dense_filler_as_primary_depth": False,
                                }
                            ],
                        },
                        "hatch_clip_audit": {
                            "audit_method": "hatch-section-fill-clip-mask-machine",
                            "passed": True,
                            "section_fill_region_count": 3,
                            "hatch_clip_violation_count": 0,
                            "entity_boundary_escape_count": 0,
                            "adjacent_view_crossing_count": 0,
                            "dimension_line_crossing_count": 0,
                            "title_block_table_bom_frame_crossing_count": 0,
                            "blank_background_leak_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": True,
                                    "section_fill_region_count": 3,
                                    "hatch_clip_violation_count": 0,
                                    "entity_boundary_escape_count": 0,
                                    "adjacent_view_crossing_count": 0,
                                    "dimension_line_crossing_count": 0,
                                    "title_block_table_bom_frame_crossing_count": 0,
                                    "blank_background_leak_count": 0,
                                }
                            ],
                        },
                        "per_sheet": [
                            {
                                "sheet": "sheet",
                                "preview_path": "sheet-preview.png",
                                "no_overlap_verdict": "pass",
                                "boundary_clearance_verdict": "pass",
                                "detail_density_verdict": "pass",
                                "text_legibility_verdict": "pass",
                                "text_integrity_verdict": "pass",
                                "text_orientation_verdict": "pass",
                                "title_block_table_notes_isolation_verdict": "pass",
                                "title_block_cell_containment_verdict": "pass",
                                "title_block_short_line_topology_verdict": "pass",
                                "annotation_margin_clearance_verdict": "pass",
                                "local_crowding_verdict": "pass",
                                "layout_collision_verdict": "pass",
                                "content_overlap_verdict": "pass",
                                "outside_frame_ink_verdict": "pass",
                                "hatch_section_fill_clipping_verdict": "pass",
                            }
                        ],
                    },
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        (bad / "renamed.dwg").write_bytes(b"0\nSECTION\n")
        (bad / "bad.json").write_text('{"title":"\\u934a\\u53bb"', encoding="utf-8")
        for target in (
            entity_only,
            overlap,
            machine_missing,
            schematic_substitute,
            frame_overflow,
            tofu_text,
            inverted_text,
            mtext_attachment,
        ):
            for child in good.iterdir():
                (target / child.name).write_bytes(child.read_bytes())
        (tofu_text / "sheet.dxf").write_text(
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n0\nSTYLE\n0\nDIMSTYLE\n0\nLTYPE\n0\nENDSEC\n"
            "0\nSECTION\n2\nBLOCKS\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n"
            "0\nTEXT\n1\n\u25a1\u25a1\u25a1\u25a1\n50\n0\n71\n0\n0\nDIMENSION\n0\nENDSEC\n0\nEOF\n",
            encoding="utf-8",
        )
        (inverted_text / "sheet.dxf").write_text(
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n0\nSTYLE\n0\nDIMSTYLE\n0\nLTYPE\n0\nENDSEC\n"
            "0\nSECTION\n2\nBLOCKS\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n"
            "0\nTEXT\n1\nTITLE\n50\n180\n71\n0\n0\nDIMENSION\n0\nENDSEC\n0\nEOF\n",
            encoding="utf-8",
        )
        (mtext_attachment / "sheet.dxf").write_text(
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n0\nSTYLE\n0\nDIMSTYLE\n0\nLTYPE\n0\nENDSEC\n"
            "0\nSECTION\n2\nBLOCKS\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n"
            "0\nMTEXT\n1\nCENTERED NOTE\n11\n1\n21\n0\n71\n5\n0\nDIMENSION\n0\nENDSEC\n0\nEOF\n",
            encoding="utf-8",
        )
        write_test_sheet_png(frame_overflow / "sheet-preview.png", overflow=True)
        (schematic_substitute / "chain-cylinder-schematic-preview.png").write_bytes(
            (schematic_substitute / "sheet-preview.png").read_bytes()
        )
        (entity_only / "manifest.json").write_text('{"files":["sheet.dwg","sheet.pdf","sheet.dxf"]}', encoding="utf-8")
        (machine_missing / "manifest.json").write_text(
            json.dumps(
                {
                    "rendered_review": {
                        "preview_paths": ["sheet-preview.png"],
                        "no_overlap_verdict": "pass",
                        "boundary_clearance_verdict": "pass",
                        "detail_density_verdict": "pass",
                        "text_legibility_verdict": "pass",
                        "title_block_table_notes_isolation_verdict": "pass",
                        "annotation_margin_clearance_verdict": "pass",
                        "local_crowding_verdict": "pass",
                        "sheet_layout_verdict": "pass",
                        "manufacturing_view_depth_verdict": "pass",
                        "layout_collision_verdict": "pass",
                        "entity_count_only_verdict": "not-used",
                        "per_sheet": [
                            {
                                "sheet": "sheet",
                                "preview_path": "sheet-preview.png",
                                "no_overlap_verdict": "pass",
                                "boundary_clearance_verdict": "pass",
                                "detail_density_verdict": "pass",
                                "text_legibility_verdict": "pass",
                                "title_block_table_notes_isolation_verdict": "pass",
                                "annotation_margin_clearance_verdict": "pass",
                                "local_crowding_verdict": "pass",
                                "layout_collision_verdict": "pass",
                            }
                        ],
                    }
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        (overlap / "manifest.json").write_text(
            json.dumps(
                {
                    "rendered_review": {
                        "preview_paths": ["sheet-preview.png"],
                        "no_overlap_verdict": "fail",
                        "boundary_clearance_verdict": "fail",
                        "detail_density_verdict": "pass",
                        "text_legibility_verdict": "pass",
                        "title_block_table_notes_isolation_verdict": "fail",
                        "title_block_cell_containment_verdict": "pass",
                        "title_block_short_line_topology_verdict": "fail",
                        "annotation_margin_clearance_verdict": "pass",
                        "local_crowding_verdict": "fail",
                        "sheet_layout_verdict": "pass",
                        "manufacturing_view_depth_verdict": "pass",
                        "layout_collision_verdict": "fail",
                        "content_overlap_verdict": "fail",
                        "entity_count_only_verdict": "not-used",
                        "machine_overlap_audit": {
                            "audit_method": "bbox-and-rendered-zone-machine",
                            "passed": False,
                            "overlap_count": 2,
                            "text_entity_overlap_count": 1,
                            "reserved_zone_collision_count": 1,
                            "title_block_table_note_collision_count": 1,
                            "annotation_collision_count": 1,
                            "frame_clearance_violation_count": 0,
                            "min_clearance_mm": 0.6,
                            "max_local_ink_density": 0.86,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": False,
                                    "overlap_count": 2,
                                    "text_entity_overlap_count": 1,
                                    "reserved_zone_collision_count": 1,
                                    "title_block_table_note_collision_count": 1,
                                    "annotation_collision_count": 1,
                                    "frame_clearance_violation_count": 0,
                                    "min_clearance_mm": 0.6,
                                    "max_local_ink_density": 0.86,
                                }
                            ],
                        },
                        "content_overlap_audit": {
                            "audit_method": "content-bbox-pairwise-machine",
                            "passed": False,
                            "registered_bbox_count": 8,
                            "checked_pair_count": 28,
                            "min_clearance_mm": 0.6,
                            "content_overlap_count": 3,
                            "view_view_overlap_count": 1,
                            "detail_frame_main_view_overlap_count": 1,
                            "table_text_grid_collision_count": 1,
                            "dimension_line_view_table_crossing_count": 1,
                            "leader_line_view_table_crossing_count": 1,
                            "balloon_geometry_collision_count": 1,
                            "bbox_helper_envelope_escape_count": 1,
                            "stale_rendered_preview_count": 0,
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": False,
                                    "registered_bbox_count": 8,
                                    "checked_pair_count": 28,
                                    "min_clearance_mm": 0.6,
                                    "content_overlap_count": 3,
                                    "view_view_overlap_count": 1,
                                    "detail_frame_main_view_overlap_count": 1,
                                    "table_text_grid_collision_count": 1,
                                    "dimension_line_view_table_crossing_count": 1,
                                    "leader_line_view_table_crossing_count": 1,
                                    "balloon_geometry_collision_count": 1,
                                    "bbox_helper_envelope_escape_count": 1,
                                    "stale_rendered_preview_count": 0,
                                }
                            ],
                        },
                        "title_block_short_line_topology_audit": {
                            "audit_method": "source-pdf-dxf-short-line-topology-machine",
                            "passed": False,
                            "baseline_source": "sheet.pdf; sheet.dxf",
                            "missing_short_table_line_count": 1,
                            "broken_cell_border_count": 1,
                            "table_grid_topology_mismatch_count": 1,
                            "diagnostic_overlay_free_title_block_crop_path": "sheet-title-crop.png",
                            "per_sheet": [
                                {
                                    "sheet": "sheet",
                                    "passed": False,
                                    "missing_short_table_line_count": 1,
                                    "broken_cell_border_count": 1,
                                    "table_grid_topology_mismatch_count": 1,
                                }
                            ],
                        },
                        "per_sheet": [
                            {
                                "sheet": "sheet",
                                "preview_path": "sheet-preview.png",
                                "no_overlap_verdict": "fail",
                                "boundary_clearance_verdict": "fail",
                                "detail_density_verdict": "pass",
                                "text_legibility_verdict": "pass",
                                "title_block_table_notes_isolation_verdict": "fail",
                                "title_block_cell_containment_verdict": "pass",
                                "title_block_short_line_topology_verdict": "fail",
                                "annotation_margin_clearance_verdict": "pass",
                                "local_crowding_verdict": "fail",
                                "layout_collision_verdict": "fail",
                                "content_overlap_verdict": "fail",
                            }
                        ],
                    }
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        good_result = audit_package(
            good,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        dimension_clamp_result = audit_package(
            good,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=1,
            min_dimensioned_dxf_files=8,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        bad_result = audit_package(
            bad,
            reference_paths=[],
            require_dwg=True,
            require_pdf=False,
            strict_cad_structure=False,
            min_dwg_count=1,
            min_dxf_count=0,
            min_pdf_count=0,
            min_total_a0=0,
            min_true_dimensions=0,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        entity_only_result = audit_package(
            entity_only,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        machine_missing_result = audit_package(
            machine_missing,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        overlap_result = audit_package(
            overlap,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        schematic_substitute_result = audit_package(
            schematic_substitute,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        frame_overflow_result = audit_package(
            frame_overflow,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        tofu_text_result = audit_package(
            tofu_text,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        inverted_text_result = audit_package(
            inverted_text,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        mtext_attachment_result = audit_package(
            mtext_attachment,
            reference_paths=[],
            require_dwg=True,
            require_pdf=True,
            strict_cad_structure=True,
            min_dwg_count=1,
            min_dxf_count=1,
            min_pdf_count=1,
            min_total_a0=0,
            min_true_dimensions=1,
            min_reference_dwg_ratio=0,
            min_geometry_entities=0,
            min_insert_count=0,
            min_arc_count=0,
            min_hatch_count=0,
            min_text_count=0,
            min_dxf_entities_per_sheet=0,
            min_dimensions_per_dxf=0,
            min_dimensioned_dxf_files=0,
            min_pdf_drawing_objects_per_a0=0,
            required_text_tokens=[],
            require_rendered_review=True,
        )
        if not good_result["passed"]:
            print(json.dumps(good_result, ensure_ascii=False, indent=2))
            return 1
        if not dimension_clamp_result["passed"]:
            print(json.dumps(dimension_clamp_result, ensure_ascii=False, indent=2))
            return 1
        clamp_requirements = dimension_clamp_result.get("requirements", {})
        if not isinstance(clamp_requirements, dict) or clamp_requirements.get("effective_min_dimensioned_dxf_files") != 1:
            print(json.dumps(dimension_clamp_result, ensure_ascii=False, indent=2))
            return 1
        if bad_result["passed"]:
            print(json.dumps(bad_result, ensure_ascii=False, indent=2))
            return 1
        if entity_only_result["passed"]:
            print(json.dumps(entity_only_result, ensure_ascii=False, indent=2))
            return 1
        if machine_missing_result["passed"]:
            print(json.dumps(machine_missing_result, ensure_ascii=False, indent=2))
            return 1
        if overlap_result["passed"]:
            print(json.dumps(overlap_result, ensure_ascii=False, indent=2))
            return 1
        if schematic_substitute_result["passed"]:
            print(json.dumps(schematic_substitute_result, ensure_ascii=False, indent=2))
            return 1
        if frame_overflow_result["passed"]:
            print(json.dumps(frame_overflow_result, ensure_ascii=False, indent=2))
            return 1
        if tofu_text_result["passed"]:
            print(json.dumps(tofu_text_result, ensure_ascii=False, indent=2))
            return 1
        if inverted_text_result["passed"]:
            print(json.dumps(inverted_text_result, ensure_ascii=False, indent=2))
            return 1
        if not mtext_attachment_result["passed"]:
            print(json.dumps(mtext_attachment_result, ensure_ascii=False, indent=2))
            return 1
        mtext_quality = mtext_attachment_result.get("candidate", {}).get("cad_text_quality", {})
        if not isinstance(mtext_quality, dict) or mtext_quality.get("upside_down_text_count") != 0 or mtext_quality.get("mirrored_text_count") != 0:
            print(json.dumps(mtext_attachment_result, ensure_ascii=False, indent=2))
            return 1
    print("self-test passed")
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("package", nargs="?", type=Path)
    parser.add_argument("--reference", action="append", default=[], type=Path, help="Reference CAD/PDF package zip or folder")
    parser.add_argument("--report-json", type=Path)
    parser.add_argument("--require-dwg", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--require-pdf", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--strict-cad-structure", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--min-dwg-count", type=int, default=8)
    parser.add_argument("--min-dxf-count", type=int, default=8)
    parser.add_argument("--min-pdf-count", type=int, default=8)
    parser.add_argument("--min-total-a0", type=float, default=3.5)
    parser.add_argument("--min-true-dimensions", type=int, default=140)
    parser.add_argument("--min-reference-dwg-ratio", type=float, default=0.35)
    parser.add_argument("--min-geometry-entities", type=int, default=12000)
    parser.add_argument("--min-insert-count", type=int, default=1)
    parser.add_argument("--min-arc-count", type=int, default=1)
    parser.add_argument("--min-hatch-count", type=int, default=200)
    parser.add_argument("--min-text-count", type=int, default=500)
    parser.add_argument("--min-dxf-entities-per-sheet", type=int, default=350)
    parser.add_argument("--min-dimensions-per-dxf", type=int, default=8)
    parser.add_argument("--min-dimensioned-dxf-files", type=int, default=8)
    parser.add_argument("--min-pdf-drawing-objects-per-a0", type=int, default=800)
    parser.add_argument("--required-text-token", action="append", default=[])
    parser.add_argument("--require-rendered-review", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    if args.self_test:
        return run_self_test()
    if args.package is None:
        raise SystemExit("package path is required unless --self-test is used")
    result = audit_package(
        args.package,
        reference_paths=args.reference,
        require_dwg=args.require_dwg,
        require_pdf=args.require_pdf,
        strict_cad_structure=args.strict_cad_structure,
        min_dwg_count=args.min_dwg_count,
        min_dxf_count=args.min_dxf_count,
        min_pdf_count=args.min_pdf_count,
        min_total_a0=args.min_total_a0,
        min_true_dimensions=args.min_true_dimensions,
        min_reference_dwg_ratio=args.min_reference_dwg_ratio,
        min_geometry_entities=args.min_geometry_entities,
        min_insert_count=args.min_insert_count,
        min_arc_count=args.min_arc_count,
        min_hatch_count=args.min_hatch_count,
        min_text_count=args.min_text_count,
        min_dxf_entities_per_sheet=args.min_dxf_entities_per_sheet,
        min_dimensions_per_dxf=args.min_dimensions_per_dxf,
        min_dimensioned_dxf_files=args.min_dimensioned_dxf_files,
        min_pdf_drawing_objects_per_a0=args.min_pdf_drawing_objects_per_a0,
        required_text_tokens=args.required_text_token,
        require_rendered_review=args.require_rendered_review,
    )
    text = json.dumps(result, ensure_ascii=True, indent=2)
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
