#!/usr/bin/env python3
"""Audit DOCX figure display size, caption adjacency, and explanation coverage."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from PIL import Image, ImageStat

from thesis_figure_contract import (
    A_NS,
    ASSET_SCHEMA,
    R_NS,
    W_NS,
    WP_NS,
    docx_body_figure_paragraphs,
    docx_drawing_object_manifest,
    docx_image_relationship_manifest,
    docx_text_height_emu,
    docx_text_width_emu,
    has_docx_figure_reference_signal,
    is_docx_figure_caption,
)


EMU_PER_CM = 360000
READABLE_HEIGHT_TOLERANCE_CM = 0.10
NATIVE_PPI_TOLERANCE = 0.5
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
A = f"{{{A_NS}}}"
WP = f"{{{WP_NS}}}"
NS = {"w": W_NS, "r": R_NS, "a": A_NS, "wp": WP_NS}
BODY_CHAPTER_RE = re.compile(r"^\s*(?:第[一二三四五六七八九十百零\d]+章|[1-9]\s+)")
BACK_MATTER_RE = re.compile(r"^\s*(?:参考文献|附录|致谢)(?:\s|$|[A-Za-z0-9一二三四五六七八九十])")
APPENDIX_CAPTION_RE = re.compile(r"^\s*(?:附图|附表)\s*[A-Za-z0-9一二三四五六七八九十]*[.．、-]")
STRUCTURAL_FIGURE_TOKENS = (
    "结构图",
    "流程图",
    "链路图",
    "逻辑图",
    "架构图",
    "关系图",
    "映射图",
    "示意图",
    "ER 图",
    "ER图",
    "模型结构",
    "证据链",
)
INDENT_FIELDS = ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def emu_to_cm(value: int) -> float:
    return round(value / EMU_PER_CM, 2)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def _paragraph_style(paragraph: ET.Element) -> str:
    pstyle = paragraph.find(".//w:pPr/w:pStyle", NS)
    if pstyle is None:
        return ""
    return pstyle.attrib.get(f"{W}val", "")


def _int_or_zero(value: str) -> int:
    try:
        return int(value or "0")
    except ValueError:
        return 0


def _style_layouts(docx_path: Path) -> tuple[dict[str, dict[str, Any]], str]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/styles.xml"))
    except Exception:
        return {}, ""
    layouts: dict[str, dict[str, Any]] = {}
    default_style_id = ""
    for style in root.findall("w:style", NS):
        if style.attrib.get(f"{W}type") != "paragraph":
            continue
        style_id = style.attrib.get(f"{W}styleId", "")
        if not style_id:
            continue
        if style.attrib.get(f"{W}default") == "1":
            default_style_id = style_id
        based_on = style.find("w:basedOn", NS)
        jc = style.find("w:pPr/w:jc", NS)
        ind = style.find("w:pPr/w:ind", NS)
        layouts[style_id] = {
            "based_on": based_on.attrib.get(f"{W}val", "") if based_on is not None else "",
            "jc": jc.attrib.get(f"{W}val", "") if jc is not None else "",
            "ind": {field: ind.attrib.get(f"{W}{field}", "") for field in INDENT_FIELDS} if ind is not None else {},
        }
    return layouts, default_style_id


def _effective_style_layout(style_id: str, layouts: dict[str, dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {"jc": "", "ind": {}}
    seen: set[str] = set()

    def apply(current_id: str) -> None:
        if not current_id or current_id in seen:
            return
        seen.add(current_id)
        current = layouts.get(current_id)
        if not current:
            return
        apply(str(current.get("based_on") or ""))
        if current.get("jc"):
            result["jc"] = str(current.get("jc") or "")
        result["ind"].update({k: v for k, v in (current.get("ind") or {}).items() if v != ""})

    apply(style_id)
    return result


def paragraph_layout_info(
    paragraph: ET.Element,
    style_layouts: dict[str, dict[str, Any]],
    default_style_id: str,
) -> dict[str, object]:
    ppr = paragraph.find("./w:pPr", NS)
    style_id = _paragraph_style(paragraph) or default_style_id
    inherited = _effective_style_layout(style_id, style_layouts) if style_id else {"jc": "", "ind": {}}
    jc_node = ppr.find("./w:jc", NS) if ppr is not None else None
    ind_node = ppr.find("./w:ind", NS) if ppr is not None else None
    direct_jc = jc_node.attrib.get(f"{W}val", "") if jc_node is not None else ""
    effective_jc = direct_jc or str(inherited.get("jc") or "")
    direct_ind = {field: ind_node.attrib.get(f"{W}{field}", "") for field in INDENT_FIELDS} if ind_node is not None else {}
    inherited_ind = dict(inherited.get("ind") or {})
    effective_ind = {
        field: direct_ind[field] if field in direct_ind and direct_ind[field] != "" else inherited_ind.get(field, "")
        for field in INDENT_FIELDS
    }
    nonzero_indent_fields = {
        field: value
        for field, value in effective_ind.items()
        if value != "" and _int_or_zero(str(value)) != 0
    }
    reasons: list[str] = []
    if effective_jc and effective_jc.lower() != "center":
        reasons.append(f"effective alignment `{effective_jc}` is not center")
    if nonzero_indent_fields:
        reasons.append(f"effective indentation is nonzero: {nonzero_indent_fields}")
    return {
        "style_id": style_id,
        "direct_alignment": direct_jc,
        "effective_alignment": effective_jc,
        "direct_indent": direct_ind,
        "effective_indent": effective_ind,
        "nonzero_indent_fields": nonzero_indent_fields,
        "layout_issue_reasons": reasons,
        "safe_image_holder_layout": not reasons,
    }


def _is_toc_cache_paragraph(paragraph: ET.Element, text: str, style: str) -> bool:
    """Identify visible TOC cache rows even when a template stores them as Normal.

    Some official templates keep TOC entries visually formatted with direct
    paragraph properties instead of TOC1/TOC2 styles. Those rows can contain
    chapter/back-matter text plus a cached page number, so body-boundary scans
    must ignore them or body figures are misclassified as back matter.
    """
    if style.upper().startswith("TOC"):
        return True
    for instr in paragraph.findall(".//w:instrText", NS):
        if "TOC" in (instr.text or "").upper():
            return True
    if paragraph.find(".//w:tab", NS) is not None and text.rstrip()[-1:].isdigit():
        return True
    return False


def body_scope_bounds(docx_path: Path) -> dict[str, int | None]:
    """Return body/back-matter paragraph bounds using the main document XML.

    The figure rows returned by thesis_figure_contract intentionally include
    appendix surfaces so appendix CAD sheets can be audited elsewhere. This
    audit's body-figure gate must not count those appendix drawings.
    """
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return {"body_start_paragraph_index": None, "back_matter_start_paragraph_index": None}
    body = root.find(f".//{W}body")
    if body is None:
        return {"body_start_paragraph_index": None, "back_matter_start_paragraph_index": None}
    body_start: int | None = None
    back_start: int | None = None
    paragraphs = body.findall(f".//{W}p")
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        style = _paragraph_style(paragraph)
        if text and _is_toc_cache_paragraph(paragraph, text, style):
            continue
        if body_start is None and text and BODY_CHAPTER_RE.match(text):
            body_start = index
            continue
        if body_start is not None and text and BACK_MATTER_RE.match(text):
            back_start = index
            break
    return {
        "body_start_paragraph_index": body_start,
        "back_matter_start_paragraph_index": back_start,
    }


def is_body_figure_scope(row: dict[str, Any], bounds: dict[str, int | None]) -> bool:
    paragraph_index = int(row.get("paragraph_index") or 0)
    text = str(row.get("text") or "").strip()
    body_start = bounds.get("body_start_paragraph_index")
    back_start = bounds.get("back_matter_start_paragraph_index")
    if body_start is not None and paragraph_index < int(body_start):
        return False
    if back_start is not None and paragraph_index >= int(back_start):
        return False
    if BACK_MATTER_RE.match(text) or APPENDIX_CAPTION_RE.match(text):
        return False
    return True


def paragraph_relationship_ids(paragraph: ET.Element) -> list[str]:
    ids: list[str] = []
    for blip in paragraph.findall(f".//{A}blip"):
        rid = blip.attrib.get(f"{R}embed") or blip.attrib.get(f"{R}link") or ""
        if rid:
            ids.append(rid)
    for node in paragraph.iter():
        if not node.tag.endswith("}imagedata"):
            continue
        rid = node.attrib.get(f"{R}id") or node.attrib.get(f"{R}embed") or ""
        if rid:
            ids.append(rid)
    return ids


def paragraph_crop_rects(paragraph: ET.Element) -> list[dict[str, int]]:
    rects: list[dict[str, int]] = []
    for rect in paragraph.findall(f".//{A}srcRect"):
        crop: dict[str, int] = {}
        for key in ("l", "t", "r", "b"):
            raw_value = rect.attrib.get(key, "0") or "0"
            try:
                crop[key] = int(raw_value)
            except ValueError:
                crop[key] = 0
        rects.append(crop)
    return rects


def paragraph_spacing_info(paragraph: ET.Element, image_height_emu: int) -> dict[str, object]:
    spacing = paragraph.find("./w:pPr/w:spacing", NS)
    if spacing is None:
        return {
            "line_rule": "",
            "line_twips": 0,
            "line_height_emu": 0,
            "exact_line_spacing_clips_inline_image": False,
        }
    line_rule = spacing.attrib.get(f"{W}lineRule", "")
    try:
        line_twips = int(spacing.attrib.get(f"{W}line", "0") or "0")
    except ValueError:
        line_twips = 0
    line_height_emu = line_twips * 635
    exact_clips = bool(
        image_height_emu > 0
        and line_rule.lower() == "exact"
        and (line_height_emu <= 0 or line_height_emu < int(image_height_emu * 0.98))
    )
    return {
        "line_rule": line_rule,
        "line_twips": line_twips,
        "line_height_emu": line_height_emu,
        "exact_line_spacing_clips_inline_image": exact_clips,
    }


def paragraph_drawing_mode_counts(paragraph: ET.Element) -> dict[str, int]:
    return {
        "inline_count": len(paragraph.findall(f".//{WP}inline")),
        "anchor_count": len(paragraph.findall(f".//{WP}anchor")),
    }


def paragraph_display_safety_map(docx_path: Path) -> dict[int, dict[str, object]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f".//{W}body")
    if body is None:
        return {}
    result: dict[int, dict[str, object]] = {}
    style_layouts, default_style_id = _style_layouts(docx_path)
    for index, paragraph in enumerate(body.findall(f".//{W}p"), start=1):
        crop_rects = paragraph_crop_rects(paragraph)
        nonzero_crop_rects = [rect for rect in crop_rects if any(value != 0 for value in rect.values())]
        result[index] = {
            "crop_rects": crop_rects,
            "nonzero_crop_rects": nonzero_crop_rects,
            "paragraph_layout": paragraph_layout_info(paragraph, style_layouts, default_style_id),
            **paragraph_drawing_mode_counts(paragraph),
        }
    return result


def paragraph_relationship_id_map(docx_path: Path) -> dict[int, list[str]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f".//{W}body")
    if body is None:
        return {}
    result: dict[int, list[str]] = {}
    for index, paragraph in enumerate(body.findall(f".//{W}p"), start=1):
        ids = paragraph_relationship_ids(paragraph)
        if ids:
            result[index] = ids
    return result


def media_by_rid(docx_path: Path) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for entry in docx_image_relationship_manifest(docx_path).values():
        if entry.get("source_part") == "word/document.xml" and entry.get("rid"):
            result[entry["rid"]] = entry
    return result


def image_bottom_band_stats(docx_path: Path, media_name: str) -> dict[str, object]:
    if not media_name:
        return {}
    try:
        with zipfile.ZipFile(docx_path) as zf:
            payload = zf.read(media_name)
    except Exception as exc:
        return {"media_error": str(exc)}
    try:
        import io

        with Image.open(io.BytesIO(payload)) as image:
            converted = image.convert("RGB")
            width, height = converted.size
            band_height = max(1, int(height * 0.12))
            band = converted.crop((0, max(0, height - band_height), width, height))
            stat = ImageStat.Stat(band)
            mean_rgb = round(sum(stat.mean) / 3, 2)
            max_stddev = round(max(stat.stddev), 2)
            return {
                "pixel_width": width,
                "pixel_height": height,
                "bottom_band_height_px": band_height,
                "bottom_band_mean_rgb": mean_rgb,
                "bottom_band_max_stddev": max_stddev,
                "bottom_blank_like": bool(mean_rgb >= 245 and max_stddev <= 3.0),
            }
    except Exception as exc:
        return {"media_error": str(exc)}


def compact_len(text: str) -> int:
    return len("".join((text or "").split()))


def is_structural_figure_caption(caption: str) -> bool:
    normalized = re.sub(r"\s+", "", caption or "").lower()
    if any(token.lower().replace(" ", "") in normalized for token in STRUCTURAL_FIGURE_TOKENS):
        return True
    return any(token in normalized for token in ("pipeline", "flow", "architecture", "structure", "diagram", "chain"))


def nearest_text(rows: list[dict[str, Any]], start: int, step: int) -> str:
    index = start
    while 0 <= index < len(rows):
        row = rows[index]
        text = str(row.get("text") or "").strip()
        if text and not row.get("has_drawing") and not row.get("is_caption"):
            return text
        index += step
    return ""


def media_sha_set(media_signature: str) -> set[str]:
    values: set[str] = set()
    for item in (media_signature or "").split(";"):
        parts = item.split("|")
        if len(parts) >= 3 and parts[-1].strip():
            values.add(parts[-1].strip().lower())
    return values


def drawing_matches_template_donor(final_drawing: dict[str, Any], template_drawing: dict[str, Any]) -> bool:
    """Match front-matter donor drawings while ignoring local relationship ids."""
    for field in ("story_part", "paragraph_index", "drawing_kind", "extent_signature"):
        if str(final_drawing.get(field, "")) != str(template_drawing.get(field, "")):
            return False
    final_hashes = media_sha_set(str(final_drawing.get("media_signature", "")))
    template_hashes = media_sha_set(str(template_drawing.get("media_signature", "")))
    if final_hashes or template_hashes:
        return bool(final_hashes) and final_hashes == template_hashes
    return str(final_drawing.get("sha256", "")) == str(template_drawing.get("sha256", ""))


def front_matter_drawings_preserved_from_template(
    final_rows: dict[str, dict[str, Any]],
    template_drawings: dict[str, dict[str, Any]],
) -> bool:
    if not final_rows or not template_drawings:
        return False
    for final_drawing in final_rows.values():
        if not any(drawing_matches_template_donor(final_drawing, template_drawing) for template_drawing in template_drawings.values()):
            return False
    return True


def audit_figure_extents(
    final_docx: Path,
    *,
    source_docx: Path | None = None,
    template_docx: Path | None = None,
    max_width_cm: float = 16.0,
    max_height_cm: float = 16.0,
    min_readable_width_cm: float = 8.0,
    min_readable_height_cm: float = 4.0,
    structural_min_readable_width_cm: float = 9.0,
    structural_min_readable_height_cm: float = 4.0,
    min_text_width_ratio: float = 0.95,
    min_native_ppi: float = 0.0,
    enforce_min_native_ppi: bool = False,
    native_ppi_target_media_name: str = "",
    native_ppi_target_caption_contains: str = "",
    require_explanations: bool = False,
    min_explanation_chars: int = 25,
    min_body_figure_count: int | None = None,
) -> dict[str, object]:
    rows = docx_body_figure_paragraphs(final_docx)
    scope_bounds = body_scope_bounds(final_docx)
    for row in rows:
        row["body_scope"] = is_body_figure_scope(row, scope_bounds)
    rel_ids_by_para = paragraph_relationship_id_map(final_docx)
    display_safety_by_para = paragraph_display_safety_map(final_docx)
    media_manifest = media_by_rid(final_docx)
    text_width_emu = docx_text_width_emu(final_docx) or 0
    text_height_emu = docx_text_height_emu(final_docx) or 0
    source_drawings = docx_drawing_object_manifest(source_docx) if source_docx is not None and source_docx.exists() else {}
    template_drawings = (
        docx_drawing_object_manifest(template_docx)
        if template_docx is not None and template_docx.exists()
        else {}
    )
    final_drawings = docx_drawing_object_manifest(final_docx)
    try:
        with zipfile.ZipFile(final_docx) as zf:
            document_root = ET.fromstring(zf.read("word/document.xml"))
            document_body = document_root.find(f".//{W}body")
            document_paragraphs = document_body.findall(f".//{W}p") if document_body is not None else []
    except Exception:
        document_paragraphs = []

    issues: list[str] = []
    figures: list[dict[str, object]] = []
    all_drawing_holders: list[dict[str, object]] = []
    zero_extent_front_matter_count = 0
    real_front_matter_count = 0
    formal_caption_count = 0
    narrative_reference_count = 0
    oversized_count = 0
    undersized_count = 0
    paragraph_margin_width_drift_count = 0
    nonzero_crop_count = 0
    exact_line_spacing_clip_count = 0
    image_holder_layout_issue_count = 0
    image_holder_noncenter_alignment_count = 0
    image_holder_abnormal_indent_count = 0
    safe_page_height_risk_count = 0
    bottom_blank_like_count = 0
    native_ppi_issue_count = 0
    native_resolution_constrained_width_count = 0
    missing_caption_count = 0
    missing_explanation_count = 0
    template_preserved_front_matter_count = 0
    non_body_figure_count = 0
    back_matter_figure_count = 0
    back_matter_caption_count = 0

    paragraph_positions = {int(row["paragraph_index"]): pos for pos, row in enumerate(rows)}
    for row in rows:
        text = str(row.get("text") or "")
        if row.get("is_caption") and not row.get("body_scope"):
            back_matter_caption_count += 1
        elif row.get("body_scope") and has_docx_figure_reference_signal(text) and not is_docx_figure_caption(text):
            narrative_reference_count += 1

    for row in rows:
        if not row.get("has_drawing"):
            continue
        paragraph_index = int(row["paragraph_index"])
        rel_ids = rel_ids_by_para.get(paragraph_index, [])
        max_cx = int(row.get("max_extent_cx_emu") or 0)
        max_cy = int(row.get("max_extent_cy_emu") or 0)
        is_real_drawing = bool(rel_ids or max_cx or max_cy)
        paragraph = (
            document_paragraphs[paragraph_index - 1]
            if 1 <= paragraph_index <= len(document_paragraphs)
            else None
        )
        display_safety = display_safety_by_para.get(paragraph_index, {})
        layout_info = dict(display_safety.get("paragraph_layout") or {})
        crop_rects = list(display_safety.get("crop_rects") or [])
        nonzero_crop_rects = list(display_safety.get("nonzero_crop_rects") or [])
        spacing_info = paragraph_spacing_info(paragraph, max_cy) if paragraph is not None else {}
        if is_real_drawing:
            holder_scope = (
                "front-matter"
                if row.get("front_matter_drawing")
                else ("body" if row.get("body_scope") else "non-body")
            )
            holder_record = {
                "paragraph_index": paragraph_index,
                "scope": holder_scope,
                "text_prefix": text[:120],
                "width_cm": emu_to_cm(max_cx),
                "height_cm": emu_to_cm(max_cy),
                "cx_emu": max_cx,
                "cy_emu": max_cy,
                "inline_count": display_safety.get("inline_count", 0),
                "anchor_count": display_safety.get("anchor_count", 0),
                "crop_rects": crop_rects,
                "nonzero_crop_rects": nonzero_crop_rects,
                "nonzero_crop": bool(nonzero_crop_rects),
                "line_rule": spacing_info.get("line_rule", ""),
                "line_twips": spacing_info.get("line_twips", 0),
                "exact_line_spacing_clips_inline_image": bool(spacing_info.get("exact_line_spacing_clips_inline_image")),
                "holder_style_id": layout_info.get("style_id", ""),
                "holder_direct_alignment": layout_info.get("direct_alignment", ""),
                "holder_effective_alignment": layout_info.get("effective_alignment", ""),
                "holder_direct_indent": layout_info.get("direct_indent", {}),
                "holder_effective_indent": layout_info.get("effective_indent", {}),
                "holder_nonzero_indent_fields": layout_info.get("nonzero_indent_fields", {}),
                "image_holder_layout_issue": not bool(layout_info.get("safe_image_holder_layout", True)),
                "holder_layout_issue_reasons": layout_info.get("layout_issue_reasons", []),
            }
            all_drawing_holders.append(holder_record)
            if nonzero_crop_rects:
                nonzero_crop_count += 1
                issues.append(
                    f"{holder_scope} image holder has nonzero picture crop a:srcRect: "
                    f"paragraph {paragraph_index} crop_rects={nonzero_crop_rects}"
                )
            if spacing_info.get("exact_line_spacing_clips_inline_image"):
                exact_line_spacing_clip_count += 1
                issues.append(
                    f"{holder_scope} image holder uses exact line spacing smaller than image height: "
                    f"paragraph {paragraph_index} line_twips={spacing_info.get('line_twips')} "
                    f"image_height_cm={emu_to_cm(max_cy)}"
                )
            if layout_info and not layout_info.get("safe_image_holder_layout", True):
                image_holder_layout_issue_count += 1
                effective_alignment = str(layout_info.get("effective_alignment") or "")
                nonzero_indent_fields = dict(layout_info.get("nonzero_indent_fields") or {})
                if effective_alignment and effective_alignment.lower() != "center":
                    image_holder_noncenter_alignment_count += 1
                if nonzero_indent_fields:
                    image_holder_abnormal_indent_count += 1
                issues.append(
                    f"{holder_scope} image holder has unsafe paragraph layout: "
                    f"paragraph {paragraph_index} "
                    f"effective_alignment=`{effective_alignment or '<missing>'}` "
                    f"nonzero_indent_fields={nonzero_indent_fields}"
                )
        if row.get("front_matter_drawing"):
            if not is_real_drawing:
                zero_extent_front_matter_count += 1
                continue
            real_front_matter_count += 1
            final_key_prefix = f"word/document.xml#p{paragraph_index}#"
            final_rows = {key: value for key, value in final_drawings.items() if key.startswith(final_key_prefix)}
            preserved = all(source_drawings.get(key) == value for key, value in final_rows.items())
            if not preserved and front_matter_drawings_preserved_from_template(final_rows, template_drawings):
                preserved = True
                template_preserved_front_matter_count += 1
            if not preserved:
                issues.append(f"front-matter drawing is not source-preserved: paragraph {paragraph_index}")
            continue
        if not row.get("body_scope"):
            if is_real_drawing:
                non_body_figure_count += 1
                if not row.get("front_matter_drawing"):
                    back_matter_figure_count += 1
            continue
        if not is_real_drawing:
            continue

        pos = paragraph_positions.get(paragraph_index, -1)
        next_row = rows[pos + 1] if pos >= 0 and pos + 1 < len(rows) else None
        caption = (
            str(next_row.get("text") or "")
            if next_row and next_row.get("is_caption") and next_row.get("body_scope")
            else ""
        )
        if not caption:
            missing_caption_count += 1
            issues.append(f"body figure paragraph {paragraph_index} lacks an immediate formal caption")
        else:
            formal_caption_count += 1

        previous_text = nearest_text(rows, pos - 1, -1) if pos >= 0 else ""
        following_text = nearest_text(rows, pos + 2, 1) if caption and pos >= 0 else nearest_text(rows, pos + 1, 1)
        explanation_ok = compact_len(previous_text) >= min_explanation_chars or compact_len(following_text) >= min_explanation_chars
        if require_explanations and not explanation_ok:
            missing_explanation_count += 1
            issues.append(f"body figure paragraph {paragraph_index} lacks nearby explanatory prose")

        width_cm = emu_to_cm(max_cx)
        height_cm = emu_to_cm(max_cy)
        width_text_ratio = (max_cx / text_width_emu) if text_width_emu else None
        rid = rel_ids[0] if rel_ids else ""
        media = media_manifest.get(rid, {})
        image_stats = image_bottom_band_stats(final_docx, media.get("media_name", ""))
        pixel_width = int(image_stats.get("pixel_width") or 0)
        pixel_height = int(image_stats.get("pixel_height") or 0)
        inserted_ppi_width_raw = pixel_width / (width_cm / 2.54) if pixel_width and width_cm else 0.0
        inserted_ppi_height_raw = pixel_height / (height_cm / 2.54) if pixel_height and height_cm else 0.0
        inserted_ppi_width = round(inserted_ppi_width_raw, 1) if inserted_ppi_width_raw else 0.0
        inserted_ppi_height = round(inserted_ppi_height_raw, 1) if inserted_ppi_height_raw else 0.0
        text_width_cm = emu_to_cm(text_width_emu) if text_width_emu else 0.0
        native_ppi_at_text_width = (
            round(pixel_width / (text_width_cm / 2.54), 1)
            if pixel_width and text_width_cm
            else 0.0
        )
        native_resolution_constrained_width = bool(
            min_native_ppi > 0
            and pixel_width > 0
            and native_ppi_at_text_width > 0
            and native_ppi_at_text_width < min_native_ppi
        )
        media_name_for_target = str(media.get("media_name") or "").replace("\\", "/")
        native_ppi_target_media_name_norm = native_ppi_target_media_name.replace("\\", "/").strip()
        native_ppi_target_caption_contains_norm = native_ppi_target_caption_contains.strip()
        native_ppi_target_match = bool(
            not native_ppi_target_media_name_norm
            and not native_ppi_target_caption_contains_norm
        )
        if native_ppi_target_media_name_norm:
            native_ppi_target_match = (
                media_name_for_target == native_ppi_target_media_name_norm
                or media_name_for_target.endswith("/" + native_ppi_target_media_name_norm)
                or Path(media_name_for_target).name == native_ppi_target_media_name_norm
            )
        if native_ppi_target_caption_contains_norm and native_ppi_target_caption_contains_norm in caption:
            native_ppi_target_match = True
        native_ppi_issue = bool(
            enforce_min_native_ppi
            and native_ppi_target_match
            and min_native_ppi > 0
            and pixel_width > 0
            and pixel_height > 0
            and min(inserted_ppi_width_raw, inserted_ppi_height_raw) < (min_native_ppi - NATIVE_PPI_TOLERANCE)
        )
        if native_resolution_constrained_width:
            native_resolution_constrained_width_count += 1
        if native_ppi_issue:
            native_ppi_issue_count += 1
            issues.append(
                "body figure below native-resolution readability threshold: "
                f"paragraph {paragraph_index} caption=`{caption[:60]}` "
                f"pixels={pixel_width}x{pixel_height} "
                f"inserted_ppi_width={inserted_ppi_width} inserted_ppi_height={inserted_ppi_height} "
                f"min_native_ppi={min_native_ppi}"
            )
        oversized_width = width_cm > max_width_cm or (text_width_emu and max_cx > int(text_width_emu * 1.02))
        oversized_height = height_cm > max_height_cm or (text_height_emu and max_cy > int(text_height_emu * 0.82))
        if oversized_width or oversized_height:
            oversized_count += 1
            if oversized_height:
                safe_page_height_risk_count += 1
            issues.append(
                "body figure exceeds display threshold: "
                f"paragraph {paragraph_index} width_cm={width_cm} height_cm={height_cm}"
            )
        structural_caption = is_structural_figure_caption(caption)
        required_min_width = structural_min_readable_width_cm if structural_caption else min_readable_width_cm
        required_min_height = structural_min_readable_height_cm if structural_caption else min_readable_height_cm
        scaled_height_at_target_width = 0
        if max_cx > 0 and max_cy > 0 and text_width_emu and min_text_width_ratio > 0:
            scaled_height_at_target_width = int(max_cy * ((text_width_emu * min_text_width_ratio) / max_cx))
        height_constrained_width = bool(
            scaled_height_at_target_width
            and (
                scaled_height_at_target_width > int(max_height_cm * EMU_PER_CM * 1.005)
                or (text_height_emu and scaled_height_at_target_width > int(text_height_emu * 0.82))
            )
        )
        undersized_width = width_cm < required_min_width and not height_constrained_width
        undersized_height = (height_cm + READABLE_HEIGHT_TOLERANCE_CM) < required_min_height
        if undersized_width or undersized_height:
            undersized_count += 1
            issues.append(
                "body figure below readability threshold: "
                f"paragraph {paragraph_index} caption=`{caption[:60]}` "
                f"width_cm={width_cm} height_cm={height_cm} "
                f"min_width_cm={required_min_width} min_height_cm={required_min_height}"
            )
        paragraph_margin_width_drift = (
            bool(text_width_emu)
            and min_text_width_ratio > 0
            and max_cx < int(text_width_emu * min_text_width_ratio)
            and not height_constrained_width
            and not (native_resolution_constrained_width and not native_ppi_issue)
        )
        if paragraph_margin_width_drift:
            paragraph_margin_width_drift_count += 1
            ratio_text = f"{width_text_ratio:.3f}" if width_text_ratio is not None else "missing"
            issues.append(
                "body figure below paragraph-margin width: "
                f"paragraph {paragraph_index} caption=`{caption[:60]}` "
                f"width_cm={width_cm} text_width_cm={emu_to_cm(text_width_emu)} "
                f"width_text_ratio={ratio_text} min_text_width_ratio={min_text_width_ratio}"
            )

        if image_stats.get("bottom_blank_like"):
            bottom_blank_like_count += 1
            issues.append(f"body figure paragraph {paragraph_index} has blank-like bottom band")

        figures.append(
            {
                "paragraph_index": paragraph_index,
                "rid": rid,
                "media": media.get("media_name", ""),
                "media_sha256": media.get("sha256", ""),
                "caption": caption,
                "previous_explanation_prefix": previous_text[:160],
                "following_explanation_prefix": following_text[:160],
                "nearby_explanation_ok": explanation_ok,
                "width_cm": width_cm,
                "height_cm": height_cm,
                "cx_emu": max_cx,
                "cy_emu": max_cy,
                "text_width_cm": emu_to_cm(text_width_emu) if text_width_emu else 0.0,
                "width_text_ratio": width_text_ratio,
                "native_ppi_min_threshold": min_native_ppi,
                "native_ppi_tolerance": NATIVE_PPI_TOLERANCE,
                "native_ppi_enforced": enforce_min_native_ppi,
                "native_ppi_target_match": native_ppi_target_match,
                "inserted_ppi_width": inserted_ppi_width,
                "inserted_ppi_height": inserted_ppi_height,
                "native_ppi_at_text_width": native_ppi_at_text_width,
                "native_resolution_constrained_width": native_resolution_constrained_width,
                "native_ppi_issue": native_ppi_issue,
                "inline_count": display_safety.get("inline_count", 0),
                "anchor_count": display_safety.get("anchor_count", 0),
                "crop_rects": crop_rects,
                "nonzero_crop_rects": nonzero_crop_rects,
                "nonzero_crop": bool(nonzero_crop_rects),
                "line_rule": spacing_info.get("line_rule", ""),
                "line_twips": spacing_info.get("line_twips", 0),
                "exact_line_spacing_clips_inline_image": bool(spacing_info.get("exact_line_spacing_clips_inline_image")),
                "holder_style_id": layout_info.get("style_id", ""),
                "holder_direct_alignment": layout_info.get("direct_alignment", ""),
                "holder_effective_alignment": layout_info.get("effective_alignment", ""),
                "holder_direct_indent": layout_info.get("direct_indent", {}),
                "holder_effective_indent": layout_info.get("effective_indent", {}),
                "holder_nonzero_indent_fields": layout_info.get("nonzero_indent_fields", {}),
                "image_holder_layout_issue": not bool(layout_info.get("safe_image_holder_layout", True)),
                "holder_layout_issue_reasons": layout_info.get("layout_issue_reasons", []),
                "height_constrained_width": height_constrained_width,
                "scaled_height_at_target_width_cm": round(emu_to_cm(scaled_height_at_target_width), 2) if scaled_height_at_target_width else 0,
                "structural_caption": structural_caption,
                "min_readable_width_cm": required_min_width,
                "min_readable_height_cm": required_min_height,
                "min_text_width_ratio": min_text_width_ratio,
                "oversized_width": oversized_width,
                "oversized_height": oversized_height,
                "undersized_width": undersized_width,
                "undersized_height": undersized_height,
                "paragraph_margin_width_drift": paragraph_margin_width_drift,
                **image_stats,
            }
        )

    body_figure_count = len(figures)
    if body_figure_count != formal_caption_count:
        issues.append(
            "body figure/caption count mismatch after formal-caption filtering: "
            f"body_figures={body_figure_count} formal_captions={formal_caption_count}"
        )
    if min_body_figure_count is not None and body_figure_count < max(0, min_body_figure_count):
        issues.append(
            "body figure count below required minimum: "
            f"body_figures={body_figure_count} min_body_figure_count={max(0, min_body_figure_count)}"
        )

    return {
        "schema": "graduation-project-builder.figure-extents-audit.v3",
        "final_docx_path": str(final_docx),
        "final_docx_sha256": sha256_file(final_docx),
        "source_docx_path": str(source_docx) if source_docx else "",
        "source_docx_sha256": sha256_file(source_docx) if source_docx is not None and source_docx.exists() else "",
        "template_docx_path": str(template_docx) if template_docx else "",
        "template_docx_sha256": sha256_file(template_docx) if template_docx is not None and template_docx.exists() else "",
        "max_width_cm": max_width_cm,
        "max_height_cm": max_height_cm,
        "min_readable_width_cm": min_readable_width_cm,
        "min_readable_height_cm": min_readable_height_cm,
        "structural_min_readable_width_cm": structural_min_readable_width_cm,
        "structural_min_readable_height_cm": structural_min_readable_height_cm,
        "readable_height_tolerance_cm": READABLE_HEIGHT_TOLERANCE_CM,
        "min_text_width_ratio": min_text_width_ratio,
        "min_native_ppi": min_native_ppi,
        "enforce_min_native_ppi": enforce_min_native_ppi,
        "native_ppi_target_media_name": native_ppi_target_media_name,
        "native_ppi_target_caption_contains": native_ppi_target_caption_contains,
        "text_width_emu": text_width_emu,
        "text_width_cm": emu_to_cm(text_width_emu) if text_width_emu else 0.0,
        "text_height_emu": text_height_emu,
        "body_start_paragraph_index": scope_bounds.get("body_start_paragraph_index"),
        "back_matter_start_paragraph_index": scope_bounds.get("back_matter_start_paragraph_index"),
        "body_figure_count": body_figure_count,
        "non_body_figure_count": non_body_figure_count,
        "back_matter_figure_count": back_matter_figure_count,
        "formal_caption_count": formal_caption_count,
        "back_matter_caption_count": back_matter_caption_count,
        "narrative_figure_reference_count": narrative_reference_count,
        "front_matter_real_drawing_count": real_front_matter_count,
        "front_matter_zero_extent_drawing_count": zero_extent_front_matter_count,
        "template_preserved_front_matter_count": template_preserved_front_matter_count,
        "all_drawing_holder_count": len(all_drawing_holders),
        "all_drawing_holder_layout_issue_count": image_holder_layout_issue_count,
        "all_drawing_holder_noncenter_alignment_count": image_holder_noncenter_alignment_count,
        "all_drawing_holder_abnormal_indent_count": image_holder_abnormal_indent_count,
        "all_drawing_nonzero_crop_count": nonzero_crop_count,
        "all_drawing_exact_line_spacing_clip_count": exact_line_spacing_clip_count,
        "oversized_count": oversized_count,
        "undersized_count": undersized_count,
        "native_ppi_issue_count": native_ppi_issue_count,
        "native_resolution_constrained_width_count": native_resolution_constrained_width_count,
        "paragraph_margin_width_drift_count": paragraph_margin_width_drift_count,
        "nonzero_crop_count": nonzero_crop_count,
        "exact_line_spacing_clip_count": exact_line_spacing_clip_count,
        "image_holder_layout_issue_count": image_holder_layout_issue_count,
        "image_holder_noncenter_alignment_count": image_holder_noncenter_alignment_count,
        "image_holder_abnormal_indent_count": image_holder_abnormal_indent_count,
        "safe_page_height_risk_count": safe_page_height_risk_count,
        "image_holder_layout_verdict": "pass" if image_holder_layout_issue_count == 0 else "fail",
        "visible_content_completeness_verdict": "pass" if nonzero_crop_count == 0 and exact_line_spacing_clip_count == 0 and safe_page_height_risk_count == 0 and image_holder_layout_issue_count == 0 else "fail",
        "bottom_blank_like_count": bottom_blank_like_count,
        "missing_caption_count": missing_caption_count,
        "missing_explanation_count": missing_explanation_count,
        "require_explanations": require_explanations,
        "min_body_figure_count": min_body_figure_count,
        "all_drawing_holders": all_drawing_holders,
        "figures": figures,
        "issues": issues,
        "native_resolution_verdict": "pass" if native_ppi_issue_count == 0 else "fail",
        "passed": not issues,
    }


def build_source_preserved_figure_manifest(
    audit_payload: dict[str, Any],
    *,
    task_card_path: str = "",
    rendered_evidence_path: str = "",
    template_docx: str = "",
) -> dict[str, Any]:
    final_docx_path = str(audit_payload.get("final_docx_path") or "")
    final_docx_sha256 = str(audit_payload.get("final_docx_sha256") or "")
    source_docx_path = str(audit_payload.get("source_docx_path") or "")
    source_docx_sha256 = str(audit_payload.get("source_docx_sha256") or "")
    manifest: dict[str, Any] = {
        "schema": ASSET_SCHEMA,
        "source_docx_path": source_docx_path,
        "source_docx_sha256": source_docx_sha256,
        "final_docx_path": final_docx_path,
        "final_docx_sha256": final_docx_sha256,
        "source_docx_role": "source-preserved-existing-figures",
        "figures": {},
        "tables": {},
        "diagrams": {},
    }
    for index, figure in enumerate(audit_payload.get("figures") or [], start=1):
        if not isinstance(figure, dict):
            continue
        media = str(figure.get("media") or "")
        rid = str(figure.get("rid") or "")
        media_sha = str(figure.get("media_sha256") or "")
        manifest["figures"][f"figure_{index}"] = {
            "caption": str(figure.get("caption") or ""),
            "family": "preserved-existing",
            "source_kind": "source-preserved",
            "caption_to_asset_mapping": media,
            "task_card": task_card_path or final_docx_path,
            "post_insertion_rendered_evidence": rendered_evidence_path or final_docx_path,
            "template_sample_baseline": template_docx or final_docx_path,
            "rendered_page_status": "pass-existing-page-rendered",
            "insertion_status": "pass-existing-figure-present",
            "mutation_intent": "no_image_mutation",
            "preservation_status": "no_image_mutation",
            "relationship_id": rid,
            "final_relationship_id": rid,
            "media_sha256": media_sha,
            "final_media_sha256": media_sha,
            "final_media_target": media,
            "final_docx_relationship_evidence": str(audit_payload.get("final_docx_path") or ""),
        }
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--source-docx")
    parser.add_argument("--report-json")
    parser.add_argument("--figure-manifest-json")
    parser.add_argument("--figure-task-card-path")
    parser.add_argument("--figure-rendered-evidence-path")
    parser.add_argument("--template-docx")
    parser.add_argument("--max-width-cm", type=float, default=16.0)
    parser.add_argument("--max-height-cm", type=float, default=16.0)
    parser.add_argument("--min-readable-width-cm", type=float, default=8.0)
    parser.add_argument("--min-readable-height-cm", type=float, default=4.0)
    parser.add_argument("--structural-min-readable-width-cm", type=float, default=9.0)
    parser.add_argument("--structural-min-readable-height-cm", type=float, default=4.0)
    parser.add_argument("--min-text-width-ratio", type=float, default=0.95)
    parser.add_argument("--min-native-ppi", type=float, default=0.0)
    parser.add_argument("--enforce-min-native-ppi", action="store_true")
    parser.add_argument("--native-ppi-target-media-name", default="")
    parser.add_argument("--native-ppi-target-caption-contains", default="")
    parser.add_argument("--require-explanations", action="store_true")
    parser.add_argument("--min-explanation-chars", type=int, default=25)
    parser.add_argument("--min-body-figure-count", type=int)
    args = parser.parse_args()

    source_docx = Path(args.source_docx).resolve() if args.source_docx else None
    template_docx = Path(args.template_docx).resolve() if args.template_docx else None
    payload = audit_figure_extents(
        Path(args.final_docx).resolve(),
        source_docx=source_docx,
        template_docx=template_docx,
        max_width_cm=args.max_width_cm,
        max_height_cm=args.max_height_cm,
        min_readable_width_cm=args.min_readable_width_cm,
        min_readable_height_cm=args.min_readable_height_cm,
        structural_min_readable_width_cm=args.structural_min_readable_width_cm,
        structural_min_readable_height_cm=args.structural_min_readable_height_cm,
        min_text_width_ratio=args.min_text_width_ratio,
        min_native_ppi=args.min_native_ppi,
        enforce_min_native_ppi=args.enforce_min_native_ppi,
        native_ppi_target_media_name=args.native_ppi_target_media_name,
        native_ppi_target_caption_contains=args.native_ppi_target_caption_contains,
        require_explanations=args.require_explanations,
        min_explanation_chars=args.min_explanation_chars,
        min_body_figure_count=args.min_body_figure_count,
    )
    text = json.dumps(payload, ensure_ascii=True, indent=2) + "\n"
    if args.report_json:
        report = Path(args.report_json).resolve()
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(text, encoding="utf-8")
    if args.figure_manifest_json:
        manifest_path = Path(args.figure_manifest_json).resolve()
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(
            json.dumps(
                build_source_preserved_figure_manifest(
                    payload,
                    task_card_path=str(Path(args.figure_task_card_path).resolve()) if args.figure_task_card_path else "",
                    rendered_evidence_path=str(Path(args.figure_rendered_evidence_path).resolve()) if args.figure_rendered_evidence_path else "",
                    template_docx=str(Path(args.template_docx).resolve()) if args.template_docx else "",
                ),
                ensure_ascii=True,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
    print(text, end="")
    return 0 if payload["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
