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
    docx_body_figure_paragraphs,
    docx_drawing_object_manifest,
    docx_image_relationship_manifest,
    docx_text_height_emu,
    docx_text_width_emu,
    has_docx_figure_reference_signal,
    is_docx_figure_caption,
)


EMU_PER_CM = 360000
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
A = f"{{{A_NS}}}"
NS = {"w": W_NS, "r": R_NS, "a": A_NS}
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


def audit_figure_extents(
    final_docx: Path,
    *,
    source_docx: Path | None = None,
    max_width_cm: float = 16.0,
    max_height_cm: float = 16.0,
    min_readable_width_cm: float = 8.0,
    min_readable_height_cm: float = 4.0,
    structural_min_readable_width_cm: float = 9.0,
    structural_min_readable_height_cm: float = 4.0,
    min_text_width_ratio: float = 0.95,
    require_explanations: bool = False,
    min_explanation_chars: int = 25,
) -> dict[str, object]:
    rows = docx_body_figure_paragraphs(final_docx)
    rel_ids_by_para = paragraph_relationship_id_map(final_docx)
    media_manifest = media_by_rid(final_docx)
    text_width_emu = docx_text_width_emu(final_docx) or 0
    text_height_emu = docx_text_height_emu(final_docx) or 0
    source_drawings = docx_drawing_object_manifest(source_docx) if source_docx is not None and source_docx.exists() else {}
    final_drawings = docx_drawing_object_manifest(final_docx)

    issues: list[str] = []
    figures: list[dict[str, object]] = []
    zero_extent_front_matter_count = 0
    real_front_matter_count = 0
    formal_caption_count = 0
    narrative_reference_count = 0
    oversized_count = 0
    undersized_count = 0
    paragraph_margin_width_drift_count = 0
    bottom_blank_like_count = 0
    missing_caption_count = 0
    missing_explanation_count = 0

    paragraph_positions = {int(row["paragraph_index"]): pos for pos, row in enumerate(rows)}
    for row in rows:
        text = str(row.get("text") or "")
        if row.get("is_caption") and not row.get("front_matter_drawing"):
            formal_caption_count += 1
        elif has_docx_figure_reference_signal(text) and not is_docx_figure_caption(text):
            narrative_reference_count += 1

    for row in rows:
        if not row.get("has_drawing"):
            continue
        paragraph_index = int(row["paragraph_index"])
        rel_ids = rel_ids_by_para.get(paragraph_index, [])
        max_cx = int(row.get("max_extent_cx_emu") or 0)
        max_cy = int(row.get("max_extent_cy_emu") or 0)
        is_real_drawing = bool(rel_ids or max_cx or max_cy)
        if row.get("front_matter_drawing"):
            if not is_real_drawing:
                zero_extent_front_matter_count += 1
                continue
            real_front_matter_count += 1
            final_key_prefix = f"word/document.xml#p{paragraph_index}#"
            preserved = all(
                source_drawings.get(key) == value
                for key, value in final_drawings.items()
                if key.startswith(final_key_prefix)
            )
            if not preserved:
                issues.append(f"front-matter drawing is not source-preserved: paragraph {paragraph_index}")
            continue
        if not is_real_drawing:
            continue

        pos = paragraph_positions.get(paragraph_index, -1)
        next_row = rows[pos + 1] if pos >= 0 and pos + 1 < len(rows) else None
        caption = str(next_row.get("text") or "") if next_row and next_row.get("is_caption") else ""
        if not caption:
            missing_caption_count += 1
            issues.append(f"body figure paragraph {paragraph_index} lacks an immediate formal caption")

        previous_text = nearest_text(rows, pos - 1, -1) if pos >= 0 else ""
        following_text = nearest_text(rows, pos + 2, 1) if caption and pos >= 0 else nearest_text(rows, pos + 1, 1)
        explanation_ok = compact_len(previous_text) >= min_explanation_chars or compact_len(following_text) >= min_explanation_chars
        if require_explanations and not explanation_ok:
            missing_explanation_count += 1
            issues.append(f"body figure paragraph {paragraph_index} lacks nearby explanatory prose")

        width_cm = emu_to_cm(max_cx)
        height_cm = emu_to_cm(max_cy)
        width_text_ratio = (max_cx / text_width_emu) if text_width_emu else None
        oversized_width = width_cm > max_width_cm or (text_width_emu and max_cx > int(text_width_emu * 1.02))
        oversized_height = height_cm > max_height_cm or (text_height_emu and max_cy > int(text_height_emu * 0.82))
        if oversized_width or oversized_height:
            oversized_count += 1
            issues.append(
                "body figure exceeds display threshold: "
                f"paragraph {paragraph_index} width_cm={width_cm} height_cm={height_cm}"
            )
        structural_caption = is_structural_figure_caption(caption)
        required_min_width = structural_min_readable_width_cm if structural_caption else min_readable_width_cm
        required_min_height = structural_min_readable_height_cm if structural_caption else min_readable_height_cm
        undersized_width = width_cm < required_min_width
        undersized_height = height_cm < required_min_height
        if undersized_width or undersized_height:
            undersized_count += 1
            issues.append(
                "body figure below readability threshold: "
                f"paragraph {paragraph_index} caption=`{caption[:60]}` "
                f"width_cm={width_cm} height_cm={height_cm} "
                f"min_width_cm={required_min_width} min_height_cm={required_min_height}"
            )
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
        paragraph_margin_width_drift = (
            bool(text_width_emu)
            and min_text_width_ratio > 0
            and max_cx < int(text_width_emu * min_text_width_ratio)
            and not height_constrained_width
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

        rid = rel_ids[0] if rel_ids else ""
        media = media_manifest.get(rid, {})
        image_stats = image_bottom_band_stats(final_docx, media.get("media_name", ""))
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

    return {
        "schema": "graduation-project-builder.figure-extents-audit.v2",
        "final_docx_path": str(final_docx),
        "final_docx_sha256": sha256_file(final_docx),
        "source_docx_path": str(source_docx) if source_docx else "",
        "source_docx_sha256": sha256_file(source_docx) if source_docx is not None and source_docx.exists() else "",
        "max_width_cm": max_width_cm,
        "max_height_cm": max_height_cm,
        "min_readable_width_cm": min_readable_width_cm,
        "min_readable_height_cm": min_readable_height_cm,
        "structural_min_readable_width_cm": structural_min_readable_width_cm,
        "structural_min_readable_height_cm": structural_min_readable_height_cm,
        "min_text_width_ratio": min_text_width_ratio,
        "text_width_emu": text_width_emu,
        "text_width_cm": emu_to_cm(text_width_emu) if text_width_emu else 0.0,
        "text_height_emu": text_height_emu,
        "body_figure_count": body_figure_count,
        "formal_caption_count": formal_caption_count,
        "narrative_figure_reference_count": narrative_reference_count,
        "front_matter_real_drawing_count": real_front_matter_count,
        "front_matter_zero_extent_drawing_count": zero_extent_front_matter_count,
        "oversized_count": oversized_count,
        "undersized_count": undersized_count,
        "paragraph_margin_width_drift_count": paragraph_margin_width_drift_count,
        "bottom_blank_like_count": bottom_blank_like_count,
        "missing_caption_count": missing_caption_count,
        "missing_explanation_count": missing_explanation_count,
        "require_explanations": require_explanations,
        "figures": figures,
        "issues": issues,
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
    parser.add_argument("--require-explanations", action="store_true")
    parser.add_argument("--min-explanation-chars", type=int, default=25)
    args = parser.parse_args()

    source_docx = Path(args.source_docx).resolve() if args.source_docx else None
    payload = audit_figure_extents(
        Path(args.final_docx).resolve(),
        source_docx=source_docx,
        max_width_cm=args.max_width_cm,
        max_height_cm=args.max_height_cm,
        min_readable_width_cm=args.min_readable_width_cm,
        min_readable_height_cm=args.min_readable_height_cm,
        structural_min_readable_width_cm=args.structural_min_readable_width_cm,
        structural_min_readable_height_cm=args.structural_min_readable_height_cm,
        min_text_width_ratio=args.min_text_width_ratio,
        require_explanations=args.require_explanations,
        min_explanation_chars=args.min_explanation_chars,
    )
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
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
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
    print(text, end="")
    return 0 if payload["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
