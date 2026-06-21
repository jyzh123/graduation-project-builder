#!/usr/bin/env python3
"""Repair figure drawing extents and image-holder paragraph safety.

This is a bounded DOCX package repair. It edits only DrawingML extent values in
``word/document.xml`` for body figure-holder paragraphs that are below the
configured text-width ratio, whose displayed height is below the readable
minimum, or whose displayed height already exceeds the configured absolute or
page-relative safe height. It also resets unsafe exact line spacing on
image-holder paragraphs in ``word/document.xml`` when the fixed line height is
smaller than the inline image extent, because that Word/WPS paragraph state can
clip the visible image body even when the drawing extent itself is correct. It
also normalizes image-holder paragraph layout so first-line/left/right/hanging
indent and body justification cannot shift the displayed image. Front-matter and
template-donor drawings are not resized, but their holder paragraphs are still
repaired because source preservation does not exempt holder-layout safety.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import io
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from PIL import Image

from thesis_figure_contract import (
    docx_body_figure_paragraphs,
    docx_image_relationship_manifest,
    docx_text_height_emu,
    docx_text_width_emu,
    paragraph_image_relationship_ids,
)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_CM = 360000
INDENT_FIELDS = ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars")

NS = {"w": W_NS, "wp": WP_NS, "a": A_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def write_docx_with_document_xml(input_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "pkg"
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp)
        (tmp / "word" / "document.xml").write_bytes(document_xml)
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmp)
        output_docx.with_suffix(".zip").replace(output_docx)


def figure_resize_plan(
    docx_path: Path,
    *,
    min_text_width_ratio: float,
    min_height_cm: float,
    max_height_cm: float,
    max_height_ratio: float,
    min_native_ppi: float = 0.0,
    target_media_name: str = "",
    target_caption_contains: str = "",
) -> dict[int, dict[str, int | float | str]]:
    text_width = docx_text_width_emu(docx_path) or 0
    text_height = docx_text_height_emu(docx_path) or 0
    min_cx = int(text_width * min_text_width_ratio) if text_width else 0
    min_cy = int(min_height_cm * EMU_PER_CM) if min_height_cm else 0
    max_cx = int(text_width) if text_width else 0
    max_safe_cy_candidates = []
    if text_height and max_height_ratio:
        max_safe_cy_candidates.append(int(text_height * max_height_ratio))
    if max_height_cm:
        max_safe_cy_candidates.append(int(max_height_cm * EMU_PER_CM))
    max_safe_cy = min(max_safe_cy_candidates) if max_safe_cy_candidates else 0
    plan: dict[int, dict[str, int | float | str]] = {}
    if not min_cx:
        return plan
    paragraph_media = paragraph_media_pixel_info(docx_path)
    target_media_name = target_media_name.replace("\\", "/").strip()
    target_caption_contains = target_caption_contains.strip()
    for row in docx_body_figure_paragraphs(docx_path):
        if not row.get("has_drawing") or row.get("is_caption") or row.get("front_matter_drawing"):
            continue
        current_cx = int(row.get("max_extent_cx_emu") or 0)
        current_cy = int(row.get("max_extent_cy_emu") or 0)
        if current_cx <= 0 or current_cy <= 0:
            continue
        paragraph_index = int(row["paragraph_index"])
        caption = str(row.get("caption") or "")
        media_info = paragraph_media.get(paragraph_index, {})
        media_name = str(media_info.get("media_name") or "")
        if target_media_name:
            normalized = media_name.replace("\\", "/")
            if not (normalized == target_media_name or normalized.endswith("/" + target_media_name) or Path(normalized).name == target_media_name):
                continue
        if target_caption_contains and target_caption_contains not in caption:
            continue
        new_cx = 0
        new_cy = 0
        resize_reason = ""
        pixel_width = int(media_info.get("pixel_width") or 0)
        pixel_height = int(media_info.get("pixel_height") or 0)
        current_ppi_width = pixel_width / (current_cx / EMU_PER_CM / 2.54) if pixel_width else 0.0
        current_ppi_height = pixel_height / (current_cy / EMU_PER_CM / 2.54) if pixel_height else 0.0
        native_max_cx_candidates = []
        if min_native_ppi > 0 and pixel_width > 0:
            native_max_cx_candidates.append(int((pixel_width / min_native_ppi) * 2.54 * EMU_PER_CM))
        if min_native_ppi > 0 and pixel_height > 0 and current_cx > 0 and current_cy > 0:
            native_max_cy = int((pixel_height / min_native_ppi) * 2.54 * EMU_PER_CM)
            native_max_cx_candidates.append(int(native_max_cy * (current_cx / current_cy)))
        native_max_cx = min(native_max_cx_candidates) if native_max_cx_candidates else 0
        if (
            min_native_ppi > 0
            and pixel_width > 0
            and pixel_height > 0
            and min(current_ppi_width, current_ppi_height) < min_native_ppi
            and native_max_cx > 0
            and native_max_cx < current_cx
        ):
            new_cx = native_max_cx
            resize_reason = "below-min-native-ppi"
        elif max_safe_cy and current_cy > max_safe_cy:
            new_cy = max_safe_cy
            new_cx = max(1, int(round(current_cx * (new_cy / current_cy))))
            resize_reason = "exceeds-safe-page-height"
        elif current_cx < min_cx:
            if native_max_cx and native_max_cx < min_cx:
                continue
            new_cx = min_cx
            resize_reason = "below-min-text-width-ratio"
        elif max_cx and current_cx > max_cx:
            new_cx = min_cx
            resize_reason = "exceeds-text-width"
        elif min_cy and current_cy < min_cy:
            new_cx = current_cx
            new_cy = min_cy
            resize_reason = "below-min-readable-height"
        else:
            continue
        proposed_cy = new_cy or int(round(current_cy * (new_cx / current_cx)))
        if max_safe_cy and proposed_cy > max_safe_cy:
            continue
        plan[paragraph_index] = {
            "old_cx": current_cx,
            "old_cy": current_cy,
            "new_cx": new_cx,
            "new_cy": proposed_cy,
            "caption": caption,
            "media_name": media_name,
            "pixel_width": pixel_width,
            "pixel_height": pixel_height,
            "native_ppi_width_before": round(current_ppi_width, 1) if current_ppi_width else 0.0,
            "native_ppi_height_before": round(current_ppi_height, 1) if current_ppi_height else 0.0,
            "min_native_ppi": min_native_ppi,
            "resize_reason": resize_reason,
            "width_text_ratio_before": current_cx / text_width if text_width else 0.0,
            "width_text_ratio_after": new_cx / text_width if text_width else 0.0,
        }
    return plan


def paragraph_media_pixel_info(docx_path: Path) -> dict[int, dict[str, int | str]]:
    media_manifest = docx_image_relationship_manifest(docx_path)
    result: dict[int, dict[str, int | str]] = {}
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
            body = root.find(".//w:body", NS)
            if body is None:
                return {}
            paragraphs = body.findall(".//w:p", NS)
            for index, paragraph in enumerate(paragraphs, start=1):
                rids = paragraph_image_relationship_ids(paragraph)
                if not rids:
                    continue
                entry = media_manifest.get(f"word/document.xml#{rids[0]}", {})
                media_name = str(entry.get("media_name") or "")
                if not media_name:
                    continue
                try:
                    payload = zf.read(media_name)
                    with Image.open(io.BytesIO(payload)) as image:
                        result[index] = {
                            "rid": rids[0],
                            "media_name": media_name,
                            "media_sha256": str(entry.get("sha256") or ""),
                            "pixel_width": int(image.width),
                            "pixel_height": int(image.height),
                        }
                except Exception:
                    result[index] = {
                        "rid": rids[0],
                        "media_name": media_name,
                        "media_sha256": str(entry.get("sha256") or ""),
                        "pixel_width": 0,
                        "pixel_height": 0,
                    }
    except Exception:
        return {}
    return result


def _paragraph_spacing(paragraph: ET.Element) -> tuple[str, int]:
    spacing = paragraph.find("./w:pPr/w:spacing", NS)
    if spacing is None:
        return "", 0
    line_rule = spacing.attrib.get(f"{{{W_NS}}}lineRule", "")
    try:
        line_twips = int(spacing.attrib.get(f"{{{W_NS}}}line", "0") or "0")
    except ValueError:
        line_twips = 0
    return line_rule, line_twips


def _paragraph_layout_needs_repair(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return True
    jc = ppr.find("./w:jc", NS)
    if jc is None or jc.attrib.get(f"{{{W_NS}}}val", "") != "center":
        return True
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        return True
    for field in INDENT_FIELDS:
        raw_value = ind.attrib.get(f"{{{W_NS}}}{field}", "")
        try:
            value = int(raw_value or "0")
        except ValueError:
            return True
        if value != 0:
            return True
    if ppr.find("./w:numPr", NS) is not None or ppr.find("./w:outlineLvl", NS) is not None:
        return True
    if ppr.find("./w:keepNext", NS) is None:
        return True
    return False


def figure_holder_safety_plan(docx_path: Path) -> dict[int, dict[str, int | str]]:
    """Return image-holder paragraphs whose exact line height can clip images."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(".//w:body", NS)
    if body is None:
        return {}
    paragraphs = body.findall(".//w:p", NS)
    plan: dict[int, dict[str, int | str]] = {}
    for row in docx_body_figure_paragraphs(docx_path):
        if not row.get("has_drawing") or row.get("is_caption"):
            continue
        paragraph_index = int(row["paragraph_index"])
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        image_height_emu = int(row.get("max_extent_cy_emu") or 0)
        if image_height_emu <= 0:
            continue
        line_rule, line_twips = _paragraph_spacing(paragraphs[paragraph_index - 1])
        line_height_emu = line_twips * 635
        if line_rule.lower() == "exact" and (line_height_emu <= 0 or line_height_emu < int(image_height_emu * 0.98)):
            plan[paragraph_index] = {
                "old_line_rule": line_rule,
                "old_line_twips": line_twips,
                "image_height_emu": image_height_emu,
                "caption": str(row.get("caption") or ""),
                "repair_reason": "exact-line-spacing-clips-inline-image",
                "new_line_rule": "auto",
                "new_line_twips": 240,
            }
    return plan


def figure_holder_layout_plan(docx_path: Path) -> dict[int, dict[str, int | str | dict[str, str]]]:
    """Return image-holder paragraphs with unsafe alignment/indent/list residue."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(".//w:body", NS)
    if body is None:
        return {}
    paragraphs = body.findall(".//w:p", NS)
    plan: dict[int, dict[str, int | str | dict[str, str]]] = {}
    for row in docx_body_figure_paragraphs(docx_path):
        if not row.get("has_drawing") or row.get("is_caption"):
            continue
        paragraph_index = int(row["paragraph_index"])
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        paragraph = paragraphs[paragraph_index - 1]
        if not _paragraph_layout_needs_repair(paragraph):
            continue
        ppr = paragraph.find("./w:pPr", NS)
        jc = ppr.find("./w:jc", NS) if ppr is not None else None
        ind = ppr.find("./w:ind", NS) if ppr is not None else None
        plan[paragraph_index] = {
            "caption": str(row.get("caption") or ""),
            "repair_reason": "image-holder-alignment-indent-list-residue",
            "old_alignment": jc.attrib.get(f"{{{W_NS}}}val", "") if jc is not None else "",
            "old_indent": {field: ind.attrib.get(f"{{{W_NS}}}{field}", "") for field in INDENT_FIELDS} if ind is not None else {},
            "new_alignment": "center",
            "new_indent": {field: "0" for field in INDENT_FIELDS},
            "new_keep_next": "1",
        }
    return plan


def _emu_to_cm(value: int) -> str:
    return f"{value / EMU_PER_CM:.4f}cm"


def _set_vml_style_dimension(style: str, *, cx: int, cy: int) -> str:
    parts = [part for part in (style or "").split(";") if part.strip()]
    seen_width = False
    seen_height = False
    updated: list[str] = []
    for part in parts:
        if ":" not in part:
            updated.append(part)
            continue
        key, _raw_value = part.split(":", 1)
        lowered = key.strip().lower()
        if lowered == "width":
            updated.append(f"{key}:{_emu_to_cm(cx)}")
            seen_width = True
        elif lowered == "height":
            updated.append(f"{key}:{_emu_to_cm(cy)}")
            seen_height = True
        else:
            updated.append(part)
    if not seen_width:
        updated.append(f"width:{_emu_to_cm(cx)}")
    if not seen_height:
        updated.append(f"height:{_emu_to_cm(cy)}")
    return ";".join(updated)


def update_vml_extents(container: ET.Element, cx: int, cy: int) -> int:
    changes = 0
    for shape in container.iter():
        if not (shape.tag.endswith("}shape") or shape.tag.endswith("}rect")):
            continue
        if not any(descendant.tag.endswith("}imagedata") for descendant in shape.iter()):
            continue
        before = dict(shape.attrib)
        shape.attrib["style"] = _set_vml_style_dimension(shape.attrib.get("style", ""), cx=cx, cy=cy)
        if "width" in shape.attrib:
            shape.attrib["width"] = _emu_to_cm(cx)
        if "height" in shape.attrib:
            shape.attrib["height"] = _emu_to_cm(cy)
        if before != shape.attrib:
            changes += 1
    return changes


def _ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(f"{{{W_NS}}}pPr")
        paragraph.insert(0, ppr)
    return ppr


def _ensure_spacing(ppr: ET.Element) -> ET.Element:
    spacing = ppr.find("./w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(f"{{{W_NS}}}spacing")
        ppr.append(spacing)
    return spacing


def _ensure_jc(ppr: ET.Element) -> ET.Element:
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.Element(f"{{{W_NS}}}jc")
        ppr.append(jc)
    return jc


def _ensure_ind(ppr: ET.Element) -> ET.Element:
    ind = ppr.find("./w:ind", NS)
    if ind is None:
        ind = ET.Element(f"{{{W_NS}}}ind")
        ppr.append(ind)
    return ind


def _ensure_keep_next(ppr: ET.Element) -> ET.Element:
    keep_next = ppr.find("./w:keepNext", NS)
    if keep_next is None:
        keep_next = ET.Element(f"{{{W_NS}}}keepNext")
        ppr.append(keep_next)
    return keep_next


def _remove_child_if_present(parent: ET.Element, tag_name: str) -> bool:
    node = parent.find(f"./w:{tag_name}", NS)
    if node is None:
        return False
    parent.remove(node)
    return True


def resize_document_xml(
    document_xml: bytes,
    plan: dict[int, dict[str, int | float | str]],
    *,
    holder_safety_plan: dict[int, dict[str, int | str]] | None = None,
    holder_layout_plan: dict[int, dict[str, int | str | dict[str, str]]] | None = None,
) -> tuple[bytes, list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    root = ET.fromstring(document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    paragraphs = body.findall(".//w:p", NS)
    changes: list[dict[str, object]] = []
    holder_changes: list[dict[str, object]] = []
    holder_layout_changes: list[dict[str, object]] = []
    for paragraph_index, resize in sorted(plan.items()):
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        paragraph = paragraphs[paragraph_index - 1]
        new_cx = str(int(resize["new_cx"]))
        new_cy = str(int(resize["new_cy"]))
        extent_changes = 0
        for extent in paragraph.findall(".//wp:extent", NS):
            before = dict(extent.attrib)
            extent.set("cx", new_cx)
            extent.set("cy", new_cy)
            if before != extent.attrib:
                extent_changes += 1
        for extent in paragraph.findall(".//a:xfrm/a:ext", NS):
            before = dict(extent.attrib)
            extent.set("cx", new_cx)
            extent.set("cy", new_cy)
            if before != extent.attrib:
                extent_changes += 1
        extent_changes += update_vml_extents(paragraph, int(resize["new_cx"]), int(resize["new_cy"]))
        if extent_changes:
            record = dict(resize)
            record.update(
                {
                    "paragraph_index": paragraph_index,
                    "paragraph_text": paragraph_text(paragraph)[:80],
                    "extent_nodes_changed": extent_changes,
                }
            )
            changes.append(record)
    for paragraph_index, repair in sorted((holder_safety_plan or {}).items()):
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        paragraph = paragraphs[paragraph_index - 1]
        ppr = _ensure_ppr(paragraph)
        spacing = _ensure_spacing(ppr)
        before = dict(spacing.attrib)
        spacing.set(f"{{{W_NS}}}lineRule", str(repair.get("new_line_rule") or "auto"))
        spacing.set(f"{{{W_NS}}}line", str(repair.get("new_line_twips") or 240))
        if before != spacing.attrib:
            record = dict(repair)
            record.update(
                {
                    "paragraph_index": paragraph_index,
                    "paragraph_text": paragraph_text(paragraph)[:80],
                    "before_spacing": before,
                    "after_spacing": dict(spacing.attrib),
                }
            )
            holder_changes.append(record)
    for paragraph_index, repair in sorted((holder_layout_plan or {}).items()):
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        paragraph = paragraphs[paragraph_index - 1]
        ppr = _ensure_ppr(paragraph)
        before_xml = ET.tostring(ppr, encoding="unicode")
        jc = _ensure_jc(ppr)
        jc.set(f"{{{W_NS}}}val", "center")
        ind = _ensure_ind(ppr)
        for field in INDENT_FIELDS:
            ind.set(f"{{{W_NS}}}{field}", "0")
        spacing = _ensure_spacing(ppr)
        spacing.set(f"{{{W_NS}}}lineRule", "auto")
        spacing.set(f"{{{W_NS}}}line", "360")
        _ensure_keep_next(ppr)
        removed_numpr = _remove_child_if_present(ppr, "numPr")
        removed_outline = _remove_child_if_present(ppr, "outlineLvl")
        after_xml = ET.tostring(ppr, encoding="unicode")
        if before_xml != after_xml:
            record = dict(repair)
            record.update(
                {
                    "paragraph_index": paragraph_index,
                    "paragraph_text": paragraph_text(paragraph)[:80],
                    "removed_numPr": removed_numpr,
                    "removed_outlineLvl": removed_outline,
                }
            )
            holder_layout_changes.append(record)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changes, holder_changes, holder_layout_changes


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument("--min-text-width-ratio", type=float, default=0.95)
    parser.add_argument("--min-height-cm", type=float, default=4.0)
    parser.add_argument("--max-height-cm", type=float, default=16.0)
    parser.add_argument("--max-height-ratio", type=float, default=0.82)
    parser.add_argument("--min-native-ppi", type=float, default=0.0)
    parser.add_argument("--target-media-name", default="")
    parser.add_argument("--target-caption-contains", default="")
    args = parser.parse_args()

    input_docx = Path(args.input_docx)
    output_docx = Path(args.output_docx)
    report_path = Path(args.report_json)

    plan = figure_resize_plan(
        input_docx,
        min_text_width_ratio=args.min_text_width_ratio,
        min_height_cm=args.min_height_cm,
        max_height_cm=args.max_height_cm,
        max_height_ratio=args.max_height_ratio,
        min_native_ppi=args.min_native_ppi,
        target_media_name=args.target_media_name,
        target_caption_contains=args.target_caption_contains,
    )
    holder_plan = figure_holder_safety_plan(input_docx)
    holder_layout = figure_holder_layout_plan(input_docx)
    with zipfile.ZipFile(input_docx, "r") as zf:
        document_xml = zf.read("word/document.xml")
    new_xml, changes, holder_changes, holder_layout_changes = resize_document_xml(
        document_xml,
        plan,
        holder_safety_plan=holder_plan,
        holder_layout_plan=holder_layout,
    )
    write_docx_with_document_xml(input_docx, output_docx, new_xml)
    report = {
        "schema": "graduation-project-builder.docx-figure-extents-repair.v1",
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "min_text_width_ratio": args.min_text_width_ratio,
        "min_height_cm": args.min_height_cm,
        "max_height_cm": args.max_height_cm,
        "max_height_ratio": args.max_height_ratio,
        "min_native_ppi": args.min_native_ppi,
        "target_media_name": args.target_media_name,
        "target_caption_contains": args.target_caption_contains,
        "planned_resize_count": len(plan),
        "changed_resize_count": len(changes),
        "planned_holder_safety_count": len(holder_plan),
        "changed_holder_safety_count": len(holder_changes),
        "planned_holder_layout_count": len(holder_layout),
        "changed_holder_layout_count": len(holder_layout_changes),
        "changes": changes,
        "holder_safety_changes": holder_changes,
        "holder_layout_changes": holder_layout_changes,
        "verdict": "pass",
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": str(output_docx), "changed_resize_count": len(changes), "changed_holder_safety_count": len(holder_changes), "changed_holder_layout_count": len(holder_layout_changes)}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
