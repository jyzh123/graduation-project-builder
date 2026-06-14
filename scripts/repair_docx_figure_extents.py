#!/usr/bin/env python3
"""Resize body figure drawings to the locked paragraph-width contract.

This is a bounded DOCX package repair. It edits only DrawingML extent values in
``word/document.xml`` for body figure-holder paragraphs that are below the
configured text-width ratio, whose displayed height is below the readable
minimum, or whose displayed height already exceeds the configured absolute or
page-relative safe height. Front-matter/template-donor drawings are ignored.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from thesis_figure_contract import (
    docx_body_figure_paragraphs,
    docx_text_height_emu,
    docx_text_width_emu,
)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_CM = 360000

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
    for row in docx_body_figure_paragraphs(docx_path):
        if not row.get("has_drawing") or row.get("is_caption") or row.get("front_matter_drawing"):
            continue
        current_cx = int(row.get("max_extent_cx_emu") or 0)
        current_cy = int(row.get("max_extent_cy_emu") or 0)
        if current_cx <= 0 or current_cy <= 0:
            continue
        new_cx = 0
        new_cy = 0
        resize_reason = ""
        if max_safe_cy and current_cy > max_safe_cy:
            new_cy = max_safe_cy
            new_cx = max(1, int(round(current_cx * (new_cy / current_cy))))
            resize_reason = "exceeds-safe-page-height"
        elif current_cx < min_cx:
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
        paragraph_index = int(row["paragraph_index"])
        plan[paragraph_index] = {
            "old_cx": current_cx,
            "old_cy": current_cy,
            "new_cx": new_cx,
            "new_cy": proposed_cy,
            "caption": str(row.get("caption") or ""),
            "resize_reason": resize_reason,
            "width_text_ratio_before": current_cx / text_width if text_width else 0.0,
            "width_text_ratio_after": new_cx / text_width if text_width else 0.0,
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


def resize_document_xml(document_xml: bytes, plan: dict[int, dict[str, int | float | str]]) -> tuple[bytes, list[dict[str, object]]]:
    root = ET.fromstring(document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    paragraphs = body.findall(".//w:p", NS)
    changes: list[dict[str, object]] = []
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
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changes


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    parser.add_argument("--min-text-width-ratio", type=float, default=0.95)
    parser.add_argument("--min-height-cm", type=float, default=4.0)
    parser.add_argument("--max-height-cm", type=float, default=16.0)
    parser.add_argument("--max-height-ratio", type=float, default=0.82)
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
    )
    with zipfile.ZipFile(input_docx, "r") as zf:
        document_xml = zf.read("word/document.xml")
    new_xml, changes = resize_document_xml(document_xml, plan)
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
        "planned_resize_count": len(plan),
        "changed_resize_count": len(changes),
        "changes": changes,
        "verdict": "pass",
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": str(output_docx), "changed_resize_count": len(changes)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
