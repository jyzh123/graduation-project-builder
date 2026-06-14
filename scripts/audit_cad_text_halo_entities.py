from __future__ import annotations

import argparse
import hashlib
import json
import math
import sys
from pathlib import Path
from typing import Any

import ezdxf
from ezdxf import bbox


TEXT_TYPES = {"TEXT", "MTEXT", "ATTRIB", "ATTDEF"}
DANGEROUS_HALO_LAYERS = {
    "HATCH",
    "SECTION",
    "FLOW",
    "HEAT",
    "CENTER",
    "HIDDEN",
    "OBJECT_THICK",
    "OBJECT_THIN",
    "THICK_SOLID",
    "THIN_SOLID",
    "VISIBLE_THICK",
    "VISIBLE_THIN",
}
DIAGONAL_TEXT_COVER_LAYERS = {"HATCH", "SECTION", "FLOW", "HEAT"}
BAD_TEXT_TOKENS = ["\ufffd", "\u25a1", "??", "\u951f", "\u95c1", "\u5a75", "\u95bb", "\u4e71\u7801"]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def sha256_paths(paths: list[Path]) -> str:
    h = hashlib.sha256()
    for path in sorted(paths, key=lambda p: str(p).lower()):
        h.update(str(path.name).encode("utf-8", errors="surrogatepass"))
        h.update(b"\0")
        h.update(sha256_file(path).encode("ascii"))
        h.update(b"\0")
    return h.hexdigest().upper()


def as_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def box_tuple(ext: Any) -> tuple[float, float, float, float] | None:
    try:
        min_x, min_y, _ = ext.extmin
        max_x, max_y, _ = ext.extmax
    except Exception:
        return None
    if not all(math.isfinite(v) for v in (min_x, min_y, max_x, max_y)):
        return None
    if max_x <= min_x or max_y <= min_y:
        return None
    return (float(min_x), float(min_y), float(max_x), float(max_y))


def expand_box(box: tuple[float, float, float, float], margin: float) -> tuple[float, float, float, float]:
    return (box[0] - margin, box[1] - margin, box[2] + margin, box[3] + margin)


def intersect_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    return (x2 - x1) * (y2 - y1)


def point_in_rect(p: tuple[float, float], rect: tuple[float, float, float, float]) -> bool:
    return rect[0] <= p[0] <= rect[2] and rect[1] <= p[1] <= rect[3]


def ccw(a: tuple[float, float], b: tuple[float, float], c: tuple[float, float]) -> float:
    return (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])


def segments_intersect(
    a: tuple[float, float],
    b: tuple[float, float],
    c: tuple[float, float],
    d: tuple[float, float],
    eps: float = 1e-8,
) -> bool:
    ab_c = ccw(a, b, c)
    ab_d = ccw(a, b, d)
    cd_a = ccw(c, d, a)
    cd_b = ccw(c, d, b)
    if (ab_c > eps and ab_d < -eps or ab_c < -eps and ab_d > eps) and (
        cd_a > eps and cd_b < -eps or cd_a < -eps and cd_b > eps
    ):
        return True
    return False


def segment_intersects_rect(
    a: tuple[float, float],
    b: tuple[float, float],
    rect: tuple[float, float, float, float],
) -> bool:
    if point_in_rect(a, rect) or point_in_rect(b, rect):
        return True
    x1, y1, x2, y2 = rect
    edges = [((x1, y1), (x2, y1)), ((x2, y1), (x2, y2)), ((x2, y2), (x1, y2)), ((x1, y2), (x1, y1))]
    return any(segments_intersect(a, b, c, d) for c, d in edges)


def box_area(a: tuple[float, float, float, float]) -> float:
    return max(0.0, a[2] - a[0]) * max(0.0, a[3] - a[1])


def entity_bbox(entity: Any) -> tuple[float, float, float, float] | None:
    try:
        return box_tuple(bbox.extents([entity], fast=False))
    except Exception:
        try:
            return box_tuple(bbox.extents([entity], fast=True))
        except Exception:
            return None


def text_value(entity: Any) -> str:
    try:
        if entity.dxftype() == "MTEXT":
            return str(entity.text)
        return str(entity.dxf.get("text", ""))
    except Exception:
        return ""


def lineweight_margin(entity: Any) -> float:
    try:
        lw = as_float(entity.dxf.get("lineweight", 25), 25.0)
    except Exception:
        lw = 25.0
    # DXF lineweight is stored in hundredths of mm. Add a floor so zero-width
    # line bboxes still catch visible crossings.
    return max(0.18, lw / 100.0 / 2.0)


def entity_segments(entity: Any) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    etype = entity.dxftype()
    try:
        if etype == "LINE":
            start = entity.dxf.start
            end = entity.dxf.end
            return [((float(start[0]), float(start[1])), (float(end[0]), float(end[1])))]
        if etype == "LWPOLYLINE":
            pts = [(float(p[0]), float(p[1])) for p in entity.get_points("xy")]
            segs = list(zip(pts, pts[1:]))
            try:
                if entity.closed and len(pts) > 1:
                    segs.append((pts[-1], pts[0]))
            except Exception:
                pass
            return segs
        if etype == "POLYLINE":
            pts = [(float(v.dxf.location[0]), float(v.dxf.location[1])) for v in entity.vertices]
            segs = list(zip(pts, pts[1:]))
            try:
                if entity.is_closed and len(pts) > 1:
                    segs.append((pts[-1], pts[0]))
            except Exception:
                pass
            return segs
    except Exception:
        return []
    return []


def collect_geometry_entities(msp: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for entity in msp:
        etype = entity.dxftype()
        if etype in TEXT_TYPES:
            continue
        if etype in {"VIEWPORT", "DIMENSION", "WIPEOUT", "CIRCLE", "ARC", "ELLIPSE"}:
            continue
        margin = lineweight_margin(entity)
        layer = str(entity.dxf.get("layer", ""))
        segments = entity_segments(entity)
        if segments:
            for start, end in segments:
                rows.append(
                    {
                        "type": etype,
                        "layer": layer,
                        "segment": (start, end),
                        "margin": margin,
                    }
                )
            continue
        box = entity_bbox(entity)
        if box is not None:
            rows.append(
                {
                    "type": etype,
                    "layer": layer,
                    "bbox": expand_box(box, margin),
                    "margin": margin,
                }
            )
    return rows


def collect_text_mask_boxes(msp: Any) -> list[tuple[float, float, float, float]]:
    boxes: list[tuple[float, float, float, float]] = []
    for entity in msp:
        if entity.dxftype() != "WIPEOUT":
            continue
        layer = str(entity.dxf.get("layer", "")).upper()
        if layer and layer != "TEXT_MASK":
            continue
        box = entity_bbox(entity)
        if box is not None:
            boxes.append(box)
    return boxes


def box_contains(container: tuple[float, float, float, float], inner: tuple[float, float, float, float], tolerance: float = 0.05) -> bool:
    return (
        container[0] <= inner[0] + tolerance
        and container[1] <= inner[1] + tolerance
        and container[2] >= inner[2] - tolerance
        and container[3] >= inner[3] - tolerance
    )


def has_text_mask(
    text_box: tuple[float, float, float, float],
    masks: list[tuple[float, float, float, float]],
    margin: float,
) -> bool:
    required = expand_box(text_box, margin)
    return any(box_contains(mask, required, tolerance=0.15) for mask in masks)


def collect_text_entities(msp: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for entity in msp:
        if entity.dxftype() not in TEXT_TYPES:
            continue
        value = text_value(entity).strip()
        if not value:
            continue
        box = entity_bbox(entity)
        if box is None:
            continue
        height = as_float(entity.dxf.get("height", 0.0), 0.0)
        rows.append(
            {
                "text": value,
                "type": entity.dxftype(),
                "layer": str(entity.dxf.get("layer", "")),
                "height": height,
                "bbox": box,
            }
        )
    return rows


def audit_dxf(path: Path, halo_mm: float, sample_limit: int) -> dict[str, Any]:
    doc = ezdxf.readfile(path)
    msp = doc.modelspace()
    texts = collect_text_entities(msp)
    geometry = collect_geometry_entities(msp)
    masks = collect_text_mask_boxes(msp)
    overlap_count = 0
    halo_count = 0
    diagonal_cover_count = 0
    mojibake_count = 0
    small_text_count = 0
    mask_protected_text_count = 0
    samples: list[dict[str, Any]] = []

    for text_row in texts:
        value = str(text_row["text"])
        if any(token in value for token in BAD_TEXT_TOKENS):
            mojibake_count += 1
            if len(samples) < sample_limit:
                samples.append({"kind": "mojibake", "text": value, "text_layer": text_row["layer"]})
        if as_float(text_row.get("height"), 0.0) and as_float(text_row.get("height"), 0.0) < 2.0:
            small_text_count += 1
        text_box = tuple(text_row["bbox"])  # type: ignore[arg-type]
        text_area = max(box_area(text_box), 0.001)
        halo_box = expand_box(text_box, halo_mm)
        mask_protected = has_text_mask(text_box, masks, margin=min(0.55, halo_mm))
        if mask_protected:
            mask_protected_text_count += 1
        for geo in geometry:
            if "segment" in geo:
                start, end = geo["segment"]
                core_hit = segment_intersects_rect(start, end, expand_box(text_box, as_float(geo.get("margin"), 0.18)))
                halo_hit = segment_intersects_rect(start, end, expand_box(halo_box, as_float(geo.get("margin"), 0.18)))
                area = text_area if core_hit else 0.0
                halo_area = text_area if halo_hit else 0.0
            else:
                geo_box = tuple(geo["bbox"])  # type: ignore[arg-type]
                area = intersect_area(text_box, geo_box)
                halo_area = intersect_area(halo_box, geo_box)
            if area > max(0.03, text_area * 0.004) and not mask_protected:
                overlap_count += 1
                if str(geo["layer"]).upper() in DIAGONAL_TEXT_COVER_LAYERS:
                    diagonal_cover_count += 1
                if len(samples) < sample_limit:
                    samples.append(
                        {
                            "kind": "text_geometry_overlap",
                            "text": value,
                            "text_layer": text_row["layer"],
                            "geometry_type": geo["type"],
                            "geometry_layer": geo["layer"],
                            "intersection_area_mm2": round(area, 4),
                        }
                    )
                continue
            if str(geo["layer"]).upper() in DANGEROUS_HALO_LAYERS:
                if halo_area > max(0.03, text_area * 0.002) and not mask_protected:
                    halo_count += 1
                    if str(geo["layer"]).upper() in DIAGONAL_TEXT_COVER_LAYERS:
                        diagonal_cover_count += 1
                    if len(samples) < sample_limit:
                        samples.append(
                            {
                                "kind": "text_exclusion_halo_violation",
                                "text": value,
                                "text_layer": text_row["layer"],
                                "geometry_type": geo["type"],
                                "geometry_layer": geo["layer"],
                                "intersection_area_mm2": round(halo_area, 4),
                            }
                        )

    passed = overlap_count == 0 and halo_count == 0 and diagonal_cover_count == 0 and mojibake_count == 0
    return {
        "dxf": str(path),
        "sha256": sha256_file(path),
        "passed": passed,
        "source_text_entity_count": len(texts),
        "checked_geometry_entity_count": len(geometry),
        "text_mask_entity_count": len(masks),
        "mask_protected_text_count": mask_protected_text_count,
        "text_geometry_overlap_count": overlap_count,
        "text_exclusion_halo_violation_count": halo_count,
        "diagonal_hatch_section_flow_text_cover_count": diagonal_cover_count,
        "mojibake_or_missing_glyph_count": mojibake_count,
        "small_text_entity_count": small_text_count,
        "samples": samples,
    }


def resolve_dxf_paths(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() == ".dxf" else []
    return sorted(input_path.rglob("*.dxf"), key=lambda p: str(p).lower())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit CAD TEXT/MTEXT safety halos from source DXF entities.")
    parser.add_argument("--dxf-root", required=True, help="DXF file or directory containing DXF files.")
    parser.add_argument("--out-json", required=True, help="Audit JSON output path.")
    parser.add_argument("--halo-mm", type=float, default=0.7, help="Text exclusion halo around source text boxes.")
    parser.add_argument("--sample-limit", type=int, default=50)
    args = parser.parse_args(argv)

    dxf_root = Path(args.dxf_root)
    out_json = Path(args.out_json)
    paths = resolve_dxf_paths(dxf_root)
    per_sheet = []
    errors = []
    for path in paths:
        try:
            per_sheet.append(audit_dxf(path, args.halo_mm, args.sample_limit))
        except Exception as exc:  # pragma: no cover - defensive for malformed third-party DXF
            errors.append({"dxf": str(path), "error": repr(exc)})

    totals = {
        "source_text_entity_count": sum(int(row.get("source_text_entity_count") or 0) for row in per_sheet),
        "checked_geometry_entity_count": sum(int(row.get("checked_geometry_entity_count") or 0) for row in per_sheet),
        "text_mask_entity_count": sum(int(row.get("text_mask_entity_count") or 0) for row in per_sheet),
        "mask_protected_text_count": sum(int(row.get("mask_protected_text_count") or 0) for row in per_sheet),
        "text_geometry_overlap_count": sum(int(row.get("text_geometry_overlap_count") or 0) for row in per_sheet),
        "text_exclusion_halo_violation_count": sum(int(row.get("text_exclusion_halo_violation_count") or 0) for row in per_sheet),
        "diagonal_hatch_section_flow_text_cover_count": sum(
            int(row.get("diagonal_hatch_section_flow_text_cover_count") or 0) for row in per_sheet
        ),
        "mojibake_or_missing_glyph_count": sum(int(row.get("mojibake_or_missing_glyph_count") or 0) for row in per_sheet),
        "small_text_entity_count": sum(int(row.get("small_text_entity_count") or 0) for row in per_sheet),
    }
    passed = (
        bool(paths)
        and not errors
        and totals["source_text_entity_count"] > 0
        and totals["text_geometry_overlap_count"] == 0
        and totals["text_exclusion_halo_violation_count"] == 0
        and totals["diagonal_hatch_section_flow_text_cover_count"] == 0
        and totals["mojibake_or_missing_glyph_count"] == 0
    )
    payload = {
        "schema": "graduation-project-builder.cad-text-halo-source-audit.v1",
        "audit_method": "source-dxf-text-bbox-geometry-halo-machine",
        "passed": passed,
        "dxf_root": str(dxf_root),
        "dxf_root_sha256": sha256_paths(paths) if paths else "",
        "sheet_count": len(paths),
        "halo_mm": args.halo_mm,
        **totals,
        "errors": errors,
        "samples": [sample for row in per_sheet for sample in row.get("samples", [])][: args.sample_limit],
        "per_sheet": per_sheet,
    }
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=True, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
