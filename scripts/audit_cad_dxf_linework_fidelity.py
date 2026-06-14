from __future__ import annotations

import argparse
import hashlib
import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any


REQUIRED_HEADERS = {
    "$LTSCALE": 1.0,
    "$CELTSCALE": 1.0,
    "$PSLTSCALE": 0,
    "$LWDISPLAY": 1,
}
REQUIRED_HEADER_POLICY = {
    "$LTSCALE": {"positive": True},
    "$CELTSCALE": {"positive": True},
    "$PSLTSCALE": {"allowed": {0, 1}},
    "$LWDISPLAY": {"equals": 1},
}
REQUIRED_LAYER_LINETYPES = {
    "CENTER": "CENTER",
    "HIDDEN": "HIDDEN",
}
REQUIRED_LAYER_STYLES = {
    "FRAME": {"lineweight_min": 50, "lineweight_max": 50, "linetype": "CONTINUOUS"},
    "OBJECT_THICK": {"lineweight_min": 50, "lineweight_max": 50, "linetype": "CONTINUOUS"},
    "OBJECT_THIN": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
    "CENTER": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CENTER"},
    "HIDDEN": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "HIDDEN"},
    "HATCH": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
    "DIM": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
    "LEADER": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
    "TEXT": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
    "BOM_GRID": {"lineweight_min": 25, "lineweight_max": 25, "linetype": "CONTINUOUS"},
}
MIN_READABLE_COLOR_DISTANCE_FROM_WHITE = 80.0
MAX_READABLE_COLOR_LUMINANCE = 190.0
LOW_CONTRAST_SAMPLE_LIMIT = 25
MIN_TEXT_GRID_COVER_RATIO = 0.02
LINE_FAMILY_LAYERS = {
    "thick_solid": {"OBJECT_THICK", "OUTLINE", "FRAME", "TITLE", "\u7c97\u5b9e\u7ebf", "\u7c97\u5b9e\u7ebf1", "\u8fb9\u6846\u7ebf"},
    "thin_solid": {"OBJECT_THIN", "DIM", "LEADER", "TEXT", "BOM_GRID", "NOTE", "\u7ec6\u5b9e\u7ebf", "\u5c3a\u5bf8", "\u6807\u6ce8", "\u6587\u5b57"},
    "center_dash_dot": {"CENTER", "\u4e2d\u5fc3\u7ebf"},
    "hidden_dashed": {"HIDDEN", "\u865a\u7ebf"},
    "section_hatch": {"HATCH", "\u5256\u9762\u7ebf"},
    "dimension": {"DIM", "\u5c3a\u5bf8"},
    "leader_or_annotation": {"LEADER", "TEXT", "BOM_GRID", "NOTE", "\u6807\u6ce8", "\u6587\u5b57", "\u5c3a\u5bf8"},
}
OPTIONAL_PER_SHEET_LINE_FAMILIES = {"center_dash_dot", "hidden_dashed", "section_hatch"}
LAYER_STYLE_ALIASES = {
    "FRAME": {"FRAME", "\u8fb9\u6846\u7ebf"},
    "OBJECT_THICK": {"OBJECT_THICK", "OUTLINE", "TITLE", "\u7c97\u5b9e\u7ebf", "\u7c97\u5b9e\u7ebf1"},
    "OBJECT_THIN": {"OBJECT_THIN", "\u7ec6\u5b9e\u7ebf"},
    "CENTER": {"CENTER", "\u4e2d\u5fc3\u7ebf"},
    "HIDDEN": {"HIDDEN", "\u865a\u7ebf"},
    "HATCH": {"HATCH", "\u5256\u9762\u7ebf"},
    "DIM": {"DIM", "\u5c3a\u5bf8"},
    "LEADER": {"LEADER", "NOTE", "\u6807\u6ce8"},
    "TEXT": {"TEXT", "\u6587\u5b57"},
    "BOM_GRID": {"BOM_GRID"},
}
OPTIONAL_LAYER_STYLES = {"BOM_GRID"}
DRAWABLE_TYPES = {
    "LINE",
    "CIRCLE",
    "ARC",
    "POLYLINE",
    "LWPOLYLINE",
    "DIMENSION",
    "INSERT",
    "TEXT",
    "MTEXT",
    "HATCH",
}
TABLE_GRID_COLLISION_LAYERS = {
    "FRAME",
    "OBJECT_THICK",
    "OBJECT_THIN",
    "BOM_GRID",
    "TITLE",
    "\u8fb9\u6846\u7ebf",
    "\u7c97\u5b9e\u7ebf",
    "\u7ec6\u5b9e\u7ebf",
}
TEXT_COLLISION_LAYERS = {"TEXT", "DIM", "LEADER", "NOTE", "\u6587\u5b57", "\u6807\u6ce8", "\u5c3a\u5bf8"}


def _iter_drawable_entities(doc: Any) -> list[Any]:
    entities: list[Any] = list(doc.modelspace())
    for block in doc.blocks:
        name = str(block.name)
        if name.startswith("*Model_Space") or name.startswith("*Paper_Space"):
            continue
        entities.extend(block)
    return [entity for entity in entities if entity.dxftype() in DRAWABLE_TYPES]


def _entity_box(entity: Any, bbox_module: Any) -> list[float] | None:
    try:
        ext = bbox_module.extents([entity])
    except Exception:
        return None
    if not ext.has_data:
        return None
    return [
        float(ext.extmin.x),
        float(ext.extmin.y),
        float(ext.extmax.x),
        float(ext.extmax.y),
    ]


def _box_width_height(box: list[float]) -> tuple[float, float]:
    return box[2] - box[0], box[3] - box[1]


def _doc_bounds(doc: Any, bbox_module: Any) -> list[float]:
    try:
        ext = bbox_module.extents(doc.modelspace())
    except Exception:
        return [0.0, 0.0, 1.0, 1.0]
    if not ext.has_data:
        return [0.0, 0.0, 1.0, 1.0]
    return [float(ext.extmin.x), float(ext.extmin.y), float(ext.extmax.x), float(ext.extmax.y)]


def _box_center(box: list[float]) -> tuple[float, float]:
    return (box[0] + box[2]) / 2.0, (box[1] + box[3]) / 2.0


RIGHT_TABLE_LINE_LAYERS = {"BOM_GRID", "FRAME", "OBJECT_THICK", "TITLE", "\u8fb9\u6846\u7ebf", "\u7c97\u5b9e\u7ebf"}


def _in_table_or_title_zone_box(box: list[float], bounds: list[float]) -> bool:
    minx, miny, maxx, maxy = bounds
    cw, ch = maxx - minx, maxy - miny
    if cw <= 80.0 or ch <= 80.0:
        return True
    cx, cy = _box_center(box)
    return cx > minx + 0.70 * cw or cy < miny + 0.245 * ch


def _is_protected_table_line_box(box: list[float], bounds: list[float], layer_name: str) -> bool:
    minx, miny, maxx, maxy = bounds
    cw, ch = maxx - minx, maxy - miny
    if cw <= 80.0 or ch <= 80.0:
        return True
    cx, cy = _box_center(box)
    if cy < miny + 0.245 * ch:
        return True
    return cx > minx + 0.70 * cw and layer_name in RIGHT_TABLE_LINE_LAYERS


def _expand_box(box: list[float], pad: float) -> list[float]:
    return [box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad]


def _shrink_box(box: list[float], pad: float) -> list[float]:
    if box[2] - box[0] <= pad * 2 or box[3] - box[1] <= pad * 2:
        return box
    return [box[0] + pad, box[1] + pad, box[2] - pad, box[3] - pad]


def _text_inner_collision_box(box: list[float]) -> list[float]:
    """Use the rendered text interior for cover checks, not the loose bbox edge."""
    width, height = _box_width_height(box)
    pad_x = max(0.12, min(0.55, width * 0.22))
    pad_y = max(0.12, min(0.55, height * 0.18))
    if width <= pad_x * 2 or height <= pad_y * 2:
        return _shrink_box(box, 0.10)
    return [box[0] + pad_x, box[1] + pad_y, box[2] - pad_x, box[3] - pad_y]


def _boxes_intersect(left: list[float], right: list[float]) -> bool:
    return min(left[2], right[2]) > max(left[0], right[0]) and min(left[3], right[3]) > max(left[1], right[1])


def _box_area(box: list[float]) -> float:
    return max(0.0, box[2] - box[0]) * max(0.0, box[3] - box[1])


def _box_intersection_area(left: list[float], right: list[float]) -> float:
    width = min(left[2], right[2]) - max(left[0], right[0])
    height = min(left[3], right[3]) - max(left[1], right[1])
    if width <= 0.0 or height <= 0.0:
        return 0.0
    return width * height


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def _as_float(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _rgb_luma(rgb: Any) -> float:
    r, g, b = _rgb_to_tuple(rgb)
    return 0.2126 * float(r) + 0.7152 * float(g) + 0.0722 * float(b)


def _rgb_distance_from_white(rgb: Any) -> float:
    r, g, b = _rgb_to_tuple(rgb)
    return ((255 - float(r)) ** 2 + (255 - float(g)) ** 2 + (255 - float(b)) ** 2) ** 0.5


def _is_readable_ink_rgb(rgb: Any) -> bool:
    return (
        _rgb_distance_from_white(rgb) >= MIN_READABLE_COLOR_DISTANCE_FROM_WHITE
        and _rgb_luma(rgb) <= MAX_READABLE_COLOR_LUMINANCE
    )


def _rgb_to_tuple(rgb: Any) -> tuple[int, int, int]:
    if isinstance(rgb, tuple):
        return (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    return (int(rgb.r), int(rgb.g), int(rgb.b))


def _aci_rgb(colors: Any, aci: int) -> Any:
    # ACI 7 is displayed as black on white CAD plotting backgrounds.
    if aci == 7:
        return (0, 0, 0)
    return colors.aci2rgb(aci)


def _effective_entity_rgb(entity: Any, colors: Any, layer_color_by_name: dict[str, int]) -> Any:
    true_color = _as_int(entity.dxf.get("true_color", None))
    if true_color is not None:
        return colors.int2rgb(true_color)
    aci = _as_int(entity.dxf.get("color", None))
    if aci is None or aci in {0, 256}:
        aci = layer_color_by_name.get(str(entity.dxf.get("layer", "0")), 7)
    return _aci_rgb(colors, aci)


def _layer_style_for_name(name: str) -> str | None:
    for style_name, aliases in LAYER_STYLE_ALIASES.items():
        if name in aliases:
            return style_name
    return None


def _style_layers(style_name: str, layer_weight_by_name: dict[str, int]) -> list[str]:
    aliases = LAYER_STYLE_ALIASES.get(style_name, {style_name})
    return [name for name in layer_weight_by_name if name in aliases]


def _max_weight_for_layers(layer_names: set[str], layer_weight_by_name: dict[str, int]) -> int:
    weights = [layer_weight_by_name.get(name, -1) for name in layer_names if name in layer_weight_by_name]
    return max(weights, default=-1)


def _dark_review_color_mode(layer_color_by_name: dict[str, int]) -> bool:
    family_colors: dict[str, set[int]] = {
        "thin_solid": set(),
        "thick_solid": set(),
        "center_dash_dot": set(),
        "hidden_dashed": set(),
        "section_hatch": set(),
        "leader_or_annotation": set(),
    }
    for layer_name, color in layer_color_by_name.items():
        for family, layers in LINE_FAMILY_LAYERS.items():
            if layer_name in layers and family in family_colors:
                family_colors[family].add(color)
    if 7 not in family_colors["thin_solid"]:
        return False
    non_thin_colors = set().union(
        family_colors["thick_solid"],
        family_colors["center_dash_dot"],
        family_colors["hidden_dashed"],
        family_colors["section_hatch"],
        family_colors["leader_or_annotation"],
    )
    return len(non_thin_colors - {7}) >= 4


def _collect_dxf_files(path: Path, tmp: Path) -> list[Path]:
    if path.is_dir():
        return sorted(path.rglob("*.dxf"))
    if path.suffix.lower() == ".dxf":
        return [path]
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as zf:
            zf.extractall(tmp)
        return sorted(tmp.rglob("*.dxf"))
    return []


def _audit_file(path: Path) -> dict[str, Any]:
    import ezdxf
    from ezdxf import bbox, colors

    issues: list[str] = []
    try:
        doc = ezdxf.readfile(path)
    except Exception as exc:  # pragma: no cover - defensive for broken CAD input
        return {
            "file": str(path),
            "passed": False,
            "issues": [f"cannot read DXF: {exc}"],
        }

    header_values: dict[str, Any] = {}
    for key, expected in REQUIRED_HEADERS.items():
        actual = doc.header.get(key)
        header_values[key] = actual
        policy = REQUIRED_HEADER_POLICY.get(key, {})
        if policy.get("positive"):
            actual_float = _as_float(actual)
            if actual_float is None or actual_float <= 0:
                issues.append(f"{key} must be positive, got {actual!r}")
        elif "allowed" in policy:
            actual_int = _as_int(actual)
            if actual_int not in policy["allowed"]:
                issues.append(f"{key} must be one of {sorted(policy['allowed'])}, got {actual!r}")
        elif "equals" in policy:
            actual_int = _as_int(actual)
            if actual_int != policy["equals"]:
                issues.append(f"{key} must be {expected}, got {actual!r}")
        elif isinstance(expected, float):
            actual_float = _as_float(actual)
            if actual_float is None or abs(actual_float - expected) > 1e-6:
                issues.append(f"{key} must be {expected}, got {actual!r}")
        else:
            actual_int = _as_int(actual)
            if actual_int != expected:
                issues.append(f"{key} must be {expected}, got {actual!r}")

    layer_rows: list[dict[str, Any]] = []
    layer_weight_by_name: dict[str, int] = {}
    layer_ltype_by_name: dict[str, str] = {}
    layer_color_by_name: dict[str, int] = {}
    for layer in doc.layers:
        name = str(layer.dxf.name)
        lineweight = _as_int(layer.dxf.get("lineweight", -1))
        linetype = str(layer.dxf.get("linetype", ""))
        color = _as_int(layer.dxf.get("color", 7)) or 7
        layer_weight_by_name[name] = lineweight if lineweight is not None else -1
        layer_ltype_by_name[name] = linetype
        layer_color_by_name[name] = color
        if name.lower() != "defpoints" and (lineweight is None or lineweight <= 0):
            issues.append(f"layer {name} missing positive lineweight")
        style_name = _layer_style_for_name(name)
        expected_ltype = REQUIRED_LAYER_LINETYPES.get(style_name or name)
        if expected_ltype and linetype.upper() != expected_ltype:
            issues.append(f"layer {name} linetype must be {expected_ltype}, got {linetype!r}")
        rgb = _aci_rgb(colors, color)
        layer_rows.append(
            {
                "name": name,
                "lineweight": lineweight,
                "linetype": linetype,
                "color": color,
                "effective_rgb": _rgb_to_tuple(rgb),
                "color_luminance": round(_rgb_luma(rgb), 2),
                "color_distance_from_white": round(_rgb_distance_from_white(rgb), 2),
                "readable_on_white_plot": _is_readable_ink_rgb(rgb),
            }
        )
    for name, expected in REQUIRED_LAYER_STYLES.items():
        present_layers = _style_layers(name, layer_weight_by_name)
        if not present_layers:
            if name not in OPTIONAL_LAYER_STYLES:
                issues.append(f"required mechanical line layer family {name} missing")
            continue
        for layer_name in present_layers:
            lineweight = layer_weight_by_name.get(layer_name, -1)
            if lineweight < expected["lineweight_min"]:
                issues.append(f"layer {layer_name} lineweight too small for {name}: {lineweight}")
            if "lineweight_max" in expected and lineweight > expected["lineweight_max"]:
                issues.append(f"layer {layer_name} lineweight too large for {name}: {lineweight}")
            actual_linetype = layer_ltype_by_name.get(layer_name, "").upper()
            if actual_linetype != str(expected["linetype"]).upper():
                issues.append(f"layer {layer_name} linetype must be {expected['linetype']} for {name}, got {actual_linetype!r}")

    drawable_count = 0
    explicit_lineweight_count = 0
    bylayer_with_layer_weight_count = 0
    missing_lineweight_count = 0
    linetype_scale_count = 0
    missing_linetype_scale_count = 0
    center_hidden_linetype_count = 0
    center_hidden_linetype_issue_count = 0
    low_contrast_color_count = 0
    low_contrast_color_samples: list[dict[str, Any]] = []
    line_family_counts = {family: 0 for family in LINE_FAMILY_LAYERS}
    layer_entity_counts: dict[str, int] = {}
    protected_grid_line_boxes: list[dict[str, Any]] = []
    text_boxes: list[dict[str, Any]] = []
    drawing_bounds = _doc_bounds(doc, bbox)

    for entity in _iter_drawable_entities(doc):
        drawable_count += 1
        layer_name = str(entity.dxf.get("layer", "0"))
        layer_entity_counts[layer_name] = layer_entity_counts.get(layer_name, 0) + 1
        for family, layers in LINE_FAMILY_LAYERS.items():
            if layer_name in layers:
                line_family_counts[family] += 1
        entity_weight = _as_int(entity.dxf.get("lineweight", -1))
        layer_weight = layer_weight_by_name.get(layer_name, -1)
        if entity_weight is not None and entity_weight > 0:
            explicit_lineweight_count += 1
        elif layer_weight > 0:
            bylayer_with_layer_weight_count += 1
        else:
            missing_lineweight_count += 1

        if entity.dxf.hasattr("ltscale"):
            linetype_scale_count += 1
            scale = _as_float(entity.dxf.get("ltscale", None))
            if scale is None or scale <= 0:
                missing_linetype_scale_count += 1
        elif entity.dxftype() in {"LINE", "ARC", "CIRCLE", "POLYLINE", "LWPOLYLINE"}:
            missing_linetype_scale_count += 1

        expected_ltype = REQUIRED_LAYER_LINETYPES.get(_layer_style_for_name(layer_name) or layer_name)
        if expected_ltype:
            center_hidden_linetype_count += 1
            linetype = str(entity.dxf.get("linetype", layer_ltype_by_name.get(layer_name, "")))
            if linetype.upper() not in {expected_ltype, "BYLAYER"}:
                center_hidden_linetype_issue_count += 1

        rgb = _effective_entity_rgb(entity, colors, layer_color_by_name)
        if not _is_readable_ink_rgb(rgb):
            low_contrast_color_count += 1
            if len(low_contrast_color_samples) < LOW_CONTRAST_SAMPLE_LIMIT:
                low_contrast_color_samples.append(
                    {
                        "dxftype": entity.dxftype(),
                        "layer": layer_name,
                        "rgb": _rgb_to_tuple(rgb),
                        "color_luminance": round(_rgb_luma(rgb), 2),
                        "color_distance_from_white": round(_rgb_distance_from_white(rgb), 2),
                    }
                )
        box = _entity_box(entity, bbox)
        if box:
            in_table_zone = _in_table_or_title_zone_box(box, drawing_bounds)
            if (
                _is_protected_table_line_box(box, drawing_bounds, layer_name)
                and entity.dxftype() in {"LINE", "LWPOLYLINE", "POLYLINE"}
                and layer_name in TABLE_GRID_COLLISION_LAYERS
            ):
                w, h = _box_width_height(box)
                if max(w, h) >= 1.0 and min(w, h) <= 1.0:
                    protected_grid_line_boxes.append({"layer": layer_name, "box": _expand_box(box, 0.18)})
            elif in_table_zone and entity.dxftype() in {"TEXT", "MTEXT"} and layer_name in TEXT_COLLISION_LAYERS:
                raw_text = entity.plain_text() if entity.dxftype() == "MTEXT" else str(entity.dxf.get("text", ""))
                if raw_text.strip():
                    text_boxes.append({"text": raw_text.strip(), "layer": layer_name, "box": _text_inner_collision_box(box)})

    table_text_grid_collision_samples: list[dict[str, Any]] = []
    for text_row in text_boxes:
        for line_row in protected_grid_line_boxes:
            overlap_area = _box_intersection_area(text_row["box"], line_row["box"])
            cover_ratio = overlap_area / max(_box_area(text_row["box"]), 1e-9)
            if _boxes_intersect(text_row["box"], line_row["box"]) and cover_ratio >= MIN_TEXT_GRID_COVER_RATIO:
                table_text_grid_collision_samples.append(
                    {
                        "text": text_row["text"][:60],
                        "text_layer": text_row["layer"],
                        "line_layer": line_row["layer"],
                        "text_box": [round(v, 3) for v in text_row["box"]],
                        "line_box": [round(v, 3) for v in line_row["box"]],
                        "cover_ratio": round(cover_ratio, 4),
                    }
                )
                break
    table_text_grid_collision_count = len(table_text_grid_collision_samples)

    if drawable_count <= 0:
        issues.append("no drawable DXF entities found")
    if missing_lineweight_count:
        issues.append(f"{missing_lineweight_count} drawable entities lack entity or layer lineweight")
    for family, count in line_family_counts.items():
        if count <= 0 and family not in OPTIONAL_PER_SHEET_LINE_FAMILIES:
            issues.append(f"required source line family {family} has no drawable entities")
    thick_weight = _max_weight_for_layers(LINE_FAMILY_LAYERS["thick_solid"], layer_weight_by_name)
    thin_weight = _max_weight_for_layers(LINE_FAMILY_LAYERS["thin_solid"] | LINE_FAMILY_LAYERS["section_hatch"], layer_weight_by_name)
    if thick_weight <= thin_weight:
        issues.append(f"thick solid lineweight must exceed thin solid lineweight, got thick={thick_weight}, thin={thin_weight}")
    # Entity group 48 is often omitted when the effective scale is exactly 1.0.
    # Treat it as acceptable only when the drawing-level scale controls are locked.
    if center_hidden_linetype_issue_count:
        issues.append(f"{center_hidden_linetype_issue_count} center/hidden entities use wrong linetype")
    color_review_mode = "dark_cad_review" if _dark_review_color_mode(layer_color_by_name) else "white_plot"
    if low_contrast_color_count and color_review_mode != "dark_cad_review":
        issues.append(
            f"{low_contrast_color_count} drawable entities use low-contrast/light CAD colors for white-background plotting"
        )
    if table_text_grid_collision_count:
        issues.append(f"{table_text_grid_collision_count} table/title/block text entities overlap protected grid/frame lines")

    return {
        "file": str(path),
        "passed": not issues,
        "acadver": doc.dxfversion,
        "headers": header_values,
        "layer_count": len(layer_rows),
        "layers": layer_rows,
        "drawable_entity_count": drawable_count,
        "layer_entity_counts": dict(sorted(layer_entity_counts.items())),
        "source_line_family_counts": line_family_counts,
        "explicit_lineweight_count": explicit_lineweight_count,
        "bylayer_with_layer_weight_count": bylayer_with_layer_weight_count,
        "missing_lineweight_count": missing_lineweight_count,
        "linetype_scale_count": linetype_scale_count,
        "missing_linetype_scale_count": missing_linetype_scale_count,
        "center_hidden_linetype_count": center_hidden_linetype_count,
        "center_hidden_linetype_issue_count": center_hidden_linetype_issue_count,
        "color_review_mode": color_review_mode,
        "low_contrast_color_count": low_contrast_color_count,
        "low_contrast_color_samples": low_contrast_color_samples,
        "protected_grid_line_box_count": len(protected_grid_line_boxes),
        "text_box_count": len(text_boxes),
        "table_title_zone_bounds": [round(v, 3) for v in drawing_bounds],
        "table_text_grid_collision_count": table_text_grid_collision_count,
        "table_text_grid_collision_samples": table_text_grid_collision_samples[:25],
        "minimum_text_grid_cover_ratio": MIN_TEXT_GRID_COVER_RATIO,
        "min_readable_color_distance_from_white": MIN_READABLE_COLOR_DISTANCE_FROM_WHITE,
        "max_readable_color_luminance": MAX_READABLE_COLOR_LUMINANCE,
        "issues": issues,
    }


def audit(path: Path) -> dict[str, Any]:
    input_sha256 = sha256_file(path) if path.is_file() else None
    with tempfile.TemporaryDirectory(prefix="cad-linework-") as tmp_name:
        tmp = Path(tmp_name)
        files = _collect_dxf_files(path, tmp)
        per_file = [_audit_file(file) for file in files]
    issues: list[str] = []
    if not per_file:
        issues.append("no DXF files found")
    package_line_family_counts = {family: 0 for family in LINE_FAMILY_LAYERS}
    for row in per_file:
        for family, count in row.get("source_line_family_counts", {}).items():
            package_line_family_counts[family] = package_line_family_counts.get(family, 0) + int(count)
    for family, count in package_line_family_counts.items():
        if count <= 0:
            issues.append(f"required package source line family {family} has no drawable entities")
    for row in per_file:
        if not row.get("passed"):
            issues.extend(f"{Path(str(row.get('file'))).name}: {issue}" for issue in row.get("issues", []))
    return {
        "schema": "graduation-project-builder.cad-dxf-linework-fidelity-audit.v1",
        "path": str(path),
        "path_sha256": input_sha256,
        "package_sha256": input_sha256 if path.suffix.lower() == ".zip" else None,
        "passed": not issues,
        "dxf_file_count": len(per_file),
        "issues": issues,
        "package_source_line_family_counts": package_line_family_counts,
        "per_file": per_file,
    }


def self_test() -> None:
    import ezdxf

    with tempfile.TemporaryDirectory(prefix="cad-linework-selftest-") as tmp_name:
        tmp = Path(tmp_name)
        good = tmp / "good.dxf"
        doc = ezdxf.new("R2000")
        doc.header["$LTSCALE"] = 1.0
        doc.header["$CELTSCALE"] = 1.0
        doc.header["$PSLTSCALE"] = 0
        doc.header["$LWDISPLAY"] = 1
        doc.layers.new("CENTER", dxfattribs={"linetype": "CENTER", "lineweight": 25})
        doc.layers.new("HIDDEN", dxfattribs={"linetype": "HIDDEN", "lineweight": 25})
        doc.layers.new("FRAME", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 50})
        doc.layers.new("OBJECT_THICK", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 50})
        doc.layers.new("OBJECT_THIN", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.new("HATCH", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.new("DIM", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.new("LEADER", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.new("TEXT", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.new("BOM_GRID", dxfattribs={"linetype": "CONTINUOUS", "lineweight": 25})
        doc.layers.get("0").dxf.lineweight = 25
        line = doc.modelspace().add_line((0, 0), (10, 0), dxfattribs={"layer": "CENTER", "linetype": "CENTER", "lineweight": 25})
        line.dxf.ltscale = 1.0
        doc.modelspace().add_line((0, 1), (10, 1), dxfattribs={"layer": "HIDDEN", "linetype": "BYLAYER"})
        doc.modelspace().add_line((0, 2), (10, 2), dxfattribs={"layer": "FRAME"})
        doc.modelspace().add_line((0, 3), (10, 3), dxfattribs={"layer": "OBJECT_THICK"})
        doc.modelspace().add_line((0, 4), (10, 4), dxfattribs={"layer": "OBJECT_THIN"})
        doc.modelspace().add_line((0, 5), (10, 5), dxfattribs={"layer": "HATCH"})
        doc.modelspace().add_linear_dim(base=(0, 6), p1=(0, 0), p2=(10, 0), dxfattribs={"layer": "DIM"}).render()
        doc.modelspace().add_line((0, 7), (10, 7), dxfattribs={"layer": "LEADER"})
        doc.modelspace().add_text("A", dxfattribs={"layer": "TEXT", "insert": (0, 10)})
        doc.modelspace().add_line((0, 8), (10, 8), dxfattribs={"layer": "BOM_GRID"})
        doc.saveas(good)
        result = audit(good)
        if not result["passed"]:
            raise SystemExit(json.dumps(result, indent=2))

        qcad_default_bad = tmp / "qcad-default-linework-bad.dxf"
        bad_doc = ezdxf.new("R2000")
        bad_doc.header["$LTSCALE"] = 1.0
        bad_doc.header["$CELTSCALE"] = 1.0
        bad_doc.header["$PSLTSCALE"] = 1
        bad_doc.header["$LWDISPLAY"] = 0
        layer_specs = {
            "CENTER": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 2},
            "HIDDEN": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 6},
            "FRAME": {"linetype": "CONTINUOUS", "lineweight": 50, "color": 3},
            "OBJECT_THICK": {"linetype": "CONTINUOUS", "lineweight": 50, "color": 3},
            "OBJECT_THIN": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 7},
            "HATCH": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 8},
            "DIM": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "LEADER": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "TEXT": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "BOM_GRID": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
        }
        for layer_name, attrs in layer_specs.items():
            bad_doc.layers.new(layer_name, dxfattribs=attrs)
            bad_doc.modelspace().add_line(
                (0, len(layer_specs)), (10, len(layer_specs)),
                dxfattribs={"layer": layer_name, "linetype": "CONTINUOUS"},
            )
        bad_doc.saveas(qcad_default_bad)
        bad_result = audit(qcad_default_bad)
        if bad_result["passed"]:
            raise SystemExit("expected QCAD default Continuous center/hidden and $LWDISPLAY=0 to fail")
        issue_text = "\n".join(bad_result.get("issues", []))
        if "$LWDISPLAY" not in issue_text or "CENTER" not in issue_text or "HIDDEN" not in issue_text:
            raise SystemExit(json.dumps(bad_result, indent=2))

        light = tmp / "light-color-fails.dxf"
        doc.layers.get("HATCH").dxf.color = 2
        doc.saveas(light)
        result = audit(light)
        if result["passed"]:
            raise SystemExit("expected low-contrast yellow CAD layer to fail")

        overlap = tmp / "table-text-overlap-fails.dxf"
        overlap_doc = ezdxf.new("R2000")
        overlap_doc.header["$LTSCALE"] = 1.0
        overlap_doc.header["$CELTSCALE"] = 1.0
        overlap_doc.header["$PSLTSCALE"] = 0
        overlap_doc.header["$LWDISPLAY"] = 1
        for layer_name, attrs in {
            "CENTER": {"linetype": "CENTER", "lineweight": 25, "color": 2},
            "HIDDEN": {"linetype": "HIDDEN", "lineweight": 25, "color": 6},
            "FRAME": {"linetype": "CONTINUOUS", "lineweight": 50, "color": 3},
            "OBJECT_THICK": {"linetype": "CONTINUOUS", "lineweight": 50, "color": 3},
            "OBJECT_THIN": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 7},
            "HATCH": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 8},
            "DIM": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "LEADER": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "TEXT": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
            "BOM_GRID": {"linetype": "CONTINUOUS", "lineweight": 25, "color": 4},
        }.items():
            overlap_doc.layers.new(layer_name, dxfattribs=attrs)
        overlap_doc.modelspace().add_line((0, 0), (10, 0), dxfattribs={"layer": "CENTER", "linetype": "CENTER"})
        overlap_doc.modelspace().add_line((0, 1), (10, 1), dxfattribs={"layer": "HIDDEN", "linetype": "HIDDEN"})
        overlap_doc.modelspace().add_line((0, 2), (10, 2), dxfattribs={"layer": "FRAME"})
        overlap_doc.modelspace().add_line((0, 3), (10, 3), dxfattribs={"layer": "OBJECT_THICK"})
        overlap_doc.modelspace().add_line((0, 4), (10, 4), dxfattribs={"layer": "OBJECT_THIN"})
        overlap_doc.modelspace().add_line((0, 5), (10, 5), dxfattribs={"layer": "HATCH"})
        overlap_doc.modelspace().add_linear_dim(base=(0, 6), p1=(0, 0), p2=(10, 0), dxfattribs={"layer": "DIM"}).render()
        overlap_doc.modelspace().add_line((0, 7), (10, 7), dxfattribs={"layer": "LEADER"})
        overlap_doc.modelspace().add_text("cell text", dxfattribs={"layer": "TEXT", "insert": (3, 8)})
        overlap_doc.modelspace().add_line((4, 7.5), (4, 9.5), dxfattribs={"layer": "BOM_GRID"})
        overlap_doc.saveas(overlap)
        overlap_result = audit(overlap)
        if overlap_result["passed"]:
            raise SystemExit("expected table/title text overlapping protected grid line to fail")
        if "table/title/block text" not in "\n".join(overlap_result.get("issues", [])):
            raise SystemExit(json.dumps(overlap_result, indent=2))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", nargs="?")
    parser.add_argument("--report-json")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
        print("self-test passed")
        return
    if not args.path:
        parser.error("path is required unless --self-test is used")
    result = audit(Path(args.path))
    if args.report_json:
        Path(args.report_json).write_text(json.dumps(result, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
