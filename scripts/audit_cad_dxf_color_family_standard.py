from __future__ import annotations

import argparse
import hashlib
import json
import math
import tempfile
import zipfile
from pathlib import Path
from typing import Any


WHITE_ACI = {7}
BYLAYER_ACI = {0, 256}
EXPECTED_FAMILIES = {
    "thin_solid": {
        "layers": {"\u7ec6\u5b9e\u7ebf", "OBJECT_THIN", "THIN_SOLID"},
        "aci": 7,
        "rgb": (255, 255, 255),
        "must_be_white": True,
    },
    "thick_solid_or_frame": {
        "layers": {"\u7c97\u5b9e\u7ebf", "\u7c97\u5b9e\u7ebf1", "\u8fb9\u6846\u7ebf", "OBJECT_THICK", "FRAME", "OUTLINE", "TITLE"},
        "aci": 3,
        "rgb": (0, 255, 0),
    },
    "dimension_annotation_text": {
        "layers": {"\u6807\u6ce8", "\u6587\u5b57", "\u5c3a\u5bf8", "DIM", "LEADER", "TEXT", "BOM_GRID", "NOTE"},
        "aci": 5,
        "rgb": (0, 0, 255),
    },
    "center_dash_dot": {
        "layers": {"\u4e2d\u5fc3\u7ebf", "CENTER"},
        "aci": 1,
        "rgb": (255, 0, 0),
    },
    "hidden_dashed": {
        "layers": {"\u865a\u7ebf", "HIDDEN"},
        "aci": 6,
        "rgb": (255, 0, 255),
    },
    "section_hatch": {
        "layers": {"\u5256\u9762\u7ebf", "HATCH"},
        "aci": 8,
        "rgb": (128, 128, 128),
    },
}
IGNORED_LAYER_NAMES = {"Defpoints"}
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
    "SOLID",
    "TRACE",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


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


def _as_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _rgb_distance(left: tuple[int, int, int], right: tuple[int, int, int]) -> float:
    return math.sqrt(sum((left[i] - right[i]) ** 2 for i in range(3)))


def _rgb_tuple(rgb: Any) -> tuple[int, int, int]:
    if isinstance(rgb, tuple):
        return (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    return (int(rgb.r), int(rgb.g), int(rgb.b))


def _layer_family(layer_name: str) -> str | None:
    for family, spec in EXPECTED_FAMILIES.items():
        if layer_name in spec["layers"]:
            return family
    return None


def _aci_or_true_rgb(entity_or_layer: Any, colors: Any) -> tuple[int | None, tuple[int, int, int] | None]:
    true_color = _as_int(entity_or_layer.dxf.get("true_color", None))
    if true_color is not None:
        rgb = colors.int2rgb(true_color)
        return None, _rgb_tuple(rgb)
    aci = _as_int(entity_or_layer.dxf.get("color", None))
    if aci is None:
        return None, None
    if aci in BYLAYER_ACI:
        return aci, None
    if aci == 7:
        return aci, (255, 255, 255)
    rgb = colors.aci2rgb(aci)
    return aci, _rgb_tuple(rgb)


def _is_white(aci: int | None, rgb: tuple[int, int, int] | None) -> bool:
    if aci in WHITE_ACI:
        return True
    return bool(rgb and _rgb_distance(rgb, (255, 255, 255)) <= 15.0)


def _matches_expected(aci: int | None, rgb: tuple[int, int, int] | None, spec: dict[str, Any]) -> bool:
    if spec.get("must_be_white"):
        return _is_white(aci, rgb)
    expected_aci = int(spec["aci"])
    expected_rgb = tuple(int(v) for v in spec["rgb"])
    return aci == expected_aci or bool(rgb and _rgb_distance(rgb, expected_rgb) <= 35.0)


def _iter_drawable_entities(doc: Any) -> list[Any]:
    entities: list[Any] = list(doc.modelspace())
    for block in doc.blocks:
        name = str(block.name)
        if name.startswith("*Model_Space") or name.startswith("*Paper_Space"):
            continue
        entities.extend(block)
    return [entity for entity in entities if entity.dxftype() in DRAWABLE_TYPES]


def _audit_file(path: Path) -> dict[str, Any]:
    import ezdxf
    from ezdxf import colors

    issues: list[str] = []
    try:
        doc = ezdxf.readfile(path)
    except Exception as exc:  # pragma: no cover
        return {"file": str(path), "passed": False, "issues": [f"cannot read DXF: {exc}"]}

    layer_rows: list[dict[str, Any]] = []
    layer_family_by_name: dict[str, str | None] = {}
    layer_color_by_name: dict[str, int | None] = {}
    family_layer_count = {family: 0 for family in EXPECTED_FAMILIES}

    for layer in doc.layers:
        name = str(layer.dxf.name)
        family = _layer_family(name)
        layer_family_by_name[name] = family
        aci, rgb = _aci_or_true_rgb(layer, colors)
        layer_color_by_name[name] = aci
        ignored = name in IGNORED_LAYER_NAMES
        if family:
            family_layer_count[family] += 1
            spec = EXPECTED_FAMILIES[family]
            if not _matches_expected(aci, rgb, spec):
                issues.append(f"layer {name} must use {family} color ACI {spec['aci']} RGB {spec['rgb']}, got ACI {aci} RGB {rgb}")
        elif not ignored and name != "0" and _is_white(aci, rgb):
            issues.append(f"non-thin layer {name} is white but only thin solid may remain white")
        elif name == "0" and _is_white(aci, rgb):
            issues.append("layer 0 is white; set it to a non-white support color so stray BYLAYER entities do not plot white")

        if family != "thin_solid" and not ignored and _is_white(aci, rgb):
            issues.append(f"layer {name} is white; only thin solid may be white")
        layer_rows.append(
            {
                "name": name,
                "family": family,
                "aci": aci,
                "rgb": rgb,
                "linetype": str(layer.dxf.get("linetype", "")),
                "lineweight": _as_int(layer.dxf.get("lineweight", None)),
            }
        )

    for family, count in family_layer_count.items():
        if count <= 0:
            issues.append(f"required color family layer missing: {family}")

    entity_rows: list[dict[str, Any]] = []
    white_override_count = 0
    non_bylayer_override_count = 0
    for entity in _iter_drawable_entities(doc):
        layer_name = str(entity.dxf.get("layer", "0"))
        family = layer_family_by_name.get(layer_name)
        entity_aci = _as_int(entity.dxf.get("color", None))
        true_color = _as_int(entity.dxf.get("true_color", None))
        if true_color is not None or (entity_aci is not None and entity_aci not in BYLAYER_ACI):
            non_bylayer_override_count += 1
            aci, rgb = _aci_or_true_rgb(entity, colors)
            if family == "thin_solid":
                if not _is_white(aci, rgb):
                    issues.append(f"thin solid entity on layer {layer_name} has non-white override ACI {aci} RGB {rgb}")
            else:
                if _is_white(aci, rgb):
                    white_override_count += 1
                    issues.append(f"non-thin entity on layer {layer_name} has white override ACI {aci} RGB {rgb}")
                elif family and not _matches_expected(aci, rgb, EXPECTED_FAMILIES[family]):
                    issues.append(f"entity on layer {layer_name} overrides expected {family} color with ACI {aci} RGB {rgb}")
            if len(entity_rows) < 50:
                entity_rows.append(
                    {
                        "dxftype": entity.dxftype(),
                        "layer": layer_name,
                        "family": family,
                        "entity_aci": aci,
                        "entity_rgb": rgb,
                    }
                )

    return {
        "file": str(path),
        "passed": not issues,
        "issues": issues,
        "layer_count": len(layer_rows),
        "layers": layer_rows,
        "family_layer_count": family_layer_count,
        "non_bylayer_entity_color_override_count": non_bylayer_override_count,
        "non_thin_white_entity_override_count": white_override_count,
        "entity_override_samples": entity_rows,
    }


def audit(path: Path) -> dict[str, Any]:
    input_sha256 = sha256_file(path) if path.is_file() else None
    with tempfile.TemporaryDirectory(prefix="cad-color-family-") as tmp_name:
        tmp = Path(tmp_name)
        files = _collect_dxf_files(path, tmp)
        per_file = [_audit_file(file) for file in files]
    issues: list[str] = []
    if not per_file:
        issues.append("no DXF files found")
    for row in per_file:
        if not row.get("passed"):
            issues.extend(f"{Path(str(row.get('file'))).name}: {issue}" for issue in row.get("issues", []))
    family_palette = {
        family: {"aci": spec["aci"], "rgb": spec["rgb"], "layers": sorted(spec["layers"])}
        for family, spec in EXPECTED_FAMILIES.items()
    }
    return {
        "schema": "graduation-project-builder.cad-dxf-color-family-standard.v1",
        "path": str(path),
        "path_sha256": input_sha256,
        "package_sha256": input_sha256 if path.suffix.lower() == ".zip" else None,
        "passed": not issues,
        "dxf_file_count": len(per_file),
        "rule_summary": "thin_solid_white_only; all other line families non-white and color-distinguishable",
        "family_palette": family_palette,
        "issues": issues,
        "per_file": per_file,
    }


def self_test() -> None:
    import ezdxf
    from ezdxf import colors

    with tempfile.TemporaryDirectory(prefix="cad-color-family-selftest-") as tmp_name:
        tmp = Path(tmp_name)
        good = tmp / "good.dxf"
        doc = ezdxf.new("R2000")
        for family, spec in EXPECTED_FAMILIES.items():
            name = sorted(spec["layers"])[0]
            doc.layers.new(name, dxfattribs={"color": spec["aci"], "lineweight": 25})
            doc.modelspace().add_line((0, len(doc.layers)), (10, len(doc.layers)), dxfattribs={"layer": name, "color": 256})
        doc.layers.get("0").dxf.color = 4
        doc.layers.get("Defpoints").dxf.color = 4
        doc.saveas(good)
        good_result = audit(good)
        if not good_result["passed"]:
            raise SystemExit(json.dumps(good_result, ensure_ascii=False, indent=2))

        bad = tmp / "bad.dxf"
        doc.layers.get(sorted(EXPECTED_FAMILIES["thick_solid_or_frame"]["layers"])[0]).dxf.color = 7
        doc.saveas(bad)
        bad_result = audit(bad)
        if bad_result["passed"]:
            raise SystemExit("expected white non-thin layer to fail")

        qcad_truecolor_bad = tmp / "qcad-truecolor-override-bad.dxf"
        bad_doc = ezdxf.new("R2010")
        bad_layer_specs = {
            "OBJECT_THIN": {"color": 7},
            "THIN_SOLID": {"color": 7},
            "FRAME": {"color": 3, "true_color": colors.rgb2int((0, 255, 80))},
            "OBJECT_THICK": {"color": 3, "true_color": colors.rgb2int((0, 255, 80))},
            "DIM": {"color": 5},
            "CENTER": {"color": 1},
            "HIDDEN": {"color": 6},
            "HATCH": {"color": 8, "true_color": colors.rgb2int((170, 170, 170))},
        }
        for idx, (layer_name, attrs) in enumerate(bad_layer_specs.items()):
            bad_doc.layers.new(layer_name, dxfattribs=attrs)
            entity_attribs = {"layer": layer_name}
            if layer_name in {"FRAME", "OBJECT_THICK"}:
                entity_attribs["true_color"] = colors.rgb2int((0, 255, 80))
            if layer_name == "HATCH":
                entity_attribs["true_color"] = colors.rgb2int((170, 170, 170))
            bad_doc.modelspace().add_line((0, idx), (10, idx), dxfattribs=entity_attribs)
        bad_doc.layers.get("0").dxf.color = 7
        bad_doc.saveas(qcad_truecolor_bad)
        qcad_bad_result = audit(qcad_truecolor_bad)
        if qcad_bad_result["passed"]:
            raise SystemExit("expected QCAD near-color true_color overrides and white layer 0 to fail")
        issue_text = "\n".join(qcad_bad_result.get("issues", []))
        if "true_color" in issue_text:
            # The report describes resolved RGB values, not the storage field.
            raise SystemExit(json.dumps(qcad_bad_result, ensure_ascii=False, indent=2))
        if "layer 0 is white" not in issue_text or "(0, 255, 0)" not in issue_text or "(128, 128, 128)" not in issue_text:
            raise SystemExit(json.dumps(qcad_bad_result, ensure_ascii=False, indent=2))

        thin_solid_alias = tmp / "thin-solid-alias.dxf"
        alias_doc = ezdxf.new("R2010")
        alias_doc.layers.new("THIN_SOLID", dxfattribs={"color": 7, "lineweight": 25})
        alias_doc.layers.new("OBJECT_THICK", dxfattribs={"color": 3, "lineweight": 50})
        alias_doc.layers.new("DIM", dxfattribs={"color": 5, "lineweight": 25})
        alias_doc.layers.new("CENTER", dxfattribs={"color": 1, "lineweight": 25})
        alias_doc.layers.new("HIDDEN", dxfattribs={"color": 6, "lineweight": 25})
        alias_doc.layers.new("HATCH", dxfattribs={"color": 8, "lineweight": 25})
        alias_doc.layers.get("0").dxf.color = 4
        alias_doc.layers.get("Defpoints").dxf.color = 4
        for idx, layer_name in enumerate(("THIN_SOLID", "OBJECT_THICK", "DIM", "CENTER", "HIDDEN", "HATCH")):
            alias_doc.modelspace().add_line((0, idx), (10, idx), dxfattribs={"layer": layer_name, "color": 256})
        alias_doc.saveas(thin_solid_alias)
        alias_result = audit(thin_solid_alias)
        if not alias_result["passed"]:
            raise SystemExit(json.dumps(alias_result, ensure_ascii=False, indent=2))


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
        Path(args.report_json).write_text(json.dumps(result, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
