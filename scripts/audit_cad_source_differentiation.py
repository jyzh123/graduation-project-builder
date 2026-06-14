from __future__ import annotations

import argparse
import collections
import hashlib
import json
import math
import tempfile
import zipfile
from pathlib import Path
from typing import Any


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


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def _collect_dxf_files(path: Path, tmp: Path) -> dict[str, Path]:
    if path.is_dir():
        files = sorted(path.rglob("*.dxf"))
    elif path.suffix.lower() == ".dxf":
        files = [path]
    elif path.suffix.lower() == ".zip":
        extract_to = tmp / path.stem
        with zipfile.ZipFile(path) as zf:
            zf.extractall(extract_to)
        files = sorted(extract_to.rglob("*.dxf"))
    else:
        files = []
    return {file.name: file for file in files}


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _round(value: float, places: int = 3) -> float:
    return round(float(value), places)


def _bbox(entity: Any) -> tuple[float, float, float, float] | None:
    kind = entity.dxftype()
    try:
        if kind == "LINE":
            pts = [(entity.dxf.start.x, entity.dxf.start.y), (entity.dxf.end.x, entity.dxf.end.y)]
        elif kind in {"CIRCLE", "ARC"}:
            c = entity.dxf.center
            r = _safe_float(entity.dxf.radius)
            pts = [(c.x - r, c.y - r), (c.x + r, c.y + r)]
        elif kind == "LWPOLYLINE":
            pts = [(point[0], point[1]) for point in entity.get_points("xy")]
        elif kind == "POLYLINE":
            pts = [(vertex.dxf.location.x, vertex.dxf.location.y) for vertex in entity.vertices]
        elif kind in {"TEXT", "MTEXT"}:
            p = entity.dxf.insert
            pts = [(p.x, p.y)]
        elif kind == "DIMENSION":
            p = entity.dxf.defpoint
            pts = [(p.x, p.y)]
        elif kind == "INSERT":
            p = entity.dxf.insert
            pts = [(p.x, p.y)]
        else:
            return None
    except Exception:
        return None
    if not pts:
        return None
    return (
        min(x for x, _ in pts),
        min(y for _, y in pts),
        max(x for x, _ in pts),
        max(y for _, y in pts),
    )


def _entity_signature(entity: Any) -> tuple[Any, ...] | None:
    kind = entity.dxftype()
    if kind not in DRAWABLE_TYPES:
        return None
    box = _bbox(entity)
    layer = str(entity.dxf.get("layer", ""))
    linetype = str(entity.dxf.get("linetype", ""))
    lineweight = str(entity.dxf.get("lineweight", ""))
    if kind == "CIRCLE":
        c = entity.dxf.center
        return (
            kind,
            layer,
            linetype,
            lineweight,
            _round(c.x),
            _round(c.y),
            _round(_safe_float(entity.dxf.radius)),
        )
    if kind == "ARC":
        c = entity.dxf.center
        return (
            kind,
            layer,
            linetype,
            lineweight,
            _round(c.x),
            _round(c.y),
            _round(_safe_float(entity.dxf.radius)),
            _round(_safe_float(entity.dxf.start_angle)),
            _round(_safe_float(entity.dxf.end_angle)),
        )
    if box:
        return (kind, layer, linetype, lineweight, *(_round(v) for v in box))
    return (kind, layer, linetype, lineweight)


def _read_profile(path: Path) -> dict[str, Any]:
    import ezdxf

    doc = ezdxf.readfile(path)
    signatures: list[tuple[Any, ...]] = []
    counts: dict[str, int] = collections.Counter()
    circles: list[tuple[float, float, float]] = []
    for entity in doc.modelspace():
        kind = entity.dxftype()
        if kind not in DRAWABLE_TYPES:
            continue
        counts[kind] += 1
        signature = _entity_signature(entity)
        if signature is not None:
            signatures.append(signature)
        if kind == "CIRCLE":
            c = entity.dxf.center
            circles.append((float(c.x), float(c.y), _safe_float(entity.dxf.radius)))

    large_circle_rows: list[dict[str, Any]] = []
    small_circles = [(x, y, r) for x, y, r in circles if 0.5 <= r <= 8.0]
    for x, y, r in circles:
        if r < 80.0:
            continue
        enclosed = 0
        near = 0
        for sx, sy, _ in small_circles:
            distance = math.hypot(sx - x, sy - y)
            if distance <= r * 1.03:
                enclosed += 1
            elif distance <= r * 1.18:
                near += 1
        if enclosed >= 25:
            large_circle_rows.append(
                {
                    "center": [_round(x, 2), _round(y, 2)],
                    "radius": _round(r, 2),
                    "enclosed_small_circle_count": enclosed,
                    "near_small_circle_count": near,
                }
            )

    return {
        "path": str(path),
        "sha256": sha256_file(path),
        "entity_count": len(signatures),
        "entity_type_counts": dict(sorted(counts.items())),
        "signature_counter": collections.Counter(signatures),
        "old_like_large_circle_rows": large_circle_rows,
    }


def _compare_profiles(
    name: str,
    current_path: Path,
    baseline_path: Path | None,
    *,
    target: bool,
    min_changed_entities: int,
    min_target_change_ratio: float,
) -> dict[str, Any]:
    issues: list[str] = []
    current = _read_profile(current_path)
    baseline = _read_profile(baseline_path) if baseline_path else None
    same_sha = baseline is not None and current["sha256"] == baseline["sha256"]

    added = removed = common = 0
    change_ratio = None
    if baseline:
        current_counter = current["signature_counter"]
        baseline_counter = baseline["signature_counter"]
        common = sum((current_counter & baseline_counter).values())
        added = sum((current_counter - baseline_counter).values())
        removed = sum((baseline_counter - current_counter).values())
        denom = max(1, current["entity_count"], baseline["entity_count"])
        change_ratio = (added + removed) / denom
        if same_sha:
            issues.append("current DXF is byte-identical to baseline")
        if target and (added + removed) < min_changed_entities and change_ratio < min_target_change_ratio:
            issues.append(
                "target sheet changed too little for a source-linework redesign "
                f"(added_plus_removed={added + removed}, change_ratio={change_ratio:.4f})"
            )
    elif target:
        issues.append("target sheet has no matching baseline DXF")

    large_circle_rows = current["old_like_large_circle_rows"]
    if large_circle_rows:
        issues.append("old-like large circle overlay still encloses many tube holes")

    return {
        "file": name,
        "target_sheet": target,
        "baseline_exists": baseline is not None,
        "same_sha_as_baseline": bool(same_sha),
        "current_sha256": current["sha256"],
        "baseline_sha256": baseline["sha256"] if baseline else None,
        "current_entity_count": current["entity_count"],
        "baseline_entity_count": baseline["entity_count"] if baseline else None,
        "common_signature_count": common,
        "added_signature_count": added,
        "removed_signature_count": removed,
        "change_ratio": change_ratio,
        "current_entity_type_counts": current["entity_type_counts"],
        "baseline_entity_type_counts": baseline["entity_type_counts"] if baseline else None,
        "old_like_large_circle_rows": large_circle_rows,
        "issues": issues,
        "passed": not issues,
    }


def audit(
    current: Path,
    baseline: Path | None,
    *,
    expected_count: int | None = None,
    target_sheets: list[str] | None = None,
    min_changed_entities: int = 40,
    min_target_change_ratio: float = 0.03,
) -> dict[str, Any]:
    target_sheets = target_sheets or ["00", "01", "03"]
    current_sha256 = sha256_file(current) if current.is_file() else None
    baseline_sha256 = sha256_file(baseline) if baseline and baseline.is_file() else None
    issues: list[str] = []
    with tempfile.TemporaryDirectory(prefix="cad-source-diff-") as tmp_name:
        tmp = Path(tmp_name)
        current_files = _collect_dxf_files(current, tmp / "current")
        baseline_files = _collect_dxf_files(baseline, tmp / "baseline") if baseline else {}
        rows = [
            _compare_profiles(
                name,
                current_path,
                baseline_files.get(name),
                target=any(name.startswith(f"EA11-V100-{token}") or name.startswith(token) for token in target_sheets),
                min_changed_entities=min_changed_entities,
                min_target_change_ratio=min_target_change_ratio,
            )
            for name, current_path in sorted(current_files.items())
        ]

    if not rows:
        issues.append("no current DXF files found")
    if expected_count is not None and len(rows) < expected_count:
        issues.append(f"current DXF count {len(rows)} is below expected {expected_count}")
    if baseline and not baseline_files:
        issues.append("no baseline DXF files found")

    identical_dxf_count = sum(1 for row in rows if row["same_sha_as_baseline"])
    changed_dxf_count = sum(1 for row in rows if row["baseline_exists"] and not row["same_sha_as_baseline"])
    target_tiny_change_count = sum(
        1
        for row in rows
        if row["target_sheet"]
        and row["baseline_exists"]
        and not row["same_sha_as_baseline"]
        and (row["added_signature_count"] + row["removed_signature_count"]) < min_changed_entities
        and (row["change_ratio"] or 0.0) < min_target_change_ratio
    )
    old_like_large_circle_overlay_count = sum(len(row["old_like_large_circle_rows"]) for row in rows)

    if baseline and identical_dxf_count:
        issues.append(f"{identical_dxf_count} DXF files are still identical to the baseline")
    if target_tiny_change_count:
        issues.append(f"{target_tiny_change_count} target sheets changed too little")
    if old_like_large_circle_overlay_count:
        issues.append(f"{old_like_large_circle_overlay_count} old-like large circle overlays remain")

    for row in rows:
        issues.extend(f"{row['file']}: {issue}" for issue in row["issues"])

    return {
        "schema": "graduation-project-builder.cad-source-differentiation-audit.v1",
        "current_path": str(current),
        "baseline_path": str(baseline) if baseline else None,
        "current_sha256": current_sha256,
        "baseline_sha256": baseline_sha256,
        "expected_dxf_count": expected_count,
        "dxf_file_count": len(rows),
        "target_sheets": target_sheets,
        "min_changed_entities": min_changed_entities,
        "min_target_change_ratio": min_target_change_ratio,
        "identical_dxf_count": identical_dxf_count,
        "changed_dxf_count": changed_dxf_count,
        "target_tiny_change_count": target_tiny_change_count,
        "old_like_large_circle_overlay_count": old_like_large_circle_overlay_count,
        "passed": not issues,
        "issues": issues,
        "per_file": rows,
    }


def self_test() -> None:
    import ezdxf

    with tempfile.TemporaryDirectory(prefix="cad-source-diff-selftest-") as tmp_name:
        tmp = Path(tmp_name)
        baseline_dir = tmp / "baseline"
        tiny_dir = tmp / "tiny"
        redesigned_dir = tmp / "redesigned"
        baseline_dir.mkdir()
        tiny_dir.mkdir()
        redesigned_dir.mkdir()

        baseline = baseline_dir / "EA11-V100-00_A0_assembly.dxf"
        doc = ezdxf.new("R2000")
        msp = doc.modelspace()
        msp.add_circle((100, 100), 95)
        for row in range(6):
            for col in range(10):
                msp.add_circle((35 + col * 14, 55 + row * 14), 3)
        doc.saveas(baseline)

        tiny = tiny_dir / baseline.name
        doc = ezdxf.readfile(baseline)
        doc.modelspace().add_line((0, 0), (1, 1))
        doc.saveas(tiny)
        tiny_result = audit(tiny_dir, baseline_dir, expected_count=1)
        if tiny_result["passed"]:
            raise SystemExit("tiny-change self-test should fail")

        redesigned = redesigned_dir / baseline.name
        doc = ezdxf.new("R2000")
        msp = doc.modelspace()
        for zone in range(3):
            x0 = 20 + zone * 70
            msp.add_lwpolyline([(x0, 20), (x0 + 50, 20), (x0 + 50, 95), (x0, 95)], close=True)
            for row in range(3):
                for col in range(5):
                    msp.add_circle((x0 + 8 + col * 9, 35 + row * 16), 2.5)
                    msp.add_line((x0 + 8 + col * 9 - 4, 35 + row * 16), (x0 + 8 + col * 9 + 4, 35 + row * 16))
                    msp.add_line((x0 + 8 + col * 9, 35 + row * 16 - 4), (x0 + 8 + col * 9, 35 + row * 16 + 4))
        doc.saveas(redesigned)
        redesigned_result = audit(redesigned_dir, baseline_dir, expected_count=1)
        if not redesigned_result["passed"]:
            raise SystemExit(json.dumps(redesigned_result, ensure_ascii=False, indent=2))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--current", required=False)
    parser.add_argument("--baseline")
    parser.add_argument("--expected-count", type=int)
    parser.add_argument("--target-sheets", default="00,01,03")
    parser.add_argument("--min-changed-entities", type=int, default=40)
    parser.add_argument("--min-target-change-ratio", type=float, default=0.03)
    parser.add_argument("--report-json")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
        print("self-test passed")
        return
    if not args.current:
        parser.error("--current is required unless --self-test is used")
    result = audit(
        Path(args.current),
        Path(args.baseline) if args.baseline else None,
        expected_count=args.expected_count,
        target_sheets=[token.strip() for token in args.target_sheets.split(",") if token.strip()],
        min_changed_entities=args.min_changed_entities,
        min_target_change_ratio=args.min_target_change_ratio,
    )
    if args.report_json:
        Path(args.report_json).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
