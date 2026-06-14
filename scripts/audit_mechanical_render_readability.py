#!/usr/bin/env python3
"""Audit rendered mechanical CAD sheets for text overlap and severe line crowding."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable

import numpy as np
from PIL import Image


TEXT_GRAPHIC_MIN_OVERLAP_PT2 = 1.5
TEXT_GRAPHIC_MIN_FRACTION = 0.015
TEXT_TEXT_MIN_FRACTION = 0.08
DEFAULT_TILE_SIZE = 96
DEFAULT_MAX_TILE_INK_DENSITY = 0.72
DEFAULT_MAX_SEVERE_TILE_COUNT = 0
INK_DISTANCE_THRESHOLD = 28.0
PDF_LINE_HIT_PAD_PT = 0.45
PDF_FILLED_RECT_MAX_AREA_PT2 = 250.0
TEXT_GRAPHIC_MIN_SUBSTANTIVE_OVERLAP_PT2 = 8.0
TEXT_GRAPHIC_MIN_SUBSTANTIVE_FRACTION = 0.12


def _safe_path(path: Path) -> str:
    """Return a JSON-safe path string that survives Windows/Chinese paths."""
    return str(path).replace("\\", "/")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def _collect_files(paths: Iterable[Path], suffix: str) -> list[Path]:
    out: list[Path] = []
    tmp_roots: list[tempfile.TemporaryDirectory[str]] = list(getattr(_collect_files, "_tmp_roots", []))
    for path in paths:
        if not path.exists():
            continue
        if path.is_dir():
            out.extend(sorted(path.rglob(f"*{suffix}")))
        elif path.suffix.lower() == suffix:
            out.append(path)
        elif path.suffix.lower() == ".zip":
            tmp = tempfile.TemporaryDirectory()
            tmp_roots.append(tmp)
            with zipfile.ZipFile(path) as zf:
                zf.extractall(tmp.name)
            out.extend(sorted(Path(tmp.name).rglob(f"*{suffix}")))
    # Keep temp directories alive long enough for callers to read files by path.
    setattr(_collect_files, "_tmp_roots", tmp_roots)
    return sorted(out)


def _area(rect: tuple[float, float, float, float]) -> float:
    return max(0.0, rect[2] - rect[0]) * max(0.0, rect[3] - rect[1])


def _intersect(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
    return (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))


def _boxes_intersect(a: tuple[float, float, float, float], b: tuple[float, float, float, float], tol: float = 0.0) -> bool:
    return not (a[2] < b[0] - tol or a[0] > b[2] + tol or a[3] < b[1] - tol or a[1] > b[3] + tol)


def _shrink_rect(rect: tuple[float, float, float, float], x_frac: float = 0.08, y_frac: float = 0.18) -> tuple[float, float, float, float]:
    w = max(0.0, rect[2] - rect[0])
    h = max(0.0, rect[3] - rect[1])
    return (rect[0] + w * x_frac, rect[1] + h * y_frac, rect[2] - w * x_frac, rect[3] - h * y_frac)


def _rect_tuple(rect: Any) -> tuple[float, float, float, float]:
    return (float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3]))


def _is_dimensionish_text(text: str) -> bool:
    clean = text.strip().replace(" ", "")
    if not clean:
        return True
    if clean in {"●", "○", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}:
        return True
    if re.fullmatch(r"[\d()（）+\-./°~～ΦφA-Za-z]+", clean):
        return True
    if re.fullmatch(r"[A-Z]-[A-Z]", clean):
        return True
    if re.fullmatch(r"[A-Z]向局部", clean):
        return True
    return False


def _is_substantive_text(text: str) -> bool:
    clean = text.strip().replace(" ", "")
    if _is_dimensionish_text(clean):
        return False
    cjk_count = len(re.findall(r"[\u4e00-\u9fff]", clean))
    if cjk_count >= 3:
        return True
    if len(clean) >= 6 and re.search(r"[A-Za-z]", clean):
        return True
    return False


def _line_hit_rect(
    p0: Any,
    p1: Any,
    pad: float,
) -> tuple[float, float, float, float]:
    x0, y0 = float(p0[0]), float(p0[1])
    x1, y1 = float(p1[0]), float(p1[1])
    return (min(x0, x1) - pad, min(y0, y1) - pad, max(x0, x1) + pad, max(y0, y1) + pad)


def _rect_edge_hit_rects(rect: Any, pad: float) -> list[tuple[float, float, float, float]]:
    x0, y0, x1, y1 = _rect_tuple(rect)
    return [
        (x0 - pad, y0 - pad, x1 + pad, y0 + pad),
        (x0 - pad, y1 - pad, x1 + pad, y1 + pad),
        (x0 - pad, y0 - pad, x0 + pad, y1 + pad),
        (x1 - pad, y0 - pad, x1 + pad, y1 + pad),
    ]


def _drawing_hit_rects(page: Any) -> list[tuple[float, float, float, float]]:
    """Convert drawing paths to narrow stroke hitboxes instead of broad path bounds."""
    hit_rects: list[tuple[float, float, float, float]] = []
    for drawing in page.get_drawings():
        width = float(drawing.get("width") or 0.5)
        pad = max(PDF_LINE_HIT_PAD_PT, width * 0.75)
        items = drawing.get("items") or []
        for item in items:
            if not item:
                continue
            op = item[0]
            if op == "l" and len(item) >= 3:
                hit_rects.append(_line_hit_rect(item[1], item[2], pad))
            elif op == "re" and len(item) >= 2:
                hit_rects.extend(_rect_edge_hit_rects(item[1], pad))
            elif op in {"c", "qu"}:
                points = []
                for value in item[1:]:
                    if hasattr(value, "x") and hasattr(value, "y"):
                        points.append((float(value.x), float(value.y)))
                    elif isinstance(value, (tuple, list)) and len(value) >= 2:
                        points.append((float(value[0]), float(value[1])))
                if points:
                    xs = [p[0] for p in points]
                    ys = [p[1] for p in points]
                    hit_rects.append((min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad))

        # Filled small graphics can genuinely cover labels. Large frames/tables are
        # handled by their edges above, not by their broad bounding rectangles.
        rect = drawing.get("rect")
        if drawing.get("fill") and rect:
            tup = _rect_tuple(rect)
            if 0.25 <= _area(tup) <= PDF_FILLED_RECT_MAX_AREA_PT2:
                hit_rects.append(tup)
        elif not items and rect:
            tup = _rect_tuple(rect)
            if 0.25 <= _area(tup) <= PDF_FILLED_RECT_MAX_AREA_PT2:
                hit_rects.append(tup)
    return hit_rects


def _pdf_text_graphic_audit(pdf: Path) -> dict[str, Any]:
    import fitz

    doc = fitz.open(str(pdf))
    page_rows: list[dict[str, Any]] = []
    total_text_spans = 0
    total_graphic_rects = 0
    total_text_text_overlap = 0
    total_text_graphic_cover = 0
    worst_text_graphic_fraction = 0.0
    worst_text_text_fraction = 0.0
    samples: list[dict[str, Any]] = []

    for page_index, page in enumerate(doc):
        text_spans: list[dict[str, Any]] = []
        text_dict = page.get_text("dict")
        for block in text_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = str(span.get("text", "")).strip()
                    bbox = span.get("bbox")
                    if not text or not bbox:
                        continue
                    rect = _rect_tuple(bbox)
                    if _area(rect) < 1.0:
                        continue
                    text_spans.append({"text": text, "bbox": rect})

        graphic_rects = _drawing_hit_rects(page)

        text_text_overlap = 0
        text_graphic_cover = 0
        graphics = np.array(graphic_rects, dtype=float) if graphic_rects else np.empty((0, 4), dtype=float)

        for i, left in enumerate(text_spans):
            left_area = _area(left["bbox"])
            for right in text_spans[i + 1 :]:
                inter = _area(_intersect(left["bbox"], right["bbox"]))
                if inter <= 0:
                    continue
                frac = inter / max(1.0, min(left_area, _area(right["bbox"])))
                worst_text_text_fraction = max(worst_text_text_fraction, frac)
                if frac >= TEXT_TEXT_MIN_FRACTION:
                    if not (_is_substantive_text(left["text"]) and _is_substantive_text(right["text"])):
                        continue
                    text_text_overlap += 1
                    if len(samples) < 20:
                        samples.append(
                            {
                                "type": "text_text_overlap",
                                "page": page_index + 1,
                                "left": left["text"][:60],
                                "right": right["text"][:60],
                                "fraction": round(frac, 4),
                                "bbox": [round(v, 2) for v in _intersect(left["bbox"], right["bbox"])],
                            }
                        )

            if graphics.size:
                if not _is_substantive_text(left["text"]):
                    continue
                shrunken_text_bbox = _shrink_rect(left["bbox"])
                l = np.array(shrunken_text_bbox, dtype=float)
                ix0 = np.maximum(l[0], graphics[:, 0])
                iy0 = np.maximum(l[1], graphics[:, 1])
                ix1 = np.minimum(l[2], graphics[:, 2])
                iy1 = np.minimum(l[3], graphics[:, 3])
                inter_areas = np.maximum(0.0, ix1 - ix0) * np.maximum(0.0, iy1 - iy0)
                if inter_areas.size:
                    max_inter = float(inter_areas.max())
                    worst_text_graphic_fraction = max(worst_text_graphic_fraction, max_inter / max(1.0, left_area))
                hit_indexes = np.where(
                    (inter_areas >= TEXT_GRAPHIC_MIN_SUBSTANTIVE_OVERLAP_PT2)
                    & ((inter_areas / max(1.0, _area(shrunken_text_bbox))) >= TEXT_GRAPHIC_MIN_SUBSTANTIVE_FRACTION)
                )[0]
                if len(hit_indexes):
                    text_graphic_cover += int(len(hit_indexes))
                    if len(samples) < 20:
                        idx = int(hit_indexes[0])
                        rect = tuple(float(v) for v in graphics[idx])
                        samples.append(
                            {
                                "type": "text_graphic_cover",
                                "page": page_index + 1,
                                "text": left["text"][:60],
                                "fraction": round(float(inter_areas[idx]) / max(1.0, _area(shrunken_text_bbox)), 4),
                                "bbox": [round(v, 2) for v in _intersect(shrunken_text_bbox, rect)],
                            }
                        )

        total_text_spans += len(text_spans)
        total_graphic_rects += len(graphic_rects)
        total_text_text_overlap += text_text_overlap
        total_text_graphic_cover += text_graphic_cover
        page_rows.append(
            {
                "page": page_index + 1,
                "text_span_count": len(text_spans),
                "graphic_rect_count": len(graphic_rects),
                "text_text_overlap_count": text_text_overlap,
                "text_graphic_cover_count": text_graphic_cover,
                "passed": text_text_overlap == 0 and text_graphic_cover == 0,
            }
        )

    doc.close()
    return {
        "pdf": _safe_path(pdf),
        "sha256": sha256_file(pdf),
        "page_count": len(page_rows),
        "text_span_count": total_text_spans,
        "graphic_rect_count": total_graphic_rects,
        "text_text_overlap_count": total_text_text_overlap,
        "text_graphic_cover_count": total_text_graphic_cover,
        "worst_text_text_overlap_fraction": round(worst_text_text_fraction, 5),
        "worst_text_graphic_cover_fraction": round(worst_text_graphic_fraction, 5),
        "samples": samples,
        "per_page": page_rows,
        "passed": total_text_text_overlap == 0 and total_text_graphic_cover == 0,
    }


def _distance_from_bg(pixel: tuple[int, int, int], bg: tuple[int, int, int]) -> float:
    return math.sqrt(sum((float(pixel[i]) - float(bg[i])) ** 2 for i in range(3)))


def _dominant_corner_background(img: Image.Image) -> tuple[int, int, int]:
    rgb = img.convert("RGB")
    arr = np.asarray(rgb, dtype=np.uint8)
    height, width = arr.shape[:2]
    s = min(32, width, height)
    samples = np.concatenate(
        [
            arr[:s, :s].reshape(-1, 3),
            arr[:s, width - s :].reshape(-1, 3),
            arr[height - s :, :s].reshape(-1, 3),
            arr[height - s :, width - s :].reshape(-1, 3),
        ],
        axis=0,
    )
    med = np.median(samples, axis=0)
    return (int(med[0]), int(med[1]), int(med[2]))


def _quantized_mode_rgb(tile: np.ndarray) -> tuple[int, int, int]:
    quantized = (tile // 8).astype(np.int32)
    keys = (
        (quantized[:, :, 0] << 16)
        + (quantized[:, :, 1] << 8)
        + quantized[:, :, 2]
    ).reshape(-1)
    values, counts = np.unique(keys, return_counts=True)
    key = int(values[int(np.argmax(counts))])
    return (((key >> 16) & 255) * 8, ((key >> 8) & 255) * 8, (key & 255) * 8)


def _tile_ink_density(tile: np.ndarray) -> tuple[float, tuple[int, int, int]]:
    bg = np.array(_quantized_mode_rgb(tile), dtype=np.int32)
    dist = np.sqrt(np.sum((tile.astype(np.int32) - bg) ** 2, axis=2))
    ink_mask = dist >= INK_DISTANCE_THRESHOLD
    return float(np.count_nonzero(ink_mask)) / float(tile.shape[0] * tile.shape[1]), tuple(int(v) for v in bg)


def _png_line_crowding_audit(png: Path, tile_size: int, max_density: float, max_severe_tiles: int) -> dict[str, Any]:
    img = Image.open(png).convert("RGB")
    width, height = img.size
    corner_bg = _dominant_corner_background(img)
    arr = np.asarray(img, dtype=np.uint8)
    severe_tiles: list[dict[str, Any]] = []
    max_observed = 0.0
    checked = 0
    observed_backgrounds: dict[tuple[int, int, int], int] = {}

    for y in range(0, height, tile_size):
        for x in range(0, width, tile_size):
            w = min(tile_size, width - x)
            h = min(tile_size, height - y)
            if w < tile_size // 2 or h < tile_size // 2:
                continue
            checked += 1
            density, local_bg = _tile_ink_density(arr[y : y + h, x : x + w])
            observed_backgrounds[local_bg] = observed_backgrounds.get(local_bg, 0) + 1
            max_observed = max(max_observed, density)
            if density > max_density:
                severe_tiles.append(
                    {
                        "bbox_px": [x, y, x + w, y + h],
                        "ink_density": round(density, 5),
                        "local_background_rgb": local_bg,
                    }
                )

    severe_tiles = sorted(severe_tiles, key=lambda row: row["ink_density"], reverse=True)
    return {
        "png": _safe_path(png),
        "sha256": sha256_file(png),
        "width_px": width,
        "height_px": height,
        "corner_background_rgb": corner_bg,
        "dominant_tile_backgrounds": [
            {"rgb": list(rgb), "tile_count": count}
            for rgb, count in sorted(observed_backgrounds.items(), key=lambda row: row[1], reverse=True)[:8]
        ],
        "tile_size_px": tile_size,
        "checked_tile_count": checked,
        "max_allowed_tile_ink_density": max_density,
        "max_observed_tile_ink_density": round(max_observed, 5),
        "severe_line_crowding_count": len(severe_tiles),
        "severe_line_crowding_samples": severe_tiles[:20],
        "passed": len(severe_tiles) <= max_severe_tiles,
    }


def audit(
    pdf_paths: list[Path],
    png_paths: list[Path],
    tile_size: int,
    max_tile_ink_density: float,
    max_severe_tiles: int,
) -> dict[str, Any]:
    pdf_rows = [_pdf_text_graphic_audit(path) for path in pdf_paths]
    png_rows = [
        _png_line_crowding_audit(path, tile_size, max_tile_ink_density, max_severe_tiles)
        for path in png_paths
    ]
    text_text_overlap_count = sum(int(row["text_text_overlap_count"]) for row in pdf_rows)
    text_graphic_cover_count = sum(int(row["text_graphic_cover_count"]) for row in pdf_rows)
    severe_line_crowding_count = sum(int(row["severe_line_crowding_count"]) for row in png_rows)
    passed = (
        bool(pdf_rows or png_rows)
        and text_text_overlap_count == 0
        and text_graphic_cover_count == 0
        and all(row["passed"] for row in png_rows)
    )
    return {
        "schema": "graduation-project-builder.mechanical-render-readability.v1",
        "audit_method": "pdf-text-bbox-plus-png-local-ink-density-machine",
        "passed": passed,
        "pdf_count": len(pdf_rows),
        "png_count": len(png_rows),
        "text_text_overlap_count": text_text_overlap_count,
        "text_graphic_cover_count": text_graphic_cover_count,
        "severe_line_crowding_count": severe_line_crowding_count,
        "tile_size_px": tile_size,
        "max_allowed_tile_ink_density": max_tile_ink_density,
        "max_allowed_severe_tiles_per_sheet": max_severe_tiles,
        "pdf_rows": pdf_rows,
        "png_rows": png_rows,
    }


def run_self_test() -> int:
    import fitz

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        good_pdf = root / "good.pdf"
        bad_pdf = root / "bad.pdf"
        for path, bad in ((good_pdf, False), (bad_pdf, True)):
            doc = fitz.open()
            page = doc.new_page(width=300, height=200)
            page.insert_text((30, 40), "CLEAR TEXT", fontsize=12)
            page.draw_line((30, 80), (250, 80), color=(0, 0, 0), width=1)
            if bad:
                page.draw_line((30, 38), (120, 38), color=(0, 0, 0), width=2)
                page.insert_text((32, 40), "CLEAR TEXT", fontsize=12)
            doc.save(path)
            doc.close()

        good_png = root / "good.png"
        bad_png = root / "bad.png"
        Image.new("RGB", (240, 160), (255, 255, 255)).save(good_png)
        img = Image.new("RGB", (240, 160), (255, 255, 255))
        for x in range(10, 230):
            for y in range(10, 150):
                if (x + y) % 4 != 0:
                    img.putpixel((x, y), (0, 0, 0))
        img.save(bad_png)

        good = audit([good_pdf], [good_png], 48, 0.42, 0)
        bad = audit([bad_pdf], [bad_png], 48, 0.42, 0)
        if not good["passed"]:
            print(json.dumps(good, ensure_ascii=False, indent=2))
            return 1
        if bad["passed"]:
            print(json.dumps(bad, ensure_ascii=False, indent=2))
            return 1
    print("self-test passed")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", action="append", default=[], type=Path, help="PDF file, folder, or ZIP to audit")
    parser.add_argument("--png", action="append", default=[], type=Path, help="PNG file, folder, or ZIP to audit")
    parser.add_argument("--report-json", type=Path)
    parser.add_argument("--tile-size", type=int, default=DEFAULT_TILE_SIZE)
    parser.add_argument("--max-tile-ink-density", type=float, default=DEFAULT_MAX_TILE_INK_DENSITY)
    parser.add_argument("--max-severe-tiles-per-sheet", type=int, default=DEFAULT_MAX_SEVERE_TILE_COUNT)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        return run_self_test()
    pdfs = _collect_files(args.pdf, ".pdf")
    pngs = _collect_files(args.png, ".png")
    result = audit(pdfs, pngs, args.tile_size, args.max_tile_ink_density, args.max_severe_tiles_per_sheet)
    text = json.dumps(result, ensure_ascii=True, indent=2)
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
