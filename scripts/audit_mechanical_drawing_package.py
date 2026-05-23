#!/usr/bin/env python3
"""Audit a mechanical drawing delivery package for CAD-source and density risks."""

from __future__ import annotations

import argparse
import hashlib
import json
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

TEXT_SUFFIXES = {".dxf", ".json", ".txt", ".csv", ".md", ".xml", ".scr", ".lsp"}
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
    "title_block_table_notes_isolation_verdict",
    "annotation_margin_clearance_verdict",
    "local_crowding_verdict",
    "sheet_layout_verdict",
    "manufacturing_view_depth_verdict",
    "layout_collision_verdict",
)
PER_SHEET_RENDERED_REVIEW_PASS_FIELDS = (
    "no_overlap_verdict",
    "boundary_clearance_verdict",
    "detail_density_verdict",
    "text_legibility_verdict",
    "title_block_table_notes_isolation_verdict",
    "annotation_margin_clearance_verdict",
    "local_crowding_verdict",
    "layout_collision_verdict",
)


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


def collect_rendered_review_payloads(payload: object) -> list[tuple[str, dict[str, object]]]:
    reviews: list[tuple[str, dict[str, object]]] = []
    seen: set[int] = set()

    def walk(value: object, location: str) -> None:
        if isinstance(value, dict):
            review = value.get("rendered_review")
            if isinstance(review, dict) and id(review) not in seen:
                seen.add(id(review))
                reviews.append((f"{location}.rendered_review", review))
            if (
                "entity_count_only_verdict" in value
                or "rendered_sheet_previews" in value
                or "preview_paths" in value
                or "manufacturing_view_depth_verdict" in value
                or "sheet_layout_verdict" in value
            ) and id(value) not in seen:
                seen.add(id(value))
                reviews.append((location, value))
            for key, child in value.items():
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
        return True
    if package_path.is_dir() and (package_path / reference).exists():
        return True
    normalized = reference.replace("\\", "/").lstrip("./")
    return normalized in member_names


def rendered_review_stats(files: Iterable[PackageFile], package_path: Path) -> dict[str, object]:
    package_files = list(files)
    member_names = {file.name.replace("\\", "/").lstrip("./") for file in package_files}
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
        for location, review in collect_rendered_review_payloads(payload):
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
    for file in files:
        if file.suffix not in TEXT_SUFFIXES:
            continue
        text, _, _ = decode_text(file.data)
        hits = text_mojibake_hits(file.name, text)
        if hits:
            text_mojibake.append({"file": file.name, "tokens": hits[:12]})

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
        "rendered_review": rendered_review_stats(files, path),
        "filename_mojibake_hits": filename_mojibake,
        "text_mojibake_hits": text_mojibake,
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
    json_manifest = candidate["json_manifest"]  # type: ignore[assignment]
    if json_manifest["failures"]:  # type: ignore[index]
        issues.append("one or more JSON manifest files are invalid or not UTF-8")
    if json_manifest["mojibake_hits"]:  # type: ignore[index]
        issues.append("one or more JSON manifest files contain mojibake-like tokens")
    rendered_review = candidate["rendered_review"]  # type: ignore[assignment]
    if require_rendered_review and not rendered_review["passed"]:  # type: ignore[index]
        issues.append("rendered-sheet visual review is missing or failing")

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
    if min_dimensions_per_dxf and min_dimensioned_dxf_files:
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
        if dimensioned_files < min_dimensioned_dxf_files:
            density_failures.append(
                {
                    "kind": "dimension_distribution",
                    "actual_dimensioned_files": dimensioned_files,
                    "required_dimensioned_files": min_dimensioned_dxf_files,
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
    if reference_dwg_bytes and min_reference_dwg_ratio:
        ratio = candidate_dwg_bytes / reference_dwg_bytes if reference_dwg_bytes else 0.0
        candidate["reference_dwg_byte_ratio"] = round(ratio, 4)
        if ratio < min_reference_dwg_ratio:
            issues.append(
                f"DWG byte density ratio below reference threshold: {ratio:.4f} < {min_reference_dwg_ratio:.4f}"
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

    return {
        "schema": "graduation-project-builder.mechanical-drawing-package-audit.v4",
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
            "min_geometry_entities": min_geometry_entities,
            "min_insert_count": min_insert_count,
            "min_arc_count": min_arc_count,
            "min_hatch_count": min_hatch_count,
            "min_text_count": min_text_count,
            "min_dxf_entities_per_sheet": min_dxf_entities_per_sheet,
            "min_dimensions_per_dxf": min_dimensions_per_dxf,
            "min_dimensioned_dxf_files": min_dimensioned_dxf_files,
            "min_pdf_drawing_objects_per_a0": min_pdf_drawing_objects_per_a0,
            "required_text_tokens": required_text_tokens,
            "require_rendered_review": require_rendered_review,
        },
        "drawing_density_checks": {
            "failures": density_failures,
        },
        "density_verdict": density_verdict,
        "manufacturing_depth": manufacturing_depth,
        "rendered_review_verdict": {
            "passed": bool(rendered_review["passed"]),  # type: ignore[index]
            "review_count": int(rendered_review["review_count"]),  # type: ignore[index]
            "accepted_review_count": int(rendered_review["accepted_review_count"]),  # type: ignore[index]
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
        entity_only.mkdir()
        overlap.mkdir()
        (good / "sheet.dwg").write_bytes(b"AC1032\x00" + b"x" * 4096)
        (good / "sheet.pdf").write_bytes(b"%PDF-1.4\n%%EOF\n")
        (good / "sheet-preview.png").write_bytes(b"\x89PNG\r\n\x1a\n")
        (good / "sheet.dxf").write_text(
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n0\nSTYLE\n0\nDIMSTYLE\n0\nLTYPE\n0\nENDSEC\n"
            "0\nSECTION\n2\nBLOCKS\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n0\nDIMENSION\n0\nENDSEC\n0\nEOF\n",
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
                    },
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        (bad / "renamed.dwg").write_bytes(b"0\nSECTION\n")
        (bad / "bad.json").write_text('{"title":"\\u934a\\u53bb"', encoding="utf-8")
        for target in (entity_only, overlap):
            for child in good.iterdir():
                (target / child.name).write_bytes(child.read_bytes())
        (entity_only / "manifest.json").write_text('{"files":["sheet.dwg","sheet.pdf","sheet.dxf"]}', encoding="utf-8")
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
                        "annotation_margin_clearance_verdict": "pass",
                        "local_crowding_verdict": "fail",
                        "sheet_layout_verdict": "pass",
                        "manufacturing_view_depth_verdict": "pass",
                        "layout_collision_verdict": "fail",
                        "entity_count_only_verdict": "not-used",
                        "per_sheet": [
                            {
                                "sheet": "sheet",
                                "preview_path": "sheet-preview.png",
                                "no_overlap_verdict": "fail",
                                "boundary_clearance_verdict": "fail",
                                "detail_density_verdict": "pass",
                                "text_legibility_verdict": "pass",
                                "title_block_table_notes_isolation_verdict": "fail",
                                "annotation_margin_clearance_verdict": "pass",
                                "local_crowding_verdict": "fail",
                                "layout_collision_verdict": "fail",
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
        if not good_result["passed"]:
            print(json.dumps(good_result, ensure_ascii=False, indent=2))
            return 1
        if bad_result["passed"]:
            print(json.dumps(bad_result, ensure_ascii=False, indent=2))
            return 1
        if entity_only_result["passed"]:
            print(json.dumps(entity_only_result, ensure_ascii=False, indent=2))
            return 1
        if overlap_result["passed"]:
            print(json.dumps(overlap_result, ensure_ascii=False, indent=2))
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
    parser.add_argument("--min-reference-dwg-ratio", type=float, default=0.25)
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
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
