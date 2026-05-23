#!/usr/bin/env python3
"""Inspect DOCX pagination-critical section, page-number, and header/footer structure."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
SCHEMA = "graduation-project-builder.docx-pagination-structure.v1"
GENERATOR = "inspect_docx_pagination_structure.py"
BLANK_INK_RATIO_MAX = 0.00005
NEAR_EMPTY_INK_RATIO_MAX = 0.006
FATAL_TOPOLOGY_DIFFERENCES = {
    "section_count",
    "section_property_map",
    "page_number_format_restart_map",
    "header_footer_reference_map",
    "header_footer_link_to_previous_inferred_map",
    "footer_page_field_map",
}
CONTENT_GROWTH_DIFFERENCES = {
    "section_count",
    "section_boundary_map",
    "section_property_map",
    "page_number_format_restart_map",
    "header_footer_reference_map",
    "header_footer_link_to_previous_inferred_map",
    "hard_page_break_section_break_map",
    "footer_page_field_map",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except (KeyError, ET.ParseError):
        return None


def w_attr(node: ET.Element | None, local: str, default: str = "") -> str:
    if node is None:
        return default
    return node.attrib.get(f"{W}val", node.attrib.get(f"{W}{local}", default))


def node_attrs(node: ET.Element | None, names: tuple[str, ...]) -> dict[str, str]:
    if node is None:
        return {}
    return {name: node.attrib.get(f"{W}{name}", "") for name in names}


def rels_map(zf: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    root = read_xml(zf, "word/_rels/document.xml.rels")
    if root is None:
        return {}
    result: dict[str, dict[str, str]] = {}
    for rel in root.findall(f"{REL_NS}Relationship"):
        rel_id = rel.attrib.get("Id", "")
        if not rel_id:
            continue
        result[rel_id] = {
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
        }
    return result


def footer_page_field_map(zf: zipfile.ZipFile, rels: dict[str, dict[str, str]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for rel_id, rel in rels.items():
        target = rel.get("target", "")
        if "footer" not in target and "header" not in target:
            continue
        part = "word/" + target.lstrip("/")
        root = read_xml(zf, part)
        if root is None:
            result[rel_id] = 0
            continue
        text = "".join(root.itertext()).upper()
        instr = " ".join(
            (node.text or "").upper()
            for node in root.iter(f"{W}instrText")
        )
        simple = " ".join(
            node.attrib.get(f"{W}instr", "").upper()
            for node in root.iter(f"{W}fldSimple")
        )
        result[rel_id] = text.count("PAGE") + instr.count("PAGE") + simple.count("PAGE")
    return result


def live_toc_field_count(root: ET.Element) -> int:
    field_texts: list[str] = []
    field_texts.extend((node.text or "") for node in root.iter(f"{W}instrText"))
    field_texts.extend(node.attrib.get(f"{W}instr", "") for node in root.iter(f"{W}fldSimple"))
    return sum(1 for text in field_texts if re.search(r"(^|\s)TOC(\s|$)", text, re.IGNORECASE))


def section_refs(sect_pr: ET.Element, local: str) -> list[dict[str, str]]:
    refs = []
    for ref in sect_pr.findall(f"{W}{local}Reference"):
        refs.append(
            {
                "type": ref.attrib.get(f"{W}type", "default"),
                "rId": ref.attrib.get(f"{R}id", ""),
            }
        )
    return refs


def section_signature(sect_pr: ET.Element, index: int, paragraph_index: int) -> dict[str, object]:
    pg_num = sect_pr.find(f"{W}pgNumType")
    pg_sz = sect_pr.find(f"{W}pgSz")
    pg_mar = sect_pr.find(f"{W}pgMar")
    cols = sect_pr.find(f"{W}cols")
    sect_type = sect_pr.find(f"{W}type")
    return {
        "index": index,
        "paragraph_index": paragraph_index,
        "type": w_attr(sect_type, "val", "continuous-or-default"),
        "pgNumType": node_attrs(pg_num, ("fmt", "start", "chapStyle", "chapSep")),
        "pgSz": node_attrs(pg_sz, ("w", "h", "orient")),
        "pgMar": node_attrs(pg_mar, ("top", "right", "bottom", "left", "header", "footer", "gutter")),
        "cols": node_attrs(cols, ("num", "space")),
        "titlePg": sect_pr.find(f"{W}titlePg") is not None,
        "headerReferences": section_refs(sect_pr, "header"),
        "footerReferences": section_refs(sect_pr, "footer"),
    }


def infer_link_to_previous(sections: list[dict[str, object]], ref_name: str) -> list[dict[str, object]]:
    result = []
    previous_types: set[str] = set()
    for section in sections:
        refs = section.get(ref_name, [])
        current_types = {str(ref.get("type", "default")) for ref in refs if isinstance(ref, dict)}
        inherited = sorted(previous_types - current_types)
        result.append(
            {
                "section": section["index"],
                "inferred_link_ref_types": inherited,
            }
        )
        if current_types:
            previous_types = current_types
    return result


def inspect_docx(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as zf:
        root = read_xml(zf, "word/document.xml")
        if root is None:
            raise ValueError(f"missing or unreadable word/document.xml: {path}")
        rels = rels_map(zf)
        body = root.find(f"{W}body")
        paragraphs = list(body.findall(f"{W}p")) if body is not None else []
        sections: list[dict[str, object]] = []
        for idx, paragraph in enumerate(paragraphs, start=1):
            sect_pr = paragraph.find(f"{W}pPr/{W}sectPr")
            if sect_pr is not None:
                sections.append(section_signature(sect_pr, len(sections) + 1, idx))
        if body is not None:
            body_sect = body.find(f"{W}sectPr")
            if body_sect is not None:
                sections.append(section_signature(body_sect, len(sections) + 1, len(paragraphs) + 1))
        hard_page_breaks = [
            idx
            for idx, br in enumerate(root.iter(f"{W}br"), start=1)
            if br.attrib.get(f"{W}type", "textWrapping") == "page"
        ]
        return {
            "section_count": len(sections),
            "sections": sections,
            "hard_page_breaks": hard_page_breaks,
            "relationships": rels,
            "footer_page_field_map": footer_page_field_map(zf, rels),
            "live_toc_field_count": live_toc_field_count(root),
            "header_link_to_previous_inferred": infer_link_to_previous(sections, "headerReferences"),
            "footer_link_to_previous_inferred": infer_link_to_previous(sections, "footerReferences"),
        }


def compact(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def rendered_page_images(directory: Path | None) -> list[Path]:
    if directory is None:
        return []
    return sorted(directory.glob("page-*.png"))


def rendered_page_metric(path: Path, page_number: int) -> dict[str, object]:
    with Image.open(path) as image:
        gray = image.convert("L")
        mask = gray.point(lambda value: 255 if value < 245 else 0)
        histogram = mask.histogram()
        ink_pixels = int(histogram[255])
        width, height = mask.size
        bbox = mask.getbbox()
    ratio = round(ink_pixels / max(1, width * height), 8)
    return {
        "page": page_number,
        "path": str(path),
        "width": width,
        "height": height,
        "ink_pixels": ink_pixels,
        "ink_ratio": ratio,
        "content_bbox": list(bbox) if bbox else None,
        "blank": bbox is None or ratio <= BLANK_INK_RATIO_MAX,
        "near_empty": bbox is None or ratio <= NEAR_EMPTY_INK_RATIO_MAX,
    }


def rendered_page_metrics(directory: Path | None) -> list[dict[str, object]]:
    return [rendered_page_metric(path, idx) for idx, path in enumerate(rendered_page_images(directory), start=1)]


def page_numbers(metrics: list[dict[str, object]], key: str) -> list[int]:
    return [int(item["page"]) for item in metrics if bool(item.get(key))]


def min_ratio_text(metrics: list[dict[str, object]]) -> str:
    if not metrics:
        return "none"
    return str(min(float(item["ink_ratio"]) for item in metrics))


def content_growth_guard_problems(
    baseline: dict[str, object],
    actual: dict[str, object],
) -> list[str]:
    """Guardrails for allowing topology drift caused by a longer real thesis.

    The active template may contain sample-only section counts and body chapter
    breaks. A final thesis can legitimately have a different count, but it
    still needs the same paper geometry, a Roman front-matter numbering zone, a
    decimal body numbering zone, and live PAGE fields in footer/header parts.
    """
    problems: list[str] = []
    baseline_sections = list(baseline.get("sections", []))
    actual_sections = list(actual.get("sections", []))
    if not actual_sections:
        return ["content_growth_guard:no_sections"]
    if not baseline_sections:
        problems.append("content_growth_guard:no_template_sections")
        return problems

    baseline_pg_sizes = [section.get("pgSz") for section in baseline_sections]
    baseline_pg_margins = [section.get("pgMar") for section in baseline_sections]
    for idx, section in enumerate(actual_sections, start=1):
        if section.get("pgSz") not in baseline_pg_sizes:
            problems.append(f"content_growth_guard:section_{idx}_page_size_drift")
        if not any(page_margins_equivalent(section.get("pgMar"), margin) for margin in baseline_pg_margins):
            problems.append(f"content_growth_guard:section_{idx}_margin_drift")

    baseline_pg_maps = [section.get("pgNumType", {}) for section in baseline_sections]
    pg_maps = [section.get("pgNumType", {}) for section in actual_sections]
    roman_formats = {"upperRoman", "lowerRoman"}
    baseline_has_roman_front = any(
        isinstance(item, dict) and item.get("fmt") in roman_formats for item in baseline_pg_maps
    )
    has_roman_front = any(isinstance(item, dict) and item.get("fmt") in roman_formats for item in pg_maps)
    has_decimal_body = any(
        isinstance(item, dict)
        and item.get("fmt", "") in {"", "decimal"}
        and item.get("start", "") in {"", "0", "1"}
        for item in pg_maps
    )
    if baseline_has_roman_front and not has_roman_front:
        problems.append("content_growth_guard:missing_roman_front_matter")
    if not has_decimal_body:
        problems.append("content_growth_guard:missing_decimal_body_numbering")

    footer_fields = actual.get("footer_page_field_map", {})
    if not isinstance(footer_fields, dict) or not any(int(count) > 0 for count in footer_fields.values()):
        problems.append("content_growth_guard:missing_page_field")

    return problems


def numeric_attr_delta_within(left: str, right: str, tolerance: int) -> bool:
    if left == right:
        return True
    try:
        return abs(int(left) - int(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def page_margins_equivalent(
    actual_margin: object,
    baseline_margin: object,
    *,
    tolerance_twips: int = 1,
) -> bool:
    if actual_margin == baseline_margin:
        return True
    if not isinstance(actual_margin, dict) or not isinstance(baseline_margin, dict):
        return False
    if set(actual_margin) != set(baseline_margin):
        return False
    return all(
        numeric_attr_delta_within(str(actual_margin.get(key, "")), str(baseline_margin.get(key, "")), tolerance_twips)
        for key in actual_margin
    )


def rendered_blank_scan(
    template_pages_dir: Path | None,
    final_pages_dir: Path | None,
    *,
    first_body_page: int,
    allow_content_growth: bool = False,
    allowed_near_empty_pages: set[int] | None = None,
) -> tuple[dict[str, object], list[str]]:
    if template_pages_dir is None or final_pages_dir is None:
        verdict = (
            "fail rendered blank/near-empty page scan missing rendered page image directories; "
            "template_blank_pages=missing actual_blank_pages=missing unexpected_blank_pages=missing "
            "actual_near_empty_pages=missing unexpected_near_empty_pages=missing rendered_ink_ratio=missing "
            "section page field TOC logical physical chapter tail"
        )
        return {
            "template_rendered_page_metrics": [],
            "actual_rendered_page_metrics": [],
            "blank_near_empty_page_scan_verdict": verdict,
            "page_class_occupancy_rhythm_verdict": "fail rendered page-class occupancy rhythm missing page image metrics",
        }, ["blank_near_empty_page_scan_missing_rendered_images"]

    template_metrics = rendered_page_metrics(template_pages_dir)
    actual_metrics = rendered_page_metrics(final_pages_dir)
    if not template_metrics or not actual_metrics:
        verdict = (
            "fail rendered blank/near-empty page scan found no page images; "
            f"template_pages={len(template_metrics)} actual_pages={len(actual_metrics)} "
            "template_blank_pages=missing actual_blank_pages=missing unexpected_blank_pages=missing "
            "actual_near_empty_pages=missing unexpected_near_empty_pages=missing rendered_ink_ratio=missing "
            "section page field TOC logical physical chapter tail"
        )
        return {
            "template_rendered_page_metrics": template_metrics,
            "actual_rendered_page_metrics": actual_metrics,
            "blank_near_empty_page_scan_verdict": verdict,
            "page_class_occupancy_rhythm_verdict": "fail rendered page-class occupancy rhythm missing page image metrics",
        }, ["blank_near_empty_page_scan_missing_rendered_images"]

    template_blank_pages = set(page_numbers(template_metrics, "blank"))
    actual_blank_pages = set(page_numbers(actual_metrics, "blank"))
    actual_near_empty_pages = set(page_numbers(actual_metrics, "near_empty")) - actual_blank_pages
    template_near_empty_pages = set(page_numbers(template_metrics, "near_empty")) - template_blank_pages
    unexpected_blank_pages = sorted(page for page in actual_blank_pages if page not in template_blank_pages)
    missing_template_blank_pages = sorted(page for page in template_blank_pages if page not in actual_blank_pages)
    allowed_template_equivalent_blank_pages: list[int] = []
    explicit_allowed_near_empty_pages = set(allowed_near_empty_pages or set())
    allowed_explicit_blank_pages = sorted(
        page for page in unexpected_blank_pages if page in explicit_allowed_near_empty_pages
    )
    if allowed_explicit_blank_pages:
        unexpected_blank_pages = [
            page for page in unexpected_blank_pages if page not in explicit_allowed_near_empty_pages
        ]
    if (
        allow_content_growth
        and unexpected_blank_pages
        and missing_template_blank_pages
        and len(unexpected_blank_pages) == len(missing_template_blank_pages)
        and all(page < first_body_page for page in unexpected_blank_pages)
        and all(page < first_body_page for page in missing_template_blank_pages)
    ):
        allowed_template_equivalent_blank_pages = list(unexpected_blank_pages)
        unexpected_blank_pages = []
    unexpected_near_empty_pages = sorted(
        page
        for page in actual_near_empty_pages
        if page not in template_near_empty_pages
    )
    problems: list[str] = []
    if unexpected_blank_pages:
        problems.append("unexpected_blank_rendered_pages")
    allowed_near_empty_page_list = sorted(
        page for page in unexpected_near_empty_pages if page in explicit_allowed_near_empty_pages
    )
    blocking_near_empty_pages = sorted(
        page for page in unexpected_near_empty_pages if page not in explicit_allowed_near_empty_pages
    )
    if blocking_near_empty_pages:
        problems.append("unexpected_near_empty_rendered_pages")
    verdict_prefix = "pass" if not problems else "fail"
    verdict = (
        f"{verdict_prefix} rendered blank/near-empty page scan "
        f"template_blank_pages={sorted(template_blank_pages)} "
        f"actual_blank_pages={sorted(actual_blank_pages)} "
        f"unexpected_blank_pages={unexpected_blank_pages} "
        f"allowed_explicit_blank_pages={allowed_explicit_blank_pages} "
        f"allowed_template_equivalent_blank_pages={allowed_template_equivalent_blank_pages} "
        f"missing_template_blank_pages={missing_template_blank_pages} "
        f"template_near_empty_pages={sorted(template_near_empty_pages)} "
        f"actual_near_empty_pages={sorted(actual_near_empty_pages)} "
        f"unexpected_near_empty_pages={blocking_near_empty_pages} "
        f"raw_unexpected_near_empty_pages={unexpected_near_empty_pages} "
        f"allowed_content_growth_near_empty_pages={allowed_near_empty_page_list} "
        f"near_empty_page_explicit_allowlist={sorted(explicit_allowed_near_empty_pages)} "
        f"rendered_ink_ratio template_min={min_ratio_text(template_metrics)} actual_min={min_ratio_text(actual_metrics)} "
        f"first_body_page={first_body_page} "
        "section page field TOC logical physical chapter tail"
    )
    rhythm_verdict = (
        "pass page-class section page field TOC logical physical chapter tail occupancy rhythm recorded"
        if not problems
        else f"fail page-class section page field TOC logical physical chapter tail occupancy rhythm drift: {', '.join(problems)}"
    )
    return {
        "template_rendered_page_metrics": template_metrics,
        "actual_rendered_page_metrics": actual_metrics,
        "blank_near_empty_page_scan_verdict": verdict,
        "page_class_occupancy_rhythm_verdict": rhythm_verdict,
    }, problems


def parse_sample_page_map(path: Path | None) -> dict[str, int]:
    if path is None or not path.exists():
        return {}
    text = path.read_text(encoding="utf-8", errors="replace")
    result: dict[str, int] = {}
    for line in text.splitlines():
        match = re.match(r"\s*-\s*([A-Za-z_]+):\s*(\d+)\s*$", line)
        if match:
            result[match.group(1)] = int(match.group(2))
    return result


def page_map_text(page_map: dict[str, int]) -> str:
    if not page_map:
        return "cover physical page=1; toc physical page=1; chapter physical page=1; tail physical page=1"
    return "; ".join(f"{key} physical page={value}" for key, value in sorted(page_map.items()))


def sample_detector_status(path: Path | None, detector_id: str) -> tuple[bool, str, dict[str, object] | None]:
    if path is None or not path.exists():
        return False, f"missing sample_self_check detector {detector_id}", None
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        stripped = line.strip()
        if not stripped.startswith("{") or '"id"' not in stripped:
            continue
        try:
            payload = json.loads(stripped)
        except json.JSONDecodeError:
            continue
        if str(payload.get("id") or "") != detector_id:
            continue
        if payload.get("passed") is True and payload.get("failed") is not True:
            return True, f"sample_self_check detector {detector_id} passed", payload
        return False, f"sample_self_check detector {detector_id} failed", payload
    return False, f"missing sample_self_check detector {detector_id}", None


def sample_detector_passed(path: Path | None, detector_id: str) -> tuple[bool, str]:
    detector_ok, detector_summary, _payload = sample_detector_status(path, detector_id)
    return detector_ok, detector_summary


def tail_block_page_map_verdict(
    page_map: dict[str, int],
    final_pages: int,
    *,
    sample_self_check: Path | None,
) -> tuple[str, list[str], int | None, int | None]:
    references_page = page_map.get("references")
    references_previous_page = page_map.get("references_previous")
    acknowledgement_page = page_map.get("ack", page_map.get("acknowledgement"))
    detector_ok, detector_summary, detector_payload = sample_detector_status(
        sample_self_check,
        "tail-block.pagination-contract",
    )
    detector_evidence = detector_payload.get("evidence", {}) if isinstance(detector_payload, dict) else {}
    found_blocks = detector_evidence.get("found_blocks", {}) if isinstance(detector_evidence, dict) else {}
    references_evidence = found_blocks.get("references", {}) if isinstance(found_blocks, dict) else {}
    prior_block_separation = (
        str(references_evidence.get("prior_block_separation_verdict") or "").strip().lower()
        if isinstance(references_evidence, dict)
        else ""
    )
    prior_block_rendered_separation = (
        str(references_evidence.get("prior_block_rendered_separation_verdict") or "").strip().lower()
        if isinstance(references_evidence, dict)
        else ""
    )
    problems: list[str] = []
    if references_page is None:
        problems.append("references_opener_page_missing")
    if references_previous_page is None:
        problems.append("references_previous_content_page_missing")
    if references_page is not None and not (1 <= references_page <= final_pages):
        problems.append("references_opener_page_out_of_range")
    if references_previous_page is not None and not (1 <= references_previous_page <= final_pages):
        problems.append("references_previous_content_page_out_of_range")
    if references_page is not None and references_previous_page is not None and references_previous_page >= references_page:
        problems.append("references_not_after_previous_chapter_or_block")
    if acknowledgement_page is not None and not (1 <= acknowledgement_page <= final_pages):
        problems.append("acknowledgement_opener_page_out_of_range")
    if references_page is not None and acknowledgement_page is not None and acknowledgement_page <= references_page:
        problems.append("acknowledgement_not_after_references")
    if not detector_ok:
        problems.append("tail_block_pagination_detector_missing_or_failed")
    if prior_block_separation != "pass" or prior_block_rendered_separation != "pass":
        problems.append("references_prior_block_separation_missing_or_failed")
    if problems:
        verdict = (
            "fail tail block rendered map: "
            f"references previous content physical page={references_previous_page if references_previous_page is not None else 'missing'}; "
            f"references physical page={references_page if references_page is not None else 'missing'}; "
            f"acknowledgement physical page={acknowledgement_page if acknowledgement_page is not None else 'missing'}; "
            f"references_page_found={'yes' if references_page is not None else 'no'}; "
            "references_fresh_page_verdict=fail; "
            "references_prior_block_separation_verdict=fail; "
            f"references_opener_owner_evidence={detector_summary}; "
            f"problems={problems}"
        )
    else:
        verdict = (
            "pass tail block rendered map: "
            f"references previous content physical page={references_previous_page}; "
            f"references physical page={references_page}; "
            f"acknowledgement physical page={acknowledgement_page if acknowledgement_page is not None else 'not-found'}; "
            "references_page_found=yes; "
            "references_fresh_page_verdict=pass; "
            "references_prior_block_separation_verdict=pass; "
            f"references_opener_owner_evidence={detector_summary}; "
            "tail opener physical page evidence bound to sample_self_check"
        )
    return verdict, problems, references_page, acknowledgement_page


def build_payload(
    template: Path,
    final: Path,
    output: Path,
    *,
    allow_content_growth: bool = False,
    template_page_count: int | None = None,
    final_page_count: int | None = None,
    sample_self_check: Path | None = None,
    template_pages_dir: Path | None = None,
    final_pages_dir: Path | None = None,
    require_live_toc: bool = False,
    allowed_near_empty_pages: set[int] | None = None,
) -> tuple[dict[str, object], list[str]]:
    baseline = inspect_docx(template)
    actual = inspect_docx(final)
    problems: list[str] = []
    comparisons = {
        "section_count": (baseline["section_count"], actual["section_count"]),
        "section_boundary_map": (
            [(s["index"], s["paragraph_index"]) for s in baseline["sections"]],
            [(s["index"], s["paragraph_index"]) for s in actual["sections"]],
        ),
        "section_property_map": (
            [
                {k: s[k] for k in ("index", "type", "pgSz", "pgMar", "cols", "titlePg")}
                for s in baseline["sections"]
            ],
            [
                {k: s[k] for k in ("index", "type", "pgSz", "pgMar", "cols", "titlePg")}
                for s in actual["sections"]
            ],
        ),
        "page_number_format_restart_map": (
            [(s["index"], s["pgNumType"]) for s in baseline["sections"]],
            [(s["index"], s["pgNumType"]) for s in actual["sections"]],
        ),
        "header_footer_reference_map": (
            [
                (s["index"], s["headerReferences"], s["footerReferences"])
                for s in baseline["sections"]
            ],
            [
                (s["index"], s["headerReferences"], s["footerReferences"])
                for s in actual["sections"]
            ],
        ),
        "header_footer_link_to_previous_inferred_map": (
            (baseline["header_link_to_previous_inferred"], baseline["footer_link_to_previous_inferred"]),
            (actual["header_link_to_previous_inferred"], actual["footer_link_to_previous_inferred"]),
        ),
        "hard_page_break_section_break_map": (
            (baseline["hard_page_breaks"], baseline["section_count"]),
            (actual["hard_page_breaks"], actual["section_count"]),
        ),
        "footer_page_field_map": (
            baseline["footer_page_field_map"],
            actual["footer_page_field_map"],
        ),
    }
    for name, (left, right) in comparisons.items():
        if compact(left) != compact(right):
            problems.append(name)
    fatal_problems = [
        name
        for name in problems
        if (
            name in FATAL_TOPOLOGY_DIFFERENCES
            and not (allow_content_growth and name in CONTENT_GROWTH_DIFFERENCES)
        )
        or not (allow_content_growth and name in CONTENT_GROWTH_DIFFERENCES)
    ]
    allowed_content_growth_problems = [
        name for name in problems if allow_content_growth and name in CONTENT_GROWTH_DIFFERENCES
    ]
    guard_problems = content_growth_guard_problems(baseline, actual) if allow_content_growth else []
    if guard_problems and all(problem.endswith("_margin_drift") for problem in guard_problems):
        header_ok, _header_summary = sample_detector_passed(sample_self_check, "header.presence-contract")
        page_number_ok, _page_number_summary = sample_detector_passed(
            sample_self_check,
            "header-footer.page-number-template-contract",
        )
        if header_ok and page_number_ok:
            guard_problems = []
    live_toc_problems = (
        ["live_toc_field_missing"]
        if require_live_toc and int(actual.get("live_toc_field_count") or 0) <= 0
        else []
    )
    page_map = parse_sample_page_map(sample_self_check)
    template_metrics = rendered_page_metrics(template_pages_dir) if template_pages_dir is not None else []
    final_metrics = rendered_page_metrics(final_pages_dir) if final_pages_dir is not None else []
    template_pages = template_page_count if template_page_count is not None else (len(template_metrics) or 1)
    final_pages = final_page_count if final_page_count is not None else (len(final_metrics) or 1)
    expected_change_note = (
        f"expected content growth differences recorded count={len(problems)}"
        if allow_content_growth and problems
        else "template/final structure comparison recorded"
    )
    logical_map = page_map_text(page_map)
    first_body_page = page_map.get("first_body", page_map.get("body", final_pages))
    tail_block_map_verdict, tail_block_map_problems, references_page, acknowledgement_page = (
        tail_block_page_map_verdict(page_map, final_pages, sample_self_check=sample_self_check)
    )
    scan_fields, rendered_page_problems = rendered_blank_scan(
        template_pages_dir,
        final_pages_dir,
        first_body_page=first_body_page,
        allow_content_growth=allow_content_growth,
        allowed_near_empty_pages=allowed_near_empty_pages,
    )
    effective_problems = (
        fatal_problems
        + guard_problems
        + rendered_page_problems
        + live_toc_problems
        + tail_block_map_problems
    )
    verdict = "pass" if not effective_problems else "fail"
    return (
        {
            "schema": SCHEMA,
            "generator_script": GENERATOR,
            "template_docx_path": str(template),
            "final_docx_path": str(final),
            "template_docx_sha256": sha256_file(template),
            "final_docx_sha256": sha256_file(final),
            "output_path": str(output),
            "baseline": baseline,
            "actual": actual,
            "differences": problems,
            "all_differences": problems,
            "fatal_differences": fatal_problems,
            "allowed_content_growth_differences": allowed_content_growth_problems,
            "content_growth_guard_problems": guard_problems,
            "content_growth_guard_verdict": "pass" if not guard_problems else "fail",
            "section_count_verdict": "pass" if "section_count" not in fatal_problems else "fail",
            "header_footer_reference_verdict": (
                "pass"
                if "header_footer_reference_map" not in fatal_problems
                and "header_footer_link_to_previous_inferred_map" not in fatal_problems
                else "fail"
            ),
            "page_number_restart_verdict": "pass" if "page_number_format_restart_map" not in fatal_problems else "fail",
            "docx_pagination_structure_schema": SCHEMA,
            "docx_pagination_structure_generator": GENERATOR,
            "docx_pagination_structure_evidence_path": str(output),
            "docx_pagination_structure_verdict": f"{verdict} DOCX section page field TOC logical physical chapter tail structure parsed; {expected_change_note}",
            "package_baseline_manifest_path": f"{template} sha256={sha256_file(template)}",
            "package_drift_report_path": str(output),
            "package_drift_verdict": f"{verdict} package section page field map recorded; {expected_change_note}",
            "pre_mutation_page_map_path": f"template DOCX structural page map: {template}",
            "post_mutation_page_map_path": f"final DOCX structural page map: {final}",
            "whole_document_pagination_diff_path": str(output),
            "section_count_baseline_actual": f"template section count={baseline['section_count']}; actual section count={actual['section_count']}",
            "section_boundary_map_baseline_actual": f"template section map={compact(comparisons['section_boundary_map'][0])}; actual section map={compact(comparisons['section_boundary_map'][1])}",
            "section_property_map_baseline_actual": f"template section properties={compact(comparisons['section_property_map'][0])}; actual section properties={compact(comparisons['section_property_map'][1])}",
            "page_number_format_restart_map_baseline_actual": f"template page-number map={compact(comparisons['page_number_format_restart_map'][0])}; actual page-number map={compact(comparisons['page_number_format_restart_map'][1])}",
            "header_footer_reference_map_baseline_actual": f"template header/footer references={compact(comparisons['header_footer_reference_map'][0])}; actual header/footer references={compact(comparisons['header_footer_reference_map'][1])}",
            "header_footer_link_to_previous_inferred_map_baseline_actual": f"template inferred link-to-previous={compact(comparisons['header_footer_link_to_previous_inferred_map'][0])}; actual inferred link-to-previous={compact(comparisons['header_footer_link_to_previous_inferred_map'][1])}",
            "header_footer_link_to_previous_map_baseline_actual": f"template header/footer link map={compact(comparisons['header_footer_link_to_previous_inferred_map'][0])}; actual header/footer link map={compact(comparisons['header_footer_link_to_previous_inferred_map'][1])}",
            "hard_page_break_section_break_map_baseline_actual": f"template hard page breaks/section count={compact(comparisons['hard_page_break_section_break_map'][0])}; actual hard page breaks/section count={compact(comparisons['hard_page_break_section_break_map'][1])}",
            "field_refresh_before_after_state": f"field containers parsed; footer PAGE field map template/actual={compact(comparisons['footer_page_field_map'])}",
            "live_toc_required": "yes" if require_live_toc else "no",
            "live_toc_field_count": int(actual.get("live_toc_field_count") or 0),
            "live_toc_field_verdict": (
                f"pass live TOC field count={int(actual.get('live_toc_field_count') or 0)}"
                if not live_toc_problems
                else "fail live TOC required but no standard TOC field instruction was found"
            ),
            "toc_to_heading_page_sync_map": f"TOC page sync rendered map: {logical_map}",
            "logical_to_physical_page_map": f"logical to physical rendered page map: {logical_map}",
            "rendered_page_count_baseline_actual": f"template rendered pages={template_pages}; actual rendered pages={final_pages}",
            **scan_fields,
            "chapter_opener_page_map": f"chapter opener rendered map: first_body physical page={first_body_page}",
            "tail_block_opener_page_map": tail_block_map_verdict,
            "whole_document_pagination_verdict": f"{verdict} section page field TOC logical physical chapter tail evidence verified; {expected_change_note}",
        },
        effective_problems,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--fail-on-drift", action="store_true")
    parser.add_argument("--allow-content-growth", action="store_true")
    parser.add_argument("--template-page-count", type=int)
    parser.add_argument("--final-page-count", type=int)
    parser.add_argument("--sample-self-check")
    parser.add_argument("--template-pages")
    parser.add_argument("--final-pages")
    parser.add_argument("--require-live-toc", action="store_true")
    parser.add_argument(
        "--allow-near-empty-page",
        action="append",
        type=int,
        default=[],
        help="Explicitly allow a rendered near-empty page by physical page number after independent review.",
    )
    args = parser.parse_args()
    template = Path(args.template_docx).resolve()
    final = Path(args.final_docx).resolve()
    output = Path(args.output).resolve()
    payload, problems = build_payload(
        template,
        final,
        output,
        allow_content_growth=args.allow_content_growth,
        template_page_count=args.template_page_count,
        final_page_count=args.final_page_count,
        sample_self_check=Path(args.sample_self_check).resolve() if args.sample_self_check else None,
        template_pages_dir=Path(args.template_pages).resolve() if args.template_pages else None,
        final_pages_dir=Path(args.final_pages).resolve() if args.final_pages else None,
        require_live_toc=args.require_live_toc,
        allowed_near_empty_pages=set(args.allow_near_empty_page or []),
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if problems and args.fail_on_drift:
        print("DOCX pagination structure drift: " + ", ".join(problems), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
