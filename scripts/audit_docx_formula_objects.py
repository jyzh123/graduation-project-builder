from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}

FORMULA_WORD = "\u5f0f"
FORMULA_NUMBER_CORE = r"\d+(?:[-.]\d+[A-Za-z]?)?"
FORMULA_NUMBER_TOKEN_RE = re.compile(rf"(?:\u5f0f\s*)?[\(\uff08]\s*(?:\u5f0f\s*)?{FORMULA_NUMBER_CORE}\s*[\)\uff09]")
FORMULA_NUMBER_CANDIDATE_CELL_RE = re.compile(
    rf"^(?:\u5f0f\s*)?[\(\uff08]\s*{FORMULA_NUMBER_CORE}\s*[\)\uff09]$"
)
FORMULA_NUMBER_CELL_RE = re.compile(r"^\u5f0f\(\d+-\d+\)$")
STRICT_FORMULA_NUMBER_TOKEN_RE = re.compile(r"\u5f0f\(\d+-\d+\)")
BARE_RENDERED_FORMULA_LABEL_RE = re.compile(r"^[\(\uff08]\s*\d{1,2}\s*[-.]\s*\d{1,3}[A-Za-z]?\s*[\)\uff09]$")
RAW_MATH_COMMAND_TOKEN_RE = re.compile(
    r"\b(?:sub|sup|frac|sqrt|over|below|above|nary|lim|eqarr)\b",
    re.I,
)
ASSIGNMENT_RE = re.compile(r"[A-Za-z][A-Za-z0-9_{}]*\s*=")
OPERATOR_RE = re.compile(r"(\+|-|\*|/|×|=|\bmin\s*\(|\bmax\s*\(|\bsqrt\s*\(|\bsum\b)", re.I)
MATH_SIGNAL_RE = re.compile(r"(=|\uff1d|[<>≤≥]|\+|-|−|\*|/|×|÷|√|π|sin|cos|tan|σ|ε|η|Δ|△)", re.I)
MATH_SYMBOL_RE = re.compile(r"[A-Za-z\u0370-\u03ff]")
PROSE_FORMULA_SKIP_RE = re.compile(
    r"^(?:\u67e5\u8868|\u67e5\u56fe|\u56e0\u4e3a|\u6240\u4ee5|\u578b\u53f7|\u957f\u5ea6|\u5f0f\u4e2d|\u5176\u4e2d|"
    r"[A-Za-zΑ-Ωα-ω]\s*[\u2014\u2013-]{2})"
)
TEXT_FORMULA_TOKEN_RE = re.compile(
    r"(?:_[A-Za-z]|\\bC_sink\\b|\\bf_sink\\b|\\bf_org\\b|\\bf_crop\\b|\\bf_soil\\b|"
    r"\\bf_climate\\b|\\bf_r\\b|\\bf_t\\b|\\bC\s*sink\\b|\\bf\s*sink\\b)",
    re.I,
)


OPERATOR_RE = re.compile(
    r"(\+|-|\*|/|×|脳|(?<=\s)x(?=\s)|=|\bmin\s*\(|\bmax\s*\(|\bsqrt\s*\(|\bsum\b)",
    re.I,
)
TEXT_FORMULA_TOKEN_RE = re.compile(
    r"(?:_[A-Za-z]|\bC_sink\b|\bf_sink\b|\bf_org\b|\bf_crop\b|\bf_soil\b|"
    r"\bf_climate\b|\bf_r\b|\bf_t\b|\bC\s*sink\b|\bf\s*sink\b)",
    re.I,
)
CODE_LINE_RE = re.compile(r"^\s*\d+\s*\|\s*")
CODE_TOKEN_RE = re.compile(
    r"(@app\.route|def\s+\w+\(|class\s+\w+|return\b|request\.|jsonify\(|"
    r"os\.path|cv2\.|self\.|import\s+\w+|from\s+\w+\s+import|"
    r"\bconst\s+\w+\s*=|\blet\s+\w+\s*=|\bvar\s+\w+\s*=|\bawait\s+fetch\b|"
    r"\bif\s+.+:|\bfor\s+.+:|\btry:|\bexcept\b|=>|</?\w+)",
    re.I,
)
STATUS_ASSIGNMENT_RE = re.compile(r"^[\u4e00-\u9fffA-Za-z\s]*(?:status|code|message|url|id|api|path)\s*=\s*[\w./:-]+$", re.I)
BODY_START_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十]+章|[1-9]\d?(?:\s|[\.．、])|绪论|引言|设计计算|结构设计)"
)
TAIL_HEADING_RE = re.compile(
    r"^\s*(?:参考文献|致谢|附录|References|Acknowledgements?|Appendix)\s*$",
    re.I,
)
APPENDIX_HEADING_RE = re.compile(r"^\s*(?:附录|Appendix)\b", re.I)
BODY_START_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百零\d]+章|[1-9]\s+[\u4e00-\u9fffA-Za-z]|绪论|引言|设计计算|结构设计)"
)
TAIL_HEADING_RE = re.compile(
    r"^\s*(?:参考文献|致谢|附录|References|Acknowledgements?|Appendix)\b",
    re.I,
)
APPENDIX_HEADING_RE = re.compile(r"^\s*(?:附录|Appendix)\b", re.I)
FORMULA_LEADIN_RE = re.compile(
    r"(?:\u4e0b\u5f0f|\u516c\u5f0f|\u8ba1\u7b97\u5f0f|\u8868\u8fbe\u5f0f|\u8868\u793a\u4e3a|\u5b9a\u4e49\u4e3a|\u53ef\u7528.*\u8868\u793a).*[\uff1a:]$"
)
FORMULA_EXPLANATION_START_RE = re.compile(r"^\s*(?:\u5176\u4e2d|\u5f0f\u4e2d|where\b)", re.I)
INEQUALITY_RE = re.compile(
    r"(?:[A-Za-z][A-Za-z0-9_{}]*\s*[<>≤≥]\s*\d|\d+(?:\.\d+)?\s*[<>≤≥]\s*[A-Za-z][A-Za-z0-9_{}]*)"
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def paragraph_math_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//m:t", NS))


def paragraph_style(paragraph: ET.Element) -> str:
    pstyle = paragraph.find(".//w:pPr/w:pStyle", NS)
    if pstyle is None:
        return ""
    return str(pstyle.get(f"{{{W}}}val") or "")


def has_math(paragraph: ET.Element) -> bool:
    return bool(
        paragraph.findall(".//m:oMath", NS)
        or paragraph.findall(".//m:oMathPara", NS)
        or paragraph.findall(".//w:object", NS)
    )


def math_count(paragraph: ET.Element) -> int:
    return len(paragraph.findall(".//m:oMath", NS)) + len(paragraph.findall(".//w:object", NS))


def omml_math_count(paragraph: ET.Element) -> int:
    return len(paragraph.findall(".//m:oMath", NS))


def legacy_formula_object_count(paragraph: ET.Element) -> int:
    return len(paragraph.findall(".//w:object", NS))


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def run_size(run: ET.Element) -> str:
    node = run.find("./w:rPr/w:sz", NS)
    return str(node.get(f"{{{W}}}val") or "") if node is not None else ""


def w_attr(node: ET.Element | None, name: str) -> str:
    return str(node.get(f"{{{W}}}{name}") or "") if node is not None else ""


def int_w_attr(node: ET.Element | None, name: str) -> int | None:
    try:
        return int(w_attr(node, name))
    except (TypeError, ValueError):
        return None


def cell_text(cell: ET.Element) -> str:
    return "".join(node.text or "" for node in cell.findall(".//w:t", NS))


def nonempty_cell_paragraph_texts(cell: ET.Element) -> list[str]:
    texts: list[str] = []
    for paragraph in cell.findall("./w:p", NS):
        text = paragraph_text(paragraph).strip()
        if text:
            texts.append(text)
    return texts


def cell_text_run_sizes(cell: ET.Element) -> list[str]:
    sizes: list[str] = []
    for run in cell.findall(".//w:r", NS):
        if run_text(run).strip():
            sizes.append(run_size(run))
    return sizes


def cell_has_no_wrap(cell: ET.Element) -> bool:
    return cell.find("./w:tcPr/w:noWrap", NS) is not None


def usable_width_twips(root: ET.Element) -> int | None:
    sect_pr = root.find(".//w:body/w:sectPr", NS)
    if sect_pr is None:
        return None
    pg_sz = sect_pr.find("./w:pgSz", NS)
    pg_mar = sect_pr.find("./w:pgMar", NS)
    page_width = int_w_attr(pg_sz, "w")
    left = int_w_attr(pg_mar, "left")
    right = int_w_attr(pg_mar, "right")
    if page_width is None or left is None or right is None:
        return None
    usable = page_width - left - right
    return usable if usable > 0 else None


def paragraph_tab_stops(paragraph: ET.Element) -> list[dict[str, int | str]]:
    stops: list[dict[str, int | str]] = []
    for tab in paragraph.findall("./w:pPr/w:tabs/w:tab", NS):
        pos = int_w_attr(tab, "pos")
        if pos is None:
            continue
        stops.append({"val": w_attr(tab, "val"), "pos": pos, "leader": w_attr(tab, "leader")})
    return stops


def paragraph_direct_tab_count(paragraph: ET.Element) -> int:
    return len(paragraph.findall("./w:r/w:tab", NS))


def paragraph_inside_table(paragraph: ET.Element) -> bool:
    return any(getattr(parent, "tag", "") == f"{{{W}}}tbl" for parent in paragraph.iterancestors())


def paragraph_ancestor_table(paragraph: ET.Element) -> ET.Element | None:
    for parent in paragraph.iterancestors():
        if getattr(parent, "tag", "") == f"{{{W}}}tbl":
            return parent
    return None


def paragraph_ancestor_row(paragraph: ET.Element) -> ET.Element | None:
    for parent in paragraph.iterancestors():
        if getattr(parent, "tag", "") == f"{{{W}}}tr":
            return parent
    return None


def row_has_formula_number_cell(row: ET.Element | None) -> bool:
    if row is None:
        return False
    for cell in row.findall("./w:tc", NS):
        compact = re.sub(r"\s+", "", cell_text(cell)).strip()
        if FORMULA_NUMBER_CANDIDATE_CELL_RE.fullmatch(compact):
            return True
    return False


def paragraph_math_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//m:t", NS)).strip()


def is_unit_like_math_text(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if not compact or len(compact) > 12:
        return False
    if re.search(r"[\d=<>鈮も墺+\-*/]", compact):
        return False
    if re.search(r"[\u4e00-\u9fff]", compact):
        return False
    return re.fullmatch(r"[A-Za-zμµα-ωΑ-Ω·⋅./()%]+", compact) is not None


def is_non_formula_table_unit_math(paragraph: ET.Element, *, formula_like: bool) -> bool:
    if formula_like or not paragraph_inside_table(paragraph):
        return False
    if row_has_formula_number_cell(paragraph_ancestor_row(paragraph)):
        return False
    return is_unit_like_math_text(paragraph_math_text(paragraph))


def pdf_line_text(line: dict[str, object]) -> str:
    spans = line.get("spans", [])
    if not isinstance(spans, list):
        return ""
    return "".join(str(span.get("text", "")) for span in spans if isinstance(span, dict))


def rounded_span_size(span: dict[str, object]) -> str:
    try:
        return f"{float(span.get('size', 0.0)):.2f}"
    except (TypeError, ValueError):
        return "0.00"


def is_rendered_formula_label_candidate(text: str) -> bool:
    """Return true for rendered formula-number labels, not formula-body values.

    PDF extraction often returns parenthesized numeric operands from equations as
    standalone text runs.  Prefixless rendered candidates are therefore limited
    to short chapter-sequence shapes such as ``(2-16)`` or ``(6.1)``; long
    decimal or integer operands such as ``(609147.9)`` are equation content.
    """

    compact = re.sub(r"\s+", "", text or "")
    if FORMULA_NUMBER_CELL_RE.fullmatch(compact):
        return True
    return BARE_RENDERED_FORMULA_LABEL_RE.fullmatch(compact) is not None


def line_bbox(line: dict[str, object]) -> tuple[float, float, float, float]:
    bbox = line.get("bbox", [0, 0, 0, 0])
    if not isinstance(bbox, (list, tuple)) or len(bbox) != 4:
        return (0.0, 0.0, 0.0, 0.0)
    try:
        return tuple(float(value) for value in bbox)  # type: ignore[return-value]
    except (TypeError, ValueError):
        return (0.0, 0.0, 0.0, 0.0)


def formula_labels_near_paragraph(paragraph: ET.Element) -> list[str]:
    labels = STRICT_FORMULA_NUMBER_TOKEN_RE.findall(paragraph_text(paragraph))
    row = paragraph_ancestor_row(paragraph)
    if row is not None:
        labels.extend(STRICT_FORMULA_NUMBER_TOKEN_RE.findall(cell_text(row)))
    return sorted(set(labels))


def audit_raw_math_command_tokens(root: ET.Element) -> dict[str, object]:
    issues: list[dict[str, object]] = []
    paragraphs = root.findall(".//w:body//w:p", NS)
    for index, paragraph in enumerate(paragraphs):
        if not has_math(paragraph):
            continue
        math_text = paragraph_math_text(paragraph)
        visible_text = paragraph_text(paragraph)
        raw_tokens = RAW_MATH_COMMAND_TOKEN_RE.findall(math_text)
        if not raw_tokens:
            continue
        issues.append(
            {
                "issue": "raw-math-command-token-visible-in-formula",
                "paragraph_index": index,
                "labels": formula_labels_near_paragraph(paragraph),
                "raw_tokens": raw_tokens[:30],
                "text": visible_text.strip()[:240],
                "math_text": math_text.strip()[:360],
                "detail": (
                    "OMML formula text contains raw command words such as sub/frac/sup/sqrt; "
                    "these are visible pseudo-formula tokens, not accepted academic formulas"
                ),
            }
        )
    return {
        "raw_math_command_token_count": len(issues),
        "raw_math_command_token_issues": issues,
    }


def audit_rendered_formula_raw_tokens(pdf_path: Path | None) -> dict[str, object]:
    if pdf_path is None:
        return {
            "rendered_raw_math_token_page_count": 0,
            "rendered_raw_math_token_issues": [],
            "rendered_near_empty_formula_page_count": 0,
            "rendered_near_empty_formula_pages": [],
        }
    try:
        import fitz  # type: ignore
    except Exception as exc:  # pragma: no cover - environment dependent
        return {
            "rendered_raw_math_token_page_count": 1,
            "rendered_raw_math_token_issues": [
                {"issue": "pymupdf-unavailable-for-raw-formula-token-audit", "detail": str(exc)}
            ],
            "rendered_near_empty_formula_page_count": 0,
            "rendered_near_empty_formula_pages": [],
        }

    raw_pages: list[dict[str, object]] = []
    near_empty_pages: list[dict[str, object]] = []
    with fitz.open(pdf_path) as doc:
        for page_number, page in enumerate(doc, start=1):
            text = page.get_text("text") or ""
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            raw_tokens = RAW_MATH_COMMAND_TOKEN_RE.findall(text)
            labels = STRICT_FORMULA_NUMBER_TOKEN_RE.findall(text)
            drawings = page.get_drawings()
            images = page.get_images(full=True)
            page_area = max(float(page.rect.width * page.rect.height), 1.0)
            drawing_area = 0.0
            for drawing in drawings:
                rect = drawing.get("rect")
                if rect is None:
                    continue
                drawing_area += max(0.0, float(rect.width)) * max(0.0, float(rect.height))
            drawing_area_ratio = drawing_area / page_area
            has_dense_visual_content = bool(
                len(images) > 0 or (len(drawings) >= 60 and drawing_area_ratio >= 0.05)
            )
            if raw_tokens:
                raw_pages.append(
                    {
                        "issue": "rendered-raw-math-command-token",
                        "page": page_number,
                        "raw_tokens": raw_tokens[:30],
                        "labels": labels[:20],
                        "line_count": len(lines),
                        "char_count": len(text.strip()),
                        "sample_lines": lines[:24],
                    }
                )
            if labels and len(lines) <= 8 and len(text.strip()) < 240 and not has_dense_visual_content:
                near_empty_pages.append(
                    {
                        "issue": "rendered-near-empty-formula-only-page",
                        "page": page_number,
                        "labels": labels[:20],
                        "line_count": len(lines),
                        "char_count": len(text.strip()),
                        "image_count": len(images),
                        "drawing_count": len(drawings),
                        "drawing_area_ratio": round(drawing_area_ratio, 4),
                        "sample_lines": lines[:24],
                        "detail": (
                            "formula-numbered page contains too little surrounding body content; "
                            "single-formula/near-empty pages caused by formula insertion are not accepted"
                        ),
                    }
                )
    return {
        "rendered_raw_math_token_page_count": len(raw_pages),
        "rendered_raw_math_token_issues": raw_pages,
        "rendered_near_empty_formula_page_count": len(near_empty_pages),
        "rendered_near_empty_formula_pages": near_empty_pages,
    }


def audit_rendered_formula_labels(pdf_path: Path | None) -> dict[str, object]:
    if pdf_path is None:
        return {
            "rendered_pdf_path": None,
            "rendered_pdf_sha256": None,
            "rendered_formula_label_split_pair_count": 0,
            "rendered_formula_label_issue_count": 0,
            "rendered_formula_number_unique_sizes": [],
            "rendered_formula_label_issues": [],
        }
    try:
        import fitz  # type: ignore
    except Exception as exc:  # pragma: no cover - environment dependent
        return {
            "rendered_pdf_path": str(pdf_path),
            "rendered_pdf_sha256": sha256(pdf_path) if pdf_path.exists() else None,
            "rendered_formula_label_split_pair_count": 0,
            "rendered_formula_label_issue_count": 1,
            "rendered_formula_number_unique_sizes": [],
            "rendered_formula_label_issues": [
                {"issue": "pymupdf-unavailable", "detail": str(exc)}
            ],
        }

    issues: list[dict[str, object]] = []
    split_pairs: list[dict[str, object]] = []
    label_sizes: list[str] = []
    label_lines: list[dict[str, object]] = []
    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc, start=1):
            page_width = float(page.rect.width)
            text_dict = page.get_text("dict")
            raw_lines: list[dict[str, object]] = []
            for block in text_dict.get("blocks", []):
                if not isinstance(block, dict):
                    continue
                for line in block.get("lines", []):
                    if isinstance(line, dict):
                        raw_lines.append(line)
            raw_lines.sort(key=lambda item: (line_bbox(item)[1], line_bbox(item)[0]))
            visible_lines: list[dict[str, object]] = []
            for line in raw_lines:
                text = re.sub(r"\s+", "", pdf_line_text(line))
                if not text:
                    continue
                bbox = line_bbox(line)
                visible_lines.append({"line": line, "text": text, "bbox": bbox})
                is_formula_label_candidate = is_rendered_formula_label_candidate(text)
                if is_formula_label_candidate:
                    right_aligned = bbox[2] >= page_width * 0.80
                    if not right_aligned:
                        issues.append(
                            {
                                "issue": "rendered-formula-label-not-right-aligned",
                                "detail": "visible formula label must sit at the right-side numbering surface, not near the equation body",
                                "page": page_index,
                                "text": text,
                                "bbox": [round(value, 2) for value in bbox],
                                "page_width": round(page_width, 2),
                                "minimum_right_edge_x": round(page_width * 0.80, 2),
                            }
                        )
                    if FORMULA_NUMBER_CELL_RE.fullmatch(text) is None:
                        issues.append(
                            {
                                "issue": "rendered-formula-label-style-invalid",
                                "detail": "visible formula label must use exact style 式(6-1): formula word prefix, ASCII parentheses, and hyphen chapter numbering",
                                "page": page_index,
                                "text": text,
                                "bbox": [round(value, 2) for value in bbox],
                            }
                        )
                    for span in line.get("spans", []):
                        if not isinstance(span, dict):
                            continue
                        span_text = re.sub(r"\s+", "", str(span.get("text", "")))
                        if FORMULA_NUMBER_CELL_RE.fullmatch(span_text):
                            label_sizes.append(rounded_span_size(span))
                    label_lines.append(
                        {
                            "page": page_index,
                            "text": text,
                            "bbox": [round(value, 2) for value in bbox],
                        }
                    )
            for index, current in enumerate(visible_lines[:-1]):
                current_text = str(current["text"])
                next_text = str(visible_lines[index + 1]["text"])
                current_bbox = current["bbox"]
                next_bbox = visible_lines[index + 1]["bbox"]
                if current_text != FORMULA_WORD or not FORMULA_NUMBER_TOKEN_RE.fullmatch(next_text):
                    continue
                if current_bbox[0] < page_width * 0.55 and next_bbox[0] < page_width * 0.55:
                    continue
                issue = {
                    "issue": "rendered-formula-label-split-line",
                    "page": page_index,
                    "formula_word_text": current_text,
                    "formula_number_text": next_text,
                    "formula_word_bbox": [round(value, 2) for value in current_bbox],
                    "formula_number_bbox": [round(value, 2) for value in next_bbox],
                }
                split_pairs.append(issue)
                issues.append(issue)

    unique_sizes = sorted({size for size in label_sizes if size and size != "0.00"})
    if len(unique_sizes) > 1:
        issues.append(
            {
                "issue": "rendered-formula-label-font-size-drift",
                "unique_sizes": unique_sizes,
                "label_lines_sample": label_lines[:40],
            }
        )
    return {
        "rendered_pdf_path": str(pdf_path),
        "rendered_pdf_sha256": sha256(pdf_path),
        "rendered_formula_label_split_pair_count": len(split_pairs),
        "rendered_formula_label_issue_count": len(issues),
        "rendered_formula_number_unique_sizes": unique_sizes,
        "rendered_formula_label_issues": issues,
    }


def audit_formula_paragraph_alignment(root: ET.Element) -> dict[str, object]:
    """Verify standalone formula paragraphs use center and right tab stops."""

    issues: list[dict[str, object]] = []
    checked: list[dict[str, object]] = []
    usable = usable_width_twips(root)
    expected_center = int(round(usable / 2)) if usable else None
    expected_right = usable
    tolerance = 180

    paragraphs = root.findall(".//w:body//w:p", NS)
    for index, paragraph in enumerate(paragraphs):
        if not has_math(paragraph):
            continue
        text = re.sub(r"\s+", "", paragraph_text(paragraph))
        if STRICT_FORMULA_NUMBER_TOKEN_RE.search(text) is None:
            continue
        if paragraph_inside_table(paragraph):
            continue
        tab_stops = paragraph_tab_stops(paragraph)
        tab_runs = paragraph_direct_tab_count(paragraph)
        center_matches = [
            stop
            for stop in tab_stops
            if stop.get("val") == "center"
            and expected_center is not None
            and abs(int(stop.get("pos", 0)) - expected_center) <= tolerance
        ]
        right_matches = [
            stop
            for stop in tab_stops
            if stop.get("val") == "right"
            and expected_right is not None
            and abs(int(stop.get("pos", 0)) - expected_right) <= tolerance
        ]
        item = {
            "paragraph_index": index,
            "text": text,
            "tab_run_count": tab_runs,
            "tab_stops": tab_stops,
            "usable_width_twips": usable,
            "expected_center_tab_twips": expected_center,
            "expected_right_tab_twips": expected_right,
        }
        checked.append(item)
        if usable is None:
            issues.append({**item, "issue": "formula-layout-usable-width-missing"})
        if not center_matches:
            issues.append({**item, "issue": "formula-missing-center-tab-stop"})
        if not right_matches:
            issues.append({**item, "issue": "formula-missing-right-tab-stop"})
        if tab_runs < 2:
            issues.append({**item, "issue": "formula-missing-center-or-right-tab-run"})

    center_missing = len([issue for issue in issues if issue.get("issue") == "formula-missing-center-tab-stop"])
    right_missing = len([issue for issue in issues if issue.get("issue") == "formula-missing-right-tab-stop"])
    return {
        "formula_paragraph_layout_checked_count": len(checked),
        "formula_paragraph_centered_count": len(checked) - center_missing,
        "formula_paragraph_right_number_count": len(checked) - right_missing,
        "formula_paragraph_layout_issue_count": len(issues),
        "formula_paragraph_layout_issues": issues,
        "formula_paragraph_layout_verdict": "pass" if not issues else "fail",
    }


def audit_formula_number_layout(root: ET.Element) -> dict[str, object]:
    """Detect right-side formula number cells that wrap or drift in typography."""

    issues: list[dict[str, object]] = []
    number_cells: list[dict[str, object]] = []
    doc_number_sizes: list[str] = []

    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        table_has_math = bool(
            table.findall(".//m:oMath", NS)
            or table.findall(".//m:oMathPara", NS)
            or table.findall(".//w:object", NS)
        )
        if not table_has_math:
            continue
        for row_index, row in enumerate(table.findall("./w:tr", NS), start=1):
            cells = row.findall("./w:tc", NS)
            row_has_math = any(
                cell.findall(".//m:oMath", NS)
                or cell.findall(".//m:oMathPara", NS)
                or cell.findall(".//w:object", NS)
                for cell in cells
            )
            if not row_has_math:
                continue
            for cell_index, cell in enumerate(cells, start=1):
                text = re.sub(r"\s+", " ", cell_text(cell)).strip()
                compact_text = re.sub(r"\s+", "", cell_text(cell)).strip()
                if not compact_text or not FORMULA_NUMBER_CANDIDATE_CELL_RE.fullmatch(compact_text):
                    continue
                paragraph_texts = nonempty_cell_paragraph_texts(cell)
                sizes = cell_text_run_sizes(cell)
                has_no_wrap = cell_has_no_wrap(cell)
                style_valid = FORMULA_NUMBER_CELL_RE.fullmatch(compact_text) is not None
                doc_number_sizes.extend(size for size in sizes if size)
                item = {
                    "table_index": table_index,
                    "row_index": row_index,
                    "cell_index": cell_index,
                    "text": text,
                    "compact_text": compact_text,
                    "label_style": "formula-word-hyphen" if style_valid else "invalid",
                    "paragraph_texts": paragraph_texts,
                    "run_sizes": sizes,
                    "has_no_wrap": has_no_wrap,
                }
                number_cells.append(item)
                if not style_valid:
                    issues.append(
                        {
                            **item,
                            "issue": "formula-number-cell-style-invalid",
                            "detail": "formula number cell must use exact visible style 式(6-1): required 式 prefix, ASCII parentheses, and hyphen chapter numbering; bare (6-1), dot (6.1), full-width variants, or split labels are invalid",
                        }
                    )
                if has_no_wrap:
                    issues.append(
                        {
                            **item,
                            "issue": "formula-number-cell-invalid-no-wrap",
                            "detail": "formula number cell must not use schema-invalid w:noWrap; prevent wrapping through legal cell width and rendered PDF verification",
                        }
                    )
                if not sizes or any(not size for size in sizes):
                    issues.append(
                        {
                            **item,
                            "issue": "formula-number-cell-missing-direct-font-size",
                            "detail": "formula number visible runs must carry direct w:sz values",
                        }
                    )
                if len(paragraph_texts) != 1:
                    issues.append(
                        {
                            **item,
                            "issue": "formula-number-cell-split-paragraphs",
                            "detail": "formula number cell must be one visible line/paragraph",
                        }
                    )
                if len({size for size in sizes if size}) > 1:
                    issues.append(
                        {
                            **item,
                            "issue": "formula-number-cell-internal-size-drift",
                            "detail": "one formula number cell has multiple direct font sizes",
                        }
                    )

    unique_doc_sizes = sorted({size for size in doc_number_sizes if size})
    if len(unique_doc_sizes) > 1:
        issues.append(
            {
                "issue": "formula-number-document-size-drift",
                "detail": "formula number cells use inconsistent direct font sizes",
                "unique_run_sizes": unique_doc_sizes,
            }
        )
    return {
        "formula_number_table_count": len({item["table_index"] for item in number_cells}),
        "formula_number_cell_count": len(number_cells),
        "formula_number_layout_issue_count": len(issues),
        "formula_number_cells": number_cells,
        "formula_number_layout_issues": issues,
    }


def audit_formula_number_requirement(
    paragraphs: list[ET.Element],
    real_formula_paragraphs: list[dict[str, object]],
    number_layout: dict[str, object],
    *,
    rendered_pdf: Path | None,
) -> dict[str, object]:
    """Fail closed when real body formulas exist without visible strict numbering.

    Earlier versions only validated formula-number table cells when they existed.
    That let a manuscript with OMML formulas and no numbering surface pass.  The
    acceptance rule is stronger: each standalone body formula group must have a
    visible strict label such as 式(2-1), and rendered PDF evidence must be bound
    whenever formula numbering is required.
    """

    body_formula_paragraphs = [
        item
        for item in real_formula_paragraphs
        if bool(item.get("body_scope")) and bool(item.get("standalone_number_required"))
    ]
    strict_label_items: list[dict[str, object]] = []
    loose_label_items: list[dict[str, object]] = []

    body_formula_indexes = {
        int(item.get("paragraph_index", -1))
        for item in body_formula_paragraphs
    }
    for index, paragraph in enumerate(paragraphs):
        if index not in body_formula_indexes:
            continue
        text = re.sub(r"\s+", "", paragraph_text(paragraph))
        if not text:
            continue
        strict_matches = STRICT_FORMULA_NUMBER_TOKEN_RE.findall(text)
        loose_matches = FORMULA_NUMBER_TOKEN_RE.findall(text)
        for label in strict_matches:
            strict_label_items.append({"paragraph_index": index, "text": label})
        for label in loose_matches:
            if STRICT_FORMULA_NUMBER_TOKEN_RE.fullmatch(label):
                continue
            loose_label_items.append({"paragraph_index": index, "text": label})

    strict_label_count = int(number_layout.get("formula_number_cell_count", 0)) + len(strict_label_items)
    issues: list[dict[str, object]] = []
    required_count = len(body_formula_paragraphs)

    if required_count and strict_label_count < required_count:
        issues.append(
            {
                "issue": "formula-number-required-missing",
                "detail": "each body formula object must have a visible strict formula number label such as 式(2-1)",
                "body_formula_group_count": required_count,
                "strict_formula_number_label_count": strict_label_count,
            }
        )
    if loose_label_items:
        issues.append(
            {
                "issue": "formula-number-label-style-invalid-outside-cell",
                "detail": "bare, dotted, full-width, or prefixless formula labels are invalid; use 式(章-序)",
                "examples": loose_label_items[:20],
            }
        )
    if required_count and rendered_pdf is None:
        issues.append(
            {
                "issue": "rendered-formula-number-audit-missing",
                "detail": "formula numbering requires --rendered-pdf evidence bound to the final PDF",
            }
        )

    return {
        "formula_number_requirement_verdict": "pass" if not issues else "fail",
        "body_formula_group_count": required_count,
        "strict_formula_number_label_count": strict_label_count,
        "loose_formula_number_label_count": len(loose_label_items),
        "formula_number_requirement_issue_count": len(issues),
        "formula_number_requirement_issues": issues,
    }


def detect_body_bounds(paragraphs: list[ET.Element]) -> tuple[int, int, int | None]:
    body_start = 0
    for index, paragraph in enumerate(paragraphs):
        text = re.sub(r"\s+", " ", paragraph_text(paragraph)).strip()
        style = paragraph_style(paragraph)
        if not style.upper().startswith("TOC") and BODY_START_RE.search(text):
            body_start = index
            break

    body_end = len(paragraphs)
    appendix_start: int | None = None
    for index in range(body_start, len(paragraphs)):
        text = re.sub(r"\s+", " ", paragraph_text(paragraphs[index])).strip()
        style = paragraph_style(paragraphs[index])
        if style.upper().startswith("TOC"):
            continue
        if appendix_start is None and APPENDIX_HEADING_RE.search(text):
            appendix_start = index
        if TAIL_HEADING_RE.search(text):
            body_end = index
            break
    return body_start, body_end, appendix_start


FORMULA_DUMP_MARKER_RE = re.compile(
    r"(?:\u53c2\u8003\u8ba1\u7b97\u516c\u5f0f\u8865\u5165|\u516c\u5f0f\u8865\u5165|\u4ee5\u4e0b\u8865\u5165\u53c2\u8003\u8ba1\u7b97\u7a3f|\u8865\u5165\u53c2\u8003\u8ba1\u7b97\u7a3f|formula\s+dump)",
    re.I,
)
FORMULA_LABEL_TEXT_RE = re.compile(r"^\s*\u5f0f[\(（]\d+[-.]\d+[A-Za-z]?[\)）]\s*$")
VISIBLE_BODY_START_RE = re.compile(r"^\s*(?:第[\d一二三四五六七八九十百]+章|[1-9]\d?(?:\.\d+)*\s+)")
VISIBLE_TAIL_HEADING_RE = re.compile(r"^\s*(?:参考文献|致谢|附录|References|Acknowledgements?|Appendix)\b", re.I)
GENERIC_REFERENCE_NARRATIVE_RE = re.compile(
    r"本项校核承接参考计算稿的计算过程，说明相关结构参数、载荷参数与强度条件的代入和判定依据。"
)


CALCULATION_PROCESS_REQUIRED_RE = re.compile(r"(?:\u5f0f\u4e2d|\u5176\u4e2d)")
CALCULATION_PROCESS_RESULT_RE = re.compile(
    r"(?:\u4ee3\u5165|\u53d6\u503c|\u8ba1\u7b97|\u5f97\u5230|\u7ed3\u679c|\u5b89\u5168\u7cfb\u6570|\u88d5\u5ea6|\u6ee1\u8db3|\u5c0f\u4e8e|\u5927\u4e8e|\u4e0d\u8d85\u8fc7|\u6821\u6838|\u5224\u5b9a|\u9009\u53d6)"
)
GENERIC_FORMULA_PROCESS_RE = re.compile(
    r"\u5f0f\s*[\(\uff08]\d+[-.]\d+[A-Za-z]?[\)\uff09]\s*\u7528\u4e8e.*?\u6b63\u6587\u8ba1\u7b97.*?\u4ee3\u5165\u503c\u4e0e\u5224\u5b9a\u7ed3\u679c\u63a5\u7eed\u670d\u52a1\u4e8e\u672c\u8282\u540e\u7eed\u6821\u6838"
)


def is_formula_narrative_text(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if len(compact) < 18:
        return False
    if FORMULA_DUMP_MARKER_RE.search(text):
        return False
    if FORMULA_LABEL_TEXT_RE.fullmatch(compact):
        return False
    if TEXT_FORMULA_TOKEN_RE.search(text) and len(compact) < 28:
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", compact))


def is_formula_calculation_process_text(text: str) -> bool:
    """Detect screenshot-like calculation explanation around a formula."""

    compact = re.sub(r"\s+", "", text or "")
    if len(compact) < 24:
        return False
    if FORMULA_DUMP_MARKER_RE.search(text):
        return False
    if GENERIC_FORMULA_PROCESS_RE.search(text):
        return False
    if not CALCULATION_PROCESS_REQUIRED_RE.search(text):
        return False
    if not CALCULATION_PROCESS_RESULT_RE.search(text):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", compact))


def direct_child_text(child: ET.Element) -> str:
    return "".join(node.text or "" for node in child.findall(".//w:t", NS))


FORMULA_NARRATIVE_BAD_STYLE_RE = re.compile(
    r"(?:Heading|Title|Subtitle|TOC|Caption|Bibliography|Reference|Keyword|Abstract)",
    re.I,
)


def direct_child_paragraph_style(child: ET.Element) -> str:
    if child.tag != f"{{{W}}}p":
        return ""
    node = child.find("./w:pPr/w:pStyle", NS)
    return str(node.get(f"{{{W}}}val") or "") if node is not None else ""


def is_formula_narrative_style_polluted(child: ET.Element, text: str) -> tuple[bool, str]:
    style = direct_child_paragraph_style(child)
    compact = re.sub(r"\s+", "", text or "")
    if style and FORMULA_NARRATIVE_BAD_STYLE_RE.search(style):
        return True, f"blocked-style:{style}"
    if re.match(r"^(?:\u5173\u952e\u8bcd|Keywords?|Key\s+words?|Abstract|摘要|参考文献)\b", compact, re.I):
        return True, "protected-surface-text"
    return False, ""


def direct_child_has_math(child: ET.Element) -> bool:
    return bool(child.findall(".//m:oMath", NS) or child.findall(".//m:oMathPara", NS))


def direct_child_is_formula_table(child: ET.Element) -> bool:
    return child.tag == f"{{{W}}}tbl" and direct_child_has_math(child)


def direct_child_is_formula_surface(child: ET.Element) -> bool:
    return child.tag in {f"{{{W}}}p", f"{{{W}}}tbl"} and direct_child_has_math(child)


def direct_child_chapter_scope(children: list[ET.Element]) -> list[bool]:
    in_body = False
    scope: list[bool] = []
    for child in children:
        text = re.sub(r"\s+", " ", direct_child_text(child)).strip()
        style = direct_child_paragraph_style(child)
        is_toc = style.upper().startswith("TOC")
        if (
            child.tag == f"{{{W}}}p"
            and not is_toc
            and (BODY_START_RE.search(text) or VISIBLE_BODY_START_RE.search(text))
        ):
            in_body = True
        if (
            child.tag == f"{{{W}}}p"
            and not is_toc
            and (TAIL_HEADING_RE.search(text) or VISIBLE_TAIL_HEADING_RE.search(text))
        ):
            in_body = False
        scope.append(in_body)
    return scope


def audit_formula_narrative_context(
    root: ET.Element,
    *,
    window: int = 4,
) -> dict[str, object]:
    """Reject formula dumps that pass object-count checks but lack body prose.

    The check is intentionally based on direct body children, so a borderless
    formula table plus its number cell is treated as one formula surface.
    """

    body = root.find(".//w:body", NS)
    if body is None:
        return {
            "formula_narrative_context_verdict": "fail",
            "formula_narrative_context_issue_count": 1,
            "formula_narrative_context_issues": [{"issue": "missing-document-body"}],
            "formula_dump_marker_count": 0,
            "formula_without_calculation_process_explanation_count": 0,
            "orphan_formula_style_issue_count": 0,
            "formula_narrative_style_issue_count": 0,
            "formula_narrative_style_issues": [],
        }

    children = list(body)
    child_scope = direct_child_chapter_scope(children)
    marker_issues: list[dict[str, object]] = []
    narrative_issues: list[dict[str, object]] = []
    process_issues: list[dict[str, object]] = []
    narrative_style_issues: list[dict[str, object]] = []
    generic_reference_narrative_indexes: list[int] = []
    for index, child in enumerate(children):
        text = re.sub(r"\s+", " ", direct_child_text(child)).strip()
        if text and FORMULA_DUMP_MARKER_RE.search(text):
            marker_issues.append(
                {
                    "issue": "formula-dump-marker-present",
                    "body_child_index": index,
                    "text": text[:120],
                }
            )
        if text and GENERIC_REFERENCE_NARRATIVE_RE.fullmatch(text):
            generic_reference_narrative_indexes.append(index)
        if not direct_child_is_formula_surface(child):
            continue

        has_context = False
        has_process = False
        nearest_context = ""
        nearest_process = ""
        stop_search = False
        for offset in range(1, max(1, window) + 1):
            for candidate_index in (index - offset, index + offset):
                if candidate_index < 0 or candidate_index >= len(children):
                    continue
                candidate = children[candidate_index]
                if direct_child_has_math(candidate):
                    continue
                candidate_text = re.sub(r"\s+", " ", direct_child_text(candidate)).strip()
                if is_formula_narrative_text(candidate_text):
                    polluted, reason = is_formula_narrative_style_polluted(candidate, candidate_text)
                    if polluted:
                        narrative_style_issues.append(
                            {
                                "issue": "formula-nearby-explanation-style-pollution",
                                "formula_body_child_index": index,
                                "context_body_child_index": candidate_index,
                                "style": direct_child_paragraph_style(candidate),
                                "reason": reason,
                                "text": candidate_text[:120],
                            }
                        )
                    has_context = True
                    if not nearest_context:
                        nearest_context = candidate_text[:120]
                    if is_formula_calculation_process_text(candidate_text):
                        has_process = True
                        nearest_process = candidate_text[:120]
                        stop_search = True
                        break
            if stop_search:
                break
        if not has_context:
            label = re.sub(r"\s+", "", text)
            narrative_issues.append(
                {
                    "issue": "formula-without-nearby-body-explanation",
                    "body_child_index": index,
                    "formula_label_or_text": label[:80],
                    "checked_window": window,
                    "nearest_context": nearest_context,
                }
            )
        elif child_scope[index] and not has_process:
            label = re.sub(r"\s+", "", text)
            process_issues.append(
                {
                    "issue": "formula-without-calculation-process-explanation",
                    "body_child_index": index,
                    "formula_label_or_text": label[:80],
                    "checked_window": window,
                    "nearest_context": nearest_context,
                    "nearest_process": nearest_process,
                    "required_shape": "nearby prose must include variable explanation such as 式中/其中 and a value/result/judgement sentence such as 代入/得到/安全系数/满足",
                }
            )

    repeated_narrative_issues: list[dict[str, object]] = []
    if len(generic_reference_narrative_indexes) > 8:
        repeated_narrative_issues.append(
            {
                "issue": "formula-generic-reference-narrative-repeated",
                "detail": "nearby formula prose is too generic and repeated; formulas must be integrated into section-specific calculation explanation",
                "repeat_count": len(generic_reference_narrative_indexes),
                "sample_body_child_indexes": generic_reference_narrative_indexes[:20],
            }
        )

    style_ids: set[str] = set()
    styles_part = root.getroottree().getroot()
    # document.xml does not contain style definitions. This placeholder is kept
    # for schema stability; style validation is populated by audit_docx after it
    # reads styles.xml.
    _ = styles_part

    issues = marker_issues + narrative_issues + process_issues + narrative_style_issues + repeated_narrative_issues
    return {
        "formula_narrative_context_verdict": "pass" if not issues else "fail",
        "formula_narrative_context_issue_count": len(issues),
        "formula_narrative_context_issues": issues,
        "formula_dump_marker_count": len(marker_issues),
        "formula_without_nearby_body_explanation_count": len(narrative_issues),
        "formula_without_calculation_process_explanation_count": len(process_issues),
        "formula_narrative_style_issue_count": len(narrative_style_issues),
        "formula_generic_reference_narrative_repeat_count": len(generic_reference_narrative_indexes),
        "formula_narrative_style_issues": narrative_style_issues,
        "orphan_formula_style_issue_count": 0,
        "orphan_formula_style_issues": [],
    }


def audit_formula_orphan_styles(root: ET.Element, style_ids: set[str]) -> dict[str, object]:
    issues: list[dict[str, object]] = []
    paragraphs = root.findall(".//w:body//w:p", NS)
    for index, paragraph in enumerate(paragraphs):
        if not has_math(paragraph):
            continue
        style = paragraph_style(paragraph)
        if style and style not in style_ids:
            issues.append(
                {
                    "issue": "formula-paragraph-orphan-style",
                    "paragraph_index": index,
                    "style": style,
                    "math_text": paragraph_math_text(paragraph)[:120],
                }
            )
    return {
        "orphan_formula_style_issue_count": len(issues),
        "orphan_formula_style_issues": issues,
    }


def normalize_formula_body_text(value: str) -> str:
    value = re.sub(r"\u5f0f\s*[\uff08(]\s*\d+\s*[-\uff0d.]\s*\d+[A-Za-z]?\s*[\uff09)]", "", value or "")
    value = re.sub(r"[\s\u3000]+", "", value)
    return value.strip().lower()


def audit_formula_duplicate_density(
    real_formula_paragraphs: list[dict[str, object]],
    min_unique_formula_ratio: float | None = None,
) -> dict[str, object]:
    bodies: list[str] = []
    for item in real_formula_paragraphs:
        text = normalize_formula_body_text(str(item.get("math_text") or item.get("text") or ""))
        if text:
            bodies.append(text)
    counts = Counter(bodies)
    duplicate_rows = [
        {"formula_text": text[:160], "count": count}
        for text, count in counts.most_common(20)
        if count > 1
    ]
    total = len(bodies)
    unique = len(counts)
    ratio = unique / total if total else 1.0
    issues: list[dict[str, object]] = []
    if min_unique_formula_ratio is not None and ratio < min_unique_formula_ratio:
        issues.append(
            {
                "issue": "formula-content-duplicate-density-too-high",
                "detail": (
                    "high-density formula delivery repeats too many identical formula bodies; "
                    "formula count cannot be satisfied by copying a small template set"
                ),
                "unique_formula_ratio": round(ratio, 4),
                "min_unique_formula_ratio": min_unique_formula_ratio,
            }
        )
    return {
        "formula_duplicate_density_verdict": "pass" if not issues else "fail",
        "formula_body_text_count": total,
        "unique_formula_body_text_count": unique,
        "duplicate_formula_body_text_count": max(0, total - unique),
        "unique_formula_ratio": round(ratio, 4),
        "min_unique_formula_ratio": min_unique_formula_ratio,
        "duplicate_formula_samples": duplicate_rows,
        "formula_duplicate_density_issue_count": len(issues),
        "formula_duplicate_density_issues": issues,
    }


def is_formula_like_text(text: str) -> tuple[bool, list[str]]:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return False, []

    reasons: list[str] = []
    if FORMULA_NUMBER_TOKEN_RE.search(compact):
        reasons.append("formula-number")
    if ASSIGNMENT_RE.search(compact):
        reasons.append("assignment")
    if OPERATOR_RE.search(compact):
        reasons.append("operator")
    if TEXT_FORMULA_TOKEN_RE.search(compact):
        reasons.append("underscore-or-known-formula-token")
    if re.search(r"\b\d+(?:\.\d+)?\s*[*/×+-]\s*\d+", compact):
        reasons.append("numeric-expression")
    if INEQUALITY_RE.search(compact):
        reasons.append("inequality")

    formula_like = "assignment" in reasons and (
        "operator" in reasons
        or "formula-number" in reasons
        or "underscore-or-known-formula-token" in reasons
        or "inequality" in reasons
    )
    if STATUS_ASSIGNMENT_RE.fullmatch(compact) and not formula_like:
        return False, ["status-or-api-assignment"]
    code_like = bool(CODE_LINE_RE.search(compact) or CODE_TOKEN_RE.search(compact))
    if code_like and not (formula_like and re.search(r"[\u4e00-\u9fff]", compact)):
        return False, ["code-or-api-line"]
    return formula_like, reasons


def cjk_char_count(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text or ""))


def requires_standalone_formula_number(paragraph: ET.Element) -> bool:
    """Identify standalone formula surfaces, excluding prose with inline math."""

    if not has_math(paragraph):
        return False
    visible_text = re.sub(r"\s+", "", paragraph_text(paragraph)).strip()
    math_text = re.sub(r"\s+", "", paragraph_math_text(paragraph)).strip()
    combined_text = visible_text + math_text
    contains_legacy_object = legacy_formula_object_count(paragraph) > 0
    contains_omml = omml_math_count(paragraph) > 0
    if row_has_formula_number_cell(paragraph_ancestor_row(paragraph)):
        return True
    if STRICT_FORMULA_NUMBER_TOKEN_RE.search(visible_text) or FORMULA_NUMBER_TOKEN_RE.search(visible_text):
        return True
    if paragraph_inside_table(paragraph):
        return False
    if PROSE_FORMULA_SKIP_RE.search(visible_text):
        return False
    if contains_legacy_object and not contains_omml:
        if not visible_text or visible_text in {".", "\uff0c", ",", "\u2235", "\u2234"}:
            return True
        if MATH_SIGNAL_RE.search(visible_text) or re.search(r"\d", visible_text):
            return True
        if cjk_char_count(visible_text) <= 2 and len(visible_text) <= 6:
            return True
        return False
    if not math_text and not MATH_SIGNAL_RE.search(visible_text):
        return False
    has_formula_signal = bool(MATH_SIGNAL_RE.search(combined_text) or MATH_SYMBOL_RE.search(math_text))
    if not has_formula_signal:
        return False
    if cjk_char_count(visible_text) > 30 and len(math_text) <= 16:
        return False
    if cjk_char_count(visible_text) <= 18:
        return True
    if len(math_text) >= 18:
        return True
    return False


def audit_docx(
    docx_path: Path,
    min_formula_count: int | None = None,
    min_body_formula_count: int | None = None,
    min_body_formula_group_count: int | None = None,
    min_strict_formula_number_label_count: int | None = None,
    rendered_pdf: Path | None = None,
    require_formula_narrative: bool = False,
    formula_narrative_window: int = 4,
    min_unique_formula_ratio: float | None = None,
) -> dict[str, object]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        try:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
        except KeyError:
            styles_root = ET.Element(f"{{{W}}}styles")

    paragraphs = root.findall(".//w:body//w:p", NS)
    omml_math_objects = root.findall(".//m:oMath", NS)
    legacy_formula_objects = root.findall(".//w:object", NS)
    math_objects = omml_math_objects + legacy_formula_objects
    math_paragraphs = root.findall(".//m:oMathPara", NS)
    body_start, body_end, appendix_start = detect_body_bounds(paragraphs)
    style_ids = {
        str(style.get(f"{{{W}}}styleId") or "")
        for style in styles_root.findall(".//w:style", NS)
        if style.get(f"{{{W}}}styleId")
    }

    formula_like_paragraphs: list[dict[str, object]] = []
    pseudo_formula_paragraphs: list[dict[str, object]] = []
    real_formula_paragraphs: list[dict[str, object]] = []
    unit_math_paragraphs: list[dict[str, object]] = []
    missing_formula_after_leadin: list[dict[str, object]] = []
    body_math_object_count = 0
    appendix_math_object_count = 0
    number_layout = audit_formula_number_layout(root)
    paragraph_layout = audit_formula_paragraph_alignment(root)
    rendered_label_layout = audit_rendered_formula_labels(rendered_pdf)
    raw_math_command_tokens = audit_raw_math_command_tokens(root)
    rendered_raw_formula_tokens = audit_rendered_formula_raw_tokens(rendered_pdf)
    narrative_context = audit_formula_narrative_context(root, window=formula_narrative_window)
    orphan_style_report = audit_formula_orphan_styles(root, style_ids)
    narrative_context["orphan_formula_style_issue_count"] = orphan_style_report[
        "orphan_formula_style_issue_count"
    ]
    narrative_context["orphan_formula_style_issues"] = orphan_style_report[
        "orphan_formula_style_issues"
    ]
    if int(orphan_style_report["orphan_formula_style_issue_count"]) > 0:
        narrative_context["formula_narrative_context_issue_count"] = int(
            narrative_context["formula_narrative_context_issue_count"]
        ) + int(orphan_style_report["orphan_formula_style_issue_count"])
        narrative_context["formula_narrative_context_issues"] = list(
            narrative_context["formula_narrative_context_issues"]
        ) + list(orphan_style_report["orphan_formula_style_issues"])
        narrative_context["formula_narrative_context_verdict"] = "fail"

    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph)
        formula_like, reasons = is_formula_like_text(text)
        contains_math = has_math(paragraph)
        paragraph_math_count = math_count(paragraph)
        if paragraph_math_count:
            if body_start <= index < body_end:
                body_math_object_count += paragraph_math_count
            elif appendix_start is not None and index >= appendix_start:
                appendix_math_object_count += paragraph_math_count
        if not formula_like and not contains_math:
            continue
        unit_math = contains_math and is_non_formula_table_unit_math(paragraph, formula_like=formula_like)
        if unit_math:
            unit_math_paragraphs.append(
                {
                    "paragraph_index": index,
                    "text": text.strip(),
                    "math_text": paragraph_math_text(paragraph),
                    "body_scope": bool(body_start <= index < body_end),
                    "reason": "table-unit-math-not-standalone-formula",
                }
            )
            continue

        item = {
            "paragraph_index": index,
            "text": text.strip(),
            "math_text": paragraph_math_text(paragraph),
            "has_formula_object": contains_math,
            "has_omml": omml_math_count(paragraph) > 0,
            "omml_math_object_count": omml_math_count(paragraph),
            "legacy_formula_object_count": legacy_formula_object_count(paragraph),
            "body_scope": bool(body_start <= index < body_end),
            "appendix_scope": bool(appendix_start is not None and index >= appendix_start),
            "standalone_number_required": contains_math and requires_standalone_formula_number(paragraph),
            "reasons": reasons,
        }
        if formula_like:
            formula_like_paragraphs.append(item)
        if contains_math:
            real_formula_paragraphs.append(item)
        elif formula_like:
            pseudo_formula_paragraphs.append(item)

    for index, paragraph in enumerate(paragraphs):
        text = re.sub(r"\s+", " ", paragraph_text(paragraph)).strip()
        if not text or has_math(paragraph) or not FORMULA_LEADIN_RE.search(text):
            continue
        next_index = None
        next_text = ""
        next_has_math = False
        next_formula_like = False
        for candidate_index in range(index + 1, len(paragraphs)):
            candidate = paragraphs[candidate_index]
            candidate_text = re.sub(r"\s+", " ", paragraph_text(candidate)).strip()
            if not candidate_text and not has_math(candidate):
                continue
            next_index = candidate_index
            next_text = candidate_text
            next_has_math = has_math(candidate)
            next_formula_like, _ = is_formula_like_text(candidate_text)
            break
        if next_has_math or next_formula_like:
            continue
        missing_formula_after_leadin.append(
            {
                "paragraph_index": index,
                "text": text,
                "next_paragraph_index": next_index,
                "next_text": next_text,
                "reason": (
                    "next paragraph starts formula explanation but no formula object or formula-like expression appears"
                    if FORMULA_EXPLANATION_START_RE.search(next_text)
                    else "formula lead-in is not followed by a formula object or formula-like expression"
                ),
            }
        )

    number_requirement = audit_formula_number_requirement(
        paragraphs,
        real_formula_paragraphs,
        number_layout,
        rendered_pdf=rendered_pdf,
    )
    rendered_label_required_missing = (
        rendered_pdf is None and int(number_layout.get("formula_number_cell_count", 0)) > 0
    )
    if rendered_label_required_missing:
        rendered_label_issues = list(rendered_label_layout["rendered_formula_label_issues"])
        rendered_label_issues.append(
            {
                "issue": "rendered-formula-label-audit-missing",
                "detail": "formula number cells exist but --rendered-pdf was not supplied",
            }
        )
        rendered_label_layout["rendered_formula_label_issues"] = rendered_label_issues
        rendered_label_layout["rendered_formula_label_issue_count"] = int(
            rendered_label_layout.get("rendered_formula_label_issue_count", 0)
        ) + 1

    effective_min_unique_formula_ratio = min_unique_formula_ratio
    if (
        effective_min_unique_formula_ratio is None
        and min_body_formula_count is not None
        and min_body_formula_count >= 100
    ):
        effective_min_unique_formula_ratio = 0.60
    duplicate_density = audit_formula_duplicate_density(
        real_formula_paragraphs,
        min_unique_formula_ratio=effective_min_unique_formula_ratio,
    )

    result = (
        "pass"
        if not pseudo_formula_paragraphs
        and not missing_formula_after_leadin
        and int(number_layout.get("formula_number_layout_issue_count", 0)) == 0
        and int(paragraph_layout.get("formula_paragraph_layout_issue_count", 0)) == 0
        and int(number_requirement.get("formula_number_requirement_issue_count", 0)) == 0
        and int(rendered_label_layout.get("rendered_formula_label_issue_count", 0)) == 0
        and int(raw_math_command_tokens.get("raw_math_command_token_count", 0)) == 0
        and int(rendered_raw_formula_tokens.get("rendered_raw_math_token_page_count", 0)) == 0
        and int(rendered_raw_formula_tokens.get("rendered_near_empty_formula_page_count", 0)) == 0
        and int(duplicate_density.get("formula_duplicate_density_issue_count", 0)) == 0
        else "fail"
    )
    report = {
        "schema": "graduation-project-builder.formula-object-audit.v1",
        "result": result,
        "docx_path": str(docx_path),
        "docx_sha256": sha256(docx_path),
        "math_object_count": len(math_objects),
        "omml_math_object_count": len(omml_math_objects),
        "legacy_formula_object_count": len(legacy_formula_objects),
        "math_paragraph_count": len(math_paragraphs),
        "body_start_paragraph_index": body_start,
        "body_end_paragraph_index": body_end,
        "appendix_start_paragraph_index": appendix_start,
        "body_math_object_count": body_math_object_count,
        "appendix_math_object_count": appendix_math_object_count,
        "non_body_math_object_count": max(0, len(math_objects) - body_math_object_count),
        "formula_like_paragraph_count": len(formula_like_paragraphs),
        "real_formula_paragraph_count": len(real_formula_paragraphs),
        "pseudo_formula_count": len(pseudo_formula_paragraphs),
        "missing_formula_after_leadin_count": len(missing_formula_after_leadin),
        "formula_number_table_count": number_layout["formula_number_table_count"],
        "formula_number_cell_count": number_layout["formula_number_cell_count"],
        "formula_number_layout_issue_count": number_layout["formula_number_layout_issue_count"],
        "formula_number_layout_issues": number_layout["formula_number_layout_issues"],
        "formula_paragraph_layout_checked_count": paragraph_layout["formula_paragraph_layout_checked_count"],
        "formula_paragraph_centered_count": paragraph_layout["formula_paragraph_centered_count"],
        "formula_paragraph_right_number_count": paragraph_layout["formula_paragraph_right_number_count"],
        "formula_paragraph_layout_issue_count": paragraph_layout["formula_paragraph_layout_issue_count"],
        "formula_paragraph_layout_issues": paragraph_layout["formula_paragraph_layout_issues"],
        "formula_paragraph_layout_verdict": paragraph_layout["formula_paragraph_layout_verdict"],
        "formula_number_requirement_verdict": number_requirement["formula_number_requirement_verdict"],
        "body_formula_group_count": number_requirement["body_formula_group_count"],
        "strict_formula_number_label_count": number_requirement["strict_formula_number_label_count"],
        "loose_formula_number_label_count": number_requirement["loose_formula_number_label_count"],
        "formula_number_requirement_issue_count": number_requirement["formula_number_requirement_issue_count"],
        "formula_number_requirement_issues": number_requirement["formula_number_requirement_issues"],
        "rendered_pdf_path": rendered_label_layout["rendered_pdf_path"],
        "rendered_pdf_sha256": rendered_label_layout["rendered_pdf_sha256"],
        "rendered_formula_label_split_pair_count": rendered_label_layout[
            "rendered_formula_label_split_pair_count"
        ],
        "rendered_formula_label_issue_count": rendered_label_layout[
            "rendered_formula_label_issue_count"
        ],
        "rendered_formula_number_unique_sizes": rendered_label_layout[
            "rendered_formula_number_unique_sizes"
        ],
        "rendered_formula_label_issues": rendered_label_layout["rendered_formula_label_issues"],
        "raw_math_command_token_count": raw_math_command_tokens["raw_math_command_token_count"],
        "raw_math_command_token_issues": raw_math_command_tokens["raw_math_command_token_issues"],
        "rendered_raw_math_token_page_count": rendered_raw_formula_tokens[
            "rendered_raw_math_token_page_count"
        ],
        "rendered_raw_math_token_issues": rendered_raw_formula_tokens[
            "rendered_raw_math_token_issues"
        ],
        "rendered_near_empty_formula_page_count": rendered_raw_formula_tokens[
            "rendered_near_empty_formula_page_count"
        ],
        "rendered_near_empty_formula_pages": rendered_raw_formula_tokens[
            "rendered_near_empty_formula_pages"
        ],
        "formula_narrative_context_verdict": narrative_context["formula_narrative_context_verdict"],
        "formula_narrative_context_issue_count": narrative_context[
            "formula_narrative_context_issue_count"
        ],
        "formula_dump_marker_count": narrative_context["formula_dump_marker_count"],
        "formula_without_nearby_body_explanation_count": narrative_context[
            "formula_without_nearby_body_explanation_count"
        ],
        "formula_without_calculation_process_explanation_count": narrative_context[
            "formula_without_calculation_process_explanation_count"
        ],
        "formula_narrative_style_issue_count": narrative_context[
            "formula_narrative_style_issue_count"
        ],
        "formula_narrative_style_issues": narrative_context["formula_narrative_style_issues"],
        "orphan_formula_style_issue_count": narrative_context["orphan_formula_style_issue_count"],
        "formula_narrative_context_issues": narrative_context["formula_narrative_context_issues"],
        "formula_duplicate_density_verdict": duplicate_density["formula_duplicate_density_verdict"],
        "formula_body_text_count": duplicate_density["formula_body_text_count"],
        "unique_formula_body_text_count": duplicate_density["unique_formula_body_text_count"],
        "duplicate_formula_body_text_count": duplicate_density["duplicate_formula_body_text_count"],
        "unique_formula_ratio": duplicate_density["unique_formula_ratio"],
        "min_unique_formula_ratio": duplicate_density["min_unique_formula_ratio"],
        "duplicate_formula_samples": duplicate_density["duplicate_formula_samples"],
        "formula_duplicate_density_issue_count": duplicate_density[
            "formula_duplicate_density_issue_count"
        ],
        "formula_duplicate_density_issues": duplicate_density[
            "formula_duplicate_density_issues"
        ],
        "pseudo_formula_paragraphs": pseudo_formula_paragraphs,
        "unit_math_paragraphs": unit_math_paragraphs,
        "unit_math_paragraph_count": len(unit_math_paragraphs),
        "missing_formula_after_leadin": missing_formula_after_leadin,
        "real_formula_paragraphs": real_formula_paragraphs,
    }
    if require_formula_narrative:
        report["formula_narrative_required"] = True
        report["formula_narrative_window"] = formula_narrative_window
        if int(report["formula_narrative_context_issue_count"]) > 0:
            report["result"] = "fail"
    if min_formula_count is not None:
        minimum = max(0, min_formula_count)
        report["min_formula_count"] = minimum
        if report["math_object_count"] < minimum:
            report["result"] = "fail"
            report["formula_count_requirement"] = (
                f"below minimum formula count: math_object_count={report['math_object_count']} "
                f"min_formula_count={minimum}"
            )
    if min_body_formula_count is not None:
        minimum_body = max(0, min_body_formula_count)
        report["min_body_formula_count"] = minimum_body
        if report["body_math_object_count"] < minimum_body:
            report["result"] = "fail"
            report["body_formula_count_requirement"] = (
                f"below minimum body formula count: body_math_object_count={report['body_math_object_count']} "
                f"min_body_formula_count={minimum_body}"
            )
    if min_body_formula_group_count is not None:
        minimum_body_groups = max(0, min_body_formula_group_count)
        report["min_body_formula_group_count"] = minimum_body_groups
        if report["body_formula_group_count"] < minimum_body_groups:
            report["result"] = "fail"
            report["body_formula_group_count_requirement"] = (
                f"below minimum body formula group count: body_formula_group_count={report['body_formula_group_count']} "
                f"min_body_formula_group_count={minimum_body_groups}"
            )
    if min_strict_formula_number_label_count is not None:
        minimum_strict_labels = max(0, min_strict_formula_number_label_count)
        report["min_strict_formula_number_label_count"] = minimum_strict_labels
        if report["strict_formula_number_label_count"] < minimum_strict_labels:
            report["result"] = "fail"
            report["strict_formula_number_label_count_requirement"] = (
                "below minimum strict formula number label count: "
                f"strict_formula_number_label_count={report['strict_formula_number_label_count']} "
                f"min_strict_formula_number_label_count={minimum_strict_labels}"
            )
    return report


def _write_minimal_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>""",
        )
        package.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""")
        package.writestr("word/styles.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}"/>""")
        package.writestr("word/document.xml", document_xml)


def self_test_raw_math_tokens() -> dict[str, object]:
    if not is_rendered_formula_label_candidate("(2-16)"):
        raise AssertionError("short bare formula labels must remain rendered-label candidates")
    if not is_rendered_formula_label_candidate("式(2-16)"):
        raise AssertionError("prefixed formula labels must remain rendered-label candidates")
    if is_rendered_formula_label_candidate("(609147.9)"):
        raise AssertionError("large parenthesized equation operands must not be rendered-label candidates")
    if is_rendered_formula_label_candidate("(0.630416750)"):
        raise AssertionError("decimal equation operands must not be rendered-label candidates")
    if is_rendered_formula_label_candidate("(123626250)"):
        raise AssertionError("large integer equation operands must not be rendered-label candidates")

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:m="{M}">
  <w:body>
    <w:p>
      <m:oMath><m:r><m:t>sub</m:t></m:r></m:oMath>
      <w:r><w:t>式(1-1)</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "raw-token-formula.docx"
        _write_minimal_docx(path, document_xml)
        report = audit_docx(path, min_formula_count=1)
    checks = {
        "raw_math_command_token_count": report["raw_math_command_token_count"],
        "result": report["result"],
    }
    if report["result"] != "fail" or int(report["raw_math_command_token_count"]) != 1:
        raise AssertionError(f"raw math token self-test did not fail closed: {checks!r}")
    return {"result": "pass", "checks": checks}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit DOCX formula authenticity and reject plain-text pseudo-formulas."
    )
    parser.add_argument("docx", type=Path, nargs="?")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--report", type=Path)
    parser.add_argument(
        "--rendered-pdf",
        type=Path,
        default=None,
        help="Fail when rendered PDF formula labels split across lines or drift in visible font size.",
    )
    parser.add_argument("--allow-no-formulas", action="store_true")
    parser.add_argument(
        "--min-formula-count",
        type=int,
        default=None,
        help="Fail when the DOCX contains fewer OMML formula objects than this count.",
    )
    parser.add_argument(
        "--min-body-formula-count",
        type=int,
        default=None,
        help="Fail when body chapters contain fewer OMML formula objects than this count.",
    )
    parser.add_argument(
        "--min-body-formula-group-count",
        type=int,
        default=None,
        help="Fail when body chapters contain fewer standalone formula groups requiring numbering than this count.",
    )
    parser.add_argument(
        "--min-strict-formula-number-label-count",
        type=int,
        default=None,
        help="Fail when there are fewer strict visible formula-number labels such as 式(2-1) than this count.",
    )
    parser.add_argument(
        "--require-formula-narrative",
        action="store_true",
        help="Fail when body formula tables are concentrated without nearby explanatory prose or contain dump markers/orphan styles.",
    )
    parser.add_argument(
        "--formula-narrative-window",
        type=int,
        default=4,
        help="Direct body-child search radius used by --require-formula-narrative.",
    )
    parser.add_argument(
        "--min-unique-formula-ratio",
        type=float,
        default=None,
        help="Fail when repeated formula bodies make the unique/total formula ratio lower than this value.",
    )
    args = parser.parse_args()

    if args.self_test:
        report = self_test_raw_math_tokens()
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0
    if args.docx is None:
        parser.error("docx is required unless --self-test is supplied")

    report = audit_docx(
        args.docx,
        min_formula_count=args.min_formula_count,
        min_body_formula_count=args.min_body_formula_count,
        min_body_formula_group_count=args.min_body_formula_group_count,
        min_strict_formula_number_label_count=args.min_strict_formula_number_label_count,
        rendered_pdf=args.rendered_pdf,
        require_formula_narrative=args.require_formula_narrative,
        formula_narrative_window=args.formula_narrative_window,
        min_unique_formula_ratio=args.min_unique_formula_ratio,
    )
    if not args.allow_no_formulas and report["math_object_count"] == 0 and report["formula_like_paragraph_count"] == 0:
        report["result"] = "fail"
        report["missing_formula_evidence"] = "no formula-like text and no OMML formula objects found"
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")

    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
