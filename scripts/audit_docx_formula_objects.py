from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}

FORMULA_NUMBER_RE = re.compile(r"\(\s*\d+(?:[-.]\d+)?\s*\)")
ASSIGNMENT_RE = re.compile(r"[A-Za-z][A-Za-z0-9_{}]*\s*=")
OPERATOR_RE = re.compile(r"(\+|-|\*|/|×|=|\bmin\s*\(|\bmax\s*\(|\bsqrt\s*\(|\bsum\b)", re.I)
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
FORMULA_LEADIN_RE = re.compile(
    r"(?:\u4e0b\u5f0f|\u516c\u5f0f|\u8ba1\u7b97\u5f0f|\u8868\u8fbe\u5f0f|\u8868\u793a\u4e3a|\u5b9a\u4e49\u4e3a|\u53ef\u7528.*\u8868\u793a).*[\uff1a:]$"
)
FORMULA_EXPLANATION_START_RE = re.compile(r"^\s*(?:\u5176\u4e2d|\u5f0f\u4e2d|where\b)", re.I)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def has_math(paragraph: ET.Element) -> bool:
    return bool(paragraph.findall(".//m:oMath", NS) or paragraph.findall(".//m:oMathPara", NS))


def is_formula_like_text(text: str) -> tuple[bool, list[str]]:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return False, []
    if CODE_LINE_RE.search(compact) or CODE_TOKEN_RE.search(compact):
        return False, ["code-or-api-line"]
    if STATUS_ASSIGNMENT_RE.fullmatch(compact):
        return False, ["status-or-api-assignment"]

    reasons: list[str] = []
    if FORMULA_NUMBER_RE.search(compact):
        reasons.append("formula-number")
    if ASSIGNMENT_RE.search(compact):
        reasons.append("assignment")
    if OPERATOR_RE.search(compact):
        reasons.append("operator")
    if TEXT_FORMULA_TOKEN_RE.search(compact):
        reasons.append("underscore-or-known-formula-token")
    if re.search(r"\b\d+(?:\.\d+)?\s*[*/×+-]\s*\d+", compact):
        reasons.append("numeric-expression")

    formula_like = "assignment" in reasons and (
        "operator" in reasons
        or "formula-number" in reasons
        or "underscore-or-known-formula-token" in reasons
    )
    return formula_like, reasons


def audit_docx(docx_path: Path) -> dict[str, object]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))

    paragraphs = root.findall(".//w:body//w:p", NS)
    math_objects = root.findall(".//m:oMath", NS)
    math_paragraphs = root.findall(".//m:oMathPara", NS)

    formula_like_paragraphs: list[dict[str, object]] = []
    pseudo_formula_paragraphs: list[dict[str, object]] = []
    real_formula_paragraphs: list[dict[str, object]] = []
    missing_formula_after_leadin: list[dict[str, object]] = []

    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph)
        formula_like, reasons = is_formula_like_text(text)
        contains_math = has_math(paragraph)
        if not formula_like and not contains_math:
            continue

        item = {
            "paragraph_index": index,
            "text": text.strip(),
            "has_omml": contains_math,
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

    result = "pass" if not pseudo_formula_paragraphs and not missing_formula_after_leadin else "fail"
    return {
        "schema": "graduation-project-builder.formula-object-audit.v1",
        "result": result,
        "docx_path": str(docx_path),
        "docx_sha256": sha256(docx_path),
        "math_object_count": len(math_objects),
        "math_paragraph_count": len(math_paragraphs),
        "formula_like_paragraph_count": len(formula_like_paragraphs),
        "real_formula_paragraph_count": len(real_formula_paragraphs),
        "pseudo_formula_count": len(pseudo_formula_paragraphs),
        "missing_formula_after_leadin_count": len(missing_formula_after_leadin),
        "pseudo_formula_paragraphs": pseudo_formula_paragraphs,
        "missing_formula_after_leadin": missing_formula_after_leadin,
        "real_formula_paragraphs": real_formula_paragraphs,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit DOCX formula authenticity and reject plain-text pseudo-formulas."
    )
    parser.add_argument("docx", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--allow-no-formulas", action="store_true")
    args = parser.parse_args()

    report = audit_docx(args.docx)
    if not args.allow_no_formulas and report["math_object_count"] == 0 and report["formula_like_paragraph_count"] == 0:
        report["result"] = "fail"
        report["missing_formula_evidence"] = "no formula-like text and no OMML formula objects found"

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
