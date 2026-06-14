from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path


ROMAN_RE = re.compile(r"^(?:[ivxlcdm]+)$", re.I)
ARABIC_RE = re.compile(r"^\d+$")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def parse_page_set(value: str) -> set[int]:
    pages: set[int] = set()
    if not value:
        return pages
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start = int(left)
            end = int(right)
            pages.update(range(min(start, end), max(start, end) + 1))
        else:
            pages.add(int(part))
    return pages


def footer_text_for_page(page) -> str:
    page_height = float(page.rect.height)
    rows: list[tuple[float, float, str]] = []
    text_dict = page.get_text("dict")
    for block in text_dict.get("blocks", []):
        if not isinstance(block, dict):
            continue
        for line in block.get("lines", []):
            if not isinstance(line, dict):
                continue
            bbox = line.get("bbox", [0, 0, 0, 0])
            if not isinstance(bbox, (list, tuple)) or len(bbox) != 4:
                continue
            try:
                x0, y0, _x1, y1 = [float(value) for value in bbox]
            except (TypeError, ValueError):
                continue
            if y0 < page_height * 0.86:
                continue
            spans = line.get("spans", [])
            if not isinstance(spans, list):
                continue
            text = "".join(str(span.get("text", "")) for span in spans if isinstance(span, dict)).strip()
            text = re.sub(r"\s+", "", text)
            if text:
                rows.append((y1, x0, text))
    rows.sort(key=lambda row: (row[0], row[1]))
    numeric_rows = [text for _y, _x, text in rows if ROMAN_RE.fullmatch(text) or ARABIC_RE.fullmatch(text)]
    return numeric_rows[-1] if numeric_rows else ""


def audit_pdf_footers(
    pdf_path: Path,
    allow_blank_pages: set[int],
    roman_pages: set[int],
    arabic_start_page: int | None,
) -> dict[str, object]:
    import fitz  # type: ignore

    issues: list[str] = []
    footer_rows: list[dict[str, object]] = []
    with fitz.open(pdf_path) as doc:
        page_count = doc.page_count
        for index, page in enumerate(doc, start=1):
            footer = footer_text_for_page(page)
            expected = "arabic" if arabic_start_page is not None and index >= arabic_start_page else "roman"
            if index in allow_blank_pages:
                expected = "blank"
            status = "pass"
            if expected == "blank":
                if footer:
                    status = "fail"
                    issues.append(f"page {index} expected blank footer but found {footer}")
            elif expected == "roman":
                if not ROMAN_RE.fullmatch(footer):
                    status = "fail"
                    issues.append(f"page {index} expected roman footer but found {footer or 'missing'}")
            elif expected == "arabic":
                if not ARABIC_RE.fullmatch(footer):
                    status = "fail"
                    issues.append(f"page {index} expected arabic footer but found {footer or 'missing'}")
            footer_rows.append(
                {
                    "physical_page": index,
                    "expected": expected,
                    "footer": footer,
                    "status": status,
                }
            )

    if roman_pages:
        observed = [row["footer"] for row in footer_rows if int(row["physical_page"]) in roman_pages]
        if not all(ROMAN_RE.fullmatch(str(value)) for value in observed):
            issues.append("declared roman page range contains missing or non-roman footer values")
    if arabic_start_page is not None:
        observed_arabic = [
            int(row["footer"])
            for row in footer_rows
            if int(row["physical_page"]) >= arabic_start_page and ARABIC_RE.fullmatch(str(row["footer"]))
        ]
        expected_arabic = list(range(1, len(observed_arabic) + 1))
        if observed_arabic != expected_arabic:
            issues.append(
                f"arabic footer sequence is not continuous from 1: observed={observed_arabic[:20]}"
            )

    return {
        "schema": "graduation-project-builder.pdf-page-footer-audit.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "result": "pass" if not issues else "fail",
        "pdf_path": str(pdf_path.resolve()),
        "pdf_sha256": sha256_file(pdf_path),
        "page_count": len(footer_rows),
        "allow_blank_pages": sorted(allow_blank_pages),
        "roman_pages": sorted(roman_pages),
        "arabic_start_page": arabic_start_page,
        "footer_rows": footer_rows,
        "footer_mismatch_count": sum(1 for row in footer_rows if row["status"] != "pass"),
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit every rendered PDF page footer/page-number row.")
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--allow-blank-pages", default="")
    parser.add_argument("--roman-pages", default="")
    parser.add_argument("--arabic-start-page", type=int, default=None)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    report = audit_pdf_footers(
        args.pdf,
        allow_blank_pages=parse_page_set(args.allow_blank_pages),
        roman_pages=parse_page_set(args.roman_pages),
        arabic_start_page=args.arabic_start_page,
    )
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report["result"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
