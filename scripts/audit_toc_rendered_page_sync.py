from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from toc_leader_audit import audit_docx_toc_dotted_leaders, normalized_label


TAIL_LABELS = {
    "\u53c2\u8003\u6587\u732e",
    "\u81f4\u8c22",
}


def _compact(value: object) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", str(value or "")).strip().lower()


def _int_or_none(value: object) -> int | None:
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _body_start_page(rows: list[dict[str, object]]) -> int | None:
    candidates: list[int] = []
    for row in rows:
        page = _int_or_none(row.get("page"))
        if page is None:
            continue
        try:
            level = int(row.get("level") or 1)
        except Exception:
            level = 1
        if level != 1:
            continue
        text = str(row.get("text") or "").strip()
        compact = _compact(text)
        if compact in {_compact(label) for label in TAIL_LABELS}:
            continue
        if re.match(r"^\d{1,2}(?:\s|\u3000)+\S", text) or re.match(
            r"^\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u7ae0",
            text,
        ):
            candidates.append(page)
    if candidates:
        return min(candidates)
    fallback = [
        _int_or_none(row.get("page"))
        for row in rows
        if str(row.get("level") or "1").strip() == "1"
    ]
    fallback = [page for page in fallback if page is not None]
    return min(fallback) if fallback else None


def expected_logical_pages(rows: list[dict[str, object]]) -> tuple[dict[str, str], int | None]:
    body_start = _body_start_page(rows)
    result: dict[str, str] = {}
    for row in rows:
        page = _int_or_none(row.get("page"))
        if page is None:
            continue
        label = normalized_label(str(row.get("text") or ""))
        if not label:
            continue
        logical_page = page - body_start + 1 if body_start is not None else page
        if logical_page < 1:
            continue
        result[label] = str(logical_page)
    return result, body_start


def audit_toc_rendered_page_sync(docx_path: Path, heading_pages_json: Path) -> tuple[dict[str, object], list[str]]:
    toc_payload, toc_issues = audit_docx_toc_dotted_leaders(docx_path)
    try:
        rows = json.loads(heading_pages_json.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        return {}, [f"rendered heading page map could not be read: {exc}"]
    if not isinstance(rows, list):
        return {}, ["rendered heading page map must be a JSON list"]

    expected, body_start = expected_logical_pages([row for row in rows if isinstance(row, dict)])
    issues: list[str] = list(toc_issues)
    comparisons: list[dict[str, object]] = []
    entries = toc_payload.get("entries") if isinstance(toc_payload, dict) else []
    if not isinstance(entries, list):
        entries = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        label = normalized_label(str(entry.get("text") or ""))
        expected_page = expected.get(label)
        if expected_page is None:
            continue
        actual_page = str(entry.get("page_number") or "").strip()
        comparison = {
            "label": label,
            "toc_page_number": actual_page,
            "rendered_logical_page": expected_page,
            "matched": actual_page == expected_page,
            "paragraph_index": entry.get("paragraph_index"),
        }
        comparisons.append(comparison)
        if actual_page != expected_page:
            issues.append(
                "TOC rendered page sync mismatch: "
                f"label={label!r} toc={actual_page!r} rendered={expected_page!r}"
            )

    for tail_label in TAIL_LABELS:
        label = normalized_label(tail_label)
        toc_has_tail = any(
            isinstance(entry, dict) and normalized_label(str(entry.get("text") or "")) == label
            for entry in entries
        )
        if toc_has_tail and label not in expected:
            issues.append(f"TOC tail-block rendered page missing from heading map: {tail_label}")

    payload: dict[str, object] = {
        "schema": "graduation-project-builder.toc-rendered-page-sync.v1",
        "docx_path": str(docx_path),
        "heading_pages_json": str(heading_pages_json),
        "body_start_physical_page": body_start,
        "passed": not issues,
        "comparison_count": len(comparisons),
        "comparisons": comparisons,
        "issues": issues,
    }
    return payload, issues


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--heading-pages-json", required=True)
    parser.add_argument("--report-json")
    args = parser.parse_args()
    payload, issues = audit_toc_rendered_page_sync(
        Path(args.docx).resolve(),
        Path(args.heading_pages_json).resolve(),
    )
    if args.report_json:
        output = Path(args.report_json).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "passed": not issues,
                "issue_count": len(issues),
                "comparison_count": payload.get("comparison_count", 0) if payload else 0,
            },
            ensure_ascii=False,
        )
    )
    return 0 if not issues else 1


if __name__ == "__main__":
    raise SystemExit(main())
