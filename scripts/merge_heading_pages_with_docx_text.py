#!/usr/bin/env python3
"""Merge rendered heading page numbers with clean DOCX heading text."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(W + "val", "") if style is not None else ""


def collect_docx_heading_rows(docx: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    rows: list[dict[str, object]] = []
    started = False
    for paragraph in root.findall("./w:body/w:p", NS):
        text = paragraph_text(paragraph)
        if not text:
            continue
        style_id = paragraph_style_id(paragraph)
        if text.startswith("第1章"):
            started = True
        if not started:
            continue
        level: int | None = None
        if style_id in {"2", "Heading1"} or text in {"参考文献", "附录", "致谢"}:
            level = 1
        elif style_id in {"3", "Heading2"}:
            level = 2
        elif style_id in {"4", "Heading3"}:
            level = 3
        if level is not None:
            rows.append({"text": text, "level": level, "style": style_id})
    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--heading-pages", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    docx_rows = collect_docx_heading_rows(args.docx.resolve())
    rendered_rows = json.loads(args.heading_pages.read_text(encoding="utf-8-sig"))
    rendered_rows = [
        row for row in rendered_rows
        if isinstance(row, dict) and isinstance(row.get("page"), int)
    ]
    if len(docx_rows) != len(rendered_rows):
        raise RuntimeError(
            f"heading count mismatch: docx={len(docx_rows)} rendered={len(rendered_rows)}"
        )
    merged = []
    for docx_row, rendered_row in zip(docx_rows, rendered_rows):
        merged.append(
            {
                "text": docx_row["text"],
                "page": int(rendered_row["page"]),
                "level": int(docx_row["level"]),
                "style": str(docx_row["style"]),
                "rendered_text": str(rendered_row.get("text") or ""),
            }
        )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(args.output.resolve()), "count": len(merged)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
