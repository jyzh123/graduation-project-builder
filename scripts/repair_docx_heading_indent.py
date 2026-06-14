#!/usr/bin/env python3
"""Package-preserving repair for DOCX title/body-heading indentation residue."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

try:
    from lxml import etree as ET
except ImportError:  # pragma: no cover - fallback for minimal Python installs
    from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("v", "urn:schemas-microsoft-com:vml")

PPR_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]


def qn(local: str) -> str:
    return f"{W}{local}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def attr(element: ET.Element | None, local: str = "val") -> str:
    if element is None:
        return ""
    return element.attrib.get(qn(local), "")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def normalized(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def paragraph_style_id(paragraph: ET.Element) -> str:
    return attr(paragraph.find("./w:pPr/w:pStyle", NS))


def paragraph_has_page_break_before(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:pageBreakBefore", NS) is not None


def is_toc_title(text: str) -> bool:
    return compact(text) in {"目录", "contents", "tableofcontents"}


def is_zh_abstract(text: str) -> bool:
    return compact(re.sub(r"[:：].*$", "", text or "")) == "摘要"


def is_en_abstract(text: str) -> bool:
    return compact(re.sub(r"[:：].*$", "", text or "")).startswith("abstract")


def is_reference_title(text: str) -> bool:
    return compact(text) in {"参考文献", "references", "bibliography"}


def is_acknowledgement_title(text: str) -> bool:
    return compact(text) in {"致谢", "谢辞", "acknowledgements", "acknowledgments"}


def is_appendix_title(text: str) -> bool:
    return compact(text) in {"附录", "appendix"}


def is_chapter_heading_text(text: str) -> bool:
    return bool(re.match(r"^\s*第\s*[0-9一二三四五六七八九十]+\s*章(?:\s+\S.*)?\s*$", text or ""))


def body_heading_level(text: str) -> int | None:
    stripped = (text or "").strip()
    if is_chapter_heading_text(stripped):
        return 1
    if re.match(r"^[1-9]\d*(?:\.\d+){3}\s+\S", stripped):
        return 4
    if re.match(r"^[1-9]\d*\.\d+\.\d+\s+\S", stripped):
        return 3
    if re.match(r"^[1-9]\d*\.\d+\s+\S", stripped):
        return 2
    return None


def title_like_kind(text: str) -> str | None:
    if is_zh_abstract(text):
        return "zh_abstract_title"
    if is_en_abstract(text):
        return "en_abstract_title"
    if is_toc_title(text):
        return "toc_title"
    if is_reference_title(text):
        return "references_title"
    if is_acknowledgement_title(text):
        return "acknowledgement_title"
    if is_appendix_title(text):
        return "appendix_title"
    level = body_heading_level(text)
    if level:
        return f"body_heading_level_{level}"
    return None


def has_toc_page_number_text(text: str) -> bool:
    return bool(re.search(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", text or "") and ("\t" in (text or "") or "." in (text or "")))


def ensure_ppr(paragraph_or_style: ET.Element) -> ET.Element:
    ppr = paragraph_or_style.find("./w:pPr", NS)
    if ppr is not None:
        return ppr
    ppr = ET.Element(qn("pPr"))
    insert_at = 0
    if paragraph_or_style.tag == qn("style"):
        for index, child in enumerate(list(paragraph_or_style)):
            if local_name(child.tag) in {"name", "basedOn", "next", "link", "uiPriority", "qFormat", "rsid"}:
                insert_at = index + 1
    paragraph_or_style.insert(insert_at, ppr)
    return ppr


def insert_ppr_child(ppr: ET.Element, child: ET.Element) -> None:
    child_order = PPR_ORDER.index(local_name(child.tag)) if local_name(child.tag) in PPR_ORDER else len(PPR_ORDER)
    for index, existing in enumerate(list(ppr)):
        existing_order = PPR_ORDER.index(local_name(existing.tag)) if local_name(existing.tag) in PPR_ORDER else len(PPR_ORDER)
        if existing_order > child_order:
            ppr.insert(index, child)
            return
    ppr.append(child)


def ensure_ind(ppr: ET.Element) -> ET.Element:
    ind = ppr.find("./w:ind", NS)
    if ind is not None:
        return ind
    ind = ET.Element(qn("ind"))
    insert_ppr_child(ppr, ind)
    return ind


def clear_indent(ppr: ET.Element) -> dict[str, str]:
    ind = ensure_ind(ppr)
    before = {local_name(k): v for k, v in ind.attrib.items()}
    for local in ("left", "right", "leftChars", "rightChars", "firstLine", "firstLineChars", "hanging", "hangingChars"):
        ind.set(qn(local), "0")
    after = {local_name(k): v for k, v in ind.attrib.items()}
    return {"before": json.dumps(before, ensure_ascii=True, sort_keys=True), "after": json.dumps(after, ensure_ascii=True, sort_keys=True)}


def style_records(styles_root: ET.Element | None) -> dict[str, dict[str, str]]:
    records: dict[str, dict[str, str]] = {}
    if styles_root is None:
        return records
    for style in styles_root.findall("./w:style", NS):
        if attr(style, "type") != "paragraph":
            continue
        style_id = attr(style, "styleId")
        if not style_id:
            continue
        records[style_id] = {
            "style_id": style_id,
            "name": attr(style.find("./w:name", NS)),
            "based_on": attr(style.find("./w:basedOn", NS)),
            "default": attr(style, "default"),
        }
    return records


def default_paragraph_style_id(styles_root: ET.Element | None, records: dict[str, dict[str, str]]) -> str:
    if styles_root is not None:
        for style in styles_root.findall("./w:style", NS):
            if attr(style, "type") == "paragraph" and attr(style, "default").lower() in {"1", "true"}:
                return attr(style, "styleId")
    for style_id, record in records.items():
        if record.get("name", "").lower() == "normal":
            return style_id
    return ""


def style_chain(style_id: str, records: dict[str, dict[str, str]]) -> list[str]:
    chain: list[str] = []
    current = style_id
    seen: set[str] = set()
    while current and current in records and current not in seen:
        seen.add(current)
        chain.append(current)
        current = records[current].get("based_on", "")
    return chain


def style_looks_like_heading(style_id: str, records: dict[str, dict[str, str]]) -> bool:
    name = records.get(style_id, {}).get("name", "").lower()
    return "heading" in name or "标题" in name


def actual_body_heading_index(
    paragraphs: list[ET.Element],
    texts: list[str],
    records: dict[str, dict[str, str]],
) -> int | None:
    for index, paragraph in enumerate(paragraphs, start=1):
        text = texts[index - 1]
        if not body_heading_level(text):
            continue
        pstyle = paragraph_style_id(paragraph)
        if pstyle and any(style_looks_like_heading(style_id, records) for style_id in style_chain(pstyle, records)):
            return index
        if paragraph_has_page_break_before(paragraph) and not has_toc_page_number_text(text):
            return index
    return None


def is_normal_like(style_id: str, records: dict[str, dict[str, str]], default_style_id: str) -> bool:
    name = records.get(style_id, {}).get("name", "").lower()
    return style_id == default_style_id or name == "normal"


def repair(source_docx: Path, output_docx: Path) -> dict[str, object]:
    before_sha = sha256_file(source_docx)
    with zipfile.ZipFile(source_docx) as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
        infos = {info.filename: info for info in zin.infolist()}

    document_root = ET.fromstring(entries["word/document.xml"])
    styles_root = ET.fromstring(entries["word/styles.xml"]) if "word/styles.xml" in entries else None
    body = document_root.find("./w:body", NS)
    paragraphs = body.findall("./w:p", NS) if body is not None else []
    texts = [text_of(p).strip() for p in paragraphs]
    records = style_records(styles_root)
    default_style_id = default_paragraph_style_id(styles_root, records)

    style_ids_to_clear: set[str] = set()
    paragraph_repairs: list[dict[str, object]] = []
    style_repairs: list[dict[str, object]] = []

    first_body = actual_body_heading_index(paragraphs, texts, records)
    toc_title_idx = next((idx for idx, text in enumerate(texts, start=1) if is_toc_title(text)), None)
    references_idx = next((idx for idx, text in enumerate(texts, start=1) if is_reference_title(text)), None)
    acknowledgement_idx = next((idx for idx, text in enumerate(texts, start=1) if is_acknowledgement_title(text)), None)
    end_body = min([idx for idx in (references_idx, acknowledgement_idx) if idx is not None and first_body is not None and idx > first_body], default=len(paragraphs) + 1)

    for index, paragraph in enumerate(paragraphs, start=1):
        text = texts[index - 1]
        kind = title_like_kind(text)
        if not kind:
            continue
        if toc_title_idx is not None and first_body is not None and toc_title_idx < index < first_body:
            # TOC result paragraphs use indentation as alignment metadata; never
            # normalize them while repairing visual title/heading residue.
            continue
        if first_body is not None and index < first_body and not (
            is_zh_abstract(text) or is_en_abstract(text) or is_toc_title(text)
        ):
            # Static/live TOC entries have their own indentation semantics.
            if has_toc_page_number_text(text):
                continue
        if first_body is not None and index >= first_body and index < end_body and body_heading_level(text):
            pass
        elif kind.startswith("body_heading"):
            continue

        pstyle = paragraph_style_id(paragraph)
        if pstyle:
            for style_id in style_chain(pstyle, records):
                if is_normal_like(style_id, records, default_style_id):
                    continue
                style_ids_to_clear.add(style_id)
        else:
            ppr = ensure_ppr(paragraph)
            change = clear_indent(ppr)
            paragraph_repairs.append(
                {
                    "paragraph_index": index,
                    "kind": kind,
                    "text": normalized(text)[:100],
                    "repair": "direct paragraph w:ind zeroed because no paragraph style is bound",
                    **change,
                }
            )

        direct_ind = paragraph.find("./w:pPr/w:ind", NS)
        if direct_ind is not None and kind.startswith("body_heading"):
            ppr = ensure_ppr(paragraph)
            change = clear_indent(ppr)
            paragraph_repairs.append(
                {
                    "paragraph_index": index,
                    "kind": kind,
                    "text": normalized(text)[:100],
                    "repair": "direct body-heading w:ind zeroed",
                    **change,
                }
            )

    if styles_root is not None:
        styles_by_id = {attr(style, "styleId"): style for style in styles_root.findall("./w:style", NS)}
        for style_id in sorted(style_ids_to_clear):
            style = styles_by_id.get(style_id)
            if style is None:
                continue
            ppr = ensure_ppr(style)
            change = clear_indent(ppr)
            style_repairs.append(
                {
                    "style_id": style_id,
                    "style_name": records.get(style_id, {}).get("name", ""),
                    "repair": "style w:ind zeroed for title/body-heading effective indentation",
                    **change,
                }
            )

    new_entries = dict(entries)
    new_entries["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    if styles_root is not None:
        new_entries["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w") as zout:
        for name, data in new_entries.items():
            info = infos.get(name)
            if info is not None:
                clone = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
                clone.compress_type = info.compress_type
                clone.comment = info.comment
                clone.extra = info.extra
                clone.internal_attr = info.internal_attr
                clone.external_attr = info.external_attr
                clone.create_system = info.create_system
                zout.writestr(clone, data)
            else:
                zout.writestr(name, data)

    after_sha = sha256_file(output_docx)
    changed_parts = [
        name
        for name in ("word/document.xml", "word/styles.xml")
        if entries.get(name) != new_entries.get(name)
    ]
    return {
        "schema": "graduation-project-builder.docx-heading-indent-repair.v1",
        "source_docx_path": str(source_docx),
        "source_docx_sha256": before_sha,
        "output_docx_path": str(output_docx),
        "output_docx_sha256": after_sha,
        "changed_parts": changed_parts,
        "style_repairs": style_repairs,
        "paragraph_repairs": paragraph_repairs,
        "preservation_policy": "package-preserving copy; only word/document.xml and word/styles.xml may change",
        "passed": output_docx.exists() and bool(style_repairs or paragraph_repairs),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    args = parser.parse_args()
    report = repair(Path(args.source_docx), Path(args.output_docx))
    Path(args.report_json).write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
