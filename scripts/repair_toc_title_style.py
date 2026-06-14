"""Repair a thesis TOC title paragraph from an approved template donor.

This helper is intentionally narrow: it mutates only ``word/document.xml`` and
only the target paragraph whose visible text normalizes to ``目录`` / ``contents``.
It preserves the DOCX package and leaves TOC entries, fields, media, comments,
headers, footers, citations, and body prose untouched.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("w", NS["w"])

PPR_AFTER_JC_TAGS = {
    W + "textDirection",
    W + "textAlignment",
    W + "textboxTightWrap",
    W + "outlineLvl",
    W + "divId",
    W + "cnfStyle",
    W + "rPr",
    W + "sectPr",
    W + "pPrChange",
}

PPR_AFTER_PAGE_BREAK_BEFORE_TAGS = {
    W + "framePr",
    W + "widowControl",
    W + "numPr",
    W + "suppressLineNumbers",
    W + "pBdr",
    W + "shd",
    W + "tabs",
    W + "suppressAutoHyphens",
    W + "kinsoku",
    W + "wordWrap",
    W + "overflowPunct",
    W + "topLinePunct",
    W + "autoSpaceDE",
    W + "autoSpaceDN",
    W + "bidi",
    W + "adjustRightInd",
    W + "snapToGrid",
    W + "spacing",
    W + "ind",
    W + "contextualSpacing",
    W + "mirrorIndents",
    W + "suppressOverlap",
    W + "jc",
    W + "textDirection",
    W + "textAlignment",
    W + "textboxTightWrap",
    W + "outlineLvl",
    W + "divId",
    W + "cnfStyle",
    W + "rPr",
    W + "sectPr",
    W + "pPrChange",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def root_namespace_declarations(document_xml: bytes) -> dict[str, str]:
    match = re.search(rb"<w:document\b(?P<attrs>[^>]*)>", document_xml[:12000], flags=re.S)
    if not match:
        return {}
    attrs = match.group("attrs")
    return {
        prefix.decode("ascii", errors="ignore"): uri.decode("utf-8", errors="ignore")
        for prefix, uri in re.findall(rb'\sxmlns:([A-Za-z0-9]+)="([^"]+)"', attrs)
    }


def serialize_document(root: ET.Element, original_document_xml: bytes) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    declarations = root_namespace_declarations(original_document_xml)
    root_start = payload.find(b"<w:document")
    if root_start < 0:
        return payload
    root_end = payload.find(b">", root_start)
    if root_end < 0:
        return payload
    root_tag = payload[root_start:root_end]
    additions: list[bytes] = []
    for prefix, uri in declarations.items():
        token = f"xmlns:{prefix}=".encode("ascii")
        if token not in root_tag:
            additions.append(f' xmlns:{prefix}="{uri}"'.encode("utf-8"))
    if not additions:
        return payload
    return payload[:root_end] + b"".join(additions) + payload[root_end:]


def visible_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS)).strip()


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def title_paragraphs(root: ET.Element) -> list[tuple[int, ET.Element]]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    found: list[tuple[int, ET.Element]] = []
    for index, paragraph in enumerate(body.findall(".//w:p", NS)):
        if compact(visible_text(paragraph)) in {"\u76ee\u5f55", "contents"}:
            found.append((index, paragraph))
    return found


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.attrib.get(W + "val", "") if style is not None else ""


def ensure_center_alignment(ppr: ET.Element) -> None:
    for jc in list(ppr.findall(W + "jc")):
        ppr.remove(jc)
    jc = ET.Element(W + "jc", {W + "val": "center"})
    for index, child in enumerate(list(ppr)):
        if child.tag in PPR_AFTER_JC_TAGS:
            ppr.insert(index, jc)
            return
    ppr.append(jc)


def append_ppr_child_in_schema_order(ppr: ET.Element, element: ET.Element) -> None:
    if element.tag == W + "jc":
        for index, child in enumerate(list(ppr)):
            if child.tag in PPR_AFTER_JC_TAGS:
                ppr.insert(index, element)
                return
    ppr.append(element)


def ensure_page_break_before(ppr: ET.Element, donor: ET.Element | None) -> bool:
    if ppr.find(W + "pageBreakBefore") is not None:
        return True
    if donor is None:
        return False
    existing = donor.find(W + "pageBreakBefore")
    if existing is None:
        return False
    page_break = deepcopy(existing)
    for index, child in enumerate(list(ppr)):
        if child.tag in PPR_AFTER_PAGE_BREAK_BEFORE_TAGS:
            ppr.insert(index, page_break)
            return True
    ppr.append(page_break)
    return True


def remove_tabs(paragraph: ET.Element) -> int:
    removed = 0
    ppr = paragraph.find(W + "pPr")
    if ppr is not None:
        for tabs in list(ppr.findall(W + "tabs")):
            ppr.remove(tabs)
            removed += 1
    for tab in list(paragraph.findall(".//w:tab", NS)):
        parent = None
        for candidate in paragraph.iter():
            if tab in list(candidate):
                parent = candidate
                break
        if parent is not None:
            parent.remove(tab)
            removed += 1
    return removed


def replace_visible_title_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        run = ET.SubElement(paragraph, W + "r")
        text_nodes = [ET.SubElement(run, W + "t")]
    first = text_nodes[0]
    first.text = text
    if text.strip() != text:
        first.set(XML_SPACE, "preserve")
    else:
        first.attrib.pop(XML_SPACE, None)
    for node in text_nodes[1:]:
        node.text = ""


def donor_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall(".//w:r", NS):
        if visible_text(run):
            return run.find(W + "rPr")
    return None


def apply_run_rpr(paragraph: ET.Element, donor_rpr: ET.Element | None) -> None:
    if donor_rpr is None:
        return
    for run in paragraph.findall(".//w:r", NS):
        if not visible_text(run):
            continue
        for rpr in list(run.findall(W + "rPr")):
            run.remove(rpr)
        run.insert(0, deepcopy(donor_rpr))


def ensure_title_run_size(paragraph: ET.Element, half_points: str | None) -> int:
    if not half_points:
        return 0
    changed = 0
    for run in paragraph.findall(".//w:r", NS):
        if not visible_text(run):
            continue
        rpr = run.find(W + "rPr")
        if rpr is None:
            rpr = ET.Element(W + "rPr")
            run.insert(0, rpr)
        for tag_name in ("sz", "szCs"):
            element = rpr.find(W + tag_name)
            if element is None:
                element = ET.SubElement(rpr, W + tag_name)
            element.set(W + "val", half_points)
        changed += 1
    return changed


def repair_toc_title(
    *,
    template_docx: Path,
    docx_path: Path,
    output_path: Path,
    report_path: Path | None = None,
    title_size_half_points: str | None = None,
) -> dict[str, object]:
    with zipfile.ZipFile(template_docx, "r") as template_zip:
        template_root = ET.fromstring(template_zip.read("word/document.xml"))
    template_titles = title_paragraphs(template_root)
    if not template_titles:
        raise RuntimeError("template DOCX has no visible TOC title paragraph")
    template_index, template_title = template_titles[0]
    template_ppr = template_title.find(W + "pPr")
    template_text = visible_text(template_title) or "\u76ee\u5f55"
    template_style = paragraph_style_id(template_title)
    template_rpr = donor_run_rpr(template_title)

    shutil.copy2(docx_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    target_titles = title_paragraphs(root)
    if not target_titles:
        raise RuntimeError("target DOCX has no visible TOC title paragraph")
    target_index, target_title = target_titles[0]
    before_text = visible_text(target_title)
    before_style = paragraph_style_id(target_title)
    before_ppr = deepcopy(target_title.find(W + "pPr"))

    old_ppr = target_title.find(W + "pPr")
    if old_ppr is not None:
        target_title.remove(old_ppr)
    new_ppr = deepcopy(template_ppr) if template_ppr is not None else ET.Element(W + "pPr")
    target_title.insert(0, new_ppr)
    ensure_center_alignment(new_ppr)
    page_break_preserved = ensure_page_break_before(new_ppr, before_ppr)
    removed_tabs = remove_tabs(target_title)
    replace_visible_title_text(target_title, template_text)
    apply_run_rpr(target_title, template_rpr)
    sized_title_runs = ensure_title_run_size(target_title, title_size_half_points)

    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    report = {
        "schema": "graduation-project-builder.toc-title-style-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "template_docx": str(template_docx.resolve()),
        "template_docx_sha256": sha256_file(template_docx),
        "source_docx": str(docx_path.resolve()),
        "source_docx_sha256": sha256_file(docx_path),
        "output_docx": str(output_path.resolve()),
        "output_docx_sha256": sha256_file(output_path),
        "changed_zip_parts": ["word/document.xml"],
        "template_title_index": template_index,
        "target_title_index": target_index,
        "template_title_text": template_text,
        "before_title_text": before_text,
        "after_title_text": visible_text(target_title),
        "template_style_id": template_style,
        "before_style_id": before_style,
        "after_style_id": paragraph_style_id(target_title),
        "forced_alignment": "center",
        "page_break_before_preserved": page_break_preserved,
        "title_size_half_points": title_size_half_points or "",
        "sized_title_runs": sized_title_runs,
        "removed_title_tabs": removed_tabs,
        "verdict": "pass",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair TOC title style from a template donor.")
    parser.add_argument("--template-docx", required=True, type=Path)
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--title-size-half-points")
    args = parser.parse_args()
    report = repair_toc_title(
        template_docx=args.template_docx,
        docx_path=args.docx,
        output_path=args.output,
        report_path=args.report,
        title_size_half_points=args.title_size_half_points,
    )
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
