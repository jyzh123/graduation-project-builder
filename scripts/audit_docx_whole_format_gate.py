#!/usr/bin/env python3
"""Fail-closed structural format audit for whole-thesis DOCX handoff."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def attr(element: ET.Element | None, local: str = "val") -> str:
    if element is None:
        return ""
    return element.attrib.get(qn(local), "")


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def paragraph_style_id(paragraph: ET.Element) -> str:
    return attr(paragraph.find("./w:pPr/w:pStyle", NS))


def has_page_break(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:br[@w:type='page']", NS) is not None


def has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def has_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tab", NS) is not None


def has_dotted_tab(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:tabs/w:tab[@w:leader='dot']", NS) is not None


def trailing_page_number(text: str) -> str:
    match = re.search(r"(\d+|[ivxlcdmIVXLCDM]+)\s*$", text or "")
    return match.group(1) if match else ""


def is_toc_title(text: str) -> bool:
    return compact(text) in {"目录", "目錄", "contents", "tableofcontents"}


def is_zh_abstract(text: str) -> bool:
    value = compact(re.sub(r"[:：].*$", "", text or ""))
    return value == "摘要"


def is_en_abstract(text: str) -> bool:
    return compact(re.sub(r"[:：].*$", "", text or "")).startswith("abstract")


def is_body_heading(text: str) -> bool:
    stripped = (text or "").strip()
    return bool(
        re.match(r"^第[0-9一二三四五六七八九十]+章", stripped)
        or re.match(r"^[1-9]\d*(?:\.\d+){0,2}\s+\S", stripped)
    )


def load_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except (KeyError, ET.ParseError):
        return None


def style_name_map(styles_root: ET.Element | None) -> dict[str, str]:
    if styles_root is None:
        return {}
    result: dict[str, str] = {}
    for style in styles_root.findall("./w:style", NS):
        style_id = attr(style, "styleId")
        name = attr(style.find("./w:name", NS))
        if style_id:
            result[style_id] = name
    return result


def visible_paragraphs(document_root: ET.Element) -> list[ET.Element]:
    body = document_root.find("./w:body", NS)
    if body is None:
        return []
    return body.findall("./w:p", NS)


def relationship_targets(root: ET.Element | None) -> dict[str, str]:
    if root is None:
        return {}
    targets: dict[str, str] = {}
    for rel in list(root):
        rid = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rid:
            targets[rid] = target
    return targets


def part_text_and_fields(zf: zipfile.ZipFile, part_name: str) -> tuple[str, str]:
    root = load_xml(zf, f"word/{part_name}")
    if root is None:
        return "", ""
    return (
        "".join(node.text or "" for node in root.findall(".//w:t", NS)),
        " ".join(
            [
                *(node.text or "" for node in root.findall(".//w:instrText", NS)),
                *(node.attrib.get(W + "instr", "") for node in root.findall(".//w:fldSimple", NS)),
            ]
        ),
    )


def audit_docx(path: Path, *, allow_builder_styles: bool = False, require_toc_field: bool = False) -> dict[str, object]:
    issues: list[str] = []
    with zipfile.ZipFile(path) as zf:
        document_root = load_xml(zf, "word/document.xml")
        if document_root is None:
            raise ValueError("word/document.xml is missing or invalid")
        styles_root = load_xml(zf, "word/styles.xml")
        rel_root = load_xml(zf, "word/_rels/document.xml.rels")
        names = set(zf.namelist())

        paragraphs = visible_paragraphs(document_root)
        texts = [text_of(p).strip() for p in paragraphs]
        styles = style_name_map(styles_root)
        used_styles: dict[str, int] = {}
        nonempty_no_style = 0
        body_no_style = 0
        body_count = 0
        builder_style_count = 0
        body_started = False
        toc_started = False
        toc_entry_count = 0
        toc_entry_with_page_count = 0
        toc_entry_with_dotted_leader_count = 0
        toc_entry_with_tab_count = 0
        page_break_count = 0
        paragraph_section_break_count = 0
        field_instructions: list[str] = []

        surface_indices = {
            "zh_abstract": None,
            "en_abstract": None,
            "toc": None,
            "first_body": None,
            "references": None,
            "appendix": None,
            "acknowledgement": None,
        }

        for index, paragraph in enumerate(paragraphs, start=1):
            text = texts[index - 1]
            style_id = paragraph_style_id(paragraph)
            if style_id:
                used_styles[style_id] = used_styles.get(style_id, 0) + 1
            elif text:
                nonempty_no_style += 1
            style_name = styles.get(style_id, style_id)
            if text and (style_id.lower().startswith("sgb") or style_name.lower().startswith("sgb")):
                builder_style_count += 1
            if has_page_break(paragraph):
                page_break_count += 1
            if has_section_break(paragraph):
                paragraph_section_break_count += 1
            instr = " ".join(
                [
                    *(node.text or "" for node in paragraph.findall(".//w:instrText", NS)),
                    *(node.attrib.get(W + "instr", "") for node in paragraph.findall(".//w:fldSimple", NS)),
                ]
            ).strip()
            if instr:
                field_instructions.append(instr)

            if surface_indices["zh_abstract"] is None and is_zh_abstract(text):
                surface_indices["zh_abstract"] = index
            if surface_indices["en_abstract"] is None and is_en_abstract(text):
                surface_indices["en_abstract"] = index
            if surface_indices["toc"] is None and is_toc_title(text):
                surface_indices["toc"] = index
                toc_started = True
                continue
            if toc_started and not body_started:
                if is_body_heading(text):
                    surface_indices["first_body"] = index
                    body_started = True
                    toc_started = False
                elif text:
                    maybe_toc_entry = has_tab(paragraph) or trailing_page_number(text) or re.match(r"^(\d+(\.\d+){0,2}|第.+章)", text)
                    if maybe_toc_entry:
                        toc_entry_count += 1
                        if trailing_page_number(text):
                            toc_entry_with_page_count += 1
                        if has_dotted_tab(paragraph):
                            toc_entry_with_dotted_leader_count += 1
                        if has_tab(paragraph):
                            toc_entry_with_tab_count += 1
            elif not body_started and surface_indices["first_body"] is None and is_body_heading(text):
                surface_indices["first_body"] = index
                body_started = True

            if body_started and text:
                body_count += 1
                if not style_id:
                    body_no_style += 1

            if surface_indices["references"] is None and compact(text) == "参考文献":
                surface_indices["references"] = index
            if surface_indices["appendix"] is None and compact(text).startswith("附录"):
                surface_indices["appendix"] = index
            if surface_indices["acknowledgement"] is None and compact(text) in {"致谢", "致謝"}:
                surface_indices["acknowledgement"] = index

        section_properties = document_root.findall(".//w:sectPr", NS)
        section_count = len(section_properties)
        rels = relationship_targets(rel_root)
        header_refs: list[str] = []
        footer_refs: list[str] = []
        page_num_types: list[dict[str, str]] = []
        for sect_pr in section_properties:
            header_refs.extend(ref.attrib.get(R + "id", "") for ref in sect_pr.findall("./w:headerReference", NS))
            footer_refs.extend(ref.attrib.get(R + "id", "") for ref in sect_pr.findall("./w:footerReference", NS))
            pg_num = sect_pr.find("./w:pgNumType", NS)
            page_num_types.append({"fmt": attr(pg_num, "fmt"), "start": attr(pg_num, "start")})

        header_parts = sorted({rels.get(rid, "") for rid in header_refs if rels.get(rid, "")})
        footer_parts = sorted({rels.get(rid, "") for rid in footer_refs if rels.get(rid, "")})
        header_texts: list[str] = []
        header_fields: list[str] = []
        footer_texts: list[str] = []
        footer_fields: list[str] = []
        for part in header_parts:
            text, fields = part_text_and_fields(zf, part)
            header_texts.append(text)
            header_fields.append(fields)
        for part in footer_parts:
            text, fields = part_text_and_fields(zf, part)
            footer_texts.append(text)
            footer_fields.append(fields)

        live_toc_count = sum(1 for instr in field_instructions if "TOC" in instr.upper())
        footer_page_field_count = sum(1 for instr in footer_fields if "PAGE" in instr.upper())

    for name, index in surface_indices.items():
        if name in {"appendix"}:
            continue
        if index is None:
            issues.append(f"missing required thesis surface: {name}")

    ordered_required = [
        surface_indices["zh_abstract"],
        surface_indices["en_abstract"],
        surface_indices["toc"],
        surface_indices["first_body"],
        surface_indices["references"],
        surface_indices["acknowledgement"],
    ]
    present_order = [value for value in ordered_required if value is not None]
    if present_order != sorted(present_order):
        issues.append(f"front/body/end matter order is wrong: {surface_indices}")

    if section_count < 3:
        issues.append(
            "whole-thesis DOCX must have separate section topology for cover, front matter, and body/end matter"
        )
    if section_count <= 1 and any(text.strip() for text in header_texts):
        issues.append("single-section thesis applies a running header to the cover/front matter")

    if require_toc_field and live_toc_count <= 0:
        issues.append("live TOC field is required but no TOC field instruction was found")
    if live_toc_count <= 0 and toc_entry_with_dotted_leader_count <= 0:
        issues.append("TOC lacks both a live TOC field and dotted-leader TOC entries")
    if toc_entry_count > 0 and toc_entry_with_page_count <= 0 and live_toc_count <= 0:
        issues.append("static TOC-like heading list lacks visible page numbers")
    if toc_entry_count > 0 and toc_entry_with_tab_count <= 0 and live_toc_count <= 0:
        issues.append("static TOC-like heading list lacks tab-separated entry/page-number structure")

    if not footer_parts:
        issues.append("no footer part is bound to the document sections")
    if footer_page_field_count <= 0:
        issues.append("footer/page-number surface lacks a PAGE field")
    if section_count >= 3:
        front_has_roman = any(item.get("fmt", "").lower() in {"roman", "upperroman", "lowerroman"} for item in page_num_types)
        body_restarts = any(item.get("start") == "1" or item.get("fmt", "").lower() in {"decimal", ""} for item in page_num_types[1:])
        if not front_has_roman:
            issues.append("front matter section does not expose a roman page-number format")
        if not body_restarts:
            issues.append("body section does not expose a restarted Arabic page-number chain")

    if builder_style_count and not allow_builder_styles:
        issues.append(f"builder-owned thesis styles are still used in visible content: {builder_style_count} paragraphs")
    if builder_style_count and body_count and body_no_style / max(body_count, 1) > 0.20:
        issues.append(f"too many body/end-matter paragraphs lack explicit style binding: {body_no_style}/{body_count}")

    return {
        "schema": "graduation-project-builder.docx-whole-format-gate.v1",
        "generator": "scripts/audit_docx_whole_format_gate.py",
        "docx_path": str(path),
        "docx_sha256": sha256_file(path),
        "counts": {
            "paragraph_count": len(paragraphs),
            "section_count": section_count,
            "paragraph_section_break_count": paragraph_section_break_count,
            "page_break_count": page_break_count,
            "live_toc_field_count": live_toc_count,
            "toc_entry_count": toc_entry_count,
            "toc_entry_with_page_count": toc_entry_with_page_count,
            "toc_entry_with_tab_count": toc_entry_with_tab_count,
            "toc_entry_with_dotted_leader_count": toc_entry_with_dotted_leader_count,
            "header_part_count": len(header_parts),
            "footer_part_count": len(footer_parts),
            "footer_page_field_count": footer_page_field_count,
            "builder_style_visible_paragraph_count": builder_style_count,
            "nonempty_no_style_paragraph_count": nonempty_no_style,
            "body_no_style_paragraph_count": body_no_style,
            "body_paragraph_count": body_count,
        },
        "surfaces": surface_indices,
        "used_styles": used_styles,
        "header_parts": header_parts,
        "footer_parts": footer_parts,
        "header_texts": header_texts,
        "footer_texts": footer_texts,
        "page_number_types": page_num_types,
        "issues": issues,
        "passed": not issues,
    }


def write_report(report: dict[str, object], path: Path | None) -> None:
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if path:
        path.write_text(text, encoding="utf-8")
    print(text)


def add_field_run(paragraph, instr: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as docx_qn

    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(docx_qn("w:fldCharType"), "begin")
    run._r.append(begin)
    run = paragraph.add_run()
    instr_text = OxmlElement("w:instrText")
    instr_text.set(docx_qn("xml:space"), "preserve")
    instr_text.text = instr
    run._r.append(instr_text)
    run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(docx_qn("w:fldCharType"), "separate")
    run._r.append(separate)
    paragraph.add_run("1")
    run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(docx_qn("w:fldCharType"), "end")
    run._r.append(end)


def set_section_page_number(section, *, fmt: str, start: int) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as docx_qn

    sect_pr = section._sectPr
    for existing in list(sect_pr.findall(docx_qn("w:pgNumType"))):
        sect_pr.remove(existing)
    pg_num = OxmlElement("w:pgNumType")
    pg_num.set(docx_qn("w:fmt"), fmt)
    pg_num.set(docx_qn("w:start"), str(start))
    sect_pr.append(pg_num)


def add_styled_paragraph(document, text: str, style_name: str):
    paragraph = document.add_paragraph(text)
    paragraph.style = style_name
    return paragraph


def self_test() -> int:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.style import WD_STYLE_TYPE

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        bad = tmp / "bad.docx"
        doc = Document()
        doc.add_paragraph("内蒙古科技大学")
        doc.add_paragraph("摘 要")
        doc.add_paragraph("Abstract")
        doc.add_paragraph("目 录")
        doc.add_paragraph("第1章 绪论")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("致 谢")
        doc.save(bad)
        bad_report = audit_docx(bad)
        if bad_report["passed"]:
            print("bad fixture unexpectedly passed", file=sys.stderr)
            return 1

        good = tmp / "good.docx"
        doc = Document()
        styles = doc.styles
        for style_name in ("GPBBody", "GPBHeading", "GPBFront"):
            if style_name not in styles:
                styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        doc.add_paragraph("内蒙古科技大学")
        doc.add_section(WD_SECTION.NEW_PAGE)
        set_section_page_number(doc.sections[1], fmt="roman", start=1)
        add_styled_paragraph(doc, "摘 要", "GPBFront")
        add_styled_paragraph(doc, "Abstract", "GPBFront")
        add_styled_paragraph(doc, "目 录", "GPBFront")
        toc = add_styled_paragraph(doc, "", "GPBFront")
        add_field_run(toc, ' TOC \\o "1-3" \\h \\z \\u ')
        doc.add_section(WD_SECTION.NEW_PAGE)
        set_section_page_number(doc.sections[2], fmt="decimal", start=1)
        add_styled_paragraph(doc, "第1章 绪论", "GPBHeading")
        add_styled_paragraph(doc, "正文内容。", "GPBBody")
        add_styled_paragraph(doc, "参考文献", "GPBHeading")
        add_styled_paragraph(doc, "[1] Source.", "GPBBody")
        add_styled_paragraph(doc, "致 谢", "GPBHeading")
        for section in doc.sections[1:]:
            section.header.is_linked_to_previous = False
            section.header.paragraphs[0].text = "内蒙古科技大学毕业设计说明书（毕业论文）"
            section.footer.is_linked_to_previous = False
            add_field_run(section.footer.paragraphs[0], " PAGE ")
        doc.save(good)
        good_report = audit_docx(good, require_toc_field=True)
        if not good_report["passed"]:
            print(json.dumps(good_report, ensure_ascii=False, indent=2), file=sys.stderr)
            return 1
    print("self-test passed")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx_path", nargs="?")
    parser.add_argument("--report-json")
    parser.add_argument("--allow-builder-styles", action="store_true")
    parser.add_argument("--require-toc-field", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test:
        return self_test()
    if not args.docx_path:
        parser.error("docx_path is required unless --self-test is used")
    report = audit_docx(
        Path(args.docx_path),
        allow_builder_styles=args.allow_builder_styles,
        require_toc_field=args.require_toc_field,
    )
    write_report(report, Path(args.report_json) if args.report_json else None)
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
