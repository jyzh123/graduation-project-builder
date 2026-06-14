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


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CURRENT_IMUST_NAME = "\u5185\u8499\u53e4\u79d1\u6280\u5927\u5b66"
CURRENT_IMUST_BOOK = "\u672c\u79d1\u751f\u6bd5\u4e1a\u8bbe\u8ba1\u8bbe\u8ba1\u4e66"
CURRENT_COVER_LABELS = {
    "title": "\u9898\u76ee\uff1a",
    "student_name": "\u5b66\u751f\u59d3\u540d\uff1a",
    "student_id": "\u5b66\u53f7\uff1a",
    "major": "\u4e13\u4e1a\uff1a",
    "class_name": "\u73ed\u7ea7\uff1a",
    "advisor": "\u6307\u5bfc\u6559\u5e08\uff1a",
}
SECTION_CHILDREN_TO_STRIP = {
    W + "headerReference",
    W + "footerReference",
    W + "pgNumType",
}

ET.register_namespace("w", NS["w"])


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def compact_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("./w:r", NS):
        if paragraph_text(run).strip():
            rpr = run.find("./w:rPr", NS)
            return deepcopy(rpr) if rpr is not None else None
    return None


def set_paragraph_text(paragraph: ET.Element, value: str) -> None:
    donor_rpr = first_run_properties(paragraph)
    ppr = None
    for child in list(paragraph):
        if child.tag == W + "pPr":
            ppr = child
        paragraph.remove(child)
    if ppr is not None:
        paragraph.append(ppr)
    if value == "":
        return
    run = ET.Element(W + "r")
    if donor_rpr is not None:
        run.append(donor_rpr)
    text = ET.SubElement(run, W + "t")
    text.text = value
    if value.startswith(" ") or value.endswith(" ") or "  " in value:
        text.set(XML_SPACE, "preserve")
    paragraph.append(run)


def body_children(root: ET.Element) -> list[ET.Element]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml does not contain w:body")
    return list(body)


def child_sect_pr(child: ET.Element) -> ET.Element | None:
    if child.tag == W + "sectPr":
        return child
    if child.tag != W + "p":
        return None
    return child.find("./w:pPr/w:sectPr", NS)


def first_section_properties(children: list[ET.Element]) -> ET.Element:
    for child in children:
        sect_pr = child_sect_pr(child)
        if sect_pr is not None:
            return deepcopy(sect_pr)
    raise RuntimeError("could not locate a section properties donor")


def make_cover_section_properties(children: list[ET.Element]) -> ET.Element:
    sect_pr = first_section_properties(children)
    for child in list(sect_pr):
        if child.tag in SECTION_CHILDREN_TO_STRIP:
            sect_pr.remove(child)
    type_el = sect_pr.find("./w:type", NS)
    if type_el is None:
        type_el = ET.Element(W + "type")
        sect_pr.insert(0, type_el)
    type_el.set(W + "val", "nextPage")
    return sect_pr


def attach_section_properties(paragraph: ET.Element, sect_pr: ET.Element) -> None:
    if paragraph.tag != W + "p":
        raise RuntimeError("cover section boundary must be attached to a paragraph")
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        paragraph.insert(0, ppr)
    for old in list(ppr.findall("./w:sectPr", NS)):
        ppr.remove(old)
    ppr.append(sect_pr)


def find_standard_cover_template_paragraphs(template_root: ET.Element) -> list[ET.Element]:
    paragraphs = template_root.findall(".//w:p", NS)
    for index, paragraph in enumerate(paragraphs):
        if compact_text(paragraph_text(paragraph)) != CURRENT_IMUST_NAME:
            continue
        tail = [compact_text(paragraph_text(item)) for item in paragraphs[index : index + 12]]
        if CURRENT_IMUST_BOOK not in tail:
            continue
        if not any(item.startswith(CURRENT_COVER_LABELS["title"]) for item in tail):
            continue
        cover_start = index
        while cover_start > 0 and index - cover_start < 3 and not paragraph_text(paragraphs[cover_start - 1]).strip():
            cover_start -= 1
        cover_end = None
        for end_index, donor in enumerate(paragraphs[index + 1 :], start=index + 1):
            if compact_text(paragraph_text(donor)) == "\u6458\u8981":
                cover_end = end_index
                break
        if cover_end is not None and cover_end > cover_start:
            return [deepcopy(donor) for donor in paragraphs[cover_start:cover_end]]
        selected: list[ET.Element] = []
        for donor in paragraphs[index:]:
            selected.append(deepcopy(donor))
            if compact_text(paragraph_text(donor)).startswith(CURRENT_COVER_LABELS["advisor"]):
                break
        if len(selected) >= 6:
            return selected

    start = None
    for index, paragraph in enumerate(paragraphs):
        if compact_text(paragraph_text(paragraph)) == "内蒙古科技大学":
            tail = [compact_text(paragraph_text(item)) for item in paragraphs[index : index + 12]]
            if "本科生毕业设计说明书（毕业论文）" in tail and any(item.startswith("题目：") for item in tail):
                start = index
                break
    if start is None:
        raise RuntimeError("could not locate standard IMUST cover skeleton in template")
    selected: list[ET.Element] = []
    for paragraph in paragraphs[start:]:
        selected.append(deepcopy(paragraph))
        if compact_text(paragraph_text(paragraph)).startswith("指导教师："):
            break
    if len(selected) < 6:
        raise RuntimeError("standard cover skeleton is too short")
    return selected


def extract_field_value(text: str, label_patterns: list[str]) -> str:
    for pattern in label_patterns:
        match = re.match(pattern, text or "")
        if match:
            return match.group(1).strip()
    return ""


def extract_cover_values(final_children: list[ET.Element]) -> dict[str, str]:
    values = {
        "title": "",
        "student_name": "",
        "student_id": "",
        "major": "",
        "class_name": "",
        "advisor": "",
    }
    for child in final_children[:40]:
        if child.tag != W + "p":
            continue
        text = paragraph_text(child).strip()
        compact = compact_text(text)
        for key, label in CURRENT_COVER_LABELS.items():
            if not values[key] and compact.startswith(label):
                values[key] = compact.split("\uff1a", 1)[1].strip()
        if not values["title"]:
            title = extract_field_value(text, [r"^\s*题\s*目\s*[:：]\s*(.*)$"])
            if title:
                values["title"] = title
        if not values["student_name"]:
            values["student_name"] = extract_field_value(text, [r"^\s*学生姓名\s*[:：]\s*(.*)$", r"^\s*姓名\s*[:：]\s*(.*)$"])
        if not values["student_id"]:
            values["student_id"] = extract_field_value(text, [r"^\s*学\s*号\s*[:：]\s*(.*)$"])
        if not values["major"]:
            values["major"] = extract_field_value(text, [r"^\s*专\s*业\s*[:：]\s*(.*)$"])
        if not values["class_name"]:
            values["class_name"] = extract_field_value(text, [r"^\s*班\s*级\s*[:：]\s*(.*)$"])
        if not values["advisor"]:
            values["advisor"] = extract_field_value(text, [r"^\s*指导教师\s*[:：]\s*(.*)$"])
        if not values["title"] and "SGB620" in compact:
            values["title"] = text
    return values


def fill_cover_paragraph(paragraph: ET.Element, values: dict[str, str]) -> tuple[str, str] | None:
    text = paragraph_text(paragraph)
    compact = compact_text(text)
    current_replacements = {
        CURRENT_COVER_LABELS["title"]: f"              \u9898    \u76ee\uff1a{values['title']}",
        CURRENT_COVER_LABELS["student_name"]: f"              \u5b66\u751f\u59d3\u540d\uff1a{values['student_name']}",
        CURRENT_COVER_LABELS["student_id"]: f"              \u5b66    \u53f7\uff1a{values['student_id']}",
        CURRENT_COVER_LABELS["major"]: f"              \u4e13    \u4e1a\uff1a{values['major']}",
        CURRENT_COVER_LABELS["class_name"]: f"\u73ed    \u7ea7\uff1a {values['class_name']}",
        CURRENT_COVER_LABELS["advisor"]: f"              \u6307\u5bfc\u6559\u5e08\uff1a{values['advisor']}",
    }
    for key, replacement in current_replacements.items():
        if compact.startswith(key):
            before = text
            set_paragraph_text(paragraph, replacement)
            return before, replacement

    replacements = {
        "题目：": f"题    目：{values['title']}",
        "学生姓名：": f"学生姓名：{values['student_name']}",
        "学号：": f"学    号：{values['student_id']}",
        "专业：": f"专    业：{values['major']}",
        "班级：": f"班    级：{values['class_name']}",
        "指导教师：": f"指导教师：{values['advisor']}",
    }
    for key, replacement in replacements.items():
        if compact.startswith(compact_text(key)):
            before = text
            set_paragraph_text(paragraph, replacement)
            return before, replacement
    return None


def find_frontmatter_start_index(children: list[ET.Element]) -> int:
    for index, child in enumerate(children):
        if child.tag != W + "p":
            continue
        if compact_text(paragraph_text(child)) == "摘要":
            return index
    raise RuntimeError("could not locate Chinese abstract title after cover")


def repair_cover(input_docx: Path, template_docx: Path, output_docx: Path, report_path: Path | None) -> dict[str, object]:
    shutil.copy2(input_docx, output_docx)
    with zipfile.ZipFile(output_docx, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}
    with zipfile.ZipFile(template_docx, "r") as zin:
        template_document_xml = zin.read("word/document.xml")

    root = ET.fromstring(members["word/document.xml"])
    template_root = ET.fromstring(template_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml does not contain w:body")
    children = list(body)
    front_start = find_frontmatter_start_index(children)
    values = extract_cover_values(children)
    donors = find_standard_cover_template_paragraphs(template_root)
    cover_section_properties = make_cover_section_properties(children)

    field_updates: list[dict[str, str]] = []
    for donor in donors:
        update = fill_cover_paragraph(donor, values)
        if update is not None:
            field_updates.append({"before": update[0], "after": update[1]})
    attach_section_properties(donors[-1], cover_section_properties)

    for child in children[:front_start]:
        body.remove(child)
    for offset, donor in enumerate(donors):
        body.insert(offset, donor)

    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in members.items():
            zout.writestr(name, payload)

    report = {
        "schema": "graduation-project-builder.standard-cover-from-school-spec-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(input_docx.resolve()),
        "source_docx_sha256": sha256_file(input_docx),
        "template_docx": str(template_docx.resolve()),
        "template_docx_sha256": sha256_file(template_docx),
        "output_docx": str(output_docx.resolve()),
        "output_docx_sha256": sha256_file(output_docx),
        "changed_zip_parts": ["word/document.xml"],
        "cover_source": "official school format document standard-cover skeleton",
        "cover_section_isolated": True,
        "cover_section_header_footer_page_number_stripped": True,
        "replaced_cover_child_count": front_start,
        "inserted_cover_paragraph_count": len(donors),
        "field_values": values,
        "field_updates": field_updates,
        "verdict": "pass",
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Replay the official school-format standard cover skeleton onto a thesis DOCX.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report")
    args = parser.parse_args()
    report = repair_cover(
        input_docx=Path(args.input_docx),
        template_docx=Path(args.template_docx),
        output_docx=Path(args.output_docx),
        report_path=Path(args.report) if args.report else None,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
