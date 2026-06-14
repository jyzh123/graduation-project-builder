#!/usr/bin/env python3
"""Center live PAGE fields in active thesis footer parts.

The default repair is bounded to footer parts referenced by non-title-page
sections that carry page numbering. When explicitly requested, the helper may
create one missing default footer relationship on the final section and copy
template margin geometry into that section. It does not touch body text,
headers, images, styles, numbering definitions, comments, or media parts.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
REL = f"{{{REL_NS}}}"
CT = f"{{{CT_NS}}}"
XML = f"{{{XML_NS}}}"
NS = {"w": W_NS, "r": R_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", REL_NS)


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"


def rels_map(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return {}
    return {
        rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
        for rel in root.findall(f"{REL}Relationship")
    }


def referenced_default_page_footer_parts(zf: zipfile.ZipFile) -> set[str]:
    document = ET.fromstring(zf.read("word/document.xml"))
    rels = rels_map(zf)
    body = document.find("./w:body", NS)
    if body is None:
        return set()
    sections = []
    for paragraph in body.findall("./w:p", NS):
        sect_pr = paragraph.find("./w:pPr/w:sectPr", NS)
        if sect_pr is not None:
            sections.append(sect_pr)
    body_sect = body.find("./w:sectPr", NS)
    if body_sect is not None:
        sections.append(body_sect)

    parts: set[str] = set()
    for sect_pr in sections:
        pg_num = sect_pr.find("./w:pgNumType", NS)
        if pg_num is None:
            continue
        if sect_pr.find("./w:titlePg", NS) is not None and not pg_num.attrib.get(qn("start")):
            continue
        for ref in sect_pr.findall("./w:footerReference", NS):
            target = rels.get(ref.attrib.get(f"{R}id", ""), "")
            if target:
                parts.add("word/" + target.lstrip("/"))
    return parts


def footer_part_has_page_field(zf: zipfile.ZipFile, part_name: str) -> bool:
    try:
        root = ET.fromstring(zf.read(part_name))
    except (KeyError, ET.ParseError):
        return False
    for node in root.findall(".//w:instrText", NS):
        if "PAGE" in (node.text or "").upper():
            return True
    for node in root.findall(".//w:fldSimple", NS):
        if "PAGE" in node.get(qn("instr"), "").upper():
            return True
    return False


def relationship_root(zf: zipfile.ZipFile) -> ET.Element:
    try:
        return ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return ET.Element(f"{REL}Relationships")


def content_types_root(zf: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(zf.read("[Content_Types].xml"))


def document_root(zf: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(zf.read("word/document.xml"))


def section_properties(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        return []
    sections: list[ET.Element] = []
    for paragraph in body.findall("./w:p", NS):
        sect_pr = paragraph.find("./w:pPr/w:sectPr", NS)
        if sect_pr is not None:
            sections.append(sect_pr)
    body_sect = body.find("./w:sectPr", NS)
    if body_sect is not None:
        sections.append(body_sect)
    return sections


def final_section(root: ET.Element) -> ET.Element | None:
    sections = section_properties(root)
    return sections[-1] if sections else None


def next_numeric_id(existing: set[str], prefix: str) -> str:
    max_seen = 0
    for value in existing:
        if value.startswith(prefix):
            suffix = value[len(prefix) :]
            if suffix.isdigit():
                max_seen = max(max_seen, int(suffix))
    return f"{prefix}{max_seen + 1}"


def next_relationship_id(rels_root: ET.Element) -> str:
    existing = {rel.attrib.get("Id", "") for rel in rels_root.findall(f"{REL}Relationship")}
    return next_numeric_id(existing, "rId")


def existing_footer_part_names(zf: zipfile.ZipFile) -> set[str]:
    return {
        name
        for name in zf.namelist()
        if name.startswith("word/footer") and name.endswith(".xml")
    }


def next_footer_part_name(zf: zipfile.ZipFile) -> str:
    existing_numbers = {
        name.removeprefix("word/footer").removesuffix(".xml")
        for name in existing_footer_part_names(zf)
    }
    next_name = next_numeric_id({f"footer{num}" for num in existing_numbers if num.isdigit()}, "footer")
    return f"word/{next_name}.xml"


def template_portrait_pg_mar(template_docx: Path | None) -> ET.Element | None:
    if template_docx is None:
        return None
    with zipfile.ZipFile(template_docx, "r") as zf:
        root = document_root(zf)
    for sect_pr in section_properties(root):
        pg_sz = sect_pr.find("./w:pgSz", NS)
        if pg_sz is not None and pg_sz.get(qn("orient")) == "landscape":
            continue
        pg_mar = sect_pr.find("./w:pgMar", NS)
        if pg_mar is not None:
            return ET.fromstring(ET.tostring(pg_mar, encoding="utf-8"))
    return None


def replace_or_insert_section_child(sect_pr: ET.Element, child: ET.Element, tag: str) -> bool:
    existing = sect_pr.find(f"./w:{tag}", NS)
    if existing is not None:
        if ET.tostring(existing, encoding="utf-8") == ET.tostring(child, encoding="utf-8"):
            return False
        index = list(sect_pr).index(existing)
        sect_pr.remove(existing)
        sect_pr.insert(index, child)
        return True
    insert_after_tags = {"pgSz", "type"} if tag == "pgMar" else set()
    insert_at = 0
    for idx, node in enumerate(list(sect_pr)):
        local = node.tag.removeprefix(W)
        if local in insert_after_tags:
            insert_at = idx + 1
    sect_pr.insert(insert_at, child)
    return True


def ensure_pg_num_type(sect_pr: ET.Element, *, fmt: str = "decimal", start: str | None = None) -> bool:
    pg_num = sect_pr.find("./w:pgNumType", NS)
    changed = False
    if pg_num is None:
        pg_num = ET.Element(qn("pgNumType"))
        insert_at = 0
        for idx, node in enumerate(list(sect_pr)):
            local = node.tag.removeprefix(W)
            if local in {"pgSz", "pgMar", "pgBorders", "lnNumType"}:
                insert_at = idx + 1
        sect_pr.insert(insert_at, pg_num)
        changed = True
    if pg_num.get(qn("fmt")) != fmt:
        pg_num.set(qn("fmt"), fmt)
        changed = True
    if start is None:
        if qn("start") in pg_num.attrib:
            del pg_num.attrib[qn("start")]
            changed = True
    elif pg_num.get(qn("start")) != start:
        pg_num.set(qn("start"), start)
        changed = True
    return changed


def ensure_default_footer_reference(sect_pr: ET.Element, rel_id: str) -> bool:
    for ref in sect_pr.findall("./w:footerReference", NS):
        if ref.get(qn("type"), "default") == "default":
            if ref.get(f"{R}id") == rel_id:
                return False
            ref.set(f"{R}id", rel_id)
            return True
    ref = ET.Element(qn("footerReference"))
    ref.set(qn("type"), "default")
    ref.set(f"{R}id", rel_id)
    insert_at = 0
    for idx, node in enumerate(list(sect_pr)):
        local = node.tag.removeprefix(W)
        if local in {"headerReference", "footerReference"}:
            insert_at = idx + 1
    sect_pr.insert(insert_at, ref)
    return True


def ensure_footer_relationship(rels_root: ET.Element, footer_part: str) -> tuple[str, bool]:
    target = footer_part.removeprefix("word/")
    for rel in rels_root.findall(f"{REL}Relationship"):
        if rel.attrib.get("Type") == FOOTER_REL_TYPE and rel.attrib.get("Target") == target:
            rel_id = rel.attrib.get("Id", "")
            if rel_id:
                return rel_id, False
    rel_id = next_relationship_id(rels_root)
    rel = ET.SubElement(rels_root, f"{REL}Relationship")
    rel.set("Id", rel_id)
    rel.set("Type", FOOTER_REL_TYPE)
    rel.set("Target", target)
    return rel_id, True


def ensure_footer_content_type(content_root: ET.Element, footer_part: str) -> bool:
    part_name = "/" + footer_part
    for override in content_root.findall(f"{CT}Override"):
        if override.attrib.get("PartName") == part_name:
            if override.attrib.get("ContentType") == FOOTER_CONTENT_TYPE:
                return False
            override.set("ContentType", FOOTER_CONTENT_TYPE)
            return True
    override = ET.SubElement(content_root, f"{CT}Override")
    override.set("PartName", part_name)
    override.set("ContentType", FOOTER_CONTENT_TYPE)
    return True


def serialize_relationships(root: ET.Element) -> bytes:
    ET.register_namespace("", REL_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_content_types(root: ET.Element) -> bytes:
    ET.register_namespace("", CT_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_word_xml(root: ET.Element) -> bytes:
    ET.register_namespace("w", W_NS)
    ET.register_namespace("r", R_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def clone_element(element: ET.Element | None) -> ET.Element | None:
    if element is None:
        return None
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def paragraph_has_page_field(paragraph: ET.Element) -> bool:
    for node in paragraph.findall(".//w:instrText", NS):
        if "PAGE" in (node.text or "").upper():
            return True
    for node in paragraph.findall(".//w:fldSimple", NS):
        if "PAGE" in node.get(qn("instr"), "").upper():
            return True
    return False


def first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall(".//w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is not None:
            return clone_element(rpr)
    return None


def ensure_center_alignment(ppr: ET.Element) -> None:
    jc = ppr.find("./w:jc", NS)
    if jc is None:
        jc = ET.SubElement(ppr, qn("jc"))
    jc.set(qn("val"), "center")


def default_page_run_properties() -> ET.Element:
    rpr = ET.Element(qn("rPr"))
    sz = ET.SubElement(rpr, qn("sz"))
    sz.set(qn("val"), "21")
    sz_cs = ET.SubElement(rpr, qn("szCs"))
    sz_cs.set(qn("val"), "21")
    return rpr


def footer_baseline_properties(zf: zipfile.ZipFile, part_name: str) -> tuple[ET.Element | None, ET.Element | None]:
    try:
        root = ET.fromstring(zf.read(part_name))
    except (KeyError, ET.ParseError):
        return None, None
    fallback_ppr: ET.Element | None = None
    fallback_rpr: ET.Element | None = None
    for paragraph in root.findall(".//w:p", NS):
        if fallback_ppr is None:
            fallback_ppr = clone_element(paragraph.find("./w:pPr", NS))
        if fallback_rpr is None:
            fallback_rpr = first_run_properties(paragraph)
        if paragraph_has_page_field(paragraph):
            return clone_element(paragraph.find("./w:pPr", NS)), first_run_properties(paragraph)
    return fallback_ppr, fallback_rpr


def centered_page_footer_xml(
    zf: zipfile.ZipFile | None = None,
    part_name: str | None = None,
) -> bytes:
    ftr = ET.Element(qn("ftr"))
    para = ET.SubElement(ftr, qn("p"))
    baseline_ppr: ET.Element | None = None
    baseline_rpr: ET.Element | None = None
    if zf is not None and part_name is not None:
        baseline_ppr, baseline_rpr = footer_baseline_properties(zf, part_name)
    ppr = clone_element(baseline_ppr) or ET.SubElement(para, qn("pPr"))
    if ppr not in list(para):
        para.append(ppr)
    ensure_center_alignment(ppr)
    run_template = clone_element(baseline_rpr) or default_page_run_properties()

    def field_run() -> ET.Element:
        run = ET.SubElement(para, qn("r"))
        if run_template is not None:
            run.append(clone_element(run_template) or default_page_run_properties())
        return run

    begin = ET.SubElement(field_run(), qn("fldChar"))
    begin.set(qn("fldCharType"), "begin")
    instr_run = field_run()
    instr = ET.SubElement(instr_run, qn("instrText"))
    instr.set(f"{XML}space", "preserve")
    instr.text = " PAGE "
    separate = ET.SubElement(field_run(), qn("fldChar"))
    separate.set(qn("fldCharType"), "separate")
    text_run = field_run()
    text = ET.SubElement(text_run, qn("t"))
    text.text = "1"
    end = ET.SubElement(field_run(), qn("fldChar"))
    end.set(qn("fldCharType"), "end")
    return ET.tostring(ftr, encoding="utf-8", xml_declaration=True)


def write_docx(input_docx: Path, output_docx: Path, replacements: dict[str, bytes]) -> None:
    with tempfile.TemporaryDirectory() as tmp_name:
        tmp = Path(tmp_name)
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp)
        for part, data in replacements.items():
            target = tmp / part
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmp)
        output_docx.with_suffix(".zip").replace(output_docx)


def plan_footer_repair(
    input_docx: Path,
    *,
    template_docx: Path | None,
    create_if_missing: bool,
    apply_template_margins: bool,
) -> tuple[dict[str, bytes], dict[str, object], int]:
    with zipfile.ZipFile(input_docx, "r") as zf:
        referenced_parts = referenced_default_page_footer_parts(zf)
        parts = {part for part in referenced_parts if footer_part_has_page_field(zf, part)}
        document = document_root(zf)
        rels_root = relationship_root(zf)
        content_root = content_types_root(zf)
        created_footer_part = ""
        section_changes: list[str] = []
        if create_if_missing and not parts:
            sect_pr = final_section(document)
            if sect_pr is None:
                return {}, {
                    "footer_parts_rewritten": [],
                    "footer_parts_created": [],
                    "section_changes": [],
                    "repair_verdict": "blocked-no-section",
                }, 1
            footer_part = next_footer_part_name(zf)
            rel_id, rel_created = ensure_footer_relationship(rels_root, footer_part)
            content_created = ensure_footer_content_type(content_root, footer_part)
            if ensure_default_footer_reference(sect_pr, rel_id):
                section_changes.append("default_footer_reference_added")
            if ensure_pg_num_type(sect_pr, fmt="decimal", start=None):
                section_changes.append("decimal_page_numbering_added")
            if apply_template_margins:
                pg_mar = template_portrait_pg_mar(template_docx)
                if pg_mar is not None and replace_or_insert_section_child(sect_pr, pg_mar, "pgMar"):
                    section_changes.append("template_portrait_margin_applied")
            created_footer_part = footer_part
            parts.add(footer_part)
            if rel_created:
                section_changes.append("footer_relationship_added")
            if content_created:
                section_changes.append("footer_content_type_added")

        replacements = {
            part: centered_page_footer_xml(zf, part)
            for part in sorted(parts)
        }
        if created_footer_part:
            replacements["word/document.xml"] = serialize_word_xml(document)
            replacements["word/_rels/document.xml.rels"] = serialize_relationships(rels_root)
            replacements["[Content_Types].xml"] = serialize_content_types(content_root)
        report = {
            "footer_parts_rewritten": sorted(parts),
            "referenced_footer_parts": sorted(referenced_parts),
        "blank_referenced_footer_parts_left_unchanged": sorted(referenced_parts - parts),
        "footer_parts_created": [created_footer_part] if created_footer_part else [],
        "footer_page_size_half_points": "21",
        "section_changes": section_changes,
            "create_if_missing": create_if_missing,
            "apply_template_margins": apply_template_margins,
            "template_docx": str(template_docx) if template_docx else "",
            "repair_verdict": "pass" if parts else "blocked-no-footer-parts",
        }
        return replacements, report, 0 if parts else 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--template-docx", help="Template DOCX used only for portrait section margin geometry.")
    parser.add_argument(
        "--create-if-missing",
        action="store_true",
        help="Create one default PAGE footer on the final section when no active footer exists.",
    )
    parser.add_argument(
        "--apply-template-margins",
        action="store_true",
        help="With --create-if-missing, copy the first portrait template section pgMar into the final section.",
    )
    args = parser.parse_args()

    input_docx = Path(args.input_docx)
    output_docx = Path(args.output_docx)
    report = Path(args.report)
    template_docx = Path(args.template_docx) if args.template_docx else None
    replacements, repair_payload, status = plan_footer_repair(
        input_docx,
        template_docx=template_docx,
        create_if_missing=args.create_if_missing,
        apply_template_margins=args.apply_template_margins,
    )
    write_docx(input_docx, output_docx, replacements)
    payload = {
        "schema": "graduation-project-builder.footer-page-number-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        **repair_payload,
    }
    report.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=True, indent=2))
    return status


if __name__ == "__main__":
    raise SystemExit(main())
