#!/usr/bin/env python3
"""Audit and optionally repair abnormal DOCX list/bullet pollution.

This gate protects thesis surfaces where Word list state is easy to miss in
XML-only checks but obvious in WPS/Word/PDF: cover/front matter, abstracts, TOC,
body headings, references title, and bibliography entries.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
XML_SPACE = f"{{{XML_NS}}}space"

BULLET_CHARS = "\u2022\u25cf\u25e6\u2219\u2043\uf0b7"
BULLET_PREFIX_RE = re.compile(rf"^[\s\u3000]*(?P<bullet>[{re.escape(BULLET_CHARS)}])[\s\u3000]*")
VISIBLE_NUMBER_RE = re.compile(r"^\s*(?:\[[0-9]{1,3}\]|[0-9]{1,3}[.、])\s+")


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


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def attr(element: ET.Element | None, local: str = "val") -> str:
    if element is None:
        return ""
    return element.attrib.get(qn(local), "")


def paragraph_style_id(paragraph: ET.Element) -> str:
    return attr(paragraph.find("./w:pPr/w:pStyle", NS))


def paragraph_num_pr(paragraph: ET.Element) -> ET.Element | None:
    return paragraph.find("./w:pPr/w:numPr", NS)


def strip_bullet_prefix(text: str) -> str:
    return BULLET_PREFIX_RE.sub("", text or "", count=1)


def has_visible_bullet_prefix(text: str) -> bool:
    return BULLET_PREFIX_RE.match(text or "") is not None


def text_nodes(paragraph: ET.Element) -> list[ET.Element]:
    return list(paragraph.findall(".//w:t", NS))


def remove_visible_bullet_prefix(paragraph: ET.Element) -> bool:
    nodes = text_nodes(paragraph)
    combined = "".join(node.text or "" for node in nodes)
    match = BULLET_PREFIX_RE.match(combined)
    if not match:
        return False
    drop = match.end()
    for node in nodes:
        value = node.text or ""
        if drop <= 0:
            break
        if len(value) <= drop:
            node.text = ""
            drop -= len(value)
        else:
            node.text = value[drop:]
            drop = 0
        if node.text and (node.text[:1].isspace() or node.text[-1:].isspace()):
            node.set(XML_SPACE, "preserve")
    return True


def is_zh_abstract_title(text: str) -> bool:
    return compact(re.sub(r"[:\uff1a].*$", "", strip_bullet_prefix(text))) == "\u6458\u8981"


def is_en_abstract_title(text: str) -> bool:
    return compact(re.sub(r"[:\uff1a].*$", "", strip_bullet_prefix(text))).startswith("abstract")


def is_zh_keyword(text: str) -> bool:
    return compact(strip_bullet_prefix(text)).startswith("\u5173\u952e\u8bcd")


def is_en_keyword(text: str) -> bool:
    value = compact(strip_bullet_prefix(text))
    return value.startswith("keywords") or value.startswith("keywords:") or value.startswith("keywords\uff1a")


def is_toc_title(text: str) -> bool:
    return compact(strip_bullet_prefix(text)) in {"\u76ee\u5f55", "contents", "tableofcontents"}


def is_reference_title(text: str) -> bool:
    return compact(strip_bullet_prefix(text)) in {"\u53c2\u8003\u6587\u732e", "references", "bibliography"}


def is_acknowledgement_title(text: str) -> bool:
    return compact(strip_bullet_prefix(text)) in {
        "\u81f4\u8c22",
        "\u8c22\u8f9e",
        "acknowledgement",
        "acknowledgements",
        "acknowledgment",
        "acknowledgments",
    }


def is_appendix_title(text: str) -> bool:
    return compact(strip_bullet_prefix(text)) in {"\u9644\u5f55", "appendix"}


def is_chapter_heading_text(text: str) -> bool:
    value = strip_bullet_prefix(text).strip()
    return bool(re.match(r"^\u7b2c\s*[0-9\u4e00-\u9fa5]+\s*\u7ae0(?:\s+\S.*)?$", value))


def body_heading_level(text: str) -> int | None:
    value = strip_bullet_prefix(text).strip()
    if is_chapter_heading_text(value):
        return 1
    if re.match(r"^[1-9]\d*(?:\.\d+){2}\s+\S", value):
        return 3
    if re.match(r"^[1-9]\d*\.\d+\s+\S", value):
        return 2
    if re.match(r"^[1-9]\d*\s+\S", value):
        return 1
    return None


def looks_like_toc_entry(text: str) -> bool:
    value = strip_bullet_prefix(text).strip()
    if not value:
        return False
    if "\t" in value:
        label = value.split("\t", 1)[0].strip()
    elif "\u2026" in value:
        label = value.split("\u2026", 1)[0].strip()
    else:
        label = re.sub(r"\s*(?:[ivxlcdmIVXLCDM]+|[0-9]+)\s*$", "", value).strip()
    return body_heading_level(label) is not None or is_reference_title(label) or is_acknowledgement_title(label)


def visible_paragraphs(root: ET.Element) -> list[ET.Element]:
    body = root.find("./w:body", NS)
    if body is None:
        return []
    return list(body.findall(".//w:p", NS))


def first_indices(texts: list[str]) -> dict[str, int | None]:
    indices: dict[str, int | None] = {
        "zh_abstract": None,
        "en_abstract": None,
        "toc": None,
        "first_body": None,
        "references": None,
        "acknowledgement": None,
        "appendix": None,
    }
    for index, text in enumerate(texts, start=1):
        if indices["zh_abstract"] is None and is_zh_abstract_title(text):
            indices["zh_abstract"] = index
        if indices["en_abstract"] is None and is_en_abstract_title(text):
            indices["en_abstract"] = index
        if indices["toc"] is None and is_toc_title(text):
            indices["toc"] = index
        if indices["first_body"] is None and indices["toc"] is not None and index > indices["toc"] and body_heading_level(text) == 1:
            indices["first_body"] = index
        if indices["references"] is None and is_reference_title(text):
            indices["references"] = index
        if indices["acknowledgement"] is None and is_acknowledgement_title(text):
            indices["acknowledgement"] = index
        if indices["appendix"] is None and is_appendix_title(text):
            indices["appendix"] = index
    return indices


def bounded_end(start: int | None, candidates: list[int | None], default: int) -> int:
    if start is None:
        return default
    valid = [value for value in candidates if value is not None and value > start]
    return min(valid) if valid else default


def classify_surface(index: int, text: str, indices: dict[str, int | None], total: int) -> str:
    zh = indices["zh_abstract"]
    en = indices["en_abstract"]
    toc = indices["toc"]
    body = indices["first_body"]
    references = indices["references"]
    acknowledgement = indices["acknowledgement"]
    appendix = indices["appendix"]
    if zh is not None and index < zh:
        return "cover_style"
    if is_zh_abstract_title(text):
        return "zh_abstract_title"
    if zh is not None and index > zh and index < bounded_end(zh, [en, toc, body], total + 1):
        return "zh_keyword_line" if is_zh_keyword(text) else "zh_abstract_body"
    if is_en_abstract_title(text):
        return "en_abstract_title"
    if en is not None and index > en and index < bounded_end(en, [toc, body], total + 1):
        return "en_keyword_line" if is_en_keyword(text) else "en_abstract_body"
    if is_toc_title(text):
        return "toc_title"
    if toc is not None and index > toc and (body is None or index < body):
        return "toc_entries"
    if references is not None and index == references:
        return "references_title"
    if references is not None and index > references and index < bounded_end(references, [appendix, acknowledgement], total + 1):
        return "references_entries"
    if body is not None and index >= body and body_heading_level(text) is not None:
        return "body_heading_levels"
    return "other"


def style_num_pr_map(styles_root: ET.Element | None) -> dict[str, ET.Element]:
    if styles_root is None:
        return {}
    result: dict[str, ET.Element] = {}
    for style in styles_root.findall("./w:style", NS):
        style_id = attr(style, "styleId")
        num_pr = style.find("./w:pPr/w:numPr", NS)
        if style_id and num_pr is not None:
            result[style_id] = num_pr
    return result


def load_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except (KeyError, ET.ParseError):
        return None


def num_pr_level(num_pr: ET.Element | None) -> tuple[str, str]:
    if num_pr is None:
        return "", "0"
    return attr(num_pr.find("./w:numId", NS)), attr(num_pr.find("./w:ilvl", NS)) or "0"


def numbering_models(numbering_root: ET.Element | None) -> dict[tuple[str, str], dict[str, str]]:
    if numbering_root is None:
        return {}
    abstract_by_id = {
        attr(item, "abstractNumId"): item
        for item in numbering_root.findall("./w:abstractNum", NS)
        if attr(item, "abstractNumId")
    }
    result: dict[tuple[str, str], dict[str, str]] = {}
    for num in numbering_root.findall("./w:num", NS):
        num_id = attr(num, "numId")
        abstract_id = attr(num.find("./w:abstractNumId", NS))
        abstract = abstract_by_id.get(abstract_id)
        if not num_id or abstract is None:
            continue
        for lvl in abstract.findall("./w:lvl", NS):
            ilvl = attr(lvl, "ilvl") or "0"
            result[(num_id, ilvl)] = {
                "numId": num_id,
                "ilvl": ilvl,
                "numFmt": attr(lvl.find("./w:numFmt", NS)),
                "lvlText": attr(lvl.find("./w:lvlText", NS)),
            }
    return result


def model_for(num_pr: ET.Element | None, models: dict[tuple[str, str], dict[str, str]]) -> dict[str, str]:
    num_id, ilvl = num_pr_level(num_pr)
    if not num_id:
        return {}
    return models.get((num_id, ilvl), {"numId": num_id, "ilvl": ilvl, "numFmt": "", "lvlText": ""})


def model_is_bullet(model: dict[str, str]) -> bool:
    fmt = (model.get("numFmt") or "").lower()
    lvl_text = model.get("lvlText") or ""
    return fmt == "bullet" or any(char in lvl_text for char in BULLET_CHARS)


def model_is_reference_numbering(model: dict[str, str]) -> bool:
    fmt = (model.get("numFmt") or "").lower()
    lvl_text = model.get("lvlText") or ""
    if model_is_bullet(model):
        return False
    return fmt in {"decimal", "decimalzero", "chinesecounting"} and "%1" in lvl_text


def bibliography_auto_numbering_allowed(surface: str, model: dict[str, str], visible_number: bool) -> bool:
    return surface == "references_entries" and model_is_reference_numbering(model) and not visible_number


def template_numpr_allowances(template_docx: Path | None) -> dict[str, bool]:
    allowed = {
        "cover_style": False,
        "zh_abstract_title": False,
        "zh_abstract_body": False,
        "zh_keyword_line": False,
        "en_abstract_title": False,
        "en_abstract_body": False,
        "en_keyword_line": False,
        "toc_title": False,
        "toc_entries": False,
        "body_heading_levels": False,
        "references_title": False,
        "references_entries": False,
    }
    if template_docx is None or not template_docx.exists():
        return allowed
    try:
        with zipfile.ZipFile(template_docx) as zf:
            document_root = load_xml(zf, "word/document.xml")
            styles_root = load_xml(zf, "word/styles.xml")
            numbering_root = load_xml(zf, "word/numbering.xml")
        if document_root is None:
            return allowed
        paragraphs = visible_paragraphs(document_root)
        texts = [text_of(paragraph).strip() for paragraph in paragraphs]
        indices = first_indices(texts)
        style_numprs = style_num_pr_map(styles_root)
        models = numbering_models(numbering_root)
        for index, paragraph in enumerate(paragraphs, start=1):
            text = texts[index - 1]
            surface = classify_surface(index, text, indices, len(paragraphs))
            if surface not in allowed:
                continue
            direct = paragraph_num_pr(paragraph)
            style_num_pr = style_numprs.get(paragraph_style_id(paragraph))
            model = model_for(direct or style_num_pr, models)
            if (direct is not None or style_num_pr is not None) and not model_is_bullet(model):
                allowed[surface] = True
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return allowed
    return allowed


def make_issue(
    *,
    index: int,
    surface: str,
    text: str,
    reason: str,
    direct_num_pr: ET.Element | None,
    style_id: str,
    style_num_pr: ET.Element | None,
    direct_model: dict[str, str],
    style_model: dict[str, str],
) -> dict[str, object]:
    return {
        "paragraph_index": index,
        "surface": surface,
        "text_prefix": text[:140],
        "reason": reason,
        "visible_bullet_prefix": has_visible_bullet_prefix(text),
        "visible_number_prefix": VISIBLE_NUMBER_RE.match(strip_bullet_prefix(text)) is not None,
        "style_id": style_id,
        "has_direct_numPr": direct_num_pr is not None,
        "has_style_numPr": style_num_pr is not None,
        "direct_numbering_model": direct_model,
        "style_numbering_model": style_model,
    }


def audit_docx_list_pollution(docx_path: Path, *, template_docx: Path | None = None) -> dict[str, object]:
    with zipfile.ZipFile(docx_path) as zf:
        document_root = load_xml(zf, "word/document.xml")
        styles_root = load_xml(zf, "word/styles.xml")
        numbering_root = load_xml(zf, "word/numbering.xml")
    if document_root is None:
        raise RuntimeError(f"word/document.xml is missing or invalid in {docx_path}")
    paragraphs = visible_paragraphs(document_root)
    texts = [text_of(paragraph).strip() for paragraph in paragraphs]
    indices = first_indices(texts)
    style_numprs = style_num_pr_map(styles_root)
    models = numbering_models(numbering_root)
    allowances = template_numpr_allowances(template_docx)
    issues: list[dict[str, object]] = []
    protected = set(allowances)
    for index, paragraph in enumerate(paragraphs, start=1):
        text = texts[index - 1]
        surface = classify_surface(index, text, indices, len(paragraphs))
        if surface not in protected:
            continue
        if not text and surface != "cover_style" and paragraph_num_pr(paragraph) is None:
            continue
        direct = paragraph_num_pr(paragraph)
        style_id = paragraph_style_id(paragraph)
        style_num_pr = style_numprs.get(style_id)
        direct_model = model_for(direct, models)
        style_model = model_for(style_num_pr, models)
        visible_bullet = has_visible_bullet_prefix(text)
        direct_bullet = model_is_bullet(direct_model)
        style_bullet = model_is_bullet(style_model)
        visible_number = VISIBLE_NUMBER_RE.match(strip_bullet_prefix(text)) is not None
        direct_ref_numbering = bibliography_auto_numbering_allowed(surface, direct_model, visible_number)
        style_ref_numbering = bibliography_auto_numbering_allowed(surface, style_model, visible_number)
        if visible_bullet:
            issues.append(
                make_issue(
                    index=index,
                    surface=surface,
                    text=text,
                    reason="visible bullet prefix is not allowed on protected thesis surfaces",
                    direct_num_pr=direct,
                    style_id=style_id,
                    style_num_pr=style_num_pr,
                    direct_model=direct_model,
                    style_model=style_model,
                )
            )
        if direct is not None and not direct_ref_numbering and (direct_bullet or not allowances.get(surface, False)):
            issues.append(
                make_issue(
                    index=index,
                    surface=surface,
                    text=text,
                    reason="direct paragraph numPr/list state is not allowed for this protected surface",
                    direct_num_pr=direct,
                    style_id=style_id,
                    style_num_pr=style_num_pr,
                    direct_model=direct_model,
                    style_model=style_model,
                )
            )
        if style_num_pr is not None and not style_ref_numbering and (style_bullet or not allowances.get(surface, False)):
            issues.append(
                make_issue(
                    index=index,
                    surface=surface,
                    text=text,
                    reason="paragraph style carries numPr/list state into a protected surface",
                    direct_num_pr=direct,
                    style_id=style_id,
                    style_num_pr=style_num_pr,
                    direct_model=direct_model,
                    style_model=style_model,
                )
            )
        if surface == "references_entries" and direct is not None and visible_number:
            issues.append(
                make_issue(
                    index=index,
                    surface=surface,
                    text=text,
                    reason="bibliography entry mixes Word numbering with visible manual numbering",
                    direct_num_pr=direct,
                    style_id=style_id,
                    style_num_pr=style_num_pr,
                    direct_model=direct_model,
                    style_model=style_model,
                )
            )
    by_surface: dict[str, int] = {}
    for issue in issues:
        key = str(issue["surface"])
        by_surface[key] = by_surface.get(key, 0) + 1
    return {
        "schema": "graduation-project-builder.docx-list-pollution-audit.v1",
        "docx_path": str(docx_path),
        "docx_sha256": sha256_file(docx_path),
        "template_docx_path": str(template_docx) if template_docx else "",
        "template_docx_sha256": sha256_file(template_docx) if template_docx and template_docx.exists() else "",
        "surface_indices": indices,
        "template_numpr_allowances": allowances,
        "issue_count": len(issues),
        "issue_count_by_surface": by_surface,
        "passed": not issues,
        "issues": issues,
    }


def remove_direct_num_pr(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return False
    removed = False
    for node in list(ppr.findall("./w:numPr", NS)):
        ppr.remove(node)
        removed = True
    return removed


def remove_style_num_pr(styles_root: ET.Element | None, style_ids: set[str]) -> int:
    if styles_root is None:
        return 0
    removed = 0
    for style in styles_root.findall("./w:style", NS):
        style_id = attr(style, "styleId")
        if style_id not in style_ids:
            continue
        ppr = style.find("./w:pPr", NS)
        if ppr is None:
            continue
        for node in list(ppr.findall("./w:numPr", NS)):
            ppr.remove(node)
            removed += 1
    return removed


def write_docx_with_parts(source_docx: Path, output_docx: Path, replacements: dict[str, bytes]) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as zin, zipfile.ZipFile(output_docx, "w") as zout:
        seen: set[str] = set()
        for item in zin.infolist():
            payload = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, payload)
            seen.add(item.filename)
        for name, payload in replacements.items():
            if name not in seen:
                zout.writestr(name, payload)


def repair_docx_list_pollution(source_docx: Path, output_docx: Path, *, template_docx: Path | None = None) -> dict[str, object]:
    if source_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy, not the source path")
    with zipfile.ZipFile(source_docx) as zf:
        document_root = load_xml(zf, "word/document.xml")
        styles_root = load_xml(zf, "word/styles.xml")
        numbering_root = load_xml(zf, "word/numbering.xml")
    if document_root is None:
        raise RuntimeError(f"word/document.xml is missing or invalid in {source_docx}")
    paragraphs = visible_paragraphs(document_root)
    texts = [text_of(paragraph).strip() for paragraph in paragraphs]
    indices = first_indices(texts)
    style_numprs = style_num_pr_map(styles_root)
    models = numbering_models(numbering_root)
    allowances = template_numpr_allowances(template_docx)
    touched: list[dict[str, object]] = []
    style_ids_to_strip: set[str] = set()
    protected = set(allowances)
    for index, paragraph in enumerate(paragraphs, start=1):
        text = texts[index - 1]
        surface = classify_surface(index, text, indices, len(paragraphs))
        if surface not in protected:
            continue
        direct = paragraph_num_pr(paragraph)
        style_id = paragraph_style_id(paragraph)
        style_num_pr = style_numprs.get(style_id)
        direct_model = model_for(direct, models)
        style_model = model_for(style_num_pr, models)
        visible_number = VISIBLE_NUMBER_RE.match(strip_bullet_prefix(text)) is not None
        direct_ref_numbering = bibliography_auto_numbering_allowed(surface, direct_model, visible_number)
        style_ref_numbering = bibliography_auto_numbering_allowed(surface, style_model, visible_number)
        changed: dict[str, object] = {
            "paragraph_index": index,
            "surface": surface,
            "text_prefix_before": text[:120],
            "style_id": style_id,
            "removed_visible_bullet_prefix": False,
            "removed_direct_numPr": False,
            "queued_style_numPr_strip": False,
        }
        if has_visible_bullet_prefix(text):
            changed["removed_visible_bullet_prefix"] = remove_visible_bullet_prefix(paragraph)
        if direct is not None and (
            (not direct_ref_numbering and (model_is_bullet(direct_model) or not allowances.get(surface, False)))
            or (surface == "references_entries" and visible_number)
        ):
            changed["removed_direct_numPr"] = remove_direct_num_pr(paragraph)
        if (
            style_id
            and style_num_pr is not None
            and not style_ref_numbering
            and (model_is_bullet(style_model) or not allowances.get(surface, False))
        ):
            style_ids_to_strip.add(style_id)
            changed["queued_style_numPr_strip"] = True
        if any(changed[key] for key in ("removed_visible_bullet_prefix", "removed_direct_numPr", "queued_style_numPr_strip")):
            touched.append(changed)
    stripped_style_count = remove_style_num_pr(styles_root, style_ids_to_strip)
    replacements = {
        "word/document.xml": ET.tostring(document_root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)
    }
    if styles_root is not None and stripped_style_count:
        replacements["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)
    write_docx_with_parts(source_docx, output_docx, replacements)
    after_audit = audit_docx_list_pollution(output_docx, template_docx=template_docx)
    return {
        "schema": "graduation-project-builder.docx-list-pollution-repair.v1",
        "source_docx": str(source_docx),
        "source_sha256": sha256_file(source_docx),
        "template_docx": str(template_docx) if template_docx else "",
        "template_sha256": sha256_file(template_docx) if template_docx and template_docx.exists() else "",
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "touched_paragraph_count": len(touched),
        "touched_paragraphs": touched,
        "stripped_style_ids": sorted(style_ids_to_strip),
        "stripped_style_numPr_count": stripped_style_count,
        "post_repair_audit": after_audit,
    }


def make_test_docx(path: Path, *, polluted: bool) -> Path:
    num_pr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if polluted else ""
    bullet = "\u2022 " if polluted else ""
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}">
  <w:body>
    <w:p><w:pPr>{num_pr}</w:pPr><w:r><w:t>{bullet}\u5c01\u9762\u6b8b\u7559</w:t></w:r></w:p>
    <w:p><w:r><w:t>\u6458\u8981</w:t></w:r></w:p>
    <w:p><w:r><w:t>\u672c\u6587\u7814\u7a76\u8bbe\u8ba1\u3002</w:t></w:r></w:p>
    <w:p><w:r><w:t>Abstract</w:t></w:r></w:p>
    <w:p><w:r><w:t>Design summary.</w:t></w:r></w:p>
    <w:p><w:r><w:t>\u76ee\u5f55</w:t></w:r></w:p>
    <w:p><w:r><w:t>\u7b2c1\u7ae0  \u7eea\u8bba\t1</w:t></w:r></w:p>
    <w:p><w:pPr>{num_pr}</w:pPr><w:r><w:t>{bullet}\u7b2c1\u7ae0  \u7eea\u8bba</w:t></w:r></w:p>
    <w:p><w:pPr>{num_pr}</w:pPr><w:r><w:t>{bullet}1.1 \u7814\u7a76\u80cc\u666f</w:t></w:r></w:p>
    <w:p><w:pPr>{num_pr}</w:pPr><w:r><w:t>{bullet}\u53c2\u8003\u6587\u732e</w:t></w:r></w:p>
    <w:p><w:pPr>{num_pr}</w:pPr><w:r><w:t>{bullet}1. GB/T 10595-2017 \u5e26\u5f0f\u8f93\u9001\u673a.</w:t></w:r></w:p>
    <w:p><w:r><w:t>\u81f4\u8c22</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>'''
    numbering_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W_NS}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''
    styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""")
        zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""")
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/numbering.xml", numbering_xml)
        zf.writestr("word/styles.xml", styles_xml)
    return path


def run_self_test() -> int:
    with tempfile.TemporaryDirectory(prefix="list_pollution_selftest_") as raw:
        td = Path(raw)
        polluted = make_test_docx(td / "polluted.docx", polluted=True)
        clean = make_test_docx(td / "clean.docx", polluted=False)
        polluted_report = audit_docx_list_pollution(polluted)
        clean_report = audit_docx_list_pollution(clean)
        repaired = td / "repaired.docx"
        repair_report = repair_docx_list_pollution(polluted, repaired)
        repaired_report = audit_docx_list_pollution(repaired)
        ok = (
            not polluted_report["passed"]
            and clean_report["passed"]
            and repair_report["post_repair_audit"]["passed"]
            and repaired_report["passed"]
        )
        print(
            json.dumps(
                {
                    "schema": "graduation-project-builder.docx-list-pollution-selftest.v1",
                    "passed": ok,
                    "polluted_issue_count": polluted_report["issue_count"],
                    "clean_issue_count": clean_report["issue_count"],
                    "repaired_issue_count": repaired_report["issue_count"],
                },
                ensure_ascii=True,
                indent=2,
            )
        )
        return 0 if ok else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit and optionally repair DOCX protected-surface list/bullet pollution.")
    parser.add_argument("--docx", type=Path, help="DOCX to audit.")
    parser.add_argument("--template-docx", type=Path, help="Optional template/sample DOCX for non-bullet numbering allowances.")
    parser.add_argument("--report-json", type=Path, help="Write audit or repair report JSON.")
    parser.add_argument("--repair-output-docx", type=Path, help="Write repaired DOCX instead of read-only audit.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return run_self_test()
    if args.docx is None:
        parser.error("--docx is required unless --self-test is used")
    if args.repair_output_docx:
        report = repair_docx_list_pollution(args.docx.resolve(), args.repair_output_docx.resolve(), template_docx=args.template_docx)
        passed = bool(report.get("post_repair_audit", {}).get("passed"))
    else:
        report = audit_docx_list_pollution(args.docx.resolve(), template_docx=args.template_docx)
        passed = bool(report.get("passed"))
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"passed": passed, "issue_count": report.get("issue_count", report.get("post_repair_audit", {}).get("issue_count", ""))}, ensure_ascii=True))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
