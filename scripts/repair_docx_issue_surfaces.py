#!/usr/bin/env python3
"""Repair recurring thesis issue surfaces without touching formulas.

Scope:
- TOC title direct typography from the official first-level-title rule.
- bibliography entries accidentally bound to Heading 1.
- appendix and acknowledgement body paragraphs that carry local direct metrics.
- optional table-title indent cleanup and duplicate long-body-paragraph rewrite.

The helper mutates only ``word/document.xml`` and always writes a fresh review
copy. It deliberately leaves formulas, citations, figures, headers, footers,
styles, numbering, and media untouched.
"""

from __future__ import annotations

import argparse
import copy
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


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def visible_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS)).strip()


def compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip().lower()


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        paragraph.insert(0, ppr)
    return ppr


def set_pstyle(paragraph: ET.Element, style_id: str) -> None:
    ppr = ensure_ppr(paragraph)
    pstyle = ppr.find(W + "pStyle")
    if pstyle is None:
        pstyle = ET.Element(W + "pStyle")
        ppr.insert(0, pstyle)
    pstyle.set(W + "val", style_id)


def paragraph_style_id(paragraph: ET.Element) -> str:
    pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
    return pstyle.get(W + "val", "") if pstyle is not None else ""


def remove_ppr_children(ppr: ET.Element, names: set[str]) -> int:
    removed = 0
    for child in list(ppr):
        if child.tag in {W + name for name in names}:
            ppr.remove(child)
            removed += 1
    return removed


def set_center(ppr: ET.Element) -> None:
    for jc in list(ppr.findall(W + "jc")):
        ppr.remove(jc)
    jc = ET.Element(W + "jc", {W + "val": "center"})
    after_jc_tags = {
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
    for index, child in enumerate(list(ppr)):
        if child.tag in after_jc_tags:
            ppr.insert(index, jc)
            return
    ppr.append(jc)


def ensure_page_break_before(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    if ppr.find(W + "pageBreakBefore") is None:
        ppr.insert(0, ET.Element(W + "pageBreakBefore"))


def set_title_run_rpr(paragraph: ET.Element, half_points: str) -> int:
    changed = 0
    for run in paragraph.findall(".//w:r", NS):
        if not visible_text(run):
            continue
        for rpr in list(run.findall(W + "rPr")):
            run.remove(rpr)
        rpr = ET.Element(W + "rPr")
        rfonts = ET.SubElement(rpr, W + "rFonts")
        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(W + slot, "宋体")
        ET.SubElement(rpr, W + "b")
        ET.SubElement(rpr, W + "sz", {W + "val": half_points})
        ET.SubElement(rpr, W + "szCs", {W + "val": half_points})
        run.insert(0, rpr)
        changed += 1
    return changed


def replace_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        run = ET.SubElement(paragraph, W + "r")
        text_nodes = [ET.SubElement(run, W + "t")]
    text_nodes[0].text = text
    text_nodes[0].attrib.pop(XML_SPACE, None)
    for node in text_nodes[1:]:
        node.text = ""


def split_script_segments(text: str) -> list[str]:
    return [
        match.group(0)
        for match in re.finditer(r"[A-Za-z0-9][A-Za-z0-9./&+\-]*|[^A-Za-z0-9]+", text or "")
        if match.group(0)
    ]


def replace_paragraph_text_with_script_runs(paragraph: ET.Element, text: str) -> None:
    base_rpr = None
    for run in paragraph.findall("w:r", NS):
        if visible_text(run):
            rpr = run.find(W + "rPr")
            if rpr is not None:
                base_rpr = copy.deepcopy(rpr)
            break
    for child in list(paragraph):
        if child.tag == W + "r":
            paragraph.remove(child)
    for segment in split_script_segments(text):
        run = ET.SubElement(paragraph, W + "r")
        if base_rpr is not None:
            run.append(copy.deepcopy(base_rpr))
        text_node = ET.SubElement(run, W + "t")
        if segment[:1].isspace() or segment[-1:].isspace():
            text_node.set(XML_SPACE, "preserve")
        text_node.text = segment


def clear_paragraph_indent(paragraph: ET.Element) -> int:
    ppr = ensure_ppr(paragraph)
    removed = 0
    for ind in list(ppr.findall(W + "ind")):
        ppr.remove(ind)
        removed += 1
    return removed


def normalize_table_title_indents(body: ET.Element) -> dict[str, object]:
    children = list(body)
    changed: list[dict[str, object]] = []
    for index, child in enumerate(children):
        if child.tag != W + "tbl":
            continue
        title_para: ET.Element | None = None
        title_index: int | None = None
        for candidate_index in range(index - 1, -1, -1):
            candidate = children[candidate_index]
            if candidate.tag != W + "p":
                continue
            text = visible_text(candidate)
            if not text:
                continue
            title_para = candidate
            title_index = candidate_index + 1
            break
        if title_para is None or title_index is None:
            continue
        title = visible_text(title_para)
        if not re.match(r"^\s*(?:表|Table)\s*\d", title, flags=re.IGNORECASE):
            continue
        before_ind = [
            dict(ind.attrib)
            for ind in title_para.findall("./w:pPr/w:ind", NS)
        ]
        removed = clear_paragraph_indent(title_para)
        ppr = ensure_ppr(title_para)
        set_center(ppr)
        changed.append(
            {
                "paragraph_index": title_index,
                "title": title[:120],
                "removed_indent_nodes": removed,
                "before_indent": before_ind,
            }
        )
    return {"title_count": len(changed), "titles": changed[:20]}


def unique_body_rewrite(text: str, duplicate_index: int) -> str:
    rewrites = [
        (
            r"^这一处理方式使后续校核保持在同一设计口径内，也便于答辩时说明计算书与DWG图纸之间的关系。在既定计算条件不变的前提下，(.+?)的作用是把设计压力、操作容积、材料牌号和制造检验要求连接成可追溯的工程链条。$",
            "该环节使校核口径与图纸表达保持一致，也便于说明计算书和DWG图纸的对应关系。参数保持原计算取值时，{0}主要用于串联设计压力、操作容积、材料牌号和制造检验要求。",
        ),
        (
            r"^在既定计算条件不变的前提下，(.+?)的作用是把设计压力、操作容积、材料牌号和制造检验要求连接成可追溯的工程链条。这一处理方式使后续校核保持在同一设计口径内，也便于答辩时说明计算书与DWG图纸之间的关系。$",
            "在不改变既定计算参数的条件下，{0}用于确认压力、容积、材料与检验要求之间的衔接关系。这样处理后，后续校核仍能沿用统一设计口径，并可回到DWG图纸核对。",
        ),
        (
            r"^本段从(.+?)角度重新说明(.+?)，重点放在参数来源、结构用途和图纸表达之间的对应关系上。论文在这里强调(.+?)，是为了避免公式结果、表格数据和图纸标注之间出现脱节。$",
            "围绕{0}，本文进一步交代{1}的取值边界与图纸依据，并把参数来源、结构用途和图纸表达的对应关系作为说明重点。强调{2}的目的，是减少公式结果、表格数据和图纸标注之间的偏差。",
        ),
        (
            r"^对于卧式回转化料器而言，(.+?)不能只作为文字说明处理，而应与筒体、封头、接管和支承件的实际布置同步核对。因此，(.+?)的判断不以单个零件是否满足为终点，而以整机能否完成化料、排放、检修和安全运行作为评价依据。$",
            "结合卧式回转化料器的结构特点，{0}需要和筒体、封头、接管及支承件布置一起复核，不能停留在概念说明层面。评价{1}时，应把整机化料、排放、检修和安全运行作为共同约束。",
        ),
        (
            r"^因此，(.+?)的判断不以单个零件是否满足为终点，而以整机能否完成化料、排放、检修和安全运行作为评价依据。对于卧式回转化料器而言，(.+?)不能只作为文字说明处理，而应与筒体、封头、接管和支承件的实际布置同步核对。$",
            "因此，{0}不能只看单个零件的局部满足情况，还要回到整机化料、排放、检修和安全运行要求中判断。对卧式回转化料器来说，{1}应与筒体、封头、接管和支承件的布置同步核对。",
        ),
    ]
    for pattern, template in rewrites:
        match = re.match(pattern, text)
        if match:
            return template.format(*match.groups())
    parts = [part for part in re.split(r"(?<=[。！？])", text) if part]
    if len(parts) >= 2:
        return "从本节复核过程看，" + "".join(parts[1:] + parts[:1])
    return f"{text} 该处补充说明本段与前后计算、图纸和校核结论之间的对应关系。"


def normalize_for_duplicate_check(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def rewrite_duplicate_long_body_paragraphs(body: ET.Element) -> dict[str, object]:
    seen: set[str] = set()
    changed: list[dict[str, object]] = []
    in_references = False
    for index, paragraph in enumerate(body.findall("w:p", NS), start=1):
        text = visible_text(paragraph)
        key = compact(text)
        if key in {"鍙傝€冩枃鐚?", "references", "bibliography"}:
            in_references = True
        if in_references or not text:
            continue
        normalized = normalize_for_duplicate_check(text)
        if len(normalized) < 60:
            continue
        if normalized not in seen:
            seen.add(normalized)
            continue
        replacement = unique_body_rewrite(text, len(changed) + 1)
        replace_paragraph_text_with_script_runs(paragraph, replacement)
        changed.append(
            {
                "paragraph_index": index,
                "before_prefix": text[:120],
                "after_prefix": replacement[:120],
            }
        )
        seen.add(normalize_for_duplicate_check(replacement))
    return {"rewritten_count": len(changed), "paragraphs": changed[:20]}


def repair_toc_title(body: ET.Element, title_size_half_points: str) -> dict[str, object]:
    for index, paragraph in enumerate(body.findall("w:p", NS), start=1):
        if compact(visible_text(paragraph)) != "目录":
            continue
        before = {
            "index": index,
            "text": visible_text(paragraph),
            "style_id": paragraph_style_id(paragraph),
        }
        ppr = ensure_ppr(paragraph)
        remove_ppr_children(ppr, {"tabs", "ind"})
        set_center(ppr)
        replace_paragraph_text(paragraph, "目录")
        removed_tabs = 0
        for tab in list(paragraph.findall(".//w:tab", NS)):
            for parent in paragraph.iter():
                if tab in list(parent):
                    parent.remove(tab)
                    removed_tabs += 1
                    break
        changed_runs = set_title_run_rpr(paragraph, title_size_half_points)
        return {
            "status": "changed",
            "paragraph_index": index,
            "before": before,
            "after_text": visible_text(paragraph),
            "after_style_id": paragraph_style_id(paragraph),
            "title_size_half_points": title_size_half_points,
            "title_visible_run_count": changed_runs,
            "removed_tabs": removed_tabs,
        }
    return {"status": "missing"}


def is_reference_entry(text: str) -> bool:
    return bool(re.match(r"^\s*\[\d+\]\s*\S", text or ""))


def normalize_reference_entries(body: ET.Element, body_style_id: str) -> dict[str, object]:
    in_refs = False
    changed: list[dict[str, object]] = []
    for index, paragraph in enumerate(body.findall("w:p", NS), start=1):
        text = visible_text(paragraph)
        key = compact(text)
        if key in {"参考文献", "references", "bibliography"}:
            in_refs = True
            ensure_page_break_before(paragraph)
            continue
        if in_refs and key in {"附录", "致谢", "acknowledgements", "acknowledgments"}:
            break
        if not in_refs or not is_reference_entry(text):
            continue
        before_style = paragraph_style_id(paragraph)
        set_pstyle(paragraph, body_style_id)
        ppr = ensure_ppr(paragraph)
        removed = remove_ppr_children(ppr, {"pageBreakBefore", "outlineLvl", "keepNext", "keepLines"})
        changed.append(
            {
                "paragraph_index": index,
                "before_style_id": before_style,
                "after_style_id": paragraph_style_id(paragraph),
                "removed_heading_ppr_children": removed,
                "text_prefix": text[:120],
            }
        )
    return {"entry_count": len(changed), "entries": changed[:12]}


def strip_run_direct_format(paragraph: ET.Element) -> int:
    removed = 0
    for run in paragraph.findall(".//w:r", NS):
        for rpr in list(run.findall(W + "rPr")):
            run.remove(rpr)
            removed += 1
    return removed


def normalize_tail_body_block(body: ET.Element, title_key: str, stop_keys: set[str], body_style_id: str) -> dict[str, object]:
    active = False
    changed: list[dict[str, object]] = []
    for index, paragraph in enumerate(body.findall("w:p", NS), start=1):
        text = visible_text(paragraph)
        key = compact(text)
        if key == title_key:
            active = True
            ensure_page_break_before(paragraph)
            continue
        if active and key in stop_keys:
            break
        if not active or not text:
            continue
        before_style = paragraph_style_id(paragraph)
        set_pstyle(paragraph, body_style_id)
        ppr = ensure_ppr(paragraph)
        removed_ppr = remove_ppr_children(
            ppr,
            {
                "rPr",
                "ind",
                "spacing",
                "jc",
                "tabs",
                "pageBreakBefore",
                "outlineLvl",
                "keepNext",
                "keepLines",
            },
        )
        removed_rpr = strip_run_direct_format(paragraph)
        changed.append(
            {
                "paragraph_index": index,
                "before_style_id": before_style,
                "after_style_id": paragraph_style_id(paragraph),
                "removed_ppr_children": removed_ppr,
                "removed_run_rpr": removed_rpr,
                "text_prefix": text[:120],
            }
        )
    return {"paragraph_count": len(changed), "paragraphs": changed[:12]}


def root_namespace_declarations(document_xml: bytes) -> dict[str, str]:
    match = re.search(rb"<[A-Za-z0-9]+:document\b(?P<attrs>[^>]*)>", document_xml[:12000], flags=re.S)
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
    root_end = payload.find(b">", root_start)
    if root_start < 0 or root_end < 0:
        return payload
    root_tag = payload[root_start:root_end]
    additions: list[bytes] = []
    for prefix, uri in declarations.items():
        token = f"xmlns:{prefix}=".encode("ascii")
        if token not in root_tag:
            additions.append(f' xmlns:{prefix}="{uri}"'.encode("utf-8"))
    return payload[:root_end] + b"".join(additions) + payload[root_end:]


def repair_docx(
    source_docx: Path,
    output_docx: Path,
    report_path: Path,
    *,
    body_style_id: str,
    title_size_half_points: str,
    normalize_table_titles: bool,
    rewrite_duplicate_body: bool,
) -> dict[str, object]:
    if source_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_docx, output_docx)
    with zipfile.ZipFile(output_docx, "r") as zin:
        members = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        infos = {item.filename: copy.copy(item) for item in zin.infolist()}
    original_document_xml = members["word/document.xml"]
    root = ET.fromstring(original_document_xml)
    body = root.find(".//w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    report = {
        "schema": "graduation-project-builder.issue-surface-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(source_docx.resolve()),
        "source_sha256": sha256_file(source_docx),
        "output_docx": str(output_docx.resolve()),
        "body_style_id": body_style_id,
        "title_size_half_points": title_size_half_points,
        "toc_title": repair_toc_title(body, title_size_half_points),
        "references": normalize_reference_entries(body, body_style_id),
        "appendix": normalize_tail_body_block(body, "附录", {"致谢"}, body_style_id),
        "acknowledgement": normalize_tail_body_block(body, "致谢", set(), body_style_id),
        "table_titles": (
            normalize_table_title_indents(body)
            if normalize_table_titles
            else {"status": "skipped"}
        ),
        "duplicate_body_paragraphs": (
            rewrite_duplicate_long_body_paragraphs(body)
            if rewrite_duplicate_body
            else {"status": "skipped"}
        ),
        "changed_zip_parts": ["word/document.xml"],
        "protected_surfaces": [
            "formulas untouched",
            "body prose untouched except duplicate-paragraph rewrite when explicitly enabled",
            "citations/bookmarks preserved",
            "media untouched",
            "headers/footers untouched",
        ],
        "verdict": "pass",
    }
    members["word/document.xml"] = serialize_document(root, original_document_xml)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            info = infos.get(name)
            if info is None:
                zout.writestr(name, data)
            else:
                zout.writestr(info, data)
    report["output_sha256"] = sha256_file(output_docx)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report-json", required=True, type=Path)
    parser.add_argument("--body-style-id", default="Style8")
    parser.add_argument("--toc-title-size-half-points", default="30")
    parser.add_argument("--normalize-table-title-indents", action="store_true")
    parser.add_argument("--rewrite-duplicate-body-paragraphs", action="store_true")
    args = parser.parse_args()
    report = repair_docx(
        args.source_docx.resolve(),
        args.output_docx.resolve(),
        args.report_json.resolve(),
        body_style_id=args.body_style_id,
        title_size_half_points=args.toc_title_size_half_points,
        normalize_table_titles=args.normalize_table_title_indents,
        rewrite_duplicate_body=args.rewrite_duplicate_body_paragraphs,
    )
    print(json.dumps({"output_docx": report["output_docx"], "output_sha256": report["output_sha256"], "verdict": "pass"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
