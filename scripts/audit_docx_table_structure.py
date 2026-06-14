#!/usr/bin/env python3
"""Audit thesis body table structure without mutating the DOCX."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
import shutil
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]
ET.register_namespace("w", W)
TABLE_TITLE_RE = re.compile(r"\s*\u8868\s*(?:[A-Za-z\uff21-\uff3a]|\d+)(?:[\-\u2014\uff0d.\uff0e]\d+)+")


def q(name: str) -> str:
    return f"{{{W}}}{name}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


PPR_CHILD_ORDER = {
    q("pStyle"): 0,
    q("keepNext"): 1,
    q("keepLines"): 2,
    q("pageBreakBefore"): 3,
    q("widowControl"): 4,
    q("numPr"): 5,
    q("spacing"): 20,
    q("ind"): 21,
    q("jc"): 22,
    q("rPr"): 50,
    q("sectPr"): 99,
}


def text_of(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def compact(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def w_attr(element: ET.Element | None, name: str) -> str | None:
    return element.get(q(name)) if element is not None else None


def paragraph_info(paragraph: ET.Element | None) -> dict[str, object]:
    ppr = paragraph.find("w:pPr", NS) if paragraph is not None else None
    ind = ppr.find("w:ind", NS) if ppr is not None else None
    jc = ppr.find("w:jc", NS) if ppr is not None else None
    pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
    keep_next = ppr.find("w:keepNext", NS) is not None if ppr is not None else False
    return {
        "text": text_of(paragraph),
        "style": w_attr(pstyle, "val"),
        "jc": w_attr(jc, "val"),
        "keepNext": keep_next,
        "ind": {
            key: ind.get(q(key))
            for key in ("firstLine", "firstLineChars", "left", "right", "hanging", "hangingChars")
            if ind is not None and ind.get(q(key)) is not None
        },
    }


def border_info(border: ET.Element | None) -> dict[str, str] | None:
    if border is None:
        return None
    return {
        key: border.get(q(key))
        for key in ("val", "sz", "space", "color")
        if border.get(q(key)) is not None
    }


def cell_borders(cell: ET.Element) -> dict[str, object]:
    tcpr = cell.find("w:tcPr", NS)
    borders = tcpr.find("w:tcBorders", NS) if tcpr is not None else None
    if borders is None:
        return {}
    result: dict[str, object] = {}
    for name in ("top", "bottom", "left", "right", "insideH", "insideV"):
        node = borders.find(f"w:{name}", NS)
        if node is not None:
            result[name] = border_info(node)
    return result


def row_flags(row: ET.Element | None) -> dict[str, bool]:
    trpr = row.find("w:trPr", NS) if row is not None else None
    return {
        "tblHeader": trpr.find("w:tblHeader", NS) is not None if trpr is not None else False,
        "cantSplit": trpr.find("w:cantSplit", NS) is not None if trpr is not None else False,
    }


def run_size(run: ET.Element) -> str | None:
    rpr = run.find("w:rPr", NS)
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    return w_attr(sz, "val")


def table_cell_margins(table: ET.Element) -> dict[str, object]:
    tblpr = table.find("w:tblPr", NS)
    margins = tblpr.find("w:tblCellMar", NS) if tblpr is not None else None
    result: dict[str, object] = {}
    if margins is None:
        return result
    for name in ("top", "bottom", "left", "right"):
        node = margins.find(f"w:{name}", NS)
        if node is not None:
            result[name] = {"w": w_attr(node, "w"), "type": w_attr(node, "type")}
    return result


def first_row_cell_texts(table: ET.Element) -> list[str]:
    first_row = table.find("w:tr", NS)
    if first_row is None:
        return []
    return [compact(text_of(cell)) for cell in first_row.findall("w:tc", NS)]


def table_matches_header(table: ET.Element, expected_cells: list[str]) -> bool:
    actual = first_row_cell_texts(table)
    expected = [compact(item) for item in expected_cells]
    return bool(expected) and actual[: len(expected)] == expected


def load_body_children(docx: Path) -> list[ET.Element]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    return list(body)


def ensure_child(parent: ET.Element, tag: str, before_tags: set[str] | None = None) -> ET.Element:
    existing = parent.find(f"w:{tag}", NS)
    if existing is not None:
        return existing
    node = ET.Element(q(tag))
    if before_tags:
        for index, child in enumerate(list(parent)):
            if child.tag in {q(name) for name in before_tags}:
                parent.insert(index, node)
                return node
    parent.append(node)
    return node


def normalize_border_order(parent: ET.Element) -> None:
    """Keep WordprocessingML border children in schema order."""
    order = {
        q("top"): 0,
        q("left"): 1,
        q("bottom"): 2,
        q("right"): 3,
        q("insideH"): 4,
        q("insideV"): 5,
        q("tl2br"): 6,
        q("tr2bl"): 7,
    }
    children = list(parent)
    children.sort(key=lambda child: order.get(child.tag, 100))
    parent[:] = children


def normalize_ppr_order(ppr: ET.Element) -> None:
    children = list(ppr)
    children.sort(key=lambda child: PPR_CHILD_ORDER.get(child.tag, 80))
    ppr[:] = children


def set_border(parent: ET.Element, name: str, val: str, sz: str = "8") -> None:
    border = ensure_child(parent, name)
    border.set(q("val"), val)
    border.set(q("sz"), sz if val == "single" else "0")
    border.set(q("space"), "0")
    border.set(q("color"), "auto")
    normalize_border_order(parent)


def repair_table(table: ET.Element) -> None:
    tblpr = ensure_child(table, "tblPr", before_tags={"tblGrid", "tr"})
    tbl_borders = ensure_child(tblpr, "tblBorders")
    set_border(tbl_borders, "top", "single")
    set_border(tbl_borders, "bottom", "single")
    set_border(tbl_borders, "left", "nil")
    set_border(tbl_borders, "right", "nil")
    set_border(tbl_borders, "insideH", "nil")
    set_border(tbl_borders, "insideV", "nil")
    rows = table.findall("w:tr", NS)
    for row_index, row in enumerate(rows):
        trpr = ensure_child(row, "trPr", before_tags={"tc"})
        if row_index == 0:
            ensure_child(trpr, "tblHeader")
        ensure_child(trpr, "cantSplit")
        is_first = row_index == 0
        is_last = row_index == len(rows) - 1
        for cell in row.findall("w:tc", NS):
            tcpr = ensure_child(cell, "tcPr", before_tags={"p"})
            borders = ensure_child(tcpr, "tcBorders")
            set_border(borders, "top", "single" if is_first else "nil")
            set_border(borders, "bottom", "single" if is_first or is_last else "nil")
            set_border(borders, "left", "nil")
            set_border(borders, "right", "nil")
            set_border(borders, "insideH", "nil")
            set_border(borders, "insideV", "nil")
            for paragraph in cell.findall("w:p", NS):
                ppr = paragraph.find("w:pPr", NS)
                if ppr is not None:
                    ind = ppr.find("w:ind", NS)
                    if ind is not None:
                        for key in ("firstLine", "firstLineChars", "left", "right", "hanging", "hangingChars"):
                            ind.attrib.pop(q(key), None)
                for run in paragraph.findall("w:r", NS):
                    if not compact(text_of(run)):
                        continue
                    rpr = ensure_child(run, "rPr", before_tags={"t", "drawing", "object", "pict"})
                    sz = ensure_child(rpr, "sz")
                    sz.set(q("val"), "21")
                    szcs = ensure_child(rpr, "szCs")
                    szcs.set(q("val"), "21")


def clone_table_title_paragraph(donor: ET.Element | None, title: str) -> ET.Element:
    paragraph = ET.Element(q("p"))
    if donor is not None:
        ppr = donor.find("w:pPr", NS)
        if ppr is not None:
            paragraph.append(ET.fromstring(ET.tostring(ppr)))
    ppr = ensure_child(paragraph, "pPr", before_tags={"r"})
    ensure_child(ppr, "keepNext")
    ind = ensure_child(ppr, "ind")
    for key in ("firstLine", "firstLineChars", "left", "right", "hanging", "hangingChars"):
        ind.set(q(key), "0")
    jc = ensure_child(ppr, "jc")
    jc.set(q("val"), "center")
    run = ET.SubElement(paragraph, q("r"))
    rpr = ET.SubElement(run, q("rPr"))
    sz = ET.SubElement(rpr, q("sz"))
    sz.set(q("val"), "21")
    szcs = ET.SubElement(rpr, q("szCs"))
    szcs.set(q("val"), "21")
    text = ET.SubElement(run, q("t"))
    text.text = title
    normalize_ppr_order(ppr)
    return paragraph


def normalize_table_title_paragraph(paragraph: ET.Element | None) -> None:
    if paragraph is None:
        return
    ppr = ensure_child(paragraph, "pPr", before_tags={"r"})
    for child in list(ppr):
        if child.tag in {q("numPr"), q("outlineLvl"), q("keepNext")}:
            ppr.remove(child)
    pstyle = ensure_child(ppr, "pStyle")
    pstyle.set(q("val"), "IMUSTTableCaption")
    ensure_child(ppr, "keepNext")
    spacing = ensure_child(ppr, "spacing")
    for key, value in (("before", "240"), ("after", "0"), ("line", "360"), ("lineRule", "auto")):
        spacing.set(q(key), value)
    ind = ensure_child(ppr, "ind")
    for key in ("firstLine", "firstLineChars", "left", "right", "leftChars", "rightChars", "hanging", "hangingChars"):
        ind.attrib.pop(q(key), None)
    jc = ensure_child(ppr, "jc")
    jc.set(q("val"), "center")
    for run in paragraph.findall("w:r", NS):
        if not compact(text_of(run)):
            continue
        rpr = ensure_child(run, "rPr", before_tags={"t", "drawing", "object", "pict"})
        sz = ensure_child(rpr, "sz")
        sz.set(q("val"), "21")
        szcs = ensure_child(rpr, "szCs")
        szcs.set(q("val"), "21")
    normalize_ppr_order(ppr)


def insert_missing_table_titles(
    body: ET.Element,
    children: list[ET.Element],
    insertions: list[dict[str, object]],
) -> list[dict[str, object]]:
    changes: list[dict[str, object]] = []
    for item in insertions:
        title = str(item.get("title") or "").strip()
        expected_cells = [str(value) for value in item.get("header_cells") or []]
        if not title or not expected_cells:
            changes.append({"kind": "table_title_insertion_skipped", "reason": "missing-title-or-header-cells"})
            continue
        matched = []
        for index, child in enumerate(children):
            if child.tag == q("tbl") and table_matches_header(child, expected_cells):
                matched.append(index)
        if len(matched) != 1:
            changes.append(
                {
                    "kind": "table_title_insertion_skipped",
                    "title": title,
                    "reason": "table-header-anchor-not-unique",
                    "match_count": len(matched),
                }
            )
            continue
        table_index = matched[0]
        _, previous_paragraph = previous_nonempty_paragraph(children, table_index)
        if TABLE_TITLE_RE.match(text_of(previous_paragraph)):
            changes.append(
                {
                    "kind": "table_title_insertion_skipped",
                    "title": title,
                    "reason": "table-already-has-formal-title",
                    "table_child_index": table_index,
                }
            )
            continue
        title_paragraph = clone_table_title_paragraph(previous_paragraph, title)
        body.insert(table_index, title_paragraph)
        children = list(body)
        changes.append(
            {
                "kind": "table_title_inserted",
                "title": title,
                "inserted_child_index": table_index,
                "table_child_index_after_insert": table_index + 1,
                "header_cells": expected_cells,
            }
        )
    return changes


def _load_json_list(path: str | None, key: str) -> list[dict[str, object]]:
    if not path:
        return []
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if isinstance(payload, dict):
        value = payload.get(key) or payload.get("items") or []
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    raise ValueError(f"{path} must contain a list or an object with {key}")


def repair_docx(input_docx: Path, output_docx: Path, title_insertions: list[dict[str, object]] | None = None) -> dict[str, object]:
    title_insertions = title_insertions or []
    with zipfile.ZipFile(input_docx, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml has no w:body")
        children = list(body)
        title_changes = insert_missing_table_titles(body, children, title_insertions)
        children = list(body)
        repaired_tables = 0
        for index, child in enumerate(children):
            if child.tag != q("tbl"):
                continue
            _, title_paragraph = previous_nonempty_paragraph(children, index)
            if not TABLE_TITLE_RE.match(text_of(title_paragraph)):
                continue
            normalize_table_title_paragraph(title_paragraph)
            repair_table(child)
            repaired_tables += 1
        output_docx.parent.mkdir(parents=True, exist_ok=True)
        tmp_output = output_docx.with_suffix(output_docx.suffix + ".tmp")
        with zipfile.ZipFile(tmp_output, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, data)
        shutil.move(str(tmp_output), str(output_docx))
    return {
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "changed_zip_parts": ["word/document.xml"],
        "table_title_changes": title_changes,
        "table_title_change_count": len(title_changes),
        "repaired_body_table_count": repaired_tables,
    }


def previous_nonempty_paragraph(children: list[ET.Element], index: int) -> tuple[int | None, ET.Element | None]:
    for prev in range(index - 1, -1, -1):
        if children[prev].tag == q("p") and compact(text_of(children[prev])):
            return prev, children[prev]
    return None, None


def audit_docx(docx: Path, rendered_pages: list[str] | None = None) -> dict[str, object]:
    children = load_body_children(docx)
    tables: list[dict[str, object]] = []
    for index, child in enumerate(children):
        if child.tag != q("tbl"):
            continue
        title_index, title_paragraph = previous_nonempty_paragraph(children, index)
        title = text_of(title_paragraph)
        if not TABLE_TITLE_RE.match(title):
            continue
        rows = child.findall("w:tr", NS)
        first_row = rows[0] if rows else None
        last_row = rows[-1] if rows else None
        first_row_borders = [cell_borders(cell) for cell in first_row.findall("w:tc", NS)] if first_row is not None else []
        last_row_borders = [cell_borders(cell) for cell in last_row.findall("w:tc", NS)] if last_row is not None else []
        cell_sizes = sorted(
            {
                size
                for row in rows
                for cell in row.findall("w:tc", NS)
                for paragraph in cell.findall("w:p", NS)
                for run in paragraph.findall("w:r", NS)
                for size in [run_size(run)]
                if size is not None and compact(text_of(run))
            }
        )
        nonzero_indent_count = 0
        for row in rows:
            for cell in row.findall("w:tc", NS):
                for paragraph in cell.findall("w:p", NS):
                    if not compact(text_of(paragraph)):
                        continue
                    ind = paragraph_info(paragraph)["ind"]
                    if isinstance(ind, dict) and any(value not in ("0", None) for value in ind.values()):
                        nonzero_indent_count += 1
        first_top = [border.get("top", {}).get("val") for border in first_row_borders]
        header_bottom = [border.get("bottom", {}).get("val") for border in first_row_borders]
        last_bottom = [border.get("bottom", {}).get("val") for border in last_row_borders]
        visible_verticals = []
        unexpected_body_horizontal_borders = []
        title_indent = paragraph_info(title_paragraph)["ind"]
        title_zero_indent = not (
            isinstance(title_indent, dict)
            and any(value not in ("0", None) for value in title_indent.values())
        )
        for row in rows:
            row_index = rows.index(row)
            is_first_row = row_index == 0
            is_last_row = row_index == len(rows) - 1
            for cell_index, border in enumerate([cell_borders(cell) for cell in row.findall("w:tc", NS)]):
                for side in ("left", "right"):
                    if border.get(side, {}).get("val") not in (None, "nil", "none"):
                        visible_verticals.append(border)
                top_val = border.get("top", {}).get("val")
                bottom_val = border.get("bottom", {}).get("val")
                if not is_first_row and top_val not in (None, "nil", "none"):
                    unexpected_body_horizontal_borders.append(
                        {"row_index": row_index, "cell_index": cell_index, "side": "top", "val": top_val}
                    )
                if not (is_first_row or is_last_row) and bottom_val not in (None, "nil", "none"):
                    unexpected_body_horizontal_borders.append(
                        {"row_index": row_index, "cell_index": cell_index, "side": "bottom", "val": bottom_val}
                    )
        row_flag_list = [row_flags(row) for row in rows]
        table_record = {
            "table_child_index": index,
            "title_child_index": title_index,
            "title": title,
            "title_info": paragraph_info(title_paragraph),
            "row_count": len(rows),
            "column_count_first_row": len(first_row.findall("w:tc", NS)) if first_row is not None else 0,
            "table_cell_margins": table_cell_margins(child),
            "row_flags": row_flag_list,
            "cell_run_sizes": cell_sizes,
            "nonzero_cell_indent_count": nonzero_indent_count,
            "title_keep_next_verdict": bool(paragraph_info(title_paragraph)["keepNext"]),
            "header_repeat_verdict": bool(first_row is not None and row_flags(first_row)["tblHeader"]),
            "row_cant_split_verdict": all(flag["cantSplit"] for flag in row_flag_list),
            "cell_size_5hao_verdict": cell_sizes == ["21"],
            "cell_zero_indent_verdict": nonzero_indent_count == 0,
            "visible_border_structure_verdict": (
                all(value == "single" for value in first_top)
                and all(value == "single" for value in header_bottom)
                and all(value == "single" for value in last_bottom)
                and not visible_verticals
                and not unexpected_body_horizontal_borders
            ),
            "title_zero_indent_verdict": title_zero_indent,
            "unexpected_body_horizontal_borders": unexpected_body_horizontal_borders,
            "unexpected_body_horizontal_border_count": len(unexpected_body_horizontal_borders),
        }
        tables.append(table_record)
    summary = {
        "body_table_count": len(tables),
        "all_table_title_keep_next": all(t["title_keep_next_verdict"] for t in tables),
        "all_table_visible_border_structure": all(t["visible_border_structure_verdict"] for t in tables),
        "all_table_title_zero_indent": all(t["title_zero_indent_verdict"] for t in tables),
        "all_table_header_repeat": all(t["header_repeat_verdict"] for t in tables),
        "all_table_rows_cant_split": all(t["row_cant_split_verdict"] for t in tables),
        "all_table_cell_5hao": all(t["cell_size_5hao_verdict"] for t in tables),
        "all_table_zero_indent": all(t["cell_zero_indent_verdict"] for t in tables),
        "rendered_pages": rendered_pages or [],
        "cross_page_table_summary": "explicit continuation evidence must be supplied by rendered-page review",
    }
    passed = bool(tables) and all(
        bool(t[key])
        for t in tables
        for key in (
            "title_keep_next_verdict",
            "title_zero_indent_verdict",
            "visible_border_structure_verdict",
            "header_repeat_verdict",
            "row_cant_split_verdict",
            "cell_size_5hao_verdict",
            "cell_zero_indent_verdict",
        )
    )
    return {
        "schema": "graduation-project-builder.docx-table-structure.v1",
        "docx": str(docx),
        "docx_sha256": sha256_file(docx),
        "passed": passed,
        "summary": summary,
        "body_tables": tables,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--report-json")
    parser.add_argument("--rendered-pages", default="")
    parser.add_argument("--fail-on-drift", action="store_true")
    parser.add_argument("--repair-output-docx")
    parser.add_argument("--title-insertions-json")
    args = parser.parse_args(argv)
    rendered_pages = [part.strip() for part in re.split(r"[;\n]", args.rendered_pages) if part.strip()]
    if args.repair_output_docx:
        repair_summary = repair_docx(
            Path(args.docx),
            Path(args.repair_output_docx),
            _load_json_list(args.title_insertions_json, "insertions"),
        )
        report = audit_docx(Path(args.repair_output_docx), rendered_pages)
        report["repair"] = repair_summary
    else:
        report = audit_docx(Path(args.docx), rendered_pages)
    if args.report_json:
        output = Path(args.report_json)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"passed": report["passed"], **report["summary"]}, ensure_ascii=True))
    return 1 if args.fail_on_drift and not report["passed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
