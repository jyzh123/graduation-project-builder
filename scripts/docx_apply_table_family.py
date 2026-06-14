from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]
ET.register_namespace("w", NS["w"])
CN_NUMBER_CHARS = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u96f6\u3007\u58f9\u8d30\u53c1\u8086\u4f0d\u9646\u67d2\u634c\u7396"
TABLE_TITLE_RE = re.compile(rf"^\s*(?:\u8868|\u7eed\u8868)\s*[0-9{CN_NUMBER_CHARS}]+(?:[-.]\d+)?\s*\S")
BODY_HEADING_RE = re.compile(r"^\s*(?:\d{1,2}\s+\S|\d{1,2}\.\d+\s+\S|\u7b2c[0-9\u4e00-\u9fff]+\u7ae0)")


def w_attr(name: str) -> str:
    return W + name


def child(parent: ET._Element | None, tag: str) -> ET._Element | None:
    return parent.find(tag) if parent is not None else None


def children(parent: ET._Element | None, tag: str) -> list[ET._Element]:
    return parent.findall(tag) if parent is not None else []


def ensure(parent: ET._Element, tag: str) -> ET._Element:
    node = parent.find(tag)
    if node is None:
        node = ET.SubElement(parent, tag)
    return node


TBL_PR_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
]

TC_PR_ORDER = [
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
]


def local_name(tag: str) -> str:
    return ET.QName(tag).localname


def insert_ordered(parent: ET._Element, node: ET._Element, order: list[str]) -> None:
    target_name = local_name(node.tag)
    target_index = order.index(target_name) if target_name in order else len(order)
    for idx, existing in enumerate(parent):
        existing_name = local_name(existing.tag)
        existing_index = order.index(existing_name) if existing_name in order else len(order)
        if existing_index > target_index:
            parent.insert(idx, node)
            return
    parent.append(node)


def ensure_ordered(parent: ET._Element, tag: str, order: list[str]) -> ET._Element:
    node = parent.find(tag)
    if node is None:
        node = ET.Element(tag)
        insert_ordered(parent, node, order)
    return node


def border(tag: str, val: str, sz: str, color: str = "000000") -> ET._Element:
    return ET.Element(
        tag,
        {
            w_attr("val"): val,
            w_attr("sz"): sz,
            w_attr("color"): color,
            w_attr("space"): "0",
        },
    )


def clear_shading(node: ET._Element) -> None:
    for shd in list(node.findall(f".//{W}shd")):
        parent = shd.getparent()
        if parent is not None:
            parent.remove(shd)


def replace_first(parent: ET._Element, tag: str, replacement: ET._Element | None, default_index: int = 0) -> None:
    current = parent.find(tag)
    if current is not None:
        idx = parent.index(current)
        parent.remove(current)
        if replacement is not None:
            parent.insert(idx, deepcopy(replacement))
    elif replacement is not None:
        parent.insert(min(default_index, len(parent)), deepcopy(replacement))


def sanitized_paragraph_properties(ppr: ET._Element | None) -> ET._Element | None:
    if ppr is None:
        return None
    clone = deepcopy(ppr)
    for tag in (W + "pStyle", W + "numPr"):
        node = clone.find(tag)
        if node is not None:
            clone.remove(node)
    return clone


def sanitized_run_properties(rpr: ET._Element | None) -> ET._Element | None:
    if rpr is None:
        return None
    clone = deepcopy(rpr)
    for tag in (W + "rStyle",):
        node = clone.find(tag)
        if node is not None:
            clone.remove(node)
    return clone


def remove_first(parent: ET._Element, tag: str) -> None:
    current = parent.find(tag)
    if current is not None:
        parent.remove(current)


def element_hash(node: ET._Element | None) -> str | None:
    if node is None:
        return None
    data = ET.tostring(node, encoding="utf-8")
    return hashlib.sha256(data).hexdigest()[:16]


def text_of(node: ET._Element) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


def set_paragraph_text_from_donor(target_p: ET._Element, donor_p: ET._Element, text: str) -> None:
    ppr = sanitized_paragraph_properties(child(donor_p, W + "pPr"))
    replace_first(target_p, W + "pPr", ppr, 0)
    donor_run = next((r for r in children(donor_p, W + "r") if text_of(r).strip()), None)
    donor_rpr = sanitized_run_properties(child(donor_run, W + "rPr") if donor_run is not None else None)
    for item in list(target_p):
        if item.tag == W + "r":
            target_p.remove(item)
    run = ET.Element(W + "r")
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    t = ET.Element(W + "t")
    t.text = text
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run.append(t)
    target_p.append(run)


def apply_paragraph_face(target_p: ET._Element, donor_p: ET._Element) -> None:
    replace_first(target_p, W + "pPr", sanitized_paragraph_properties(child(donor_p, W + "pPr")), 0)
    donor_run = next((r for r in children(donor_p, W + "r") if text_of(r).strip()), None)
    donor_rpr = sanitized_run_properties(child(donor_run, W + "rPr") if donor_run is not None else None)
    for run in children(target_p, W + "r"):
        replace_first(run, W + "rPr", donor_rpr, 0)


def table_header(tbl: ET._Element) -> tuple[str, ...]:
    title_mode, header_row = table_title_mode(tbl)
    rows = children(tbl, W + "tr")
    if not rows:
        return ()
    if title_mode == "first_merged_row" and header_row is not None:
        return tuple(text_of(tc).strip() for tc in children(header_row, W + "tc"))
    return tuple(text_of(tc).strip() for tc in children(rows[0], W + "tc"))


def table_title_mode(tbl: ET._Element) -> tuple[str, ET._Element | None]:
    rows = children(tbl, W + "tr")
    if not rows:
        return "none", None
    first_row = rows[0]
    first_cells = children(first_row, W + "tc")
    if not first_cells:
        return "none", None
    first_cell = first_cells[0]
    paragraphs = [p for p in children(first_cell, W + "p") if text_of(p).strip()]
    if len(first_cells) == 1 and any(is_table_title_paragraph(paragraph) for paragraph in paragraphs):
        return "first_merged_row", first_row
    return "external_paragraph", None


def is_table_title_paragraph(paragraph: ET._Element | None) -> bool:
    return paragraph is not None and bool(TABLE_TITLE_RE.match(text_of(paragraph).strip()))


def is_body_heading_paragraph(paragraph: ET._Element) -> bool:
    return bool(BODY_HEADING_RE.match(text_of(paragraph).strip()))


def body_start_child_index(body: ET._Element) -> int | None:
    for index, node in enumerate(body):
        if node.tag == W + "p" and is_body_heading_paragraph(node):
            return index
    return None


def cell_text_contamination_issues(tbl: ET._Element, table_index: int, *, title_mode: str = "external_paragraph") -> list[str]:
    issues: list[str] = []
    rows = children(tbl, W + "tr")
    for cell_index, cell in enumerate(tbl.findall(".//w:tc", NS), 1):
        if title_mode == "first_merged_row" and rows:
            first_row_cells = children(rows[0], W + "tc")
            if first_row_cells and cell in first_row_cells:
                continue
        paragraphs = [p for p in children(cell, W + "p") if text_of(p).strip()]
        for paragraph_index, paragraph in enumerate(paragraphs, 1):
            text = text_of(paragraph).strip()
            if is_table_title_paragraph(paragraph):
                issues.append(
                    f"target table {table_index} cell {cell_index} paragraph {paragraph_index} contains a table title"
                )
            if re.match(r"^\s*(?:\u56fe|figure|fig\.)\s*\d", text, re.I):
                issues.append(
                    f"target table {table_index} cell {cell_index} paragraph {paragraph_index} contains a figure caption"
                )
            if len(text) >= 80 and any(mark in text for mark in ("\u3002", ".", "\uff1b", ";")):
                issues.append(
                    f"target table {table_index} cell {cell_index} paragraph {paragraph_index} looks like body prose"
                )
        if len(paragraphs) > 2 and any(len(text_of(p).strip()) >= 40 for p in paragraphs):
            issues.append(f"target table {table_index} cell {cell_index} contains multiple long paragraphs")
    return issues


def find_table_contexts(root: ET._Element, *, body_only: bool = True) -> list[dict[str, object]]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    contexts: list[dict[str, object]] = []
    start_index = body_start_child_index(body) if body_only else None
    for node_index, node in enumerate(body):
        if body_only and start_index is not None and node_index < start_index:
            continue
        if body_only and start_index is None:
            continue
        if node.tag == W + "p":
            continue
        elif node.tag == W + "tbl":
            previous = body[node_index - 1] if node_index > 0 else None
            title_p = previous if previous is not None and previous.tag == W + "p" and is_table_title_paragraph(previous) else None
            title_mode = "external_paragraph" if title_p is not None else "none"
            if title_p is None:
                rows = children(node, W + "tr")
                if rows:
                    first_row_cells = children(rows[0], W + "tc")
                    if len(first_row_cells) == 1:
                        first_cell = first_row_cells[0]
                        title_paragraph = next((p for p in children(first_cell, W + "p") if is_table_title_paragraph(p)), None)
                        if title_paragraph is not None:
                            title_p = title_paragraph
                            title_mode = "first_merged_row"
            contexts.append(
                {
                    "tbl": node,
                    "title_p": title_p,
                    "title_mode": title_mode,
                    "node_index": node_index,
                    "body_only": body_only,
                    "structure_issues": cell_text_contamination_issues(node, len(contexts) + 1, title_mode=title_mode),
                }
            )
    return contexts


def lift_first_cell_table_titles(
    root: ET._Element,
    *,
    body_only: bool = True,
    donor_title_modes_by_header: dict[tuple[str, ...], str] | None = None,
) -> list[dict[str, object]]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    changes: list[dict[str, object]] = []
    start_index = body_start_child_index(body) if body_only else None
    table_index = 0
    for node_index, node in list(enumerate(body)):
        if node.tag != W + "tbl":
            continue
        if body_only and start_index is not None and node_index < start_index:
            continue
        if body_only and start_index is None:
            continue
        table_index += 1
        previous = body[node_index - 1] if node_index > 0 else None
        if previous is not None and previous.tag == W + "p" and is_table_title_paragraph(previous):
            continue
        title_mode, _ = table_title_mode(node)
        if title_mode != "first_merged_row":
            continue
        donor_mode = None
        if donor_title_modes_by_header is not None:
            donor_mode = donor_title_modes_by_header.get(table_header(node))
        if donor_mode not in {None, "external_paragraph"}:
            continue
        rows = children(node, W + "tr")
        if not rows:
            continue
        cells = children(rows[0], W + "tc")
        if not cells:
            continue
        first_cell = cells[0]
        for paragraph in list(children(first_cell, W + "p")):
            if not is_table_title_paragraph(paragraph):
                continue
            title_text = text_of(paragraph).strip()
            title_paragraph = deepcopy(paragraph)
            first_cell.remove(paragraph)
            if not children(first_cell, W + "p"):
                first_cell.append(ET.Element(W + "p"))
            body.insert(body.index(node), title_paragraph)
            changes.append(
                {
                    "target_table_index": table_index,
                    "lifted_title": title_text,
                    "source": "first-header-cell",
                }
            )
            break
    return changes


def first_or_last(items: list[ET._Element], index: int) -> ET._Element:
    if not items:
        raise ValueError("donor table has no rows")
    return items[min(max(index, 0), len(items) - 1)]


def row_pattern(donor_rows: list[ET._Element], target_index: int, target_count: int) -> ET._Element:
    if target_index == 0:
        return first_or_last(donor_rows, 0)
    if target_index == target_count - 1 and len(donor_rows) > 2:
        return donor_rows[-1]
    return first_or_last(donor_rows, min(target_index, max(len(donor_rows) - 2, 1)))


def apply_cell_face(target_tc: ET._Element, donor_tc: ET._Element) -> None:
    replace_first(target_tc, W + "tcPr", child(donor_tc, W + "tcPr"), 0)
    donor_p = next((p for p in children(donor_tc, W + "p") if text_of(p).strip()), None)
    if donor_p is None:
        donor_p = child(donor_tc, W + "p")
    if donor_p is None:
        return
    for p in children(target_tc, W + "p"):
        apply_paragraph_face(p, donor_p)


def apply_table_from_donor(target_tbl: ET._Element, donor_tbl: ET._Element) -> None:
    replace_first(target_tbl, W + "tblPr", child(donor_tbl, W + "tblPr"), 0)
    replace_first(target_tbl, W + "tblGrid", child(donor_tbl, W + "tblGrid"), 1)
    target_rows = children(target_tbl, W + "tr")
    donor_rows = children(donor_tbl, W + "tr")
    for row_index, target_row in enumerate(target_rows):
        donor_row = row_pattern(donor_rows, row_index, len(target_rows))
        replace_first(target_row, W + "trPr", child(donor_row, W + "trPr"), 0)
        target_cells = children(target_row, W + "tc")
        donor_cells = children(donor_row, W + "tc")
        for cell_index, target_tc in enumerate(target_cells):
            donor_tc = first_or_last(donor_cells, cell_index)
            apply_cell_face(target_tc, donor_tc)


def ensure_ppr(paragraph: ET._Element) -> ET._Element:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        paragraph.insert(0, ppr)
    return ppr


def ensure_rpr(run: ET._Element) -> ET._Element:
    rpr = run.find(W + "rPr")
    if rpr is None:
        rpr = ET.Element(W + "rPr")
        run.insert(0, rpr)
    return rpr


def ensure_rfonts(rpr: ET._Element) -> ET._Element:
    rfonts = rpr.find(W + "rFonts")
    if rfonts is None:
        rfonts = ET.Element(W + "rFonts")
        rpr.insert(0, rfonts)
    return rfonts


SONG = "\u5b8b\u4f53"
TIMES_NR = "TimesNewRoman"


TABLE_CAPTION_PARAGRAPH_BASELINE = {
    "align": "center",
    "before": "",
    "after": "3",
    "line": "265",
    "lineRule": "auto",
    "left": "10",
    "right": "122",
    "hanging": "10",
    "firstLine": "",
    "firstLineChars": "",
    "leftChars": "",
    "rightChars": "",
    "hangingChars": "",
}

TABLE_HEADER_CELL_PARAGRAPH_BASELINE = {
    "align": "center",
    "before": "",
    "after": "0",
    "line": "",
    "lineRule": "",
    "left": "",
    "right": "",
    "hanging": "",
    "firstLine": "",
    "firstLineChars": "",
    "leftChars": "",
    "rightChars": "",
    "hangingChars": "",
}

TABLE_BODY_CELL_PARAGRAPH_BASELINE = dict(TABLE_HEADER_CELL_PARAGRAPH_BASELINE)


def normalize_table_body_numeric_cells(
    tbl: ET._Element,
    *,
    east_asia: str = SONG,
    ascii_font: str = SONG,
    hansi_font: str = SONG,
    size_half_points: str = "21",
) -> dict[str, object]:
    changed_cells: list[dict[str, object]] = []
    rows = children(tbl, W + "tr")
    for row_index, row in enumerate(rows):
        if row_index == 0:
            continue
        for cell_index, cell in enumerate(children(row, W + "tc")):
            cell_text = text_of(cell).strip()
            if not re.fullmatch(r"\d+(?:\.\d+)?", cell_text):
                continue
            cell_changed = False
            for paragraph in children(cell, W + "p"):
                ppr = ensure_ppr(paragraph)
                jc = ppr.find(W + "jc")
                if jc is None:
                    jc = ET.Element(W + "jc")
                    ppr.append(jc)
                if jc.get(w_attr("val")) != "center":
                    jc.set(w_attr("val"), "center")
                    cell_changed = True
                for run in children(paragraph, W + "r"):
                    if not text_of(run).strip():
                        continue
                    rpr = ensure_rpr(run)
                    rfonts = ensure_rfonts(rpr)
                    for key, value in {
                        "eastAsia": east_asia,
                        "ascii": ascii_font,
                        "hAnsi": hansi_font,
                    }.items():
                        if rfonts.get(w_attr(key)) != value:
                            rfonts.set(w_attr(key), value)
                            cell_changed = True
                    for tag in (W + "sz", W + "szCs"):
                        node = rpr.find(tag)
                        if node is None:
                            node = ET.Element(tag)
                            rpr.append(node)
                        if node.get(w_attr("val")) != size_half_points:
                            node.set(w_attr("val"), size_half_points)
                            cell_changed = True
            if cell_changed:
                changed_cells.append(
                    {"row_index": row_index + 1, "cell_index": cell_index + 1, "text": cell_text}
                )
    return {
        "changed_cell_count": len(changed_cells),
        "changed_cells": changed_cells,
        "eastAsia": east_asia,
        "ascii": ascii_font,
        "hAnsi": hansi_font,
        "size_half_points": size_half_points,
    }


def remove_child(parent: ET._Element, tag: str) -> None:
    current = parent.find(tag)
    if current is not None:
        parent.remove(current)


def set_or_clear_attr(node: ET._Element, key: str, value: str | None) -> None:
    attr = w_attr(key)
    if value in {None, ""}:
        node.attrib.pop(attr, None)
    else:
        node.set(attr, str(value))


def set_paragraph_direct_metrics(paragraph: ET._Element, baseline: dict[str, str]) -> None:
    ppr = ensure_ppr(paragraph)
    for tag in (W + "tabs", W + "outlineLvl", W + "numPr"):
        node = ppr.find(tag)
        if node is not None:
            ppr.remove(node)

    ind_keys = ("left", "right", "hanging", "firstLine", "firstLineChars", "leftChars", "rightChars", "hangingChars")
    ind_values = {key: baseline.get(key, "") for key in ind_keys}
    ind = ppr.find(W + "ind")
    if any(value not in {"", None} for value in ind_values.values()):
        if ind is None:
            ind = ET.Element(W + "ind")
            ppr.append(ind)
        for key, value in ind_values.items():
            set_or_clear_attr(ind, key, value)
    elif ind is not None:
        ppr.remove(ind)

    spacing_keys = ("before", "after", "line", "lineRule")
    spacing_values = {key: baseline.get(key, "") for key in spacing_keys}
    spacing = ppr.find(W + "spacing")
    if any(value not in {"", None} for value in spacing_values.values()):
        if spacing is None:
            spacing = ET.Element(W + "spacing")
            ppr.append(spacing)
        for key, value in spacing_values.items():
            set_or_clear_attr(spacing, key, value)
    elif spacing is not None:
        ppr.remove(spacing)

    jc = ppr.find(W + "jc")
    align = baseline.get("align", "")
    if align:
        if jc is None:
            jc = ET.Element(W + "jc")
            ppr.append(jc)
        jc.set(w_attr("val"), align)
    elif jc is not None:
        ppr.remove(jc)


def set_paragraph_cell_metrics(paragraph: ET._Element, *, header: bool = False) -> None:
    set_paragraph_direct_metrics(
        paragraph,
        TABLE_HEADER_CELL_PARAGRAPH_BASELINE if header else TABLE_BODY_CELL_PARAGRAPH_BASELINE,
    )


def set_run_font_metrics(
    run: ET._Element,
    *,
    east_asia: str = SONG,
    ascii_font: str = SONG,
    hansi_font: str = SONG,
    cs_font: str = SONG,
    size_half_points: str = "21",
    bold: bool | None = None,
) -> None:
    rpr = ensure_rpr(run)
    rfonts = ensure_rfonts(rpr)
    for key, value in {
        "eastAsia": east_asia,
        "ascii": ascii_font,
        "hAnsi": hansi_font,
        "cs": cs_font,
    }.items():
        rfonts.set(w_attr(key), value)
        rfonts.attrib.pop(w_attr(f"{key}Theme"), None)
    for tag in (W + "sz", W + "szCs"):
        node = rpr.find(tag)
        if node is None:
            node = ET.Element(tag)
            rpr.append(node)
        node.set(w_attr("val"), size_half_points)
    if bold is not None:
        for tag in (W + "b", W + "bCs"):
            node = rpr.find(tag)
            if bold:
                if node is None:
                    node = ET.Element(tag)
                    rpr.append(node)
                node.attrib.pop(w_attr("val"), None)
            elif node is not None:
                rpr.remove(node)


def visible_run_text(run: ET._Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//" + W + "t"))


def text_has_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def text_has_latin_or_digit(text: str) -> bool:
    return any(("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch.isdigit() for ch in text)


def split_script_chunks(text: str) -> list[tuple[str, str]]:
    chunks: list[tuple[str, str]] = []
    current_kind = ""
    current = ""
    for ch in text:
        if "\u4e00" <= ch <= "\u9fff":
            kind = "cjk"
        elif ("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch.isdigit():
            kind = "latin"
        else:
            kind = current_kind or "neutral"
        if current and kind != current_kind and kind != "neutral":
            chunks.append((current_kind or "neutral", current))
            current = ch
            current_kind = kind
        else:
            current += ch
            if kind != "neutral":
                current_kind = kind
    if current:
        chunks.append((current_kind or "neutral", current))
    return chunks


def replace_run_text(run: ET._Element, text: str) -> None:
    text_nodes = run.findall(".//" + W + "t")
    if not text_nodes:
        t = ET.Element(W + "t")
        run.append(t)
        text_nodes = [t]
    text_nodes[0].text = text
    if text.startswith(" ") or text.endswith(" "):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in text_nodes[1:]:
        node.text = ""


def set_caption_run_metrics(run: ET._Element, kind: str) -> None:
    if kind == "latin":
        set_run_font_metrics(
            run,
            east_asia=TIMES_NR,
            ascii_font=TIMES_NR,
            hansi_font=TIMES_NR,
            cs_font=TIMES_NR,
            size_half_points="21",
            bold=False,
        )
    else:
        set_run_font_metrics(
            run,
            east_asia=SONG,
            ascii_font=SONG,
            hansi_font=SONG,
            cs_font=SONG,
            size_half_points="21",
            bold=False,
        )


def normalize_caption_runs(paragraph: ET._Element) -> None:
    for run in list(children(paragraph, W + "r")):
        text = visible_run_text(run)
        if not text.strip():
            continue
        chunks = split_script_chunks(text)
        if len(chunks) <= 1:
            kind = "latin" if text_has_latin_or_digit(text) and not text_has_cjk(text) else "cjk"
            set_caption_run_metrics(run, kind)
            continue
        parent = run.getparent()
        if parent is None:
            continue
        insert_at = parent.index(run)
        parent.remove(run)
        for offset, (kind, chunk_text) in enumerate(chunks):
            new_run = deepcopy(run)
            replace_run_text(new_run, chunk_text)
            set_caption_run_metrics(new_run, "latin" if kind == "latin" else "cjk")
            parent.insert(insert_at + offset, new_run)


def normalize_body_table_structure(tbl: ET._Element) -> dict[str, object]:
    rows = children(tbl, W + "tr")
    border_result = normalize_three_line_table_borders(tbl)
    body_text_cells = 0
    for row_index, row in enumerate(rows):
        trpr = ensure(row, W + "trPr")
        if trpr.find(W + "cantSplit") is None:
            trpr.append(ET.Element(W + "cantSplit"))
        tbl_header = trpr.find(W + "tblHeader")
        if row_index == 0:
            if tbl_header is None:
                trpr.append(ET.Element(W + "tblHeader"))
        elif tbl_header is not None:
            trpr.remove(tbl_header)
        for cell in children(row, W + "tc"):
            tcpr = ensure(cell, W + "tcPr")
            tc_borders = ensure_ordered(tcpr, W + "tcBorders", TC_PR_ORDER)
            tc_borders.clear()
            if row_index == 0:
                tc_borders.extend(
                    [
                        border(W + "top", "single", "4"),
                        border(W + "left", "none", "0"),
                        border(W + "bottom", "single", "4"),
                        border(W + "right", "none", "0"),
                    ]
                )
            elif row_index == len(rows) - 1:
                tc_borders.extend(
                    [
                        border(W + "top", "none", "0"),
                        border(W + "left", "none", "0"),
                        border(W + "bottom", "single", "4"),
                        border(W + "right", "none", "0"),
                    ]
                )
            else:
                tc_borders.extend(
                    [
                        border(W + "top", "none", "0"),
                        border(W + "left", "none", "0"),
                        border(W + "bottom", "none", "0"),
                        border(W + "right", "none", "0"),
                    ]
                )
            for paragraph in children(cell, W + "p"):
                set_paragraph_cell_metrics(paragraph, header=(row_index == 0))
                for run in children(paragraph, W + "r"):
                    if not text_of(run).strip():
                        continue
                    set_run_font_metrics(
                        run,
                        size_half_points="21",
                        bold=False,
                    )
                    body_text_cells += 1
    return {
        **border_result,
        "row_count": len(rows),
        "formatted_text_run_count": body_text_cells,
        "header_repeat_set": bool(rows),
        "cant_split_rows_set": len(rows),
    }


def ensure_table_title_keep_next(title_p: ET._Element | None) -> bool:
    if title_p is None:
        return False
    ppr = ensure_ppr(title_p)
    keep_next = ppr.find(W + "keepNext")
    if keep_next is None:
        ppr.insert(0, ET.Element(W + "keepNext"))
        return True
    return False


def normalize_table_title(title_p: ET._Element | None) -> bool:
    if title_p is None:
        return False
    changed_keep_next = ensure_table_title_keep_next(title_p)
    set_paragraph_direct_metrics(title_p, TABLE_CAPTION_PARAGRAPH_BASELINE)
    normalize_caption_runs(title_p)
    return changed_keep_next


def patch_all_body_tables(docx_path: Path, output_path: Path | None = None) -> dict[str, object]:
    out_path = output_path or docx_path
    if output_path:
        shutil.copy2(docx_path, out_path)
    members, root = load_docx_xml(out_path)
    body_tables = find_table_contexts(root, body_only=True)
    table_reports = []
    title_keep_next_updates: list[int] = []
    for table_index, context in enumerate(body_tables, 1):
        tbl = context["tbl"]
        if normalize_table_title(context.get("title_p")):  # type: ignore[arg-type]
            title_keep_next_updates.append(table_index)
        result = normalize_body_table_structure(tbl)  # type: ignore[arg-type]
        table_reports.append({"table_index": table_index, **result})
    write_docx_xml(out_path, members, root)
    return {
        "source_docx": str(docx_path),
        "output_docx": str(out_path),
        "table_count": len(table_reports),
        "title_keep_next_updates": title_keep_next_updates,
        "tables": table_reports,
    }


def normalize_three_line_table_borders(tbl: ET._Element) -> dict[str, object]:
    tbl_pr = ensure(tbl, W + "tblPr")
    clear_shading(tbl_pr)
    tbl_style = tbl_pr.find(W + "tblStyle")
    removed_style = False
    if tbl_style is not None:
        tbl_pr.remove(tbl_style)
        removed_style = True
    tbl_layout = ensure_ordered(tbl_pr, W + "tblLayout", TBL_PR_ORDER)
    tbl_layout.set(w_attr("type"), "fixed")
    tbl_jc = ensure_ordered(tbl_pr, W + "jc", TBL_PR_ORDER)
    tbl_jc.set(w_attr("val"), "center")
    tbl_borders = ensure_ordered(tbl_pr, W + "tblBorders", TBL_PR_ORDER)
    tbl_borders.clear()
    tbl_borders.extend(
        [
            border(W + "top", "single", "4"),
            border(W + "left", "none", "0"),
            border(W + "bottom", "single", "4"),
            border(W + "right", "none", "0"),
            border(W + "insideH", "none", "0"),
            border(W + "insideV", "none", "0"),
        ]
    )
    header_cell_count = 0
    rows = children(tbl, W + "tr")
    if rows:
        for cell in children(rows[0], W + "tc"):
            header_cell_count += 1
            tc_pr = ensure(cell, W + "tcPr")
            clear_shading(tc_pr)
            tc_borders = ensure_ordered(tc_pr, W + "tcBorders", TC_PR_ORDER)
            tc_borders.clear()
            tc_borders.extend(
                [
                    border(W + "top", "single", "4"),
                    border(W + "left", "none", "0"),
                    border(W + "bottom", "single", "4"),
                    border(W + "right", "none", "0"),
                ]
            )
    return {"removed_table_style": removed_style, "header_cell_border_count": header_cell_count}


def normalize_table_caption_number(text: str) -> str:
    return re.sub(r"^表\s*(\d+)[-.](\d+)", r"表\1-\2", text.strip())


def border_summary(tbl: ET._Element) -> dict[str, dict[str, str | None]]:
    result: dict[str, dict[str, str | None]] = {}
    borders = child(child(tbl, W + "tblPr"), W + "tblBorders")
    if borders is None:
        return result
    for item in borders:
        result[ET.QName(item).localname] = {
            "val": item.get(w_attr("val")),
            "sz": item.get(w_attr("sz")),
            "color": item.get(w_attr("color")),
        }
    return result


def table_metrics(tbl: ET._Element, title_p: ET._Element | None) -> dict[str, object]:
    tbl_pr = child(tbl, W + "tblPr")
    tbl_w = child(tbl_pr, W + "tblW")
    tbl_style = child(tbl_pr, W + "tblStyle")
    tbl_layout = child(tbl_pr, W + "tblLayout")
    jc = child(tbl_pr, W + "jc")
    grid = child(tbl, W + "tblGrid")
    rows = children(tbl, W + "tr")
    row_heights = []
    for row in rows:
        trh = child(child(row, W + "trPr"), W + "trHeight")
        row_heights.append(
            {
                "val": trh.get(w_attr("val")) if trh is not None else None,
                "hRule": trh.get(w_attr("hRule")) if trh is not None else None,
            }
        )
    first_p = None
    first_r = None
    if rows and children(rows[0], W + "tc"):
        first_p = child(children(rows[0], W + "tc")[0], W + "p")
        first_r = child(first_p, W + "r") if first_p is not None else None
    return {
        "title_text": text_of(title_p).strip() if title_p is not None else "",
        "title_mode": "first_merged_row" if title_p is not None and title_p.getparent() is not None and title_p.getparent().tag == W + "tc" else "external_paragraph" if title_p is not None else "none",
        "title_pPr_hash": element_hash(child(title_p, W + "pPr") if title_p is not None else None),
        "title_rPr_hash": element_hash(child(next((r for r in children(title_p, W + "r") if text_of(r).strip()), None), W + "rPr") if title_p is not None else None),
        "rows": len(rows),
        "cols": len(children(rows[0], W + "tc")) if rows else 0,
        "header": list(table_header(tbl)),
        "tblStyle": tbl_style.get(w_attr("val")) if tbl_style is not None else None,
        "tblW": {
            "w": tbl_w.get(w_attr("w")) if tbl_w is not None else None,
            "type": tbl_w.get(w_attr("type")) if tbl_w is not None else None,
        },
        "jc": jc.get(w_attr("val")) if jc is not None else None,
        "tblLayout": tbl_layout.get(w_attr("type")) if tbl_layout is not None else None,
        "grid": [col.get(w_attr("w")) for col in children(grid, W + "gridCol")],
        "borders": border_summary(tbl),
        "row_heights": row_heights,
        "first_header_pPr_hash": element_hash(child(first_p, W + "pPr")),
        "first_header_rPr_hash": element_hash(child(first_r, W + "rPr")),
        "tblPr_hash": element_hash(tbl_pr),
        "tblGrid_hash": element_hash(grid),
    }


def load_docx_xml(path: Path) -> tuple[dict[str, bytes], ET._Element]:
    with zipfile.ZipFile(path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}
    return members, ET.fromstring(members["word/document.xml"])


def collect_style_ids(node: ET._Element | None) -> set[str]:
    if node is None:
        return set()
    ids: set[str] = set()
    for elem in node.iter():
        if elem.tag in {W + "pStyle", W + "rStyle", W + "tblStyle"}:
            value = elem.get(w_attr("val"))
            if value:
                ids.add(value)
    return ids


def style_hash(style: ET._Element | None) -> str | None:
    if style is None:
        return None
    return hashlib.sha256(ET.tostring(style, encoding="utf-8")).hexdigest()[:16]


def copy_missing_styles(
    target_members: dict[str, bytes], template_members: dict[str, bytes], style_ids: set[str]
) -> tuple[list[str], dict[str, str], dict[str, dict[str, str | None]]]:
    if not style_ids or "word/styles.xml" not in target_members or "word/styles.xml" not in template_members:
        return [], {}, {}
    target_styles = ET.fromstring(target_members["word/styles.xml"])
    template_styles = ET.fromstring(template_members["word/styles.xml"])
    target_by_id = {
        style.get(w_attr("styleId")): style
        for style in target_styles.findall("w:style", NS)
        if style.get(w_attr("styleId"))
    }
    template_by_id = {
        style.get(w_attr("styleId")): style
        for style in template_styles.findall("w:style", NS)
        if style.get(w_attr("styleId"))
    }
    copied: list[str] = []
    id_map: dict[str, str] = {}
    collision_report: dict[str, dict[str, str | None]] = {}

    def unique_style_id(base_id: str) -> str:
        candidate = f"tplTbl_{base_id}"
        if candidate not in target_by_id:
            return candidate
        suffix = 2
        while f"{candidate}_{suffix}" in target_by_id:
            suffix += 1
        return f"{candidate}_{suffix}"

    def ensure_template_table_style(style_id: str) -> str | None:
        if style_id in id_map:
            return id_map[style_id]
        template_style = template_by_id.get(style_id)
        if template_style is None or template_style.get(w_attr("type")) != "table":
            return None
        target_style = target_by_id.get(style_id)
        template_hash = style_hash(template_style)
        target_hash = style_hash(target_style)
        target_type = target_style.get(w_attr("type")) if target_style is not None else None
        if target_style is not None and target_type == "table" and target_hash == template_hash:
            id_map[style_id] = style_id
            return style_id
        new_id = style_id if target_style is None else unique_style_id(style_id)
        style = deepcopy(template_style)
        style.set(w_attr("styleId"), new_id)
        based_on = style.find("w:basedOn", NS)
        if based_on is not None:
            base_id = based_on.get(w_attr("val"))
            if base_id:
                mapped_base = ensure_template_table_style(base_id)
                if mapped_base:
                    based_on.set(w_attr("val"), mapped_base)
                else:
                    style.remove(based_on)
        target_styles.append(style)
        target_by_id[new_id] = style
        copied.append(new_id)
        id_map[style_id] = new_id
        if target_style is not None and new_id != style_id:
            collision_report[style_id] = {
                "template_hash": template_hash,
                "target_hash": target_hash,
                "target_type": target_type,
                "mapped_style_id": new_id,
            }
        return new_id

    for style_id in sorted(style_ids):
        ensure_template_table_style(style_id)

    target_members["word/styles.xml"] = ET.tostring(
        target_styles, encoding="utf-8", xml_declaration=True, standalone=False
    )
    return copied, id_map, collision_report


def normalize_table_reference_text(text: str) -> str:
    return re.sub(r"表(\s*)(\d+)[.](\d+)", r"表\1\2-\3", text)


def normalize_table_references(root: ET._Element) -> int:
    changed = 0
    for text_node in root.xpath(".//w:t", namespaces=NS):
        value = text_node.text or ""
        normalized = normalize_table_reference_text(value)
        if normalized != value:
            text_node.text = normalized
            changed += 1
    return changed


def write_docx_xml(path: Path, members: dict[str, bytes], root: ET._Element) -> None:
    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True, standalone=False)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)


def parse_donor_map(value: str | None) -> dict[int, int]:
    if not value:
        return {}
    result: dict[int, int] = {}
    for pair in value.split(","):
        if not pair.strip():
            continue
        left, right = pair.split(":", 1)
        result[int(left)] = int(right)
    return result


def body_table_texts(contexts: list[dict[str, object]]) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for index, context in enumerate(contexts, 1):
        tbl = context["tbl"]  # type: ignore[assignment]
        rows = []
        for row in children(tbl, W + "tr"):  # type: ignore[arg-type]
            rows.append([text_of(cell).strip() for cell in children(row, W + "tc")])
        result.append(
            {
                "table_index": index,
                "title": text_of(context["title_p"]).strip() if context.get("title_p") is not None else "",
                "header": list(table_header(tbl)),  # type: ignore[arg-type]
                "rows": rows,
            }
        )
    return result


def copy_body_table_from_source(
    docx_path: Path,
    source_docx_path: Path,
    output_path: Path,
    table_index: int,
) -> dict[str, object]:
    """Replace one body table with the same-index body table from a locked source.

    This intentionally copies only the selected w:tbl node in word/document.xml.
    It is used when a later format pass corrupts table content while the locked
    source/review copy still owns the correct semantic table rows.
    """

    if table_index < 1:
        raise ValueError("--table-index must be >= 1")
    shutil.copy2(docx_path, output_path)
    target_members, target_root = load_docx_xml(output_path)
    source_members, source_root = load_docx_xml(source_docx_path)
    target_contexts = find_table_contexts(target_root, body_only=True)
    source_contexts = find_table_contexts(source_root, body_only=True)
    if len(target_contexts) < table_index:
        raise ValueError(f"Target body table index {table_index} not found")
    if len(source_contexts) < table_index:
        raise ValueError(f"Source body table index {table_index} not found")

    target_context = target_contexts[table_index - 1]
    source_context = source_contexts[table_index - 1]
    target_tbl = target_context["tbl"]  # type: ignore[assignment]
    source_tbl = source_context["tbl"]  # type: ignore[assignment]
    target_parent = target_tbl.getparent()
    if target_parent is None:
        raise ValueError("Target table has no parent")
    target_position = target_parent.index(target_tbl)

    before_target = body_table_texts(target_contexts)
    before_source = body_table_texts(source_contexts)
    target_parent.remove(target_tbl)
    target_parent.insert(target_position, deepcopy(source_tbl))
    after_contexts = find_table_contexts(target_root, body_only=True)
    after_target = body_table_texts(after_contexts)

    style_ids = collect_style_ids(source_tbl)
    copied_styles, style_id_map, style_collisions = copy_missing_styles(target_members, source_members, style_ids)
    if style_id_map:
        copied_tbl = after_contexts[table_index - 1]["tbl"]  # type: ignore[index]
        for elem in copied_tbl.iter():  # type: ignore[union-attr]
            if elem.tag in {W + "pStyle", W + "rStyle", W + "tblStyle"}:
                style_id = elem.get(w_attr("val"))
                mapped = style_id_map.get(style_id or "")
                if mapped:
                    elem.set(w_attr("val"), mapped)

    write_docx_xml(output_path, target_members, target_root)
    return {
        "source_docx": str(source_docx_path),
        "target_docx": str(docx_path),
        "output_docx": str(output_path),
        "body_table_index": table_index,
        "replacement_node_index": target_context.get("node_index"),
        "source_title": before_source[table_index - 1]["title"],
        "target_title_before": before_target[table_index - 1]["title"],
        "target_title_after": after_target[table_index - 1]["title"],
        "source_header": before_source[table_index - 1]["header"],
        "target_header_before": before_target[table_index - 1]["header"],
        "target_header_after": after_target[table_index - 1]["header"],
        "source_rows": before_source[table_index - 1]["rows"],
        "target_rows_before": before_target[table_index - 1]["rows"],
        "target_rows_after": after_target[table_index - 1]["rows"],
        "body_table_count_before": len(target_contexts),
        "body_table_count_after": len(after_contexts),
        "copied_style_ids": copied_styles,
        "style_id_map": style_id_map,
        "style_collisions": style_collisions,
    }


def apply_template_table_styles(
    docx_path: Path,
    template_path: Path,
    output_path: Path,
    donor_map: dict[int, int] | None = None,
    skip_indexes: set[int] | None = None,
    lift_cell_captions: bool = False,
) -> dict[str, object]:
    donor_map = donor_map or {}
    skip_indexes = skip_indexes or set()
    shutil.copy2(docx_path, output_path)
    target_members, target_root = load_docx_xml(output_path)
    template_members, template_root = load_docx_xml(template_path)
    template_contexts = find_table_contexts(template_root, body_only=False)
    donor_title_modes_by_header: dict[tuple[str, ...], str] = {}
    donor_by_header: dict[tuple[str, ...], int] = {}
    for index, context in enumerate(template_contexts, 1):
        header = table_header(context["tbl"])  # type: ignore[arg-type]
        donor_by_header.setdefault(header, index)
        donor_title_modes_by_header.setdefault(header, str(context.get("title_mode") or "none"))
    lifted_cell_captions = (
        lift_first_cell_table_titles(
            target_root,
            body_only=True,
            donor_title_modes_by_header=donor_title_modes_by_header,
        )
        if lift_cell_captions
        else []
    )
    target_contexts = find_table_contexts(target_root, body_only=True)

    before: dict[str, object] = {}
    after: dict[str, object] = {}
    applied: list[dict[str, object]] = []
    applied_table_nodes: list[ET._Element] = []
    structure_issues: list[str] = []
    required_style_ids: set[str] = set()
    for target_index, context in enumerate(target_contexts, 1):
        target_tbl = context["tbl"]  # type: ignore[assignment]
        target_title = context["title_p"]  # type: ignore[assignment]
        target_title_mode = str(context.get("title_mode") or "none")
        before[str(target_index)] = table_metrics(target_tbl, target_title)  # type: ignore[arg-type]
        for issue in context.get("structure_issues", []):  # type: ignore[union-attr]
            structure_issues.append(str(issue))
        if target_index in skip_indexes:
            continue
        donor_index = donor_map.get(target_index)
        if donor_index is None:
            donor_index = donor_by_header.get(table_header(target_tbl))
        if donor_index is None:
            continue
        if target_title is None:
            structure_issues.append(
                f"target table {target_index} lacks an immediately preceding standalone table title paragraph"
            )
            continue
        donor_context = template_contexts[donor_index - 1]
        donor_tbl = donor_context["tbl"]  # type: ignore[assignment]
        donor_title = donor_context["title_p"]  # type: ignore[assignment]
        donor_title_mode = str(donor_context.get("title_mode") or "none")
        if len(table_header(target_tbl)) != len(table_header(donor_tbl)):  # type: ignore[arg-type]
            continue
        if target_title is not None and donor_title is not None and donor_title_mode == target_title_mode:
            set_paragraph_text_from_donor(
                target_title,  # type: ignore[arg-type]
                donor_title,  # type: ignore[arg-type]
                normalize_table_caption_number(text_of(target_title)),  # type: ignore[arg-type]
            )
        elif target_title is not None and donor_title is not None and donor_title_mode != target_title_mode:
            structure_issues.append(
                f"target table {target_index} title mode `{target_title_mode}` does not match donor title mode `{donor_title_mode}`"
            )
            continue
        donor_tbl_style = child(child(donor_tbl, W + "tblPr"), W + "tblStyle")  # type: ignore[arg-type]
        if donor_tbl_style is not None and donor_tbl_style.get(w_attr("val")):
            required_style_ids.add(donor_tbl_style.get(w_attr("val")))
        apply_table_from_donor(target_tbl, donor_tbl)  # type: ignore[arg-type]
        applied_table_nodes.append(target_tbl)  # type: ignore[arg-type]
        applied.append(
            {
                "target_table_index": target_index,
                "donor_table_index": donor_index,
                "target_header": list(table_header(target_tbl)),
                "donor_header": list(table_header(donor_tbl)),  # type: ignore[arg-type]
                "target_title_after": text_of(target_title).strip() if target_title is not None else "",
                "donor_title": text_of(donor_title).strip() if donor_title is not None else "",
            }
        )

    copied_styles, style_id_map, style_collisions = copy_missing_styles(target_members, template_members, required_style_ids)
    for target_tbl in applied_table_nodes:
        tbl_style = child(child(target_tbl, W + "tblPr"), W + "tblStyle")
        if tbl_style is None:
            continue
        style_id = tbl_style.get(w_attr("val"))
        mapped_style_id = style_id_map.get(style_id or "")
        if mapped_style_id and mapped_style_id != style_id:
            tbl_style.set(w_attr("val"), mapped_style_id)

    normalized_reference_text_nodes = normalize_table_references(target_root)

    for target_index, context in enumerate(target_contexts, 1):
        after[str(target_index)] = table_metrics(context["tbl"], context["title_p"])  # type: ignore[arg-type]

    write_docx_xml(output_path, target_members, target_root)
    return {
        "source_docx": str(docx_path),
        "template_docx": str(template_path),
        "output_docx": str(output_path),
        "skip_indexes": sorted(skip_indexes),
        "applied": applied,
        "required_style_ids": sorted(required_style_ids),
        "copied_template_styles": copied_styles,
        "template_style_id_map": style_id_map,
        "template_style_collisions": style_collisions,
        "lifted_cell_captions": lifted_cell_captions,
        "normalized_table_reference_text_nodes": normalized_reference_text_nodes,
        "structure_issues": structure_issues,
        "before": before,
        "after": after,
    }


def patch_table(docx_path: Path, table_index: int, family: str, output_path: Path | None = None) -> Path:
    if family != "wps_second_three_line_rendered":
        raise ValueError(f"Unsupported family: {family}")
    out_path = output_path or docx_path
    if output_path:
        shutil.copy2(docx_path, out_path)

    members, root = load_docx_xml(out_path)
    tables = root.findall(".//w:tbl", NS)
    if len(tables) < table_index:
        raise ValueError(f"Table index {table_index} not found")
    tbl = tables[table_index - 1]

    tbl_pr = ensure(tbl, W + "tblPr")
    clear_shading(tbl_pr)
    tbl_style = tbl_pr.find(W + "tblStyle")
    if tbl_style is not None:
        tbl_pr.remove(tbl_style)

    tbl_borders = ensure_ordered(tbl_pr, W + "tblBorders", TBL_PR_ORDER)
    tbl_borders.clear()
    tbl_borders.extend(
        [
            border(W + "top", "single", "12"),
            border(W + "left", "none", "0"),
            border(W + "bottom", "single", "12"),
            border(W + "right", "none", "0"),
            border(W + "insideH", "none", "0"),
            border(W + "insideV", "none", "0"),
        ]
    )

    rows = tbl.findall(W + "tr")
    if rows:
        header_cells = rows[0].findall(W + "tc")
        for cell in header_cells:
            tc_pr = ensure(cell, W + "tcPr")
            clear_shading(tc_pr)
            tc_borders = ensure_ordered(tc_pr, W + "tcBorders", TC_PR_ORDER)
            tc_borders.clear()
            tc_borders.extend(
                [
                    border(W + "top", "single", "12"),
                    border(W + "bottom", "single", "6"),
                ]
            )

    for row in rows:
        for cell in row.findall(W + "tc"):
            tc_pr = ensure(cell, W + "tcPr")
            clear_shading(tc_pr)
            for para in cell.findall(f".//{W}p"):
                ppr = para.find(W + "pPr")
                if ppr is not None:
                    clear_shading(ppr)
                for run in para.findall(f"{W}r"):
                    rpr = run.find(W + "rPr")
                    if rpr is not None:
                        clear_shading(rpr)

    write_docx_xml(out_path, members, root)
    return out_path


def patch_table_body_numeric_cells(
    docx_path: Path,
    table_index: int,
    output_path: Path | None = None,
    *,
    normalize_borders: bool = False,
) -> dict[str, object]:
    out_path = output_path or docx_path
    if output_path:
        shutil.copy2(docx_path, out_path)
    members, root = load_docx_xml(out_path)
    body_tables = find_table_contexts(root, body_only=True)
    if len(body_tables) < table_index:
        raise ValueError(f"Body table index {table_index} not found")
    tbl = body_tables[table_index - 1]["tbl"]  # type: ignore[index]
    result = normalize_table_body_numeric_cells(tbl)  # type: ignore[arg-type]
    border_result = normalize_three_line_table_borders(tbl) if normalize_borders else {}
    write_docx_xml(out_path, members, root)
    return {
        "source_docx": str(docx_path),
        "output_docx": str(out_path),
        "table_index": table_index,
        "border_normalization": border_result,
        **result,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply locked thesis table-family formatting to DOCX tables.")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--table-index", type=int, default=1)
    parser.add_argument("--family")
    parser.add_argument("--output")
    parser.add_argument("--template")
    parser.add_argument("--apply-template-styles", action="store_true")
    parser.add_argument("--donor-map", help="Comma-separated target:donor table mapping, for example 2:2,3:3,4:13")
    parser.add_argument("--skip-indexes", default="", help="Comma-separated target table indexes to leave unchanged.")
    parser.add_argument(
        "--lift-cell-captions",
        action="store_true",
        help="Move table titles trapped in the first header cell into standalone paragraphs before applying styles.",
    )
    parser.add_argument("--audit-json")
    parser.add_argument(
        "--normalize-body-numeric-cells",
        action="store_true",
        help="Normalize body numeric cells in one body table to the locked thesis metric-cell font/alignment baseline.",
    )
    parser.add_argument(
        "--normalize-body-table-borders",
        action="store_true",
        help="Normalize the selected body table to the canonical three-line table border family.",
    )
    parser.add_argument(
        "--normalize-all-body-tables",
        action="store_true",
        help="Normalize every body table to the locked thesis three-line table family plus cell fonts, header repeat, and no row splitting.",
    )
    parser.add_argument(
        "--copy-body-table-from-source",
        help="Copy the selected same-index body table from this locked source DOCX before later table normalization.",
    )
    args = parser.parse_args()

    if args.copy_body_table_from_source:
        if not args.output:
            raise SystemExit("--output is required with --copy-body-table-from-source")
        audit = copy_body_table_from_source(
            docx_path=Path(args.docx),
            source_docx_path=Path(args.copy_body_table_from_source),
            output_path=Path(args.output),
            table_index=args.table_index,
        )
        if args.audit_json:
            Path(args.audit_json).parent.mkdir(parents=True, exist_ok=True)
            Path(args.audit_json).write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"patched_docx={args.output}")
        print(f"copied_body_table_index={args.table_index}")
        return 0

    if args.normalize_all_body_tables:
        output_path = Path(args.output) if args.output else Path(args.docx)
        audit = patch_all_body_tables(
            docx_path=Path(args.docx),
            output_path=output_path,
        )
        if args.audit_json:
            Path(args.audit_json).parent.mkdir(parents=True, exist_ok=True)
            Path(args.audit_json).write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"patched_docx={output_path}")
        print(f"patched_body_tables={audit['table_count']}")
        return 0

    if args.normalize_body_numeric_cells:
        output_path = Path(args.output) if args.output else Path(args.docx)
        audit = patch_table_body_numeric_cells(
            docx_path=Path(args.docx),
            table_index=args.table_index,
            output_path=output_path,
            normalize_borders=args.normalize_body_table_borders,
        )
        if args.audit_json:
            Path(args.audit_json).parent.mkdir(parents=True, exist_ok=True)
            Path(args.audit_json).write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"patched_docx={output_path}")
        print(f"changed_numeric_cells={audit['changed_cell_count']}")
        return 0

    if args.apply_template_styles:
        if not args.template:
            raise SystemExit("--template is required with --apply-template-styles")
        if not args.output:
            raise SystemExit("--output is required with --apply-template-styles")
        skip_indexes = {int(item) for item in args.skip_indexes.split(",") if item.strip()}
        audit = apply_template_table_styles(
            docx_path=Path(args.docx),
            template_path=Path(args.template),
            output_path=Path(args.output),
            donor_map=parse_donor_map(args.donor_map),
            skip_indexes=skip_indexes,
            lift_cell_captions=args.lift_cell_captions,
        )
        if args.audit_json:
            Path(args.audit_json).parent.mkdir(parents=True, exist_ok=True)
            Path(args.audit_json).write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
        if audit.get("structure_issues"):
            for issue in audit["structure_issues"]:
                print(f"table structure issue: {issue}")
            return 2
        print(f"patched_docx={args.output}")
        print(f"applied_tables={len(audit['applied'])}")
        return 0

    if not args.family:
        raise SystemExit("--family is required unless --apply-template-styles is used")
    out_path = patch_table(
        docx_path=Path(args.docx),
        table_index=args.table_index,
        family=args.family,
        output_path=Path(args.output) if args.output else None,
    )
    print(f"patched_docx={out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
