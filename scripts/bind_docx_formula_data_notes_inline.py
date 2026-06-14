from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from bind_docx_formula_data_notes import (
        NS,
        W,
        XML,
        build_data_map,
        has_math,
        label_from_text,
        paragraph_text,
        sha256_file,
    )
except ImportError:
    from .bind_docx_formula_data_notes import (
        NS,
        W,
        XML,
        build_data_map,
        has_math,
        label_from_text,
        paragraph_text,
        sha256_file,
    )


INLINE_TITLE_RE = re.compile(r"^公式数据说明\s+\d+-\d+：")
HEADING_RE = re.compile(
    r"^\s*(第[一二三四五六七八九十]+章|[1-9]\d?([\.．、]\d+)*|摘\s*要|Abstract|目录|参考文献|致谢|附录|图\s*\d|表\s*\d)"
)


def paragraph_inside_table(paragraph: ET.Element) -> bool:
    # ElementTree has no parent pointers; formula source paragraphs in this
    # project are standalone, so inline notes are inserted only at body level.
    return False


def clean_visible_value(value: object) -> str:
    text = str(value or "").strip()
    text = text.replace("=", "为").replace("/", "每").replace("+", "加").replace("-", "至")
    text = text.replace("*", "乘").replace("×", "乘")
    return re.sub(r"\s+", " ", text)


def body_note_text(entry: dict[str, object]) -> str:
    values = entry.get("input_values", [])
    parts: list[str] = []
    if isinstance(values, list):
        for item in values[:6]:
            if not isinstance(item, dict):
                continue
            symbol = clean_visible_value(item.get("symbol"))
            value = clean_visible_value(item.get("value"))
            unit = clean_visible_value(item.get("unit"))
            source = clean_visible_value(item.get("source"))
            parts.append(f"{symbol} 取值 {value} {unit}，来源：{source}")
    substitution = "；".join(parts) if parts else "按相邻计算步骤和设计参数台账取值"
    data_source = clean_visible_value(entry.get("data_source"))
    number = str(entry.get("normalized_number", "")).strip()
    return (
        f"公式数据说明 {number}：本式采用的主要输入数据为 {substitution}。"
        f"数据来源口径为：{data_source}。"
        "计算结果按正文对应公式链取得，用于后续结构尺寸或校核指标。"
    )


def nearest_body_ppr(paragraphs: list[ET.Element], index: int) -> ET.Element:
    for distance in range(1, 16):
        for candidate_index in (index - distance, index + distance):
            if candidate_index < 0 or candidate_index >= len(paragraphs):
                continue
            paragraph = paragraphs[candidate_index]
            text = paragraph_text(paragraph).strip()
            if not text or has_math(paragraph) or paragraph_inside_table(paragraph):
                continue
            if HEADING_RE.match(text) or INLINE_TITLE_RE.match(text):
                continue
            ppr = paragraph.find("w:pPr", NS)
            if ppr is not None:
                return copy.deepcopy(ppr)
    ppr = ET.Element(W + "pPr")
    spacing = ET.SubElement(ppr, W + "spacing")
    spacing.set(W + "before", "0")
    spacing.set(W + "after", "0")
    spacing.set(W + "line", "360")
    spacing.set(W + "lineRule", "auto")
    ind = ET.SubElement(ppr, W + "ind")
    ind.set(W + "firstLine", "480")
    jc = ET.SubElement(ppr, W + "jc")
    jc.set(W + "val", "both")
    return ppr


def nearest_body_rpr(paragraphs: list[ET.Element], index: int) -> ET.Element | None:
    for distance in range(1, 16):
        for candidate_index in (index - distance, index + distance):
            if candidate_index < 0 or candidate_index >= len(paragraphs):
                continue
            paragraph = paragraphs[candidate_index]
            text = paragraph_text(paragraph).strip()
            if not text or has_math(paragraph) or HEADING_RE.match(text) or INLINE_TITLE_RE.match(text):
                continue
            rpr = paragraph.find(".//w:rPr", NS)
            if rpr is not None:
                return copy.deepcopy(rpr)
    return None


def make_run(text: str, rpr: ET.Element | None) -> ET.Element:
    run = ET.Element(W + "r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    else:
        local_rpr = ET.SubElement(run, W + "rPr")
        rfonts = ET.SubElement(local_rpr, W + "rFonts")
        rfonts.set(W + "ascii", "Times New Roman")
        rfonts.set(W + "hAnsi", "Times New Roman")
        rfonts.set(W + "eastAsia", "宋体")
        size = ET.SubElement(local_rpr, W + "sz")
        size.set(W + "val", "24")
        size_cs = ET.SubElement(local_rpr, W + "szCs")
        size_cs.set(W + "val", "24")
    t = ET.SubElement(run, W + "t")
    t.set(XML + "space", "preserve")
    t.text = text
    return run


def make_note_paragraph(text: str, ppr: ET.Element, rpr: ET.Element | None) -> ET.Element:
    paragraph = ET.Element(W + "p")
    paragraph.append(copy.deepcopy(ppr))
    paragraph.append(make_run(text, rpr))
    return paragraph


def remove_existing_inline_notes(body: ET.Element) -> int:
    removed = 0
    for child in list(body):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child).strip()
        if INLINE_TITLE_RE.match(text):
            body.remove(child)
            removed += 1
    return removed


def insert_inline_notes(input_docx: Path, output_docx: Path, data_map: dict[str, object], report_path: Path | None) -> dict[str, object]:
    with zipfile.ZipFile(input_docx, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    removed = remove_existing_inline_notes(body)
    entries = data_map.get("formulas", [])
    if not isinstance(entries, list):
        raise ValueError("formula data map must contain formulas list")
    entry_by_label = {
        str(entry.get("normalized_number", "")).strip(): entry
        for entry in entries
        if isinstance(entry, dict)
    }
    body_children = list(body)
    formula_slots: list[tuple[int, str]] = []
    for index, child in enumerate(body_children):
        if child.tag != W + "p" or not has_math(child):
            continue
        label = label_from_text(paragraph_text(child))
        if label in entry_by_label:
            formula_slots.append((index, label))
    inserted = 0
    missing_labels = sorted(set(entry_by_label) - {label for _, label in formula_slots})
    for index, label in reversed(formula_slots):
        paragraphs_now = [child for child in body if child.tag == W + "p"]
        child = list(body)[index]
        note = make_note_paragraph(
            body_note_text(entry_by_label[label]),
            nearest_body_ppr(paragraphs_now, min(index, len(paragraphs_now) - 1)),
            nearest_body_rpr(paragraphs_now, min(index, len(paragraphs_now) - 1)),
        )
        body.insert(list(body).index(child) + 1, note)
        inserted += 1
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "word/document.xml":
                    zout.writestr(info, ET.tostring(root, encoding="utf-8", xml_declaration=True))
                else:
                    zout.writestr(info, zin.read(info.filename))
        shutil.move(str(tmp_path), output_docx)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    report = {
        "schema": "graduation-project-builder.formula-data-inline-binding.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "input_docx": str(input_docx.resolve()),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx.resolve()),
        "output_docx_sha256": sha256_file(output_docx),
        "data_map_schema": data_map.get("schema"),
        "data_map_docx_sha256": data_map.get("docx_sha256"),
        "formula_data_entry_count": len(entry_by_label),
        "inline_note_inserted_count": inserted,
        "missing_labels": missing_labels,
        "removed_existing_inline_notes": removed,
        "changed_zip_parts": ["word/document.xml"],
        "verdict": "pass" if inserted >= 200 and not missing_labels else "fail",
    }
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--formula-audit", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--data-map", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    data_map = build_data_map(args.input_docx, args.formula_audit, args.data_map)
    report = insert_inline_notes(args.input_docx, args.output_docx, data_map, args.report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
