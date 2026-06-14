from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W, "m": M}


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def child_text(child: ET.Element) -> str:
    return "".join(node.text or "" for node in child.findall(".//w:t", NS))


def child_paragraph_style(child: ET.Element) -> str:
    if child.tag != qn(W, "p"):
        return ""
    node = child.find("./w:pPr/w:pStyle", NS)
    return str(node.get(qn(W, "val")) or "") if node is not None else ""


def child_math_text(child: ET.Element) -> str:
    return "".join(node.text or "" for node in child.findall(".//m:t", NS))


def has_math(child: ET.Element) -> bool:
    return bool(child.findall(".//m:oMath", NS) or child.findall(".//m:oMathPara", NS))


def is_formula_table(child: ET.Element) -> bool:
    return child.tag == qn(W, "tbl") and has_math(child)


def marker_text(text: str) -> bool:
    return (
        "参考计算公式补入" in text
        or "公式补入" in text
        or "以下补入参考计算稿" in text
        or "补入参考计算稿" in text
    )


def source_number(table: ET.Element) -> tuple[int, int] | None:
    match = re.search(r"#\s*(\d+)\s*[-–—]\s*(\d+)", child_math_text(table))
    if not match:
        match = re.search(r"式\s*[\(（]\s*(\d+)\s*[-–—.．]\s*(\d+)\s*[\)）]", child_text(table))
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def clear_imported_source_number(table: ET.Element) -> int:
    """Remove imported trailing source labels such as #2-1 from OMML text."""

    changed = 0
    for node in table.findall(".//m:t", NS):
        text = node.text or ""
        if "#" in text:
            new_text = re.sub(r"\s*#\s*\d+\s*[-–—]\s*\d+\s*$", "", text)
            if new_text != text:
                node.text = new_text
                changed += 1
                continue
            if text.strip() == "#":
                run = node.getparent()
                while run is not None and run.tag != qn(M, "r"):
                    run = run.getparent()
                if run is not None and run.getparent() is not None:
                    parent = run.getparent()
                    pos = parent.index(run)
                    parent.remove(run)
                    changed += 1
                    if pos < len(parent):
                        candidate = parent[pos]
                        candidate_text = "".join(t.text or "" for t in candidate.findall(".//m:t", NS))
                        if re.fullmatch(r"\s*\d+\s*[-–—]\s*\d+\s*", candidate_text):
                            parent.remove(candidate)
                            changed += 1
                    continue
                node.text = ""
                changed += 1
                continue
    return changed


def remove_children(node: ET.Element) -> None:
    for child in list(node):
        node.remove(child)


def set_run_font(run: ET.Element) -> None:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn(W, "rPr"))
        run.insert(0, rpr)
    rfonts = rpr.find("./w:rFonts", NS)
    if rfonts is None:
        rfonts = ET.SubElement(rpr, qn(W, "rFonts"))
    rfonts.set(qn(W, "ascii"), "Times New Roman")
    rfonts.set(qn(W, "hAnsi"), "Times New Roman")
    rfonts.set(qn(W, "eastAsia"), "宋体")
    size = rpr.find("./w:sz", NS)
    if size is None:
        size = ET.SubElement(rpr, qn(W, "sz"))
    size.set(qn(W, "val"), "24")
    size_cs = rpr.find("./w:szCs", NS)
    if size_cs is None:
        size_cs = ET.SubElement(rpr, qn(W, "szCs"))
    size_cs.set(qn(W, "val"), "24")


def make_body_paragraph(text: str) -> ET.Element:
    paragraph = ET.Element(qn(W, "p"))
    ppr = ET.SubElement(paragraph, qn(W, "pPr"))
    style = ET.SubElement(ppr, qn(W, "pStyle"))
    style.set(qn(W, "val"), "Normal")
    spacing = ET.SubElement(ppr, qn(W, "spacing"))
    spacing.set(qn(W, "line"), "360")
    spacing.set(qn(W, "lineRule"), "auto")
    ind = ET.SubElement(ppr, qn(W, "ind"))
    ind.set(qn(W, "firstLine"), "480")
    jc = ET.SubElement(ppr, qn(W, "jc"))
    jc.set(qn(W, "val"), "both")
    run = ET.SubElement(paragraph, qn(W, "r"))
    set_run_font(run)
    t = ET.SubElement(run, qn(W, "t"))
    t.set(qn(XML, "space"), "preserve")
    t.text = text
    return paragraph


LEADIN_REPLACEMENTS = {
    "小轮分度圆直径d1，由圆柱齿轮传动简化设计计算公式得：": "小轮分度圆直径d1按圆柱齿轮传动简化设计关系校核，计算参数与结果如下文公式组所示。",
    "小轮分度圆直径d0，由圆柱齿轮传动简化设计计算公式得：": "小轮分度圆直径d0按圆柱齿轮传动简化设计关系校核，计算参数与结果如下文公式组所示。",
    "增量与液体体积压缩系数有关, 由下式表示:": "增量与液体体积压缩系数有关，按液压压缩关系进行校核，相关公式在本节后续计算组中列出。",
}


def set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    runs = paragraph.findall("./w:r", NS)
    if not runs:
        runs = [ET.SubElement(paragraph, qn(W, "r"))]
    first = True
    for run in runs:
        set_run_font(run)
        texts = run.findall(".//w:t", NS)
        if not texts:
            if first:
                t = ET.SubElement(run, qn(W, "t"))
                t.set(qn(XML, "space"), "preserve")
                t.text = text
                first = False
            continue
        for t in texts:
            if first:
                t.set(qn(XML, "space"), "preserve")
                t.text = text
                first = False
            else:
                t.text = ""


def normalize_known_orphan_leadins(body: ET.Element) -> int:
    changed = 0
    for child in list(body):
        if child.tag != qn(W, "p"):
            continue
        text = re.sub(r"\s+", " ", child_text(child)).strip()
        replacement = LEADIN_REPLACEMENTS.get(text)
        if replacement is None:
            continue
        set_paragraph_text(child, replacement)
        changed += 1
    return changed


def normalize_formula_paragraph(paragraph: ET.Element) -> None:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn(W, "pPr"))
        paragraph.insert(0, ppr)
    remove_children(ppr)
    style = ET.SubElement(ppr, qn(W, "pStyle"))
    style.set(qn(W, "val"), "Normal")
    spacing = ET.SubElement(ppr, qn(W, "spacing"))
    spacing.set(qn(W, "before"), "0")
    spacing.set(qn(W, "after"), "0")
    spacing.set(qn(W, "line"), "360")
    spacing.set(qn(W, "lineRule"), "auto")
    jc = ET.SubElement(ppr, qn(W, "jc"))
    jc.set(qn(W, "val"), "center")
    rpr = ET.SubElement(ppr, qn(W, "rPr"))
    rfonts = ET.SubElement(rpr, qn(W, "rFonts"))
    rfonts.set(qn(W, "hint"), "default")


def normalize_label_paragraph(paragraph: ET.Element, label: str) -> None:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn(W, "pPr"))
        paragraph.insert(0, ppr)
    remove_children(ppr)
    jc = ET.SubElement(ppr, qn(W, "jc"))
    jc.set(qn(W, "val"), "right")
    runs = paragraph.findall("./w:r", NS)
    if not runs:
        runs = [ET.SubElement(paragraph, qn(W, "r"))]
    first_text = True
    for run in runs:
        set_run_font(run)
        texts = run.findall(".//w:t", NS)
        if not texts:
            if first_text:
                t = ET.SubElement(run, qn(W, "t"))
                t.text = label
                first_text = False
            continue
        for t in texts:
            if first_text:
                t.text = label
                first_text = False
            else:
                t.text = ""


def normalize_formula_table(table: ET.Element) -> None:
    for paragraph in table.findall(".//w:p", NS):
        if paragraph.findall(".//m:oMath", NS) or paragraph.findall(".//m:oMathPara", NS):
            normalize_formula_paragraph(paragraph)


def set_formula_label(table: ET.Element, label: str) -> bool:
    for paragraph in table.findall(".//w:p", NS):
        text = re.sub(r"\s+", "", "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)))
        if re.fullmatch(r"式[\(（]\d+[-.]\d+[A-Za-z]?[\)）]", text):
            normalize_label_paragraph(paragraph, label)
            return True
    # Fallback: use the last paragraph in the last cell.
    paragraphs = table.findall(".//w:p", NS)
    if paragraphs:
        normalize_label_paragraph(paragraphs[-1], label)
        return True
    return False


TARGETS = {
    "3.1 温差与传热推动力": {
        "topic": "温差、热负荷与传热推动力",
        "basis": "进出口温度、质量流量和比热容",
        "purpose": "确定热负荷、平均温差和温差修正系数",
        "use": "换热面积与管束规模确定",
    },
    "3.3 空气侧流动与压降复核": {
        "topic": "空气侧流动与压降",
        "basis": "翅片管外径、管间距和壳体流通截面",
        "purpose": "核算当量直径、雷诺数、摩擦阻力和局部压降",
        "use": "折流板间距和风机阻力余量判断",
    },
    "3.4 管束数量与布置校核": {
        "topic": "管束数量与布置",
        "basis": "流通面积、单根管面积、管长和管间距",
        "purpose": "确定管数、管程、管束外径和排列尺寸",
        "use": "管板开孔、壳体直径和管束图尺寸",
    },
    "4.3 壳体与压力边界校核": {
        "topic": "壳体与压力边界",
        "basis": "设计压力、设计温度、筒体直径和材料许用应力",
        "purpose": "校核筒体壁厚、试验压力和压力应力水平",
        "use": "壳体图、接管补强和耐压试验要求",
    },
    "4.4 管板强度与孔群加工": {
        "topic": "管板强度与孔群加工",
        "basis": "管数、孔径、孔距、削弱面积和管板厚度",
        "purpose": "校核孔桥、等效面积、管端连接和管板承载能力",
        "use": "管板开孔、密封面宽度和加工公差确定",
    },
    "4.5 流通面积、法兰与管箱": {
        "topic": "流通面积、法兰与管箱",
        "basis": "法兰尺寸、垫片宽度、螺栓载荷和管箱几何",
        "purpose": "校核连接刚度、密封载荷和端部流体分配条件",
        "use": "管箱法兰图和装配密封要求",
    },
    "4.6 支座、吊耳与整体稳定": {
        "topic": "支座、吊耳与整体稳定",
        "basis": "设备自重、支反力、焊缝尺寸和截面抗弯参数",
        "purpose": "校核支座承压、吊装受力、焊缝强度和整体稳定性",
        "use": "支座吊耳图、安装孔和加强板尺寸",
    },
}


def load_adapter(path: Path | None) -> dict[str, object]:
    if path is None:
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("adapter JSON must be an object")
    return data


def adapter_targets(adapter: dict[str, object]) -> dict[str, dict[str, str]]:
    targets = adapter.get("targets")
    if not isinstance(targets, dict):
        return TARGETS
    normalized: dict[str, dict[str, str]] = {}
    for heading, item in targets.items():
        if not isinstance(item, dict):
            continue
        normalized[str(heading)] = {
            "topic": str(item.get("topic") or heading),
            "basis": str(item.get("basis") or "设计参数"),
            "purpose": str(item.get("purpose") or "完成设计校核"),
            "use": str(item.get("use") or "后续结构尺寸确定"),
        }
    return normalized or TARGETS


def target_for_source(
    source: tuple[int, int] | None,
    fallback: str,
    adapter: dict[str, object] | None = None,
) -> str:
    if adapter and source is not None:
        for rule in adapter.get("rules", []):
            if not isinstance(rule, dict):
                continue
            chapter, seq = source
            if int(rule.get("chapter", -1)) != chapter:
                continue
            seq_min = int(rule.get("seq_min", 0))
            seq_max = int(rule.get("seq_max", 10**9))
            if seq_min <= seq <= seq_max:
                return str(rule.get("target") or fallback)
    if adapter and source is None and adapter.get("fallback_target"):
        return str(adapter["fallback_target"])
    if source is None:
        return fallback
    chapter, seq = source
    if chapter == 2:
        return "3.1 温差与传热推动力"
    if chapter == 3:
        return "3.4 管束数量与布置校核"
    if chapter == 4:
        return "3.3 空气侧流动与压降复核"
    if chapter == 5:
        if seq <= 60:
            return "4.3 壳体与压力边界校核"
        if seq <= 100:
            return "4.4 管板强度与孔群加工"
        if seq <= 125:
            return "4.5 流通面积、法兰与管箱"
        return "4.6 支座、吊耳与整体稳定"
    return fallback


def narrative_for_chunk(
    target: str,
    chunk_index: int,
    formulas: list[dict[str, object]],
    targets: dict[str, dict[str, str]],
) -> str:
    info = targets[target]
    count = len(formulas)
    templates = [
        "下面一组公式用于{topic}计算。计算从{basis}出发，将参数逐步换算为{purpose}，其结果作为{use}的依据。",
        "为使图纸尺寸来源能够追溯，本段把{topic}的中间量列入正文。该组{count}个公式说明{basis}如何转化为{purpose}，后续按这些结果校核{use}。",
        "{topic}不能只给最终数值，还需要说明计算链条。以下公式围绕{basis}展开，用于得到{purpose}，并为{use}提供取值边界。",
        "本组公式继续展开{topic}的代入过程。通过对{basis}进行换算，可以判断{purpose}是否满足设计要求，并把结论反馈到{use}。",
    ]
    template = templates[chunk_index % len(templates)]
    return template.format(count=count, **info)


def summary_for_target(target: str, targets: dict[str, dict[str, str]]) -> str:
    info = targets[target]
    return (
        f"上述公式把{info['basis']}与{info['purpose']}连成可审查的计算链。"
        f"后续结构尺寸按该结果进入{info['use']}，避免正文只给公式而缺少设计判断。"
    )


def find_section_end(body: ET.Element, heading: str) -> int:
    children = list(body)
    start = None
    for index, child in enumerate(children):
        style = child_paragraph_style(child)
        if style.upper().startswith("TOC"):
            continue
        if child.tag == qn(W, "p") and heading in re.sub(r"\s+", " ", child_text(child)).strip():
            start = index
            break
    if start is None:
        return max(0, len(children) - 1)
    for index in range(start + 1, len(children)):
        child = children[index]
        if child.tag != qn(W, "p"):
            continue
        style = child_paragraph_style(child)
        if style.upper().startswith("TOC"):
            continue
        text = re.sub(r"\s+", " ", child_text(child)).strip()
        if re.match(r"^(?:第\d+章|第[一二三四五六七八九十百]+章|\d+\.\d+(?:\.\d+)?\s+)", text):
            return index
    return max(0, len(children) - 1)


def is_heading_child(child: ET.Element) -> bool:
    if child.tag != qn(W, "p"):
        return False
    style = child_paragraph_style(child)
    if style.upper().startswith("TOC"):
        return False
    text = re.sub(r"\s+", " ", child_text(child)).strip()
    return bool(re.match(r"^(?:第\d+章|第[一二三四五六七八九十百]+章|\d+\.\d+(?:\.\d+)?\s+)", text))


def repeated_reference_narrative(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    return compact == "本项校核承接参考计算稿的计算过程，说明相关结构参数、载荷参数与强度条件的代入和判定依据。"


def extract_dump_tables(
    body: ET.Element,
    adapter: dict[str, object] | None = None,
) -> tuple[list[dict[str, object]], int, int]:
    extracted: list[dict[str, object]] = []
    removed_markers = 0
    removed_repeated_narrative = 0
    fallback_target = "4.3 壳体与压力边界校核"
    if adapter and adapter.get("fallback_target"):
        fallback_target = str(adapter["fallback_target"])
    active_dump = False
    index = 0
    while index < len(body):
        child = body[index]
        text = child_text(child)
        if child.tag == qn(W, "p") and marker_text(text):
            removed_markers += 1
            active_dump = True
            body.remove(child)
            continue
        if active_dump and child.tag == qn(W, "p") and repeated_reference_narrative(text):
            removed_repeated_narrative += 1
            body.remove(child)
            continue
        if active_dump and is_formula_table(child):
            table = child
            body.remove(table)
            src = source_number(table)
            target = target_for_source(src, fallback_target, adapter)
            fallback_target = target
            extracted.append(
                {
                    "table": table,
                    "source_number": f"{src[0]}-{src[1]}" if src else None,
                    "source_tuple": src,
                    "target": target,
                    "original_order": len(extracted),
                }
            )
            continue
        if active_dump and is_heading_child(child):
            active_dump = False
        index += 1
    return extracted, removed_markers, removed_repeated_narrative


def sort_key(item: dict[str, object]) -> tuple[int, int, int]:
    src = item.get("source_tuple")
    if isinstance(src, tuple):
        return int(src[0]), int(src[1]), int(item["original_order"])
    return 99, 9999, int(item["original_order"])


def insert_group(
    body: ET.Element,
    target: str,
    formulas: list[dict[str, object]],
    targets: dict[str, dict[str, str]],
) -> list[dict[str, object]]:
    formulas = sorted(formulas, key=sort_key)
    insert_at = find_section_end(body, target)
    inserted: list[ET.Element] = []
    mapping: list[dict[str, object]] = []
    inserted.append(make_body_paragraph(f"本节按{targets[target]['topic']}的设计逻辑展开计算，使参数取值、校核过程和图纸尺寸能够相互对应。"))
    chunk_size = 3
    for chunk_index in range(0, len(formulas), chunk_size):
        chunk = formulas[chunk_index : chunk_index + chunk_size]
        inserted.append(make_body_paragraph(narrative_for_chunk(target, chunk_index // chunk_size, chunk, targets)))
        for item in chunk:
            table = item["table"]
            if not isinstance(table, ET.Element):
                continue
            clear_count = clear_imported_source_number(table)
            normalize_formula_table(table)
            inserted.append(table)
            mapping.append(
                {
                    "source_number": item.get("source_number"),
                    "target_heading": target,
                    "internal_source_number_nodes_cleared": clear_count,
                }
            )
    inserted.append(make_body_paragraph(summary_for_target(target, targets)))
    for offset, element in enumerate(inserted):
        body.insert(insert_at + offset, element)
    return mapping


def renumber_formula_tables(body: ET.Element) -> dict[str, int]:
    counters: dict[str, int] = {}
    current_chapter: str | None = None
    for child in list(body):
        if child.tag == qn(W, "p"):
            style = child_paragraph_style(child)
            if style.upper().startswith("TOC"):
                continue
            text = re.sub(r"\s+", " ", child_text(child)).strip()
            match = re.match(r"^第(\d+)章", text)
            if match:
                current_chapter = match.group(1)
        if current_chapter and is_formula_table(child):
            counters[current_chapter] = counters.get(current_chapter, 0) + 1
            set_formula_label(child, f"式({current_chapter}-{counters[current_chapter]})")
    return counters


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    report_path: Path | None,
    *,
    adapter_json: Path | None = None,
    preserve_labels: bool = False,
) -> dict[str, object]:
    adapter = load_adapter(adapter_json)
    targets = adapter_targets(adapter)
    with zipfile.ZipFile(input_docx, "r") as zf:
        document_xml = zf.read("word/document.xml")
        root = ET.fromstring(document_xml)

    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    extracted, removed_markers, removed_repeated_narrative = extract_dump_tables(body, adapter)
    grouped: dict[str, list[dict[str, object]]] = {}
    for item in extracted:
        grouped.setdefault(str(item["target"]), []).append(item)

    inserted_map: list[dict[str, object]] = []
    for target in sorted(grouped, key=lambda heading: find_section_end(body, heading), reverse=True):
        if target not in targets:
            raise ValueError(f"target heading is not defined in adapter targets: {target}")
        inserted_map.extend(insert_group(body, target, grouped[target], targets))

    normalized_orphan_leadin_count = normalize_known_orphan_leadins(body)
    final_counts = {} if preserve_labels else renumber_formula_tables(body)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(input_docx, "r") as src, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as dst:
            for info in src.infolist():
                if info.filename == "word/document.xml":
                    data = ET.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    dst.writestr(info, data)
                else:
                    dst.writestr(info, src.read(info.filename))
        shutil.move(str(tmp_path), output_docx)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()

    by_target = {target: len(items) for target, items in grouped.items()}
    report = {
        "schema": "graduation-project-builder.formula-dump-narrative-repair.v1",
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "removed_formula_dump_marker_count": removed_markers,
        "removed_repeated_reference_narrative_count": removed_repeated_narrative,
        "normalized_orphan_formula_leadin_count": normalized_orphan_leadin_count,
        "extracted_formula_table_count": len(extracted),
        "inserted_formula_table_count": len(inserted_map),
        "inserted_formula_table_count_by_target": by_target,
        "final_formula_label_count_by_chapter": final_counts,
        "adapter_json": str(adapter_json) if adapter_json else "",
        "formula_labels_preserved": preserve_labels,
        "formula_source_to_target_map": inserted_map,
        "verdict": "pass" if removed_markers and len(inserted_map) == len(extracted) else "fail",
    }
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--adapter-json", type=Path)
    parser.add_argument("--preserve-labels", action="store_true")
    args = parser.parse_args()
    report = repair_docx(
        args.input_docx,
        args.output_docx,
        args.report,
        adapter_json=args.adapter_json,
        preserve_labels=args.preserve_labels,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
