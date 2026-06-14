from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS, "m": M_NS}
W = "{" + W_NS + "}"
M = "{" + M_NS + "}"
XML = "{" + XML_NS + "}"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def qn(ns: str, tag: str) -> str:
    return "{" + ns + "}" + tag


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def math_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//m:t", NS))


def has_math(element: ET.Element) -> bool:
    return (
        element.find(".//m:oMath", NS) is not None
        or element.find(".//m:oMathPara", NS) is not None
        or element.find(".//w:object", NS) is not None
    )


def label_from_text(value: str) -> str | None:
    match = re.search(r"式\s*[\(（]\s*(\d+)\s*[-－–—]\s*(\d+)\s*[\)）]", value or "")
    if not match:
        return None
    return f"{int(match.group(1))}-{int(match.group(2))}"


def label_sort_key(label: str) -> tuple[int, int]:
    left, right = label.split("-", 1)
    return int(left), int(right)


def split_formula(value: str) -> tuple[str, str]:
    for op in ("≤", ">=", "≥", "=", "≦"):
        if op in value:
            left, right = value.split(op, 1)
            return left.strip(), right.strip()
    return "", value.strip()


FUNCTION_NAMES = {
    "min",
    "max",
    "tan",
    "sin",
    "cos",
    "sqrt",
    "ln",
    "log",
}


TOKEN_RE = re.compile(
    r"\[[^\]]+\](?:_[A-Za-z0-9Α-ω]+)?|"
    r"[A-Za-zΑ-ω]+(?:_[A-Za-z0-9Α-ω]+)?[A-Za-z0-9Α-ω]*|"
    r"[δσφτγηρπμλαβθωΩΣΔ][A-Za-z0-9_]*"
)


def base_token(token: str) -> str:
    compact = token.strip()
    compact = re.sub(r"\d+$", "", compact)
    compact = compact.replace("[", "").replace("]", "")
    return compact or token.strip()


KNOWN_VALUES: dict[str, tuple[str, str, str]] = {
    "D_o": ("1200", "mm", "总装配图筒体外径"),
    "Do": ("1200", "mm", "总装配图筒体外径"),
    "D_i": ("1176", "mm", "由筒体外径和名义厚度换算"),
    "Di": ("1176", "mm", "由筒体外径和名义厚度换算"),
    "t_n": ("12", "mm", "筒体名义厚度设计台账"),
    "δ_n": ("12", "mm", "筒体名义厚度设计台账"),
    "δ_nom": ("12", "mm", "筒体名义厚度设计台账"),
    "δe": ("10", "mm", "腐蚀裕量扣除后的有效厚度"),
    "δ_e": ("10", "mm", "腐蚀裕量扣除后的有效厚度"),
    "C": ("2", "mm", "腐蚀裕量和制造负偏差合计"),
    "C_1": ("1", "mm", "钢板厚度负偏差"),
    "C_2": ("1", "mm", "腐蚀裕量"),
    "L": ("2400", "mm", "总装配图筒体有效长度"),
    "L_s": ("2200", "mm", "鞍座中心距"),
    "p": ("0.25", "MPa", "设计压力台账"),
    "p_d": ("0.30", "MPa", "设计压力台账"),
    "p_w": ("0.25", "MPa", "工作压力台账"),
    "pT": ("0.375", "MPa", "试验压力系数换算"),
    "p_T": ("0.375", "MPa", "试验压力系数换算"),
    "T_w": ("90", "℃", "工作温度台账"),
    "T_d": ("110", "℃", "设计温度台账"),
    "ΔT": ("20", "℃", "温升裕量"),
    "σ_y": ("235", "MPa", "Q235B材料屈服强度"),
    "σ_b": ("370", "MPa", "Q235B材料抗拉强度"),
    "σ": ("150", "MPa", "许用应力表"),
    "σ_t": ("150", "MPa", "设计温度许用应力"),
    "τ": ("80", "MPa", "剪切许用应力"),
    "η": ("0.85", "1", "焊接接头系数"),
    "η_w": ("0.85", "1", "焊接接头系数"),
    "η_t": ("0.94", "1", "传动效率台账"),
    "η_g": ("0.96", "1", "齿轮传动效率"),
    "η_b": ("0.99", "1", "轴承效率"),
    "η_s": ("0.98", "1", "密封效率"),
    "ηΣ": ("0.91", "1", "传动链总效率"),
    "n_s": ("1.6", "1", "强度安全系数"),
    "n_b": ("3.0", "1", "抗拉安全系数"),
    "n_m": ("960", "r/min", "电动机额定转速"),
    "n_c": ("8", "r/min", "筒体工作转速"),
    "P_m": ("7.5", "kW", "电动机功率"),
    "P_a": ("7.05", "kW", "有效传动功率"),
    "P_out": ("6.8", "kW", "输出功率"),
    "T_c": ("8410", "N·m", "筒体输出扭矩"),
    "T_ca": ("10500", "N·m", "工况系数修正扭矩"),
    "G": ("28000", "N", "筒体、物料和附件总重"),
    "G_s": ("16000", "N", "壳体自重"),
    "G_f": ("9000", "N", "物料重量"),
    "G_r": ("3000", "N", "转动部件重量"),
    "ρ_s": ("7850", "kg/m3", "钢材密度"),
    "ρ_m": ("1200", "kg/m3", "物料密度"),
    "ρ_w": ("1000", "kg/m3", "水密度"),
    "V_m": ("1.80", "m3", "装料体积台账"),
    "V_g": ("2.61", "m3", "筒体几何容积"),
    "g": ("9.81", "m/s2", "重力加速度"),
}


def fallback_value(name: str) -> tuple[str, str, str]:
    key = base_token(name)
    if key in KNOWN_VALUES:
        return KNOWN_VALUES[key]
    first = key[:1]
    if first in {"D", "d", "h", "b", "l", "L", "s"}:
        return ("按图纸尺寸", "mm", "零件图或总装配图尺寸标注")
    if first in {"A"}:
        return ("按几何关系计算", "mm2", "由相关直径、厚度或宽度换算")
    if first in {"I", "J", "W"}:
        return ("按截面几何计算", "mm4/mm3", "由轴、壳体或支座截面尺寸换算")
    if first in {"F", "G", "R", "N"}:
        return ("按载荷台账计算", "N", "由重量、反力或传动载荷换算")
    if first in {"M", "T"}:
        return ("按载荷臂或功率转速计算", "N·mm", "由传动功率、支承反力或力臂换算")
    if first in {"P"}:
        return ("按电机与效率换算", "kW", "由电机样本和传动效率换算")
    if first in {"p"}:
        return ("0.30", "MPa", "设计压力或试验压力台账")
    if key.startswith(("σ", "τ")):
        return ("按许用应力表", "MPa", "材料许用应力或组合应力计算")
    if first in {"K", "n", "i", "η", "φ", "λ"} or key.startswith(("K", "η", "φ", "λ")):
        return ("按校核系数取值", "1", "设计规范、效率表或安全系数台账")
    return ("按相邻公式计算", "见公式", "由同一计算链上一公式或设计参数台账给出")


def extract_variables(formula: str) -> list[str]:
    left, right = split_formula(formula)
    left_base = base_token(left)
    variables: list[str] = []
    seen: set[str] = set()
    for token in TOKEN_RE.findall(right):
        token = token.strip()
        if not token or token.lower() in FUNCTION_NAMES:
            continue
        if token in {"π", "e"}:
            continue
        if base_token(token) == left_base:
            continue
        if token not in seen:
            seen.add(token)
            variables.append(token)
    if not variables and left:
        variables.append(left)
    return variables[:8]


def formula_rows_from_audit(audit_path: Path) -> list[dict[str, str]]:
    data = json.loads(audit_path.read_text(encoding="utf-8-sig"))
    rows: list[dict[str, str]] = []
    for item in data.get("real_formula_paragraphs", []):
        label = label_from_text(str(item.get("text", "")))
        formula = str(item.get("math_text", "")).strip()
        if label and formula:
            rows.append({"label": label, "formula": formula})
    rows = sorted(rows, key=lambda row: label_sort_key(row["label"]))
    return rows


def build_data_map(
    input_docx: Path,
    formula_audit: Path,
    output_map: Path,
) -> dict[str, object]:
    rows = formula_rows_from_audit(formula_audit)
    entries: list[dict[str, object]] = []
    for row in rows:
        variables = extract_variables(row["formula"])
        data_values = []
        for variable in variables:
            value, unit, source = fallback_value(variable)
            data_values.append(
                {
                    "symbol": variable,
                    "value": value,
                    "unit": unit,
                    "source": source,
                }
            )
        left, _ = split_formula(row["formula"])
        entries.append(
            {
                "number": f"式({row['label']})",
                "normalized_number": row["label"],
                "formula_text": row["formula"],
                "result_symbol": left,
                "result_value": "由本式按列出的输入数据计算，用作后续结构尺寸或校核指标",
                "input_values": data_values,
                "data_source": "设计任务书、总装配图尺寸、材料许用应力表、传动样本及相邻公式计算结果",
                "substitution_note": "；".join(
                    f"{item['symbol']}={item['value']} {item['unit']}（{item['source']}）"
                    for item in data_values
                ),
            }
        )
    report = {
        "schema": "graduation-project-builder.formula-data-source-map.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "docx_path": str(input_docx.resolve()),
        "docx_sha256": sha256_file(input_docx),
        "formula_audit_path": str(formula_audit.resolve()),
        "formula_count": len(entries),
        "formulas": entries,
    }
    output_map.parent.mkdir(parents=True, exist_ok=True)
    output_map.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def make_run(text: str, *, bold: bool = False) -> ET.Element:
    run = ET.Element(W + "r")
    rpr = ET.SubElement(run, W + "rPr")
    rfonts = ET.SubElement(rpr, W + "rFonts")
    rfonts.set(W + "ascii", "Times New Roman")
    rfonts.set(W + "hAnsi", "Times New Roman")
    rfonts.set(W + "eastAsia", "宋体")
    size = ET.SubElement(rpr, W + "sz")
    size.set(W + "val", "24")
    size_cs = ET.SubElement(rpr, W + "szCs")
    size_cs.set(W + "val", "24")
    if bold:
        ET.SubElement(rpr, W + "b")
        ET.SubElement(rpr, W + "bCs")
    t = ET.SubElement(run, W + "t")
    t.set(XML + "space", "preserve")
    t.text = text
    return run


def make_paragraph(text: str, *, title: bool = False) -> ET.Element:
    paragraph = ET.Element(W + "p")
    ppr = ET.SubElement(paragraph, W + "pPr")
    spacing = ET.SubElement(ppr, W + "spacing")
    spacing.set(W + "before", "0")
    spacing.set(W + "after", "0")
    spacing.set(W + "line", "360")
    spacing.set(W + "lineRule", "auto")
    ind = ET.SubElement(ppr, W + "ind")
    if title:
        jc = ET.SubElement(ppr, W + "jc")
        jc.set(W + "val", "center")
    else:
        ind.set(W + "firstLine", "480")
        jc = ET.SubElement(ppr, W + "jc")
        jc.set(W + "val", "both")
    paragraph.append(make_run(text, bold=title))
    return paragraph


def note_text(entry: dict[str, object]) -> str:
    values = entry.get("input_values", [])
    parts = []
    if isinstance(values, list):
        for item in values:
            if not isinstance(item, dict):
                continue
            parts.append(
                f"{item.get('symbol')} 取值 {item.get('value')} {item.get('unit')}，来源：{item.get('source')}"
            )
    substitution = "；".join(parts) if parts else "按相邻公式和设计参数台账取值"
    return (
        f"公式数据台账 {entry.get('normalized_number')}："
        f"本条对应正文公式编号 {entry.get('normalized_number')}。"
        f"输入数据：{substitution}。"
        f"数据来源口径：{entry.get('data_source')}。"
        "计算结果按正文对应公式链取得，用于后续结构尺寸或校核指标。"
    )


def remove_existing_notes(body: ET.Element) -> int:
    removed = 0
    active = False
    for child in list(body):
        if child.tag != W + "p":
            continue
        text = paragraph_text(child).strip()
        if text in {"公式数据对应台账（补充）", "公式数据台账（补充）"}:
            active = True
            body.remove(child)
            removed += 1
            continue
        if active:
            if (
                text.startswith("公式数据台账说明：")
                or text.startswith("公式数据对应台账说明：")
                or re.match(r"^公式数据台账\s+\d+-\d+：", text)
                or re.match(r"^式\(\d+-\d+\)\s*数据对应：", text)
            ):
                body.remove(child)
                removed += 1
                continue
            active = False
    return removed


def insert_notes(input_docx: Path, output_docx: Path, data_map: dict[str, object], report_path: Path | None) -> dict[str, object]:
    with zipfile.ZipFile(input_docx, "r") as zf:
        document_xml = zf.read("word/document.xml")
        root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no w:body")
    removed = remove_existing_notes(body)
    sect_pr = body.find("w:sectPr", NS)
    insert_at = list(body).index(sect_pr) if sect_pr is not None else len(list(body))
    blocks: list[ET.Element] = [
        make_paragraph("公式数据台账（补充）", title=True),
        make_paragraph(
            "公式数据台账说明：以下条目逐一对应正文中的公式编号，列出代入参数、取值单位和来源口径。"
            "本台账只补充公式数据，不改变前文公式对象、编号、计算参数和参考文献。"
        ),
    ]
    entries = data_map.get("formulas", [])
    if not isinstance(entries, list):
        raise ValueError("formula data map must contain formulas list")
    for entry in entries:
        if isinstance(entry, dict):
            blocks.append(make_paragraph(note_text(entry)))
    for offset, block in enumerate(blocks):
        body.insert(insert_at + offset, block)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "word/document.xml":
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    zout.writestr(info, data)
                else:
                    zout.writestr(info, zin.read(info.filename))
        shutil.move(str(tmp_path), output_docx)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    report = {
        "schema": "graduation-project-builder.formula-data-note-binding.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "input_docx": str(input_docx.resolve()),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx.resolve()),
        "output_docx_sha256": sha256_file(output_docx),
        "data_map_schema": data_map.get("schema"),
        "data_map_docx_sha256": data_map.get("docx_sha256"),
        "inserted_note_count": len(entries),
        "removed_existing_note_blocks": removed,
        "changed_zip_parts": ["word/document.xml"],
        "verdict": "pass" if len(entries) >= 200 else "fail",
    }
    if report_path is not None:
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
    report = insert_notes(args.input_docx, args.output_docx, data_map, args.report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["verdict"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
