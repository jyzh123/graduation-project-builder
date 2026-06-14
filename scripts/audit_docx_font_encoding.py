#!/usr/bin/env python3
"""Audit a DOCX for mojibake and bibliography font-slot drift."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = NS["w"]

TEXT_PARTS = {
    "word/document.xml",
    "word/styles.xml",
    "word/fontTable.xml",
    "word/comments.xml",
}

FONT_ATTR_RE = re.compile(r'w:(?:ascii|hAnsi|eastAsia|cs|name|val)="([^"]+)"')
VISIBLE_TEXT_RE = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.DOTALL)
WESTERN_REFERENCE_CHARS = set("[](){}.,;:/\\-+_=&%#@'\"<>")
REFERENCE_ENTRY_RE = re.compile(r"^\s*(?:\[(\d{1,3})\]|(\d{1,3})[\.\u3001])\s*\S")
BIBLIOGRAPHY_VISIBLE_LABEL_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\d+[\.\u3001])")
MIN_BIBLIOGRAPHY_ENTRY_CONTENT_CHARS = 8
FONT_FAMILY_KEYS = ("eastAsia", "ascii", "hAnsi", "cs")
FONT_THEME_KEYS = ("eastAsiaTheme", "asciiTheme", "hAnsiTheme", "csTheme")
CHINESE_FONT_SIZE_HALF_POINTS = {
    "初号": "84",
    "小初": "72",
    "一号": "52",
    "小一": "48",
    "二号": "44",
    "小二": "36",
    "三号": "32",
    "小三": "30",
    "四号": "28",
    "小四": "24",
    "五号": "21",
    "小五": "18",
    "六号": "15",
    "小六": "13",
    "七号": "11",
    "八号": "10",
}
REFERENCE_ENTRY_POLICY_HINTS = ("序号", "内容", "顶格", "方括号", "条目")
KNOWN_CJK_FONT_FAMILIES = ("华文中宋", "宋体", "黑体", "楷体", "仿宋", "SimSun")
CHINESE_SIZE_NAME_RE = "|".join(
    re.escape(name) for name in sorted(CHINESE_FONT_SIZE_HALF_POINTS, key=len, reverse=True)
)
KNOWN_CJK_FONT_RE = "|".join(re.escape(name) for name in sorted(KNOWN_CJK_FONT_FAMILIES, key=len, reverse=True))

# Build suspicious tokens programmatically where possible so this source does
# not itself contain replacement characters or private-use mojibake samples.
SUSPICIOUS_GENERIC_TOKENS = (
    "\ufffd",
    "?" * 4,
    chr(0x95C2) + "?",
    chr(0x95C1) + "?",
    chr(0x7F02) + "?",
)

SUSPICIOUS_FONT_TOKENS = (
    chr(0x59D2) + chr(0x6D9B) + chr(0x57B3) + chr(0x7F4D) + "?",
    chr(0x940E) + chr(0x747E) + chr(0xE0C4) + chr(0x7F4D) + "?",
    chr(0x6FE1) + chr(0x8BFE) + chr(0x6E39) + chr(0x7F4D) + "?",
    chr(0x5A34) + chr(0x72B2) + chr(0x7042) + chr(0x9E2C) + "?",
    chr(0x7035) + chr(0x90A6) + chr(0x55DB) + chr(0x9482) + "?",
)


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def iter_parts(zf: zipfile.ZipFile) -> list[str]:
    names: list[str] = []
    for name in zf.namelist():
        if name in TEXT_PARTS or re.fullmatch(r"word/(header|footer)\d+\.xml", name):
            names.append(name)
    return sorted(names)


def decode_part(zf: zipfile.ZipFile, name: str) -> str:
    return zf.read(name).decode("utf-8", errors="replace")


def contains_suspicious_token(text: str, tokens: tuple[str, ...]) -> bool:
    return any(token in text for token in tokens)


def normalize(value: str) -> str:
    return re.sub(r"\s+", "", value or "").lower()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def contains_cjk(text: str) -> bool:
    return any("\u3400" <= char <= "\u9fff" for char in text)


def contains_western_reference_text(text: str) -> bool:
    return any(char.isascii() and (char.isalnum() or char in WESTERN_REFERENCE_CHARS) for char in text)


def contains_western_letters(text: str) -> bool:
    return any(char.isascii() and char.isalpha() for char in text)


def script_class(text: str) -> str:
    has_cjk = contains_cjk(text)
    has_western = contains_western_reference_text(text)
    if has_cjk and has_western:
        return "mixed"
    if has_cjk:
        return "cjk"
    if has_western:
        return "latin"
    return "neutral"


def is_reference_entry_text(text: str) -> bool:
    return bool(REFERENCE_ENTRY_RE.match(text or ""))


def is_reference_label_only_text(text: str) -> bool:
    if not BIBLIOGRAPHY_VISIBLE_LABEL_RE.match(text or ""):
        return False
    content = BIBLIOGRAPHY_VISIBLE_LABEL_RE.sub("", text or "", count=1)
    content = re.sub(r"^\s*[\.\u3001\u3002\]\)）:：,，;；-]+", "", content)
    return len(re.sub(r"\s+", "", content)) < MIN_BIBLIOGRAPHY_ENTRY_CONTENT_CHARS


def reference_entry_number(text: str) -> int | None:
    match = REFERENCE_ENTRY_RE.match(text or "")
    if not match:
        return None
    return int(match.group(1) or match.group(2))


def is_bibliography_heading(text: str) -> bool:
    normalized = normalize(text)
    if normalized in {normalize("\u53c2\u8003\u6587\u732e"), "references", "bibliography"}:
        return True
    wanted = {normalize("参考文献"), "references", "bibliography"}
    for alias in wanted:
        if normalized in {alias, f"{alias}:", f"{alias}："}:
            return True
    return False


def contains_bibliography_marker(text: str) -> bool:
    normalized = normalize(text)
    if normalize("\u53c2\u8003\u6587\u732e") in normalized:
        return True
    if is_reference_entry_text(text):
        return False
    if is_bibliography_heading(text):
        return True
    return any(alias in normalized for alias in (normalize("参考文献"), "references", "bibliography"))


def is_tail_heading_after_bibliography(text: str) -> bool:
    normalized = normalize(text)
    if normalized in {normalize("\u81f4\u8c22"), normalize("\u9644\u5f55")}:
        return True
    return normalized in {
        normalize("致谢"),
        normalize("附录"),
        "acknowledgements",
        "acknowledgments",
        "appendix",
    }


def iter_body_paragraphs(root: ET.Element) -> list[tuple[int, ET.Element, str]]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    result: list[tuple[int, ET.Element, str]] = []
    for index, child in enumerate(list(body)):
        if child.tag != w("p"):
            continue
        result.append((index, child, paragraph_text(child).strip()))
    return result


def bibliography_entry_paragraph_elements(root: ET.Element) -> list[tuple[int, ET.Element]]:
    paragraphs = iter_body_paragraphs(root)
    candidates: list[tuple[int, int, list[tuple[int, ET.Element]]]] = []
    for position, (_body_index, _paragraph, text) in enumerate(paragraphs):
        if not text or not contains_bibliography_marker(text):
            continue
        skipped_non_entries = 0
        entries: list[tuple[int, ET.Element]] = []
        for lookahead_position in range(position + 1, len(paragraphs)):
            body_index, paragraph, candidate_text = paragraphs[lookahead_position]
            if not candidate_text:
                continue
            if is_tail_heading_after_bibliography(candidate_text) and entries:
                break
            has_auto_numbering = paragraph.find("./w:pPr/w:numPr", NS) is not None
            if is_reference_entry_text(candidate_text) or has_auto_numbering:
                entries.append((body_index, paragraph))
                continue
            if entries:
                break
            skipped_non_entries += 1
            if skipped_non_entries > 5:
                break
        if entries:
            first_number = reference_entry_number(paragraph_text(entries[0][1]).strip())
            score = len(entries) * 10
            if is_bibliography_heading(text):
                score += 30
            if first_number == 1:
                score += 10
            if entries[0][1].find("./w:pPr/w:numPr", NS) is not None:
                score += 10
            score -= skipped_non_entries
            candidates.append((score, position, entries))
    if not candidates:
        return []
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]


def bibliography_label_only_paragraphs(docx_path: Path) -> list[tuple[int, str]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    entry_indexes = {
        index for index, _paragraph in bibliography_entry_paragraph_elements(root)
    }
    hits: list[tuple[int, str]] = []
    inside_bibliography = False
    for body_index, paragraph, text in iter_body_paragraphs(root):
        if is_bibliography_heading(text):
            inside_bibliography = True
            continue
        if inside_bibliography and is_tail_heading_after_bibliography(text):
            break
        has_auto_numbering = paragraph.find("./w:pPr/w:numPr", NS) is not None
        has_entry_signal = (
            body_index in entry_indexes
            or (inside_bibliography and (has_auto_numbering or BIBLIOGRAPHY_VISIBLE_LABEL_RE.match(text or "")))
        )
        if not has_entry_signal:
            continue
        if is_reference_label_only_text(text) or (
            has_auto_numbering and len(re.sub(r"\s+", "", text or "")) < MIN_BIBLIOGRAPHY_ENTRY_CONTENT_CHARS
        ):
            hits.append((body_index, text))
    return hits


def signature_from_rpr(rpr: ET.Element | None) -> dict[str, str]:
    signature: dict[str, str] = {}
    if rpr is None:
        return signature
    fonts = rpr.find("./w:rFonts", NS)
    if fonts is not None:
        for key in (*FONT_FAMILY_KEYS, *FONT_THEME_KEYS):
            value = fonts.get(w(key))
            if value:
                signature[key] = value
    size = rpr.find("./w:sz", NS)
    if size is not None and size.get(w("val")):
        signature["size"] = size.get(w("val")) or ""
    size_cs = rpr.find("./w:szCs", NS)
    if size_cs is not None and size_cs.get(w("val")):
        signature["sizeCs"] = size_cs.get(w("val")) or ""
    if rpr.find("./w:b", NS) is not None:
        signature["bold"] = "yes"
    if rpr.find("./w:i", NS) is not None:
        signature["italic"] = "yes"
    if rpr.find("./w:u", NS) is not None:
        signature["underline"] = "yes"
    vert_align = rpr.find("./w:vertAlign", NS)
    if vert_align is not None and vert_align.get(w("val")):
        signature["vertAlign"] = vert_align.get(w("val")) or ""
    return signature


def run_signature(run: ET.Element) -> dict[str, str]:
    rpr = run.find("./w:rPr", NS)
    return signature_from_rpr(rpr)


def paragraph_num_pr(paragraph: ET.Element) -> dict[str, str]:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return {}
    num_pr = ppr.find("./w:numPr", NS)
    if num_pr is None:
        return {}
    result: dict[str, str] = {}
    ilvl = num_pr.find("./w:ilvl", NS)
    num_id = num_pr.find("./w:numId", NS)
    if ilvl is not None and ilvl.get(w("val")):
        result["ilvl"] = ilvl.get(w("val")) or ""
    if num_id is not None and num_id.get(w("val")):
        result["numId"] = num_id.get(w("val")) or ""
    return result


def load_numbering_root(docx_path: Path) -> ET.Element | None:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            return ET.fromstring(zf.read("word/numbering.xml"))
    except KeyError:
        return None


def numbering_model_for_num_id(numbering_root: ET.Element | None, num_id: str) -> dict[str, str]:
    if numbering_root is None or not num_id:
        return {}
    num_node = next(
        (
            node
            for node in numbering_root.findall("./w:num", NS)
            if node.get(w("numId")) == num_id
        ),
        None,
    )
    if num_node is None:
        return {}
    abstract_ref = num_node.find("./w:abstractNumId", NS)
    abstract_id = abstract_ref.get(w("val"), "") if abstract_ref is not None else ""
    abstract = next(
        (
            node
            for node in numbering_root.findall("./w:abstractNum", NS)
            if node.get(w("abstractNumId")) == abstract_id
        ),
        None,
    )
    if abstract is None:
        return {"numId": num_id, "abstractNumId": abstract_id}
    lvl = abstract.find("./w:lvl", NS)
    lvl_text = lvl.find("./w:lvlText", NS) if lvl is not None else None
    num_fmt = lvl.find("./w:numFmt", NS) if lvl is not None else None
    ind = lvl.find("./w:pPr/w:ind", NS) if lvl is not None else None
    return {
        "numId": num_id,
        "abstractNumId": abstract_id,
        "lvlText": lvl_text.get(w("val"), "") if lvl_text is not None else "",
        "numFmt": num_fmt.get(w("val"), "") if num_fmt is not None else "",
        "left": ind.get(w("left"), "") if ind is not None else "",
        "hanging": ind.get(w("hanging"), "") if ind is not None else "",
    }


def bibliography_numbering_models(docx_path: Path, entries: list[dict[str, object]]) -> list[dict[str, str]]:
    numbering_root = load_numbering_root(docx_path)
    seen: set[str] = set()
    models: list[dict[str, str]] = []
    for entry in entries:
        num_pr = entry.get("numPr")
        if not isinstance(num_pr, dict):
            continue
        num_id = str(num_pr.get("numId", ""))
        if not num_id or num_id in seen:
            continue
        seen.add(num_id)
        models.append(numbering_model_for_num_id(numbering_root, num_id))
    return models


def paragraph_style_id(paragraph: ET.Element) -> str:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return ""
    pstyle = ppr.find("./w:pStyle", NS)
    return pstyle.get(w("val"), "") if pstyle is not None else ""


def run_style_id(run: ET.Element) -> str:
    rpr = run.find("./w:rPr", NS)
    if rpr is None:
        return ""
    rstyle = rpr.find("./w:rStyle", NS)
    return rstyle.get(w("val"), "") if rstyle is not None else ""


def load_style_tables(docx_path: Path) -> dict[str, object]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
    except KeyError:
        return {"docDefaults": {}, "styles": {}}

    doc_defaults = signature_from_rpr(styles_root.find("./w:docDefaults/w:rPrDefault/w:rPr", NS))
    styles: dict[str, dict[str, object]] = {}
    for style in styles_root.findall("./w:style", NS):
        style_id = style.get(w("styleId"), "")
        if not style_id:
            continue
        based_on = style.find("./w:basedOn", NS)
        name = style.find("./w:name", NS)
        styles[style_id] = {
            "basedOn": based_on.get(w("val"), "") if based_on is not None else "",
            "name": name.get(w("val"), "") if name is not None else "",
            "type": style.get(w("type"), ""),
            "signature": signature_from_rpr(style.find("./w:rPr", NS)),
        }
    return {"docDefaults": doc_defaults, "styles": styles}


def style_chain_signature(style_id: str, styles: dict[str, dict[str, object]], seen: set[str] | None = None) -> dict[str, str]:
    if not style_id or style_id not in styles:
        return {}
    if seen is None:
        seen = set()
    if style_id in seen:
        return {}
    seen.add(style_id)
    row = styles[style_id]
    signature: dict[str, str] = {}
    based_on = str(row.get("basedOn", ""))
    signature.update(style_chain_signature(based_on, styles, seen))
    own_signature = row.get("signature", {})
    if isinstance(own_signature, dict):
        signature.update({key: str(value) for key, value in own_signature.items() if value})
    return signature


def merge_signature_with_owner(
    result: dict[str, str],
    owners: dict[str, str],
    signature: dict[str, str],
    owner: str,
) -> None:
    for key, value in signature.items():
        if value:
            result[key] = value
            owners[key] = owner


def effective_run_signature(
    run: ET.Element,
    paragraph_style: str,
    style_tables: dict[str, object],
) -> tuple[dict[str, str], dict[str, str]]:
    effective: dict[str, str] = {}
    owners: dict[str, str] = {}
    doc_defaults = style_tables.get("docDefaults", {})
    if isinstance(doc_defaults, dict):
        merge_signature_with_owner(effective, owners, {key: str(value) for key, value in doc_defaults.items()}, "docDefaults")

    styles_raw = style_tables.get("styles", {})
    styles: dict[str, dict[str, object]] = styles_raw if isinstance(styles_raw, dict) else {}
    if paragraph_style:
        merge_signature_with_owner(
            effective,
            owners,
            style_chain_signature(paragraph_style, styles),
            f"pStyle:{paragraph_style}",
        )

    character_style = run_style_id(run)
    if character_style:
        merge_signature_with_owner(
            effective,
            owners,
            style_chain_signature(character_style, styles),
            f"rStyle:{character_style}",
        )

    merge_signature_with_owner(effective, owners, run_signature(run), "direct")
    return effective, owners


def font_canonical(value: str) -> str:
    normalized = re.sub(r"[\s()（）]+", "", value or "").lower()
    aliases = {
        "simsun": "宋体",
        "songti": "宋体",
        "宋体": "宋体",
        "timesnewroman": "timesnewroman",
    }
    return aliases.get(normalized, normalized)


def font_canonical(value: str) -> str:
    raw_parts = [part for part in re.split(r"[;；,，/]+", value or "") if part.strip()]
    aliases = {
        "simsun": "\u5b8b\u4f53",
        "songti": "\u5b8b\u4f53",
        "\u5b8b\u4f53": "\u5b8b\u4f53",
        "simhei": "\u9ed1\u4f53",
        "\u9ed1\u4f53": "\u9ed1\u4f53",
        "kaiti": "\u6977\u4f53",
        "\u6977\u4f53": "\u6977\u4f53",
        "fangsong": "\u4eff\u5b8b",
        "\u4eff\u5b8b": "\u4eff\u5b8b",
        "timesnewroman": "timesnewroman",
    }
    for part in raw_parts or [value or ""]:
        normalized = re.sub(r"[\s()\uff08\uff09]+", "", part or "").lower()
        if normalized in aliases:
            return aliases[normalized]
    normalized = re.sub(r"[\s()\uff08\uff09]+", "", value or "").lower()
    return aliases.get(normalized, normalized)


def font_matches(actual: str, expected: str) -> bool:
    return bool(actual and expected and font_canonical(actual) == font_canonical(expected))


def resolve_font_size_half_points(size_name: str | None = None, size_half_points: str | None = None) -> str | None:
    if size_name and size_half_points:
        raise ValueError("choose either a Chinese font-size name or a half-point value, not both")
    if size_name:
        normalized = size_name.strip()
        if normalized not in CHINESE_FONT_SIZE_HALF_POINTS:
            raise ValueError(f"unknown Chinese font-size name: {size_name}")
        return CHINESE_FONT_SIZE_HALF_POINTS[normalized]
    if size_half_points:
        normalized = size_half_points.strip()
        if not normalized.isdigit() or int(normalized) <= 0:
            raise ValueError(f"invalid half-point font size: {size_half_points}")
        return normalized
    return None


def font_size_name_for_half_points(size_half_points: str | None) -> str | None:
    if not size_half_points:
        return None
    for name, value in CHINESE_FONT_SIZE_HALF_POINTS.items():
        if value == size_half_points:
            return name
    return None


def positive_bibliography_stats(docx_path: Path) -> tuple[int, int]:
    entries = collect_bibliography_entries(docx_path)
    run_count = sum(len(entry.get("script_runs") or entry.get("runs") or []) for entry in entries)
    return len(entries), run_count


def bibliography_content_completeness_hits(docx_path: Path) -> list[str]:
    return [
        f"paragraph {index + 1} bibliography entry has only a label or too little substantive content: `{text}`"
        for index, text in bibliography_label_only_paragraphs(docx_path)
    ]


def bibliography_intrinsic_format_hits(docx_path: Path) -> list[str]:
    entries = collect_bibliography_entries(docx_path)
    issues: list[str] = []
    for entry_index, entry in enumerate(entries, start=1):
        entry_text = str(entry.get("text", "")).strip()
        match = REFERENCE_ENTRY_RE.match(entry_text)
        if not match:
            continue
        leading_token_len = len(match.group(0).strip())
        consumed = 0
        leading_rows: list[dict[str, object]] = []
        for row in entry.get("script_runs", []):  # type: ignore[union-attr]
            if not isinstance(row, dict):
                continue
            text = str(row.get("text", ""))
            if not text:
                continue
            if consumed < leading_token_len:
                leading_rows.append(row)
            consumed += len(text)
            if consumed >= leading_token_len:
                break
        for row in leading_rows:
            signature = row.get("signature")
            if not isinstance(signature, dict):
                continue
            vert_align = str(signature.get("vertAlign", ""))
            if vert_align in {"superscript", "subscript"}:
                issues.append(
                    f"entry {entry_index} bibliography leading number has {vert_align}; "
                    f"reference entry numbers must be normal baseline text; "
                    f"text=`{str(row.get('text', ''))[:60] or entry_text[:60]}`"
                )
    return issues


def _wps_payload_value(payload: dict[str, object], *names: str) -> object:
    for name in names:
        if name in payload:
            return payload[name]
    return None


def validate_wps_named_size_evidence(
    evidence_path: Path,
    *,
    docx_path: Path,
    expected_size_name: str | None,
    expected_size_half_points: str | None,
    expected_entry_count: int,
) -> list[str]:
    issues: list[str] = []
    try:
        payload = json.loads(evidence_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        return [f"WPS named-size evidence is not valid JSON: {evidence_path} ({exc})"]

    schema = str(payload.get("schema") or "")
    if schema != "graduation-project-builder.wps-reference-entry-ui-font.v1":
        issues.append(
            "WPS named-size evidence must use schema "
            "graduation-project-builder.wps-reference-entry-ui-font.v1"
        )

    verdict = str(_wps_payload_value(payload, "verdict", "overallVerdict") or "").lower()
    if verdict != "pass":
        issues.append(f"WPS named-size evidence verdict is not pass: {evidence_path}")

    raw_docx_path = _wps_payload_value(payload, "docxPath", "docx_path", "docx")
    if raw_docx_path:
        resolved_docx_path = Path(str(raw_docx_path))
        if not resolved_docx_path.is_absolute():
            resolved_docx_path = (evidence_path.parent / resolved_docx_path).resolve()
        else:
            resolved_docx_path = resolved_docx_path.resolve()
        if resolved_docx_path != docx_path.resolve():
            issues.append(
                f"WPS named-size evidence targets {resolved_docx_path} instead of {docx_path.resolve()}"
            )
    else:
        issues.append("WPS named-size evidence must record docxPath/docx_path")

    docx_sha = str(_wps_payload_value(payload, "docxSha256", "docx_sha256") or "")
    expected_docx_sha = sha256_file(docx_path)
    if docx_sha.lower() != expected_docx_sha.lower():
        issues.append("WPS named-size evidence DOCX sha256 does not match the audited DOCX")

    if expected_size_name:
        payload_size_name = str(
            _wps_payload_value(payload, "expectedWpsDisplaySizeName", "expectedDisplaySizeName", "expected_size_name")
            or ""
        )
        if payload_size_name != expected_size_name:
            issues.append(
                f"WPS named-size evidence expected display size is {payload_size_name or '<missing>'}, "
                f"expected {expected_size_name}"
            )
    if expected_size_half_points:
        payload_half_points = str(
            _wps_payload_value(payload, "expectedSizeHalfPoints", "expected_size_half_points")
            or ""
        )
        if payload_half_points != expected_size_half_points:
            issues.append(
                f"WPS named-size evidence expected half-points is {payload_half_points or '<missing>'}, "
                f"expected {expected_size_half_points}"
            )

    entries = payload.get("entries")
    checked_entry_count = _wps_payload_value(payload, "checkedEntryCount", "checked_entry_count")
    if checked_entry_count is None and isinstance(entries, list):
        checked_entry_count = len(entries)
    try:
        checked_entry_count_int = int(str(checked_entry_count))
    except (TypeError, ValueError):
        issues.append("WPS named-size evidence must record checkedEntryCount or an entries array")
        checked_entry_count_int = -1
    if expected_entry_count > 0 and checked_entry_count_int != expected_entry_count:
        issues.append(
            f"WPS named-size evidence checked {checked_entry_count_int} bibliography entries; "
            f"expected {expected_entry_count}"
        )

    if isinstance(entries, list):
        for index, entry in enumerate(entries, start=1):
            if not isinstance(entry, dict):
                issues.append(f"WPS named-size evidence entry {index} is not an object")
                continue
            entry_verdict = str(entry.get("verdict") or "").lower()
            display_name = str(
                entry.get("inferredWpsDisplaySizeName")
                or entry.get("wpsDisplaySizeName")
                or entry.get("displaySizeName")
                or ""
            )
            if entry_verdict != "pass":
                issues.append(f"WPS named-size evidence entry {index} verdict is not pass")
            if expected_size_name and display_name != expected_size_name:
                issues.append(
                    f"WPS named-size evidence entry {index} display size is {display_name or '<missing>'}, "
                    f"expected {expected_size_name}"
                )
    else:
        issues.append("WPS named-size evidence must include all-entry entries[] proof")

    return issues


def extract_reference_font_policy(docx_path: Path) -> dict[str, str]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            document_root = ET.fromstring(zf.read("word/document.xml"))
    except KeyError:
        return {}
    paragraph_lines = [
        paragraph_text(paragraph).strip()
        for paragraph in document_root.findall(".//w:p", NS)
        if paragraph_text(paragraph).strip()
    ]
    text_node_lines = [node.text or "" for node in document_root.findall(".//w:t", NS)]
    visible_text = "\n".join(paragraph_lines)
    reference_policy_texts = [line for line in paragraph_lines if "参考文献" in line]
    reference_entry_policy_texts = [
        line for line in reference_policy_texts if any(hint in line for hint in REFERENCE_ENTRY_POLICY_HINTS)
    ]
    search_texts = reference_policy_texts + text_node_lines + [visible_text]
    policy: dict[str, str] = {}
    for text in reference_entry_policy_texts:
        entry_match = re.search(
            rf"(?:用|为|采用)?\s*(?P<size>{CHINESE_SIZE_NAME_RE})\s*"
            rf"(?P<cjk>{KNOWN_CJK_FONT_RE})"
            rf"(?:和|与|、|及|\s)*"
            rf"(?P<latin>Times\s*New\s*Roman|[A-Za-z][A-Za-z ]*[A-Za-z])?",
            text,
            re.IGNORECASE,
        )
        if not entry_match:
            continue
        size_name = entry_match.group("size")
        if size_name in CHINESE_FONT_SIZE_HALF_POINTS:
            size_half_points = CHINESE_FONT_SIZE_HALF_POINTS[size_name]
            policy["sizeName"] = size_name
            policy["size"] = size_half_points
            policy["sizeCs"] = size_half_points
        cjk_family = entry_match.group("cjk")
        if cjk_family:
            policy["eastAsia"] = cjk_family
        latin_family = entry_match.group("latin") or ""
        if latin_family:
            latin = re.sub(r"\s+", " ", latin_family.strip())
            if normalize(latin) == normalize("Times New Roman"):
                latin = "Times New Roman"
            policy["latin"] = latin
        break
    for text in search_texts:
        for pattern in (
            r"(?:参考文献中)?(?:中文|汉字)(?:使用|为|用|字体为|字体使用)\s*([A-Za-z\u4e00-\u9fff]+)",
            r"(?:参考文献中)?(?:中文|汉字)[^。；;，,）)]{0,12}(?:使用|为|用|字体为|字体使用)\s*([A-Za-z\u4e00-\u9fff]+)",
            r"汉字用\s*([^，,。；;\s]+)",
        ):
            east_asia_match = re.search(pattern, text)
            if east_asia_match:
                raw = east_asia_match.group(1).strip()
                for known_family in ("宋体", "黑体", "楷体", "仿宋", "SimSun", "Times New Roman"):
                    if known_family in raw:
                        policy["eastAsia"] = known_family
                        break
                if "eastAsia" not in policy:
                    policy["eastAsia"] = re.sub(r"(?:小四|四号|五号|5号字|5号|号字|号)$", "", raw).strip()
                break
        if "eastAsia" in policy:
            break
    for text in search_texts:
        for pattern in (
            r"(?:数字[、,，和]?英文|英文[、,，和]?数字|数字、英文、标点符号|英文|西文)"
            r"[^。；;，,）)]{0,16}(?:使用|为|用|字体为|字体使用|字体均是|均是|是)\s*"
            r"([A-Za-z][A-Za-z ]*[A-Za-z])\s*(?:体|字体|[0-9０-９]|号|，|,|。|；|;|\)|）)",
            r"英文用\s*([A-Za-z][A-Za-z ]*[A-Za-z])\s*体",
        ):
            latin_match = re.search(pattern, text)
            if latin_match:
                latin = re.sub(r"\s+", " ", latin_match.group(1).strip())
                latin = latin.replace("NewRoman", "New Roman")
                policy["latin"] = latin
                break
        if "latin" in policy:
            break
    return policy


def collect_bibliography_entries(docx_path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    style_tables = load_style_tables(docx_path)

    entries: list[dict[str, object]] = []
    for body_index, child in bibliography_entry_paragraph_elements(root):
        text = paragraph_text(child).strip()
        if not text:
            continue

        runs: list[dict[str, object]] = []
        p_style = paragraph_style_id(child)
        for run in child.findall("w:r", NS):
            if run.find(".//w:txbxContent", NS) is not None:
                continue
            run_text = paragraph_text(run)
            if not run_text.strip():
                continue
            effective_signature, effective_owners = effective_run_signature(run, p_style, style_tables)
            runs.append(
                {
                    "text": run_text,
                    "script": script_class(run_text),
                    "signature": run_signature(run),
                    "effective_signature": effective_signature,
                    "effective_owners": effective_owners,
                }
            )
        entries.append(
            {
                "text": text,
                "script_runs": runs,
                "numPr": paragraph_num_pr(child),
                "pStyle": p_style,
                "body_child_index": body_index,
            }
        )
    return entries


def expected_bibliography_font_models(entries: list[dict[str, object]]) -> tuple[dict[str, str], dict[str, str]]:
    cjk_expected: dict[str, str] = {}
    latin_expected: dict[str, str] = {}
    cjk_score = -1
    latin_score = -1
    for entry in entries:
        for row in entry.get("script_runs", []):  # type: ignore[union-attr]
            if not isinstance(row, dict):
                continue
            script = str(row.get("script", ""))
            signature = row.get("signature")
            if not isinstance(signature, dict):
                continue
            effective_signature = row.get("effective_signature", {})
            effective_owners = row.get("effective_owners", {})
            if not isinstance(effective_signature, dict):
                effective_signature = {}
            if not isinstance(effective_owners, dict):
                effective_owners = {}

            def locked_value(key: str) -> str:
                direct_value = str(signature.get(key, ""))
                if direct_value:
                    return direct_value
                effective_value = str(effective_signature.get(key, ""))
                owner = str(effective_owners.get(key, ""))
                if effective_value and font_owner_is_locked(owner):
                    return effective_value
                return ""

            if script in {"cjk", "mixed"}:
                candidate = {
                    key: locked_value(key)
                    for key in ("eastAsia", "eastAsiaTheme", "size", "sizeCs")
                }
                candidate_score = sum(1 for key in ("eastAsia", "size", "sizeCs") if candidate.get(key))
                if candidate_score > cjk_score:
                    cjk_expected = candidate
                    cjk_score = candidate_score
            if script in {"latin", "mixed"}:
                candidate = {
                    key: locked_value(key)
                    for key in ("ascii", "hAnsi", "cs", "asciiTheme", "hAnsiTheme", "csTheme", "size", "sizeCs")
                }
                candidate_score = sum(1 for key in ("ascii", "hAnsi", "cs", "size", "sizeCs") if candidate.get(key))
                if candidate_score > latin_score:
                    latin_expected = candidate
                    latin_score = candidate_score
    return cjk_expected, latin_expected


def apply_reference_font_policy(
    cjk_expected: dict[str, str],
    latin_expected: dict[str, str],
    policy: dict[str, str],
) -> None:
    east_asia = policy.get("eastAsia", "")
    latin = policy.get("latin", "")
    if east_asia and not cjk_expected.get("eastAsia"):
        cjk_expected["eastAsia"] = east_asia
    if latin:
        for key in ("ascii", "hAnsi", "cs"):
            if not latin_expected.get(key):
                latin_expected[key] = latin
    for size_key in ("size", "sizeCs"):
        if policy.get(size_key):
            cjk_expected.setdefault(size_key, policy[size_key])
            latin_expected.setdefault(size_key, policy[size_key])


def apply_explicit_reference_font_policy(
    cjk_expected: dict[str, str],
    latin_expected: dict[str, str],
    *,
    cjk_font: str | None = None,
    latin_font: str | None = None,
) -> None:
    if cjk_font:
        cjk_expected["eastAsia"] = cjk_font
    if latin_font:
        for key in ("ascii", "hAnsi", "cs"):
            latin_expected[key] = latin_font


def apply_reference_size_policy(
    cjk_expected: dict[str, str],
    latin_expected: dict[str, str],
    expected_size_half_points: str | None,
) -> None:
    if not expected_size_half_points:
        return
    for model in (cjk_expected, latin_expected):
        model["size"] = expected_size_half_points
        model["sizeCs"] = expected_size_half_points


def font_owner_is_locked(owner: str) -> bool:
    return owner == "direct" or owner.startswith("pStyle:") or owner.startswith("rStyle:")


def assert_expected_font(
    *,
    issues: list[str],
    entry_index: int,
    run_index: int,
    text: str,
    script_label: str,
    key: str,
    expected: str,
    signature: dict[str, str],
    effective_signature: dict[str, str],
    effective_owners: dict[str, str],
) -> None:
    if not expected:
        return
    direct_value = str(signature.get(key, ""))
    effective_value = str(effective_signature.get(key, ""))
    owner = str(effective_owners.get(key, ""))
    if not font_matches(effective_value, expected):
        issues.append(
            f"entry {entry_index} run {run_index} {script_label} effective font-chain drift: "
            f"expected {key}=`{expected}` but effective `{effective_value}` from `{owner or 'unresolved'}`; "
            f"direct `{direct_value}`; text=`{text[:60]}`"
        )
        return
    if not font_owner_is_locked(owner):
        theme_key = f"{key}Theme"
        issues.append(
            f"entry {entry_index} run {run_index} {script_label} font family not locked at direct/style level: "
            f"expected {key}=`{expected}` but owner is `{owner or 'unresolved'}` "
            f"(theme `{signature.get(theme_key, effective_signature.get(theme_key, ''))}`); text=`{text[:60]}`"
        )


def bibliography_font_slot_hits(
    reference_docx: Path,
    final_docx: Path,
    *,
    expected_size_half_points: str | None = None,
    expected_size_name: str | None = None,
    wps_named_size_evidence_valid: bool = False,
    bibliography_cjk_font: str | None = None,
    bibliography_latin_font: str | None = None,
) -> list[str]:
    reference_entries = collect_bibliography_entries(reference_docx)
    final_entries = collect_bibliography_entries(final_docx)
    reference_policy = extract_reference_font_policy(reference_docx)
    issues: list[str] = []
    if not reference_entries and not any(
        reference_policy.get(key) for key in ("eastAsia", "latin", "size", "sizeCs")
    ):
        return [f"reference bibliography baseline not found in {reference_docx}"]
    if not final_entries:
        return [f"final bibliography entries not found in {final_docx}"]

    if reference_entries:
        cjk_expected, latin_expected = expected_bibliography_font_models(reference_entries)
    else:
        cjk_expected, latin_expected = {}, {}
    apply_reference_font_policy(cjk_expected, latin_expected, reference_policy)
    apply_explicit_reference_font_policy(
        cjk_expected,
        latin_expected,
        cjk_font=bibliography_cjk_font,
        latin_font=bibliography_latin_font,
    )
    if expected_size_half_points and reference_entries:
        donor_sizes = {
            value
            for model in (cjk_expected, latin_expected)
            for key in ("size", "sizeCs")
            for value in [str(model.get(key, ""))]
            if value
        }
        conflicting_sizes = sorted(size for size in donor_sizes if size != expected_size_half_points)
        if conflicting_sizes:
            for model in (cjk_expected, latin_expected):
                for key in ("size", "sizeCs"):
                    if str(model.get(key, "")) in conflicting_sizes:
                        model.pop(key, None)
    apply_reference_size_policy(cjk_expected, latin_expected, expected_size_half_points)
    preflight_issue_count = len(issues)
    cjk_has_comparable_model = bool(cjk_expected) and any(
        cjk_expected.get(key) for key in ("eastAsia", "eastAsiaTheme", "size", "sizeCs")
    )
    latin_has_comparable_model = bool(latin_expected) and any(
        latin_expected.get(key)
        for key in ("ascii", "hAnsi", "cs", "asciiTheme", "hAnsiTheme", "csTheme", "size", "sizeCs")
    )
    if not cjk_has_comparable_model:
        issues.append("reference bibliography lacks a Chinese run donor for eastAsia/size comparison")
    if not latin_has_comparable_model:
        issues.append("reference bibliography lacks a Western run donor for ascii/hAnsi/cs/size comparison")
    if not cjk_expected.get("eastAsia"):
        issues.append("reference bibliography lacks a locked Chinese font-family donor or instruction-derived eastAsia policy")
    if not any(latin_expected.get(key) for key in ("ascii", "hAnsi", "cs")):
        issues.append("reference bibliography lacks a locked Western font-family donor or instruction-derived Times/Latin policy")
    if len(issues) > preflight_issue_count:
        return issues

    reference_auto_models = bibliography_numbering_models(reference_docx, reference_entries)
    final_auto_models = bibliography_numbering_models(final_docx, final_entries)
    reference_uses_auto_numbering = any(
        model.get("lvlText") == "[%1]" and model.get("numFmt") == "decimal"
        for model in reference_auto_models
    )
    final_uses_valid_auto_numbering = any(
        model.get("lvlText") == "[%1]" and model.get("numFmt") == "decimal"
        for model in final_auto_models
    )
    if reference_uses_auto_numbering:
        manual_visible_labels = all(re.match(r"^\[\d+\]", str(entry.get("text", "")).strip()) for entry in final_entries)
        final_has_no_numpr = all(not entry.get("numPr") for entry in final_entries)
        if not final_uses_valid_auto_numbering:
            if manual_visible_labels and final_has_no_numpr:
                pass
            else:
                issues.append("final bibliography lacks the template automatic numbering model lvlText=[%1], numFmt=decimal")
        if any(re.match(r"^\[\d+\]", str(entry.get("text", "")).strip()) for entry in final_entries) and any(entry.get("numPr") for entry in final_entries):
            issues.append("final bibliography mixes visible manual prefixes with automatic numbering")

    reference_uses_split_runs = any(len(entry.get("script_runs", [])) > 1 for entry in reference_entries)
    for entry_index, entry in enumerate(final_entries, start=1):
        entry_text = str(entry.get("text", "")).strip()
        entry_runs = entry.get("script_runs", [])
        if (
            reference_uses_split_runs
            and isinstance(entry_runs, list)
            and len(entry_runs) <= 1
            and script_class(entry_text) == "mixed"
        ):
            issues.append(
                f"entry {entry_index} collapsed to one bibliography text run while the reference baseline uses split runs; "
                f"text=`{entry_text[:60]}`"
            )
        if entry.get("numPr") and not re.match(r"^\[\d+\]", entry_text) and not reference_uses_auto_numbering:
            issues.append(
                f"entry {entry_index} uses automatic bibliography numbering; numbering.xml marker font-slot audit is required before pass"
            )
        for run_index, row in enumerate(entry.get("script_runs", []), start=1):  # type: ignore[union-attr]
            if not isinstance(row, dict):
                continue
            text = str(row.get("text", ""))
            script = str(row.get("script", ""))
            signature = row.get("signature")
            if not isinstance(signature, dict):
                continue
            effective_signature = row.get("effective_signature", {})
            if not isinstance(effective_signature, dict):
                effective_signature = {}
            effective_owners = row.get("effective_owners", {})
            if not isinstance(effective_owners, dict):
                effective_owners = {}
            if re.match(r"^\s*\[\d+\]", text) and str(signature.get("vertAlign", "")) in {"superscript", "subscript"}:
                issues.append(
                    f"entry {entry_index} bibliography leading number is {str(signature.get('vertAlign', ''))}; "
                    f"reference entry numbers must be normal baseline text; text=`{text[:60] or entry_text[:60]}`"
                )
            if script == "mixed":
                issues.append(
                    f"entry {entry_index} run {run_index} mixed-script run was not split; "
                    f"text=`{text[:60] or entry_text[:60]}`"
                )
            if script in {"cjk", "mixed"}:
                for key, expected in cjk_expected.items():
                    if key == "sizeCs" and cjk_expected.get("size"):
                        continue
                    owner = str(effective_owners.get(key, ""))
                    direct_value = str(signature.get(key, ""))
                    effective_value = str(effective_signature.get(key, ""))
                    if key in FONT_FAMILY_KEYS:
                        if font_matches(direct_value, expected) or (
                            font_matches(effective_value, expected)
                            and font_owner_is_locked(owner)
                        ):
                            continue
                    elif direct_value == expected or (
                        effective_value == expected
                        and font_owner_is_locked(owner)
                    ):
                        continue
                    if expected:
                        issues.append(
                            f"entry {entry_index} run {run_index} Chinese font-slot drift: "
                            f"expected {key}=`{expected}` but found `{signature.get(key, '')}`; "
                            f"text=`{text[:60] or entry_text[:60]}`"
                        )
                        break
                assert_expected_font(
                    issues=issues,
                    entry_index=entry_index,
                    run_index=run_index,
                    text=text or entry_text,
                    script_label="Chinese",
                    key="eastAsia",
                    expected=cjk_expected.get("eastAsia", ""),
                    signature={key: str(value) for key, value in signature.items()},
                    effective_signature={key: str(value) for key, value in effective_signature.items()},
                    effective_owners={key: str(value) for key, value in effective_owners.items()},
                )
            if script in {"latin", "mixed"}:
                if (
                    contains_western_letters(text)
                    and latin_expected.get("sizeCs")
                    and not latin_expected.get("size")
                    and not signature.get("size")
                ):
                    issues.append(
                        f"entry {entry_index} run {run_index} Western visible font-size slot missing: "
                        f"expected size=`{latin_expected.get('sizeCs')}` mirrored from template sizeCs but found empty; "
                        f"text=`{text[:60] or entry_text[:60]}`"
                    )
                    break
                for font_key in ("ascii", "hAnsi", "cs"):
                    assert_expected_font(
                        issues=issues,
                        entry_index=entry_index,
                        run_index=run_index,
                        text=text or entry_text,
                        script_label="Western",
                        key=font_key,
                        expected=latin_expected.get(font_key, ""),
                        signature={key: str(value) for key, value in signature.items()},
                        effective_signature={key: str(value) for key, value in effective_signature.items()},
                        effective_owners={key: str(value) for key, value in effective_owners.items()},
                    )
                for key, expected in latin_expected.items():
                    if key == "sizeCs" and latin_expected.get("size"):
                        continue
                    if expected and str(signature.get(key, "")) != expected:
                        owner = str(effective_owners.get(key, ""))
                        direct_value = str(signature.get(key, ""))
                        effective_value = str(effective_signature.get(key, ""))
                        if key in FONT_FAMILY_KEYS:
                            if font_matches(direct_value, expected) or (
                                font_matches(effective_value, expected)
                                and font_owner_is_locked(owner)
                            ):
                                continue
                        elif effective_value == expected and font_owner_is_locked(owner):
                            continue
                        issues.append(
                            f"entry {entry_index} run {run_index} Western font-slot drift: "
                            f"expected {key}=`{expected}` but found `{signature.get(key, '')}`; "
                            f"text=`{text[:60] or entry_text[:60]}`"
                        )
                        break
            if len(issues) >= 24:
                return issues
    return issues


def make_report(
    docx_path: Path,
    checked_parts: list[str],
    font_hits: list[str],
    text_hits: list[str],
    generic_hits: list[str],
    *,
    reference_docx: Path | None,
    bibliography_hits: list[str],
    expected_size_half_points: str | None,
    expected_size_name: str | None,
    size_policy_source: str,
    bibliography_entry_count: int,
    bibliography_checked_run_count: int,
    bibliography_content_format_model_status: str,
    bibliography_content_format_model_source: str,
    bibliography_content_completeness_hits: list[str],
    wps_ui_evidence_path: Path | None,
    wps_ui_evidence_verdict: str,
) -> str:
    bibliography_status = (
        "fail"
        if bibliography_hits
        else ("pass" if reference_docx is not None else "intrinsic-only-pass")
    )
    result = "pass" if not (font_hits or text_hits or generic_hits or bibliography_hits) else "fail"
    lines = [
        "# DOCX Font And Encoding Audit",
        "",
        f"- docx path: {docx_path}",
        f"- docx sha256: {sha256_file(docx_path)}",
        f"- checked parts: {'; '.join(checked_parts)}",
        f"- suspicious font-name hits: {len(font_hits)}",
        f"- suspicious visible-text hits: {len(text_hits)}",
        f"- suspicious generic-token hits: {len(generic_hits)}",
        f"- bibliography reference docx path: {reference_docx if reference_docx is not None else 'none'}",
        f"- bibliography reference docx sha256: {sha256_file(reference_docx) if reference_docx is not None else 'none'}",
        f"- bibliography expected size name: {expected_size_name if expected_size_name else 'template-derived'}",
        f"- bibliography expected size half-points: {expected_size_half_points if expected_size_half_points else 'template-derived'}",
        f"- bibliography size policy source: {size_policy_source}",
        f"- bibliography entry count: {bibliography_entry_count}",
        f"- bibliography checked run count: {bibliography_checked_run_count}",
        f"- bibliography content-format model checks: {bibliography_content_format_model_status}",
        f"- bibliography content-format model source: {bibliography_content_format_model_source}",
        f"- bibliography empty-entry/content completeness checks: {'fail' if bibliography_content_completeness_hits else 'pass'}",
        f"- bibliography empty-entry/content completeness hits: {len(bibliography_content_completeness_hits)}",
        f"- bibliography named-size WPS evidence path: {wps_ui_evidence_path if wps_ui_evidence_path is not None else 'not-required'}",
        f"- bibliography named-size WPS evidence verdict: {wps_ui_evidence_verdict}",
        f"- bibliography font-slot checks: {bibliography_status}",
        f"- bibliography font-slot hits: {len(bibliography_hits)}",
        f"- result: {result}",
        "",
        "## Font Hits",
    ]
    if font_hits:
        lines.extend(f"- {item}" for item in font_hits)
    else:
        lines.append("- none")
    lines.extend(["", "## Visible Text Hits"])
    if text_hits:
        lines.extend(f"- {item}" for item in text_hits)
    else:
        lines.append("- none")
    lines.extend(["", "## Generic Token Hits"])
    if generic_hits:
        lines.extend(f"- {item}" for item in generic_hits)
    else:
        lines.append("- none")
    lines.extend(["", "## Bibliography Font-Slot Hits"])
    if bibliography_hits:
        lines.extend(f"- {item}" for item in bibliography_hits)
    else:
        lines.append("- none")
    if bibliography_content_completeness_hits:
        lines.extend(["", "## Bibliography Content Completeness Hits"])
        lines.extend(f"- {item}" for item in bibliography_content_completeness_hits)
    return "\n".join(lines) + "\n"


def audit_docx(docx_path: Path) -> tuple[list[str], list[str], list[str], list[str]]:
    checked_parts: list[str] = []
    font_hits: list[str] = []
    text_hits: list[str] = []
    generic_hits: list[str] = []

    with zipfile.ZipFile(docx_path) as zf:
        for name in iter_parts(zf):
            payload = decode_part(zf, name)
            checked_parts.append(name)

            if contains_suspicious_token(payload, SUSPICIOUS_GENERIC_TOKENS):
                generic_hits.append(f"{name}: suspicious generic token present")

            for match in FONT_ATTR_RE.finditer(payload):
                value = match.group(1)
                if contains_suspicious_token(value, SUSPICIOUS_FONT_TOKENS):
                    font_hits.append(f"{name}: font attribute contains mojibake `{value}`")

            for text_match in VISIBLE_TEXT_RE.finditer(payload):
                visible = text_match.group(1)
                if contains_suspicious_token(visible, SUSPICIOUS_GENERIC_TOKENS):
                    text_hits.append(f"{name}: visible text contains suspicious token `{visible[:80]}`")

    return checked_parts, font_hits, text_hits, generic_hits


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit a DOCX for font-name, encoding, and bibliography font-slot corruption.")
    parser.add_argument("docx_path", help="Path to the DOCX to audit")
    parser.add_argument("--reference-docx", help="Template/reference DOCX used to audit bibliography Chinese/Western font slots")
    parser.add_argument("--bibliography-size-name", help="Exact Chinese size name for bibliography entries, for example 五号")
    parser.add_argument("--bibliography-size-half-points", help="Exact bibliography entry size in DOCX half-points, for example 21 for 五号")
    parser.add_argument("--bibliography-cjk-font", help="Instruction-derived Chinese font family for bibliography entries")
    parser.add_argument("--bibliography-latin-font", help="Instruction-derived Western font family for bibliography entries")
    parser.add_argument("--bibliography-wps-ui-evidence-json", help="WPS UI all-entry named-size evidence JSON")
    parser.add_argument("--require-wps-named-size-evidence", action="store_true", help="Fail when explicit named-size WPS UI evidence is absent")
    parser.add_argument("--report", help="Optional markdown report output path")
    args = parser.parse_args()

    docx_path = Path(args.docx_path).resolve()
    reference_docx = Path(args.reference_docx).resolve() if args.reference_docx else None
    try:
        expected_size_half_points = resolve_font_size_half_points(args.bibliography_size_name, args.bibliography_size_half_points)
    except ValueError as exc:
        parser.error(str(exc))
    reference_policy = extract_reference_font_policy(reference_docx) if reference_docx else {}
    if expected_size_half_points is None and reference_policy:
        expected_size_half_points = reference_policy.get("size") or reference_policy.get("sizeCs")
    expected_size_name = (
        args.bibliography_size_name
        or font_size_name_for_half_points(expected_size_half_points)
        or reference_policy.get("sizeName")
    )
    size_policy_source = (
        "explicit-named-size-cli"
        if args.bibliography_size_name
        else ("explicit-half-points-cli" if args.bibliography_size_half_points else "template-derived")
    )
    checked_parts, font_hits, text_hits, generic_hits = audit_docx(docx_path)
    bibliography_entry_count, bibliography_checked_run_count = positive_bibliography_stats(docx_path)
    bibliography_hits = bibliography_intrinsic_format_hits(docx_path)
    bibliography_empty_hits = bibliography_content_completeness_hits(docx_path)
    bibliography_hits.extend(bibliography_empty_hits)
    bibliography_content_format_model_status = "not-applicable"
    bibliography_content_format_model_source = "no bibliography entries detected"
    if bibliography_entry_count > 0:
        if reference_docx:
            bibliography_content_format_model_status = "pass"
            bibliography_content_format_model_source = str(reference_docx)
        else:
            bibliography_content_format_model_status = "fail"
            bibliography_content_format_model_source = "missing --reference-docx"
            bibliography_hits.append(
                "bibliography content-format model requires --reference-docx; "
                "intrinsic-only bibliography audit cannot prove reference entry paragraph/run formatting"
            )
    wps_ui_evidence_path = Path(args.bibliography_wps_ui_evidence_json).resolve() if args.bibliography_wps_ui_evidence_json else None
    wps_ui_evidence_verdict = "not-required"
    if args.require_wps_named_size_evidence or args.bibliography_size_name:
        if wps_ui_evidence_path is None:
            bibliography_hits.append("WPS named-size evidence missing for explicit bibliography size-name policy")
            wps_ui_evidence_verdict = "missing"
        else:
            wps_issues = validate_wps_named_size_evidence(
                wps_ui_evidence_path,
                docx_path=docx_path,
                expected_size_name=expected_size_name,
                expected_size_half_points=expected_size_half_points,
                expected_entry_count=bibliography_entry_count,
            )
            bibliography_hits.extend(wps_issues)
            wps_ui_evidence_verdict = "pass" if not wps_issues else "fail"
    if reference_docx:
        model_issues = bibliography_font_slot_hits(
            reference_docx,
            docx_path,
            expected_size_half_points=expected_size_half_points,
            expected_size_name=expected_size_name,
            wps_named_size_evidence_valid=wps_ui_evidence_verdict == "pass",
            bibliography_cjk_font=args.bibliography_cjk_font,
            bibliography_latin_font=args.bibliography_latin_font,
        )
        bibliography_hits.extend(model_issues)
        if model_issues:
            bibliography_content_format_model_status = "fail"
    report = make_report(
        docx_path,
        checked_parts,
        font_hits,
        text_hits,
        generic_hits,
        reference_docx=reference_docx,
        bibliography_hits=bibliography_hits,
        expected_size_half_points=expected_size_half_points,
        expected_size_name=expected_size_name,
        size_policy_source=size_policy_source,
        bibliography_entry_count=bibliography_entry_count,
        bibliography_checked_run_count=bibliography_checked_run_count,
        bibliography_content_format_model_status=bibliography_content_format_model_status,
        bibliography_content_format_model_source=bibliography_content_format_model_source,
        bibliography_content_completeness_hits=bibliography_empty_hits,
        wps_ui_evidence_path=wps_ui_evidence_path,
        wps_ui_evidence_verdict=wps_ui_evidence_verdict,
    )

    if args.report:
        report_path = Path(args.report).resolve()
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(report, encoding="utf-8")
    else:
        sys.stdout.write(report)

    return 0 if not (font_hits or text_hits or generic_hits or bibliography_hits) else 1


if __name__ == "__main__":
    raise SystemExit(main())
