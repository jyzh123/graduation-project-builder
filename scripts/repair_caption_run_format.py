#!/usr/bin/env python3
"""Replay template figure/table caption run formatting onto a DOCX copy."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CAPTION_RE = re.compile(
    r"^\s*(?P<label>图|表|续表)\s*"
    r"(?P<number>[0-9A-Za-z一二三四五六七八九十零〇壹贰叁肆伍陆柒捌玖]+"
    r"(?:[\-－\.．][0-9A-Za-z一二三四五六七八九十零〇壹贰叁肆伍陆柒捌玖]+)*)"
    r"\s*(?P<title>.+?)\s*$"
)
CAPTION_PROSE_PREFIXES = (
    "中的",
    "中",
    "所示",
    "显示",
    "展示",
    "体现",
    "说明",
    "可以",
    "可见",
    "为",
    "是",
)

ET.register_namespace("w", W_NS)


def w(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def is_body_heading_text(text: str) -> bool:
    stripped = text.strip()
    return bool(re.match(r"^\d+\s+\S", stripped) or re.match(r"^\d+(?:\.\d+)+\s+\S", stripped))


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def run_visible_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def run_kind(text: str) -> str | None:
    if any("\u4e00" <= ch <= "\u9fff" for ch in text):
        return "cjk"
    if any(ch.isascii() and (ch.isalpha() or ch.isdigit()) for ch in text):
        return "latin"
    return None


def character_kind(char: str) -> str:
    if "\u4e00" <= char <= "\u9fff":
        return "cjk"
    if char.isascii() and (char.isalnum() or char.isspace() or char in "-.,;:()[]/\\"):
        return "latin"
    if ord(char) < 128:
        return "latin"
    return "neutral"


def next_non_neutral(chars: list[str], start: int) -> str | None:
    for char in chars[start:]:
        kind = character_kind(char)
        if kind != "neutral":
            return kind
    return None


def split_caption_text(text: str) -> list[tuple[str, str]]:
    segments: list[tuple[str, str]] = []
    current_kind: str | None = None
    current_chars: list[str] = []
    chars = list(text)
    for index, char in enumerate(chars):
        kind = character_kind(char)
        if kind == "neutral":
            kind = current_kind or next_non_neutral(chars, index + 1) or "cjk"
        if current_kind is None:
            current_kind = kind
        if kind != current_kind:
            if current_chars:
                segments.append((current_kind, "".join(current_chars)))
            current_kind = kind
            current_chars = [char]
        else:
            current_chars.append(char)
    if current_chars and current_kind:
        segments.append((current_kind, "".join(current_chars)))
    return [(kind, payload) for kind, payload in segments if payload]


def caption_kind(text: str) -> str | None:
    match = CAPTION_RE.match(text.strip())
    if not match:
        return None
    title = (match.group("title") or "").strip()
    if not title:
        return None
    compact_title = re.sub(r"\s+", "", title)
    if compact_title.startswith(CAPTION_PROSE_PREFIXES):
        return None
    if len(title) > 80 and re.search(r"[\u3002\uff0c\uff1b,;]", title):
        return None
    return "figure" if match.group("label") == "图" else "table"


def clone_rpr(run: ET.Element | None) -> ET.Element | None:
    if run is None:
        return None
    rpr = run.find("w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def donor_models(donor: ET.Element) -> tuple[ET.Element | None, ET.Element | None, ET.Element | None]:
    fallback_run: ET.Element | None = None
    cjk_run: ET.Element | None = None
    latin_run: ET.Element | None = None
    for run in donor.findall("w:r", NS):
        text = run_visible_text(run)
        if not text.strip():
            continue
        if fallback_run is None:
            fallback_run = run
        kind = run_kind(text)
        if kind == "cjk" and cjk_run is None:
            cjk_run = run
        if kind == "latin" and latin_run is None:
            latin_run = run
    fallback = clone_rpr(fallback_run)
    return clone_rpr(cjk_run) or fallback, clone_rpr(latin_run) or fallback, fallback


def make_run(text: str, rpr: ET.Element | None) -> ET.Element:
    run = ET.Element(w("r"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    text_node = ET.SubElement(run, w("t"))
    if text[:1].isspace() or text[-1:].isspace() or "  " in text:
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    return run


def ensure_keep_next(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(w("pPr"))
        paragraph.insert(0, ppr)
    if ppr.find("w:keepNext", NS) is None:
        ppr.append(ET.Element(w("keepNext")))


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.Element(tag)
        parent.append(child)
    return child


def remove_children(parent: ET.Element, tags: set[str]) -> None:
    for child in list(parent):
        if child.tag in tags:
            parent.remove(child)


def insert_before_any(parent: ET.Element, child: ET.Element, later_tags: set[str]) -> None:
    for index, existing in enumerate(list(parent)):
        if existing.tag in later_tags:
            parent.insert(index, child)
            return
    parent.append(child)


def caption_ppr_without_wrapping(donor: ET.Element) -> ET.Element | None:
    ppr = donor.find("w:pPr", NS)
    cloned = copy.deepcopy(ppr) if ppr is not None else ET.Element(w("pPr"))
    for tag in ("w:framePr", "w:outlineLvl", "w:numPr"):
        for node in list(cloned.findall(tag, NS)):
            cloned.remove(node)
    return cloned


def style_id_from_paragraph(paragraph: ET.Element) -> str | None:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.get(w("val")) if style is not None else None


def ensure_pstyle(ppr: ET.Element, style_id: str) -> None:
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        pstyle = ET.Element(w("pStyle"))
        ppr.insert(0, pstyle)
    pstyle.set(w("val"), style_id)


def find_style(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for style in styles_root.findall("w:style", NS):
        if style.get(w("styleId")) == style_id:
            return style
    return None


def find_caption_style_by_name(styles_root: ET.Element) -> ET.Element | None:
    for style in styles_root.findall("w:style", NS):
        name = style.find("w:name", NS)
        name_value = (name.get(w("val")) if name is not None else "") or ""
        if name_value.strip().lower() in {"caption", "题注", "图题", "表题", "图名", "表名"}:
            return style
    return None


def load_styles(members: dict[str, bytes]) -> ET.Element | None:
    payload = members.get("word/styles.xml")
    if payload is None:
        return None
    return ET.fromstring(payload)


def write_styles(members: dict[str, bytes], styles_root: ET.Element) -> None:
    members["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)


def clone_caption_style(
    members: dict[str, bytes],
    template_members: dict[str, bytes],
    donors: dict[str, ET.Element],
) -> dict[str, object]:
    donor_styles_root = load_styles(template_members)
    target_styles_root = load_styles(members)
    copied: list[str] = []
    missing: list[str] = []
    caption_style_ids = sorted({style_id_from_paragraph(donor) or "Caption" for donor in donors.values()})
    if donor_styles_root is None:
        missing = caption_style_ids
        return {"copied_style_ids": copied, "missing_donor_style_ids": missing, "styles_changed": False}
    if target_styles_root is None:
        target_styles_root = copy.deepcopy(donor_styles_root)
        write_styles(members, target_styles_root)
        return {
            "copied_style_ids": caption_style_ids,
            "missing_donor_style_ids": [],
            "styles_changed": True,
            "styles_copy_mode": "copied-template-styles-part",
        }
    for style_id in caption_style_ids:
        donor_style = find_style(donor_styles_root, style_id)
        if donor_style is None and style_id == "Caption":
            donor_style = find_caption_style_by_name(donor_styles_root)
            if donor_style is not None:
                style_id = donor_style.get(w("styleId")) or style_id
        if donor_style is None:
            missing.append(style_id)
            continue
        existing = find_style(target_styles_root, style_id)
        cloned = copy.deepcopy(donor_style)
        if existing is not None:
            index = list(target_styles_root).index(existing)
            target_styles_root.remove(existing)
            target_styles_root.insert(index, cloned)
        else:
            target_styles_root.append(cloned)
        copied.append(style_id)
    if copied:
        write_styles(members, target_styles_root)
    return {"copied_style_ids": copied, "missing_donor_style_ids": missing, "styles_changed": bool(copied)}


def direct_bookmarks(paragraph: ET.Element) -> list[ET.Element]:
    kept: list[ET.Element] = []
    for child in list(paragraph):
        if child.tag not in {w("pPr"), w("r")}:
            kept.append(copy.deepcopy(child))
    return kept


def set_caption_from_donor(target: ET.Element, donor: ET.Element, kind: str) -> int:
    text = paragraph_text(target)
    ppr = caption_ppr_without_wrapping(donor)
    style_id = style_id_from_paragraph(donor) or "Caption"
    if ppr is None:
        ppr = ET.Element(w("pPr"))
    ensure_pstyle(ppr, style_id)
    kept = direct_bookmarks(target)
    target[:] = []
    if ppr is not None:
        target.append(ppr)
    for item in kept:
        target.append(item)
    cjk_rpr, latin_rpr, fallback_rpr = donor_models(donor)
    run_count = 0
    for segment_kind, payload in split_caption_text(text):
        rpr = latin_rpr if segment_kind == "latin" else cjk_rpr
        target.append(make_run(payload, rpr or fallback_rpr))
        run_count += 1
    if kind == "table":
        ensure_keep_next(target)
    if paragraph_text(target) != text:
        raise RuntimeError(f"caption text changed during repair: {text!r} -> {paragraph_text(target)!r}")
    return run_count


def load_document(path: Path) -> tuple[dict[str, bytes], ET.Element]:
    with zipfile.ZipFile(path) as zf:
        members = {name: zf.read(name) for name in zf.namelist()}
    return members, ET.fromstring(members["word/document.xml"])


def caption_donors(template_docx: Path) -> dict[str, ET.Element]:
    _members, root = load_document(template_docx)
    donors: dict[str, ET.Element] = {}
    for paragraph in root.findall(".//w:body/w:p", NS):
        kind = caption_kind(paragraph_text(paragraph))
        if kind and kind not in donors:
            donors[kind] = paragraph
    # Some official templates contain only one caption exemplar. Figure and
    # table captions share the same visible typography in that case, so keep
    # the repair usable instead of silently skipping one caption family.
    if "figure" not in donors and "table" in donors:
        donors["figure"] = donors["table"]
    if "table" not in donors and "figure" in donors:
        donors["table"] = donors["figure"]
    return donors


def repair_captions(source_docx: Path, template_docx: Path, output_docx: Path) -> dict[str, object]:
    if source_docx.resolve() == output_docx.resolve():
        raise RuntimeError("output DOCX must be a fresh review copy")
    donors = caption_donors(template_docx)
    if not donors:
        raise RuntimeError(f"no figure/table caption donor found in template: {template_docx}")
    template_members, _template_root = load_document(template_docx)
    members, root = load_document(source_docx)
    style_copy_report = clone_caption_style(members, template_members, donors)
    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    touched: list[dict[str, object]] = []
    body_started = False
    for child_index, child in enumerate(list(body)):
        if child.tag != w("p"):
            continue
        text = paragraph_text(child).strip()
        if not body_started and is_body_heading_text(text):
            body_started = True
        if not body_started:
            continue
        if normalized(text) in {normalized("参考文献"), normalized("致谢"), normalized("谢辞"), normalized("附录")}:
            break
        kind = caption_kind(text)
        if kind is None:
            continue
        donor = donors.get(kind)
        if donor is None:
            continue
        run_count = set_caption_from_donor(child, donor, kind)
        touched.append(
            {
                "body_child_index": child_index,
                "kind": kind,
                "text": text,
                "run_count_after": run_count,
            }
        )

    if not touched:
        raise RuntimeError(f"no body figure/table captions found in source DOCX: {source_docx}")

    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    changed_zip_parts = ["word/document.xml"]
    if style_copy_report.get("styles_changed"):
        changed_zip_parts.append("word/styles.xml")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    return {
        "schema": "graduation-project-builder.caption-run-format-repair.v1",
        "source_docx": str(source_docx),
        "source_sha256": sha256_file(source_docx),
        "template_docx": str(template_docx),
        "template_sha256": sha256_file(template_docx),
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "touched_scope": "word/document.xml figure/table caption paragraphs and donor caption style definitions only",
        "changed_zip_parts": changed_zip_parts,
        "caption_style_copy": style_copy_report,
        "caption_count": len(touched),
        "captions": touched,
        "verdict": "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair DOCX figure/table caption run formatting from template donors.")
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    report = repair_captions(
        Path(args.source_docx).resolve(),
        Path(args.template_docx).resolve(),
        Path(args.output_docx).resolve(),
    )
    report_path = Path(args.report).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_docx": report["output_docx"], "caption_count": report["caption_count"], "verdict": "pass"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
