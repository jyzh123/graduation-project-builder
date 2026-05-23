#!/usr/bin/env python3
"""Audit and optionally repair non-black visible font colors in thesis DOCX files."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
STORY_PARTS = {
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
}
STYLE_PARTS = ("word/styles.xml", "word/stylesWithEffects.xml")
BLACK_VALUES = {"", "000000", "auto", "AUTO"}
COLOR_ATTRS_TO_CLEAR = (
    "themeColor",
    "themeTint",
    "themeShade",
    "themeFill",
    "themeFillTint",
    "themeFillShade",
)

ET.register_namespace("w", W_NS)


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def visible_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(f".//{qn('t')}")).strip()


def is_black_color(color: ET.Element) -> bool:
    value = color.attrib.get(qn("val"), "")
    theme = color.attrib.get(qn("themeColor"), "")
    return value in BLACK_VALUES and not theme


def force_black(color: ET.Element) -> None:
    color.attrib[qn("val")] = "000000"
    for local in COLOR_ATTRS_TO_CLEAR:
        color.attrib.pop(qn(local), None)


def used_style_ids(root: ET.Element) -> set[str]:
    ids: set[str] = set()
    for tag in ("pStyle", "rStyle"):
        for element in root.findall(f".//{qn(tag)}"):
            value = element.attrib.get(qn("val"), "")
            if value:
                ids.add(value)
    return ids


def story_direct_color_issues(part: str, root: ET.Element) -> list[dict[str, object]]:
    issues: list[dict[str, object]] = []
    for run_index, run in enumerate(root.findall(f".//{qn('r')}"), start=1):
        text = visible_text(run)
        if not text:
            continue
        for color in run.findall(f".//{qn('color')}"):
            if is_black_color(color):
                continue
            issues.append({
                "kind": "direct-run-color",
                "part": part,
                "run_index": run_index,
                "text": text[:80],
                "value": color.attrib.get(qn("val"), ""),
                "themeColor": color.attrib.get(qn("themeColor"), ""),
            })
    return issues


def style_color_issues(part: str, root: ET.Element, used_ids: set[str]) -> list[dict[str, object]]:
    issues: list[dict[str, object]] = []
    for style in root.findall(qn("style")):
        style_id = style.attrib.get(qn("styleId"), "")
        if style_id not in used_ids:
            continue
        name = style.find(qn("name"))
        name_value = name.attrib.get(qn("val"), "") if name is not None else ""
        for color in style.findall(f".//{qn('color')}"):
            if is_black_color(color):
                continue
            issues.append({
                "kind": "used-style-color",
                "part": part,
                "styleId": style_id,
                "styleName": name_value,
                "value": color.attrib.get(qn("val"), ""),
                "themeColor": color.attrib.get(qn("themeColor"), ""),
            })
    return issues


def copy_docx_with_replacements(source: Path, output: Path, replacements: dict[str, bytes]) -> None:
    if source.resolve() == output.resolve():
        raise ValueError("input and output DOCX paths must be different")
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=str(output.parent))
    os.close(fd)
    Path(tmp_name).unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = replacements.get(item.filename)
                if payload is None:
                    payload = zin.read(item.filename)
                info = zipfile.ZipInfo(item.filename, item.date_time)
                info.comment = item.comment
                info.extra = item.extra
                info.internal_attr = item.internal_attr
                info.external_attr = item.external_attr
                info.create_system = item.create_system
                info.compress_type = item.compress_type
                zout.writestr(info, payload)
        os.replace(tmp_name, output)
    finally:
        Path(tmp_name).unlink(missing_ok=True)


def audit_docx(docx_path: Path, *, repair_output: Path | None = None) -> dict[str, object]:
    issues: list[dict[str, object]] = []
    replacements: dict[str, bytes] = {}
    used_ids: set[str] = set()
    with zipfile.ZipFile(docx_path, "r") as zin:
        existing = set(zin.namelist())
        story_roots: dict[str, ET.Element] = {}
        for part in sorted(STORY_PARTS & existing):
            root = ET.fromstring(zin.read(part))
            story_roots[part] = root
            used_ids.update(used_style_ids(root))
            issues.extend(story_direct_color_issues(part, root))
        for part in STYLE_PARTS:
            if part not in existing:
                continue
            root = ET.fromstring(zin.read(part))
            issues.extend(style_color_issues(part, root, used_ids))
    if repair_output is not None:
        with zipfile.ZipFile(docx_path, "r") as zin:
            existing = set(zin.namelist())
            for part in sorted(STORY_PARTS & existing):
                root = ET.fromstring(zin.read(part))
                changed = False
                for color in root.findall(f".//{qn('color')}"):
                    if is_black_color(color):
                        continue
                    force_black(color)
                    changed = True
                if changed:
                    replacements[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            for part in STYLE_PARTS:
                if part not in existing:
                    continue
                root = ET.fromstring(zin.read(part))
                changed = False
                for style in root.findall(qn("style")):
                    style_id = style.attrib.get(qn("styleId"), "")
                    if style_id not in used_ids:
                        continue
                    for color in style.findall(f".//{qn('color')}"):
                        if is_black_color(color):
                            continue
                        force_black(color)
                        changed = True
                if changed:
                    replacements[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        copy_docx_with_replacements(docx_path, repair_output, replacements)
    return {
        "schema": "graduation-project-builder.docx-font-color-audit.v1",
        "generator": "scripts/audit_docx_font_color.py",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "docx_path": str(docx_path),
        "docx_sha256": sha256_file(docx_path),
        "repair_output_docx_path": str(repair_output) if repair_output else "",
        "repair_output_docx_sha256": sha256_file(repair_output) if repair_output and repair_output.exists() else "",
        "used_style_ids": sorted(used_ids),
        "nonblack_color_count": len(issues),
        "issues": issues,
        "changed_parts": sorted(replacements),
        "passed": not issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx_path", type=Path)
    parser.add_argument("--repair-output-docx", type=Path)
    parser.add_argument("--report-json", type=Path)
    args = parser.parse_args()
    report = audit_docx(args.docx_path, repair_output=args.repair_output_docx)
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["passed"] or args.repair_output_docx else 1


if __name__ == "__main__":
    raise SystemExit(main())
