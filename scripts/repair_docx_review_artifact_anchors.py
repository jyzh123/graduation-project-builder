#!/usr/bin/env python3
"""Restore DOCX comment anchors from a source manuscript by stable paragraph id.

This helper is intentionally narrow: it restores missing commentRangeStart,
commentRangeEnd, and commentReference markers on final paragraphs whose
`w14:paraId`/`w15:paraId` matches the source. It does not change comment text,
mark comments done, accept revisions, or rewrite surrounding prose.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
W = f"{{{NS['w']}}}"
W14 = f"{{{NS['w14']}}}"
W15 = f"{{{NS['w15']}}}"

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_id(paragraph: ET.Element) -> str:
    return paragraph.attrib.get(f"{W14}paraId") or paragraph.attrib.get(f"{W15}paraId") or paragraph.attrib.get("paraId") or ""


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def has_picture(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None


def picture_child_indices(paragraph: ET.Element) -> list[int]:
    result: list[int] = []
    for index, child in enumerate(list(paragraph)):
        if (
            child.find(".//w:drawing", NS) is not None
            or child.find(".//w:pict", NS) is not None
            or child.find(".//a:blip", NS) is not None
            or child.find(".//v:imagedata", NS) is not None
        ):
            result.append(index)
    return result


def source_comment_anchor_map(source_docx: Path) -> dict[str, dict[str, object]]:
    with zipfile.ZipFile(source_docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    rows: dict[str, dict[str, object]] = {}
    for paragraph in root.findall(".//w:p", NS):
        pid = paragraph_id(paragraph)
        if not pid:
            continue
        ids = {
            child.attrib.get(f"{W}id", "")
            for child in paragraph.iter()
            if child.tag in {f"{W}commentRangeStart", f"{W}commentRangeEnd", f"{W}commentReference"}
            and child.attrib.get(f"{W}id", "")
        }
        if not ids:
            continue
        for comment_id in ids:
            rows[comment_id] = {
                "comment_id": comment_id,
                "para_id": pid,
                "source_text": paragraph_text(paragraph),
                "source_has_picture": has_picture(paragraph),
            }
    return rows


def final_anchor_ids(root: ET.Element) -> set[tuple[str, str]]:
    found: set[tuple[str, str]] = set()
    for child in root.iter():
        if child.tag not in {f"{W}commentRangeStart", f"{W}commentRangeEnd", f"{W}commentReference"}:
            continue
        comment_id = child.attrib.get(f"{W}id", "")
        if comment_id:
            found.add((child.tag.rsplit("}", 1)[-1], comment_id))
    return found


def make_anchor(tag_name: str, comment_id: str) -> ET.Element:
    element = ET.Element(f"{W}{tag_name}")
    element.set(f"{W}id", comment_id)
    return element


def make_reference_run(comment_id: str) -> ET.Element:
    run = ET.Element(f"{W}r")
    ref = ET.SubElement(run, f"{W}commentReference")
    ref.set(f"{W}id", comment_id)
    return run


def paragraph_content_child_indices(paragraph: ET.Element) -> list[int]:
    result: list[int] = []
    for index, child in enumerate(list(paragraph)):
        if child.tag in {f"{W}pPr", f"{W}commentRangeStart", f"{W}commentRangeEnd", f"{W}bookmarkStart", f"{W}bookmarkEnd"}:
            continue
        if (
            child.find(".//w:t", NS) is not None
            or child.find(".//w:drawing", NS) is not None
            or child.find(".//w:pict", NS) is not None
            or child.find(".//w:fldChar", NS) is not None
            or child.find(".//w:instrText", NS) is not None
            or child.find(".//w:tab", NS) is not None
        ):
            result.append(index)
    return result


def restore_anchor_on_paragraph(paragraph: ET.Element, comment_id: str) -> bool:
    present = final_anchor_ids(paragraph)
    required = {
        ("commentRangeStart", comment_id),
        ("commentRangeEnd", comment_id),
        ("commentReference", comment_id),
    }
    if required.issubset(present):
        return False
    picture_indices = picture_child_indices(paragraph)
    content_indices = picture_indices or paragraph_content_child_indices(paragraph)
    if not content_indices:
        return False
    first_content = content_indices[0]
    last_content = content_indices[-1]
    if ("commentRangeStart", comment_id) not in present:
        paragraph.insert(first_content, make_anchor("commentRangeStart", comment_id))
        last_content += 1
    if ("commentRangeEnd", comment_id) not in present:
        paragraph.insert(last_content + 1, make_anchor("commentRangeEnd", comment_id))
        last_content += 1
    if ("commentReference", comment_id) not in present:
        paragraph.insert(last_content + 1, make_reference_run(comment_id))
    return True


def repair(source_docx: Path, final_docx: Path, output_docx: Path, comment_ids: set[str] | None = None) -> dict[str, object]:
    source_rows = source_comment_anchor_map(source_docx)
    with zipfile.ZipFile(final_docx) as zin:
        members = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    root = ET.fromstring(members["word/document.xml"])
    by_para_id = {paragraph_id(paragraph): paragraph for paragraph in root.findall(".//w:p", NS) if paragraph_id(paragraph)}

    restored: list[dict[str, object]] = []
    skipped: list[dict[str, object]] = []
    for comment_id, row in sorted(source_rows.items(), key=lambda item: int(item[0]) if re.fullmatch(r"\d+", item[0]) else item[0]):
        if comment_ids is not None and comment_id not in comment_ids:
            continue
        pid = str(row["para_id"])
        paragraph = by_para_id.get(pid)
        if paragraph is None:
            skipped.append({"comment_id": comment_id, "para_id": pid, "reason": "matching final paragraph not found"})
            continue
        if restore_anchor_on_paragraph(paragraph, comment_id):
            restored.append({"comment_id": comment_id, "para_id": pid, "final_text": paragraph_text(paragraph)})
        else:
            skipped.append({"comment_id": comment_id, "para_id": pid, "reason": "matching final paragraph has no restorable content or already complete"})

    members["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_output = Path(temp_dir) / "out.docx"
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in members.items():
                zout.writestr(name, data)
        output_docx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(temp_output, output_docx)

    return {
        "schema": "graduation-project-builder.repair-docx-review-artifact-anchors.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(source_docx.resolve()),
        "source_docx_sha256": sha256_file(source_docx),
        "input_final_docx": str(final_docx.resolve()),
        "input_final_docx_sha256": sha256_file(final_docx),
        "output_docx": str(output_docx.resolve()),
        "output_docx_sha256": sha256_file(output_docx),
        "restored": restored,
        "skipped": skipped,
    }


def parse_comment_ids(value: str | None) -> set[str] | None:
    if not value:
        return None
    return {part.strip() for part in value.split(",") if part.strip()}


def main() -> int:
    parser = argparse.ArgumentParser(description="Restore missing DOCX comment anchors by source paragraph id.")
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--final-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--comment-ids", help="Optional comma-separated comment ids to restore.")
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    report = repair(
        Path(args.source_docx),
        Path(args.final_docx),
        Path(args.output_docx),
        parse_comment_ids(args.comment_ids),
    )
    report_path = Path(args.report)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
