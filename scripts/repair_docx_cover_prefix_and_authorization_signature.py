#!/usr/bin/env python3
"""Repair a DOCX cover admin prefix and copy the author signature to authorization.

This is a narrow front-matter helper for graduation-project-builder runs. It
mutates only ``word/document.xml`` and leaves media, relationships, styles,
numbering, headers, footers, comments, bookmarks, fields, citations, and tables
untouched.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xml": "http://www.w3.org/XML/1998/namespace",
}

W = "{%s}" % NS["w"]
WP = "{%s}" % NS["wp"]
PIC = "{%s}" % NS["pic"]

FORBIDDEN_COVER_TEXTS = {
    "附件8",
    "河北北方学院学士学位论文模板",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def para_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS))


def compact(value: str) -> str:
    return "".join((value or "").split())


def is_paragraph(node: etree._Element) -> bool:
    return node.tag == W + "p"


def runs(p: etree._Element) -> list[etree._Element]:
    return [node for node in p if node.tag == W + "r"]


def set_run_text(run: etree._Element, text: str) -> None:
    rpr = run.find("w:rPr", namespaces=NS)
    preserved_rpr = copy.deepcopy(rpr) if rpr is not None else None
    for child in list(run):
        run.remove(child)
    if preserved_rpr is not None:
        run.append(preserved_rpr)
    t = etree.Element(W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{%s}space" % NS["xml"], "preserve")
    t.text = text
    run.append(t)


def max_drawing_id(root: etree._Element) -> int:
    result = 0
    for node in root.xpath(".//wp:docPr | .//pic:cNvPr", namespaces=NS):
        raw = node.get("id")
        if raw and raw.isdigit():
            result = max(result, int(raw))
    return result


def retarget_cloned_drawing_ids(root: etree._Element, drawing_run: etree._Element) -> list[int]:
    next_id = max_drawing_id(root) + 1
    assigned: list[int] = []
    for node in drawing_run.xpath(".//wp:docPr | .//pic:cNvPr", namespaces=NS):
        if node.get("id") is not None:
            node.set("id", str(next_id))
            assigned.append(next_id)
            next_id += 1
    return assigned


def drawing_embed_ids(run: etree._Element) -> list[str]:
    ids: list[str] = []
    for blip in run.xpath(".//a:blip", namespaces=NS):
        rid = blip.get("{%s}embed" % NS["r"])
        if rid:
            ids.append(rid)
    return ids


def remove_cover_prefix_paragraphs(body: etree._Element) -> list[dict[str, Any]]:
    children = list(body)
    cover_title_seen_early = any(
        is_paragraph(node) and "学士学位论文" in para_text(node)
        for node in children[:12]
    )
    if not cover_title_seen_early:
        raise RuntimeError("cover title was not found in the expected front-matter window")

    removals: list[tuple[int, etree._Element, str, str]] = []
    for idx, node in enumerate(children[:12]):
        if not is_paragraph(node):
            continue
        text = para_text(node)
        normalized = compact(text)
        if normalized in FORBIDDEN_COVER_TEXTS:
            removals.append((idx, node, text, "forbidden-template-prefix"))

    # Some earlier repairs cleared the text but left the two template-prefix
    # paragraph nodes in place. Remove only the first two blank paragraphs when
    # the cover title/logo window proves this is the school front matter.
    if not removals:
        for idx, node in enumerate(children[:2]):
            if is_paragraph(node) and compact(para_text(node)) == "":
                removals.append((idx, node, para_text(node), "blank-former-template-prefix"))

    if not removals:
        return []

    removed: list[dict[str, Any]] = []
    for idx, node, text, reason in sorted(removals, key=lambda item: item[0], reverse=True):
        body.remove(node)
        removed.append({"body_child_index": idx, "text": text, "reason": reason})
    removed.reverse()
    return removed


def find_authorization_anchor(paragraphs: list[etree._Element]) -> int:
    for idx, p in enumerate(paragraphs):
        if "学位论文版权使用授权书" in para_text(p):
            return idx
    raise RuntimeError("authorization heading was not found")


def copy_author_signature(root: etree._Element, body: etree._Element) -> dict[str, Any]:
    paragraphs = [node for node in body if is_paragraph(node)]
    source_idx = None
    source_para = None
    for idx, p in enumerate(paragraphs):
        text = compact(para_text(p))
        if "论文作者（签名）：" in text and "指导教师确认（签名）：" in text and p.xpath(".//w:drawing", namespaces=NS):
            source_idx = idx
            source_para = p
            break
    if source_para is None or source_idx is None:
        raise RuntimeError("source originality-statement author signature drawing was not found")

    auth_idx = find_authorization_anchor(paragraphs)
    target_idx = None
    target_para = None
    for idx in range(auth_idx + 1, min(auth_idx + 12, len(paragraphs))):
        p = paragraphs[idx]
        text = compact(para_text(p))
        if (
            "论文作者（签名）：" in text
            and "指导教师（签名）：" in text
            and "指导教师确认（签名）：" not in text
        ):
            target_idx = idx
            target_para = p
            break
    if target_para is None or target_idx is None:
        raise RuntimeError("target authorization author signature paragraph was not found")

    source_drawing_run = None
    source_after_run = None
    for run in runs(source_para):
        if run.xpath(".//w:drawing", namespaces=NS):
            source_drawing_run = run
        elif source_drawing_run is not None and source_after_run is None and para_text(run):
            source_after_run = run
    if source_drawing_run is None:
        raise RuntimeError("source signature drawing run was not found")

    target_text_runs = [run for run in runs(target_para) if para_text(run)]
    if not target_text_runs:
        raise RuntimeError("target signature paragraph has no text run to preserve formatting from")

    before_run = copy.deepcopy(target_text_runs[0])
    set_run_text(before_run, "论文作者（签名）：    ")

    drawing_run = copy.deepcopy(source_drawing_run)
    assigned_ids = retarget_cloned_drawing_ids(root, drawing_run)

    after_run = copy.deepcopy(source_after_run if source_after_run is not None else target_text_runs[0])
    set_run_text(after_run, "              指导教师（签名）：")

    for run in runs(target_para):
        target_para.remove(run)
    target_para.append(before_run)
    target_para.append(drawing_run)
    target_para.append(after_run)

    return {
        "source_paragraph_index": source_idx,
        "target_paragraph_index": target_idx,
        "source_text_before": para_text(source_para),
        "target_text_after": para_text(target_para),
        "source_drawing_count": len(source_para.xpath(".//w:drawing", namespaces=NS)),
        "target_drawing_count": len(target_para.xpath(".//w:drawing", namespaces=NS)),
        "reused_embed_ids": drawing_embed_ids(drawing_run),
        "new_drawing_object_ids": assigned_ids,
    }


def replace_document_xml(src: Path, dst: Path, document_xml: bytes) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(dst.parent)) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
        shutil.move(str(tmp_path), str(dst))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def repair_docx(src: Path, dst: Path) -> dict[str, Any]:
    with zipfile.ZipFile(src, "r") as zf:
        original_document_xml = zf.read("word/document.xml")

    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    root = etree.fromstring(original_document_xml, parser)
    body = root.find("w:body", namespaces=NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")

    removed_prefix = remove_cover_prefix_paragraphs(body)
    signature = copy_author_signature(root, body)

    remaining_forbidden = []
    for idx, p in enumerate([node for node in body if is_paragraph(node)]):
        text = para_text(p)
        if compact(text) in FORBIDDEN_COVER_TEXTS:
            remaining_forbidden.append({"paragraph_index": idx, "text": text})

    if remaining_forbidden:
        raise RuntimeError(f"forbidden cover prefix text still remains: {remaining_forbidden!r}")

    modified_document_xml = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    if modified_document_xml == original_document_xml:
        raise RuntimeError("repair produced no word/document.xml change")

    replace_document_xml(src, dst, modified_document_xml)
    return {
        "schema": "graduation-project-builder.cover-prefix-authorization-signature-repair.v1",
        "input_docx": str(src),
        "input_sha256": sha256_file(src),
        "output_docx": str(dst),
        "output_sha256": sha256_file(dst),
        "changed_parts": ["word/document.xml"],
        "removed_cover_prefix_paragraphs": removed_prefix,
        "remaining_forbidden_cover_texts": remaining_forbidden,
        "authorization_signature_copy": signature,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--audit-json", type=Path)
    args = parser.parse_args(argv)

    report = repair_docx(args.input, args.output)
    if args.audit_json:
        args.audit_json.parent.mkdir(parents=True, exist_ok=True)
        args.audit_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
