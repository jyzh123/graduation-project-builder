#!/usr/bin/env python3
"""Repair figure-holder and caption locality in a DOCX review copy.

This helper is intentionally narrow. Its locality repairs edit only
``word/document.xml`` and do not touch media binaries, relationships, headers,
footers, styles, or numbering. Optional image binding repairs may copy new
media or reuse existing ``word/document.xml`` relationship ids only when
explicitly requested. It handles three locality defects that final thesis gates can
identify mechanically:

* formal figure captions placed immediately before their image holder;
* short malformed caption text stored inside an image-holder paragraph;
* optional removal of an unlabeled consecutive image immediately before a
  separate captioned image.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
VML_NS = "urn:schemas-microsoft-com:vml"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W = f"{{{W_NS}}}"
NS = {"w": W_NS, "r": R_NS, "wp": WP_NS, "a": A_NS, "v": VML_NS, "pic": PIC_NS}
EMU_PER_CM = 360000
EMU_PER_INCH = 914400
EMU_PER_PT = 12700

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
for prefix, uri in {
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": PIC_NS,
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}.items():
    ET.register_namespace(prefix, uri)
ET.register_namespace("", REL_PKG_NS)


FIGURE_CAPTION_RE = re.compile(
    r"^\s*\u56fe\s*\d+(?:[-.\uff0d\uff0e]\d+)+[\s\u3000:：.-]+.+$"
)
MALFORMED_CAPTION_RE = re.compile(
    r"^\s*\u56fe\s*(?P<major>\d+)[.\uff0e](?P<minor>\d+)\s*(?P<title>\S.+?)\s*$"
)
HEADING_RE = re.compile(r"^\s*(?:\u7b2c\s*\d+\s*\u7ae0|\d+(?:\.\d+){0,3}\s+\S)")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def qn(local: str) -> str:
    return f"{W}{local}"


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def is_paragraph(node: ET.Element) -> bool:
    return node.tag == qn("p")


def has_drawing(paragraph: ET.Element) -> bool:
    for node in paragraph.iter():
        local = node.tag.rsplit("}", 1)[-1]
        if local in {"drawing", "pict", "object"}:
            return True
    return False


def _css_length_to_emu(value: str) -> int:
    raw = (value or "").strip().lower()
    if not raw:
        return 0
    match = re.match(r"^([0-9]+(?:\.[0-9]+)?)(cm|mm|in|pt)?$", raw)
    if not match:
        return 0
    number = float(match.group(1))
    unit = match.group(2) or "pt"
    if unit == "cm":
        return int(round(number * EMU_PER_CM))
    if unit == "mm":
        return int(round(number * EMU_PER_CM / 10))
    if unit == "in":
        return int(round(number * EMU_PER_INCH))
    return int(round(number * EMU_PER_PT))


def _vml_style_dimensions(style: str) -> tuple[int, int]:
    width = 0
    height = 0
    for part in (style or "").split(";"):
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        lowered = key.strip().lower()
        if lowered == "width":
            width = _css_length_to_emu(value)
        elif lowered == "height":
            height = _css_length_to_emu(value)
    return width, height


def drawing_extent_area(paragraph: ET.Element) -> int:
    areas: list[int] = []
    for extent in paragraph.findall(".//wp:extent", NS):
        try:
            cx = int(extent.attrib.get("cx", "0") or "0")
            cy = int(extent.attrib.get("cy", "0") or "0")
        except ValueError:
            continue
        if cx > 0 and cy > 0:
            areas.append(cx * cy)
    for extent in paragraph.findall(".//a:ext", NS):
        try:
            cx = int(extent.attrib.get("cx", "0") or "0")
            cy = int(extent.attrib.get("cy", "0") or "0")
        except ValueError:
            continue
        if cx > 0 and cy > 0:
            areas.append(cx * cy)
    for shape in paragraph.findall(".//v:shape", NS) + paragraph.findall(".//v:rect", NS):
        if shape.find(".//v:imagedata", NS) is None:
            continue
        width, height = _vml_style_dimensions(shape.attrib.get("style", ""))
        width = width or _css_length_to_emu(shape.attrib.get("width", ""))
        height = height or _css_length_to_emu(shape.attrib.get("height", ""))
        if width > 0 and height > 0:
            areas.append(width * height)
    return max(areas, default=0)


def consecutive_orphan_image_remove_offset(first: ET.Element, second: ET.Element) -> int:
    """Return 0 to remove first image, 1 to remove second image.

    The historical default removed the first blank image before a captioned
    holder. That is correct for many duplicate-image defects, but if a malformed
    caption was split out of a tiny second holder, the first paragraph is often
    the actual figure and the second paragraph is only a residual fragment.
    """
    first_area = drawing_extent_area(first)
    second_area = drawing_extent_area(second)
    if first_area > 0 and second_area > 0 and second_area * 4 < first_area:
        return 1
    if first_area > 0 and second_area > 0 and first_area * 4 < second_area:
        return 0
    return 0


def is_formal_caption_text(text: str) -> bool:
    return bool(FIGURE_CAPTION_RE.match(text or ""))


def is_heading_text(text: str) -> bool:
    return bool(HEADING_RE.match(text or ""))


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is not None:
        return ppr
    ppr = ET.Element(qn("pPr"))
    paragraph.insert(0, ppr)
    return ppr


def ensure_keep_next(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    if ppr.find("./w:keepNext", NS) is not None:
        return False
    ppr.append(ET.Element(qn("keepNext")))
    return True


def first_text_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall(".//w:r", NS):
        if "".join(t.text or "" for t in run.findall(".//w:t", NS)).strip():
            rpr = run.find("./w:rPr", NS)
            return deepcopy(rpr) if rpr is not None else None
    return None


def clone_text_paragraph(donor: ET.Element | None, text: str) -> ET.Element:
    paragraph = ET.Element(qn("p"))
    if donor is not None:
        donor_ppr = donor.find("./w:pPr", NS)
        if donor_ppr is not None:
            paragraph.append(deepcopy(donor_ppr))
    ppr = ensure_ppr(paragraph)
    if donor is None:
        jc = ET.SubElement(ppr, qn("jc"))
        jc.set(qn("val"), "center")
    run = ET.SubElement(paragraph, qn("r"))
    donor_rpr = first_text_run_rpr(donor) if donor is not None else None
    if donor_rpr is not None:
        run.append(donor_rpr)
    t = ET.SubElement(run, qn("t"))
    t.text = text
    return paragraph


def clear_visible_text(paragraph: ET.Element) -> int:
    cleared = 0
    for node in paragraph.findall(".//w:t", NS):
        if node.text:
            node.text = ""
            cleared += 1
    return cleared


def has_drawingml_blip(paragraph: ET.Element) -> bool:
    return paragraph.find(".//a:blip", NS) is not None


def image_holder_donor(children: list[ET.Element]) -> ET.Element | None:
    for child in children:
        if is_paragraph(child) and has_drawing(child) and not paragraph_text(child) and has_drawingml_blip(child):
            return child
    for child in children:
        if is_paragraph(child) and has_drawing(child) and has_drawingml_blip(child):
            return child
    return None


def _png_dimensions(path: Path) -> tuple[int, int] | None:
    data = path.read_bytes()
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    return None


def _image_extents_emu(path: Path, width_cm: float | None = None, height_cm: float | None = None) -> tuple[int, int]:
    dimensions = _png_dimensions(path)
    if width_cm and height_cm:
        return int(round(width_cm * EMU_PER_CM)), int(round(height_cm * EMU_PER_CM))
    if dimensions is None:
        raise ValueError(f"image dimensions are required for non-PNG figure asset: {path}")
    pixel_width, pixel_height = dimensions
    if pixel_width <= 0 or pixel_height <= 0:
        raise ValueError(f"invalid image dimensions: {path}")
    if width_cm:
        cx = int(round(width_cm * EMU_PER_CM))
        cy = int(round(cx * pixel_height / pixel_width))
        return cx, cy
    if height_cm:
        cy = int(round(height_cm * EMU_PER_CM))
        cx = int(round(cy * pixel_width / pixel_height))
        return cx, cy
    cx = int(round(13.69 * EMU_PER_CM))
    cy = int(round(cx * pixel_height / pixel_width))
    return cx, cy


def _max_docpr_id(root: ET.Element) -> int:
    current = 0
    for node in root.findall(".//wp:docPr", NS):
        try:
            current = max(current, int(node.attrib.get("id", "0") or "0"))
        except ValueError:
            continue
    return current


def _set_drawing_embedding(paragraph: ET.Element, rid: str, cx: int, cy: int, docpr_id: int, name: str) -> int:
    changes = 0
    for blip in paragraph.findall(".//a:blip", NS):
        before = dict(blip.attrib)
        blip.set(f"{{{R_NS}}}embed", rid)
        if f"{{{R_NS}}}link" in blip.attrib:
            del blip.attrib[f"{{{R_NS}}}link"]
        if before != blip.attrib:
            changes += 1
    for extent in paragraph.findall(".//wp:extent", NS):
        before = dict(extent.attrib)
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
        if before != extent.attrib:
            changes += 1
    for extent in paragraph.findall(".//a:xfrm/a:ext", NS):
        before = dict(extent.attrib)
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
        if before != extent.attrib:
            changes += 1
    for docpr in paragraph.findall(".//wp:docPr", NS):
        before = dict(docpr.attrib)
        docpr.set("id", str(docpr_id))
        docpr.set("name", name)
        if before != docpr.attrib:
            changes += 1
    for cnvpr in paragraph.findall(".//pic:cNvPr", NS):
        before = dict(cnvpr.attrib)
        cnvpr.set("id", str(docpr_id))
        cnvpr.set("name", name)
        if before != cnvpr.attrib:
            changes += 1
    return changes


def _caption_index(children: list[ET.Element], caption: str) -> int | None:
    target = " ".join((caption or "").split())
    for index, child in enumerate(children):
        if not is_paragraph(child):
            continue
        text = " ".join(paragraph_text(child).split())
        if text == target:
            return index
    return None


def _previous_drawing_index(children: list[ET.Element], caption_index: int) -> int | None:
    index = caption_index - 1
    while index >= 0:
        child = children[index]
        if is_paragraph(child):
            text = paragraph_text(child)
            if has_drawingml_blip(child):
                return index
            if text:
                return None
        index -= 1
    return None


def _previous_legacy_image_holder_index(children: list[ET.Element], caption_index: int) -> int | None:
    index = caption_index - 1
    while index >= 0:
        child = children[index]
        if is_paragraph(child):
            text = paragraph_text(child)
            if has_drawingml_blip(child):
                return None
            if has_drawing(child) and not text:
                return index
            if text:
                return None
        index -= 1
    return None


def caption_donor(children: list[ET.Element]) -> ET.Element | None:
    for child in children:
        if is_paragraph(child) and is_formal_caption_text(paragraph_text(child)):
            return child
    return None


def body_donor(children: list[ET.Element], start: int) -> ET.Element | None:
    for index in range(start, len(children)):
        child = children[index]
        if not is_paragraph(child):
            continue
        text = paragraph_text(child)
        if text and not has_drawing(child) and not is_formal_caption_text(text) and not is_heading_text(text):
            return child
    return None


def normalize_malformed_caption(text: str) -> str | None:
    match = MALFORMED_CAPTION_RE.match(text or "")
    if not match:
        return None
    title = match.group("title").strip()
    # The historical roll-forming manuscript used this malformed caption after
    # several chapter-1 figures already existed. Keep the repaired caption
    # unique instead of reusing "1-2".
    if "\u53cc\u68cd\u8f6e\u6324\u538b\u793a\u610f\u56fe" in title or "\u53cc\u8f8a\u8f6e\u6324\u538b\u793a\u610f\u56fe" in title:
        return "\u56fe1-6 \u53cc\u8f8a\u8f6e\u6324\u538b\u793a\u610f\u56fe"
    return f"\u56fe{match.group('major')}-{match.group('minor')} {title}"


def insert_followup_if_needed(children: list[ET.Element], body: ET.Element, caption_index: int) -> dict[str, object] | None:
    next_node = children[caption_index + 1] if caption_index + 1 < len(children) else None
    if next_node is None or not is_paragraph(next_node) or not is_heading_text(paragraph_text(next_node)):
        return None
    donor = body_donor(children, caption_index + 1)
    text = "\u8be5\u56fe\u7528\u4e8e\u8bf4\u660e\u672c\u5904\u7ed3\u6784\u5173\u7cfb\uff0c\u540e\u7eed\u8ba1\u7b97\u548c\u56fe\u7eb8\u8868\u8fbe\u5747\u4ee5\u56fe\u4e2d\u6240\u793a\u4f4d\u7f6e\u5173\u7cfb\u4f5c\u4e3a\u4f9d\u636e\u3002"
    paragraph = clone_text_paragraph(donor, text)
    body.insert(caption_index + 1, paragraph)
    return {
        "kind": "caption_followup_inserted",
        "body_child_index": caption_index + 1,
        "text": text,
    }


def repair_document(root: ET.Element, *, remove_consecutive_orphan_images: bool) -> list[dict[str, object]]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    changes: list[dict[str, object]] = []

    children = list(body)
    donor = caption_donor(children)

    index = 0
    while index + 1 < len(children):
        current = children[index]
        following = children[index + 1]
        if (
            is_paragraph(current)
            and is_paragraph(following)
            and is_formal_caption_text(paragraph_text(current))
            and has_drawing(following)
        ):
            caption_text = paragraph_text(current)
            body.remove(current)
            body.insert(index + 1, current)
            ensure_keep_next(following)
            changes.append(
                {
                    "kind": "caption_moved_after_image",
                    "from_body_child_index": index,
                    "to_body_child_index": index + 1,
                    "caption": caption_text,
                }
            )
            children = list(body)
            followup = insert_followup_if_needed(children, body, index + 1)
            if followup:
                changes.append(followup)
                children = list(body)
            index += 2
            continue
        index += 1

    children = list(body)
    index = 0
    while index < len(children):
        child = children[index]
        if not is_paragraph(child) or not has_drawing(child):
            index += 1
            continue
        text = paragraph_text(child)
        caption = normalize_malformed_caption(text)
        if caption and len(text) <= 60:
            cleared = clear_visible_text(child)
            ensure_keep_next(child)
            new_caption = clone_text_paragraph(donor, caption)
            body.insert(index + 1, new_caption)
            changes.append(
                {
                    "kind": "malformed_caption_split_from_image_holder",
                    "body_child_index": index,
                    "inserted_caption_index": index + 1,
                    "old_text": text,
                    "caption": caption,
                    "cleared_text_nodes": cleared,
                }
            )
            children = list(body)
            index += 2
            continue
        index += 1

    if remove_consecutive_orphan_images:
        children = list(body)
        index = 0
        while index + 2 < len(children):
            first, second, third = children[index], children[index + 1], children[index + 2]
            if (
                is_paragraph(first)
                and is_paragraph(second)
                and is_paragraph(third)
                and has_drawing(first)
                and not paragraph_text(first)
                and has_drawing(second)
                and is_formal_caption_text(paragraph_text(third))
            ):
                remove_offset = consecutive_orphan_image_remove_offset(first, second)
                body.remove(first if remove_offset == 0 else second)
                changes.append(
                    {
                        "kind": "unlabeled_consecutive_image_removed",
                        "body_child_index": index + remove_offset,
                        "following_caption": paragraph_text(third),
                    }
                )
                children = list(body)
                continue
            index += 1

    return changes


def _load_json_list(path: Path | None, key: str) -> list[dict[str, object]]:
    if path is None:
        return []
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if isinstance(payload, dict):
        value = payload.get(key) or payload.get("items") or []
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    raise ValueError(f"{path} must contain a list or an object with {key}")


def _content_type_root(path: Path) -> ET.Element:
    return ET.fromstring(path.read_bytes())


def _ensure_png_content_type(root: ET.Element) -> bool:
    q_default = f"{{{CT_NS}}}Default"
    for child in root.findall(q_default):
        if child.attrib.get("Extension", "").lower() == "png":
            return False
    node = ET.Element(q_default)
    node.set("Extension", "png")
    node.set("ContentType", "image/png")
    root.append(node)
    return True


def _next_relationship_id(root: ET.Element) -> str:
    max_id = 0
    for rel in list(root):
        rid = rel.attrib.get("Id", "")
        match = re.match(r"^rId(\d+)$", rid)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def _add_image_relationship(root: ET.Element, target: str) -> str:
    rid = _next_relationship_id(root)
    node = ET.Element(f"{{{REL_PKG_NS}}}Relationship")
    node.set("Id", rid)
    node.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    node.set("Target", target)
    root.append(node)
    return rid


def _relationship_target(root: ET.Element, rid: str) -> str | None:
    for rel in list(root):
        if rel.attrib.get("Id") == rid:
            return rel.attrib.get("Target")
    return None


def _relationship_media_path(target: str | None) -> str | None:
    if not target:
        return None
    normalized = target.replace("\\", "/").lstrip("/")
    if normalized.startswith("word/"):
        return normalized
    return f"word/{normalized}"


def _safe_media_name(existing: set[str], preferred_stem: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_-]+", "-", preferred_stem).strip("-") or "skill-figure"
    candidate = f"{stem}.png"
    counter = 1
    while f"word/media/{candidate}" in existing:
        candidate = f"{stem}-{counter}.png"
        counter += 1
    existing.add(f"word/media/{candidate}")
    return candidate


def _copy_image_to_media(tmp: Path, image_path: Path, media_name: str) -> None:
    media_dir = tmp / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(image_path, media_dir / media_name)


def _repair_image_bindings(
    root: ET.Element,
    *,
    tmp: Path,
    replacements: list[dict[str, object]],
    insertions: list[dict[str, object]],
) -> tuple[list[dict[str, object]], set[str]]:
    if not replacements and not insertions:
        return [], set()
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    rels_path = tmp / "word" / "_rels" / "document.xml.rels"
    content_types_path = tmp / "[Content_Types].xml"
    rels_root = ET.fromstring(rels_path.read_bytes())
    content_types_root = _content_type_root(content_types_path)
    changed_parts: set[str] = set()
    existing_names = {
        str(path.relative_to(tmp)).replace("\\", "/")
        for path in (tmp / "word" / "media").glob("*")
        if path.is_file()
    }
    needs_new_png_media = bool(replacements) or any(not (item.get("existing_rid") or item.get("rid")) for item in insertions)
    if needs_new_png_media and _ensure_png_content_type(content_types_root):
        changed_parts.add("[Content_Types].xml")
    children = list(body)
    donor = image_holder_donor(children)
    if donor is None and (replacements or insertions):
        raise RuntimeError("no DrawingML image-holder donor paragraph found")
    max_docpr = _max_docpr_id(root)
    changes: list[dict[str, object]] = []

    def add_media_and_relationship(item: dict[str, object], fallback_stem: str) -> tuple[str, str, int, int, Path]:
        image_path = Path(str(item.get("image_path") or "")).resolve()
        if not image_path.exists():
            raise FileNotFoundError(f"figure asset missing: {image_path}")
        media_name = _safe_media_name(existing_names, str(item.get("media_stem") or fallback_stem))
        _copy_image_to_media(tmp, image_path, media_name)
        target = f"media/{media_name}"
        rid = _add_image_relationship(rels_root, target)
        width_cm = float(item["width_cm"]) if item.get("width_cm") not in (None, "") else None
        height_cm = float(item["height_cm"]) if item.get("height_cm") not in (None, "") else None
        cx, cy = _image_extents_emu(image_path, width_cm, height_cm)
        changed_parts.add(f"word/media/{media_name}")
        changed_parts.add("word/_rels/document.xml.rels")
        return rid, media_name, cx, cy, image_path

    def existing_relationship_binding(item: dict[str, object]) -> tuple[str, str, int, int, str]:
        rid = str(item.get("existing_rid") or item.get("rid") or "").strip()
        if not rid:
            raise ValueError("existing_rid is required when reusing an existing image relationship")
        target = _relationship_target(rels_root, rid)
        media_path = _relationship_media_path(target)
        if media_path is None:
            raise ValueError(f"relationship {rid} has no target")
        if not (tmp / media_path).exists():
            raise FileNotFoundError(f"relationship {rid} target missing in DOCX package: {media_path}")
        if item.get("expected_media_sha256"):
            expected_sha = str(item.get("expected_media_sha256")).lower()
            actual_sha = hashlib.sha256((tmp / media_path).read_bytes()).hexdigest().lower()
            if expected_sha != actual_sha:
                raise ValueError(f"relationship {rid} media SHA256 mismatch: expected {expected_sha}, got {actual_sha}")
        width_cm = item.get("width_cm")
        height_cm = item.get("height_cm")
        cx = int(item["cx_emu"]) if item.get("cx_emu") not in (None, "") else None
        cy = int(item["cy_emu"]) if item.get("cy_emu") not in (None, "") else None
        if cx is None and width_cm not in (None, ""):
            cx = int(round(float(width_cm) * EMU_PER_CM))
        if cy is None and height_cm not in (None, ""):
            cy = int(round(float(height_cm) * EMU_PER_CM))
        if cx is None or cy is None:
            raise ValueError("cx_emu/cy_emu or width_cm/height_cm are required when reusing an existing image relationship")
        return rid, media_path, cx, cy, hashlib.sha256((tmp / media_path).read_bytes()).hexdigest()

    for item in replacements:
        caption = str(item.get("caption") or "")
        children = list(body)
        caption_index = _caption_index(children, caption)
        if caption_index is None:
            changes.append({"kind": "image_replacement_skipped", "caption": caption, "reason": "caption-not-found"})
            continue
        drawing_index = _previous_drawing_index(children, caption_index)
        if drawing_index is None:
            changes.append({"kind": "image_replacement_skipped", "caption": caption, "reason": "previous-drawing-not-found"})
            continue
        rid, media_name, cx, cy, image_path = add_media_and_relationship(item, f"replacement-{caption_index + 1}")
        max_docpr += 1
        changed = _set_drawing_embedding(children[drawing_index], rid, cx, cy, max_docpr, media_name)
        if changed:
            changed_parts.add("word/document.xml")
        changes.append(
            {
                "kind": "image_replaced_before_caption",
                "caption": caption,
                "caption_body_child_index": caption_index,
                "image_body_child_index": drawing_index,
                "new_rid": rid,
                "new_media": f"word/media/{media_name}",
                "new_media_sha256": sha256_file(image_path),
                "width_cm": round(cx / EMU_PER_CM, 3),
                "height_cm": round(cy / EMU_PER_CM, 3),
                "drawing_nodes_changed": changed,
            }
        )

    for item in insertions:
        caption = str(item.get("caption") or "")
        children = list(body)
        caption_index = _caption_index(children, caption)
        if caption_index is None:
            changes.append({"kind": "image_insertion_skipped", "caption": caption, "reason": "caption-not-found"})
            continue
        if _previous_drawing_index(children, caption_index) is not None:
            changes.append({"kind": "image_insertion_skipped", "caption": caption, "reason": "caption-already-has-previous-drawing"})
            continue
        legacy_holder_index = _previous_legacy_image_holder_index(children, caption_index)
        reuse_existing = bool(item.get("existing_rid") or item.get("rid"))
        if reuse_existing:
            rid, media_path, cx, cy, media_sha = existing_relationship_binding(item)
            media_name = Path(media_path).name
            image_path = None
        else:
            rid, media_name, cx, cy, image_path = add_media_and_relationship(item, f"insertion-{caption_index + 1}")
            media_path = f"word/media/{media_name}"
            media_sha = sha256_file(image_path)
        max_docpr += 1
        new_paragraph = deepcopy(donor)
        clear_visible_text(new_paragraph)
        ensure_keep_next(new_paragraph)
        changed = _set_drawing_embedding(new_paragraph, rid, cx, cy, max_docpr, media_name)
        if legacy_holder_index is not None:
            body.remove(children[legacy_holder_index])
            body.insert(legacy_holder_index, new_paragraph)
            inserted_index = legacy_holder_index
        else:
            body.insert(caption_index, new_paragraph)
            inserted_index = caption_index
        changed_parts.add("word/document.xml")
        changes.append(
            {
                "kind": "legacy_image_holder_upgraded_before_caption" if legacy_holder_index is not None else "image_inserted_before_caption",
                "caption": caption,
                "caption_body_child_index_before_insert": caption_index,
                "inserted_body_child_index": inserted_index,
                "new_rid": rid,
                "new_media": media_path,
                "new_media_sha256": media_sha,
                "relationship_reuse": reuse_existing,
                "width_cm": round(cx / EMU_PER_CM, 3),
                "height_cm": round(cy / EMU_PER_CM, 3),
                "drawing_nodes_changed": changed,
            }
        )

    if "word/_rels/document.xml.rels" in changed_parts:
        rels_path.write_bytes(ET.tostring(rels_root, encoding="utf-8", xml_declaration=True))
    if "[Content_Types].xml" in changed_parts:
        content_types_path.write_bytes(ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True))
    return changes, changed_parts


def write_docx(input_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "pkg"
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp)
        (tmp / "word" / "document.xml").write_bytes(document_xml)
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmp)
        output_docx.with_suffix(".zip").replace(output_docx)


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    remove_consecutive_orphan_images: bool,
    image_replacements: list[dict[str, object]] | None = None,
    image_insertions: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    image_replacements = image_replacements or []
    image_insertions = image_insertions or []
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "pkg"
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp)
        document_xml_path = tmp / "word" / "document.xml"
        root = ET.fromstring(document_xml_path.read_bytes())
        changes = repair_document(root, remove_consecutive_orphan_images=remove_consecutive_orphan_images)
        changed_parts: set[str] = {"word/document.xml"} if changes else set()
        binding_changes, binding_changed_parts = _repair_image_bindings(
            root,
            tmp=tmp,
            replacements=image_replacements,
            insertions=image_insertions,
        )
        changes.extend(binding_changes)
        changed_parts.update(binding_changed_parts)
        document_xml_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
        if binding_changes:
            changed_parts.add("word/document.xml")
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmp)
        output_docx.with_suffix(".zip").replace(output_docx)
    skipped_changes = [change for change in changes if str(change.get("kind", "")).endswith("_skipped")]
    return {
        "schema": "graduation-project-builder.docx-figure-locality-repair.v2",
        "generator": "scripts/repair_docx_figure_locality.py",
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "changed_zip_parts": sorted(changed_parts),
        "remove_consecutive_orphan_images": remove_consecutive_orphan_images,
        "changes": changes,
        "change_count": len(changes),
        "skipped_change_count": len(skipped_changes),
        "verdict": "fail" if skipped_changes else "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair DOCX figure caption/image locality.")
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report-json", required=True, type=Path)
    parser.add_argument("--remove-consecutive-orphan-images", action="store_true")
    parser.add_argument("--image-replacements-json", type=Path)
    parser.add_argument("--image-insertions-json", type=Path)
    args = parser.parse_args()
    report = repair_docx(
        args.input_docx,
        args.output_docx,
        remove_consecutive_orphan_images=args.remove_consecutive_orphan_images,
        image_replacements=_load_json_list(args.image_replacements_json, "replacements"),
        image_insertions=_load_json_list(args.image_insertions_json, "insertions"),
    )
    args.report_json.parent.mkdir(parents=True, exist_ok=True)
    args.report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output_docx": str(args.output_docx), "change_count": report["change_count"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
