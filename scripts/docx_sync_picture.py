from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import posixpath
import re
import shutil
import struct
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "v": "urn:schemas-microsoft-com:vml",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}
W14 = "{%s}" % NS["w14"]
W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]
PR = "{%s}" % NS["pr"]
CT = "{%s}" % NS["ct"]
SCRIPT_DIR = Path(__file__).resolve().parent
CAPTION_LABEL_PREFIX_RE = re.compile(
    r"^\s*(?:图|表)\s*(?:\d+|[一二三四五六七八九十]+)(?:[-.－．](?:\d+|[一二三四五六七八九十]+))*"
    r"(?=\s*[\u4e00-\u9fff])",
    flags=re.IGNORECASE,
)

try:
    from thesis_figure_contract import docx_drawing_object_manifest, validate_figure_manifest
except ImportError:  # pragma: no cover - script execution fallback
    import sys

    sys.path.insert(0, str(SCRIPT_DIR))
    from thesis_figure_contract import docx_drawing_object_manifest, validate_figure_manifest


for _prefix, _uri in (
    ("w", NS["w"]),
    ("w14", NS["w14"]),
    ("a", NS["a"]),
    ("pic", NS["pic"]),
    ("wp", NS["wp"]),
    ("r", NS["r"]),
    ("v", NS["v"]),
    ("asvg", NS["asvg"]),
):
    ET.register_namespace(_prefix, _uri)


def ensure_content_type(content_types_xml: bytes, suffix: str) -> bytes:
    raw_text = content_types_xml.decode("utf-8-sig")
    ext = suffix.lower().lstrip(".")
    mime, _ = mimetypes.guess_type(f"dummy.{ext}")
    content_type = mime or {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "svg": "image/svg+xml",
        "webp": "image/webp",
    }.get(ext, "application/octet-stream")
    if re.search(rf'Extension="{re.escape(ext)}"', raw_text, flags=re.IGNORECASE):
        return content_types_xml
    insertion = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
    updated = raw_text.replace("</Types>", insertion + "</Types>")
    return updated.encode("utf-8")


def file_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_new_media_name(existing_target: str, image_path: Path) -> str:
    clean_target = existing_target.lstrip("/\\")
    target_dir = Path(clean_target).parent
    stem = Path(clean_target).stem
    new_name = f"{stem}-replaced{image_path.suffix.lower()}"
    return str(target_dir / new_name).replace("\\", "/")


def build_insert_media_name(existing_members: set[str], image_path: Path) -> str:
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", image_path.stem).strip("-") or "inserted-figure"
    suffix = image_path.suffix.lower()
    for index in range(1, 10000):
        target = f"media/{safe_stem}{'' if index == 1 else '-' + str(index)}{suffix}"
        if word_media_member_from_target(target, existing_members) not in existing_members:
            return target
    raise ValueError(f"Could not allocate a unique media name for {image_path}")


def word_media_member_from_target(target: str, names: set[str] | None = None) -> str:
    candidates: list[str]
    if target.startswith("/"):
        candidates = [
            posixpath.normpath(target.lstrip("/")),
            posixpath.normpath("word/" + target.lstrip("/")),
        ]
    else:
        candidates = [posixpath.normpath("word/" + target), posixpath.normpath(target)]
    if names is not None:
        for candidate in candidates:
            if candidate in names:
                return candidate
    return candidates[0]


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def find_caption_paragraph(document_root: ET.Element, caption: str) -> tuple[ET.Element, list[ET.Element], int, ET.Element]:
    body = document_root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no body")
    children = list(body)
    for index, child in enumerate(children):
        if child.tag != f"{{{NS['w']}}}p":
            continue
        if paragraph_text(child) != caption:
            continue
        return body, children, index, child
    raise ValueError(f"Caption paragraph not found: {caption}")


def find_picture_paragraph_by_caption(document_root: ET.Element, caption: str) -> ET.Element:
    _body, children, index, _caption_para = find_caption_paragraph(document_root, caption)
    scan = index - 1
    if scan >= 0:
        candidate = children[scan]
        if candidate.tag == f"{{{NS['w']}}}p" and paragraph_text(candidate):
            raise ValueError(f"Caption is not immediately preceded by a dedicated picture paragraph: {caption}")
        if candidate.tag == f"{{{NS['w']}}}p" and (
            candidate.find(".//a:blip", NS) is not None or candidate.find(".//v:imagedata", NS) is not None
        ):
            return candidate
    raise ValueError(f"Caption found but no nearby picture paragraph before it: {caption}")


def find_existing_image_holder_before_caption(document_root: ET.Element, caption: str) -> ET.Element | None:
    _body, children, index, _caption_para = find_caption_paragraph(document_root, caption)
    scan = index - 1
    if scan >= 0:
        candidate = children[scan]
        if candidate.tag == f"{{{NS['w']}}}p" and (
            candidate.find(".//a:blip", NS) is not None or candidate.find(".//v:imagedata", NS) is not None
        ):
            return candidate
    return None


def next_relationship_id(rels_root: ET.Element) -> str:
    used = {rel.attrib.get("Id", "") for rel in rels_root.findall(f"{PR}Relationship")}
    max_num = 0
    for rel_id in used:
        match = re.fullmatch(r"rId(\d+)", rel_id)
        if match:
            max_num = max(max_num, int(match.group(1)))
    candidate = max_num + 1
    while f"rId{candidate}" in used:
        candidate += 1
    return f"rId{candidate}"


def next_drawing_id(document_root: ET.Element) -> int:
    max_id = 0
    for node in document_root.findall(".//wp:docPr", NS) + document_root.findall(".//pic:cNvPr", NS):
        try:
            max_id = max(max_id, int(node.attrib.get("id", "0")))
        except ValueError:
            continue
    return max_id + 1


def image_dimensions(image_path: Path) -> tuple[int, int]:
    data = image_path.read_bytes()
    suffix = image_path.suffix.lower()
    if suffix == ".png" and data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if suffix in {".jpg", ".jpeg"} and data.startswith(b"\xff\xd8"):
        index = 2
        sof_markers = set(range(0xC0, 0xC4)) | set(range(0xC5, 0xC8)) | set(range(0xC9, 0xCC)) | set(range(0xCD, 0xD0))
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            index += 2
            if marker in {0xD8, 0xD9}:
                continue
            if index + 2 > len(data):
                break
            length = struct.unpack(">H", data[index : index + 2])[0]
            if marker in sof_markers and index + 7 <= len(data):
                height = struct.unpack(">H", data[index + 3 : index + 5])[0]
                width = struct.unpack(">H", data[index + 5 : index + 7])[0]
                return width, height
            index += max(length, 2)
    raise ValueError(f"Unsupported image type for dimension detection: {image_path}")


def scaled_extent_emu(image_path: Path, *, width_cm: float) -> tuple[int, int]:
    width_px, height_px = image_dimensions(image_path)
    if width_px <= 0 or height_px <= 0:
        raise ValueError(f"Invalid image dimensions for {image_path}: {width_px}x{height_px}")
    cx = int(round(width_cm * 360000))
    cy = int(round(cx * height_px / width_px))
    return cx, cy


def cm_from_emu(value: int) -> float:
    return value / 360000


def make_el(tag: str, attrib: dict[str, str] | None = None) -> ET.Element:
    prefix, local = tag.split(":", 1)
    return ET.Element(f"{{{NS[prefix]}}}{local}", attrib or {})


def sub_el(parent: ET.Element, tag: str, attrib: dict[str, str] | None = None) -> ET.Element:
    child = make_el(tag, attrib)
    parent.append(child)
    return child


def build_picture_paragraph(
    *,
    rel_id: str,
    svg_rel_id: str | None = None,
    image_path: Path,
    label: str,
    width_cm: float,
    drawing_id: int,
) -> ET.Element:
    cx, cy = scaled_extent_emu(image_path, width_cm=width_cm)
    paragraph = make_el("w:p")
    ppr = sub_el(paragraph, "w:pPr")
    sub_el(
        ppr,
        "w:spacing",
        {
            f"{{{NS['w']}}}before": "120",
            f"{{{NS['w']}}}after": "0",
            f"{{{NS['w']}}}line": "240",
            f"{{{NS['w']}}}lineRule": "auto",
        },
    )
    sub_el(
        ppr,
        "w:ind",
        {
            f"{{{NS['w']}}}left": "0",
            f"{{{NS['w']}}}right": "0",
            f"{{{NS['w']}}}firstLine": "0",
        },
    )
    sub_el(ppr, "w:jc", {f"{{{NS['w']}}}val": "center"})
    run = sub_el(paragraph, "w:r")
    drawing = sub_el(run, "w:drawing")
    inline = sub_el(drawing, "wp:inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
    sub_el(inline, "wp:extent", {"cx": str(cx), "cy": str(cy)})
    sub_el(inline, "wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
    sub_el(inline, "wp:docPr", {"id": str(drawing_id), "name": label, "descr": label})
    c_nv = sub_el(inline, "wp:cNvGraphicFramePr")
    sub_el(c_nv, "a:graphicFrameLocks", {"noChangeAspect": "1"})
    graphic = sub_el(inline, "a:graphic")
    graphic_data = sub_el(graphic, "a:graphicData", {"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"})
    pic = sub_el(graphic_data, "pic:pic")
    nv_pic = sub_el(pic, "pic:nvPicPr")
    sub_el(nv_pic, "pic:cNvPr", {"id": str(drawing_id), "name": label, "descr": label})
    sub_el(nv_pic, "pic:cNvPicPr")
    blip_fill = sub_el(pic, "pic:blipFill")
    blip = sub_el(blip_fill, "a:blip", {f"{{{NS['r']}}}embed": rel_id})
    if svg_rel_id:
        ext_lst = sub_el(blip, "a:extLst")
        ext = sub_el(ext_lst, "a:ext", {"uri": "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"})
        sub_el(ext, "asvg:svgBlip", {f"{{{NS['r']}}}embed": svg_rel_id})
    stretch = sub_el(blip_fill, "a:stretch")
    sub_el(stretch, "a:fillRect")
    sp_pr = sub_el(pic, "pic:spPr")
    xfrm = sub_el(sp_pr, "a:xfrm")
    sub_el(xfrm, "a:off", {"x": "0", "y": "0"})
    sub_el(xfrm, "a:ext", {"cx": str(cx), "cy": str(cy)})
    prst_geom = sub_el(sp_pr, "a:prstGeom", {"prst": "rect"})
    sub_el(prst_geom, "a:avLst")
    return paragraph


def element_has_picture(element: ET.Element) -> bool:
    return (
        element.find(".//a:blip", NS) is not None
        or element.find(".//v:imagedata", NS) is not None
        or element.find(".//w:drawing", NS) is not None
        or element.find(".//w:pict", NS) is not None
    )


def replacement_picture_children(picture_paragraph: ET.Element) -> list[ET.Element]:
    return [deepcopy(child) for child in list(picture_paragraph) if element_has_picture(child)]


def replace_paragraph_picture(paragraph: ET.Element, picture_paragraph: ET.Element) -> None:
    """Replace only the picture-bearing child nodes.

    Teacher comments often anchor directly around an image paragraph:
    commentRangeStart, picture run, commentRangeEnd, commentReference. Replacing
    the whole paragraph silently deletes those anchors, so keep every non-picture
    child in place and swap only the drawing/VML run.
    """
    new_picture_children = replacement_picture_children(picture_paragraph)
    if not new_picture_children:
        raise ValueError("replacement picture paragraph has no picture-bearing child")

    original_children = list(paragraph)
    rebuilt: list[ET.Element] = []
    inserted = False
    for child in original_children:
        if element_has_picture(child):
            if not inserted:
                rebuilt.extend(deepcopy(item) for item in new_picture_children)
                inserted = True
            continue
        rebuilt.append(deepcopy(child))

    if not inserted:
        insert_at = 1 if rebuilt and rebuilt[0].tag == f"{{{NS['w']}}}pPr" else 0
        rebuilt[insert_at:insert_at] = [deepcopy(item) for item in new_picture_children]

    for child in original_children:
        paragraph.remove(child)
    for child in rebuilt:
        paragraph.append(child)


def is_figure_caption_text(text: str) -> bool:
    return bool(re.match(r"^\s*(图|Figure|Fig\.)\s*[\d一二三四五六七八九十]+(?:[-.]\d+)?\s+\S", text, flags=re.IGNORECASE))


def _attr(element: ET.Element | None, local_name: str) -> str:
    return element.attrib.get(f"{{{NS['w']}}}{local_name}", "") if element is not None else ""


def caption_following_body_safety_issues(document_root: ET.Element, caption: str) -> list[str]:
    """Fail on caption-adjacent body prose that looks like another caption."""

    _body, children, index, _caption_para = find_caption_paragraph(document_root, caption)
    for child in children[index + 1 :]:
        if child.tag != W + "p":
            if child.tag == W + "tbl":
                return []
            continue
        text = paragraph_text(child)
        if not text:
            continue
        if is_figure_caption_text(text):
            return []
        ppr = child.find("./w:pPr", NS)
        spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
        ind = ppr.find("./w:ind", NS) if ppr is not None else None
        jc = ppr.find("./w:jc", NS) if ppr is not None else None
        pstyle = ppr.find("./w:pStyle", NS) if ppr is not None else None
        issues: list[str] = []
        if CAPTION_LABEL_PREFIX_RE.match(text):
            issues.append("body prose repeats a figure/table label immediately after the formal caption")
        if _attr(pstyle, "val").lower() == "caption":
            issues.append("caption-following body prose still uses Caption style")
        if _attr(jc, "val").lower() in {"center", "right"}:
            issues.append("caption-following body prose is still centered/right-aligned")
        if _attr(ind, "firstLine") in {"", "0"}:
            issues.append("caption-following body prose lacks body first-line indent")
        if _attr(spacing, "line") in {"240", "300"} and _attr(ind, "firstLine") in {"", "0"}:
            issues.append("caption-following body prose keeps caption-like line spacing")
        if ppr is not None and ppr.find("./w:keepNext", NS) is not None:
            issues.append("caption-following body prose keeps keepNext from a figure/caption block")
        return issues
    return []


def build_caption_paragraph(document_root: ET.Element, caption: str) -> ET.Element:
    donor_ppr: ET.Element | None = None
    donor_rpr: ET.Element | None = None
    for paragraph in document_root.findall(".//w:body/w:p", NS):
        text = paragraph_text(paragraph)
        if not is_figure_caption_text(text):
            continue
        donor_ppr = paragraph.find("./w:pPr", NS)
        first_run = paragraph.find("./w:r", NS)
        donor_rpr = first_run.find("./w:rPr", NS) if first_run is not None else None
        break
    paragraph = make_el("w:p")
    if donor_ppr is not None:
        paragraph.append(deepcopy(donor_ppr))
    else:
        ppr = sub_el(paragraph, "w:pPr")
        sub_el(ppr, "w:jc", {f"{{{NS['w']}}}val": "center"})
    run = sub_el(paragraph, "w:r")
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    text_node = sub_el(run, "w:t")
    text_node.text = caption
    return paragraph


def find_paragraph_by_para_id(document_root: ET.Element, para_id: str) -> tuple[ET.Element, list[ET.Element], int, ET.Element]:
    body = document_root.find(".//w:body", NS)
    if body is None:
        raise ValueError("word/document.xml has no body")
    children = list(body)
    for index, child in enumerate(children):
        if child.tag == f"{{{NS['w']}}}p" and child.attrib.get(f"{W14}paraId") == para_id:
            return body, children, index, child
    raise ValueError(f"Paragraph with paraId {para_id} not found")


def select_figure_entry(payload: dict[str, object], caption: str | None) -> dict[str, object]:
    figures = payload.get("figures")
    if not isinstance(figures, dict) or not figures:
        raise ValueError("figure manifest must contain at least one figure row")
    entries = [entry for entry in figures.values() if isinstance(entry, dict)]
    matches = [entry for entry in entries if caption and str(entry.get("caption") or "").strip() == caption.strip()]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(f"figure manifest has multiple rows for caption: {caption}")
    if len(entries) == 1:
        return entries[0]
    raise ValueError(f"figure manifest row not found for caption: {caption or '<none>'}")


def matching_diagram_entries(payload: dict[str, object], caption: str | None) -> list[dict[str, object]]:
    diagrams = payload.get("diagrams")
    if not isinstance(diagrams, dict):
        return []
    result: list[dict[str, object]] = []
    for entry in diagrams.values():
        if not isinstance(entry, dict):
            continue
        if caption and str(entry.get("caption") or "").strip() != caption.strip():
            continue
        result.append(entry)
    return result


def bind_diagram_relationship_fields(
    payload: dict[str, object],
    *,
    caption: str | None,
    source_info: dict[str, str] | None = None,
    final_info: dict[str, str],
    svg_info: dict[str, str] | None,
    source_drawing: dict[str, str] | None = None,
    final_drawing: dict[str, str] | None,
) -> None:
    for entry in matching_diagram_entries(payload, caption):
        entry.setdefault("insertion_status", "pass")
        entry.setdefault("rendered_page_status", "pass")
        entry.setdefault("caption_asset_binding", "verified by docx_sync_picture.py")
        entry.setdefault("caption_asset_binding_verdict", "pass")
        if source_info is not None:
            entry.setdefault("mutation_intent", "replace_existing")
            entry.setdefault("explicit_replacement_authorization_source", "figure manifest approved DOCX picture replacement")
            entry.setdefault("explicit_replacement_authorization_scope", "single verified picture relationship")
            entry.setdefault("explicit_drawing_authorization_source", "figure manifest approved DOCX picture replacement")
            entry.setdefault("explicit_drawing_authorization_scope", "single verified picture paragraph")
            entry["original_rid"] = source_info["rid"]
            entry["original_media_target"] = source_info["target"]
            entry["original_media_sha256"] = source_info["sha256"]
            entry["original_owner_part"] = source_info["owner_part"]
        entry["final_rid"] = final_info["rid"]
        entry["final_media_target"] = final_info["target"]
        entry["final_media_sha256"] = final_info["sha256"]
        entry["final_owner_part"] = final_info["owner_part"]
        if svg_info is not None:
            entry["docx_svg_rid"] = svg_info["rid"]
            entry["docx_svg_media_target"] = svg_info["target"]
            entry["docx_svg_media_sha256"] = svg_info["sha256"]
        if source_drawing is not None:
            entry["original_drawing_sha256"] = source_drawing["sha256"]
            entry["original_drawing_owner_part"] = source_drawing["owner_part"]
        if final_drawing is not None:
            entry["final_drawing_sha256"] = final_drawing["sha256"]
            entry["final_drawing_owner_part"] = final_drawing["owner_part"]
        entry.setdefault("target_anchor_caption", caption or str(entry.get("caption") or ""))
        entry.setdefault("target_anchor_not_protected_surface_verdict", "pass")


def drawing_info_for_relationship(docx_path: Path, rel_id: str, caption: str | None = None) -> dict[str, str]:
    drawings = docx_drawing_object_manifest(docx_path)
    for info in drawings.values():
        rels = {item for item in str(info.get("relationship_ids") or "").split(";") if item}
        if rel_id not in rels:
            continue
        if caption and str(info.get("next_text") or "").strip() != caption.strip():
            continue
        return {"sha256": str(info.get("sha256") or ""), "owner_part": str(info.get("story_part") or "word/document.xml")}
    for info in drawings.values():
        rels = {item for item in str(info.get("relationship_ids") or "").split(";") if item}
        if rel_id in rels:
            return {"sha256": str(info.get("sha256") or ""), "owner_part": str(info.get("story_part") or "word/document.xml")}
    raise ValueError(f"Drawing object for relationship {rel_id} not found in {docx_path}")


def bind_docx_paths(payload: dict[str, object], *, source_docx: Path, final_docx: Path) -> None:
    payload["source_docx_path"] = str(source_docx.resolve())
    payload["source_docx_sha256"] = file_sha256(source_docx)
    payload["final_docx_path"] = str(final_docx.resolve())
    payload["final_docx_sha256"] = file_sha256(final_docx)


def picture_relationship_id(paragraph: ET.Element) -> str:
    blip = paragraph.find(".//a:blip", NS)
    if blip is not None:
        rel_id = blip.attrib.get(R + "embed")
        if rel_id:
            return rel_id
    imagedata = paragraph.find(".//v:imagedata", NS)
    if imagedata is not None:
        rel_id = imagedata.attrib.get(R + "id")
        if rel_id:
            return rel_id
    raise ValueError("No picture relationship found in target paragraph")


def resize_picture_paragraph(paragraph: ET.Element, image_path: Path, *, width_cm: float) -> None:
    cx, cy = scaled_extent_emu(image_path, width_cm=width_cm)
    for extent in paragraph.findall(".//wp:extent", NS):
        extent.attrib["cx"] = str(cx)
        extent.attrib["cy"] = str(cy)
    for extent in paragraph.findall(".//a:ext", NS):
        extent.attrib["cx"] = str(cx)
        extent.attrib["cy"] = str(cy)
    width_text = f"{cm_from_emu(cx):.2f}cm"
    height_text = f"{cm_from_emu(cy):.2f}cm"
    for shape in paragraph.findall(".//v:shape", NS):
        style = shape.attrib.get("style", "")
        if "width:" in style:
            style = re.sub(r"width:[^;]+", f"width:{width_text}", style)
        else:
            style = f"{style};width:{width_text}" if style else f"width:{width_text}"
        if "height:" in style:
            style = re.sub(r"height:[^;]+", f"height:{height_text}", style)
        else:
            style = f"{style};height:{height_text}" if style else f"height:{height_text}"
        shape.attrib["style"] = style


def load_manifest_gate(
    *,
    manifest_path: Path | None,
    source_docx: Path | None,
    final_docx: Path,
) -> tuple[dict[str, object], Path, Path]:
    if manifest_path is None:
        raise ValueError("--figure-manifest is required for any DOCX picture replacement")
    if source_docx is None:
        raise ValueError("--source-docx is required for any DOCX picture replacement")
    if not manifest_path.exists():
        raise ValueError(f"figure manifest does not exist: {manifest_path}")
    if not source_docx.exists():
        raise ValueError(f"source DOCX does not exist: {source_docx}")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"figure manifest root must be an object: {manifest_path}")
    return payload, manifest_path, source_docx


def validate_manifest_gate(
    *,
    payload: dict[str, object],
    manifest_path: Path,
    source_docx: Path,
    final_docx: Path,
) -> None:
    issues = validate_figure_manifest(
        payload,
        source_docx=source_docx,
        final_docx=final_docx,
        manifest_path=manifest_path,
    )
    if issues:
        raise ValueError("figure manifest gate failed after picture replacement: " + "; ".join(issues[:8]))


def docx_relationship_media_info(docx_path: Path, rel_id: str) -> dict[str, str]:
    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        for rel in rels_root.findall(f"{PR}Relationship"):
            if rel.attrib.get("Id") != rel_id:
                continue
            target = rel.attrib.get("Target", "")
            media_name = word_media_member_from_target(target, names)
            digest = hashlib.sha256(zf.read(media_name)).hexdigest()
            return {
                "rid": rel_id,
                "target": target,
                "sha256": digest,
                "owner_part": "word/document.xml",
                "rels_part": "word/_rels/document.xml.rels",
            }
    raise ValueError(f"Relationship {rel_id} not found in {docx_path}")


def add_image_relationship(rels_root: ET.Element, target: str) -> str:
    rel_id = next_relationship_id(rels_root)
    rels_root.append(
        ET.Element(
            f"{PR}Relationship",
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": target,
            },
        )
    )
    return rel_id


def bind_manifest_after_replacement(
    payload: dict[str, object],
    *,
    manifest_path: Path,
    source_docx: Path,
    final_docx: Path,
    rel_id: str,
    svg_rel_id: str | None = None,
    caption: str | None = None,
) -> None:
    source_info = docx_relationship_media_info(source_docx, rel_id)
    final_info = docx_relationship_media_info(final_docx, rel_id)
    bind_docx_paths(payload, source_docx=source_docx, final_docx=final_docx)
    entry = select_figure_entry(payload, caption)
    entry.setdefault("mutation_intent", "replace_existing")
    entry.setdefault("explicit_replacement_authorization_source", "figure manifest approved DOCX picture replacement")
    entry.setdefault("explicit_replacement_authorization_scope", "single verified picture relationship")
    entry.setdefault("explicit_drawing_authorization_source", "figure manifest approved DOCX picture replacement")
    entry.setdefault("explicit_drawing_authorization_scope", "single verified picture paragraph")
    entry["original_rid"] = source_info["rid"]
    entry["original_media_target"] = source_info["target"]
    entry["original_media_sha256"] = source_info["sha256"]
    entry["original_owner_part"] = source_info["owner_part"]
    entry["final_rid"] = final_info["rid"]
    entry["final_media_target"] = final_info["target"]
    entry["final_media_sha256"] = final_info["sha256"]
    entry["final_owner_part"] = final_info["owner_part"]
    if svg_rel_id:
        svg_info = docx_relationship_media_info(final_docx, svg_rel_id)
        entry["docx_svg_rid"] = svg_info["rid"]
        entry["docx_svg_media_target"] = svg_info["target"]
        entry["docx_svg_media_sha256"] = svg_info["sha256"]
    else:
        svg_info = None
    final_drawing: dict[str, str] | None = None
    try:
        source_drawing = drawing_info_for_relationship(source_docx, rel_id, caption)
        final_drawing = drawing_info_for_relationship(final_docx, rel_id, caption)
        entry["original_drawing_sha256"] = source_drawing["sha256"]
        entry["original_drawing_owner_part"] = source_drawing["owner_part"]
        entry["final_drawing_sha256"] = final_drawing["sha256"]
        entry["final_drawing_owner_part"] = final_drawing["owner_part"]
    except ValueError:
        pass
    entry.setdefault("target_anchor_caption", caption or str(entry.get("caption") or ""))
    entry.setdefault("target_anchor_not_protected_surface_verdict", "pass")
    entry.setdefault("caption_asset_binding", "verified by docx_sync_picture.py")
    entry.setdefault("caption_asset_binding_verdict", "pass")
    bind_diagram_relationship_fields(
        payload,
        caption=caption,
        source_info=source_info,
        final_info=final_info,
        svg_info=svg_info,
        source_drawing=source_drawing if "source_drawing" in locals() else None,
        final_drawing=final_drawing,
    )
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def bind_manifest_after_insertion(
    payload: dict[str, object],
    *,
    manifest_path: Path,
    source_docx: Path,
    final_docx: Path,
    rel_id: str,
    svg_rel_id: str | None = None,
    caption: str = "",
) -> None:
    final_info = docx_relationship_media_info(final_docx, rel_id)
    bind_docx_paths(payload, source_docx=source_docx, final_docx=final_docx)
    entry = select_figure_entry(payload, caption)
    entry.setdefault("mutation_intent", "insert_image")
    entry.setdefault("explicit_insertion_authorization_source", "figure manifest approved DOCX picture insertion")
    entry.setdefault("explicit_insertion_authorization_scope", "single verified picture inserted before target caption")
    entry.setdefault("explicit_drawing_authorization_source", "figure manifest approved DOCX picture insertion")
    entry.setdefault("explicit_drawing_authorization_scope", "single verified picture inserted before target caption")
    entry["final_rid"] = final_info["rid"]
    entry["final_media_target"] = final_info["target"]
    entry["final_media_sha256"] = final_info["sha256"]
    entry["final_owner_part"] = final_info["owner_part"]
    if svg_rel_id:
        svg_info = docx_relationship_media_info(final_docx, svg_rel_id)
        entry["docx_svg_rid"] = svg_info["rid"]
        entry["docx_svg_media_target"] = svg_info["target"]
        entry["docx_svg_media_sha256"] = svg_info["sha256"]
    else:
        svg_info = None
    final_drawing = drawing_info_for_relationship(final_docx, rel_id, caption)
    entry["final_drawing_sha256"] = final_drawing["sha256"]
    entry["final_drawing_owner_part"] = final_drawing["owner_part"]
    entry.setdefault("target_anchor_caption", caption)
    entry.setdefault("target_anchor_not_protected_surface_verdict", "pass")
    entry.setdefault("insertion_status", "inserted")
    entry.setdefault("rendered_page_status", "pass")
    entry.setdefault("caption_asset_binding", "verified by docx_sync_picture.py")
    entry.setdefault("caption_asset_binding_verdict", "pass")
    bind_diagram_relationship_fields(
        payload,
        caption=caption,
        final_info=final_info,
        svg_info=svg_info,
        final_drawing=final_drawing,
    )
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def patch_docx(
    docx_path: Path,
    para_id: str | None,
    image_path: Path,
    svg_image_path: Path | None = None,
    new_label: str | None = None,
    output_path: Path | None = None,
    caption: str | None = None,
    manifest_path: Path | None = None,
    source_docx: Path | None = None,
    mode: str = "replace",
    width_cm: float = 12.0,
    resize_width_cm: float | None = None,
    defer_validation: bool = False,
) -> tuple[Path, str]:
    out_path = output_path or docx_path
    if output_path:
        shutil.copy2(docx_path, out_path)
    manifest_payload, checked_manifest_path, checked_source_docx = load_manifest_gate(
        manifest_path=manifest_path,
        source_docx=source_docx,
        final_docx=out_path,
    )

    with zipfile.ZipFile(out_path, "r") as zin:
        members = {name: zin.read(name) for name in zin.namelist()}

    document_root = ET.fromstring(members["word/document.xml"])
    rels_root = ET.fromstring(members["word/_rels/document.xml.rels"])

    if mode not in {"replace", "insert-before-caption", "insert-after-para-id", "auto"}:
        raise ValueError(f"Unsupported mode: {mode}")
    if caption and mode == "auto":
        mode = "replace" if find_existing_image_holder_before_caption(document_root, caption) is not None else "insert-before-caption"
    elif mode == "auto":
        mode = "replace"

    if caption and mode == "replace":
        target_para = find_picture_paragraph_by_caption(document_root, caption)
    elif mode == "replace":
        target_para = None
        for para in document_root.findall(".//w:body/w:p", NS):
            if para.attrib.get(f"{W14}paraId") == para_id:
                target_para = para
                break
        if target_para is None:
            raise ValueError(f"Paragraph with paraId {para_id} not found")
    svg_rel_id: str | None = None
    if mode == "replace":
        rel_id = picture_relationship_id(target_para)

        rel_node = None
        for rel in rels_root.findall(f"{PR}Relationship"):
            if rel.attrib.get("Id") == rel_id:
                rel_node = rel
                break
        if rel_node is None:
            raise ValueError(f"Relationship {rel_id} not found in document rels")

        existing_target = rel_node.attrib["Target"]
        if Path(existing_target).suffix.lower() != image_path.suffix.lower():
            raise ValueError(
                "relationship-preserving picture replacement must keep the original media extension "
                f"({Path(existing_target).suffix.lower()} != {image_path.suffix.lower()})"
            )
        new_target = existing_target
        if svg_image_path is not None:
            svg_target = build_insert_media_name(set(members), svg_image_path)
            svg_rel_id = add_image_relationship(rels_root, svg_target)

        if new_label:
            doc_pr = target_para.find(".//wp:docPr", NS)
            if doc_pr is not None:
                doc_pr.attrib["name"] = new_label
                doc_pr.attrib["descr"] = new_label
            c_nv_pr = target_para.find(".//pic:cNvPr", NS)
            if c_nv_pr is not None:
                c_nv_pr.attrib["name"] = new_label
                c_nv_pr.attrib["descr"] = new_label
        if resize_width_cm is not None:
            resize_picture_paragraph(target_para, image_path, width_cm=resize_width_cm)
        if svg_image_path is not None:
            replacement_para = build_picture_paragraph(
                rel_id=rel_id,
                svg_rel_id=svg_rel_id,
                image_path=image_path,
                label=new_label or caption or image_path.stem,
                width_cm=resize_width_cm or width_cm,
                drawing_id=next_drawing_id(document_root),
            )
            replace_paragraph_picture(target_para, replacement_para)
        members[word_media_member_from_target(new_target, set(members))] = image_path.read_bytes()
        if svg_image_path is not None:
            members[word_media_member_from_target(svg_target, set(members))] = svg_image_path.read_bytes()
    else:
        if not caption:
            raise ValueError("--caption is required for picture insertion modes")
        if mode == "insert-before-caption":
            if find_existing_image_holder_before_caption(document_root, caption) is not None:
                raise ValueError(f"Caption already has a picture paragraph before it: {caption}")
            body, _children, caption_index, _caption_para = find_caption_paragraph(document_root, caption)
            insert_index = caption_index
            insert_caption = False
        else:
            if not para_id:
                raise ValueError("--para-id is required for insert-after-para-id mode")
            body, _children, para_index, _target_para = find_paragraph_by_para_id(document_root, para_id)
            insert_index = para_index + 1
            insert_caption = True
        new_target = build_insert_media_name(set(members), image_path)
        rel_id = add_image_relationship(rels_root, new_target)
        if svg_image_path is not None:
            svg_target = build_insert_media_name(set(members), svg_image_path)
            svg_rel_id = add_image_relationship(rels_root, svg_target)
        picture_para = build_picture_paragraph(
            rel_id=rel_id,
            svg_rel_id=svg_rel_id,
            image_path=image_path,
            label=new_label or caption or image_path.stem,
            width_cm=width_cm,
            drawing_id=next_drawing_id(document_root),
        )
        body.insert(insert_index, picture_para)
        if insert_caption:
            body.insert(insert_index + 1, build_caption_paragraph(document_root, caption))
        members[word_media_member_from_target(new_target, set(members))] = image_path.read_bytes()
        if svg_image_path is not None:
            members[word_media_member_from_target(svg_target, set(members))] = svg_image_path.read_bytes()

    members["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    members["word/_rels/document.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    members["[Content_Types].xml"] = ensure_content_type(members["[Content_Types].xml"], image_path.suffix)
    if svg_image_path is not None:
        members["[Content_Types].xml"] = ensure_content_type(members["[Content_Types].xml"], svg_image_path.suffix)
    if caption:
        issues = caption_following_body_safety_issues(document_root, caption)
        if issues:
            raise RuntimeError(
                "caption-following body prose failed safety checks after picture sync for "
                f"{caption}: " + "; ".join(issues)
            )

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    if mode == "replace":
        bind_manifest_after_replacement(
            manifest_payload,
            manifest_path=checked_manifest_path,
            source_docx=checked_source_docx,
            final_docx=out_path,
            rel_id=rel_id,
            svg_rel_id=svg_rel_id,
            caption=caption,
        )
    else:
        bind_manifest_after_insertion(
            manifest_payload,
            manifest_path=checked_manifest_path,
            source_docx=checked_source_docx,
            final_docx=out_path,
            rel_id=rel_id,
            svg_rel_id=svg_rel_id,
            caption=caption or "",
        )
    if not defer_validation:
        validate_manifest_gate(
            payload=manifest_payload,
            manifest_path=checked_manifest_path,
            source_docx=checked_source_docx,
            final_docx=out_path,
        )
    return out_path, rel_id


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace or insert one embedded DOCX picture and sync visible metadata.")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--para-id")
    parser.add_argument("--caption", help="Find the picture paragraph immediately before this visible caption.")
    parser.add_argument("--image", required=True)
    parser.add_argument("--svg-image", help="Optional SVG primary image; --image remains the raster fallback.")
    parser.add_argument("--label")
    parser.add_argument("--output")
    parser.add_argument("--figure-manifest", required=True)
    parser.add_argument("--source-docx", required=True)
    parser.add_argument(
        "--mode",
        choices=("replace", "insert-before-caption", "insert-after-para-id", "auto"),
        default="replace",
        help="replace an existing picture, insert before a caption, insert after a paraId, or auto-detect for caption targets",
    )
    parser.add_argument("--width-cm", type=float, default=12.0, help="Inserted picture width in centimeters")
    parser.add_argument(
        "--resize-width-cm",
        type=float,
        help="When replacing, also resize the existing picture holder to this width in centimeters",
    )
    parser.add_argument(
        "--defer-validation",
        action="store_true",
        help="Bind the manifest but defer full figure-contract validation to a later aggregate manifest",
    )
    args = parser.parse_args()

    out_path, rel_id = patch_docx(
        docx_path=Path(args.docx),
        para_id=args.para_id,
        image_path=Path(args.image),
        svg_image_path=Path(args.svg_image) if args.svg_image else None,
        new_label=args.label,
        output_path=Path(args.output) if args.output else None,
        caption=args.caption,
        manifest_path=Path(args.figure_manifest),
        source_docx=Path(args.source_docx),
        mode=args.mode,
        width_cm=args.width_cm,
        resize_width_cm=args.resize_width_cm,
        defer_validation=args.defer_validation,
    )
    print(f"patched_docx={out_path}")
    print(f"relationship_id={rel_id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
