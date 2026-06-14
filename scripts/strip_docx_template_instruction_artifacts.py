#!/usr/bin/env python3
"""Strip visible template instruction artifacts from thesis DOCX front matter.

This is a generic skill-local helper. Project-local thesis adapters may point to
templates and content manifests, but template-instruction cleanup belongs here.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = "{%s}" % NS["w"]
V = "{%s}" % NS["v"]

INSTRUCTION_KEYWORDS = (
    "\u5b57\u4f53",  # font family
    "\u5b57\u53f7",  # font size
    "\u4e2d\u6587\u9898\u76ee",
    "\u82f1\u6587\u9898\u76ee",
    "\u9898\u76ee\u53ea\u6709\u4e00\u884c",
    "\u987b\u5220\u9664\u672c\u884c",
    "\u6bb5\u524d",
    "\u6bb5\u540e",
    "\u884c\u8ddd",
    "\u56fa\u5b9a\u503c",
    "\u5c45\u4e2d",
    "\u9876\u683c",
    "\u586b\u5199",
    "\u66ff\u6362",
    "\u91c7\u7528",
    "\u9ed1\u4f53",
    "\u5b8b\u4f53",
    "\u4eff\u5b8b",
    "\u5c0f\u4e94",
    "\u5c0f\u56db\u53f7",
    "\u4e09\u53f7",
    "\u78c5",
    "\u4e8c\u9009\u4e00",
    "\u8bba\u6587\u3001\u8bbe\u8ba1\u4e8c\u9009\u4e00",
    "\u5173\u952e\u8bcd",
    "\u6458\u8981\u5355\u72ec\u6210\u9875",
    "\u76ee\u5f55\u5355\u72ec\u6210\u9875",
    "\u6bcf\u4e2a\u5173\u952e\u8bcd",
    "\u4e24\u4e2a\u6c49\u5b57",
    "\u4e0d\u7528\u963f\u62c9\u4f2f\u6570\u5b57",
    "\u65e0\u9700\u586b\u5199",
    "\u7b54\u8fa9",
    "\u4e00\u8fa9",
    "\u4e8c\u8fa9",
    "\u624b\u7b7e",
    "\u9ed1\u8272\u4e2d\u6027\u7b14",
    "\u9875\u7709",
    "\u8868\u9898",
    "\u56fe\u5e8f",
    "times new roman",
    "arial",
    "pt",
)

PROCESS_NOTE_PHRASES = (
    "\u672c\u6587\u5728\u5f15\u7528\u65f6\u6309\u6b63\u6587\u9996\u6b21\u51fa\u73b0\u987a\u5e8f\u7f16\u53f7",
    "\u53c2\u8003\u6587\u732e\u533a\u4e5f\u6309\u540c\u4e00\u987a\u5e8f\u5217\u51fa",
    "\u907f\u514d\u51fa\u73b0\u56fe\u9898\u3001\u8868\u9898\u548c\u6807\u9898\u4e0a\u6302\u5f15\u7528\u7f16\u53f7\u7684\u95ee\u9898",
    "\u8865\u5145\u5f15\u7528\u8bf4\u660e",
    "\u5148\u5b8c\u6210pdf+dwg\u5de5\u7a0b\u56fe\u7eb8\u5305",
    "\u518d\u6309\u56fe\u7eb8\u53c2\u6570\u53cd\u63a8\u8ba1\u7b97\u4e66\u548c\u8bf4\u660e\u4e66\u5185\u5bb9",
    "\u4ea4\u4ed8pdf\u4e0edwg\u6e90\u6587\u4ef6",
    "\u5df2\u4f5c\u4e3a\u72ec\u7acbaa0\u56fe\u7eb8\u4ea4\u4ed8",
    "\u6e90\u6587\u4ef6\u5305\u542bpdf\u3001dwg\u548cdxf\u683c\u5f0f",
)


def normalize(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "").strip().lower()


def load_document_xml(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def load_xml_part(docx_path: Path, part_name: str) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read(part_name))


def xml_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", NS))


def direct_paragraph_text(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for run in paragraph.findall("./w:r", NS):
        if run.find(".//w:txbxContent", NS) is not None:
            continue
        parts.append("".join(node.text or "" for node in run.findall(".//w:t", NS)))
    return "".join(parts)


def textbox_texts(paragraph: ET.Element) -> list[str]:
    return [xml_text(node).strip() for node in paragraph.findall(".//w:txbxContent", NS) if xml_text(node).strip()]


def has_vml_image_payload(paragraph: ET.Element) -> bool:
    return paragraph.find(".//v:imagedata", NS) is not None


def has_real_image_payload(paragraph: ET.Element) -> bool:
    if paragraph.find(".//v:imagedata", NS) is not None:
        return True
    if paragraph.find(".//a:blip", NS) is not None:
        return True
    return any(element.tag.endswith("}blip") for element in paragraph.iter())


def has_arrow_or_line_artifact(paragraph: ET.Element) -> bool:
    if has_vml_image_payload(paragraph):
        return False
    for element in paragraph.iter():
        if element.tag in {V + "line", V + "shape"}:
            style = " ".join(str(value).lower() for value in element.attrib.values())
            if element.tag == V + "line" or "arrow" in style or "callout" in style:
                return True
        if element.tag.endswith("}stroke"):
            style = " ".join(str(value).lower() for value in element.attrib.values())
            if "arrow" in style:
                return True
    return False


def is_instruction_text(text: str) -> bool:
    lowered = normalize(text)
    if not lowered:
        return False
    strong_instruction_tokens = (
        "\u987b\u5220\u9664\u672c\u884c",
        "\u5220\u9664\u672c\u884c",
        "\u65e0\u9700\u586b\u5199",
        "\u65e0\u987b\u586b\u5199",
        "\u9898\u76ee\u53ea\u6709\u4e00\u884c",
        "\u8bba\u6587\u3001\u8bbe\u8ba1\u4e8c\u9009\u4e00",
        "\u65e5\u671f\u4e3a\u7b54\u8fa9\u4e4b\u524d",
        "\u672a\u901a\u8fc7\u4e00\u8fa9\u8005\u5199\u4e8c\u8fa9\u4e4b\u524d\u7684\u65e5\u671f",
        "\u672c\u4eba\u9ed1\u8272\u4e2d\u6027\u7b14\u624b\u7b7e",
        "\u6559\u5e08\u9ed1\u8272\u4e2d\u6027\u7b14\u624b\u7b7e",
    )
    if any(normalize(token) in lowered for token in strong_instruction_tokens):
        return True
    if len(lowered) <= 30 and any(
        normalize(token) in lowered
        for token in (
            "\u4e2d\u6587\u9898\u76ee",
            "\u82f1\u6587\u9898\u76ee",
            "\u8bba\u6587\u3001\u8bbe\u8ba1",
        )
    ):
        return True
    weak_content_tokens = {
        normalize("\u91c7\u7528"),  # adopted/used; common in real thesis prose
        normalize("\u7b54\u8fa9"),  # defense/demo; can appear in real system-scope prose
        normalize("\u4e00\u8fa9"),
        normalize("\u4e8c\u8fa9"),
    }
    hits = sum(
        1
        for token in INSTRUCTION_KEYWORDS
        if normalize(token) in lowered and normalize(token) not in weak_content_tokens
    )
    hard_format_tokens = (
        "\u5b57\u4f53",
        "\u5b57\u53f7",
        "\u9ed1\u4f53",
        "\u5b8b\u4f53",
        "\u4eff\u5b8b",
        "\u5c0f\u4e94",
        "\u5c0f\u56db",
        "\u4e09\u53f7",
        "\u6bb5\u524d",
        "\u6bb5\u540e",
        "\u884c\u8ddd",
        "\u5c45\u4e2d",
        "\u9876\u683c",
        "\u7a7a\u4e00\u884c",
        "\u4e24\u4e2a\u6c49\u5b57",
        "\u4e8c\u9009\u4e00",
        "\u987b\u5220\u9664\u672c\u884c",
        "\u65e0\u9700\u586b\u5199",
        "\u5355\u72ec\u6210\u9875",
        "\u624b\u7b7e",
        "\u9ed1\u8272\u4e2d\u6027\u7b14",
        "timesnewroman",
        "pt",
    )
    hard_hits = sum(1 for token in hard_format_tokens if normalize(token) in lowered)
    if "\u5355\u72ec\u6210\u9875" in lowered and lowered.startswith(normalize("\u6ce8")):
        return True
    if any(token in lowered for token in ("timesnewroman", "arial")) and any(token in lowered for token in ("\u5b57\u4f53", "\u5b57", "pt")):
        return True
    if "\u4e8c\u9009\u4e00" in lowered and ("\u8bba\u6587" in lowered or "\u8bbe\u8ba1" in lowered):
        return True
    return hits >= 2 and hard_hits >= 1


def is_instruction_note_text(text: str) -> bool:
    lowered = normalize(text)
    if not lowered.startswith(normalize("\u6ce8")):
        return False
    return any(
        token in lowered
        for token in (
            normalize("\u6458\u8981\u5355\u72ec\u6210\u9875"),
            normalize("\u76ee\u5f55\u5355\u72ec\u6210\u9875"),
            normalize("\u5355\u72ec\u6210\u9875"),
        )
    )


def is_process_note_text(text: str) -> bool:
    """Detect generated process notes that describe citation mechanics, not thesis content."""
    lowered = normalize(text)
    if not lowered:
        return False
    hits = sum(1 for token in PROCESS_NOTE_PHRASES if normalize(token) in lowered)
    if hits >= 2:
        return True
    if "pdf+dwg" in lowered and "\u53cd\u63a8" in lowered and "\u8bf4\u660e\u4e66" in lowered:
        return True
    if "pdf+dwg" in lowered and ("\u5de5\u7a0b\u56fe\u7eb8\u5305" in lowered or "\u56fe\u7eb8\u5305" in lowered):
        return True
    if "\u4ea4\u4ed8" in lowered and "pdf" in lowered and "dwg" in lowered:
        return True
    if "\u6e90\u6587\u4ef6\u5305\u542bpdf" in lowered and "dwg" in lowered:
        return True
    return lowered.startswith(normalize("\u8865\u5145\u5f15\u7528\u8bf4\u660e"))


def table_cell_texts(table: ET.Element) -> list[str]:
    return [xml_text(cell).strip() for cell in table.findall(".//w:tc", NS) if xml_text(cell).strip()]


def is_removable_instruction_table(table: ET.Element) -> bool:
    """Return true only for isolated instruction tables, not real cover metadata tables."""
    cells = table_cell_texts(table)
    table_text = " ".join(cells).strip()
    if not table_text or not is_instruction_text(table_text):
        return False
    normalized = normalize(table_text)
    strong_delete_markers = (
        normalize("\u987b\u5220\u9664\u672c\u884c"),
        normalize("\u5220\u9664\u672c\u884c"),
        normalize("\u65e0\u9700\u586b\u5199"),
        normalize("\u586b\u5199\u8bf4\u660e"),
        normalize("\u6ce8\uff1a"),
        normalize("\u6ce8:"),
    )
    if not any(marker in normalized for marker in strong_delete_markers):
        return False
    substantive_cells = [
        cell
        for cell in cells
        if cell
        and not is_instruction_text(cell)
        and not is_instruction_note_text(cell)
        and not re.fullmatch(r"[\s\u25a1:锛氥€?\-_/\\|]+", cell)
    ]
    # Front-matter field tables often mix labels/instructions with real project
    # values. Such tables must be preserved and repaired by a surface-aware lane.
    return not substantive_cells


def heading_level(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    sep = r"[\s\u25a1]+"
    if re.match(rf"^\d{{1,2}}{sep}\S", stripped):
        return 1
    if re.match(r"^\u7b2c[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u7ae0", stripped):
        return 1
    return None


def body_start_index(paragraphs: list[ET.Element]) -> int:
    for idx, paragraph in enumerate(paragraphs):
        if heading_level(direct_paragraph_text(paragraph).strip() or xml_text(paragraph).strip()) is not None:
            return idx
    return len(paragraphs)


def body_start_child_index(children: list[ET.Element]) -> int:
    for idx, child in enumerate(children):
        if child.tag == W + "p" and heading_level(direct_paragraph_text(child).strip() or xml_text(child).strip()) is not None:
            return idx
    return len(children)


def has_page_break(paragraph: ET.Element) -> bool:
    return any(node.attrib.get(W + "type") == "page" for node in paragraph.findall(".//w:br", NS))


def remove_descendants(paragraph: ET.Element, tag: str) -> int:
    removed = 0
    for element in list(paragraph.findall(".//" + tag)):
        parent = _parent(paragraph, element)
        if parent is not None:
            parent.remove(element)
            removed += 1
    return removed


def remove_descendants_if(paragraph: ET.Element, tag: str, predicate) -> int:
    removed = 0
    for element in list(paragraph.findall(".//" + tag)):
        if not predicate(element):
            continue
        parent = _parent(paragraph, element)
        if parent is not None:
            parent.remove(element)
            removed += 1
    return removed


def _parent(root: ET.Element, target: ET.Element) -> ET.Element | None:
    for candidate in root.iter():
        for child in list(candidate):
            if child is target:
                return candidate
    return None


def paragraph_has_visible_payload(paragraph: ET.Element) -> bool:
    return bool(direct_paragraph_text(paragraph).strip()) or has_page_break(paragraph) or has_real_image_payload(paragraph)


def element_has_real_image_payload(element: ET.Element) -> bool:
    if element.find(".//v:imagedata", NS) is not None:
        return True
    if element.find(".//a:blip", NS) is not None:
        return True
    return any(child.tag.endswith("}blip") for child in element.iter())


def element_has_textbox(element: ET.Element) -> bool:
    return element.find(".//w:txbxContent", NS) is not None


def is_arrow_or_line_element(element: ET.Element) -> bool:
    if element.find(".//v:imagedata", NS) is not None:
        return False
    if element.tag in {V + "line", V + "shape"}:
        style = " ".join(str(value).lower() for value in element.attrib.values())
        return element.tag == V + "line" or "arrow" in style or "callout" in style
    if element.tag.endswith("}stroke"):
        style = " ".join(str(value).lower() for value in element.attrib.values())
        return "arrow" in style
    return False


def collect_instruction_artifacts_from_root(
    root: ET.Element,
    *,
    scope: str = "front-matter",
    part_name: str = "word/document.xml",
) -> list[dict[str, object]]:
    body = root.find("w:body", NS)
    container = body if body is not None else root
    children = list(container)
    child_limit = body_start_child_index(children) if scope == "front-matter" and body is not None else len(children)
    paragraphs = list(container.findall(".//w:p", NS)) if body is None else [child for child in list(container) if child.tag == W + "p"]
    limit = body_start_index(paragraphs) if scope == "front-matter" and body is not None else len(paragraphs)
    artifacts: list[dict[str, object]] = []
    for child_idx, child in enumerate(children[:child_limit]):
        if child.tag != W + "tbl":
            continue
        table_text = xml_text(child).strip()
        if is_removable_instruction_table(child):
            artifacts.append(
                {
                    "paragraph_index": child_idx,
                    "part_name": part_name,
                    "kind": "instruction-table",
                    "direct_text": "",
                    "artifact_text": table_text,
                }
            )
    for idx, paragraph in enumerate(paragraphs[:limit]):
        direct_text = direct_paragraph_text(paragraph).strip()
        embedded = [text for text in textbox_texts(paragraph) if is_instruction_text(text)]
        if embedded:
            artifacts.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "textbox-shape",
                    "direct_text": direct_text,
                    "artifact_text": " | ".join(embedded),
                }
            )
        if has_arrow_or_line_artifact(paragraph):
            artifacts.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "vml-arrow-or-line",
                    "direct_text": direct_text,
                    "artifact_text": "arrow-or-line-shape-without-image-payload",
                }
            )
        if direct_text and is_instruction_note_text(direct_text):
            artifacts.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "note-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
        elif direct_text and is_process_note_text(direct_text):
            artifacts.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "process-note-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
        elif direct_text and is_instruction_text(direct_text):
            artifacts.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "instruction-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
    return artifacts


def collect_instruction_artifacts(docx_path: Path, *, scope: str = "front-matter") -> list[dict[str, object]]:
    if scope != "all-parts":
        return collect_instruction_artifacts_from_root(load_document_xml(docx_path), scope=scope)
    artifacts: list[dict[str, object]] = []
    with zipfile.ZipFile(docx_path) as zf:
        part_names = [
            name
            for name in zf.namelist()
            if name == "word/document.xml"
            or re.fullmatch(r"word/(?:header|footer|footnotes|endnotes)\d*\.xml", name)
        ]
        for part_name in sorted(part_names):
            try:
                root = ET.fromstring(zf.read(part_name))
            except Exception:
                continue
            artifacts.extend(
                collect_instruction_artifacts_from_root(root, scope="all", part_name=part_name)
            )
    return artifacts


def strip_instruction_artifacts_from_root(
    root: ET.Element,
    *,
    scope: str = "front-matter",
    part_name: str = "word/document.xml",
) -> dict[str, object]:
    body = root.find("w:body", NS)
    container = body if body is not None else root
    paragraphs = list(container.findall(".//w:p", NS)) if body is None else [child for child in list(container) if child.tag == W + "p"]
    removed_shapes = 0
    removed_notes = 0
    removed_tables = 0
    removed_empty_paragraphs = 0
    artifacts_before: list[dict[str, object]] = []

    children = list(container)
    child_limit = body_start_child_index(children) if scope == "front-matter" and body is not None else len(children)
    for child_idx, child in enumerate(list(children[:child_limit])):
        if child.tag != W + "tbl":
            continue
        table_text = xml_text(child).strip()
        if is_removable_instruction_table(child):
            artifacts_before.append(
                {
                    "paragraph_index": child_idx,
                    "part_name": part_name,
                    "kind": "instruction-table",
                    "direct_text": "",
                    "artifact_text": table_text,
                }
            )
            container.remove(child)
            removed_tables += 1

    paragraphs = list(container.findall(".//w:p", NS)) if body is None else [child for child in list(container) if child.tag == W + "p"]
    limit = body_start_index(paragraphs) if scope == "front-matter" and body is not None else len(paragraphs)

    for idx, paragraph in enumerate(list(paragraphs[:limit])):
        direct_text = direct_paragraph_text(paragraph).strip()
        embedded = [text for text in textbox_texts(paragraph) if is_instruction_text(text)]
        if embedded:
            artifacts_before.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "textbox-shape",
                    "direct_text": direct_text,
                    "artifact_text": " | ".join(embedded),
                }
            )
            removed_shapes += remove_descendants_if(
                paragraph,
                W + "drawing",
                lambda element: element_has_textbox(element) and not element_has_real_image_payload(element),
            )
            removed_shapes += remove_descendants_if(
                paragraph,
                W + "pict",
                lambda element: element_has_textbox(element) and not element_has_real_image_payload(element),
            )
            if not paragraph_has_visible_payload(paragraph):
                parent = _parent(container, paragraph) or _parent(root, paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed_empty_paragraphs += 1
            continue
        if has_arrow_or_line_artifact(paragraph):
            artifacts_before.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "vml-arrow-or-line",
                    "direct_text": direct_text,
                    "artifact_text": "arrow-or-line-shape-without-image-payload",
                }
            )
            removed_shapes += remove_descendants_if(paragraph, V + "line", is_arrow_or_line_element)
            removed_shapes += remove_descendants_if(paragraph, V + "shape", is_arrow_or_line_element)
            removed_shapes += remove_descendants_if(paragraph, V + "stroke", is_arrow_or_line_element)
            if not paragraph_has_visible_payload(paragraph):
                parent = _parent(container, paragraph) or _parent(root, paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed_empty_paragraphs += 1
            continue
        if direct_text and is_instruction_note_text(direct_text):
            artifacts_before.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "note-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
            if not has_page_break(paragraph):
                parent = _parent(container, paragraph) or _parent(root, paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed_notes += 1
            continue
        if direct_text and is_process_note_text(direct_text):
            artifacts_before.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "process-note-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
            if not has_page_break(paragraph):
                parent = _parent(container, paragraph) or _parent(root, paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed_notes += 1
            continue
        if direct_text and is_instruction_text(direct_text):
            artifacts_before.append(
                {
                    "paragraph_index": idx,
                    "part_name": part_name,
                    "kind": "instruction-paragraph",
                    "direct_text": direct_text,
                    "artifact_text": direct_text,
                }
            )
            if not has_page_break(paragraph):
                parent = _parent(container, paragraph) or _parent(root, paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed_notes += 1

    return {
        "part_name": part_name,
        "artifacts_before": artifacts_before,
        "removed_shapes": removed_shapes,
        "removed_tables": removed_tables,
        "removed_note_paragraphs": removed_notes,
        "removed_empty_paragraphs": removed_empty_paragraphs,
    }


def strip_instruction_artifacts(input_docx: Path, output_docx: Path, *, scope: str = "front-matter") -> dict[str, object]:
    part_scope = "all" if scope == "all-parts" else scope
    target_parts: dict[str, bytes] = {}
    part_reports: list[dict[str, object]] = []

    with zipfile.ZipFile(input_docx) as zin:
        part_names = [
            name
            for name in zin.namelist()
            if name == "word/document.xml"
            or (scope == "all-parts" and re.fullmatch(r"word/(?:header|footer|footnotes|endnotes)\d*\.xml", name))
        ]
        for part_name in sorted(part_names):
            try:
                root = ET.fromstring(zin.read(part_name))
            except Exception:
                continue
            report = strip_instruction_artifacts_from_root(root, scope=part_scope, part_name=part_name)
            part_reports.append(report)
            if (
                report["removed_shapes"]
                or report["removed_tables"]
                or report["removed_note_paragraphs"]
                or report["removed_empty_paragraphs"]
            ):
                target_parts[part_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output_docx.with_name(output_docx.name + ".tmp")
    with zipfile.ZipFile(input_docx) as zin, zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = target_parts.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    if output_docx.exists():
        output_docx.unlink()
    shutil.move(str(temp_output), str(output_docx))

    artifacts_before: list[dict[str, object]] = []
    for report in part_reports:
        artifacts_before.extend(report.get("artifacts_before", []))

    return {
        "input": str(input_docx),
        "output": str(output_docx),
        "scope": scope,
        "artifacts_before": artifacts_before,
        "part_reports": part_reports,
        "removed_shapes": sum(int(report.get("removed_shapes", 0)) for report in part_reports),
        "removed_tables": sum(int(report.get("removed_tables", 0)) for report in part_reports),
        "removed_note_paragraphs": sum(int(report.get("removed_note_paragraphs", 0)) for report in part_reports),
        "removed_empty_paragraphs": sum(int(report.get("removed_empty_paragraphs", 0)) for report in part_reports),
        "artifacts_after": collect_instruction_artifacts(output_docx, scope=scope),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Input DOCX")
    parser.add_argument("--output", help="Output DOCX; defaults to in-place")
    parser.add_argument("--scope", choices=["front-matter", "all", "all-parts"], default="front-matter")
    parser.add_argument("--report-json", help="Optional JSON report path")
    parser.add_argument("--check-only", action="store_true", help="Report artifacts without modifying the DOCX")
    parser.add_argument("--fail-on-artifacts", action="store_true", help="Return non-zero if artifacts are found")
    args = parser.parse_args()

    input_docx = Path(args.input).resolve()
    output_docx = Path(args.output).resolve() if args.output else input_docx
    if args.check_only:
        report = {
            "input": str(input_docx),
            "scope": args.scope,
            "artifacts": collect_instruction_artifacts(input_docx, scope=args.scope),
        }
    else:
        report = strip_instruction_artifacts(input_docx, output_docx, scope=args.scope)
    if args.report_json:
        report_path = Path(args.report_json).resolve()
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    artifacts = report.get("artifacts") if args.check_only else report.get("artifacts_after")
    if args.fail_on_artifacts and artifacts:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
