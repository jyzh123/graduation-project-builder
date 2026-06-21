"""Repair TOC entry paragraphs so page-number tabs use dotted leaders."""

from __future__ import annotations

import argparse
from copy import deepcopy
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

from toc_leader_audit import NS, W, audit_docx_toc_dotted_leaders, choose_right_tab, collect_toc_entry_paragraphs, is_body_level1_heading, is_toc_title, load_styles, page_number_after_last_tab, page_number_tail_after_last_tab, page_number_tail_is_render_safe, paragraph_style_id, run_tab_segments, tail_text_node_runs_after_final_tab, w_attr


def ensure_child(parent: ET.Element, tag: str, *, first: bool = False) -> ET.Element:
    child = parent.find(tag, NS)
    if child is not None:
        return child
    child = ET.Element(W + tag.split(":", 1)[1])
    if first:
        parent.insert(0, child)
    else:
        parent.append(child)
    return child


def patch_entry_tabs(paragraph: ET.Element, *, leader: str, default_pos: str) -> bool:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        paragraph.insert(0, ppr)
    tabs = ppr.find("w:tabs", NS)
    if tabs is None:
        tabs = ET.Element(W + "tabs")
        ppr.append(tabs)
    direct_tabs = list(tabs.findall("w:tab", NS))
    target = choose_right_tab(direct_tabs)
    changed = False
    if target is None or (w_attr(target, "val") or "left") != "right":
        target = ET.Element(W + "tab")
        target.set(W + "val", "right")
        target.set(W + "pos", default_pos)
        tabs.append(target)
        changed = True
    for tab in direct_tabs + [target]:
        if (w_attr(tab, "val") or "left") != "right":
            continue
        if w_attr(tab, "pos") != default_pos:
            tab.set(W + "pos", default_pos)
            changed = True
        if tab.attrib.get(W + "leader") != leader:
            tab.set(W + "leader", leader)
            changed = True
    return changed


def style_name(style: ET.Element) -> str:
    name = style.find("w:name", NS)
    return w_attr(name, "val").strip().lower()


def style_id(style: ET.Element) -> str:
    return style.attrib.get(W + "styleId", "")


def template_toc_style_ppr_by_style(reference_styles_xml: bytes | None) -> dict[str, ET.Element]:
    if reference_styles_xml is None:
        return {}
    try:
        root = ET.fromstring(reference_styles_xml)
    except ET.ParseError:
        return {}
    style_by_sid = {style_id(style): style for style in root.findall("w:style", NS) if style_id(style)}

    def effective_outline(style: ET.Element, seen: set[str] | None = None) -> str | None:
        seen = seen or set()
        ppr = style.find("w:pPr", NS)
        outline = ppr.find("w:outlineLvl", NS) if ppr is not None else None
        if outline is not None:
            return w_attr(outline, "val")
        based_on = style.find("w:basedOn", NS)
        parent_id = w_attr(based_on, "val")
        if not parent_id or parent_id in seen:
            return None
        parent = style_by_sid.get(parent_id)
        if parent is None:
            return None
        return effective_outline(parent, seen | {parent_id})

    donors: dict[str, ET.Element] = {}
    for style in root.findall("w:style", NS):
        sid = style_id(style)
        name = style_name(style)
        if not sid or not name.startswith("toc "):
            continue
        ppr = style.find("w:pPr", NS)
        donor = deepcopy(ppr) if ppr is not None else ET.Element(W + "pPr")
        if donor.find("w:outlineLvl", NS) is None:
            outline_val = effective_outline(style)
            if outline_val is not None:
                outline = ET.Element(W + "outlineLvl")
                outline.set(W + "val", outline_val)
                donor.insert(0, outline)
        if list(donor):
            donors[sid] = donor
    return donors


def replace_style_ppr_from_template(style: ET.Element, donor_ppr: ET.Element | None) -> bool:
    if donor_ppr is None:
        return False
    existing = style.find("w:pPr", NS)
    existing_payload = ET.tostring(existing, encoding="utf-8") if existing is not None else b""
    donor_payload = ET.tostring(donor_ppr, encoding="utf-8")
    if existing_payload == donor_payload:
        return False
    donor = deepcopy(donor_ppr)
    if existing is not None:
        index = list(style).index(existing)
        style.remove(existing)
        style.insert(index, donor)
    else:
        style.append(donor)
    return True


def patch_toc_style_tabs(
    styles_xml: bytes | None,
    used_style_ids: set[str],
    *,
    leader: str,
    default_pos: str,
    reference_style_ppr_donors: dict[str, ET.Element] | None = None,
) -> tuple[bytes | None, list[str], list[str]]:
    if styles_xml is None:
        return None, [], []
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return styles_xml, [], []

    changed_style_ids: list[str] = []
    template_metric_style_ids: list[str] = []
    for style in root.findall("w:style", NS):
        sid = style_id(style)
        name = style_name(style)
        if sid not in used_style_ids and not name.startswith("toc "):
            continue
        if replace_style_ppr_from_template(style, (reference_style_ppr_donors or {}).get(sid)):
            template_metric_style_ids.append(sid or name)
        ppr = style.find("w:pPr", NS)
        if ppr is None:
            continue
        tabs = ppr.find("w:tabs", NS)
        if tabs is None:
            continue
        style_changed = False
        for tab in tabs.findall("w:tab", NS):
            if (w_attr(tab, "val") or "left") != "right":
                continue
            if w_attr(tab, "pos") != default_pos:
                tab.set(W + "pos", default_pos)
                style_changed = True
            if tab.attrib.get(W + "leader") != leader:
                tab.set(W + "leader", leader)
                style_changed = True
        if style_changed:
            changed_style_ids.append(sid or name)
    if not changed_style_ids and not template_metric_style_ids:
        return styles_xml, [], []
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed_style_ids, template_metric_style_ids


def load_styles_from_xml(styles_xml: bytes | None) -> dict[str, object]:
    if styles_xml is None:
        return {}
    try:
        ET.fromstring(styles_xml)
    except ET.ParseError:
        return {}

    class _StyleZip:
        def read(self, name: str) -> bytes:
            if name != "word/styles.xml":
                raise KeyError(name)
            return styles_xml

    return load_styles(_StyleZip())  # type: ignore[arg-type,return-value]


def template_toc_ppr_by_style(reference_document_xml: bytes | None, reference_styles_xml: bytes | None) -> dict[str, ET.Element]:
    if reference_document_xml is None:
        return {}
    try:
        reference_root = ET.fromstring(reference_document_xml)
    except ET.ParseError:
        return {}
    reference_styles = load_styles_from_xml(reference_styles_xml)
    reference_entries, _issues = collect_toc_entry_paragraphs(reference_root, reference_styles)
    donors: dict[str, ET.Element] = {}
    for paragraph in reference_entries:
        sid = paragraph_style_id(paragraph)
        if not sid or sid in donors:
            continue
        ppr = paragraph.find("w:pPr", NS)
        if ppr is not None:
            donors[sid] = deepcopy(ppr)
    return donors


def replace_entry_ppr_from_template(paragraph: ET.Element, donor_ppr: ET.Element | None) -> bool:
    if donor_ppr is None:
        return False
    existing = paragraph.find("w:pPr", NS)
    existing_payload = ET.tostring(existing, encoding="utf-8") if existing is not None else b""
    donor_payload = ET.tostring(donor_ppr, encoding="utf-8")
    if existing_payload == donor_payload:
        return False
    if existing is not None:
        paragraph.remove(existing)
    paragraph.insert(0, deepcopy(donor_ppr))
    return True


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def first_text_run_rpr(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall(".//w:r", NS):
        if "".join(node.text or "" for node in run.findall(".//w:t", NS)).strip():
            rpr = run.find("w:rPr", NS)
            return deepcopy(rpr) if rpr is not None else None
    return None


def rpr_without_underline(rpr: ET.Element | None) -> ET.Element | None:
    if rpr is None:
        return None
    cloned = deepcopy(rpr)
    for node in list(cloned.findall("w:rStyle", NS)):
        cloned.remove(node)
    for node in list(cloned.findall("w:u", NS)):
        cloned.remove(node)
    return cloned


def remove_toc_hyperlink_residue(rpr: ET.Element | None) -> bool:
    if rpr is None:
        return False
    changed = False
    for node in list(rpr.findall("w:rStyle", NS)):
        rpr.remove(node)
        changed = True
    for node in list(rpr.findall("w:u", NS)):
        rpr.remove(node)
        changed = True
    return changed


def split_static_toc_text(text: str) -> tuple[str, str] | None:
    value = (text or "").strip()
    if not value:
        return None
    match = re.match(r"^(?P<label>.+?)(?P<page>\d+|[ivxlcdmIVXLCDM]+)\s*$", value)
    if not match:
        return None
    label = match.group("label").rstrip(".\u2026\t ")
    page = match.group("page")
    if not label or not page:
        return None
    return label, page


def split_tabbed_toc_text_with_trailing_page(paragraph: ET.Element) -> bool:
    tab_count, segments = run_tab_segments(paragraph)
    if not tab_count or page_number_after_last_tab(paragraph):
        return False
    if move_trailing_page_number_after_final_tab(paragraph):
        return True
    if paragraph.find(".//w:hyperlink", NS) is not None:
        return False
    if not segments:
        return False
    leading = "".join(segments[:-1]).strip() or segments[0].strip()
    split = split_static_toc_text(leading)
    if split is None:
        return False
    label, page = split
    base_rpr = rpr_without_underline(first_text_run_rpr(paragraph))
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    append_run(paragraph, label, base_rpr)
    append_run(paragraph, rpr=base_rpr, tab=True)
    append_run(paragraph, page, base_rpr)
    return True


def text_nodes_for_direct_runs(paragraph: ET.Element) -> list[tuple[int, ET.Element, ET.Element]]:
    rows: list[tuple[int, ET.Element, ET.Element]] = []
    for run_index, run in enumerate(paragraph.iter(W + "r")):
        for node in run.findall("w:t", NS):
            rows.append((run_index, run, node))
    return rows


def direct_run_has_tab(run: ET.Element) -> bool:
    return run.find("w:tab", NS) is not None


def move_trailing_page_number_after_final_tab(paragraph: ET.Element) -> bool:
    """Move a misplaced TOC page token to the final tab tail.

    This handles rows such as `1.1 研究背景  1<TAB>`, where the paragraph
    structure still has a final tab but the page number was written before it.
    The repair preserves existing paragraph properties and most run structure.
    """

    runs = list(paragraph.iter(W + "r"))
    tab_run_indexes = [idx for idx, run in enumerate(runs) if direct_run_has_tab(run)]
    if not tab_run_indexes:
        return False
    parent_map = {child: parent for parent in paragraph.iter() for child in list(parent)}
    final_tab_run_index = tab_run_indexes[-1]
    before_nodes = [
        (run_index, run, node)
        for run_index, run, node in text_nodes_for_direct_runs(paragraph)
        if run_index < final_tab_run_index
    ]
    after_nodes = [
        (run_index, run, node)
        for run_index, run, node in text_nodes_for_direct_runs(paragraph)
        if run_index > final_tab_run_index
    ]
    page = ""
    source_rpr: ET.Element | None = None
    for _run_index, run, node in reversed(before_nodes):
        value = node.text or ""
        match = re.match(r"^(?P<label>.*?)(?P<gap>[\s\u3000]+)(?P<page>\d+|[ivxlcdmIVXLCDM]+)\s*$", value)
        if not match:
            continue
        label = match.group("label").rstrip()
        page = match.group("page")
        if not label:
            continue
        node.text = label + " "
        rpr = run.find("w:rPr", NS)
        source_rpr = deepcopy(rpr) if rpr is not None else None
        break
    if not page:
        return False
    for _run_index, _run, node in reversed(after_nodes):
        if (node.text or "").strip():
            continue
        node.text = page
        return True
    tab_run = runs[final_tab_run_index]
    tab_parent = parent_map.get(tab_run)
    if tab_parent is None:
        return False
    tab_position = list(tab_parent).index(tab_run)
    page_run = ET.Element(W + "r")
    if source_rpr is not None:
        page_run.append(source_rpr)
    text_node = ET.Element(W + "t")
    text_node.text = page
    page_run.append(text_node)
    tab_parent.insert(tab_position + 1, page_run)
    return True


def text_nodes_after_final_tab(paragraph: ET.Element) -> list[ET.Element]:
    nodes: list[ET.Element] = []
    seen_tab = False
    for node in paragraph.iter():
        if node.tag == W + "tab":
            seen_tab = True
            nodes = []
            continue
        if node.tag != W + "t":
            continue
        text = node.text or ""
        if "\t" in text:
            seen_tab = True
            nodes = [node]
            continue
        if seen_tab:
            nodes.append(node)
    return nodes


def normalize_page_number_tail_after_final_tab(paragraph: ET.Element) -> bool:
    page = page_number_after_last_tab(paragraph)
    tail = page_number_tail_after_last_tab(paragraph)
    if not page or page_number_tail_is_render_safe(tail, page):
        return False
    if tail.strip() != page:
        return False
    nodes = text_nodes_after_final_tab(paragraph)
    if not nodes:
        return False
    for index in range(len(nodes) - 1, -1, -1):
        text = nodes[index].text or ""
        if "\t" not in text:
            continue
        prefix, raw_tail = text.rsplit("\t", 1)
        following_tail = "".join(node.text or "" for node in nodes[index + 1 :])
        if (raw_tail + following_tail).strip() != page:
            return False
        changed = False
        new_text = prefix + "\t" + page
        if text != new_text:
            nodes[index].text = new_text
            changed = True
        for node in nodes[index + 1 :]:
            if node.text:
                node.text = ""
                changed = True
        return changed
    changed = False
    nonempty_nodes = [(node, node.text or "") for node in nodes if (node.text or "").strip()]
    if not nonempty_nodes:
        return False
    if "".join(text.strip() for _node, text in nonempty_nodes) != page:
        return False
    page_node = nonempty_nodes[0][0]
    for node in nodes:
        text = node.text or ""
        if node is page_node:
            new_text = page
            if text != new_text:
                node.text = new_text
                changed = True
            continue
        if text:
            node.text = ""
            changed = True
    return changed


def remove_unsafe_page_number_tail_run_scaling(paragraph: ET.Element) -> bool:
    page = page_number_after_last_tab(paragraph)
    if not page:
        return False
    rows = tail_text_node_runs_after_final_tab(paragraph)
    nonempty_rows = [(run, node.text or "") for run, node in rows if (node.text or "").strip()]
    if not nonempty_rows:
        return False
    if "".join(text.rsplit("\t", 1)[-1].strip() for _run, text in nonempty_rows) != page:
        return False
    changed = False
    for run, _text in nonempty_rows:
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            continue
        for scale in list(rpr.findall("w:w", NS)):
            rpr.remove(scale)
            changed = True
    return changed


def append_repairable_tabbed_toc_candidates(
    root: ET.Element,
    styles: dict[str, object],
    entries: list[ET.Element],
) -> list[ET.Element]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    paragraphs = list(body.iter(W + "p"))
    title_index: int | None = None
    for index, child in enumerate(paragraphs):
        if is_toc_title(child, styles):  # type: ignore[arg-type]
            title_index = index
            break
    if title_index is None:
        return []
    added: list[ET.Element] = []
    known = {id(paragraph) for paragraph in entries}
    seen_toc_like = False
    for child in paragraphs[title_index + 1 :]:
        if id(child) in known:
            seen_toc_like = True
            continue
        text = paragraph_text(child)
        tab_count, _segments = run_tab_segments(child)
        if not text.strip():
            continue
        if tab_count:
            seen_toc_like = True
        if seen_toc_like and not tab_count and is_body_level1_heading(child):  # type: ignore[arg-type]
            break
        if re.match(r"^\s*(?:\d{1,2}\s+|第\s*\d+\s*章|第\d+章|绗)", text) and not tab_count_has_missing_page(child):
            break
        if tab_count_has_missing_page(child) and split_static_toc_text(text) is not None:
            entries.append(child)
            added.append(child)
            known.add(id(child))
            seen_toc_like = True
    return added


def tab_count_has_missing_page(paragraph: ET.Element) -> bool:
    tab_count, _segments = run_tab_segments(paragraph)
    return bool(tab_count) and not bool(page_number_after_last_tab(paragraph))


def append_run(paragraph: ET.Element, text: str = "", rpr: ET.Element | None = None, *, tab: bool = False) -> None:
    run = ET.Element(W + "r")
    if rpr is not None:
        run.append(deepcopy(rpr))
    if tab:
        run.append(ET.Element(W + "tab"))
    if text:
        node = ET.Element(W + "t")
        node.text = text
        run.append(node)
    paragraph.append(run)


def split_concatenated_entry_runs(paragraph: ET.Element) -> bool:
    if split_tabbed_toc_text_with_trailing_page(paragraph):
        return True
    tab_count, _segments = run_tab_segments(paragraph)
    text = paragraph_text(paragraph)
    if tab_count and "\t" not in text:
        # Still remove direct underline and hyperlink character-style residue
        # from visible TOC runs. WPS/LibreOffice render Hyperlink style as
        # underlined even when no direct w:u is present on the run.
        changed = False
        for run in paragraph.findall(".//w:r", NS):
            rpr = run.find("w:rPr", NS)
            if remove_toc_hyperlink_residue(rpr):
                changed = True
        return changed
    split = split_static_toc_text(text)
    if split is None:
        return False
    if paragraph.find(".//w:hyperlink", NS) is not None:
        return False
    label, page = split
    base_rpr = rpr_without_underline(first_text_run_rpr(paragraph))
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    append_run(paragraph, label, base_rpr)
    append_run(paragraph, rpr=base_rpr, tab=True)
    append_run(paragraph, page, base_rpr)
    return True


def repair_document_xml(
    document_xml: bytes,
    styles_xml: bytes | None,
    *,
    leader: str,
    default_pos: str,
    reference_document_xml: bytes | None = None,
    reference_styles_xml: bytes | None = None,
) -> tuple[bytes, bytes | None, dict[str, object]]:
    root = ET.fromstring(document_xml)
    styles = load_styles_from_xml(styles_xml)
    entries, collection_issues = collect_toc_entry_paragraphs(root, styles)
    added_repairable_entries = append_repairable_tabbed_toc_candidates(root, styles, entries)
    used_style_ids = {paragraph_style_id(paragraph) for paragraph in entries if paragraph_style_id(paragraph)}
    template_ppr_donors = template_toc_ppr_by_style(reference_document_xml, reference_styles_xml)
    template_style_ppr_donors = template_toc_style_ppr_by_style(reference_styles_xml)
    patched_styles_xml, changed_style_ids, template_metric_style_ids = patch_toc_style_tabs(
        styles_xml,
        used_style_ids,
        leader=leader,
        default_pos=default_pos,
        reference_style_ppr_donors=template_style_ppr_donors,
    )
    changed_indexes: list[int] = []
    template_metric_indexes: list[int] = []
    split_indexes: list[int] = []
    page_tail_normalized_indexes: list[int] = []
    page_tail_scaling_repaired_indexes: list[int] = []
    body = root.find(".//w:body", NS)
    paragraphs = list(body.iter(W + "p")) if body is not None else []
    for paragraph in entries:
        index = paragraphs.index(paragraph) + 1 if paragraph in paragraphs else -1
        if replace_entry_ppr_from_template(paragraph, template_ppr_donors.get(paragraph_style_id(paragraph))):
            template_metric_indexes.append(index)
        if split_concatenated_entry_runs(paragraph):
            split_indexes.append(index)
        if normalize_page_number_tail_after_final_tab(paragraph):
            page_tail_normalized_indexes.append(index)
        if remove_unsafe_page_number_tail_run_scaling(paragraph):
            page_tail_scaling_repaired_indexes.append(index)
        if patch_entry_tabs(paragraph, leader=leader, default_pos=default_pos):
            changed_indexes.append(index)
    report: dict[str, object] = {
        "schema": "graduation-project-builder.toc-dotted-leader-repair.v1",
        "changed_paragraph_indexes": changed_indexes,
        "template_metric_replayed_paragraph_indexes": template_metric_indexes,
        "split_concatenated_entry_indexes": split_indexes,
        "page_tail_normalized_paragraph_indexes": page_tail_normalized_indexes,
        "page_tail_scaling_repaired_paragraph_indexes": page_tail_scaling_repaired_indexes,
        "entry_count": len(entries),
        "repairable_tabbed_missing_page_entry_count": len(added_repairable_entries),
        "collection_issues": collection_issues,
        "leader": leader,
        "default_right_tab_pos": default_pos,
        "changed_toc_style_ids": changed_style_ids,
        "template_metric_replayed_style_ids": template_metric_style_ids,
    }
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), patched_styles_xml, report


def repair_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    leader: str,
    default_pos: str,
    reference_docx: Path | None = None,
) -> dict[str, object]:
    reference_document_xml: bytes | None = None
    reference_styles_xml: bytes | None = None
    if reference_docx is not None:
        with zipfile.ZipFile(reference_docx, "r") as reference:
            reference_document_xml = reference.read("word/document.xml")
            try:
                reference_styles_xml = reference.read("word/styles.xml")
            except KeyError:
                reference_styles_xml = None
    with zipfile.ZipFile(input_docx, "r") as source:
        document_xml = source.read("word/document.xml")
        try:
            styles_xml = source.read("word/styles.xml")
        except KeyError:
            styles_xml = None
        patched_document_xml, patched_styles_xml, report = repair_document_xml(
            document_xml,
            styles_xml,
            leader=leader,
            default_pos=default_pos,
            reference_document_xml=reference_document_xml,
            reference_styles_xml=reference_styles_xml,
        )
        output_docx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                if info.filename == "word/document.xml":
                    payload = patched_document_xml
                elif info.filename == "word/styles.xml" and patched_styles_xml is not None:
                    payload = patched_styles_xml
                else:
                    payload = source.read(info.filename)
                target.writestr(info, payload)
    report["input_docx"] = str(input_docx)
    report["output_docx"] = str(output_docx)
    if reference_docx is not None:
        report["reference_docx"] = str(reference_docx)
    audit_payload, audit_issues = audit_docx_toc_dotted_leaders(output_docx)
    report["post_repair_audit"] = audit_payload
    report["post_repair_issues"] = audit_issues
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair TOC dotted leaders in a DOCX while preserving package parts.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--leader", default="dot")
    parser.add_argument("--default-right-tab-pos", default="8647")
    parser.add_argument("--reference-docx")
    args = parser.parse_args()

    input_docx = Path(args.input).resolve()
    output_docx = Path(args.output).resolve()
    report_path = Path(args.report).resolve()
    report = repair_docx(
        input_docx,
        output_docx,
        leader=args.leader,
        default_pos=args.default_right_tab_pos,
        reference_docx=Path(args.reference_docx).resolve() if args.reference_docx else None,
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    issues = report.get("post_repair_issues")
    if isinstance(issues, list) and issues:
        print("TOC dotted-leader repair failed audit: " + "; ".join(str(issue) for issue in issues[:6]))
        return 1
    print(f"TOC dotted-leader repair passed: {output_docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
