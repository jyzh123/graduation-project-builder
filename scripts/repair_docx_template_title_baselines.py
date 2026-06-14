#!/usr/bin/env python3
"""Apply template-owned title baselines to fixed thesis title surfaces.

This canonical helper is intentionally narrow. It copies paragraph properties
and non-text run properties for the Chinese abstract title, English abstract
title, bibliography title, and acknowledgement title from a locked template
DOCX into a candidate DOCX without touching body text, tables, figures, TOC, or
headers/footers.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable

from lxml import etree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"

SURFACE_LABELS = {
    "zh_abstract": ("\u6458\u8981",),
    "en_abstract": ("abstract",),
    "references": ("\u53c2\u8003\u6587\u732e", "references", "bibliography"),
    "acknowledgement": ("\u81f4\u8c22", "\u8c22\u8f9e", "acknowledgements", "acknowledgments"),
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def text_of(paragraph: ET._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def compact(value: str) -> str:
    return "".join((value or "").split()).lower()


def surface_name(text: str) -> str | None:
    normalized = compact(text)
    for name, labels in SURFACE_LABELS.items():
        if normalized in {compact(label) for label in labels}:
            return name
    return None


def body_paragraphs(root: ET._Element) -> list[ET._Element]:
    return list(root.xpath("//w:body/w:p", namespaces=NS))


def find_surface_paragraphs(root: ET._Element) -> dict[str, ET._Element]:
    found: dict[str, ET._Element] = {}
    for paragraph in body_paragraphs(root):
        name = surface_name(text_of(paragraph))
        if name and name not in found:
            found[name] = paragraph
    return found


def non_text_run_props(run: ET._Element) -> ET._Element | None:
    rpr = run.find("w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def visible_runs(paragraph: ET._Element) -> list[ET._Element]:
    return [
        run
        for run in paragraph.findall("w:r", NS)
        if "".join(run.xpath(".//w:t/text()", namespaces=NS)).strip()
    ]


def replace_or_insert_child(parent: ET._Element, child: ET._Element, tag: str, insert_at: int = 0) -> None:
    existing = parent.find(f"w:{tag}", NS)
    if existing is not None:
        parent.remove(existing)
    parent.insert(insert_at, copy.deepcopy(child))


def copy_title_baseline(
    target: ET._Element,
    donor: ET._Element,
    *,
    remove_page_break_before: bool = False,
) -> dict[str, object]:
    target_text = text_of(target)
    donor_ppr = donor.find("w:pPr", NS)
    if donor_ppr is not None:
        donor_ppr_copy = copy.deepcopy(donor_ppr)
        if remove_page_break_before:
            page_break = donor_ppr_copy.find("w:pageBreakBefore", NS)
            if page_break is not None:
                donor_ppr_copy.remove(page_break)
        replace_or_insert_child(target, donor_ppr_copy, "pPr", insert_at=0)

    donor_runs = visible_runs(donor)
    donor_rpr = non_text_run_props(donor_runs[0]) if donor_runs else None
    changed_runs = 0
    for run in visible_runs(target):
        existing = run.find("w:rPr", NS)
        if existing is not None:
            run.remove(existing)
        if donor_rpr is not None:
            run.insert(0, copy.deepcopy(donor_rpr))
        changed_runs += 1
    return {
        "text": target_text,
        "donor_text": text_of(donor),
        "target_visible_run_count": changed_runs,
            "donor_has_ppr": donor_ppr is not None,
            "donor_has_rpr": donor_rpr is not None,
            "removed_page_break_before": remove_page_break_before,
    }


def iter_zip_infos(source: zipfile.ZipFile) -> Iterable[zipfile.ZipInfo]:
    return sorted(source.infolist(), key=lambda item: item.filename)


def repair(input_docx: Path, template_docx: Path, output_docx: Path, report_json: Path) -> int:
    before_hash = sha256_file(input_docx)
    template_hash = sha256_file(template_docx)
    with zipfile.ZipFile(template_docx, "r") as template_zip:
        template_root = ET.fromstring(template_zip.read("word/document.xml"))
    donor_map = find_surface_paragraphs(template_root)
    missing_donors = sorted(set(SURFACE_LABELS) - set(donor_map))
    if missing_donors:
        report_json.parent.mkdir(parents=True, exist_ok=True)
        report_json.write_text(
            json.dumps(
                {
                    "schema": "graduation-project-builder.template-title-baseline-repair.v1",
                    "input_docx_path": str(input_docx),
                    "template_docx_path": str(template_docx),
                    "missing_template_surfaces": missing_donors,
                    "verdict": "failed",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return 1

    changed_rows: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory() as td:
        work = Path(td)
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(work)
        doc_xml = work / "word" / "document.xml"
        parser = ET.XMLParser(remove_blank_text=False)
        tree = ET.parse(str(doc_xml), parser)
        root = tree.getroot()
        target_map = find_surface_paragraphs(root)
        missing_targets = sorted(set(SURFACE_LABELS) - set(target_map))
        if missing_targets:
            report_json.parent.mkdir(parents=True, exist_ok=True)
            report_json.write_text(
                json.dumps(
                    {
                        "schema": "graduation-project-builder.template-title-baseline-repair.v1",
                        "input_docx_path": str(input_docx),
                        "template_docx_path": str(template_docx),
                        "missing_target_surfaces": missing_targets,
                        "verdict": "failed",
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            return 1
        paragraph_positions = {id(p): i for i, p in enumerate(body_paragraphs(root), start=1)}
        for name in SURFACE_LABELS:
            row = copy_title_baseline(
                target_map[name],
                donor_map[name],
                remove_page_break_before=name in {"references", "acknowledgement"},
            )
            row["surface"] = name
            row["paragraph_index"] = paragraph_positions.get(id(target_map[name]))
            changed_rows.append(row)
        doc_xml.write_bytes(ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True))

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(
            output_docx, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for item in iter_zip_infos(zin):
                data = (work / item.filename).read_bytes() if not item.is_dir() else b""
                zi = zipfile.ZipInfo(item.filename)
                zi.date_time = item.date_time
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = item.external_attr
                zout.writestr(zi, data)
    after_hash = sha256_file(output_docx)
    report_json.parent.mkdir(parents=True, exist_ok=True)
    report_json.write_text(
        json.dumps(
            {
                "schema": "graduation-project-builder.template-title-baseline-repair.v1",
                "input_docx_path": str(input_docx),
                "input_docx_sha256": before_hash,
                "template_docx_path": str(template_docx),
                "template_docx_sha256": template_hash,
                "output_docx_path": str(output_docx),
                "output_docx_sha256": after_hash,
                "changed_parts": ["word/document.xml"],
                "changes": changed_rows,
                "verdict": "pass" if len(changed_rows) == len(SURFACE_LABELS) else "failed",
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return 0 if len(changed_rows) == len(SURFACE_LABELS) else 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--template-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report-json", required=True)
    args = parser.parse_args()
    return repair(
        Path(args.input_docx),
        Path(args.template_docx),
        Path(args.output_docx),
        Path(args.report_json),
    )


if __name__ == "__main__":
    raise SystemExit(main())
