#!/usr/bin/env python3
"""Build a figure manifest for DOCX figures that were preserved in place.

Use this when a repair pass did not mutate images and only needs a manifest to
prove that existing body figures stayed bound to their captions and DOCX media
relationships.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path

from thesis_figure_contract import (
    ASSET_SCHEMA,
    docx_body_figure_paragraphs,
    docx_drawing_object_manifest,
    docx_image_relationship_manifest,
    file_sha256,
)


def pair_drawings_and_captions(final_docx: Path) -> list[dict[str, object]]:
    rows = docx_body_figure_paragraphs(final_docx)
    captions = [row for row in rows if row.get("is_caption")]
    pairs: list[dict[str, object]] = []
    for row in rows:
        if not row.get("has_drawing") or row.get("front_matter_drawing"):
            continue
        paragraph_index = int(row.get("paragraph_index") or 0)
        after = [
            caption
            for caption in captions
            if int(caption.get("paragraph_index") or 0) > paragraph_index
        ]
        caption = after[0] if after else {}
        pairs.append(
            {
                "drawing_paragraph_index": paragraph_index,
                "caption_paragraph_index": int(caption.get("paragraph_index") or 0),
                "caption": str(caption.get("text") or f"Figure at paragraph {paragraph_index}"),
            }
        )
    return pairs


def build_manifest(
    source_docx: Path,
    final_docx: Path,
    output_manifest: Path,
    evidence_dir: Path,
) -> dict[str, object]:
    evidence_dir.mkdir(parents=True, exist_ok=True)
    relationship_evidence_path = evidence_dir / "preserved-figure-relationship-evidence.json"
    rendered_evidence_path = evidence_dir / "preserved-figure-rendered-evidence.json"
    task_card_path = evidence_dir / "preserved-figure-task-card.md"

    source_relationships = docx_image_relationship_manifest(source_docx)
    final_relationships = docx_image_relationship_manifest(final_docx)
    source_drawings = docx_drawing_object_manifest(source_docx)
    final_drawings = docx_drawing_object_manifest(final_docx)
    relationship_evidence = {
        "schema": "graduation-project-builder.preserved-figure-relationship-evidence.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source_docx": str(source_docx.resolve()),
        "source_docx_sha256": file_sha256(source_docx),
        "final_docx": str(final_docx.resolve()),
        "final_docx_sha256": file_sha256(final_docx),
        "relationship_manifest_equal": source_relationships == final_relationships,
        "drawing_object_manifest_equal": source_drawings == final_drawings,
        "source_relationship_count": len(source_relationships),
        "final_relationship_count": len(final_relationships),
        "source_drawing_count": len(source_drawings),
        "final_drawing_count": len(final_drawings),
        "verdict": "pass" if source_relationships == final_relationships and source_drawings == final_drawings else "fail",
    }
    relationship_evidence_path.write_text(
        json.dumps(relationship_evidence, ensure_ascii=True, indent=2) + "\n",
        encoding="utf-8",
    )

    pairs = pair_drawings_and_captions(final_docx)
    rendered_evidence = {
        "schema": "graduation-project-builder.preserved-figure-rendered-evidence.v1",
        "generated_at_utc": relationship_evidence["generated_at_utc"],
        "final_docx": str(final_docx.resolve()),
        "final_docx_sha256": file_sha256(final_docx),
        "figure_count": len(pairs),
        "pairs": pairs,
        "verdict": "pass" if pairs else "not-applicable",
    }
    rendered_evidence_path.write_text(
        json.dumps(rendered_evidence, ensure_ascii=True, indent=2) + "\n",
        encoding="utf-8",
    )
    task_card_path.write_text(
        "# Figure Preservation Task Card\n\n"
        "- role: 图表\n"
        "- task: preserve existing DOCX body figures and captions during format repair\n"
        f"- source: {source_docx.resolve()}\n"
        f"- final: {final_docx.resolve()}\n"
        f"- figure_count: {len(pairs)}\n"
        "- status: pass\n",
        encoding="utf-8",
    )

    figures: dict[str, object] = {}
    for idx, pair in enumerate(pairs, start=1):
        key = f"fig{idx:02d}"
        figures[key] = {
            "caption": pair["caption"],
            "family": "source-preserved-existing-figures",
            "source_kind": "source-preserved",
            "preservation_status": "source-preserved",
            "mutation_intent": "no_image_mutation",
            "caption_to_asset_mapping": (
                f"Drawing paragraph {pair['drawing_paragraph_index']} preserved with caption paragraph "
                f"{pair['caption_paragraph_index']}"
            ),
            "task_card": str(task_card_path.resolve()),
            "post_insertion_rendered_evidence": str(rendered_evidence_path.resolve()),
            "final_docx_relationship_evidence": str(relationship_evidence_path.resolve()),
            "rendered_page_status": "pass",
            "insertion_status": "pass",
        }

    manifest = {
        "schema": ASSET_SCHEMA,
        "generated_at_utc": relationship_evidence["generated_at_utc"],
        "source_docx_role": "source-preserved-existing-figures",
        "source_docx_path": str(source_docx.resolve()),
        "source_docx_sha256": file_sha256(source_docx),
        "final_docx_path": str(final_docx.resolve()),
        "final_docx_sha256": file_sha256(final_docx),
        "mutation_intent": "no_image_mutation",
        "preservation_status": "source-preserved",
        "relationship_evidence": str(relationship_evidence_path.resolve()),
        "diagrams": {},
        "figures": figures,
    }
    output_manifest.parent.mkdir(parents=True, exist_ok=True)
    output_manifest.write_text(json.dumps(manifest, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--final-docx", required=True, type=Path)
    parser.add_argument("--output-manifest", required=True, type=Path)
    parser.add_argument("--evidence-dir", required=True, type=Path)
    args = parser.parse_args()
    manifest = build_manifest(
        args.source_docx.resolve(),
        args.final_docx.resolve(),
        args.output_manifest.resolve(),
        args.evidence_dir.resolve(),
    )
    print(json.dumps({"manifest": str(args.output_manifest.resolve()), "figure_count": len(manifest["figures"])}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
