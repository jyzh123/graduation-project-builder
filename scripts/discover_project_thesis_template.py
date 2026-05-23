#!/usr/bin/env python3
"""Discover likely thesis template files under a project root."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path


TEMPLATE_EXTENSIONS = {".docx", ".doc"}
EXCLUDE_NAME_PREFIXES = ("~$",)
EXCLUDE_DIR_NAMES = {
    ".git",
    ".codex",
    "__pycache__",
    "node_modules",
    "dist",
    "build",
    "archive",
    "backups",
    "outputs",
    "reports",
}
POSITIVE_TOKENS = (
    "\u6a21\u677f",
    "\u6807\u51c6",
    "\u6821\u5bf9",
    "thesis",
    "template",
    "standard",
)
STRONG_TEMPLATE_STEMS = {
    "\u683c\u5f0f\u6a21\u677f",
    "\u8bba\u6587\u683c\u5f0f\u6a21\u677f",
    "\u6bd5\u4e1a\u8bba\u6587\u683c\u5f0f\u6a21\u677f",
    "thesis-template",
    "template",
}
WEAK_POSITIVE_TOKENS = (
    "\u683c\u5f0f",
    "\u8bba\u6587",
    "\u6bd5\u4e1a",
)
NEGATIVE_TOKENS = (
    "review",
    "review-pass",
    "copy",
    "draft",
    "pass",
    "pass0",
    "sample-output",
    "generated",
    "acceptance",
    "format-align",
    "format-repair",
    "\u4fee\u590d",
    "\u6062\u590d",
    "\u683c\u5f0f\u5bf9\u9f50",
    "\u4fee\u6539\u7a3f",
    "\u7ec8\u7a3f",
    "\u8349\u7a3f",
    "\u6837\u7a3f",
    "\u8f93\u51fa",
)
GENERATED_DIR_PREFIXES = (
    "format-align-run-",
    "format-figure-pagination-audit-",
    "format-recovery-audit-",
    "format-recovery-run-",
    "format-repair-run-",
    "skill-audit-run-",
    "template-format-audit-run-",
)
NEGATIVE_DIR_PARTS = {
    ".codex",
    "integration-temp",
    "integration",
    "temp",
    "tmp",
    "cache",
    "evidence",
    "review",
    "stage",
    "rendered",
    "pages",
    "outputs",
    "reports",
}
OFFICIAL_TOKENS = (
    "撰写与装订规范",
    "装订规范",
    "规范",
    "标准封面",
    "teacher",
    "approved",
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def score_candidate(path: Path) -> int:
    lowered = path.name.lower()
    stem_lowered = path.stem.lower()
    path_parts_lower = tuple(part.lower() for part in path.parts)
    score = 0
    if stem_lowered in {stem.lower() for stem in STRONG_TEMPLATE_STEMS}:
        score += 120
    for token in POSITIVE_TOKENS:
        if token.lower() in lowered:
            score += 20
    for token in WEAK_POSITIVE_TOKENS:
        if token.lower() in lowered:
            score += 3
    for token in NEGATIVE_TOKENS:
        if token.lower() in lowered:
            score -= 30
    for token in OFFICIAL_TOKENS:
        if token.lower() in lowered:
            score += 35
    if any(part in NEGATIVE_DIR_PARTS for part in path_parts_lower):
        score -= 120
    generated_dir = any(part.lower().startswith(GENERATED_DIR_PREFIXES) for part in path.parts)
    if generated_dir:
        if any(token in lowered for token in ("\u6a21\u677f", "template")):
            score -= 5
        else:
            score -= 80
    if ".codex" in path_parts_lower:
        score -= 300
    if path.suffix.lower() == ".docx":
        score += 3
    try:
        size = path.stat().st_size
    except OSError:
        size = 0
    if size > 0:
        score += 1
    return score


def iter_candidates(project_root: Path) -> list[Path]:
    candidates: list[Path] = []
    for path in project_root.rglob("*"):
        if not path.is_file():
            continue
        if path.name.startswith(EXCLUDE_NAME_PREFIXES):
            continue
        if path.suffix.lower() not in TEMPLATE_EXTENSIONS:
            continue
        rel_parts = {part.lower() for part in path.relative_to(project_root).parts[:-1]}
        if rel_parts & EXCLUDE_DIR_NAMES:
            continue
        candidates.append(path)
    candidates.sort(key=lambda item: (-score_candidate(item), len(item.parts), str(item).lower()))
    return candidates


def build_report(project_root: Path, limit: int) -> dict[str, object]:
    candidates = iter_candidates(project_root)
    selected = candidates[0] if candidates else None
    generated_at = datetime.now(timezone.utc)
    authority_class = "none"
    if selected is not None:
        lowered = selected.name.lower()
        if any(token.lower() in lowered for token in OFFICIAL_TOKENS):
            authority_class = "official-format-spec"
        elif any(token.lower() in lowered for token in POSITIVE_TOKENS):
            authority_class = "project-template-candidate"
        else:
            authority_class = "weak-project-document"
    return {
        "schema": "graduation-project-builder.template-discovery.v1",
        "generation_stage": "pre-mutation-template-lock",
        "generated_at_utc": generated_at.isoformat().replace("+00:00", "Z"),
        "generated_at_unix": generated_at.timestamp(),
        "generator": "scripts/discover_project_thesis_template.py",
        "project_root": str(project_root),
        "patterns": sorted(f"*{suffix}" for suffix in TEMPLATE_EXTENSIONS),
        "candidate_count": len(candidates),
        "candidates": [
            {
                "path": str(path),
                "score": score_candidate(path),
                "sha256": sha256_file(path),
                "size": path.stat().st_size,
                "under_codex": ".codex" in {part.lower() for part in path.parts},
            }
            for path in candidates[:limit]
        ],
        "selected_template_path": str(selected) if selected else None,
        "selected_template_fingerprint": sha256_file(selected) if selected else None,
        "selected_template_authority_class": authority_class,
        "selection_reason": (
            "highest score from template-name tokens, directory-risk filtering, and authority-weighted selection"
            if selected
            else "no .doc/.docx thesis template candidates found"
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Discover thesis template candidates in a project root.")
    parser.add_argument("--project-root", required=True)
    parser.add_argument("--output")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--fail-on-none", action="store_true")
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    if not project_root.exists() or not project_root.is_dir():
        raise SystemExit(f"project root is not a directory: {project_root}")

    report = build_report(project_root, args.limit)
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        output = Path(args.output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    if args.fail_on_none and not report["selected_template_path"]:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
