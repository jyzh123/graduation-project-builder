"""Shared utility helpers for validate_skill_gate."""

from __future__ import annotations

from pathlib import Path

try:
    from .validate_skill_gate_registry_core import EXPLICIT_VALUES, PLACEHOLDER_VALUES
except ImportError:
    from validate_skill_gate_registry_core import EXPLICIT_VALUES, PLACEHOLDER_VALUES


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def read_lines(path: Path) -> list[str]:
    return read_text(path).splitlines()


def normalize(value: str) -> str:
    return " ".join(value.split()).strip()


def normalized_text(path: Path) -> str:
    return normalize(read_text(path)).lower()


def parse_line_value(line: str) -> str:
    _, _, value = line.partition(":")
    return normalize(value)


def raw_line_value(line: str) -> str:
    _, _, value = line.partition(":")
    return value.strip()


def is_explicit(value: str) -> bool:
    return value not in PLACEHOLDER_VALUES


def is_explicit_or_none(value: str) -> bool:
    return value in EXPLICIT_VALUES or is_explicit(value)


def is_explicit_none(value: str) -> bool:
    normalized = normalize(value).lower()
    return value in EXPLICIT_VALUES or normalized.startswith("not-applicable-with-reason")


def find_lines_with_prefix(lines: list[str], prefix: str) -> list[str]:
    return [line for line in lines if normalize(line).lower().startswith(prefix.lower())]


def contains_any(text: str, tokens: tuple[str, ...] | set[str]) -> bool:
    lowered = normalize(text).lower()
    return any(token.lower() in lowered for token in tokens)


def split_path_values(value: str) -> list[str]:
    return [part.strip() for part in value.split(";") if part.strip()]


def resolve_record_path(raw_path: str, anchor_path: Path) -> Path:
    candidate = Path(raw_path)
    if candidate.is_absolute():
        return candidate
    return (anchor_path.parent / candidate).resolve()


def validate_existing_path(path: Path, *, require_nonempty_file: bool) -> list[str]:
    if not path.exists():
        return [f"path does not exist: {path}"]
    if require_nonempty_file:
        if path.is_dir():
            return [f"path must be a file, not a directory: {path}"]
        try:
            if path.stat().st_size <= 0:
                return [f"file is empty: {path}"]
        except OSError as exc:
            return [f"path could not be inspected: {path} ({exc})"]
    return []


def build_raw_context(raw_values: dict[str, str], prefixes: list[str] | tuple[str, ...]) -> str:
    return " ".join(raw_values.get(prefix, "") for prefix in prefixes)


def append_missing_explicit_prefix_issues(
    *,
    values: dict[str, str],
    prefixes: list[str] | tuple[str, ...],
    issues: list[str],
    location: Path,
    message: str,
) -> None:
    for prefix in prefixes:
        value = values[prefix]
        if value in EXPLICIT_VALUES or not is_explicit(value):
            issues.append(f"{message}: {location} ({prefix})")


def append_incomplete_prefix_issues(
    *,
    values: dict[str, str],
    prefixes: list[str] | tuple[str, ...],
    issues: list[str],
    location: Path,
    message: str,
) -> None:
    for prefix in prefixes:
        if not is_explicit_or_none(values[prefix]):
            issues.append(f"{message}: {location} ({prefix})")
