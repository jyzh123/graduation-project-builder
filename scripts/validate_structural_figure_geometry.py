#!/usr/bin/env python3
"""Validate source-scale geometry for thesis structural figures."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

try:
    from .thesis_figure_contract import drawio_er_geometry_report, drawio_structural_geometry_report
except ImportError:
    from thesis_figure_contract import drawio_er_geometry_report, drawio_structural_geometry_report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--drawio", required=True, help="Path to the draw.io source file.")
    parser.add_argument(
        "--family",
        default="er",
        choices=["er", "flowchart", "structure", "use-case"],
        help="Structural figure family to validate.",
    )
    parser.add_argument("--output", required=True, help="JSON report path.")
    args = parser.parse_args()

    drawio_path = Path(args.drawio).resolve()
    output_path = Path(args.output).resolve()
    if args.family == "er":
        report = drawio_er_geometry_report(drawio_path)
    else:
        report = drawio_structural_geometry_report(drawio_path, family=args.family)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if report.get("verdict") != "pass":
        for issue in report.get("issues", []):
            print(issue)
        return 1
    print(f"geometry validation passed: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
