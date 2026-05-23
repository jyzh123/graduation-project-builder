# Skill Maintenance Case: Orthogonal Connector Law

- date: 2026-05-20
- mode: skill-maintenance
- trigger: user reported that figure connectors violated the rule that lines must not pass through boxes, must connect box edge to box edge, and must use right-angle corners
- scope: graduation-project-builder skill bundle only; no thesis DOCX/PDF artifacts were modified

## Change

- Added `CORE-FIGURE-009` to require boundary-to-boundary orthogonal connectors for draw.io-backed structural figures.
- Hardened `geometry-and-layout.md` so connectors must use visible source/target nodes, real node perimeters, `edgeStyle=orthogonalEdgeStyle`, and `rounded=0`.
- Extended `validate_structural_figure_geometry.py` beyond ER diagrams so `flowchart`, `structure`, and `use-case` sources can produce source-scale geometry reports.
- Extended `thesis_figure_contract.py` to reject invisible routers, source/targetless line segments, non-orthogonal edge styles, rounded/curved connectors, and through-node connector paths.
- Added targeted selftest coverage for non-orthogonal structural connectors.

## Validation

- `py -3 -m py_compile scripts\thesis_figure_contract.py scripts\validate_structural_figure_geometry.py scripts\validate_skill_gate_registry_bundle.py scripts\selftest_skill_flow.py`: pass
- JSON parse of `references\rule-owner-map.json`: pass
- `py -3 scripts\selftest_skill_flow.py --case figure_manifest_structural_router_collision_rejected --case figure_manifest_structural_line_crosses_node_rejected --case figure_manifest_structural_non_orthogonal_connector_rejected --quiet`: pass
- `py -3 scripts\check_utf8_clean.py --root . --json`: pass, checked 238 files
- `py -3 scripts\validate_skill_gate.py --skill-root .`: pass
- `py -3 scripts\validate_structural_figure_geometry.py --family structure --drawio <fig4-2 drawio> --output .codex\skill-maintenance\20260520-orthogonal-connector-law\fig4-2-geometry-report.json`: pass
