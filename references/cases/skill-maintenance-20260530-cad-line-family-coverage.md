# Skill Maintenance Case: CAD Line Family Coverage

Date: 2026-05-30

## Trigger

The user reported that mechanical CAD drawings still needed proper engineering line expression: thick and thin solid lines, dash-dot center lines, dashed hidden lines, hatch/section lines, dimensions, leaders, text, and table/frame lines. A read-only audit agent confirmed that the existing lineweight/linetype audit checked positive lineweights and CENTER/HIDDEN linetypes but did not require complete source line-family coverage on every handoff.

## Rule Change

- `CORE-FIGURE-010` now requires source line family coverage as part of the lineweight/linetype fidelity surface.
- `scripts/audit_cad_dxf_linework_fidelity.py` now checks required mechanical CAD line families, layer lineweight ranges, CENTER/HIDDEN linetypes, and package SHA binding.
- `assets/final-acceptance-template.md`, `assets/mechanical-cad-acceptance-template.md`, and validator registries now expose line-family coverage acceptance fields.

## Validation

- `python -m py_compile scripts/audit_cad_dxf_linework_fidelity.py scripts/validate_skill_gate_registry_core.py scripts/validate_skill_gate_registry_records.py`: pass
- `python scripts/audit_cad_dxf_linework_fidelity.py --self-test`: pass
- `python -c "json.load(open('references/rule-owner-map.json', encoding='utf-8'))"`: pass
- `python scripts/check_utf8_clean.py --root . --json`: pass
- `python scripts/validate_skill_gate.py --skill-root .`: pass

## Local Project Evidence

- Project: `D:\项目\刮板输送机`
- Final CAD package: `D:\项目\刮板输送机\最终交付\SGB620-80T刮板输送机图纸包_PDF-DWG-DXF_线型颜色区分重出版.zip`
- Final package SHA256: `FC54D738CB4FDFD97ECDF89CBEF93D74A6877F8F4AC699B8B0BF09B49E3CC093`
- Line-family audit: `D:\项目\刮板输送机\.codex\graduation-project-builder\20260530-full-redo-thesis-drawings\reports\cad-dxf-linework-fidelity-redo-v3.json`
- Scoped mechanical audit: `D:\项目\刮板输送机\.codex\graduation-project-builder\20260530-full-redo-thesis-drawings\reports\mechanical-drawing-audit-line-style-redo-v3-scoped.json`

## Boundary

This case hardens line-family evidence. It does not downgrade existing full mechanical CAD parity gates for density, rendered overlap, or reference-package complexity.
