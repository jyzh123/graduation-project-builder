# Skill Maintenance Case: SGB620/80T v10 Six-Sheet CAD Baseline

## Trigger

The user reported that earlier SGB620/80T CAD outputs were too simple, had repeated frame overflow and overlap regressions, and later asked to use the current accepted CAD package as the default CAD annotation/detail standard for future tasks.

## Durable Rule Change

- `CORE-FIGURE-010` keeps user-provided CAD/DWG/PDF packages above stored defaults.
- When no stronger current-run user sample exists, the active fallback is the v10 six-sheet SGB620/80T baseline under `references/visual-style-samples/mechanical-cad/sgb62080t-v10-six-sheet-baseline/`.
- The default six-sheet SGB620/80T conveyor set is A0 total assembly, A1 head drive assembly, A1 middle trough assembly, A1 tail tensioning assembly, A2 scraper part, and A2 sprocket shaft component.
- The older v5 three-sheet baseline is secondary extension evidence only and cannot be the sole default for new SGB620/80T-like conveyor packages.
- A locked six-sheet formal package must expose `effective_min_dimensioned_dxf_files=6` or equivalent evidence instead of being rejected by a generic eight-sheet dimension-distribution threshold.

## Validation Targets

- `scripts/audit_mechanical_drawing_package.py --self-test`
- `scripts/check_utf8_clean.py --root <skill-root> --json`
- `scripts/validate_skill_gate.py --skill-root <skill-root>`
- exact v10 SGB620/80T mechanical drawing package audit with `passed=true`
