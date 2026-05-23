# Skill Maintenance 20260522: Mechanical CAD Package Parity

## Trigger

The user reported that a mechanical graduation-design drawing set was sketch-level, not comparable to the provided CAD packages, and that the final package needed to follow the same format as a user-provided DWG drawing zip.

## Durable Rule

- Added `CORE-FIGURE-010` in `references/thesis/figure-rules/baseline-and-sourcing.md`.
- The rule requires mechanical graduation drawing tasks to inspect user-provided CAD/PDF packages as the drawing baseline, including package structure, file source format, sheet sizes, title blocks, linework density, dimension/detail depth, and no-overlap rendered review.
- The rule blocks claiming "same format" when the required source format is DWG but no real DWG writer/converter is available.
- The rule explicitly forbids renaming DXF/SVG/PDF/PNG files as `.dwg`.

## Validation Scope

This rule now has executable coverage through `scripts/audit_mechanical_drawing_package.py`. The auditor does not replace expert drafting review, but it fails closed on the repeatable error class: fake or invalid DWG/PDF headers, invalid or mojibake JSON manifests, sparse or structure-less DXF exports, below-threshold sheet workload, one-sheet-only detail concentration, weak per-sheet drawing-object density, insufficient distributed dimensioning, missing mechanical-detail tokens, and DWG byte-density far below the user-provided CAD package baseline. The mojibake detector must avoid over-broad normal Chinese characters that create false positives while still rejecting replacement characters and typical replacement-character-style encoding damage.

The 2026-05-22 follow-up hardening upgraded the auditor to the v2 report schema. Mechanical CAD handoff evidence must now expose `density_verdict` and `manufacturing_depth`, including true dimension totals, geometry entity totals, DWG reference ratio, block insert count, arc count, hatch/section count, and text annotation count. The default command-line thresholds are intentionally no longer sketch-level: they require a multi-sheet DWG/DXF/PDF package, A0-equivalent workload, distributed dimensions, per-sheet entity density, and minimum block/arc/hatch/text depth unless a run explicitly documents a narrower non-production drawing task.

The 2026-05-23 follow-up hardening upgraded the auditor to the v3 report schema. Mechanical CAD handoff evidence must now expose `rendered_review_verdict` in addition to `density_verdict` and `manufacturing_depth`. A package manifest must carry `rendered_review` evidence with real preview paths, per-sheet rows, pass verdicts for no-overlap, text legibility, sheet layout, manufacturing-view depth, and an explicit rejection of entity-count-only acceptance. This closes the false-pass where a package has enough entities but the rendered drawings still look like simplified sketches or show text, tables, title blocks, drawing frames, dimension labels, or annotations overlapping.

The 2026-05-23 acceptance-record follow-up added `assets/mechanical-cad-acceptance-template.md` and a dedicated `validate_skill_gate.py --gate-record` path for CAD-only drawing packages. CAD handoff records must bind the exact final delivery ZIP, audited CAD ZIP, DWG ZIP, combined PDF, strict v3 audit JSON, rendered-review evidence path, no-overlap verdict, text/table/frame overlap verdict, and entity-count-only false-pass rejection. This prevents a CAD-specific evidence record from failing only because it is not a DOCX thesis final-acceptance record, while still failing closed on stale artifacts, missing DWG/PDF outputs, stale audit SHA values, missing rendered review, or visual-overlap shortcuts.

The 2026-05-23 SGB620/80T density closeout follow-up added one more durable lesson: when a mechanical CAD package fails only one sheet on rendered detail density, the preferred repair is not to inject extra floating blocks into page whitespace. The safer high-quality repair is to raise real drawing depth inside the existing view envelopes of that sheet, such as reducer internal ribs/bolt rows, chain-run hidden lines, guide-slot split lines, and bearing/section local structure, while preserving guard distances to BOM tables, title blocks, technical notes, and dimension bands. This was the difference between clearing the sheet-level density gate and reintroducing local crowding.

The 2026-05-23 default-baseline promotion follow-up promoted the accepted SGB620/80T package itself into the skill bundle as the stored fallback default mechanical CAD baseline. Future mechanical CAD tasks must use the copied SGB620/80T benchmark sheets plus the copied v5 audit JSON when the current run does not include a stronger user-provided CAD/PDF package. This promotion is intentionally weaker than a user sample: a stronger current-run user sample wins and must still override the stored baseline.
