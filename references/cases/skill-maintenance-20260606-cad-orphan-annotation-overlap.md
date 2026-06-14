# Skill Maintenance Case: CAD Orphan Annotation And Text Overlap

- date: 2026-06-06
- owner rule: `CORE-FIGURE-010`, `CORE-FIGURE-018`
- affected files:
  - `SKILL.md`
  - `references/thesis/thesis-figure-generation-rules.md`
  - `references/thesis/figure-rules/baseline-and-sourcing.md`
  - `references/rule-owner-map.json`
  - `assets/mechanical-cad-acceptance-template.md`

## Trigger

The user reported that a mechanical graduation-design CAD drawing package still had overlapping text and isolated annotation text that did not visibly bind to any drawing content after a prior package-level PASS.

## Durable Rule

For mechanical CAD graduation-design tasks, every visible CAD text entity must be owned by a dimension, leader, balloon, datum/roughness/tolerance symbol, view/detail/section label, title-block/BOM/table cell, technical-requirement note block, or documented local feature list. Unowned free text, unsupported floating text, unbound scattered text, dimension-like text without anchor, and user-reported overlap/crop defects are hard blockers until the delivered DWG/DXF/PDF/PNG sheets are regenerated and audited against the exact final package SHA256. The same rule family also locks CAD text as normal readable CAD text, requires table/title-block/BOM cell padding evidence, strict 0.5 mm / 0.25 mm source lineweight evidence when requested, and current DXF/PNG/regeneration-manifest SHA binding so old PASS evidence cannot close a new drawing complaint.

## Enforcement Notes

The active run checklist, task cards, rendered-review JSON, mechanical CAD acceptance record, and final package audit must all expose the orphan/annotation-binding surface. Old rendered-readability PASS records become stale when the user reports remaining overlap or orphan text.
