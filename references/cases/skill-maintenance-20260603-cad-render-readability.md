# Skill Maintenance: CAD Render Readability

- date: 2026-06-03
- owner rule: `CORE-FIGURE-010`
- owner file: `references/thesis/figure-rules/baseline-and-sourcing.md`
- primary validator: `scripts/audit_mechanical_render_readability.py`
- gate validator: `scripts/validate_skill_gate_record_gate.py::check_mechanical_cad_acceptance_record`

## Trigger

The MB670 drawing repair run received repeated user feedback that final sheet PDFs had linework merged into unreadable clusters and text/table content still overlapped or was covered by drawing strokes. Prior lineweight and high-DPI preview fixes did not guarantee that the delivered standard sheets themselves were readable.

## Root Cause

The mechanical CAD rule chain had rendered no-overlap, local-crowding, frame, hatch, and text-legibility requirements, but CAD-only acceptance could still pass without a dedicated exact-final standard-sheet audit for:

- PDF text overlapping other text.
- PDF text covered by actual vector drawing strokes or filled small graphics.
- PNG/PDF rendered linework crowding in local sheet regions.

This left room for stale evidence, zoom-page-only evidence, or broad rendered-review summaries to mask defects in the standard drawing sheets.

## Change

Added and wired `scripts/audit_mechanical_render_readability.py` so the CAD acceptance record must bind a report from the exact final standard PDF and PNG renders with:

- `passed=true`
- `text_text_overlap_count=0`
- `text_graphic_cover_count=0`
- `severe_line_crowding_count=0`

The audit uses PDF text bounding boxes, actual drawing stroke/edge hitboxes instead of broad table rectangles, and PNG local-background ink density so dark CAD view backgrounds do not count as linework.

## Validation

Required validation for this maintenance change:

- `python scripts/audit_mechanical_render_readability.py --self-test`
- `python -m py_compile scripts/audit_mechanical_render_readability.py scripts/validate_skill_gate_record_gate.py`
- rule-owner JSON parse
- UTF-8 clean check
- `python scripts/validate_skill_gate.py --skill-root <skill-root>`
- exact project final standard-sheet render audit before handoff
