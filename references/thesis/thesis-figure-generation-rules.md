# Thesis Figure Generation Rules

Use this file as the routing layer whenever a thesis task includes figures, screenshots, tables, charts, or diagram replacement.

## Scope

Apply the thesis figure rule set in all thesis-related modes:

- `program-plus-thesis`
- `thesis-only`
- `format-repair-only` when figure or table work is involved
- `skill-maintenance` when the figure or mechanical-drawing rules, validators, templates, or gates are being changed

## Default Rule

Thesis figure generation is not optional polish. If the thesis topic, template, or chapter semantics imply that figures should exist, treat missing figures as an unfinished state.

## Activation Rule

- Activate the thesis figure lane before any DOCX mutation whenever a teacher comment, user feedback item, existing caption, nearby prose, planned replacement paragraph, generated asset, or helper-script output mentions or implies a figure, screenshot, chart, diagram, flowchart, ER/database diagram, architecture/module diagram, use-case diagram, sequence diagram, figure readability issue, wrong image, image order issue, caption mismatch, or redraw/replacement request.
- Comment-driven thesis revision must convert each figure-related comment into a figure inventory row and a per-figure task card before editing the manuscript. The task card must preserve the comment id or anchor, target caption/location, inferred figure family, planned source path, sample lock, pre-insertion evidence path, post-insertion rendered evidence path, and final status.
- Broad thesis rewriting, template alignment, or format repair does not exempt incidental generated images from this lane. If an image is generated, replaced, recaptured, or rebound in the DOCX, classify it as runtime screenshot, code screenshot, data chart, or structural figure and record it in the figure plan.
- If user feedback says all figures, a figure class, or a figure set is noncompliant, route the run as a whole-set figure defect. The active run must inventory every figure in the current thesis copy and create one task card per affected figure before any DOCX mutation.
- Raster data charts may use charting libraries when the chart is genuinely data-driven and style-reviewed. Structural figures such as flowcharts, architecture diagrams, ER diagrams, use-case diagrams, sequence diagrams, and module trees must use the structural-figure production path rather than ad hoc raster drawing.
- If a final thesis DOCX contains figure captions or embedded images but the active run has no figure inventory or manifest, the run is incomplete unless the inventory explicitly marks every figure surface as not-applicable with evidence.

## Sequential-Diagram Routing Rule

- If a figure title, caption, or nearby thesis prose describes ordered stages, handoff sequence, evidence chain, processing chain, result chain, or any other stepwise left-to-right / top-to-bottom progression, route that figure into the flowchart family by default.
- Do not treat a sequential `示意图`, `证据链图`, `链路图`, or similarly named figure as exempt from flowchart rules just because the caption does not literally contain the word `流程图`.
- When a figure semantically represents a sequence of steps, it must inherit the active flowchart sample lock and the corresponding flowchart review gates.

## Child Files

- `references/thesis/figure-rules/baseline-and-sourcing.md`: figure scope, source priority, required figure families, baseline style, and completion/failure rules
- `references/thesis/figure-rules/review-gates.md`: style review, cleanliness review, text legibility, sample-lock, and insertion gates
- `references/thesis/figure-rules/geometry-and-layout.md`: connector geometry, ER/tree layout, text containment, and collision-prevention rules
- `references/thesis/figure-rules/workflow-and-checklists.md`: structural-figure SOP, preflight, and final acceptance checklists
- `references/user-feedback/thesis-workflow.md` rule `FB-THESIS-010`: whole-set figure defects, full inventory, per-figure task cards, replacement relationship evidence, and rendered-page closure
- `references/rule-owner-map.json`: durable figure rule owner index and validator/selftest coverage map

## Loading Rule

- Load only the child files relevant to the current figure subtask instead of bulk-loading every figure rule file.
- For a normal structural-figure task, the default load set is:
  - `baseline-and-sourcing.md`
  - `review-gates.md`
  - `geometry-and-layout.md`
  - `workflow-and-checklists.md`
- For a current-user `material-only-reuse` instruction such as "only pull images from the material document", "do not redraw", "do not generate", or "use the supplemental DOCX only for missing figures", load:
  - `baseline-and-sourcing.md` for `CORE-FIGURE-019` and `CORE-FIGURE-020`
  - `workflow-and-checklists.md` for the material-only manifest and acceptance rows
  - do not load draw.io production as an execution path unless the user later explicitly cancels material-only reuse
- For a runtime screenshot-only task, load `baseline-and-sourcing.md` and `review-gates.md` first, then add other child files only if the task expands into structural-figure repair.
- For mechanical CAD, drawing-package, source-linework, CAD color-family differentiation, redraw/different-from-baseline, official-CAD-command requirements, CAD open-view structural coherence complaints, teacher/reference drawing alignment, CAD text mojibake/missing-glyph boxes, upside-down/mirrored CAD text, text/entity overlap, orphan/free-floating CAD text, unbound scattered labels, external CAD case references including jixie5-style same-type drawing pages, or PDF-only-change complaints, load `baseline-and-sourcing.md` before CAD mutation and route `CORE-FIGURE-010` through `CORE-FIGURE-018` as applicable. In particular, `CORE-FIGURE-013` / CAD Source Linework Differentiation requires a source-linework differentiation audit, PDF-only change rejection, minor-entity-move-only rejection, and source-to-render derivation evidence; `CORE-FIGURE-018` requires annotation ownership evidence for leaders, dimensions, balloons, note labels, and any dimension-like or scattered text.
- For mechanical CAD, drawing-package, source-linework, CAD color-family differentiation, redraw/different-from-baseline, official-CAD-command requirements, CAD open-view structural coherence complaints, teacher/reference drawing alignment, CAD text mojibake/missing-glyph boxes, upside-down/mirrored CAD text, external CAD case references including jixie5-style same-type drawing pages, or PDF-only-change complaints, load `baseline-and-sourcing.md` before CAD mutation and route `CORE-FIGURE-010` through `CORE-FIGURE-017` as applicable. In particular, `CORE-FIGURE-013` / CAD Source Linework Differentiation requires a source-linework differentiation audit, PDF-only change rejection, minor-entity-move-only rejection, and source-to-render derivation evidence.

## Parent Boundary

- Keep this parent file short and routing-oriented.
- Do not duplicate detailed figure constraints here once they have been moved into child files.
- If a child file becomes too heavy, split it further and update `SKILL.md`, `FILE-ROLE-INDEX.md`, and validation logic in the same turn.
