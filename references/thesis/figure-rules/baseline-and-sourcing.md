# Thesis Figure Rules: Baseline And Sourcing

Use this file for figure scope, priority order, required figure types, sourcing rules, and baseline style expectations.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current figure subtask.
- Apply this file together with `references/thesis/thesis-figure-generation-rules.md`.

## Scope

Apply these rules in all thesis-related modes:

- `program-plus-thesis`
- `thesis-only`
- `format-repair-only` when figure or table work is involved

## Default Rule

Thesis figure generation is not optional polish. If the thesis topic, template, or chapter semantics imply that figures should exist, treat missing figures as an unfinished state.

## Priority Order

When deciding how a figure should look, use this order of authority:

1. user-provided figure samples or screenshots
2. actual figure/image samples inside the active school template or accepted sample thesis
3. current school template or accepted sample thesis rules for caption, image-holder, spacing, and pagination
4. stored skill samples and stable thesis format rules in this skill
5. conservative default academic diagram style

Do not skip a higher-priority visual source just because a lower-priority default already exists.

### CORE-FIGURE-005. Template Figure Samples Precede Skill Samples (Mandatory)

- Before generating, replacing, resizing, or approving thesis figures, inspect the active school template and accepted sample thesis for real figure or image examples, including the image body, image-holder paragraph, caption paragraph, spacing around the figure block, page occupancy, and whether one or multiple figures appear on a page.
- If a template or accepted sample contains any usable figure/image example, treat that example as the mandatory figure-format baseline for page occupancy, figure width, internal label readability, line weight, border style, caption separation, image-to-caption spacing, and figure-block pagination.
- Skill-internal visual samples under `references/visual-style-samples/figures/` are fallback sources only after the current run records that the active template and accepted sample thesis do not contain a usable figure/image example for the target family or page class.
- A figure task card, figure manifest, or review evidence record must name either the template/sample figure baseline that was used or the documented no-template-figure-sample verdict plus the exact skill-internal fallback sample path.
- Do not claim a figure set is compliant when the run used a skill-internal sample without first checking and recording whether the active template or accepted sample thesis contains a usable figure/image example.

### CORE-FIGURE-006. Existing Figure Replacement Requires Explicit Authorization And Source-To-Final Binding (Mandatory)

- Existing thesis figures are preserve-by-default surfaces. Do not redraw, replace, recapture, or rebind an existing embedded image merely because a broad format pass, comment pass, or helper script is already touching nearby content.
- A replacement or redraw is allowed only when the current user instruction, teacher comment, or approved figure task card explicitly authorizes that exact figure family, caption, or anchor.
- The figure manifest row must record `mutation_intent`, explicit replacement authorization source and scope, original/final relationship id, original/final media target, original/final media SHA256, target anchor caption, target anchor chapter, protected-surface location verdict, caption-to-asset binding, rendered evidence, and final DOCX relationship evidence.
- If the original media hash, relationship id, target caption, or media target changes without this source-to-final binding, the figure task fails even when the replacement image looks visually plausible.
- Changing a DOCX drawing object is also a figure mutation even when the media file is unchanged. The manifest must authorize and bind original/final drawing object SHA256 plus owner part before accepting changes to `wp:extent`, `wp:inline`/`wp:anchor`, relationship id set, or image-to-caption adjacency.
- The manifest itself must bind the compared documents with top-level `source_docx_path`, `source_docx_sha256`, `final_docx_path`, and `final_docx_sha256`. A sidecar task card, acceptance summary, or final-only path is not enough.
- The replacement check applies to the full Word package, not only `word/document.xml`. Image relationships in headers, footers, footnotes, endnotes, comments, and other story parts are preserve-by-default and must be authorized with the same original/final rid, target, SHA256, and owner-part binding.
- A newly inserted thesis image also requires explicit insertion authorization, final rid/target/SHA256 binding, caption-to-asset mapping, and proof that the target anchor is not in cover, declaration, abstracts, keywords, TOC, headers, footers, or page-number protected surfaces unless the official template owns that image surface.
- If source and final DOCX media relationship manifests differ, the run is an image mutation even when the task text only says `format repair`, `content cleanup`, or similar wording.
- If source and final DOCX drawing object manifests differ by size, inline/anchor mode, relationship set, or caption adjacency, the run is an image mutation even when media relationship manifests are identical.
- Structural replacement figures must still use draw.io source, SVG export, and raster fallback. Mermaid, Pillow/PIL, hand-drawn PNG, fixed-coordinate quick drawings, or other draft-only raster paths cannot be accepted as final structural figure sources.

## Figure Types That Usually Must Be Covered

When relevant to the thesis, generate or capture these by default:

- overall system architecture diagram
- module or functional structure diagram
- business flowchart
- use-case diagram
- sequence diagram
- ER diagram / entity-relationship figure / database entity-relationship figure
- data analysis chart
- real program screenshots for implementation pages
- testing or runtime evidence screenshots when the chapter needs them
- real code screenshots when the chapter or teacher comment explicitly requires code-first explanation

## Inventory And Manifest Rule

### CORE-FIGURE-003. Condition-Driven Required Figures Must Have Checklist Rows (Mandatory)

- Every thesis figure surface in the active scope must have one inventory row before manuscript mutation. Sources include teacher comments, user feedback, existing captions, planned captions, generated assets, replacement screenshots, and helper-script image outputs.
- The inventory must also include condition-driven required figures inferred from the thesis body and project evidence, not only figures already present in captions. For example, a database design chapter that contains database concept design, main entities, or entity-relationship prose requires a `database_er_diagram` row even when the current manuscript has no ER caption yet.
- Missing required figure rows are hard failures. A required figure row may not be hidden behind a generic architecture figure, a database table, a prose-only entity list, or a broad `figure set complete` statement.
- The inventory row must state the figure id, comment id or source anchor when any exists, chapter/location, inferred family, source kind, chosen style source, source asset paths, task-card path, pre-insertion evidence path, post-insertion rendered evidence path, final DOCX relationship evidence, status, and skip reason.
- Do not allow a figure plan to pass with a generic `figure set complete` summary. Each affected figure and each planned family must have its own status.
- Thesis generation and revision must create and validate a figure asset manifest before DOCX assembly whenever the final DOCX will contain embedded images or figure captions. If no figure surfaces are in scope, record the no-figure reason explicitly instead of leaving manifest fields blank.
- The figure asset manifest must carry a `required_figures` checklist when condition-driven figures are inferred. Each required row must name the requirement id, family, expected caption, status/verdict, task card, post-insertion rendered evidence, and final DOCX relationship evidence.

## Whole-Set Figure Defect Rule

- If the user reports that all figures, a group of figures, or the thesis figure set does not meet requirements, expand the scope to every figure in the active thesis copy unless the user explicitly narrows it.
- The inventory must include every formal caption and every embedded image relationship in that active copy. Do not rely on a sample page, old exported PNG folder, or a prior figure list without re-reading the target DOCX.
- Each affected figure must have its own task card, replacement asset path, manifest entry, final DOCX relationship evidence, rendered-page evidence, and pass/fail/skipped verdict.
- A whole-set figure repair cannot pass from old image reuse, screenshot existence, caption correctness, DOCX validity, PDF export success, or a single successful replacement.
- Final handoff must name the exact final DOCX/PDF path and state whether structural figures were embedded as SVG-primary or PNG fallback for renderer compatibility.

## Generation Rules

- If the user provides a figure sample, visually match that sample first.
- If the figure is a runtime UI image, prefer a real screenshot from the running system.
- Do not use schematic UI mockups, placeholder diagrams, manually redrawn pages, or `示意图` images as substitutes for runtime screenshots.
- For Streamlit or other browser-based program pages, the default screenshot path must be a real Google Chrome render driven through Chrome DevTools Protocol full-page capture.
- Do not default to one-shot `chrome --headless --screenshot`, Edge headless screenshot, fake evidence-page generation, or static redraws when the requirement is a real program screenshot.
- The preferred runtime-page capture pattern is:
  1. start the real program page locally
  2. open the target page in real Google Chrome
  3. wait for the target page's key DOM content and charts to finish rendering
  4. capture the full page through Chrome DevTools Protocol rather than a blind command-line screenshot
  5. visually inspect the output before insertion
- If Chrome DevTools Protocol full-page capture is available, treat it as the mandatory default screenshot source for thesis runtime pages unless the user explicitly requests another path.
- For system-description or implementation chapters, the project must have authentic full-page system screenshots available as evidence assets whenever those chapters describe real system pages or flows.
- These authentic full-page screenshots do not have to all be inserted into the thesis. Selection still depends on chapter semantics, template limits, and user direction.
- Do not treat cropped fragments, mockups, design placeholders, or blank exports as substitutes for authentic full-page system screenshots.
- If the real system cannot be started or the target page cannot be reached, stop treating the screenshot task as complete; record the blocker explicitly instead of silently replacing it with a schematic image.
- For desktop GUI programs, do not downgrade the real-screenshot requirement to `widget.render`, `widget.grab`, `QWidget.grab`, offscreen widget painting, canvas export, or any component-only snapshot. Use a real visible-window or OS/desktop capture path and record the capture rectangle/window title evidence before insertion.
- For chapter-4 or equivalent implementation screenshot families, lock a route-to-caption mapping before insertion:
  - caption text
  - route URL
  - readiness cue
  - accepted screenshot asset path
  - embedded media relationship after insertion when replacing an existing image
- If that map is missing, the screenshot family is not ready for final insertion because wrong-asset rebinding can silently put a structural figure into a runtime screenshot slot.

## Code Screenshot Rule

- Treat code screenshots as a separate figure family from runtime page screenshots.
- When the chapter, teacher comment, or user direction requires code screenshots, the source must be real project code rendered from an actual editor or code pane.
- Do not use synthetic code cards, poster panels, pseudocode mockups, or LLM-styled illustration panels as thesis code screenshots.
- Default accepted code screenshot characteristics:
  - real code from the current project
  - syntax highlighting when the chosen editor provides it
  - visible line numbers
  - enough surrounding code context to read a coherent fragment rather than only a tiny isolated snippet
  - no manual in-image title such as `关键代码`, `核心代码`, `对应 4.3.x`, or similar explanatory labels
- If the user or approved sample requires code-pane-only crops, crop away editor window chrome and keep only the real code pane with line numbers.
- Do not treat a code-pane-only crop as violating the runtime full-page screenshot rule; code screenshots and runtime screenshots are different figure families with different acceptance surfaces.
- If a figure has been upgraded from a sample or schematic image into a real screenshot, update the caption text, alt text, and nearby thesis prose in the same pass so they no longer call it `示意图` or `样例图`.

## Algorithm Result Figure Rule

### CORE-FIGURE-004. Algorithm Detection And Recognition Result Figures Must Be Authentic (Mandatory)

- Treat YOLOv8 detection results, DBNet text-detection results, CRNN/OCR recognition comparisons, preprocessing result strips, detection-box overlays, recognition success/failure examples, and similar model-output figures as `algorithm_result` figures.
- An `algorithm_result` figure must come from one of these traceable sources:
  - a real program run from the current project
  - a saved result image produced by the current project or its experiment scripts
  - a user-provided real result screenshot or output image
  - a reproducible external evidence asset explicitly approved for the thesis
- The figure manifest must record the accepted result image path, caption-to-asset mapping, and at least one provenance source such as source/input image path, generation script path, model/output log path, existing result source, or user-provided asset evidence.
- An accepted algorithm-result image must contain readable rendered result content. A blank image, near-empty export, dominant solid-color block, or purple placeholder remains a failure even when provenance fields and pixel dimensions are present.
- Do not use a manually drawn composition, generic drug-box illustration, mock detector overlay, fake OCR panel, placeholder image, or builder-created `示意图` / `样例图` as a substitute for a required algorithm result figure.
- If the real system, model, result images, or user-provided output assets cannot be found, stop the algorithm-result figure lane and record a blocker. Do not silently downgrade the required image to a schematic figure.
- If a caption or nearby prose still calls an algorithm result figure `示意图`, `样例图`, `mock`, or `placeholder`, the figure can pass only when the manifest proves real provenance and records an explicit authenticity verdict. Otherwise the figure is failed.

## Structural Figure Rule

- If the figure is analytical or structural, draw it instead of faking it with placeholder images.
- If the thesis chapter is design-oriented, use design figures; do not replace them with runtime screenshots unless the user explicitly wants that.
- If the thesis chapter is implementation-oriented, prefer real interface or runtime evidence.
- For thesis entity-attribute figures or entity field-structure figures that sit beside Chinese database-design tables, prefer Chinese attribute labels that map one-to-one to the table semantics.
- Do not default those figure labels to raw English field names such as `room_id`, `inventory_date`, or `booked_stock` when the surrounding thesis tables and prose are presenting the same fields in Chinese.
- If the current user explicitly wants the figure to line up with a database table while still remaining Chinese-readable, use Chinese semantic labels first and keep primary-key hints or equivalent role hints only as lightweight supplements such as `（PK）` when needed.
- For final thesis structural figures, use draw.io source files and direct draw.io exports as the mandatory production path.
- For final thesis structural figures that originate from draw.io or another vector authoring path, insert them into the thesis through an SVG-primary DOCX embed path by default.
- Thesis generation must create and validate a figure asset manifest before DOCX assembly. The manifest must record family, source kind, draw.io path, SVG path, raster fallback, caption, chapter, insertion status, and rendered-page status for every planned figure.
- For ER and other dense structural figures, the manifest must also record `geometry_validation_report`, `source_scale_bbox_map`, `inserted_scale_geometry_evidence`, `dense_zone_crop_evidence`, and `collision_check_verdict`; these fields are acceptance evidence, not optional notes.
- For comment-driven, replacement, or incidental figure generation, the same manifest must also record task-card path and evidence paths for every planned or touched figure.
- Figure family is not only a manifest field supplied by the caller. The canonical figure contract must infer family from title, caption, and explanation text; wording such as `流程`, `链路`, `步骤`, `过程`, `处理`, `pipeline`, or `workflow` routes the figure to the flowchart contract even if the manifest declares `structure`.
- A structural thesis figure cannot be accepted from a plain `image_path` alone. It must have a draw.io source, an SVG export, and a raster fallback; the PNG may only serve as fallback evidence.
- Every final-DOCX figure or embedded media relationship must be covered by the structured `figure.scope-manifest-contract` detector in the exact `sample_self_check` report used for handoff. A manifest row must include the per-figure task card, caption-to-asset mapping, final-DOCX relationship evidence, post-insertion rendered-page evidence, insertion status, and rendered-page status. Missing evidence, a stale source image that was never inserted, or a DOCX image without a manifest row blocks acceptance.
- An ER structural thesis figure cannot be accepted until the canonical figure contract validates its draw.io source against `scripts/validate_structural_figure_geometry.py` and the geometry report verdict is `pass`.
- A PNG fallback or renderer-generated fallback image may exist inside the DOCX package, but it must not be the only surviving insertion path for a structural figure when an approved SVG source exists.
- Do not treat Mermaid, ad hoc Pillow drawing, fixed-coordinate scripts, or any other quick-diagram path as an acceptable final-source substitute for thesis structural figures.
- When a structural figure family already has a stored sample, a draw.io path, or a draw.io-first requirement, Mermaid and other quick-diagram syntaxes may be used only for planning drafts and must not be accepted as the final production asset for that family.
- White background is mandatory for thesis structural figures unless an approved sample explicitly overrides it.
- Do not use gray box fills, tinted node fills, pseudo-title bars, or other non-white decorative panels in final thesis structural figures.
- All text in thesis figures must be black unless a stronger approved sample explicitly overrides that.
- Text inside thesis figures must prioritize legibility over density or compactness.
- If text becomes hard to read at expected thesis page size, enlarge the layout, reduce local density, or wrap labels before acceptance.
- Do not place figure titles, figure numbers, or caption-like text inside the image canvas.
- Do not place explanatory note boxes, summary panels, interpretation bullets, or other prose explanation blocks inside the image canvas; move that explanation into the thesis body text instead.
- Treat `in-image title plus external DOCX caption` as a hard acceptance failure.
- When doing thesis format alignment, do not stop at caption position and page placement. The internal style of the figure itself must also match the approved sample or template.
- Internal style alignment must cover:
  - text font family and visual weight
  - text size hierarchy
  - stroke and connector thickness
  - node border style
  - fill usage
  - spacing and padding rhythm
  - overall visual density

## Architecture Sample Lock Rule

For thesis architecture-family figures, if the user has provided the grouped layered architecture screenshot recorded in:

- `references/visual-style-samples/figures/figure-layered-architecture-sample-20260419.jpg`

then that screenshot is the mandatory default style source unless the current run provides a newer stronger architecture sample.

Treat the following figure intents as part of this architecture family when no more specific stronger sample overrides it:

- overall system architecture diagrams
- data-acquisition architecture diagrams
- grouped module architecture diagrams
- layered processing architecture diagrams

### Current Recorded Architecture Sample Characteristics

The locked grouped architecture sample uses this visual grammar:

- one large outer boundary frame
- one short top-centered system or architecture label inside the outer frame when the sample shows that label
- one centered inner grouped frame that contains the main module set
- optional narrow left and right side pillars for supporting modules
- several aligned inner rectangles for core modules
- one or more wide horizontal layer boxes for data, processing, or output layers
- monochrome black-and-white academic styling
- white background and thin black borders
- short centered labels
- minimal or no connector emphasis unless the relation cannot be conveyed by grouping

### Mandatory Visual Grammar For The Locked Architecture Sample

The default thesis architecture figure must follow this visual grammar:

- use grouped containment first and connectors second
- preserve the outer-frame plus inner-grouped-frame composition rather than converting the figure into a loose tree
- keep left and right side pillars when the content naturally maps to side supporting modules
- keep inner module rectangles aligned in clean rows or columns
- keep wide horizontal layer boxes centered and visually subordinate to the outer frame
- allow at most one short top-centered internal system/container label when the sample requires it
- do not place figure numbers, figure captions, or long explanatory text inside the canvas
- keep the entire figure monochrome and white-background
- keep borders thin and consistent across the whole figure

### Hard Failure Rule For Thesis Architecture Figures

Treat an architecture figure as failing style lock when any of these are true:

- it falls back to a root-node tree or arrow fan-out layout even though the grouped layered sample could express the same content
- it omits the outer frame or the centered grouped core area and becomes a loose set of unrelated boxes
- it adds colored fills, shadows, decorative icons, title bars, or presentation styling
- it places more than one internal title-like banner or uses a long sentence as the top label
- it uses arrows as the main visual grammar when containment and alignment should be doing the explanatory work
- side pillars, layer boxes, or grouped modules drift out of alignment and no longer resemble the locked sample family
- the resulting figure looks like a software slide or generic block diagram rather than a grouped textbook architecture figure

### Architecture Internal-Label Exception

The generic thesis figure rule against in-canvas titles still applies to figure numbers, captions, and banner-style titles.

For this locked architecture family only, one short top-centered system/container label inside the outer frame is allowed when:

- the approved sample itself uses that label position
- the label is a system or architecture name rather than a figure caption
- the label remains visually integrated with the outer frame and does not become a decorative title bar

## Completion Rules

A figure task is not complete until all of these are true:

- the figure matches the intended chapter semantics
- the figure style matches the approved visual source
- every planned structural figure family has its own resolved status instead of being collapsed into a single generic "diagram complete" claim
- the caption text is correct

### CORE-FIGURE-010. Mechanical CAD Drawing Packages Must Match The User-Provided Package Standard (Mandatory)

When a graduation design task is for a mechanical product, machine part, assembly, conveyor, transmission, hydraulic mechanism, tooling fixture, or similar engineering deliverable and the user provides a CAD/PDF drawing package as the sample authority, treat that package as the drawing baseline rather than as decorative reference material.

The run must inspect and record the sample package before producing drawings:

- package structure, folder names, file extensions, and naming pattern
- real source format such as DWG, DXF, DWT, PDF, or mixed exports
- sheet size classes such as A0/A1/A2/A3 and total drawing workload
- border, coordinate grid, title block, revision block, materials/BOM tables, technical-requirement blocks, and signature fields
- lineweight family, hidden/center/section/hatching usage, dimension-density level, balloons, weld/roughness/tolerance symbols, local detail views, sectional views, and notes
- whether the sample is an editable CAD-source package or a rendered PDF-only package

Sketch-style vector drawings, sparse schematic redraws, raster-only illustrations, or simplified "CAD-like" pictures do not satisfy a mechanical graduation drawing package when the sample shows dense manufacturing/assembly drawings. If the required final source format is DWG and the environment has no AutoCAD, ODA, QCAD Professional, DraftSight, BricsCAD, or equivalent DWG writer/converter, the run must record DWG generation as a blocker and may only hand off DXF/PDF as an explicitly labeled substitute. Do not rename DXF, SVG, PDF, or PNG files as `.dwg`, and do not claim "same format as the sample" unless the handed-off package contains real files in the same source format family.

When the current run does not contain a stronger user-provided mechanical CAD package, the mandatory default mechanical CAD baseline is the stored SGB620/80T passed package baseline under `references/visual-style-samples/mechanical-cad/default-mechanical-cad-baseline-sgb62080t.md`, together with `sgb62080t-sheet-11-default-baseline.png`, `sgb62080t-sheet-13-default-baseline.png`, `sgb62080t-sheet-16-default-baseline.png`, and `sgb62080t-default-baseline-audit-v5.json`. This fallback must be used as both the visual baseline and the complexity baseline. A stronger current-run user-provided sample wins over the stored fallback baseline.

For a three-sheet-A0-equivalent mechanical drawing requirement, the output must show academic/manufacturing drawing depth, not only occupy A0 page area. Acceptance evidence must include a package manifest, rendered-sheet previews, a rendered no-overlap review for drawing frames, title blocks, BOM/material tables, technical-requirement tables, notes, dimension text, and view geometry, plus explicit rendered verdicts for boundary clearance, detail density, title-block/table/notes isolation, annotation-margin clearance, and local crowding; readable text/font evidence, sheet-size/workload evidence, and a source-format capability verdict are still required.
- When DWG/PDF/DXF drawings are handed off, run `scripts/audit_mechanical_drawing_package.py` on the exact final package or drawing folder. If the user provided CAD/PDF samples, pass them as `--reference` inputs and bind the report path, package path, package SHA256, verdict, and any density/encoding/CAD-structure blockers before claiming the drawing package matches the sample.
- For sample-based mechanical CAD work, use distributed density gates rather than only total file counts: per-DXF entity minimums, dimensions spread across multiple sheets, per-PDF drawing-object density by sheet size, required engineering tokens such as sectional/detail/technical-requirement/BOM/weld/roughness terms, and a DWG byte-density ratio against the user-provided CAD source package. A single dense strip, one overloaded sheet, or a normal Chinese character falsely detected as mojibake cannot clear or block acceptance by itself; the machine gate must expose those details for the audit lane to review.
- For this error class, the audit report must use the v4 mechanical package schema or a stricter successor and expose `density_verdict`, `manufacturing_depth`, and `rendered_review_verdict`; a pass that lacks dimension totals, geometry totals, DWG reference ratio, block/arc/hatch/text depth, rendered-review fields, per-sheet density failure details, or any clearance/isolation verdict that still leaves title blocks, tables, notes, or annotations crowded is not a final acceptance record.
- For this error class, entity counts, dimension counts, byte ratios, file counts, and PDF parse counts are necessary but never sufficient. The package manifest must include `rendered_review` evidence with real preview paths, per-sheet rows, `no_overlap_verdict=pass`, `boundary_clearance_verdict=pass`, `detail_density_verdict=pass`, `title_block_table_notes_isolation_verdict=pass`, `annotation_margin_clearance_verdict=pass`, `local_crowding_verdict=pass`, `text_legibility_verdict=pass`, `sheet_layout_verdict=pass`, `manufacturing_view_depth_verdict=pass`, and `entity_count_only_verdict` set to a rejection value such as `not-used`. A package that passes only from entities or dimensions is an entity-count-only false-pass and must be blocked.
- The exact literal keyword `annotation margin clearance` must remain present in the routed mechanical CAD rule text and validator-facing documentation so `CORE-FIGURE-010` owner-map keyword checks can bind this rendered-review verdict without ambiguity.
- If rendered previews show sketch-like simplified outlines, missing manufacturing views, missing section/detail/hatching/local-detail depth, overlapped text, overlapped dimension labels, crowded notes, annotation text that sits hard against the drawing frame, tables colliding with the title block or drawing frame, clipped borders, unreadable annotation text, or any frame/title-block/table overlap, the mechanical package fails even when the static CAD entity-density gate passes.
- When only one or two sheets fail the distributed detail-density gate, prefer adding real mechanical detail inside the existing main view or section envelopes of the failing sheet before adding new floating panels, new detached tables, or extra note boxes. Safe first choices are repeated hidden/center lines, guide-slot lines, bolt-hole rows, hatch/section refinements, and reducer/chain/seat internal structure lines that stay inside an already reserved view box. Avoid clearing the gate by stuffing new blocks into page gaps if that move increases local crowding risk near title blocks, BOM tables, technical-requirement notes, or dimension bands.
- the numbering is correct
- the placement is correct
- the figure is actually embedded in the final thesis file
- the figure remains fully visible and readable on rendered thesis pages after insertion scaling

## Failure Rules

Do not mark figure work complete when any of these still happen:

- required figures are still missing
- only one surviving flowchart exists while other planned structural figure families such as architecture or ER remain unresolved
- the figure style visibly deviates from the approved sample
- a screenshot is used where a design diagram is required
- a design figure is used where runtime evidence is required
- captions and figures do not match
- the source image exists but the final thesis file does not contain the real embedded figure
- a structural figure has an approved SVG source, but the final thesis file still embeds it only as a PNG without an SVG primary path
- a runtime screenshot caption points to one page or module while the embedded image clearly shows a different route or module
- two different runtime screenshot captions share one stale embedded media asset without explicit approval
- authentic full-page system screenshots are required by chapter semantics, but the project does not actually have them as evidence assets
- a browser screenshot was taken through a known-bad skeleton-screen path even though Chrome DevTools Protocol full-page capture was available
- a code screenshot is still a synthetic panel, pseudocode card, or sample image rather than a real code capture
- a figure that now uses a real screenshot still carries caption or nearby wording that falsely calls it `示意图` or `样例图`
- a generic Mermaid quick diagram is being used as the final structural figure even though the current figure family is sample-locked or draw.io-first
- a Mermaid crow's-foot `erDiagram` with table-style field lists is delivered while the active ER sample lock still expects Chen-style entity rectangles, relation diamonds, and attribute ellipses
- any node or panel uses gray fill when the approved figure family is white-background academic style
- any inner box, node, or child panel visibly extends beyond its parent frame, layer box, or outer system boundary
- any connector passes through an unrelated box, ellipse, diamond, or text label
- any connector is drawn by center-to-center approximation instead of boundary-to-boundary geometry
- the figure contains an internal title even though the thesis uses external captions
- the figure is embedded but the rendered page still clips the figure because the image paragraph inherits unsafe body-text spacing
- the figure contains both an internal title and an external thesis caption
- the figure text is so cramped, stacked, or overlapped that a normal reader cannot clearly read it at thesis page size

## Practical Review Checklist

Before closing a thesis task with figures, check:

- figure list against chapter needs
- each figure against the approved style source
- each figure against its caption
- each figure against chapter semantics
- final document embed success
- final visual layout after insertion

## Flowchart Sample Rule

When the user provides a thesis flowchart sample, treat it as a direct style target.

When the user has not provided a stronger newer flowchart sample for the current run, the mandatory default flowchart style source is:

- `references/visual-style-samples/figures/figure-flowchart-vertical-sample-01.svg`

Treat this file as the locked fallback flowchart source-of-truth, not as a loose inspiration board.

### Current recorded flowchart sample characteristics

- monochrome black-and-white academic style
- top-down vertical layout
- a single centered vertical main chain is the default composition for sequential flows
- rounded-rectangle terminators for start and end nodes
- square-corner process rectangles with thin-to-medium dark borders
- a single centered diamond decision node below the start node
- centered Chinese text inside nodes
- explicit branch labels such as `真` and `假` placed beside the outgoing decision branches
- straight vertical connectors and simple horizontal split / reconverge connectors
- left-right branch symmetry with one clean merge lane before the downstream central step
- no gradients, icons, shadows, or presentation-style decoration
- no in-canvas title, note box, legend, or explanatory prose

### Mandatory Visual Grammar For The Locked Flowchart Sample

The default thesis flowchart must follow this visual grammar:

- the main reading path must move straight downward along one centered vertical axis
- start and end nodes use rounded rectangles rather than plain rectangles or ellipses
- decisions use a single centered diamond rather than a generic box or prose callout
- process steps use plain white rectangles with uniform dark borders
- arrows remain dark, straight, simple, and orthogonal where possible
- branch labels remain short and sit outside the diamond near the outgoing paths
- branch results reconverge cleanly before the downstream central step when the logic structure converges
- the composition remains vertically readable and textbook-like rather than UI-like or infographic-like
- labels remain short, centered, and Chinese-first unless the source artifact is inherently English
- if the business logic is sequential rather than branching, draw it as one vertical start-to-end chain instead of a horizontal or snake-like multi-row route

### Hard Failure Rule For Thesis Flowcharts

For thesis flowcharts, treat the following as style failures even if the semantics are correct:

- using colored fills, gradients, shadows, icons, title bars, or decorative panels
- spreading one sequential process across two horizontal rows or a snake-like left-right route when one vertical chain would fit
- placing the first process node off to the left or right while the start node sits elsewhere, breaking the centered top-down reading path
- drawing start or end as plain rectangles when the locked sample uses rounded terminators
- replacing the central decision diamond with a rectangle, sentence block, or freeform callout
- putting `真 / 假` style branch outcomes into large note boxes instead of simple branch labels
- letting branches diverge asymmetrically without a content-driven reason
- skipping the merge lane when the business logic clearly converges
- adding an in-image title such as `流程图` or `业务流程图`
- producing a flowchart that looks like a slide graphic or software block diagram instead of a textbook thesis flowchart
- delivering a sequential `证据链图`, `链路图`, `结果链图`, or similarly staged figure as a loose connection diagram instead of a real flowchart
- inserting a flowchart at a size that makes the rendered page show only part of the figure or pushes the actual figure body onto the next page while leaving a large white gap on the previous page

Use this as the default fallback flowchart style when the user has not provided a newer or stronger sample.

## Use-Case Diagram Sample Lock Rule

When the task is a thesis use-case diagram and the user has not provided a stronger current-run sample, the mandatory default style source is:

- `references/visual-style-samples/figures/figure-use-case-diagram-sample-01.svg`

Treat that sample as the locked fallback visual family for actor, use-case ellipse, line weight, and overall sparsity, but apply the current durable default-route override below unless a stronger current-run source explicitly wins.

### Current Recorded Use-Case Diagram Characteristics

- monochrome black-and-white academic style
- white background with no decorative fills, gradients, or shadows
- actors drawn as simple stick figures outside the system area
- use cases drawn as plain white ellipses with thin black borders
- labels are short Chinese nouns or verb phrases
- connectors are simple thin black lines with plain arrowheads by default
- the diagram stays airy and avoids dense crossing

### Additional Boundary Rule For Thesis Use-Case Diagrams

- default thesis use-case style is no outermost system boundary rectangle
- unless the user explicitly asks to restore a boundary, or the active approved sample/template clearly requires one, do not add a surrounding system box just because UML tools expose it by default
- if a stronger current-run source still requires a system boundary rectangle, keep the boundary as a plain white rectangle with a thin black border
- do not turn the system boundary into a routing corridor that multiple actor lines stack through
- actor-to-use-case association lines must still preserve the same clean visual family as the stored sample even when a system boundary is present
- for sparse single-actor / single-column use-case layouts, direct fan-out arrow routes are the default target and should not be replaced by artificial trunk corridors or decorative orthogonal detours

### Hard Failure Rule For Thesis Use-Case Diagrams

Treat the figure as failing style alignment when any of these are true:

- a default run adds an outermost system boundary even though no stronger current-run source required one
- an actor-to-use-case association line passes through another unrelated use-case ellipse
- the bottom use cases on the student side or teacher side are connected by a line that visually cuts through a neighboring use-case ellipse
- several actor lines are stacked into one narrow shared lane along the system boundary and become hard to distinguish
- an association line starts from empty space instead of the actor contour, or its arrowhead does not terminate on the target ellipse boundary
- the figure preserves actors and use cases, but the association layout reads like a crowded wiring sketch rather than a thesis use-case diagram

## ER Diagram Sample Lock Rule

For thesis database-design figures, treat the following names as the same figure family and apply the same ER rules:

- ER diagram
- entity-relationship diagram
- database entity-relationship diagram
- 实体关系图
- 数据库实体关系图

When the task is a thesis ER diagram and the user has not provided a stronger ER sample for the current run, the mandatory default style source is:

- `references/visual-style-samples/figures/figure-er-diagram-sample-01.svg`

Treat this file as the locked fallback ER style source, not as a loose inspiration board.

### Current Recorded ER Sample Characteristics

The locked ER sample `figure-er-diagram-sample-01.svg` uses a thesis-style Chen ER grammar. Treat the following as the source-of-truth description of that sample:

- monochrome black-and-white academic style
- white background with no decorative fills, shadows, gradients, icons, or UI-like panels
- entity nodes drawn as plain white rectangles with thin black borders
- relationship nodes drawn as plain white diamonds with thin black borders
- attribute nodes drawn as plain white ellipses with thin black borders
- all connectors drawn as thin straight black lines
- cardinalities shown as simple `1` and `n` labels placed close to the relevant relationship lines
- Chinese labels centered inside shapes
- labels are short nouns or verbs such as entity names, attribute names, and relationship names
- entity boxes contain only the entity name, not a table-style field list
- attributes are detached from entities as separate ellipses instead of being embedded inside the entity rectangle
- relationships are expressed by diamonds rather than by direct entity-to-entity labeled lines
- overall composition is centered, airy, and diagrammatic rather than dense, tabular, or schema-like

### Mandatory Visual Grammar For The Locked ER Sample

The default thesis ER diagram must follow this visual grammar:

- Chen-style ER notation rather than table-card, schema-card, or UML class-card notation
- entity-relationship figures are ER figures; do not redraw an `实体关系图` as a table-field box sketch just because the caption contains `数据库`
- entity rectangles contain only the entity title
- each important attribute is drawn as a separate ellipse when shown
- each important relationship is drawn as a separate diamond
- cardinalities are written as `1` / `n`, not as prose notes
- use concise labels and limited attributes so the figure keeps the same light visual density as the stored sample
- keep the figure as a pure ER diagram; do not turn it into a mixed `ER + explanatory memo` canvas
- if too many entities or attributes would break the sample's visual grammar, split the content into multiple ER figures or move secondary explanation into the thesis body text

### Recorded Deviation Pattern From Failed ER Repairs

A recurrent failure occurs when the agent preserves semantic corrections but drifts away from the stored ER sample's visual grammar.

This produces figures that may be semantically closer to the codebase but still fail style alignment because they:

- look like simplified database schema cards instead of a thesis ER diagram
- connect entities directly by labeled lines without relation diamonds
- place explanatory prose boxes inside the figure as a substitute for proper ER notation
- keep too many field names inside rectangles and therefore read like table-definition snapshots
- visually resemble a custom engineering note more than the stored thesis sample
- keep all semantics in one overloaded figure instead of preserving the sample's restrained density

### Hard Failure Rule

For thesis ER diagrams, treat the following as style failures even if the semantics are correct:

- drawing the figure as entity-to-entity table cards instead of rectangle-diamond-ellipse ER grammar
- omitting relationship diamonds where the sample uses them
- omitting attribute ellipses when the sample's figure family depends on them
- adding a large internal note or legend box that dominates the figure surface
- using dense field-list rectangles as the primary visual form
- producing a figure that a reviewer would reasonably classify as a data model sketch instead of a thesis ER diagram
- replacing short centered labels with long prose statements inside the figure
- solving density problems by collapsing attributes into entity rectangles instead of reducing scope or splitting the figure
- keeping a top title such as `ER图`, `实体关系图`, or `数据库实体关系图` inside the image canvas instead of using the external thesis caption

## Current SVG Template Assets

Prefer these stored SVG assets as reusable visual references when they match the target figure type:

- `references/visual-style-samples/figures/figure-er-diagram-sample-01.svg`
- `references/visual-style-samples/figures/figure-system-structure-tree-sample-01.svg`
- `references/visual-style-samples/figures/figure-use-case-diagram-sample-01.svg`
- `references/visual-style-samples/figures/figure-flowchart-vertical-sample-01.svg`
- `references/visual-style-samples/formulas/formula-layout-sample-01.svg`
- `references/visual-style-samples/toc/toc-style-sample-01.svg`

Use them as default fallback references when the user has not provided a stronger sample in the current task.
