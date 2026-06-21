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

1. explicit current-user source instructions, including material-only reuse and no-redraw instructions
2. user-provided figure samples or screenshots
3. actual figure/image samples inside the active school template or accepted sample thesis
4. current school template or accepted sample thesis rules for caption, image-holder, spacing, and pagination
5. stored skill samples and stable thesis format rules in this skill
6. conservative default academic diagram style

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

### CORE-FIGURE-019. User-Directed Material-Only Figure Reuse Overrides Redraw (Mandatory)

- When the current user explicitly says that thesis images must be pulled only from a material document, source DOCX, supplied screenshot bundle, or a named supplemental document, the active run must enter `material-only-reuse` mode for that paper.
- In `material-only-reuse` mode, do not generate, redraw, reinterpret, trace, or rebuild figures, even when a caption or nearby prose has flowchart, workflow, architecture, ER, use-case, sequence, or other structural semantics.
- A material-only structural or flowchart figure is accepted as a source-preserved raster figure. The validator must check material provenance and final embedded-media binding instead of demanding draw.io, SVG, geometry, or native reconstruction evidence.
- Approved sources must be locked in priority order. The primary material document is searched first. A supplemental document may be used only for figure slots that are missing from the primary material source, and the evidence must record the missing-figure reason.
- Each material-only figure row must record `material_only_reuse=true`, `no_redraw_user_override=true`, the current user override, material source path and SHA256, inventory path, material anchor/caption/index, extracted image path and SHA256, final embedded media SHA256, source-match verdict, generated-substitute rejection verdict, and material-only reuse verdict.
- A draw.io file, SVG, generated PNG, AI-created bitmap, screenshot of a redrawn diagram, or draw.io wrapper containing an imported material image is not a valid substitute for this mode. If such an asset is present, the run must record it as rejected rather than use it for the final DOCX.
- If the run cannot prove that the final embedded image hash came from the primary material document or an allowed supplemental source, the figure fails even if it is visually clear.

### CORE-FIGURE-020. Horizontal Body Figures Default To Paragraph Width (Mandatory)

- For all thesis DOCX work, horizontal/landscape figures, wide screenshots, system-interface screenshots, and other readable raster figures must default to the正文 usable paragraph width.
- The expected width is the section text width, computed from page width minus left and right margins, not a hard-coded centimeter value.
- Preserve aspect ratio. A tall portrait figure may be constrained by safe page height, but a horizontal figure or system screenshot that can fit at paragraph width must not be left as a small centered thumbnail.
- Paragraph-width placement must not override native raster readability. Before enlarging a raster figure to the正文 usable width, compute the inserted native PPI from the embedded image pixel dimensions and the displayed size. If the source pixels would fall below the active native-PPI threshold at正文 width, the run must first look for a higher-resolution image in the locked material source. If no higher-resolution source exists, do not stretch the low-resolution raster; constrain the displayed extent to the largest size that satisfies the native-PPI threshold, record `native_resolution_constrained_width=true`, and bind the missing-high-resolution-source evidence.
- User-reported blurry figures must enable native-PPI enforcement in `scripts/audit_docx_figure_extents.py`. A figure inserted at low native PPI is failed even when its outer selection frame reaches正文 width, because the visible text remains unreadable after PDF conversion.
- Figure extent evidence must report text width, actual inserted width, width-to-text ratio, and whether any height constraint prevented paragraph-width placement.
- Figure extent evidence for blurry-image complaints must also report pixel width/height, inserted PPI, native PPI at正文 width, native-resolution constraint verdict, and native-PPI issue count.
- The acceptance keyword for low-pixel raster shrinkage is `native-resolution constrained width`; records may expose it as `native_resolution_constrained_width`, but the rule owner remains this paragraph.
- The image-holder paragraph itself must be part of the insertion-stage audit. Body figures must be centered in a dedicated holder paragraph or an equivalent direct-format holder baseline with zero effective first-line, left, right, hanging, and character-unit indents. A body first-line indent, body-text first-line indent, hanging indent, list residue, or non-centered alignment that shifts the picture body away from the paragraph margins is a hard failure even when the drawing extent is correct.
- The insertion-stage audit must cover all drawing-holder paragraphs, including front matter, body, appendix, headers/footers when present, and any non-caption holder paragraph that carries a real drawing.
- The acceptance keyword for unsafe holder indentation is `abnormal image-holder indent`; records may expose it through abnormal-indent counts or holder effective-indent fields.
- Width/height evidence alone is not enough. The same insertion-stage audit must also reject hidden visible-content clipping, including nonzero DrawingML `a:srcRect` crop rectangles, legacy/VML crop attributes when present, image-holder paragraphs whose exact line spacing is smaller than the inline image extent, and drawing extents that exceed the safe page-height threshold.
- A figure whose outer Word selection frame is paragraph-width but whose visible picture body is offset by inherited body indentation, cropped, clipped, truncated, shown as a thin strip, or split across pages is failed. Clear the holder layout/crop/paragraph-safety defect or resize/move the figure block, then rerender the touched page and adjacent page before handoff.
- For user-reported image display-incomplete or image indentation complaints, final evidence must include a visible-content completeness verdict, image-holder layout verdict, image-holder layout issue count, nonzero crop count, exact-line-spacing clipping count, safe page-height verdict, and rendered page or page-region review for each affected figure.
- Source preservation does not exempt holder layout safety. Every real drawing-holder paragraph in `word/document.xml`, including cover, front-matter, logo, banner, body, appendix, and other non-caption image paragraphs, must be audited for zero effective abnormal indentation, centered or donor-authorized effective alignment, non-clipping line spacing, and zero visible crop. Front-matter/template images are not resized by the paragraph-width rule unless explicitly authorized, but their holder paragraph must not inherit body first-line/hanging/list indentation or clipping line spacing.
- Final acceptance for user-reported small-image complaints must bind `scripts/audit_docx_figure_extents.py` evidence with a passing paragraph-margin width verdict.

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
- A draw.io structural source must contain native mxGraph shapes and connectors. Imported `shape=image` cells, `data:image` payloads, image URLs, pasted material screenshots, or generated PNG/SVG artwork inside a draw.io wrapper are raster impostors and must fail validation.
- When a user-provided material document is the source authority, bind the material file SHA256, inventory row, anchor/caption, extracted-preview SHA when available, and source-match verdict. This proves provenance only; it never waives the draw.io/SVG/raster fallback contract for structural figures.
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

Formal mechanical drawings and CAD thesis figures must have real engineering-drawing provenance. A schematic, concept figure, sketch, mockup, placeholder, sample image, AI-generated illustration, or raster-only "looks like CAD" asset may be used only as an internal draft, never as a substitute for any required assembly drawing, part drawing, CAD sheet render, CAD appendix sheet, or thesis figure that the task/template/teacher requires as a formal drawing. The mechanical CAD audit must record `formal CAD source provenance` and `schematic/concept substitute rejection` evidence; any package, figure manifest, appendix, caption, nearby prose, filename, or metadata that presents `示意图`, `概念图`, `草图`, `简图`, `样例图`, `schematic`, `concept`, `sketch`, `mock`, or `placeholder` as the final required drawing is a hard handoff blocker.

### CORE-FIGURE-016. Official CAD Command Route And External Case Sourcing Must Be Proven (Mandatory)

For formal mechanical CAD drawing mutation, and especially when the user says drawings must be changed with commands inside CAD, the default production path is a real CAD official command route before any drawing mutation. Acceptable evidence includes an installed AutoCAD/compatible CAD executable or COM automation route, ODA Drawings Explorer executing its own CAD `.scr` command script, QCAD/qcadcmd executing its own CAD script/command workflow, or another CAD application that actually opens the drawing, runs CAD commands, and saves/plots the output. The evidence must include a disposable command-route test before production mutation, the exact executable path or COM ProgID, product/version evidence, the command script path, the command log, and generated DWG/DXF/PDF hashes. ODA File Converter, ezdxf, SVG/PDF generators, image renderers, or library-only DXF writers may be used for conversion, inspection, or generating a CAD command script, but they are not a substitute for the required `official CAD command route`. If no official CAD command route is available, the CAD mutation lane is blocked and the handoff must not claim that drawings were redrawn, traced, plotted, dimensioned, or restyled in CAD. The run must record `CAD official command route verdict`, `CAD official command test log`, `CAD executable or COM ProgID evidence`, and `non-CAD fallback rejection verdict` before accepting any DWG/DXF/PDF drawing-mutation claim.

An official CAD command route proves only that the drawing mutation path used a CAD application. It does not clear downstream CAD hard gates by itself. QCAD/qcadcmd, AutoCAD, or another official command workflow must still be followed by current parseable evidence for source color-family, lineweight/linetype fidelity, `$LWDISPLAY`, CENTER/HIDDEN source linetype normalization, Layer `0` color, entity true-color overrides, frame overflow, content overlap, rendered readability, source-to-render derivation, and package JSON parseability before handoff.

External CAD case pages or download-site previews may be used only as reference-sourcing evidence for layout, readability, title-block/BOM placement, dimension/leader density, and view organization. They must not be copied verbatim as final project geometry, title-block data, proprietary part structure, or unlicensed DWG content unless the current run records a license or user-owned source authority. A "trace the website drawing" request must be implemented as an original project drawing that borrows only the reference's drawing grammar; the manifest must record `external CAD case reference URL`, `reference-use restriction`, `no verbatim geometry copying verdict`, and the project-specific parameter/design source used for the final sheets. When a user points to a download site such as jixie5 as a same-type drawing source, the run must also record an external-case annotation checklist covering view hierarchy, dimension placement, leader/balloon routing, title block, BOM/material table, technical notes, and protected blank/reserved zones before any CAD command mutation.

### CORE-FIGURE-017. CAD Open-View Structure Must Be Complete And Reference-Aligned (Mandatory)

When a user or teacher says the drawing is unreadable in CAD, looks like linework pasted together, shows only scattered parts, lacks a complete assembled object, or asks to follow teacher reference drawings / a good reference PDF such as `2222.pdf`, the acceptance surface is the CAD open view and structural composition, not only the exported PDF/PNG sheet. The final drawing must open in CAD as a coherent mechanical object with recognizable main views and required local/detail/section views. A package fails when the assembly sheet is just detached parts, unrelated local blocks, dense line bundles, or a collage that cannot be identified as the designed machine even if entity counts, colors, lineweights, and rendered-overlap audits pass.

For this surface, the run must lock the strongest available structural baseline before mutation: user-provided good PDF/image, teacher-provided DWG/DXF/PDF reference, or project reference drawings in the named reference directory. The final sheet may change parameters, markings, and non-proprietary details, but it must preserve the reference drawing grammar: complete main assembly structure, subordinate views located like engineering views rather than loose fragments, meaningful annotations, leader/balloon callouts, dimensions, title block, BOM/material table when present, and technical notes. Do not copy protected geometry verbatim when the reference is not owned by the user; redraw an original project sheet that follows the reference's level of detail and view organization.

For teacher-reference redraw work, the CAD production sequence is part of acceptance: first trace or redraw the complete object structure and view hierarchy from the locked reference baseline inside the official CAD command route, then add or repair dimensions, leaders/balloons, title block, BOM/material table, and technical notes. Do not start from loose part blocks, disconnected local details, or table/title-block restyling and later call the result an assembly drawing. The audit must record `structure-first redraw workflow verdict`, `dimension/leader/title-block second-pass verdict`, and `loose-part collage rejection verdict` when the user asks to follow a teacher template or reference drawing directory.

For A0 total assembly sheets of process or rotary equipment, structural maturity is a first-class gate. The open view must make the machine's design idea readable without relying on the title alone: process/material flow, drive chain, support/load path, shell or frame, end/head/flange interfaces, shaft/seal/bearing or equivalent rotating support, internal working elements, main section, local/detail views, BOM/title block, dimensions, and technical notes must be arranged as one coherent design. A large outline with decorative hatching, a simple cylinder with scattered details, or an entity-dense sheet that does not explain how the equipment works fails even when it has valid DWG headers, correct colors, and no obvious overlap. The audit must record `a0_process_flow_verdict`, `a0_drive_chain_verdict`, `a0_support_load_path_verdict`, `a0_internal_working_element_verdict`, and `a0_design_intent_readability_verdict` before accepting this class of A0 assembly drawing.

Acceptance must include `mechanical drawing reference baseline path`, `mechanical drawing CAD open-view close-up evidence path`, `mechanical drawing CAD open-view structural coherence verdict`, `mechanical drawing complete assembly/object recognizability verdict`, `mechanical drawing scattered-parts rejection verdict`, `mechanical drawing reference-view trace alignment verdict`, `mechanical drawing annotation/leader/title-block completeness verdict`, and `mechanical drawing external-case annotation checklist path`. The close-up evidence must be diagnostic-overlay-free CAD-open screenshots, CAD-exported crops, or equivalent review images of the final source drawing, including the main assembly area and any user-reported dense zones. Small full-sheet thumbnails, PDF-only overview renders, or screenshots covered by audit overlays cannot close this rule.

When the current run does not contain a stronger user-provided mechanical CAD package, the mandatory default mechanical CAD baseline is the stored SGB620/80T v12 six-sheet passed package baseline under `references/visual-style-samples/mechanical-cad/default-mechanical-cad-baseline-sgb62080t.md`, together with `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-00-a0-overall-assembly.png`, `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-01-a1-head-drive-assembly.png`, `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-02-a1-middle-trough-assembly.png`, `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-03-a1-tail-tension-assembly.png`, `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-04-a2-scraper-part.png`, `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-05-a2-sprocket-shaft-component.png`, `sgb62080t-v12-six-sheet-baseline/mechanical-drawing-package-v12-fixed.json`, and `sgb62080t-v12-six-sheet-baseline/rendered-review-v12.json`. This fallback must be used as both the visual baseline and the complexity baseline. A stronger current-run user-provided sample wins over the stored fallback baseline. The older v5 three-sheet sample files and the earlier v10 six-sheet files remain historical comparison evidence only and must not be used as the default for new SGB620/80T-like conveyor drawing packages.

For SGB620/80T-like conveyor A0 total assembly sheets, the A0 overall assembly layout baseline is a sheet-composition gate, not a density-only gate. The A0 sheet must visually follow the accepted user/sample total-assembly drawing family: two full-length conveyor views with upper and lower full-length views across the main sheet width, sectional/detail views kept in subordinate lower zones, a title block and BOM right zone at the right or lower-right protected area, balloon callouts routed from assemblies to the BOM, and a recorded balloon-to-BOM row match. A single long conveyor strip with unrelated local blocks, a dense entity field, or a title-block-only restyle cannot satisfy the A0 total assembly sheet even if entity counts, A0-equivalent workload, and outside-frame audits pass. Acceptance evidence must expose `a0_reference_layout_verdict`, `two_full_length_views_verdict`, `bom_right_zone_verdict`, and `balloon_bom_row_match_verdict` or equivalent fields before claiming the A0 total assembly matches the reference layout.

For a three-sheet-A0-equivalent mechanical drawing requirement, the output must show academic/manufacturing drawing depth, not only occupy A0 page area. Acceptance evidence must include a package manifest, rendered-sheet previews, a rendered no-overlap review for drawing frames, title blocks, BOM/material tables, technical-requirement tables, notes, dimension text, and view geometry, plus explicit rendered verdicts for boundary clearance, detail density, title-block/table/notes isolation, annotation-margin clearance, and local crowding; readable text/font evidence, sheet-size/workload evidence, and a source-format capability verdict are still required. A visual-review string alone is not enough: the package must include a machine overlap audit that reports zero text-entity overlaps, zero reserved-zone collisions, zero title-block/table/note collisions, zero annotation collisions, zero frame-clearance violations, sheet-level minimum clearance, and local ink-density metrics for every rendered sheet. It must also include an outside-frame ink audit on the rendered PNG/PDF sheets that detects the outer drawing frame independently from view geometry and reports zero independent ink components outside that frame. It must also include a hatch/section fill clipping audit that proves every section fill is clipped to its owning entity boundary and reports zero adjacent view crossing, zero dimension line crossing, zero title block/table/BOM/frame crossing, and zero blank background leak. It must also include a text legibility audit with CAD text height, preview DPI, minimum rendered text pixel height, and per-sheet pass rows; current mechanical package acceptance requires at least 3.8 mm CAD text height and 40 px rendered text height in the final preview audit unless a stronger school/sample rule applies. It must also include a manufacturing complexity audit with at least eight real manufacturing detail families per sheet so the drawing is not a sketch-level outline or dense-filler-only pass.

Mechanical CAD text integrity and text orientation are hard gates. CAD/DXF/PDF/rendered text with mojibake, replacement characters, tofu/missing-glyph boxes, square placeholder glyphs, unreadable missing-font boxes, empty required labels, upside-down text, mirrored/backward text, or 180-degree title/table/dimension/annotation text blocks the drawing package. Acceptance must bind machine-readable `text_integrity_audit`, `text_orientation_audit`, and `cad_text_quality_verdict` evidence with zero mojibake/tofu/missing glyph count, zero missing required drawing text count, zero upside-down text count, and zero mirrored text count. Standard 90-degree dimension conventions may be accepted only when the audit distinguishes them from 180-degree/upside-down or mirrored text; a manual screenshot statement cannot clear this gate.

DXF text-orientation validators must parse entity semantics, not only raw group-code numbers. For `TEXT`, `ATTRIB`, and `ATTDEF`, group code `71` may carry text generation flags and must still reject mirrored or upside-down flags. For `MTEXT`, group code `71` is an attachment point, so it must not be counted as a generation flag; this rule is keyed as `MTEXT attachment point` and `DXF MTEXT group 71` in the owner map. `MTEXT` orientation must instead be checked from rotation and direction-vector groups such as `11` and `21`. A valid centered `MTEXT` note or dimension label with attachment point `71=5` is not mirrored or upside down unless its rotation/direction evidence proves that defect.
- When DWG/PDF/DXF drawings are handed off, run `scripts/audit_mechanical_drawing_package.py` on the exact final package or drawing folder. If the user provided CAD/PDF samples, pass them as `--reference` inputs and bind the report path, package path, package SHA256, verdict, and any density/encoding/CAD-structure blockers before claiming the drawing package matches the sample.
- For sample-based mechanical CAD work, use distributed density gates rather than only total file counts: per-DXF entity minimums, dimensions spread across multiple sheets, per-PDF drawing-object density by sheet size, required engineering tokens such as sectional/detail/technical-requirement/BOM/weld/roughness terms, and a DWG byte-density ratio against the user-provided CAD source package. When a reference CAD/DWG package is present, `audit_mechanical_drawing_package.py` must enforce an effective DWG byte-density ratio floor of at least `0.35` even if a caller passes a lower threshold. A single dense strip, one overloaded sheet, or a normal Chinese character falsely detected as mojibake cannot clear or block acceptance by itself; the machine gate must expose those details for the audit lane to review.
- For this error class, the audit report must use the v4 mechanical package schema or a stricter successor and expose `density_verdict`, `manufacturing_depth`, `rendered_review_verdict`, `rendered_frame_overflow_verdict`, a passing `machine_overlap_audit`, a passing `outside_frame_ink_audit`, a passing `hatch_clip_audit`, a passing `text_legibility_audit`, passing `text_integrity_audit`, `text_orientation_audit`, and `cad_text_quality_verdict` evidence, and a passing `manufacturing_complexity_audit`; a pass that lacks dimension totals, geometry totals, DWG reference ratio, block/arc/hatch/text depth, rendered-review fields, machine collision metrics, outside-frame independent pixel-component metrics, hatch/section fill clipping evidence, text entity overlap count evidence fixed at zero, rendered text-height metrics, zero mojibake/tofu/missing-glyph count evidence, zero upside-down/mirrored text evidence, real manufacturing detail-family evidence, per-sheet density failure detail, or any clearance/isolation verdict that still leaves title blocks, tables, notes, annotations, or figure text unreadable is not a final acceptance record.
- For this error class, entity counts, dimension counts, byte ratios, file counts, and PDF parse counts are necessary but never sufficient. The package manifest must include `rendered_review` evidence with real preview paths, per-sheet rows, `no_overlap_verdict=pass`, `boundary_clearance_verdict=pass`, `detail_density_verdict=pass`, `title_block_table_notes_isolation_verdict=pass`, `annotation_margin_clearance_verdict=pass`, `local_crowding_verdict=pass`, `text_legibility_verdict=pass`, `sheet_layout_verdict=pass`, `manufacturing_view_depth_verdict=pass`, `outside_frame_ink_verdict=pass`, `hatch_section_fill_clipping_verdict=pass`, `content_overlap_verdict=pass`, `entity_count_only_verdict` set to a rejection value such as `not-used`, `machine_overlap_audit.passed=true` with zero collision counts and zero `text_entity_overlap_count`, zero `table_text_grid_collision_count`, zero `view_detail_overlap_count`, zero `dimension_table_crossing_count`, zero `leader_view_crossing_count`, zero `balloon_geometry_collision_count`, zero `bbox_helper_undercoverage_count`, `outside_frame_ink_audit.passed=true` with zero outside-frame independent ink/text/leader/hatch/table/title-block components in every rendered PNG/PDF sheet, `hatch_clip_audit.passed=true` with zero entity-boundary escape, adjacent-view crossing, dimension-line crossing, title-block/table/BOM/frame crossing, and blank-background leak counts, `text_legibility_audit.passed=true` with minimum rendered text height at or above the current threshold, and `manufacturing_complexity_audit.passed=true` with multiple manufacturing detail families per sheet. A package that passes only from entities, dimensions, or pass-shaped visual-review strings is an entity-count-only false-pass or self-reported-rendered-review false-pass and must be blocked.
- For this error class, the rendered review must also block content overlap false-passes: view boxes, detail boxes, tables, notes, title blocks, leader zones, dimension bands, and balloon callouts must be checked as separate content envelopes, and any overlap between unrelated envelopes is a failure even if the outer drawing frame still fits. The exact literal keywords `content-overlap audit`, `view-view overlap`, `detail-frame-main-view overlap`, `table-text/grid collision`, `dimension-line cross-view collision`, `dimension-line table-zone intrusion`, `leader-line cross-view collision`, `leader-line view crossing`, `balloon/detail collision`, `balloon geometry collision`, `view-box overlap`, and `bbox-helper-envelope audit` must remain present in the routed mechanical CAD rule text and validator-facing documentation.
- Protected mechanical drawing table zones are hard exclusion zones, not spare blank space. The title block, BOM/material table, technical-requirement table, revision/signature table, and their readable text cells must be protected from all unrelated drawing geometry. Any main view, local/detail view, section view, hatch block, dimension line, extension line, leader line, balloon circle, callout text, shaft/sprocket/gearbox geometry, or loose inspection symbol that touches or crosses a protected table/BOM/title-block envelope is a hard failure, even when the outer drawing frame still fits and the table grid itself remains visible. A repair that adds a new local/detail view over the BOM or pushes view geometry into the right-side table area is a false pass. Acceptance evidence must expose `protected-table-zone intrusion audit`, `view-geometry table-zone intrusion count=0`, `detail-view table-zone intrusion count=0`, `leader/balloon table-zone intrusion count=0`, `dimension table-zone intrusion count=0`, and `title-block/BOM protected-zone intrusion count=0` or equivalent machine-verifiable fields bound to the exact final DXF/PDF/PNG package.
- If drawings are regenerated after a rendered-review record was created, all no-overlap evidence must be regenerated and rebound to the current DXF/PDF/PNG file set. A rendered review, contact sheet, or CAD audit whose creation time or bound package SHA predates the latest drawing output is stale and cannot be used to close a user-reported overlap defect.
- If the user reports repeated overlap or mismatched background hatching, the refreshed evidence must include a preview freshness check comparing every sheet render, review JPEG, contact sheet, rendered-review JSON, and package SHA against the final DXF/PDF/PNG output. Any stale rendered preview freshness mismatch blocks handoff even when older reports say pass.
- When a body figure or CAD preview is inserted into the thesis, leave a deliberate safety buffer from the text width instead of stretching to the absolute margin. For dense mechanical previews, a target width around `0.88` to `0.90` of the available text width is the default safe zone unless the current sample proves a different ratio is needed; an image that technically fits but visually kisses the frame or leaves no breathing room is still a layout defect and must be shrunk and rerendered.
- Do not use diagonal/cross/spiral/ground-line hatching as non-corresponding background fill, page shading, or complexity filler. Any hatch-like or section-fill-like linework must have a visible owning entity boundary, must visually correspond to a cut/material surface, and must be removed when it cannot be distinguished from background texture at rendered preview scale.
- Title-block, BOM/material table, technical-requirement table, tolerance-frame, and other table-owned text must be cell-bounded. Labels and values must be centered, wrapped, or scaled inside their intended cell rectangle; any text touching/crossing a cell border, floating outside the table grid, or positioned by an unrelated free coordinate as if it were table content is a handoff blocker. A renderer-side clip that hides overflow is not sufficient unless the source text placement also records the owning cell and remains readable at final preview scale.
- Title-block/BOM/table grid completeness is a source-and-render hard gate when a user sample, source PDF, source DXF, or accepted prior sheet is available. The audit must compare protected table/title-block short horizontal and vertical line segments, cell intersections, right-side author/page cells, signature/date rows, and BOM row dividers against the locked baseline; missing short table line count, broken cell-border count, and table-grid topology mismatch count must all be zero. A long-line or entity-count-only pass cannot clear a user-reported missing-line defect. The exact literal keywords `title-block short-line topology audit`, `missing short table line count`, `broken cell-border count`, and `table-grid topology mismatch count` must remain present in this routed mechanical CAD rule text and validator-facing documentation.
- Table/title-block text and protected grid/frame lines must also be checked at source level. The linework fidelity audit must expose `table_text_grid_collision_count` and fail when table-owned text, title-block values, author/student information, page-number text, BOM text, signature/date labels, or normal annotation text intersects or is covered by protected table/grid/frame linework. Rendered review crops must also be diagnostic-overlay-free so audit rectangles or heatmaps do not hide whether the final drawing itself has covered text. The exact literal keywords `table_text_grid_collision_count`, `title-block text-grid collision`, `covered text by grid line`, and `diagnostic-overlay-free title-block crop` must remain present in this routed mechanical CAD rule text and validator-facing documentation.
- Ordinary note text must also have an explicit drawing owner. Text may be owned by a dimension, leader, datum/surface symbol, sheet title, or a visible callout box; otherwise it is unowned free text. The exact literal keyword `annotation ownership audit` must remain present in the routed mechanical CAD rule text and validator-facing documentation, and the audit must report zero `unowned free text` and zero `unsupported floating text` before handoff.
- Title block, BOM/material table, technical-requirement table, tolerance-frame, and drawing-frame table zones are protected zones. Non-table view geometry, local-detail/detail-frame geometry, dimension lines, dimension text, leader lines, hatch strokes, center lines, or free annotations must not touch or enter those protected zones and must keep the recorded minimum clearance, defaulting to at least 3.5 mm for A2/A0 rendered reviews unless a stronger school/sample rule applies. Local detail panels, enlarged sections, auxiliary views, and section/detail frames are view geometry for this rule; they may not be added as floating complexity fillers in the bottom band when they collide with or visually press against the title block, BOM, technical notes, or border. The exact literal keywords `reserved-zone intrusion audit`, `dimension-line table-zone intrusion`, and `view-geometry reserved-zone intrusion` must remain present in the routed mechanical CAD rule text and validator-facing documentation so protected table-zone intrusion cannot pass from text-cell evidence alone.
- The exact literal keyword `annotation margin clearance` must remain present in the routed mechanical CAD rule text and validator-facing documentation so `CORE-FIGURE-010` owner-map keyword checks can bind this rendered-review verdict without ambiguity.
- Mechanical CAD rendered PNG/PDF sheets have zero tolerance for border overflow. The drawing-frame outer rectangle must be detected as its own frame, and every independent pixel/ink component outside that frame must be classified. Outside the drawing frame there must be no view geometry, dimensions, text, leader lines, hatch/section fill, tables, title-block content, BOM content, notes, revision/signature content, page ornaments, or stray pixels from clipping/export. The exact literal keywords `outside-frame ink audit`, `outside-frame independent ink component`, `outside-frame text component`, `outside-frame leader component`, `outside-frame hatch component`, and `outside-frame table/title-block component` must remain present in the routed mechanical CAD rule text and validator-facing documentation; a stale `rendered_review` pass that lacks this audit cannot close a current border overflow defect.
- Mechanical CAD rendered sheets must also respect the inner-frame safe margin. The engineering content itself, including view geometry, leaders, leader text, dimension lines/text, detail blocks, notes, tables, and inspection bands, must remain inside the detected inner safe strip for the sheet family. On the standard SGB620/80T A0 baseline, the right-side safe boundary is 1145 mm unless a stricter current-run sample explicitly records a narrower safe strip. The exact literal keywords `inner-frame safe-margin audit`, `inner-frame safe-boundary intrusion`, `leader-text inner-frame intrusion`, `view-geometry inner-frame intrusion`, and `dimension-text inner-frame intrusion` must remain present in the routed mechanical CAD rule text and validator-facing documentation; an outer-frame pass alone cannot close an inner-frame overflow defect.
- Mechanical CAD exports must preserve lineweight and linetype fidelity across DXF, DWG, PNG, and PDF. If the final PDF visually changes CAD strokes, makes hidden/center lines look continuous, thickens fine structure lines, or alters dash spacing compared with the CAD source, the drawing package is failing even when all entities remain inside the frame. DXF/DWG generators must either set stable layer/entity lineweights and linetype scales or record a tool-backed reason why the active CAD converter controls them; for generated DXF this normally means preserving `$LTSCALE`, `$CELTSCALE`, `$PSLTSCALE`, `$LWDISPLAY`, entity or layer lineweight through DXF group `370`, and linetype scale through DXF group `48`. The source line family coverage must prove that coarse/fine mechanical line expression is not cosmetic: thick solid visible outlines, thin solid detail/dimension/table/text lines, center dash-dot lines, hidden dashed lines, section hatch lines, dimension lines, leader/annotation lines, and title/BOM/table grid lines must be distinct by CAD layer, linetype, color, and lineweight before PDF/DWG export. For generated or command-redrawn DXF, CENTER and HIDDEN layers must carry real source linetype names, not only hand-segmented short line fragments on a Continuous layer; center/hidden entities must be BYLAYER or explicitly use the matching CENTER/HIDDEN linetype. A source file with `$LWDISPLAY=0`, missing CENTER/HIDDEN linetypes, or Continuous center/hidden layers cannot pass from a visually dashed PDF. The lineweight/linetype fidelity audit must also reject low-contrast/light CAD colors at source level by checking layer/entity color and true-color values for white-background plotting, before relying on PNG/PDF rendered ink contrast. The lineweight/linetype fidelity audit must bind the exact audited CAD package with `package_sha256`; a pass-shaped linework report without a package hash cannot close a final CAD handoff. The exact literal keywords `lineweight/linetype fidelity audit`, `source line family coverage`, `CENTER/HIDDEN source linetype normalization`, `source-level low-contrast CAD color rejection`, `thick solid`, `thin solid`, `center dash-dot`, `hidden dashed`, `section hatch`, `package_sha256`, `DXF group 370`, `DXF group 48`, `$LWDISPLAY`, `$LTSCALE`, and `$PSLTSCALE` must remain present in this rule so PDF conversion cannot close a user-reported linework defect from a geometry-only audit.
- CAD layers routed into thick-solid, frame, or title-block families must keep thick-family lineweight. A layer literally named `TITLE` is a title-block thick-family layer unless a stronger user sample explicitly separates it, so `TITLE thick-family lineweight` must meet the same minimum as `OBJECT_THICK`/`FRAME`; color-only restyling must not preserve or introduce an already-too-small title-block lineweight.
- Mechanical CAD PDF renders must preserve true engineering sheet page boxes. The final PDF page box, not only the visible image aspect ratio or filename, must be recognized as the intended A-series sheet size such as A0, A1, or A2, and the package-level `estimated PDF sheet workload` must meet the required A0-equivalent workload. If a renderer, plotting backend, or export pipeline resets the canvas to a compact preview size, the run must repair the exporter and rerun the package audit; do not claim completion from a visually similar small PDF. A raster-only PDF replacement is not an acceptable workaround when vector drawing-object density is part of the package audit. The exact literal keywords `PDF page-box sheet-size audit`, `estimated PDF sheet workload`, `A-series page box`, and `compact preview size` must remain present in this rule so page-size regressions are blocked together with linework, overlap, and overflow defects.
- Mechanical CAD rendered PNG/PDF previews must have readable stroke contrast. Yellow, cyan, gray, or near-white line colors that are technically present but disappear on a white plotting background are a hard failure. The package audit must expose a `rendered ink contrast audit`, `readable ink ratio`, `worst readable ink ratio`, and `minimum readable ink ratio`; a final package with `rendered_ink_contrast_verdict.passed=false` cannot pass even when entity counts, DWG/DXF counts, and frame-overflow checks pass.
- `rendered_review_verdict.passed=false` is a package-level hard failure. A final mechanical drawing audit cannot report overall `passed=true` when the rendered review is missing, stale, disabled by a low caller flag, or failing. For formal mechanical CAD packages, rendered review is effectively required whenever strict CAD structure or any reference CAD sample is in scope.
- Mechanical CAD rule fixes must be wired into the execution flow, not only recorded as prose. When a user-reported CAD defect creates or changes a hard rule, the same run must update the active checklist/task card, modify the actual source DXF/DWG or canonical generator, rerun official CAD conversion where required, regenerate PDF/PNG previews, create or refresh rendered-review JSON with all validator-required sub-audits, and rerun `scripts/audit_mechanical_drawing_package.py` on the exact final package. A rule-only change without current generated outputs and current package-audit evidence is incomplete. The exact literal keywords `CAD rule-to-flow binding`, `active checklist enforcement`, `regenerate DXF/DWG/PDF/PNG after rule change`, `rendered-review JSON refresh`, and `final package audit rerun` must remain present in this rule so future runs cannot pass by documenting a restriction without executing it.
- When the current user supplies a stronger CAD/DWG/PDF sample and asks to use the current accepted drawing as the default standard, lock that sample or accepted current CAD package as the run baseline before mutation. A later package may not pass by reusing an older simplified SGB620/80T baseline, by renaming old sheets, or by changing only the title block. For formal graduation mechanical design tasks requiring a defined sheet set, the manifest must list every required sheet by name and size, such as A0 total assembly, A1 head drive, A1 middle trough, A1 tail tensioning, A2 scraper part, and A2 sprocket shaft assembly; missing, merged, or schematic substitutes are hard failures.
- For SGB620/80T-like six-sheet conveyor packages, the default formal sheet set is A0 total assembly, A1 head drive assembly, A1 middle trough assembly, A1 tail tensioning assembly, A2 scraper part, and A2 sprocket shaft component. If a user writes `机尾装紧部`, interpret it as the mechanical `机尾张紧部` sheet unless the user later gives a different official title. A locked six-sheet manifest must not be failed only because a generic dimension-distribution setting expected eight dimensioned DXF files; `audit_mechanical_drawing_package.py` must clamp `min_dimensioned_dxf_files` to the available required formal DXF sheet count and expose `effective_min_dimensioned_dxf_files=6` or the equivalent evidence.
- When the user locks a formal sheet set whose A0-equivalent workload is exactly the requested workload, pass that exact value to `audit_mechanical_drawing_package.py --min-total-a0` and record it in the manifest. For the SGB620/80T six-sheet set A0 + 3*A1 + 2*A2, the required user-locked A0-equivalent workload is `3.0`; do not let the generic default threshold reject this exact six-sheet package.
- Border overflow and content overlap are coupled hard gates. A repair that scales the drawing down to stop outer-frame overflow but creates table/text/view/dimension/leader overlap is still failing; a repair that removes overlap but pushes any independent ink outside the outer frame or inner safe margin is also failing. Moving or resizing local-detail panels to clear a title-block/table intrusion also invalidates old rendered-review evidence, because the same edit can create new crowding or border pressure elsewhere. Acceptance must include both current rendered overflow evidence and current content-overlap evidence bound to the same final DXF/PDF/PNG package SHA. These gates must be rerun after every lineweight, scale, color-family, linetype, title-block, local-detail panel, view-layout, or conversion change; an old no-overlap report from before a source linework/color/layout repair is stale by definition.
- Mechanical CAD rendered readability is a hard gate when the user reports linework that visually merges into one mass, indistinguishable overlapping lines, text overlap, or text covered by drawing geometry. The repair must change the final delivered sheet set itself by splitting, rescaling, decongesting, moving labels, or removing stray cover geometry; adding a separate zoom booklet, high-DPI preview, or visual note does not close the defect when the delivered standard sheets still fail. Do not treat layer-name-only cleanup as sufficient for table, BOM, note, dimension, or title-block readability: if rendered `text_graphic_cover` or `text_text_overlap` evidence remains, the final CAD source must move/redraw the text, redraw the table with safe cell padding, or remove the specific covering stroke/stray grid from the delivered sheet itself, then rerun the exact standard-sheet audit. DXF/DWG source lineweight compliance alone does not close a user-reported visual merging defect: if the final PDF/PNG has excessive rendered effective lineweight, failed line-cluster distinguishability, or a failing standard-sheet render density gate, the plotting/export scale, view scale, local decongestion, or delivered sheet geometry must be repaired and all final renders regenerated. If the user reports that lines remain visually merged after a default rendered-readability PASS, the run must escalate to a `small-tile line-cluster challenge` using stricter rendered tiles and crop-based visual review; a broad-tile density pass cannot close the defect while smaller tiles or human-reviewed crops still show indistinguishable line bundles or text/line cover. Run `scripts/audit_mechanical_render_readability.py` on the exact final PDF and PNG sheet renders and bind `mechanical render readability audit`, `text_text_overlap_count=0`, `text_graphic_cover_count=0`, `severe_line_crowding_count=0`, `max_observed_tile_ink_density`, and reviewed PDF/PNG SHA256 evidence before handoff. A pass claim is invalid if this audit is missing, stale, run on earlier v67/v66 renders, or run only on optional local enlargement pages instead of the final delivered sheet set.
- Source CAD lineweight and final rendered stroke width are separate acceptance surfaces. When the user requires source lineweights such as green thick/frame lines at `0.5 mm` and other colored lines at `0.25 mm`, the DXF/DWG source must keep those values, but PDF/PNG plotting may use a documented `render-only stroke cap` or `render-only lineweight scaling` so adjacent mechanical lines, dimension text, table text, hidden lines, and center lines remain readable in the delivered standard-sheet render. A source-lineweight pass cannot force the preview to stay visually thick, and a render-only thinning pass cannot alter or downgrade the audited DXF/DWG source lineweight policy. The acceptance record must state the source lineweight verdict, render-only stroke cap/scaling, and whether the delivered standard sheets themselves, not only optional zoom crops, pass the strict readability audit.
- A `small-tile line-cluster challenge` is a pressure test for linework, not a replacement for semantic review. If the raw small-tile density count flags isolated title/BOM text glyphs, dimension text, green/cyan marker dots, hole callout circles, or other annotation symbols while `text_text_overlap_count=0`, `text_graphic_cover_count=0`, table/text proximity passes, and diagnostic-overlay-free crops show readable content rather than merged mechanical line bundles, record those hits as `semantic_density_false_positive` instead of a linework blocker. The run must still repair any crop that shows indistinguishable adjacent mechanical lines, text covered by strokes, table text touching grid lines, or filled symbols that visually obscure required labels. Acceptance in this case requires a `semantic small-tile density classification` report, no-overlay crop evidence, and an explicit blocker count for real line bundles. The acceptance record must expose the exact small-tile threshold, failed sheet list, `requires_crop_review` count, `line_bundle_blocker_count`, diagnostic-overlay-free crop review path, and final PDF/PNG SHA evidence so a raw-density false positive cannot hide a current user-reported line-cluster defect.
- When a standard sheet looks unreadable because the engineering views occupy too little of the plotted frame, the repair must record a `standard-sheet view occupancy` verdict and fix the final sheet by increasing the usable view scale, reducing empty legacy table/frame space, or lowering the rendered effective stroke width. Source layer lineweight compliance and optional zoom pages cannot close a line-cluster complaint when the actual delivered standard sheet remains too small or fuzzy to distinguish adjacent lines.
- Crop-based evidence for line-cluster or table/text review must include a `diagnostic-overlay-free crop review` path or an explicit note that any red boxes, heatmap outlines, or similar markers are audit overlays only and are not present in the final PDF/PNG/DXF. A human or agent review cannot reject or accept the final drawing solely from diagnostic rectangles that cover the crop; it must compare the unmarked final sheet crop or final PDF/PNG render.
- Mechanical CAD graduation-design sheets must not leave drawing text as loose decoration. Every visible text entity must be owned by exactly one drawing surface before handoff: a dimension object or dimension line, leader/callout/balloon, datum/roughness/tolerance symbol, view/detail/section label, title block/BOM/material table cell, technical-requirement note block, or a documented local feature list. Standalone dimension-like fragments, isolated note fragments, or scattered labels that do not visibly attach to geometry are blockers even when they do not overlap anything. Examples include a lone `m=2`, a loose `关键尺寸` list outside a note/table, or labels such as `安装高`, `定位高`, and `截面高` placed in empty space without leader/dimension ownership. The audit must expose `unbound scattered text`, `dimension-like text without anchor`, `orphan text examples`, `owner-zone coverage`, `user-reported crop binding review`, and `annotation ownership audit` evidence with blocker counts fixed at zero.
- If a user reports text overlap, text covered by lines, or text floating away from the drawing content after a prior PASS, all previous rendered-readability and annotation-ownership evidence is stale. The next run must add explicit checklist rows for the reported crop zones, regenerate the standard DWG/DXF/PDF/PNG sheets, and include diagnostic-overlay-free crops for those zones. Moving text into a new note block is acceptable only when the note block has a visible border/anchor, enough padding, and no intrusion into the drawing frame, BOM, title block, dimensions, leaders, or main view.
- User-reported CAD screenshots or crops are hard review zones, not examples that can be overridden by a full-sheet PASS. For every reported zone, the final record must bind a `user-reported text-cover crop audit path`, `diagnostic-overlay-free after-crop paths`, and `reported crop blocker count=0`; the reviewed crop must come from the exact final standard DWG/DXF/PDF/PNG sheet family after the last drawing mutation. A broad rendered-readability PASS, full-sheet thumbnail, or line-family/source audit cannot close the defect while any reported crop still shows covered, crossed, garbled, or unowned text.
- Every visible CAD text bounding box requires a clear no-ink padding zone before handoff. Section hatch, diagonal hatching, ground/support lines, flow arrows, center lines, hidden lines, leader/dimension lines, table/grid strokes, load-path strokes, and stray construction lines must not cross the glyph bbox or its recorded text exclusion halo. Acceptance must expose `text exclusion halo audit`, `text exclusion halo violation count=0`, `diagonal hatch/section/flow text-cover count=0`, and `reported crop blocker count=0`. This applies to normal labels and to dimension-like text such as diameter, height, inlet/outlet, centerline, flow-direction, support-load, BOM/balloon, and title-block fields; a dimension or leader may point to a feature but may not visually cut through the readable text.
- Diagonal hatch, section fill, process-flow arrows, support/load-path strokes, and other engineering guide lines are blocker geometry when they cross or sit behind text in a way that makes the label hard to read. The source drawing must clip, break, move, or redraw those strokes around the text halo, or move the label to a bound leader/note/cell zone with padding. Do not hide this with opaque raster masks, diagnostic overlays, optional zoom pages, or draw-order tricks that leave the final editable CAD source visually ambiguous.
- CAD text must stay as normal CAD text, not decorative or broken geometry. The source must use valid `TEXT`, `MTEXT`, `ATTRIB`, `ATTDEF`, or dimension text with an approved readable font/style. For mechanical CAD graduation drawings, the target Chinese CAD text style should be `HANYI_CHANGFANGSONG` / 汉仪长仿宋体 when available; if that exact font is not installed or licensed in the local environment, the run must record the fallback font file/name, bundle or point to the readable fallback when allowed, and prove that the fallback does not create mojibake, tofu boxes, missing glyphs, or unreadable Chinese text. Exploded text outlines, artistic text, vectorized glyph strokes, unsupported SHX/font placeholders, missing-font fallback boxes, or style records that make Chinese text unreadable fail the CAD text style/font gate even when PDF OCR appears readable. Acceptance must include a `CAD text style/font audit`, `normal CAD text entity verdict`, `target CAD Chinese font style`, `actual CAD font file`, `font fallback recorded verdict`, `artistic/vectorized text count=0`, and `unsupported CAD font/style count=0`.
- CAD text integrity must be proven on the final DWG itself, not only on the editable source DXF or a PDF/PNG preview. When Chinese text, title-block text, BOM text, or technical requirements have ever shown `???`, mojibake, tofu boxes, missing glyph boxes, or other unreadable fallback in a user screenshot or CAD open-view review, the final DWG must be re-converted or backread to DXF through a recorded CAD/ODA/QCAD route and audited again. Acceptance must expose `DWG backread text integrity audit`, `DWG backread question mark text count=0`, `DWG backread mojibake/tofu/missing glyph count=0`, `missing required drawing text count=0`, `bundled CAD Chinese fallback font`, and diagnostic-overlay-free title-block crops from the final DWG backread render. A source-DXF pass, PDF OCR pass, or white-background preview pass cannot close a current CAD mojibake defect without this final DWG backread evidence.
- User-reported stray or meaningless line artifacts are treated as delivered drawing geometry until proved otherwise. For lower-left frame-area artifacts, construction remnants, isolated `FRAME`/`HIDDEN` short segments, orphan helper crosses, and unexplained leader-like strokes, the repair must delete or bind the entity in the final CAD source and then rerun both source DXF and final DWG backread audits. Acceptance must expose `lower-left orphan line artifact audit`, `source lower-left orphan line artifact count=0`, `DWG backread lower-left orphan line artifact count=0`, and diagnostic-overlay-free lower-left crops from the final DWG backread render. A crop or PDF export that merely hides the artifact does not pass.
- Mechanical CAD table/title-block/BOM cells require measurable padding, not only closed grid lines. Text must stay inside its owning cell or note block and clear every border by the recorded minimum padding. The default minimum cell padding is `1.0 mm` unless a stronger school/template/CAD sample requires more. Acceptance must expose `min cell padding mm`, `cell padding violation count=0`, and `title-block cell containment`.
- Source lineweight policy is exact for CAD graduation-design handoffs when the user requires thick `0.5 mm` and thin `0.25 mm`: thick/frame/title/object-thick families must be DXF group 370 `50`, and thin/dimension/leader/text/table/center/hidden/hatch families must be DXF group 370 `25` unless a locked sample explicitly overrides the value. A range-only lineweight pass is not enough for this user-reported defect class; the lineweight audit and acceptance record must expose `source thick lineweight required`, `source thick lineweight observed`, `source thick lineweight mismatch count=0`, `source thin lineweight required`, `source thin lineweight observed`, and `source thin lineweight mismatch count=0`.
- Final CAD handoff must bind every regenerated artifact family, not only the combined ZIP or DWG ZIP. The final acceptance record must include exact current paths and SHA256 values for the final delivery ZIP, audited CAD package ZIP, DWG package ZIP, exact DXF package path, combined PDF, exact PNG render package path, and a drawing regeneration manifest. It must also expose a current package SHA binding verdict before handoff. A package with a stale rendered-review PASS, missing DXF/PNG SHA binding, or a regeneration manifest produced before the last CAD/DXF/DWG/PDF/PNG write is blocked.
- The exact literal keywords `hatch/section fill clipping`, `entity boundary`, `adjacent view`, `dimension line`, and `blank background leak` must remain present in the routed mechanical CAD rule text and validator-facing documentation so `CORE-FIGURE-010` owner-map keyword checks can bind this hatch/background verdict without ambiguity.
- The exact literal keyword `title-block cell containment` must remain present in the routed mechanical CAD rule text and validator-facing documentation so `CORE-FIGURE-010` owner-map keyword checks can bind table-cell text containment without ambiguity.
- If rendered previews show sketch-like simplified outlines, missing manufacturing views, missing section/detail/hatching/local-detail depth, overlapped text, overlapped dimension labels, crowded notes, annotation text that sits hard against the drawing frame, tables colliding with the title block or drawing frame, clipped borders, unreadable annotation text, or any frame/title-block/table overlap, the mechanical package fails even when the static CAD entity-density gate passes.
- When only one or two sheets fail the distributed detail-density gate, prefer adding real mechanical detail inside the existing main view or section envelopes of the failing sheet before adding new floating panels, new detached tables, or extra note boxes. Safe first choices are repeated hidden/center lines, guide-slot lines, bolt-hole rows, hatch/section refinements, and reducer/chain/seat internal structure lines that stay inside an already reserved view box. Avoid clearing the gate by stuffing new blocks into page gaps if that move increases local crowding risk near title blocks, BOM tables, technical-requirement notes, or dimension bands.

### CORE-FIGURE-018. Mechanical CAD Annotations Must Bind To The Correct Target (Mandatory)

- Mechanical CAD annotation acceptance is semantic, not only geometric. A drawing can pass border, density, text, and overlap checks while still failing because a leader, balloon, datum mark, roughness symbol, tolerance frame, or dimension label has drifted away from the part it is supposed to identify.
- Annotation ownership also covers ordinary CAD `TEXT`/`MTEXT`, not only formal leader entities. A text object is accepted only when it is inside a protected table/title-block/note cell, inside a view/detail label zone, or bound by a leader, dimension anchor, balloon, datum, section marker, or documented feature list. Unowned free text, unsupported floating text, unbound scattered text, and dimension-like text without anchor must be counted and fixed to zero before handoff.
- A text label that merely names a nearby part is not bound unless the final render shows a visible leader/arrow/dimension relation, a table row/cell relation, or an owner-zone relation recorded in the annotation ownership audit. Text that overlaps a balloon circle, crosses a green/colored arc, sits on a dimension line, or is visually covered by geometry fails both overlap and annotation-binding gates.
- Every formal callout, leader, balloon, boxed tolerance frame, datum tag, surface-roughness mark, and dimension must have an explicit owner before handoff:
  - source entity id or handle
  - leader start/end or dimension definition points
  - target view or local-detail frame
  - nearest target geometry or feature family
  - intended part/feature name or BOM row when applicable
  - distance from the annotation endpoint to the target geometry
  - uniqueness verdict when multiple possible targets are nearby
- Balloon-to-BOM binding is a hard gate. Balloon numbers must map to a valid BOM/detail row or a documented local feature list. Missing numbers, duplicate numbers, orphan balloons, balloons pointing to the wrong part family, or BOM rows without matching drawing targets block the CAD handoff.
- Dimension anchor validity is a hard gate. Dimension definition points must anchor to real geometry such as an edge, hole, shaft centerline, center mark, or construction centerline that belongs to the same view. A dimension anchored in blank space, title block/table zones, adjacent views, decoration lines, or an unrelated detail frame blocks handoff.
- Leader and annotation drift must be audited after every render, scale, postprocess, CAD-command conversion, or layout compaction step. A source-level pass is stale if the PDF/PNG/DWG render changes scale, page box, or entity placement afterward.
- Source-level collision checks are not enough for this rule. Final acceptance must include a machine-readable annotation ownership audit and either a machine-readable `leader-target-binding` / `dimension-anchor` / `balloon-bom-crosswalk` report or a rendered-review record with overlay crops that identify high-risk annotations and their target geometry. The final acceptance record must include the counts `unowned_free_text_count=0`, `unsupported_floating_text_count=0`, `unbound_scattered_text_count=0`, and `dimension_like_text_without_anchor_count=0`.
- A rendered screenshot with mobile preview chrome, compressed chat scaling, or viewer controls cannot be the only proof of annotation binding. Use the exported PDF/PNG/DWG sheet render and, when needed, overlay the leader endpoints, target bounding boxes, and BOM row ids.
- All annotation-binding evidence JSON must parse with a standard JSON parser. A pass-shaped report with invalid backslash escapes, mojibake paths, stale final paths, or unopenable evidence files is a failed evidence surface.

### CORE-FIGURE-011. Mechanical CAD Appendices Must Embed Final Sheet Renders (Mandatory)

When a thesis or design说明书 is delivered together with mechanical CAD/DWG/DXF/PDF drawings, the appendix cannot pass from text-only drawing descriptions, schematic figures, old preview images, or body thumbnails. The appendix must contain the real final sheet renders from the exact final CAD package, and the audit must compare embedded DOCX media hashes against the final CAD PNG/PDF-render source hashes.

If the current user explicitly says CAD-converted images cannot be used in the Word manuscript, this current-user instruction overrides the default appendix-embed requirement for that run. In that case, do not insert CAD-rendered PNG/PDF previews into the Word正文 or appendix; instead, the Word appendix should list the drawing titles, sheet sizes, and separate DWG/PDF package names, and the final delivery package must still include the real DWG/DXF/PDF drawings with their own CAD package audit.

The CAD appendix must not use schematic, concept, sketch, placeholder, or sample-image wording or media to stand in for final sheet renders. When a real CAD-derived sheet render is embedded, captions and nearby prose should identify the source sheet, view, or drawing number, not call the formal sheet a schematic or sample figure.

- For mechanical CAD graduation projects, create a figure inventory row for every final CAD sheet that is expected in the appendix; at minimum, the appendix must cover the required A0-equivalent workload, and when the user asks for the complete drawing package, every final sheet render must be represented.
- Appendix drawing rows must name the exact source CAD package, sheet code, sheet size, title, source PNG/PDF render path, embedded media relationship, embedded media SHA256, and rendered thesis-page evidence.
- Do not accept an appendix that only says drawings are supplied separately, lists three drawing titles, or inserts small body thumbnails while the actual appendix has no sheet images.
- Do not accept old preview hashes from earlier drawing packages after the CAD package has been revised; final DOCX media hashes must be bound to the current final CAD package hashes.
- Captions and nearby wording must not call final CAD sheets `示意图`, `样例图`, or `简图`; if body explanatory diagrams remain in the manuscript, they must be clearly separated from the appendix drawing package and cannot satisfy the appendix drawing requirement.
- When the final CAD package is expanded or revised, stale workload wording such as `三张A0等效工作量图纸`, `3.5张A0等效工作量`, or `Three A0-equivalent drawings` must be treated as a hard appendix-binding failure unless it is updated to the current final sheet count and A0-equivalent workload.
- Run `scripts/audit_docx_cad_appendix_binding.py` on the exact final DOCX and exact final CAD package/folder before handoff. A pass claim is invalid when it lacks the appendix binding report path, final DOCX SHA256, final CAD package/folder path, matched sheet count, missing sheet count, and final verdict.

### CORE-FIGURE-012. Mechanical CAD Appendix Does Not Satisfy Body Figure Requirements (Mandatory)

When a mechanical-design thesis includes a complete drawing appendix, the body still needs explanatory figures wherever the design narrative, calculation chapter, structural chapter, or user feedback expects visual support. Appendix sheet renders prove drawing-package completeness; they do not replace body figures that explain structure, flow paths, heat-transfer arrangement, force/checking locations, or drawing-system relationships.

- Do not mark body figure coverage as complete merely because every CAD sheet render appears in the appendix.
- For mechanical CAD/design说明书 topics, create separate inventory rows for正文 explanatory figures and appendix drawing sheets. The same CAD render may be cited by the body, but a body figure must be physically embedded in a正文 figure block with a formal caption and nearby explanatory prose.
- If the manuscript has no正文 figures after TOC and before references/appendix, treat the body figure surface as incomplete unless the user or school template explicitly forbids body figures.
- Run `scripts/audit_docx_figure_extents.py --min-body-figure-count <required>` on the exact final DOCX before handoff when body figures are required. The report must bind `body_figure_count`, `formal_caption_count`, final DOCX SHA256, and `min_body_figure_count`.
- The final acceptance record for this surface must expose `body figure audit path`, `body figure audit verdict`, `body figure count`, `body figure minimum count`, and `body figure final DOCX SHA256` so appendix drawing evidence cannot silently substitute for missing正文 figures.
- Appendix CAD binding evidence from `scripts/audit_docx_cad_appendix_binding.py` and body figure evidence from `scripts/audit_docx_figure_extents.py` are separate acceptance surfaces. Passing one cannot substitute for the other.
- A body figure that is an authorized excerpt from a verified mechanical CAD drawing package should be marked as `mechanical-cad` / `verified-cad-png` in the figure manifest. It is not a draw.io structural diagram and must not be forced through the draw.io/SVG structural-figure contract when the authoritative source is the CAD drawing package itself.
- Mechanical CAD body figures and appendix renders must also keep nearby prose consistent with their source family. Do not describe a raster CAD sheet, CAD excerpt, or force-location sketch with structural-diagram trigger wording such as `结构`, `关系`, `流程`, `链路`, `步骤`, `过程`, `处理`, `顺序`, `阶段`, `架构`, `模块`, `功能`, `层次`, `实体`, `用例`, `时序`, `diagram`, `draw.io`, or `svg` immediately around a figure reference unless the manifest actually contains a matching draw.io/SVG `diagrams` entry. Use mechanical drawing wording such as dimensions, installation position, force point, section, assembly position, and drawing number instead. The `validate_figure_manifest` structural-signal gate must remain fail-closed: a final DOCX with nonzero structural_context_count and no matching diagram manifest entry cannot be accepted as a mechanical-CAD body figure pass.
- When the user says正文插图不要 CAD 转图、CAD 转出来的图不能用、or equivalent, formal CAD sheet renders are allowed only in the appendix drawing section or the separate DWG/DXF/PDF drawing package. Body figures must be non-CAD explanatory diagrams that match the surrounding article text, such as scheme relationship diagrams, force decomposition sketches, reducer transmission schematic diagrams, parameter-flow diagrams, or calculation-model sketches.
- A body figure that is merely a full CAD sheet render, a cropped CAD title-block/sheet preview, or a PNG/PDF converted from the formal DWG/DXF package fails this user surface even if the same image is valid for the appendix. If body figures are kept, each must have a caption and nearby prose that matches the exact calculation or structural explanation it supports.
- It is acceptable to reduce the number of body figures when the remaining figures are relevant and readable. Do not keep extra body figures only to show the CAD sheets; keep CAD sheet renders in the appendix and in the drawing package evidence.

### CORE-FIGURE-015. Body Figures Must Not Use CAD Sheet Renders When User Rejects CAD-Converted Images (Mandatory)

When the user rejects CAD-converted images in the Word manuscript, CAD sheet renders are allowed only in the appendix when the user explicitly approves appendix renders; otherwise they belong only in the separate DWG/DXF/PDF drawing package. Body figures must be non-CAD explanatory diagrams that match the surrounding article text. The run may reduce the body figure count, but every retained body figure must support nearby正文 prose and must not be a formal DWG/DXF package render, cropped CAD title-block image, or CAD package PNG/PDF preview.

### CORE-FIGURE-013. CAD Source Linework Differentiation Must Be Proven At Source Level (Mandatory)

When the user asks for mechanical drawings to be changed, redrawn, fully refreshed, made different from another package, or specifically says the lines cannot look the same, the acceptance surface is the CAD source linework, not only rendered PDF or PNG output. The run must compare final DXF/DWG source files against the locked baseline or user-provided reference package and record a `source-linework differentiation audit` before claiming the drawing package is new.

- A PDF-only change is not evidence of CAD source redesign. The audit must bind source CAD package SHA256, final CAD package SHA256, combined PDF SHA256 when present, and sheet-level DXF hashes. If source CAD files are unchanged or unbound, the `PDF-only change rejection` verdict must fail even when the PDF bytes differ.
- A whole-package redraw request cannot pass with identical source sheets. The audit must expose `changed source sheet count`, `identical source sheet count`, `changed source entity count`, per-target-sheet entity deltas, layer/entity-family deltas, and `minor-entity-move-only rejection` so moving one or two entities, title text, or a border cannot satisfy a source-linework complaint.
- Target sheets named by the user, or obvious high-risk sheets such as assembly, bundle/end-view, tubesheet, and hole-array sheets, must have a minimum source geometry delta. If the baseline contains an old-like large circle/ellipse overlay around dense tube-hole arrays, the final source must remove or materially redesign that motif and the audit must expose `old-like large circle overlay count = 0`.
- Final PDF/PNG renders must be traceable to the audited CAD source package through a `source-to-render derivation` record or equivalent package manifest. A rendered preview that is not rebound to the final CAD source hashes is stale evidence and cannot close a linework defect.
- Run `scripts/audit_cad_source_differentiation.py` whenever CAD linework differentiation is in scope, and bind the report path, report verdict, source/final package hashes, changed source sheet count, changed source entity count, PDF-only change rejection verdict, minor-entity-move-only rejection verdict, source-to-render derivation verdict, and old-like large circle overlay count before handoff.

### CORE-FIGURE-014. CAD Color Families Must Distinguish Line Families (Mandatory)

When the user supplies a CAD color sample or complains that drawing lines look undifferentiated, CAD source colors are an acceptance surface, not cosmetic preview settings. The final DXF/DWG source and derived PDF/PNG renders must use color families that make each line family visually distinguishable while preserving the existing lineweight and linetype fidelity.

- The default mechanical CAD color-family standard for dark CAD review is: `细实线` / `OBJECT_THIN` / `THIN_SOLID` remains white only; `粗实线`, `粗实线1`, `边框线`, `OBJECT_THICK`, and `FRAME` use bright green; `标注`, `尺寸`, `文字`, `DIM`, `LEADER`, `TEXT`, and `BOM_GRID` use cyan; `中心线` / `CENTER` uses yellow; `虚线` / `HIDDEN` uses magenta; `剖面线` / `HATCH` uses gray or another clearly non-white hatch color. Layer `0` must also be non-white so stray BYLAYER entities do not remain white.
- Only the thin solid family may be white. Any non-thin layer, entity override, true-color override, block entity, annotation, dimension, frame, centerline, hidden line, hatch/section line, table grid, title-block line, or support layer that remains white is a hard failure.
- Color-family changes must not be used to hide source regressions. The run must preserve geometry, lineweights, linetypes, linetype scales, text positions, and title-block contents unless a separate authorized drawing mutation owns those surfaces.
- Generated CAD sources should use layer-controlled color, lineweight, and linetype by default. Entity-level color or true-color overrides are allowed only when they are recorded, intentional, and exactly match the owning line-family standard; stale true-color values such as near-green instead of bright green or near-gray instead of the required hatch gray are hard failures. The audit must expose a `BYLAYER or exact family overrides` verdict so a dark-review PDF cannot hide source-layer color drift.
- Color-family changes must preserve or repair source lineweight fidelity at the same time. Any non-white title-block/frame/thick layer, including `TITLE`, that remains below the thick-family lineweight threshold fails even when its color family is correct.
- PDF/PNG renders made for this color-family standard should use a dark CAD review background or another recorded preview background where white thin solid lines remain visible. A white-background plot where white thin lines disappear cannot close this rule by itself.
- Run `scripts/audit_cad_dxf_color_family_standard.py` on the exact final DXF folder, DXF file, or ZIP whenever CAD colors are in scope, and bind the color-family audit report path, verdict, package/path SHA256, thin-solid-white-only verdict, non-thin-white count, expected color family count, and entity color override count before handoff. Also keep the existing lineweight/linetype fidelity audit bound to the same final source package.
- CAD audit evidence files are machine evidence, not console notes. The final handoff must not cite a CAD linework, package, color-family, rendered-readability, frame-overflow, or overlap report when the referenced JSON file cannot be parsed by a standard JSON parser. If stdout shows pass but the `--report-json` file is invalid because of Windows path escaping, mojibake, or a truncated write, rerun or repair the canonical audit script first and regenerate the report. A parse-invalid JSON file is a hard evidence failure even when the drawing package itself appears visually acceptable.

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
