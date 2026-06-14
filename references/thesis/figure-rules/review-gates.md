# Thesis Figure Rules: Review Gates

Use this file for pre-insertion and post-insertion figure review gates, style review, cleanliness review, and text-legibility review.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current figure subtask.
- Apply this file together with `references/thesis/thesis-figure-generation-rules.md` and `references/review-figure-style-checklist.md`.

## Post-Drawing Review Gate

After any thesis figure is generated or captured, run this gate before insertion and before final delivery.

### Mandatory Check 1: Style Compliance

Check whether the figure style matches thesis expectations rather than slide-deck or product-demo aesthetics.

A figure fails style compliance if any of these are true:

- strong PPT-like colored title bars or decorative panels dominate the image
- gradients, shadows, glossy buttons, icon-heavy decoration, or presentation-style embellishment appear without an approved sample requiring them
- runtime screenshots are used in a design-only chapter where a structural diagram is required
- analytical or structural figures are drawn in a UI-card style instead of an academic diagram style
- excessive English UI labels dominate a Chinese thesis figure when Chinese labeling is expected

Default pass style when no stronger sample exists:

- black-and-white or near-monochrome academic style
- white background
- thin dark borders and connector lines
- restrained grayscale accents only when needed for readability
- centered composition and even spacing
- Chinese labels required for thesis diagrams unless the source artifact is inherently English
- internal text, line weight, and spacing rhythm consistent with the active thesis figure family

### Mandatory Check 2: Image Cleanliness

Check whether the figure surface is clean and free of unrelated overlay elements.

A figure fails cleanliness if any of these are visible:

- screenshot tool overlays
- OCR buttons such as text-extraction widgets
- browser utility overlays, floating toolbars, or extension UI
- operating-system notifications, watermarks, task-switch hints, or unrelated windows
- selection boxes, resize handles, crop marks, drawing-tool canvas chrome, or unrelated editor UI
- draw.io-export fallback text such as `Text is not SVG - cannot display`, export-help links, or other renderer fallback notices

For code screenshots, treat legitimate editor chrome as controlled by the active sample or current user instruction:

- keep it when the approved figure family is an editor-window screenshot
- crop it away when the approved figure family is code-pane-only

If either mandatory check fails, the figure is not acceptable for thesis use and must be regenerated or cleaned before insertion.

### Mandatory Check 2A: Runtime Screenshot Capture Path

Check whether a runtime screenshot came from the approved capture path rather than a historically unreliable shortcut.

A runtime screenshot fails this check if any of these are true:

- it was produced by blind headless `--screenshot` export without a rendered-content readiness gate
- it was produced by a known-bad capture path that previously generated Streamlit skeleton pages or partial placeholder renders
- it was produced by Qt/PyQt/Tk/wx or browser component rendering APIs such as `widget.render`, `widget.grab`, `QWidget.grab`, canvas export, or offscreen control painting while still being labeled as a system runtime screenshot
- it was replaced by a generated evidence page, mock page, or manually redrawn substitute while still being labeled as a screenshot
- it was replaced by a schematic illustration, UI `示意图` placeholder composition, or any non-runtime image while the chapter actually requires a real running-page screenshot
- the capture operator cannot identify the Chrome-based full-page capture path that produced the current image

Default required pass path for browser-based thesis runtime pages:

- real page running locally
- Google Chrome used as the rendering browser
- page readiness verified against key text and chart elements
- full-page screenshot captured through Chrome DevTools Protocol
- output visually reviewed before insertion

If this check fails, the screenshot must be regenerated through the approved Chrome path before it can be inserted or accepted.

### CORE-FIGURE-008. Runtime Screenshots Must Not Use Widget Render Or Grab Substitutes (Mandatory)

- A runtime screenshot for a thesis must be captured from the real running program surface, not from a toolkit's widget rendering, offscreen paint, canvas export, or component snapshot API.
- For desktop GUI programs, acceptable evidence must identify the launched executable/script, the visible window title, the OS-level or desktop-capture method, the window rectangle or monitor/desktop geometry, the accepted screenshot path, and a full-window coverage verdict.
- For browser programs, acceptable evidence must identify the local route/page URL, readiness cue, browser capture method, viewport/window geometry, accepted screenshot path, and full-page/full-window coverage verdict.
- The following capture kinds or method descriptions are hard failures for runtime screenshot slots: `real_pyqt_widget_runtime_capture`, `widget.render`, `widget.grab`, `QWidget.grab`, `QPixmap.grabWidget`, `offscreen widget render`, `canvas.toDataURL`, and equivalent component-only render paths.
- A widget-rendered image may be used only as an explicitly labeled UI design/mockup asset, never as `运行界面截图`, `系统截图`, `真实截图`, or a chapter-4 runtime evidence figure.
- If the current evidence proves only that the embedded DOCX media hash equals a widget-rendered image, the figure remains failed because hash equality proves insertion, not screenshot authenticity.

### Mandatory Check 2B: Code Screenshot Authenticity And Crop Discipline

Check whether a code screenshot is a real code capture and whether its crop matches the active figure family.

A code screenshot fails this check if any of these are true:

- it is a synthetic poster, pseudocode panel, fake editor card, or manually designed code mockup rather than a real project code capture
- it contains manual in-image titles such as `关键代码`, `核心代码`, `对应 4.3.x`, or similar explanatory labels
- it omits line numbers even though the approved code-screenshot family requires them
- it captures too little code context to read the target fragment coherently when the user explicitly asked for more complete code context
- it keeps editor window chrome even though the approved family for the current run is code-pane-only
- it is a real code screenshot, but the caption or nearby prose still describes it as `样例图` or `示意图`

Default pass pattern for thesis code screenshots:

- real code from the current project
- line numbers visible
- enough surrounding code context to read the fragment
- no artificial in-image title
- crop choice follows the current user instruction or approved sample

If this check fails, regenerate or recrop the code screenshot before insertion.

### Mandatory Check 3: Text Legibility

Check whether text inside the figure remains clearly readable at the expected thesis insertion size.

A figure fails text legibility if any of these are true:

- node labels look cramped, blurred, or too small at normal page-reading zoom
- text is technically present but machine vision cannot reliably distinguish the label content after insertion sizing
- local text density was increased just to keep the whole figure on one page
- Word or WPS insertion scaling makes labels materially less readable than the source figure
- readable output would require changing the thesis page size, page orientation, or other manuscript-level layout settings before the figure itself has been repaired
- the in-figure label size is visibly smaller than the external figure caption size on the rendered thesis page
- the in-figure labels are not bold enough to remain visually dominant against borders and whitespace at normal thesis reading zoom

If text is not clearly readable, treat the figure as failed even if geometry and styling are otherwise correct.

Hard rule for thesis structural figures:

- every in-figure label must be explicitly enlarged and bolded rather than relying on draw.io defaults
- every in-figure label must be visually no smaller than the external DOCX figure caption on the rendered thesis page
- if the rendered page shows that the caption is easier to read than the internal figure labels, the figure automatically fails review and must be redrawn or reinserted at a larger effective scale

Required readability recovery order:

1. enlarge the figure-internal font size
2. enlarge the related node, ellipse, diamond, attribute, or actor-label area
3. reduce local density or redistribute the internal layout
4. increase visual weight to bold or the approved heavier weight for the figure family
5. enlarge the useful drawing area inside the image canvas
6. increase the DOCX insertion width for the figure block while preserving pagination integrity
7. only if a higher-precedence source explicitly requires it, consider a document-level page-layout change

Do not use page-size enlargement, page-orientation changes, or manuscript-wide layout mutations as the default first fix for unreadable structural figures.

### Mandatory Check 3A: Inserted-Scale Collision Review

Do not accept a figure only because the standalone PNG or SVG looks clean at native resolution.

A figure fails inserted-scale collision review if any of these are true after thesis-width insertion sizing:

- labels that were barely clear in the source now touch borders or separator lines
- class attributes or node text look underlined because the text row and border collapse together after scaling
- connector labels and route lines overlap after insertion scaling even though they were barely separated in the source
- bottom padding inside a node disappears at thesis reading size

Required correction path:

1. preview the figure at the exact intended thesis insertion width
2. inspect every dense local region at local zoom
3. redraw the source figure with more padding or wider lanes
4. reinsert and recheck

Do not rely on changing only Word paragraph spacing or image scaling to rescue a source figure that already fails this check.

## In-Figure Language Gate

For Chinese theses, every structural figure, flowchart, ER diagram, architecture diagram, module diagram, and manually redrawn analysis figure must use Chinese labels inside the figure by default.

A figure fails this gate if any of these are true:

- English labels dominate the figure when a Chinese label would convey the same business meaning
- ER entities or attributes are shown mainly as raw table names, field names, or English schema cards instead of Chinese entity/attribute names
- flowchart and business-process nodes use English verbs such as `Start`, `Process`, `End`, `Login`, or `Admin` when they can be written as `开始`, `处理`, `结束`, `登录`, or `管理员`
- English code/database/API identifiers are kept inside the figure without a recorded necessity reason
- the final figure manifest has no explicit in-figure-language verdict

English may remain only when the label is a literal code symbol, database field, route, API name, library name, model name, command, acronym, or other technical identifier that must be preserved exactly. Even then, the figure should use a Chinese main label and place the identifier in parentheses or move the identifier into the surrounding prose/table when space is tight. The exception reason must be recorded in the figure task card and asset manifest.

For ER diagrams, use Chinese business labels for entities, attributes, and relationship names. Raw table and field identifiers such as `article`, `comments`, `user_id`, or `created_at` may be explained in body prose or an adjacent field table, but they must not dominate the ER figure surface.

### Mandatory Check 3B: Source-Scale Structural Collision Report

For ER diagrams and other dense draw.io-backed structural figures, a local crop review is not enough by itself.

Before insertion, the run must produce a validator report from the draw.io source. For ER diagrams, the report must come from `scripts/validate_structural_figure_geometry.py` or the equivalent `thesis_figure_contract.py` validation path and must record:

- source-scale shape bounding boxes
- relation diamond to attribute ellipse overlap and clearance checks
- all-shape overlap checks
- pass/fail collision verdict

The post-insertion rendered-page review must then confirm that the same dense zones remain separated at thesis scale. A pass claim is invalid when the source report is missing, the report verdict is not `pass`, dense-zone crop evidence is missing, or the final rendered page is reviewed only as a whole page without local dense-zone evidence.

### Mandatory Check 4: Post-Insertion Paragraph Safety

Check whether the paragraph that holds the inserted figure is safe for image rendering inside Word or WPS.

A figure fails post-insertion paragraph safety if any of these are true:

- the image paragraph still inherits body-text fixed line height
- the image paragraph uses body first-line indent or other body indentation residues
- the rendered page shows only a thin strip, clipped edge, or otherwise truncated image body
- the rendered page shows the paragraph after the caption split into left/right fragments beside the figure caption or image body
- the first explanatory body paragraph immediately below a formal figure/table caption begins by repeating the formal label, such as `图5-1...`, `图 5-1 ...`, or `表4-1...`, instead of using body prose such as `该图...` or `从图中可以看出...`
- the image is `wp:inline` but its holder paragraph line height is smaller than the inserted image extent, causing later text to overlay or wrap around the apparent figure block
- the figure caption paragraph carries `w:framePr` with wrapping behavior, because the following body paragraph can flow to the left and right of the caption instead of starting below it
- later insertion scaling makes the image readable in the source file but clipped in the rendered page
- the previous rendered page is left near-empty because the figure-holder paragraph, caption, or paired explanation paragraph was forced onto the next page as a detached block

Required correction path:

1. reset the figure-holder paragraph into an image-safe block paragraph
2. for inline images, set the holder paragraph to a centered no-indent block with an exact or at-least line height that is no smaller than the image height
3. remove `w:framePr` and wrap behavior from the external figure caption paragraph
4. rerender the affected page
5. confirm that the full figure body is visible and that the caption and following explanation text occupy separate normal body lines, with no repeated figure/table label prefix in the explanatory paragraph, before keeping the insertion

Do not keep a figure insertion that passes XML or media-count checks but fails rendered-page visibility.

### Mandatory Check 4B: Figure Block Pagination Integrity

Check whether the figure body, figure caption, and immediately associated explanation paragraph still form a coherent block after pagination.

A figure fails pagination integrity if any of these are true:

- the caption is visible on one page while the figure body is clipped or appears on another page
- a near-empty rendered page remains immediately before the figure because a large block was pushed forward by unsafe sizing or keep rules
- the figure body is intact, but the explanation paragraph is separated by a page break that destroys the intended reading order

Required correction path:

1. reduce the inserted figure size or redraw the source figure to a thesis-safe aspect ratio
2. re-check the figure-holder and caption paragraphs
3. rerender the affected page neighborhood
4. keep the figure only after the full block reads continuously on rendered pages

### CORE-FIGURE-002. Whole-Set Figure Caption And Pagination Gate (Mandatory)

When a thesis format run touches figures, captions, image-holder paragraphs, chapter pagination near figures, or user feedback reports figure/caption/page-break failures, the run must perform this gate on the exact active DOCX before any manuscript mutation and again before handoff.

The gate must produce a figure inventory from the DOCX package, not from remembered captions or a prior report. The inventory must reconcile:

- every `w:drawing` in the body
- every formal visible figure caption paragraph
- each embedded image relationship id and media target
- nearby source text that names a figure number
- rendered page evidence for every affected figure block

Treat the run as failed if any of these are true:

- a body figure has no adjacent formal caption and no explicit skipped-with-reason task card
- a formal figure caption contains only the number, such as `图5-1`, without a descriptive figure name
- adding a missing caption would collide with existing downstream figure numbering and no renumbering plan is locked
- drawing count, caption count, and figure inventory rows are not reconciled
- a structural figure lacks draw.io/SVG source evidence but is still claimed as accepted
- a runtime screenshot lacks route-to-asset or real capture evidence but is still claimed as accepted
- rendered evidence does not show the figure body and full caption together in the intended reading order
- rendered or OOXML evidence shows a caption-following explanatory paragraph that still starts with the formal figure/table number rather than a normal body-text lead-in

A PDF export, matching media count, valid DOCX package, or a sampled page review cannot pass this gate by itself. Final handoff is blocked until the inventory, per-figure task cards, asset/source evidence, relationship evidence, and rendered page evidence all agree.

The deterministic enforcement path for the DOCX-side part of this gate is `scripts/thesis_figure_contract.py`. Final acceptance must reject incomplete figure captions, orphan body drawings, image-caption block mismatches, and generated per-figure evidence manifests whose verdict is `pass` while the figure contract summary is failed.

### Mandatory Check 4A: Figure Caption Paragraph Integrity

Check whether the figure caption below the inserted image still behaves like a thesis caption paragraph rather than a body paragraph or a heading residue.

A figure fails caption-paragraph integrity if any of these are true:

- the visible figure number line is no longer centered like the template caption style
- the caption is merged into the same paragraph as surrounding body text
- the caption inherits heading-like outline level, bold rhythm, or spacing that disrupts the figure block
- the figure number format drifts across nearby figures in the same chapter after image replacement or reinsertion

Required correction path:

1. isolate the caption as its own paragraph block
2. reset caption alignment, spacing, and outline level to the approved figure-caption class
3. rerender the page and verify the full image-plus-caption block

Do not treat a visually close-enough caption as acceptable when its paragraph class or numbering rhythm has already drifted.

### Mandatory Check 4C: SVG-Primary Structural Figure Insertion

Check whether a draw.io-backed or otherwise vector-backed structural figure was inserted through the approved SVG-primary path instead of a PNG-only shortcut.

A structural figure fails this check if any of these are true:

- the final DOCX contains only a PNG relation for that structural figure while an approved SVG source asset exists
- the figure was reinserted through a raster-only path even though the current run produced an SVG export
- the SVG source exists on disk, but the final thesis replacement path never bound the figure block to that SVG asset

Required correction path:

1. confirm the approved SVG source asset
2. replace the figure through an insertion path that yields an SVG-primary embed
3. rerender the touched page
4. confirm that the rendered page shows the new geometry and the package still contains the SVG-backed insertion path

Do not accept a structural figure as complete when the source-of-truth is vector but the thesis still embeds only a raster insertion.

### Enforcement Rule

Do not insert or keep a figure in the final thesis when style compliance, image cleanliness, text legibility, or post-insertion paragraph safety has not been checked explicitly.

Treat these checks as required completion criteria, not optional polish.

### CORE-FIGURE-007. Figure Insertions Must Not Land In TOC Or Front Matter (Mandatory)

- Before inserting or replacing any image, prove the target anchor is outside cover, declaration, abstract, keyword, TOC, and other front-matter protected zones unless the official template itself owns that exact image surface.
- The target anchor proof must name the verified caption, heading scope, paragraph/container location, and `target_anchor_not_protected_surface_verdict`.
- This proof is required for every image mutation, including Chinese-described operations such as `替换图片`, `插入截图`, `图像`, `绘图`, or `图表`, and including mutations inferred only from source-to-final DOCX media differences.
- The exact final DOCX must be scanned for drawings, objects, or legacy pictures before the first real body chapter, including nested table/SDT/textbox paragraphs and VML/WPS `w:pict` shapes. Any unapproved image surface in this range fails the figure contract.
- Do not treat media counts, caption counts, package validity, PDF export, or a broad page screenshot as enough evidence for protected-zone safety.
- If a broad repair accidentally inserts an image into TOC/front matter, stop the figure lane and restart from a clean source or an audit-only diagnosis; do not patch around the contaminated output.

## Sample-Lock Rule

When a thesis figure needs a sample lock, the agent must inspect the active template or accepted sample thesis before using any stored skill sample. A stored thesis figure sample for the target figure type may be used only as the documented fallback after the current run records that no usable template/sample figure exists.

Required sequence:

1. inspect the active template and accepted sample thesis for real figure/image examples first
2. if a usable template/sample figure exists, extract its visual skeleton, including page occupancy, image-holder paragraph behavior, image-to-caption spacing, layout pattern, spacing logic, border style, node style, label size, and caption separation rule
3. if no usable template/sample figure exists, record that absence and identify the matching stored skill sample asset; this recorded fallback path is mandatory before any skill-internal sample is used
4. preserve the selected skeleton as the default drawing frame
5. replace only the business content with the current project's real modules, entities, relationships, or runtime pages

Do not start by freely redesigning the figure and only later try to "move it closer" to the sample. That workflow causes avoidable visual drift.

### Recorded Failure Pattern

A recurrent failure occurs when the agent prioritizes semantic mapping of real project content first, but does not treat the stored thesis sample as a hard layout constraint.

This leads to figures that are:

- content-related but visually unlike the thesis sample
- structurally rearranged in an ad hoc way
- closer to custom engineering sketches than accepted thesis figures

### Prevention Rule

For architecture diagrams, ER diagrams, module trees, use-case figures, and flowcharts:

- sample-first, content-second
- keep the sample's visual grammar stable
- only substitute project-specific labels, entities, and relationships
- if the sample and the project content conflict, resolve the conflict with minimal visual change rather than a full redraw in a new style

### Architecture Sample-Specific Lock

For thesis architecture diagrams, if no stronger current-run sample overrides it, the locked visual source is:

- `references/visual-style-samples/figures/figure-layered-architecture-sample-20260419.jpg`

Before approving any architecture-family figure, explicitly compare it against that stored sample and reject it if it drifts into another visual family.

During that comparison, explicitly check all of the following against the stored sample:

1. an outer frame exists and encloses the actual architecture body
2. a centered grouped core area exists rather than a loose box cloud
3. side pillars remain narrow and vertically aligned when the content uses side supporting modules
4. inner module rectangles are aligned and evenly spaced
5. wide horizontal layer boxes remain centered and clearly subordinate to the outer frame
6. the figure stays monochrome and white-background
7. connectors do not dominate the visual reading path when grouping alone should explain the structure
8. any allowed internal top label is short, centered, and acts only as a system/container name
9. the figure still looks like the same grouped textbook architecture family at first glance

Treat the architecture figure as failing sample lock when any of these are true:

- it reads like a root-node tree, a flowchart, or a generic block diagram instead of the grouped layered sample family
- the outer frame is present but the main modules are not organized into a centered grouped core
- arrows become the main composition device even though the sample relies on containment and alignment
- multiple internal title bars or long prose labels appear inside the canvas
- side pillars, inner group boxes, or layer boxes break alignment and no longer preserve the sample's compartment rhythm

### Architecture Inserted-Scale Review Addendum

For architecture-family figures, do not stop at a full-page glance after insertion.

Explicitly verify at inserted thesis scale that:

1. side pillars still read as distinct narrow compartments rather than oversized stray boxes
2. horizontal layer boxes still keep visible top and bottom padding
3. the top internal system/container label, if used, does not collapse into a pseudo-caption line
4. the outer frame still has visible margin around every inner compartment

If any of these checks fail, redraw the source figure instead of relying on Word scaling to rescue it.

### ER Sample-Specific Lock

For thesis database-design figures, treat `ER图`, `实体关系图`, and `数据库实体关系图` as the same ER figure family.

For thesis ER diagrams, if no stronger user-provided sample overrides it, the locked visual source is:

- `references/visual-style-samples/figures/figure-er-diagram-sample-01.svg`

Before approving any thesis ER figure, explicitly compare it against that stored sample and reject it if it drifts into another visual family.

During that comparison, explicitly check all of the following against the stored sample:

1. entities are rectangles rather than table cards
2. relations are diamonds rather than direct labeled connectors
3. attributes are ellipses rather than embedded field lists
4. labels are short, centered, and Chinese-first; for Chinese theses the visible ER entity, attribute, and relationship labels must be Chinese unless a recorded literal-identifier exception applies
5. line weight and shape borders stay thin and uniform
6. visual density stays moderate and the figure still reads as a diagram instead of a schema sheet
7. no large prose legend or explanatory memo block appears inside the image canvas unless the user explicitly required it
8. the figure still looks like the same visual family as the stored sample at first glance
9. no top title such as `ER图`, `实体关系图`, or `数据库实体关系图` remains inside the image canvas

Treat the ER figure as failing sample lock when any of these are true:

- entities are drawn as table-style schema cards rather than plain ER entity rectangles
- relations are expressed only by direct labeled lines and relationship diamonds are absent
- attributes are embedded as field lists when the sample expects detached attribute ellipses
- English/raw schema labels dominate the figure when Chinese business labels would express the same entities, attributes, or relationships
- a prose legend or explanation panel occupies a major portion of the image canvas
- the resulting figure looks closer to a database design memo, schema cheat sheet, or UML/data-model sketch than to the stored thesis ER sample
- the figure preserves semantic correctness but still abandons the sample's rectangle-diamond-ellipse grammar
- the figure still keeps an internal title even though the thesis caption already names the figure outside the image

### Additional Pre-Insertion Gate

Before accepting a generated structural figure, verify these three things together:

- it matches the real project content
- it matches the stored thesis sample style
- it does not introduce new layout inventions that make it look like a different visual family

## Text Fit And Boundary Gate

Structural figures must pass text fit and boundary validation before they can be inserted.

### Hard Constraints

- do not place the figure title inside the image canvas; thesis figure titles belong in the document caption only
- every text label must remain fully inside its own node or reserved label area with visible padding on all sides
- text bounding boxes from different nodes must not overlap each other
- text may not touch or cross a box border, ellipse border, diamond border, or connector line
- the drawn figure must keep a visible outer margin from the image boundary; no node, line, or label may be clipped by the canvas edge
- if a node cannot contain its label at the current size, the layout must be expanded or the label must be wrapped and resized before acceptance
- no structural node or panel may use gray fill, tinted fill, or shaded title bars when the required style is white-background academic output
- no inner box, child rectangle, or subordinate panel may extend beyond its parent frame, layer box, or outer boundary

### Required Pre-Insertion Review

After drawing any ER diagram, use-case diagram, flowchart, or structure diagram, explicitly check all of the following:

1. no in-image title exists
2. every label is fully contained in its shape or reserved label lane
3. no pair of labels overlap
4. no label intersects any connector or border
5. no foreground element is clipped by the image edge
6. text remains readable after expected insertion scaling
7. no explanatory paragraph, summary box, or note panel remains inside the image canvas
8. no connector passes through any unrelated shape border
9. no connector passes through the visible text area of any shape, including its own target shape
10. every connector visibly terminates on a shape boundary or a deliberately reserved connector lane rather than across the label body
11. every inner box stays fully inside its parent frame with visible padding on all sides

If any check fails, redraw the figure instead of relying on later Word scaling to hide the defect.

### Required Post-Insertion Size Review

After inserting or replacing any thesis image object in DOCX, verify the actual image object size in the final DOCX and the rendered page:

- the inserted image width must not exceed the current page text area / body版心
- the inserted image height must leave room for the external figure caption and the required follow-up explanatory paragraph
- do not rely on a later manual Word resize as the proof; the stored DOCX drawing extent must already be safe
- if the figure is clipped, partially hidden, or forces the caption/prose off the local page block, treat the figure insertion as failed even when the source image itself is valid
- the final self-check evidence must include the image-size safety result for the exact handoff DOCX when image insertion was in scope

## Recorded Failure Pattern: Text-Shape Collision And Canvas Clipping

A recurrent failure occurred when the agent used fixed coordinates to draw thesis figures but did not validate text boxes against node bounds or edge margins.

This caused:

- text touching or crossing shape borders
- nearby labels colliding with each other
- titles being drawn inside the figure itself
- shapes or labels being clipped after insertion and scaling in Word

### Prevention Rule

For thesis structural figures, the agent must treat text layout validation as a hard completion gate rather than optional polish.

Minimum required safeguards:

- reserve dedicated text channels for attribute nodes and actor labels instead of packing them by approximation
- prefer larger canvases and wider spacing over dense layouts
- run a bounding-box pass after drawing to confirm label containment, label separation, and edge clearance
- never approve a figure whose correctness depends on Word later shrinking it into place

## Recorded Failure Pattern: ER Diagram Semantics Correct But Visual Family Wrong

A recurrent failure occurs when an ER diagram uses the right entities and relationships, but is drawn as a generic module box diagram, table-like block diagram, or freeform architecture chart instead of the approved ER family.

This causes:

- entity boxes that look like system modules rather than database entities
- missing attribute ellipses
- missing relationship diamonds
- weak or absent cardinality marking
- visible drift from the stored ER sample even when the content is semantically related

### Prevention Rule

When the active figure family is the thesis ER sample:

- entities must be rectangles
- attributes must be ellipses
- relationships must be diamonds
- cardinality labels such as `1` and `n` must be explicit on the relationship lines
- do not replace this visual grammar with module cards, tables, or architecture lanes even if those are easier to generate

## Recorded Failure Pattern: Flowchart Semantics Correct But Visual Family Wrong

A recurrent failure occurs when a thesis flowchart uses the right business steps but is drawn as a generic box diagram, software architecture route chart, or slide-style process graphic instead of the locked textbook flowchart family.

This causes:

- start and end nodes that lose the rounded terminator shape
- decision points that drift into rectangles or sentence blocks instead of a centered diamond
- branch outcomes that are shoved into large note boxes rather than simple `真 / 假` labels
- branch paths that diverge or reconnect in ad hoc ways and no longer resemble the stored sample
- visible drift from the stored flowchart sample even when the content remains semantically related

### Prevention Rule

When the active figure family is the thesis flowchart sample:

- start and end must use rounded terminators
- the main decision must use a centered diamond
- branch labels such as `真` and `假` must remain outside the decision node near the branch paths
- process steps must remain plain white rectangles with uniform dark borders
- converging logic should use a clean shared merge lane before the next central step
- do not replace this visual grammar with software module cards, swimlanes, banner panels, or poster-like compositions

## Flowchart Sample-Specific Lock

For thesis flowcharts, if no stronger user-provided sample overrides it, the locked visual source is:

- `references/visual-style-samples/figures/figure-flowchart-vertical-sample-01.svg`

Before approving any thesis flowchart, explicitly compare it against that stored sample and reject it if it drifts into another visual family.

During that comparison, explicitly check all of the following against the stored sample:

1. start and end use rounded terminators
2. the decision node is a centered diamond
3. process nodes are plain white rectangles with thin-to-medium dark borders
4. branch labels such as `真 / 假` stay outside the diamond and close to the outgoing paths
5. connectors remain monochrome, straight, and visually simple
6. the branch layout stays balanced and reconverges cleanly when the logic structure converges

### Use-Case Diagram Review Addendum

For thesis use-case diagrams, explicitly check all of the following before approval:

1. every actor stays outside the system area
2. every use case remains a standalone ellipse rather than being visually cut by a connector
3. no actor-to-use-case association line passes through an unrelated use-case ellipse
4. no bottom use case such as `成绩查询` or `错题回顾` is connected through another use-case ellipse on the same row
5. the system boundary, if present, is only a boundary and not a shared traffic lane for stacked actor lines
6. right-side actors with multiple use cases use separate outer approach lanes or clearly separated bends instead of one overlapped corridor
7. no association line visibly stops short of the target ellipse or looks like it dies in open space
8. no association line is so ambiguous that a reviewer must guess which use case it belongs to
9. no association line crosses through use-case text, even if the ellipse outline itself is not fully crossed
10. no association line uses the labeled centerline of a use-case ellipse as a transit lane to another use case
11. when a shared trunk, shared lane, or central connector spine is used, the actor association line must visibly terminate on that exact trunk or lane rather than dying on the system boundary or on a nearby unconnected segment
12. when two use-case columns share one internal trunk, every branch must meet that trunk at a real shared point instead of relying on visual near-touch or antialiasing coincidence

Treat the figure as failed if any one of these checks fails, even when all labels remain readable.
7. no in-image title, legend, decorative panel, or explanatory memo appears
8. the figure still looks like the same monochrome textbook flowchart family at first glance

Treat the flowchart as failing sample lock when any of these are true:

- start or end are drawn as generic boxes instead of rounded terminators
- the decision node is replaced by a rectangle, callout, or sentence block
- branch labels are missing, buried inside shapes, or replaced by oversized note boxes
- one branch floats or reconnects in an ad hoc way that breaks the stored sample's symmetry
- the result looks like a presentation infographic or software architecture sketch instead of a thesis flowchart

## Dense-Zone Local Review Gate

Full-figure visual inspection is not sufficient for dense structural figures.

Before approving a regenerated ER diagram, use-case diagram, or similar multi-connector figure, explicitly perform a local crop review on every dense zone, including at minimum:

- every shared trunk or shared route lane in a use-case diagram
- every entity side that hosts both an attribute connector and a relationship connector
- every region where two labels, two attributes, or two relations sit in the same horizontal or vertical band

Treat the figure as failed if a local crop review shows any of the following even when the zoomed-out full figure looked acceptable:

- two attributes or labels occupying the same coordinate band and overlapping after export
- a connector that appears to attach globally but is visibly offset in the dense local crop
- a line that only seems to reach a node because of low zoom, while the local crop shows a gap, double lane, or border penetration

Required correction path:

1. export the figure
2. crop the dense local zones
3. inspect those crops at readable size
4. redraw the source figure if any dense zone fails

## Recorded Failure Pattern: Text Paragraphs Masquerading As Figures

A recurrent failure occurs when a supposed thesis figure is assembled from normal body paragraphs, arrow characters, or line-broken text blocks instead of being inserted as a real drawn object.

This causes:

- figure content that behaves like text during pagination
- captions separating from the intended figure body
- TOC and heading repair tools misreading pseudo-figure text as ordinary body paragraphs
- final pages where the "figure" is visibly just text, not an actual diagram

### Prevention Rule

For technical route diagrams, architecture diagrams, ER diagrams, use-case diagrams, and flowcharts:

- the figure body must be a real image or drawing object
- do not use stacked text paragraphs, arrow symbols, or table cells to imitate a figure
- if the figure body is not a real drawable object, the figure task is still failing even if the visible text roughly resembles the intended diagram
