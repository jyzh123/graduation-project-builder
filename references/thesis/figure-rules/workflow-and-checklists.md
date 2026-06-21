# Thesis Figure Rules: Workflow And Checklists

Use this file for the structural-figure SOP and the final preflight or acceptance checklists.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current figure subtask.
- Apply this file together with `references/thesis/thesis-figure-generation-rules.md`.

## Thesis Structural Figure SOP

Use this SOP for architecture diagrams, module trees, ER figures, use-case diagrams, and flow-oriented thesis design figures.

If the current user has locked `material-only-reuse` mode under `CORE-FIGURE-019`, do not run this SOP as a drawing workflow. Treat structural captions as source-preserved material figures and run the material provenance, final embedded-media binding, generated-substitute rejection, and paragraph-width placement checks instead.

### Step 1: Lock The Visual Family

- inspect the active template and accepted sample thesis for real figure/image examples first
- if the template contains a usable figure/image example, lock that template example before any stored skill sample
- if the template has no usable figure/image example, record that absence and identify the closest stored skill sample
- treat the sample as the layout skeleton, not as a loose inspiration
- preserve its overall visual family before replacing labels with project content

### Step 2: Freeze The Content Scope

- use only real project modules, entities, pages, or relationships
- remove anything that cannot be defended from code, SQL, or runtime evidence
- avoid adding decorative or speculative nodes just to make the figure look fuller

### Step 3: Separate The Layers

For any structural figure, divide the canvas into stable zones before drawing:

- title and caption stay outside the image
- node area
- connector area
- attribute area when needed

Do not let these zones collapse into each other.

### Step 4: Draw Geometry Conservatively

- align nodes first
- add connectors second
- add attributes last
- align directly related parent-child nodes vertically whenever possible, so the connector can be a single straight vertical line
- if a line would need to pass through a box, diamond, ellipse, or text block, relayout instead of forcing the line through
- if a line would need several elbows, first try to move the nodes into a cleaner column or branch layout; avoid bendy detours
- prefer extra whitespace over dense packing

### Step 4A: Flowchart Terminators Are Mandatory

- For thesis flowcharts, include an explicit start terminator and an explicit end terminator unless a stronger approved sample explicitly forbids that grammar.
- The default required terminator form is a rounded or elliptical white node with a thin dark outline, consistent with the stored vertical flowchart sample.
- Do not accept a sequence of process rectangles with arrows as a complete thesis flowchart when no start or end state is shown.
- If a user reports that a thesis workflow figure has no start/end nodes, redraw the figure instead of treating the omission as an acceptable simplification.

### Step 4B: Use-Case Diagram Routing Plan Is Mandatory

Before drawing a thesis use-case diagram:

1. lock the actor-to-use-case adjacency matrix first
2. choose the outer framing rule first:
   - no outermost system boundary by default
   - only restore a boundary when the user or approved sample/template explicitly requires it
3. reserve a dedicated ellipse field for each use-case column or row
4. choose a route family for the whole figure:
   - direct fan-out arrow lines when the layout is sparse enough, especially for one actor plus one use-case column
   - clearly separated outer bypass lanes when several use cases sit on the same side
5. lock where each actor-side line will leave the actor contour and where each target-side arrowhead will land on the ellipse boundary
6. reject any layout sketch where an actor line would need to pass through another use-case ellipse, share a narrow boundary corridor, grind along an ellipse edge, or terminate ambiguously

Do not start by freely placing ellipses and then "seeing where the lines can fit later". For use-case diagrams, route planning must happen before the final node placement is accepted.

### Step 4C: Use-Case Diagram Execution Flow Is Mandatory

After the routing plan is locked, execute the use-case figure in this exact order:

1. draw all actors first, and add an outermost system boundary only when a stronger current-run source explicitly requires it
2. place all use-case ellipses without any connectors and confirm that no ellipse field is being reused as a route corridor
3. choose the simplest readable route family before drawing:
   - smooth direct fan-out arrow lines when the layout is sparse enough
   - clean separated outer lanes only when direct fan-out would create collisions
4. draw routes for one actor at a time, with each route leaving from a real actor contour point and ending with its arrowhead on the exact target ellipse boundary
5. after each actor is finished, inspect that local region at normal zoom before starting the next actor
6. if the local region is crowded, relayout nodes and add whitespace before continuing; do not solve crowding by borrowing borders or forcing edge-hugging routes
7. export the source figure and run a standalone visual review on the exported image
8. only after the exported image passes, replace the thesis figure and render the touched thesis page
9. judge acceptance from the rendered thesis page, not from the draw.io canvas alone

Treat the figure as failed if the workflow jumps from drawing straight to thesis insertion without a standalone export review, if the rendered thesis page is not checked after replacement, or if the final line family still looks knotted even though the connectors are technically legal.

### Step 4D: Grouped Architecture Layout Plan Is Mandatory

Before drawing a thesis architecture-family figure that follows the locked layered architecture sample:

1. lock whether the figure belongs to the grouped layered architecture family rather than the older tree family
2. lock the outer frame, centered grouped core frame, side-pillar usage, and horizontal layer-box count before placing module text
3. lock whether the sample requires one short top-centered internal system/container label
4. assign every module to one of the allowed grouped regions first:
   - left side pillar
   - centered grouped core
   - right side pillar
   - horizontal layer boxes
5. place compartments before adding any connectors
6. only add connectors if grouping alone cannot explain the relation
7. export the figure and compare it directly against the locked architecture sample before thesis insertion

Do not start an architecture figure by placing a root box and then improvising arrows outward. If the current family is the grouped layered architecture sample, the grouped frames must be planned first and the connector set must remain secondary.

### Step 5: Run The Four Mandatory Checks

Before insertion, check all four:

1. sample consistency
2. style compliance
3. image cleanliness
4. geometry legality

Geometry legality must include:

- valid endpoints
- tree closure
- ER channel separation
- sufficient vertical and horizontal clearance
- validator-backed source-scale collision report when the figure is ER or another dense draw.io-backed structural figure

### Step 6: Local Zoom Review

Do not judge only from the full canvas.
Review at local zoom on all dense regions.
If any cropped region still feels tangled, stacked, or ambiguous, the figure fails and must be redrawn.

### Step 7: Insert Only After Passing

A figure may enter the thesis only after:

- content is defensible
- style matches the approved sample family
- no overlay or unrelated UI exists
- no line, box, diamond, ellipse, or text collision remains
- text remains clearly readable at expected page size

### Step 7A: Thesis Replacement And Render Verification

For any repaired thesis figure, especially a repeated-failure use-case diagram:

1. identify the exact target paragraph, picture object, or media relation that the thesis currently uses
2. replace that specific target and record the replacement path
3. do not trust a successful mutation message by itself; verify the actual embedded media or the rendered page
4. export the thesis through a real renderer such as WPS, Word, or LibreOffice
5. inspect the touched rendered page at local zoom and confirm that the inserted figure is the new figure, not an older cached or unchanged image
6. verify the local reading order is `image -> formal caption -> explanatory body paragraph`, and that the explanatory body paragraph plus nearby same-block explanatory prose use body formatting and do not begin by repeating a figure/table label or label cluster such as `图5-1`, `表4-1`, `图3-4和图3-5`, or `该图和图3-5`
7. if the rendered page still shows the old geometry, treat the replacement as failed and repair the replacement path before any further drawing pass

## Thesis Diagram Preflight Checklist

Before drawing any thesis structural figure, explicitly confirm all of the following:

1. the target figure type is identified
2. the full planned structural figure family list for the current chapter is locked and counted
2A. condition-driven required figures have been inferred from the thesis body and project evidence; database concept design or entity-relationship prose has a locked `database_er_diagram` row
3. the approved visual source has been selected
3A. active template or accepted-sample figure/image examples have been inspected before any skill-internal fallback sample is used
3B. if no template figure/image sample exists, the selected skill-internal fallback sample path is recorded in the task card and manifest
4. the figure will use a white background
5. all figure text will be black
6. no in-image title or caption text will be drawn
7. the chosen authoring path is draw.io source plus draw.io export, unless `material-only-reuse` mode is explicitly locked for this paper
8. connector geometry will be boundary-to-boundary, orthogonal, vertical-first when possible, and right-angled rather than center-to-center, diagonal, curved, source/targetless, bend-heavy, or routed through frames
9. a post-drawing connector/collision check will be run and will reject through-frame, through-node, free-line, invisible-router, and non-orthogonal routes
10. a post-insertion placement check will be run
10A. a post-insertion caption-sibling body check will verify body donor formatting and no repeated figure/table label prefix in the first explanatory paragraph
11. every flowchart will include visible start and end terminator nodes unless a stronger sample explicitly overrides that grammar
12. for a use-case diagram, the actor-to-use-case routing sketch has been locked before the first connector is drawn
13. the replacement verification path for the final thesis page is known in advance
14. for a sparse use-case layout, the default target is smooth direct fan-out rather than nested orthogonal routing
15. for a use-case diagram, the default framing choice has been checked and no outermost system boundary will be added unless a stronger current-run source explicitly requires it
16. for a use-case diagram, actor-side connector origins and target-side arrowhead landing points are locked before drawing starts
17. if the figure is a draw.io-backed structural figure, the intended thesis insertion path is SVG-primary rather than PNG-only
18. if the figure contains parent frames, lanes, or layer boxes, the child-box padding plan is locked before export
19. for architecture-family figures, the outer-frame / grouped-core / side-pillar composition is locked before any module text is placed
20. for architecture-family figures, the run has explicitly chosen whether one short top-centered internal system/container label is required by the approved sample
21. for architecture-family figures, connectors have been classified as necessary or unnecessary before drawing starts
22. the figure asset manifest has been generated by the canonical skill script and every structural figure entry has draw.io, SVG, and raster fallback evidence before Word insertion
22A. item 22 is the default structural-figure path; if the entry is locked as `material-only-reuse`, the manifest must instead bind source material, extracted image, final embedded media, no-redraw override, and generated-substitute rejection evidence
23. for ER diagrams and non-ER structural figures, `scripts/validate_structural_figure_geometry.py` or the equivalent canonical figure-contract path will write a source-scale geometry report before Word insertion
23A. every workflow/process/step/chain/sequence figure has been classified as `flowchart` in the manifest, and the flowchart row names draw.io source, SVG export, raster fallback, source-scale geometry report, source-to-inserted geometry evidence, post-insertion rendered evidence, final DOCX relationship evidence, pass collision verdict, pass rendered-page status, and pass/inserted insertion status, unless `CORE-FIGURE-019` material-only reuse is locked and the row instead binds primary/supplemental material provenance plus final embedded-media hash evidence
24. every teacher/user comment or existing caption that mentions figure work has been converted into a figure inventory row and a filled task card
25. every figure task card records comment/source anchor, selected sample lock, draw.io/SVG/raster evidence when structural, source-scale geometry validation report when ER/dense structural, pre-insertion evidence, post-insertion rendered evidence, dense-zone crop evidence when applicable, final DOCX relationship evidence, and a pass/fail/skipped verdict
26. no generated image created during a broad thesis rewrite remains unclassified outside the figure inventory
27. when material-only reuse is locked, every affected figure row records the material source path/SHA256, extracted image SHA256, final embedded media SHA256, missing-primary-source reason for supplemental images, generated-substitute rejection verdict, and material-only reuse verdict

If any item is not true, stop and fix the plan before drawing.

## Thesis Diagram Post-Drawing Acceptance Checklist

A structural figure may be inserted only after every item below passes:

1. white background only
2. all text black
3. no in-image title or caption text
4. no gray-filled or tinted node
5. no connector crossing unrelated shapes
6. no connector crossing labels
7. all connector endpoints land on real boundaries
7A. parent-child connectors use straight vertical lines whenever possible, and no connector keeps avoidable elbow bends or U-shaped detours
8. no clipping at the image boundary
9. labels fit cleanly inside nodes or reserved label lanes
10. rendered figure matches the approved sample family
11. inserted-page machine vision or equivalent visual inspection can still read the figure text clearly
12. every thesis flowchart shows a visible start terminator and a visible end terminator
13. every planned structural figure family for the current figure plan has its own pass, fail, or skipped status
14. no sample-locked or draw.io-first structural figure family is still represented only by a generic Mermaid draft
15. every thesis structural figure is backed by a draw.io source file and a draw.io export file unless the current user explicitly locked `material-only-reuse`
16. every repaired use-case diagram has passed both standalone export review and rendered-thesis-page review
17. the rendered thesis page visibly contains the new figure geometry rather than an older embedded image
18. the final use-case line family reads as smooth and easy to trace rather than woven or knot-like
19. no default-run use-case diagram keeps an outermost system boundary unless a stronger current-run source explicitly required it
20. every use-case association line starts on the actor contour and its arrowhead visibly terminates on the intended ellipse boundary
21. no use-case association line borrows an ellipse edge, frame edge, or page edge as a route lane
22. every draw.io-backed structural figure that has an SVG export is inserted into the thesis through an SVG-primary path
23. no child box or subordinate panel crosses beyond its parent frame on the exported image or rendered thesis page
24. every architecture-family figure still reads as a grouped layered textbook architecture figure rather than a tree or generic block diagram
25. every architecture-family figure keeps its outer frame, grouped core, side pillars, and layer boxes aligned after export and insertion
26. any allowed internal top system/container label in an architecture-family figure remains short, centered, and clearly different from a figure caption
27. architecture-family figures do not rely on decorative or excessive connector routing when grouping alone should explain the structure
28. the final DOCX contains raster-renderable image relationships for every draw.io-backed structural figure family while the manifest retains draw.io/SVG provenance evidence
29. every affected figure has a linked inventory row, task card, asset-manifest entry, and review evidence record
29A. every condition-driven required figure has a `required_figures` manifest row and a matching final DOCX caption; missing rows or missing captions are blocking failures
30. no figure-related teacher/user comment remains open without a linked figure task-card verdict
31. every ER figure has a pass source-scale geometry validation report, a pass relation-attribute collision verdict, and post-insertion dense-zone rendered evidence
32. every flowchart has a pass source-scale geometry validation report, pass collision verdict, pass source-to-inserted geometry verdict, post-insertion rendered-page evidence, and final DOCX relationship evidence, unless it is locked as `material-only-reuse`; a flowchart represented only as a generic PNG, Mermaid export, AI-generated image, Pillow output, unclassified raster, or a draw.io file containing only pasted/imported image cells is not acceptable
33. when user material is the source authority but no material-only reuse instruction exists, the manifest binds the material path, SHA256, inventory anchor, and source-match verdict, and the structural figure still remains native draw.io/SVG/fallback output rather than a pasted material image
34. when material-only reuse is locked, the manifest must prove that the embedded final DOCX media is the extracted material image and must prove that any supplemental-source image was used only after the primary material source was missing that figure

If any item fails, redraw the figure instead of patching it in Word.
