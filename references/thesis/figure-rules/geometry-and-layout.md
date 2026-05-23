# Thesis Figure Rules: Geometry And Layout

Use this file for connector geometry, ER layout, tree layout, text containment, and dense-figure collision prevention.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current figure subtask.
- Apply this file together with `references/thesis/thesis-figure-generation-rules.md`.

## Final DOCX Paragraph-Margin Width Rule

For ordinary body figures, runtime screenshots, result screenshots, and structural diagrams, the default final DOCX display width is the available paragraph text width, not a small readability floor.

- Compute the body text width from section page width minus left and right margins; if the specific figure paragraph has narrower left/right indent, use that paragraph content width as the stricter baseline.
- Insert final body figures near the paragraph margin width unless the active template, locked sample, or current user instruction explicitly records a smaller accepted width.
- Treat `8.0 cm` and `9.0 cm` as minimum readability floors only. They do not authorize a figure that is visibly compressed relative to the surrounding paragraph width.
- Final evidence must report text width, displayed figure width, width/text-width ratio, and rendered-page review after reflow.

## Connector Geometry Rule

Structural figures must satisfy connector geometry correctness before they can be inserted into the thesis.

### Hard Constraints

Every connector line must obey all of these rules:

- it must start from a real source node boundary
- it must end at a real target node boundary
- it must be a boundary-bound draw.io edge with real visible `source` and `target` nodes, not a free line, center-to-center line, invisible router point, or source/targetless manual segment
- it must use orthogonal routing with right-angle bends; in draw.io terms use `edgeStyle=orthogonalEdgeStyle` with `rounded=0`, and redraw any connector that needs a diagonal or curved leg
- it must not extend beyond the intended target or point into empty space
- it must not terminate at a location where no node or relation symbol exists
- it must not cross through boxes, diamonds, ellipses, or text unless the sample explicitly requires that layout
- it must not merge into nearby borders in a way that makes the target ambiguous

### Boundary-To-Boundary Orthogonal Connector Law

For thesis flowcharts and structural figures, connector geometry is a hard acceptance surface, not a visual preference.

- A connector may leave a shape only from that shape's perimeter and may enter the next shape only at the next shape's perimeter.
- A connector must never run through the interior of its source, target, an intermediate process box, a decision diamond, an ellipse, an unrelated group frame, or any text-bearing node.
- A connector must not borrow a box border, page edge, or container frame as a routing lane.
- A connector must be drawn as an orthogonal draw.io edge between visible nodes. Freehand lines, center-to-center strokes, invisible point routers, and source/targetless `mxPoint` routes are rejected.
- Every bend must be a right angle. If a route appears to need a diagonal, curved, or rounded corner to avoid another node, the layout must be expanded or rearranged instead.
- When an edge cannot connect boundary-to-boundary without crossing a frame, the figure must be relaid out before export; do not hide the violation by making the line thinner, moving it behind the box, or relying on arrowheads.

### Separation Rule

Lines, boxes, diamonds, ellipses, and text must keep visible separation.

Treat these as failures:

- line-box overlap that makes the box border hard to read
- line-text overlap that reduces text legibility
- line-shape overlap that makes entity or relation boundaries unclear
- multiple connectors tangling together without clear branching logic
- branch lines leaving the layout area or visually implying a connection that does not exist

### Required Review

After drawing a structural figure, explicitly verify:

1. every line has a valid destination
2. every branch corresponds to a real node or relation symbol
3. every connector is a boundary-to-boundary orthogonal edge with right-angle bends
4. no connector visually collides with text or shape borders
5. no connector route crosses the interior of any unrelated frame, box, diamond, ellipse, or text node
6. the final geometry still matches the stored sample style

If any of these checks fail, the figure must be redrawn or relaid out before insertion.

## Tree And ER Layout Rule

When drawing thesis structure trees or ER figures, layout hierarchy must be validated separately from generic connector correctness.

### Tree Figure Closure Rule

For tree-style architecture or module figures:

- the parent horizontal branch line must stop at the outermost child branch centers; it must not extend past them
- each child connector must terminate at the centerline of its child box top edge
- no branch line may visually continue beyond the valid child range and imply nonexistent children
- sibling boxes must have enough spacing so vertical branch lines do not visually merge into neighboring borders

Treat these as failures:

- the top branch line extends farther than the actual leftmost or rightmost child target
- vertical child lines land between boxes or beyond box width
- the tree branch visually overshoots the last child and looks unfinished or misleading

### Grouped Architecture Layout Rule

For grouped layered architecture figures that follow the locked architecture screenshot family:

- the outer boundary frame must fully enclose every inner frame, side pillar, module box, and layer box with visible padding
- the centered grouped core frame must remain visually centered within the outer frame rather than drifting to one side
- left and right side pillars must stay vertically aligned and should not exceed the visual weight of the centered grouped core
- inner module rectangles should align to a stable row or column rhythm; do not stagger them arbitrarily
- wide horizontal layer boxes must remain fully inside the grouped core frame and keep visible padding above and below
- if a short top-centered internal system/container label is used, reserve a dedicated top band for it inside the outer frame so it does not collide with the grouped core
- containment should explain the hierarchy first; do not add connector routes that cut across compartments unless the content truly requires them

Treat these as failures:

- an inner frame, side pillar, or layer box touches or crosses the outer frame border
- the grouped core is visibly off-center without a content-driven reason
- side pillars float independently and no longer read as part of the same architecture composition
- horizontal layer boxes are so tall or so close together that the architecture loses its grouped textbook rhythm
- a top internal system/container label collides with the grouped core or reads like an accidental in-image caption
- connectors cut through boxed compartments that should be explained by containment and alignment instead

### ER Figure Layer Separation Rule

For ER figures:

- entity rectangles, relation diamonds, and attribute ellipses must occupy separate layout bands or channels
- relation diamonds must not sit directly on top of attribute ellipses or their connector stems
- main entity-to-relation lines and attribute lines must use different spatial lanes whenever overlap would occur
- if an entity has many attributes, fan them out or move them to one side instead of stacking them under a busy relation junction

Treat these as failures:

- a relation diamond overlaps an attribute ellipse
- an attribute connector passes through a relation symbol
- a main relationship line shares the same center axis as multiple attribute stems and causes ambiguity
- box, line, diamond, and ellipse boundaries become hard to distinguish because of crowding

### ER Sample-Style Density Rule

The stored ER sample for this skill is not a dense schema diagram. It uses selective attributes and restrained node counts.

When matching that sample:

- do not attempt to show every available field from the entity class or table
- show only the small set of attributes needed for the thesis explanation
- prefer fewer, cleaner attributes over a visually crowded but “more complete” figure
- if the real project has too many important entities for one clean sample-style figure, split the ER content into multiple figures instead of collapsing the grammar into direct line labels or field-heavy rectangles

Treat these as failures:

- using attribute reduction only in some entities while turning other entities into mini schema tables
- preserving all semantics in one crowded canvas at the cost of abandoning the sample's light visual rhythm
- replacing detached attribute ellipses with embedded fields because the figure became too dense

### ER Semantics Boundary Rule

For thesis ER figures, semantic correctness must be checked before any layout polish.

- only draw a solid entity-to-entity relationship when the current codebase or schema contains a real persisted linking field such as `customerId`, `contactId`, `userId`, or an explicit join table
- do not draw workflow progression, lifecycle conversion, or business-stage evolution as if it were a database relationship
- do not turn likely future relations or domain assumptions into current ER links unless the repository actually persists them
- when a business flow such as `Lead -> Customer` or `Lead -> Opportunity` is important to explain, present it as a dashed note, explanatory label, or separate workflow figure instead of a solid ER relation
- if a support entity is currently independent in the codebase, keep it independent in the ER figure even if a more complete production system would probably relate it to another table

Treat these as hard failures:

- drawing `Lead -> Customer`, `Lead -> Opportunity`, or similar conversion flow as a database foreign-key relation when only service-layer conversion exists
- drawing `Course -> Contract` or another business-plausible link when no persisted key or join structure exists in the current project
- mixing real persisted relations and explanatory workflow arrows in the same visual grammar so the reader cannot distinguish schema truth from process semantics
- using an ER figure to compensate for a missing business-flow diagram

### Required Geometry Review Addendum

For every structural figure, explicitly check:

1. branch lines are closed to the real child span in tree figures
2. ER relations and attributes use separate spatial channels
3. every connector uses boundary-to-boundary orthogonal routing with right-angle bends
4. no node border is crossed by an unrelated connector
5. no connector endpoint appears to target empty space, a node interior, or an unintended symbol

### ER Side-Channel Preference

When an ER entity has both relationship lines and multiple attributes, prefer this default layout:

- attributes placed on the left or right side of the entity in a dedicated attribute channel
- relationship diamonds placed only in the horizontal relation lane between entities
- avoid placing attributes directly below an entity if a relation lane or another relation diamond exists nearby

Treat this as a hard failure:

- an attribute ellipse touches or crowds a relation diamond
- an attribute connector shares the same bend or vertical path as a main relation connector
- an entity border becomes a transit path for unrelated lines because attributes were placed too close to the relation area

### ER Vertical Clearance Rule

When an ER attribute group sits above or below another relationship lane, leave at least one full symbol-height of clear vertical space between them.

Do not treat tiny gaps as acceptable just because shapes do not mathematically intersect.

Treat this as a failure:

- local zoom still makes attributes, connectors, and lower relation diamonds feel visually stacked or compressed
- a viewer can plausibly read two separate lanes as one crowded junction

### CORE-FIGURE-001. Structural Figure Geometry Collision Evidence Must Be Validator-Backed (Mandatory)

The ER layer-separation rules above are not complete until they are measured in the source drawing and preserved after insertion.

For every thesis ER diagram, the figure lane must produce a source-scale geometry record before insertion. The record must include:

- draw.io source path and SHA256-equivalent provenance in the figure task or manifest
- source-scale shape bounding boxes for every entity rectangle, relation diamond, and attribute ellipse
- relation-attribute collision verdict
- all-shape overlap verdict
- relation-to-attribute minimum-clearance threshold and actual clearance values
- dense-zone crop evidence paths for every local relation/attribute channel
- inserted-scale rendered-page geometry evidence after DOCX replacement

The canonical source-scale validator is `scripts/validate_structural_figure_geometry.py`. For ER figures, `scripts/thesis_figure_contract.py` must also reject a manifest when an ER draw.io source has a relation diamond overlapping an attribute ellipse, a relation-attribute clearance below the threshold, a missing geometry validation report, missing source-scale bbox evidence, missing inserted-scale geometry/collision evidence, missing dense-zone crop evidence, or any non-pass collision / overlap / source-to-inserted geometry verdict.

Do not accept ER geometry from prose claims, full-page screenshots, source-image existence, or manual local zoom alone. Missing `geometry_validation_report`, missing source-scale shape bounding boxes, missing inserted-scale collision evidence, missing dense-zone crop evidence, missing relation-attribute collision verdict, or missing selftest coverage is an enforcement gap and blocks final acceptance.

### CORE-FIGURE-009. Structural Connectors Must Be Boundary-Bound And Orthogonal (Mandatory)

Every draw.io-backed structural figure family, including flowcharts, use-case diagrams, architecture diagrams, module trees, and ER diagrams, must pass the boundary-to-boundary orthogonal connector law before insertion or handoff.

The canonical geometry check must reject:

- connectors without real visible source and target node ids
- connectors that use invisible point/router vertices to route through or around nodes
- connectors that omit `edgeStyle=orthogonalEdgeStyle` or use rounded, curved, diagonal, or freehand routing
- source/targetless line segments that substitute for real node-to-node edges
- any connector path that crosses the interior of a non-endpoint box, diamond, ellipse, unrelated group frame, or text-bearing node

For non-ER structural figures, `scripts/validate_structural_figure_geometry.py --family flowchart|structure|use-case` must write the source-scale geometry report before Word insertion. For ER diagrams, the same script may still use the ER-specific family, but ER-specific checks do not replace the boundary-to-boundary orthogonal connector law.

## Use-Case Diagram Lane Rule

For thesis use-case diagrams:

- actors must remain outside the system boundary or use-case field
- the default layout family is actor plus use-case ellipses without an outermost system boundary, unless a stronger current-run source explicitly requires that boundary
- use cases must occupy dedicated ellipse lanes that are not reused as routing channels
- lock the actor-to-use-case adjacency matrix, node layout, and route family before final ellipse placement; do not place ellipses freely and search for routes afterward
- actor-to-use-case lines must approach the target ellipse through reserved empty space rather than through another ellipse's horizontal band
- every actor-to-use-case line must start from the visible actor contour boundary, not from empty space near the actor
- every actor-to-use-case association should use a plain arrowhead by default, pointing toward the target use-case ellipse unless a stronger current-run source explicitly overrides that convention
- the arrowhead must terminate exactly on the target ellipse boundary rather than inside the ellipse, above/below it, or in nearby blank space
- when one actor is associated with several use cases on the same side, distribute the approaches across separate outer lanes or separated bend points
- when a sparse layout allows it, prefer smooth direct fan-out lines over nested multi-bend routing
- do not let an association line grind along the top edge or bottom edge of a use-case ellipse as a pseudo-route lane
- if the layout becomes crowded, solve it by redistributing nodes and adding whitespace rather than by borrowing page edges, frame borders, or ellipse borders as routing corridors
- for lower-row use cases such as `成绩查询` and `错题回顾`, prefer a bottom or outer bypass lane instead of a straight line that cuts through a neighboring use case

Treat these as hard failures:

- a default run adds an outermost system boundary even though no stronger current-run source required one
- a teacher or student association line crosses a neighboring use-case ellipse to reach a lower target
- the right boundary of the system rectangle is reused as a shared trunk for several actor lines and the routes become visually merged
- a lower use case is readable by itself, but its connecting line still passes through the border or text area of another use case
- an actor-to-use-case line ends in empty space, stops short of the target boundary, or visually looks broken halfway through the route
- an actor-to-use-case line starts from a point that is not on the actor contour
- an actor-to-use-case line or arrowhead enters the interior of the target ellipse instead of terminating on its boundary
- an association line skims along the top or bottom edge of a neighboring use-case ellipse and reads like it is crushing or borrowing that ellipse border
- a connector uses the page edge, frame edge, or other borrowed border as an improvised routing lane
- the route geometry is so tangled that a reviewer cannot tell which actor is connected to which use case at normal page zoom
- the lines are technically legal, but the full route family still looks woven, knotted, or visually tangled

### Required Geometry Review Addendum For Use-Case Diagrams

For every thesis use-case diagram, explicitly verify:

1. each actor-to-use-case path has its own readable route
2. no use-case ellipse is intersected by an unrelated connector
3. any boundary-adjacent route still keeps visible clearance from the neighboring ellipses
4. the final route geometry remains understandable at normal thesis zoom without tracing effort
5. every route has a visually complete source-to-target path with no broken or dangling segment
6. the full line family reads as smooth fan-out or clean lane routing rather than a woven bundle
7. every actor-side connector origin is visibly attached to the actor contour
8. every target-side arrowhead visibly terminates on the intended ellipse boundary and nowhere else
9. no route is borrowing an ellipse edge, page edge, or frame edge as a transport lane

## Class Diagram And Deployment Diagram Clearance Rule

For thesis class diagrams, development views, and deployment views, text containment must be validated at both source scale and inserted thesis scale.

### Hard Constraints

- no class attribute row may touch, cross, or visually merge with the class box border or the class header separator
- class boxes must be tall enough that the last attribute row still keeps visible bottom padding
- connector labels such as protocol labels, config-key labels, or relation verbs must occupy their own reserved label lane rather than sitting on top of a line
- slanted connectors must not pass under a label closely enough to make the line and text read as one fused stroke
- when a connector approaches a cylinder, class box, or dashed resource box, the route must keep visible clearance from any nearby text label
- if a box contains multiple wrapped lines, the line spacing and inner padding must be increased before export rather than relying on Word scaling

Treat these as hard failures:

- attribute text underlining itself because the glyph baseline visually coincides with a border or separator line
- a route label such as `HTTP / JSON`, `CHROMA_PATH`, or `deepseek_api_key` touching or crossing a connector stroke
- bottom-row boxes whose text touches the border after insertion scaling
- connector lines that visually continue through a label or make the label look struck through

### Required Review Addendum

For class diagrams, development views, and deployment views, explicitly verify:

1. the lowest text row in every box still has visible bottom padding
2. no connector label shares the same pixel lane as a connector stroke
3. every label remains readable after resizing the figure to the intended thesis insertion width
4. no relation or route label is forced into a crowded junction just to keep the canvas compact

## Comprehensive Structural Figure Rules From Repeated Failures

The following constraints are mandatory for thesis ER diagrams, use-case diagrams, flowcharts, and similar structural figures.

### Figure Constraint A. Caption Separation Rule

- never draw figure titles, figure numbers, or caption text inside the image canvas
- captions belong only to the thesis document body under the inserted image

### Figure Constraint B. Boundary Safety Rule

- reserve visible outer margins on all four sides of the image canvas
- no box, ellipse, diamond, line, arrowhead, text, or cardinality label may touch the canvas edge
- do not rely on Word scaling to rescue a near-edge layout; redraw first, insert later

### Figure Constraint C. Text Containment Rule

- every label must remain fully inside its own shape or dedicated label lane
- text may not cross a border, sit on a connector, or touch a shape edge
- if the label does not fit, enlarge the shape, wrap the text, or move the node; do not squeeze it

### Figure Constraint D. Connector Endpoint Rule

- every connector must start exactly on the source shape boundary and end exactly on the target shape boundary
- do not draw center-to-center connectors when they pass through a shape interior
- for rectangles, diamonds, and ellipses, compute boundary intersections rather than approximating by eye
- if a script cannot compute or validate boundary intersections, that script must not be used for final thesis structural figures

### Figure Constraint E. Connector Non-Overlap Rule

- connectors must not pass through any unrelated box, ellipse, diamond, or text label
- avoid using one dense central channel for many connectors in structural figures
- if several connectors approach the same entity, distribute them across independent lanes or sides

### Figure Constraint F. Relationship Node Clearance Rule

- diamonds must have their own free area and may not overlap attribute ellipses or entity boxes
- cardinality labels such as 1 and n must not touch lines, borders, or nearby labels
- relation diamonds must not be placed on top of a vertical trunk that also carries an attribute node

### Figure Constraint G. Attribute Layout Rule For ER Diagrams

- attribute ellipses must be organized into clear groups or bands around the entity, not packed by rough symmetry alone
- when an entity has a lower relation chain, do not place an attribute ellipse on that same lower centerline unless there is proven clearance
- use wider canvases and separated top, side, and bottom channels instead of allowing fan-out lines to cross each other densely

### Figure Constraint H. Flowchart Arrow Rule

- arrows must terminate at the border of the next process box, not in blank space near the target
- vertical and horizontal flow segments should be continuous and unambiguous
- if a branch drops to a lower row, the final arrowhead must visibly enter the next node boundary

### Figure Constraint I. Whole-Set Replacement Rule

- when the user points out a class of figure defects, identify every figure number in that class and verify that each corresponding image in the final document was actually replaced
- do not assume fixing one sample figure fixes the whole chapter

### Figure Constraint J. Mandatory Final Review Gate

Before delivery, verify all structural figures against this checklist:

1. no in-image title or caption text exists
2. no non-white or gray-filled node exists when the required style is white-background
3. no shape is clipped by the canvas edge
4. no label is clipped, overlapped, or sitting on a connector
5. every connector starts and ends exactly on shape boundaries
6. no connector crosses an unrelated shape or label
7. relation diamonds and attribute ellipses have clear separation
8. every user-reported defective figure number has been rechecked in the final document, not just in the source image folder

If any item fails, the figure work is incomplete.

Checklist routing note:

- the thesis diagram preflight checklist and the post-drawing acceptance checklist live in `references/thesis/figure-rules/workflow-and-checklists.md`
- use this geometry file for layout legality and collision rules, and use the workflow/checklist file for the explicit go/no-go checklist pass

## Finalized Constraints From Repeated Diagram Repairs

These constraints summarize the full set of recurring failures observed in thesis figure work for this project class. Treat them as default hard rules for all future thesis diagram generation.

### A. No Center-To-Center Drawing

- do not connect shapes by drawing from one center point to another center point
- lines must be trimmed to shape boundaries at both ends
- if a connector would visually pass through a box interior, reroute it or redesign the layout

### B. No Shared Collision Channels

- do not let multiple connectors stack onto one narrow corridor if the reader can no longer tell which line belongs to which node
- when several relations converge, spread them across separate lanes or separate sides of the entity
- if a dense region still looks ambiguous at normal thesis zoom, the figure fails

### C. Bottom Chain Reservation Rule

- when an ER figure has a vertical relation chain below the main entity, reserve that lower center channel for the relation only
- do not place an attribute ellipse on that same centerline below the entity
- do not place attribute labels where they visually compete with cardinality labels or relation diamonds

### D. Shape Dominance Rule

- no ellipse, diamond, line, or label may visually cover an entity box or make its border unreadable
- entity rectangles must remain visually dominant and fully legible after all connectors are drawn
- if a connector cluster makes the rectangle edge hard to read, the layout must be widened or rerouted

### E. Thesis Insertion Safety Rule

- even if the source image looks acceptable, verify the embedded Word result at thesis scale
- if a figure approaches the page boundary after insertion, reduce insertion width or enlarge source margins before delivery
- do not approve a figure that only works in the raw PNG but clips or crowds in the final document

### F. Figure Set Verification Rule

For any chapter with multiple related figures, verify all of the following in the final document, not only in the source asset folder:

1. each target figure number was actually replaced
2. no old defective image remains in neighboring figure numbers of the same section
3. all inserted figures still satisfy boundary, text-fit, and connector rules after Word scaling
4. the figure family remains visually consistent across the whole section
