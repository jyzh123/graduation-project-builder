# Figure Style Review Checklist

Use this checklist for every thesis figure before and after insertion.

## Figure Task Intake Gate

## Pre-Drawing Gate

- [ ] figure type is identified
- [ ] the planned structural figure family list is locked and counted before drawing begins
- [ ] style source is locked (user sample / template / accepted sample / stored fallback)
- [ ] chapter semantics are identified
- [ ] figure task scope is clear (new draw / redraw / replace / insert only)
- [ ] authentic full-page system screenshots exist as evidence assets when the chapter describes real system pages or flows

## Figure Review Gate

## Pre-Insertion Review

- [ ] style matches approved source
- [ ] internal figure style matches approved source, not only caption placement or page position
- [ ] white background unless approved sample overrides
- [ ] all text is black unless approved sample overrides
- [ ] figure text remains clearly readable at expected thesis page size
- [ ] every in-figure label is explicitly enlarged and bold enough for thesis reading
- [ ] every in-figure label is visually no smaller than the external figure caption on the rendered page
- [ ] internal font family and visual weight match the figure family
- [ ] internal text-size hierarchy matches the figure family
- [ ] line weight / connector thickness / border thickness match the figure family
- [ ] fill usage and internal spacing rhythm match the figure family
- [ ] no in-image title or caption text
- [ ] no duplicated figure naming between the image canvas and the external DOCX caption
- [ ] every planned structural figure family has its own source path and review status instead of being hidden behind one completed diagram
- [ ] every thesis structural figure has a draw.io source file and a matching draw.io export asset
- [ ] every ER figure has a source-scale geometry validation report with shape bounding boxes and a pass relation-attribute collision verdict
- [ ] sample-locked or draw.io-first structural figures are not being delivered as generic Mermaid quick diagrams
- [ ] every thesis flowchart has visible start and end terminator nodes unless a stronger sample explicitly overrides that grammar
- [ ] use-case diagrams default to no outermost system boundary unless the active user instruction or approved sample/template explicitly requires one
- [ ] use-case diagrams with sparse actor-to-ellipse layouts use direct fan-out arrow routes unless a denser layout proves that another family is necessary
- [ ] sequential thesis flowcharts stay on one centered top-down vertical chain rather than a horizontal snake layout
- [ ] any sequential `证据链`, `链路`, `处理过程`, or similarly stepwise figure has been routed into the flowchart family rather than reviewed as a generic box diagram
- [ ] architecture-family figures that use the 2026-04-19 layered sample keep grouped outer-frame composition rather than falling back to a loose tree or arrow fan-out layout
- [ ] architecture-family figures keep a centered grouped core, aligned side pillars when used, and subordinate horizontal layer boxes
- [ ] architecture-family figures do not use more than one short internal top system/container label, and that label is not a caption surrogate
- [ ] ER figures match the active ER sample grammar rather than a crow's-foot field-list substitute when Chen ER lock is active
- [ ] code screenshots use real project code captures rather than synthetic code cards or pseudocode panels
- [ ] code screenshots keep line numbers when the approved family requires them
- [ ] code screenshots keep enough surrounding code context when the user explicitly asks for fuller code fragments
- [ ] code screenshot crop follows the active requirement for this run, including code-pane-only cropping when explicitly required
- [ ] connectors land on shape boundaries
- [ ] use-case actor connectors start from the visible actor contour rather than nearby blank space
- [ ] use-case arrowheads terminate exactly on the target ellipse boundary rather than inside the ellipse or in empty space
- [ ] connectors do not cross unrelated shapes or text
- [ ] use-case connectors do not grind along ellipse borders, frame borders, or page edges as improvised route lanes
- [ ] no clipping or overlap
- [ ] dense structural regions have local crop evidence and inserted-scale collision evidence, not only a whole-page screenshot
- [ ] figure text is not cramped, stacked into unreadable clusters, or visually merged by over-dense layout
- [ ] numbering and caption text are prepared correctly

## Post-Insertion Review

- [ ] rendered page shows the figure clearly
- [ ] use-case standalone export review passed before the thesis-page replacement review
- [ ] machine-vision or equivalent visual inspection can still read the figure text clearly after insertion
- [ ] rendered-page comparison confirms that the external figure caption is not visually larger or clearer than the internal figure labels
- [ ] the rendered page does not clip, truncate, or crop the figure body at page boundaries
- [ ] the rendered page does not leave a near-empty previous page while the figure body or figure-caption pair is forced onto the next page
- [ ] thesis format review includes the internal content of every inserted figure, not only page layout, caption position, or surrounding paragraph spacing
- [ ] table-like figures, code screenshots, and data screenshots are rejected if internal text is cramped, garbled, overlapping, or unreadable at normal thesis reading size
- [ ] architecture-family figures still read as grouped layered architecture diagrams after insertion, not as collapsed generic box clusters
- [ ] caption matches the figure
- [ ] captions and nearby prose do not falsely call a real screenshot `示意图` or `样例图`
- [ ] figure, caption, and required explanatory paragraph remain adjacent and in the intended order
- [ ] the first explanatory paragraph after the caption is normal body prose, matches the body donor format, and does not begin by repeating the figure/table label such as `图5-1`, `图 5-1`, or `表4-1`
- [ ] insertion did not break surrounding paragraph readability
- [ ] figure is actually embedded in the final document
- [ ] screenshot insertion decisions match chapter semantics rather than the mistaken assumption that every authentic screenshot must be inserted
- [ ] runtime screenshots do not show skeleton screens, placeholder blocks, empty loading states, or broken partial renders
- [ ] runtime screenshots were not produced by widget render/grab, offscreen component painting, canvas export, or any component-only snapshot path

## False-Pass Guard

- A figure task does not pass only because the image file exists, the DOCX contains media relationships, `officecli validate` passes, or the outline still shows expected headings.
- A figure task does not pass only because one flowchart looks finished while other planned structural figure families remain unresolved.
- A figure task does not pass when the caption remains visible but the actual flowchart body is clipped, partially off-page, or delayed to the next page after a near-empty gap.
- For runtime screenshots and inserted figures, visual rendered-page review is the acceptance source of truth.
- If the screenshot is visibly damaged, partially rendered, clipped, detached from its caption, detached from its required explanation, or otherwise unfit for thesis reading, treat the figure task as failed even when structure-level checks pass.

## Failure Rule

If any check fails, the figure task is still incomplete.
