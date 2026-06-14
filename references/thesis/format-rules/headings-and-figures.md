# Thesis Format Rules: Headings And Figures

Use this file for heading repair and figure-format repair rules that sit on the formatting layer.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current format-repair subtask.
- Apply this file together with `references/thesis/thesis-format-rules.md`.

## 6. Heading Rules

- Each heading level is an independent formatting surface.
- For chapter titles, second-level headings, caption lines, headers, footers, and other visually critical heading-adjacent surfaces, the approved sample's real paragraph instance is the final formatting source of truth, not the style name alone.
- If the template style definition and the approved sample's rendered paragraph instance diverge, follow the rendered paragraph instance.
- Do not declare a heading task complete only because the paragraph has the expected style name or styleId. Completion depends on rendered-page agreement with the approved sample.
- Third-level headings and fourth-level headings must also be treated as independent template-owned style classes. Do not infer their final appearance only from upper-level headings or generic body formatting.
- Second-level and third-level headings fail if the style name is correct but direct run font, size, bold, spacing, line-height, or body-format residue overrides the approved heading baseline.
- A third-level heading is still failing if the style name says `Heading 3` / `标题 3` but direct font, size, spacing, or line-height residue from a former body paragraph still overrides the approved third-level baseline.
- Bind headings to the template's real heading styles. A centered or bolded `Normal` paragraph is not an acceptable substitute.
- After any template-style import or style-XML replacement, rerun a heading-class audit. Do not assume imported template styles preserved correct heading bindings.
- If a chapter title or section heading picks up a numbering style, symbol-list style, or any other non-heading class after template import, that is a hard failure.
- Treat first-line indent and left indent on headings as explicit audit items after template import. A heading with residual indent from a numbering/list style is not acceptable even if the font and size look roughly correct.
- For first-level chapter titles specifically, verify the final paragraph against the approved sample paragraph instance item by item:
  - paragraph alignment
  - paragraph line-spacing rule and actual line spacing
  - paragraph space-before and space-after
  - page-break-before behavior
  - singular body-chapter pagination owner
  - direct run font family, size, and weight
  - numbering text spacing such as `第 1 章 绪论`
- If a first-level heading still differs materially from the sample after style binding, do not keep tuning the style definition blindly. Rewrite the paragraph's direct formatting from the approved sample instance instead.
- Body first-level chapter pagination has one owner: use the template heading/page-break owner when present, otherwise add one explicit `pageBreakBefore` owner to each first-level chapter heading; do not combine hard page-break paragraphs with `pageBreakBefore` for the same opener.
- Figure follow-up explanation paragraphs are body text, not captions. If generated follow-up text begins with a caption-like prefix such as `图 3-1`, rewrite it to a body-text reference such as `该图...` before insertion so caption detectors do not create duplicate figure captions.
- Figure follow-up explanation paragraphs are body text, not captions. If generated follow-up text begins with a caption-like prefix such as `图3-1`, `图 3-1`, or `表3-1`, rewrite it to a body-text reference such as `该图...`, `该表...`, or `从图中可以看出...` before insertion so caption detectors do not create duplicate figure captions or caption-like body prose.
- When converting a former body paragraph into a heading paragraph, explicitly clear the paragraph's direct body-format residues instead of relying on style reassignment alone.
- Mandatory heading cleanup items after body-to-heading conversion include:
  - first-line indent
  - character-unit first-line indent
  - left indent inherited from body text
  - body-style justification when the template heading should be left-aligned or centered
  - any other residual direct paragraph properties that override the heading style
- A heading paragraph that still renders with body-text indent or body-text alignment after style rebinding is a failed heading repair.
- Do not identify headings with broad heuristics across the full document such as "paragraph starts with a digit". That can capture TOC entries, numbered lists, table contents, and other non-heading lines.
- Use `officecli query` / `get` to scope candidate heading paragraphs by style, path, and nearby structure before any heading remapping pass.
- Heading remapping must run only on the verified body-text zone after the true first chapter heading, not on front matter or TOC pages.
- Before inserting any missing chapter title, verify on rendered body pages that the chapter title is actually absent. Do not insert chapter titles based only on neighboring paragraph text in the raw DOCX tree.
- Keep numbering ownership singular:
  - if the template supplies numbering, do not hardcode the same numbering into heading text
  - if the builder supplies visible numbering, strip residual `numPr` from heading styles and heading paragraphs
- If the teacher template visibly requires literal heading prefixes such as `第 1 章` / `第 1 节` / ... in the main-body heading family, generate those prefixes explicitly.
- Keep tail blocks unnumbered by default:
  - `结论`
  - `参考文献`
  - `致谢`
  - `附录`
- Third-level headings must follow the template numbering literally, for example `1.1.7`, not malformed forms like `1.17`.
- After heading repair, also check:
  - nested `pPr`
  - residual `numPr`
  - correct styleId matching such as `1 / 2 / 3` instead of only `Heading1 / Heading2 / Heading3`
  - chapter page-break behavior
- If a centered heading should have no first-line indent, explicitly zero the indent values rather than only removing an inherited indent node.
- After heading repair, explicitly verify that TOC lines still use TOC styles and that body chapter titles still use heading styles. If either side swaps roles, the repair is invalid and must be reverted.
- After heading repair, rerender at least one representative body first page and compare the real chapter-title paragraph against the approved sample page. If the page-level heading rhythm still differs materially, the heading repair is incomplete.

## 7. Figure Rules

- Figures must remain visible on rendered pages.
- Figures must not squeeze nearby body text into narrow side columns or fragmented vertical text.
- Prefer figure placement strategies that preserve paragraph readability first.
- Figure captions belong below the image in a centered paragraph.
- Treat the caption paragraph and the caption-adjacent explanatory body paragraph as two different formatting surfaces.
- The first body paragraph below a figure caption must not repeat the formal figure number as its opening token; use `该图...`, `从图中可以看出...`, or another normal body-text lead-in instead.
- Nearby body paragraphs in the same figure/table block must not begin with a figure/table-number cluster such as `图3-4和图3-5...` or `该图和图3-5...`; use `该组图...`, `上述结果...`, or another body-text lead-in so the paragraph is not read as a caption continuation.
- A figure caption may use caption styling, but any explanatory sentence such as `图 4-3 展示...` or other nearby narrative text must revert to the approved body-text class instead of inheriting the caption paragraph style, heading style, or any other figure-adjacent emphasis style.
- If a figure insertion or renumbering pass leaves the nearby explanatory paragraph centered, bold, blue, underlined, caption-sized, or otherwise visually aligned with the caption instead of the body baseline, treat that figure block as polluted and repair the body paragraph explicitly.
- Figure captions must also follow the approved sample's real caption paragraph instance. Do not treat the style name alone as proof that a caption matches the template.
- Do not rebuild figure captions or image-holder paragraphs by copying the nearest caption/holder pair from the current drifted working draft unless that draft pair has first been re-approved as the locked baseline.
- The default authority order for figure-caption and image-holder baselines is:
  - active local template paragraph instance
  - teacher-approved sample paragraph instance
  - already accepted manuscript copy
  - current working draft only after explicit visual re-approval
- A figure-repair path fails if it restores captions or image-holder paragraphs from a local draft pair that is already style-drifted.
- Treat code-title paragraphs and code blocks as independent non-body surfaces when they exist in the manuscript.
- Code titles such as `代码 1 ...` or `程序清单 1 ...` must follow the approved code-title baseline instead of falling back to body-text or heading formatting.
- Code blocks must keep their own code font family, size, indent, spacing-before/after, and line-spacing baseline instead of inheriting ordinary body paragraph formatting.
- After any body-normalization, heading-remapping, or caption cleanup pass, explicitly verify that code titles and code blocks were not rewritten into the body baseline.
- Do not push every figure onto a standalone page unless that is the deliberate fallback to prevent layout corruption.
- If the current visible artifact still shows missing figures, the figure problem remains unresolved regardless of media counts or drawing objects.
- Treat exact image and formula micro-positioning as a manual-confirmation surface unless rendering verification clearly proves it is correct.
- Thesis figure-format repair must also verify figure-internal style rather than only external placement.
- When figure-internal style, figure replacement, screenshot authenticity, flowchart grammar, ER visual family, architecture/module style, use-case routing, sequence-diagram style, or generated figure provenance is in scope, this formatting file is not enough. Load `references/thesis/thesis-figure-generation-rules.md`, the routed `references/thesis/figure-rules/` child files, and the figure task/plan templates before mutation.
- A caption/holder format pass cannot close a figure-related user comment unless the linked figure task card proves the internal figure style and source contract also passed.
- After figure insertion, the figure-holder paragraph must be treated as an image-safe block instead of inheriting body text line-height constraints. If a fixed body-text line height clips the figure body, the figure insertion fails.
- Required image-safe paragraph checks after figure insertion:
  - centered alignment
  - no body-text first-line indent
  - no inherited character-unit first-line indent or left-indent residue such as `firstLineChars` / `leftChars`
  - no fixed small line-height that clips the inline image
  - no heading-like / TOC-like / caption-like style binding or outline metadata on the image-holder paragraph
  - no body-style family binding or body-style `basedOn` chain on the image-holder paragraph when the approved sample uses a dedicated non-body holder baseline
  - surrounding spacing compatible with full image visibility
- For non-screenshot figures, internal style must be checked against the approved template or sample, including:
  - font family and visual weight of text inside the figure
  - font size hierarchy inside the figure
  - line and connector thickness
  - border thickness and border color
  - fill behavior
  - arrowhead style when arrows exist
  - internal spacing, box padding, and overall visual density
- If figure placement is correct but the figure-internal style still differs materially from the approved visual source, the figure block is still failing format review.
- If the figure exists in the DOCX but the rendered page shows only a thin strip, clipped corner, or other line-height clipping artifact, treat the figure task as failed and reset the image paragraph before any further layout tuning.
- Figure pagination review must also detect abnormal blank pages inside the local figure block, not only missing images or missing captions.
- A figure block fails pagination review if editing the figure introduces a visually blank page before the caption, after the caption, or between the figure and its required explanatory paragraph without a template-backed reason.
- A rendered page that keeps the figure/caption visible but leaves a large unexplained white area before the next thesis paragraph is still a figure-adjacent pagination defect when the template baseline does not show the same occupancy. If the user asks to fill the area with text, repair the nearby body prose under `FB-COPY-004` and rerender the exact page range before claiming the local figure block is fixed.
- Do not mark a figure-adjacent blank-space repair as complete from object counts, caption presence, PDF export, or screenshots alone. The evidence must include the inserted body-text paragraph identifiers, the affected rendered page images before/after when available, and an occupancy verdict for the figure page and the immediately following page.
