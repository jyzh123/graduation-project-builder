# User Feedback Persistence: Template And Layout

Use this file for durable template-following, TOC, heading, front-matter, and layout corrections.

## Enforcement Status

- Every numbered rule in this file is mandatory when this file is loaded for the current subtask.
- Apply these rules together with `references/user-feedback-persistence.md`.

## Routing Note

- This file is the durable user-correction and override layer for template-following, TOC, heading, and front-matter work.
- For canonical structural rules, use:
  - `references/thesis/format-rules/front-matter-and-toc.md`
  - `references/thesis/format-rules/headings-and-figures.md`
  - `references/thesis/thesis-format-sop.md`
- Use this file to preserve durable corrections about how the canonical rules must be interpreted after repeated user feedback.

## Template And Layout Rules

### FB-LAYOUT-001 (legacy 13). Updated Local Thesis Samples Override Remembered Formatting Assumptions (Mandatory)

- If the project directory contains a newly updated thesis sample, formatting guide, or school requirement document, re-read those local files before doing more thesis formatting work.
- Do not keep using older remembered TOC, cover, or heading assumptions after the user says the local template changed.

### FB-LAYOUT-002 (legacy 14). TOC Must Match The Sample Visually, Not Just Semantically (Mandatory)

- A TOC is still wrong if it has the right headings but looks like a plain list.
- Use `references/thesis/format-rules/front-matter-and-toc.md` as the canonical structural rule source for TOC repair.
- The durable override from repeated user feedback is that semantic correctness alone is not enough; the final TOC must also visually match the accepted sample.
- Match the sample's visible pattern for hierarchy, dotted leaders, right-aligned page numbers, and title-vs-body distinction after the canonical TOC structure pass is already clean.
- Treat wrong TOC indentation, wrong level rhythm, wrong title weight, or wrong visible page-number column alignment as hard TOC failures even when the entry text and page numbers are semantically correct.
- When the user provides a TOC screenshot, treat that screenshot as the authoritative visual target.
- If a revision renames body headings, explicitly recheck the visible TOC text; TOC acceptance fails when the body heading is corrected but the TOC still shows the old title.
- If any later mutation can change page flow, including image display-size repair, body spacing repair, reference repair, or tail-block pagination repair, rerun rendered heading-page collection and compare the visible TOC page-number column against the rendered logical page map. A TOC that keeps correct dotted leaders but stale page numbers is still a failed TOC repair.
- If the user says the current manuscript TOC is visually wrong, do not treat that manuscript's TOC as the restoration baseline for the next refresh pass.
- When a school template and the current working draft disagree, the template or teacher-approved sample wins even if the draft TOC looks internally consistent.
- For legacy `.doc` school templates, assume the visible TOC may be driven by direct paragraph formatting rather than rich modern TOC styles, and extract the visible paragraph metrics from the template pages before styling the refreshed TOC.
- When a script claims to search only body headings or body paragraphs, do not accept style-name filtering alone as sufficient TOC exclusion. The targeting logic must also exclude TOC content-control paragraphs such as `w:sdt` / `InToc=True` so body rebuild steps cannot land inside the TOC block.
- When the canonical builder creates or updates a static TOC from a template, it must clone the template TOC entry paragraphs and preserve their run/tab/page-number structure; a `para.text` rewrite that keeps only the visible words is still a failed TOC build.
- When the active template shows a live TOC, TOC content control, or field-backed TOC cache, the target manuscript must derive its TOC from that same template-owned structure or record an explicit blocked static-fallback exception. A semantically correct handwritten list, default Word/WPS refreshed TOC, or static paragraph block is not acceptable for `1:1` template alignment.
- TOC evidence must record the template TOC implementation type, source paragraph/control path, field instruction when present, style/direct-format donor, per-level paragraph metrics, visible-run typography for entry/tab/page-number runs, and rendered page occupancy rhythm. A pass claim that lacks any one of those template-vs-target fields must fail closed.
- If the teacher screenshot shows a Word/WPS live TOC selection/content-control frame or grey selected entry blocks, treat that as evidence of the expected TOC editing surface, not as incidental UI decoration. The repaired DOCX must preserve the template's editable TOC behavior unless the final record explicitly says the TOC remains blocked.

### FB-LAYOUT-003 (legacy 15). Thesis Figures Must Stay In The Narrative Flow (Mandatory)

- Do not solve figure-placement issues by forcing every figure onto its own standalone page unless the template explicitly requires it.
- Figures should stay inside the relevant section, with explanatory paragraphs before or after them.
- The caption must be placed below the image in a centered paragraph.
- For implementation chapters that discuss multiple pages or modules, each screenshot must stay adjacent to the paragraph block that explains that exact page or module.
- If a subsection contains multiple screenshots, interleave them in the same discussion order instead of turning the subsection into `all prose first, all screenshots later`.
- During pagination or reflow repair, treat the explanatory paragraph block, screenshot paragraph, and caption paragraph as one local narrative unit.

### FB-LAYOUT-004 (legacy 16). Visible Final Pages Outrank Internal DOCX/PDF Evidence (Mandatory)

- Embedded media files, drawing relationships, and image xrefs are not sufficient proof that figures are correct.
- The real acceptance criterion is the user-visible final page.
- If the user says figures are still missing, assume the output is still wrong until rendered pages prove otherwise.

### FB-LAYOUT-005 (legacy 17). Do Not Mix Manual Heading Numbers With Template Numbering Systems (Mandatory)

- When following a thesis template, do not hardcode chapter or subsection numbers into heading text and also rely on template heading styles that may carry numbering semantics.
- Numbering must come from exactly one place.
- Preferred rule when using a teacher-provided template: keep heading text pure and let the template heading style system own numbering.

### FB-LAYOUT-006 (legacy 18). Standard Way To Repair Thesis Headings (Mandatory)

- The stable repair path for thesis headings is:
  - keep main-body heading text pure, without hardcoded `第1章`, `1.1`, `4.2.3` text inside the actual heading paragraphs
  - attach the main-body headings to the template's real heading styles so the navigation pane can recognize them
  - generate TOC display text separately, where numbered display like `第1章 绪论` and `1.1 研究背景` is allowed
  - never let the same numbering responsibility exist in both main-body heading text and template heading numbering at the same time
  - if a former body paragraph is converted into a heading, clear direct body-format residues such as first-line indent, character-unit first-line indent, left indent, and body-text justification instead of relying on style reassignment alone
- Treat this as the default standard way to modify headings under this skill unless the user explicitly asks for manual numbering in the main body itself.
- Operationalized implementation:
  - the main thesis builder should output pure heading text into the main body
  - a dedicated TOC-formatting step may post-process the DOCX so the TOC keeps sample-matching numbering, dotted leaders, and indentation
  - treat this two-step split as the standard heading/TOC repair flow when templates are fragile

### FB-LAYOUT-007 (legacy 19). Front Matter Should Reuse Sample Paragraph Formatting Directly (Mandatory)

- Use `references/thesis/format-rules/front-matter-and-toc.md` as the canonical structural rule source for front-matter repair.
- The durable override from repeated user feedback is that front-matter blocks should prefer copying sample paragraph formatting directly and then replacing only variable text.
- Do not restyle these front-matter blocks from scratch if a teacher-provided sample is available.
- Front-matter review must include checking the first cover paragraphs for stray inline or floating media objects, not just visible text fields.
- Preserve legitimate school banner or logo objects, but remove accidental body screenshots or other stray media that drift into the cover.
- If the content manifest supplies a thesis title, the generated cover and declaration text must not retain the template sample title or title placeholder.

### FB-LAYOUT-008 (legacy 20). Clear Inherited Headers Explicitly When Reusing Sample DOCX Files (Mandatory)

- Teacher-provided sample thesis files may carry inherited section headers from unrelated chapters or unrelated sample documents.
- When reusing a sample DOCX as the build baseline, explicitly break header links to previous sections and clear inherited headers before finalizing the document.
- Treat this as a standard cleanup step in template-following thesis generation.
- When the locked template has no visible running headers, the final DOCX must not add institutional or sample running headers on the cover, declaration/front matter, TOC, or body pages. A static-looking school-name header is still a format drift if the template header baseline is empty.

### FB-LAYOUT-009 (legacy 21). Thesis Finalizer Chains Must Not Abort On Successful `SystemExit` (Mandatory)

- When the main thesis builder launches fix-up scripts through `runpy`, catch `SystemExit(0)` and continue.
- Otherwise the build can silently stop after the first successful finalizer and leave later TOC, heading, or figure fixes unapplied.

### FB-LAYOUT-010 (legacy 22). Front Matter Repair Should Copy Sample Metrics, Not Just Text (Mandatory)

- A stable front-matter repair requires more than replacing the text.
- Copy the sample paragraph metrics for cover lines and declaration blocks, including first-line indent, line spacing, alignment, and title spacing.
- Then replace only the variable fields such as title, college, class, teacher, and date.

### FB-LAYOUT-011 (legacy 23). Title Drift Usually Means Style Binding Drift (Mandatory)

- If a title page, abstract title, chapter title, tail-block title, or section title does not match the teacher template, do not only inspect font size or alignment.
- First verify that the paragraph is still attached to the template's real heading style.
- A visually similar `Normal` paragraph should be treated as broken formatting, because it also breaks navigation recognition and makes later template-based repairs unstable.

### FB-LAYOUT-012 (legacy 23A). Body Drift Can Also Be A Style-Binding Failure, Not Only A Font-Metric Failure (Mandatory)

- If the user reports that WPS or Word shows inconsistent body styles, do not stop after comparing font size, alignment, or first-line indent.
- Also inspect whether the body paragraphs are explicitly attached to the intended body style in XML.
- A paragraph family that only looks correct because of copied direct formatting, while the office app still treats those paragraphs as style-less or default-fallback content, is a failed repair state.
- When repairing such a case, restore both:
  - the visible body metrics
  - the explicit body-style binding

### FB-LAYOUT-013 (legacy 23B). Body-Normalization Passes Must Exclude Cover Lines, Table Captions, Continuation Titles, And Table Cells (Mandatory)

- If the current pass is repairing body drift, do not let the same body-normalization logic rewrite protected non-body surfaces into the body baseline.
- The exclusion set must explicitly include:
  - cover/title-page lines
  - declaration / commitment rows
  - figure captions
  - table captions
  - continuation titles such as `续表`
  - table-cell paragraphs
- If a pass makes cover text or table text inherit the body font, body size, body first-line indent, or body justification, classify that as a failed mixed-surface repair rather than as a minor cosmetic drift.
- Cover identity tables are part of that exclusion set. They must be replayed from the active template donor, including table geometry and cell-level formatting, before target field values are refilled.
- Recovery path:
  1. restore the polluted protected surface from the approved sample paragraph instance
  2. rerun body normalization only on the verified body paragraph family
  3. rerender the affected cover/table pages before handoff

### FB-LAYOUT-014 (legacy 23C). XML-Level Body Rewrites Must Preserve Real Body Paragraph Instances, Not Only Text (Mandatory)

- If a thesis pass rewrites body text through `word/document.xml`, raw paragraph-text replacement, or another XML-level text path, do not assume that keeping the visible sentence text is enough.
- Treat every rewritten body paragraph as a full paragraph-instance surface.
- The repair must preserve or explicitly rebuild all of the following together:
  - explicit body-style binding such as `w:pStyle`
  - first-line indent
  - left/right indent
  - paragraph alignment
  - line spacing
  - spacing before/after
  - run-level font mapping needed for the active body baseline
- Do not assume the document's `Normal` style or default body style definition already carries the approved first-line indent. Many accepted theses keep part of the body baseline on the paragraph instance rather than on the style definition alone.
- If a body rewrite path clears the paragraph instance and only writes new text back, treat that path as body-format-unsafe until it proves that the full body instance baseline is restored.
- After any XML-level body rewrite, rerun the body-style audit on the exact post-edit deliverable. An earlier body-style audit becomes stale immediately once the rewritten manuscript differs again.

### FB-LAYOUT-015 (legacy 24). Heading Levels Must Follow The Template Separately (Mandatory)

- Treat each heading level as an independent formatting surface.
- Third-level headings must keep template numbering like `1.1.7`; malformed compressed numbering like `1.17` is a formatting bug.
- If the teacher template visibly shows numbered main-body chapter headings like `第1章 绪论 / 第2章 ...`, generate those chapter-heading strings explicitly when needed, while keeping terminal blocks such as `结论`, `致谢`, and `参考文献` unnumbered unless the template explicitly requires another convention.

### FB-LAYOUT-016 (legacy 25). Abstract Keyword Labels Must Be Bold, But Only The Labels (Mandatory)

- In thesis abstract pages, treat labels such as `摘要：`, `关键词：`, `Abstract:`, and `Key words:` as bold labels by default.
- Do not let these labels fall back to ordinary body-text weight during abstract assembly or final formatting passes.
- Bold the label text itself, not the entire keyword string, unless the template explicitly requires full-line bolding.
- The keyword content after the label must remain in the keyword-content/body-like donor face. It must not inherit abstract-title, heading, caption, TOC-title, or title-like character formatting even when the label/content run split is structurally correct.

### FB-LAYOUT-017 (legacy 26). Figures Must Stay Visible Without Breaking Paragraph Readability (Mandatory)

- If the user reports that figures are still missing or covered by text, treat plain inline picture insertion as unresolved.
- For thesis templates like this one, use a figure-placement strategy that keeps figures visible on rendered pages and does not squeeze nearby body text into narrow side columns or fragmented vertical text.
- Preserve complete paragraph readability even if that means changing figure placement strategy.
- The current visible artifact outranks earlier internal evidence such as embedded media counts or drawing objects.

### FB-LAYOUT-018 (legacy 27). When High-Level DOCX APIs Drop TOC Or Heading Formatting, Use XML-Level Fixes (Mandatory)

- Follow the canonical DOCX toolchain order first.
- If those high-level operations still cannot preserve TOC indentation, dotted leaders, heading spacing, or style bindings reliably, switch once to XML-level normalization instead of repeatedly tweaking high-level properties.
- For TOC, rebuild paragraph `pPr` cleanly with explicit `TOC1/TOC2`, tabs, indentation, and tab runs.
- For headings, fix style binding, strip unintended numbering, flatten nested `pPr`, and target the real template styleId rather than only English aliases like `Heading1`.

### FB-LAYOUT-019 (legacy 28). Tail Blocks Stay Unnumbered And Undivided (Mandatory)

- Terminal blocks such as `结论`, `致谢`, and `参考文献` should be treated as terminal thesis blocks, not as numbered chapters, unless the template explicitly requires numbering.
- By default they should not be split into subsection hierarchies such as `7.1` or `A.1` unless the user explicitly provides a template requiring that structure.

### FB-LAYOUT-020 (legacy 29). Do Not Mix Multiple TOC Repair Strategies In One Build Chain (Mandatory)

- A fragile TOC becomes unstable when the build uses multiple overlapping mechanisms:
  - handwritten TOC text
  - style copying
  - XML patching
  - stale template leftovers
- Pick one TOC strategy and keep it clean.
- If XML repair is used, replace TOC paragraph XML cleanly instead of layering new `pPr` fragments onto old ones.

### FB-LAYOUT-021 (legacy 30A). Template Alignment Must Be Proven On Rendered Page Classes, Not Inferred From Structure (Mandatory)

- When a local sample thesis exists, do not conclude that the current thesis matches the template after checking only heading hierarchy, TOC presence, page count, or paragraph style names.
- Required rendered comparison set before a thesis may be described as template-aligned:
  - cover
  - Chinese abstract
  - English abstract
  - TOC
  - first body chapter page
  - one figure page
  - one table page
  - references
  - acknowledgement
- If that page-class comparison has not happened on the exact final deliverable or exact review copy being handed off, the template-alignment claim is invalid.

### FB-LAYOUT-022 (legacy 30B). Front-Matter Numbering Chains Must Be Section-Verified Before TOC Acceptance (Mandatory)

- If the sample or school guide uses a front-matter numbering split such as `摘要 I / Abstract II / 目录 III / 正文 1`, do not infer that split from rendered TOC text alone.
- Before TOC acceptance, verify in a real office application that the title paragraphs for:
  - Chinese abstract
  - English abstract
  - TOC
  - first body chapter
  report the expected adjusted page numbers or section-numbering state.
- If Word COM is unstable but WPS COM is available, use WPS COM as the numbering-authority fallback instead of guessing the split from PDF pages.
- If the manuscript still uses the wrong section chain and the builder merely edits visible TOC strings to look correct, treat that as a failed repair.

### FB-LAYOUT-023 (legacy 30C). TOC Helper Text And Placeholder Leakage Are Hard Failures (Mandatory)

- Do not leave visible helper text such as `TOC_PLACEHOLDER`, `TOC_PLACEHOLDER_PASS2`, static staging entries, or other build-only markers anywhere in the thesis.
- Check both the DOCX text layer and the rendered PDF/page images.

### FB-LAYOUT-064. Unreadable Structural Figures Require Source Typography And Rendered-Page Repair (Mandatory)

- If a user says thesis figures are too small or unreadable, do not only increase the DOCX image width.
- Repair the source figure as well: increase font sizes, enlarge boxes, reduce excess whitespace, split long labels, and recheck connector/text collisions before rebuilding the document.
- Insert key structural figures near the maximum safe body text width unless the template imposes a smaller figure width.
- The canonical final-DOCX display audit must treat structural figures narrower than `9.0 cm` as unreadable by default unless a locked template/sample or explicit user instruction records a smaller accepted width.
- The same audit must treat runtime screenshots, pop-up screenshots, result screenshots, and ordinary UI images below `8.0 cm` width or `4.0 cm` height as unreadable by default; a `3.0 cm`-class technical minimum is not acceptable final thesis evidence.
- Export rendered PDF pages around the affected figures and inspect the visible page result; internal image dimensions or DOCX media counts alone do not prove readability.
- During the rendered review, also check nearby captions, explanatory paragraphs, and page headers because figure reflow can expose stale template headers or detached figure-caption blocks.
- If helper text is absent from the outline but still visible on rendered pages, the TOC task is still failing.

### FB-LAYOUT-071. Body Figures Must Align With Paragraph Text Margins By Default (Mandatory)

- Do not treat the minimum readable width (`8.0 cm` for runtime/UI images and `9.0 cm` for structural figures) as the default insertion width.
- Unless the active school template, locked sample, or current user instruction explicitly sets a smaller accepted width, body figures must be inserted close to the available text width between the paragraph left and right margins.
- Landscape and near-landscape body figures are paragraph-width figures by default for every thesis run. When width is greater than height, set the displayed width to the available text width and preserve aspect ratio; do not leave horizontal screenshots or diagrams at a small mid-page width merely because the image is already legible in the raw media file.
- The final-DOCX display audit must record the body text width and each body figure's width/text-width ratio; figures below the default ratio gate are compressed evidence, not acceptable repaired figures.
- When the user reports screenshots or body figures are too small, final acceptance must bind the canonical figure-extents JSON for the exact final DOCX SHA256 and a pass paragraph-margin width verdict. A path-only or prose-only evidence file is not enough.
- Source readability repair remains mandatory for dense structural figures. A paragraph-width insertion cannot compensate for unreadable source typography, connector collisions, or crowded labels.
- A final handoff that claims "figures fixed" must include rendered-page review after the width change so captions, page breaks, and nearby paragraphs are not broken by reflow.

### FB-LAYOUT-063. Template-Owned Styles Must Not Be Invented By The Builder (Mandatory)

- Do not create new `Heading 1`, `Heading 2`, `Heading 3`, `Normal`, TOC, abstract, caption, table, reference, or cover styles to compensate for a fragile school template unless the active template itself exposes no usable donor and the run records that blocker explicitly.
- Do not assign builder-chosen Chinese fonts, Western fonts, English fonts, font sizes, paragraph spacing, or line spacing to template-owned thesis surfaces.
- Before writing or repairing a template-owned surface, lock a real donor paragraph or table-cell instance from the active school template or an approved local sample.
- If a required donor cannot be found, stop that surface as blocked rather than approximating it with application defaults or manually created styles.
- Creating English-named heading styles only to satisfy a checker is not a valid template repair when the template uses direct formatting or local style names.
- Abstract body prefixes, English abstract content runs, TOC entry runs, table-cell runs, and chapter-opening page-break owners are template-owned donor surfaces; the builder may not substitute locally invented font or pagination policy for them.

### FB-LAYOUT-024 (legacy 30D). English Title And Abstract Must Not Split Into A Fake Nearly-Empty Page (Mandatory)

- If the English thesis title appears by itself on a mostly empty rendered page and the `Abstract` heading starts on the next page, treat that as a front-matter pagination failure unless the approved sample explicitly uses that layout.
- Correct order alone is not enough; the English title block and the English abstract block must also follow the sample's occupancy pattern.

### FB-LAYOUT-025 (legacy 30E). Runtime Screenshot Slots Must Use A Locked Caption-Route-Asset Map (Mandatory)

- For implementation chapters with runtime screenshots, do not treat image replacement as a generic "replace figure" step.
- Before replacement, lock for every runtime screenshot slot:
  - visible caption text
  - intended route URL
  - readiness cue
  - accepted screenshot asset path
  - embedded media relationship after insertion when replacing an existing image
- If two different runtime captions end up bound to the same stale structural image or to the wrong screenshot asset, the chapter fails image review even when the images themselves render cleanly.
- A route/caption/path match is still not enough: the accepted screenshot must contain real rendered page content. Blank exports, near-empty images, dominant solid-color blocks, and purple placeholder captures fail even when their dimensions and route evidence look correct.

### FB-LAYOUT-026 (legacy 48). Repair Heading And TOC Structure Before TOC Styling (Mandatory)

- Use `references/thesis/format-rules/front-matter-and-toc.md` as the canonical structural rule source for this sequence.
- The durable override here is priority: repair heading-to-TOC structure first, then do TOC styling.
- Do not jump directly into TOC indentation, dotted-leader, or font cleanup while the heading-to-TOC mapping is still wrong.

### FB-LAYOUT-027 (legacy 55). Figure Captions Must Be Targeted As Exact Caption Paragraphs, Not As Generic Figure-Number Mentions (Mandatory)

- Do not locate a figure block by broad substring matches such as any paragraph that merely contains `图 4-1` or similar text.
- Body sentences like `见图4-1` are not valid figure-caption anchors.
- Target figure replacement on the exact caption paragraph, or on a paragraph whose text clearly starts with the figure caption token.
- If an existing DOCX stores an inline image and the visible caption text in the same paragraph, split that mixed paragraph into:
  - one image-only paragraph
  - one caption-only paragraph
- Only after that split may the figure replacement or figure-format repair proceed.

### FB-LAYOUT-028 (legacy 56). Visible Figure And Table Captions Must Exclude Editorial Notes And Build Provenance (Mandatory)

- Do not leave visible caption text with editorial provenance or build-process remarks such as `（据实验结果文档整理）`, `（自动生成）`, `（草稿）`, `（占位）`, or similar parenthetical notes unless the approved sample explicitly uses that wording.
- For thesis delivery, visible caption text should normally contain only the numbering token and the approved title text for that figure or table.
- Source provenance, extraction notes, and build reminders belong in automation records, evidence logs, or operator notes, not in the reader-facing caption paragraph.
- If a builder or repair script finds a caption anchor that already includes an editorial note, strip that note before insertion, then rerender and recheck the caption on the page.
- A caption pass fails acceptance if the numbering is correct but the visible caption still carries workflow remarks that are absent from the approved template or sample.

### FB-LAYOUT-029 (legacy 57). Runtime Screenshot References Must Be Checked Against The Current Asset, Not Just The File Name (Mandatory)

- If the thesis references a runtime screenshot by path, verify that the referenced file is the current accepted screenshot rather than an older intermediate export with a similar name.
- Do not assume that the newest-looking filename such as `fixed`, `wide`, or `final` is actually the correct image without visual inspection.
- Before final DOCX export, check the exact screenshot path used in the manuscript against the current accepted image asset.
- If the main system screenshot, result screenshot, or tab screenshot is stale, cropped, clipped, or visibly wrong, replace the manuscript reference and regenerate the DOCX in the same turn.
- Do not accept a screenshot from size-only evidence. The current asset check must include rendered visual content, not just file existence, pixel dimensions, or a pass-shaped filename.

### FB-LAYOUT-030 (legacy 58). Figure And Table Object Paragraphs Must Not Inherit Body Paragraph Formatting (Mandatory)

- When global body formatting is applied, figure-holder paragraphs and table-cell paragraphs must be excluded from raw body-paragraph inheritance or explicitly reset afterward.
- Required image-holder paragraph fallback:
  - centered alignment
  - no body-text first-line indent residue
  - template-safe line spacing; if the locked holder donor has no direct spacing,
    do not synthesize direct `w:spacing` values just to make the holder look
    normalized
  - keep-with-next enabled when the next paragraph is the caption
- Required table-cell paragraph fallback:
  - no first-line indent
  - restrained academic font size
  - single line spacing
  - alignment chosen deliberately for readability rather than inherited body justification
- If figures or tables still show body-style fixed line spacing, two-character first-line indent, or paragraph justification side effects after export, treat the layout pass as failed.

### FB-LAYOUT-031 (legacy 58A). Figure Holder And Figure Caption Paragraphs Must Clear Direct Indent And List Residues (Mandatory)

- Repeated thesis repair runs can leave figure-holder paragraphs or figure-caption paragraphs with direct paragraph residues even after the visible style name looks correct.
- Treat these residues as a dedicated failure class.
- Required cleanup for every affected figure-holder paragraph:
  - centered alignment
  - clear direct first-line, character-unit first-line, left, hanging,
    `leftChars`, and `hangingChars` residues unless the locked approved holder
    donor explicitly proves those fields must be written
  - clear direct paragraph spacing when the approved holder donor inherits
    spacing from its style rather than writing direct `before`, `after`,
    `line`, or `lineRule` attributes
  - keep-with-next enabled when the next paragraph is the caption
- Do not leave a figure-holder paragraph bound to the body-text style family or to a style chain that still inherits the body paragraph's indent grammar.
- Do not repair a holder drift by writing `0` indents or fixed `120/240`
  spacing when the self-check/template baseline expects those direct fields to
  be absent; absence and explicit zero are different DOCX surfaces.
- If `Normal` or the current body style definition carries first-line indent via `w:ind w:firstLine`, `w:firstLineChars`, or related character-unit indent fields, treat a holder paragraph that still resolves through that chain as failed even when its visible text is empty.
- For the current IMUST mechanical-design baseline, a figure-holder paragraph must keep the dedicated `ThesisImageHolder` / `Thesis Image Holder` paragraph family rather than being downgraded to `Normal`. The direct holder metrics are centered alignment, `before=120`, `after=0`, `line=360`, `lineRule=auto`, `firstLine=0`, `left=0`, `right=0`, `firstLineChars=0`, `leftChars=0`, `rightChars=0`, no hanging/list/outline residue, and `keepNext` when followed by a caption. Older `line=240` holder evidence is stale for this profile because it can conflict with the non-clipping image-holder safety gate.
- Required cleanup for every affected figure-caption paragraph:
  - centered alignment
  - first-line indent explicitly set to zero
  - no inherited left indent or hanging indent
  - no list numbering, bullet numbering, or residual `numPr`
- A caption paragraph with visible centering but hidden bullet/list metadata is still a failed repair because it can reintroduce abnormal caption offset on later edits or in another renderer.

### FB-LAYOUT-032 (legacy 58B). Table Captions, Continuation Titles, And Table Cells Must Clear Body Indent Residues Separately (Mandatory)

- Table-related surfaces must not be normalized by pretending they are ordinary body paragraphs.
- Trigger this rule for user wording such as `表格样式`, `表格缩进`, `表格文字偏右`, `三线表`, `表头线`, `中线`, `TableGrid`, `continuation table`, or `continued table`.
- Required cleanup for every affected table-caption or continuation-title paragraph:
  - centered alignment when that is the approved sample behavior
  - first-line indent explicitly set to zero
  - no inherited left indent or hanging indent
  - `keepNext` enabled when the next block is the table start or continued table fragment
- Required cleanup for every affected table-cell paragraph:
  - first-line indent explicitly set to zero
  - no inherited left/right body indent residue
  - no body-text justification carried over by accident
- If a rendered page still shows a `续表` title, table caption, or table-cell line starting from the body indent column, treat the table lane as still failing.

### FB-LAYOUT-033 (legacy 58C). Cross-Page Tables Must Carry The Approved `续表` Family On Continuation Pages (Mandatory)

- When the approved local sample or teacher template uses continuation titles such as `续表 4-3`, continuation pages of the same table must reproduce that family explicitly.
- Trigger this rule for user wording such as `表格跨页`, `跨页表格`, `续表`, `续表标题`, `表格断页`, `表题和表格跨页`, `repeated header`, or `continuation title`.
- Do not treat a repeated table header row by itself as sufficient proof that the continuation page is correctly labeled.
- The continuation title must be:
  - a standalone paragraph above the continued table fragment
  - outside the table grid
  - formatted from the approved table-caption / continuation-caption baseline rather than from generic body text
- If a continuation page has no visible continuation title while the template requires one, or if the title is pushed into the table grid or into a body paragraph, the repair is incomplete.

### FB-LAYOUT-034 (legacy 58D). Figure Caption And Image-Holder Baselines Must Not Be Sourced From A Drifted Working Draft (Mandatory)

- When repairing thesis figure blocks, do not pick the current working draft's caption paragraph or image-holder paragraph as the default donor just because it is nearby and convenient.
- If the current local draft already has caption drift, image-holder line-spacing drift, or class-binding drift, reusing that local pair will amplify the error across the whole manuscript.
- Required authority order for figure-caption and image-holder baseline extraction:
  - active local template
  - teacher-approved local sample
  - already accepted manuscript copy
  - current working draft only after explicit visual proof that the pair still matches the approved baseline
- If the figure lane touches many figures at once, lock one approved caption baseline instance and one approved image-holder baseline instance before any batch rewrite starts.
- Do not accept a repair path that clones a drifted local `caption_ref` / `holder_ref` pair and then tries to correct only centering or `keep-with-next` afterward.
- A figure-repair lane is incomplete if the final manuscript still shows:
  - figure captions collapsed into the generic `Normal` family instead of the approved caption family
  - image-holder paragraphs using the wrong line-spacing rhythm relative to the approved baseline
  - apparent local offset or pseudo-indent caused by wrong holder baseline even when explicit first-line indent is already zero

### FB-LAYOUT-035 (legacy 59). Content-Only Thesis Rewrites Must Not Globally Refresh Fragile TOC Blocks (Mandatory)

- When the current run is a thesis content rewrite or argument-strengthening pass rather than a dedicated TOC/pagination repair pass, treat the front matter and TOC block as protected surfaces.
- Do not run whole-document `Fields.Update`, automatic `TOC.Update`, or equivalent global foreground refreshes on WPS/template-driven `.docx` files just because body paragraphs were rewritten.
- On fragile thesis templates, those global refreshes can pull cover, abstract, keyword, or TOC-title paragraphs into the visible TOC and then shift visible pagination.
- Safe default workflow for content-only passes:
  - rewrite only the confirmed body chapter ranges
  - leave the existing TOC block untouched unless the task explicitly includes TOC repair
  - if the TOC must be repaired, calculate heading pages first and rebuild the visible TOC as its own isolated step after structure review
- Never mix body-content rewriting and blind global TOC refresh in the same automatic pass.

### FB-LAYOUT-036 (legacy 62). Figure Description Paragraphs Must Not Break The Figure-Caption Block (Mandatory)

- When a thesis figure needs explanatory text, do not treat the image paragraph, description paragraph, and caption paragraph as independent movable fragments.
- If the chosen layout is:
  - image
  - description
  - caption
  - then these three paragraphs must be handled as one pagination block.
- If the chosen layout is:
  - description
  - image
  - caption
  - then the description must still stay attached to the figure-caption block rather than being left on the previous page by itself.
- Never let a figure caption disappear from the visible page because the image moved and the description or caption was left behind.
- Never let a figure title/caption become visually orphaned at the page top or page bottom after a content-enrichment pass.
- If the full block does not fit, move the block boundary before the figure block instead of splitting the image, description, and caption across pages arbitrarily.

### FB-LAYOUT-037 (legacy 65). Pagination Review Must Start At The Cover Page, Not The First Body Chapter (Mandatory)

- If the user reports a pagination problem, do not narrow the inspection to body chapters first.
- The stable review order is:
  1. cover page
  2. Chinese abstract page(s)
  3. English abstract page(s)
  4. TOC page(s)
  5. first body chapter page
  6. later body chapter pages
  7. end matter
- A pagination review is incomplete if it can explain body chapter breaks but has not verified whether the cover, abstracts, and TOC already spill into the wrong pages.

### FB-LAYOUT-038 (legacy 69). Caption Detection Must Not Capture Explanatory Paragraphs (Mandatory)

- Caption-detection logic must also reject explanatory body paragraphs that begin with figure or table numbers but continue directly into verbs, for example `图 4-1 展示了...` or `表 2-6 汇总了...`.
- For automated targeting, require the real caption token shape such as `图 4-1 空格 标题` or `表 2-6 空格 标题` instead of any paragraph that merely starts with `图 4-1` or `表 2-6`.
- Body prose that begins with ordinary words such as `图像预处理`, `图像识别`, or `图像融合` is not a figure caption and must not trigger a figure-manifest requirement by itself. Only an image/drawing action, target surface, or actual source-to-final media/drawing diff may escalate it to image mutation scope.
- If a layout, pagination, or replacement script marks a normal explanatory paragraph as a caption candidate, treat the run as a failure and restore the body-paragraph baseline before continuing.

### FB-LAYOUT-039 (legacy 72). Implementation Module Screenshots Must Follow Code-Then-Result Order When The Review Comment Requires It (Mandatory)

- If a teacher comment or approved sample requires `先代码，再运行截图` in an implementation chapter, do not leave the code screenshots collected only in a later standalone section while the module subsections still show runtime screenshots first.
- For each affected module, build the local figure block in this order:
  - code screenshot
  - code screenshot caption
  - runtime screenshot
  - runtime screenshot caption
- Update the surrounding body sentence so that figure references point to the new code figure number first and the runtime figure number second when both are mentioned.
- After inserting new code screenshots, renumber all downstream figure captions and update later in-text figure references in the same pass.
- If the newly inserted code screenshots are visible but later captions or `见图` references still keep the old numbers, treat the repair as failing.

### FB-LAYOUT-040 (legacy 73). Caption-Adjacent Narrative Paragraphs Must Rebind To Body Text After Figure/Table Edits (Mandatory)

- Repeated user feedback shows a recurring failure pattern: after figure or table insertion, nearby narrative paragraphs such as `图 2-13 展示了...` or `表 3-4 所示...` can accidentally inherit the caption or table-title formatting.
- Treat this as a dedicated regression class during every figure or table pass.
- Required rule:
  - caption paragraph keeps caption formatting
  - table-title paragraph keeps table-title formatting
  - nearby explanatory paragraph keeps body formatting
- Do not insert, duplicate, or rewrite explanatory paragraphs by reusing the caption style, table-title style, or a section-heading style.
- After any figure/table insertion or renumbering pass, explicitly inspect the first explanatory paragraph near the edited block and verify that it is still body text in both document-internal data and rendered-page appearance.

### FB-LAYOUT-041 (legacy 74). Abstract Blocks Must Not Be Repaired As Body Paragraphs In Disguise (Mandatory)

- Repeated user feedback confirms a failure pattern where the builder centers and bolds the abstract title, or bolds the keyword line, but leaves those paragraphs attached to the generic body-paragraph class.
- Trigger this rule for `英文摘要缩进`, `Abstract 缩进`, `English abstract indent`, `firstLineChars`, `摘要缩进`, `摘要空格`, and any user report that the abstract/keyword block lost its original formatting.
- Treat that outcome as failed abstract repair even when the rendered page looks close.
- The stable repair path is:
  - extract a real abstract-title instance, abstract-body instance, and keyword-line instance from the approved template or accepted sample
  - rebuild the Chinese abstract block and the English abstract block as dedicated front-matter surfaces
  - only then apply narrow direct-format exceptions when needed
- Do not accept an abstract block where the title and keyword lines are still fundamentally `正文` paragraphs with cosmetic overrides.

### FB-LAYOUT-042 (legacy 75). Missing Abstract Baseline Is A Real Blocker, Not A License To Guess (Mandatory)

- If the active local template does not actually contain usable abstract formatting instances, do not silently reuse generic body formatting as a fallback.
- Record that the abstract baseline is missing, then lock one approved source of truth before continuing:
  - a teacher-provided full thesis sample
  - the current accepted manuscript's abstract block if the user confirms it is the baseline
  - another explicitly approved local sample
- Until one of those is locked, treat abstract formatting as unresolved rather than "good enough".

### FB-LAYOUT-043 (legacy 75A). Abstract Repair Lanes Must Stay Isolated From Body, Citation, And TOC Mutation (Mandatory)

- Repeated failures show that abstract corruption is usually not a pure wording problem; it happens when the abstract/front-matter block is repaired through a generic body-edit or generic recovery path.
- When a pass touches any abstract surface, lock that pass as an abstract/front-matter lane, not as an ordinary body-text lane.
- In one abstract repair pass, the allowed write scope is only:
  - Chinese abstract title
  - Chinese abstract body
  - Chinese keyword line
  - English abstract title
  - English abstract body
  - English keyword line
  - the immediate rendered-page evidence for those abstract pages
- In that same pass, do not run or mix in:
  - generic body helpers
  - broad `word/document.xml` body replacement
  - citation normalizers or citation finalizers
  - bibliography renumbering
  - TOC refresh or TOC rebuild
  - field refresh that can disturb front matter or bookmark chains
- If citations, bibliography, TOC, or body paragraphs also need repair, close the abstract lane first, rerender and verify the six abstract surfaces, then open a new isolated pass for the later surface family.
- Do not restore abstract pages from fixed paragraph numbers, fixed DOCX offsets, first-run style guesses, or review-copy rollback heuristics.
- The stable abstract recovery path is to clone six real approved baseline instances and replace only the intended text payload for those six surfaces.

### FB-LAYOUT-044 (legacy 75B). Abstract Text Replacement Must Preserve Paragraph Instances And Keyword Run Boundaries (Mandatory)

- Repeated failures show that abstract edits can look text-correct while still being format-wrong because the pass used whole-paragraph text replacement in WPS automation, `python-docx`, OOXML, or another helper that collapsed the original paragraph instance.
- Treat the following as protected abstract surfaces during content edits, not only during pure format repair:
  - Chinese abstract body paragraph direct indents
  - English abstract body paragraph direct indents
  - Chinese keyword label/body run split
  - English keyword label/body run split
  - approved abstract style binding such as `Body Text` versus `Normal`
- Do not update an abstract body or keyword line by assigning one replacement paragraph text string unless the pass immediately replays the approved abstract baseline instance for that exact surface.
- If `关键词：` / `Key words:` / `Keywords:` collapses into one undifferentiated run, if the keyword content becomes bold together with the label, or if body-paragraph left/right/first-line indents disappear after the text edit, treat the abstract pass as failed even when the wording itself is correct.
- After any abstract text rewrite, rerun six-surface abstract self-check on the exact post-edit DOCX. A pre-edit abstract audit becomes stale immediately once the abstract text changes again.

### FB-LAYOUT-045 (legacy 75C). Format Complaints Must Trigger Full Surface-Face Checks, Not Symptom-Only Checks (Mandatory)

- Repeated user feedback shows a recurring failure pattern: the builder verifies only whether a page still exists, whether text overflows, or whether an obvious hotspot was patched, while the actual surface face still differs from the template.
- When the user reports a format problem anywhere in the thesis, treat the report as a request to verify the full face of the affected template-owned surface family.
- Full face means at minimum:
  - style binding and outline/list state
  - Chinese and Western/English font families, including `ascii`, `hAnsi`, `eastAsia`, and `cs` mappings when present
  - font size, bold, italic, underline, color, highlight, and superscript/subscript
  - alignment, first-line indent, left/right/hanging indent, line spacing, and spacing before/after
  - tabs, dotted leaders, page-break-before, keep-with-next, keep-lines, borders, and shading when that surface uses them
  - run boundaries for labels, citations, fields, hyperlinks, page numbers, and keyword labels
- The same-class audit must cover sibling surfaces, not only the one paragraph named in the complaint:
  - abstract complaint: all six abstract surfaces
  - TOC complaint: TOC title plus every used TOC level
  - heading complaint: the same heading level plus adjacent body paragraphs
  - table complaint: table title, header cells, body cells, borders, continuation title when present, and nearby body paragraphs
  - figure complaint: image-holder paragraph, caption paragraph, figure size/position, and nearby explanatory paragraph
  - header/footer complaint: touched header/footer family plus page-number presentation
  - references or acknowledgement complaint: title paragraph, entry/body paragraphs, indentation, spacing, and page opener
- Do not accept a repair that leaves the same surface family split between template-derived formatting and builder/default formatting.

### FB-LAYOUT-046 (legacy 75D). Builder Defaults Must Never Override Template Font Families (Mandatory)

- Repeated failures show that custom English or Western font choices can enter through local helper defaults, office-application defaults, or manual XML patches.
- For thesis template-following work, no repair path may choose a new English, Western, Chinese, or East Asian font because it is common, visually close, or convenient.
- Required font source order:
  1. explicit current-user rule
  2. active school template paragraph/run instance
  3. teacher-approved sample paragraph/run instance
  4. already accepted manuscript instance for the same surface
- If none of those sources exists, record the font baseline as blocked and do not proceed with a guessed font.
- For mixed Chinese/English paragraphs, copy both the East Asian and Western font mappings from the baseline. A run passes only when the DOCX font attributes and the rendered page both match the locked baseline behavior.
- Any helper script that contains hardcoded surface fonts must be treated as unsafe for thesis formatting unless those values are parameterized from the locked baseline before writing.

### FB-LAYOUT-047 (legacy 87). When The User Explicitly Requires WPS Auto-Generated TOC, Do Not Substitute Manual Or Word-Only TOC Workflows (Mandatory)

- If the user explicitly says the TOC must be generated in WPS first and styled afterward, treat that as the required workflow rather than as a preference.
- Do not replace that workflow with:
  - hand-written TOC text
  - XML-only pseudo-TOC reconstruction
  - Word-only auto-TOC generation when WPS automation is available
- The accepted order in that case is:
  1. remove stale TOC blocks
  2. generate one real TOC field in WPS
  3. refresh page numbers
  4. style the resulting TOC paragraphs

### FB-LAYOUT-048 (legacy 88). Before Rebuilding A TOC, Demote All Front-Matter Paragraphs That Could Be Misread As Headings (Mandatory)

- Before any automatic TOC generation, explicitly demote front-matter paragraphs such as:
  - cover title
  - Chinese abstract title/body/keywords
  - English abstract title/body/keywords
  - TOC title
- Do not leave those paragraphs carrying heading styles, outline levels, or hidden style residues that can pollute the regenerated TOC.
- If the rebuilt TOC unexpectedly captures front-matter content, treat that as a front-matter demotion failure first, not as a TOC styling issue.

### FB-LAYOUT-049 (legacy 89). Auto-Generated TOC Styling Must Follow The TOC Styles, Not The Heading Styles (Mandatory)

- After a real TOC field is generated, style its visible entries through the TOC paragraph styles such as `TOC 1/2/3` or `目录 1/2/3`.
- Do not restyle TOC paragraphs by copying the chapter-heading styles into the TOC block.
- If the visible TOC still looks like chapter headings in miniature rather than a TOC, the repair is incomplete even if the entry text and page numbers are correct.
- Do not assume a built-in WPS/Word TOC refresh will preserve the current TOC appearance.
- Before refresh, lock one real local paragraph instance for the TOC title and each used TOC level.
- After refresh, restore the refreshed TOC title and TOC level paragraphs to that locked local baseline in the same pass.
- If the TOC text is correct but the title, indentation, dotted leaders, page-number alignment, or level-by-level font rhythm drift away from the approved sample after refresh, treat the TOC pass as failed.
- If the approved sample implements part of the TOC look through direct paragraph formatting instead of reusable TOC style definitions, restore those direct paragraph metrics explicitly rather than assuming the style name alone is sufficient.
- TOC visible-run restoration must preserve the donor's direct-run shape, including inherited/theme-only runs. Do not write builder-chosen `宋体`/`Times New Roman`,字号, or bold flags into TOC entry/page-number runs when the approved donor expects inherited fonts, theme fonts, an explicit `w:b val=0` only, or an empty direct `w:rPr`.

### Wrong-Draft TOC Baselines Must Be Rejected Explicitly (Mandatory)

- Repeated regressions show a distinct failure mode: the builder refreshes a TOC correctly, then restores formatting from the current draft's already-wrong TOC instead of from the approved template/sample.
- Treat this as a release-blocking workflow error.
- Required baseline source order for TOC restoration:
  - active school template
  - teacher-approved local sample
  - already accepted manuscript copy
  - current working draft only after visual proof of alignment
- If the user has just complained that the TOC still does not match the template, the current working draft is automatically disqualified as the TOC style baseline for that run.

### FB-LAYOUT-050 (legacy 90). Heading-Indent Repair Must Clear Both Standard And Character-Unit Indents (Mandatory)

- If a main-body heading still appears visually indented after paragraph formatting has been reset, assume there may be character-unit indent residues in addition to standard first-line indent.
- Required heading-indent cleanup checks:
  - `firstLineIndent`
  - `leftIndent`
  - `CharacterUnitFirstLineIndent`
  - `CharacterUnitLeftIndent`
- A heading task is not complete until all four are verified clear when the template expects flush-left headings.
- DOCX-level acceptance must inspect both paragraph-level `w:pPr/w:ind` and the effective paragraph-style chain in `word/styles.xml`.
- The hard-fail residue set includes `w:left`, `w:right`, `w:leftChars`, `w:rightChars`, `w:firstLine`, `w:firstLineChars`, `w:hanging`, `w:hangingChars`, and `w:numPr` on real body headings.
- A heading style that inherits body `Normal` first-line indentation because it lacks an explicit zero `w:ind` is still a failed heading repair even when the visible paragraph has the correct `pStyle` name.
- The gate must cover body level 1, level 2, level 3, and level 4 heading instances separately, excluding TOC field results and front-matter titles from the body-heading evidence.

### FB-LAYOUT-051 (legacy 91). Fourth-Level Heading Paragraph Instances Must Be Repaired Explicitly (Mandatory)

- A fourth-level heading is not accepted merely because the style name says `论文4级标题`.
- If the real paragraph instance still keeps body-text first-line indent, body-text justification, or former list-item rhythm, treat the repair as failed.
- Required repair state for project body fourth-level headings:
  - standalone heading paragraph
  - left alignment unless the approved sample explicitly centers it
  - zero first-line indent
  - zero left indent and zero hanging indent
  - keep-with-next enabled

### FB-LAYOUT-052 (legacy 92). When Large Blank Space Appears Before A Figure, Prefer Shrinking The Figure And Binding The Figure-Caption Block More Tightly (Mandatory)

- If a thesis page shows a large blank area while the next page begins with a figure, do not accept that layout by default.
- First try reducing the figure size and tightening the image paragraph and caption paragraph spacing so the figure block can fit on the previous page cleanly.
- Treat the image paragraph and caption paragraph as one local pagination block during this repair.
- Do not leave abnormal visual gaps between the image and its caption, and do not preserve wasteful page-bottom whitespace when a slightly smaller figure would solve the problem.
- For implementation chapters with many runtime screenshots, also check whether real empty body paragraphs or leftover separator paragraphs around subsection boundaries are helping create the visible white gap.
- Stable repair order for this failure class:
  1. remove real empty paragraphs in the touched local chapter range
  2. rerender and recheck page occupancy
  3. if wasteful white space still remains, shrink the affected screenshot block moderately
  4. rerender again and confirm the page looks denser without making the screenshot unreadable
- Do not jump straight to large image-size changes before ruling out removable blank paragraphs in the same local layout zone.

### FB-LAYOUT-053 (legacy 93). Body Sentences Must Never Reappear As TOC Or Heading Entries After Content Or Citation Repair (Mandatory)

- If a late-stage content pass or citation repair causes a full body sentence to appear in the document outline, TOC, or heading map, treat that as paragraph-style contamination, not as a TOC-styling defect.
- Do not repair that incident by styling the TOC only.
- Required recovery path:
  - identify the polluted body paragraph
  - demote it back to the approved body-paragraph class in the source manuscript or thesis generator
  - regenerate the document
  - only then refresh or rebuild the TOC
- If the outline still contains any full claim sentence instead of only real headings, the manuscript is not handoff-ready.

### FB-LAYOUT-054 (legacy 94). TOC Placeholder Blocks Must Be Rebuilt Cleanly When Template Residue Persists (Mandatory)

- If a teacher sample or prior review copy leaves stale TOC field residue, placeholder paragraphs, or hidden TOC content inside the visible TOC zone, do not write new TOC text directly on top of those old paragraphs.
- Required recovery path:
  1. remove the old TOC block from the visible TOC zone
  2. rebuild a clean TOC paragraph block
  3. apply TOC styles to the rebuilt paragraphs
- Writing line by line into stale TOC placeholders is not acceptable when that residue can concatenate page numbers, heading text, or later body text into the same visible paragraph.

### FB-LAYOUT-055 (legacy 95). Inserted Thesis Content Must Clone Local Paragraph Formatting Before Text Replacement (Mandatory)

- When adding new TOC entries, headings, captions, or body paragraphs into an existing thesis DOCX, do not treat text insertion and format repair as separate optional steps.
- The default repair path is:
  1. find the nearest accepted local paragraph instance of the same class
  2. clone or reproduce that paragraph's real formatting and style binding
  3. replace only the text payload
- This applies separately to:
  - TOC level 1/2/3 entries
  - chapter headings
  - subsection headings
  - body paragraphs
  - figure captions and table titles
- A run fails acceptance if the inserted text is semantically correct but still keeps plain default formatting while the surrounding document uses a different local style.
- At minimum, synchronize style binding, font family, font size, bold/italic state, alignment, first-line indent, spacing, line spacing, and keep-with-next/keep-lines behavior with the nearest accepted sample paragraph of the same class.

### FB-LAYOUT-056 (legacy 96). Body Font Baseline Must Come From A Real Local Sample Paragraph, Not A Builder Default (Mandatory)

- For thesis generation or thesis format repair, the final body font and body size must be extracted from a real accepted local sample paragraph instance before the builder applies global body formatting.
- Do not let hardcoded paragraph classes inside a local builder script override that extracted sample baseline.
- If the accepted local sample or template body text is `楷体 12pt`, the builder must not silently rewrite the thesis body into `宋体 12pt` just because the script's paragraph-class table says so.
- A formatting run fails if the body font differs from the approved sample while the builder never recorded a sample-derived reason for that change.
- When inserted or rewritten Chinese body paragraphs contain inline English identifiers, API paths, function names, or code-like tokens, do not rely on DOM-only font labels to judge correctness.
- Required check for that paragraph class:
  - verify the real body-font baseline against a local accepted paragraph instance
  - verify the run-level font settings do not leave Chinese body text effectively rendering as `Times New Roman`
  - verify the rendered page does not show abnormal stretched spacing that makes the paragraph look like a broken full-justified block
- If a newly inserted Chinese body paragraph renders with abnormal spreading under `justify`, treat that as a baseline failure in font and/or alignment, not as a harmless renderer quirk.

### FB-LAYOUT-057 (legacy 97). Template Pagination Must Not Be Rebuilt From A Fixed Section Skeleton (Mandatory)

- Do not assume every thesis should be rebuilt as one fixed sequence such as cover, declaration, authorization, Chinese abstract, English abstract, TOC, and body with one hardcoded section break after each block.
- Section boundaries, page-break ownership, front-matter zones, and numbering restarts must be derived from the active local template or approved sample, not from a reusable builder skeleton alone.
- If the builder uses a fixed section topology or fixed section-count assumptions, verify that topology against the actual local template before writing the final manuscript.
- A pagination pass fails if the section topology was imposed by script habit rather than confirmed from the active template.

### FB-LAYOUT-058 (legacy 98). Page-Break Ownership Must Be Singular In Template-Driven Thesis Builds (Mandatory)

- Do not assign chapter pagination responsibility to both explicit inserted page breaks and template heading/page-break-before behavior at the same time.
- If chapter starts are already owned by verified heading or section formatting from the template, do not add another forced page break before those chapters.
- If the builder adds explicit page breaks for every chapter, references block, or acknowledgement block, verify first that the template does not already create those breaks.
- A run fails pagination review if duplicated break ownership creates blank pages, near-empty pages, or large unexplained white areas before chapters or terminal blocks.
- If the user explicitly asks to add fixed page breaks only for first-level chapters, restrict the repair scope to the confirmed chapter-level heading instances and do not promote second-level subsections into new-page starts as a side effect.
- Stable default for this request class:
  - keep Chapter 1 on the existing first body page boundary unless the template explicitly requires another break
  - apply explicit page-break ownership only to later first-level chapter headings such as Chapters 2-6
  - rerender and verify that every first-level chapter starts on a fresh page while second-level headings remain governed by normal local flow unless separately requested

### FB-LAYOUT-059 (legacy 99). Footer And Page-Number Presentation Must Be Locked As Its Own Baseline (Mandatory)

- Do not infer footer correctness from header correctness, from the mere presence of page numbers, or from a generic page-number API call alone.
- Extract a real accepted sample instance for each footer family that matters in the current run, including:
  - front matter when it differs from the body
  - main body footer or body page-number presentation
  - tail-block footer when conclusion, references, acknowledgement, or appendix use another presentation
- Lock and verify at least:
  - alignment
  - horizontal position on the rendered page
  - font family and size
  - visible numbering style and restart behavior
  - whether the footer is intentionally blank or visible on each section family
  - odd/even or first-page differences when the sample shows them
- Also verify that footer/page-number paragraphs do not carry abnormal first-line indent, left indent, or right indent residues that visibly push the footer away from the approved sample position.
- For legacy `.doc` templates or fragile converted samples, use rendered pages to confirm footer placement and appearance instead of trusting style names or section XML alone.
- A footer repair fails acceptance if the page number exists but the visible footer presentation still does not match the approved template or sample.

### FB-LAYOUT-060 (legacy 100). Tail-Block Titles Must Clone The Real Template Title Baseline, Not Rebuild Font Names By Hand (Mandatory)

- Tail-block title indentation is part of the cloned baseline. A title pass must compare left indent, right indent, first-line or hanging indent, alignment, tab behavior, and centerline against the approved paragraph instance.
- Tail-block title indentation must be recorded as a named evidence surface when references, acknowledgement, conclusion, appendix, or similar end-matter openers are touched.
- Evidence must compare the template and target rendered left-x geometry for tail-block titles and nearby body lines. Font, size, and visible-title checks alone cannot prove that the indentation is correct.

- For `结论`, `参考文献`, `致谢`, `附录`, and similar tail-block titles, prefer copying the real approved title paragraph instance from the template or accepted sample and then replacing only the visible title text.
- Do not synthesize those title paragraphs by manually rebuilding East Asian font names, size, centering, and spacing when a real baseline already exists.
- If a script-generated tail-block title drifts into unreadable font labels, fallback-font drift, or mojibake inside the office app, reject that repair path and restore the title from the real baseline.

### FB-LAYOUT-061 (legacy 100A). TOC Refresh Must Be Followed By Explicit Format Restoration (Mandatory)

- Repeated user feedback shows a distinct failure pattern: after a built-in TOC update, the TOC content refreshes but the visible TOC formatting drifts and is not restored.
- Treat this as its own failure class rather than as a generic TOC-styling issue.
- Required workflow for any run that refreshes the TOC:
  1. extract one accepted local paragraph instance for the TOC title and each used TOC level
  2. run the built-in TOC refresh
  3. restore the refreshed TOC paragraphs to the extracted local formatting baseline
  4. render the TOC page and verify the restored format visually
- Do not hand off a thesis where the TOC was refreshed but still remains in default application styling.
- Do not transplant the entire TOC block from a renderer-refreshed review copy into the official manuscript as a shortcut for TOC repair.
- If the refreshed TOC in a review copy falls back to generic `TOC1/TOC2/TOC3` styling or changes front-matter page display from forms like `摘要I` to `摘要1`, treat that review copy as a content-only refresh artifact and not as the final style source.
- The final official manuscript must restore the TOC title and TOC levels to the approved local baseline before handoff, even when the TOC text and page numbers already look correct.

### FB-LAYOUT-062 (legacy 100B). Tail-Block First Pages Must Restore A Verified Opener Owner (Mandatory)

- Repeated regressions show that `结论`, `参考文献`, `致谢`, and similar tail-block titles can survive while the real fresh-page owner for the opener is lost.
- Treat this as a dedicated failure class rather than as a generic chapter-start or heading-style issue.
- When the school format requires the sequence `参考文献` before `致谢`, the DOCX body order, TOC order, and rendered opener pages must all follow that order. A manuscript with `致谢` before `参考文献`, or with both titles on the same rendered page, is not acceptable.
- Tail-block title repair must set a single opener owner, such as paragraph-owned `w:pageBreakBefore`, and remove redundant hard page-break runs on the title. A title that relies on leftover inline hard breaks is not a verified opener owner.
- The formal `references` opener must be rendered after the previous real content block, not only after a generic tail-block marker. Evidence must record the previous content physical page, the `references` physical page, and `references_prior_block_separation_verdict=pass`; the same page is a failure even if the title paragraph still has a page-start owner in XML.
- Required workflow for each touched tail block:
  1. lock the approved title baseline paragraph instance
  2. lock one page-start owner only for that opener
  3. verify the page immediately before the opener, the previous real content page, and the opener page on rendered output
  4. verify that the DOCX or office-application state still shows that intended opener owner rather than a one-off rendered coincidence
- If a tail-block title shares a page with the prior block, depends on duplicate owners, or no longer has a verified opener owner, the repair fails.

Header-Aware Body/Tail Split Detail

- If the approved sample shows school-name headers on even pages and current chapter-title headers on odd pages, do not leave the body header saved as fixed text like `第一章 绪论` after Chapter 1.
- Required repair path for this template class:
  - inspect `headerReference` bindings and section boundaries instead of editing visible header text only
  - make the main-body odd-page header follow the real first-level chapter heading, for example with a heading-driven header field such as `STYLEREF`
  - create separate end-matter sections when the sample expects dedicated visible headers such as `结论`, `参考文献`, or `致谢`
- If the body header is dynamic but the tail blocks still inherit the numbered chapter title, the repair is incomplete.
- If the tail blocks are fixed correctly but later numbered chapters still show the first chapter title, the repair is incomplete.

Rendered Header/Footer Acceptance Detail

- For thesis header or footer repair, do not hand off the DOCX after checking only header XML, section bindings, outline data, or field codes.
- Required acceptance path:
  - export the current review DOCX to a rendered format such as PDF
  - inspect the touched page tops or full touched pages through machine vision
  - confirm the visible header/footer text matches the sample on the rendered page
- A repair fails acceptance if a field-based header looks structurally correct in XML but still renders as `错误！未定义样式。`, stale text, blank text, or another visible header error on the page.

### FB-LAYOUT-065. Regenerated Thesis Covers Must Clone Accepted Local Donors Instead Of Rebuilding Field Rows By Hand (Mandatory)

- Repeated cover regressions show that cover pages are especially fragile when generation scripts rewrite title and field rows with generic paragraph formatting.
- If a project already has an accepted cover donor or last-known-good manuscript, treat that donor as the owner of:
  - title underline behavior
  - spacer paragraphs between title and field block
  - field-row indentation and alignment
  - visible field underlines
  - cover page break and section-break ownership
  - school logo/image relationship geometry
- Do not rebuild cover fields by applying the title paragraph style or a generic cover style to every row.
- If the visible values need to change, replace only the field text inside the accepted donor structure and preserve the donor paragraph/run properties.
- The visible underline geometry is part of the donor structure. Do not let the underline length expand or shrink with the replacement text; every value line must keep the locked donor span, line weight, and baseline position.
- If the donor uses the same underline geometry across multiple cover rows, preserve that equality. A cover row whose underline becomes shorter or longer than the other rows because the value text changed is a failed repair.
- After regeneration, export the exact final DOCX to rendered pages and inspect at least the cover page plus the next front-matter page.
- A cover repair fails if the first page looks visually fixed but the generator still contains the same hand-built cover logic that will damage the cover on the next rebuild.

### FB-LAYOUT-066. Front-Matter, TOC, And Abstract Rebuilds Must Preserve Page-Class Donors (Mandatory)

- Repeated regressions show that a thesis can look locally repaired while the rebuilt front matter has drifted as a page class.
- Any full or canonical thesis rebuild must lock donor evidence for cover, Chinese abstract, English abstract, TOC, and first body page before writing those surfaces.
- The builder must not use a fixed front-matter skeleton, generic section breaks, default TOC styles, or current-draft wrong baselines to rebuild those page classes.
- Required post-build checks:
  - cover title/field rows still match the accepted local donor and do not drift onto the next page
  - Chinese abstract, English abstract, and TOC remain separated according to the active template without abnormal blank pages
  - TOC title and entries keep template color/font/indent/leader/page-number geometry
  - TOC content includes Chinese abstract and English abstract entries when the active template expects them
  - abstract body text is split into real Word paragraphs and does not contain manual `w:br` line breaks from newline-joined source text
- If the rendered front matter fails any of these checks, the script path that produced it must be treated as unsafe until the donor-lock and detector coverage are repaired.

### FB-LAYOUT-067. Current Template Feedback Must Override Old Table And Cover Memories (Mandatory)

- When a current user correction says the cover, table, TOC, header, or body format still differs from the template, treat that correction as a live template-alignment incident, not as a cosmetic preference.
- For cover pages, value fields must be replaced inside the original value area while preserving the donor underline, row baseline, label cell, field-row alignment, spacer rows, and page-break ownership. A cover page that visually lacks the fill-in underline at the value position is still failed even when the text is correct.
- Cover underline geometry must be verified as a rendered width baseline, not inferred from text length. The repaired cover is invalid if the underline length changes per row, even when the typed content fits neatly.
- For thesis tables, the active donor decides whether the table title is an external paragraph or an in-table first merged title row. Do not move an in-table donor title outside the grid, and do not move an external donor title into a cell.
- Three-line-table evidence must separately record the top rule, header separator rule, any donor-required body middle rule, bottom rule, title mode, width, cell typography, and rendered crop metrics. A table with only a top and bottom rule cannot pass when the donor shows a header separator or a middle rule.
- If a detector, fallback memory, or script assumes `caption inside table = failure`, it must first exempt donor-backed `first_merged_row` table titles. Otherwise the detector is unsafe for templates that place table titles inside the table.
- The final user-reported issue ledger must carry one row per live complaint, including cover underline/alignment, table title mode, missing table rules, header/TOC drift, body style drift, citation superscripts, abnormal blank pages, and AI-sounding prose when those surfaces were reported.

### FB-LAYOUT-068. Official Format Text Must Be Extracted Before 1:1 Template Mutation (Mandatory)

- When the user provides both a school format-requirement document and a format template, do not rely on visual sampling, old memory, or the current damaged manuscript as the format authority.
- Before any whole-thesis, full-paper, format-repair, or `1:1` template-alignment mutation, extract a written requirement profile from the official format document and bind it to the active template profile.
- The extraction must cover at least:
  - cover/title page title fields, identity fields, underline or table-field geometry, and school logo/banner behavior
  - Chinese abstract title/body/keyword line and English abstract title/body/key-words line
  - TOC title, entry hierarchy, leaders, and page-number column
  - body heading levels and body paragraph Chinese/Western font rules
  - odd/even headers, section links, footer, and page-number field presentation
  - citation or annotation marker style, including superscript requirements when the school document says markers are upper-right/superscript
  - reference title, reference entry content standard, numbering/order, font slots, indentation, and line spacing
  - required document order and page-break or section-break ownership for cover, declarations, abstracts, TOC, body, references, acknowledgement, and appendix
- If the official format document says, for example, Chinese body text is `宋体 小四` and Western body text is `Times New Roman 小四`, that textual rule overrides renderer convenience and older fallback memory.
- If the official format document says the body odd-page header is the thesis title and the even-page header is the school/year phrase, the repair must verify odd/even header behavior by section and rendered page rather than editing a single visible header string.
- A run that cannot extract or read the format-requirement document must switch that surface to blocked or audit-only. It may not claim `1:1` alignment from page screenshots alone.

### FB-LAYOUT-069. Template Fonts Are Authoritative; Renderer-Safe Substitution Is Forbidden (Mandatory)

- Do not replace a template-required font with a renderer-safe substitute such as `NSimSun`, `SimSun-ExtB`, `DengXian`, `Microsoft YaHei`, or another nearby font merely because LibreOffice, a PDF renderer, or another conversion path has missing glyphs or poor Chinese rendering.
- If the active school template or official rule requires `宋体`, `黑体`, `楷体`, or `Times New Roman`, preserve that requirement in the DOCX font slots and use a renderer that can inspect it correctly.
- If all available renderers misrender a required school font, record a renderer limitation and use Word/WPS UI evidence or DOCX effective-font-chain evidence. Do not silently change the thesis font and still describe the result as template-aligned.
- A body-font repair must prove the expected East Asian, ASCII, hAnsi, and complex-script font ownership from the template or official text. A pass based on visual similarity, style name, theme alias, or substitute-font PDF appearance is invalid.
- If the school template exports a style font value as a semicolon alias list such as `宋体;SimSun` or `黑体;SimHei`, treat that list as a reference-only fallback set. The final DOCX must keep one concrete font family in each OOXML font slot; do not write semicolon alias lists into the manuscript to satisfy baseline comparison.
- This rule applies to body text, cover fields, abstract labels/content, TOC entries, captions, table cells, headers, footers, page numbers, citation markers, references, acknowledgement, and appendix text.

### FB-LAYOUT-070. Cover, Header, Footer, Abstract, And Page-Number Surfaces Must Be Protected As Donor Families (Mandatory)

- Treat cover, declaration/title front matter, Chinese abstract, English abstract, TOC, header, footer, page-number fields, references, and acknowledgement as protected donor families before any body or style normalization pass.
- Cover repair must clone the accepted local donor structure and replace only variable text. It must not rebuild the cover from generic paragraphs, flatten cover tables, lose title wrapping, remove school text/logo/banner objects, or damage declaration blocks.
- When the active donor is Attachment 8-style cover material, the Attachment 8 cover skeleton is a protected donor surface: top template header lines, logo paragraph position, degree-title position, `题目：` paragraph position, identity-field row order, date row position, and the cover section boundary before declaration/front-matter page must be checked against the donor. A gate that checks only media presence, placeholder removal, or field text is a false pass.
- Clean final covers may remove explanatory callout prose only when the remaining skeleton keeps the donor paragraph order and vertical rhythm; removing the callouts must not collapse the logo, title, topic, identity rows, or date upward.
- Abstract repair must keep the official label/content split. Chinese `摘要`/`关键词` and English `Abstract`/`Key words` surfaces must not inherit body normalization, heading normalization, or bibliography formatting.
- Header/footer repair must verify section-level bindings, odd/even or first-page differences, link-to-previous state, visible header/footer text, and page-number field behavior across front matter, TOC, body, references, acknowledgement, and appendix when present.
- Header repair must compare the template's full visible header string, not only the semantic chapter title. If the template's right header displays a chapter-number component plus title, such as `第一章 绪论`, `第二章 照明系统设计`, or an equivalent numbered pattern, a target header that shows only `绪论` or only `照明系统设计` is a hard failure.
- Header evidence must record the expected display-string source, chapter-number component, chapter-title component, observed rendered string, section/header part source, and verdict for the first page of each chapter plus at least one later body page and every present tail-block family.
- Cover, header, and front-matter horizontal rules must be template-proven before they are kept. A target-only paragraph border, table border, header bottom border, shape line, or imported donor residue must be removed when the active template or approved donor does not show the same rendered line in the same page region. Header-line evidence must identify the concrete OOXML source, such as `w:pBdr/w:bottom`, and compare template versus target values instead of relying on the line looking formal.
- A cover repair that only replays visible cover text is incomplete when the cover remains in the same section as abstracts, TOC, or body pages. The canonical cover repair must create a cover-only section boundary from the approved template/page-setup donor and strip cover-section `headerReference`, `footerReference`, and `pgNumType` unless the locked template explicitly proves a visible cover header/footer/page number.
- Page-number repair must preserve a Word/WPS page-number field when the template uses one. Hand-typed page numbers or page numbers that only look correct before field refresh are failed.
- Whole-document blank-page review must classify whether a blank or near-empty page is template-owned, odd/even-section-owned, or abnormal. Abnormal blank pages caused by duplicate page-break owners, stale section breaks, empty paragraphs, field refresh, or oversized blocks must be fixed before handoff.
- A repair that changes body text or body style without freezing and diffing these donor families is a style-blast-radius repair and cannot pass as a local body-font fix.

### FB-LAYOUT-072. Inserted Or Replaced Body Prose Must Not Inherit Heading, Caption, Or Title Formatting (Mandatory)

- When adding, expanding, rewriting, cleaning, or replacing thesis body paragraphs, clone or preserve the nearest valid body paragraph's paragraph properties and run properties unless the active template or teacher instruction explicitly requires another body donor.
- This rule applies even when a cleanup is described as text-only and does not lengthen the paragraph. A whole-paragraph text replacement that leaves the paragraph visually different from neighboring body prose is a failed repair.
- Inserted or replaced body prose must not inherit chapter-title, heading, TOC, caption, cover, abstract-title, reference-title, acknowledgement-title, figure-caption, table-caption, or title-page formatting.
- Treat these as hard failures: body prose rendered bold/oversized like a heading; body prose centered like a title or caption; body prose assigned an outline level or heading/caption style; body prose appears in the TOC; body prose keeps wrong line spacing, first-line indent, paragraph alignment, font size, or run formatting compared with the local body donor; body prose starts a page as a title-like orphan immediately before a chapter opener; body prose changes the next chapter opener's visual hierarchy.
- After a formal figure caption, table title, figure holder, or table object, the next explanatory paragraph is body prose unless it is itself a formal caption/title. It must be reset to the body donor family and must not keep `Caption`, table-title, heading/title, `keepNext`, center alignment, caption line spacing, zero first-line indent, or caption/title direct run formatting.
- Any image, screenshot, chart, or figure replacement that touches a figure holder or caption must reacquire the figure block from the verified caption after insertion, then inspect nearby explanatory body paragraphs. Newly inserted body prose after the caption must be cloned from a local body donor, not from the caption paragraph, image-holder paragraph, table title, or the mutating COM paragraph chain.
- The first explanatory paragraph after a formal figure/table caption must not begin by repeating a visible figure/table number such as `图5-1...`, `图 5-1 ...`, or `表4-1...`; rewrite it as body prose such as `该图...`, `该表...`, `从图中可以看出...`, or `上述结果...`.
- Caption detection must treat lead-ins such as `图4-7显示...`, `图4-7所示...`, `图4-7说明...`, `图4-7表明...`, and similar `figure/table number + narrative verb` sentences as explanatory body prose, not as additional formal captions. These paragraphs must be body-aligned and audited for caption-adjacent contamination even when they start with a formal-looking figure/table label.
- The same prohibition applies to nearby explanatory body prose in the same figure/table block, not only the first paragraph after the caption. If a later nearby paragraph starts with a figure/table-number cluster such as `图3-4和图3-5...` or `该图和图3-5...`, rewrite it to normal body wording such as `该组图...` or `上述结果...` and rebind it to the body donor family.
- Caption/table sibling body contamination is a named blocker: when a table/figure-adjacent explanatory paragraph keeps caption/table-title formatting, the final record must expose a `caption/table sibling body contamination verdict` backed by the body-style audit and rendered touched-page evidence.
- If a nearby explanatory paragraph does not start with a figure/table-number cluster but carries caption/title metrics, such as centered alignment, `keepNext`, zero first-line indent, caption line spacing, title-like font size, or a caption/title style id, it is still caption contamination and must be repaired before any figure task can pass.
- If a rewrite or cleanup touches an existing paragraph, compare the final paragraph against neighboring same-family body paragraphs and repair both paragraph properties and character run properties before handoff. Do not accept a paragraph whose text is correct but whose visible Word/WPS style differs from surrounding body text.
- Acceptance must compare inserted or replaced body paragraphs against a local body donor using both DOCX paragraph/run metrics and rendered page images. A pass based only on style names, XML counts, package counts, or PDF existence is invalid, and a user screenshot or report of visible body-style mismatch overrides XML-only evidence.
- If a content expansion, cleanup, or style normalization changes pagination or appears near a figure/table/caption, the format lane must review the touched pages, the previous/next rendered pages, and any chapter opener moved by the change before handoff.

### FB-LAYOUT-073. Thesis Deliverables Must Not Retain Blue Or Theme-Colored Visible Text Unless The Active Template Requires It (Mandatory)

- A thesis DOCX/PDF handoff must not leave visible heading, title, caption, citation, TOC, body, reference, acknowledgement, header, footer, or page-number text in blue or another theme accent color unless the locked active template or official school requirement explicitly requires that exact color.
- Do not treat PDF export, reference count, image count, or successful Office validation as proof of font-color correctness. Word built-in styles such as `Heading 1`, `Heading 2`, `Heading 3`, `Title`, `Caption`, and `Hyperlink` can keep blue theme colors in `styles.xml` even when direct body runs look black.
- Before final thesis handoff after generation, whole-thesis revision, format repair, figure insertion, citation repair, or style normalization, run the canonical font-color audit on the exact final DOCX. The audit must inspect both direct run colors and the colors of styles actually used in the manuscript.
- If the audit finds a used style or direct run with non-black visible color, repair that style/run to black, clear theme-color attributes, rerender the PDF, and rerun the audit on the exact final DOCX. A handoff with known non-black visible text is a failed format repair unless the active template evidence proves that color is required.

### FB-LAYOUT-074. Protected Thesis Surfaces Must Not Carry Abnormal Bullet Or List State (Mandatory)

- A thesis DOCX/PDF handoff must not leave abnormal project-symbol bullets, Word list bullets, Unicode bullet prefixes, or other visible bullet prefixes on protected surfaces such as cover/front matter, abstract titles and bodies, TOC title and entries, body headings, references title, bibliography entries, appendix, or acknowledgement.
- Do not treat the absence of a visible bullet character in `word/document.xml` as enough. Inspect both direct paragraph `w:numPr` and style-inherited `w:numPr`, then resolve the referenced `numbering.xml` model to ensure it is not `numFmt=bullet` or a bullet-like `lvlText`.
- If headings use manually written numbers such as `第1章`, `1.1`, or `1.1.1`, their heading styles must not also carry automatic bullet/list numbering. If bibliography entries use visible manual numbers such as `1.` or `[1]`, the same paragraph must not also carry Word list state unless the active template explicitly proves that exact automatic numbering model.
- TOC field caches and PDF exports must be checked after repair. A DOCX that looks clean in XML but renders `•`, `U+F0B7`, or similar project-symbol bullets in Word/WPS/PDF is still a failed format repair.
- Before final thesis handoff after generation, whole-thesis revision, format repair, TOC refresh, citation/reference repair, or style normalization, run `scripts/audit_docx_list_pollution.py` on the exact final DOCX and include its verdict in the whole-format gate evidence. Any failure blocks handoff.
