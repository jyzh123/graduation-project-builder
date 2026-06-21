# Protected Surface Evidence And Agent Audit Contract

Use this file as the common evidence contract for thesis-format generation, revision, repair, and read-only audit.
Surface-specific files still own local repair rules; this file owns the record shape, proof standard, and agent-audit handoff contract.

## Enforcement Status

- Every rule in this file is mandatory when thesis formatting, thesis generation, thesis revision, format repair, or full-paper format audit is in scope.
- This file must be loaded with `references/thesis/thesis-format-rules.md`, `references/thesis/thesis-format-class-review.md`, and `references/agents/agent-lanes.md` for any substantial thesis-format run.
- If this contract conflicts with a surface-specific rule file, use this contract for evidence shape and audit fields, and use the surface-specific file for local formatting details.
- A run may not claim template alignment, whole-thesis format pass, or abstract/TOC/reference pass while any required field below is absent, stale, generic, or contradicted by rendered evidence.

## FMT-EVID-001. Protected Surface Ids Are Canonical

The following ids are canonical for protected high-risk format surfaces:

- `cover_style`
- `declaration_or_title_front_matter`
- `zh_abstract_title`
- `zh_abstract_body`
- `zh_keyword_line`
- `en_abstract_title`
- `en_abstract_body`
- `en_keyword_line`
- `toc_title`
- `toc_entries`
- `toc_dotted_leaders`
- `toc_page_number_column`
- `body_heading_levels`
- `body_text`
- `figure_table_captions_and_holders`
- `body_citation_superscripts`
- `review_comments_and_change_marks`
- `references_title`
- `references_entries`
- `acknowledgement_title`
- `acknowledgement_body`
- `appendix_title`
- `appendix_body`
- `header`
- `footer`
- `page_numbers`
- `whole_document_pagination`

The active checklist, format-repair task record, agent task cards, surface inventory, high-risk surface matrix, review evidence records, and final acceptance record must use these ids exactly.
Generic names such as `abstract`, `toc`, `front matter`, `reference block`, `page`, or `font checked` are not valid substitutes.

## FMT-EVID-002. Required Evidence Record Shape

Each present protected surface must have its own review evidence record.
One evidence record must not be reused to prove multiple protected surfaces.

Each evidence record must include:

- target surface id and exact target identifier
- baseline surface id using the same canonical id
- reviewed output path and SHA256
- baseline source path and SHA256
- target DOCX paragraph/run path
- baseline DOCX paragraph/run path
- paragraph metrics: style binding, outline/list state, alignment, indentation, line spacing, spacing before/after, tabs/leaders, keep rules, page-break ownership
- run metrics: font slots, size, bold, italic, underline, color, highlight, superscript/subscript, field/bookmark/hyperlink behavior when applicable
- for `body_heading_levels`: level 1/2/3/4 baseline verdicts, direct run typography, effective style/basedOn font chain, `w:sz` and `w:szCs`, explicit bold on/off state, paragraph-dialog metrics, absence of body-format residue, and TOC/chapter-start synchronization
- rendered PDF path and rendered region/page image path
- distinct template and target rendered region image paths
- numeric rendered geometry metrics for both baseline and actual: bounding box, x/y position, width/height, line height or row y-delta, spacing before/after, indentation/tab position when applicable, and page-region occupancy
- rendered surface geometry verdict derived from those numeric metrics
- for whole-document pagination: section boundary map, section property baseline/actual map, page-number restart map, header/footer link-to-previous map, hard page-break and section-break map, field-refresh before/after state, TOC-to-heading page sync map, logical-to-physical page map, rendered page count baseline/actual, blank or near-empty page scan, chapter opener page map, tail-block opener page map, and a whole-document pagination verdict
- for `body_citation_superscripts`: source and final citation marker run inventories, citation paragraph ids, marker text, run index, `w:vertAlign`, font size, color, underline, hyperlink/bookmark host behavior, punctuation-side placement, first-appearance order, controlled-change ledger when markers are intentionally renumbered, and source-to-final superscript retention verdict
- for `review_comments_and_change_marks`: source and final package part inventories for `word/comments.xml`, `word/commentsExtended.xml`, `word/people.xml`, body comment anchors (`w:commentRangeStart`, `w:commentRangeEnd`, `w:commentReference`), tracked-change marks (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`), bookmark/cross-reference anchors, part-level SHA256 values, id/text digest maps, explicit strip-approval evidence when applicable, and source-to-final review-artifact preservation verdict
- metric-by-metric comparison verdict
- final row verdict: `pass`, `fail`, or `needs-manual-review`
- blocker or explicit `none`

Supporting artifacts such as schema validation, PDF export success, full-page screenshot, page-image existence, visible text presence, style-name presence, page-order check, or a generic "looks correct" note cannot be the evidence record. A rendered image also cannot pass alone; the evidence must extract numeric template-vs-target geometry for the surface being judged.

Measurement provenance is mandatory. A protected-surface evidence record must name the independent measurement JSON path, measurement JSON SHA256, measurement schema, measurement generator script and command, reviewed output path/SHA256, template path/SHA256, and measurement provenance verdict. The acceptance-record generator may aggregate evidence but must not be the measurement producer. A pass-shaped Markdown record that cannot be traced to an independent measurement JSON for the same final DOCX SHA256 is blocked.

## FMT-EVID-003. Effective Font Chain Is A Format Gate

Typography is part of thesis format, not a supporting detail.
For every protected surface and every touched surface that contains visible text, font evidence must resolve the effective font chain for `ascii`, `hAnsi`, `eastAsia`, and `cs`.

The chain must record:

- direct run `w:rPr` and `w:rFonts`
- character style and basedOn chain when present
- paragraph style and basedOn chain
- `docDefaults`
- theme major/minor mappings
- WPS/Word UI displayed font names when available
- final resolved value or theme alias for each slot
- template baseline value or alias for each slot
- per-slot comparison verdict

Theme/default aliases such as `minorHAnsi`, `minorEastAsia`, `Calibri (Body)`, or `宋体（正文）` pass only when the locked template baseline for the same surface uses the same effective alias.
An empty direct run font is not a match unless the inherited effective chain is resolved and compared.

## FMT-EVID-004. Agent Lane Integration

The controller must route protected surfaces through the canonical lane roster defined in `references/agents/agent-lanes.md`.

Minimum ownership:

- `总控`: workflow classification, checklist, manifest, dispatch, final merge
- `格式`: template lock, surface inventory, baseline extraction, typography/page/TOC checks
- `内容`: content changes and semantic parity checks
- `图表`: figure/table surfaces
- `引用`: citation markers, hyperlinks, bibliography order and bibliography entries
- `程序`: program evidence when thesis content depends on implementation truth
- `验收`: final record assembly and validation commands
- `审核`: independent or sequential audit verdict

For each lane, the task card must record active/skipped/not-applicable status, protected surface ids owned by the lane, evidence paths produced or reviewed, blockers, and whether the evidence was created for the exact current output.
Inactive lanes must remain in the complete roster with a concrete reason.

The audit lane must reject handoff when:

- a protected surface lacks an owner
- a task card uses generic surface names
- an evidence path is stale or points to a different output
- the final acceptance record omits the run manifest, task cards, or exact output SHA256
- a multi-agent claim lacks authorization source, spawned agent ids or fallback reason, and audit verdict
- a dispatch wave exceeds the live-agent cap or lacks audit attendance

## FMT-EVID-005. Baseline Lock And Mutation Order

Before thesis DOCX mutation, the `格式` lane must lock:

- active template path and SHA256
- template profile path
- template/profile selected before mutation verdict
- mandatory surface inventory path
- high-risk surface matrix path
- page-class rendered baseline targets
- font-chain resolver evidence plan

Mutation order must remain surface-bounded.
Do not let one helper script own cover, abstract, TOC, body headings, body text, references, headers, footers, and pagination in one pass.
After each mutation cycle, rerender the relevant pages and produce a mutation-level audit verdict before the next unrelated mutation.

Read-only audit runs must still produce the surface inventory, evidence paths, and audit verdict, but must mark mutation fields as `not-applicable-with-reason`.

## FMT-EVID-006. TOC, Page Numbers, Citations, And References Are Protected Format Surfaces

TOC evidence must cover:

- title
- every used entry level
- body-heading coverage map from every real level-1/2/3 body heading to its visible TOC cache row
- dotted leaders and right-tab position
- page-number column
- per-entry visible page number vs rendered heading page
- front-matter numbering system vs main-body numbering system
- rendered visual geometry against the locked template baseline, including:
  - template rendered TOC page or region image, distinct from the target image path
  - target rendered TOC page or region image, distinct from the template image path
  - numeric title bounding box and title-to-first-entry gap for both baseline and actual
  - numeric first-entry and per-row bounding boxes for both baseline and actual
  - numeric per-level left x positions / indentation offsets for both baseline and actual
  - numeric row y-deltas and line-spacing rhythm for both baseline and actual
  - numeric dotted-leader start x, end x, and density for both baseline and actual
  - numeric right-tab / page-number x column for both baseline and actual
  - numeric row count per page and page occupancy rhythm for both baseline and actual
  - metric-by-metric visual-geometry verdict

Citation/reference evidence must cover:

- citation first-appearance order
- one citation marker per sentence when the stricter rule is active
- single-number markers only
- superscript visual style
- exact final DOCX SHA256 and source-to-final citation-marker run inventory when the run rewrites any citation-bearing body paragraph
- internal hyperlink/bookmark integrity when required by the current run
- bibliography entry count retention
- reference title and entry metrics
- mixed Chinese/Western bibliography font slots by run, not by whole paragraph

Page-number evidence must cover rendered page tokens, logical-to-physical page mapping, footer typography, and restart behavior.
Whole-document pagination evidence must cover every section boundary and every page-numbering restart, not only visible page tokens on sampled pages.

TOC text, font-chain, page-number, field/bookmark, and style-name evidence are necessary but not sufficient. A TOC that has correct entries, visible leaders, and page numbers but is compressed into a dense one-page/default-application layout fails unless the rendered geometry also matches the template or an explicit content-growth exception is recorded and approved. Natural-language geometry claims such as `template equals actual`, `matched`, `looks consistent`, or `visual pass` are forbidden unless the same record also carries the numeric baseline/actual measurements listed above.

TOC visual-geometry evidence follows the exact-output binding rule in FMT-EVID-007. After any DOCX, TOC, pagination, section/page-number, body-length, or front-matter mutation, prior TOC visual-geometry evidence is stale until regenerated for the current output.

## FMT-EVID-007. Acceptance Merge Rule

The final acceptance record must name:

- this contract file in `loaded references`
- active template path, SHA256, and profile path
- agent run manifest path
- all role task card paths
- mandatory surface inventory path
- high-risk surface matrix path
- per-protected-surface evidence path and verdict
- font-family baseline audit evidence path
- citation and bibliography audit paths when references are present
- rendered-page evidence paths
- TOC visual-geometry evidence path and verdict when TOC is present or user-reported
- TOC paragraph-and-typography metric evidence path and verdict when TOC is present or user-reported
- TOC implementation family/live-field parity evidence path and verdict when the template contains a TOC field, content control, or editable TOC cache
- header expected display string source, rendered full-display evidence path, and chapter-number preservation verdict when headers are present or user-reported
- whole-document pagination evidence path and verdict for every thesis-format, thesis-revision, thesis-generation, or full-paper audit run
- exact output path and SHA256
- audit verdict and open blockers

The agent run manifest and each relevant task card must carry the same protected-surface evidence map, TOC visual-geometry evidence path/verdict, TOC paragraph-and-typography evidence path/verdict, and whole-document pagination evidence path/verdict. Agent audit is incomplete if these fields exist only in the final acceptance record.

If any required protected surface has `fail`, `blocked`, blank evidence, stale evidence, or `needs-manual-review` without explicit user-approved manual acceptance, the handoff status is not `pass`.

FMT-EVID-007 is the canonical owner for exact-output evidence binding. In final acceptance, protected-surface records, TOC visual-geometry JSON, TOC paragraph-and-typography JSON, and whole-document pagination JSON must all bind to the exact final DOCX path and reviewed output SHA256 named in the final acceptance record. Evidence for a previous review copy, intermediate pass, or same-content different path is stale and must block handoff. During an in-progress mutation cycle, a locked current review copy may be used only as cycle-local evidence; it cannot substitute for final-acceptance evidence unless it is the exact final DOCX path and SHA256.

## FMT-EVID-008. All Template-Owned Surfaces Require Rendered Numeric Geometry

The rendered-geometry rule is not a TOC-only rule.
Every present, touched, user-reported, or template-owned thesis format surface must be judged with numeric rendered template-vs-target geometry before a format pass can be claimed.

This includes cover and declaration rows, Chinese and English abstracts, keyword lines, TOC, heading levels, body text, figure holders, figure/table captions, table titles and cells, citation superscripts, reference title and entries, acknowledgement, appendix, headers, footers, and page numbers.

For each surface, the evidence must compare a distinct rendered template region image against a distinct rendered target region image and record numeric baseline/actual measurements for bounding box, x/y position, width/height, line height or row y-delta, spacing before/after, indentation or tab behavior when applicable, and page-region occupancy.

Font-chain, style-name, XML schema, page-order, screenshot, PDF export, visible text, and TOC-specific checks are supporting signals only.
They do not replace the all-surface numeric rendered template-vs-target geometry record.

## FMT-EVID-009. TOC Level Style Binding Requires Paragraph And Typography Metrics

The existing TOC rule remains active: TOC title and every used TOC level are independent template-owned formatting classes, and each level must have its own locked style/baseline evidence.
This rule strengthens enforcement rather than replacing that style-binding rule.

For protected TOC surfaces (`toc_title`, `toc_entries`, `toc_dotted_leaders`, and `toc_page_number_column`), evidence must record Word/WPS paragraph-dialog and typography metrics for the TOC title and every used TOC level.
The used-level set must be explicit. A TOC with level 2, level 3, or level 4 entries cannot pass from title plus level-1 evidence.
The record must compare template and actual values for:

- style id/name, outline/list state, and whether direct paragraph formatting is part of the approved baseline
- font family, font size, bold/weight, and other visible run-style differences for TOC title and each used TOC level
- visible-run direct typography for TOC title text, each used level's entry text runs, tab/leader runs, and page-number runs, including direct `w:rPr`, `w:rFonts` script slots, theme font slots, size, sizeCs, and weight
- paragraph before/after spacing
- line-spacing mode and numeric value
- left, right, first-line, and hanging indentation in points and/or character units
- right-tab stop position and dotted-leader type
- page-number tab/column behavior when applicable
- scale/compression verdict proving the TOC was not proportionally shrunken, squeezed, or left in default application styling
- used-level inventory and evidence map proving one row or linked evidence record for each level that appears in the target TOC

Style names such as `TOC 1`, `TOC 2`, or `TOC 3` are necessary signals but not enough.
A TOC evidence record fails when it proves only style binding, visual geometry, font chain, page-number correctness, or field/bookmark state while omitting the paragraph-dialog and typography metric comparison.
A TOC evidence record also fails when it collapses visible-run typography into a single paragraph/style font value. If the locked template donor has direct run font properties for entry text, tab/leader, or page-number runs, the target must carry matching direct run properties or an explicit same-surface template-approved inherited baseline; empty target run properties cannot pass by style-name equality alone.

## FMT-EVID-010. All Style-Bound Surfaces Require Paragraph-Dialog And Typography Metrics

The TOC paragraph-and-typography rule is a specialized instance of a broader rule.
Every present, touched, user-reported, or template-owned thesis format surface that has paragraph or run styling must turn style-binding claims into WPS/Word paragraph-dialog and typography evidence before a format pass can be claimed.

This includes cover and declaration rows, Chinese and English abstract title/body/keyword lines, body heading levels, body text, figure holders, figure captions, table titles, table cells, citation superscripts when paragraph-hosted, reference title, reference entries, acknowledgement title/body, appendix title/body, headers, footers, page-number paragraphs, and any school-template front-matter row.

For each such surface, evidence must record template and actual values for:

- style id/name, style type, basedOn chain, and whether direct paragraph formatting is part of the approved baseline
- outline/list state, numbering/list binding, and field/bookmark/hyperlink host behavior when applicable
- paragraph alignment, spacing before/after, line-spacing mode and numeric value
- left, right, first-line, and hanging indentation in points and/or character units
- tab stops, leader type, keep-with-next, keep-lines, widow/orphan, page-break-before, and section/page ownership when applicable
- visible typography: font family, font size, bold/weight, italic, underline, color, highlight, superscript/subscript, and script-specific font slots
- scale/compression verdict proving the surface was not accepted after proportional shrinking, default-application fallback, or dense layout drift

Final acceptance is incomplete if this proof exists only as prose in a rule file, a broad `baseline metrics` field, a rendered screenshot, a style-name check, a font-chain-only audit, or a geometry-only audit.
The evidence record must expose dedicated paragraph-dialog / typography fields, and the gate validator must reject missing, generic, nonnumeric, blocked, or not-checked values.
For figure/table caption surfaces, absence of a formal caption donor in the template or target is not a pass-shaped exception. The hard-field producer must still expose numeric `template` and `actual` paragraph/typography fields using an explicit surface fallback, or the surface fails closed.
For direct DOCX metric validation, `body_text` means real body prose only. TOC rows, page-number rows, figure/table captions, image-holder paragraphs, keyword lines, and front-matter abstract surfaces must be excluded from `body_text`; Chinese and English abstract bodies must be detected both as standalone-title blocks and inline label-plus-content forms such as `摘 要：...` and `Abstract: ...`.
Body-text typography evidence must compare the resolved template/body baseline against every real body-prose run family that was touched or user-reported, including direct `w:sz`/`w:szCs`, paragraph style size, inherited style chain size, and caption/table sibling paragraph metrics. A paragraph whose visible body prose becomes smaller or denser than the body baseline fails even if its style name is `Normal`.
Final acceptance must bind the body-style audit summaries for style binding, Normal baseline, body family consistency, heading contamination, mixed-script font separation, direct visible metrics, and explicit `result: pass`. Missing mixed-script or direct-visible-metrics summaries are not pass-shaped evidence for `body_text`.

Header evidence additionally requires the structured `header.presence-contract` detector record in the exact `sample_self_check` report used for handoff. The detector must prove section-level effective header references, non-empty header parts, template header-token preservation, and rendered-page header-token visibility on body and tail sample pages. A visible header on one page, a DOCX header part existing somewhere, or a broad `header_ok` Boolean is not enough.

When the template header includes a chapter-number component, the header evidence must also expose `expected full display string`, `expected chapter-number component`, `expected chapter-title component`, `observed rendered full display string`, and `chapter-number preservation verdict` for each chapter page sampled. A header that preserves the title but drops the numbering component fails even if `header.presence-contract`, title consistency, section linkage, or field-code checks otherwise look correct.

Whole-format audit evidence must also expose the protected surface defect hard fields produced by `scripts/audit_docx_whole_format_gate.py` under `surface_checks`. A pass-shaped whole-format report is invalid unless all of these named checks exist and pass: `cover_media`, `front_matter_hard_fields`, `header_full_display_string`, `toc_page_number_right_tab`, `references_entries_font_size`, `acknowledgement_title_style`, and `footer_page_number_font_size`. The validator must fail closed when any field is absent, prose-only, stale, not bound to the exact final DOCX SHA256, or contradicts the DOCX package structure. `cover_media` is required only when the locked template or approved sample proves a cover-zone media surface; when the baseline cover is text-only, the report must bind the template SHA256 and record `required=false` rather than inventing cover media.

Protected-surface review evidence records must carry matching hard fields instead of only general `baseline metrics` prose:

- `cover_style`: `cover media/icon requirement baseline`, `cover media/icon relationship ids baseline/actual`, `cover media/icon package targets baseline/actual`, and `cover media/icon binding verdict`
- cover/front-matter surfaces: `front-matter hard-field paragraph metrics baseline/actual`, `front-matter hard-field run typography baseline/actual`, and `front-matter hard-field verdict`
- `header`: `header expected full display string`, `header observed rendered full display string`, and `header full-display string verdict`
- `toc_dotted_leaders` and `toc_page_number_column`: `TOC right-tab stop semantic baseline/actual`, `TOC page-number column right alignment baseline/actual`, `TOC page-number tab leader ownership baseline/actual`, and `TOC per-entry right-tab/page-number verdict`
- `references_entries`: `references entries font-size baseline/actual`, `references entries per-entry font-size map`, and `references entries font-size verdict`
- `acknowledgement_title`: `acknowledgement title style baseline/actual` and `acknowledgement title paragraph style verdict`
- `footer` and `page_numbers`: `footer page-number font-size baseline/actual`, `footer page-number run path map`, and `footer page-number font-size verdict`

These fields are hard fields. A rendered screenshot, visual statement, section count, font-chain-only audit, citation count, TOC page-number text check, or generic end-matter evidence cannot substitute for them.

## FMT-EVID-011. Whole-Document Pagination Requires Section/Page/Field Evidence

Whole-document pagination is a protected format surface, not a side effect of successful PDF rendering.
It must be represented by the canonical id `whole_document_pagination` whenever thesis formatting, thesis generation, thesis revision, format repair, whole-paper review, TOC repair, header/footer repair, body rewrite, chapter-boundary work, figure/table insertion, reference/acknowledgement/appendix repair, or page-number correction is in scope.

Whole-document pagination evidence must record:

- package baseline manifest and package drift report for DOCX parts that affect layout
- pre-mutation and post-mutation page maps for all rendered pages
- section boundary map and section property baseline/actual map
- page-number format and restart map for front matter, body, and tail blocks
- header/footer link-to-previous map
- hard page-break and section-break map
- field-refresh before/after state, including TOC field state when a live TOC is required
- TOC-to-heading page sync map for every visible TOC entry
- logical-to-physical page map and displayed page-token map
- rendered page count baseline/actual
- blank or near-empty page scan verdict
- chapter opener page map, including the page before each opener and the first paragraph block under the opener
- references, acknowledgement, appendix, and other tail-block opener page map
- page-class coverage and occupancy rhythm verdict
- final whole-document pagination verdict

The final acceptance record, format-repair task record, review evidence record, agent run manifest, and relevant task cards must all name the whole-document pagination evidence path and verdict.
The validator must reject pass claims when these fields are absent, stale, generic, derived from sampled pages only, or contradicted by rendered page images.
Hard page-break or section-break drift is not content-growth drift. If a protected boundary page break disappears, moves across front matter / TOC / body / tail-block boundaries, or its count decreases relative to the locked source, the pagination surface must fail even when the manuscript text was expanded.

## FMT-EVID-012. Heading And End-Matter Indentation Evidence Is Mandatory

Heading and end-matter indentation defects are protected-surface defects, not visual cleanup notes.
When any thesis-format run touches broad body styles, paragraph defaults, heading styles, references, acknowledgement, appendix, pagination, TOC refresh, or user-reported indentation/offset issues, the evidence set must include independent rows for:

- `body_heading_levels`
- `references_title`
- `references_entries`
- `acknowledgement_title`
- `acknowledgement_body`
- `appendix_title` when present
- `appendix_body` when present

Each row must prove the same exact output path and SHA256 named in final acceptance.
The row must include paragraph-dialog metrics for template and actual alignment, left/right/first-line/hanging indent, tab stops, line-spacing mode/value, spacing before/after, keep/list/page-break ownership, style id/name, and body-format residue state.
It must also include run typography values for font slots, font size, bold, underline, color, and effective font chain.

Rendered evidence must compare template and target region images with numeric title centerline or left-x, text bounding box, line-height/y-delta, page occupancy, and bibliography hanging/left-x geometry when applicable.
A protected title such as `acknowledgement_title` must fail when template and actual alignment, left/right/first-line/hanging indent, tab stop, centerline, left-x, or x-position metrics differ, even if the evidence record says `pass`.
A generic tail-block screenshot, page-start check, reference-entry-only audit, heading-style-name check, or body-style audit cannot substitute for these surface-bound indentation rows.

## FMT-EVID-013. Format-Preservation Promises Require Chapter-Level Evidence

When a user, teacher comment, checklist, task card, or agent statement promises to preserve thesis formatting, the acceptance evidence must prove chapter-level preservation rather than only surface-level existence.

Required evidence fields:

- format-preservation promise source and trigger wording
- touched chapter ids and heading texts
- chapter format diff path
- touched chapter rendered evidence paths
- chapter opener pagination evidence for the page before the opener, the opener page, and the first paragraph block
- per-touched-chapter heading style/rPr/font/spacing comparison
- per-touched-chapter body paragraph style-family comparison
- direct paragraph-dialog and typography metrics for representative changed and unchanged paragraphs inside the chapter
- non-target protected-surface format preservation verdict
- structured `chapter.format-preservation-contract` detector result from `scripts/sample_self_check.py`

The detector and final acceptance must bind to the exact final DOCX path and SHA256. Evidence for an earlier review copy, a sampled-only page, or a different chapter is stale.

If the detector is missing, not-applicable without a concrete no-chapter/no-promise reason, failed, or has an empty evidence object, the thesis run is blocked. If a named chapter such as `第七章` is reported damaged, the evidence must cover that whole chapter range, not only the screenshot area or the paragraph being rewritten.

A local repair that touches only front-matter protected surfaces such as
`zh_keyword_line` and `en_keyword_line` may carry a concrete
`not-applicable-local-keyword-only` chapter detector verdict, but only when the
transaction targets no body/chapter surface and makes no format-preservation
claim. The local repair must still provide protected-surface freeze, target
render, blast-radius, and cross-surface regression evidence for the touched and
sibling surfaces.

## FMT-EVID-014. Review Artifacts And Citation Runs Require Source-To-Final Diff Evidence

Review comments, tracked-change marks, bookmarks, field anchors, hyperlinks, and body citation superscript runs are preservation surfaces, not optional editor metadata.

Before any DOCX text mutation, the run must create a source inventory for:

- DOCX review package parts: `word/comments.xml`, `word/commentsExtended.xml`, `word/people.xml`, and related relationship/content-type entries
- body comment anchors: `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference`
- tracked changes: `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo`
- bookmarks, hyperlink anchors, field-code hosts, footnote/endnote anchors when present
- body citation marker runs with paragraph id, visible marker text, run index, superscript state, font slots, color, underline, and hyperlink/bookmark host

After mutation, the final evidence must compare that inventory against the exact final DOCX path and SHA256. A pass requires all source review artifacts and citation superscript runs to be preserved, or an explicit source-to-final controlled-change ledger for each intentional change.

For body citation hyperlinks, the comparison is source-relative. A source
citation marker that already lacks a hyperlink/bookmark host must not be
reported as `lost hyperlink host`; the diff must instead record that hyperlink
preservation is not applicable for that marker. A source marker that has an
internal hyperlink or field host and loses it in the final DOCX remains a hard
failure.

The final acceptance text for this rule must say the exact final DOCX path and SHA256, `w:vertAlign=superscript`, and source-to-final controlled-change ledger explicitly, not just imply them through nearby prose.

The following are hard failures:

- stripping `word/comments*.xml`, `word/people.xml`, comment anchors, or tracked-change marks to make the page look clean
- generating only a no-comment copy and calling it the final reviewed deliverable
- replacing a citation-bearing paragraph with one plain text run
- accepting a citation audit that records only the document path and not the exact final DOCX SHA256
- letting the audit lane pass comment-implementation coverage while review artifacts or citation superscripts regressed

If the user explicitly requests a clean copy without comments, create it as a separate preview or no-comments deliverable. The run must still preserve or archive the review-artifact-bearing copy and record the explicit user request in the final acceptance record.
## FMT-EVID-015. External Format Report Ledgers Must Bind Every Report-Owned Surface

External format reports are protected-surface evidence sources. A report-driven repair must produce a normalized ledger with:

- external report source path, extracted report paths, optional comment-bearing DOCX path, report timestamp, template name, official overview error count, expanded issue-row count, comment-anchor issue count, and stats-page issue count
- stable row identifiers for paragraph/report/comment rows, surface id, detector id or issue family, actual value, expected value, target text when available, source report file, and blocking/advisory status
- explicit distinction between official overview totals and expanded cell/paragraph sub-issues
- exact final DOCX path and SHA256 for every final report-equivalent audit
- evidence paths for structure order, TOC paragraph spacing, abstract/keyword spacing and alignment, body heading spacing, table cell line rule/alignment, caption typography, reference label spacing, tail-title typography, header/footer/page-number surfaces, and report-comment cleanup when present
- manual-disposition rows for report reminders that require author judgement, such as abstract word-count reduction

A final handoff is blocked when report-owned rows are left as prose notes, when stats-page structure/header/footer/page-number issues are omitted from the ledger, when a row has no final evidence binding, or when the final audit points to a stale DOCX/PDF hash.

For Fanyu/format-audit packages that include a comment-bearing DOCX, recognizing a comment issue family as supported is not enough. The report-equivalent audit must close each blocking comment/stat family with a concrete verifier row. Page setup families are mandatory hard fields: `页眉内容` must verify the fixed left header text `沈阳科技学院学士学位论文` plus a non-empty right chapter/tail-block title on every header part, and the right title must match the chapter/tail-title set owned by the section that references that header part rather than any arbitrary legal chapter title; `页码` `字号问题` must verify PAGE-field/page-number runs at `小五` (`w:sz=18` and `w:szCs=18`) on every footer part; `空白页问题` must remain failed unless rendered PDF or whole-document pagination evidence proves no blank rendered page. These rows may not be downgraded into ledger summaries.

## FMT-EVID-016. Figure And Table Caption Style Drift Must Be Source-Relative

The protected surface `figure_table_captions_and_holders` is a source-relative
freeze surface, not a single sampled caption row.

For every thesis DOCX mutation, the protected-surface diff must extract and
compare the full source and final collection of:

- formal figure-caption paragraphs
- formal table-caption or table-title paragraphs, including in-table title rows
  when they are visible paragraph surfaces
- image-holder paragraphs that contain `w:drawing` or `w:pict`

Each row must include paragraph order, text, style id, paragraph-property hash,
run-property hashes, run count, and drawing/picture hashes when present. If a
non-caption target such as `body_text`, `toc_entries`, `page_numbers`, or
`references_entries` changes any figure/table caption paragraph style, direct
run typography, keep-with-next setting, alignment, spacing, indentation, or
holder drawing signature, the protected-surface diff must mark
`figure_table_captions_and_holders` as an unauthorized non-target change.

Body-style audits that only detect caption sibling contamination are not enough
to close this rule. The transaction must bind the source-to-final protected
surface JSON emitted by `scripts/audit_docx_protected_surface_diff.py` on the
exact final DOCX SHA256, and the transaction validator must parse that JSON
instead of trusting a pass-shaped acceptance statement.
