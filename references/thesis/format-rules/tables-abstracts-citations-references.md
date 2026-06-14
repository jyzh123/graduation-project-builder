# Thesis Format Rules: Tables, Abstracts, Citations, And References

Use this file for table-format, abstract-format, citation-format, and reference-block rules.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current format-repair subtask.
- Apply this file together with `references/thesis/thesis-format-rules.md`.

## 8. Table Rules

### FMT-TABLE-001. Body Table Headers Must Match The Locked Template Family

- For every body table, the table lane must lock the active table authority before mutation and classify the table family, such as glossary/role table, database field table, testing table, environment table, or project-specific summary table.
- Header text is part of table format, not only content. If the locked template family uses a specific header set such as `序号 / 词汇 / 描述`, `列   名 / 数据类型 / 说  明 / 备注`, or `测试项目 / 测试步骤 / 预计结果 / 实际结果`, the repaired table must use that header set unless the task record records a template-backed reason for a different family.
- Do not preserve source-table headers such as `测试项 / 预期结果 / 结果` or `表名 / 核心作用 / 关键字段` when the active template family requires a different header schema and the table can be expressed in that schema without losing meaning.
- Adding a missing template-required column, such as `测试步骤`, is table-format repair scope when the user reports table-header mismatch.
- The final evidence must list, for each touched body table: active donor table, donor header row, final header row, column count, family verdict, and any semantic adaptation rationale.
- A table cannot pass from border checks alone while its header row remains source-specific and incompatible with the template authority.

### FMT-TABLE-002. Three-Line Table Repair Must Verify Title Binding, Middle Rule, Width, And Cell Typography

- Body-table repair must verify all of the following against the locked template or school rule: donor table title mode, title keep-with-next binding when the donor uses an external title paragraph, title font/alignment/spacing, table width, top rule, header-bottom middle rule, bottom rule, absence of vertical borders, header-cell typography, body-cell typography, and cell paragraph alignment.
- The table title mode is a protected template-owned structure. The table lane must record `donor_title_mode` and `target_title_mode` for every touched body table. Valid modes are `external_paragraph`, `first_merged_row`, or another explicitly documented `other_template_owned_mode` backed by the donor.
- If the donor title mode is `first_merged_row`, the final table must keep the first merged title row, its `gridSpan`, row/cell properties, paragraph properties, and run properties. Do not lift that title into an outside paragraph only because a generic three-line-table memory says titles belong above the table.
- If the donor title mode is `external_paragraph`, the final table title must remain a dedicated paragraph immediately before the table and must keep with the following table. Moving it into a cell is a structure failure.
- A helper script may repair a title-mode mismatch only when the locked donor proves the target mode is wrong. A helper that cannot determine the donor title mode must stop with a blocker rather than guessing.
- A repaired table cannot pass because its header text is correct or because it roughly resembles a three-line table. The table lane must compare and record the locked donor's table properties, table grid columns, row heights, cell margins, cell widths, table layout mode, table alignment, table style id, title paragraph properties, title run properties, header/body cell paragraph properties, and header/body run properties before and after repair.
- When a template donor table exists with the same family and column count, the repair must apply the donor table's hard fields to the target table rather than reconstructing a builder-chosen approximation. Generic `TableGrid`, auto-width tables, or previously remembered fallback widths are invalid when a real donor table is available.
- Template table style ids are not portable by name or id alone. If the target document already contains the same style id with a different XML definition or a different style type, the table lane must clone the template table style under a collision-free id, rewrite the target table's `tblStyle` binding to that cloned donor style, and record the original id, target hash/type, template hash, and mapped style id. A table cannot pass when `tblStyle` points at a target-local style definition that merely has the same id as the donor.
- Table captions must use the numbering separator required by the school rule or donor sample, such as `表3.1`, and must not retain builder-generated alternatives such as `表3-1` when the active requirement shows dot-separated table numbers.
- A table with only top and bottom borders but no header-bottom rule is not a complete three-line table when the school rule requires `三线表`.
- A header-bottom rule must not be faked with header-cell paragraph borders (`w:pBdr`) or visible run underlines (`w:u`). Those create per-cell or per-word underlines instead of the locked table-family separator. Final table evidence must inspect and clear unintended header paragraph borders/run underlines while keeping the separator on the locked table or cell border surface.
- A table title placed after the table is a failed table surface even when the visible text is correct.
- A body table that still relies on empty/inherited header fonts after repair is blocked until the effective font chain is recorded and compared with the template header/body cell baseline.
- The validator or final acceptance record must have hard fields for donor title mode, target title mode, title-mode verdict, keep-with-next when external, donor header row, final header row, header-bottom middle-rule verdict, border family, width, header/body font size and bold state, and rendered template-vs-target comparison paths.

- Tables should follow the authorized school sample or project standard.
- Use true three-line tables by default unless the template explicitly requires another format.
- Do not place table directly after table without explanatory text between them.
- For thesis-format repair, verify both table borders and the relationship between table, table title, and surrounding text.
- Table-content anomalies are format failures. An incorrect field list, incorrect column meaning, or table-body content that no longer matches the surrounding thesis text still counts as thesis format-repair scope.
- If the current pass does not explicitly target tables, any table-style, table-cell-paragraph, or table-caption drift that appears after the pass is a tooling failure, not an acceptable side effect.
- Before any table generation, redraw, or repair pass, lock exactly one active table authority for the current manuscript:
  - explicit user instruction
  - current school template example
  - approved sample manuscript
  - project-local WPS preset lock
- Do not let a table pass infer its active family from a mixed bundle of historical notes in `references/thesis-table-style-memory.md`.
- If two historical project-specific table rules could both plausibly apply, the run must record why one authority is active for the current manuscript and why the others are out of scope.
- If a body paragraph, explanatory sentence, table title, caption fragment, or similar narrative text has drifted into a table cell, treat that as a table-local structure failure before any wording, tone, or style cleanup.
- Do not continue polishing text that is still trapped inside a table cell as if it were already back in body flow.
- A table-repair lane must explicitly verify that stray body or caption text is outside the table before handoff.
- Every external-mode table title must be a dedicated paragraph immediately before its table object, not a cell paragraph and not separated from the table by an empty/body paragraph.
- Every external-mode table title paragraph must keep with the following table and must not carry list numbering, outline level, or auto-numbering residue that can create strange visible numbering.
- Every in-table donor title row must be preserved as an in-table title row and must not be normalized into the external-mode rule.
- Table title typography and paragraph metrics must match the approved table-title paragraph instance; a correct `表 x` text string or correct style name is not enough if direct run formatting has drifted into the body baseline.
- Table-cell paragraph typography must be checked against the approved header/body cell baselines after table repair. Correct three-line borders do not prove the table text itself still matches the template.
- After any table insertion, table repair, or table-caption edit, the exact final DOCX must pass a table-caption binding check, not only a three-line-table border check.
- Table title binding, table-cell baseline, and table-family checks are generation-time gates. Do not postpone them to optional final polish after a sample has already been handed off.
- A generated thesis sample with table-cell text inheriting body fonts, body first-line indent, or builder-created font defaults is blocked even if the table borders visually resemble a three-line table.

### FMT-TABLE-003. Cross-Page Table Continuation Evidence Must Be Explicit

- When table repair, table insertion, or table QA is in scope, final evidence must state whether each real body table spans more than one rendered page.
- For every cross-page table, the table lane must record the table id/title, first page, continuation pages, rendered page-image paths, and the continuation-title policy from the locked template/sample.
- If the locked template/sample requires a continuation title family such as `续表 3-2`, every continuation page must have a standalone continuation-title paragraph above the continued table fragment.
- Continuation-title evidence must prove that the title is outside the table grid, centered or aligned according to the locked donor, first-line indent is zero, and `keepNext` binds the title to the continued fragment.
- A repeated header row alone is not a substitute for the required continuation title when the donor requires a continuation-title family.
- If no body table spans pages, the acceptance record must still say so explicitly and name the rendered table pages that prove no continuation pages exist.

### FMT-TABLE-004. Table Pagination Must Preserve Header Repeat And Row Integrity

- The table lane must verify `tblHeader` on header rows whenever a table can cross a rendered page, even if the current final layout fits on one page.
- Every body table row must be checked for row-splitting policy (`cantSplit` or a donor-backed equivalent). A row split that strands a partial row under a title, caption, or figure is a pagination failure.
- Table-title and table-body page binding must be checked from rendered output, not only from `keepNext` in OOXML.
- Final table evidence must include a rendered-page line/border check or page crop for every table page. A DOCX-only claim that top/header/bottom borders exist is insufficient when the user reported visible table breakage.
- The final acceptance record must include table continuation evidence paths, table continuation summary, cross-page rendered pages, continuation-title outside-grid verdict, and row-split/header-repeat verdict.
- The acceptance field label is `table row split/header repeat verdict`; it must pass or give a donor-backed not-applicable reason.

## No Table-As-Image Gate

- Thesis tabular data must be inserted as a real Word table object (`w:tbl`) with editable cells, not as a screenshot, PNG/JPG/SVG image, canvas export, or embedded picture that visually imitates a table.
- Do not caption a table screenshot or raster table as a formal `表 x.x`; it fails even when the visible rows and borders look correct.
- Do not caption a pure data table image as a `图 x.x` to bypass table requirements. If the content is mainly rows and columns, convert it into a real Word table and give it a formal table title.
- If the source artifact is a running-system screenshot that happens to contain a table, keep it only when the figure purpose is to prove the page UI, and the caption/prose must call it a page screenshot, not a thesis data table.
- Any figure/table acceptance record must explicitly state that no formal table is being represented by a raster image. When tabular data is present, final DOCX evidence must show an adjacent `w:tbl` object, title binding, and three-line-table styling.

### FMT-CAPTION-001. Caption Counts Are Not Figure/Table Format Evidence

- Figure and table caption counts are inventory signals only. They do not prove caption format, numbering separator, title binding, or image/table adjacency.
- The figure/table lane must enumerate every formal figure caption and every formal table title in the body, excluding cover/front-matter layout tables from body-table counts.
- Each caption/title must compare against the active template donor for numbering text, separator style, paragraph alignment, spacing, font slots, font size, keep-with-next behavior, and the adjacent image or table object.
- A count-only pass such as `10 figure captions, 6 table captions, 8 tables` is blocked when any caption uses the wrong separator, loses its descriptive title, is detached from its object, or is backed only by generic inventory evidence.

## 9. Abstract Rules

### FMT-ABSTRACT-001. Keyword Label And Content Must Be Structurally Split

- Keyword-line separation is a structural DOCX requirement, not only a visual formatting preference.
- For `zh_keyword_line` and `en_keyword_line`, the keyword label token, separator, and keyword content must be machine-verified from the final DOCX `word/document.xml` run structure.
- A pass requires:
  - the label run or label run sequence is isolated from the content run sequence
  - the compact label run text exactly equals the expected label token and separator, not a prefix that also contains the first keyword
  - the content run count is at least one
  - the content runs do not inherit label-only bolding unless the locked template baseline explicitly proves full-line bolding for the same surface
  - the evidence names the label text, separator, content run count, label font/bold or template-approved strong face, content font/bold state, and extraction method
- Evidence that only says `Chinese keyword label/content run split confirmed: yes` or `English keyword label/content run split confirmed: yes` is not enough. The validator must reject keyword-surface evidence unless the same record contains the machine-extracted fields and the final DOCX structure passes direct inspection.
- The required machine-extracted evidence fields include the keyword run split extraction method, label run text, label isolation verdict, content run count, label bold or strong-face state, content bold state, separator, and keyword run split verdict.
- Even when the keyword line is stored as multiple runs, it still fails if the label run and the content run collapse to the same effective bold or strong-face state when the locked template requires `label bold, content normal`.
- A visible line such as `关键词：内容` or `Key Words: content` is still failed when the label and content are stored in one uniform-format run.
- A label run is not isolated when its compact text starts with the expected label but continues into the first keyword. The validator must compare compact label-run text by exact equality against the expected label token and separator, and must emit `label run must contain only the label text` for that failure.
- A keyword paragraph must not use a heading, title, TOC-title, TOC-entry, caption, or outline-level paragraph style. A keyword line with correct text and correct label/content run split still fails if its `w:pStyle`, resolved style name, or direct `w:outlineLvl` makes it behave as a title/heading surface.
- A keyword content run must not inherit abstract-title, title, heading, caption, TOC-title, or title-like character formatting. A keyword line with correct label/content run split still fails when the content run carries a title/heading character style, title-size direct font, strong title face, or other effective title-like run formatting. The only exception is a locked donor baseline for the same protected keyword surface that explicitly proves full-line title-like formatting is required.
- Treat this as a named `keyword content title-style contamination` hard failure in rule-owner maps, surface evidence, and acceptance records.

- For thesis format-repair, abstract problems are not layout-only issues. Chinese/English abstract inconsistency is a format failure and must be repaired in the same pass.
- Do not treat repairing only the TOC-to-abstract seam, page break, or front-matter page order as proof that the abstract block itself is fixed.
- An abstract-format pass is complete only after all six front-matter surfaces have been checked independently:
  - Chinese abstract title
  - Chinese abstract body
  - Chinese keyword line
  - English abstract title
  - English abstract body
  - English keyword line
- The six-surface abstract check must include the full formatting face of each surface, not only text and pagination:
  - paragraph style binding and outline level
  - Chinese and Western font family attributes
  - font size, bold, italic, underline, color, and superscript/subscript state
  - alignment, indentation, line spacing, and spacing before/after
  - keyword label/content run split and label-only bolding
  - abnormal ASCII spaces, full-width spaces, tabs, or stretched full-justification artifacts
- Abstract font-family evidence must resolve the effective font chain through direct run properties, character style, paragraph style, basedOn chain, docDefaults, theme major/minor mappings, and WPS/Word UI displayed font names. A run with no explicit `w:rFonts`, or a style that displays as `Calibri (正文)` / theme body font, cannot be marked as matching merely because the direct run properties are empty.
- English technical terms inside the Chinese abstract and English abstract body text must use the Western font mapping from the locked template/sample baseline. Do not repair them with an ad hoc English font or size chosen by the builder.
- If the user reports an abstract format issue, the next pass may not check only whether the abstract remains on one page or visibly overflows. It must repair and re-check Chinese abstract title, Chinese body, Chinese keywords, English title, English body, and English keywords as one protected front-matter family.
- The Chinese keyword label must be bold.
- `Key words:` in the English abstract must be bold.
- If the approved sample or active rule requires `Key words:` and the manuscript still uses `Keywords:` or another label form, treat that as an abstract-format failure rather than a wording-only preference.
- Only the label text should be bold by default; the keyword content after the label should remain normal unless the template explicitly requires full-line bolding.
- The keyword label and the keyword content must be stored and formatted as separate text runs whenever the document toolchain permits it.
- Do not leave the keyword label and keyword content merged into one uniform-format run when the required result is "label bold, content normal".
- Keyword-line evidence must cover both row-level formatting and label/content run-level formatting. The record must state whether the label is bold, whether keyword content is non-bold unless the template requires otherwise, what separator is used, how runs are split, and which East Asian / Western / complex-script font slots each run uses.
- A keyword-line repair that first applies one paragraph-wide font or bold state and only then appends label/content runs is still failed unless the final DOCX proves distinct effective label/content formatting at run level.
- English abstract verification is not only an English-page typography check. It must be compared against the final Chinese abstract for topic, method, system functions, result claims, and keywords; stale project descriptions, omissions, or extra claims block format acceptance because the formatted front matter would describe the wrong thesis.
- Abstract body text should remain intact paragraph blocks and should not be broken into awkward fragments by formatting logic.
- Abstract body generation must preserve the template first-line strategy: template square placeholders such as `□□` mean required blank spaces and must be converted to spaces in the generated manuscript; replacement content must stay in the template content run rather than the placeholder-prefix run, and no visible `□` may remain.
- English abstract content fonts must come from the template's effective content run or explicit adapter baseline, not from placeholder prefix runs or builder-chosen English fonts.
- Keyword lines must preserve template label, separator, and content run structure; label/content merge or missing label/content run separation is a hard generation failure.
- Keyword formatting donors must come from a real keyword or abstract surface in the locked template. A TOC entry, visible template instruction paragraph, red annotation, or front-matter format note may not donate keyword paragraph or run properties.
- If the template keyword donor is instruction-like or TOC-styled, the helper must fall back to a target abstract-body donor and record that fallback in the report, for example `target_abstract_body_fallback`; silent borrowing from TOC or instruction text is a failed keyword repair.
- Table-cell generation must choose the donor run by content family. Latin-only or digit-only table cells must use the template's Latin/digit cell run donor when one exists instead of forcing the CJK cell run onto English identifiers or numeric fields.
- If one abstract language version has been revised while the other still reflects stale wording, stale claims, stale structure, or stale scope, the abstract block remains formatting-incomplete.
- If an abstract title or keyword line only receives direct formatting while still remaining bound to the generic body paragraph class, the abstract repair is incomplete. Treat that as a front-matter class-binding failure, not as an acceptable shortcut.
- Abstract six-surface parity is a generation-time gate. A thesis generator must verify Chinese abstract title, Chinese abstract body, Chinese keyword line, English abstract title, English abstract body, and English keyword line before handoff.
- If keyword label/content runs merge into one run, if only the whole line is bold, or if English technical terms use builder-chosen Western font mappings, the abstract block is blocked even when page order and visible text look acceptable.
- An abstract or keyword repair is also blocked when the implementation rebuilds the whole line or paragraph from one inherited donor run and then tries to patch font fields afterward. That path destroys the original label/content or mixed-script structure before verification and may be used only when the canonical builder owns the full replay plus evidence.
- Abstract and front-matter generation must remove template instruction artifacts such as visible callout text boxes, arrows, font-size notes, keyword instructions, and `注:摘要单独成页` notes; any such residue is a hard failure, not a cosmetic warning.

### FMT-ABSTRACT-002. Chinese Abstract Mixed-Script Fonts Must Be Verified Per Run

- English letters, abbreviations, package names, model names, percentages, and other Latin/digit fragments inside `zh_abstract_body` are independent font-chain surfaces, not incidental text inside a Chinese paragraph.
- A Chinese abstract body pass requires the final DOCX to be machine-inspected from `word/document.xml`, not only checked through the first body run, paragraph style name, rendered screenshot, or broad effective-font prose.
- The checker must locate the Chinese abstract title, collect every non-empty body paragraph before `关键词`, and inspect every run containing ASCII letters or digits.
- For each Latin/digit run, the checker must resolve `ascii`, `hAnsi`, and `cs` through direct run properties and the paragraph style baseline. A direct or resolved builder/default font such as `Calibri` fails unless the locked template baseline for that same surface explicitly proves it.
- If a Chinese abstract body paragraph contains Latin/digit text but no separately auditable Latin/digit run can be found, the surface is blocked until the run structure is repaired or the template baseline proves the same structure.
- Evidence that only says `Chinese abstract body confirmed: yes`, checks the first content run, or compares one paragraph-level font summary is not enough. The final DOCX must be re-read and the actual Latin/digit run slots must pass direct inspection.

### FMT-ABSTRACT-003. English Abstract Indentation Must Be Baseline-Bound And Directly Audited

- English abstract body indentation is a protected front-matter surface. It may not be repaired by leading spaces, full-width spaces, tabs, or a visually similar body paragraph.
- When the user reports English abstract indentation, abstract spacing, or `Abstract` page formatting, the final evidence must inspect the English abstract title, English abstract body, and English keyword line from the final DOCX.
- The English abstract body evidence must record the locked baseline source, paragraph index or OOXML path, `w:ind` values (`firstLine`, `firstLineChars`, `left`, `right` when present), alignment, leading-space verdict, style binding, and rendered page path.
- A pass requires either direct `w:ind` values matching the locked template/sample baseline or an explicit template-backed zero-indent rule. A hard-coded two-character value without baseline proof is not enough for new manuscripts.
- If `jc=both` or other full-justification creates visibly stretched English words after indentation repair, the rendered page review must fail until the alignment is corrected against the baseline.
- The final acceptance record must include an English abstract indentation evidence path and an English abstract indentation verdict whenever this issue was user-reported.
- When an approved abstract baseline profile is supplied, the six protected abstract surfaces are bound to that profile as the donor. The checker must compare paragraph metrics and direct/effective font slots against the profile surface itself and must not mix in style-inheritance values from a different active template document for the Latin/digit run decision.

## 10. Citation And Reference Rules

Canonical owner boundary: this section is the format router for citation and reference surfaces. Final acceptance authority for body citation superscript state, first-appearance order, numeric marker shape, hyperlink target correctness, and stale-report rejection belongs to `references/user-feedback/citations-and-bibliography.md` rules `FB-CITE-002`, `FB-CITE-034`, `FB-CITE-035`, and `FB-CITE-049`, enforced by `scripts/audit_thesis_citations.py` and `scripts/validate_skill_gate_record_gate.py`. Do not clear citation acceptance from this format section alone.

- Citation markers should use superscript formatting and be placed before the sentence-ending punctuation mark.
- Citation markers that support document-internal jumps must still render like thesis body superscripts rather than generic hyperlink styling.
- Citation markers are their own run-level surface. They must remain separate superscript runs with the template-approved marker font/position, not plain body text, not hyperlink-default blue/underline text, and not merged into a rewritten sentence run.
- Accepted citation-marker visual style:
  - black text
  - no underline
  - superscript preserved
- If a citation marker is clickable but still shows blue text, underline, or other default hyperlink formatting, treat that as a formatting failure.
- Verify citation order from first appearance in body text, not only from the bibliography list.
- Treat citation order from first appearance in body text as its own acceptance gate.
- Do not infer citation-order correctness from preserved bibliography numbering, existing hyperlink targets, or a successful citation-style rebuild alone.
- Enforce one sentence at most one citation marker.
- Enforce one citation marker exactly one citation number.
- Do not leave merged forms such as `[1-3]`, `[1,2]`, or `[1][2]` in the final body.
- Every bibliography entry must be cited at least once in the thesis body.
- When citation numbers are repaired, keep the bibliography numbering synchronized with the body.
- When repairing from an existing manuscript, the final bibliography must not contain fewer real entries than the source/review-copy unless the user explicitly requested entry deletion and the post-deletion citation audit proves the new list is synchronized.
- Do not treat "citation audit passed" as sufficient if a later rewrite or repair has silently dropped bibliography entries; bibliography count and bibliography formatting are separate proof surfaces.
- Treat citation scan and citation repair as their own pass before spacing out special headings or other formatting-owned blocks.
- When editing citation text in DOCX, assume the visible marker may be split across multiple runs. Inspect those runs with `officecli view annotated` or targeted `get` calls, and do not trust `paragraph.text` replacement as a real edit path.
- Treat every citation-bearing body paragraph as a protected mixed-format surface. Do not collapse body text and citation markers into one replacement run or one raw paragraph-text assignment.
- When a sentence around a citation marker is rewritten, preserve the citation marker as its own superscript run and re-check the rendered punctuation order immediately.
- After run-level deletion or replacement, check for leftover punctuation or empty-run artifacts.

### Citation Override Gate

- If the user explicitly requires stricter citation controls, those controls override the generic thesis defaults immediately.
- Current hard-override citation rule: each sentence may contain at most one citation marker.
- Current hard-override citation rule: each citation marker may contain exactly one citation number.
- Current hard-override citation rule: merged forms such as `[1-3]`, `[1,2]`, and `[1][2]` are not allowed.
- Current hard-override citation rule: every citation marker must be superscript and must sit on the left side of the punctuation mark it belongs to.
- Current hard-override citation rule: every bibliography entry must be cited at least once in the thesis body.

## 11. References, Acknowledgement, Appendix

- Treat references, acknowledgement, and appendix as independent formatting blocks.
- References should use the required bibliography layout, not generic body text.
- Reference title, reference entry paragraphs, acknowledgement title, acknowledgement body, appendix title, and appendix body are separate protected classes. They may not be normalized into body/Normal style as a shortcut; each class must be checked against its own approved donor or an explicit missing-baseline blocker.
- Tail-block donor selection must be class-locked. A reference-entry donor may come only from a real bibliography entry after the formal reference title, never from the reference title, TOC row, acknowledgement, appendix, or body heading. Appendix body formatting may not borrow numbered bibliography-entry formatting; if the template has no appendix body donor, use an approved body-prose donor or fail closed with a missing-baseline blocker.
- Appendix body paragraphs often contain drawing numbers, part numbers, model identifiers, sheet codes, and other Latin/digit strings. If the approved appendix/body-prose donor has no Latin run baseline, the checker must still compare paragraph metrics and CJK typography but must not fail only because the final appendix text contains Latin/digit tokens. Latin-run slots are blocking for appendix body only when the donor explicitly provides a Latin baseline or a separate Latin/digit appendix donor is locked.

### FMT-REF-001. References Title And Entries Must Bind To The Formal Tail Block

- The `references_title` detector must bind to the formal end-matter title paragraph only. A TOC row such as `参考文献38`, a body mention, or any paragraph whose compact text is not exactly the reference title cannot serve as the `references_title` surface.
- The formal `references_title` is also a pagination surface: it must start after the previous real content block on rendered output, and its evidence must not pass unless the previous-content page and reference-title page are both recorded with a passing prior-block separation verdict.
- When rendered PDF text wraps the previous real content paragraph, the checker must locate the previous-content page by normalized stable excerpts instead of requiring the full DOCX paragraph string to appear verbatim on one PDF text line/page.
- The `references_entries` detector must start after the formal `references_title` paragraph and must bind only to bibliography-entry paragraphs such as `[1] ...` or the template-approved bibliography numbering form.
- Rendered/full-thesis bibliography entry counters must accept the active template's approved visible numbering forms, including `[1] ...`, `［1］ ...`, `1. ...`, `1．...`, and `1、...`, while still requiring the entries to occur inside the formal reference block.
- If the measured `references_entries` paragraph is the reference title, a TOC row, an acknowledgement/appendix title, a generic heading, or a non-entry paragraph, the evidence is failed as `title-as-entry` or wrong-target binding.
- Reference evidence must record the DOCX paragraph/run path or paragraph index for the formal title and first real entry. Page presence, entry count, or tail-page order cannot prove reference formatting by itself.
- Visible bibliography label spacing must be locked from the active school template, approved sample, or validated project adapter for the exact document. Do not hard-code a global compact `[1]内容` family or a spaced `[1] 内容` family into the skill; either family can be correct only when the current template/adapter evidence proves it. The final reference audit must report which label-content spacing family was locked and must bind that verdict to the exact final DOCX/PDF pair.

### FMT-REF-002. Reference Entry Format Cannot Pass From Count Or Citation Checks

- Bibliography entry count, citation order, hyperlink integrity, or visible `[n]` numbering are supporting signals only.
- A reference-entry pass requires the `references_title` and `references_entries` protected-surface records to pass paragraph-dialog metrics, typography metrics, rendered geometry, entry font-slot checks, and exact-output SHA binding.
- A final acceptance record must not synthesize `references entry format verdict: pass` from count inventory, citation audit, or broad bibliography status when either `references_title` or `references_entries` has failed, stale, generic, or wrong-target evidence.
- For the current SGB620/80T graduation-project run, count-only success and citation-chain success are supporting signals only. They cannot override a visible bibliography family failure, a `label-content spacing none` failure, or a bibliography block that was copied from generic body text instead of replayed from a real donor.
- Hanging indent alone is not sufficient proof that the bibliography format is correct.
- Reference-entry indentation is a baseline-owned surface. The checker must compare each final bibliography entry against a real template/reference entry for left indent, right indent, first-line indent, hanging indent, character indents, line spacing, spacing before/after, style binding, and run fonts.
- Reference-entry run fonts are a mixed-script surface. For each bibliography run that contains Chinese text, the East Asian font slot (`eastAsia` or its template-approved theme equivalent) must match the locked Chinese reference donor. For each bibliography run that contains English letters, digits, URL, DOI, ASCII punctuation, or bracketed numbers, the Western/complex-script slots (`ascii`, `hAnsi`, and `cs`, or their template-approved theme equivalents) must match the locked Western reference donor.
- Reference-entry font size is a hard field for every bibliography entry, not a sampled typography note. Evidence must expose `references entries font-size baseline/actual`, a per-entry font-size map, and a pass/fail verdict bound to the exact final DOCX SHA256. A bibliography count, citation-order pass, visible `[n]` numbering, or first-entry-only sample cannot clear the `references_entries` surface.
- Do not repair references by classifying whole entries as "Chinese entries" or "foreign entries" and applying one paragraph-wide font. A single bibliography entry can contain Chinese, English, numbers, DOI, URL, and punctuation; the checker must validate every relevant run slot rather than only the first run or the paragraph metrics.
- A bibliography run that mixes Chinese characters with English letters, digits, DOI/URL text, ASCII punctuation, or bracketed reference numbers is a structure failure unless the active template explicitly proves the same mixed run as the donor. The default repair is to split the entry by script and apply the locked Chinese and Western donors separately.
- The DOCX font/encoding audit must run with the active template or approved sample as `--reference-docx` whenever references are present. A font audit that reports only mojibake status and does not report bibliography font-slot checks is not valid completion evidence for reference repair.
- If bibliography numbering is generated by `w:numbering.xml` instead of visible `[n]` runs, the numbering marker font slots are also part of the reference-entry surface. Until those numbering-level slots are audited against the template donor, the repair cannot be marked as passed.
- A bibliography entry that still carries the body paragraph style, body first-line indent, or body paragraph family after font/size repair is failed. Treat this as a reference-entry class-binding failure even when the text and citation order are correct.
- A bibliography repair is blocked if it checks only entry count, citation order, or link integrity while leaving abnormal entry indentation, unexpected hanging indent, body first-line indent residue, or theme-font alias drift.
- First verify that the rendered page has actually entered the bibliography section, then check heading, paragraph rhythm, line spacing, indentation, and visual density as a dedicated block.
- Verify their order and formatting against the template before delivery.
- The title paragraphs for `参考文献`, `致谢`, and similar tail blocks are template-owned title surfaces. Repair them from a real template/sample title paragraph instance rather than from generic body formatting or a hand-scripted font-name reconstruction path.
- `acknowledgement_title` must expose its actual title style binding, paragraph metrics, run typography, and font size. A centered visible `致谢` string or generic tail-block title evidence is insufficient unless `acknowledgement title style baseline/actual` and `acknowledgement title paragraph style verdict` pass.
- Footer page numbers are a typography surface. Footer evidence must expose the page-number field/run path and font size with `footer page-number font-size baseline/actual`, `footer page-number run path map`, and `footer page-number font-size verdict`; a PAGE field, correct numbering restart, or rendered page token alone cannot pass the footer/page-number surface.
- If a repair path writes mojibake or unreadable font-family names into those title paragraphs, treat the result as a hard failure even when the renderer falls back to a readable font.

## 12. Broad Format Scope

- In thesis format-repair, "format issues" include layout, pagination, spacing, numbering, typography, and content-delivery anomalies that break the formatted meaning of the thesis.
- Typography includes the complete font face for every touched surface: Chinese font, Western/English font, complex-script font when present, size, weight, color, emphasis, and run splitting. Do not narrow typography to only Chinese font size or visible boldness.
- Typography evidence must distinguish explicitly written fonts from inherited effective fonts. If the font is inherited from a style, basedOn style, docDefaults, or theme alias, the evidence must name that source and compare it to the template baseline before any pass verdict.
- A user-reported format issue on any protected surface must trigger a same-class audit. The pass is incomplete if it fixes only the visible hotspot while leaving sibling surfaces with the same baseline mismatch.
- A formula-numbering block that survives only as a standalone paragraph under the equation is still a format failure when the active template requires same-line numbering at the far-right edge.
- Treat the following as format failures that must be repaired inside the thesis-format pass:
  - Chinese and English abstract mismatch
  - incorrect text content inside a formatted block
  - incorrect table content
  - incorrect figure content
  - caption-reference-image mismatch
  - caption-reference-table mismatch
  - stale sample text left inside a front-matter, abstract, figure, table, or other template-owned block
  - mismatches between正文描述、图题、表题 and the actual delivered figure/table content
