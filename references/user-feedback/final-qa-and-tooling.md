# User Feedback Persistence: Final QA And Tooling

Use this file for durable final-review, render-verification, abstract, pagination, and thesis DOCX tooling corrections.

## Enforcement Status

- Every numbered rule in this file is mandatory when this file is loaded for the current subtask.
- Apply these rules together with `references/user-feedback-persistence.md` and the active thesis toolchain references.

## Final QA And Tooling Rules

### QA-FINAL-001 (legacy 31). Final Thesis QA Must Check Body Font And Size Consistency (Mandatory)

- After all content merges and formatting edits, run a whole-document consistency scan on body paragraphs.
- Verify that the main body uses one unified font/size baseline rather than relying on visible similarity.
- Treat mixed body size, mixed East Asia font mapping, or accidental style fallback to `Normal` as release-blocking formatting errors.

### QA-FINAL-002 (legacy 31A). Body Paragraphs Must Be Explicitly Bound To The Body Style, Not Left On Implicit Default Fallback (Mandatory)

- Do not accept a thesis body merely because most body paragraphs visually resemble the sample.
- For the main body paragraph family, verify two surfaces separately:
  - the visible body formatting matches the sample or template
  - the paragraph XML is explicitly bound to the intended body style rather than relying on omitted `pStyle` fallback or scattered direct formatting alone
- If a repair pass bulk-formats body paragraphs through paragraph properties or run properties but leaves the body family without explicit style binding, treat that pass as incomplete.
- If the sample uses the template default body style, still write the explicit binding in the repaired manuscript when the user's office app would otherwise display the paragraph as style-less or inconsistently styled.

### QA-FINAL-003 (legacy 32). TOC Correctness Requires Order And Visual Structure (Mandatory)

- A TOC is still wrong if all headings are present but the order is reversed, indentation is flattened, or page-number presentation breaks.
- Final TOC QA must confirm:
- heading order is top-to-bottom correct
- indentation reflects heading depth
- visible TOC styles still match the approved local template or accepted sample rather than generic renderer defaults
- front-matter page-display strings still match the approved numbering convention, including forms like `摘要I / ABSTRACTII / 目录III` when the sample uses that pattern
- a TOC pass still fails if the visible entries fall back to generic styles such as `TOC1/TOC2/TOC3` while the approved sample uses a different local style presentation

### QA-FINAL-004 (legacy 33). Blank Artifacts And Root-Cause Repair Are Mandatory (Mandatory)

- In thesis formatting work, abnormal blank pages, abnormal blank lines, and abnormal extra spaces are format errors, not cosmetic noise.
- Treat any of these artifacts as release-blocking until they are removed or explicitly justified by the school template.
- Check them in both structure review and rendered-page review.
- If rendered-page review still shows abnormal blank pages, do not stop at reporting them.
- Identify the exact cause before handoff:
  - residual blank paragraphs near figures or tables
  - floating-shape anchor drift
  - oversized inline figures
  - manual page breaks or section breaks
  - TOC/body duplication or other structural drift
- Remove the cause and rerun rendered-page review.
- A thesis with known abnormal blank pages is incomplete even if all other format classes look acceptable.
- Page references must remain aligned and readable.
- Heading lines must not break or wrap abnormally after the repair.

### QA-FINAL-005 (legacy 34). Final PDF Page Review Is Mandatory (Mandatory)

- Before handing off a thesis, render the final PDF and inspect:
- front matter pages
- at least one image-heavy middle page
- references page
- acknowledgement page
- final tail page
- Use this review to catch blank trailing pages, misplaced captions, TOC breakage, and font drift that paragraph-level DOCX inspection may miss.
- `officecli view html` is not a valid replacement for this step. Final PDF page review must inspect actual rendered PDF pages or page images derived from that PDF.

### QA-FINAL-006 (legacy 35). Abstract Formatting Is A Protected Surface (Mandatory)

- Do not treat the Chinese abstract page and English abstract page as ordinary body paragraphs during whole-document normalization.
- Check the abstract body and keyword lines as dedicated front-matter surfaces.
- Do not report the abstract as fixed if the run only repaired TOC-to-abstract seam pagination or page-order drift.
- Final abstract QA must explicitly cover six independent surfaces:
  - Chinese abstract title
  - Chinese abstract body
  - Chinese keyword line
  - English abstract title
  - English abstract body
  - English keyword line
- A global font or spacing pass that fixes the body but disturbs abstract layout is still a failed delivery.
- If abstract formatting is wrong, restore the abstract paragraphs and keyword lines from the template source rather than relying only on generic body normalization.
- When English abstract indentation is user-reported, final QA must include an English abstract indentation evidence path and verdict. The evidence must record the final DOCX paragraph index or OOXML path, direct/effective `w:ind` values, alignment, leading-space verdict, and rendered English abstract page path.

### QA-FINAL-007 (legacy 36). Every Chapter Must Start On A New Page (Mandatory)

- Thesis chapter-level headings must begin on a fresh page.
- Do not assume template pagination will survive later insertions, deletions, or spacing cleanup.
- During finalization, explicitly preserve or reinsert page breaks before each chapter heading and verify this on rendered pages.

### QA-FINAL-008 (legacy 37). English Abstract Must Be A Translation Of The Chinese Abstract (Mandatory)

- Do not independently rewrite the English abstract with new structure, extra claims, or different emphasis.
- The English abstract must be a faithful translation of the Chinese abstract.
- It may be polished into natural English, but its information coverage, paragraph intent, and scope must match the Chinese source.
- The same rule applies to keywords: the English keywords should correspond to the Chinese keywords rather than introducing a different set.

### QA-FINAL-009 (legacy 38). Default Thesis DOCX Toolchain Is OfficeCLI First (Mandatory)

- For thesis `.docx` work under this skill, use `references/tooling-dependencies.md` as the canonical toolchain source.
- Use `references/thesis/thesis-format-sop.md` as the canonical execution-order and OfficeCLI-routing source.
- Do not replace that canonical toolchain with ad hoc `python-docx`-first mutation, broad paragraph rewrites, or mixed unsafely overlapping repair paths.
- Final QA should confirm that the run actually followed the canonical DOCX path rather than only producing a superficially correct-looking output.

### QA-FINAL-010 (legacy 38A). Missing Formula Numbers Are A Format Failure (Mandatory)

- Treat missing formula numbers as a thesis format failure, not as a content-only omission.
- A formula task is incomplete if the manuscript contains real equation objects but no visible numbering where the active template requires formula numbers.
- Final QA must verify both surfaces separately:
- formula object presence
- visible formula numbering family, placement, and alignment
- rendered-page confirmation that the formula number reaches the far right end of the formula line
- rendered-page confirmation that the formula number does not exceed the active section's usable right boundary
- If the accepted template uses chapter-based numbering such as `（式4-1）`, `（式4-2）`, keep that family and do not downgrade to unnumbered formulas or generic flat numbering.
- If the formula number is present but stops short of the line end, treat that as a failed numbering placement and repair the paragraph's numbering surface explicitly, for example by restoring the template's center/right tab-stop grammar when that is the active layout family.
- If the formula number is pushed beyond the right margin by a fixed tab stop or over-wide borderless table, treat that as the same class of failed numbering placement; the repair must recompute tab stops or table width from the active section page width minus margins.

- If the formula number survives only as a standalone paragraph below the equation, treat that as a release-blocking failure even when the number text itself is correct.
- Do not accept a formula evidence summary that describes the accepted result as `numbering paragraph` or `numbering paragraphs` when the required surface is same-line right-aligned numbering.

### QA-FINAL-011 (legacy 39). Image Paragraphs Must Not Inherit Body Line Spacing (Mandatory)

- After inserting or replacing any thesis figure, do not leave the image paragraph on the body-text line-spacing rule such as fixed 22pt.
- The paragraph that contains only the image object must be normalized separately:
- center aligned
- no first-line indent
- no left or right indent
- single line spacing
- spacing before and after set explicitly for the figure block
- no heading-like, TOC-like, or caption-like paragraph class
- no outline level or list metadata that could let TOC refresh treat the image paragraph as an outline candidate
- The caption must be in its own following paragraph and must not share the image paragraph.
- If the visible result still shows a flattened or hidden image, treat that as an image-paragraph formatting failure rather than only an image-size problem.

### QA-FINAL-012 (legacy 63). Touched Figure And Table Pages Must All Be Re-Rendered After Late Content Passes (Mandatory)

- If a late-stage thesis pass adds explanatory text, changes caption adjacency, centers tables, or modifies paragraph spacing near figures or tables, do not rely on earlier figure/table render checks.
- Re-render every touched figure page and every touched table page after that pass.
- For each touched figure block, verify all of the following on the rendered page:
- the image is fully visible
- the caption is visible
- the description paragraph is visible if one is required
- the image, description, and caption remain in the intended order
- no part of the block is stranded onto a different page without the rest
- For each touched table block, verify all of the following on the rendered page:
- the caption is not orphaned at page bottom
- the header row is not stranded without data rows
- the active border family still matches the required table style
- If a touched figure or table page was not re-rendered after a late content/layout pass, the thesis-format task is incomplete.

### QA-FINAL-013 (legacy 68). Structure-Level Checks Must Not Be Mistaken For Thesis Acceptance (Mandatory)

- Do not treat `officecli validate`, DOCX media counts, paragraph outlines, heading presence, or skill-gate success as proof that the thesis itself is ready to submit.
- These checks only prove partial structure, not rendered-page quality.
- Final thesis acceptance must still be blocked when any of the following remain true on rendered pages:
  - a figure style visibly deviates from the approved sample
  - a screenshot is damaged, blank, partially rendered, or stuck in a loading skeleton state
  - an image is detached from its caption
  - a figure or table is detached from its required explanatory paragraph
  - page-level formatting remains visibly inconsistent with the template
- Final thesis acceptance must also be blocked when the current run requires paper-only literature and one or more bibliography items are still non-paper sources, unverifiable sources, or body dependencies on rejected sources.
- If structure-level checks pass but rendered-page review still fails, report the thesis as failing. Do not describe that state as "complete" or "ready".
- Final acceptance cannot use a local or sampled page check as proof of whole-thesis template alignment. A whole-thesis, full-paper, `1:1`, template-aligned, or submission-ready claim requires a page-class coverage matrix covering cover, title/front matter, Chinese abstract, English abstract, TOC, body/chapter, figure, table, references, acknowledgement, and appendix.
- Final acceptance must bind thesis figure work to the canonical figure asset manifest and figure contract. A `passed` figure-family summary without draw.io/SVG/raster fallback evidence for structural figures, or without source-scale geometry validation and relation-attribute collision evidence for ER/dense structural figures, is not acceptance evidence.
- Final acceptance must carry a user-reported issue ledger whenever the user names a surface problem such as abstract, keyword, references, bibliography, citations, TOC, body style, table style, or figure style. Each ledger item needs the user's wording, the affected surface, the expected fix, an evidence path, and a final verdict.
- The user-reported issue ledger is mandatory for any named thesis surface, including TOC font drift, `宋体（正文）` theme-font alias drift, reference-entry indentation, body-font drift, table style drift, and figure style drift. A summary that says "checked" or "fixed" without one ledger row per named issue is not acceptance evidence.
- For figure-related comments or feedback, final acceptance must also carry the figure comment conversion checklist, figure plan, per-figure task-card paths, per-figure evidence manifest, figure asset manifest, and rendered-page review evidence. A broad `figure style checked` summary is not acceptance evidence.
- Final acceptance must carry the mandatory thesis surface inventory path plus explicit verdicts for cover style, abstract and keyword surfaces, TOC visual baseline, reference-entry format, and appendix format. A thesis handoff that lacks any one of those fields is incomplete even if the changed paragraph or chapter passed local review.
- Final acceptance must also carry a high-risk thesis format surface matrix path and an overall high-risk thesis format surface verdict. This matrix exists specifically to prevent repeated omissions of cover, abstract, keyword, TOC, reference-entry, and appendix formatting checks.
- Generated or regenerated thesis deliverables must include a non-empty `sample_self_check` report for the exact handed-off DOCX; omitting the report is itself a delivery blocker.
- A sample self-check report that contains `smoke acceptance mode: yes`, `smoke-only; blocked for delivery`, `full thesis content gate failed`, or a critical failed detector blocks final acceptance even when the surrounding acceptance record says `passed`.

### QA-FINAL-014 (legacy 69). Caption Wording And Footer Presentation Are Final Acceptance Surfaces (Mandatory)

- Final thesis QA must treat visible figure/table caption wording and visible footer/page-number presentation as independent acceptance surfaces rather than as cosmetic polish.
- Required caption QA items:
  - the caption wording matches the approved local sample or template class
  - the caption contains no editorial provenance note, build note, placeholder note, or similar extra parenthetical text unless the sample explicitly requires it
  - the caption paragraph still uses the approved caption class rather than a body-text fallback
  - when the manuscript contains code titles such as `代码 1 ...` or `程序清单 1 ...`, those code-title paragraphs must also follow the approved code-title baseline instead of falling back to ordinary body text
  - code blocks beneath those code titles must keep their own code-format baseline rather than inheriting the surrounding Chinese body paragraph font, body indent, or body line spacing
- Required footer QA items:
  - the footer and page number match the approved sample in alignment, font rhythm, visibility, and numbering presentation
  - the review explicitly checks footer appearance on rendered pages, not only section settings or field presence
- If either the caption wording or the footer presentation still differs from the approved baseline, the manuscript is not submission-ready even when the structure, section count, and visible page numbers look plausible.

### QA-FINAL-015 (legacy 72). Figure Pagination Must Keep The Image Paragraph And Caption On The Same Page (Mandatory)

- A figure is still failing if the image appears on one page and the caption drops to the next page, even when the caption paragraph text exists in the DOCX.
- Treat the minimum accepted figure unit as:
- image-holder paragraph
- figure caption paragraph
- Required default pagination behavior for thesis figures:
- image-holder paragraph uses a non-body layout baseline
- image-holder paragraph keeps with the following caption paragraph
- if the full figure block does not fit in the remaining page area, move the block boundary before the figure instead of letting the caption orphan
- If a user screenshot shows a visible figure with no caption below it, do not classify that as a missing-title problem alone. Treat it as a figure-pagination failure.

### QA-FINAL-016 (legacy 73). Table Captions Must Sit Above The Table And Stay Bound To The Table Start (Mandatory)

- In thesis body chapters, table captions belong above the table by default unless the approved template explicitly overrides that layout.
- The minimum accepted table-start pagination unit is:
- table caption paragraph
- header row
- first data row
- Required default behavior:
- table-caption paragraph uses `keepNext`
- table-caption paragraph stays on the same page as the start of the table
- the table header separator line remains visible on the first rendered page of the table
- If a user screenshot shows the caption on one page while the table body starts on the next page, or shows the table body with no visible caption above it, treat that as a release-blocking pagination failure.

### QA-FINAL-017 (legacy 73A). Cross-Page Tables Must Show A Dedicated Continuation Title On Every Continuation Page (Mandatory)

- If a real thesis table spans more than one rendered page and the active template or sample uses continuation titles such as `续表 3-2`, do not leave later pages with only the carried table body.
- The continuation page must begin with a standalone continuation-title paragraph above the continued table fragment.
- Required continuation-title behavior:
  - continuation title wording follows the active sample or template family such as `续表 3-2`
  - continuation title remains outside the table grid rather than being merged into the header row
  - continuation title is centered, zero-indent, and `keepNext`
  - continuation title stays visually attached to the continued table fragment on that page
- Do not accept a continuation page that repeats the original first-page table title blindly when the approved sample requires a distinct `续表` form.
- During final QA, render every continuation page of every touched cross-page table and verify the continuation-title surface explicitly.
- Final acceptance must include table continuation evidence paths, a table continuation summary, cross-page table rendered pages, a continuation-title outside-grid verdict, and a table row split/header repeat verdict. If no body table spans pages, say that explicitly and still list the rendered table pages reviewed.

### QA-FINAL-018 (legacy 73B). Global Body Normalization Must Not Rewrite Cover Text, Table Captions, Or Table Cells Into The Body Baseline (Mandatory)

- A whole-document body-format pass must not be allowed to re-sync protected non-body surfaces into the body font, body size, body justification, or body first-line indent.
- At minimum, the protected exclusion set includes:
  - cover/title-page text lines
  - declaration / commitment front-matter rows
  - body heading level paragraphs
  - figure captions
  - table captions
  - cross-page continuation titles such as `续表`
  - table-cell paragraphs
  - tail-block opener titles such as references, acknowledgement, conclusion, and appendix
  - reference title and reference-entry paragraphs
  - acknowledgement title/body paragraphs
  - appendix title/body paragraphs
- If a pass makes cover text or table text look visually closer to the body paragraph family than to the locked sample instance for that surface, treat the pass as surface contamination rather than as successful normalization.
- Recovery order:
  - restore the polluted non-body surface from the approved local sample instance
  - keep that surface excluded from the next body-normalization pass
  - rerender the affected pages before handoff

### QA-FINAL-019 (legacy 73C). Image Holders, Figure Captions, Table Captions, Continuation Titles, And Table Cells Must Clear Body First-Line Indent Residue (Mandatory)

- Treat body first-line indent leakage into media and table surfaces as a dedicated hard-failure family.
- The following paragraph classes must not inherit the body first-line indent or body-text justification:
  - body heading level paragraphs
  - image-holder paragraphs
  - figure-caption paragraphs
  - table-caption paragraphs
  - continuation-title paragraphs such as `续表`
  - table-cell paragraphs
  - tail-block opener title paragraphs
  - reference title and reference-entry paragraphs
  - acknowledgement title/body paragraphs
  - appendix title/body paragraphs
- Required cleanup baseline:
  - first-line indent explicitly zero
  - no inherited left indent or hanging indent residue
  - alignment comes from the approved sample for that class rather than from nearby body text
- If a rendered page still shows caption offset, table-cell text shifted right, a heading or tail-block title shifted from its template center/left baseline, acknowledgement text inheriting the wrong body indent, a bibliography block with the wrong left-x/hanging geometry, or an image paragraph starting from the body-text indent column, the layout pass is incomplete.

### QA-FINAL-020 (legacy 74). Rendered-Page QA Must Be Performed On The Exact Review-Copy Path (Mandatory)

- Do not trust a successful `officecli query`, `outline`, or package inspection from one DOCX path when the user may be viewing another similarly named draft.
- For late-stage thesis QA, verify the exact review-copy path that is being handed to the user.
- When a user screenshot conflicts with the builder's internal checks, first confirm file identity by comparing one or more visible sentinel strings such as chapter titles, figure captions, or table captions from that exact path.
- A run is not accepted if the builder inspected one file while the user reviewed a different file.

### QA-FINAL-021 (legacy 75). Figure-And-Table Pagination QA Must Explicitly Detect In-Block Blank Pages (Mandatory)

- Pagination review is incomplete if it checks only whether a figure caption or table caption exists somewhere in the document.
- For every touched figure block and table block, explicitly inspect the rendered page sequence around that block and determine whether an abnormal blank page has appeared inside the block's logical neighborhood.
- Treat any of the following as release-blocking pagination failures:
- a page that is visually blank except for residual spacing caused by a figure block
- a page that contains only a stranded image paragraph while the caption or required narrative moves elsewhere
- a page that contains only a stranded table caption while the table starts on the next page
- a blank page inserted between a figure/table object and its required caption or explanatory paragraph
- Root-cause analysis must consider:
- residual `pageBreakBefore`
- oversized image-holder spacing
- keep-with-next or keep-lines misuse
- blank paragraphs left around figures or tables
- table-start pagination drift
- Do not treat abnormal blank pages near figures or tables as generic document noise. They must be tied back to a concrete edited block and repaired before handoff.

### QA-FINAL-022 (legacy 76). Table Format Is Its Own Final Acceptance Gate, Not A Secondary Cosmetic Check (Mandatory)

- Final thesis QA must verify table format as an independent acceptance surface in addition to pagination.
- Required checks include:
- three-line border geometry or the stronger template-specific border family
- caption position and caption formatting
- header-row emphasis and separator line
- internal vertical separators when required by the project rule
- absence of unintended full-grid fallback borders
- readable row spacing and cell paragraph layout
- A table page does not pass final QA merely because the caption and table start stay on the same page. The active table style itself must also match the required standard.
- If pagination repair changes table width, row height, or caption adjacency, rerun the table-format check on those pages in the same pass.
- The accepted evidence must include both DOCX table-structure inspection and rendered/PDF evidence for the top rule, header separator rule, and bottom rule; either side alone is not enough after table-style user feedback.

### QA-FINAL-023 (legacy 76A). Final Table QA Must Enumerate The Entire Table Family, Not Only Touched Tables (Mandatory)

- Do not rely on `touched pages only`, ad hoc page sampling, or generic `officecli view issues` output to conclude that thesis tables are acceptable.
- Required final-QA sequence when the manuscript contains real tables:
- enumerate every body table on the exact review-copy path
- record the total table count
- inspect the structural border family of every table
- render every table page and verify the visible border hierarchy
- record continuation status for every table, including `no cross-page body tables detected` when applicable
- Treat the following as release-blocking failures:
- one or more tables never entered the audit scope
- all table borders remain on the same Word default full-grid geometry while the project requires a thesis three-line family or a stronger sample-derived border hierarchy
- some tables were rendered and checked while sister tables in the same manuscript family were skipped
- If the run cannot prove that every real table was enumerated and reviewed, the thesis-format task is incomplete.

### QA-FINAL-024 (legacy 77). Thesis Automation Must Lock One Verified Python Interpreter Path Per Run (Mandatory)

- If the default `python` or `py` launcher is unstable, missing, or resolves to the wrong environment, do not keep using launcher aliases opportunistically across the same thesis run.
- Recovery order:
  - probe installed interpreter paths explicitly
  - verify required packages for the current subtask on each candidate
  - choose one working interpreter path
  - use that absolute interpreter path consistently for every helper script in the current run
- Do not treat "Python exists" as sufficient. The chosen interpreter must be validated against the actual package set required by the current workflow, such as `python-docx`, `selenium`, `pymysql`, or thesis-specific document tooling.
- If different subtasks require different interpreters, record that split explicitly; do not silently mix launchers.

### QA-FINAL-025 (legacy 78). Local Runtime Screenshot QA Must Detect Port Collisions Before Blaming Routes Or Templates (Mandatory)

- When a thesis run requires screenshots from a local web system, do not assume that `127.0.0.1:5000` or another default port belongs to the current project.
- If live browser access returns `404`, an unexpected login page, or the wrong app surface while offline route inspection still says the route exists, first suspect port collision or wrong-process binding.
- Required recovery sequence:
  - inspect the listening process on the target port
  - compare it with the process that was just started for the current project
  - if another process is already bound, move the review server to a unique port and re-verify
  - only after port identity is confirmed may the builder classify the incident as a route or template failure
- For screenshot capture, the runtime base URL must be the verified project-owned URL, not merely a convenient localhost default.

### QA-FINAL-026 (legacy 79). Paragraph-Clone Content Passes Must Be Followed By Body-Format Re-Normalization And Render Review (Mandatory)

- After any thesis content pass that clones paragraphs, duplicates formatted blocks, or inserts new explanatory paragraphs through `python-docx` or equivalent DOM editing, do not assume the inserted body text inherited the correct final formatting.
- Common failure signs:
  - newly inserted body paragraphs inherit run-level bolding
  - image explanation paragraphs inherit caption formatting
  - newly inserted body paragraphs retain wrong spacing, indent, or pagination behavior
  - newly inserted body paragraphs fall back to single-line or `1x` line spacing while neighboring body paragraphs still use the approved academic body baseline
  - an inserted or rewritten body zone keeps one or more empty `Normal`/body paragraphs between two visible body paragraphs after the content pass
- Required recovery sequence:
  - normalize the affected zone's body paragraphs back to the approved body baseline
  - delete residual empty body paragraphs that were left behind by content replacement or paragraph insertion helpers unless the template explicitly requires a visual blank line there
  - normalize captions separately as caption paragraphs rather than as body text
  - re-check heading paragraphs that border the edited zone
  - rerender the touched pages before handoff
- If a content pass changed text correctly but left the body visually denser, bolder, or structurally inconsistent than the template, the pass is incomplete.

Render Review Page-Number Distinction Detail

- When a thesis uses cover/front-matter pages, roman-numeral prelim pages, or restarted arabic numbering, do not assume a Word-reported page number maps directly to the exported PDF page index.
- Before page-targeted rendered review, explicitly determine whether the selected page number is:
  - a visible thesis page number
  - or a physical page position in the exported PDF
- If the run finds a chapter title or tail-block title through Word or DOCX inspection and then exports the thesis to PDF, verify the physical rendered page by sentinel text rather than relying only on `page N => PDF[N-1]`.
- A render-review pass fails if the inspected PNG/PDF page is not first confirmed to contain the expected local sentinel text for that exact review target.

### QA-FINAL-027 (legacy 94). Abstract And Keyword Rewrites Must Preserve The Original Front-Matter Run Structure (Mandatory)

- When rewriting thesis abstract, keyword, English abstract, or English keyword paragraphs inside an existing `.docx`, do not collapse the whole paragraph into one newly generated run by default.
- Preserve the front-matter split between:
  - label run such as `摘要：`, `关键词：`, `Abstract:`, `Key words:`
  - following body-text run(s) that use the template's normal abstract-body formatting
- A content edit that keeps the wording but turns the whole abstract or keyword paragraph into title-sized or fully bold text is a failed abstract repair.
- If the active tool cannot preserve the original run-level formatting safely, switch to a structure-preserving path such as:
  - restoring the paragraph from the accepted baseline and then replacing only the body-text run
  - XML-level patching limited to `word/document.xml`
- Treat abstract and keyword formatting drift after a content pass as a release-blocking failure, not as a cosmetic issue.

### QA-FINAL-028 (legacy 95). Content-Only DOCX Rewrites Must Prefer Run-Preserving Or Document-Only Patching On Pagination-Sensitive Manuscripts (Mandatory)

- When editing an existing thesis manuscript for content only, do not use a broad paragraph-rebuild path if that path may rewrite package parts beyond `word/document.xml`.

### QA-FINAL-029 (legacy 102). Table QA Cannot Rely On DOM Properties When The User Is Complaining About Visual Breakage (Mandatory)

- If the user reports that table text is still indented, clipped, wrapped strangely, or visually unreadable, do not accept `officecli get`, table XML, or paragraph-property inspection as sufficient proof of correctness.
- Required recovery sequence:
- render the exact review-copy pages that contain the complained-about tables
- inspect those rendered pages with machine vision
- identify whether the failure is caused by table width, column width, font size, first-line indent residue, line-spacing residue, or partial off-page clipping
- repair the table
- rerender the same pages and recheck them before handoff
- The final evidence must name the rendered page paths and the table-structure audit path; a prose claim that the table is now a three-line table is not enough.
- If the rendered page still shows split decimals, detached percent signs, fragmented headers, or clipped mean/rank columns, the table-format task is still failing even if the underlying DOCX properties look normalized.

### QA-FINAL-030 (legacy 103). Multi-Column Satisfaction Tables Must Pass A Page-Fit Review As A Whole Family (Mandatory)

- For thesis satisfaction-analysis tables with many narrow numeric columns, treat page-fit as a dedicated acceptance gate.
- Unacceptable rendered-page symptoms include:
- decimals or percent signs wrapping onto a separate line
- mean or ranking values breaking into stacked fragments
- header text fragmenting into visually chaotic vertical stacks
- the rightmost columns drifting into the page edge or clipping outside the visible table block
- one table in the same four-table family remaining visibly denser or narrower than the others after a repair pass
- Required repair behavior:
- measure or estimate the table against the real text block width of the page, not against arbitrary default twip values alone
- normalize the full table family consistently when they share the same structure, instead of patching only one table and assuming the others will be fine
- rerender every touched table page in that family after the pass
- Do not stop after a partial improvement on one table while sister tables in the same family still fail visually.
- On pagination-sensitive manuscripts, the default safe order is:
  1. preserve the current accepted manuscript as the baseline
  2. apply the smallest possible text rewrite
  3. prefer run-preserving edits first
  4. if that fails, patch only `word/document.xml`
  5. avoid unexpected rewrites to `styles`, `theme`, `header`, `footer`, `numbering`, or `settings`
- If a content-only pass unexpectedly changes package parts outside `word/document.xml`, treat the pass as unsafe until the drift is explained and intentionally accepted.
- If the user reports that pagination disappeared or chapter/front-matter layout shifted after a wording-only pass, first suspect paragraph/run reconstruction or multi-part DOCX rewrites before blaming text-length change alone.

### QA-FINAL-031 (legacy 96). Any Check, Review, Or Acceptance Pass Must Include Machine-Vision Page Inspection (Mandatory)

- Under this skill, do not treat structure-only inspection as sufficient for any task that is framed as checking, reviewing, verifying, validating, auditing, or accepting an output.
- Required rule:
  - if the run involves checking a thesis, document, screenshot, page layout, figure placement, pagination, header/footer state, cover, abstract, TOC, or other visual deliverable
  - then the run must include a machine-vision inspection step on the actual rendered or user-facing pages in addition to structural inspection
- Structural inspection includes tools such as:
  - `officecli outline`
  - DOCX XML inspection
  - paragraph/style queries
  - page-number extraction
  - package-part validation
- Machine-vision inspection must be used to confirm at least:
  - visible page composition
  - abnormal blank space
  - figure visibility
  - caption placement
  - TOC visual integrity
  - header/footer visible state
  - whether the checked file really matches what the user sees

### QA-FINAL-032 (legacy 97). Citation QA Must Include A Body-Only Coverage Audit And A Broken-Sentence Audit (Mandatory)

- A thesis citation QA pass is incomplete if it checks only that citation markers exist somewhere in the document.
- Required minimum audit items:
  - every bibliography item has at least one citation in the thesis body unless the user explicitly accepts surplus references
  - no body sentence begins with a citation marker such as `[n]指出` or `[n]表明`
  - no range-style citation like `[20]至[25]` remains in body prose when the active citation rules require one source per marker
- During this audit, exclude:
  - TOC lines
  - front matter
  - figure captions
  - table captions
  - bibliography paragraphs themselves
- If the audit scope is not body-only, the citation QA result is not reliable enough for acceptance.
- If a run reports a check result without machine-vision review on the relevant pages, treat that check as incomplete rather than "passed".
- A report that cites only HTML previews, DOM extraction, XML inspection, or `officecli` structure output without rendered PDF page images must be treated as failing this rule.

- Final acceptance must store the exact body-citation audit report path for the exact review-copy or final-deliverable DOCX that was checked.

### QA-FINAL-033 (legacy 98). Content-Only Thesis Passes Must Run A Format-Regression Audit (Mandatory)

- After any thesis pass that claims to change wording only, run a regression audit against the accepted pre-edit baseline.
- Required audit surfaces:
  - one unchanged citation-bearing paragraph
  - one rewritten citation-bearing paragraph when citations were touched
  - one untouched table block when tables exist
  - one rewritten body paragraph
  - one nearby heading block
- Required audit items:
  - first-line indent
  - line spacing
  - effective font mapping
  - superscript citation preservation
  - table caption and table border stability
- If any unchanged surface drifts during a wording-only pass, the pass is unsafe and must be rebuilt from the accepted baseline.

### QA-FINAL-034 (legacy 99). Package Drift Outside `word/document.xml` Is A Release-Blocking Signal For Content-Only Passes (Mandatory)

- On pagination-sensitive thesis manuscripts, a content-only pass should not silently rewrite package parts such as:
  - `word/styles.xml`
  - `word/settings.xml`
  - `word/fontTable.xml`
  - `word/numbering.xml`
  - `word/header*.xml`
  - `word/footer*.xml`
  - `word/_rels/document.xml.rels`
  - `[Content_Types].xml`
- If those parts drift during a pass that was not explicitly changing them, do not hand off the result as a minor wording update.
- First classify the drift:
  - intentional structural repair
  - office-application rewrite side effect
  - unsafe paragraph / run reconstruction
- If the drift is not explicitly required by the task, reject the pass and restart from the accepted baseline using a narrower mutation path.

### QA-FINAL-035 (legacy 100). Touched Thesis Review Surfaces Must Have Review Evidence Records Before Handoff (Mandatory)

- If a thesis run modifies body paragraphs, headings, TOC entries, figure blocks, table blocks, captions, front matter, or pagination, do not hand off the result unless the corresponding review surfaces have review evidence records.
- The acceptance record must reference those review evidence record files rather than only raw screenshots or a prose statement that review was performed.
- Minimum rule:
- paragraph-level thesis edits require paragraph-review evidence
- touched rendered pages require rendered-page evidence
- touched figures require figure-review evidence
- If the run changed a thesis surface but the acceptance record cannot point to a valid review evidence record for that surface, treat the run as incomplete even if the visible document currently looks correct.

### QA-FINAL-036 (legacy 101). Thesis Modification Runs Must Include Blast-Radius Touched-Page Review Before Handoff (Mandatory)

- If a thesis run modifies a paragraph, caption, figure block, table block, heading, TOC entry, page break, spacing rule, or any other pagination-sensitive surface, do not stop at the immediate local paragraph review.
- Before handoff, run and record a touched-page review that covers the modified page neighborhood, including:
  - the modified page
- adjacent pages when page flow may have shifted
- any extra pages visibly affected by figure, table, TOC, or blank-space drift
- The acceptance record must reference a touched-page review evidence record for this blast-radius review.
- If a builder edits one place and other pages can still drift without a recorded touched-page review, the thesis-format run is incomplete even if the immediate local paragraph check passed.

### QA-FINAL-037 (legacy 104). Figure Readability Fixes Must Not Default To Changing The Manuscript Page Layout (Mandatory)

- If a thesis figure is readable in the standalone image but unreadable after insertion, first classify that as a figure-source readability failure rather than as a page-layout failure.
- Required repair order:
  - increase figure-internal font size
  - enlarge the figure's internal nodes, attributes, or label containers
  - reduce internal density and rebalance the layout
  - then rerender the same thesis page at the intended insertion width
- Do not default to increasing paper size, changing page orientation, widening the text block, or otherwise altering the manuscript page layout just to rescue a small or dense structural figure.
- Page-layout changes for figure readability require a higher-precedence reason such as an explicit user instruction or a template-backed requirement.

### QA-FINAL-038 (legacy 105). Figure QA Must Explicitly Reject Connectors That Cross Text Or Shape Borders (Mandatory)

- Treat connector-to-text and connector-to-border collisions as their own release-blocking review class rather than as a vague subset of general overlap.
- Required rendered-page checks for every touched structural figure:
  - no connector passes through the visible text area of any label
  - no connector passes through an unrelated rectangle, ellipse, diamond, or other shape border
  - no connector enters a target shape by cutting across its label area instead of reaching the boundary cleanly
- If the user reports that a line is crossing text or a shape border, do not answer with source-diagram inspection alone.
- Rerender the exact thesis page, inspect the collision on the rendered page, redraw the source figure, and rerender again before handoff.

### QA-FINAL-039 (legacy 105A). Final Submission Claims Must Be Bound To The Exact Final Deliverable And A Full Sequential Render Review (Mandatory)

- Do not allow a final `可提交`, `ready to submit`, or equivalent claim unless all of the following are true:
  - the exact final named deliverable DOCX path is locked
  - the exact final named deliverable PDF path is locked when PDF export is required
  - the verdict is based on that exact final artifact rather than an earlier review copy or intermediate draft
  - rendered-page review was performed sequentially from the cover page through the final end-matter page
- A run fails this gate if it relied on:
  - structure-only inspection
  - outline or XML checks alone
  - sampled page review only
  - a similarly named but non-final artifact
- If cover identity fields or other protected front-matter placeholders are still unresolved, the artifact may be called `near-final` only when the user has not yet waived those placeholders. It must not be called submit-ready by default.

### QA-FINAL-040 (legacy 106). Final QA Must Record A Sample-Comparison Pass On The Exact Deliverable (Mandatory)

- Final thesis QA is incomplete if it only lists that rendered pages were reviewed.
- When a local sample thesis exists, final review must also record which exact rendered pages of the final deliverable were compared against which sample page classes.
- At minimum, record the comparison for:
  - cover
  - Chinese abstract
  - English abstract
  - TOC
  - first body page
  - one figure page
  - one table page
  - references
  - acknowledgement
- If the user reports that the delivered manuscript still looks far from the sample, missing sample-comparison evidence is a process failure rather than a cosmetic miss.

### QA-FINAL-041 (legacy 107). Visible Citation Anchor Leakage Must Hard-Fail Final QA (Mandatory)

- If rendered body pages show internal bookmark names, helper text, or field-display artifacts such as `cite_ref_1`, `bookmark_3`, `HYPERLINK`, or similar internal identifiers near citation markers, the citation task has failed.
- Do not accept the manuscript merely because the DOCX package contains hyperlinks or bookmarks.
- The accepted visible result is that only the citation marker itself appears to the reader.

### QA-FINAL-042 (legacy 108). Runtime Screenshot Family Must Be Route-Verified During Final QA (Mandatory)

- For chapter-4 or equivalent implementation screenshots, final QA must verify more than image readability.
- Required acceptance check per runtime screenshot slot:
  - caption text matches the intended page or module
  - screenshot content matches the intended route or page state
  - no structural diagram appears in a runtime screenshot slot
  - no duplicated screenshot asset is silently reused for two different runtime captions unless explicitly approved
- If a route-based screenshot family was refreshed, record the route-to-caption mapping in the final QA artifacts.

### QA-FINAL-043 (legacy 109). TOC Visual Baseline Restoration Is A Final Acceptance Surface (Mandatory)

- Do not accept a TOC repair merely because the heading list and page numbers are correct.
- Final TOC QA must verify all of the following on rendered TOC pages:
  - TOC title matches the locked baseline in font, size, alignment, and spacing
  - TOC level-1, level-2, and level-3 entries match the locked indentation and line-spacing pattern
  - dotted leaders and the right-aligned page-number column match the locked baseline
  - the TOC page count or page-occupancy rhythm has not collapsed into a denser default-style layout unless the added entries truly require a different count
- A run fails this gate if the refreshed TOC still looks like a default application TOC even when the visible text is correct.

### QA-FINAL-044 (legacy 109A). Tail-Block First-Page Pagination Is A Final Acceptance Surface (Mandatory)

- Final QA must treat touched tail-block openers such as `结论`, `参考文献`, `致谢`, and `附录` as independent acceptance surfaces.
- Required checks on the exact rendered deliverable:
  - the title paragraph matches the locked baseline
  - the opener starts on a fresh page
  - the `references` opener page is strictly after the previous real content page, with `references previous content physical page=...` and `references_prior_block_separation_verdict=pass` recorded in the tail-block opener map
  - the page immediately before the opener does not still carry a merged boundary
  - references and acknowledgement remain separate terminal blocks in the rendered sequence
  - the acceptance record names the evidence files and summary for the checked tail-block openers
- A run fails this gate if other pagination looks correct but a tail-block opener lost its verified page-start owner or merged into the prior page.

### QA-FINAL-045 (legacy 110). Final QA Must Run A DOCX Font-Name And Mojibake Audit On The Exact Deliverable (Mandatory)

- Do not rely only on rendered-page appearance when helper scripts, XML patches, or PowerShell automation touched font-bearing or text-bearing thesis surfaces.
- Run a DOCX-internal audit on the exact final deliverable or exact review copy for:
  - visible `w:t` text on touched surfaces
  - `w:rFonts` and related font-family attributes in `document.xml`, `styles.xml`, `fontTable.xml`, and touched header/footer parts
- Hard-failure examples include:
  - corrupted East Asian font names such as mojibake in `w:rFonts`
  - WPS or Word showing unreadable font labels on references, acknowledgement, conclusion, captions, TOC, or headings
  - title text that appears only through fallback while the underlying DOCX font binding is broken
- Treat this as a tooling failure, not as an acceptable fallback.

### QA-FINAL-046 (legacy 111). Final QA Must Include Surface-Face Parity For Every Touched Formatting Class (Mandatory)

- Final QA must not accept a thesis-format pass based only on `officecli validate`, outline shape, page count, no-overflow observations, or visible text correctness.
- For every touched or user-reported formatting class, record a surface-face parity verdict against the locked template/sample baseline.
- Surface-face parity must be class-specific, not only document-wide or body-style-wide. The final QA lane must verify separate class bindings for every heading level, body text, figure holder/body, figure caption, table title, table-cell text, abstract title/body, keyword label/content, reference title, reference entry, citation superscript, acknowledgement title/content, appendix title/content, header, footer, and page number whenever that class is present or user-reported.
- A final record that says references, acknowledgement, appendix, abstract, figure captions, table titles, or keywords were checked while those paragraphs are still bound to body/Normal style without a matching template donor is a failed acceptance record.
- For full-thesis or user-reported format repair, the surface-face parity evidence must be a class matrix, not prose. It must carry one row per present or reported surface with: surface id, template donor id/path, expected style id/name, final style id/name, paragraph metrics verdict, run-font verdict, rendered-page evidence path, and final verdict.
- A single row such as `body passed`, `all surfaces checked`, or `format checked` is not valid evidence for references, acknowledgement, appendix, abstract, keyword, TOC, figure, table, or citation surfaces.
- The mandatory thesis surface inventory is required for every thesis generation, thesis revision, and thesis format repair run. It must include cover, declaration/title front matter, abstract, keyword, TOC, references, acknowledgement, appendix, header, footer, and page-number rows with statuses, reasons, evidence paths, and final verdicts.
- The parity verdict must explicitly cover:
  - style binding and outline/list state
  - paragraph alignment, indentation, spacing, line spacing, tabs, keep rules, and page-break ownership
  - Chinese, Western/English, and complex-script font mappings
  - font size, bold, italic, underline, color, highlight, superscript/subscript, and label/content run splits
  - field, hyperlink, bookmark, page-number, TOC leader, formula-number, citation-marker, table-border, or image-holder behavior when present
- The exact review copy or final deliverable must pass both:
  - a DOCX-internal audit of those properties
  - a rendered-page machine-vision review of the touched page neighborhood
- If any property is unknown because the baseline was not extracted, the final QA verdict for that surface is blocked rather than passed.
- If a repair uses a builder default font, application default font, or guessed English font where a template baseline should exist, the final QA verdict is failed even when the rendered page looks close.

### QA-FINAL-047. Sample Self-Check Failures Block Delivery Claims (Mandatory)

- Do not describe a thesis sample, rebuilt manuscript, or generated draft as successful when `sample_self_check` reports any critical subcheck as failed, `未通过`, `blocked`, or equivalent.
- The final QA interpretation of `sample_self_check` must inspect the report body, not only the process exit code.
- Critical subchecks include abstract-surface parity, TOC content and visible format, body-style binding, table caption binding, table-cell baseline, figure-block locality, image-holder safety, citation audit, and font/encoding audit.
- The required detector set includes `figure.scope-manifest-contract`; when figure or media surfaces exist, the report and acceptance side must validate it with source-to-final evidence, including a `source_docx` binding rather than final-only DOCX evidence.
- A smoke fixture may use explicit `--smoke-acceptance` only to keep detector tests running; the corresponding report must be marked as smoke-only and blocked for delivery.
- If a generated sample is handed to the user while a critical subcheck is blocked, record that handoff as a workflow failure and repair the skill gate before generating the next sample.

### QA-FINAL-048. Smoke And Placeholder Manuscripts Cannot Become Full Thesis Evidence (Mandatory)

- A manuscript containing smoke, placeholder, minimal/manual-review, detector-fixture, `Example Research`, template-alignment note, self-check, or skill-gate wording must fail the full-thesis delivery gate.
- `--smoke-acceptance` is allowed only for detector fixtures. It must force the self-check and acceptance record into non-deliverable status even when individual detector dimensions pass.
- Final acceptance must fail if the self-check report says `smoke acceptance mode: yes`, `smoke-only; blocked for delivery`, or `full thesis content gate failed`.
- Do not use `complete_sample_smoke` or any shortest/manual-review fixture as evidence that the skill can write a new complete thesis.
- A real new-thesis test must show cover preservation, Chinese abstract, English abstract, TOC, body chapters, figures, tables, references, acknowledgement, rendered page evidence, and content-completeness checks on the exact final DOCX/PDF.

### QA-FINAL-049. Local Surface Passes Cannot Become Whole-Thesis Passes (Mandatory)

- Every pass verdict must declare scope: local surface, specialty repair, or whole thesis.
- A local reference-entry/font repair pass cannot clear whole-thesis delivery while body-style, pagination, protected-surface, figure, table, citation, or surface-inventory gates still have blockers.
- After any rule tightening, prior pass reports become stale under the new rule until the exact same final DOCX SHA256 is re-audited with the new validator.
- The final handoff wording must be generated from the whole-thesis gate result. It must not be inferred manually from a single local audit file.
- The validation command in a final acceptance record must run `validate_skill_gate.py` or `validate_skill_gate.cmd` with `--gate-record <that exact record>`.

### QA-FINAL-050. Repeated User-Reported Thesis Defects Need Named Detector Closure (Mandatory)

- When the user reports a concrete surface defect, the next thesis delivery cannot pass from a generic `format checked` or local visual claim.
- The acceptance record and `sample_self_check` report must name the exact detector family that guards the defect, including:
  - cover donor/layout drift and cover-next-page pagination
  - front-matter page separation, especially Chinese abstract to English abstract blank-page or same-page drift
  - TOC title color/font pollution and missing Chinese/English abstract entries
  - TOC underline pollution, including underlined title text, underlined entries, underlined leader/page-number runs, or accidental hyperlink/visited-hyperlink style inheritance
  - invisible or clipped figures, including SVG fallback leakage and raster relationship failure
  - acknowledgement title/body typography drift
  - abstract body paragraphs containing manual `w:br` line breaks instead of real Word paragraphs
  - abstract body indentation loss, including zeroed first-line indentation or manual-space substitutes
  - keyword line run contamination where keyword content inherits the label/title bold face or heading class
  - second-level and third-level heading style drift, including unfamiliar font families, body-format residue, or lost heading level evidence
  - whole-document blank-page or near-empty-page regressions caused by hard page breaks, section breaks, field refresh, or page-detection mistakes
  - table title loss, table-title/table-body misbinding, table-title text trapped in a cell, or a real third-level heading misclassified as a table title
  - heading paragraphs carrying abnormal left, first-line, hanging, tab, or rendered left-x indentation against the locked template level baseline
  - tail-block opener titles, including references, acknowledgement, conclusion, and appendix, carrying body-indent residue or rendered centerline/left-x drift
  - acknowledgement title/body indentation, paragraph-dialog metrics, font color, or rendered text-box geometry drift
  - reference-title paragraph indentation, alignment, paragraph-dialog metrics, font color, or rendered centerline/left-x geometry drift
  - reference-entry paragraph indentation and rendered left-x geometry drift
  - reference-entry font family, named size, direct-run font slot, or style-class pollution
  - English keyword label/content splitting where the `Key Words:` colon must remain part of the label
  - runtime screenshot, code screenshot, algorithm detection, OCR, YOLOv8, DBNet, CRNN, or recognition-result figures being replaced by synthetic schematics, mockups, placeholders, or unverifiable sample images
  - template-instruction leakage where the final DOCX/PDF still shows font/size/alignment notes, `论文、设计二选一`, `须删除本行`, signature/date instructions, or other wording that only describes the template instead of thesis content
  - visible red template-format pollution, including direct red runs, red inherited from paragraph/run styles, red highlights, red underline color, and red text inside header/footer/footnote/endnote parts
- A repair that touches `Normal`, `正文`, docDefaults, theme fonts, default paragraph spacing, TOC styles, table styles, caption styles, or any broad run/paragraph style family has style-blast-radius risk. Before handoff, final acceptance must record a protected-surface freeze manifest, a post-mutation diff, and pass verdicts for at least TOC visible-run typography, TOC underline pollution, table authority, table local structure, table rendered family, body-style binding, surface-face parity, sibling surfaces, references, acknowledgement, appendix, and whole-document pagination.
- Repeated-defect or style-blast-radius closure must also carry a valid thesis mutation transaction record from `references/thesis/thesis-mutation-transaction.md`. The record must bind the freeze manifest, post-mutation diff, target render, blast-radius render, cross-surface regression report, target surface ids, write owner, final DOCX path, and final DOCX SHA256 into one validated transaction. Separate detector pass snippets do not satisfy this rule unless the transaction validator passes.
- A local body-style or paragraph-spacing repair cannot pass with TOC/table fields marked `not-applicable`. Changing body defaults must automatically escalate TOC and table verification even when the user only asked about body text.
- If any protected surface that was not supposed to change differs between the pre-mutation freeze and final DOCX, the run is failed unless the user explicitly requested that surface change and it has its own detector evidence.
- Repeated surface-defect closure is an all-mentioned-surfaces gate. If one user feedback item names abstract/keyword, heading, pagination, figure/screenshot/provenance, table, and references together, the user-reported issue ledger must include each named surface family and must map it to a real detector/evidence path. A ledger that covers only one or two of the named surfaces is a failed closure.
- Template donors must be classified before use: reusable donor data is limited to layout, paragraph properties, typography, style bindings, table/caption geometry, header/footer structure, and field containers; visible explanatory words in the template are non-copyable instruction artifacts and must be stripped from every DOCX text-bearing part before final acceptance.
- A pass result is valid only when the exact post-edit DOCX/PDF paths and detector outputs are recorded for those named issues.
- If an issue was previously reported in the current project history and the final detector is missing, stale, not-applicable, smoke-only, or scoped to a different artifact, the gate must fail.

### QA-FINAL-051. Heading And End-Matter Indentation Drift Needs Surface-Bound Evidence (Mandatory)

- When the user reports abnormal indentation, offset titles, shifted headings, reference indentation, acknowledgement layout, body-style pollution, or a fix-one-break-another formatting regression, the next handoff must treat indentation as its own detector family rather than a side effect of body-style or pagination checks.
- The detector family must cover the exact protected surface ids that are present or user-reported:
  - `body_heading_levels`
  - `references_title`
  - `references_entries`
  - `acknowledgement_title`
  - `acknowledgement_body`
  - `appendix_title`
  - `appendix_body`
  - `whole_document_pagination` for opener-page ownership
- Each surface row must have its own evidence path and pass verdict. A single `tail block passed`, `reference format passed`, `heading style passed`, or `body style passed` row cannot prove these surfaces.
- The evidence must expose paragraph-dialog metrics for template and actual values: alignment, left/right/first-line/hanging indent, tab stops, spacing before/after, line-spacing mode/value, keep/list/page-break ownership, style id/name, font slots, font size, bold/underline/color, and body-format residue state.
- The rendered comparison must expose numeric template-vs-target geometry for title centerline or left-x, text bounding box, line-height/y-delta, page-region occupancy, and bibliography hanging/left-x position when applicable.
- Final acceptance must include dedicated detector fields for `tail-block title indentation`, `references_title`, `references_entries`, `acknowledgement_title`, and `acknowledgement_body`. Missing, stale, not-applicable, sampled-only, manual-only, or generic visual evidence fails the gate.

### QA-FINAL-052. User-Named Catastrophic Format Failures Require One Ledger Row Per Surface (Mandatory)

- When the user reports a catastrophic thesis format failure list, the next handoff must treat each named item as an independent release-blocking surface, not as a generic `format checked` claim.
- The user-reported issue ledger must include one row for every named surface in the user's wording, including:
  - cover damaged or cover layout drift
  - header wrong or missing
  - footer wrong, page-number wrong, or page-number field loss
  - Chinese abstract, English abstract, keyword, or abstract translation mismatch
  - abnormal blank pages, near-empty pages, or duplicate page breaks
  - citation superscripts lost or citation markers converted to baseline text
  - reference content, numbering, GB/T format, font, indentation, or order errors
  - body font, body size, body style binding, or substituted font errors
- Each row must bind to a surface-specific evidence path and final verdict. A shared generic PDF, office validation result, page count, or prose statement cannot satisfy multiple rows unless the evidence file itself has one explicit row per surface id.
- If any named surface lacks an evidence path, a baseline source, a final DOCX SHA256 binding, or a pass verdict, the task remains incomplete.
- The handoff must not say `已完成`, `1:1`, `可提交`, `格式已修复`, or equivalent completion wording while any named surface row is missing, stale, sampled-only, or failed.

### QA-FINAL-053. Full-Document Render Review Is Required After Whole-Thesis Format Repair (Mandatory)

- For whole-thesis revision, full-paper formatting, or `1:1` template-alignment work, do not sample only selected pages.
- Final QA must create a sequential rendered-page review from cover through the last end-matter page, with page-class labels for at least cover, declaration/title front matter, Chinese abstract, English abstract, TOC, each chapter opener, at least one normal body page, every figure/table page that changed, references, acknowledgement, appendix when present, and final page.
- Blank-page and near-empty-page detection must cover the whole rendered document. It is not enough to review only front matter or only the pages that changed.
- Header/footer/page-number review must include rendered evidence from front matter, TOC, first body page, an odd body page, an even body page, references, acknowledgement, and appendix when present.
- If a rendered full-document pass is unavailable because Word/WPS/LibreOffice/PDF rasterization is blocked, the run may produce an audit report but must not claim final 1:1 visual alignment.
- A final review that uses only DOCX XML, office validation, PDF export success, screenshots of a few pages, or page-image existence is not final QA for this class of task.

### QA-FINAL-054. Official Template Rule Profile Must Be Cited In Final QA (Mandatory)

- When an official school format document or format template is provided, final QA must cite the extracted template rule profile and the active template profile used for mutation.
- The acceptance record must name evidence for these profile rows:
  - cover title and identity field requirements
  - abstract and keyword label/content requirements
  - body Chinese and Western font requirements
  - odd/even header requirements
  - footer and page-number field requirements
  - citation superscript or annotation marker requirements
  - reference title and entry requirements
  - document order and page-break/section-break requirements
- The final QA verdict must compare the final DOCX against the extracted rule profile and the rendered template/sample, not against the damaged current draft or a prior failed repair output.
- If the template rule profile contradicts an older skill memory, the official current project template and school document win. The final QA evidence must record that override rather than silently following memory.
- A final acceptance record that lacks the template rule profile path, active template profile path, or source document fingerprint for the school rule/template cannot support a template-alignment handoff.

### QA-FINAL-055. Final Comment-Completion Claims Need Comment-Resolution Audit (Mandatory)

- A final acceptance record may not claim `all-comments-resolved`, `verified comments resolved`, `批注已改完`, `所有批注已解决`, or equivalent wording unless it names a comment-resolution ledger and a comment-resolution audit report for the exact final DOCX.
- The audit must reopen the final DOCX and, when available, the source DOCX. It must compare comment ids, text digests, done states, anchors, ledger status, evidence paths, and final DOCX SHA256.
- If a final DOCX has any done comments, the acceptance record must name the ledger authorizing those done states even when the handoff does not explicitly say all comments are resolved.
- If any comment remains open in the final DOCX, or remains partial, blocked, unknown, orphaned, missing from the ledger, or closed from evidence that does not match the comment's subissues, the handoff status must be blocked rather than pass. A fixed ledger row cannot override a Word/WPS open-comment state for final comment-completion claims.
- Comment-resolution evidence generated before the last DOCX mutation is stale and cannot support final QA.

### QA-FINAL-056. Common Pre-Submission Thesis Checklist Must Pass (Mandatory)

- Before any whole-thesis handoff, run a common pre-submission checklist on the exact final DOCX/PDF and record the evidence path.
- The checklist must explicitly verify:
  - at least 20 bibliography entries are present
  - every bibliography entry has at least one matching in-text citation marker in the body
  - figures have standalone titles below the figure, and tables have standalone titles above the table
  - figure/table internal text is one Chinese named size smaller than body text where editable text exists; for this project that means 五号 when正文 uses 小四
  - tables are editable Word tables, not table screenshots
  - ordinary diagrams/charts are not pasted screenshots used to bypass proper figure generation; runtime, code, algorithm detection, OCR, YOLOv8, DBNet, CRNN, or recognition-result screenshots/images are allowed only when the figure lane records real-source provenance
  -正文 does not contain blank lines or large unexplained whitespace blocks
  - table title and first table rows remain on the same rendered page unless an approved continuation-title rule applies
  - figure image and figure title remain on the same rendered page
  - visible manuscript length is at least 15000 Chinese-character/word units under the active school rule
  - the TOC starts from Chinese abstract/摘要 and includes English abstract/Abstract before the body chapters
  - `总结与展望` is an independent final body chapter before references/acknowledgement/appendix, appears in the TOC as a chapter-level entry, and is not reduced to a subsection named `结论` or mixed into the testing chapter
- If any item is missing, stale, sampled-only, or justified only by a manual visual claim, the final handoff is blocked.
- If this checklist conflicts with a more specific project instruction, record the override in the checklist.
- Project-specific figure/result-image provenance overrides must be recorded in the checklist rather than silently weakening this common gate.

### QA-FINAL-057. User-Reported Protected Visual Defects Require Template-Bound Render Geometry (Mandatory)

- When the user report, screenshot, or review comment points to TOC visual drift, abstract/keyword typography drift, header horizontal-line drift, footer/page-number position drift, references pagination/page-break drift, or visible body-font/body-size/body-style abnormality, do not close the issue from XML inspection, style names, `officecli validate`, PDF export success, page-image existence, or sampled-page review alone.
- The acceptance record must bind every named surface to:
  - the exact reviewed DOCX/PDF path and final DOCX SHA256
  - the active template or approved sample path and fingerprint
  - rendered template-vs-target full-page evidence for the same logical/physical page class
  - rendered key-surface crop or region evidence for the complained-about TOC row/level, abstract title/body/keyword block, header line/text region, footer/page-number region, references opener/entry block, or body-font page
  - numeric geometry or typography evidence, including left-x/centerline, bounding box, line-height/y-delta, page occupancy, font size, and paragraph/run typography where relevant
  - one ledger row per user-named surface, with detector id, evidence path, and pass verdict
- TOC issues must cover all rendered TOC pages and every used level, including dotted leaders, page-number column, row density, visible-run typography, and page-class occupancy.
- Abstract issues must cover Chinese abstract title/body/keyword line and English abstract title/body/keyword line as separate surfaces. A body-style pass or front-matter order pass cannot substitute for those six typography/paragraph checks.
- Header/footer/page-number issues must cover rendered front matter, TOC, first body page, representative body page, references, acknowledgement, and appendix when present. A header/footer XML or field-code pass is not enough when a visible line, stale header text, or page-number position is reported.
- References pagination issues must cover the references title/opener, first references page, continuation pages when present, the page immediately before the opener, and the pagination owner that makes the opener start on a fresh page.
- Body visible-font issues must cover the user-reported page and enough same-family body pages to prove the repair is not a sampled-only local patch. If body defaults, `Normal`, docDefaults, theme fonts, or shared body styles were touched, the proof must escalate to same-family chapter/body-page coverage.
- Evidence that is stale, sampled-only, current-draft-as-baseline, XML-only, structure-only, or missing either the full-page binding or key-surface binding is failed evidence even when individual detector summaries say `pass`.
- Final handoff text may not say `passed with limitations`, `core checks passed`, `structural pass`, or equivalent caveat wording for these surfaces. The correct state is blocked until the template-bound rendered evidence passes.

### QA-FINAL-058. Thesis Content Expansion Requires Post-Mutation Machine-Vision Body Review (Mandatory)

- After any thesis DOCX text mutation that inserts, expands, rewrites, or lengthens body paragraphs, final QA must render or rasterize the exact output being handed off and run a machine-vision or equivalent page-image review on every touched page plus the page-flow/blast-radius neighborhood.
- PDF export success, page count stability, XML/package diff counts, media counts, paragraph counts, style names, body character counts, and broad page-image existence are supporting evidence only. They cannot close content expansion or body paragraph insertion.
- The review must explicitly check that inserted body prose did not render as a title, chapter heading, section heading, centered opener, bold oversized block, heading-outline paragraph, or top-of-page orphan block attached to the next chapter opener.
- The acceptance record must include `content mutation rendered-page review path`, `content mutation machine-vision verdict`, `inserted body heading-contamination verdict`, `touched-page/blast-radius machine-vision evidence paths`, and `format lane post-mutation rendered audit verdict`.
- The `content mutation machine-vision verdict` must be pass-shaped and must name the reviewed exact DOCX/PDF path, page images, touched paragraph/page ids, body-vs-heading comparison, and final body-contamination verdict.
- A final acceptance record that says `XML only`, `PDF export only`, `page count only`, `manual visual only`, `sampled-only`, `not checked`, `no machine vision`, or any equivalent substitute proof for this surface is failed evidence.
- Multi-agent, sequential-fallback, or single-agent audit cannot accept worker claims for content expansion unless the audit parses the rendered machine-vision evidence and confirms the format lane reviewed the post-mutation exact output.
