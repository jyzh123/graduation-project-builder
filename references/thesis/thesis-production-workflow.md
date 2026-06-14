# Thesis Production Workflow

Use this as the default end-to-end workflow when turning a finished program into a matching graduation thesis.

## Enforcement Status

- Highest-priority writing rule: every newly written paragraph and every modified paragraph must be followed immediately by a rendered-page machine-vision review of that exact paragraph region before the next paragraph is touched.
- That paragraph-level review must verify both:
  - formatting against the active template or sample
  - content truth against the real project, local subsection purpose, and current user requirements
- The formatting side of that paragraph review must explicitly check font family, font size, font color, run-level emphasis drift, and abnormal spacing artifacts.
- If either side fails, repair the same paragraph immediately and rerender it before continuing.
- Every workflow step, gate, and completion rule in this file is mandatory when this file is loaded for the current thesis-production run.
- Do not downgrade any step here into optional guidance unless a higher-precedence source explicitly overrides it.
- No-implicit-skip rule: every required thesis step, gate, or figure family must end as passed, failed, or explicitly skipped with a reason. Implicit skip is a workflow failure.

## Goal

The thesis workflow is complete when the work has progressed from:

1. program understanding
2. chapter drafting
3. figure and table generation
4. code-snippet extraction
5. Word assembly
6. troubleshooting-guided rebuilds when needed
7. final formatting against the user's template

After step 7, the task is considered finished.

## Workflow

### 0. Paragraph-Level Machine-Vision Cycle

Apply this micro-cycle before and during every thesis-writing action:

1. write or modify exactly one paragraph
2. render the affected page or local page region
3. review that paragraph region with machine vision
4. confirm both format and content are correct
5. only then continue to the next paragraph

Mandatory interpretation:

- chapter-level review is not a substitute for paragraph-level review
- page-level review is not sufficient if it does not explicitly confirm the just-written or just-modified paragraph
- if the paragraph introduces a nearby caption, figure explanation, code explanation, or citation sentence, include those immediately affected lines in the same local rendered review
- if the paragraph rewrite changes line count, paragraph height, or local block order enough to move nearby content on the page, expand the same review cycle into a pagination-sensitive local review
- the minimum pagination-sensitive local review must cover:
  - the edited paragraph region
  - the immediately adjacent paragraph or block above
  - the immediately adjacent paragraph, heading, figure, table, or caption below
  - the next rendered page when the edited block sits near a page bottom or before a figure/table/caption cluster
- if the edited paragraph sits in the last page neighborhood of one chapter or the first page neighborhood of the next chapter, expand the same cycle further into a chapter-boundary review
- the minimum chapter-boundary review must cover:
  - the last affected page of the earlier chapter
  - the chapter opener page of the next chapter
  - the chapter title paragraph and the first body paragraph beneath it
- if the paragraph fails review, do not continue drafting later paragraphs first
- the rendered review for that paragraph cycle must be recordable through a review evidence record file for later acceptance review
- before handoff, any thesis modification run must also record a touched-page or blast-radius review evidence record that covers the modified page neighborhood
- a touched-page review is incomplete if it confirms the edited paragraph only but leaves local pagination status unknown for a nearby heading, figure, table, caption, or next-page continuation block
- a touched-page review is also incomplete if a touched chapter boundary exists but the earlier page and chapter-opener page were not both reviewed as a pair

### 0A. Thesis Lane Lock

Before the main thesis workflow continues, explicitly lock all of the following:

1. current lane:
   - full drafting
   - content-only revision
   - format-only repair
2. whether content is frozen
3. current master manuscript path
4. paragraph-level rendered review path
5. required figure families for the current run
6. whether real runtime screenshots are required
7. word-count target or explicit user override
8. where rendered review evidence record files will be stored for later acceptance review

If any of these remain unknown, the thesis run is not ready to proceed.

### 0B. Final Artifact Lock And Submission-Readiness Gate

Before any run may claim a thesis is `ready to submit`, `可提交`, or equivalent, explicitly lock all of the following:

1. current master manuscript path
2. current review-copy path
3. exact final submission DOCX path
4. exact final submission PDF path when PDF export is required
5. whether cover identity fields are complete, intentionally blank by user waiver, or still unresolved
6. whether final review will be:
   - sampled review only
   - or full sequential rendered review from cover to end matter

Mandatory interpretation:

- do not let the master manuscript, review copy, and final submission artifact collapse into one ambiguous `current draft` concept
- a submission-readiness claim is invalid if it was made against the wrong file path, an earlier intermediate draft, or a review copy that is not the final named deliverable
- sampled review is not sufficient for a submission-readiness claim
- if cover identity fields, protected front matter, references, acknowledgement, or end-matter formatting are unresolved, the run may still be described as structurally improved or near-final, but not as submission-ready
- unresolved cover identity fields block a submission-readiness claim unless the current user explicitly waives them for the current run

### 1. Read the real project first

Before writing the thesis:

- read the repository structure
- read README and docs
- read main app entrypoints
- read data-processing scripts
- read database sync or schema files
- read routes, views, controllers, and key business logic

The thesis must reflect the actual project, not a guessed system.

### 2. Use the remembered default thesis sample first

Use `references/thesis/thesis-template-learning.md` as the default remembered sample.

Do not re-discover the same sample each time unless:

- the user provides a different template
- the user explicitly asks to replace the current default structure

If a new template is provided, then extract its chapter skeleton, heading depth, numbering style, and figure/table placement habits.

### 3. Build the thesis skeleton

Create:

- thesis blueprint
- outline
- figure and table plan
- front/back matter notes

At this stage, decide which real project modules map to:

- system analysis
- system design
- system implementation
- system testing

Design-chapter completeness gate:

- if the thesis contains a design-oriented chapter, that chapter is incomplete until the required design figures are planned explicitly
- the minimum required structural design set is:
  - overall system architecture diagram
  - business or core workflow diagram
  - ER diagram or equivalent database structure figure when persistence exists
- if the template, chapter semantics, or current-user instruction imply additional structural figures, add them to the plan before moving into final formatting
- runtime screenshots do not satisfy this structural-design requirement by themselves

Structural-figure family closure gate:

- once the figure plan names multiple structural figure families, treat each family as a separate required surface rather than as a shared "diagram task"
- before leaving figure generation, record every planned structural figure family as passed, failed, or explicitly skipped with a reason
- if the figure plan lists architecture, ER, business flow, and pricing flow, completion of one flowchart does not close the other three families
- do not proceed into Word assembly while any planned structural figure family is still unresolved, still sample-unchecked, or still only present as a draft-form placeholder

### 4. Draft the thesis text

Write chapter content in order:

- abstract and abstract-en
- chapter 1 introduction
- chapter 2 system analysis
- chapter 3 system design
- chapter 4 system implementation
- chapter 5 system testing
- chapter 6 conclusion and outlook

Drafting execution rule:

- do not write a whole subsection first and review afterward
- after each paragraph is added or changed, run the paragraph-level machine-vision cycle from section `0`
- do not treat raw inline citation text preseeded in draft source paragraphs as an acceptable intermediate state for delivery
- if chapter 4 or chapter 5 prose is generated from repository facts, first classify each sentence as either:
  - project-fact sentence
  - external-source sentence
- only external-source sentences may enter the citation finalizer path; project-fact sentences must remain uncited unless the user explicitly requires another convention
- if the current drafting or revision pass touches a displayed formula, do not materialize the formula through plain text paragraph insertion
- displayed formulas must go through a formula-aware path that preserves:
  - equation object authenticity
  - template-matching numbering family
  - same-line right-aligned numbering placement when required
- if no formula-aware path is available yet, stop and build that path before continuing thesis drafting

Implementation-chapter screenshot drafting rule:

- when chapter 4 or another implementation subsection discusses multiple pages or modules with multiple screenshots, draft it as repeated local blocks rather than as one prose block plus one screenshot block
- default local block order:
  1. subsection lead-in paragraph
  2. page or module implementation paragraph
  3. the corresponding screenshot
  4. the corresponding caption
  5. next page or module implementation paragraph
- if the current manuscript instead has all screenshots grouped at the bottom of the subsection, treat that as a content-structure failure and reflow the subsection before continuing

## Writing constraints

- each subsection must start with a lead-in paragraph
- each subsection should normally be at least 300 Chinese characters
- unless the user explicitly overrides it, the full thesis body should default to no less than 10000 Chinese characters
- For mechanical-design whole-thesis rebuilds, lock the stronger profile before drafting or final QA unless the current user or school template gives a different target: at least 20000 visible Chinese-character/word units, at least 60 substantive bibliography entries, at least 10 foreign/English bibliography entries, formulas as real Word equation objects in body calculation/design chapters, body figures in the main narrative, and all final CAD sheet renders embedded again in the appendix.
- The mechanical-design profile is a single profile, not separate optional preferences. A pass on appendix drawings, formula count, reference count, or body figures cannot compensate for a miss in another profile field.
- do not leave sections as one-sentence placeholders
- keep wording consistent with the implemented system
- treat the English abstract as a translation of the Chinese abstract, not as an independently rewritten summary

Full-thesis delivery gate:

- a full thesis generation run must not be proven by a smoke fixture, minimal/manual-review fixture, or detector-only sample
- the exact final DOCX/PDF must contain, at minimum, cover/front administrative page preservation, Chinese abstract, English abstract, TOC, body chapters, at least one figure/screenshot surface when the project is software-based, at least one table surface, references, and acknowledgement
- the rendered output must not contain smoke, placeholder, minimal/manual-review, detector-fixture, self-check, skill-gate, `Example Research`, template-alignment, or other build-provenance/meta text
- the rendered output must not contain U+FFFD replacement glyphs, citations inserted inside ordinary tokens, fake bibliography entries, or repeated long body paragraphs caused by generator fallback
- the acceptance record for a full-thesis test must fail if `sample_self_check` reports `full thesis content gate failed`, `smoke-only; blocked for delivery`, or `smoke acceptance mode: yes`

Canonical engine and local adapter split:

- full thesis production must use canonical skill scripts for general thesis-making behavior
- project-local generation files are allowed only as adapter/profile/manifest files for one current template and project
- a local adapter must pass `scripts/validate_thesis_local_adapter.py` before it can be used by the canonical builder or recorded as evidence
- if the project needs a new general DOCX assembly capability, add it to the canonical skill scripts and tests rather than generating a project-local builder
- if the project needs template-specific details, record them as adapter data: donor paragraph identifiers, template fingerprint, field labels, surface expectations, content/evidence paths, screenshot paths, bibliography inputs, and locked output paths

## Humanization policy

Apply language humanization only during thesis content drafting, rewriting, or abstract-polish work.

Do not invoke language-humanization rules during:

- `format-repair-only` runs that are purely about template compliance
- global formatting passes
- pagination, heading, TOC, caption, or spacing repair

Chinese thesis text policy:

- when drafting or rewriting the Chinese abstract or Chinese body chapters, apply `humanizer-zh` after the factual-consistency pass
- use it to remove AI-sounding filler, empty significance claims, template transitions, promotional wording, and customer-service tone
- do not let the pass introduce new facts, remove technical detail, or turn formal thesis prose into casual chat
- this is mandatory for Chinese abstract drafting, Chinese abstract rewriting, and Chinese body-section drafting or rewriting
- do not treat `humanizer-zh` as optional once the current run is producing new Chinese thesis prose or rewriting existing Chinese thesis prose

English abstract policy:

- draft the English abstract from the finalized Chinese abstract first
- then apply `humanizer` as a smoothing pass for unnatural English phrasing
- keep the English abstract semantically aligned with the Chinese abstract
- do not add claims, metrics, modules, limitations, or conclusions that are absent from the Chinese abstract
- do not use `humanizer` to turn the English abstract into an independently expanded summary
- this is mandatory for English abstract drafting, English abstract rewriting, and any English thesis-body rewriting that the current run intentionally performs

Mandatory recording rule for humanizer routing:

- every thesis drafting or thesis rewriting run must explicitly record:
  - whether `humanizer-zh`, `humanizer`, or both were used
  - which language surfaces they touched
  - which surfaces were intentionally excluded
  - one or more evidence paths that record the processed paragraph or abstract block
  - for each processed paragraph, the evidence must include before text, after text, target language, processed paragraph ID, and skill name
- if the run is `format-repair-only`, the humanizer route must be recorded explicitly as `none`
- do not hand off a thesis writing run while the humanizer route remains implicit or while humanizer evidence is only a path, screenshot, summary, or checklist note

Execution order for abstract work:

1. finish or repair the Chinese abstract against the real project
2. humanize the Chinese abstract with `humanizer-zh` when needed
3. translate the Chinese abstract into English
4. humanize the English abstract with `humanizer` when needed
5. run a semantic-alignment check between the two abstracts before assembly

### 5. Generate thesis figures from the real program

Do not stop at prose. Produce the figures the sample expects.

Typical figure sources:

- architecture and technical architecture: derived from repository structure
- workflow diagrams: derived from actual program flow
- DFD diagrams: derived from data movement in scripts and app routes
- ER diagram: derived from schema or database sync code
- chapter 4 images: prefer real page screenshots from the running program
- chapter 5 images: use real charts or test screenshots

Figure production rule:

- use `references/thesis/thesis-figure-generation-rules.md` as the canonical rule source for figure authoring, styling, text legibility, and insertion readiness
- use `references/review-figure-style-checklist.md` as the canonical review gate before and after insertion
- if the thesis lacks the expected non-screenshot figures, the thesis is incomplete even if the prose has been drafted
- if a design chapter still lacks architecture, workflow, or ER figures after prose drafting, stop and complete figure generation before entering format-only work
- if a structural figure family is sample-locked or draw.io-first, do not accept a generic Mermaid quick diagram as the final production asset for that family

### 6. Add core code for each function

For each concrete function discussed in the implementation chapter:

- extract a matching code snippet from the real codebase
- place the snippet near the relevant subsection
- do not invent code

Typical targets:

- login
- home/dashboard
- query page
- map or visualization
- data generation or data import
- data cleaning
- statistical analysis
- database sync

### 7. Assemble a Word draft

Create a Word draft that includes:

- title
- abstract
- abstract-en
- visible directory or TOC
- body chapters
- figures
- tables
- code snippets

Do not mix deployment instructions into the thesis document.

Bibliography truth gate:

- before final formatting or final handoff, classify every bibliography item as either:
  - queryable paper
  - non-paper source
  - unverifiable source
- if the current user requires real papers only, the bibliography fails this gate unless every kept item is both:
  - real
  - queryable as a paper through an academic discovery path
- if a non-paper or unverifiable item is removed during this gate, rerun a local body audit on every sentence that cited it and either:
  - replace the support with a real queryable paper
  - or rewrite the sentence as an uncited project fact when appropriate
- do not treat bibliography cleanup as complete if the list has been corrected but the body still depends on removed non-paper evidence
- when the current run requires paper-only literature, execute `scripts/audit_paper_only_bibliography.py` on the exact review-copy or final-deliverable DOCX and persist the markdown report path into the acceptance record
- do not enter final formatting acceptance while the paper-only bibliography audit still reports non-paper items, unverified items, or body dependencies on rejected sources

Citation assembly gate:

- before delivery, run a citation-only audit on the final chapter 4 and chapter 5 body paragraphs
- the audit must explicitly fail if any of the following remain:
  - plain baseline bracket citations such as `[n]` left as ordinary body text
  - chapter 4 or chapter 5 project-fact sentences carrying guessed literature numbers without a directly traceable source statement
  - citation markers inserted by whole-paragraph text assignment rather than a citation-aware run-preserving path
- if the citation-only audit fails, do not continue into final handoff or final formatting acceptance

Word-assembly preflight for figures:

- before assembly, compare the actual figure assets against the current figure plan
- verify that every planned structural figure family has a resolved status and a source path
- verify that every planned structural figure family uses a draw.io source path and a draw.io export path when the figure is a thesis structural diagram rather than a runtime screenshot
- verify that implementation-chapter runtime screenshots are inserted next to the paragraph block that explains them, rather than collected as a gallery at the end of the subsection
- if any planned structural figure family is missing, unresolved, or still represented only by a quick draft artifact, stop and return to figure generation

### 8. Apply remembered troubleshooting fixes during assembly

Before finalizing the thesis Word draft, read and apply `references/thesis-troubleshooting-log.md`.

This is now part of the standard workflow, not an optional note.

### 9. Final stage: template-only formatting

Once the thesis content, figures, tables, and code snippets are all in the Word draft, the remaining work should be treated as formatting only:

- adjust according to the user-provided template
- adjust fonts, line spacing, heading styles, page breaks
- adjust figure captions and table captions
- adjust visible directory or TOC layout

At this stage, the task should not reopen content generation unless the user explicitly requests content changes.

Completion boundary for format-only stage:

- format-only work may start only after required design figures, runtime screenshots, and tables are already present in the manuscript
- if a later rendered-page review reveals that a required figure family is still missing, treat that as a workflow failure and return to figure generation instead of trying to hand off a format-only result

Before any global formatting pass such as font replacement, spacing normalization, caption cleanup, or heading cleanup:

1. identify the true current master manuscript
2. confirm that manuscript already contains the latest references, superscripts, screenshots, and inserted code snippets
3. only then run the formatting transform

If this preflight is skipped, formatting work can silently regress citation chains or newer thesis content. Treat that outcome as a workflow failure and restart from the newest master manuscript.

If the default `python` or `py` launcher fails during thesis automation:

1. recover a usable interpreter first by locating installed Python runtimes
2. verify the required packages for the current document pipeline
3. continue the planned scripted workflow with that recovered interpreter path
4. only fall back to non-Python tooling if no usable interpreter can be recovered

Do not treat a broken default launcher as permission to abandon the scripted workflow immediately.

### Heading Repair Guard

Before any heading-repair or chapter-title rebuild step:

1. identify the front-matter zone
2. identify the TOC zone
3. identify the first true body chapter page
4. restrict heading repair to the confirmed body zone only

Never run heading remapping on the entire paragraph list without excluding TOC and front matter first.

If chapter-title insertion is needed:

1. verify on rendered pages that the target chapter title is actually missing
2. insert only after that confirmation
3. immediately rerender the TOC and the modified chapter start page

If TOC lines become headings, chapter titles duplicate, or first-level headings disappear after the pass, treat the branch as invalid and revert immediately.

### Structural Figure Hard Gate

After any structural figure is drawn and before insertion, run the pre-insertion gate from:

- `references/thesis/thesis-figure-generation-rules.md`
- `references/review-figure-style-checklist.md`

If either gate still fails, the figure must be redrawn before insertion.

### Post-Insertion Figure Gate

After inserting a structural figure into the DOCX, run the post-insertion gate from:

- `references/thesis/thesis-figure-generation-rules.md`
- `references/review-figure-style-checklist.md`

Do not move on to the next figure until the current one passes those post-insertion checks.

After the formatting stage and before handoff, run a mandatory final thesis QA checklist:

## Thesis Completion Review

1. verify citations are not attached to any heading, title, figure caption, table caption, or code caption
2. verify citations are placed at the end of the claim-bearing sentence, immediately before sentence-ending punctuation, unless the user explicitly requires anchor-level in-sentence placement
3. verify `scripts/audit_thesis_citations.py` passes on the exact review-copy or final-deliverable DOCX and persist that report path into the acceptance record
4. verify TOC order, indentation, and page-number presentation on rendered pages
5. verify body font and font size are uniform across normal body paragraphs
6. verify figures are centered and captions are separated correctly
7. verify implementation-chapter screenshots are interleaved with their corresponding explanatory paragraphs instead of being grouped at the end of a subsection
8. verify the rendered tail pages do not contain unintended blank pages
9. verify Chinese abstract and English abstract formatting separately from normal body formatting
10. verify every chapter-level heading starts on a new page in the rendered output
11. verify the English abstract is semantically aligned with the Chinese abstract and does not introduce extra or missing claims
12. verify Chinese abstract and Chinese body text do not still contain obvious AI filler after the final content pass
13. verify any `humanizer` or `humanizer-zh` pass preserved thesis tone, technical specificity, and factual boundaries
14. when the current run requires paper-only literature, verify every bibliography item is a real queryable paper rather than documentation, financial reports, or general web pages
15. when a non-paper source has been replaced by a paper, verify the citing body sentence was also rewritten so the new paper actually supports that sentence

If any item fails, the thesis is not handoff-ready.

When applying global formatting:

- exclude the abstract and abstract-en blocks from body-style normalization unless the target transformation is specifically for abstract formatting
- preserve or reapply explicit page breaks before every chapter-level heading
- when abstract formatting needs repair, copy paragraph formatting directly from the template source manuscript instead of reconstructing it heuristically

## Correction-handling rule

If the user says the thesis formatting, template-following, front matter, TOC, figure placement, or page structure is wrong:

1. record the correction immediately in the project learning log
2. translate the correction into a concrete formatting countermeasure checklist
3. compare the current output against the real template and rules document, not against generic thesis conventions
4. continue fixing until the corrected rule is visibly reflected in the generated Word output

Do not treat a user formatting correction as a one-line note. It must change the workflow and the next rebuild.

If the user continues to direct revisions after that, each accepted revision must also be written back into the skill or its references immediately, not postponed until the end of the project.

If the user indicates that the current thesis output is still broadly unqualified, escalate from local correction mode to full rebuild mode:

1. go back to the real repository and re-read the implemented system
2. verify whether the program can be run locally
3. if it can run, obtain real screenshots for implementation-related sections instead of using placeholders
4. rebuild from the original thesis source or template baseline rather than stacking patches on a bad intermediate file
5. perform visual verification on the rebuilt result before handoff

If the user reports that a generated structural figure contains line crossings through unrelated shapes, duplicate titles inside the image, gray fills on a white-background requirement, or other direct violations of the thesis figure rules:

1. stop using the current figure generation path immediately
2. treat the current figure source as invalid
3. update the skill rules before drawing again
4. rebuild the affected figure set from a compliant source path instead of patching the invalid images in place

## Completion rule

After the thesis has reached the state:

- content drafted
- figures generated
- tables inserted
- code snippets inserted
- Word file assembled
- troubleshooting lessons applied
- only template formatting remains

the workflow is considered complete enough to hand over for final formatting.

## Submission-Claim Rule

- Do not issue a final `可提交` or equivalent verdict until the exact final submission artifact has passed:
  - structure review
  - content-consistency review
  - final rendered-page review from cover through end matter in page order
- A run that passes only structure review, XML validation, outline inspection, or sampled page checks must be reported as incomplete for submission purposes.

## Thesis Image Generation Default

Apply this workflow in every thesis-related mode:

- `program-plus-thesis`
- `thesis-only`
- `format-repair-only` when figures, screenshots, tables, or diagrams are in scope

### Mandatory rules

- If the user provides figure samples, screenshot samples, or template figure examples, use them as the highest-priority visual target.
- If no sample is provided, follow the current template and the stable thesis figure rules.
- 设计类章节需要结构图、流程图、ER 图，如果模板没有给出样式，从skill内部的样式抽取样式，不算可选项。
- After generating or replacing a figure, verify chapter-semantic match, caption correctness, numbering, placement, and actual embed success in the final document.
- Do not close a thesis task while required figures are still missing or visibly inconsistent with the approved style source.
