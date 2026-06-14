# Thesis Format Rules: General And DOCX Safety

Use this file for baseline principles, safe DOCX editing rules, and detection scope.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current format-repair subtask.
- Apply this file together with `references/thesis/thesis-format-rules.md`.

## 1. General Principles

- Treat the school template or teacher sample as the formatting baseline.
- Use baseline priority in this order:
  - school or department official thesis specification
  - user-approved sample or explicit formatting requirement
  - common Chinese graduation-thesis defaults when no school-specific baseline exists
- If the baseline is only a general default, state that clearly instead of implying school-level compliance.
- Keep content revision and format repair as separate passes.
- Run a detection pass before editing. Do not start by patching the DOCX blindly.
- If the thesis output becomes visibly polluted, rebuild from a clean baseline instead of stacking more local patches.
- Treat cover, declarations, abstracts, TOC, body chapters, references, acknowledgement, and appendix as separate formatting surfaces.
- Final acceptance is based on rendered pages, not only on paragraph style names or DOCX internals.
- Treat every template-owned surface as a full paragraph/run instance, not as text plus cosmetic styling. The required surface set includes cover lines, declarations, Chinese and English abstracts, keywords, TOC title and levels, headings, body paragraphs, figure holders, figure captions, explanatory paragraphs near figures/tables, table titles, table cells, continuation titles, formulas, code titles, code blocks, references, acknowledgements, headers, footers, and page numbers.
- For each touched surface, extract and record the approved baseline before repair: style id/name, outline level, alignment, first-line/left/right/hanging indent, line spacing, spacing before/after, tabs and leaders, keep-with-next/keep-lines/page-break-before, borders/shading when visible, run font family for ascii/eastAsia/hAnsi/cs, font size, bold, italic, underline, color, highlight, superscript/subscript, and field/bookmark/hyperlink behavior when present.
- Each thesis surface family must bind to its own locked template class. Do not bind references, reference titles, acknowledgements, appendix titles/content, abstract titles/body, keyword labels/content, figure captions, table titles, table-cell text, headings, or citation-marker runs to the generic body/Normal class unless the active template proves that the same style binding is the approved donor for that exact surface.
- A format audit is incomplete if it checks only body paragraphs. It must separately verify style binding and paragraph/run baselines for the touched or user-reported classes: every heading level, body text, figure body/holder, figure caption, table title, table-cell text, abstract title, abstract body, keyword label, keyword content, reference title, reference entry, citation superscript, acknowledgement title/content, appendix title/content, header, footer, and page number.
- A format audit is incomplete if it checks only body paragraphs. It must separately verify style binding and paragraph/run baselines for the touched or user-reported classes: every heading level, body text, figure body/holder, figure caption, table title, table-cell text, abstract title, abstract body, keyword label, keyword content, reference title, reference entry, citation superscript, review comments and tracked changes when present, acknowledgement title/content, appendix title/content, header, footer, and page number.
- Body text is not a paragraph-style-only surface. If a touched or user-reported body paragraph contains both Chinese text and English letters, identifiers, API routes, file names, model names, or digit-heavy ASCII tokens, the audit must check run boundaries and the resolved `ascii` / `hAnsi` / `eastAsia` / `cs` font chain for that same paragraph instead of accepting one paragraph-wide font verdict.
- Body text audit must fail on global or local shrinkage, not only heading-like enlargement. If direct run size, style-chain size, or rendered density makes real body prose smaller than the locked body baseline, the body-text surface is polluted even when the paragraph style name is `Normal`.
- A surface does not pass merely because it remains on the same rendered page, does not overflow, or has the correct visible text. It must pass both DOCX-internal parity against the baseline instance and rendered-page visual review on the exact review copy.
- If the user explicitly points out a format problem on a surface, the next repair pass must audit the complete surface family rather than only the symptom shown in the screenshot. For example, an abstract complaint triggers all six abstract surfaces; a table complaint triggers title, grid/borders, header cells, body cells, and nearby body paragraphs; a heading complaint triggers the same heading level and adjacent body paragraphs.
- Do not create or hardcode fallback font families for Chinese, Western, or English text in any thesis surface when a template/sample baseline exists. If a helper script or manual XML patch needs a font value, copy the value from the locked baseline instance, including separate East Asian and Western font attributes. If the baseline is missing, stop and record the blocker instead of inventing `Times New Roman`, `Arial`, `宋体`, `黑体`, or any other default.

## 2. Safe DOCX Editing

- Work on a copied output file, not on the original template.
- Do not replace the accepted manuscript or current review copy by copying an unrelated sample thesis over it just to inherit layout. For sample-following rewrite work, keep the target manuscript path stable and import only the explicitly locked surfaces or paragraph metrics from the sample.
- Use `references/tooling-dependencies.md` as the canonical toolchain source for thesis `.docx` work.
- Use `references/thesis/thesis-format-sop.md` as the canonical execution-order and OfficeCLI routing source.
- Before editing, confirm `officecli` is available and use the baseline inspection set required by the active SOP/tooling references.
- Prefer `officecli open` / `close` plus targeted `set`, `add`, `move`, `remove`, or `batch` edits for in-place repair.
- Use `officecli raw` / `raw-set` only when high-level operations cannot safely express TOC, field, bookmark, tab-stop, tracked-change, or style-binding changes.
- Do not default to `python-docx` or other library-wide whole-paragraph rewrites when `officecli` can target the affected paragraph, run, table, image, or heading directly.
- Do not let local helper code, script defaults, WPS defaults, LibreOffice defaults, or office-application fallback styles become the formatting authority. They may only apply values already extracted from the approved baseline or explicitly supplied by the user.
- Do not rebuild an existing thesis manuscript by creating a new generic DOCX body and then copying visible text into it. That route discards section ownership, field behavior, style bindings, numbering, tabs, page breaks, media relationships, table geometry, and direct run formatting.
- In `whole-thesis-revision` and `local-surface-repair`, the default mutation model is structure-preserving edit on a locked review copy. Full-body reconstruction is allowed only after reclassifying as `new-thesis-production`, locking the active template/profile, and recording that the prior manuscript is used only as content input.
- If the previous output lost TOC page numbers, defaulted headings to blue Word theme styles, changed page margins, reduced page breaks, or compressed appendix code by generic code-block styling, treat that output as contaminated and restart from a clean source before further format work.
- For content-only thesis revision on an accepted manuscript, do not use Word/WPS open-save or Word COM as the primary mutation path.
- Before any content-only thesis pass, record a package baseline for:
  - `word/document.xml`
  - `word/styles.xml`
  - `word/settings.xml`
  - `word/fontTable.xml`
  - `word/numbering.xml` when present
  - `word/header*.xml` / `word/footer*.xml` when present
  - `word/_rels/document.xml.rels`
  - `[Content_Types].xml`
- In a content-only pass, unexpected drift outside `word/document.xml` is a tooling failure until explained and intentionally accepted.
- Do not rewrite a thesis paragraph by assigning one new paragraph text string when the paragraph contains superscript citations, mixed bold/normal labels, manual line breaks, or other meaningful run-level formatting.
- Do not rewrite a touched body paragraph that contains mixed Chinese and English content by assigning one concatenated replacement string or by rebuilding the paragraph from a single donor run. Preserve the original run boundaries when they are still valid, or replay the locked donor run family with separate Chinese and Western runs before handoff.
- Do not "repair" mixed-script body text with a hardcoded token whitelist. If the paragraph contains English words, identifiers, routes, filenames, or digit-heavy ASCII fragments, the split/font decision must come from the locked donor baseline or canonical builder logic rather than a local allowlist.
- Do not strip comments, tracked changes, bookmarks, field anchors, hyperlinks, citation bookmarks, or citation superscript runs as collateral cleanup. If the user wants a comment-free copy, generate it as a separate preview from a review-artifact-preserving source and record explicit approval plus source-to-final review-artifact diff evidence.
- In `new-thesis-production`, if the source DOCX is an old-topic manuscript used only to carry the school layout/template and the final DOCX intentionally replaces the subject, old-source bookmarks, field anchors, hyperlinks, and citation numbers may be excluded from preservation only through the SHA-bound `new-thesis-source-artifact-disposition` path owned by `EXEC-MAINT-059`. The final DOCX must still pass independent review-artifact generation, body-citation audit, bibliography-content audit, font audit, and whole-format gates. Do not use this route for local repairs or comment-driven revision.
- Before rewriting any citation-bearing paragraph, inventory each citation marker run and its host relationship. After rewriting, verify every marker remains a separate superscript run with the approved visual style or has a citation-lane controlled-change ledger.
- Do not rewrite TOC blocks, cover label rows, abstract title/body/keyword lines, heading paragraphs, or footer/page-number paragraphs by assigning one concatenated plain-text string. Those surfaces are paragraph-structure-sensitive and must preserve tabs, leaders, page breaks, fields, and per-entry paragraph boundaries.
- Treat sample footer literals such as borrowed arabic page numbers, roman numerals, and sample dates as unsafe carry-over data. If a sample DOCX is reused as a structural baseline, rebuild the target numbering and footer fields explicitly instead of inheriting those literal values.
- When inserting a new body paragraph, copy paragraph metrics from a real accepted body paragraph instance first:
  - first-line indent
  - line spacing
  - alignment
  - space before / after
  - effective font mapping
- For `conclusion`, `references`, `acknowledgement`, `appendix`, and similar tail-block title paragraphs, copy the real approved title paragraph instance first and then replace only the visible title text.
- Do not synthesize tail-block title typography by manually scripting East Asian font-family literals when a real template/sample title instance already exists.
- Do not use paragraph-wide wildcard rewrites across the cover, abstracts, keyword lines, or TOC block.
- Use Unicode-safe edits for Chinese title and chapter strings.
- Do not anchor repairs on TOC text alone.
- Do not rely on fixed paragraph indexes for figure success.
- Do not rely on fixed `Word COM` paragraph indexes for body-text, heading, caption, or reference edits.
- Do not assume `python-docx` paragraph indexes and `Word COM` paragraph indexes refer to the same physical paragraph.
- When tables exist, treat `Word COM` paragraph numbering as unsafe for direct write targeting unless the candidate paragraph has first been verified as outside a table.
- When the builder calls fix-up scripts through `runpy`, catch `SystemExit(0)` and continue so the full finalizer chain can run.
- Avoid export paths that reopen and save the cleaned DOCX in a way that reintroduces numbering or style drift.
- Keep the repair order stable when multiple passes exist:
  - page setup
  - paragraph formatting
  - citations
  - references
  - special headings
  - header and footer
- Do not let one bulk helper script own cover, abstracts, TOC, body headings, body text, pagination, and header/footer in the same pass. Split the pass by protected surface family and verify each family before the next one.
- Do not run style-unification passes before pagination, section-boundary, and TOC repair are stable.
- Treat combined screenshots, image replacement, and empty-paragraph cleanup as high-risk structural mutations. After those actions, rerun outline, text, issues, and rendered-page review before delivery.
- Prefer detecting against normalized text when headings may already contain added spacing or tabs. For example, compare a cleaned tail-block title string instead of raw paragraph text only.
- Do not use broad substring matches for heading detection when TOC lines can contain similar text.

## 3. Detection Scope

Before repair, audit at least these surfaces:

- page setup
- style definitions that own body text, headings, and TOC
- paragraph alignment, spacing, and indentation
- font family, font size, and font color at both paragraph and run level
- unexpected run-level formatting such as bold, italic, underline, highlight, shading, and hyperlink-like color drift
- loss of protected run-level formatting such as superscript citations or split label/body runs
- loss of mixed-script body-text run boundaries where Chinese text and English/ASCII content require separate font-slot handling
- loss of review comments, comment anchors, tracked-change marks, bookmarks, field hosts, hyperlinks, or citation bookmarks
- abnormal blank pages, abnormal blank lines, and abnormal extra spaces
- abnormal extra spaces must include:
  - repeated ASCII spaces
  - stray full-width spaces
  - stray tabs
  - mixed-spacing artifacts introduced by copy/paste or DOCX rewrites
- table borders and title placement
- figure placement and caption placement
- figure-internal style, including internal typography, stroke weight, fill behavior, spacing rhythm, and connector style when the figure is not a raw program screenshot
- header and footer content and typography
- citations and references

If the task is explicitly format-repair-only, still do the detection pass first and then repair in a controlled order.
