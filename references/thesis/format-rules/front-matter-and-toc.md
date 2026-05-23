# Thesis Format Rules: Front Matter And TOC

Use this file for front-matter formatting and TOC-specific repair rules.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current format-repair subtask.
- Apply this file together with `references/thesis/thesis-format-rules.md`.

## Routing Note

- This file is the canonical structural rule source for front-matter and TOC repair.
- For durable user-correction overrides and sample-following interpretation, also load `references/user-feedback/template-and-layout.md` when relevant.
- For purely visual comparison fallback, use `references/thesis-layout-visual-memory.md`.

## 4. Front Matter Rules

- Reuse sample paragraph formatting directly for cover, declaration, authorization, abstract titles, and TOC title whenever possible.
- Do not invent front-matter section boundaries from a fixed builder skeleton. Extract the real front-matter zone boundaries and numbering scheme from the active local template or approved sample first.
- When repairing a cover, preserve the template cover structure. Prefer directly reusing the school template or an approved sample cover layout, and only make minimal field corrections inside that existing cover block instead of rebuilding the cover structure from scratch.
- Copy sample paragraph metrics together with the text structure:
  - first-line indent
  - line spacing
  - title centering
  - title spacing
  - signature-line rhythm
- Abstract and TOC acceptance must be surface-level, not page-class-level. Chinese abstract title, Chinese abstract body, Chinese keyword line, English abstract title, English abstract body, English keyword line, TOC title, TOC entries for every used level, TOC dotted leaders, and TOC page-number column each require an independent row with baseline donor, DOCX paragraph/run evidence, rendered page or region evidence, metric-by-metric comparison, and final verdict.
- For each protected abstract and TOC row, font evidence must record the actual effective font chain, not only the visible result or directly written `w:rPr`. The evidence must resolve direct run fonts, character style, paragraph style, basedOn styles, docDefaults, theme major/minor mappings, and WPS/Word UI displayed font names, then compare the resolved result against the template baseline for `ascii`, `hAnsi`, `eastAsia`, and `cs`.
- Theme/default font aliases such as `Calibri (Body)`, `Calibri (正文)`, `minorHAnsi`, `minorEastAsia`, or any `w:*Theme` font attribute are accepted only when the locked template baseline proves the same effective alias for that same protected surface. Otherwise they are font-format drift.
- The evidence record for each protected abstract or TOC row must name the exact protected surface id in both `target identifier` and `baseline surface id` (`zh_abstract_title`, `zh_abstract_body`, `zh_keyword_line`, `en_abstract_title`, `en_abstract_body`, `en_keyword_line`, `toc_title`, `toc_entries`, `toc_dotted_leaders`, or `toc_page_number_column`). A generic `abstract`, `toc`, `front matter`, or page-class id is not acceptable, and one evidence record must not be reused to prove multiple protected surfaces.
- A front-matter evidence record that only says `title present`, `keyword line present`, `page I`, `page II`, `TOC entries visible`, `page numbers corrected`, or similar existence/page-order language is under-specified and cannot support a pass verdict for abstract or TOC formatting.
- Cover and declaration paragraphs must be verified against the approved paragraph instances at paragraph and run level. A cover paragraph that visually becomes body text is still failing even if its text remains correct.
- Cover and declaration/commitment pages must not share the same rendered page.
- If the declaration title or declaration body appears on the same page as the cover title block, treat front-matter repair as failed and split the pages before any later body cleanup.
- Cover label rows, abstract label rows, keyword label rows, and English front-matter rows are template-owned independent classes. Do not batch-apply generic body paragraph rules to them.
- Cover identity tables are also template-owned independent surfaces. A cover table must keep the active template's table structure, row count, column widths, borders/underlines, cell margins, paragraph alignment, spacing, and run typography; repair may only refill target-specific values inside the template value cells.
- Do not treat a cover identity table as a body table, a three-line table, or a normal body paragraph container. If a body/table normalization pass makes cover cells inherit body text metrics or removes template spacer/value rows, restore the cover table from the template donor before handoff.
- Chinese abstract title, Chinese abstract body, Chinese keyword line, English abstract title, English abstract body, and English keyword line must be treated as six independent front-matter formatting surfaces.
- A centered and bolded body paragraph is not an acceptable substitute for a repaired abstract block. If the abstract title or keyword line still remains on the generic body-paragraph class after repair, treat the abstract block as still failing even when the rendered page looks roughly correct.
- If no approved abstract sample paragraph instance can be extracted from the active template source, record that gap explicitly and treat the abstract repair as blocked on baseline evidence rather than silently approximating the abstract with ordinary body formatting.
- Clear inherited headers explicitly when reusing a sample DOCX so unrelated chapter names do not leak into body pages.
- Front-matter review must be sequential and rendered-page-first:
  - inspect the cover page first
  - then inspect the Chinese abstract page(s)
  - then inspect the English abstract page(s)
  - then inspect the TOC page(s)
  - only after those pass may the builder inspect the first body chapter page
- During that sequential review, check actual page occupancy rather than only page order. A page that is technically present but visually near-empty because of a duplicated hard break still fails front-matter pagination review.
- Template-learned same-page groups are hard constraints, not suggestions. If the active template keeps the Chinese thesis title and English thesis title in the same cover title block, the generated or repaired manuscript must keep those title paragraphs on the same rendered page.
- The canonical builder must write a template profile that records front-matter page-class markers, same-page groups, separated-page pairs, and the source of each relationship. If the profile cannot determine a critical cover/title/abstract/TOC relationship, generation must stop instead of guessing pagination.
- `sample_self_check` must validate final rendered PDF occupancy against the template profile. A cover title block split across rendered pages, an English title drifting to a near-empty page, or any missing profile member page is a blocking failure even if the ordinary cover/abstract/TOC order check passes.
- Template profile paragraph indices describe the template source, not the final manuscript. Self-check must locate final cover title members from the final manuscript's actual cover title text before rendered-page lookup; using template indices directly on the final DOCX is a detector bug.
- Cover title instruction artifacts such as `中文题目（一号黑体字）`, `英文题目`, `如题目只有一行`, and `须删除本行` must be removed from generated manuscripts. If those drawing/textbox instructions remain visible, the cover page is not template-aligned and front-matter pagination review is blocked.
- If the cover page already contains abstract text, if the abstract page already contains TOC content, or if the TOC page already contains body text, the front-matter repair is still failing regardless of later body-page correctness.
- Chinese abstract, English abstract, TOC, and first body chapter must be separated on rendered pages; if English abstract and TOC share one rendered page, or TOC and body share one rendered page, the run is blocked even when all strings are present.
- Front-matter format repair must be treated broadly rather than as layout-only work. If the cover text itself is wrong, incomplete, stale, sample-placeholder text, or semantically mismatched to the thesis, that still counts as a front-matter formatting failure and must be repaired in the same pass.
- For generation from a template, if the content manifest provides a thesis title, the builder must replace the template sample title in the cover and declaration text; any residual sample title or title placeholder is a hard blocker.
- Cover identity fields must be replaced inside their original value runs and only inside the cover identity zone. Do not treat spaced labels such as `学    号` or `院    系` as title candidates, and do not rewrite declaration or signature rows just because they contain words such as `指导教师`.
- Front-matter acceptance also requires one office-application numbering-state check for the title paragraphs of the Chinese abstract, English abstract, TOC, and first body chapter when the template uses distinct front-matter numbering. Do not rely only on rendered TOC strings.
- If a borrowed sample footer or front-matter footer still shows the sample's literal date, roman numeral, or arabic page number after repair, treat front-matter repair as failing until the target manuscript's own numbering strategy has been rebuilt and rerendered.

## 5. TOC Rules

- Audit heading hierarchy before rebuilding the TOC.
- Treat TOC paragraphs as a protected formatting surface. Do not let heading-repair logic, chapter-title insertion logic, or generic "numbered line" detection touch TOC paragraphs.
- TOC correctness includes:
  - heading coverage
  - level-based indentation
  - dotted leaders
  - right-aligned page numbers
  - TOC body styling distinct from normal body paragraphs
  - numbering from exactly one source
- TOC title must remain a standalone title paragraph and must not be merged with the first TOC entry. `toc_entries` is not a valid substitute for per-level evidence: every TOC level actually present in the final manuscript must have its own baseline metrics, internal paragraph/run evidence, rendered-region evidence, and row-level verdict.
- TOC page-number acceptance must be per-entry. For each visible TOC entry, the visible page number must match both the corresponding heading's rendered page and the manuscript's displayed numbering system; the right tab, dotted leader endpoint, and page-number column position must match the locked template baseline.
- TOC visual acceptance must include a rendered geometry comparison against the template or approved sample. Required geometry evidence includes numeric title and first-entry bounding boxes, numeric per-row bounding boxes, numeric per-level left x positions, numeric row y-deltas / line spacing, numeric dotted-leader start/end/density, numeric page-number x column, numeric row count per page, numeric title-to-first-entry gap, and numeric page occupancy rhythm for both baseline and actual.
- TOC visual-geometry exact-output binding is owned by `FMT-EVID-007` in `references/thesis/format-rules/protected-surface-evidence-contract.md`; this file defines the TOC-local geometry metrics and repair behavior.
- A TOC cannot pass from content, font-chain, style-name, page-number, or screenshot existence alone. If rendered geometry evidence is missing, generic, or not compared metric by metric with the template, the TOC is still failing.
- A TOC cannot pass from natural-language geometry claims alone. Phrases such as `template equals actual`, `matched`, or `visual pass` must be rejected unless the same evidence record names distinct template and actual rendered images and gives the numeric baseline/actual geometry values.
- The TOC repair sequence is mandatory:
  - first generate or refresh a clean TOC block
  - then verify that every TOC line is a real heading entry and that no body paragraph, caption, keyword line, or other non-heading text has been pulled into the TOC
  - only after the TOC content, hierarchy, and page references are verified clean may the builder extract sample formatting for TOC title and each TOC level and apply the final visual tuning
- Do not replace a multi-paragraph TOC block or TOC content-control region with one plain paragraph or one concatenated text payload. If the TOC title and entries collapse into one paragraph, the TOC repair has failed regardless of whether the text strings are present.
- For `1:1` template alignment, the TOC implementation form is itself a protected formatting surface. If the active template uses `w:sdt`, `w:fldSimple`, `w:instrText` with a `TOC` field instruction, hyperlinks, bookmarks, smart tags, or another wrapper around visible TOC runs, the target must preserve or deliberately reconstruct the same accepted wrapper family. A static reconstructed TOC that merely imitates the visible rows cannot pass when the template's editable TOC structure was available.
- Canonical static TOC generation must clone real template TOC paragraph instances and rewrite only the visible label/page text while preserving `pPr`, runs, tabs, leaders, and page-number run structure.
- Do not update TOC page numbers with `para.text` or any whole-paragraph plain-text assignment; update only page-number text nodes after the template tab run.
- TOC refresh/page-sync acceptance must include run-level direct-font restoration from the locked template TOC donor. Paragraph style names such as `TOC 1` are not enough: refreshed entries must not carry body-heading direct font overrides unless that same direct font is present in the template TOC donor for that level and run side.
- TOC helper scripts must traverse visible runs nested inside WPS smart tags, content controls, hyperlinks, or other wrappers. A repair that inspects only direct child `w:r` nodes can miss the actual displayed TOC text, tab/leader, or page-number run and is not an acceptable enforcement path.
- The run-level restoration proof must be validator-backed hard evidence, not a prose note. For the TOC title and every used TOC level, record template/actual visible-run typography separately for entry text, tab/leader, and page-number runs, including direct `w:rPr`, `w:rFonts` script/theme slots, size/sizeCs, and weight. A TOC whose paragraph spacing and tab stops match but whose visible runs lost the template donor font properties still fails.
- Before any WPS/Word built-in TOC refresh, extract and lock one real accepted local paragraph instance for:
  - TOC title
  - TOC level 1
  - TOC level 2
  - TOC level 3
- That TOC baseline lock must also record concrete visible metrics rather than only a source path:
  - title font / size / spacing
  - level-1 indentation / line spacing / tab stop
  - level-2 indentation / line spacing / tab stop
  - level-3 indentation / line spacing / tab stop
  - dotted-leader behavior
  - expected TOC page count or occupancy rhythm on rendered pages
- The baseline lock must preserve the existing per-level style-binding rule and expand it into Word/WPS paragraph-dialog metrics. For TOC title and every used TOC level, record template and actual style id/name, outline/list state, direct paragraph-format owner, font family, font size, weight, spacing before/after, line-spacing mode/value, left/right/first-line/hanging indent, tab-stop position, and leader type.
- A TOC that is "style-bound" but whose dialog metrics show smaller font size, reduced line spacing, collapsed before/after spacing, changed indentation, or shifted tab stops is still a failed TOC repair. Treat the failure as an enforcement/evidence failure, not as absence of the original style-binding rule.
- The source of that locked TOC baseline is constrained:
  - first choice: the active school template or teacher-approved sample
  - second choice: another already accepted local manuscript whose TOC has been visually confirmed against the template
  - last resort only: the current working manuscript, and only if its TOC has already been proven visually template-aligned in the current run
- If the user reports that the current manuscript's TOC style is already wrong, do not extract the restoration baseline from that same manuscript.
- Treat those TOC paragraph instances as the post-refresh restoration baseline rather than as optional reference material.
- A built-in TOC refresh is not accepted as complete until the refreshed TOC has been restored to that locked baseline in the same pass.
- Required post-refresh restoration targets:
  - TOC title font, size, weight, alignment, and spacing
  - TOC level-by-level indentation
  - dotted tab leaders
  - right-aligned page numbers
  - per-level font and spacing differences
- If the approved sample TOC spans more than one rendered page, the restored TOC must also preserve the approved occupancy rhythm unless the content growth explicitly requires another page.
- A refresh that keeps correct entries and page numbers but collapses the TOC into a denser one-page default layout is still a failed TOC restoration.
- A rendered TOC that is visibly too dense, squeezed into a different page count, uses the wrong level hierarchy weight/size, or shifts indentation/leader/page-number geometry away from the template is a hard failure even if the text, field, font chain, and page numbers are otherwise correct.
- If the refresh produces a semantically correct TOC that visually falls back to default app styling, the TOC repair still fails.
- If the refresh produces semantically correct entries but restores them to the pre-refresh wrong local look instead of the template look, the TOC repair still fails.
- If the builder styles the TOC before verifying that the TOC content is clean, treat that TOC pass as failing and restart from TOC generation instead of continuing local formatting tweaks.
- If the TOC page and the first body chapter page collapse onto the same rendered page after a rewrite, treat front-matter / TOC repair as structurally failing even if the TOC strings and page numbers look correct in the DOCX text layer.
- Treat the TOC title, TOC level-1 entries, TOC level-2 entries, TOC level-3 entries, and TOC level-4 entries as template-owned, independent formatting classes. Do not merge the TOC title into the first entry line, do not let one TOC level inherit another level's final visual style, and do not reuse body heading or body paragraph styles as final TOC styles.
- A TOC repair fails if the title and levels are separated only by indentation while still sharing one undifferentiated font-style class. Determine final TOC font, size, weight, indentation, and tab-stop settings for the TOC title and every TOC level from the formatting template first.
- TOC page numbers must use the thesis's displayed numbering system rather than raw physical page indexes.
- If the template uses roman numerals for front matter and arabic numerals restarting at `1` for the main body, the TOC must reflect that exact split.
- All pages before the Chinese abstract, English abstract, and TOC block belong to the cover/front-cover area and must not display page numbers unless the school template explicitly overrides this.
- If front matter and main body still share one numbering section, or if the main body does not restart from arabic page `1`, treat the TOC as structurally failing even when indentation, fonts, and dotted leaders look correct.
- Do not accept a TOC where later entries are still showing stale physical-page values after front-matter or section-boundary changes.
- When a teacher sample or screenshot shows the expected TOC, match that visual pattern directly.
- Do not mix multiple TOC repair strategies in one build chain.
- Prefer one clean TOC path: `officecli` high-level inspection first, then either one `officecli raw` repair pass or one office-app refresh path if the live field must be updated.
- If XML repair is used, replace TOC paragraph XML cleanly instead of layering new `pPr` fragments onto old ones.
- Treat TOC refresh as a separate confirmation step when the environment cannot reliably update the live Word field.
- If a live TOC is required by the current run, a visible static TOC fallback is a failed state, not a graceful fallback.
- Before any heading remapping pass, explicitly identify the TOC block boundaries and exclude them from title detection and title insertion.
- A field marked `dirty=true` is not equivalent to a refreshed TOC. If the rendered TOC still shows stale page numbers, the TOC repair is not complete.
- If a helper paragraph or staging paragraph such as `TOC_PLACEHOLDER`, `TOC_PLACEHOLDER_PASS2`, or similar survives into the visible DOCX or rendered PDF, the TOC task has failed even when the field itself exists.
- Legacy template detection rule:
  - do not assume the template TOC is stored as a modern TOC content control or as clean `TOC 1/2/3` style definitions
  - for school `.doc` templates converted to `.docx`, inspect whether the visible TOC relies on:
    - direct paragraph formatting on ordinary paragraphs
    - old TOC paragraph styles with sparse style definitions
    - or a field/content-control block
  - restore the refreshed TOC to the visible paragraph metrics of the template's real TOC entries, even when those metrics come mainly from direct formatting rather than style definitions
- TOC implementation-parity evidence rule:
  - record the template TOC implementation type and the target TOC implementation type before any pass claim
  - record whether the template and target keep the same live-field/content-control/static-template family
  - record the template field instruction and target field instruction when either side contains a TOC field
  - record the exact source and target XML paths for the TOC title, each used TOC level, tab/leader run, and page-number run
  - reject `implementation parity: not checked`, `same enough`, `style-bound`, `visible pass`, or any prose-only substitute
- Tab-stop fidelity rule:
  - treat TOC right-tab position and dotted-leader position as template-owned hard metrics
  - do not accept a TOC whose text is correct but whose page-number tab stop drifts from the template, because this will visibly shift the dotted leaders and page-number column
- Visual geometry evidence rule:
  - treat TOC title position, entry row positions, row spacing, level indentation, leader density, page-number column, and page occupancy rhythm as template-owned hard metrics
  - compare the target and template rendered TOC pages or cropped TOC regions, not only their DOCX paragraph metrics
  - carry the FMT-EVID-007 exact-output binding fields used by final acceptance
  - store the comparison in a review evidence record that follows `assets/review-evidence-template.md`
  - reject any TOC pass claim whose evidence says only that entries, fonts, leaders, or page numbers are present

## 5A. TOC Content-Loss Hard Failures

- Treat TOC content loss as a hard failure, not as a partial formatting issue.
- A TOC pass fails if the TOC title is missing, the TOC entries are missing, multiple entries collapse into one paragraph, entries are replaced by an empty paragraph run, or only a static placeholder remains.
- A TOC pass also fails if the output keeps body headings but drops their corresponding TOC lines, or if the final rendered TOC no longer contains the expected chapter and subsection sequence.
- Do not hand off a thesis sample when the TOC checker has not confirmed both content coverage and visible TOC structure on the exact output path.
- Static TOC page numbers must be computed from rendered heading pages relative to the first body chapter page, then written into the existing template page-number run after the tab. Do not leave every generated TOC entry at page `1`.
- TOC fonts must be compared against the effective template baseline at run level and style level. A theme-font alias such as `宋体（正文）`, `minorEastAsia`, `minorHAnsi`, or any `w:*Theme` font attribute is not equivalent to an explicit template font such as `宋体` unless the template baseline itself uses that same theme attribute.
- A TOC repair is blocked if it reports `passed` after only checking content, bookmarks, or page numbers while skipping TOC title/entry font family, theme-font attributes, indentation, tab stops, dotted leader, line spacing, and page-number column alignment.

### FMT-TOC-001. TOC And Front Matter Must Not Receive Body Images Or Body Content

- The cover, declarations, abstracts, keyword lines, TOC title, TOC entries, and other front-matter protected zones must not receive body paragraphs, body headings, figure captions, screenshots, drawings, chart images, or generated placeholders during thesis mutation.
- A TOC/front-matter page may contain an image only when the official template defines that exact protected-surface image and the transaction records it as an authorized edit to that protected surface.
- A DOCX structure scan must reject `w:drawing`, `w:object`, legacy picture/image data, figure captions, or body-heading styled paragraphs inside the TOC/front-matter zone when they were not explicitly authorized by the template/profile.
- A run that inserts a figure into TOC/front matter is contaminated. Do not continue mutating that output as if it were a valid review copy.

### FMT-TOC-002. Front-Matter TOC Structure Repair Must Be Boundary-Locked

- A helper that removes TOC contamination must first locate the actual TOC title and may delete only non-TOC paragraphs or tables between that TOC title and the first real body level-1 heading.
- Such a helper must not carry a pre-scan `toc_started` state into the mutation loop; cover, declaration, abstract, keyword, and English abstract paragraphs before the TOC title are outside the deletion window.
- Canonical structure repair for TOC contamination, duplicate front-matter page breaks, declaration signature residual pages, abstract-title style repair, level-1 heading baseline replay, or reference-residue cleanup must run through `scripts/repair_thesis_frontmatter_toc_structure.py` or an equivalent canonical helper with the same transaction locks.
- Canonical front-matter/body page-number section repair must run through `scripts/repair_front_matter_page_numbering.py` or an equivalent canonical helper with the same transaction locks. It must write a new review-copy path, bind input/output SHA256 values, replace only `word/document.xml`, fail closed when the Chinese abstract or first body heading cannot be located, and preserve the source package while establishing lower-roman front matter and arabic body restart semantics.
- Rendered page sync for TOC repairs must be audited by `scripts/audit_toc_rendered_page_sync.py` or an equivalent canonical rendered page sync detector before a pass-shaped TOC/front-matter handoff.
- The helper report must bind input/output DOCX paths and SHA256 values, list the explicit operations, list every deleted paragraph/table text excerpt, list any locked reference-residue relocation, and list the exact changed ZIP parts.
- The only allowed DOCX package mutations for this helper are `word/document.xml` and, when a missing abstract title style must be added, `word/styles.xml`.
- The helper must not modify media files, relationship parts, comments, tracked changes, headers, footers, body prose, figure captions, table bodies, bibliography entries, or whole-document section ownership outside the locked target surfaces. The only citation-run exception is the locked `reference-residue` operation, which may move the existing residue paragraph and replace its repair-note wording while preserving its source citation marker runs and hyperlink anchors.
- When replaying level-1 heading baseline, the helper must preserve the chapter opener's existing `w:pageBreakBefore` owner unless a transaction explicitly authorizes a pagination strategy change.
- A level-1 body heading must not carry both `w:pageBreakBefore` and an explicit page-break run. Canonical `duplicate-page-breaks` repair may remove the explicit page-break run while preserving the paragraph-owned `w:pageBreakBefore`; otherwise rendered blank pages can appear before later chapters.
- A first body heading must also not carry an explicit page-break run immediately after a front-matter section-break paragraph; that section break already owns the page transition, so the extra hard break creates a false blank page.
- If the body contains `参考文献` or `致谢`, the visible TOC cache must contain matching tail-block entries in the order `参考文献` before `致谢`; a dotted-leader-only pass without those tail entries is a false pass.
- Static TOC synchronization must collect heading pages through `scripts/collect_heading_pages_word.py` and update cache rows through `scripts/update_static_toc.py` with fullwidth dot normalization, so numbered headings using `U+FF0E` separators are not dropped before tail entries are checked.
- If a body table, caption, figure, or paragraph appears between the TOC title and the first body level-1 heading, the TOC is contaminated. Canonical repair must either relocate an intact caption/table group back to a verified body anchor or fail closed; it must not silently delete body content from the protected TOC range.
- When repairing declaration signature residual pages, the helper may remove only unanchored blank spacer paragraphs inside the declaration block and may tighten declaration-title spacing. It must preserve bookmarks, fields, comments, drawings, section breaks, signature text, and all non-declaration front-matter content.
- After this helper runs, the mutation transaction must verify changed ZIP parts, media/relationship unchanged status, review-artifact preservation, citation-run preservation, TOC contamination removal, abstract style presence, heading baseline state, and rendered front-matter/TOC/body opener review before any pass claim.
