# User Feedback Persistence: Citations And Bibliography

Use this file for durable citation, bibliography, and citation-placement corrections.

If the same thesis-scope complaint also includes visible bibliography label
family issues or page-flow/page-number drift, load
`references/user-feedback/maintenance-and-structure.md` as the sibling router
before acting.

## Enforcement Status

- Every numbered rule in this file is mandatory when this file is loaded for the current subtask.
- Apply these rules together with `references/user-feedback-persistence.md` and the active citation policy files.

## Citation And Bibliography Rules

### FB-CITE-001 (legacy 10). One Sentence Should Use One Superscript Citation Marker (Mandatory)

- When inserting thesis citations as superscripts, do not append multiple independent superscript markers to the same sentence.
- Use the sentence as the citation insertion unit during auto-formatting checks.
- One sentence may cite only one source by default.
- If one sentence appears to need multiple sources, rewrite the prose into separate citation-bearing sentences instead of combining or grouping sources, unless the user explicitly overrides this rule.

### FB-CITE-002 (legacy 11). Thesis Superscripts Must Be Single-Number And Sequential (Mandatory)

- If the user asks for thesis superscript citations, clear old superscripts before rebuilding them.
- One superscript bracket may contain only one number.
- One superscript bracket may correspond to only one source.
- Do not use grouped forms like `[7-13]`, `[7][8]`, or `[7,8]` in thesis superscript output unless the user explicitly overrides this rule.
- Citations must follow the order in which they appear in the text.
- Each sentence may contain at most one superscript citation.
- Citation-marker audits must only apply this numeric-marker rule to bracket tokens whose content is numeric citation syntax. Non-numeric formula or variable brackets such as `[sigma]`, `[MPa]`, or Greek-symbol brackets are formula/prose surfaces and must not be counted as bibliography citation markers.
- Validate these rules on the exact final DOCX before delivery with the canonical citation audit. `BODY_CITATION_ORDER`, `BODY_CITATION_MULTI_NUMBER_MARKER`, or `BODY_CITATION_NOT_SUPERSCRIPT` are blocking failures; a hand-written pass verdict cannot override them.

### FB-CITE-003 (legacy 12). Thesis Font Fixes Must Use The Current Master Manuscript (Mandatory)

- Before applying a thesis-wide font replacement or other global formatting pass, first verify which DOCX is the real current master manuscript.
- Do not run whole-document formatting on an older branch file when references, superscripts, screenshots, or content updates already exist in a newer branch.
- If the latest citation chain exists only in a newer draft, the font-fix pass must use that newer draft as input.
- Treat reference count, superscript presence, screenshot presence, and current output filename as mandatory preflight checks before any global formatting rewrite.

### FB-CITE-004 (legacy 30). Citations Must Not Land On Titles Or Non-Body Surfaces (Mandatory)

- Thesis citations must never be attached to:
- chapter titles
- section titles
- subsection titles
- figure captions
- table captions
- code captions
- Citations belong in body-text sentences only.
- When a strict sentence-based citation rule is active, place the citation marker at the end of the claim-bearing sentence, immediately before the sentence-ending punctuation mark, rather than on an arbitrary term anchor.
- If a citation-placement script cannot identify a valid body-text sentence for the citation, stop and flag the case for manual placement logic instead of attaching the citation to a title, caption, or non-body location.

### FB-CITE-005 (legacy 46). Citation Sources Must Be Real And Publicly Verifiable (Mandatory)

- When adding or repairing thesis references, use only real sources that can be verified through authoritative discovery paths.
- Prefer CNKI and Google Scholar first for academic references.
- If a source cannot be verified through CNKI, Google Scholar, official publisher pages, or official institutional pages, do not keep treating it as a trusted reference by default.
- Replace unverifiable, fabricated, weak aggregator, or low-trust sources when a stronger real source is available.
- Before final handoff, explicitly confirm that every kept reference is real and queryable.
- If the current user requires the bibliography to consist of real papers only, treat the accepted source class as:
  - journal papers
  - conference papers
  - degree theses or dissertations
  - academically published review papers
- Under that paper-only requirement, do not keep the following in the thesis bibliography unless the user explicitly re-allows them:
  - product documentation
  - framework documentation
  - database documentation
  - company financial reports
  - news posts
  - industry release pages
  - general official websites that are not themselves papers
- A source is not accepted as a paper-only bibliography item merely because it is real online content. It must also be queryable as a paper through an academic discovery path such as CNKI, Google Scholar, Semantic Scholar, DOI landing pages, or an official publisher / repository record.
- If a bibliography item is real but not a paper, classify it separately as a non-paper source instead of silently keeping it inside a paper-only literature list.
- When the current run requires paper-only literature, run the dedicated paper-only bibliography audit against the exact review-copy or final-deliverable DOCX before handoff.
- Record the bibliography audit result in a standalone markdown report and in the final acceptance record instead of describing the check only in conversational prose.

### FB-CITE-006 (legacy 47). Thesis Citations Must Use Quoted Source Sentences With Superscript Hyperlinks (Mandatory)

- When the user requires strict thesis citation insertion, do not cite only by paraphrasing.
- Insert citations on the basis of a sentence that quotes or directly carries one traceable source statement from the original paper or source.
- Under this strict rule, one citation sentence should map to one source only.
- Place the citation marker as a superscript before the sentence-ending punctuation mark.
- The superscript citation marker must hyperlink to the corresponding bibliography item in the thesis using an internal document jump path equivalent to the Word-style `Ctrl+left-click` navigation behavior used by TOC entries.
- The visual style of the citation marker must remain body-compatible:
- black text
- no underline
- superscript preserved
- Do not leave thesis citation markers rendered in the default blue-underlined hyperlink look.
- If a sentence cannot be backed by a directly traceable source statement, do not present it as already compliant with this citation rule.

### FB-CITE-007 (legacy 53). Citation First-Appearance Order Must Be Audited Separately (Mandatory)

- Do not assume citation order is already correct just because existing bibliography numbers were preserved.
- Do not assume citation order is correct just because hyperlink jumps, superscript styling, or one-to-one citation mapping were rebuilt successfully.
- Run a separate audit that checks citation numbering against the order of first appearance in the thesis body text.
- Persist that audit as a standalone markdown report on the exact review-copy or final-deliverable DOCX rather than describing the result only in conversational prose.
- During repair, treat this chain as one linked sequence:
- body first-appearance order
- bibliography numbering order
- citation anchor order
- If these three orders diverge, the citation task is still incomplete.

### FB-CITE-008 (legacy 54). Bibliography Formatting Must Be Rebuilt From A Real Template Entry, Then Re-verified At Run Level (Mandatory)

- Do not repair `参考文献` by reusing generic body-text formatting or chapter-heading formatting.
- Use a real bibliography entry paragraph from the school template or accepted sample as the formatting baseline for bibliography entries.
- Copy bibliography paragraph metrics from that real entry sample first, then verify the entry runs separately.
- After copying the template entry metrics, explicitly re-check:
- bold
- italic
- underline
- hyperlink carryover
- If bibliography entries still inherit heading-like emphasis or any other unintended run-level styling after the paragraph-level repair, the bibliography task is still incomplete.
- Bibliography visible-number family must also be locked from rendered evidence before acceptance.
- If the active school template or sample uses dot-style visible bibliography numbering, do not infer the bibliography label family from body citation brackets.
- Bind the label-family decision to the rendered bibliography label geometry report and a separate decision note that explains why bibliography labels are distinct from body citation markers.
- A bibliography entry that uses Word automatic non-bullet numbering is not manual numbering by itself. Validators must not mark such entries as list pollution solely because they carry `w:numPr`; visible manual labels mixed with automatic numbering remain a hard failure.

### FB-CITE-009 (legacy 54A). Bibliography Entry Baselines Must Be Locked Before Deleting Or Rebuilding The Old Bibliography Block (Mandatory)

- Do not delete the old bibliography block first and then sample the "next paragraph" or similar nearby paragraph as the entry-style donor.
- If the bibliography title paragraph remains but the old entries have already been removed, nearby paragraphs can shift and the sampled donor may silently become:
  - the bibliography title itself
  - the following tail-block title such as `致谢`
  - another unrelated heading-like paragraph
- Required safe order:
  1. lock one approved bibliography title paragraph instance
  2. lock one approved bibliography entry paragraph instance
  3. only then remove or rebuild the old bibliography entries
- Do not infer the bibliography entry baseline from mutable positional logic such as `heading_index + 1` after the old block has already been edited.
- If rebuilt bibliography entries collapse into the same style family as the bibliography title, treat that as a bibliography-baseline selection failure rather than as a generic formatting drift.

### FB-CITE-010 (legacy 54B). Bibliography Content Mutations Must Replay The Approved Entry Run Model Before Acceptance (Mandatory)

- Replacing bibliography text is not enough. After any bibliography entry content mutation, replay the approved bibliography entry baseline from the active template or accepted sample onto that exact entry paragraph before acceptance.
- The required replay scope includes both paragraph-level and run-level surfaces:
  - bibliography paragraph style binding
  - indentation and spacing
  - Chinese-body run font model
  - Latin / number / punctuation run font model
  - numbering-token versus entry-text run separation when the sample uses it
- Do not accept a rebuilt bibliography entry that survives only as one homogeneous text run or one homogeneous font family when the approved sample uses separate Chinese and Latin run models.
- After the replay step, rerun bibliography-specific audit on the exact post-edit DOCX. A bibliography audit captured before the content mutation becomes stale immediately.

### FB-CITE-011 (legacy 56). Citation Hyperlink Counts Must Be Verified On The Official Deliverable (Mandatory)

- A thesis citation task is not complete when hyperlink insertion works only on an intermediate probe copy or only inside a builder substep.
- Verify the final official deliverable itself after the last COM save and after the last package-sanitation step.
- The acceptance check must confirm all of the following together:
- the body citation markers still exist
- the markers are superscript, black, and non-underlined
- the internal hyperlink count equals the actual number of cited bibliography items in the body
- A `w:instrText` field instruction such as `HYPERLINK \l "cite_ref_1"` is not visible clickable citation text. It must not count as a valid body citation hyperlink unless a `w:hyperlink` element wraps the visible `[n]` marker run.
- If a later COM save strips or fails to persist the internal citation links, rerun citation-link rebuilding as its own finalizer instead of assuming the earlier citation pass already succeeded.

### FB-CITE-012 (legacy 56A). Internal Citation Links Must Not Expose Anchor Names As Visible Text (Mandatory)

- Do not build thesis citation jumps through a field or wrapper whose visible result becomes the bookmark name, anchor id, or helper label.
- Hard-failure examples include visible text such as:
  - `cite_ref_1`
  - `ref_anchor_3`
  - `bookmark_7`
  - any other internal identifier that appears alongside or instead of `[1]`
- The final visible citation surface must contain only the citation marker text itself, usually `[n]`, while the jump target remains invisible to the reader.
- If a citation-link implementation leaks anchor names on rendered pages, discard that implementation path and rebuild the citation marker with a safer hyperlink representation before handoff.
- The audit scope must scan visible `w:t` text across all `word/*.xml` story parts, not only main-body paragraphs before the bibliography. Headers, footers, text boxes, footnotes, endnotes, comments, and other visible story-part text cannot expose citation anchor helper labels.
- Treat visible numeric, nonnumeric, or bare-prefix helper labels as the same hard failure, including `cite_ref_1`, `cite_ref_`, `ref_anchor_x`, `bookmark_abc`, and concatenated forms such as `[2]cite_ref_1`.
- `w:instrText` field instructions such as `HYPERLINK \l "cite_ref_1"` are not visible pollution by themselves, but their field result zone must not contain any `cite_ref`, `ref_anchor`, or `bookmark` helper text.
- Final acceptance must bind a dedicated citation-anchor pollution audit generated after the last DOCX mutation and after PDF export. The record must include the exact final DOCX SHA256, the rendered PDF path/SHA256 when a PDF is delivered, visible DOCX hit count `0`, field-result hit count `0`, rendered PDF hit count `0`, and a pass verdict.
- A passing citation-anchor pollution audit only proves that internal helper labels are not visible to the reader. It must never be used as a substitute for canonical citation-object validation.
- If the exact final DOCX still fails `scripts/audit_thesis_citations.py` with `BODY_CITATION_NOT_SUPERSCRIPT`, `BODY_CITATION_NOT_HYPERLINK`, `BODY_CITATION_ORDER`, wrong-target, visual-style, or coupled-chain errors, the handoff remains blocked even when DOCX/PDF anchor-pollution counts are all zero.

### FB-CITE-013 (legacy 56B). Citation-Bearing Body Rewrites Must Replay Citation Objects And Re-Audit On The Post-Edit Deliverable (Mandatory)

- When a body sentence with citations is rewritten, do not stop after the sentence text becomes correct.
- Required same-pass recovery order:
  1. rewrite the host sentence content through a citation-aware path
  2. replay the approved citation object model for the touched markers
  3. restore the expected visual style
  4. rerun citation audit on the exact post-edit deliverable
- The required citation-object replay scope includes:
  - isolated citation-marker run or wrapper
  - superscript state
  - internal hyperlink jump
  - black text
  - no underline
  - no inherited bold / italic residue
- If a post-edit DOCX still shows plain body-text `[n]`, blue-underlined citation markers, or citation text merged back into an ordinary body run, treat that as a failed citation replay rather than as a minor formatting miss.
- If an anchor-pollution cleanup removes visible `cite_ref` text but leaves the citation marker as a plain body run or non-clickable text, the same pass must rebuild the superscript hyperlink object and rerun the canonical citation audit before any completion claim.

### FB-CITE-014 (legacy 83). Broken Citation Placeholders Must Be Treated As Citation Corruption, Not As Normal Body Text (Mandatory)

- If a thesis body paragraph shows artifacts such as trailing `0。`, `0，`, or other orphan numeric remnants at a place where a citation marker should exist, treat that as citation corruption.
- Do not leave those artifacts in body text and do not hand-fix only the punctuation around them.
- Recover the intended citation number from the bibliography mapping, nearby citation sequence, or the immediately surrounding drafting history before continuing.
- Citation repair is incomplete if the visible placeholder is removed but the corresponding bibliography numbering is not reconciled.

### FB-CITE-015 (legacy 84). Bibliography Renumbering Must Start From The Current Body Citation Chain, Not From The Existing Reference List (Mandatory)

- When bibliography numbering has drifted, first extract the actual citation numbers used in the thesis body in order of first appearance.
- Build the new bibliography numbering from that body chain, then rewrite the bibliography entries to match it.
- Do not assume the existing bibliography order is still authoritative once the body citation chain has been modified, repaired, or partially rebuilt.
- If body markers and bibliography numbering disagree, the body order wins unless the user explicitly overrides it.

### FB-CITE-016 (legacy 85). Bibliography Rewrites Must Preserve Entry Order Deliberately And Must Not Reverse The List By Insertion Side Effects (Mandatory)

- When rewriting bibliography entries programmatically, do not rely on repeated `insert_paragraph_before` or similar operations without checking the resulting order.
- If the rebuild path inserts entries one by one relative to a fixed anchor, explicitly verify that the final visible order is still `[1]`, `[2]`, `[3]` ... rather than the reversed sequence.
- Treat accidental bibliography reversal as a critical formatting failure because it silently breaks every in-text citation mapping.

### FB-CITE-017 (legacy 86). Citation And Bibliography Passes Must Also Clear Run-Level Bold And Italic Residues (Mandatory)

- Citation and bibliography repair is not complete when the numbers are correct but the citation runs or bibliography runs still inherit bold, italic, hyperlink color, or underline from earlier formatting passes.
- For thesis body citations:
  - keep superscript
  - keep black text
  - remove bold, italic, and underline unless the template explicitly requires otherwise
- For bibliography entries:
  - remove bold and italic from the entry text by default
  - keep academic bibliography font sizing and spacing separate from chapter-heading formatting
- Bibliography font repair must work at run font-slot level, not by whole paragraph or whole-entry language classification:
  - Chinese characters use the locked reference-entry Chinese donor slot, normally `eastAsia`.
  - English letters, digits, DOI, URL, ASCII punctuation, and bracketed reference numbers use the locked reference-entry Western donor slots, normally `ascii`, `hAnsi`, and `cs`.
  - A run containing both scripts is a structure failure unless the active template explicitly proves that mixed run as the donor; by default split or rewrite the run model before handoff.
  - The font/encoding audit must include bibliography font-slot checks against the active template or approved sample, not only mojibake checks.
  - Checking only the first CJK run, first Latin run, paragraph style, indentation, or citation order is not enough to pass bibliography font repair.
- Run-level residue removal is mandatory after any paragraph-level bibliography rewrite.

### FB-CITE-018 (legacy 87). Broken Citation Sentence Starts Must Be Rewritten Into Complete Host Sentences (Mandatory)

- If a thesis body paragraph contains patterns such as `[1]指出`, `[2]表明`, `[3]提出`, or similar sentence-start citation fragments, do not treat that as an acceptable citation style.
- Treat those patterns as broken host sentences rather than as a harmless citation-position variant.
- Required repair rule:
  - rewrite the sentence so the narrative subject and predicate are complete
  - move the citation marker to the end of the claim-bearing sentence
  - keep the citation as a superscript before sentence-ending punctuation
- Do not leave a new sentence beginning with a citation marker after a previous sentence has already ended.

### FB-CITE-019 (legacy 88). Bibliography Coverage Audit Must Compare The Reference List Against The Body-Only Citation Chain (Mandatory)

- A citation pass is incomplete if it only checks that bibliography entries exist and that some citation markers appear somewhere in the document.
- Required audit scope:
  - extract citation numbers from the thesis body only
  - exclude front matter, TOC, figure captions, table captions, and the bibliography block itself
  - compare that body-only citation chain against the bibliography entry list
- Unless the user explicitly asks to keep uncited references, every bibliography entry must have at least one real body citation.
- If a bibliography item exists without a body citation, either add a justified body citation or remove the unused bibliography item.
- Treat the exact body-citation audit report path as a final acceptance artifact, not as an optional troubleshooting note.

### FB-CITE-020 (legacy 89). Citation Repair Must Prefer Source-Of-Truth Rewrites Over Live Paragraph-String Replacement On The Review DOCX (Mandatory)

- Do not use broad paragraph-string replacement on a live thesis review copy as the default way to repair many citation sentences.
- This repair path can silently change paragraph classes, promote body paragraphs into heading-like paragraphs, and pollute the TOC or outline.
- Preferred repair order:
  1. repair the citation-bearing source manuscript or thesis generator text first
  2. regenerate the thesis from that corrected source
  3. only use narrow live-DOCX edits for isolated exceptions after structure review
- If a citation repair pass causes a normal body sentence to appear as a heading or TOC entry, treat the citation pass as structurally failed and restart from the source manuscript instead of trying to patch the TOC locally.

### FB-CITE-021 (legacy 90). Repeated Citation Finalizers Must Not Rebuild Already-Processed Citation Markers In Place (Mandatory)

- If a thesis citation finalizer needs to rerun on the same DOCX, it must first detect whether a visible citation marker has already been converted into a processed citation run or wrapped object.
- Do not add a second hyperlink wrapper, second citation pass, or second marker rebuild around an already-processed citation marker inside the same live DOCX.
- Preferred recovery path:
  1. go back to the clean source-generated DOCX
  2. rerun one clean citation finalizer pass
  3. verify the visible marker text is unchanged apart from the intended superscript styling
- If repeated citation processing starts splitting words, duplicating markers, or fragmenting the host sentence, treat the citation pass as failed and restart from a clean DOCX instead of layering more local fixes.

### FB-CITE-022 (legacy 91). Design And Implementation Chapters Must Separate Project Facts From Literature Claims (Mandatory)

- In chapter 4, chapter 5, or equivalent design and implementation chapters, do not force literature citations onto sentences that primarily describe the real local project.
- The following surfaces are project-fact statements by default and must not receive a literature citation unless the sentence explicitly discusses an external method, paper, or framework claim:
  - real module boundaries in the current repository
  - local route paths, config files, ports, and deployment paths
  - local database tables, object relations, and permission logic
  - local UI pages, buttons, forms, and operating steps
- If a sentence is derived from repository inspection rather than from a traceable external source statement, present it as project truth, not as literature-backed prose.
- If a paragraph mixes project facts and external theory, split the prose so only the source-backed sentence carries the citation.
- If a non-paper source is removed from the bibliography because the current run requires paper-only literature, do not leave the old project sentence or industry-data sentence untouched with the same citation number.
- Rewrite the affected sentence so it is either:
  - supported by a real queryable paper
  - or presented as a project fact without literature citation
- Do not solve a paper-only bibliography requirement by changing the bibliography list alone while leaving body sentences still semantically dependent on deleted non-paper evidence.
- Under paper-only review, this body-dependency recheck is mandatory during both thesis generation and thesis inspection.

### FB-CITE-023 (legacy 92). Do Not Preseed Raw Citation Markers Into Thesis Generator Text (Mandatory)

- Do not treat raw inline markers such as `[1]`, `[2]`, or `[11]` inside generator JSON, drafting notes, or paragraph-source text as a valid finished citation state.
- For thesis generation and chapter rewrites, keep citation planning separate from body-text generation.
- Only materialize citation markers during a citation-aware finalizer that can preserve superscript runs, punctuation-side placement, hyperlink behavior, and first-appearance order.
- If a generation path can only append plain bracket numbers to paragraph strings, stop and treat that path as citation-incomplete rather than shipping those raw markers into the official manuscript.

### FB-CITE-024 (legacy 93). Template Bibliography Samples Must Hard-Fail Final Review (Mandatory)

- If the references block still contains template sample entries, placeholder author names, format notes, or instructional prose such as `作者姓名`, `期刊或杂志名称`, `参考文献（小三号黑体）`, `文献类型标志说明`, or similar school-template scaffolding, the thesis must fail review immediately.
- Do not treat a visually formatted reference list as acceptable when its contents are still sample text from the template.
- Remove template scaffolding before any bibliography formatting pass is considered complete.

### FB-CITE-025 (legacy 94). Non-Empty Bibliography With Zero Body Citations Must Hard-Fail (Mandatory)

- If the thesis contains a non-empty bibliography but the body-citation audit reports zero real body citations, treat that as a hard failure.
- Do not hand off a thesis where the bibliography exists only as a document-ending reading list with no corresponding in-body citation chain.
- Under this rule, a bibliography is only considered valid when at least one real body-text sentence cites it and the full chain can then be audited for order and coverage.

### FB-CITE-026 (legacy 95). Any Bibliography Change Must Trigger Same-Pass Body Citation Synchronization (Mandatory)

- Do not treat bibliography editing as an isolated tail-block task.
- If the run adds, removes, replaces, reorders, or renumbers any bibliography entry, the thesis body citation chain must be re-audited in the same pass.
- Required same-pass checks after any bibliography mutation:
- body citations still point to the intended sources
- any newly added kept reference is either cited in the body or explicitly removed again before handoff
- any removed or replaced reference no longer leaves a stale citation number, stale claim dependency, or broken citation marker in the body
- bibliography numbering still matches body first-appearance order when numbering is in scope
- Do not hand off a thesis where the reference list changed but the body still reflects the old citation mapping.
- Treat `bibliography updated, body citations later` as a hard workflow failure rather than an acceptable intermediate completion state.
- A bibliography mutation is also incomplete when the visible `参考文献` block changes page start, loses its first-page opener, or collapses onto the wrong page sequence after export. The rendered references page and the first continuation page, if any, must be rechecked in the same pass.

### FB-CITE-027 (legacy 96). Any Post-Audit Citation-Bearing Body Mutation Makes The Previous Citation Audit Stale (Mandatory)

- If a thesis run mutates citation-bearing body paragraphs, bibliography numbering, bibliography entries, or citation wrappers after a citation audit report has already been generated, treat that older citation audit as stale immediately.
- This stale-audit rule applies equally to:
  - ordinary paragraph rewrites
  - `word/document.xml` body replacement
  - paragraph-string replacement
  - live DOCX sentence patching
  - bibliography-only rewrites that can change citation mapping
- Do not keep carrying the old citation audit path into the task record, acceptance record, or handoff once the manuscript has changed again.
- Before handoff after such a mutation:
  1. rerun `scripts/normalize_thesis_citation_chain.py` when citation markers, bibliography numbering, or citation wrappers may have drifted
  2. rerun `scripts/audit_thesis_citations.py` on the exact final or exact review-copy DOCX that will be handed off
  3. replace the old citation audit evidence path with the new report path
- A citation audit report whose `document path` names a different DOCX from the exact rendered deliverable is invalid evidence, even if the report itself says `pass`.

### FB-CITE-042. References Pagination Must Be Treated As Part Of The Bibliography Surface (Mandatory)

- A visible numbering fix that still drifts the bibliography opener or the continuation page is not done. Treat the tail block as one reflow unit and rerender it whenever the bibliography page start or continuation page shifts.
- Do not treat references pagination as a side effect of citation repair.
- The `参考文献` title paragraph, the first reference entry, and any continuation page are one coupled tail-block surface. If the opener page changes, page count changes, or the block merges into the prior tail block, the bibliography repair is still failing even if the entries themselves are correct.
- After any bibliography or citation pass, inspect the rendered references page plus the next page, and compare the opener page against the locked donor. A page-fit failure is not cured by correct numbering alone.

### FB-CITE-043. Citation Hyperlinks Must Resolve To Bibliography Entries, Not Cover Or Front Matter (Mandatory)

- A citation hyperlink is not valid just because it is clickable.
- The internal jump target for each body citation must resolve to the matching bibliography entry inside the `参考文献` block, not to the cover, abstract, TOC, header, footer, or any other front-matter surface.
- If a body citation anchor points to a bookmark that lives before the bibliography opener, or to a bookmark that is not inside the bibliography entry range, treat the citation chain as failed even when the visible marker looks correct.

### FB-CITE-028. Citation Normalizers Must Preserve Bibliography Paragraph OOXML (Mandatory)

- A citation or bibliography finalizer may move existing bibliography paragraphs to synchronize first-appearance order, but it must not rebuild an entry by deleting all existing runs and writing the whole reference as one homogeneous text run.
- The only default in-place bibliography mutations allowed inside `scripts/normalize_thesis_citation_chain.py` are:
  - replacing the visible leading bracket number, such as `[8]` to `[1]`
  - adding or refreshing invisible internal bookmark anchors needed by body citation hyperlinks
- Any bibliography content change beyond the leading bracket number must be followed in the same pass by approved reference-entry baseline replay and a bibliography-specific DOCX audit. A citation-order pass alone is not allowed to own reference-entry typography.
- A citation normalizer must fail closed before rewriting a body paragraph when the paragraph is not a plain body-text host. Hard-fail inputs include headings, captions, TOC entries, outline-level paragraphs, paragraphs with fields/drawings/bookmarks/hyperlinks/comments, and paragraphs with multiple non-empty text runs whose direct formatting would be lost by paragraph reconstruction.
- After any normalizer run, the post-edit deliverable must be checked for bibliography run preservation. A result where reference entries collapse into one run per entry is a failed normalizer result unless the locked approved bibliography donor explicitly uses that same one-run model.

### FB-CITE-029. Long Western Bibliography Runs Need A Visible `w:sz` Size Slot (Mandatory)

- A bibliography audit that only checks `w:szCs` can miss WPS/Word rendering shrinkage on long English reference entries.
- When the locked template bibliography donor exposes only `w:szCs` but the repaired entry contains long Western words, the repair helper must mirror the donor `w:szCs` value into `w:sz` for the Western run.
- This size value is not builder-chosen; it is derived from the locked bibliography donor's complex-script size and exists to make the visible Western run render at the intended size.
- The bibliography font audit must reject long Western bibliography runs that have no visible `w:sz` slot after repair.

### FB-CITE-030. Bibliography Font Audit Must Resolve The Effective WPS Font Chain (Mandatory)

- Do not mark reference-entry formatting as passing only because `w:sz` / `w:szCs` exists or because `docDefaults` happens to contain a compatible font.
- The bibliography detector must skip TOC/static page-number rows such as `参考文献21` and must stop the entry block when the `[n]` reference sequence ends, so template instructions after sample references are not treated as bibliography entries.
- When a template entry has no direct `w:rFonts`, lock the bibliography font policy from a real template instruction or approved sample instance before repair. If neither exists, fail instead of guessing.
- Reference-entry audit must report the direct run font slots and the effective font chain through character style, paragraph style, basedOn, docDefaults, and theme aliases.
- A final reference run may pass only when Chinese text is locked to the approved East Asian font and Western letters, numbers, DOI/URL text, and bracketed reference numbers are locked to the approved Latin font through direct run formatting or an approved style. Falling through to theme/default ownership such as WPS body/default font is a failure.
- Rendered-page review of the references page remains required because XML font-slot equality alone can still miss visual density and pagination compression.

### FB-CITE-031. Bibliography Size Names Must Match WPS Exactly (Mandatory)

- Do not accept an approximate point size for a Chinese named-size requirement.
- If the active requirement says `五号`, the bibliography entry runs must use exact DOCX half-point value `21`, equivalent to 10.5 pt, and WPS must show the named size `五号` when the reference text is selected.
- A WPS toolbar display of `10`, `10.5`, `11`, or any other numeric/nearby value is not an acceptable substitute for the named-size requirement when the user or template calls it `五号`.
- A reference-entry audit that is run under an explicit named-size requirement must fail any `w:sz` or `w:szCs` value other than the exact mapped half-point value. For `五号`, w:sz=22 / w:szCs=22 is a failure even if it looks close.
- The repair helper must record the expected named size and half-point value in its report whenever it forces a named-size correction.

### FB-CITE-032. Bibliography Font Pass Requires Positive Evidence And WPS Named-Size Coverage (Mandatory)

- A bibliography font audit pass must not be inferred from `result: pass` or `bibliography font-slot checks: pass` text alone.
- The audit report must bind to the exact DOCX by SHA256, bind the reference/template DOCX by SHA256, and record the active size-policy source.
- A pass report must record positive coverage counts for bibliography entries and checked runs. A zero-hit report with no entry/run coverage is stale or under-scoped evidence, not a pass.
- When the size-policy source is an explicit named-size requirement such as `五号`, the final gate must require WPS UI evidence for every bibliography entry, not only entry `[1]` or a sampled range.
- The WPS evidence must record the exact DOCX SHA256, WPS-displayed named size, expected half-point value, checked entry count, per-entry verdicts, and overall pass verdict.
- WPS named-size evidence collection must recognize every approved bibliography label form used by the active template or final manuscript, including bracket labels like `[1]`, manual dotted labels like `1.`, fullwidth dotted labels like `1．`, Chinese comma labels like `1、`, and Word automatic numbering labels. A zero-entry WPS evidence file is a failed detector run when the final references block visibly contains entries.
- If the official template contains mixed or contradictory reference-entry donor sizes but the school/document requirement and WPS evidence prove the selected named size, the font audit may use the explicit named-size evidence to override only the conflicting donor size while still comparing font families, run splitting, numbering, and other reference-entry formatting.
### FB-CITE-033. Bibliography Entry Indentation Must Be Verified Structurally And On Rendered Geometry (Mandatory)

- Reference-entry indentation cannot pass from DOCX paragraph properties alone.
- When the active template or approved sample uses Word numbering (`w:numPr` / numbering definitions) for bibliography labels, a repair cannot pass by copying only `w:ind left/hanging` values onto manual `[n]` text. It must either preserve/replay the numbering model and remove duplicate manual labels, or produce rendered geometry evidence proving the manual-label fallback is visually equivalent for every entry, including two-digit labels.
- After any bibliography content, numbering, citation, font, or tail-block repair, run both:
  - reference-entry paragraph baseline comparison against a real template or accepted sample entry
  - rendered PDF geometry comparison against the template or accepted sample references page
- The rendered check must compare the left x position of every visible `[n]` entry start against the locked baseline with a numeric tolerance.
- Template instruction rows, TOC-like static page-number rows, reference title text, acknowledgement text, and appendix text must be excluded from the bibliography baseline.
- A final references pass fails if any entry has extra left indent, hanging indent drift, body-style first-line indent residue, or a rendered `[n]` x position shifted right from the template baseline.
- When the locked school template has conflicting bibliography-label evidence, for example prose rules mention bracketed Arabic numbers but the later example block shows `1.` labels, do not rely on the example block alone. Create a label-family decision record from the rule prose, rendered example, and current user correction before repair.
- If the user explicitly reports that visible bibliography labels are still wrong and the official prose rule supports bracketed numbers, the repair must convert the bibliography labels to visible `[n]` form while keeping the paragraphs ordinary text paragraphs. Do not reintroduce Word automatic numbering, `w:numPr`, hanging indent, or body-style first-line indent to obtain the brackets.
- If a run deliberately keeps Arabic-dot bibliography labels because the template and current user both accept that family, the same indentation check applies to visible `n.` labels. Do not impose `w:ind left/hanging` list-style formatting unless the template sample shows hanging continuation lines; if the sample wraps continuation lines back to the label margin, bibliography entries must remove hanging indent and be audited with plain-label geometry.
- Bibliography label repair acceptance must include a rendered PDF geometry check over the exact final PDF: every visible bibliography label `[1]` through `[N]` or `1.` through `N.` must be detected in order according to the chosen label-family decision, and their left `x0` positions must be internally consistent with zero or explicitly justified tolerance. A structure-only audit that reports `60` entries but does not prove rendered label alignment is incomplete.
- `sample_self_check` must receive the reference PDF when available so this rendered geometry detector cannot silently become `not-applicable` during canonical thesis builds.

### FB-CITE-034. Citation Superscripts Need Source-To-Final Run Preservation Evidence (Mandatory)

- A citation audit pass must bind to the exact final DOCX path and SHA256. A report with only `document path` and `result: pass` is not enough for thesis handoff.
- When a body paragraph containing citation markers is rewritten, the run must first inventory the source citation markers, then compare the post-edit final DOCX against that inventory.
- Required citation-run evidence includes paragraph id or paragraph index, marker text, run index, `w:vertAlign=superscript`, font size, color, underline, bold/italic residue, hyperlink or bookmark host, punctuation-side placement, and first-appearance order.
- The comparison must be occurrence-level, not citation-number-level. If a paragraph contains two visible `[14]` markers and only one remains superscript or hyperlink-hosted, the final citation surface fails even though the same citation number still has one valid run elsewhere.
- If the source DOCX lacks stable paragraph ids, paragraph index alone is not enough to prove occurrence continuity. The source-to-final diff must also compare a normalized host paragraph text digest and fail when the same paragraph index keeps a valid-looking marker after the host sentence has changed.
- A source citation marker may be renumbered or moved only through a citation-lane controlled-change ledger that explains the intended mapping and proves the final bibliography chain remains synchronized.
- A final thesis must fail if citation markers are converted into plain body text, merged into a rewritten sentence run, lose superscript, gain hyperlink-blue or underline styling, or are audited only before a later DOCX mutation.

### FB-CITE-049. Final Citation Audit Report Is Non-Substitutable Evidence (Mandatory)

- When the exact final DOCX contains body citation markers or a references block, final acceptance must bind `citation audit evidence path`, `citation audit final DOCX SHA256`, and `citation audit source-to-final run diff path`.
- The evidence path must be produced by `scripts/audit_thesis_citations.py` against the exact final DOCX after the last mutation. Summary prose, manual verdict text, screenshot review, bibliography count checks, hyperlink count checks, or source-to-final run diffs cannot substitute for this report.
- The gate must recompute the citation audit on the exact final DOCX and fail if the report is missing, stale, targets another DOCX, has a different SHA256, carries any error code, or if the recomputed audit fails.
- The same non-substitutable audit owns both visible citation superscript preservation and first-appearance order. `BODY_CITATION_NOT_SUPERSCRIPT` and `BODY_CITATION_ORDER` must block final acceptance even when `body citation superscripts preservation verdict` or `citation-reference coupled-chain verdict` is pass-shaped.

### FB-CITE-035. Citation Superscripts And Reference Entries Are A Coupled Preservation Surface (Mandatory)

- Treat body citation markers and the references block as one coupled chain during thesis format repair.
- Before any mutation that can touch body text, styles, fields, hyperlinks, bookmarks, references, or section content, freeze:
  - body-only citation marker inventory
  - citation marker run properties, especially superscript, font size, color, underline, and hyperlink/bookmark host
  - first-appearance order
  - bibliography entry count, numbering, order, and source text
  - bibliography title and entry paragraph/run baseline from the active template or official requirement
- After mutation, compare the final DOCX against that freeze. The citation chain fails if any marker loses superscript, becomes baseline text, changes to hyperlink-blue/underlined text, leaks an anchor name, moves onto a title/caption/TOC/front-matter surface, or points to the wrong bibliography entry.
- The references block fails if entries are truncated, replaced with template sample text, reordered by insertion side effects, collapsed into one homogeneous run model when the donor uses separate run slots, or left in a generic body font/indentation family.
- A bibliography content change must preserve or deliberately update the body citation chain in the same pass. A reference-only format repair cannot hand off while body citation audit is stale or missing.
- Final evidence must include both a citation audit report and a bibliography baseline/content-format report, bound to the same final DOCX path and SHA256.

### FB-CITE-036. Official Reference Format Requirements Override Generic Bibliography Memory (Mandatory)

- When the official school format document specifies the reference style, use it as the active bibliography policy before falling back to generic memory.
- If the school document requires references after the body, before acknowledgement, numbered by body citation order, and formatted according to `GB/T 7714-2005`, the final bibliography audit must check those conditions explicitly.
- If the school document specifies the reference title font, title indentation, entry Chinese font, entry Western font, or entry size, those fields must be locked in the bibliography baseline profile and verified on the final DOCX.
- Do not accept a references block only because entry count, citation order, or DOI/source verifiability passes. Content standard, entry ordering, title formatting, entry paragraph metrics, run font slots, and rendered geometry are separate required checks.
- Template sample bibliography entries, instructional text, format notes, and placeholder author/title text must be removed before final acceptance. A visually styled but semantically placeholder bibliography is a hard failure.

### FB-CITE-037. Reference Content Mutations Must Be Exact And Citation-Preserving (Mandatory)

- Replacing one bibliography entry to satisfy a teacher comment is a citation-surface mutation, not a plain text find-and-replace.
- A bibliography content mutation must use an exact planned old/new text map, bind the source and final DOCX path/SHA256, and prove that only `word/document.xml` changed unless a separate transaction explicitly authorizes more.
- The helper may replace the host sentence text and the bibliography entry text, but it must not rebuild citation markers, delete or recreate citation bookmarks, collapse bibliography formatting, accept/remove comments, touch relationships, media, styles, numbering, headers, footers, TOC entries, or page-number fields.
- After replacement, rerun citation and review-artifact preservation checks against the exact final DOCX. A reference-content report is not enough by itself if citation markers, comments, bookmarks, or body-citation run state drifted.
- If the replacement is meant to satisfy a year/source-count comment, the plan must state the required year/source token, and the final evidence must show that token in the replaced entry rather than relying on a conversational claim.

### FB-CITE-038. Bibliography Entries Must Not Combine Manual Labels With Word Numbering (Mandatory)

- A reference entry must use one numbering authority, not two.
- If a bibliography paragraph has Word numbering properties such as `w:numPr`, the visible entry text must not also begin with a manual label such as `[1]`, `1.`, or `1、`.
- A bibliography-format repair that replays paragraph metrics from a numbered donor must explicitly choose one numbering model:
  - automatic Word numbering with the visible manual prefix removed
  - manual visible numbering with paragraph `w:numPr` removed
- If the visible reference text already begins with a manual label, remove the paragraph-level Word numbering property or rebuild the entry through the approved bibliography helper before handoff.
- A reference block that passes citation order but can render as automatic numbering plus manual numbering is a hard bibliography-format failure.
- `scripts/audit_thesis_citations.py` must report `BIBLIOGRAPHY_MANUAL_AND_AUTO_NUMBERING` on the exact final DOCX when such overlap exists.

### FB-CITE-038A. Bibliography Entry Repair Must Not Regress The References Title Layer (Mandatory)

- A bibliography-entry formatting pass may touch the `参考文献` title only to preserve or restore its approved title-layer formatting.
- If a previous title/hard-field repair has already materialized the main title font size on `参考文献`, a later bibliography-entry repair must not replace that run formatting with a smaller entry-donor or body-text run model.
- When the bibliography helper replays paragraph metrics from a template donor, it must rematerialize the title's own style run formatting on the visible title run and rerun the whole-format gate afterwards.
- If the live review copy already has `参考文献` bound to a main-title layer such as `Heading1`, `Title`, or a visible size at or above 30 half-points, the bibliography helper must preserve that title paragraph and must not downgrade it to the template's bibliography-entry style.
- Treat a post-bibliography report such as `references title lacks the template main-title layer/font size` as a regression from the bibliography pass, not as a separate cosmetic issue.

### FB-CITE-039. Body-Style Audits Must Exclude Hidden Field Instructions (Mandatory)

- Word field instructions such as `HYPERLINK \l "cite_ref_1"` are not rendered body text and must not be counted as visible paragraph prose by body-style audits.
- Body-style audits must build visible text from rendered text nodes such as `w:t` and tab nodes, not from raw XML `itertext()` that also includes `w:instrText`.
- This exclusion must not hide real citation leakage. Visible marker text such as `[n]`, leaked anchor labels, blue-underlined hyperlinks, and bookmark names remain citation-surface evidence and must still be audited by the citation gate.
- A final `scripts/audit_docx_body_style.py` report that shows hidden field instructions inside body paragraph excerpts is stale or under-scoped evidence and cannot be used as a final body-style pass/fail decision.
- Body-style font-alias checks must stop before the references and acknowledgement tail blocks. Reference-entry font slots are still mandatory, but they belong to `scripts/audit_docx_font_encoding.py` and bibliography-specific gates, not to body prose line-spacing evidence.

### FB-CITE-040. Bibliography Entry Numbers Must Stay Baseline Text (Mandatory)

- Visible bibliography entry numbers such as `[12]` are reference-list labels, not body citation markers.
- The leading bracket-number token of every bibliography entry must not carry `w:vertAlign=superscript` or `w:vertAlign=subscript`, even when the citation-order audit and bibliography entry count pass.
- This check must run on the exact final DOCX even when no separate reference/template DOCX is supplied. Template font-slot comparison is additional evidence, not the owner of the baseline-number rule.
- A final bibliography audit that reports zero checked reference-entry runs while a `参考文献` block exists is under-scoped evidence and cannot satisfy handoff.

### FB-CITE-041. Bibliography Content Runs Require A Bound Format Model (Mandatory)

- Reference-entry content formatting cannot pass from entry count, citation order, or leading-number checks alone.
- When bibliography entries exist in the final DOCX, the DOCX font/encoding audit must bind a reference/template DOCX or equivalent approved bibliography format model and record a pass `bibliography content-format model checks` field.
- Intrinsic-only checks may still reject obvious defects, but they cannot prove that the content runs follow the school/template model. If no `--reference-docx` or approved model source is available, the bibliography content-format model check must fail closed.
- The model comparison must cover the reference-entry content runs, not only the bracket number. It must detect collapsed entries, mixed run font drift, run-slot loss, size drift, and paragraph/run formatting divergence from the donor.
- Final acceptance for reference or bibliography repair must bind the DOCX font/encoding audit evidence path for the exact final DOCX; a pass-shaped summary without the content-format model source is stale or under-scoped evidence.

### FB-CITE-044. Body Citation Labels And Bibliography Labels Are Separate Surfaces (Mandatory)

- When the official school rule or template says正文 citation markers use bracketed superscripts such as `[1]`, keep the body citation marker text in that form.
- When the same school rule or template shows only bibliography examples as Arabic-dot labels such as `1.`, do not force the reference list to display bracket labels merely because the body marker uses `[1]`; first check the prose rule and current user correction.
- The bibliography entry may display `1.` while still carrying the invisible internal bookmark `cite_ref_1` used by the body `[1]` hyperlink only when the label-family decision record accepts Arabic-dot bibliography labels for that deliverable.
- Citation audits must resolve body `[n]` hyperlinks to bibliography bookmarks rather than relying on the visible bibliography label text being `[n]`.
- A bibliography-label repair is not complete until body citation marker style and reference-list label style are reported separately. In particular, body citations may remain bracketed superscripts while the reference list uses either visible Arabic-dot entries or visible bracketed entries according to the explicit label-family decision record.
- Bibliography audits must evaluate the reference-list visible label style independently from the body citation marker style.

### FB-CITE-045. Citation Occurrence Assignment Must Cover Every Bibliography Entry When Markers Are Available (Mandatory)

- If a manuscript has enough body citation markers to cite every bibliography entry but the first-appearance chain skips entries, for example `[1], [3], [5]...` or repeats only the last item, treat it as citation-chain corruption.
- The canonical repair path is `scripts/normalize_thesis_citation_chain.py --assign-all-bibliography-by-occurrence`, which assigns the first N visible body citation markers to bibliography entries `1..N` by occurrence order and keeps later markers on the last bibliography entry.
- This repair must preserve citation markers as isolated superscript internal hyperlinks and must refresh `cite_ref_N` bookmarks for every bibliography entry.
- The repair must not delete unused bibliography entries to make the citation audit pass. It must either assign existing body markers to cover all entries or fail closed when there are too few markers.
- After occurrence assignment, rerun `scripts/audit_thesis_citations.py` and the bibliography school-requirements audit on the exact post-edit DOCX.

### FB-CITE-046. Bibliography Entries Must Contain Substantive Content After The Label (Mandatory)

- A reference paragraph that contains only a visible label such as `[1]`, `1.`, `1、`, or only automatic Word numbering plus a bookmark is not a bibliography entry for delivery.
- The references block fails when any entry has fewer than eight non-space characters of substantive content after the visible label or automatic numbering marker.
- This content-completeness gate is independent from citation order, entry count, hyperlink target, and reference-entry font checks. Those checks cannot turn a label-only paragraph into a valid reference.
- `scripts/audit_thesis_citations.py` must report `BIBLIOGRAPHY_ENTRY_CONTENT_MISSING` for label-only or content-missing entries on the exact final DOCX.
- `scripts/audit_docx_font_encoding.py` must report `bibliography empty-entry/content completeness checks: fail` instead of treating label-only entries as "no bibliography entries detected" or `not-applicable`.
- `scripts/audit_bibliography_school_requirements.py` must include the empty/content-missing count so source-count and paper-only gates cannot pass with numbered placeholders.
- Final acceptance for reference or bibliography work must bind a citation or bibliography audit report that proves zero empty/content-missing reference entries.

### FB-CITE-047. Locked Bibliography Count Profiles Must Drive The Audit Command (Mandatory)

- When the current user, task book, school rule, or accepted project profile locks a bibliography target, do not leave the bibliography audit on its generic defaults.
- For mechanical-design whole-thesis rebuilds, use the locked profile `--min-reference-count 60 --min-foreign-count 10` unless the current user or school template gives a stronger or different target.
- If the profile also specifies journal, paper-only, or recent-literature floors, pass those floors through the matching audit options instead of only mentioning them in prose.
- A bibliography handoff fails when the report says the required count is met but the command/evidence still used a lower generic threshold.
- When the active template uses Word automatic numbering for reference labels, the school-requirements count audit must treat that automatic numbering model as a valid single numbering authority; it must not force manual `[n]` text only to satisfy a text-extraction detector.
- Final acceptance must record the effective bibliography profile, the audit command thresholds, total reference count, foreign/non-web count, and zero empty/content-missing entries on the exact final DOCX.

### FB-CITE-048. User-Corrected Bibliography Label Family Must Be Enforced On The Final Render (Mandatory)

- Do not confuse body citation markers with bibliography-entry labels. If the school template says body citations use superscript bracket markers such as `[1]` but its reference-list examples and prose say `序号.` / `1.` / `2.`, the bibliography list must use the Arabic-dot family while body citations keep bracket markers.
- When the official school rule, active template profile, accepted sample, or current user's correction selects a visible bibliography label family, the final reference list must use exactly that family for every entry. A later repair must not silently revert the entries to another family, automatic numbering, a mixed `[n]` and `n.` block, duplicate automatic plus manual labels, or label-only paragraphs.
- When the user has already reported the bibliography labels as wrong in the same project family, treat the latest explicit user correction plus the official template evidence as a locked label-family decision. Re-read the official template before choosing between `bracket` and `dot` instead of carrying forward an earlier failed repair assumption.
- If the current user locks the compact bracket family, the required visible form is exactly `[1]内容` through `[n]内容`: the opening bracket, number, closing bracket, and substantive entry content must be contiguous with no intervening whitespace. `[1] 内容`, `[n] 内容`, `1. 内容`, Word automatic numbering plus hidden bracket text, or mixed bracket/dot labels are hard bibliography-format failures for that locked family.
- `label-content spacing none` is a hard constraint when the compact bracket family is locked. It is not a cosmetic preference and must survive template replay, bibliography content replacement, citation normalization, PDF export, and any final formatting pass.
- The bibliography label-family decision record must name: official prose evidence, rendered template/example evidence, current user correction, chosen visible label family (`bracket` or `dot`), and the exact final DOCX/PDF paths being audited.
- Acceptance must include `bibliography label-family decision path` and a rendered final-PDF label geometry check for every visible bibliography label, produced by `scripts/audit_pdf_bibliography_labels.py --expected-style <chosen-family>`. A DOCX text scan that says `reference_count=60` is not enough when the user's complaint is a visible label defect.
- Count-only, citation-only, or source-type audits cannot prove bibliography format correctness. `reference_count`, `unique_citation_count`, `body_unique_citations`, paper/source counts, hyperlink counts, or first-appearance order may support bibliography acceptance, but they do not prove visible label family, label/content spacing, paragraph/run baseline, or rendered label geometry.
- The rendered bibliography-label audit must not rely on a built-in default label family. `--expected-style` is mandatory and must come from the bibliography label-family decision record or an equivalent explicit current-user/template decision.
- When replaying bibliography entry formatting after the label family is locked, use the canonical `scripts/repair_bibliography_entry_format.py --visible-label-family <chosen-family>` path rather than an ad hoc paragraph-string replacement.
- Do not produce a reference list by copying body paragraphs, abstract text, TOC rows, captions, appendix rows, acknowledgement text, template instructions, or any other non-bibliography block into the references style. The enforceable shorthand is: must not paste body paragraphs. Reference entries must be substantive bibliography records, then have the locked bibliography donor/baseline replayed; pasted prose that only looks like a reference paragraph is not an accepted bibliography surface.
- If the current user, official template, or accepted sample specifies a compact visible form such as `[1]内容`, the label/content spacing is part of the locked visible bibliography-label decision. Use `scripts/repair_bibliography_entry_format.py --visible-label-family bracket` with `--label-content-spacing none` and audit the rendered PDF with `scripts/audit_pdf_bibliography_labels.py --expected-style bracket` plus `--expected-label-content-spacing none`; `[1] 内容` is still a failure for that locked compact form.
- The rendered label detector must ignore continuation-line page ranges such as `216.` or `458.` when they are outside the locked `1..N` bibliography label range. Page ranges must not be counted as reference-entry labels or as fallback label-family failures.
- When the current label-family decision explicitly locks compact bracket labels, `scripts/audit_bibliography_school_requirements.py` must be run with `--expected-numbering-style bracket`; it must not fail the final manuscript merely because a template example block also contains `1.` labels. The report must still expose the template-derived style separately so the override is auditable rather than silent.
- If a later template replay or bibliography repair changes the final PDF after that check, the rendered label geometry check is stale and must be rerun. The final PDF must show the chosen family for `1..N` and zero fallback labels from the non-chosen family inside the required range.
- Reference-entry font size is a separate surface from formula, figure-caption, and table-caption size. If the school guide says formulas/captions use 五号 but says body text uses 小四 and does not give a smaller reference-entry rule, do not force bibliography entries to 五号. Use the locked bibliography donor/template prose policy, then rerun the font-slot and rendered-page checks on the final DOCX/PDF.
- Reference-entry visible font-size comparison must use the visible script slot: Chinese and Western runs compare `w:sz` when the donor has it. Do not fail a Chinese or Western reference run only because the donor also carries a different `w:szCs`; `w:szCs` is a complex-script slot and is not the visible CJK/Latin size authority when `w:sz` is present.
- The reference-entry font size is separate from formula and caption size, so formula/caption evidence must not be reused as bibliography-entry font evidence.
- For gate keyword compatibility, visible font-size comparison uses w:sz before w:szCs whenever the donor exposes both slots.
