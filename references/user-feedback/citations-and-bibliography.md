# User Feedback Persistence: Citations And Bibliography

Use this file for durable citation, bibliography, and citation-placement corrections.

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
- Validate these rules before delivery instead of assuming the citation script got them right.

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
### FB-CITE-033. Bibliography Entry Indentation Must Be Verified Structurally And On Rendered Geometry (Mandatory)

- Reference-entry indentation cannot pass from DOCX paragraph properties alone.
- When the active template or approved sample uses Word numbering (`w:numPr` / numbering definitions) for bibliography labels, a repair cannot pass by copying only `w:ind left/hanging` values onto manual `[n]` text. It must either preserve/replay the numbering model and remove duplicate manual labels, or produce rendered geometry evidence proving the manual-label fallback is visually equivalent for every entry, including two-digit labels.
- After any bibliography content, numbering, citation, font, or tail-block repair, run both:
  - reference-entry paragraph baseline comparison against a real template or accepted sample entry
  - rendered PDF geometry comparison against the template or accepted sample references page
- The rendered check must compare the left x position of every visible `[n]` entry start against the locked baseline with a numeric tolerance.
- Template instruction rows, TOC-like static page-number rows, reference title text, acknowledgement text, and appendix text must be excluded from the bibliography baseline.
- A final references pass fails if any entry has extra left indent, hanging indent drift, body-style first-line indent residue, or a rendered `[n]` x position shifted right from the template baseline.
- `sample_self_check` must receive the reference PDF when available so this rendered geometry detector cannot silently become `not-applicable` during canonical thesis builds.

### FB-CITE-034. Citation Superscripts Need Source-To-Final Run Preservation Evidence (Mandatory)

- A citation audit pass must bind to the exact final DOCX path and SHA256. A report with only `document path` and `result: pass` is not enough for thesis handoff.
- When a body paragraph containing citation markers is rewritten, the run must first inventory the source citation markers, then compare the post-edit final DOCX against that inventory.
- Required citation-run evidence includes paragraph id or paragraph index, marker text, run index, `w:vertAlign=superscript`, font size, color, underline, bold/italic residue, hyperlink or bookmark host, punctuation-side placement, and first-appearance order.
- The comparison must be occurrence-level, not citation-number-level. If a paragraph contains two visible `[14]` markers and only one remains superscript or hyperlink-hosted, the final citation surface fails even though the same citation number still has one valid run elsewhere.
- If the source DOCX lacks stable paragraph ids, paragraph index alone is not enough to prove occurrence continuity. The source-to-final diff must also compare a normalized host paragraph text digest and fail when the same paragraph index keeps a valid-looking marker after the host sentence has changed.
- A source citation marker may be renumbered or moved only through a citation-lane controlled-change ledger that explains the intended mapping and proves the final bibliography chain remains synchronized.
- A final thesis must fail if citation markers are converted into plain body text, merged into a rewritten sentence run, lose superscript, gain hyperlink-blue or underline styling, or are audited only before a later DOCX mutation.

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

### FB-CITE-038. Bibliography Entries Must Not Combine Manual Bracket Numbers With Word Numbering (Mandatory)

- A reference entry must use one numbering authority, not two.
- If a bibliography paragraph has Word numbering properties such as `w:numPr`, the visible entry text must not also begin with a manual bracket number such as `[1]`.
- If the visible reference text already begins with `[n]`, remove the paragraph-level Word numbering property or rebuild the entry through the approved bibliography helper before handoff.
- A reference block that passes citation order but can render as automatic `[n]` plus manual `[n]` is a hard bibliography-format failure.
- `scripts/audit_thesis_citations.py` must report `BIBLIOGRAPHY_MANUAL_AND_AUTO_NUMBERING` on the exact final DOCX when such overlap exists.

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
