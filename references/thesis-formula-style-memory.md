# Thesis Formula Style Memory

This file stores durable formula-style rules learned from user-provided thesis samples.

## Current learned style

- Formula blocks should appear centered on the page.
- Formulas should use a restrained academic math style, not presentation styling.
- Variables and metric names should appear in italic math style when the sample shows them that way.
- Fractions should use standard stacked fraction layout with a clear horizontal rule.
- Formula lines should keep generous vertical spacing between blocks to avoid crowding.
- Overall composition should stay clean and monochrome, with no decorative color or boxed background.
- Long formulas should remain visually balanced and centered rather than squeezed to one side.
- Formula numbering should follow the active template or accepted manuscript numbering family rather than an arbitrary automation default.
- When the template already uses chapter-based numbering, keep that numbering family. Unless a current school template explicitly proves a different convention, accepted visible numbering must use the formula-word label family `式(章-序)`, for example `式(6-1)`, rather than bare `(6-1)`, dot-style `(6.1)`, full-width variants, or generic `(1) ... (8)`.
- Formula numbers should remain right-aligned as a dedicated numbering surface rather than being visually fused into the math expression itself.
- Formula numbers must sit at the far right end of the formula line in the rendered page, not merely appear after the equation with a small gap.
- Far-right formula numbering must still remain inside the usable page area. A right tab, formula table, or other numbering surface that extends past the right page margin is a hard layout failure even when the number is visually far to the right.

## Operational rule

- When the user provides a formula sample, treat that visual sample as the primary source for formula appearance.
- If the template does not define formula appearance clearly, use the stored default formula style in this file.
- Validate rendered formula appearance by visual comparison against the sample or template, not only by checking equation objects internally.
- Validate numbering style and numbering scope against the template or accepted sample separately from the equation object's internal correctness.
- If a formula is inserted into an existing thesis DOCX, verify both:
  - the equation is a real equation object or the template-approved equivalent
  - the visible numbering still matches the active template convention

## Hard Constraint

### FMT-FORMULA-001. Formula Authenticity Object Gate (Mandatory)

- Every standalone formula in a thesis DOCX must be a real Word equation object (`m:oMath` / `m:oMathPara`) or an explicitly approved equivalent recorded in the evidence. A paragraph that only looks like a formula through ordinary text, underscores, spacing, `x` multiplication, or right-side numbering is a hard failure.
- Do not invent or apply a generic formula style when an existing thesis template or user-provided formula image is available.
- Template and user-provided formula screenshots take precedence over every default remembered style in this file.
- If a generated formula does not visually match the template or sample, stop and revise; do not continue rolling formula edits through the manuscript.
- Do not use text-based fallback formulas that introduce encoding artifacts into thesis prose when the target document expects template-matching equation presentation.
- Do not silently replace template-style chapter numbering with flat document-wide numbering just because sequential numbering is easier to automate.
- A centered paragraph in Times New Roman or another math-like text font is still a failure if it only visually imitates a formula but is not a real equation object.
- Treat “plain text formula masquerading as equation” as a hard failure even when the visible glyphs look roughly acceptable in PDF export.

## Required default formula surface

- centered academic formula layout with italic math variables and stacked fractions
- template-matching numbering placement and numbering scope

## Additional formula lessons recorded on 2026-04-16

- Formula layout and formula numbering are separate acceptance surfaces; both must match the template.
- If the accepted manuscript already shows chapter-based numbering, preserve that numbering family during later insertions, reformatting, or equation-object recovery.
- A formula pass is still failing if the equations are real objects but the visible numbering changes from chapter-based numbering to generic sequential numbering.
- A formula pass is still failing if the numbering exists but does not reach the line's rightmost numbering position on the rendered page.
- A formula pass is still failing if the numbering reaches a hard-coded right edge outside the active section's usable width. Paragraph-tab layouts must compute the center and right tab stops from the active section page width minus margins, and table layouts must keep their total width inside that same usable width.

## Formula authenticity enforcement chain

- Run `scripts/audit_docx_formula_objects.py` on the exact final DOCX before acceptance generation; the JSON report must use schema `graduation-project-builder.formula-object-audit.v1`, bind to the final DOCX SHA256, and record `math_object_count`, `formula_like_paragraph_count`, and `pseudo_formula_count`.
- Final acceptance must carry `formula object audit evidence path` and `formula object preservation summary` whenever formulas or formula-like paragraphs are present.
- `pseudo_formula_count > 0` is a hard gate failure even if PDF export, screenshots, visible numbering, or centered paragraph styling look acceptable.
- The validator must re-audit the final DOCX instead of trusting the summary text, and selftests must include a fake-pass record where a plain text formula tries to claim formula-object success.
- When a formula-like expression appears inside an otherwise valid prose paragraph, repair must preserve the surrounding prose and replace only the exact formula anchor with an OMML object. Use `scripts/docx_formula_number_table.py --replace-inline-text-formulas` with a JSON formula map for this narrow repair path instead of rebuilding the whole paragraph or leaving the expression as plain text.

### FMT-FORMULA-002. Formula Lead-In Must Be Followed By A Formula Surface (Mandatory)

- A thesis paragraph ending with a formula lead-in such as `下式`, `公式`, `计算式`, `表示为`, or `定义为` plus a colon must be followed by a real formula object or an explicitly formula-like expression before explanatory prose such as `其中` / `式中`.
- If the next non-empty paragraph starts explanation text but no equation appears, the manuscript has a missing-formula defect, not a harmless wording issue.
- The repair options are: insert the intended formula as a real equation object with template-matching numbering, or rewrite the lead-in so it no longer promises a formula that is not present.
- `scripts/audit_docx_formula_objects.py` must report `missing_formula_after_leadin_count` and fail the exact final DOCX when this pattern is present.

### FMT-FORMULA-003. Mechanical Design Thesis Formula Density Gate (Mandatory)

- For mechanical, structural, transmission, hoisting, crane, conveyor, reducer, tooling, fixture, CAD/drawing-package, or similar engineering-design theses, the final manuscript must contain at least 200 real Word equation objects unless the current user instruction, task book, or school template explicitly sets a different numeric bound.
- Formula-density evidence must be bound to the exact final DOCX SHA256. Run `scripts/audit_docx_formula_objects.py --min-formula-count 200 --min-body-formula-count 200` for this project class and record `math_object_count`, `body_math_object_count`, `min_formula_count`, `min_body_formula_count`, and the report path before handoff.
- Plain-text equations, image-only formulas, calculation-table prose, formulas hidden only in CAD drawings, and formulas placed only in an appendix do not count toward the mechanical-design body minimum. The count must come from real OMML formula objects in the thesis body calculation/design chapters unless a current user instruction or school template explicitly authorizes a formula appendix as the primary calculation surface.
- If a CAD/mechanical thesis has fewer than 200 real equation objects, treat the calculation chapter as incomplete and continue content repair; do not close the run by saying the document merely needs later manual formula expansion.
- Earlier near-hundred formula language is superseded for CAD/mechanical drawing theses. If the user says `CAD类`, `图纸类`, `机械设计`, or equivalent and does not give a lower explicit number, set both `--min-formula-count` and `--min-body-formula-count` to at least 200.
- A high-density mechanical formula pass must distribute formulas across the real design/calculation chapters instead of concentrating them in one appendix, one summary table, or one artificial formula dump. The acceptance summary must report both total and body formula counts and must say which high-density threshold was used.
- When the run inserts many formula blocks, do not leave plain-text arithmetic fragments in nearby prose as if they were formulas. Text such as `Gd=1800` or `54.6+28.0=82.6` must be either converted into real OMML formulas or rewritten as ordinary narrative; otherwise it is pseudo-formula contamination even if the real equation count is high.

### FMT-FORMULA-004. Mechanical Formula Placement Must Be In Body Chapters (Mandatory)

- Mechanical-design calculation formulas belong in the正文 calculation, thermal-design, structural-design, or strength-check chapters, not only in appendix drawings, tables, or supporting notes.
- A thesis may still keep calculation details or full drawing sheets in appendices, but appendix material cannot be used as the sole evidence that the body explains the design calculations.
- When a user reports missing body formulas or says formulas should be in正文, repair the body chapter first and run `scripts/audit_docx_formula_objects.py --min-body-formula-count <required>` on the exact final DOCX before handoff.
- Final acceptance and scoped handoff must explicitly report `body_math_object_count`; a pass claim that reports only total `math_object_count` is incomplete for mechanical-design theses.

### FMT-FORMULA-005. Formula Number Cells Must Stay Single-Line And Style-Consistent (Mandatory)

- A right-side formula number is a protected numbering surface. It must not wrap, split, or render as separate visible lines such as `式` on one line and `(5.13)` on the next line.
- If the target formula style uses a borderless table, the number cell must contain exactly one non-empty paragraph for the visible formula number. Multiple non-empty paragraphs, manual line breaks, or separate runs that force a line split are a hard formula-layout failure.
- Do not use schema-invalid `w:noWrap` inside table-cell properties as the way to prevent wrapping. Formula-number no-wrap compliance must come from legal OOXML geometry: a fixed table, a sufficiently wide right number column, exactly one visible number paragraph, and rendered PDF proof that the label stays on one line.
- The visible formula-number run size and font model must stay consistent across all formula-number cells in the manuscript unless the active template explicitly proves a different size for a specific formula class.
- Formula-number family, font family, font size, and alignment must be locked from the active template, accepted sample, or current explicit user correction before repair. If the locked surface requires bare `(chapter-sequence)` labels such as `(6-1)`, that family may be accepted only as the recorded template/user override; otherwise the default `式(6-1)` family remains the fallback. Mixed families in the same manuscript are a hard failure.
- Every formula number must be rendered as one visible label on one physical line, right-aligned at the formula numbering surface, and inside the usable page area. A number that is present in XML but appears under the equation, near the equation instead of at the right numbering position, outside the margin, or split across lines is not accepted.
- A body formula object without any visible formula-number surface is a hard failure. The formula audit must report `formula_number_requirement_verdict=pass`, `formula_number_requirement_issue_count=0`, `strict_formula_number_label_count` greater than or equal to `body_formula_group_count`, and no `formula-number-required-missing` issues; `formula_number_cell_count=0` is not a pass when body formulas exist.
- Rendered PDF evidence is mandatory whenever formula numbering is required. A JSON report produced without `--rendered-pdf` may diagnose equation objects, but it cannot clear final formula-number acceptance for a manuscript that contains body formulas.
- For the current SGB620/80T graduation-project run, the right-side formula number must keep the locked donor's visible font family and visible font size. A number that wraps, splits, or uses a different visible size/style than the donor is a hard failure even if the equation object is real.
- Do not repair formula overflow by shrinking only selected formula-number labels, switching some labels to another font family, changing only two-digit labels, or letting later labels inherit a different size/style. Formula-number typography must be consistent across the full checked set unless a locked template/sample records a deliberate exception.
- Rendered formula-number font checks must measure only actual formula-label lines, such as `式(6-1)`, not ordinary prose lines containing words like `式中` or `公式`.
- Visible formula-number labels must use the exact `式(6-1)` family by default: required `式` prefix, ASCII parentheses, and a hyphen between chapter and sequence. Bare `(6-1)`, dot numbering such as `(6.1)` or `式(6.1)`, full-width parentheses such as `式（6-1）`, and split labels where `式` appears on one line and the number appears on another are formula-number style failures unless the active school template explicitly overrides the family.
- Audit keywords for this family are: bare `(6-1)` rejected; dot `(6.1)` rejected; full-width parentheses rejected; split `式` plus number rejected.
- Do not shrink only the formula-number run to make an overlong equation fit. Repair the equation/table geometry instead: widen the math cell, reduce the equation expression through a template-approved math layout, or move the formula to a valid standalone formula block while preserving the right-side number.
- Formula-number layout acceptance must be bound to the exact final DOCX and final rendered PDF. It must include `formula_number_layout_issue_count=0`, `rendered_formula_label_split_pair_count=0`, and `rendered_formula_label_issue_count=0` from `scripts/audit_docx_formula_objects.py --rendered-pdf <final.pdf>`. A report that contains `formula-number-cell-style-invalid` or `rendered-formula-label-style-invalid` is failing even when the equation objects are real and the number is right-aligned.
- Standalone paragraph-based formula layout must also prove the equation line owns a real center tab stop for the OMML formula surface and a real right tab stop for the formula name/number. Visible tab characters alone are not enough; `scripts/audit_docx_formula_objects.py` must report `formula_paragraph_layout_verdict=pass`, `formula_paragraph_layout_issue_count=0`, `formula_paragraph_centered_count` equal to the checked standalone formula count, and `formula_paragraph_right_number_count` equal to the checked standalone formula count. A formula paragraph with `w:tab` runs but no matching center/right `w:tabs/w:tab` definitions is a hard failure because Word/WPS will fall back to default tab stops instead of centered formula plus far-right numbering geometry.
- Rendered formula-label audit must scan every visible `式(章-序)` label in the final PDF, not only labels that already appear near the page's right side. A label rendered beside the equation body or in the page middle must fail as `rendered-formula-label-not-right-aligned` even if its DOCX paragraph contains the expected text.
- Rendered evidence must cover the actual formula-number labels, not sampled nearby prose or count-only equation evidence. The audit must fail closed when formula numbers exist but rendered extraction cannot prove single-line placement, right-side alignment, consistent visible font size, and the locked label family.
- The final acceptance record must carry a formula rendered-label geometry path, rendered split count, and visible font-size verdict. XML-only evidence cannot close a complaint that was visible in the PDF.
- If formula-number cells exist, `scripts/audit_docx_formula_objects.py` must fail when no final rendered PDF is supplied. Missing `--rendered-pdf` is not a pass-shaped not-checked state for formula numbering because right-side number wrapping and visible font-size drift are rendered defects.
- Acceptance generators must pass the exact final PDF into the formula audit and summarize formula numbering as a rendered-label geometry verdict. A formula-numbered manuscript may not record `not-applicable no formula-numbering repair requested` when formula-number cells are present.
- In high-density OMML calculation chapters, PDF text extraction may map math glyphs to U+FFFD even when the DOCX text layer and the rendered formula layout are correct. This is not a body-text乱码 pass: the DOCX XML/text layer must contain zero U+FFFD, every affected PDF page must contain rendered formula-number evidence, and the replacement-character tolerance in `scripts/sample_self_check.py` must scale from the actual OMML/formula-label count rather than a fixed low cap.

### FMT-FORMULA-006. Calculation Source Formulas Must Be Carried Into The Final Manuscript (Mandatory)

- When the project contains a calculation reference, formula map, design-calculation draft, teacher-provided calculation sheet, or earlier locked calculation baseline, formula repair must first create or select a source formula inventory before adding or approving body formulas.
- The final manuscript cannot pass by merely adding generic high-density equations. It must prove that the source calculation formulas have been carried into the final DOCX, with source formula numbers mapped to visible final formula numbers and real OMML formula objects.
- If a source formula inventory exists, run `scripts/audit_docx_formula_source_coverage.py` on the exact final DOCX with every locked `--source-map` and require complete source coverage unless the current user explicitly excludes a source formula.
- When the user says the reference material already contains calculation formulas, the reference material itself must be locked as a formula source before handoff. For DOCX references, run `scripts/audit_docx_formula_source_coverage.py` with `--source-docx <reference.docx>` in addition to any generated `--source-map`; do not satisfy the source-coverage rule only with an internally generated formula batch.
- A final-delivery copy in the visible project handoff folder must be re-audited after promotion. A pass on a hidden work-copy does not clear this rule when the visible `最终交付` DOCX/PDF still contains stale pseudo-formulas or is missing the reference formulas.
- Mechanical-design formula expansion may add additional derived checks after the source formulas are covered, but those additions must not replace, hide, or renumber away the formulas already present in the reference calculation material.
- A pass claim is incomplete when it reports only `math_object_count` or body count and omits source coverage evidence. Acceptance must bind the source-map paths, source formula count, matched source formula count, missing source formula count, coverage ratio, and final DOCX SHA256.
- If source formulas are missing from the final DOCX, keep the run open and insert the missing formulas into the relevant body calculation/design chapter instead of saying the document has enough formula objects overall.

### FMT-FORMULA-007. Formula Additions Must Have Nearby Body Explanation (Mandatory)

- A high-density formula repair is not complete when it merely raises the OMML equation count. Added formulas must be embedded in the relevant calculation, thermal-design, structural-design, pressure-boundary, or strength-check body section.
- Do not leave working labels such as `参考计算公式补入`, `公式补入`, or similar staging text in the final manuscript. Such labels are content pollution and a hard failure.
- A formula table or standalone formula surface in the thesis body must have nearby body prose explaining the calculation purpose, variable meaning, value source, or design judgment. A run of formula tables separated only by blank paragraphs and formula numbers is a formula dump even when every formula is a real Word equation object.
- For mechanical-design calculation prose, a nearby explanation is not sufficient if it is only a generic filler sentence. Each formula group must follow the screenshot-like calculation chain: a body paragraph explains why the calculation is needed, the real equation object carries the numbered formula, nearby ordinary body prose explains variables with `式中` or `其中`, the values are substituted or stated with words such as `代入`, `取值`, `计算`, or `得到`, and the paragraph ends with a design judgement such as `满足`, `不超过`, `安全系数`, `裕度`, `校核`, `判定`, or `选取`.
- Repeated boilerplate such as `式(2-1)用于...正文计算，代入值与判定结果接续服务于本节后续校核` is not a valid formula explanation. It has the shape of an audit placeholder rather than a calculation process and must be repaired before handoff.
- Imported source formula numbers such as `#2-1` must not remain inside the visible equation body when the final manuscript already provides its own right-side formula number. The final number surface must be the only visible formula number for that formula.
- Formula paragraphs must not use undefined, imported, or orphan styles. If a formula import introduces a style id that does not exist in `styles.xml`, normalize the formula paragraph to the locked formula style or to the active template-compatible formula paragraph settings before handoff.
- Nearby explanatory prose added around formula blocks must remain ordinary thesis body text. It must not inherit heading, title, caption, keyword, abstract, reference-list, TOC, or bibliography styles, and it must not turn protected labels such as `关键词` or `Key words` into the style used for a body explanation.
- Nearby explanatory prose for formula calculations is still a protected body-text surface. When it mixes Chinese prose with Latin variables, units, digits, identifiers, or ASCII punctuation, the repair must preserve or restore body mixed-script run and font-slot separation instead of writing one collapsed visible run. A formula narrative pass cannot override a failing body-style audit; rerun `scripts/audit_docx_body_style.py` after formula-process rewrites and repair with the canonical body-run helper when `body mixed-script font summary` fails.
- For mechanical-design theses, run `scripts/audit_docx_formula_objects.py --require-formula-narrative` on the exact final DOCX. The final acceptance evidence must report `formula_dump_marker_count`, `formula_without_nearby_body_explanation_count`, `formula_without_calculation_process_explanation_count`, `formula_narrative_style_issue_count`, `orphan_formula_style_issue_count`, and the formula narrative-context verdict.
- For high-density mechanical formula deliveries, formula count is not enough. When the final requirement is around 100 formulas or more, `scripts/audit_docx_formula_objects.py` must also expose formula duplicate-density evidence, including `unique_formula_body_text_count`, `duplicate_formula_body_text_count`, `unique_formula_ratio`, `min_unique_formula_ratio`, and `formula_duplicate_density_verdict`. Repeating a small formula template set until the document reaches 200 formulas is a false pass.

### FMT-FORMULA-008. Raw Math Command Tokens And Formula-Only Pages Are Hard Failures (Mandatory)

- A formula object is not authentic if its visible DOCX/OMML text layer or rendered PDF exposes raw construction tokens such as `sub`, `sup`, `frac`, `sqrt`, `over`, `below`, `above`, `nary`, `lim`, or `eqarr` as the equation body. These are pseudo-formula command words, not academic formulas, even when the paragraph contains `m:oMath`.
- Mechanical formula-density repair must not satisfy the 200-formula threshold by adding one tiny raw-token formula per page. Formula-numbered pages that are near-empty or contain only the header/footer plus one formula label are pagination failures unless a current template explicitly authorizes a full-page derivation sheet and the page is independently allowlisted with rendered evidence.
- Generated formula paragraphs must not inherit chapter-heading/list/page-break behavior. A formula paragraph carrying `Heading1`, `outlineLvl`, `pageBreakBefore`, `keepNext`, or list numbering from a template paragraph is a structural formula failure when it forces formula-only pages or pollutes the body outline.
- `scripts/audit_docx_formula_objects.py` must report `raw_math_command_token_count`, `rendered_raw_math_token_page_count`, and `rendered_near_empty_formula_page_count`. Any non-zero value is a hard failure for final acceptance.
- Source formula maps that use structured segment keys such as `sub` or `frac` must be converted into real OMML structures before insertion. The keys themselves must never appear in final `m:t` formula text or rendered PDF lines.
- When this failure is found after a high-density formula pass, treat the previous formula-count pass as a false positive; remove or replace the affected formula blocks and rerun both DOCX and rendered-PDF formula audits on the exact final output.

### FMT-FORMULA-009. Body Formulas Must Not Be Chapter-Front Formula Dumps (Mandatory)

- Mechanical-design formulas must appear inside the relevant正文 discussion, calculation, design-selection, or strength-check subsection. A visible run labelled like `公式组...计算如下`, `formula group`, `公式补入`, `计算公式汇总`, or similar staging wording immediately followed by several standalone formulas is a formula dump even when every formula is a real OMML object.
- Do not place a pile of formulas at the beginning of a chapter, before the design narrative, or as a chapter-opening batch that readers must interpret without local variable explanation and design judgment. Chapter openers may introduce calculation scope, but the formulas themselves must be distributed into the subsections where the load, power, gear, shaft, weld, bearing, or tensioning calculation is discussed.
- A high-density formula target such as 200 formulas does not override readability. Formula count must be achieved through正文 calculation coverage, not by inserting four-equation blocks under every heading with repetitive labels.
- If the final manuscript contains repeated marker paragraphs such as `公式组2-00-...计算如下：`, repair them by deleting the marker text and merging the formula sequence into nearby正文 prose, or by removing nonessential duplicate formulas while preserving the required calculation chain. Acceptance must report `formula_front_dump_marker_count=0` or an equivalent local audit verdict.
- When a user says formulas should only be in正文, treat appendix-only formulas, chapter-front formula piles, and stand-alone formula batches before the first body explanation as failing surfaces. Rerun the formula narrative audit after repair and bind the exact final DOCX/PDF.

## Additional formula lessons recorded on 2026-04-17

- Formula authenticity and formula appearance are separate acceptance surfaces.
- A formula task is not complete when the manuscript still contains centered text paragraphs such as `F(P)=...` or `R_local=...` in place of real equation objects.
- After repairing a pseudo-formula into a real equation object, verify both:
  - `officecli query ... equation` (or equivalent) can detect the equation object
  - rendered-page review still shows the formula at the intended location with correct spacing
- For inline prose formulas, also verify that the prose before and after the formula anchor is still present and that `scripts/audit_docx_formula_objects.py` reports `pseudo_formula_count=0`.

## Additional formula lessons recorded on 2026-04-19

- In this local environment, WPS COM can detect `OMaths` successfully but may hang on fine-grained formula write operations such as numbering-surface edits.
- When that happens, a narrow `word/document.xml` fallback may be used for numbering-surface repair if and only if the run still preserves the real equation object.
- A locally verified fallback on 2026-04-19 converted one formula paragraph into a borderless three-cell table with:
  - left spacer cell
  - middle cell carrying the original equation object
  - right cell carrying the formula number
- This fallback is acceptable only when rendered-page review confirms all of the following together:
  - the equation object still renders as a real formula
  - the formula remains visually centered in the main equation surface
  - the number stays at the far-right numbering surface
  - the table borders remain invisible on the rendered page
- Do not treat this borderless-table fallback as self-proving. Its acceptance still depends on the rendered page rather than only on package structure.

If the template does not define a stronger fixed style, this stored formula sample is the default required formula reference for thesis completion.
