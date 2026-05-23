# Thesis Troubleshooting: Blank Pages And End Matter

Use this file for abnormal blank-page incidents, reference/end-matter drift, and late-stage content-closure failures.

## Enforcement Status

- Every rule in this file is mandatory when this file is loaded for the current recovery subtask.
- Apply this file together with `references/thesis/thesis-troubleshooting-log.md`.

## Blank Page Root-Cause Rule

- If rendered pages show abnormal blank pages around figures, diagrams, or tables, treat the blank pages as a structural failure, not as a late cosmetic issue.
- Common root causes:
  - repeated blank paragraphs left in front of image-anchor paragraphs
  - floating shapes anchored far from their intended narrative position
  - oversized inline images that force content onto later pages
  - caption/table insertion logic that introduces hidden structural drift
  - duplicated chapter-start ownership, such as a moved body section restart plus an inline page-break run still embedded inside the first chapter title paragraph
- Correct recovery path:
  - inspect the affected page range with paragraph-to-page mapping
  - inspect nearby blank paragraphs, figure-holder paragraphs, inline shapes, and floating shapes
  - inspect the local chapter/title paragraph XML when the blank page appears immediately before a chapter opener, references page, or acknowledgement page
  - if both a section boundary and a title-level forced page break are active for the same opener, remove the redundant owner and keep one verified chapter-start mechanism
  - remove unnecessary blank paragraphs
  - re-anchor or resize the relevant figure block
  - when the user explicitly requests text-based filling, add or relocate real subsection prose near the affected figure/table block instead of using empty paragraphs or purely mechanical image movement
  - rerun rendered-page review before continuing

## Front-Matter Blank Page Recovery Rule

- A long customized title inserted into a template cover table can push the cover section boundary onto the next rendered page, producing a blank page before the declaration.
- Treat this as a rendered front-matter failure even if DOCX structure validation passes.
- Do not accept a pagination evidence record that only compares page counts. It must inspect rendered page images and list `template_blank_pages`, `actual_blank_pages`, `unexpected_blank_pages`, `actual_near_empty_pages`, and `unexpected_near_empty_pages`.
- Correct recovery path:
  - compare the rendered cover/declaration pages against the template
  - inspect blank paragraphs around the cover table, date paragraph, and section boundary
  - remove only unnecessary blank paragraphs or reduce the title/table overflow in the repair copy
  - rerender until the declaration follows the cover without an unintended blank page

## Reference Boundary Rule

- `参考文献` is a terminal top-level section, not a continuation of Chapter 6.
- The formal `参考文献` title must not share a rendered physical page with the previous chapter, conclusion, or other real content block. A repair is accepted only when the evidence records the previous content page, the `参考文献` page, and a passing prior-block separation verdict.
- Do not leave any "reference-discussion" or copied bibliography explanation paragraphs inside Chapter 6 once the final references section is appended.
- Before delivery, scan the body and confirm:
  - Chapter 6 ends with its own conclusion content
  - `致谢` and `参考文献` are separate terminal blocks
  - bibliography entries appear only under `参考文献`
- Also check for duplicate insertion bugs:
  - do not accept a repair only because one PDF export happens to place `参考文献` or `致谢` on a new page; also verify that the title paragraph still has one durable opener owner such as a section boundary or `pageBreakBefore`
  - references accidentally appended once before the terminal sections
  - then appended again under the formal `参考文献` heading
- If duplicate insertion exists, remove the earlier block and keep only the terminal reference section.

## Late-Stage Content Closure Rule

- Near the end of thesis delivery, visible process narration can be a bigger defect than residual spacing issues.
- Typical failure signs:
  - phrases equivalent to "current draft", "later template adjustment", "continue polishing", or "approaching final submission"
  - conclusion pages that read like project status reports instead of thesis conclusions
- Correct recovery path:
  - inspect the experiment tail and conclusion pages directly
  - rewrite them into formal result-analysis and conclusion language
  - only then continue template-level formatting repair
