# Thesis Format Class Review

Use this file when a thesis-format run needs explicit class-by-class checking instead of ad hoc spot checks.

## 0. Enforcement Status

- Every review rule in this file is mandatory when this file is loaded.
- If a touched class is not reviewed under this file's rules, the format run must be treated as incomplete.

## 1. Required Status Values

Every touched class must be recorded as one of:

- `pass`
- `fail`
- `needs-manual-review`

If a touched class is unrecorded, the run is incomplete.

For thesis generation, thesis revision, and thesis format repair, also create a mandatory thesis surface inventory before mutation. This inventory is broader than "touched classes": it records every protected thesis surface that could otherwise be silently skipped.

Each inventory row must use one of these statuses:

- `present-active`
- `present-unchanged-reviewed`
- `not-present`
- `not-applicable-with-reason`
- `blocked`

Each row must also record surface id, baseline or donor, evidence path, final verdict, and reason. A blank reason, missing evidence for a present surface, missing pass verdict for a present surface, or any `blocked` status means the thesis run is not deliverable.

## 2. Required Format Classes

### FMT-CLASS-001. Class Review Must Cover Every Present Thesis Surface (Mandatory)

When a thesis-format run claims whole-thesis, full-paper, 1:1 template alignment, template-aligned completion, or responds to a user-reported format problem, the class review surface list must be at least as broad as the final acceptance surface-face parity list.

Check these separately when they are present, touched, user-reported, or can be affected by pagination:

- cover field labels and values
- title/front matter same-page groups
- declarations and other front-matter blocks
- Chinese abstract title
- Chinese abstract body
- Chinese keyword label
- Chinese keyword content
- English abstract title
- English abstract body
- English keyword label
- English keyword content
- TOC title
- TOC level-1 entries
- TOC level-2 entries
- TOC level-3 entries
- TOC implementation family and live-field/content-control structure
- TOC tabs, dotted leaders, and page-number column
- level-1 headings
- level-2 headings
- level-3 headings
- body text
- code titles
- code blocks
- formula objects
- formula numbering
- figure holder paragraphs
- figure captions
- figure follow-up explanatory paragraphs
- table titles
- table family/border model
- table-cell text
- body citation superscripts
- review comments and tracked changes
- references title
- reference entries
- acknowledgement title
- acknowledgement body
- appendix title
- appendix body
- headers
- header left fixed/school token
- header right full display string
- header chapter-number component
- footers
- page numbers
- whole-document pagination
- section breaks
- chapter-start pagination
- tail-block pagination

For each class, compare the final manuscript against the locked active template/profile for style binding, direct run fonts, font slots, font size, bold/italic/underline, color, alignment, indentation, line spacing, spacing before/after, tabs, keep/page-break rules, table borders, image holder geometry, and rendered page placement where applicable. A visible-text-only check is not a class review.

### FMT-CLASS-003. Mandatory Thesis Surface Inventory Must Cover Front Matter And End Matter (Mandatory)

Before any thesis DOCX mutation, the format lane must externalize a mandatory thesis surface inventory with at least these rows:

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
- `figure_table_captions_and_holders`
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
- `review_comments_and_change_marks`

The inventory is mandatory even for a local repair because local edits can move or contaminate sibling surfaces. Cover, abstract, keyword, TOC, and references rows are critical deliverable rows: if they are absent, unverified, or marked not applicable, the run must stay failed or audit-only rather than handing off a completed thesis.

The front-matter coverage matrix must prove cover, declaration/title front matter, Chinese abstract, Chinese keywords, English abstract, English keywords, and TOC visual structure. The end-matter coverage matrix must prove references title, reference entries, acknowledgement when present, appendix when present, header/footer, and page numbers. The final acceptance record must name both matrix paths and must carry explicit verdict fields for cover style, abstract and keyword surfaces, TOC visual baseline, reference-entry format, and appendix format.

The high-risk thesis format surface matrix is separate from the general inventory. It must include cover style, Chinese abstract title/body/keyword line, English abstract title/body/keyword line, TOC title/entries/dotted leaders/page-number column, TOC implementation family/live-field structure when the template exposes one, body heading levels, header full display string including chapter-number component when the template uses one, references title/entries, acknowledgement title/body, and appendix title/body. The final acceptance record must name the matrix path and carry an overall high-risk thesis format surface verdict. If any of those rows is missing, unchecked, marked blocked, or contradicted by rendered evidence, the run must not be handed off as fixed.

If the source manuscript or comment-carrier DOCX includes review comments or tracked changes, add `review_comments_and_change_marks` to the high-risk matrix and prove source-to-final preservation or explicit user-approved removal. A final copy may not silently drop comment parts or tracked-change parts while the comments surface still exists in the source.

The abstract and TOC rows are not allowed to be generic page-class rows. `zh_keyword_line` and `en_keyword_line` must include label/content run-split evidence under the row, and each TOC row must map to the actual TOC levels and page-number column used in the final manuscript. Do not write `Chinese abstract passed`, `English abstract passed`, `TOC passed`, or one shared `front matter passed` row as a substitute for the required row set.

Appendix rows are allowed to be `not-applicable-with-reason` only when the active template or source manuscript proves there is no appendix. When appendix content exists, appendix title and appendix body require the same rendered evidence and pass verdict as the other high-risk surfaces.

### FMT-CLASS-004. High-Risk Thesis Format Surfaces Need A Separate Release Gate (Mandatory)

Every thesis generation, thesis revision, and thesis format repair record must include:

- a high-risk thesis format surface matrix path
- a high-risk thesis format surface verdict
- row-level verdicts for cover style, Chinese abstract title/body/keyword line, English abstract title/body/keyword line, TOC title/entries/dotted leaders/page-number column, TOC implementation family/live-field structure when present, body heading levels, header full display string and chapter-number component when present, references title/entries, acknowledgement title/body, and appendix title/body
- evidence records for abstract and TOC rows that confirm the exact protected surface, not only the surrounding page class or a raw screenshot

This gate is required even when the active request names only one paragraph, one heading, or one figure. It exists because these front-matter and end-matter surfaces have repeatedly been omitted by local-only repairs.

### FMT-CLASS-002. Class Review Records Must Name Owner Detector And Acceptance Evidence (Mandatory)

Every reviewed class must record:

- one write owner lane or `read-only`
- the donor or baseline source
- the detector or script used, or the explicit blocker when no detector exists
- rendered evidence path when the surface is visible
- final acceptance evidence field that will carry the result

If a present or user-reported class has no owner, no detector, no donor, or no acceptance evidence field, the run is incomplete and must not be handed off as fixed.

## 3. Table Review

When tables are touched, verify:

- caption position
- caption alignment
- border model
- header-row emphasis
- body-cell alignment
- width behavior
- page-split behavior

If three-line-table style is active, verify:

- top border exists
- header separator exists
- bottom border exists
- internal vertical separators match the locked active table variant
- outermost left border state matches the locked active table variant
- outermost right border state matches the locked active table variant
- body-row horizontal separator state matches the locked active table variant
- when WPS built-in style is the editor-side authority, judge the final border family from an unselected page view or exported PDF rather than from selected-table editor guides

## 4. Figure Review

When figures are touched, verify:

- intended image is visible
- caption is visible
- caption is outside the image
- image and caption stay paired
- surrounding text does not intrude into the image area

## 5. Heading Review

When headings are touched, verify:

- style binding is correct
- indentation is correct
- numbering display matches the template strategy
- chapter-level headings start on a new page when required

## 6. TOC Review

Verify:

- order is correct
- implementation family matches the locked template when the template uses a live field or content-control TOC
- indentation reflects heading depth
- dotted leaders are correct when required
- page numbers align
- no stale or duplicate TOC blocks remain

## 7. Inspection Order

Use this order by default:

1. front matter
2. main-text hierarchy
3. body text
4. figures and tables
5. end matter
6. final TOC and pagination recheck

## 8. Final Completion Review

Before claiming completion:

1. verify touched classes internally from DOCX state
2. verify touched classes on rendered pages
3. verify TOC after all later fixes
4. verify chapter starts after all later fixes
5. verify every changed figure/table block visually

If either internal verification or rendered-page verification is missing, the format task is not complete.

## 9. Minimal Check Record Template

Use this structure when logging a real run:

```md
- class: level-2 headings
  status: pass
  owner: format
  donor: active template level-2 heading paragraph
  detector: body-style audit + rendered page review
  acceptance_evidence: heading baseline preservation summary + touched-page review evidence
  evidence: outline + rendered-page review

- class: table style
  status: pass
  owner: table/format
  donor: active template table family
  detector: table-local structure evidence + rendered table baseline comparison
  acceptance_evidence: table authority evidence paths + table rendered baseline comparison evidence paths
  evidence: rendered table pages + DOCX border inspection
```
