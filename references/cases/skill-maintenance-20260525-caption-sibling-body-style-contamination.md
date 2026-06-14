# Skill Maintenance Case: Caption/Table Sibling Body-Style Contamination

- date: 2026-05-25
- mode: skill-maintenance
- rule owner: `FB-LAYOUT-072`
- affected validator: `scripts/audit_docx_body_style.py`
- regression coverage: `scripts/selftest_skill_flow.py::case_audit_body_style_caption_sibling_pollution_rejected`; `scripts/selftest_skill_flow.py::case_audit_body_style_caption_sibling_body_valid`

## Defect

Repeated thesis DOCX repair sessions left正文 paragraphs after figure/table captions or table objects rendered like captions/table titles. The text content was body prose, but the paragraph kept caption/title style state such as center alignment, caption line spacing, zero first-line indent, `keepNext`, or a `Caption` style chain.

## Repair

`FB-LAYOUT-072` now states that the paragraph after a formal figure caption, table title, figure holder, or table object is body prose unless it is itself a formal caption/title. `scripts/audit_docx_body_style.py` adds `caption_sibling_body_contamination_records`, which rejects caption/title/heading style chains and caption-like paragraph/run metrics on that sibling body paragraph.

## Acceptance

The negative regression fixture builds a table-title-plus-table block and a following explanatory paragraph that wrongly uses the caption family. The body-style audit must fail with `after table object keeps caption/title formatting`.

The positive regression fixture builds the same table-title-plus-table block followed by a correctly formatted body paragraph. The body-style audit must pass so the detector does not overreject valid explanatory prose after tables.
