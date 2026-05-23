# Skill Maintenance Case: Reference Pagination Loss

- date: 2026-05-14
- mode: skill-maintenance
- trigger: user reported that reference-section pagination was lost after repeated thesis repairs and had to be manually restored
- thesis artifact mutation: none

## Root Cause

The existing pagination gate verified section/page-number topology, blank/near-empty rendered pages, and a generic tail-block page map. It did not require executable proof that the `references` opener still had a single pagination owner after mutation. A repair could therefore remove `w:pageBreakBefore`, a preceding page break, or a preceding section break while still producing pass-shaped whole-document pagination evidence.

## Durable Fix

- `EXEC-MAINT-073` now requires `references_page_found=yes`, `references_fresh_page_verdict=pass`, and `references_opener_owner_evidence` in the tail-block opener map.
- `sample_self_check.py` emits `tail-block.pagination-contract` and fails when the `references` opener lacks exactly one pagination owner.
- `inspect_docx_pagination_structure.py` binds tail-block page-map evidence to that detector instead of accepting generic page-map text.
- `generate_thesis_acceptance_record.py`, `validate_skill_gate_record_evidence.py`, and `validate_skill_gate_record_gate.py` reject missing, stale, failed, or pass-shaped reference pagination evidence.
- 2026-05-20 hardening adds prior-block separation proof: the tail-block opener map must record `references previous content physical page=...` and `references_prior_block_separation_verdict=pass`, and the sample detector fails when the previous content page is missing or the same as the formal references opener page.

## Regression Coverage

- `case_sample_self_check_tail_block_pagination_loss_rejected`
- `case_sample_self_check_tail_block_previous_page_merge_rejected`
- `case_acceptance_generator_reference_pagination_loss_rejected`
- `case_acceptance_generator_reference_prior_page_token_missing_rejected`
- `case_acceptance_generator_missing_tail_block_pagination_detector_rejected`
