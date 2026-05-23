# Skill Maintenance Case: Citation Hyperlink And Occurrence Gate

- date: 2026-05-23
- mode: skill-maintenance plus scoped thesis DOCX citation preservation verification
- trigger: user reported that citation superscripts and hyperlinks were lost while the skill gate still allowed the run to pass
- thesis artifact mutation: visible handoff copy only; no thesis prose rewrite

## Root Cause

The old citation acceptance path had several false-pass gaps. It could treat `w:instrText` field instructions that mentioned `HYPERLINK cite_ref_n` as hyperlink evidence even when the visible `[n]` marker was not wrapped by a real `w:hyperlink`. It also relied too much on final-DOCX citation surface checks and citation-number aggregation, so source citation surfaces could be stripped or moved without enough occurrence-level evidence. Reference-heading boundary detection could drift in polluted/mojibake contexts, and no-`paraId` DOCX fixtures could be matched only by paragraph index even when the host sentence changed.

## Durable Fix

- `scripts/audit_thesis_citations.py` now requires visible `[n]` marker runs inside `w:hyperlink` and requires those anchors to target `cite_ref_n` bookmarks inside the bibliography range.
- `scripts/audit_docx_review_artifacts.py` now uses the canonical bibliography boundary, records `paragraph_text_digest`, and rejects no-`paraId` same-index occurrences when the citation host text changed.
- `scripts/validate_skill_gate_record_gate.py` now triggers required citation evidence from source or final citation surfaces and enforces `citation-reference coupled-chain verdict` when citations/references are present.
- `references/user-feedback/citations-and-bibliography.md` records that `w:instrText` alone is not visible clickable text and that no-`paraId` occurrence continuity needs host text digest comparison.
- `references/rule-owner-map.json` links these rules to validator and selftest owners, including router anchors for hard-enforced citation rules.

## Regression Coverage

- `case_body_citation_instr_text_only_hyperlink_rejected`
- `case_body_citation_visible_anchor_leak_rejected`
- `case_body_citation_snapshot_stops_at_references_heading`
- `case_body_citation_occurrence_moved_without_ledger_rejected`
- `case_body_citation_no_para_id_host_changed_rejected`
- `case_gate_source_citation_surface_stripped_rejected`
- `case_gate_citation_reference_coupled_chain_missing_rejected`
- `case_body_citation_diff_source_hyperlink_removed_rejected`
- `case_body_citation_diff_duplicate_occurrence_plain_text_rejected`

## Validation

- `py -3 -m py_compile scripts\audit_docx_review_artifacts.py scripts\validate_skill_gate_record_gate.py scripts\selftest_skill_flow.py` PASS
- Targeted citation/gate selftests PASS, including the two new no-`paraId` and coupled-chain cases
- `py -3 scripts\check_utf8_clean.py --root . --json` PASS
- `py -3 scripts\validate_skill_gate.py --skill-root .` PASS
- Exact output citation audit PASS on the visible DOCX handoff
- Source-to-final review/citation preservation audit PASS
