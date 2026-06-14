# Skill Maintenance 20260524 Whole-Format Template Header

- run id: skill-maintenance-20260524-whole-format-template-header
- trigger: local thesis audit found that the whole-format gate rejected a valid template/static institutional running header and that the repair helper could still write template-incompatible header text
- mode: skill-maintenance with local thesis audit evidence
- changed rule: EXEC-MAINT-076
- changed scripts: scripts/audit_docx_whole_format_gate.py; scripts/repair_docx_whole_format_structure.py; scripts/generate_thesis_acceptance_record.py; scripts/validate_skill_gate_record_core.py; scripts/selftest_skill_flow.py
- changed references/indices: references/user-feedback/maintenance-and-structure.md; references/rule-owner-map.json; FILE-ROLE-INDEX.md; DURABLE-RULE-PROMOTION-AUDIT.md
- root cause: the gate and helper had drifted toward one template's header policy. Static institutional headers on cover/TOC/body sections were treated as leaks, while the helper could synthesize `目 录`, a body chapter title, or a hard-coded school header instead of deriving the current template header.
- durable correction: the gate allows safe static institutional headers while still rejecting body/TOC heading leaks and visible cover PAGE fields; the helper derives header text from current section/header relationships and recognizes localized Heading1 style names or outline level 0.
- local project evidence: `D:\项目\旅游景点情感分析与研究\.codex\graduation-project-builder\20260523-continue-thesis-revision\reports\clean-final-v14-current-whole-format-gate-after-skill-fix.json`
- validation target: py_compile changed scripts; `audit_docx_whole_format_gate.py --self-test`; targeted selftests `whole_format_cover_media_template_text_only_valid` and `repair_docx_whole_format_toc_body_section_headers_valid`; rule-owner JSON parse; `scripts/check_utf8_clean.py --root <skill-root> --json`; `scripts/validate_skill_gate.py --skill-root <skill-root>`
