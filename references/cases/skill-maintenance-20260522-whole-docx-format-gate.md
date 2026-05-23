# Skill Maintenance 20260522 Whole DOCX Format Gate

- run id: skill-maintenance-20260522-whole-docx-format-gate
- trigger: user reported SGB620/80T thesis cover, TOC, header/footer, style, color, and global format false pass
- mode: skill-maintenance with project audit evidence
- changed rule: EXEC-MAINT-076
- changed scripts: scripts/audit_docx_whole_format_gate.py; scripts/validate_skill_gate_record_core.py; scripts/validate_skill_gate_record_gate.py; scripts/validate_skill_gate_registry_core.py; scripts/validate_skill_gate_registry_records.py; scripts/generate_thesis_acceptance_record.py; scripts/selftest_skill_flow.py
- changed templates/indices: assets/final-acceptance-template.md; references/rule-owner-map.json; FILE-ROLE-INDEX.md; DURABLE-RULE-PROMOTION-AUDIT.md; SKILL.md
- exact current project audit: D:\项目\刮板输送机\.codex\graduation-project-builder\20260522-sgb620-80t-skill-gated-rebuild\evidence\whole-format-current.json
- current project gate result: fail; section_count=1; live_toc_field_count=0; footer_page_field_count=0; builder_style_visible_paragraph_count=468
- validation: py_compile changed scripts pass; audit_docx_whole_format_gate.py --self-test pass; targeted whole-format/font-color selftests pass; rule-owner JSON parse pass; check_utf8_clean pass; validate_skill_gate.py --skill-root pass
- thesis mutation status: blocked until the SGB620/80T manuscript is rebuilt or repaired to satisfy the new exact-output whole-DOCX structural gate
