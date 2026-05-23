# Skill Maintenance Case: Content Expansion Visual Gate

- case id: skill-maintenance-20260516-content-expansion-visual-gate
- task mode: skill-maintenance
- user trigger: explicit graduation-project-builder skill repair after a thesis body expansion allowed visible heading/title contamination in body text
- lock status: pass
- lock created before mutation?: yes
- loaded entrypoint: SKILL.md
- loaded routed references: references/user-feedback-persistence.md; references/user-feedback/final-qa-and-tooling.md; references/user-feedback/template-and-layout.md; references/agents/agent-lanes.md
- active checklist path: this file
- agent mode: single-agent-no-auth
- sequential audit fallback id: controller-audit-role
- mutation allowed verdict: pass canonical skill maintenance

## Problem

The previous thesis body expansion workflow could pass from XML/package counts, PDF export, and broad rendered-page existence. It did not force post-mutation machine-vision review of the exact output pages touched by inserted body paragraphs. The audit lane could therefore accept record-shape evidence while a rendered body paragraph visually inherited heading/title formatting.

## Required Repair

- Add a final QA rule for thesis DOCX text mutation/content expansion requiring exact-output rendered-page machine-vision review.
- Add a layout rule that treats heading/title-looking inserted body prose as a hard failure.
- Make the audit lane reject multi-agent or fallback records that lack post-mutation rendered evidence for content edits.
- Add final acceptance and agent manifest fields for content-mutation rendered review, body-heading contamination, and format-lane post-mutation audit.
- Add validator and selftest coverage so XML-only/PDF-export-only evidence fails for content expansion.

## Action Audit

- action: classify request and load skill-maintenance references
- owner: controller
- audit owner: controller-audit-role
- verdict: pass

- action: patch canonical rule files, templates, validator, selftests, owner map, and file-role index
- owner: controller
- audit owner: controller-audit-role
- verdict: pass

## Handoff Gate

- final validation required: py_compile changed scripts; JSON parse rule-owner map; targeted selftests for content expansion visual gate; UTF-8 clean; validate_skill_gate.py --skill-root .
- validation result: pass
- handoff allowed verdict: pass
