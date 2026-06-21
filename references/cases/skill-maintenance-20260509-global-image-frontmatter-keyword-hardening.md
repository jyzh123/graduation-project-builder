# Skill Maintenance Run 2026-05-09: Global Image, Front-Matter, And Keyword Hardening

## Run Manifest

- run_id: skill-maintenance-20260509-global-image-frontmatter-keyword-hardening
- task_mode: skill-maintenance
- subtask: global repair of graduation-project-builder gates after user reported image-ban violations, front-matter drift, and keyword label/content formatting confusion
- scope: only `C:\Users\Administrator\.agents\skills\graduation-project-builder`; thesis DOCX files under the project workspace were not modified
- authorization_source: user stated this session defaults to multi-agent authorization and requested multi-agent coordination
- agent_mode: reused existing live agents for read-only review plus controller implementation; earlier fresh spawn attempts failed in platform handoff and were recorded as sequential fallback in the prior run
- max_concurrent_live_agents: 6 existing live agents plus controller orchestration, no new live-agent overflow
- dispatch_wave_plan: wave1 existing read-only agents for scripts/rules/records/gate review; wave2 controller rule and record mutation; wave3 validation and sequential final audit
- audit_presence_by_wave: read-only audit prompts in wave1; controller sequential audit in waves 2 and 3
- concurrency_limit_verdict: pass
- required_lanes: controller; content-worker; format-worker; figure-worker; citation-worker; program-worker; acceptance-worker; audit
- complete_role_roster: controller; content-worker; format-worker; figure-worker; citation-worker; program-worker; acceptance-worker; audit
- role_attendance_matrix: controller active; figure-worker active; format-worker active; acceptance-worker active; audit active; content-worker not-applicable; citation-worker not-applicable; program-worker not-applicable
- not_applicable_lanes_with_reasons: content-worker=no thesis prose edit; citation-worker=no citation/bibliography mutation; program-worker=no runtime project mutation
- spawned_agent_ids: existing live agents `019e0abd-c13b-7011-9f6f-322a6e7a834c`, `019e0abd-c53a-7a51-a4f0-113312d7ad71`, `019e0abd-cabe-77e0-8f2e-30680e104a43`, `019e0abd-cfce-75d2-8d56-df2a2ec1dd59`, `019e0ac9-48b5-74e1-94de-cd2c864ce3c9`, `019e0ac9-846c-72b1-a9ec-2242b04d18af`
- spawned_agent_aliases_zh: 019e0abd-c13b-7011-9f6f-322a6e7a834c=审核-rules; 019e0abd-c53a-7a51-a4f0-113312d7ad71=审核-scripts; 019e0abd-cabe-77e0-8f2e-30680e104a43=审核-records; 019e0abd-cfce-75d2-8d56-df2a2ec1dd59=验收; 019e0ac9-48b5-74e1-94de-cd2c864ce3c9=审核-routing; 019e0ac9-846c-72b1-a9ec-2242b04d18af=审核-validation
- audit_agent_id: controller-final-audit
- sequential_audit_fallback_id: controller-final-audit
- action_audit_scope: inspect scripts, rule owners, transaction rules, maintenance rules, owner map, file-role index, durable audit, and targeted selftests for image/front-matter/keyword gates
- mutation_audit_scope: rule docs, owner map, file-role index, durable audit, and this case record
- handoff_status: pass after targeted validation; full aggregate fast-thesis-records command timed out and was split into passing batches

## Active Checklist

- figure manifest fail-closed source/final DOCX path and SHA binding: implemented in script and documented in owner rules
- whole-package media relationship scan across body/header/footer/comment/footnote/endnote story parts: implemented in script and documented in owner rules
- unauthorized image replacement, final-only media addition, and source media removal rejected: implemented with strict figure contract and mapped selftests
- image mutation inferred from Chinese image terms or source-to-final media diff: implemented in transaction validator and documented in transaction rules
- image mutation inferred from source-to-final drawing-object diff, including size, inline/anchor, relationship-set, and caption-adjacency changes with unchanged media bytes: implemented in transaction validator and figure contract
- target anchor must prove outside protected front matter for every image mutation unless official-template protected-image authorization passes: implemented and selftested
- direct picture replacement helper requires figure manifest/source DOCX and immediate-caption targeting: implemented in script and documented in maintenance rules
- keyword label/content split requires exact compact label-run equality, not prefix matching: implemented in validator evidence path and documented in abstract rules
- `FMT-ABSTRACT-001` and `QA-FINAL-047` are marked enforceable in `references/rule-owner-map.json`: implemented
- thesis DOCX files in the project workspace remain unmodified: required scope constraint

## Changed Files

- `references/thesis/format-rules/tables-abstracts-citations-references.md`
- `references/thesis/thesis-mutation-transaction.md`
- `references/user-feedback/maintenance-and-structure.md`
- `references/rule-owner-map.json`
- `FILE-ROLE-INDEX.md`
- `DURABLE-RULE-PROMOTION-AUDIT.md`
- `references/cases/skill-maintenance-20260509-global-image-frontmatter-keyword-hardening.md`

Script changes from the immediately preceding implementation pass are part of this closure:

- `scripts/thesis_figure_contract.py`
- `scripts/validate_thesis_mutation_transaction.py`
- `scripts/sample_self_check.py`
- `scripts/build_canonical_thesis.py`
- `scripts/docx_sync_picture.py`
- `scripts/validate_skill_gate_record_evidence.py`
- `scripts/selftest_skill_flow.py`

## Validation Plan

- `py -3 -m py_compile scripts/thesis_figure_contract.py scripts/validate_thesis_mutation_transaction.py scripts/sample_self_check.py scripts/build_canonical_thesis.py scripts/docx_sync_picture.py scripts/validate_skill_gate_record_evidence.py scripts/selftest_skill_flow.py`
- `py -3 -m json.tool references/rule-owner-map.json`
- `py -3 scripts/validate_skill_gate.py --skill-root .`
- `py -3 scripts/check_utf8_clean.py --root . --json`
- targeted selftests for strict figure manifest binding, transaction image mutation routing, and keyword label/content split
- targeted drawing-object selftests: `figure_manifest_drawing_extent_changed_rejected`, `figure_manifest_caption_adjacency_changed_rejected`, `transaction_drawing_extent_without_image_words_requires_manifest`, and `transaction_caption_adjacency_without_image_words_requires_manifest`
- `py -3 scripts/selftest_skill_flow.py --suite fast-thesis-records --quiet`, or split per-case execution if the single command exceeds the local timeout

## Validation Results

- `py -3 -m py_compile scripts/thesis_figure_contract.py scripts/validate_thesis_mutation_transaction.py scripts/generate_thesis_acceptance_record.py scripts/sample_self_check.py scripts/build_canonical_thesis.py scripts/docx_sync_picture.py scripts/validate_skill_gate_record_evidence.py scripts/validate_skill_gate_record_gate.py scripts/run_integration_gate.py scripts/selftest_skill_flow.py`: PASS
- `py -3 -m json.tool references/rule-owner-map.json > $null`: PASS
- `py -3 scripts/check_utf8_clean.py --root . --json`: PASS, 162 files checked, 0 issues
- `py -3 scripts/validate_skill_gate.py --skill-root .`: PASS
- `py -3 scripts/selftest_skill_flow.py --quiet` targeted 17-case core image/transaction/keyword/helper batch: PASS
- `py -3 scripts/run_integration_gate.py --case serialized_helper_mutation_chain --quiet`: PASS
- `py -3 scripts/selftest_skill_flow.py --suite fast-thesis-records --quiet`: TIMED OUT after about 424 seconds; not counted as pass
- Split `fast-thesis-records` batches covering the listed 51 existing cases plus the 4 new drawing-object cases: PASS in batches of 10/10/10/11/13

## Remaining Risks

- The skill can now reject more old manifests than before. This is intentional fail-closed behavior.
- Full rendered DOCX review still depends on available Office/WPS/rendering tools in the real project run; this maintenance run only validates the skill gates and selftest fixtures.
