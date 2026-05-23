---
name: graduation-project-builder
description: Finish a graduation project or course-design deliverable end to end across common software and thesis-delivery workflows; use whenever the user explicitly names graduation-project-builder or links this skill.
---

# Graduation Project Builder

Use this skill as a delivery finisher. The result should be runnable, explainable,
tested, demonstrable, packaged when needed, and supported by a matching thesis
when thesis work is in scope.

`SKILL.md` is the router and hard-gate entry point. Keep detailed rules in
focused files under `references/`, reusable templates in `assets/`, and
executable checks in `scripts/`.

## 0A. Explicit Invocation Bootstrap

- On explicit invocation, create the lock/checklist/audit bootstrap before project inspection, DOCX inspection, code search, browser checks, helper scripts, smoke checks, mutation, or handoff; only then inspect project files or execute work.
- If non-control action happened before the lock, mark contaminated/reference-only drift and restart from a fresh lock.

## 0. Cross-Cutting Gates

- Every loaded workflow, gate, and referenced rule is mandatory unless a higher-precedence source explicitly overrides it.
- Load the smallest focused reference set that covers the task, then work from an external checklist or task record rather than memory or ad hoc judgement.
- When the user explicitly invokes this skill, create a skill-invocation lock from `assets/skill-invocation-lock-template.md` before mutation or handoff. A missing, incomplete, or blocked lock means audit-only until resolved.
- Treat `$graduation-project-builder`, `graduation-project-builder`, or a direct path/link to this `SKILL.md` as an execution-state change, not reference context; before the lock/checklist/routed references/audit record exist, only routed skill reading and control-artifact creation are allowed.
- If a recognized invocation proceeds through ad hoc scripts, narrow smoke checks, sampled screenshots, local helper output, or hand-written final summary before the lock and canonical gate are active, it is a bypass attempt; stop, mark audit-only, and restart from the lock.
- Track each required step as `passed`, `failed`, or `explicitly skipped` with a
  reason. Treat an unmarked required step as incomplete.
- Lock condition-driven requirements before execution, including auth,
  persistence, admin surfaces, analytics/reporting, screenshots, figure
  families, citation/bibliography handling, and rendered-review availability.
  An implicit skip is a workflow failure.
- Read the real project before proposing or changing work.
- Preserve the existing stack, document format, template, and directory layout
  unless they are structurally broken or the user asks to change them.
- Default UI text, thesis text, evidence notes, and handoff summaries to Chinese
  unless the user says otherwise.
- Verify the exact output path being handed off.

Always load `references/user-feedback-persistence.md` for this skill. It is the
durable feedback router; after reading it, load only the child files that match
the current task.

If a root-level `INCIDENT-*-THESIS-LANE-FREEZE.md` is `ACTIVE`, thesis lanes are
audit-only unless the task is canonical skill maintenance.

## 0C. Skill Maintenance Consolidation Gate

- A skill change is incomplete until rules and workflow have been consolidated into the active source-of-truth chain.
- Every new or changed durable rule must have one canonical owner file, an exposed router path, a `references/rule-owner-map.json` entry, and either validator or selftest coverage when the rule is enforceable.
- Do not leave the same rule as disconnected fragments across `SKILL.md`,
  `references/`, `assets/`, scripts, and state notes.
- After changing this skill, run a consolidation pass that checks: owner file, router exposure, file-role index, templates, validators/generators, selftests, acceptance handoff text, and project/session state records.
- Skill audit, read-only review, diagnosis, repair, and maintenance all use
  `skill-maintenance` routing.

## 1. Mode Split

Choose exactly one primary mode before acting:

- `program-only`: software delivery, debugging, startup, packaging, or demo work.
- `thesis-only`: thesis writing, revision, evidence, references, or DOCX output.
- `program-plus-thesis`: both the running project and thesis must stay aligned.
- `format-repair-only`: bounded DOCX format repair without content rewriting.
- `skill-maintenance`: this skill bundle itself is being changed, audited,
  reviewed, diagnosed, or maintained.

Do not let thesis-formatting rules dominate a program-only task. Do not let
program delivery work bypass thesis gates when a thesis deliverable is part of
the request.

## 2. Core Working Rules

- Reconstruct missing requirements from the repository and visible artifacts.
- Prefer autonomous execution after direction is clear.
- Prefer believable, runnable delivery over wide but shallow feature lists.
- If the current task touches files, verify the exact output path that will be
  handed off.
- If a reference contains `Child Files`, `Loading Rule`, or `Parent Boundary`,
  continue into the relevant child files before executing.
- Do not downgrade an explicit skill invocation into "read the skill then act
  normally." Completion requires evidence that the skill controlled execution:
  skill lock path, active checklist path, agent/audit record, and exact
  `validate_skill_gate.py --gate-record <record>` result when a substantial
  handoff is made.

## 3. Active Rule Externalization

For every substantial run:

1. identify the current mode and subtask
2. load only the relevant routed references
3. externalize the active checklist
4. externalize the agent run manifest and task cards when audit/lane routing is
   in scope
5. execute against the checklist
6. validate against the checklist and acceptance gate before handoff

For any substantial invocation, load `references/agents/agent-lanes.md` and use:

- `assets/agents/agent-run-manifest-template.md`
- `assets/agents/agent-task-card-template.md`

For small read-only behavior, record a lightweight action-audit note in the
handoff. For substantial runs, create a manifest/task-card record or name the
explicit fallback reason.

Agent and audit hard gates:

- For every substantial run, if the current user has explicitly authorized subagents, delegation, or parallel agent work for this turn and parallel subagents are available, the controller must split work into the relevant worker lanes plus an independent audit agent and record the dispatch in an agent run manifest before handoff.
- If explicit authorization is absent, do not claim multi-agent execution; either ask for authorization before any substantial multi-agent decomposition or record the run as `single-agent-no-auth` without pretending that lanes were spawned.
- Agent visible names are platform-assigned; record Chinese role aliases such as `总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, and `审核` beside the real system agent ids instead of claiming the UI display name was customized.
- The audit lane must always use the Chinese role alias `审核` in task cards, run manifests, prompts, and acceptance records.
- Create the complete canonical role roster first (`总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, `审核`), then route each touched surface family to its corresponding worker lane. A lane that has no work for the current request must still have a task card marked `not-applicable` or `skipped-with-reason`; it must not disappear from the run.
- Do not interpret the complete eight-role roster as permission to keep eight spawned agents alive at once. Codex supports at most six simultaneous live agents; if more than six roles are active, split the run into dispatch waves or sequential fallback batches, and keep the audit lane present in every wave.

Subagent authorization only decides whether agents are spawned. It does not
decide whether audit rules are loaded. If spawning is unavailable, record
sequential-audit-fallback evidence through `references/agents/agent-lanes.md`.

## 4. Program Workflow

For `program-only`, and for the program part of `program-plus-thesis`, load:

- `references/program/workflow-standard.md`
- `references/program/verification-matrix.md`
- `references/program/stack-adaptation.md` when the stack is unclear
- `references/program/executable-automation.md` when startup or packaging is in
  scope
- `references/program/packaging-rules.md` when a delivery bundle is required
- `references/tooling-dependencies.md`

Then fill `assets/program-gap-checklist.md`, build missing work in vertical
slices, verify the main user/admin/analytics flows that apply, and review with:

- `references/review-program-checklist.md`
- `references/review-delivery-bundle-checklist.md`

## 5. Thesis Workflow

For `thesis-only`, `format-repair-only`, and the thesis phase of
`program-plus-thesis`, load the focused thesis references for the touched
surface. Do not stop at this file.

Baseline references:

- `references/thesis/thesis-production-workflow.md` when producing or drafting a
  new thesis, shortest draft, sample manuscript, or full manuscript
- `references/thesis/thesis-workflow-map.md`
- `references/thesis/thesis-mutation-transaction.md`
- `references/thesis/thesis-execution-contract.md`
- `references/thesis/thesis-format-rules.md`
- `references/thesis/thesis-format-sop.md`
- `references/thesis/thesis-format-class-review.md`
- `references/thesis/format-rules/protected-surface-evidence-contract.md`
- `references/user-feedback/maintenance-and-structure.md` for DOCX hard gates
  involving body heading levels, style blast radius, format preservation,
  review artifacts, citation superscripts, blocked evidence, and final
  acceptance false-pass prevention
- `maintenance-and-structure.md` owns helper-script, smoke audit, and active
  checklist routing when thesis helper behavior is involved
- `references/thesis/thesis-figure-generation-rules.md` when figures,
  screenshots, diagrams, or captions are in scope
- `references/thesis/thesis-template-learning.md` when template or sample
  behavior is in scope
- `references/thesis/thesis-troubleshooting-log.md` when recovering from drift,
  blank pages, TOC issues, media loss, or repeated repair failure
- `references/thesis-table-style-memory.md` when tables are in scope
- `references/thesis-formula-style-memory.md` when formulas are in scope; formula acceptance must include `scripts/audit_docx_formula_objects.py` evidence when formulas or formula-like text are present
- `references/policy/cnki-citation-policy.md` when citation source policy is
  needed

Mandatory thesis gate lines:

- For thesis generation, including shortest draft generation, smoke samples, or manual-review test manuscripts, do not bypass `sample_self_check`, acceptance-record generation, and `validate_skill_gate` on the exact output path being handed off.
- A smoke or final-verification summary that only checks `officecli validate`, PDF export, media counts, page-image existence, old-term counts, phrase presence, or broad screenshot existence is an intermediate artifact. It cannot be cited as final acceptance even if the filename contains `final-acceptance`.
- Shortest drafts and test manuscripts are not exempt from abstract-surface parity, citation rules, three-line-table authority, structural-figure provenance, formula-example requirements, or header/footer baseline checks when those surfaces are present or explicitly requested.
- If the user explicitly invokes `graduation-project-builder` for thesis work, reading `SKILL.md` alone is not enough. The run must continue into the routed thesis child files needed for that subtask, externalize one active checklist, and execute against that checklist rather than falling back to ad hoc local judgement.
- Before any thesis mutation or project-local helper creation, run `scripts/scan_project_local_thesis_helpers.py --project-root <project-root> --fail-on-risk` on the real project root. If it reports risky local thesis helper scripts, the lane is audit-only until a clean-source restart or canonical-helper replacement is recorded.
- If the user asks for whole-thesis / full-paper / `1:1` template alignment, do not silently narrow the verification scope to selected chapters, visible hotspots, or body-only surfaces. Cover, Chinese abstract, English abstract, TOC, body, figures/tables, references, and acknowledgement remain in scope unless the user explicitly excludes them.
- Runtime screenshot slots must show real runtime screenshots. Structural
  diagrams, stale assets, blank captures, or mismatched media in those slots
  block handoff.
- Visible citation anchor leakage is a hard failure. Rendered body text must not
  expose `cite_ref`, bookmark names, or hyperlink helper text.
- Live TOC is a hard surface when required. A final manuscript with a
  handwritten/static TOC, mixed static-plus-field TOC, or leaked TOC placeholder
  is not complete.
- Whole-thesis generation or revision must pass `scripts/audit_docx_whole_format_gate.py` on the exact final DOCX and bind its path, verdict, and SHA256 in final acceptance; narrower font/color/body-style/PDF-export checks cannot clear section, TOC, header/footer, page-number, surface-order, or builder-style failures.
- Minimum page-class sample comparison set: cover, Chinese abstract, English
  abstract, TOC, first body chapter page, one figure page, one table page,
  references, and acknowledgement.

Local helper hard gates:

- Do not auto-generate project-local thick thesis rewrite scripts that replicate heading, abstract, figure, caption, table, bibliography, or pagination logic outside the canonical skill bundle. If a local script is unavoidable, it must be a thin wrapper that delegates to canonical skill scripts/helpers with locked target paths and locked surface scope.
- Do not create or use workspace-local thick thesis generation scripts that replicate heading, abstract, TOC, body, figure, table, bibliography, pagination, font, or template-donor logic outside the canonical skill bundle. A Codex workspace helper for thesis generation must be either disposable orchestration only or a thin wrapper around canonical skill scripts with locked inputs, locked outputs, and no independent business-format policy.
- For thesis generation, canonical skill scripts are the only allowed owner of DOCX business-format behavior. If the canonical script cannot express the required project content, extend the canonical skill script and its tests first instead of creating an ad hoc workspace builder.
- General thesis-making scripts belong inside this canonical skill bundle. Project-local files may only provide template/project-specific adapter data, template profiles, content manifests, locked-path run manifests, or thin wrappers that delegate to canonical skill scripts.
- For new thesis generation from validated local adapter/content manifests, use `scripts/build_canonical_thesis.py` as the canonical general builder before adding any new project-local orchestration.
- A project-local thesis adapter is valid only after `scripts/validate_thesis_local_adapter.py` passes on that adapter. If local template-specific requirements cannot fit the adapter schema, extend the canonical adapter validator and canonical builder path first instead of putting a generic thesis builder in the project directory.

Project-local helper preflight is required for thesis generation, thesis
revision, and format repair whenever a real project root is involved. Record the
scanner report path, risk count, disposition, and clean-source restart status.

Thesis execution order:

1. classify the workflow through `references/thesis/thesis-workflow-map.md`
2. create a mutation transaction from
   `references/thesis/thesis-mutation-transaction.md` before DOCX mutation
3. discover and lock the active template/profile before format mutation
4. create surface inventory and high-risk surface matrix when thesis format is
   in scope
5. modify only through the selected workflow path
6. render or otherwise inspect touched pages and blast-radius pages
7. generate or update the acceptance record
8. run the gate on the exact final output

The thesis workflow map owns `new-thesis-production`,
`whole-thesis-revision`, `local-surface-repair`,
`content-only-paragraph-revision`, and `audit-only`; do not invent an ad hoc
thesis lane.

For header/footer, section, chapter-start, and tail-block title repairs, follow
the review-copy-first promotion workflow in
`references/thesis/thesis-format-sop.md`; do not promote back to the master
manuscript until rendered review passes.

For thesis handoff review, use:

- `references/review-thesis-format-checklist.md`
- `references/review-thesis-content-consistency-checklist.md`
- `references/review-figure-style-checklist.md` when figures are in scope

## 6. Delivery Templates

Use templates instead of improvising structure:

- `assets/program-gap-checklist.md`
- `assets/thesis-blueprint-template.md`
- `assets/figure-task-template.md`
- `assets/figure-plan-template.md`
- `assets/format-repair-task-template.md`
- `assets/final-acceptance-template.md`
- `assets/paper-only-bibliography-review-template.md`
- `assets/humanizer-evidence-template.md`
- `assets/review-evidence-template.md`
- `assets/user-reported-issue-ledger-template.md`
- `assets/agents/agent-run-manifest-template.md`
- `assets/agents/agent-task-card-template.md`

## 7. Acceptance Gate

Before handoff on a substantial run:

1. write an acceptance record from `assets/final-acceptance-template.md`
2. run `scripts/validate_skill_gate.py --gate-record <this exact acceptance record>`
   or the `.cmd` wrapper with the same exact record/output binding
3. run `scripts/run_integration_gate.py` for thesis DOCX work, DOCX toolchain
   changes, or required integration coverage
4. state exact output paths, evidence paths, and any skipped verification reason

If the gate returns non-zero, the run is incomplete.

Acceptance-record requirements:

- the acceptance record must name the active references, active checklists, condition locks, blockers, skips, exact output paths, evidence paths, and validation result
- the acceptance record must also name the agent authorization source, agent mode, required lanes, Chinese role aliases, spawned agent ids or fallback reason, audit agent id or sequential audit fallback id, and agent run manifest path for any substantial run
- when a review or verification surface was used, the acceptance record must
  cite the corresponding artifact path or explicit path list rather than only a
  prose claim
- evidence paths should point to review-evidence-template.md records
- all verification surfaces must have a review-evidence path or an explicit
  skipped/blocked reason in the acceptance record
- after thesis citation or bibliography work, evidence paths must include a
  body-citation audit report path from `scripts/audit_thesis_citations.py`
- after thesis modification work that touched visible text or font-bearing DOCX
  surfaces, evidence paths must include a DOCX font/encoding audit report path
- after thesis modification work, evidence paths must include touched-page
  review evidence in addition to paragraph-review evidence
- do not allow skipped items to remain blank; use an explicit reason or
  explicit `none`

## 8. Memory Policy

Keep `SKILL.md` as short as the active hard-gate layer allows. Store detailed
rules in the right focused owner:

- reusable workflow rules: `references/user-feedback-persistence.md`
- thesis formatting rules: `references/thesis/*.md` and
  `references/thesis/format-rules/*.md`
- thesis workflow and transaction rules:
  `references/thesis/thesis-workflow-map.md` and
  `references/thesis/thesis-mutation-transaction.md`
- execution-layer path locks, review-copy naming, renderer handoff, and lock
  recovery: `references/thesis/thesis-execution-contract.md` and
  `references/tooling-dependencies.md`
- agent lane rules: `references/agents/agent-lanes.md`
- table-specific visual rules: `references/thesis-table-style-memory.md`
- checklists: `references/review-*.md`
- templates: `assets/*.md`

Do not turn `SKILL.md` into a topic-detail dump. If a focused rule file becomes
too heavy, split it into smaller topic-owned files instead of expanding this
entry file.

## 9. Resource Map

Core:

- `references/user-feedback-persistence.md`
- `references/tooling-dependencies.md`
- `references/rule-owner-map.json`
- `references/agents/agent-lanes.md`

Program:

- `references/program/workflow-standard.md`
- `references/program/stack-adaptation.md`
- `references/program/verification-matrix.md`
- `references/program/executable-automation.md`
- `references/program/packaging-rules.md`

Thesis:

- `references/thesis/thesis-production-workflow.md`
- `references/thesis/thesis-workflow-map.md`
- `references/thesis/thesis-mutation-transaction.md`
- `references/thesis/thesis-format-rules.md`
- `references/thesis/format-rules/protected-surface-evidence-contract.md`
- `references/thesis/thesis-format-sop.md`
- `references/thesis/thesis-execution-contract.md`
- `references/thesis/thesis-format-class-review.md`
- `references/thesis/thesis-figure-generation-rules.md`
- `references/thesis/thesis-template-learning.md`
- `references/thesis/thesis-troubleshooting-log.md`
- `references/thesis/thesis-companion.md`
- `references/thesis-table-style-memory.md`
- `references/thesis-formula-style-memory.md`

Review:

- `references/review-program-checklist.md`
- `references/review-delivery-bundle-checklist.md`
- `references/review-thesis-format-checklist.md`
- `references/review-thesis-content-consistency-checklist.md`
- `references/review-figure-style-checklist.md`

Policy and persistence:

- `references/policy/cnki-citation-policy.md`
- `references/user-feedback-persistence.md`
- `references/user-feedback/maintenance-and-structure.md`
- `references/rule-owner-map.json`
- `references/agents/agent-lanes.md`
- `references/tooling-dependencies.md`

Scripts:

- `scripts/build_canonical_thesis.py`
- `scripts/validate_thesis_local_adapter.py`
- `scripts/scan_project_local_thesis_helpers.py`
- `scripts/discover_project_thesis_template.py`
- `scripts/thesis_template_profile.py`
- `scripts/sample_self_check.py`
- `scripts/generate_thesis_acceptance_record.py`
- `scripts/audit_thesis_citations.py`
- `scripts/audit_docx_font_encoding.py`
- `scripts/audit_docx_body_style.py`
- `scripts/inspect_docx_pagination_structure.py`
- `scripts/docx_apply_table_family.py`
- `scripts/docx_formula_number_table.py`
- `scripts/docx_sync_picture.py`
- `scripts/normalize_thesis_citation_chain.py`
- `scripts/repair_thesis_surface_format.py`
- `scripts/run_integration_gate.py`
- `scripts/validate_skill_gate.py`
- `scripts/validate_skill_gate.cmd`
