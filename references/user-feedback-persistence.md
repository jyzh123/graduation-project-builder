# User Feedback Persistence

Use this file as the routing layer for durable corrections learned whenever `graduation-project-builder` is used to create or modify files.

## Promotion Rule

- Do not keep effective user corrections only inside the current workspace.
- Same-turn durable promotion is allowed only when the current run is explicitly a canonical skill-audit run on the canonical skill bundle.
- Do not promote a project-local correction run directly into global skill rules. Record it locally first unless the run has been escalated into a canonical skill audit.
- Before durable promotion, update `DURABLE-RULE-PROMOTION-AUDIT.md` and confirm the required validator / selftest / integration results for the affected surfaces.
- Write durable corrections back into the skill references or scripts in the same turn only after that promotion gate has passed.
- Project-local `.learnings/LEARNINGS.md` is still required, but it is not the final destination for reusable graduation-project rules.
- This applies across sessions and across projects whenever this skill is actively used for creation or modification work.

## Incident Freeze Override

- If a root-level file matching `INCIDENT-*-THESIS-LANE-FREEZE.md` exists and its status is `ACTIVE`, suspend the default same-turn durable-promotion rule for thesis-related incidents.
- During that freeze, write new findings only into the active incident report, conflict matrix, or regression-fixture documentation unless the current task is the skill-audit task itself.
- Do not treat project-local thesis repair lessons as promotion-ready durable rules while the freeze is active.
- Once the freeze is cleared, resume normal durable-promotion behavior only after the integration gate and regression review pass.

## Enforcement Rule

- Every active rule in `references/rule-owner-map.json` is mandatory by default when the owning child file is loaded.
- Legacy numeric rule labels are aliases only. Use the manifest IDs such as `FB-LAYOUT-001` or `QA-FINAL-001` as the stable rule identity.
- Treat each numbered rule as a must-follow execution rule, not as optional guidance.
- A routed rule may be overridden only by a higher-precedence source:
  - explicit current-user instruction
  - official school template or official school rule
  - a more specific focused reference file that this router explicitly delegates to
- If no higher-precedence source overrides a routed rule, the rule must be enforced during execution and review.

## Child Files

- `references/user-feedback/program-delivery.md`: namespace `FB-PROG-*`
- `references/user-feedback/thesis-workflow.md`: namespace `FB-THESIS-*`
- `references/user-feedback/content-and-copy.md`: namespace `FB-COPY-*`
- `references/user-feedback/citations-and-bibliography.md`: namespace `FB-CITE-*`
- `references/user-feedback/template-and-layout.md`: namespace `FB-LAYOUT-*`
- `references/user-feedback/final-qa-and-tooling.md`: namespace `QA-FINAL-*`
- `references/user-feedback/maintenance-and-structure.md`: namespace `EXEC-MAINT-*`
- `references/agents/agent-lanes.md`: rules `AGENT-CTRL-001, AGENT-ALIAS-001, AGENT-ROSTER-001, AGENT-CONCURRENCY-001, AGENT-WORK-001, AGENT-FORMAT-001, AGENT-FORMAT-002, AGENT-FORMAT-003, AGENT-FORMAT-004, AGENT-FORMAT-005, AGENT-FORMAT-006, AGENT-AUDIT-001, AGENT-AUDIT-002, AGENT-AUDIT-003, AGENT-AUDIT-004, AGENT-FALLBACK-001, AGENT-CARD-001, AGENT-EVIDENCE-001`
- `references/rule-owner-map.json`: canonical owner manifest for the rule IDs above and their legacy numeric aliases
- `assets/agents/agent-task-card-template.md`: task-card template for agent lane work
- `assets/agents/agent-run-manifest-template.md`: run-manifest template for authorization, lane dispatch, fallback, and audit evidence

## Owner Map Coverage Fields

- `references/rule-owner-map.json` is also the coverage manifest for enforceable durable rules.
- When a rule is marked `enforcement_required: true`, it must name `validator_owner` or `selftest_owner`; optional `template_owner`, `generator_owner`, `required_load_modes`, and `acceptance_fields` record the rest of the enforcement chain.
- A skill-maintenance run may not claim consolidation complete when a newly enforceable rule has no validator/selftest owner or points to a missing path/anchor.
- Semantic-only or not-yet-mechanically-enforceable rules should remain unmarked as `enforcement_required` until their coverage owner is added.

## Loading Rule

- Load only the child files relevant to the current subtask instead of bulk-loading every durable correction file.
- When a thesis-scope task mixes bibliography repair, visible reference label
  family complaints, or page-flow/page-number complaints, load both
  `references/user-feedback/citations-and-bibliography.md` and
  `references/user-feedback/maintenance-and-structure.md` together.
- When the current user explicitly invokes `graduation-project-builder`, start a fail-closed skill-invocation lock from `assets/skill-invocation-lock-template.md` before mutation or handoff, and route EXEC-MAINT-065 through `references/user-feedback/maintenance-and-structure.md`.
- If the explicit invocation is recognized but execution would continue without the lock, active checklist, routed references, audit record, and canonical gate path, route EXEC-MAINT-071 through `references/user-feedback/maintenance-and-structure.md`; the lane is audit-only until the bypass risk is cleared.
- If any project inspection, DOCX inspection, code search, browser check, helper script, smoke check, or handoff text happens before the lock/checklist/audit bootstrap, route EXEC-MAINT-072 through `references/user-feedback/maintenance-and-structure.md`; mark the current run as contaminated/reference-only drift and restart from a fresh lock before mutation or handoff.
- For every substantial run invoked through this skill, load `references/agents/agent-lanes.md`, externalize `assets/agents/agent-run-manifest-template.md`, and fill `assets/agents/agent-task-card-template.md` for the controller and audit lanes before mutation or handoff; task cards are required for every canonical role lane (`总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, `审核`); skipped or not-applicable lanes must be recorded with reasons.
- For every execution behavior invoked through this skill, keep `references/agents/agent-lanes.md` active, record an action-level audit owner, and persist action-audit scope, cadence, and verdicts in the run manifest, task card, or lightweight audit entry before any completion claim.
- The agent run manifest must record the explicit authorization source, agent mode, required lanes, spawned agent ids or fallback reason, and audit verdict before any completion claim.
- The agent run manifest and task cards must record Chinese role aliases beside system agent ids; the audit alias must be `审核`.
- When a new durable correction is too detailed for this router, add it to the most specific child file or split out a new child file when necessary.
- If a new child file is introduced, update `SKILL.md`, `FILE-ROLE-INDEX.md`, and validation logic in the same turn.

## Parent Boundary

- Keep this file short and routing-oriented.
- Do not duplicate full rule bodies here once they have been moved into child files.
- Use this file to explain ownership, loading order, and where future durable corrections belong.
