# Agent Lane Rules

## Controller Lane
### AGENT-CTRL-001. Controller Lane Owns Decomposition And Final Merge (Mandatory)
- The controller lane owns decomposition, lane assignment, dependency tracking, and final merge.
- The controller lane must keep one active task card per lane and must record the current lane state, authorization source, agent mode, and audit owner before handoff.
- Every execution behavior must be represented as an action cycle with an audit owner before it starts. Action cycles include inspection, classification, rule loading, checklist externalization, planning, tool execution, mutation, verification, packaging, and handoff.
- The controller must keep the audit lane in attendance for action cycles even when no file is changed. Read-only or planning actions still require an action-level audit verdict.
- Every substantial invocation of `graduation-project-builder` must create a controller task card and an audit task card before any file mutation or handoff; task cards are mandatory for every canonical role lane; inactive lanes must still be present as `not-applicable` or `skipped-with-reason`.
- The substantial-invocation rule above is a minimum record-shape rule, not a limit on audit coverage; non-substantial skill behavior still requires action-cycle audit attendance and evidence.
- Every substantial invocation must keep the audit lane active for the whole run. The controller may not treat audit as a final-only checklist when files, thesis surfaces, generated assets, helper scripts, or delivery records are modified.
- Before each mutation, the controller must name the expected audit scope; after each mutation cycle, the controller must record the audit verdict before starting the next unrelated mutation.
- If the current user has explicitly authorized subagents, delegation, or parallel agent work for this turn, the controller must dispatch the touched surface families to the corresponding worker lanes and record the dispatch in an agent run manifest before handoff.
- Complete role roster coverage is not permission to spawn the whole roster at once. Codex supports at most six simultaneous live agents, so the controller must plan dispatch waves when active roles would exceed that cap.
- If explicit authorization is absent, the controller must not claim multi-agent execution; it must stay single-agent, record `agent_mode: single-agent-no-auth`, create no spawned-worker claim, record any worker lanes as skipped with reasons, and still complete the audit lane as `sequential-audit-fallback`.
- Do not claim a multi-agent run in the handoff record without an authorization source, a run manifest, and audit verdict.

## Role Aliases
### AGENT-ALIAS-001. Chinese Role Aliases Must Be Recorded With System Agent IDs (Mandatory)
- Platform-assigned agent display names cannot be controlled by this skill; the skill must not claim that the visible UI name was changed.
- Every controller, worker, and audit task card must record `role_alias_zh` and `system_agent_id` so the intended Chinese role name can be checked against the actual spawned or sequential executor.
- Every task card must also record `lane_alias_zh`, `owner_alias_zh`, `spawn_agent_alias_zh`, and `audit_agent_alias_zh` beside the canonical lane id and system agent id.
- Use the canonical Chinese role aliases `总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, and `审核`; the audit lane alias must be `审核`.
- Canonical lane ids remain authoritative; Chinese role aliases are display and audit labels, not replacements for `lane`, `owner`, `spawn_agent_id`, or `audit_agent_id`.
- The agent run manifest and final acceptance record must include the Chinese role alias mapping beside spawned agent ids or sequential fallback ids.
- Agent prompts should address the assigned role alias, for example `你是本轮的“审核”agent`, but completion evidence must still record the real system agent id.
- Agent prompts, task cards, manifests, and acceptance records must use the clean literal alias map `controller=总控`, `content-worker=内容`, `format-worker=格式`, `figure-worker=图表`, `citation-worker=引用`, `program-worker=程序`, `acceptance-worker=验收`, and `audit=审核`. Short role-family keys such as `content` or `format` are labels only and must not replace canonical lane ids in `lane_alias_map_zh`.
- Do not copy role aliases from terminal output, lossy console captures, or previously corrupted records. If any role alias is not exactly one of the eight clean aliases in the literal alias map, stop before dispatch or handoff, regenerate the prompt or record from the clean literal alias map, and record the repair in the audit lane.


## Complete Role Roster
### AGENT-ROSTER-001. Full Canonical Role Roster Must Be Present (Mandatory)
- Every substantial invocation of `graduation-project-builder` must create a complete role roster before execution: `总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, and `审核`.
- `required_lanes` means the full canonical roster for this skill, not only the lanes selected by the controller. A run may mark a lane `active`, `not-applicable`, or `skipped-with-reason`, but it may not omit the lane.
- Every canonical role must have a task card or lightweight action-audit entry with `role_applicability`, `attendance_status`, and either active evidence or a concrete not-applicable/skipped reason.
- Small read-only behavior that does not mutate files may use one durable lightweight action-audit note instead of a full eight-role task-card matrix; it must still name the audit owner, checked rule source, verdict, and blocker state.
- The controller must not claim multi-agent compliance when only selected worker lanes were created. Selective lane creation is a workflow failure.
- The audit lane must compare the manifest, task cards, and final acceptance record against the full canonical roster and reject handoff if any canonical role is missing, lacks status, or lacks a reason for not participating.

## Concurrency Cap
### AGENT-CONCURRENCY-001. Live Agent Concurrency Is Capped At Six (Mandatory)
- The complete canonical role roster is a record and responsibility requirement, not a requirement to keep eight spawned agents alive at the same time.
- A `parallel-subagents` run must record `max_concurrent_live_agents`, `live_agent_count_plan`, `dispatch_wave_plan`, `audit_presence_by_wave`, and `concurrency_limit_verdict` in the run manifest and final acceptance record.
- `max_concurrent_live_agents` must never exceed `6`.
- If more than six roles are active, the controller must split execution into waves or sequential fallback batches. Each wave must include the audit lane or a documented sequential audit fallback; no wave may run without audit attendance.
- The audit lane must reject handoff when the manifest claims full-roster execution but omits the concurrency plan, exceeds six simultaneous live agents, or fails to show audit presence for every dispatch wave.
## Worker Lanes
### AGENT-WORK-001. Worker Lanes Stay Bounded To One Task Card (Mandatory)
- Each worker lane must stay on one bounded task card and must not replan other lanes.
- Each worker lane must return status, changed files, blockers, and evidence pointers before the controller closes the lane.
- Create the complete canonical role roster first (`总控`, `内容`, `格式`, `图表`, `引用`, `程序`, `验收`, `审核`), then route each touched surface family to its corresponding worker lane. A lane that has no work for the current request must still have a task card marked `not-applicable` or `skipped-with-reason`; it must not disappear from the run.
- A worker lane may not claim a surface family owned by another lane or rewrite the audit verdict.

### AGENT-FORMAT-001. Format Lane Must Discover And Lock The Active Thesis Template Before Mutation (Mandatory)
- The `格式` lane must run `scripts/discover_project_thesis_template.py` against the project root before any thesis-format mutation unless the user provided an explicit template path in the current request.
- The `格式` lane task card must record the project template discovery root, discovery patterns, discovered candidate template paths, selected template reason, active template source type, active template path lock, active template fingerprint, active template profile path, and whether the template was selected before mutation.
- The active template source type must be one of `project-auto-discovered`, `user-provided`, or `teacher-approved-sample`.
- The active template profile must be generated from the active template path before mutation and must match the active template fingerprint.
- The `审核` lane must reject a formatting handoff when the `格式` task card lacks template discovery evidence, active template fingerprint, active template profile, or rendered template-alignment evidence.
- If the selected template cannot be profiled or fingerprinted, the `格式` lane is blocked; it must not infer formatting from the current broken manuscript, Word built-in styles, or a previous memory note.

### AGENT-FORMAT-002. Format Lane Revalidates After Content Mutation (Mandatory)
- When the `内容` lane writes or rewrites a thesis paragraph, the `格式` lane must classify the touched surface before the next paragraph is edited.
- The `格式` lane task card must record the touched paragraph identifiers, touched surface family, sibling surfaces, blast-radius pages, stale audits, and required rerender targets for each content mutation batch.
- If a content edit changes line count, figure/table adjacency, heading position, citation state, tail-block position, or page flow, the run is no longer pure content-only; the controller must switch to `local-surface-repair` or `whole-thesis-revision` before more mutation.
- The `格式` lane must not accept a reused rendered page, self-check, font audit, body-style audit, citation audit, or acceptance record unless it can prove that evidence was produced against the exact current review copy or final output after the content mutation.
- The `审核` lane must reject a handoff when the `格式` card lacks touched-surface, blast-radius, stale-audit, rerender, or exact-output evidence for changed content.

### AGENT-FORMAT-003. Format Lane Must Use The Full Surface Owner Matrix (Mandatory)
- The `格式` lane must use `references/thesis/thesis-format-class-review.md` as the surface owner matrix for thesis formatting and template-alignment work.
- The `格式` lane must also use `references/thesis/format-rules/protected-surface-evidence-contract.md` as the shared evidence contract for protected-surface ids, effective font-chain proof, and final acceptance handoff fields.
- A surface family may not be checked by a generic body/Normal rule when the class-review matrix names a more specific surface.
- The `格式` lane must record one write owner for each touched or user-reported surface and must list surfaces explicitly excluded from generic scripts before mutation.
- If a generic helper script would touch multiple protected surface families in one pass, the controller must split the pass or mark the run blocked.

### AGENT-FORMAT-004. Format Lane Must Bind Evidence To Canonical Protected Surface IDs (Mandatory)
- The `格式` lane task card must list the canonical protected surface ids in scope, the owner lane for each id, the evidence record path for each id, the exact reviewed output path, and the exact reviewed output SHA256; this is the protected-surface evidence contract binding for the format lane.
- Protected surface ids must match `references/thesis/format-rules/protected-surface-evidence-contract.md`; generic labels such as `abstract`, `toc`, `front matter`, `references`, or `font checked` are invalid in task cards and manifests.
- The `格式` lane must reject a pass verdict when one evidence path is reused for multiple protected surfaces, when an evidence record lacks effective font-chain fields for a text surface, when any present/touched/user-reported template-owned surface lacks numeric rendered template-vs-target geometry, when any present/touched/user-reported template-owned surface has only style-binding prose without WPS/Word paragraph-dialog and typography hard fields, when TOC evidence lacks the additional TOC visual-geometry comparison fields, when TOC evidence proves only per-level style binding without Word/WPS paragraph-dialog and typography metrics, or when a rendered artifact was produced from a different output path.
- The `审核` lane must compare the protected surface owner map in the `格式` card, the agent run manifest, and the final acceptance record before handoff.

### AGENT-FORMAT-005. Format Lane Must Reconcile Local Repair With Whole-Document Pagination (Mandatory)
- The `格式` lane must treat `whole_document_pagination` as a canonical protected surface for every thesis-format run, even when the user reports only one local page, TOC, heading, footer, reference, or spacing defect.
- After any mutation that can change flow, section boundaries, headers/footers, TOC page references, chapter openers, figures, tables, references, acknowledgement, appendix, or page numbers, the lane must regenerate the whole-document pagination evidence path before handoff.
- The whole-document pagination evidence must include package baseline/drift, pre/post page maps, section boundary/property maps, page-number restart map, header/footer link-to-previous map, hard page-break and section-break map, field-refresh before/after state, TOC-to-heading page sync map, logical-to-physical page map, rendered page count baseline/actual, blank-page scan, chapter opener map, tail-block opener map, and final verdict.
- Local rendered screenshots, sampled page checks, chapter-start summaries, tail-block summaries, or PDF export success cannot substitute for this evidence.

### AGENT-FORMAT-006. Format Lane Must Freeze Review Artifacts And Citation Runs (Mandatory)
- Before any thesis DOCX text mutation, the `格式` lane must record whether the source manuscript or comment-carrier DOCX contains comments, comment anchors, tracked changes, bookmarks, fields, hyperlinks, and body citation superscript runs.
- The `格式` lane task card must name the source review-artifact inventory path, source body-citation run inventory path, post-mutation review-artifact diff path, and post-mutation citation-run diff path.
- If comments or tracked changes exist in the source, the protected surface id `review_comments_and_change_marks` must appear in the surface owner map, task cards, inventory, high-risk matrix, and final acceptance record unless an explicit user-approved no-comments preview route is recorded.
- If body citations exist or references are present, `body_citation_superscripts` must appear as a protected surface and the citation audit must bind to the final DOCX SHA256.
- The `审核` lane must reject a format handoff when these inventories or diffs are missing, stale, or not bound to the exact source and final DOCX files.

## Audit Lane
### AGENT-AUDIT-001. Audit Lane Owns Evidence Review And Verdict (Mandatory)
- The audit lane owns verification, evidence review, and the final pass-or-fail verdict for every skill behavior, including non-substantial, read-only, substantial, and single-agent-no-auth records.
- The audit lane must cite the exact artifacts it reviewed and must name every blocker that remains open.
- When explicit authorization is absent, the controller must still complete the audit role sequentially in the manifest and must not pretend a separate audit agent was spawned.
- If explicit authorization exists or sequential fallback is needed, the audit lane is mandatory and must review the run manifest, worker task cards, and exact evidence artifacts before final verdict.
- The audit lane must cite the authorization source and dispatch mode it verified, not only the files it inspected.
- The audit lane must verify that `graduation-project-builder` was actually invoked for the run, that the routed skill references required by the task mode were loaded, and that the active checklist was externalized before execution.
- The audit lane must verify each action cycle, including read-only actions, planning decisions, tool calls, generated artifacts, mutations, checks, and handoff statements.
- The audit lane must verify user-request compliance, loaded-rule compliance, changed-file scope, and whether any mutation invalidated earlier evidence.
- The audit lane must reject handoff when a change cycle lacks a mutation-level audit verdict, even if the final aggregate evidence appears complete.

### AGENT-AUDIT-002. Audit Lane Must Parse Agent Records Not Just Path-Check Them (Mandatory)
- The audit lane must review the contents of the agent run manifest and every lane task card referenced by the final acceptance record.
- A path-only check is not enough: the audit must verify lane alias, system agent id, authorization source, agent mode, audit verdict, template lock fields, touched surfaces, evidence paths, and blockers in the referenced records.
- For format work, the audit must compare the `格式` task card against the final acceptance record and reject mismatches in active template path, fingerprint, profile path, selected-before-mutation status, template alignment verdict, touched surface families, and evidence paths.
- For thesis-format work, the audit must also compare the protected surface ids, owner map, reviewed output path/SHA256, per-surface evidence paths, effective font-chain verdicts, TOC visual-geometry verdicts, TOC paragraph-and-typography verdicts, citation/reference evidence paths, and rendered page evidence paths against `references/thesis/format-rules/protected-surface-evidence-contract.md`.
- For thesis DOCX mutation, the audit must independently compare source and final DOCX review-artifact inventories and body-citation run inventories. It must inspect or require machine evidence for `word/comments*.xml`, `word/people.xml`, comment anchors, tracked-change marks, bookmarks, fields, hyperlinks, citation bookmarks, citation marker run separation, superscript state, visual style, and final DOCX SHA256 binding.
- The audit must reject stale evidence when the evidence record does not identify the reviewed output path. During mutation-cycle review, that path may be the locked current review copy for that cycle; during final acceptance, it must be the exact final DOCX path required by FMT-EVID-007.

### AGENT-AUDIT-003. Audit Lane Must Reject Stale Or Generator-Synthesized Format Evidence (Mandatory)
- The `审核` lane must reject final acceptance when TOC, style-bound surface, page-class, or whole-document pagination evidence was synthesized from broad statuses such as `toc_ok`, `page_classes_ok`, `rendered pdf exists`, `chapter_ok`, `tail_block_ok`, or `sample_self_check passed`.
- The audit must verify that the final acceptance record, manifest, task cards, surface inventory, high-risk matrix, and review evidence records all point to the same exact final output path and SHA256 required by FMT-EVID-007, and that evidence creation is later than the last mutation.
- The audit must compare measurement provenance for TOC visual geometry, TOC paragraph/typography, all-surface paragraph/typography, and whole-document pagination. Missing producer path, missing JSON/report path, reused evidence, or hand-entered natural-language measurements block handoff.
- The audit must reject handoff when any evidence or helper report states that comments, tracked changes, or citation runs were stripped, cleaned, accepted, or regenerated without explicit user approval and a source-to-final controlled-change ledger.
- The audit must reject a citation audit that has `result: pass` but lacks exact final DOCX SHA256 binding, source-to-final citation-marker run evidence after text mutation, or regeneration after the last DOCX write.

### AGENT-AUDIT-004. Audit Lane Must Enforce Rendered Review After Content Expansion (Mandatory)
- When the `content-worker` lane inserts, expands, rewrites, or lengthens thesis body paragraphs, the `audit` lane must reject final handoff unless the `format-worker` lane reviewed rendered page images from the exact post-mutation output.
- The audit must parse, not path-check, the content-mutation rendered review record and confirm it names touched paragraph/page ids, rendered page images, body donor comparison, heading/title contamination checks, and pass-shaped machine-vision verdicts.
- The audit must reject XML/package-count-only, PDF-export-only, page-count-only, style-name-only, manual-only, sampled-only, stale, or `not checked` evidence for body paragraph insertion or expansion.
- A multi-agent claim is not valid when content and format workers return before post-mutation rendered evidence exists, even if the final controller summary says format was preserved.

#### Action Cycle Audit Requirement
- An action cycle is any observable behavior taken under this skill, including reading files, searching, classifying a task, selecting references, creating a checklist, planning, running tools, editing files, generating outputs, validating, or reporting completion.
- Each action cycle must have an audit entry that names the action category, action owner, skill rules considered, expected outcome, checks performed, verdict, and blockers.
- A mutation cycle is a high-risk action cycle and must also satisfy the mutation-specific audit requirements below.
- Missing action-level audit evidence is a workflow failure even when no files were changed.

#### Mutation Cycle Audit Requirement
- A mutation cycle is any change to repository code, generated assets, thesis/DOCX/PDF files, helper scripts, checklists, manifests, acceptance records, or skill/reference files.
- Each mutation cycle must have an audit entry that names the changed paths, expected behavior, loaded skill rules, checks performed, verdict, and blockers.
- The audit entry may be produced by a spawned independent audit agent when allowed, or by the sequential audit fallback when spawning is unavailable or unauthorized; in both cases the manifest must make the execution mode explicit.
- A later successful audit does not erase a missing earlier mutation audit. Missing cycle-level audit evidence is a workflow failure that must be recorded and repaired before handoff.

## Sequential Fallback
### AGENT-FALLBACK-001. Sequential Fallback Is Mandatory When Parallel Subagents Are Unavailable (Mandatory)
- If parallel subagents are unavailable, the controller must run the controller step, each worker lane, and the audit lane sequentially in that order.
- Sequential fallback must be explicit in the run log; it is not acceptable to imply parallel execution when the lanes were actually run one after another.
- Sequential multi-agent fallback is allowed only when explicit authorization exists but parallel subagents are unavailable or blocked by environment limits.
- Single-agent sequential audit fallback is different from sequential multi-agent fallback: when explicit authorization is absent, it is required for audit attendance and must not be described as spawned multi-agent execution.
- Sequential fallback must be recorded as `agent_mode: sequential-fallback`, not as parallel execution.
- If explicit authorization is absent, do not fake sequential multi-agent execution; record `agent_mode: single-agent-no-auth`, complete the audit lane as `sequential-audit-fallback`, or pause for authorization before any substantial spawned-worker decomposition.

## Task Card Rules
### AGENT-CARD-001. Task Cards Must Lock Lane Inputs Outputs Fallback And Evidence (Mandatory)
- Every lane must have a task card with the lane name, role_alias_zh, lane_alias_zh, owner_alias_zh, spawn_agent_alias_zh, audit_agent_alias_zh, system_agent_id, objective, inputs, outputs, dependencies, owner, authorization_source, agent_mode, spawn_status, fallback_mode, audit_agent_id or sequential_audit_fallback_id, action_audit_scope, action_audit_verdict_cadence, action_audit_verdicts, mutation_audit_scope, mutation_audit_verdicts, and evidence fields.
- In the legacy `audit_agent_id` requirement above, a documented `sequential_audit_fallback_id` satisfies the audit identity field when the run is `single-agent-no-auth` or `sequential-fallback`; do not invent a spawned audit id for fallback records.
- A task card is incomplete if it omits authorization_source, agent_mode, spawn_status, fallback mode, or the evidence requirement for the lane.
- Every substantial run must also externalize one run manifest that records required lanes, spawned agent ids or spawn-skipped reasons, fallback mode, audit verdict cadence, action-level audit verdicts, mutation-level audit verdicts, and final audit verdict.
- Every non-substantial skill behavior must still record an action-level audit entry with action owner, audit owner, action category, checks, verdict, and fallback mode.

## Audit Evidence Rules
### AGENT-EVIDENCE-001. Audit Evidence Must Name Reviewed Artifacts Checks Verdicts And Gaps (Mandatory)
- Audit evidence must name the reviewed files, the command or check that produced the evidence, the verdict, and any blocking gaps.
- Audit evidence must name the reviewed action cycles, including read-only and planning actions that produced no changed files.
- Audit evidence must be durable enough to re-run the same check without guessing the source artifacts.
- Audit evidence must also name the agent run manifest path, authorization source, dispatched agent ids, and audit lane verdict so the claim can be rechecked.
