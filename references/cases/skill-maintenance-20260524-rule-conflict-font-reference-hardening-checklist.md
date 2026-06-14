# Active Checklist

- run id: skill-maintenance-20260524-rule-conflict-font-reference-hardening
- mode: skill-maintenance
- user request: update graduation-project-builder; consolidate rules so they do not conflict; fix repeated errors around missing drawings, invalid thesis format, empty references, bibliography English font, keyword style pollution, and visibly wrong fonts

## Control

- [passed] Explicit skill invocation detected and routed to skill-maintenance.
- [passed] Pre-lock reads recorded as contaminated/reference-only drift.
- [passed] Skill invocation lock created before any mutation in this restarted run.
- [passed] Routed references loaded: user-feedback persistence, maintenance-and-structure, and agent-lanes.
- [passed] Agent manifest and full role roster task cards created before mutation.
- [passed] Parallel/subagent audit recorded with real system agent ids; new spawn attempt hit thread limit and sequential fallback was recorded.

## Consolidation

- [passed] Identify the canonical owner for each repeated defect and remove competing or disconnected rule fragments.
- [passed] Verify rule-owner-map entries for new/changed durable rules.
- [passed] Verify FILE-ROLE-INDEX exposes new/changed active files.
- [passed] Verify templates, generators, validators, and selftests agree on required fields.
- [passed] Keep thesis-format rules from overriding CAD delivery rules and keep skill-maintenance separated from manuscript mutation.

## Frequent Defects To Harden

- [passed] Reference entries that contain only numbering or empty labels must fail citation/bibliography audit and gate.
- [passed] Bibliography English/Latin runs must preserve the required English font and be covered by font/encoding audit evidence.
- [passed] Keywords must not inherit the keyword title style; keyword label and keyword body must be distinguishable.
- [passed] Wrong visible fonts must fail on the relevant protected surface instead of passing via broad style names.
- [passed] Drawing/CAD delivery cannot be treated as complete when the expected drawing package is missing or stale.

## Verification

- [passed] Reproduce current failing selftest or targeted gate failure.
- [passed] Patch the fixture or implementation without weakening fail-closed gates.
- [passed] Run Python compile checks for touched scripts.
- [passed] Run JSON validation for rule-owner-map.
- [passed] Run registry validation.
- [passed] Run UTF-8 clean check.
- [passed] Run targeted selftests for bibliography/font/style false-pass blockers.
- [passed] Run canonical skill bundle gate.
- [passed] Write final audit with commands, results, changed paths, remaining blockers, and audit verdict.
