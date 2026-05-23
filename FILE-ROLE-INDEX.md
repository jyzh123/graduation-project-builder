# File Role Index

This file defines which files under `graduation-project-builder` are active, which are backups, which are archives, and which are historical notes.

Use this index to avoid mixing active rule files with snapshots or old repair branches.

## 1. Active Files

These files are the current source of truth for execution.

### Core entry

- `SKILL.md`
- `memory.md`

Skill boundary rule:

- `SKILL.md` owns mode split, cross-cutting gates, acceptance gates, and routing reminders.
- `SKILL.md` owns the explicit-invocation bootstrap order: after `graduation-project-builder` is invoked, lock/checklist/audit control artifacts must exist before project inspection, helper execution, or handoff.
- execution-layer tool ordering, shell invocation recipes, renderer staging, and lock recovery belong in `references/thesis/thesis-execution-contract.md` or `references/tooling-dependencies.md`
- do not use `SKILL.md` to store tool-specific PowerShell, COM, XML-part, or renderer recipes when a focused reference can own them
- after any skill maintenance edit, consolidate the rule and process into one owner/routing chain; do not leave the active behavior scattered across uncross-linked notes, templates, and scripts

### Active folders

- `agents/`
- `assets/`

### Program references

- `references/program/workflow-standard.md`
- `references/program/verification-matrix.md`
- `references/program/stack-adaptation.md`
- `references/program/executable-automation.md`
- `references/program/packaging-rules.md`

### Thesis references

- Router file:
  - `references/thesis/thesis-format-rules.md`
- Router file:
  - `references/thesis/thesis-figure-generation-rules.md`
- Router file:
  - `references/thesis/thesis-troubleshooting-log.md`
- `references/thesis/thesis-production-workflow.md`
- `references/thesis/thesis-workflow-map.md`
- `references/thesis/thesis-mutation-transaction.md`
- `references/thesis/thesis-format-rules.md`
- `references/thesis/format-rules/general-and-docx-safety.md`
- `references/thesis/format-rules/protected-surface-evidence-contract.md`
- `references/thesis/format-rules/front-matter-and-toc.md`
- `references/thesis/format-rules/headings-and-figures.md`
- `references/thesis/format-rules/tables-abstracts-citations-references.md`
- `references/thesis/format-rules/repair-logging-and-technical-notes.md`
- `references/thesis/thesis-format-sop.md`
- `references/thesis/thesis-execution-contract.md`
- `references/thesis/thesis-format-class-review.md`
- `references/thesis/thesis-figure-generation-rules.md`
- `references/thesis/figure-rules/baseline-and-sourcing.md`
- `references/thesis/figure-rules/review-gates.md`
- `references/thesis/figure-rules/geometry-and-layout.md`
- `references/thesis/figure-rules/workflow-and-checklists.md`
- `references/thesis/thesis-template-learning.md`
- `references/thesis/canonical-thesis-engine-contract.md`
- `references/thesis/thesis-troubleshooting-log.md`
- `references/thesis/troubleshooting/recovery-basics.md`
- `references/thesis/troubleshooting/media-and-figure-recovery.md`
- `references/thesis/troubleshooting/template-and-toc-rebuild.md`
- `references/thesis/troubleshooting/blank-pages-and-end-matter.md`
- `references/thesis/thesis-companion.md`

### Shared references

- Router file:
  - `references/user-feedback-persistence.md`
- `references/user-feedback-persistence.md`
- `references/rule-owner-map.json`
- `references/user-feedback/program-delivery.md`
- `references/user-feedback/thesis-workflow.md`
- `references/user-feedback/content-and-copy.md`
- `references/user-feedback/citations-and-bibliography.md`
- `references/user-feedback/template-and-layout.md`
- `references/user-feedback/final-qa-and-tooling.md`
- `references/user-feedback/maintenance-and-structure.md`
- `references/agents/agent-lanes.md`
- `references/thesis-table-style-memory.md`
- `references/thesis-figure-style-memory.md`
- `references/thesis-formula-style-memory.md`
- `references/thesis-layout-visual-memory.md`
- `references/thesis-source-expansion-and-no-code.md`
- `references/tooling-dependencies.md`

### Review checklists

- `references/review-program-checklist.md`
- `references/review-delivery-bundle-checklist.md`
- `references/review-thesis-format-checklist.md`
- `references/review-thesis-content-consistency-checklist.md`
- `references/review-figure-style-checklist.md`

### Policy references

- `references/policy/cnki-citation-policy.md`

### Templates

- `assets/program-gap-checklist.md`
- `assets/thesis-blueprint-template.md`
- `assets/figure-task-template.md`
- `assets/figure-plan-template.md`
- `assets/format-repair-task-template.md`
- `assets/final-acceptance-template.md`
- `assets/mechanical-cad-acceptance-template.md`
- `assets/skill-invocation-lock-template.md`
- `assets/paper-only-bibliography-review-template.md`
- `assets/review-evidence-template.md`
- `assets/humanizer-evidence-template.md`
- `assets/user-reported-issue-ledger-template.md`
- `assets/agents/agent-task-card-template.md`
- `assets/agents/agent-run-manifest-template.md`

Template responsibility rule:

- execution-critical locks must exist in the templates before they are enforced by the validator
- explicit `graduation-project-builder` invocations must start from `assets/skill-invocation-lock-template.md`; this lock is the first proof that the routed skill path, checklist, helper preflight, mutation allowance, and final gate binding are active
- explicit-invocation bootstrap drift is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-072`; the router is `references/user-feedback-persistence.md`, the validator owner is `scripts/validate_skill_gate_record_gate.py::validate_skill_invocation_lock`, and selftest coverage is in `scripts/selftest_skill_flow.py`
- pagination/reference false-pass closure is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-073`; the validator owners are `scripts/inspect_docx_pagination_structure.py`, `scripts/validate_skill_gate_record_evidence.py`, `scripts/validate_skill_gate_record_core.py`, `scripts/validate_skill_gate_record_gate.py`, `scripts/generate_thesis_acceptance_record.py`, and `scripts/sample_self_check.py`, including the `tail-block.pagination-contract` detector that protects the references opener pagination owner and the rendered prior-block separation proof
- user-reported protected visual false-pass closure is owned by `references/user-feedback/final-qa-and-tooling.md` rule `QA-FINAL-057`; final acceptance must carry template-vs-target rendered geometry evidence plus full-page/key-surface binding fields for reported TOC, abstract/keyword, header, footer, page-number, references, and body-font surfaces, and `scripts/validate_skill_gate_record_gate.py` fails records that rely on XML-only, structure-only, stale, sampled-only, or caveated-pass evidence for those reported surfaces
- thesis content expansion rendered-body false-pass closure is owned by `references/user-feedback/final-qa-and-tooling.md` rule `QA-FINAL-058`, `references/user-feedback/template-and-layout.md` rule `FB-LAYOUT-072`, and `references/agents/agent-lanes.md` rule `AGENT-AUDIT-004`; final acceptance must carry exact-output content-mutation rendered review evidence, machine-vision verdict, inserted-body heading-contamination verdict, touched-page/blast-radius evidence, and format-lane post-mutation audit verdict, and `scripts/validate_skill_gate_record_gate.py` fails XML-only, PDF-export-only, manual-only, stale, sampled-only, or unchecked content-expansion records
- thesis blue/theme-colored visible text closure is owned by `references/user-feedback/template-and-layout.md` rule `FB-LAYOUT-073`; final acceptance must carry an exact-output `scripts/audit_docx_font_color.py` audit path and verdict proving that direct run colors and used style colors are black/auto or explicitly template-authorized
- mechanical CAD package handoff closure is owned by `references/thesis/figure-rules/baseline-and-sourcing.md` rule `CORE-FIGURE-010`; CAD-only drawing deliveries may use `assets/mechanical-cad-acceptance-template.md`, and `scripts/validate_skill_gate_record_gate.py::check_mechanical_cad_acceptance_record` must bind the exact final package, audited CAD package, DWG package, combined PDF, strict v4 audit JSON, rendered no-overlap review, boundary-clearance/detail-density/title-block-table-notes-isolation/annotation-margin/local-crowding verdicts, and entity-count-only false-pass rejection before `validate_skill_gate.py --gate-record` can pass
- body opener/header title consistency closure is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-074`; the validator owners are `scripts/validate_skill_gate_record_gate.py::check_gate_record` and `scripts/validate_skill_gate_record_gate.py::validate_body_opener_header_title_evidence`, with selftest coverage in `scripts/selftest_skill_flow.py`
- protected visual surface false-pass closure is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-075`; `scripts/sample_self_check.py`, `scripts/generate_thesis_acceptance_record.py`, and `scripts/validate_skill_gate_record_gate.py` must share the gate-required detector set, including `header-footer.page-number-template-contract`, and selftests must reject limitation/caveat wording and missing rendered-geometry evidence for abstract/header/footer/page-number complaints
- whole-thesis DOCX structural release closure is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-076`; final acceptance must carry an exact-output `scripts/audit_docx_whole_format_gate.py` JSON path and verdict proving section topology, TOC implementation, header/footer PAGE fields, page-number chain, surface order, builder-style contamination, and style-binding risk on the handed-off DOCX SHA
- explicit skill invocation anti-bypass fields belong in both `assets/skill-invocation-lock-template.md` and `assets/final-acceptance-template.md`; they must prove source type, active skill takeover, prohibited-bypass review, canonical gate requirement, no narrow/smoke substitute, failed-evidence escalation, no mutation before lock, and final handoff allowance
- if a new gate field is added to validation, update the owning template in the same turn
- thesis content-writing planning fields such as humanizer routing belong in `assets/thesis-blueprint-template.md`
- run-level agent authorization, lane dispatch, Chinese role aliases, system agent id mapping, fallback, and audit supervision evidence belong in `assets/agents/agent-run-manifest-template.md`
- protected-surface ids, effective font-chain proof, all-surface paragraph-dialog / typography proof, TOC paragraph-and-typography proof, per-surface evidence maps, owner maps, reviewed output hashes, and agent audit handoff fields belong in `references/thesis/format-rules/protected-surface-evidence-contract.md` and must appear as fields in the task, manifest, task-card, evidence, and acceptance templates before enforcement
- review comments/change marks, comment anchors, bookmarks, fields/hyperlinks, and body citation superscript runs are protected surfaces; source inventories, final diffs, explicit disposal approval, and exact final DOCX SHA binding must be present in templates before validators can pass them
- empty-paragraph bookmark disposal is owned by `references/user-feedback/maintenance-and-structure.md` rule `EXEC-MAINT-059` and enforced by `scripts/audit_docx_review_artifacts.py`; only SHA-bound `empty-paragraph-bookmark-disposition` records that prove no text, field, hyperlink, comment, tracked-change, image/table, break, or citation host was removed may convert a missing bookmark from blocking to controlled
- comment-resolution ledgers are required for teacher/user comment completion claims; final acceptance must carry the ledger path, audit report path, audit verdict, source/final DOCX bindings, and per-comment evidence before `all comments resolved` wording can pass
- final execution verdict fields such as humanizer evidence, surface ownership locks, and smoke-audit summaries belong in `assets/final-acceptance-template.md`
- release/baseline promotion fields for scoped thesis repairs belong in `assets/final-acceptance-template.md` and `assets/format-repair-task-template.md`; they must expose baseline promotion status, promotion evidence, release blocker ledger, unresolved blocker count, and scoped artifact next-baseline verdict before any candidate can be treated as the next handoff baseline
- paragraph-level humanizer proof such as before text, after text, target language, paragraph ID, and skill name belongs in `assets/humanizer-evidence-template.md`

### Scripts

- `scripts/extract_thesis_template.py`
- `scripts/discover_project_thesis_template.py`
- `scripts/check_utf8_clean.py`
- `scripts/audit_paper_only_bibliography.py`
- `scripts/audit_thesis_comment_resolution.py`
- `scripts/audit_thesis_citations.py`
- `scripts/audit_docx_body_style.py`
- `scripts/audit_docx_frontmatter_structure.py`
- `scripts/audit_docx_table_structure.py`
- `scripts/audit_docx_figure_extents.py`
- `scripts/audit_mechanical_drawing_package.py`
- `scripts/audit_docx_font_color.py`
- `scripts/audit_docx_whole_format_gate.py`
- `scripts/audit_docx_font_encoding.py`
- `scripts/repair_bibliography_entry_format.py`
- `scripts/repair_thesis_reference_content.py`
- `scripts/python_runtime.py`
- `scripts/generate_compatibility_export_notes.py`
- `scripts/validate_skill_gate_bundle.py`
- `scripts/validate_skill_gate_records.py`
- `scripts/validate_skill_gate_record_core.py`
- `scripts/validate_skill_gate_record_evidence.py`
- `scripts/validate_skill_gate_record_format.py`
- `scripts/validate_skill_gate_record_gate.py`
- `scripts/validate_skill_gate_registry.py`
- `scripts/validate_skill_gate_registry_bundle.py`
- `scripts/validate_skill_gate_registry_core.py`
- `scripts/validate_skill_gate_registry_records.py`
- `scripts/validate_skill_gate_utils.py`
- `scripts/generate_acceptance_report.py`
- `scripts/generate_delivery_scaffold.py`
- `scripts/generate_project_blueprint.py`
- `scripts/generate_thesis_acceptance_record.py`
- `scripts/generate_thesis_blueprint.py`
- `scripts/generate_validate_gate_registries.py`
- `scripts/measure_toc_paragraph_typography.py`
- `scripts/export_thesis_drawio_figure.py`
- `scripts/validate_structural_figure_geometry.py`
- `scripts/thesis_template_profile.py`
- `scripts/thesis_figure_contract.py`
- `scripts/repair_thesis_frontmatter_toc_structure.py`
- `scripts/repair_front_matter_page_numbering.py`
- `scripts/build_canonical_thesis.py`
- `scripts/strip_docx_template_instruction_artifacts.py`
- `scripts/inspect_docx_pagination_structure.py`
- `scripts/audit_docx_formula_objects.py`
- `scripts/build_minimal_template_thesis.py`
- `scripts/collect_heading_pages_word.py`
- `scripts/copy_locked_docx.py`
- `scripts/docx_apply_table_family.py`
- `scripts/docx_formula_number_table.py`
- `scripts/docx_sync_picture.py`
- `scripts/normalize_thesis_citation_chain.py`
- `scripts/repair_thesis_surface_format.py`
- `scripts/repair_thesis_comment_content_surfaces.py`
- `scripts/repair_docx_picture_display_extents.py`
- `scripts/pdf_to_pages.py`
- `scripts/sample_self_check.py`
- `scripts/scan_project_local_thesis_helpers.py`
- `scripts/update_static_toc.py`
- `scripts/validate_thesis_local_adapter.py`
- `scripts/validate_thesis_mutation_transaction.py`
- `scripts/selftest_skill_flow.py`
- `scripts/run_integration_gate.py`
- `scripts/validate_skill_gate.cmd`
- `scripts/validate_skill_gate.py`
- `scripts/wps_export_pdf.ps1`

Script responsibility rule:

- `scripts/validate_skill_gate.py` is the canonical text gate for bundle structure, routing coverage, and evidence-record completeness
- `scripts/python_runtime.py` centralizes active-interpreter resolution for Python-driven validation, selftest, and rebuild helpers
- `scripts/discover_project_thesis_template.py` is the canonical template-candidate discovery helper for thesis format lanes; it identifies candidate DOCX/DOC templates before mutation and must not infer format from a broken manuscript.
- `scripts/extract_thesis_template.py` is a compatibility template-extraction helper; new production template facts should flow through `scripts/thesis_template_profile.py`.
- `scripts/audit_docx_font_encoding.py` is the canonical DOCX font-slot and encoding audit for exact thesis outputs, including bibliography mixed-run font drift, bibliography leading-number baseline checks that reject `w:vertAlign` superscript/subscript on visible `[n]` entry labels even without a template DOCX, TOC-safe bibliography block detection, direct/style/effective WPS font-chain validation, exact DOCX/reference SHA binding, positive entry/run coverage, bound bibliography content-format model checks (`FB-CITE-041`), template-derived bibliography size enforcement, and explicit Chinese named-size half-point enforcement such as 五号=`21` with WPS named-size evidence when required. It must reject explicit half-point overrides that conflict with a bound template/reference bibliography donor.
- `scripts/audit_docx_font_color.py` is the canonical DOCX visible font-color audit and bounded repair helper for exact thesis outputs. It audits direct run colors plus the colors of actually used paragraph/run styles in `styles.xml` and `stylesWithEffects.xml`, rejects blue/theme-accent thesis text unless template-authorized, and can write a repair output that forces those used colors to black while clearing theme-color attributes.
- `scripts/audit_docx_whole_format_gate.py` is the canonical whole-thesis DOCX structural release gate for exact thesis handoffs. It audits the handed-off DOCX package for section topology, front/body/end surface order, live or template-authorized TOC evidence, footer PAGE fields, front/body page-number chains, header/footer part binding, builder-owned style contamination, and excessive unstyled body paragraphs. It is not a replacement for rendered page-class review, whole-document pagination JSON, font-color audit, or body-style audit; it is the fail-closed structural gate that prevents those narrower checks from passing a globally broken manuscript.
- `scripts/audit_wps_reference_entry_ui_font.ps1` is the canonical WPS/Word COM read-only UI-font evidence producer for bibliography named-size requirements; it selects each reference-entry paragraph, records the displayed/inferred named size, per-entry font fields, exact DOCX SHA256, and emits `graduation-project-builder.wps-reference-entry-ui-font.v1` JSON for the DOCX font audit.
- `scripts/repair_bibliography_entry_format.py` is the canonical bounded reference-entry format replay helper for repairing already-collapsed bibliography entries from a locked template donor; it may only rewrite bibliography entry paragraphs inside `word/document.xml`, must preserve visible entry text and existing citation bookmarks, must lock reference-entry font families from direct donor slots or template instruction policy, must support explicit named-size repair such as 五号=`21`, and must report package drift before handoff.
- `scripts/repair_thesis_reference_content.py` is the canonical bounded reference-content replacement helper for verified bibliography/comment repairs. It may only replace exact planned text in `word/document.xml`, must write a source/final SHA-bound report, must prove only `word/document.xml` changed, and must not rebuild citation markers, bookmarks, comments, relationships, media, styles, numbering, headers, footers, or whole bibliography formatting.
- `scripts/audit_thesis_citations.py` is the canonical body-citation and bibliography linkage audit for thesis DOCX outputs. Its report must include the exact audited document path and document SHA256 so stale path-only citation audits cannot satisfy final acceptance, and it must verify that citation hyperlinks resolve to bibliography entries inside the `参考文献` block rather than to cover or front-matter bookmarks.
- `scripts/audit_docx_review_artifacts.py` is the canonical source-to-final preservation auditor for DOCX review artifacts and body citation runs; it reopens the source and final DOCX packages to compare `word/comments*.xml`, people parts, comment anchors by anchor type and count, tracked changes including property-change elements, bookmarks, fields, hyperlinks, and citation marker run state across main document, header, footer, footnote, endnote, and comment story parts. Citation hyperlink preservation is source-relative: a final marker is failed for lost hyperlink host only when the source marker had a hyperlink/bookmark host. Missing bookmarks remain blocking unless `--controlled-bookmark-disposition` provides a SHA-bound empty-paragraph disposition that proves the missing anchor had no visible or functional host. Its validators must reject pass-shaped or stale reports even when path/SHA fields look current. Use its `--json-output` option for persistent JSON reports so Windows shell redirection cannot create UTF-16 or mojibake audit records.
- `scripts/audit_thesis_comment_resolution.py` is the canonical semantic comment-closure auditor. It reads `word/comments.xml`, `word/commentsExtended.xml`, comment anchors across story parts, source/final DOCX SHA256 values, and a JSON comment-resolution ledger. It fails closed when done/resolved state changes lack ledger authorization, all-comments-resolved claims leave any open final DOCX comments or open/partial/blocked/missing/orphan ledger rows, comment text is altered without disposal approval, or figure comments are closed by size-only evidence while crop/provenance/redraw/content/readability subissues remain.
- Fixed comment-resolution rows must also bind each claimed fix to a detector surface, subissue, detector id, detector report path, and pass verdict for the exact final DOCX; a prose note or checked Word/WPS done state is not enough.
- `scripts/close_docx_comments_from_ledger.py` is the bounded comment-state closer. It may only change `word/commentsExtended.xml` in a new DOCX path after a source-bound fixed comment-resolution ledger exists; it must not remove comment text, anchors, relationships, tracked changes, media, or document body content. Its output still requires a fresh comment-resolution audit and SHA-bound detector reports before handoff.
- `scripts/audit_paper_only_bibliography.py` is the canonical literature-source classifier for paper-only bibliography requirements.
- `scripts/generate_compatibility_export_notes.py` regenerates the compatibility retirement and external-audit notes from validator registry metadata
- `scripts/validate_skill_gate_bundle.py`, `scripts/validate_skill_gate_records.py`, `scripts/validate_skill_gate_record_core.py`, `scripts/validate_skill_gate_record_evidence.py`, `scripts/validate_skill_gate_record_format.py`, `scripts/validate_skill_gate_record_gate.py`, `scripts/validate_skill_gate_registry.py`, `scripts/validate_skill_gate_registry_bundle.py`, `scripts/validate_skill_gate_registry_core.py`, `scripts/validate_skill_gate_registry_records.py`, and `scripts/validate_skill_gate_utils.py` are the split validator implementation and must stay import-compatible with the compatibility aggregators; `record_core` owns per-report schemas including DOCX/font audit SHA, bibliography content-format model binding, citation audit final DOCX SHA, and WPS named-size evidence checks; `record_gate` owns final acceptance semantics, scoped-repair release/baseline promotion closure, and cross-record consistency between final acceptance, agent manifest/task cards, protected review artifacts, body citation run evidence, and the filled format-repair task record; `record_format` owns format-task-card semantics including baseline promotion gate and release blocker ledger checks, `record_evidence` owns review evidence semantics, robust body/abstract direct-surface classification, acknowledgement-title indentation/position drift rejection, acknowledgement-body paragraph/typography/title-contamination drift rejection, exclusion of TOC rows and figure/table captions from `body_text`, inline abstract body detection such as `摘 要：...` / `Abstract: ...`, and the registry files own required-line/schema data.
- compatibility-only exports such as `SCRIPT_RUNTIME_FORBIDDEN_TOKENS_BY_FILE` and `SEMANTIC_ONLY_RULE_FILES` must stay explicitly marked with replacement source, retention tier, external caller retirement status/evidence, direct importer inventory, retirement checklist, and removal condition metadata and, where applicable, mirror their active source-of-truth structures
- `scripts/generate_validate_gate_registries.py` regenerates the split validator registry files in deterministic ASCII-safe form
- `scripts/validate_thesis_local_adapter.py` validates project-local thesis adapter manifests; generic thesis-making behavior remains owned by canonical skill scripts
- `scripts/validate_thesis_mutation_transaction.py` validates thesis mutation transaction records for required evidence schemas, target surfaces, one write owner, non-target protected-surface changes, local-surface scope claims, TOC page-number right-edge metrics, exact source/final/template/review DOCX SHA binding, distinct source/final DOCX paths, and strict figure-manifest contract enforcement for image mutations. It treats source-to-final media relationship diffs and drawing-object diffs as image mutations, including changes to drawing size, inline/anchor mode, relationship set, and caption adjacency, while ignoring pure paragraph-index key shifts when stable story/media/extent/caption signatures match. It distinguishes mutation intent from protected sibling freeze lists: `protected_sibling_surfaces` records blast-radius surfaces but must not by itself trigger image manifest requirements, fixed-locator rejection, comment-driven behavior, whole-thesis-claim rejection, or chapter format-preservation detector requirements when the actual target/operation text is a non-body local surface such as keyword run structure and no format-preservation promise is made. Use `--json-output` for persistent JSON validation reports so Windows shell redirection cannot create UTF-16 or mojibake records. It passes the transaction source/final DOCX and manifest path into the canonical figure contract so source/final media and drawing comparison plus manifest-relative path resolution cannot be bypassed by pass-shaped transaction fields.
- `scripts/thesis_figure_contract.py` owns figure/screenshot/diagram manifest validation, source-to-final media relationship comparison, drawing-object preservation checks, runtime-screenshot authenticity checks, structural-figure provenance checks, and figure-surface preservation logic. It treats preceding body-prose rewrites before an otherwise unchanged figure as text changes rather than figure mutations; media relationships, drawing extents, caption text, and caption adjacency remain protected figure surfaces.
- `scripts/audit_docx_figure_extents.py` is the final-DOCX display-size audit for inserted thesis figures. It enforces readable-width floors and the `FB-LAYOUT-071` paragraph-margin width rule by reporting body text width, width/text-width ratio, and `paragraph_margin_width_drift_count`; body figures below the default ratio are compressed evidence even when they exceed the `8.0 cm` or `9.0 cm` floors.
- `scripts/build_canonical_thesis.py` is the canonical general thesis builder for validated local adapter/content manifests; it owns generic template-driven DOCX assembly, active template SHA256 lock enforcement, cover title replacement, cover identity value-run replacement, template-square-to-space front-matter normalization, language-bounded abstract prefix/content-run replacement, keyword run preservation, TOC paragraph cloning, rendered logical page sync, template-donor TOC run-font restoration, table-cell content-family donor selection, figure follow-up caption deconfliction, body chapter pagination owners, rendering, self-check, and acceptance chaining. Its self-check and acceptance calls must pass the source/template DOCX through `--source-docx`, and its figure manifest completion path must write top-level source/final DOCX path and SHA256 bindings before final figure contract validation.
- `references/thesis/thesis-workflow-map.md` is the canonical thesis process router for new full thesis production, whole-thesis revision, local surface repair, content-only paragraph revision, and audit-only lanes; other thesis rule files provide surface details after this router selects the workflow.
- `references/thesis/thesis-mutation-transaction.md` is the canonical transaction owner for DOCX-writing thesis workflows; it owns source/template/review/final path locks, protected-surface freeze manifest, post-mutation diff, rendered target/blast-radius review evidence, cross-surface regression report, source-to-final review-artifact and body-citation run preservation evidence, write-owner proof, local-vs-whole pass scope, and final DOCX SHA binding.
- `scripts/thesis_template_profile.py` is the canonical template-profile extractor for front-matter page-class occupancy, same-page title groups, separated front-matter pairs, and critical surface markers used by generation and self-check.
- `scripts/thesis_figure_contract.py` is the canonical figure-manifest contract for figure family inference, final-DOCX figure-surface detection, adjacent image-caption block parsing, official-caption recognition that separates `图3-1  标题` / `图3-1：标题` from explanatory prose such as `图3-1展示...`, incomplete caption-name rejection, runtime screenshot route/capture/readiness/accepted-asset evidence, runtime screenshot full-window geometry and cropped top-left rejection, widget render/grab substitute rejection, source-preserved runtime screenshot evidence rejection, runtime/algorithm rendered image blank, solid-color-block, and purple-placeholder rejection, algorithm-result accepted image provenance and authenticity evidence, draw.io/SVG/raster fallback closure, flowchart terminator checks, ER source-scale collision checks, inserted-scale and dense-zone evidence enforcement, manifest coverage for structural figure signals, source geometry report enforcement, and SVG-primary DOCX relationship validation. When a final DOCX or manifest indicates figure/image/replacement surface, it fails closed unless the manifest itself binds top-level source/final DOCX paths and SHA256 values; it resolves manifest-relative asset paths, scans `word/**/_rels/*.rels` across body/header/footer/comment/footnote/endnote story parts, verifies original and final media rid/target/SHA bindings, fails on missing media package targets or missing media hashes, traverses nested body/table/SDT/textbox paragraphs for front-matter or TOC drawings while allowing source-preserved official front-matter images, rejects DrawingML and VML/WPS `w:pict` image shapes wider than the available text width, compares source/final drawing-object manifests for size, inline/anchor, relationship-set, and caption-adjacency drift, rejects unauthorized media/drawing changes/removals/additions, and requires owner part disambiguation for duplicate relationship identities.
- Figure display-height is part of the same figure contract: body and non-body story drawings must fail when their final extent exceeds the safe page-height occupancy threshold, even if their width fits the text area and media bytes are unchanged.
- `scripts/repair_thesis_frontmatter_toc_structure.py` is the canonical package-preserving helper for boundary-locked front-matter/TOC structure recovery. It supports only explicit operations `toc-contamination`, `duplicate-page-breaks`, `abstract-style`, `heading1-baseline`, `reference-residue`, and `frontmatter-signature-residual`; it may change only `word/document.xml` and, when a missing abstract-title style is created, `word/styles.xml`; it must preserve chapter `w:pageBreakBefore` owners during Heading1 replay and must not touch media, relationships, comments, tracked changes, headers, footers, figure captions, table bodies, or bibliography entries. The locked `duplicate-page-breaks` operation may remove explicit page-break runs from level-1 body headings that already own `w:pageBreakBefore`, including later chapter headings, while preserving the paragraph-owned break. The locked `reference-residue` operation repairs the bad repair-note wording and relocates the existing citation paragraph after the final pre-existing body citation while preserving source citation marker runs and hyperlink anchors. The locked `frontmatter-signature-residual` operation may remove only unanchored blank declaration spacer paragraphs and tighten declaration-title spacing while preserving bookmarks, fields, comments, drawings, section breaks, and signature text.
- `scripts/repair_front_matter_page_numbering.py` is the canonical package-preserving helper for front-matter/body page-number section repair. It may write only a new review-copy path, may replace only `word/document.xml`, inserts or updates lower-roman front-matter section numbering from the Chinese abstract through the TOC, restores decimal body numbering from the first body chapter, and removes only the duplicate first-body `pageBreakBefore` that is made redundant by the new section break. It must fail closed when the Chinese abstract or first body heading cannot be located.
- `scripts/strip_docx_template_instruction_artifacts.py` is the canonical reusable DOCX helper for removing visible template instruction callouts, text boxes, arrows, and front-matter note residue from generated or repaired manuscripts while preserving real cover/declaration/title tables unless an isolated instruction table has explicit delete/fill markers and no thesis field values.
- `scripts/run_integration_gate.py` is the canonical DOCX integration gate for real `officecli` / renderer / file-lock / rendered-page behavior; `--quiet` keeps stdout to a machine-readable summary while detailed logs stay in the report file. Its sample-self-check lane must pass a rendered reference PDF when available and include regression cases for abstract manual line breaks and rendered bibliography geometry drift.
- `scripts/export_thesis_drawio_figure.py` is the canonical bounded draw.io export helper for thesis structural figures and must remove draw.io fallback leakage such as `Text is not SVG - cannot display` before DOCX insertion
- `scripts/validate_structural_figure_geometry.py` is the canonical source-scale geometry validator for ER and non-ER structural figures, including flowcharts and use-case diagrams; it writes the figure geometry report consumed by `thesis_figure_contract.py` and final acceptance evidence, and rejects source/targetless, non-orthogonal, rounded/curved, invisible-router, and through-node connector routes.
- `scripts/build_minimal_template_thesis.py` is the canonical direct-template shortest/manual-review thesis builder and must preserve draw.io structure provenance, explicit body-style binding, formula numbering, citation normalization, and the downstream self-check / acceptance chain
- `scripts/sample_self_check.py` is the canonical thesis self-check detector runner for generated, rebuilt, and repaired manuscripts; it must not be bypassed for any handed-off thesis DOCX and must block delivery when figure/image surfaces lack a manifest, structural figures are missing diagram entries, abstract body paragraphs contain manual `w:br` line breaks, TOC visible runs gain underline/hyperlink-style pollution, reference-entry rendered geometry drifts from the template/reference PDF, image dimension checks are disabled by wording while body images exist, cover identity value-line detection binds to bibliography legends such as `文献类型`, or abstract/keyword donor evidence is polluted by TOC/template-instruction/reference rows. It must emit required Detector Registry entries for `heading.baseline-contract`, `toc.visible-format-contract`, and `figure.family-style-contract`, and `scripts/validate_skill_gate_record_gate.py` must reject final records when those entries are missing or lack evidence. Its figure contract detector accepts `--source-docx` and must pass manifest-bound source DOCX, final DOCX, and manifest path into `thesis_figure_contract.py` so media diffs cannot be validated from final-only or reference-only evidence.
- `scripts/generate_thesis_acceptance_record.py` is the canonical acceptance-record generator for thesis outputs and must point every evidence field at the exact reviewed output path; it owns page-class coverage matrix generation with class-bound target identifiers and per-class evidence paths, and accepts an explicit real project root for project-local helper preflight scanning. It must populate source/final review-artifact inventory/diff fields, source/final body-citation run inventory/diff fields, citation audit final DOCX SHA256, citation source-to-final run diff fields, and explicit skill-invocation anti-bypass fields in both the generated lock and final acceptance record. It passes `--source-docx`, final DOCX, and asset manifest path into the figure contract so final acceptance cannot validate figure media using final-only evidence. If the source or final DOCX carries comments, it must require `--comment-resolution-ledger`, run the comment-resolution audit against source/final DOCX, and record ledger path, audit report path, and verdict before any final pass. It must parse protected-surface verdict fields instead of synthesizing pass rows from evidence existence, count inventory, or broad citation/bibliography status, and it must expose acknowledgement-body paragraph-dialog, direct-run typography, first-line indent, and title-contamination verdict fields separately from the generic acknowledgement-body verdict. It must not hide scanner metadata: project-local helper preflight reports carry schema, project root, scanner command, timestamp, exit status, and risk count.
- `scripts/measure_toc_paragraph_typography.py` is the canonical producer for TOC paragraph-dialog and typography JSON; it compares the template TOC title and each used level against the exact final DOCX for style id/name, spacing, line spacing, indentation, tabs/leaders, paragraph-level font size, visible-run direct `w:rPr` / `w:rFonts` / size / weight for entry text, tab/leader, and page-number runs, plus scale/compression verdict before acceptance generation.
- `scripts/inspect_docx_pagination_structure.py` is the canonical DOCX structure parser for whole-document pagination evidence; it compares template and final `sectPr`, `pgNumType`, header/footer references, inferred link-to-previous behavior, hard page breaks, section counts, footer PAGE-field containers, rendered page-count drift, and blank/near-empty rendered pages across front matter, body, and tail pages before acceptance generation may claim pagination pass. Hard page-break or section-break drift is fatal even when content growth is otherwise allowed.
- `scripts/repair_front_matter_page_numbering.py` is the bounded canonical helper for front-matter/body page-number section repair. It writes only a new review-copy path, replaces only `word/document.xml`, accepts both standalone Chinese abstract titles and inline `摘 要：...` labels as the front-matter start sentinel, and fails closed when the Chinese abstract or first body heading cannot be located.
- `scripts/repair_footer_page_numbers.py` is the bounded canonical helper for footer PAGE-field repair. By default it rewrites only existing active default footer parts referenced by numbered sections; only with explicit `--create-if-missing` may it create one default footer part on the final section, add the matching relationship/content-type override, add decimal `pgNumType`, and optionally copy portrait template `pgMar` via `--apply-template-margins`. It must preserve body text, comments, images/media, styles, bibliography text, and TOC entries, and its selftest covers OOXML namespace serialization so generated packages remain Office/LibreOffice-loadable.
- `scripts/measure_surface_hardfields.py` is the canonical protected-surface hard-field measurement helper for rendered geometry plus DOCX paragraph typography; its verdicts must come from numeric template-vs-target comparisons and must fail on typography, font-size, position, crop-size, ink-ratio, or blank-crop drift. It owns formal-tail-block binding for `references_title` and `references_entries`, must exclude TOC rows from reference evidence, must not reconcile reference-title/reference-entry first-line, hanging, target, title-as-entry, font-size, or weight drift into a pass-shaped result, and must make `--fail-on-drift` fail on paragraph/typography drift as well as rendered-geometry drift. For `figure_table_captions_and_holders`, missing formal figure/table caption donors are not allowed to collapse to prose such as `no formal captions found`; the helper must emit numeric template/actual paragraph and typography fallback fields from selected surface metrics or fail closed.
- `scripts/audit_docx_body_style.py` is the canonical body-style and Normal-baseline audit for DOCX thesis outputs; it also rejects body prose paragraphs that inherit heading-style or outline-level formatting and binds the report to the exact final DOCX SHA256. Its caption filter must treat `图3-1展示...`, `图4-4中的...`, and similar explanatory figure-reference paragraphs as body prose, not captions, so these paragraphs remain subject to body line-spacing, indentation, style-binding, and mixed-script font checks. In strict mode it must also check real body heading level alignment/spacing/indent/direct run size, require formal figure captions to be immediately preceded by an image-holder paragraph, normalize school-template font alias lists only on the reference side, and still reject semicolon alias lists in the final DOCX.
- `scripts/repair_docx_font_alias_slots.py` is the package-preserving global font-alias audit/repair helper for final DOCX outputs. It scans all `word/*.xml` parts for semicolon font alias lists in `w:rFonts`, reports exact part/attribute bindings, and can write a new review copy with each slot reduced to one concrete font family so template-export aliases such as `宋体;SimSun` never remain in the manuscript.
- Body-style audits must fail closed when instance-level body paragraph spacing and indentation evidence is too sparse or lacks hard metrics; sparse template samples cannot silently excuse whole-body line-spacing drift.
- `scripts/audit_docx_frontmatter_structure.py` is the canonical front-matter structure and abstract/keyword hard-field audit. It must expose exact paragraph metrics for Chinese/English abstract and keyword surfaces, including line spacing, first-line indent, label/content run split, label bold state, and content-run weight, so English abstract indentation and keyword label/content separation cannot pass from broad structure-only evidence.
- `scripts/audit_docx_table_structure.py` is the canonical non-mutating DOCX table structure audit. It must enumerate every formal body table, verify title binding, visible three-line border structure, header repeat, row split policy, cell 5号 size, and cell-indent cleanup before a table-style repair can pass.
- `scripts/audit_docx_figure_extents.py` is the canonical final-DOCX figure display-size and caption-adjacency audit for comment-driven figure-size repairs. It distinguishes real body figures from front-matter drawings and narrative figure-reference prose, binds source/final DOCX path and SHA, rejects oversized width/height, structural figures below the strict readable width/height thresholds (`structural_min_readable_width_cm` defaults to 9.0), non-structural/runtime screenshot figures below the strict readable thresholds (`min_readable_width_cm` defaults to 8.0 and `min_readable_height_cm` defaults to 4.0), or blank-bottom images, and requires immediate formal captions plus nearby explanatory prose when requested.
- `scripts/audit_mechanical_drawing_package.py` is the canonical mechanical drawing package gate for graduation-design CAD deliverables. It checks exact DWG/PDF/DXF package contents, real DWG/PDF headers, UTF-8 JSON manifests, mojibake-like text without over-broad normal-Chinese false positives, DXF CAD table/style/dimension structure, A0-equivalent sheet workload when PDFs are parseable, distributed per-sheet DXF/PDF drawing density, required mechanical-detail tokens, DWG byte-density parity against user-provided sample packages, and manifest-bound rendered-sheet review. Its v4 schema must expose `density_verdict`, `manufacturing_depth`, and `rendered_review_verdict` so sketch-level drawings, text/table/frame overlap, unreadable rendered annotations, missing manufacturing-view depth, boundary/clearance failures, title-block/table/notes crowding, and entity-count-only false-passes fail instead of passing from file counts or entity counts alone.
- `scripts/audit_docx_formula_objects.py` is the canonical formula-authenticity audit for DOCX thesis outputs; it counts OMML equation objects, rejects ordinary text paragraphs that look like standalone formulas, and fails when a formula lead-in such as `下式` / `公式` / `表示为` is followed by explanation text instead of a formula surface, so final acceptance cannot pass from text that only visually imitates an equation or promises a missing equation.
- `scripts/scan_project_local_thesis_helpers.py` is the canonical hard preflight scanner for risky project-local thick thesis helper scripts. Thesis acceptance and format-task records must carry its report path, risk count, disposition, and canonical-source restart status before validator pass; the report itself must be schema/versioned and include project root, scanner command, generated timestamp, scanner exit status, and risk count.
- `scripts/normalize_thesis_citation_chain.py` is the canonical bounded citation-chain normalizer; it may patch citation markers, move existing bibliography paragraphs, patch visible leading reference numbers, and add invisible citation bookmarks, but it must fail closed rather than collapsing bibliography entries or multi-run body paragraphs into one rebuilt run. It is not allowed to own final bibliography font policy without the reference font-slot audit. `scripts/audit_thesis_citations.py` owns the final check that visible manual `[n]` bibliography numbers are not combined with paragraph-level Word automatic numbering.
- `scripts/repair_thesis_surface_format.py` is a bounded thesis surface-format repair helper for contaminated recovery passes; it may only replace locked abstract body/keyword text from a plan, replay template-donor paragraph/run hard fields onto abstract and body prose paragraphs, normalize locked body-heading direct metrics when `normalize_body_headings` is explicitly enabled, and clean image-holder paragraph residue. Keyword repair must reject TOC/template-instruction/reference-legend donors and record target abstract-body fallback when used. It must not independently resize images or mutate drawing extents; display-size repair belongs to a transaction/figure-manifest wrapper with source/final drawing hashes. It must not rewrite TOC entries, page numbering, citations, bibliography entries, headers, footers, or whole-document structure.
- `scripts/repair_thesis_comment_content_surfaces.py` is a bounded comment-content repair helper for existing review-copy manuscripts. It may apply locked paragraph text/format, keyword, caption, reference-entry, and figure-surface plans only when anchors are unique and the target surface matches the plan. Media replacement, image insertion, or display-extent mutation requires `--source-docx`, `--transaction-record`, and `--figure-manifest` and validates the figure manifest after mutation. Global `all_ascii_runs` font repair is blocked unless the plan carries a surface allowlist or the caller passes explicit global authorization.
- `scripts/repair_docx_picture_display_extents.py` is the bounded display-extent helper for teacher-comment image-size repairs. It requires the input DOCX to match `--source-docx`, requires `--transaction-record` and `--figure-manifest`, validates transaction binding to source SHA, intended output path, and manifest path before mutation, and re-runs the canonical figure contract after writing the output DOCX.
- `scripts/repair_template_surface_baselines.py` is the canonical bounded helper for replaying locked template surface baselines onto an existing thesis review copy. For `--surfaces cover`, it owns cover paragraph donor replay and cover identity table donor replay: the helper clones the active template cover table structure/geometry/cell formatting and refills only target-specific value cells, so cover tables do not inherit body-table or body-paragraph formatting.
- `scripts/update_static_toc.py` is a compatibility CLI for static TOC page synchronization; canonical builders or repair lanes must own the TOC policy and call this only with locked baseline evidence. When used, it must restore locked template TOC direct `w:rPr` on visible entry text, tab/leader, and page-number runs, including nested WPS run wrappers, before TOC paragraph/typography acceptance; it must normalize fullwidth-dot numbered headings before heading-level detection, require mapped tail entries such as references and acknowledgement to already exist, and refuse one-run TOC rows that lack a separable tab, leader, or trailing page-number segment instead of replacing the whole entry with a page number.
- `scripts/audit_toc_rendered_page_sync.py` is the fail-closed TOC page-number validator that compares static visible TOC page numbers against a rendered heading-page map from `collect_heading_pages_word.py` or equivalent rendered PDF evidence. It catches stale TOC tail pages after figure/body pagination changes, including references and acknowledgement page drift, and must be rerun after any mutation that can change page flow.
- `scripts/docx_sync_picture.py`, `scripts/docx_formula_number_table.py`, and `scripts/docx_apply_table_family.py` are bounded DOCX repair helpers for one protected surface family per call. `docx_sync_picture.py` may replace a picture only when `--figure-manifest` and `--source-docx` are supplied, must target the dedicated image paragraph immediately before the verified caption or a verified image paragraph id, must preserve the original media relationship target/extension, must write final source/final DOCX and media rid/target/SHA bindings back into the manifest after mutation, and must re-run the figure manifest contract after mutation. `docx_apply_table_family.py` also owns the template-donor table hard-field copy path for table repair: body-table range protection, immediately adjacent standalone table-title enforcement, title paragraph/run properties, table properties, table grid, row properties, cell properties, cell paragraph properties, cell run properties, schema-ordered `tblBorders`/`tcBorders` insertion when rebuilding border surfaces, table reference separator normalization, cell-contamination blocking, and collision-safe cloning of template table styles when a target document already uses the donor style id for a different definition or style type.
- `scripts/collect_heading_pages_word.py`, `scripts/copy_locked_docx.py`, `scripts/pdf_to_pages.py`, `scripts/wps_export_pdf.ps1`, and `scripts/inspect_docx_pagination_structure.py` are execution-layer helpers for rendered review, lock-safe staging, and page-image evidence. `inspect_docx_pagination_structure.py` must fail closed on raw unexpected near-empty pages unless each page is explicitly allowlisted with page-class/root-cause evidence. `collect_heading_pages_word.py` must normalize fullwidth dot separators before collecting body heading levels so static TOC sync does not miss headings such as `1 U+FF0E 1` or drop final references/acknowledgement page rows.
- `scripts/repair_docx_openxml_compat.py` is a bounded DOCX compatibility helper for recovery passes that have schema-sensitive namespace, property-order, or WPS-renumbered built-in styleId drift. It may rewrite only selected XML parts, preserve the package, restore WordprocessingML `pPr`/`rPr`/`style`/`tblPr`/`tcPr` child ordering, including table and cell border child placement, add missing namespace declarations required by `mc:Ignorable`, and canonicalize built-in style ids such as `1 -> Normal`, `6 -> TOC1`, and `7 -> TOC2`; it must not edit visible thesis text, media, comments, relationships, citations, headers, footers, or field content.
- `scripts/repair_docx_submission_blockers.py` is a bounded package-preserving repair helper for final thesis submission blockers that are document-structural but not media/comment/citation changes. It may patch only `word/document.xml` to convert visible square placeholders to spaces, add missing front-matter TOC entries, add a standard live TOC field wrapper around the existing visible TOC cache, and add one small academic results table when a final manuscript contains no table at all; it must report source/final SHA256 values and must not touch media, relationships, comments, citations, styles, headers, or footers.
- `scripts/docx_formula_number_table.py` is the bounded canonical formula-object conversion and formula-numbering helper. It may wrap an existing OMML formula with a right-side number, or replace exact text formula anchors from a JSON formula map with OMML formula objects inside a borderless numbering table or paragraph tab layout; its DOCX package mutation scope is `word/document.xml`. Paragraph tab stops and borderless table widths must be computed from the active section's usable page width, not hard-coded beyond the right page margin.
- `scripts/repair_docx_review_artifact_anchors.py` is the bounded canonical review-artifact anchor restorer. It may only restore missing comment range/reference markers by stable source paragraph id and picture-anchor scope in `word/document.xml`; it must not edit comment text, mark comments done, accept revisions, or rewrite surrounding prose.
- retired thick thesis rewrite helpers are archived under `references/archive/legacy-thesis-scripts/20260429-frontmatter-drift/`; `scripts/align_target_thesis.py`, `scripts/build_pass4_docx.py`, and `scripts/rebuild_complete_sample.py` must not return to the active script tree or production helper list.
- active `scripts/` are scanned for broad thesis rewrite patterns such as fixed paragraph indexes, `paragraph.clear()`, Word `Range.Style =`, broad `word/document.xml` rewrites, and figure drawing plus DOCX insertion without a manifest route; any allowed low-level DOCX helper must be named in the active-script policy allowlist with a bounded surface owner.
- front-matter and end-matter surface matrices are validator-owned evidence records: the final gate must validate required rows and per-surface evidence rather than accepting those matrix fields as path-only placeholders.
- `scripts/selftest_skill_flow.py` is the canonical regression suite for workflow failures. Its default `--suite fast` is a bounded lightweight suite for text gates, validator coverage, small DOCX fixtures, stale evidence rejection, validation-command `--gate-record` enforcement, DOCX/font audit SHA/WPS named-size evidence regressions, reference-entry leading-number superscript regressions, front-matter inline abstract page-number repair, strict structural-figure readable-width/height regressions, source-to-final preservation regressions for deleted comments, stripped tracked changes, lost citation superscripts, stale citation-audit content, same-count stale citation records, non-main story part revisions, stripped extra `comments*.xml` parts, missing generator `--source-docx`, strict figure-manifest source/final DOCX binding regressions, VML/WPS pict extent regressions, rendered screenshot/result blank and purple-placeholder regressions, whole-package media relationship replacement/removal/addition regressions, drawing-object extent/adjacency/index-shift transaction regressions, reference-residue citation-preservation regressions, acceptance-generator source-docx and comment-resolution propagation, sample self-check manifest-source binding, body-style audit coverage for narrative figure-reference paragraphs, cover value-line false-positive exclusion, image-dimension disabled-wording detection, abstract/keyword donor-pollution detection, and transaction-level figure-manifest false-pass regressions; `--suite integration` and `--suite all` run inventory parity plus one complete `run_integration_gate.py --case all` path. Positive fixtures must use real evidence extensions such as `.png` page images and `.pdf` render evidence instead of txt/md/html stand-ins.
- `scripts/validate_skill_gate_record_gate.py::validate_skill_invocation_lock` owns explicit skill-invocation fail-closed validation. It must reject missing anti-bypass lock fields, reference-only execution, narrow/smoke gate substitutes, mutation before lock, non-canonical final gate paths, and blocked evidence hidden as caveats or residual risk.
- `scripts/validate_skill_gate.cmd` is a thin Windows launcher and must avoid user-bound Python paths in favor of `%LocalAppData%` discovery plus non-stub fallback
- duplicate registry keys, router range drift, helper-script target-path guessing, and protected-surface ownership overlap are validator-level failures, not only prose-rule violations
- `references/rule-owner-map.json` is the canonical owner and coverage manifest for durable rule IDs, legacy aliases, owning files, router exposure, required keywords, and validator/selftest ownership metadata
- if a new workflow rule is meant to be hard and cross-session, it is incomplete until either the validator or the selftest suite can fail on it
- if a skill edit changes rules or workflow, it is incomplete until `SKILL.md`, the canonical owner reference, `references/rule-owner-map.json`, affected templates, validators/generators/selftests, and state records are either updated or explicitly recorded as not applicable

### Informational but current

- `DURABLE-RULE-PROMOTION-AUDIT.md`
- `COMPATIBILITY-EXPORT-RETIREMENT.md`
- `COMPATIBILITY-EXPORT-EXTERNAL-AUDIT.md`
- `COMPATIBILITY-EXPORT-EXTERNAL-AUDIT-RECORD.md`
- `INSTALL.md`
- `PUBLISHING.md`
- `README-WORKBUDDY-ZH.md`
- `BASELINE-20260327.md`
- `MIRROR.md`
- `RULE-CONFLICT-REVIEW.md`
- `INCIDENT-20260419-THESIS-LANE-FREEZE.md`
- `INCIDENT-20260419-CONFLICT-MATRIX.md`
- `references/cases/thesis-docx-integration-regressions-20260419.md`
- `references/cases/skill-maintenance-20260509-transaction-figure-manifest.md`
- `references/cases/skill-maintenance-20260509-strict-figure-manifest-binding.md`
- `references/cases/skill-maintenance-20260510-comment-figure-transaction-fail-closed.md`
- `references/cases/skill-maintenance-20260511-pass037-detector-closure.md`
- `references/cases/skill-maintenance-20260511-pass038-frontmatter-toc-structure-helper.md`
- `references/cases/skill-maintenance-20260511-runtime-widget-capture-rejection.md`
- `references/cases/skill-maintenance-20260511-body-style-field-instr-text.md`
- `references/cases/skill-maintenance-20260520-orthogonal-connector-law.md`
- `references/cases/skill-maintenance-20260522-mechanical-cad-package-parity.md`
- `references/cases/skill-maintenance-20260523-citation-hyperlink-occurrence-gate.md`

## 2. Backup Files

These files are old preserved copies. They are not the active source of truth.

- `references/archive/backups/SKILL.md.backup-20260410-final-mojibake-fix`
- `references/archive/backups/SKILL.md.backup-20260410-utf8bomfix`
- `references/archive/backups/SKILL.md.bak-20260323-221556`
- `references/archive/backups/SKILL.md.bak-clean-20260325-231050`
- `references/archive/backups/SKILL.md.bak-merge-20260409-1919`
- `references/archive/backups/memory.md.bak-clean-20260325-231050`

Rule:

- never edit a backup file as part of normal active maintenance unless the task is explicitly archival cleanup
- never treat a backup file as the live rule source when a current active file already exists

## 3. Archive Files

These files are historical snapshots for reference only.

- `references/archive/2026-03-27-skill-snapshot.md`

Rule:

- do not load archive files by default during execution
- use them only when tracing history, comparing rule drift, or recovering deleted context

## 4. Historical Note Files

These files record transitional or pre-clean states. They are not the active source of truth.

- `memory.pre-clean-20260328.md`

Rule:

- keep for historical audit or migration reference
- do not use as the active memory baseline when `memory.md` already exists

## 5. Visual Sample Assets

These files are supporting visual references, not rule texts.

- `references/visual-style-samples/**`

Rule:

- use them only when the current task needs figure, formula, table, or TOC visual imitation, or when a mechanical CAD task needs the stored fallback baseline
- for future mechanical CAD package tasks, `references/visual-style-samples/mechanical-cad/` is the active fallback baseline set only when the current run has no stronger user-provided CAD/PDF package sample

## 6. Execution Rule

- If a file appears in both an active path and a backup/archive path, always use the active path.
- If a user asks to update the skill, update active files first unless they explicitly ask for archival cleanup.
- If a future cleanup removes backups or archives, update this index in the same turn.
