# Skill Maintenance Case: SGB620 Full Redo Hardening

Date: 2026-05-30

## Trigger

The user required the SGB620/80T scraper conveyor thesis and drawings to be rebuilt, not patched. Repeated defects included CAD frame overflow, content overlap after scaling, linework that did not distinguish thick/thin/center/hidden/hatch/dimension/leader/table families, bibliography label regressions away from compact `[1]content`, formula-density false passes, and final theses using schematic or stale CAD images instead of the revised drawing package.

## Rule Change

- `CORE-FIGURE-010` now explicitly couples frame-overflow and content-overlap gates and requires the current accepted CAD package or stronger user sample to be locked as the run baseline.
- `audit_mechanical_drawing_package.py` now fails closed when `rendered_review_verdict.passed=false`, even if a caller passes `--no-require-rendered-review`, whenever strict CAD structure or reference CAD samples are in scope. It also enforces an effective sample-based DWG byte-density floor of `0.35` and adds a rendered ink contrast/readability audit for near-white CAD line colors.
- `audit_cad_dxf_linework_fidelity.py` now checks layer/entity color and true-color values in addition to lineweight, linetype, and source line family coverage. Low-contrast or light CAD colors that will disappear on white-background plotting are rejected at source level before rendered PNG/PDF review.
- Mechanical drawing manifests for formal graduation projects must list the required sheet set and cannot pass by merging, omitting, or renaming old simplified sheets.
- Mechanical CAD PDF renders must preserve true A-series page boxes; compact preview-size PDFs are blocked even when the visible drawing looks similar. The audit must expose estimated PDF sheet workload and keep vector drawing-object evidence rather than replacing the handoff with a raster-only PDF.
- `CORE-FIGURE-012` now explicitly blocks mechanical CAD body/appendix figure prose from accidentally triggering the structural draw.io/SVG figure contract. If ordinary CAD renders use words such as structure, relationship, workflow, diagram, draw.io, or svg around the figure reference, `validate_figure_manifest` must fail unless the manifest has a matching structural diagram entry. Mechanical CAD figure prose should instead use drawing terms such as dimensions, installation position, force point, section, assembly position, and drawing number.
- `FB-CITE-048`, `FMT-FORMULA-003`, and `FMT-FORMULA-008` remain the owning gates for compact `[n]content` bibliography labels, 200+ body OMML formulas, raw-token rejection, and formula-only page rejection.
- Citation punctuation regressions must be caught before handoff: body text must use `claim[n]。`, not `claim。[n]`; canonical content validation, citation normalization, and final citation audit must all reject or repair the latter form.
- Body mixed-script audits must not count pure numeric citation markers such as `[12]` as ordinary Western body text after the citation has already become a superscript hyperlink; otherwise valid Chinese body paragraphs with only citation-number ASCII become false failures.
- Repeated page-flow complaints are handled as whole-document pagination evidence, not only local page screenshots: reference block, appendix, acknowledgement, header/footer, TOC, and formula-number page flow must be regenerated after the last DOCX mutation.
- Figure image-holder paragraphs are audited against their real template paragraph metrics, not only against image visibility. The canonical builder now writes the required direct holder metrics during generation and postprocess: dedicated `ThesisImageHolder` style, `before=120`, `after=0`, `line=360`, `lineRule=auto`, `firstLine=0`, `left=0`, `right=0`, `firstLineChars=0`, `leftChars=0`, and `rightChars=0`, with hanging leftovers cleared.
- Static TOC entries must replay the donor's run-level and paragraph-level baseline after any field/page refresh. For this SGB620 template the canonical visible TOC baseline is exact 20 pt line spacing, level indents `0/420/840`, visible run fonts `宋体` plus `Times New Roman`, size `24` half-points, and explicit `w:b w:val="0"` rather than missing bold tags or theme fonts.
- End-matter paragraphs must not be treated as ordinary body text. The acknowledgement body is generated from the locked acknowledgement/reference tail-block donor and no longer receives a forced two-character first-line indent when the template baseline has no such direct indent.
- Compatibility tools that open or validate DOCX may rewrite package internals and invalidate SHA-bound acceptance evidence. Final skill-gate evidence must be captured after the last such tool invocation, or the canonical builder must be rerun and the gate rerun before handoff.

## Validation

This case records the rule consolidation for the active SGB620/80T run. The current project still needs exact-output CAD, DOCX, PDF, and final package audits before it can be accepted.
