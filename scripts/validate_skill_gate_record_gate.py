"""Gate-record checks for validate_skill_gate."""

from __future__ import annotations

__all__ = ["check_gate_record"]

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from .validate_skill_gate_registry_core import (
        EXPLICIT_VALUES,
        FINAL_ACCEPTANCE_SCHEMA,
        IMAGE_EXTENSIONS,
        PLACEHOLDER_VALUES,
        PDF_EXTENSIONS,
    )
    from .validate_skill_gate_registry_records import (
        EVIDENCE_FIELD_TO_TYPE,
        FORMULA_HINT_TOKENS,
        NON_CONTENT_THESIS_HINT_TOKENS,
        NUMBERING_PARAGRAPH_TOKENS,
        PATH_FIELD_PREFIXES,
        REQUIRED_NON_NONE_PATHS_BY_MODE,
    )
    from .validate_skill_gate_record_core import (
        check_thesis_citation_audit_report,
        check_docx_citation_anchor_pollution_audit_report,
        check_docx_body_style_audit_report,
        check_docx_font_audit_report,
        check_docx_font_color_audit_report,
        check_docx_list_pollution_audit_report,
        check_docx_whole_format_gate_report,
        check_humanizer_evidence_record,
        format_repair_task_touches_surface,
        sha256_file,
        validate_project_local_helper_preflight_fields,
        validate_template_lock_fields,
        validate_gate_single_prefix,
    )
    from .validate_skill_gate_record_evidence import (
        check_effective_font_evidence_record,
        check_required_surface_evidence_record,
        check_review_evidence_record,
    )
    from .validate_skill_gate_record_format import check_format_repair_task_record
    from .thesis_figure_contract import (
        docx_figure_surface_summary,
        drawio_structural_geometry_issues,
        final_docx_manifest_requirement_issues,
        final_docx_figure_surface_issues,
        validate_structural_geometry_report,
        validate_figure_manifest,
    )
    from .validate_thesis_mutation_transaction import (
        load_record as load_transaction_record,
        resolve_path as resolve_transaction_path,
        validate_transaction_record,
    )
    from .audit_docx_review_artifacts import (
        APPROVED_NONPRESERVATION_CITATION_SCOPE,
        STRICT_CITATION_PRESERVATION_SCOPE,
        WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE,
        validate_citation_run_reports,
        validate_review_artifact_reports,
    )
    from .audit_docx_cad_appendix_binding import audit_docx_cad_appendix_binding
    from .audit_docx_formula_objects import audit_docx as audit_formula_objects
    from .audit_thesis_citations import audit_docx as audit_body_citations
    from .audit_thesis_comment_resolution import (
        collect_comment_snapshot,
        has_all_comments_claim,
        validate_comment_resolution_ledger,
    )
    from .validate_skill_gate_utils import (
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )
except ImportError:
    from validate_skill_gate_registry_core import (
        EXPLICIT_VALUES,
        FINAL_ACCEPTANCE_SCHEMA,
        IMAGE_EXTENSIONS,
        PLACEHOLDER_VALUES,
        PDF_EXTENSIONS,
    )
    from validate_skill_gate_registry_records import (
        EVIDENCE_FIELD_TO_TYPE,
        FORMULA_HINT_TOKENS,
        NON_CONTENT_THESIS_HINT_TOKENS,
        NUMBERING_PARAGRAPH_TOKENS,
        PATH_FIELD_PREFIXES,
        REQUIRED_NON_NONE_PATHS_BY_MODE,
    )
    from validate_skill_gate_record_core import (
        check_thesis_citation_audit_report,
        check_docx_citation_anchor_pollution_audit_report,
        check_docx_body_style_audit_report,
        check_docx_font_audit_report,
        check_docx_font_color_audit_report,
        check_docx_list_pollution_audit_report,
        check_docx_whole_format_gate_report,
        check_humanizer_evidence_record,
        format_repair_task_touches_surface,
        sha256_file,
        validate_project_local_helper_preflight_fields,
        validate_template_lock_fields,
        validate_gate_single_prefix,
    )
    from validate_skill_gate_record_evidence import (
        check_effective_font_evidence_record,
        check_required_surface_evidence_record,
        check_review_evidence_record,
    )
    from validate_skill_gate_record_format import check_format_repair_task_record
    from thesis_figure_contract import (
        docx_figure_surface_summary,
        drawio_structural_geometry_issues,
        final_docx_manifest_requirement_issues,
        final_docx_figure_surface_issues,
        validate_structural_geometry_report,
        validate_figure_manifest,
    )
    from validate_thesis_mutation_transaction import (
        load_record as load_transaction_record,
        resolve_path as resolve_transaction_path,
        validate_transaction_record,
    )
    from audit_docx_review_artifacts import (
        APPROVED_NONPRESERVATION_CITATION_SCOPE,
        STRICT_CITATION_PRESERVATION_SCOPE,
        WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE,
        validate_citation_run_reports,
        validate_review_artifact_reports,
    )
    from audit_docx_cad_appendix_binding import audit_docx_cad_appendix_binding
    from audit_docx_formula_objects import audit_docx as audit_formula_objects
    from audit_thesis_citations import audit_docx as audit_body_citations
    from audit_thesis_comment_resolution import (
        collect_comment_snapshot,
        has_all_comments_claim,
        validate_comment_resolution_ledger,
    )
    from validate_skill_gate_utils import (
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )


SKILL_ROOT = Path(__file__).resolve().parents[1]
COMMENTS_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
THESIS_MODES = {"thesis-only", "format-repair-only", "program-plus-thesis"}
THESIS_ONLY_SOURCE_ROLE_PREFIXES = {
    "- comment-resolution source DOCX path:",
    "- comment-resolution source DOCX SHA256:",
    "- figure source DOCX path:",
    "- figure source DOCX SHA256:",
}
THESIS_WORKFLOWS = {
    "new-thesis-production",
    "whole-thesis-revision",
    "local-surface-repair",
    "content-only-paragraph-revision",
    "audit-only",
}
WHOLE_REBUILD_CITATION_SCOPE_WORKFLOWS = {"new-thesis-production", "whole-thesis-revision"}
STRICT_CITATION_SCOPE_WORKFLOWS = {"local-surface-repair", "content-only-paragraph-revision"}
MECHANICAL_CAD_ACCEPTANCE_SCHEMA_VALUE = "graduation-project-builder.mechanical-cad-acceptance.v2"
MECHANICAL_CAD_AUDIT_SCHEMA_VALUES = {
    "graduation-project-builder.mechanical-drawing-package-audit.v4",
    "graduation-project-builder.mechanical-drawing-package-audit.v5",
    "graduation-project-builder.mechanical-drawing-package-audit.v6",
}
MECHANICAL_CAD_ACCEPTANCE_HEADINGS = (
    "# Mechanical CAD Acceptance Template",
    "## Outputs",
    "## Audit Evidence",
    "## Rendered Review",
    "## Scope",
    "## Validation",
)
MECHANICAL_CAD_ACCEPTANCE_PREFIXES = (
    "- acceptance schema:",
    "- task mode:",
    "- subtask:",
    "- exact final delivery package path:",
    "- exact final delivery package sha256:",
    "- exact audited CAD package path:",
    "- exact audited CAD package sha256:",
    "- exact DWG package path:",
    "- exact DWG package sha256:",
    "- exact DXF package path:",
    "- exact DXF package sha256:",
    "- exact combined PDF path:",
    "- exact combined PDF sha256:",
    "- exact PNG render package path:",
    "- exact PNG render package sha256:",
    "- drawing regeneration manifest path:",
    "- drawing regeneration manifest sha256:",
    "- current package SHA binding verdict:",
    "- mechanical drawing package audit path:",
    "- mechanical drawing package audit verdict:",
    "- mechanical drawing formal CAD source provenance verdict:",
    "- mechanical drawing schematic/concept substitute rejection verdict:",
    "- CAD official command route verdict:",
    "- CAD official command test log:",
    "- CAD executable or COM ProgID evidence:",
    "- non-CAD fallback rejection verdict:",
    "- external CAD case reference URL:",
    "- reference-use restriction:",
    "- no verbatim geometry copying verdict:",
    "- mechanical drawing reference baseline path:",
    "- mechanical drawing CAD open-view close-up evidence path:",
    "- mechanical drawing CAD open-view structural coherence verdict:",
    "- mechanical drawing complete assembly/object recognizability verdict:",
    "- mechanical drawing scattered-parts rejection verdict:",
    "- mechanical drawing reference-view trace alignment verdict:",
    "- mechanical drawing annotation/leader/title-block completeness verdict:",
    "- mechanical drawing structure-first redraw workflow verdict:",
    "- mechanical drawing dimension/leader/title-block second-pass verdict:",
    "- mechanical drawing loose-part collage rejection verdict:",
    "- mechanical drawing external-case annotation checklist path:",
    "- mechanical drawing A0 process flow verdict:",
    "- mechanical drawing A0 drive chain verdict:",
    "- mechanical drawing A0 support load path verdict:",
    "- mechanical drawing A0 internal working element verdict:",
    "- mechanical drawing A0 design intent readability verdict:",
    "- mechanical drawing A0 overall assembly layout baseline verdict:",
    "- mechanical drawing A0 upper full-length view count:",
    "- mechanical drawing A0 lower full-length view count:",
    "- mechanical drawing A0 BOM right-zone verdict:",
    "- mechanical drawing A0 balloon-to-BOM row match verdict:",
    "- mechanical drawing source linework delta audit path:",
    "- mechanical drawing source linework delta verdict:",
    "- mechanical drawing source CAD package sha256:",
    "- mechanical drawing final CAD package sha256:",
    "- mechanical drawing changed source sheet count:",
    "- mechanical drawing identical source sheet count:",
    "- mechanical drawing changed source entity count:",
    "- mechanical drawing linework family delta count:",
    "- mechanical drawing old-like large circle overlay count:",
    "- mechanical drawing PDF-only change rejection verdict:",
    "- mechanical drawing minor-entity-move-only rejection verdict:",
    "- mechanical drawing source-to-PDF derivation verdict:",
    "- mechanical drawing rendered review evidence paths:",
    "- mechanical drawing rendered no-overlap verdict:",
    "- mechanical drawing boundary clearance verdict:",
    "- mechanical drawing detail density verdict:",
    "- mechanical drawing title block/table/notes isolation verdict:",
    "- mechanical drawing title-block cell containment verdict:",
    "- mechanical drawing title-block short-line topology audit verdict:",
    "- mechanical drawing missing short table line count:",
    "- mechanical drawing broken cell-border count:",
    "- mechanical drawing table-grid topology mismatch count:",
    "- mechanical drawing diagnostic-overlay-free title-block crop path:",
    "- mechanical drawing annotation margin clearance verdict:",
    "- mechanical drawing local crowding verdict:",
    "- mechanical drawing text/table/frame overlap verdict:",
    "- mechanical render readability audit path:",
    "- mechanical render readability verdict:",
    "- mechanical render readability source lineweight verdict:",
    "- mechanical render readability render-only stroke cap:",
    "- mechanical render readability render-only lineweight scaling:",
    "- text text overlap count:",
    "- text graphic cover count:",
    "- severe line crowding count:",
    "- mechanical render readability small-tile threshold:",
    "- mechanical render readability failed sheet list:",
    "- mechanical render readability requires crop review count:",
    "- mechanical render readability line bundle blocker count:",
    "- mechanical render readability diagnostic-overlay-free crop review path:",
    "- mechanical render readability standard-sheet view occupancy verdict:",
    "- mechanical render readability final PDF/PNG SHA evidence:",
    "- mechanical drawing user-reported text-cover crop audit path:",
    "- mechanical drawing diagnostic-overlay-free after-crop paths:",
    "- mechanical drawing reported crop blocker count:",
    "- mechanical drawing text exclusion halo audit path:",
    "- mechanical drawing text exclusion halo audit verdict:",
    "- mechanical drawing text exclusion halo violation count:",
    "- mechanical drawing diagonal hatch/section/flow text-cover count:",
    "- mechanical drawing machine overlap audit verdict:",
    "- mechanical drawing content-overlap audit verdict:",
    "- mechanical drawing content-overlap evidence path:",
    "- mechanical drawing registered content bbox count:",
    "- mechanical drawing checked content pair count:",
    "- mechanical drawing content overlap count:",
    "- mechanical drawing view-view overlap count:",
    "- mechanical drawing detail-frame-main-view overlap count:",
    "- mechanical drawing table text/grid collision count:",
    "- mechanical drawing table_text_grid_collision_count:",
    "- mechanical drawing dimension-line view/table crossing count:",
    "- mechanical drawing leader-line view/table crossing count:",
    "- mechanical drawing balloon geometry collision count:",
    "- mechanical drawing bbox helper envelope escape count:",
    "- mechanical drawing stale rendered preview count:",
    "- mechanical drawing outside-frame ink audit verdict:",
    "- mechanical drawing outside-frame ink audit evidence path:",
    "- mechanical drawing outside-frame independent ink component count:",
    "- mechanical drawing outside-frame text component count:",
    "- mechanical drawing outside-frame leader component count:",
    "- mechanical drawing outside-frame hatch/section component count:",
    "- mechanical drawing outside-frame table/title-block component count:",
    "- mechanical drawing inner-frame safe-margin audit verdict:",
    "- mechanical drawing inner-frame safe-margin evidence path:",
    "- mechanical drawing right safe-boundary intrusion count:",
    "- mechanical drawing leader-text inner-frame intrusion count:",
    "- mechanical drawing view-geometry inner-frame intrusion count:",
    "- mechanical drawing dimension-text inner-frame intrusion count:",
    "- mechanical drawing annotation ownership audit verdict:",
    "- mechanical drawing unowned free text count:",
    "- mechanical drawing unsupported floating text count:",
    "- mechanical drawing unbound scattered text count:",
    "- mechanical drawing dimension-like text without anchor count:",
    "- mechanical drawing orphan text examples path:",
    "- mechanical drawing owner-zone coverage verdict:",
    "- mechanical drawing user-reported crop binding review path:",
    "- mechanical drawing user-reported crop binding review verdict:",
    "- mechanical drawing min cell padding mm:",
    "- mechanical drawing cell padding violation count:",
    "- CAD text style/font audit path:",
    "- CAD text style/font audit verdict:",
    "- normal CAD text entity verdict:",
    "- unsupported CAD font/style count:",
    "- artistic/vectorized text count:",
    "- mechanical drawing reserved-zone intrusion audit verdict:",
    "- mechanical drawing protected-table-zone intrusion audit verdict:",
    "- mechanical drawing reserved-zone intrusion count:",
    "- mechanical drawing protected-table-zone intrusion count:",
    "- mechanical drawing dimension-line table-zone intrusion count:",
    "- mechanical drawing dimension-text table-zone intrusion count:",
    "- mechanical drawing view-geometry table-zone intrusion count:",
    "- mechanical drawing detail-view table-zone intrusion count:",
    "- mechanical drawing leader/balloon table-zone intrusion count:",
    "- mechanical drawing dimension table-zone intrusion count:",
    "- mechanical drawing title-block/BOM protected-zone intrusion count:",
    "- mechanical drawing view-geometry reserved-zone intrusion count:",
    "- mechanical drawing hatch/section fill clipping audit verdict:",
    "- mechanical drawing hatch/section fill clipping evidence path:",
    "- mechanical drawing hatch/section fill boundary violation count:",
    "- mechanical drawing hatch/section fill adjacent-view crossing count:",
    "- mechanical drawing hatch/section fill dimension-line crossing count:",
    "- mechanical drawing hatch/section fill title-block/table/BOM/frame crossing count:",
    "- mechanical drawing hatch/section fill blank-background leak count:",
    "- mechanical drawing text entity overlap count:",
    "- mechanical drawing text legibility machine audit verdict:",
    "- mechanical drawing minimum CAD text height mm:",
    "- mechanical drawing minimum rendered text height px:",
    "- mechanical drawing text integrity audit path:",
    "- mechanical drawing text integrity audit verdict:",
    "- mechanical drawing mojibake/tofu/missing glyph count:",
    "- mechanical drawing missing required drawing text count:",
    "- mechanical drawing text orientation audit path:",
    "- mechanical drawing text orientation audit verdict:",
    "- mechanical drawing upside-down text count:",
    "- mechanical drawing mirrored text count:",
    "- mechanical drawing lineweight/linetype fidelity audit path:",
    "- mechanical drawing lineweight/linetype fidelity audit verdict:",
    "- mechanical drawing lineweight/linetype fidelity package sha256:",
    "- source thick lineweight required:",
    "- source thick lineweight observed:",
    "- source thick lineweight mismatch count:",
    "- source thin lineweight required:",
    "- source thin lineweight observed:",
    "- source thin lineweight mismatch count:",
    "- mechanical drawing color-family audit path:",
    "- mechanical drawing color-family audit verdict:",
    "- mechanical drawing color-family package/path sha256:",
    "- mechanical drawing thin-solid-white-only verdict:",
    "- mechanical drawing non-thin white layer/entity count:",
    "- mechanical drawing expected color family count:",
    "- mechanical drawing entity color override count:",
    "- mechanical drawing source line family coverage verdict:",
    "- mechanical drawing PDF page-box sheet-size audit verdict:",
    "- mechanical drawing rendered ink contrast audit verdict:",
    "- mechanical drawing worst readable ink ratio:",
    "- mechanical drawing minimum readable ink ratio:",
    "- mechanical drawing effective DWG byte-density ratio:",
    "- mechanical drawing estimated PDF sheet workload:",
    "- mechanical drawing A-series page box count:",
    "- mechanical drawing thick solid family count:",
    "- mechanical drawing thin solid family count:",
    "- mechanical drawing center dash-dot family count:",
    "- mechanical drawing hidden dashed family count:",
    "- mechanical drawing section hatch family count:",
    "- mechanical drawing manufacturing complexity audit verdict:",
    "- mechanical drawing minimum manufacturing detail family count:",
    "- mechanical drawing entity-count-only false-pass verdict:",
    "- thesis DOCX mutation verdict:",
    "- validation command:",
    "- validation result:",
)
MECHANICAL_CAD_PASS_VALUES = {"pass", "passed", "yes", "true", "ok"}
MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM = 3.8
MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX = 40.0
MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES = 8
MECHANICAL_CAD_NO_OVERLAP_VALUES = MECHANICAL_CAD_PASS_VALUES | {
    "no-overlap",
    "no_overlap",
    "no overlap",
    "clear",
}
MECHANICAL_CAD_ENTITY_ONLY_REJECT_VALUES = {
    "not-used",
    "not_used",
    "not-used-as-acceptance",
    "not_used_as_acceptance",
    "not-entity-only",
    "not_entity_only",
    "rejected",
    "visual-reviewed",
    "visual_reviewed",
    "blocked",
}
MECHANICAL_FORMULA_MIN_COUNT = 200
MECHANICAL_FORMULA_CONTEXT_TOKENS = {
    "mechanical",
    "mechanism",
    "structural design",
    "crane",
    "gantry",
    "cantilever",
    "reducer",
    "conveyor",
    "cad",
    "drawing package",
    "mechanical drawing",
    "机械",
    "机构",
    "结构设计",
    "起重机",
    "门式",
    "悬臂",
    "减速器",
    "输送机",
    "图纸",
    "装配图",
}
MECHANICAL_FORMULA_EXCLUSION_TOKENS = {
    "formula explicitly excluded",
    "formula not in scope",
    "calculation depth explicitly excluded",
    "no formula work requested",
    "明确排除公式",
    "不处理公式",
    "不要求公式",
}
CAD_APPENDIX_FIELD_PREFIXES = (
    "- CAD appendix binding audit path:",
    "- CAD appendix binding audit verdict:",
    "- CAD appendix matched sheet count:",
    "- CAD appendix missing sheet count:",
    "- CAD appendix final DOCX SHA256:",
    "- CAD appendix final package path:",
)
FULL_SCOPE_TOKENS = (
    "whole thesis",
    "whole-thesis",
    "full thesis",
    "full-thesis",
    "full paper",
    "full-paper",
    "1:1",
    "template alignment",
    "template-aligned",
    "submission-ready",
    "ready to submit",
    "whole document",
    "complete thesis",
    "complete paper",
    "all sections",
    "template parity",
    "format parity",
    "teacher template",
    "approved sample",
    "submit ready",
)
CANONICAL_ROLE_LANES = ("controller", "content-worker", "format-worker", "figure-worker", "citation-worker", "program-worker", "acceptance-worker", "audit")
CANONICAL_ROLE_ALIASES = ("总控", "内容", "格式", "图表", "引用", "程序", "验收", "审核")

SKILL_INVOCATION_LOCK_REQUIRED_PREFIXES = (
    "- skill name:",
    "- user invocation source:",
    "- invocation detected:",
    "- lock created before mutation?:",
    "- run start order verdict:",
    "- task mode:",
    "- subtask:",
    "- project root:",
    "- requested mutation?:",
    "- thesis/docx surface touched?:",
    "- loaded entrypoint:",
    "- loaded routed references:",
    "- active checklist path:",
    "- agent run manifest path:",
    "- lane task card paths:",
    "- project-local helper preflight report path:",
    "- project-local helper risk count:",
    "- project-local helper disposition:",
    "- mutation transaction record path:",
    "- mutation allowed verdict:",
    "- blocked reason:",
    "- exact output path:",
    "- exact output sha256:",
    "- final gate record path:",
    "- final gate command:",
    "- final gate verdict:",
    "- explicit invocation source type:",
    "- skill activation status:",
    "- rule engine takeover verdict:",
    "- prohibited bypasses checked:",
    "- canonical gate required?:",
    "- narrow/smoke gate substitute used?:",
    "- failed evidence escalation verdict:",
    "- no project-local thick helper execution before preflight?:",
    "- no non-control action before lock?:",
    "- no mutation before lock?:",
    "- final handoff allowed verdict:",
    "- blocked evidence disposition:",
)


def _contains_blocking_status_token(value: str) -> bool:
    lowered = (value or "").lower()
    for phrase in ("not checked", "not-applicable", "not applicable"):
        if phrase in lowered:
            return True
    return re.search(r"\b(fail|failed|missing|pending|blocked|stale)\b", lowered) is not None

FULL_SCOPE_CLAIMS = {
    "full-thesis-template-alignment",
    "full-paper-template-alignment",
    "whole-thesis-template-alignment",
    "submission-ready",
}
SAMPLE_SELF_CHECK_BLOCK_TOKENS = (
    "smoke acceptance mode: yes",
    "smoke-only; blocked for delivery",
    "full thesis content gate failed",
    "deliverable critical gate: blocked",
)
SMOKE_SUMMARY_FINAL_ACCEPTANCE_TOKENS = (
    "officecli validate",
    "pdf export",
    "rendered pdf",
    "media count",
    "image count",
    "page-image",
    "page images",
    "old terms count",
    "old crawler terms count",
    "required phrase checks",
    "phrase presence",
    "screenshot existence",
)
PASS_SHAPED_HANDOFF_TOKENS = (
    "handoff status: pass",
    "known caveats: none",
    "validation result: pass",
    "officecli validate: pass",
    "officecli validate passed",
)
REQUIRED_SAMPLE_SELF_CHECK_DETECTORS = (
    "header.presence-contract",
    "header-footer.page-number-template-contract",
    "figure.scope-manifest-contract",
    "figure.family-style-contract",
    "figure.image-dimension-contract",
    "cover.identity-value-line-contract",
    "abstract.template-style-contract",
    "heading.baseline-contract",
    "toc.visible-format-contract",
    "body.style-contamination-contract",
    "endmatter.indentation-contract",
    "tail-block.pagination-contract",
    "chapter.format-preservation-contract",
    "common.pre-submission-checklist",
    "toc.page-number-column-right-edge",
)
PAGE_CLASS_ALIASES = {
    "cover": ("cover", "front cover"),
    "title/front matter": ("title page", "title/front matter", "front matter", "cover_zh_title", "cover_en_title", "title page"),
    "chinese abstract": ("chinese abstract", "zh_abstract", "zh abstract"),
    "english abstract": ("english abstract", "en_abstract", "en abstract"),
    "toc": ("toc", "table of contents"),
    "body/chapter": ("body", "chapter", "first body"),
    "figure": ("figure", "diagram"),
    "table": ("table",),
    "references": ("references", "bibliography"),
    "acknowledgement": ("acknowledgement", "acknowledgment", "ack"),
    "appendix": ("appendix", "appendices"),
}
PAGE_CLASS_PASS_TOKENS = {"pass", "passed"}
PAGE_CLASS_BAD_TOKENS = {
    "fail",
    "failed",
    "missing",
    "not checked",
    "sampled-only",
    "sample only",
    "pending",
    "blocked",
    "unresolved",
}
ABSTRACT_LEDGER_SURFACES = (
    "chinese abstract title",
    "chinese abstract body",
    "chinese keyword line",
    "english abstract title",
    "english abstract body",
    "english keyword line",
)
FRONT_MATTER_FINAL_EVIDENCE_FIELDS = {
    "cover_style": ("- cover style evidence path:", "- cover style verdict:"),
    "declaration_or_title_front_matter": (
        "- declaration/title front matter evidence path:",
        "- declaration/title front matter verdict:",
    ),
    "zh_abstract_title": ("- Chinese abstract title evidence path:", "- Chinese abstract title verdict:"),
    "zh_abstract_body": ("- Chinese abstract body evidence path:", "- Chinese abstract body verdict:"),
    "zh_keyword_line": ("- Chinese keyword line evidence path:", "- Chinese keyword line verdict:"),
    "en_abstract_title": ("- English abstract title evidence path:", "- English abstract title verdict:"),
    "en_abstract_body": ("- English abstract body evidence path:", "- English abstract body verdict:"),
    "en_keyword_line": ("- English keyword line evidence path:", "- English keyword line verdict:"),
    "toc_title": ("- TOC title evidence path:", "- TOC title verdict:"),
    "toc_entries": ("- TOC entries evidence path:", "- TOC entries verdict:"),
    "toc_dotted_leaders": ("- TOC dotted leaders evidence path:", "- TOC dotted leaders verdict:"),
    "toc_page_number_column": (
        "- TOC page-number column evidence path:",
        "- TOC page-number column verdict:",
    ),
    "body_heading_levels": ("- body heading levels evidence path:", "- body heading levels verdict:"),
    "body_text": ("- body text evidence path:", "- body text verdict:"),
    "body_citation_superscripts": (
        "- body citation superscripts evidence path:",
        "- body citation superscripts verdict:",
    ),
    "review_comments_and_change_marks": (
        "- review comments/change marks evidence path:",
        "- review comments/change marks verdict:",
    ),
    "references_title": ("- references title evidence path:", "- references title verdict:"),
    "references_entries": ("- references entries evidence path:", "- references entries verdict:"),
    "acknowledgement_title": ("- acknowledgement title evidence path:", "- acknowledgement title verdict:"),
    "acknowledgement_body": ("- acknowledgement body evidence path:", "- acknowledgement body verdict:"),
    "header": ("- header evidence path:", "- header verdict:"),
    "footer": ("- footer evidence path:", "- footer verdict:"),
    "page_numbers": ("- page numbers evidence path:", "- page numbers verdict:"),
}

ENGLISH_ABSTRACT_INDENTATION_FIELDS = (
    "- English abstract indentation evidence path:",
    "- English abstract indentation verdict:",
)

TABLE_CONTINUATION_EVIDENCE_FIELDS = (
    "- table continuation evidence paths:",
    "- table continuation summary:",
    "- cross-page table rendered pages:",
    "- continuation title outside-grid verdict:",
    "- table row split/header repeat verdict:",
)


def inspect_live_toc_fields(docx_path: Path) -> dict[str, int]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return {"count": 0, "locked_count": 0}
    count = 0
    locked_count = 0
    field_stack: list[dict[str, object]] = []
    for node in root.iter():
        if node.tag == f"{W}fldSimple":
            instr = node.attrib.get(f"{W}instr", "")
            if re.search(r"(^|\s)TOC(\s|$)", instr, re.IGNORECASE):
                count += 1
                if node.attrib.get(f"{W}fldLock", "").lower() == "true":
                    locked_count += 1
            continue
        if node.tag == f"{W}fldChar":
            field_type = node.attrib.get(f"{W}fldCharType", "")
            if field_type == "begin":
                field_stack.append(
                    {
                        "instr": "",
                        "locked": node.attrib.get(f"{W}fldLock", "").lower() == "true",
                    }
                )
            elif field_type == "end" and field_stack:
                field = field_stack.pop()
                instr = str(field.get("instr", ""))
                if re.search(r"(^|\s)TOC(\s|$)", instr, re.IGNORECASE):
                    count += 1
                    if bool(field.get("locked")):
                        locked_count += 1
            continue
        if node.tag == f"{W}instrText" and field_stack:
            field_stack[-1]["instr"] = str(field_stack[-1].get("instr", "")) + (node.text or "")
    return {"count": count, "locked_count": locked_count}


def count_live_toc_fields(docx_path: Path) -> int:
    return inspect_live_toc_fields(docx_path)["count"]
ABSTRACT_REPORT_TOKENS = ("abstract", "keyword", "\u6458\u8981", "\u5173\u952e\u8bcd")
ABSTRACT_INDENT_REPORT_TOKENS = (
    "english abstract indent",
    "abstract indent",
    "firstlinechars",
    "\u82f1\u6587\u6458\u8981\u7f29\u8fdb",
    "\u6458\u8981\u7f29\u8fdb",
    "\u6458\u8981\u7a7a\u683c",
)
REFERENCE_REPORT_TOKENS = ("references", "bibliography", "citation", "\u53c2\u8003\u6587\u732e")
FIGURE_REPORT_TOKENS = (
    "figure",
    "diagram",
    "flowchart",
    "draw.io",
    "drawio",
    "svg",
    "screenshot",
    "runtime screenshot",
    "algorithm result",
    "algorithm-result",
    "yolo",
    "yolov8",
    "dbnet",
    "crnn",
    "ocr",
    "\u56fe",
    "\u622a\u56fe",
    "\u793a\u610f\u56fe",
    "\u68c0\u6d4b\u6548\u679c",
    "\u8bc6\u522b\u7ed3\u679c",
)
FIGURE_WIDTH_USER_REPORTED_TOKENS = (
    "screenshot",
    "system screenshot",
    "figure width",
    "image width",
    "image indent",
    "figure indent",
    "image-holder",
    "image holder",
    "holder indent",
    "too small",
    "not readable",
    "display incomplete",
    "incomplete display",
    "clipped",
    "cropped",
    "truncated",
    "paragraph margin",
    "text width",
    "width/text-width",
    "\u7cfb\u7edf\u622a\u56fe",
    "\u622a\u56fe",
    "\u56fe\u7247",
    "\u56fe\u7247\u7f29\u8fdb",
    "\u56fe\u7247\u6bb5\u843d",
    "\u56fe\u50cf\u7f29\u8fdb",
    "\u56fe\u6bb5\u843d\u7f29\u8fdb",
    "\u5c3a\u5bf8",
    "\u592a\u5c0f",
    "\u770b\u4e0d\u6e05",
    "\u663e\u793a\u4e0d\u5168",
    "\u663e\u793a\u4e0d\u5b8c\u6574",
    "\u88ab\u88c1\u5207",
    "\u88c1\u5207",
    "\u88c1\u526a",
    "\u8de8\u9875",
    "\u6587\u5b57\u6bb5\u843d",
    "\u9f50\u5e73",
    "\u6bb5\u843d\u5bbd\u5ea6",
)
ALGORITHM_CODE_USER_REPORTED_TOKENS = (
    "algorithm code",
    "code visibility",
    "source code",
    "real code",
    "missing code",
    "\u7b97\u6cd5\u4ee3\u7801",
    "\u65b0\u589e\u4ee3\u7801",
    "\u4ee3\u7801\u6ca1\u770b\u89c1",
    "\u771f\u5b9e\u4ee3\u7801",
    "\u6e90\u4ee3\u7801",
)
TOC_DIGITS_CACHE_USER_REPORTED_TOKENS = (
    "toc digits",
    "toc cache",
    "toc number",
    "toc page",
    "bookmark",
    "anchor",
    "1. 1",
    "2. 1",
    "broken toc",
    "\u76ee\u5f55\u6570\u5b57",
    "\u76ee\u5f55\u9875\u7801",
    "\u76ee\u5f55",
    "\u6570\u5b57",
    "\u951a\u70b9",
    "\u4e66\u7b7e",
    "\u6539\u574f",
)
FIGURE_MATERIAL_GENERATED_USER_REPORTED_TOKENS = (
    "material.doc",
    "source material",
    "material source",
    "generated image",
    "generated substitute",
    "ai generated",
    "draw.io",
    "drawio",
    "impostor",
    "wrapped image",
    "\u7d20\u6750.doc",
    "\u7d20\u6750",
    "\u751f\u6210\u56fe",
    "\u5192\u5145",
    "\u4f2a draw.io",
    "\u4e0d\u662fdraw.io",
    "\u4e0d\u662f draw.io",
    "\u6d41\u7a0b\u56fe",
    "\u753b\u7684\u56fe",
)
FIGURE_MATERIAL_ONLY_USER_REPORTED_TOKENS = (
    "material-only",
    "material only",
    "source-only",
    "source only",
    "no redraw",
    "no redrawn",
    "do not redraw",
    "do not generate",
    "supplemental docx",
    "\u53ea\u4ece\u7d20\u6750",
    "\u53ea\u9700\u8981\u4ece\u7d20\u6750",
    "\u4ece\u7d20\u6750.doc\u62c9\u53d6",
    "\u4e0d\u9700\u8981\u81ea\u5df1\u91cd\u65b0\u753b",
    "\u4e0d\u8981\u81ea\u5df1\u91cd\u65b0\u753b",
    "\u4e0d\u5f97\u91cd\u753b",
    "\u4e0d\u5f97\u751f\u6210",
    "\u7f3a\u7684\u56fe",
    "\u8865\u5145 docx",
)
BODY_OPENER_HEADER_TITLE_TOKENS = (
    "title/header",
    "header title",
    "running header",
    "body opener",
    "chapter opener",
    "section opener",
    "opener title",
    "title mismatch",
    "header mismatch",
    "wrong title",
    "wrong header",
    "body title",
    "same physical page",
    "pagination title",
    "\u6807\u9898",
    "\u9875\u7709",
    "\u7ae0\u8282",
    "\u7ae0\u6807\u9898",
    "\u7ae0\u8282\u6807\u9898",
    "\u7bc7\u7ae0",
    "\u7ae0\u9996",
    "\u7ae0\u5f00\u5934",
    "\u6b63\u6587\u6807\u9898",
    "\u9875\u7709\u6807\u9898",
    "\u9875\u811a\u6807\u9898",
    "\u7eea\u8bba",
    "\u7ed3\u8bba",
)
BODY_OPENER_HEADER_TITLE_EVIDENCE_TOKENS = (
    "body opener",
    "running header",
    "expected",
    "observed",
    "physical page",
    "negative evidence",
    "final verdict",
)
USER_REPORTED_VISUAL_SURFACE_TOKENS = (
    "toc",
    "table of contents",
    "abstract",
    "keyword",
    "header",
    "footer",
    "page number",
    "page-number",
    "page_numbers",
    "references",
    "bibliography",
    "body",
    "body font",
    "body style",
    "font",
    "page break",
    "pagination",
    "\u76ee\u5f55",
    "\u6458\u8981",
    "\u5173\u952e\u8bcd",
    "\u9875\u7709",
    "\u9875\u811a",
    "\u9875\u7801",
    "\u53c2\u8003\u6587\u732e",
    "\u6b63\u6587",
    "\u5b57\u4f53",
    "\u5206\u9875",
)
USER_REPORTED_VISUAL_CONTEXT_TOKENS = (
    "screenshot",
    "visible",
    "visual",
    "rendered",
    "looks",
    "wrong",
    "abnormal",
    "font",
    "style",
    "format",
    "pagination",
    "page break",
    "page",
    "drift",
    "\u622a\u56fe",
    "\u53ef\u89c1",
    "\u89c6\u89c9",
    "\u683c\u5f0f",
    "\u5b57\u4f53",
    "\u5206\u9875",
    "\u6362\u9875",
    "\u9875",
    "\u5f02\u5e38",
)
USER_REPORTED_VISUAL_DEFECT_FIELDS = (
    "- user-reported visual defect surfaces:",
    "- user-reported visual defect render-geometry evidence path:",
    "- user-reported visual defect template-vs-target binding verdict:",
    "- user-reported visual defect full-page/key-surface binding verdict:",
)
USER_REPORTED_VISUAL_BAD_TOKENS = {
    "failed",
    "missing",
    "not checked",
    "sample only",
    "sampled-only",
    "stale",
    "xml only",
    "xml-only",
    "structure only",
    "structure-only",
    "officecli only",
    "current draft as baseline",
    "blocked",
}
USER_REPORTED_VISUAL_NOT_APPLICABLE_TOKENS = {
    "not-applicable",
    "not applicable",
}
USER_REPORTED_VISUAL_NOT_APPLICABLE_CONTEXT_TOKENS = {
    "user-reported visual",
    "template-vs-target",
    "full-page",
    "full page",
    "key-surface",
    "key surface",
    "rendered template",
    "rendered target",
    "target/actual/final",
    "surface geometry",
    "surface crop",
    "template rendered region",
    "actual rendered region",
    "toc visual geometry",
    "final verdict",
}
USER_REPORTED_VISUAL_EVIDENCE_TOKENS = (
    "template",
    "rendered",
    "geometry",
    "surface",
    "final verdict",
)
CONTENT_MUTATION_VISUAL_TRIGGER_TOKENS = (
    "content expansion",
    "expand content",
    "body expansion",
    "text mutation",
    "body paragraph insertion",
    "inserted body paragraph",
    "inserted paragraph",
    "body paragraph rewrite",
    "caption pollution",
    "figure caption",
    "figure replacement",
    "image replacement",
    "screenshot paragraph",
    "caption-adjacent body",
    "caption sibling body",
    "content-only-paragraph-revision",
    "visible chinese characters",
    "word count",
    "lengthen",
    "rewrite body",
    "\u6269\u5199",
    "\u5b57\u6570",
    "\u6b63\u6587\u6269\u5199",
    "\u6b63\u6587\u6bb5\u843d",
    "\u63d2\u5165\u6bb5\u843d",
    "\u6dfb\u52a0\u6bb5\u843d",
    "\u56fe\u9898",
    "\u56fe\u540d",
    "\u56fe\u7247",
    "\u622a\u56fe",
    "\u65b0\u63d2\u5165\u6b63\u6587",
    "\u56fe\u9898\u6c61\u67d3",
    "\u56fe\u540d\u6c61\u67d3",
    "\u66ff\u6362\u56fe\u7247",
    "\u63d2\u5165\u56fe\u7247",
    "\u6539\u5199",
)
CONTENT_MUTATION_VISUAL_REQUIRED_FIELDS = (
    "- content mutation rendered-page review path:",
    "- content mutation machine-vision verdict:",
    "- inserted body heading-contamination verdict:",
    "- caption/table sibling body contamination verdict:",
    "- touched-page/blast-radius machine-vision evidence paths:",
    "- format lane post-mutation rendered audit verdict:",
)
CONTENT_MUTATION_VISUAL_BAD_TOKENS = {
    "failed",
    "missing",
    "not checked",
    "sample only",
    "sampled-only",
    "stale",
    "xml only",
    "xml-only",
    "structure only",
    "structure-only",
    "pdf export only",
    "pdf-export-only",
    "page count only",
    "page-count-only",
    "manual only",
    "manual-only",
    "manual visual only",
    "no machine vision",
    "without machine vision",
    "page-image existence",
    "page image existence",
    "blocked",
    "title-like",
    "heading-like remains",
}
CONTENT_MUTATION_VISUAL_EVIDENCE_TOKENS = (
    "machine-vision",
    "rendered",
    "exact output",
    "touched paragraph",
    "touched page",
    "body donor",
    "body-vs-heading",
    "heading contamination",
    "final verdict",
)
BIBLIOGRAPHY_COMMENT_TOKENS = ("references", "reference format", "bibliography", "literature")
FORBIDDEN_PASS_WITH_CAVEAT_TOKENS = (
    "passed with caveats",
    "pass with caveats",
    "passed with limitations",
    "pass with limitation",
    "mostly passed",
    "near pass",
    "sampled, not every paragraph",
    "sampled rendered review",
    "structural pass only",
    "structure-only pass",
    "font audit limitation accepted",
    "donor limitation accepted",
)
FORBIDDEN_STATIC_TOC_PASS_TOKENS = (
    "static toc fallback",
    "static table of contents fallback",
    "static outline-style table rather than",
    "handwritten/static toc",
)


USER_REPORTED_TRIGGER_TOKENS = (
    "user-reported",
    "reported",
    "complaint",
    "feedback",
    "user said",
    "user named",
    "\u7528\u6237\u53cd\u9988",
    "\u7528\u6237\u6307\u51fa",
    "\u70b9\u540d",
    "\u53cd\u590d\u63d0\u53ca",
)
EXPLICIT_USER_RULE_PROBLEM_TOKENS = (
    "problem",
    "issue",
    "wrong",
    "drift",
    "mismatch",
    "font",
    "indent",
    "style",
    "format",
    "\u95ee\u9898",
    "\u9519\u8bef",
    "\u5b57\u4f53",
    "\u7f29\u8fdb",
    "\u683c\u5f0f",
)
SURFACE_LEDGER_REQUIREMENTS = {
    "TOC": (
        ("toc", "table of contents", "underline", "underlined", "\u76ee\u5f55", "\u4e0b\u5212\u7ebf"),
        ("toc", "table of contents", "underline", "\u76ee\u5f55", "\u4e0b\u5212\u7ebf"),
    ),
    "body": (
        ("body", "body style", "normal baseline", "normal", "docdefaults", "\u6b63\u6587", "\u5168\u6587\u5b57\u4f53", "\u6b63\u6587\u6837\u5f0f", "\u6837\u5f0f\u6a21\u677f"),
        ("body", "body style", "normal", "\u6b63\u6587", "\u6b63\u6587\u6837\u5f0f"),
    ),
    "table": (
        (
            "table",
            "table style",
            "table family",
            "lost table",
            "three-line table",
            "tablegrid",
            "cross-page table",
            "continued table",
            "continuation table",
            "continuation title",
            "\u8868\u683c",
            "\u8868\u9898",
            "\u8868\u9898\u4e22\u5931",
            "\u8868\u683c\u6837\u5f0f",
            "\u4e09\u7ebf\u8868",
            "\u8868\u5934\u7ebf",
            "\u4e2d\u7ebf",
            "\u8868\u683c\u8de8\u9875",
            "\u8de8\u9875\u8868\u683c",
            "\u7eed\u8868",
            "\u7eed\u8868\u6807\u9898",
            "\u4e22\u5931",
        ),
        (
            "table",
            "table style",
            "three-line table",
            "continuation",
            "\u8868\u683c",
            "\u8868\u9898",
            "\u8868\u683c\u6837\u5f0f",
            "\u4e09\u7ebf\u8868",
            "\u8868\u683c\u8de8\u9875",
            "\u7eed\u8868",
        ),
    ),
    "figure": (
        (
            "figure",
            "diagram",
            "flowchart",
            "draw.io",
            "drawio",
            "svg",
            "screenshot",
            "algorithm result",
            "algorithm-result",
            "yolo",
            "yolov8",
            "dbnet",
            "crnn",
            "ocr",
            "\u56fe",
            "\u753b\u56fe",
            "\u622a\u56fe",
            "\u793a\u610f\u56fe",
            "\u68c0\u6d4b\u6548\u679c",
            "\u8bc6\u522b\u7ed3\u679c",
        ),
        ("figure", "diagram", "flowchart", "draw.io", "drawio", "svg", "screenshot", "algorithm result", "\u56fe", "\u622a\u56fe"),
    ),
    "heading": (
        (
            "heading",
            "heading 2",
            "heading 3",
            "body_heading_levels",
            "body_text",
            "body text",
            "mixed-script body text",
            "body heading levels",
            "title indentation",
            "heading indentation",
            "abnormal indentation",
            "left indent",
            "hanging indent",
            "first-line indent",
            "rendered left-x",
            "centerline",
            "\u4e8c\u7ea7\u6807\u9898",
            "\u4e09\u7ea7\u6807\u9898",
            "\u6807\u9898\u6837\u5f0f",
            "\u6b63\u6587\u6807\u9898",
            "\u6807\u9898",
            "\u5f02\u5e38\u7f29\u8fdb",
        ),
        ("heading", "body_heading_levels", "body heading levels", "body_text", "body text", "mixed-script body text", "indentation", "\u4e8c\u7ea7\u6807\u9898", "\u4e09\u7ea7\u6807\u9898", "\u6b63\u6587\u6807\u9898", "\u6b63\u6587", "\u6807\u9898", "\u7f29\u8fdb"),
    ),
    "body-opener-header-title": (
        BODY_OPENER_HEADER_TITLE_TOKENS,
        (
            "body opener",
            "running header",
            "header title",
            "title/header",
            "chapter opener",
            "\u6b63\u6587\u6807\u9898",
            "\u9875\u7709",
            "\u9875\u7709\u6807\u9898",
            "\u7ae0\u9996",
            "\u7eea\u8bba",
            "\u7ed3\u8bba",
        ),
    ),
    "header-footer-page-number": (
        (
            "header",
            "footer",
            "page number",
            "page-number",
            "page_numbers",
            "running header",
            "header line",
            "horizontal line",
            "footer position",
            "\u9875\u7709",
            "\u9875\u811a",
            "\u9875\u7801",
            "\u6a2a\u7ebf",
        ),
        (
            "header",
            "footer",
            "page number",
            "header line",
            "horizontal line",
            "\u9875\u7709",
            "\u9875\u811a",
            "\u9875\u7801",
            "\u6a2a\u7ebf",
        ),
    ),
    "pagination": (
        (
            "pagination",
            "blank page",
            "blank-page",
            "near-empty",
            "whole_document_pagination",
            "whole-document pagination",
            "\u5206\u9875",
            "\u7a7a\u9875",
            "\u8fd1\u7a7a\u9875",
            "\u5168\u6587\u5206\u9875",
        ),
        ("pagination", "whole_document_pagination", "whole-document pagination", "\u5206\u9875", "\u7a7a\u9875", "\u5168\u6587\u5206\u9875"),
    ),
    "template-format": (
        ("template-format", "template format", "template alignment", "\u6a21\u677f\u683c\u5f0f", "\u683c\u5f0f\u6a21\u677f", "\u683c\u5f0f"),
        ("template-format", "template", "\u6a21\u677f", "\u683c\u5f0f"),
    ),
    "font": (
        ("font", "font family", "theme font", "wrong font", "\u5b57\u4f53", "\u5b8b\u4f53\uff08\u6b63\u6587\uff09"),
        ("font", "font family", "\u5b57\u4f53"),
    ),
    "references": (
        ("reference", "references", "bibliography", "references_title", "references_entries", "left indent", "hanging indent", "rendered left-x", "\u53c2\u8003\u6587\u732e", "\u7f29\u8fdb"),
        ("reference", "references", "bibliography", "references_title", "references_entries", "left-x", "indentation", "\u53c2\u8003\u6587\u732e", "\u7f29\u8fdb"),
    ),
    "citation": (
        ("citation", "superscript", "\u5f15\u7528", "\u4e0a\u6807"),
        ("citation", "superscript", "\u5f15\u7528", "\u4e0a\u6807"),
    ),
    "acknowledgement": (
        ("acknowledgement", "acknowledgment", "acknowledgement_title", "acknowledgement_body", "title indentation", "body indentation", "\u81f4\u8c22", "\u7f29\u8fdb"),
        ("acknowledgement", "acknowledgment", "acknowledgement_title", "acknowledgement_body", "indentation", "\u81f4\u8c22", "\u7f29\u8fdb"),
    ),
    "appendix": (
        ("appendix", "\u9644\u5f55"),
        ("appendix", "\u9644\u5f55"),
    ),
    "cross-surface-regression": (
        (
            "cross-surface",
            "regression",
            "style-blast-radius",
            "fix one",
            "break another",
            "repeat",
            "repeated",
            "loop",
            "cycle",
            "body style",
            "normal baseline",
            "body style normal",
            "toc underline",
            "table style lost",
            "\u53cd\u590d\u72af\u9519",
            "\u5faa\u73af",
            "\u4fee\u4e00\u6b21",
            "\u574f\u4e00\u6b21",
            "\u4fee\u4e00\u4e2a\u5730\u65b9",
            "\u53e6\u4e00\u4e2a\u5730\u65b9",
            "\u4e0b\u5212\u7ebf",
            "\u8868\u683c\u6837\u5f0f",
            "\u4e22\u5931",
        ),
        ("cross-surface", "regression", "style-blast-radius", "\u4fee\u4e00\u4e2a\u5730\u65b9", "\u53e6\u4e00\u4e2a\u5730\u65b9"),
    ),
    "format-preservation": (
        (
            "preserve formatting",
            "format preservation",
            "do not break format",
            "same format",
            "chapter format",
            "\u4e0d\u8981\u7834\u574f\u683c\u5f0f",
            "\u4fdd\u6301\u539f\u683c\u5f0f",
            "\u4e0d\u6539\u683c\u5f0f",
            "\u683c\u5f0f\u672a\u7834\u574f",
            "\u7b2c\u4e03\u7ae0",
        ),
        ("format preservation", "chapter format", "\u4fdd\u6301\u683c\u5f0f", "\u7ae0\u8282\u683c\u5f0f"),
    ),
}

MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES = {
    "cover_style": ("cover_style", "cover style", "cover", "\u5c01\u9762"),
    "declaration_or_title_front_matter": (
        "declaration_or_title_front_matter",
        "declaration",
        "title front matter",
        "title/front matter",
        "\u58f0\u660e",
        "\u9898\u540d\u9875",
        "\u627f\u8bfa",
    ),
    "zh_abstract_title": ("zh_abstract_title", "chinese abstract title", "\u4e2d\u6587\u6458\u8981\u6807\u9898"),
    "zh_abstract_body": ("zh_abstract_body", "chinese abstract body", "\u4e2d\u6587\u6458\u8981\u6b63\u6587"),
    "zh_keyword_line": ("zh_keyword_line", "chinese keyword line", "\u4e2d\u6587\u5173\u952e\u8bcd"),
    "en_abstract_title": ("en_abstract_title", "english abstract title"),
    "en_abstract_body": ("en_abstract_body", "english abstract body"),
    "en_keyword_line": ("en_keyword_line", "english keyword line", "key words", "keywords"),
    "toc_title": ("toc_title", "toc title", "table of contents title", "\u76ee\u5f55\u6807\u9898"),
    "toc_entries": ("toc_entries", "toc entries", "table of contents entries", "\u76ee\u5f55\u6761\u76ee"),
    "toc_dotted_leaders": ("toc_dotted_leaders", "dotted leaders", "\u70b9\u5f15\u5bfc\u7ebf"),
    "toc_page_number_column": ("toc_page_number_column", "page-number column", "page number column", "\u9875\u7801\u5217"),
    "body_heading_levels": ("body_heading_levels", "body heading levels", "heading levels", "\u6b63\u6587\u6807\u9898"),
    "body_text": ("body_text", "body text", "mixed-script body text", "\u6b63\u6587", "\u6b63\u6587\u6bb5\u843d"),
    "body_citation_superscripts": (
        "body_citation_superscripts",
        "body citation superscripts",
        "citation superscript",
        "\u5f15\u7528\u4e0a\u6807",
    ),
    "review_comments_and_change_marks": (
        "review_comments_and_change_marks",
        "review comments",
        "comment anchors",
        "tracked changes",
        "\u6279\u6ce8",
        "\u4fee\u8ba2",
    ),
    "figure_table_captions_and_holders": (
        "figure_table_captions_and_holders",
        "figure/table captions",
        "figure holder",
        "table title",
        "\u56fe\u9898",
        "\u8868\u9898",
    ),
    "references_title": ("references_title", "references title", "bibliography title", "\u53c2\u8003\u6587\u732e\u6807\u9898"),
    "references_entries": (
        "references_entries",
        "reference entries",
        "references entries",
        "bibliography entries",
        "\u53c2\u8003\u6587\u732e\u6761\u76ee",
    ),
    "acknowledgement_title": ("acknowledgement_title", "acknowledgement title", "\u81f4\u8c22\u6807\u9898"),
    "acknowledgement_body": ("acknowledgement_body", "acknowledgement body", "\u81f4\u8c22\u6b63\u6587"),
    "appendix_title": ("appendix_title", "appendix title", "\u9644\u5f55\u6807\u9898"),
    "appendix_body": ("appendix_body", "appendix body", "\u9644\u5f55\u6b63\u6587"),
    "header": ("header", "\u9875\u7709"),
    "footer": ("footer", "\u9875\u811a"),
    "page_numbers": ("page_numbers", "page numbers", "\u9875\u7801"),
    "whole_document_pagination": (
        "whole_document_pagination",
        "whole document pagination",
        "full document pagination",
        "\u5168\u6587\u5206\u9875",
    ),
}

MANDATORY_THESIS_SURFACE_ALLOWED_STATUSES = {
    "present-active",
    "present-unchanged-reviewed",
    "not-present",
    "not-applicable-with-reason",
    "blocked",
}
MANDATORY_THESIS_CRITICAL_PRESENT_SURFACES = {
    "cover_style",
    "declaration_or_title_front_matter",
    "zh_abstract_title",
    "zh_abstract_body",
    "zh_keyword_line",
    "en_abstract_title",
    "en_abstract_body",
    "en_keyword_line",
    "toc_title",
    "toc_entries",
    "toc_dotted_leaders",
    "toc_page_number_column",
    "body_heading_levels",
    "body_text",
    "body_citation_superscripts",
    "references_title",
    "references_entries",
    "acknowledgement_title",
    "acknowledgement_body",
    "header",
    "footer",
    "page_numbers",
    "whole_document_pagination",
}
MANDATORY_THESIS_VERDICT_PREFIXES = (
    "- cover style verdict:",
    "- declaration/title front matter verdict:",
    "- abstract and keyword surface verdict:",
    "- TOC visual baseline verdict:",
    "- TOC visual geometry verdict:",
    "- TOC paragraph-and-typography verdict:",
    "- body heading levels verdict:",
    "- body citation superscripts verdict:",
    "- review comments/change marks verdict:",
    "- format preservation promise verdict:",
    "- chapter format preservation detector verdict:",
    "- non-target format preservation verdict:",
    "- heading level 1 verdict:",
    "- heading level 2 verdict:",
    "- heading level 3 verdict:",
    "- heading level 4 verdict:",
    "- heading direct rPr/font/size/bold/spacing verdict:",
    "- body text verdict:",
    "- references title verdict:",
    "- references entries verdict:",
    "- acknowledgement title verdict:",
    "- acknowledgement body verdict:",
    "- acknowledgement body paragraph-dialog metrics verdict:",
    "- acknowledgement body direct-run typography verdict:",
    "- acknowledgement body first-line indent verdict:",
    "- acknowledgement body title-contamination verdict:",
    "- header verdict:",
    "- body opener/header title consistency verdict:",
    "- footer verdict:",
    "- page numbers verdict:",
    "- whole-document pagination verdict:",
    "- references entry format verdict:",
    "- review comments/change marks preservation verdict:",
    "- body citation superscripts preservation verdict:",
    "- appendix format verdict:",
)
HIGH_RISK_THESIS_FORMAT_SURFACE_ALIASES = {
    surface_id: MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES[surface_id]
    for surface_id in (
        "cover_style",
        "zh_abstract_title",
        "zh_abstract_body",
        "zh_keyword_line",
        "en_abstract_title",
        "en_abstract_body",
        "en_keyword_line",
        "toc_title",
        "toc_entries",
        "toc_dotted_leaders",
        "toc_page_number_column",
        "body_heading_levels",
        "body_text",
        "body_citation_superscripts",
        "review_comments_and_change_marks",
        "references_title",
        "references_entries",
        "acknowledgement_title",
        "acknowledgement_body",
        "appendix_title",
        "appendix_body",
    )
}
HIGH_RISK_OPTIONAL_WITH_REASON_SURFACES = {"appendix_title", "appendix_body", "review_comments_and_change_marks"}
FRONT_MATTER_SURFACE_COVERAGE_ALIASES = {
    surface_id: MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES[surface_id]
    for surface_id in (
        "cover_style",
        "declaration_or_title_front_matter",
        "zh_abstract_title",
        "zh_abstract_body",
        "zh_keyword_line",
        "en_abstract_title",
        "en_abstract_body",
        "en_keyword_line",
        "toc_title",
        "toc_entries",
        "toc_dotted_leaders",
        "toc_page_number_column",
        "header",
        "footer",
        "page_numbers",
    )
}
END_MATTER_SURFACE_COVERAGE_ALIASES = {
    surface_id: MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES[surface_id]
    for surface_id in (
        "references_title",
        "references_entries",
        "acknowledgement_title",
        "acknowledgement_body",
        "appendix_title",
        "appendix_body",
        "header",
        "footer",
        "page_numbers",
        "whole_document_pagination",
    )
}
SURFACE_COVERAGE_OPTIONAL_WITH_REASON = {"appendix_title", "appendix_body"}


def extract_docx_comment_text(docx_path: Path) -> str:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            if "word/comments.xml" not in zf.namelist():
                return ""
            root = ET.fromstring(zf.read("word/comments.xml"))
    except Exception:
        return ""
    return "".join(root.itertext())


def read_optional_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""


def transaction_figure_manifest_paths(transaction_record_path: Path) -> list[Path]:
    try:
        payload = load_transaction_record(transaction_record_path)
    except Exception:
        return []
    paths: list[Path] = []
    for key in ("figure_manifest_path", "figure_asset_manifest_path"):
        value = payload.get(key)
        if isinstance(value, dict):
            value = value.get("path")
        resolved = resolve_transaction_path(value, transaction_record_path)
        if resolved is not None:
            paths.append(resolved)
    return paths


AGENT_RUN_MANIFEST_REQUIRED_PREFIXES = (
    "- run_id:",
    "- task_mode:",
    "- authorization_source:",
    "- agent_mode:",
    "- max_concurrent_live_agents:",
    "- live_agent_count_plan:",
    "- dispatch_wave_plan:",
    "- audit_presence_by_wave:",
    "- concurrency_limit_verdict:",
    "- required_lanes:",
    "- complete_role_roster:",
    "- role_attendance_matrix:",
    "- not_applicable_lanes_with_reasons:",
    "- role_alias_map_zh:",
    "- lane_alias_map_zh:",
    "- required_lane_aliases_zh:",
    "- audit_role_alias_zh:",
    "- spawned_agent_ids:",
    "- audit_agent_id:",
    "- sequential_audit_fallback_id:",
    "- audit_spawn_or_fallback_mode:",
    "- audit_verdict:",
    "- audit_verdict_cadence:",
    "- action_audit_scope:",
    "- action_audit_verdict_cadence:",
    "- action_audit_verdicts:",
    "- action_cycles:",
    "- action_categories:",
    "- action_owner_map:",
    "- protected_surface_contract_path:",
    "- protected_surface_contract_loaded:",
    "- protected_surface_id_set:",
    "- protected_surface_owner_map:",
    "- protected_surface_evidence_map:",
    "- surface_paragraph_and_typography_evidence_paths:",
    "- surface_paragraph_and_typography_verdict:",
    "- toc_visual_geometry_evidence_paths:",
    "- toc_visual_geometry_verdict:",
    "- toc_paragraph_and_typography_evidence_paths:",
    "- toc_paragraph_and_typography_verdict:",
    "- toc_visible_run_typography_evidence_paths:",
    "- toc_visible_run_typography_verdict:",
    "- whole_document_pagination_evidence_path:",
    "- whole_document_pagination_verdict:",
    "- content_mutation_rendered_review_path:",
    "- content_mutation_machine_vision_verdict:",
    "- inserted_body_heading_contamination_verdict:",
    "- caption_table_sibling_body_contamination_verdict:",
    "- figure_user_material_source_path:",
    "- figure_user_material_source_sha256:",
    "- figure_material_source_inventory_path:",
    "- figure_material_source_binding_verdict:",
    "- figure_generated_substitute_rejection_verdict:",
    "- figure_extents_audit_path:",
    "- figure_paragraph_margin_width_verdict:",
    "- runtime_screenshot_full_window_geometry_evidence_paths:",
    "- runtime_screenshot_full_window_capture_verdict:",
    "- flowchart_semantic_trigger_inventory_path:",
    "- flowchart_drawio_source_evidence_paths:",
    "- flowchart_svg_export_evidence_paths:",
    "- flowchart_raster_fallback_evidence_paths:",
    "- flowchart_geometry_report_paths:",
    "- flowchart_rendered_page_evidence_paths:",
    "- flowchart_final_docx_relationship_evidence_paths:",
    "- flowchart_manifest_contract_verdict:",
    "- algorithm_code_visibility_evidence_path:",
    "- algorithm_code_visibility_verdict:",
    "- toc_digits_cache_preservation_evidence_path:",
    "- toc_digits_cache_preservation_verdict:",
    "- touched_page_blast_radius_machine_vision_evidence_paths:",
    "- format_lane_post_mutation_rendered_audit_verdict:",
    "- protected_surface_reviewed_output_sha256:",
    "- protected_surface_contract_verdict:",
    "- lane_task_card_paths:",
)

AGENT_TASK_CARD_REQUIRED_PREFIXES = (
    "- card_id:",
    "- run_id:",
    "- lane:",
    "- role_alias_zh:",
    "- role_applicability:",
    "- attendance_status:",
    "- not_applicable_reason:",
    "- skip_reason:",
    "- lane_alias_zh:",
    "- owner:",
    "- owner_alias_zh:",
    "- system_agent_id:",
    "- authorization_source:",
    "- agent_mode:",
    "- run_manifest_path:",
    "- objective:",
    "- inputs:",
    "- outputs:",
    "- dependencies:",
    "- spawn_requested:",
    "- spawn_status:",
    "- spawn_agent_id:",
    "- spawn_agent_alias_zh:",
    "- fallback_mode:",
    "- audit_agent_id:",
    "- sequential_audit_fallback_id:",
    "- audit_agent_alias_zh:",
    "- audit_required:",
    "- audit_spawn_or_fallback_mode:",
    "- action_audit_scope:",
    "- action_audit_verdict_cadence:",
    "- action_audit_verdicts:",
    "- action_cycles:",
    "- audit_verdict:",
    "- supervised_by:",
    "- figure_user_material_source_path:",
    "- figure_user_material_source_sha256:",
    "- figure_material_source_inventory_path:",
    "- figure_material_source_binding_verdict:",
    "- figure_generated_substitute_rejection_verdict:",
    "- figure_extents_audit_path:",
    "- figure_paragraph_margin_width_verdict:",
    "- runtime_screenshot_full_window_geometry_evidence_paths:",
    "- runtime_screenshot_full_window_capture_verdict:",
    "- flowchart_semantic_trigger_inventory_path:",
    "- flowchart_drawio_source_evidence_paths:",
    "- flowchart_svg_export_evidence_paths:",
    "- flowchart_raster_fallback_evidence_paths:",
    "- flowchart_geometry_report_paths:",
    "- flowchart_rendered_page_evidence_paths:",
    "- flowchart_final_docx_relationship_evidence_paths:",
    "- flowchart_manifest_contract_verdict:",
    "- algorithm_code_visibility_evidence_path:",
    "- algorithm_code_visibility_verdict:",
    "- toc_digits_cache_preservation_evidence_path:",
    "- toc_digits_cache_preservation_verdict:",
    "- evidence_required:",
    "- evidence_paths:",
)

FORMAT_TASK_CARD_REQUIRED_PREFIXES = (
    "- project_template_discovery_root:",
    "- template_discovery_patterns:",
    "- discovered_candidate_template_paths:",
    "- candidate_template_selection_reason:",
    "- active_template_source_type:",
    "- active_template_path_lock:",
    "- active_template_fingerprint:",
    "- active_template_profile_path:",
    "- active_template_selected_before_mutation:",
    "- template_alignment_evidence_required:",
    "- template_alignment_evidence_paths:",
    "- template_alignment_verdict:",
    "- mandatory_thesis_surface_inventory_path:",
    "- front_matter_surface_coverage_matrix_path:",
    "- end_matter_surface_coverage_matrix_path:",
    "- high_risk_thesis_format_surface_matrix_path:",
    "- surface_inventory_verdict:",
    "- touched_paragraph_ids:",
    "- touched_surface_families:",
    "- sibling_surfaces:",
    "- blast_radius_pages:",
    "- stale_audits:",
    "- rerender_targets:",
    "- reviewed_output_path:",
    "- reviewed_output_sha256:",
    "- protected_surface_contract_path:",
    "- protected_surface_contract_loaded:",
    "- canonical_protected_surface_ids_in_scope:",
    "- protected_surface_owner_map:",
    "- protected_surface_evidence_map:",
    "- surface_paragraph_and_typography_evidence_paths:",
    "- surface_paragraph_and_typography_verdict:",
    "- toc_visual_geometry_evidence_paths:",
    "- toc_visual_geometry_verdict:",
    "- toc_paragraph_and_typography_evidence_paths:",
    "- toc_paragraph_and_typography_verdict:",
    "- toc_visible_run_typography_evidence_paths:",
    "- toc_visible_run_typography_verdict:",
    "- whole_document_pagination_evidence_path:",
    "- whole_document_pagination_verdict:",
    "- content_mutation_rendered_review_path:",
    "- content_mutation_machine_vision_verdict:",
    "- inserted_body_heading_contamination_verdict:",
    "- figure_user_material_source_path:",
    "- figure_user_material_source_sha256:",
    "- figure_material_source_inventory_path:",
    "- figure_material_source_binding_verdict:",
    "- figure_generated_substitute_rejection_verdict:",
    "- figure_extents_audit_path:",
    "- figure_paragraph_margin_width_verdict:",
    "- runtime_screenshot_full_window_geometry_evidence_paths:",
    "- runtime_screenshot_full_window_capture_verdict:",
    "- flowchart_semantic_trigger_inventory_path:",
    "- flowchart_drawio_source_evidence_paths:",
    "- flowchart_svg_export_evidence_paths:",
    "- flowchart_raster_fallback_evidence_paths:",
    "- flowchart_geometry_report_paths:",
    "- flowchart_rendered_page_evidence_paths:",
    "- flowchart_final_docx_relationship_evidence_paths:",
    "- flowchart_manifest_contract_verdict:",
    "- algorithm_code_visibility_evidence_path:",
    "- algorithm_code_visibility_verdict:",
    "- toc_digits_cache_preservation_evidence_path:",
    "- toc_digits_cache_preservation_verdict:",
    "- touched_page_blast_radius_machine_vision_evidence_paths:",
    "- format_lane_post_mutation_rendered_audit_verdict:",
    "- protected_surface_reviewed_output_sha256:",
    "- protected_surface_contract_verdict:",
)

FORMAT_CARD_TO_GATE_PREFIX = {
    "- active_template_source_type:": "- active template source type:",
    "- active_template_path_lock:": "- active template path lock:",
    "- active_template_fingerprint:": "- active template fingerprint:",
    "- active_template_profile_path:": "- active template profile path:",
    "- active_template_selected_before_mutation:": "- active template selected before mutation?:",
    "- template_alignment_verdict:": "- template alignment verdict:",
}

AGENT_PROTECTED_FIELD_TO_GATE_PREFIX = {
    "- protected_surface_contract_path:": "- protected-surface evidence contract path:",
    "- protected_surface_contract_loaded:": "- protected-surface evidence contract loaded?:",
    "- protected_surface_id_set:": "- canonical protected surface id set:",
    "- canonical_protected_surface_ids_in_scope:": "- canonical protected surface id set:",
    "- protected_surface_owner_map:": "- protected-surface owner map:",
    "- protected_surface_evidence_map:": "- protected-surface evidence map:",
    "- surface_paragraph_and_typography_evidence_paths:": "- surface paragraph-and-typography audit evidence path:",
    "- surface_paragraph_and_typography_verdict:": "- surface paragraph-and-typography verdict:",
    "- toc_visual_geometry_evidence_paths:": "- toc visual geometry evidence paths:",
    "- toc_visual_geometry_verdict:": "- TOC visual geometry verdict:",
    "- toc_paragraph_and_typography_evidence_paths:": "- toc paragraph-and-typography evidence paths:",
    "- toc_paragraph_and_typography_verdict:": "- TOC paragraph-and-typography verdict:",
    "- toc_visible_run_typography_evidence_paths:": "- toc visible-run typography evidence paths:",
    "- toc_visible_run_typography_verdict:": "- TOC visible-run typography verdict:",
    "- whole_document_pagination_evidence_path:": "- whole-document pagination evidence path:",
    "- whole_document_pagination_verdict:": "- whole-document pagination verdict:",
    "- content_mutation_rendered_review_path:": "- content mutation rendered-page review path:",
    "- content_mutation_machine_vision_verdict:": "- content mutation machine-vision verdict:",
    "- inserted_body_heading_contamination_verdict:": "- inserted body heading-contamination verdict:",
    "- caption_table_sibling_body_contamination_verdict:": "- caption/table sibling body contamination verdict:",
    "- figure_user_material_source_path:": "- figure user material source path:",
    "- figure_user_material_source_sha256:": "- figure user material source SHA256:",
    "- figure_material_source_inventory_path:": "- figure material-source inventory path:",
    "- figure_material_source_binding_verdict:": "- figure material-source binding verdict:",
    "- figure_generated_substitute_rejection_verdict:": "- figure generated-substitute rejection verdict:",
    "- figure_extents_audit_path:": "- figure extents audit path:",
    "- figure_paragraph_margin_width_verdict:": "- figure paragraph-margin width verdict:",
    "- figure_visible_content_completeness_verdict:": "- figure visible-content completeness verdict:",
    "- figure_nonzero_crop_count:": "- figure nonzero crop count:",
    "- figure_exact_line_spacing_clipping_count:": "- figure exact-line-spacing clipping count:",
    "- figure_image_holder_layout_verdict:": "- figure image-holder layout verdict:",
    "- figure_image_holder_layout_issue_count:": "- figure image-holder layout issue count:",
    "- runtime_screenshot_full_window_geometry_evidence_paths:": "- runtime screenshot full-window geometry evidence paths:",
    "- runtime_screenshot_full_window_capture_verdict:": "- runtime screenshot full-window capture verdict:",
    "- flowchart_semantic_trigger_inventory_path:": "- flowchart semantic trigger inventory path:",
    "- flowchart_drawio_source_evidence_paths:": "- flowchart draw.io source evidence paths:",
    "- flowchart_svg_export_evidence_paths:": "- flowchart SVG export evidence paths:",
    "- flowchart_raster_fallback_evidence_paths:": "- flowchart raster fallback evidence paths:",
    "- flowchart_geometry_report_paths:": "- flowchart source-scale geometry report paths:",
    "- flowchart_rendered_page_evidence_paths:": "- flowchart rendered-page evidence paths:",
    "- flowchart_final_docx_relationship_evidence_paths:": "- flowchart final DOCX relationship evidence paths:",
    "- flowchart_manifest_contract_verdict:": "- flowchart manifest contract verdict:",
    "- algorithm_code_visibility_evidence_path:": "- algorithm code visibility evidence path:",
    "- algorithm_code_visibility_verdict:": "- algorithm code visibility verdict:",
    "- toc_digits_cache_preservation_evidence_path:": "- TOC digits/cache preservation evidence path:",
    "- toc_digits_cache_preservation_verdict:": "- TOC digits/cache preservation verdict:",
    "- touched_page_blast_radius_machine_vision_evidence_paths:": "- touched-page/blast-radius machine-vision evidence paths:",
    "- format_lane_post_mutation_rendered_audit_verdict:": "- format lane post-mutation rendered audit verdict:",
    "- protected_surface_reviewed_output_sha256:": "- protected-surface reviewed output sha256:",
    "- protected_surface_contract_verdict:": "- protected-surface evidence contract verdict:",
    "- protected-surface evidence contract path:": "- protected-surface evidence contract path:",
    "- protected-surface evidence contract loaded?:": "- protected-surface evidence contract loaded?:",
    "- canonical protected surface id set:": "- canonical protected surface id set:",
    "- protected-surface owner map:": "- protected-surface owner map:",
    "- protected-surface evidence map:": "- protected-surface evidence map:",
    "- all-surface paragraph-dialog / typography evidence record paths:": "- surface paragraph-and-typography audit evidence path:",
    "- TOC visual geometry evidence record paths:": "- toc visual geometry evidence paths:",
    "- TOC paragraph-and-typography evidence record paths:": "- toc paragraph-and-typography evidence paths:",
    "- TOC visible-run typography evidence record paths:": "- toc visible-run typography evidence paths:",
    "- whole-document pagination evidence record path:": "- whole-document pagination evidence path:",
    "- protected-surface reviewed output sha256:": "- protected-surface reviewed output sha256:",
    "- protected-surface evidence contract verdict:": "- protected-surface evidence contract verdict:",
}

AGENT_PROTECTED_PATH_PREFIXES = {
    "- protected_surface_contract_path:",
    "- protected_surface_evidence_map:",
    "- surface_paragraph_and_typography_evidence_paths:",
    "- toc_visual_geometry_evidence_paths:",
    "- toc_paragraph_and_typography_evidence_paths:",
    "- toc_visible_run_typography_evidence_paths:",
    "- whole_document_pagination_evidence_path:",
    "- content_mutation_rendered_review_path:",
    "- figure_user_material_source_path:",
    "- figure_material_source_inventory_path:",
    "- figure_extents_audit_path:",
    "- runtime_screenshot_full_window_geometry_evidence_paths:",
    "- flowchart_semantic_trigger_inventory_path:",
    "- flowchart_drawio_source_evidence_paths:",
    "- flowchart_svg_export_evidence_paths:",
    "- flowchart_raster_fallback_evidence_paths:",
    "- flowchart_geometry_report_paths:",
    "- flowchart_rendered_page_evidence_paths:",
    "- flowchart_final_docx_relationship_evidence_paths:",
    "- algorithm_code_visibility_evidence_path:",
    "- toc_digits_cache_preservation_evidence_path:",
    "- touched_page_blast_radius_machine_vision_evidence_paths:",
    "- protected-surface evidence contract path:",
    "- protected-surface evidence map:",
    "- all-surface paragraph-dialog / typography evidence record paths:",
    "- TOC visual geometry evidence record paths:",
    "- TOC paragraph-and-typography evidence record paths:",
    "- TOC visible-run typography evidence record paths:",
    "- whole-document pagination evidence record path:",
}


def validate_agent_concurrency_fields(
    *,
    max_live_agents: str,
    live_agent_count_plan: str,
    dispatch_wave_plan: str,
    audit_presence_by_wave: str,
    concurrency_limit_verdict: str,
    kind: str,
) -> list[str]:
    issues: list[str] = []
    try:
        max_live = int(max_live_agents)
    except (TypeError, ValueError):
        issues.append(f"{kind} max concurrent live agents must be an integer no greater than 6")
    else:
        if max_live < 1 or max_live > 6:
            issues.append(f"{kind} max concurrent live agents must be between 1 and 6")
    for label, value in (
        ("live agent count plan", live_agent_count_plan),
        ("dispatch wave plan", dispatch_wave_plan),
        ("audit presence by wave", audit_presence_by_wave),
    ):
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"{kind} must explicitly record {label}")
    audit_presence_norm = normalize(audit_presence_by_wave).lower()
    if "audit" not in audit_presence_norm and "\u5ba1\u6838" not in audit_presence_by_wave:
        issues.append(f"{kind} audit presence by wave must name audit presence")
    if concurrency_limit_verdict != "pass":
        issues.append(f"{kind} concurrency limit verdict must be pass")
    return issues


def parse_prefixed_record(path: Path, prefixes: tuple[str, ...], kind: str) -> tuple[dict[str, str], dict[str, str], list[str]]:
    issues: list[str] = []
    try:
        lines = read_lines(path)
    except UnicodeDecodeError as exc:
        return {}, {}, [f"{kind} is not valid UTF-8: {path} ({exc})"]
    except OSError as exc:
        return {}, {}, [f"{kind} could not be read: {path} ({exc})"]
    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    for prefix in prefixes:
        matched = find_lines_with_prefix(lines, prefix)
        if len(matched) != 1:
            issues.append(f"{kind} must contain exactly one '{prefix}' line: {path}")
            continue
        values[prefix] = parse_line_value(matched[0])
        raw_values[prefix] = raw_line_value(matched[0])
    return values, raw_values, issues


def skill_lock_passish(value: str) -> bool:
    normalized = normalize(value).lower()
    return (
        normalized == "pass"
        or normalized.startswith("pass ")
        or normalized.startswith("passed")
        or normalized in {"yes", "verified", "audit-only-pass"}
    )


def skill_lock_noish(value: str) -> bool:
    normalized = normalize(value).lower()
    return normalized in {"no", "false", "not used", "not-used", "none", "no substitute", "no-substitute"}


def skill_lock_blocked_evidence_ok(value: str) -> bool:
    normalized = normalize(value).lower()
    if not normalized:
        return False
    if normalized in {"none", "no blocked evidence", "no-blocked-evidence", "not-applicable"}:
        return True
    if contains_any(normalized, {"known caveat", "remaining risk", "residual", "deferred", "ignored", "handwave"}):
        return False
    return skill_lock_passish(normalized) or normalized.startswith("escalated") or normalized.startswith("blocked handoff")


def _is_mechanical_cad_final_acceptance_prefix(prefix: str) -> bool:
    lowered = prefix.lower()
    return (
        "mechanical drawing" in lowered
        or "mechanical render readability" in lowered
        or "cad appendix" in lowered
        or "cad official command" in lowered
        or "cad executable or com progid" in lowered
        or "external cad case" in lowered
        or "non-cad fallback" in lowered
        or "no verbatim geometry" in lowered
        or "reference-use restriction" in lowered
        or "exact final cad" in lowered
        or "exact dwg package" in lowered
        or "exact combined drawing pdf" in lowered
        or lowered in {"- text text overlap count:", "- text graphic cover count:", "- severe line crowding count:"}
    )


def _mechanical_cad_final_acceptance_scope(record_text_lower: str) -> bool:
    return contains_any(
        record_text_lower,
        {
            "mechanical cad",
            "cad package",
            "dwg package",
            "drawing package",
            "mechanical drawing package audit",
            "cad appendix binding",
            "exact final cad delivery package",
            "刮板输送机图纸",
            "机械图纸",
            "图纸包",
        },
    )


def _mechanical_cad_final_acceptance_scope_from_record(record_lines: list[str]) -> bool:
    context_parts: list[str] = []
    for prefix in (
        "- subtask:",
        "- thesis lane lock:",
        "- explicit user overrides:",
        "- exact final CAD delivery package path:",
        "- exact DWG package path:",
        "- exact combined drawing PDF path:",
        "- mechanical drawing package audit path:",
        "- mechanical drawing package exact package path:",
        "- CAD appendix binding audit path:",
    ):
        lines = find_lines_with_prefix(record_lines, prefix)
        if not lines:
            continue
        value = parse_line_value(lines[0])
        if is_explicit(value) and not is_explicit_none(value):
            context_parts.append(raw_line_value(lines[0]))
    return _mechanical_cad_final_acceptance_scope("\n".join(context_parts).lower())


def _mechanical_cad_field_maps(record_lines: list[str]) -> tuple[dict[str, str], dict[str, str], list[str]]:
    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    issues: list[str] = []
    for prefix in MECHANICAL_CAD_ACCEPTANCE_PREFIXES:
        lines = find_lines_with_prefix(record_lines, prefix)
        if len(lines) != 1:
            issues.append(f"mechanical CAD acceptance record must contain exactly one '{prefix}' line")
            continue
        values[prefix] = parse_line_value(lines[0])
        raw_values[prefix] = raw_line_value(lines[0])
    return values, raw_values, issues


def _mechanical_cad_passish(value: str) -> bool:
    text = normalize(value).lower()
    return text in MECHANICAL_CAD_PASS_VALUES or text.startswith(("pass;", "pass ", "passed;", "passed "))


def _mechanical_cad_float(value: object) -> float | None:
    text = normalize(str(value)).lower()
    for token in ("mm", "px", "毫米", "像素", ">= ", ">=", "约", "：", ":"):
        text = text.replace(token, "")
    text = text.strip()
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _mechanical_cad_int(value: object) -> int | None:
    parsed = _mechanical_cad_float(value)
    if parsed is None:
        return None
    return int(parsed)


def _mechanical_cad_resolve_file(
    *,
    record_path: Path,
    raw_value: str,
    label: str,
    required_suffixes: set[str] | None = None,
) -> tuple[Path | None, list[str]]:
    issues: list[str] = []
    raw_paths = split_path_values(raw_value)
    if not raw_paths:
        return None, [f"mechanical CAD acceptance {label} must name a file path"]
    resolved = resolve_record_path(raw_paths[0], record_path)
    issues.extend(validate_existing_path(resolved, require_nonempty_file=True))
    if required_suffixes is not None and resolved.suffix.lower() not in required_suffixes:
        issues.append(
            f"mechanical CAD acceptance {label} must use one of {sorted(required_suffixes)}: {resolved}"
        )
    return resolved, issues


def _mechanical_cad_validate_sha(path: Path | None, expected: str, label: str) -> list[str]:
    if path is None:
        return []
    value = normalize(expected).upper()
    if not re.fullmatch(r"[0-9A-F]{64}", value):
        return [f"mechanical CAD acceptance {label} SHA256 must be a 64-hex value"]
    actual = sha256_file(path).upper()
    if actual != value:
        return [f"mechanical CAD acceptance {label} SHA256 mismatch: expected={value} actual={actual} path={path}"]
    return []


def _mechanical_cad_gate_record_arg(command_text: str) -> str:
    match = re.search(r"--gate-record(?:=|\s+)(\"[^\"]+\"|'[^']+'|`[^`]+`|\S+)", command_text)
    if match:
        return match.group(1).strip().strip("`").strip('"').strip("'")
    return ""


def _mechanical_cad_validate_validation_command(command_text: str, record_path: Path) -> list[str]:
    issues: list[str] = []
    normalized_command = normalize(command_text)
    if "validate_skill_gate" not in normalized_command or "--gate-record" not in normalized_command:
        issues.append("mechanical CAD acceptance validation command must call validate_skill_gate with --gate-record")
        return issues
    gate_record_arg = _mechanical_cad_gate_record_arg(command_text)
    if not gate_record_arg:
        issues.append("mechanical CAD acceptance validation command must name the exact gate record")
        return issues
    referenced = resolve_record_path(gate_record_arg, record_path)
    try:
        if referenced.resolve() != record_path.resolve():
            issues.append(f"mechanical CAD acceptance validation command --gate-record must point to this record: {referenced}")
    except OSError:
        issues.append(f"mechanical CAD acceptance validation command --gate-record could not be resolved: {gate_record_arg}")
    return issues


def _mechanical_cad_audit_path_value(report: dict[str, object], key: str) -> str:
    value = report.get(key)
    if isinstance(value, str):
        return value
    candidate = report.get("candidate")
    if isinstance(candidate, dict):
        nested = candidate.get(key)
        if isinstance(nested, str):
            return nested
    return ""


def _mechanical_cad_acceptance_to_report_field(prefix: str) -> str:
    mapping = {
        "- mechanical drawing rendered no-overlap verdict:": "no_overlap_verdict",
        "- mechanical drawing boundary clearance verdict:": "boundary_clearance_verdict",
        "- mechanical drawing detail density verdict:": "detail_density_verdict",
        "- mechanical drawing title block/table/notes isolation verdict:": "title_block_table_notes_isolation_verdict",
        "- mechanical drawing title-block cell containment verdict:": "title_block_cell_containment_verdict",
        "- mechanical drawing title-block short-line topology audit verdict:": "title_block_short_line_topology_verdict",
        "- mechanical drawing annotation margin clearance verdict:": "annotation_margin_clearance_verdict",
        "- mechanical drawing local crowding verdict:": "local_crowding_verdict",
        "- mechanical drawing text/table/frame overlap verdict:": "no_overlap_verdict",
        "- mechanical drawing text integrity audit verdict:": "text_integrity_verdict",
        "- mechanical drawing text orientation audit verdict:": "text_orientation_verdict",
        "- mechanical drawing content-overlap audit verdict:": "content_overlap_verdict",
        "- mechanical drawing outside-frame ink audit verdict:": "outside_frame_ink_verdict",
        "- mechanical drawing inner-frame safe-margin audit verdict:": "inner_frame_safe_margin_verdict",
        "- mechanical drawing hatch/section fill clipping audit verdict:": "hatch_section_fill_clipping_verdict",
    }
    return mapping.get(prefix, "")


def _mechanical_cad_list_empty(value: str) -> bool:
    text = normalize(str(value)).lower()
    return text in {"[]", "none", "no failed sheets", "no-failed-sheets", "0", ""}


def _mechanical_cad_final_sha_evidence_present(value: str) -> bool:
    text = normalize(str(value))
    if not text or "not-applicable" in text.lower():
        return False
    return bool(re.search(r"[0-9a-fA-F]{64}", text))


def _mechanical_cad_resolved_json(
    *,
    record_path: Path,
    raw_values: dict[str, str],
    prefix: str,
    label: str,
) -> tuple[Path | None, dict[str, object], list[str]]:
    issues: list[str] = []
    raw_value = raw_values.get(prefix, "")
    raw_paths = split_path_values(raw_value)
    if not raw_paths:
        return None, {}, [f"{label} must name a JSON evidence path"]
    resolved = resolve_record_path(raw_paths[0], record_path)
    issues.extend(validate_existing_path(resolved, require_nonempty_file=True))
    if resolved.suffix.lower() != ".json":
        issues.append(f"{label} must be JSON: {resolved}")
    if issues:
        return resolved, {}, issues
    try:
        payload = json.loads(resolved.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        return resolved, {}, [f"{label} must be readable UTF-8 JSON: {resolved} ({exc})"]
    if not isinstance(payload, dict):
        return resolved, {}, [f"{label} must contain a JSON object: {resolved}"]
    return resolved, payload, []


def _mechanical_cad_validate_render_readability_fields(
    *,
    record_path: Path,
    values: dict[str, str],
    raw_values: dict[str, str],
) -> list[str]:
    issues: list[str] = []
    for prefix in (
        "- mechanical render readability audit path:",
        "- mechanical render readability verdict:",
        "- text text overlap count:",
        "- text graphic cover count:",
        "- severe line crowding count:",
        "- mechanical render readability small-tile threshold:",
        "- mechanical render readability failed sheet list:",
        "- mechanical render readability requires crop review count:",
        "- mechanical render readability line bundle blocker count:",
        "- mechanical render readability diagnostic-overlay-free crop review path:",
        "- mechanical render readability standard-sheet view occupancy verdict:",
        "- mechanical render readability final PDF/PNG SHA evidence:",
        "- mechanical drawing user-reported text-cover crop audit path:",
        "- mechanical drawing diagnostic-overlay-free after-crop paths:",
        "- mechanical drawing reported crop blocker count:",
        "- mechanical drawing text exclusion halo audit path:",
        "- mechanical drawing text exclusion halo audit verdict:",
        "- mechanical drawing text exclusion halo violation count:",
        "- mechanical drawing diagonal hatch/section/flow text-cover count:",
    ):
        value = values.get(prefix, "")
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"mechanical CAD render readability evidence must bind {prefix[2:-1]}")
    if issues:
        return issues
    if not _mechanical_cad_passish(values["- mechanical render readability verdict:"]):
        issues.append("mechanical CAD render readability verdict must be pass-shaped")
    for prefix in (
        "- text text overlap count:",
        "- text graphic cover count:",
        "- severe line crowding count:",
        "- mechanical render readability requires crop review count:",
        "- mechanical render readability line bundle blocker count:",
    ):
        count = _mechanical_cad_int(values[prefix])
        if count is None or count != 0:
            issues.append(f"mechanical CAD render readability {prefix[2:]} must be numeric 0")
    if not _mechanical_cad_list_empty(values["- mechanical render readability failed sheet list:"]):
        issues.append("mechanical CAD render readability failed sheet list must be empty")
    if not _mechanical_cad_passish(values["- mechanical render readability standard-sheet view occupancy verdict:"]):
        issues.append("mechanical CAD render readability standard-sheet view occupancy verdict must be pass-shaped")
    if not _mechanical_cad_final_sha_evidence_present(values["- mechanical render readability final PDF/PNG SHA evidence:"]):
        issues.append("mechanical CAD render readability final PDF/PNG SHA evidence must include SHA256 rows")
    if not _mechanical_cad_passish(values["- mechanical drawing text exclusion halo audit verdict:"]):
        issues.append("mechanical CAD text exclusion halo audit verdict must be pass-shaped")
    for prefix in (
        "- mechanical drawing reported crop blocker count:",
        "- mechanical drawing text exclusion halo violation count:",
        "- mechanical drawing diagonal hatch/section/flow text-cover count:",
    ):
        count = _mechanical_cad_int(values[prefix])
        if count is None or count != 0:
            issues.append(f"mechanical CAD text-cover crop/halo evidence {prefix[2:]} must be numeric 0")
    for raw_path in split_path_values(
        raw_values["- mechanical render readability diagnostic-overlay-free crop review path:"]
    ):
        crop_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(crop_path, require_nonempty_file=True))
        if crop_path.suffix.lower() not in IMAGE_EXTENSIONS | PDF_EXTENSIONS:
            issues.append(f"mechanical CAD render readability crop review path must be image/PDF: {crop_path}")
    for prefix in (
        "- mechanical drawing user-reported text-cover crop audit path:",
        "- mechanical drawing text exclusion halo audit path:",
    ):
        for raw_path in split_path_values(raw_values[prefix]):
            audit_path = resolve_record_path(raw_path, record_path)
            issues.extend(validate_existing_path(audit_path, require_nonempty_file=True))
    for raw_path in split_path_values(raw_values["- mechanical drawing diagnostic-overlay-free after-crop paths:"]):
        crop_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(crop_path, require_nonempty_file=True))
        if crop_path.suffix.lower() not in IMAGE_EXTENSIONS | PDF_EXTENSIONS:
            issues.append(f"mechanical CAD diagnostic-overlay-free after-crop path must be image/PDF: {crop_path}")

    _, report, path_issues = _mechanical_cad_resolved_json(
        record_path=record_path,
        raw_values=raw_values,
        prefix="- mechanical render readability audit path:",
        label="mechanical CAD render readability report",
    )
    issues.extend(path_issues)
    if path_issues:
        return issues
    if report.get("schema") != "graduation-project-builder.mechanical-render-readability.v1":
        issues.append("mechanical CAD render readability report must use mechanical-render-readability schema v1")
    if report.get("passed") is not True:
        issues.append("mechanical CAD render readability report must have passed=true")
    for prefix, report_key in (
        ("- text text overlap count:", "text_text_overlap_count"),
        ("- text graphic cover count:", "text_graphic_cover_count"),
        ("- severe line crowding count:", "severe_line_crowding_count"),
    ):
        report_count = _mechanical_cad_int(report.get(report_key))
        record_count = _mechanical_cad_int(values[prefix])
        if report_count is None:
            issues.append(f"mechanical CAD render readability report missing integer {report_key}")
        elif report_count != 0:
            issues.append(f"mechanical CAD render readability report {report_key} must be 0")
        if report_count is not None and record_count is not None and report_count != record_count:
            issues.append(f"mechanical CAD render readability record {prefix[2:]} differs from report {report_key}")
    pdf_rows = report.get("pdf_rows")
    png_rows = report.get("png_rows")
    if not isinstance(pdf_rows, list) or not isinstance(png_rows, list) or not pdf_rows or not png_rows:
        issues.append("mechanical CAD render readability report must include nonempty pdf_rows and png_rows")
    for row_family, rows in (("pdf_rows", pdf_rows), ("png_rows", png_rows)):
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                issues.append(f"mechanical CAD render readability {row_family} rows must be objects")
                continue
            if row.get("passed") is not True:
                issues.append(f"mechanical CAD render readability {row_family} row must pass")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", str(row.get("sha256") or "")):
                issues.append(f"mechanical CAD render readability {row_family} row must include sha256")
    return issues


def _mechanical_cad_validate_linework_and_color_fields(
    *,
    record_path: Path,
    values: dict[str, str],
    raw_values: dict[str, str],
    expected_package_sha: str,
) -> list[str]:
    issues: list[str] = []
    for prefix in (
        "- mechanical drawing lineweight/linetype fidelity audit verdict:",
        "- mechanical drawing lineweight/linetype fidelity audit path:",
        "- mechanical drawing lineweight/linetype fidelity package sha256:",
        "- mechanical drawing source line family coverage verdict:",
        "- mechanical drawing color-family audit path:",
        "- mechanical drawing color-family audit verdict:",
        "- mechanical drawing color-family package/path sha256:",
        "- mechanical drawing thin-solid-white-only verdict:",
        "- mechanical drawing non-thin white layer/entity count:",
        "- mechanical drawing expected color family count:",
        "- mechanical drawing entity color override count:",
        "- mechanical drawing thick solid family count:",
        "- mechanical drawing thin solid family count:",
        "- mechanical drawing center dash-dot family count:",
        "- mechanical drawing hidden dashed family count:",
        "- mechanical drawing section hatch family count:",
    ):
        value = values.get(prefix, "")
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"mechanical CAD linework/color evidence must bind {prefix[2:-1]}")
    if issues:
        return issues

    if not _mechanical_cad_passish(values["- mechanical drawing lineweight/linetype fidelity audit verdict:"]):
        issues.append("mechanical CAD lineweight/linetype fidelity verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing source line family coverage verdict:"]):
        issues.append("mechanical CAD source line family coverage verdict must be pass-shaped")
    linework_sha = normalize(values["- mechanical drawing lineweight/linetype fidelity package sha256:"]).upper()
    if not re.fullmatch(r"[0-9A-F]{64}", linework_sha):
        issues.append("mechanical CAD lineweight/linetype fidelity package sha256 must be 64-hex")
    elif expected_package_sha and linework_sha != expected_package_sha.upper():
        issues.append("mechanical CAD lineweight/linetype fidelity package sha256 must match final CAD package")
    for prefix in (
        "- mechanical drawing thick solid family count:",
        "- mechanical drawing thin solid family count:",
        "- mechanical drawing center dash-dot family count:",
        "- mechanical drawing hidden dashed family count:",
        "- mechanical drawing section hatch family count:",
    ):
        count = _mechanical_cad_int(values[prefix])
        if count is None or count <= 0:
            issues.append(f"mechanical CAD {prefix[2:]} must be numeric > 0")

    _, linework_report, linework_path_issues = _mechanical_cad_resolved_json(
        record_path=record_path,
        raw_values=raw_values,
        prefix="- mechanical drawing lineweight/linetype fidelity audit path:",
        label="mechanical CAD lineweight/linetype fidelity report",
    )
    issues.extend(linework_path_issues)
    if linework_path_issues:
        return issues
    if linework_report.get("schema") != "graduation-project-builder.cad-dxf-linework-fidelity-audit.v1":
        issues.append("mechanical CAD lineweight/linetype fidelity report must use cad-dxf-linework-fidelity schema v1")
    if linework_report.get("passed") is not True:
        issues.append("mechanical CAD lineweight/linetype fidelity report must have passed=true")
    if linework_report.get("issues") not in ([], None):
        issues.append("mechanical CAD lineweight/linetype fidelity report must not carry unresolved issues")
    report_linework_sha = normalize(
        str(linework_report.get("package_sha256") or linework_report.get("path_sha256") or "")
    ).upper()
    if report_linework_sha and report_linework_sha != linework_sha:
        issues.append("mechanical CAD lineweight/linetype package sha256 differs from linework fidelity report")
    if expected_package_sha and report_linework_sha and report_linework_sha != expected_package_sha.upper():
        issues.append("mechanical CAD linework fidelity report package sha256 must match final CAD package")
    package_family_counts = linework_report.get("package_source_line_family_counts")
    if not isinstance(package_family_counts, dict):
        package_family_counts = {}
        per_file_counts = linework_report.get("per_file")
        if isinstance(per_file_counts, list):
            for row in per_file_counts:
                if not isinstance(row, dict):
                    continue
                counts = row.get("source_line_family_counts")
                if not isinstance(counts, dict):
                    continue
                for family, value in counts.items():
                    package_family_counts[str(family)] = (
                        (_mechanical_cad_int(package_family_counts.get(str(family))) or 0)
                        + (_mechanical_cad_int(value) or 0)
                    )
    linework_family_field_map = {
        "- mechanical drawing thick solid family count:": "thick_solid",
        "- mechanical drawing thin solid family count:": "thin_solid",
        "- mechanical drawing center dash-dot family count:": "center_dash_dot",
        "- mechanical drawing hidden dashed family count:": "hidden_dashed",
        "- mechanical drawing section hatch family count:": "section_hatch",
    }
    for prefix, family in linework_family_field_map.items():
        report_count = _mechanical_cad_int(package_family_counts.get(family))
        record_count = _mechanical_cad_int(values[prefix])
        if report_count is None or report_count <= 0:
            issues.append(f"mechanical CAD linework fidelity report must include positive {family} count")
        if report_count is not None and record_count is not None and report_count != record_count:
            issues.append(f"mechanical CAD acceptance {prefix[2:]} differs from linework fidelity report {family}")
    for family in ("dimension", "leader_or_annotation"):
        report_count = _mechanical_cad_int(package_family_counts.get(family))
        if report_count is None or report_count <= 0:
            issues.append(f"mechanical CAD linework fidelity report must include positive {family} count")
    linework_per_file = linework_report.get("per_file")
    if not isinstance(linework_per_file, list) or not linework_per_file:
        issues.append("mechanical CAD linework fidelity report must include per_file rows")
    else:
        for row in linework_per_file:
            if not isinstance(row, dict):
                issues.append("mechanical CAD linework fidelity per_file rows must be objects")
                continue
            if row.get("passed") is not True:
                issues.append("mechanical CAD linework fidelity per_file row must pass")
            headers = row.get("headers")
            if not isinstance(headers, dict) or "$LTSCALE" not in headers or "$PSLTSCALE" not in headers:
                issues.append("mechanical CAD linework fidelity per_file row must bind $LTSCALE and $PSLTSCALE headers")
            if (_mechanical_cad_int(row.get("missing_lineweight_count")) or 0) != 0:
                issues.append("mechanical CAD linework fidelity per_file row must have missing_lineweight_count=0")
            if (_mechanical_cad_int(row.get("center_hidden_linetype_issue_count")) or 0) != 0:
                issues.append("mechanical CAD linework fidelity per_file row must have center_hidden_linetype_issue_count=0")

    _, color_report, path_issues = _mechanical_cad_resolved_json(
        record_path=record_path,
        raw_values=raw_values,
        prefix="- mechanical drawing color-family audit path:",
        label="mechanical CAD color-family audit report",
    )
    issues.extend(path_issues)
    if path_issues:
        return issues
    if color_report.get("schema") != "graduation-project-builder.cad-dxf-color-family-standard.v1":
        issues.append("mechanical CAD color-family audit report must use cad-dxf-color-family-standard schema v1")
    if color_report.get("passed") is not True:
        issues.append("mechanical CAD color-family audit report must have passed=true")
    if color_report.get("issues") not in ([], None):
        issues.append("mechanical CAD color-family audit report must not carry unresolved issues")
    color_sha = normalize(
        str(color_report.get("package_sha256") or color_report.get("path_sha256") or "")
    ).upper()
    record_color_sha = normalize(values["- mechanical drawing color-family package/path sha256:"]).upper()
    if not re.fullmatch(r"[0-9A-F]{64}", record_color_sha):
        issues.append("mechanical CAD color-family package/path sha256 must be 64-hex")
    elif color_sha and record_color_sha != color_sha:
        issues.append("mechanical CAD color-family package/path sha256 differs from color audit report")
    if expected_package_sha and record_color_sha != expected_package_sha.upper():
        issues.append("mechanical CAD color-family package/path sha256 must match final CAD package")
    if not _mechanical_cad_passish(values["- mechanical drawing color-family audit verdict:"]):
        issues.append("mechanical CAD color-family audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing thin-solid-white-only verdict:"]):
        issues.append("mechanical CAD thin-solid-white-only verdict must be pass-shaped")
    non_thin_white_count = _mechanical_cad_int(values["- mechanical drawing non-thin white layer/entity count:"])
    if non_thin_white_count is None or non_thin_white_count != 0:
        issues.append("mechanical CAD non-thin white layer/entity count must be numeric 0")
    override_count = _mechanical_cad_int(values["- mechanical drawing entity color override count:"])
    if override_count is None or override_count != 0:
        issues.append("mechanical CAD entity color override count must be numeric 0")
    expected_family_count = _mechanical_cad_int(values["- mechanical drawing expected color family count:"])
    palette = color_report.get("family_palette")
    if not isinstance(palette, dict) or len(palette) < 6:
        issues.append("mechanical CAD color-family audit report must include at least six family_palette entries")
    elif expected_family_count != len(palette):
        issues.append("mechanical CAD expected color family count field must match color audit family_palette")
    rule_summary = normalize(str(color_report.get("rule_summary") or "")).lower()
    if "thin_solid_white_only" not in rule_summary or "non-white" not in rule_summary:
        issues.append("mechanical CAD color-family audit rule_summary must lock thin-solid white and non-thin non-white")
    per_file = color_report.get("per_file")
    if not isinstance(per_file, list) or not per_file:
        issues.append("mechanical CAD color-family audit report must include per_file rows")
        return issues
    aggregate_non_thin_white = 0
    aggregate_override = 0
    non_thin_white_layers = 0
    for row in per_file:
        if not isinstance(row, dict):
            issues.append("mechanical CAD color-family audit per_file rows must be objects")
            continue
        if row.get("passed") is not True:
            issues.append("mechanical CAD color-family audit per_file row must pass")
        aggregate_non_thin_white += _mechanical_cad_int(row.get("non_thin_white_entity_override_count")) or 0
        aggregate_override += _mechanical_cad_int(row.get("non_bylayer_entity_color_override_count")) or 0
        for layer in row.get("layers") or []:
            if not isinstance(layer, dict):
                continue
            rgb = layer.get("rgb")
            is_white = layer.get("aci") == 7
            if isinstance(rgb, list) and len(rgb) >= 3:
                try:
                    is_white = is_white or sum((float(rgb[index]) - 255.0) ** 2 for index in range(3)) ** 0.5 <= 15.0
                except (TypeError, ValueError):
                    pass
            if is_white and str(layer.get("family") or "") != "thin_solid" and str(layer.get("name") or "") != "Defpoints":
                non_thin_white_layers += 1
    if aggregate_non_thin_white != 0 or non_thin_white_layers != 0:
        issues.append("mechanical CAD color-family audit must have zero non-thin white entities/layers")
    if aggregate_override != 0:
        issues.append("mechanical CAD color-family audit must have zero entity color overrides")
    if non_thin_white_count == 0 and (aggregate_non_thin_white + non_thin_white_layers) != non_thin_white_count:
        issues.append("mechanical CAD non-thin white layer/entity field differs from color audit report")
    if override_count == 0 and aggregate_override != override_count:
        issues.append("mechanical CAD entity color override count field differs from color audit report")
    return issues


MECHANICAL_CAD_FINAL_PACKAGE_SURFACE_PREFIXES = (
    "- mechanical drawing rendered no-overlap verdict:",
    "- mechanical drawing boundary clearance verdict:",
    "- mechanical drawing detail density verdict:",
    "- mechanical drawing title block/table/notes isolation verdict:",
    "- mechanical drawing title-block cell containment verdict:",
    "- mechanical drawing title-block short-line topology audit verdict:",
    "- mechanical drawing missing short table line count:",
    "- mechanical drawing broken cell-border count:",
    "- mechanical drawing table-grid topology mismatch count:",
    "- mechanical drawing diagnostic-overlay-free title-block crop path:",
    "- mechanical drawing annotation margin clearance verdict:",
    "- mechanical drawing local crowding verdict:",
    "- mechanical drawing text/table/frame overlap verdict:",
    "- mechanical drawing machine overlap audit verdict:",
    "- mechanical drawing content-overlap audit verdict:",
    "- mechanical drawing content-overlap evidence path:",
    "- mechanical drawing registered content bbox count:",
    "- mechanical drawing checked content pair count:",
    "- mechanical drawing content overlap count:",
    "- mechanical drawing view-view overlap count:",
    "- mechanical drawing detail-frame-main-view overlap count:",
    "- mechanical drawing table text/grid collision count:",
    "- mechanical drawing table_text_grid_collision_count:",
    "- mechanical drawing dimension-line view/table crossing count:",
    "- mechanical drawing leader-line view/table crossing count:",
    "- mechanical drawing balloon geometry collision count:",
    "- mechanical drawing bbox helper envelope escape count:",
    "- mechanical drawing stale rendered preview count:",
    "- mechanical drawing outside-frame ink audit verdict:",
    "- mechanical drawing outside-frame ink audit evidence path:",
    "- mechanical drawing outside-frame independent ink component count:",
    "- mechanical drawing outside-frame text component count:",
    "- mechanical drawing outside-frame leader component count:",
    "- mechanical drawing outside-frame hatch/section component count:",
    "- mechanical drawing outside-frame table/title-block component count:",
    "- mechanical drawing annotation ownership audit verdict:",
    "- mechanical drawing unowned free text count:",
    "- mechanical drawing unsupported floating text count:",
    "- mechanical drawing reserved-zone intrusion audit verdict:",
    "- mechanical drawing protected-table-zone intrusion audit verdict:",
    "- mechanical drawing reserved-zone intrusion count:",
    "- mechanical drawing protected-table-zone intrusion count:",
    "- mechanical drawing dimension-line table-zone intrusion count:",
    "- mechanical drawing dimension-text table-zone intrusion count:",
    "- mechanical drawing view-geometry table-zone intrusion count:",
    "- mechanical drawing detail-view table-zone intrusion count:",
    "- mechanical drawing leader/balloon table-zone intrusion count:",
    "- mechanical drawing dimension table-zone intrusion count:",
    "- mechanical drawing title-block/BOM protected-zone intrusion count:",
    "- mechanical drawing view-geometry reserved-zone intrusion count:",
    "- mechanical drawing hatch/section fill clipping audit verdict:",
    "- mechanical drawing hatch/section fill clipping evidence path:",
    "- mechanical drawing hatch/section fill boundary violation count:",
    "- mechanical drawing hatch/section fill adjacent-view crossing count:",
    "- mechanical drawing hatch/section fill dimension-line crossing count:",
    "- mechanical drawing hatch/section fill title-block/table/BOM/frame crossing count:",
    "- mechanical drawing hatch/section fill blank-background leak count:",
    "- mechanical drawing text entity overlap count:",
    "- mechanical drawing text legibility machine audit verdict:",
    "- mechanical drawing minimum CAD text height mm:",
    "- mechanical drawing minimum rendered text height px:",
    "- mechanical drawing text integrity audit path:",
    "- mechanical drawing text integrity audit verdict:",
    "- mechanical drawing mojibake/tofu/missing glyph count:",
    "- mechanical drawing missing required drawing text count:",
    "- mechanical drawing text orientation audit path:",
    "- mechanical drawing text orientation audit verdict:",
    "- mechanical drawing upside-down text count:",
    "- mechanical drawing mirrored text count:",
    "- mechanical drawing manufacturing complexity audit verdict:",
    "- mechanical drawing minimum manufacturing detail family count:",
    "- mechanical drawing entity-count-only false-pass verdict:",
)


def _mechanical_cad_require_final_surface_fields(values: dict[str, str]) -> list[str]:
    issues: list[str] = []
    for prefix in MECHANICAL_CAD_FINAL_PACKAGE_SURFACE_PREFIXES:
        value = values.get(prefix, "")
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"mechanical CAD final acceptance must bind {prefix[2:-1]}")
    return issues


def _mechanical_cad_record_count(values: dict[str, str], prefix: str) -> int | None:
    value = values.get(prefix, "")
    if not is_explicit(value) or is_explicit_none(value):
        return None
    return _mechanical_cad_int(value)


def _mechanical_cad_check_section(
    *,
    report: dict[str, object],
    key: str,
    label: str,
    issues: list[str],
) -> dict[str, object]:
    section = report.get(key)
    if not isinstance(section, dict) or section.get("passed") is not True:
        issues.append(f"mechanical CAD package audit must include passing {label}")
        return {}
    return section


def _mechanical_cad_check_zero_counts(
    *,
    section: dict[str, object],
    label: str,
    report_fields: tuple[str, ...],
    issues: list[str],
    values: dict[str, str] | None = None,
    field_prefix_map: dict[str, str] | None = None,
) -> None:
    for report_field in report_fields:
        report_count = _mechanical_cad_int(section.get(report_field))
        if report_count is None or report_count != 0:
            issues.append(f"mechanical CAD package audit {label}.{report_field} must be 0")
        if values is not None and field_prefix_map is not None:
            prefix = field_prefix_map.get(report_field)
            if prefix:
                record_count = _mechanical_cad_record_count(values, prefix)
                if record_count is not None and report_count is not None and record_count != report_count:
                    issues.append(f"mechanical CAD final acceptance {prefix[2:]} differs from {label}.{report_field}")


def _mechanical_cad_validate_rendered_review_payload(
    *,
    review: dict[str, object],
    label: str,
    values: dict[str, str],
) -> list[str]:
    issues: list[str] = []
    if not isinstance(review, dict) or review.get("passed") is not True:
        return [f"mechanical CAD package audit must include passing {label}"]

    machine_overlap = _mechanical_cad_check_section(
        report=review,
        key="machine_overlap_audit",
        label=f"{label}.machine_overlap_audit",
        issues=issues,
    )
    if machine_overlap:
        _mechanical_cad_check_zero_counts(
            section=machine_overlap,
            label=f"{label}.machine_overlap_audit",
            report_fields=(
                "overlap_count",
                "text_entity_overlap_count",
                "reserved_zone_collision_count",
                "title_block_table_note_collision_count",
                "annotation_collision_count",
                "frame_clearance_violation_count",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "text_entity_overlap_count": "- mechanical drawing text entity overlap count:",
            },
        )

    content_overlap = _mechanical_cad_check_section(
        report=review,
        key="content_overlap_audit",
        label=f"{label}.content_overlap_audit",
        issues=issues,
    )
    if content_overlap:
        for report_field, prefix in (
            ("registered_bbox_count", "- mechanical drawing registered content bbox count:"),
            ("checked_pair_count", "- mechanical drawing checked content pair count:"),
        ):
            report_count = _mechanical_cad_int(content_overlap.get(report_field))
            record_count = _mechanical_cad_record_count(values, prefix)
            if report_count is None or report_count <= 0:
                issues.append(f"mechanical CAD package audit {label}.content_overlap_audit.{report_field} must be > 0")
            if record_count is not None and report_count is not None and record_count != report_count:
                issues.append(f"mechanical CAD final acceptance {prefix[2:]} differs from {label}.content_overlap_audit.{report_field}")
        _mechanical_cad_check_zero_counts(
            section=content_overlap,
            label=f"{label}.content_overlap_audit",
            report_fields=(
                "content_overlap_count",
                "view_view_overlap_count",
                "detail_frame_main_view_overlap_count",
                "table_text_grid_collision_count",
                "dimension_line_view_table_crossing_count",
                "leader_line_view_table_crossing_count",
                "balloon_geometry_collision_count",
                "bbox_helper_envelope_escape_count",
                "stale_rendered_preview_count",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "content_overlap_count": "- mechanical drawing content overlap count:",
                "view_view_overlap_count": "- mechanical drawing view-view overlap count:",
                "detail_frame_main_view_overlap_count": "- mechanical drawing detail-frame-main-view overlap count:",
                "table_text_grid_collision_count": "- mechanical drawing table text/grid collision count:",
                "dimension_line_view_table_crossing_count": "- mechanical drawing dimension-line view/table crossing count:",
                "leader_line_view_table_crossing_count": "- mechanical drawing leader-line view/table crossing count:",
                "balloon_geometry_collision_count": "- mechanical drawing balloon geometry collision count:",
                "bbox_helper_envelope_escape_count": "- mechanical drawing bbox helper envelope escape count:",
                "stale_rendered_preview_count": "- mechanical drawing stale rendered preview count:",
            },
        )

    title_topology = _mechanical_cad_check_section(
        report=review,
        key="title_block_short_line_topology_audit",
        label=f"{label}.title_block_short_line_topology_audit",
        issues=issues,
    )
    if title_topology:
        _mechanical_cad_check_zero_counts(
            section=title_topology,
            label=f"{label}.title_block_short_line_topology_audit",
            report_fields=(
                "missing_short_table_line_count",
                "broken_cell_border_count",
                "table_grid_topology_mismatch_count",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "missing_short_table_line_count": "- mechanical drawing missing short table line count:",
                "broken_cell_border_count": "- mechanical drawing broken cell-border count:",
                "table_grid_topology_mismatch_count": "- mechanical drawing table-grid topology mismatch count:",
            },
        )

    outside_frame = _mechanical_cad_check_section(
        report=review,
        key="outside_frame_ink_audit",
        label=f"{label}.outside_frame_ink_audit",
        issues=issues,
    )
    if outside_frame:
        _mechanical_cad_check_zero_counts(
            section=outside_frame,
            label=f"{label}.outside_frame_ink_audit",
            report_fields=(
                "outside_frame_independent_ink_component_count",
                "outside_frame_text_component_count",
                "outside_frame_leader_component_count",
                "outside_frame_hatch_section_component_count",
                "outside_frame_table_title_block_component_count",
                "max_outside_component_area_px",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "outside_frame_independent_ink_component_count": "- mechanical drawing outside-frame independent ink component count:",
                "outside_frame_text_component_count": "- mechanical drawing outside-frame text component count:",
                "outside_frame_leader_component_count": "- mechanical drawing outside-frame leader component count:",
                "outside_frame_hatch_section_component_count": "- mechanical drawing outside-frame hatch/section component count:",
                "outside_frame_table_title_block_component_count": "- mechanical drawing outside-frame table/title-block component count:",
            },
        )

    cell_containment = _mechanical_cad_check_section(
        report=review,
        key="cell_containment_audit",
        label=f"{label}.cell_containment_audit",
        issues=issues,
    )
    if cell_containment:
        _mechanical_cad_check_zero_counts(
            section=cell_containment,
            label=f"{label}.cell_containment_audit",
            report_fields=("outside_cell_count", "border_touch_count", "unowned_table_text_count", "clipped_overflow_count"),
            issues=issues,
        )

    annotation_ownership = _mechanical_cad_check_section(
        report=review,
        key="annotation_ownership_audit",
        label=f"{label}.annotation_ownership_audit",
        issues=issues,
    )
    if annotation_ownership:
        _mechanical_cad_check_zero_counts(
            section=annotation_ownership,
            label=f"{label}.annotation_ownership_audit",
            report_fields=("unowned_free_text_count", "unsupported_floating_text_count"),
            issues=issues,
            values=values,
            field_prefix_map={
                "unowned_free_text_count": "- mechanical drawing unowned free text count:",
                "unsupported_floating_text_count": "- mechanical drawing unsupported floating text count:",
            },
        )

    reserved_zone = _mechanical_cad_check_section(
        report=review,
        key="reserved_zone_intrusion_audit",
        label=f"{label}.reserved_zone_intrusion_audit",
        issues=issues,
    )
    if reserved_zone:
        _mechanical_cad_check_zero_counts(
            section=reserved_zone,
            label=f"{label}.reserved_zone_intrusion_audit",
            report_fields=(
                "intrusion_count",
                "geometry_title_block_intrusion_count",
                "dimension_line_table_zone_intrusion_count",
                "dimension_text_table_zone_intrusion_count",
                "leader_table_zone_intrusion_count",
                "view_geometry_reserved_zone_intrusion_count",
                "bom_table_intrusion_count",
                "technical_requirement_intrusion_count",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "intrusion_count": "- mechanical drawing reserved-zone intrusion count:",
                "dimension_line_table_zone_intrusion_count": "- mechanical drawing dimension-line table-zone intrusion count:",
                "dimension_text_table_zone_intrusion_count": "- mechanical drawing dimension-text table-zone intrusion count:",
                "view_geometry_reserved_zone_intrusion_count": "- mechanical drawing view-geometry reserved-zone intrusion count:",
            },
        )

    hatch_clip = _mechanical_cad_check_section(
        report=review,
        key="hatch_clip_audit",
        label=f"{label}.hatch_clip_audit",
        issues=issues,
    )
    if hatch_clip:
        _mechanical_cad_check_zero_counts(
            section=hatch_clip,
            label=f"{label}.hatch_clip_audit",
            report_fields=(
                "hatch_clip_violation_count",
                "entity_boundary_escape_count",
                "adjacent_view_crossing_count",
                "dimension_line_crossing_count",
                "title_block_table_bom_frame_crossing_count",
                "blank_background_leak_count",
            ),
            issues=issues,
            values=values,
            field_prefix_map={
                "hatch_clip_violation_count": "- mechanical drawing hatch/section fill boundary violation count:",
                "adjacent_view_crossing_count": "- mechanical drawing hatch/section fill adjacent-view crossing count:",
                "dimension_line_crossing_count": "- mechanical drawing hatch/section fill dimension-line crossing count:",
                "title_block_table_bom_frame_crossing_count": "- mechanical drawing hatch/section fill title-block/table/BOM/frame crossing count:",
                "blank_background_leak_count": "- mechanical drawing hatch/section fill blank-background leak count:",
            },
        )

    text_legibility = _mechanical_cad_check_section(
        report=review,
        key="text_legibility_audit",
        label=f"{label}.text_legibility_audit",
        issues=issues,
    )
    if text_legibility and (
        (_mechanical_cad_float(text_legibility.get("min_cad_text_height_mm")) or 0.0)
        < MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM
        or (_mechanical_cad_float(text_legibility.get("min_rendered_text_height_px")) or 0.0)
        < MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX
    ):
        issues.append(f"mechanical CAD package audit {label}.text_legibility_audit is below required thresholds")

    manufacturing_complexity = _mechanical_cad_check_section(
        report=review,
        key="manufacturing_complexity_audit",
        label=f"{label}.manufacturing_complexity_audit",
        issues=issues,
    )
    if manufacturing_complexity and (
        (_mechanical_cad_int(manufacturing_complexity.get("min_feature_family_count")) or 0)
        < MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES
    ):
        issues.append(f"mechanical CAD package audit {label}.manufacturing_complexity_audit is below required family count")

    verdict_map = {
        "- mechanical drawing rendered no-overlap verdict:": "no_overlap_verdict",
        "- mechanical drawing boundary clearance verdict:": "boundary_clearance_verdict",
        "- mechanical drawing detail density verdict:": "detail_density_verdict",
        "- mechanical drawing title block/table/notes isolation verdict:": "title_block_table_notes_isolation_verdict",
        "- mechanical drawing title-block cell containment verdict:": "title_block_cell_containment_verdict",
        "- mechanical drawing title-block short-line topology audit verdict:": "title_block_short_line_topology_verdict",
        "- mechanical drawing annotation margin clearance verdict:": "annotation_margin_clearance_verdict",
        "- mechanical drawing local crowding verdict:": "local_crowding_verdict",
        "- mechanical drawing text/table/frame overlap verdict:": "no_overlap_verdict",
        "- mechanical drawing content-overlap audit verdict:": "content_overlap_verdict",
        "- mechanical drawing outside-frame ink audit verdict:": "outside_frame_ink_verdict",
        "- mechanical drawing hatch/section fill clipping audit verdict:": "hatch_section_fill_clipping_verdict",
    }
    for prefix, report_field in verdict_map.items():
        record_value = normalize(values.get(prefix, "")).lower()
        report_value = normalize(str(review.get(report_field, ""))).lower()
        if prefix in {
            "- mechanical drawing rendered no-overlap verdict:",
            "- mechanical drawing text/table/frame overlap verdict:",
        }:
            if report_value and report_value not in MECHANICAL_CAD_NO_OVERLAP_VALUES:
                issues.append(f"mechanical CAD package audit {label}.{report_field} must be pass/no-overlap")
        elif report_value and report_value not in MECHANICAL_CAD_PASS_VALUES:
            issues.append(f"mechanical CAD package audit {label}.{report_field} must be pass-shaped")
        if record_value and report_value and record_value != report_value:
            issues.append(f"mechanical CAD final acceptance {prefix[2:]} differs from {label}.{report_field}")

    return issues


def _mechanical_cad_validate_final_package_audit_report(
    *,
    package_report: dict[str, object],
    values: dict[str, str],
    expected_package_sha: str,
) -> list[str]:
    issues: list[str] = []
    if package_report.get("schema") not in MECHANICAL_CAD_AUDIT_SCHEMA_VALUES:
        issues.append("mechanical CAD package audit must use v4-or-stricter schema")
    if package_report.get("passed") is not True:
        issues.append("mechanical CAD package audit must have passed=true")
    if package_report.get("issues") not in ([], None):
        issues.append("mechanical CAD package audit must not carry unresolved issues")
    package_sha = normalize(str(package_report.get("package_sha256") or package_report.get("sha256") or "")).upper()
    expected_sha = normalize(expected_package_sha).upper()
    if expected_sha and re.fullmatch(r"[0-9A-F]{64}", expected_sha) and package_sha and package_sha != expected_sha:
        issues.append("mechanical CAD package audit SHA256 must match exact final CAD package SHA256")

    frame_overflow = package_report.get("rendered_frame_overflow_verdict")
    if not isinstance(frame_overflow, dict) or frame_overflow.get("passed") is not True:
        issues.append("mechanical CAD package audit must include passing rendered_frame_overflow_verdict")
    else:
        _mechanical_cad_check_zero_counts(
            section=frame_overflow,
            label="rendered_frame_overflow_verdict",
            report_fields=("outside_frame_component_count", "outside_frame_pixel_count", "max_outside_component_area_px"),
            issues=issues,
        )

    rendered_review = package_report.get("rendered_review_verdict")
    issues.extend(
        _mechanical_cad_validate_rendered_review_payload(
            review=rendered_review if isinstance(rendered_review, dict) else {},
            label="rendered_review_verdict",
            values=values,
        )
    )

    candidate = package_report.get("candidate")
    candidate_review = candidate.get("rendered_review") if isinstance(candidate, dict) else None
    issues.extend(
        _mechanical_cad_validate_rendered_review_payload(
            review=candidate_review if isinstance(candidate_review, dict) else {},
            label="candidate.rendered_review",
            values=values,
        )
    )
    candidate_frame = candidate.get("rendered_frame_overflow") if isinstance(candidate, dict) else None
    if isinstance(candidate_frame, dict):
        if candidate_frame.get("passed") is not True:
            issues.append("mechanical CAD package audit candidate.rendered_frame_overflow must pass")
        else:
            _mechanical_cad_check_zero_counts(
                section=candidate_frame,
                label="candidate.rendered_frame_overflow",
                report_fields=("outside_frame_component_count", "outside_frame_pixel_count", "max_outside_component_area_px"),
                issues=issues,
            )
    else:
        issues.append("mechanical CAD package audit must include candidate.rendered_frame_overflow")

    entity_only_value = normalize(values.get("- mechanical drawing entity-count-only false-pass verdict:", "")).lower()
    if entity_only_value and entity_only_value not in MECHANICAL_CAD_ENTITY_ONLY_REJECT_VALUES:
        issues.append("mechanical CAD final acceptance must reject entity-count-only false-pass")
    return issues


def check_mechanical_cad_acceptance_record(record_path: Path, record_lines: list[str]) -> list[str]:
    issues: list[str] = []
    normalized_lines = {normalize(line) for line in record_lines if normalize(line)}
    for heading in MECHANICAL_CAD_ACCEPTANCE_HEADINGS:
        if heading not in normalized_lines:
            issues.append(f"mechanical CAD acceptance record missing marker: {heading}")

    values, raw_values, field_issues = _mechanical_cad_field_maps(record_lines)
    issues.extend(field_issues)
    if field_issues:
        return issues

    if values["- acceptance schema:"] != MECHANICAL_CAD_ACCEPTANCE_SCHEMA_VALUE:
        issues.append(
            "mechanical CAD acceptance schema must be "
            f"{MECHANICAL_CAD_ACCEPTANCE_SCHEMA_VALUE}: {values['- acceptance schema:'] or 'missing'}"
        )
    if not is_explicit(values["- task mode:"]) or is_explicit_none(values["- task mode:"]):
        issues.append("mechanical CAD acceptance task mode must be explicit")
    if not is_explicit(values["- subtask:"]) or is_explicit_none(values["- subtask:"]):
        issues.append("mechanical CAD acceptance subtask must be explicit")

    final_zip, final_zip_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact final delivery package path:"],
        label="final delivery package",
        required_suffixes={".zip"},
    )
    audited_package, audited_package_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact audited CAD package path:"],
        label="audited CAD package",
        required_suffixes={".zip"},
    )
    dwg_zip, dwg_zip_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact DWG package path:"],
        label="DWG package",
        required_suffixes={".zip"},
    )
    dxf_zip, dxf_zip_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact DXF package path:"],
        label="DXF package",
        required_suffixes={".zip"},
    )
    combined_pdf, combined_pdf_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact combined PDF path:"],
        label="combined PDF",
        required_suffixes={".pdf"},
    )
    png_zip, png_zip_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- exact PNG render package path:"],
        label="PNG render package",
        required_suffixes={".zip"},
    )
    regeneration_manifest, regeneration_manifest_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- drawing regeneration manifest path:"],
        label="drawing regeneration manifest",
        required_suffixes={".json", ".md"},
    )
    audit_report_path, audit_report_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- mechanical drawing package audit path:"],
        label="audit report",
        required_suffixes={".json"},
    )
    source_delta_report_path, source_delta_report_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- mechanical drawing source linework delta audit path:"],
        label="source linework delta audit report",
        required_suffixes={".json"},
    )
    render_readability_report_path, render_readability_report_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- mechanical render readability audit path:"],
        label="render readability audit report",
        required_suffixes={".json"},
    )
    title_block_crop_path, title_block_crop_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- mechanical drawing diagnostic-overlay-free title-block crop path:"],
        label="diagnostic-overlay-free title-block crop",
        required_suffixes=IMAGE_EXTENSIONS | PDF_EXTENSIONS,
    )
    font_audit_path, font_audit_issues = _mechanical_cad_resolve_file(
        record_path=record_path,
        raw_value=raw_values["- CAD text style/font audit path:"],
        label="CAD text style/font audit report",
        required_suffixes={".json"},
    )
    issues.extend(final_zip_issues)
    issues.extend(audited_package_issues)
    issues.extend(dwg_zip_issues)
    issues.extend(dxf_zip_issues)
    issues.extend(combined_pdf_issues)
    issues.extend(png_zip_issues)
    issues.extend(regeneration_manifest_issues)
    issues.extend(audit_report_issues)
    issues.extend(source_delta_report_issues)
    issues.extend(render_readability_report_issues)
    issues.extend(title_block_crop_issues)
    issues.extend(font_audit_issues)
    issues.extend(
        _mechanical_cad_validate_sha(
            final_zip,
            values["- exact final delivery package sha256:"],
            "final delivery package",
        )
    )
    issues.extend(
        _mechanical_cad_validate_sha(
            audited_package,
            values["- exact audited CAD package sha256:"],
            "audited CAD package",
        )
    )
    final_source_sha = normalize(values["- mechanical drawing final CAD package sha256:"]).upper()
    source_baseline_sha = normalize(values["- mechanical drawing source CAD package sha256:"]).upper()
    if not re.fullmatch(r"[0-9A-F]{64}", final_source_sha):
        issues.append("mechanical CAD acceptance final CAD package SHA256 must be a 64-hex value")
    elif final_source_sha != values["- exact audited CAD package sha256:"].upper():
        issues.append("mechanical CAD acceptance final CAD package SHA256 must match exact audited CAD package SHA256")
    if not re.fullmatch(r"[0-9A-F]{64}", source_baseline_sha):
        issues.append("mechanical CAD acceptance source CAD package SHA256 must be a 64-hex value")
    issues.extend(
        _mechanical_cad_validate_sha(
            dwg_zip,
            values["- exact DWG package sha256:"],
            "DWG package",
        )
    )
    issues.extend(
        _mechanical_cad_validate_sha(
            dxf_zip,
            values["- exact DXF package sha256:"],
            "DXF package",
        )
    )
    issues.extend(
        _mechanical_cad_validate_sha(
            combined_pdf,
            values["- exact combined PDF sha256:"],
            "combined PDF",
        )
    )
    issues.extend(
        _mechanical_cad_validate_sha(
            png_zip,
            values["- exact PNG render package sha256:"],
            "PNG render package",
        )
    )
    issues.extend(
        _mechanical_cad_validate_sha(
            regeneration_manifest,
            values["- drawing regeneration manifest sha256:"],
            "drawing regeneration manifest",
        )
    )
    issues.extend(
        _mechanical_cad_validate_render_readability_fields(
            record_path=record_path,
            values=values,
            raw_values=raw_values,
        )
    )
    issues.extend(
        _mechanical_cad_validate_linework_and_color_fields(
            record_path=record_path,
            values=values,
            raw_values=raw_values,
            expected_package_sha=values["- exact audited CAD package sha256:"],
        )
    )

    if not _mechanical_cad_passish(values["- mechanical drawing package audit verdict:"]):
        issues.append("mechanical CAD acceptance audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing formal CAD source provenance verdict:"]):
        issues.append("mechanical CAD acceptance formal CAD source provenance verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing schematic/concept substitute rejection verdict:"]):
        issues.append("mechanical CAD acceptance schematic/concept substitute rejection verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- CAD official command route verdict:"]):
        issues.append("mechanical CAD acceptance CAD official command route verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- non-CAD fallback rejection verdict:"]):
        issues.append("mechanical CAD acceptance non-CAD fallback rejection verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- no verbatim geometry copying verdict:"]):
        issues.append("mechanical CAD acceptance no verbatim geometry copying verdict must be pass-shaped")
    reference_baseline_paths = split_path_values(raw_values["- mechanical drawing reference baseline path:"])
    if not reference_baseline_paths:
        issues.append("mechanical CAD acceptance must bind at least one teacher/user reference baseline path")
    for raw_path in reference_baseline_paths:
        baseline_path = resolve_record_path(raw_path, record_path)
        if baseline_path.exists() and baseline_path.is_dir():
            issues.extend(validate_existing_path(baseline_path, require_nonempty_file=False))
            continue
        issues.extend(validate_existing_path(baseline_path, require_nonempty_file=True))
        if baseline_path.suffix.lower() not in IMAGE_EXTENSIONS | PDF_EXTENSIONS | {".dwg", ".dxf", ".zip", ".json"}:
            issues.append(
                "mechanical CAD reference baseline must be a CAD/PDF/image/ZIP/JSON file or directory: "
                f"{baseline_path}"
            )
    closeup_evidence_paths = split_path_values(
        raw_values["- mechanical drawing CAD open-view close-up evidence path:"]
    )
    if not closeup_evidence_paths:
        issues.append("mechanical CAD acceptance must bind at least one CAD-open close-up evidence path")
    for raw_path in closeup_evidence_paths:
        closeup_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(closeup_path, require_nonempty_file=True))
        if closeup_path.suffix.lower() not in IMAGE_EXTENSIONS | PDF_EXTENSIONS | {".json"}:
            issues.append(f"mechanical CAD open-view close-up evidence must be image/PDF/JSON: {closeup_path}")
    for key in (
        "- mechanical drawing CAD open-view structural coherence verdict:",
        "- mechanical drawing complete assembly/object recognizability verdict:",
        "- mechanical drawing scattered-parts rejection verdict:",
        "- mechanical drawing reference-view trace alignment verdict:",
        "- mechanical drawing annotation/leader/title-block completeness verdict:",
    ):
        if not _mechanical_cad_passish(values[key]):
            issues.append(f"mechanical CAD acceptance {key[2:]} must be pass-shaped")
    annotation_checklist_paths = split_path_values(
        raw_values["- mechanical drawing external-case annotation checklist path:"]
    )
    if not annotation_checklist_paths:
        issues.append("mechanical CAD acceptance must bind external/reference annotation checklist path")
    for raw_path in annotation_checklist_paths:
        checklist_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(checklist_path, require_nonempty_file=True))
        if checklist_path.suffix.lower() not in {".json", ".md", ".txt", ".csv", ".xlsx", ".pdf"}:
            issues.append(
                "mechanical CAD external/reference annotation checklist must be JSON/MD/TXT/CSV/XLSX/PDF: "
                f"{checklist_path}"
            )
    if not is_explicit(values["- CAD executable or COM ProgID evidence:"]) or is_explicit_none(
        values["- CAD executable or COM ProgID evidence:"]
    ):
        issues.append("mechanical CAD acceptance must name the CAD executable path or COM ProgID evidence")
    cad_command_log_paths = split_path_values(raw_values["- CAD official command test log:"])
    if not cad_command_log_paths:
        issues.append("mechanical CAD acceptance must bind at least one CAD official command test log path")
    for raw_path in cad_command_log_paths:
        log_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(log_path, require_nonempty_file=True))
        if log_path.suffix.lower() not in {".log", ".txt", ".json"}:
            issues.append(f"mechanical CAD official command test log must be a log/txt/json file: {log_path}")
    external_ref_value = values["- external CAD case reference URL:"]
    if is_explicit(external_ref_value) and not is_explicit_none(external_ref_value):
        if not is_explicit(values["- reference-use restriction:"]) or is_explicit_none(
            values["- reference-use restriction:"]
        ):
            issues.append("mechanical CAD acceptance must record reference-use restriction for external CAD case references")
    if not _mechanical_cad_passish(values["- mechanical drawing source linework delta verdict:"]):
        issues.append("mechanical CAD acceptance source linework delta verdict must be pass-shaped")
    for key in (
        "- mechanical drawing changed source sheet count:",
        "- mechanical drawing changed source entity count:",
        "- mechanical drawing linework family delta count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count <= 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric > 0")
    for key in (
        "- mechanical drawing identical source sheet count:",
        "- mechanical drawing old-like large circle overlay count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    for key in (
        "- mechanical drawing PDF-only change rejection verdict:",
        "- mechanical drawing minor-entity-move-only rejection verdict:",
        "- mechanical drawing source-to-PDF derivation verdict:",
    ):
        if not _mechanical_cad_passish(values[key]):
            issues.append(f"mechanical CAD acceptance {key[2:]} must be pass-shaped")
    if normalize(values["- mechanical drawing rendered no-overlap verdict:"]).lower() not in MECHANICAL_CAD_NO_OVERLAP_VALUES:
        issues.append("mechanical CAD acceptance rendered no-overlap verdict must be pass/no-overlap")
    for key in (
        "- mechanical drawing boundary clearance verdict:",
        "- mechanical drawing detail density verdict:",
        "- mechanical drawing title block/table/notes isolation verdict:",
        "- mechanical drawing title-block cell containment verdict:",
        "- mechanical drawing title-block short-line topology audit verdict:",
        "- mechanical drawing annotation margin clearance verdict:",
        "- mechanical drawing local crowding verdict:",
    ):
        if not _mechanical_cad_passish(values[key]):
            issues.append(f"mechanical CAD acceptance {key[2:]} must be pass-shaped")
    if normalize(values["- mechanical drawing text/table/frame overlap verdict:"]).lower() not in MECHANICAL_CAD_NO_OVERLAP_VALUES:
        issues.append("mechanical CAD acceptance text/table/frame overlap verdict must be pass/no-overlap")
    if not _mechanical_cad_passish(values["- mechanical render readability verdict:"]):
        issues.append("mechanical CAD acceptance render readability verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical render readability source lineweight verdict:"]):
        issues.append("mechanical CAD acceptance render readability source lineweight verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical render readability render-only stroke cap:"]):
        issues.append("mechanical CAD acceptance render-only stroke cap field must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical render readability render-only lineweight scaling:"]):
        issues.append("mechanical CAD acceptance render-only lineweight scaling field must be pass-shaped")
    for key in (
        "- text text overlap count:",
        "- text graphic cover count:",
        "- severe line crowding count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    if not _mechanical_cad_passish(values["- mechanical drawing machine overlap audit verdict:"]):
        issues.append("mechanical CAD acceptance machine overlap audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing content-overlap audit verdict:"]):
        issues.append("mechanical CAD acceptance content-overlap audit verdict must be pass-shaped")
    for key in (
        "- mechanical drawing registered content bbox count:",
        "- mechanical drawing checked content pair count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count <= 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric > 0")
    for key in (
        "- mechanical drawing content overlap count:",
        "- mechanical drawing view-view overlap count:",
        "- mechanical drawing detail-frame-main-view overlap count:",
        "- mechanical drawing table text/grid collision count:",
        "- mechanical drawing table_text_grid_collision_count:",
        "- mechanical drawing dimension-line view/table crossing count:",
        "- mechanical drawing leader-line view/table crossing count:",
        "- mechanical drawing balloon geometry collision count:",
        "- mechanical drawing bbox helper envelope escape count:",
        "- mechanical drawing stale rendered preview count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    if not _mechanical_cad_passish(values["- mechanical drawing outside-frame ink audit verdict:"]):
        issues.append("mechanical CAD acceptance outside-frame ink audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing inner-frame safe-margin audit verdict:"]):
        issues.append("mechanical CAD acceptance inner-frame safe-margin audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing annotation ownership audit verdict:"]):
        issues.append("mechanical CAD acceptance annotation ownership audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing owner-zone coverage verdict:"]):
        issues.append("mechanical CAD acceptance owner-zone coverage verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing user-reported crop binding review verdict:"]):
        issues.append("mechanical CAD acceptance user-reported crop binding review verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing text exclusion halo audit verdict:"]):
        issues.append("mechanical CAD acceptance text exclusion halo audit verdict must be pass-shaped")
    for key in (
        "- mechanical drawing reported crop blocker count:",
        "- mechanical drawing text exclusion halo violation count:",
        "- mechanical drawing diagonal hatch/section/flow text-cover count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    if not _mechanical_cad_passish(values["- CAD text style/font audit verdict:"]):
        issues.append("mechanical CAD acceptance text style/font audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- normal CAD text entity verdict:"]):
        issues.append("mechanical CAD acceptance normal CAD text entity verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- current package SHA binding verdict:"]):
        issues.append("mechanical CAD acceptance current package SHA binding verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing reserved-zone intrusion audit verdict:"]):
        issues.append("mechanical CAD acceptance reserved-zone intrusion audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing hatch/section fill clipping audit verdict:"]):
        issues.append("mechanical CAD acceptance hatch/section fill clipping audit verdict must be pass-shaped")
    for key in (
        "- mechanical drawing reserved-zone intrusion count:",
        "- mechanical drawing dimension-line table-zone intrusion count:",
        "- mechanical drawing dimension-text table-zone intrusion count:",
        "- mechanical drawing view-geometry reserved-zone intrusion count:",
        "- mechanical drawing unowned free text count:",
        "- mechanical drawing unsupported floating text count:",
        "- mechanical drawing unbound scattered text count:",
        "- mechanical drawing dimension-like text without anchor count:",
        "- mechanical drawing cell padding violation count:",
        "- unsupported CAD font/style count:",
        "- artistic/vectorized text count:",
        "- source thick lineweight mismatch count:",
        "- source thin lineweight mismatch count:",
        "- mechanical drawing hatch/section fill boundary violation count:",
        "- mechanical drawing hatch/section fill adjacent-view crossing count:",
        "- mechanical drawing hatch/section fill dimension-line crossing count:",
        "- mechanical drawing hatch/section fill title-block/table/BOM/frame crossing count:",
        "- mechanical drawing hatch/section fill blank-background leak count:",
        "- mechanical drawing text entity overlap count:",
        "- mechanical drawing outside-frame independent ink component count:",
        "- mechanical drawing outside-frame text component count:",
        "- mechanical drawing outside-frame leader component count:",
        "- mechanical drawing outside-frame hatch/section component count:",
        "- mechanical drawing outside-frame table/title-block component count:",
        "- mechanical drawing right safe-boundary intrusion count:",
        "- mechanical drawing leader-text inner-frame intrusion count:",
        "- mechanical drawing view-geometry inner-frame intrusion count:",
        "- mechanical drawing dimension-text inner-frame intrusion count:",
        "- mechanical drawing missing short table line count:",
        "- mechanical drawing broken cell-border count:",
        "- mechanical drawing table-grid topology mismatch count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    min_padding = _mechanical_cad_float(values["- mechanical drawing min cell padding mm:"])
    if min_padding is None or min_padding < 1.0:
        issues.append("mechanical CAD acceptance min cell padding mm must be numeric and >= 1.0")
    for key, expected in (
        ("- source thick lineweight required:", "50"),
        ("- source thick lineweight observed:", "50"),
        ("- source thin lineweight required:", "25"),
        ("- source thin lineweight observed:", "25"),
    ):
        value = normalize(values[key])
        if expected not in value and f"0.{expected}" not in value:
            issues.append(f"mechanical CAD acceptance {key[2:]} must include {expected} or 0.{expected}")
    if not _mechanical_cad_passish(values["- mechanical drawing text legibility machine audit verdict:"]):
        issues.append("mechanical CAD acceptance text legibility machine audit verdict must be pass-shaped")
    cad_text_height = _mechanical_cad_float(values["- mechanical drawing minimum CAD text height mm:"])
    if cad_text_height is None or cad_text_height < MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM:
        issues.append(
            "mechanical CAD acceptance minimum CAD text height must be "
            f">= {MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM:g} mm"
        )
    rendered_text_height = _mechanical_cad_float(values["- mechanical drawing minimum rendered text height px:"])
    if rendered_text_height is None or rendered_text_height < MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX:
        issues.append(
            "mechanical CAD acceptance minimum rendered text height must be "
            f">= {MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX:g} px"
        )
    if not _mechanical_cad_passish(values["- mechanical drawing text integrity audit verdict:"]):
        issues.append("mechanical CAD acceptance text integrity audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(values["- mechanical drawing text orientation audit verdict:"]):
        issues.append("mechanical CAD acceptance text orientation audit verdict must be pass-shaped")
    for key in (
        "- mechanical drawing mojibake/tofu/missing glyph count:",
        "- mechanical drawing missing required drawing text count:",
        "- mechanical drawing upside-down text count:",
        "- mechanical drawing mirrored text count:",
    ):
        count = _mechanical_cad_int(values[key])
        if count is None or count != 0:
            issues.append(f"mechanical CAD acceptance {key[2:]} must be numeric 0")
    if not _mechanical_cad_passish(values["- mechanical drawing manufacturing complexity audit verdict:"]):
        issues.append("mechanical CAD acceptance manufacturing complexity audit verdict must be pass-shaped")
    detail_family_count = _mechanical_cad_int(values["- mechanical drawing minimum manufacturing detail family count:"])
    if detail_family_count is None or detail_family_count < MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES:
        issues.append(
            "mechanical CAD acceptance manufacturing detail family count must be "
            f">= {MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES}"
        )
    if (
        normalize(values["- mechanical drawing entity-count-only false-pass verdict:"]).lower()
        not in MECHANICAL_CAD_ENTITY_ONLY_REJECT_VALUES
    ):
        issues.append("mechanical CAD acceptance must reject entity-count-only acceptance")
    thesis_docx_value = normalize(values["- thesis DOCX mutation verdict:"]).lower()
    if thesis_docx_value not in {"not-touched", "none", "not-applicable", "paused"}:
        issues.append("mechanical CAD acceptance must not imply thesis DOCX mutation in the CAD-only lane")

    if source_delta_report_path is not None and not source_delta_report_issues:
        try:
            source_delta_report = json.loads(source_delta_report_path.read_text(encoding="utf-8"))
        except Exception as exc:
            issues.append(
                f"mechanical CAD source linework delta report is not valid UTF-8 JSON: {source_delta_report_path} ({exc})"
            )
        else:
            if source_delta_report.get("schema") != "graduation-project-builder.cad-source-differentiation-audit.v1":
                issues.append("mechanical CAD source linework delta report must use cad-source-differentiation schema v1")
            if source_delta_report.get("passed") is not True:
                issues.append("mechanical CAD source linework delta report must have passed=true")
            if source_delta_report.get("issues") not in ([], None):
                issues.append("mechanical CAD source linework delta report must not carry unresolved issues")
            changed_sheet_count = _mechanical_cad_int(source_delta_report.get("changed_dxf_count"))
            identical_sheet_count = _mechanical_cad_int(source_delta_report.get("identical_dxf_count"))
            old_overlay_count = _mechanical_cad_int(source_delta_report.get("old_like_large_circle_overlay_count"))
            if changed_sheet_count is not None and changed_sheet_count != _mechanical_cad_int(values["- mechanical drawing changed source sheet count:"]):
                issues.append("mechanical CAD acceptance changed source sheet count differs from source linework delta report")
            if identical_sheet_count is not None and identical_sheet_count != _mechanical_cad_int(values["- mechanical drawing identical source sheet count:"]):
                issues.append("mechanical CAD acceptance identical source sheet count differs from source linework delta report")
            if old_overlay_count is not None and old_overlay_count != _mechanical_cad_int(values["- mechanical drawing old-like large circle overlay count:"]):
                issues.append("mechanical CAD acceptance old-like large circle overlay count differs from source linework delta report")
            per_file = source_delta_report.get("per_file")
            if isinstance(per_file, list):
                report_entity_delta = sum(
                    (_mechanical_cad_int(row.get("added_signature_count")) or 0)
                    + (_mechanical_cad_int(row.get("removed_signature_count")) or 0)
                    for row in per_file
                    if isinstance(row, dict)
                )
                if report_entity_delta and report_entity_delta != _mechanical_cad_int(values["- mechanical drawing changed source entity count:"]):
                    issues.append("mechanical CAD acceptance changed source entity count differs from source linework delta report")
            report_current_sha = normalize(str(source_delta_report.get("current_sha256") or "")).upper()
            report_baseline_sha = normalize(str(source_delta_report.get("baseline_sha256") or "")).upper()
            if report_current_sha and report_current_sha != final_source_sha:
                issues.append("mechanical CAD acceptance final CAD package SHA256 differs from source linework delta report")
            if report_baseline_sha and report_baseline_sha != source_baseline_sha:
                issues.append("mechanical CAD acceptance source CAD package SHA256 differs from source linework delta report")

    if render_readability_report_path is not None and not render_readability_report_issues:
        try:
            render_readability_report = json.loads(render_readability_report_path.read_text(encoding="utf-8"))
        except Exception as exc:
            issues.append(
                f"mechanical CAD render readability report is not valid UTF-8 JSON: {render_readability_report_path} ({exc})"
            )
        else:
            if render_readability_report.get("schema") != "graduation-project-builder.mechanical-render-readability.v1":
                issues.append("mechanical CAD render readability report must use mechanical-render-readability schema v1")
            if render_readability_report.get("passed") is not True:
                issues.append("mechanical CAD render readability report must have passed=true")
            for prefix, report_key in (
                ("- text text overlap count:", "text_text_overlap_count"),
                ("- text graphic cover count:", "text_graphic_cover_count"),
                ("- severe line crowding count:", "severe_line_crowding_count"),
            ):
                report_count = _mechanical_cad_int(render_readability_report.get(report_key))
                record_count = _mechanical_cad_int(values[prefix])
                if report_count is None:
                    issues.append(f"mechanical CAD render readability report missing integer {report_key}")
                elif report_count != 0:
                    issues.append(f"mechanical CAD render readability report {report_key} must be 0")
                if report_count is not None and record_count is not None and report_count != record_count:
                    issues.append(
                        f"mechanical CAD acceptance {prefix[2:]} differs from render readability report {report_key}"
                    )

    rendered_paths = split_path_values(raw_values["- mechanical drawing rendered review evidence paths:"])
    if not rendered_paths:
        issues.append("mechanical CAD acceptance must bind rendered review evidence paths")
    for raw_path in rendered_paths:
        rendered_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(rendered_path, require_nonempty_file=True))
        if rendered_path.suffix.lower() not in IMAGE_EXTENSIONS | PDF_EXTENSIONS:
            issues.append(f"mechanical CAD rendered review evidence must be image/PDF: {rendered_path}")

    content_overlap_evidence_paths = split_path_values(raw_values["- mechanical drawing content-overlap evidence path:"])
    if not content_overlap_evidence_paths:
        issues.append("mechanical CAD acceptance must bind content-overlap audit evidence path")
    for raw_path in content_overlap_evidence_paths:
        content_overlap_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(content_overlap_path, require_nonempty_file=True))
        if content_overlap_path.suffix.lower() not in {".json"}:
            issues.append(f"mechanical CAD content-overlap audit evidence must be JSON: {content_overlap_path}")

    hatch_evidence_paths = split_path_values(raw_values["- mechanical drawing hatch/section fill clipping evidence path:"])
    if not hatch_evidence_paths:
        issues.append("mechanical CAD acceptance must bind hatch/section fill clipping evidence path")
    for raw_path in hatch_evidence_paths:
        hatch_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(hatch_path, require_nonempty_file=True))
        if hatch_path.suffix.lower() not in {".json"}:
            issues.append(f"mechanical CAD hatch/section fill clipping evidence must be JSON: {hatch_path}")

    outside_frame_evidence_paths = split_path_values(raw_values["- mechanical drawing outside-frame ink audit evidence path:"])
    if not outside_frame_evidence_paths:
        issues.append("mechanical CAD acceptance must bind outside-frame ink audit evidence path")
    for raw_path in outside_frame_evidence_paths:
        outside_frame_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(outside_frame_path, require_nonempty_file=True))
        if outside_frame_path.suffix.lower() not in {".json"}:
            issues.append(f"mechanical CAD outside-frame ink audit evidence must be JSON: {outside_frame_path}")

    inner_frame_evidence_paths = split_path_values(raw_values["- mechanical drawing inner-frame safe-margin evidence path:"])
    if not inner_frame_evidence_paths:
        issues.append("mechanical CAD acceptance must bind inner-frame safe-margin evidence path")
    for raw_path in inner_frame_evidence_paths:
        inner_frame_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(inner_frame_path, require_nonempty_file=True))
        if inner_frame_path.suffix.lower() != ".json":
            issues.append(f"mechanical CAD inner-frame safe-margin evidence must be JSON: {inner_frame_path}")

    text_integrity_paths = split_path_values(raw_values["- mechanical drawing text integrity audit path:"])
    if not text_integrity_paths:
        issues.append("mechanical CAD acceptance must bind text integrity audit path")
    for raw_path in text_integrity_paths:
        text_integrity_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(text_integrity_path, require_nonempty_file=True))
        if text_integrity_path.suffix.lower() != ".json":
            issues.append(f"mechanical CAD text integrity audit evidence must be JSON: {text_integrity_path}")

    text_orientation_paths = split_path_values(raw_values["- mechanical drawing text orientation audit path:"])
    if not text_orientation_paths:
        issues.append("mechanical CAD acceptance must bind text orientation audit path")
    for raw_path in text_orientation_paths:
        text_orientation_path = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(text_orientation_path, require_nonempty_file=True))
        if text_orientation_path.suffix.lower() != ".json":
            issues.append(f"mechanical CAD text orientation audit evidence must be JSON: {text_orientation_path}")

    if audit_report_path is not None and not audit_report_issues:
        try:
            report = json.loads(audit_report_path.read_text(encoding="utf-8"))
        except Exception as exc:
            issues.append(f"mechanical CAD audit report is not valid UTF-8 JSON: {audit_report_path} ({exc})")
        else:
            if report.get("schema") not in MECHANICAL_CAD_AUDIT_SCHEMA_VALUES:
                issues.append("mechanical CAD audit report must use v4-or-stricter mechanical drawing package schema")
            if report.get("passed") is not True:
                issues.append("mechanical CAD audit report must have passed=true")
            if report.get("issues") not in ([], None):
                issues.append("mechanical CAD audit report must not carry unresolved issues")
            report_path_value = _mechanical_cad_audit_path_value(report, "package_path")
            report_sha_value = _mechanical_cad_audit_path_value(report, "package_sha256") or _mechanical_cad_audit_path_value(report, "sha256")
            if audited_package is not None and report_path_value:
                try:
                    if Path(report_path_value).resolve() != audited_package.resolve():
                        issues.append("mechanical CAD audit report package_path differs from acceptance audited package path")
                except OSError:
                    issues.append("mechanical CAD audit report package_path could not be resolved")
            if report_sha_value and report_sha_value.upper() != values["- exact audited CAD package sha256:"].upper():
                issues.append("mechanical CAD audit report package SHA256 differs from acceptance audited package SHA256")
            for key in ("density_verdict", "manufacturing_depth", "rendered_review_verdict"):
                section = report.get(key)
                if not isinstance(section, dict) or section.get("passed") is not True:
                    issues.append(f"mechanical CAD audit report {key}.passed must be true")
            formal_source = report.get("formal_cad_source_verdict")
            if not isinstance(formal_source, dict) or formal_source.get("passed") is not True:
                issues.append("mechanical CAD audit report formal_cad_source_verdict.passed must be true")
            else:
                if (_mechanical_cad_int(formal_source.get("editable_cad_source_count")) or 0) <= 0:
                    issues.append("mechanical CAD audit report must prove at least one real DWG/DXF source file")
                if formal_source.get("schematic_substitute_hits"):
                    issues.append("mechanical CAD audit report must have zero schematic/concept substitute hits")
                source_verdict = normalize(values["- mechanical drawing formal CAD source provenance verdict:"]).lower()
                substitute_verdict = normalize(
                    values["- mechanical drawing schematic/concept substitute rejection verdict:"]
                ).lower()
                if source_verdict not in MECHANICAL_CAD_PASS_VALUES:
                    issues.append("mechanical CAD acceptance formal CAD source provenance field must be pass-shaped")
                if substitute_verdict not in MECHANICAL_CAD_PASS_VALUES:
                    issues.append("mechanical CAD acceptance schematic/concept substitute field must be pass-shaped")
                report_substitute_verdict = normalize(
                    str(formal_source.get("schematic_substitute_rejection_verdict", ""))
                ).lower()
                if report_substitute_verdict not in MECHANICAL_CAD_PASS_VALUES:
                    issues.append(
                        "mechanical CAD audit report formal_cad_source_verdict "
                        "schematic_substitute_rejection_verdict must be pass-shaped"
                    )
            candidate = report.get("candidate")
            extension_counts = candidate.get("extension_counts") if isinstance(candidate, dict) else {}
            if isinstance(extension_counts, dict):
                if int(extension_counts.get(".dwg", 0)) <= 0:
                    issues.append("mechanical CAD audit report must count real DWG files")
                if int(extension_counts.get(".pdf", 0)) <= 0:
                    issues.append("mechanical CAD audit report must count PDF drawings")
            cad_text_quality = report.get("cad_text_quality_verdict")
            if not isinstance(cad_text_quality, dict) or cad_text_quality.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing cad_text_quality_verdict")
            else:
                text_quality_field_map = {
                    "- mechanical drawing mojibake/tofu/missing glyph count:": "mojibake_or_missing_glyph_count",
                    "- mechanical drawing missing required drawing text count:": "missing_required_drawing_text_count",
                    "- mechanical drawing upside-down text count:": "upside_down_text_count",
                    "- mechanical drawing mirrored text count:": "mirrored_text_count",
                }
                for prefix, report_field in text_quality_field_map.items():
                    report_count = _mechanical_cad_int(cad_text_quality.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report cad_text_quality_verdict.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"cad_text_quality_verdict.{report_field}"
                        )
            rendered_verdict = report.get("rendered_review_verdict")
            if isinstance(rendered_verdict, dict) and int(rendered_verdict.get("accepted_review_count", 0)) <= 0:
                issues.append("mechanical CAD audit report must bind at least one accepted rendered review")
            frame_overflow_verdict = report.get("rendered_frame_overflow_verdict")
            if not isinstance(frame_overflow_verdict, dict) or frame_overflow_verdict.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_frame_overflow_verdict")
            else:
                frame_overflow_component_count = _mechanical_cad_int(
                    frame_overflow_verdict.get("outside_frame_component_count")
                )
                frame_overflow_pixel_count = _mechanical_cad_int(frame_overflow_verdict.get("outside_frame_pixel_count"))
                frame_overflow_max_area = _mechanical_cad_int(frame_overflow_verdict.get("max_outside_component_area_px"))
                expected_outside_components = _mechanical_cad_int(
                    values["- mechanical drawing outside-frame independent ink component count:"]
                )
                if frame_overflow_component_count is None or frame_overflow_component_count != 0:
                    issues.append("mechanical CAD audit report rendered_frame_overflow_verdict.outside_frame_component_count must be 0")
                if frame_overflow_pixel_count is None or frame_overflow_pixel_count != 0:
                    issues.append("mechanical CAD audit report rendered_frame_overflow_verdict.outside_frame_pixel_count must be 0")
                if frame_overflow_max_area is None or frame_overflow_max_area != 0:
                    issues.append("mechanical CAD audit report rendered_frame_overflow_verdict.max_outside_component_area_px must be 0")
                if expected_outside_components != frame_overflow_component_count:
                    issues.append(
                        "mechanical CAD acceptance outside-frame independent ink component count differs from "
                        "audit report rendered_frame_overflow_verdict.outside_frame_component_count"
                    )
            machine_overlap = rendered_verdict.get("machine_overlap_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(machine_overlap, dict) or machine_overlap.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.machine_overlap_audit")
            else:
                report_text_overlap = _mechanical_cad_int(machine_overlap.get("text_entity_overlap_count"))
                expected_text_overlap = _mechanical_cad_int(values["- mechanical drawing text entity overlap count:"])
                if report_text_overlap is None or report_text_overlap != 0:
                    issues.append("mechanical CAD audit report machine_overlap_audit.text_entity_overlap_count must be 0")
                if expected_text_overlap != report_text_overlap:
                    issues.append(
                        "mechanical CAD acceptance text entity overlap count differs from audit report "
                        "machine_overlap_audit.text_entity_overlap_count"
                    )
            content_overlap = rendered_verdict.get("content_overlap_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(content_overlap, dict) or content_overlap.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.content_overlap_audit")
            else:
                content_positive_field_map = {
                    "- mechanical drawing registered content bbox count:": "registered_bbox_count",
                    "- mechanical drawing checked content pair count:": "checked_pair_count",
                }
                for prefix, report_field in content_positive_field_map.items():
                    report_count = _mechanical_cad_int(content_overlap.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count <= 0:
                        issues.append(f"mechanical CAD audit report content_overlap_audit.{report_field} must be > 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"content_overlap_audit.{report_field}"
                        )
                content_zero_field_map = {
                    "- mechanical drawing content overlap count:": "content_overlap_count",
                    "- mechanical drawing view-view overlap count:": "view_view_overlap_count",
                    "- mechanical drawing detail-frame-main-view overlap count:": "detail_frame_main_view_overlap_count",
                    "- mechanical drawing table text/grid collision count:": "table_text_grid_collision_count",
                    "- mechanical drawing dimension-line view/table crossing count:": "dimension_line_view_table_crossing_count",
                    "- mechanical drawing leader-line view/table crossing count:": "leader_line_view_table_crossing_count",
                    "- mechanical drawing balloon geometry collision count:": "balloon_geometry_collision_count",
                    "- mechanical drawing bbox helper envelope escape count:": "bbox_helper_envelope_escape_count",
                    "- mechanical drawing stale rendered preview count:": "stale_rendered_preview_count",
                }
                for prefix, report_field in content_zero_field_map.items():
                    report_count = _mechanical_cad_int(content_overlap.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report content_overlap_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"content_overlap_audit.{report_field}"
                        )
            outside_frame_ink = rendered_verdict.get("outside_frame_ink_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(outside_frame_ink, dict) or outside_frame_ink.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.outside_frame_ink_audit")
            else:
                outside_frame_field_map = {
                    "- mechanical drawing outside-frame independent ink component count:": "outside_frame_independent_ink_component_count",
                    "- mechanical drawing outside-frame text component count:": "outside_frame_text_component_count",
                    "- mechanical drawing outside-frame leader component count:": "outside_frame_leader_component_count",
                    "- mechanical drawing outside-frame hatch/section component count:": "outside_frame_hatch_section_component_count",
                    "- mechanical drawing outside-frame table/title-block component count:": "outside_frame_table_title_block_component_count",
                }
                for prefix, report_field in outside_frame_field_map.items():
                    report_count = _mechanical_cad_int(outside_frame_ink.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report outside_frame_ink_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"outside_frame_ink_audit.{report_field}"
                        )
            inner_frame_safe_margin = (
                rendered_verdict.get("inner_frame_safe_margin_audit") if isinstance(rendered_verdict, dict) else None
            )
            if not isinstance(inner_frame_safe_margin, dict) or inner_frame_safe_margin.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.inner_frame_safe_margin_audit")
            else:
                inner_frame_field_map = {
                    "- mechanical drawing right safe-boundary intrusion count:": "right_safe_boundary_intrusion_count",
                    "- mechanical drawing leader-text inner-frame intrusion count:": "leader_text_inner_frame_intrusion_count",
                    "- mechanical drawing view-geometry inner-frame intrusion count:": "view_geometry_inner_frame_intrusion_count",
                    "- mechanical drawing dimension-text inner-frame intrusion count:": "dimension_text_inner_frame_intrusion_count",
                }
                source_sections = inner_frame_safe_margin.get("source_sections")
                if not isinstance(source_sections, list) or not {
                    "machine_overlap_audit",
                    "content_overlap_audit",
                    "reserved_zone_intrusion_audit",
                    "outside_frame_ink_audit",
                }.issubset({str(item) for item in source_sections}):
                    issues.append("mechanical CAD audit report inner_frame_safe_margin_audit must cite source machine sections")
                if (_mechanical_cad_float(inner_frame_safe_margin.get("min_safe_margin_mm")) or 0.0) < 2.0:
                    issues.append("mechanical CAD audit report inner_frame_safe_margin_audit.min_safe_margin_mm must be >= 2.0")
                for prefix, report_field in inner_frame_field_map.items():
                    report_count = _mechanical_cad_int(inner_frame_safe_margin.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report inner_frame_safe_margin_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"inner_frame_safe_margin_audit.{report_field}"
                        )
            cell_containment = rendered_verdict.get("cell_containment_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(cell_containment, dict) or cell_containment.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.cell_containment_audit")
            else:
                for field in (
                    "outside_cell_count",
                    "border_touch_count",
                    "unowned_table_text_count",
                    "clipped_overflow_count",
                ):
                    report_count = _mechanical_cad_int(cell_containment.get(field))
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report cell_containment_audit.{field} must be 0")
            annotation_ownership = (
                rendered_verdict.get("annotation_ownership_audit") if isinstance(rendered_verdict, dict) else None
            )
            if not isinstance(annotation_ownership, dict) or annotation_ownership.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.annotation_ownership_audit")
            else:
                annotation_field_map = {
                    "- mechanical drawing unowned free text count:": "unowned_free_text_count",
                    "- mechanical drawing unsupported floating text count:": "unsupported_floating_text_count",
                }
                for prefix, report_field in annotation_field_map.items():
                    report_count = _mechanical_cad_int(annotation_ownership.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report annotation_ownership_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"annotation_ownership_audit.{report_field}"
                        )
            reserved_zone_intrusion = (
                rendered_verdict.get("reserved_zone_intrusion_audit") if isinstance(rendered_verdict, dict) else None
            )
            if not isinstance(reserved_zone_intrusion, dict) or reserved_zone_intrusion.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.reserved_zone_intrusion_audit")
            else:
                reserved_field_map = {
                    "- mechanical drawing reserved-zone intrusion count:": "intrusion_count",
                    "- mechanical drawing protected-table-zone intrusion count:": "protected_table_zone_intrusion_count",
                    "- mechanical drawing dimension-line table-zone intrusion count:": "dimension_line_table_zone_intrusion_count",
                    "- mechanical drawing dimension-text table-zone intrusion count:": "dimension_text_table_zone_intrusion_count",
                    "- mechanical drawing view-geometry table-zone intrusion count:": "view_geometry_table_zone_intrusion_count",
                    "- mechanical drawing detail-view table-zone intrusion count:": "detail_view_table_zone_intrusion_count",
                    "- mechanical drawing leader/balloon table-zone intrusion count:": "leader_balloon_table_zone_intrusion_count",
                    "- mechanical drawing dimension table-zone intrusion count:": "dimension_table_zone_intrusion_count",
                    "- mechanical drawing title-block/BOM protected-zone intrusion count:": "title_block_bom_protected_zone_intrusion_count",
                    "- mechanical drawing view-geometry reserved-zone intrusion count:": "view_geometry_reserved_zone_intrusion_count",
                }
                for prefix, report_field in reserved_field_map.items():
                    report_count = _mechanical_cad_int(reserved_zone_intrusion.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report reserved_zone_intrusion_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report "
                            f"reserved_zone_intrusion_audit.{report_field}"
                        )
            hatch_clip = rendered_verdict.get("hatch_clip_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(hatch_clip, dict) or hatch_clip.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.hatch_clip_audit")
            else:
                hatch_field_map = {
                    "- mechanical drawing hatch/section fill boundary violation count:": "hatch_clip_violation_count",
                    "- mechanical drawing hatch/section fill adjacent-view crossing count:": "adjacent_view_crossing_count",
                    "- mechanical drawing hatch/section fill dimension-line crossing count:": "dimension_line_crossing_count",
                    "- mechanical drawing hatch/section fill title-block/table/BOM/frame crossing count:": "title_block_table_bom_frame_crossing_count",
                    "- mechanical drawing hatch/section fill blank-background leak count:": "blank_background_leak_count",
                }
                for prefix, report_field in hatch_field_map.items():
                    report_count = _mechanical_cad_int(hatch_clip.get(report_field))
                    expected_count = _mechanical_cad_int(values[prefix])
                    if report_count is None or report_count != 0:
                        issues.append(f"mechanical CAD audit report hatch_clip_audit.{report_field} must be 0")
                    if expected_count != report_count:
                        issues.append(f"mechanical CAD acceptance {prefix[2:]} differs from audit report hatch_clip_audit.{report_field}")
            text_legibility = rendered_verdict.get("text_legibility_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(text_legibility, dict) or text_legibility.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.text_legibility_audit")
            elif (
                (_mechanical_cad_float(text_legibility.get("min_cad_text_height_mm")) or 0.0)
                < MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM
                or (_mechanical_cad_float(text_legibility.get("min_rendered_text_height_px")) or 0.0)
                < MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX
            ):
                issues.append("mechanical CAD audit report text_legibility_audit is below required legibility thresholds")
            text_integrity = rendered_verdict.get("text_integrity_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(text_integrity, dict) or text_integrity.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.text_integrity_audit")
            else:
                for field in ("mojibake_or_missing_glyph_count", "missing_required_drawing_text_count"):
                    count = _mechanical_cad_int(text_integrity.get(field))
                    if count is None or count != 0:
                        issues.append(f"mechanical CAD audit report text_integrity_audit.{field} must be 0")
            text_orientation = rendered_verdict.get("text_orientation_audit") if isinstance(rendered_verdict, dict) else None
            if not isinstance(text_orientation, dict) or text_orientation.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.text_orientation_audit")
            else:
                for field in ("upside_down_text_count", "mirrored_text_count"):
                    count = _mechanical_cad_int(text_orientation.get(field))
                    if count is None or count != 0:
                        issues.append(f"mechanical CAD audit report text_orientation_audit.{field} must be 0")
            manufacturing_complexity = (
                rendered_verdict.get("manufacturing_complexity_audit") if isinstance(rendered_verdict, dict) else None
            )
            if not isinstance(manufacturing_complexity, dict) or manufacturing_complexity.get("passed") is not True:
                issues.append("mechanical CAD audit report must include passing rendered_review_verdict.manufacturing_complexity_audit")
            elif (
                (_mechanical_cad_int(manufacturing_complexity.get("min_feature_family_count")) or 0)
                < MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES
            ):
                issues.append("mechanical CAD audit report manufacturing_complexity_audit is below required family count")
            rendered_candidate = candidate.get("rendered_review") if isinstance(candidate, dict) else None
            if not isinstance(rendered_candidate, dict):
                issues.append("mechanical CAD audit report candidate.rendered_review must exist")
            else:
                candidate_machine_overlap = rendered_candidate.get("machine_overlap_audit")
                if not isinstance(candidate_machine_overlap, dict) or candidate_machine_overlap.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.machine_overlap_audit must pass")
                elif (_mechanical_cad_int(candidate_machine_overlap.get("text_entity_overlap_count")) or 0) != 0:
                    issues.append(
                        "mechanical CAD audit report candidate.rendered_review.machine_overlap_audit.text_entity_overlap_count must be 0"
                    )
                candidate_content_overlap = rendered_candidate.get("content_overlap_audit")
                if not isinstance(candidate_content_overlap, dict) or candidate_content_overlap.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.content_overlap_audit must pass")
                elif (
                    (_mechanical_cad_int(candidate_content_overlap.get("registered_bbox_count")) or 0) <= 0
                    or (_mechanical_cad_int(candidate_content_overlap.get("checked_pair_count")) or 0) <= 0
                    or any(
                        (_mechanical_cad_int(candidate_content_overlap.get(field)) or 0) != 0
                        for field in (
                            "content_overlap_count",
                            "view_view_overlap_count",
                            "detail_frame_main_view_overlap_count",
                            "table_text_grid_collision_count",
                            "dimension_line_view_table_crossing_count",
                            "leader_line_view_table_crossing_count",
                            "balloon_geometry_collision_count",
                            "bbox_helper_envelope_escape_count",
                            "stale_rendered_preview_count",
                        )
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.content_overlap_audit counts are invalid")
                candidate_outside_frame = rendered_candidate.get("outside_frame_ink_audit")
                if not isinstance(candidate_outside_frame, dict) or candidate_outside_frame.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.outside_frame_ink_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_outside_frame.get(field)) or 0) != 0
                    for field in (
                        "outside_frame_independent_ink_component_count",
                        "outside_frame_text_component_count",
                        "outside_frame_leader_component_count",
                        "outside_frame_hatch_section_component_count",
                        "outside_frame_table_title_block_component_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.outside_frame_ink_audit counts must all be 0")
                candidate_inner_frame = rendered_candidate.get("inner_frame_safe_margin_audit")
                if not isinstance(candidate_inner_frame, dict) or candidate_inner_frame.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.inner_frame_safe_margin_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_inner_frame.get(field)) or 0) != 0
                    for field in (
                        "right_safe_boundary_intrusion_count",
                        "leader_text_inner_frame_intrusion_count",
                        "view_geometry_inner_frame_intrusion_count",
                        "dimension_text_inner_frame_intrusion_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.inner_frame_safe_margin_audit counts must all be 0")
                candidate_cell_containment = rendered_candidate.get("cell_containment_audit")
                if not isinstance(candidate_cell_containment, dict) or candidate_cell_containment.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.cell_containment_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_cell_containment.get(field)) or 0) != 0
                    for field in (
                        "outside_cell_count",
                        "border_touch_count",
                        "unowned_table_text_count",
                        "clipped_overflow_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.cell_containment_audit counts must all be 0")
                candidate_annotation_ownership = rendered_candidate.get("annotation_ownership_audit")
                if (
                    not isinstance(candidate_annotation_ownership, dict)
                    or candidate_annotation_ownership.get("passed") is not True
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.annotation_ownership_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_annotation_ownership.get(field)) or 0) != 0
                    for field in (
                        "unowned_free_text_count",
                        "unsupported_floating_text_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.annotation_ownership_audit counts must all be 0")
                candidate_reserved_zone = rendered_candidate.get("reserved_zone_intrusion_audit")
                if not isinstance(candidate_reserved_zone, dict) or candidate_reserved_zone.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.reserved_zone_intrusion_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_reserved_zone.get(field)) or 0) != 0
                    for field in (
                        "intrusion_count",
                        "geometry_title_block_intrusion_count",
                        "dimension_line_table_zone_intrusion_count",
                        "dimension_text_table_zone_intrusion_count",
                        "leader_table_zone_intrusion_count",
                        "view_geometry_reserved_zone_intrusion_count",
                        "bom_table_intrusion_count",
                        "technical_requirement_intrusion_count",
                        "protected_table_zone_intrusion_count",
                        "view_geometry_table_zone_intrusion_count",
                        "detail_view_table_zone_intrusion_count",
                        "leader_balloon_table_zone_intrusion_count",
                        "dimension_table_zone_intrusion_count",
                        "title_block_bom_protected_zone_intrusion_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.reserved_zone_intrusion_audit counts must all be 0")
                candidate_hatch_clip = rendered_candidate.get("hatch_clip_audit")
                if not isinstance(candidate_hatch_clip, dict) or candidate_hatch_clip.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.hatch_clip_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_hatch_clip.get(field)) or 0) != 0
                    for field in (
                        "hatch_clip_violation_count",
                        "entity_boundary_escape_count",
                        "adjacent_view_crossing_count",
                        "dimension_line_crossing_count",
                        "title_block_table_bom_frame_crossing_count",
                        "blank_background_leak_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.hatch_clip_audit counts must all be 0")
                candidate_text_legibility = rendered_candidate.get("text_legibility_audit")
                if not isinstance(candidate_text_legibility, dict) or candidate_text_legibility.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_legibility_audit must pass")
                elif (
                    (_mechanical_cad_float(candidate_text_legibility.get("min_cad_text_height_mm")) or 0.0)
                    < MECHANICAL_CAD_MIN_ACCEPTANCE_CAD_TEXT_MM
                    or (_mechanical_cad_float(candidate_text_legibility.get("min_rendered_text_height_px")) or 0.0)
                    < MECHANICAL_CAD_MIN_ACCEPTANCE_RENDERED_TEXT_PX
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_legibility_audit is below thresholds")
                candidate_text_integrity = rendered_candidate.get("text_integrity_audit")
                if not isinstance(candidate_text_integrity, dict) or candidate_text_integrity.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_integrity_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_text_integrity.get(field)) or 0) != 0
                    for field in (
                        "mojibake_or_missing_glyph_count",
                        "missing_required_drawing_text_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_integrity_audit counts must all be 0")
                candidate_text_orientation = rendered_candidate.get("text_orientation_audit")
                if not isinstance(candidate_text_orientation, dict) or candidate_text_orientation.get("passed") is not True:
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_orientation_audit must pass")
                elif any(
                    (_mechanical_cad_int(candidate_text_orientation.get(field)) or 0) != 0
                    for field in (
                        "upside_down_text_count",
                        "mirrored_text_count",
                    )
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.text_orientation_audit counts must all be 0")
                candidate_manufacturing_complexity = rendered_candidate.get("manufacturing_complexity_audit")
                if (
                    not isinstance(candidate_manufacturing_complexity, dict)
                    or candidate_manufacturing_complexity.get("passed") is not True
                ):
                    issues.append("mechanical CAD audit report candidate.rendered_review.manufacturing_complexity_audit must pass")
                elif (
                    (_mechanical_cad_int(candidate_manufacturing_complexity.get("min_feature_family_count")) or 0)
                    < MECHANICAL_CAD_MIN_ACCEPTANCE_DETAIL_FAMILIES
                ):
                    issues.append(
                        "mechanical CAD audit report candidate.rendered_review.manufacturing_complexity_audit is below required family count"
                    )
                for prefix in (
                    "- mechanical drawing rendered no-overlap verdict:",
                    "- mechanical drawing boundary clearance verdict:",
                    "- mechanical drawing detail density verdict:",
                    "- mechanical drawing title block/table/notes isolation verdict:",
                    "- mechanical drawing title-block cell containment verdict:",
                    "- mechanical drawing annotation margin clearance verdict:",
                    "- mechanical drawing local crowding verdict:",
                    "- mechanical drawing text/table/frame overlap verdict:",
                    "- mechanical drawing text integrity audit verdict:",
                    "- mechanical drawing text orientation audit verdict:",
                    "- mechanical drawing content-overlap audit verdict:",
                    "- mechanical drawing outside-frame ink audit verdict:",
                    "- mechanical drawing inner-frame safe-margin audit verdict:",
                    "- mechanical drawing hatch/section fill clipping audit verdict:",
                ):
                    report_field = _mechanical_cad_acceptance_to_report_field(prefix)
                    if not report_field:
                        continue
                    report_value = normalize(str(rendered_candidate.get(report_field, ""))).lower()
                    expected_value = normalize(values[prefix]).lower()
                    if prefix in {
                        "- mechanical drawing rendered no-overlap verdict:",
                        "- mechanical drawing text/table/frame overlap verdict:",
                    }:
                        if report_value not in MECHANICAL_CAD_NO_OVERLAP_VALUES:
                            issues.append(f"mechanical CAD audit report rendered_review.{report_field} must be pass/no-overlap")
                    else:
                        if report_value not in MECHANICAL_CAD_PASS_VALUES:
                            issues.append(f"mechanical CAD audit report rendered_review.{report_field} must be pass-shaped")
                    if expected_value != report_value:
                        issues.append(
                            f"mechanical CAD acceptance {prefix[2:]} differs from audit report rendered_review.{report_field}"
                        )

    issues.extend(_mechanical_cad_validate_validation_command(raw_values["- validation command:"], record_path))
    if values["- validation result:"] != "pass":
        issues.append("mechanical CAD acceptance validation result must be pass")
    return issues


def validate_skill_invocation_lock(
    *,
    record_path: Path,
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    task_mode: str,
    selected_workflow: str,
    final_docx_path: Path | None,
    expected_output_sha256: str,
) -> list[str]:
    issues: list[str] = []
    lock_value = gate_values.get("- skill invocation lock path:", "")
    lock_raw = gate_raw_values.get("- skill invocation lock path:", "")
    lock_verdict = gate_values.get("- skill invocation lock verdict:", "")
    if not is_explicit(lock_value) or is_explicit_none(lock_value):
        return ["skill invocation lock path must name a lock record for explicit graduation-project-builder runs"]
    if not skill_lock_passish(lock_verdict):
        issues.append(f"skill invocation lock verdict must be pass-shaped: {lock_verdict or 'missing'}")
    raw_paths = split_path_values(lock_raw)
    if not raw_paths:
        return issues + ["skill invocation lock path did not contain a resolvable path"]
    lock_path = resolve_record_path(raw_paths[0], record_path)
    issues.extend(validate_existing_path(lock_path, require_nonempty_file=True))
    if issues:
        return issues

    values, raw_values, parse_issues = parse_prefixed_record(
        lock_path,
        SKILL_INVOCATION_LOCK_REQUIRED_PREFIXES,
        "skill invocation lock",
    )
    issues.extend(parse_issues)
    if parse_issues:
        return issues

    if values["- skill name:"] != "graduation-project-builder":
        issues.append("skill invocation lock skill name must be graduation-project-builder")
    if values["- invocation detected:"].lower() not in {"yes", "true", "detected"}:
        issues.append("skill invocation lock must record invocation detected: yes")
    if values["- lock created before mutation?:"].lower() != "yes":
        issues.append("skill invocation lock must be created before mutation")
    if not skill_lock_passish(values["- run start order verdict:"]):
        issues.append("skill invocation lock run start order verdict must be pass-shaped")
    if values["- task mode:"] != task_mode:
        issues.append("skill invocation lock task mode differs from final acceptance record")
    entrypoint = values["- loaded entrypoint:"].lower()
    routed = values["- loaded routed references:"].lower()
    if "skill.md" not in entrypoint or "graduation-project-builder" not in entrypoint:
        issues.append("skill invocation lock must bind the graduation-project-builder SKILL.md entrypoint")
    if "user-feedback-persistence" not in routed:
        issues.append("skill invocation lock must record routed user-feedback-persistence loading")

    for label, value in (
        ("active checklist path", values["- active checklist path:"]),
        ("agent run manifest path", values["- agent run manifest path:"]),
        ("lane task card paths", values["- lane task card paths:"]),
    ):
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"skill invocation lock must record {label} or an explicit fallback path")

    mutation_allowed = values["- mutation allowed verdict:"].lower()
    if contains_any(mutation_allowed, {"blocked", "failed", "fail", "missing", "stale"}):
        issues.append("skill invocation lock mutation allowed verdict is blocking")
    if task_mode in THESIS_MODES and selected_workflow != "audit-only" and not skill_lock_passish(mutation_allowed):
        issues.append("skill invocation lock must allow mutation before a thesis DOCX mutation can pass")
    if selected_workflow == "audit-only" and "audit-only" not in mutation_allowed and not skill_lock_passish(mutation_allowed):
        issues.append("skill invocation lock audit-only workflow must record audit-only or pass mutation allowance")

    risk_text = values["- project-local helper risk count:"]
    try:
        risk_count = int(risk_text)
    except ValueError:
        issues.append("skill invocation lock project-local helper risk count must be numeric")
        risk_count = -1
    disposition = values["- project-local helper disposition:"].lower()
    if risk_count > 0 and disposition not in {
        "clean-source-restart-completed",
        "canonical-helper-replacement-completed",
        "clean",
    }:
        if skill_lock_passish(mutation_allowed):
            issues.append(
                "skill invocation lock cannot allow mutation when risky project-local helpers exist "
                "without a clean-source restart or canonical-helper replacement"
            )

    final_gate_command = values["- final gate command:"]
    if "validate_skill_gate" not in final_gate_command or "--gate-record" not in final_gate_command:
        issues.append("skill invocation lock final gate command must call validate_skill_gate with --gate-record")
    if not skill_lock_passish(values["- final gate verdict:"]):
        issues.append("skill invocation lock final gate verdict must be pass-shaped before handoff")
    if not is_explicit(values["- explicit invocation source type:"]) or is_explicit_none(values["- explicit invocation source type:"]):
        issues.append("skill invocation lock explicit invocation source type must be recorded")
    if not skill_lock_passish(values["- skill activation status:"]) and values["- skill activation status:"].lower() not in {"active", "activated"}:
        issues.append("skill invocation lock skill activation status must be active/pass-shaped")
    if not skill_lock_passish(values["- rule engine takeover verdict:"]):
        issues.append("skill invocation lock rule engine takeover verdict must be pass-shaped")
    if not skill_lock_passish(values["- prohibited bypasses checked:"]):
        issues.append("skill invocation lock prohibited bypasses checked must be pass-shaped")
    if values["- canonical gate required?:"].lower() != "yes":
        issues.append("skill invocation lock canonical gate required must be yes")
    if not skill_lock_noish(values["- narrow/smoke gate substitute used?:"]):
        issues.append("skill invocation lock narrow/smoke gate substitute used must be no")
    if not skill_lock_passish(values["- failed evidence escalation verdict:"]):
        issues.append("skill invocation lock failed evidence escalation verdict must be pass-shaped")
    preflight_helper_value = values["- no project-local thick helper execution before preflight?:"].lower()
    if preflight_helper_value not in {"yes", "not-applicable"}:
        issues.append("skill invocation lock must record no project-local thick helper execution before preflight")
    if values["- no non-control action before lock?:"].lower() != "yes":
        issues.append("skill invocation lock no non-control action before lock must be yes")
    if values["- no mutation before lock?:"].lower() != "yes":
        issues.append("skill invocation lock no mutation before lock must be yes")
    if not skill_lock_passish(values["- final handoff allowed verdict:"]):
        issues.append("skill invocation lock final handoff allowed verdict must be pass-shaped before handoff")
    if not skill_lock_blocked_evidence_ok(values["- blocked evidence disposition:"]):
        issues.append("skill invocation lock blocked evidence disposition must close or escalate blocked evidence, not hide it as caveat/risk")
    final_gate_record_raw = raw_values["- final gate record path:"]
    if is_explicit(values["- final gate record path:"]) and not is_explicit_none(values["- final gate record path:"]):
        final_gate_record_path = resolve_record_path(final_gate_record_raw, lock_path)
        try:
            if final_gate_record_path.resolve() != record_path.resolve():
                issues.append("skill invocation lock final gate record path must match the acceptance record being validated")
        except OSError:
            issues.append("skill invocation lock final gate record path could not be resolved")

    if final_docx_path is not None and is_explicit(values["- exact output path:"]) and not is_explicit_none(values["- exact output path:"]):
        output_path = resolve_record_path(raw_values["- exact output path:"], lock_path)
        try:
            if output_path.resolve() != final_docx_path.resolve():
                issues.append("skill invocation lock exact output path differs from final acceptance exact output DOCX")
        except OSError:
            issues.append("skill invocation lock exact output path could not be resolved")
    output_sha = values["- exact output sha256:"].strip()
    if expected_output_sha256 and len(expected_output_sha256) == 64 and output_sha.lower() != expected_output_sha256.lower():
        issues.append("skill invocation lock exact output sha256 differs from final acceptance reviewed output sha256")

    if not is_explicit(gate_values.get("- skill invocation source type:", "")) or is_explicit_none(gate_values.get("- skill invocation source type:", "")):
        issues.append("gate record skill invocation source type must be recorded for explicit skill invocation")
    for prefix in (
        "- skill invocation verified:",
        "- skill activation status:",
        "- skill takeover verdict:",
        "- skill bypass prevention verdict:",
        "- canonical gate enforcement verdict:",
        "- blocked evidence escalation verdict:",
        "- final handoff allowed by skill lock?:",
        "- routed references verified:",
        "- active checklist verified:",
        "- user request compliance verdict:",
        "- loaded rule compliance verdict:",
    ):
        value = gate_values.get(prefix, "")
        if not skill_lock_passish(value):
            issues.append(f"gate record {prefix} must be pass-shaped for explicit skill invocation")
    if not skill_lock_noish(gate_values.get("- narrow/smoke substitute gate used?:", "")):
        issues.append("gate record narrow/smoke substitute gate used must be no for explicit skill invocation")
    if gate_values.get("- no non-control action before lock?:", "").lower() != "yes":
        issues.append("gate record no non-control action before lock must be yes for explicit skill invocation")
    return issues


def normalized_agent_path_value(raw_value: str, base_path: Path) -> tuple[str, ...]:
    parts = split_path_values(raw_value)
    normalized_parts: list[str] = []
    for part in parts:
        item = part.strip().strip("`'\"")
        if not item:
            continue
        if "=" in item:
            key, value = item.split("=", 1)
            key_text = normalize(key)
            value_text = value.strip().strip("`'\"")
            if value_is_pathlike(value_text):
                value_text = str(resolve_record_path(value_text, base_path).resolve()).lower()
            else:
                value_text = normalize(value_text)
            normalized_parts.append(f"{key_text}={value_text}")
            continue
        if value_is_pathlike(item):
            normalized_parts.append(str(resolve_record_path(item, base_path).resolve()).lower())
        else:
            normalized_parts.append(normalize(item))
    if not normalized_parts and raw_value:
        normalized_parts.append(normalize(raw_value))
    return tuple(sorted(normalized_parts))


def compare_agent_record_field_to_gate(
    issues: list[str],
    *,
    kind: str,
    record_path_for_values: Path,
    record_values: dict[str, str],
    record_raw_values: dict[str, str],
    record_prefix: str,
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    gate_record_path: Path,
) -> None:
    gate_prefix = AGENT_PROTECTED_FIELD_TO_GATE_PREFIX.get(record_prefix)
    if not gate_prefix:
        return
    record_value = record_values.get(record_prefix, "")
    gate_value = gate_values.get(gate_prefix, "")
    if record_prefix in AGENT_PROTECTED_PATH_PREFIXES:
        record_comparable = normalized_agent_path_value(
            record_raw_values.get(record_prefix, record_value),
            record_path_for_values,
        )
        gate_comparable = normalized_agent_path_value(
            gate_raw_values.get(gate_prefix, gate_value),
            gate_record_path,
        )
        if record_comparable != gate_comparable:
            issues.append(f"{kind} {record_prefix} differs from final acceptance record {gate_prefix}")
        return
    if normalize(record_value) != normalize(gate_value):
        issues.append(f"{kind} {record_prefix} differs from final acceptance record {gate_prefix}")


def validate_agent_records(
    *,
    record_path: Path,
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    rendered_docx_path: Path | None,
) -> list[str]:
    issues: list[str] = []
    manifest_raw = gate_raw_values.get("- agent run manifest path:", "")
    cards_raw = gate_raw_values.get("- lane task card evidence paths:", "")
    if is_explicit_none(gate_values.get("- agent run manifest path:", "")) or not is_explicit(manifest_raw):
        return issues
    if is_explicit_none(gate_values.get("- lane task card evidence paths:", "")) or not is_explicit(cards_raw):
        return issues

    manifest_path = resolve_record_path(manifest_raw, record_path)
    issues.extend(validate_existing_path(manifest_path, require_nonempty_file=True))
    if issues:
        return issues
    manifest_values, manifest_raw_values, manifest_issues = parse_prefixed_record(
        manifest_path,
        AGENT_RUN_MANIFEST_REQUIRED_PREFIXES,
        "agent run manifest",
    )
    issues.extend(manifest_issues)
    if not manifest_issues:
        if manifest_values["- agent_mode:"] != gate_values.get("- agent mode:", ""):
            issues.append("agent run manifest agent_mode differs from final acceptance record")
        if manifest_values["- authorization_source:"] != gate_values.get("- agent authorization source:", ""):
            issues.append("agent run manifest authorization_source differs from final acceptance record")
        concurrency_pairs = (
            ("- max_concurrent_live_agents:", "- max concurrent live agents:"),
            ("- live_agent_count_plan:", "- live agent count plan:"),
            ("- dispatch_wave_plan:", "- dispatch wave plan:"),
            ("- audit_presence_by_wave:", "- audit presence by wave:"),
            ("- concurrency_limit_verdict:", "- concurrency limit verdict:"),
        )
        for manifest_prefix, gate_prefix in concurrency_pairs:
            if manifest_values[manifest_prefix] != gate_values.get(gate_prefix, ""):
                issues.append(f"agent run manifest {manifest_prefix} differs from final acceptance record")
        issues.extend(
            validate_agent_concurrency_fields(
                max_live_agents=manifest_values["- max_concurrent_live_agents:"],
                live_agent_count_plan=manifest_values["- live_agent_count_plan:"],
                dispatch_wave_plan=manifest_values["- dispatch_wave_plan:"],
                audit_presence_by_wave=manifest_values["- audit_presence_by_wave:"],
                concurrency_limit_verdict=manifest_values["- concurrency_limit_verdict:"],
                kind="agent run manifest",
            )
        )
        if manifest_values["- audit_role_alias_zh:"] != "\u5ba1\u6838":
            issues.append("agent run manifest audit_role_alias_zh must be \u5ba1\u6838")
        if manifest_values["- audit_verdict:"] != "pass":
            issues.append("agent run manifest audit_verdict must be pass")
        manifest_lines = read_optional_text(manifest_path).splitlines()
        manifest_consistency_values: dict[str, str] = {}
        for manifest_prefix, gate_prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX.items():
            if manifest_prefix in manifest_values:
                manifest_consistency_values[gate_prefix] = manifest_values[manifest_prefix]
        for raw_prefix, gate_prefix in (
            ("- handoff_status:", "- handoff status:"),
            ("- handoff status:", "- handoff status:"),
            ("- blockers:", "- blockers:"),
            ("- blocker_summary:", "- blockers:"),
            ("- known caveats:", "- known caveats:"),
            ("- validation result:", "- validation result:"),
        ):
            matches = find_lines_with_prefix(manifest_lines, raw_prefix)
            if matches:
                manifest_consistency_values[gate_prefix] = parse_line_value(matches[0])
        manifest_consistency_values.setdefault("- audit verdict:", manifest_values["- audit_verdict:"])
        issues.extend(validate_blocking_state_consistency(manifest_consistency_values, location=manifest_path))
        for manifest_prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX:
            if manifest_prefix not in manifest_values:
                continue
            compare_agent_record_field_to_gate(
                issues,
                kind="agent run manifest",
                record_path_for_values=manifest_path,
                record_values=manifest_values,
                record_raw_values=manifest_raw_values,
                record_prefix=manifest_prefix,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                gate_record_path=record_path,
            )
        for lane in CANONICAL_ROLE_LANES:
            if lane not in manifest_values.get("- required_lanes:", ""):
                issues.append(f"agent run manifest required_lanes must include full canonical role roster: missing {lane}")
            if lane not in manifest_values.get("- complete_role_roster:", ""):
                issues.append(f"agent run manifest complete_role_roster missing canonical lane: {lane}")
            if lane not in manifest_values.get("- role_attendance_matrix:", ""):
                issues.append(f"agent run manifest role_attendance_matrix missing canonical lane: {lane}")
        for alias in CANONICAL_ROLE_ALIASES:
            if alias not in manifest_values.get("- complete_role_roster:", ""):
                issues.append(f"agent run manifest complete_role_roster missing role alias: {alias}")

    card_paths = [resolve_record_path(raw_path, record_path) for raw_path in split_path_values(cards_raw)]
    manifest_card_paths = {
        str(resolve_record_path(raw_path, manifest_path))
        for raw_path in split_path_values(manifest_raw_values.get("- lane_task_card_paths:", ""))
    }
    seen_lanes: set[str] = set()
    for card_path in card_paths:
        issues.extend(validate_existing_path(card_path, require_nonempty_file=True))
        if not card_path.exists():
            continue
        if manifest_card_paths and str(card_path.resolve()) not in {str(Path(p).resolve()) for p in manifest_card_paths}:
            issues.append(f"lane task card is not listed by agent run manifest: {card_path}")
        card_values, card_raw_values, card_issues = parse_prefixed_record(
            card_path,
            AGENT_TASK_CARD_REQUIRED_PREFIXES,
            "agent task card",
        )
        issues.extend(card_issues)
        if card_issues:
            continue
        lane = card_values["- lane:"]
        seen_lanes.add(lane)
        if card_values["- run_manifest_path:"]:
            card_manifest = resolve_record_path(card_raw_values["- run_manifest_path:"], card_path)
            if card_manifest.resolve() != manifest_path.resolve():
                issues.append(f"agent task card run_manifest_path does not match final acceptance manifest: {card_path}")
        if card_values["- agent_mode:"] != gate_values.get("- agent mode:", ""):
            issues.append(f"agent task card agent_mode differs from final acceptance record: {card_path}")
        if card_values["- authorization_source:"] != gate_values.get("- agent authorization source:", ""):
            issues.append(f"agent task card authorization_source differs from final acceptance record: {card_path}")
        if card_values["- audit_agent_alias_zh:"] != "\u5ba1\u6838":
            issues.append(f"agent task card audit_agent_alias_zh must be \u5ba1\u6838: {card_path}")
        if card_values["- audit_verdict:"] != "pass":
            issues.append(f"agent task card audit_verdict must be pass: {card_path}")
        if card_values["- attendance_status:"] not in {"active", "not-applicable", "skipped-with-reason", "blocked"}:
            issues.append(f"agent task card attendance_status must be active/not-applicable/skipped-with-reason/blocked: {card_path}")
        if card_values["- attendance_status:"] in {"not-applicable", "skipped-with-reason"} and not (is_explicit(card_values["- not_applicable_reason:"]) or is_explicit(card_values["- skip_reason:"])):
            issues.append(f"inactive agent task card must record not_applicable_reason or skip_reason: {card_path}")
        for card_prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX:
            if card_prefix not in card_values:
                continue
            compare_agent_record_field_to_gate(
                issues,
                kind=f"agent task card {card_path}",
                record_path_for_values=card_path,
                record_values=card_values,
                record_raw_values=card_raw_values,
                record_prefix=card_prefix,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                gate_record_path=record_path,
            )

        if card_values["- attendance_status:"] != "active":
            continue
        is_format_card = "format" in lane.lower() or card_values["- role_alias_zh:"] in {"\u683c\u5f0f", "format"} or card_values["- lane_alias_zh:"] in {"\u683c\u5f0f", "format"}
        if not is_format_card:
            continue
        format_values, format_raw_values, format_issues = parse_prefixed_record(
            card_path,
            FORMAT_TASK_CARD_REQUIRED_PREFIXES,
            "format agent task card",
        )
        issues.extend(format_issues)
        if format_issues:
            continue
        card_lines = read_optional_text(card_path).splitlines()
        card_consistency_values: dict[str, str] = {}
        for format_prefix, gate_prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX.items():
            if format_prefix in format_values:
                card_consistency_values[gate_prefix] = format_values[format_prefix]
        for raw_prefix, gate_prefix in (
            ("- status:", "- handoff status:"),
            ("- blockers:", "- blockers:"),
            ("- known caveats:", "- known caveats:"),
            ("- validation result:", "- validation result:"),
        ):
            matches = find_lines_with_prefix(card_lines, raw_prefix)
            if matches:
                card_consistency_values[gate_prefix] = parse_line_value(matches[0])
        card_consistency_values.setdefault("- audit verdict:", card_values["- audit_verdict:"])
        issues.extend(validate_blocking_state_consistency(card_consistency_values, location=card_path))
        for card_prefix, gate_prefix in FORMAT_CARD_TO_GATE_PREFIX.items():
            card_value = format_values[card_prefix]
            gate_value = gate_values.get(gate_prefix, "")
            if card_value and gate_value and card_value != gate_value:
                issues.append(f"format task card differs from final acceptance record: {card_prefix} vs {gate_prefix}")
        for prefix in FORMAT_TASK_CARD_REQUIRED_PREFIXES:
            if format_values[prefix] in EXPLICIT_VALUES or not is_explicit(format_values[prefix]):
                issues.append(f"format task card field must not be blank or none: {card_path} ({prefix})")
        for format_prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX:
            if format_prefix not in format_values:
                continue
            compare_agent_record_field_to_gate(
                issues,
                kind=f"format task card {card_path}",
                record_path_for_values=card_path,
                record_values=format_values,
                record_raw_values=format_raw_values,
                record_prefix=format_prefix,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                gate_record_path=record_path,
            )
        if format_values["- template_alignment_verdict:"] != "pass":
            issues.append(f"format task card template_alignment_verdict must be pass: {card_path}")
        reviewed_output = resolve_record_path(format_raw_values["- reviewed_output_path:"], card_path)
        if rendered_docx_path is not None and reviewed_output.resolve() != rendered_docx_path.resolve():
            issues.append(f"format task card reviewed_output_path does not match rendered/final DOCX: {card_path}")
        if reviewed_output.exists():
            reviewed_sha = format_values["- reviewed_output_sha256:"]
            if len(reviewed_sha) != 64 or any(ch not in "0123456789abcdefABCDEF" for ch in reviewed_sha):
                issues.append(f"format task card reviewed_output_sha256 must be a 64-hex value: {card_path}")
            elif sha256_file(reviewed_output).lower() != reviewed_sha.lower():
                issues.append(f"format task card reviewed_output_sha256 does not match reviewed output: {card_path}")

    required_lanes = gate_values.get("- required lanes:", "")
    final_roster = gate_values.get("- complete role roster:", "")
    final_attendance = gate_values.get("- role attendance matrix:", "")
    final_inactive_reasons = gate_values.get("- not-applicable lanes with reasons:", "")
    all_role_paths = gate_values.get("- all role task card paths:", "")
    for lane in CANONICAL_ROLE_LANES:
        if lane not in required_lanes:
            issues.append(f"final acceptance required lanes must include full canonical role roster: missing {lane}")
        if lane not in final_roster:
            issues.append(f"final acceptance complete role roster missing canonical lane: {lane}")
        if lane not in final_attendance:
            issues.append(f"final acceptance role attendance matrix missing canonical lane: {lane}")
        if not any(lane == seen or lane in seen for seen in seen_lanes):
            issues.append(f"final acceptance requires canonical role task card but none was found: {lane}")
        if lane not in all_role_paths:
            issues.append(f"final acceptance all role task card paths must include canonical lane: {lane}")
    for alias in CANONICAL_ROLE_ALIASES:
        if alias not in final_roster:
            issues.append(f"final acceptance complete role roster missing role alias: {alias}")
        if alias not in gate_values.get("- agent role aliases zh:", "") and alias not in gate_values.get("- required lane aliases zh:", ""):
            issues.append(f"final acceptance missing Chinese role alias: {alias}")
    inactive_status_tokens = ("not-applicable", "skipped-with-reason")
    if any(token in final_attendance for token in inactive_status_tokens) and not is_explicit(final_inactive_reasons):
        issues.append("final acceptance inactive role attendance must include not-applicable lanes with reasons")
    if gate_values.get("- audit full roster verdict:", "") != "pass":
        issues.append("final acceptance audit full roster verdict must be pass")
    return issues


def value_is_pathlike(value: str) -> bool:
    lowered = value.lower()
    return any(ext in lowered for ext in (".md", ".txt", ".json", ".png", ".jpg", ".jpeg", ".pdf"))


def extract_pathlike_candidates(value: str) -> list[str]:
    candidates: list[str] = []
    for part in split_path_values(value):
        part = part.strip().strip("`'\"")
        if value_is_pathlike(part) and not is_explicit_none(normalize(part)):
            candidates.append(part)
    path_pattern = re.compile(
        r"(?P<path>(?:[A-Za-z]:\\|/)[^;\n\r\t|`\"']+?\.(?:md|txt|json|png|jpe?g|pdf))",
        re.IGNORECASE,
    )
    for match in path_pattern.finditer(value):
        candidate = match.group("path").strip()
        if candidate not in candidates:
            candidates.append(candidate)
    return candidates


def iter_json_dict_nodes(value: object):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from iter_json_dict_nodes(child)
    elif isinstance(value, list):
        for child in value:
            yield from iter_json_dict_nodes(child)


def node_matches_alias(node: dict[str, object], aliases: tuple[str, ...]) -> bool:
    text = json.dumps(node, ensure_ascii=False, sort_keys=True).lower()
    return any(alias in text for alias in aliases)


def node_has_pass_verdict(node: dict[str, object]) -> bool:
    for key in ("verdict", "result", "status", "passed"):
        if key not in node:
            continue
        value = normalize(str(node.get(key, ""))).lower()
        if key == "passed":
            if value == "true":
                return True
            continue
        if value in PAGE_CLASS_PASS_TOKENS:
            return True
    return False


def node_has_evidence_path(node: dict[str, object]) -> bool:
    for key, value in node.items():
        lowered_key = str(key).lower()
        if "path" in lowered_key or "evidence" in lowered_key or "artifact" in lowered_key:
            if value_is_pathlike(str(value)) and not is_explicit_none(normalize(str(value))):
                return True
    return False


def iter_node_evidence_paths(node: dict[str, object]) -> list[str]:
    paths: list[str] = []
    for key, value in node.items():
        lowered_key = str(key).lower()
        if "path" not in lowered_key and "evidence" not in lowered_key and "artifact" not in lowered_key:
            continue
        paths.extend(extract_pathlike_candidates(str(value)))
    return paths


def validate_page_class_evidence_paths(raw_paths: list[str], matrix_path: Path, acceptance_mode: str) -> list[str]:
    issues: list[str] = []
    for raw_path in raw_paths:
        resolved = resolve_record_path(raw_path, matrix_path)
        path_issues = validate_existing_path(resolved, require_nonempty_file=True)
        issues.extend(path_issues)
        if path_issues:
            continue
        if resolved.suffix.lower() in {".md", ".txt"}:
            issues.extend(
                check_review_evidence_record(
                    resolved,
                    expected_type="thesis-rendered-page",
                    acceptance_mode=acceptance_mode,
                )
            )
    return issues


def validate_json_page_class_entries(payload: object, matrix_path: Path, acceptance_mode: str) -> list[str]:
    issues: list[str] = []
    nodes = list(iter_json_dict_nodes(payload))
    evidence_path_labels: dict[str, set[str]] = {}
    for label, aliases in PAGE_CLASS_ALIASES.items():
        class_nodes = [node for node in nodes if node_matches_alias(node, aliases)]
        if not class_nodes:
            issues.append(f"page-class coverage matrix lacks a per-class row for {label}: {matrix_path}")
            continue
        if not any(node_has_pass_verdict(node) for node in class_nodes):
            issues.append(f"page-class coverage matrix lacks an exact pass verdict for {label}: {matrix_path}")
        if not any(node_has_evidence_path(node) for node in class_nodes):
            issues.append(f"page-class coverage matrix lacks a per-class evidence path for {label}: {matrix_path}")
        if not any(str(node.get("target_identifier") or "").strip() for node in class_nodes):
            issues.append(f"page-class coverage matrix lacks target_identifier for {label}: {matrix_path}")
        if not any(str(node.get("rendered_page_or_region") or node.get("target_region") or "").strip() for node in class_nodes):
            issues.append(f"page-class coverage matrix lacks rendered_page_or_region for {label}: {matrix_path}")
        class_paths: list[str] = []
        for node in class_nodes:
            class_paths.extend(iter_node_evidence_paths(node))
        issues.extend(validate_page_class_evidence_paths(class_paths, matrix_path, acceptance_mode))
        for raw_path in class_paths:
            resolved = str(resolve_record_path(raw_path, matrix_path))
            evidence_path_labels.setdefault(resolved, set()).add(label)
    for resolved, labels in sorted(evidence_path_labels.items()):
        if len(labels) <= 1:
            continue
        text = read_optional_text(Path(resolved)).lower()
        if not all(label.lower() in text for label in labels):
            issues.append(
                "page-class coverage matrix reuses one evidence path for multiple page classes without class-specific rows: "
                f"{resolved} -> {', '.join(sorted(labels))}"
            )
    return issues


def validate_text_page_class_entries(text: str, matrix_path: Path, acceptance_mode: str) -> list[str]:
    issues: list[str] = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    lowered_lines = [line.lower() for line in lines]
    for label, aliases in PAGE_CLASS_ALIASES.items():
        matched_lines = [
            lowered
            for lowered in lowered_lines
            if any(alias in lowered for alias in aliases)
        ]
        if not matched_lines:
            issues.append(f"page-class coverage matrix lacks a per-class row for {label}: {matrix_path}")
            continue
        pass_lines = [
            line
            for line in matched_lines
            if re.search(r"\bpass(?:ed)?\b", line)
            and not contains_any(line, PAGE_CLASS_BAD_TOKENS)
        ]
        if not pass_lines:
            issues.append(f"page-class coverage matrix lacks an exact pass verdict for {label}: {matrix_path}")
        if not any(value_is_pathlike(line) for line in matched_lines):
            issues.append(f"page-class coverage matrix lacks a per-class evidence path for {label}: {matrix_path}")
        class_paths: list[str] = []
        for line in matched_lines:
            class_paths.extend(extract_pathlike_candidates(line))
        issues.extend(validate_page_class_evidence_paths(class_paths, matrix_path, acceptance_mode))
    return issues


def inventory_node_text(node: dict[str, object]) -> str:
    return json.dumps(node, ensure_ascii=False, sort_keys=True).lower()


def inventory_status_from_text(text: str) -> str:
    for status in sorted(MANDATORY_THESIS_SURFACE_ALLOWED_STATUSES, key=len, reverse=True):
        if status in text:
            return status
    return ""


def inventory_node_has_reason(node: dict[str, object]) -> bool:
    for key, value in node.items():
        if "reason" in str(key).lower():
            reason = normalize(str(value))
            return is_explicit(reason) and not is_explicit_none(reason)
    return False


def inventory_text_row_has_reason(row: str, status: str) -> bool:
    cells = inventory_text_row_cells(row)
    if len(cells) >= 6:
        reason = normalize(cells[5])
        return is_explicit(reason) and not is_explicit_none(reason)
    lowered = row.lower()
    if "reason" in lowered:
        after_reason = re.split(r"\breason\b\s*[:=| -]*", row, maxsplit=1, flags=re.I)
        if len(after_reason) == 2:
            return is_explicit(normalize(after_reason[1])) and not is_explicit_none(normalize(after_reason[1]))
        return True
    after_status = lowered.split(status, 1)[1] if status in lowered else ""
    return len(normalize(after_status)) >= 8 and not is_explicit_none(normalize(after_status))


def inventory_text_row_cells(row: str) -> list[str]:
    if "|" not in row:
        return []
    return [cell.strip() for cell in row.strip().strip("|").split("|")]


def inventory_text_row_surface_text(row: str) -> str:
    cells = inventory_text_row_cells(row)
    return cells[0].lower() if cells else row.lower()


def inventory_text_row_evidence_text(row: str) -> str:
    cells = inventory_text_row_cells(row)
    return cells[3] if len(cells) >= 4 else row


def inventory_text_row_verdict_text(row: str) -> str:
    cells = inventory_text_row_cells(row)
    return cells[4] if len(cells) >= 5 else row


def inventory_row_passes(row_text: str) -> bool:
    lowered = inventory_text_row_verdict_text(row_text).lower()
    return bool(re.search(r"\bpass(?:ed)?\b", lowered)) and not contains_any(
        lowered,
        {
            "fail",
            "failed",
            "missing",
            "not checked",
            "pending",
            "blocked",
            "unresolved",
            "sampled-only",
            "sample only",
            "drift",
        },
    )


STRICT_FRONT_MATTER_EVIDENCE_SURFACES = set(MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES)


def validate_inventory_evidence_paths(
    raw_paths: list[str],
    inventory_path: Path,
    *,
    surface_id: str | None = None,
    acceptance_mode: str = "format-repair-only",
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    strict_required = surface_id in STRICT_FRONT_MATTER_EVIDENCE_SURFACES
    strict_text_seen = False
    strict_pass_seen = False
    strict_failures: list[str] = []
    for raw_path in raw_paths:
        resolved = resolve_record_path(raw_path, inventory_path)
        path_issues = validate_existing_path(resolved, require_nonempty_file=True)
        issues.extend(path_issues)
        if path_issues or not strict_required:
            continue
        if resolved.suffix.lower() not in {".md", ".txt"}:
            continue
        strict_text_seen = True
        current_issues = check_required_surface_evidence_record(
            resolved,
            expected_type="thesis-rendered-page",
            acceptance_mode=acceptance_mode,
            required_surface_id=surface_id or "",
            expected_reviewed_output=expected_reviewed_output,
            expected_reviewed_sha256=expected_reviewed_sha256,
        )
        if current_issues:
            strict_failures.extend(current_issues)
        else:
            strict_pass_seen = True
    if strict_required and not strict_pass_seen:
        if not strict_text_seen:
            issues.append(
                f"{surface_id} evidence must include a text review evidence record with surface-level baseline comparison: {inventory_path}"
            )
        issues.extend(strict_failures)
    return issues


def validate_json_thesis_surface_inventory(
    payload: object,
    inventory_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    nodes = list(iter_json_dict_nodes(payload))
    for surface_id, aliases in MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES.items():
        surface_nodes = [node for node in nodes if node_matches_alias(node, (surface_id, *aliases))]
        if not surface_nodes:
            issues.append(f"mandatory thesis surface inventory missing required surface row: {surface_id}")
            continue
        valid_row_seen = False
        for node in surface_nodes:
            row_text = inventory_node_text(node)
            status = inventory_status_from_text(row_text)
            if not status:
                continue
            valid_row_seen = True
            if status == "blocked":
                issues.append(f"mandatory thesis surface inventory row is blocked: {surface_id}")
            if surface_id in MANDATORY_THESIS_CRITICAL_PRESENT_SURFACES and status in {
                "not-present",
                "not-applicable-with-reason",
            }:
                issues.append(f"critical thesis surface cannot be marked absent in acceptance inventory: {surface_id}")
            if status.startswith("present"):
                if not node_has_evidence_path(node):
                    issues.append(f"mandatory thesis surface inventory present row lacks evidence path: {surface_id}")
                if not inventory_row_passes(row_text):
                    issues.append(f"mandatory thesis surface inventory present row lacks pass verdict: {surface_id}")
                issues.extend(
                    validate_inventory_evidence_paths(
                        iter_node_evidence_paths(node),
                        inventory_path,
                        surface_id=surface_id,
                        acceptance_mode=acceptance_mode,
                        expected_reviewed_output=expected_reviewed_output,
                        expected_reviewed_sha256=expected_reviewed_sha256,
                    )
                )
            elif not inventory_node_has_reason(node):
                issues.append(f"mandatory thesis surface inventory absent/not-applicable row lacks reason: {surface_id}")
        if not valid_row_seen:
            issues.append(f"mandatory thesis surface inventory row lacks allowed status: {surface_id}")
    return issues


def validate_text_thesis_surface_inventory(
    text: str,
    inventory_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    rows = [
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    lowered_rows = [(row, row.lower()) for row in rows]
    for surface_id, aliases in MANDATORY_THESIS_SURFACE_INVENTORY_ALIASES.items():
        search_aliases = tuple(alias.lower() for alias in (surface_id, *aliases))
        matched_rows = [
            row
            for row, lowered in lowered_rows
            if any(alias in inventory_text_row_surface_text(row) for alias in search_aliases)
        ]
        if not matched_rows:
            issues.append(f"mandatory thesis surface inventory missing required surface row: {surface_id}")
            continue
        row = matched_rows[0]
        lowered = row.lower()
        status = inventory_status_from_text(lowered)
        if not status:
            issues.append(f"mandatory thesis surface inventory row lacks allowed status: {surface_id}")
            continue
        if status == "blocked":
            issues.append(f"mandatory thesis surface inventory row is blocked: {surface_id}")
        if surface_id in MANDATORY_THESIS_CRITICAL_PRESENT_SURFACES and status in {
            "not-present",
            "not-applicable-with-reason",
        }:
            issues.append(f"critical thesis surface cannot be marked absent in acceptance inventory: {surface_id}")
        if status.startswith("present"):
            if not inventory_row_passes(row):
                issues.append(f"mandatory thesis surface inventory present row lacks pass verdict: {surface_id}")
            raw_paths = extract_pathlike_candidates(inventory_text_row_evidence_text(row))
            if not raw_paths:
                issues.append(f"mandatory thesis surface inventory present row lacks evidence path: {surface_id}")
            issues.extend(
                validate_inventory_evidence_paths(
                    raw_paths,
                    inventory_path,
                    surface_id=surface_id,
                    acceptance_mode=acceptance_mode,
                    expected_reviewed_output=expected_reviewed_output,
                    expected_reviewed_sha256=expected_reviewed_sha256,
                )
            )
        elif not inventory_text_row_has_reason(row, status):
            issues.append(f"mandatory thesis surface inventory absent/not-applicable row lacks reason: {surface_id}")
    return issues


def validate_thesis_surface_inventory(
    inventory_path: Path,
    record_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(inventory_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(inventory_path)
    lowered = text.lower()
    if contains_any(
        lowered,
        {
            "sampled-only",
            "sample only",
            "not checked",
            "pending",
            "unresolved",
        },
    ) or re.search(r"\b(fail|failed)\b(?!\s*[:=]\s*false)", lowered):
        issues.append(f"mandatory thesis surface inventory records unresolved coverage in {record_path}")
    payload: object | None = None
    if inventory_path.suffix.lower() == ".json":
        try:
            payload = json.loads(text)
        except Exception as exc:
            issues.append(f"mandatory thesis surface inventory is not valid JSON: {inventory_path} ({exc})")
    if payload is not None:
        issues.extend(
            validate_json_thesis_surface_inventory(
                payload,
                inventory_path,
                acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    else:
        issues.extend(
            validate_text_thesis_surface_inventory(
                text,
                inventory_path,
                acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_json_high_risk_surface_matrix(
    payload: object,
    matrix_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    nodes = list(iter_json_dict_nodes(payload))
    for surface_id, aliases in HIGH_RISK_THESIS_FORMAT_SURFACE_ALIASES.items():
        surface_nodes = [node for node in nodes if node_matches_alias(node, (surface_id, *aliases))]
        if not surface_nodes:
            issues.append(f"high-risk thesis format surface matrix missing required surface row: {surface_id}")
            continue
        valid_row_seen = False
        for node in surface_nodes:
            row_text = inventory_node_text(node)
            status = inventory_status_from_text(row_text)
            if not status:
                continue
            valid_row_seen = True
            if status == "blocked":
                issues.append(f"high-risk thesis format surface matrix row is blocked: {surface_id}")
            if status in {"not-present", "not-applicable-with-reason"}:
                if surface_id not in HIGH_RISK_OPTIONAL_WITH_REASON_SURFACES:
                    issues.append(f"high-risk thesis format surface cannot be marked absent: {surface_id}")
                elif not inventory_node_has_reason(node):
                    issues.append(f"high-risk appendix surface row lacks not-applicable reason: {surface_id}")
                continue
            if status.startswith("present"):
                if not node_has_evidence_path(node):
                    issues.append(f"high-risk thesis format surface row lacks evidence path: {surface_id}")
                if not inventory_row_passes(row_text):
                    issues.append(f"high-risk thesis format surface row lacks pass verdict: {surface_id}")
                issues.extend(
                    validate_inventory_evidence_paths(
                        iter_node_evidence_paths(node),
                        matrix_path,
                        surface_id=surface_id,
                        acceptance_mode=acceptance_mode,
                        expected_reviewed_output=expected_reviewed_output,
                        expected_reviewed_sha256=expected_reviewed_sha256,
                    )
                )
        if not valid_row_seen:
            issues.append(f"high-risk thesis format surface matrix row lacks allowed status: {surface_id}")
    return issues


def validate_text_high_risk_surface_matrix(
    text: str,
    matrix_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    rows = [
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    lowered_rows = [(row, row.lower()) for row in rows]
    for surface_id, aliases in HIGH_RISK_THESIS_FORMAT_SURFACE_ALIASES.items():
        search_aliases = tuple(alias.lower() for alias in (surface_id, *aliases))
        matched_rows = [
            row
            for row, lowered in lowered_rows
            if any(alias in inventory_text_row_surface_text(row) for alias in search_aliases)
        ]
        if not matched_rows:
            issues.append(f"high-risk thesis format surface matrix missing required surface row: {surface_id}")
            continue
        row = matched_rows[0]
        lowered = row.lower()
        status = inventory_status_from_text(lowered)
        if not status:
            issues.append(f"high-risk thesis format surface matrix row lacks allowed status: {surface_id}")
            continue
        if status == "blocked":
            issues.append(f"high-risk thesis format surface matrix row is blocked: {surface_id}")
        if status in {"not-present", "not-applicable-with-reason"}:
            if surface_id not in HIGH_RISK_OPTIONAL_WITH_REASON_SURFACES:
                issues.append(f"high-risk thesis format surface cannot be marked absent: {surface_id}")
            elif not inventory_text_row_has_reason(row, status):
                issues.append(f"high-risk appendix surface row lacks not-applicable reason: {surface_id}")
            continue
        if status.startswith("present"):
            if not inventory_row_passes(row):
                issues.append(f"high-risk thesis format surface row lacks pass verdict: {surface_id}")
            raw_paths = extract_pathlike_candidates(inventory_text_row_evidence_text(row))
            if not raw_paths:
                issues.append(f"high-risk thesis format surface row lacks evidence path: {surface_id}")
            issues.extend(
                validate_inventory_evidence_paths(
                    raw_paths,
                    matrix_path,
                    surface_id=surface_id,
                    acceptance_mode=acceptance_mode,
                    expected_reviewed_output=expected_reviewed_output,
                    expected_reviewed_sha256=expected_reviewed_sha256,
                )
            )
    return issues


def validate_high_risk_surface_matrix(
    matrix_path: Path,
    record_path: Path,
    acceptance_mode: str,
    *,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(matrix_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(matrix_path)
    lowered = text.lower()
    if contains_any(
        lowered,
        {
            "sampled-only",
            "sample only",
            "not checked",
            "pending",
            "unresolved",
        },
    ) or re.search(r"\b(fail|failed)\b(?!\s*[:=]\s*false)", lowered):
        issues.append(f"high-risk thesis format surface matrix records unresolved coverage in {record_path}")
    payload: object | None = None
    if matrix_path.suffix.lower() == ".json":
        try:
            payload = json.loads(text)
        except Exception as exc:
            issues.append(f"high-risk thesis format surface matrix is not valid JSON: {matrix_path} ({exc})")
    if payload is not None:
        issues.extend(
            validate_json_high_risk_surface_matrix(
                payload,
                matrix_path,
                acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    else:
        issues.extend(
            validate_text_high_risk_surface_matrix(
                text,
                matrix_path,
                acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_surface_coverage_node(
    node: dict[str, object],
    matrix_path: Path,
    *,
    matrix_label: str,
    surface_id: str,
    acceptance_mode: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    row_text = inventory_node_text(node)
    status = inventory_status_from_text(row_text)
    if not status:
        issues.append(f"{matrix_label} surface coverage matrix row lacks allowed status: {surface_id}")
        return issues
    if status == "blocked":
        issues.append(f"{matrix_label} surface coverage matrix row is blocked: {surface_id}")
    if status in {"not-present", "not-applicable-with-reason"}:
        if surface_id not in SURFACE_COVERAGE_OPTIONAL_WITH_REASON:
            issues.append(f"{matrix_label} surface cannot be marked absent: {surface_id}")
        elif not inventory_node_has_reason(node):
            issues.append(f"{matrix_label} appendix surface row lacks not-applicable reason: {surface_id}")
        return issues
    if status.startswith("present"):
        if not node_has_evidence_path(node):
            issues.append(f"{matrix_label} surface coverage matrix row lacks evidence path: {surface_id}")
        if not inventory_row_passes(row_text):
            issues.append(f"{matrix_label} surface coverage matrix row lacks pass verdict: {surface_id}")
        issues.extend(
            validate_inventory_evidence_paths(
                iter_node_evidence_paths(node),
                matrix_path,
                surface_id=surface_id,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_text_surface_coverage_row(
    row: str,
    matrix_path: Path,
    *,
    matrix_label: str,
    surface_id: str,
    acceptance_mode: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    lowered = row.lower()
    status = inventory_status_from_text(lowered)
    if not status:
        issues.append(f"{matrix_label} surface coverage matrix row lacks allowed status: {surface_id}")
        return issues
    if status == "blocked":
        issues.append(f"{matrix_label} surface coverage matrix row is blocked: {surface_id}")
    if status in {"not-present", "not-applicable-with-reason"}:
        if surface_id not in SURFACE_COVERAGE_OPTIONAL_WITH_REASON:
            issues.append(f"{matrix_label} surface cannot be marked absent: {surface_id}")
        elif not inventory_text_row_has_reason(row, status):
            issues.append(f"{matrix_label} appendix surface row lacks not-applicable reason: {surface_id}")
        return issues
    if status.startswith("present"):
        if not inventory_row_passes(row):
            issues.append(f"{matrix_label} surface coverage matrix row lacks pass verdict: {surface_id}")
        raw_paths = extract_pathlike_candidates(inventory_text_row_evidence_text(row))
        if not raw_paths:
            issues.append(f"{matrix_label} surface coverage matrix row lacks evidence path: {surface_id}")
        issues.extend(
            validate_inventory_evidence_paths(
                raw_paths,
                matrix_path,
                surface_id=surface_id,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_json_surface_coverage_matrix(
    payload: object,
    matrix_path: Path,
    *,
    matrix_label: str,
    required_aliases: dict[str, tuple[str, ...]],
    acceptance_mode: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    nodes = list(iter_json_dict_nodes(payload))
    for surface_id, aliases in required_aliases.items():
        surface_nodes = [node for node in nodes if node_matches_alias(node, (surface_id, *aliases))]
        if not surface_nodes:
            issues.append(f"{matrix_label} surface coverage matrix missing required surface row: {surface_id}")
            continue
        issues.extend(
            validate_surface_coverage_node(
                surface_nodes[0],
                matrix_path,
                matrix_label=matrix_label,
                surface_id=surface_id,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_text_surface_coverage_matrix(
    text: str,
    matrix_path: Path,
    *,
    matrix_label: str,
    required_aliases: dict[str, tuple[str, ...]],
    acceptance_mode: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    rows = [
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    lowered_rows = [(row, row.lower()) for row in rows]
    for surface_id, aliases in required_aliases.items():
        search_aliases = tuple(alias.lower() for alias in (surface_id, *aliases))
        matched_rows = [
            row
            for row, _lowered in lowered_rows
            if any(alias in inventory_text_row_surface_text(row) for alias in search_aliases)
        ]
        if not matched_rows:
            issues.append(f"{matrix_label} surface coverage matrix missing required surface row: {surface_id}")
            continue
        issues.extend(
            validate_text_surface_coverage_row(
                matched_rows[0],
                matrix_path,
                matrix_label=matrix_label,
                surface_id=surface_id,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def validate_surface_coverage_matrix(
    matrix_path: Path,
    record_path: Path,
    *,
    matrix_label: str,
    required_aliases: dict[str, tuple[str, ...]],
    acceptance_mode: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(matrix_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(matrix_path)
    lowered = text.lower()
    if contains_any(
        lowered,
        {
            "sampled-only",
            "sample only",
            "not checked",
            "pending",
            "unresolved",
        },
    ) or re.search(r"\b(fail|failed)\b(?!\s*[:=]\s*false)", lowered):
        issues.append(f"{matrix_label} surface coverage matrix records unresolved coverage in {record_path}")
    payload: object | None = None
    if matrix_path.suffix.lower() == ".json":
        try:
            payload = json.loads(text)
        except Exception as exc:
            issues.append(f"{matrix_label} surface coverage matrix is not valid JSON: {matrix_path} ({exc})")
    if payload is not None:
        issues.extend(
            validate_json_surface_coverage_matrix(
                payload,
                matrix_path,
                matrix_label=matrix_label,
                required_aliases=required_aliases,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    else:
        issues.extend(
            validate_text_surface_coverage_matrix(
                text,
                matrix_path,
                matrix_label=matrix_label,
                required_aliases=required_aliases,
                acceptance_mode=acceptance_mode,
                expected_reviewed_output=expected_reviewed_output,
                expected_reviewed_sha256=expected_reviewed_sha256,
            )
        )
    return issues


def surface_verdict_passes(value: str, *, allow_not_applicable_with_reason: bool = False) -> bool:
    lowered = normalize(value).lower()
    if re.search(r"\bpass(?:ed)?\b", lowered) and not contains_any(
        lowered,
        {"fail", "failed", "missing", "not checked", "pending", "blocked", "unresolved", "drift"},
    ):
        return True
    if allow_not_applicable_with_reason and (
        "not-applicable-with-reason" in lowered or "not-present" in lowered
    ):
        return "reason" in lowered and not contains_any(lowered, {"blank", "missing reason"})
    return False


def user_reported_field_gate_required(context: str, tokens: tuple[str, ...]) -> bool:
    lowered = context.lower()
    return contains_any(lowered, set(tokens))


def require_gate_path_field(
    issues: list[str],
    *,
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    record_path: Path,
    prefix: str,
    label: str,
) -> None:
    value = gate_values.get(prefix, "")
    if not is_explicit(value) or is_explicit_none(value):
        issues.append(f"gate record lacks {label}: {prefix}")
        return
    raw_value = gate_raw_values.get(prefix, value)
    for raw_path in split_path_values(raw_value):
        resolved = resolve_record_path(raw_path, record_path)
        issues.extend(validate_existing_path(resolved, require_nonempty_file=True))


def require_gate_pass_verdict(
    issues: list[str],
    *,
    gate_values: dict[str, str],
    prefix: str,
    label: str,
) -> None:
    value = gate_values.get(prefix, "")
    if not surface_verdict_passes(value) or contains_any(
        value,
        {"not-applicable", "none", "missing", "not checked", "stale", "pending", "blocked"},
    ):
        issues.append(f"gate record {label} must be pass: {prefix}")


BLOCKING_CONSISTENCY_FIELDS = (
    ("- protected-surface evidence contract verdict:", False),
    ("- surface paragraph-and-typography verdict:", False),
    ("- TOC visual geometry verdict:", False),
    ("- TOC paragraph-and-typography verdict:", False),
    ("- whole-document pagination verdict:", False),
    ("- high-risk thesis format surface verdict:", False),
    ("- header presence verdict:", True),
    ("- header rendered verdict:", True),
    ("- figure manifest contract verdict:", True),
)


def blocking_verdict_reasons(values: dict[str, str]) -> list[str]:
    reasons: list[str] = []
    for prefix, allow_not_applicable in BLOCKING_CONSISTENCY_FIELDS:
        value = values.get(prefix, "")
        if not value:
            continue
        lowered = normalize(value).lower()
        if "program-only" in lowered and "no thesis" in lowered:
            continue
        if not surface_verdict_passes(value, allow_not_applicable_with_reason=allow_not_applicable):
            reasons.append(f"{prefix} {value}")
    return reasons


def validate_blocking_state_consistency(values: dict[str, str], *, location: Path | str) -> list[str]:
    issues: list[str] = []
    reasons = blocking_verdict_reasons(values)
    if not reasons:
        return issues
    location_text = str(location)
    if values.get("- handoff status:", "").strip().lower() == "pass":
        issues.append(f"{location_text} has blocking evidence but handoff status is pass: {'; '.join(reasons[:4])}")
    if values.get("- audit verdict:", "").strip().lower() == "pass":
        issues.append(f"{location_text} has blocking evidence but audit verdict is pass: {'; '.join(reasons[:4])}")
    if is_explicit_none(values.get("- blockers:", "")):
        issues.append(f"{location_text} has blocking evidence but blockers is none: {'; '.join(reasons[:4])}")
    if is_explicit_none(values.get("- known caveats:", "")):
        issues.append(f"{location_text} has blocking evidence but known caveats is none: {'; '.join(reasons[:4])}")
    if values.get("- validation result:", "").strip().lower() == "pass":
        issues.append(f"{location_text} has blocking evidence but validation result is pass: {'; '.join(reasons[:4])}")
    return issues


NON_PROMOTION_TOKENS = {
    "candidate-only",
    "candidate only",
    "audit-only",
    "audit only",
    "blocked",
    "not promoted",
    "not-promoted",
    "unverified",
    "unresolved",
    "stale",
    "pending",
    "not baseline",
    "not a baseline",
}

SCOPED_REPAIR_CONTEXT_TOKENS = {
    "local-surface",
    "local surface",
    "content-only",
    "content only",
    "scoped",
    "specialty",
    "touched-surfaces-only",
    "touched surfaces only",
    "single-surface",
    "single surface",
    "candidate artifact",
    "candidate-only",
    "audit-only",
}


def parse_release_blocker_count(value: str) -> int | None:
    stripped = normalize(value).strip()
    return int(stripped) if re.fullmatch(r"\d+", stripped) else None


def is_scoped_thesis_repair_context(
    *,
    selected_workflow: str,
    subtask: str,
    verification_scope: str,
    rebuild_class: str,
    touched_surfaces: str,
) -> bool:
    if selected_workflow in {"local-surface-repair", "content-only-paragraph-revision"}:
        return True
    context = " ".join([subtask, verification_scope, rebuild_class, touched_surfaces]).lower()
    return contains_any(context, SCOPED_REPAIR_CONTEXT_TOKENS)


def validate_baseline_promotion_state(
    values: dict[str, str],
    *,
    task_mode: str,
    selected_workflow: str,
    record_path: Path,
) -> list[str]:
    issues: list[str] = []
    if task_mode not in THESIS_MODES:
        return issues

    handoff_pass = values.get("- handoff status:", "").strip().lower() == "pass"
    known_caveats = values.get("- known caveats:", "")
    baseline_status = values.get("- baseline promotion status:", "")
    baseline_evidence = values.get("- baseline promotion evidence path:", "")
    blocker_ledger = values.get("- release blocker ledger path:", "")
    blocker_count_text = values.get("- unresolved release blocker count:", "")
    next_baseline_verdict = values.get("- scoped artifact next-baseline verdict:", "")
    subtask = values.get("- subtask:", "")
    verification_scope = values.get("- verification scope claim:", "")
    rebuild_class = values.get("- rebuild class:", "")
    touched_surfaces = values.get("- touched template-owned surface families:", "")
    scoped_context = is_scoped_thesis_repair_context(
        selected_workflow=selected_workflow,
        subtask=subtask,
        verification_scope=verification_scope,
        rebuild_class=rebuild_class,
        touched_surfaces=touched_surfaces,
    )
    blocker_count = parse_release_blocker_count(blocker_count_text)
    if blocker_count is None:
        issues.append("gate record unresolved release blocker count must be a nonnegative integer")
    if handoff_pass and blocker_count not in (0, None):
        if scoped_context:
            issues.append(
                "local-surface repair cannot be handed off as pass while unresolved release blockers remain"
            )
        else:
            issues.append("thesis repair cannot be handed off as pass while unresolved release blockers remain")
    if handoff_pass and not surface_verdict_passes(baseline_status):
        issues.append("gate record baseline promotion status must be pass before a thesis handoff can pass")
    if handoff_pass and (
        not is_explicit(baseline_evidence) or is_explicit_none(baseline_evidence)
    ):
        issues.append("gate record baseline promotion evidence path must be present before a thesis handoff can pass")
    if handoff_pass and (
        not is_explicit(blocker_ledger) or is_explicit_none(blocker_ledger)
    ):
        issues.append("gate record release blocker ledger path must be present before a thesis handoff can pass")
    if handoff_pass and not surface_verdict_passes(next_baseline_verdict):
        issues.append("gate record scoped artifact next-baseline verdict must be pass before a thesis handoff can pass")
    if handoff_pass and contains_any(next_baseline_verdict, NON_PROMOTION_TOKENS):
        issues.append(
            "gate record scoped artifact next-baseline verdict blocks release-baseline promotion"
        )
    if handoff_pass and not is_explicit_none(known_caveats):
        issues.append("gate record known caveats must be none before baseline promotion")

    helper_risk_count = parse_release_blocker_count(values.get("- project-local helper risk count:", ""))
    helper_disposition = values.get("- project-local helper disposition:", "")
    restart_status = values.get("- canonical source restart required?:", "")
    helper_summary = values.get("- project-local helper script preflight summary:", "")
    completed_restart = (
        helper_disposition == "clean-source-restart-completed"
        or restart_status in {"completed", "clean-source-restart-completed"}
    )
    helper_risky = (
        helper_risk_count is not None
        and helper_risk_count > 0
        or contains_any(helper_summary, {"failed", "contaminated", "thick", "risky project-local"})
    )
    if handoff_pass and helper_risky and not completed_restart and surface_verdict_passes(baseline_status):
        issues.append(
            "gate record baseline promotion must be blocked until clean-source restart completes when risky project-local helpers are detected"
        )
    return issues


def first_resolved_path(raw_value: str, record_path: Path) -> Path | None:
    for raw_path in split_path_values(raw_value):
        if not is_explicit_none(normalize(raw_path)):
            return resolve_record_path(raw_path, record_path)
    return None


def review_inventory_source_docx_path(source_report_path: Path | None) -> Path | None:
    if source_report_path is None or not source_report_path.exists():
        return None
    try:
        text = source_report_path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return None
    for line in text.splitlines():
        if not line.lower().startswith("- docx path:"):
            continue
        value = line.partition(":")[2].strip()
        if not value:
            return None
        path = Path(value)
        if not path.is_absolute():
            path = (source_report_path.parent / path).resolve()
        else:
            path = path.resolve()
        return path
    return None


def citation_inventory_source_docx_path(source_report_path: Path | None) -> Path | None:
    return review_inventory_source_docx_path(source_report_path)


def docx_has_figure_caption(docx_path: Path | None) -> bool:
    if docx_path is None or not docx_path.exists() or docx_path.suffix.lower() != ".docx":
        return False
    return bool(docx_figure_surface_summary(docx_path).get("figure_caption_count"))


def docx_has_body_table_or_caption(docx_path: Path | None) -> bool:
    if docx_path is None or not docx_path.exists() or docx_path.suffix.lower() != ".docx":
        return False
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return False
    if root.findall(f".//{W}tbl"):
        return True
    texts = [node.text or "" for node in root.findall(f".//{W}t")]
    joined = "\n".join(texts)
    return bool(re.search(r"(^|\n)\s*(表|Table)\s*\d+[\-.．、]?\d*", joined, flags=re.IGNORECASE))


def sample_self_check_detector_registry(text: str) -> dict[str, dict[str, object]]:
    detectors: dict[str, dict[str, object]] = {}
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("{") or '"id"' not in stripped:
            continue
        try:
            item = json.loads(stripped)
        except json.JSONDecodeError:
            continue
        detector_id = str(item.get("id") or "")
        if detector_id:
            detectors[detector_id] = item
    return detectors


def validate_sample_self_check_report(
    report_path: Path,
    record_path: Path,
    *,
    expected_final_docx: Path | None = None,
    expected_final_sha256: str = "",
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(report_path, require_nonempty_file=True))
    if issues:
        return issues
    raw_text = read_optional_text(report_path)
    text = raw_text.lower()
    if expected_final_docx is not None:
        binding_path = ""
        binding_sha = ""
        for line in raw_text.splitlines():
            stripped = line.strip()
            lowered = stripped.lower()
            if lowered.startswith("- final docx path:") or lowered.startswith("- document path:"):
                binding_path = stripped.split(":", 1)[1].strip()
            elif lowered.startswith("- final docx sha256:"):
                binding_sha = stripped.split(":", 1)[1].strip()
        if not binding_path:
            issues.append(f"sample self-check report missing final DOCX path binding: {report_path}")
        else:
            try:
                if Path(binding_path).resolve() != expected_final_docx.resolve():
                    issues.append(
                        "sample self-check report final DOCX path does not match the gate record output: "
                        f"{report_path}"
                    )
            except OSError:
                issues.append(f"sample self-check report final DOCX path is not resolvable: {report_path}")
        expected_sha = expected_final_sha256.strip()
        if not expected_sha and expected_final_docx.exists():
            expected_sha = sha256_file(expected_final_docx)
        if not binding_sha:
            issues.append(f"sample self-check report missing final DOCX sha256 binding: {report_path}")
        elif expected_sha and binding_sha.lower() != expected_sha.lower():
            issues.append(
                "sample self-check report final DOCX sha256 does not match the gate record output: "
                f"{report_path}"
            )
    for token in SAMPLE_SELF_CHECK_BLOCK_TOKENS:
        if token in text:
            issues.append(f"sample self-check report is delivery-blocked in {record_path}: {token}")
    detectors = sample_self_check_detector_registry(raw_text)
    blocking_failed_detectors = [
        detector_id
        for detector_id, detector in detectors.items()
        if detector.get("blocking") is True
        and (detector.get("failed") is True or detector.get("passed") is False)
    ]
    if blocking_failed_detectors:
        issues.append(
            "sample self-check report contains blocking failed detector results: "
            f"{report_path}: {', '.join(sorted(blocking_failed_detectors))}"
        )
    for detector_id in REQUIRED_SAMPLE_SELF_CHECK_DETECTORS:
        detector = detectors.get(detector_id)
        if detector is None:
            issues.append(f"sample self-check report missing required detector `{detector_id}`: {report_path}")
            continue
        if detector.get("passed") is not True or detector.get("failed") is True:
            issues.append(f"sample self-check required detector `{detector_id}` is not pass-shaped: {report_path}")
        evidence = detector.get("evidence")
        if not isinstance(evidence, dict) or not evidence:
            issues.append(f"sample self-check detector `{detector_id}` has no evidence object: {report_path}")
    return issues


def validate_page_class_coverage_matrix(matrix_path: Path, record_path: Path, acceptance_mode: str) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(matrix_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(matrix_path)
    lowered = text.lower()
    if re.search(r"\b(sampled-only|sample only|not checked|missing|pending|blocked|unresolved)\b", lowered) or re.search(
        r"\b(fail|failed)\b(?!\s*[:=]\s*false)", lowered
    ):
        issues.append(f"page-class coverage matrix records unresolved or sampled-only coverage: {matrix_path}")
    payload: object | None = None
    if matrix_path.suffix.lower() == ".json":
        try:
            payload = json.loads(text)
        except Exception as exc:
            issues.append(f"page-class coverage matrix is not valid JSON: {matrix_path} ({exc})")
    searchable = json.dumps(payload, ensure_ascii=False, sort_keys=True).lower() if payload is not None else lowered
    missing = [
        label for label, aliases in PAGE_CLASS_ALIASES.items()
        if not any(alias in searchable for alias in aliases)
    ]
    if missing:
        issues.append(
            f"page-class coverage matrix missing required page classes in {record_path}: {', '.join(missing)}"
        )
    if missing:
        return issues
    if "pass" not in searchable and "passed" not in searchable:
        issues.append(f"page-class coverage matrix lacks pass verdicts: {matrix_path}")
    if not value_is_pathlike(searchable):
        issues.append(f"page-class coverage matrix lacks evidence paths: {matrix_path}")
    if payload is not None:
        issues.extend(validate_json_page_class_entries(payload, matrix_path, acceptance_mode))
    else:
        issues.extend(validate_text_page_class_entries(text, matrix_path, acceptance_mode))
    return issues


def validate_user_issue_ledger(
    ledger_path: Path,
    record_path: Path,
    *,
    abstract_required: bool,
    references_required: bool,
    required_surface_labels: tuple[str, ...] = (),
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(ledger_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(ledger_path)
    lowered = text.lower()
    required_tokens = ("issue id", "user wording", "surface", "expected fix", "evidence path", "final verdict")
    missing_tokens = [token for token in required_tokens if token not in lowered]
    if missing_tokens:
        issues.append(
            f"user-reported issue ledger is missing required fields in {record_path}: {', '.join(missing_tokens)}"
        )
    verdict_values = [
        line.split(":", 1)[1].strip()
        for line in text.splitlines()
        if line.lower().strip().startswith("- final verdict:")
    ]
    if not verdict_values:
        issues.append(f"user-reported issue ledger has no final verdict rows: {ledger_path}")
    for value in verdict_values:
        if normalize(value).lower() != "pass":
            issues.append(
                f"user-reported issue ledger final verdict must be exactly pass, found `{value or 'blank'}`: {ledger_path}"
            )
    if abstract_required:
        missing_abstract = [token for token in ABSTRACT_LEDGER_SURFACES if token not in lowered]
        if missing_abstract:
            issues.append(
                "user-reported issue ledger missing abstract surface coverage in "
                f"{record_path}: {', '.join(missing_abstract)}"
            )
    if references_required and not contains_any(lowered, {"reference", "references", "bibliography", "citation"}):
        issues.append(f"user-reported issue ledger missing references/bibliography coverage in {record_path}")
    for surface in required_surface_labels:
        _trigger_tokens, ledger_tokens = SURFACE_LEDGER_REQUIREMENTS.get(surface, ((), ()))
        if ledger_tokens and not contains_any(lowered, set(ledger_tokens)):
            issues.append(f"user-reported issue ledger missing {surface} coverage in {record_path}")
    evidence_values = [
        line.split(":", 1)[1].strip()
        for line in text.splitlines()
        if line.lower().strip().startswith("- evidence path:")
    ]
    if not evidence_values:
        issues.append(f"user-reported issue ledger has no evidence path rows: {ledger_path}")
    for value in evidence_values:
        if not value or is_explicit_none(normalize(value)):
            issues.append(f"user-reported issue ledger evidence path is empty/none: {ledger_path}")
            continue
        for raw_path in split_path_values(value):
            resolved = resolve_record_path(raw_path, ledger_path)
            issues.extend(validate_existing_path(resolved, require_nonempty_file=True))
    return issues


def validate_body_opener_header_title_evidence(
    evidence_path: Path,
    record_path: Path,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(evidence_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(evidence_path)
    lowered = text.lower()
    missing_tokens = [
        token
        for token in BODY_OPENER_HEADER_TITLE_EVIDENCE_TOKENS
        if token not in lowered
    ]
    if missing_tokens:
        issues.append(
            "body opener/header title consistency evidence lacks required tokens in "
            f"{record_path}: {', '.join(missing_tokens)}"
        )
    if contains_any(
        lowered,
        {
            "failed",
            "missing",
            "not checked",
            "sample only",
            "sampled-only",
            "stale",
            "unresolved",
            "blocked",
            "mismatch remains",
            "needs review",
        },
    ):
        issues.append("body opener/header title consistency evidence records unresolved title/header drift")
    return issues


def user_reported_visual_defect_required(user_issue_context: str) -> bool:
    lowered = user_issue_context.lower()
    visual_context = {
        "screenshot",
        "visible",
        "visual",
        "rendered",
        "looks",
        "wrong",
        "format",
        "drift",
        "abnormal",
        "\u622a\u56fe",
        "\u53ef\u89c1",
        "\u89c6\u89c9",
        "\u683c\u5f0f",
        "\u5f02\u5e38",
        "\u9519",
    }
    toc_issue = contains_any(lowered, {"toc", "table of contents", "\u76ee\u5f55"}) and contains_any(
        lowered,
        visual_context,
    )
    references_pagination_issue = contains_any(
        lowered,
        {"references", "bibliography", "\u53c2\u8003\u6587\u732e"},
    ) and contains_any(
        lowered,
        {"pagination", "page break", "page-break", "fresh page", "continuation page", "\u5206\u9875", "\u6362\u9875", "\u65ad\u9875"},
    )
    body_font_issue = contains_any(lowered, {"body", "\u6b63\u6587"}) and contains_any(
        lowered,
        {"body font", "body size", "body style", "font", "font size", "typeface", "style abnormal", "\u5b57\u4f53", "\u5b57\u53f7", "\u6837\u5f0f"},
    ) and contains_any(
        lowered,
        visual_context,
    )
    abstract_issue = contains_any(lowered, {"abstract", "keyword", "\u6458\u8981", "\u5173\u952e\u8bcd"}) and contains_any(
        lowered,
        visual_context | {"font", "font size", "style", "\u5b57\u4f53", "\u5b57\u53f7", "\u6837\u5f0f", "\u7edf\u4e00"},
    )
    header_footer_page_number_issue = contains_any(
        lowered,
        {
            "header",
            "footer",
            "page number",
            "page-number",
            "page_numbers",
            "running header",
            "header line",
            "horizontal line",
            "footer position",
            "\u9875\u7709",
            "\u9875\u811a",
            "\u9875\u7801",
            "\u6a2a\u7ebf",
            "\u6807\u53f7",
        },
    ) and contains_any(lowered, visual_context | {"position", "line", "\u4f4d\u7f6e", "\u6a2a\u7ebf"})
    return toc_issue or references_pagination_issue or body_font_issue or abstract_issue or header_footer_page_number_issue


def validate_user_reported_visual_defect_evidence(
    evidence_path: Path,
    record_path: Path,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(evidence_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(evidence_path)
    lowered = text.lower()
    missing_tokens = [
        token
        for token in USER_REPORTED_VISUAL_EVIDENCE_TOKENS
        if token not in lowered
    ]
    if missing_tokens:
        issues.append(
            "user-reported visual defect evidence is missing required template/rendered/geometry fields "
            f"in {record_path}: {', '.join(missing_tokens)}"
        )
    if not contains_any(lowered, {"target", "actual", "final"}):
        issues.append(f"user-reported visual defect evidence must bind the target/actual/final artifact: {evidence_path}")
    if not contains_any(lowered, {"full-page", "full page", "page class", "physical page", "logical page"}):
        issues.append(f"user-reported visual defect evidence must include full-page binding: {evidence_path}")
    if not contains_any(lowered, {"key-surface", "key surface", "crop", "region", "bbox", "bounding box"}):
        issues.append(f"user-reported visual defect evidence must include key-surface/crop binding: {evidence_path}")
    bad_token_scan_lines = []
    path_like_prefixes = (
        "- final docx path:",
        "- final pdf path:",
        "- reviewed output path:",
        "- template path:",
        "- template rendered region image path:",
        "- actual rendered region image path:",
        "- target actual final artifact binding:",
        "- surface evidence paths:",
        "- final docx sha256:",
    )
    for line in lowered.splitlines():
        stripped = line.strip()
        if any(stripped.startswith(prefix) for prefix in path_like_prefixes):
            continue
        bad_token_scan_lines.append(line)
    if contains_any("\n".join(bad_token_scan_lines), USER_REPORTED_VISUAL_BAD_TOKENS):
        issues.append(f"user-reported visual defect evidence contains blocked or substitute proof wording: {evidence_path}")
    for line in lowered.splitlines():
        if contains_any(line, USER_REPORTED_VISUAL_NOT_APPLICABLE_CONTEXT_TOKENS) and contains_any(
            line,
            USER_REPORTED_VISUAL_NOT_APPLICABLE_TOKENS,
        ):
            issues.append(f"user-reported visual defect evidence contains blocked or substitute proof wording: {evidence_path}")
            break
    return issues


def thesis_content_mutation_visual_gate_required(context: str) -> bool:
    lowered = context.lower()
    if not contains_any(lowered, CONTENT_MUTATION_VISUAL_TRIGGER_TOKENS):
        return False
    if contains_any(lowered, {"audit-only", "no mutation", "read-only", "not-applicable"}):
        return contains_any(lowered, {"content expansion", "body paragraph insertion", "\u6269\u5199", "\u5b57\u6570"})
    return True


def validate_content_mutation_rendered_evidence(
    evidence_path: Path,
    record_path: Path,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(evidence_path, require_nonempty_file=True))
    if issues:
        return issues
    text = read_optional_text(evidence_path)
    lowered = text.lower()
    missing_tokens = [
        token
        for token in CONTENT_MUTATION_VISUAL_EVIDENCE_TOKENS
        if token not in lowered
    ]
    if missing_tokens:
        issues.append(
            "content expansion machine-vision evidence is missing required rendered body-contamination fields "
            f"in {record_path}: {', '.join(missing_tokens)}"
        )
    if contains_any(lowered, CONTENT_MUTATION_VISUAL_BAD_TOKENS):
        issues.append(f"content expansion machine-vision evidence contains blocked substitute proof wording: {evidence_path}")
    return issues


def active_format_task_user_issue_context(format_task_context: str) -> str:
    issue_prefixes = (
        "- touched blocks this round:",
        "- rendered pages to inspect:",
        "- rendered sentinel texts to confirm:",
        "- specific format classes to verify:",
        "- style-blast-radius trigger:",
    )
    values: list[str] = []
    for line in format_task_context.splitlines():
        stripped = line.lower().strip()
        if not stripped.startswith(issue_prefixes):
            continue
        value = raw_line_value(line)
        if is_explicit(value) and not is_explicit_none(normalize(value)):
            values.append(value)
    return "\n".join(values)


def validate_figure_contract_manifest(
    manifest_path: Path,
    final_docx: Path | None,
    record_path: Path,
    *,
    source_docx: Path | None = None,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(manifest_path, require_nonempty_file=True))
    if issues:
        return issues
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"figure asset manifest is not valid JSON in {record_path}: {manifest_path} ({exc})"]
    issues.extend(
        validate_figure_manifest(
            manifest,
            final_docx=final_docx,
            source_docx=source_docx,
            manifest_path=manifest_path,
        )
    )
    return issues


def _path_equivalent(left: str | Path | None, right: str | Path | None) -> bool:
    if left is None or right is None:
        return False
    try:
        return str(Path(left).resolve()).lower() == str(Path(right).resolve()).lower()
    except Exception:
        return str(left).strip().lower() == str(right).strip().lower()


def _load_json_evidence(path: Path, label: str) -> tuple[dict[str, object] | None, list[str]]:
    issues: list[str] = []
    issues.extend(validate_existing_path(path, require_nonempty_file=True))
    if issues:
        return None, issues
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return None, [f"{label} evidence must be readable JSON: {path} ({exc})"]
    if not isinstance(payload, dict):
        return None, [f"{label} evidence JSON must be an object: {path}"]
    return payload, []


def _json_passes(payload: dict[str, object]) -> bool:
    passed = payload.get("passed")
    if passed is True:
        return True
    for key in ("result", "verdict", "status"):
        value = payload.get(key)
        if isinstance(value, str) and surface_verdict_passes(value):
            return True
    return False


def _validate_bound_final_docx(
    payload: dict[str, object],
    *,
    evidence_path: Path,
    label: str,
    final_docx: Path | None,
) -> list[str]:
    issues: list[str] = []
    final_path_value = str(
        payload.get("final_docx_path")
        or payload.get("reviewed_output_path")
        or payload.get("docx_path")
        or ""
    ).strip()
    final_sha_value = str(
        payload.get("final_docx_sha256")
        or payload.get("reviewed_output_sha256")
        or payload.get("docx_sha256")
        or ""
    ).strip()
    if not final_path_value:
        issues.append(f"{label} evidence lacks final_docx_path: {evidence_path}")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", final_sha_value or ""):
        issues.append(f"{label} evidence lacks 64-hex final_docx_sha256: {evidence_path}")
    if final_docx is not None and final_docx.exists():
        if final_path_value and not _path_equivalent(final_path_value, final_docx):
            issues.append(f"{label} evidence final_docx_path does not match exact final DOCX: {evidence_path}")
        expected_sha = sha256_file(final_docx)
        if final_sha_value and final_sha_value.upper() != expected_sha.upper():
            issues.append(f"{label} evidence final_docx_sha256 does not match exact final DOCX: {evidence_path}")
    return issues


def validate_figure_extents_user_evidence(
    evidence_path: Path,
    record_path: Path,
    *,
    final_docx: Path | None = None,
) -> list[str]:
    payload, issues = _load_json_evidence(evidence_path, "figure extents")
    if payload is None:
        return issues
    if payload.get("schema") not in {
        "graduation-project-builder.figure-extents-audit.v2",
        "graduation-project-builder.figure-extents-audit.v3",
    }:
        issues.append(f"figure extents evidence has wrong schema in {record_path}: {evidence_path}")
    if not _json_passes(payload):
        issues.append(f"figure extents evidence must pass in {record_path}: {evidence_path}")
    issues.extend(
        _validate_bound_final_docx(payload, evidence_path=evidence_path, label="figure extents", final_docx=final_docx)
    )
    try:
        text_width_cm = float(payload.get("text_width_cm") or 0)
    except (TypeError, ValueError):
        text_width_cm = 0.0
    if text_width_cm <= 0:
        issues.append(f"figure extents evidence must record text_width_cm > 0: {evidence_path}")
    try:
        body_figure_count = int(payload.get("body_figure_count") or 0)
    except (TypeError, ValueError):
        body_figure_count = 0
    if body_figure_count <= 0:
        issues.append(f"figure extents evidence must include body figures for width repair: {evidence_path}")
    try:
        drift_count = int(payload.get("paragraph_margin_width_drift_count") or 0)
    except (TypeError, ValueError):
        drift_count = -1
    if drift_count != 0:
        issues.append(f"figure extents evidence paragraph_margin_width_drift_count must be 0: {evidence_path}")
    for key, label in (
        ("nonzero_crop_count", "nonzero crop count"),
        ("exact_line_spacing_clip_count", "exact-line-spacing clipping count"),
        ("image_holder_layout_issue_count", "image-holder layout issue count"),
        ("image_holder_abnormal_indent_count", "image-holder abnormal indent count"),
        ("image_holder_noncenter_alignment_count", "image-holder noncenter alignment count"),
        ("safe_page_height_risk_count", "safe page-height risk count"),
    ):
        if key not in payload:
            issues.append(f"figure extents evidence missing {label}: {evidence_path}")
            continue
        try:
            count = int(payload.get(key) or 0)
        except (TypeError, ValueError):
            count = -1
        if count != 0:
            issues.append(f"figure extents evidence {label} must be 0: {evidence_path}")
    visible_verdict = str(payload.get("visible_content_completeness_verdict") or "pass").strip().lower()
    if visible_verdict not in {"pass", "passed"}:
        issues.append(f"figure extents evidence visible_content_completeness_verdict must pass: {evidence_path}")
    layout_verdict = str(payload.get("image_holder_layout_verdict") or "").strip().lower()
    if layout_verdict not in {"pass", "passed"}:
        issues.append(f"figure extents evidence image_holder_layout_verdict must pass: {evidence_path}")
    figures = payload.get("figures")
    if not isinstance(figures, list) or not figures:
        issues.append(f"figure extents evidence must include per-figure rows: {evidence_path}")
    else:
        for index, row in enumerate(figures, start=1):
            if not isinstance(row, dict):
                issues.append(f"figure extents row {index} is not an object: {evidence_path}")
                continue
            if "width_text_ratio" not in row:
                issues.append(f"figure extents row {index} lacks width_text_ratio: {evidence_path}")
            if row.get("paragraph_margin_width_drift") is True:
                issues.append(f"figure extents row {index} still has paragraph_margin_width_drift=true: {evidence_path}")
            if row.get("nonzero_crop") is True:
                issues.append(f"figure extents row {index} still has nonzero_crop=true: {evidence_path}")
            if row.get("exact_line_spacing_clips_inline_image") is True:
                issues.append(f"figure extents row {index} still has exact_line_spacing_clips_inline_image=true: {evidence_path}")
            if "holder_effective_alignment" not in row:
                issues.append(f"figure extents row {index} lacks holder_effective_alignment: {evidence_path}")
            if "holder_effective_indent" not in row:
                issues.append(f"figure extents row {index} lacks holder_effective_indent: {evidence_path}")
            if row.get("image_holder_layout_issue") is True:
                issues.append(f"figure extents row {index} still has image_holder_layout_issue=true: {evidence_path}")
    if payload.get("issues") not in ([], None):
        issues.append(f"figure extents evidence carries unresolved issues: {evidence_path}")
    return issues


def validate_algorithm_code_visibility_evidence(
    evidence_path: Path,
    record_path: Path,
    *,
    final_docx: Path | None = None,
) -> list[str]:
    payload, issues = _load_json_evidence(evidence_path, "algorithm code visibility")
    if payload is None:
        return issues
    if payload.get("schema") != "graduation-project-builder.algorithm-code-visibility.v1":
        issues.append(f"algorithm code visibility evidence has wrong schema in {record_path}: {evidence_path}")
    if not _json_passes(payload):
        issues.append(f"algorithm code visibility evidence must pass in {record_path}: {evidence_path}")
    issues.extend(
        _validate_bound_final_docx(payload, evidence_path=evidence_path, label="algorithm code visibility", final_docx=final_docx)
    )
    code_file_value = str(payload.get("code_file_path") or payload.get("source_code_path") or "").strip()
    if not code_file_value:
        issues.append(f"algorithm code visibility evidence lacks code_file_path: {evidence_path}")
    else:
        code_file = Path(code_file_value)
        issues.extend(validate_existing_path(code_file, require_nonempty_file=True))
        recorded_sha = str(payload.get("code_file_sha256") or payload.get("source_code_sha256") or "").strip()
        if not re.fullmatch(r"[0-9a-fA-F]{64}", recorded_sha or ""):
            issues.append(f"algorithm code visibility evidence lacks 64-hex code_file_sha256: {evidence_path}")
        elif code_file.exists() and recorded_sha.lower() != sha256_file(code_file):
            issues.append(f"algorithm code visibility evidence code_file_sha256 mismatch: {evidence_path}")
    line_start = payload.get("visible_line_start") or payload.get("line_start")
    line_end = payload.get("visible_line_end") or payload.get("line_end")
    try:
        start_int = int(line_start)
        end_int = int(line_end)
    except (TypeError, ValueError):
        issues.append(f"algorithm code visibility evidence must record numeric visible line range: {evidence_path}")
    else:
        if start_int <= 0 or end_int < start_int:
            issues.append(f"algorithm code visibility evidence visible line range is invalid: {evidence_path}")
    try:
        visible_count = int(payload.get("visible_code_line_count") or 0)
    except (TypeError, ValueError):
        visible_count = 0
    if visible_count <= 0:
        issues.append(f"algorithm code visibility evidence must record visible_code_line_count > 0: {evidence_path}")
    for key, label in (
        ("docx_binding_verdict", "DOCX binding"),
        ("code_visibility_verdict", "code visibility"),
    ):
        if not surface_verdict_passes(str(payload.get(key) or "")):
            issues.append(f"algorithm code visibility evidence {label} verdict must pass: {evidence_path}")
    return issues


def validate_toc_digits_cache_evidence(
    evidence_path: Path,
    record_path: Path,
    *,
    final_docx: Path | None = None,
) -> list[str]:
    payload, issues = _load_json_evidence(evidence_path, "TOC digits/cache preservation")
    if payload is None:
        return issues
    if payload.get("schema") != "graduation-project-builder.toc-digits-cache-preservation.v1":
        issues.append(f"TOC digits/cache preservation evidence has wrong schema in {record_path}: {evidence_path}")
    if not _json_passes(payload):
        issues.append(f"TOC digits/cache preservation evidence must pass in {record_path}: {evidence_path}")
    issues.extend(
        _validate_bound_final_docx(payload, evidence_path=evidence_path, label="TOC digits/cache preservation", final_docx=final_docx)
    )
    try:
        row_count = int(payload.get("toc_cache_row_count") or 0)
    except (TypeError, ValueError):
        row_count = 0
    if row_count <= 0:
        issues.append(f"TOC digits/cache evidence must record toc_cache_row_count > 0: {evidence_path}")
    for key, label in (
        ("digit_spacing_verdict", "digit spacing"),
        ("page_number_column_verdict", "page-number column"),
        ("bookmark_anchor_verdict", "bookmark anchor"),
        ("hyperlink_anchor_verdict", "hyperlink anchor"),
        ("before_after_cache_verdict", "before/after cache"),
    ):
        if not surface_verdict_passes(str(payload.get(key) or "")):
            issues.append(f"TOC digits/cache evidence {label} verdict must pass: {evidence_path}")
    for key, label in (
        ("broken_digit_row_count", "broken digit row"),
        ("missing_page_number_count", "missing page-number"),
        ("missing_bookmark_anchor_count", "missing bookmark anchor"),
        ("missing_hyperlink_anchor_count", "missing hyperlink anchor"),
    ):
        try:
            count = int(payload.get(key) or 0)
        except (TypeError, ValueError):
            issues.append(f"TOC digits/cache evidence {label} count must be numeric: {evidence_path}")
            continue
        if count != 0:
            issues.append(f"TOC digits/cache evidence {label} count must be 0: {evidence_path}")
    return issues


def validate_flowchart_acceptance_fields(
    *,
    issues: list[str],
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    record_path: Path,
    final_docx: Path | None,
) -> None:
    required_path_fields = (
        ("- flowchart semantic trigger inventory path:", "flowchart semantic trigger inventory path"),
        ("- flowchart draw.io source evidence paths:", "flowchart draw.io source evidence paths"),
        ("- flowchart SVG export evidence paths:", "flowchart SVG export evidence paths"),
        ("- flowchart raster fallback evidence paths:", "flowchart raster fallback evidence paths"),
        ("- flowchart source-scale geometry report paths:", "flowchart source-scale geometry report paths"),
        ("- flowchart rendered-page evidence paths:", "flowchart rendered-page evidence paths"),
        ("- flowchart final DOCX relationship evidence paths:", "flowchart final DOCX relationship evidence paths"),
    )
    for prefix, label in required_path_fields:
        require_gate_path_field(
            issues,
            gate_values=gate_values,
            gate_raw_values=gate_raw_values,
            record_path=record_path,
            prefix=prefix,
            label=label,
        )
    require_gate_pass_verdict(
        issues,
        gate_values=gate_values,
        prefix="- flowchart manifest contract verdict:",
        label="flowchart manifest contract verdict",
    )

    drawio_paths = [
        resolve_record_path(raw_path, record_path)
        for raw_path in split_path_values(gate_raw_values.get("- flowchart draw.io source evidence paths:", ""))
    ]
    geometry_paths = [
        resolve_record_path(raw_path, record_path)
        for raw_path in split_path_values(gate_raw_values.get("- flowchart source-scale geometry report paths:", ""))
    ]
    for drawio_path in drawio_paths:
        if drawio_path.suffix.lower() != ".drawio":
            issues.append(f"flowchart draw.io source evidence must use .drawio files: {drawio_path}")
            continue
        issues.extend(validate_existing_path(drawio_path, require_nonempty_file=True))
        if drawio_path.exists():
            for issue in drawio_structural_geometry_issues(drawio_path, family="flowchart"):
                issues.append(f"flowchart draw.io source invalid: {issue}")
    for geometry_path in geometry_paths:
        issues.extend(validate_existing_path(geometry_path, require_nonempty_file=True))
        if not geometry_path.exists() or not drawio_paths:
            continue
        matching_drawio = drawio_paths[0]
        issues.extend(
            f"flowchart source-scale geometry report invalid: {issue}"
            for issue in validate_structural_geometry_report(geometry_path, matching_drawio)
        )
    for prefix, suffixes, label in (
        ("- flowchart SVG export evidence paths:", {".svg"}, "flowchart SVG export"),
        ("- flowchart raster fallback evidence paths:", IMAGE_EXTENSIONS, "flowchart raster fallback"),
        ("- flowchart rendered-page evidence paths:", IMAGE_EXTENSIONS | PDF_EXTENSIONS | {".md", ".json"}, "flowchart rendered-page evidence"),
        ("- flowchart final DOCX relationship evidence paths:", {".json", ".md", ".txt"}, "flowchart final DOCX relationship evidence"),
        ("- flowchart semantic trigger inventory path:", {".json", ".md", ".txt", ".csv"}, "flowchart semantic trigger inventory"),
    ):
        for raw_path in split_path_values(gate_raw_values.get(prefix, "")):
            path = resolve_record_path(raw_path, record_path)
            issues.extend(validate_existing_path(path, require_nonempty_file=True))
            if path.exists() and suffixes and path.suffix.lower() not in suffixes:
                issues.append(f"{label} has unsupported file type: {path}")
    if final_docx is not None and final_docx.exists():
        for raw_path in split_path_values(gate_raw_values.get("- flowchart final DOCX relationship evidence paths:", "")):
            path = resolve_record_path(raw_path, record_path)
            if path.suffix.lower() != ".json" or not path.exists():
                continue
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except Exception as exc:
                issues.append(f"flowchart final DOCX relationship evidence JSON is unreadable: {path} ({exc})")
                continue
            if isinstance(payload, dict):
                issues.extend(
                    _validate_bound_final_docx(
                        payload,
                        evidence_path=path,
                        label="flowchart final DOCX relationship",
                        final_docx=final_docx,
                    )
                )


def figure_manifest_has_er_diagrams(manifest_path: Path | None) -> bool:
    if manifest_path is None or not manifest_path.exists():
        return False
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return False
    diagrams = manifest.get("diagrams") if isinstance(manifest, dict) else None
    if not isinstance(diagrams, dict):
        return False
    for entry in diagrams.values():
        if not isinstance(entry, dict):
            continue
        family = str(entry.get("family") or entry.get("inferred_family") or "").strip().lower()
        if family == "er":
            return True
    return False


def _int_from_gate_value(value: str) -> int | None:
    match = re.search(r"-?\d+", value or "")
    if not match:
        return None
    try:
        return int(match.group(0))
    except ValueError:
        return None


def _mechanical_cad_context_required(gate_values: dict[str, str], record_text_lower: str) -> bool:
    cad_field_prefixes = (
        "- exact final CAD delivery package path:",
        "- mechanical drawing package audit path:",
        "- mechanical drawing package exact package path:",
        "- exact combined drawing PDF path:",
    )
    if any(
        is_explicit(gate_values.get(prefix, ""))
        and not is_explicit_none(gate_values.get(prefix, ""))
        for prefix in cad_field_prefixes
    ):
        return True
    return False


def _mechanical_formula_minimum_required(record_context: str) -> bool:
    lowered = record_context.lower()
    if contains_any(lowered, MECHANICAL_FORMULA_EXCLUSION_TOKENS):
        return False
    return contains_any(lowered, MECHANICAL_FORMULA_CONTEXT_TOKENS)


def _validate_cad_appendix_binding_record(
    *,
    record_path: Path,
    gate_values: dict[str, str],
    gate_raw_values: dict[str, str],
    final_docx_path: Path | None,
    record_text_lower: str,
) -> list[str]:
    issues: list[str] = []
    if not _mechanical_cad_context_required(gate_values, record_text_lower):
        return issues
    if final_docx_path is None or not final_docx_path.exists():
        issues.append("CAD appendix binding requires an existing exact final DOCX path")
        return issues

    for prefix in (
        "- mechanical drawing package audit path:",
        "- mechanical drawing package audit verdict:",
        "- mechanical drawing formal CAD source provenance verdict:",
        "- mechanical drawing schematic/concept substitute rejection verdict:",
    ):
        value = gate_values.get(prefix, "")
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"thesis final acceptance with mechanical CAD drawings must bind {prefix[2:-1]}")
    if not _mechanical_cad_passish(gate_values.get("- mechanical drawing package audit verdict:", "")):
        issues.append("mechanical drawing package audit verdict must be pass-shaped")
    if not _mechanical_cad_passish(gate_values.get("- mechanical drawing formal CAD source provenance verdict:", "")):
        issues.append("mechanical drawing formal CAD source provenance verdict must be pass-shaped")
    if not _mechanical_cad_passish(
        gate_values.get("- mechanical drawing schematic/concept substitute rejection verdict:", "")
    ):
        issues.append("mechanical drawing schematic/concept substitute rejection verdict must be pass-shaped")
    issues.extend(_mechanical_cad_require_final_surface_fields(gate_values))
    expected_cad_package_sha = (
        gate_values.get("- exact final CAD delivery package sha256:", "")
        or gate_values.get("- mechanical drawing package exact package sha256:", "")
    )
    issues.extend(
        _mechanical_cad_validate_render_readability_fields(
            record_path=record_path,
            values=gate_values,
            raw_values=gate_raw_values,
        )
    )
    issues.extend(
        _mechanical_cad_validate_linework_and_color_fields(
            record_path=record_path,
            values=gate_values,
            raw_values=gate_raw_values,
            expected_package_sha=expected_cad_package_sha,
        )
    )
    package_audit_value = gate_values.get("- mechanical drawing package audit path:", "")
    if is_explicit(package_audit_value) and not is_explicit_none(package_audit_value):
        package_audit_path = resolve_record_path(
            gate_raw_values.get("- mechanical drawing package audit path:", ""),
            record_path,
        )
        try:
            package_report = json.loads(package_audit_path.read_text(encoding="utf-8"))
        except Exception as exc:
            issues.append(f"mechanical drawing package audit path must be readable JSON: {package_audit_path} ({exc})")
        else:
            formal_source = package_report.get("formal_cad_source_verdict")
            if not isinstance(formal_source, dict) or formal_source.get("passed") is not True:
                issues.append("mechanical drawing package audit must pass formal CAD source provenance")
            elif formal_source.get("schematic_substitute_hits"):
                issues.append("mechanical drawing package audit must report zero schematic/concept substitute hits")
            issues.extend(
                _mechanical_cad_validate_final_package_audit_report(
                    package_report=package_report,
                    values=gate_values,
                    expected_package_sha=expected_cad_package_sha,
                )
            )

    for prefix in CAD_APPENDIX_FIELD_PREFIXES:
        value = gate_values.get(prefix, "")
        if not is_explicit(value) or is_explicit_none(value):
            issues.append(f"thesis final acceptance with mechanical CAD drawings must bind {prefix[2:-1]}")

    audit_path_value = gate_values.get("- CAD appendix binding audit path:", "")
    package_value = gate_values.get("- CAD appendix final package path:", "")
    if not is_explicit(audit_path_value) or is_explicit_none(audit_path_value):
        return issues
    audit_path = resolve_record_path(gate_raw_values.get("- CAD appendix binding audit path:", ""), record_path)
    try:
        report = json.loads(audit_path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"CAD appendix binding audit path must be readable JSON: {audit_path} ({exc})")
        return issues
    if report.get("schema") != "graduation-project-builder.docx-cad-appendix-binding-audit.v1":
        issues.append(f"CAD appendix binding audit has wrong schema: {audit_path}")
    if report.get("passed") is not True:
        issues.append("CAD appendix binding audit report must be passing")
    package_only_user_override = str(report.get("binding_mode") or "").strip().lower() == "package-only-user-override"
    if package_only_user_override:
        override_reason = str(report.get("user_override_reason") or "").strip().lower()
        if not override_reason or not contains_any(override_reason, {"user", "explicit", "cad", "package"}):
            issues.append("CAD appendix package-only override must record the explicit user/package reason")
    verdict = gate_values.get("- CAD appendix binding audit verdict:", "")
    if not _mechanical_cad_passish(verdict):
        issues.append("CAD appendix binding audit verdict must be pass-shaped")

    final_docx_sha = sha256_file(final_docx_path)
    recorded_docx_sha = str(report.get("docx_sha256", ""))
    field_docx_sha = gate_values.get("- CAD appendix final DOCX SHA256:", "")
    if recorded_docx_sha and recorded_docx_sha.upper() != final_docx_sha.upper():
        issues.append("CAD appendix binding audit docx_sha256 must match exact final DOCX")
    if field_docx_sha and is_explicit(field_docx_sha) and not is_explicit_none(field_docx_sha):
        if field_docx_sha.upper() != final_docx_sha.upper():
            issues.append("CAD appendix final DOCX SHA256 field must match exact final DOCX")

    matched_count = int(report.get("matched_cad_png_count") or 0)
    missing_count = int(report.get("missing_cad_png_count") or 0)
    cad_png_count = int(report.get("cad_png_count") or 0)
    if not package_only_user_override:
        if matched_count <= 0 or cad_png_count <= 0:
            issues.append("CAD appendix binding audit must prove at least one final CAD sheet render was embedded")
        if missing_count != 0:
            issues.append("CAD appendix binding audit must have zero missing final CAD sheet renders")
        if matched_count < cad_png_count:
            issues.append("CAD appendix binding audit matched sheet count must cover every CAD PNG sheet render")
    field_matched = _int_from_gate_value(gate_values.get("- CAD appendix matched sheet count:", ""))
    field_missing = _int_from_gate_value(gate_values.get("- CAD appendix missing sheet count:", ""))
    if field_matched != matched_count:
        issues.append("CAD appendix matched sheet count field must match the audit report")
    if package_only_user_override:
        if field_missing != missing_count:
            issues.append("CAD appendix missing sheet count field must match the package-only audit report")
    elif field_missing != 0 or field_missing != missing_count:
        issues.append("CAD appendix missing sheet count field must be 0 and match the audit report")

    if is_explicit(package_value) and not is_explicit_none(package_value):
        package_path = resolve_record_path(gate_raw_values.get("- CAD appendix final package path:", ""), record_path)
    else:
        package_path = Path(str(report.get("cad_source_path", "")))
    if not package_path.exists():
        issues.append(f"CAD appendix final package path does not exist: {package_path}")
    else:
        report_package_sha = str(report.get("cad_source_sha256", ""))
        if package_path.is_file() and report_package_sha:
            actual_package_sha = sha256_file(package_path)
            if report_package_sha.upper() != actual_package_sha.upper():
                issues.append("CAD appendix binding audit CAD package SHA256 must match the final package")
        if not package_only_user_override:
            try:
                live_report = audit_docx_cad_appendix_binding(
                    final_docx_path,
                    package_path,
                    min_cad_images_in_appendix=1,
                    require_all_cad_pngs=True,
                )
            except Exception as exc:
                issues.append(f"CAD appendix binding live audit could not run: {exc}")
            else:
                if live_report.get("passed") is not True:
                    issues.append("CAD appendix binding live audit must pass on exact final DOCX and CAD package")
                if int(live_report.get("matched_cad_png_count") or 0) != matched_count:
                    issues.append("CAD appendix binding audit matched count differs from live audit")
                if int(live_report.get("missing_cad_png_count") or 0) != missing_count:
                    issues.append("CAD appendix binding audit missing count differs from live audit")
    return issues


def check_gate_record(record_path: Path) -> list[str]:
    issues: list[str] = []
    if not record_path.exists():
        return [f"gate record not found: {record_path}"]

    record_lines = read_lines(record_path)
    record_text_lower = "\n".join(record_lines).lower()
    if (
        f"- acceptance schema: {MECHANICAL_CAD_ACCEPTANCE_SCHEMA_VALUE}" in record_text_lower
        or "# mechanical cad acceptance template" in record_text_lower
    ):
        return check_mechanical_cad_acceptance_record(record_path, record_lines)
    if (
        ("final acceptance" in record_text_lower or "final-acceptance" in record_path.name.lower())
        and contains_any(record_text_lower, SMOKE_SUMMARY_FINAL_ACCEPTANCE_TOKENS)
        and (
            contains_any(record_text_lower, PASS_SHAPED_HANDOFF_TOKENS)
            or sum(1 for token in SMOKE_SUMMARY_FINAL_ACCEPTANCE_TOKENS if token in record_text_lower) >= 3
        )
        and "# final acceptance template" not in record_text_lower
    ):
        issues.append(
            "smoke summary cannot be final acceptance; use assets/final-acceptance-template.md "
            "and validate the exact record with validate_skill_gate.py --gate-record"
        )
    for token in FORBIDDEN_PASS_WITH_CAVEAT_TOKENS:
        if token in record_text_lower:
            issues.append(f"gate record contains forbidden caveated-pass wording: {token}")
    for token in FORBIDDEN_STATIC_TOC_PASS_TOKENS:
        if token in record_text_lower:
            issues.append(f"gate record contains forbidden static-TOC fallback wording: {token}")
    normalized_lines = {normalize(line) for line in record_lines if normalize(line)}
    task_mode_lines = find_lines_with_prefix(record_lines, "- task mode:")
    task_mode = parse_line_value(task_mode_lines[0]) if task_mode_lines else ""
    selected_workflow_lines = find_lines_with_prefix(record_lines, "- selected thesis workflow:")
    selected_workflow = parse_line_value(selected_workflow_lines[0]) if selected_workflow_lines else ""
    if task_mode in THESIS_MODES and selected_workflow not in THESIS_WORKFLOWS:
        issues.append(
            "gate record selected thesis workflow must be one of "
            f"{', '.join(sorted(THESIS_WORKFLOWS))}: {selected_workflow or 'missing'}"
        )
    explicit_override_lines = find_lines_with_prefix(record_lines, "- explicit user overrides:")
    explicit_override_text = raw_line_value(explicit_override_lines[0]) if explicit_override_lines else ""
    paper_only_required = "paper-only literature" in normalize(explicit_override_text).lower()

    review_copy_path: Path | None = None
    final_docx_path: Path | None = None
    review_copy_promotion_binding = ""
    rendered_docx_lines = find_lines_with_prefix(record_lines, "- review-copy path actually rendered:")
    if rendered_docx_lines:
        rendered_docx_value = parse_line_value(rendered_docx_lines[0])
        if is_explicit(rendered_docx_value) and not is_explicit_none(rendered_docx_value):
            review_copy_path = resolve_record_path(raw_line_value(rendered_docx_lines[0]), record_path)
    promotion_binding_lines = find_lines_with_prefix(record_lines, "- review-copy promotion binding:")
    if promotion_binding_lines:
        review_copy_promotion_binding = parse_line_value(promotion_binding_lines[0])
    exact_output_lines = find_lines_with_prefix(record_lines, "- exact output paths:")
    if exact_output_lines:
        for raw_path in split_path_values(raw_line_value(exact_output_lines[0])):
            resolved = resolve_record_path(raw_path, record_path)
            if resolved.suffix.lower() == ".docx":
                final_docx_path = resolved
                break
    if task_mode in THESIS_MODES and final_docx_path is None:
        issues.append("thesis gate record exact output paths must name the final DOCX handed off")
    final_docx_has_citation_surface = False
    final_citation_surface_removed = False
    if task_mode in THESIS_MODES and final_docx_path is not None and final_docx_path.exists():
        try:
            final_citation_audit = audit_body_citations(final_docx_path)
        except Exception as exc:
            issues.append(f"thesis gate record final DOCX citation audit could not run: {exc}")
        else:
            final_docx_has_citation_surface = (
                final_citation_audit.body_citation_paragraph_count > 0
                or final_citation_audit.bibliography_item_count > 0
            )
            if not final_citation_audit.passed:
                error_codes = ", ".join(final_citation_audit.error_codes) or "unknown"
                issues.append(
                    "thesis gate record final DOCX canonical citation audit failed on the exact output "
                    f"({error_codes}); body citation superscript/order/coupled-chain verdict text cannot override "
                    "scripts/audit_thesis_citations.py"
                )
            if not final_docx_has_citation_surface and contains_any(
                record_text_lower,
                {"citation", "citations", "bibliography", "reference", "\u5f15\u7528", "\u53c2\u8003\u6587\u732e"},
                ):
                final_citation_surface_removed = True
    live_toc_required_lines = find_lines_with_prefix(record_lines, "- live TOC required this round?:")
    live_toc_required_value = parse_line_value(live_toc_required_lines[0]) if live_toc_required_lines else ""
    live_toc_required = live_toc_required_value.lower() in {"yes", "true", "required", "pass"}
    if (
        task_mode in THESIS_MODES
        and selected_workflow in {"new-thesis-production", "whole-thesis-revision"}
        and live_toc_required_value.lower() in {"", "no", "false", "not-applicable", "none"}
    ):
        issues.append("whole-thesis thesis gate records must preserve live TOC requirement instead of marking it not required")
        live_toc_required = True
    if task_mode in THESIS_MODES and live_toc_required and final_docx_path is not None:
        toc_field_state = inspect_live_toc_fields(final_docx_path)
        live_toc_count = toc_field_state["count"]
        locked_toc_count = toc_field_state["locked_count"]
        if live_toc_count <= 0:
            issues.append("live TOC is required but the final DOCX has no standard TOC field instruction")
        elif locked_toc_count < live_toc_count:
            issues.append(
                "live TOC is required but not every standard TOC field is locked with w:fldLock=true"
            )
        live_toc_field_verdict_lines = find_lines_with_prefix(record_lines, "- live TOC field verdict:")
        if live_toc_field_verdict_lines:
            verdict_value = parse_line_value(live_toc_field_verdict_lines[0]).lower()
            if live_toc_count <= 0 and "pass" in verdict_value:
                issues.append("live TOC field verdict is pass-shaped while the final DOCX has no TOC field")
            if locked_toc_count < live_toc_count and "pass" in verdict_value:
                issues.append("live TOC field verdict is pass-shaped while the final DOCX TOC field is not locked")
    whole_format_path_prefix = "- final DOCX whole-format structural audit path:"
    whole_format_verdict_prefix = "- final DOCX whole-format structural audit verdict:"
    whole_format_required = task_mode in THESIS_MODES and selected_workflow in {
        "new-thesis-production",
        "whole-thesis-revision",
    }
    whole_format_path_lines = find_lines_with_prefix(record_lines, whole_format_path_prefix)
    whole_format_verdict_lines = find_lines_with_prefix(record_lines, whole_format_verdict_prefix)
    if whole_format_required:
        if len(whole_format_path_lines) != 1:
            issues.append(
                f"whole-thesis gate records must contain exactly one '{whole_format_path_prefix}' line"
            )
        if len(whole_format_verdict_lines) != 1:
            issues.append(
                f"whole-thesis gate records must contain exactly one '{whole_format_verdict_prefix}' line"
            )
    if len(whole_format_path_lines) > 1:
        issues.append(f"gate record must contain at most one '{whole_format_path_prefix}' line")
    if len(whole_format_verdict_lines) > 1:
        issues.append(f"gate record must contain at most one '{whole_format_verdict_prefix}' line")
    if whole_format_path_lines:
        whole_format_value = parse_line_value(whole_format_path_lines[0])
        whole_format_raw = raw_line_value(whole_format_path_lines[0])
        if whole_format_required and (not is_explicit(whole_format_value) or is_explicit_none(whole_format_value)):
            issues.append("whole-thesis final acceptance must bind an exact-output whole-format structural audit path")
        elif is_explicit(whole_format_value) and not is_explicit_none(whole_format_value):
            whole_format_report = first_resolved_path(whole_format_raw, record_path)
            if whole_format_report is None:
                issues.append("whole-format structural audit path could not be resolved")
            else:
                issues.extend(check_docx_whole_format_gate_report(whole_format_report, final_docx_path))
    if whole_format_verdict_lines:
        whole_format_verdict = parse_line_value(whole_format_verdict_lines[0]).lower()
        if whole_format_required and "pass" not in whole_format_verdict:
            issues.append("whole-thesis final acceptance whole-format structural audit verdict must be pass-shaped")
    list_pollution_path_prefix = "- final DOCX list-pollution audit path:"
    list_pollution_verdict_prefix = "- final DOCX list-pollution audit verdict:"
    whole_format_list_pollution_prefix = "- whole-format list_pollution_audit verdict:"
    list_pollution_required = task_mode in THESIS_MODES
    list_pollution_path_lines = find_lines_with_prefix(record_lines, list_pollution_path_prefix)
    list_pollution_verdict_lines = find_lines_with_prefix(record_lines, list_pollution_verdict_prefix)
    whole_format_list_pollution_lines = find_lines_with_prefix(record_lines, whole_format_list_pollution_prefix)
    if list_pollution_required:
        if len(list_pollution_path_lines) != 1:
            issues.append(f"thesis gate records must contain exactly one '{list_pollution_path_prefix}' line")
        if len(list_pollution_verdict_lines) != 1:
            issues.append(f"thesis gate records must contain exactly one '{list_pollution_verdict_prefix}' line")
        if len(whole_format_list_pollution_lines) != 1:
            issues.append(f"thesis gate records must contain exactly one '{whole_format_list_pollution_prefix}' line")
    if len(list_pollution_path_lines) > 1:
        issues.append(f"gate record must contain at most one '{list_pollution_path_prefix}' line")
    if len(list_pollution_verdict_lines) > 1:
        issues.append(f"gate record must contain at most one '{list_pollution_verdict_prefix}' line")
    if len(whole_format_list_pollution_lines) > 1:
        issues.append(f"gate record must contain at most one '{whole_format_list_pollution_prefix}' line")
    if list_pollution_path_lines:
        list_pollution_value = parse_line_value(list_pollution_path_lines[0])
        list_pollution_raw = raw_line_value(list_pollution_path_lines[0])
        if list_pollution_required and (
            not is_explicit(list_pollution_value) or is_explicit_none(list_pollution_value)
        ):
            issues.append("thesis final acceptance must bind an exact-output list-pollution audit path")
        elif is_explicit(list_pollution_value) and not is_explicit_none(list_pollution_value):
            list_pollution_report = first_resolved_path(list_pollution_raw, record_path)
            if list_pollution_report is None:
                issues.append("list-pollution audit path could not be resolved")
            else:
                issues.extend(check_docx_list_pollution_audit_report(list_pollution_report, final_docx_path))
    if list_pollution_verdict_lines:
        list_pollution_verdict = parse_line_value(list_pollution_verdict_lines[0]).lower()
        if list_pollution_required and "pass" not in list_pollution_verdict:
            issues.append("thesis final acceptance list-pollution audit verdict must be pass-shaped")
    if whole_format_list_pollution_lines:
        whole_format_list_pollution_verdict = parse_line_value(whole_format_list_pollution_lines[0]).lower()
        if list_pollution_required and "pass" not in whole_format_list_pollution_verdict:
            issues.append("thesis final acceptance whole-format list_pollution_audit verdict must be pass-shaped")
    font_color_path_prefix = "- final DOCX font-color audit path:"
    font_color_verdict_prefix = "- final DOCX font-color audit verdict:"
    font_color_required = task_mode in THESIS_MODES
    font_color_path_lines = find_lines_with_prefix(record_lines, font_color_path_prefix)
    font_color_verdict_lines = find_lines_with_prefix(record_lines, font_color_verdict_prefix)
    if font_color_required:
        if len(font_color_path_lines) != 1:
            issues.append(f"thesis gate records must contain exactly one '{font_color_path_prefix}' line")
        if len(font_color_verdict_lines) != 1:
            issues.append(f"thesis gate records must contain exactly one '{font_color_verdict_prefix}' line")
    if len(font_color_path_lines) > 1:
        issues.append(f"gate record must contain at most one '{font_color_path_prefix}' line")
    if len(font_color_verdict_lines) > 1:
        issues.append(f"gate record must contain at most one '{font_color_verdict_prefix}' line")
    if font_color_path_lines:
        font_color_value = parse_line_value(font_color_path_lines[0])
        font_color_raw = raw_line_value(font_color_path_lines[0])
        if font_color_required and (not is_explicit(font_color_value) or is_explicit_none(font_color_value)):
            issues.append("thesis final acceptance must bind an exact-output font-color audit path")
        elif is_explicit(font_color_value) and not is_explicit_none(font_color_value):
            font_color_report = first_resolved_path(font_color_raw, record_path)
            if font_color_report is None:
                issues.append("font-color audit path could not be resolved")
            else:
                issues.extend(check_docx_font_color_audit_report(font_color_report, final_docx_path))
    if font_color_verdict_lines:
        font_color_verdict = parse_line_value(font_color_verdict_lines[0]).lower()
        if font_color_required and "pass" not in font_color_verdict:
            issues.append("thesis final acceptance font-color audit verdict must be pass-shaped")
    if task_mode in THESIS_MODES and review_copy_path is not None and final_docx_path is not None:
        try:
            if review_copy_path.resolve() != final_docx_path.resolve():
                binding_text = review_copy_promotion_binding.lower()
                binding_names_paths = (
                    review_copy_path.name in review_copy_promotion_binding
                    and final_docx_path.name in review_copy_promotion_binding
                )
                binding_claims_promotion = any(
                    token in binding_text
                    for token in ("promotion", "promoted", "handoff", "delivery", "final", "\u4ea4\u4ed8", "\u63d0\u5347")
                )
                if (
                    not is_explicit(review_copy_promotion_binding)
                    or is_explicit_none(review_copy_promotion_binding)
                    or not binding_names_paths
                    or not binding_claims_promotion
                ):
                    issues.append(
                        "thesis gate record review-copy path actually rendered must match exact output DOCX "
                        "or record explicit review-copy promotion binding; review copies cannot substitute "
                        "for the handed-off final DOCX"
                    )
        except OSError:
            issues.append("thesis gate record review-copy/final DOCX paths could not be resolved for exact-output binding")
    rendered_docx_path = final_docx_path
    if task_mode in THESIS_MODES and rendered_docx_path is not None and rendered_docx_path.exists():
        source_review_artifact_lines = find_lines_with_prefix(record_lines, "- source review-artifact inventory path:")
        source_docx_for_initial_figure_check = None
        if source_review_artifact_lines:
            source_review_artifact_value = parse_line_value(source_review_artifact_lines[0])
            if is_explicit(source_review_artifact_value) and not is_explicit_none(source_review_artifact_value):
                source_review_artifact_path = first_resolved_path(
                    raw_line_value(source_review_artifact_lines[0]),
                    record_path,
                )
                source_docx_for_initial_figure_check = review_inventory_source_docx_path(source_review_artifact_path)
        issues.extend(
            final_docx_figure_surface_issues(
                rendered_docx_path,
                source_docx=source_docx_for_initial_figure_check,
                source_docx_role="format_template" if source_docx_for_initial_figure_check else "",
            )
        )

    for marker in FINAL_ACCEPTANCE_SCHEMA["headings"]:
        if marker not in normalized_lines:
            issues.append(f"gate record missing marker: {marker}")

    for prefix, expected_count in FINAL_ACCEPTANCE_SCHEMA["repeated_prefix_counts"].items():
        lines = find_lines_with_prefix(record_lines, prefix)
        if len(lines) != expected_count:
            issues.append(
                f"gate record must contain {expected_count} occurrences of '{prefix}' but found {len(lines)}"
            )
            continue
        for line in lines:
            value = parse_line_value(line)
            if prefix == "- failed:":
                if not is_explicit_none(value):
                    issues.append(f"gate record still has failed checks: {normalize(line)}")
            else:
                if value not in EXPLICIT_VALUES and value in PLACEHOLDER_VALUES:
                    issues.append(f"gate record has incomplete field: {normalize(line)}")

    mechanical_cad_scope = _mechanical_cad_final_acceptance_scope_from_record(record_lines)
    active_single_prefixes = [
        prefix
        for prefix in FINAL_ACCEPTANCE_SCHEMA["single_prefixes"]
        if task_mode in THESIS_MODES or prefix not in THESIS_ONLY_SOURCE_ROLE_PREFIXES
        if mechanical_cad_scope or not _is_mechanical_cad_final_acceptance_prefix(prefix)
    ]

    for prefix in active_single_prefixes:
        lines = find_lines_with_prefix(record_lines, prefix)
        if len(lines) != 1:
            issues.append(f"gate record must contain exactly one '{prefix}' line")
            continue
        issues.extend(
            validate_gate_single_prefix(
                prefix=prefix,
                line=lines[0],
                value=parse_line_value(lines[0]),
                task_mode=task_mode,
                paper_only_required=paper_only_required,
                record_lines=record_lines,
                record_path=record_path,
            )
        )

    gate_values = {
        prefix: parse_line_value(find_lines_with_prefix(record_lines, prefix)[0])
        for prefix in active_single_prefixes
        if find_lines_with_prefix(record_lines, prefix)
    }
    gate_raw_values = {
        prefix: raw_line_value(find_lines_with_prefix(record_lines, prefix)[0])
        for prefix in active_single_prefixes
        if find_lines_with_prefix(record_lines, prefix)
    }
    for prefix in FINAL_ACCEPTANCE_SCHEMA.get("repeated_prefix_counts", {}):
        if prefix in {"- failed:", "- passed:", "- skipped:"}:
            continue
        lines = find_lines_with_prefix(record_lines, prefix)
        if lines:
            gate_values.setdefault(prefix, parse_line_value(lines[0]))
            gate_raw_values.setdefault(prefix, raw_line_value(lines[0]))
    if any(
        len(find_lines_with_prefix(record_lines, prefix)) != 1
        for prefix in active_single_prefixes
    ):
        return issues
    rendered_pdf_path = (
        first_resolved_path(gate_raw_values.get("- rendered PDF path:", ""), record_path)
        if is_explicit(gate_values.get("- rendered PDF path:", ""))
        and not is_explicit_none(gate_values.get("- rendered PDF path:", ""))
        else None
    )
    expected_surface_output_path = rendered_docx_path if task_mode in THESIS_MODES else None
    expected_surface_output_sha256 = gate_values.get("- protected-surface reviewed output sha256:", "").strip()
    if task_mode in THESIS_MODES:
        issues.extend(
            _validate_cad_appendix_binding_record(
                record_path=record_path,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                final_docx_path=final_docx_path,
                record_text_lower=record_text_lower,
            )
        )
    issues.extend(
        validate_skill_invocation_lock(
            record_path=record_path,
            gate_values=gate_values,
            gate_raw_values=gate_raw_values,
            task_mode=task_mode,
            selected_workflow=selected_workflow,
            final_docx_path=rendered_docx_path,
            expected_output_sha256=expected_surface_output_sha256,
        )
    )
    if task_mode in THESIS_MODES:
        issues.extend(validate_blocking_state_consistency(gate_values, location=record_path))
        issues.extend(
            validate_baseline_promotion_state(
                gate_values,
                task_mode=task_mode,
                selected_workflow=selected_workflow,
                record_path=record_path,
            )
        )
    if task_mode in THESIS_MODES:
        citation_audit_final_sha = gate_values.get("- citation audit final DOCX SHA256:", "").strip()
        er_structural_prefixes = (
            "- structural source-scale bbox map paths:",
            "- structural inserted-scale geometry evidence paths:",
            "- structural dense-zone crop evidence paths:",
            "- inserted-scale collision evidence paths:",
            "- structural relation-attribute collision verdict:",
            "- structural shape-overlap verdict:",
            "- structural inserted-scale collision verdict:",
            "- structural source-to-inserted geometry verdict:",
        )
        for prefix in er_structural_prefixes:
            lines = find_lines_with_prefix(record_lines, prefix)
            if len(lines) > 1:
                issues.append(f"gate record must contain at most one '{prefix}' line")
            elif len(lines) == 1:
                gate_values[prefix] = parse_line_value(lines[0])
                gate_raw_values[prefix] = raw_line_value(lines[0])

        if expected_surface_output_path is None:
            issues.append("thesis gate record must bind protected-surface evidence to a rendered final DOCX path")
        elif expected_surface_output_path.exists():
            actual_output_sha256 = sha256_file(expected_surface_output_path)
            if len(expected_surface_output_sha256) != 64 or any(
                ch not in "0123456789abcdefABCDEF" for ch in expected_surface_output_sha256
            ):
                issues.append("thesis gate record protected-surface reviewed output sha256 must be a 64-hex value")
            elif actual_output_sha256.lower() != expected_surface_output_sha256.lower():
                issues.append(
                    "thesis gate record protected-surface reviewed output sha256 does not match the rendered final DOCX"
                )
            if len(citation_audit_final_sha) != 64 or any(
                ch not in "0123456789abcdefABCDEF" for ch in citation_audit_final_sha
            ):
                issues.append("thesis gate record citation audit final DOCX SHA256 must be a 64-hex value")
            elif actual_output_sha256.lower() != citation_audit_final_sha.lower():
                issues.append(
                    "thesis gate record citation audit final DOCX SHA256 does not match the rendered final DOCX"
                )
        source_review_artifact_inventory_path = first_resolved_path(
            gate_raw_values.get("- source review-artifact inventory path:", ""), record_path
        ) if is_explicit(gate_values.get("- source review-artifact inventory path:", "")) and not is_explicit_none(gate_values.get("- source review-artifact inventory path:", "")) else None
        final_review_artifact_diff_path = first_resolved_path(
            gate_raw_values.get("- final review-artifact diff path:", ""), record_path
        ) if is_explicit(gate_values.get("- final review-artifact diff path:", "")) and not is_explicit_none(gate_values.get("- final review-artifact diff path:", "")) else None
        source_body_citation_run_inventory_path = first_resolved_path(
            gate_raw_values.get("- source body-citation run inventory path:", ""), record_path
        ) if is_explicit(gate_values.get("- source body-citation run inventory path:", "")) and not is_explicit_none(gate_values.get("- source body-citation run inventory path:", "")) else None
        final_body_citation_run_diff_path = first_resolved_path(
            gate_raw_values.get("- final body-citation run diff path:", ""), record_path
        ) if is_explicit(gate_values.get("- final body-citation run diff path:", "")) and not is_explicit_none(gate_values.get("- final body-citation run diff path:", "")) else None
        citation_audit_source_to_final_run_diff_path = first_resolved_path(
            gate_raw_values.get("- citation audit source-to-final run diff path:", ""), record_path
        ) if is_explicit(gate_values.get("- citation audit source-to-final run diff path:", "")) and not is_explicit_none(gate_values.get("- citation audit source-to-final run diff path:", "")) else None
        comment_resolution_ledger_path = first_resolved_path(
            gate_raw_values.get("- comment-resolution ledger path:", ""), record_path
        ) if is_explicit(gate_values.get("- comment-resolution ledger path:", "")) and not is_explicit_none(gate_values.get("- comment-resolution ledger path:", "")) else None
        comment_resolution_audit_report_path = first_resolved_path(
            gate_raw_values.get("- comment-resolution audit report path:", ""), record_path
        ) if is_explicit(gate_values.get("- comment-resolution audit report path:", "")) and not is_explicit_none(gate_values.get("- comment-resolution audit report path:", "")) else None
        comment_resolution_source_docx_path = first_resolved_path(
            gate_raw_values.get("- comment-resolution source DOCX path:", ""), record_path
        ) if is_explicit(gate_values.get("- comment-resolution source DOCX path:", "")) and not is_explicit_none(gate_values.get("- comment-resolution source DOCX path:", "")) else None
        if source_review_artifact_inventory_path is not None and final_review_artifact_diff_path is not None:
            issues.extend(
                validate_review_artifact_reports(
                    source_review_artifact_inventory_path,
                    final_review_artifact_diff_path,
                    expected_final_docx=rendered_docx_path,
                )
            )
        source_docx_has_citation_surface = False
        citation_inventory_source_docx = citation_inventory_source_docx_path(source_body_citation_run_inventory_path)
        if citation_inventory_source_docx is None:
            citation_inventory_source_docx = review_inventory_source_docx_path(source_review_artifact_inventory_path)
        if citation_inventory_source_docx is not None and citation_inventory_source_docx.exists():
            try:
                source_citation_audit = audit_body_citations(citation_inventory_source_docx)
            except Exception as exc:
                issues.append(f"thesis gate record source DOCX citation audit could not run: {exc}")
            else:
                source_docx_has_citation_surface = (
                    source_citation_audit.body_citation_paragraph_count > 0
                    or source_citation_audit.bibliography_item_count > 0
                )
        citation_field_presence = any(
            find_lines_with_prefix(record_lines, prefix)
            for prefix in (
                "- source body-citation run inventory path:",
                "- final body-citation run diff path:",
                "- citation audit source-to-final run diff path:",
                "- citation audit evidence path:",
                "- citation audit path:",
                "- body citation superscripts preservation verdict:",
                "- citation-reference coupled-chain verdict:",
            )
        )
        citation_evidence_required = (
            task_mode in THESIS_MODES
            and (
                citation_field_presence
                or final_docx_has_citation_surface
                or source_docx_has_citation_surface
                or final_citation_surface_removed
            )
        )
        if citation_evidence_required:
            verdict = gate_values.get("- body citation superscripts preservation verdict:", "")
            if not is_explicit(verdict) or is_explicit_none(verdict) or not contains_any(verdict, {"pass", "passed"}):
                issues.append(
                    "thesis gate record with citations must record a passing body citation superscripts preservation verdict"
                )
            citation_audit_report_path = None
            citation_audit_value = gate_values.get("- citation audit evidence path:", "")
            if is_explicit(citation_audit_value) and not is_explicit_none(citation_audit_value):
                citation_audit_report_path = first_resolved_path(
                    gate_raw_values.get("- citation audit evidence path:", ""),
                    record_path,
                )
            if citation_audit_report_path is None:
                issues.append(
                    "thesis gate record with citations must name a canonical citation audit evidence path "
                    "generated by scripts/audit_thesis_citations.py; text verdicts are not sufficient"
                )
            else:
                issues.extend(
                    check_thesis_citation_audit_report(
                        citation_audit_report_path,
                        expected_docx_path=rendered_docx_path,
                    )
                )
            citation_anchor_pollution_report_path = None
            citation_anchor_pollution_value = gate_values.get("- citation anchor pollution audit path:", "")
            if is_explicit(citation_anchor_pollution_value) and not is_explicit_none(citation_anchor_pollution_value):
                citation_anchor_pollution_report_path = first_resolved_path(
                    gate_raw_values.get("- citation anchor pollution audit path:", ""),
                    record_path,
                )
            if citation_anchor_pollution_report_path is None:
                issues.append(
                    "thesis gate record with citations must name a canonical citation anchor pollution audit path "
                    "generated by scripts/audit_docx_citation_anchor_pollution.py"
                )
            else:
                issues.extend(
                    check_docx_citation_anchor_pollution_audit_report(
                        citation_anchor_pollution_report_path,
                        expected_docx_path=rendered_docx_path,
                        expected_pdf_path=rendered_pdf_path,
                    )
                )
            citation_anchor_verdict = gate_values.get("- citation anchor pollution verdict:", "")
            if (
                not is_explicit(citation_anchor_verdict)
                or is_explicit_none(citation_anchor_verdict)
                or not contains_any(citation_anchor_verdict, {"pass", "passed"})
                or _contains_blocking_status_token(citation_anchor_verdict)
            ):
                issues.append("thesis gate record with citations must record a passing citation anchor pollution verdict")
            for prefix in (
                "- citation anchor visible DOCX hit count:",
                "- citation anchor field-result hit count:",
                "- citation anchor rendered PDF hit count:",
            ):
                raw_count = normalize(gate_values.get(prefix, ""))
                if raw_count in {"none", "not-applicable", "not applicable"} and prefix == "- citation anchor rendered PDF hit count:" and rendered_pdf_path is None:
                    continue
                if not re.fullmatch(r"\d+", raw_count):
                    issues.append(f"thesis gate record citation anchor pollution count must be a nonnegative integer: {prefix}")
                elif int(raw_count) != 0:
                    issues.append(f"thesis gate record citation anchor pollution count must be zero: {prefix} {raw_count}")
            citation_anchor_final_sha = gate_values.get("- citation anchor pollution final DOCX SHA256:", "").strip()
            if len(citation_anchor_final_sha) != 64 or any(
                ch not in "0123456789abcdefABCDEF" for ch in citation_anchor_final_sha
            ):
                issues.append("thesis gate record citation anchor pollution final DOCX SHA256 must be a 64-hex value")
            elif rendered_docx_path is not None and rendered_docx_path.exists():
                actual_anchor_docx_sha = sha256_file(rendered_docx_path)
                if actual_anchor_docx_sha.lower() != citation_anchor_final_sha.lower():
                    issues.append(
                        "thesis gate record citation anchor pollution final DOCX SHA256 does not match the rendered final DOCX"
                    )
            coupled_verdict = gate_values.get("- citation-reference coupled-chain verdict:", "")
            if (
                not is_explicit(coupled_verdict)
                or is_explicit_none(coupled_verdict)
                or not contains_any(coupled_verdict, {"pass", "passed"})
                or _contains_blocking_status_token(coupled_verdict)
            ):
                issues.append(
                    "thesis gate record with citations must record a passing citation-reference coupled-chain verdict"
                )
            if source_body_citation_run_inventory_path is None:
                issues.append(
                    "thesis gate record with citations must name a source body-citation run inventory path"
                )
            if final_body_citation_run_diff_path is None:
                issues.append("thesis gate record with citations must name a final body-citation run diff path")
            if citation_audit_source_to_final_run_diff_path is None:
                issues.append(
                    "thesis gate record with citations must name a citation audit source-to-final run diff path"
                )
        if source_body_citation_run_inventory_path is not None and final_body_citation_run_diff_path is not None:
            allowed_citation_preservation_scopes = {STRICT_CITATION_PRESERVATION_SCOPE}
            if selected_workflow in WHOLE_REBUILD_CITATION_SCOPE_WORKFLOWS:
                allowed_citation_preservation_scopes.add(WHOLE_REBUILD_CITATION_PRESERVATION_SCOPE)
                allowed_citation_preservation_scopes.add(APPROVED_NONPRESERVATION_CITATION_SCOPE)
            elif selected_workflow in STRICT_CITATION_SCOPE_WORKFLOWS:
                allowed_citation_preservation_scopes = {STRICT_CITATION_PRESERVATION_SCOPE}
            issues.extend(
                validate_citation_run_reports(
                    source_body_citation_run_inventory_path,
                    final_body_citation_run_diff_path,
                    expected_final_docx=rendered_docx_path,
                    allowed_preservation_scopes=allowed_citation_preservation_scopes,
                )
            )
        if citation_audit_source_to_final_run_diff_path is not None and final_body_citation_run_diff_path is not None:
            if citation_audit_source_to_final_run_diff_path.resolve() != final_body_citation_run_diff_path.resolve():
                issues.append(
                    "thesis gate record citation audit source-to-final run diff path must match the final body-citation run diff path"
                )
        if rendered_docx_path is not None and rendered_docx_path.exists():
            comment_snapshot = collect_comment_snapshot(rendered_docx_path)
            subtask_for_comments = raw_line_value(find_lines_with_prefix(record_lines, "- subtask:")[0])
            verification_scope_for_comments = gate_values.get("- verification scope claim:", "")
            comment_context = "\n".join(
                [
                    verification_scope_for_comments,
                    explicit_override_text,
                    subtask_for_comments,
                    gate_values.get("- review comments/change marks verdict:", ""),
                    gate_values.get("- review comments/change marks preservation verdict:", ""),
                    gate_values.get("- handoff status:", ""),
                    gate_values.get("- audit verdict:", ""),
                ]
            )
            all_comments_claim = has_all_comments_claim(comment_context)
            comment_revision_context = comment_snapshot.comment_count > 0 and contains_any(
                comment_context.lower(),
                {"comment", "comments", "teacher comment", "review comment", "\u6279\u6ce8"},
            )
            if comment_snapshot.done_count and comment_resolution_ledger_path is None:
                issues.append("final DOCX has done comments but gate record lacks comment-resolution ledger path")
            if (all_comments_claim or comment_revision_context) and comment_resolution_ledger_path is None:
                issues.append("comment-driven or all-comments-resolved gate record lacks comment-resolution ledger path")
            if comment_resolution_ledger_path is not None:
                comment_resolution_source_docx = (
                    comment_resolution_source_docx_path
                    or review_inventory_source_docx_path(source_review_artifact_inventory_path)
                )
                issues.extend(
                    validate_comment_resolution_ledger(
                        comment_resolution_ledger_path,
                        final_docx=rendered_docx_path,
                        source_docx=comment_resolution_source_docx,
                        assert_all_resolved=all_comments_claim or comment_revision_context,
                    )
                )
                if comment_resolution_audit_report_path is None:
                    issues.append("comment-resolution ledger requires a comment-resolution audit report path")
                else:
                    issues.extend(validate_existing_path(comment_resolution_audit_report_path, require_nonempty_file=True))
                if not surface_verdict_passes(gate_values.get("- comment-resolution audit verdict:", "")):
                    issues.append("comment-resolution audit verdict must be pass when a comment-resolution ledger is present")
    agent_mode = gate_values.get("- agent mode:", "")
    agent_authorization_source = gate_values.get("- agent authorization source:", "")
    max_concurrent_live_agents = gate_values.get("- max concurrent live agents:", "")
    live_agent_count_plan = gate_values.get("- live agent count plan:", "")
    dispatch_wave_plan = gate_values.get("- dispatch wave plan:", "")
    audit_presence_by_wave = gate_values.get("- audit presence by wave:", "")
    concurrency_limit_verdict = gate_values.get("- concurrency limit verdict:", "")
    required_lanes = gate_values.get("- required lanes:", "")
    spawned_agent_ids = gate_values.get("- spawned agent ids:", "")
    spawned_agent_aliases = gate_values.get("- spawned agent aliases zh:", "")
    sequential_fallback_reason = gate_values.get("- sequential fallback reason:", "")
    audit_role_alias = gate_values.get("- audit role alias zh:", "")
    audit_agent_id = gate_values.get("- audit agent id:", "")
    sequential_audit_fallback_id = gate_values.get("- sequential audit fallback id:", "")
    agent_run_manifest_value = gate_values.get("- agent run manifest path:", "")
    lightweight_action_audit_value = gate_values.get("- lightweight action-audit entry path:", "")
    lane_task_card_value = gate_values.get("- lane task card evidence paths:", "")
    if agent_mode not in {"single-agent-no-auth", "parallel-subagents", "sequential-fallback"}:
        issues.append(f"gate record agent mode is not a recognized GPB dispatch mode: {agent_mode or 'missing'}")
    if audit_role_alias != "\u5ba1\u6838":
        issues.append("gate record audit role alias zh must be \u5ba1\u6838")
    issues.extend(
        validate_agent_concurrency_fields(
            max_live_agents=max_concurrent_live_agents,
            live_agent_count_plan=live_agent_count_plan,
            dispatch_wave_plan=dispatch_wave_plan,
            audit_presence_by_wave=audit_presence_by_wave,
            concurrency_limit_verdict=concurrency_limit_verdict,
            kind="gate record",
        )
    )
    if required_lanes != "none":
        for lane in CANONICAL_ROLE_LANES:
            if lane not in required_lanes:
                issues.append(f"gate record required lanes must include full canonical role roster: missing {lane}")
    if (
        (not is_explicit(agent_run_manifest_value) or is_explicit_none(agent_run_manifest_value))
        and (not is_explicit(lightweight_action_audit_value) or is_explicit_none(lightweight_action_audit_value))
    ):
        issues.append("gate record must name an agent run manifest path or lightweight action-audit entry path")
    if not is_explicit(lane_task_card_value) or is_explicit_none(lane_task_card_value):
        issues.append("gate record must name lane task card evidence paths")
    if is_explicit_none(audit_agent_id) and is_explicit_none(sequential_audit_fallback_id):
        issues.append("gate record must name audit agent id or sequential audit fallback id")
    if agent_mode == "single-agent-no-auth":
        if not is_explicit_none(agent_authorization_source):
            issues.append("single-agent-no-auth gate record must use agent authorization source: none")
        if not is_explicit_none(spawned_agent_ids):
            issues.append("single-agent-no-auth gate record must not claim spawned agent ids")
        if not is_explicit_none(audit_agent_id):
            issues.append("single-agent-no-auth gate record must not claim spawned audit agent id; use sequential audit fallback id")
        if is_explicit_none(sequential_audit_fallback_id):
            issues.append("single-agent-no-auth gate record must name sequential audit fallback id")
    elif agent_mode == "parallel-subagents":
        if is_explicit_none(agent_authorization_source):
            issues.append("parallel-subagents gate record must name the explicit authorization source")
        if is_explicit_none(spawned_agent_ids) or is_explicit_none(spawned_agent_aliases):
            issues.append("parallel-subagents gate record must name spawned agent ids and aliases")
        if not is_explicit_none(sequential_fallback_reason):
            issues.append("parallel-subagents gate record must keep sequential fallback reason as none")
    elif agent_mode == "sequential-fallback":
        if is_explicit_none(agent_authorization_source):
            issues.append("sequential-fallback gate record must name the explicit authorization source")
        if is_explicit_none(sequential_fallback_reason):
            issues.append("sequential-fallback gate record must name the fallback reason")
        if not is_explicit_none(audit_agent_id):
            issues.append("sequential-fallback gate record must not claim spawned audit agent id; use sequential audit fallback id")
        if is_explicit_none(sequential_audit_fallback_id):
            issues.append("sequential-fallback gate record must name sequential audit fallback id")

    issues.extend(
        validate_agent_records(
            record_path=record_path,
            gate_values=gate_values,
            gate_raw_values=gate_raw_values,
            rendered_docx_path=rendered_docx_path,
        )
    )

    if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}:
        issues.extend(
            validate_template_lock_fields(
                record_path=record_path,
                values=gate_values,
                raw_values=gate_raw_values,
                record_kind="gate record",
                explicit_source_text=explicit_override_text,
            )
        )
        transaction_owner = gate_values.get("- thesis mutation transaction owner path:", "")
        transaction_record_value = gate_values.get("- thesis mutation transaction record path:", "")
        transaction_validator_result = gate_values.get("- thesis mutation transaction validator result:", "")
        non_target_change_verdict = gate_values.get("- non-target protected surface change verdict:", "")
        local_scope_claim_verdict = gate_values.get("- local-surface whole-thesis claim verdict:", "")
        transaction_final_sha = gate_values.get("- thesis mutation transaction final docx sha256:", "")
        if "thesis-mutation-transaction.md" not in transaction_owner:
            issues.append("thesis gate record must route DOCX mutation through thesis-mutation-transaction.md")
        transaction_required = selected_workflow != "audit-only"
        transaction_record_path = first_resolved_path(
            gate_raw_values.get("- thesis mutation transaction record path:", ""), record_path
        ) if is_explicit(transaction_record_value) and not is_explicit_none(transaction_record_value) else None
        if transaction_required and transaction_record_path is None:
            issues.append("thesis DOCX mutation requires a thesis mutation transaction record path")
        if transaction_record_path is not None:
            issues.extend(
                validate_transaction_record(
                    transaction_record_path,
                    expected_final_docx=rendered_docx_path,
                )
            )
            gate_manifest_path = first_resolved_path(
                gate_raw_values.get("- figure asset manifest path:", ""), record_path
            )
            transaction_manifest_paths = transaction_figure_manifest_paths(transaction_record_path)
            if gate_manifest_path is not None and transaction_manifest_paths:
                try:
                    gate_manifest_resolved = gate_manifest_path.resolve()
                    if not any(path.resolve() == gate_manifest_resolved for path in transaction_manifest_paths):
                        issues.append("figure asset manifest path differs from thesis mutation transaction record")
                except OSError:
                    issues.append("figure asset manifest path cannot be resolved for transaction binding")
        if transaction_required:
            if not surface_verdict_passes(transaction_validator_result):
                issues.append("thesis mutation transaction validator result must be pass")
            if not surface_verdict_passes(non_target_change_verdict):
                issues.append("thesis mutation transaction must pass non-target protected surface change verdict")
            if selected_workflow == "local-surface-repair" and not surface_verdict_passes(local_scope_claim_verdict):
                issues.append("local-surface transaction must explicitly reject whole-thesis claim promotion")
            if expected_surface_output_path is not None and expected_surface_output_path.exists():
                expected_sha = sha256_file(expected_surface_output_path).lower()
                if len(transaction_final_sha) != 64 or transaction_final_sha.lower() != expected_sha:
                    issues.append("thesis mutation transaction final docx sha256 must match the rendered final DOCX")

    humanizer_route = parse_line_value(find_lines_with_prefix(record_lines, "- humanizer route decision:")[0])
    humanizer_lang = parse_line_value(find_lines_with_prefix(record_lines, "- humanizer target language:")[0])
    humanizer_scope = parse_line_value(find_lines_with_prefix(record_lines, "- humanizer scope:")[0])
    humanizer_evidence_seen_skills: set[str] = set()
    humanizer_evidence_seen_languages: set[str] = set()

    for prefix in PATH_FIELD_PREFIXES:
        lines = find_lines_with_prefix(record_lines, prefix)
        if len(lines) != 1:
            continue
        value = parse_line_value(lines[0])
        raw_value = raw_line_value(lines[0])
        if prefix in REQUIRED_NON_NONE_PATHS_BY_MODE.get(task_mode, set()):
            if prefix == "- bibliography audit evidence path:" and not paper_only_required:
                pass
            elif is_explicit_none(value) or not is_explicit(value):
                issues.append(
                    f"gate record field must not be none for mode {task_mode}: {normalize(lines[0])}"
                )
                continue
        if is_explicit_none(value):
            continue
        if not is_explicit(value):
            continue

        for raw_path in split_path_values(raw_value):
            resolved = resolve_record_path(raw_path, record_path)
            if prefix == "- filled format-repair task record path:":
                issues.extend(check_format_repair_task_record(resolved))
            elif prefix == "- citation audit evidence path:":
                issues.extend(check_thesis_citation_audit_report(resolved, expected_docx_path=rendered_docx_path))
            elif prefix == "- reference rendered label geometry path:":
                try:
                    label_report = json.loads(resolved.read_text(encoding="utf-8"))
                except Exception as exc:
                    issues.append(f"reference rendered label geometry path must be readable JSON: {resolved} ({exc})")
                else:
                    if label_report.get("schema") != "graduation-project-builder.pdf-bibliography-label-audit.v1":
                        issues.append(f"reference rendered label geometry path has wrong schema: {resolved}")
                    if label_report.get("result") != "pass":
                        issues.append("reference rendered label geometry result must be pass")
                    expected_family = str(label_report.get("expected_label_family", "")).lower()
                    chosen_family = str(label_report.get("chosen_label_family", "")).lower()
                    if expected_family not in {"bracket", "dot"}:
                        issues.append("reference rendered label geometry expected_label_family must be bracket or dot")
                    if chosen_family != expected_family:
                        issues.append("reference rendered label geometry chosen_label_family must match expected_label_family")
                    min_count = int(label_report.get("min_count") or 60)
                    selected_count_key = (
                        "bracket_label_count" if expected_family == "bracket" else "dot_label_count"
                    )
                    if int(label_report.get(selected_count_key) or 0) < min_count:
                        issues.append(
                            f"reference rendered label geometry must prove at least {min_count} {expected_family} labels"
                        )
                    unexpected_core_count = int(label_report.get("unexpected_core_label_count") or 0)
                    if unexpected_core_count:
                        issues.append("reference rendered label geometry must not show fallback labels inside 1..min_count")
                    if label_report.get("label_geometry_verdict") != "pass":
                        issues.append("reference rendered label geometry verdict must be pass")
            elif prefix == "- bibliography label-family decision path:":
                try:
                    decision_text = resolved.read_text(encoding="utf-8")
                except Exception as exc:
                    issues.append(f"bibliography label-family decision path must be readable: {resolved} ({exc})")
                else:
                    lowered_decision = decision_text.lower()
                    locks_bracket = "chosen visible label family: bracket" in lowered_decision or "chosen label family: bracket" in lowered_decision
                    locks_dot = "chosen visible label family: dot" in lowered_decision or "chosen label family: dot" in lowered_decision
                    if not locks_bracket and not locks_dot:
                        issues.append("bibliography label-family decision must explicitly lock chosen visible label family: bracket or dot")
                    if locks_bracket and locks_dot:
                        issues.append("bibliography label-family decision must not lock both bracket and dot")
                    if locks_dot and "[1]" in decision_text and "body citation" not in lowered_decision:
                        issues.append("bibliography label-family decision choosing dot must distinguish bibliography labels from body citation brackets")
            elif prefix == "- formula rendered label geometry path:":
                try:
                    formula_render_report = json.loads(resolved.read_text(encoding="utf-8"))
                except Exception as exc:
                    issues.append(f"formula rendered label geometry path must be readable JSON: {resolved} ({exc})")
                else:
                    if int(formula_render_report.get("rendered_formula_label_issue_count") or 0) != 0:
                        issues.append("formula rendered label geometry report must have rendered_formula_label_issue_count=0")
                    if int(formula_render_report.get("rendered_formula_label_split_pair_count") or 0) != 0:
                        issues.append("formula rendered label geometry report must have rendered_formula_label_split_pair_count=0")
                    sizes = formula_render_report.get("rendered_formula_number_unique_sizes") or []
                    if isinstance(sizes, list) and len(sizes) > 1:
                        issues.append("formula rendered label geometry report must show one visible font-size family")
            elif prefix == "- body style audit evidence path:":
                issues.extend(check_docx_body_style_audit_report(resolved, expected_final_docx_path=rendered_docx_path))
            elif prefix == "- docx font/encoding audit evidence path:":
                issues.extend(check_docx_font_audit_report(resolved, expected_docx_path=rendered_docx_path))
            elif prefix == "- humanizer evidence path:":
                evidence_issues, seen_skills, seen_languages = check_humanizer_evidence_record(resolved)
                issues.extend(evidence_issues)
                humanizer_evidence_seen_skills.update(seen_skills)
                humanizer_evidence_seen_languages.update(seen_languages)
            elif prefix == "- font-family baseline audit evidence path:":
                issues.extend(
                    check_effective_font_evidence_record(
                        resolved,
                        expected_type=EVIDENCE_FIELD_TO_TYPE[prefix],
                        acceptance_mode=task_mode,
                    )
                )
            elif prefix in EVIDENCE_FIELD_TO_TYPE:
                issues.extend(
                    check_review_evidence_record(
                        resolved,
                        expected_type=EVIDENCE_FIELD_TO_TYPE[prefix],
                        acceptance_mode=task_mode,
                    )
                )
            else:
                issues.extend(validate_existing_path(resolved, require_nonempty_file=False))
                if prefix == "- rendered PDF path:" and resolved.suffix.lower() not in PDF_EXTENSIONS:
                    issues.append(f"gate record rendered PDF path is not a pdf: {resolved}")
                if prefix == "- full rendered page/footer map path:":
                    try:
                        footer_report = json.loads(resolved.read_text(encoding="utf-8"))
                    except Exception as exc:
                        issues.append(f"full rendered page/footer map path must be readable JSON: {resolved} ({exc})")
                    else:
                        if footer_report.get("schema") != "graduation-project-builder.pdf-page-footer-audit.v1":
                            issues.append(f"full rendered page/footer map path has wrong schema: {resolved}")
                        if footer_report.get("result") != "pass":
                            issues.append("full rendered page/footer map result must be pass")
                        if int(footer_report.get("footer_mismatch_count") or 0) != 0:
                            issues.append("full rendered page/footer map must have footer_mismatch_count=0")
                        if rendered_pdf_path and rendered_pdf_path.exists():
                            expected_pdf_sha = sha256_file(rendered_pdf_path)
                            recorded_pdf_sha = str(footer_report.get("pdf_sha256", ""))
                            if recorded_pdf_sha and recorded_pdf_sha.upper() != expected_pdf_sha.upper():
                                issues.append("full rendered page/footer map pdf_sha256 must match rendered PDF")
                if prefix == "- page-image artifact paths:" and resolved.suffix.lower() not in IMAGE_EXTENSIONS:
                    issues.append(f"gate record page-image artifact is not an image: {resolved}")

    bibliography_completeness_verdict = gate_values.get("- bibliography empty-entry/content completeness verdict:", "")
    if task_mode in THESIS_MODES:
        if not surface_verdict_passes(bibliography_completeness_verdict):
            issues.append("thesis gate record bibliography empty-entry/content completeness verdict must be pass")
        bibliography_completeness_paths = gate_raw_values.get(
            "- bibliography empty-entry/content completeness evidence path:",
            "",
        )
        if is_explicit(bibliography_completeness_verdict) and not is_explicit_none(bibliography_completeness_verdict):
            resolved_completeness_paths = [
                resolve_record_path(raw_path, record_path)
                for raw_path in split_path_values(bibliography_completeness_paths)
            ]
            has_citation_completeness_report = any(
                path.exists() and "- bibliography empty/content-missing entries: 0" in read_optional_text(path)
                for path in resolved_completeness_paths
            )
            has_font_completeness_report = any(
                path.exists()
                and "- bibliography empty-entry/content completeness checks: pass" in read_optional_text(path)
                and "- bibliography empty-entry/content completeness hits: 0" in read_optional_text(path)
                for path in resolved_completeness_paths
            )
            if not has_citation_completeness_report:
                issues.append("bibliography completeness evidence must include a citation audit with zero empty/content-missing entries")
            if not has_font_completeness_report:
                issues.append("bibliography completeness evidence must include a font/encoding audit with zero empty-entry/content hits")

    record_text = "\n".join(record_lines)
    verification_scope_claim = gate_values.get("- verification scope claim:", "")
    sample_self_check_value = gate_values.get("- sample self-check report path:", "")
    page_class_matrix_value = gate_values.get("- page-class coverage matrix evidence path:", "")
    mandatory_surface_inventory_value = gate_values.get("- mandatory thesis surface inventory path:", "")
    front_matter_surface_matrix_value = gate_values.get("- front matter surface coverage matrix path:", "")
    end_matter_surface_matrix_value = gate_values.get("- end matter surface coverage matrix path:", "")
    high_risk_surface_matrix_value = gate_values.get("- high-risk thesis format surface matrix path:", "")
    user_issue_ledger_value = gate_values.get("- user-reported issue ledger path:", "")
    figure_asset_manifest_value = gate_values.get("- figure asset manifest path:", "")
    figure_family_summary_for_contract = gate_values.get("- figure family style summary:", "")
    figure_contract_summary = gate_values.get("- figure source/style contract summary:", "")
    structural_source_bbox_value = gate_values.get("- structural source-scale bbox map paths:", "")
    structural_inserted_geometry_value = gate_values.get("- structural inserted-scale geometry evidence paths:", "")
    structural_dense_crop_value = gate_values.get("- structural dense-zone crop evidence paths:", "")
    inserted_collision_evidence_value = gate_values.get("- inserted-scale collision evidence paths:", "")
    structural_relation_attribute_verdict = gate_values.get("- structural relation-attribute collision verdict:", "")
    structural_shape_overlap_verdict = gate_values.get("- structural shape-overlap verdict:", "")
    structural_inserted_collision_verdict = gate_values.get("- structural inserted-scale collision verdict:", "")
    structural_source_to_inserted_verdict = gate_values.get("- structural source-to-inserted geometry verdict:", "")

    if task_mode in THESIS_MODES:
        sample_report_path = first_resolved_path(
            gate_raw_values.get("- sample self-check report path:", ""), record_path
        ) if is_explicit(sample_self_check_value) and not is_explicit_none(sample_self_check_value) else None
        if sample_report_path is not None:
            issues.extend(
                validate_sample_self_check_report(
                    sample_report_path,
                    record_path,
                    expected_final_docx=expected_surface_output_path,
                    expected_final_sha256=expected_surface_output_sha256,
                )
            )

        if not is_explicit(mandatory_surface_inventory_value) or is_explicit_none(mandatory_surface_inventory_value):
            issues.append("thesis gate record lacks mandatory thesis surface inventory path")
        else:
            inventory_path = first_resolved_path(
                gate_raw_values.get("- mandatory thesis surface inventory path:", ""), record_path
            )
            if inventory_path is not None:
                issues.extend(
                    validate_thesis_surface_inventory(
                        inventory_path,
                        record_path,
                        task_mode,
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        for value, field_name in (
            (front_matter_surface_matrix_value, "front matter surface coverage matrix path"),
            (end_matter_surface_matrix_value, "end matter surface coverage matrix path"),
        ):
            if not is_explicit(value) or is_explicit_none(value):
                issues.append(f"thesis gate record lacks {field_name}")
        if is_explicit(front_matter_surface_matrix_value) and not is_explicit_none(front_matter_surface_matrix_value):
            front_matrix_path = first_resolved_path(
                gate_raw_values.get("- front matter surface coverage matrix path:", ""), record_path
            )
            if front_matrix_path is not None:
                issues.extend(
                    validate_surface_coverage_matrix(
                        front_matrix_path,
                        record_path,
                        matrix_label="front matter",
                        required_aliases=FRONT_MATTER_SURFACE_COVERAGE_ALIASES,
                        acceptance_mode=task_mode,
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        if is_explicit(end_matter_surface_matrix_value) and not is_explicit_none(end_matter_surface_matrix_value):
            end_matrix_path = first_resolved_path(
                gate_raw_values.get("- end matter surface coverage matrix path:", ""), record_path
            )
            if end_matrix_path is not None:
                issues.extend(
                    validate_surface_coverage_matrix(
                        end_matrix_path,
                        record_path,
                        matrix_label="end matter",
                        required_aliases=END_MATTER_SURFACE_COVERAGE_ALIASES,
                        acceptance_mode=task_mode,
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        if not is_explicit(high_risk_surface_matrix_value) or is_explicit_none(high_risk_surface_matrix_value):
            issues.append("thesis gate record lacks high-risk thesis format surface matrix path")
        else:
            high_risk_matrix_path = first_resolved_path(
                gate_raw_values.get("- high-risk thesis format surface matrix path:", ""), record_path
            )
            if high_risk_matrix_path is not None:
                issues.extend(
                    validate_high_risk_surface_matrix(
                        high_risk_matrix_path,
                        record_path,
                        task_mode,
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        for prefix in MANDATORY_THESIS_VERDICT_PREFIXES:
            verdict_value = gate_values.get(prefix, "")
            allow_optional = prefix == "- appendix format verdict:"
            if not surface_verdict_passes(verdict_value, allow_not_applicable_with_reason=allow_optional):
                issues.append(f"thesis gate record {prefix} must be pass or an appendix not-applicable-with-reason verdict")
        high_risk_verdict = gate_values.get("- high-risk thesis format surface verdict:", "")
        if not surface_verdict_passes(high_risk_verdict):
            issues.append("thesis gate record high-risk thesis format surface verdict must be pass")
        keyword_title_contamination_verdict = gate_values.get(
            "- keyword content title-style contamination verdict:",
            "",
        )
        if not surface_verdict_passes(keyword_title_contamination_verdict):
            issues.append("thesis gate record keyword content title-style contamination verdict must be pass")

        if gate_values.get("- forbidden substitute evidence used?:", "") != "no":
            issues.append("thesis gate record must explicitly state forbidden substitute evidence used?: no")
        protected_surface_evidence_usage: dict[str, list[str]] = {}
        for surface_id, (path_prefix, verdict_prefix) in FRONT_MATTER_FINAL_EVIDENCE_FIELDS.items():
            verdict_value = gate_values.get(verdict_prefix, "")
            if not surface_verdict_passes(verdict_value):
                issues.append(f"thesis gate record {verdict_prefix} must be pass")
            path_value = gate_values.get(path_prefix, "")
            if not is_explicit(path_value) or is_explicit_none(path_value):
                issues.append(f"thesis gate record {path_prefix} must name surface-level review evidence")
                continue
            for raw_path in split_path_values(gate_raw_values.get(path_prefix, "")):
                resolved = resolve_record_path(raw_path, record_path)
                protected_surface_evidence_usage.setdefault(str(resolved), []).append(surface_id)
                issues.extend(
                    check_required_surface_evidence_record(
                        resolved,
                        expected_type="thesis-rendered-page",
                        acceptance_mode=task_mode,
                        required_surface_id=surface_id,
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        for resolved_path, surface_ids in protected_surface_evidence_usage.items():
            unique_surface_ids = sorted(set(surface_ids))
            if len(unique_surface_ids) > 1:
                issues.append(
                    "protected thesis surface evidence record must prove a single protected surface only: "
                    f"{resolved_path} used for {', '.join(unique_surface_ids)}"
                )
        toc_geometry_path_value = gate_values.get("- toc visual geometry evidence paths:", "")
        if not is_explicit(toc_geometry_path_value) or is_explicit_none(toc_geometry_path_value):
            issues.append("thesis gate record must name TOC visual geometry evidence paths")
        else:
            for raw_path in split_path_values(gate_raw_values.get("- toc visual geometry evidence paths:", "")):
                resolved = resolve_record_path(raw_path, record_path)
                issues.extend(
                    check_required_surface_evidence_record(
                        resolved,
                        expected_type="thesis-rendered-page",
                        acceptance_mode=task_mode,
                        required_surface_id="toc_entries",
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        toc_paragraph_path_value = gate_values.get("- toc paragraph-and-typography evidence paths:", "")
        if not is_explicit(toc_paragraph_path_value) or is_explicit_none(toc_paragraph_path_value):
            issues.append("thesis gate record must name TOC paragraph-and-typography evidence paths")
        else:
            for raw_path in split_path_values(gate_raw_values.get("- toc paragraph-and-typography evidence paths:", "")):
                resolved = resolve_record_path(raw_path, record_path)
                issues.extend(
                    check_required_surface_evidence_record(
                        resolved,
                        expected_type="thesis-rendered-page",
                        acceptance_mode=task_mode,
                        required_surface_id="toc_entries",
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )
        toc_visible_run_path_value = gate_values.get("- toc visible-run typography evidence paths:", "")
        if not is_explicit(toc_visible_run_path_value) or is_explicit_none(toc_visible_run_path_value):
            issues.append("thesis gate record must name TOC visible-run typography evidence paths")
        else:
            for raw_path in split_path_values(gate_raw_values.get("- toc visible-run typography evidence paths:", "")):
                resolved = resolve_record_path(raw_path, record_path)
                issues.extend(
                    check_required_surface_evidence_record(
                        resolved,
                        expected_type="thesis-rendered-page",
                        acceptance_mode=task_mode,
                        required_surface_id="toc_entries",
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )

        whole_pagination_verdict = gate_values.get("- whole-document pagination verdict:", "")
        if not surface_verdict_passes(whole_pagination_verdict):
            issues.append("thesis gate record whole-document pagination verdict must be pass")
        whole_pagination_path_value = gate_values.get("- whole-document pagination evidence path:", "")
        if not is_explicit(whole_pagination_path_value) or is_explicit_none(whole_pagination_path_value):
            issues.append("thesis gate record must name whole-document pagination evidence path")
        else:
            for raw_path in split_path_values(gate_raw_values.get("- whole-document pagination evidence path:", "")):
                resolved = resolve_record_path(raw_path, record_path)
                issues.extend(
                    check_required_surface_evidence_record(
                        resolved,
                        expected_type="thesis-rendered-page",
                        acceptance_mode=task_mode,
                        required_surface_id="whole_document_pagination",
                        expected_reviewed_output=expected_surface_output_path,
                        expected_reviewed_sha256=expected_surface_output_sha256,
                    )
                )

        exact_output_line = find_lines_with_prefix(record_lines, "- exact output paths:")
        if exact_output_line:
            for raw_path in split_path_values(raw_line_value(exact_output_line[0])):
                candidate = resolve_record_path(raw_path, record_path)
                if candidate.suffix.lower() not in {".md", ".txt", ".json", ".log"} or not candidate.exists():
                    continue
                candidate_text = read_optional_text(candidate).lower()
                for token in SAMPLE_SELF_CHECK_BLOCK_TOKENS:
                    if token in candidate_text:
                        issues.append(
                            f"gate record exact output includes delivery-blocked self-check text: {candidate} ({token})"
                        )

        format_task_context = ""
        format_task_line_for_context = find_lines_with_prefix(record_lines, "- filled format-repair task record path:")
        if format_task_line_for_context:
            format_task_value = parse_line_value(format_task_line_for_context[0])
            if is_explicit(format_task_value) and not is_explicit_none(format_task_value):
                format_task_path_for_context = resolve_record_path(
                    raw_line_value(format_task_line_for_context[0]), record_path
                )
                format_task_text_for_context = read_optional_text(format_task_path_for_context)
                context_prefixes = (
                    "- explicit user rule:",
                    "- touched blocks this round:",
                    "- specific format classes to verify:",
                    "- rendered pages to inspect:",
                )
                format_task_context = "\n".join(
                    line for line in format_task_text_for_context.splitlines()
                    if any(line.lower().startswith(prefix.lower()) for prefix in context_prefixes)
                )

        subtask_context = raw_line_value(find_lines_with_prefix(record_lines, "- subtask:")[0])
        full_scope_context = "\n".join(
            [
                verification_scope_claim,
                explicit_override_text,
                subtask_context,
                gate_raw_values.get("- thesis rendered-page verification summary:", ""),
                gate_raw_values.get("- machine-vision summary by touched surface:", ""),
                format_task_context,
            ]
        ).lower()
        full_scope_required = verification_scope_claim in FULL_SCOPE_CLAIMS or contains_any(
            full_scope_context, FULL_SCOPE_TOKENS
        )
        formula_context_parts = [explicit_override_text, subtask_context, format_task_context]
        if is_explicit(user_issue_ledger_value) and not is_explicit_none(user_issue_ledger_value):
            for raw_path in split_path_values(gate_raw_values.get("- user-reported issue ledger path:", "")):
                ledger_path = resolve_record_path(raw_path, record_path)
                formula_context_parts.append(read_optional_text(ledger_path))
        for prefix in (
            "- formula source coverage audit path:",
            "- formula source map paths:",
            "- formula source DOCX paths:",
            "- formula source count:",
            "- formula source matched count:",
            "- formula source missing count:",
            "- formula source coverage ratio:",
            "- formula object preservation summary:",
            "- formula numbering surface summary:",
        ):
            value = gate_values.get(prefix, "")
            if is_explicit(value) and not is_explicit_none(value):
                formula_context_parts.append(value)
        formula_context = "\n".join(formula_context_parts).lower()
        mechanical_context_parts = list(formula_context_parts)
        for prefix in (
            "- exact final CAD delivery package path:",
            "- mechanical drawing package audit path:",
            "- mechanical drawing package exact package path:",
        ):
            value = gate_values.get(prefix, "")
            if is_explicit(value) and not is_explicit_none(value):
                mechanical_context_parts.append(value)
        mechanical_formula_context = "\n".join(mechanical_context_parts).lower()
        mechanical_formula_required = _mechanical_formula_minimum_required(mechanical_formula_context)
        formula_context_required = contains_any(formula_context, FORMULA_HINT_TOKENS) or mechanical_formula_required
        min_formula_count = MECHANICAL_FORMULA_MIN_COUNT if mechanical_formula_required else 0
        min_body_formula_count = MECHANICAL_FORMULA_MIN_COUNT if mechanical_formula_required else 0
        formula_audit_report: dict[str, object] | None = None
        formula_surface_present = False
        if (
            task_mode in THESIS_MODES
            and rendered_docx_path is not None
            and rendered_docx_path.exists()
            and (zipfile.is_zipfile(rendered_docx_path) or formula_context_required)
        ):
            try:
                formula_audit_report = audit_formula_objects(
                    rendered_docx_path,
                    min_formula_count=min_formula_count if min_formula_count else None,
                    min_body_formula_count=min_body_formula_count if min_body_formula_count else None,
                    rendered_pdf=rendered_pdf_path if rendered_pdf_path and rendered_pdf_path.exists() else None,
                )
            except Exception as exc:
                if formula_context_required:
                    issues.append(f"thesis gate record formula object audit could not read final DOCX: {exc}")
            else:
                formula_math_count = int(formula_audit_report.get("math_object_count") or 0)
                formula_body_math_count = int(formula_audit_report.get("body_math_object_count") or 0)
                formula_like_count = int(formula_audit_report.get("formula_like_paragraph_count") or 0)
                formula_pseudo_count = int(formula_audit_report.get("pseudo_formula_count") or 0)
                formula_surface_present = formula_math_count > 0 or formula_like_count > 0
                if formula_pseudo_count > 0:
                    examples = formula_audit_report.get("pseudo_formula_paragraphs") or []
                    example_text = ""
                    if isinstance(examples, list) and examples:
                        first = examples[0]
                        if isinstance(first, dict):
                            example_text = str(first.get("text", ""))[:120]
                    issues.append(
                        "thesis gate record final DOCX contains plain-text pseudo-formula paragraphs: "
                        f"{formula_pseudo_count}"
                        + (f" (example: {example_text})" if example_text else "")
                    )
                formula_number_layout_issue_count = int(
                    formula_audit_report.get("formula_number_layout_issue_count") or 0
                )
                rendered_formula_label_issue_count = int(
                    formula_audit_report.get("rendered_formula_label_issue_count") or 0
                )
                if formula_number_layout_issue_count > 0:
                    issues.append(
                        "thesis gate record final DOCX formula number cells have layout defects: "
                        f"{formula_number_layout_issue_count}"
                    )
                formula_number_requirement_issue_count = int(
                    formula_audit_report.get("formula_number_requirement_issue_count") or 0
                )
                if formula_number_requirement_issue_count > 0:
                    issues.append(
                        "thesis gate record final DOCX formula numbering requirement failed: "
                        f"{formula_number_requirement_issue_count}"
                    )
                if rendered_formula_label_issue_count > 0:
                    issues.append(
                        "thesis gate record rendered PDF formula labels have split-line or font-size defects: "
                        f"{rendered_formula_label_issue_count}"
                    )
                if formula_context_required and not formula_surface_present:
                    issues.append("thesis gate record formula task has no detectable formula object in the final DOCX")
                if min_formula_count and formula_math_count < min_formula_count:
                    issues.append(
                        "mechanical thesis formula object count is below minimum "
                        f"{min_formula_count}: math_object_count={formula_math_count}"
                    )
                if min_body_formula_count and formula_body_math_count < min_body_formula_count:
                    issues.append(
                        "mechanical thesis body formula object count is below minimum "
                        f"{min_body_formula_count}: body_math_object_count={formula_body_math_count}"
                    )
                formula_object_audit_value = gate_values.get("- formula object audit evidence path:", "")
                if (formula_surface_present or formula_context_required) and (
                    not is_explicit(formula_object_audit_value) or is_explicit_none(formula_object_audit_value)
                ):
                    issues.append("thesis gate record with formula-required context must name formula object audit evidence path")
                elif is_explicit(formula_object_audit_value) and not is_explicit_none(formula_object_audit_value):
                    formula_audit_path = resolve_record_path(
                        gate_raw_values.get("- formula object audit evidence path:", ""),
                        record_path,
                    )
                    try:
                        recorded_formula_audit = json.loads(formula_audit_path.read_text(encoding="utf-8"))
                    except Exception as exc:
                        issues.append(f"formula object audit evidence path must be readable JSON: {formula_audit_path} ({exc})")
                    else:
                        if recorded_formula_audit.get("schema") != "graduation-project-builder.formula-object-audit.v1":
                            issues.append(f"formula object audit evidence has wrong schema: {formula_audit_path}")
                        expected_sha = sha256_file(rendered_docx_path)
                        recorded_sha = str(recorded_formula_audit.get("docx_sha256", ""))
                        if recorded_sha and recorded_sha.upper() != expected_sha.upper():
                            issues.append("formula object audit evidence docx_sha256 must match final DOCX")
                        if int(recorded_formula_audit.get("pseudo_formula_count") or 0) != formula_pseudo_count:
                            issues.append("formula object audit evidence pseudo_formula_count differs from live final DOCX audit")
                        if int(recorded_formula_audit.get("formula_number_layout_issue_count") or 0) != 0:
                            issues.append("formula object audit evidence formula_number_layout_issue_count must be 0")
                        if str(recorded_formula_audit.get("formula_number_requirement_verdict", "")).lower() != "pass":
                            issues.append("formula object audit evidence formula_number_requirement_verdict must be pass")
                        if int(recorded_formula_audit.get("formula_number_requirement_issue_count") or 0) != 0:
                            issues.append("formula object audit evidence formula_number_requirement_issue_count must be 0")
                        if int(recorded_formula_audit.get("body_formula_group_count") or 0) > 0 and int(
                            recorded_formula_audit.get("strict_formula_number_label_count") or 0
                        ) < int(recorded_formula_audit.get("body_formula_group_count") or 0):
                            issues.append("formula object audit evidence strict_formula_number_label_count must cover body_formula_group_count")
                        if int(recorded_formula_audit.get("rendered_formula_label_issue_count") or 0) != 0:
                            issues.append("formula object audit evidence rendered_formula_label_issue_count must be 0")
                        if int(recorded_formula_audit.get("rendered_formula_label_split_pair_count") or 0) != 0:
                            issues.append("formula object audit evidence rendered_formula_label_split_pair_count must be 0")
                        if str(recorded_formula_audit.get("formula_narrative_context_verdict", "")).lower() != "pass":
                            issues.append("formula object audit evidence formula_narrative_context_verdict must be pass")
                        if int(recorded_formula_audit.get("formula_dump_marker_count") or 0) != 0:
                            issues.append("formula object audit evidence formula_dump_marker_count must be 0")
                        if int(recorded_formula_audit.get("formula_without_nearby_body_explanation_count") or 0) != 0:
                            issues.append(
                                "formula object audit evidence formula_without_nearby_body_explanation_count must be 0"
                            )
                        if int(recorded_formula_audit.get("formula_narrative_style_issue_count") or 0) != 0:
                            issues.append("formula object audit evidence formula_narrative_style_issue_count must be 0")
                        if int(recorded_formula_audit.get("orphan_formula_style_issue_count") or 0) != 0:
                            issues.append("formula object audit evidence orphan_formula_style_issue_count must be 0")
                        if min_formula_count:
                            if str(recorded_formula_audit.get("formula_duplicate_density_verdict", "")).lower() != "pass":
                                issues.append("formula object audit evidence formula_duplicate_density_verdict must be pass")
                            if float(recorded_formula_audit.get("unique_formula_ratio") or 0.0) <= 0.0:
                                issues.append("formula object audit evidence unique_formula_ratio must be positive")
                            if int(recorded_formula_audit.get("duplicate_formula_body_text_count") or 0) < 0:
                                issues.append("formula object audit evidence duplicate_formula_body_text_count must be non-negative")
                            density_field = gate_values.get("- formula duplicate-density verdict:", "")
                            unique_ratio_field = gate_values.get("- unique formula ratio:", "")
                            duplicate_count_field = gate_values.get("- duplicate formula body text count:", "")
                            if not _mechanical_cad_passish(density_field):
                                issues.append("formula duplicate-density verdict field must be pass-shaped")
                            if not is_explicit(unique_ratio_field) or is_explicit_none(unique_ratio_field):
                                issues.append("unique formula ratio field must be explicit for high-density mechanical formula audits")
                            if not is_explicit(duplicate_count_field) or is_explicit_none(duplicate_count_field):
                                issues.append("duplicate formula body text count field must be explicit for high-density mechanical formula audits")
                        recorded_sizes = recorded_formula_audit.get("rendered_formula_number_unique_sizes") or []
                        if isinstance(recorded_sizes, list) and len(recorded_sizes) > 1:
                            issues.append("formula object audit evidence rendered formula label font sizes must be consistent")
                        if rendered_pdf_path and rendered_pdf_path.exists():
                            expected_pdf_sha = sha256_file(rendered_pdf_path)
                            recorded_pdf_sha = str(recorded_formula_audit.get("rendered_pdf_sha256", ""))
                            if recorded_pdf_sha and recorded_pdf_sha.upper() != expected_pdf_sha.upper():
                                issues.append("formula object audit evidence rendered_pdf_sha256 must match rendered PDF")
                        if formula_surface_present and recorded_formula_audit.get("result") != "pass":
                            issues.append("formula object audit evidence result must be pass when formula surface is present")
                        if min_formula_count:
                            recorded_math_count = int(recorded_formula_audit.get("math_object_count") or 0)
                            recorded_body_math_count = int(recorded_formula_audit.get("body_math_object_count") or 0)
                            recorded_minimum = int(recorded_formula_audit.get("min_formula_count") or 0)
                            recorded_body_minimum = int(recorded_formula_audit.get("min_body_formula_count") or 0)
                            if recorded_minimum < min_formula_count:
                                issues.append(
                                    "mechanical thesis formula audit evidence must bind "
                                    f"min_formula_count={min_formula_count}"
                                )
                            if recorded_body_minimum < min_body_formula_count:
                                issues.append(
                                    "mechanical thesis formula audit evidence must bind "
                                    f"min_body_formula_count={min_body_formula_count}"
                                )
                            if recorded_math_count < min_formula_count:
                                issues.append(
                                    "mechanical thesis formula audit evidence math_object_count is below "
                                    f"minimum {min_formula_count}"
                                )
                            if recorded_body_math_count < min_body_formula_count:
                                issues.append(
                                    "mechanical thesis formula audit evidence body_math_object_count is below "
                                    f"minimum {min_body_formula_count}"
                                )
                            if recorded_formula_audit.get("result") != "pass":
                                issues.append("mechanical thesis formula audit evidence result must be pass")
                formula_object_summary_value_for_audit = gate_values.get("- formula object preservation summary:", "")
                if (formula_surface_present or formula_context_required) and not formula_object_summary_value_for_audit.startswith("passed"):
                    issues.append("thesis gate record formula object preservation summary must be passed when formula context or formula surface is present")
        if full_scope_required:
            if not is_explicit(page_class_matrix_value) or is_explicit_none(page_class_matrix_value):
                issues.append(
                    "gate record claims whole-thesis/template alignment but lacks page-class coverage matrix evidence path"
                )
            else:
                matrix_path = first_resolved_path(
                    gate_raw_values.get("- page-class coverage matrix evidence path:", ""), record_path
                )
                if matrix_path is not None:
                    issues.extend(validate_page_class_coverage_matrix(matrix_path, record_path, task_mode))

        cross_surface_user_reported = False
        user_reported_trigger_context = "\n".join([explicit_override_text, subtask_context]).lower()
        explicit_user_override_active = is_explicit(explicit_override_text) and not is_explicit_none(
            normalize(explicit_override_text)
        )
        explicit_user_rule_values = [
            raw_line_value(line)
            for line in format_task_context.splitlines()
            if line.lower().strip().startswith("- explicit user rule:")
        ]
        active_format_task_context = active_format_task_user_issue_context(format_task_context)
        user_issue_context = "\n".join(
            [
                explicit_override_text,
                subtask_context,
                "\n".join(explicit_user_rule_values),
                active_format_task_context,
            ]
        ).lower()
        user_reported_trigger_context = "\n".join(
            [
                user_reported_trigger_context,
                "\n".join(explicit_user_rule_values),
            ]
        ).lower()
        explicit_user_rule_active = any(
            is_explicit(value)
            and not is_explicit_none(normalize(value))
            and contains_any(value.lower(), set(EXPLICIT_USER_RULE_PROBLEM_TOKENS))
            for value in explicit_user_rule_values
        )
        user_reported_context = explicit_user_override_active or explicit_user_rule_active or contains_any(
            user_reported_trigger_context,
            set(USER_REPORTED_TRIGGER_TOKENS),
        )
        abstract_issue_required = user_reported_context and contains_any(user_issue_context, ABSTRACT_REPORT_TOKENS)
        abstract_indent_issue_required = user_reported_context and contains_any(
            user_issue_context,
            ABSTRACT_INDENT_REPORT_TOKENS,
        )
        references_issue_required = user_reported_context and contains_any(user_issue_context, REFERENCE_REPORT_TOKENS)
        body_opener_header_title_required = user_reported_context and contains_any(
            user_issue_context,
            set(BODY_OPENER_HEADER_TITLE_TOKENS),
        )
        visual_defect_required = user_reported_context and user_reported_visual_defect_required(user_issue_context)
        required_ledger_surfaces = tuple(
            label
            for label, (trigger_tokens, _ledger_tokens) in SURFACE_LEDGER_REQUIREMENTS.items()
            if user_reported_context and contains_any(user_issue_context, set(trigger_tokens))
        )
        cross_surface_user_reported = "cross-surface-regression" in required_ledger_surfaces
        if abstract_issue_required or references_issue_required or required_ledger_surfaces:
            if not is_explicit(user_issue_ledger_value) or is_explicit_none(user_issue_ledger_value):
                issues.append(
                    "gate record references user-reported thesis surfaces but lacks user-reported issue ledger path"
                )
            else:
                ledger_path = first_resolved_path(
                    gate_raw_values.get("- user-reported issue ledger path:", ""), record_path
                )
                if ledger_path is not None:
                    issues.extend(
                        validate_user_issue_ledger(
                            ledger_path,
                            record_path,
                            abstract_required=abstract_issue_required,
                            references_required=references_issue_required,
                            required_surface_labels=required_ledger_surfaces,
                        )
                    )
        if visual_defect_required:
            for prefix in USER_REPORTED_VISUAL_DEFECT_FIELDS:
                field_value = gate_values.get(prefix, "")
                if not is_explicit(field_value) or is_explicit_none(field_value):
                    issues.append(f"gate record lacks {prefix} for user-reported visual defect closure")
            visual_surfaces_value = gate_values.get("- user-reported visual defect surfaces:", "")
            if not contains_any(visual_surfaces_value, set(USER_REPORTED_VISUAL_SURFACE_TOKENS)):
                issues.append("gate record user-reported visual defect surfaces must name the reported protected visual surface family")
            visual_evidence_value = gate_values.get("- user-reported visual defect render-geometry evidence path:", "")
            if is_explicit(visual_evidence_value) and not is_explicit_none(visual_evidence_value):
                for raw_path in split_path_values(
                    gate_raw_values.get("- user-reported visual defect render-geometry evidence path:", "")
                ):
                    evidence_path = resolve_record_path(raw_path, record_path)
                    issues.extend(validate_user_reported_visual_defect_evidence(evidence_path, record_path))
            for prefix, label in (
                (
                    "- user-reported visual defect template-vs-target binding verdict:",
                    "template-vs-target binding verdict",
                ),
                (
                    "- user-reported visual defect full-page/key-surface binding verdict:",
                    "full-page/key-surface binding verdict",
                ),
            ):
                verdict_value = gate_values.get(prefix, "")
                if not surface_verdict_passes(verdict_value) or contains_any(
                    verdict_value,
                    USER_REPORTED_VISUAL_BAD_TOKENS,
                ):
                    issues.append(f"gate record user-reported visual defect {label} must be pass with full rendered binding")
        content_mutation_context = "\n".join(
            [
                explicit_override_text,
                subtask_context,
                selected_workflow,
                gate_raw_values.get("- thesis mutation transaction subtype:", ""),
                gate_raw_values.get("- thesis mutation transaction target surfaces:", ""),
                gate_raw_values.get("- touched template-owned surface families:", ""),
                gate_raw_values.get("- machine-vision summary by touched surface:", ""),
                gate_raw_values.get("- thesis rendered-page verification summary:", ""),
                format_task_context,
            ]
        )
        if task_mode in {"thesis-only", "program-plus-thesis"} and thesis_content_mutation_visual_gate_required(
            content_mutation_context
        ):
            for prefix in CONTENT_MUTATION_VISUAL_REQUIRED_FIELDS:
                field_value = gate_values.get(prefix, "")
                if not is_explicit(field_value) or is_explicit_none(field_value):
                    issues.append(f"gate record lacks {prefix} for content expansion machine-vision/body contamination closure")
            content_render_path_value = gate_values.get("- content mutation rendered-page review path:", "")
            if is_explicit(content_render_path_value) and not is_explicit_none(content_render_path_value):
                for raw_path in split_path_values(
                    gate_raw_values.get("- content mutation rendered-page review path:", "")
                ):
                    evidence_path = resolve_record_path(raw_path, record_path)
                    issues.extend(validate_content_mutation_rendered_evidence(evidence_path, record_path))
            touched_blast_value = gate_values.get("- touched-page/blast-radius machine-vision evidence paths:", "")
            if is_explicit(touched_blast_value) and not is_explicit_none(touched_blast_value):
                for raw_path in split_path_values(
                    gate_raw_values.get("- touched-page/blast-radius machine-vision evidence paths:", "")
                ):
                    evidence_path = resolve_record_path(raw_path, record_path)
                    issues.extend(validate_content_mutation_rendered_evidence(evidence_path, record_path))
            for prefix, label in (
                ("- content mutation machine-vision verdict:", "content mutation machine-vision verdict"),
                ("- inserted body heading-contamination verdict:", "inserted body heading-contamination verdict"),
                ("- caption/table sibling body contamination verdict:", "caption/table sibling body contamination verdict"),
                ("- format lane post-mutation rendered audit verdict:", "format lane post-mutation rendered audit verdict"),
            ):
                verdict_value = gate_values.get(prefix, "")
                if not surface_verdict_passes(verdict_value) or contains_any(
                    verdict_value,
                    CONTENT_MUTATION_VISUAL_BAD_TOKENS,
                ):
                    issues.append(f"gate record {label} must be pass with exact-output rendered machine-vision evidence")
        if abstract_indent_issue_required:
            for prefix in ENGLISH_ABSTRACT_INDENTATION_FIELDS:
                matching_lines = find_lines_with_prefix(record_lines, prefix)
                if not matching_lines:
                    issues.append(f"gate record missing English abstract indentation field: {prefix}")
                    continue
                value = parse_line_value(matching_lines[0])
                if not is_explicit(value) or value in EXPLICIT_VALUES:
                    issues.append(f"gate record English abstract indentation field must not be empty: {normalize(matching_lines[0])}")
            indent_verdict_lines = find_lines_with_prefix(record_lines, "- English abstract indentation verdict:")
            if indent_verdict_lines and not surface_verdict_passes(parse_line_value(indent_verdict_lines[0])):
                issues.append("gate record English abstract indentation verdict must be pass when English abstract indentation is user-reported")

        if body_opener_header_title_required:
            title_evidence_value = gate_values.get("- body opener/header title consistency evidence path:", "")
            title_verdict_value = gate_values.get("- body opener/header title consistency verdict:", "")
            if not is_explicit(title_evidence_value) or is_explicit_none(title_evidence_value):
                issues.append("gate record lacks body opener/header title consistency evidence path")
            else:
                title_evidence_path = first_resolved_path(
                    gate_raw_values.get("- body opener/header title consistency evidence path:", ""),
                    record_path,
                )
                if title_evidence_path is not None:
                    issues.extend(validate_body_opener_header_title_evidence(title_evidence_path, record_path))
            if not surface_verdict_passes(title_verdict_value) or contains_any(
                title_verdict_value,
                {
                    "failed",
                    "missing",
                    "not checked",
                    "sample only",
                    "sampled-only",
                    "stale",
                    "unresolved",
                    "blocked",
                    "not-applicable",
                },
            ):
                issues.append("gate record body opener/header title consistency verdict must be pass when title/header drift is user-reported")

        user_reported_acceptance_context = "\n".join([explicit_override_text, subtask_context]).lower()
        if user_reported_field_gate_required(
            user_reported_acceptance_context,
            FIGURE_WIDTH_USER_REPORTED_TOKENS,
        ):
            require_gate_path_field(
                issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                prefix="- figure extents audit path:",
                label="figure extents audit path for user-reported screenshot/body-figure width repair",
            )
            for raw_path in split_path_values(gate_raw_values.get("- figure extents audit path:", "")):
                evidence_path = resolve_record_path(raw_path, record_path)
                issues.extend(
                    validate_figure_extents_user_evidence(
                        evidence_path,
                        record_path,
                        final_docx=rendered_docx_path,
                    )
                )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- figure paragraph-margin width verdict:",
                label="figure paragraph-margin width verdict",
            )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- figure visible-content completeness verdict:",
                label="figure visible-content completeness verdict",
            )

        if user_reported_field_gate_required(
            user_reported_acceptance_context,
            ALGORITHM_CODE_USER_REPORTED_TOKENS,
        ):
            require_gate_path_field(
                issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                prefix="- algorithm code visibility evidence path:",
                label="algorithm code visibility evidence path for user-reported missing code",
            )
            for raw_path in split_path_values(gate_raw_values.get("- algorithm code visibility evidence path:", "")):
                evidence_path = resolve_record_path(raw_path, record_path)
                issues.extend(
                    validate_algorithm_code_visibility_evidence(
                        evidence_path,
                        record_path,
                        final_docx=rendered_docx_path,
                    )
                )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- algorithm code visibility verdict:",
                label="algorithm code visibility verdict",
            )

        if user_reported_field_gate_required(
            user_reported_acceptance_context,
            TOC_DIGITS_CACHE_USER_REPORTED_TOKENS,
        ):
            require_gate_path_field(
                issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                prefix="- TOC digits/cache preservation evidence path:",
                label="TOC digits/cache preservation evidence path for user-reported TOC numbering/cache repair",
            )
            for raw_path in split_path_values(gate_raw_values.get("- TOC digits/cache preservation evidence path:", "")):
                evidence_path = resolve_record_path(raw_path, record_path)
                issues.extend(
                    validate_toc_digits_cache_evidence(
                        evidence_path,
                        record_path,
                        final_docx=rendered_docx_path,
                    )
                )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- TOC digits/cache preservation verdict:",
                label="TOC digits/cache preservation verdict",
            )

        material_or_generated_required = (
            contains_any(
                user_reported_acceptance_context,
                {
                    "material.doc",
                    "source material",
                    "material source",
                    "\u7d20\u6750.doc",
                    "\u7d20\u6750",
                },
            )
            or (
                contains_any(
                    user_reported_acceptance_context,
                    {
                        "generated image",
                        "generated substitute",
                        "ai generated",
                        "impostor",
                        "wrapped image",
                        "\u751f\u6210\u56fe",
                        "\u5192\u5145",
                        "\u4f2a draw.io",
                        "\u4e0d\u662fdraw.io",
                        "\u4e0d\u662f draw.io",
                    },
                )
                and contains_any(
                    user_reported_acceptance_context,
                    {
                        "draw.io",
                        "drawio",
                        "figure",
                        "diagram",
                        "flowchart",
                        "\u56fe",
                        "\u6d41\u7a0b\u56fe",
                        "\u753b\u7684\u56fe",
                    },
                )
            )
        )
        material_only_required = contains_any(
            user_reported_acceptance_context,
            FIGURE_MATERIAL_ONLY_USER_REPORTED_TOKENS,
        )
        drawio_impostor_required = (
            contains_any(
                user_reported_acceptance_context,
                {
                    "generated image",
                    "generated substitute",
                    "ai generated",
                    "impostor",
                    "wrapped image",
                    "not draw.io",
                    "not drawio",
                    "not a draw.io",
                    "\u751f\u6210\u56fe",
                    "\u5192\u5145",
                    "\u4f2a draw.io",
                    "\u4e0d\u662fdraw.io",
                    "\u4e0d\u662f draw.io",
                },
            )
            and contains_any(
                user_reported_acceptance_context,
                {
                    "draw.io",
                    "drawio",
                    "flowchart",
                    "diagram",
                    "figure",
                    "\u6d41\u7a0b\u56fe",
                    "\u753b\u7684\u56fe",
                    "\u56fe",
                },
            )
        )
        if material_or_generated_required:
            require_gate_path_field(
                issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                prefix="- figure user material source path:",
                label="figure user material source path for material/generated-substitute complaint",
            )
            material_sha_value = gate_values.get("- figure user material source SHA256:", "")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", material_sha_value or ""):
                issues.append("gate record figure user material source SHA256 must be a 64-hex value")
            require_gate_path_field(
                issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                prefix="- figure material-source inventory path:",
                label="figure material-source inventory path for material/generated-substitute complaint",
            )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- figure material-source binding verdict:",
                label="figure material-source binding verdict",
            )
            require_gate_pass_verdict(
                issues,
                gate_values=gate_values,
                prefix="- figure generated-substitute rejection verdict:",
                label="figure generated-substitute rejection verdict",
            )
            if material_only_required:
                require_gate_pass_verdict(
                    issues,
                    gate_values=gate_values,
                    prefix="- figure material-only reuse verdict:",
                    label="figure material-only reuse verdict",
                )
                supplemental_value = gate_values.get("- figure supplemental material source path:", "")
                supplemental_ledger_value = gate_values.get("- figure supplemental missing-figure ledger path:", "")
                if is_explicit(supplemental_value) and not is_explicit_none(supplemental_value):
                    require_gate_path_field(
                        issues,
                        gate_values=gate_values,
                        gate_raw_values=gate_raw_values,
                        record_path=record_path,
                        prefix="- figure supplemental material source path:",
                        label="figure supplemental material source path for material-only missing figures",
                    )
                    supplemental_sha_value = gate_values.get("- figure supplemental material source SHA256:", "")
                    if not re.fullmatch(r"[0-9a-fA-F]{64}", supplemental_sha_value or ""):
                        issues.append("gate record figure supplemental material source SHA256 must be a 64-hex value when supplemental source is used")
                    require_gate_path_field(
                        issues,
                        gate_values=gate_values,
                        gate_raw_values=gate_raw_values,
                        record_path=record_path,
                        prefix="- figure supplemental missing-figure ledger path:",
                        label="figure supplemental missing-figure ledger path",
                    )
                    require_gate_pass_verdict(
                        issues,
                        gate_values=gate_values,
                        prefix="- figure supplemental missing-source verdict:",
                        label="figure supplemental missing-source verdict",
                    )
                elif is_explicit(supplemental_ledger_value) and not is_explicit_none(supplemental_ledger_value):
                    issues.append("gate record supplemental missing-figure ledger is present but supplemental material source path is missing")
        if drawio_impostor_required and not material_only_required:
            validate_flowchart_acceptance_fields(
                issues=issues,
                gate_values=gate_values,
                gate_raw_values=gate_raw_values,
                record_path=record_path,
                final_docx=rendered_docx_path,
            )

        figure_contract_failed = contains_any(
            figure_contract_summary,
            {"failed", "missing", "generic", "png-only", "raster-only", "no manifest", "no svg", "no draw.io"},
        )
        if material_only_required and figure_contract_failed and "no draw.io" in figure_contract_summary:
            figure_contract_failed = contains_any(
                figure_contract_summary,
                {"failed", "missing", "generic", "png-only", "raster-only", "no manifest", "no svg"},
            )
        if figure_contract_failed:
            issues.append("gate record figure source/style contract summary indicates unresolved figure contract drift")
        final_docx_manifest_issues = (
            final_docx_manifest_requirement_issues(rendered_docx_path)
            if rendered_docx_path is not None and rendered_docx_path.exists()
            else []
        )
        figure_summary_active = (
            figure_family_summary_for_contract.startswith("passed")
            or figure_contract_summary.startswith("passed")
            or contains_any("\n".join([subtask_context, explicit_override_text, format_task_context]), FIGURE_REPORT_TOKENS)
            or docx_has_figure_caption(rendered_docx_path)
            or bool(final_docx_manifest_issues)
        )
        manifest_path = first_resolved_path(
            gate_raw_values.get("- figure asset manifest path:", ""), record_path
        ) if is_explicit(figure_asset_manifest_value) and not is_explicit_none(figure_asset_manifest_value) else None
        if figure_summary_active or manifest_path is not None:
            if manifest_path is None:
                if final_docx_manifest_issues:
                    issues.extend(final_docx_manifest_issues)
                else:
                    issues.append("gate record figure work requires a figure asset manifest path")
            else:
                source_docx_for_figure_contract = first_resolved_path(
                    gate_raw_values.get("- figure source DOCX path:", ""), record_path
                ) if is_explicit(gate_values.get("- figure source DOCX path:", "")) and not is_explicit_none(gate_values.get("- figure source DOCX path:", "")) else None
                if source_docx_for_figure_contract is None:
                    source_docx_for_figure_contract = review_inventory_source_docx_path(
                        source_review_artifact_inventory_path
                    )
                issues.extend(
                    validate_figure_contract_manifest(
                        manifest_path,
                        rendered_docx_path,
                        record_path,
                        source_docx=source_docx_for_figure_contract,
                    )
                )
                if figure_manifest_has_er_diagrams(manifest_path):
                    for label, value in (
                        ("structural source-scale bbox map paths", structural_source_bbox_value),
                        ("structural inserted-scale geometry evidence paths", structural_inserted_geometry_value),
                        ("structural dense-zone crop evidence paths", structural_dense_crop_value),
                        ("inserted-scale collision evidence paths", inserted_collision_evidence_value),
                    ):
                        if not is_explicit(value) or is_explicit_none(value) or value.strip().lower().startswith("not-applicable"):
                            issues.append(f"gate record ER figure work requires {label}")
                    for label, value in (
                        ("structural relation-attribute collision verdict", structural_relation_attribute_verdict),
                        ("structural shape-overlap verdict", structural_shape_overlap_verdict),
                        ("structural inserted-scale collision verdict", structural_inserted_collision_verdict),
                        ("structural source-to-inserted geometry verdict", structural_source_to_inserted_verdict),
                    ):
                        if not surface_verdict_passes(value):
                            issues.append(f"gate record ER figure work requires pass {label}")
            figure_review_value = gate_values.get("- figure review evidence paths:", "")
            if not is_explicit(figure_review_value) or is_explicit_none(figure_review_value):
                issues.append("gate record figure work requires figure review evidence paths")
            if not figure_contract_summary.startswith("passed") and figure_contract_summary not in {"not-applicable", "none"}:
                issues.append("gate record figure source/style contract summary must be passed for figure work")

    subtask_text = " ".join(find_lines_with_prefix(record_lines, "- subtask:")).lower()

    helper_target_lock = parse_line_value(find_lines_with_prefix(record_lines, "- helper-script target path lock:")[0])
    helper_scripts_planned = parse_line_value(find_lines_with_prefix(record_lines, "- helper scripts planned this round:")[0])
    helper_script_provenance_summary = parse_line_value(find_lines_with_prefix(record_lines, "- helper script provenance summary:")[0])
    delegated_canonical_helper_paths = parse_line_value(find_lines_with_prefix(record_lines, "- delegated canonical helper paths:")[0])
    project_local_helper_script_preflight_summary = parse_line_value(
        find_lines_with_prefix(record_lines, "- project-local helper script preflight summary:")[0]
    )
    project_local_helper_preflight_report_path = parse_line_value(
        find_lines_with_prefix(record_lines, "- project-local helper preflight report path:")[0]
    )
    project_local_helper_risk_count = parse_line_value(
        find_lines_with_prefix(record_lines, "- project-local helper risk count:")[0]
    )
    project_local_helper_disposition = parse_line_value(
        find_lines_with_prefix(record_lines, "- project-local helper disposition:")[0]
    )
    canonical_source_restart_required = parse_line_value(
        find_lines_with_prefix(record_lines, "- canonical source restart required?:")[0]
    )
    source_manuscript_genealogy_path = parse_line_value(
        find_lines_with_prefix(record_lines, "- source manuscript genealogy path:")[0]
    )
    source_retention_manifest_path = parse_line_value(
        find_lines_with_prefix(record_lines, "- source retention manifest path:")[0]
    )
    source_retention_ratio = parse_line_value(
        find_lines_with_prefix(record_lines, "- source retention ratio:")[0]
    )
    source_retention_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- source retention verdict:")[0]
    )
    rebuild_class = parse_line_value(find_lines_with_prefix(record_lines, "- rebuild class:")[0])
    clean_source_restart_source_path = parse_line_value(
        find_lines_with_prefix(record_lines, "- clean-source restart source path:")[0]
    )
    contaminated_baseline_disposition = parse_line_value(
        find_lines_with_prefix(record_lines, "- contaminated-baseline disposition:")[0]
    )
    ownership_lock = parse_line_value(find_lines_with_prefix(record_lines, "- protected-surface ownership lock:")[0])
    smoke_summary = parse_line_value(find_lines_with_prefix(record_lines, "- post-script smoke-audit summary:")[0])
    body_style_binding_summary = parse_line_value(find_lines_with_prefix(record_lines, "- body style binding summary:")[0])
    normal_baseline_summary = parse_line_value(find_lines_with_prefix(record_lines, "- Normal baseline preservation summary:")[0])
    body_family_summary = parse_line_value(find_lines_with_prefix(record_lines, "- body paragraph family consistency summary:")[0])
    body_heading_contamination_lines = find_lines_with_prefix(record_lines, "- body heading contamination summary:")
    body_heading_contamination_summary = (
        parse_line_value(body_heading_contamination_lines[0])
        if body_heading_contamination_lines
        else "missing"
    )
    surface_face_parity_summary = parse_line_value(find_lines_with_prefix(record_lines, "- surface-face parity summary:")[0])
    surface_paragraph_typography_summary = parse_line_value(find_lines_with_prefix(record_lines, "- surface paragraph-and-typography summary:")[0])
    surface_paragraph_typography_verdict = parse_line_value(find_lines_with_prefix(record_lines, "- surface paragraph-and-typography verdict:")[0])
    sibling_surface_summary = parse_line_value(find_lines_with_prefix(record_lines, "- sibling-surface audit summary:")[0])
    style_blast_radius_summary = parse_line_value(find_lines_with_prefix(record_lines, "- style-blast-radius escalation summary:")[0])
    cross_surface_regression_verdict = parse_line_value(find_lines_with_prefix(record_lines, "- cross-surface regression verdict:")[0])
    toc_underline_pollution_verdict = parse_line_value(find_lines_with_prefix(record_lines, "- TOC underline pollution verdict:")[0])
    table_style_regression_verdict = parse_line_value(find_lines_with_prefix(record_lines, "- table style regression verdict:")[0])
    font_family_baseline_summary = parse_line_value(find_lines_with_prefix(record_lines, "- font-family baseline summary:")[0])
    builder_default_font_summary = parse_line_value(find_lines_with_prefix(record_lines, "- builder/default font rejection summary:")[0])
    code_title_summary = parse_line_value(find_lines_with_prefix(record_lines, "- code title formatting summary:")[0])
    code_block_summary = parse_line_value(find_lines_with_prefix(record_lines, "- code block formatting summary:")[0])
    abstract_baseline_summary = parse_line_value(find_lines_with_prefix(record_lines, "- abstract baseline preservation summary:")[0])
    heading_baseline_summary = parse_line_value(find_lines_with_prefix(record_lines, "- heading baseline preservation summary:")[0])
    heading_summary = parse_line_value(find_lines_with_prefix(record_lines, "- heading family preservation summary:")[0])
    toc_summary = parse_line_value(find_lines_with_prefix(record_lines, "- TOC / bookmark integrity summary:")[0])
    toc_visible_summary = parse_line_value(find_lines_with_prefix(record_lines, "- TOC visible format summary:")[0])
    toc_restoration_summary = parse_line_value(
        find_lines_with_prefix(record_lines, "- TOC baseline restoration summary:")[0]
    )
    toc_rhythm_summary = parse_line_value(
        find_lines_with_prefix(record_lines, "- TOC visual rhythm summary:")[0]
    )
    toc_geometry_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- TOC visual geometry verdict:")[0]
    )
    toc_paragraph_typography_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- TOC paragraph-and-typography verdict:")[0]
    )
    toc_visible_run_typography_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- TOC visible-run typography verdict:")[0]
    )
    format_preservation_promise_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- format preservation promise verdict:")[0]
    )
    chapter_format_preservation_detector_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- chapter format preservation detector verdict:")[0]
    )
    non_target_format_preservation_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- non-target format preservation verdict:")[0]
    )
    chapter_summary = parse_line_value(find_lines_with_prefix(record_lines, "- chapter-start pagination summary:")[0])
    tail_block_summary = parse_line_value(find_lines_with_prefix(record_lines, "- tail-block pagination summary:")[0])
    whole_document_pagination_summary = parse_line_value(
        find_lines_with_prefix(record_lines, "- whole-document pagination summary:")[0]
    )
    whole_document_pagination_verdict = parse_line_value(
        find_lines_with_prefix(record_lines, "- whole-document pagination verdict:")[0]
    )
    table_summary = parse_line_value(find_lines_with_prefix(record_lines, "- custom-layout result-table preservation summary:")[0])
    table_authority_summary = parse_line_value(find_lines_with_prefix(record_lines, "- table authority summary:")[0])
    table_local_structure_summary = parse_line_value(find_lines_with_prefix(record_lines, "- table-local structure summary:")[0])
    table_title_mode_lines = find_lines_with_prefix(record_lines, "- table title-mode summary:")
    table_title_mode_summary = parse_line_value(table_title_mode_lines[0]) if table_title_mode_lines else ""
    table_header_middle_rule_lines = find_lines_with_prefix(record_lines, "- table header-bottom middle-rule summary:")
    table_header_middle_rule_summary = parse_line_value(table_header_middle_rule_lines[0]) if table_header_middle_rule_lines else ""
    figure_family_summary = parse_line_value(find_lines_with_prefix(record_lines, "- figure family style summary:")[0])
    table_family_summary = parse_line_value(find_lines_with_prefix(record_lines, "- table rendered-family summary:")[0])
    wps_preset_summary = parse_line_value(find_lines_with_prefix(record_lines, "- WPS preset application summary:")[0])
    image_rendered_summary = parse_line_value(find_lines_with_prefix(record_lines, "- image replacement rendered summary:")[0])
    image_metadata_summary = parse_line_value(find_lines_with_prefix(record_lines, "- image metadata sync summary:")[0])
    formula_object_summary = parse_line_value(find_lines_with_prefix(record_lines, "- formula object preservation summary:")[0])
    formula_numbering_summary = parse_line_value(find_lines_with_prefix(record_lines, "- formula numbering surface summary:")[0])
    header_placement_summary = parse_line_value(find_lines_with_prefix(record_lines, "- header placement summary:")[0])
    footer_indent_summary = parse_line_value(find_lines_with_prefix(record_lines, "- footer indent summary:")[0])
    footer_baseline_typography_summary = parse_line_value(find_lines_with_prefix(record_lines, "- footer baseline typography summary:")[0])
    bibliography_baseline_summary = parse_line_value(find_lines_with_prefix(record_lines, "- bibliography baseline summary:")[0])
    bibliography_numbering_summary = parse_line_value(find_lines_with_prefix(record_lines, "- bibliography numbering summary:")[0])
    bibliography_comment_aware_repair_summary = parse_line_value(find_lines_with_prefix(record_lines, "- bibliography comment-aware repair summary:")[0])
    project_local_helper_script_risk_summary = parse_line_value(find_lines_with_prefix(record_lines, "- project-local helper script risk summary:")[0])

    toc_touched_in_gate = False
    table_touched_in_gate = False
    style_blast_radius_touched_in_gate = False
    tail_block_touched_in_gate = False
    format_task_line = find_lines_with_prefix(record_lines, "- filled format-repair task record path:")
    if format_task_line:
        format_task_value = parse_line_value(format_task_line[0])
        if is_explicit(format_task_value) and format_task_value not in EXPLICIT_VALUES:
            format_task_path = resolve_record_path(raw_line_value(format_task_line[0]), record_path)
            toc_touched_in_gate = format_repair_task_touches_surface(format_task_path, "toc")
            table_touched_in_gate = format_repair_task_touches_surface(format_task_path, "table")
            style_blast_radius_touched_in_gate = format_repair_task_touches_surface(format_task_path, "style_blast_radius")
            tail_block_touched_in_gate = format_repair_task_touches_surface(format_task_path, "tail_block")
            if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}:
                try:
                    format_task_lines = read_lines(format_task_path)
                except Exception:
                    format_task_lines = []
                for prefix in (
                    "- project template discovery root:",
                    "- discovered candidate template paths:",
                    "- candidate template selection reason:",
                    "- active template source type:",
                    "- active template path lock:",
                    "- active template fingerprint:",
                    "- active template profile path:",
                    "- active template selected before mutation?:",
                    "- template alignment verdict:",
                    "- mandatory thesis surface inventory path:",
                    "- front matter surface coverage matrix path:",
                    "- end matter surface coverage matrix path:",
                    "- high-risk thesis format surface matrix path:",
                    "- project-local helper preflight report path:",
                    "- project-local helper risk count:",
                    "- project-local helper disposition:",
                    "- canonical source restart required?:",
                    "- source manuscript genealogy path:",
                    "- source retention manifest path:",
                    "- source retention ratio:",
                    "- source retention verdict:",
                    "- rebuild class:",
                    "- clean-source restart source path:",
                ):
                    gate_value = normalize(gate_raw_values.get(prefix, ""))
                    task_matches = find_lines_with_prefix(format_task_lines, prefix)
                    task_value = normalize(raw_line_value(task_matches[0])) if task_matches else ""
                    if gate_value and task_value and gate_value != task_value:
                        issues.append(
                            f"gate record template lock field differs from filled format-repair task record: {prefix}"
                        )
            format_task_values: dict[str, str] = {}
            format_task_raw_values: dict[str, str] = {}
            for prefix in AGENT_PROTECTED_FIELD_TO_GATE_PREFIX:
                task_matches = find_lines_with_prefix(format_task_lines, prefix)
                if not task_matches:
                    continue
                if len(task_matches) != 1:
                    issues.append(
                        f"filled format-repair task record must contain exactly one protected hard field: {prefix}"
                    )
                    continue
                format_task_values[prefix] = parse_line_value(task_matches[0])
                format_task_raw_values[prefix] = raw_line_value(task_matches[0])
            for format_task_prefix in format_task_values:
                compare_agent_record_field_to_gate(
                    issues,
                    kind=f"filled format-repair task record {format_task_path}",
                    record_path_for_values=format_task_path,
                    record_values=format_task_values,
                    record_raw_values=format_task_raw_values,
                    record_prefix=format_task_prefix,
                    gate_values=gate_values,
                    gate_raw_values=gate_raw_values,
                    gate_record_path=record_path,
                )

    if task_mode in THESIS_MODES and not table_touched_in_gate and docx_has_body_table_or_caption(final_docx_path):
        table_touched_in_gate = True

    if helper_target_lock not in {"locked", "not-applicable", "n/a"}:
        issues.append("gate record helper-script target path lock must be locked or not-applicable")
    helper_scripts_planned_active = is_explicit(helper_scripts_planned) and not is_explicit_none(helper_scripts_planned)
    if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"} and helper_scripts_planned_active:
        if helper_script_provenance_summary not in {"canonical-skill-bundle", "project-local-thin-wrapper"}:
            issues.append("gate record helper script provenance summary must be canonical-skill-bundle or project-local-thin-wrapper for thesis helper activity")
        if not is_explicit(delegated_canonical_helper_paths) or is_explicit_none(delegated_canonical_helper_paths):
            issues.append("gate record must declare delegated canonical helper paths when helper scripts are planned for thesis work")
        else:
            delegated_lines = find_lines_with_prefix(record_lines, "- delegated canonical helper paths:")
            for raw_path in split_path_values(raw_line_value(delegated_lines[0])):
                resolved = resolve_record_path(raw_path, record_path)
                try:
                    resolved.relative_to(SKILL_ROOT)
                except ValueError:
                    issues.append(f"delegated canonical helper path escapes canonical skill bundle: {resolved}")
    if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}:
        if not is_explicit(project_local_helper_script_preflight_summary) or project_local_helper_script_preflight_summary in EXPLICIT_VALUES:
            issues.append("thesis-related gate record must explicitly record a project-local helper script preflight summary")
        if not is_explicit(project_local_helper_preflight_report_path) or is_explicit_none(project_local_helper_preflight_report_path):
            issues.append("thesis-related gate record must explicitly record a project-local helper preflight report path")
        if not re.fullmatch(r"\d+", project_local_helper_risk_count or ""):
            issues.append("thesis-related gate record must explicitly record a numeric project-local helper risk count")
        if not is_explicit(project_local_helper_disposition) or project_local_helper_disposition in EXPLICIT_VALUES:
            issues.append("thesis-related gate record must explicitly record a project-local helper disposition")
        if not is_explicit(canonical_source_restart_required) or (
            canonical_source_restart_required in EXPLICIT_VALUES and canonical_source_restart_required != "no"
        ):
            issues.append("thesis-related gate record must explicitly record canonical source restart required status")
        for label, value in (
            ("source manuscript genealogy path", source_manuscript_genealogy_path),
            ("source retention manifest path", source_retention_manifest_path),
            ("source retention verdict", source_retention_verdict),
            ("rebuild class", rebuild_class),
        ):
            if not is_explicit(value) or is_explicit_none(value):
                issues.append(f"thesis-related gate record must explicitly record {label}")
        if source_retention_ratio == "not-applicable-with-reason":
            pass
        elif is_explicit(source_retention_ratio) and not is_explicit_none(source_retention_ratio):
            try:
                ratio_value = float(source_retention_ratio)
            except ValueError:
                issues.append("thesis-related gate record source retention ratio must be numeric or not-applicable-with-reason")
            else:
                if not 0.0 <= ratio_value <= 1.0:
                    issues.append("thesis-related gate record source retention ratio must be between 0 and 1")
        elif "new-thesis-production" not in rebuild_class and "format-repair" not in rebuild_class:
            issues.append("thesis-related gate record source retention ratio cannot be none outside new-thesis or format-repair workflows")
        if contains_any(source_retention_verdict, {"failed", "fail", "lost", "missing", "blocked"}):
            issues.append("thesis-related gate record source retention verdict is blocking")
        if canonical_source_restart_required in {"yes", "required", "clean-source-restart-required"}:
            if not is_explicit(clean_source_restart_source_path) or is_explicit_none(clean_source_restart_source_path):
                issues.append("clean-source restart required but no clean-source restart source path is recorded")
        if not is_explicit(contaminated_baseline_disposition) or contaminated_baseline_disposition in EXPLICIT_VALUES:
            issues.append("thesis-related gate record must explicitly record a contaminated-baseline disposition")
        issues.extend(
            validate_project_local_helper_preflight_fields(
                record_path=record_path,
                values=gate_values,
                raw_values=gate_raw_values,
                record_kind="gate record",
                task_mode=task_mode,
                risk_summary=project_local_helper_script_risk_summary,
            )
        )
    if ownership_lock not in {"locked", "not-applicable", "n/a"}:
        issues.append("gate record protected-surface ownership lock must be locked or not-applicable")
    if smoke_summary in {"missing", "skipped", "failed", "fail"}:
        issues.append("gate record post-script smoke-audit summary indicates an unresolved failure")
    if contains_any(
        body_style_binding_summary,
        {"failed", "missing pstyle", "implicit fallback", "style-less", "styleless", "no explicit binding", "default fallback"},
    ):
        issues.append("gate record body style binding summary indicates unresolved body-style binding drift")
    if contains_any(
        normal_baseline_summary,
        {"failed", "baseline drift", "normal drift", "wrong font", "wrong size", "wrong alignment", "mismatch"},
    ):
        issues.append("gate record Normal baseline preservation summary indicates unresolved default-body-style drift")
    if contains_any(
        body_family_summary,
        {"failed", "fragmented", "mixed body", "mixed family", "fallback_only", "style-less", "styleless"},
    ):
        issues.append("gate record body paragraph family consistency summary indicates unresolved body-family fragmentation")
    if (
        task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}
        and not body_heading_contamination_summary.startswith("passed")
        and contains_any(
        body_heading_contamination_summary,
        {"failed", "heading-style", "heading style", "outline-level", "contamination", "missing"},
        )
    ):
        issues.append("gate record body heading contamination summary indicates body prose may be styled as a heading")
    if not surface_face_parity_summary.startswith("passed") and contains_any(
        surface_face_parity_summary,
        {
            "failed",
            "missing",
            "not checked",
            "body/normal",
            "normal style",
            "body style",
            "generic body",
            "generic normal",
            "wrong class",
            "class-binding failure",
            "style binding drift",
            "reference entry bound to body",
            "reference title bound to body",
            "caption bound to body",
            "keyword bound to body",
        },
    ):
        issues.append("gate record surface-face parity summary indicates unresolved class-specific style binding drift")
    if not surface_paragraph_typography_summary.startswith("passed") and contains_any(
        surface_paragraph_typography_summary,
        {
            "failed",
            "missing",
            "not checked",
            "style name only",
            "broad baseline metrics",
            "generic metrics",
            "paragraph dialog missing",
            "paragraph-dialog missing",
            "font chain only",
            "geometry only",
            "default application",
            "compression",
            "compressed",
            "scaled",
        },
    ):
        issues.append("gate record surface paragraph-and-typography summary indicates unresolved WPS/Word paragraph-dialog or typography drift")
    if not surface_paragraph_typography_verdict.startswith(("pass", "passed")):
        issues.append("gate record surface paragraph-and-typography verdict must be pass")
    acknowledgement_body_guard_fields = (
        ("- acknowledgement body paragraph-dialog metrics verdict:", "acknowledgement body paragraph-dialog metrics verdict"),
        ("- acknowledgement body direct-run typography verdict:", "acknowledgement body direct-run typography verdict"),
        ("- acknowledgement body first-line indent verdict:", "acknowledgement body first-line indent verdict"),
        ("- acknowledgement body title-contamination verdict:", "acknowledgement body title-contamination verdict"),
    )
    for prefix, label in acknowledgement_body_guard_fields:
        value = gate_values.get(prefix, "")
        if not surface_verdict_passes(value) or contains_any(
            value,
            {
                "failed",
                "missing",
                "not checked",
                "not-checked",
                "sample only",
                "sampled-only",
                "stale",
                "generic",
                "title contamination",
                "title-contamination",
                "heading-like",
            },
        ):
            issues.append(f"gate record {label} indicates unresolved acknowledgement body format drift")
    if not sibling_surface_summary.startswith("passed") and contains_any(
        sibling_surface_summary,
        {"failed", "missing", "not checked", "only sampled", "partial", "sibling drift", "same family not checked"},
    ):
        issues.append("gate record sibling-surface audit summary indicates incomplete same-family surface review")
    style_blast_required = task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"} and (
        style_blast_radius_touched_in_gate or cross_surface_user_reported
    )
    if task_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"}:
        cross_surface_path_value = gate_values.get("- cross-surface regression freeze evidence path:", "")
        if not is_explicit(cross_surface_path_value) or is_explicit_none(cross_surface_path_value):
            issues.append("thesis gate record must name cross-surface regression freeze evidence path")
        if not surface_verdict_passes(cross_surface_regression_verdict) or contains_any(
            cross_surface_regression_verdict,
            {"failed", "missing", "not checked", "sample only", "sampled-only", "stale", "unreviewed", "not-applicable"},
        ):
            issues.append("gate record cross-surface regression verdict indicates protected-surface regression was not closed")
        if not surface_verdict_passes(toc_underline_pollution_verdict) or contains_any(
            toc_underline_pollution_verdict,
            {"failed", "missing", "not checked", "stale", "not-applicable", "underlined remains", "hyperlink style remains"},
        ):
            issues.append("gate record TOC underline pollution verdict indicates unresolved TOC underline or hyperlink-style contamination")
        if not surface_verdict_passes(table_style_regression_verdict) or contains_any(
            table_style_regression_verdict,
            {"failed", "missing", "not checked", "stale", "not-applicable", "lost", "regressed", "wrong family"},
        ):
            issues.append("gate record table style regression verdict indicates unresolved table-style regression")
    if style_blast_required:
        if not is_explicit(style_blast_radius_summary) or contains_any(
            style_blast_radius_summary,
            {"none", "not-applicable", "missing", "not checked", "sample only", "sampled-only", "no escalation"},
        ):
            issues.append("style-blast-radius repair must record explicit cross-surface escalation summary")
        for prefix in (
            "- toc visible-run typography evidence paths:",
            "- table authority evidence paths:",
            "- table-local structure evidence paths:",
            "- table rendered baseline comparison evidence paths:",
            "- cross-surface regression freeze evidence path:",
        ):
            value = gate_values.get(prefix, "")
            if not is_explicit(value) or is_explicit_none(value):
                issues.append(f"style-blast-radius repair must carry non-empty sibling evidence: {prefix}")
        if contains_any(table_authority_summary, {"not-applicable", "none", "missing", "failed", "drift"}):
            issues.append("style-blast-radius repair must close table authority instead of marking it not-applicable")
        if contains_any(table_local_structure_summary, {"not-applicable", "none", "missing", "failed", "drift"}):
            issues.append("style-blast-radius repair must close table-local structure instead of marking it not-applicable")
    if contains_any(
        font_family_baseline_summary,
        {
            "failed",
            "missing",
            "not checked",
            "wrong font",
            "theme alias",
            "minorEastAsia",
            "minorHAnsi",
            "builder",
            "no effective font chain",
            "generic evidence",
            "style name only",
            "direct rpr only",
            "visible text only",
        },
    ):
        issues.append("gate record font-family baseline summary indicates unresolved font baseline drift")
    builder_default_font_summary_lower = builder_default_font_summary.lower()
    builder_default_negated_safe = (
        "no builder/default/guessed" in builder_default_font_summary_lower
        or "no builder/default" in builder_default_font_summary_lower
        or (
            "no builder" in builder_default_font_summary_lower
            and "no default" in builder_default_font_summary_lower
            and "no guessed" in builder_default_font_summary_lower
        )
    )
    if contains_any(
        builder_default_font_summary,
        {
            "failed",
            "allowed",
            "used",
            "guessed",
            "builder-chosen",
            "default font",
            "custom font",
            "no effective font chain",
            "style name only",
        },
    ) and not builder_default_negated_safe:
        issues.append("gate record builder/default font rejection summary indicates unsafe builder/default font usage")
    if contains_any(
        code_title_summary,
        {"failed", "wrong font", "wrong size", "wrong spacing", "wrong line", "body-like", "not restored"},
    ):
        issues.append("gate record code title formatting summary indicates unresolved code-title formatting drift")
    if contains_any(
        code_block_summary,
        {"failed", "wrong font", "wrong size", "wrong spacing", "wrong line", "body-like", "non-monospace", "not restored"},
    ):
        issues.append("gate record code block formatting summary indicates unresolved code-block formatting drift")
    if contains_any(
        abstract_baseline_summary,
        {"failed", "drift", "wrong font", "wrong size", "wrong style", "wrong label", "body-like", "not restored"},
    ):
        issues.append("gate record abstract baseline preservation summary indicates unresolved abstract-surface drift")
    if contains_any(
        heading_baseline_summary,
        {"failed", "drift", "wrong font", "wrong size", "wrong spacing", "wrong line", "residue", "not restored"},
    ):
        issues.append("gate record heading baseline preservation summary indicates unresolved heading-baseline drift")
    if contains_any(heading_summary, {"normal", "lost", "dropped", "wrong style", "failed", "drift"}):
        issues.append("gate record heading family preservation summary indicates unresolved heading-style drift")
    if contains_any(toc_summary, {"stale", "drift", "broken", "failed", "bookmark error", "undefined bookmark"}):
        issues.append("gate record TOC / bookmark integrity summary indicates unresolved TOC or bookmark drift")
    if contains_any(toc_visible_summary, {"failed", "drift", "wrong", "mismatch", "unstyled", "default", "broken"}):
        issues.append("gate record TOC visible format summary indicates unresolved visible TOC formatting drift")
    if toc_touched_in_gate:
        for prefix in (
            "- toc restoration evidence paths:",
            "- toc rendered baseline comparison evidence paths:",
            "- toc visual geometry evidence paths:",
            "- toc paragraph-and-typography evidence paths:",
            "- toc visible-run typography evidence paths:",
        ):
            line = find_lines_with_prefix(record_lines, prefix)[0]
            value = parse_line_value(line)
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record field must not be empty when TOC is touched: {normalize(line)}")
        for value, label in (
            (toc_restoration_summary, "TOC baseline restoration summary"),
            (toc_rhythm_summary, "TOC visual rhythm summary"),
            (toc_geometry_verdict, "TOC visual geometry verdict"),
            (toc_paragraph_typography_verdict, "TOC paragraph-and-typography verdict"),
            (toc_visible_run_typography_verdict, "TOC visible-run typography verdict"),
        ):
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record must explicitly record {label} when TOC is touched")
        if contains_any(toc_restoration_summary, {"default", "not restored", "failed", "drift", "missing"}):
            issues.append("gate record TOC baseline restoration summary indicates unresolved TOC style drift")
        if contains_any(toc_rhythm_summary, {"compressed", "single page", "one page", "failed", "drift", "missing"}):
            issues.append("gate record TOC visual rhythm summary indicates unresolved TOC occupancy drift")
        if not surface_verdict_passes(toc_geometry_verdict) or contains_any(
            toc_geometry_verdict,
            {"compressed", "dense", "default", "font only", "content only", "not checked", "missing", "drift"},
        ):
            issues.append("gate record TOC visual geometry verdict indicates unresolved TOC rendered-geometry drift")
        if not surface_verdict_passes(toc_paragraph_typography_verdict) or contains_any(
            toc_paragraph_typography_verdict,
            {
                "style only",
                "font only",
                "geometry only",
                "compressed",
                "dense",
                "default",
                "smaller",
                "shrunken",
                "scaled down",
                "proportionally scaled",
                "not checked",
                "missing",
                "drift",
            },
        ):
            issues.append("gate record TOC paragraph-and-typography verdict indicates unresolved TOC paragraph or typography drift")
        if not surface_verdict_passes(toc_visible_run_typography_verdict) or contains_any(
            toc_visible_run_typography_verdict,
            {
                "style only",
                "paragraph only",
                "geometry only",
                "direct rpr missing",
                "rpr missing",
                "run font",
                "font only",
                "default",
                "theme alias",
                "not checked",
                "missing",
                "drift",
                "mismatch",
                "failed",
            },
        ):
            issues.append("gate record TOC visible-run typography verdict indicates unresolved visible run direct typography drift")
    if tail_block_touched_in_gate:
        for prefix in (
            "- tail-block pagination evidence paths:",
            "- tail-block rendered opener comparison evidence paths:",
        ):
            line = find_lines_with_prefix(record_lines, prefix)[0]
            value = parse_line_value(line)
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record field must not be empty when tail blocks are touched: {normalize(line)}")
        if not is_explicit(tail_block_summary) or tail_block_summary in EXPLICIT_VALUES:
            issues.append("gate record must explicitly record tail-block pagination summary when tail blocks are touched")
        if contains_any(
            tail_block_summary,
            {
                "failed",
                "drift",
                "same page",
                "shared page",
                "merged",
                "missing",
                "not checked",
                "lost owner",
                "owner missing",
                "pagebreakbefore=0",
                "no fresh page",
            },
        ):
            issues.append("gate record tail-block pagination summary indicates unresolved tail-block pagination drift")
    if table_touched_in_gate:
        for prefix in (
            "- table authority evidence paths:",
            "- table-local structure evidence paths:",
            "- table rendered baseline comparison evidence paths:",
            "- table title-mode evidence paths:",
            "- table header-bottom middle-rule evidence paths:",
            "- table continuation evidence paths:",
            "- table continuation summary:",
            "- cross-page table rendered pages:",
            "- continuation title outside-grid verdict:",
            "- table row split/header repeat verdict:",
        ):
            matching_lines = find_lines_with_prefix(record_lines, prefix)
            if not matching_lines:
                issues.append(f"gate record missing table evidence field when tables are touched: {prefix}")
                continue
            line = matching_lines[0]
            value = parse_line_value(line)
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"gate record field must not be empty when tables are touched: {normalize(line)}")
        table_authority_summary = parse_line_value(find_lines_with_prefix(record_lines, "- table authority summary:")[0])
        table_local_structure_summary = parse_line_value(find_lines_with_prefix(record_lines, "- table-local structure summary:")[0])
        if not is_explicit(table_authority_summary) or table_authority_summary in EXPLICIT_VALUES:
            issues.append("gate record must explicitly record table authority summary when tables are touched")
        if not is_explicit(table_local_structure_summary) or table_local_structure_summary in EXPLICIT_VALUES:
            issues.append("gate record must explicitly record table-local structure summary when tables are touched")
        if not is_explicit(table_title_mode_summary) or table_title_mode_summary in EXPLICIT_VALUES:
            issues.append("gate record must explicitly record table title-mode summary when tables are touched")
        if not is_explicit(table_header_middle_rule_summary) or table_header_middle_rule_summary in EXPLICIT_VALUES:
            issues.append("gate record must explicitly record table header-bottom middle-rule summary when tables are touched")
        table_continuation_summary_lines = find_lines_with_prefix(record_lines, "- table continuation summary:")
        table_continuation_summary = (
            parse_line_value(table_continuation_summary_lines[0])
            if table_continuation_summary_lines
            else "missing"
        )
        continuation_outside_grid_lines = find_lines_with_prefix(record_lines, "- continuation title outside-grid verdict:")
        row_split_header_lines = find_lines_with_prefix(record_lines, "- table row split/header repeat verdict:")
        if contains_any(table_continuation_summary, {"missing", "not checked", "sample only", "sampled-only", "unresolved"}):
            issues.append("gate record table continuation summary indicates cross-page table evidence was not closed")
        for label, matching_lines in (
            ("continuation title outside-grid verdict", continuation_outside_grid_lines),
            ("table row split/header repeat verdict", row_split_header_lines),
        ):
            value = parse_line_value(matching_lines[0]) if matching_lines else "missing"
            if not surface_verdict_passes(value, allow_not_applicable_with_reason=True):
                issues.append(f"gate record {label} must be pass or not-applicable-with-reason with rendered evidence")
        if contains_any(table_authority_summary, {"ambiguous", "guessed", "memory only", "failed", "drift", "missing", "multiple"}):
            issues.append("gate record table authority summary indicates unresolved table-authority selection")
        if contains_any(
            table_local_structure_summary,
            {
                "failed",
                "drift",
                "inside cell",
                "in cell",
                "trapped in cell",
                "caption in cell",
                "title in cell",
                "body in cell",
                "single cell",
                "misplaced into cell",
                "misplaced into table",
                "table-local contamination",
                "table71",
            },
        ):
            issues.append("gate record table-local structure summary indicates unresolved text-in-table-cell contamination")
        for token in ("title", "keep", "header-bottom", "middle rule"):
            if token not in table_local_structure_summary:
                issues.append(f"gate record table-local structure summary must name checked table surface token: {token}")
        if contains_any(table_title_mode_summary, {"guessed", "failed", "drift", "missing", "not checked", "mismatch", "builder-chosen"}):
            issues.append("gate record table title-mode summary indicates unresolved donor title-mode mismatch")
        if contains_any(table_header_middle_rule_summary, {"failed", "drift", "missing", "not checked", "absent", "lost", "no middle"}):
            issues.append("gate record table header-bottom middle-rule summary indicates unresolved three-line table middle-rule loss")
    if contains_any(chapter_summary, {"failed", "drift", "wrong page", "same page", "not checked", "missing"}):
        issues.append("gate record chapter-start pagination summary indicates unresolved chapter pagination drift")
    for value, label in (
        (format_preservation_promise_verdict, "format preservation promise verdict"),
        (chapter_format_preservation_detector_verdict, "chapter format preservation detector verdict"),
        (non_target_format_preservation_verdict, "non-target format preservation verdict"),
    ):
        if not surface_verdict_passes(value) or contains_any(
            value,
            {"failed", "drift", "not checked", "missing", "sample only", "sampled-only", "stale", "damaged"},
        ):
            issues.append(f"gate record {label} indicates unresolved chapter-level format preservation drift")
    if contains_any(
        whole_document_pagination_summary,
        {"failed", "drift", "wrong page", "same page", "not checked", "missing", "sample only", "sampled-only", "stale"},
    ):
        issues.append("gate record whole-document pagination summary indicates unresolved full-document pagination drift")
    if not surface_verdict_passes(whole_document_pagination_verdict) or contains_any(
        whole_document_pagination_verdict,
        {"failed", "drift", "not checked", "missing", "sample only", "sampled-only", "stale", "pdf export only"},
    ):
        issues.append("gate record whole-document pagination verdict indicates unresolved full-document pagination drift")
    if contains_any(table_summary, {"overwrite", "overwrote", "restyle", "lost tuning", "wrapped", "failed"}):
        issues.append("gate record custom-layout result-table preservation summary indicates unresolved table-restyle drift")
    if contains_any(figure_family_summary, {"failed", "drift", "generic", "mermaid", "wrong family", "sample mismatch"}):
        issues.append("gate record figure family style summary indicates unresolved figure-family style drift")
    if contains_any(table_family_summary, {"failed", "grid", "wrong family", "wrong style", "fallback mismatch", "shading", "tinted"}):
        issues.append("gate record table rendered-family summary indicates unresolved rendered table-family drift")
    if contains_any(wps_preset_summary, {"claimed applied but not verified", "false claim", "mismatch"}):
        issues.append("gate record WPS preset application summary makes an unsafe editor-authority claim")
    if contains_any(image_rendered_summary, {"failed", "old image", "stale", "wrong image", "blank"}):
        issues.append("gate record image replacement rendered summary indicates unresolved rendered-image replacement drift")
    if contains_any(image_metadata_summary, {"failed", "stale", "old name", "old alt", "mismatch"}):
        issues.append("gate record image metadata sync summary indicates unresolved picture metadata drift")
    if contains_any(formula_object_summary, {"failed", "plain text", "lost object", "missing object"}):
        issues.append("gate record formula object preservation summary indicates unresolved formula-object drift")
    if contains_any(formula_numbering_summary, {"failed", "missing", "below equation", "wrong side", "not far-right"}):
        issues.append("gate record formula numbering surface summary indicates unresolved formula-numbering drift")
    if contains_any(header_placement_summary, {"failed", "drift", "wrong position", "misaligned", "offset"}):
        issues.append("gate record header placement summary indicates unresolved header-placement drift")
    if contains_any(footer_indent_summary, {"failed", "drift", "indent", "offset", "wrong position"}):
        issues.append("gate record footer indent summary indicates unresolved footer-indent drift")
    if contains_any(
        footer_baseline_typography_summary,
        {"failed", "drift", "wrong font", "wrong size", "wrong style", "not restored"},
    ):
        issues.append("gate record footer baseline typography summary indicates unresolved footer-baseline typography drift")
    if contains_any(
        bibliography_baseline_summary,
        {
            "failed",
            "drift",
            "wrong",
            "mismatch",
            "abnormal",
            "unexpected",
            "indent drift",
            "hanging drift",
            "left indent drift",
            "first-line drift",
            "body indent residue",
            "\u7f29\u8fdb\u6f02\u79fb",
            "\u5f02\u5e38\u7f29\u8fdb",
        },
    ):
        issues.append("gate record bibliography baseline summary indicates unresolved bibliography formatting drift")
    if rendered_docx_path is not None:
        comment_text = extract_docx_comment_text(rendered_docx_path)
        if any(token.lower() in comment_text.lower() for token in BIBLIOGRAPHY_COMMENT_TOKENS):
            if not bibliography_baseline_summary.startswith("passed"):
                issues.append("gate record bibliography baseline summary does not prove repair while bibliography-related comments remain")
            if not bibliography_numbering_summary.startswith("passed"):
                issues.append("gate record bibliography numbering summary does not prove repair while bibliography-related comments remain")
            if bibliography_comment_aware_repair_summary != "passed":
                issues.append("gate record bibliography comment-aware repair summary indicates unresolved bibliography-comment repair")
    if project_local_helper_script_risk_summary.startswith("failed"):
        completed_restart = (
            project_local_helper_disposition == "clean-source-restart-completed"
            or canonical_source_restart_required in {"completed", "clean-source-restart-completed"}
        )
        if not contains_any(
            project_local_helper_script_preflight_summary,
            {"failed", "contaminated", "thick", "risky project-local"},
        ):
            issues.append("gate record project-local helper script preflight summary does not acknowledge detected thick local scripts")
        if not contains_any(
            contaminated_baseline_disposition,
            {"restart", "clean source", "clean-source", "audit-only", "blocked", "completed", "not used"},
        ):
            issues.append("gate record contaminated-baseline disposition must require clean-source restart or audit-only when risky project-local scripts are detected")
        if completed_restart:
            if project_local_helper_disposition != "clean-source-restart-completed":
                issues.append("gate record project-local helper disposition must be clean-source-restart-completed when restart is completed")
            if canonical_source_restart_required not in {"completed", "clean-source-restart-completed"}:
                issues.append("gate record canonical source restart status must be completed when restart is completed")
            if not is_explicit(clean_source_restart_source_path) or is_explicit_none(clean_source_restart_source_path):
                issues.append("gate record clean-source restart completed but no clean-source restart source path is recorded")
            if helper_script_provenance_summary != "canonical-skill-bundle":
                issues.append("gate record clean-source restart completion must use canonical-skill-bundle helper provenance")
            if not contains_any(contaminated_baseline_disposition, {"completed", "clean source", "clean-source", "not used"}):
                issues.append("gate record contaminated-baseline disposition must state completed clean-source restart and non-use of contaminated helpers")
        elif project_local_helper_disposition not in {"audit-only", "clean-source-restart-required"}:
            issues.append("gate record project-local helper disposition must be audit-only or clean-source-restart-required when risky project-local scripts are detected")
        if not completed_restart and canonical_source_restart_required not in {"yes", "required", "clean-source-restart-required"}:
            issues.append("gate record canonical source restart must be required when risky project-local scripts are detected")
        if not completed_restart:
            issues.append("gate record project-local helper script risk summary indicates thick project-local thesis rewrite scripts were detected")
    elif contains_any(contaminated_baseline_disposition, {"restart", "clean source", "clean-source", "audit-only", "blocked"}):
        issues.append("gate record contaminated-baseline disposition conflicts with a clean project-local helper script risk summary")

    if task_mode == "format-repair-only":
        if humanizer_route != "none":
            issues.append("format-repair-only gate record must set humanizer route decision to none")
        if humanizer_lang != "none":
            issues.append("format-repair-only gate record must set humanizer target language to none")
        if humanizer_scope != "none":
            issues.append("format-repair-only gate record must set humanizer scope to none")
    elif task_mode in {"thesis-only", "program-plus-thesis"}:
        if not is_explicit(humanizer_route) or humanizer_route in EXPLICIT_VALUES:
            issues.append("thesis-related gate record must explicitly record a humanizer route decision")
        if humanizer_route == "none":
            if not contains_any(subtask_text, NON_CONTENT_THESIS_HINT_TOKENS):
                issues.append("thesis-related gate record cannot use humanizer route none unless the subtask is explicitly non-content")
        else:
            if not is_explicit(humanizer_lang) or humanizer_lang in EXPLICIT_VALUES:
                issues.append("active humanizer route must record target language")
            if not is_explicit(humanizer_scope) or humanizer_scope in EXPLICIT_VALUES:
                issues.append("active humanizer route must record scope")
            if humanizer_route == "humanizer-zh" and not contains_any(humanizer_lang, {"zh", "chinese", "\u4e2d\u6587"}):
                issues.append("humanizer-zh route must target Chinese text")
            if humanizer_route == "humanizer" and not contains_any(humanizer_lang, {"en", "english", "\u82f1\u6587"}):
                issues.append("humanizer route must target English text")
            if humanizer_route == "both" and not contains_any(humanizer_lang, {"both", "zh+en", "\u4e2d\u6587+\u82f1\u6587", "\u4e2d\u82f1"}):
                issues.append("humanizer route both must declare both languages")
            if humanizer_route == "humanizer-zh":
                if "humanizer-zh" not in humanizer_evidence_seen_skills:
                    issues.append("humanizer-zh route must include humanizer evidence with skill name humanizer-zh")
                if humanizer_evidence_seen_skills - {"humanizer-zh"}:
                    issues.append("humanizer-zh route evidence must not include non-Chinese humanizer skill names")
                if "zh" not in humanizer_evidence_seen_languages:
                    issues.append("humanizer-zh route must include humanizer evidence with target language zh/Chinese")
            if humanizer_route == "humanizer":
                if "humanizer" not in humanizer_evidence_seen_skills:
                    issues.append("humanizer route must include humanizer evidence with skill name humanizer")
                if humanizer_evidence_seen_skills - {"humanizer"}:
                    issues.append("humanizer route evidence must not include Chinese humanizer skill names")
                if "en" not in humanizer_evidence_seen_languages:
                    issues.append("humanizer route must include humanizer evidence with target language en/English")
            if humanizer_route == "both":
                missing_skills = {"humanizer-zh", "humanizer"} - humanizer_evidence_seen_skills
                if missing_skills:
                    issues.append(
                        "humanizer route both must include evidence for skill names: "
                        + ", ".join(sorted(missing_skills))
                    )
                missing_languages = {"zh", "en"} - humanizer_evidence_seen_languages
                if missing_languages:
                    issues.append(
                        "humanizer route both must include evidence for target languages: "
                        + ", ".join(sorted(missing_languages))
                    )

    if task_mode == "format-repair-only":
        for prefix in (
            "- thesis rendered-page verification summary:",
            "- machine-vision summary by touched surface:",
        ):
            lines = find_lines_with_prefix(record_lines, prefix)
            if not lines:
                continue
            value = normalize(lines[0]).lower()
            if "html snapshots" in value or "rendered html" in value:
                issues.append(f"format-repair-only gate record cannot rely on HTML snapshots: {normalize(lines[0])}")

    record_text = "\n".join(record_lines)
    if contains_any(record_text, FORMULA_HINT_TOKENS) and contains_any(
        record_text, NUMBERING_PARAGRAPH_TOKENS
    ):
        issues.append("gate record cannot accept formula numbering paragraphs as a final result")

    return issues
