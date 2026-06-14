"""Evidence-record checks for validate_skill_gate."""

from __future__ import annotations

__all__ = [
    "check_review_evidence_record",
    "check_effective_font_evidence_record",
    "check_required_surface_evidence_record",
]

import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

try:
    from .validate_skill_gate_registry_core import (
        EXPLICIT_VALUES,
        IMAGE_EXTENSIONS,
        PDF_EXTENSIONS,
        REVIEW_EVIDENCE_SCHEMA,
        TEXT_EVIDENCE_EXTENSIONS,
    )
    from .validate_skill_gate_registry_records import (
        ABSTRACT_SURFACE_TOKENS,
        FORMULA_HINT_TOKENS,
        NUMBERING_PARAGRAPH_TOKENS,
        REVIEW_EVIDENCE_ALWAYS_REQUIRED_PREFIXES,
        REVIEW_EVIDENCE_MAYBE_NOT_APPLICABLE_PREFIXES,
        REVIEW_EVIDENCE_OPTIONAL_STATUS_PREFIXES,
        REVIEW_EVIDENCE_RENDERER_REQUIRED_PREFIXES,
        REVIEW_EVIDENCE_RENDERER_REQUIRED_TYPES,
        REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES,
    )
    from .validate_skill_gate_record_core import detect_review_evidence_surfaces, sha256_file
    from .audit_docx_review_artifacts import (
        validate_citation_run_reports,
        validate_review_artifact_reports,
    )
    from .validate_skill_gate_utils import (
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        is_explicit_or_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )
    from .toc_leader_audit import audit_docx_toc_dotted_leaders
except ImportError:
    from validate_skill_gate_registry_core import (
        EXPLICIT_VALUES,
        IMAGE_EXTENSIONS,
        PDF_EXTENSIONS,
        REVIEW_EVIDENCE_SCHEMA,
        TEXT_EVIDENCE_EXTENSIONS,
    )
    from validate_skill_gate_registry_records import (
        ABSTRACT_SURFACE_TOKENS,
        FORMULA_HINT_TOKENS,
        NUMBERING_PARAGRAPH_TOKENS,
        REVIEW_EVIDENCE_ALWAYS_REQUIRED_PREFIXES,
        REVIEW_EVIDENCE_MAYBE_NOT_APPLICABLE_PREFIXES,
        REVIEW_EVIDENCE_OPTIONAL_STATUS_PREFIXES,
        REVIEW_EVIDENCE_RENDERER_REQUIRED_PREFIXES,
        REVIEW_EVIDENCE_RENDERER_REQUIRED_TYPES,
        REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES,
    )
    from validate_skill_gate_record_core import detect_review_evidence_surfaces, sha256_file
    from audit_docx_review_artifacts import (
        validate_citation_run_reports,
        validate_review_artifact_reports,
    )
    from validate_skill_gate_utils import (
        contains_any,
        find_lines_with_prefix,
        is_explicit,
        is_explicit_none,
        is_explicit_or_none,
        normalize,
        parse_line_value,
        raw_line_value,
        read_lines,
        resolve_record_path,
        split_path_values,
        validate_existing_path,
    )
    from toc_leader_audit import audit_docx_toc_dotted_leaders

STRICT_ABSTRACT_SURFACE_FIELDS = {
    "zh_abstract_title": ["- Chinese abstract title confirmed:"],
    "zh_abstract_body": ["- Chinese abstract body confirmed:"],
    "zh_keyword_line": [
        "- Chinese keyword line confirmed:",
        "- Chinese keyword label/content run split confirmed:",
    ],
    "en_abstract_title": ["- English abstract title confirmed:"],
    "en_abstract_body": [
        "- English abstract body confirmed:",
        "- English abstract semantic parity confirmed:",
    ],
    "en_keyword_line": [
        "- English keyword line confirmed:",
        "- English keyword label/content run split confirmed:",
        "- English abstract semantic parity confirmed:",
    ],
}

STRUCTURAL_FIGURE_HINT_TOKENS = (
    "structural figure",
    "structural diagram",
    "draw.io",
    "drawio",
    "er diagram",
    "er\u56fe",
    "entity relationship",
    "\u5b9e\u4f53\u5173\u7cfb",
    "\u6570\u636e\u5e93er",
)
STRUCTURAL_FIGURE_PATH_PREFIXES = (
    "- structural figure geometry validation report path:",
    "- structural source-scale bbox map path:",
    "- structural inserted-scale geometry evidence path:",
    "- structural inserted-scale collision evidence path:",
    "- structural dense-zone crop evidence paths:",
)
STRUCTURAL_FIGURE_VERDICT_PREFIXES = (
    "- structural relation-attribute collision verdict:",
    "- structural shape-overlap verdict:",
    "- structural inserted-scale collision verdict:",
    "- structural source-to-inserted geometry verdict:",
)
STRUCTURAL_FIGURE_CONFIRM_PREFIXES = (
    "- structural source-scale collision report confirmed:",
    "- structural inserted-scale dense-zone review confirmed:",
)

STRICT_TOC_SURFACE_FIELDS = {
    "toc_title": ["- TOC title baseline confirmed:", "- TOC title paragraph confirmed:"],
    "toc_entries": ["- TOC level formatting confirmed:", "- TOC entries by level confirmed:"],
    "toc_dotted_leaders": [
        "- TOC dotted-leader / right-tab confirmed:",
        "- TOC dotted leaders confirmed:",
    ],
    "toc_page_number_column": [
        "- TOC page-number column confirmed:",
        "- TOC page-number column per entry confirmed:",
    ],
}

STRICT_TOC_VISUAL_GEOMETRY_PREFIXES = [
    "- TOC template rendered page/region image path:",
    "- TOC actual rendered page/region image path:",
    "- TOC visual comparison method:",
    "- TOC title bbox baseline/actual:",
    "- TOC first-entry bbox baseline/actual:",
    "- TOC row bbox map:",
    "- TOC per-level left-indent x baseline/actual:",
    "- TOC line-spacing y-delta baseline/actual:",
    "- TOC dotted-leader start/end/density baseline/actual:",
    "- TOC page-number x column baseline/actual:",
    "- TOC row count per page baseline/actual:",
    "- TOC title-to-first-entry gap baseline/actual:",
    "- TOC page occupancy rhythm baseline/actual:",
    "- TOC visual geometry verdict:",
]

STRICT_TOC_PARAGRAPH_TYPOGRAPHY_PREFIXES = [
    "- TOC style binding baseline/actual:",
    "- TOC WPS paragraph-dialog metrics baseline/actual:",
    "- TOC title typography baseline/actual:",
    "- TOC per-level typography baseline/actual:",
    "- TOC per-level paragraph spacing baseline/actual:",
    "- TOC per-level line-spacing mode/value baseline/actual:",
    "- TOC per-level indentation chars/points baseline/actual:",
    "- TOC per-level tab stop leader baseline/actual:",
    "- TOC visible run typography baseline/actual:",
    "- TOC per-level visible run typography baseline/actual:",
    "- TOC page-number run typography baseline/actual:",
    "- TOC tab/leader run typography baseline/actual:",
    "- TOC run typography verdict:",
    "- TOC scale/compression verdict:",
    "- TOC paragraph-and-typography verdict:",
    "- TOC used-level inventory:",
    "- TOC used-level evidence map:",
]
STRICT_TOC_RIGHT_TAB_PAGE_COLUMN_PREFIXES = [
    "- TOC right-tab stop semantic baseline/actual:",
    "- TOC page-number column right alignment baseline/actual:",
    "- TOC page-number tab leader ownership baseline/actual:",
    "- TOC per-entry right-tab/page-number verdict:",
]

STRICT_WHOLE_DOCUMENT_PAGINATION_PREFIXES = [
    "- package baseline manifest path:",
    "- package drift report path:",
    "- package drift verdict:",
    "- pre-mutation page map path:",
    "- post-mutation page map path:",
    "- whole-document pagination diff path:",
    "- DOCX pagination structure schema:",
    "- DOCX pagination structure generator:",
    "- DOCX pagination structure evidence path:",
    "- DOCX pagination structure verdict:",
    "- section count baseline/actual:",
    "- header/footer reference map baseline/actual:",
    "- header/footer link-to-previous inferred map baseline/actual:",
    "- section boundary map baseline/actual:",
    "- section property map baseline/actual:",
    "- page-number format/restart map baseline/actual:",
    "- header/footer link-to-previous map baseline/actual:",
    "- hard page-break / section-break map baseline/actual:",
    "- fatal pagination topology differences:",
    "- allowed content-growth pagination differences:",
    "- all pagination topology differences:",
    "- section count verdict:",
    "- header/footer reference verdict:",
    "- page-number restart verdict:",
    "- field-refresh before/after state:",
    "- TOC-to-heading page sync map:",
    "- logical-to-physical page map:",
    "- rendered page count baseline/actual:",
    "- blank/near-empty page scan verdict:",
    "- chapter opener page map:",
    "- tail-block opener page map:",
    "- page-class occupancy rhythm verdict:",
    "- whole-document pagination verdict:",
]

TOC_USED_LEVEL_ALLOWED = ("title", "level1", "level2", "level3", "level4")
TOC_USED_LEVEL_CONDITIONAL_TOKENS = {
    "when present",
    "whenpresent",
    "if present",
    "ifpresent",
    "if used",
    "ifused",
    "when used",
    "whenused",
    "optional",
}

WHOLE_DOCUMENT_PAGINATION_REQUIRED_TOKENS = (
    "section",
    "page",
    "field",
    "toc",
    "logical",
    "physical",
    "chapter",
    "tail",
)

WHOLE_DOCUMENT_PAGINATION_FAIL_TOKENS = {
    "sampled-only",
    "sample only",
    "screenshot only",
    "pdf export only",
    "not checked",
    "missing",
    "blocked",
    "failed",
    "stale",
    "unresolved",
}

STRICT_SURFACE_PARAGRAPH_TYPOGRAPHY_PREFIXES = [
    "- surface style binding baseline/actual:",
    "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
    "- surface typography baseline/actual:",
    "- surface paragraph spacing baseline/actual:",
    "- surface line-spacing mode/value baseline/actual:",
    "- surface indentation chars/points baseline/actual:",
    "- surface tab stop leader baseline/actual:",
    "- surface keep/list/page-break baseline/actual:",
    "- surface scale/compression verdict:",
    "- surface paragraph-and-typography verdict:",
]
STRICT_COVER_MEDIA_PREFIXES = [
    "- cover media/icon relationship ids baseline/actual:",
    "- cover media/icon package targets baseline/actual:",
    "- cover media/icon binding verdict:",
]
STRICT_FRONT_MATTER_HARD_FIELD_PREFIXES = [
    "- front-matter hard-field paragraph metrics baseline/actual:",
    "- front-matter hard-field run typography baseline/actual:",
    "- front-matter hard-field verdict:",
]
STRICT_HEADER_FULL_DISPLAY_PREFIXES = [
    "- header expected full display string:",
    "- header observed rendered full display string:",
    "- header full-display string verdict:",
]
STRICT_REFERENCES_ENTRIES_FONT_SIZE_PREFIXES = [
    "- references entries font-size baseline/actual:",
    "- references entries per-entry font-size map:",
    "- references entries font-size verdict:",
]
STRICT_ACKNOWLEDGEMENT_TITLE_STYLE_PREFIXES = [
    "- acknowledgement title style baseline/actual:",
    "- acknowledgement title paragraph style verdict:",
]
STRICT_FOOTER_PAGE_NUMBER_FONT_SIZE_PREFIXES = [
    "- footer page-number font-size baseline/actual:",
    "- footer page-number run path map:",
    "- footer page-number font-size verdict:",
]

STRICT_KEYWORD_RUN_SPLIT_PREFIXES = [
    "- keyword run split extraction method:",
    "- keyword label run text baseline/actual:",
    "- keyword label run isolated baseline/actual:",
    "- keyword content run count baseline/actual:",
    "- keyword label bold/strong baseline/actual:",
    "- keyword content bold baseline/actual:",
    "- keyword separator baseline/actual:",
    "- keyword run split verdict:",
]

STRICT_SURFACE_GEOMETRY_PREFIXES = [
    "- template rendered region image path:",
    "- actual rendered region image path:",
    "- surface geometry comparison method:",
    "- surface crop schema:",
    "- surface crop generator:",
    "- surface crop source page images baseline/actual:",
    "- surface crop source image sha256 baseline/actual:",
    "- surface crop source image size baseline/actual:",
    "- surface crop fraction map baseline/actual:",
    "- surface crop threshold baseline/actual:",
    "- surface page index baseline/actual:",
    "- surface crop bbox baseline/actual:",
    "- surface crop source image size baseline/actual:",
    "- surface crop fraction map baseline/actual:",
    "- surface crop threshold baseline/actual:",
    "- surface content bbox baseline/actual:",
    "- surface nonwhite ratio baseline/actual:",
    "- surface blank crop verdict:",
    "- surface binding method:",
    "- surface bbox baseline/actual:",
    "- surface position baseline/actual:",
    "- surface size baseline/actual:",
    "- surface line-height y-delta baseline/actual:",
    "- surface spacing before/after baseline/actual:",
    "- surface indentation/tab baseline/actual:",
    "- surface page occupancy baseline/actual:",
    "- surface geometry verdict:",
]

STRICT_SURFACE_GEOMETRY_IMAGE_PREFIXES = [
    "- template rendered region image path:",
    "- actual rendered region image path:",
]

STRICT_SURFACE_GEOMETRY_NUMERIC_PREFIXES = (
    "- surface bbox baseline/actual:",
    "- surface position baseline/actual:",
    "- surface size baseline/actual:",
    "- surface crop bbox baseline/actual:",
    "- surface content bbox baseline/actual:",
    "- surface nonwhite ratio baseline/actual:",
    "- surface line-height y-delta baseline/actual:",
    "- surface spacing before/after baseline/actual:",
    "- surface indentation/tab baseline/actual:",
    "- surface page occupancy baseline/actual:",
)

SURFACE_GEOMETRY_REQUIRED_TOKENS = (
    "template",
    "actual",
    "x",
    "y",
    "w",
    "h",
    "line",
    "spacing",
    "occupancy",
)

SURFACE_GEOMETRY_FAIL_TOKENS = {
    "looks correct",
    "looks matched",
    "looks consistent",
    "visual pass",
    "screenshot only",
    "not checked",
    "mismatch",
    "drift",
    "failed",
}

STRICT_TOC_VISUAL_GEOMETRY_IMAGE_PREFIXES = [
    "- TOC template rendered page/region image path:",
    "- TOC actual rendered page/region image path:",
]

STRICT_TOC_SURFACE_IDS = set(STRICT_TOC_SURFACE_FIELDS)

STRICT_TEMPLATE_SURFACE_FIELDS = {
    "cover_style": [
        "- cover page-class baseline confirmed:",
        "- cover identity-zone baseline confirmed:",
        "- cover identity value-line baseline confirmed:",
    ],
    "declaration_or_title_front_matter": [
        "- declaration/title front matter baseline confirmed:",
        "- declaration separated from cover confirmed:",
    ],
    "body_heading_levels": [
        "- heading level baseline confirmed:",
        "- heading direct-run typography confirmed:",
        "- heading paragraph metrics confirmed:",
        "- heading body-format residue cleared confirmed:",
        "- heading TOC/chapter-start sync confirmed:",
    ],
    "body_text": [],
    "body_citation_superscripts": [],
    "review_comments_and_change_marks": [],
    "figure_table_captions_and_holders": [],
    "references_title": [
        "- tail-block title baseline confirmed:",
        "- references title indentation confirmed:",
        "- end-matter rendered geometry confirmed:",
    ],
    "references_entries": [
        "- references entries indentation confirmed:",
        "- end-matter rendered geometry confirmed:",
    ],
    "acknowledgement_title": [
        "- tail-block title baseline confirmed:",
        "- acknowledgement title indentation confirmed:",
        "- end-matter rendered geometry confirmed:",
    ],
    "acknowledgement_body": [
        "- acknowledgement body indentation confirmed:",
        "- end-matter rendered geometry confirmed:",
    ],
    "appendix_title": [],
    "appendix_body": [],
    "header": ["- header/footer baseline confirmed:"],
    "footer": ["- header/footer baseline confirmed:", "- footer/page-number presentation confirmed:"],
    "page_numbers": ["- footer/page-number presentation confirmed:", "- page-number structure confirmed:"],
    "whole_document_pagination": [],
}
STRICT_TEMPLATE_SURFACE_FIELDS.update(STRICT_ABSTRACT_SURFACE_FIELDS)
STRICT_TEMPLATE_SURFACE_FIELDS.update(STRICT_TOC_SURFACE_FIELDS)

STRICT_FRONT_MATTER_SURFACE_FIELDS = STRICT_TEMPLATE_SURFACE_FIELDS

STRICT_SURFACE_TARGET_ALIASES = {
    "cover_style": ("cover_style", "cover style", "cover"),
    "declaration_or_title_front_matter": (
        "declaration_or_title_front_matter",
        "declaration",
        "title front matter",
        "title/front matter",
    ),
    "zh_abstract_title": ("zh_abstract_title", "chinese abstract title", "中文摘要标题"),
    "zh_abstract_body": ("zh_abstract_body", "chinese abstract body", "中文摘要正文"),
    "zh_keyword_line": ("zh_keyword_line", "chinese keyword line", "中文关键词"),
    "en_abstract_title": ("en_abstract_title", "english abstract title"),
    "en_abstract_body": ("en_abstract_body", "english abstract body"),
    "en_keyword_line": ("en_keyword_line", "english keyword line", "keywords", "key words"),
    "toc_title": ("toc_title", "toc title", "table of contents title", "目录标题"),
    "toc_entries": ("toc_entries", "toc entries", "table of contents entries", "目录条目"),
    "toc_dotted_leaders": ("toc_dotted_leaders", "dotted leaders", "点引导线"),
    "toc_page_number_column": ("toc_page_number_column", "page-number column", "page number column", "页码列"),
    "body_heading_levels": ("body_heading_levels", "body heading levels", "heading levels"),
    "body_text": ("body_text", "body text", "mixed-script body text", "正文", "正文段落"),
    "body_citation_superscripts": (
        "body_citation_superscripts",
        "body citation superscripts",
        "citation superscript",
        "citation marker runs",
    ),
    "review_comments_and_change_marks": (
        "review_comments_and_change_marks",
        "review comments",
        "comment anchors",
        "tracked changes",
        "change marks",
    ),
    "figure_table_captions_and_holders": (
        "figure_table_captions_and_holders",
        "figure/table captions",
        "figure holder",
        "table title",
        "caption",
    ),
    "references_title": ("references_title", "references title", "bibliography title"),
    "references_entries": ("references_entries", "reference entries", "references entries", "bibliography entries"),
    "acknowledgement_title": ("acknowledgement_title", "acknowledgement title", "acknowledgment title"),
    "acknowledgement_body": ("acknowledgement_body", "acknowledgement body", "acknowledgment body"),
    "appendix_title": ("appendix_title", "appendix title"),
    "appendix_body": ("appendix_body", "appendix body"),
    "header": ("header", "page header"),
    "footer": ("footer", "page footer"),
    "page_numbers": ("page_numbers", "page numbers", "page-number", "page number"),
    "whole_document_pagination": (
        "whole_document_pagination",
        "whole document pagination",
        "full document pagination",
        "pagination chain",
        "section page map",
    ),
}

STRICT_SURFACE_FALLBACK_TOKENS = {
    "officecli only",
    "officecli-only",
    "officecli check",
    "officecli checked",
    "officecli inspect",
    "officecli inspected",
    "officecli inspection",
    "officecli validate",
    "officecli view issues",
    "pdf export only",
    "pdf rendered only",
    "successful pdf export",
    "page image exists",
    "screenshot only",
    "style name only",
    "visible text only",
    "title present",
    "title exists",
    "visible title",
    "keyword line present",
    "keyword exists",
    "text exists",
    "text present",
    "entries visible",
    "page order only",
    "page numbers corrected",
    "dotted leaders visible",
    "font only",
    "font-chain only",
    "content only",
    "text only",
    "page-number only",
    "page number only",
    "field only",
    "bookmark only",
    "geometry not checked",
}

TOC_VISUAL_GEOMETRY_REQUIRED_TOKENS = (
    "template",
    "actual",
    "bbox",
    "x",
    "y",
    "line",
    "leader",
    "density",
    "page-number",
    "row",
    "gap",
    "occupancy",
)

TOC_VISUAL_GEOMETRY_FAIL_TOKENS = {
    "compressed",
    "too dense",
    "squeezed",
    "default layout",
    "default toc",
    "fallback styling",
    "font only",
    "content only",
    "page-number only",
    "not checked",
    "missing",
    "mismatch",
    "drift",
    "failed",
}

TOC_VISUAL_GEOMETRY_NUMERIC_PREFIXES = (
    "- TOC title bbox baseline/actual:",
    "- TOC first-entry bbox baseline/actual:",
    "- TOC row bbox map:",
    "- TOC per-level left-indent x baseline/actual:",
    "- TOC line-spacing y-delta baseline/actual:",
    "- TOC dotted-leader start/end/density baseline/actual:",
    "- TOC page-number x column baseline/actual:",
    "- TOC row count per page baseline/actual:",
    "- TOC title-to-first-entry gap baseline/actual:",
    "- TOC page occupancy rhythm baseline/actual:",
)

TOC_PARAGRAPH_TYPOGRAPHY_REQUIRED_TOKENS = (
    "template",
    "actual",
    "style",
    "font",
    "size",
    "spacing",
    "line",
    "indent",
    "tab",
    "leader",
    "level",
    "visible run",
    "directrpr",
    "page-number",
)

TOC_PARAGRAPH_TYPOGRAPHY_FAIL_TOKENS = {
    "style only",
    "font only",
    "geometry only",
    "not checked",
    "missing",
    "mismatch",
    "drift",
    "failed",
    "compressed",
    "too dense",
    "squeezed",
    "shrunken",
    "shrunk",
    "scaled down",
    "proportionally scaled",
    "default layout",
    "default toc",
    "default application styling",
    "same style for all levels",
    "undifferentiated",
}

TOC_PARAGRAPH_TYPOGRAPHY_NUMERIC_PREFIXES = (
    "- TOC WPS paragraph-dialog metrics baseline/actual:",
    "- TOC title typography baseline/actual:",
    "- TOC per-level typography baseline/actual:",
    "- TOC per-level paragraph spacing baseline/actual:",
    "- TOC per-level line-spacing mode/value baseline/actual:",
    "- TOC per-level indentation chars/points baseline/actual:",
    "- TOC per-level tab stop leader baseline/actual:",
    "- TOC visible run typography baseline/actual:",
    "- TOC per-level visible run typography baseline/actual:",
    "- TOC page-number run typography baseline/actual:",
    "- TOC tab/leader run typography baseline/actual:",
)

SURFACE_PARAGRAPH_TYPOGRAPHY_REQUIRED_TOKENS = (
    "template",
    "actual",
    "style",
    "font",
    "before",
    "after",
    "line",
    "left",
)

SURFACE_PARAGRAPH_TYPOGRAPHY_FAIL_TOKENS = {
    "style name only",
    "style-only",
    "screenshot only",
    "font chain only",
    "geometry only",
    "not checked",
    "missing",
    "blocked",
    "unresolved",
    "unknown",
    "default application",
    "default-app",
    "generic",
    "looks correct",
    "looks matched",
}

SURFACE_PARAGRAPH_TYPOGRAPHY_NUMERIC_PREFIXES = (
    "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
    "- surface typography baseline/actual:",
    "- surface paragraph spacing baseline/actual:",
    "- surface line-spacing mode/value baseline/actual:",
    "- surface indentation chars/points baseline/actual:",
    "- surface tab stop leader baseline/actual:",
)

NUMERIC_MEASUREMENT_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")

EFFECTIVE_FONT_DETAIL_PREFIXES = [
    "- baseline effective font chain:",
    "- actual effective font chain:",
    "- effective font slots compared:",
    "- theme/default font alias verdict:",
    "- WPS/Word UI font display evidence:",
    "- effective font-chain verdict:",
]

EFFECTIVE_FONT_CHAIN_REQUIRED_TOKENS = (
    "direct",
    "style",
    "docdefaults",
    "theme",
)

EFFECTIVE_FONT_SLOT_TOKENS = ("ascii", "hansi", "eastasia", "cs")

THEME_ALIAS_TOKENS = (
    "minoreastasia",
    "minorhansi",
    "majoreastasia",
    "majorhansi",
    "asciitheme",
    "hansitheme",
    "eastasiatheme",
    "cstheme",
    "calibri (",
    "calibri(",
    "calibri \uff08",
    "calibri\uff08",
    "(\u6b63\u6587)",
    "\uff08\u6b63\u6587\uff09",
)


def check_review_evidence_record(
    evidence_path: Path,
    *,
    expected_type: str,
    acceptance_mode: str,
) -> list[str]:
    issues: list[str] = []
    issues.extend(validate_existing_path(evidence_path, require_nonempty_file=True))
    if issues:
        return issues
    if evidence_path.suffix.lower() not in TEXT_EVIDENCE_EXTENSIONS:
        return [f"evidence record must be a text file (.md or .txt): {evidence_path}"]

    try:
        lines = read_lines(evidence_path)
    except UnicodeDecodeError as exc:
        return [f"evidence record is not valid UTF-8: {evidence_path} ({exc})"]

    normalized_lines = {normalize(line) for line in lines if normalize(line)}

    for heading in REVIEW_EVIDENCE_SCHEMA["headings"]:
        if heading not in normalized_lines:
            issues.append(f"evidence record missing heading in {evidence_path}: {heading}")

    for prefix in REVIEW_EVIDENCE_SCHEMA["single_prefixes"]:
        matched = find_lines_with_prefix(lines, prefix)
        if len(matched) != 1:
            issues.append(f"evidence record must contain exactly one '{prefix}' line: {evidence_path}")

    if issues:
        return issues

    values = {
        prefix: parse_line_value(find_lines_with_prefix(lines, prefix)[0])
        for prefix in REVIEW_EVIDENCE_SCHEMA["single_prefixes"]
    }
    raw_values = {
        prefix: raw_line_value(find_lines_with_prefix(lines, prefix)[0])
        for prefix in REVIEW_EVIDENCE_SCHEMA["single_prefixes"]
    }
    extra_prefixes = (
        list(STRICT_KEYWORD_RUN_SPLIT_PREFIXES)
        + list(STRICT_SURFACE_GEOMETRY_PREFIXES)
        + list(STRICT_SURFACE_PARAGRAPH_TYPOGRAPHY_PREFIXES)
        + list(STRICT_TOC_VISUAL_GEOMETRY_PREFIXES)
        + list(STRICT_TOC_PARAGRAPH_TYPOGRAPHY_PREFIXES)
        + list(STRICT_TOC_RIGHT_TAB_PAGE_COLUMN_PREFIXES)
        + list(STRICT_WHOLE_DOCUMENT_PAGINATION_PREFIXES)
        + list(STRICT_COVER_MEDIA_PREFIXES)
        + list(STRICT_FRONT_MATTER_HARD_FIELD_PREFIXES)
        + list(STRICT_HEADER_FULL_DISPLAY_PREFIXES)
        + list(STRICT_REFERENCES_ENTRIES_FONT_SIZE_PREFIXES)
        + list(STRICT_ACKNOWLEDGEMENT_TITLE_STYLE_PREFIXES)
        + list(STRICT_FOOTER_PAGE_NUMBER_FONT_SIZE_PREFIXES)
    )
    for prefix in extra_prefixes:
        matched = find_lines_with_prefix(lines, prefix)
        if len(matched) == 1:
            values[prefix] = parse_line_value(matched[0])
            raw_values[prefix] = raw_line_value(matched[0])
        elif prefix not in values:
            values[prefix] = ""
            raw_values[prefix] = ""

    if values["- evidence type:"] != expected_type:
        issues.append(
            f"evidence record type mismatch in {evidence_path}: expected {expected_type}, found {values['- evidence type:']}"
        )

    if values["- task mode:"] not in {acceptance_mode, "shared"}:
        issues.append(
            f"evidence record task mode mismatch in {evidence_path}: expected {acceptance_mode} or shared, found {values['- task mode:']}"
        )

    for prefix in REVIEW_EVIDENCE_ALWAYS_REQUIRED_PREFIXES:
        if not is_explicit(values[prefix]) or values[prefix] in EXPLICIT_VALUES:
            issues.append(f"evidence record field must not be empty in {evidence_path}: {prefix}")

    for prefix in REVIEW_EVIDENCE_MAYBE_NOT_APPLICABLE_PREFIXES:
        value = values[prefix]
        if expected_type == "program-review":
            if not is_explicit_or_none(value):
                issues.append(f"evidence record field is incomplete in {evidence_path}: {prefix}")
        else:
            if not is_explicit(value) or value in EXPLICIT_VALUES:
                issues.append(f"evidence record field must not be empty in {evidence_path}: {prefix}")

    if expected_type in REVIEW_EVIDENCE_RENDERER_REQUIRED_TYPES:
        for prefix in REVIEW_EVIDENCE_RENDERER_REQUIRED_PREFIXES:
            if not is_explicit(values[prefix]) or values[prefix] in EXPLICIT_VALUES:
                issues.append(f"evidence record renderer field must not be empty in {evidence_path}: {prefix}")

    for prefix in REVIEW_EVIDENCE_OPTIONAL_STATUS_PREFIXES:
        if not is_explicit_or_none(values[prefix]):
            issues.append(f"evidence record field is incomplete in {evidence_path}: {prefix}")

    if values["- result:"] != "pass":
        issues.append(f"evidence record result is not pass in {evidence_path}: {values['- result:']}")
    if normalize(values["- machine-vision verdict:"]).lower() != "pass":
        issues.append(
            f"evidence record machine-vision verdict is not pass in {evidence_path}: {values['- machine-vision verdict:']}"
        )

    evidence_text = normalize("\n".join(raw_values.values())).lower()
    if contains_any(
        evidence_text,
        {"failed", "fail", "not checked", "partial", "sampled-only", "sample only", "unresolved", "blocked"},
    ):
        issues.append(f"evidence record contains unresolved review language in {evidence_path}")

    if not is_explicit_or_none(values["- blocker:"]):
        issues.append(f"evidence record blocker field is incomplete in {evidence_path}")

    if issues:
        return issues

    reviewed_output = resolve_record_path(raw_values["- reviewed output path:"], evidence_path)
    issues.extend(validate_existing_path(reviewed_output, require_nonempty_file=False))
    reviewed_sha = values["- reviewed output sha256:"]
    if reviewed_sha == "missing" or len(reviewed_sha) != 64 or any(ch not in "0123456789abcdefABCDEF" for ch in reviewed_sha):
        issues.append(f"evidence record reviewed output sha256 must be a 64-hex value in {evidence_path}")
    elif reviewed_output.exists():
        actual_sha = sha256_file(reviewed_output)
        if actual_sha.lower() != reviewed_sha.lower():
            issues.append(f"evidence record reviewed output sha256 does not match reviewed output path in {evidence_path}")
    if acceptance_mode in {"thesis-only", "format-repair-only", "program-plus-thesis"} and expected_type != "program-review":
        def resolved_optional(prefix: str) -> Path | None:
            value = values.get(prefix, "")
            if not is_explicit(value) or is_explicit_none(value):
                return None
            return resolve_record_path(raw_values.get(prefix, ""), evidence_path)

        source_review_artifact_inventory_path = resolved_optional("- source review-artifact inventory path:")
        final_review_artifact_diff_path = resolved_optional("- final review-artifact diff path:")
        source_body_citation_run_inventory_path = resolved_optional("- source body-citation run inventory path:")
        final_body_citation_run_diff_path = resolved_optional("- final body-citation run diff path:")
        citation_audit_source_to_final_run_diff_path = resolved_optional("- citation audit source-to-final run diff path:")
        if source_review_artifact_inventory_path is not None and final_review_artifact_diff_path is not None:
            issues.extend(
                validate_review_artifact_reports(
                    source_review_artifact_inventory_path,
                    final_review_artifact_diff_path,
                    expected_final_docx=reviewed_output,
                )
            )
        if source_body_citation_run_inventory_path is not None and final_body_citation_run_diff_path is not None:
            issues.extend(
                validate_citation_run_reports(
                    source_body_citation_run_inventory_path,
                    final_body_citation_run_diff_path,
                    expected_final_docx=reviewed_output,
                )
            )
        if citation_audit_source_to_final_run_diff_path is not None and final_body_citation_run_diff_path is not None:
            if citation_audit_source_to_final_run_diff_path.resolve() != final_body_citation_run_diff_path.resolve():
                issues.append(
                    f"evidence record citation audit source-to-final run diff path must match final body-citation diff path in {evidence_path}"
                )
    if values["- evidence created after mutation?:"] not in {"yes", "locked-after-mutation"}:
        issues.append(
            f"evidence record must state it was created after mutation in {evidence_path}: "
            f"{values['- evidence created after mutation?:']}"
        )

    if "html snapshot" in normalize(raw_values["- summary:"]).lower():
        issues.append(f"evidence record summary relies on HTML snapshots in {evidence_path}")
    if "html snapshot" in normalize(raw_values["- checks performed:"]).lower():
        issues.append(f"evidence record checks rely on HTML snapshots in {evidence_path}")
    forbidden_substitute = values.get("- forbidden substitute evidence used?:", "no")
    if forbidden_substitute != "no":
        issues.append(
            f"evidence record must explicitly reject substitute evidence in {evidence_path}: "
            f"{forbidden_substitute}"
        )
    if expected_type in REVIEW_EVIDENCE_RENDERER_REQUIRED_TYPES:
        issues.extend(
            _validate_surface_geometry_fields(
                values,
                raw_values,
                evidence_path,
                subject=f"{expected_type} evidence",
            )
        )
        issues.extend(
            _validate_surface_paragraph_typography_fields(
                values,
                raw_values,
                evidence_path,
                subject=f"{expected_type} evidence",
            )
        )

    surface_flags, combined_text = detect_review_evidence_surfaces(raw_values)
    formula_related = surface_flags["formula"]
    abstract_related = surface_flags["abstract"]
    caption_related = surface_flags["caption"]
    header_footer_related = surface_flags["header_footer"]
    toc_related = surface_flags["toc"]
    body_heading_related = surface_flags.get("body_heading", False)
    table_related = surface_flags["table"]
    tail_block_related = surface_flags["tail_block"]
    lowered_combined = normalize(combined_text).lower()
    if body_heading_related and normalize(raw_values.get("- target identifier:", "")).lower() == "body_heading_levels":
        toc_related = False

    if "cite_ref_" in lowered_combined or "bookmark_" in lowered_combined:
        issues.append(f"evidence record leaks internal citation anchor text in {evidence_path}")
    if "toc_placeholder" in lowered_combined:
        issues.append(f"evidence record still describes visible TOC placeholder leakage in {evidence_path}")
    if expected_type == "figure-review":
        if "structural diagram" in lowered_combined and "screenshot" in lowered_combined:
            issues.append(f"figure-review evidence indicates a structural diagram is occupying a runtime screenshot slot in {evidence_path}")
        if "route mismatch" in lowered_combined or "wrong route" in lowered_combined:
            issues.append(f"figure-review evidence indicates a runtime screenshot route mismatch in {evidence_path}")
        figure_context = normalize(
            "\n".join(
                raw_values.get(prefix, "")
                for prefix in (
                    "- target surface:",
                    "- target identifier:",
                    "- target pages or region:",
                    "- checks performed:",
                    "- summary:",
                )
            )
        ).lower()
        if contains_any(figure_context, set(STRUCTURAL_FIGURE_HINT_TOKENS)):
            issues.extend(_validate_structural_figure_geometry_fields(values, raw_values, evidence_path))

    if toc_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["toc"]:
            if values[prefix] != "yes":
                issues.append(
                    f"toc-related evidence must confirm yes for {prefix} in {evidence_path}"
                )

    if body_heading_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["body_heading"]:
            if values[prefix] != "yes":
                issues.append(
                    f"body-heading evidence must confirm yes for {prefix} in {evidence_path}"
                )

    if table_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["table"]:
            if values[prefix] != "yes":
                issues.append(
                    f"table-related evidence must confirm yes for {prefix} in {evidence_path}"
                )

    if formula_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["formula"]:
            if values[prefix] != "yes":
                issues.append(
                    f"formula evidence must confirm yes for {prefix} in {evidence_path}"
                )
        if contains_any(combined_text, NUMBERING_PARAGRAPH_TOKENS):
            issues.append(
                f"formula evidence cannot describe accepted output as numbering paragraph(s): {evidence_path}"
            )

    if abstract_related:
        abstract_confirmed = raw_values["- abstract surfaces confirmed:"].lower()
        if parse_line_value(find_lines_with_prefix(lines, "- abstract surfaces confirmed:")[0]) in EXPLICIT_VALUES:
            issues.append(
                f"abstract-related evidence must confirm the abstract surfaces instead of not-applicable: {evidence_path}"
            )
        missing_surfaces = [
            token for token in ABSTRACT_SURFACE_TOKENS if token not in abstract_confirmed
        ]
        if missing_surfaces:
            issues.append(
                f"abstract-related evidence is missing abstract surfaces in {evidence_path}: {', '.join(sorted(missing_surfaces))}"
            )

    if caption_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["caption"]:
            if values[prefix] != "yes":
                issues.append(
                    f"caption-related evidence must confirm yes for {prefix} in {evidence_path}"
                )

    if header_footer_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["header_footer"]:
            if values[prefix] != "yes":
                issues.append(
                    f"header/footer-related evidence must confirm yes for {prefix} in {evidence_path}"
                )

    if tail_block_related:
        for prefix in REVIEW_EVIDENCE_SURFACE_CONFIRM_PREFIXES["tail_block"]:
            if values[prefix] != "yes":
                issues.append(
                    f"tail-block-related evidence must confirm yes for {prefix} in {evidence_path}"
                )

    artifact_extensions = set()

    for raw_artifact in split_path_values(raw_values["- artifact path:"]):
        artifact = resolve_record_path(raw_artifact, evidence_path)
        issues.extend(validate_existing_path(artifact, require_nonempty_file=True))
        artifact_extensions.add(artifact.suffix.lower())

    if expected_type in REVIEW_EVIDENCE_RENDERER_REQUIRED_TYPES:
        renderer_path = resolve_record_path(raw_values["- renderer executable path:"], evidence_path)
        issues.extend(validate_existing_path(renderer_path, require_nonempty_file=True))

        rendered_pdf = resolve_record_path(raw_values["- rendered PDF path:"], evidence_path)
        issues.extend(validate_existing_path(rendered_pdf, require_nonempty_file=True))
        if rendered_pdf.suffix.lower() not in PDF_EXTENSIONS:
            issues.append(f"rendered PDF path is not a pdf in {evidence_path}: {rendered_pdf}")

        page_images = split_path_values(raw_values["- page image path list:"])
        if not page_images:
            issues.append(f"page image path list is empty in {evidence_path}")
        for raw_image in page_images:
            image_path = resolve_record_path(raw_image, evidence_path)
            issues.extend(validate_existing_path(image_path, require_nonempty_file=True))
            if image_path.suffix.lower() not in IMAGE_EXTENSIONS:
                issues.append(f"page image path is not an image in {evidence_path}: {image_path}")

    if expected_type in {"thesis-rendered-page", "touched-page-review", "figure-review"} and artifact_extensions and artifact_extensions.issubset({".html", ".txt", ".md"}):
        issues.append(f"rendered evidence points only to html/txt/md artifacts in {evidence_path}")

    return issues


def _validate_effective_font_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    *,
    subject: str,
) -> list[str]:
    issues: list[str] = []
    for prefix in EFFECTIVE_FONT_DETAIL_PREFIXES:
        value = values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"{subject} lacks effective font-chain detail {prefix} in {evidence_path}")

    if issues:
        return issues

    baseline_chain = raw_values.get("- baseline effective font chain:", "")
    actual_chain = raw_values.get("- actual effective font chain:", "")
    slots = raw_values.get("- effective font slots compared:", "")
    alias_verdict = values.get("- theme/default font alias verdict:", "")
    ui_evidence = raw_values.get("- WPS/Word UI font display evidence:", "")
    font_chain_verdict = values.get("- effective font-chain verdict:", "")

    combined_chain = normalize(f"{baseline_chain} {actual_chain} {ui_evidence}").lower()
    for token in EFFECTIVE_FONT_CHAIN_REQUIRED_TOKENS:
        if token not in combined_chain:
            issues.append(
                f"{subject} effective font chain must record {token} inheritance/source in {evidence_path}"
            )

    normalized_slots = normalize(slots).lower()
    for token in EFFECTIVE_FONT_SLOT_TOKENS:
        if token not in normalized_slots:
            issues.append(
                f"{subject} effective font slot comparison must include {token} in {evidence_path}"
            )

    if alias_verdict not in {"pass", "passed"} and not alias_verdict.startswith("passed"):
        issues.append(f"{subject} theme/default font alias verdict must be pass in {evidence_path}")
    if font_chain_verdict not in {"pass", "passed"}:
        issues.append(f"{subject} effective font-chain verdict must be pass in {evidence_path}")

    baseline_norm = normalize(baseline_chain).lower()
    actual_norm = normalize(f"{actual_chain} {ui_evidence}").lower()
    for token in THEME_ALIAS_TOKENS:
        if token in actual_norm and token not in baseline_norm:
            issues.append(
                f"{subject} actual effective font uses theme/default alias not present in the baseline: {evidence_path}"
            )
            break

    return issues


def check_effective_font_evidence_record(
    evidence_path: Path,
    *,
    expected_type: str,
    acceptance_mode: str,
) -> list[str]:
    issues = check_review_evidence_record(
        evidence_path,
        expected_type=expected_type,
        acceptance_mode=acceptance_mode,
    )
    if issues:
        return issues

    values, raw_values, load_issues = _load_review_evidence_values(evidence_path)
    issues.extend(load_issues)
    if issues:
        return issues
    issues.extend(
        _validate_effective_font_fields(
            values,
            raw_values,
            evidence_path,
            subject="font-family baseline evidence",
        )
    )
    return issues


def _load_review_evidence_values(evidence_path: Path) -> tuple[dict[str, str], dict[str, str], list[str]]:
    try:
        lines = read_lines(evidence_path)
    except UnicodeDecodeError as exc:
        return {}, {}, [f"evidence record is not valid UTF-8: {evidence_path} ({exc})"]
    values: dict[str, str] = {}
    raw_values: dict[str, str] = {}
    for prefix in REVIEW_EVIDENCE_SCHEMA["single_prefixes"]:
        matched = find_lines_with_prefix(lines, prefix)
        if len(matched) == 1:
            values[prefix] = parse_line_value(matched[0])
            raw_values[prefix] = raw_line_value(matched[0])
    extra_prefixes = (
        list(STRICT_KEYWORD_RUN_SPLIT_PREFIXES)
        + list(STRICT_SURFACE_GEOMETRY_PREFIXES)
        + list(STRICT_SURFACE_PARAGRAPH_TYPOGRAPHY_PREFIXES)
        + list(STRICT_TOC_VISUAL_GEOMETRY_PREFIXES)
        + list(STRICT_TOC_PARAGRAPH_TYPOGRAPHY_PREFIXES)
        + list(STRICT_TOC_RIGHT_TAB_PAGE_COLUMN_PREFIXES)
        + list(STRICT_WHOLE_DOCUMENT_PAGINATION_PREFIXES)
        + list(STRICT_COVER_MEDIA_PREFIXES)
        + list(STRICT_FRONT_MATTER_HARD_FIELD_PREFIXES)
        + list(STRICT_HEADER_FULL_DISPLAY_PREFIXES)
        + list(STRICT_REFERENCES_ENTRIES_FONT_SIZE_PREFIXES)
        + list(STRICT_ACKNOWLEDGEMENT_TITLE_STYLE_PREFIXES)
        + list(STRICT_FOOTER_PAGE_NUMBER_FONT_SIZE_PREFIXES)
    )
    for prefix in extra_prefixes:
        matched = find_lines_with_prefix(lines, prefix)
        if len(matched) == 1:
            values[prefix] = parse_line_value(matched[0])
            raw_values[prefix] = raw_line_value(matched[0])
        elif prefix not in values:
            values[prefix] = ""
            raw_values[prefix] = ""
    return values, raw_values, []


def _is_pass_value(value: str) -> bool:
    lowered = normalize(value).lower()
    return lowered == "yes" or lowered.startswith("pass")


def _structural_figure_geometry_is_not_applicable(
    values: dict[str, str],
    raw_values: dict[str, str],
) -> bool:
    structural_prefixes = (
        STRUCTURAL_FIGURE_PATH_PREFIXES
        + STRUCTURAL_FIGURE_VERDICT_PREFIXES
        + STRUCTURAL_FIGURE_CONFIRM_PREFIXES
    )
    present_values = [
        normalize(raw_values.get(prefix, "") or values.get(prefix, "")).lower()
        for prefix in structural_prefixes
        if normalize(raw_values.get(prefix, "") or values.get(prefix, ""))
    ]
    if not present_values:
        return False
    combined = " ".join(present_values)
    if "no structural figure" not in combined:
        return False
    return all("not-applicable-with-reason" in value for value in present_values)


def _validate_structural_figure_geometry_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
) -> list[str]:
    if _structural_figure_geometry_is_not_applicable(values, raw_values):
        return []
    issues: list[str] = []
    for prefix in STRUCTURAL_FIGURE_PATH_PREFIXES:
        value = values.get(prefix, "")
        lowered = normalize(value).lower()
        if not is_explicit(value) or value in EXPLICIT_VALUES or lowered in {"none", "not-applicable"} or "not-applicable" in lowered:
            issues.append(f"structural figure evidence lacks required geometry path field {prefix} in {evidence_path}")
            continue
        if prefix == "- structural figure geometry validation report path:":
            for raw_path in split_path_values(raw_values.get(prefix, "")):
                resolved = resolve_record_path(raw_path, evidence_path)
                issues.extend(validate_existing_path(resolved, require_nonempty_file=True))
                if not resolved.exists():
                    continue
                try:
                    report = json.loads(resolved.read_text(encoding="utf-8"))
                except Exception as exc:
                    issues.append(f"structural figure geometry report is not valid JSON in {evidence_path}: {resolved} ({exc})")
                    continue
                if report.get("schema") != "graduation-project-builder.structural-figure-geometry.v1":
                    issues.append(f"structural figure geometry report schema mismatch in {evidence_path}: {resolved}")
                if report.get("verdict") != "pass":
                    issues.append(f"structural figure geometry report verdict is not pass in {evidence_path}: {resolved}")
                if report.get("issues") not in ([], None):
                    issues.append(f"structural figure geometry report still lists issues in {evidence_path}: {resolved}")
                vertices = report.get("vertices")
                if not isinstance(vertices, list) or not vertices:
                    issues.append(f"structural figure geometry report lacks vertex bbox map in {evidence_path}: {resolved}")
                else:
                    for item in vertices:
                        if not isinstance(item, dict) or not isinstance(item.get("bbox"), dict):
                            issues.append(f"structural figure geometry report vertex lacks bbox object in {evidence_path}: {resolved}")
                            break
                family = str(report.get("family") or "").lower()
                if family and family != "er":
                    edge_count = report.get("edge_count")
                    line_segment_count = report.get("line_segment_count")
                    if not isinstance(edge_count, int) or not isinstance(line_segment_count, int):
                        issues.append(f"non-ER structural figure geometry report lacks connector counts in {evidence_path}: {resolved}")
                    elif edge_count + line_segment_count <= 0:
                        issues.append(f"non-ER structural figure geometry report has no connector geometry in {evidence_path}: {resolved}")
        elif prefix in {
            "- structural inserted-scale geometry evidence path:",
            "- structural inserted-scale collision evidence path:",
            "- structural dense-zone crop evidence paths:",
        }:
            for raw_path in split_path_values(raw_values.get(prefix, "")):
                resolved = resolve_record_path(raw_path, evidence_path)
                issues.extend(validate_existing_path(resolved, require_nonempty_file=True))
    for prefix in STRUCTURAL_FIGURE_VERDICT_PREFIXES:
        if not _is_pass_value(values.get(prefix, "")):
            issues.append(f"structural figure evidence must pass {prefix} in {evidence_path}")
    for prefix in STRUCTURAL_FIGURE_CONFIRM_PREFIXES:
        if not _is_pass_value(values.get(prefix, "")):
            issues.append(f"structural figure evidence must confirm {prefix} in {evidence_path}")
    return issues


def _surface_alias_seen(text: str, required_surface_id: str) -> bool:
    lowered = normalize(text).lower()
    return any(alias.lower() in lowered for alias in STRICT_SURFACE_TARGET_ALIASES[required_surface_id])


def _surface_template_instruction_reconciled(raw_values: dict[str, str]) -> bool:
    """Accept measured surfaces whose template donor is an instruction placeholder.

    Some official school templates contain explanatory placeholder paragraphs for
    references and acknowledgements. The hard metric evidence must still include
    measured template/actual rows, but a pass-shaped reconciliation line from the
    protected-surface checker prevents those placeholders from being treated as
    final thesis style donors.
    """
    metric_prefixes = (
        "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
        "- surface typography baseline/actual:",
        "- surface indentation chars/points baseline/actual:",
        "- surface indentation/tab baseline/actual:",
    )
    has_measured_metrics = any(
        "template" in normalize(raw_values.get(prefix, "")).lower()
        and "actual" in normalize(raw_values.get(prefix, "")).lower()
        for prefix in metric_prefixes
    )
    verdict_text = " ".join(
        raw_values.get(prefix, "")
        for prefix in (
            "- surface paragraph-and-typography verdict:",
            "- surface scale/compression verdict:",
        )
    )
    lowered = normalize(verdict_text).lower()
    return (
        has_measured_metrics
        and lowered.startswith(("pass", "passed"))
        and "reconciled" in lowered
        and "template-instruction" in lowered
        and "sample_self_check" in lowered
    )


def _validate_baseline_source_path(raw_value: str, evidence_path: Path, required_surface_id: str) -> list[str]:
    issues: list[str] = []
    match = re.match(r"^(?P<path>.+?)\s+(?:sha256=|sha256:)\s*(?P<sha>[0-9a-fA-F]{64})\s*$", raw_value.strip())
    if not match:
        return [
            f"surface evidence for {required_surface_id} must record baseline source as '<path> sha256=<64-hex>' in {evidence_path}"
        ]
    source_path = resolve_record_path(match.group("path").strip(), evidence_path)
    issues.extend(validate_existing_path(source_path, require_nonempty_file=True))
    if not issues:
        actual_sha = sha256_file(source_path)
        expected_sha = match.group("sha").lower()
        if actual_sha.lower() != expected_sha:
            issues.append(
                f"surface evidence for {required_surface_id} baseline source sha256 does not match path in {evidence_path}"
            )
    return issues


def _validate_rendered_region_images(raw_value: str, evidence_path: Path, required_surface_id: str) -> list[str]:
    issues: list[str] = []
    image_values = split_path_values(raw_value)
    if not image_values:
        return [f"surface evidence for {required_surface_id} must name rendered region image path(s) in {evidence_path}"]
    for raw_image in image_values:
        image_path = resolve_record_path(raw_image, evidence_path)
        issues.extend(validate_existing_path(image_path, require_nonempty_file=True))
        if image_path.suffix.lower() not in IMAGE_EXTENSIONS:
            issues.append(
                f"surface evidence for {required_surface_id} rendered region path is not an image in {evidence_path}: {image_path}"
            )
    return issues


def _validate_surface_geometry_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    subject: str,
) -> list[str]:
    issues: list[str] = []
    combined_geometry_values: list[str] = []
    for prefix in STRICT_SURFACE_GEOMETRY_PREFIXES:
        value = values.get(prefix, "")
        raw_value = raw_values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"{subject} lacks rendered numeric surface-geometry detail {prefix} in {evidence_path}")
        if prefix not in STRICT_TOC_VISUAL_GEOMETRY_IMAGE_PREFIXES and prefix not in STRICT_SURFACE_GEOMETRY_IMAGE_PREFIXES:
            combined_geometry_values.append(f"{prefix} {raw_value}")

    for prefix in STRICT_SURFACE_GEOMETRY_IMAGE_PREFIXES:
        issues.extend(
            _validate_rendered_region_images(
                raw_values.get(prefix, ""),
                evidence_path,
                subject,
            )
        )

    template_images = {
        resolve_record_path(raw_image, evidence_path)
        for raw_image in split_path_values(raw_values.get("- template rendered region image path:", ""))
    }
    actual_images = {
        resolve_record_path(raw_image, evidence_path)
        for raw_image in split_path_values(raw_values.get("- actual rendered region image path:", ""))
    }
    if template_images and actual_images and template_images.intersection(actual_images):
        issues.append(f"{subject} must use distinct template and actual rendered region images in {evidence_path}")

    geometry_text = normalize(" ".join(combined_geometry_values)).lower()
    if not issues:
        for token in SURFACE_GEOMETRY_REQUIRED_TOKENS:
            if token not in geometry_text:
                issues.append(f"{subject} surface geometry must record {token} in {evidence_path}")
        if contains_any(geometry_text, SURFACE_GEOMETRY_FAIL_TOKENS):
            issues.append(f"{subject} surface geometry indicates unresolved or generic geometry evidence in {evidence_path}")

    for prefix in STRICT_SURFACE_GEOMETRY_NUMERIC_PREFIXES:
        raw_value = raw_values.get(prefix, "")
        normalized_value = normalize(raw_value).lower()
        measurements = NUMERIC_MEASUREMENT_RE.findall(normalized_value)
        has_baseline = "template" in normalized_value or "baseline" in normalized_value
        has_actual = "actual" in normalized_value or "target" in normalized_value
        if len(measurements) < 2 or not (has_baseline and has_actual):
            issues.append(
                f"{subject} must include numeric baseline/actual rendered surface geometry measurement for {prefix} in {evidence_path}"
            )

    geometry_verdict = values.get("- surface geometry verdict:", "")
    if geometry_verdict not in {"pass", "passed"} and not geometry_verdict.startswith("passed"):
        issues.append(f"{subject} must include a pass rendered surface geometry verdict in {evidence_path}")
    blank_verdict = values.get("- surface blank crop verdict:", "")
    if blank_verdict not in {"pass", "passed"} and not blank_verdict.startswith("passed"):
        issues.append(f"{subject} must include a pass surface blank crop verdict in {evidence_path}")
    if "none" in normalize(raw_values.get("- surface content bbox baseline/actual:", "")).lower():
        issues.append(f"{subject} must include nonblank template and actual surface content bboxes in {evidence_path}")
    return issues


def _validate_surface_paragraph_typography_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    subject: str,
) -> list[str]:
    issues: list[str] = []
    combined_detail_values: list[str] = []
    for prefix in STRICT_SURFACE_PARAGRAPH_TYPOGRAPHY_PREFIXES:
        value = values.get(prefix, "")
        raw_value = raw_values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"{subject} lacks WPS/Word paragraph-dialog or typography hard field {prefix} in {evidence_path}")
        combined_detail_values.append(raw_value)

    detail_text = normalize(" ".join(combined_detail_values)).lower()
    if not issues:
        for token in SURFACE_PARAGRAPH_TYPOGRAPHY_REQUIRED_TOKENS:
            if token not in detail_text:
                issues.append(f"{subject} paragraph-dialog/typography evidence must record {token} in {evidence_path}")
        if contains_any(detail_text, SURFACE_PARAGRAPH_TYPOGRAPHY_FAIL_TOKENS):
            issues.append(f"{subject} paragraph-dialog/typography evidence indicates unresolved or generic style binding in {evidence_path}")

    for prefix in SURFACE_PARAGRAPH_TYPOGRAPHY_NUMERIC_PREFIXES:
        raw_value = raw_values.get(prefix, "")
        normalized_value = normalize(raw_value).lower()
        measurements = NUMERIC_MEASUREMENT_RE.findall(normalized_value)
        has_baseline = "template" in normalized_value or "baseline" in normalized_value
        has_actual = "actual" in normalized_value or "target" in normalized_value
        if len(measurements) < 2 or not (has_baseline and has_actual):
            issues.append(
                f"{subject} must include numeric baseline/actual WPS/Word paragraph-dialog or typography measurement for {prefix} in {evidence_path}"
            )

    paragraph_verdict = values.get("- surface paragraph-and-typography verdict:", "")
    if paragraph_verdict not in {"pass", "passed"} and not paragraph_verdict.startswith("passed"):
        issues.append(f"{subject} must include a pass surface paragraph-and-typography verdict in {evidence_path}")
    scale_verdict = values.get("- surface scale/compression verdict:", "")
    if scale_verdict not in {"pass", "passed"} and not scale_verdict.startswith("passed"):
        issues.append(f"{subject} must include a pass surface scale/compression verdict in {evidence_path}")
    return issues


def _validate_required_prefix_group(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
    prefixes: list[str],
    *,
    label: str,
    require_pass_verdict_prefix: str | None = None,
    required_tokens: tuple[str, ...] = (),
) -> list[str]:
    issues: list[str] = []
    combined_values: list[str] = []
    for prefix in prefixes:
        value = values.get(prefix, "")
        raw_value = raw_values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(
                f"surface evidence for {required_surface_id} lacks {label} hard field {prefix} in {evidence_path}"
            )
        combined_values.append(f"{prefix} {raw_value}")
    combined = normalize(" ".join(combined_values)).lower()
    for token in required_tokens:
        if token not in combined:
            issues.append(
                f"surface evidence for {required_surface_id} {label} must record {token} in {evidence_path}"
            )
    if require_pass_verdict_prefix:
        verdict = values.get(require_pass_verdict_prefix, "")
        if verdict not in {"pass", "passed"} and not verdict.startswith("passed"):
            issues.append(
                f"surface evidence for {required_surface_id} must include pass for {require_pass_verdict_prefix} in {evidence_path}"
            )
    return issues


def _validate_surface_defect_hard_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    if required_surface_id == "cover_style":
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_COVER_MEDIA_PREFIXES,
                label="cover media/icon binding",
                require_pass_verdict_prefix="- cover media/icon binding verdict:",
                required_tokens=("rid", "media"),
            )
        )
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_FRONT_MATTER_HARD_FIELD_PREFIXES,
                label="cover/front-matter hard-field typography",
                require_pass_verdict_prefix="- front-matter hard-field verdict:",
                required_tokens=("template", "actual", "font", "size", "spacing"),
            )
        )
    if required_surface_id in {
        "declaration_or_title_front_matter",
        "zh_abstract_title",
        "zh_abstract_body",
        "zh_keyword_line",
        "en_abstract_title",
        "en_abstract_body",
        "en_keyword_line",
        "toc_title",
        "toc_entries",
    }:
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_FRONT_MATTER_HARD_FIELD_PREFIXES,
                label="front-matter hard-field typography",
                require_pass_verdict_prefix="- front-matter hard-field verdict:",
                required_tokens=("template", "actual", "font", "size", "spacing"),
            )
        )
    if required_surface_id == "header":
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_HEADER_FULL_DISPLAY_PREFIXES,
                label="header full-display string",
                require_pass_verdict_prefix="- header full-display string verdict:",
                required_tokens=("expected", "observed"),
            )
        )
    if required_surface_id == "references_entries":
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_REFERENCES_ENTRIES_FONT_SIZE_PREFIXES,
                label="references entries font size",
                require_pass_verdict_prefix="- references entries font-size verdict:",
                required_tokens=("template", "actual", "size"),
            )
        )
    if required_surface_id == "acknowledgement_title":
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_ACKNOWLEDGEMENT_TITLE_STYLE_PREFIXES,
                label="acknowledgement title style",
                require_pass_verdict_prefix="- acknowledgement title paragraph style verdict:",
                required_tokens=("template", "actual", "style"),
            )
        )
    if required_surface_id in {"footer", "page_numbers"}:
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_FOOTER_PAGE_NUMBER_FONT_SIZE_PREFIXES,
                label="footer page-number font size",
                require_pass_verdict_prefix="- footer page-number font-size verdict:",
                required_tokens=("template", "actual", "size", "page"),
            )
        )
    return issues


def _split_template_actual_metric_segments(raw_value: str) -> tuple[str, str] | None:
    match = re.search(
        r"(?:template|baseline)\s+(?P<template>.*?);\s*(?:actual|target)\s+(?P<actual>.*)",
        raw_value,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    return match.group("template"), match.group("actual")


def _surface_metric_map(segment: str) -> dict[str, str]:
    text = normalize(segment)
    text = re.sub(r"first\s*line", "firstline", text, flags=re.IGNORECASE)
    text = re.sub(r"line\s*mode", "linemode", text, flags=re.IGNORECASE)
    text = re.sub(r"line\s*value", "linevalue", text, flags=re.IGNORECASE)
    metrics: dict[str, str] = {}
    for match in re.finditer(
        r"\b(?P<key>style|font|size|weight|alignment|outline|left-x|centerline|left|right|firstline|hanging|tab|before|after|linemode|linevalue|line|x|y|w|h)\s*=?\s*(?P<value>Times\s+New\s+Roman|[-+]?\d+(?:\.\d+)?(?:pt|px|%)?|[\w\u4e00-\u9fff._-]+)",
        text,
        flags=re.IGNORECASE,
    ):
        metrics[match.group("key").lower()] = match.group("value").lower()
    return metrics


def _numeric_metric_value(value: str) -> float | None:
    match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _surface_image_sizes(raw_values: dict[str, str]) -> tuple[float | None, float | None]:
    segments = _split_template_actual_metric_segments(
        raw_values.get("- surface crop source image size baseline/actual:", "")
    )
    if segments is None:
        return None, None
    template_metrics = _surface_metric_map(segments[0])
    actual_metrics = _surface_metric_map(segments[1])
    return _numeric_metric_value(template_metrics.get("w", "")), _numeric_metric_value(actual_metrics.get("w", ""))


def _normalized_x_metric_matches(
    template_value: str,
    actual_value: str,
    template_width: float | None,
    actual_width: float | None,
    *,
    tolerance: float = 0.0025,
) -> bool:
    template_x = _numeric_metric_value(template_value)
    actual_x = _numeric_metric_value(actual_value)
    if template_x is None or actual_x is None or not template_width or not actual_width:
        return template_value == actual_value
    return abs((template_x / template_width) - (actual_x / actual_width)) <= tolerance


def _parse_int_list_token(text: str, token: str) -> list[int] | None:
    match = re.search(rf"{re.escape(token)}\s*=\s*\[(?P<body>[^\]]*)\]", text, flags=re.IGNORECASE)
    if not match:
        return None
    body = match.group("body").strip()
    if not body:
        return []
    result: list[int] = []
    for item in re.findall(r"-?\d+", body):
        result.append(int(item))
    return result


def _validate_acknowledgement_title_indent_position(
    raw_values: dict[str, str],
    evidence_path: Path,
) -> list[str]:
    issues: list[str] = []
    if _surface_template_instruction_reconciled(raw_values):
        return issues
    checked_keys: set[str] = set()
    template_width, actual_width = _surface_image_sizes(raw_values)
    prefixes = (
        "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
        "- surface indentation chars/points baseline/actual:",
        "- surface indentation/tab baseline/actual:",
        "- surface bbox baseline/actual:",
        "- surface position baseline/actual:",
    )
    critical_keys = {"alignment", "left", "right", "firstline", "hanging", "tab", "x", "left-x", "centerline"}
    for prefix in prefixes:
        segments = _split_template_actual_metric_segments(raw_values.get(prefix, ""))
        if segments is None:
            continue
        template_metrics = _surface_metric_map(segments[0])
        actual_metrics = _surface_metric_map(segments[1])
        for key in sorted(critical_keys & template_metrics.keys() & actual_metrics.keys()):
            checked_keys.add(key)
            if key in {"x", "left-x", "centerline"}:
                matches = _normalized_x_metric_matches(
                    template_metrics[key],
                    actual_metrics[key],
                    template_width,
                    actual_width,
                )
            else:
                matches = template_metrics[key] == actual_metrics[key]
            if not matches:
                issues.append(
                    "surface evidence for acknowledgement_title has title indentation/position drift "
                    f"for {key} in {prefix} template={template_metrics[key]} actual={actual_metrics[key]} "
                    f"in {evidence_path}"
                )
    if not checked_keys & {"alignment", "left", "firstline", "hanging", "tab", "x", "left-x", "centerline"}:
        issues.append(
            f"surface evidence for acknowledgement_title must expose comparable title alignment/indentation/left-x metrics in {evidence_path}"
        )
    return issues


def _validate_acknowledgement_body_hard_metrics(
    raw_values: dict[str, str],
    evidence_path: Path,
) -> list[str]:
    issues: list[str] = []
    if _surface_template_instruction_reconciled(raw_values):
        return issues
    checked_keys: set[str] = set()
    prefixes = (
        "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
        "- surface typography baseline/actual:",
        "- surface indentation chars/points baseline/actual:",
        "- surface indentation/tab baseline/actual:",
    )
    critical_keys = {
        "style",
        "font",
        "size",
        "weight",
        "alignment",
        "outline",
        "left",
        "right",
        "firstline",
        "hanging",
        "tab",
        "before",
        "after",
        "linemode",
        "linevalue",
        "line",
    }
    required_keys = {"style", "font", "size", "weight", "alignment", "outline", "firstline", "left", "hanging"}
    for prefix in prefixes:
        segments = _split_template_actual_metric_segments(raw_values.get(prefix, ""))
        if segments is None:
            continue
        template_metrics = _surface_metric_map(segments[0])
        actual_metrics = _surface_metric_map(segments[1])
        for key in sorted(critical_keys & template_metrics.keys() & actual_metrics.keys()):
            checked_keys.add(key)
            if template_metrics[key] != actual_metrics[key]:
                issues.append(
                    "surface evidence for acknowledgement_body has body paragraph/typography drift "
                    f"for {key} in {prefix} template={template_metrics[key]} actual={actual_metrics[key]} "
                    f"in {evidence_path}"
                )
    missing_keys = sorted(required_keys - checked_keys)
    if missing_keys:
        issues.append(
            "surface evidence for acknowledgement_body must expose comparable body style, typography, "
            f"alignment, and indentation metrics in {evidence_path}; missing {', '.join(missing_keys)}"
        )
    return issues


def _validate_references_surface_hard_metrics(
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    if _surface_template_instruction_reconciled(raw_values):
        return issues
    checked_keys: set[str] = set()
    reconciled_title_surface = (
        required_surface_id == "references_title"
        and normalize(raw_values.get("- surface paragraph-and-typography verdict:", ""))
        .lower()
        .startswith(("pass", "passed"))
        and "reconciled" in normalize(raw_values.get("- surface paragraph-and-typography verdict:", "")).lower()
    )
    effective_font_chain_passed = (
        required_surface_id == "references_entries"
        and normalize(raw_values.get("- effective font-chain verdict:", "")).lower().startswith(("pass", "passed"))
        and is_explicit(raw_values.get("- effective font slots compared:", ""))
    )
    prefixes = (
        "- surface WPS/Word paragraph-dialog metrics baseline/actual:",
        "- surface typography baseline/actual:",
        "- surface indentation chars/points baseline/actual:",
        "- surface indentation/tab baseline/actual:",
    )
    critical_keys = {"style", "font", "size", "weight", "alignment", "left", "right", "firstline", "hanging", "tab"}
    for prefix in prefixes:
        segments = _split_template_actual_metric_segments(raw_values.get(prefix, ""))
        if segments is None:
            continue
        template_metrics = _surface_metric_map(segments[0])
        actual_metrics = _surface_metric_map(segments[1])
        for key in sorted(critical_keys & template_metrics.keys() & actual_metrics.keys()):
            checked_keys.add(key)
            if key == "font" and effective_font_chain_passed:
                continue
            if reconciled_title_surface:
                continue
            if template_metrics[key] != actual_metrics[key]:
                issues.append(
                    f"surface evidence for {required_surface_id} has reference paragraph/typography drift "
                    f"for {key} in {prefix} template={template_metrics[key]} actual={actual_metrics[key]} "
                    f"in {evidence_path}"
                )
    if not checked_keys & {"firstline", "hanging", "left", "tab", "font", "size", "weight"}:
        issues.append(
            f"surface evidence for {required_surface_id} must expose comparable reference title/entry paragraph and typography metrics in {evidence_path}"
        )
    return issues


def _validate_toc_visual_geometry_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    combined_geometry_values: list[str] = []
    for prefix in STRICT_TOC_VISUAL_GEOMETRY_PREFIXES:
        value = values.get(prefix, "")
        raw_value = raw_values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(
                f"surface evidence for {required_surface_id} lacks TOC visual-geometry detail {prefix} in {evidence_path}"
            )
        if prefix not in STRICT_TOC_VISUAL_GEOMETRY_IMAGE_PREFIXES:
            combined_geometry_values.append(f"{prefix} {raw_value}")

    for prefix in STRICT_TOC_VISUAL_GEOMETRY_IMAGE_PREFIXES:
        issues.extend(
            _validate_rendered_region_images(
                raw_values.get(prefix, ""),
                evidence_path,
                required_surface_id,
            )
        )

    template_images = {
        resolve_record_path(raw_image, evidence_path)
        for raw_image in split_path_values(raw_values.get("- TOC template rendered page/region image path:", ""))
    }
    actual_images = {
        resolve_record_path(raw_image, evidence_path)
        for raw_image in split_path_values(raw_values.get("- TOC actual rendered page/region image path:", ""))
    }
    if template_images and actual_images and template_images.intersection(actual_images):
        issues.append(
            f"surface evidence for {required_surface_id} must use distinct template and actual TOC rendered images in {evidence_path}"
        )

    geometry_text = normalize(" ".join(combined_geometry_values)).lower()
    if not issues:
        if "ink" not in geometry_text:
            issues.append(
                f"surface evidence for {required_surface_id} TOC visual geometry must prove rendered ink-pixel measurement in {evidence_path}"
            )
        if "fixed-proportion" in geometry_text or "synthetic" in geometry_text:
            issues.append(
                f"surface evidence for {required_surface_id} TOC visual geometry must not use fixed-proportion or synthetic row boxes in {evidence_path}"
            )
        for token in TOC_VISUAL_GEOMETRY_REQUIRED_TOKENS:
            if token not in geometry_text:
                issues.append(
                    f"surface evidence for {required_surface_id} TOC visual geometry must record {token} in {evidence_path}"
                )
        if contains_any(geometry_text, TOC_VISUAL_GEOMETRY_FAIL_TOKENS):
            issues.append(
                f"surface evidence for {required_surface_id} TOC visual geometry indicates unresolved geometry drift in {evidence_path}"
            )

    for prefix in TOC_VISUAL_GEOMETRY_NUMERIC_PREFIXES:
        raw_value = raw_values.get(prefix, "")
        normalized_value = normalize(raw_value).lower()
        measurements = NUMERIC_MEASUREMENT_RE.findall(normalized_value)
        has_baseline = "template" in normalized_value or "baseline" in normalized_value
        has_actual = "actual" in normalized_value or "target" in normalized_value
        if len(measurements) < 2 or not (has_baseline and has_actual):
            issues.append(
                f"surface evidence for {required_surface_id} must include numeric baseline/actual TOC geometry measurement for {prefix} in {evidence_path}"
            )

    geometry_verdict = values.get("- TOC visual geometry verdict:", "")
    normalized_geometry_verdict = normalize(geometry_verdict).lower()
    if not re.search(r"\bpass(?:ed)?\b", normalized_geometry_verdict) or contains_any(
        normalized_geometry_verdict,
        {"fail", "failed", "missing", "not checked", "pending", "blocked", "unresolved", "drift"},
    ):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass TOC visual geometry verdict in {evidence_path}"
        )

    return issues


def parse_toc_used_levels(raw_inventory: str) -> set[str]:
    normalized = normalize(raw_inventory).lower()
    return {level for level in TOC_USED_LEVEL_ALLOWED if level in normalized}


def parse_toc_used_level_map(raw_map: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in split_path_values(raw_map.replace(",", ";")):
        text = item.strip()
        if "=" not in text:
            continue
        key, value = text.split("=", 1)
        level = normalize(key).lower()
        if level in TOC_USED_LEVEL_ALLOWED:
            result[level] = value.strip()
    return result


def _validate_toc_used_level_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    raw_inventory = raw_values.get("- TOC used-level inventory:", "")
    raw_map = raw_values.get("- TOC used-level evidence map:", "")
    normalized_inventory = normalize(raw_inventory).lower()
    if any(token in normalized_inventory for token in TOC_USED_LEVEL_CONDITIONAL_TOKENS):
        issues.append(
            f"surface evidence for {required_surface_id} TOC used-level inventory must be explicit, not conditional, in {evidence_path}"
        )
    used_levels = parse_toc_used_levels(raw_inventory)
    if "title" not in used_levels or not any(level.startswith("level") for level in used_levels):
        issues.append(
            f"surface evidence for {required_surface_id} TOC used-level inventory must name title and each actual TOC level in {evidence_path}"
        )
    evidence_map = parse_toc_used_level_map(raw_map)
    for level in sorted(used_levels, key=TOC_USED_LEVEL_ALLOWED.index):
        mapped = evidence_map.get(level, "")
        mapped_norm = normalize(mapped).lower()
        if not mapped:
            issues.append(
                f"surface evidence for {required_surface_id} TOC used-level evidence map missing {level} in {evidence_path}"
            )
            continue
        if mapped_norm in {normalize(required_surface_id).lower(), "toc", "tocentries"}:
            issues.append(
                f"surface evidence for {required_surface_id} TOC used-level evidence map must point {level} to a level-specific row or linked evidence, not a generic surface id, in {evidence_path}"
            )
        if level not in mapped_norm:
            issues.append(
                f"surface evidence for {required_surface_id} TOC used-level evidence map value for {level} must include the level id in {evidence_path}"
            )
    extra_levels = sorted(set(evidence_map) - used_levels, key=TOC_USED_LEVEL_ALLOWED.index)
    if extra_levels:
        issues.append(
            f"surface evidence for {required_surface_id} TOC used-level evidence map names levels absent from the used-level inventory: {', '.join(extra_levels)} in {evidence_path}"
        )
    return issues


def _validate_toc_paragraph_typography_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    combined_values: list[str] = []
    for prefix in STRICT_TOC_PARAGRAPH_TYPOGRAPHY_PREFIXES:
        value = values.get(prefix, "")
        raw_value = raw_values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(
                f"surface evidence for {required_surface_id} lacks TOC paragraph/typography detail {prefix} in {evidence_path}"
            )
        combined_values.append(f"{prefix} {raw_value}")

    detail_text = normalize(" ".join(combined_values)).lower()
    if not issues:
        for token in TOC_PARAGRAPH_TYPOGRAPHY_REQUIRED_TOKENS:
            if token not in detail_text:
                issues.append(
                    f"surface evidence for {required_surface_id} TOC paragraph/typography must record {token} in {evidence_path}"
                )
        if contains_any(detail_text, TOC_PARAGRAPH_TYPOGRAPHY_FAIL_TOKENS):
            issues.append(
                f"surface evidence for {required_surface_id} TOC paragraph/typography indicates unresolved paragraph or typography drift in {evidence_path}"
            )

    for prefix in TOC_PARAGRAPH_TYPOGRAPHY_NUMERIC_PREFIXES:
        raw_value = raw_values.get(prefix, "")
        normalized_value = normalize(raw_value).lower()
        measurements = NUMERIC_MEASUREMENT_RE.findall(normalized_value)
        has_baseline = "template" in normalized_value or "baseline" in normalized_value
        has_actual = "actual" in normalized_value or "target" in normalized_value
        if len(measurements) < 2 or not (has_baseline and has_actual):
            issues.append(
                f"surface evidence for {required_surface_id} must include numeric baseline/actual TOC paragraph/typography measurement for {prefix} in {evidence_path}"
            )

    paragraph_verdict = values.get("- TOC paragraph-and-typography verdict:", "")
    if paragraph_verdict not in {"pass", "passed"} and not paragraph_verdict.startswith("passed"):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass TOC paragraph-and-typography verdict in {evidence_path}"
        )
    scale_verdict = values.get("- TOC scale/compression verdict:", "")
    if scale_verdict not in {"pass", "passed"} and not scale_verdict.startswith("passed"):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass TOC scale/compression verdict in {evidence_path}"
        )
    run_typography_verdict = values.get("- TOC run typography verdict:", "")
    if run_typography_verdict not in {"pass", "passed"} and not run_typography_verdict.startswith("passed"):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass TOC run typography verdict in {evidence_path}"
        )
    if required_surface_id in {"toc_dotted_leaders", "toc_page_number_column"}:
        issues.extend(
            _validate_required_prefix_group(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
                STRICT_TOC_RIGHT_TAB_PAGE_COLUMN_PREFIXES,
                label="TOC right-tab/page-number-column semantics",
                require_pass_verdict_prefix="- TOC per-entry right-tab/page-number verdict:",
                required_tokens=("right", "tab", "page-number", "leader"),
            )
        )
    issues.extend(_validate_toc_used_level_fields(values, raw_values, evidence_path, required_surface_id))

    return issues


def _validate_whole_document_pagination_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    combined = normalize("\n".join(raw_values.get(prefix, "") for prefix in STRICT_WHOLE_DOCUMENT_PAGINATION_PREFIXES)).lower()
    blank_scan = normalize(raw_values.get("- blank/near-empty page scan verdict:", "")).lower()
    text_declares_allowed_pass = (
        raw_values.get("- fatal pagination topology differences:", "").strip() == "[]"
        and values.get("- DOCX pagination structure verdict:", "").startswith(("pass", "passed"))
        and values.get("- whole-document pagination verdict:", "").startswith(("pass", "passed"))
    )

    for prefix in STRICT_WHOLE_DOCUMENT_PAGINATION_PREFIXES:
        value = values.get(prefix, "")
        if prefix in {
            "- fatal pagination topology differences:",
            "- allowed content-growth pagination differences:",
            "- all pagination topology differences:",
        } and value == "[]":
            continue
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(
                f"surface evidence for {required_surface_id} lacks whole-document pagination field {prefix} in {evidence_path}"
            )

    if contains_any(combined, WHOLE_DOCUMENT_PAGINATION_FAIL_TOKENS) and not text_declares_allowed_pass:
        issues.append(
            f"surface evidence for {required_surface_id} indicates unresolved whole-document pagination drift in {evidence_path}"
        )
    for token in (
        "template_blank_pages=",
        "actual_blank_pages=",
        "unexpected_blank_pages=[]",
        "actual_near_empty_pages=",
        "unexpected_near_empty_pages=[]",
        "rendered_ink_ratio",
    ):
        if token not in blank_scan:
            issues.append(
                f"surface evidence for {required_surface_id} blank/near-empty page scan must record {token} in {evidence_path}"
            )

    for token in WHOLE_DOCUMENT_PAGINATION_REQUIRED_TOKENS:
        if token not in combined:
            issues.append(
                f"surface evidence for {required_surface_id} whole-document pagination must record {token} evidence in {evidence_path}"
            )
    if values.get("- DOCX pagination structure schema:", "") != "graduation-project-builder.docx-pagination-structure.v1":
        issues.append(
            f"surface evidence for {required_surface_id} must cite canonical DOCX pagination structure schema v1 in {evidence_path}"
        )
    if values.get("- DOCX pagination structure generator:", "") != "inspect_docx_pagination_structure.py":
        issues.append(
            f"surface evidence for {required_surface_id} must cite inspect_docx_pagination_structure.py as DOCX pagination generator in {evidence_path}"
        )
    structure_path = resolve_record_path(raw_values.get("- DOCX pagination structure evidence path:", ""), evidence_path)
    try:
        payload = json.loads(structure_path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(
            f"surface evidence for {required_surface_id} must point to readable DOCX pagination JSON in {evidence_path}: {structure_path} ({exc})"
        )
        payload = {}
    if isinstance(payload, dict):
        if payload.get("schema") != "graduation-project-builder.docx-pagination-structure.v1":
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON has wrong schema: {structure_path}"
            )
        if payload.get("generator_script") != "inspect_docx_pagination_structure.py":
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON has wrong generator_script: {structure_path}"
            )
        reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
        payload_final_path = payload.get("final_docx_path")
        if payload_final_path:
            resolved_payload_final = Path(str(payload_final_path))
            if not resolved_payload_final.is_absolute():
                resolved_payload_final = (structure_path.parent / resolved_payload_final).resolve()
            else:
                resolved_payload_final = resolved_payload_final.resolve()
            if resolved_payload_final != reviewed_output.resolve():
                issues.append(
                    f"surface evidence for {required_surface_id} DOCX pagination JSON targets {resolved_payload_final} instead of reviewed output {reviewed_output.resolve()}: {structure_path}"
                )
        else:
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON must include final_docx_path: {structure_path}"
            )
        payload_final_sha = str(payload.get("final_docx_sha256", ""))
        if not re.fullmatch(r"[0-9a-fA-F]{64}", payload_final_sha):
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON must include a 64-hex final_docx_sha256: {structure_path}"
            )
        elif reviewed_output.exists() and payload_final_sha.lower() != sha256_file(reviewed_output).lower():
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON final_docx_sha256 does not match reviewed output: {structure_path}"
            )
        fatal_differences = payload.get("fatal_differences")
        if fatal_differences is None:
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON must include fatal_differences: {structure_path}"
            )
        elif fatal_differences:
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON contains fatal differences: {fatal_differences}"
            )
        for key in ("section_count_verdict", "header_footer_reference_verdict", "page_number_restart_verdict"):
            if str(payload.get(key, "")).strip().lower() not in {"pass", "passed"}:
                issues.append(
                    f"surface evidence for {required_surface_id} DOCX pagination JSON must set {key}=pass: {structure_path}"
                )
        json_blank_scan = normalize(str(payload.get("blank_near_empty_page_scan_verdict", ""))).lower()
        raw_unexpected_near_empty = _parse_int_list_token(json_blank_scan, "raw_unexpected_near_empty_pages")
        unexpected_near_empty = _parse_int_list_token(json_blank_scan, "unexpected_near_empty_pages")
        allowed_near_empty = _parse_int_list_token(json_blank_scan, "allowed_content_growth_near_empty_pages") or []
        allowlist_near_empty = _parse_int_list_token(json_blank_scan, "near_empty_page_explicit_allowlist") or []
        if raw_unexpected_near_empty:
            unapproved = [
                page
                for page in raw_unexpected_near_empty
                if page not in allowed_near_empty or page not in allowlist_near_empty
            ]
            if unapproved:
                issues.append(
                    f"surface evidence for {required_surface_id} DOCX pagination JSON has unapproved raw unexpected near-empty pages {unapproved}: {structure_path}"
                )
        if unexpected_near_empty:
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON contains unexpected near-empty pages {unexpected_near_empty}: {structure_path}"
            )
        tail_block_map = str(payload.get("tail_block_opener_page_map", "")).lower()
        for token in (
            "references previous content physical page=",
            "references physical page=",
            "references_page_found=yes",
            "references_fresh_page_verdict=pass",
            "references_prior_block_separation_verdict=pass",
            "references_opener_owner_evidence=",
            "tail-block.pagination-contract",
        ):
            if token not in tail_block_map:
                issues.append(
                    f"surface evidence for {required_surface_id} DOCX pagination JSON tail_block_opener_page_map lacks {token}: {structure_path}"
                )
        if (
            "references previous content physical page=missing" in tail_block_map
            or "references physical page=missing" in tail_block_map
            or "references_fresh_page_verdict=fail" in tail_block_map
            or "references_prior_block_separation_verdict=fail" in tail_block_map
        ):
            issues.append(
                f"surface evidence for {required_surface_id} DOCX pagination JSON records lost references pagination: {structure_path}"
            )

    numeric_fields = (
        "- rendered page count baseline/actual:",
        "- section count baseline/actual:",
        "- section boundary map baseline/actual:",
        "- page-number format/restart map baseline/actual:",
    )
    for prefix in numeric_fields:
        if not any(char.isdigit() for char in raw_values.get(prefix, "")):
            issues.append(
                f"surface evidence for {required_surface_id} must include numeric baseline/actual pagination data for {prefix} in {evidence_path}"
            )
    raw_tail_map = raw_values.get("- tail-block opener page map:", "").lower()
    for token in (
        "references previous content physical page=",
        "references physical page=",
        "references_page_found=yes",
        "references_fresh_page_verdict=pass",
        "references_prior_block_separation_verdict=pass",
        "references_opener_owner_evidence=",
        "tail-block.pagination-contract",
    ):
        if token not in raw_tail_map:
            issues.append(
                f"surface evidence for {required_surface_id} tail-block opener page map must record {token} in {evidence_path}"
            )

    structure_verdict = values.get("- DOCX pagination structure verdict:", "")
    if not structure_verdict.startswith(("pass", "passed")):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass DOCX pagination structure verdict in {evidence_path}"
        )
    for prefix in (
        "- section count verdict:",
        "- header/footer reference verdict:",
        "- page-number restart verdict:",
    ):
        if values.get(prefix, "") not in {"pass", "passed"}:
            issues.append(
                f"surface evidence for {required_surface_id} must include pass for {prefix} in {evidence_path}"
            )
    verdict = values.get("- whole-document pagination verdict:", "")
    if not verdict.startswith(("pass", "passed")):
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass whole-document pagination verdict in {evidence_path}"
        )
    return issues


DOCX_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
DOCX_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOCX_V = "{urn:schemas-microsoft-com:vml}"


def _docx_normalize_font_name(value: str) -> str:
    compact = normalize(value).lower()
    if compact in {"timesnewroman", "timesnewromanpsmt"}:
        return "Times New Roman"
    if compact in {"simsun", "songti", "\u5b8b\u4f53", "宋体"}:
        return "SimSun"
    return value


def _load_docx_document_root(docx_path: Path) -> tuple[ET.Element | None, list[str]]:
    if docx_path.suffix.lower() != ".docx":
        return None, [f"reviewed output must be a DOCX for protected surface artifact scans, found: {docx_path}"]
    if not docx_path.exists():
        return None, [f"reviewed output DOCX does not exist for protected surface artifact scan: {docx_path}"]
    try:
        with zipfile.ZipFile(docx_path) as zf:
            return ET.fromstring(zf.read("word/document.xml")), []
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return None, [f"reviewed output DOCX artifact scan failed for {docx_path}: {exc}"]


def _docx_visible_text(root: ET.Element) -> str:
    return "\n".join("".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)) for paragraph in root.findall(".//w:p", DOCX_NS))


def _validate_cover_placeholder_residue_docx(docx_path: Path) -> list[str]:
    root, issues = _load_docx_document_root(docx_path)
    if root is None:
        return issues
    text = _docx_visible_text(root)
    residue_tokens = (
        "中文题目",
        "英文题目",
        "须删除本行",
        "须替换为本论文名称",
        "字体字号",
        "会务组织系统的设计与实现",
        "Conference Organization System Based on Wechat Public Platform",
    )
    found = [token for token in residue_tokens if token in text]
    if found:
        issues.append("cover/front-matter template placeholder residue remains in reviewed DOCX: " + "; ".join(found))
    return issues


def _docx_has_vml_image_payload(node: ET.Element) -> bool:
    return node.find(".//v:imagedata", DOCX_NS) is not None


def _validate_front_matter_vml_residue_docx(docx_path: Path) -> list[str]:
    root, issues = _load_docx_document_root(docx_path)
    if root is None:
        return issues
    residuals: list[str] = []
    for paragraph_index, paragraph in enumerate(root.findall(".//w:body/w:p", DOCX_NS), start=1):
        if _docx_has_vml_image_payload(paragraph):
            continue
        for element in paragraph.iter():
            if element.tag in {DOCX_V + "line", DOCX_V + "shape"}:
                style = " ".join(str(value).lower() for value in element.attrib.values())
                if element.tag == DOCX_V + "line" or "arrow" in style or "callout" in style:
                    residuals.append(f"paragraph {paragraph_index} {element.tag.split('}')[-1]}")
                    break
            if element.tag.endswith("}stroke"):
                style = " ".join(str(value).lower() for value in element.attrib.values())
                if "arrow" in style:
                    residuals.append(f"paragraph {paragraph_index} stroke-arrow")
                    break
    if residuals:
        issues.append("front-matter VML arrow/line/callout residue remains in reviewed DOCX: " + "; ".join(residuals[:8]))
    return issues


def _docx_paragraph_style_fonts(root: ET.Element, docx_path: Path) -> dict[str, dict[str, str]]:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError):
        return {}
    result: dict[str, dict[str, str]] = {}
    for style in styles_root.findall("w:style", DOCX_NS):
        if style.attrib.get(DOCX_W + "type") != "paragraph":
            continue
        style_id = style.attrib.get(DOCX_W + "styleId", "")
        if not style_id:
            continue
        rfonts = style.find("w:rPr/w:rFonts", DOCX_NS)
        signature = {"ascii": "", "hAnsi": "", "cs": "", "eastAsia": ""}
        if rfonts is not None:
            for key in signature:
                signature[key] = _docx_normalize_font_name(rfonts.attrib.get(DOCX_W + key, ""))
        result[style_id] = signature
    return result


def _docx_run_font_slots(run: ET.Element, style_fonts: dict[str, str]) -> dict[str, str]:
    signature = {"ascii": "", "hAnsi": "", "cs": "", "eastAsia": ""}
    rpr = run.find("w:rPr", DOCX_NS)
    if rpr is not None:
        rfonts = rpr.find("w:rFonts", DOCX_NS)
        if rfonts is not None:
            for key in signature:
                signature[key] = _docx_normalize_font_name(rfonts.attrib.get(DOCX_W + key, ""))
    for key, value in style_fonts.items():
        if not signature.get(key):
            signature[key] = value
    return signature


def _docx_has_ascii_alnum(text: str) -> bool:
    return any(char.isascii() and char.isalnum() for char in text)


def _docx_has_ascii_alpha(text: str) -> bool:
    return any(char.isascii() and char.isalpha() for char in text)


def _zh_abstract_body_paragraphs(root: ET.Element) -> list[ET.Element]:
    paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    title_seen = False
    result: list[ET.Element] = []
    for paragraph in paragraphs:
        visible = "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()
        compact = re.sub(r"[\s\u25a1]+", "", visible).lower()
        if not title_seen:
            if compact == "摘要":
                title_seen = True
            continue
        if compact.startswith("关键词") or compact.startswith("關鍵詞"):
            break
        if visible:
            result.append(paragraph)
    return result


def _compact_body_text(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "").lower()


def _body_text_paragraphs(root: ET.Element) -> list[ET.Element]:
    paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    toc_seen = False
    main_body_started = False
    reference_zone_started = False
    result: list[ET.Element] = []
    terminal_markers = {"参考文献", "references", "bibliography", "致谢", "acknowledgement", "acknowledgment", "附录", "appendix"}
    for paragraph in paragraphs:
        visible = "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()
        if not visible:
            continue
        compact = _compact_body_text(visible)
        if not toc_seen:
            if compact in {"目录", "contents", "tableofcontents"}:
                toc_seen = True
            continue
        if compact in { _compact_body_text(item) for item in terminal_markers }:
            reference_zone_started = True
            continue
        if reference_zone_started:
            continue
        if re.match(r"^(?:第[0-9一二三四五六七八九十]+章|chapter\d+|\d+(?:\.\d+){0,3})", compact, re.IGNORECASE):
            main_body_started = True
            continue
        if not main_body_started:
            continue
        if paragraph.find(".//w:drawing", DOCX_NS) is not None or paragraph.find(".//w:pict", DOCX_NS) is not None:
            continue
        result.append(paragraph)
    return result


def _docx_paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()


def _compact_docx_text(text: str) -> str:
    return re.sub(r"[\s\u25a1\u3000]+", "", text or "").lower()


def _is_docx_body_heading(text: str) -> bool:
    stripped = re.sub(r"\s+", " ", text or "").strip()
    compact = _compact_docx_text(stripped)
    return bool(
        re.match(r"^(?:\d{1,2}(?:\.\d{1,2}){0,3})(?:\s+|\u3000+|[^\d.])", stripped)
        or re.match(r"^\d{1,2}(?:\.\d{1,2}){0,3}[\u4e00-\u9fffA-Za-z]", compact)
        or re.match(r"^\u7b2c[0-9\u4e00-\u9fff]+\u7ae0", compact)
        or stripped.lower().startswith("chapter ")
    )


def _is_docx_body_heading_paragraph(paragraph: ET.Element, text: str) -> bool:
    if _is_docx_body_heading(text):
        return True
    if _is_docx_caption_or_table_marker(text) or _is_docx_non_body_terminal(text):
        return False
    if paragraph.find(".//w:tab", DOCX_NS) is not None or paragraph.find("./w:pPr/w:tabs", DOCX_NS) is not None:
        return False

    style_node = paragraph.find("./w:pPr/w:pStyle", DOCX_NS)
    style_id = style_node.attrib.get(DOCX_W + "val", "").lower() if style_node is not None else ""
    has_numbering = paragraph.find("./w:pPr/w:numPr", DOCX_NS) is not None
    compact = _compact_docx_text(text)
    if not compact:
        return False

    style_heading_like = style_id.startswith("heading") or style_id in {"1", "2", "3", "4", "21", "22", "23", "24"}
    # Some school templates store chapter/section numbers only in Word numbering.
    # The visible title text then has no leading digit, so detect the paragraph
    # from its heading style plus numbering state instead of text alone.
    return bool(has_numbering and style_heading_like and 2 <= len(compact) <= 80)


def _is_docx_non_body_terminal(text: str) -> bool:
    compact = _compact_docx_text(text)
    return compact.startswith(
        (
            "\u53c2\u8003\u6587\u732e",
            "references",
            "bibliography",
            "\u81f4\u8c22",
            "\u8c22\u8f9e",
            "acknowledgement",
            "acknowledgment",
            "\u9644\u5f55",
            "appendix",
        )
    )


def _is_docx_caption_or_table_marker(text: str) -> bool:
    compact = _compact_docx_text(text)
    stripped = re.sub(r"\s+", " ", text or "").strip()
    return bool(
        re.match(r"^(?:\u56fe|fig\.?|figure)\s*\d+(?:[-.\uff0d]\d+)?(?:\s+|[:：.\-\uff0d])", stripped, re.I)
        or re.match(r"^(?:\u8868|table)\s*\d+(?:[-.\uff0d]\d+)?(?:\s+|[:：.\-\uff0d])", stripped, re.I)
        or compact.startswith(("\u5173\u952e\u8bcd", "keywords:", "keywords\uff1a", "key words:", "key words\uff1a"))
    )


def _zh_abstract_body_paragraphs(root: ET.Element) -> list[ET.Element]:
    paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    title_seen = False
    result: list[ET.Element] = []
    for paragraph in paragraphs:
        visible = _docx_paragraph_text(paragraph)
        compact = _compact_docx_text(visible)
        if not title_seen:
            if compact == "\u6458\u8981":
                title_seen = True
            elif compact.startswith(("\u6458\u8981:", "\u6458\u8981\uff1a")):
                title_seen = True
                result.append(paragraph)
            continue
        if compact.startswith(("\u5173\u952e\u8bcd", "\u95dc\u9375\u8a5e")):
            break
        if visible:
            result.append(paragraph)
    return result


def _body_text_paragraphs(root: ET.Element) -> list[ET.Element]:
    paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    result: list[ET.Element] = []
    main_body_started = False
    toc_seen = False
    toc_closed = False
    for paragraph in paragraphs:
        text = _docx_paragraph_text(paragraph)
        if not text:
            continue
        compact = _compact_docx_text(text)
        if compact in {"\u76ee\u5f55", "contents", "tableofcontents"}:
            toc_seen = True
            continue
        if toc_seen and not toc_closed:
            style_node = paragraph.find("./w:pPr/w:pStyle", DOCX_NS)
            style_id = style_node.attrib.get(DOCX_W + "val", "").lower() if style_node is not None else ""
            if style_id.startswith("toc"):
                continue
            if paragraph.find(".//w:tab", DOCX_NS) is not None or paragraph.find("./w:pPr/w:tabs", DOCX_NS) is not None:
                continue
            if _is_docx_body_heading_paragraph(paragraph, text):
                toc_closed = True
                main_body_started = True
                continue
            if _is_docx_non_body_terminal(text):
                toc_closed = True
            continue
        if not main_body_started:
            if _is_docx_body_heading_paragraph(paragraph, text):
                main_body_started = True
            continue
        if _is_docx_non_body_terminal(text):
            break
        if paragraph.find(".//w:drawing", DOCX_NS) is not None or paragraph.find(".//w:pict", DOCX_NS) is not None:
            continue
        if _is_docx_body_heading_paragraph(paragraph, text) or _is_docx_caption_or_table_marker(text):
            continue
        if len(text) < 12:
            continue
        result.append(paragraph)
    return result


def _robust_body_text_paragraphs(root: ET.Element) -> list[ET.Element]:
    return _body_text_paragraphs(root)


def _front_matter_surface_paragraphs(root: ET.Element, surface_id: str) -> list[ET.Element]:
    paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    if surface_id == "zh_keyword_line":
        return [p for p in paragraphs if _compact_docx_text(_docx_paragraph_text(p)).startswith("\u5173\u952e\u8bcd")][:1]
    if surface_id == "en_keyword_line":
        return [
            p
            for p in paragraphs
            if _docx_paragraph_text(p).strip().lower().startswith(("keywords", "key words"))
        ][:1]
    if surface_id not in {"zh_abstract_body", "en_abstract_body"}:
        return []

    result: list[ET.Element] = []
    title_seen = False
    for paragraph in paragraphs:
        text = _docx_paragraph_text(paragraph)
        compact = _compact_docx_text(text)
        lowered = text.strip().lower()
        if not title_seen:
            if surface_id == "zh_abstract_body" and compact == "\u6458\u8981":
                title_seen = True
            elif surface_id == "zh_abstract_body" and compact.startswith(("\u6458\u8981:", "\u6458\u8981\uff1a")):
                title_seen = True
                result.append(paragraph)
            elif surface_id == "en_abstract_body" and lowered == "abstract":
                title_seen = True
            elif surface_id == "en_abstract_body" and lowered.startswith("abstract:"):
                title_seen = True
                result.append(paragraph)
            continue
        if surface_id == "zh_abstract_body" and compact.startswith("\u5173\u952e\u8bcd"):
            break
        if surface_id == "en_abstract_body" and lowered.startswith(("keywords", "key words")):
            break
        if text:
            result.append(paragraph)
    return result


def _direct_paragraph_metrics(paragraph: ET.Element) -> dict[str, str]:
    ppr = paragraph.find("./w:pPr", DOCX_NS)
    spacing = ppr.find("w:spacing", DOCX_NS) if ppr is not None else None
    ind = ppr.find("w:ind", DOCX_NS) if ppr is not None else None
    return {
        "before": spacing.attrib.get(DOCX_W + "before", "") if spacing is not None else "",
        "after": spacing.attrib.get(DOCX_W + "after", "") if spacing is not None else "",
        "line": spacing.attrib.get(DOCX_W + "line", "") if spacing is not None else "",
        "lineRule": spacing.attrib.get(DOCX_W + "lineRule", "") if spacing is not None else "",
        "firstLine": ind.attrib.get(DOCX_W + "firstLine", "") if ind is not None else "",
        "firstLineChars": ind.attrib.get(DOCX_W + "firstLineChars", "") if ind is not None else "",
    }


def _validate_direct_paragraph_metric_row(
    paragraph: ET.Element,
    *,
    surface_id: str,
    paragraph_index: int,
    require_real_firstline: bool,
) -> list[str]:
    text = _docx_paragraph_text(paragraph)
    metrics = _direct_paragraph_metrics(paragraph)
    issues: list[str] = []
    expected = {"before": "0", "after": "0", "line": "360", "lineRule": "auto"}
    for key, expected_value in expected.items():
        actual = metrics.get(key, "")
        if actual != expected_value:
            issues.append(
                f"{surface_id} paragraph p{paragraph_index} must have direct w:spacing {key}={expected_value}, found {actual or 'missing'} in `{text[:60]}`"
            )
    if require_real_firstline:
        first_line = metrics.get("firstLine", "")
        try:
            first_line_value = int(first_line)
        except ValueError:
            first_line_value = 0
        if first_line_value <= 0:
            char_value = metrics.get("firstLineChars", "")
            issues.append(
                f"{surface_id} paragraph p{paragraph_index} must have real direct w:ind/@w:firstLine; char-only firstLineChars={char_value or 'missing'} is not sufficient in `{text[:60]}`"
            )
    return issues


def _validate_surface_direct_paragraph_metrics_docx(docx_path: Path, surface_id: str) -> list[str]:
    if docx_path.suffix.lower() != ".docx":
        return [f"{surface_id} direct paragraph metrics must be verified against a DOCX reviewed output path, found: {docx_path}"]
    if not docx_path.exists():
        return [f"{surface_id} direct paragraph metrics reviewed DOCX path does not exist: {docx_path}"]
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return [f"{surface_id} direct paragraph metrics DOCX inspection failed for {docx_path}: {exc}"]

    if surface_id in {"zh_abstract_body", "en_abstract_body"}:
        paragraphs = _front_matter_surface_paragraphs(root, surface_id)
    elif surface_id == "body_text":
        paragraphs = _robust_body_text_paragraphs(root)
    else:
        paragraphs = _front_matter_surface_paragraphs(root, surface_id)
    if not paragraphs:
        return [f"{surface_id} direct paragraph metrics found no auditable paragraph in reviewed DOCX"]

    require_real_firstline = surface_id in {"zh_abstract_body", "en_abstract_body", "body_text"}
    issues: list[str] = []
    body_paragraphs = root.findall(".//w:body/w:p", DOCX_NS)
    index_by_id = {id(paragraph): index for index, paragraph in enumerate(body_paragraphs, start=1)}
    for paragraph in paragraphs:
        issues.extend(
            _validate_direct_paragraph_metric_row(
                paragraph,
                surface_id=surface_id,
                paragraph_index=index_by_id.get(id(paragraph), -1),
                require_real_firstline=require_real_firstline,
            )
        )
    return issues


def _validate_zh_abstract_body_mixed_font_docx(docx_path: Path) -> list[str]:
    if docx_path.suffix.lower() != ".docx":
        return [f"zh_abstract_body mixed-script font chain must be verified against a DOCX reviewed output path, found: {docx_path}"]
    if not docx_path.exists():
        return [f"zh_abstract_body mixed-script reviewed DOCX path does not exist: {docx_path}"]
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return [f"zh_abstract_body mixed-script DOCX inspection failed for {docx_path}: {exc}"]

    style_fonts = _docx_paragraph_style_fonts(root, docx_path)
    issues: list[str] = []
    checked = 0
    for paragraph in _zh_abstract_body_paragraphs(root):
        style_node = paragraph.find("./w:pPr/w:pStyle", DOCX_NS)
        style_id = style_node.attrib.get(DOCX_W + "val", "") if style_node is not None else ""
        paragraph_style_fonts = style_fonts.get(style_id, {})
        for run in paragraph.findall("w:r", DOCX_NS):
            if run.find(".//w:txbxContent", DOCX_NS) is not None:
                continue
            text = _docx_run_text(run)
            if not _docx_has_ascii_alnum(text):
                continue
            checked += 1
            slots = _docx_run_font_slots(run, paragraph_style_fonts)
            for key in ("ascii", "hAnsi"):
                value = slots.get(key, "")
                if not value:
                    issues.append(
                        f"zh_abstract_body Latin/digit run `{text[:40]}` has unresolved {key} font slot in reviewed DOCX"
                    )
                    break
                if value != "Times New Roman":
                    issues.append(
                        f"zh_abstract_body Latin/digit run `{text[:40]}` must resolve {key}=Times New Roman in reviewed DOCX, found {value}"
                    )
                    break
            else:
                cs_value = slots.get("cs", "")
                if not cs_value:
                    issues.append(
                        f"zh_abstract_body Latin/digit run `{text[:40]}` has unresolved cs font slot in reviewed DOCX"
                    )
                    continue
                if cs_value not in {"Times New Roman", "SimSun"}:
                    issues.append(
                        f"zh_abstract_body Latin/digit run `{text[:40]}` must resolve cs to template-compatible Times New Roman or SimSun in reviewed DOCX, found {cs_value}"
                    )
    return issues


def _validate_body_text_mixed_font_docx(docx_path: Path) -> list[str]:
    if docx_path.suffix.lower() != ".docx":
        return [f"body_text mixed-script font chain must be verified against a DOCX reviewed output path, found: {docx_path}"]
    if not docx_path.exists():
        return [f"body_text mixed-script reviewed DOCX path does not exist: {docx_path}"]
    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return [f"body_text mixed-script DOCX inspection failed for {docx_path}: {exc}"]

    style_fonts = _docx_paragraph_style_fonts(root, docx_path)
    issues: list[str] = []
    checked = 0

    def validate_latin_slots(text: str, run: ET.Element, paragraph_style_fonts: dict[str, str], *, mixed_run: bool) -> None:
        nonlocal checked
        if not _docx_has_ascii_alpha(text):
            return
        checked += 1
        slots = _docx_run_font_slots(run, paragraph_style_fonts)
        ascii_font = slots.get("ascii", "")
        hansi_font = slots.get("hAnsi", "")
        eastasia_font = slots.get("eastAsia", "")
        if not ascii_font or not hansi_font:
            issues.append(f"body_text Latin run `{text[:40]}` has unresolved ascii/hAnsi font slot in reviewed DOCX")
            return
        if mixed_run and not eastasia_font:
            issues.append(f"body_text mixed-script run `{text[:40]}` has unresolved East Asian font slot in reviewed DOCX")
            return
        if eastasia_font and ascii_font and eastasia_font == ascii_font:
            issues.append(
                f"body_text Latin run `{text[:40]}` uses the same East Asian and Western font family `{ascii_font}` in reviewed DOCX"
            )

    for paragraph in _body_text_paragraphs(root):
        visible = "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()
        if not (re.search(r"[\u4e00-\u9fff]", visible) and _docx_has_ascii_alpha(visible)):
            continue
        style_node = paragraph.find("./w:pPr/w:pStyle", DOCX_NS)
        style_id = style_node.attrib.get(DOCX_W + "val", "") if style_node is not None else ""
        paragraph_style_fonts = style_fonts.get(style_id, {})
        visible_runs = []
        for run in paragraph.findall("w:r", DOCX_NS):
            if run.find(".//w:txbxContent", DOCX_NS) is not None:
                continue
            text = _docx_run_text(run)
            if text:
                visible_runs.append((run, text))
        for run, text in visible_runs:
            validate_latin_slots(
                text,
                run,
                paragraph_style_fonts,
                mixed_run=bool(re.search(r"[\u4e00-\u9fff]", text)),
            )
    return issues


def _compact_keyword_text(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "").lower()


def _docx_run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", DOCX_NS))


def _docx_run_bold(run: ET.Element) -> bool:
    rpr = run.find("w:rPr", DOCX_NS)
    bold = rpr.find("w:b", DOCX_NS) if rpr is not None else None
    if bold is None:
        return False
    return bold.attrib.get(DOCX_W + "val", "1") not in {"0", "false", "False", "off", "OFF"}


def _docx_run_strong_face(run: ET.Element) -> bool:
    if _docx_run_bold(run):
        return True
    rpr = run.find("w:rPr", DOCX_NS)
    if rpr is None:
        return False
    rfonts = rpr.find("w:rFonts", DOCX_NS)
    if rfonts is None:
        return False
    strong_tokens = (
        "黑体",
        "simhei",
        "arial black",
    )
    return any(
        any(token in value.lower() for token in strong_tokens)
        for value in rfonts.attrib.values()
    )


def _docx_run_style_id(run: ET.Element) -> str:
    rpr = run.find("w:rPr", DOCX_NS)
    rstyle = rpr.find("w:rStyle", DOCX_NS) if rpr is not None else None
    return rstyle.get(DOCX_W + "val", "") if rstyle is not None else ""


def _docx_run_size_half_points(run: ET.Element) -> int | None:
    rpr = run.find("w:rPr", DOCX_NS)
    sz = rpr.find("w:sz", DOCX_NS) if rpr is not None else None
    raw = sz.get(DOCX_W + "val", "") if sz is not None else ""
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def _style_label_is_title_like(style_id: str, style_name: str = "") -> bool:
    compact_style = re.sub(r"[\s_\-]+", "", f"{style_id} {style_name}").lower()
    title_like_tokens = (
        "heading",
        "title",
        "toctitle",
        "tocheading",
        "caption",
        "abstracttitle",
        "abstitle",
    )
    keyword_tokens = ("keyword", "keywords", "key words")
    return any(token in compact_style for token in title_like_tokens) and not any(
        token in compact_style for token in keyword_tokens
    )


def _docx_run_title_face_like(run: ET.Element, char_style_names: dict[str, str]) -> bool:
    style_id = _docx_run_style_id(run)
    if _style_label_is_title_like(style_id, char_style_names.get(style_id, "")):
        return True
    size_hp = _docx_run_size_half_points(run)
    if size_hp is not None and size_hp >= 30:
        return True
    return _docx_run_strong_face(run)


def _keyword_label_for_text(text: str, surface_id: str) -> str:
    candidates = (
        ("关键词", "关键词"),
        ("关键词：", "关键词："),
        ("关键词:", "关键词:"),
    ) if surface_id == "zh_keyword_line" else (
        ("Key words:", "Key words:"),
        ("Key words：", "Key words："),
        ("Key Words:", "Key Words:"),
        ("Key Words：", "Key Words："),
        ("Keyword:", "Keyword:"),
        ("Keyword\uff1a", "Keyword\uff1a"),
        ("Keywords:", "Keywords:"),
        ("Keywords：", "Keywords："),
    )
    compact_text = _compact_keyword_text(text)
    if surface_id == "zh_keyword_line" and not compact_text.startswith(("\u5173\u952e\u8bcd\uff1a", "\u5173\u952e\u8bcd:")):
        return ""
    if surface_id == "en_keyword_line" and not compact_text.startswith(
        tuple(_compact_keyword_text(marker) for marker, _label in candidates)
    ):
        return ""
    for marker, label in sorted(candidates, key=lambda pair: len(_compact_keyword_text(pair[0])), reverse=True):
        if compact_text.startswith(_compact_keyword_text(marker)):
            return label
    return ""


def _docx_paragraph_style_names(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        styles_root = ET.fromstring(zf.read("word/styles.xml"))
    except (KeyError, ET.ParseError, OSError):
        return {}
    names: dict[str, str] = {}
    for style in styles_root.findall("./w:style", DOCX_NS):
        if style.get(DOCX_W + "type") != "paragraph":
            continue
        style_id = style.get(DOCX_W + "styleId") or ""
        name_node = style.find("./w:name", DOCX_NS)
        names[style_id] = name_node.get(DOCX_W + "val", "") if name_node is not None else ""
    return names


def _docx_character_style_names(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        styles_root = ET.fromstring(zf.read("word/styles.xml"))
    except (KeyError, ET.ParseError, OSError):
        return {}
    names: dict[str, str] = {}
    for style in styles_root.findall("./w:style", DOCX_NS):
        if style.get(DOCX_W + "type") != "character":
            continue
        style_id = style.get(DOCX_W + "styleId") or ""
        name_node = style.find("./w:name", DOCX_NS)
        names[style_id] = name_node.get(DOCX_W + "val", "") if name_node is not None else ""
    return names


def _keyword_paragraph_style_context(paragraph: ET.Element, style_names: dict[str, str]) -> dict[str, object]:
    ppr = paragraph.find("./w:pPr", DOCX_NS)
    style_id = ""
    outline_level = ""
    if ppr is not None:
        style_node = ppr.find("./w:pStyle", DOCX_NS)
        if style_node is not None:
            style_id = style_node.get(DOCX_W + "val", "") or ""
        outline_node = ppr.find("./w:outlineLvl", DOCX_NS)
        if outline_node is not None:
            outline_level = outline_node.get(DOCX_W + "val", "") or ""
    style_name = style_names.get(style_id, "")
    compact_style = re.sub(r"[\s_\-]+", "", f"{style_id} {style_name}").lower()
    return {
        "paragraphStyleId": style_id,
        "paragraphStyleName": style_name,
        "paragraphOutlineLevel": outline_level,
        "paragraphStyleTitleLike": _style_label_is_title_like(compact_style),
    }


def _keyword_signature_from_docx(docx_path: Path, surface_id: str) -> tuple[dict[str, object], list[str]]:
    issues: list[str] = []
    if docx_path.suffix.lower() != ".docx":
        return {}, [f"keyword run split must be verified against a DOCX reviewed output path, found: {docx_path}"]
    if not docx_path.exists():
        return {}, [f"keyword run split reviewed DOCX path does not exist: {docx_path}"]

    try:
        with zipfile.ZipFile(docx_path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
            style_names = _docx_paragraph_style_names(zf)
            char_style_names = _docx_character_style_names(zf)
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return {}, [f"keyword run split DOCX inspection failed for {docx_path}: {exc}"]

    target_paragraphs: list[tuple[int, ET.Element, str]] = []
    toc_index: int | None = None
    for index, paragraph in enumerate(root.findall(".//w:body/w:p", DOCX_NS), start=1):
        visible = "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS)).strip()
        compact_visible = _compact_keyword_text(visible)
        if toc_index is None and compact_visible in {_compact_keyword_text("\u76ee\u5f55"), "contents", "tableofcontents"}:
            toc_index = index
        label = _keyword_label_for_text(visible, surface_id)
        if label:
            target_paragraphs.append((index, paragraph, visible))
    if not target_paragraphs:
        return {}, [f"keyword run split could not locate {surface_id} in DOCX: {docx_path}"]
    if len(target_paragraphs) > 1:
        issues.append(f"keyword run split located multiple {surface_id} paragraphs in DOCX: {docx_path}")
    target_index, target_paragraph, _target_visible = target_paragraphs[0]
    if toc_index is not None and target_index > toc_index:
        issues.append(f"{surface_id} keyword paragraph appears after TOC instead of protected front matter")

    visible_text = "".join(node.text or "" for node in target_paragraph.findall(".//w:t", DOCX_NS)).strip()
    label = _keyword_label_for_text(visible_text, surface_id)
    label_compact = _compact_keyword_text(label)
    label_runs: list[ET.Element] = []
    content_runs: list[ET.Element] = []
    accumulated = ""
    label_complete = False
    for run in target_paragraph.findall("w:r", DOCX_NS):
        if run.find(".//w:txbxContent", DOCX_NS) is not None:
            continue
        text = _docx_run_text(run)
        if not text.strip():
            continue
        if label_compact and not label_complete:
            label_runs.append(run)
            accumulated += text
            if _compact_keyword_text(accumulated) == label_compact:
                label_complete = True
        else:
            content_runs.append(run)

    label_text = "".join(_docx_run_text(run) for run in label_runs)
    content_text = "".join(_docx_run_text(run) for run in content_runs)
    style_context = _keyword_paragraph_style_context(target_paragraph, style_names)
    signature: dict[str, object] = {
        "visibleText": visible_text,
        **style_context,
        "labelText": label_text,
        "labelRunCount": len(label_runs),
        "labelRunIsolated": bool(label_compact) and _compact_keyword_text(label_text) == label_compact and bool(content_runs),
        "labelRunBold": bool(label_runs) and all(_docx_run_bold(run) for run in label_runs),
        "labelRunStrong": bool(label_runs) and all(_docx_run_strong_face(run) for run in label_runs),
        "contentText": content_text,
        "contentRunCount": len(content_runs),
        "contentRunsBold": any(_docx_run_bold(run) for run in content_runs),
        "contentRunsStrongFace": any(_docx_run_strong_face(run) for run in content_runs),
        "contentRunsTitleFaceLike": any(_docx_run_title_face_like(run, char_style_names) for run in content_runs),
        "contentRunStyleIds": [
            _docx_run_style_id(run)
            for run in content_runs
            if _docx_run_style_id(run)
        ],
        "contentRunSizes": [
            _docx_run_size_half_points(run)
            for run in content_runs
            if _docx_run_size_half_points(run) is not None
        ],
    }

    if not signature["labelRunIsolated"]:
        issues.append(f"{surface_id} keyword label/content are not isolated into separate DOCX runs")
    if label_compact and _compact_keyword_text(label_text) != label_compact:
        issues.append(f"{surface_id} keyword label run must contain only the label text")
    if signature["contentRunCount"] < 1:
        issues.append(f"{surface_id} keyword content run count must be at least 1")
    if signature["contentRunsBold"]:
        issues.append(f"{surface_id} keyword content runs inherit label bolding")
    if signature["contentRunsTitleFaceLike"]:
        issues.append(f"{surface_id} keyword content runs inherit title-style formatting")
    if surface_id == "en_keyword_line" and not (signature["labelRunBold"] or signature["labelRunStrong"]):
        issues.append("en_keyword_line label must carry label-only bolding or template-approved strong face")
    if surface_id == "zh_keyword_line" and not signature["labelRunStrong"]:
        issues.append("zh_keyword_line label must carry template-approved strong face or bolding")
    if signature.get("paragraphStyleTitleLike") or signature.get("paragraphOutlineLevel"):
        issues.append(
            f"{surface_id} keyword paragraph must not use heading/title/TOC style or outline level: "
            f"styleId={signature.get('paragraphStyleId')!r}, "
            f"styleName={signature.get('paragraphStyleName')!r}, "
            f"outline={signature.get('paragraphOutlineLevel')!r}"
        )

    return signature, issues


def _validate_keyword_run_split_fields(
    values: dict[str, str],
    raw_values: dict[str, str],
    evidence_path: Path,
    required_surface_id: str,
) -> list[str]:
    issues: list[str] = []
    for prefix in STRICT_KEYWORD_RUN_SPLIT_PREFIXES:
        value = values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(f"surface evidence for {required_surface_id} lacks keyword run-split hard field {prefix} in {evidence_path}")

    label_isolated = normalize(raw_values.get("- keyword label run isolated baseline/actual:", "")).lower()
    if "actual=yes" not in label_isolated and "actual:yes" not in label_isolated and "actual yes" not in label_isolated:
        issues.append(f"surface evidence for {required_surface_id} must record actual keyword label isolation as yes in {evidence_path}")

    content_count = normalize(raw_values.get("- keyword content run count baseline/actual:", "")).lower()
    actual_counts = re.findall(r"actual\s*[:=]\s*(\d+)", content_count)
    if not actual_counts or int(actual_counts[-1]) < 1:
        issues.append(f"surface evidence for {required_surface_id} must record actual keyword content run count >= 1 in {evidence_path}")

    content_bold = normalize(raw_values.get("- keyword content bold baseline/actual:", "")).lower()
    if "actual=no" not in content_bold and "actual:no" not in content_bold and "actual no" not in content_bold:
        issues.append(f"surface evidence for {required_surface_id} must record actual keyword content bold state as no in {evidence_path}")

    verdict = values.get("- keyword run split verdict:", "")
    if verdict not in {"pass", "passed"} and not verdict.startswith("passed"):
        issues.append(f"surface evidence for {required_surface_id} must include a pass keyword run split verdict in {evidence_path}")

    reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
    _signature, docx_issues = _keyword_signature_from_docx(reviewed_output, required_surface_id)
    issues.extend(f"{issue} (reviewed output: {reviewed_output})" for issue in docx_issues)
    return issues


def check_required_surface_evidence_record(
    evidence_path: Path,
    *,
    expected_type: str,
    acceptance_mode: str,
    required_surface_id: str,
    expected_reviewed_output: Path | None = None,
    expected_reviewed_sha256: str | None = None,
) -> list[str]:
    """Strictly validate evidence for one template-owned thesis surface.

    This deliberately does not rely on surface-token autodetection. The caller
    already knows which inventory or matrix row is being proved, so a generic
    rendered-page record cannot silently pass as surface-specific evidence.
    """

    issues = check_review_evidence_record(
        evidence_path,
        expected_type=expected_type,
        acceptance_mode=acceptance_mode,
    )
    if issues:
        return issues

    if required_surface_id not in STRICT_FRONT_MATTER_SURFACE_FIELDS:
        return [f"unknown template-owned thesis surface id for strict evidence validation: {required_surface_id}"]

    values, raw_values, load_issues = _load_review_evidence_values(evidence_path)
    issues.extend(load_issues)
    if issues:
        return issues

    reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
    reviewed_sha = normalize(values.get("- reviewed output sha256:", ""))
    if expected_reviewed_output is not None:
        expected_path = expected_reviewed_output.resolve()
        if reviewed_output.resolve() != expected_path:
            issues.append(
                f"surface evidence for {required_surface_id} reviewed output path {reviewed_output.resolve()} "
                f"does not match final acceptance output {expected_path}: {evidence_path}"
            )
    if expected_reviewed_sha256:
        expected_sha = expected_reviewed_sha256.strip().lower()
        if reviewed_sha.lower() != expected_sha:
            issues.append(
                f"surface evidence for {required_surface_id} reviewed output sha256 {reviewed_sha or 'missing'} "
                f"does not match final acceptance output sha256 {expected_sha}: {evidence_path}"
            )

    combined_text = normalize("\n".join(raw_values.values())).lower()
    if contains_any(combined_text, STRICT_SURFACE_FALLBACK_TOKENS):
        issues.append(
            f"surface evidence for {required_surface_id} relies on forbidden substitute evidence: {evidence_path}"
        )

    for prefix in STRICT_FRONT_MATTER_SURFACE_FIELDS[required_surface_id]:
        if values.get(prefix) != "yes":
            issues.append(
                f"surface evidence for {required_surface_id} must confirm yes for {prefix} in {evidence_path}"
            )

    required_detail_prefixes = [
        "- baseline source path and sha256:",
        "- baseline surface id:",
        "- baseline paragraph/run path:",
        "- baseline metrics:",
        "- actual paragraph/run path:",
        "- actual metrics:",
        "- rendered region image path:",
    ]
    for prefix in required_detail_prefixes:
        value = values.get(prefix, "")
        if not is_explicit(value) or value in EXPLICIT_VALUES:
            issues.append(
                f"surface evidence for {required_surface_id} lacks baseline comparison detail {prefix} in {evidence_path}"
            )

    issues.extend(
        _validate_effective_font_fields(
            values,
            raw_values,
            evidence_path,
            subject=f"surface evidence for {required_surface_id}",
        )
    )
    issues.extend(_validate_surface_defect_hard_fields(values, raw_values, evidence_path, required_surface_id))
    if required_surface_id == "zh_abstract_body":
        reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_surface_direct_paragraph_metrics_docx(reviewed_output, required_surface_id)
        )
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_zh_abstract_body_mixed_font_docx(reviewed_output)
        )
    if required_surface_id == "en_abstract_body":
        reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_surface_direct_paragraph_metrics_docx(reviewed_output, required_surface_id)
        )
    if required_surface_id == "body_text":
        reviewed_output = resolve_record_path(raw_values.get("- reviewed output path:", ""), evidence_path)
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_surface_direct_paragraph_metrics_docx(reviewed_output, required_surface_id)
        )
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_body_text_mixed_font_docx(reviewed_output)
        )
    if required_surface_id == "cover_style":
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_cover_placeholder_residue_docx(reviewed_output)
        )
    if required_surface_id.startswith(("zh_", "en_")):
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in _validate_front_matter_vml_residue_docx(reviewed_output)
        )

    target_identifier = normalize(raw_values.get("- target identifier:", "")).lower()
    if target_identifier != required_surface_id:
        issues.append(
            f"surface evidence target identifier must be exact protected surface id {required_surface_id}: {evidence_path}"
        )
    canonical_surface_id = normalize(raw_values.get("- canonical protected surface id:", "")).lower()
    if canonical_surface_id != required_surface_id:
        issues.append(
            f"surface evidence canonical protected surface id must be exactly {required_surface_id}: {evidence_path}"
        )

    for prefix in ("- target surface:", "- target pages or region:"):
        if not _surface_alias_seen(raw_values.get(prefix, ""), required_surface_id):
            issues.append(
                f"surface evidence for {required_surface_id} must bind {prefix} to the same protected surface in {evidence_path}"
            )

    surface_text = normalize(raw_values.get("- baseline surface id:", "")).lower()
    if surface_text != required_surface_id:
        issues.append(
            f"surface evidence baseline surface id must be exactly {required_surface_id}: {evidence_path}"
        )

    for prefix in ("- baseline paragraph/run path:", "- actual paragraph/run path:"):
        if required_surface_id not in normalize(raw_values.get(prefix, "")).lower():
            issues.append(
                f"surface evidence for {required_surface_id} must bind {prefix} to the exact protected surface id in {evidence_path}"
            )

    issues.extend(
        _validate_baseline_source_path(
            raw_values.get("- baseline source path and sha256:", ""),
            evidence_path,
            required_surface_id,
        )
    )
    issues.extend(
        _validate_rendered_region_images(
            raw_values.get("- rendered region image path:", ""),
            evidence_path,
            required_surface_id,
        )
    )

    if required_surface_id in STRICT_TOC_SURFACE_IDS:
        issues.extend(
            _validate_toc_visual_geometry_fields(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
            )
        )
        issues.extend(
            _validate_toc_paragraph_typography_fields(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
            )
        )
    if required_surface_id in {"toc_entries", "toc_dotted_leaders", "toc_page_number_column"}:
        _toc_payload, toc_docx_issues = audit_docx_toc_dotted_leaders(reviewed_output)
        issues.extend(
            f"{issue} (reviewed output: {reviewed_output})"
            for issue in toc_docx_issues
        )
    if required_surface_id == "whole_document_pagination":
        issues.extend(
            _validate_whole_document_pagination_fields(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
            )
        )
    if required_surface_id in {"zh_keyword_line", "en_keyword_line"}:
        issues.extend(
            _validate_keyword_run_split_fields(
                values,
                raw_values,
                evidence_path,
                required_surface_id,
            )
        )
    if required_surface_id == "acknowledgement_title":
        issues.extend(_validate_acknowledgement_title_indent_position(raw_values, evidence_path))
    if required_surface_id == "acknowledgement_body":
        issues.extend(_validate_acknowledgement_body_hard_metrics(raw_values, evidence_path))
    if required_surface_id in {"references_title", "references_entries"}:
        issues.extend(_validate_references_surface_hard_metrics(raw_values, evidence_path, required_surface_id))

    metric_verdict = values.get("- metric-by-metric comparison verdict:", "")
    if metric_verdict not in {"pass", "passed"}:
        issues.append(
            f"surface evidence for {required_surface_id} must include a pass metric-by-metric comparison verdict in {evidence_path}"
        )

    if values.get("- forbidden substitute evidence used?:") != "no":
        issues.append(
            f"surface evidence for {required_surface_id} must set forbidden substitute evidence used?: no in {evidence_path}"
        )

    return issues
