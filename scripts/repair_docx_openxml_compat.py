#!/usr/bin/env python3
"""Repair schema-sensitive OOXML ordering without rewriting thesis content.

This helper is intentionally narrow. It copies the DOCX package, rewrites only
selected XML parts, restores known WordprocessingML property child order, and
adds missing namespace declarations required by mc:Ignorable values.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16SE_NS = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16_NS = "http://schemas.microsoft.com/office/word/2018/wordml"
W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
W16SDTDH_NS = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPS_CUSTOM_DATA_NS = "http://www.wps.cn/officeDocument/2013/wpsCustomData"
VML_OFFICE_NS = "urn:schemas-microsoft-com:office:office"

W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
MC = f"{{{MC_NS}}}"

DEFAULT_PARTS = (
    "word/document.xml",
    "word/styles.xml",
    "word/stylesWithEffects.xml",
    "word/numbering.xml",
    "word/settings.xml",
)
STYLE_REFERENCE_PART_RE = re.compile(
    r"^word/(?:document|styles|numbering|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$"
)
IGNORABLE_PREFIX_URIS = {
    "w14": W14_NS,
    "w15": W15_NS,
    "w16se": W16SE_NS,
    "w16cid": W16CID_NS,
    "w16": W16_NS,
    "w16cex": W16CEX_NS,
    "w16sdtdh": W16SDTDH_NS,
    "wp14": WP14_NS,
    "wpg": WPG_NS,
    "wps": WPS_NS,
    "wpsCustomData": WPS_CUSTOM_DATA_NS,
}

PPR_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]

RPR_ORDER = [
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
]

STYLE_ORDER = [
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
]

SECTPR_ORDER = [
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
    "sectPrChange",
]

TBLPR_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
]

TR_ORDER = [
    "tblPrEx",
    "trPr",
]

TCPR_ORDER = [
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
]

BORDER_ORDER = [
    "top",
    "left",
    "bottom",
    "right",
    "insideH",
    "insideV",
    "tl2br",
    "tr2bl",
]

TC_MARGIN_ORDER = [
    "top",
    "left",
    "bottom",
    "right",
]

TBLSTYLEPR_ORDER = [
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
]

MATHPR_ORDER = [
    "mathFont",
    "brkBin",
    "brkBinSub",
    "smallFrac",
    "dispDef",
    "lMargin",
    "rMargin",
    "defJc",
    "preSp",
    "postSp",
    "interSp",
    "intraSp",
    "wrapIndent",
    "wrapRight",
    "intLim",
    "naryLim",
]

MATH_RPR_ORDER = [
    "lit",
    "nor",
    "scr",
    "sty",
    "brk",
    "aln",
]

NUMBERING_ROOT_ORDER = {
    "numPicBullet": 0,
    "abstractNum": 1,
    "num": 2,
    "numIdMacAtCleanup": 3,
}

SETTINGS_ORDER = [
    "writeProtection",
    "view",
    "zoom",
    "removePersonalInformation",
    "removeDateAndTime",
    "doNotDisplayPageBoundaries",
    "displayBackgroundShape",
    "printPostScriptOverText",
    "printFractionalCharacterWidth",
    "printFormsData",
    "embedTrueTypeFonts",
    "embedSystemFonts",
    "saveSubsetFonts",
    "saveFormsData",
    "mirrorMargins",
    "alignBordersAndEdges",
    "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter",
    "gutterAtTop",
    "hideSpellingErrors",
    "hideGrammaticalErrors",
    "activeWritingStyle",
    "proofState",
    "formsDesign",
    "attachedTemplate",
    "linkStyles",
    "stylePaneFormatFilter",
    "stylePaneSortMethod",
    "documentType",
    "mailMerge",
    "revisionView",
    "trackRevisions",
    "doNotTrackMoves",
    "doNotTrackFormatting",
    "documentProtection",
    "autoFormatOverride",
    "styleLockTheme",
    "styleLockQFSet",
    "defaultTabStop",
    "autoHyphenation",
    "consecutiveHyphenLimit",
    "hyphenationZone",
    "doNotHyphenateCaps",
    "showEnvelope",
    "summaryLength",
    "clickAndTypeStyle",
    "defaultTableStyle",
    "evenAndOddHeaders",
    "bookFoldRevPrinting",
    "bookFoldPrinting",
    "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin",
    "doNotShadeFormData",
    "noPunctuationKerning",
    "characterSpacingControl",
    "printTwoOnOne",
    "strictFirstAndLastChars",
    "noLineBreaksAfter",
    "noLineBreaksBefore",
    "savePreviewPicture",
    "doNotValidateAgainstSchema",
    "saveInvalidXml",
    "ignoreMixedContent",
    "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml",
    "saveXmlDataOnly",
    "useXSLTWhenSaving",
    "saveThroughXslt",
    "showXMLTags",
    "alwaysMergeEmptyNamespace",
    "updateFields",
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "docVars",
    "rsids",
    "mathPr",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
]

BUILTIN_STYLE_ID_BY_TYPE_AND_NAME = {
    ("paragraph", "Normal"): "Normal",
    ("character", "Default Paragraph Font"): "DefaultParagraphFont",
    ("table", "Normal Table"): "TableNormal",
    ("paragraph", "Body Text"): "BodyText",
    ("paragraph", "Body Text 2"): "BodyText2",
    ("paragraph", "Body Text Indent"): "BodyTextIndent",
    ("paragraph", "Body Text Indent 2"): "BodyTextIndent2",
    ("paragraph", "TOC 1"): "TOC1",
    ("paragraph", "TOC 2"): "TOC2",
    ("paragraph", "TOC 3"): "TOC3",
    ("paragraph", "TOC 4"): "TOC4",
    ("paragraph", "TOC 5"): "TOC5",
    ("paragraph", "toc 1"): "TOC1",
    ("paragraph", "toc 2"): "TOC2",
    ("paragraph", "toc 3"): "TOC3",
    ("paragraph", "toc 4"): "TOC4",
    ("paragraph", "toc 5"): "TOC5",
    ("paragraph", "Heading 1"): "Heading1",
    ("paragraph", "Heading 2"): "Heading2",
    ("paragraph", "Heading 3"): "Heading3",
    ("paragraph", "Heading 4"): "Heading4",
    ("paragraph", "Heading 5"): "Heading5",
    ("paragraph", "Heading 6"): "Heading6",
    ("paragraph", "Heading 7"): "Heading7",
    ("paragraph", "Heading 8"): "Heading8",
    ("paragraph", "Heading 9"): "Heading9",
}

STYLE_REFERENCE_TAGS = {
    "pStyle",
    "rStyle",
    "tblStyle",
    "basedOn",
    "next",
    "link",
    "numStyleLink",
    "styleLink",
}


for prefix, uri in {
    "": PKG_REL_NS,
    "w": W_NS,
    "r": R_NS,
    "mc": MC_NS,
    "a": A_NS,
    "wp": WP_NS,
    "w14": W14_NS,
    "w15": W15_NS,
    "wp14": WP14_NS,
    "wps": WPS_NS,
    "wpsCustomData": WPS_CUSTOM_DATA_NS,
}.items():
    ET.register_namespace(prefix, uri)


def rel_qn(local: str) -> str:
    return f"{{{PKG_REL_NS}}}{local}"


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[1]
    return tag


def namespace_uri(tag: str) -> str:
    if tag.startswith("{"):
        return tag[1:].split("}", 1)[0]
    return ""


def order_children(parent: ET.Element, order: list[str]) -> bool:
    children = list(parent)
    if len(children) < 2:
        return False
    index = {name: idx for idx, name in enumerate(order)}

    def key(item: tuple[int, ET.Element]) -> tuple[int, int]:
        original_idx, child = item
        name = local_name(child.tag)
        return index.get(name, len(order)), original_idx

    sorted_children = [child for _, child in sorted(enumerate(children), key=key)]
    if sorted_children == children:
        return False
    parent[:] = sorted_children
    return True


def order_math_run_children(root: ET.Element) -> int:
    """Put OMML run properties before Word run properties for strict validators."""
    reordered = 0
    for element in root.iter(f"{M}r"):
        children = list(element)
        if len(children) < 2:
            continue

        def key(item: tuple[int, ET.Element]) -> tuple[int, int]:
            original_idx, child = item
            if child.tag == f"{M}rPr":
                return (0, original_idx)
            if child.tag == qn("rPr"):
                return (1, original_idx)
            if child.tag == f"{M}t":
                return (2, original_idx)
            return (3, original_idx)

        sorted_children = [child for _, child in sorted(enumerate(children), key=key)]
        if sorted_children != children:
            element[:] = sorted_children
            reordered += 1
    return reordered


def merge_duplicate_paragraph_properties(root: ET.Element) -> int:
    """Collapse duplicate paragraph property blocks into the first w:pPr."""

    removed = 0
    for paragraph in root.iter(qn("p")):
        pprs = [child for child in list(paragraph) if child.tag == qn("pPr")]
        if len(pprs) < 2:
            continue
        primary = pprs[0]
        existing_child_names = {local_name(child.tag) for child in list(primary)}
        for duplicate in pprs[1:]:
            for child in list(duplicate):
                child_name = local_name(child.tag)
                if child_name not in existing_child_names:
                    duplicate.remove(child)
                    primary.append(child)
                    existing_child_names.add(child_name)
            paragraph.remove(duplicate)
            removed += 1
        if list(paragraph).index(primary) != 0:
            paragraph.remove(primary)
            paragraph.insert(0, primary)
    return removed


def reorder_property_containers(root: ET.Element) -> dict[str, int]:
    counts = {
        "p": 0,
        "duplicate_pPr": 0,
        "pPr": 0,
        "rPr": 0,
        "style": 0,
        "sectPr": 0,
        "tblPr": 0,
        "tr": 0,
        "tcPr": 0,
        "tblStylePr": 0,
        "pBdr": 0,
        "tblBorders": 0,
        "tcBorders": 0,
        "tcMar": 0,
        "tblCellMar": 0,
        "mathPr": 0,
        "math_r": 0,
        "math_rPr": 0,
    }
    counts["duplicate_pPr"] = merge_duplicate_paragraph_properties(root)
    counts["math_r"] = order_math_run_children(root)
    for paragraph in root.iter(qn("p")):
        ppr = paragraph.find(qn("pPr"))
        if ppr is not None and list(paragraph).index(ppr) != 0:
            paragraph.remove(ppr)
            paragraph.insert(0, ppr)
            counts["p"] += 1
    for ppr in root.iter(qn("pPr")):
        if order_children(ppr, PPR_ORDER):
            counts["pPr"] += 1
    for rpr in root.iter(qn("rPr")):
        if order_children(rpr, RPR_ORDER):
            counts["rPr"] += 1
    for style in root.iter(qn("style")):
        if order_children(style, STYLE_ORDER):
            counts["style"] += 1
    for sectpr in root.iter(qn("sectPr")):
        if order_children(sectpr, SECTPR_ORDER):
            counts["sectPr"] += 1
    for tblpr in root.iter(qn("tblPr")):
        if order_children(tblpr, TBLPR_ORDER):
            counts["tblPr"] += 1
    for tr in root.iter(qn("tr")):
        if order_children(tr, TR_ORDER):
            counts["tr"] += 1
    for tcpr in root.iter(qn("tcPr")):
        if order_children(tcpr, TCPR_ORDER):
            counts["tcPr"] += 1
    for tblstylepr in root.iter(qn("tblStylePr")):
        if order_children(tblstylepr, TBLSTYLEPR_ORDER):
            counts["tblStylePr"] += 1
    for borders in root.iter(qn("pBdr")):
        if order_children(borders, BORDER_ORDER):
            counts["pBdr"] += 1
    for borders in root.iter(qn("tblBorders")):
        if order_children(borders, BORDER_ORDER):
            counts["tblBorders"] += 1
    for borders in root.iter(qn("tcBorders")):
        if order_children(borders, BORDER_ORDER):
            counts["tcBorders"] += 1
    for margin in root.iter(qn("tcMar")):
        if order_children(margin, TC_MARGIN_ORDER):
            counts["tcMar"] += 1
    for margin in root.iter(qn("tblCellMar")):
        if order_children(margin, TC_MARGIN_ORDER):
            counts["tblCellMar"] += 1
    for mathpr in root.iter(f"{M}mathPr"):
        if order_children(mathpr, MATHPR_ORDER):
            counts["mathPr"] += 1
    for math_rpr in root.iter(f"{M}rPr"):
        if order_children(math_rpr, MATH_RPR_ORDER):
            counts["math_rPr"] += 1
    return counts


def remove_duplicate_bookmark_ends(root: ET.Element) -> int:
    """Remove duplicate w:bookmarkEnd elements while preserving the first pair close."""
    seen: set[str] = set()
    removed = 0

    def visit(parent: ET.Element) -> None:
        nonlocal removed
        for child in list(parent):
            if child.tag == qn("bookmarkEnd"):
                bookmark_id = child.attrib.get(qn("id"), "")
                if bookmark_id in seen:
                    parent.remove(child)
                    removed += 1
                    continue
                seen.add(bookmark_id)
            visit(child)

    visit(root)
    return removed


def remove_duplicate_bookmark_starts(root: ET.Element) -> int:
    """Remove duplicate w:bookmarkStart elements with the same id."""
    seen: set[str] = set()
    removed = 0

    def visit(parent: ET.Element) -> None:
        nonlocal removed
        for child in list(parent):
            if child.tag == qn("bookmarkStart"):
                bookmark_id = child.attrib.get(qn("id"), "")
                if bookmark_id in seen:
                    parent.remove(child)
                    removed += 1
                    continue
                seen.add(bookmark_id)
            visit(child)

    visit(root)
    return removed


def remove_false_no_wrap_values(root: ET.Element) -> int:
    """Drop schema-hostile noWrap=false elements from table cells."""
    removed = 0

    def visit(parent: ET.Element) -> None:
        nonlocal removed
        for child in list(parent):
            if child.tag == qn("noWrap") and child.attrib.get(qn("val"), "").strip() in {"0", "false", "off"}:
                parent.remove(child)
                removed += 1
                continue
            visit(child)

    visit(root)
    return removed


def normalize_on_off_values(root: ET.Element) -> int:
    """Convert non-canonical boolean values on common on/off elements."""
    normalized = 0
    for element in root.iter():
        name = local_name(element.tag)
        value = element.attrib.get(qn("val"))
        if value is None:
            continue
        clean_value = value.strip().lower()
        if name in {"cantSplit", "tblHeader", "keepNext", "keepLines", "b", "bCs", "i", "iCs"}:
            if clean_value in {"true", "1", "on"}:
                element.attrib.pop(qn("val"), None)
                normalized += 1
            elif clean_value == "false":
                element.attrib[qn("val")] = "0"
                normalized += 1
    return normalized


def normalize_start_end_justification(root: ET.Element) -> int:
    """Map newer logical justification values to Word 2007 schema values."""
    normalized = 0
    mapping = {"start": "left", "end": "right"}
    for element in root.iter():
        if local_name(element.tag) not in {"jc", "lvlJc"}:
            continue
        value = element.attrib.get(qn("val"))
        if value in mapping:
            element.attrib[qn("val")] = mapping[value]
            normalized += 1
    return normalized


def normalize_font_charset_attrs(root: ET.Element) -> int:
    """Drop non-schema charset metadata emitted by some Office/WPS stacks."""
    normalized = 0
    for element in root.iter():
        if local_name(element.tag) != "charset":
            continue
        for attr in list(element.attrib):
            if local_name(attr) == "characterSet":
                del element.attrib[attr]
                normalized += 1
    return normalized


def normalize_text_space_attrs(root: ET.Element) -> int:
    """Move invalid w:space on text nodes to xml:space while leaving borders alone."""
    normalized = 0
    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
    for element in root.iter(qn("t")):
        value = element.attrib.pop(qn("space"), None)
        if value is None:
            continue
        if element.attrib.get(xml_space) != value:
            element.attrib[xml_space] = value
        normalized += 1
    return normalized


def normalize_shading_values(root: ET.Element) -> int:
    """Add the required w:val value to shading records that only carry color data."""
    normalized = 0
    for element in root.iter(qn("shd")):
        if qn("val") not in element.attrib:
            element.attrib[qn("val")] = "clear"
            normalized += 1
    return normalized


def remove_conflicting_math_sty_with_normal_text(root: ET.Element) -> int:
    """Drop OMML m:sty when m:nor is present in the same math run properties."""
    removed = 0
    for math_rpr in root.iter(f"{M}rPr"):
        has_nor = any(child.tag == f"{M}nor" for child in list(math_rpr))
        if not has_nor:
            continue
        for child in list(math_rpr):
            if child.tag == f"{M}sty":
                math_rpr.remove(child)
                removed += 1
    return removed


def cap_vml_shapetype_spt(root: ET.Element) -> int:
    """Keep legacy VML shapetype spt values inside the local validator range."""
    normalized = 0
    for element in root.iter():
        if local_name(element.tag) != "shapetype":
            continue
        for attr in list(element.attrib):
            if local_name(attr) != "spt":
                continue
            try:
                value = int(element.attrib[attr])
            except ValueError:
                continue
            if value > 202:
                element.attrib[attr] = "202"
                normalized += 1
    return normalized


def normalize_vml_shapetype_type_attrs(root: ET.Element) -> int:
    """Drop invalid VML shapetype type= references from fallback definitions."""
    normalized = 0
    for element in root.iter():
        if namespace_uri(element.tag) != "urn:schemas-microsoft-com:vml" or local_name(element.tag) != "shapetype":
            continue
        type_value = element.attrib.pop("type", None)
        if type_value is None:
            continue
        normalized += 1
    return normalized


def remove_orphan_note_settings(root: ET.Element) -> dict[str, int]:
    """Remove settings-level footnote/endnote refs when no note parts exist."""
    removed = {"footnotePr": 0, "endnotePr": 0}
    if local_name(root.tag) != "settings":
        return removed
    for name in ("footnotePr", "endnotePr"):
        for node in list(root.findall(qn(name))):
            root.remove(node)
            removed[name] += 1
    return removed


def remove_strict_validator_extension_attrs(root: ET.Element) -> int:
    """Drop extension attributes rejected by the local schema validator."""
    removed = 0
    for element in root.iter():
        for attr in list(element.attrib):
            if attr in {
                f"{{{W15_NS}}}restartNumberingAfterBreak",
                f"{{{W16CID_NS}}}durableId",
            }:
                del element.attrib[attr]
                removed += 1
    return removed


def remove_strict_validator_extension_elements(root: ET.Element) -> int:
    """Drop extension elements rejected by the local schema validator."""
    removed = 0

    def visit(parent: ET.Element) -> None:
        nonlocal removed
        for child in list(parent):
            if child.tag == f"{{{W14_NS}}}ligatures":
                parent.remove(child)
                removed += 1
                continue
            visit(child)

    visit(root)
    return removed


def reorder_numbering_root_children(root: ET.Element) -> int:
    """Restore WordprocessingML numbering root order: abstractNum definitions before nums."""
    if local_name(root.tag) != "numbering":
        return 0
    children = list(root)
    ordered = [
        child
        for _index, child in sorted(
            enumerate(children),
            key=lambda item: (NUMBERING_ROOT_ORDER.get(local_name(item[1].tag), 99), item[0]),
        )
    ]
    if ordered == children:
        return 0
    root[:] = ordered
    return 1


def reorder_settings_root_children(root: ET.Element) -> int:
    """Restore WordprocessingML settings root order for schema validators."""
    if local_name(root.tag) != "settings":
        return 0
    return 1 if order_children(root, SETTINGS_ORDER) else 0


def ignorable_prefixes(root: ET.Element) -> set[str]:
    value = root.attrib.get(f"{MC}Ignorable", "")
    return {token for token in re.split(r"\s+", value.strip()) if token}


def normalize_ignorable_prefixes(root: ET.Element) -> dict[str, object]:
    """Keep mc:Ignorable limited to known Office compatibility prefixes."""
    value = root.attrib.get(f"{MC}Ignorable", "")
    tokens = [token for token in re.split(r"\s+", value.strip()) if token]
    kept = [token for token in tokens if token in IGNORABLE_PREFIX_URIS]
    removed = [token for token in tokens if token not in IGNORABLE_PREFIX_URIS]
    if tokens:
        if kept:
            root.attrib[f"{MC}Ignorable"] = " ".join(dict.fromkeys(kept))
        else:
            root.attrib.pop(f"{MC}Ignorable", None)
    return {
        "removed_invalid_ignorable_prefixes": sorted(set(removed)),
        "kept_ignorable_prefixes": sorted(set(kept)),
    }


def markup_compatibility_prefixes(root: ET.Element) -> set[str]:
    prefixes = set(ignorable_prefixes(root))
    for element in root.iter():
        requires = element.attrib.get("Requires", "")
        prefixes.update(token for token in re.split(r"\s+", requires.strip()) if token)
    return prefixes


def inject_missing_namespace_declarations(xml_text: str, prefixes: set[str]) -> tuple[str, list[str]]:
    missing = [prefix for prefix in sorted(prefixes) if prefix in IGNORABLE_PREFIX_URIS and f"xmlns:{prefix}=" not in xml_text]
    if not missing:
        return order_ignorable_namespace_declarations(xml_text, prefixes), []
    match = re.search(r"<(?!\?)(?!!--)([A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*(?:\s[^<>]*)?>", xml_text)
    if not match:
        return xml_text, []
    insertion = "".join(f' xmlns:{prefix}="{IGNORABLE_PREFIX_URIS[prefix]}"' for prefix in missing)
    root_tag = match.group(0)
    ignorable_match = re.search(r"\s+[A-Za-z0-9_.-]+:Ignorable=", root_tag)
    if ignorable_match:
        insert_at = match.start() + ignorable_match.start()
    else:
        insert_at = match.end() - (2 if xml_text[match.end() - 2 : match.end()] == "/>" else 1)
    xml_text = xml_text[:insert_at] + insertion + xml_text[insert_at:]
    return order_ignorable_namespace_declarations(xml_text, prefixes), missing


def order_ignorable_namespace_declarations(xml_text: str, prefixes: set[str]) -> str:
    """Place namespace declarations used by mc:Ignorable before that attribute for strict validators."""
    match = re.search(r"<(?!\?)(?!!--)([A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*(?:\s[^<>]*)?>", xml_text)
    if not match:
        return xml_text
    root_tag = match.group(0)
    ignorable_match = re.search(r"\s+[A-Za-z0-9_.-]+:Ignorable=", root_tag)
    if not ignorable_match:
        return xml_text
    declarations: list[str] = []
    rewritten = root_tag
    for prefix in sorted(prefixes):
        uri = IGNORABLE_PREFIX_URIS.get(prefix)
        if not uri:
            continue
        attr_re = re.compile(rf'\s+xmlns:{re.escape(prefix)}="{re.escape(uri)}"')
        attr_match = attr_re.search(rewritten)
        if not attr_match:
            continue
        declarations.append(attr_match.group(0))
        rewritten = rewritten[: attr_match.start()] + rewritten[attr_match.end() :]
    if not declarations:
        return xml_text
    ignorable_match = re.search(r"\s+[A-Za-z0-9_.-]+:Ignorable=", rewritten)
    if not ignorable_match:
        return xml_text
    rewritten = rewritten[: ignorable_match.start()] + "".join(declarations) + rewritten[ignorable_match.start() :]
    return xml_text[: match.start()] + rewritten + xml_text[match.end() :]


def build_builtin_style_id_map(styles_payload: bytes) -> dict[str, str]:
    """Map WPS-renumbered built-in style IDs back to stable OOXML IDs."""
    try:
        root = ET.fromstring(styles_payload)
    except ET.ParseError:
        return {}
    existing_ids = {
        style.attrib.get(qn("styleId"), "")
        for style in root.iter(qn("style"))
        if style.attrib.get(qn("styleId"), "")
    }
    remap: dict[str, str] = {}
    for style in root.iter(qn("style")):
        style_id = style.attrib.get(qn("styleId"), "")
        style_type = style.attrib.get(qn("type"), "")
        name = style.find(qn("name"))
        name_value = name.attrib.get(qn("val"), "") if name is not None else ""
        canonical_id = BUILTIN_STYLE_ID_BY_TYPE_AND_NAME.get((style_type, name_value))
        if not style_id or not canonical_id or style_id == canonical_id:
            continue
        if canonical_id in existing_ids:
            continue
        remap[style_id] = canonical_id
        existing_ids.add(canonical_id)
    return remap


def apply_builtin_style_id_map(root: ET.Element, style_id_map: dict[str, str]) -> dict[str, int]:
    counts = {"style_ids": 0, "style_references": 0}
    if not style_id_map:
        return counts
    for element in root.iter():
        name = local_name(element.tag)
        if name == "style":
            style_id = element.attrib.get(qn("styleId"), "")
            mapped = style_id_map.get(style_id)
            if mapped:
                element.attrib[qn("styleId")] = mapped
                counts["style_ids"] += 1
        if name in STYLE_REFERENCE_TAGS:
            value = element.attrib.get(qn("val"), "")
            mapped = style_id_map.get(value)
            if mapped:
                element.attrib[qn("val")] = mapped
                counts["style_references"] += 1
    return counts


def remove_style_tbl_layout(root: ET.Element) -> int:
    """Remove table-layout properties from table style definitions.

    Word often tolerates w:tblLayout in style table properties, but strict
    OpenXML validators reject it in this context. Actual document tables are
    left untouched; this only normalizes style definitions.
    """
    removed = 0
    for style in root.iter(qn("style")):
        for tblpr in style.iter(qn("tblPr")):
            for child in list(tblpr):
                if child.tag == qn("tblLayout"):
                    tblpr.remove(child)
                    removed += 1
    return removed


def repair_xml_part(
    payload: bytes,
    *,
    style_id_map: dict[str, str] | None = None,
    part_name: str = "",
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(payload)
    ignorable_normalization = normalize_ignorable_prefixes(root)
    prefixes = markup_compatibility_prefixes(root)
    style_counts = apply_builtin_style_id_map(root, style_id_map or {})
    removed_duplicate_bookmark_starts = remove_duplicate_bookmark_starts(root)
    removed_duplicate_bookmark_ends = remove_duplicate_bookmark_ends(root)
    removed_false_no_wrap = remove_false_no_wrap_values(root)
    normalized_on_off_values = normalize_on_off_values(root)
    normalized_start_end_justification = normalize_start_end_justification(root)
    normalized_font_charset_attrs = normalize_font_charset_attrs(root)
    normalized_text_space_attrs = normalize_text_space_attrs(root)
    normalized_shading_values = normalize_shading_values(root)
    removed_conflicting_math_sty = remove_conflicting_math_sty_with_normal_text(root)
    capped_vml_shapetype_spt = cap_vml_shapetype_spt(root)
    normalized_vml_shapetype_type_attrs = normalize_vml_shapetype_type_attrs(root)
    removed_orphan_note_settings = remove_orphan_note_settings(root)
    removed_extension_attrs = remove_strict_validator_extension_attrs(root)
    removed_extension_elements = remove_strict_validator_extension_elements(root)
    removed_style_tbl_layout = (
        remove_style_tbl_layout(root)
        if part_name in {"word/styles.xml", "word/stylesWithEffects.xml"}
        else 0
    )
    reordered_numbering_root = reorder_numbering_root_children(root)
    reordered_settings_root = reorder_settings_root_children(root)
    counts = reorder_property_containers(root)
    xml_text = ET.tostring(root, encoding="unicode", xml_declaration=True)
    xml_text, added_prefixes = inject_missing_namespace_declarations(xml_text, prefixes)
    changed_payload = xml_text.encode("utf-8")
    return changed_payload, {
        "canonicalized_builtin_style_ids": style_counts["style_ids"],
        "canonicalized_builtin_style_references": style_counts["style_references"],
        "builtin_style_id_map": dict(sorted((style_id_map or {}).items())),
        "removed_duplicate_bookmark_starts": removed_duplicate_bookmark_starts,
        "removed_duplicate_bookmark_ends": removed_duplicate_bookmark_ends,
        "removed_false_no_wrap": removed_false_no_wrap,
        "normalized_on_off_values": normalized_on_off_values,
        "normalized_start_end_justification": normalized_start_end_justification,
        "normalized_font_charset_attrs": normalized_font_charset_attrs,
        "normalized_text_space_attrs": normalized_text_space_attrs,
        "normalized_shading_values": normalized_shading_values,
        "removed_conflicting_math_sty": removed_conflicting_math_sty,
        "capped_vml_shapetype_spt": capped_vml_shapetype_spt,
        "normalized_vml_shapetype_type_attrs": normalized_vml_shapetype_type_attrs,
        "removed_orphan_note_settings": removed_orphan_note_settings,
        "removed_strict_validator_extension_attrs": removed_extension_attrs,
        "removed_strict_validator_extension_elements": removed_extension_elements,
        "removed_style_tbl_layout": removed_style_tbl_layout,
        "reordered_numbering_root": reordered_numbering_root,
        "reordered_settings_root": reordered_settings_root,
        "reordered_p": counts["p"],
        "removed_duplicate_pPr": counts["duplicate_pPr"],
        "reordered_pPr": counts["pPr"],
        "reordered_rPr": counts["rPr"],
        "reordered_style": counts["style"],
        "reordered_sectPr": counts["sectPr"],
        "reordered_tblPr": counts["tblPr"],
        "reordered_tr": counts["tr"],
        "reordered_tcPr": counts["tcPr"],
        "reordered_tblStylePr": counts["tblStylePr"],
        "reordered_pBdr": counts["pBdr"],
        "reordered_tblBorders": counts["tblBorders"],
        "reordered_tcBorders": counts["tcBorders"],
        "reordered_tcMar": counts["tcMar"],
        "reordered_tblCellMar": counts["tblCellMar"],
        "reordered_mathPr": counts["mathPr"],
        "reordered_math_r": counts["math_r"],
        "reordered_math_rPr": counts["math_rPr"],
        "ignorable_prefixes": sorted(prefixes),
        **ignorable_normalization,
        "added_namespace_prefixes": added_prefixes,
    }


def repair_relationships_part(payload: bytes) -> tuple[bytes, dict[str, object]]:
    """Normalize .rels parts away from ns0-prefixed ElementTree output.

    LibreOffice is stricter than python-docx for some relationship parts.  A
    package relationships file serialized as ``ns0:Relationships`` can validate
    as XML yet fail conversion, so keep these package parts in the conventional
    default namespace form used by Word.
    """
    root = ET.fromstring(payload)
    renamed_root = 0
    renamed_children = 0
    if local_name(root.tag) == "Relationships" and root.tag != rel_qn("Relationships"):
        root.tag = rel_qn("Relationships")
        renamed_root = 1
    for child in root:
        if local_name(child.tag) == "Relationship" and child.tag != rel_qn("Relationship"):
            child.tag = rel_qn("Relationship")
            renamed_children += 1
    body = ET.tostring(root, encoding="unicode", xml_declaration=False, short_empty_elements=True)
    xml_text = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n" + body
    return xml_text.encode("utf-8"), {
        "relationship_root_normalized": renamed_root,
        "relationship_children_normalized": renamed_children,
    }


def read_latest_part(archive: zipfile.ZipFile, name: str) -> bytes:
    for item in reversed(archive.infolist()):
        if item.filename == name:
            with archive.open(item) as handle:
                return handle.read()
    raise KeyError(name)


def copy_docx_with_replacements(source: Path, output: Path, replacements: dict[str, bytes]) -> None:
    if source.resolve() == output.resolve():
        raise ValueError("input and output DOCX paths must be different")
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=str(output.parent))
    os.close(fd)
    Path(tmp_name).unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            unique_items: list[zipfile.ZipInfo] = []
            seen: set[str] = set()
            for item in reversed(zin.infolist()):
                if item.filename in seen:
                    continue
                seen.add(item.filename)
                unique_items.append(item)
            for item in reversed(unique_items):
                payload = replacements.get(item.filename)
                if payload is None:
                    with zin.open(item) as source_entry:
                        payload = source_entry.read()
                info = zipfile.ZipInfo(item.filename, item.date_time)
                info.comment = item.comment
                info.extra = item.extra
                info.internal_attr = item.internal_attr
                info.external_attr = item.external_attr
                info.create_system = item.create_system
                info.compress_type = item.compress_type
                zout.writestr(info, payload)
        os.replace(tmp_name, output)
    finally:
        Path(tmp_name).unlink(missing_ok=True)


def repair_docx(source: Path, output: Path, parts: list[str]) -> dict[str, object]:
    replacements: dict[str, bytes] = {}
    part_reports: list[dict[str, object]] = []
    with zipfile.ZipFile(source, "r") as zin:
        existing = set(zin.namelist())
        style_id_map = (
            build_builtin_style_id_map(read_latest_part(zin, "word/styles.xml"))
            if "word/styles.xml" in existing
            else {}
        )
        expanded_parts = set(parts)
        expanded_parts.update(name for name in existing if name.endswith(".rels"))
        expanded_parts.update(
            name
            for name in existing
            if STYLE_REFERENCE_PART_RE.match(name)
        )
        if style_id_map:
            expanded_parts.update(name for name in existing if STYLE_REFERENCE_PART_RE.match(name))
        for part in sorted(expanded_parts):
            if part not in existing:
                part_reports.append({"part": part, "status": "missing"})
                continue
            original = read_latest_part(zin, part)
            try:
                if part.endswith(".rels"):
                    repaired, details = repair_relationships_part(original)
                else:
                    repaired, details = repair_xml_part(original, style_id_map=style_id_map, part_name=part)
            except ET.ParseError as exc:
                part_reports.append({"part": part, "status": "parse-error", "message": str(exc)})
                continue
            changed = repaired != original
            if changed:
                replacements[part] = repaired
            part_reports.append({"part": part, "status": "changed" if changed else "unchanged", **details})
    copy_docx_with_replacements(source, output, replacements)
    return {
        "schema": "graduation-project-builder.docx-openxml-compat-repair.v1",
        "generator": "scripts/repair_docx_openxml_compat.py",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source_docx_path": str(source),
        "source_docx_sha256": sha256_file(source),
        "final_docx_path": str(output),
        "final_docx_sha256": sha256_file(output),
        "parts_requested": parts,
        "builtin_style_id_map": dict(sorted(style_id_map.items())),
        "changed_parts": sorted(replacements),
        "part_reports": part_reports,
        "verdict": "pass" if replacements else "pass-noop",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair narrow OOXML compatibility issues in a DOCX package.")
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--report-json", type=Path)
    parser.add_argument("--parts", nargs="*", default=list(DEFAULT_PARTS))
    args = parser.parse_args()

    report = repair_docx(args.input_docx, args.output_docx, list(args.parts))
    if args.report_json:
        args.report_json.parent.mkdir(parents=True, exist_ok=True)
        args.report_json.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
