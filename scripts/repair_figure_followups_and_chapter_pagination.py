#!/usr/bin/env python3
"""Repair figure follow-up prose and chapter-start pagination ownership.

This script is intentionally bounded to two surfaces:
- missing body prose immediately after figure captions
- missing explicit page-start ownership on level-1 body chapter headings

It edits only word/document.xml in a fresh review copy and preserves existing
media, relationships, styles, numbering, headers, and footers.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = f"{{{W_NS}}}"
NS = {"w": W_NS, "r": R_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


def qn(local: str) -> str:
    return f"{W}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "")


def heading_level(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    if "\u2026" in stripped or "\t" in stripped:
        return None
    sep = r"[\s\u25a1]+"
    if re.match(r"^\u7b2c\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{1,3}\s*\u7ae0", stripped):
        return 1
    first_level_number = re.match(rf"^(\d{{1,2}}){sep}\S", stripped)
    if first_level_number and 1 <= int(first_level_number.group(1)) <= 20:
        return 1
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}{sep}\S", stripped):
        return 2
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}\.\d{{1,2}}{sep}\S", stripped):
        return 3
    return None


def is_toc_heading(text: str) -> bool:
    compact = compact_text(text).lower()
    if compact in {compact_text("\u76ee\u5f55").lower(), "contents", "tableofcontents"}:
        return True
    return compact_text(text).lower() in {compact_text("目录").lower(), "contents", "tableofcontents"}


def strip_toc_page_number(text: str) -> str:
    value = str(text or "").strip()
    if "\t" in value:
        return value.split("\t", 1)[0].strip()
    if "\u2026" in value:
        return value.split("\u2026", 1)[0].strip()
    return re.sub(r"\s*\d+\s*$", "", value).strip()


def looks_like_toc_entry(text: str) -> bool:
    value = str(text or "").strip()
    if not value:
        return False
    if "\t" in value or "\u2026" in value:
        return heading_level(strip_toc_page_number(value)) is not None
    return bool(re.search(r"\d+\s*$", value)) and heading_level(strip_toc_page_number(value)) is not None


def is_reference_heading(text: str) -> bool:
    compact = compact_text(text).lower()
    if compact in {compact_text("\u53c2\u8003\u6587\u732e").lower(), "references", "bibliography"}:
        return True
    return compact_text(text).lower() in {compact_text("参考文献").lower(), "references", "bibliography"}


def has_image(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:drawing", NS) is not None
        or paragraph.find(".//w:pict", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
    )


def has_page_break(paragraph: ET.Element) -> bool:
    return any(br.attrib.get(qn("type"), "textWrapping") == "page" for br in paragraph.findall(".//w:br", NS))


def has_page_break_before(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:pageBreakBefore", NS) is not None


def has_section_break(paragraph: ET.Element) -> bool:
    return paragraph.find("./w:pPr/w:sectPr", NS) is not None


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def add_page_break_before(paragraph: ET.Element) -> bool:
    ppr = ensure_ppr(paragraph)
    if ppr.find("./w:pageBreakBefore", NS) is not None:
        return False
    style = ppr.find("./w:pStyle", NS)
    page_break = ET.Element(qn("pageBreakBefore"))
    if style is not None:
        children = list(ppr)
        ppr.insert(children.index(style) + 1, page_break)
    else:
        ppr.insert(0, page_break)
    return True


def paragraph_style_id(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.attrib.get(qn("val"), "") if style is not None else ""


def paragraph_style_map(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read("word/styles.xml"))
    except KeyError:
        return {}
    result: dict[str, str] = {}
    for style in root.findall("./w:style", NS):
        style_id = style.attrib.get(qn("styleId"), "")
        name = style.find("./w:name", NS)
        result[style_id] = name.attrib.get(qn("val"), "") if name is not None else ""
    return result


def is_figure_caption(text: str) -> bool:
    return re.match(r"^\s*图\s*\d+\s*[-－—]\s*\d+\s+\S+", text or "") is not None


def is_table_caption(text: str) -> bool:
    return re.match(r"^\s*表\s*\d+\s*[-－—]\s*\d+\s+\S+", text or "") is not None


def followup_text_for_caption(caption: str) -> str:
    normalized = re.sub(r"\s+", " ", caption.strip())
    explicit = {
        "图4-1 系统整体架构图": "系统整体架构图展示了前端应用、后端服务、数据库和文件存储之间的协作关系，为后续数据库设计和功能实现提供结构依据。",
        "图5-1 前端整体架构图": "前端整体架构图说明了页面组件、路由管理、接口请求和状态维护之间的组织关系，体现了前端实现的主要模块边界。",
        "图5-2 登录页面": "登录页面承担用户身份校验入口功能，通过账号密码输入、记住密码选项和登录结果反馈完成系统访问前的认证交互。",
        "图5-3 注册与密码重置页面": "注册与密码重置页面覆盖新用户开户注册和遗忘密码恢复两类场景，通过邮箱验证和表单校验提高账户操作的可靠性。",
        "图5-4 首页仪表盘页面": "首页仪表盘页面集中展示分类、热门资源和最新资源，为用户进入平台后的资源浏览与快速访问提供入口。",
        "图5-5 资源列表页面": "资源列表页面通过分类筛选、分页和资源卡片呈现平台资源，便于用户按主题快速检索和比较学习资料。",
        "图5-6 资源详情页面": "资源详情页面展示资源描述、元数据以及下载收藏操作，是用户确认资源内容并执行后续学习行为的核心页面。",
        "图5-7 资源上传页面": "资源上传页面用于提交资源文件及其分类、名称和描述信息，配合后端审核流程保证学习资源进入平台前具备基本的规范性。",
        "图5-8 我的收藏页面": "我的收藏页面将收藏夹管理与资源卡片展示结合起来，方便用户按个人学习计划维护和访问已收藏资源。",
        "图5-9 个人中心页面": "个人中心页面集中展示用户资料、学习行为数据和常用操作入口，便于用户维护个人信息并查看资源使用情况。",
    }
    if normalized in explicit:
        return explicit[normalized]
    title = re.sub(r"^\s*图\s*\d+\s*[-－—]\s*\d+\s*", "", caption).strip()
    if "ER" in title.upper() or "实体关系" in title or "数据库" in title:
        return (
            f"该图展示了{title}中主要实体、属性和实体间联系，能够直观说明用户、资源、分类、"
            "收藏、下载历史和审核记录等数据对象之间的结构关系，为后续逻辑表设计提供依据。"
        )
    return f"该图展示了{title}的主要界面和功能关系，能够辅助说明本节相关模块的实现方式。"


def first_body_paragraph_donor(paragraphs: list[ET.Element]) -> tuple[ET.Element | None, ET.Element | None]:
    body_started = False
    for paragraph in paragraphs:
        text = paragraph_text(paragraph).strip()
        if is_reference_heading(text):
            break
        if heading_level(text) == 1:
            body_started = True
            continue
        if not body_started:
            continue
        if not text or heading_level(text) is not None or is_figure_caption(text) or is_table_caption(text):
            continue
        if has_image(paragraph):
            continue
        ppr = paragraph.find("./w:pPr", NS)
        rpr = None
        for run in paragraph.findall("./w:r", NS):
            if paragraph_text(run).strip():
                rpr = run.find("./w:rPr", NS)
                break
        return deepcopy(ppr) if ppr is not None else None, deepcopy(rpr) if rpr is not None else None
    return None, None


def make_text_paragraph(text: str, donor_ppr: ET.Element | None, donor_rpr: ET.Element | None) -> ET.Element:
    paragraph = ET.Element(qn("p"))
    if donor_ppr is not None:
        paragraph.append(deepcopy(donor_ppr))
    run = ET.SubElement(paragraph, qn("r"))
    if donor_rpr is not None:
        run.append(deepcopy(donor_rpr))
    t = ET.SubElement(run, qn("t"))
    t.text = text
    return paragraph


def figure_number_key(text: str) -> str | None:
    match = re.search(r"图\s*(\d+)\s*[-－—]\s*(\d+)", text or "")
    if not match:
        return None
    return f"{match.group(1)}-{match.group(2)}"


def references_different_figure(current_caption: str, next_text: str) -> bool:
    current = figure_number_key(current_caption)
    if current is None:
        return False
    refs = [
        f"{match.group(1)}-{match.group(2)}"
        for match in re.finditer(r"图\s*(\d+)\s*[-－—]\s*(\d+)", next_text or "")
    ]
    return any(ref != current for ref in refs)


def repair_document(
    root: ET.Element,
    styles: dict[str, str],
    *,
    repair_figure_followups: bool = True,
    repair_chapter_pagination: bool = True,
) -> dict[str, object]:
    body = root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    donor_ppr, donor_rpr = first_body_paragraph_donor([node for node in list(body) if node.tag == qn("p")])
    inserted_followups: list[dict[str, object]] = []
    removed_empty_after_caption: list[int] = []

    if repair_figure_followups:
        idx = 0
        children = list(body)
        body_started = False
        while idx < len(children):
            child = children[idx]
            if child.tag != qn("p"):
                idx += 1
                continue
            text = paragraph_text(child).strip()
            if heading_level(text) == 1:
                body_started = True
            if not body_started or is_reference_heading(text) or not is_figure_caption(text):
                idx += 1
                continue

            next_idx = idx + 1
            if next_idx < len(children):
                next_child = children[next_idx]
                next_text = paragraph_text(next_child).strip() if next_child.tag == qn("p") else ""
                if next_child.tag == qn("p") and not next_text and not has_image(next_child) and not has_section_break(next_child):
                    body.remove(next_child)
                    removed_empty_after_caption.append(next_idx)
                    children.pop(next_idx)

            next_child = children[next_idx] if next_idx < len(children) else None
            next_text = paragraph_text(next_child).strip() if next_child is not None and next_child.tag == qn("p") else ""
            next_style_name = styles.get(paragraph_style_id(next_child), "") if next_child is not None and next_child.tag == qn("p") else ""
            needs_followup = (
                next_child is None
                or next_child.tag != qn("p")
                or not next_text
                or has_image(next_child)
                or is_figure_caption(next_text)
                or is_table_caption(next_text)
                or heading_level(next_text) is not None
                or next_style_name.lower().startswith("heading")
                or next_style_name.lower().startswith("toc")
                or references_different_figure(text, next_text)
            )
            if needs_followup:
                followup = make_text_paragraph(followup_text_for_caption(text), donor_ppr, donor_rpr)
                body.insert(next_idx, followup)
                children.insert(next_idx, followup)
                inserted_followups.append({"caption": text, "inserted_after_body_child_index": idx + 1})
                idx += 1
            idx += 1

    chapter_headings: list[tuple[int, ET.Element, str]] = []
    added_page_break_before: list[str] = []
    removed_empty_before_chapter: list[int] = []
    if repair_chapter_pagination:
        children = list(body)
        toc_seen = False
        for idx, child in enumerate(children):
            if child.tag != qn("p"):
                continue
            text = paragraph_text(child).strip()
            style_name = styles.get(paragraph_style_id(child), "")
            if is_toc_heading(text):
                toc_seen = True
                continue
            if not toc_seen:
                continue
            if "\t" in text or style_name.lower().startswith("toc") or looks_like_toc_entry(text):
                continue
            if is_reference_heading(text):
                break
            if heading_level(text) == 1:
                chapter_headings.append((idx, child, text))

    for idx, paragraph, text in chapter_headings:
        while idx > 0:
            previous = next((node for node in reversed(children[:idx]) if node.tag == qn("p")), None)
            if previous is None:
                break
            previous_index = children.index(previous)
            if (
                paragraph_text(previous).strip()
                or has_image(previous)
                or has_page_break(previous)
                or has_section_break(previous)
            ):
                break
            body.remove(previous)
            children.pop(previous_index)
            removed_empty_before_chapter.append(previous_index)
            idx = children.index(paragraph)
        previous = next((node for node in reversed(children[:idx]) if node.tag == qn("p")), None)
        has_owner = has_page_break_before(paragraph) or (
            previous is not None and (has_page_break(previous) or has_section_break(previous))
        )
        if not has_owner and add_page_break_before(paragraph):
            added_page_break_before.append(text)

    return {
        "inserted_figure_followups": inserted_followups,
        "removed_empty_after_caption": removed_empty_after_caption,
        "removed_empty_before_chapter": removed_empty_before_chapter,
        "chapter_page_break_before_added": added_page_break_before,
    }


def write_docx_with_document_xml(input_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    with tempfile.TemporaryDirectory() as tmp_name:
        tmp = Path(tmp_name)
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp)
        (tmp / "word" / "document.xml").write_bytes(document_xml)
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmp)
        zip_path = output_docx.with_suffix(".zip")
        zip_path.replace(output_docx)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--chapter-pagination-only", action="store_true")
    args = parser.parse_args()

    input_docx = Path(args.input_docx)
    output_docx = Path(args.output_docx)
    report = Path(args.report)
    with zipfile.ZipFile(input_docx, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        styles = paragraph_style_map(zf)
    changes = repair_document(
        root,
        styles,
        repair_figure_followups=not args.chapter_pagination_only,
        repair_chapter_pagination=True,
    )
    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    write_docx_with_document_xml(input_docx, output_docx, xml_bytes)
    payload = {
        "schema": "graduation-project-builder.figure-followup-chapter-pagination-repair.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_docx": str(input_docx),
        "input_docx_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_docx_sha256": sha256_file(output_docx),
        "chapter_pagination_only": bool(args.chapter_pagination_only),
        **changes,
    }
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
