#!/usr/bin/env python3
"""Rewrite Chinese/English thesis abstracts into triad paragraphs.

The helper is intentionally narrow: it mutates only ``word/document.xml`` and
only the paragraphs between the abstract headings and keyword lines. Existing
paragraph properties and run properties from the source abstract body/keyword
paragraphs are reused so the local template formatting remains stable.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "xml": "http://www.w3.org/XML/1998/namespace"}
W = "{%s}" % NS["w"]


ZH_ABSTRACT = [
    "随着在线旅游平台评论数量不断增加，评论文本逐渐成为理解游客体验、服务感知和消费反馈的重要数据来源。相比传统问卷或人工阅读方式，在线评论具有数量大、表达分散、情绪倾向不完全显性等特点，直接统计难以完整呈现评论中的情感分布和文本特征。因此，面向本地旅游评论样本构建可导入、可预处理、可计算和可视化的情感分析流程，能够为旅游评论文本研究提供更清晰的数据处理路径。",
    "针对上述问题，本文设计并实现了一个基于 Python 与 Qt 的景点情感分析系统，围绕评论导入、文本预处理、情感计算、结果可视化和 SQLite 本地存储构建完整分析流程。实验采用项目中的 ChnSentiCorp.csv 公开中文在线旅游与酒店评论语料，共获得有效样本 7765 条，其中正向评论 5322 条、负向评论 2443 条。系统利用 Jieba 分词、SnowNLP 情感概率和情感词典进行融合判断，并将情感分布、高频词和分组统计结果以图表形式展示。实验结果表明，系统能够完成公开评论样本的本地化处理、情感分类评价和可视化展示，相关分组结果反映的是样本内统计差异，不直接代表具体景点在真实平台上的总体口碑。",
    "本研究的意义在于将公开中文旅游评论语料的本地化处理、情感计算、样本统计和可视化展示整合为一个可操作的分析流程，为在线旅游评论的情感识别和文本特征观察提供参考。系统结果能够帮助研究者从样本层面了解评论情绪结构和高频表达特征，也为后续在更大规模数据、更多分类模型和更细粒度旅游场景中的情感分析研究提供基础。"
]

ZH_KEYWORDS = "关键词：旅游评论；情感分析；文本处理；可视化分析；Qt桌面系统"

EN_ABSTRACT = [
    "With the continuous growth of reviews on online travel platforms, review texts have become an important data source for understanding tourist experience, service perception, and consumer feedback. Compared with traditional questionnaires or manual reading, online reviews are large in volume, scattered in expression, and often implicit in sentiment tendency, which makes direct statistics insufficient for presenting sentiment distribution and textual characteristics. Therefore, building a sentiment analysis process that supports local data import, preprocessing, sentiment calculation, and visualization can provide a clearer data-processing path for tourism review analysis.",
    "To address these problems, this thesis designs and implements a scenic spot sentiment analysis system based on Python and Qt. The system builds an integrated workflow covering review import, text preprocessing, sentiment calculation, result visualization and local SQLite storage. The experiment uses the public ChnSentiCorp.csv Chinese online tourism and hotel review corpus in the project directory, containing 7,765 valid samples, including 5,322 positive reviews and 2,443 negative reviews. Jieba word segmentation, SnowNLP sentiment probability and a sentiment lexicon are combined to judge review polarity, and sentiment distribution, high-frequency words and grouped statistics are presented through charts. The results show that the system can process public review samples locally, evaluate sentiment classification and visualize the analysis results, while the grouped results only describe statistical differences within the sample and do not represent the overall reputation of specific attractions on real platforms.",
    "The significance of this study lies in integrating local processing, sentiment calculation, sample statistics, and visualization of public Chinese tourism review data into an operable analysis workflow. The results help observe sentiment structure and high-frequency textual features at the sample level, and provide a foundation for future sentiment analysis research involving larger datasets, more classification models, and more fine-grained tourism scenarios."
]

EN_KEYWORDS = "Key words: tourism reviews; sentiment analysis; text processing; visualization analysis; Qt desktop system"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def text_of(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS))


def compact(text: str) -> str:
    return "".join((text or "").split()).lower()


def paragraph_runs(p: etree._Element) -> list[etree._Element]:
    return [node for node in p if node.tag == W + "r"]


def clone_para_with_text(source: etree._Element, text: str) -> etree._Element:
    p = copy.deepcopy(source)
    for child in list(p):
        if child.tag != W + "pPr":
            p.remove(child)
    template_run = None
    for run in paragraph_runs(source):
        if text_of(run):
            template_run = run
            break
    if template_run is None:
        template_run = etree.Element(W + "r")
    run = copy.deepcopy(template_run)
    for child in list(run):
        if child.tag != W + "rPr":
            run.remove(child)
    t = etree.Element(W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{%s}space" % NS["xml"], "preserve")
    t.text = text
    run.append(t)
    p.append(run)
    return p


def set_para_text_in_place(p: etree._Element, text: str) -> None:
    new_p = clone_para_with_text(p, text)
    p.clear()
    p.attrib.update(new_p.attrib)
    for child in list(new_p):
        p.append(child)


def set_run_text_preserving_rpr(run: etree._Element, text: str) -> None:
    for child in list(run):
        if child.tag != W + "rPr":
            run.remove(child)
    if text:
        t = etree.Element(W + "t")
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t.set("{%s}space" % NS["xml"], "preserve")
        t.text = text
        run.append(t)


def split_keyword_line(text: str) -> tuple[str, str] | None:
    for marker in ("关键词：", "Key words:", "Keywords:", "Key words：", "Keywords："):
        if text.startswith(marker):
            return marker, text[len(marker) :]
    return None


def set_keyword_text_in_place(p: etree._Element, text: str) -> None:
    parts = split_keyword_line(text)
    text_runs = [run for run in paragraph_runs(p) if text_of(run)]
    if parts is None or len(text_runs) < 2:
        set_para_text_in_place(p, text)
        return
    label, content = parts
    set_run_text_preserving_rpr(text_runs[0], label)
    set_run_text_preserving_rpr(text_runs[1], content)
    for run in text_runs[2:]:
        set_run_text_preserving_rpr(run, "")


def find_para(paragraphs: list[etree._Element], predicate: Any, start: int = 0) -> int:
    for idx in range(start, len(paragraphs)):
        if predicate(compact(text_of(paragraphs[idx])), text_of(paragraphs[idx])):
            return idx
    raise RuntimeError("target paragraph not found")


def replace_range(body: etree._Element, paragraphs: list[etree._Element], start: int, end_exclusive: int, source_para: etree._Element, new_texts: list[str]) -> None:
    if end_exclusive <= start:
        raise RuntimeError("invalid replacement range")
    first = paragraphs[start]
    parent_children = list(body)
    body_indices = [parent_children.index(paragraphs[idx]) for idx in range(start, end_exclusive)]
    for body_idx in sorted(body_indices, reverse=True):
        body.remove(parent_children[body_idx])
    insert_at = body_indices[0]
    for offset, text in enumerate(new_texts):
        body.insert(insert_at + offset, clone_para_with_text(source_para, text))


def repair_docx(input_docx: Path, output_docx: Path) -> dict[str, Any]:
    with zipfile.ZipFile(input_docx, "r") as zf:
        original_xml = zf.read("word/document.xml")
    root = etree.fromstring(original_xml)
    body = root.find("w:body", namespaces=NS)
    if body is None:
        raise RuntimeError("missing w:body")
    paragraphs = [node for node in body if node.tag == W + "p"]

    zh_title = find_para(paragraphs, lambda c, t: c == "摘要")
    zh_kw = find_para(paragraphs, lambda c, t: c.startswith("关键词"), zh_title + 1)
    en_title = find_para(paragraphs, lambda c, t: c == "abstract", zh_kw + 1)
    en_kw = find_para(paragraphs, lambda c, t: c.startswith("keywords:") or c.startswith("keywords") or c.startswith("keywords：") or c.startswith("key words:"), en_title + 1)

    old = {
        "zh_body": [text_of(paragraphs[i]) for i in range(zh_title + 1, zh_kw)],
        "zh_keywords": text_of(paragraphs[zh_kw]),
        "en_body": [text_of(paragraphs[i]) for i in range(en_title + 1, en_kw)],
        "en_keywords": text_of(paragraphs[en_kw]),
    }

    replace_range(body, paragraphs, zh_title + 1, zh_kw, paragraphs[zh_title + 1], ZH_ABSTRACT)
    paragraphs = [node for node in body if node.tag == W + "p"]
    zh_kw = find_para(paragraphs, lambda c, t: c.startswith("关键词"), zh_title + 1)
    set_keyword_text_in_place(paragraphs[zh_kw], ZH_KEYWORDS)
    en_title = find_para(paragraphs, lambda c, t: c == "abstract", zh_kw + 1)
    set_para_text_in_place(paragraphs[en_title], "ABSTRACT")
    en_kw = find_para(paragraphs, lambda c, t: c.startswith("keywords:") or c.startswith("keywords") or c.startswith("keywords：") or c.startswith("key words:"), en_title + 1)
    replace_range(body, paragraphs, en_title + 1, en_kw, paragraphs[en_title + 1], EN_ABSTRACT)
    paragraphs = [node for node in body if node.tag == W + "p"]
    en_kw = find_para(paragraphs, lambda c, t: c.startswith("keywords:") or c.startswith("keywords") or c.startswith("keywords：") or c.startswith("key words:"), en_title + 1)
    set_keyword_text_in_place(paragraphs[en_kw], EN_KEYWORDS)

    modified_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(output_docx.parent)) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                data = modified_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
        shutil.move(str(tmp_path), str(output_docx))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()

    return {
        "schema": "graduation-project-builder.abstract-triad-repair.v1",
        "input_docx": str(input_docx),
        "input_sha256": sha256_file(input_docx),
        "output_docx": str(output_docx),
        "output_sha256": sha256_file(output_docx),
        "changed_parts": ["word/document.xml"],
        "old": old,
        "new": {
            "zh_body": ZH_ABSTRACT,
            "zh_keywords": ZH_KEYWORDS,
            "en_body": EN_ABSTRACT,
            "en_keywords": EN_KEYWORDS,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--audit-json", type=Path)
    args = parser.parse_args()
    report = repair_docx(args.input, args.output)
    if args.audit_json:
        args.audit_json.parent.mkdir(parents=True, exist_ok=True)
        args.audit_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
