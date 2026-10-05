#!/usr/bin/env python3
"""Read-only, cross-platform content signals from a SCAU thesis .docx.

Only the standard library is required. Findings concern the rule transcriptions
in this repository; they do not certify academic accuracy or school compliance.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from typing import Any
import xml.etree.ElementTree as ET
import zipfile


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SKILL_ROOT = Path(__file__).resolve().parent.parent
NS = {"w": W[1:-1]}
HAN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\U00020000-\U0002fa1f]")
EN_WORD = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
CN_KEYWORDS = re.compile(r"^\s*关键词\s*([:：])\s*(.*)$", re.DOTALL)
EN_KEYWORDS = re.compile(r"^\s*(Key\s+words|Keywords)\s*([:：])\s*(.*)$", re.I | re.DOTALL)
EN_ABSTRACT = re.compile(r"^\s*Abstract\s*[:：]\s*(.*)$", re.I | re.DOTALL)
CHAPTER = re.compile(r"^(\d+(?:\.\d+){0,3})\s+(.+)$")
CN_CHAPTER = re.compile(r"^(第\s*[0-9零〇一二三四五六七八九十百千两]+\s*章)\s*(.+)$")
HEADING_STYLE = re.compile(r"^(?:Heading|标题)\s*([1-4])$", re.I)
FUNCTION_WORDS = {"a", "an", "the", "and", "or", "of", "in", "on", "to", "for", "by", "with", "from", "at", "as", "into", "via"}


@dataclass
class Paragraph:
    index: int
    text: str
    style: str = ""
    outline_level: int | None = None
    in_table: bool = False
    in_toc: bool = False
    content_control_placeholder: bool = False

    def location(self) -> dict[str, Any]:
        return {"part": "word/document.xml", "paragraph": self.index,
                "in_table": self.in_table, "text_excerpt": self.text[:160]}


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text)


def explicit_heading(paragraph: Paragraph) -> bool:
    return (paragraph.outline_level is not None and 0 <= paragraph.outline_level <= 3) or bool(HEADING_STYLE.match(paragraph.style))


def body_heading(paragraph: Paragraph) -> tuple[str, str] | None:
    """A number alone is insufficient evidence that prose is a chapter heading."""
    match = CHAPTER.match(paragraph.text) or CN_CHAPTER.match(paragraph.text)
    if not match:
        return None
    number, title = match.groups()
    if not explicit_heading(paragraph):
        # These are parser safeguards, not school limits on chapter titles.
        if re.fullmatch(r"(?:19|20|21)\d{2}", number) or re.search(r"[，。；！？,.;!?\n]", title):
            return None
    return number, title


def source(comment_ids: list[int] | None = None, section: str | None = None,
           strength: str = "template_comment", file: str | None = None) -> dict[str, Any]:
    return {"file": file or ("references/scau-template-comments.md" if comment_ids else "references/format-rules.md"),
            "comment_ids": comment_ids or [], "section": section,
            "rule_strength": strength,
            "rule_verification": "repository_transcription_only"}


def _visible_nodes(element: ET.Element):
    """Do not let deleted text or nested text boxes act as document headings."""
    if element.tag in {W + "del", W + "moveFrom", W + "txbxContent"}:
        return
    yield element
    for child in element:
        yield from _visible_nodes(child)


def read_docx(input_path: Path) -> tuple[list[Paragraph], dict[str, Any]]:
    with zipfile.ZipFile(input_path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        styles_root = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else None
    parents = {child: parent for parent in document.iter() for child in parent}
    styles: dict[str, dict[str, Any]] = {}
    if styles_root is not None:
        for style in styles_root.findall("w:style", NS):
            name = style.find("w:name", NS)
            based = style.find("w:basedOn", NS)
            outline = style.find("w:pPr/w:outlineLvl", NS)
            styles[style.get(W + "styleId", "")] = {
                "name": name.get(W + "val", "") if name is not None else "",
                "based": based.get(W + "val", "") if based is not None else "",
                "outline": int(outline.get(W + "val", "9")) if outline is not None else None,
            }

    def style_properties(style_id: str) -> tuple[bool, int | None]:
        seen: set[str] = set()
        toc = False
        outline = None
        while style_id and style_id not in seen:
            seen.add(style_id)
            item = styles.get(style_id, {})
            name = item.get("name", "")
            toc = toc or bool(re.match(r"^(?:TOC|目录)\s*[1-9]$", style_id, re.I) or re.match(r"^(?:TOC|目录)\s*[1-9]$", name, re.I))
            if outline is None:
                outline = item.get("outline")
                heading_name = HEADING_STYLE.match(style_id) or HEADING_STYLE.match(name)
                if outline is None and heading_name:
                    outline = int(heading_name.group(1)) - 1
            style_id = item.get("based", "")
        return toc, outline

    paragraphs: list[Paragraph] = []
    field_stack: list[dict[str, Any]] = []
    excluded_nested = 0
    for index, element in enumerate(document.findall(".//w:body//w:p", NS), 1):
        ancestors = []
        current = parents.get(element)
        while current is not None:
            ancestors.append(current)
            current = parents.get(current)
        if any(item.tag in {W + "del", W + "moveFrom", W + "txbxContent"} for item in ancestors):
            excluded_nested += 1
            continue
        style_node = element.find("w:pPr/w:pStyle", NS)
        style_id = style_node.get(W + "val", "") if style_node is not None else ""
        style_toc, outline = style_properties(style_id)
        direct_outline = element.find("w:pPr/w:outlineLvl", NS)
        if direct_outline is not None:
            outline = int(direct_outline.get(W + "val", "9"))
        sdt_toc = False
        placeholder = False
        for ancestor in ancestors:
            if ancestor.tag != W + "sdt":
                continue
            gallery = ancestor.find("w:sdtPr/w:docPartObj/w:docPartGallery", NS)
            if gallery is not None:
                sdt_toc = sdt_toc or gallery.get(W + "val", "").lower() in {"table of contents", "目录"}
            placeholder = placeholder or ancestor.find("w:sdtPr/w:showingPlcHdr", NS) is not None
        text_parts = []
        field_toc = any(field["toc"] for field in field_stack)
        for node in _visible_nodes(element):
            if node.tag == W + "fldSimple":
                field_toc = field_toc or bool(re.search(r"\bTOC\b", node.get(W + "instr", ""), re.I))
            elif node.tag == W + "fldChar":
                kind = node.get(W + "fldCharType", "")
                if kind == "begin":
                    field_stack.append({"toc": False, "instruction": ""})
                elif kind == "end" and field_stack:
                    field_toc = field_toc or field_stack[-1]["toc"]
                    field_stack.pop()
            elif node.tag == W + "instrText" and field_stack:
                field_stack[-1]["instruction"] += node.text or ""
                field_stack[-1]["toc"] = bool(re.search(r"\bTOC\b", field_stack[-1]["instruction"], re.I))
                field_toc = field_toc or field_stack[-1]["toc"]
            elif node.tag == W + "t":
                text_parts.append(node.text or "")
            elif node.tag == W + "tab":
                text_parts.append("\t")
            elif node.tag in {W + "br", W + "cr"}:
                text_parts.append("\n")
        text = "".join(text_parts).strip()
        # Plain text contents pasted from another document often have no TOC style.
        # Explicit heading metadata outweighs a terminal tab/page-number heuristic.
        # Actual TOC styles, containers and fields still take precedence.
        heading_metadata = (outline is not None and 0 <= outline <= 3) or bool(HEADING_STYLE.match(style_id))
        plain_toc = not heading_metadata and bool(re.search(r"(?:\t|\.{2,}|…{2,})\s*[0-9ivxlcdm]+\s*$", text, re.I))
        paragraphs.append(Paragraph(index, text, style_id, outline,
                                    any(a.tag == W + "tbl" for a in ancestors),
                                    style_toc or sdt_toc or field_toc or plain_toc, placeholder))
    return paragraphs, {
        "paragraph_index_basis": "1-based XML document paragraph order, including table cells; not Word page numbers",
        "excluded_nested_paragraphs": excluded_nested,
        "tracked_change_elements": sum(len(document.findall(".//w:" + tag, NS)) for tag in ("ins", "del", "moveFrom", "moveTo")),
        "unclosed_fields": len(field_stack),
    }


def section_key(paragraph: Paragraph) -> str | None:
    if paragraph.in_toc or not paragraph.text:
        return None
    text = compact(paragraph.text)
    if text in {"本科毕业论文", "本科毕业设计", "本科毕业论文(或设计)", "本科毕业论文（或设计）"}:
        return "cover"
    if re.fullmatch(r"(?:华南农业大学)?(?:本科毕业论文(?:[（(]设计[）)])?)?原创性声明", text):
        return "originality_statement"
    if re.fullmatch(r"(?:华南农业大学)?(?:本科毕业论文(?:[（(]设计[）)])?)?(?:论文)?使用授权声明", text):
        return "authorization_statement"
    if paragraph.in_table:
        return None
    if text == "摘要":
        return "chinese_abstract"
    if EN_ABSTRACT.match(paragraph.text):
        return "english_abstract"
    if text == "目录":
        return "contents"
    if text in {"英文缩略词", "英文缩略词（符号表）", "英文缩略词(符号表)"}:
        return "abbreviation_list"
    if text == "参考文献":
        return "references"
    if text == "致谢":
        return "acknowledgements"
    if re.fullmatch(r"附录[A-ZＡ-Ｚ](?:.+)?", text):
        return "appendix"
    if body_heading(paragraph):
        return "body"
    return None


def inspect_provenance(skill_root: Path) -> dict[str, Any]:
    manifest_path = skill_root / "assets/official-2024/manifest.json"
    files = []
    if manifest_path.is_file():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        for item in manifest.get("required_files", []):
            path = manifest_path.parent / item["filename"]
            status = "not_imported"
            actual_hash = None
            if path.is_file():
                digest = hashlib.sha256()
                with path.open("rb") as handle:
                    for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                        digest.update(chunk)
                actual_hash = digest.hexdigest().upper()
                status = "hash_matches_manifest" if actual_hash == item.get("sha256", "").upper() else "hash_mismatch"
            files.append({"filename": item["filename"], "status": status, "observed_sha256": actual_hash})
    verified = bool(files) and all(item["status"] == "hash_matches_manifest" for item in files)
    return {"rule_basis": "repository rule and comment transcriptions",
            "rule_verification": "repository_transcription_only",
            "official_package_status": "hash_matches_manifest" if verified else "not_fully_verified",
            "official_files": files,
            "note": "文件哈希匹配只验证来源文件身份；本脚本不自动重新核对 PDF/DOC 原文，也不确认部门或导师后续要求。"}


def inspect_content(input_path: Path, skill_root: Path = SKILL_ROOT) -> dict[str, Any]:
    input_path = Path(input_path).expanduser().resolve()
    if input_path.suffix.lower() != ".docx":
        raise ValueError("Content inspection accepts .docx only; convert legacy .doc with Word first.")
    paragraphs, extraction = read_docx(input_path)
    findings: list[dict[str, Any]] = []

    def add(rule_id: str, status: str, severity: str, category: str,
            location: Any, observed: Any, expected: Any, rule_source: dict[str, Any], note: str = ""):
        findings.append({"rule_id": rule_id, "status": status, "severity": severity,
                         "category": category, "source": rule_source, "location": location,
                         "observed": observed, "expected": expected, "note": note})

    locations: dict[str, list[Paragraph]] = {}
    active_abstract = None
    for paragraph in paragraphs:
        key = section_key(paragraph)
        if active_abstract and not paragraph.in_toc:
            keyword_pattern = CN_KEYWORDS if active_abstract == "chinese_abstract" else EN_KEYWORDS
            if keyword_pattern.match(paragraph.text):
                active_abstract = None
            elif key == "body" and not explicit_heading(paragraph):
                # Normal-style numeric sentences/lists inside an abstract are prose.
                key = None
            elif key:
                active_abstract = None
        if key:
            locations.setdefault(key, []).append(paragraph)
            if key in {"chinese_abstract", "english_abstract"}:
                active_abstract = key
    required = {"cover", "originality_statement", "authorization_statement", "chinese_abstract", "english_abstract", "contents", "body"}
    ordered_sections = ["cover", "originality_statement", "authorization_statement", "chinese_abstract", "english_abstract", "abbreviation_list", "contents", "body", "references", "appendix", "acknowledgements"]
    sections = {}
    for key in ordered_sections:
        detected = locations.get(key, [])
        scope = "expected_template_structure" if key in required else "optional_or_template_module"
        sections[key] = {"status": "detected" if detected else "not_detected", "scope": scope,
                         "locations": [item.location() for item in detected],
                         "note": "检测到标签不表示该模块内容完整、学术有效或格式合规。"}
        if not detected and key in required:
            add("structure." + key, "manual_confirm", "warning", "structure", {"part": "word/document.xml"},
                "未检测到独立标题或标签", "核对该模板默认模块是否适用于本稿，并确认存在及命名/位置", source(section="§5", strength="template_example"),
                "仅按可见正文检测；目录条目不能替代正文模块，非标准标题和文本框需人工确认。")
        if len(detected) > 1 and key not in {"body", "appendix"}:
            add("structure.duplicate." + key, "manual_confirm", "warning", "structure",
                [item.location() for item in detected], len(detected), "核对重复标签；本次统计采用首个候选", source(section="§5", strength="template_example"))

    def abstract_parts(key: str, keyword_pattern: re.Pattern) -> tuple[list[Paragraph], list[str], Paragraph | None, str]:
        candidates = locations.get(key, [])
        if not candidates:
            return [], [], None, "not_detected"
        start = candidates[0]
        selected: list[Paragraph] = []
        texts: list[str] = []
        keyword = None
        boundary = "section_boundary_without_keyword"
        if key == "english_abstract":
            body = EN_ABSTRACT.match(start.text).group(1).strip()
            if body:
                selected.append(start)
                texts.append(body)
        after_start = False
        for paragraph in paragraphs:
            if paragraph.index == start.index:
                after_start = True
                continue
            if not after_start or paragraph.in_toc:
                continue
            if keyword_pattern.match(paragraph.text):
                keyword = paragraph
                boundary = "keyword_label"
                break
            boundary_key = section_key(paragraph)
            if boundary_key and (boundary_key != "body" or explicit_heading(paragraph)):
                break
            if paragraph.text:
                selected.append(paragraph)
                texts.append(paragraph.text)
        if keyword is None:
            add("abstract.keyword_label_missing." + key, "manual_confirm", "warning", "abstract",
                start.location(), "未在该摘要边界前找到关键词标签", "中文 关键词： / 英文 Key words:", source([30 if key == "chinese_abstract" else 35]),
                "摘要计数可能混入下一部分的题目、作者或单位；须先确认边界。")
        return selected, texts, keyword, boundary

    abstract_stats = {}
    keyword_stats = {}
    for key, pattern in (("chinese_abstract", CN_KEYWORDS), ("english_abstract", EN_KEYWORDS)):
        selected, texts, keywords, boundary = abstract_parts(key, pattern)
        body_text = "\n".join(texts)
        stats = {"han_characters": len(HAN.findall(body_text)),
                 "visible_characters_excluding_whitespace": len(re.sub(r"\s", "", body_text)),
                 "english_word_tokens": len(EN_WORD.findall(body_text)),
                 "body_paragraphs": [p.index for p in selected], "boundary": boundary,
                 "count_basis": "摘要正文候选，不含摘要标签和关键词；Han仅汉字，visible含标点、数字和西文；均不是 Word 字数。"}
        abstract_stats[key] = stats
        if key in locations and not body_text.strip():
            certain_empty = boundary == "keyword_label" and len(locations[key]) == 1
            add("abstract.empty." + key, "confirmed" if certain_empty else "manual_confirm",
                "error" if certain_empty else "warning", "content_completion",
                locations[key][0].location(), "摘要标签后未检测到正文",
                "补入经过用户确认的摘要正文" if certain_empty else "先确认摘要边界及实际正文是否存在",
                source([29 if key == "chinese_abstract" else 34]),
                "只有独立摘要标签到真实关键词标签之间为空时，才确认空摘要；缺边界或重复标签均需人工确认。")
        if key == "chinese_abstract" and body_text:
            visible_count = stats["visible_characters_excluding_whitespace"]
            stats["recorded_range"] = [300, 600]
            stats["range_status"] = "within_recorded_range" if 300 <= visible_count <= 600 else "manual_confirm"
            if not 300 <= visible_count <= 600:
                add("abstract.cn_length", "manual_confirm", "warning", "abstract", [p.location() for p in selected],
                    {"han_characters": stats["han_characters"], "visible_characters": visible_count},
                    "批注 C29 记录为 300 至 600 字；须核对学校统计口径", source([29]),
                    "此处用非空白可见字符筛查，不能将汉字数或此计数冒充学校/Word 字数。")
            citation_pattern = re.compile(r"[（(][^（）()\n]{0,100}(?:19|20)\d{2}[a-z]?[^（）()\n]{0,30}[）)]|\[\d+(?:\s*[-,，]\s*\d+)*\]")
            for paragraph, paragraph_text in zip(selected, texts):
                candidates = citation_pattern.findall(paragraph_text)
                if candidates:
                    add("abstract.cn_citation_candidate", "manual_confirm", "warning", "abstract", paragraph.location(),
                        candidates, "批注 C29 不引用参考文献；核对候选是否确属引用", source([29]),
                        "年份、区间、实验编号可能匹配此模式；不会自动删除。")
        if keywords is not None:
            match = pattern.match(keywords.text)
            content = match.group(2) if key == "chinese_abstract" else match.group(3)
            delimiter = "；" if key == "chinese_abstract" else ";"
            keyword_list = [item.strip() for item in re.split(r"[;；]", content) if item.strip()]
            ambiguous = not re.search(r"[;；]", content) and bool(re.search(r"[,，、]", content))
            keyword_stats[key] = {"items": keyword_list, "count": len(keyword_list),
                                  "count_status": "manual_confirm" if ambiguous else "measured",
                                  "location": keywords.location()}
            cids = [30] if key == "chinese_abstract" else [35]
            wrong = ";" if delimiter == "；" else "；"
            if wrong in content:
                add("keywords.separator." + key, "confirmed", "error", "keywords", keywords.location(),
                    content, "关键词间用" + ("全角分号 ；" if delimiter == "；" else "半角分号 ;"), source(cids))
            if content.rstrip().endswith((";", "；", ",", "，", ".", "。", "、", ":", "：")):
                add("keywords.trailing_punctuation." + key, "confirmed", "error", "keywords", keywords.location(),
                    content[-30:], "末尾不加标点", source(cids))
            if not content.strip():
                add("keywords.empty." + key, "confirmed", "error", "content_completion", keywords.location(),
                    "关键词内容为空", "补入经用户确认的关键词", source(cids))
            elif ambiguous:
                add("keywords.ambiguous_separator." + key, "manual_confirm", "warning", "keywords", keywords.location(),
                    content, "用分号明确关键词边界后再计数", source(cids), "逗号也可能属于一个关键词内部。")
            elif key == "chinese_abstract" and not 3 <= len(keyword_list) <= 5:
                add("keywords.cn_count", "confirmed", "error", "keywords", keywords.location(),
                    len(keyword_list), "3 至 5 个中文关键词", source([30]))
            if key == "english_abstract":
                lower_words = [word for word in EN_WORD.findall(content) if word[0].islower() and word.lower() not in FUNCTION_WORDS]
                if lower_words:
                    add("keywords.en_capitalization", "manual_confirm", "warning", "keywords", keywords.location(),
                        lower_words, "批注 C35：实词首字母大写；核对学名、缩写等例外", source([35]), "不会自动改写关键词或学名。")
                if match.group(1).lower() != "key words" or match.group(2) != ":":
                    add("keywords.en_label_variant", "manual_confirm", "info", "keywords", keywords.location(),
                        match.group(1) + match.group(2), "模板标签 Key words:", source([35]))

    english_title_candidate = None
    cn_keyword = keyword_stats.get("chinese_abstract", {}).get("location", {}).get("paragraph")
    en_start = locations.get("english_abstract", [])
    if cn_keyword is not None and en_start:
        candidates = [p for p in paragraphs if cn_keyword < p.index < en_start[0].index
                      and p.text and not p.in_toc and not p.in_table and section_key(p) is None]
        if candidates:
            title = candidates[0]
            english_title_candidate = {"text": title.text, "location": title.location(), "status": "manual_confirm"}
            lower_words = [word for word in EN_WORD.findall(title.text)
                           if word[0].islower() and word.lower() not in FUNCTION_WORDS]
            if lower_words:
                add("title.en_capitalization_candidate", "manual_confirm", "warning", "title", title.location(),
                    lower_words, "批注 C31：英文题目实词首字母大写；先确认该段是否为题目及专有名词例外", source([31]),
                    "通过中文关键词到 Abstract 的首段推测题目，不设置没有来源的标题长度限制，不自动改写。")

    # Report a candidate, never erase a real author or research identifier.
    residuals = [("repeated_template_text", re.compile(r"文本文本|标题标题")),
                 ("template_marker", re.compile(r"\bXXXX\b|\bEnglish Title\b|\bSong Nianxiu\b|\bText text\b", re.I))]
    for paragraph in paragraphs:
        if paragraph.in_toc:
            continue
        for marker, pattern in residuals:
            matches = pattern.findall(paragraph.text)
            if matches:
                add("template." + marker, "manual_confirm", "warning", "template_placeholder", paragraph.location(),
                    matches, "确认是否为模板残留，替换必须有用户提供的真实内容", source(section="§5", strength="workflow_default"),
                    "命中字符串只是候选；姓名、编码或引用原句可能合法，不能据此认定学术内容错误。")
        if paragraph.content_control_placeholder:
            add("template.content_control", "manual_confirm", "warning", "template_placeholder", paragraph.location(),
                "内容控件标记为 showingPlcHdr", "核对控件是否仍为占位提示", source(section="§5", strength="workflow_default"))

    body_headings = [p for p in locations.get("body", []) if body_heading(p)[0].count(".") == 0]
    if body_headings:
        heading_texts = [body_heading(p)[1] for p in body_headings]
        missing_suggestions = [label for label, alternatives in (
            ("前言", ("前言", "绪论", "引言")), ("材料与方法", ("材料", "方法")),
            ("结果与分析", ("结果",)), ("讨论与结论", ("讨论", "结论")))
            if not any(any(term in text for term in alternatives) for text in heading_texts)]
        if missing_suggestions:
            add("body.recommended_science_structure", "suggested", "info", "structure",
                [p.location() for p in body_headings], heading_texts, {"recommended_modules_not_matched": missing_suggestions},
                source(section="§4", strength="template_example"),
                "学校给出农科/理科参考结构；专业或研究类型可另有章节，不据此判定缺章。")

    if "chinese_abstract" in locations and "english_abstract" in locations:
        add("abstract.bilingual_consistency", "manual_confirm", "warning", "academic_content",
            [locations[k][0].location() for k in ("chinese_abstract", "english_abstract")],
            {"keyword_counts": {k: v["count"] for k, v in keyword_stats.items()}},
            "人工逐项核对中英文目的、方法、结果、结论及关键词含义一致", source(section="§3", strength="manual", file="references/content-audit.md"),
            "字数或关键词数量相同不等于翻译一致；本脚本不验证实验数据、引用真实性、论证或翻译。")
    status_counts = Counter(item["status"] for item in findings)
    return {"schema_version": 1, "file": str(input_path), "read_only": True,
            "judgement_basis": {"docx_text_structure": "observed", "rendered_layout": "manual_confirm",
                                "school_compliance": "not_determined", "academic_accuracy": "not_determined",
                                "note": "confirmed 仅表示可见文本与已转录规则的具体不符，不等于学校原文件已复验。"},
            "rule_provenance": inspect_provenance(skill_root), "extraction": extraction,
            "sections": sections, "abstract_stats": abstract_stats, "keyword_stats": keyword_stats,
            "english_title_candidate": english_title_candidate,
            "findings": findings, "summary": {"confirmed_issues": status_counts["confirmed"],
                                               "manual_confirm": status_counts["manual_confirm"],
                                               "suggested": status_counts["suggested"]},
            "limitations": ["只读 word/document.xml 可见正文，排除目录、删除文本和文本框；未检查页眉页脚、脚注、图片或渲染页。",
                            "非标准标题、纯文本目录和混合摘要段落可能影响边界，检测到标签不代表模块完成。",
                            "模板中的 optional/back matter 占位只列候选，不自动判为硬错误。",
                            "须结合官方原文件、部门/导师要求、Word 格式检查和逐页预览；不得宣称全文符合学校。",
                            "修订未接受/拒绝时统计使用当前插入文本并排除删除文本，须由用户确认最终修订状态。"]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_path", help="Input .docx; never modified")
    parser.add_argument("--output", help="Optional JSON report path; defaults to stdout")
    args = parser.parse_args()
    try:
        input_path = Path(args.input_path).expanduser().resolve()
        output = Path(args.output).expanduser().resolve() if args.output else None
        if output == input_path or (output is not None and output.exists() and input_path.exists()
                                   and os.path.samefile(output, input_path)):
            raise ValueError("Report output must not overwrite the input document.")
        report = inspect_content(input_path)
        payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        if output:
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(payload, encoding="utf-8")
        else:
            print(payload, end="")
        return 0
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
