#!/usr/bin/env python
"""Collect Word thesis statistics and structure hints for format-audit reports."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


PLACEHOLDER_RULES = {
    "common": [
        r"文本文本",
        r"\bXXXX\b",
        r"English Title",
        r"Song Nianxiu",
        r"Text text",
        r"\b标题标题\b",
    ],
    "references": [
        r"陈爱东",
        r"尼葛洛庞蒂",
        r"Yang Y, Chen F",
        r"Rafae1 C G",
    ],
    "abbreviation_list": [
        r"英文缩写",
        r"英文全称",
        r"中文全称",
    ],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect a Word thesis file and emit a JSON report summary."
    )
    parser.add_argument("input_path", help="Path to the input .doc or .docx file")
    parser.add_argument(
        "--output",
        dest="output_path",
        help="Optional JSON output path; prints to stdout when omitted",
    )
    return parser.parse_args()


def require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("This script requires Windows and Microsoft Word.")


def resolve_input(path_arg: str) -> Path:
    input_path = Path(path_arg).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Input file must have a .doc or .docx extension.")
    return input_path


def normalize_text(text: str) -> str:
    return (
        text.replace("\r", "\n")
        .replace("\x07", "\n")
        .replace("\x0b", "\n")
        .replace("\f", "\n")
    )


def first_match(text: str, pattern: str) -> re.Match[str] | None:
    return re.search(pattern, text, flags=re.MULTILINE)


def section_slice(text: str, start_pattern: str, end_patterns: list[str]) -> str:
    start = first_match(text, start_pattern)
    if not start:
        return ""
    end_pos = len(text)
    for pat in end_patterns:
        match = re.search(pat, text[start.end() :], flags=re.MULTILINE)
        if match:
            end_pos = min(end_pos, start.end() + match.start())
    return text[start.end() : end_pos].strip()


def count_cn_chars(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def count_en_words(text: str) -> int:
    return len(re.findall(r"\b[A-Za-z][A-Za-z-]*\b", text))


def infer_heading_level(text: str, outline_level: int) -> str | None:
    numbered = re.match(r"^(\d+(?:\.\d+)*)\s+", text)
    if numbered:
        depth = numbered.group(1).count(".") + 1
        if 1 <= depth <= 4:
            return str(depth)
    if re.match(r"^(参\s*考\s*文\s*献|致\s*谢|附录\s*[A-ZＡ-Ｚ])", text):
        return "1"
    if 1 <= outline_level <= 4:
        return str(outline_level)
    return None


def collect_heading_info(document) -> tuple[dict[str, int], list[dict[str, str]]]:
    heading_counts = {"1": 0, "2": 0, "3": 0, "4": 0}
    heading_samples: list[dict[str, str]] = []
    for para in document.Paragraphs:
        level = int(para.OutlineLevel)
        if 1 <= level <= 4:
            text = para.Range.Text.strip().replace("\r", " ")
            if not text:
                continue
            key = infer_heading_level(text, level)
            if key is None:
                continue
            heading_counts[key] += 1
            heading_samples.append({"level": key, "text": text})
    return heading_counts, heading_samples


def collect_caption_counts(paragraphs) -> dict[str, int]:
    counts = {
        "figure_captions": 0,
        "table_captions": 0,
        "continued_tables": 0,
        "formula_labels": 0,
    }
    for para in paragraphs:
        text = para.Range.Text.strip().replace("\r", " ")
        if not text:
            continue
        if re.match(r"^图\s*[A-Z]?\d+(?:-\d+)?\s+", text):
            counts["figure_captions"] += 1
        if re.match(r"^表\s*[A-Z]?\d+(?:-\d+)?\s+", text):
            counts["table_captions"] += 1
        if re.match(r"^续表\s*[A-Z]?\d+(?:-\d+)?\s+", text):
            counts["continued_tables"] += 1
        if re.search(r"（式\s*[A-Z]?\d+(?:-\d+)?）", text):
            counts["formula_labels"] += 1
    return counts


def collect_reference_entry_count(paragraphs) -> int:
    in_references = False
    count = 0
    for para in paragraphs:
        text = para.Range.Text.strip().replace("\r", " ")
        if not text:
            continue
        if re.match(r"^参\s*考\s*文\s*献$", text):
            in_references = True
            continue
        if in_references and re.match(r"^(附录\s*[A-ZＡ-Ｚ]|致\s*谢)$", text):
            break
        if in_references:
            count += 1
    return count


def collect_section_presence(text: str) -> dict[str, bool]:
    checks = {
        "cover": r"本科毕业论文|本科毕业设计",
        "originality_statement": r"原创性声明",
        "authorization_statement": r"使用授权声明",
        "chinese_abstract": r"摘\s*要",
        "english_abstract": r"\bAbstract:",
        "abbreviation_list": r"英文缩略词",
        "contents": r"目\s*录",
        "body": r"^\d+\s+",
        "references": r"参\s*考\s*文\s*献",
        "appendix": r"附录\s*[A-ZＡ-Ｚ]",
        "acknowledgements": r"致\s*谢",
    }
    return {key: bool(re.search(pattern, text, flags=re.MULTILINE)) for key, pattern in checks.items()}


def collect_section_texts(text: str) -> dict[str, str]:
    return {
        "chinese_abstract": section_slice(
            text,
            r"(?m)^\s*摘\s*要\s*$",
            [r"(?m)^\s*关键词[:：]", r"(?m)^\s*Abstract:", r"(?m)^\s*English Title", r"(?m)^\s*目\s*录\s*$"],
        ),
        "english_abstract": section_slice(
            text,
            r"(?m)^\s*Abstract:",
            [r"(?m)^\s*Key words:", r"(?m)^\s*英文缩略词", r"(?m)^\s*目\s*录\s*$"],
        ),
        "abbreviation_list": section_slice(
            text,
            r"(?m)^\s*英文缩略词.*$",
            [r"(?m)^\s*目\s*录\s*$", r"(?m)^\d+\s+"],
        ),
        "references": section_slice(
            text,
            r"(?m)^\s*参\s*考\s*文\s*献\s*$",
            [r"(?m)^\s*附录\s*[A-ZＡ-Ｚ].*$", r"(?m)^\s*致\s*谢\s*$"],
        ),
        "appendix": section_slice(
            text,
            r"(?m)^\s*附录\s*[A-ZＡ-Ｚ].*$",
            [r"(?m)^\s*致\s*谢\s*$"],
        ),
        "acknowledgements": section_slice(
            text,
            r"(?m)^\s*致\s*谢\s*$",
            [],
        ),
    }


def has_placeholder(text: str, extra_patterns: list[str] | None = None) -> bool:
    if not text:
        return False
    patterns = list(PLACEHOLDER_RULES["common"])
    if extra_patterns:
        patterns.extend(extra_patterns)
    return any(re.search(pattern, text, flags=re.MULTILINE) for pattern in patterns)


def build_section_audit(
    presence: dict[str, bool], section_texts: dict[str, str]
) -> dict[str, dict[str, str]]:
    scopes = {
        "cover": "hard_required",
        "originality_statement": "hard_required",
        "authorization_statement": "hard_required",
        "chinese_abstract": "hard_required",
        "english_abstract": "hard_required",
        "abbreviation_list": "optional_template_module",
        "contents": "hard_required",
        "body": "hard_required",
        "references": "template_module",
        "appendix": "optional_template_module",
        "acknowledgements": "template_module",
    }

    audits: dict[str, dict[str, str]] = {}
    for key, scope in scopes.items():
        if not presence.get(key):
            audits[key] = {
                "status": "missing",
                "scope": scope,
                "note": "未检测到该模块。",
            }
            continue

        extra_patterns = PLACEHOLDER_RULES.get(key, None)
        if has_placeholder(section_texts.get(key, ""), extra_patterns):
            audits[key] = {
                "status": "template_placeholder",
                "scope": scope,
                "note": "保留模板占位 / 本轮未处理。",
            }
        else:
            audits[key] = {
                "status": "confirmed",
                "scope": scope,
                "note": "已检测到非占位内容。",
            }
    return audits


def inspect_document(input_path: Path) -> dict[str, object]:
    try:
        import pythoncom
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Word inspection. Install it with `pip install pywin32`."
        ) from exc

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(input_path),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
            Revert=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )

        wd_stat_words = 0
        wd_stat_lines = 1
        wd_stat_pages = 2
        wd_stat_characters = 3
        wd_stat_paragraphs = 4
        wd_stat_characters_with_spaces = 5
        wd_stat_far_east_characters = 6

        full_text = normalize_text(document.Content.Text)
        section_presence = collect_section_presence(full_text)
        section_texts = collect_section_texts(full_text)
        section_audit = build_section_audit(section_presence, section_texts)
        heading_counts, heading_samples = collect_heading_info(document)
        caption_counts = collect_caption_counts(document.Paragraphs)

        report = {
            "file": str(input_path),
            "source_type": input_path.suffix.lower(),
            "judgement_basis": {
                "word_structure_stats": "confirmed",
                "rendered_layout": "manual_confirm",
                "note": "本脚本只提供基础统计、结构信号和模板占位识别，不能替代最终版式判定。",
            },
            "word_statistics": {
                "pages": int(document.ComputeStatistics(wd_stat_pages)),
                "words": int(document.ComputeStatistics(wd_stat_words)),
                "lines": int(document.ComputeStatistics(wd_stat_lines)),
                "paragraphs": int(document.ComputeStatistics(wd_stat_paragraphs)),
                "characters": int(document.ComputeStatistics(wd_stat_characters)),
                "characters_with_spaces": int(
                    document.ComputeStatistics(wd_stat_characters_with_spaces)
                ),
                "far_east_characters": int(
                    document.ComputeStatistics(wd_stat_far_east_characters)
                ),
            },
            "counts": {
                "tables": int(document.Tables.Count),
                "footnotes": int(document.Footnotes.Count),
                "endnotes": int(document.Endnotes.Count),
                "comments": int(document.Comments.Count),
                "toc_tables": int(document.TablesOfContents.Count),
                "references_estimated": collect_reference_entry_count(document.Paragraphs),
                **caption_counts,
            },
            "sections": section_presence,
            "section_audit": section_audit,
            "heading_counts": heading_counts,
            "heading_samples": heading_samples[:30],
            "abstract_stats": {
                "chinese_abstract_cn_chars": count_cn_chars(section_texts["chinese_abstract"]),
                "english_abstract_en_words": count_en_words(section_texts["english_abstract"]),
            },
        }
        return report
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def main() -> int:
    try:
        args = parse_args()
        require_windows()
        input_path = resolve_input(args.input_path)
        report = inspect_document(input_path)
        serialized = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output_path:
            output_path = Path(args.output_path).expanduser().resolve()
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(serialized, encoding="utf-8")
        else:
            print(serialized)
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
