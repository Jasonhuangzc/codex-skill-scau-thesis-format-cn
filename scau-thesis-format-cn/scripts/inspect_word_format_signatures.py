#!/usr/bin/env python
"""Inspect Word character-level format signatures for high-risk front-matter items."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect character-level font and bold signatures in a Word thesis."
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
    return text.replace("\r", "").replace("\x07", "").strip()


def is_bold(value) -> bool:
    return value not in (0, False, None, "")


def normalize_font_name(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def size_matches(observed: object, expected: float, tolerance: float = 0.6) -> bool:
    if observed is None:
        return False
    try:
        return abs(float(observed) - expected) <= tolerance
    except (TypeError, ValueError):
        return False


def char_signature(char_range) -> dict[str, object]:
    font = char_range.Font
    size = float(font.Size) if font.Size else None
    return {
        "bold": is_bold(font.Bold),
        "ascii_font": normalize_font_name(font.NameAscii or font.Name),
        "far_east_font": normalize_font_name(font.NameFarEast or font.Name),
        "size": round(size, 1) if size is not None else None,
    }


def collect_segments(paragraph) -> list[dict[str, object]]:
    segments: list[dict[str, object]] = []
    current: dict[str, object] | None = None
    for index in range(1, paragraph.Range.Characters.Count + 1):
        char_range = paragraph.Range.Characters(index)
        char_text = char_range.Text
        if char_text in {"\r", "\x07"}:
            continue
        signature = char_signature(char_range)
        if current is None or any(current[key] != signature[key] for key in ("bold", "ascii_font", "far_east_font", "size")):
            current = {
                "text": char_text,
                "bold": signature["bold"],
                "ascii_font": signature["ascii_font"],
                "far_east_font": signature["far_east_font"],
                "size": signature["size"],
            }
            segments.append(current)
        else:
            current["text"] = str(current["text"]) + char_text
    return segments


def non_empty_paragraphs(document) -> list[dict[str, object]]:
    items: list[dict[str, object]] = []
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs(index)
        text = normalize_text(paragraph.Range.Text)
        if not text:
            continue
        items.append(
            {
                "index": index,
                "text": text,
                "paragraph": paragraph,
            }
        )
    return items


def find_first(paragraphs: list[dict[str, object]], pattern: str) -> dict[str, object] | None:
    regex = re.compile(pattern)
    for item in paragraphs:
        if regex.search(str(item["text"])):
            return item
    return None


def previous_items(paragraphs: list[dict[str, object]], item: dict[str, object], count: int) -> list[dict[str, object]]:
    position = paragraphs.index(item)
    start = max(0, position - count)
    return paragraphs[start:position]


def next_nonempty_after(paragraphs: list[dict[str, object]], item: dict[str, object] | None) -> dict[str, object] | None:
    if item is None:
        return None
    position = paragraphs.index(item)
    for candidate in paragraphs[position + 1 :]:
        if str(candidate["text"]).strip():
            return candidate
    return None


def next_body_paragraph_after(paragraphs: list[dict[str, object]], item: dict[str, object] | None) -> dict[str, object] | None:
    if item is None:
        return None
    position = paragraphs.index(item)
    for candidate in paragraphs[position + 1 :]:
        text = str(candidate["text"]).strip()
        if not text:
            continue
        if re.match(r"^\d+(?:\.\d+){0,3}\s+", text):
            continue
        if re.match(r"^(参\s*考\s*文\s*献|致\s*谢|附录\s*[A-ZＡ-Ｚ])$", text):
            continue
        return candidate
    return None


def next_matching_paragraph(
    paragraphs: list[dict[str, object]],
    item: dict[str, object] | None,
    *,
    skip_patterns: list[str] | None = None,
) -> dict[str, object] | None:
    if item is None:
        return None
    compiled = [re.compile(pattern) for pattern in (skip_patterns or [])]
    position = paragraphs.index(item)
    for candidate in paragraphs[position + 1 :]:
        text = str(candidate["text"])
        if not text.strip():
            continue
        if any(pattern.search(text) for pattern in compiled):
            continue
        return candidate
    return None


def find_first_after(
    paragraphs: list[dict[str, object]],
    item: dict[str, object] | None,
    pattern: str,
) -> dict[str, object] | None:
    if item is None:
        return None
    regex = re.compile(pattern)
    position = paragraphs.index(item)
    for candidate in paragraphs[position + 1 :]:
        if regex.search(str(candidate["text"])):
            return candidate
    return None


def find_first_after_position(
    paragraphs: list[dict[str, object]],
    position: int,
    pattern: str,
) -> dict[str, object] | None:
    regex = re.compile(pattern)
    for candidate in paragraphs:
        if int(candidate["paragraph"].Range.Start) <= position:
            continue
        if regex.search(str(candidate["text"])):
            return candidate
    return None


def ensure_segments(item: dict[str, object] | None) -> list[dict[str, object]]:
    if item is None:
        return []
    segments = item.get("segments")
    if isinstance(segments, list):
        return segments
    segments = collect_segments(item["paragraph"])
    item["segments"] = segments
    return segments


def dominant_segment(segments: list[dict[str, object]]) -> dict[str, object] | None:
    filtered = [
        segment
        for segment in segments
        if str(segment["text"]).strip()
    ]
    if not filtered:
        return None
    return max(filtered, key=lambda segment: len(re.sub(r"\s+", "", str(segment["text"]))))


def leading_label_signature(paragraph_text: str, segments: list[dict[str, object]], label: str) -> tuple[dict[str, object] | None, dict[str, object] | None]:
    compact_text = paragraph_text.lstrip()
    if not compact_text.startswith(label):
        return None, None
    label_chars = len(label)
    consumed = 0
    label_segments: list[dict[str, object]] = []
    body_segment = None
    for segment in segments:
        text = str(segment["text"])
        if consumed < label_chars:
            label_segments.append(segment)
            consumed += len(text)
            continue
        if text.strip():
            body_segment = segment
            break
    label_segment = dominant_segment(label_segments)
    return label_segment, body_segment


def leading_label_signature_fast(paragraph, label: str, visible_limit: int = 120) -> tuple[dict[str, object] | None, dict[str, object] | None]:
    label_chars = len(label)
    compact = ""
    label_segments: list[dict[str, object]] = []
    current_body: dict[str, object] | None = None
    current_label: dict[str, object] | None = None
    visible_seen = 0

    for index in range(1, paragraph.Range.Characters.Count + 1):
        char_range = paragraph.Range.Characters(index)
        char_text = char_range.Text
        if char_text in {"\r", "\x07"}:
            continue
        visible_seen += 1
        signature = char_signature(char_range)
        compact += char_text
        target = label_segments if len(compact) <= label_chars else None

        if target is not None:
            if current_label is None or any(
                current_label[key] != signature[key]
                for key in ("bold", "ascii_font", "far_east_font", "size")
            ):
                current_label = {
                    "text": char_text,
                    "bold": signature["bold"],
                    "ascii_font": signature["ascii_font"],
                    "far_east_font": signature["far_east_font"],
                    "size": signature["size"],
                }
                label_segments.append(current_label)
            else:
                current_label["text"] = str(current_label["text"]) + char_text
        elif char_text.strip():
            if current_body is None:
                current_body = {
                    "text": char_text,
                    "bold": signature["bold"],
                    "ascii_font": signature["ascii_font"],
                    "far_east_font": signature["far_east_font"],
                    "size": signature["size"],
                }
            elif all(current_body[key] == signature[key] for key in ("bold", "ascii_font", "far_east_font", "size")):
                current_body["text"] = str(current_body["text"]) + char_text
            else:
                break

        if visible_seen >= visible_limit and current_body is not None:
            break

    if not compact.lstrip().startswith(label):
        return None, None
    return dominant_segment(label_segments), current_body


def contains_font(segment: dict[str, object] | None, expected: str, key: str) -> bool:
    if not segment:
        return False
    return expected.lower() in str(segment.get(key, "")).lower()


def compact_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text)


def paragraph_format_signature(paragraph) -> dict[str, object]:
    paragraph_format = paragraph.Range.ParagraphFormat
    return {
        "alignment": getattr(paragraph_format, "Alignment", None),
        "line_spacing_rule": getattr(paragraph_format, "LineSpacingRule", None),
        "line_spacing": getattr(paragraph_format, "LineSpacing", None),
        "character_unit_first_line_indent": getattr(paragraph_format, "CharacterUnitFirstLineIndent", None),
        "character_unit_left_indent": getattr(paragraph_format, "CharacterUnitLeftIndent", None),
    }


def paragraph_check(
    item: dict[str, object] | None,
    *,
    expected_ascii: str | None = None,
    expected_far_east: str | None = None,
    expected_bold: bool | None = None,
    expected_size: float | None = None,
    expected_note: str,
) -> dict[str, object]:
    if item is None:
        return {
            "status": "manual_confirm",
            "expected": expected_note,
            "note": "未定位到对应段落。",
        }
    dominant = dominant_segment(ensure_segments(item))
    if dominant is None:
        return {
            "status": "manual_confirm",
            "paragraph_text": item["text"],
            "expected": expected_note,
            "note": "未能提取稳定字符格式。",
        }

    checks: list[bool] = []
    if expected_ascii is not None:
        checks.append(contains_font(dominant, expected_ascii, "ascii_font"))
    if expected_far_east is not None:
        checks.append(contains_font(dominant, expected_far_east, "far_east_font"))
    if expected_bold is not None:
        checks.append(bool(dominant["bold"]) == expected_bold)
    if expected_size is not None:
        checks.append(size_matches(dominant["size"], expected_size))

    status = "confirmed" if checks and all(checks) else "suggested"
    return {
        "status": status,
        "paragraph_text": item["text"],
        "expected": expected_note,
        "observed": dominant,
    }


def inline_label_body_check(item: dict[str, object] | None, label: str, expected_note: str) -> dict[str, object]:
    if item is None:
        return {
            "status": "manual_confirm",
            "expected": expected_note,
            "note": "未定位到对应段落。",
        }
    label_segment, body_segment = leading_label_signature_fast(item["paragraph"], label)
    if label_segment is None:
        return {
            "status": "manual_confirm",
            "paragraph_text": item["text"],
            "expected": expected_note,
            "note": "段落未按预期标签起始。",
        }

    checks = [
        bool(label_segment["bold"]) is True,
        contains_font(label_segment, "Times New Roman", "ascii_font"),
        body_segment is not None,
        bool(body_segment["bold"]) is False if body_segment is not None else False,
        contains_font(body_segment, "Times New Roman", "ascii_font") if body_segment is not None else False,
    ]
    status = "confirmed" if all(checks) else "suggested"
    return {
        "status": status,
        "paragraph_text": item["text"],
        "expected": expected_note,
        "observed": {
            "label_segment": label_segment,
            "body_segment": body_segment,
        },
    }


def chinese_keyword_label_check(item: dict[str, object] | None) -> dict[str, object]:
    expected_note = "“关键词：”标签用黑体，后续关键词内容回到正文样式，不把整行都做成同一强调格式。"
    if item is None:
        return {
            "status": "manual_confirm",
            "expected": expected_note,
            "note": "未定位到关键词段落。",
        }
    label_segment, body_segment = leading_label_signature_fast(item["paragraph"], "关键词：")
    if label_segment is None:
        label_segment, body_segment = leading_label_signature_fast(item["paragraph"], "关键词:")
    if label_segment is None:
        return {
            "status": "manual_confirm",
            "paragraph_text": item["text"],
            "expected": expected_note,
            "note": "段落未按“关键词：”起始。",
        }
    checks = [
        contains_font(label_segment, "黑体", "far_east_font"),
        body_segment is not None,
        contains_font(body_segment, "宋体", "far_east_font") if body_segment is not None else False,
    ]
    status = "confirmed" if all(checks) else "suggested"
    return {
        "status": status,
        "paragraph_text": item["text"],
        "expected": expected_note,
        "observed": {
            "label_segment": label_segment,
            "body_segment": body_segment,
        },
    }


def body_paragraph_check(item: dict[str, object] | None, expected_note: str) -> dict[str, object]:
    if item is None:
        return {
            "status": "manual_confirm",
            "expected": expected_note,
            "note": "未定位到稳定的正文样本段落。",
        }
    dominant = dominant_segment(ensure_segments(item))
    if dominant is None:
        return {
            "status": "manual_confirm",
            "paragraph_text": item["text"],
            "expected": expected_note,
            "note": "未能提取稳定字符格式。",
        }
    text = str(item["text"])
    has_western = bool(re.search(r"[A-Za-z0-9]", text))
    checks = [
        contains_font(dominant, "宋体", "far_east_font"),
        bool(dominant["bold"]) is False,
        size_matches(dominant["size"], 12),
    ]
    if has_western:
        checks.append(contains_font(dominant, "Times New Roman", "ascii_font"))
    status = "confirmed" if all(checks) else "suggested"
    return {
        "status": status,
        "paragraph_text": item["text"],
        "expected": expected_note,
        "observed": {
            **dominant,
            "paragraph_format": paragraph_format_signature(item["paragraph"]),
        },
    }


def reference_entry_check(item: dict[str, object] | None) -> dict[str, object]:
    expected_note = "参考文献条目使用宋体 + Times New Roman 小四；如涉及悬挂缩进，应与模板一致。"
    if item is None:
        return {
            "status": "manual_confirm",
            "expected": expected_note,
            "note": "未定位到参考文献正文样本。",
        }
    segments = ensure_segments(item)
    visible_segments = [segment for segment in segments if str(segment["text"]).strip()]
    if not visible_segments:
        return {
            "status": "manual_confirm",
            "paragraph_text": item["text"],
            "expected": expected_note,
            "note": "未提取到可见字符段。",
        }
    chinese_segments = [
        segment for segment in visible_segments if re.search(r"[\u4e00-\u9fff]", str(segment["text"]))
    ]
    western_segments = [
        segment for segment in visible_segments if re.search(r"[A-Za-z0-9]", str(segment["text"]))
    ]
    paragraph_sig = paragraph_format_signature(item["paragraph"])
    first_indent = paragraph_sig.get("character_unit_first_line_indent")
    checks = [
        any(contains_font(segment, "宋体", "far_east_font") for segment in chinese_segments) if chinese_segments else True,
        any(contains_font(segment, "Times New Roman", "ascii_font") for segment in western_segments) if western_segments else True,
        all(size_matches(segment.get("size"), 12) for segment in visible_segments),
        first_indent is not None and float(first_indent) <= -1.5,
    ]
    status = "confirmed" if all(checks) else "suggested"
    return {
        "status": status,
        "paragraph_text": item["text"],
        "expected": expected_note,
        "observed": {
            "segments": visible_segments[:12],
            "paragraph_format": paragraph_sig,
        },
    }


def toc_special_entry_checks(document) -> dict[str, dict[str, object]]:
    targets = {
        "references_contents_entry_spacing": "参考文献",
        "acknowledgements_contents_entry_spacing": "致谢",
    }
    if int(document.TablesOfContents.Count) == 0:
        return {
            key: {
                "status": "manual_confirm",
                "expected": f"目录中的“{target}”不保留标题行中的字间空格。",
                "note": "文档中未检测到目录域。",
            }
            for key, target in targets.items()
        }

    toc = document.TablesOfContents(1)
    checks: dict[str, dict[str, object]] = {
        key: {
            "status": "manual_confirm",
            "expected": f"目录中的“{target}”应写成“{target}”，不要保留标题行字间空格。",
            "note": "未在目录中定位到对应条目。",
        }
        for key, target in targets.items()
    }

    for index in range(1, toc.Range.Paragraphs.Count + 1):
        paragraph = toc.Range.Paragraphs(index)
        raw_text = normalize_text(paragraph.Range.Text)
        if not raw_text:
            continue
        entry_text = raw_text.split("\t", 1)[0].strip()
        compact_entry = compact_spaces(entry_text)
        for key, target in targets.items():
            if compact_entry != target:
                continue
            checks[key] = {
                "status": "confirmed" if entry_text == target else "suggested",
                "paragraph_text": raw_text,
                "expected": f"目录中的“{target}”应写成“{target}”，不要保留标题行字间空格。",
                "observed": {
                    "entry_text": entry_text,
                    "compact_entry_text": compact_entry,
                },
            }
    return checks


def inspect_document(input_path: Path) -> dict[str, object]:
    try:
        import pythoncom
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Word format inspection. Install it with `pip install pywin32`."
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
        paragraphs = non_empty_paragraphs(document)
        abstract_item = find_first(paragraphs, r"^Abstract:")
        key_words_item = find_first(paragraphs, r"^Key words:")
        chinese_keywords_item = find_first(paragraphs, r"^关键词[:：]")
        abstract_title_item = find_first(paragraphs, r"^摘\s*要$")
        contents_title_item = find_first(paragraphs, r"^目\s*录$")
        references_title_item = find_first(paragraphs, r"^参\s*考\s*文\s*献$")
        acknowledgements_title_item = find_first(paragraphs, r"^致\s*谢$")
        toc_end = int(document.TablesOfContents(1).Range.End) if int(document.TablesOfContents.Count) else 0
        first_body_heading_item = find_first_after_position(paragraphs, toc_end, r"^\d+\s+")
        if first_body_heading_item is None:
            first_body_heading_item = find_first(paragraphs, r"^\d+\s+")

        english_context = previous_items(paragraphs, abstract_item, 3) if abstract_item else []
        english_title_item = english_context[0] if len(english_context) == 3 else None
        english_author_item = english_context[1] if len(english_context) == 3 else None
        english_affiliation_item = english_context[2] if len(english_context) == 3 else None
        body_sample_item = next_matching_paragraph(
            paragraphs,
            first_body_heading_item,
            skip_patterns=[
                r"^\d+(?:\.\d+){0,3}\s+",
                r"^(图|表|续表|注[:：])",
                r"^（式",
            ],
        )
        reference_body_item = next_nonempty_after(paragraphs, references_title_item)
        acknowledgement_body_item = next_nonempty_after(paragraphs, acknowledgements_title_item)
        toc_checks = toc_special_entry_checks(document)

        report = {
            "file": str(input_path),
            "source_type": input_path.suffix.lower(),
            "judgement_basis": {
                "word_character_format": "confirmed",
                "rendered_layout": "manual_confirm",
                "note": "本脚本用于检查字体、字号、加粗边界、目录特殊条目和末尾模块样式，不替代页面版式审查。",
            },
            "checks": {
                "abstract_title_font": paragraph_check(
                    abstract_title_item,
                    expected_far_east="黑体",
                    expected_note="“摘        要”标题使用黑体；模板批注未要求把此项单独判为显式加粗。",
                ),
                "english_title_format": paragraph_check(
                    english_title_item,
                    expected_ascii="Times New Roman",
                    expected_bold=True,
                    expected_note="英文题目使用 Times New Roman，且显式加粗。",
                ),
                "english_author_format": paragraph_check(
                    english_author_item,
                    expected_ascii="Times New Roman",
                    expected_bold=False,
                    expected_note="英文作者姓名使用 Times New Roman，且不作为显式加粗项。",
                ),
                "english_affiliation_format": paragraph_check(
                    english_affiliation_item,
                    expected_ascii="Times New Roman",
                    expected_bold=False,
                    expected_note="英文作者单位使用 Times New Roman，且不作为显式加粗项。",
                ),
                "abstract_label_body_format": inline_label_body_check(
                    abstract_item,
                    "Abstract:",
                    "“Abstract:”标签加粗，紧随其后的摘要正文不加粗，且都使用 Times New Roman。",
                ),
                "keywords_cn_label_body_format": chinese_keyword_label_check(chinese_keywords_item),
                "keywords_en_label_body_format": inline_label_body_check(
                    key_words_item,
                    "Key words:",
                    "“Key words:”标签加粗，后续关键词内容不加粗，且都使用 Times New Roman。",
                ),
                "contents_title_font": paragraph_check(
                    contents_title_item,
                    expected_far_east="黑体",
                    expected_note="“目        录”标题使用黑体；重点检查标题间距，不把其显式加粗作为默认硬规则。",
                ),
                "body_sample_font": body_paragraph_check(
                    body_sample_item,
                    expected_note="正文示例段落使用中文宋体、西文 Times New Roman、小四号。",
                ),
                "references_title_font": paragraph_check(
                    references_title_item,
                    expected_far_east="黑体",
                    expected_size=14,
                    expected_note="“参  考  文  献”标题使用黑体，重点检查字间距和字号，不擅自追加显式加粗要求。",
                ),
                "references_body_font": reference_entry_check(reference_body_item),
                "acknowledgements_title_font": paragraph_check(
                    acknowledgements_title_item,
                    expected_far_east="黑体",
                    expected_size=14,
                    expected_note="“致        谢”标题使用黑体，重点检查标题字距和字号。",
                ),
                "acknowledgements_body_font": body_paragraph_check(
                    acknowledgement_body_item,
                    "致谢正文使用中文宋体、西文 Times New Roman、小四号，不作为显式加粗项。",
                ),
                **toc_checks,
            },
            "guardrails": [
                "只有模板批注明确写了“加粗”，才把 bold 当硬规则。",
                "黑体、楷体这类字体要求本身不等于必须打开 Word 的 Bold 属性。",
                "正文、参考文献、致谢被修改后，要回看字符级字体和字号，不要只看内容是否正确。",
                "每次刷新目录后，都要复查目录中的“参考文献”和“致谢”是否被标题空格污染。",
                "页面位置、跨页、目录页码等仍需配合渲染页面审查。",
            ],
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
        pythoncom.CoUninitialize()


def main() -> int:
    try:
        require_windows()
        args = parse_args()
        input_path = resolve_input(args.input_path)
        report = inspect_document(input_path)
        payload = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output_path:
            Path(args.output_path).expanduser().resolve().write_text(payload, encoding="utf-8")
        else:
            print(payload)
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
