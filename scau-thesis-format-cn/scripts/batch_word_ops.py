#!/usr/bin/env python
"""Apply a batch of Word COM operations in a single session and save once."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


WD_FIND_STOP = 0
WD_REPLACE_ALL = 2
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_PAGE_BREAK = 7
WD_GO_TO_BOOKMARK = -1
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_JUSTIFY = 3
WD_LINE_SPACE_1PT5 = 1
SMALL_FOUR_PT = 12
LEVEL1_HEADING_PT = 14


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Apply a JSON batch of Word COM operations in one Word session."
    )
    parser.add_argument("input_path", help="Path to the input .doc or .docx file")
    parser.add_argument("plan_path", help="Path to the JSON plan file")
    parser.add_argument(
        "--output",
        dest="output_path",
        help="Optional output path. Defaults to overwriting the input file.",
    )
    return parser.parse_args()


def resolve_path(path_arg: str, suffixes: set[str] | None = None) -> Path:
    path = Path(path_arg).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if suffixes and path.suffix.lower() not in suffixes:
        raise ValueError(f"Unexpected file type for {path}")
    return path


def require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("This script requires Windows and Microsoft Word.")


def load_plan(plan_path: Path) -> list[dict[str, object]]:
    data = json.loads(plan_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Plan file must be a JSON list of operations.")
    return data


def open_document(path: Path):
    try:
        import pythoncom
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Word COM operations. Install it with `pip install pywin32`."
        ) from exc

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = word.Documents.Open(
        str(path),
        ReadOnly=False,
        ConfirmConversions=False,
        AddToRecentFiles=False,
        Revert=False,
        OpenAndRepair=True,
        NoEncodingDialog=True,
    )
    return pythoncom, word, document


def find_range(document, find_text: str):
    rng = document.Content
    finder = rng.Find
    finder.ClearFormatting()
    finder.Text = find_text
    finder.Forward = True
    finder.Wrap = WD_FIND_STOP
    if finder.Execute():
        return rng
    raise ValueError(f"Anchor text not found: {find_text}")


def paragraph_text(paragraph) -> str:
    return str(paragraph.Range.Text).replace("\r", "").replace("\x07", "").strip()


def compact_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text)


def find_paragraph_index(document, pattern: str, start_index: int = 1) -> int | None:
    regex = re.compile(pattern)
    for index in range(start_index, document.Paragraphs.Count + 1):
        if regex.match(paragraph_text(document.Paragraphs(index))):
            return index
    return None


def copy_basic_font_format(target_range, template_range) -> None:
    target_font = target_range.Font
    template_font = template_range.Font
    if template_font.Name:
        target_font.Name = template_font.Name
    if template_font.NameAscii:
        target_font.NameAscii = template_font.NameAscii
    if template_font.NameFarEast:
        target_font.NameFarEast = template_font.NameFarEast
    if template_font.NameOther:
        target_font.NameOther = template_font.NameOther
    if template_font.NameBi:
        target_font.NameBi = template_font.NameBi
    if template_font.Size:
        target_font.Size = template_font.Size
    target_font.Bold = template_font.Bold
    target_font.Italic = template_font.Italic
    target_font.Underline = template_font.Underline


def best_template_range(document, insertion_point: int):
    if insertion_point > 0:
        return document.Range(insertion_point - 1, insertion_point)
    content = document.Content
    if content.End > content.Start:
        return document.Range(content.Start, content.Start + 1)
    return None


def apply_font_preset(target_range, *, far_east: str, ascii_font: str, size: float, bold: bool) -> None:
    font = target_range.Font
    font.NameFarEast = far_east
    font.NameAscii = ascii_font
    font.NameOther = ascii_font
    font.NameBi = ascii_font
    font.Name = ascii_font
    font.Size = size
    font.Bold = -1 if bold else 0


def apply_paragraph_preset(
    target_range,
    *,
    first_line_indent_chars: float | None,
    left_indent_chars: float | None,
    line_spacing: float | None,
    line_spacing_rule: int | None,
    alignment: int | None,
) -> None:
    paragraph_format = target_range.ParagraphFormat
    if first_line_indent_chars is not None:
        paragraph_format.CharacterUnitFirstLineIndent = first_line_indent_chars
    if left_indent_chars is not None:
        paragraph_format.CharacterUnitLeftIndent = left_indent_chars
    if line_spacing_rule is not None:
        paragraph_format.LineSpacingRule = line_spacing_rule
    if line_spacing is not None:
        paragraph_format.LineSpacing = line_spacing
    if alignment is not None:
        paragraph_format.Alignment = alignment


def cleanup_contents_entries(document, result: list[dict[str, object]]) -> None:
    cleaned: list[dict[str, object]] = []
    replacements = {
        "参考文献": "参考文献",
        "致谢": "致谢",
    }
    for toc_index in range(1, document.TablesOfContents.Count + 1):
        toc = document.TablesOfContents(toc_index)
        for para_index in range(1, toc.Range.Paragraphs.Count + 1):
            paragraph = toc.Range.Paragraphs(para_index)
            raw_text = str(paragraph.Range.Text).replace("\x07", "").rstrip("\r")
            if "\t" not in raw_text:
                continue
            entry_text, page_text = raw_text.split("\t", 1)
            compact_entry = compact_spaces(entry_text)
            replacement = replacements.get(compact_entry)
            if replacement and entry_text != replacement:
                paragraph.Range.Text = f"{replacement}\t{page_text}\r"
                cleaned.append(
                    {
                        "toc_index": toc_index,
                        "paragraph_index": para_index,
                        "original_entry": entry_text,
                        "updated_entry": replacement,
                    }
                )
    result.append(
        {
            "action": "cleanup_contents_entries",
            "result": "applied",
            "updated_entries": cleaned,
        }
    )


def refresh_contents(
    document,
    result: list[dict[str, object]],
    *,
    cleanup_special_entries: bool = False,
    mode: str = "full",
) -> None:
    for toc in document.TablesOfContents:
        if mode == "page_numbers_only":
            toc.UpdatePageNumbers()
        else:
            toc.Update()
    result.append({"action": "refresh_contents", "result": "applied", "mode": mode})
    if cleanup_special_entries:
        cleanup_contents_entries(document, result)


def normalize_section_font_range(
    document,
    *,
    start_pattern: str,
    end_patterns: list[str],
    title_far_east: str,
    title_ascii: str,
    title_size: float,
    title_bold: bool,
    body_far_east: str,
    body_ascii: str,
    body_size: float,
    body_bold: bool,
    title_first_line_indent_chars: float | None = None,
    body_first_line_indent_chars: float | None = None,
    body_left_indent_chars: float | None = None,
    body_line_spacing: float | None = None,
    body_line_spacing_rule: int | None = None,
    body_alignment: int | None = None,
) -> dict[str, object]:
    start_index = find_paragraph_index(document, start_pattern)
    if start_index is None:
        return {"result": "not_found"}

    end_index = document.Paragraphs.Count + 1
    for pattern in end_patterns:
        match_index = find_paragraph_index(document, pattern, start_index + 1)
        if match_index is not None:
            end_index = min(end_index, match_index)

    title_para = document.Paragraphs(start_index)
    apply_font_preset(
        title_para.Range,
        far_east=title_far_east,
        ascii_font=title_ascii,
        size=title_size,
        bold=title_bold,
    )
    apply_paragraph_preset(
        title_para.Range,
        first_line_indent_chars=title_first_line_indent_chars,
        left_indent_chars=0.0,
        line_spacing=18.0,
        line_spacing_rule=WD_LINE_SPACE_1PT5,
        alignment=WD_ALIGN_PARAGRAPH_LEFT,
    )

    body_count = 0
    for index in range(start_index + 1, end_index):
        paragraph = document.Paragraphs(index)
        text = paragraph_text(paragraph)
        if not text:
            continue
        apply_font_preset(
            paragraph.Range,
            far_east=body_far_east,
            ascii_font=body_ascii,
            size=body_size,
            bold=body_bold,
        )
        apply_paragraph_preset(
            paragraph.Range,
            first_line_indent_chars=body_first_line_indent_chars,
            left_indent_chars=body_left_indent_chars,
            line_spacing=body_line_spacing,
            line_spacing_rule=body_line_spacing_rule,
            alignment=body_alignment,
        )
        body_count += 1

    return {
        "result": "applied",
        "start_paragraph": start_index,
        "end_paragraph_exclusive": end_index,
        "body_paragraphs_updated": body_count,
    }


def normalize_tail_section_fonts(document, operation: dict[str, object], result: list[dict[str, object]]) -> None:
    requested_sections = operation.get("sections", ["references", "acknowledgements"])
    if not isinstance(requested_sections, list):
        raise ValueError("`sections` must be a list when using normalize_tail_section_fonts.")

    section_results: dict[str, object] = {}
    if "references" in requested_sections:
        section_results["references"] = normalize_section_font_range(
            document,
            start_pattern=r"^参\s*考\s*文\s*献$",
            end_patterns=[r"^附录\s*[A-ZＡ-Ｚ].*$", r"^致\s*谢$"],
            title_far_east="黑体",
            title_ascii="Times New Roman",
            title_size=LEVEL1_HEADING_PT,
            title_bold=False,
            body_far_east="宋体",
            body_ascii="Times New Roman",
            body_size=SMALL_FOUR_PT,
            body_bold=False,
            title_first_line_indent_chars=0.0,
            body_first_line_indent_chars=-2.0,
            body_left_indent_chars=0.0,
            body_line_spacing=18.0,
            body_line_spacing_rule=WD_LINE_SPACE_1PT5,
            body_alignment=WD_ALIGN_PARAGRAPH_JUSTIFY,
        )
    if "acknowledgements" in requested_sections:
        section_results["acknowledgements"] = normalize_section_font_range(
            document,
            start_pattern=r"^致\s*谢$",
            end_patterns=[],
            title_far_east="黑体",
            title_ascii="Times New Roman",
            title_size=LEVEL1_HEADING_PT,
            title_bold=False,
            body_far_east="宋体",
            body_ascii="Times New Roman",
            body_size=SMALL_FOUR_PT,
            body_bold=False,
            title_first_line_indent_chars=0.0,
            body_first_line_indent_chars=2.0,
            body_left_indent_chars=0.0,
            body_line_spacing=18.0,
            body_line_spacing_rule=WD_LINE_SPACE_1PT5,
            body_alignment=WD_ALIGN_PARAGRAPH_JUSTIFY,
        )

    result.append(
        {
            "action": "normalize_tail_section_fonts",
            "result": "applied",
            "sections": section_results,
        }
    )


def get_anchor_range(document, operation: dict[str, object]):
    bookmark = operation.get("bookmark")
    if isinstance(bookmark, str) and bookmark:
        if document.Bookmarks.Exists(bookmark):
            return document.Bookmarks(bookmark).Range
        raise ValueError(f"Bookmark not found: {bookmark}")

    anchor_text = operation.get("anchor_text")
    if isinstance(anchor_text, str) and anchor_text:
        return find_range(document, anchor_text)

    raise ValueError("Operation requires either `bookmark` or `anchor_text`.")


def replace_text(document, operation: dict[str, object], result: list[dict[str, object]]) -> None:
    find_text = str(operation["find_text"])
    replace_text_value = str(operation.get("replace_text", ""))
    replaced_count = 0
    search_start = document.Content.Start

    while search_start <= document.Content.End:
        rng = document.Range(search_start, document.Content.End)
        finder = rng.Find
        finder.ClearFormatting()
        finder.Text = find_text
        finder.Forward = True
        finder.Wrap = WD_FIND_STOP
        if not finder.Execute():
            break

        original_start = rng.Start
        template_range = best_template_range(document, original_start)
        rng.Text = replace_text_value
        if replace_text_value:
            inserted = document.Range(original_start, original_start + len(replace_text_value))
            if template_range is not None:
                copy_basic_font_format(inserted, template_range)
            search_start = inserted.End
        else:
            search_start = original_start
        replaced_count += 1

    result.append(
        {
            "action": "replace_text",
            "find_text": find_text,
            "replace_text": replace_text_value,
            "result": "applied" if replaced_count else "not_found",
            "replaced_count": replaced_count,
        }
    )


def insert_text_after(document, operation: dict[str, object], result: list[dict[str, object]]) -> None:
    rng = get_anchor_range(document, operation)
    rng.Collapse(WD_COLLAPSE_END)
    insert_text = str(operation.get("text", ""))
    insert_start = rng.End
    template_range = best_template_range(document, insert_start)
    rng.InsertAfter(insert_text)
    if insert_text and template_range is not None:
        inserted = document.Range(insert_start, insert_start + len(insert_text))
        copy_basic_font_format(inserted, template_range)
    result.append({"action": "insert_text_after", "result": "applied"})


def insert_page_break_before(document, operation: dict[str, object], result: list[dict[str, object]]) -> None:
    rng = get_anchor_range(document, operation)
    rng.Collapse(WD_COLLAPSE_START)
    rng.InsertBreak(WD_PAGE_BREAK)
    result.append({"action": "insert_page_break_before", "result": "applied"})


def insert_image_after(document, operation: dict[str, object], result: list[dict[str, object]]) -> None:
    image_path = resolve_path(str(operation["image_path"]))
    rng = get_anchor_range(document, operation)
    rng.Collapse(WD_COLLAPSE_END)
    document.InlineShapes.AddPicture(str(image_path), False, True, rng)
    result.append(
        {
            "action": "insert_image_after",
            "image_path": str(image_path),
            "result": "applied",
        }
    )


def apply_operations(document, operations: list[dict[str, object]]) -> list[dict[str, object]]:
    results: list[dict[str, object]] = []
    for operation in operations:
        action = operation.get("action")
        if action == "replace_text":
            replace_text(document, operation, results)
        elif action == "insert_text_after":
            insert_text_after(document, operation, results)
        elif action == "insert_page_break_before":
            insert_page_break_before(document, operation, results)
        elif action == "insert_image_after":
            insert_image_after(document, operation, results)
        elif action == "refresh_contents":
            refresh_contents(
                document,
                results,
                cleanup_special_entries=bool(operation.get("cleanup_special_entries", False)),
                mode=str(operation.get("mode", "full")),
            )
        elif action == "cleanup_contents_entries":
            cleanup_contents_entries(document, results)
        elif action == "normalize_tail_section_fonts":
            normalize_tail_section_fonts(document, operation, results)
        else:
            raise ValueError(f"Unsupported action: {action}")
    return results


def save_document(document, input_path: Path, output_path: Path | None) -> Path:
    if output_path is None:
        target = input_path
    else:
        target = output_path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    document.SaveAs2(str(target), AddToRecentFiles=False)
    return target


def main() -> int:
    pythoncom_module = None
    word = None
    document = None
    try:
        require_windows()
        args = parse_args()
        input_path = resolve_path(args.input_path, {".doc", ".docx"})
        plan_path = resolve_path(args.plan_path, {".json"})
        operations = load_plan(plan_path)
        output_path = Path(args.output_path).expanduser().resolve() if args.output_path else None

        pythoncom_module, word, document = open_document(input_path)
        results = apply_operations(document, operations)
        saved_path = save_document(document, input_path, output_path)
        print(
            json.dumps(
                {
                    "input_file": str(input_path),
                    "saved_file": str(saved_path),
                    "operation_results": results,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
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
        if pythoncom_module is not None:
            try:
                pythoncom_module.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
