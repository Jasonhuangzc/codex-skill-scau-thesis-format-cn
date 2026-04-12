#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document

from word_template_utils import (
    apply_three_line_table_format,
    collapse_ws,
    default_output_path,
    delete_range,
    donor_key_for_paragraph_text,
    find_heading_donors,
    find_next_section_anchor,
    find_paragraph_by_regex,
    find_paragraph_by_text,
    format_continued_table_caption,
    insert_paragraph_after,
    insert_paragraph_before,
    insert_table_after,
    insert_table_before,
    iter_block_items,
    normalize_heading_text,
    same_block,
    set_cell_text,
)


@dataclass
class MdBlock:
    kind: str
    text: str | None = None
    level: int | None = None
    rows: list[list[str]] | None = None
    note: str | None = None


HEADING_RE = re.compile(r"^(#{1,4})\s+(.+?)\s*$")
TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*$")


def parse_md_row(line: str) -> list[str]:
    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|"):
        text = text[:-1]
    return [cell.strip() for cell in text.split("|")]


def is_table_start(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines):
        return False
    current = lines[index]
    nxt = lines[index + 1]
    return "|" in current and TABLE_SEPARATOR_RE.match(nxt) is not None


def normalize_md_heading(level: int, text: str) -> str:
    collapsed = collapse_ws(text)
    if level == 1:
        return normalize_heading_text(collapsed)
    return collapsed


def parse_markdown(path: Path) -> list[MdBlock]:
    lines = path.read_text(encoding="utf-8").splitlines()
    blocks: list[MdBlock] = []
    buffer: list[str] = []
    index = 0

    def flush_paragraph() -> None:
        nonlocal buffer
        text = collapse_ws(" ".join(buffer))
        if text:
            blocks.append(MdBlock(kind="paragraph", text=text))
        buffer = []

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if not stripped:
            flush_paragraph()
            index += 1
            continue

        heading_match = HEADING_RE.match(line)
        if heading_match:
            flush_paragraph()
            level = len(heading_match.group(1))
            title = normalize_md_heading(level, heading_match.group(2))
            blocks.append(MdBlock(kind="heading", text=title, level=level))
            index += 1
            continue

        if is_table_start(lines, index):
            flush_paragraph()
            header = parse_md_row(lines[index])
            index += 2
            rows = [header]
            while index < len(lines) and "|" in lines[index]:
                rows.append(parse_md_row(lines[index]))
                index += 1
            blocks.append(MdBlock(kind="table", rows=rows))
            continue

        buffer.append(stripped)
        index += 1

    flush_paragraph()
    return blocks


def combine_table_blocks(blocks: list[MdBlock]) -> list[MdBlock]:
    combined: list[MdBlock] = []
    index = 0
    while index < len(blocks):
        block = blocks[index]
        if (
            block.kind == "paragraph"
            and donor_key_for_paragraph_text(block.text or "") == "table_caption"
            and index + 1 < len(blocks)
            and blocks[index + 1].kind == "table"
        ):
            note = None
            next_index = index + 2
            if (
                next_index < len(blocks)
                and blocks[next_index].kind == "paragraph"
                and donor_key_for_paragraph_text(blocks[next_index].text or "") == "note"
            ):
                note = blocks[next_index].text
                next_index += 1
            combined.append(MdBlock(kind="table_block", text=block.text, rows=blocks[index + 1].rows, note=note))
            index = next_index
            continue
        combined.append(block)
        index += 1
    return combined


def estimate_row_weight(row_values: list[str]) -> int:
    longest = max((len(collapse_ws(value)) for value in row_values), default=0)
    return max(1, math.ceil(longest / 22))


def split_rows(rows: list[list[str]]) -> list[list[list[str]]]:
    if len(rows) <= 13:
        return [rows]
    header = rows[0]
    body_rows = rows[1:]
    segments: list[list[list[str]]] = []
    current_body: list[list[str]] = []
    current_weight = 0
    for row in body_rows:
        row_weight = estimate_row_weight(row)
        if current_body and (len(current_body) >= 12 or current_weight + row_weight > 16):
            segments.append([header, *current_body])
            current_body = []
            current_weight = 0
        current_body.append(row)
        current_weight += row_weight
    if current_body:
        segments.append([header, *current_body])
    return segments or [rows]


def find_insertion_anchor(document: Document, chapter_title: str | None, insert_before_regex: str) -> tuple[object | None, bool]:
    if chapter_title:
        current = find_paragraph_by_text(document, chapter_title)
        if current is not None:
            blocks = list(iter_block_items(document))
            current_index = next(
                idx for idx, block in enumerate(blocks) if same_block(block, current)
            )
            return find_next_section_anchor(blocks, current_index), True

    anchor = find_paragraph_by_regex(document, insert_before_regex, flags=re.IGNORECASE)
    if anchor is None:
        raise RuntimeError(f"Could not find insertion anchor matching regex: {insert_before_regex}")
    return anchor, False


def render_table(table, rows: list[list[str]]) -> None:
    for row_index, row_values in enumerate(rows):
        if row_index >= len(table.rows):
            table.add_row()
        row = table.rows[row_index]
        while len(row.cells) < len(row_values):
            row._tr.add_tc()
        for col_index, value in enumerate(row_values):
            set_cell_text(row.cells[col_index], value)
    apply_three_line_table_format(table)


def insert_table_block_before(document: Document, anchor, block: MdBlock, donors: dict[str, object]):
    rows = block.rows or []
    if not rows:
        raise RuntimeError("Markdown table block has no rows.")
    current_anchor = anchor
    if block.note:
        current_anchor = insert_paragraph_before(current_anchor, donors["note"], block.note)
    for index, segment in reversed(list(enumerate(split_rows(rows)))):
        caption_text = block.text or ""
        if index > 0:
            caption_text = format_continued_table_caption(caption_text)
        current_anchor = insert_paragraph_before(current_anchor, donors["table_caption"], caption_text)
        table = insert_table_before(document, current_anchor, rows=len(segment), cols=len(segment[0]))
        render_table(table, segment)
        current_anchor = table
    return current_anchor


def insert_table_block_after(document: Document, anchor, block: MdBlock, donors: dict[str, object]):
    rows = block.rows or []
    if not rows:
        raise RuntimeError("Markdown table block has no rows.")
    current_anchor = anchor
    for index, segment in enumerate(split_rows(rows)):
        caption_text = block.text or ""
        if index > 0:
            caption_text = format_continued_table_caption(caption_text)
        current_anchor = insert_paragraph_after(current_anchor, donors["table_caption"], caption_text)
        table = insert_table_after(document, current_anchor, rows=len(segment), cols=len(segment[0]))
        render_table(table, segment)
        current_anchor = table
    if block.note:
        current_anchor = insert_paragraph_after(current_anchor, donors["note"], block.note)
    return current_anchor


def insert_block_before(document: Document, anchor, block: MdBlock, donors: dict[str, object]):
    if block.kind == "heading":
        donor = donors[f"heading{block.level}"]
        return insert_paragraph_before(anchor, donor, block.text or "")
    if block.kind == "paragraph":
        donor = donors[donor_key_for_paragraph_text(block.text or "")]
        return insert_paragraph_before(anchor, donor, block.text or "")
    if block.kind == "table":
        rows = block.rows or []
        if not rows:
            raise RuntimeError("Markdown table block has no rows.")
        table = insert_table_before(document, anchor, rows=len(rows), cols=len(rows[0]))
        render_table(table, rows)
        return table
    if block.kind == "table_block":
        return insert_table_block_before(document, anchor, block, donors)
    raise ValueError(f"Unsupported block kind: {block.kind}")


def insert_block_after(document: Document, anchor, block: MdBlock, donors: dict[str, object]):
    if block.kind == "heading":
        donor = donors[f"heading{block.level}"]
        return insert_paragraph_after(anchor, donor, block.text or "")
    if block.kind == "paragraph":
        donor = donors[donor_key_for_paragraph_text(block.text or "")]
        return insert_paragraph_after(anchor, donor, block.text or "")
    if block.kind == "table":
        rows = block.rows or []
        if not rows:
            raise RuntimeError("Markdown table block has no rows.")
        table = insert_table_after(document, anchor, rows=len(rows), cols=len(rows[0]))
        render_table(table, rows)
        return table
    if block.kind == "table_block":
        return insert_table_block_after(document, anchor, block, donors)
    raise ValueError(f"Unsupported block kind: {block.kind}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Insert a Markdown chapter into the thesis Word template.")
    parser.add_argument("--docx", required=True, help="Path to the input .docx file")
    parser.add_argument("--chapter-file", required=True, help="Markdown file containing the chapter draft")
    parser.add_argument("--output", help="Output .docx path. Defaults to overwrite-safe sibling path.")
    parser.add_argument(
        "--insert-before-regex",
        default=r"^参\s*考\s*文\s*献$",
        help="Fallback regex anchor when the chapter heading does not yet exist in the document",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    chapter_path = Path(args.chapter_file).resolve()
    output_path = Path(args.output).resolve() if args.output else default_output_path(docx_path, "_章节插入")

    blocks = combine_table_blocks(parse_markdown(chapter_path))
    if not blocks:
        raise RuntimeError(f"No usable content found in Markdown file: {chapter_path}")

    chapter_title = None
    if blocks[0].kind == "heading" and blocks[0].level == 1:
        chapter_title = blocks[0].text

    document = Document(docx_path)
    donors = find_heading_donors(document)
    anchor, replacing_existing = find_insertion_anchor(document, chapter_title, args.insert_before_regex)

    if replacing_existing and chapter_title:
        existing_heading = find_paragraph_by_text(document, chapter_title)
        if existing_heading is None:
            raise RuntimeError(f"Expected to find existing chapter heading: {chapter_title}")
        delete_range(document, existing_heading, anchor)

    if anchor is not None:
        current_anchor = anchor
        for block in reversed(blocks):
            current_anchor = insert_block_before(document, current_anchor, block, donors)
    else:
        existing_blocks = list(iter_block_items(document))
        if not existing_blocks:
            raise RuntimeError("Document has no body blocks to append after.")
        cursor = existing_blocks[-1]
        for block in blocks:
            cursor = insert_block_after(document, cursor, block, donors)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    print(
        json.dumps(
            {
                "output": str(output_path),
                "replaced_existing_chapter": replacing_existing,
                "chapter_title": chapter_title,
                "source_markdown": str(chapter_path),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
