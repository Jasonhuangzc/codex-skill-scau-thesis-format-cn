#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from word_template_utils import (
    default_output_path,
    find_heading_donors,
    insert_paragraph_before,
    insert_paragraph_after,
    iter_block_items,
    normalize_keyword_heading,
    remove_block,
    replace_paragraph_text_from_donor,
    set_east_asia_font,
    set_hanging_indent_chars,
    same_block,
)


SECTION_MAP = {
    "中文文献": "cn",
    "英文文献": "en",
}


def parse_reference_source(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    grouped = {"cn": [], "en": []}
    current: str | None = None
    saw_named_sections = any(
        SECTION_MAP.get(re.sub(r"^#+\s*", "", line).strip()) is not None
        for line in lines
        if line.strip().startswith("#")
    )

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("#"):
            heading = re.sub(r"^#+\s*", "", line).strip()
            current = SECTION_MAP.get(heading)
            if current is None and saw_named_sections:
                current = None
            continue
        if line.startswith("|"):
            continue
        if current is not None:
            entry = line.lstrip("-* ").strip()
            if entry:
                grouped[current].append(entry)
        elif not saw_named_sections:
            entry = line.lstrip("-* ").strip()
            if entry:
                grouped["cn"].append(entry)

    combined = grouped["cn"] + grouped["en"]
    if not combined:
        raise RuntimeError(f"No reference entries found in source file: {path}")
    return combined


def find_reference_heading(document: Document) -> Paragraph:
    for paragraph in document.paragraphs:
        if normalize_keyword_heading(paragraph.text) == "参考文献":
            return paragraph
    raise RuntimeError("Could not find the 参考文献 heading in the document.")


def find_reference_range(document: Document, heading: Paragraph):
    blocks = list(iter_block_items(document))
    heading_index = next(idx for idx, block in enumerate(blocks) if same_block(block, heading))
    end_anchor = None
    for block in blocks[heading_index + 1 :]:
        if not isinstance(block, Paragraph):
            continue
        normalized = normalize_keyword_heading(block.text)
        if normalized.startswith("附录") or normalized == "致谢":
            end_anchor = block
            break
    return blocks, heading_index, end_anchor


def find_reference_donor(document: Document, heading: Paragraph, end_anchor) -> Paragraph:
    blocks, heading_index, _ = find_reference_range(document, heading)
    for block in blocks[heading_index + 1 :]:
        if end_anchor is not None and same_block(block, end_anchor):
            break
        if isinstance(block, Paragraph) and block.text.strip():
            return block
    donors = find_heading_donors(document)
    return donors["body"]


def clear_existing_entries(document: Document, heading: Paragraph, end_anchor) -> None:
    clearing = False
    for block in list(iter_block_items(document)):
        if same_block(block, heading):
            clearing = True
            continue
        if clearing and end_anchor is not None and same_block(block, end_anchor):
            break
        if clearing:
            remove_block(block)


def format_reference_paragraph(paragraph: Paragraph, donor: Paragraph, text: str) -> Paragraph:
    replace_paragraph_text_from_donor(paragraph, donor, text)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    set_hanging_indent_chars(paragraph, chars=2.0)
    if paragraph.runs:
        set_east_asia_font(paragraph.runs[-1], east_asia_font="宋体", ascii_font="Times New Roman", size_pt=12)
    return paragraph


def insert_entries(document: Document, heading: Paragraph, end_anchor, donor: Paragraph, entries: list[str]) -> None:
    if end_anchor is not None:
        for entry in entries:
            paragraph = insert_paragraph_before(end_anchor, donor, "")
            format_reference_paragraph(paragraph, donor, entry)
    else:
        cursor = heading
        for entry in entries:
            paragraph = insert_paragraph_after(cursor, donor, "")
            format_reference_paragraph(paragraph, donor, entry)
            cursor = paragraph


def reformat_existing_entries(document: Document, heading: Paragraph, end_anchor, donor: Paragraph) -> int:
    count = 0
    started = False
    for block in list(iter_block_items(document)):
        if same_block(block, heading):
            started = True
            continue
        if started and end_anchor is not None and same_block(block, end_anchor):
            break
        if started and isinstance(block, Paragraph) and block.text.strip():
            format_reference_paragraph(block, donor, block.text)
            count += 1
    return count


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch insert and format a thesis reference section.")
    parser.add_argument("--docx", required=True, help="Input .docx path")
    parser.add_argument("--references-file", help="Markdown or text file containing reference entries")
    parser.add_argument("--output", help="Output .docx path")
    parser.add_argument("--reformat-only", action="store_true", help="Only reformat the existing reference section")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    output_path = Path(args.output).resolve() if args.output else default_output_path(docx_path, "_参考文献")
    document = Document(docx_path)

    heading = find_reference_heading(document)
    _, _, end_anchor = find_reference_range(document, heading)
    donor = find_reference_donor(document, heading, end_anchor)

    if args.reformat_only:
        count = reformat_existing_entries(document, heading, end_anchor, donor)
    else:
        if not args.references_file:
            raise ValueError("--references-file is required unless --reformat-only is used.")
        entries = parse_reference_source(Path(args.references_file).resolve())
        clear_existing_entries(document, heading, end_anchor)
        heading = find_reference_heading(document)
        _, _, end_anchor = find_reference_range(document, heading)
        insert_entries(document, heading, end_anchor, donor, entries)
        count = len(entries)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    print(json.dumps({"output": str(output_path), "entry_count": count, "reformat_only": args.reformat_only}, ensure_ascii=False))


if __name__ == "__main__":
    main()
