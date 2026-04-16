#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document

from reference_order_utils import inspect_reference_sequence
from word_template_utils import normalize_keyword_heading


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect thesis bibliography ordering against the SCAU rule.")
    parser.add_argument("--docx", required=True, help="Input thesis .docx path")
    parser.add_argument("--output", help="Optional JSON output path")
    return parser.parse_args()


def find_reference_entries(document: Document) -> list[str]:
    started = False
    entries: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        normalized = normalize_keyword_heading(text)
        if normalized == "参考文献":
            started = True
            continue
        if not started:
            continue
        if not text:
            continue
        if normalized == "致谢" or normalized.startswith("附录"):
            break
        if re.match(r"^\d+(?:\.\d+){0,3}\s+", text):
            break
        entries.append(text)
    return entries


def main() -> int:
    args = parse_args()
    docx_path = Path(args.docx).resolve()
    report = inspect_reference_sequence(find_reference_entries(Document(docx_path)))
    report["docx"] = str(docx_path)

    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        output_path = Path(args.output).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(payload, encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
