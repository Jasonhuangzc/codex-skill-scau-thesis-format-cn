#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def collapse_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def load_comments(docx_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(docx_path) as zf:
        if "word/comments.xml" not in zf.namelist():
            return {}
        root = ET.fromstring(zf.read("word/comments.xml"))
    comments: dict[str, str] = {}
    for comment in root.findall("w:comment", NS):
        cid = comment.attrib.get(f"{{{NS['w']}}}id", "")
        text = "".join(node.text or "" for node in comment.findall(".//w:t", NS))
        comments[cid] = collapse_ws(text)
    return comments


def load_anchor_rows(docx_path: Path, comments: dict[str, str]) -> list[dict[str, object]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return []

    rows: list[dict[str, object]] = []
    paragraph_index = 0
    for paragraph in body.iterfind(".//w:p", NS):
        comment_ids = [
            node.attrib.get(f"{{{NS['w']}}}id", "")
            for node in paragraph.findall(".//w:commentRangeStart", NS)
        ]
        if not comment_ids:
            paragraph_index += 1
            continue
        anchor_text = collapse_ws("".join(node.text or "" for node in paragraph.findall(".//w:t", NS)))
        rows.append(
            {
                "paragraph_index": paragraph_index,
                "comment_ids": comment_ids,
                "anchor_text": anchor_text,
                "rules": [comments.get(cid, "") for cid in comment_ids],
            }
        )
        paragraph_index += 1
    return rows


def markdown_escape(value: object) -> str:
    text = str(value)
    return text.replace("|", "\\|").replace("\n", " ")


def to_markdown(rows: list[dict[str, object]]) -> str:
    lines = [
        "| Anchor paragraph | Comment IDs | Anchor text | Rule |",
        "| --- | --- | --- | --- |",
    ]
    for row in rows:
        lines.append(
            "| "
            + markdown_escape(row["paragraph_index"])
            + " | "
            + markdown_escape(", ".join(row["comment_ids"]))  # type: ignore[arg-type]
            + " | "
            + markdown_escape(row["anchor_text"])
            + " | "
            + markdown_escape(" / ".join(row["rules"]))  # type: ignore[arg-type]
            + " |"
        )
    return "\n".join(lines) + "\n"


def build_payload(docx_path: Path) -> dict[str, object]:
    comments = load_comments(docx_path)
    rows = load_anchor_rows(docx_path, comments)
    seen_ids = {cid for row in rows for cid in row["comment_ids"]}  # type: ignore[index]
    return {
        "docx_path": str(docx_path),
        "comment_count": len(comments),
        "anchor_rows": rows,
        "unanchored_comments": {
            cid: text for cid, text in comments.items() if cid not in seen_ids
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract Word comments and anchor paragraphs from a .docx file.")
    parser.add_argument("docx", help="Path to the source .docx file")
    parser.add_argument("--json-out", help="Optional path for JSON output")
    parser.add_argument("--markdown-out", help="Optional path for Markdown table output")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")
    if docx_path.suffix.lower() != ".docx":
        raise ValueError("This script only supports .docx files.")

    payload = build_payload(docx_path)
    rows = payload["anchor_rows"]  # type: ignore[assignment]
    markdown = to_markdown(rows)  # type: ignore[arg-type]

    if args.json_out:
        json_path = Path(args.json_out).resolve()
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.markdown_out:
        md_path = Path(args.markdown_out).resolve()
        md_path.write_text(markdown, encoding="utf-8")

    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
