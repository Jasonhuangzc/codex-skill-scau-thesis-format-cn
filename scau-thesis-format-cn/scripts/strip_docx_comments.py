#!/usr/bin/env python3
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

COMMENT_PARTS = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
}
COMMENT_REL_SUFFIXES = (
    "/comments",
    "/commentsExtended",
    "/commentsIds",
)


def remove_comment_markup(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for node in list(root.iter(f"{{{W_NS}}}{tag}")):
            parent = find_parent(root, node)
            if parent is not None:
                parent.remove(node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def find_parent(root, child):
    for parent in root.iter():
        for node in list(parent):
            if node is child:
                return parent
    return None


def strip_relationships(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for rel in list(root):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith(COMMENT_REL_SUFFIXES):
            root.remove(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def strip_content_types(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for override in list(root):
        part_name = override.attrib.get("PartName", "").lstrip("/")
        if part_name in COMMENT_PARTS:
            root.remove(override)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_docx(input_path: Path, output_path: Path) -> None:
    with zipfile.ZipFile(input_path, "r") as source_zip, zipfile.ZipFile(output_path, "w") as target_zip:
        for info in source_zip.infolist():
            if info.filename in COMMENT_PARTS:
                continue

            data = source_zip.read(info.filename)
            if info.filename in {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml"}:
                data = remove_comment_markup(data)
            elif info.filename == "word/_rels/document.xml.rels":
                data = strip_relationships(data)
            elif info.filename == "[Content_Types].xml":
                data = strip_content_types(data)

            target_zip.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a clean copy of a .docx file without Word comments.")
    parser.add_argument("docx", help="Input .docx path")
    parser.add_argument("--output", help="Output .docx path")
    args = parser.parse_args()

    input_path = Path(args.docx).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input .docx not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise ValueError("Only .docx files are supported.")

    output_path = Path(args.output).resolve() if args.output else input_path.with_name(f"{input_path.stem}_无批注.docx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    process_docx(input_path, output_path)
    print(output_path)


if __name__ == "__main__":
    main()
