#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from insert_figure_blocks import load_manifest, resolve_image_paths, should_insert_page_break
from word_com_utils import (
    WordComError,
    WordSession,
    add_picture_to_paragraph,
    clone_donor_paragraph,
    find_paragraph_by_regex,
    insert_page_break_paragraph,
    maybe_open_fallback_document,
    resolve_donors,
    set_centered_picture_paragraph,
    set_keep_flags,
)


def safe_default_output_path(docx_path: Path) -> Path:
    return docx_path.with_name(f"{docx_path.stem}_插图插入_com{docx_path.suffix}")


def emit_text(text: str, *, stderr: bool = False) -> None:
    stream = sys.stderr if stderr else sys.stdout
    payload = text if text.endswith("\n") else text + "\n"
    encoding = getattr(stream, "encoding", None) or "utf-8"
    stream.buffer.write(payload.encode(encoding, errors="backslashreplace"))
    stream.flush()


def insert_entry_after(app, document, anchor, entry: dict, donors: dict[str, object], base_dir: Path):
    images = resolve_image_paths(entry, base_dir)
    width_cm = float(entry.get("width_cm", 12.0))
    if should_insert_page_break(entry, images, width_cm):
        anchor = insert_page_break_paragraph(document, anchor, donors["body"], position="after")

    cursor = anchor
    for image_path in images:
        picture_paragraph = clone_donor_paragraph(document, cursor, donors["body"], position="after", text="")
        set_centered_picture_paragraph(picture_paragraph)
        add_picture_to_paragraph(app, picture_paragraph, image_path, width_cm=width_cm)
        cursor = picture_paragraph

    caption_paragraph = clone_donor_paragraph(
        document,
        cursor,
        donors["figure_caption"],
        position="after",
        text=entry["caption"],
    )
    set_keep_flags(caption_paragraph, keep_with_next=bool(entry.get("note")), keep_together=True)
    cursor = caption_paragraph

    if entry.get("note"):
        note_paragraph = clone_donor_paragraph(
            document,
            cursor,
            donors["note"],
            position="after",
            text=entry["note"],
        )
        set_keep_flags(note_paragraph, keep_with_next=False, keep_together=True)
        cursor = note_paragraph
    return cursor


def insert_entry_before(app, document, anchor, entry: dict, donors: dict[str, object], base_dir: Path):
    images = resolve_image_paths(entry, base_dir)
    width_cm = float(entry.get("width_cm", 12.0))
    current_anchor = anchor
    if should_insert_page_break(entry, images, width_cm):
        current_anchor = insert_page_break_paragraph(document, current_anchor, donors["body"], position="before")

    if entry.get("note"):
        note_paragraph = clone_donor_paragraph(
            document,
            current_anchor,
            donors["note"],
            position="before",
            text=entry["note"],
        )
        set_keep_flags(note_paragraph, keep_with_next=False, keep_together=True)
        current_anchor = note_paragraph

    caption_paragraph = clone_donor_paragraph(
        document,
        current_anchor,
        donors["figure_caption"],
        position="before",
        text=entry["caption"],
    )
    set_keep_flags(caption_paragraph, keep_with_next=bool(entry.get("note")), keep_together=True)
    current_anchor = caption_paragraph

    for image_path in reversed(images):
        picture_paragraph = clone_donor_paragraph(document, current_anchor, donors["body"], position="before", text="")
        set_centered_picture_paragraph(picture_paragraph)
        add_picture_to_paragraph(app, picture_paragraph, image_path, width_cm=width_cm)
        current_anchor = picture_paragraph

    return current_anchor


def main() -> None:
    parser = argparse.ArgumentParser(description="Insert figure blocks into a thesis .docx through Word COM.")
    parser.add_argument("--docx", required=True, help="Path to the input .docx")
    parser.add_argument("--manifest", required=True, help="Path to the figure manifest JSON")
    parser.add_argument("--fallback-template-docx", help="Official template .docx used when the working copy lacks figure-caption or note donors.")
    parser.add_argument("--output", help="Output .docx path")
    parser.add_argument("--visible", action="store_true", help="Show the Word window while inserting figures")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    manifest_path = Path(args.manifest).resolve()
    output_path = Path(args.output).resolve() if args.output else safe_default_output_path(docx_path)
    fallback_path = Path(args.fallback_template_docx).resolve() if args.fallback_template_docx else None
    entries = load_manifest(manifest_path)
    required_donors = ["body", "figure_caption"]
    if any(entry.get("note") for entry in entries):
        required_donors.append("note")

    with WordSession(visible=args.visible) as session:
        document = session.open_document(docx_path, read_only=False)
        with maybe_open_fallback_document(session, fallback_path, docx_path) as fallback_document:
            base_dir = manifest_path.parent
            for index, entry in enumerate(entries, start=1):
                try:
                    donors = resolve_donors(document, required_donors, fallback_document=fallback_document)
                    anchor = find_paragraph_by_regex(
                        document,
                        entry["anchor_regex"],
                        occurrence=int(entry.get("occurrence", 1)),
                        body_only=True,
                    )
                    position = (entry.get("position") or "after").lower()
                    if position == "after":
                        insert_entry_after(session.app, document, anchor, entry, donors, base_dir)
                    elif position == "before":
                        insert_entry_before(session.app, document, anchor, entry, donors, base_dir)
                    else:
                        raise WordComError(f"Unsupported figure insertion position: {position}")
                except Exception as exc:
                    raise WordComError(
                        f"Figure {index} failed: {entry.get('caption', '<no caption>')} | {exc}"
                    ) from exc

            saved_path = session.save_document(document, output_path)

    emit_text(
        json.dumps(
            {
                "output": str(saved_path),
                "manifest": str(manifest_path),
                "entries": len(entries),
                "backend": "word-com",
            },
            ensure_ascii=True,
        )
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        emit_text(
            json.dumps(
                {
                    "step": "figures-com",
                    "error": str(exc),
                    "recovery_hint": "确认 Windows 本机已安装 Word，且当前文档没有被手工锁定；必要时改回 python-docx backend 重试。",
                },
                ensure_ascii=True,
            ),
            stderr=True,
        )
        raise SystemExit(1)
