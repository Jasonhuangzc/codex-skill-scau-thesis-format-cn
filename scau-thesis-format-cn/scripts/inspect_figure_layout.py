#!/usr/bin/env python
"""Inspect Word thesis figure layout and emit structured figure-audit results."""

from __future__ import annotations

import argparse
import json
import re
import tempfile
from pathlib import Path
from typing import Any

from export_word_to_pdf import export_to_pdf
from render_pdf_pages import render_pdf_pages


WD_ACTIVE_END_PAGE_NUMBER = 3

FIGURE_CAPTION_RE = re.compile(r"^(图\s*([A-Z]?\d+(?:-\d+)?))\s+(.+)$")
TABLE_CAPTION_RE = re.compile(r"^(?:表|续表)\s*[A-Z]?\d+(?:-\d+)?\s+")
NOTE_RE = re.compile(r"^注[:：]")
HEADING_RE = re.compile(r"^(?:\d+(?:\.\d+){0,3}\s+|参\s*考\s*文\s*献|致\s*谢|附录\s*[A-ZＡ-Ｚ])")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect figure layout from a Word thesis and emit JSON."
    )
    parser.add_argument("input_path", help="Path to the input .doc or .docx file")
    parser.add_argument(
        "--pdf",
        dest="pdf_path",
        help="Optional existing PDF path. If omitted, export the Word file to PDF first.",
    )
    parser.add_argument(
        "--output-dir",
        dest="output_dir",
        help="Directory for rendered page images and optional exported PDF.",
    )
    parser.add_argument(
        "--target-regex",
        help="Only audit figure labels matching this regex, e.g. '^图3-[1-8]$'. Defaults to all figures.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="Render DPI. Defaults to 200.",
    )
    parser.add_argument(
        "--output",
        dest="output_path",
        help="Optional JSON output path; prints to stdout when omitted.",
    )
    return parser.parse_args()


def status(level: str, message: str, **extra: Any) -> dict[str, Any]:
    payload: dict[str, Any] = {"level": level, "message": message}
    if extra:
        payload["evidence"] = extra
    return payload


def resolve_input(path_arg: str) -> Path:
    input_path = Path(path_arg).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Input file must have a .doc or .docx extension.")
    return input_path


def ensure_output_dir(output_dir_arg: str | None, input_path: Path) -> Path:
    if output_dir_arg:
        output_dir = Path(output_dir_arg).expanduser().resolve()
    else:
        root = Path(tempfile.gettempdir()) / "thesis-format-check-figure-audit"
        output_dir = root / input_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def normalize_text(text: str) -> str:
    return text.replace("\r", " ").replace("\n", " ").strip()


def collect_word_data(input_path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    try:
        import pythoncom
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for figure layout inspection. Install it with `pip install pywin32`."
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

        paragraphs: list[dict[str, Any]] = []
        for idx, para in enumerate(document.Paragraphs, start=1):
            text = normalize_text(para.Range.Text)
            if not text:
                continue
            paragraphs.append(
                {
                    "index": idx,
                    "text": text,
                    "start": int(para.Range.Start),
                    "end": int(para.Range.End),
                    "page": int(para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)),
                }
            )

        graphics: list[dict[str, Any]] = []
        for idx, shape in enumerate(document.InlineShapes, start=1):
            graphics.append(
                {
                    "kind": "inline_shape",
                    "index": idx,
                    "start": int(shape.Range.Start),
                    "page": int(shape.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)),
                }
            )
        for idx, shape in enumerate(document.Shapes, start=1):
            anchor = shape.Anchor
            graphics.append(
                {
                    "kind": "floating_shape",
                    "index": idx,
                    "start": int(anchor.Start),
                    "page": int(anchor.Information(WD_ACTIVE_END_PAGE_NUMBER)),
                }
            )
        graphics.sort(key=lambda item: item["start"])
        return paragraphs, graphics
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


def find_note(paragraphs: list[dict[str, Any]], start_idx: int) -> dict[str, Any] | None:
    for item in paragraphs[start_idx + 1 :]:
        text = item["text"]
        if FIGURE_CAPTION_RE.match(text) or TABLE_CAPTION_RE.match(text) or HEADING_RE.match(text):
            return None
        if NOTE_RE.match(text):
            return item
        if text:
            return None
    return None


def find_next_heading(paragraphs: list[dict[str, Any]], start_idx: int) -> dict[str, Any] | None:
    for item in paragraphs[start_idx + 1 :]:
        if HEADING_RE.match(item["text"]):
            return item
    return None


def map_graphic(caption: dict[str, Any], graphics: list[dict[str, Any]]) -> dict[str, Any] | None:
    candidates = [
        graphic
        for graphic in graphics
        if graphic["start"] < caption["start"] and graphic["page"] in {caption["page"], max(1, caption["page"] - 1)}
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda item: item["start"])


def build_figure_entry(
    caption: dict[str, Any],
    caption_idx: int,
    paragraphs: list[dict[str, Any]],
    graphics: list[dict[str, Any]],
) -> dict[str, Any]:
    match = FIGURE_CAPTION_RE.match(caption["text"])
    assert match is not None
    figure_label = match.group(1).replace(" ", "")
    note = find_note(paragraphs, caption_idx)
    next_heading = find_next_heading(paragraphs, caption_idx)
    graphic = map_graphic(caption, graphics)

    graphic_page = graphic["page"] if graphic else None
    note_page = note["page"] if note else None
    next_heading_page = next_heading["page"] if next_heading else None

    if graphic_page is None:
        caption_order = status(
            "manual_confirm",
            "未能自动映射到图体对象，需要人工对照页面图像确认图题是否在图下。",
            caption_page=caption["page"],
        )
    elif graphic_page == caption["page"]:
        caption_order = status(
            "confirmed",
            "图体与图题位于同一页，且图题在图体之后。",
            figure_page=graphic_page,
            caption_page=caption["page"],
        )
    else:
        caption_order = status(
            "suggested",
            "图体与图题不在同一页，存在图题跨页或图块断裂风险。",
            figure_page=graphic_page,
            caption_page=caption["page"],
        )

    if note is None:
        note_order = status(
            "manual_confirm",
            "未检测到紧随图题后的图注段落；若该图应有图注或资料来源，需要人工确认。",
            caption_page=caption["page"],
        )
    elif note_page == caption["page"]:
        note_order = status(
            "confirmed",
            "图注位于图题之后且与图题同页。",
            caption_page=caption["page"],
            note_page=note_page,
        )
    else:
        note_order = status(
            "suggested",
            "图注与图题不在同一页，存在图注跨页风险。",
            caption_page=caption["page"],
            note_page=note_page,
        )

    figure_block_pages = sorted(
        {
            page
            for page in [graphic_page, caption["page"], note_page]
            if page is not None
        }
    )
    rendered_pages = sorted(
        {
            page
            for page in [graphic_page, caption["page"], note_page, next_heading_page]
            if page is not None
        }
    )

    split_risk = status(
        "confirmed" if len(figure_block_pages) <= 1 else "suggested",
        "图块集中在单页内。" if len(figure_block_pages) <= 1 else "图体、图题或图注跨页分布，需要重点复查。",
        figure_block_pages=figure_block_pages,
    )

    if next_heading is None:
        next_heading_risk = status(
            "manual_confirm",
            "未检测到后续标题，无法判断是否影响下一节标题起始。",
        )
    elif next_heading_page in figure_block_pages:
        next_heading_risk = status(
            "suggested",
            "后续标题与当前图块落在同页，需人工确认起页是否合理。",
            next_heading=next_heading["text"],
            next_heading_page=next_heading_page,
        )
    else:
        next_heading_risk = status(
            "confirmed",
            "后续标题未与当前图块挤在同一页。",
            next_heading=next_heading["text"],
            next_heading_page=next_heading_page,
        )

    readability = status(
        "manual_confirm",
        "脚本无法可靠自动判断图片压缩后是否清晰可读，需要结合渲染页面人工复核。",
        pages=rendered_pages,
    )

    return {
        "figure_id": figure_label,
        "caption_text": caption["text"],
        "caption_page": caption["page"],
        "note_page": note_page,
        "graphic_page": graphic_page,
        "next_heading_page": next_heading_page,
        "next_heading_text": next_heading["text"] if next_heading else None,
        "figure_block_pages": figure_block_pages,
        "rendered_pages_checked": rendered_pages,
        "caption_order_status": caption_order,
        "note_order_status": note_order,
        "split_risk": split_risk,
        "next_heading_risk": next_heading_risk,
        "readability_status": readability,
        "manual_confirm": [
            "图体细节、字号和压缩可读性仍需结合页面图像人工确认。"
        ],
    }


def build_figure_audit(
    input_path: Path, pdf_path: Path, output_dir: Path, target_regex: str | None, dpi: int
) -> dict[str, Any]:
    paragraphs, graphics = collect_word_data(input_path)
    target = re.compile(target_regex) if target_regex else None

    figure_entries: list[dict[str, Any]] = []
    for idx, paragraph in enumerate(paragraphs):
        match = FIGURE_CAPTION_RE.match(paragraph["text"])
        if not match:
            continue
        figure_id = match.group(1).replace(" ", "")
        if target and not target.search(figure_id):
            continue
        figure_entries.append(build_figure_entry(paragraph, idx, paragraphs, graphics))

    pages_to_render = sorted(
        {
            page
            for entry in figure_entries
            for page in entry["rendered_pages_checked"]
            if page is not None
        }
    )
    render_dir = output_dir / "rendered_pages"
    render_result = render_pdf_pages(pdf_path, render_dir, pages=pages_to_render, dpi=dpi)
    page_image_map = {
        item["page_number"]: item["image_path"] for item in render_result["rendered_pages"]
    }

    for entry in figure_entries:
        entry["page_images"] = {
            page: page_image_map[page]
            for page in entry["rendered_pages_checked"]
            if page in page_image_map
        }

    summary = {
        "figure_count": len(figure_entries),
        "confirmed_count": sum(
            1
            for entry in figure_entries
            if entry["split_risk"]["level"] == "confirmed"
            and entry["caption_order_status"]["level"] == "confirmed"
            and entry["note_order_status"]["level"] == "confirmed"
        ),
        "high_risk_figures": [
            entry["figure_id"]
            for entry in figure_entries
            if entry["split_risk"]["level"] == "suggested"
            or entry["caption_order_status"]["level"] == "suggested"
            or entry["note_order_status"]["level"] == "suggested"
            or entry["next_heading_risk"]["level"] == "suggested"
        ],
        "manual_confirm_figures": [
            entry["figure_id"]
            for entry in figure_entries
            if any(
                entry[key]["level"] == "manual_confirm"
                for key in [
                    "caption_order_status",
                    "note_order_status",
                    "next_heading_risk",
                    "readability_status",
                ]
            )
        ],
    }

    return {
        "source_file": str(input_path),
        "pdf_path": str(pdf_path),
        "render_basis": {
            "word_structure": "confirmed",
            "rendered_pdf": "confirmed",
            "page_images": "confirmed",
            "renderer": render_result["renderer"],
            "fallback_reason": render_result["fallback_reason"],
        },
        "figure_page_map": figure_entries,
        "summary": summary,
        "still_manual_confirm": [
            "图像压缩后是否明显不可读，仍建议在渲染页图上人工复核。",
            "若文档中存在特殊浮动对象或手工排版，图体映射可能需要人工复查。",
        ],
    }


def main() -> int:
    try:
        args = parse_args()
        input_path = resolve_input(args.input_path)
        output_dir = ensure_output_dir(args.output_dir, input_path)
        if args.pdf_path:
            pdf_path = Path(args.pdf_path).expanduser().resolve()
        else:
            pdf_path = output_dir / f"{input_path.stem}.pdf"
            export_to_pdf(input_path, pdf_path)

        result = build_figure_audit(
            input_path=input_path,
            pdf_path=pdf_path,
            output_dir=output_dir,
            target_regex=args.target_regex,
            dpi=args.dpi,
        )
        serialized = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output_path:
            output_path = Path(args.output_path).expanduser().resolve()
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(serialized, encoding="utf-8")
        else:
            print(serialized)
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        import sys

        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
