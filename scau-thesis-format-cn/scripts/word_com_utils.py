#!/usr/bin/env python3
from __future__ import annotations

import re
import subprocess
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client


BODY_START_REGEX = r"^\s*1\s+绪论\s*$"
BODY_CANDIDATE_MIN_LEN = 12

WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_PAGE_BREAK = 7
POINTS_PER_CM = 28.3464566929


class WordComError(RuntimeError):
    pass


def normalize_paragraph_text(text: str) -> str:
    return text.replace("\r", "").replace("\x07", "").strip()


def cleanup_hidden_winword_processes() -> None:
    command = (
        "$procs = Get-Process WINWORD -ErrorAction SilentlyContinue | "
        "Where-Object { $_.MainWindowHandle -eq 0 }; "
        "if ($procs) { $procs | Stop-Process -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=True,
        text=True,
        check=False,
    )


def wait_for_word(delay_sec: float = 0.08) -> None:
    time.sleep(delay_sec)
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


class WordSession:
    def __init__(self, *, visible: bool = False, retries: int = 2, startup_delay_sec: float = 1.0):
        self.visible = visible
        self.retries = max(1, retries)
        self.startup_delay_sec = startup_delay_sec
        self.app = None
        self._documents = []

    def __enter__(self):
        cleanup_hidden_winword_processes()
        pythoncom.CoInitialize()
        last_error = None
        for attempt in range(1, self.retries + 1):
            try:
                app = win32com.client.DispatchEx("Word.Application")
                app.Visible = self.visible
                app.DisplayAlerts = 0
                app.ScreenUpdating = False
                self.app = app
                return self
            except Exception as exc:  # pragma: no cover - COM-specific retry path
                last_error = exc
                time.sleep(self.startup_delay_sec)
        pythoncom.CoUninitialize()
        raise WordComError(f"Failed to start Word COM after {self.retries} attempts: {last_error}") from last_error

    def __exit__(self, exc_type, exc, tb):
        try:
            for document in reversed(self._documents):
                try:
                    document.Close(False)
                except Exception:
                    pass
        finally:
            if self.app is not None:
                try:
                    self.app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def open_document(self, path: Path, *, read_only: bool = False):
        if self.app is None:
            raise WordComError("Word session is not active.")
        document = self.app.Documents.Open(
            str(path),
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=self.visible,
        )
        self._documents.append(document)
        return document

    def save_document(self, document, output_path: Path | None = None) -> Path:
        if output_path is None:
            document.Save()
            return Path(document.FullName)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        document.SaveAs2(str(output_path))
        return output_path


def find_body_start_index(document, *, body_start_regex: str = BODY_START_REGEX) -> int:
    pattern = re.compile(body_start_regex)
    paragraphs = document.Paragraphs
    for index in range(1, paragraphs.Count + 1):
        text = normalize_paragraph_text(paragraphs.Item(index).Range.Text)
        if pattern.match(text):
            return index
    return 1


def iter_body_paragraphs(document, *, body_start_regex: str = BODY_START_REGEX):
    start = find_body_start_index(document, body_start_regex=body_start_regex)
    paragraphs = document.Paragraphs
    for index in range(start, paragraphs.Count + 1):
        yield paragraphs.Item(index)


def is_body_candidate(text: str) -> bool:
    if len(text) < BODY_CANDIDATE_MIN_LEN:
        return False
    if re.match(r"^(第\s*\d+\s*章|\d+(\.\d+){0,3})\s*", text):
        return False
    if re.match(r"^(摘要|ABSTRACT|Abstract|关键词|Key words|目录|参考文献|附录|致谢)\b", text, re.IGNORECASE):
        return False
    if re.match(r"^(图|表)\s*\d+\s*[-－—]\s*\d+", text):
        return False
    if re.match(r"^(注[:：]|资料来源[:：])", text):
        return False
    return True


def matches_figure_caption(text: str, style_name: str = "") -> bool:
    if "figure_caption" in style_name.lower():
        return True
    return bool(re.match(r"^图\s*\d+(\s*[-－—]\s*\d+)?\b", text))


def matches_table_caption(text: str, style_name: str = "") -> bool:
    if "table" in style_name.lower() and "caption" in style_name.lower():
        return True
    return bool(re.match(r"^(续表|表)\s*\d+(\s*[-－—]\s*\d+)?\b", text))


def matches_note(text: str) -> bool:
    return bool(re.match(r"^(注[:：]|资料来源[:：])", text))


def collect_basic_donors(document) -> dict[str, object]:
    donors: dict[str, object] = {}
    for paragraph in iter_body_paragraphs(document):
        text = normalize_paragraph_text(paragraph.Range.Text)
        style_name = str(paragraph.Range.Style) if paragraph.Range.Style is not None else ""
        if not text:
            continue
        if "figure_caption" not in donors and matches_figure_caption(text, style_name):
            donors["figure_caption"] = paragraph
            continue
        if "table_caption" not in donors and matches_table_caption(text, style_name):
            donors["table_caption"] = paragraph
            continue
        if "note" not in donors and matches_note(text):
            donors["note"] = paragraph
            continue
        if "body" not in donors and is_body_candidate(text):
            donors["body"] = paragraph
    if "figure_caption" not in donors or "table_caption" not in donors or "note" not in donors:
        for index in range(1, document.Paragraphs.Count + 1):
            paragraph = document.Paragraphs.Item(index)
            text = normalize_paragraph_text(paragraph.Range.Text)
            style_name = str(paragraph.Range.Style) if paragraph.Range.Style is not None else ""
            if "figure_caption" not in donors and matches_figure_caption(text, style_name):
                donors["figure_caption"] = paragraph
                continue
            if "table_caption" not in donors and matches_table_caption(text, style_name):
                donors["table_caption"] = paragraph
                continue
            if "note" not in donors and matches_note(text):
                donors["note"] = paragraph
    return donors


def resolve_donors(document, required: list[str], fallback_document=None) -> dict[str, object]:
    donors = collect_basic_donors(document)
    missing = [key for key in required if key not in donors]
    if missing and fallback_document is not None:
        fallback_donors = collect_basic_donors(fallback_document)
        for key in missing:
            if key in fallback_donors:
                donors[key] = fallback_donors[key]
        missing = [key for key in required if key not in donors]
    if missing:
        raise WordComError(f"Missing COM donors: {', '.join(missing)}")
    return donors


def find_paragraph_by_regex(document, pattern: str, *, occurrence: int = 1, body_only: bool = False):
    regex = re.compile(pattern)
    hits = []
    paragraphs = iter_body_paragraphs(document) if body_only else (
        document.Paragraphs.Item(i) for i in range(1, document.Paragraphs.Count + 1)
    )
    for paragraph in paragraphs:
        if regex.search(normalize_paragraph_text(paragraph.Range.Text)):
            hits.append(paragraph)
    if len(hits) < occurrence:
        raise WordComError(f"Anchor regex {pattern!r} matched {len(hits)} paragraphs, need {occurrence}.")
    return hits[occurrence - 1]


def find_marker_paragraph(document, marker: str):
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs.Item(index)
        if normalize_paragraph_text(paragraph.Range.Text) == marker:
            return paragraph
    raise WordComError(f"Could not find temporary marker paragraph: {marker}")


def insert_marker_paragraph(document, anchor_paragraph, *, position: str) -> object:
    marker = f"__CODEX_WORD_MARKER_{uuid.uuid4().hex}__"
    marker_text = marker + "\r"
    range_obj = anchor_paragraph.Range.Duplicate
    if position == "after":
        range_obj.Collapse(WD_COLLAPSE_END)
        start = range_obj.Start
        range_obj.InsertAfter(marker_text)
    elif position == "before":
        range_obj.Collapse(WD_COLLAPSE_START)
        start = range_obj.Start
        range_obj.InsertBefore(marker_text)
    else:
        raise WordComError(f"Unsupported insertion position: {position}")
    marker_range = document.Range(start, start + len(marker_text))
    wait_for_word()
    return marker_range.Paragraphs.Item(1)


def set_paragraph_text(paragraph, text: str) -> None:
    content_range = paragraph.Range.Duplicate
    if content_range.End > content_range.Start:
        content_range.End -= 1
    content_range.Text = text


def apply_donor_format(target_paragraph, donor_paragraph) -> None:
    try:
        target_paragraph.Range.Style = str(donor_paragraph.Range.Style)
    except Exception:
        pass

    target_format = target_paragraph.Range.ParagraphFormat
    donor_format = donor_paragraph.Range.ParagraphFormat
    paragraph_fields = [
        "Alignment",
        "LineSpacingRule",
        "LineSpacing",
        "SpaceBefore",
        "SpaceAfter",
        "LeftIndent",
        "RightIndent",
        "FirstLineIndent",
        "CharacterUnitFirstLineIndent",
        "CharacterUnitLeftIndent",
        "CharacterUnitRightIndent",
        "KeepWithNext",
        "KeepTogether",
        "WidowControl",
        "PageBreakBefore",
        "DisableLineHeightGrid",
    ]
    for field in paragraph_fields:
        try:
            setattr(target_format, field, getattr(donor_format, field))
        except Exception:
            continue

    target_font = target_paragraph.Range.Font
    donor_font = donor_paragraph.Range.Font
    font_fields = [
        "Name",
        "NameAscii",
        "NameFarEast",
        "NameOther",
        "Size",
        "Bold",
        "Italic",
        "Underline",
        "Color",
    ]
    for field in font_fields:
        try:
            setattr(target_font, field, getattr(donor_font, field))
        except Exception:
            continue


def clone_donor_paragraph(document, anchor_paragraph, donor_paragraph, *, position: str, text: str = ""):
    target_paragraph = insert_marker_paragraph(document, anchor_paragraph, position=position)
    apply_donor_format(target_paragraph, donor_paragraph)
    set_paragraph_text(target_paragraph, text)
    wait_for_word()
    return target_paragraph


def set_keep_flags(paragraph, *, keep_with_next: bool | None = None, keep_together: bool | None = None) -> None:
    fmt = paragraph.Range.ParagraphFormat
    if keep_with_next is not None:
        fmt.KeepWithNext = bool(keep_with_next)
    if keep_together is not None:
        fmt.KeepTogether = bool(keep_together)


def set_centered_picture_paragraph(paragraph) -> None:
    paragraph.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    set_keep_flags(paragraph, keep_with_next=True, keep_together=True)


def add_picture_to_paragraph(app, paragraph, image_path: Path, *, width_cm: float) -> None:
    content_range = paragraph.Range.Duplicate
    if content_range.End > content_range.Start:
        content_range.End -= 1
    content_range.Text = ""
    content_range.Collapse(WD_COLLAPSE_START)
    inline_shape = content_range.InlineShapes.AddPicture(
        FileName=str(image_path),
        LinkToFile=False,
        SaveWithDocument=True,
    )
    inline_shape.LockAspectRatio = True
    inline_shape.Width = width_cm * POINTS_PER_CM
    wait_for_word(0.12)


def insert_page_break_paragraph(document, anchor_paragraph, donor_paragraph, *, position: str):
    page_break_paragraph = clone_donor_paragraph(
        document,
        anchor_paragraph,
        donor_paragraph,
        position=position,
        text="",
    )
    break_range = page_break_paragraph.Range.Duplicate
    if break_range.End > break_range.Start:
        break_range.End -= 1
    break_range.InsertBreak(WD_PAGE_BREAK)
    set_keep_flags(page_break_paragraph, keep_with_next=True, keep_together=True)
    wait_for_word()
    return page_break_paragraph


@contextmanager
def maybe_open_fallback_document(session: WordSession, fallback_path: Path | None, current_path: Path):
    if fallback_path is None:
        yield None
        return
    if fallback_path.resolve() == current_path.resolve():
        yield None
        return
    fallback_doc = session.open_document(fallback_path, read_only=True)
    yield fallback_doc
