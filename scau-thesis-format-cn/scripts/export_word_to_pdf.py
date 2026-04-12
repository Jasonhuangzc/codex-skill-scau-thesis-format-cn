#!/usr/bin/env python
"""Export a .doc or .docx file to PDF with Microsoft Word on Windows."""

from __future__ import annotations

import argparse
import os
import tempfile
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export a Word document (.doc/.docx) to PDF using Microsoft Word."
    )
    parser.add_argument("input_path", help="Path to the input .doc or .docx file")
    parser.add_argument(
        "--output",
        dest="output_path",
        help="Optional output PDF path; defaults to input file with .pdf suffix",
    )
    parser.add_argument(
        "--pid-file",
        dest="pid_file",
        help="Optional path to write the spawned WINWORD process id",
    )
    return parser.parse_args()


def require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("This script requires Windows and Microsoft Word.")


def resolve_paths(input_arg: str, output_arg: str | None) -> tuple[Path, Path]:
    input_path = Path(input_arg).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Input file must have a .doc or .docx extension.")

    if output_arg:
        output_path = Path(output_arg).expanduser().resolve()
    else:
        output_path = input_path.with_suffix(".pdf")

    if output_path.suffix.lower() != ".pdf":
        raise ValueError("Output path must end with .pdf.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    return input_path, output_path


def export_to_pdf(input_path: Path, output_path: Path, pid_file: Path | None = None) -> None:
    try:
        import pythoncom
        import win32com.client  # type: ignore
        import win32process  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Word PDF export. Install it with `pip install pywin32`."
        ) from exc

    pythoncom.CoInitialize()
    word = None
    document = None
    wd_export_format_pdf = 17
    wd_do_not_save_changes = 0
    temp_output = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        if pid_file is not None:
            try:
                _, pid = win32process.GetWindowThreadProcessId(word.Hwnd)
                pid_file.write_text(str(pid), encoding="utf-8")
            except Exception:
                pass
        document = word.Documents.Open(
            str(input_path),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
            Revert=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        temp_dir = output_path.parent
        fd, temp_name = tempfile.mkstemp(prefix=output_path.stem + "_", suffix=".pdf", dir=str(temp_dir))
        os.close(fd)
        temp_output = Path(temp_name)
        Path(temp_name).unlink(missing_ok=True)
        document.ExportAsFixedFormat(
            OutputFileName=str(temp_output),
            ExportFormat=wd_export_format_pdf,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
        )
        if not temp_output.exists():
            raise RuntimeError("Word export did not produce a PDF file.")
        temp_output.replace(output_path)
    finally:
        if document is not None:
            try:
                document.Close(wd_do_not_save_changes)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if temp_output is not None and temp_output.exists():
            try:
                temp_output.unlink()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> int:
    try:
        args = parse_args()
        require_windows()
        input_path, output_path = resolve_paths(args.input_path, args.output_path)
        pid_file = Path(args.pid_file).expanduser().resolve() if args.pid_file else None
        export_to_pdf(input_path, output_path, pid_file=pid_file)
        print(output_path)
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
