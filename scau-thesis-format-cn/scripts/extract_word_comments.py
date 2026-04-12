#!/usr/bin/env python
"""Extract Word comment scope and text with Word COM."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract Word comments as structured JSON."
    )
    parser.add_argument("input_path", help="Path to the input .doc or .docx file")
    return parser.parse_args()


def resolve_input(path_arg: str) -> Path:
    path = Path(path_arg).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if path.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Input file must be .doc or .docx.")
    return path


def main() -> int:
    word = None
    document = None
    try:
        if sys.platform != "win32":
            raise RuntimeError("This script requires Windows and Microsoft Word.")
        try:
            import pythoncom
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 is required for Word comment extraction. Install it with `pip install pywin32`."
            ) from exc

        args = parse_args()
        input_path = resolve_input(args.input_path)

        pythoncom.CoInitialize()
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

        items = []
        for index in range(1, document.Comments.Count + 1):
            comment = document.Comments(index)
            items.append(
                {
                    "index": index,
                    "scope": comment.Scope.Text.replace("\r", " ").replace("\x07", " ").strip(),
                    "comment": comment.Range.Text.replace("\r", " ").strip(),
                }
            )
        print(json.dumps(items, ensure_ascii=False, indent=2))
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
        try:
            import pythoncom  # type: ignore

            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
