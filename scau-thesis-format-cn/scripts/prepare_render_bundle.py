#!/usr/bin/env python
"""Prepare a stable rendered-page bundle from a Word or PDF source."""

from __future__ import annotations

import argparse
import json
import os
import tempfile
import subprocess
import sys
import time
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Word to PDF with timeout protection, then render PDF pages."
    )
    parser.add_argument("input_path", help="Path to the input .doc/.docx/.pdf file")
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory where the PDF and rendered PNG pages will be stored",
    )
    parser.add_argument(
        "--pdf-output",
        help="Optional PDF output path. Defaults to output-dir/<input-stem>.pdf",
    )
    parser.add_argument(
        "--pages",
        help="Pages to render, e.g. 1,3-5. Defaults to all pages.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="Target DPI for rendering. Defaults to 200.",
    )
    parser.add_argument(
        "--export-timeout",
        type=int,
        default=240,
        help="Seconds to wait for Word-to-PDF export before failing. Defaults to 240.",
    )
    parser.add_argument(
        "--export-retries",
        type=int,
        default=1,
        help="Retry count for Word export on timeout or transient failure. Defaults to 1.",
    )
    parser.add_argument(
        "--reuse-pdf-if-fresh",
        action="store_true",
        help="Reuse an existing PDF when it is newer than the Word source.",
    )
    return parser.parse_args()


def resolve_input(path_arg: str) -> Path:
    path = Path(path_arg).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if path.suffix.lower() not in {".doc", ".docx", ".pdf"}:
        raise ValueError("Input file must be .doc, .docx, or .pdf.")
    return path


def resolve_pdf_output(input_path: Path, output_dir: Path, output_arg: str | None) -> Path:
    if output_arg:
        output_path = Path(output_arg).expanduser().resolve()
    else:
        output_path = output_dir / f"{input_path.stem}.pdf"
    if output_path.suffix.lower() != ".pdf":
        raise ValueError("PDF output path must end with .pdf.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def is_pdf_fresh(word_path: Path, pdf_path: Path) -> bool:
    return pdf_path.exists() and pdf_path.stat().st_mtime >= word_path.stat().st_mtime


def export_word_pdf(
    input_path: Path,
    pdf_path: Path,
    timeout_seconds: int,
    retries: int,
    reuse_if_fresh: bool,
) -> dict[str, object]:
    if reuse_if_fresh and is_pdf_fresh(input_path, pdf_path):
        return {
            "status": "reused",
            "pdf_path": str(pdf_path),
            "attempts": [],
            "timeout_seconds": timeout_seconds,
            "reason": "existing_pdf_is_newer_than_word_source",
        }

    script_path = Path(__file__).with_name("export_word_to_pdf.py")
    attempts: list[dict[str, object]] = []
    max_attempts = max(1, retries + 1)
    for attempt in range(1, max_attempts + 1):
        started = time.time()
        fd, pid_file_name = tempfile.mkstemp(prefix="word-export-", suffix=".pid", dir=str(pdf_path.parent))
        os.close(fd)
        pid_file = Path(pid_file_name)
        pid_file.unlink(missing_ok=True)
        command = [
            sys.executable,
            "-X",
            "utf8",
            str(script_path),
            str(input_path),
            "--output",
            str(pdf_path),
            "--pid-file",
            str(pid_file),
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
            )
        except subprocess.TimeoutExpired:
            kill_result = None
            pid_value = None
            if pid_file.exists():
                try:
                    pid_value = int(pid_file.read_text(encoding="utf-8").strip())
                    kill_completed = subprocess.run(
                        ["taskkill", "/PID", str(pid_value), "/T", "/F"],
                        capture_output=True,
                        text=True,
                    )
                    kill_result = {
                        "pid": pid_value,
                        "returncode": kill_completed.returncode,
                        "stdout": kill_completed.stdout.strip(),
                        "stderr": kill_completed.stderr.strip(),
                    }
                except Exception as exc:
                    kill_result = {"pid": pid_value, "error": str(exc)}
            attempts.append(
                {
                    "attempt": attempt,
                    "status": "timeout",
                    "elapsed_seconds": round(time.time() - started, 2),
                    "cleanup": kill_result,
                }
            )
            try:
                pid_file.unlink(missing_ok=True)
            except Exception:
                pass
            continue

        attempts.append(
            {
                "attempt": attempt,
                "status": "ok" if completed.returncode == 0 else "failed",
                "elapsed_seconds": round(time.time() - started, 2),
                "stdout": completed.stdout.strip(),
                "stderr": completed.stderr.strip(),
                "returncode": completed.returncode,
            }
        )
        try:
            pid_file.unlink(missing_ok=True)
        except Exception:
            pass
        if completed.returncode == 0 and pdf_path.exists():
            return {
                "status": "exported",
                "pdf_path": str(pdf_path),
                "attempts": attempts,
                "timeout_seconds": timeout_seconds,
            }

    raise RuntimeError(json.dumps({"status": "export_failed", "attempts": attempts}, ensure_ascii=False))


def main() -> int:
    try:
        args = parse_args()
        input_path = resolve_input(args.input_path)
        output_dir = Path(args.output_dir).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        source_type = input_path.suffix.lower()
        if source_type == ".pdf":
            pdf_path = input_path
            export_stage = {
                "status": "skipped",
                "reason": "source_is_pdf",
                "pdf_path": str(pdf_path),
                "attempts": [],
            }
        else:
            pdf_path = resolve_pdf_output(input_path, output_dir, args.pdf_output)
            export_stage = export_word_pdf(
                input_path,
                pdf_path,
                timeout_seconds=args.export_timeout,
                retries=args.export_retries,
                reuse_if_fresh=args.reuse_pdf_if_fresh,
            )

        render_dir = output_dir / "rendered_pages"
        render_script = Path(__file__).with_name("render_pdf_pages.py")
        command = [
            sys.executable,
            "-X",
            "utf8",
            str(render_script),
            str(pdf_path),
            "--output-dir",
            str(render_dir),
            "--dpi",
            str(args.dpi),
        ]
        if args.pages:
            command.extend(["--pages", args.pages])
        started = time.time()
        completed = subprocess.run(command, capture_output=True, text=True)
        if completed.returncode != 0:
            raise RuntimeError(
                json.dumps(
                    {
                        "status": "render_failed",
                        "stdout": completed.stdout.strip(),
                        "stderr": completed.stderr.strip(),
                        "returncode": completed.returncode,
                    },
                    ensure_ascii=False,
                )
            )
        render_stage = json.loads(completed.stdout)
        render_stage["elapsed_seconds"] = round(time.time() - started, 2)

        result = {
            "input_file": str(input_path),
            "source_type": source_type,
            "output_dir": str(output_dir),
            "export_stage": export_stage,
            "render_stage": render_stage,
            "audit_basis": {
                "exported_pdf": "confirmed",
                "rendered_page_images": "confirmed",
                "pdf_text_extraction_dependency": "not_required",
            },
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
