#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
OFFICIAL_DIR = SKILL_ROOT / "assets" / "official-2024"
TEMPLATE_DIR = SKILL_ROOT / "assets" / "template"
MANIFEST_PATH = OFFICIAL_DIR / "manifest.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import the official 2024 SCAU thesis files into the public skill workspace."
    )
    parser.add_argument(
        "--source-dir",
        required=True,
        help="Directory containing the three official 2024 SCAU thesis files.",
    )
    parser.add_argument(
        "--skip-hash-check",
        action="store_true",
        help="Skip SHA256 validation. Not recommended unless the school released a newer but identically named file.",
    )
    return parser.parse_args()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def load_manifest() -> dict[str, object]:
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("This importer requires Windows because it regenerates the template via Microsoft Word COM.")


def import_files(source_dir: Path, skip_hash_check: bool) -> list[dict[str, str]]:
    manifest = load_manifest()
    copied: list[dict[str, str]] = []
    OFFICIAL_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)

    for item in manifest["required_files"]:  # type: ignore[index]
        filename = item["filename"]  # type: ignore[index]
        expected_hash = item["sha256"]  # type: ignore[index]
        source_path = source_dir / filename
        if not source_path.exists():
            raise FileNotFoundError(f"Missing official file: {source_path}")
        observed_hash = sha256(source_path)
        if not skip_hash_check and observed_hash != expected_hash:
            raise RuntimeError(
                f"SHA256 mismatch for {filename}: expected {expected_hash}, observed {observed_hash}"
            )
        target_path = OFFICIAL_DIR / filename
        shutil.copy2(source_path, target_path)
        copied.append(
            {
                "filename": filename,
                "source": str(source_path),
                "target": str(target_path),
                "sha256": observed_hash,
            }
        )
    return copied


def regenerate_template_assets() -> dict[str, object]:
    source_doc = OFFICIAL_DIR / "附件6.华南农业大学本科毕业论文（设计）格式模板.doc"
    template_doc = TEMPLATE_DIR / "scau-undergrad-thesis-template.doc"
    template_docx = TEMPLATE_DIR / "scau-undergrad-thesis-template.docx"
    preview_pdf = TEMPLATE_DIR / "scau-undergrad-thesis-template-preview.pdf"

    shutil.copy2(source_doc, template_doc)

    powershell = rf"""
$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open('{source_doc}', $false, $true)
  $doc.SaveAs([ref]'{template_docx}', [ref]16)
  $doc.ExportAsFixedFormat('{preview_pdf}', 17)
  $doc.Close($false)
  $doc = $null
  $word.Quit()
  $word = $null
}} finally {{
  if ($doc -ne $null) {{ try {{ $doc.Close($false) }} catch {{}} }}
  if ($word -ne $null) {{ try {{ $word.Quit() }} catch {{}} }}
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", powershell],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            json.dumps(
                {
                    "step": "regenerate_template_assets",
                    "returncode": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
                ensure_ascii=False,
                indent=2,
            )
        )

    comments_result = subprocess.run(
        [sys.executable, str(SCRIPT_DIR / "extract_docx_comments.py"), str(template_docx)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if comments_result.returncode != 0:
        raise RuntimeError(
            json.dumps(
                {
                    "step": "extract_docx_comments",
                    "returncode": comments_result.returncode,
                    "stdout": comments_result.stdout,
                    "stderr": comments_result.stderr,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    comment_payload = json.loads(comments_result.stdout)
    return {
        "template_doc": str(template_doc),
        "template_docx": str(template_docx),
        "preview_pdf": str(preview_pdf),
        "comment_count": int(comment_payload.get("comment_count", 0)),
    }


def main() -> int:
    try:
        require_windows()
        args = parse_args()
        source_dir = Path(args.source_dir).expanduser().resolve()
        if not source_dir.exists():
            raise FileNotFoundError(f"Source directory not found: {source_dir}")

        copied = import_files(source_dir, args.skip_hash_check)
        derived = regenerate_template_assets()
        print(
            json.dumps(
                {
                    "source_dir": str(source_dir),
                    "copied_files": copied,
                    "derived_assets": derived,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
