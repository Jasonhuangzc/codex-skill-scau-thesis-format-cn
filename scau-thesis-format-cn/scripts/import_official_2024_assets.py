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
        help="Deprecated: changed files must use a separately verified manifest, not bypass the 2024 hashes.",
    )
    parser.add_argument("--verify-only", action="store_true", help="Verify all source files without copying or Word conversion; works on any platform.")
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


def verify_files(source_dir: Path, skip_hash_check: bool = False) -> list[dict[str, str]]:
    if skip_hash_check:
        raise ValueError("--skip-hash-check cannot certify the 2024 package. Verify a changed school version separately and update its manifest before importing.")
    manifest = load_manifest()
    verified: list[dict[str, str]] = []
    for item in manifest["required_files"]:  # type: ignore[index]
        filename = item["filename"]  # type: ignore[index]
        expected_hash = item["sha256"]  # type: ignore[index]
        source_path = source_dir / filename
        if not source_path.exists():
            raise FileNotFoundError(f"Missing official file: {source_path}")
        observed_hash = sha256(source_path)
        if observed_hash != expected_hash.upper():
            raise RuntimeError(
                f"SHA256 mismatch for {filename}: expected {expected_hash}, observed {observed_hash}"
            )
        target_path = OFFICIAL_DIR / filename
        verified.append(
            {
                "filename": filename,
                "source": str(source_path),
                "target": str(target_path),
                "sha256": observed_hash,
            }
        )
    return verified


def import_files(source_dir: Path, skip_hash_check: bool) -> list[dict[str, str]]:
    # Validate the entire set before changing any existing local package file.
    verified = verify_files(source_dir, skip_hash_check)
    OFFICIAL_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    for item in verified:
        source_path, target_path = Path(item["source"]), Path(item["target"])
        if source_path.resolve() != target_path.resolve():
            shutil.copy2(source_path, target_path)
    return verified


def regenerate_template_assets() -> dict[str, object]:
    source_doc = OFFICIAL_DIR / "附件6.华南农业大学本科毕业论文（设计）格式模板.doc"
    template_doc = TEMPLATE_DIR / "scau-undergrad-thesis-template.doc"
    template_docx = TEMPLATE_DIR / "scau-undergrad-thesis-template.docx"
    preview_pdf = TEMPLATE_DIR / "scau-undergrad-thesis-template-preview.pdf"

    def ps_path(path: Path) -> str:
        return str(path).replace("'", "''")

    shutil.copy2(source_doc, template_doc)

    powershell = rf"""
$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open('{ps_path(source_doc)}', $false, $true)
  $doc.SaveAs([ref]'{ps_path(template_docx)}', [ref]16)
  $doc.ExportAsFixedFormat('{ps_path(preview_pdf)}', 17)
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
    comment_count = int(comment_payload.get("comment_count", 0))
    if comment_count != 50:
        raise RuntimeError(f"Converted 2024 template comment count mismatch: expected 50, observed {comment_count}. Recheck conversion before using derived assets.")
    return {
        "template_doc": str(template_doc),
        "template_docx": str(template_docx),
        "preview_pdf": str(preview_pdf),
        "comment_count": comment_count,
    }


def main() -> int:
    try:
        args = parse_args()
        source_dir = Path(args.source_dir).expanduser().resolve()
        if not source_dir.exists():
            raise FileNotFoundError(f"Source directory not found: {source_dir}")

        if args.verify_only:
            print(json.dumps({"mode": "verify_only", "verified_files": verify_files(source_dir, args.skip_hash_check),
                              "derived_assets": "not_checked"}, ensure_ascii=False, indent=2))
            return 0
        require_windows()

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
