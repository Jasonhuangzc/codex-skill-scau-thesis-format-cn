#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
TEMPLATE_DOCX = SKILL_ROOT / "assets" / "template" / "scau-undergrad-thesis-template.docx"

TEXT_EXTENSIONS = {".md", ".py", ".yaml", ".yml", ".json", ".toml", ".txt", ".ps1"}
DEFAULT_BANNED_TOKENS: list[str] = []


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run a generic smoke test for the SCAU thesis-format skill."
    )
    parser.add_argument(
        "--workspace",
        help="Optional workspace root for the smoke test. Defaults to a temporary directory.",
    )
    parser.add_argument(
        "--output",
        help="Optional JSON output path.",
    )
    parser.add_argument(
        "--keep-workspace",
        action="store_true",
        help="Keep the generated smoke-test workspace.",
    )
    parser.add_argument(
        "--banned-token",
        action="append",
        default=[],
        help="Additional residue token to scan for. Can be passed multiple times.",
    )
    return parser.parse_args()


def run_json(command: list[str], cwd: Path) -> dict[str, object]:
    result = subprocess.run(
        command,
        cwd=str(cwd),
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
                    "command": command,
                    "returncode": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    stdout = result.stdout.strip()
    if not stdout:
        return {}
    return json.loads(stdout)


def run_plain(command: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        command,
        cwd=str(cwd),
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
                    "command": command,
                    "returncode": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    return result


def write_generic_workspace(workspace: Path) -> tuple[Path, Path]:
    drafts_dir = workspace / "drafts"
    drafts_dir.mkdir(parents=True, exist_ok=True)

    metadata = {
        "thesis_title_zh": "基于模板驱动的本科论文装版与终稿审查流程研究",
        "thesis_title_en": "Template-Driven Undergraduate Thesis Assembly and Final Format Audit Workflow",
        "college": "测试学院",
        "college_en": "Test College",
        "major": "测试专业",
        "student_name_zh": "测试学生",
        "english_name": "Test Student",
        "student_id": "2024000001",
        "advisor_name_zh": "测试导师",
        "advisor_title": "教授",
        "submission_date": "2026年4月12日",
        "university_en": "South China Agricultural University",
        "city_en": "Guangzhou",
        "postal_code": "510642",
        "abstract_zh": "本文围绕本科毕业论文的模板装版与终稿格式审查流程展开，重点验证模板锚点、章节回灌和局部修订是否能够在不破坏原有版式的前提下稳定执行，为终稿阶段的低返工、高一致性处理提供方法支持。",
        "keywords_zh": ["模板装版", "终稿审查", "Word模板"],
        "abstract_en": "This study evaluates a template-driven workflow for undergraduate thesis assembly and final format audit, with emphasis on front-matter anchors, large chapter backfill, and local revision operations that should preserve the original layout and font signatures.",
        "keywords_en": ["Template Assembly", "Final Audit", "Word Formatting"],
    }
    metadata_path = workspace / "metadata.json"
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    chapter_text = """# 第1章 绪论

## 1.1 研究背景

近年来，本科毕业论文越来越依赖真实 Word 模板进行装版与终稿审查。若在终稿阶段仍需频繁人工修补，则容易造成格式漂移、图表断裂与目录失配等问题。

## 1.2 技术路线

本研究采用模板锚点映射、章节回灌、渲染复核与局部替换相结合的方式，降低终稿阶段的返工成本，并保证字体、字号和段落样式稳定继承。
"""
    chapter_path = drafts_dir / "chapter1.md"
    chapter_path.write_text(chapter_text, encoding="utf-8")
    return metadata_path, chapter_path


def scan_project_specific_residue(tokens: list[str]) -> list[dict[str, str]]:
    hits: list[dict[str, str]] = []
    for path in SKILL_ROOT.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in TEXT_EXTENSIONS:
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        for token in tokens:
            if token in text:
                hits.append({"file": str(path), "token": token})
    return hits


def read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    args = parse_args()
    if sys.platform != "win32":
        raise RuntimeError("This smoke test requires Windows because Word format-signature checks use Word COM.")
    if not TEMPLATE_DOCX.exists():
        raise RuntimeError(
            "Bundled template .docx is missing. In the public repo, run scripts/import_official_2024_assets.py first."
        )

    created_temp_root = args.workspace is None
    root = Path(args.workspace).resolve() if args.workspace else Path(tempfile.mkdtemp(prefix="scau-skill-smoke-"))
    workspace = root / "workspace" if created_temp_root else root
    workspace.mkdir(parents=True, exist_ok=True)

    try:
        metadata_path, chapter_path = write_generic_workspace(workspace)
        work_output_dir = workspace / "_scau_thesis_output"
        working_docx = work_output_dir / "scau_thesis_working.docx"
        chapter_docx = work_output_dir / "chapter_inserted.docx"
        replaced_docx = work_output_dir / "chapter_replaced.docx"
        replace_plan = work_output_dir / "replace_plan.json"
        frontmatter_signatures = work_output_dir / "frontmatter_signatures.json"
        replaced_signatures = work_output_dir / "chapter_replaced_signatures.json"

        comment_payload = run_json(
            [sys.executable, str(SCRIPT_DIR / "extract_docx_comments.py"), str(TEMPLATE_DOCX)],
            SKILL_ROOT,
        )

        run_json(
            [sys.executable, str(SCRIPT_DIR / "fill_scau_frontmatter.py"), "--workspace", str(workspace)],
            SKILL_ROOT,
        )
        run_plain(
            [sys.executable, str(SCRIPT_DIR / "inspect_word_format_signatures.py"), str(working_docx), "--output", str(frontmatter_signatures)],
            SKILL_ROOT,
        )

        run_json(
            [
                sys.executable,
                str(SCRIPT_DIR / "insert_markdown_chapter.py"),
                "--docx",
                str(working_docx),
                "--chapter-file",
                str(chapter_path),
                "--output",
                str(chapter_docx),
            ],
            SKILL_ROOT,
        )

        replace_plan.write_text(
            json.dumps(
                [
                    {
                        "action": "replace_text",
                        "find_text": "终稿阶段的返工成本",
                        "replace_text": "终稿阶段的人工返工成本",
                    }
                ],
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        replace_result = run_json(
            [
                sys.executable,
                str(SCRIPT_DIR / "batch_word_ops.py"),
                str(chapter_docx),
                str(replace_plan),
                "--output",
                str(replaced_docx),
            ],
            SKILL_ROOT,
        )
        run_plain(
            [sys.executable, str(SCRIPT_DIR / "inspect_word_format_signatures.py"), str(replaced_docx), "--output", str(replaced_signatures)],
            SKILL_ROOT,
        )

        frontmatter_checks = read_json(frontmatter_signatures).get("checks", {})
        replaced_checks = read_json(replaced_signatures).get("checks", {})

        required_frontmatter = {
            "english_title_format": "confirmed",
            "abstract_label_body_format": "confirmed",
            "keywords_cn_label_body_format": "confirmed",
            "keywords_en_label_body_format": "confirmed",
        }
        required_replaced = {
            "body_sample_font": "confirmed",
            "references_contents_entry_spacing": "confirmed",
            "acknowledgements_contents_entry_spacing": "confirmed",
        }

        residue_tokens = DEFAULT_BANNED_TOKENS + list(args.banned_token)
        residue_hits = scan_project_specific_residue(residue_tokens)
        comment_count = int(comment_payload.get("comment_count", 0))
        replaced_count = (
            replace_result.get("operation_results", [{}])[0].get("replaced_count", 0)  # type: ignore[index]
            if isinstance(replace_result.get("operation_results"), list)
            else 0
        )

        report = {
            "workspace": str(workspace),
            "official_template_comment_count": {
                "status": "pass" if comment_count == 50 else "fail",
                "observed": comment_count,
                "expected": 50,
            },
            "reusability_scan": {
                "status": "pass" if not residue_hits else "fail",
                "tokens": residue_tokens,
                "hits": residue_hits,
            },
            "frontmatter_fill": {
                "working_docx": str(working_docx),
                "checks": {
                    key: {
                        "expected": expected,
                        "observed": ((frontmatter_checks.get(key) or {}).get("status")),
                    }
                    for key, expected in required_frontmatter.items()
                },
            },
            "chapter_backfill": {
                "chapter_docx": str(chapter_docx),
                "status": "pass" if chapter_docx.exists() else "fail",
            },
            "small_replace": {
                "replaced_docx": str(replaced_docx),
                "replaced_count": replaced_count,
                "checks": {
                    key: {
                        "expected": expected,
                        "observed": ((replaced_checks.get(key) or {}).get("status")),
                    }
                    for key, expected in required_replaced.items()
                },
            },
        }

        overall_pass = (
            report["official_template_comment_count"]["status"] == "pass"
            and report["reusability_scan"]["status"] == "pass"
            and report["chapter_backfill"]["status"] == "pass"
            and replaced_count == 1
            and all(item["observed"] == item["expected"] for item in report["frontmatter_fill"]["checks"].values())
            and all(item["observed"] == item["expected"] for item in report["small_replace"]["checks"].values())
        )
        report["overall_status"] = "pass" if overall_pass else "fail"

        payload = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output:
            Path(args.output).resolve().write_text(payload, encoding="utf-8")
        else:
            print(payload)
        return 0 if overall_pass else 1
    finally:
        if created_temp_root and not args.keep_workspace:
            shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
