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


def write_generic_workspace(workspace: Path) -> tuple[Path, Path, Path]:
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
    references_text = """# 最终参考文献著录初稿

## 中文文献

杨三. 模板化终稿审查流程研究[D]. 广州: 测试大学, 2025.
李四. 学位论文装版自动化方法[J]. 测试学报, 2023, 10(2): 12-18.

## 英文文献

Zhang X, Miller T. Audit workflow for thesis formatting[J]. Journal of Test Systems, 2024, 8(1): 10-18.
Brown A, Smith J. Template-guided Word repair for final theses[J]. Document Engineering Review, 2022, 5(3): 44-53.
"""
    references_path = drafts_dir / "references_final.md"
    references_path.write_text(references_text, encoding="utf-8")
    return metadata_path, chapter_path, references_path


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
        metadata_path, chapter_path, references_path = write_generic_workspace(workspace)
        work_output_dir = workspace / "_scau_thesis_output"
        working_docx = work_output_dir / "scau_thesis_working.docx"
        chapter_docx = work_output_dir / "chapter_inserted.docx"
        replaced_docx = work_output_dir / "chapter_replaced.docx"
        finalized_docx = work_output_dir / "chapter_finalized.docx"
        refs_docx = work_output_dir / "chapter_refs.docx"
        replace_plan = work_output_dir / "replace_plan.json"
        finalize_plan = work_output_dir / "finalize_plan.json"
        frontmatter_signatures = work_output_dir / "frontmatter_signatures.json"
        replaced_signatures = work_output_dir / "chapter_replaced_signatures.json"
        finalized_signatures = work_output_dir / "chapter_finalized_signatures.json"
        reference_order_report = work_output_dir / "reference_order.json"

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
        finalize_plan.write_text(
            json.dumps(
                [
                    {
                        "action": "ensure_page_break_before",
                        "section": "english_abstract",
                    },
                    {
                        "action": "normalize_body_paragraph_layout",
                        "first_line_indent_chars": 2.0,
                        "left_indent_chars": 0.0,
                        "line_spacing": 18.0,
                        "line_spacing_rule": 1,
                        "alignment": 3,
                    },
                    {
                        "action": "normalize_table_cells",
                        "target": "all",
                        "apply_fonts": False,
                        "first_line_indent_chars": 0.0,
                        "left_indent_chars": 0.0,
                        "line_spacing": 18.0,
                        "line_spacing_rule": 1,
                        "alignment": 1,
                    },
                    {
                        "action": "normalize_table_cells",
                        "target": "abbreviation",
                        "apply_fonts": True,
                        "far_east_font": "宋体",
                        "ascii_font": "Times New Roman",
                        "size": 12,
                        "first_line_indent_chars": 0.0,
                        "left_indent_chars": 0.0,
                        "line_spacing": 18.0,
                        "line_spacing_rule": 1,
                        "alignment": 1,
                    },
                    {
                        "action": "normalize_tail_section_fonts",
                        "sections": ["references", "acknowledgements"],
                    },
                    {
                        "action": "finalize_contents",
                        "mode": "full",
                        "update_fields": True,
                    }
                ],
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        run_json(
            [
                sys.executable,
                str(SCRIPT_DIR / "batch_word_ops.py"),
                str(replaced_docx),
                str(finalize_plan),
                "--output",
                str(finalized_docx),
            ],
            SKILL_ROOT,
        )
        run_plain(
            [sys.executable, str(SCRIPT_DIR / "inspect_word_format_signatures.py"), str(replaced_docx), "--output", str(replaced_signatures)],
            SKILL_ROOT,
        )
        run_plain(
            [sys.executable, str(SCRIPT_DIR / "inspect_word_format_signatures.py"), str(finalized_docx), "--output", str(finalized_signatures)],
            SKILL_ROOT,
        )
        run_json(
            [
                sys.executable,
                str(SCRIPT_DIR / "insert_reference_batch.py"),
                "--docx",
                str(chapter_docx),
                "--references-file",
                str(references_path),
                "--output",
                str(refs_docx),
            ],
            SKILL_ROOT,
        )
        run_plain(
            [sys.executable, str(SCRIPT_DIR / "inspect_reference_order.py"), "--docx", str(refs_docx), "--output", str(reference_order_report)],
            SKILL_ROOT,
        )

        frontmatter_checks = read_json(frontmatter_signatures).get("checks", {})
        replaced_checks = read_json(replaced_signatures).get("checks", {})
        finalized_checks = read_json(finalized_signatures).get("checks", {})
        reference_order_checks = read_json(reference_order_report)

        required_frontmatter = {
            "english_title_format": "confirmed",
            "abstract_label_body_format": "confirmed",
            "keywords_cn_label_body_format": "confirmed",
            "keywords_en_label_body_format": "confirmed",
            "abstract_section_page_break": "confirmed",
            "abbreviation_table_format": "confirmed",
            "table_cells_centered": "confirmed",
        }
        required_replaced = {
            "body_sample_font": "confirmed",
            "body_first_line_indent_2chars": "confirmed",
            "body_line_spacing_1p5": "confirmed",
            "references_contents_entry_spacing": "confirmed",
            "acknowledgements_contents_entry_spacing": "confirmed",
        }
        required_finalized = {
            "contents_font_pairing": "confirmed",
            "body_first_line_indent_2chars": "confirmed",
            "body_line_spacing_1p5": "confirmed",
            "table_cells_centered": "confirmed",
            "abbreviation_table_format": "confirmed",
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
            "contents_finalize": {
                "finalized_docx": str(finalized_docx),
                "checks": {
                    key: {
                        "expected": expected,
                        "observed": ((finalized_checks.get(key) or {}).get("status")),
                    }
                    for key, expected in required_finalized.items()
                },
            },
            "reference_sorting": {
                "refs_docx": str(refs_docx),
                "status": reference_order_checks.get("status"),
                "entry_count": reference_order_checks.get("entry_count"),
                "issue_count": len(reference_order_checks.get("issues", [])),
            },
        }

        overall_pass = (
            report["official_template_comment_count"]["status"] == "pass"
            and report["reusability_scan"]["status"] == "pass"
            and report["chapter_backfill"]["status"] == "pass"
            and replaced_count == 1
            and all(item["observed"] == item["expected"] for item in report["frontmatter_fill"]["checks"].values())
            and all(item["observed"] == item["expected"] for item in report["small_replace"]["checks"].values())
            and all(item["observed"] == item["expected"] for item in report["contents_finalize"]["checks"].values())
            and report["reference_sorting"]["status"] == "confirmed"
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
