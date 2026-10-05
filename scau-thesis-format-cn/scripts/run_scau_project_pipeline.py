#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from word_template_utils import contents_paragraph_elements, delete_range, normalize_heading_text, normalize_keyword_heading


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
BUNDLED_TEMPLATE_DOCX = SKILL_ROOT / "assets" / "template" / "scau-undergrad-thesis-template.docx"


class SkillStepError(RuntimeError):
    pass


def emit_text(text: str, *, stderr: bool = False) -> None:
    stream = sys.stderr if stderr else sys.stdout
    payload = text if text.endswith("\n") else text + "\n"
    encoding = getattr(stream, "encoding", None) or "utf-8"
    stream.buffer.write(payload.encode(encoding, errors="backslashreplace"))
    stream.flush()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Project-level runner for the SCAU thesis Word-template workflow."
    )
    parser.add_argument("--project-root", required=True, help="Thesis workspace root.")
    parser.add_argument("--official-template-docx", help="Official template .docx used for frontmatter bootstrap and donor fallback.")
    parser.add_argument("--metadata-file", help="Path to metadata JSON. Supports thesis_metadata.json and generic metadata.json names.")
    parser.add_argument("--docx", help="Base working .docx file.")
    parser.add_argument("--chapter-file", help="Markdown chapter draft to insert.")
    parser.add_argument("--figures-root", help="Root containing figure folders such as 实验结果3-*.")
    parser.add_argument("--figure-number-prefix", help="Only include figure folders matching this prefix, such as 3-.")
    parser.add_argument("--tables-manifest", help="Table manifest JSON path for table-block insertion.")
    parser.add_argument("--references-file", help="Reference source markdown/text file.")
    parser.add_argument("--manifest-output", help="Output path for the generated figure manifest.")
    parser.add_argument("--output", help="Final output .docx path.")
    parser.add_argument(
        "--figure-backend",
        choices=["python-docx", "word-com"],
        default="python-docx",
        help="Backend for figure insertion. Use word-com for large local Word documents on Windows.",
    )
    parser.add_argument(
        "--word-visible",
        action="store_true",
        help="Show the Word window when a Word COM backend is used.",
    )
    parser.add_argument(
        "--chapter-heading",
        help="Normalized chapter heading in Word, such as '3  结果与分析'. Used when trimming template sample body.",
    )
    parser.add_argument("--skip-frontmatter", action="store_true", help="Do not refresh cover, declarations, and abstract frontmatter.")
    parser.add_argument("--skip-generate-figures", action="store_true", help="Do not rerun figure generation scripts.")
    parser.add_argument("--skip-chapter", action="store_true", help="Do not insert chapter Markdown.")
    parser.add_argument("--skip-figures", action="store_true", help="Do not insert figure blocks.")
    parser.add_argument("--skip-tables", action="store_true", help="Do not insert table blocks.")
    parser.add_argument("--skip-references", action="store_true", help="Do not insert references.")
    parser.add_argument(
        "--keep-template-body",
        action="store_true",
        help="Keep the template's sample body chapters instead of trimming them after insertion.",
    )
    parser.add_argument("--trim-template-body", action="store_true", help="Explicitly remove a sample body only after matching every body block against the official template.")
    parser.add_argument("--finalize-contents", action="store_true", help="Opt in to Word COM field/TOC refresh on Windows; does not normalize unrelated document layout.")
    parser.add_argument("--replace-media", action="store_true", help="Explicitly permit chapter media removal after checking a complete reconstruction manifest; never inferred from figure insertion.")
    parser.add_argument("--replace-tables", action="store_true", help="Explicitly permit chapter table removal after checking a complete reconstruction manifest.")
    return parser.parse_args()


def ensure_sibling_script(name: str) -> Path:
    path = SCRIPT_DIR / name
    if not path.exists():
        raise FileNotFoundError(f"未找到 skill 脚本: {path}")
    return path


def run_step(step_name: str, command: list[str], cwd: Path) -> str:
    result = subprocess.run(
        command,
        check=False,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.stdout.strip():
        emit_text(f"{step_name}: completed", stderr=True)
    if result.returncode != 0:
        recovery_hints = {
            "frontmatter": "检查 thesis_metadata.json、官方模板 docx 路径，以及封面段落锚点是否仍与学校模板一致。",
            "figure-manifest": "检查图目录命名、覆盖配置文件和章节 Markdown 是否存在；必要时先关闭 --run-generators。",
            "chapter": "检查章节 Markdown 的一级标题是否规范，以及工作稿中 donor 样式是否还存在。",
            "tables": "检查表格 manifest、Markdown 表格文件和表题格式是否符合 `# 表x-x ...`。",
            "figures": "检查锚点 regex、图片文件路径，以及是否应传入 --official-template-docx 作为 donor fallback。",
            "references": "检查参考文献源文件格式，确认 `## 中文文献` / `## 英文文献` 结构完整。",
            "contents-finalize": "检查本机是否为 Windows + Word + pywin32 环境，以及目录域是否能正常更新。",
        }
        raise SkillStepError(
            json.dumps(
                {
                    "step": step_name,
                    "returncode": result.returncode,
                    "command": command,
                    "stdout": result.stdout.strip(),
                    "stderr": result.stderr.strip(),
                    "recovery_hint": recovery_hints.get(step_name, "查看 stderr 并从当前步骤重新运行。"),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    return result.stdout.strip()


def normalized_heading_from_markdown(path: Path) -> str:
    from insert_markdown_chapter import parse_markdown

    return parse_markdown(path)[0].text


def chapter_tag_from_heading(heading: str) -> str:
    match = re.match(r"^(\d+)\s{2,}(.+)$", heading)
    if match:
        return f"第{match.group(1)}章"
    return "章节"


def discover_work_output_dir(project_root: Path) -> Path:
    for candidate in (
        project_root / "论文终稿",
        project_root / "work",
        project_root / "output",
        project_root / "outputs",
    ):
        if candidate.exists() and candidate.is_dir():
            return candidate
    return project_root / "_scau_thesis_output"


def discover_manifest_dir(project_root: Path) -> Path:
    work_dir = discover_work_output_dir(project_root)
    if work_dir.name == "论文终稿":
        return work_dir / "装版清单"
    return work_dir / "manifests"


def discover_template_docx(project_root: Path) -> Path:
    candidates = [
        project_root / "论文撰写规范" / "附件6_格式模板_转存.docx",
        project_root / "论文撰写规范" / "附件6.华南农业大学本科毕业论文（设计）格式模板.docx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    workspace_named = sorted(project_root.rglob("*格式模板*.docx"))
    if workspace_named:
        return workspace_named[0]

    english_named = sorted(project_root.rglob("*template*.docx"))
    if english_named:
        return english_named[0]

    if BUNDLED_TEMPLATE_DOCX.exists():
        return BUNDLED_TEMPLATE_DOCX
    raise FileNotFoundError("未找到可用的华农论文模板 docx。公开仓库首次使用前请先运行 scripts/import_official_2024_assets.py，或显式传入 --official-template-docx。")


def discover_metadata_file(project_root: Path) -> Path:
    direct = project_root / "thesis_metadata.json"
    if direct.exists():
        return direct
    generic = project_root / "metadata.json"
    if generic.exists():
        return generic
    matches = sorted(project_root.rglob("thesis_metadata.json")) + sorted(project_root.rglob("metadata.json"))
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError("发现多个 metadata JSON，请显式传入 --metadata-file。")
    raise FileNotFoundError("未找到 metadata JSON，请显式传入 --metadata-file。")


def discover_docx_path(project_root: Path) -> Path:
    work_dir = discover_work_output_dir(project_root)
    preferred = [
        work_dir / "毕业论文终稿_工作版.docx",
        work_dir / "scau_thesis_working.docx",
    ]
    for candidate in preferred:
        if candidate.exists():
            return candidate
    return preferred[0]


def discover_figures_root(project_root: Path) -> Path:
    candidates = [
        project_root / "论文草稿",
        project_root / "figures",
        project_root / "images",
        project_root / "assets" / "figures",
        project_root / "图件",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return project_root


def discover_references_file(project_root: Path) -> Path:
    candidates = [
        project_root / "论文草稿" / "最终参考文献著录初稿.md",
        project_root / "references.md",
        project_root / "references.txt",
        project_root / "bibliography.md",
        project_root / "bibliography.txt",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    markdown_matches = sorted(project_root.rglob("*参考文献*.md")) + sorted(project_root.rglob("*reference*.md"))
    if markdown_matches:
        return markdown_matches[0]
    return candidates[0]


def default_manifest_path(project_root: Path, figure_prefix: str | None, heading: str | None) -> Path:
    chapter_number = None
    if figure_prefix:
        chapter_number = figure_prefix.split("-")[0]
    elif heading:
        match = re.match(r"^(\d+)\s", heading)
        if match:
            chapter_number = match.group(1)
    filename = f"第{chapter_number}章图块清单.json" if chapter_number else "图块清单.json"
    return discover_manifest_dir(project_root) / filename


def default_output_path(project_root: Path, heading: str | None) -> Path:
    tag = chapter_tag_from_heading(heading) if heading else "装版"
    work_dir = discover_work_output_dir(project_root)
    if work_dir.name == "论文终稿":
        return work_dir / f"毕业论文终稿_工作版_{tag}装版.docx"
    return work_dir / f"scau_thesis_{tag}_assembled.docx"


def template_body_bounds(document):
    start = None
    end = None
    contents = contents_paragraph_elements(document)
    for paragraph in document.paragraphs:
        if paragraph._p in contents:
            continue
        text = normalize_heading_text(paragraph.text)
        if start is None and re.match(r"^1\s+\S", text):
            start = paragraph
        elif start is not None and normalize_keyword_heading(text) == "参考文献":
            end = paragraph
            break
    if start is None or end is None:
        raise RuntimeError("无法唯一定位模板样例正文（第1章至参考文献之前），拒绝宽区间删除。")
    children = list(document.element.body)
    return start, end, children[children.index(start._p):children.index(end._p)]


def body_signature(document, elements) -> tuple:
    xml = tuple(etree.tostring(element, method="c14n") for element in elements)
    relation_ids = set()
    relationship_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for element in elements:
        for node in element.iter():
            for attribute, value in node.attrib.items():
                if attribute.startswith(relationship_ns):
                    relation_ids.add(value)
    related = []
    for relation_id in sorted(relation_ids):
        relation = document.part.rels[relation_id]
        payload = relation.target_ref if relation.is_external else relation.target_part.blob
        related.append((relation_id, relation.reltype, payload))
    return xml, tuple(related)


def strip_template_body_placeholders(docx_path: Path, template_docx: Path, output_path: Path) -> Path:
    document = Document(docx_path)
    official = Document(template_docx)
    start, end, elements = template_body_bounds(document)
    _, _, official_elements = template_body_bounds(official)
    if body_signature(document, elements) != body_signature(official, official_elements):
        raise RuntimeError("工作稿正文与官方模板样例不完全一致，拒绝裁剪；保留其他章节并逐章回灌。")
    if any(any(element.iter(qn("w:sectPr"))) for element in elements):
        raise RuntimeError("样例正文包含分节符，需要先确认节边界后再裁剪。")
    delete_range(document, start, end)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def publish_output(source: Path, destination: Path) -> None:
    # Publish only a completed pipeline, and keep a prior output intact on failure.
    with tempfile.NamedTemporaryFile(prefix=".scau-thesis-", suffix=".docx", dir=destination.parent, delete=False) as handle:
        staging = Path(handle.name)
    try:
        shutil.copyfile(source, staging)
        os.replace(staging, destination)
    finally:
        staging.unlink(missing_ok=True)


def main() -> None:
    args = parse_args()
    project_root = Path(args.project_root).resolve()
    if not project_root.is_dir():
        raise FileNotFoundError(f"论文工作目录不存在: {project_root}")
    if args.keep_template_body and args.trim_template_body:
        raise ValueError("--keep-template-body 与 --trim-template-body 不能同时使用。")
    if args.finalize_contents and sys.platform != "win32":
        raise ValueError("--finalize-contents 需要 Windows + Microsoft Word + pywin32。")
    if args.figure_backend == "word-com" and not args.skip_figures and sys.platform != "win32":
        raise ValueError("word-com 图件后端仅可在 Windows 使用。")

    chapter_path = Path(args.chapter_file).resolve() if args.chapter_file else None
    chapter_heading = args.chapter_heading
    if not args.skip_chapter:
        if chapter_path is None:
            raise ValueError("未跳过章节插入时，必须提供 --chapter-file。")
        source_heading = normalized_heading_from_markdown(chapter_path)
        if chapter_heading and normalize_heading_text(chapter_heading) != normalize_heading_text(source_heading):
            raise ValueError("--chapter-heading 与 Markdown 的章标题不一致。")
        chapter_heading = source_heading

    metadata_file = None
    if not args.skip_frontmatter:
        metadata_file = Path(args.metadata_file).resolve() if args.metadata_file else discover_metadata_file(project_root)
    docx_path = Path(args.docx).resolve() if args.docx else discover_docx_path(project_root)
    bootstrap = not docx_path.exists() and not args.skip_frontmatter
    if not docx_path.exists() and not bootstrap:
        raise FileNotFoundError(f"工作稿不存在，局部回灌需要已有 --docx: {docx_path}")

    tables_manifest = Path(args.tables_manifest).resolve() if args.tables_manifest and not args.skip_tables else None
    official_template_docx = Path(args.official_template_docx).resolve() if args.official_template_docx else None
    requires_template = bootstrap or args.trim_template_body
    if official_template_docx is None and (requires_template or not args.skip_chapter or not args.skip_figures or tables_manifest is not None):
        try:
            official_template_docx = discover_template_docx(project_root)
        except FileNotFoundError:
            if requires_template:
                raise
    for label, path in (("metadata", metadata_file), ("official template", official_template_docx), ("tables manifest", tables_manifest)):
        if path is not None and not path.is_file():
            raise FileNotFoundError(f"{label} 不存在: {path}")

    figures_root = Path(args.figures_root).resolve() if args.figures_root else discover_figures_root(project_root)
    refs_path = None
    if not args.skip_references:
        refs_path = Path(args.references_file).resolve() if args.references_file else discover_references_file(project_root)
        if not refs_path.is_file():
            raise FileNotFoundError(f"参考文献源文件不存在: {refs_path}")
    manifest_output = Path(args.manifest_output).resolve() if args.manifest_output else default_manifest_path(project_root, args.figure_number_prefix, chapter_heading)
    final_output = Path(args.output).resolve() if args.output else default_output_path(project_root, chapter_heading)
    current_docx = official_template_docx if bootstrap else docx_path
    if final_output in {current_docx, docx_path, official_template_docx}:
        raise ValueError("输出必须与输入工作稿/官方模板分开，以保留可恢复的原稿。")
    editing = args.trim_template_body or not args.skip_frontmatter or not args.skip_chapter or not args.skip_figures or tables_manifest is not None or not args.skip_references
    if editing:
        source_document = Document(current_docx)
        if source_document.element.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"):
            raise RuntimeError("工作稿含未处理的修订；先单独确认/处理修订，再进行结构回灌。")
    if not args.skip_figures:
        manifest_output.parent.mkdir(parents=True, exist_ok=True)
    final_output.parent.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(tempfile.mkdtemp(prefix="thesis-word-pipeline-"))
    steps_run = []
    template_body_trimmed = False
    contents_finalized = False

    def execute(step_name: str, script_name: str, options: list[str], next_docx: Path | None = None) -> None:
        nonlocal current_docx
        run_step(step_name, [sys.executable, str(ensure_sibling_script(script_name)), *options], project_root)
        steps_run.append(step_name)
        if next_docx is not None:
            current_docx = next_docx

    try:
        if args.trim_template_body or (bootstrap and not args.keep_template_body and not args.skip_chapter):
            current_docx = strip_template_body_placeholders(current_docx, official_template_docx, temp_dir / "template_body_cleared.docx")
            template_body_trimmed = True
            steps_run.append("trim-template-body")
        if not args.skip_frontmatter:
            target = temp_dir / "frontmatter_filled.docx"
            execute("frontmatter", "fill_scau_frontmatter.py", ["--workspace", str(project_root), "--meta", str(metadata_file), "--template", str(current_docx), "--output", str(target)], target)
        if not args.skip_figures:
            options = ["--figures-root", str(figures_root), "--output", str(manifest_output)]
            if chapter_path is not None:
                options.extend(["--chapter-file", str(chapter_path)])
            if args.figure_number_prefix:
                options.extend(["--number-prefix", args.figure_number_prefix])
            if not args.skip_generate_figures:
                options.append("--run-generators")
            execute("figure-manifest", "generate_figure_manifest_from_dirs.py", options)
        if not args.skip_chapter:
            target = temp_dir / "chapter_inserted.docx"
            options = ["--docx", str(current_docx), "--chapter-file", str(chapter_path), "--output", str(target)]
            if official_template_docx is not None:
                options.extend(["--fallback-template-docx", str(official_template_docx)])
            if args.replace_media:
                options.append("--replace-media")
            if args.replace_tables:
                options.append("--replace-tables")
            execute("chapter", "insert_markdown_chapter.py", options, target)
        if tables_manifest is not None:
            target = temp_dir / "tables_inserted.docx"
            options = ["--docx", str(current_docx), "--manifest", str(tables_manifest), "--output", str(target)]
            if official_template_docx is not None:
                options.extend(["--fallback-template-docx", str(official_template_docx)])
            execute("tables", "insert_table_blocks.py", options, target)
        if not args.skip_figures:
            target = temp_dir / "figures_inserted.docx"
            options = ["--docx", str(current_docx), "--manifest", str(manifest_output), "--output", str(target)]
            if official_template_docx is not None:
                options.extend(["--fallback-template-docx", str(official_template_docx)])
            if args.figure_backend == "word-com" and args.word_visible:
                options.append("--visible")
            execute("figures", "insert_figure_blocks_com.py" if args.figure_backend == "word-com" else "insert_figure_blocks.py", options, target)
        if not args.skip_references:
            target = temp_dir / "references_inserted.docx"
            execute("references", "insert_reference_batch.py", ["--docx", str(current_docx), "--references-file", str(refs_path), "--output", str(target)], target)
        if args.finalize_contents:
            plan = temp_dir / "finalize_contents_plan.json"
            target = temp_dir / "contents_finalized.docx"
            plan.write_text(json.dumps([{"action": "finalize_contents", "mode": "full", "update_fields": True}], ensure_ascii=False, indent=2), encoding="utf-8")
            execute("contents-finalize", "batch_word_ops.py", [str(current_docx), str(plan), "--output", str(target)], target)
            contents_finalized = True
        publish_output(current_docx, final_output)
    except Exception as exc:
        detail = json.loads(str(exc)) if isinstance(exc, SkillStepError) else {"step": "runner", "error": str(exc)}
        detail.update({"recovery_dir": str(temp_dir), "last_successful_docx": str(current_docx), "output_published": False, "input_docx_preserved": True})
        raise SkillStepError(json.dumps(detail, ensure_ascii=False, indent=2)) from exc
    else:
        shutil.rmtree(temp_dir)

    emit_text(json.dumps({
        "output": str(final_output),
        "manifest": str(manifest_output) if not args.skip_figures else None,
        "base_docx": str(official_template_docx if bootstrap else docx_path),
        "frontmatter_inserted": not args.skip_frontmatter,
        "chapter_inserted": not args.skip_chapter,
        "tables_inserted": tables_manifest is not None,
        "figures_inserted": not args.skip_figures,
        "figure_backend": None if args.skip_figures else args.figure_backend,
        "references_inserted": not args.skip_references,
        "figure_generators_rerun": not args.skip_generate_figures and not args.skip_figures,
        "chapter_heading": chapter_heading,
        "template_body_trimmed": template_body_trimmed,
        "contents_finalized": contents_finalized,
        "contents_refresh_required": editing and not contents_finalized,
        "input_docx_preserved": True,
        "steps_run": steps_run,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        if isinstance(exc, SkillStepError):
            emit_text(str(exc), stderr=True)
        else:
            emit_text(
                json.dumps(
                    {
                        "step": "runner",
                        "error": str(exc),
                        "recovery_hint": "检查当前参数组合是否属于 frontmatter-only / chapter-only / figure-only / table-only / reference-only 之一。",
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                stderr=True,
            )
        raise SystemExit(1)
