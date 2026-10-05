#!/usr/bin/env python3
from __future__ import annotations

import argparse
from copy import deepcopy
import json
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from word_template_utils import ensure_output_copy, insert_paragraph_before


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
BUNDLED_TEMPLATE_DOCX = SKILL_ROOT / "assets" / "template" / "scau-undergrad-thesis-template.docx"


def discover_workspace_root(start: Path) -> Path:
    for candidate in [start, *start.parents]:
        if (candidate / "thesis_metadata.json").exists():
            return candidate
        generic_meta = list(candidate.glob("*metadata*.json"))
        if generic_meta:
            return candidate
    return start


def discover_meta(workspace: Path) -> Path:
    meta_path = workspace / "thesis_metadata.json"
    if meta_path.exists():
        return meta_path
    generic_meta = sorted(workspace.rglob("*metadata*.json"))
    if generic_meta:
        return generic_meta[0]
    matches = list(workspace.rglob("thesis_metadata.json"))
    if not matches:
        raise FileNotFoundError("Could not find thesis_metadata.json.")
    return matches[0]


def discover_template(workspace: Path) -> Path:
    preferred = workspace / "论文撰写规范" / "附件6_格式模板_转存.docx"
    if preferred.exists():
        return preferred

    candidates = sorted(workspace.rglob("*格式模板*.docx"))
    if candidates:
        return candidates[0]

    english_named = sorted(workspace.rglob("*template*.docx"))
    if english_named:
        return english_named[0]

    if BUNDLED_TEMPLATE_DOCX.exists():
        return BUNDLED_TEMPLATE_DOCX
    raise FileNotFoundError(
        "Could not find a converted .docx thesis template. In the public repo, run scripts/import_official_2024_assets.py or pass --template explicitly."
    )


def discover_work_output_dir(workspace: Path) -> Path:
    for candidate in (
        workspace / "论文终稿",
        workspace / "work",
        workspace / "output",
        workspace / "outputs",
    ):
        if candidate.exists() and candidate.is_dir():
            return candidate
    return workspace / "_scau_thesis_output"


def default_output_path_for_workspace(workspace: Path) -> Path:
    output_dir = discover_work_output_dir(workspace)
    if output_dir.name == "论文终稿":
        return output_dir / "毕业论文终稿_工作版.docx"
    return output_dir / "scau_thesis_working.docx"


def load_metadata(meta_path: Path) -> dict:
    return json.loads(meta_path.read_text(encoding="utf-8"))


def coerce_text(value: object, default: str) -> str:
    if value is None:
        return default
    text = str(value).strip()
    return text if text else default


def coerce_keywords(value: object, *, sep: str, default: str) -> str:
    if value is None:
        return default
    if isinstance(value, list):
        if any(not isinstance(item, str) or not item.strip() for item in value):
            raise ValueError("Keyword arrays must contain non-empty strings.")
        parts = [item.strip() for item in value]
        return sep.join(parts) if parts else default
    text = str(value).strip()
    return text if text else default


def set_run_font(
    run,
    east_asia_font: str,
    ascii_font: str,
    *,
    size_pt: float | None = None,
    bold: bool | None = None,
) -> None:
    run.font.name = ascii_font
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    run_pr = run._element.get_or_add_rPr()
    run_fonts = run_pr.rFonts
    if run_fonts is None:
        run_fonts = OxmlElement("w:rFonts")
        run_pr.insert(0, run_fonts)
    run_fonts.set(qn("w:ascii"), ascii_font)
    run_fonts.set(qn("w:hAnsi"), ascii_font)
    run_fonts.set(qn("w:cs"), ascii_font)
    run_fonts.set(qn("w:eastAsia"), east_asia_font)


def fill_run(
    paragraph,
    run_index: int,
    text: str,
    east_asia_font: str = "宋体",
    ascii_font: str = "Times New Roman",
    *,
    size_pt: float | None = None,
    bold: bool | None = None,
) -> None:
    if run_index >= len(paragraph.runs):
        raise IndexError(f"Run index {run_index} out of range for paragraph: {paragraph.text!r}")
    run = paragraph.runs[run_index]
    set_run_text(run, text)
    set_run_font(run, east_asia_font, ascii_font, size_pt=size_pt, bold=bold)


def set_run_text(run, text: str) -> None:
    references = [deepcopy(node) for node in run._r.findall(qn("w:commentReference"))]
    run.text = text
    for reference in references:
        run._r.append(reference)


def clear_other_runs(paragraph, keep_indices: set[int]) -> None:
    for idx, run in enumerate(paragraph.runs):
        if idx not in keep_indices:
            set_run_text(run, "")


def reset_paragraph_runs(paragraph) -> None:
    for run in paragraph.runs:
        set_run_text(run, "")


def set_label_body_runs(
    paragraph,
    *,
    label: str,
    body: str,
    label_east_asia: str,
    label_ascii: str,
    body_east_asia: str,
    body_ascii: str,
    label_size_pt: float | None = None,
    body_size_pt: float | None = None,
    label_bold: bool | None = None,
    body_bold: bool | None = None,
) -> None:
    reset_paragraph_runs(paragraph)
    if paragraph.runs:
        label_run = paragraph.runs[0]
    else:
        label_run = paragraph.add_run()
    set_run_text(label_run, label)
    set_run_font(label_run, label_east_asia, label_ascii, size_pt=label_size_pt, bold=label_bold)

    body_run = paragraph.runs[1] if len(paragraph.runs) > 1 else paragraph.add_run()
    set_run_text(body_run, body)
    set_run_font(body_run, body_east_asia, body_ascii, size_pt=body_size_pt, bold=body_bold)


def split_submission_date(text: str) -> tuple[str, str, str]:
    parts = re.findall(r"\d+", text)
    if len(parts) < 3:
        raise ValueError(f"Invalid submission date format: {text}")
    if len(parts) != 3:
        raise ValueError(f"Invalid submission date format: {text}")
    date(*(int(part) for part in parts))
    return parts[0], parts[1], parts[2]


def resolve_frontmatter(doc: Document) -> dict:
    """Resolve stable labels before writing, including on an already filled copy."""
    paragraphs = doc.paragraphs
    compact = [re.sub(r"\s+", "", p.text) for p in paragraphs]

    def unique(name: str, pattern: str, start: int = 0, end: int | None = None) -> int:
        matches = [i for i in range(start, len(paragraphs) if end is None else end)
                   if re.match(pattern, compact[i])]
        if len(matches) != 1:
            raise RuntimeError(f"Ambiguous or missing frontmatter anchor {name}: {matches}")
        return matches[0]

    declaration = unique("originality_statement", r"^(?:华南农业大学)?(?:本科毕业(?:论文|设计)(?:[（(]设计[）)])?)?原创性声明$")
    fields = {}
    for name, pattern in {
        "paper_type": r"^本科毕业(?:论文|设计)(?:[（(]或设计[）)])?$",
        "college": r"^学院[:：]", "major": r"^专业[:：]",
        "student_name_zh": r"^姓名[:：]", "student_id": r"^学号[:：]",
        "advisor": r"^指导教师[:：]", "submission_date": r"^提交日期[:：]",
    }.items():
        fields[name] = unique(name, pattern, end=declaration)
    title_candidates = [i for i in range(fields["paper_type"] + 1, fields["college"])
                        if compact[i]]
    if len(title_candidates) != 1:
        raise RuntimeError("Cannot uniquely locate the cover title between paper type and college.")
    fields["thesis_title_zh"] = title_candidates[0]
    fields["abstract_heading"] = unique("chinese_abstract", r"^摘要$", start=declaration)
    fields["keywords_zh"] = unique("keywords_zh", r"^关键词[:：]", start=fields["abstract_heading"])
    fields["abstract_en"] = unique("abstract_en", r"^Abstract[:：]", start=fields["keywords_zh"])
    fields["keywords_en"] = unique("keywords_en", r"^Keywords[:：]", start=fields["abstract_en"])
    english_lines = [i for i in range(fields["keywords_zh"] + 1, fields["abstract_en"])
                     if compact[i]]
    if len(english_lines) != 3:
        raise RuntimeError("Expected English title, author and affiliation before Abstract:.")
    fields.update(zip(("thesis_title_en", "english_name", "affiliation"), english_lines))
    zh_body = [paragraphs[i] for i in range(fields["abstract_heading"] + 1, fields["keywords_zh"])]
    if not zh_body:
        raise RuntimeError("Chinese abstract has no donor paragraph before 关键词：.")
    fields["abstract_zh_paragraphs"] = zh_body
    fields["abstract_en_paragraphs"] = paragraphs[fields["abstract_en"]:fields["keywords_en"]]
    return {key: (paragraphs[value] if isinstance(value, int) else value)
            for key, value in fields.items()}


def validate_template(doc: Document) -> None:
    fields = resolve_frontmatter(doc)
    touched = [value for value in fields.values() if hasattr(value, "_p")]
    touched.extend(fields["abstract_zh_paragraphs"])
    touched.extend(fields["abstract_en_paragraphs"])
    for paragraph in touched:
        if paragraph._p.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:hyperlink | .//w:fldChar | .//w:fldSimple | .//w:sdt | .//w:drawing | .//w:pict | .//w:object | .//m:oMath | .//m:oMathPara"):
            raise RuntimeError("Frontmatter contains revisions, fields or complex objects; use a targeted Word edit to preserve them.")
    for name, required_run_index in {
        "college": 6, "major": 5, "student_name_zh": 5, "student_id": 5,
        "advisor": 9, "submission_date": 9,
    }.items():
        if required_run_index >= len(fields[name].runs):
            raise RuntimeError(
                f"Template structure mismatch at {name}: "
                f"expected run index {required_run_index}, got only {len(fields[name].runs)} runs."
            )
    for name, start, end in (
        ("college", 7, None), ("major", 6, None), ("student_name_zh", 6, None),
        ("student_id", 6, None), ("advisor", 5, 9), ("advisor", 10, None),
    ):
        if any(run.text.strip(" \t\r\n_＿") for run in fields[name].runs[start:end]):
            raise RuntimeError(f"Cover run mapping changed or field was split at {name}; use a targeted Word edit instead of leaving stale text.")
    for index, run in enumerate(fields["submission_date"].runs):
        if index not in {2, 5, 9} and re.search(r"\d", run.text):
            raise RuntimeError("Submission date run mapping changed; refusing to mix old and new date digits.")


def replace_cover(doc: Document, meta: dict, paper_type: str) -> None:
    fields = resolve_frontmatter(doc)
    fill_run(fields["paper_type"], 0, paper_type, size_pt=36, bold=True)
    clear_other_runs(fields["paper_type"], {0})

    fill_run(fields["thesis_title_zh"], 0, meta["thesis_title_zh"], east_asia_font="黑体", size_pt=22, bold=True)
    clear_other_runs(fields["thesis_title_zh"], {0})

    fill_run(fields["college"], 6, meta["college"], size_pt=15, bold=False)
    fill_run(fields["major"], 5, meta["major"], size_pt=15, bold=False)
    fill_run(fields["student_name_zh"], 5, meta["student_name_zh"], size_pt=15, bold=False)
    fill_run(fields["student_id"], 5, str(meta["student_id"]), east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=15, bold=False)
    fill_run(fields["advisor"], 4, meta["advisor_name_zh"], size_pt=15, bold=False)
    fill_run(fields["advisor"], 9, meta["advisor_title"], size_pt=15, bold=False)

    year, month, day = split_submission_date(meta["submission_date"])
    fill_run(fields["submission_date"], 2, year, east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=15, bold=False)
    fill_run(fields["submission_date"], 5, month, east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=15, bold=False)
    fill_run(fields["submission_date"], 9, day, east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=15, bold=False)


def write_abstract_paragraphs(paragraphs, boundary, text: str, *, english: bool) -> None:
    parts = [part.strip() for part in text.splitlines() if part.strip()]
    if not parts:
        raise ValueError("Abstract content cannot be empty when supplied.")
    donor = paragraphs[0]
    while len(paragraphs) < len(parts):
        paragraph = insert_paragraph_before(boundary, donor, "")
        paragraph.paragraph_format.page_break_before = False
        paragraphs.append(paragraph)
    for index, paragraph in enumerate(paragraphs):
        if index >= len(parts):
            reset_paragraph_runs(paragraph)
            continue
        if english and index == 0:
            set_label_body_runs(
                paragraph, label="Abstract:", body=f" {parts[index]}",
                label_east_asia="Times New Roman", label_ascii="Times New Roman",
                body_east_asia="Times New Roman", body_ascii="Times New Roman",
                label_size_pt=12, body_size_pt=12, label_bold=True, body_bold=False,
            )
        else:
            reset_paragraph_runs(paragraph)
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            set_run_text(run, parts[index])
            set_run_font(run, "Times New Roman" if english else "宋体", "Times New Roman", size_pt=12, bold=False)
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if not english or index > 0:
            ind = paragraph._p.get_or_add_pPr().find(qn("w:ind"))
            if ind is None:
                ind = OxmlElement("w:ind")
                paragraph._p.get_or_add_pPr().append(ind)
            for attr in ("hanging", "hangingChars"):
                ind.attrib.pop(qn(f"w:{attr}"), None)
            ind.set(qn("w:firstLineChars"), "200")
            ind.set(qn("w:firstLine"), "480")


def replace_abstract_frontmatter(doc: Document, meta: dict) -> None:
    fields = resolve_frontmatter(doc)

    zh_keywords = coerce_keywords(meta.get("keywords_zh"), sep="；", default="")
    en_keywords = coerce_keywords(meta.get("keywords_en"), sep="; ", default="")

    if "abstract_zh" in meta:
        write_abstract_paragraphs(fields["abstract_zh_paragraphs"], fields["keywords_zh"],
                                  coerce_text(meta["abstract_zh"], ""), english=False)
    if "abstract_en" in meta:
        write_abstract_paragraphs(fields["abstract_en_paragraphs"], fields["keywords_en"],
                                  coerce_text(meta["abstract_en"], ""), english=True)
    if "keywords_zh" in meta:
        if not zh_keywords:
            raise ValueError("Chinese keywords cannot be empty when supplied.")
        set_label_body_runs(
            fields["keywords_zh"],
            label="关键词：",
            body=zh_keywords,
            label_east_asia="黑体",
            label_ascii="Times New Roman",
            body_east_asia="宋体",
            body_ascii="Times New Roman",
            label_size_pt=12,
            body_size_pt=12,
            label_bold=False,
            body_bold=False,
        )

    fill_run(fields["thesis_title_en"], 0, meta["thesis_title_en"], east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=14, bold=True)
    clear_other_runs(fields["thesis_title_en"], {0})
    fill_run(fields["english_name"], 0, meta["english_name"], east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=12, bold=False)
    clear_other_runs(fields["english_name"], {0})
    affiliation = (
        f"（{meta['college_en']}, {meta['university_en']}, "
        f"{meta['city_en']} {meta['postal_code']}, China）"
    )
    fill_run(fields["affiliation"], 0, affiliation, east_asia_font="Times New Roman", ascii_font="Times New Roman", size_pt=12, bold=False)
    clear_other_runs(fields["affiliation"], {0})
    if "keywords_en" in meta:
        if not en_keywords:
            raise ValueError("English keywords cannot be empty when supplied.")
        set_label_body_runs(
            fields["keywords_en"],
            label="Key words:",
            body=f" {en_keywords}" if en_keywords and not en_keywords.startswith(" ") else en_keywords,
            label_east_asia="Times New Roman",
            label_ascii="Times New Roman",
            body_east_asia="Times New Roman",
            body_ascii="Times New Roman",
            label_size_pt=12,
            body_size_pt=12,
            label_bold=True,
            body_bold=False,
        )


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def paragraph_has_page_break(paragraph) -> bool:
    xml = paragraph._p.xml
    return 'w:type="page"' in xml or "<w:lastRenderedPageBreak" in xml



def ensure_frontmatter_page_breaks(doc: Document) -> dict:
    title = resolve_frontmatter(doc)["thesis_title_en"]
    title.paragraph_format.page_break_before = True
    return {
        "english_abstract_page_break": "applied",
        "paragraph_index": list(doc.element.body).index(title._p),
    }


def normalize_cover_gap(doc: Document) -> dict:
    paragraphs = doc.paragraphs
    declaration_index = None
    for index, paragraph in enumerate(paragraphs):
        if "原创性声明" in paragraph.text:
            declaration_index = index
            break
    if declaration_index is None:
        return {"declaration_found": False, "removed_blank_paragraphs": 0, "kept_page_break_paragraphs": 0}

    date_paragraph = resolve_frontmatter(doc)["submission_date"]
    cover_end_index = next(i for i, p in enumerate(paragraphs) if p._p is date_paragraph._p)
    mid_paragraphs = list(paragraphs[cover_end_index + 1 : declaration_index])
    removed = 0
    kept_breaks = 0
    for paragraph in mid_paragraphs:
        text = paragraph.text.strip()
        has_page_break = paragraph_has_page_break(paragraph) or bool(paragraph._p.xpath("./w:pPr/w:sectPr"))
        has_structure = bool(paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object | .//w:bookmarkStart | .//w:bookmarkEnd | .//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference | .//w:fldChar | .//w:sdt"))
        if has_page_break or has_structure:
            kept_breaks += 1
            continue
        if not text:
            remove_paragraph(paragraph)
            removed += 1
    return {
        "declaration_found": True,
        "declaration_index": declaration_index,
        "removed_blank_paragraphs": removed,
        "kept_page_break_paragraphs": kept_breaks,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill the South China Agricultural University thesis front matter from a metadata JSON file.")
    parser.add_argument("--workspace", help="Workspace root. Defaults to current directory or its parents.")
    parser.add_argument("--meta", help="Path to metadata JSON. Supports thesis_metadata.json and generic metadata.json names.")
    parser.add_argument("--template", help="Path to the converted template .docx")
    parser.add_argument("--output", help="Path to output .docx. Defaults to a detected working directory, then falls back to _scau_thesis_output/scau_thesis_working.docx")
    parser.add_argument("--paper-type", default="本科毕业论文", help="Either 本科毕业论文 or 本科毕业设计")
    args = parser.parse_args()

    start = Path(args.workspace).resolve() if args.workspace else Path.cwd().resolve()
    workspace = discover_workspace_root(start)
    meta_path = Path(args.meta).resolve() if args.meta else discover_meta(workspace)
    template_path = Path(args.template).resolve() if args.template else discover_template(workspace)
    output_path = (
        Path(args.output).resolve()
        if args.output
        else default_output_path_for_workspace(workspace)
    )
    ensure_output_copy(output_path, template_path, meta_path)

    meta = load_metadata(meta_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(template_path)
    validate_template(doc)
    replace_cover(doc, meta, args.paper_type)
    replace_abstract_frontmatter(doc, meta)
    cover_gap_report = normalize_cover_gap(doc)
    frontmatter_page_breaks = ensure_frontmatter_page_breaks(doc)
    doc.save(output_path)
    print(
        json.dumps(
            {
                "output": str(output_path),
                "frontmatter_checks": {
                    "cover_to_declaration": cover_gap_report,
                    "frontmatter_page_breaks": frontmatter_page_breaks,
                },
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
