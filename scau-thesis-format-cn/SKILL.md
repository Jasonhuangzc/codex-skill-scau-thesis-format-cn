---
name: scau-thesis-format-cn
description: Fill the official South China Agricultural University undergraduate thesis Word template with user-provided thesis content, then audit the rendered Word/PDF formatting and iterate targeted repairs until the thesis is submission-ready without changing academic meaning.
---

# SCAU Thesis Format CN

Use one closed loop for South China Agricultural University undergraduate theses:

1. start from the real Word template
2. fill user content into the template
3. audit the rendered Word/PDF result
4. repair the specific formatting or insertion problem
5. recheck until the confirmed issues converge

Act as a template-driven formatter and rendered-layout auditor, not a content rewriter.

## Core guardrails

- Use the school template as the layout source of truth.
- For this skill, the only official source-of-truth package is the 2024 revision set under `assets/official-2024/`:
  - `附件1-5.华南农业大学本科毕业论文（设计）撰写规范（封面模板、原创性声明及使用授权声明、正文结构参考样式、参考文献著录规则、评分参考标准）.pdf`
  - `附件6.华南农业大学本科毕业论文（设计）格式模板.doc`
  - `关于印发《华南农业大学本科毕业论文（设计）撰写规范》（2024年修订）的通知.doc`
- Preserve academic meaning. Fix structure, layout, fonts, numbering, captions, bibliography format, and template compliance only.
- Prefer targeted repair over rebuilding the whole document.
- For large local Word files on Windows, do not run one huge all-in-one pass; split into staged passes and save each stage to a new copy with `SaveAs2`.
- Treat rendered pages as the final evidence for pagination, caption placement, hanging indent, continued tables, and page-number alignment.
- Treat Word structure evidence as the final evidence for fonts, bold boundaries, label/body split formatting, TOC special entries, and tail-section style drift.
- If the school rules are silent, mark the point as `以模板为准` rather than inventing a hard rule.

## Bundled template assets

- Default working template: `assets/template/scau-undergrad-thesis-template.docx`
- Original source template: `assets/template/scau-undergrad-thesis-template.doc`
- Preview reference: `assets/template/scau-undergrad-thesis-template-preview.pdf`
- Official 2024 standard package: `assets/official-2024/`

In the public GitHub repository, the official school files and the derived template assets may be absent until the maintainer or user imports them locally.
Before first use in the public repo, run `scripts/import_official_2024_assets.py` to populate `assets/official-2024/` and regenerate the working `.docx` and preview PDF.
The bundled `.docx` and preview PDF, when present, are derived from the official 2024 `附件6` Word template.
If the user supplies a newer official template, only switch after confirming it supersedes the 2024 package and then refresh the comment mapping before bulk insertion.
If the workspace does not follow the original Chinese folder names, keep using explicit paths or let the scripts fall back to the bundled template plus `_scau_thesis_output`.

## Read these references as needed

- `references/workflow.md`
  - use before choosing front matter, chapter, figure, table, or reference insertion mode
- `references/scau-frontmatter-map.md`
  - use before filling cover, Chinese abstract, English abstract, keywords, and author blocks
- `references/scau-template-comments.md`
  - use before touching any teacher-visible template-constrained region
- `references/format-rules.md`
  - use for school rules and confirmed project conventions
- `references/template-comment-rules.md`
  - use for teacher-visible details such as `目录`, `参考文献`, `致谢`, `Abstract:`, `Key words:`, citation punctuation, and subfigure labels
- `references/windows-word-com.md`
  - use when the file is local, large, and needs repeated in-place repair
- `references/project-pipeline.md`
  - use when the current thesis workspace already has stable chapter, figure, table, and reference paths
- `references/reference-import.md`, `references/block-manifest.md`, `references/table-manifest.md`
  - use for bibliography, figure, and table insertion payloads

## Main workflow

1. Identify the mode:
   - initialize a working thesis document
   - fill front matter
   - insert one chapter
   - insert figures or tables
   - insert or reformat references
   - full thesis assembly
   - final format audit
   - audit and repair loop
2. Choose the template source:
   - prefer the bundled `assets/template/scau-undergrad-thesis-template.docx`
   - if the user provides a newer official template, switch to that file
3. Build or refresh the working document:
   - keep a commented working copy
   - keep a separate clean submission copy for the end
4. Fill content into the template:
   - front matter: `scripts/fill_scau_frontmatter.py`
     - it can fill cover metadata, Chinese abstract, Chinese keywords, English abstract, and English keywords when those values are present in the metadata JSON
   - Markdown chapter insertion: `scripts/insert_markdown_chapter.py`
   - figures: `scripts/insert_figure_blocks.py` or `scripts/insert_figure_blocks_com.py`
   - tables: `scripts/insert_table_blocks.py`
   - references: `scripts/insert_reference_batch.py`
   - project-level orchestration: `scripts/run_scau_project_pipeline.py`
5. Export and audit:
   - structure and statistics: `scripts/inspect_word_report.py`
   - font, label/body split, TOC special entries, bibliography/acknowledgement style drift: `scripts/inspect_word_format_signatures.py`
   - Word to PDF: `scripts/export_word_to_pdf.py` or `scripts/export_word_preview.ps1`
   - render pages: `scripts/render_pdf_pages.py`
   - figure block audit when needed: `scripts/inspect_figure_layout.py`
6. Classify each issue:
   - insertion problem
   - template-mapping problem
   - rendered-layout problem
   - font/style drift problem
   - bibliography/citation problem
7. Repair only the right layer:
   - content landed in the wrong structural place: rerun the relevant insertion script
   - TOC, tail sections, page breaks, or local font drift: use `scripts/batch_word_ops.py`
   - citations or bibliography formatting only: fix the reference section and recheck
8. Re-export, rerender, and re-audit until the remaining issues are only `manual_confirm` items or acceptable template choices.
9. Only at the end, create the clean submission copy:
   - `scripts/finalize_submission_copy.ps1`
   - `scripts/strip_docx_comments.py` if direct OOXML comment stripping is safer

## Maintainer acceptance

When updating this skill for open-source reuse, preserve these acceptance conditions:

1. `可复用`
   - do not leave thesis-topic-specific details from any one project
   - the skill must still work when the workspace uses generic names such as `metadata.json`, `work`, or `_scau_thesis_output`
2. `模板吃透`
   - rules must stay aligned with the 2024 official source package and the 50 template comments extracted from `附件6`
   - front-matter anchor mapping must match the current converted template
3. `装版与小修可验证`
   - large chapter backfill from Markdown must still succeed
   - small wording replacement must still preserve body fonts and paragraph formatting

Use `scripts/smoke_test_scau_skill.py` after template, rule, or script changes. The smoke test verifies:

- no project-specific residue in text files
- bundled template comment count still matches the expected 2024 template
- front matter can be filled into the bundled template
- one generic chapter can be inserted
- one local wording replacement can be applied without body-font drift

When validating that a project-specific skill has been generalized, pass the former project terms through repeated `--banned-token` flags so the residue scan stays project-aware without hard-coding those terms into the shared skill.

## Repair routing

Use these default repair routes.

- If the problem is `front matter mapping`, go back to `fill_scau_frontmatter.py`.
- If the problem is `chapter text landed with wrong structure`, go back to `insert_markdown_chapter.py`.
- If the problem is `figure block order or pagination`, go back to the figure insertion step or repair in one Word COM session.
- If the problem is `table continuation or note placement`, go back to table insertion or repair in place.
- If the problem is `目录 special entry spacing`, use `cleanup_contents_entries`.
- If the problem is `参考文献 / 致谢 title-body font drift`, use `normalize_tail_section_fonts`.
- If the problem is `full TOC refresh is too heavy`, do not force a full field update just to clean `参考文献` and `致谢`; use `page_numbers_only` refresh or TOC cleanup alone.

## Word COM repair rules

When the thesis is large or the user is iterating near the end:

- always start from comments + revisions together, not comments only
- in each repair stage:
  - keep one Word session
  - apply only one coherent group of operations
  - save to a new copy with `SaveAs2`
  - re-export once
  - re-audit once
- for the next stage, open the newest copy and repeat
- always disable revision recording before bulk text replacement, otherwise new revisions will keep growing

For `scripts/batch_word_ops.py`, prefer these actions:

- `set_track_revisions`
  - set `enabled: false` before replacement-heavy stages
- `accept_all_revisions`
  - accept legacy revisions before content replacement stages when the user asks for a clean baseline
- `delete_all_comments`
  - optional and usually only for final clean-copy stages
- `replace_text`
  - now uses style-preserving replacement rather than a raw global replace
- `insert_text_after`
  - now copies surrounding font information for inserted text
- `normalize_ascii_digit_font`
  - use for targeted western-letter/digit font normalization
  - default wildcard is `[A-Za-z0-9.]@` (includes periods so decimal values like `0.05` are normalized in one pass)
- `refresh_contents`
  - use `mode: "full"` only when headings really changed
  - use `mode: "page_numbers_only"` when only pagination changed
  - set `update_fields: true` for the final TOC pass so field results settle before cleanup
- `cleanup_contents_entries`
  - removes heading-line spacing leakage from TOC `参考文献` and `致谢`
- `normalize_contents_fonts`
  - only changes TOC range (does not touch end-of-document `参考文献` / `致谢` sections)
  - normalize TOC Chinese chars to `宋体`; English/digits/`.` to `Times New Roman`
- `finalize_contents`
  - runs `refresh_contents(update_fields=true) -> cleanup_contents_entries -> normalize_contents_fonts` in one precise sequence
- `normalize_tail_section_fonts`
  - restores `参考文献` and `致谢` title/body fonts, bibliography hanging indent, acknowledgement first-line indent, and 1.5-line spacing
- `--log-jsonl`
  - write per-stage progress events for long documents so stuck points are traceable

## High-risk final checks

Always prioritize these before declaring the thesis submission-ready:

- `摘        要`, `关键词：`, `Abstract:`, `Key words:` label/body font boundaries
- main body paragraph font pairing, small-four size, and first-line indent
- `参  考  文  献` heading spacing versus TOC `参考文献` no-spacing
- `致        谢` heading spacing versus TOC `致谢` no-spacing
- bibliography entry font pairing, punctuation, hanging indent, and language grouping
- bibliography ordering rule:
  - Chinese first
  - western-language and Russian references second
  - Chinese entries sorted by the first author's surname in Hanyu Pinyin order
  - western-language and Russian entries sorted by the first author's surname in alphabetical order
- citation punctuation width and ordering
- TOC high-risk pair:
  - TOC `参考文献` / `致谢` entries must not keep heading-character spacing
  - TOC font pairing must be Chinese `宋体`, English/digits/`.` `Times New Roman`
- body repeated punctuation check for accidental duplicates such as `。。`, `，，`, `；；`, `：：`
- body paragraph line spacing should remain `1.5` (`LineSpacingRule = 1`), excluding table-cell paragraphs
- figure caption below figure, figure note below caption, no figure-caption split across pages
- continued table headers and table-note placement
- TOC page numbers aligned with the rendered file

## What not to do

- Do not recreate the thesis layout from scratch in a blank document.
- Do not treat PDF text extraction as the main basis for Chinese layout decisions.
- Do not use full TOC refresh as the default final-step cleanup when a lighter cleanup path is enough.
- Do not rewrite prose just to solve a formatting problem.
- Do not remove comments before the working copy has been verified.
- Do not leave `TrackRevisions` enabled during bulk global replacements.
- Do not rely on `document.Save()` for large final-stage documents when `SaveAs2` stage copies are feasible.

## Expected outputs

If the user asks to assemble the thesis:

- return the working Word file path
- return the preview PDF path if exported
- list any confirmed format issues that still remain

If the user asks to audit:

- return a compact report with:
  - basic file facts
  - structure overview
  - statistics if available
  - high-risk issues first
  - issue classification by repair layer
  - `问题 -> 已采取修正 -> 修正后复查结果` when repairs were made

If the user asks for a final submission version:

- produce the clean submission copy
- confirm whether comments were removed
- confirm whether TOC cleanup and tail-section font checks were re-run
- state any remaining `manual_confirm` items explicitly
