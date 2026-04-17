# Project Pipeline

Use this reference when the thesis workspace already has stable paths and the work should be executed through the project runner:

- `scripts/run_scau_project_pipeline.py`

The default source-of-truth package remains the 2024 official files under `assets/official-2024/`.

## Official partial scopes

The runner now officially supports these safe scopes:

- `frontmatter-only`
  Keep `--skip-chapter --skip-figures --skip-tables --skip-references`
- `chapter-only`
  Keep `--skip-frontmatter --skip-figures --skip-tables --skip-references`
- `figure-only`
  Keep `--skip-frontmatter --skip-chapter --skip-tables --skip-references`
- `table-only`
  Keep `--skip-frontmatter --skip-chapter --skip-figures --skip-references`
- `reference-only`
  Keep `--skip-frontmatter --skip-chapter --skip-figures --skip-tables`

## Important runner behavior

- If the current working docx already exists, `frontmatter` refresh operates on that working copy rather than regenerating from the official template from scratch.
- If the working copy lacks figure or table donor paragraphs, the runner passes `--official-template-docx` through to the block inserters as fallback donor source.
- Figure manifest generation prefers chapter-text anchors such as `如图3-1所示` over a broad `图3-1` regex.
- The runner can route only the figure step through Word COM by passing `--figure-backend word-com`.
- On Windows, the runner now ends with a TOC finalization pass:
  - ensure the English abstract starts with an explicit page break
  - normalize正文 paragraphs back to first-line indent `2` and `1.5` line spacing
  - normalize table-cell paragraphs to centered alignment
  - normalize the `英文缩略词（符号表）` table to `宋体 + Times New Roman` small-four, centered, `1.5` line spacing
  - update fields
  - clean TOC `参考文献` / `致谢` entries
  - normalize TOC Chinese characters to `宋体`
  - normalize TOC English, digits, and `.` to `Times New Roman`

## Recommended pattern

For a workspace that already follows one stable thesis-folder layout, prefer:

- `--official-template-docx` pointing to `论文撰写规范\附件6_格式模板_转存.docx`
- `--metadata-file` pointing to `thesis_metadata.json`
- `--chapter-file` pointing to the current chapter Markdown
- `--figure-number-prefix` matching the chapter number when figure folders are grouped by chapter
- when the working document is already large and the current pass is mainly figure insertion, add `--figure-backend word-com`

These are convenience defaults, not hard requirements. If the workspace uses other names, pass explicit paths or let the runner fall back to the bundled template and `_scau_thesis_output`.

## Backend boundary

- `python-docx` remains the primary backend for frontmatter, chapters, tables, and references.
- `word-com` is currently the official Windows large-document backend for figure insertion.
- Mixed runs are supported. A chapter can still be inserted through `python-docx`, and the figure step can then switch to `word-com` in the same runner invocation.

## Failure handling

When a sub-step fails, the runner reports:

- current step name
- executed command
- stdout and stderr
- a recovery hint for that step

This is the preferred entry point whenever the job is larger than a single isolated insertion.
