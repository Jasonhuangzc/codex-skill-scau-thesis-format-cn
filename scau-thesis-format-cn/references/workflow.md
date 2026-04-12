# Workflow

Use this sequence whenever the task is to move thesis material into a Word template.

For this skill, the default standard package is the 2024 official set under `assets/official-2024/`. Do not silently mix in older school templates or unofficial local copies.

## 1. Discover inputs

Look for these files first:

- metadata: `thesis_metadata.json`
- official template: converted `.docx` version of the school template
- current working document: often under `论文终稿`, `work`, or `_scau_thesis_output`
- chapter drafts: Markdown, `.docx`, or direct user text
- figure assets: image files under `Data`, `figures`, `图片`, or chapter-specific folders
- table material: editable tables, spreadsheet exports, or user text

These are discovery preferences, not required names. The scripts also accept explicit paths and can fall back to the bundled template plus a generic output directory.

If only a `.doc` template exists, convert it to `.docx` with Word before doing any structured edits.

## 2. Choose the mode

Use one mode at a time:

- `frontmatter-bootstrap`
  Fill cover, Chinese abstract, English abstract, and keyword blocks.
  Runner equivalent: keep `--skip-chapter --skip-figures --skip-tables --skip-references`.
- `chapter-insert`
  Insert or revise chapter text while preserving heading and body styles.
  Runner equivalent: keep `--skip-frontmatter --skip-figures --skip-tables --skip-references`.
- `figure-insert`
  Insert figures, captions, subfigure labels, and figure notes.
  Runner equivalent: keep `--skip-frontmatter --skip-chapter --skip-tables --skip-references`.
  For large Windows working documents with repeated figure revisions, prefer the Word COM path:
  `scripts/insert_figure_blocks_com.py` or runner `--figure-backend word-com`.
- `figure-grid-compose`
  Compose multi-panel figures such as grouped HE sections or fluorescence panels into a final combined image.
- `figure-manifest-build`
  Scan one-folder-per-figure directories and auto-build the JSON manifest that `insert_figure_blocks.py` needs.
- `table-insert`
  Insert tables, table titles, continued tables, and table notes.
- `table-block-insert`
  Insert tables from table manifests or standalone Markdown table files, including automatic continued-table splitting.
  Runner equivalent: keep `--skip-frontmatter --skip-chapter --skip-figures --skip-references`.
- `chapter-markdown-insert`
  Insert a whole chapter from Markdown, including heading levels and pipe tables.
- `reference-batch-insert`
  Replace the template bibliography sample with the final verified reference draft and re-apply hanging indent.
  Runner equivalent: keep `--skip-frontmatter --skip-chapter --skip-figures --skip-tables`.
- `project-pipeline-run`
  Run one thesis-specific chapter pass or one official partial-scope operation when workspace paths are already stable.
- `final-clean`
  Refresh and verify the document, then create a clean copy if requested.

## 3. Read the constraint layer

Before bulk edits, inspect the template comments:

- if the template matches the bundled South China Agricultural University template, read `scau-template-comments.md`
- if the template may have changed, run `scripts/extract_docx_comments.py` on the real template and compare the output
- when cover or abstract anchor positions matter, also check `scau-frontmatter-map.md`

Treat comments as binding instructions unless the user explicitly says the school template has been superseded.

## 4. Edit in the safe order

Use this order:

1. fill front matter
2. insert chapter headings and body text
   - use `scripts/insert_markdown_chapter.py` when the source chapter is Markdown
3. insert figures and tables
   - use `scripts/insert_table_blocks.py` when table captions, notes, or continued-table behavior need to be controlled explicitly
   - use `scripts/generate_figure_manifest_from_dirs.py` when the workspace keeps each figure in a dedicated folder with a final PNG and a `_图注.txt` file
   - use `scripts/compose_panel_grid.py` or the `layout: "grid"` manifest mode for multi-panel figures
   - use `scripts/insert_figure_blocks.py` for image blocks
   - on Windows large-document figure jobs, switch only the figure step to `scripts/insert_figure_blocks_com.py`
4. update references and back matter
   - use `scripts/insert_reference_batch.py` for the bibliography
5. when the workspace paths are already stable and the task is one whole chapter pass or a safe partial scope
   - use `scripts/run_scau_project_pipeline.py`
   - for runner mode details, read `project-pipeline.md`
   - for large Windows figure insertion, add `--figure-backend word-com`
6. export preview and verify
7. make a clean copy only at the end
   - use `scripts/strip_docx_comments.py` when the goal is only to remove comments
   - use `scripts/finalize_submission_copy.ps1` when comments, fields, and directory all need a Word-based refresh

Do not start by deleting sample content across the whole document. Replace the document section by section.

## 5. Handle comments safely

- Keep comments in the working copy.
- Do not remove comments just because the corresponding text was replaced.
- If a comment is attached to a paragraph that acts as a style or structure donor, keep that paragraph position stable.
- Only create a clean no-comment copy after the user asks for the submission file.

## 6. Verify rendered output

Use rendered output for these checks:

- figure and caption are on the same page
- table continuation and repeated table header
- bibliography hanging indent
- abstract block spacing
- table and figure note placement
- chapter heading pagination
- directory refresh status

Use `scripts/export_word_preview.ps1` to export PDF preview. Then inspect visually or use the `doc` skill if available.

## 7. Final pass checklist

Confirm:

- title pages and declarations follow the template
- abstracts and keywords match school rules
- chapter numbering is continuous
- figure and table numbering is continuous
- references heading and bibliography formatting are correct
- appendix numbering follows appendix rules
- acknowledgements heading is correct
- directory has been refreshed after正文 changes
