# Word COM Mode

Use this mode on Windows when the thesis `.docx` is already large and the current job is dominated by repeated figure insertion or figure revision.

## Why this mode exists

- `python-docx` is still the primary structured-edit backend for chapter Markdown, table conversion, and reference formatting.
- A large thesis working copy can become slow when the same document is rewritten many times just to revise figures.
- The Word COM backend opens the real `.docx` inside Word, inserts figure blocks in one session, and saves once at the end.

## Current official coverage

The COM backend is officially supported for:

- `figure-only` partial-scope jobs
- the figure step inside `scripts/run_scau_project_pipeline.py`

Current entry points:

- `scripts/insert_figure_blocks_com.py`
- `scripts/run_scau_project_pipeline.py --figure-backend word-com`

## What still stays on python-docx

These flows still use `python-docx` as the source of truth:

- frontmatter filling
- Markdown chapter insertion
- standalone table insertion
- reference batch insertion

This is intentional. The COM backend is currently focused on the highest-payoff Windows large-document path rather than replacing every structured edit surface at once.

## Recommended use cases

Prefer Word COM when all of these are true:

- running on Windows with local Microsoft Word installed
- the working thesis document is already large
- the task is mainly figure insertion or repeated figure correction
- you want Word to handle pagination and save behavior directly

Prefer the existing python-docx path when:

- the task is chapter-heavy rather than figure-heavy
- you are running outside Windows Word
- you need deterministic OOXML-level transforms for tables or references

## Notes

- The COM figure backend uses the same manifest schema as `insert_figure_blocks.py`.
- Donor lookup still honors official-template fallback, but it is resolved through live Word paragraphs.
- The runner can mix backends: chapter and tables via `python-docx`, figures via `word-com`.
