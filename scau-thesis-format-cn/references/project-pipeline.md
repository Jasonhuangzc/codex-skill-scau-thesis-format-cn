# Project Pipeline

Use `scripts/run_scau_project_pipeline.py` for a complete chapter pass or one partial operation. The authoritative package remains the 2024 official files under `assets/official-2024/`; an explicitly supplied template must be checked against that package before use.

## Supported scopes

| Scope | Flags to keep |
| --- | --- |
| frontmatter-only | `--skip-chapter --skip-figures --skip-tables --skip-references` |
| chapter-only | `--skip-frontmatter --skip-figures --skip-tables --skip-references` |
| figure-only | `--skip-frontmatter --skip-chapter --skip-tables --skip-references` |
| table-only | `--skip-frontmatter --skip-chapter --skip-figures --skip-references` |
| reference-only | `--skip-frontmatter --skip-chapter --skip-figures --skip-tables` |
| copy/no-op | all five `--skip-*` flags |

Partial scopes require an existing working `.docx`. Metadata is loaded only for frontmatter work. A converted official template is required for first bootstrap or explicit template-body trimming; other insertions can use an existing document's donors without an extra template. If a template is available, missing chapter/figure/table donor styles can fall back to it. Only the donor types actually used by the chapter are required.

## Content preservation

- Frontmatter refresh starts from the existing working copy when it exists.
- Chapter insertion replaces the unique body chapter with the same chapter number, even if its title changes. Repeating the same backfill does not add another chapter.
- Existing working documents keep all other chapters by default. The runner no longer deletes everything between `1 绪论` and the inserted chapter.
- On first bootstrap directly from the official template, an inserted chapter can replace the template's sample body. `--keep-template-body` preserves those samples instead.
- `--trim-template-body` explicitly removes a sample body only when every XML body block and its linked asset payload matches the selected official template. Changed text, tables, images, or section boundaries cause a refusal. Trimming includes tables and content controls, rather than removing only paragraphs.
- Unresolved tracked revisions block structural backfill. Chapter replacement also refuses an ambiguous/duplicate chapter, a missing following section boundary, or a target range containing fields, section boundaries, embedded objects, footnotes, or equations. Handle those cases through a targeted Word edit first. Word heading/outline properties identify following chapters; a directly formatted numbered paragraph with an ambiguous boundary causes a refusal instead of a guessed deletion range.
- Existing chapter images are protected by default. Before explicitly using `--replace-media`, check a complete manifest mapping each existing image, number, caption, note, and source asset to its rebuilt counterpart. The runner passes this flag through only when explicitly requested; enabling figure insertion alone does not grant it.
- Existing chapter tables are protected if the Markdown contains no table payload. Before explicitly using `--replace-tables`, check a complete reconstruction manifest for each table's caption, data, note, and continuation behavior. Footnotes, equations, fields, and embedded objects have no structural-deletion override. A text-only revision of such a chapter should use a targeted Word edit.
- Output must be separate from the input document and official template. The completed result is published atomically only after all requested steps succeed; a failed later step leaves a previous output intact.

## Markdown chapter contract

The draft starts with exactly one numbered level-1 chapter title, such as `# 第3章 结果与分析` or `# 3 结果与分析`. Level 2–4 headings use numbers belonging to that chapter, with matching depth (`## 3.1`, `### 3.1.1`). Duplicate heading numbers, another chapter's subheading, and multiple level-1 chapters are rejected before insertion. If supplied, `--chapter-heading` must agree with the draft.

TOC styles (including inherited styles), multi-paragraph TOC fields, and TOC content controls are excluded when selecting body anchors. A real Word heading using a tab between its number and title is still a body heading; a tab alone is not treated as directory evidence.

Supported content includes paragraphs, inline emphasis, headings, and rectangular pipe tables. Table separators can have or omit outer pipes; `\|` preserves a literal pipe in a cell. Ragged rows fail with a line number. Fenced code, image Markdown, and headings beyond level 4 need a separate/manual treatment instead of silently entering the thesis as literal markup. Use figure manifests for images.

Put a numbered table title immediately before its table and an optional `注：`/`资料来源：` paragraph after it. The importer renders the title above the table, keeps it with the following table, repeats headers and titles for continued tables, preserves all data rows, and leaves one blank line around a complete titled table block. Estimated table splitting still needs rendered-page review.

## Directory refresh and backend boundary

`python-docx` is the primary backend for frontmatter, chapters, tables, and references. `--figure-backend word-com` selects the Windows Word COM figure backend for large documents; `--word-visible` displays its Word window.

The runner reports `contents_refresh_required` after content changes. A full assembly must refresh and verify the directory/fields before submission. On Windows with Word and pywin32, `--finalize-contents` opts into a final field/TOC refresh and TOC spacing/font cleanup. It does not perform a broad normalization of unrelated body, table, or tail-section layout. On other platforms, refresh fields in Word before export and keep that check unconfirmed until verified.

Figure-manifest generation prefers chapter-text anchors such as `如图3-1所示`. `--skip-generate-figures` avoids rerunning generation scripts while still rebuilding the manifest. For local chapter revisions, skip figure work when its assets and anchors are unchanged.

## Output and recovery

The final stdout is one JSON report including the output, selected scope, `steps_run`, whether samples were trimmed, and the field-refresh state. Stage progress goes to stderr.

On a staged failure, the error JSON records the step, command, stdout/stderr, and recovery hint. It also includes `recovery_dir` and `last_successful_docx`; keep that intermediate copy to rerun only the failed scope. The original input remains preserved, and `output_published` is false. Successful runs remove their temporary working directory.

Cross-platform content-preservation regressions run without Word or imported official assets:

```bash
python -m unittest discover -s tests -p 'test_backfill.py' -v
```

These tests verify semantic block order, source preservation, repeated replacement, safe failure, Markdown validation, donor font pairing, and sample-body cleanup. They do not establish pagination or final visual compliance; export and review the finished Word/PDF separately.
