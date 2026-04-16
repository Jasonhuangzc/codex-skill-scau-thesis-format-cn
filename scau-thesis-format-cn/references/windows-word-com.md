# Windows Word COM Path

Use this note when the user is on Windows, has local Microsoft Word, and the thesis file is large or frequently edited.

## Why this path matters

- For large `.docx` thesis files with repeated insertions, page-break fixes, and figure-block adjustments, prefer driving Word itself through `pywin32` instead of repeatedly rewriting the OOXML package.
- Keep one Word session open, apply all planned edits, and save once at the end.
- This is especially useful for:
- repeated figure insertion
- moving captions or notes
- adding page breaks before headings
- refreshing contents after edits
- cleaning `参考文献` and `致谢` TOC entries after refresh
- restoring `参考文献` and `致谢` title/body fonts after content edits
- performing a sequence of small format repairs on a large file

## When to prefer Word COM over python-docx

- The source file is already a local Word document.
- The thesis is large, image-heavy, or close to the final submission state.
- The task needs Word-native layout behavior, not just plain paragraph editing.
- Multiple edits must be made in one pass and then re-rendered for recheck.

## Guardrails

- Treat comments and revisions as two separate review layers. Do not assume comments cover all teacher-required changes.
- For final-stage large files, use staged passes instead of one huge pass:
  1. review-state stage (disable tracking / accept old revisions)
  2. main content replacement stage
  3. formatting-only stage (e.g., Latin-name italics, TOC cleanup)
  4. optional English abstract typography stage
- In each stage, open once, run the stage plan, then `SaveAs2` to a new file. Avoid in-place `document.Save()` as the default.
- Prefer the bundled `scripts/batch_word_ops.py` helper when the task can be expressed as a small repair plan.
- Prefer style-preserving replacements and insertions so revised text keeps the surrounding Word font pairing instead of falling back to a wrong default font.
- Do not force a full TOC update when only page numbers or final TOC cleanup are needed; prefer `page_numbers_only` refresh or a standalone TOC cleanup pass.
- Prefer writing to a copy unless the user explicitly wants in-place edits.
- Re-export to PDF and re-render pages after the repair pass.
- For render preparation, prefer `scripts/prepare_render_bundle.py` so Word export runs in a child process with timeout protection.
- If export times out, clean up the Word instance created for that export instead of killing every local Word process.
- If a fresh exported PDF already exists and is newer than the Word source, reuse it to avoid needless repeated export.

## Recommended `batch_word_ops.py` stage plans

Use small JSON plans and run them in separate invocations:

1. Review-state stage:
   - `set_track_revisions` with `enabled: false`
   - `accept_all_revisions`
2. Main replacement stage:
   - one or more `replace_text` actions
   - optional `normalize_ascii_digit_font` for western letters/digits font cleanup
     - recommended wildcard: `[A-Za-z0-9.]@`
     - this includes period `.` so decimal numbers and abbreviations are covered
3. Format cleanup stage:
   - prefer one-shot `finalize_contents` for TOC finalization:
     - `refresh_contents`
     - `cleanup_contents_entries` for TOC `参考文献` / `致谢` de-spacing
     - `normalize_contents_fonts` (TOC-only: Chinese `宋体`, English/digits/`.` `Times New Roman`)
   - `normalize_tail_section_fonts`

For long files, pass `--log-jsonl <path>` to record per-operation start/finish events and locate the exact stuck action.
