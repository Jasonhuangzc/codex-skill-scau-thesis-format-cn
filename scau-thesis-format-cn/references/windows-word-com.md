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

- Open the document once.
- Apply all planned operations in sequence.
- Save once at the end.
- Prefer the bundled `scripts/batch_word_ops.py` helper when the task can be expressed as a small repair plan.
- Prefer style-preserving replacements and insertions so revised text keeps the surrounding Word font pairing instead of falling back to a wrong default font.
- Do not force a full TOC update when only page numbers or final TOC cleanup are needed; prefer `page_numbers_only` refresh or a standalone TOC cleanup pass.
- Prefer writing to a copy unless the user explicitly wants in-place edits.
- Re-export to PDF and re-render pages after the repair pass.
- For render preparation, prefer `scripts/prepare_render_bundle.py` so Word export runs in a child process with timeout protection.
- If export times out, clean up the Word instance created for that export instead of killing every local Word process.
- If a fresh exported PDF already exists and is newer than the Word source, reuse it to avoid needless repeated export.
