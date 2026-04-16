# South China Agricultural University Thesis Format Rules

Use this file when checking a South China Agricultural University undergraduate thesis against the official template and related comment-derived conventions.

The only source-of-truth set for this skill is the 2024 revision package:

- `assets/official-2024/附件1-5.华南农业大学本科毕业论文（设计）撰写规范（封面模板、原创性声明及使用授权声明、正文结构参考样式、参考文献著录规则、评分参考标准）.pdf`
- `assets/official-2024/附件6.华南农业大学本科毕业论文（设计）格式模板.doc`
- `assets/official-2024/关于印发《华南农业大学本科毕业论文（设计）撰写规范》（2024年修订）的通知.doc`

The bundled working template `assets/template/scau-undergrad-thesis-template.docx` is a converted derivative of the official 2024 `附件6` Word file.
In the public repository, these files are expected to be imported locally through `scripts/import_official_2024_assets.py` before first use.

For template-comment-sensitive items such as `目录`, `参考文献`, `致谢`, keywords, citation punctuation, and subfigure labels, also read `references/template-comment-rules.md`.

## 1. Confirmed school rules from the PDF

These rules are explicit in the 2024 `附件1-5`.

### 1.1 Body structure

- The school gives a recommended farm/science structure:
  - `1 前言`
  - `2 材料与方法`
  - `3 结果与分析`
  - `4 讨论与结论`
- This is a reference structure for the正文. Other format rules still follow the school-wide requirements.

### 1.2 In-text citation rules

- Use the `著者-出版年` system.
- Citation parentheses use Chinese full-width round brackets.
- If the author is mentioned in the sentence, put the year in parentheses after the author name.
- If the author is not mentioned in the sentence, put `作者, 年份` in parentheses after the cited statement.
- For two authors:
  - Chinese use `和`
  - English use `and`
- For three or more authors:
  - Chinese use `等`
  - English use `et al.`
- For the same author with multiple years, list later items by year only.
- For the same author in the same year, distinguish entries with `a`, `b`, `c`.
- For one location citing multiple sources:
  - list Chinese first, then foreign-language references
  - separate items with semicolons
- Inside citation parentheses, punctuation uses half-width English symbols.
- After commas and semicolons inside citation parentheses, keep one half-width space.
- When quoting a precise page repeatedly, use `年份: 页码`.

### 1.3 Bibliography list rules

- List only sources actually cited in the正文.
- Do not number bibliography entries.
- Group by language in this order for this skill:
  - Chinese
  - western-language and Russian entries as the foreign-language block
  - if another foreign-language block appears, treat it as part of the foreign-language block unless the school later publishes a stricter sub-order rule
- Chinese entries are sorted by the first author's pinyin surname.
- Western-language entries are sorted by the first author's surname.
- Chinese bibliography text uses `宋体` small-four.
- Western text and digits use `Times New Roman` small-four.
- Punctuation in both Chinese and English bibliography entries uses half-width `Times New Roman` symbols.
- Every bibliography entry ends with a half-width period.
- After commas, periods, colons, and semicolons in bibliography entries, keep one half-width space where the template examples show it.
- For authors:
  - list all authors when there are up to three
  - list the first three only when there are more than three
  - add `等` for Chinese or `et al.` for English
- Western author names use `surname + initials`; initials are uppercase and have no abbreviation dots.
- Foreign journal names may use the full title or a standard abbreviation.
- If abbreviated, do not use abbreviation dots, and separate words with one space.
- Use hanging indent of two characters for each bibliography entry.
- Two-character Chinese names do not have an internal space.
- Page ranges use the short hyphen `-`, not wave dashes.
- Between adjacent brackets such as `) [` keep one half-width space.
- URLs do not add spaces around punctuation.

### 1.4 Reference entry patterns explicitly given in the PDF

- Book: `作者. 书名[M]. 版次. 出版地: 出版者, 出版年: 起止页码.`
- Chapter in book or conference volume: `析出作者. 析出题名[类型]//原文献作者. 原文献题名. 出版地: 出版社, 出版年: 起止页码.`
- Journal article: `作者. 题名[J]. 期刊名, 出版年, 卷号(期号): 起止页码.`
- Newspaper: `作者. 题名[N]. 报刊名, 年-月-日(版次).`
- Thesis: `作者. 题名[D]. 授予单位所在地: 授予单位, 授予年份: 起止页码.`
- Report: `作者. 题名[R]. 报告地: 报告会主办单位, 年份.`
- Online material: `作者. 题名[文献类型/OL]. (上传或更新日期) [引用日期]. 获取和访问路径.`

## 2. Template examples from the DOC

These are format examples from the 2024 `附件6`. Use them as layout evidence when the PDF is silent.

### 2.1 Front matter and abstracts

- The Chinese abstract page uses:
  - `摘        要`
  - abstract body
  - `关键词：关键词；关键词；...`
- The Chinese abstract is expected to stay around `300-600` words and normally does not cite references.
- The English abstract section uses:
  - English title
  - author name
  - affiliation line in parentheses
  - `Abstract:`
  - `Key words:`
- In the English abstract body, commas, periods, colons, and semicolons are followed by one half-width space.
- The second and later English abstract paragraphs use a two-character first-line indent according to the template comment.
- The template includes an optional `英文缩略词（符号表）` section.
- The template includes `目录`, then the正文.
- The template shows a front-matter sequence of:
  - cover
  - originality statement
  - authorization statement
  - Chinese abstract
  - English abstract
  - optional abbreviation list
  - table of contents
  - body
  - references
  - appendices
  - acknowledgements
- The statement pages do not carry page numbers.

### 2.2 Heading levels

- The template shows heading levels such as:
  - `1`
  - `1.1`
  - `1.1.1`
  - `1.1.1.1`
- When checking a draft, focus on continuity and style consistency rather than forcing unnecessary extra levels.
- Level-1 heading: `黑体 + Times New Roman`, four-point size class, 1.5-line spacing, left aligned.
- Level-2 heading: `黑体 + Times New Roman`, small-four, 1.5-line spacing, left aligned.
- Level-3 heading: `楷体 + Times New Roman`, small-four, 1.5-line spacing, left aligned.
- Level-4 heading: `楷体 + Times New Roman`, small-four, 1.5-line spacing, left aligned.
- The heading number and the heading text are separated by one character space.
- Body text uses `宋体 + Times New Roman`, small-four, first-line indent of two characters, and 1.5-line spacing.

### 2.3 Figures, tables, formulas

- Table title example: `表1  表名`
- Continued table example: `续表2  表名`
- Table note example: `注：...`
- Figure caption example: `图1  图名`
- Figure note example: `注：...`
- Subfigure example in the template: `(a)  分图名`, `(b)  分图名`
- Formula numbering example: `（式1）`
- Tables and figures leave one blank line from surrounding正文 in the layout example comments.
- Tables use a three-line-table structure with heavier top and bottom rules and a lighter middle rule.
- Table body text uses five-point size class and single spacing in the template comment.
- Table notes use small-five in the template comment.
- Figures and their main captions must not be split across pages.
- Formula blocks are centered, numbered continuously, and their numbers sit at the right end of the line.
- Footnotes are continuous and use small-five single spacing.

### 2.4 Sensitive comment-derived checks

These come from explicit comments in the 2024 `附件6` and should be treated as high-priority final audit checks.

- `关键词`:
  - the label `关键词：` keeps the template label/body distinction and should not force the whole line into one uniform emphasis style
  - Chinese keywords are separated by full-width semicolons
  - the last keyword has no punctuation
- `Key words`:
  - the label `Key words:` is explicitly bold in the template comment, but the keyword content after it is not
  - English keywords are separated by half-width semicolons
  - the last keyword has no punctuation
  - each content word begins with an uppercase letter
- `英文摘要格式边界`:
  - `Abstract:` is explicitly bold
  - the abstract body immediately after `Abstract:` is not
  - English title is explicitly bold
  - English author and affiliation are not explicitly marked as bold in the template comment
- `目录`:
  - update the table of contents after heading edits
  - the final TOC pass is: update fields -> clean the TOC `参考文献` / `致谢` entries -> normalize TOC fonts
  - after each TOC update, re-check and clean the `参考文献` and `致谢` entries so they do not keep the heading-line character spacing
  - show contents down to level 3
  - keep left-aligned entries, right-aligned page numbers, and leader dots
  - `参考文献` and `致谢` in the contents do not insert character spacing
  - TOC Chinese characters use `宋体`; English, digits, and `.` use `Times New Roman`
- `目录标题`:
  - the title itself uses `目        录`
- `参考文献` title:
  - the heading itself uses spaced characters in the title line according to the template comment
  - the title line keeps heading fonts, but bibliography entry paragraphs return to `宋体 + Times New Roman` small-four
- `致谢` title:
  - the heading itself uses the template-spaced title line, but the contents entry keeps `致谢` without inserted character spacing
  - the title line keeps heading fonts, but the acknowledgement body returns to `宋体 + Times New Roman` small-four
- Citation examples:
  - citation brackets are full-width
  - punctuation inside citation brackets is half-width
  - comma and semicolon are followed by one half-width space
- Bibliography examples:
  - punctuation is half-width
  - punctuation spacing follows the template comments
- Subfigure labels:
  - the template comment uses `(a)`, `(b)`, `(c)` style directly under each panel
- `图表版式`:
  - figure and table blocks keep blank lines from正文
  - figures do not split from their captions across pages
  - continued tables repeat the table header
- `页码`:
  - statement pages do not show page numbers
- `公式与脚注`:
  - formula numbers are right aligned without leader dashes
  - footnotes use continuous numbering

## 3. Reusable conventions for SCAU science theses

These conventions are reusable defaults rather than current-project assumptions.

- A common science-thesis structure is:
  - `前言`
  - `材料与方法`
  - `结果与分析`
  - `讨论与结论`
- If the user later splits conclusion into a separate chapter, check continuity rather than forcing one fixed structure.
- Figure captions and notes should stay in Chinese unless the school or department explicitly requires another language.
- For final Huanong Word-format audit, prefer the template-comment subfigure style `(a)`, `(b)`, `(c)` unless the supervisor explicitly fixes another house rule.
- Put subfigure labels under each panel and the main figure title under the whole figure.
- Do not embed the main figure title inside the image.
- For editable figure text, use:
  - Chinese: `宋体`
  - western letters and digits: `Times New Roman`
- Keep table titles above tables and notes below tables.

## 4. Final Word audit priorities

When the user provides a Word file for final review, prioritize these checks in order:

1. Rendered-page checks
   - page order
   - whether major sections start consistently
   - page numbering continuity
   - no-page-number handling on statement pages
   - table of contents alignment with actual headings
   - figure caption and table title positions
   - hanging indent in the bibliography
   - continued-table layout
   - figure or caption page splitting
   - teacher-visible punctuation and bracket details in citations, bibliography, keywords, and subfigure labels
   - TOC cleanup state for `参考文献` and `致谢`
   - post-edit font drift in正文, bibliography, and acknowledgement paragraphs
2. Text-structure checks
   - heading numbering continuity
   - heading-style consistency by level
   - figure and table numbering continuity
   - abstract and keyword labels
   - appendix and acknowledgement labels
3. Bibliography and citation checks
  - author-year format in the body
  - entry completeness and punctuation
  - Chinese versus English entry style differences
  - bibliography ordering:
    - Chinese references first
    - western-language and Russian references after Chinese
    - Chinese entries sorted by the first author's surname in Hanyu Pinyin order
    - western-language and Russian entries sorted by the first author's surname in alphabetical order
4. Terminology and mixed-language consistency
  - abbreviations
  - species names
  - concentration and unit formats
  - repeated English term variants
5. High-visibility cleanup checks
   - TOC `参考文献` / `致谢` entry spacing
   - TOC Chinese-vs-western font pairing
   - repeated punctuation such as `。。` and `，，`
   - body paragraph `1.5` line spacing, excluding table-cell paragraphs

If the rendered file has not been inspected, do not mark the thesis as fully compliant.

### 4.1 Render priority and fallback

- Preferred evidence chain:
  - `Word report`
  - `exported PDF`
  - `rendered page images`
- If `pdftoppm` or Poppler is unavailable, render PDF pages with PyMuPDF instead of stopping the audit.
- If PDF text extraction is garbled, keep the rendered-page audit moving; do not use garbled extraction as the main reason to stop.

### 4.2 Figure-special audit priorities

When the task is figure-only review, change the order to:

1. figure page mapping
2. figure body, caption, and note order
3. split risk across pages
4. next-heading start risk
5. readability manual confirmation
6. broader textual consistency only after the figure block is stable

### 4.3 Template boundary handling

- `封面`, `原创性声明`, `使用授权声明`, and `目录` remain hard checks.
- `英文缩略词`, `附录`, `致谢`, and still-unfinished `参考文献` may be reported as:
  - `保留模板占位 / 本轮未处理`
  - not as direct format errors by default
- Distinguish:
  - real formatting defects
  - content completion issues
  - template placeholders intentionally left for a later round

## 5. Report expectations

For a final thesis format audit, the report should ideally include:

1. Audit target basics
   - file name
   - source type
   - whether the judgment is based on Word text only, rendered PDF pages, or both
2. Structure overview
   - presence of cover, statements, abstracts, contents, body, references, appendices, acknowledgements
   - chapter count and heading depth
3. Basic statistics
   - page count
   - Word statistics
   - Chinese abstract character count
   - English abstract word count
   - heading counts by level
   - figure, table, formula, footnote, and reference counts where derivable
4. Format completion assessment
   - front matter
   - contents
   - body headings and paragraphs
   - figures
   - tables
   - formulas and footnotes
   - references
   - appendices and acknowledgements
5. High-risk issues summary
   - teacher-visible punctuation and bracket issues
   - page split, caption placement, TOC refresh, and page number problems
6. Detailed issues list
7. Still-unverified items
   - anything blocked by missing render output or ambiguous supervisor requirements

### 5.1 Repair-closure reporting

For each high-risk issue, prefer this structure:

1. `问题`
2. `已采取修正`
3. `修正后复查结果`

When possible, propose the smallest effective repair action rather than a vague suggestion.

### 5.2 Structured conclusion levels

Use these levels consistently in machine-readable outputs:

- `confirmed`
- `suggested`
- `manual_confirm`

### 5.3 Figure-special structured outputs

For figure audit outputs, include at least:

- `figure_page_map`
- `render_basis`
- `caption_order_status`
- `note_order_status`
- `split_risk`
- `next_heading_risk`
- `page_images`
- `still_manual_confirm`

## 6. Ambiguity handling

- Treat the PDF as the authority for explicit school rules, especially bibliography and citation format.
- Treat the DOC as the authority for layout examples such as abstract blocks, heading levels, figure captions, table captions, and continued tables.
- If the PDF is silent and the DOC only shows an example rather than a hard rule, prefer consistency with the template and with the rest of the thesis.
- If the project convention conflicts with the final official template or the supervisor's explicit requirement, follow the final required version consistently throughout the thesis.
