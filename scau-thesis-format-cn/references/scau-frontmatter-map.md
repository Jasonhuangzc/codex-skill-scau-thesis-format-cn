# South China Agricultural University Frontmatter Map

This map matches the current converted school template used in this thesis workflow.
Its source-of-truth is the 2024 official `附件6` Word template under `assets/official-2024/`, converted into `assets/template/scau-undergrad-thesis-template.docx`.

## Cover paragraph anchors

These zero-based indices describe the initial converted template only. The filler resolves stable labels on every run and checks cover run positions; indices may change after cover-gap cleanup or multi-paragraph abstracts.

| Paragraph index | Expected anchor text | Replacement rule |
| --- | --- | --- |
| 1 | `本科毕业论文(或设计)` | Replace with `本科毕业论文` or `本科毕业设计` |
| 3 | `论文（或设计）题目` | Replace with Chinese thesis title |
| 10 | `学    院:` | Replace run 6 with college full name |
| 11 | `专    业:` | Replace run 5 with major full name |
| 12 | `姓    名:` | Replace run 5 with student Chinese name |
| 13 | `学    号:` | Replace run 5 with student ID |
| 14 | `指导教师:` | Replace run 4 with advisor name and run 9 with title |
| 15 | `提交日期：` | Replace runs 2, 5, and 9 with year, month, day |

## Abstract anchor paragraphs

| Paragraph index | Expected anchor text | Replacement rule |
| --- | --- | --- |
| 38 | `摘        要` | Keep heading, use as Chinese abstract anchor |
| 39 | sample Chinese abstract text | Replace with real Chinese abstract |
| 41 | `关键词：` sample | Replace with Chinese keywords |
| 42 | `English Title` | Replace with English thesis title |
| 43 | `Song Nianxiu` | Replace with English author name |
| 44 | affiliation sample | Replace with English affiliation line |
| 45 | `Abstract:` sample | Replace with English abstract |
| 47 | `Key words:` sample | Replace with English keywords |

## Safety checks before writing

Before using the map:

- identify the unique cover labels, declaration heading, abstract heading, keyword labels and Abstract label
- locate the cover title between paper type and college, and the three English title/author/affiliation lines before Abstract
- validate cover run positions and non-empty trailing/split field runs; stop on missing/ambiguous anchors or a changed run map
- do not require historical paragraph indices when processing an already-filled working copy

If the template changed, refresh the map with `scripts/extract_docx_comments.py` and update the workflow instead of forcing the old indices.

## Repeat updates and source values

`fill_scau_frontmatter.py` preserves metadata-omitted abstract and keyword blocks. Supplied abstracts are written as actual paragraphs, retaining the label/body font distinction; explicitly empty supplied content is rejected. Chinese and English abstracts may grow beyond their initial two donor paragraphs without shifting later labels. The cover metadata and English identity/affiliation keys remain required by this script.

Blank cover-gap paragraphs containing section properties are retained. English title `pageBreakBefore` is a chosen stable layout implementation; template compliance still needs rendered page review. These synthetic regression checks do not verify a newly imported official template.
