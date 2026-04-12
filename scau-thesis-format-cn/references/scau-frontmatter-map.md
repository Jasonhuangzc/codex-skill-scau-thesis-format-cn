# South China Agricultural University Frontmatter Map

This map matches the current converted school template used in this thesis workflow.
Its source-of-truth is the 2024 official `附件6` Word template under `assets/official-2024/`, converted into `assets/template/scau-undergrad-thesis-template.docx`.

## Cover paragraph anchors

Paragraph indices are zero-based within `Document.paragraphs`.

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

- confirm the template paragraph count still covers these anchors
- confirm the expected anchor text still appears in the mapped paragraphs
- stop if the anchor text has moved or been replaced by another template version

If the template changed, refresh the map with `scripts/extract_docx_comments.py` and update the workflow instead of forcing the old indices.
