# South China Agricultural University Template Comments

This file captures the comments embedded in the current converted school template.
The current source-of-truth template is the 2024 official `附件6` under `assets/official-2024/`, and the bundled working `.docx` is converted from that file.

The `Anchor paragraph` column follows the full document paragraph order produced by `scripts/extract_docx_comments.py`. It is not always the same as `python-docx` `Document.paragraphs` after tables appear. For frontmatter scripting, use `scau-frontmatter-map.md`.

## Cover and front matter

| Comment ID | Anchor paragraph | Anchor text | Rule |
| --- | --- | --- | --- |
| 0 | 1 | `本科毕业论文(或设计)` | 宋体，小初号，加粗，居中；毕业论文写“本科毕业论文”，毕业设计写“本科毕业设计”。 |
| 1 | 3 | `论文（或设计）题目` | 题目黑体二号，加粗，居中。 |
| 2 | 10 | `学    院:` | 宋体，小三号，学院全称。 |
| 3 | 11 | `专    业:` | 宋体，小三号，专业全称。 |
| 4 | 12 | `姓    名:` | 宋体，小三号。 |
| 5 | 13 | `学    号:` | Times New Roman，小三号。 |
| 6, 7 | 14 | `指导教师:` | 宋体，小三号。 |
| 8 | 15 | `提交日期：` | 日期为 Times New Roman，小三号。 |
| 26 | 37 | blank declaration page paragraph | 声明页不加页码。 |
| 28 | 38 | `摘        要` | 标题“摘要”中间空 4 个字距，黑体四号，居中，1.5 倍行距。 |
| 29 | 39 | Chinese abstract sample | 中文摘要正文：宋体，西文 Times New Roman，小四号，两端对齐，首行缩进 2 字距，1.5 倍行距；300 至 600 字；不引用参考文献。 |
| 30 | 41 | `关键词：关键词；关键词；关键词；关键词；关键词` | “关键词：”黑体，小四号；关键词宋体，西文 Times New Roman，小四号；3 至 5 个；关键词间用全角分号；末尾不加标点。 |
| 31 | 42 | `English Title` | 英文题目 Times New Roman，四号，加粗，居中；实词首字母大写。 |
| 32 | 43 | `Song Nianxiu` | 作者英文名 Times New Roman，小四号，居中；姓和名首字母大写。 |
| 33 | 44 | affiliation sample | 作者单位含学院、学校、广州、邮编、中国；Times New Roman，小四号，前后括号，居中。 |
| 34 | 45 | English abstract sample | `Abstract:` 加粗；英文摘要 Times New Roman，小四号，两端对齐，1.5 倍行距；第二段及以后首行缩进 2 字距；英文标点后有一个半角空格。 |
| 35 | 47 | `Key words: ...` | `Key words:` 加粗；关键词 Times New Roman，小四号；关键词间用半角分号；末尾不加标点；实词首字母大写。 |
| 43 | 48 | `英文缩略词（符号表）` | 标题黑体四号，居中；表格为三线表；中文宋体小四号，英文 Times New Roman 小四号，居中。 |
| 44 | 73 | `目 录` | 标题“目录”空 4 个字距，黑体四号；目录显示到 3 级；小四号；页码右对齐，有前导符；正文修改后刷新目录。 |
| 45 | unanchored | no anchor in current XML paragraph scan | 目录中的“参考文献”“致谢”字间不空格；成绩评定表不编入目录。 |

## Headings,正文, figures, tables, formulas, footnotes

| Comment ID | Anchor paragraph | Anchor text | Rule |
| --- | --- | --- | --- |
| 47 | 98 | `1 绪论` | 一级标题：黑体四号，西文 Times New Roman，左对齐，1.5 倍行距；题序与标题间空 1 个字距。 |
| 50 | 99 | `1.1 标题` | 二级标题：黑体小四号，左对齐，1.5 倍行距；题序与标题间空 1 个字距。 |
| 51 | 100 | body sample | 正文：中文宋体，西文 Times New Roman，小四号，首行缩进 2 字距，1.5 倍行距。 |
| 52 | 102 | citation sample | 两位作者：中文用“和”，英文用“and”。 |
| 53 | 103 | citation sample | 三位及以上作者：仅写第一作者，其后加“等”或“et al.”。 |
| 54 | 105 | citation sample | 圆括号中文状态；括号内标点英文状态；逗号和分号后有一个半角空格；中文文献在前，外文在后；各组内按年份递增。 |
| 56 | 110 | `表1 表名` | 表格与正文上下各空一行；三线表；表从 1 起全文连续编号；表号后空 1 字距写表题；表题五号居中。 |
| 57 | 120 | `XXXX` | 表内内容：中文宋体，西文 Times New Roman，五号，居中，单倍行距。 |
| 58 | 141 | `注：文本文本文本文本` | 表注一般与表起始位置相同；引用文献时在表下列资料来源；表注小五号，1.5 倍行距。 |
| 62 | 144 | `1.1.1 标题` | 三级标题：楷体小四号，左对齐，1.5 倍行距；题序与标题间空 1 个字距。 |
| 64 | 146 | formula sample | 公式居中，全文连续编号，编号在右边行末，不加虚线；编号五号；含公式段落可设最小值 20 磅。 |
| 66 | 147 | footnote sample | 页下脚注全文连续编号；中文宋体，西文 Times New Roman，小五号，两端对齐，单倍行距。 |
| 67 | 149 | `图1 图名` | 插图与正文上下各空一行；图从 1 起全文连续编号；图号后空 1 字距写图题；图题五号居中；图与图题不可拆分到两页。 |
| 68 | 150 | `注：文本文本` | 图注一般与图起始位置相同；引用图时在图下列资料来源；图注小五号，1.5 倍行距。 |
| 71 | 153 | `1.1.1.1 标题` | 四级标题：楷体小四号，左对齐，1.5 倍行距；题序与标题间空 1 个字距。 |
| 72 | 160 | `(a) 分图名` | 多个分图时，各分图依次为 `(a) (b) (c)`；各分图有分图名；主图名位于全部分图下方正中。 |
| 78 | 218 | `编号` | 续表需重复表编号，且每页重复表头。 |

## References and back matter

| Comment ID | Anchor paragraph | Anchor text | Rule |
| --- | --- | --- | --- |
| 121 | 293 | `参 考 文 献` | 标题字间空 1 个字距，黑体四号，居中；条目中文宋体、英文 Times New Roman，小四号；不编序号；换行悬挂缩进 2 字距；中文在前、外文在后。 |
| 124 | 294 | reference sample | 标点用英文半角；逗号、句号、冒号、分号后有一个半角空格。 |
| 125 | 294 | reference sample | 三人及以内列全部作者；三人以上列前三人后加“等”或“et al.”。 |
| 126 | 294 | reference sample | 中文姓名为两个字时中间不留空。 |
| 127 | 294 | reference sample | 期刊无卷号时用“出版年(期号)”；无期号时用“出版年，卷号”。 |
| 128 | 294 | reference sample | 页码范围使用短线 `-`。 |
| 133 | 303 | white paper sample | 括号之间加一个半角空格。 |
| 134 | 304 | URL sample | 网址中的标点前后不加空格。 |
| 135 | 305 | book sample | 图书第 1 版不标注版次；后续版次中英文按规范写。 |
| 140 | 311 | `附录A 标题标题` | 附录序号与标题间空 1 个字距；附录中的图、表、公式单独编号，如图 A1、表 A1、式 A1。 |
| 143 | 385 | `致 谢` | 标题“致谢”空 4 个字距，黑体四号，居中，1.5 倍行距。 |
| 144 | 386 | acknowledgement sample | 致谢正文：中文宋体，西文 Times New Roman，小四号，两端对齐，首行缩进 2 字距，1.5 倍行距。 |

## Use of this reference

- Use it as the default constraint layer for the current South China Agricultural University template.
- If a new template version appears, re-extract comments with `scripts/extract_docx_comments.py` and compare before bulk formatting.
- When a rule here conflicts with the school's latest written notice, flag the conflict and let the user decide whether to follow the newer notice or the inherited template.
