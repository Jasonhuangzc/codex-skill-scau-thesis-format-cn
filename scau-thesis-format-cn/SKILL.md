---
name: scau-thesis-format-cn
description: 按华南农业大学（SCAU/华农）本科毕业论文官方模板回灌用户内容，依据指定规范检查内容完整性与格式，并核对真实 Word/PDF 页面；用于装版、局部更新、终稿审查及修复，不代写研究结论。
---

# SCAU Thesis Format CN

从真实模板或已有工作稿出发，执行“确认规则与输入 → 回灌选定内容 → 内容核对 → Word 结构和渲染检查 → 定向修复 → 复查”。保留学术含义、数据和引用关系，不为满足篇幅或格式要求编造内容。

## 先确认文件与任务范围

- 区分模板初始化、局部回灌、全文装版、仅审查、审查并修复、清洁提交版。局部更新只处理选定章节或模块。
- 默认标准包为 `assets/official-2024/manifest.json` 所列三份 2024 文件；工作模板为 `assets/template/scau-undergrad-thesis-template.docx`，由附件6转存。公开仓库不附带这些原文件，须本地导入或传入已核实来源的模板。
- 官方原文件在场时核对 manifest 的 SHA256，并记录适用版本。缺原文件时可按仓库批注摘录做预检，但注明“原文未复验”；缺模板只阻塞装版，不阻塞已有 DOCX 的只读内容预检。
- 用户指定其他正式版本或院系文件时，先阅读并定位差异，再以用户要求的适用文件建立规则映射；不混用版本，不直接套用旧锚点或把新文件当作已验证的 2024 包。
- 找到源稿、元数据、图片、表格和文献，建立“来源 → 目标模块”清单。通用文件名和显式路径均可，禁止依赖某个论文项目的目录名。
- 原稿和模板保留；输出到工作副本。封面身份信息、摘要、参考文献或数据缺失时列出缺项，保留已有内容，不用占位文本覆盖真实内容。

## 按需读取

| 任务 | 必读参考 |
| --- | --- |
| 内容核对、文件依据、验收范围 | [content-audit.md](references/content-audit.md)、[format-rules.md](references/format-rules.md) |
| 模板或教师可见约束 | [scau-template-comments.md](references/scau-template-comments.md)、[template-comment-rules.md](references/template-comment-rules.md) |
| 封面与摘要 | [scau-frontmatter-map.md](references/scau-frontmatter-map.md) |
| 内容装版路线 | [workflow.md](references/workflow.md) |
| 多步骤或局部回灌 | [project-pipeline.md](references/project-pipeline.md) |
| 图、表、文献载荷 | [block-manifest.md](references/block-manifest.md)、[table-manifest.md](references/table-manifest.md)、[reference-import.md](references/reference-import.md) 中对应项 |
| Windows 大稿和定向修复 | [windows-word-com.md](references/windows-word-com.md)、[word-com-mode.md](references/word-com-mode.md) |

## 检查内容时遵循文件证据

每个发现包含 `规则来源（文件/页码或批注ID） → 目标位置 → 观察值 → 要求 → 状态 → 修复或待确认动作`。

- 显式要求、模板版式值、工作流建议和人工判断分开记录。农科四章结构是参考结构，不能强行改掉其他学科的合法章节组织。
- 按附件6批注29核对中文摘要 300–600 字、引用候选；按批注30核对中文关键词 3–5 个、全角分号及末尾无标点。记录计数口径，边界值与混合文字需确认，不能自行补写摘要。
- 批注34、35提供英文摘要与关键词格式；没有英文摘要固定字数的摘录依据，不添加字数阈值。中英文题目、摘要、关键词的含义及数据一致性需要逐项比对。
- 检查模板占位残留、空模块、章节和图表编号、文内图表指向及引用与文献的双向对应。正则只能提供引用候选，不能独自认定来源真实性或引用完整。
- 摘要是否覆盖目的、方法、结果、结论，研究数据是否一致、来源是否支持主张，由源稿与证据判断。无法从现有材料证实的内容列入 `manual_confirm`。
- 附录和缩略词表按实际需要；工作阶段可保留待填模块，终稿中的必需内容缺失、样例文献、占位致谢必须列为未完成。

先运行跨平台只读预检：

```bash
python scripts/inspect_thesis_content.py working.docx --output content-audit.json
```

它核对 DOCX 文本与结构信号，并输出来源、位置和未验证范围。不能替代学术核查、Word 字体检查或页面渲染。旧 `inspect_word_report.py` 的统计和 `detected` 状态只说明检测到了模块。

## 合理高效回灌

1. 回灌前核对目标标题、范围、所需样式 donor 和源载荷。Markdown 一次提供一个带编号章；不支持或不明确的结构先转换为受支持的内容，不把复杂对象悄悄丢掉。
2. 根据范围选择脚本：
   - 封面/摘要：`scripts/fill_scau_frontmatter.py`。使用标签定位，校验 cover runs；拆分字段或字段后残留非空文本时停止，改用 Word 定向编辑。摘要换行保留段落，省略的摘要/关键词保持原值。
   - 单章：`scripts/insert_markdown_chapter.py`。替换目标章内容，保留其他章节；标题改名按章节编号定位。
   - 图：`scripts/insert_figure_blocks.py`；Windows 大稿可用 `scripts/insert_figure_blocks_com.py`。
   - 表：`scripts/insert_table_blocks.py`；图表另有 manifest 时避免与 Markdown 表重复插入。
   - 文献：`scripts/insert_reference_batch.py`；源条目先核验，按适用文件排序，不自动补全未知著录信息。
   - 多步编排：`scripts/run_scau_project_pipeline.py`。局部范围使用相应 skip 参数，读取 project-pipeline.md，不要求被跳过步骤的输入。
3. 首轮从模板初始化时，只在样例区与模板吻合的前提下清理样例；已有工作稿默认保留其他章节。重复运行需核对章标题、图表顺序和旧表残留。目标章有既有图片或无 Markdown 载荷的表时默认停止；只有核对完整重建清单后才使用 `--replace-media` / `--replace-tables`（runner 同名参数）。图表单独插入脚本不保证幂等，重跑前核对已插对象，避免重复。
4. 检查保存后的 Word：目标段落/表格/图片数量、次序、源文本与目标文本对应、非目标模块是否保持。编号、图题、表题、文献排序等允许变动须单列，不用全文字数相等代替内容核对。
5. 正文或目录相关条目改变后，最终刷新目录。runner 的 `--finalize-contents` 为 Windows Word 的显式目录收尾，不自动对全文做宽泛样式修复。
6. 大稿保留阶段副本。每阶段只做一组相关操作，用一个 Word 会话完成该组，保存一次；失败从最新有效阶段恢复。不要每改一段就启动 Word、更新全目录或重导出整篇。

回灌含修订、内容控件、域、公式、嵌套表或其他复杂对象时，先确认脚本覆盖范围；不能保证保留的区间采用 Word 定向编辑，不将只处理普通段落的脚本强套到复杂文档。

## 格式检查与修复

- `scripts/inspect_word_report.py`：Windows Word 基础统计与结构信号。
- `scripts/inspect_word_format_signatures.py`：字符格式、标签/正文边界、目录和尾部样式；区分代表段落抽样与扫描数量。
- `scripts/export_word_to_pdf.py` 或 `scripts/export_word_preview.ps1`：Word 导出。
- `scripts/render_pdf_pages.py`：页面图片；必要时用 `scripts/inspect_figure_layout.py` 检查图块。

字体、字号、标签加粗和缩进以 Word 结构为证；分页、图与图题同页、表题/表注位置、续表、声明页无页码、目录页码以真实页面为证。两字距缩进可用字符单位或结合字体验证等效磅值；摘要独立页可通过段前分页或分页型分节实现。模板未明写的值不得宣称为学校硬性规定。

修复选择正确层：落错位置回到对应回灌步骤；目录、局部字体、分页问题用 `scripts/batch_word_ops.py` 做目标操作。常见操作包括 `replace_text`、`ensure_page_break_before`、`normalize_tail_section_fonts`、`cleanup_contents_entries`、`normalize_contents_fonts` 和 `finalize_contents`，细节见 Word COM references。

批量替换前关闭修订记录；已有修订须保留或按用户要求处理。标题变更需完整目录更新；仅页码改变可更新页码；清理目录里的“参考文献”“致谢”空格在域更新后完成。不要把模板例值归一化到所有布局表或非目标区。

修复后重查受影响内容和版式；最终另做整体复查。若相同问题连续两轮没有改善，保留有效副本、报告阻碍和人工处理位置，不无限循环或重建全文。

## 验证与交付

维护代码后先运行公开仓库可执行的回归：

```bash
python -m unittest discover -s tests -v
```

覆盖内容定位、章节替换、表题顺序、重复回灌和局部范围保留。使用合成文档，不等同官方模板验收。具备 Windows Word 和导入模板时另运行 `scripts/smoke_test_scau_skill.py`，验证官方模板50条批注、真实 frontmatter、章回灌和小修字体签名。未执行的环境检查明确列出。

交付工作 Word、实际已生成的预览 PDF、内容与格式报告；列明文件版本、依据、已检查范围、未完成内容和 `manual_confirm` 项。`confirmed` 仅用于有证据的具体检测项，不能据此宣称全文符合规定或可提交。

用户需要提交版时，在工作稿验收后另存清洁副本：`scripts/finalize_submission_copy.ps1` 或 `scripts/strip_docx_comments.py`。检查批注/修订处理结果，复核目录与尾部格式，并再次确认清洁副本页面。未完成内容或缺渲染证据时如实说明。
