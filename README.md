# SCAU / 华农本科毕业论文格式 Skill

`scau-thesis-format-cn` 是一个面向 **华南农业大学（SCAU / 华农）本科毕业论文（设计）** 的本地 skill。  
它适合接入 **Codex、Claude Code、Trae、Antigravity** 以及其他支持本地 skill / prompt / agent workflow 的环境，用于：

- 按华农官方 Word 模板装版论文
- 回灌摘要、章节、图表、参考文献
- 审查真实 Word / PDF 页面版式
- 修复目录、字体、图题、表题、参考文献、分页等终稿格式问题

关键词：`华农`、`SCAU`、`华南农业大学`、`本科毕业论文`、`毕业论文格式`、`论文模板`、`Word 模板`、`格式审查`、`格式检查`、`参考文献格式`、`图表格式`、`thesis template`、`Word formatting`、`final thesis audit`

## 这是什么

这不是一个普通的“规则清单”或者“文本扫错器”。  
它是一套完整的 **模板装版 + 真实版式审查 + 定向修复 + 复查收敛** 工作流。

核心能力分成两块：

### 1. 模板装版能力

- 从华农官方 Word 模板出发，而不是自己重建排版
- 写入封面、摘要、Abstract、关键词
- 将 Markdown 章节整章回灌到模板中
- 插入图、表、参考文献
- 保持标题层级、段落格式、正文字体、目录结构与模板一致
- 参考文献按“中文在前、外文在后；中文按第一著者姓氏汉语拼音字母顺序，西文和俄文按第一著者姓氏字母顺序”收口
- 最终目录按“更新域 -> 清理 `参考文献` / `致谢` 空格 -> 目录中文宋体、英文数字和 `.` 为 Times New Roman”收尾
- 对大文档支持 Word COM 单会话批处理，减少反复改写造成的版式漂移

### 2. 强格式审查能力

- 审查真实 Word 结构，而不只看纯文本
- 导出 PDF 后审查真实页面，而不只看文本抽取结果
- 检查图题位置、图注位置、图块跨页、续表、目录页码、分页稳定性
- 检查 `摘要 / Abstract / Key words / 关键词 / 参考文献 / 致谢 / 目录` 等高风险位置
- 检查标题样式、字号、加粗边界、目录特殊条目、参考文献与正文引用格式
- 检查目录字体、`参考文献 / 致谢` 目录项空格、正文 `1.5` 倍行距、重复标点如 `。。` / `，，`
- 小修文字后复查字体和段落样式，避免“改对了内容，改坏了格式”

## 适用官方文件

这个 skill **严格以华农 2024 版官方三文件为准**。

唯一标准来源是：

1. `附件1-5.华南农业大学本科毕业论文（设计）撰写规范（封面模板、原创性声明及使用授权声明、正文结构参考样式、参考文献著录规则、评分参考标准）.pdf`
2. `附件6.华南农业大学本科毕业论文（设计）格式模板.doc`
3. `关于印发《华南农业大学本科毕业论文（设计）撰写规范》（2024年修订）的通知.doc`

这三份文件的文件名和 SHA256 清单在：

- [scau-thesis-format-cn/assets/official-2024/manifest.json](./scau-thesis-format-cn/assets/official-2024/manifest.json)

## 快速开始

### 1. 克隆仓库

```powershell
git clone https://github.com/Jasonhuangzc/codex-skill-scau-thesis-format-cn.git
cd codex-skill-scau-thesis-format-cn
```

### 2. 导入华农 2024 官方三文件

先准备好包含那三份官方文件的目录，然后运行：

```powershell
python .\scau-thesis-format-cn\scripts\import_official_2024_assets.py --source-dir "C:\path\to\华农2024官方三文件目录"
```

这个脚本会自动完成：

- 校验三文件是否齐全
- 校验 SHA256
- 拷贝到 `assets/official-2024/`
- 从 `附件6.doc` 生成：
  - `assets/template/scau-undergrad-thesis-template.doc`
  - `assets/template/scau-undergrad-thesis-template.docx`
  - `assets/template/scau-undergrad-thesis-template-preview.pdf`
- 抽取模板批注并确认评论数

### 3. 放到你的本地 skill 目录

实际 skill 目录是：

- `scau-thesis-format-cn/`

把它放到你所使用的工具支持的本地 skill / prompt / agent 目录中即可。

## 在不同工具里怎么调用

下面这些写法的目标是一致的：**显式告诉代理使用 `scau-thesis-format-cn`，并说明你是要装版、审查还是修复。**

### Codex

如果你在 Codex 中使用本地 skill，通常放到：

- `~/.codex/skills/scau-thesis-format-cn`

调用示例：

```text
使用 $scau-thesis-format-cn，把这篇华农本科毕业论文内容装入官方模板，并输出工作版 Word。
```

```text
使用 $scau-thesis-format-cn，审查这份论文终稿的 Word 和 PDF 格式，重点检查目录、参考文献、图题图注和分页。
```

### Claude Code

如果你的 Claude Code 工作流支持本地 skills / reusable prompts / workflow folders，把 `scau-thesis-format-cn/` 放到对应目录，然后在任务里显式点名它。

调用示例：

```text
Use the local skill "scau-thesis-format-cn" to assemble this SCAU undergraduate thesis into the official Word template and then audit the rendered formatting.
```

```text
Use "scau-thesis-format-cn" only for final-format review. Do not rewrite academic content. Focus on references, TOC, captions, and page layout.
```

### Trae

如果你在 Trae 中维护项目级 agent workflows / prompt packs / local skills，可以把本目录作为一个本地工作流能力包导入。

调用示例：

```text
调用 scau-thesis-format-cn，对这份华农本科毕业论文做模板装版，并生成可继续修订的工作版。
```

```text
调用 scau-thesis-format-cn，对终稿做格式复查，按“问题 -> 修复动作 -> 复查结果”输出报告。
```

### Antigravity

如果你在 Antigravity 中使用本地 agent skills 或项目内 workflow 目录，导入 `scau-thesis-format-cn/` 后可直接按 skill 名调用。

调用示例：

```text
Use scau-thesis-format-cn for final thesis formatting only. Start from the official SCAU template, then inspect Word/PDF layout and repair format issues.
```

### 其他支持本地 skill 的环境

通用原则很简单：

1. 把 `scau-thesis-format-cn/` 放到该环境可读取的本地 skill / workflow 目录
2. 在任务中显式点名 `scau-thesis-format-cn`
3. 说明当前目标是：
   - 装版
   - 终稿格式审查
   - 审查并修复

通用调用模板：

```text
使用 scau-thesis-format-cn，按华农 2024 官方模板处理这份本科毕业论文。
先装版 / 审查 / 修复，再输出结果。
不要改学术内容，只处理格式、模板和版式问题。
```

## 典型使用场景

### 场景 1：把论文内容装进华农模板

```text
使用 scau-thesis-format-cn，把我提供的封面信息、摘要、正文、图表和参考文献装入华农官方模板，生成工作版 Word。
```

### 场景 2：终稿格式审查

```text
使用 scau-thesis-format-cn，检查这份论文终稿的格式是否符合华农 2024 规范。
重点看目录、参考文献、图题图注、分页、摘要、Abstract、关键词和标题层级。
```

### 场景 3：检查并修复

```text
使用 scau-thesis-format-cn，先输出格式审查报告，再修复明确的格式问题，最后做一轮复查。
不要改学术内容。
```

## 仓库结构

```text
codex-skill-scau-thesis-format-cn/
├── README.md
├── LICENSE
├── .gitignore
└── scau-thesis-format-cn/
    ├── SKILL.md
    ├── agents/openai.yaml
    ├── assets/
    │   ├── official-2024/
    │   └── template/
    ├── references/
    └── scripts/
```

## 核心脚本

### 模板与内容装版

- `scripts/import_official_2024_assets.py`
- `scripts/fill_scau_frontmatter.py`
- `scripts/insert_markdown_chapter.py`
- `scripts/insert_figure_blocks.py`
- `scripts/insert_figure_blocks_com.py`
- `scripts/insert_table_blocks.py`
- `scripts/insert_reference_batch.py`
- `scripts/run_scau_project_pipeline.py`

### 格式审查与修复

- `scripts/inspect_word_report.py`
- `scripts/inspect_word_format_signatures.py`
- `scripts/inspect_figure_layout.py`
- `scripts/export_word_to_pdf.py`
- `scripts/render_pdf_pages.py`
- `scripts/batch_word_ops.py`

## 可信性和维护方式

这套 skill 当前已经完成这些验证：

- 以华农 2024 官方三文件为唯一标准来源
- 模板批注抽取数为 `50`
- 整章回灌测试通过
- 小修替换后字体签名检查通过
- 项目私货残留扫描通过

如果华农模板后续更新，建议按这个顺序维护：

1. 替换本地官方三文件
2. 重新运行 `import_official_2024_assets.py`
3. 检查批注数量和锚点映射
4. 运行 `smoke_test_scau_skill.py`

## 适用范围

如果你要找的是这些能力：

- 华农论文格式装版
- 华农毕业论文模板审查
- SCAU Word 模板自动填充
- 华南农业大学本科毕业论文终稿格式复查
- 参考文献、图表、目录、分页的最终检查

这个仓库就是对应的 skill。

## 许可证

本仓库中的代码、脚本、规则说明和通用文档按 MIT 许可开放，见：

- [LICENSE](./LICENSE)
