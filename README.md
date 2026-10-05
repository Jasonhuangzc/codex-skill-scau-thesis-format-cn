# 华农本科毕业论文 Skill

`scau-thesis-format-cn` 用于华南农业大学（SCAU）本科论文（设计）的内容检查、模板回灌与 Word/PDF 格式复核。可在 Codex 及支持本地 skill 的工具中使用。

## 主要能力

- **内容检查**：核对摘要、关键词、模块和模板占位，报告规则出处、位置与待确认项。
- **稳定回灌**：填入封面、摘要、章节、图表和文献；局部更新保留其他章节，默认保护既有图表。
- **格式复核**：检查字体、目录、图题表题和分页；定向修复，保留工作副本与失败恢复路径。

## 安装

```bash
git clone https://github.com/Jasonhuangzc/codex-skill-scau-thesis-format-cn.git
cd codex-skill-scau-thesis-format-cn
python -m pip install -r requirements.txt
```

将 `scau-thesis-format-cn/` 放入本地 skill 目录；Codex 通常为 `~/.codex/skills/scau-thesis-format-cn`。

默认依据为华农 **2024 版官方三文件**，文件名与 SHA256 见 [manifest](scau-thesis-format-cn/assets/official-2024/manifest.json)。仓库不附带原文件或衍生模板；装版前在 Windows + Microsoft Word 环境导入：

```bash
python scau-thesis-format-cn/scripts/import_official_2024_assets.py --source-dir "C:/path/to/official-files"
```

导入会校验整套文件并生成模板和预览。仅核验可加 `--verify-only`，支持任意平台。

## 使用

在工具中直接说明任务，例如：

```text
使用 $scau-thesis-format-cn，检查这份论文的内容与格式，修复明确问题并输出工作版和检查报告。
```

对已有 DOCX 做只读内容预检：

```bash
python scau-thesis-format-cn/scripts/inspect_thesis_content.py working.docx --output content-audit.json
```

仅更新某一章（Markdown 以 `# 第3章 结果与分析` 等编号标题开头）：

```bash
python scau-thesis-format-cn/scripts/run_scau_project_pipeline.py --project-root . --docx working.docx --chapter-file chapter3.md --output updated.docx --skip-frontmatter --skip-figures --skip-tables --skip-references
```

目标章含未重建图表、公式、域或修订时需先处理，复杂对象采用 Word 定向编辑。

更多说明：[完整流程](scau-thesis-format-cn/SKILL.md) · [内容检查](scau-thesis-format-cn/references/content-audit.md) · [局部回灌与恢复](scau-thesis-format-cn/references/project-pipeline.md)。

## 验证与边界

```bash
python -m unittest discover -s tests -v
```

当前 64 项合成文档/CLI 回归及 GitHub CI 已通过。真实模板需另跑 Windows Word 冒烟测试并检查导出页面。

内容预检仅需 Python 标准库；Word COM、字体检查、模板转存与目录刷新需要 Windows + Word。更新章节后需刷新目录；Windows runner 可加 `--finalize-contents`。

原文件缺失时报告“摘录依据，原文未复验”。中英文语义、数据与文献真实性需源材料核对；文本预检不代表全文合规，不补写研究结果。

[MIT License](LICENSE)
