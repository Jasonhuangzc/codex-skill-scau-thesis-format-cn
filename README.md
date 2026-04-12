# SCAU / 华农本科毕业论文格式 Skill

`codex-skill-scau-thesis-format-cn` 是一个面向 **华南农业大学（SCAU / 华农）本科毕业论文（设计）** 的 Codex 本地 skill。  
它不是普通的“文本规则检查器”，而是一个围绕 **Word 模板装版、PDF 渲染审查、定向格式修复、终稿收敛复查** 设计的完整流程工具。

关键词：`华农`、`SCAU`、`华南农业大学`、`本科毕业论文`、`毕业论文格式`、`论文模板`、`Word 模板`、`格式审查`、`格式检查`、`参考文献格式`、`图表格式`、`Codex skill`

## 这是什么

这个仓库公开的是一个可复用的本地 skill：

- 用华农官方模板装版论文内容
- 检查真实 Word / PDF 页面版式
- 修复目录、字体、标题、图表、参考文献等格式问题
- 循环复查，直到接近提交版

仓库里的实际 skill 目录是：

- `scau-thesis-format-cn/`

## 它和普通论文格式检查器的区别

这套 skill 的核心不是“扫文本”，而是三层一起工作：

1. **模板层**
   - 从华农官方 Word 模板出发，而不是自己重建一份版式。
2. **结构层**
   - 检查 Word 中的字体、字号、加粗边界、目录特殊条目、参考文献和致谢样式。
3. **渲染层**
   - 导出 PDF，再看真实页面，处理图题位置、图注跨页、目录页码、续表、分页等问题。

这意味着它能处理很多文本提取做不到的终稿问题，比如：

- 图题在不在图下
- 图注有没有被挤到下一页
- 目录页码是否真的对齐
- `参考文献` 和 `致谢` 在目录里有没有错误带入标题字间距
- 小修文字后正文是不是还保持 `宋体 + Times New Roman + 小四`

## 唯一标准来源：华农 2024 官方三文件

这个 skill **严格以华农 2024 版官方三文件为准**，不是按旧版模板、网上二手模板或个人习惯规则写的。

唯一标准来源是：

1. `附件1-5.华南农业大学本科毕业论文（设计）撰写规范（封面模板、原创性声明及使用授权声明、正文结构参考样式、参考文献著录规则、评分参考标准）.pdf`
2. `附件6.华南农业大学本科毕业论文（设计）格式模板.doc`
3. `关于印发《华南农业大学本科毕业论文（设计）撰写规范》（2024年修订）的通知.doc`

这三份文件的文件名和 SHA256 清单在：

- [scau-thesis-format-cn/assets/official-2024/manifest.json](./scau-thesis-format-cn/assets/official-2024/manifest.json)

## 为什么公开仓库里不直接放官方原件

这个仓库默认 **不直接分发** 华农官方 PDF / Word 原件，也不默认提交由其派生出的模板 `.docx` 和预览 PDF。

原因很简单：

- 这套 skill 必须严格以官方文件为准
- 但公开仓库不适合默认假定这些学校文件可以自由再分发

因此这里采用更稳的模式：

1. 公开代码、规则、批注映射和审查脚本
2. 由使用者在本地导入官方三文件
3. 导入脚本自动校验文件名和 SHA256
4. 再本地生成工作模板和预览 PDF

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

### 3. 安装 skill

把 `scau-thesis-format-cn/` 放到你的 `$CODEX_HOME/skills/` 下。

Windows 常见路径类似：

- `C:\Users\<你的用户名>\.codex\skills\scau-thesis-format-cn`

### 4. 跑一遍维护冒烟测试

```powershell
python .\scau-thesis-format-cn\scripts\smoke_test_scau_skill.py --banned-token 你旧项目中的专有词
```

这个测试会验证：

- skill 中没有残留某一篇论文项目的私货
- 官方模板批注数仍正确
- 前置部分可写入
- 一整章 Markdown 可回灌
- 小修替换后字体、字号和目录特殊条目不漂

## 这个 skill 能做什么

### 装版

- 写入封面、摘要、Abstract、关键词
- 把 Markdown 章节回灌进 Word 模板
- 插入图、表、参考文献
- 维护标题层级和模板样式

### 审查

- 检查图题、图注、续表、目录页码、分页
- 检查标题样式、正文样式、目录特殊条目
- 检查参考文献与正文引用格式
- 检查 `Abstract:` / `Key words:` 等高风险边界

### 修复

- 小修文字且尽量保持原字体
- 清理目录中的 `参考文献` / `致谢` 空格污染
- 恢复 `参考文献` / `致谢` 的标题与正文样式
- 在 Word COM 会话里集中修复终稿问题，减少大文档反复重写

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

## 可信性和维护策略

这套 skill 当前已经做过这些验证：

- 以华农 2024 官方三文件为唯一标准来源
- 模板批注抽取数为 `50`
- 整章回灌测试通过
- 小修替换后字体签名检查通过
- 项目私货残留扫描通过

如果后面华农模板更新，建议按这个顺序维护：

1. 替换本地官方三文件
2. 重新运行 `import_official_2024_assets.py`
3. 检查批注数量和锚点映射
4. 运行 `smoke_test_scau_skill.py`

## 适用范围

这个 skill 是为 **SCAU / 华农本科毕业论文（设计）终稿格式工作流** 准备的。  
如果你需要的是：

- 华农论文格式装版
- 华农毕业论文模板审查
- SCAU Word 模板自动填充
- 本科论文 Word / PDF 版式复查

这就是对应的仓库。

## 许可证

本仓库中的代码、脚本、规则说明和通用文档按 MIT 许可开放，见：

- [LICENSE](./LICENSE)

注意：

- MIT 只覆盖仓库作者编写的代码和说明
- 不自动覆盖你本地导入的华农官方文件
- 官方文件及其派生模板仍应按原始来源和适用规则使用
