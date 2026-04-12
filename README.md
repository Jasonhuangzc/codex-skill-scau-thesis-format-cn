# codex-skill-scau-thesis-format-cn

公开版 Codex 本地 skill，用于按华南农业大学 2024 本科毕业论文（设计）规范进行：

- 模板装版
- Word / PDF 真实版式审查
- 定向格式修复
- 终稿收敛复查

仓库中的实际 skill 目录是：

- `scau-thesis-format-cn/`

## 公开版边界

本仓库默认**不直接分发**华农官方 PDF / Word 原件及其派生模板文件。

原因很直接：

- 这类学校官方文件的再分发权限不适合在公开仓库里默认假定
- 但 skill 又必须严格以 2024 官方三文件为准

因此本仓库采用下面这条更稳的模式：

1. 代码、规则层、批注映射、审查与修复脚本公开
2. 官方三文件由使用者在本地导入
3. 导入脚本会校验文件名和 SHA256
4. 再本地生成工作用 `.docx` 和预览 PDF

## 官方标准包

本 skill 的唯一标准来源是华农 2024 版三文件：

1. `附件1-5.华南农业大学本科毕业论文（设计）撰写规范（封面模板、原创性声明及使用授权声明、正文结构参考样式、参考文献著录规则、评分参考标准）.pdf`
2. `附件6.华南农业大学本科毕业论文（设计）格式模板.doc`
3. `关于印发《华南农业大学本科毕业论文（设计）撰写规范》（2024年修订）的通知.doc`

导入清单和 hash 在：

- [manifest.json](./scau-thesis-format-cn/assets/official-2024/manifest.json)

## 安装

把 `scau-thesis-format-cn/` 放到你的 `$CODEX_HOME/skills/` 下。

如果你本地已经在 Windows 上使用 Codex，一般路径类似：

- `C:\Users\<你自己的用户名>\.codex\skills\scau-thesis-format-cn`

## 首次使用前导入官方文件

先准备好上面那三个官方文件所在目录，然后运行：

```powershell
python .\scau-thesis-format-cn\scripts\import_official_2024_assets.py --source-dir "C:\path\to\官方三文件所在目录"
```

这个脚本会：

- 校验三文件是否齐全
- 校验 SHA256
- 拷贝到 `assets/official-2024/`
- 从 `附件6.doc` 生成：
  - `assets/template/scau-undergrad-thesis-template.docx`
  - `assets/template/scau-undergrad-thesis-template-preview.pdf`
- 再抽取批注，确认模板评论数

## 维护冒烟测试

导入官方文件后，可以运行：

```powershell
python .\scau-thesis-format-cn\scripts\smoke_test_scau_skill.py --banned-token 你的旧项目关键词
```

这个冒烟测试会验证：

- skill 中没有残留某个具体论文项目的细节
- 官方模板批注数仍正确
- 前置部分可写入
- 一整章 Markdown 可回灌
- 小修替换后字体和目录特殊条目不漂

## 许可证

本仓库中的代码、说明和通用规则建议按 MIT 许可使用，见：

- [LICENSE](./LICENSE)

注意：

- MIT 仅覆盖本仓库中作者编写的代码与说明
- 不自动覆盖你本地导入的华农官方文件
- 导入后的官方文件及其派生模板，仍应按原始来源和适用规则使用
