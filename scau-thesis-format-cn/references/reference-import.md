# Reference Batch Import

Use `scripts/insert_reference_batch.py` when the thesis already has a verified bibliography draft and the remaining work is to place it into the Word template with correct paragraph formatting.

## Supported source shape

The source file can be Markdown or plain text.

Preferred Markdown shape:

```md
# 最终参考文献著录初稿

## 中文文献

张三. 某环境因子对鱼类肠道健康影响的研究[D]. 南京: 某大学, 2021.

李四, 王五, 赵六, 等. 新兴污染物生态毒理研究进展[J]. 环境科学, 2025, 46(3): 1868-1884.

## 英文文献

Smith J, Chen R, Brown T, et al. Effects of emerging contaminants on fish intestinal health[J]. Journal of Hazardous Materials, 2024, 470: 134157.
```

The script reads:

- all entries under `## 中文文献`
- all entries under `## 英文文献`

and inserts them in that order.

Within the Chinese block, the script sorts entries by the first author's surname in Hanyu Pinyin order.

Within the foreign-language block, the script sorts entries by the first author's surname in alphabetical order. The implementation uses one foreign-language block; verify language subgroup requirements against the applicable PDF before using a mixed-language bibliography.

## Commands

Insert and replace the existing reference sample entries:

```powershell
python scripts/insert_reference_batch.py `
  --docx "work\thesis_working.docx" `
  --references-file "drafts\references_final.md" `
  --output "work\thesis_working_refs.docx"
```

Reformat an already inserted reference section and re-apply hanging indent:

```powershell
python scripts/insert_reference_batch.py `
  --docx "work\thesis_working_refs.docx" `
  --reformat-only `
  --output "work\thesis_working_refs_fixed.docx"
```

## What the script does

- finds the `参考文献` heading in the template
- replaces the existing sample entries unless `--reformat-only` is used
- keeps Chinese entries first and foreign-language entries second
- sorts Chinese entries by the first author's surname in Hanyu Pinyin order
- sorts the foreign-language block by normalized author keys; this implementation does not certify cross-language collation rules
- clones the template reference paragraph style from the sample bibliography
- reapplies one-and-a-half line spacing and zero paragraph spacing
- reapplies hanging indent as two characters through OOXML character-based indentation

## Recommended source shape

- Keep one final bibliography draft file, for example `drafts/references_final.md`.
- The file should already be grouped by language.
- The insertion script re-sorts each implemented language group before writing. Confirm the applicable school rule and special surnames first; the algorithm is not evidence of official language subgroup order.

## Collation evidence

Chinese insertion requires `pypinyin`; raw Unicode and unspecified Chinese locale sorting do not establish Hanyu Pinyin surname order. Audit reports downgrade unavailable collation to `manual_confirm`. Compound/polyphonic surnames and language subgroup ordering need verification against the applicable file. Sorting never establishes whether an entry is cited or whether its source is real.

Reference heading discovery excludes real TOC fields/styles and rejects duplicate body anchors. Review data, fields, section boundaries and unsupported objects in the replacement range require targeted Word editing; reformat-only also refuses hyperlinks that its ordinary-run writer cannot safely preserve. Output paths must differ from source files.
