# Table Block Manifest

Use this JSON format with `scripts/insert_table_blocks.py`.

## Minimal example

```json
[
  {
    "anchor_regex": "主要指标汇总见表3-3",
    "position": "after",
    "table_file": "tables/表3-3_主要指标汇总表.md"
  }
]
```

## Continued-table example

```json
[
  {
    "anchor_regex": "见表2-2",
    "position": "after",
    "table_file": "tables/表2-2_主要实验材料、试剂与仪器.md",
    "max_body_rows_per_segment": 6,
    "continued_suffix": "（续表）"
  }
]
```

## Supported fields

- `anchor_regex`
  Required. Python regex used to locate the paragraph that controls insertion.
- `occurrence`
  Optional. Use when the regex matches more than one paragraph. Default is `1`.
- `position`
  Optional. Either `after` or `before`. Default is `after`.
- `table_file`
  Optional. Markdown file containing one table block. The script reads the first `# 表x-x ...` line as the caption, the first Markdown pipe table as the table body, and trailing `注：` or `资料来源：` lines as the note.
- `caption`
  Optional. Override caption text from `table_file`.
- `note`
  Optional. Override note text from `table_file`.
- `rows`
  Optional. Inline table rows. Use this when no `table_file` is provided.
- `max_body_rows_per_segment`
  Optional. When provided and the table body exceeds this row count, the script splits the table into multiple segments and generates continued-table captions automatically.
- `max_segment_weight`
  Optional. Heuristic row-height budget for auto splitting. Useful when rows contain long text and raw row count alone is not enough.
- `continued_suffix`
  Optional. Default is `（续表）`.

## What auto continued tables do

- repeat the same table number on each continued segment
- append `（续表）` by default to the continued caption
- repeat the table header row on each segment
- keep the table note only after the last segment

## Important limitation

This feature is structure-driven, not page-render-driven.

That means:

- the script splits tables according to `max_body_rows_per_segment` or `max_segment_weight`
- Word still decides the final page breaks during layout
- for high-risk tables with very tall cells, export a PDF preview and verify whether the split point is reasonable

## Recommendation

- Keep source tables in one dedicated folder such as `tables/`.
- For short result tables, normally no continued-table split is needed.
- For long materials/instruments tables, set `max_body_rows_per_segment` explicitly when you want deterministic continued-table output.
