# Figure Block Manifest

Use this JSON format with either:

- `scripts/insert_figure_blocks.py`
- `scripts/insert_figure_blocks_com.py`

The same figure manifest works for both backends.

## Minimal example

```json
[
  {
    "anchor_regex": "如图3-1所示",
    "position": "after",
    "image": "figures/figure3-1.png",
    "caption": "图3-1 不同处理组样品荧光图像",
    "page_break_before": false
  }
]
```

## Full example

```json
[
  {
    "anchor_regex": "如图3-1所示",
    "occurrence": 1,
    "position": "after",
    "image": "figures/figure3-1.png",
    "caption": "图3-1 不同处理组样品荧光图像",
    "note": "注：CK 为对照组，L、M、H 分别表示不同处理组。",
    "width_cm": 12.5
  },
  {
    "anchor_regex": "如图3-3所示",
    "position": "after",
    "layout": "grid",
    "grid": {
      "rows": 2,
      "cols": 2,
      "cell_width_px": 1500,
      "cell_height_px": 1100,
      "title_height_px": 120
    },
    "panels": [
      { "image": "figures/figure3-3a.png", "title": "N-C" },
      { "image": "figures/figure3-3b.png", "title": "N-L" },
      { "image": "figures/figure3-3c.png", "title": "N-M" },
      { "image": "figures/figure3-3d.png", "title": "N-H" }
    ],
    "caption": "图3-3 不同处理组组织切片结果",
    "note": "注：脚本会自动生成 (a) 至 (d) 的分图标记并合成为最终插图。",
    "composite_output": "figures/figure3-3-composite.png",
    "width_cm": 11.8
  }
]
```

## Field rules

- `anchor_regex`
  Required. Python regular expression used to locate the paragraph that controls insertion.
- `occurrence`
  Optional. Use when the regex matches more than one paragraph. Default is `1`.
- `position`
  Optional. Either `after` or `before`. Default is `after`.
- `image`
  Optional. Single image path.
- `images`
  Optional. Multiple image paths. The current script inserts them as vertically stacked centered image paragraphs.
- `layout`
  Optional. Set to `grid` or `panel-grid` to auto-compose a multi-panel figure before insertion.
- `grid`
  Optional. Grid layout settings when `layout` is `grid`, such as `rows`, `cols`, `cell_width_px`, `cell_height_px`, `title_height_px`, `padding_px`, `gutter_x_px`, and `gutter_y_px`.
- `panels`
  Optional. Array of panel objects for grid composition. Each panel accepts `image`, optional `label`, and optional `title`.
- `composite_output`
  Optional. Where to save the generated composite image. If omitted, the script writes a temporary PNG under the system temp directory.
- `caption`
  Required. Main figure caption such as `图3-3 不同处理组组织切片结果`.
- `note`
  Optional. Figure note or source note placed below the caption.
- `width_cm`
  Optional. Render width in centimeters. Default is `12.0`.
- `page_break_before`
  Optional. When `true`, insert a hard page break before the figure block. Use this for tall figures with long notes when you already know the block should start on a fresh page.
- `_meta`
  Optional. Generator metadata. The insertion script ignores this field, but it can record whether `caption`, `anchor_regex`, or `width_cm` came from overrides or automatic inference.

## Recommendation

- For four-panel comparison figures, prefer `layout: "grid"` with `panels`, so the script can generate a `2 x 2` panel layout and put each subfigure label directly under its own panel.
- For tall figures with long notes, set `page_break_before: true` when you already know the whole block should start on a fresh page.
