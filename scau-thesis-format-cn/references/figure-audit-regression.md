# Figure Audit Regression Checklist

Use this checklist after changing the figure-audit workflow or after a large figure insertion pass.

## Minimal regression checklist

1. `图全部插入成功`
   - All target figures such as `图3-1` to `图3-8` appear in the rendered PDF and have page mappings.
2. `图题在图下`
   - Each figure caption is on the same page as its figure body and follows the body.
3. `图注在图题下`
   - If a figure note or source note exists, it follows the caption and stays on the expected page.
4. `图块不跨页`
   - Figure body, caption, and note do not split across pages.
5. `标题起页合理`
   - The next section or subsection heading is not awkwardly forced onto the same crowded page as the figure block unless clearly intentional.
6. `复查后问题已收敛`
   - Every previously reported high-risk figure issue has a corresponding recheck result.

## Structured outputs to review

- `figure_page_map`
- `caption_order_status`
- `note_order_status`
- `split_risk`
- `next_heading_risk`
- `page_images`
- `still_manual_confirm`
