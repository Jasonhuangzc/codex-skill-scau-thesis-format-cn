#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import re
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps


DEFAULT_GRID = {
    "padding_px": 80,
    "gutter_x_px": 70,
    "gutter_y_px": 90,
    "cell_width_px": 1500,
    "cell_height_px": 1100,
    "title_height_px": 120,
    "border_width_px": 3,
    "background_color": "#FFFFFF",
    "border_color": "#000000",
    "title_color": "#000000",
    "font_size_px": 56,
}

FONT_CANDIDATES = [
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/arial.ttf",
]


def sanitize_stem(text: str) -> str:
    cleaned = re.sub(r"[\\\\/:*?\"<>|]+", "_", text.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:80] or "panel_grid"


def load_font(size: int, font_path: str | None = None):
    candidates = [font_path] if font_path else []
    candidates.extend(FONT_CANDIDATES)
    for candidate in candidates:
        if not candidate:
            continue
        path = Path(candidate)
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size=size)
            except OSError:
                continue
    return ImageFont.load_default()


def auto_label(index: int) -> str:
    return f"({chr(ord('a') + index)})"


def resolve_path(raw: str, base_dir: Path) -> Path:
    path = Path(raw)
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Panel image not found: {path}")
    return path


def build_panel_title(panel: dict, index: int, auto_labels: bool) -> str:
    label = panel.get("label")
    if not label and auto_labels:
        label = auto_label(index)
    title = panel.get("title", "").strip()
    if label and title:
        return f"{label} {title}"
    return (label or title or "").strip()


def compute_grid_shape(panel_count: int, rows: int | None, cols: int | None) -> tuple[int, int]:
    if rows and cols:
        if rows * cols < panel_count:
            raise ValueError("Grid rows * cols is smaller than panel count.")
        return rows, cols
    if rows:
        return rows, math.ceil(panel_count / rows)
    if cols:
        return math.ceil(panel_count / cols), cols
    cols = math.ceil(math.sqrt(panel_count))
    rows = math.ceil(panel_count / cols)
    return rows, cols


def fit_image(image: Image.Image, max_width: int, max_height: int) -> Image.Image:
    image = ImageOps.exif_transpose(image).convert("RGB")
    ratio = min(max_width / image.width, max_height / image.height)
    new_size = (max(1, int(image.width * ratio)), max(1, int(image.height * ratio)))
    return image.resize(new_size, Image.Resampling.LANCZOS)


def draw_centered_text(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], text: str, font, fill: str) -> None:
    if not text:
        return
    left, top, right, bottom = box
    bbox = draw.multiline_textbbox((0, 0), text, font=font, align="center", spacing=8)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = left + (right - left - text_width) / 2
    y = top + (bottom - top - text_height) / 2
    draw.multiline_text((x, y), text, font=font, fill=fill, align="center", spacing=8)


def compose_from_spec(spec: dict, base_dir: Path, output_path: Path | None = None) -> Path:
    panels = spec.get("panels")
    if not isinstance(panels, list) or not panels:
        raise ValueError("Grid spec must include a non-empty panels array.")

    grid_cfg = dict(DEFAULT_GRID)
    grid_cfg.update(spec.get("grid", {}))
    rows, cols = compute_grid_shape(
        len(panels),
        spec.get("rows") or grid_cfg.get("rows"),
        spec.get("cols") or grid_cfg.get("cols"),
    )

    padding = int(grid_cfg["padding_px"])
    gutter_x = int(grid_cfg["gutter_x_px"])
    gutter_y = int(grid_cfg["gutter_y_px"])
    cell_width = int(grid_cfg["cell_width_px"])
    cell_height = int(grid_cfg["cell_height_px"])
    title_height = int(grid_cfg["title_height_px"])
    border_width = int(grid_cfg["border_width_px"])
    background_color = str(grid_cfg["background_color"])
    border_color = str(grid_cfg["border_color"])
    title_color = str(grid_cfg["title_color"])
    font_size = int(grid_cfg["font_size_px"])
    font = load_font(font_size, grid_cfg.get("font_path"))
    auto_labels = bool(spec.get("auto_labels", True))

    canvas_width = padding * 2 + cols * cell_width + max(cols - 1, 0) * gutter_x
    canvas_height = padding * 2 + rows * (cell_height + title_height) + max(rows - 1, 0) * gutter_y
    canvas = Image.new("RGB", (canvas_width, canvas_height), background_color)
    draw = ImageDraw.Draw(canvas)

    for index, panel in enumerate(panels):
        row = index // cols
        col = index % cols
        cell_left = padding + col * (cell_width + gutter_x)
        cell_top = padding + row * (cell_height + title_height + gutter_y)
        image_box = (cell_left, cell_top, cell_left + cell_width, cell_top + cell_height)
        title_box = (cell_left, cell_top + cell_height, cell_left + cell_width, cell_top + cell_height + title_height)

        image_path = resolve_path(panel["image"], base_dir)
        with Image.open(image_path) as image:
            fitted = fit_image(image, cell_width - border_width * 2, cell_height - border_width * 2)
        paste_x = cell_left + (cell_width - fitted.width) // 2
        paste_y = cell_top + (cell_height - fitted.height) // 2
        canvas.paste(fitted, (paste_x, paste_y))
        if border_width > 0:
            draw.rectangle(image_box, outline=border_color, width=border_width)

        panel_title = build_panel_title(panel, index, auto_labels)
        draw_centered_text(draw, title_box, panel_title, font, title_color)

    if output_path is None:
        output_dir = Path(tempfile.gettempdir()) / "thesis-word-template-cn"
        output_dir.mkdir(parents=True, exist_ok=True)
        stem = sanitize_stem(spec.get("output_stem") or spec.get("caption") or "panel_grid")
        output_path = output_dir / f"{stem}.png"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Compose a multi-panel figure grid from a JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to JSON spec file")
    parser.add_argument("--output", help="Output PNG path")
    args = parser.parse_args()

    spec_path = Path(args.spec).resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    output_path = Path(args.output).resolve() if args.output else None
    result = compose_from_spec(spec, spec_path.parent, output_path=output_path)
    print(result)


if __name__ == "__main__":
    main()
