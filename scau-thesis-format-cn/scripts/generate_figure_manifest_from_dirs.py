#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

from PIL import Image


FIGURE_DIR_RE = re.compile(r"^实验结果(?P<number>\d+-\d+)_(?P<title>.+)$")
SIDE_CAR_NAMES = ("装版配置.json", "figure-manifest.json", "manifest.json")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a figure-block manifest by scanning thesis figure directories."
    )
    parser.add_argument("--figures-root", required=True, help="Root directory containing figure folders.")
    parser.add_argument("--output", required=True, help="Output JSON manifest path.")
    parser.add_argument(
        "--number-prefix",
        help="Only include figure numbers that start with this prefix, such as 3-",
    )
    parser.add_argument(
        "--position",
        default="after",
        choices=["before", "after"],
        help="Whether each figure block should be inserted before or after the first matching anchor paragraph.",
    )
    parser.add_argument(
        "--run-generators",
        action="store_true",
        help="Run any *_生成脚本.py scripts found in each figure directory before manifest generation.",
    )
    parser.add_argument("--chapter-file", help="Optional chapter Markdown file used to infer more precise anchor phrases.")
    parser.add_argument("--overrides-file", help="Optional JSON overrides keyed by figure number or folder name.")
    return parser.parse_args()


def figure_dirs(root: Path, number_prefix: str | None) -> list[tuple[str, str, Path]]:
    discovered: list[tuple[str, str, Path]] = []
    for child in root.iterdir():
        if not child.is_dir():
            continue
        match = FIGURE_DIR_RE.match(child.name)
        if not match:
            continue
        number = match.group("number")
        if number_prefix and not number.startswith(number_prefix):
            continue
        title = match.group("title").replace("_", " ").strip()
        discovered.append((number, title, child))
    return sorted(discovered, key=lambda item: tuple(int(part) for part in item[0].split("-")))


def run_generators(figure_dir: Path) -> None:
    for script in sorted(figure_dir.glob("*_生成脚本.py")):
        subprocess.run([sys.executable, str(script)], check=True, cwd=str(figure_dir))


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def load_overrides(path: Path | None) -> dict:
    if path is None:
        return {}
    return load_json(path)


def load_sidecar(figure_dir: Path) -> dict:
    candidates = [figure_dir / name for name in SIDE_CAR_NAMES]
    candidates.extend(figure_dir.glob("*_装版配置.json"))
    for candidate in candidates:
        if candidate.exists():
            return load_json(candidate)
    return {}


def resolve_override(overrides: dict, number: str, figure_dir: Path) -> dict:
    for key in (number, figure_dir.name, figure_dir.name.replace("_", " ")):
        value = overrides.get(key)
        if isinstance(value, dict):
            return value
    return {}


def choose_image(figure_dir: Path) -> Path:
    preferred_stems = [figure_dir.name, figure_dir.name.replace("_", " ")]
    for stem in preferred_stems:
        for ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
            candidate = figure_dir / f"{stem}{ext}"
            if candidate.exists():
                return candidate

    raster_files = [
        path
        for path in figure_dir.iterdir()
        if path.is_file() and path.suffix.lower() in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
    ]
    if not raster_files:
        raise FileNotFoundError(f"未找到可插入 Word 的图片文件: {figure_dir}")
    return max(raster_files, key=lambda path: path.stat().st_size)


def read_note(figure_dir: Path) -> str:
    preferred = figure_dir / f"{figure_dir.name}_图注.txt"
    if preferred.exists():
        return preferred.read_text(encoding="utf-8").strip()
    candidates = sorted(figure_dir.glob("*_图注.txt"))
    if not candidates:
        return ""
    return candidates[0].read_text(encoding="utf-8").strip()


def infer_width_cm(image_path: Path) -> float:
    with Image.open(image_path) as image:
        width, height = image.size
    ratio = width / max(height, 1)
    if ratio >= 2.1:
        return 15.0
    if ratio >= 1.45:
        return 14.0
    if ratio >= 1.15:
        return 12.8
    return 11.5


def normalize_override_title(text: str) -> str:
    normalized = text.strip()
    if normalized.startswith("图"):
        match = re.match(r"^图\d+(?:[-–]\d+)?\s+(.+)$", normalized)
        if match:
            return match.group(1).strip()
    return normalized


def load_chapter_paragraphs(path: Path | None) -> list[str]:
    if path is None:
        return []
    content = path.read_text(encoding="utf-8")
    chunks = [re.sub(r"\s+", " ", chunk).strip() for chunk in re.split(r"\n\s*\n", content)]
    return [chunk for chunk in chunks if chunk]


def infer_anchor_from_chapter(number: str, chapter_paragraphs: list[str]) -> str | None:
    direct_patterns = [
        f"如图{number}所示",
        f"结果见图{number}",
        f"见图{number}",
        f"图{number}所示",
        f"图{number}显示",
        f"图{number}可见",
    ]
    for paragraph in chapter_paragraphs:
        for phrase in direct_patterns:
            if phrase in paragraph:
                return re.escape(phrase)
    return None


def default_anchor_regex(number: str) -> str:
    return rf"(?:如图{re.escape(number)}所示|结果见图{re.escape(number)}|见图{re.escape(number)}|图{re.escape(number)}所示|图{re.escape(number)}显示|图{re.escape(number)}可见)"


def manual_review_fields(meta_sources: dict[str, str]) -> list[str]:
    review_fields = []
    for field, source in meta_sources.items():
        if source == "auto":
            review_fields.append(field)
    return review_fields


def build_entry(
    number: str,
    title: str,
    figure_dir: Path,
    position: str,
    *,
    chapter_paragraphs: list[str],
    overrides: dict,
) -> dict:
    image_path = choose_image(figure_dir)
    note = read_note(figure_dir)
    sidecar = load_sidecar(figure_dir)
    override = resolve_override(overrides, number, figure_dir)

    merged = {**sidecar, **override}

    caption_title = title
    caption_source = "auto"
    if merged.get("caption"):
        caption_title = normalize_override_title(str(merged["caption"]))
        caption_source = "override"

    chapter_anchor = infer_anchor_from_chapter(number, chapter_paragraphs)
    anchor_regex = chapter_anchor
    anchor_source = "auto"
    if anchor_regex is None:
        anchor_regex = default_anchor_regex(number)
    if merged.get("anchor_regex"):
        anchor_regex = str(merged["anchor_regex"])
        anchor_source = "override"
    elif chapter_anchor:
        anchor_source = "chapter"

    width_cm = infer_width_cm(image_path)
    width_source = "auto"
    if merged.get("width_cm") is not None:
        width_cm = float(merged["width_cm"])
        width_source = "override"

    page_break_before = None
    if merged.get("page_break_before") is not None:
        page_break_before = bool(merged["page_break_before"])
    elif len(note) >= 120 and width_cm >= 14.0:
        page_break_before = True

    meta_sources = {
        "caption": caption_source,
        "anchor_regex": anchor_source,
        "width_cm": width_source,
    }

    return {
        "anchor_regex": anchor_regex,
        "occurrence": 1,
        "position": position,
        "image": str(image_path),
        "width_cm": width_cm,
        "caption": f"图{number} {caption_title}",
        "note": note,
        "page_break_before": page_break_before,
        "_meta": {
            "figure_number": number,
            "figure_dir": str(figure_dir),
            "sources": meta_sources,
            "requires_review": manual_review_fields(meta_sources),
        },
    }


def main() -> None:
    args = parse_args()
    figures_root = Path(args.figures_root).resolve()
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    chapter_paragraphs = load_chapter_paragraphs(Path(args.chapter_file).resolve()) if args.chapter_file else []
    overrides = load_overrides(Path(args.overrides_file).resolve()) if args.overrides_file else {}

    entries: list[dict] = []
    for number, title, figure_dir in figure_dirs(figures_root, args.number_prefix):
        if args.run_generators:
            run_generators(figure_dir)
        entries.append(
            build_entry(
                number,
                title,
                figure_dir,
                args.position,
                chapter_paragraphs=chapter_paragraphs,
                overrides=overrides,
            )
        )

    output_path.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(output_path), "entries": len(entries)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
