#!/usr/bin/env python
"""Render PDF pages to PNGs with Poppler or a PyMuPDF fallback."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render selected PDF pages to PNG images with a stable fallback."
    )
    parser.add_argument("input_path", help="Path to the input PDF file")
    parser.add_argument(
        "--pages",
        help="Pages to render, e.g. 1,3-5. Defaults to all pages.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory where rendered PNG pages will be written",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="Target DPI for rendering. Defaults to 200.",
    )
    return parser.parse_args()


def resolve_pdf(path_arg: str) -> Path:
    pdf_path = Path(path_arg).expanduser().resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF.")
    return pdf_path


def parse_page_spec(page_spec: str | None, page_count: int) -> list[int]:
    if not page_spec:
        return list(range(1, page_count + 1))

    pages: set[int] = set()
    for part in page_spec.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            start = int(start_text)
            end = int(end_text)
            if start > end:
                raise ValueError(f"Invalid page range: {token}")
            pages.update(range(start, end + 1))
        else:
            pages.add(int(token))

    invalid = [p for p in pages if p < 1 or p > page_count]
    if invalid:
        raise ValueError(f"Page(s) out of range: {invalid}")
    return sorted(pages)


def render_with_pdftoppm(
    pdf_path: Path, pages: list[int], output_dir: Path, dpi: int
) -> list[dict[str, object]]:
    rendered: list[dict[str, object]] = []
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise FileNotFoundError("pdftoppm is not available")

    for page_number in pages:
        prefix = output_dir / f"page_{page_number:03d}"
        cmd = [
            pdftoppm,
            "-png",
            "-r",
            str(dpi),
            "-f",
            str(page_number),
            "-l",
            str(page_number),
            str(pdf_path),
            str(prefix),
        ]
        subprocess.run(cmd, check=True, capture_output=True)
        generated = prefix.with_name(prefix.name + "-1.png")
        target = output_dir / f"page_{page_number:03d}.png"
        generated.replace(target)
        rendered.append(
            {
                "page_number": page_number,
                "image_path": str(target),
            }
        )
    return rendered


def render_with_pymupdf(
    pdf_path: Path, pages: list[int], output_dir: Path, dpi: int
) -> list[dict[str, object]]:
    try:
        import fitz  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "PyMuPDF is required when pdftoppm is unavailable. Install it with `pip install pymupdf`."
        ) from exc

    rendered: list[dict[str, object]] = []
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)
    with fitz.open(pdf_path) as document:
        for page_number in pages:
            page = document.load_page(page_number - 1)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            target = output_dir / f"page_{page_number:03d}.png"
            pix.save(str(target))
            rendered.append(
                {
                    "page_number": page_number,
                    "image_path": str(target),
                    "width": pix.width,
                    "height": pix.height,
                }
            )
    return rendered


def get_page_count(pdf_path: Path) -> int:
    try:
        import fitz  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "PyMuPDF is required to inspect PDF page counts. Install it with `pip install pymupdf`."
        ) from exc
    with fitz.open(pdf_path) as document:
        return document.page_count


def render_pdf_pages(
    pdf_path: Path, output_dir: Path, pages: list[int] | None = None, dpi: int = 200
) -> dict[str, object]:
    output_dir.mkdir(parents=True, exist_ok=True)
    page_count = get_page_count(pdf_path)
    target_pages = parse_page_spec(
        ",".join(str(page) for page in pages) if pages else None, page_count
    )

    renderer = "pdftoppm"
    fallback_reason = None
    try:
        rendered = render_with_pdftoppm(pdf_path, target_pages, output_dir, dpi)
    except Exception as exc:
        renderer = "pymupdf"
        fallback_reason = str(exc)
        rendered = render_with_pymupdf(pdf_path, target_pages, output_dir, dpi)

    return {
        "pdf": str(pdf_path),
        "output_dir": str(output_dir),
        "renderer": renderer,
        "fallback_reason": fallback_reason,
        "page_count": page_count,
        "rendered_pages": rendered,
    }


def main() -> int:
    try:
        args = parse_args()
        pdf_path = resolve_pdf(args.input_path)
        output_dir = Path(args.output_dir).expanduser().resolve()
        page_count = get_page_count(pdf_path)
        pages = parse_page_spec(args.pages, page_count)
        result = render_pdf_pages(pdf_path, output_dir, pages=pages, dpi=args.dpi)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:  # pragma: no cover - CLI wrapper
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
