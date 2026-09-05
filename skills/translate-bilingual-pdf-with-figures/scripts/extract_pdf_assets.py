#!/usr/bin/env python3
"""Inventory an academic PDF and extract reviewable pages, text blocks, and images."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

import fitz


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract page renders, text, positioned blocks, and embedded images from a PDF."
    )
    parser.add_argument("source", type=Path, help="Input PDF")
    parser.add_argument("--output-dir", type=Path, required=True, help="Extraction directory")
    parser.add_argument("--dpi", type=int, default=150, help="Page-render DPI (default: 150)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = args.source.expanduser().resolve()
    output_dir = args.output_dir.expanduser().resolve()
    if not source.is_file():
        raise SystemExit(f"Source PDF not found: {source}")
    if args.dpi < 72 or args.dpi > 600:
        raise SystemExit("--dpi must be between 72 and 600")

    pages_dir = output_dir / "pages"
    images_dir = output_dir / "images"
    pages_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(source)
    zoom = args.dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    page_records: list[dict[str, Any]] = []
    block_records: list[dict[str, Any]] = []
    image_records: dict[int, dict[str, Any]] = {}
    image_occurrences: list[dict[str, int]] = []
    text_chunks: list[str] = []
    total_chars = 0

    for page_index, page in enumerate(doc):
        page_number = page_index + 1
        preview_name = f"page-{page_number:04d}.png"
        preview_path = pages_dir / preview_name
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        pix.save(preview_path)

        page_text = page.get_text("text") or ""
        total_chars += len(page_text.strip())
        text_chunks.append(f"\n===== PAGE {page_number} =====\n{page_text.rstrip()}\n")

        raw_blocks = page.get_text("blocks", sort=False)
        blocks: list[dict[str, Any]] = []
        for block_index, block in enumerate(raw_blocks):
            x0, y0, x1, y1, text_value, block_no, block_type = block[:7]
            blocks.append(
                {
                    "index": block_index,
                    "block_no": int(block_no),
                    "block_type": int(block_type),
                    "bbox": [round(float(v), 3) for v in (x0, y0, x1, y1)],
                    "text": str(text_value).rstrip(),
                }
            )
        block_records.append(
            {
                "page": page_number,
                "width": round(float(page.rect.width), 3),
                "height": round(float(page.rect.height), 3),
                "blocks": blocks,
            }
        )

        page_image_xrefs: list[int] = []
        for occurrence_index, image_info in enumerate(page.get_images(full=True)):
            xref = int(image_info[0])
            page_image_xrefs.append(xref)
            image_occurrences.append(
                {"page": page_number, "occurrence": occurrence_index + 1, "xref": xref}
            )
            if xref in image_records:
                continue
            extracted = doc.extract_image(xref)
            data = extracted["image"]
            extension = str(extracted.get("ext") or "bin").lower()
            image_name = f"image-xref-{xref:05d}.{extension}"
            image_path = images_dir / image_name
            image_path.write_bytes(data)
            image_records[xref] = {
                "xref": xref,
                "file": str(image_path.relative_to(output_dir)),
                "extension": extension,
                "width": int(extracted.get("width") or 0),
                "height": int(extracted.get("height") or 0),
                "colorspace": int(extracted.get("colorspace") or 0),
                "bytes": len(data),
                "sha256": sha256_bytes(data),
            }

        page_records.append(
            {
                "page": page_number,
                "width_points": round(float(page.rect.width), 3),
                "height_points": round(float(page.rect.height), 3),
                "text_characters": len(page_text.strip()),
                "block_count": len(blocks),
                "preview": str(preview_path.relative_to(output_dir)),
                "image_xrefs": page_image_xrefs,
            }
        )

    doc.close()
    (output_dir / "source.txt").write_text("".join(text_chunks), encoding="utf-8")
    (output_dir / "blocks.json").write_text(
        json.dumps(block_records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    manifest = {
        "source": str(source),
        "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "page_count": len(page_records),
        "render_dpi": args.dpi,
        "text_characters": total_chars,
        "likely_scanned": total_chars < max(100, len(page_records) * 40),
        "unique_embedded_images": len(image_records),
        "image_occurrences": image_occurrences,
        "images": list(image_records.values()),
        "pages": page_records,
        "review_warning": (
            "Embedded image objects are not equivalent to article figures. "
            "Use page renders to identify complete compound figures and reading order."
        ),
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(json.dumps({
        "output_dir": str(output_dir),
        "pages": len(page_records),
        "text_characters": total_chars,
        "likely_scanned": manifest["likely_scanned"],
        "unique_embedded_images": len(image_records),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
