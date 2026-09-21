#!/usr/bin/env python3
"""Emit a bounded, read-only structural summary of a PPTX using Paper APIs."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Iterable

from pptx import Presentation
from pptx.inspect import inspect_deck, inspect_text


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_slides(value: str | None, slide_count: int) -> list[int]:
    if value is None:
        return list(range(slide_count))
    positions: list[int] = []
    for token in value.split(","):
        token = token.strip()
        if not token:
            continue
        number = int(token)
        if number < 1 or number > slide_count:
            raise ValueError(f"slide number {number} outside 1..{slide_count}")
        index = number - 1
        if index not in positions:
            positions.append(index)
    return positions


def flatten_shapes(shapes: Iterable, parent: tuple[str, ...] = ()) -> list[dict]:
    records: list[dict] = []
    for shape in shapes:
        path = parent + (shape.name,)
        records.append(
            {
                "path": list(path),
                "shape_id": shape.shape_id,
                "name": shape.name,
                "kind": shape.kind,
                "z_index": shape.z_index,
                "placeholder_type": shape.placeholder_type,
                "geometry_emu": {
                    "x": shape.x,
                    "y": shape.y,
                    "cx": shape.cx,
                    "cy": shape.cy,
                    "rotation": shape.rotation,
                },
                "text_block_count": shape.text_block_count,
                "table": shape.table,
                "chart": shape.chart,
                "image": shape.image,
                "autofit": shape.autofit,
            }
        )
        records.extend(flatten_shapes(shape.children, path))
    return records


def summarize(path: Path, slide_spec: str | None, include_text: bool, limit: int) -> dict:
    prs = Presentation(str(path))
    manifest = inspect_deck(prs)
    selected = parse_slides(slide_spec, manifest.slide_count)
    slides: list[dict] = []
    for index in selected:
        item = manifest.slides[index]
        shapes = flatten_shapes(item.shapes)
        slide_record: dict = {
            "slide_number": index + 1,
            "slide_index_zero_based": index,
            "partname": item.part,
            "slide_id": item.slide_id,
            "layout_name": item.layout_name,
            "has_notes": item.has_notes,
            "alternate_content_count": item.alternate_content_count,
            "shape_count": len(shapes),
            "shapes_truncated": len(shapes) > limit,
            "shapes": shapes[:limit],
        }
        if include_text:
            text = inspect_text(prs.slides[index])
            blocks = [block.to_dict() for block in text.blocks]
            slide_record["text"] = {
                "blind_region_count": text.blind_region_count,
                "block_count": len(blocks),
                "blocks_truncated": len(blocks) > limit,
                "blocks": blocks[:limit],
            }
        slides.append(slide_record)
    return {
        "schema": "paper-pptx-skill-inspection",
        "version": 1,
        "source": str(path.resolve()),
        "sha256": sha256(path),
        "slide_width_emu": manifest.slide_width,
        "slide_height_emu": manifest.slide_height,
        "slide_count": manifest.slide_count,
        "masters": list(manifest.masters),
        "selected_slide_numbers": [index + 1 for index in selected],
        "slides": slides,
        "scope_note": (
            "This is a Paper structural/text inspection, not a complete rendering, "
            "chart-ownership, notes, or unsupported-content inventory."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--slides", help="Comma-separated one-based slide numbers")
    parser.add_argument("--text", action="store_true", help="Include bounded text blocks")
    parser.add_argument("--limit", type=int, default=100, help="Maximum shapes/blocks per slide")
    parser.add_argument("--output", type=Path, help="Write JSON here instead of stdout")
    args = parser.parse_args()
    if args.limit < 1:
        parser.error("--limit must be positive")
    if not args.pptx.is_file():
        parser.error(f"input does not exist: {args.pptx}")

    payload = summarize(args.pptx, args.slides, args.text, args.limit)
    rendered = json.dumps(payload, indent=2, ensure_ascii=False) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
