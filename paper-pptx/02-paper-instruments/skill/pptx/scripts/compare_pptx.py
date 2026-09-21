#!/usr/bin/env python3
"""Compare lineage-related PPTX files with bounded Paper package/deck reports."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from pptx import Presentation
from pptx.diff import diff_decks
from pptx.package import diff_package


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def truncate_list(values: list, limit: int) -> dict:
    return {
        "count": len(values),
        "truncated": len(values) > limit,
        "items": values[:limit],
    }


def compare(source: Path, candidate: Path, detail: str, limit: int, full: bool) -> dict:
    source_prs = Presentation(str(source))
    candidate_prs = Presentation(str(candidate))
    package = diff_package(str(source), str(candidate)).to_dict()
    deck_error = None
    try:
        deck = diff_decks(str(source), str(candidate), detail=detail).to_dict()
    except Exception as exc:
        from pptx.errors import PaperRefusal

        if not isinstance(exc, PaperRefusal):
            raise
        deck = None
        deck_error = {"exception_class": type(exc).__name__, "message": str(exc)}

    if full:
        package_payload: dict = package
        deck_payload: dict | None = deck
    else:
        package_payload = {
            "schema": package["schema"],
            "version": package["version"],
            "deltas": truncate_list(package["deltas"], limit),
        }
        deck_payload = (
            {
                "schema": deck["schema"],
                "version": deck["version"],
                "detail": deck["detail"],
                "slides_added": truncate_list(deck["slides_added"], limit),
                "slides_removed": truncate_list(deck["slides_removed"], limit),
                "slides_moved": truncate_list(deck["slides_moved"], limit),
                "slide_changes": truncate_list(deck["slide_changes"], limit),
                "package_changes": truncate_list(deck["package_changes"], limit),
            }
            if deck is not None
            else None
        )

    return {
        "schema": "paper-pptx-skill-comparison",
        "version": 1,
        "complete": deck_error is None,
        "source": {
            "path": str(source.resolve()),
            "sha256": sha256(source),
            "slide_count": len(source_prs.slides),
            "size_bytes": source.stat().st_size,
        },
        "candidate": {
            "path": str(candidate.resolve()),
            "sha256": sha256(candidate),
            "slide_count": len(candidate_prs.slides),
            "size_bytes": candidate.stat().st_size,
        },
        "package_diff": package_payload,
        "deck_diff": deck_payload,
        "deck_diff_refusal": deck_error,
        "scope_note": (
            "Deck matching uses permanent slide IDs and is intended for lineage-related "
            "files. Package/deck diffs are not visual comparison or PowerPoint evidence."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--detail", choices=("structure", "text", "full"), default="structure")
    parser.add_argument("--limit", type=int, default=100, help="Maximum records per list")
    parser.add_argument("--full", action="store_true", help="Do not truncate Paper reports")
    parser.add_argument("--output", type=Path, help="Write JSON here instead of stdout")
    args = parser.parse_args()
    if args.limit < 1:
        parser.error("--limit must be positive")
    for path in (args.source, args.candidate):
        if not path.is_file():
            parser.error(f"input does not exist: {path}")

    payload = compare(args.source, args.candidate, args.detail, args.limit, args.full)
    rendered = json.dumps(payload, indent=2, ensure_ascii=False) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0 if payload["complete"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
