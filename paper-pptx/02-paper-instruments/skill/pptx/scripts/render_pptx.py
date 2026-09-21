#!/usr/bin/env python3
"""Render a PPTX to slide PNGs through LibreOffice and Poppler.

Rendering is read-only evidence from LibreOffice, not Microsoft PowerPoint compatibility.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


def require_command(name: str) -> str:
    path = shutil.which(name)
    if path is None:
        raise RuntimeError(f"required command is unavailable on PATH: {name}")
    return path


def render(source: Path, output_dir: Path, dpi: int, force: bool) -> dict:
    soffice = require_command("soffice")
    pdftoppm = require_command("pdftoppm")
    output_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(output_dir.glob("slide-*.png"))
    if existing and not force:
        raise RuntimeError(
            f"{output_dir} already contains slide PNGs; use --force to replace them"
        )
    if force:
        for path in existing:
            path.unlink()

    with tempfile.TemporaryDirectory(prefix="paper-pptx-render-") as temporary:
        temp = Path(temporary)
        profile = temp / "lo-profile"
        conversion = subprocess.run(
            [
                soffice,
                "--headless",
                f"-env:UserInstallation={profile.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp),
                str(source.resolve()),
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if conversion.returncode != 0:
            raise RuntimeError(
                "LibreOffice conversion failed: "
                + (conversion.stderr.strip() or conversion.stdout.strip())
            )
        pdf = temp / f"{source.stem}.pdf"
        if not pdf.is_file():
            raise RuntimeError(
                "LibreOffice did not produce the expected PDF: "
                + (conversion.stdout.strip() or conversion.stderr.strip())
            )

        prefix = temp / "slide"
        raster = subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), str(pdf), str(prefix)],
            check=False,
            capture_output=True,
            text=True,
        )
        if raster.returncode != 0:
            raise RuntimeError(
                "Poppler rendering failed: "
                + (raster.stderr.strip() or raster.stdout.strip())
            )
        rendered = sorted(temp.glob("slide-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
        if not rendered:
            raise RuntimeError("Poppler produced no slide images")
        outputs: list[str] = []
        for number, image in enumerate(rendered, 1):
            destination = output_dir / f"slide-{number:03d}.png"
            shutil.copy2(image, destination)
            outputs.append(str(destination.resolve()))

    return {
        "schema": "paper-pptx-skill-render",
        "version": 1,
        "source": str(source.resolve()),
        "renderer": "LibreOffice PDF export plus Poppler pdftoppm",
        "dpi": dpi,
        "slide_count": len(outputs),
        "images": outputs,
        "scope_note": (
            "These images show LibreOffice rendering only. They do not prove Microsoft "
            "PowerPoint open-without-Repair behavior or nonvisual package preservation."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--json-output", type=Path)
    args = parser.parse_args()
    if not args.pptx.is_file():
        parser.error(f"input does not exist: {args.pptx}")
    if args.dpi < 36 or args.dpi > 600:
        parser.error("--dpi must be between 36 and 600")

    try:
        payload = render(args.pptx, args.output_dir, args.dpi, args.force)
    except RuntimeError as exc:
        parser.exit(1, f"error: {exc}\n")
    rendered = json.dumps(payload, indent=2, ensure_ascii=False) + "\n"
    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
