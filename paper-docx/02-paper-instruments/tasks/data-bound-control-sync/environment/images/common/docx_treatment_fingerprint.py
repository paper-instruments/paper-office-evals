#!/usr/bin/env python3
"""Fingerprint and validate the system-Python treatment without importing cwd."""

from __future__ import annotations

import argparse
import hashlib
import importlib
import importlib.metadata
import importlib.util
import json
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any


SCAN_ROOTS = (Path("/home/agent"), Path("/tmp"), Path("/app"))


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def distribution_fingerprint(name: str) -> dict[str, Any] | None:
    try:
        dist = importlib.metadata.distribution(name)
    except importlib.metadata.PackageNotFoundError:
        return None

    owned: list[dict[str, Any]] = []
    for relative in sorted(dist.files or (), key=str):
        located = Path(dist.locate_file(relative))
        if not located.is_file():
            continue
        owned.append(
            {
                "path": str(relative),
                "sha256": sha256(located),
                "size": located.stat().st_size,
            }
        )
    aggregate = hashlib.sha256()
    for item in owned:
        aggregate.update(item["path"].encode("utf-8"))
        aggregate.update(b"\0")
        aggregate.update(item["sha256"].encode("ascii"))
        aggregate.update(b"\0")
    return {
        "requested_name": name,
        "metadata_name": dist.metadata.get("Name"),
        "version": dist.version,
        "metadata_path": str(Path(dist._path).resolve()),
        "file_count": len(owned),
        "record_aggregate_sha256": aggregate.hexdigest(),
    }


def module_details() -> dict[str, Any] | None:
    spec = importlib.util.find_spec("docx")
    if spec is None:
        return None
    module = importlib.import_module("docx")
    origin = getattr(module, "__file__", None)
    return {
        "origin": None if origin is None else str(Path(origin).resolve()),
        "origin_sha256": None if origin is None else sha256(Path(origin)),
        "version": getattr(module, "__version__", None),
        "paper_version": getattr(module, "__paper_version__", None),
    }


def alternate_install_candidates(expected: str) -> list[str]:
    forbidden = {
        "paper": ("python_docx",),
        "python-docx": ("paper_docx",),
        "neither": ("paper_docx", "python_docx"),
    }[expected]
    matches: list[str] = []
    for root in SCAN_ROOTS:
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if ".cache" in path.parts:
                continue
            name = path.name.casefold().replace("-", "_")
            if (
                path.is_dir()
                and name.endswith(".dist_info")
                and any(name.startswith(prefix) for prefix in forbidden)
            ):
                matches.append(str(path))
    return sorted(matches)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def validate(expected: str, payload: dict[str, Any]) -> None:
    paper = payload["distributions"]["paper-docx"]
    upstream = payload["distributions"]["python-docx"]
    module = payload["docx_module"]
    require(
        not payload["agent_install_candidates"],
        "agent-writable location contains the alternate DOCX distribution: "
        f"{payload['agent_install_candidates']}",
    )
    require(payload["marker"] == expected, "image treatment marker mismatch")
    if expected == "paper":
        require(
            paper is not None and paper["version"] == "0.2.0", "Paper 0.2.0 missing"
        )
        require(upstream is None, "upstream distribution present in Paper treatment")
        require(
            module is not None and module["paper_version"] == "0.2.0",
            "Paper sentinel missing",
        )
        require(
            payload["doctor_exit_code"] == 0,
            "paper-docx-doctor rejected Paper treatment",
        )
    elif expected == "python-docx":
        require(paper is None, "Paper distribution present in upstream treatment")
        require(
            upstream is not None and upstream["version"] == "1.2.0",
            "python-docx 1.2.0 missing",
        )
        require(
            module is not None and module["paper_version"] is None,
            "Paper sentinel in upstream treatment",
        )
    elif expected == "neither":
        require(
            paper is None and upstream is None,
            "DOCX distribution present in raw treatment",
        )
        require(module is None, "docx import is available in raw treatment")
    else:
        raise RuntimeError(f"unknown expected treatment: {expected}")


def comparison_view(payload: dict[str, Any]) -> dict[str, Any]:
    """Return immutable package identity fields, excluding runtime user details."""
    return {
        "schema": payload.get("schema"),
        "version": payload.get("version"),
        "expected": payload.get("expected"),
        "marker": payload.get("marker"),
        "python": payload.get("python"),
        "distributions": payload.get("distributions"),
        "docx_module": payload.get("docx_module"),
        "doctor_exit_code": payload.get("doctor_exit_code"),
        "agent_install_candidates": payload.get("agent_install_candidates"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--expected",
        choices=("paper", "python-docx", "neither"),
        default=os.environ.get("DOCX_ABLATION_CONDITION"),
    )
    parser.add_argument(
        "--check", type=Path, help="Compare against a build-time fingerprint"
    )
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if not args.expected:
        parser.error("--expected or DOCX_ABLATION_CONDITION is required")

    marker_path = Path("/opt/docx-ablation/condition")
    doctor_exit_code = None
    if args.expected == "paper":
        doctor_exit_code = subprocess.run(
            [sys.executable, "-I", "-m", "paper_docx_doctor"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        ).returncode

    payload: dict[str, Any] = {
        "schema": "paper-docx-ablation-treatment",
        "version": 1,
        "expected": args.expected,
        "marker": marker_path.read_text(encoding="utf-8").strip(),
        "uid": os.getuid(),
        "user": os.environ.get("USER"),
        "python": {
            "executable": sys.executable,
            "version": platform.python_version(),
            "implementation": platform.python_implementation(),
        },
        "distributions": {
            "paper-docx": distribution_fingerprint("paper-docx"),
            "python-docx": distribution_fingerprint("python-docx"),
            "lxml": distribution_fingerprint("lxml"),
            "typing-extensions": distribution_fingerprint("typing-extensions"),
        },
        "docx_module": module_details(),
        "doctor_exit_code": doctor_exit_code,
        "agent_install_candidates": alternate_install_candidates(args.expected),
    }
    validate(args.expected, payload)

    status = 0
    if args.check is not None:
        baseline = json.loads(args.check.read_text(encoding="utf-8"))
        current_view = comparison_view(payload)
        baseline_view = comparison_view(baseline)
        differences = sorted(
            key
            for key in set(current_view) | set(baseline_view)
            if current_view.get(key) != baseline_view.get(key)
        )
        payload["matches_baseline"] = not differences
        payload["differences"] = differences
        status = 0 if not differences else 3
    else:
        payload["matches_baseline"] = True
        payload["differences"] = []

    encoded = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        temporary = args.output.with_suffix(args.output.suffix + ".tmp")
        temporary.write_text(encoded, encoding="utf-8")
        temporary.replace(args.output)
    else:
        sys.stdout.write(encoded)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
