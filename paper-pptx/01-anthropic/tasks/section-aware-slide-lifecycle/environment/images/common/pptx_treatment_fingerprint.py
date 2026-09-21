#!/usr/bin/env python3
"""Create and verify an immutable PPTX ablation treatment fingerprint."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import importlib.util
import json
import os
from pathlib import Path
import sys
from typing import Any


SCHEMA = "paper-pptx-ablation-treatment-fingerprint"
SCHEMA_VERSION = 1
PRESENTATION_DISTRIBUTIONS = ("paper-pptx", "python-pptx")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical_sha256(value: Any) -> str:
    payload = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(payload).hexdigest()


def distribution_snapshot(name: str) -> dict[str, Any] | None:
    try:
        distribution = importlib.metadata.distribution(name)
    except importlib.metadata.PackageNotFoundError:
        return None

    files: list[dict[str, Any]] = []
    for relative in sorted(distribution.files or [], key=lambda item: str(item)):
        path = Path(distribution.locate_file(relative))
        if not path.is_file() or path.suffix == ".pyc" or "__pycache__" in path.parts:
            continue
        files.append(
            {
                "path": str(relative),
                "size": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )

    return {
        "name": name,
        "version": distribution.version,
        "metadata_path": str(Path(distribution._path).resolve()),
        "files_sha256": canonical_sha256(files),
        "files": files,
    }


def module_tree_snapshot(root: Path) -> dict[str, Any]:
    files: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix == ".pyc" or "__pycache__" in path.parts:
            continue
        files.append(
            {
                "path": path.relative_to(root).as_posix(),
                "size": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
    return {"root": str(root.resolve()), "files_sha256": canonical_sha256(files)}


def pptx_module_snapshot() -> dict[str, Any] | None:
    spec = importlib.util.find_spec("pptx")
    if spec is None:
        return None

    origin = Path(spec.origin).resolve() if spec.origin else None
    search_locations = list(spec.submodule_search_locations or [])
    root = Path(search_locations[0]).resolve() if search_locations else None
    compatibility_version = None
    paper_version = None
    import_error = None
    try:
        import pptx  # type: ignore

        compatibility_version = getattr(pptx, "__version__", None)
        paper_version = getattr(pptx, "__paper_version__", None)
    except Exception as exc:  # pragma: no cover - retained as fingerprint evidence
        import_error = f"{type(exc).__name__}: {exc}"

    owners = sorted(importlib.metadata.packages_distributions().get("pptx", []))
    return {
        "origin": str(origin) if origin else None,
        "owners": owners,
        "compatibility_version": compatibility_version,
        "paper_version": paper_version,
        "import_error": import_error,
        "tree": module_tree_snapshot(root) if root and root.is_dir() else None,
    }


def suspicious_workdir_entries(workdir: Path | None) -> list[str]:
    if workdir is None or not workdir.exists():
        return []

    matches: list[str] = []
    for path in workdir.rglob("*"):
        name = path.name.lower().replace("-", "_")
        if name == "__pycache__":
            continue
        suspicious = (
            (path.is_dir() and name == "pptx" and (path / "__init__.py").is_file())
            or (path.is_file() and name == "pptx.py")
            or (
                path.is_dir()
                and name.startswith("paper_pptx")
                and name.endswith(".dist_info")
            )
            or (
                path.is_dir()
                and name.startswith("python_pptx")
                and name.endswith(".dist_info")
            )
        )
        if suspicious:
            matches.append(path.relative_to(workdir).as_posix())
    return sorted(matches)


def suspicious_agent_entries(condition: str) -> list[str]:
    """Find the opposite treatment distribution in agent-writable locations."""
    forbidden = {
        "paper": ("python_pptx",),
        "python-pptx": ("paper_pptx",),
        "raw": ("paper_pptx", "python_pptx"),
    }[condition]
    roots = (Path("/home/agent"), Path("/tmp"))
    matches: list[str] = []
    for root in roots:
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if ".cache" in path.parts:
                continue
            name = path.name.lower().replace("-", "_")
            if name == "__pycache__":
                continue
            suspicious = (
                path.is_dir()
                and name.endswith(".dist_info")
                and any(name.startswith(prefix) for prefix in forbidden)
            )
            if suspicious:
                matches.append(str(path))
    return sorted(matches)


def condition_assertions(
    condition: str,
    distributions: dict[str, dict[str, Any] | None],
    module: dict[str, Any] | None,
) -> list[str]:
    errors: list[str] = []
    paper = distributions["paper-pptx"]
    upstream = distributions["python-pptx"]

    if condition == "paper":
        if paper is None or paper.get("version") != "0.1.2":
            errors.append("paper condition requires exactly paper-pptx==0.1.2")
        if upstream is not None:
            errors.append("paper condition must not contain python-pptx metadata")
        if module is None:
            errors.append("paper condition requires an importable pptx package")
        else:
            if module.get("owners") != ["paper-pptx"]:
                errors.append(
                    "paper condition requires paper-pptx to be the sole pptx owner"
                )
            if module.get("compatibility_version") != "1.0.2":
                errors.append("paper condition requires pptx.__version__ == 1.0.2")
            if module.get("paper_version") != "0.1.2":
                errors.append(
                    "paper condition requires pptx.__paper_version__ == 0.1.2"
                )
    elif condition == "python-pptx":
        if upstream is None or upstream.get("version") != "1.0.2":
            errors.append("python-pptx condition requires exactly python-pptx==1.0.2")
        if paper is not None:
            errors.append("python-pptx condition must not contain paper-pptx metadata")
        if module is None:
            errors.append("python-pptx condition requires an importable pptx package")
        else:
            if module.get("owners") != ["python-pptx"]:
                errors.append(
                    "python-pptx condition requires python-pptx to be the sole pptx owner"
                )
            if module.get("compatibility_version") != "1.0.2":
                errors.append(
                    "python-pptx condition requires pptx.__version__ == 1.0.2"
                )
            if module.get("paper_version") is not None:
                errors.append(
                    "python-pptx condition must not expose a Paper version sentinel"
                )
    elif condition == "raw":
        if paper is not None or upstream is not None:
            errors.append(
                "raw condition must contain neither presentation distribution"
            )
        if module is not None:
            errors.append("raw condition must not have an importable pptx package")
    else:
        errors.append(f"unknown treatment condition: {condition!r}")
    return errors


def create_snapshot(condition: str, workdir: Path | None) -> dict[str, Any]:
    distributions = {
        name: distribution_snapshot(name) for name in PRESENTATION_DISTRIBUTIONS
    }
    module = pptx_module_snapshot()
    assertions = condition_assertions(condition, distributions, module)
    return {
        "schema": SCHEMA,
        "schema_version": SCHEMA_VERSION,
        "condition": condition,
        "python": {
            "version": ".".join(str(part) for part in sys.version_info[:3]),
            "executable": str(Path(sys.executable).resolve()),
        },
        "distributions": distributions,
        "pptx_module": module,
        "workdir_pptx_candidates": suspicious_workdir_entries(workdir),
        "agent_pptx_candidates": suspicious_agent_entries(condition),
        "condition_assertions": assertions,
    }


def comparison_view(snapshot: dict[str, Any]) -> dict[str, Any]:
    """Return only immutable treatment fields used for exact comparison."""
    return {
        "schema": snapshot.get("schema"),
        "schema_version": snapshot.get("schema_version"),
        "condition": snapshot.get("condition"),
        "python": snapshot.get("python"),
        "distributions": snapshot.get("distributions"),
        "pptx_module": snapshot.get("pptx_module"),
        "workdir_pptx_candidates": snapshot.get("workdir_pptx_candidates"),
        "agent_pptx_candidates": snapshot.get("agent_pptx_candidates"),
        "condition_assertions": snapshot.get("condition_assertions"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--condition",
        choices=("paper", "python-pptx", "raw"),
        default=os.environ.get("PPTX_ABLATION_CONDITION"),
    )
    parser.add_argument("--workdir", type=Path)
    parser.add_argument("--check", type=Path, help="Compare against a baseline JSON")
    parser.add_argument(
        "--output", type=Path, help="Write the full report to this path"
    )
    args = parser.parse_args()
    if not args.condition:
        parser.error("--condition or PPTX_ABLATION_CONDITION is required")

    current = create_snapshot(args.condition, args.workdir)
    report: dict[str, Any]
    status = 0
    if args.check:
        baseline = json.loads(args.check.read_text())
        baseline_view = comparison_view(baseline)
        current_view = comparison_view(current)
        differences: list[str] = []
        for key in baseline_view:
            if baseline_view[key] != current_view[key]:
                differences.append(key)
        report = {
            "schema": "paper-pptx-ablation-treatment-verification",
            "schema_version": 1,
            "matches_baseline": not differences,
            "differences": differences,
            "baseline_sha256": sha256_file(args.check),
            "current": current,
        }
        status = 0 if not differences else 3
    else:
        report = current
        status = 0 if not current["condition_assertions"] else 2

    rendered = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered)
    else:
        sys.stdout.write(rendered)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
