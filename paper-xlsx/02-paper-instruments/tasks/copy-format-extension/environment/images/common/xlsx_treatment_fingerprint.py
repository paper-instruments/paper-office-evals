#!/usr/bin/env python3
"""Emit and optionally verify the authoritative XLSX treatment identity."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import importlib.util
import json
import os
import sys
from pathlib import Path


SCAN_ROOTS = (Path("/home/agent"), Path("/tmp"), Path("/app"))
IGNORED_PARTS = {".cache", "__pycache__"}
EXPECTED_IDENTITIES = {
    "paper": {
        "paper_xlsx_distribution": "0.2.1",
        "openpyxl_distribution": None,
        "openpyxl_import_version": "3.1.5",
        "paper_sentinel": "0.2.1",
        "openpyxl_import_root": "/usr/local/lib/python3.13/site-packages/openpyxl",
    },
    "openpyxl": {
        "paper_xlsx_distribution": None,
        "openpyxl_distribution": "3.1.5",
        "openpyxl_import_version": "3.1.5",
        "paper_sentinel": None,
        "openpyxl_import_root": "/usr/local/lib/python3.13/site-packages/openpyxl",
    },
}


def package_tree_hash(root: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(root.rglob("*.py")):
        digest.update(path.relative_to(root).as_posix().encode())
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    return digest.hexdigest()


def distribution_version(name: str):
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return None


def alternate_install_candidates() -> list[str]:
    """Find active-looking spreadsheet installs below writable runtime roots."""
    matches: set[str] = set()
    for root in SCAN_ROOTS:
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if any(part in IGNORED_PARTS for part in path.parts):
                continue
            name = path.name.casefold().replace("-", "_")
            is_spreadsheet_metadata = (
                path.is_dir()
                and name.endswith(".dist_info")
                and name.startswith(("paper_xlsx", "openpyxl"))
            )
            is_shadow_package = (
                (path.is_dir() and name == "openpyxl" and (path / "__init__.py").is_file())
                or (path.is_file() and name in {"openpyxl.py", "paper_xlsx.py"})
            )
            if is_spreadsheet_metadata or is_shadow_package:
                matches.add(str(path))
    return sorted(matches)


def expected_identity_errors(condition: str, actual: dict) -> list[str]:
    expected = EXPECTED_IDENTITIES.get(condition)
    if expected is None:
        return [f"unsupported treatment condition: {condition!r}"]
    errors = [
        f"{key}: expected {value!r}, got {actual.get(key)!r}"
        for key, value in expected.items()
        if actual.get(key) != value
    ]
    import_root = actual.get("openpyxl_import_root")
    if import_root is None:
        errors.append("openpyxl import root is missing")
    elif any(
        Path(import_root).is_relative_to(root)
        for root in SCAN_ROOTS
        if root.exists()
    ):
        errors.append(f"openpyxl import root is writable treatment space: {import_root}")
    if not actual.get("openpyxl_tree_sha256"):
        errors.append("openpyxl source-tree hash is missing")
    return errors


def fingerprint(condition: str) -> dict:
    spec = importlib.util.find_spec("openpyxl")
    import_root = None
    sentinel = None
    import_version = None
    tree_hash = None
    if spec is not None and spec.origin:
        import openpyxl

        import_root = str(Path(spec.origin).resolve().parent)
        sentinel = getattr(openpyxl, "__paper_version__", None)
        import_version = getattr(openpyxl, "__version__", None)
        tree_hash = package_tree_hash(Path(import_root))
    result = {
        "schema": "paper-xlsx-ablation-treatment",
        "version": 2,
        "condition": condition,
        "python": sys.version.split()[0],
        "paper_xlsx_distribution": distribution_version("paper-xlsx"),
        "openpyxl_distribution": distribution_version("openpyxl"),
        "openpyxl_import_root": import_root,
        "openpyxl_import_version": import_version,
        "paper_sentinel": sentinel,
        "openpyxl_tree_sha256": tree_hash,
        "agent_install_candidates": alternate_install_candidates(),
    }
    errors = expected_identity_errors(condition, result)
    result["matches_expected_identity"] = not errors
    result["identity_errors"] = errors
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--condition", default=os.environ.get("XLSX_ABLATION_CONDITION", "unknown")
    )
    parser.add_argument("--check", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    actual = fingerprint(args.condition)
    status = 0
    payload_value = dict(actual)
    if args.check:
        expected = json.loads(args.check.read_text())
        expected.pop("matches_baseline", None)
        expected.pop("differences", None)
        differences = sorted(
            key
            for key in set(actual) | set(expected)
            if actual.get(key) != expected.get(key)
        )
        payload_value["matches_baseline"] = not differences
        payload_value["differences"] = differences
        status = 0 if not differences else 1
    else:
        payload_value["matches_baseline"] = True
        payload_value["differences"] = []
    if not actual["matches_expected_identity"] or actual["agent_install_candidates"]:
        status = 1
    payload = json.dumps(payload_value, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload)
    else:
        print(payload, end="")
    return status


if __name__ == "__main__":
    raise SystemExit(main())
