#!/usr/bin/env python3
"""Run a Harbor grader with an isolated score channel and atomic publication."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def run_cleanup(grader: Path) -> None:
    subprocess.run(
        [sys.executable, str(grader), "--cleanup-candidate-processes"],
        check=False,
        close_fds=True,
    )


def remove_untrusted_path(path: Path) -> None:
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.exists():
        shutil.rmtree(path)


def reset_verifier_directory(path: Path) -> None:
    """Clear verifier outputs without removing Harbor's bind-mount root."""

    if path.is_symlink() or path.is_file():
        path.unlink()
        path.mkdir(parents=True, mode=0o700)
    elif path.exists():
        for child in path.iterdir():
            remove_untrusted_path(child)
    else:
        path.mkdir(parents=True, mode=0o700)
    path.chmod(0o700)


def score_fd_path(fd: int) -> str:
    proc_path = Path(f"/proc/self/fd/{fd}")
    if proc_path.exists():
        return str(proc_path)
    return f"/dev/fd/{fd}"


def run(root: Path, grader: Path, verifier_dir: Path) -> None:
    os.umask(0o077)
    run_cleanup(grader)
    reset_verifier_directory(verifier_dir)
    private_dir = Path(tempfile.mkdtemp(prefix=".score-", dir=verifier_dir))
    reward_draft = private_dir / "reward"
    grading_draft = private_dir / "grading.json"
    reward_fd = os.open(reward_draft, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        subprocess.run(
            [
                sys.executable,
                str(grader),
                str(root),
                score_fd_path(reward_fd),
                str(grading_draft),
            ],
            check=True,
            close_fds=True,
            pass_fds=(reward_fd,),
        )
        os.fsync(reward_fd)
    finally:
        os.close(reward_fd)

    run_cleanup(grader)
    if not reward_draft.is_file() or not grading_draft.is_file():
        raise RuntimeError("grader did not produce both score artifacts")
    os.replace(reward_draft, verifier_dir / "reward.txt")
    os.replace(grading_draft, verifier_dir / "grading.json")
    (verifier_dir / "reward.txt").chmod(0o600)
    (verifier_dir / "grading.json").chmod(0o600)
    private_dir.rmdir()


def publish_failure(verifier_dir: Path, exc: BaseException) -> None:
    """Publish a valid zero reward even when the verifier process fails."""

    reset_verifier_directory(verifier_dir)
    private_dir = Path(tempfile.mkdtemp(prefix=".score-", dir=verifier_dir))
    reward_draft = private_dir / "reward"
    grading_draft = private_dir / "grading.json"
    reward_draft.write_text("0.000000\n")
    grading_draft.write_text(
        json.dumps(
            {
                "reward": 0.0,
                "hard_failures": [
                    f"verifier runner abort: {type(exc).__name__}: {str(exc)[:1000]}"
                ],
                "checks": [],
            },
            indent=2,
            sort_keys=True,
        )
        + "\n"
    )
    os.replace(reward_draft, verifier_dir / "reward.txt")
    os.replace(grading_draft, verifier_dir / "grading.json")
    private_dir.rmdir()


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: secure_runner.py ROOT GRADER VERIFIER_DIR", file=sys.stderr)
        return 2
    verifier_dir = Path(sys.argv[3])
    try:
        run(Path(sys.argv[1]), Path(sys.argv[2]), verifier_dir)
    except BaseException as exc:
        publish_failure(verifier_dir, exc)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
