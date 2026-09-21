#!/usr/bin/env bash
set -euo pipefail
fixture_root="${HARBOR_TASK_ROOT:-/app}"
mkdir -p "$fixture_root/evals/revision-review"
cp /solution/report.json "$fixture_root/evals/revision-review/report.json"
