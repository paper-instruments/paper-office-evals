#!/usr/bin/env bash
set -euo pipefail
fixture_root="${HARBOR_TASK_ROOT:-/app}"
mkdir -p "$fixture_root/evals/appearance-preserving-slide-handoff"
cp /solution/output.pptx "$fixture_root/evals/appearance-preserving-slide-handoff/output.pptx"
