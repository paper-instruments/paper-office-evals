#!/usr/bin/env bash
set -u

baseline=/opt/pptx-ablation/treatment-fingerprint.json
verifier_dir=/logs/verifier
artifact_dir=/logs/artifacts
report="$verifier_dir/treatment-fingerprint-after.json"
before="$artifact_dir/treatment-fingerprint-before.json"

mkdir -p "$verifier_dir" "$artifact_dir"
python3 /usr/local/bin/pptx-ablation-fingerprint \
  --check "$baseline" \
  --workdir /app \
  --output "$report"
status=$?

if [ ! -s "$before" ] || ! python3 - "$before" <<'PY'
import json
import sys

try:
    valid = json.load(open(sys.argv[1], encoding="utf-8")).get("matches_baseline") is True
except Exception:
    valid = False
raise SystemExit(0 if valid else 1)
PY
then
  status=4
fi
cp "$report" "$artifact_dir/treatment-fingerprint-after.json"
exit "$status"
