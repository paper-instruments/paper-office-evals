#!/usr/bin/env bash
set -u

artifact_dir=/logs/artifacts
report="$artifact_dir/treatment-fingerprint-before.json"
mkdir -p "$artifact_dir"

if python3 /usr/local/bin/docx-ablation-fingerprint \
  --expected "$DOCX_ABLATION_CONDITION" \
  --check /opt/docx-ablation/treatment-fingerprint.json \
  --output "$report"; then
  echo 0 > "$artifact_dir/treatment-fingerprint-before.status"
else
  echo 1 > "$artifact_dir/treatment-fingerprint-before.status"
  exit 70
fi

exec "$@"
