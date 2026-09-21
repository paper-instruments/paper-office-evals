#!/usr/bin/env bash
set -u

artifact_dir=/logs/artifacts
report="$artifact_dir/treatment-fingerprint-before.json"
mkdir -p "$artifact_dir"

if python3 /usr/local/bin/pptx-ablation-fingerprint \
  --check /opt/pptx-ablation/treatment-fingerprint.json \
  --workdir /app \
  --output "$report"; then
  echo 0 > "$artifact_dir/treatment-fingerprint-before.status"
else
  echo 1 > "$artifact_dir/treatment-fingerprint-before.status"
  exit 70
fi

exec "$@"
