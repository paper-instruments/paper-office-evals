#!/bin/sh
set -eu

# Keep the fingerprinted system treatment first while leaving unrelated user-site
# tools available to supported agent harnesses.
export PYTHONPATH=/usr/local/lib/python3.13/site-packages
export PYTHONSAFEPATH=1

artifact_dir=/logs/artifacts
mkdir -p "$artifact_dir"
if python3 /usr/local/bin/xlsx-ablation-fingerprint \
    --check /opt/xlsx-ablation/treatment-fingerprint.json \
    --output "$artifact_dir/treatment-fingerprint-before.json"; then
  echo 0 > "$artifact_dir/treatment-fingerprint-before.status"
else
  echo 1 > "$artifact_dir/treatment-fingerprint-before.status"
  exit 70
fi

exec "$@"
