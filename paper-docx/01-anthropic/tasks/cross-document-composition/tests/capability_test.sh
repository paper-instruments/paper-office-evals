#!/bin/bash
set +e
umask 077

mkdir -p /logs/verifier
rm -f /logs/verifier/score.txt /logs/verifier/reward.txt /logs/verifier/grading.json
unset HARBOR_SCORE_PATH HARBOR_REWARD_PATH HARBOR_GRADING_PATH

score_channel=$(python3 /tests/grader.py cross-document-composition 2> /logs/verifier/test-output.txt)
status=$?

if ! printf '%s\n' "$score_channel" | python3 -c '
import json
import math
import os
import sys
import tempfile

root = "/logs/verifier"
record = json.load(sys.stdin)
if set(record) != {"schema", "version", "score", "grading"}:
    raise SystemExit("invalid score channel fields")
if record["schema"] != "docx-grader-score-channel" or record["version"] != 1:
    raise SystemExit("invalid score channel schema")
score = record["score"]
if isinstance(score, bool) or not isinstance(score, (int, float)):
    raise SystemExit("invalid score type")
score = float(score)
if not math.isfinite(score) or not 0.0 <= score <= 1.0:
    raise SystemExit("score outside [0, 1]")
grading = record["grading"]
if not isinstance(grading, dict) or grading.get("score") != round(score, 6):
    raise SystemExit("grading payload does not match score")

def publish(name, payload):
    descriptor, temporary = tempfile.mkstemp(prefix=".grader-owned-", dir=root)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, os.path.join(root, name))
    finally:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass

publish("score.txt", f"{score:.6f}\n")
publish("reward.txt", f"{score:.6f}\n")
publish("grading.json", json.dumps(grading, indent=2, sort_keys=True) + "\n")
'; then
  printf '0\n' > /logs/verifier/.grader-owned-zero
  mv -f /logs/verifier/.grader-owned-zero /logs/verifier/score.txt
  cp /logs/verifier/score.txt /logs/verifier/reward.txt
  python3 - <<'PY'
import json
from pathlib import Path

Path("/logs/verifier/grading.json").write_text(json.dumps({
    "schema": "paper-docx-harbor-grading",
    "version": 2,
    "score": 0.0,
    "hard_failures": ["verifier failed before publishing a valid score envelope"],
}, indent=2, sort_keys=True) + "\n")
PY
fi

if [ "$status" -ne 0 ]; then
  cat /logs/verifier/test-output.txt
fi

exit 0
