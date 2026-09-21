#!/usr/bin/env bash
set +e
set -u
umask 077

ROOT="${HARBOR_TASK_ROOT:-$(pwd)}"
verifier_dir=/logs/verifier
mkdir -p "$verifier_dir"
chmod 700 "$verifier_dir"
rm -f "$verifier_dir/reward.txt" "$verifier_dir/grading.json"

kill_agent_processes() {
  if command -v pkill >/dev/null 2>&1 && id agent >/dev/null 2>&1; then
    pkill -KILL -u "$(id -u agent)" >/dev/null 2>&1 || true
  fi
}

# Agent-phase children are never part of verification and must not survive into
# the predictable Harbor publication namespace.
kill_agent_processes

private_dir="$(mktemp -d /tmp/paper-pptx-verifier.XXXXXX)"
chmod 700 "$private_dir"
cleanup() {
  kill_agent_processes
  rm -rf "$private_dir"
}
trap cleanup EXIT HUP INT TERM

grader_status=0
exec 9>"$private_dir/result-envelope.json"
python3 /tests/grader.py "$ROOT" --result-fd 9
grader_status=$?
exec 9>&-

# Grader-spawned solution/probe processes run in supervised groups. This final
# sweep also removes any agent-owned process that tried to detach before score
# publication.
kill_agent_processes

if [ "$grader_status" -eq 0 ] && [ -s "$private_dir/result-envelope.json" ]; then
  python3 - "$private_dir/result-envelope.json" "$private_dir/reward.txt" "$private_dir/grading.json" <<'PY'
import json
import pathlib
import sys

envelope = json.loads(pathlib.Path(sys.argv[1]).read_text())
if envelope.get("schema") != "paper-pptx-verifier-result-envelope":
    raise SystemExit(2)
reward = float(envelope["reward"])
grading = envelope["grading"]
if not 0.0 <= reward <= 1.0 or not isinstance(grading, dict):
    raise SystemExit(3)
pathlib.Path(sys.argv[2]).write_text(("%.4f" % reward).rstrip("0").rstrip(".") + "\n")
pathlib.Path(sys.argv[3]).write_text(json.dumps(grading, indent=2, sort_keys=True) + "\n")
PY
  grader_status=$?
fi

if [ "$grader_status" -ne 0 ] || [ ! -s "$private_dir/reward.txt" ]; then
  printf '0\n' > "$private_dir/reward.txt"
  python3 - "$private_dir/grading.json" "$grader_status" <<'PY'
import json
import pathlib
import sys

pathlib.Path(sys.argv[1]).write_text(json.dumps({
    "schema": "paper-pptx-harbor-grading",
    "version": 2,
    "reward": 0.0,
    "hard_failures": ["verifier runner failed before score publication (status=%s)" % sys.argv[2]],
}, indent=2, sort_keys=True) + "\n")
PY
fi

install -m 600 "$private_dir/reward.txt" "$verifier_dir/.reward.txt.$$"
install -m 600 "$private_dir/grading.json" "$verifier_dir/.grading.json.$$"
mv -f "$verifier_dir/.grading.json.$$" "$verifier_dir/grading.json"
mv -f "$verifier_dir/.reward.txt.$$" "$verifier_dir/reward.txt"
exit 0
