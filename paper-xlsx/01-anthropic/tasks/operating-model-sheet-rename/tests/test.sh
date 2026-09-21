#!/usr/bin/env bash
set +e
umask 077

/tests/capability_test.sh
capability_status=$?

if ! /usr/local/bin/xlsx-ablation-verify-treatment; then
  mkdir -p /logs/verifier
  printf '0\n' > /logs/verifier/reward.txt
  printf '0\n' > /logs/verifier/score.txt
  cat > /logs/verifier/grading.json <<'JSON'
{
  "schema": "paper-office-treatment-integrity-failure",
  "version": 1,
  "score": 0.0,
  "hard_failures": ["treatment package changed after agent execution"]
}
JSON
  exit 0
fi

exit "$capability_status"
