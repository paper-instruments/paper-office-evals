#!/bin/sh
set -eu
python /tests/secure_runner.py /app /tests/grader.py /logs/verifier
