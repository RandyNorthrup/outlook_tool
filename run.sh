#!/usr/bin/env bash
# Outlook PST/OST Tool - launcher for macOS / Linux.
# Uses the local .venv if present, otherwise the system Python 3.
set -euo pipefail
cd "$(dirname "$0")"

if [ -f .venv/bin/activate ]; then
    # shellcheck disable=SC1091
    . .venv/bin/activate
    PY=python
elif command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=python
fi

exec "$PY" main.py
