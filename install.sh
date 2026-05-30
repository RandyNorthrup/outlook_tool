#!/usr/bin/env bash
# Outlook PST/OST Tool - setup for macOS / Linux.
# Creates a local .venv and installs the dependencies from requirements.txt.
set -euo pipefail
cd "$(dirname "$0")"

echo "========================================"
echo " Outlook PST/OST Tool - Setup"
echo "========================================"

# Locate a Python 3 interpreter.
if command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
else
    echo "ERROR: Python 3 is not installed or not on PATH." >&2
    echo "Install it from https://www.python.org/downloads/" >&2
    exit 1
fi
echo "Using: $("$PY" --version)"

# Create the virtual environment if it does not exist yet.
if [ ! -d .venv ]; then
    echo "Creating virtual environment in .venv ..."
    "$PY" -m venv .venv
fi

# shellcheck disable=SC1091
. .venv/bin/activate
python -m pip install --upgrade pip

echo "Installing dependencies ..."
if ! pip install -r requirements.txt; then
    echo
    echo "WARNING: pip install failed."
    echo "libpff-python is the only dependency that may need extra steps:"
    echo "  conda install -c conda-forge libpff-python"
    echo "  or see https://github.com/libyal/libpff for a manual build."
    exit 1
fi

echo
echo "Setup complete. Launch the app with:  ./run.sh"
