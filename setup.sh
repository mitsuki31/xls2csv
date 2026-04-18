#!/usr/bin/env bash

set -e

VENV_DIR=".venv"
PYTHON=""

if command -v python3 >/dev/null 2>&1; then
    PYTHON="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON="python"
else
    echo "Python not found. Please install Python v3 and ensure it's in PATH." >&2
    exit 1
fi

echo "[1/3] Creating virtual environment..."
$PYTHON -m venv "$VENV_DIR"

echo "[2/3] Upgrading pip (inside venv)..."
"$VENV_DIR/bin/python" -m pip install --upgrade pip

echo "[3/3] Installing dependencies..."
"$VENV_DIR/bin/pip" install -r requirements.txt

echo ""
echo "Setup complete."
echo ""
echo "Activate with: source $VENV_DIR/bin/activate"
echo "Run with: $VENV_DIR/bin/python -m xls2csv.cli --help"

unset VENV_DIR
unset PYTHON