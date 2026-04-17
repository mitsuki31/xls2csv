#!/usr/bin/env bash

set -e

VENV_DIR=".venv"

echo "[1/3] Creating virtual environment..."
python3 -m venv "$VENV_DIR"

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