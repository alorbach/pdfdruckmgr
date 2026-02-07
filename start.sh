#!/bin/bash
set -e

echo "Creating/using venv..."
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate

echo "Installing dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo ""
echo "Starting Druck Manager..."
python druckmgr.py
