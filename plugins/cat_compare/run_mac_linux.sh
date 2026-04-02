#!/usr/bin/env bash
set -e

# 1) Create venv if it doesn't exist
if [ ! -d ".venv" ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
fi

# 2) Activate venv
source .venv/bin/activate

# 3) Install deps (cached after first run)
echo "Installing requirements..."
pip install --upgrade pip
pip install -r requirements.txt

# 4) Run the app
echo "Starting Config Assessment Tool on http://127.0.0.1:5000 ..."
python comparewebapp/app.py
