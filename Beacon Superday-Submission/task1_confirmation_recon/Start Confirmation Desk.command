#!/bin/bash
# Double-click to open the Confirmation Desk in your browser (Mac).
cd "$(dirname "$0")"
if [ ! -x .venv/bin/streamlit ]; then
  echo "First run: setting up (takes a minute)..."
  python3 -m venv .venv && ./.venv/bin/pip install -q -r requirements.txt || { echo "Setup failed - see README"; read; exit 1; }
fi
command -v tesseract >/dev/null || echo "Note: Tesseract is not installed - scanned PDFs will go to 'Needs your OK' for typing in. (brew install tesseract)"
./.venv/bin/streamlit run desk_app.py
