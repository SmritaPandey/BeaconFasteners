#!/usr/bin/env bash
# One command for graders: checks prerequisites, installs into a local venv, runs both
# tasks on the case data, generates the vendor forms, and runs the test suites.
#
#   ./run_all.sh              # data in /data (the case default)
#   ./run_all.sh path/to/data # data somewhere else
#
# Outputs land in ./output. Nothing outside this folder is touched.
set -euo pipefail

DATA="${1:-/data}"
HERE="$(cd "$(dirname "$0")" && pwd)"
OUT="$HERE/output"
say() { printf '\n\033[1m== %s\033[0m\n' "$*"; }

say "Checking prerequisites"
command -v python3 >/dev/null || { echo "python3 not found - install Python 3.9+"; exit 1; }
python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' || { echo "Python 3.9+ required"; exit 1; }
for f in confirmations open_pos.csv vendor_master.csv beacon_erp.db; do
  [ -e "$DATA/$f" ] || { echo "Missing $DATA/$f - pass the data folder: ./run_all.sh /path/to/data"; exit 1; }
done
if command -v tesseract >/dev/null; then
  echo "tesseract: $(tesseract --version 2>&1 | head -1)"
else
  echo "WARNING: tesseract not installed - the 4 scanned Continental PDFs will be listed as 'could not read'."
  echo "         Install it for full results: brew install tesseract  |  sudo apt install tesseract-ocr"
fi
echo "data: $DATA"

say "Installing Python packages into .venv (first run only takes a minute)"
[ -x "$HERE/.venv/bin/python" ] || python3 -m venv "$HERE/.venv"
PY="$HERE/.venv/bin/python"
"$PY" -m pip install -q --upgrade pip >/dev/null
"$PY" -m pip install -q -r "$HERE/task1_confirmation_recon/requirements.txt"

mkdir -p "$OUT"
say "Task 1 - PO confirmation check (Lisa's workbook)"
cd "$HERE/task1_confirmation_recon"
"$PY" recon.py --confirmations "$DATA/confirmations" --pos "$DATA/open_pos.csv" \
               --vendors "$DATA/vendor_master.csv" --erp "$DATA/beacon_erp.db" --out "$OUT" 2>&1 | grep -v "fitz"

say "Vendor forms - acknowledgment form per open PO, supplier pack per vendor"
"$PY" vendor_forms.py ack     --pos "$DATA/open_pos.csv" --vendors "$DATA/vendor_master.csv" --erp "$DATA/beacon_erp.db" \
                              --out "$OUT/vendor_forms/acknowledgment_forms" | tail -1
"$PY" vendor_forms.py onboard --pos "$DATA/open_pos.csv" --vendors "$DATA/vendor_master.csv" --erp "$DATA/beacon_erp.db" \
                              --out "$OUT/vendor_forms/supplier_packs" | tail -1

say "Task 2 - vendor performance readout (plant manager)"
cd "$HERE/task2_vendor_readout"
"$PY" build_readout.py --erp "$DATA/beacon_erp.db" --out "$OUT"

say "Tests"
cd "$HERE/task1_confirmation_recon" && BEACON_DATA="$DATA" "$PY" -m unittest discover -s tests 2>&1 | grep -v fitz | tail -3
cd "$HERE/task2_vendor_readout"     && BEACON_DATA="$DATA" "$PY" -m unittest discover -s tests 2>&1 | tail -3

say "Done. Open these:"
ls -1 "$OUT"/PO_Confirmation_Check_*.xlsx "$OUT"/Vendor_Readout.xlsx "$OUT"/vendor_readout.html 2>/dev/null
echo "Optional browser app for Lisa: cd task1_confirmation_recon && ../.venv/bin/streamlit run desk_app.py"
