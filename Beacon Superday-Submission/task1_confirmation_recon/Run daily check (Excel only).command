#!/bin/bash
# For anyone who prefers Excel only: reads the PDFs in ./inbox, writes today's workbook to ./output and opens it.
cd "$(dirname "$0")"
[ -x .venv/bin/python ] || { python3 -m venv .venv && ./.venv/bin/pip install -q -r requirements.txt; }
DATA="${BEACON_DATA:-/data}"
INBOX="${INBOX:-$DATA/confirmations}"
./.venv/bin/python recon.py --confirmations "$INBOX" --pos "$DATA/open_pos.csv" --vendors "$DATA/vendor_master.csv" \
   --erp "$DATA/beacon_erp.db" --out output && open output/PO_Confirmation_Check_*.xlsx
