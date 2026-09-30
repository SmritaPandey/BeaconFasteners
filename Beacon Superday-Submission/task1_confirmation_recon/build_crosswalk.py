"""
Seed pn_crosswalk.csv (vendor part number -> Beacon part number) from ERP history.

The ERP `confirmation` table stores the vendor's PN but leaves part_id empty for
vendors that use their own numbers. Joining back to po_line on (po_number,
line_no) recovers the Beacon part each vendor PN was actually confirmed against.

Key is (vendor_id, vendor_pn), NOT vendor_pn alone: 'K-1050' is one part at
Heritage and a different part at Ostmark.

This file is the seed of a future part-number master: every row carries where
the mapping came from, how confident we are, and whether a human has reviewed it.

Usage:  python3 build_crosswalk.py --erp /path/to/beacon_erp.db
"""
import argparse
import csv
import os
import sqlite3

COLUMNS = ["vendor_id", "vendor_name", "vendor_pn", "beacon_pn", "description",
           "mapping_source", "mapping_confidence", "times_seen", "first_seen", "last_seen",
           "review_status"]

SQL = """
SELECT h.vendor_id, v.vendor_name, c.vendor_pn, l.part_id, p.description,
       COUNT(*), MIN(c.doc_date), MAX(c.doc_date)
FROM confirmation c
JOIN po_line l   ON l.po_number = c.po_number AND l.line_no = c.line_no
JOIN po_header h ON h.po_number = c.po_number
LEFT JOIN vendor_master v ON v.vendor_id = h.vendor_id
LEFT JOIN part_master p   ON p.part_id = l.part_id
WHERE c.vendor_pn IS NOT NULL AND c.vendor_pn <> ''
GROUP BY 1, 2, 3, 4, 5
ORDER BY 1, 3
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--erp", required=True)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "pn_crosswalk.csv"))
    a = ap.parse_args()
    rows = sqlite3.connect(a.erp).execute(SQL).fetchall()
    # A vendor PN that maps to more than one Beacon part is ambiguous -> needs review
    seen = {}
    for r in rows:
        seen.setdefault((r[0], r[2]), set()).add(r[3])
    with open(a.out, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
        for vid, vname, vpn, bpn, desc, n, first, last in rows:
            ambiguous = len(seen[(vid, vpn)]) > 1
            w.writerow([vid, vname, vpn, bpn, desc, "ERP confirmation history (po_number+line_no join)",
                        "LOW" if ambiguous else "HIGH", n, first, last,
                        "NEEDS REVIEW - ambiguous" if ambiguous else "APPROVED (historical)"])
    print("wrote %d mappings -> %s" % (len(rows), a.out))


if __name__ == "__main__":
    main()
