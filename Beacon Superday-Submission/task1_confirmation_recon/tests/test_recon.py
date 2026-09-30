"""
Tests for the business logic. Run from task1_confirmation_recon/:

    python3 -m unittest discover -s tests -v

Most tests use SYNTHETIC POs and documents built in code (clearly labelled
below) so each rule is tested in isolation. The last class runs the real
case data end-to-end if it is present and checks the known discrepancies.
"""
import os
import sys
import unittest
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import parsers  # noqa: E402
import recon    # noqa: E402

AS_OF = date(2026, 5, 17)


# ---------------------------------------------------------------- SYNTHETIC fixtures
def po_line(po="PO-9000000001", line=1, pn="BAR-X", qty=100, price=2.0, desc="Synthetic bar",
            vendor="V001", required=date(2026, 6, 1), po_date=date(2026, 5, 1)):
    return dict(po_number=po, po_date=po_date, vendor_id=vendor, vendor_name="Synthetic Vendor",
                line_number=line, our_pn=pn, description=desc, qty=float(qty), price=price, required=required)


def doc(po="PO-9000000001", lines=(), vendor="V001", doc_date=date(2026, 5, 5), doc_type="ACK",
        part_of=None, file="synthetic.pdf", raw="x"):
    return dict(vendor_id=vendor, template="synthetic", doc_type=doc_type, po_number=po, doc_date=doc_date,
                vendor_ref=None, part_of=part_of, supersedes_prior=doc_type == "REVISED_ACK",
                lines=list(lines), warnings=[], text_source="text", raw_text=raw, file=file)


def dline(pn="BAR-X", qty=100, price=2.0, promise=date(2026, 6, 1), line_no=1, ccy="USD", desc=None):
    return dict(line_no=line_no, vendor_pn=pn, description=desc, qty=float(qty), uom=None,
                unit_price=price, currency=ccy, promise_date=promise, promise_note=None)


def run(pos, docs, xw=None, fx=None):
    res, exc, rev, sug, ext = recon.reconcile(pos, docs, xw or {}, fx or {}, AS_OF)
    return res, exc, rev, sug


# ---------------------------------------------------------------- matching & checks
class TestReconcile(unittest.TestCase):

    def test_exact_match_is_clean(self):
        res, exc, rev, _ = run([po_line()], [doc(lines=[dline()])])
        r = res[("PO-9000000001", 1)]
        self.assertEqual(r["issues"], [])
        self.assertEqual(r["confidence"], "HIGH")
        self.assertIn("L1", r["match"])
        self.assertEqual((exc, rev), ([], []))

    def test_missing_line_is_possible_drop(self):
        pos = [po_line(line=1, pn="A"), po_line(line=2, pn="B")]
        res, *_ = run(pos, [doc(lines=[dline(pn="A")])])
        self.assertIn("DROPPED", res[("PO-9000000001", 2)]["issues"])
        self.assertEqual(recon.ISSUE["DROPPED"][1], "POSSIBLE_DROPPED_LINE")
        self.assertEqual(res[("PO-9000000001", 1)]["issues"], [])

    def test_no_document_at_all(self):
        res, *_ = run([po_line()], [])
        self.assertIn("NO_CONF", res[("PO-9000000001", 1)]["issues"])

    def test_quantity_short_and_over(self):
        res, *_ = run([po_line()], [doc(lines=[dline(qty=95)])])
        r = res[("PO-9000000001", 1)]
        self.assertIn("QTY_SHORT", r["issues"])
        self.assertEqual(r["qty_var"], -5)
        res, *_ = run([po_line()], [doc(lines=[dline(qty=120)])])
        self.assertIn("QTY_OVER", res[("PO-9000000001", 1)]["issues"])

    def test_price_increase_flagged_with_dollar_impact(self):
        res, *_ = run([po_line(qty=1000, price=3.92)], [doc(lines=[dline(qty=1000, price=4.0102)])])
        r = res[("PO-9000000001", 1)]
        self.assertIn("PRICE_UP", r["issues"])
        self.assertAlmostEqual(r["impact"], 90.2, places=1)

    def test_eur_price_within_fx_tolerance_not_flagged(self):
        fx = {("2026-05", "EUR"): 1.087}
        res, *_ = run([po_line(price=1240.0)], [doc(lines=[dline(price=1140.80, ccy="EUR")])], fx=fx)
        r = res[("PO-9000000001", 1)]
        self.assertNotIn("PRICE_UP", r["issues"])
        self.assertAlmostEqual(r["conf_price_usd"], 1240.05, places=1)

    def test_eur_price_outside_tolerance_flagged(self):
        fx = {("2026-05", "EUR"): 1.087}
        res, *_ = run([po_line(price=1240.0)], [doc(lines=[dline(price=1200.0, ccy="EUR")])], fx=fx)
        self.assertIn("PRICE_UP", res[("PO-9000000001", 1)]["issues"])

    def test_late_promise(self):
        res, *_ = run([po_line(required=date(2026, 6, 1))], [doc(lines=[dline(promise=date(2026, 6, 15))])])
        r = res[("PO-9000000001", 1)]
        self.assertIn("LATE", r["issues"])
        self.assertEqual(r["days_late"], 14)

    def test_early_promise_is_fine(self):
        res, *_ = run([po_line(required=date(2026, 6, 1))], [doc(lines=[dline(promise=date(2026, 5, 20))])])
        self.assertNotIn("LATE", res[("PO-9000000001", 1)]["issues"])

    def test_doc_dated_before_po_is_inconsistent(self):
        res, *_ = run([po_line(po_date=date(2026, 5, 10))], [doc(doc_date=date(2026, 5, 2), lines=[dline()])])
        self.assertIn("DATE_ODD", res[("PO-9000000001", 1)]["issues"])

    # ------------------------------------------------ part-number resolution
    def test_vendor_pn_via_approved_crosswalk(self):
        xw = {("V002", "APH-441"): dict(vendor_id="V002", vendor_pn="APH-441", beacon_pn="CHB-9472-3",
                                        review_status="APPROVED (historical)")}
        res, _, rev, sug = run([po_line(pn="CHB-9472-3", vendor="V002")],
                               [doc(vendor="V002", lines=[dline(pn="APH-441")])], xw=xw)
        r = res[("PO-9000000001", 1)]
        self.assertEqual(r["confidence"], "HIGH")
        self.assertIn("L2", r["match"])
        self.assertEqual((rev, sug), ([], {}))

    def test_crosswalk_is_vendor_scoped(self):
        # Same vendor PN, different vendor -> must NOT use the other vendor's mapping
        xw = {("V002", "K-1050"): dict(vendor_id="V002", vendor_pn="K-1050", beacon_pn="CHB-5520",
                                       review_status="APPROVED (historical)")}
        res, *_ = run([po_line(pn="TOOL-PUNCH-STD", vendor="V005")],
                      [doc(vendor="V005", lines=[dline(pn="K-1050")])], xw=xw)
        self.assertNotIn("L2", res[("PO-9000000001", 1)]["match"])

    def test_pending_crosswalk_entry_not_auto_used(self):
        xw = {("V002", "NEW-1"): dict(vendor_id="V002", vendor_pn="NEW-1", beacon_pn="BAR-X",
                                      review_status="PENDING REVIEW")}
        res, _, rev, _ = run([po_line(vendor="V002")], [doc(vendor="V002", lines=[dline(pn="NEW-1")])], xw=xw)
        self.assertNotIn("L2", res[("PO-9000000001", 1)]["match"])
        self.assertEqual(len(rev), 1)

    def test_unknown_vendor_pn_inferred_goes_to_review(self):
        res, _, rev, sug = run([po_line(pn="CHB-9472-4", vendor="V002", price=0.86, qty=25000)],
                               [doc(vendor="V002", lines=[dline(pn="APH-441-OS", price=0.86, qty=25000)])])
        r = res[("PO-9000000001", 1)]
        self.assertEqual(r["confidence"], "MEDIUM")
        self.assertIn("PN_INFERRED", r["issues"])
        self.assertEqual(rev[0]["candidate"], "CHB-9472-4")
        self.assertEqual(sug[("V002", "APH-441-OS")]["review_status"], "PENDING REVIEW")

    def test_ambiguous_part_mapping_is_not_guessed(self):
        # Two PO lines with identical qty and price; vendor uses an unknown PN and no line number
        pos = [po_line(line=1, pn="A"), po_line(line=2, pn="B")]
        dl = dline(pn="ZZZ-UNKNOWN", line_no=None)
        res, exc, _, _ = run(pos, [doc(lines=[dl])])
        self.assertEqual(len(exc), 1)
        self.assertEqual(exc[0]["code"], "UNMATCHED_DOC_LINE")
        self.assertIn("equally likely", exc[0]["detail"])
        # ...and both PO lines then show as not confirmed rather than one silently "matched"
        self.assertIn("DROPPED", res[("PO-9000000001", 1)]["issues"])
        self.assertIn("DROPPED", res[("PO-9000000001", 2)]["issues"])

    def test_description_match_when_vendor_has_no_pn(self):
        pos = [po_line(pn="PLAT-PASV", desc="Passivation per AMS 2700, per piece", price=0.08)]
        res, *_ = run(pos, [doc(lines=[dline(pn="PASV", desc="Passivation per AMS 2700, per piece", price=None)])])
        r = res[("PO-9000000001", 1)]
        self.assertIn("L3", r["match"])
        self.assertIn("NO_PRICE", r["issues"])

    # ------------------------------------------------ multiple documents
    def test_revised_confirmation_supersedes_original(self):
        d1 = doc(doc_date=date(2026, 5, 6), lines=[dline(qty=90)], file="orig.pdf", raw="a")
        d2 = doc(doc_date=date(2026, 5, 11), doc_type="REVISED_ACK", lines=[dline(qty=100)], file="rev.pdf", raw="b")
        res, *_ = run([po_line()], [d1, d2])
        r = res[("PO-9000000001", 1)]
        self.assertEqual(r["conf_qty"], 100)
        self.assertNotIn("QTY_SHORT", r["issues"])
        self.assertIn("SUPERSEDED", r["issues"])
        self.assertTrue(d1["status"].startswith("Superseded"))

    def test_split_parts_are_summed_not_superseded(self):
        d1 = doc(part_of=(1, 2), lines=[dline(qty=60, promise=date(2026, 5, 28))], file="p1.pdf", raw="a")
        d2 = doc(part_of=(2, 2), lines=[dline(qty=40, promise=date(2026, 6, 5))], file="p2.pdf", raw="b")
        res, *_ = run([po_line(required=date(2026, 6, 1))], [d1, d2])
        r = res[("PO-9000000001", 1)]
        self.assertEqual(r["conf_qty"], 100)
        self.assertIn("SPLIT", r["issues"])
        self.assertIn("LATE", r["issues"])      # second delivery is after required
        self.assertEqual(r["promise"], date(2026, 6, 5))

    def test_exact_duplicate_pdf_ignored(self):
        d1 = doc(lines=[dline(qty=50)], file="a.pdf", raw="same content")
        d2 = doc(lines=[dline(qty=50)], file="a_copy.pdf", raw="same   content")
        kept = recon.mark_duplicates([d1, d2])
        self.assertEqual([d["file"] for d in kept], ["a.pdf"])
        self.assertEqual(d2["duplicate_of"], "a.pdf")

    def test_confirmation_for_unknown_po(self):
        res, exc, *_ = run([po_line()], [doc(lines=[dline()]), doc(po="PO-9999999999", lines=[dline()], file="rogue.pdf", raw="r")])
        self.assertEqual([e["code"] for e in exc], ["PO_NOT_OPEN"])

    def test_ack_without_detail(self):
        d = doc(doc_type="ACK_NO_DETAIL", lines=[])
        d["warnings"] = ["no qty or date"]
        res, *_ = run([po_line()], [d])
        self.assertIn("NO_DETAIL", res[("PO-9000000001", 1)]["issues"])

    def test_validation_flags_bad_values(self):
        d = doc(lines=[dline(qty=0, price=-1, promise=None)])
        recon.validate_doc(d)
        self.assertEqual(len(d["warnings"]), 3)


# ---------------------------------------------------------------- normalization / parsing
class TestNormalization(unittest.TestCase):

    def test_number_formats(self):
        self.assertEqual(parsers._num("1,500"), 1500)
        self.assertEqual(parsers._num("$4.0102"), 4.0102)
        self.assertEqual(parsers._num("€1140.8000"), 1140.8)
        self.assertEqual(parsers._num("1.140,80"), 1140.8)   # German
        self.assertEqual(parsers._num("16,74"), 16.74)       # German decimal comma

    def test_german_dates_and_calendar_weeks(self):
        self.assertEqual(parsers._parse_promise_de("08.06.2026")[0], date(2026, 6, 8))
        d, note = parsers._parse_promise_de("KW 20-22 / 2026")
        self.assertEqual(d, date(2026, 5, 31))              # Sunday of ISO week 22
        self.assertIn("05/11", note)                         # Monday of ISO week 20
        self.assertEqual(parsers._parse_promise_de("KW 21 / 2026")[0], date(2026, 5, 24))

    def test_quickship_email_revised_and_invoice(self):
        # SYNTHETIC email text in QuickShip's style
        t = ("From: orders@quickship-ind.com\nDate: 05/11/2026\nSubject: Order confirmation - PO PO-4500050030 (revised)\n"
             "Please disregard our earlier confirmation.\nMISC-SPR-001 qty 1,500 @ $0.1800 ea ship 05/21/2026\n")
        d = parsers.parse_quickship(t)
        self.assertEqual(d["doc_type"], "REVISED_ACK")
        self.assertEqual(d["lines"][0]["qty"], 1500)
        t2 = t.replace("Order confirmation", "Invoice").replace("(revised)", "").replace("disregard", "see")
        self.assertEqual(parsers.parse_quickship(t2)["doc_type"], "INVOICE")

    def test_ostmark_line(self):
        # SYNTHETIC text in Ostmark's layout
        t = ("Ostmark Werkzeug GmbH\nIhre Bestellung / Your PO: PO-4500050027\nDatum / Date: 01.05.2026\n"
             "Auftrag-Nr.: OST-1\n1         OST-CAR-A-100              (siehe Bestellung)      50        €16.7440    KW 20-22 / 2026\n")
        d = parsers.parse_ostmark(t)
        self.assertEqual(d["doc_date"], date(2026, 5, 1))
        ln = d["lines"][0]
        self.assertEqual((ln["vendor_pn"], ln["qty"], ln["unit_price"], ln["currency"]), ("OST-CAR-A-100", 50, 16.744, "EUR"))
        self.assertIsNotNone(ln["promise_note"])

    def test_unrecognised_layout_is_flagged_not_guessed(self):
        self.assertEqual(parsers.TEMPLATES[0][0].search("Some Other Vendor Inc"), None)


# ---------------------------------------------------------------- real data, end to end
DATA = os.environ.get("BEACON_DATA", os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))),
                                                  "Beacon Superday-Candidate", "data"))


@unittest.skipUnless(os.path.isdir(os.path.join(DATA, "confirmations")), "case data not found (set BEACON_DATA)")
class TestEndToEnd(unittest.TestCase):
    """Runs on the real case files. Expected results were verified by reading each PDF by hand."""

    @classmethod
    def setUpClass(cls):
        import tempfile
        out = recon.run(os.path.join(DATA, "confirmations"), os.path.join(DATA, "open_pos.csv"),
                        os.path.join(DATA, "vendor_master.csv"), os.path.join(DATA, "beacon_erp.db"),
                        None, "2026-05-17", tempfile.mkdtemp())
        cls.stats, cls.res, cls.exc, cls.sug = out["stats"], out["results"], out["exceptions"], out["suggestions"]

    def issues(self, po, line):
        return self.res[("PO-45000500%02d" % po, line)]["issues"]

    def test_every_pdf_read_and_every_line_accounted(self):
        self.assertEqual(self.stats["PDFs processed"], 35)
        self.assertEqual(self.stats["  unreadable (manual)"], 0)
        self.assertEqual(self.stats["CHECK every PDF has a status"], "PASS")
        self.assertEqual(self.stats["CHECK every open PO line is accounted for"], "PASS")

    def test_known_discrepancies(self):
        self.assertIn("DROPPED", self.issues(1, 2))          # Apex dropped line 2
        self.assertIn("QTY_SHORT", self.issues(7, 1))        # 1,425 of 1,500
        self.assertIn("PRICE_UP", self.issues(2, 1))         # $4.0102 vs $3.92
        self.assertIn("LATE", self.issues(22, 1))            # Continental +21 days (OCR doc)
        self.assertIn("LATE", self.issues(12, 1))            # Heritage +14 days
        self.assertIn("LATE", self.issues(16, 1))            # Liberty split, 2nd part late
        self.assertIn("NO_DETAIL", self.issues(19, 1))       # Liberty ack w/o schedule
        self.assertIn("INVOICE_ONLY", self.issues(32, 1))    # QuickShip invoice
        self.assertIn("VAGUE_DATE", self.issues(27, 1))      # Ostmark KW range

    def test_revised_email_used_not_original(self):
        r = self.res[("PO-4500050030", 1)]
        self.assertEqual(r["conf_qty"], 500)
        self.assertNotIn("QTY_SHORT", r["issues"])

    def test_rogue_po_flagged(self):
        self.assertEqual([e["po"] for e in self.exc if e["code"] == "PO_NOT_OPEN"], ["PO-4500060619"])

    def test_ostmark_eur_prices_not_false_alarms(self):
        for key in [(26, 1), (27, 1), (27, 2), (28, 1), (29, 1), (29, 2)]:
            self.assertNotIn("PRICE_UP", self.issues(*key))

    def test_renumbered_vendor_pns_go_to_review(self):
        self.assertEqual(set(self.sug), {("V002", "APH-441-OS"), ("V005", "OST-CAR-A-100")})

    def test_clean_lines_stay_clean(self):
        self.assertEqual(self.issues(9, 1), [])     # Heritage APH-441 via crosswalk
        self.assertEqual(self.issues(6, 1), [])     # Apex exact


if __name__ == "__main__":
    unittest.main()
