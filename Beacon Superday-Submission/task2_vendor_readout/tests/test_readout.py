"""
Tests for the vendor readout. Run from task2_vendor_readout/:

    BEACON_DATA=/data python3 -m unittest discover -s tests -v

The rule tests are self-contained. The ERP tests pin the headline numbers the
readout and the deck quote, so a change in the cleaning logic can't silently
move them.
"""
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import forward_risk as fr  # noqa: E402

DATA = os.environ.get("BEACON_DATA", os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))),
                                                  "Beacon Superday-Candidate", "data"))
ERP = os.path.join(DATA, "beacon_erp.db")


class TestFlagRule(unittest.TestCase):
    def test_clean_line_is_ok(self):
        self.assertEqual(fr._flag(0, 0.95, 40)[0], "OK")

    def test_promise_after_need(self):
        self.assertEqual(fr._flag(3, 0.95, 40)[0], "WATCH")
        self.assertEqual(fr._flag(14, 0.95, 40)[0], "HIGH")

    def test_poor_track_record(self):
        self.assertEqual(fr._flag(None, 0.55, 20)[0], "WATCH")
        self.assertEqual(fr._flag(None, 0.40, 20)[0], "HIGH")

    def test_too_little_history_is_not_judged(self):
        self.assertEqual(fr._flag(None, 0.0, 3)[0], "OK")


@unittest.skipUnless(os.path.exists(ERP), "case data not found (set BEACON_DATA)")
class TestOnCaseData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from analysis import build
        cls.a = build(ERP)

    def test_received_value_after_cleaning(self):
        self.assertAlmostEqual(float(self.a["monthly"]["Total"].sum()), 28187145.87, places=2)

    def test_call_first_is_continental(self):
        s = self.a["score"]
        self.assertEqual(s.on_time.idxmin(), "V004")
        self.assertEqual(int(s.loc["V004"].severe_lines), 61)
        self.assertEqual(int(s.severe_lines.sum()), 63)

    def test_backtest_numbers_quoted_in_readout(self):
        bt = fr.backtest(self.a["rec"])
        self.assertEqual((bt["severe_caught"], bt["severe"]), (59, 63))
        self.assertGreater(bt["precision"], bt["base_rate"])

    def test_forward_book(self):
        fw = fr.forward_book(self.a)          # ERP promises only
        self.assertEqual(len(fw), 80)
        self.assertAlmostEqual(fw.open_value.sum(), 3320355.0, places=0)
        self.assertTrue((fw[fw.tier == "HIGH"].vendor_id == "V004").any())


if __name__ == "__main__":
    unittest.main()
