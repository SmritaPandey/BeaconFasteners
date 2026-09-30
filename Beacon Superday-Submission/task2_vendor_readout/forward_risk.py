"""
Looking forward: which OPEN orders are likely to arrive late?

The rest of the readout looks back. This module turns the same history into a
watch list for the next weeks, so planning can act before the part is late.

The rule is deliberately simple, so a buyer can check it by hand. A line is flagged when:
  * the vendor has already promised a date after our required date, or
  * this vendor has delivered THIS part on time less than 60% of the time
    (at least 5 past lines).
HIGH = promised more than 7 days late, or on-time history under 50%.

backtest() replays the same rule over the 8 months of history, using only
lines received before each PO was placed. The readout prints its hit rate, so
nobody has to take the rule on faith.

Promise dates come from the ERP's live confirmations and, when Task 1 has run,
from this week's acknowledgment PDFs (output/confirmation_lines_extracted.csv).
The newer of the two wins.
"""
import os

import pandas as pd

MIN_HISTORY = 5          # past lines needed before a vendor+part rate is trusted
WATCH_RATE = 0.60        # on-time rate below this -> flagged
HIGH_RATE = 0.50         # on-time rate below this -> HIGH
HIGH_PROMISE_DAYS = 7    # promised more than this many days after need -> HIGH


def _flag(promise_gap, rate, n):
    """(tier, reasons) for one line. promise_gap: promised - required in days (None if no promise)."""
    reasons, tier = [], "OK"
    if promise_gap is not None and promise_gap > 0:
        tier = "HIGH" if promise_gap > HIGH_PROMISE_DAYS else "WATCH"
        reasons.append("vendor promises %d day(s) after we need it" % promise_gap)
    if rate is not None and n >= MIN_HISTORY and rate < WATCH_RATE:
        tier = "HIGH" if (rate < HIGH_RATE or tier == "HIGH") else "WATCH"
        reasons.append("on time %.0f%% of the last %d deliveries of this part" % (100 * rate, n))
    return tier, reasons


def backtest(rec):
    """Replay the rule on history (no look-ahead). Returns the numbers the readout quotes."""
    r = rec[rec["on_time"].notna()].sort_values("po_dt")
    flags = []
    for _, row in r.iterrows():
        h = r[(r.vendor_id == row.vendor_id) & (r.part_id == row.part_id) & (r.last_rcv < row.po_dt)]
        rate = float(h.on_time.astype(float).mean()) if len(h) else None
        gap = int((row.promised - row.required).days) if pd.notna(row.promised) else None
        flags.append(_flag(gap, rate, len(h))[0] != "OK")
    r = r.assign(flag=flags)
    late = ~r.on_time.astype(bool)
    severe = r.days_late > 7
    return dict(lines=len(r), flagged=int(r.flag.sum()), late=int(late.sum()), severe=int(severe.sum()),
                severe_caught=int((r.flag & severe).sum()), late_caught=int((r.flag & late).sum()),
                precision=float(late[r.flag].mean()), base_rate=float(late.mean()),
                unflagged_late_rate=float(late[~r.flag].mean()))


def _task1_promises(path):
    """{(po, line): (promised Timestamp, doc_date Timestamp)} from Task 1's extracted lines."""
    if not path or not os.path.exists(path):
        return {}
    t = pd.read_csv(path)
    t = t[t.promised_date.notna()]
    out = {}
    for _, r in t.sort_values("doc_date").iterrows():
        out[(r.po_number, int(r.line_no))] = (pd.Timestamp(r.promised_date), pd.Timestamp(r.doc_date))
    return out


def forward_book(a, task1_lines=None, horizon_days=60):
    """Open PO lines due after the as-of date (within horizon), with a risk tier and plain-English reasons."""
    as_of = a["as_of"]
    line = a["line"].merge(a["hdr"][["po_number", "vendor_id", "po_date", "status"]], on="po_number")
    line = line[(line.status == "OPEN")].copy()
    line["required"] = pd.to_datetime(line.required_date)
    line["po_dt"] = pd.to_datetime(line.po_date)
    line = line[(line.required > as_of) & (line.required <= as_of + pd.Timedelta(days=horizon_days))]
    rec = a["rec"]
    got = rec.set_index(["po_number", "line_no"]).rec_qty.to_dict()
    conf = a["conf"][a["conf"].superseded_by.isna()]
    erp_prom = {(r.po_number, int(r.line_no)): (pd.Timestamp(r.promised_date), pd.Timestamp(r.doc_date))
                for r in conf.itertuples() if pd.notna(r.promised_date)}
    t1 = _task1_promises(task1_lines)
    hist = rec[rec.on_time.notna()]
    grp = {k: g for k, g in hist.groupby(["vendor_id", "part_id"])}
    desc = a["part"].set_index("part_id").description.to_dict()

    rows = []
    for r in line.itertuples():
        key = (r.po_number, int(r.line_no))
        open_qty = r.qty_ordered - got.get(key, 0)
        if open_qty <= 0:
            continue
        cands = [(v[1], v[0], src) for src, v in (("ERP confirmation", erp_prom.get(key)),
                                                  ("This week's acknowledgment", t1.get(key))) if v]
        promised, src = (max(cands)[1], max(cands)[2]) if cands else (None, "None on file")
        gap = int((promised - r.required).days) if promised is not None else None
        g = grp.get((r.vendor_id, r.part_id))
        n = 0 if g is None else len(g)
        rate = float(g.on_time.astype(float).mean()) if n else None
        med_late = float(g.days_late[g.days_late > 0].median()) if n and (g.days_late > 0).any() else None
        lt_actual = float(g.lt_actual.median()) if n else None
        tier, reasons = _flag(gap, rate, n)
        allowed = (r.required - r.po_dt).days
        if lt_actual and allowed < lt_actual - 3:
            reasons.append("we allowed %d days; this part usually takes %d" % (allowed, lt_actual))
        if promised is None and tier == "OK" and (r.required - as_of).days <= 21:
            reasons.append("no confirmation on file")
        rows.append(dict(
            tier=tier, vendor_id=r.vendor_id, vendor=a["vend"][r.vendor_id], po_number=r.po_number, line_no=int(r.line_no),
            part_id=r.part_id, description=desc.get(r.part_id, ""), open_qty=open_qty, open_value=open_qty * r.unit_price,
            required=r.required, days_to_due=int((r.required - as_of).days), promised=promised, promise_source=src,
            hist_n=n, hist_on_time=rate, typical_delay=med_late, why="; ".join(reasons) or "track record and promise look fine",
            action={"HIGH": "Ask the vendor for a firm date now; warn planning if the job can't move.",
                    "WATCH": "Confirm the ship date a week before it is due.",
                    "OK": "No action."}[tier]))
    df = pd.DataFrame(rows)
    if len(df):
        df["rank"] = df.tier.map({"HIGH": 0, "WATCH": 1, "OK": 2})
        df = df.sort_values(["rank", "required", "open_value"], ascending=[True, True, False]).drop(columns="rank")
    return df
