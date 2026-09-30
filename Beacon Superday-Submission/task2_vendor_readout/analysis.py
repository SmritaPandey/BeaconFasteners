"""
Vendor performance analysis on the Beacon ERP extract (beacon_erp.db).

Every metric here states its numerator, denominator, period and exclusions in
METRIC_DEFINITIONS so the readout can print them next to the numbers.

Data traps handled (all verified independently - see DATA_QUALITY):
  * receipt_txn.txn_date is MM/DD/YYYY text (everything else is ISO)
  * 50 'RV' reversals (packing slip REV-KEYING-ERR) each cancel a duplicate 'R'
    keyed the day after the real one -> drop BOTH rows for shipment-level metrics
  * 1 duplicate 'R' was never reversed -> detected by rule, excluded
  * po_line.qty_received is stale on 169 lines -> never used; receipts are truth
  * confirmation rows: 30 superseded -> use superseded_by IS NULL only
  * confirmation feed stops at doc_date 2026-04-03 -> confirmation metrics only
    for POs dated on/before the last confirmed PO date
  * Ostmark (EUR vendor): po_line.unit_price is already USD (its EUR confirmations
    are exactly 0.92 x PO price) -> no FX on received value
"""
import sqlite3
from datetime import date

import numpy as np
import pandas as pd

LATE_GRACE_DAYS = 7          # "severely late" threshold (see README for why 7)
EARLY_PERIOD = ("2025-09", "2025-12")
LATE_PERIOD = ("2026-01", "2026-05")


def load(db_path):
    con = sqlite3.connect("file:%s?mode=ro" % db_path, uri=True)
    q = lambda s: pd.read_sql_query(s, con)
    t = dict(
        vendor=q("select * from vendor_master"),
        part=q("select * from part_master"),
        hdr=q("select * from po_header"),
        line=q("select * from po_line"),
        rcv=q("select * from receipt_txn"),
        conf=q("select * from confirmation"),
        fx=q("select * from fx_rate"),
        hold=q("select * from qc_hold"),
    )
    con.close()
    return t


def clean_receipts(rcv, line):
    """Returns (valid shipments, net-quantity rows, list of data-quality notes)."""
    r = rcv.copy()
    r["rdate"] = pd.to_datetime(r["txn_date"], format="%m/%d/%Y")
    notes = []

    # 1. Keying-error pairs: each RV reverses the R immediately before it
    rv = r[r.action_type == "RV"]
    prev = r.set_index("txn_id").reindex(rv.txn_id - 1)
    pair_ok = ((prev["action_type"].values == "R") & (prev["qty"].values == -rv["qty"].values) &
               (prev["po_number"].values == rv["po_number"].values) & (prev["line_no"].values == rv["line_no"].values))
    reversed_ids = set((rv.txn_id - 1)[pair_ok]) | set(rv.txn_id)
    notes.append(("Keyed-in-error receipts reversed (RV + the R it cancels)", int(pair_ok.sum()),
                  "Dropped both rows for shipment dates/counts; nets to zero for value."))

    ship = r[(r.action_type == "R") & ~r.txn_id.isin(reversed_ids)].copy()

    # 2. Unreversed duplicate: same line+qty as an earlier R within 7 days AND
    #    it pushes the line above ordered qty
    ship = ship.sort_values(["po_number", "line_no", "rdate", "txn_id"])
    ordered = line.set_index(["po_number", "line_no"])["qty_ordered"]
    dup_ids = []
    for (po, ln), g in ship.groupby(["po_number", "line_no"]):
        cum = 0
        seen = []
        for _, row in g.iterrows():
            is_dup = any(q == row.qty and 0 <= (row.rdate - d).days <= 7 for q, d in seen)
            if is_dup and cum + row.qty > ordered[(po, ln)]:
                dup_ids.append(int(row.txn_id))
                continue
            cum += row.qty
            seen.append((row.qty, row.rdate))
    ship = ship[~ship.txn_id.isin(dup_ids)]
    notes.append(("Unreversed duplicate receipts (same line+qty within 7 days, pushes line over ordered)",
                  len(dup_ids), "Excluded: txn %s" % ", ".join(map(str, dup_ids))))
    return ship, dup_ids, notes


def build(db_path, as_of=None):
    t = load(db_path)
    hdr, line, conf = t["hdr"], t["line"], t["conf"]
    vend = t["vendor"].set_index("vendor_id")["vendor_name"]
    short = {vid: n.split()[0].replace("&", "") for vid, n in vend.items()}

    ship, dup_ids, dq = clean_receipts(t["rcv"], line)
    as_of = pd.Timestamp(as_of) if as_of else ship["rdate"].max()

    ln = line.merge(hdr[["po_number", "vendor_id", "po_date", "status", "buyer"]], on="po_number")
    ln["required"] = pd.to_datetime(ln["required_date"])
    ln["po_dt"] = pd.to_datetime(ln["po_date"])
    ln = ln[ln.status != "CANCELLED"].copy()

    agg = ship.groupby(["po_number", "line_no"]).agg(rec_qty=("qty", "sum"), first_rcv=("rdate", "min"),
                                                     last_rcv=("rdate", "max"), shipments=("qty", "size")).reset_index()
    ln = ln.merge(agg, on=["po_number", "line_no"], how="left")
    ln["rec_qty"] = ln["rec_qty"].fillna(0)
    ln["value_ordered"] = ln.qty_ordered * ln.unit_price

    live = conf[conf.superseded_by.isna()][["po_number", "line_no", "confirmed_price", "currency",
                                            "promised_date", "doc_date", "vendor_pn"]]
    ln = ln.merge(live, on=["po_number", "line_no"], how="left")
    ln["promised"] = pd.to_datetime(ln["promised_date"])
    conf_cutoff = pd.to_datetime(hdr.merge(conf[["po_number"]].drop_duplicates(), on="po_number")["po_date"]).max()

    # ---------------- received value by month (net of reversals, dup excluded) --------
    r_all = t["rcv"].copy()
    r_all = r_all[~r_all.txn_id.isin(dup_ids)]
    r_all["rdate"] = pd.to_datetime(r_all["txn_date"], format="%m/%d/%Y")
    r_all = r_all.merge(line[["po_number", "line_no", "unit_price", "part_id"]], on=["po_number", "line_no"]) \
                 .merge(hdr[["po_number", "vendor_id"]], on="po_number")
    r_all["value"] = r_all.qty * r_all.unit_price
    r_all["month"] = r_all.rdate.dt.strftime("%Y-%m")
    monthly = r_all.pivot_table(index="month", columns="vendor_id", values="value", aggfunc="sum", fill_value=0)
    monthly["Total"] = monthly.sum(axis=1)
    ship_m = ship.merge(hdr[["po_number", "vendor_id"]], on="po_number")
    ship_m["month"] = ship_m.rdate.dt.strftime("%Y-%m")
    monthly_counts = ship_m.groupby("month").agg(receipts=("txn_id", "size"), vendors=("vendor_id", "nunique"))
    first_day, last_day = r_all.rdate.min(), r_all.rdate.max()

    # naive comparisons (for the data-quality slide)
    naive_R_only = (r_all[r_all.action_type == "R"].value.sum() +
                    (t["rcv"][t["rcv"].txn_id.isin(dup_ids)].merge(line, on=["po_number", "line_no"])
                     .eval("qty*unit_price").sum()))
    naive_qty_received = (line.qty_received * line.unit_price).sum()
    fxm = t["fx"][t["fx"].currency == "EUR"].set_index("month")["rate_to_usd"]
    v5 = r_all[r_all.vendor_id == "V005"]
    v5_if_eur = (v5.value * v5.month.map(fxm)).sum() - v5.value.sum()

    # ---------------- delivery (due lines) ----------------
    due = ln[ln.required <= as_of].copy()
    rec = due[due.rec_qty > 0].copy()
    rec["days_late"] = (rec.last_rcv - rec.required).dt.days
    rec["on_time"] = rec.days_late <= 0
    rec["on_time_first"] = (rec.first_rcv - rec.required).dt.days <= 0
    rec["full"] = rec.rec_qty >= rec.qty_ordered
    rec["otif"] = rec.on_time & rec.full
    rec["severe"] = rec.days_late > LATE_GRACE_DAYS
    rec["value_rcv"] = np.minimum(rec.rec_qty, rec.qty_ordered) * rec.unit_price
    rec["req_month"] = rec.required.dt.strftime("%Y-%m")
    rec["lt_allowed"] = (rec.required - rec.po_dt).dt.days
    rec["lt_promised"] = (rec.promised - rec.po_dt).dt.days
    rec["lt_actual"] = (rec.last_rcv - rec.po_dt).dt.days
    wc = rec[rec.promised.notna()].copy()
    wc["promise_gap"] = (wc.promised - wc.required).dt.days
    wc["on_time_vs_promise"] = wc.last_rcv <= wc.promised

    # ---------------- open / past due ----------------
    open_ln = ln[(ln.status == "OPEN") & (ln.rec_qty < ln.qty_ordered)].copy()
    open_ln["balance_qty"] = open_ln.qty_ordered - open_ln.rec_qty
    open_ln["balance_value"] = open_ln.balance_qty * open_ln.unit_price
    past_due = open_ln[open_ln.required < as_of].copy()
    past_due["days_past_due"] = (as_of - past_due.required).dt.days
    past_due["kind"] = np.where(past_due.rec_qty > 0, "Short-shipped balance left open", "Nothing received")

    # ---------------- commercial ----------------
    c = ln[ln.confirmed_price.notna() & (ln.currency == "USD")].copy()
    c["price_var"] = c.confirmed_price - c.unit_price
    above = c[c.price_var > 0.00005].copy()
    above["exposure_ordered"] = above.price_var * above.qty_ordered
    above["exposure_received"] = above.price_var * np.minimum(above.rec_qty, above.qty_ordered)
    above["conf_month"] = above.doc_date.str[:7]
    eur = ln[ln.currency == "EUR"]
    eur_ratio = (eur.confirmed_price / eur.unit_price).round(4).unique().tolist()

    # ---------------- confirmation behaviour (window with a live feed) ----------------
    win = ln[ln.po_dt <= conf_cutoff]
    sup = conf[conf.superseded_by.notna()].merge(conf[["conf_id", "confirmed_qty", "promised_date"]]
                                                 .rename(columns={"conf_id": "superseded_by", "confirmed_qty": "new_qty",
                                                                  "promised_date": "new_prom"}), on="superseded_by") \
                                          .merge(hdr[["po_number", "vendor_id"]], on="po_number")
    sup["kind"] = np.where(sup.new_qty != sup.confirmed_qty, "qty raised", "date pushed out")

    # ---------------- holds (weak signal) ----------------
    holds = t["hold"].merge(hdr[["po_number", "vendor_id"]], on="po_number")

    # ---------------- per-vendor scorecard ----------------
    rows = []
    total_val = monthly["Total"].sum()
    for vid in vend.index:
        R = rec[rec.vendor_id == vid]
        C = wc[wc.vendor_id == vid]
        W = win[win.vendor_id == vid]
        PD = past_due[past_due.vendor_id == vid]
        A = above[above.vendor_id == vid]
        H = holds[holds.vendor_id == vid]
        early = R[(R.req_month >= EARLY_PERIOD[0]) & (R.req_month <= EARLY_PERIOD[1])]
        late = R[(R.req_month >= LATE_PERIOD[0]) & (R.req_month <= LATE_PERIOD[1])]
        received_pos = ship.merge(hdr[["po_number", "vendor_id"]], on="po_number").query("vendor_id == @vid").po_number.nunique()
        rows.append(dict(
            vendor_id=vid, vendor=vend[vid], short=short[vid],
            received_value=monthly[vid].sum() if vid in monthly else 0.0,
            share=(monthly[vid].sum() / total_val) if vid in monthly else 0.0,
            pos=hdr[(hdr.vendor_id == vid) & (hdr.status != "CANCELLED")].po_number.nunique(),
            due_lines=len(R),
            on_time=R.on_time.mean(), on_time_n=int(R.on_time.sum()),
            late_lines=int((~R.on_time).sum()),
            on_time_first=R.on_time_first.mean(),
            otif=R.otif.mean(),
            severe_lines=int(R.severe.sum()), severe_pct=R.severe.mean(),
            severe_value=R[R.severe].value_rcv.sum(),
            late_value=R[~R.on_time].value_rcv.sum(),
            avg_days_late_when_late=R[~R.on_time].days_late.mean(),
            median_days_late_when_late=R[~R.on_time].days_late.median(),
            max_days_late=R.days_late.max(),
            p90_days_late=R.days_late.quantile(0.9),
            on_time_with_grace=(R.days_late <= LATE_GRACE_DAYS).mean(),
            lines_w_promise=len(C),
            pct_promised_late=(C.promise_gap > 0).mean(), avg_promise_gap=C.promise_gap.mean(),
            on_time_vs_promise=C.on_time_vs_promise.mean(),
            lt_allowed=R.lt_allowed.median(), lt_promised=R.lt_promised.median(), lt_actual=R.lt_actual.median(),
            multi_ship_pct=(R.shipments > 1).mean(),
            conf_coverage=W.promised_date.notna().mean(), conf_window_lines=len(W),
            revisions=int((sup.vendor_id == vid).sum()),
            pushouts=int(((sup.vendor_id == vid) & (sup.kind == "date pushed out")).sum()),
            past_due_lines=len(PD), past_due_value=PD.balance_value.sum(),
            past_due_nothing=int((PD.rec_qty == 0).sum()),
            short_open_lines=int((PD.rec_qty > 0).sum()),
            price_above_lines=len(A), price_exposure=A.exposure_ordered.sum(),
            holds=len(H), holds_open=int((H.released == "N").sum()), holds_per_100po=100.0 * len(H) / max(received_pos, 1),
            ot_early=early.on_time.mean(), ot_early_n=len(early),
            ot_late=late.on_time.mean(), ot_late_n=len(late),
        ))
    score = pd.DataFrame(rows).set_index("vendor_id")

    # Fisher exact test on early vs late on-time (is the trend real?)
    from scipy.stats import fisher_exact
    pvals = {}
    for vid in score.index:
        R = rec[rec.vendor_id == vid]
        e = R[(R.req_month >= EARLY_PERIOD[0]) & (R.req_month <= EARLY_PERIOD[1])]
        l = R[(R.req_month >= LATE_PERIOD[0]) & (R.req_month <= LATE_PERIOD[1])]
        tbl = [[int(e.on_time.sum()), int((~e.on_time).sum())], [int(l.on_time.sum()), int((~l.on_time).sum())]]
        pvals[vid] = fisher_exact(tbl)[1] if min(len(e), len(l)) else np.nan
    score["trend_p"] = pd.Series(pvals)

    monthly_ot = rec.pivot_table(index="req_month", columns="vendor_id", values="on_time", aggfunc="mean")
    monthly_ot_n = rec.pivot_table(index="req_month", columns="vendor_id", values="on_time", aggfunc="size")

    # V004 service-level breakdown
    v4 = rec[rec.vendor_id == "V004"].groupby("part_id").agg(lines=("on_time", "size"), on_time=("on_time", "mean"),
                                                             avg_late=("days_late", lambda s: s[s > 0].mean()))

    dq += [
        ("receipt_txn.txn_date format", "MM/DD/YYYY", "Parsed explicitly; SQLite date() returns NULL on all rows."),
        ("po_line.qty_received disagrees with receipt_txn", int((line.set_index(["po_number", "line_no"]).qty_received
                                                               != t["rcv"].groupby(["po_number", "line_no"]).qty.sum()
                                                               .reindex(line.set_index(["po_number", "line_no"]).index).fillna(0)).sum()),
         "Field ignored; receipts are the source of truth."),
        ("Superseded confirmations", int(conf.superseded_by.notna().sum()), "Only the live (latest) confirmation is used."),
        ("Confirmation feed ends (last doc_date)", conf.doc_date.max(),
         "Confirmation metrics use POs dated on/before %s only." % conf_cutoff.date()),
        ("Cancelled POs", int((hdr.status == "CANCELLED").sum()), "Excluded everywhere (none have receipts)."),
        ("Ostmark EUR confirmations / PO price ratio", ", ".join(map(str, eur_ratio)),
         "PO prices are USD already; no FX applied to received value."),
        ("qc_hold dates before the PO existed", int((t["hold"].merge(hdr, on="po_number").eval("hold_date < po_date")).sum()),
         "Hold dates unreliable -> holds shown as counts only, not ranked on."),
        ("Unused/decoy tables", "gl_account, ship_method, user_master, plant_calendar, mrp_message",
         "No link to vendor performance (MRP messages reflect Beacon demand changes)."),
    ]

    return dict(
        as_of=as_of, first_day=first_day, last_day=last_day, conf_cutoff=conf_cutoff,
        vend=vend, short=short, monthly=monthly, monthly_counts=monthly_counts, score=score,
        monthly_ot=monthly_ot, monthly_ot_n=monthly_ot_n, rec=rec, past_due=past_due, above=above,
        superseded=sup, holds=holds, v4_parts=v4, dq=dq,
        naive=dict(correct=total_val, r_only=naive_R_only, qty_received_field=naive_qty_received,
                   ostmark_fx_again=v5_if_eur),
        part=t["part"], conf=conf, line=line, hdr=hdr,
    )


METRIC_DEFINITIONS = [
    ("Received value", "Sum of net receipt qty (R minus reversals; unreversed duplicate excluded) x PO line unit price, "
                       "bucketed by receipt month. USD. Receipts carry no price, so this is value at PO price, not invoiced cost."),
    ("Due lines", "Non-cancelled PO lines with required date on/before the as-of date and at least one valid receipt."),
    ("On-time %", "Due lines whose LAST valid receipt (the one that completes the line) is on/before the required date / due lines."),
    ("Severely late", "Due lines whose last receipt is more than %d days after the required date." % LATE_GRACE_DAYS),
    ("$ severely late", "Received value (capped at ordered qty) on severely late lines."),
    ("Promised late %", "Live confirmations whose promised date is after our required date / lines with a live confirmation."),
    ("On-time vs promise", "Lines received on/before the vendor's own promised date / lines with a live confirmation."),
    ("Lead time (median days from PO date)", "Allowed = required - PO date; Quoted = promised - PO date; Actual = last receipt - PO date."),
    ("Past-due open", "OPEN lines with required date before as-of and received qty < ordered; value = balance x unit price."),
    ("Price above PO", "Live USD confirmations priced above the PO line; exposure = price difference x ordered qty."),
    ("Trend", "On-time % for lines required Sep-Dec 2025 vs Jan-May 2026; Fisher exact test p-value (p<0.05 = real change)."),
]
