"""
Excel workbook for Lisa. Sheet order is the order she'd work in:

  1. Action List       - only what needs her, worst first, with what to do
  2. Review Required   - things the tool is NOT sure about (never auto-accepted)
  3. All PO Lines      - every open PO line, PO vs vendor side by side (audit trail)
  4. Vendor Summary    - counts per vendor
  5. Draft Emails      - one chase email per vendor, ready to edit
  6. Documents         - every PDF: how it was read, what it is, what happened to it
  7. PN Crosswalk      - vendor PN -> Beacon PN, incl. new ones pending review
  8. Run Summary       - counts + validation checks (did every PDF / PO line get handled?)
"""
from collections import defaultdict
from datetime import datetime

import xlsxwriter

from recon import ISSUE, SEV_LABEL, action_text, fmt_d


def _dt(d):
    return datetime.combine(d, datetime.min.time())


def write_workbook(path, results, exceptions, review, suggestions, docs, vendors, xw, as_of, item_status=None, backlog=None):
    """item_status: {item_key: (status, note, first_seen)} from the Desk memory, so Lisa's
    Done/Notes carry forward and items she already handled sink to the bottom."""
    item_status = item_status or {}
    wb = xlsxwriter.Workbook(path)
    H = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "text_wrap": True,
                       "valign": "top", "border": 1})
    T = wb.add_format({"bold": True, "font_size": 14})
    W = wb.add_format({"text_wrap": True, "valign": "top"})
    D = wb.add_format({"num_format": "mm/dd/yyyy", "valign": "top"})
    N = wb.add_format({"num_format": "#,##0", "valign": "top"})
    P = wb.add_format({"num_format": "$#,##0.0000", "valign": "top"})
    PCT = wb.add_format({"num_format": "+0.00%;-0.00%", "valign": "top"})
    M = wb.add_format({"num_format": "$#,##0", "valign": "top"})
    B = wb.add_format({"bold": True})
    sev_fmt = {1: wb.add_format({"bg_color": "#F8CBAD", "valign": "top", "bold": True}),
               2: wb.add_format({"bg_color": "#FFE699", "valign": "top"}),
               3: wb.add_format({"bg_color": "#DDEBF7", "valign": "top"}),
               4: wb.add_format({"bg_color": "#E2EFDA", "valign": "top"}),
               5: wb.add_format({"bg_color": "#E2EFDA", "valign": "top"})}
    yellow = wb.add_format({"bg_color": "#FFF2CC", "text_wrap": True, "valign": "top"})

    def vname(vid):
        return vendors.get(vid, {}).get("vendor_name", vid or "?")

    def vmail(vid):
        return vendors.get(vid, {}).get("ap_email", "")

    def header(ws, row, cols, widths):
        for c, (h, w) in enumerate(zip(cols, widths)):
            ws.write(row, c, h, H)
            ws.set_column(c, c, w)

    rows = sorted(results.values(), key=lambda r: (r["sev"], -r["impact"], r["required"]))
    n_action = sum(1 for r in rows if r["sev"] <= 3) + len(exceptions)

    # ---- 1. Action List --------------------------------------------------
    ws = wb.add_worksheet("Action List")
    ws.write(0, 0, "PO confirmation check - as of %s:  %d open PO lines | %d items need attention | %d lines fully clean" % (
        fmt_d(as_of), len(rows), n_action, sum(1 for r in rows if r["sev"] >= 4)), T)
    ws.write(1, 0, "Worst first. Red = act today, yellow = this week, blue = check. Type Yes in 'Done?' and add Notes - the next "
                   "run reads them back, so handled items stay handled. Uncertain matches are on 'Review Required'.", W)
    cols = ["Priority", "Issue type", "Vendor", "PO", "Line", "Our part", "Vendor PN", "What happened", "What to do",
            "Detail", "Ordered qty", "Confirmed qty", "Qty var", "PO price", "Confirmed price (USD)", "Price var %",
            "$ at stake", "Required", "Promised", "Days late", "Match confidence", "Source file(s)", "Vendor email",
            "Done?", "Notes", "Open since", "key"]
    widths = [12, 20, 20, 15, 5, 14, 13, 28, 50, 50, 9, 9, 8, 10, 11, 9, 10, 11, 11, 6, 10, 18, 24, 8, 25, 11, 4]
    header(ws, 3, cols, widths)
    ws.set_row(3, 30)
    grey = wb.add_format({"font_color": "#8A8A8A", "valign": "top", "text_wrap": True})

    def memory(rr, key):
        st, note, since = item_status.get(key, ("Open", "", as_of.isoformat()))
        ws.write(rr, 23, "Yes" if st == "Done" else "", W)
        ws.write(rr, 24, note or "", W)
        ws.write(rr, 25, ("NEW today" if since == as_of.isoformat() else since) if since else "", W)
        ws.write(rr, 26, key)
        return st == "Done"

    done_rows = []
    rr = 4
    for e in sorted(exceptions, key=lambda e: e["sev"]):
        ws.write(rr, 0, SEV_LABEL[e["sev"]], sev_fmt[e["sev"]])
        ws.write(rr, 1, e["code"], W)
        ws.write(rr, 2, vname(e["vendor"]), W)
        ws.write(rr, 3, e["po"], W)
        ws.write(rr, 7, e["issue"], W)
        ws.write(rr, 8, e["action"], W)
        ws.write(rr, 9, e["detail"], W)
        ws.write(rr, 21, e["file"], W)
        ws.write(rr, 22, vmail(e["vendor"]), W)
        memory(rr, "doc|%s|%s" % (e["file"], e["code"]))
        rr += 1
    line_rows = [r for r in rows if r["sev"] <= 3]
    line_rows.sort(key=lambda r: item_status.get("%s|%d" % (r["po_number"], r["line_number"]), ("Open",))[0] == "Done")
    for r in line_rows:
        shown = [i for i in r["issues"] if ISSUE[i][0] <= 3]
        is_done = item_status.get("%s|%d" % (r["po_number"], r["line_number"]), ("Open",))[0] == "Done"
        ws.write(rr, 0, "Done" if is_done else SEV_LABEL[r["sev"]], grey if is_done else sev_fmt[r["sev"]])
        ws.write(rr, 1, ", ".join(dict.fromkeys(ISSUE[i][1] for i in shown)), W)
        ws.write(rr, 2, r["vendor_name"], W)
        ws.write(rr, 3, r["po_number"], W)
        ws.write(rr, 4, r["line_number"], W)
        ws.write(rr, 5, r["our_pn"], W)
        ws.write(rr, 6, r["vendor_pn"], W)
        ws.write(rr, 7, "; ".join(ISSUE[i][2] for i in shown), W)
        ws.write(rr, 8, action_text(r), W)
        ws.write(rr, 9, "; ".join(r["notes"]), W)
        ws.write_number(rr, 10, r["qty"], N)
        if r["conf_qty"] is not None:
            ws.write_number(rr, 11, r["conf_qty"], N)
            ws.write_number(rr, 12, r["qty_var"], N)
        ws.write_number(rr, 13, r["price"], P)
        if r["conf_price_usd"] is not None:
            ws.write_number(rr, 14, r["conf_price_usd"], P)
            ws.write_number(rr, 15, r["price_var_pct"] / 100.0, PCT)
        if r["impact"]:
            ws.write_number(rr, 16, r["impact"], M)
        ws.write_datetime(rr, 17, _dt(r["required"]), D)
        if r["promise"]:
            ws.write_datetime(rr, 18, _dt(r["promise"]), D)
        if r["days_late"]:
            ws.write_number(rr, 19, r["days_late"], N)
        ws.write(rr, 20, r["confidence"], W)
        ws.write(rr, 21, r["files"], W)
        ws.write(rr, 22, vmail(r["vendor_id"]), W)
        memory(rr, "%s|%d" % (r["po_number"], r["line_number"]))
        rr += 1
    ws.set_column(26, 26, 4, None, {"hidden": True})
    ws.freeze_panes(4, 4)
    ws.autofilter(3, 0, max(rr - 1, 4), len(cols) - 1)
    ws.data_validation(4, 23, max(rr - 1, 4), 23, {"validate": "list", "source": ["Yes", "Waiting on vendor", "No"]})

    # ---- 2. Review Required ---------------------------------------------
    ws = wb.add_worksheet("Review Required")
    ws.write(0, 0, "The tool is not certain about these. Nothing here was silently accepted - check each one.", T)
    cols = ["What to review", "Vendor", "PO", "Line", "Vendor PN / source value", "Tool's best guess",
            "Description / what was read", "How matched", "Confidence", "Why it needs review",
            "Recommended check", "Source file", "Reviewed? (Y/N)"]
    widths = [18, 22, 15, 5, 16, 14, 40, 26, 10, 60, 45, 18, 10]
    header(ws, 2, cols, widths)
    for i, v in enumerate(review, 3):
        vals = [v["kind"], vname(v["vendor"]), v["po"], v["line"], v["source_pn"], v["candidate"], v["description"],
                v["method"], v["confidence"], v["reason"], v["recommended"], v["file"]]
        for c, x in enumerate(vals):
            ws.write(i, c, x if x is not None else "", W)
    ws.freeze_panes(3, 0)
    ws.autofilter(2, 0, max(len(review) + 2, 3), len(cols) - 1)

    # ---- 3. All PO Lines -------------------------------------------------
    ws = wb.add_worksheet("All PO Lines")
    cols = ["Status", "Vendor", "PO", "Line", "PO date", "Our part", "Vendor PN", "Description", "Ordered qty",
            "Confirmed qty", "Qty var", "PO price (USD)", "Confirmed price", "Ccy", "Confirmed price (USD)",
            "Price var %", "Required", "Promised", "Date var (days)", "Days until required", "Issue types",
            "Notes", "How matched", "Confidence", "Source file(s)", "First promise (never overwritten)", "Times vendor moved date"]
    widths = [12, 22, 15, 5, 11, 14, 13, 30, 9, 9, 8, 10, 10, 5, 11, 9, 11, 11, 8, 9, 30, 60, 26, 10, 22, 13, 9]
    header(ws, 0, cols, widths)
    ws.set_row(0, 30)
    for i, r in enumerate(sorted(results.values(), key=lambda r: (r["vendor_id"], r["po_number"], r["line_number"])), 1):
        ws.write(i, 0, SEV_LABEL[r["sev"]], sev_fmt[r["sev"]])
        ws.write(i, 1, r["vendor_name"])
        ws.write(i, 2, r["po_number"])
        ws.write(i, 3, r["line_number"])
        ws.write_datetime(i, 4, _dt(r["po_date"]), D)
        ws.write(i, 5, r["our_pn"])
        ws.write(i, 6, r["vendor_pn"])
        ws.write(i, 7, r["description"])
        ws.write_number(i, 8, r["qty"], N)
        if r["conf_qty"] is not None:
            ws.write_number(i, 9, r["conf_qty"], N)
            ws.write_number(i, 10, r["qty_var"], N)
        ws.write_number(i, 11, r["price"], P)
        if r["conf_price"] is not None:
            ws.write_number(i, 12, r["conf_price"], P)
            ws.write(i, 13, r["conf_ccy"])
            ws.write_number(i, 14, r["conf_price_usd"], P)
            ws.write_number(i, 15, r["price_var_pct"] / 100.0, PCT)
        ws.write_datetime(i, 16, _dt(r["required"]), D)
        if r["promise"]:
            ws.write_datetime(i, 17, _dt(r["promise"]), D)
        if r["date_var"] is not None:
            ws.write_number(i, 18, r["date_var"], N)
        ws.write_number(i, 19, r["days_to_required"], N)
        ws.write(i, 20, ", ".join(dict.fromkeys(ISSUE[x][1] for x in r["issues"])) or "MATCH")
        ws.write(i, 21, "; ".join(r["notes"]))
        ws.write(i, 22, r["match"])
        ws.write(i, 23, r["confidence"])
        ws.write(i, 24, r["files"])
        if r.get("first_promise"):
            ws.write_datetime(i, 25, _dt(r["first_promise"]), D)
        if r.get("promise_changes"):
            ws.write_number(i, 26, r["promise_changes"], N)
    ws.freeze_panes(1, 4)
    ws.autofilter(0, 0, len(results), len(cols) - 1)

    # ---- 4. Vendor Summary -----------------------------------------------
    ws = wb.add_worksheet("Vendor Summary")
    keys = ["DROPPED", "NO_CONF", "NO_DETAIL", "QTY_SHORT", "LATE", "PRICE_UP", "DATE_ODD", "PN_INFERRED"]
    cols = ["Vendor", "Open POs", "Open PO lines", "Lines confirmed", "Clean lines", "Lines needing action"] + \
           [ISSUE[k][1] for k in keys] + ["$ at stake", "Docs for POs not open"]
    header(ws, 0, cols, [30] + [12] * (len(cols) - 1))
    ws.set_row(0, 45)
    by_v = defaultdict(list)
    for r in results.values():
        by_v[r["vendor_id"]].append(r)
    for i, vid in enumerate(sorted(by_v), 1):
        rs = by_v[vid]
        ws.write(i, 0, vname(vid))
        ws.write(i, 1, len(set(r["po_number"] for r in rs)))
        ws.write(i, 2, len(rs))
        ws.write(i, 3, sum(1 for r in rs if r["conf_qty"] is not None))
        ws.write(i, 4, sum(1 for r in rs if r["sev"] >= 4))
        ws.write(i, 5, sum(1 for r in rs if r["sev"] <= 3))
        for c, k in enumerate(keys, 6):
            n = sum(1 for r in rs if k in r["issues"])
            if n:
                ws.write(i, c, n)
        ws.write_number(i, 6 + len(keys), sum(r["impact"] for r in rs), M)
        n = sum(1 for e in exceptions if e["vendor"] == vid)
        if n:
            ws.write(i, 7 + len(keys), n)

    # ---- 5. Draft Emails -------------------------------------------------
    ws = wb.add_worksheet("Draft Emails")
    header(ws, 0, ["Vendor", "To", "Draft email - review and edit before sending (the tool never sends anything)"], [26, 28, 110])
    for i, (vid, (to, body)) in enumerate(sorted(draft_emails(results, exceptions, vendors).items()), 1):
        ws.write(i, 0, vname(vid), W)
        ws.write(i, 1, to, W)
        ws.write(i, 2, body, W)
        ws.set_row(i, 15 * (body.count("\n") + 2))

    # ---- 6. Documents ----------------------------------------------------
    ws = wb.add_worksheet("Documents")
    cols = ["File", "Vendor", "Layout recognised", "Read via", "Doc type", "PO", "Doc date", "Vendor ref",
            "Lines read", "Status", "Warnings"]
    header(ws, 0, cols, [18, 24, 34, 8, 22, 15, 11, 14, 8, 32, 70])
    for i, d in enumerate(sorted(docs, key=lambda d: d["file"]), 1):
        ws.write(i, 0, d["file"])
        ws.write(i, 1, vname(d["vendor_id"]) if d["vendor_id"] else "?")
        ws.write(i, 2, d["template"])
        ws.write(i, 3, (d.get("text_source") or "-").upper())
        ws.write(i, 4, d["doc_type"] + (" (part %d of %d)" % d["part_of"] if d["part_of"] else ""))
        ws.write(i, 5, d["po_number"] or "?")
        if d["doc_date"]:
            ws.write_datetime(i, 6, _dt(d["doc_date"]), D)
        ws.write(i, 7, d["vendor_ref"] or "")
        ws.write(i, 8, len(d["lines"]))
        ws.write(i, 9, d.get("status", ""))
        ws.write(i, 10, "; ".join(dict.fromkeys(d["warnings"])))
    ws.autofilter(0, 0, len(docs), len(cols) - 1)

    # ---- 7. PN Crosswalk -------------------------------------------------
    ws = wb.add_worksheet("PN Crosswalk")
    ws.write(0, 0, "Vendor part number -> Beacon part number, keyed by (vendor, vendor PN). Yellow rows are PENDING REVIEW and are "
                   "not used for auto-matching. To approve one, type APPROVED in 'Review status' (or REJECTED) - the next run "
                   "picks it up. This table is the seed of Beacon's vendor-part master.", W)
    ws.set_row(0, 45)
    cols = ["Vendor", "Vendor PN", "Beacon PN", "Description", "Mapping source", "Confidence", "Times seen",
            "First seen", "Last seen", "Review status", "Note", "vendor_id"]
    header(ws, 2, cols, [26, 15, 15, 30, 48, 10, 9, 11, 11, 22, 45, 4])
    merged = dict(xw)
    for k, s in suggestions.items():
        merged.setdefault(k, s)
    pend = lambda x: str(x.get("review_status", "")).startswith("PENDING")
    i = 3
    for x in sorted(merged.values(), key=lambda x: (not pend(x), x["vendor_id"], x["vendor_pn"])):
        old = [y["vendor_pn"] for y in merged.values() if y["vendor_id"] == x["vendor_id"] and y["beacon_pn"] == x["beacon_pn"]
               and y["vendor_pn"] != x["vendor_pn"]]
        shared = [y["vendor_id"] for y in merged.values() if y["vendor_pn"] == x["vendor_pn"] and y["vendor_id"] != x["vendor_id"]]
        note = []
        if pend(x) and old:
            note.append("Our history has this part from this vendor as %s - vendor appears to have renumbered" % ", ".join(old))
        if shared:
            note.append("Same PN used by %s for a different part - always key on vendor + PN" % ", ".join(vname(v) for v in shared))
        vals = [vname(x["vendor_id"]), x["vendor_pn"], x["beacon_pn"], x.get("description", ""), x.get("mapping_source", ""),
                x.get("mapping_confidence", ""), x.get("times_seen", ""), x.get("first_seen", ""), x.get("last_seen", ""),
                x.get("review_status", ""), "; ".join(note), x["vendor_id"]]
        for c, v in enumerate(vals):
            ws.write(i, c, v if v is not None else "", yellow if pend(x) else None)
        i += 1
    ws.set_column(11, 11, 4, None, {"hidden": True})
    ws.data_validation(3, 9, max(i - 1, 3), 9, {"validate": "list", "source": [
        "PENDING REVIEW", "APPROVED", "REJECTED", "APPROVED (historical)"]})

    # ---- 8. Old Open Balances (from ERP history) ---------------------------
    if backlog:
        ws = wb.add_worksheet("Old Open Balances")
        ws.write(0, 0, "From the ERP: PO lines past their required date with quantity still open (%d lines, $%s). Months-old "
                       "short-shipped balances are usually vendor short-ships nobody closed - chase the balance or short-close the "
                       "line so MRP stops counting on it." % (len(backlog), format(round(sum(b["balance_value"] for b in backlog)), ",")), W)
        ws.set_row(0, 45)
        cols = ["Vendor", "PO", "Line", "Part", "Ordered", "Received", "Open balance", "Balance $", "Required",
                "Days past due", "What it is", "Last receipt", "Suggested action"]
        header(ws, 2, cols, [28, 16, 6, 16, 10, 10, 12, 12, 11, 10, 30, 11, 50])
        for i, b in enumerate(backlog, 3):
            act = ("Chase: nothing received yet." if b["kind"].startswith("Nothing") else
                   "Ask vendor if the balance will ship; if not, short-close the line." if b["days_past_due"] <= 60 else
                   "Probably never shipping: confirm with vendor and short-close.")
            vals = [vname(b["vendor_id"]), b["po_number"], b["line_no"], b["part_id"], b["ordered"], b["received"],
                    b["balance"], b["balance_value"]]
            for c, v in enumerate(vals):
                ws.write(i, c, v, M if c == 7 else (N if c in (4, 5, 6) else None))
            ws.write_datetime(i, 8, _dt(b["required"]), D)
            ws.write_number(i, 9, b["days_past_due"], N)
            ws.write(i, 10, b["kind"])
            if b["last_receipt"]:
                ws.write_datetime(i, 11, _dt(b["last_receipt"]), D)
            ws.write(i, 12, act, W)
        ws.autofilter(2, 0, 2 + len(backlog), len(cols) - 1)
        ws.freeze_panes(3, 2)

    # ---- 9. Run Summary --------------------------------------------------
    stats = run_stats(results, exceptions, review, docs, suggestions)
    ws = wb.add_worksheet("Run Summary")
    ws.set_column(0, 0, 55)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 70)
    ws.write(0, 0, "Run summary and validation checks - as of %s" % fmt_d(as_of), T)
    for i, (k, v) in enumerate(stats.items(), 2):
        ws.write(i, 0, k, B if k.startswith("CHECK") else None)
        ws.write(i, 1, v)
    wb.close()
    return stats


def draft_emails(results, exceptions, vendors, buyer="Lisa", skip_keys=()):
    """{vendor_id: (to_address, email_text)} for every vendor with something to ask.
    skip_keys: (po, line) pairs Lisa already closed, so they aren't asked again."""
    by_v = defaultdict(list)
    for r in results.values():
        by_v[r["vendor_id"]].append(r)
    out = {}
    for vid in sorted(set(by_v) | set(e["vendor"] for e in exceptions)):
        items = [r for r in by_v.get(vid, []) if r["sev"] <= 3 and set(r["issues"]) - {"PN_INFERRED", "INVOICE_ONLY", "OCR"}
                 and (r["po_number"], r["line_number"]) not in skip_keys]
        exc = [e for e in exceptions if e["vendor"] == vid and e["code"] != "UNREADABLE_DOCUMENT"]
        bullets = []
        for r in items:
            b = []
            if "DROPPED" in r["issues"]:
                b.append("line %d (%s, qty %d) is not on your acknowledgment - please confirm you will ship it" % (
                    r["line_number"], r["our_pn"], r["qty"]))
            if "DECLINED" in r["issues"]:
                b.append("you indicated you cannot supply line %d (%s, qty %d) - is there any alternative (partial qty, "
                         "later date, substitute material)?" % (r["line_number"], r["our_pn"], r["qty"]))
            if "FORM_INCOMPLETE" in r["issues"]:
                b.append("line %d on the acknowledgment form is incomplete (%s) - please complete it and send the form back" % (
                    r["line_number"], "; ".join(n.replace("Form problem: ", "") for n in r["notes"] if n.startswith("Form problem"))))
            if "NO_CONF" in r["issues"]:
                b.append("we have not received an acknowledgment - please confirm qty, price and ship date")
            if "NO_DETAIL" in r["issues"]:
                b.append("please confirm quantity and a firm ship date")
            if "QTY_SHORT" in r["issues"]:
                b.append("you confirmed %d but we ordered %d - will the balance ship?" % (r["conf_qty"], r["qty"]))
            if "QTY_OVER" in r["issues"]:
                b.append("you confirmed %d but we ordered %d - please ship only %d" % (r["conf_qty"], r["qty"], r["qty"]))
            if "PRICE_UP" in r["issues"] or "PRICE_DOWN" in r["issues"]:
                b.append("confirmed price %.4f %s does not match our PO price of $%.4f - please correct" % (
                    r["conf_price"], r["conf_ccy"], r["price"]))
            if "LATE" in r["issues"]:
                b.append("promise date %s is after our required date %s - can you ship by %s?" % (
                    fmt_d(r["promise"]), fmt_d(r["required"]), fmt_d(r["required"])))
            if "VAGUE_DATE" in r["issues"]:
                b.append("please give a specific ship date rather than a week range")
            if "DATE_ODD" in r["issues"]:
                b.append("the dates on your acknowledgment look inconsistent - please re-confirm PO number and ship date")
            if b:
                bullets.append("  - %s line %d: %s" % (r["po_number"], r["line_number"], "; ".join(b)))
        for e in exc:
            if e["code"] == "PO_NOT_OPEN":
                bullets.append("  - %s: you sent an acknowledgment for this PO number, but it is not an open Beacon PO. "
                               "Please do not ship against it and let us know which PO you meant." % e["po"])
            else:
                bullets.append("  - %s: %s - please check." % (e["po"], e["issue"]))
        if not bullets:
            continue
        body = ["Subject: Beacon Fasteners - open items on your order acknowledgments", "", "Hello,", "",
                "We reviewed your recent acknowledgments and need your help with the following:", ""] + bullets + \
               ["", "Thank you,", buyer, "Beacon Fasteners - Purchasing"]
        out[vid] = (vendors.get(vid, {}).get("ap_email", ""), "\n".join(body))
    return out


def run_stats(results, exceptions, review, docs, suggestions):
    used = [d for d in docs if not d.get("duplicate_of")]
    matched_lines = sum(1 for r in results.values() if r["conf_qty"] is not None)
    by_sev = defaultdict(int)
    for r in results.values():
        by_sev[r["sev"]] += 1
    unaccounted = [k for k, r in results.items()
                   if r["conf_qty"] is None and not set(r["issues"]) & {"DROPPED", "NO_CONF", "NO_DETAIL", "DECLINED", "FORM_INCOMPLETE"}]
    no_status = [d["file"] for d in docs if not d.get("status")]
    return {
        "PDFs processed": len(docs),
        "  read from text layer": sum(1 for d in docs if d.get("text_source") == "text"),
        "  read by OCR (scans)": sum(1 for d in docs if d.get("text_source") == "ocr"),
        "  Beacon forms (no reading needed)": sum(1 for d in docs if d.get("text_source") == "form"),
        "  unreadable (manual)": sum(1 for d in docs if d["doc_type"] == "UNREADABLE"),
        "  duplicates ignored": len(docs) - len(used),
        "  superseded by a later revision": sum(1 for d in docs if str(d.get("status", "")).startswith("Superseded")),
        "  for POs not on the open list": sum(1 for d in docs if d.get("status") == "PO not on open list"),
        "Open PO lines": len(results),
        "  matched to a confirmation line": matched_lines,
        "  possible dropped lines": sum(1 for r in results.values() if "DROPPED" in r["issues"]),
        "  no confirmation at all": sum(1 for r in results.values() if "NO_CONF" in r["issues"]),
        "  acknowledged without qty/date": sum(1 for r in results.values() if "NO_DETAIL" in r["issues"]),
        "  declined by vendor (Beacon form)": sum(1 for r in results.values() if "DECLINED" in r["issues"]),
        "  incomplete on Beacon form": sum(1 for r in results.values() if "FORM_INCOMPLETE" in r["issues"]),
        "Priority 1 (act today) lines": by_sev[1],
        "Priority 2 (this week) lines": by_sev[2],
        "Priority 3 (check) lines": by_sev[3],
        "Clean lines": by_sev[4] + by_sev[5],
        "Document-level exceptions": len(exceptions),
        "Review-required items": len(review),
        "New vendor PNs pending review": len(suggestions),
        "CHECK every PDF has a status": "PASS" if not no_status else "FAIL: %s" % no_status,
        "CHECK every open PO line is accounted for": "PASS" if not unaccounted else "FAIL: %s" % unaccounted,
    }
