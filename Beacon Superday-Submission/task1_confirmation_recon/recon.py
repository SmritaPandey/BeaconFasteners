"""
Beacon PO confirmation reconciler - matching and checks.

Pipeline:  PDF -> parsers.parse_file() -> normalized doc
           docs + open PO list -> reconcile() -> one result per open PO line,
                                               + exceptions + review queue
           results -> report.write_workbook() -> Excel for Lisa

Everything here is deterministic. Nothing is auto-accepted that we are not
sure of: uncertain part-number matches go to the Review Required sheet.

Usage:
  python3 recon.py --confirmations DIR --pos open_pos.csv --vendors vendor_master.csv
                   [--erp beacon_erp.db] [--as-of YYYY-MM-DD] [--out DIR]
"""
import argparse
import csv
import glob
import hashlib
import os
import re
import sqlite3
from collections import defaultdict
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))

# --------------------------------------------------------------------------
# Tolerances (see README for why)
# --------------------------------------------------------------------------
PRICE_TOL_USD = 0.00005      # any real change in a USD price is flagged
PRICE_TOL_FX_PCT = 1.0       # EUR prices: flag only if >1% off after FX (monthly-rate noise is ~0.3%)
DEFAULT_EUR_USD = 1.09       # only if no ERP fx table is available

# Issue catalogue: key -> (severity, issue_type code, plain-English label)
# Severity 1 = act today, 2 = act this week, 3 = check / review, 4 = FYI
ISSUE = {
    "NO_CONF":      (1, "NO_CONFIRMATION",       "No confirmation received"),
    "DROPPED":      (1, "POSSIBLE_DROPPED_LINE", "Line missing from vendor's confirmation"),
    "DECLINED":     (1, "VENDOR_DECLINED_LINE",  "Vendor says it cannot supply this line"),
    "FORM_INCOMPLETE": (2, "FORM_INCOMPLETE",    "Vendor's acknowledgment form left this line incomplete"),
    "QTY_SHORT":    (2, "QUANTITY_SHORT",        "Confirmed qty short"),
    "LATE":         (2, "DATE_LATE",             "Promise date after required date"),
    "PRICE_UP":     (2, "PRICE_MISMATCH",        "Price increase"),
    "NO_DETAIL":    (2, "MISSING_REQUIRED_FIELD", "Acknowledged with no qty/date"),
    "QTY_OVER":     (3, "QUANTITY_OVER",         "Confirmed qty over"),
    "PRICE_DOWN":   (3, "PRICE_MISMATCH",        "Price decrease"),
    "INVOICE_ONLY": (3, "INVOICE_NOT_ACK",       "Vendor sent an invoice, not an acknowledgment"),
    "PN_INFERRED":  (3, "PART_NUMBER_UNMATCHED", "Vendor part no. not in crosswalk - matched by qty/price"),
    "DATE_ODD":     (3, "DATE_INCONSISTENT",     "Dates on the document don't make sense"),
    "VAGUE_DATE":   (3, "DATE_AMBIGUOUS",        "Promise is a week range, not a date"),
    "SPLIT":        (4, "SPLIT_CONFIRMATION",    "Confirmed across multiple documents/deliveries"),
    "NO_PRICE":     (4, "PRICE_NOT_STATED",      "Vendor did not state a price"),
    "OCR":          (4, "OCR_SOURCE",            "Read from a scanned image (OCR) or by AI"),
    "SUPERSEDED":   (4, "REVISED_CONFIRMATION",  "Earlier confirmation replaced by a revision"),
}
SEV_LABEL = {1: "1 - Act today", 2: "2 - This week", 3: "3 - Check", 4: "OK", 5: "OK"}

# Match confidence. Identifier-based matches are HIGH. A match inferred from
# qty + price is MEDIUM: on this data every inference is an exact qty AND
# exact price hit on a single-vendor PO, but it is still a guess about the
# vendor's part number, so it goes to review rather than into the crosswalk.
HIGH, MEDIUM, LOW = "HIGH", "MEDIUM", "LOW"


def norm(s):
    s = (s or "").lower().replace(" per ib", " per lb")
    return re.sub(r"[^a-z0-9]", "", s)


def fmt_d(d):
    return d.strftime("%m/%d/%Y") if d else ""


# --------------------------------------------------------------------------
# Inputs
# --------------------------------------------------------------------------

def load_pos(path):
    lines = []
    with open(path, newline="") as f:
        for r in csv.DictReader(f):
            lines.append(dict(
                po_number=r["po_number"].strip(),
                po_date=datetime.strptime(r["po_date"], "%Y-%m-%d").date(),
                vendor_id=r["vendor_id"], vendor_name=r["vendor_name"],
                line_number=int(r["line_number"]), our_pn=r["our_pn"].strip(),
                description=r["our_description"], qty=float(r["qty_ordered"]),
                price=float(r["unit_price"]),
                required=datetime.strptime(r["required_date"], "%Y-%m-%d").date()))
    return lines


def load_vendors(path):
    with open(path, newline="") as f:
        return {r["vendor_id"]: r for r in csv.DictReader(f)}


def load_crosswalk(path):
    """Only APPROVED mappings are used for auto-matching."""
    xw = {}
    if path and os.path.exists(path):
        with open(path, newline="") as f:
            for r in csv.DictReader(f):
                xw[(r["vendor_id"], r["vendor_pn"])] = r
    return xw


def load_fx(erp):
    """{('2026-05','EUR'): 1.0902, ...} from the ERP, else empty (-> default rate)."""
    if erp and os.path.exists(erp):
        try:
            con = sqlite3.connect("file:%s?mode=ro" % erp, uri=True)
            return {(m, c): r for m, c, r in con.execute("select month, currency, rate_to_usd from fx_rate")}
        except sqlite3.Error:
            pass
    return {}


def fx_rate(fx, ccy, d):
    if ccy == "USD":
        return 1.0
    return fx.get((d.strftime("%Y-%m"), ccy), DEFAULT_EUR_USD if ccy == "EUR" else None)


# --------------------------------------------------------------------------
# Document hygiene
# --------------------------------------------------------------------------

def mark_duplicates(docs):
    """Identical content sent twice (re-sent email, same PDF saved twice) is
    ignored after the first copy so it can't double-count a split delivery."""
    seen = {}
    for d in docs:
        h = hashlib.sha1(re.sub(r"\s+", " ", d.get("raw_text", "")).encode()).hexdigest()
        if h in seen:
            d["duplicate_of"] = seen[h]
        else:
            seen[h] = d["file"]
            d["duplicate_of"] = None
    return [d for d in docs if not d["duplicate_of"]]


def validate_doc(d):
    """Sanity checks on extracted values. Problems become warnings, not crashes."""
    for dl in d["lines"]:
        if dl.get("qty") is None or dl["qty"] <= 0:
            d["warnings"].append("Non-positive or missing quantity on a line")
        if dl.get("unit_price") is not None and dl["unit_price"] <= 0:
            d["warnings"].append("Non-positive price on a line")
        if dl.get("promise_date") is None:
            d["warnings"].append("Missing promise date on a line")
    if not d.get("po_number"):
        d["warnings"].append("No PO number found")


def pick_docs(docs):
    """For one PO: drop superseded docs, keep declared split parts together.
    Returns (active_docs, superseded_docs)."""
    docs = sorted(docs, key=lambda d: (d["doc_date"] or date.min))
    if len(docs) <= 1 or all(d["part_of"] for d in docs):
        return docs, []
    # Otherwise the latest full document wins (revisions, re-sends, invoice after ack)
    return [docs[-1]], docs[:-1]


# --------------------------------------------------------------------------
# Part-number / line resolution (hierarchical)
# --------------------------------------------------------------------------

def resolve_line(dl, doc, po_lines, xw, taken):
    """Which PO line does this vendor doc line refer to?
    Returns (po_line or None, method, confidence, note).

    Level 1  Beacon PN printed on the doc                  HIGH
    Level 2  Approved (vendor, vendor PN) crosswalk entry   HIGH
    Level 3  Normalized description equals PO description  HIGH
    Level 4  Inferred: qty + price (+ line no.) on this PO  MEDIUM/LOW -> review
    Level 5  Nothing -> unmatched doc line (exception)
    """
    vid, vpn = doc["vendor_id"], dl.get("vendor_pn")
    by_pn = defaultdict(list)
    for pl in po_lines:
        by_pn[pl["our_pn"]].append(pl)

    def choose(cands):
        if len(cands) == 1:
            return cands[0]
        for c in cands:  # same part on several lines: prefer line no, then qty
            if dl.get("line_no") == c["line_number"]:
                return c
        for c in cands:
            if dl.get("qty") == c["qty"]:
                return c
        return cands[0]

    if vpn and vpn in by_pn:
        return choose(by_pn[vpn]), "L1 Beacon PN on document", HIGH, None

    note = None
    x = xw.get((vid, vpn)) if vpn else None
    if x and x.get("review_status", "").startswith("APPROVED"):
        if x["beacon_pn"] in by_pn:
            return choose(by_pn[x["beacon_pn"]]), "L2 Crosswalk %s -> %s" % (vpn, x["beacon_pn"]), HIGH, None
        note = "Crosswalk says %s = %s, but that part is not on this PO" % (vpn, x["beacon_pn"])

    d = norm(dl.get("description"))
    if d and "seepo" not in d and "siehebestellung" not in d:
        hits = [pl for pl in po_lines if norm(pl["description"]) == d]
        if hits:
            return choose(hits), "L3 Description match", HIGH, note

    open_lines = [pl for pl in po_lines if (pl["po_number"], pl["line_number"]) not in taken]
    scored = []
    for pl in open_lines:
        s = 0
        if dl.get("qty") == pl["qty"]:
            s += 2
        usd = dl.get("unit_price_usd")
        if usd is not None and (abs(usd - pl["price"]) < 0.0001 if dl.get("currency") == "USD"
                                else abs(usd - pl["price"]) / pl["price"] * 100 <= PRICE_TOL_FX_PCT):
            s += 2
        if dl.get("line_no") == pl["line_number"]:
            s += 1
        scored.append((s, pl))
    best_score = max([s for s, _ in scored] or [0])
    top = [pl for s, pl in scored if s == best_score]
    if best_score >= 2 and len(top) > 1:
        # Two PO lines fit equally well -> refuse to guess
        return None, "L5 Unmatched (ambiguous)", None, \
            "Could be line %s - equally likely; needs a human" % " or ".join(str(p["line_number"]) for p in top)
    if best_score >= 4:
        return top[0], "L4 Inferred (qty+price%s)" % ("+line no" if best_score >= 5 else ""), MEDIUM, note
    if best_score >= 2:
        return top[0], "L4 Inferred (weak: score %d)" % best_score, LOW, note
    return None, "L5 Unmatched", None, note


# --------------------------------------------------------------------------
# Reconciliation
# --------------------------------------------------------------------------

def reconcile(po_lines, docs, xw, fx, as_of):
    pos = defaultdict(list)
    for pl in po_lines:
        pos[pl["po_number"]].append(pl)
    docs_by_po = defaultdict(list)
    for d in docs:
        d.setdefault("status", "Used")
        docs_by_po[d["po_number"]].append(d)

    results = {}       # (po, line) -> dict
    exceptions = []    # not tied to an open PO line
    review = []        # uncertain things a human should look at
    suggestions = {}   # new crosswalk rows inferred this run
    extracted = []     # normalized confirmation lines (ERP-loadable)

    for po, lines in pos.items():
        vdocs = [d for d in docs_by_po.get(po, []) if d["doc_type"] != "UNREADABLE"]
        active, superseded = pick_docs(vdocs)
        for s in superseded:
            s["status"] = "Superseded by %s" % active[-1]["file"]

        matched = defaultdict(list)   # line key -> [(doc, doc_line, method, conf)]
        taken = set()
        for doc in active:
            for dl in doc["lines"]:
                if dl.get("unit_price") is not None:
                    rate = fx_rate(fx, dl.get("currency", "USD"), lines[0]["po_date"])
                    dl["unit_price_usd"] = dl["unit_price"] * rate if rate else None
                pl, method, conf, note = resolve_line(dl, doc, lines, xw, taken)
                if pl is None:
                    exceptions.append(dict(
                        sev=2, code="UNMATCHED_DOC_LINE", vendor=doc["vendor_id"], po=po, file=doc["file"],
                        issue="Line on vendor doc doesn't match any line on our PO",
                        detail="%s qty %s @ %s %s. %s" % (dl.get("vendor_pn") or dl.get("description"),
                                                          dl.get("qty"), dl.get("unit_price"), dl.get("currency"), note or ""),
                        action="Ask vendor what this line is - they may have mixed up POs."))
                    continue
                key = (po, pl["line_number"])
                taken.add(key)
                matched[key].append((doc, dl, method, conf))
                vpn = dl.get("vendor_pn")
                if conf != HIGH:
                    hist = [x["vendor_pn"] for x in xw.values()
                            if x["vendor_id"] == doc["vendor_id"] and x["beacon_pn"] == pl["our_pn"]]
                    review.append(dict(
                        kind="Part number", vendor=doc["vendor_id"], po=po, line=pl["line_number"],
                        source_pn=vpn, candidate=pl["our_pn"], description=pl["description"],
                        method=method, confidence=conf, file=doc["file"],
                        reason="Vendor PN '%s' is not in the approved crosswalk. It matched PO line %d on %s.%s" % (
                            vpn, pl["line_number"], method.split("(")[-1].rstrip(")"),
                            " ERP history knows this part from this vendor as %s - vendor may have renumbered." % ", ".join(hist) if hist else ""),
                        recommended="Confirm with vendor/drawing; if right, set review_status=APPROVED in pn_crosswalk.csv"))
                    if vpn and vpn != pl["our_pn"]:
                        suggestions[(doc["vendor_id"], vpn)] = dict(
                            vendor_id=doc["vendor_id"], vendor_pn=vpn, beacon_pn=pl["our_pn"],
                            description=pl["description"],
                            mapping_source="Inferred from %s (%s)" % (doc["file"], method),
                            mapping_confidence=conf, times_seen=1,
                            first_seen=doc["doc_date"].isoformat(), last_seen=doc["doc_date"].isoformat(),
                            review_status="PENDING REVIEW")
                own = dl.get("own_pn")
                if own and own.upper() != "SAME" and own != pl["our_pn"] and (doc["vendor_id"], own) not in xw:
                    # Vendor told us its own part number on Beacon's form: a strong hint, but still Lisa's call
                    suggestions[(doc["vendor_id"], own)] = dict(
                        vendor_id=doc["vendor_id"], vendor_pn=own, beacon_pn=pl["our_pn"], description=pl["description"],
                        mapping_source="Vendor-declared on acknowledgment form %s" % doc["file"], mapping_confidence=HIGH,
                        times_seen=1, first_seen=doc["doc_date"].isoformat(), last_seen=doc["doc_date"].isoformat(),
                        review_status="PENDING REVIEW")
                extracted.append(dict(
                    po_number=po, line_no=pl["line_number"], vendor_pn=vpn, part_id=pl["our_pn"],
                    confirmed_qty=dl.get("qty"), confirmed_price=dl.get("unit_price"), currency=dl.get("currency"),
                    promised_date=dl.get("promise_date"), doc_date=doc["doc_date"], source_file=doc["file"],
                    match_method=method, match_confidence=conf))

        for pl in lines:
            key = (po, pl["line_number"])
            r = dict(pl, issues=[], notes=[], vendor_pn="", conf_qty=None, conf_price=None, conf_ccy=None,
                     conf_price_usd=None, promise=None, days_late=None, date_var=None, qty_var=None,
                     price_var=None, price_var_pct=None, impact=0.0, match="", confidence="",
                     files=", ".join(sorted(set(d["file"] for d in active))))
            hits = matched.get(key, [])
            for s in superseded:
                r["issues"].append("SUPERSEDED")
                r["notes"].append("%s (%s) replaced by %s" % (s["file"], fmt_d(s["doc_date"]), active[-1]["file"]))
            if not vdocs:
                r["issues"].append("NO_CONF")
                r["notes"].append("PO issued %s, %d days before %s" % (
                    fmt_d(pl["po_date"]), (as_of - pl["po_date"]).days, fmt_d(as_of)))
            elif not hits and all(d["doc_type"] == "ACK_NO_DETAIL" for d in active):
                r["issues"].append("NO_DETAIL")
                r["notes"].append("; ".join(active[0]["warnings"]))
            elif not hits and any(pl["line_number"] in d.get("declined", {}) for d in active):
                r["issues"].append("DECLINED")
                r["impact"] = pl["qty"] * pl["price"]
                r["notes"].append("Vendor's reason: %s" % "; ".join(
                    d["declined"][pl["line_number"]] for d in active if pl["line_number"] in d.get("declined", {})))
            elif not hits and any(pl["line_number"] in d.get("incomplete", {}) for d in active):
                r["issues"].append("FORM_INCOMPLETE")
                r["notes"].append("Form problem: %s" % "; ".join(
                    d["incomplete"][pl["line_number"]] for d in active if pl["line_number"] in d.get("incomplete", {})))
            elif not hits:
                r["issues"].append("DROPPED")
                r["impact"] = pl["qty"] * pl["price"]
                r["notes"].append("Vendor confirmed %d other line(s) on this PO but not this one. "
                                  "Not proof it's dropped - but nothing on file says they'll ship it." %
                                  sum(len(d["lines"]) for d in active))
            else:
                _check_line(r, pl, hits, fx)
            r["issues"] = list(dict.fromkeys(r["issues"]))
            r["sev"] = min([ISSUE[i][0] for i in r["issues"]] or [5])
            r["days_to_required"] = (pl["required"] - as_of).days
            results[key] = r

    _document_exceptions(docs, pos, po_lines, results, exceptions, review)
    return results, exceptions, review, suggestions, extracted


def _check_line(r, pl, hits, fx):
    docs_used = [h[0] for h in hits]
    r["files"] = ", ".join(sorted(set(d["file"] for d in docs_used)))
    r["match"] = "; ".join(sorted(set(h[2] for h in hits)))
    r["confidence"] = min((h[3] for h in hits), key=[HIGH, MEDIUM, LOW].index)
    r["vendor_pn"] = ", ".join(sorted(set(h[1].get("vendor_pn") or "" for h in hits) - {""}))
    r["conf_qty"] = sum(h[1]["qty"] or 0 for h in hits)
    proms = [h[1]["promise_date"] for h in hits if h[1]["promise_date"]]
    r["promise"] = max(proms) if proms else None
    prices = [h[1] for h in hits if h[1]["unit_price"] is not None]

    if len(hits) > 1:
        r["issues"].append("SPLIT")
        r["notes"].append("Deliveries: " + "; ".join("%d on %s" % (h[1]["qty"], fmt_d(h[1]["promise_date"])) for h in hits))
    if r["confidence"] != HIGH:
        r["issues"].append("PN_INFERRED")
        r["notes"].append("Vendor PN '%s' matched by qty/price - see Review Required" % r["vendor_pn"])
    if any(d["text_source"] in ("ocr", "ai") for d in docs_used):
        r["issues"].append("OCR")
    for h in hits:
        if h[1].get("vendor_reason"):
            r["notes"].append("Vendor's reason: %s" % h[1]["vendor_reason"])
    if any(d["doc_type"] == "INVOICE" for d in docs_used):
        r["issues"].append("INVOICE_ONLY")
        r["notes"].append("Invoice says goods already shipped - make sure receiving expects it")

    # Quantity
    r["qty_var"] = r["conf_qty"] - pl["qty"]
    if r["qty_var"] < 0:
        r["issues"].append("QTY_SHORT")
        r["impact"] = max(r["impact"], -r["qty_var"] * pl["price"])
        r["notes"].append("Short %d (%.0f%%)" % (-r["qty_var"], -100.0 * r["qty_var"] / pl["qty"]))
    elif r["qty_var"] > 0:
        r["issues"].append("QTY_OVER")
        r["notes"].append("Over by %d" % r["qty_var"])

    # Price
    if not prices:
        r["issues"].append("NO_PRICE")
    else:
        p = prices[-1]
        rate = fx_rate(fx, p["currency"], pl["po_date"])
        r["conf_price"], r["conf_ccy"] = p["unit_price"], p["currency"]
        r["conf_price_usd"] = p["unit_price"] * rate
        r["price_var"] = r["conf_price_usd"] - pl["price"]
        r["price_var_pct"] = 100.0 * r["price_var"] / pl["price"]
        if p["currency"] != "USD":
            flag = abs(r["price_var_pct"]) > PRICE_TOL_FX_PCT
        else:
            flag = abs(r["price_var"]) > PRICE_TOL_USD
        if flag:
            r["issues"].append("PRICE_UP" if r["price_var"] > 0 else "PRICE_DOWN")
            ext = abs(r["price_var"]) * (r["conf_qty"] or pl["qty"])
            r["impact"] = max(r["impact"], ext)
            r["notes"].append("Price %+.2f%% (%+.4f/ea = %s$%.0f on this line)" % (
                r["price_var_pct"], r["price_var"], "+" if r["price_var"] > 0 else "-", ext))
        elif p["currency"] != "USD":
            r["notes"].append("%s %.4f x %.4f (ERP fx, PO month) = $%.2f, %+.2f%% vs PO - within FX tolerance" % (
                p["currency"], p["unit_price"], rate, r["conf_price_usd"], r["price_var_pct"]))

    # Dates
    if r["promise"]:
        r["date_var"] = (r["promise"] - pl["required"]).days
        if r["date_var"] > 0:
            r["days_late"] = r["date_var"]
            r["issues"].append("LATE")
            r["notes"].append("Promised %d days after required" % r["days_late"])
    for h in hits:
        if h[1].get("promise_note"):
            r["issues"].append("VAGUE_DATE")
            r["notes"].append(h[1]["promise_note"])
            break
    for d in docs_used:
        if d["doc_date"] and d["doc_date"] < pl["po_date"]:
            r["issues"].append("DATE_ODD")
            r["notes"].append("%s is dated %s, %d day(s) BEFORE the PO date %s - wrong PO reference, or PO re-issued?" % (
                d["file"], fmt_d(d["doc_date"]), (pl["po_date"] - d["doc_date"]).days, fmt_d(pl["po_date"])))
        if r["promise"] and d["doc_date"] and r["promise"] < d["doc_date"]:
            r["issues"].append("DATE_ODD")
            r["notes"].append("Promise %s is before the confirmation was sent (%s) - already shipped, or a typo?" % (
                fmt_d(r["promise"]), fmt_d(d["doc_date"])))


def _document_exceptions(docs, pos, po_lines, results, exceptions, review):
    for d in docs:
        if d["doc_type"] == "UNREADABLE":
            d["status"] = "NOT READ - manual"
            exceptions.append(dict(sev=1, code="UNREADABLE_DOCUMENT", vendor=d.get("vendor_id") or "?",
                                   po=d.get("po_number") or "?", file=d["file"], issue="Could not read document",
                                   detail="; ".join(d["warnings"]), action="Open the PDF and reconcile by hand."))
        elif d["po_number"] not in pos:
            d["status"] = "PO not on open list"
            same_part = [pl for pl in po_lines if pl["vendor_id"] == d["vendor_id"]
                         and any(dl.get("vendor_pn") == pl["our_pn"] for dl in d["lines"])]
            hint = ""
            unconf = [p for p in same_part
                      if set(results[(p["po_number"], p["line_number"])]["issues"]) & {"DROPPED", "NO_CONF"}]
            if unconf:
                hint += " Possibly meant for %s (not confirmed on its own ack)." % ", ".join(
                    "%s L%d qty %d" % (p["po_number"], p["line_number"], p["qty"]) for p in unconf)
            if same_part:
                hint += " Other open %s lines for this part: %s." % (d["vendor_id"], ", ".join(
                    "%s L%d qty %d" % (p["po_number"], p["line_number"], p["qty"]) for p in same_part))
            exceptions.append(dict(
                sev=1, code="PO_NOT_OPEN", vendor=d["vendor_id"], po=d["po_number"], file=d["file"],
                issue="Confirmation for a PO that is NOT on the open PO list",
                detail="; ".join("%s qty %d @ %s, ship %s" % (dl.get("vendor_pn") or dl.get("description"), dl["qty"],
                                                             dl.get("unit_price"), fmt_d(dl.get("promise_date")))
                                 for dl in d["lines"]) + "." + hint,
                action="Tell vendor this is not an open Beacon PO - don't ship against it. Check for a typo'd PO number or a closed/cancelled order."))
        if d["text_source"] in ("ocr", "ai") and d["doc_type"] != "UNREADABLE":
            review.append(dict(kind="OCR spot-check" if d["text_source"] == "ocr" else "AI-read document", vendor=d["vendor_id"], po=d["po_number"], line="",
                               source_pn="", candidate="", description="; ".join(
                                   "%s qty %s promise %s" % (dl.get("description"), dl.get("qty"), fmt_d(dl.get("promise_date")))
                                   for dl in d["lines"]),
                               method="Tesseract OCR" if d["text_source"] == "ocr" else d["template"],
                               confidence=MEDIUM if d["text_source"] == "ocr" else LOW, file=d["file"],
                               reason="Scanned document; numbers were read by OCR." if d["text_source"] == "ocr"
                               else "No template for this layout; an AI model read it. Nothing from it is trusted until checked.",
                               recommended="Glance at the PDF and confirm qty/date match what's shown here."))
        for dl in d["lines"]:
            if dl.get("promise_note"):
                review.append(dict(kind="Date interpretation", vendor=d["vendor_id"], po=d["po_number"], line=dl.get("line_no"),
                                   source_pn=dl.get("vendor_pn"), candidate=fmt_d(dl.get("promise_date")),
                                   description=dl.get("promise_note"), method="ISO calendar week", confidence=MEDIUM,
                                   file=d["file"], reason="Vendor gave a calendar-week range; we assume the latest day.",
                                   recommended="Ask vendor for a specific ship date."))


# --------------------------------------------------------------------------
# Plain-English action text
# --------------------------------------------------------------------------

def action_text(r):
    i, v = r["issues"], r["vendor_name"]
    parts = []
    if "DROPPED" in i:
        parts.append("Ask %s whether they will ship line %d (%s x %d) - it is not on their acknowledgment. "
                     "If not, receiving will be short." % (v, r["line_number"], r["our_pn"], r["qty"]))
    if "DECLINED" in i:
        parts.append("%s says it cannot supply line %d (%s x %d). Find another source or re-plan, and tell planning today." % (
            v, r["line_number"], r["our_pn"], r["qty"]))
    if "FORM_INCOMPLETE" in i:
        parts.append("Send the form back to %s: line %d is incomplete." % (v, r["line_number"]))
    if "NO_CONF" in i:
        parts.append("Chase %s for an acknowledgment." % v)
    if "NO_DETAIL" in i:
        parts.append("Ask for a firm qty and ship date.")
    if "QTY_SHORT" in i:
        parts.append("Ask whether the %d-pc balance ships later or is cancelled; re-order if needed." % (r["qty"] - r["conf_qty"]))
    if "QTY_OVER" in i:
        parts.append("Confirm we only need %d; refuse the overage." % r["qty"])
    if "PRICE_UP" in i:
        parts.append("Push back: PO price $%.4f, confirmed $%.4f. Get it corrected before the invoice reaches AP." % (
            r["price"], r["conf_price_usd"]))
    if "PRICE_DOWN" in i:
        parts.append("Price lower than PO - confirm it's intended and update the PO.")
    if "LATE" in i:
        parts.append("Promise is %d days after we need it - ask to expedite, or warn planning." % r["days_late"])
    if "DATE_ODD" in i:
        parts.append("Ask vendor to re-confirm the PO number and ship date.")
    if "VAGUE_DATE" in i and "LATE" not in i:
        parts.append("Ask for a specific ship date.")
    if "PN_INFERRED" in i:
        parts.append("Confirm the vendor's part number (Review Required tab).")
    if "INVOICE_ONLY" in i:
        parts.append("Tell receiving to expect it; hold the invoice until received.")
    return " ".join(parts) or "No action."


# --------------------------------------------------------------------------

def write_extracted_csv(path, extracted):
    cols = ["po_number", "line_no", "vendor_pn", "part_id", "confirmed_qty", "confirmed_price", "currency",
            "promised_date", "doc_date", "source_file", "match_method", "match_confidence"]
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for e in extracted:
            w.writerow({k: (e[k].isoformat() if isinstance(e[k], date) else e[k]) for k in cols})


def read_document(path, vendors=None, manual=None):
    """One place that turns a file into a normalized doc:
    template parser (+OCR) -> optional AI fallback -> lines typed in by hand."""
    from parsers import parse_file
    try:
        d = parse_file(path)
    except Exception as e:  # corrupt file etc. - never kill the run
        d = dict(vendor_id=None, template="-", doc_type="UNREADABLE", po_number=None, doc_date=None,
                 vendor_ref=None, part_of=None, supersedes_prior=False, lines=[], text_source="-",
                 raw_text="", warnings=["Could not open file: %s" % e])
    if d["doc_type"] == "UNREADABLE" and manual:
        d = manual_doc(manual, d.get("raw_text", ""))
    elif d["doc_type"] == "UNREADABLE":
        import ai_reader
        if ai_reader.enabled():
            ai = ai_reader.extract(path, vendors)
            if ai:
                ai["raw_text"] = d.get("raw_text", "")
                d = ai
    d["file"] = os.path.basename(path)
    validate_doc(d)
    return d


def manual_doc(m, raw_text=""):
    """A doc built from lines Lisa typed in for a PDF nothing could read."""
    def dt(s):
        return datetime.strptime(s, "%Y-%m-%d").date() if s else None
    return dict(vendor_id=m["vendor_id"], template="Entered by hand", doc_type=m.get("doc_type", "ACK"),
                po_number=m["po_number"], doc_date=dt(m.get("doc_date")), vendor_ref=m.get("vendor_ref"),
                part_of=None, supersedes_prior=False, text_source="manual", raw_text=raw_text,
                warnings=["Lines entered by hand by %s" % m.get("entered_by", "buyer")],
                lines=[dict(line_no=l.get("line_no"), vendor_pn=l.get("vendor_pn") or None, description=l.get("description"),
                            qty=float(l["qty"]), uom=None, unit_price=None if l.get("unit_price") in (None, "") else float(l["unit_price"]),
                            currency=l.get("currency") or "USD", promise_date=dt(l.get("promise_date")), promise_note=None)
                       for l in m["lines"]])


def run(confirmations, pos_csv, vendors_csv, erp=None, crosswalk=None, as_of=None, out=".", memory=None):
    """One daily run from the command line. Uses the same memory as the Desk app:
    1. read back Lisa's Done/Notes and part-number approvals from the last workbook
    2. register any PDFs in the folder we haven't seen (by content hash)
    3. run the shared check (workflow.check) and write today's workbook."""
    import workflow
    from report import write_workbook
    from store import Store, file_hash

    os.makedirs(out, exist_ok=True)
    store = Store(memory or os.path.join(out, "confirmation_desk.db"))
    store.seed_crosswalk_csv(crosswalk or os.path.join(HERE, "pn_crosswalk.csv"))
    prev = sorted(glob.glob(os.path.join(out, "PO_Confirmation_Check_*.xlsx")), key=os.path.getmtime)
    feedback = workflow.import_excel_feedback(store, prev[-1]) if prev else 0

    vendors = load_vendors(vendors_csv)
    parsed = {}
    known = store.known_hashes()
    new_docs = 0
    for f in sorted(glob.glob(os.path.join(confirmations, "*.pdf")) + glob.glob(os.path.join(confirmations, "*.xlsx"))):
        h = file_hash(f)
        if h in known:
            store.relocate_document(h, f)   # same file, maybe a new folder or another PC: read it from here
            continue
        parsed[f] = read_document(f, vendors)
        store.add_document(h, os.path.basename(f), f, parsed[f])
        known.add(h)
        new_docs += 1

    def reader(m, vendors, manual):
        if m["stored_path"] not in parsed or manual:
            parsed[m["stored_path"]] = read_document(m["stored_path"], vendors, manual=manual)
        return parsed[m["stored_path"]]

    day = datetime.strptime(as_of, "%Y-%m-%d").date() if as_of else None
    last = workflow.check(store, workflow.Paths(pos_csv, vendors_csv, erp), reader=reader, as_of=day)

    xlsx = os.path.join(out, "PO_Confirmation_Check_%s.xlsx" % last["as_of"].isoformat())
    stats = write_workbook(xlsx, last["results"], last["exceptions"], last["review"], last["suggestions"], last["docs"],
                           last["vendors"], store.crosswalk(), last["as_of"], item_status=workflow.item_status(store),
                           backlog=last.get("backlog"))
    stats = dict([("New documents this run", new_docs), ("Edits read back from last workbook", feedback),
                  ("Items auto-resolved since last run", last["resolved"])] + list(stats.items()))
    write_extracted_csv(os.path.join(out, "confirmation_lines_extracted.csv"), last["extracted"])
    store.export_crosswalk(os.path.join(out, "pn_crosswalk_updated.csv"))
    return dict(xlsx=xlsx, stats=stats, results=last["results"], exceptions=last["exceptions"], review=last["review"],
                suggestions=last["suggestions"], store=store)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--confirmations", required=True, help="folder of vendor acknowledgment PDFs")
    ap.add_argument("--pos", required=True, help="open PO list CSV")
    ap.add_argument("--vendors", required=True, help="vendor master CSV")
    ap.add_argument("--erp", help="ERP sqlite (optional; used for monthly FX rates)")
    ap.add_argument("--crosswalk", default=os.path.join(HERE, "pn_crosswalk.csv"))
    ap.add_argument("--as-of", help="YYYY-MM-DD (default: latest document date in the batch)")
    ap.add_argument("--out", default="output", help="where the workbook (and the memory file) go")
    ap.add_argument("--memory", help="memory file (default: <out>/confirmation_desk.db) - share it with the Desk app")
    a = ap.parse_args()
    out = run(a.confirmations, a.pos, a.vendors, a.erp, a.crosswalk, a.as_of, a.out, a.memory)
    for k, v in out["stats"].items():
        print("  %-42s %s" % (k, v))
    print("  -> %s" % out["xlsx"])


if __name__ == "__main__":
    main()
