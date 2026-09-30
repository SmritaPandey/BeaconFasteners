"""
Beacon Confirmation Desk - the daily tool for purchasing.

Drop in the day's vendor acknowledgments; get back a to-do list that remembers
yesterday. Run with:   streamlit run desk_app.py   (or double-click the launcher)

Where things live (DESK_HOME, default ./desk_data):
  desk.db         memory: documents, to-do items + notes, part-number decisions
  inbox/          every PDF received (kept, never modified)
  open_pos.csv    today's open-PO export from the ERP (replace it in Settings)
  vendor_master.csv, beacon_erp.db (optional, for vendor history)
"""
import json
import os
import shutil
import sys
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import recon   # noqa: E402
import report  # noqa: E402
import workflow  # noqa: E402
from store import DONE, OPEN, RESOLVED, Store, file_hash  # noqa: E402

HOME = os.environ.get("DESK_HOME", os.path.join(HERE, "desk_data"))
SAMPLE = os.environ.get("BEACON_SAMPLE_DATA") or ("/data" if os.path.isdir("/data/confirmations") else "")
INBOX, EXPORTS = os.path.join(HOME, "inbox"), os.path.join(HOME, "exports")
P_POS, P_VEND, P_ERP = (os.path.join(HOME, f) for f in ("open_pos.csv", "vendor_master.csv", "beacon_erp.db"))
SEV_LABEL = {1: "🔴 Today", 2: "🟠 This week", 3: "🔵 Check"}

st.set_page_config(page_title="Beacon Confirmation Desk", page_icon="📋", layout="wide")
st.markdown("""<style>
.block-container {padding-top: 1.6rem; max-width: 1400px}
div[data-testid="stMetricValue"] {font-size: 1.9rem}
.small {font-size: .88rem; color: #6b7280}
</style>""", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# Setup (first run copies the sample data if provided)
# --------------------------------------------------------------------------

def bootstrap():
    for d in (HOME, INBOX, EXPORTS):
        os.makedirs(d, exist_ok=True)
    for src, dst in (("open_pos.csv", P_POS), ("vendor_master.csv", P_VEND), ("beacon_erp.db", P_ERP)):
        if not os.path.exists(dst) and SAMPLE and os.path.exists(os.path.join(SAMPLE, src)):
            shutil.copy(os.path.join(SAMPLE, src), dst)


@st.cache_resource
def get_store():
    bootstrap()
    s = Store(os.path.join(HOME, "desk.db"))
    s.seed_crosswalk_csv(os.path.join(HERE, "pn_crosswalk.csv"))
    return s


@st.cache_data(show_spinner=False)
def parse_cached(path, h, manual_json, vendors_json):
    """Reading PDFs (esp. OCR) is the slow part - do each file once."""
    d = recon.read_document(path, json.loads(vendors_json), manual=json.loads(manual_json) if manual_json else None)
    d["hash"] = h
    return d


@st.cache_data(show_spinner=False)
def vendor_history(erp_path, mtime):
    return workflow.vendor_history(erp_path)


def inputs_ok():
    return os.path.exists(P_POS) and os.path.exists(P_VEND)


# --------------------------------------------------------------------------
# The check: read every stored document, reconcile, update the to-do list
# --------------------------------------------------------------------------

def run_check(store):
    """The shared daily check (workflow.py) - same logic as the command line."""
    def reader(m, vendors, manual):
        return parse_cached(m["stored_path"], m["file_hash"], json.dumps(manual) if manual else "", json.dumps(vendors))
    hist = vendor_history(P_ERP, os.path.getmtime(P_ERP) if os.path.exists(P_ERP) else 0)
    st.session_state["last"] = workflow.check(store, workflow.Paths(P_POS, P_VEND, P_ERP), reader=reader, hist=hist)
    return st.session_state["last"]


def add_files(store, files):
    known = store.known_hashes()
    added, dup = [], []
    for f in files:
        data = f.getvalue()
        path = os.path.join(INBOX, "tmp_" + f.name)
        with open(path, "wb") as fh:
            fh.write(data)
        h = file_hash(path)
        if h in known:
            os.remove(path)
            dup.append(f.name)
            continue
        final = os.path.join(INBOX, "%s_%s" % (datetime.now().strftime("%Y%m%d-%H%M%S"), f.name))
        os.replace(path, final)
        d = recon.read_document(final, recon.load_vendors(P_VEND))
        store.add_document(h, f.name, final, d)
        known.add(h)
        added.append(f.name)
    return added, dup


def pdf_images(path, dpi=110):
    import fitz
    doc = fitz.open(path)
    return [p.get_pixmap(dpi=dpi).tobytes("png") for p in doc]


def doc_path(store, file_name):
    for m in store.documents():
        if m["file_name"] == file_name:
            return m["stored_path"]
    return None


# --------------------------------------------------------------------------
# Pages
# --------------------------------------------------------------------------

def page_today(store, last):
    vendors = last["vendors"]
    items = store.items(OPEN)
    st.subheader("Today's list")
    c = st.columns(5)
    c[0].metric("Act today", sum(1 for i in items if i["severity"] == 1))
    c[1].metric("This week", sum(1 for i in items if i["severity"] == 2))
    c[2].metric("Check", sum(1 for i in items if i["severity"] == 3))
    c[3].metric("New since last check", last["new"])
    c[4].metric("Fixed by vendors since last check", last["resolved"])

    if not items:
        st.success("Nothing open. Every open PO line matches its confirmation.")
        return
    df = pd.DataFrame([{
        "key": i["item_key"], "Done": False, "New": "🆕" if i["first_seen"] == last["as_of"].isoformat() else "",
        "Priority": SEV_LABEL.get(i["severity"], ""), "Vendor": vendors.get(i["vendor_id"], {}).get("vendor_name", i["vendor_id"]),
        "PO": i["po_number"], "Line": i["line_no"], "What happened": i["title"], "What to do": i["action"],
        "$ at stake": i["impact"] or None, "My note": i["note"] or "", "Open since": i["first_seen"], "Details": i["detail"],
        "PDF": i["source_files"]} for i in items])
    df["Line"] = df["Line"].astype("Int64")
    st.caption("Tick **Done** when handled and add a note if you like - it saves automatically and stays done tomorrow. "
               "Items also close on their own when a vendor sends a corrected confirmation.")
    edited = st.data_editor(
        df, key="today_editor", hide_index=True, width="stretch", height=min(620, 60 + 36 * len(df)),
        disabled=[c for c in df.columns if c not in ("Done", "My note")],
        column_order=["Done", "New", "Priority", "Vendor", "PO", "Line", "What happened", "What to do", "$ at stake",
                      "My note", "Open since", "Details", "PDF"],
        column_config={
            "Done": st.column_config.CheckboxColumn("Done", width="small"),
            "New": st.column_config.TextColumn("", width="small"),
            "What to do": st.column_config.TextColumn(width="large"),
            "What happened": st.column_config.TextColumn(width="medium"),
            "$ at stake": st.column_config.NumberColumn(format="$%.0f", width="small"),
            "My note": st.column_config.TextColumn(width="medium"),
            "Details": st.column_config.TextColumn(width="large"),
        })
    changed = False
    for (_, a), (_, b) in zip(df.iterrows(), edited.iterrows()):
        if b["Done"] and not a["Done"]:
            store.set_item(a["key"], status=DONE, note=b["My note"])
            changed = True
        elif b["My note"] != a["My note"]:
            store.set_item(a["key"], note=b["My note"])
    if changed:
        st.toast("Saved. Marked done.")
        st.rerun()

    left, right = st.columns([1, 1])
    with left:
        with st.expander("✉️ Draft emails to vendors (copy, edit, send from Outlook)"):
            skip = {(i["po_number"], i["line_no"]) for i in store.items(DONE)}
            emails = report.draft_emails(last["results"], last["exceptions"], vendors,
                                         buyer=store.settings()["buyer_name"], skip_keys=skip)
            if not emails:
                st.write("No vendor emails needed.")
            for vid, (to, body) in emails.items():
                st.markdown("**%s** &nbsp; <span class='small'>%s</span>" % (vendors[vid]["vendor_name"], to), unsafe_allow_html=True)
                st.code(body, language=None)
    with right:
        with st.expander("📄 Show the vendor's PDF for an item"):
            files = sorted({f for i in items for f in (i["source_files"] or "").split(", ") if f})
            pick = st.selectbox("Document", files, key="pdf_pick")
            p = doc_path(store, pick) if pick else None
            if p:
                for img in pdf_images(p):
                    st.image(img, width="stretch")

    done = store.items(DONE)[:15] + store.items(RESOLVED)[:15]
    if done:
        with st.expander("Recently closed (%d)" % len(done)):
            st.dataframe(pd.DataFrame([{"Status": i["status"], "Why": i["status_reason"], "PO": i["po_number"], "Line": i["line_no"],
                                        "What it was": i["title"], "Note": i["note"]} for i in done]),
                         hide_index=True, width="stretch")


def page_review(store, last):
    st.subheader("Needs your OK")
    st.caption("The tool never guesses silently. These are the things it isn't sure about.")
    vendors = last["vendors"]
    cw = store.crosswalk()
    pending = [x for x in cw.values() if x["review_status"].startswith("PENDING")]

    st.markdown("#### Vendor part numbers it hasn't seen before")
    if not pending:
        st.write("None waiting.")
    for x in pending:
        old = [y["vendor_pn"] for y in cw.values() if y["vendor_id"] == x["vendor_id"] and y["beacon_pn"] == x["beacon_pn"]
               and y["vendor_pn"] != x["vendor_pn"] and y["review_status"].startswith("APPROVED")]
        with st.container(border=True):
            a, b = st.columns([3, 2])
            a.markdown("**%s** calls a part **`%s`**. The tool thinks it is our **`%s`** (%s)." % (
                vendors.get(x["vendor_id"], {}).get("vendor_name", x["vendor_id"]), x["vendor_pn"], x["beacon_pn"], x["description"]))
            a.markdown("<span class='small'>Why: %s. Seen %s time(s), last %s.%s</span>" % (
                x["mapping_source"], x["times_seen"], x["last_seen"],
                " Our history has this part from them as <b>%s</b> - they may have renumbered." % ", ".join(old) if old else ""),
                unsafe_allow_html=True)
            parts = sorted({p["our_pn"] for p in last["po_lines"] if p["vendor_id"] == x["vendor_id"]} | {x["beacon_pn"]})
            choice = b.selectbox("Beacon part", parts, index=parts.index(x["beacon_pn"]), key="pn_" + x["vendor_pn"])
            c1, c2 = b.columns(2)
            if c1.button("✅ Yes, use this", key="ok_" + x["vendor_pn"], width="stretch"):
                store.review_mapping(x["vendor_id"], x["vendor_pn"], True, store.settings()["buyer_name"], beacon_pn=choice)
                run_check(store)
                st.toast("Saved. %s will match automatically from now on." % x["vendor_pn"])
                st.rerun()
            if c2.button("❌ Not right", key="no_" + x["vendor_pn"], width="stretch"):
                store.review_mapping(x["vendor_id"], x["vendor_pn"], False, store.settings()["buyer_name"])
                run_check(store)
                st.rerun()

    st.markdown("#### Scanned or AI-read documents to glance at")
    ok = set(store.settings().get("checked_scans", []))
    scans = [d for d in last["docs"] if d.get("text_source") in ("ocr", "ai") and d["hash"] not in ok]
    if not scans:
        st.write("None waiting.")
    for d in scans:
        with st.container(border=True):
            a, b = st.columns([1, 1])
            p = doc_path(store, d["file"])
            if p:
                a.image(pdf_images(p, dpi=90)[0], width="stretch")
            b.markdown("**%s** &nbsp; %s &nbsp; %s" % (d["file"], d.get("po_number") or "", "(read by OCR)" if d["text_source"] == "ocr" else "(read by AI)"))
            b.dataframe(pd.DataFrame([{"Part / description": l.get("vendor_pn") or l.get("description"), "Qty": l["qty"],
                                       "Price": "not stated" if l.get("unit_price") is None else "%.4f" % l["unit_price"], "Promise": recon.fmt_d(l.get("promise_date"))} for l in d["lines"]]),
                        hide_index=True, width="stretch")
            if b.button("👍 Matches the PDF", key="scan_" + d["hash"]):
                store.save_setting("checked_scans", sorted(ok | {d["hash"]}))
                st.rerun()

    st.markdown("#### Documents nothing could read")
    bad = [d for d in last["docs"] if d["doc_type"] == "UNREADABLE"]
    if not bad:
        st.write("None - every document was read.")
    for d in bad:
        with st.container(border=True):
            manual_entry(store, d, last)


def manual_entry(store, d, last):
    """Type the lines in by hand - prefilled from the PO so it's mostly checking, not typing."""
    vendors = last["vendors"]
    st.markdown("**%s** - %s" % (d["file"], "; ".join(d["warnings"])))
    p = doc_path(store, d["file"])
    a, b = st.columns([1, 1])
    if p:
        a.image(pdf_images(p, dpi=90)[0], width="stretch")
    vid = b.selectbox("Vendor", list(vendors), format_func=lambda v: vendors[v]["vendor_name"], key="mv_" + d["hash"],
                      index=list(vendors).index(d["vendor_id"]) if d.get("vendor_id") in vendors else 0)
    pos = sorted({l["po_number"] for l in last["po_lines"] if l["vendor_id"] == vid})
    po = b.selectbox("PO", pos, key="mp_" + d["hash"], index=pos.index(d["po_number"]) if d.get("po_number") in pos else 0)
    dd = b.date_input("Document date", value=last["as_of"], key="md_" + d["hash"])
    pl = [l for l in last["po_lines"] if l["po_number"] == po]
    grid = pd.DataFrame([{"line_no": l["line_number"], "vendor_pn": l["our_pn"], "description": l["description"], "qty": l["qty"],
                          "unit_price": l["price"], "currency": "USD", "promise_date": l["required"].isoformat()} for l in pl])
    b.caption("Pre-filled from our PO. Change only what the vendor's document says differently; delete lines they didn't confirm.")
    g = b.data_editor(grid, num_rows="dynamic", hide_index=True, key="mg_" + d["hash"], width="stretch")
    if b.button("Save these lines", key="ms_" + d["hash"], type="primary"):
        store.save_manual(d["hash"], dict(vendor_id=vid, po_number=po, doc_date=dd.isoformat(), doc_type="ACK",
                                          entered_by=store.settings()["buyer_name"], lines=g.to_dict("records")),
                          store.settings()["buyer_name"])
        run_check(store)
        st.toast("Saved and re-checked.")
        st.rerun()


def page_ahead(store, last):
    st.subheader("Chase & due soon")
    cfg = store.settings()
    as_of, hist, vendors = last["as_of"], last["hist"], last["vendors"]
    horizon = as_of + timedelta(days=int(cfg["due_soon_days"]))
    rows = []
    for r in last["results"].values():
        if r["required"] > horizon:
            continue
        h = hist.get(r["vendor_id"], {})
        risk, why = "Low", []
        if r["conf_qty"] is None:
            risk = "High"
            why.append("not confirmed")
        if r["days_late"]:
            risk = "High"
            why.append("promised %d days late" % r["days_late"])
        if "QTY_SHORT" in r["issues"] or "DROPPED" in r["issues"]:
            risk = "High"
            why.append("short / dropped")
        if h and r["promise"] and not r["days_late"] and h["meets_promise"] < 0.8:
            risk = "Medium" if risk == "Low" else risk
            why.append("vendor met only %.0f%% of past promises" % (100 * h["meets_promise"]))
        if r["days_to_required"] < 0 and risk == "Low":
            risk = "Medium"
            why.append("past due - check with receiving")
        rows.append({"rank": {"High": 0, "Medium": 1, "Low": 2}[risk],
                     "Risk": {"High": "🔴 High", "Medium": "🟠 Medium", "Low": "🟢 Low"}[risk], "Required": r["required"],
                     "Days left": r["days_to_required"], "Vendor": r["vendor_name"], "PO": r["po_number"], "Line": r["line_number"],
                     "Part": r["our_pn"], "Qty": r["qty"], "Promised": r["promise"], "Why": ", ".join(why) or "confirmed on time"})
    st.markdown("**Due in the next %d days (or already past due)** as of %s" % (cfg["due_soon_days"], recon.fmt_d(as_of)))
    if rows:
        st.dataframe(pd.DataFrame(rows).sort_values(["rank", "Required"]).drop(columns="rank"), hide_index=True, width="stretch")
    else:
        st.write("Nothing due in that window.")
    st.caption("Past-due lines are shown with negative days left - check with receiving whether they arrived; this tool "
               "only sees confirmations, not receipts.")

    st.markdown("**Waiting for an acknowledgment**")
    wait = [{"Vendor": r["vendor_name"], "PO": r["po_number"], "Line": r["line_number"], "PO date": r["po_date"],
             "Days waiting": (as_of - r["po_date"]).days,
             "Status": "Chase now" if (as_of - r["po_date"]).days >= int(cfg["chase_after_days"]) else "Normal wait"}
            for r in last["results"].values() if "NO_CONF" in r["issues"]]
    if wait:
        st.dataframe(pd.DataFrame(wait).sort_values("Days waiting", ascending=False), hide_index=True, width="stretch")
    else:
        st.write("Every open PO has an acknowledgment on file.")


def page_all(store, last):
    st.subheader("All open PO lines")
    rows = []
    for r in sorted(last["results"].values(), key=lambda r: (r["vendor_name"], r["po_number"], r["line_number"])):
        rows.append({"Status": {1: "🔴", 2: "🟠", 3: "🔵"}.get(r["sev"], "✅"), "Vendor": r["vendor_name"], "PO": r["po_number"],
                     "Line": r["line_number"], "Part": r["our_pn"], "Vendor PN": r["vendor_pn"], "Ordered": r["qty"],
                     "Confirmed": r["conf_qty"], "PO price": r["price"], "Confirmed $": r["conf_price_usd"],
                     "Required": r["required"], "Promised": r["promise"],
                     "Issues": ", ".join(dict.fromkeys(recon.ISSUE[i][1] for i in r["issues"])) or "MATCH",
                     "Matched by": r["match"], "Confidence": r["confidence"], "PDF": r["files"]})
    df = pd.DataFrame(rows)
    q = st.text_input("Filter (vendor, PO, part...)", key="all_filter")
    if q:
        df = df[df.apply(lambda row: q.lower() in " ".join(map(str, row.values)).lower(), axis=1)]
    st.dataframe(df, hide_index=True, width="stretch", height=560,
                 column_config={"PO price": st.column_config.NumberColumn(format="$%.4f"),
                                "Confirmed $": st.column_config.NumberColumn(format="$%.4f")})
    st.markdown("**Documents on file**")
    st.dataframe(pd.DataFrame([{"File": m["file_name"], "Vendor": m["vendor_id"], "PO": m["po_number"], "Type": m["doc_type"],
                                "Doc date": m["doc_date"], "Read via": m["read_via"], "Status": m["status"], "Received": m["received_at"]}
                               for m in store.documents()]), hide_index=True, width="stretch")


def page_history(store, last):
    st.subheader("Vendor history")
    hist, vendors = last["hist"], last["vendors"]
    if not hist:
        st.info("Add the ERP history file (beacon_erp.db) in Settings to see how each vendor has performed.")
        return
    st.caption("From ERP purchasing history. Used above to warn you when a vendor's promise is less reliable than it looks.")
    st.dataframe(pd.DataFrame([{"Vendor": vendors.get(v, {}).get("vendor_name", v), "Lines delivered": h["due"],
                                "On time vs our need": h["on_time"], "Kept its own promise": h["meets_promise"],
                                "Typical delay when late (days)": h["median_late"], "Lines >7 days late": h["severe"],
                                "Promises later than we need": h["promised_late"], "Lead time we allow": h["lt_allowed"],
                                "Lead time they quote": h["lt_quoted"], "Lead time actual": h["lt_actual"]} for v, h in hist.items()]),
                 hide_index=True, width="stretch",
                 column_config={k: st.column_config.ProgressColumn(k, format="percent", min_value=0, max_value=1)
                                for k in ("On time vs our need", "Kept its own promise", "Promises later than we need")})


def page_settings(store, last):
    st.subheader("Settings")
    cfg = store.settings()
    a, b = st.columns(2)
    with a:
        st.markdown("**Today's open PO list** - export it from the ERP each morning and drop it here.")
        up = st.file_uploader("Open PO list (CSV)", type=["csv"], key="pos_up")
        if up is not None:
            df = pd.read_csv(up)
            need = {"po_number", "po_date", "vendor_id", "vendor_name", "line_number", "our_pn", "our_description",
                    "qty_ordered", "unit_price", "required_date"}
            missing = need - set(df.columns)
            if missing:
                st.error("That file is missing columns: %s. Export the 'Open PO lines' report again." % ", ".join(sorted(missing)))
            elif st.button("Use this open PO list", type="primary"):
                df.to_csv(P_POS, index=False)
                run_check(store)
                st.success("Loaded %d open PO lines and re-checked." % len(df))
        st.markdown("**Vendor list** (names and email addresses)")
        vup = st.file_uploader("Vendor master (CSV)", type=["csv"], key="vend_up")
        if vup is not None and st.button("Use this vendor list"):
            pd.read_csv(vup).to_csv(P_VEND, index=False)
            st.success("Vendor list updated.")
        st.markdown("**ERP history** (optional, for vendor history and exchange rates): %s" %
                    ("✅ loaded" if os.path.exists(P_ERP) else "not loaded - copy beacon_erp.db into %s" % HOME))
    with b:
        st.markdown("**Rules**")
        chase = st.number_input("Chase a vendor when there's no acknowledgment after (days)", 1, 30, int(cfg["chase_after_days"]))
        soon = st.number_input("'Due soon' looks ahead (days)", 3, 60, int(cfg["due_soon_days"]))
        tol = st.number_input("Allowed price difference for EUR vendors after exchange rate (%)", 0.0, 5.0,
                              float(cfg["price_tol_fx_pct"]), step=0.1)
        name = st.text_input("Buyer name (signs the draft emails)", cfg["buyer_name"])
        fixed = st.text_input("Check as of date (leave blank for today)", cfg.get("as_of", ""), help="YYYY-MM-DD")
        if st.button("Save rules", type="primary"):
            for k, v in (("chase_after_days", chase), ("due_soon_days", soon), ("price_tol_fx_pct", tol), ("buyer_name", name)):
                store.save_setting(k, v)
            store.save_setting("as_of", fixed.strip())
            run_check(store)
            st.success("Saved.")
    st.divider()
    st.markdown("**Exports**")
    c = st.columns(3)
    xw_path = store.export_crosswalk(os.path.join(EXPORTS, "pn_crosswalk.csv"))
    if os.path.exists(xw_path):
        c[0].download_button("Part-number crosswalk (CSV)", open(xw_path, "rb").read(), "pn_crosswalk.csv")
    path = os.path.join(EXPORTS, "confirmation_lines.csv")
    recon.write_extracted_csv(path, last["extracted"])
    c[1].download_button("Confirmed lines, ERP-ready (CSV)", open(path, "rb").read(), "confirmation_lines.csv")
    st.markdown("**Recent checks**")
    st.dataframe(pd.DataFrame(store.runs(10)), hide_index=True, width="stretch")
    st.caption("Everything is stored in %s. Back it up by copying that folder." % HOME)


def excel_bytes(store, last):
    status = workflow.item_status(store)
    path = os.path.join(EXPORTS, "PO_Confirmation_Check_%s.xlsx" % last["as_of"].isoformat())
    report.write_workbook(path, last["results"], last["exceptions"], last["review"], last["suggestions"], last["docs"],
                          last["vendors"], store.crosswalk(), last["as_of"], item_status=status)
    return open(path, "rb").read(), os.path.basename(path)


# --------------------------------------------------------------------------

def main():
    store = get_store()
    with st.sidebar:
        st.markdown("## 📋 Confirmation Desk")
        st.caption("Beacon Fasteners - Purchasing")
        if not inputs_ok():
            st.error("Load today's open PO list and the vendor list in Settings first.")
        files = st.file_uploader("Drop vendor confirmations (PDF)", type=["pdf"], accept_multiple_files=True, key="up_%d" % st.session_state.get("up_n", 0))
        if files and st.button("Check them", type="primary", width="stretch"):
            with st.spinner("Reading %d document(s)..." % len(files)):
                added, dup = add_files(store, files)
                run_check(store)
            st.session_state["up_n"] = st.session_state.get("up_n", 0) + 1
            st.session_state["flash"] = "Read %d new document(s)%s." % (len(added), "; %d already on file were skipped" % len(dup) if dup else "")
            st.rerun()
        if SAMPLE and not store.documents() and os.path.isdir(os.path.join(SAMPLE, "confirmations")):
            if st.button("Load this week's sample confirmations", width="stretch"):
                import glob
                class F:  # minimal stand-in for an uploaded file
                    def __init__(self, p): self.name, self.p = os.path.basename(p), p
                    def getvalue(self): return open(self.p, "rb").read()
                with st.spinner("Reading sample documents (scans take a few seconds)..."):
                    add_files(store, [F(p) for p in sorted(glob.glob(os.path.join(SAMPLE, "confirmations", "*.pdf")))])
                    run_check(store)
                st.rerun()
        page = st.radio("Go to", ["Today's list", "Needs your OK", "Chase & due soon", "All open POs", "Vendor history", "Settings"],
                        label_visibility="collapsed")

    if not inputs_ok():
        page_settings(store, dict(extracted=[]))
        return
    last = st.session_state.get("last") or run_check(store)
    with st.sidebar:
        st.caption("Checked as of **%s** · %d documents on file" % (recon.fmt_d(last["as_of"]), len(last["docs"])))
        if (date.today() - last["as_of"]).days > 1 and not store.settings().get("as_of"):
            st.caption("Dates shown as of the newest confirmation (sample data).")
        if last["docs"]:
            data, name = excel_bytes(store, last)
            st.download_button("⬇️ Download Excel", data, name, width="stretch")
        if st.button("Re-check", width="stretch"):
            run_check(store)
            st.rerun()
    if st.session_state.get("flash"):
        st.success(st.session_state.pop("flash"))
    if not last["docs"]:
        st.info("Drop the day's vendor confirmations in the sidebar to start.")
    {"Today's list": page_today, "Needs your OK": page_review, "Chase & due soon": page_ahead, "All open POs": page_all,
     "Vendor history": page_history, "Settings": page_settings}[page](store, last)


main()
