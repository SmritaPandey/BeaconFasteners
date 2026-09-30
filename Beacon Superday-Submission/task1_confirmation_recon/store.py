"""
Confirmation Desk memory: one SQLite file that remembers across days.

What it keeps
  documents   every confirmation PDF ever received (by content hash, so a
              re-sent PDF is recognised and not processed twice)
  manual_docs lines Lisa typed in by hand for a PDF the tool couldn't read
  items       every issue the tool has raised, with Lisa's status and notes.
              An item stays open across runs until Lisa marks it done or a
              newer confirmation fixes it (then it auto-resolves, with the reason)
  crosswalk   vendor part number -> Beacon part number, with who approved it
  promise_history  every qty/price/date a vendor ever confirmed per PO line,
              append-only. The FIRST promise is never overwritten, so a vendor
              that keeps pushing dates can't end up looking "on time".
  runs        a log of each check (when, how many docs, how many open items)
  vendor_profiles  what each vendor declared on Beacon's supplier information pack:
              contacts, acknowledgment turnaround, certifications, and per part their
              part number and standard lead time (vendor_forms.py)

Why SQLite: a single file, no server, backs up by copying, and the same
schema moves to SQL Server/Postgres when IT wants it on a shared server.
"""
import csv
import hashlib
import json
import os
import functools
import sqlite3
import threading
from datetime import datetime

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
  file_hash TEXT PRIMARY KEY, file_name TEXT, stored_path TEXT, received_at TEXT,
  vendor_id TEXT, po_number TEXT, doc_type TEXT, doc_date TEXT, read_via TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS manual_docs (
  file_hash TEXT PRIMARY KEY, payload TEXT, entered_by TEXT, entered_at TEXT);
CREATE TABLE IF NOT EXISTS items (
  item_key TEXT PRIMARY KEY, po_number TEXT, line_no INTEGER, vendor_id TEXT, issue_code TEXT,
  severity INTEGER, title TEXT, detail TEXT, action TEXT, impact REAL, source_files TEXT,
  first_seen TEXT, last_seen TEXT, status TEXT, status_reason TEXT, note TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS crosswalk (
  vendor_id TEXT, vendor_pn TEXT, beacon_pn TEXT, description TEXT, mapping_source TEXT,
  mapping_confidence TEXT, times_seen INTEGER, first_seen TEXT, last_seen TEXT,
  review_status TEXT, reviewed_by TEXT, reviewed_at TEXT, PRIMARY KEY (vendor_id, vendor_pn));
CREATE TABLE IF NOT EXISTS runs (
  run_at TEXT, as_of TEXT, docs_total INTEGER, docs_new INTEGER, open_items INTEGER, resolved_items INTEGER);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS vendor_profiles (
  vendor_id TEXT PRIMARY KEY, company TEXT, parts TEXT, problems TEXT, lead_time_gaps TEXT,
  file_name TEXT, received_at TEXT);
CREATE TABLE IF NOT EXISTS promise_history (
  po_number TEXT, line_no INTEGER, source_file TEXT, doc_date TEXT, confirmed_qty REAL,
  confirmed_price REAL, currency TEXT, promised_date TEXT, recorded_at TEXT,
  PRIMARY KEY (po_number, line_no, source_file));
"""

DEFAULT_SETTINGS = {
    "chase_after_days": 3,          # no acknowledgment this many days after PO date -> chase
    "due_soon_days": 14,            # look-ahead window for the "due soon" list
    "price_tol_fx_pct": 1.0,        # EUR/other-currency tolerance after FX
    "buyer_name": "Lisa",
}

OPEN, DONE, RESOLVED = "Open", "Done", "Resolved"


def now():
    return datetime.now().isoformat(timespec="seconds")


def file_hash(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def _locked(fn):
    """One connection is shared by the app's browser sessions (threads):
    serialise every call so two tabs can't trip over each other."""
    @functools.wraps(fn)
    def wrap(self, *a, **kw):
        with self._lock:
            return fn(self, *a, **kw)
    return wrap


class Store:
    def __init__(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self._lock = threading.RLock()
        self.con = sqlite3.connect(path, check_same_thread=False, timeout=30)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA journal_mode=WAL")   # readers don't block the writer
        self.con.executescript(SCHEMA)
        self.con.commit()

    # ---------------- settings ----------------
    @_locked
    def settings(self):
        s = dict(DEFAULT_SETTINGS)
        for r in self.con.execute("select key, value from settings"):
            s[r["key"]] = json.loads(r["value"])
        return s

    @_locked
    def save_setting(self, key, value):
        self.con.execute("insert or replace into settings values (?, ?)", (key, json.dumps(value)))
        self.con.commit()

    # ---------------- documents ----------------
    @_locked
    def known_hashes(self):
        return {r[0] for r in self.con.execute("select file_hash from documents")}

    @_locked
    def add_document(self, h, name, stored_path, d):
        self.con.execute("""insert or replace into documents values (?,?,?,?,?,?,?,?,?,?)""", (
            h, name, stored_path, now(), d.get("vendor_id"), d.get("po_number"), d.get("doc_type"),
            d["doc_date"].isoformat() if d.get("doc_date") else None, d.get("text_source"), d.get("status")))
        self.con.commit()

    @_locked
    def relocate_document(self, h, path):
        """The memory keeps the path a document was first read from. If the folder moved (or the
        memory file was copied to another machine), point it at where the file is now."""
        self.con.execute("update documents set stored_path=? where file_hash=? and stored_path<>?", (path, h, path))
        self.con.commit()

    @_locked
    def update_doc_status(self, h, status):
        self.con.execute("update documents set status=? where file_hash=?", (status, h))
        self.con.commit()

    @_locked
    def documents(self):
        return [dict(r) for r in self.con.execute("select * from documents order by received_at desc, file_name")]

    @_locked
    def save_manual(self, h, payload, who):
        self.con.execute("insert or replace into manual_docs values (?,?,?,?)", (h, json.dumps(payload), who, now()))
        self.con.commit()

    @_locked
    def manual(self, h):
        r = self.con.execute("select payload from manual_docs where file_hash=?", (h,)).fetchone()
        return json.loads(r[0]) if r else None

    # ---------------- crosswalk ----------------
    @_locked
    def seed_crosswalk_csv(self, path):
        """Load the ERP-derived crosswalk once; never overwrite Lisa's decisions."""
        if not os.path.exists(path):
            return 0
        n = 0
        with open(path, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                cur = self.con.execute("select 1 from crosswalk where vendor_id=? and vendor_pn=?",
                                       (r["vendor_id"], r["vendor_pn"])).fetchone()
                if cur:
                    continue
                self.con.execute("insert into crosswalk values (?,?,?,?,?,?,?,?,?,?,?,?)", (
                    r["vendor_id"], r["vendor_pn"], r["beacon_pn"], r.get("description", ""), r.get("mapping_source", ""),
                    r.get("mapping_confidence", ""), int(r.get("times_seen") or 0), r.get("first_seen", ""),
                    r.get("last_seen", ""), r.get("review_status", ""), "ERP history", now()))
                n += 1
        self.con.commit()
        return n

    @_locked
    def crosswalk(self):
        return {(r["vendor_id"], r["vendor_pn"]): dict(r) for r in self.con.execute("select * from crosswalk")}

    @_locked
    def suggest_mapping(self, s):
        cur = self.con.execute("select review_status, times_seen from crosswalk where vendor_id=? and vendor_pn=?",
                               (s["vendor_id"], s["vendor_pn"])).fetchone()
        if cur:
            if cur["review_status"].startswith("PENDING"):
                self.con.execute("update crosswalk set last_seen=?, times_seen=times_seen+1 where vendor_id=? and vendor_pn=?",
                                 (s["last_seen"], s["vendor_id"], s["vendor_pn"]))
        else:
            self.con.execute("insert into crosswalk values (?,?,?,?,?,?,?,?,?,?,?,?)", (
                s["vendor_id"], s["vendor_pn"], s["beacon_pn"], s.get("description", ""), s["mapping_source"],
                s["mapping_confidence"], 1, s["first_seen"], s["last_seen"], "PENDING REVIEW", None, None))
        self.con.commit()

    @_locked
    def review_mapping(self, vendor_id, vendor_pn, approve, who, beacon_pn=None):
        status = "APPROVED" if approve else "REJECTED"
        if beacon_pn:
            self.con.execute("update crosswalk set beacon_pn=? where vendor_id=? and vendor_pn=?", (beacon_pn, vendor_id, vendor_pn))
        self.con.execute("update crosswalk set review_status=?, reviewed_by=?, reviewed_at=? where vendor_id=? and vendor_pn=?",
                         (status, who, now(), vendor_id, vendor_pn))
        self.con.commit()

    @_locked
    def export_crosswalk(self, path):
        rows = [dict(r) for r in self.con.execute("select * from crosswalk order by vendor_id, vendor_pn")]
        if rows:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(rows[0]))
                w.writeheader()
                w.writerows(rows)
        return path

    # ---------------- vendor profiles (supplier information pack) ----------------
    @_locked
    def save_vendor_profile(self, res, file_name):
        """Keep the vendor's declarations; their part numbers go to the review queue, never
        straight to APPROVED (a vendor can be wrong about which of our parts is which)."""
        self.con.execute("insert or replace into vendor_profiles values (?,?,?,?,?,?,?)", (
            res["vendor_id"], json.dumps(res["company"], default=str), json.dumps(res["parts"], default=str),
            json.dumps(res["problems"]), json.dumps(res["lead_time_gaps"]), file_name, now()))
        day = now()[:10]
        for m in res["mappings"]:
            if not self.con.execute("select 1 from crosswalk where vendor_id=? and vendor_pn=?",
                                    (m["vendor_id"], m["vendor_pn"])).fetchone():
                self.con.execute("insert into crosswalk values (?,?,?,?,?,?,?,?,?,?,?,?)", (
                    m["vendor_id"], m["vendor_pn"], m["beacon_pn"], m.get("note", ""),
                    "Vendor-declared (supplier information pack %s)" % file_name, "HIGH", 1, day, day,
                    "PENDING REVIEW", None, None))
        self.con.commit()

    @_locked
    def vendor_profiles(self):
        out = {}
        for r in self.con.execute("select * from vendor_profiles"):
            d = dict(r)
            for k in ("company", "parts", "problems", "lead_time_gaps"):
                d[k] = json.loads(d[k] or "null")
            out[d["vendor_id"]] = d
        return out

    # ---------------- items (the to-do list that survives across days) ----------------
    @_locked
    def sync_items(self, current, as_of):
        """current: {item_key: dict(...)} from today's check.
        New -> Open.  Seen again -> refresh details (keep Lisa's status/note).
        Previously open but gone -> Resolved automatically, with the reason."""
        ts = now()
        existing = {r["item_key"]: dict(r) for r in self.con.execute("select * from items")}
        new = resolved = 0
        for k, it in current.items():
            if k in existing:
                e = existing[k]
                status = e["status"]
                reason = e["status_reason"]
                if status == RESOLVED:          # problem came back -> reopen
                    status, reason = OPEN, "Reappeared on %s" % as_of
                elif status == DONE and not set(it["issue_code"].split(",")) <= set((e["issue_code"] or "").split(",")):
                    status, reason = OPEN, "Reopened %s: new problem on this line (%s)" % (as_of, it["issue_code"])
                self.con.execute("""update items set issue_code=?, severity=?, title=?, detail=?, action=?, impact=?,
                                    source_files=?, last_seen=?, status=?, status_reason=?, updated_at=? where item_key=?""",
                                 (it["issue_code"], it["severity"], it["title"], it["detail"], it["action"], it["impact"],
                                  it["source_files"], as_of, status, reason, ts, k))
            else:
                new += 1
                self.con.execute("insert into items values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                    k, it["po_number"], it["line_no"], it["vendor_id"], it["issue_code"], it["severity"], it["title"],
                    it["detail"], it["action"], it["impact"], it["source_files"], as_of, as_of, OPEN, None, "", ts))
        for k, e in existing.items():
            if k not in current and e["status"] == OPEN:
                resolved += 1
                self.con.execute("update items set status=?, status_reason=?, updated_at=? where item_key=?",
                                 (RESOLVED, "No longer flagged as of %s (newer confirmation, or PO closed)" % as_of, ts, k))
        self.con.commit()
        return new, resolved

    @_locked
    def items(self, status=None):
        q = "select * from items"
        args = ()
        if status:
            q += " where status=?"
            args = (status,)
        return [dict(r) for r in self.con.execute(q + " order by severity, impact desc, po_number, line_no", args)]

    @_locked
    def set_item(self, key, status=None, note=None):
        if status is not None:
            self.con.execute("update items set status=?, status_reason=?, updated_at=? where item_key=?",
                             (status, "Marked %s by buyer" % status.lower(), now(), key))
        if note is not None:
            self.con.execute("update items set note=?, updated_at=? where item_key=?", (note, now(), key))
        self.con.commit()

    # ---------------- promise history ----------------
    @_locked
    def record_promise(self, e):
        iso = lambda d: d.isoformat() if hasattr(d, "isoformat") else d
        self.con.execute("insert or ignore into promise_history values (?,?,?,?,?,?,?,?,?)", (
            e["po_number"], e["line_no"], e["source_file"], iso(e["doc_date"]), e["confirmed_qty"], e["confirmed_price"],
            e["currency"], iso(e["promised_date"]), now()))
        self.con.commit()

    @_locked
    def promise_summary(self):
        """{(po, line): first_promise, latest_promise, changes} - ordered by the vendor's document date."""
        from datetime import date as _d
        out = {}
        for r in self.con.execute("""select po_number, line_no, doc_date, promised_date from promise_history
                                     where promised_date is not null order by po_number, line_no, doc_date, recorded_at"""):
            k = (r["po_number"], r["line_no"])
            p = _d.fromisoformat(r["promised_date"])
            if k not in out:
                out[k] = dict(first_promise=p, latest_promise=p, changes=0, _docs={r["doc_date"]: p})
            else:
                h = out[k]
                if r["doc_date"] not in h["_docs"] and p != h["latest_promise"]:
                    h["changes"] += 1
                h["_docs"][r["doc_date"]] = p
                h["latest_promise"] = p
        for h in out.values():
            h.pop("_docs")
        return out

    @_locked
    def log_run(self, as_of, total, new_docs, open_items, resolved):
        self.con.execute("insert into runs values (?,?,?,?,?,?)", (now(), as_of, total, new_docs, open_items, resolved))
        self.con.commit()

    @_locked
    def runs(self, n=10):
        return [dict(r) for r in self.con.execute("select * from runs order by run_at desc limit ?", (n,))]
