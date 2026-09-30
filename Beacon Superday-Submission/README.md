# Beacon Fasteners: Confirmation Desk + Vendor Readout

Tinicum FDE super-day case study. Two deliverables share one engine:

| | For | What it is | Open this |
|---|---|---|---|
| **Task 1** | Lisa, senior buyer | **Confirmation Desk.** Reads vendor acknowledgment PDFs, checks them against the open PO list and produces a to-do list that remembers yesterday. Use it as an Excel workbook, or as a simple browser app. | `output/PO_Confirmation_Check_2026-05-17.xlsx` |
| **Task 2** | Plant manager | **Vendor performance readout** built from the ERP extract: received value by month, a vendor-by-vendor comparison, and which vendor to call first and why. | `output/vendor_readout.html` (or `output/Vendor_Readout.xlsx`) |
| **Vendor forms** | Vendors, via Lisa | **Fix it at the source.** Beacon's own Excel acknowledgment form per PO, and a supplier information pack per vendor, both pre-filled from the ERP. Returned forms load straight into the daily check. | `output/vendor_forms/` |
| Deck | Tinicum panel | The 30-minute walkthrough | `presentation/Beacon_FDE_Casestudy.pptx` |

---

## 1. Run it on the data (graders)

Needs Python 3.9+ and [Tesseract](https://tesseract-ocr.github.io/) for the scanned PDFs (`brew install tesseract` / `apt install tesseract-ocr`).

```bash
cd task1_confirmation_recon
python3 -m pip install -r requirements.txt

# Task 1 - the Excel workbook Lisa opens in the morning
python3 recon.py --confirmations /data/confirmations --pos /data/open_pos.csv \
                 --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output

# Task 2 - the plant manager's readout (workbook + one-page HTML)
cd ../task2_vendor_readout
python3 build_readout.py --erp /data/beacon_erp.db --out ../output

# Tests (59 checks; the end-to-end ones read the case data from BEACON_DATA and are skipped without it)
cd ../task1_confirmation_recon && BEACON_DATA=/data python3 -m unittest discover -s tests -v
```

The zip ships the generated outputs but not the tool's memory file, so the first run above starts clean and reads all 35 PDFs (the 4 scans take a few seconds each). Runs after that remember what they have seen.

`--erp` is optional for Task 1. Without it, the tool uses a default EUR rate and skips the vendor-history hints.

## 2. Use it day to day (Lisa)

**Option A: the browser app.** Double-click `task1_confirmation_recon/Start Confirmation Desk.command` (Mac) or `.bat` (Windows). The first run sets itself up, then your browser opens:

1. Drag the day's acknowledgment PDFs into the sidebar and click **Check them**.
2. **Today's list:** red means act today, orange means this week, blue means check. Each row says *what happened* and *what to do*. Tick **Done** and add a note; both are saved.
3. **Needs your OK:** new vendor part numbers (one click to approve), scanned pages shown next to what was read from them, and a form for any PDF nothing could read. The form is pre-filled from the PO, so it's mostly checking rather than typing.
4. **Chase & due soon:** the next 14 days, ranked by risk, plus POs waiting for an acknowledgment.
5. **Draft emails:** one ready-to-edit email per vendor to copy into Outlook. The tool never sends anything.
6. **Download Excel** at any time.
7. **Vendor forms:** download the acknowledgment form to attach to a PO, or a vendor's supplier pack; upload a returned pack to check it (section 4).

**Option B: Excel only.** Run `Run daily check (Excel only).command`, or the `recon.py` command above, each morning. Type `Yes` in **Done?** and add **Notes** in the workbook, and type `APPROVED` next to a part number on the **PN Crosswalk** tab. The next run reads those edits back, so handled items stay handled.

Both options share the same memory, so Excel and the app always agree.

### What "remembers" means

| Situation | What the tool does |
|---|---|
| Same PDF sent twice | Recognised by its content and ignored |
| You mark an item Done | Stays done tomorrow, with your note |
| Vendor sends a corrected acknowledgment | The item closes by itself, with the reason |
| A new problem appears on a line you'd closed | Reopens it and says why |
| Vendor keeps moving its promise date | The **first** promise is never overwritten, so the history stays honest |
| You approve a vendor part number once | It matches automatically from then on and joins the part-number master |

The `TestDayTwo` test runs exactly this sequence: day 1 on the real files, Lisa's edits in Excel, and a revised vendor PDF on day 2.

---

## 3. How it works

```
PDF ──► parsers.py (one parser per vendor layout; OCR for scans; optional AI for unknown layouts)
          │  normalized: vendor, PO, lines (part, qty, price, currency, promise date)
          ▼
recon.py  match each document line to a PO line:
          L1 Beacon part no. on the document ► L2 approved vendor-PN crosswalk ► L3 description
          ► L4 qty+price (MEDIUM confidence, goes to review) ► L5 unmatched (never guessed)
          then check: missing line · qty · price (EUR converted at the ERP rate, 1% tolerance)
          · date vs need · date sanity · revisions / split confirmations / duplicates
          ▼
workflow.py  daily check: update memory (store.py, one SQLite file) - to-do items,
             promise history, part-number decisions
          ▼
report.py (Excel)    desk_app.py (Streamlit)    ◄── same engine, two front doors
```

**Deliberately deterministic.** All 35 PDFs parse with template rules, so every number is traceable to a line of text, the run costs nothing, and nothing leaves the building. The AI reader (`ai_reader.py`, Claude with schema-validated output) is only a fallback for layouts nobody has written a parser for. It is **off** unless IT sets `BEACON_AI_FALLBACK=1`, because aerospace paperwork can be export-controlled. Whatever it reads always goes to *Needs your OK*.

**Adding a vendor layout** means writing one ~20-line parse function in `parsers.py` and adding it to `TEMPLATES`. Until then, that vendor's PDFs are flagged "could not read" (or typed in via the form), never silently skipped.

## 4. Past and future: using the ERP, and fixing discrepancies at the source

Reading six PDF layouts better every year is the wrong long-term goal. The durable fix is for vendors to answer in Beacon's own format, with the answers checked as they type. `vendor_forms.py` does this with plain Excel files, because vendors already work by email and it needs no portal, hosting or logins.

**The past (the ERP extract)** feeds everything below:

| From 8 months of ERP history | Used for |
|---|---|
| Per vendor + part: POs, last price, planned / quoted / actual lead time (medians) | Pre-filling the supplier pack, and the lead-time gap report for planning |
| Approved vendor part numbers (mined from confirmations) | Pre-filling "Your part no." so vendors confirm rather than type |
| 40 PO lines past due with quantity still open ($128k, incl. Apex $77k) | New **Old Open Balances** sheet in Lisa's workbook and on *Chase & due soon*: chase or short-close |
| Vendor on-time and promise-keeping history | Warnings next to today's promises (already in the desk) |

**The future: two forms.**

1. **PO acknowledgment form** (`output/vendor_forms/acknowledgment_forms/`, one per open PO; send it with the PO)
   - Grey cells are our PO and are locked. Rows can't be deleted, so a line can't vanish silently.
   - For every line the vendor chooses *Accept as ordered*, *Accept with changes* or *Cannot supply*, and gives **one real delivery date**. Excel refuses "KW 20-22" and "TBD".
   - A qty or price change needs a reason code (material surcharge, capacity, partial now, …).
   - Returned: drop it in the sidebar with the PDFs. It matches on Beacon part numbers at HIGH confidence, with no OCR.
   - *Cannot supply* becomes **VENDOR_DECLINED_LINE** (act today, with the vendor's reason) instead of "possible dropped line". A blank or invalid answer becomes **FORM_INCOMPLETE**, and the draft email asks for exactly that line back.
2. **Supplier information pack** (`output/vendor_forms/supplier_packs/`, once per vendor; refresh yearly)
   - Pre-filled with every part we have bought from that vendor, and our record of planned / quoted / actual lead time.
   - The vendor confirms its own part numbers, standard lead times, MOQ, price basis, surcharges, acknowledgment contact and turnaround, and certifications (AS9100 / Nadcap) with expiry dates.
   - Returned (*Vendor forms* page): the file is validated, and a reply email lists what is missing.
     - Declared part numbers go to *Needs your OK*, never straight to APPROVED. A part number given for two different parts is refused, and a renumbering is flagged against ERP history.
     - A **lead-time gap** report tells planning where MRP is wrong. Continental's standard is 28 days; we plan 19.

```bash
cd task1_confirmation_recon
python3 vendor_forms.py ack     --pos /data/open_pos.csv --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output/vendor_forms/acknowledgment_forms
python3 vendor_forms.py onboard --pos /data/open_pos.csv --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output/vendor_forms/supplier_packs
python3 vendor_forms.py check   ../output/vendor_forms/SYNTHETIC_vendor_replies/SYNTHETIC_Continental_supplier_pack.xlsx
```

`output/vendor_forms/SYNTHETIC_vendor_replies/` holds two **made-up** vendor replies for demos: Ostmark declining a line, and Continental returning its pack. They are not case data. Dropping the Ostmark one into the desk shows a *Vendor declined line* item.

**Rollout, and why not a portal yet:** start with the vendors that cause the most reading work (Continental's scans, Ostmark's German week ranges, Heritage's part numbers). PDFs keep working for everyone else, so it is not a big-bang change. Once most vendors use the form, the same fields become a small web form or an EDI 855 feed with no change to the engine.

## 5. What the tool found this week (35 PDFs, 44 open PO lines)

| | |
|---|---|
| Act today | Apex **dropped PO-4500050001 line 2** (1,500 pcs A286 bar, ≈$7.1k). Apex also sent a confirmation for **PO-4500060619**, which is not a Beacon PO. |
| This week | A 75-pc short (PO-007) · a +2.3% price increase (PO-002) · promises 21 / 16 / 14 / 7 days late (Continental, Ostmark, Heritage, Liberty split) · Liberty acknowledged PO-019 with no qty or date |
| Check | Revised QuickShip email (the original is superseded) · QuickShip sent an **invoice** in place of an acknowledgment · Continental promise dates earlier than the confirmation itself · documents dated before their PO · 2 renumbered vendor part numbers |
| Clean | 22 lines |

## 6. Vendor readout: headline

**Call Continental Quality Heat Treat first.**
- It is on time on 42% of lines; the other vendors range from 69% to 99%.
- Of all lines that arrived more than 7 days late, 61 of 63 are Continental's.
- It promises a date after our need date on 91% of lines.
- It has not improved over the eight months.
- Part of the problem is ours: our planned heat-treat lead time is 19 days, Continental quotes 28 and delivers in about 25.

Apex is second, for commercial reasons rather than delivery: 85% of spend, CRES bar prices up 1–2% since February on every confirmation while every PO stayed at $3.92, and $77k of short-shipped balances left open.

The ERP extract has traps that change the answer. Each was independently re-derived before use. They are listed in the readout and in the workbook's *Data Quality* tab.

## 7. Layout

```
task1_confirmation_recon/   parsers, recon, workflow, store, report, desk_app, ai_reader, vendor_forms, tests/, launchers
task2_vendor_readout/       analysis.py (metrics + data-trap handling), build_readout.py (workbook + HTML)
output/                     generated from the case data (+ output/vendor_forms/: forms to send vendors)
presentation/               deck + speaker notes
```

The case inputs (PDFs, CSVs, ERP extract) are in `../Beacon Superday-Candidate/data/` in this repository. The tests find them there, or set `BEACON_DATA`.
