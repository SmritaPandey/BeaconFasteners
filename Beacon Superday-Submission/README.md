# Beacon Fasteners · PO Confirmation Control Tower & Vendor Readout

Tinicum FDE super-day case. The goal was not to "extract data from PDFs": it was to give Beacon's buyer a **daily, auditable action queue** and give the plant manager an **evidence-based answer about vendors**, without asking the plant to change systems first.

| Deliverable | For | Answers | Open |
|---|---|---|---|
| **Task 1: Confirmation Control Tower** | Lisa, senior buyer | *Which PO lines need my attention today, and what do I do?* | `output/PO_Confirmation_Check_2026-05-17.xlsx` |
| **Task 2: Vendor performance readout** | Plant manager | *Received value by month · how vendors compare · who to call first and why · which open orders to watch* | `output/Vendor_Readout.pdf` (also `.html`, `.xlsx`) |
| **Vendor forms** | Vendors, via Lisa | Beacon's own acknowledgment form and supplier pack, pre-filled from the ERP, so discrepancies are prevented, not just detected | `output/vendor_forms/` |
| **Presentation** | Tinicum panel | What I built and why, what I didn't, where it breaks, what I'd ask | `presentation/Beacon_FDE_Casestudy.pptx` |

---

## Two ways to run the same tool

| Mode | How | What it shows |
|---|---|---|
| **Without AI (default)** | `./run_all.sh /data` | Template parsers + OCR + deterministic analytics. Every number is traceable, and nothing leaves the building. |
| **With Claude (optional)** | `export ANTHROPIC_API_KEY=...`, then `recon.py --ai` and `task2_vendor_readout/ask.py "question"` | Claude reads PDFs no template can, and answers the plant manager's questions from the computed tables. Tested live, with results in `output/ai_evidence/` (section 6b). |

The deliverables and the numbers are identical in both modes. AI only adds coverage (new layouts, no OCR installed) and a way to ask questions.

## 1. Run it (graders)

**One command** (macOS/Linux). It checks prerequisites, installs into a local `.venv`, runs both tasks, generates the vendor forms and runs all tests:

```bash
./run_all.sh /data          # or the path to the case data folder
```

Windows: `run_all.bat C:\path\to\data`.

**Prerequisites:** Python 3.9+, and [Tesseract](https://tesseract-ocr.github.io/) for the 4 scanned PDFs:
- macOS: `brew install tesseract`
- Linux: `sudo apt install tesseract-ocr`
- Windows: the [UB-Mannheim installer](https://github.com/UB-Mannheim/tesseract/wiki). It is found in `C:\Program Files\Tesseract-OCR` automatically; elsewhere, set `BEACON_TESSERACT` to the `.exe`.

Without Tesseract everything still runs, and the scans are listed as "could not read" instead of being silently skipped.

Tested from a fresh unzip, in a path with spaces, on Python 3.9 and 3.13 (Linux). All dependencies install as prebuilt packages on Windows x64 and macOS (Apple Silicon and Intel).

<details><summary>Step by step instead</summary>

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r task1_confirmation_recon/requirements.txt

cd task1_confirmation_recon
python3 recon.py --confirmations /data/confirmations --pos /data/open_pos.csv \
                 --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output
python3 vendor_forms.py ack     --pos /data/open_pos.csv --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output/vendor_forms/acknowledgment_forms
python3 vendor_forms.py onboard --pos /data/open_pos.csv --vendors /data/vendor_master.csv --erp /data/beacon_erp.db --out ../output/vendor_forms/supplier_packs
BEACON_DATA=/data python3 -m unittest discover -s tests        # 61 tests

cd ../task2_vendor_readout
python3 build_readout.py --erp /data/beacon_erp.db --out ../output
BEACON_DATA=/data python3 -m unittest discover -s tests        # 9 tests
```
</details>

Expected on the case data: 35 PDFs read (31 text, 4 OCR), 44 open PO lines, 1 act-today line + 1 document-level exception, 22 clean lines, and both built-in checks `PASS`. Task 2 prints `Call first: Continental Quality Heat Treat`.

The zip ships the generated outputs but not the tool's memory file, so your first run starts clean. `Vendor_Readout.pdf` is produced only when Playwright + Chromium are available; otherwise print the HTML.

---

## 2. Task 1 · PO Confirmation Control Tower

**The decision:** which open PO lines need Lisa today, and what should she do? The hard part is what is *not* on the page: a line the vendor silently dropped. So the tool starts from the **open PO list** and looks for each line in the acknowledgments, never the other way round.

### Pipeline

```
PDF ──► parsers.py   one parser per vendor layout (6) · OCR for scans · German labels and "KW 20-22" week ranges
         │           → canonical record: vendor, PO, line, part, qty, price, currency, promise date, source file
         ▼
recon.py  tiered match of each document line to a PO line, then the checks below
         ▼
workflow.py  memory (store.py, one SQLite file): to-do items, Lisa's notes, promise history, part-number decisions
         ▼
report.py ──► Excel workbook          desk_app.py ──► optional browser app (same engine, same memory)
```

### Matching: automate the certain, route the uncertain to a human

| Level | Evidence | Confidence | What happens | This week |
|---|---|---|---|---|
| L1 | Beacon part number printed on the vendor document | HIGH | auto-match | 23 lines |
| L2 | approved (vendor, vendor part no.) crosswalk entry | HIGH | auto-match | 5 |
| L3 | normalized description equals the PO description | HIGH | auto-match | 9 |
| L4 | only qty + price (+ line no.) agree | MEDIUM | matched for the checks, **always** sent to *Review Required* | 5 |
| L5 | nothing fits, or two PO lines fit equally | none | left unmatched; never guessed | 0 |

84% of open lines matched automatically at HIGH confidence; every lower-confidence match is visible for review.

### Exception rules

| Issue type | Rule | Severity | Recommended action |
|---|---|---|---|
| `POSSIBLE_DROPPED_LINE` | vendor acknowledged the PO but this line is missing | Act today | Ask vendor whether it ships; warn receiving |
| `VENDOR_DECLINED_LINE` | vendor chose *Cannot supply* on Beacon's form | Act today | Re-source or re-plan; tell planning |
| `NO_CONFIRMATION` | no acknowledgment after the chase window (3 days, configurable) | Act today | Chase the vendor |
| `PO_NOT_OPEN` | acknowledgment for a PO that is not on the open list | Act today | Tell vendor not to ship; find the intended PO |
| `QUANTITY_SHORT` | confirmed qty < ordered | This week | Balance later, or re-order |
| `PRICE_MISMATCH` (up) | USD: any change; EUR: > 1% after ERP FX for the PO month | This week | Push back before the invoice reaches AP |
| `DATE_LATE` | promise after required date | This week | Expedite, or warn planning |
| `MISSING_REQUIRED_FIELD` | acknowledged with no qty/date | This week | Ask for a firm qty and date |
| `FORM_INCOMPLETE` | Beacon form line blank or date not a real date | This week | Send the form back (draft email lists the line) |
| `PART_NUMBER_UNMATCHED` | matched only at L4 | Check | Approve or reject the vendor part number |
| `DATE_INCONSISTENT` / `DATE_AMBIGUOUS` | document dated before the PO, promise before the document, or a week range | Check | Ask vendor to re-confirm |
| `QUANTITY_OVER`, `PRICE_MISMATCH` (down), `INVOICE_NOT_ACK` | as named | Check | Confirm, or tell receiving / AP |

Within a severity, the Action List is ordered by **money at stake** (line value for a dropped line, the variance × quantity for price/qty), then by required date. Revisions supersede earlier acknowledgments, split confirmations are summed, and a re-sent PDF is recognised by its content hash.

### The workbook: one tab per job

| Tab | For | Contents |
|---|---|---|
| **Action List** | Lisa | Only exceptions, worst first, each with *what happened* and *what to do*. `Done?` and `Notes` are read back on the next run |
| Review Required | Lisa | Every uncertain match, with the reason and the evidence |
| All PO Lines | Audit | One row per open PO line: ordered vs confirmed, match method, confidence, source file, first promise (never overwritten) |
| Dropped & Unconfirmed | Buyer / receiving | Lines with no usable confirmation |
| Price Variances | Buyer / AP | PO vs confirmed price, per unit and on the line |
| Date Variances | Buyer / planner | Required vs promised, days late, how often the vendor moved the date |
| Quantity Variances | Buyer / receiving | Short and over |
| Unmatched Vendor Lines | Master data | Unknown vendor part numbers, lines that fit no PO line, POs not open |
| Vendor Summary · Draft Emails | Buyer | Per-vendor counts, and one ready-to-edit email per vendor (never sent by the tool) |
| Documents · PN Crosswalk · Run Summary | Audit / master data | Every file with its status; the part-number crosswalk (type `APPROVED`); run statistics and the built-in checks |
| Old Open Balances | Buyer | From the ERP: 40 past-due lines with qty still open ($128k): chase or short-close |
| Read Me | Everyone | Refresh steps, rules, tolerances, limits, escalation |

### Day to day

- **Excel only:** run `Run daily check (Excel only).command` (or `recon.py`) each morning after exporting the open-PO list.
- **Browser app (optional):** `Start Confirmation Desk.command` / `.bat`. Drag in PDFs, approve part numbers in one click, see scans next to what was read, download the workbook. Both share one memory file, so they always agree.

| Situation | What the tool does |
|---|---|
| Same PDF sent twice | recognised by content and ignored |
| Lisa marks an item Done | stays done tomorrow, with her note |
| Vendor sends a corrected acknowledgment | the item closes by itself, with the reason |
| Vendor keeps moving its date | the **first** promise is kept, so the history stays honest |
| Lisa approves a vendor part number | it matches automatically from then on |

---

## 3. Task 2 · Vendor performance readout

The ERP extract was undocumented, so nothing was assumed from column names: the schema was profiled, joins were validated on samples, and every metric states its numerator, denominator and exclusions (*Definitions* tab).

**Headline:** call **Continental Quality Heat Treat** first. It was on time on 42% of 134 due lines (other vendors: 69–99%), accounts for 61 of the 63 lines that arrived more than 7 days late, promises after our need date on 91% of lines, and is not improving. Part of it is Beacon's: we plan 19 days for heat treat, Continental quotes 28 and delivers in about 25. **Apex** is the second call: commercial, not delivery. It is 85% of received value, CRES bar has been confirmed 1–2% above PO since February, and $77k of short-shipped balances are still open.

**Looking ahead** (`forward_risk.py`): 80 open lines ($3.32M) are due in the next 60 days.
- **3** are at high risk of arriving late.
- **26 lines ($1.84M)** due within 14 days have no acknowledgment on file.
- The flag rule, backtested on 8 months without look-ahead, would have flagged **59 of the 63** lines that arrived more than 7 days late. Roughly 1 in 3 flagged lines was late, so it is a watch list, not a forecast.

**Why no 0–100 vendor score:** the weights would be invented. Apex looks "worst" on late dollars only because it is 85% of spend; its late lines are late by a median of 1 day. The readout shows absolute and relative measures side by side, with denominators, and ranks the call by severity and business impact.

**Data traps found and handled** (each re-derived independently):
- receipt dates are `MM/DD/YYYY` text, so SQLite date functions return NULL;
- 50 keyed-in-error reversal pairs, plus 1 duplicate that was never reversed;
- `po_line.qty_received` is wrong on 169 lines;
- 30 confirmations were superseded;
- the confirmation feed stops on 2026-04-03;
- Ostmark's PO prices are already USD (converting again would add $76k);
- `qc_hold` dates are unreliable.

Taken at face value, received value would be misstated by between +$1.3M and −$2.9M. The corrected figure is **$28,187,145.87**.

**Refresh monthly:** `python3 build_readout.py --erp <new extract> --out <folder>`. If Task 1 has run, its extracted promise dates feed the forward view automatically.

---

## 4. Part numbers and vendor forms: build for the cleanup

Vendor part numbers are a **master-data problem**, not an OCR problem. The crosswalk (`pn_crosswalk.csv` → `output/pn_crosswalk_updated.csv`) is keyed on **(vendor, vendor part no.)**, never the part number alone: `K-1050` means two different parts at two vendors. Each row keeps its source (ERP history, inferred this week, vendor-declared, buyer-approved), confidence, times seen, first and last seen, review status, reviewer and date. New mappings are only ever **proposed**: a buyer approves them, and only approved rows are used for matching. Two vendor renumberings were caught this week against ERP history.

`vendor_forms.py` stops discrepancies at the source, using plain Excel because vendors already work by email:
- **PO acknowledgment form** (one per open PO, sent with it):
  - Our lines are locked, and rows can't be deleted.
  - Every line needs *Accept as ordered*, *Accept with changes* or *Cannot supply*.
  - Excel accepts only a real delivery date, and a change to qty or price needs a reason code.
  - A returned form loads into the daily check at HIGH confidence, with no OCR.
- **Supplier information pack** (once per vendor, pre-filled from the ERP):
  - The vendor confirms its part numbers (they go to review), standard lead times, contacts, acknowledgment turnaround and certification expiry.
  - The returned pack is validated, and a reply email lists what is missing.
  - A lead-time gap report tells planning where MRP is wrong.

`output/vendor_forms/SYNTHETIC_vendor_replies/` holds two **made-up** replies for demos (not case data).

---

## 5. Before trusting it in production

1. **Shadow week:** Lisa keeps her manual check; the tool runs alongside. Count misses (target: zero) and false alarms, and tune the tolerances with her.
2. **Coverage report:** the Run Summary shows per-run counts of text / OCR / unreadable documents, and the checks `every PDF has a status` and `every open PO line is accounted for` must say `PASS`.
3. **Spot-check the Review Required tab** and every OCR'd document against the PDF image. The browser app shows them side by side.
4. **Adding a vendor layout:** one ~20-line parse function in `parsers.py`, registered in `TEMPLATES`, plus a test in `tests/test_recon.py`. Until then that vendor's PDFs are flagged, never skipped. (Or send them the acknowledgment form.)

## 6. Deliberately not built

- ERP write-back.
- Automatic vendor emails (drafts only).
- LLM-first extraction.
- A master-data platform.
- A generic chatbot (the Q&A is scoped to the readout data and quotes its figures).
- A vendor portal.
- Cloud hosting.
- Causal claims from 8 months of data.

## 6b. Where AI is used (optional, tested live)

Every number in the deliverables is computed deterministically. **Claude is used in exactly two places.** Both are optional and run only when `ANTHROPIC_API_KEY` is set. Neither decides anything on its own. Evidence is in `output/ai_evidence/`.

| Where | What Claude does | Guardrails | Live result on the case data |
|---|---|---|---|
| **Task 1: AI reader** (`recon.py --ai`, or the Desk *Settings* switch) | Reads a PDF no template can: a new vendor layout, a scan when Tesseract is missing, another language | Strict JSON schema; the same validation and PO checks as the parsers; everything it reads goes to *Needs your OK* | Agreed with the parsers on **125 / 125** qty/price/date fields across all 35 PDFs; the full check without OCR matched the OCR run on **44 / 44** lines |
| **Task 2: vendor Q&A** (`task2_vendor_readout/ask.py "question"`, or *Vendor history → Ask* in the Desk) | Answers the plant manager in plain English | Sees only the cleaned, computed readout tables, never the raw ERP; must quote each figure and its table; no causal claims; no arithmetic of its own (totals are precomputed) | Every quoted figure was verified against an independent recomputation; it refused to claim downtime the data can't show |

```bash
export ANTHROPIC_API_KEY=...        # your own key; never stored in any file
python3 task2_vendor_readout/ask.py --erp /data/beacon_erp.db "Which open orders are at risk in the next two weeks?"
python3 task1_confirmation_recon/ai_reader.py --benchmark /data/confirmations --vendors /data/vendor_master.csv
```

- **Why AI isn't the primary reader:** aerospace paperwork can be export-controlled, so sending PDFs to a cloud API is IT's decision. The deterministic parsers also make every value traceable.
- **Why no vector database:** the computed tables fit in one cached prompt (about 32k tokens). The Q&A is retrieval-grounded; similarity search would only add a way to miss the right row.
- **Requirements:** the Claude SDK needs Python 3.10+. On 3.9 everything else runs.
- **Cost:** a question costs cents, because the data block is prompt-cached after the first question. The 35-PDF benchmark costs about $1–2.

## 7. Known limitations

- **Unknown layouts:** a new vendor layout is unreadable until a parser is added.
- **OCR:** can misread digits; every scan goes to review.
- **Stale export:** an out-of-date open-PO export causes false "dropped" or "not open" flags.
- **Received value:** priced at PO price, not invoiced cost; there is no AP data in the extract.
- **Promise metrics:** stop at 2026-04-03, where the ERP confirmation feed ends.
- **Quality ranking:** not possible, because QC hold dates are unreliable.
- **Production impact:** there is no production data, so the readout can rank exposure but can't prove cost.

## 8. Layout

```
run_all.sh / run_all.bat     one-command run for graders
task1_confirmation_recon/    parsers · recon · workflow · store · report · vendor_forms · desk_app · ai_reader · tests/ · launchers
task2_vendor_readout/        analysis (cleaning + metrics) · forward_risk (watch list + backtest) · build_readout (xlsx/html/pdf) · ask (optional Q&A) · tests/
output/                      generated from the case data
presentation/                deck (build_deck.js regenerates it) and screenshots
```
