# Live AI test evidence (Claude `claude-opus-5-5`, run on the case data)

AI is used in exactly two places, both optional (they run only when `ANTHROPIC_API_KEY` is set). Every number in the main deliverables is still computed deterministically.

## 1. AI reader: Task 1, PDFs no template can read

`ai_reader_benchmark_35_pdfs.txt`

Claude read all 35 PDFs, and its output was compared field by field with the template parsers.

| Result | Value |
|---|---|
| Documents read | 35 / 35 |
| PO numbers correct | 35 / 35 |
| Quantity / price / date fields agreeing with the parsers | 125 / 125 |

This includes the 4 scanned Continental PDFs and the German Ostmark documents.

**End to end:** the daily check ran with `--ai` on a machine *without* Tesseract, so Claude read the 4 scans instead of OCR. The workbook matched the OCR run on **44 / 44 lines**.

That test found two integration bugs, both fixed and covered by tests:
- a plain printed date was flagged as "ambiguous";
- documents with no text layer were treated as duplicates of each other.

## 2. Vendor Q&A: Task 2 (`task2_vendor_readout/ask.py`, and the "Ask about your vendors" box in the Desk app)

**Grounding:** Claude answers only from the cleaned, computed readout tables (about 32k tokens, prompt-cached). It never sees the raw ERP.

**Run 1** (`vendor_qa_run1_before_fix.txt`)
- Every figure quoted from the data was exact.
- One total it added up itself was wrong: it said "about $1.44M" where the correct figure is $1,708,250.
- It correctly refused to claim production downtime, which the data cannot show, and named the data that would answer the question.

**Fix:** open-order totals by vendor, risk tier and due window are now precomputed, and the model is instructed never to add numbers up itself.

**Run 2** (`vendor_qa_run2_after_fix.txt`)
- Every figure was checked against an independent recomputation, and all were correct.
- Where a total was not precomputed, it said so and pointed to the workbook instead of guessing.

**Why no vector database:** the computed tables fit easily in one cached prompt. Retrieval by similarity would add a way to miss the right row without adding accuracy.
