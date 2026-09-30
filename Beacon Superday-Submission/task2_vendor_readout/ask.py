"""
Ask the vendor readout a question in plain English (optional, needs a Claude API key).

    python3 ask.py --erp /data/beacon_erp.db "Why is Continental late, and what should I ask them?"
    python3 ask.py --erp /data/beacon_erp.db          # interactive: ask several questions

How it stays trustworthy:
  * Claude never sees the raw ERP tables (they contain traps: MM/DD/YYYY dates, keyed-in
    reversals, a stale qty_received field). It sees only the CLEANED, COMPUTED tables the
    readout itself is built from - vendor scorecard, monthly value, past-due lines, price
    exceptions, open orders at risk, data-quality notes and metric definitions.
  * It must quote the figures it uses and name the table they come from, say "not in the
    data" rather than guess, and never claim causes the data can't show.
  * It calculates nothing new that matters: every number it can cite already exists in the
    readout workbook, so an answer can be checked against Vendor_Readout.xlsx.

Credentials: the standard ANTHROPIC_API_KEY environment variable (never stored in files).
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
MODEL = "claude-opus-5-5"

SYSTEM = """You are the analyst behind Beacon Fasteners' vendor performance readout. Beacon makes aerospace fasteners; \
the plant manager and the senior buyer ask you questions about vendors.

Answer ONLY from the data below: tables computed from 8 months of ERP history after cleaning (the raw ERP had errors \
that are listed under data_quality). Rules:
- Quote the specific figures you rely on and name the table each comes from, e.g. (scorecard: on_time 42%).
- If the data does not contain the answer, say so plainly and name the data that would be needed. Never estimate or invent numbers.
- Do not claim causes the data cannot show (there is no production, quality-cost or invoice data). Say "the data shows X; it does not show why".
- Distinguish received value, ordered value and open value; say which you mean.
- Be concise and practical: lead with the answer in one or two sentences, then the evidence, then (if useful) a suggested next step.
- Percentages in the data are fractions (0.42 = 42%). Money is USD at PO price.

DATA (JSON):
"""

def _records(df, cols=None, n=None):
    df = df if cols is None else df[[c for c in cols if c in df.columns]]
    if n:
        df = df.head(n)
    return json.loads(df.to_json(orient="records", date_format="iso", double_precision=4))

def build_context(erp, task1_lines=None):
    """The computed tables, as one JSON string. Same numbers the readout workbook shows."""
    sys.path.insert(0, HERE)
    from analysis import METRIC_DEFINITIONS, build
    from forward_risk import backtest, forward_book
    a = build(erp)
    s = a["score"].reset_index()
    monthly = a["monthly"].rename(columns=a["vend"].to_dict())
    fw = forward_book(a, task1_lines)
    ctx = dict(
        as_of=str(a["as_of"].date()), history_from=str(a["first_day"].date()), history_to=str(a["last_day"].date()),
        vendors={k: v for k, v in a["vend"].items()},
        scorecard=_records(s),
        received_value_by_month_usd=json.loads(monthly.round(2).to_json(orient="index")),
        on_time_by_required_month=json.loads(a["monthly_ot"].rename(columns=a["vend"].to_dict()).round(3).to_json(orient="index")),
        continental_by_service=_records(a["v4_parts"].reset_index()),
        past_due_open_lines=_records(a["past_due"].sort_values("balance_value", ascending=False),
                                     ["vendor_id", "po_number", "line_no", "part_id", "qty_ordered", "rec_qty", "balance_qty",
                                      "balance_value", "required", "days_past_due", "kind"]),
        confirmed_above_po_price=_records(a["above"], ["vendor_id", "po_number", "line_no", "part_id", "unit_price",
                                                      "confirmed_price", "price_var", "exposure_ordered", "conf_month"]),
        revised_confirmations=_records(a["superseded"], ["vendor_id", "po_number", "line_no", "confirmed_qty", "new_qty",
                                                          "promised_date", "new_prom", "kind"]),
        open_orders_next_60_days=_records(fw, ["tier", "vendor", "po_number", "line_no", "part_id", "open_qty", "open_value",
                                               "required", "days_to_due", "promised", "promise_source", "hist_on_time",
                                               "hist_n", "why"]),
        watch_list_backtest=backtest(a["rec"]),
        data_quality=[dict(issue=k, value=str(v), handling=h) for k, v, h in a["dq"]],
        metric_definitions=dict(METRIC_DEFINITIONS),
    )
    return json.dumps(ctx, default=str)

class Advisor:
    """Holds the data context and the conversation, so follow-up questions work."""

    def __init__(self, erp, task1_lines=None, client=None):
        import anthropic
        self.client = client or anthropic.Anthropic()
        self.context = build_context(erp, task1_lines)
        self.history = []

    def ask(self, question):
        """Returns the answer text. Raises anthropic errors to the caller (CLI / app show them)."""
        messages = self.history + [{"role": "user", "content": question}]
        response = self.client.beta.messages.create(
            model=MODEL,
            max_tokens=16000,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            output_config={"effort": "medium"},
            # The data block is identical on every question, so it is cached after the first call.
            system=[{"type": "text", "text": SYSTEM + self.context, "cache_control": {"type": "ephemeral"}}],
            messages=messages,
        )
        if response.stop_reason == "refusal":
            return "The model declined to answer this question."
        text = "".join(b.text for b in response.content if b.type == "text").strip()
        if response.stop_reason == "max_tokens":
            text += "\n\n[answer cut off - ask a narrower question]"
        self.history = messages + [{"role": "assistant", "content": response.content}]
        self.last_usage = response.usage
        return text

def main():
    ap = argparse.ArgumentParser(description="Ask the vendor readout a question (needs ANTHROPIC_API_KEY).")
    ap.add_argument("question", nargs="*")
    ap.add_argument("--erp", required=True)
    ap.add_argument("--task1-lines", default=os.path.join(HERE, "..", "output", "confirmation_lines_extracted.csv"))
    a = ap.parse_args()
    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("Set ANTHROPIC_API_KEY to use this. The readout itself (build_readout.py) needs no key.")
    import anthropic
    adv = Advisor(a.erp, a.task1_lines)
    questions = [" ".join(a.question)] if a.question else None
    try:
        if questions:
            print(adv.ask(questions[0]))
            return
        print("Ask about Beacon's vendors (empty line to quit).")
        while True:
            q = input("\n> ").strip()
            if not q:
                break
            print("\n" + adv.ask(q))
    except anthropic.AuthenticationError:
        sys.exit("The API key was rejected.")
    except anthropic.APIStatusError as e:
        sys.exit("Claude API error %s: %s" % (e.status_code, e.message))
    except anthropic.APIConnectionError:
        sys.exit("Could not reach the Claude API (network).")

if __name__ == "__main__":
    main()
