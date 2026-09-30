"""
Optional AI fallback for confirmation layouts the tool has no parser for.

OFF by default. Beacon makes aerospace parts; POs and vendor paperwork can be
export-controlled (ITAR/EAR) or customer-confidential. Sending a PDF to a cloud
API is an IT / compliance decision, so this only runs when an administrator sets

    BEACON_AI_FALLBACK=1        (and Anthropic credentials, e.g. ANTHROPIC_API_KEY)

What it does: sends ONE unreadable PDF to Claude and asks for the same fields
our parsers extract, as schema-validated JSON. What it never does: decide
anything. Every line it reads is marked "AI-read" and goes to Review Required,
so Lisa confirms it against the PDF before it counts.

The regular path (template parsers + OCR) stays fully offline.
"""
import base64
import json
import os
import re
from datetime import datetime

MODEL = "claude-opus-5-5"

SCHEMA = {
    "type": "object",
    "properties": {
        "vendor_name": {"type": "string"},
        "po_number": {"type": "string", "description": "Beacon PO number as printed, e.g. PO-4500050001; empty if none"},
        "doc_type": {"type": "string", "enum": ["ACK", "REVISED_ACK", "INVOICE", "ACK_NO_DETAIL", "OTHER"]},
        "doc_date": {"type": "string", "description": "YYYY-MM-DD, empty if none"},
        "vendor_ref": {"type": "string"},
        "lines": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "line_no": {"type": "integer", "description": "0 if not printed"},
                "vendor_pn": {"type": "string", "description": "part number exactly as printed; empty if none"},
                "description": {"type": "string"},
                "qty": {"type": "number"},
                "unit_price": {"type": "number", "description": "-1 if no price is printed"},
                "currency": {"type": "string", "description": "ISO code, e.g. USD, EUR"},
                "promise_date": {"type": "string", "description": "YYYY-MM-DD; if only a week or range is given, the LAST day; empty if none"},
                "promise_text": {"type": "string", "description": "the promise/ship date exactly as printed"},
            },
            "required": ["line_no", "vendor_pn", "description", "qty", "unit_price", "currency", "promise_date", "promise_text"],
            "additionalProperties": False}},
        "notes": {"type": "string", "description": "anything a buyer should know: partial, revision, backorder, substitution"},
    },
    "required": ["vendor_name", "po_number", "doc_type", "doc_date", "vendor_ref", "lines", "notes"],
    "additionalProperties": False,
}

PROMPT = """This PDF is a supplier's response to a purchase order from Beacon Fasteners (an aerospace fastener maker).
Extract exactly what is printed. Do not infer values that are not on the page, and do not translate part numbers.
Dates: convert to YYYY-MM-DD (European documents use DD.MM.YYYY; 'KW' means ISO calendar week).
If the document is an invoice rather than an acknowledgment, set doc_type INVOICE. If it acknowledges the order
without quantities or dates, use ACK_NO_DETAIL and leave lines empty."""


def enabled():
    if os.environ.get("BEACON_AI_FALLBACK") != "1":
        return False
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return False
    return True


def _date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date() if s else None
    except ValueError:
        return None


def to_doc(data, vendors):
    """Map the model's JSON onto the parser's normalized doc shape."""
    vid = None
    name = (data.get("vendor_name") or "").lower()
    for k, v in (vendors or {}).items():
        first = v["vendor_name"].lower().split()[0]
        if first and first in name:
            vid = k
            break
    m = re.search(r"PO[-\s]?(\d{10})", data.get("po_number") or "")
    lines = []
    for l in data.get("lines", []):
        lines.append(dict(line_no=l["line_no"] or None, vendor_pn=l["vendor_pn"] or None, description=l["description"] or None,
                          qty=float(l["qty"]), uom=None, unit_price=None if l["unit_price"] < 0 else float(l["unit_price"]),
                          currency=(l["currency"] or "USD").upper(), promise_date=_date(l["promise_date"]),
                          promise_note=("AI read the date '%s' as %s" % (l["promise_text"], l["promise_date"]))
                          if l["promise_text"] and l["promise_text"] != l["promise_date"] else None))
    warnings = ["Read by AI (no template for this layout) - check every value against the PDF."]
    if data.get("notes"):
        warnings.append("AI note: " + data["notes"])
    if not vid:
        warnings.append("Vendor '%s' not matched to vendor master." % data.get("vendor_name"))
    return dict(vendor_id=vid, template="AI fallback (%s)" % MODEL, doc_type=data["doc_type"] if data["doc_type"] != "OTHER" else "UNREADABLE",
                po_number="PO-" + m.group(1) if m else None, doc_date=_date(data.get("doc_date")),
                vendor_ref=data.get("vendor_ref") or None, part_of=None, supersedes_prior=data["doc_type"] == "REVISED_ACK",
                lines=lines, warnings=warnings, text_source="ai")


def extract(path, vendors=None, client=None):
    """Returns a normalized doc, or None if the call fails (the doc then stays 'read by hand')."""
    import anthropic
    client = client or anthropic.Anthropic()
    with open(path, "rb") as f:
        pdf = base64.standard_b64encode(f.read()).decode("utf-8")
    try:
        response = client.beta.messages.create(
            model=MODEL,
            max_tokens=16000,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            output_config={"effort": "medium", "format": {"type": "json_schema", "schema": SCHEMA}},
            messages=[{"role": "user", "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf}},
                {"type": "text", "text": PROMPT}]}],
        )
    except anthropic.RateLimitError:
        return None
    except anthropic.APIStatusError:
        return None
    except anthropic.APIConnectionError:
        return None
    if response.stop_reason in ("refusal", "max_tokens"):
        return None
    text = next((b.text for b in response.content if b.type == "text"), None)
    if not text:
        return None
    return to_doc(json.loads(text), vendors)
