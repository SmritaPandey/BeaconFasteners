"""
Turn a vendor order-acknowledgment PDF into a normalized record.

One parser per vendor template (6 today). Deterministic regex, not an LLM:
the templates are stable, the output has to be auditable, and a wrong guess
on a PO line costs Lisa more than a "couldn't read this, please look" flag.
Anything that doesn't match a known template is returned as UNREADABLE so it
lands on the exceptions sheet instead of being silently dropped.
"""
import io
import re
from datetime import date, datetime

import fitz  # PyMuPDF


# --------------------------------------------------------------------------
# Text extraction
# --------------------------------------------------------------------------

def extract_text(path):
    """Return (text, source) where source is 'text' or 'ocr'."""
    doc = fitz.open(path)
    text = "\n".join(p.get_text("text", sort=True) for p in doc)
    if len(text.strip()) > 50:
        return text, "text"
    # No text layer -> scanned PDF -> OCR
    import pytesseract
    from PIL import Image
    pages = []
    for p in doc:
        pix = p.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        pages.append(pytesseract.image_to_string(img, config="--psm 6"))
    return "\n".join(pages), "ocr"


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

PO_RE = re.compile(r"PO[-\s]?(\d{10})")


def _po(text):
    m = PO_RE.search(text or "")
    return "PO-" + m.group(1) if m else None


def _num(s):
    """Parse '1,500' / '1.140,80' / '1140.8000' / '€16.7440' -> float."""
    s = s.strip().replace("$", "").replace("€", "").replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):   # German: 1.140,80
            s = s.replace(".", "").replace(",", ".")
        else:                              # US: 1,140.80
            s = s.replace(",", "")
    elif "," in s:
        # '1,500' (thousands) vs '16,74' (German decimal)
        if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    return float(s)


def _us_date(s):
    return datetime.strptime(s.strip(), "%m/%d/%Y").date()


def _de_date(s):
    return datetime.strptime(s.strip(), "%d.%m.%Y").date()


def _parse_promise_de(s):
    """Ostmark: '08.06.2026' or 'KW 20-22 / 2026' or 'KW 21 / 2026'.
    Returns (date, note). For a week range we take the END of the last
    week - the latest date the vendor has actually committed to."""
    s = s.strip()
    m = re.search(r"\d{2}\.\d{2}\.\d{4}", s)
    if m:
        return _de_date(m.group(0)), None
    m = re.search(r"KW\s*(\d{1,2})(?:\s*-\s*(\d{1,2}))?\s*/\s*(\d{4})", s)
    if m:
        w1, w2, yr = int(m.group(1)), int(m.group(2) or m.group(1)), int(m.group(3))
        start = date.fromisocalendar(yr, w1, 1)
        end = date.fromisocalendar(yr, w2, 7)
        return end, "Vague promise '%s' (calendar weeks) = %s to %s; using %s" % (
            s, start.strftime("%m/%d"), end.strftime("%m/%d"), end.strftime("%m/%d"))
    return None, "Could not read promise date '%s'" % s


def _doc(**kw):
    base = dict(vendor_id=None, template=None, doc_type="ACK", po_number=None,
                doc_date=None, vendor_ref=None, part_of=None,
                supersedes_prior=False, lines=[], warnings=[])
    base.update(kw)
    return base


def _line(**kw):
    base = dict(line_no=None, vendor_pn=None, description=None, qty=None,
                uom=None, unit_price=None, currency="USD", promise_date=None,
                promise_note=None)
    base.update(kw)
    return base


# --------------------------------------------------------------------------
# Vendor templates
# --------------------------------------------------------------------------

# Apex + Heritage share a table layout:  1  PN  description  qty  $price  mm/dd/yyyy
TABLE_LINE = re.compile(
    r"^\s*(\d{1,3})\s+(\S+)\s+(.+?)\s{2,}([\d,]+)\s+\$([\d.,]+)\s+(\d{2}/\d{2}/\d{4})\s*$")


def _table_lines(text):
    out = []
    for ln in text.splitlines():
        m = TABLE_LINE.match(ln)
        if m:
            out.append(_line(line_no=int(m.group(1)), vendor_pn=m.group(2),
                             description=m.group(3).strip(), qty=_num(m.group(4)),
                             unit_price=_num(m.group(5)),
                             promise_date=_us_date(m.group(6))))
    return out


def parse_apex(text):
    d = _doc(vendor_id="V001", template="Apex ERP acknowledgment")
    d["po_number"] = _po(re.search(r"Customer PO:\s*(\S+)", text).group(1))
    d["doc_date"] = _us_date(re.search(r"Acknowledgment Date:\s*(\S+)", text).group(1))
    m = re.search(r"Apex Order #:\s*(\S+)", text)
    d["vendor_ref"] = m.group(1) if m else None
    d["lines"] = _table_lines(text)
    return d


def parse_heritage(text):
    d = _doc(vendor_id="V002", template="Heritage sales order confirmation")
    d["po_number"] = _po(re.search(r"Your reference:\s*(\S+)", text).group(1))
    d["doc_date"] = _us_date(re.search(r"Issued:\s*(\S+)", text).group(1))
    m = re.search(r"SO #\s*(\S+)", text)
    d["vendor_ref"] = m.group(1) if m else None
    d["lines"] = _table_lines(text)
    return d


LIB_LINE = re.compile(r"^\s*(.+?)\s{2,}([\d,]+)\s+([A-Z]{2,6})\s+(\d{2}/\d{2}/\d{4})\s*$")


def parse_liberty(text):
    d = _doc(vendor_id="V003", template="Liberty process order ack")
    m = re.search(r"PO Reference:\s*(PO-\d+)(?:\s*\(part\s*(\d+)\s*of\s*(\d+)\))?", text)
    d["po_number"] = m.group(1)
    if m.group(2):
        d["part_of"] = (int(m.group(2)), int(m.group(3)))
    d["doc_date"] = _us_date(re.search(r"\bDate:\s*(\S+)", text).group(1))
    m = re.search(r"Process Order #:\s*(\S+)", text)
    d["vendor_ref"] = m.group(1) if m else None
    for ln in text.splitlines():
        m = LIB_LINE.match(ln)
        if m and "Description" not in ln:
            d["lines"].append(_line(description=m.group(1).strip(), qty=_num(m.group(2)),
                                    vendor_pn=m.group(3), unit_price=None,
                                    promise_date=_us_date(m.group(4))))
    if not d["lines"]:
        d["doc_type"] = "ACK_NO_DETAIL"
        d["warnings"].append("Vendor acknowledged the PO but gave no quantity or date "
                             "('schedule to be confirmed').")
    return d


def parse_continental(text):
    """Scanned + OCR'd. Tolerate common OCR slips (Ib->lb, stray punctuation)."""
    d = _doc(vendor_id="V004", template="Continental scanned confirmation (OCR)")
    t = text.replace("Ibs", "lbs").replace(" Ib", " lb")
    d["po_number"] = _po(re.search(r"Customer PO[:.]?\s*(\S+)", t).group(1))
    d["doc_date"] = _us_date(re.search(r"\bDate[:.]?\s*(\d{2}/\d{2}/\d{4})", t).group(1))
    m = re.search(r"WO[:.]*\s*(CQH-\d+)", t)
    d["vendor_ref"] = m.group(1) if m else None
    desc = re.search(r"Description[:.]?\s*(.+)", t)
    qty = re.search(r"Quantity[:.]?\s*([\d,]+)\s*(\w+)?", t)
    prom = re.search(r"Promise[:.]?\s*(\d{2}/\d{2}/\d{4})", t)
    if desc and qty:
        d["lines"].append(_line(description=desc.group(1).strip(), qty=_num(qty.group(1)),
                                uom=(qty.group(2) or "").lower() or None,
                                promise_date=_us_date(prom.group(1)) if prom else None))
    else:
        d["warnings"].append("OCR could not find description/quantity - check scan by hand.")
    d["warnings"].append("Scanned document read by OCR - spot-check numbers.")
    return d


OST_LINE = re.compile(
    r"^\s*(\d{1,3})\s+(\S+)\s+(.+?)\s{2,}([\d.,]+)\s+€\s*([\d.,]+)\s+(.+?)\s*$")


def parse_ostmark(text):
    d = _doc(vendor_id="V005", template="Ostmark Auftragsbestaetigung (DE/EN)")
    d["po_number"] = _po(re.search(r"Your PO:\s*(\S+)", text).group(1))
    d["doc_date"] = _de_date(re.search(r"Date:\s*(\d{2}\.\d{2}\.\d{4})", text).group(1))
    m = re.search(r"Auftrag-Nr\.:\s*(\S+)", text)
    d["vendor_ref"] = m.group(1) if m else None
    for ln in text.splitlines():
        m = OST_LINE.match(ln)
        if m:
            pdate, note = _parse_promise_de(m.group(6))
            d["lines"].append(_line(line_no=int(m.group(1)), vendor_pn=m.group(2),
                                    description=m.group(3).strip(), qty=_num(m.group(4)),
                                    unit_price=_num(m.group(5)), currency="EUR",
                                    promise_date=pdate, promise_note=note))
    return d


QS_LINE = re.compile(r"(\S+)\s+qty\s+([\d,]+)\s+@\s+\$([\d.,]+)\s+ea\s+ship\s+(\d{2}/\d{2}/\d{4})", re.I)


def parse_quickship(text):
    d = _doc(vendor_id="V006", template="QuickShip email (free text)")
    subj = re.search(r"Subject:\s*(.+)", text)
    subj = subj.group(1) if subj else ""
    d["po_number"] = _po(subj) or _po(text)
    d["doc_date"] = _us_date(re.search(r"^Date:\s*(\S+)", text, re.M).group(1))
    low = text.lower()
    if "invoice" in subj.lower():
        d["doc_type"] = "INVOICE"
        d["warnings"].append("Vendor sent an INVOICE, not an order acknowledgment"
                             + (" - says goods already shipped" if "shipped" in low else "")
                             + ". Check with receiving; AP will see this.")
    if "revised" in subj.lower() or "disregard" in low:
        d["doc_type"] = "REVISED_ACK"
        d["supersedes_prior"] = True
    for m in QS_LINE.finditer(text):
        d["lines"].append(_line(vendor_pn=m.group(1), qty=_num(m.group(2)),
                                unit_price=_num(m.group(3)),
                                promise_date=_us_date(m.group(4))))
    return d


TEMPLATES = [
    # (signature regex, parser)
    (re.compile(r"APEX BAR", re.I), parse_apex),
    (re.compile(r"Heritage.*Cold|HCH-\d", re.I | re.S), parse_heritage),
    (re.compile(r"Liberty\s+Surface", re.I), parse_liberty),
    (re.compile(r"CONTINENTAL QUALITY", re.I), parse_continental),
    (re.compile(r"Ostmark", re.I), parse_ostmark),
    (re.compile(r"quickship", re.I), parse_quickship),
]


def parse_file(path):
    if path.lower().endswith(".xlsx"):
        import vendor_forms
        try:
            d = vendor_forms.read_ack_form(path)
        except Exception as e:
            d, err = None, e
        else:
            err = None
        if d is None:
            d = _doc(template="unknown", doc_type="UNREADABLE", text_source="form", raw_text="",
                     warnings=["Excel file is not a Beacon acknowledgment form%s - read by hand." % (" (%s)" % err if err else "")])
        return d
    text, source = extract_text(path)
    for sig, fn in TEMPLATES:
        if sig.search(text):
            try:
                d = fn(text)
            except Exception as e:  # vendor recognised but layout changed
                d = _doc(template=fn.__name__, doc_type="UNREADABLE",
                         warnings=["Recognised vendor but could not parse (%s). "
                                   "Layout may have changed - read by hand." % e])
                d["po_number"] = _po(text)
            break
    else:
        d = _doc(template="unknown", doc_type="UNREADABLE",
                 warnings=["Unrecognised document layout - read by hand."])
        d["po_number"] = _po(text)
    d["text_source"] = source
    d["raw_text"] = text
    if d["doc_type"] not in ("UNREADABLE", "ACK_NO_DETAIL") and not d["lines"]:
        d["warnings"].append("No line items found on document - read by hand.")
    return d
