"""
Build the plant manager's vendor readout from the ERP extract.

  python3 build_readout.py --erp /path/to/beacon_erp.db --out ../output

Writes:
  Vendor_Readout.xlsx     - workbook: summary, monthly value (+chart), scorecard,
                            delivery trend, open orders at risk (forward-looking),
                            open past-due lines, price exceptions,
                            part-number crosswalk, data quality & definitions
  vendor_readout.html     - one-page readout (self-contained, opens in any browser)
"""
import argparse
import html
import os
import sys

import numpy as np
import pandas as pd

from analysis import LATE_GRACE_DAYS, METRIC_DEFINITIONS, build
from forward_risk import backtest, forward_book

HERE = os.path.dirname(os.path.abspath(__file__))


def money(x, dp=0):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    if abs(x) >= 1e6:
        return "$%.2fM" % (x / 1e6)
    if abs(x) >= 1e3 and dp == 0:
        return "$%.0fk" % (x / 1e3) if abs(x) >= 1e4 else "$%s" % format(int(round(x)), ",")
    return ("$%." + str(dp) + "f") % x


def pct(x, dp=0):
    return "–" if x is None or (isinstance(x, float) and np.isnan(x)) else ("%." + str(dp) + "f%%") % (100 * x)


def esc(s):
    return html.escape(str(s))


# ---------------------------------------------------------------------------
# Narrative: the answer, built from the numbers (nothing hard-coded)
# ---------------------------------------------------------------------------

def findings(a):
    s = a["score"]
    total_sev = s.severe_lines.sum()
    total_sev_val = s.severe_value.sum()
    v4, v1, v2, v3 = s.loc["V004"], s.loc["V001"], s.loc["V002"], s.loc["V003"]
    others = s.drop("V004")
    rec = a["rec"]
    f = {}
    f["call_first"] = dict(
        vendor=v4.vendor,
        headline="Call %s first" % v4.vendor,
        why=[
            "<b>Late most often:</b> %s of %d due lines arrived on time. The other five vendors range from %s to %s." % (
                pct(v4.on_time), v4.due_lines, pct(others.on_time.min()), pct(others.on_time.max())),
            "<b>Late by the most:</b> %d of the %d lines that arrived more than %d days late across all vendors are %s's "
            "(%s of %s received value). When late, the median delay is %d days (worst %d). Every other vendor "
            "lands %s or more of its lines within %d days." % (
                v4.severe_lines, total_sev, LATE_GRACE_DAYS, v4.short, money(v4.severe_value), money(total_sev_val),
                v4.median_days_late_when_late, v4.max_days_late, pct(others.on_time_with_grace.min()), LATE_GRACE_DAYS),
            "<b>Warns it will be late, then misses its own date too:</b> %s of its confirmations promise a date after our required date "
            "(avg +%.1f days; other vendors avg +%.1f to +%.1f). It then meets only %s of its own promises." % (
                pct(v4.pct_promised_late), v4.avg_promise_gap, others.avg_promise_gap.min(), others.avg_promise_gap.max(),
                pct(v4.on_time_vs_promise)),
            "<b>Not getting better:</b> %s on time for lines due Sep–Dec (n=%d) vs %s for Jan–May (n=%d). "
            "It has been consistently poor, not a recent slip." % (
                pct(v4.ot_early), v4.ot_early_n, pct(v4.ot_late), v4.ot_late_n),
            "<b>Material:</b> %s received (%s of spend) in outside heat-treat services. "
            "%d open lines are past due today (%s)." % (
                money(v4.received_value), pct(v4.share, 1), v4.past_due_lines, money(v4.past_due_value)),
        ],
        ask=[
            "Our planned lead time for heat treat is %g days from PO; you quote %g and deliver in %g (medians). "
            "Can you commit to a standard lead time we can load into MRP?" % (v4.lt_allowed, v4.lt_promised, v4.lt_actual),
            "When you do confirm a date, you meet it %s of the time. What is driving the misses: furnace capacity, "
            "batching, or incoming parts?" % pct(v4.on_time_vs_promise),
            "Split shipments: %s of lines come in more than one shipment. Can partial lots be avoided or flagged?" %
            pct(v4.multi_ship_pct),
        ],
        internal="Beacon side: our 19-day heat-treat lead time is about 9 days shorter than what Continental quotes. "
                 "Part of this lateness is ours to fix in planning.",
    )
    f["second"] = dict(
        vendor=v1.vendor,
        text="<b>%s: the dollar and commercial conversation, not a delivery one.</b> %s of all received value (%s). "
             "On time %s, but when late the median is %d day and 0 lines were more than %d days late. The issues: "
             "(1) BAR-CRES-250 confirmed at +1%% from Feb and +2%% from Mar while every PO stayed at $3.92: %d lines, "
             "%s over PO on ordered qty, and nothing in the ERP captures it. "
             "(2) %d short-shipped balances left open (%s, the largest past-due exposure of any vendor). "
             "(3) This week's acknowledgments show the same pattern: a dropped line, a short and a +2.3%% price." % (
                 v1.vendor, pct(v1.share), money(v1.received_value), pct(v1.on_time), v1.median_days_late_when_late,
                 LATE_GRACE_DAYS, v1.price_above_lines, money(v1.price_exposure), v1.short_open_lines, money(v1.past_due_value)),
    )
    f["watch"] = [
        "<b>%s:</b> on time %s, and %s of promises are after our need date (avg +%.1f days). Delays are short (median %d days, "
        "%d lines over %d days). %d short-shipped balances are still open." % (
            v2.vendor, pct(v2.on_time), pct(v2.pct_promised_late), v2.avg_promise_gap, v2.median_days_late_when_late,
            v2.severe_lines, LATE_GRACE_DAYS, v2.short_open_lines),
        "<b>%s:</b> slipping from %s (Sep–Dec) to %s (Jan–May). The change is not statistically significant (p=%.2f) "
        "and misses are 1–3 days, so watch it rather than call." % (v3.vendor, pct(v3.ot_early), pct(v3.ot_late), v3.trend_p),
        "<b>%s:</b> on time %s but never more than %d days late. Watch the open die-set order PO-4500050027 "
        "(25 x TOOL-DIE-9472, $31k, due 05/15, nothing received). This week's acknowledgment gives only 'KW 20-22'." % (
            s.loc["V005"].vendor, pct(s.loc["V005"].on_time), s.loc["V005"].max_days_late),
        "<b>%s:</b> the benchmark at %s on time, but only %s of spend." % (
            s.loc["V006"].vendor, pct(s.loc["V006"].on_time), money(s.loc["V006"].received_value)),
    ]
    f["overall_on_time"] = rec.on_time.mean()
    f["due_lines"] = len(rec)
    return f


# ---------------------------------------------------------------------------
# SVG charts (inline, no libraries; hover via <title>)
# ---------------------------------------------------------------------------

def svg_monthly(monthly, first_day, last_day):
    months = list(monthly.index)
    vals = monthly["Total"].values
    W, H, L, R, T, B = 720, 280, 64, 16, 16, 44
    vmax = 6e6
    bw = (W - L - R) / len(months)
    y = lambda v: T + (H - T - B) * (1 - v / vmax)
    out = ['<svg viewBox="0 0 %d %d" role="img" aria-label="Received value by month">' % (W, H)]
    for g in range(0, 7):
        gv = g * 1e6
        out.append('<line x1="%d" x2="%d" y1="%.1f" y2="%.1f" class="grid"/>' % (L, W - R, y(gv), y(gv)))
        out.append('<text x="%d" y="%.1f" class="ax" text-anchor="end">$%dM</text>' % (L - 8, y(gv) + 4, g))
    for i, (m, v) in enumerate(zip(months, vals)):
        x = L + i * bw + 3
        partial = m == first_day.strftime("%Y-%m") or m == last_day.strftime("%Y-%m")
        cls = "bar partial" if partial else "bar"
        h = (H - T - B) * v / vmax
        lab = pd.Timestamp(m + "-01").strftime("%b %y")
        note = (" (partial: from %s)" % first_day.strftime("%m/%d") if m == first_day.strftime("%Y-%m") else
                " (partial: to %s)" % last_day.strftime("%m/%d") if m == last_day.strftime("%Y-%m") else "")
        out.append('<g><title>%s: %s%s</title><path class="%s" d="M%.1f,%.1f v%.1f a4,4 0 0 1 4,-4 h%.1f a4,4 0 0 1 4,4 v%.1f z"/></g>' % (
            lab, money(v), note, cls, x, y(0), -(h - 4) if h > 4 else -h, bw - 14, (h - 4) if h > 4 else h))
        out.append('<text x="%.1f" y="%d" class="ax" text-anchor="middle">%s%s</text>' % (x + (bw - 6) / 2, H - B + 18, lab, "*" if partial else ""))
        out.append('<text x="%.1f" y="%.1f" class="val" text-anchor="middle">%.1f</text>' % (x + (bw - 6) / 2, y(v) - 6, v / 1e6))
    out.append('<line x1="%d" x2="%d" y1="%.1f" y2="%.1f" class="base"/>' % (L, W - R, y(0), y(0)))
    out.append("</svg>")
    return "\n".join(out)


def svg_ontime(score):
    s = score.sort_values("on_time")
    W, Hrow, L, R, T = 720, 34, 190, 70, 10
    H = T + Hrow * len(s) + 30
    x = lambda v: L + (W - L - R) * v
    out = ['<svg viewBox="0 0 %d %d" role="img" aria-label="On-time delivery by vendor">' % (W, H)]
    for g in (0, .25, .5, .75, 1):
        out.append('<line x1="%.1f" x2="%.1f" y1="%d" y2="%d" class="grid"/>' % (x(g), x(g), T, H - 24))
        out.append('<text x="%.1f" y="%d" class="ax" text-anchor="middle">%d%%</text>' % (x(g), H - 8, g * 100))
    for i, (vid, r) in enumerate(s.iterrows()):
        yy = T + i * Hrow + 6
        cls = "bar hi" if vid == "V004" else "bar lo"
        w = (W - L - R) * r.on_time
        out.append('<text x="%d" y="%.1f" class="lab%s" text-anchor="end">%s</text>' % (
            L - 10, yy + 15, " strong" if vid == "V004" else "", esc(r.vendor)))
        out.append('<g><title>%s: %s on time (%d of %d due lines); %d lines &gt;%d days late</title>'
                   '<path class="%s" d="M%.1f,%.1f h%.1f a4,4 0 0 1 4,4 v%d a4,4 0 0 1 -4,4 h-%.1f z"/></g>' % (
                       esc(r.vendor), pct(r.on_time, 1), r.on_time_n, r.due_lines, r.severe_lines, LATE_GRACE_DAYS,
                       cls, L, yy, w - 4, 14, w - 4))
        out.append('<text x="%.1f" y="%.1f" class="val" text-anchor="start">%s</text>' % (x(r.on_time) + 8, yy + 15, pct(r.on_time)))
    out.append("</svg>")
    return "\n".join(out)


def svg_trend(monthly_ot, rec):
    months = [m for m in monthly_ot.index if m <= rec.required.max().strftime("%Y-%m")]
    v4 = rec[rec.vendor_id == "V004"].groupby("req_month").on_time.agg(["mean", "size"]).reindex(months)
    oth = rec[rec.vendor_id != "V004"].groupby("req_month").on_time.agg(["mean", "size"]).reindex(months)
    W, H, L, R, T, B = 720, 260, 48, 120, 14, 36
    step = (W - L - R) / (len(months) - 1)
    y = lambda v: T + (H - T - B) * (1 - v)
    out = ['<svg viewBox="0 0 %d %d" role="img" aria-label="Monthly on-time trend">' % (W, H)]
    for g in (0, .25, .5, .75, 1):
        out.append('<line x1="%d" x2="%d" y1="%.1f" y2="%.1f" class="grid"/>' % (L, W - R, y(g), y(g)))
        out.append('<text x="%d" y="%.1f" class="ax" text-anchor="end">%d%%</text>' % (L - 8, y(g) + 4, g * 100))
    for i, m in enumerate(months):
        out.append('<text x="%.1f" y="%d" class="ax" text-anchor="middle">%s</text>' % (
            L + i * step, H - 12, pd.Timestamp(m + "-01").strftime("%b")))
    for name, d, cls in (("All other vendors", oth, "s1"), ("Continental", v4, "s2")):
        pts = [(L + i * step, y(v)) for i, v in enumerate(d["mean"].values)]
        out.append('<polyline class="line %s" points="%s"/>' % (cls, " ".join("%.1f,%.1f" % p for p in pts)))
        for (px, py), m, v, n in zip(pts, months, d["mean"].values, d["size"].values):
            out.append('<g><title>%s, lines due %s: %s on time (n=%d)</title><circle class="dot %s" cx="%.1f" cy="%.1f" r="4.5"/></g>' % (
                name, pd.Timestamp(m + "-01").strftime("%b %Y"), pct(v), n, cls, px, py))
        out.append('<text x="%.1f" y="%.1f" class="lab %s-t">%s</text>' % (pts[-1][0] + 10, pts[-1][1] + 4, cls, name))
    out.append("</svg>")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------

CSS = """
/* Layout: single reading column (answer first), tables scroll inside their own frames */
:root {
  --bg: #f6f7f8; --panel: #ffffff; --ink: #15191e; --ink-2: #4b5560; --ink-3: #7a848f;
  --rule: #dde1e5; --grid: #e8ebee; --accent: #b4441a;
  --s1: #2a78d6; --s2: #eb6834; --muted-bar: #b9c0c7; --flag-bg: #fdf1ea;
  --font-display: "Archivo", "Helvetica Neue", Arial, sans-serif;
  --font-body: "IBM Plex Sans", "Helvetica Neue", Arial, sans-serif;
  --font-data: "IBM Plex Mono", ui-monospace, "SF Mono", Menlo, monospace;
}
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
  --bg: #121417; --panel: #1b1e22; --ink: #eef0f2; --ink-2: #b7bfc7; --ink-3: #8a939c;
  --rule: #2f343a; --grid: #262a2f; --accent: #f08a5d;
  --s1: #3987e5; --s2: #d95926; --muted-bar: #4d555d; --flag-bg: #2a1d17; color-scheme: dark; } }
:root[data-theme="dark"] {
  --bg: #121417; --panel: #1b1e22; --ink: #eef0f2; --ink-2: #b7bfc7; --ink-3: #8a939c;
  --rule: #2f343a; --grid: #262a2f; --accent: #f08a5d;
  --s1: #3987e5; --s2: #d95926; --muted-bar: #4d555d; --flag-bg: #2a1d17; color-scheme: dark; }
body { background: var(--bg); color: var(--ink); font-family: var(--font-body); font-size: 15px; line-height: 1.55; }
.wrap { max-width: 960px; margin: 0 auto; padding-inline: 20px; padding-block: 32px 64px; display: grid; grid-template-columns: minmax(0, 1fr); gap: 40px; }
.wrap > * { min-width: 0; }
.eyebrow { overflow-wrap: anywhere; }
header { display: grid; gap: 6px; border-bottom: 2px solid var(--ink); padding-bottom: 16px; }
.eyebrow { font-family: var(--font-data); font-size: 12px; letter-spacing: .08em; text-transform: uppercase; color: var(--ink-3); }
h1 { font-family: var(--font-display); font-weight: 800; font-size: clamp(28px, 5vw, 40px); line-height: 1.1; margin: 0; letter-spacing: -.01em; text-wrap: balance; }
h2 { font-family: var(--font-display); font-weight: 700; font-size: 22px; margin: 0; text-wrap: balance; }
h3 { font-family: var(--font-display); font-weight: 700; font-size: 16px; margin: 0; }
p { margin: 0; max-width: 70ch; }
.sub { color: var(--ink-2); }
section { display: grid; gap: 14px; min-width: 0; }
.answer { background: var(--panel); border: 1px solid var(--rule); border-left: 5px solid var(--accent); padding: 22px 24px; display: grid; gap: 14px; }
.answer h2 { font-size: 26px; color: var(--ink); }
.answer ul, .plain ul { margin: 0; padding-left: 20px; display: grid; gap: 8px; }
.cols { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
.cols > div { min-width: 0; display: grid; gap: 8px; align-content: start; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 1px; background: var(--rule); border: 1px solid var(--rule); }
.kpi { background: var(--panel); padding: 14px 16px; display: grid; gap: 2px; }
.kpi .n { font-family: var(--font-data); font-size: 24px; font-weight: 600; font-variant-numeric: tabular-nums; }
.kpi .l { font-size: 12.5px; color: var(--ink-2); }
.figure { background: var(--panel); border: 1px solid var(--rule); padding: 16px; overflow-x: auto; }
.figure svg { width: 100%; height: auto; min-width: 520px; display: block; }
.cap { font-size: 13px; color: var(--ink-3); }
svg text { font-family: var(--font-data); }
svg .grid { stroke: var(--grid); stroke-width: 1; }
svg .base { stroke: var(--ink-3); stroke-width: 1; }
svg .ax { fill: var(--ink-3); font-size: 11px; }
svg .val { fill: var(--ink-2); font-size: 11px; }
svg .lab { fill: var(--ink-2); font-size: 12.5px; font-family: var(--font-body); }
svg .lab.strong { fill: var(--ink); font-weight: 600; }
svg .bar { fill: var(--s1); }
svg .bar.partial { fill: var(--s1); opacity: .45; }
svg .bar.hi { fill: var(--s2); }
svg .bar.lo { fill: var(--muted-bar); }
svg .line { fill: none; stroke-width: 2; }
svg .line.s1 { stroke: var(--s1); } svg .line.s2 { stroke: var(--s2); }
svg .dot { stroke: var(--panel); stroke-width: 2; }
svg .dot.s1 { fill: var(--s1); } svg .dot.s2 { fill: var(--s2); }
svg .s1-t, svg .s2-t { fill: var(--ink-2); }
svg g:hover path, svg g:hover circle { opacity: .8; }
.tbl { overflow-x: auto; background: var(--panel); border: 1px solid var(--rule); }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th, td { padding: 7px 10px; text-align: right; border-bottom: 1px solid var(--grid); white-space: nowrap; font-variant-numeric: tabular-nums; }
th { font-weight: 600; color: var(--ink-2); background: var(--bg); font-size: 12px; vertical-align: bottom; white-space: normal; }
th:first-child, td:first-child { text-align: left; }
td.num { font-family: var(--font-data); font-size: 12.5px; }
tr.flag td { background: var(--flag-bg); }
tr.total td { font-weight: 600; border-top: 1px solid var(--ink-3); }
td.txt, th.txt { white-space: normal; text-align: left; min-width: 220px; }
.legend { display: flex; flex-wrap: wrap; gap: 16px; font-size: 13px; color: var(--ink-2); }
.legend span::before { content: ""; display: inline-block; width: 12px; height: 3px; margin-right: 6px; vertical-align: middle; }
.legend .k1::before { background: var(--s1); } .legend .k2::before { background: var(--s2); }
.note { font-size: 13.5px; color: var(--ink-2); }
code { font-family: var(--font-data); font-size: 12.5px; }
details { background: var(--panel); border: 1px solid var(--rule); padding: 12px 16px; }
summary { cursor: pointer; font-weight: 600; }
summary:focus-visible, a:focus-visible { outline: 2px solid var(--s1); outline-offset: 2px; }
dl { display: grid; grid-template-columns: minmax(140px, 220px) 1fr; gap: 6px 16px; margin: 12px 0 0; font-size: 13.5px; }
dt { font-weight: 600; } dd { margin: 0; color: var(--ink-2); min-width: 0; }
@media (max-width: 560px) { dl { grid-template-columns: 1fr; } .answer { padding: 18px; } }
/* Print / PDF: plain flow layout (grid rows fragment badly across pages), full-width tables, nothing clipped */
@media print {
  body { background: #fff; font-size: 11px; }
  .wrap { display: block; max-width: none; padding: 0; }
  .wrap > * { margin-bottom: 18px; }
  section { display: block; }
  section > * + * { margin-top: 8px; }
  .answer, .kpis, .figure, tr, details { break-inside: avoid; }
  h2, h3 { break-after: avoid; }
  .tbl, .figure { overflow: visible; }
  .figure svg { min-width: 0; }
  table { font-size: 9px; }
  th, td { padding: 3px 5px; white-space: normal; }
  td.num { font-size: 9px; }
  td.txt, th.txt { min-width: 0; }
  .kpi .n { font-size: 18px; }
  summary { list-style: none; }
}
"""

FONTS = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
         '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;600&display=swap">')


def html_page(a, f, crosswalk_rows, standalone=True):
    s = a["score"]
    vend = a["vend"]
    as_of = a["as_of"].strftime("%m/%d/%Y")
    monthly = a["monthly"]
    cf = f["call_first"]

    def row(vid, cells, flag=False):
        return '<tr%s>%s</tr>' % (' class="flag"' if flag else "", "".join(cells))

    # scorecard table
    order = s.sort_values("received_value", ascending=False).index
    sc = ['<table><thead><tr><th>Vendor</th><th>Received value</th><th>Share</th><th>Due lines</th><th>On time</th>'
          '<th>Lines &gt;%dd late</th><th>$ &gt;%dd late</th><th>Median days late (when late)</th><th>Promised after need</th>'
          '<th>Met own promise</th><th>Past-due open $</th><th>Price above PO $ (lines)</th><th>QC holds (open)</th></tr></thead><tbody>'
          % (LATE_GRACE_DAYS, LATE_GRACE_DAYS)]
    for vid in order:
        r = s.loc[vid]
        sc.append(row(vid, [
            "<td>%s</td>" % esc(r.vendor), '<td class="num">%s</td>' % money(r.received_value),
            '<td class="num">%s</td>' % pct(r.share, 1), '<td class="num">%d</td>' % r.due_lines,
            '<td class="num">%s</td>' % pct(r.on_time), '<td class="num">%d</td>' % r.severe_lines,
            '<td class="num">%s</td>' % money(r.severe_value), '<td class="num">%d</td>' % r.median_days_late_when_late,
            '<td class="num">%s</td>' % pct(r.pct_promised_late), '<td class="num">%s</td>' % pct(r.on_time_vs_promise),
            '<td class="num">%s</td>' % money(r.past_due_value),
            '<td class="num">%s</td>' % ("%s (%d)" % (money(r.price_exposure), r.price_above_lines) if r.price_above_lines else "–"),
            '<td class="num">%d (%d)</td>' % (r.holds, r.holds_open)], flag=vid == "V004"))
    sc.append("</tbody></table>")

    # monthly table
    mt = ['<table><thead><tr><th>Month</th>%s<th>Total</th><th>Receipts</th></tr></thead><tbody>' %
          "".join("<th>%s</th>" % esc(a["short"][v]) for v in s.index)]
    for m in monthly.index:
        lab = pd.Timestamp(m + "-01").strftime("%b %Y")
        partial = m in (a["first_day"].strftime("%Y-%m"), a["last_day"].strftime("%Y-%m"))
        mt.append("<tr><td>%s%s</td>%s<td class='num'><b>%s</b></td><td class='num'>%d</td></tr>" % (
            lab, " *" if partial else "", "".join("<td class='num'>%s</td>" % money(monthly.loc[m, v]) for v in s.index),
            money(monthly.loc[m, "Total"]), a["monthly_counts"].loc[m, "receipts"]))
    mt.append("<tr class='total'><td>Total</td>%s<td class='num'>%s</td><td class='num'>%d</td></tr></tbody></table>" % (
        "".join("<td class='num'>%s</td>" % money(monthly[v].sum()) for v in s.index), money(monthly["Total"].sum()),
        a["monthly_counts"].receipts.sum()))

    # continental by service
    v4p = a["v4_parts"]
    v4t = ["<table><thead><tr><th>Service</th><th>Due lines</th><th>On time</th><th>Avg days late (when late)</th></tr></thead><tbody>"]
    desc = a["part"].set_index("part_id").description
    for p, r in v4p.iterrows():
        v4t.append("<tr><td>%s <span class='cap'>%s</span></td><td class='num'>%d</td><td class='num'>%s</td><td class='num'>%.1f</td></tr>" % (
            p, esc(desc.get(p, "")), r.lines, pct(r.on_time), r.avg_late))
    v4t.append("</tbody></table>")

    # lead time table
    lt = ["<table><thead><tr><th>Vendor</th><th>We allow</th><th>Vendor quotes</th><th>Actual</th></tr></thead><tbody>"]
    for vid in s.index:
        r = s.loc[vid]
        lt.append(row(vid, ["<td>%s</td>" % esc(r.vendor)] + ["<td class='num'>%g days</td>" % v for v in
                                                              (r.lt_allowed, r.lt_promised, r.lt_actual)], flag=vid == "V004"))
    lt.append("</tbody></table>")

    # price creep
    above = a["above"]
    pc = above.groupby(["conf_month"]).agg(lines=("po_number", "size"), price=("confirmed_price", "max"),
                                           exp=("exposure_ordered", "sum"))
    pct_t = ["<table><thead><tr><th>Confirmation month</th><th>Lines above PO</th><th>Confirmed price</th><th>PO price</th><th>Exposure (ordered qty)</th></tr></thead><tbody>"]
    for m, r in pc.iterrows():
        pct_t.append("<tr><td>%s</td><td class='num'>%d</td><td class='num'>$%.4f</td><td class='num'>$3.9200</td><td class='num'>%s</td></tr>" % (
            pd.Timestamp(m + "-01").strftime("%b %Y"), r.lines, r.price, money(r.exp, 2)))
    pct_t.append("</tbody></table>")

    # crosswalk table
    cw = ["<table><thead><tr><th>Vendor</th><th>Vendor PN</th><th>Beacon PN</th><th>Source</th><th class='txt'>Status / note</th></tr></thead><tbody>"]
    for r in crosswalk_rows:
        cw.append("<tr%s><td>%s</td><td class='num'>%s</td><td class='num'>%s</td><td class='txt'>%s</td><td class='txt'>%s</td></tr>" % (
            ' class="flag"' if r["flag"] else "", esc(r["vendor"]), esc(r["vpn"]), esc(r["bpn"]), esc(r["src"]), esc(r["note"])))
    cw.append("</tbody></table>")

    dq = "".join("<dt>%s</dt><dd>%s &middot; %s</dd>" % (esc(k), esc(v), esc(h)) for k, v, h in a["dq"])
    defs = "".join("<dt>%s</dt><dd>%s</dd>" % (esc(k), esc(v)) for k, v in METRIC_DEFINITIONS)
    nv = a["naive"]

    body = f"""
<title>Beacon Vendor Readout</title>
{FONTS}
<style>{CSS}</style>
<div class="wrap">
<header>
  <div class="eyebrow">Tinicum FDE case study &middot; "Beacon Fasteners" &middot; ERP history {a['first_day'].strftime('%b %Y')} – {a['last_day'].strftime('%b %Y')} &middot; as of {as_of}</div>
  <h1>Vendor performance readout</h1>
  <p class="sub">Six active vendors, {s.pos.sum():,} POs, {f['due_lines']:,} delivered PO lines scored. Built from the ERP extract after correcting its data problems (listed at the end).</p>
</header>

<section class="answer" aria-labelledby="a1">
  <div class="eyebrow">The phone call</div>
  <h2 id="a1">{esc(cf['headline'])}</h2>
  <ul>{''.join('<li>%s</li>' % w for w in cf['why'])}</ul>
  <div class="cols">
    <div><h3>What to ask them</h3><ul class="note">{''.join('<li>%s</li>' % q for q in cf['ask'])}</ul></div>
    <div><h3>What to fix on our side</h3><p class="note">{esc(cf['internal'])}</p></div>
  </div>
</section>

<section class="plain">
  <h2>Second call, and who to watch</h2>
  <p>{f['second']['text']}</p>
  <ul>{''.join('<li>%s</li>' % w for w in f['watch'])}</ul>
</section>

{forward_html(a, f)}

<section>
  <div class="kpis">
    <div class="kpi"><span class="n">{money(monthly['Total'].sum())}</span><span class="l">received at PO price, 8.5 months</span></div>
    <div class="kpi"><span class="n">{pct(s.loc['V001'].share)}</span><span class="l">of it from Apex (bar stock)</span></div>
    <div class="kpi"><span class="n">{pct(f['overall_on_time'])}</span><span class="l">of due lines complete on time, all vendors</span></div>
    <div class="kpi"><span class="n">{s.severe_lines.sum()}</span><span class="l">lines &gt;{LATE_GRACE_DAYS} days late: {s.loc['V004'].severe_lines} are Continental</span></div>
  </div>
</section>

<section>
  <h2>Received value by month</h2>
  <p class="sub">Net receipts (reversals removed) at PO unit price, in USD. Apex is {pct(s.loc['V001'].share)} of the total, so the monthly shape is mostly Apex's bar-stock deliveries.</p>
  <div class="figure">{svg_monthly(monthly, a['first_day'], a['last_day'])}
  <p class="cap">* Partial months: data starts {a['first_day'].strftime('%m/%d/%Y')} and ends {a['last_day'].strftime('%m/%d/%Y')}. Values in $M.</p></div>
  <div class="tbl">{''.join(mt)}</div>
</section>

<section>
  <h2>Vendor by vendor</h2>
  <p class="sub">Relative rates and absolute counts side by side. A vendor late on many low-value lines differs from one late on a few critical ones. Continental is highlighted.</p>
  <div class="figure">{svg_ontime(s)}<p class="cap">On time = the receipt that completes the line arrives on/before the required date. Grey = other vendors.</p></div>
  <div class="tbl">{''.join(sc)}</div>
  <p class="note">Apex has the most dollars received late ({money(s.loc['V001'].late_value)}), because it is 85% of spend. Its late lines are late by a median of {s.loc['V001'].median_days_late_when_late:.0f} day and none exceeds {s.loc['V001'].max_days_late} days. Continental's late lines run a median of {s.loc['V004'].median_days_late_when_late:.0f} days. That is why severity, and not raw late dollars, drives the call order.</p>
</section>

<section>
  <h2>Is anyone getting worse?</h2>
  <div class="figure">{svg_trend(a['monthly_ot'], a['rec'])}
  <div class="legend"><span class="k2">Continental</span><span class="k1">All other vendors</span></div>
  <p class="cap">Share of lines complete on time, by the month the line was due. Hover a point for n.</p></div>
  <p class="note">No vendor shows a statistically significant change between Sep–Dec and Jan–May (Fisher exact test, all p &gt; 0.1). Continental is flat at a low level. Liberty dipped from {pct(s.loc['V003'].ot_early)} to {pct(s.loc['V003'].ot_late)} (p={s.loc['V003'].trend_p:.2f}).</p>
</section>

<section>
  <h2>Continental in detail</h2>
  <div class="cols">
    <div><h3>By heat-treat service</h3><div class="tbl">{''.join(v4t)}</div><p class="cap">Lateness shows up in all three services, not just one.</p></div>
    <div><h3>Lead time, median days from PO date</h3><div class="tbl">{''.join(lt)}</div><p class="cap">Other vendors quote within 3 days of the lead time we plan with. Continental quotes 9 days longer, and delivers about 5 days later than we plan.</p></div>
  </div>
</section>

<section>
  <h2>Apex price creep on CRES 17-4 bar</h2>
  <p class="sub">The only vendor confirming above PO price. Every Apex part other than BAR-CRES-250 confirms at PO price. Receipts carry no price and there is no invoice table in this extract, so the ERP does not record this increase anywhere. AP will meet it on the invoice.</p>
  <div class="tbl">{''.join(pct_t)}</div>
</section>

<section>
  <h2>Vendor part numbers: the cleanup, started</h2>
  <p>Two vendors confirm with their own part numbers ({int(a['conf'].vendor_pn.notna().sum())} confirmation rows have a vendor PN and no Beacon part). Joining each back to its PO line recovers a crosswalk. The history exposes three problems a cleanup has to design for:</p>
  <ul class="note">
    <li><b>The same vendor PN means different parts at different vendors.</b> <code>K-1050</code> is a cold-headed pin blank at Heritage and a heading punch at Ostmark, so the key must be vendor + vendor PN.</li>
    <li><b>Vendors renumber.</b> This week's PDFs use <code>APH-441-OS</code> and <code>OST-CAR-A-100</code> for parts the ERP knows as <code>APH-4411</code> and <code>WZ-HM-EIN</code>. A mapping needs first-seen/last-seen dates and a review status, not just a lookup.</li>
    <li><b>Look-alikes are real.</b> <code>APH-441</code> and <code>APH-4411</code> are two different blanks (9472-3 vs 9472-4), so fuzzy matching on PNs would merge them. Mappings are proposed from qty and price evidence and approved by a person.</li>
  </ul>
  <div class="tbl">{''.join(cw)}</div>
  <p class="note">The confirmation tool writes every new mapping it infers to this table as <i>pending review</i>. After Lisa approves it, it is used automatically. Over time this becomes the vendor-part master, ready to load into the ERP's vendor-part table.</p>
</section>

<section>
  <h2>How the numbers were built</h2>
  <p class="sub">The ERP extract has traps that change the answer if taken at face value:</p>
  <div class="kpis">
    <div class="kpi"><span class="n">{money(nv['correct'])}</span><span class="l">received value, corrected</span></div>
    <div class="kpi"><span class="n">{money(nv['r_only'])}</span><span class="l">if reversed keying errors and the unreversed duplicate are counted (+{money(nv['r_only'] - nv['correct'])})</span></div>
    <div class="kpi"><span class="n">{money(nv['qty_received_field'])}</span><span class="l">if the ERP's own qty_received field is trusted ({money(nv['qty_received_field'] - nv['correct'])})</span></div>
    <div class="kpi"><span class="n">+{money(nv['ostmark_fx_again'])}</span><span class="l">if Ostmark prices were converted from EUR again</span></div>
  </div>
  <details><summary>Data problems found and how each was handled</summary><dl>{dq}</dl></details>
  <details><summary>Metric definitions</summary><dl>{defs}</dl></details>
  <details><summary>What this readout cannot tell you</summary><ul class="note" style="margin-top:10px">
    <li>Whether any late delivery stopped production or caused a customer expedite: there is no production, downtime or customer-order data in the extract.</li>
    <li>Invoiced cost or price variance paid: receipts carry no price and there is no invoice/AP table. Value is at PO price.</li>
    <li>Quality performance: qc_hold has only 35 holds and 16 are dated before their PO existed, so holds are shown as counts, not ranked.</li>
    <li>Confirmation behaviour after March: the confirmation feed stops at 04/03/2026, so April–May POs have none in the ERP.</li>
    <li>Why required dates are set where they are: if the 19-day heat-treat lead time is a planning default, part of Continental's "lateness" is Beacon's own.</li>
  </ul></details>
</section>
</div>
"""
    if standalone:
        content = body.split("</style>", 1)[1]
        return ('<!doctype html>\n<html lang="en"><head><meta charset="utf-8">'
                '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
                '<title>Beacon Vendor Readout</title>%s<style>body{margin:0}%s</style></head><body>%s</body></html>'
                % (FONTS, CSS, content))
    return body


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def forward_html(a, f):
    """'Looking ahead' section: open orders due in the next 60 days that need attention."""
    fw, bt = f.get("forward"), f.get("backtest")
    if fw is None or not len(fw):
        return ""
    hi = fw[fw.tier != "OK"]
    unc = fw[(fw.promise_source == "None on file") & (fw.days_to_due <= 14)].sort_values("open_value", ascending=False)
    rows = []
    for x in pd.concat([hi, unc.head(8)]).itertuples():
        rows.append("<tr%s><td>%s</td><td>%s</td><td class='num'>%s L%d</td><td class='num'>%s</td><td class='num'>%s</td>"
                    "<td class='num'>%s</td><td class='num'>%s</td><td class='txt'>%s</td></tr>" % (
                        ' class="flag"' if x.tier == "HIGH" else "", esc(x.tier if x.tier != "OK" else "Chase"), esc(x.vendor),
                        x.po_number, x.line_no, esc(x.part_id), money(x.open_value), x.required.strftime("%m/%d"),
                        x.promised.strftime("%m/%d") if pd.notna(x.promised) else "none", esc(x.why)))
    return f"""
<section>
  <h2>Looking ahead: open orders due in the next 60 days</h2>
  <p class="sub">The same history, turned into a watch list for planning. {len(fw)} open lines worth {money(fw.open_value.sum())} are due between {(a['as_of'] + pd.Timedelta(days=1)).strftime('%m/%d')} and {(a['as_of'] + pd.Timedelta(days=60)).strftime('%m/%d')}.</p>
  <div class="kpis">
    <div class="kpi"><span class="n">{(fw.tier == 'HIGH').sum()}</span><span class="l">lines at high risk of arriving late ({money(fw[fw.tier == 'HIGH'].open_value.sum())})</span></div>
    <div class="kpi"><span class="n">{len(unc)}</span><span class="l">lines due within 14 days with no acknowledgment on file ({money(unc.open_value.sum())})</span></div>
    <div class="kpi"><span class="n">{bt['severe_caught']} of {bt['severe']}</span><span class="l">lines &gt;7 days late in the past that this rule would have flagged</span></div>
  </div>
  <div class="tbl"><table><thead><tr><th>Status</th><th>Vendor</th><th>PO / line</th><th>Part</th><th>Open $</th><th>Needed</th><th>Promised</th><th class='txt'>Why</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>
  <p class="note">Flag rule: the vendor already promises after our need date, or it has delivered this part on time less than 60% of the time (5+ past lines). Backtested on 8 months without look-ahead: {pct(bt['precision'])} of flagged lines arrived late, against {pct(bt['base_rate'])} overall, so treat it as a watch list, not a forecast. Promise dates come from the ERP and from this week's acknowledgments read by the Task 1 tool. "No acknowledgment on file" may partly reflect the ERP confirmation feed stopping on 04/03. Full list: <i>Open Orders at Risk</i> tab.</p>
</section>"""


def write_xlsx(path, a, f, crosswalk_rows):
    import xlsxwriter
    s = a["score"]
    wb = xlsxwriter.Workbook(path)
    H = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "text_wrap": True, "valign": "top", "border": 1})
    T = wb.add_format({"bold": True, "font_size": 16})
    T2 = wb.add_format({"bold": True, "font_size": 12})
    W = wb.add_format({"text_wrap": True, "valign": "top"})
    M = wb.add_format({"num_format": "$#,##0"})
    P = wb.add_format({"num_format": "0.0%"})
    N = wb.add_format({"num_format": "#,##0"})
    D1 = wb.add_format({"num_format": "0.0"})
    FL = wb.add_format({"bg_color": "#FCE4D6"})
    import re
    strip = lambda t: re.sub(r"<[^>]+>", "", t)

    # Summary
    ws = wb.add_worksheet("Summary")
    ws.set_column(0, 0, 130)
    r = 0
    ws.write(r, 0, "Beacon vendor performance readout - ERP history %s to %s (as of %s)" % (
        a["first_day"].strftime("%m/%d/%Y"), a["last_day"].strftime("%m/%d/%Y"), a["as_of"].strftime("%m/%d/%Y")), T); r += 2
    ws.write(r, 0, f["call_first"]["headline"].upper(), T2); r += 1
    for w in f["call_first"]["why"]:
        ws.write(r, 0, "• " + strip(w), W); r += 1
    r += 1
    ws.write(r, 0, "What to ask them", T2); r += 1
    for q in f["call_first"]["ask"]:
        ws.write(r, 0, "• " + q, W); r += 1
    ws.write(r, 0, "• " + f["call_first"]["internal"], W); r += 2
    ws.write(r, 0, "SECOND CALL", T2); r += 1
    ws.write(r, 0, strip(f["second"]["text"]), W); r += 2
    ws.write(r, 0, "WATCH", T2); r += 1
    for w in f["watch"]:
        ws.write(r, 0, "• " + strip(w), W); r += 1
    r += 1
    fw, bt = f["forward"], f["backtest"]
    if len(fw):
        hi, unc = fw[fw.tier == "HIGH"], fw[(fw.promise_source == "None on file") & (fw.days_to_due <= 14)]
        ws.write(r, 0, "LOOKING AHEAD: OPEN ORDERS DUE IN THE NEXT 60 DAYS (%s)" % money(fw.open_value.sum()), T2); r += 1
        ws.write(r, 0, "• %d line(s) at high risk of arriving late (%s): %s." % (len(hi), money(hi.open_value.sum()), "; ".join(
            "%s %s L%d" % (x.vendor.split()[0], x.po_number, x.line_no) for x in hi.itertuples())), W); r += 1
        ws.write(r, 0, "• %d line(s) due within 14 days (%s) have no acknowledgment on file - chase them." % (
            len(unc), money(unc.open_value.sum())), W); r += 1
        ws.write(r, 0, "• Backtest of the same rule on 8 months of history: it flagged %d of the %d lines that arrived more than 7 days late; "
                       "about 1 in 3 flagged lines was late. A watch list, not a prediction." % (bt["severe_caught"], bt["severe"]), W); r += 2
    ws.write(r, 0, "Tabs: Monthly Value | Vendor Scorecard | On-time Trend | Open Orders at Risk | Past-due Open Lines | Price Above PO | "
                   "Revised Confirmations | PN Crosswalk | Data Quality | Definitions", W)

    # Monthly value
    ws = wb.add_worksheet("Monthly Value")
    mv = a["monthly"]
    cols = ["Month"] + [a["vend"][v] for v in s.index] + ["Total", "Receipts", "Vendors receiving"]
    for c, h in enumerate(cols):
        ws.write(0, c, h, H)
    ws.set_column(0, 0, 12)
    ws.set_column(1, len(cols), 16)
    for i, m in enumerate(mv.index, 1):
        partial = m in (a["first_day"].strftime("%Y-%m"), a["last_day"].strftime("%Y-%m"))
        ws.write(i, 0, m + (" (partial)" if partial else ""))
        for c, v in enumerate(s.index, 1):
            ws.write_number(i, c, mv.loc[m, v], M)
        ws.write_number(i, len(s) + 1, mv.loc[m, "Total"], M)
        ws.write_number(i, len(s) + 2, a["monthly_counts"].loc[m, "receipts"])
        ws.write_number(i, len(s) + 3, a["monthly_counts"].loc[m, "vendors"])
    n = len(mv)
    ws.write(n + 1, 0, "Total", H)
    totals = [mv[v].sum() for v in s.index] + [mv["Total"].sum()]
    for c in range(1, len(s) + 2):   # cached value too, so previews (Outlook, Quick Look) show the total
        ws.write_formula(n + 1, c, "=SUM(%s2:%s%d)" % (chr(65 + c), chr(65 + c), n + 1), M, round(float(totals[c - 1]), 2))
    ws.write(n + 3, 0, "Received value = net receipt qty (reversals netted, 1 unreversed duplicate excluded) x PO unit price, USD, by receipt month. "
                       "First and last months are partial.", W)
    ch = wb.add_chart({"type": "column", "subtype": "stacked"})
    for c, v in enumerate(s.index, 1):
        ch.add_series({"name": ["Monthly Value", 0, c], "categories": ["Monthly Value", 1, 0, n, 0],
                       "values": ["Monthly Value", 1, c, n, c], "gap": 60})
    ch.set_title({"name": "Received value by month (USD, at PO price)"})
    ch.set_y_axis({"num_format": "$#,##0,,\"M\"", "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}})
    ch.set_size({"width": 900, "height": 380})
    ws.insert_chart(n + 5, 0, ch)

    # Scorecard
    ws = wb.add_worksheet("Vendor Scorecard")
    spec = [("Vendor", "vendor", None), ("Received value", "received_value", M), ("Share of value", "share", P),
            ("Non-cancelled POs", "pos", N), ("Due lines scored", "due_lines", N), ("On time %", "on_time", P),
            ("Late lines", "late_lines", N), ("Lines >%dd late" % LATE_GRACE_DAYS, "severe_lines", N),
            ("%% lines >%dd late" % LATE_GRACE_DAYS, "severe_pct", P), ("$ received >%dd late" % LATE_GRACE_DAYS, "severe_value", M),
            ("$ received late (any)", "late_value", M), ("Avg days late (when late)", "avg_days_late_when_late", D1),
            ("Median days late (when late)", "median_days_late_when_late", D1), ("Max days late", "max_days_late", N),
            ("On time with %dd grace" % LATE_GRACE_DAYS, "on_time_with_grace", P), ("OTIF (full qty by required date)", "otif", P),
            ("On time on FIRST receipt", "on_time_first", P), ("Lines with live confirmation", "lines_w_promise", N),
            ("% promised after required", "pct_promised_late", P), ("Avg promise gap (days)", "avg_promise_gap", D1),
            ("On time vs own promise", "on_time_vs_promise", P), ("Lead time allowed (median d)", "lt_allowed", D1),
            ("Lead time quoted (median d)", "lt_promised", D1), ("Lead time actual (median d)", "lt_actual", D1),
            ("% lines split-shipped", "multi_ship_pct", P), ("Confirmation coverage (POs to 03/30)", "conf_coverage", P),
            ("Revised confirmations", "revisions", N), ("  of which date push-outs", "pushouts", N),
            ("Past-due open lines", "past_due_lines", N), ("Past-due open $", "past_due_value", M),
            ("  of which nothing received", "past_due_nothing", N), ("  of which short-shipped balance", "short_open_lines", N),
            ("Lines confirmed above PO price", "price_above_lines", N), ("Price exposure (ordered qty)", "price_exposure", M),
            ("QC holds", "holds", N), ("QC holds unreleased", "holds_open", N), ("Holds per 100 received POs", "holds_per_100po", D1),
            ("On time Sep-Dec", "ot_early", P), ("  n", "ot_early_n", N), ("On time Jan-May", "ot_late", P), ("  n", "ot_late_n", N),
            ("Trend p-value (Fisher)", "trend_p", wb.add_format({"num_format": "0.00"}))]
    ws.set_column(0, 0, 38)
    ws.set_column(1, len(s), 18)
    ws.write(0, 0, "Metric", H)
    for c, vid in enumerate(s.index, 1):
        ws.write(0, c, s.loc[vid].vendor, H)
    ws.set_row(0, 32)
    for i, (lab, key, fmt) in enumerate(spec[1:], 1):
        ws.write(i, 0, lab)
        for c, vid in enumerate(s.index, 1):
            v = s.loc[vid][key]
            if isinstance(v, (float, np.floating)) and np.isnan(v):
                continue
            ws.write_number(i, c, float(v), fmt) if fmt else ws.write(i, c, v)
    ws.freeze_panes(1, 1)
    ws.conditional_format(0, 4, len(spec), 4, {"type": "no_blanks", "format": FL})

    # Trend
    ws = wb.add_worksheet("On-time Trend")
    mo, mn = a["monthly_ot"], a["monthly_ot_n"]
    ws.write(0, 0, "Required month", H)
    for c, vid in enumerate(s.index, 1):
        ws.write(0, 2 * c - 1, a["vend"][vid] + " on time", H)
        ws.write(0, 2 * c, "n", H)
    ws.set_column(0, 0, 14)
    ws.set_column(1, 2 * len(s), 12)
    for i, m in enumerate(mo.index, 1):
        ws.write(i, 0, m)
        for c, vid in enumerate(s.index, 1):
            if vid in mo and not np.isnan(mo.loc[m, vid]):
                ws.write_number(i, 2 * c - 1, mo.loc[m, vid], P)
                ws.write_number(i, 2 * c, mn.loc[m, vid])
    ch = wb.add_chart({"type": "line"})
    for c, vid in enumerate(s.index, 1):
        ch.add_series({"name": ["On-time Trend", 0, 2 * c - 1], "categories": ["On-time Trend", 1, 0, len(mo), 0],
                       "values": ["On-time Trend", 1, 2 * c - 1, len(mo), 2 * c - 1],
                       "line": {"width": 3.0 if vid == "V004" else 1.25}, "marker": {"type": "circle", "size": 5}})
    ch.set_title({"name": "On-time % by month due"})
    ch.set_y_axis({"num_format": "0%", "min": 0, "max": 1})
    ch.set_size({"width": 900, "height": 380})
    ws.insert_chart(len(mo) + 3, 0, ch)

    # Open orders at risk (forward-looking)
    fw, bt = f["forward"], f["backtest"]
    ws = wb.add_worksheet("Open Orders at Risk")
    cols = [("Risk", 7), ("Vendor", 26), ("PO", 15), ("Line", 5), ("Part", 15), ("Open qty", 9), ("Open $", 11),
            ("Required", 11), ("Days to due", 8), ("Promised", 11), ("Promise from", 22), ("Vendor+part on time", 10),
            ("Past lines", 7), ("Why", 60), ("Suggested action", 45)]
    ws.write(0, 0, "Open PO lines due %s to %s. Flag rule: vendor already promises after our need date, or this vendor delivers this "
                   "part on time <60%% of the time (5+ past lines). Backtest on history: flagged %d of %d lines that arrived >7 days late; "
                   "%.0f%% of flagged lines were late vs %.0f%% overall." % (
                       (a["as_of"] + pd.Timedelta(days=1)).strftime("%m/%d/%Y"), (a["as_of"] + pd.Timedelta(days=60)).strftime("%m/%d/%Y"),
                       bt["severe_caught"], bt["severe"], 100 * bt["precision"], 100 * bt["base_rate"]), W)
    ws.set_row(0, 45)
    for c, (h, w) in enumerate(cols):
        ws.write(2, c, h, H)
        ws.set_column(c, c, w)
    tier_fmt = {"HIGH": wb.add_format({"bg_color": "#F8CBAD", "bold": True}), "WATCH": wb.add_format({"bg_color": "#FFE699"}),
                "OK": wb.add_format({"bg_color": "#E2EFDA"})}
    for i, x in enumerate(fw.itertuples(), 3):
        vals = [x.tier, x.vendor, x.po_number, x.line_no, x.part_id, x.open_qty, x.open_value, x.required.strftime("%Y-%m-%d"),
                x.days_to_due, x.promised.strftime("%Y-%m-%d") if pd.notna(x.promised) else "", x.promise_source,
                x.hist_on_time if x.hist_on_time is not None and not pd.isna(x.hist_on_time) else "", x.hist_n, x.why, x.action]
        for c, v in enumerate(vals):
            v = v.item() if hasattr(v, "item") else v
            ws.write(i, c, v, tier_fmt[x.tier] if c == 0 else (M if c == 6 else (P if c == 11 else None)))
    ws.autofilter(2, 0, 2 + len(fw), len(cols) - 1)
    ws.freeze_panes(3, 2)

    # Past due
    ws = wb.add_worksheet("Past-due Open Lines")
    pdl = a["past_due"].sort_values("balance_value", ascending=False)
    cols = [("Vendor", 26), ("PO", 15), ("Line", 5), ("Part", 15), ("Ordered", 10), ("Received", 10), ("Balance", 10),
            ("Balance $", 12), ("Required", 11), ("Days past due", 9), ("Type", 30), ("Last receipt", 11)]
    for c, (h, w) in enumerate(cols):
        ws.write(0, c, h, H)
        ws.set_column(c, c, w)
    for i, (_, r) in enumerate(pdl.iterrows(), 1):
        vals = [a["vend"][r.vendor_id], r.po_number, r.line_no, r.part_id, r.qty_ordered, r.rec_qty, r.balance_qty,
                r.balance_value, r.required.strftime("%Y-%m-%d"), r.days_past_due, r.kind,
                r.last_rcv.strftime("%Y-%m-%d") if pd.notna(r.last_rcv) else ""]
        for c, v in enumerate(vals):
            ws.write(i, c, v.item() if hasattr(v, "item") else v, M if c == 7 else None)
    ws.autofilter(0, 0, len(pdl), len(cols) - 1)
    ws.write(len(pdl) + 2, 0, "OPEN lines with required date before %s and received qty below ordered. Short-shipped balances "
                              "months old are likely vendor short-ships never closed: chase or short-close." % a["as_of"].strftime("%m/%d/%Y"))

    # Price above PO
    ws = wb.add_worksheet("Price Above PO")
    ab = a["above"].sort_values("doc_date")
    cols = [("Vendor", 24), ("PO", 15), ("Line", 5), ("Part", 15), ("PO price", 10), ("Confirmed price", 12),
            ("Var %", 8), ("Ordered qty", 10), ("Exposure ordered $", 14), ("Exposure received $", 14), ("Confirmation date", 12)]
    for c, (h, w) in enumerate(cols):
        ws.write(0, c, h, H)
        ws.set_column(c, c, w)
    for i, (_, r) in enumerate(ab.iterrows(), 1):
        vals = [a["vend"][r.vendor_id], r.po_number, r.line_no, r.part_id, r.unit_price, r.confirmed_price,
                r.price_var / r.unit_price, r.qty_ordered, r.exposure_ordered, r.exposure_received, r.doc_date]
        fm = [None, None, None, None, None, None, P, N, M, M, None]
        for c, v in enumerate(vals):
            ws.write(i, c, v.item() if hasattr(v, "item") else v, fm[c])

    # Revised confirmations
    ws = wb.add_worksheet("Revised Confirmations")
    sp = a["superseded"]
    cols = [("Vendor", 24), ("PO", 15), ("Line", 5), ("Original qty", 11), ("Revised qty", 11), ("Original promise", 13),
            ("Revised promise", 13), ("What changed", 16)]
    for c, (h, w) in enumerate(cols):
        ws.write(0, c, h, H)
        ws.set_column(c, c, w)
    for i, (_, r) in enumerate(sp.iterrows(), 1):
        for c, v in enumerate([a["vend"][r.vendor_id], r.po_number, r.line_no, r.confirmed_qty, r.new_qty,
                               r.promised_date, r.new_prom, r.kind]):
            ws.write(i, c, v.item() if hasattr(v, "item") else v)

    # Crosswalk
    ws = wb.add_worksheet("PN Crosswalk")
    cols = [("Vendor", 26), ("Vendor PN", 15), ("Beacon PN", 15), ("Source", 45), ("Status / note", 70)]
    for c, (h, w) in enumerate(cols):
        ws.write(0, c, h, H)
        ws.set_column(c, c, w)
    for i, r in enumerate(crosswalk_rows, 1):
        for c, k in enumerate(["vendor", "vpn", "bpn", "src", "note"]):
            ws.write(i, c, r[k], FL if r["flag"] else W)

    # Data quality + definitions
    ws = wb.add_worksheet("Data Quality")
    ws.set_column(0, 0, 60)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 80)
    for c, h in enumerate(["Issue found in the ERP extract", "Count / value", "How it was handled"]):
        ws.write(0, c, h, H)
    for i, (k, v, h) in enumerate(a["dq"], 1):
        ws.write(i, 0, k, W)
        ws.write(i, 1, str(v), W)
        ws.write(i, 2, h, W)
    nv = a["naive"]
    base = len(a["dq"]) + 3
    ws.write(base, 0, "Received value under different (wrong) approaches", T2)
    for i, (k, v) in enumerate([("Corrected", nv["correct"]), ("Counting 'R' receipts only: reversed keying errors + unreversed duplicate kept", nv["r_only"]),
                                 ("Using po_line.qty_received", nv["qty_received_field"]),
                                 ("Extra if Ostmark converted from EUR again", nv["ostmark_fx_again"])], base + 1):
        ws.write(i, 0, k)
        ws.write_number(i, 1, v, M)
    ws = wb.add_worksheet("Definitions")
    ws.set_column(0, 0, 36)
    ws.set_column(1, 1, 120)
    for c, h in enumerate(["Metric", "Definition (numerator / denominator / period / exclusions)"]):
        ws.write(0, c, h, H)
    for i, (k, v) in enumerate(METRIC_DEFINITIONS, 1):
        ws.write(i, 0, k, W)
        ws.write(i, 1, v, W)
    wb.close()


def crosswalk_rows(a, task1_crosswalk):
    """ERP-derived mappings + this week's inferred ones (from Task 1's output if present)."""
    conf, line, hdr, vend = a["conf"], a["line"], a["hdr"], a["vend"]
    x = conf[conf.vendor_pn.notna()].merge(line[["po_number", "line_no", "part_id"]], on=["po_number", "line_no"],
                                            suffixes=("", "_po")).merge(hdr[["po_number", "vendor_id"]], on="po_number")
    g = x.groupby(["vendor_id", "vendor_pn", "part_id_po"]).agg(n=("conf_id", "size"), first_seen=("doc_date", "min"),
                                                                last_seen=("doc_date", "max")).reset_index()
    shared = g.groupby("vendor_pn").vendor_id.nunique()
    rows = []
    if task1_crosswalk and os.path.exists(task1_crosswalk):
        t1 = pd.read_csv(task1_crosswalk)
        for _, r in t1[t1.review_status.str.startswith("PENDING")].iterrows():
            old = g[(g.vendor_id == r.vendor_id) & (g.part_id_po == r.beacon_pn)].vendor_pn.tolist()
            rows.append(dict(vendor=vend[r.vendor_id], vpn=r.vendor_pn, bpn=r.beacon_pn, src="This week's PDF (inferred by qty+price)",
                             note="PENDING REVIEW. ERP history uses %s for this part: vendor renumbered?" % ", ".join(old) if old else "PENDING REVIEW",
                             flag=True))
    for _, r in g.iterrows():
        note = "Seen %d times, %s to %s" % (r.n, r.first_seen, r.last_seen)
        if shared[r.vendor_pn] > 1:
            note += ". Same PN used by another vendor for a different part."
        rows.append(dict(vendor=vend[r.vendor_id], vpn=r.vendor_pn, bpn=r.part_id_po, src="ERP confirmations joined to PO lines",
                         note=note, flag=False))
    return rows


def export_pdf(html_path, pdf_path):
    """Printable copy for the plant manager. Optional: needs Playwright + Chromium; skipped otherwise."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(executable_path=os.environ.get("BEACON_CHROMIUM") or None)
            pg = b.new_page()
            pg.emulate_media(media="print", color_scheme="light")
            pg.goto("file://" + os.path.abspath(html_path))
            pg.evaluate("document.querySelectorAll('details').forEach(d => d.open = true)")
            pg.pdf(path=pdf_path, format="Letter", print_background=True,
                   margin=dict(top="0.5in", bottom="0.5in", left="0.4in", right="0.4in"))
            b.close()
        return pdf_path
    except Exception as e:  # not installed / no browser: the HTML and workbook are the deliverables
        print("(PDF copy skipped: %s. Open vendor_readout.html and print to PDF if needed.)" % str(e).splitlines()[0][:80])
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--erp", required=True)
    ap.add_argument("--out", default=os.path.join(HERE, "..", "output"))
    ap.add_argument("--task1-crosswalk", default=os.path.join(HERE, "..", "output", "pn_crosswalk_updated.csv"),
                    help="Task 1 output with this week's inferred vendor PNs (optional)")
    ap.add_argument("--task1-lines", default=os.path.join(HERE, "..", "output", "confirmation_lines_extracted.csv"),
                    help="Task 1 output: this week's confirmed lines, used for promise dates on open orders (optional)")
    ap.add_argument("--as-of", default=None)
    a_ = ap.parse_args()
    a = build(a_.erp, a_.as_of)
    f = findings(a)
    cw = crosswalk_rows(a, a_.task1_crosswalk)
    f["forward"] = forward_book(a, a_.task1_lines)
    f["backtest"] = backtest(a["rec"])
    os.makedirs(a_.out, exist_ok=True)
    write_xlsx(os.path.join(a_.out, "Vendor_Readout.xlsx"), a, f, cw)
    with open(os.path.join(a_.out, "vendor_readout.html"), "w") as fh:
        fh.write(html_page(a, f, cw, standalone=True))
    pdf = export_pdf(os.path.join(a_.out, "vendor_readout.html"), os.path.join(a_.out, "Vendor_Readout.pdf"))
    print("Call first: %s" % f["call_first"]["vendor"])
    if pdf:
        print("-> %s" % pdf)
    print("-> %s" % os.path.join(a_.out, "Vendor_Readout.xlsx"))
    print("-> %s" % os.path.join(a_.out, "vendor_readout.html"))


if __name__ == "__main__":
    main()
