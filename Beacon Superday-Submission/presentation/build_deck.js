// Builds presentation/Beacon_FDE_Casestudy.pptx
// Every number on these slides comes from the two tools' outputs on the case data
// (output/PO_Confirmation_Check_2026-05-17.xlsx and output/Vendor_Readout.xlsx).
// Run:  npm install pptxgenjs && node build_deck.js
const path = require("path");
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Beacon PO Confirmation Control Tower";

const C = {
  ink: "1F2A33", slate: "4E5D6A", muted: "7D8B97", tint: "EEF2F5", line: "D5DCE2", white: "FFFFFF",
  accent: "D9622B", red: "B83227", amber: "C98A06", blue: "2E6DB4", green: "2E7D4F", paleRed: "FBEAE8",
};
const HEAD = "Cambria", BODY = "Calibri";
const W = 13.333, M = 0.6;
const A = (f) => path.join(__dirname, "assets", f);
let n = 0;

function base(kicker, title, dark = false) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.ink : C.white };
  n += 1;
  if (kicker) s.addText(kicker.toUpperCase(), { x: M, y: 0.35, w: 9, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: C.accent, charSpacing: 2, margin: 0, isTextBox: true });
  if (title) s.addText(title, { x: M, y: 0.65, w: W - 2 * M, h: 0.75, fontFace: HEAD, fontSize: 24, bold: true, color: dark ? C.white : C.ink, margin: 0, valign: "top", isTextBox: true });
  if (n > 1) s.addText(`Beacon Fasteners · FDE case   ${n}`, { x: W - 4.6, y: 7.08, w: 4.0, h: 0.25, fontFace: BODY, fontSize: 9, color: dark ? "9AA7B2" : C.muted, align: "right", margin: 0, isTextBox: true });
  return s;
}

function card(s, x, y, w, h, fill = C.tint) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
}

function stat(s, x, y, w, big, label, color = C.ink, size = 36) {
  s.addText(big, { x, y, w, h: 0.65, fontFace: HEAD, fontSize: size, bold: true, color, margin: 0, isTextBox: true });
  s.addText(label, { x, y: y + 0.66, w: w - 0.2, h: 0.55, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
}

function bullets(s, items, opts) {
  const runs = items.map((it, i) => {
    const o = { bullet: true, breakLine: i < items.length - 1 };
    if (Array.isArray(it)) return [{ text: it[0], options: { ...o, bold: true, breakLine: false } }, { text: it[1], options: { breakLine: i < items.length - 1 } }];
    return [{ text: it, options: o }];
  }).flat();
  s.addText(runs, { fontFace: BODY, fontSize: 14, color: C.ink, paraSpaceAfter: 6, valign: "top", margin: 0.05, isTextBox: true, ...opts });
}

function badge(s, x, y, label, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color }, line: { color } });
  s.addText(label, { x, y, w: 0.42, h: 0.42, fontFace: BODY, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

function table(s, header, rows, opts) {
  const hdr = header.map((h) => ({ text: h, options: { bold: true, color: C.white, fill: { color: C.ink }, fontSize: opts.hsize || 11 } }));
  const body = rows.map((r, i) => r.map((c) => (typeof c === "object" && c !== null && c.text !== undefined ? c : { text: String(c), options: {} }))
    .map((c) => ({ text: c.text, options: { fill: { color: opts.highlight && opts.highlight(i) ? C.paleRed : (i % 2 ? C.white : "F6F8FA") }, ...c.options } })));
  s.addTable([hdr, ...body], { fontFace: BODY, fontSize: opts.size || 11, color: C.ink, border: { type: "solid", pt: 0.5, color: C.line }, valign: "middle", margin: 0.05, autoPage: false, ...opts.t });
}

// ------------------------------------------------------------------ title
{
  const s = base(null, null, true);
  // motif: a short "queue" of rounded bars, the worst item highlighted
  [0, 1, 2, 3, 4].forEach((i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.35, y: 1.35 + i * 0.62, w: 3.35 - i * 0.28, h: 0.42, rectRadius: 0.08,
      fill: { color: i === 0 ? C.accent : "2C3A46" }, line: { color: i === 0 ? C.accent : "2C3A46" } });
  });
  s.addText("TINICUM · FDE SUPER-DAY CASE", { x: M, y: 1.35, w: 8, h: 0.35, fontFace: BODY, fontSize: 13, bold: true, color: C.accent, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("Beacon PO Confirmation\nControl Tower", { x: M, y: 1.85, w: 8.6, h: 1.6, fontFace: HEAD, fontSize: 36, bold: true, color: C.white, margin: 0, valign: "top", isTextBox: true });
  s.addText("A daily buyer worklist, and an evidence-led view of supplier risk", { x: M, y: 3.7, w: 8.4, h: 0.5, fontFace: HEAD, fontSize: 20, italic: true, color: "CFD8DF", margin: 0, isTextBox: true });
  const items = [["For Lisa", "35 vendor PDFs in, one worst-first action queue out, in Excel"],
    ["For the plant manager", "8 months of ERP history: who to call first, and which open orders to watch"],
    ["For Beacon", "Vendor forms and a part-number crosswalk that stop the mess at the source"]];
  items.forEach(([k, v], i) => {
    s.addText(k, { x: M, y: 4.75 + i * 0.5, w: 2.4, h: 0.4, fontFace: BODY, fontSize: 14, bold: true, color: C.accent, margin: 0, isTextBox: true });
    s.addText(v, { x: M + 2.4, y: 4.75 + i * 0.5, w: 9.5, h: 0.4, fontFace: BODY, fontSize: 14, color: C.white, margin: 0, isTextBox: true });
  });
  s.addText("Smrita Pandey", { x: M, y: 6.55, w: 6, h: 0.35, fontFace: BODY, fontSize: 13, color: "9AA7B2", margin: 0, isTextBox: true });
}

// ------------------------------------------------------------------ exec
{
  const s = base("Executive summary", "What Lisa acts on today, and who the plant calls first");
  const colW = (W - 2 * M - 0.4) / 2;
  [[M, "LISA · TASK 1", "PO Confirmation Control Tower"], [M + colW + 0.4, "PLANT MANAGER · TASK 2", "Vendor performance readout"]].forEach(([x, k, t]) => {
    card(s, x, 1.65, colW, 5.2);
    s.addText(k, { x: x + 0.3, y: 1.85, w: colW - 0.6, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.accent, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.3, y: 2.15, w: colW - 0.6, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: C.ink, margin: 0, isTextBox: true });
  });
  const x1 = M + 0.3, x2 = M + colW + 0.7, sw = (colW - 0.6) / 3;
  stat(s, x1, 2.8, sw, "84%", "of 44 lines auto-matched at high confidence", C.ink, 26);
  stat(s, x1 + sw, 2.8, sw, "22", "lines fully clean: Lisa never looks at them", C.green, 26);
  stat(s, x1 + 2 * sw, 2.8, sw, "2", "items to act on today", C.red, 26);
  bullets(s, [
    ["Silently dropped line: ", "Apex left PO-4500050001 line 2 (1,500 pcs) off its acknowledgment."],
    ["Stray PO: ", "a confirmation arrived for PO-4500060619, which is not an open Beacon PO."],
  ], { x: x1, y: 4.25, w: colW - 0.6, h: 2.4, fontSize: 13 });
  stat(s, x2, 2.8, sw, "42%", "Continental on time; other vendors 69-99%", C.red, 26);
  stat(s, x2 + sw, 2.8, sw, "61/63", "of the lines >7 days late are Continental's", C.red, 26);
  stat(s, x2 + 2 * sw, 2.8, sw, "$1.84M", "due in 14 days, not acknowledged", C.amber, 26);
  bullets(s, [
    ["Call Continental first. ", "Late most often, by the most, not improving; part of it is our 19-day plan vs their 28-day quote."],
    ["Call Apex second, commercially. ", "85% of spend; CRES bar 1-2% over PO since Feb; $77k short-shipped balances open."],
  ], { x: x2, y: 4.25, w: colW - 0.6, h: 2.4, fontSize: 13 });
}

// ------------------------------------------------------------------ heard
{
  const s = base("What I heard", "Three decisions behind the case, and what I built for each");
  const cols = [
    ["Lisa, senior buyer", "\"Which PO lines need me today, and what do I do?\"",
     "~90 min a day reading ~35 PDFs against ~50 open POs: clean ERP output, scans, one German, one vendor with its own part numbers.",
     "Misses become receiving shorts, AP fights and customer expedites. The hardest one is a line that is simply not there.",
     "A worst-first action queue in Excel, with memory"],
    ["Plant manager", "\"Which supplier do I call first, and why?\"",
     "8 months of ERP history nobody has analysed; vendor decisions run on anecdote.",
     "The wrong vendor gets the call, and a cause that is partly ours (planning lead times) stays hidden.",
     "A readout with one answer, the evidence and the denominators"],
    ["Beacon", "\"How do we stop this recurring?\"",
     "Every vendor sends its own layout and part numbers; nothing is captured in a structured way.",
     "The same reconciliation is repeated every day, and the part-number mess grows.",
     "Vendor forms plus a governed part-number crosswalk"],
  ];
  const cw = (W - 2 * M - 0.6) / 3;
  cols.forEach(([who, q, today, cost, built], i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.6, cw, 5.25);
    s.addText(who.toUpperCase(), { x: x + 0.25, y: 1.75, w: cw - 0.5, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.accent, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(q, { x: x + 0.25, y: 2.08, w: cw - 0.5, h: 0.85, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    s.addText([{ text: "Today  ", options: { bold: true, color: C.muted } }, { text: today, options: {} }], { x: x + 0.25, y: 3.0, w: cw - 0.5, h: 1.2, fontFace: BODY, fontSize: 12.5, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    s.addText([{ text: "Cost of a miss  ", options: { bold: true, color: C.muted } }, { text: cost, options: {} }], { x: x + 0.25, y: 4.25, w: cw - 0.5, h: 1.15, fontFace: BODY, fontSize: 12.5, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: 5.6, w: cw - 0.4, h: 1.0, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.line } });
    s.addText([{ text: "Built  ", options: { bold: true, color: C.accent } }, { text: built, options: { bold: true } }], { x: x + 0.35, y: 5.6, w: cw - 0.7, h: 1.0, fontFace: BODY, fontSize: 13, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
  });
}

// ------------------------------------------------------------------ 3b approach / time split
{
  const s = base("Approach · how I split the time", "Riskiest first: the data, then Task 1 end to end, then Task 2");
  const steps = [
    ["1", "Profile the data", "Before any code: 6 PDF layouts (4 scans, 1 German, 1 vendor with its own part numbers); an undocumented ERP with traps that change the answer.", "A wrong number would sink both tasks, and it was the cheapest thing to check first."],
    ["2", "Task 1 end to end", "PDF → match → checks → Excel. All 35 PDFs through the pipeline before polishing any one parser.", "Lisa's pain is daily, and a dropped line can only be seen when the whole pipeline works."],
    ["3", "Task 2 on the same data", "Cleaned receipts, then value by month, the vendor table, and who to call first, with definitions.", "The ERP work feeds Task 1 too: FX rates, the part-number crosswalk, vendor history."],
    ["4", "Only then, extras", "Memory across days, the desk app, the vendor forms. Each was added only after both tasks ran.", "\"Useful beats impressive\": nothing extra went in until the basics were done."],
    ["5", "Stop and write", "This deck, the README, and a clean run from the zip, as a grader would do it.", "The presentation counts as much as the code."],
  ];
  const y0 = 1.65, rh = 1.02;
  s.addText("What", { x: M + 0.6, y: y0 - 0.05, w: 2.6, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.muted, margin: 0, isTextBox: true });
  s.addText("What it covered", { x: M + 3.3, y: y0 - 0.05, w: 5.0, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.muted, margin: 0, isTextBox: true });
  s.addText("Why in this order", { x: M + 8.55, y: y0 - 0.05, w: 3.5, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.muted, margin: 0, isTextBox: true });
  steps.forEach(([k, t, d, why], i) => {
    const y = y0 + 0.3 + i * rh;
    card(s, M, y, W - 2 * M, rh - 0.12, i % 2 ? C.white : C.tint);
    badge(s, M + 0.1, y + (rh - 0.12 - 0.42) / 2, k, C.accent);
    s.addText(t, { x: M + 0.6, y, w: 2.6, h: rh - 0.12, fontFace: BODY, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    s.addText(d, { x: M + 3.3, y, w: 5.0, h: rh - 0.12, fontFace: BODY, fontSize: 12, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    s.addText(why, { x: M + 8.55, y, w: W - 2 * M - 8.7, h: rh - 0.12, fontFace: BODY, fontSize: 12, italic: true, color: C.slate, margin: 0, valign: "middle", isTextBox: true });
  });
}


// ------------------------------------------------------------------ 4 product decision
{
  const s = base("Task 1 · Design choices", "An exception list Lisa opens in Excel, not a dashboard to learn");
  const lw = 7.3;
  const rows = [
    ["1", "Excel workbook is the product", "Worst first, in words. Lisa's Done and Notes carry over to tomorrow."],
    ["2", "Optional browser app, same engine", "Drag in PDFs, approve part numbers in one click, see scans beside what was read."],
    ["3", "Memory in one local file", "Re-sent PDFs ignored; corrections close items; the first promise is kept."],
    ["4", "Deterministic first, AI as fallback", "One small parser per layout. Claude reads only what no parser can."],
  ];
  rows.forEach(([k, t, d], i) => {
    const y = 1.75 + i * 1.22;
    badge(s, M, y + 0.05, k, C.accent);
    s.addText(t, { x: M + 0.6, y, w: lw - 0.6, h: 0.38, fontFace: BODY, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(d, { x: M + 0.6, y: y + 0.38, w: lw - 0.6, h: 0.75, fontFace: BODY, fontSize: 13, color: C.slate, margin: 0, valign: "top", isTextBox: true });
  });
  const rx = M + lw + 0.4, rw = W - M - rx;
  card(s, rx, 1.7, rw, 5.1);
  s.addText("Rules it never breaks", { x: rx + 0.3, y: 1.9, w: rw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink, margin: 0, isTextBox: true });
  bullets(s, [
    ["Never guesses silently: ", "uncertain matches go to review."],
    ["\"Possible\" dropped line, ", "not \"dropped\"."],
    ["Never sends anything: ", "emails are drafts."],
    ["Every PDF and PO line ", "accounted for, every run."],
  ], { x: rx + 0.3, y: 2.4, w: rw - 0.6, h: 4.2, fontSize: 14 });
}


// ------------------------------------------------------------------ 5 how it works
{
  const s = base("Task 1 · How it works", "Five deterministic steps from vendor PDF to Excel");
  const steps = [
    ["Read", "Text layer: 31 PDFs\nOCR: 4 Continental scans\nClaude: optional fallback"],
    ["Normalize", "Dates, units, decimals\nGerman labels, \"KW 20-22\"\nEUR→USD at ERP rate"],
    ["Match", "Vendor, PO, then each line via the 5-level ladder below"],
    ["Check", "Missing line · qty\nprice · date vs need\ndate sanity · revisions"],
    ["Report", "Excel workbook + memory (SQLite)\n8 sheets"],
  ];
  const bw = 2.2, gap = (W - 2 * M - 5 * bw) / 4;
  steps.forEach(([t, d], i) => {
    const x = M + i * (bw + gap);
    card(s, x, 1.7, bw, 1.95, i === 4 ? "FBE9DF" : C.tint);
    s.addText(t, { x: x + 0.15, y: 1.8, w: bw - 0.3, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: i === 4 ? C.accent : C.ink, margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.15, y: 2.22, w: bw - 0.3, h: 1.35, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
    if (i < 4) s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: x + bw + gap / 2 - 0.09, y: 2.55, w: 0.18, h: 0.26, rotate: 90, fill: { color: C.muted }, line: { color: C.muted } });
  });
  s.addText("Part-number matching ladder: identifiers first, inference only with review", { x: M, y: 3.95, w: 10, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  const conf = (t, c) => ({ text: t, options: { bold: true, color: c } });
  table(s, ["Level", "Evidence", "Confidence", "What happens"], [
    ["L1", "Beacon part number printed on the vendor's document", conf("HIGH", C.green), "Auto-match"],
    ["L2", "Approved (vendor, vendor PN) entry in the crosswalk", conf("HIGH", C.green), "Auto-match; the crosswalk grows every time Lisa approves"],
    ["L3", "Normalized description equals the PO description", conf("HIGH", C.green), "Auto-match"],
    ["L4", "Only qty + price (+ line number) agree on this PO", conf("MEDIUM", C.amber), "Matched for the checks, but sent to Review Required"],
    ["L5", "Nothing agrees", conf("UNMATCHED", C.red), "Shown as unmatched; never guessed"],
  ], { t: { x: M, y: 4.4, w: W - 2 * M, colW: [0.8, 5.0, 1.6, 4.73], rowH: 0.4 }, size: 12 });
}


// ------------------------------------------------------------------ workbook
{
  const s = base("Task 1 · Lisa's workbook", "Lisa's 7am view: exceptions only, worst first, next step");
  const iw = 8.3, ih = iw * 433 / 1041;
  s.addImage({ path: A("workbook_action_list.png"), x: M, y: 1.65, w: iw, h: ih, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  s.addText("Action List tab, real output on this week's 35 PDFs (first columns shown).", { x: M, y: 1.75 + ih, w: iw, h: 0.3, fontFace: BODY, fontSize: 11, italic: true, color: C.muted, margin: 0, isTextBox: true });
  bullets(s, [
    ["Done? and Notes ", "are read back on the next run: handled items stay handled."],
    ["Draft Emails: ", "one ready-to-edit email per vendor; the tool never sends."],
    ["Read Me tab: ", "refresh steps, rules, tolerances, escalation."],
  ], { x: M, y: 2.25 + ih, w: iw, h: 1.6, fontSize: 13 });
  const rx = M + iw + 0.35, rw = W - M - rx;
  card(s, rx, 1.65, rw, 5.2);
  s.addText("One tab per job", { x: rx + 0.25, y: 1.8, w: rw - 0.5, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  const tabs = [["Buyer", "Action List · Dropped & Unconfirmed · Draft Emails"], ["AP", "Price Variances"], ["Planner", "Date Variances"],
    ["Receiving", "Quantity Variances"], ["Master data", "Unmatched Vendor Lines · PN Crosswalk"],
    ["Audit", "All PO Lines · Documents · Review Required · Run Summary"], ["From the ERP", "Old Open Balances ($128k to chase or close)"]];
  s.addText(tabs.map(([k, v], i) => [{ text: k + "\n", options: { bold: true, color: C.accent, fontSize: 11 } }, { text: v, options: { breakLine: i < tabs.length - 1, fontSize: 12 } }]).flat(),
    { x: rx + 0.25, y: 2.3, w: rw - 0.5, h: 4.4, fontFace: BODY, color: C.ink, paraSpaceAfter: 5, valign: "top", margin: 0, isTextBox: true });
}

// ------------------------------------------------------------------ 6 what it caught
{
  const s = base("Task 1 · Run on this week's 35 PDFs (as of 05/17/2026)", "23 things to do, worst first, and 22 lines Lisa can skip");
  const P = (t, c) => ({ text: t, options: { bold: true, color: c } });
  const T = P("Today", C.red), Wk = P("This week", C.amber), Ck = P("Check", C.blue);
  table(s, ["When", "Vendor", "PO / line", "What happened", "What to do"], [
    [T, "Apex", "4500050001 / 2", "Line missing from acknowledgment (BAR-A286-250 × 1,500)", "Ask if it ships, or receiving will be short"],
    [T, "Apex", "4500060619", "Confirmation for a PO that is not open (probably meant for …0001)", "Tell vendor not to ship against it"],
    [Wk, "Apex", "4500050007 / 1", "Confirmed 1,425 of 1,500 (-75 pcs)", "Balance later, or cancelled?"],
    [Wk, "Apex", "4500050002 / 1", "Price $3.9200 → $4.0102 (+2.3%)", "Push back before the invoice reaches AP"],
    [Wk, "Continental", "4500050022 / 1", "Promise 21 days after we need it", "Expedite, or warn planning"],
    [Wk, "Liberty", "4500050019 / 1", "Acknowledged with no qty or date", "Ask for a firm qty and date"],
    [Wk, "Ostmark (DE)", "4500050027 / 1-2", "Promise is a week range \"KW 20-22\", 16 days late", "Pin down a real date"],
    [Ck, "Heritage, Ostmark", "3 lines", "New vendor PNs APH-441-OS, OST-CAR-A-100", "One-click approve (Review Required)"],
    [Ck, "QuickShip", "4500050032", "Sent an invoice instead of an acknowledgment", "Tell receiving; hold the invoice"],
  ], { t: { x: M, y: 1.65, w: 9.0, colW: [1.0, 1.35, 1.35, 3.1, 2.2], rowH: 0.47 }, size: 11 });
  const rx = M + 9.3, rw = W - M - rx;
  card(s, rx, 1.65, rw, 5.05);
  stat(s, rx + 0.25, 1.8, rw - 0.5, "2 · 8 · 13", "act today · this week · check", C.ink, 28);
  stat(s, rx + 0.25, 2.95, rw - 0.5, "22", "lines fully clean: never shown on the Action List", C.green, 28);
  stat(s, rx + 0.25, 4.1, rw - 0.5, "11", "items on Review Required (incl. OCR'd scans)", C.amber, 28);
  s.addText("Built-in checks, every run:\n✓ every PDF has a status\n✓ every open PO line accounted for\n✓ revision superseded original (QuickShip)", { x: rx + 0.25, y: 5.5, w: rw - 0.5, h: 1.15, fontFace: BODY, fontSize: 11, color: C.slate, margin: 0, valign: "top", isTextBox: true });
}


// ------------------------------------------------------------------ 8b fix at source
{
  const s = base("Task 1 · Past feeds future", "Fix it at the source: Beacon's own forms, pre-filled from the ERP");
  const lw = 5.6;
  const cards = [
    ["PO acknowledgment form", "One Excel file per PO, sent with it", [
      "Our lines locked; rows can't be deleted, so nothing drops silently",
      "Every line: Accept / Accept with changes / Cannot supply",
      "One real delivery date; Excel refuses \"KW 20-22\"",
      "Reason code for any qty or price change",
      "Returned file drops into the same check: HIGH match, no OCR"]],
    ["Supplier information pack", "Once per vendor, pre-filled from 8 months of ERP", [
      "Every part we buy from them, with planned / quoted / actual lead time",
      "They confirm their part numbers → Lisa's approval queue",
      "Standard lead times → gap report for planning",
      "Contacts, turnaround, AS9100 / Nadcap expiry"]],
  ];
  cards.forEach(([t, sub, items], i) => {
    const y = 1.6 + i * 2.8, h = i === 0 ? 2.65 : 2.3;
    card(s, M, y, lw, h);
    s.addText(t, { x: M + 0.25, y: y + 0.12, w: lw - 0.5, h: 0.38, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(sub, { x: M + 0.25, y: y + 0.48, w: lw - 0.5, h: 0.3, fontFace: BODY, fontSize: 12, italic: true, color: C.accent, margin: 0, isTextBox: true });
    bullets(s, items, { x: M + 0.25, y: y + 0.82, w: lw - 0.5, h: h - 0.9, fontSize: 12, paraSpaceAfter: 2 });
  });
  const rx = M + lw + 0.35, rw = W - M - rx;
  const h1 = rw * 400 / 1380, h2 = rw * 430 / 1380;
  s.addText("Ostmark answers on the form: line 2 is now \"Vendor declined\" (act today, with the reason), not a guess", { x: rx, y: 1.6, w: rw, h: 0.5, fontFace: BODY, fontSize: 12, bold: true, color: C.ink, margin: 0, valign: "top", isTextBox: true });
  s.addImage({ path: A("today_after_form_c.png"), x: rx, y: 2.1, w: rw, h: h1, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  const y2 = 2.1 + h1 + 0.3;
  s.addText("Continental returns its pack: MRP plans 19 days, their standard is 28", { x: rx, y: y2, w: rw, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addImage({ path: A("pack_gaps_c.png"), x: rx, y: y2 + 0.35, w: rw, h: h2, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  s.addText("Vendor replies in the demo are synthetic. Plain Excel: no portal, no logins; PDFs keep working for everyone else.", { x: rx, y: 6.55, w: rw, h: 0.4, fontFace: BODY, fontSize: 11, italic: true, color: C.muted, margin: 0, isTextBox: true });
}


// ------------------------------------------------------------------ 9 data traps
{
  const s = base("Task 2 · Method", "Undocumented ERP: profiled first, 8 traps found");
  bullets(s, [
    ["Receipt dates are MM/DD/YYYY text: ", "SQLite date() returns NULL on every row."],
    ["50 reversal pairs ", "(keyed in error, then reversed) plus 1 unreversed duplicate (txn 701546)."],
    ["po_line.qty_received is wrong ", "on 169 lines; receipts are used instead."],
    ["30 superseded confirmations: ", "only the latest is used."],
    ["Confirmation feed stops 2026-04-03: ", "promise metrics use POs to 03/30 only."],
    ["Ostmark PO prices are already USD: ", "converting again adds about $76k."],
    ["qc_hold dates are unreliable ", "(16 before the PO existed): shown as counts only."],
    ["5 decoy tables ", "with no link to vendor performance (GL, MRP, users, …)."],
  ], { x: M, y: 1.65, w: 6.6, h: 5.1, fontSize: 14 });
  s.addChart(pres.charts.BAR, [{ name: "Received value ($M)", labels: ["Corrected (used)", "Raw 'R' receipts", "po_line.qty_received"], values: [28.19, 29.50, 25.30] }], {
    x: M + 6.9, y: 1.6, w: W - 2 * M - 6.9, h: 4.3, barDir: "col", chartColors: [C.ink, C.muted, C.muted],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$0.00\"M\"", dataLabelFontSize: 13, dataLabelColor: C.ink,
    showTitle: true, title: "Received value, Sep 2025 to May 2026, three ways", titleFontSize: 14, titleColor: C.ink, titleFontFace: BODY,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 12, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    valAxisMinVal: 0, valAxisMaxVal: 33,
  });
  s.addText("The naive answers are wrong by +$1.3M and -$2.9M. Every number in the readout says which records it counts (Definitions tab).", { x: M + 6.9, y: 6.0, w: W - 2 * M - 6.9, h: 0.8, fontFace: BODY, fontSize: 13, italic: true, color: C.slate, margin: 0, isTextBox: true });
}


// ------------------------------------------------------------------ 10 monthly value
{
  const s = base("Task 2 · Received value by month", "$28.2M received in 8 months, 85% of it from Apex");
  const months = ["Sep*", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May*"];
  const apex = [0, 3.080, 4.339, 2.943, 4.489, 2.586, 3.154, 2.716, 0.672];
  const total = [0.114, 3.533, 4.786, 3.668, 5.092, 3.126, 3.723, 3.223, 0.921];
  const other = total.map((t, i) => +(t - apex[i]).toFixed(3));
  s.addChart(pres.charts.BAR, [
    { name: "Apex Bar & Tube", labels: months, values: apex },
    { name: "Other 5 vendors", labels: months, values: other },
  ], {
    x: M, y: 1.6, w: 8.6, h: 5.1, barDir: "col", barGrouping: "stacked", chartColors: [C.ink, C.accent],
    showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: C.slate,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 12, valAxisLabelColor: C.muted, valAxisLabelFontSize: 11, valAxisLabelFormatCode: "$0\"M\"",
    valGridLine: { color: "E6EAEE", size: 0.5 }, catGridLine: { style: "none" },
    showTitle: true, title: "Received value by receipt month ($M, at PO price)", titleFontSize: 14, titleColor: C.ink, titleFontFace: BODY,
  });
  const rx = M + 9.0, rw = W - M - rx;
  stat(s, rx, 1.8, rw, "$28.19M", "received value, 09/06/2025 to 05/17/2026");
  stat(s, rx, 3.0, rw, "85%", "Apex Bar & Tube ($23.98M): a concentration question, not only a performance one");
  stat(s, rx, 4.35, rw, "$3.1-5.1M", "range across full months (Oct to Apr); no clear trend");
  s.addText("* Sep and May are partial months. Reversals netted, 1 duplicate excluded.", { x: rx, y: 5.9, w: rw, h: 0.6, fontFace: BODY, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
}


// ------------------------------------------------------------------ 11 vendor comparison
{
  const s = base("Task 2 · Vendor-by-vendor", "Continental stands out on delivery; Apex on dollars");
  const R = (t) => ({ text: t, options: { bold: true, color: C.red } });
  table(s, ["Vendor", "Received", "Share", "On time", "Lines >7d late", "Median days late (when late)", "Promised after need", "Past-due open $", "On time Sep-Dec → Jan-May"], [
    ["Apex Bar & Tube", "$23.98M", "85.1%", "87%", "0", "1", "19%", R("$76.7k"), "87% → 87%"],
    ["Heritage Cold Heading", "$1.71M", "6.1%", "72%", "2", "3", "60%", "$4.1k", "70% → 74%"],
    [{ text: "Continental Quality Heat Treat", options: { bold: true } }, "$1.51M", "5.4%", R("42%"), R("61"), R("17"), R("91%"), "$14.0k", R("45% → 40%")],
    ["Ostmark Werkzeug (DE)", "$0.87M", "3.1%", "69%", "0", "2", "20%", "$31.9k", "71% → 68%"],
    ["Liberty Surface Finishing", "$0.11M", "0.4%", "91%", "0", "1", "20%", "$1.1k", "95% → 87%"],
    ["QuickShip Industrial", "$0.01M", "0.05%", "99%", "0", "1", "16%", "$0.04k", "98% → 100%"],
  ], { t: { x: M, y: 1.65, w: W - 2 * M, colW: [2.75, 1.05, 0.85, 0.9, 1.1, 1.5, 1.3, 1.3, 1.38], rowH: 0.52 }, size: 13, hsize: 11, highlight: (i) => i === 2 });
  s.addText([
    { text: "On time ", options: { bold: true } }, { text: "= the last receipt that completes the line arrives on or before our required date. Denominators: due lines with a receipt (Apex 222, Heritage 205, Continental 134, Ostmark 55, Liberty 189, QuickShip 110). ", options: {} },
    { text: "Trend ", options: { bold: true } }, { text: "= Fisher exact test; no vendor's change is statistically significant (Liberty p=0.13). ", options: {} },
    { text: "Not ranked: ", options: { bold: true } }, { text: "QC holds (dates unreliable). No composite score: the weights would be invented.", options: {} },
  ], { x: M, y: 5.5, w: W - 2 * M, h: 1.2, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
}


// ------------------------------------------------------------------ 12 call continental
{
  const s = base("Task 2 · Who to call first", "Call Continental first: late most often and by the most");
  const sw = 2.35;
  stat(s, M, 1.7, sw, "42%", "of 134 due lines on time (others 69-99%)", C.red, 40);
  stat(s, M + sw, 1.7, sw, "61/63", "lines >7 days late across all vendors are theirs ($660k)", C.red, 40);
  stat(s, M + 2 * sw, 1.7, sw, "91%", "of their confirmations promise after our need date", C.red, 40);
  bullets(s, [
    ["How bad: ", "median 17 days late when late, worst 28. Every other vendor lands 99%+ of lines within 7 days."],
    ["Not a recent slip: ", "45% on time Sep-Dec (n=56) vs 40% Jan-May (n=78)."],
    ["Misses its own dates: ", "meets only 58% of its own promises."],
    ["Material: ", "$1.51M of heat treat; 7 open lines past due today ($14k)."],
  ], { x: M, y: 3.1, w: 3 * sw - 0.2, h: 2.4, fontSize: 14 });
  card(s, M, 5.55, 3 * sw - 0.2, 1.2, "FBE9DF");
  s.addText([{ text: "Beacon's side too: ", options: { bold: true, color: C.accent } }, { text: "we plan 19 days for heat treat, they quote 28 and deliver in 24.5 (medians). Part of this lateness is our planning. Take that to the call.", options: { color: C.ink } }],
    { x: M + 0.2, y: 5.62, w: 3 * sw - 0.6, h: 1.05, fontFace: BODY, fontSize: 13, margin: 0, valign: "middle", isTextBox: true });
  const cx = M + 3 * sw + 0.2;
  const vend = ["Continental", "Ostmark", "Heritage", "Apex", "Liberty", "QuickShip"], ot = [42, 69, 72, 87, 91, 99];
  s.addChart(pres.charts.BAR, [{ name: "On time %", labels: vend, values: ot }], {
    x: cx, y: 1.6, w: W - M - cx, h: 5.2, barDir: "bar", chartColors: [C.red, C.muted, C.muted, C.muted, C.muted, C.muted],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"", dataLabelFontSize: 12, dataLabelColor: C.ink,
    showTitle: true, title: "On-time % by vendor (due lines, Sep 2025 to May 2026)", titleFontSize: 13, titleColor: C.ink, titleFontFace: BODY,
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 12, catAxisOrientation: "maxMin", valAxisHidden: true, valAxisMaxVal: 110, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  });
}


// ------------------------------------------------------------------ ahead
{
  const s = base("Task 2 · Looking ahead", "From rear-view to a watch list for open orders");
  const sw = 2.35;
  stat(s, M, 1.7, sw, "$3.32M", "80 open lines due in the next 60 days", C.ink, 32);
  stat(s, M + sw, 1.7, sw, "$1.84M", "26 lines due in 14 days, no acknowledgment", C.amber, 32);
  stat(s, M, 3.05, sw, "3", "lines at high risk (2 Continental, 1 Heritage)", C.red, 32);
  stat(s, M + sw, 3.05, sw, "59 of 63", "past lines >7 days late the rule would have flagged", C.green, 32);
  card(s, M, 4.45, 2 * sw - 0.2, 2.45);
  s.addText("The rule (checkable by hand)", { x: M + 0.2, y: 4.55, w: 2 * sw - 0.6, h: 0.35, fontFace: HEAD, fontSize: 13, bold: true, color: C.ink, margin: 0, isTextBox: true });
  bullets(s, [
    "Flag if the vendor already promises after our need date, or delivers this part on time <60% (5+ past lines).",
    "Backtested on 8 months, no look-ahead: 33% of flagged lines were late vs 22% overall. A watch list, not a forecast.",
    "Promise dates come from the ERP and from this week's PDFs read by Task 1.",
  ], { x: M + 0.2, y: 4.95, w: 2 * sw - 0.6, h: 1.9, fontSize: 11, paraSpaceAfter: 3 });
  const rx = M + 2 * sw + 0.2, rw = W - M - rx, ih = rw * 914 / 960;
  const h = Math.min(ih, 5.25), w = h * 960 / 914;
  s.addImage({ path: A("readout_lookahead.png"), x: rx + (rw - w) / 2, y: 1.6, w, h, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
}

// ------------------------------------------------------------------ ai
{
  const s = base("AI · where it earns its place", "Claude reads and explains; it never decides or calculates");
  const cw = (W - 2 * M - 0.4) / 2;
  [[M, "TASK 1 · AI READER", "Reads any PDF no template can", [["125 / 125", "qty, price and date fields matched the parsers on all 35 PDFs"], ["44 / 44", "lines identical to the OCR run with no Tesseract installed"]]],
   [M + cw + 0.4, "TASK 2 · VENDOR Q&A", "Grounded, plain-English answers", [["100%", "of quoted figures verified against an independent recomputation"], ["0", "causal claims: it said downtime is not in the data"]]]].forEach(([x, k, t, stats]) => {
    card(s, x, 1.6, cw, 2.55);
    s.addText(k, { x: x + 0.3, y: 1.75, w: cw - 0.6, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, color: C.accent, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.3, y: 2.05, w: cw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
    stats.forEach(([big, lab], i) => stat(s, x + 0.3 + i * (cw - 0.6) / 2, 2.6, (cw - 0.6) / 2, big, lab, C.green, 26));
  });
  const ih = 1.95, iw = ih * 1270 / 255;
  s.addImage({ path: A("ai_chat.png"), x: (W - iw) / 2, y: 4.4, w: iw, h: ih, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.2 } });
  s.addText("Sees only cleaned tables · quotes every figure · totals precomputed · AI-read items go to review · off without a key", { x: M, y: 4.5 + ih, w: W - 2 * M, h: 0.35, fontFace: BODY, fontSize: 12, italic: true, color: C.slate, margin: 0, isTextBox: true });
}

// ------------------------------------------------------------------ 14 part-number cleanup
{
  const s = base("Task 2 · Toward the part-number cleanup", "The crosswalk grows from daily work and seeds the part master");
  const pend = (t) => ({ text: t, options: { bold: true, color: C.amber } });
  table(s, ["Vendor", "Vendor PN", "Beacon PN", "Source", "Status"], [
    ["Heritage", "APH-441-OS", "CHB-9472-4", "This week's PDF (qty+price)", pend("PENDING: ERP had APH-4411")],
    ["Ostmark", "OST-CAR-A-100", "TOOL-INS-CAR", "This week's PDF (qty+price)", pend("PENDING: ERP had WZ-HM-EIN")],
    ["Heritage", "APH-125 / 318 / 441 / 4411", "CHB-3301 / 7715 / 9472-3 / 9472-4", "ERP confirmations × PO lines", "Seen 30-44 times each"],
    ["Heritage", { text: "K-1050", options: { bold: true, color: C.red } }, "CHB-5520", "ERP confirmations × PO lines", "Seen 41 times"],
    ["Ostmark", { text: "K-1050", options: { bold: true, color: C.red } }, "TOOL-PUNCH-STD", "ERP confirmations × PO lines", "Same PN, different part!"],
    ["Ostmark", "WZ-9472-DS / WZ-HM-EIN", "TOOL-DIE-9472 / TOOL-INS-CAR", "ERP confirmations × PO lines", "Seen 17-20 times each"],
  ], { t: { x: M, y: 1.65, w: 8.4, colW: [1.0, 2.0, 2.1, 1.75, 1.55], rowH: 0.55 }, size: 11 });
  const rx = M + 8.8, rw = W - M - rx;
  card(s, rx, 1.65, rw, 5.1);
  s.addText("For the cleanup", { x: rx + 0.3, y: 1.85, w: rw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  bullets(s, [
    ["Key on (vendor, vendor PN): ", "K-1050 is two different parts."],
    ["Vendors renumber quietly: ", "2 caught this week."],
    ["Governed: ", "source, confidence, status, approver and date on every row."],
    ["Proposed, not imposed: ", "a buyer approves; only approved rows match."],
    ["Never overwrite: ", "old numbers retire, history still joins."],
  ], { x: rx + 0.3, y: 2.3, w: rw - 0.6, h: 4.4, fontSize: 12, paraSpaceAfter: 4 });
  s.addText("Mined from the ERP: 8 mappings. Task 1 adds new ones as Lisa approves them (output/pn_crosswalk_updated.csv).", { x: M, y: 5.8, w: 8.4, h: 0.8, fontFace: BODY, fontSize: 12, italic: true, color: C.slate, margin: 0, isTextBox: true });
}


// ------------------------------------------------------------------ 15 not built
{
  const s = base("Scope", "What I deliberately did not build, and why");
  const items = [
    ["ERP integration or write-back", "No API access, and a wrong write is expensive. The CSV export is the contract; integration comes after trust."],
    ["Auto-emailing vendors", "The tool drafts emails; Lisa sends them. Vendor relationships are hers."],
    ["LLM-first extraction", "6 stable layouts parse deterministically and auditably. Claude is the tested fallback, not the default."],
    ["Full master-data (MDM) platform", "The crosswalk file, with source, confidence and review status, is the seed. The platform comes later."],
    ["A 0-100 vendor score", "The weights would be invented. An evidence table with denominators is more honest."],
    ["A chatbot or a big web app", "Lisa needs a morning action queue, not another interface. Excel is where she works; the small app is optional."],
  ];
  const cw = (W - 2 * M - 0.6) / 3, ch = 2.35;
  items.forEach(([t, d], i) => {
    const x = M + (i % 3) * (cw + 0.3), y = 1.7 + Math.floor(i / 3) * (ch + 0.3);
    card(s, x, y, cw, ch);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.28, w: 0.36, h: 0.36, fill: { color: C.white }, line: { color: C.muted, width: 1.5 } });
    s.addText("✕", { x: x + 0.25, y: y + 0.28, w: 0.36, h: 0.36, fontFace: BODY, fontSize: 13, bold: true, color: C.muted, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.75, y: y + 0.22, w: cw - 0.95, h: 0.5, fontFace: BODY, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    s.addText(d, { x: x + 0.25, y: y + 0.9, w: cw - 0.5, h: ch - 1.05, fontFace: BODY, fontSize: 13, color: C.slate, margin: 0, valign: "top", isTextBox: true });
  });
}


// ------------------------------------------------------------------ 16 where it falls over
{
  const s = base("Limits", "Where this falls over, and what catches it when it does");
  table(s, ["Failure mode", "What happens today", "Mitigation / next step"], [
    [{ text: "A new vendor or a changed layout", options: { bold: true } }, "No parser matches; the PDF shows as unreadable", "Claude reads it (tested: 125/125 fields); Beacon's form removes the problem"],
    [{ text: "A poor scan", options: { bold: true } }, "OCR can misread a digit", "Every OCR'd doc goes to Needs your OK, shown next to the image"],
    [{ text: "Two PO lines with the same qty and price", options: { bold: true } }, "An L4 inference could pair the wrong lines", "L4 is never HIGH; always goes to review"],
    [{ text: "A stale open-PO CSV", options: { bold: true } }, "False \"PO not open\" or \"possible dropped line\"", "Pull the export each morning; show its as-of date"],
    [{ text: "Task 2: value at PO price", options: { bold: true } }, "Not invoiced cost; no AP data in the extract", "Add invoice lines, then true price variance"],
    [{ text: "Task 2: no production-impact data", options: { bold: true } }, "Can rank exposure, but can't prove cost", "Link to expedites and downtime; no causal claims until then"],
  ], { t: { x: M, y: 1.7, w: W - 2 * M, colW: [3.4, 4.1, 4.63], rowH: 0.66 }, size: 13, hsize: 13 });
}


// ------------------------------------------------------------------ 17 questions
{
  const s = base("What I wish I'd asked", "Questions for Lisa and the plant manager, each tied to the data");
  const colW = (W - 2 * M - 0.4) / 2;
  const qs = [
    ["LISA", [
      "How often is a missing line really dropped, vs. confirmed later?",
      "Which miss costs most: a short, a price or a date?",
      "Apex CRES bar +1-2% since Feb: a surcharge that never reached the PO?",
      "Which date rules: our required date, the promise, or MRP need?",
      "What would you need to see before you stop opening every PDF?",
    ]],
    ["PLANT MANAGER", [
      "What does a late heat-treat lot cost: downtime, overtime, an expedite?",
      "Who set the 19-day heat-treat lead time? Continental quotes 28.",
      "Is there a second qualified (Nadcap) heat treater?",
      "Is Apex at 85% of spend a deliberate single source?",
      "Which decision should this drive: re-source, renegotiate or re-plan?",
    ]],
  ];
  qs.forEach(([who, list], c) => {
    const x = M + c * (colW + 0.4);
    card(s, x, 1.6, colW, 5.25);
    s.addText(who, { x: x + 0.3, y: 1.75, w: colW - 0.6, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: C.accent, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(list.map((q, i) => ({ text: q, options: { bullet: { type: "number" }, breakLine: i < list.length - 1 } })), {
      x: x + 0.3, y: 2.1, w: colW - 0.6, h: 4.65, fontFace: BODY, fontSize: 16, color: C.ink, paraSpaceAfter: 14, valign: "top", margin: 0.03, isTextBox: true,
    });
  });
}


// ------------------------------------------------------------------ 18 data wish-list + next steps
{
  const s = base("Data I wish I'd had · 30 / 60 / 90", "What would sharpen the answers, and how this goes into use");
  const lw = 6.2;
  s.addText("Data I wish I'd had (absent from the extract)", { x: M, y: 1.6, w: lw, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  bullets(s, [
    ["Invoice / AP lines: ", "actual price paid, and therefore true price variance"],
    ["Production impact: ", "downtime, expedites, late customer shipments"],
    ["The full vendor master: ", "6 vendors in the files vs ~30 active suppliers"],
    ["PO change and cancel history, ", "and the confirmation feed after 04/03"],
    ["Supplier agreements: ", "lead-time commitments, surcharge clauses, strategic status"],
  ], { x: M, y: 2.05, w: lw, h: 4.7, fontSize: 14 });
  const rx = M + lw + 0.4, rw = W - M - rx;
  const steps = [
    ["Days 1-30", "Shadow week with Lisa (target: zero misses, time saved measured); schedule the 6am run; Continental call with the lead-time data."],
    ["Days 31-60", "Supplier packs and acknowledgment forms to Continental, Ostmark, Heritage; planning fixes MRP lead times; clear the $128k backlog."],
    ["Days 61-90", "Approved crosswalk loaded as the ERP vendor-part cross-reference; readout refreshed monthly; decide on a portal or EDI 855."],
    ["Portfolio", "Package the pattern (parsers + rules + workbook) for the next Tinicum company with the same supplier-paperwork problem."],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.6 + i * 1.3;
    card(s, rx, y, rw, 1.15, i === 0 ? "FBE9DF" : C.tint);
    s.addText(t, { x: rx + 0.25, y: y + 0.12, w: 1.3, h: 0.4, fontFace: HEAD, fontSize: 14, bold: true, color: C.accent, margin: 0, isTextBox: true });
    s.addText(d, { x: rx + 1.55, y: y + 0.1, w: rw - 1.75, h: 0.95, fontFace: BODY, fontSize: 13, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
  });
}


// ------------------------------------------------------------------ close
{
  const s = base(null, null, true);
  s.addText("CLOSING", { x: M, y: 1.3, w: 8, h: 0.35, fontFace: BODY, fontSize: 13, bold: true, color: C.accent, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("From a stack of PDFs to a controlled daily workflow", { x: M, y: 1.75, w: 12, h: 0.9, fontFace: HEAD, fontSize: 34, bold: true, color: C.white, margin: 0, isTextBox: true });
  bullets(s, [
    ["Lisa ", "gets a prioritised, auditable exception queue instead of 35 PDFs."],
    ["The plant manager ", "gets a supplier view tied to business impact, not anecdote, and a watch list for what's coming."],
    ["Humans stay in control ", "where commercial judgment or low-confidence extraction matters."],
    ["Next is not a bigger dashboard: ", "one shadow week with Lisa, measure time saved and misses caught, then productionise what earned it."],
    ["Reusable across the portfolio: ", "any company that receives unstructured supplier paperwork has the same shape of problem."],
  ], { x: M, y: 2.95, w: 11.5, h: 3.4, fontSize: 17, color: C.white, paraSpaceAfter: 10 });
  s.addText("Thank you  ·  questions?", { x: M, y: 6.45, w: 8, h: 0.4, fontFace: HEAD, fontSize: 18, italic: true, color: "CFD8DF", margin: 0, isTextBox: true });
}

// ------------------------------------------------------------------ 13 apex + watch
{
  const s = base("Appendix · Second call and watch list", "Call Apex second, about money not delivery; watch three others");
  const lw = 6.2;
  card(s, M, 1.65, lw, 5.1, "FBE9DF");
  s.addText("Apex Bar & Tube: a commercial call", { x: M + 0.3, y: 1.85, w: lw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText("85% of received value ($23.98M). 87% on time, and when late the median is 1 day.", { x: M + 0.3, y: 2.3, w: lw - 0.6, h: 0.6, fontFace: BODY, fontSize: 13, color: C.slate, margin: 0, isTextBox: true });
  bullets(s, [
    ["Price creep: ", "BAR-CRES-250 confirmed +1% from February and +2% from March, while every PO stayed at $3.92. 11 lines, $8.5k over PO, and nothing in the ERP flags it."],
    ["Short-ships left open: ", "8 balances, $77k, the largest past-due exposure of any vendor (oldest 215 days)."],
    ["Same pattern this week: ", "a dropped line, a 75-pc short and a +2.3% price on its acknowledgments."],
  ], { x: M + 0.3, y: 3.0, w: lw - 0.6, h: 3.6, fontSize: 14 });
  const rx = M + lw + 0.4, rw = W - M - rx;
  const watch = [
    ["Heritage Cold Heading", "72% on time; 60% of promises after need date. Delays short (median 3 days). 10 short-ship balances open.", C.amber],
    ["Liberty Surface Finishing", "95% → 87% on time, but not significant (p=0.13) and misses are 1-3 days. Watch, don't call.", C.amber],
    ["Ostmark Werkzeug (DE)", "69% on time, never >6 days late. Open die set PO-4500050027: $31k, due 05/15, nothing received; ack says only \"KW 20-22\".", C.amber],
    ["QuickShip Industrial", "The benchmark: 99% on time, but only $13k of spend.", C.green],
  ];
  watch.forEach(([t, d, col], i) => {
    const y = 1.65 + i * 1.3;
    card(s, rx, y, rw, 1.15);
    s.addShape(pres.shapes.OVAL, { x: rx + 0.22, y: y + 0.2, w: 0.2, h: 0.2, fill: { color: col }, line: { color: col } });
    s.addText(t, { x: rx + 0.55, y: y + 0.1, w: rw - 0.75, h: 0.38, fontFace: BODY, fontSize: 14, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(d, { x: rx + 0.55, y: y + 0.46, w: rw - 0.75, h: 0.65, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
  });
}


// ------------------------------------------------------------------ 8 memory
{
  const s = base("Appendix · The past and future problem", "A desk that remembers: tomorrow's list shows only what's new");
  table(s, ["Situation", "What the tool does"], [
    ["Same PDF emailed twice", "Recognised by its content hash and ignored"],
    ["Lisa marks an item Done", "Stays done tomorrow, with her note"],
    ["Vendor sends a corrected acknowledgment", "The open item closes by itself, with the reason"],
    ["A new problem appears on a line she had closed", "Reopens it and says why"],
    ["Vendor keeps moving the promise date", "The FIRST promise is never overwritten, so the history stays honest"],
    ["Lisa approves a vendor part number once", "Matches automatically from then on and joins the part-number crosswalk"],
  ], { t: { x: M, y: 1.7, w: 7.9, colW: [3.4, 4.5], rowH: 0.62 }, size: 13, hsize: 13 });
  const rx = M + 8.3, rw = W - M - rx;
  card(s, rx, 1.7, rw, 4.95);
  s.addText("Why keep the first promise", { x: rx + 0.3, y: 1.9, w: rw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText("If every revised date overwrites the last one, every PO ends up looking on time and vendor performance can't be measured. The same trap is in the ERP: 30 confirmations were superseded.", { x: rx + 0.3, y: 2.75, w: rw - 0.6, h: 1.4, fontFace: BODY, fontSize: 13, color: C.slate, margin: 0, valign: "top", isTextBox: true });
  s.addText("Tested", { x: rx + 0.3, y: 4.2, w: rw - 0.6, h: 0.35, fontFace: HEAD, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText("70 automated tests. TestDayTwo runs day 1 on the real PDFs, applies Lisa's Excel edits, then feeds a revised vendor PDF on day 2 and checks each row above.", { x: rx + 0.3, y: 4.6, w: rw - 0.6, h: 1.9, fontFace: BODY, fontSize: 13, color: C.slate, margin: 0, valign: "top", isTextBox: true });
}



// ------------------------------------------------------------------ 7 screenshots
{
  const s = base("Appendix · What Lisa sees", "Today's list, plus the things the tool is not sure about");
  const h1 = 5.25, w1 = h1 * 1630 / 1240;
  s.addImage({ path: A("today_c.png"), x: M, y: 1.6, w: w1, h: h1, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  const rx = M + w1 + 0.4, rw = W - M - rx, h2 = rw * 740 / 1630;
  s.addImage({ path: A("review_c.png"), x: rx, y: 1.6, w: rw, h: h2, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  s.addText("Needs your OK: the vendor renamed a part. The tool shows why it thinks so and what our history calls it. One click approves it for good.", { x: rx, y: 1.7 + h2, w: rw, h: 1.0, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
  const h3 = rw * 530 / 1630;
  s.addImage({ path: A("review_ocr_c.png"), x: rx, y: 2.9 + h2, w: rw, h: h3, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
  s.addText("Scans are shown next to what OCR read, so checking takes seconds, not retyping.", { x: rx, y: 3.0 + h2 + h3, w: rw, h: 0.6, fontFace: BODY, fontSize: 12, color: C.slate, margin: 0, valign: "top", isTextBox: true });
}



pres.writeFile({ fileName: path.join(__dirname, "Beacon_FDE_Casestudy.pptx") }).then((f) => console.log("->", f));
