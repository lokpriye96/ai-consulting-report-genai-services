// Builds outputs/TCS_AI_Deflation_Strategy.pptx. Every number comes from data/derived.json
// (built by src/derive.py from Project 1 and Project 2 outputs) or from data/new_sources.csv (N1-N8).
// Run: node src/build_deck.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
// Optional: writes the palette into the PowerPoint theme. Without it the deck still builds; only theme-linked colours fall back to Office defaults.
let applyTheme = async () => console.log("apply_theme.js not found: theme colours not embedded");
try { ({ applyTheme } = require(path.join(process.env.PPTX_SKILL || "", "scripts/apply_theme.js"))); } catch (e) { /* optional */ }

const ROOT = path.resolve(__dirname, "..");
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "data/derived.json"), "utf8"));
const T = D.tcs, P = D.peers, M = D.market;
const mn = v => String(v).replace("-", "−");
const sg = (v, d = 1) => (v > 0 ? "+" : v < 0 ? "−" : "") + Math.abs(v).toFixed(d) + "%";
const OUT = path.join(ROOT, "outputs/TCS_AI_Deflation_Strategy.pptx");

// ---------- theme ----------
const THEME = {
  name: "AI Deflation Case",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "16232A", lt1: "FFFFFF", dk2: "0F4C55", lt2: "EEF2F1",
    accent1: "0F4C55", accent2: "C98A0B", accent3: "B63A2E", accent4: "2F7D5B",
    accent5: "6B7A80", accent6: "9FB8BC", hlink: "0F4C55", folHlink: "6B7A80",
  },
};
const HEX = { ink: "16232A", teal: "0F4C55", amber: "C98A0B", red: "B63A2E", green: "2F7D5B", grey: "6B7A80", light: "EEF2F1", pale: "9FB8BC", line: "D5DDDB", white: "FFFFFF" };
const FONT = "+mn-lt";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.author = "Utkarsh Saxena";
pres.company = "Independent portfolio project";
pres.title = "Getting ahead of AI deflation: a growth agenda for TCS";
pres.subject = "AI-assisted consulting report (Portfolio Project 3)";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const FOOT = "Independent portfolio case on public data · not affiliated with or endorsed by TCS · not investment advice";
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.white },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.28, w: 12.3, h: 0.3, fontSize: 11, bold: true, color: C.accent2, charSpacing: 1, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.58, w: 12.3, h: 0.95, fontFace: THEME.headFontFace, fontSize: 24, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 7.1, w: 10.5, h: 0.25, fontSize: 9, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.1, w: 0.5, h: 0.25, fontSize: 9, color: HEX.grey, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HEX.teal },
  objects: [],
});

const tb = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontSize: 14, color: C.text1, valign: "top" }, o));
const source = (s, text) => tb(s, "Source: " + text, { x: 0.5, y: 6.62, w: 12.3, h: 0.4, fontSize: 10, color: C.accent5, valign: "bottom" });
const content = (sec, kicker, title) => {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
};
const card = (s, x, y, w, h, fill = HEX.light, name) =>
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, objectName: name });
const chartBase = () => ({
  catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.grey, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, dataLabelFontSize: 11, dataLabelColor: HEX.ink,
  valGridLine: { color: "E3E8E7", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  titleFontFace: FONT, titleFontSize: 12, titleColor: HEX.ink,
  catAxisLineShow: true, catAxisLineColor: "BFC9C7", valAxisLineShow: false,
});

// ================= 1. Title =================
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
  tb(s, "PORTFOLIO PROJECT 3 OF 3  ·  AI-ASSISTED CONSULTING REPORT", { x: 0.7, y: 0.8, w: 11.5, h: 0.35, fontSize: 12, bold: true, color: HEX.pale, charSpacing: 1 });
  tb(s, "Getting ahead of AI deflation", { x: 0.7, y: 1.6, w: 11.5, h: 1.1, fontFace: "Cambria", fontSize: 44, bold: true, color: C.background1 });
  tb(s, "A growth agenda for TCS: how an Indian IT services leader should respond when AI shrinks its core business", { x: 0.7, y: 2.75, w: 9.8, h: 1.0, fontSize: 20, color: C.background1 });
  const stats = [[`$${T.ai_runrate_bn.toFixed(1)} bn`, "TCS AI revenue run-rate"], [sg(T.core_growth_pct), "core (non-AI) revenue per year"], [`~${D.ai_needed_flat_pct}%`, "AI growth needed just to stand still"]];
  stats.forEach(([big, lab], i) => {
    const x = 0.7 + i * 3.7;
    tb(s, big, { x, y: 4.25, w: 3.4, h: 0.75, fontFace: "Cambria", fontSize: 36, bold: true, color: i === 0 ? C.background1 : "F2C46B" });
    tb(s, lab, { x, y: 5.0, w: 3.4, h: 0.4, fontSize: 14, color: HEX.pale });
  });
  tb(s, "Utkarsh Saxena  ·  November 2026  ·  Built on Project 1 (AI competitive intelligence tracker) and Project 2 (GenAI services market analytics)", { x: 0.7, y: 6.15, w: 11.8, h: 0.35, fontSize: 12, color: C.background1 });
  tb(s, FOOT, { x: 0.7, y: 6.55, w: 11.8, h: 0.3, fontSize: 10, color: HEX.pale });
  s.addNotes("Independent case built on my two earlier projects. The client framing (TCS leadership) is hypothetical; every number is from Project 1, Project 2 or a cited public source.");
}

// ================= 2. Executive summary =================
{
  const s = content("Introduction", "EXECUTIVE SUMMARY",
    `TCS's $${T.ai_runrate_bn.toFixed(1)} bn AI business is real, but it is not yet growing the company: the priority is to stop the core from shrinking`);
  const scr = [
    ["Situation", `AI is ${T.ai_share_pct}% of TCS revenue ($${T.ai_runrate_bn.toFixed(1)} bn run-rate, up ${T.ai_3q_growth_pct}% in three quarters). Bookings are healthy: book-to-bill ${T.book_to_bill.toFixed(2)}.`],
    ["Complication", `The non-AI core is shrinking ${Math.abs(T.core_growth_pct)}% a year. TCS grew ${T.yoy_usd_pct}% in USD vs ${M.gartner_it_services_growth_pct}% for the IT services market. At today's core decline, AI must grow ~${D.ai_needed_flat_pct}% a year just to keep revenue flat.`],
    ["Resolution", `The core is the bigger lever: 1 pt of core growth adds ~${D.elasticity.total_pts_per_core_pt} pt to total growth; 10 pts of AI growth add only ~${D.elasticity.total_pts_per_10pt_ai.toFixed(1)} pt. Four moves follow.`],
  ];
  scr.forEach(([lab, txt], i) => {
    const y = 1.75 + i * 1.55;
    card(s, 0.5, y, 6.5, 1.38, HEX.light, "scr-" + lab);
    tb(s, lab.toUpperCase(), { x: 0.7, y: y + 0.15, w: 6.1, h: 0.3, fontSize: 12, bold: true, color: i === 1 ? C.accent3 : C.accent1 });
    tb(s, txt, { x: 0.7, y: y + 0.45, w: 6.1, h: 0.9, fontSize: 14 });
  });
  card(s, 7.3, 1.75, 5.5, 4.48, HEX.teal, "recs-panel");
  tb(s, "RECOMMENDATIONS, IN PRIORITY ORDER", { x: 7.55, y: 1.92, w: 5.0, h: 0.3, fontSize: 12, bold: true, color: "F2C46B" });
  const recs = [
    ["Lead the re-pricing", "Move the annuity book to outcome pricing before clients force it"],
    ["Turn alliances into certified capacity", "Fewer platforms; named, certified delivery teams"],
    ["Win share, not just deals", "AI-led operations, GCC takeovers and scale M&A"],
    ["Manage \"net-new AI\"", "Run and report the business on core growth, not the AI line"],
  ];
  recs.forEach(([h, d], i) => {
    const y = 2.35 + i * 0.95;
    s.addShape(pres.shapes.OVAL, { x: 7.55, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: "F2C46B" }, line: { color: "F2C46B" }, objectName: "rec-num-" + (i + 1) });
    tb(s, String(i + 1), { x: 7.55, y: y + 0.05, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: C.text1, align: "center", valign: "middle" });
    tb(s, h, { x: 8.15, y, w: 4.5, h: 0.32, fontSize: 15, bold: true, color: C.background1 });
    tb(s, d, { x: 8.15, y: y + 0.33, w: 4.5, h: 0.5, fontSize: 13, color: "DDE7E6" });
  });
  source(s, "Project 2 (TCS Q2 FY27, calendar Q3 2026; run-rate = quarter × 4); Gartner IT services growth 2026 via Project 2; core-vs-AI arithmetic in src/derive.py.");
  s.addNotes("Pyramid: one governing thought, three SCR lines, four recommendations. The key number is the 32%: AI must grow that fast just to offset the core decline.");
}

// ================= 3. Question and objectives =================
{
  const s = content("Introduction", "THE QUESTION",
    "How should TCS respond to AI deflation over the next 24–36 months?");
  tb(s, "Three framings considered", { x: 0.5, y: 1.62, w: 7.4, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  const hdr = { bold: true, color: HEX.white, fill: { color: HEX.teal }, fontSize: 12, fontFace: FONT };
  const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 12, color: HEX.ink, fontFace: FONT, valign: "middle" }, o) });
  const rows = [
    [{ text: "Client and question", options: hdr }, { text: "Strength of evidence", options: hdr }, { text: "Verdict", options: hdr }],
    [cell("A. TCS: how to respond to AI deflation", { bold: true }), cell("Company-reported AI revenue (grade A), four quarters of history, clearest tension: strong bookings, slow growth"), cell("Chosen", { bold: true, color: HEX.green })],
    [cell("B. Infosys: why AI is not lifting growth"), cell("Fastest core decline (−5.9%/yr), but Infosys uses the broadest AI definition, so the finding is the least robust"), cell("Not chosen", { color: HEX.grey })],
    [cell("C. A GCC / enterprise buyer: how to capture AI savings in vendor contracts"), cell("Strong question, but Projects 1 and 2 hold vendor-side data, not contract or buyer data"), cell("Not chosen", { color: HEX.grey })],
  ];
  s.addTable(rows, { x: 0.5, y: 2.05, w: 7.4, colW: [2.4, 3.9, 1.1], rowH: [0.4, 0.95, 0.95, 0.95], border: { type: "solid", color: HEX.line, pt: 0.75 }, fill: { color: HEX.white }, margin: 0.08, objectName: "options-table" });
  card(s, 8.3, 1.62, 4.5, 4.8, HEX.light, "objectives");
  tb(s, "Objectives", { x: 8.5, y: 1.78, w: 4.1, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  tb(s, [
    { text: "Test whether AI is net-additive for TCS", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Size the gap AI has to close", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Prioritise 3–4 moves, each with KPIs and risks", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Sequence them into a 36-month roadmap", options: { bullet: { type: "number" } } },
  ], { x: 8.5, y: 2.2, w: 4.1, h: 1.75, fontSize: 14, paraSpaceAfter: 4 });
  tb(s, "Scope", { x: 8.5, y: 4.05, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  tb(s, `TCS vs 7 peers from Projects 1 and 2; public data; latest quarter = calendar Q3 2026. Out of scope: the HyperVault data-centre business (a separate capex case) and the share price.`, { x: 8.5, y: 4.38, w: 4.1, h: 1.4, fontSize: 13 });
  tb(s, "Decision owner (hypothetical): CEO and strategy office", { x: 8.5, y: 5.85, w: 4.1, h: 0.45, fontSize: 12, italic: true, color: C.accent5 });
  source(s, "Project 2 core-vs-AI analysis (sql/03_core_vs_ai.sql) and evidence grades; author's framing.");
}

// ================= 4. Situation =================
pres.addSection({ title: "Situation and complication" });
{
  const s = content("Situation and complication", "SITUATION",
    `AI is already a $${T.ai_runrate_bn.toFixed(1)} bn, ${Math.round(T.ai_share_pct)}% business for TCS, and bookings are healthy`);
  s.addChart(pres.charts.BAR, [{ name: "Annualised AI revenue, USD bn", labels: T.ai_series.map(r => r.period), values: T.ai_series.map(r => r.usd_bn) }], Object.assign(chartBase(), {
    x: 0.5, y: 1.65, w: 7.2, h: 4.8, barDir: "col", barGapWidthPct: 55, chartColors: [HEX.teal],
    showTitle: true, title: "TCS annualised AI services revenue, USD bn (company-reported)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 14, dataLabelFontBold: true,
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 3.6, valGridLine: { style: "none" }, catAxisLabelFontSize: 12,
  }));
  const stats = [
    [`${T.ai_share_pct}%`, "of revenue is AI", `$${T.ai_runrate_bn.toFixed(1)} bn of a $${T.revenue_runrate_bn.toFixed(1)} bn run-rate`, HEX.teal],
    [`+${T.ai_3q_growth_pct}%`, "AI growth in three quarters", "Q3 FY26 to Q2 FY27", HEX.teal],
    [T.book_to_bill.toFixed(2), "book-to-bill", `TCV $${T.tcv_q_bn} bn in the quarter; second highest of the eight`, HEX.teal],
    [String(T.p1_events), "verified AI moves (Project 1)", `widest frontier-lab bench: ${T.p1_frontier_labs.split(",").join(", ")}`, HEX.amber],
  ];
  stats.forEach(([big, lab, sub, col], i) => {
    const y = 1.65 + i * 1.2;
    card(s, 8.1, y, 4.7, 1.05, HEX.light, "stat-" + i);
    tb(s, big, { x: 8.3, y: y + 0.12, w: 1.6, h: 0.8, fontFace: "Cambria", fontSize: 28, bold: true, color: col, valign: "middle" });
    tb(s, lab, { x: 9.95, y: y + 0.12, w: 2.75, h: 0.35, fontSize: 14, bold: true });
    tb(s, sub, { x: 9.95, y: y + 0.47, w: 2.75, h: 0.5, fontSize: 11, color: C.accent5 });
  });
  source(s, "TCS quarterly releases via Project 2 (data/raw/ai_disclosures.csv: TCS_Q3FY26_PR, TCS_Q4FY26_DH, TCS_Q1FY27_PR, TCS_Q2FY27_QUARTR); book-to-bill from Project 2 scorecard; Project 1 verified events.");
}

// ================= 5. Complication 1 =================
{
  const s = content("Situation and complication", "COMPLICATION 1",
    "But AI is replacing TCS's core business rather than adding to it, and TCS is losing share");
  const disc = ["Infosys", "TCS", "HCLTech", "Accenture"];
  const coreVals = disc.map(k => P[k].core_growth_pct);
  s.addChart(pres.charts.BAR, [{ name: "Core growth", labels: disc, values: coreVals }], Object.assign(chartBase(), {
    x: 0.5, y: 1.65, w: 6.0, h: 4.25, barDir: "col", barGapWidthPct: 60,
    chartColors: coreVals.map((v, i) => disc[i] === "TCS" ? HEX.red : v < 0 ? "D9928A" : HEX.green),
    invertedColors: coreVals.map((v, i) => disc[i] === "TCS" ? HEX.red : v < 0 ? "D9928A" : HEX.green),
    showTitle: true, title: "Core (non-AI) revenue growth, annualised %",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "+0.0;−0.0", dataLabelFontBold: true,
    valAxisHidden: true, valAxisMinVal: -7, valAxisMaxVal: 6, valGridLine: { style: "none" }, catAxisLabelPos: "low",
  }));
  const aiG = disc.map(k => `${k} ${Math.round(P[k].ai_growth_pct)}%`).join(" · ");
  tb(s, `Over the same windows, AI revenue grew: ${aiG} a year`, { x: 0.5, y: 5.95, w: 6.0, h: 0.55, fontSize: 12, color: C.accent5 });

  const share = Object.fromEntries(fs.readFileSync(path.join(ROOT, "data/inputs/01_market_share.csv"), "utf8").trim().split("\n").slice(1).map(l => { const c = l.split(","); return [c[0], parseFloat(c[6])]; }));
  const order = Object.keys(share).sort((a, b) => share[b] - share[a]);
  const labs = ["IT services market", ...order];
  const vals = [M.gartner_it_services_growth_pct, ...order.map(k => share[k])];
  const cols = labs.map(l => l === "IT services market" ? HEX.amber : l === "TCS" ? HEX.red : HEX.pale);
  s.addChart(pres.charts.BAR, [{ name: "YoY growth USD", labels: labs, values: vals }], Object.assign(chartBase(), {
    x: 6.9, y: 1.65, w: 5.9, h: 4.85, barDir: "bar", barGapWidthPct: 40, chartColors: cols, invertedColors: cols,
    showTitle: true, title: "Revenue growth YoY in USD, % (latest quarter)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", catAxisOrientation: "maxMin",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 15, valGridLine: { style: "none" }, catAxisLabelFontSize: 11,
  }));
  source(s, "Project 2 sql/03_core_vs_ai.sql (Infosys, TCS: Q3 FY26→Q2 FY27; HCLTech: Q2 FY26→Q2 FY27; Accenture FY24→FY25) and sql/01_market_share.sql; market = Gartner IT services spending growth 2026 (GARTNER_IT_FEB26). Capgemini includes the WNS acquisition.");
  s.addNotes(`Core shrinkage is an upper bound: some 'AI growth' may be existing work relabelled. The eight firms together grew ${M.eight_yoy_usd_pct}% vs the market's ${M.gartner_it_services_growth_pct}%.`);
}

// ================= 6. Complication 2 =================
{
  const s = content("Situation and complication", "COMPLICATION 2",
    "Clients are taking the AI dividend through price, and AI announcements are not turning into growth");
  const facts = [
    ["25–30%", "less: what clients now want to pay for the same output", "Reuters (N1)"],
    ["~80%", "of TCS's finance, HR and BPO contracts are outcome-priced, double the late-2023 share", "Reuters (N1)"],
    ["2% to 5%", "IT services deflation, historic vs now (Everest Group estimate)", "Everest Group (N3)"],
    ["< 24 mo", "clients renegotiate contracts within 24 months of signing, not at renewal", "HFS Research (N2)"],
  ];
  facts.forEach(([big, txt, src], i) => {
    const x = 0.5 + (i % 2) * 3.2, y = 1.7 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 3.0, 2.15, HEX.light, "fact-" + i);
    tb(s, big, { x: x + 0.2, y: y + 0.15, w: 2.6, h: 0.65, fontFace: "Cambria", fontSize: 28, bold: true, color: i < 2 ? C.accent3 : C.accent2 });
    tb(s, txt, { x: x + 0.2, y: y + 0.82, w: 2.6, h: 1.0, fontSize: 13 });
    tb(s, src, { x: x + 0.2, y: y + 1.8, w: 2.6, h: 0.28, fontSize: 10, color: C.accent5 });
  });
  const firms = Object.keys(P);
  const series = [{ name: "X", values: firms.map(k => P[k].p1_events) }, { name: "Revenue growth YoY cc %", values: firms.map(k => P[k].yoy_cc_pct), labels: firms.map(k => k === "TCS" ? "TCS, HCLTech" : k === "HCLTech" ? "" : k) }];
  s.addChart(pres.charts.SCATTER, series, Object.assign(chartBase(), {
    x: 7.1, y: 1.65, w: 5.7, h: 4.55, chartColors: [HEX.teal], lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 10,
    showTitle: true, title: `Project 1 AI announcements vs growth (Spearman ρ = ${mn(M.spearman_events_vs_growth)})`,
    showLabel: true, dataLabelFormatScatter: "custom", dataLabelPosition: "r", dataLabelFontSize: 10,
    valAxisMinVal: 0, valAxisMaxVal: 13, valAxisMajorUnit: 2, catAxisMinVal: 6, catAxisMaxVal: 11, catAxisMajorUnit: 1,
    showValAxisTitle: true, valAxisTitle: "Revenue growth YoY, constant currency %", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.grey, valAxisTitleFontFace: FONT,
    showCatAxisTitle: true, catAxisTitle: "Verified AI events per firm, Jun 2025 – Nov 2026", catAxisTitleFontSize: 10, catAxisTitleColor: HEX.grey, catAxisTitleFontFace: FONT,
    catAxisLabelFontSize: 10, catGridLine: { style: "none" },
  }));
  tb(s, "TCS, HCLTech and Wipro announced most and grew slowest", { x: 7.1, y: 6.22, w: 5.7, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  source(s, "N1 Outsource Accelerator reporting Reuters (28 Aug 2026); N2 HFS Research (30 Jun 2026); N3 Let's Data Science reporting Moneycontrol/Everest Group (27 Oct 2026); Project 2 sql/05_momentum_vs_money.sql. Eight firms only: directional, not proof.");
}

// ================= 7. The arithmetic =================
{
  const s = content("Situation and complication", "SO WHAT: THE ARITHMETIC",
    `At today's core decline, TCS's AI business must grow ~${D.ai_needed_flat_pct}% a year just for the company to stand still`);
  tb(s, "TCS total revenue growth over one year, % (derived)", { x: 0.5, y: 1.65, w: 7.4, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  const heat = v => (v < 1 ? "F1D3CF" : v < 5 ? "F6E3B8" : "CFE6DA");
  const H = { bold: true, color: HEX.white, fill: { color: HEX.teal }, fontSize: 13, fontFace: FONT, align: "center", valign: "middle" };
  const rows = [[{ text: "Core growth / yr  ↓    AI growth / yr  →", options: Object.assign({}, H, { align: "left" }) }, ...D.sensitivity.ai_cases_pct.map(a => ({ text: `+${a}%`, options: H }))]];
  D.sensitivity.rows.forEach((r, i) => {
    rows.push([
      { text: (i === 0 ? `${sg(r.core_growth_pct)}  (today)` : sg(r.core_growth_pct)), options: { bold: true, fontSize: 14, color: HEX.ink, fontFace: FONT, fill: { color: HEX.light }, valign: "middle" } },
      ...r.cells.map(v => ({ text: sg(v), options: { fontSize: 18, bold: true, color: HEX.ink, fontFace: FONT, align: "center", valign: "middle", fill: { color: heat(v) } } })),
    ]);
  });
  s.addTable(rows, { x: 0.5, y: 2.1, w: 7.4, colW: [3.2, 1.4, 1.4, 1.4], rowH: [0.5, 0.7, 0.7, 0.7, 0.7], border: { type: "solid", color: HEX.white, pt: 2 }, objectName: "sensitivity" });
  tb(s, `Base: AI $${T.ai_runrate_bn.toFixed(1)} bn and core $${T.core_runrate_bn.toFixed(1)} bn run-rate. For reference, Project 2's Base scenario assumes AI growth halves to 20–35% a year.`, { x: 0.5, y: 5.85, w: 7.4, h: 0.6, fontSize: 12, color: C.accent5 });

  const call = [
    ["~" + D.elasticity.total_pts_per_core_pt + " pt", `of total growth from each 1 pt of core growth`],
    ["~" + D.elasticity.total_pts_per_10pt_ai.toFixed(1) + " pt", "of total growth from each 10 pts of AI growth"],
    [`${D.ai_needed_for_8pct[0].ai_growth_pct}% vs ${D.ai_needed_for_8pct[2].ai_growth_pct}%`, "AI growth needed for 8% total growth with the core at −3.6% vs flat"],
  ];
  call.forEach(([big, txt], i) => {
    const y = 1.65 + i * 1.5;
    card(s, 8.3, y, 4.5, 1.32, i === 2 ? HEX.teal : HEX.light, "call-" + i);
    tb(s, big, { x: 8.5, y: y + 0.12, w: 4.1, h: 0.6, fontFace: "Cambria", fontSize: 26, bold: true, color: i === 2 ? "F2C46B" : C.accent1 });
    tb(s, txt, { x: 8.5, y: y + 0.72, w: 4.1, h: 0.55, fontSize: 13, color: i === 2 ? C.background1 : C.text1 });
  });
  tb(s, "Implication: stopping the core decline is worth more than accelerating AI. TCS management's own goal of high-single-digit growth (N4) is out of reach on AI alone.", { x: 8.3, y: 6.15, w: 4.5, h: 0.6, fontSize: 12, bold: true, color: C.accent3 });
  source(s, "Derived in src/derive.py from Project 2 TCS core-vs-AI split (Q2 FY27 × 4). One-year arithmetic, not a forecast. N4 Business Today reporting JM Financial (8 Sep 2026).");
}

// ================= 8. Resolution overview =================
pres.addSection({ title: "Recommendations" });
const RECS = [
  { n: 1, t: "Lead the re-pricing", what: "Move the application and infrastructure annuity book to outcome pricing on TCS's terms, before clients force renegotiation", impact: "High", time: "6–18 months", why: `Attacks the ${sg(T.core_growth_pct)} core decline directly` },
  { n: 2, t: "Turn alliances into certified capacity", what: "Concentrate on 2–3 platforms; sell named, certified delivery teams rather than partnership badges", impact: "Medium", time: "6–12 months", why: `Announcements do not predict growth (ρ = ${mn(M.spearman_events_vs_growth)})` },
  { n: 3, t: "Win share, not just deals", what: "AI-led operations, GCC takeovers and scale M&A to turn strong bookings into growth", impact: "High", time: "12–36 months", why: `Book-to-bill ${T.book_to_bill.toFixed(2)}, yet growth ${T.yoy_usd_pct}% vs market ${M.gartner_it_services_growth_pct}%` },
  { n: 4, t: "Manage \"net-new AI\"", what: "Run and report the business on core growth, AI bookings and revenue per employee, not the AI revenue line", impact: "Enabler", time: "0–6 months", why: `AI share does not predict growth (ρ = ${mn(M.spearman_ai_vs_growth)})` },
];
{
  const s = content("Recommendations", "RESOLUTION",
    "Four moves, led by re-pricing the core before clients force it");
  RECS.forEach((r, i) => {
    const x = 0.5 + i * 3.1, y = 1.7, w = 2.9;
    card(s, x, y, w, 4.75, i === 0 ? HEX.teal : HEX.light, "rec-card-" + r.n);
    const dark = i === 0;
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fill: { color: dark ? "F2C46B" : HEX.teal }, line: { color: dark ? "F2C46B" : HEX.teal }, objectName: "rec-ov-" + r.n });
    tb(s, String(r.n), { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fontSize: 18, bold: true, align: "center", valign: "middle", color: dark ? C.text1 : C.background1 });
    tb(s, r.t, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: 0.75, fontFace: "Cambria", fontSize: 17, bold: true, color: dark ? C.background1 : C.text1 });
    tb(s, r.what, { x: x + 0.2, y: y + 1.65, w: w - 0.4, h: 1.45, fontSize: 13, color: dark ? C.background1 : C.text1 });
    tb(s, "Why: " + r.why, { x: x + 0.2, y: y + 3.1, w: w - 0.4, h: 0.65, fontSize: 12, italic: true, color: dark ? "DDE7E6" : C.accent5 });
    tb(s, [{ text: "Impact  ", options: { bold: true } }, { text: r.impact, options: { breakLine: true } }, { text: "Time to impact  ", options: { bold: true } }, { text: r.time }],
      { x: x + 0.2, y: y + 3.9, w: w - 0.4, h: 0.65, fontSize: 12, color: dark ? "F2C46B" : C.accent1 });
  });
  source(s, "Author's assessment of impact and timing, based on the evidence on slides 4–7.");
}

// ================= 9 & 10. Recommendation detail =================
const DETAIL = [
  {
    r: RECS[0],
    rationale: [
      `The core (${sg(T.core_growth_pct)}/yr) is the biggest drag on growth (slide 7)`,
      "Clients want 25–30% less and reopen deals within 24 months (N1, N2)",
      "TCS already runs ~80% of finance, HR and BPO contracts on outcomes (N1): extend the model to the larger IT annuity book",
    ],
    kpis: [
      `Core (non-AI) growth: ${sg(T.core_growth_pct)}/yr → ≥ 0% within 24 months`,
      "Share of revenue on outcome or transaction pricing (start disclosing)",
      "Price change achieved on renewals",
      "EBIT margin vs the FY27 exit target of 25%+ (N4)",
    ],
    risks: [
      "TCS carries delivery risk → price off measured baselines; cap downside per contract",
      "Revenue dips before it recovers → start with renewals due in the next 24 months",
    ],
  },
  {
    r: RECS[1],
    rationale: [
      `Project 1 announcement counts correlate ${mn(M.spearman_events_vs_growth)} with growth`,
      `Widest bench (${T.p1_frontier_lab_events} frontier labs, ${T.p1_partnerships} partnerships), yet slowest grower of the four largest firms`,
      `Rivals sell capacity: Cognizant has committed ${(D.deals.cognizant_openai_staff / 1000).toFixed(0)},000+ staff to OpenAI, ${(D.deals.cognizant_gemini_staff / 1000).toFixed(0)},000 to Gemini`,
    ],
    kpis: [
      "Certified practitioners per priority platform (publish the number)",
      "AI bookings and win rate per alliance",
      "Partner-sourced pipeline value",
    ],
    risks: [
      "Model leadership shifts → keep a 2–3 platform core, review it every year",
      "Partner channel conflict → agree co-sell rules up front",
    ],
  },
  {
    r: RECS[2],
    rationale: [
      `Book-to-bill ${T.book_to_bill.toFixed(2)}, but growth ${T.yoy_usd_pct}% vs the market's ${M.gartner_it_services_growth_pct}%: bookings are not becoming share`,
      "GCC takeovers (e.g. Best Buy, Q2 FY27) win deals: four Indian firms used them in 2026 (Project 1)",
      `Fastest grower Capgemini bought scale (WNS, $${(D.deals.capgemini_wns_usd_m / 1000).toFixed(1)} bn); TCS's largest deal: $${(T.largest_acq_usd_m / 1000).toFixed(1)} bn`,
    ],
    kpis: [
      `Share of Gartner IT services: ${T.share_of_gartner_pct}% → rising each year`,
      "Conversion of TCV into revenue growth",
      "Number and value of GCC takeovers",
      "Inorganic revenue contribution",
    ],
    risks: [
      "Integration and overpaying → screen targets against moves 1 and 2",
      "Margin dilution on captive transfers → price on a multi-year productivity glide path",
    ],
  },
  {
    r: RECS[3],
    rationale: [
      `AI share of revenue does not predict growth (ρ = ${mn(M.spearman_ai_vs_growth)}); definitions differ across firms`,
      "TCS's CEO: AI-assisted delivery is not automatically AI revenue (N5); the AI line can rise while the company stalls",
      `Revenue per employee $${T.rev_per_employee_k}k vs peer median $${D.rpe_peer_median_k}k (Accenture $${P.Accenture.rev_per_employee_k}k)`,
    ],
    kpis: [
      "Core growth reported each quarter beside AI revenue",
      "AI share of bookings (leading indicator)",
      `Revenue per employee: $${T.rev_per_employee_k}k → peer median $${D.rpe_peer_median_k}k`,
    ],
    risks: [
      "Disclosure exposes the core decline → publish it with the move-1 plan",
      "Definition gaming → audit a sample of AI-tagged contracts each year",
    ],
  },
];
const detailSlide = (pair, title) => {
  const s = content("Recommendations", "RECOMMENDATIONS IN DETAIL", title);
  const colX = [0.5, 4.75, 8.85], colW = [4.05, 3.9, 3.95];
  const heads = ["Rationale (evidence)", "KPIs to track", "Risks → mitigation"];
  heads.forEach((h, j) => tb(s, h, { x: colX[j], y: 1.6, w: colW[j], h: 0.3, fontSize: 13, bold: true, color: C.accent5 }));
  pair.forEach((d, i) => {
    const y = 1.98 + i * 2.33;
    card(s, 0.5, y, 12.3, 2.18, HEX.light, "detail-" + d.r.n);
    tb(s, `${d.r.n}. ${d.r.t}`, { x: 0.7, y: y + 0.12, w: 11.9, h: 0.35, fontFace: "Cambria", fontSize: 16, bold: true, color: C.accent1 });
    [d.rationale, d.kpis, d.risks].forEach((list, j) => {
      tb(s, list.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < list.length - 1 } })),
        { x: colX[j] + 0.2, y: y + 0.52, w: colW[j] - 0.3, h: 1.6, fontSize: 12, paraSpaceAfter: 3 });
    });
  });
  return s;
};
{
  const s = detailSlide([DETAIL[0], DETAIL[1]], "Moves 1–2: re-price the core on TCS's terms, and sell capacity rather than partnerships");
  source(s, "N1 Reuters via Outsource Accelerator; N2 HFS Research; N4 Business Today/JM Financial; Project 1 events CTS-008, CTS-009; Project 2 sql/05. KPI targets are proposals, not forecasts.");
}
{
  const s = detailSlide([DETAIL[2], DETAIL[3]], "Moves 3–4: convert strong bookings into share, and run the business on net-new AI");
  source(s, "Project 1 events TCS-008, TCS-009, CAP-001 and key findings; Project 2 sql/01, sql/04, sql/06; N5 Business Standard interview (13 Jul 2026). KPI targets are proposals.");
}

// ================= 11. Roadmap =================
pres.addSection({ title: "Roadmap and next steps" });
{
  const s = content("Roadmap and next steps", "IMPLEMENTATION ROADMAP",
    "A 36-month roadmap: measure first, re-price in waves, then scale the share-gain plays");
  const x0 = 3.3, colW = 3.15, phases = ["Phase 1 · 0–6 months", "Phase 2 · 6–18 months", "Phase 3 · 18–36 months"];
  phases.forEach((p, j) => {
    card(s, x0 + j * colW, 1.65, colW - 0.1, 0.45, HEX.teal, "phase-" + j);
    tb(s, p, { x: x0 + j * colW, y: 1.65, w: colW - 0.1, h: 0.45, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  const plan = [
    ["4. Net-new AI", ["Internal dashboard: core growth, AI bookings, revenue per employee", "Disclose core growth and AI bookings quarterly", "Tie leadership incentives to core growth"]],
    ["1. Re-pricing", ["Pilot outcome pricing on renewals due within 24 months", "Roll out across the annuity book in waves", "Outcome pricing is the default for new annuity deals"]],
    ["2. Alliances", ["Choose 2–3 priority platforms; set certification targets", "Launch named, certified delivery teams", "Annual alliance review on bookings, not announcements"]],
    ["3. Share gain", ["Package a GCC-takeover offer; build an M&A screen", "Execute GCC transfers and 1–2 scale deals", "Integrate; track share vs the market"]],
  ];
  plan.forEach(([lab, cells], i) => {
    const y = 2.22 + i * 0.98;
    tb(s, lab, { x: 0.5, y, w: 2.7, h: 0.85, fontFace: "Cambria", fontSize: 15, bold: true, color: C.accent1, valign: "middle" });
    cells.forEach((t, j) => {
      card(s, x0 + j * colW, y, colW - 0.1, 0.85, j === 0 ? "DCE8E6" : HEX.light, `rm-${i}-${j}`);
      tb(s, t, { x: x0 + j * colW + 0.12, y: y + 0.05, w: colW - 0.34, h: 0.75, fontSize: 12, valign: "middle" });
    });
  });
  const gates = ["Gate: core-growth baseline agreed; pilot renewals chosen", "Gate: core growth stops falling; first scale deal decided", "Gate: core growth ≥ 0%; share of the market rising"];
  gates.forEach((g, j) => tb(s, "◆ " + g, { x: x0 + j * colW, y: 6.17, w: colW - 0.1, h: 0.45, fontSize: 11, bold: true, color: C.accent2 }));
  tb(s, "Decision gates", { x: 0.5, y: 6.17, w: 2.7, h: 0.45, fontSize: 13, bold: true, color: C.accent2 });
  source(s, "Author's proposal. Phases start from the decision date; move 4 comes first because it sets the baseline the other moves are measured against.");
}

// ================= 12. Next steps and limitations =================
{
  const s = content("Roadmap and next steps", "NEXT STEPS AND LIMITATIONS",
    "Next steps: validate the numbers, get contract-level pricing data, and test the moves with experts");
  card(s, 0.5, 1.65, 6.0, 4.85, HEX.light, "next-steps");
  tb(s, "Next 4 weeks", { x: 0.7, y: 1.8, w: 5.6, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
  const steps = [
    ["Validate", "Check headline figures against TCS call transcripts; close Project 2's open verification rows"],
    ["Separate deflation from relabelling", "Collect contract-level pricing evidence (renewal discounts, rate cards)"],
    ["Test with experts", "6–8 interviews with CIOs, GCC heads and former deal leads on moves 1 and 3"],
    ["Decide", "Leadership workshop to agree priorities, KPI targets and owners"],
  ];
  steps.forEach(([h, d], i) => {
    const y = 2.3 + i * 1.02;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y, w: 0.42, h: 0.42, fill: { color: HEX.teal }, line: { color: HEX.teal }, objectName: "ns-" + i });
    tb(s, String(i + 1), { x: 0.7, y, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    tb(s, h, { x: 1.3, y: y - 0.02, w: 5.0, h: 0.32, fontSize: 14, bold: true });
    tb(s, d, { x: 1.3, y: y + 0.3, w: 5.0, h: 0.6, fontSize: 13 });
  });
  tb(s, "Limitations (read before quoting a number)", { x: 6.9, y: 1.65, w: 5.9, h: 0.35, fontSize: 16, bold: true, color: C.accent3 });
  const lim = [
    "AI revenue definitions differ by firm; TCS reports \"AI services\" revenue",
    "Core shrinkage is an upper bound: some AI growth may be relabelled work",
    "TCS core growth rests on two points three quarters apart, annualised",
    "Run-rate = quarter × 4, which ignores seasonality",
    "Slide 7 is one-year arithmetic, not a forecast; KPI targets are proposals",
    "Correlations use 8 firms: directional, not proof",
    "Several sources are press reports of company or analyst figures (marked)",
    "HyperVault (up to 1 GW of AI data centres; revenue expected 18–24 months from kick-off, N8) is out of scope",
  ];
  tb(s, lim.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < lim.length - 1 } })), { x: 6.9, y: 2.1, w: 5.9, h: 4.4, fontSize: 13, paraSpaceAfter: 5 });
  source(s, "Project 2 README limitations; N8 People Matters (10 Oct 2025).");
}

// ================= 13. Appendix: sources =================
pres.addSection({ title: "Appendix" });
{
  const s = content("Appendix", "APPENDIX · SOURCES AND METHOD", "Sources and method");
  tb(s, "Method", { x: 0.5, y: 1.55, w: 4.0, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  tb(s, [
    { text: "Projects 1 and 2 outputs copied into data/inputs/", options: { bullet: true, breakLine: true } },
    { text: "src/derive.py computes every deck number into data/derived.json", options: { bullet: true, breakLine: true } },
    { text: "src/build_deck.js reads only derived.json and the new-source list", options: { bullet: true, breakLine: true } },
    { text: "8 new public sources (N1–N8) for pricing and strategy context", options: { bullet: true, breakLine: true } },
    { text: "Basis: calendar Q3 2026; run-rate = quarter × 4; USD", options: { bullet: true, breakLine: true } },
    { text: "Built with Claude as an AI assistant; author set the question, checked figures and owns the conclusions", options: { bullet: true } },
  ], { x: 0.5, y: 1.95, w: 4.0, h: 4.5, fontSize: 12, paraSpaceAfter: 4 });
  const src = [
    ["P1", "AI-Powered CI Tracker: github.com/lokpriye96/ai-ci-tracker (72 events, 70 verified)"],
    ["P2", "GenAI Services Market Analytics: github.com/lokpriye96/genai-services-market-analytics"],
    ["P2", "via P2: TCS Q3 FY26–Q2 FY27 releases; Gartner IT services and AI spending (2026); Kotak deflation view"],
    ["N1", "Outsource Accelerator reporting Reuters, 28 Aug 2026: outcome pricing at TCS; 25–30% client ask"],
    ["N2", "HFS Research, 30 Jun 2026: contracts renegotiated within 24 months"],
    ["N3", "Let's Data Science reporting Moneycontrol, 27 Oct 2026: Everest Group deflation 2% → 5%"],
    ["N4", "Business Today reporting JM Financial, 8 Sep 2026: TCS growth ambition, margin target"],
    ["N5", "Business Standard, 13 Jul 2026: TCS CEO on what counts as AI revenue"],
    ["N6", "Tata Group, 20 Nov 2025: TPG to invest up to Rs 8,820 crore in HyperVault; 'largest AI-led technology services company'"],
    ["N7", "Inc42, 5 Sep 2026: HyperVault campus of up to 1 GW in Hyderabad"],
    ["N8", "People Matters, 10 Oct 2025: AI infrastructure revenue 18–24 months from kick-off"],
  ];
  const rows = src.map(([id, t]) => [{ text: id, options: { bold: true, color: HEX.teal, fontSize: 11, fontFace: FONT } }, { text: t, options: { fontSize: 11, color: HEX.ink, fontFace: FONT } }]);
  s.addTable(rows, { x: 4.9, y: 1.55, w: 7.9, colW: [0.55, 7.35], rowH: 0.4, border: { type: "solid", color: HEX.line, pt: 0.5 }, margin: 0.05, objectName: "sources-table" });
  tb(s, "Full URLs: data/new_sources.csv and data/inputs/p2_sources_used.csv in the GitHub repo", { x: 4.9, y: 6.2, w: 7.9, h: 0.35, fontSize: 11, italic: true, color: C.accent5 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
