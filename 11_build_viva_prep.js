// Builds the viva preparation notes (separate from the dissertation). Run: node 11_build_viva_prep.js && python3 tools_pdf.py outputs/viva/Viva_Preparation.docx Viva_Preparation.pdf
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, ImageRun, TableOfContents, Footer, PageNumber, LevelFormat, PageBreak,
} = require("docx");

const S = JSON.parse(fs.readFileSync("outputs/final/simple_results.json", "utf8"));
const E = JSON.parse(fs.readFileSync("event_study/results.json", "utf8"));
const R = JSON.parse(fs.readFileSync("outputs/final/results.json", "utf8").replace(/\bNaN\b/g, "null"));
const XC = JSON.parse(fs.readFileSync("outputs/final/excel_cells.json", "utf8"));

// ---------------------------------------------------------------- formatting of numbers
const f2 = (x) => Number(x).toFixed(2);
const f1 = (x) => Number(x).toFixed(1);
const f3 = (x) => Number(x).toFixed(3);
const p3 = (x) => (Number(x) < 0.001 ? "< 0.001" : Number(x).toFixed(3));
const pe = (x) => (Number(x) < 0.001 ? "p < 0.001" : `p = ${Number(x).toFixed(3)}`);
const pc = (x) => (100 * Number(x)).toFixed(2);
const pcs = (x) => (Number(x) >= 0 ? "+" : "") + pc(x);
const neg = (x) => f2(-Number(x));

const W1 = S.H1.welch, W2 = S.H2.welch, WC = S.H2_core.welch;
const R1 = S.H1.reg_simple, R2 = S.H2.reg_simple;
const M1 = S.H1.reg_months, M2 = S.H2.reg_months, MC = S.H2_core.reg_months;
const rob = (s) => S.H1.robustness.find((r) => r.spec.startsWith(s));
const NH = "Non-HFC NBFCs (10)", HF = "HFCs (6, exempt)", CL = "Consumer lenders (Bajaj, SBI Cards, Poonawalla)";
const DNH = "Non-HFC NBFCs minus HFCs", DCL = "Consumer lenders minus HFCs";
const EP = (ek, p) => E[ek].NIFTY50.portfolios[p];
const CS = (ek) => E[ek].NIFTY50.cross_section["[0]"];
const PD = (ek) => E[ek].placebo_days[DNH];
const FC = S.firm.cof, FB = S.firm.bank_share;
const sdOf = (ek, p) => EP(ek, p)["[0]"].car / EP(ek, p)["[0]"].t;      // residual SD of the portfolio (one-day window)
const lakh = (x) => (Number(x) / 100000).toFixed(1);
const welchSE = (w) => Math.sqrt(w.sd_post ** 2 / w.n_post + w.sd_pre ** 2 / w.n_pre);

// ---------------------------------------------------------------- document helpers
const FONT = "Arial";
const run = (t) => (typeof t === "string" ? new TextRun(t) : t);
const P = (text, opts = {}) => new Paragraph({ spacing: { after: 120, line: 300 }, alignment: opts.align || AlignmentType.LEFT,
  keepNext: !!opts.keepNext, children: (Array.isArray(text) ? text : [text]).map(run) });
const B = (t) => new TextRun({ text: t, bold: true });
const I = (t) => new TextRun({ text: t, italics: true });
const H1 = (t, br = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: br, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, keepNext: true, children: [new TextRun(t)] });
const bullet = (t) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 70, line: 290 }, children: (Array.isArray(t) ? t : [t]).map(run) });
const numbered = (ref) => (t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 70, line: 290 }, children: (Array.isArray(t) ? t : [t]).map(run) });
const QA = (q, a) => [
  new Paragraph({ keepNext: true, spacing: { before: 160, after: 60, line: 290 }, children: [new TextRun({ text: "Q. " + q, bold: true, color: "1F3864" })] }),
  ...(Array.isArray(a) ? a : [a]).map((x) => (typeof x === "string" ? new Paragraph({ spacing: { after: 80, line: 290 }, indent: { left: 360 }, children: [new TextRun(x)] }) : x)),
];
const ind = (t) => new Paragraph({ spacing: { after: 80, line: 290 }, indent: { left: 360 }, children: (Array.isArray(t) ? t : [t]).map(run) });
const box = (title, lines) => {
  const BD = { style: BorderStyle.SINGLE, size: 6, color: "8EA9DB" };
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({ width: { size: 9026, type: WidthType.DXA }, borders: { top: BD, bottom: BD, left: BD, right: BD },
      shading: { fill: "EEF3FB", type: ShadingType.CLEAR, color: "auto" }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ spacing: { after: 80 }, children: [B(title)] }),
        ...lines.map((l) => new Paragraph({ spacing: { after: 60, line: 280 }, children: (Array.isArray(l) ? l : [l]).map(run) }))] })] })] });
};
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
function table(head, rows, widths) {
  const cell = (t, i, isHead) => new TableCell({ width: { size: widths[i], type: WidthType.DXA }, borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
    shading: isHead ? { fill: "DCE6F2", type: ShadingType.CLEAR, color: "auto" } : undefined, margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.LEFT, children: [new TextRun({ text: String(t), bold: isHead, size: 18 })] })] });
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, true)) }), ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, false)) }))] });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });
const img = (f, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join("outputs/final", f)), transformation: { width: w, height: h } })] });

const C = [];

// ================================================================= title page
C.push(new Paragraph({ spacing: { before: 2600, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VIVA PREPARATION NOTES", bold: true, size: 40 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Capital Requirements and Credit Supply: The Effect of the RBI's November 2023 Risk-Weight Measures on Bank Consumer Credit and Bank Funding of NBFCs in India, 2021 to 2026", size: 26 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Nityam Gupta, MBA (Full-Time) 2025-27, FMS Delhi", size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Guide: Dr. Dezy Kumari", size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200 }, children: [new TextRun({ text: "Personal notes for the viva. Not part of the submission.", italics: true, size: 20 })] }));
C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Contents", bold: true, size: 30 })] }));
C.push(new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }));

// ================================================================= 1. how to use
C.push(H1("1. How to use these notes"));
C.push(P("These notes cover what an examiner is likely to ask about the dissertation, with short answers in plain words. Every number here is taken from the same results as the dissertation, so it matches the tables exactly."));
C.push(bullet([B("Section 2"), " is a two-minute summary to open the viva with. Practise saying it out loud until it takes under two minutes."]));
C.push(bullet([B("Section 3"), " lists the numbers worth knowing by heart. Examiners often ask for a number and then ask what it means."]));
C.push(bullet([B("Sections 4 to 7"), " explain the policy, the data and each method: what it does, why it was used, what the result was and how to defend it."]));
C.push(bullet([B("Section 8"), " has the questions most likely to come up, with suggested answers. Read the answers, then put them in your own words."]));
C.push(bullet([B("Section 9"), " lists the weak points. Examiners respect a student who knows the limits of the work. Admit them calmly and explain what the evidence still shows."]));
C.push(bullet([B("Sections 10 to 12"), " are a formula sheet, a map of where each number sits in the dissertation and the Excel files, and a glossary."]));
C.push(box("Three rules for the viva", [
  "1. Answer the question asked, then stop. Short answers invite good follow-up questions; long ones invite hard ones.",
  "2. Never claim more than the evidence shows. Say \"significant at 10 per cent\" when it is, not \"significant\".",
  "3. If you do not know, say so and say how you would find out. \"I did not test that; I would do it by ...\" is a good answer.",
]));

// ================================================================= 2. two-minute summary
C.push(H1("2. The two-minute summary"));
C.push(P(I("Say this in your own words. The bold phrases are the points to hit.")));
C.push(P([B("The policy. "), "In November 2023 the RBI raised risk weights on two things: unsecured consumer credit, such as personal loans and credit cards, and banks' loans to NBFCs. Housing finance companies were exempt. In February 2025 it withdrew the higher risk weight on bank loans to NBFCs, from April 2025, but kept the consumer-credit one."]));
C.push(P([B("The question. "), "Did these measures actually slow the targeted credit, and how did the market judge them? I test two hypotheses. H1: bank credit to cards and personal loans slowed relative to exempt household credit. H2: bank credit to NBFCs slowed relative to banks' total lending."]));
C.push(P([B("Data. "), "Three sources: monthly RBI data on bank credit by sector; daily share prices of 16 listed NBFCs and HFCs; and quarterly data I compiled from 382 company investor presentations."]));
C.push(P([B("Method. "), "For the credit data I compare the growth of targeted credit with similar credit that was not targeted, before and after the circular, using Welch t-tests and regressions with a post-policy dummy and month dummies. For share prices I use a standard event study with the market model. For the firm data I use t-tests between NBFCs and the exempt HFCs."]));
C.push(P([B("Findings. "), `Bank credit to NBFCs grew about ${f1(-12 * W2.diff)} percentage points a year more slowly than total bank credit after the circular, significant at 5 per cent once seasonality is controlled for (${pe(M2.p)}), and it recovered after the rollback. Targeted consumer credit grew about ${f1(-12 * W1.diff)} points a year more slowly than exempt credit, significant at 10 per cent (${pe(W1.p)}). The event study is sharper: NBFC shares fell ${pc(-EP("E1", NH)["[0]"].car)} per cent relative to the market on the first trading day after the circular and rose ${pc(EP("E2", NH)["[0]"].car)} per cent after the rollback, while HFC shares did not move on either day. Two consumer lenders reported capital-ratio cuts of 3 to 4 percentage points. NBFCs' funding costs and borrowing mix did not change differently from HFCs'.`]));
C.push(P([B("What it means, and the limits. "), "The measures appear to have slowed the targeted credit, and investors saw them as a cost to the lenders covered. But the circular was a package: it also asked lenders to set limits on unsecured lending, banks had plenty of capital, and credit risk was rising at the same time. So I attribute the effect to the package as a whole, not to the capital charge alone."]));

// ================================================================= 3. numbers to know
C.push(H1("3. Numbers to know by heart"));
C.push(table(["Result", "Number", "What it means"], [
  ["H1 Welch t-test (Table 7.1)", `${f2(W1.diff)} pp a month, ${f1(12 * W1.diff)} a year; t = ${f2(W1.t)}; ${pe(W1.p)}`, "Targeted consumer credit grew more slowly than exempt credit after the circular. Significant at 10 per cent."],
  ["H1 with month dummies (Table 7.4)", `${f2(M1.b)}; ${pe(M1.p)}; Newey-West ${pe(M1.p_nw)}`, "Same conclusion once seasonality is held constant."],
  ["H2 Welch t-test", `${f2(W2.diff)} pp a month, ${f1(12 * W2.diff)} a year; ${pe(W2.p)}`, "Bank credit to NBFCs slowed relative to total credit; not significant without seasonal controls."],
  ["H2 with month dummies", `${f2(M2.b)}; ${pe(M2.p)}; Newey-West ${pe(M2.p_nw)}`, "Significant at 5 per cent once seasonality is controlled for. The headline H2 result."],
  ["H2, narrower benchmark", `Welch ${f2(WC.diff)} (${pe(WC.p)}); month dummies ${f2(MC.b)} (${pe(MC.p)})`, "Result holds when NBFCs, cards and personal loans are taken out of the benchmark."],
  ["R squared, with month dummies", `${f2(R1.r2)} rises to ${f2(M1.r2)} (H1) and ${f2(M2.r2)} (H2)`, "Seasonal patterns explain about half the month-to-month noise."],
  ["Placebo (fake date Nov-2022)", `H1 ${pe(S.H1.placebo_welch.p)}; H2 ${pe(S.H2.placebo_welch.p)}`, "No break before the policy."],
  ["Pre-policy trend", `H1 ${f2(S.H1.pretrend.slope_per_year)} a year (${pe(S.H1.pretrend.p)}); H2 ${f2(S.H2.pretrend.slope_per_year)} (${pe(S.H2.pretrend.p)})`, "No significant trend before the policy."],
  ["H2 after the rollback", `${f2(S.H2.rollback_welch.diff)} pp a month; ${pe(S.H2.rollback_welch.p)}`, "Gap moved back by about the size of the original fall; not significant."],
  ["H1, personal loans only", `${f2(rob("Other personal loans only").diff)}; ${pe(rob("Other personal loans only").p_welch)}`, "H1 is driven mainly by personal loans."],
  ["Economic size", `About Rs ${lakh(S.H1.shortfall.amount)} lakh crore (H1) and Rs ${lakh(S.H2.shortfall.amount)} lakh crore (H2) by January 2025`, "Rough upper-end estimate of credit not extended."],
  ["Circular, NBFCs, day 0 (Table 8.2)", `${pc(EP("E1", NH)["[0]"].car)} per cent; ${pe(EP("E1", NH)["[0]"].p)}`, "NBFC shares fell more than the market predicted."],
  ["Circular, HFCs, day 0", `${pc(EP("E1", HF)["[0]"].car)} per cent; ${pe(EP("E1", HF)["[0]"].p)}`, "Exempt HFCs did not move."],
  ["Circular, NBFCs minus HFCs", `${pc(EP("E1", DNH)["[0]"].car)} per cent; ${pe(EP("E1", DNH)["[0]"].p)}`, "The gap between the groups is significant."],
  ["Circular, consumer lenders, 3 days", `${pc(EP("E1", CL)["[0,+2]"].car)} per cent; ${pe(EP("E1", CL)["[0,+2]"].p)}`, "Strongest reaction: lenders hit by the consumer-credit rule."],
  ["Sign test, circular", `${CS("E1").sign_nonhfc.k} of ${CS("E1").sign_nonhfc.n} NBFCs fell; ${pe(CS("E1").sign_nonhfc.p)}`, "Not driven by one or two stocks."],
  ["Placebo days, circular", `${PD("E1").days_as_large} of ${PD("E1").days} ordinary days as large`, "Day 0 was more extreme than every day in the estimation window."],
  ["Rollback, NBFCs, day 0 (Table 8.3)", `${pcs(EP("E2", NH)["[0]"].car)} per cent; ${pe(EP("E2", NH)["[0]"].p)}`, "Mirror image of the circular."],
  ["Rollback, NBFCs minus HFCs", `${pcs(EP("E2", DNH)["[0]"].car)} per cent; ${pe(EP("E2", DNH)["[0]"].p)}; ${PD("E2").days_as_large} of ${PD("E2").days} days as large`, "Significant difference again."],
  ["Firm level (Table 9.2)", `Cost of funds +${f2(FC.all.mean_change_nonhfc)} vs +${f2(FC.all.mean_change_hfc)} (${pe(FC.all.p)}); bank share +${f2(FB.all.mean_change_nonhfc)} vs +${f2(FB.all.mean_change_hfc)} (${pe(FB.all.p)})`, "No difference between NBFCs and HFCs."],
  ["Company disclosures (Table 9.1)", "Bajaj Finance CRAR -290 bps; SBI Cards about -400 bps", "Direct evidence that the risk weights cut capital ratios."],
  ["Sample sizes", `H1 ${W1.n_pre} months before, ${W1.n_post} after; H2 ${W2.n_pre} and ${W2.n_post}; 16 firms; 220-day estimation window; 382 presentations`, ""],
], [2500, 3000, 3526]));

// ================================================================= 4. the policy
C.push(H1("4. The policy in brief"));
C.push(H2("4.1 What the November 2023 circular did"));
C.push(P("RBI circular RBI/2023-24/85, \"Regulatory measures towards consumer credit and bank credit to NBFCs\", dated 16 November 2023. Its main parts:"));
C.push(table(["Measure", "Change", "Exempt"], [
  ["Consumer credit of banks", "Risk weight 100% to 125%, on outstanding as well as new loans", "Housing, education, vehicle and gold-secured loans"],
  ["Retail loans of NBFCs", "Risk weight 100% to 125%", "Same, plus microfinance and SHG loans"],
  ["Credit card receivables", "Banks 125% to 150%; NBFCs 100% to 125%", "None"],
  ["Bank loans to NBFCs", "+25 percentage points where the rating-based weight was below 100%", "HFCs, core investment companies, loans eligible as priority-sector lending"],
  ["Credit standards", "Review exposure limits; set board-approved limits on unsecured consumer credit", "Deadline 29 February 2024"],
], [2300, 3700, 3026]));
C.push(P("The risk weights applied with immediate effect; only the credit-standards paragraph had the February 2024 deadline. That is why the first post-policy month in the tests is December 2023."));
C.push(H2("4.2 The February 2025 rollback"));
C.push(P("On 25 February 2025 the RBI issued two circulars (RBI/2024-25/119 and /120). The extra 25 points on bank loans to NBFCs was withdrawn from 1 April 2025, and banks' microfinance loans in the nature of consumer credit went back to 100 per cent. The consumer-credit and card risk weights were not changed. This partial reversal is what allows the H2 result to be checked in reverse."));
C.push(H2("4.3 What a risk weight means in rupees (Figure 2.1)"));
C.push(P("Banks must hold total capital of at least 9 per cent of risk-weighted assets plus a 2.5 per cent conservation buffer, so 11.5 per cent. NBFCs must keep a CRAR of at least 15 per cent."));
C.push(table(["Loan of Rs 100", "Before", "After", "Why"], [
  ["Bank personal loan", "Rs 11.50", "Rs 14.38", "11.5% of Rs 100, then 11.5% of Rs 125"],
  ["Bank credit card balance", "Rs 14.38", "Rs 17.25", "Risk weight 125% to 150%"],
  ["Bank loan to an NBFC (50% rating weight)", "Rs 5.75", "about Rs 8.6", "Risk weight 50% to 75%; the 50% weight is illustrative"],
  ["NBFC personal loan", "Rs 15.00", "Rs 18.75", "15% CRAR on 100%, then on 125%"],
], [2900, 1300, 1400, 3426]));
C.push(P("If asked why capital matters: equity is the most expensive source of funding, so needing more of it per loan raises the cost of that loan. A lender can respond by charging more, lending less, raising capital or moving to exempt lending."));
C.push(H2("4.4 Other dates worth knowing"));
C.push(bullet("Repo rate: raised by 250 basis points from May 2022 to February 2023, to 6.5 per cent, and held there until February 2025. So the policy came during a pause in rate changes."));
C.push(bullet("Banks' CRAR was 16.8 per cent in September 2023 (RBI Financial Stability Report, December 2023), well above the minimum."));
C.push(bullet("Bajaj Finance: RBI barred its eCOM and Insta EMI Card lending on 15 November 2023, one day before the circular; lifted 2 May 2024."));
C.push(bullet("IIFL Finance: gold-loan business restricted 4 March 2024; lifted 19 September 2024."));
C.push(bullet("HDFC Ltd merged into HDFC Bank on 1 July 2023."));
C.push(bullet("Equity raises: Bajaj Finance's board approved a Rs 10,000 crore raise on 5 October 2023 and the QIP opened on 6 November 2023; Cholamandalam's placement opened on 28 September 2023. Both came before the circular."));
C.push(bullet("2004 precedent: in December 2004 the RBI raised the consumer-credit risk weight from 100 to 125 per cent, the same move it made in 2023."));

// ================================================================= 5. data
C.push(H1("5. The data and the choices behind it", false));
C.push(H2("5.1 RBI bank credit by sector"));
C.push(bullet("Source: RBI Database on Indian Economy (DBIE), Table 15, \"Deployment of Gross Bank Credit by Major Sectors\". Monthly, January 2019 to July 2026, 91 observations, Rs crore outstanding."));
C.push(bullet("Each observation is the last reporting Friday of the month. The November 2023 observation is dated 17 November, one day after the circular, so it is counted as before the policy. A bank's loan book on 17 November cannot reflect a circular issued the previous evening. The first post-policy observation is 29 December 2023."));
C.push(bullet("Growth is the monthly change in the natural log of outstanding credit, times 100. This is close to the percentage change, and log changes add up over time, which makes the economic-size calculation simple."));
C.push(bullet("The HDFC merger (1 July 2023) moved HDFC Ltd's loans onto a bank's books. Bank credit to housing, personal loans and total credit would jump, and bank credit to HFCs and NBFCs would drop. The RBI publishes series that exclude the merger for July 2023 to June 2025, and those are used. The merger month itself (28 July 2023) is dropped for the affected series, which is why H2 has 31 pre-policy months against 32 for H1."));
C.push(bullet("Window for the main tests: April 2021 to January 2025. It ends in January 2025 so the post-policy period finishes before the rollback was announced in February 2025."));
C.push(H2("5.2 Share prices"));
C.push(bullet("Daily closing prices from Investing.com, 1 September 2022 to 28 March 2025, for the 16 stocks and the Nifty 50 and Nifty 500. 641 trading days, so 640 daily returns."));
C.push(bullet("Prices are adjusted for splits and bonuses. This was checked by looking for one-day moves above 15 per cent: there are six, none near the event dates; the largest is IIFL in March 2024, the day of its gold-loan restriction."));
C.push(H2("5.3 Firm data from investor presentations"));
C.push(bullet("16 firms: 10 NBFCs (Arman, Bajaj Finance, Cholamandalam, CreditAccess Grameen, IIFL, M&M Financial, Muthoot, Poonawalla, SBI Cards, Shriram) and 6 HFCs (AAVAS, Can Fin, Home First, LIC Housing, PNB Housing, Repco)."));
C.push(bullet("382 quarterly presentations, Q1 FY2021 to Q4 FY2026. Cost of funds: 366 company-quarters. Bank share of borrowings: 286 company-quarters, 15 firms (Poonawalla does not disclose its borrowing mix). Every value is recorded with the presentation and page it came from."));
C.push(H2("5.4 Why the data were re-collected (Section 4.4)"));
C.push(P("This is likely to come up, and it is a strength. The project started with a ready-made firm dataset. When I checked it against the companies' own filings, several fields turned out not to be reported data: \"unsecured\" assets for five lenders were a fixed share of total assets in every quarter, and bank borrowing shares were wrong for several firms (Arman 71 per cent in the dataset against 32 per cent in its filing; Muthoot 41 against 65). Headline numbers such as total assets were generally right. So every firm-level number in the dissertation was re-collected from the presentations, and the original dataset is not used anywhere."));
C.push(H2("5.5 The Excel files"));
C.push(bullet([B("Dissertation_Data_Compilation.xlsx: "), "all the data, with sources. Raw inputs in blue; growth rates, gaps and returns are Excel formulas."]));
C.push(bullet([B("Dissertation_Tests.xlsx: "), "every test redone in Excel, one sheet per test, with the dissertation's value next to the Excel value. Only the Newey-West standard errors are Python output, because Excel has no function for them."]));
C.push(P("If an examiner wants to see a calculation, open the Tests workbook. Section 11 gives the exact sheet and cell for each headline number."));

// ================================================================= 6. methods: bank credit
C.push(H1("6. The bank credit tests, explained"));
C.push(H2("6.1 The growth gap"));
C.push(P("Each month I take the growth of targeted credit and subtract the growth of comparable credit that the rule did not touch:"));
C.push(bullet([B("Gap1 (H1): "), "average growth of credit cards and other personal loans, minus average growth of housing, vehicle and education loans."]));
C.push(bullet([B("Gap2 (H2): "), "growth of bank credit to NBFCs excluding HFCs, minus growth of total non-food credit."]));
C.push(P("Why a gap: anything that moves both series together, such as interest rates or the economy, cancels out. If the circular slowed targeted credit, the gap should be lower after November 2023 than before. The whole design is just a test of whether the average gap fell."));
C.push(...QA("Why these control groups?", [`For H1, housing, vehicle and education loans are household loans that banks make to similar borrowers, but they were exempt. For H2, total non-food credit is banks' overall lending; bank credit to NBFCs is only about 7 per cent of it, so it barely affects the benchmark. I also use a narrower benchmark that takes NBFCs, cards and personal loans out, and the result is similar (${f2(MC.b)}, ${pe(MC.p)}).`]));
C.push(...QA("Why not gold loans as a control?", ["Gold loans were exempt, but bank gold lending grew at 62 per cent a year after the policy because gold prices rose sharply and banks reclassified agricultural gold loans as retail gold loans. That has nothing to do with the policy and would distort the comparison."]));
C.push(H2("6.2 Welch t-test"));
C.push(P("Compares the average gap after the circular with the average before it. Welch's version does not assume the two periods have the same variance. Here they clearly do not: for H1 the standard deviation is " + f2(W1.sd_pre) + " before and " + f2(W1.sd_post) + " after."));
C.push(box("H1 worked example", [
  `Mean before ${f3(W1.mean_pre)}, mean after ${f3(W1.mean_post)}: difference ${f3(W1.diff)} pp a month.`,
  `Standard error = square root of (${f3(W1.sd_post)}² / ${W1.n_post} + ${f3(W1.sd_pre)}² / ${W1.n_pre}) = ${f3(welchSE(W1))}.`,
  `t = ${f3(W1.diff)} / ${f3(welchSE(W1))} = ${f2(W1.t)}, with Welch degrees of freedom ${f1(W1.df)}. Two-sided ${pe(W1.p)}.`,
  `Per year: ${f3(W1.diff)} x 12 = ${f1(12 * W1.diff)} percentage points.`,
]));
C.push(P(`For H2 the difference is ${f2(W2.diff)} (t = ${f2(W2.t)}, ${pe(W2.p)}). H2 is not significant in the simple t-test because quarter-end lending to NBFCs makes the series very seasonal, which the regression handles.`));
C.push(H2("6.3 Regression with dummy variables"));
C.push(P([I("Gap(t) = a + d x Post(t) + month effects + error")]));
C.push(bullet([B("Post "), "is 1 from December 2023 onward and 0 before. Its coefficient d is the change in the average gap after the circular."]));
C.push(bullet([B("Month dummies: "), "eleven dummies, one for each month from February to December. January is the base month; using twelve would make them add up to the constant (the dummy-variable trap)."]));
C.push(bullet([B("The constant "), `is the expected gap in a January before the policy: ${f2(M1.const)} for H1 and ${f2(M2.const)} for H2. The H2 constant is negative because January is a weak month for bank lending to NBFCs after the December quarter-end peak.`]));
C.push(bullet([B("Why month dummies: "), "the before and after periods contain a different mix of calendar months. The 14 post-policy months run from December 2023 to January 2025, so December and January appear twice. If those months are seasonally strong or weak, a simple comparison is biased. The dummies compare like with like."]));
C.push(table(["", "H1 (Model 1)", "H2 (Model 2)"], [
  ["Coefficient on Post (d)", `${f2(M1.b)} (SE ${f2(M1.se)})`, `${f2(M2.b)} (SE ${f2(M2.se)})`],
  ["t, p-value", `${f2(M1.t)}, ${p3(M1.p)}`, `${f2(M2.t)}, ${p3(M2.p)}`],
  ["R squared / adjusted", `${f2(M1.r2)} / ${f2(M1.adj_r2)}`, `${f2(M2.r2)} / ${f2(M2.adj_r2)}`],
  ["F (overall), p", `${f2(M1.F)}, ${p3(M1.F_p)}`, `${f2(M2.F)}, ${p3(M2.F_p)}`],
  ["Month dummies jointly, p", p3(M1.months_joint_p), p3(M2.months_joint_p)],
  ["Observations, residual df", `${M1.n}, ${M1.df_res}`, `${M2.n}, ${M2.df_res}`],
], [3200, 2900, 2926]));
C.push(...QA("Why is adjusted R squared so much lower than R squared?", [`Adjusted R squared penalises extra regressors. There are 12 regressors and only ${M1.n} observations, so the penalty is large: ${f2(M1.r2)} becomes ${f2(M1.adj_r2)}. The month dummies still add explanatory power overall, which the F-test on them confirms (p = ${p3(M1.months_joint_p)}).`]));
C.push(...QA("What does the ANOVA table test?", [`Whether all the regressors together explain any of the variation in the gap. Null: every coefficient is zero. F = ${f2(M1.F)} with ${M1.df_reg} and ${M1.df_res} degrees of freedom gives p = ${p3(M1.F_p)} for H1, so the model as a whole is significant.`]));
C.push(H2("6.4 Diagnostic tests"));
C.push(table(["Test", "Checks", "H1", "H2", "How to read it"], [
  ["Durbin-Watson", "Correlation between one month's error and the next", f2(M1.dw), f2(M2.dw), "Near 2 means little first-order autocorrelation"],
  ["Jarque-Bera", "Are residuals normal? (uses skewness and kurtosis)", `${f2(M1.jb)} (p ${p3(M1.jb_p)})`, `${f2(M2.jb)} (p ${p3(M2.jb_p)})`, "p above 0.05: normality not rejected"],
  ["Breusch-Pagan", "Is the error variance constant?", `${f2(M1.bp)} (p ${p3(M1.bp_p)})`, `${f2(M2.bp)} (p ${p3(M2.bp_p)})`, "p above 0.05: constant variance not rejected"],
  ["Largest VIF", "Are the regressors too closely related?", f2(M1.max_vif), f2(M2.max_vif), "Below 10 (here below 2): no multicollinearity problem"],
], [1500, 2500, 1350, 1350, 2326]));
C.push(P("A fair caveat to volunteer: with 45 or 46 observations these tests have little power, so \"not rejected\" is weak reassurance. That is one reason the Newey-West check is also reported."));
C.push(H2("6.5 Newey-West standard errors"));
C.push(P(`Newey-West standard errors allow the errors to have unequal variance and to be correlated with the errors of nearby months (three lags here). The coefficient does not change, only its standard error. Here they are smaller than the classical ones: ${f2(M1.se_nw)} against ${f2(M1.se)} for H1, giving ${pe(M1.p_nw)}; ${f2(M2.se_nw)} against ${f2(M2.se)} for H2, giving ${pe(M2.p_nw)}.`));
C.push(...QA("Why are the Newey-West standard errors smaller? Usually they are larger.", [
  "Mainly because the months after the circular are much less volatile than the months before it. The classical standard error assumes one common variance for all months, so it gives the 14 calm post-policy months the larger pre-policy variance. A method that allows unequal variances, like Newey-West or Welch, does not, and the standard error falls. Autocorrelation is small (Durbin-Watson near 2), so it does not push the other way much.",
  "Because Newey-West standard errors can be unreliable in small samples, I treat the classical p-values as the main result and Newey-West as a check. I would not lean on p = 0.004.",
]));
C.push(H2("6.6 Checks"));
C.push(table(["Check", "Why", "Result"], [
  ["Placebo date, November 2022 (pre-policy data only)", "If the gap was already falling, a fake date would also show a drop", `H1 t-test ${f2(S.H1.placebo_welch.diff)} (p ${p3(S.H1.placebo_welch.p)}), regression ${f2(S.H1.placebo_reg_months.b)} (p ${p3(S.H1.placebo_reg_months.p)}); H2 ${f2(S.H2.placebo_welch.diff)} (p ${p3(S.H2.placebo_welch.p)}), ${f2(S.H2.placebo_reg_months.b)} (p ${p3(S.H2.placebo_reg_months.p)}). No break.`],
  ["Pre-policy trend", "A downward drift before the policy would explain part of the fall", `H1 ${f2(S.H1.pretrend.slope_per_year)} a year (p ${p3(S.H1.pretrend.p)}); H2 ${f2(S.H2.pretrend.slope_per_year)} (p ${p3(S.H2.pretrend.p)}). Not significant, but H1 drifts down slightly, so the H1 estimate is an upper bound.`],
  ["Drop one series (H1)", "Is the result driven by one series?", `Personal loans only ${f2(rob("Other personal loans only").diff)} (p ${p3(rob("Other personal loans only").p_welch)}); cards only ${f2(rob("Credit cards only").diff)} (p ${p3(rob("Credit cards only").p_welch)}). Negative whichever control is dropped.`],
  ["Narrower H2 benchmark", "NBFC credit is part of total credit", `Welch ${f2(WC.diff)} (p ${p3(WC.p)}); regression ${f2(MC.b)} (p ${p3(MC.p)})`],
  ["H2 after the rollback", "If H2 is right, the gap should recover", `${f2(S.H2.rollback_welch.diff)} pp a month (p ${p3(S.H2.rollback_welch.p)}), with ${S.H2.rollback_welch.n_pre} and ${S.H2.rollback_welch.n_post} months. Right direction and size; too few months for significance.`],
], [2600, 2800, 3626]));
C.push(box("An extra check you can mention if asked (not in the dissertation)", [
  "If the main window starts in July 2021 instead of April 2021, which drops the April to June 2021 COVID wave, both results get stronger:",
  "H1: Welch -1.00 (p = 0.011), with month dummies -0.98 (p = 0.014).",
  "H2: Welch -1.34 (p = 0.048), with month dummies -1.44 (p = 0.024).",
  "So the April 2021 start makes the reported results more conservative, not less.",
]));
C.push(H2("6.7 Economic size"));
C.push(P(`If the gap fell by ${f3(-W1.diff)} points a month for ${S.H1.shortfall.months} months, targeted credit ended up lower by exp(${f3(-W1.diff)} x ${S.H1.shortfall.months} / 100) - 1 = ${f1(S.H1.shortfall.pct)} per cent relative to the comparison path. On Rs ${lakh(S.H1.shortfall.outstanding)} lakh crore outstanding in January 2025, that is about Rs ${lakh(S.H1.shortfall.amount)} lakh crore. For H2: ${f1(S.H2.shortfall.pct)} per cent of Rs ${lakh(S.H2.shortfall.outstanding)} lakh crore, about Rs ${lakh(S.H2.shortfall.amount)} lakh crore.`));
C.push(P("Call these rough upper-end figures: they assume the pre-policy gap would have continued, and fast-growing credit usually slows on its own."));

// ================================================================= 7. event study and firm level
C.push(H1("7. The event study and the firm tests, explained", false));
C.push(H2("7.1 The idea"));
C.push(P("If investors thought the rules would hurt certain lenders, those lenders' shares should fall when the news came out, compared with what the market as a whole would have predicted. The exempt HFCs are the natural comparison: same industry, same funding markets, but not covered by the bank-to-NBFC rule."));
C.push(H2("7.2 Step by step"));
C.push(numbered("n1")([B("Event day (day 0). "), "Both announcements came after the market closed. The earliest news reports of the circular were at 5:26 pm on 16 November 2023 (market closes at 3:30 pm), so day 0 is Friday 17 November. The rollback was reported from 9:31 pm on 25 February 2025 and the market was shut on 26 February for Mahashivratri, so day 0 is Thursday 27 February."]));
C.push(numbered("n1")([B("Market model. "), "For each stock, regress its daily return on the Nifty 50's daily return over trading days -250 to -31 before the event (220 days). This gives alpha and beta: what the stock normally does, given the market."]));
C.push(numbered("n1")([B("Abnormal return (AR). "), "Actual return minus (alpha + beta x market return) on the event day."]));
C.push(numbered("n1")([B("Cumulative abnormal return (CAR). "), "Sum of ARs over a window: [0], [0,+1], [0,+2] and [-1,+1]."]));
C.push(numbered("n1")([B("Test. "), "t = CAR / (residual SD x square root of the number of days). The residual SD is how big a normal day's abnormal return is, from the estimation window. p comes from the t-distribution with 218 degrees of freedom."]));
C.push(numbered("n1")([B("Portfolios. "), "All firms share the same event day, so their abnormal returns are correlated and cannot be treated as 16 independent tests. I average returns into equal-weighted portfolios (10 NBFCs, 6 HFCs, 3 consumer lenders) and run the same steps on each portfolio. For the comparison I build a difference portfolio: NBFC return minus HFC return each day."]));
C.push(box("Worked example: NBFC portfolio, day 0 of the circular", [
  `CAR [0] = ${pc(EP("E1", NH)["[0]"].car)} per cent. Residual SD of the portfolio = ${pc(sdOf("E1", NH))} per cent.`,
  `t = ${pc(EP("E1", NH)["[0]"].car)} / ${pc(sdOf("E1", NH))} = ${f2(EP("E1", NH)["[0]"].t)}; ${pe(EP("E1", NH)["[0]"].p)}.`,
  "In words: a typical day's surprise for this portfolio is under 1 per cent, and on day 0 it was more than three times that.",
]));
C.push(H2("7.3 Results"));
C.push(table(["Portfolio", "Circular, day 0", "Rollback, day 0"], [
  ["Non-HFC NBFCs (10)", `${pc(EP("E1", NH)["[0]"].car)}% (p ${p3(EP("E1", NH)["[0]"].p)})`, `${pcs(EP("E2", NH)["[0]"].car)}% (p ${p3(EP("E2", NH)["[0]"].p)})`],
  ["HFCs (6, exempt)", `${pc(EP("E1", HF)["[0]"].car)}% (p ${p3(EP("E1", HF)["[0]"].p)})`, `${pcs(EP("E2", HF)["[0]"].car)}% (p ${p3(EP("E2", HF)["[0]"].p)})`],
  ["NBFCs minus HFCs", `${pc(EP("E1", DNH)["[0]"].car)}% (p ${p3(EP("E1", DNH)["[0]"].p)})`, `${pcs(EP("E2", DNH)["[0]"].car)}% (p ${p3(EP("E2", DNH)["[0]"].p)})`],
  ["Consumer lenders (3)", `${pc(EP("E1", CL)["[0]"].car)}% (p ${p3(EP("E1", CL)["[0]"].p)})`, `${pcs(EP("E2", CL)["[0]"].car)}% (p ${p3(EP("E2", CL)["[0]"].p)})`],
  ["Consumer lenders minus HFCs", `${pc(EP("E1", DCL)["[0]"].car)}% (p ${p3(EP("E1", DCL)["[0]"].p)})`, `${pcs(EP("E2", DCL)["[0]"].car)}% (p ${p3(EP("E2", DCL)["[0]"].p)})`],
  ["Sign test (NBFCs moving the predicted way)", `${CS("E1").sign_nonhfc.k} of 10 (p ${p3(CS("E1").sign_nonhfc.p)})`, `${CS("E2").sign_nonhfc.k} of 10 (p ${p3(CS("E2").sign_nonhfc.p)})`],
  ["Ordinary days with a difference as large", `${PD("E1").days_as_large} of ${PD("E1").days}`, `${PD("E2").days_as_large} of ${PD("E2").days} (p about ${f2(PD("E2").empirical_p)})`],
], [3600, 2700, 2726]));
C.push(P("The pattern to stress: NBFCs fell on the tightening and rose on the reversal, while HFCs stayed flat both times. A mirror-image reaction like this is hard to explain by anything other than the measure itself."));
C.push(img("fig7_event_car.png", 560, 213));
C.push(H2("7.4 The extra event-study checks"));
C.push(bullet([B("Sign test. "), "If the news had no effect, each NBFC would be equally likely to rise or fall. Nine or more of ten falling has probability 11/1024, about 0.011, so two-sided p = 0.021."]));
C.push(bullet([B("Placebo days. "), "Compare day 0's NBFC-minus-HFC abnormal return with the same number on each of the 220 ordinary days. This needs no assumption about the t-distribution. After the circular none of the 220 days was as extreme; after the rollback, 7 were."]));
C.push(bullet([B("Nifty 500 as the market. "), `Almost identical results (circular NBFCs ${pc(E.E1.NIFTY500.portfolios[NH]["[0]"].car)}%, rollback ${pcs(E.E2.NIFTY500.portfolios[NH]["[0]"].car)}%; Table C.1).`]));
C.push(bullet([B("Without Bajaj Finance "), `(embargo announced the day before): circular ${pc(EP("E1", "Non-HFC NBFCs excl. Bajaj (9)")["[0]"].car)}% (p ${p3(EP("E1", "Non-HFC NBFCs excl. Bajaj (9)")["[0]"].p)}).`]));
C.push(bullet([B("Without the microfinance lenders "), `(the RBI eased microfinance risk weights on the rollback day): rollback ${pcs(EP("E2", "Non-HFC NBFCs excl. MFIs (8)")["[0]"].car)}% (p ${p3(EP("E2", "Non-HFC NBFCs excl. MFIs (8)")["[0]"].p)}).`]));
C.push(bullet([B("Bank dependence within NBFCs. "), `Regressing each NBFC's day-0 CAR on its bank share of borrowings gives no relationship (p = ${p3(E.E1.bankshare_within_nonhfc.p)} and ${p3(E.E2.bankshare_within_nonhfc.p)}). Investors seem to have reacted to whether a lender was covered, not to its exact funding mix. With 9 firms this test has little power.`]));
C.push(H2("7.5 Firm-level tests (Chapter 9)"));
C.push(P(`For each firm, average cost of funds and bank share before the policy (to Q2 FY2024) and after it (Q3 FY2024 to Q4 FY2025). Then two tests: a Welch t-test comparing the change for the 10 NBFCs with the change for the 6 HFCs, and a paired t-test of before against after within each group.`));
C.push(table(["", "NBFCs", "HFCs", "Between groups (Welch)", "Within (paired)"], [
  ["Cost of funds, change (pp)", `+${f2(FC.all.mean_change_nonhfc)}`, `+${f2(FC.all.mean_change_hfc)}`, `p ${p3(FC.all.p)}`, `NBFC ${p3(FC["paired_Non-HFC NBFC"].p)}; HFC ${p3(FC.paired_HFC.p)}`],
  ["Bank share, change (pp)", `+${f2(FB.all.mean_change_nonhfc)}`, `+${f2(FB.all.mean_change_hfc)}`, `p ${p3(FB.all.p)}`, `NBFC ${p3(FB["paired_Non-HFC NBFC"].p)}; HFC ${p3(FB.paired_HFC.p)}`],
], [2500, 1200, 1200, 2000, 2126]));
C.push(P("How to explain it: funding costs rose for everyone because of the rate cycle (bank loans reprice with a lag), so there is no NBFC-specific effect. The NBFCs' bank share rose less than the HFCs', which fits banks lending relatively less to NBFCs, but the difference is not significant. With 16 firms that differ a lot from each other, only a large effect would show up."));

// ================================================================= 8. questions
C.push(H1("8. Likely questions and suggested answers"));
C.push(H2("8.1 Motivation and contribution"));
C.push(...QA("Why this topic?", ["It was one of the RBI's biggest interventions in retail credit since the IL&FS crisis, it targeted specific types of credit on a known date, and it was partly reversed. That makes it a clean before-and-after setting. It also matters to NBFC managers and investors, because NBFCs depend heavily on bank funding."]));
C.push(...QA("What is new in your work?", ["Three things. A formal comparison of targeted and exempt credit in RBI data, using the merger-adjusted series, with placebo and trend checks. An event study of both the tightening and the rollback, which gives a second test of each channel. And a firm-level dataset built from 382 investor presentations, with every number traced to its page. Existing discussion of these measures is mostly descriptive."]));
C.push(...QA("Your title says 2021 to 2026, but the tests end in January 2025. Why?", ["The data run to July 2026 and are used for the rollback check and the descriptive tables. The main tests stop in January 2025 so that the post-policy period ends before the rollback was announced in February 2025. Otherwise the rollback would contaminate the post-policy average."]));
C.push(...QA("Why only two hypotheses when the synopsis had five?", ["My guide advised me to focus on one or two well-defined hypotheses. H1 and H2 correspond to the two risk-weight measures in the circular. The other three (capital, a shift to secured lending, recovery after the rollback) are discussed with the evidence available in Section 10.3, but not tested formally."]));
C.push(H2("8.2 Method choices"));
C.push(...QA("Is this a difference-in-differences design?", ["It uses the same basic idea, comparing the change in targeted credit with the change in exempt credit. I implement it in the simplest way: I compute the gap between the two each month and test whether its average changed, with a t-test and a regression with a dummy. There is no panel regression with fixed effects. With a handful of aggregate series and one event date, a simple, transparent test is easier to check and to explain, and the event study gives an independent second test."]));
C.push(...QA("Why Welch's t-test rather than the ordinary t-test?", [`The ordinary t-test assumes equal variance before and after. The gap was much more volatile before the policy (SD ${f2(W1.sd_pre)} against ${f2(W1.sd_post)} for H1), so that assumption fails. Welch's test does not need it.`]));
C.push(...QA("Why log growth?", ["The change in the log is very close to the percentage change for small changes, it treats rises and falls symmetrically, and monthly log changes add up, which makes it easy to turn a monthly effect into a 14-month effect."]));
C.push(...QA("Why do you report two-sided p-values when your hypotheses are one-sided?", [`To be conservative. The hypotheses predict a fall, so a one-sided test would be justified and would halve the p-values: H1's Welch p would be ${f3(W1.p / 2)}, and H2 with month dummies ${f3(M2.p / 2)}. I preferred not to rely on that.`]));
C.push(...QA("Why is 10 per cent significance acceptable?", ["I treat 5 per cent as the standard and report 10 per cent as weak evidence, as set out in Table 5.3. With only 45 or 46 monthly observations, tests have limited power, so a 10 per cent result in the predicted direction, backed by other evidence, is informative but not conclusive. I say exactly that in the dissertation."]));
C.push(...QA("Why the market model and not CAPM or a multi-factor model?", ["Brown and Warner (1985) show that with daily data the simple market model works as well as more complex models for detecting abnormal returns. For one-day windows the choice of model matters very little, because the expected return over a single day is tiny compared with the abnormal return."]));
C.push(...QA("Why the Nifty 50 and not a financial-sector index?", ["A financial-sector index includes several of my own sample firms, such as Bajaj Finance and Shriram Finance, and it is dominated by banks, which were themselves affected by the circular. Using it would absorb part of the effect I am trying to measure. The broad market is the cleaner benchmark, and the Nifty 500 gives the same results."]));
C.push(...QA("Why an estimation window of -250 to -31 days?", ["About a year of data gives stable estimates of alpha and beta, and stopping 30 trading days (about six weeks) before the event keeps any early leak of the news out of the estimate of normal behaviour. This is the standard set-up in MacKinlay (1997)."]));
C.push(...QA("Why equal-weighted portfolios?", ["So that each firm counts the same. In a value-weighted portfolio Bajaj Finance would dominate, and Bajaj had its own RBI embargo the day before the circular. I also show the result without Bajaj, which is almost the same."]));
C.push(...QA("Why use portfolios instead of testing each stock?", ["All the firms react on the same day, so their abnormal returns are correlated. Testing 16 stocks as if they were independent would overstate the significance. Testing a portfolio, with its own volatility measured over the estimation window, handles that correlation."]));
C.push(H2("8.3 Data"));
C.push(...QA("Why is the 17 November 2023 observation treated as pre-policy?", ["RBI data refer to the last reporting Friday of the month. The circular came out on the evening of 16 November, so loans outstanding on 17 November could not have responded. The first month that can reflect the policy is the 29 December observation."]));
C.push(...QA("How did you handle the HDFC merger?", ["The merger moved HDFC Ltd's loans onto HDFC Bank's books in July 2023, which inflates bank credit to housing, personal loans and total credit and deflates bank credit to HFCs and NBFCs. The RBI publishes series excluding the merger for July 2023 to June 2025; I use those for the growth rates, and I drop the merger month itself."]));
C.push(...QA("Why Investing.com? Is it reliable?", ["It provides NSE closing prices adjusted for splits and bonuses, and I checked the adjustment by looking for one-day moves above 15 per cent. There are only six, none near the event dates, and the largest is a genuine event (IIFL's gold-loan restriction)."]));
C.push(...QA("How reliable is the firm data, given that definitions differ?", ["Each firm defines cost of funds and bank borrowing slightly differently, so I never compare levels across firms. I compare each firm's own change over time, which removes the definitional differences as long as a firm keeps its definition. Appendix B lists the definitions."]));
C.push(...QA("Your declaration mentions AI tools. What were they used for?", ["Keep this short and factual. For example: \"They helped me pull figures out of the 382 investor presentations and helped with the Python code. Every figure is recorded with the presentation and page it came from, so each one can be checked, and every test is repeated in Excel in the second workbook. I am happy to walk you through any of the calculations.\" Then stop."]));
C.push(H2("8.4 Results and interpretation"));
C.push(...QA("H1 has p = 0.057. You failed to reject the null at 5 per cent. How can you say H1 is supported?", [`I say it is supported at the 10 per cent level, and that the credit-data evidence for H1 is weaker than for H2. Three things support it beyond the p-value: the result is driven by personal loans, which on their own are strongly significant (${pe(rob("Other personal loans only").p_welch)}); the consumer lenders' shares fell ${pc(-EP("E1", CL)["[0,+2]"].car)} per cent over three days after the circular; and two of them reported capital-ratio cuts of 3 to 4 points.`]));
C.push(...QA("H2 is not significant in the t-test (p = 0.108) but is significant in the regression (p = 0.047). Isn't that cherry-picking?", ["No. Bank lending to NBFCs is very seasonal: it spikes at quarter-ends and in the festive season. The post-policy period contains a different mix of months from the pre-policy period. The regression with month dummies compares like with like, so it is the better test for this series. Both tests were planned from the start and both are reported."]));
C.push(...QA("Is the effect causal?", ["I would say the evidence is consistent with a causal effect but does not prove it. The comparison groups remove common shocks like interest rates; the placebo and trend tests show no fall before the policy; and the share prices moved on the day of the news, in opposite directions on the tightening and the reversal. But other things changed at the same time, mainly rising delinquencies on unsecured loans, and the circular also asked for exposure limits. So I attribute the effect to the package, not to the risk weight alone."]));
C.push(...QA("Could the slowdown be demand rather than supply?", ["Partly, yes. Lower bank credit to NBFCs fits both banks lending less and NBFCs borrowing less, since NBFCs' own lending also slowed in 2024. The event study helps here: share prices fell on the announcement day, before any change in demand could happen, which suggests investors expected a cost to NBFCs. NBFC bond and commercial-paper issuance data would help separate the two; that is left for future work."]));
C.push(...QA("Fast-growing credit slows down anyway. How do you know it was the policy?", ["That is a fair concern; the RBI acted because these segments were growing fast. The trend test finds no significant slowdown before November 2023, and the placebo date finds no break. But the H1 gap did drift down slightly, so I treat the H1 estimate as an upper bound and call the economic-size numbers rough upper-end figures."]));
C.push(...QA("Banks had a CRAR of 16.8 per cent. Why would a higher risk weight bind at all?", ["For most banks it probably did not make capital requirements bind. That is why I say part of the effect may come from the exposure limits in the same circular and from the signal of the RBI's concern. Banks also manage to internal targets above the minimum, so even with headroom, extra capital per loan has a cost. Testing whether banks with less spare capital cut more would need bank-level data."]));
C.push(...QA("Why is the H2 effect diluted?", ["The RBI's series for bank credit to NBFCs includes public-sector and infrastructure lenders with AAA ratings or government ownership, and microfinance lenders whose bank loans are often priority-sector. For many of them the increase was small or did not apply. So the aggregate estimate understates the effect on NBFCs that were fully covered."]));
C.push(...QA("Why did CreditAccess Grameen rise 12 per cent on the rollback day?", ["On the same day the RBI also eased risk weights on banks' microfinance loans, which matters directly for a microfinance lender. That is why I also show the NBFC portfolio without the two microfinance lenders; it still rose 3.06 per cent (p = 0.002)."]));
C.push(...QA("The HFCs drifted down after the circular. Doesn't that weaken the comparison?", ["The clear difference between the groups is in the immediate reaction, on day 0 and day +1. Over longer windows the HFCs also drifted down and the difference is no longer significant. I report that openly. The day-0 reaction is the cleanest measure, because it is the day the news became public."]));
C.push(...QA("Was the circular anticipated?", ["The RBI had expressed concern about unsecured lending before the circular, so investors may have partly expected some action. If so, the day-0 reaction understates the full effect, which would make my estimates conservative."]));
C.push(...QA("Why don't the firm-level tests show anything?", ["Three reasons. Interest rates were rising for everyone, so funding costs went up for NBFCs and HFCs alike. Firms differ a lot from each other, so with 16 firms only a large group difference would be significant. And the firm data are quarterly with company-specific definitions. The firm data do show the capital impact directly, through the Bajaj Finance and SBI Cards disclosures."]));
C.push(...QA("What does the Bajaj Finance segment data show?", ["Its consumer segments fell from 37.1 to 35.7 per cent of assets and then recovered to 37.1 per cent by March 2025. But the dip matches the RBI's separate embargo on two of its products (15 November 2023 to 2 May 2024), so it cannot be attributed to the risk weight. For the largest consumer NBFC, the risk weight did not cause a lasting shift away from consumer lending."]));
C.push(H2("8.5 Implications and next steps"));
C.push(...QA("What should a policymaker take from this?", ["Targeted risk weights can slow credit to a specific segment without a general tightening, and the bank-to-NBFC lever worked quickly in both directions. But because the circular combined risk weights with exposure limits when banks had spare capital, regulators should see such packages as signals as well as capital measures."]));
C.push(...QA("What should an NBFC manager take from this?", ["Dependence on bank funding is a regulatory risk, and investors price it on the day. A more diversified funding base (bonds, external commercial borrowing, securitisation) reduces it."]));
C.push(...QA("What would you do with more time or data?", ["Use bank-level data to test whether banks with less spare capital cut targeted lending more; that would separate the capital channel from the signal. Use NBFC bond and commercial-paper data to see whether NBFCs replaced bank funding. And extend the work to unlisted NBFCs, which may depend more on banks."]));
C.push(...QA("If you started again, what would you change?", ["I would collect the firm data from the filings from the start instead of beginning with a ready-made dataset, and I would plan the event study from the beginning, since it turned out to give the sharpest evidence."]));

// ================================================================= 9. weak points
C.push(H1("9. Weak points and how to handle them"));
C.push(P("Examiners will probe these. The best approach is to raise them yourself if the chance comes, then explain what still stands."));
C.push(table(["Weak point", "What to say"], [
  ["H1 is only significant at 10 per cent", `True. I call it weak evidence. It is driven by personal loans (p ${p3(rob("Other personal loans only").p_welch)} alone), and the consumer lenders' share prices and capital disclosures point the same way.`],
  ["H2 is not significant in the simple t-test", "Bank lending to NBFCs is highly seasonal; with month dummies it is significant at 5 per cent. Both tests are reported."],
  ["Only 45 or 46 monthly observations", "Yes, so I avoid overstating results and back the credit data with the event study, which uses daily data."],
  ["Newey-West p-values look too good", "They are smaller mainly because the post-policy months are calmer. I treat them as a check, not the main result."],
  ["Package, not just a capital charge", "Agreed; I say so in the abstract. The results measure the effect of the circular as a whole."],
  ["Rising delinquencies at the same time", "This affects the credit data more than the event study, because share prices reacted on the day of the announcement."],
  ["Fast growth might have slowed anyway", "No significant pre-trend or placebo break; H1 is treated as an upper bound."],
  ["Other news near event days (Bajaj embargo, microfinance easing)", "Results hold without Bajaj and without the microfinance lenders."],
  ["Only 16 listed firms", "The event study cannot speak for small unlisted NBFCs, which may be more bank-dependent; I list this as a limitation."],
  ["Firm-level tests are not significant", "Reported as supporting, descriptive evidence; the sample is small and the rate cycle affected everyone."],
  ["H1 measures bank credit, not NBFC lending", "Correct, and the title says \"bank consumer credit\". NBFCs' own consumer lending is seen only for Bajaj."],
], [3200, 5826]));

// ================================================================= 10. formula sheet
C.push(H1("10. Formula sheet"));
const F = (name, formula, note) => [new Paragraph({ keepNext: true, spacing: { before: 120, after: 40 }, children: [B(name)] }),
  new Paragraph({ spacing: { after: 40 }, indent: { left: 360 }, children: [I(formula)] }),
  ...(note ? [new Paragraph({ spacing: { after: 80 }, indent: { left: 360 }, children: [new TextRun({ text: note, size: 19 })] })] : [])];
C.push(...F("Monthly growth", "g(t) = 100 x [ln(credit at t) - ln(credit at t-1)]", "Close to the percentage change; monthly values add up."));
C.push(...F("Welch t-test", "t = (mean after - mean before) / sqrt(s1²/n1 + s2²/n2)", "Degrees of freedom: (v1 + v2)² / [v1²/(n1-1) + v2²/(n2-1)], with v = s²/n."));
C.push(...F("Regression", "Gap(t) = a + d x Post(t) + c2 x M2(t) + ... + c12 x M12(t) + e(t)", "d is the change in the gap after the policy, holding the calendar month constant."));
C.push(...F("R squared and adjusted R squared", "R² = SS regression / SS total;  adjusted R² = 1 - (1 - R²)(n - 1)/(n - k - 1)", "k = 12 regressors."));
C.push(...F("F-test (ANOVA)", "F = (SS regression / k) / (SS residual / (n - k - 1))", ""));
C.push(...F("Durbin-Watson", "DW = sum of (e(t) - e(t-1))² / sum of e(t)²", "About 2 means no first-order autocorrelation; towards 0 positive, towards 4 negative."));
C.push(...F("Jarque-Bera", "JB = n/6 x [S² + (K - 3)²/4]", "S = skewness, K = kurtosis of residuals; compare with chi-squared, 2 df."));
C.push(...F("Breusch-Pagan", "LM = n x R² from regressing squared residuals on the regressors", "Compare with chi-squared with k df."));
C.push(...F("VIF", "VIF(j) = 1 / (1 - R²(j)), where R²(j) is from regressing regressor j on the others", "Above 10 is usually a concern."));
C.push(...F("Market model", "R(i,t) = alpha(i) + beta(i) x R(market,t) + e(i,t)", "Estimated on days -250 to -31."));
C.push(...F("Abnormal and cumulative abnormal return", "AR(t) = R(t) - [alpha + beta x R(market,t)];  CAR = sum of AR over the window", ""));
C.push(...F("CAR t-statistic", "t = CAR / (residual SD x sqrt(L)), L = days in the window", "Degrees of freedom: 220 - 2 = 218."));
C.push(...F("Sign test", "p = 2 x P(X >= k) with X ~ Binomial(10, 0.5)", "9 of 10: 2 x 11/1024 = 0.021."));
C.push(...F("Paired t-test", "t = mean change / (SD of changes / sqrt(n))", "Used within each group of firms."));
C.push(...F("Economic size", "shortfall % = 100 x [exp(-d x months / 100) - 1]", "d = monthly fall in the gap (pp), months = 14."));

// ================================================================= 11. where to find
C.push(H1("11. Where to find each number"));
C.push(table(["Result", "Dissertation", "Dissertation_Tests.xlsx"], [
  ["Welch t-tests", "Table 7.1", `H1_Tests ${XC["H1_Tests.welch_diff"]} (difference), ${XC["H1_Tests.welch_p"]} (p); same cells in H2_Tests and H2_Alt_Tests`],
  ["Model summary, ANOVA", "Tables 7.2, 7.3", "H1_Tests and H2_Tests, column U, Regression 2 sections"],
  ["Coefficient on Post", "Table 7.4", `H1_Tests ${XC["H1_Tests.post_b"]} and ${XC["H1_Tests.post_p"]}; all coefficients in columns X to AC`],
  ["Newey-West", "Table 7.4", "NeweyWest_Output"],
  ["Diagnostics", "Table 7.5", `H1_Tests from ${XC["H1_Tests.dw"]} (Durbin-Watson) down; VIF in column AC`],
  ["Placebo and trend", "Table 7.6", "H1_Placebo, H2_Placebo"],
  ["Rollback", "Table 7.6", "H2_Rollback"],
  ["Drop-one checks", "Table 7.7", "H1_DropOne (summary at the top)"],
  ["Economic size", "Section 7.5", "H1_Tests and H2_Tests, bottom of column U"],
  ["Market model", "Table 8.1", "ES_MarketModel"],
  ["Circular CARs", "Table 8.2", `ES_E1: NBFCs ${XC["ES_E1.Non-HFC NBFCs (10)"]}; NBFCs minus HFCs ${XC["ES_E1.Non-HFC NBFCs minus HFCs"]}`],
  ["Rollback CARs", "Table 8.3", `ES_E2: NBFCs ${XC["ES_E2.Non-HFC NBFCs (10)"]}`],
  ["Firm day-0 ARs", "Table 8.4", "ES_E1 and ES_E2, section B, column H"],
  ["Sign test, bank-share regression", "Section 8.2 to 8.4", "ES_E1 and ES_E2, sections C and D"],
  ["Placebo days", "Figure 8.2", "ES_PlaceboDays"],
  ["Nifty 500", "Table C.1", "ES_Nifty500"],
  ["Firm tests", "Table 9.2", "Firm_Tests"],
  ["Raw data", "Chapter 4", "Data sheets; full data with sources in Dissertation_Data_Compilation.xlsx"],
], [2600, 1900, 4526]));

// ================================================================= 12. glossary
C.push(H1("12. Glossary"));
C.push(table(["Term", "Meaning"], [
  ["Risk weight", "The percentage of a loan counted as risk-weighted assets. A 125% weight makes Rs 100 of loans count as Rs 125."],
  ["CRAR", "Capital to risk-weighted assets ratio. Regulatory capital divided by risk-weighted assets."],
  ["Capital conservation buffer", "An extra 2.5% of risk-weighted assets that banks hold above the 9% minimum."],
  ["NBFC / HFC", "Non-banking financial company / housing finance company. HFCs were exempt from the bank-to-NBFC measure."],
  ["Priority-sector lending", "Lending to sectors the RBI wants banks to serve (such as agriculture and small borrowers). Bank loans to NBFCs that qualify were exempt."],
  ["Last reporting Friday", "The date to which each monthly RBI credit observation refers."],
  ["Growth gap", "Growth of targeted credit minus growth of comparable exempt credit, in percentage points a month."],
  ["Dummy variable", "A variable that is 1 or 0, such as Post (1 after the policy) or a month dummy."],
  ["Placebo test", "Repeating a test at a date when nothing happened, to check that it does not find an effect."],
  ["Abnormal return", "A stock's actual return minus the return the market model predicts."],
  ["Estimation window", "The 220 trading days used to estimate each stock's normal relationship with the market."],
  ["Event window", "The days over which abnormal returns are added up, such as [0] or [0,+2]."],
  ["Difference portfolio", "Daily return of one portfolio minus another (NBFCs minus HFCs), tested like any portfolio."],
  ["Newey-West standard errors", "Standard errors that allow for unequal variance and correlation between nearby observations."],
  ["Paired t-test", "A t-test on the change within the same units (each firm before and after)."],
  ["QIP", "Qualified institutional placement: a way for listed companies to raise equity from institutional investors."],
], [2600, 6426]));

// ================================================================= document
const FOOT = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] });
const doc = new Document({
  creator: "Nityam Gupta", title: "Viva preparation notes",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: FONT, color: "1F3864" }, paragraph: { spacing: { before: 120, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: FONT, color: "2E4A7A" }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: FONT }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: "n1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 340 } } } }] },
  ] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: FOOT() }, children: C }],
});
fs.mkdirSync("outputs/viva", { recursive: true });
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync("outputs/viva/Viva_Preparation.docx", buf); console.log("outputs/viva/Viva_Preparation.docx written", buf.length); });
