// Builds Dissertation.docx from outputs/final/results.json and figures. Run: node 06_build_dissertation.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, ImageRun, TableOfContents, Footer, PageNumber, LevelFormat, PageBreak,
  NumberFormat, Header, PageBorderDisplay, PageBorderOffsetFrom, PageBorderZOrder, TabStopType,
} = require("docx");

const R = JSON.parse(fs.readFileSync("outputs/final/results.json", "utf8").replace(/\bNaN\b/g, "null"));
const FIG = (f) => fs.readFileSync(path.join("outputs/final", f));
const f2 = (x) => Number(x).toFixed(2);
const f1 = (x) => Number(x).toFixed(1);
const p3 = (x) => (Number(x) < 0.001 ? "< 0.001" : Number(x).toFixed(3));
const H1 = R.H1, H2 = R.H2, GT = R.growth_table;

// ---------------------------------------------------------------- helpers
const FONT = "Times New Roman";
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 360 }, alignment: opts.align || AlignmentType.JUSTIFIED,
  children: (Array.isArray(text) ? text : [text]).map((t) => (typeof t === "string" ? new TextRun(t) : t)),
  ...opts.p,
});
const B = (t) => new TextRun({ text: t, bold: true });
const I = (t) => new TextRun({ text: t, italics: true });
const Hd = (text, level, noBreak = false) => new Paragraph({ heading: level, spacing: { before: 240, after: 120 },
  alignment: level === HeadingLevel.HEADING_1 ? AlignmentType.CENTER : undefined, pageBreakBefore: level === HeadingLevel.HEADING_1 && !noBreak,
  children: [new TextRun(level === HeadingLevel.HEADING_1 ? text.toUpperCase() : text)] });
const bullet = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 80, line: 340 },
  children: (Array.isArray(text) ? text : [text]).map((t) => (typeof t === "string" ? new TextRun(t) : t)) });
const num = (text) => new Paragraph({ numbering: { reference: "num", level: 0 }, spacing: { after: 80, line: 340 },
  children: (Array.isArray(text) ? text : [text]).map((t) => (typeof t === "string" ? new TextRun(t) : t)) });
const caption = (t) => new Paragraph({ keepNext: true, keepLines: true, alignment: AlignmentType.LEFT, spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: t, bold: true, size: 22 })] });
const note = (t) => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200 },
  children: [new TextRun({ text: t, italics: true, size: 18 })] });
const fig = (file, w, h, cap, src) => [caption(cap),
  new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, children: [new ImageRun({ type: "png", data: FIG(file), transformation: { width: w, height: h } })] }),
  note(src)];
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
// borderless two-column block (cover page, signature blocks)
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const twoCol = (left, right, opts = {}) => new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
  borders: { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB },
  rows: [new TableRow({ children: [left, right].map((col, k) => new TableCell({ width: { size: 4680, type: WidthType.DXA },
    borders: { top: NB, bottom: NB, left: NB, right: NB },
    children: col.map((t, j) => new Paragraph({ alignment: k === 1 && opts.rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT, spacing: { after: 140 },
      children: [new TextRun({ text: t, bold: opts.boldFirst ? j === 0 : !!opts.bold, underline: opts.boldFirst && j === 0 ? {} : undefined })] })) })) })] });

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const CONTENT_W = 9026; // A4 with 1" margins, DXA
function table(head, rows, widths) {
  const tot = widths.reduce((a, b) => a + b, 0);
  const cell = (t, i, isHead) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
    shading: isHead ? { fill: "E7EDF5", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
      children: [new TextRun({ text: String(t), bold: isHead, size: 19 })] })],
  });
  return new Table({ width: { size: tot, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, true)) }),
           ...rows.map((r) => new TableRow({ children: r.map((t, i) => cell(t, i, false)) }))] });
}

// ---------------------------------------------------------------- numbers used in text
const h1a = H1.its.annualised_change, h2a = H2.its.annualised_change;
const gtCC = GT["Credit cards"], gtPL = GT["Other personal loans"], gtNB = GT["NBFCs ex-HFC"], gtNF = GT["Non-food credit"];
const gtH = GT["Housing"], gtV = GT["Vehicle loans"], gtE = GT["Education"], gtG = GT["Gold loans"], gtHF = GT["HFCs"];
const gtCore = GT["Non-food ex NBFC, cards, other PL"];
const PRE = "Pre-policy (Apr-21 to mid-Nov-23)", POST = "Post-policy (Dec-23 to Mar-25)", ROLL = "After rollback (Apr-25 to Jul-26)";
const fc = R.firm_cof, fch = R.firm_cof_hfc, fb = R.firm_bankshare, bw = R.bajaj_within, cnt = R.counts;
const MG = R.magnitudes, FD = R.firm_desc, A = R.audit;
const lakh = (x) => (Number(x) / 100000).toFixed(1);
const fdv = (v, g, p) => FD.find((r) => r.Variable.startsWith(v) && r.Group === g && r.Period.startsWith(p));
const I1 = R.inference.H1, I2 = R.inference.H2;
const im = (T, m) => T.find((r) => r.Method.startsWith(m));
const MS1 = R.model_summary[0], MS2 = R.model_summary[1];
const pct = (d, n) => f1(100 * (Math.exp(-d * n / 100) - 1));          // cumulative shortfall in per cent
const TC = A.trend_cf, PT = A.pretrend, EV = A.event;
const sub = (lbl) => A.bankshare_subsamples.find((r) => r.label.startsWith(lbl));
const RC = A.recheck;
const chgF = (c) => A.bankshare_firm_change[c];

// ---------------------------------------------------------------- content
const C = [];

// Student and guide details
const ST = { name: "Nityam Gupta", prog: "MBA (Full-Time)", batch: "2025–27", roll: "FT-25-235", date: "03/10/2026" };
const GD = { name: "Dr. Dezy Kumari", desig: "Assistant Professor, Finance", fac: "Faculty of Management Studies", uni: "University of Delhi" };
const TITLE = "Capital Requirements and Credit Supply: The Effect of the RBI's November 2023 Risk-Weight Measures on Bank Consumer Credit and Bank Funding of NBFCs in India, 2021–2026";

// Cover page (template layout)
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 360, after: 480 },
  children: [new ImageRun({ type: "png", data: fs.readFileSync("assets/fms_logo.png"), transformation: { width: 250, height: 121 } })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: "FINAL YEAR DISSERTATION", bold: true, size: 28 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 720, line: 360 }, children: [new TextRun({ text: `“${TITLE}”`, bold: true, italics: true, size: 30 })] }));
C.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 1200, line: 360 }, children: [new TextRun({ text: "A project report submitted in partial fulfilment of the requirements for the degree of Master of Business Administration (Full-Time) from the Faculty of Management Studies (FMS), University of Delhi.", bold: true })] }));
C.push(twoCol(["Submitted by:", ST.name, `${ST.prog}, Batch ${ST.batch}`, `Roll no: ${ST.roll}`, `Date: ${ST.date}`],
  ["Under guidance of:", GD.name, GD.desig, GD.fac, GD.uni], { boldFirst: true }));
const COVER_END = C.length;

// Certificate
C.push(Hd("Certificate", HeadingLevel.HEADING_1, true));
C.push(P(`Certified that the dissertation entitled “${TITLE}”, submitted by ${ST.name}, ${ST.prog}, Batch ${ST.batch}, Roll no. ${ST.roll}, in partial fulfilment of the requirements for the degree of Master of Business Administration of the Faculty of Management Studies, University of Delhi, is a bona fide record of research carried out by the student under my supervision. All sources of data and published work used have been cited. The work has not been submitted for any other degree or diploma.`));
C.push(P(`Date: ${ST.date}`, { p: { spacing: { before: 480, after: 1200 } } }));
C.push(twoCol([GD.name, "(Project Guide)", GD.desig, GD.fac, GD.uni],
  [ST.name, `${ST.prog}, ${ST.batch}`, `Roll no: ${ST.roll}`], { rightAlign: true }));

// Declaration
C.push(Hd("Declaration", HeadingLevel.HEADING_1));
C.push(P(`I hereby declare that this dissertation titled “${TITLE}” is my original work carried out under the guidance of Dr. Dezy Kumari, Assistant Professor (Finance), Faculty of Management Studies, University of Delhi. It has not been submitted elsewhere for any degree or diploma. All sources of data and published work used have been acknowledged.`));
C.push(P([B("Use of AI tools. "), "Artificial intelligence (AI) tools were used for data scraping, that is, extracting figures from the companies' quarterly investor presentations, and for the Python-based data analysis. Every extracted value is recorded with the document and page it came from, so that it can be checked against the original source (Section 4.3 reports the checks carried out)."]));
C.push(P("Signature: ____________________", { align: AlignmentType.LEFT, p: { spacing: { before: 720, after: 120 } } }));
for (const t of [ST.name, `${ST.prog}, Batch ${ST.batch}`, `Roll no: ${ST.roll}`, `Date: ${ST.date}`])
  C.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 80 }, children: [new TextRun(t)] }));

// Acknowledgement
C.push(Hd("Acknowledgement", HeadingLevel.HEADING_1));
C.push(P("I would like to express my sincere gratitude to my guide, Dr. Dezy Kumari, Assistant Professor (Finance), Faculty of Management Studies, University of Delhi, for her guidance throughout this dissertation. Her advice to focus the study on a small number of well-defined hypotheses shaped the final form of this work, and her feedback at each stage helped me turn a broad question into a focused piece of research."));
C.push(P("I am grateful to the faculty members of the Faculty of Management Studies, whose teaching in finance, economics and research methods gave me the foundation for this study, and to the administrative and library staff for their support."));
C.push(P("I thank my batchmates for the many discussions, suggestions and words of encouragement during the long months of collecting and analysing the data. Their company made the work lighter."));
C.push(P("Above all, I thank my parents and my sister, whose patience, encouragement and constant belief in me kept me going. This work would not have been possible without their support."));
for (const t of [ST.name, `${ST.prog}, Batch ${ST.batch}`, `Roll no: ${ST.roll}`])
  C.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: t === ST.name ? 720 : 0, after: 80 }, children: [new TextRun(t)] }));

// Abstract
C.push(Hd("Abstract", HeadingLevel.HEADING_1));
C.push(P(`On 16 November 2023 the Reserve Bank of India (RBI) raised the risk weights on two channels of credit: banks' and NBFCs' unsecured consumer credit (personal loans and credit cards), and banks' lending to non-banking financial companies (NBFCs), with housing finance companies (HFCs) exempt. This dissertation asks whether these measures slowed the targeted flows of bank credit. Using monthly RBI data on the sectoral deployment of bank credit, it compares targeted with exempt credit categories before and after the circular (a difference-in-differences design), using the RBI's merger-adjusted series to remove the effect of the HDFC merger. Bank credit to NBFCs (excluding HFCs) grew about ${f1(-h2a)} percentage points a year more slowly than other bank credit after the circular; the estimate is marginally significant (p = ${p3(H2.its.p)} in the main specification and between ${p3(im(I2, "Month dummies, White").p)} and ${p3(im(I2, "Month dummies, fixed-b").p)} once calendar-month seasonality is controlled for), and the gap reversed after the RBI withdrew this measure in April 2025, although the rebound is not statistically significant. Targeted bank consumer credit (cards and other personal loans) grew about ${f1(-h1a)} percentage points a year more slowly than exempt household credit. This estimate is consistent in sign across specifications and significant with autocorrelation-robust inference (p = ${p3(H1.its.p)}; fixed-b p = ${p3(im(I1, "Fixed-b HAC").p)}), but not with classical standard errors (p = ${p3(im(I1, "Classical").p)}), and it shrinks to ${f2(-TC.H1.from_jul21.delta)} percentage points a month if a mild pre-policy decline in the gap is assumed to continue. Placebo tests on a fictitious November 2022 date show no effect. Because the circular also required board-approved exposure limits, and because unsecured credit was growing unusually fast and delinquencies were rising, the estimates measure the combined effect of the RBI's November 2023 package rather than a pure capital-cost effect. Firm-level data extracted from ${cnt.decks} investor presentations of ${cnt.firms} listed NBFCs and HFCs show that the higher risk weights cut two consumer lenders' capital ratios by about 3–4 percentage points; evidence on how NBFCs adjusted their funding costs, funding mix and lending is descriptive and inconclusive.`));
C.push(P([B("Keywords: "), "risk weights, capital requirements, credit supply, NBFCs, consumer credit, macroprudential policy, difference-in-differences, India"]));

// Table of contents
C.push(new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 480 }, children: [new TextRun({ text: "Table of Contents", bold: true, size: 26 })] }));
C.push(new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }));
const FRONT_END = C.length;

// Executive summary
C.push(Hd("Executive Summary", HeadingLevel.HEADING_1, true));
C.push(P([B("Context. "), "On 16 November 2023 the Reserve Bank of India raised the risk weights on unsecured consumer credit (personal loans and credit cards) and on banks' loans to NBFCs (housing finance companies exempt), and required lenders to set board-approved limits on unsecured consumer exposures by 29 February 2024. The bank-to-NBFC risk weight was withdrawn from 1 April 2025; the consumer-credit risk weight was not."]));
C.push(P([B("Objective. "), "To measure whether the November 2023 measures slowed bank credit through each targeted channel, and to document how listed NBFCs adjusted."]));
C.push(P([B("Data. "), `Monthly RBI data on sectoral deployment of bank credit (January 2019 to July 2026), using the RBI's merger-adjusted series for the period affected by the HDFC merger; and firm-level data on 16 listed NBFCs and HFCs (cost of funds, bank share of borrowings, Bajaj Finance segment assets and company statements), extracted with AI assistance from ${cnt.decks} quarterly investor presentations with every value logged to its source page.`]));
C.push(P([B("Method. "), "Difference-in-differences: credit categories that faced the higher risk weight are compared with closely related exempt categories before and after the circular. Inference uses Newey–West standard errors with the rule-of-thumb lag length and t-distribution p-values, checked against fixed-b, bootstrap, seasonal and quarterly alternatives; placebo, pre-trend and event-study tests check the comparison groups."]));
C.push(P([B("Key findings.")]));
C.push(bullet(`Bank credit to NBFCs (excluding HFCs) grew about ${f1(-h2a)} percentage points a year more slowly than total bank credit after the circular. The effect is marginally significant (p = ${p3(H2.its.p)}; ${p3(im(I2, "Month dummies, White").p)}–${p3(im(I2, "Month dummies, fixed-b").p)} with seasonal controls; up to ${p3(im(I2, "Fixed-b HAC").p)} without them) and reversed after the April 2025 rollback, though the rebound is not significant (p = ${p3(H2.rollback.p)}).`));
C.push(bullet(`Targeted bank consumer credit grew about ${f1(-h1a)} percentage points a year more slowly than exempt household credit. The sign is consistent across specifications and the effect is significant with autocorrelation-robust inference (p = ${p3(H1.its.p)}; fixed-b p = ${p3(im(I1, "Fixed-b HAC").p)}), but not with classical OLS (p = ${p3(im(I1, "Classical").p)}) or quarterly averages (p = ${p3(im(I1, "Quarterly").p)}), and it is sensitive to assumptions about pre-policy trends.`));
C.push(bullet(`Taken at face value, the estimates imply roughly ₹${lakh(MG.h1_shortfall)} lakh crore less targeted consumer credit and ₹${lakh(MG.h2_shortfall)} lakh crore less bank funding to NBFCs by January 2025 than if pre-policy relative growth had continued; both figures inherit the uncertainty above.`));
C.push(bullet("The estimates capture the whole November 2023 package (higher risk weights plus required exposure limits) at a time when banks were well capitalised and delinquencies on unsecured loans were rising, so they should not be read as a pure capital-requirement effect."));
C.push(bullet("Firm-level evidence is descriptive and inconclusive: two consumer lenders reported capital-ratio cuts of about 3–4 percentage points, but cost-of-funds, borrowing-mix and segment-growth tests show no significant effect, and the main within-firm case (Bajaj Finance) is confounded by an RBI embargo on two of its lending products."));
C.push(P([B("Implications. "), "The evidence is consistent with targeted risk weights slowing bank credit to the targeted segments, most clearly for bank lending to NBFCs, which slowed and then recovered when the measure was withdrawn. The size of the effects is uncertain, and the study cannot separate the capital cost from the supervisory signal of the package."]));

// Rationale of the study
C.push(Hd("Rationale of the Study", HeadingLevel.HEADING_1));
C.push(P("NBFCs now supply a large share of India's retail credit and depend heavily on banks for funding. A regulatory change that raises the cost of both unsecured consumer lending and bank lending to NBFCs therefore matters to the regulator, to banks, to NBFC managers and to investors. The study is worthwhile for four reasons."));
C.push(num([B("Policy relevance. "), "The November 2023 measures were one of the RBI's most significant macroprudential interventions since the IL&FS crisis, and were partly reversed in 2025. Whether they worked, and through which channel, is relevant to how such tools are used in future."]));
C.push(num([B("A natural experiment. "), "The circular changed risk weights for clearly defined categories on a known date, left comparable categories untouched and was later partly withdrawn. This allows the change to be compared with untargeted credit that faced the same interest-rate cycle."]));
C.push(num([B("Managerial relevance. "), "For NBFC managers the episode shows how dependence on a single funding source translates into regulatory risk, and how capital, pricing and funding mix can be used to absorb such a shock."]));
C.push(num([B("Gap in evidence. "), "Existing discussion of the measures is largely descriptive. There is little formal evidence that separates the consumer-credit and bank-funding channels or that uses NBFCs' own disclosures to trace their response."]));

// ================= Chapter 1
C.push(Hd("1. Introduction", HeadingLevel.HEADING_1));
C.push(Hd("1.1 Background and motivation", HeadingLevel.HEADING_2));
C.push(P("Non-banking financial companies (NBFCs) have become one of the main channels of retail credit in India. Between 2021 and 2023, unsecured consumer lending — personal loans and credit cards — grew far faster than overall credit, and NBFCs both lent directly to households and borrowed heavily from banks to fund that lending. The Reserve Bank of India (RBI) viewed this combination as a build-up of risk: rapid unsecured lending on the asset side, and growing interdependence between banks and NBFCs on the funding side."));
C.push(P("On 16 November 2023 the RBI responded with a single circular. First, it increased the risk weight on consumer credit — excluding housing, education, vehicle and gold-backed loans, and (for NBFCs) microfinance loans — from 100 per cent to 125 per cent, and raised risk weights on credit card receivables. Second, it increased by 25 percentage points the risk weight that banks attach to their loans to most NBFCs, while exempting housing finance companies (HFCs) and loans eligible for priority-sector classification. Third, it required banks and NBFCs to review their sectoral exposure limits and put in place board-approved limits for unsecured consumer credit by 29 February 2024 (RBI, 2023a). In February 2025 the RBI withdrew the higher risk weight on bank loans to NBFCs from 1 April 2025, but left the consumer-credit risk weights in place (RBI, 2025b)."));
C.push(P("This policy is a useful natural experiment. It changed risk weights for clearly defined categories of credit at a known date, left closely comparable categories untouched, and was later partly reversed. That design makes it possible to compare targeted credit with untargeted credit that faced the same interest-rate cycle. It also has limits, discussed throughout: the circular was a package rather than a pure capital shock, and the RBI acted because the targeted credit was growing unusually fast."));
C.push(Hd("1.2 Research problem", HeadingLevel.HEADING_2));
C.push(P("Market commentary since 2023 suggests that unsecured lending slowed after the RBI's measures, but it is not clear how much of the slowdown was caused by the measures, which channel they worked through, and how NBFCs responded. This study measures the change in bank credit through each targeted channel relative to exempt credit, and documents what listed NBFCs disclosed about their capital, funding and lending."));
C.push(Hd("1.3 Objectives and hypotheses", HeadingLevel.HEADING_2));
C.push(P("Following the guide's direction to focus on one or two hypotheses, the study tests the two levers of the November 2023 circular, both measured in RBI data on bank credit:"));
C.push(bullet([B("H1 (consumer-credit channel): "), "The November 2023 measures reduced the growth of bank credit to the targeted consumer categories (credit cards and other personal loans) relative to exempt household credit categories."]));
C.push(bullet([B("H2 (bank-funding channel): "), "The higher risk weight on bank exposures to NBFCs reduced the growth of bank credit to NBFCs (excluding the exempt HFCs) relative to banks' overall lending."]));
C.push(P("These hypotheses correspond to H1 and H2 of the original synopsis, re-specified so that each is tested on the group that the regulation targeted. H1 measures bank consumer credit, not NBFCs' own consumer lending, which is observed only for one firm (Bajaj Finance); the title of the dissertation reflects this. The synopsis versions — firm-level effects on NBFCs — are examined as supporting, descriptive evidence (Chapter 8)."));
C.push(Hd("1.4 Contribution", HeadingLevel.HEADING_2));
C.push(P(`The study makes two contributions. First, it provides a before-and-after comparison of targeted and exempt credit categories in official RBI data, using the RBI's merger-adjusted series, with placebo, pre-trend and event-study checks and a systematic comparison of inference methods. Second, it assembles a firm-level dataset from ${cnt.decks} quarterly investor presentations of 16 listed NBFCs and HFCs — cost of funds, bank share of borrowings, segment assets and company disclosures of the policy's capital impact — with every value logged to its source page.`));
C.push(Hd("1.5 Structure", HeadingLevel.HEADING_2));
C.push(P("Chapter 2 describes the regulatory background and Chapter 3 reviews the literature. Chapter 4 describes the data and how they were checked. Chapter 5 sets out the theoretical framework, variables and formal hypotheses, and Chapter 6 the econometric methodology. Chapters 7 and 8 present the results for H1 and H2, together with the firm-level evidence. Chapter 9 collects the model diagnostics, the sensitivity of inference and the pre-trend checks for both models. Chapter 10 discusses the findings and threats to validity, and Chapter 11 concludes."));

// ================= Chapter 2
C.push(Hd("2. Regulatory Background", HeadingLevel.HEADING_1));
C.push(Hd("2.1 Risk weights and capital requirements", HeadingLevel.HEADING_2));
C.push(P("Under the Basel framework applied in India, banks and NBFCs must hold regulatory capital equal to a minimum percentage of their risk-weighted assets (RWA). Each exposure is multiplied by a risk weight; a 125 per cent risk weight means that every ₹100 lent counts as ₹125 of RWA. Raising the risk weight therefore increases the capital a lender must hold against the same loan or, equivalently, lowers its capital-to-risk-weighted-assets ratio (CRAR) if capital does not change. For an NBFC that must keep CRAR above 15 per cent, or a bank that must meet its own capital requirements, a higher risk weight raises the effective cost of supplying that type of credit."));
C.push(Hd("2.2 The November 2023 circular", HeadingLevel.HEADING_2));
C.push(P("The RBI's circular of 16 November 2023 (“Regulatory measures towards consumer credit and bank credit to NBFCs”, RBI/2023-24/85) contained the measures summarised in Table 2.1."));
C.push(caption("Table 2.1: Main measures, November 2023 and February 2025"));
C.push(table(["Measure", "Change", "Exemptions / notes"], [
  ["Consumer credit (banks)", "Risk weight 100% → 125%", "Excludes housing, education, vehicle and gold-secured loans"],
  ["Retail loans (NBFCs)", "Risk weight 100% → 125%", "Also excludes microfinance/SHG loans"],
  ["Credit card receivables", "Banks 125% → 150%; NBFCs 100% → 125%", "—"],
  ["Bank exposures to NBFCs", "+25 percentage points where the rating-based risk weight was below 100%", "Excludes HFCs, core investment companies and loans eligible for priority-sector classification"],
  ["Credit standards", "Review sectoral exposure limits; board-approved limits for all unsecured consumer credit", "To be implemented by 29 February 2024"],
  ["February 2025 circulars", "Higher risk weight on bank exposures to NBFCs withdrawn from 1 April 2025; bank microfinance loans in the nature of consumer credit moved back to 100%", "Other consumer-credit risk weights unchanged"],
], [2300, 3600, 3460]));
C.push(note("Source: RBI circulars RBI/2023-24/85 (16 November 2023), RBI/2024-25/119 and RBI/2024-25/120 (25 February 2025)."));
C.push(P("Three features of the circular matter for the empirical design. The consumer-credit measure targeted unsecured personal loans and cards but left housing, vehicle and education loans untouched, so exempt household categories provide a comparison group for H1. The bank-to-NBFC measure applied to NBFCs but not to HFCs or to priority-sector-eligible on-lending (which covers much bank lending to NBFC-MFIs), so bank credit to HFCs and to the banking system as a whole provide comparison groups for H2. And the circular was a package: alongside the risk weights it required board-approved exposure limits, so any effect measured in the data is the effect of the package, not of the capital charge alone."));
C.push(Hd("2.3 Macroeconomic and banking setting", HeadingLevel.HEADING_2));
C.push(P("The policy came after the RBI had raised the repo rate by 250 basis points between May 2022 and February 2023, to 6.5 per cent, where it stayed until February 2025. Funding costs across the financial system were therefore already rising during 2022–23 as loans repriced. Banks were well capitalised: the system-wide CRAR of scheduled commercial banks was 16.8 per cent in September 2023 (RBI Financial Stability Report, December 2023), well above the regulatory minimum. A 25-percentage-point rise in the risk weight on a few loan categories was therefore unlikely to make capital requirements bind for most banks; its effect may have worked as much through supervisory signalling as through the cost of capital (Section 10.2)."));
C.push(Hd("2.4 The NBFC sector before the policy", HeadingLevel.HEADING_2));
C.push(P("NBFCs are regulated by the RBI but, unlike banks, cannot accept demand deposits and do not have access to the RBI's liquidity facilities in the same way. They fund themselves through bank loans, bonds (non-convertible debentures), commercial paper, securitisation and, for some, public deposits. The default of IL&FS in 2018 showed how quickly funding can dry up for NBFCs that depend on short-term market borrowing, and banks subsequently became an even more important source of NBFC funding. In October 2021 the RBI introduced a Scale-Based Regulation framework, placing NBFCs in base, middle, upper and top layers with progressively stricter requirements; most of the firms in this study fall in the middle or upper layers."));
C.push(P(`By October 2023 the two channels targeted by the RBI were large. Outstanding bank credit was ₹${lakh(MG.cards_oct23)} lakh crore to credit cards and ₹${lakh(MG.pl_oct23)} lakh crore in other personal loans, and bank credit to NBFCs excluding HFCs stood at ₹${lakh(MG.nbfc_oct23)} lakh crore — about ${f1(100 * MG.nbfc_oct23 / MG.nonfood_oct23)} per cent of banks' total non-food credit. All three had been growing at more than 20 per cent a year (Table 4.1), well above overall credit growth. Note that the RBI series for bank credit to NBFCs includes public financial institutions and infrastructure NBFCs, many of which are rated AAA or are government-owned; for loans to such borrowers the rating-based risk weight is often already low, and loans that qualify as priority-sector on-lending were exempt. The aggregate H2 estimate is therefore a diluted, intent-to-treat effect.`));

// ================= Chapter 3
C.push(Hd("3. Literature Review", HeadingLevel.HEADING_1));
C.push(Hd("3.1 Capital requirements and credit supply", HeadingLevel.HEADING_2));
C.push(P("A large literature shows that capital requirements affect the supply of bank credit. Kashyap and Stein (2004) explain how risk-sensitive capital rules can amplify credit cycles. Aiyar, Calomiris and Wieladek (2014) use changes in bank-specific capital requirements in the United Kingdom to show that higher requirements reduce lending, while part of the gap is filled by foreign branches outside the regulator's reach. Jiménez, Ongena, Peydró and Saurina (2017) find that Spain's dynamic provisioning — a capital buffer — dampened credit cycles, with effects concentrated in more constrained banks. Behn, Haselmann and Wachtel (2016) show that when model-based risk weights rose during the 2008 crisis, German banks cut lending more for loans whose risk weights increased. Basten (2020) studies Switzerland's sectoral countercyclical capital buffer on mortgages and finds that banks raised mortgage rates, with more capital-constrained banks raising them more. Together these studies suggest that raising risk weights on a category of loans should reduce credit supply to that category, which is the logic behind H1, and that the size of the effect depends on how close banks are to their capital constraints."));
C.push(Hd("3.2 Bank funding and shadow banks", HeadingLevel.HEADING_2));
C.push(P("The bank lending channel literature (Bernanke and Blinder, 1988; Khwaja and Mian, 2008) shows that shocks to banks' ability or willingness to lend are passed on to borrowers who depend on bank credit. Khwaja and Mian (2008) use Pakistani loan-level data to show that firms borrowing from affected banks could not fully replace the lost funding. Irani, Iyer, Meisenzahl and Peydró (2021) show that tighter capital regulation of banks shifts lending activity toward non-bank lenders. In India, Acharya, Khandwala and Öncü (2013) document the close funding links between banks and NBFCs. Raising the risk weight on bank loans to NBFCs is a direct shock to this channel; H2 tests whether it reduced bank credit to NBFCs."));
C.push(Hd("3.3 Macroprudential policy and targeted risk weights", HeadingLevel.HEADING_2));
C.push(P("Higher risk weights on specific loan categories are a macroprudential tool: instead of tightening credit across the economy through interest rates, the regulator raises the cost of supplying the type of credit it considers risky. India has a domestic precedent. During the 2004–08 credit boom the RBI raised sectoral risk weights several times: on housing loans to individuals (from 50 to 75 per cent in December 2004), on commercial real estate (from 100 to 125 per cent in July 2005, and further later), and on consumer credit and capital-market exposures [VERIFY: exact dates and levels of the 2004–06 consumer-credit and capital-market risk-weight increases]. The studies above suggest three predictions for such a measure. First, credit growth in the targeted category should slow relative to similar untargeted categories. Second, lenders may respond not only by lending less but by raising prices, raising capital or shifting toward exempt lending. Third, part of the targeted activity may move to lenders or funding sources outside the measure's reach (Aiyar, Calomiris and Wieladek, 2014; Irani et al., 2021). The November 2023 circular is unusual because it targeted both a type of borrower (unsecured consumers) and a lender-funding link (banks to NBFCs) at the same time."));
C.push(Hd("3.4 Methodological literature", HeadingLevel.HEADING_2));
C.push(P("Policy evaluations of this kind typically compare affected and unaffected groups before and after the change (difference-in-differences). Two inference issues are relevant here. With aggregate monthly series, errors are serially correlated, so standard errors robust to autocorrelation (Newey and West, 1987) or to cross-sectional and serial dependence in panels (Driscoll and Kraay, 1998) are needed. In short samples these estimators can be badly sized, and fixed-b inference (Kiefer and Vogelsang, 2005) gives more reliable p-values. With a small number of firms, conventional cluster-robust standard errors over-reject; the wild cluster bootstrap-t (Cameron, Gelbach and Miller, 2008), with six-point weights for very few clusters (Webb, 2023), is the recommended remedy. Callaway, Goodman-Bacon and Sant'Anna (2024) discuss the interpretation of designs with continuous treatment intensity, relevant to the firm-level exposure tests."));
C.push(Hd("3.5 Research gap", HeadingLevel.HEADING_2));
C.push(P("Existing evidence on the November 2023 measures is largely descriptive (RBI Financial Stability Reports and market commentary). There is little formal evidence that separates the consumer-credit and bank-funding channels, uses exempt categories as controls, or documents how listed NBFCs adjusted using their own disclosures. This study addresses that gap."));

// ================= Chapter 4
C.push(Hd("4. Data", HeadingLevel.HEADING_1));
C.push(Hd("4.1 Aggregate data: RBI sectoral deployment of bank credit", HeadingLevel.HEADING_2));
C.push(P(`The main dataset is the RBI's table “Deployment of Gross Bank Credit by Major Sectors” (Database on Indian Economy, Table 15), with ${cnt.rbi_months} monthly observations from January 2019 to July 2026. Each observation refers to the last reporting Friday of the month (for example 17 November 2023 and 29 December 2023). It reports outstanding bank credit (₹ crore) to credit cards, other personal loans, housing, vehicle loans, education, loans against gold jewellery, NBFCs, HFCs and total non-food credit. Bank credit to NBFCs excluding HFCs is computed as NBFC credit minus HFC credit. Growth is measured as the monthly change in the natural logarithm of outstanding credit (×100, approximately per cent per month). The data come from a return covering banks that account for about 95 per cent of non-food credit.`));
C.push(P("One adjustment is required. On 1 July 2023 HDFC Ltd, then India's largest HFC, merged into HDFC Bank. Its loan book moved onto a bank balance sheet and banks' loans to HDFC Ltd disappeared. For bank credit to NBFCs and HFCs this is a one-off level drop, so only the merger month (28 July 2023) is excluded. For housing, other personal loans and total non-food credit the merger is not only a level shift: after July 2023 their growth also includes growth of the former HDFC Ltd book. The RBI publishes these three series excluding the impact of the merger from July 2023 to June 2025, and this study uses those merger-adjusted series for growth over that period, which covers the whole post-policy estimation window. After June 2025 only the reported series are available."));
C.push(caption("Table 4.1: Annualised growth of bank credit by category (per cent per year)"));
const gtRow = (name, o) => [name, f1(o[PRE]), f1(o[POST]), f1(o[ROLL])];
C.push(table(["Category", "Pre-policy (Apr-21 to mid-Nov-23)", "Post-policy (Dec-23 to Mar-25)", "After rollback (Apr-25 to Jul-26)"], [
  gtRow("Credit cards (targeted, H1)", gtCC), gtRow("Other personal loans (targeted, H1)", gtPL),
  gtRow("Housing (exempt)", gtH), gtRow("Vehicle loans (exempt)", gtV), gtRow("Education (exempt)", gtE), gtRow("Gold loans (exempt)", gtG),
  gtRow("NBFCs ex-HFC (targeted, H2)", gtNB), gtRow("HFCs (exempt)", gtHF), gtRow("Total non-food credit (benchmark, H2)", gtNF),
  gtRow("Non-food excl. NBFCs, cards, other PL (alt. benchmark)", gtCore),
], [3560, 1900, 1950, 1950]));
C.push(note("Average monthly log growth × 12. Merger month excluded; RBI merger-adjusted series used for housing, other personal loans and non-food credit from August 2023 to June 2025. The pre-policy period ends with the observation of 17 November 2023, one day after the circular. Source: RBI DBIE Table 15; author's calculations."));
C.push(P(`Table 4.1 shows the raw pattern. Bank credit to credit cards grew ${f1(gtCC[PRE])} per cent a year before the policy and ${f1(gtCC[POST])} per cent after; other personal loans slowed from ${f1(gtPL[PRE])} to ${f1(gtPL[POST])} per cent. Exempt categories slowed much less (vehicle ${f1(gtV[PRE])} → ${f1(gtV[POST])}) or accelerated (housing ${f1(gtH[PRE])} → ${f1(gtH[POST])}; education ${f1(gtE[PRE])} → ${f1(gtE[POST])}). Bank credit to NBFCs excluding HFCs slowed from ${f1(gtNB[PRE])} to ${f1(gtNB[POST])} per cent, while total non-food credit eased only from ${f1(gtNF[PRE])} to ${f1(gtNF[POST])} per cent. After the April 2025 rollback, bank credit to NBFCs recovered to ${f1(gtNB[ROLL])} per cent. Card credit — whose higher risk weight was not reversed — slowed further to ${f1(gtCC[ROLL])} per cent, but this cannot be read as evidence that the risk weight still binds: delinquencies on cards and small personal loans were rising over the same period (Section 10.2). Bank gold loans grew very rapidly after the policy (${f1(gtG[POST])} per cent a year), but this mainly reflects a sharp rise in gold prices and the reclassification of banks' agricultural gold loans as retail gold loans (Business Standard, 2025), so it is weak evidence of a policy-driven shift toward secured credit.`));
C.push(Hd("4.2 Firm-level data: extracted from investor presentations", HeadingLevel.HEADING_2));
C.push(P(`The firm-level sample consists of 16 listed NBFCs and HFCs: six HFCs (AAVAS Financiers, Can Fin Homes, Home First Finance, LIC Housing Finance, PNB Housing Finance and Repco Home Finance) and ten NBFCs (Arman Financial Services, Bajaj Finance, Cholamandalam Investment & Finance, CreditAccess Grameen, IIFL Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, Poonawalla Fincorp, SBI Cards and Payment Services, and Shriram Finance). For each company, all quarterly investor presentations from Q1 FY2021 to Q4 FY2026 were collected — ${cnt.decks} documents in total. The following variables were extracted from them by AI-assisted extraction, with each value logged to the presentation and page it came from:`));
C.push(bullet([B("Cost of funds"), ` (${cnt.cof_obs} company-quarters, all 16 firms): the company's reported cost of borrowing; where not reported, annualised finance cost divided by average borrowings from the same presentation.`]));
C.push(bullet([B("Bank share of borrowings"), ` (${cnt.bankshare_obs} company-quarters, ${cnt.bankshare_firms} firms): the share of borrowings from banks shown in each company's borrowing-mix slide (Poonawalla Fincorp does not disclose its mix).`]));
C.push(bullet([B("Pre-policy bank dependence"), ": the bank share of borrowings at 30 September 2023, checked against each company's Q2 FY2024 presentation."]));
C.push(bullet([B("Bajaj Finance segment-wise assets under management"), " (24 quarters), separating consumer segments affected by the rule from exempt segments (Table 8.3)."]));
C.push(bullet([B("Company statements"), " quantifying the effect of the RBI's measures on capital ratios and pricing."]));
C.push(Hd("4.3 Data checks", HeadingLevel.HEADING_2));
C.push(P(`Every extracted value was re-checked automatically against the text of the presentation it is attributed to. Of ${RC.total} cost-of-funds and bank-share values, ${RC.computed} are computed by formula (for example finance cost divided by average borrowings) and ${RC.no_text} were read from chart labels that do not appear in the text layer; neither can be matched verbatim. Of the remaining ${RC.checkable} values, ${RC.found} (${f1(100 * RC.found / RC.checkable)} per cent) were found verbatim in the cited presentation; the ${RC.checkable - RC.found} that were not are mostly ratios calculated from rupee amounts (Repco Home Finance, Muthoot Finance).`));
C.push(P("The project began with a pre-compiled firm-level dataset (file nbfc_quarterly_data-v3.csv) [VERIFY: student to state who supplied this file or where it was obtained]. Checking it against the companies' own filings showed that several key fields were not reported data. For five lenders, “unsecured” assets were a fixed percentage of total assets in every quarter; for Shriram Finance, the series mixed an assumed 9 per cent share with two quarters of actual personal-loan figures. Bank borrowing shares differed materially from the filings for several firms (for example Arman Financial 71 per cent in the dataset versus 32 per cent in its filing; Muthoot Finance 41 versus 65 per cent). Headline figures such as total assets and gross NPA ratios were generally accurate. All firm-level variables used in this dissertation were therefore re-extracted from primary sources, and the original dataset is not used in any result reported here. An early test using the unverified bank shares appeared to show a significant effect on borrowing costs; with verified shares the effect disappeared."));
C.push(caption("Table 4.2: Firm sample and pre-policy bank dependence (30 September 2023)"));
const ver = [
  ["AAVAS Financiers", "HFC", "44.5", "49.6"], ["Arman Financial Services", "NBFC (MFI)", "71.0", "32.3"], ["Bajaj Finance", "NBFC", "23.0", "23.0"],
  ["Can Fin Homes", "HFC", "54.0", "57.0"], ["Cholamandalam Inv. & Fin.", "NBFC", "53.5", "60.0"], ["CreditAccess Grameen", "NBFC (MFI)", "58.0", "52.9"],
  ["Home First Finance", "HFC", "57.5", "54.0"], ["IIFL Finance", "NBFC", "56.0", "55.3"], ["LIC Housing Finance", "HFC", "24.5", "33.0"],
  ["M&M Financial Services", "NBFC", "56.5", "57.2"], ["Muthoot Finance", "NBFC (gold)", "41.0", "65.4"], ["PNB Housing Finance", "HFC", "40.2", "40.2"],
  ["Poonawalla Fincorp", "NBFC", "64.0", "not disclosed"], ["Repco Home Finance", "HFC", "74.0", "75.5"], ["SBI Cards & Payment", "NBFC (cards)", "68.5", "77.0"],
  ["Shriram Finance", "NBFC", "30.5", "25.6"]];
C.push(table(["Company", "Type", "Bank share in original dataset (%)", "Bank share verified from filing (%)"], ver, [3226, 1600, 2267, 2267]));
C.push(note("Definitions follow each company's borrowing-mix slide (e.g. some include financial institutions). Source: each company's Q2 FY2024 investor presentation (30 September 2023). Shriram's figure (term loans, 25.58%) appears in its Q2 FY2024 presentation and is repeated as the comparative column in its Q2 FY2025 presentation."));
C.push(caption("Table 4.3: Descriptive statistics of firm-level variables"));
C.push(table(["Variable", "Group", "Period", "N", "Mean", "SD", "Min", "Max"], FD.map((r) => [r.Variable, r.Group, r.Period, r.N, f2(r.Mean), f2(r.SD), f2(r.Min), f2(r.Max)]), [2126, 1500, 1500, 600, 900, 900, 900, 934]));
C.push(note("Company-quarters, Q1 FY2021 – Q4 FY2026. Pre = up to Q2 FY2024 (September 2023); Post = Q3 FY2024 onward."));
C.push(P(`Table 4.3 summarises the firm-level data. Average cost of funds rose for both groups after the policy — from ${f2(fdv("Cost","Non-HFC NBFCs","Pre").Mean)} to ${f2(fdv("Cost","Non-HFC NBFCs","Post").Mean)} per cent for non-HFC NBFCs and, by more, from ${f2(fdv("Cost","HFCs","Pre").Mean)} to ${f2(fdv("Cost","HFCs","Post").Mean)} per cent for the exempt HFCs — reflecting the general rise in interest rates. The bank share of borrowings rose for both groups: by about ${f1(fdv("Bank","HFCs","Post").Mean - fdv("Bank","HFCs","Pre").Mean)} percentage points for HFCs and ${f1(fdv("Bank","Non-HFC NBFCs","Post").Mean - fdv("Bank","Non-HFC NBFCs","Pre").Mean)} for non-HFC NBFCs. These raw comparisons are unbalanced (firms enter and leave the panel) and motivate the formal tests in Chapter 8.`));
C.push(Hd("4.4 Descriptive statistics and correlations of the RBI series", HeadingLevel.HEADING_2));
const SHORT = { "Credit cards": "Cards", "Other personal loans": "Oth. PL", "Housing": "Housing", "Vehicle loans": "Vehicle", "Education": "Educ.", "NBFCs ex-HFC": "NBFC ex-HFC", "HFCs": "HFCs", "Non-food credit": "Non-food" };
C.push(caption("Table 4.4: Descriptive statistics of monthly bank credit growth (per cent per month)"));
C.push(table(["Series", "Period", "N", "Mean", "SD", "Min", "Max"], R.rbi_desc.map((r) => [r.Series, r.Period, r.N, f2(r.Mean), f2(r.SD), f2(r.Min), f2(r.Max)]),
  [2660, 1600, 700, 1100, 1100, 1100, 1100]));
C.push(note("Monthly log growth × 100. Pre-policy: April 2021 – 17 November 2023; post-policy: 29 December 2023 – January 2025. Merger month excluded; merger-adjusted series used where available. Source: RBI DBIE Table 15."));
C.push(P("Monthly growth of the targeted series is volatile: the standard deviation of monthly credit-card growth is larger than its mean. Much of this volatility is seasonal (festive-season card spending, quarter-end lending to NBFCs), which is why the regression models explain only a small share of month-to-month variation even when the post-policy shift is large, and why Chapter 9 checks the results with calendar-month controls."));
C.push(caption("Table 4.5: Correlation matrix of monthly bank credit growth, April 2021 – January 2025"));
C.push(table(["", ...R.rbi_corr.cols.map((c) => SHORT[c])], R.rbi_corr.rows.map((row, i) => [SHORT[R.rbi_corr.cols[i]], ...row.map((v) => f2(v))]),
  [1360, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000]));
C.push(note("Pearson correlations of monthly log growth rates."));
C.push(P("Correlations between categories are low to moderate, so each series carries largely independent information. The control groups are therefore chosen on economic grounds (similar borrowers, exempt from the rule) rather than on statistical co-movement, and the tests compare average growth before and after the policy rather than relying on month-to-month co-movement."));

// ================= Chapter 5 (Theoretical framework)
C.push(Hd("5. Theoretical Framework", HeadingLevel.HEADING_1));
C.push(Hd("5.1 Conceptual framework", HeadingLevel.HEADING_2));
C.push(P("A risk weight determines how much capital a lender must hold against each rupee of a loan. When the risk weight on a category of loans rises, the lender either has to hold more capital (which is costly) or lend less in that category; if it has ample capital, the change may matter mainly as a signal of supervisory concern. The November 2023 circular raised risk weights on two links in the credit chain at once and, in the same circular, required board-approved limits on unsecured consumer credit. Table 5.1 traces each link from the regulatory change to the variable observed in the data and the model that tests it."));
C.push(caption("Table 5.1: Transmission channels, observable variables and models"));
C.push(table(["Channel", "Mechanism", "Observable variable", "Model"], [
  ["Consumer credit", "Higher capital per rupee of unsecured consumer loans and required exposure limits → banks slow or reprice such lending", "Monthly growth of bank credit to cards and other personal loans relative to exempt household credit", "Model 1 (H1)"],
  ["Bank funding of NBFCs", "Higher capital per rupee lent to NBFCs → banks slow lending to NBFCs (HFCs exempt)", "Monthly growth of bank credit to NBFCs ex-HFC relative to total non-food credit", "Model 2 (H2)"],
  ["NBFC funding cost", "Bank-dependent, non-exempt NBFCs face dearer or scarcer bank loans", "Quarterly cost of funds of each firm", "Model 3"],
  ["NBFC funding mix", "Non-exempt NBFCs substitute bonds, CP and other funding for bank loans", "Bank share of each firm's borrowings", "Model 4"],
  ["NBFC lending mix", "NBFCs slow affected consumer segments relative to exempt segments", "Segment-wise AUM growth of Bajaj Finance", "Model 5"],
], [1700, 3000, 3060, 1600]));
C.push(Hd("5.2 Variables", HeadingLevel.HEADING_2));
C.push(caption("Table 5.2: Definition of variables"));
C.push(table(["Variable", "Type", "Definition", "Source"], [
  ["Gap1(t)", "Dependent (Model 1)", "Average monthly log growth (×100) of bank credit to credit cards and other personal loans minus that of housing, vehicle and education loans", "RBI DBIE Table 15"],
  ["Gap2(t)", "Dependent (Model 2)", "Monthly log growth (×100) of bank credit to NBFCs excluding HFCs minus that of total non-food bank credit", "RBI DBIE Table 15"],
  ["g(i,t)", "Dependent (panel)", "Monthly log growth (×100) of bank credit to category i", "RBI DBIE Table 15"],
  ["COF(f,q)", "Dependent (Model 3)", "Cost of funds of firm f in quarter q, per cent", "Investor presentations"],
  ["BankShare(f,q)", "Dependent (Model 4)", "Bank borrowings as a share of total borrowings, per cent", "Investor presentations"],
  ["AUMgrowth(s,q)", "Dependent (Model 5)", "Quarterly log growth of Bajaj Finance segment AUM", "Investor presentations"],
  ["Post(t)", "Independent", "Monthly data: 1 from the observation of 29 December 2023 (the observation of 17 November 2023 is pre-policy); quarterly data: 1 from Q3 FY2024", "RBI circular"],
  ["Treated(i) × Post(t)", "Independent", "1 for targeted categories / firms / segments after the policy", "RBI circular"],
  ["BankShare2023(f) × NonHFC(f) × Post(q)", "Independent (Model 3)", "Verified bank share at 30 September 2023, for non-HFC NBFCs only, interacted with Post", "Q2 FY2024 presentations"],
  ["NonHFC(f) × Post(q)", "Independent (Models 3–4)", "1 for NBFCs other than HFCs after the policy", "RBI circular"],
  ["Fixed effects", "Control", "Category / firm / segment and month / quarter effects; calendar-month dummies in robustness tests", "—"],
], [2300, 1700, 3860, 1500]));
C.push(Hd("5.3 Model 1: consumer-credit channel (H1)", HeadingLevel.HEADING_2));
C.push(P([I("Gap1(t) = α₁ + δ₁ · Post(t) + ε(t)")], { align: AlignmentType.CENTER }));
C.push(P([B("Null hypothesis H0: "), "δ₁ = 0 — growth of targeted bank consumer credit did not change relative to exempt household credit after the circular."]));
C.push(P([B("Alternative hypothesis H1: "), "δ₁ < 0 — growth of targeted bank consumer credit fell relative to exempt household credit after the circular."]));
C.push(Hd("5.4 Model 2: bank-funding channel (H2)", HeadingLevel.HEADING_2));
C.push(P([I("Gap2(t) = α₂ + δ₂ · Post(t) + ε(t)")], { align: AlignmentType.CENTER }));
C.push(P([B("Null hypothesis H0: "), "δ₂ = 0 — growth of bank credit to NBFCs did not change relative to total bank credit after the circular."]));
C.push(P([B("Alternative hypothesis H1: "), "δ₂ < 0 — growth of bank credit to NBFCs (excluding HFCs) fell relative to total bank credit after the circular."]));
C.push(Hd("5.5 Supporting firm-level models (Models 3–5)", HeadingLevel.HEADING_2));
C.push(P([I("Model 3: COF(f,q) = μ(f) + λ(q) + β₁ · BankShare2023(f) × Post(q) + β₂ · NonHFC(f) × Post(q) + β₃ · BankShare2023(f) × NonHFC(f) × Post(q) + ε(f,q)")], { align: AlignmentType.CENTER }));
C.push(P([I("Model 4: BankShare(f,q) = μ(f) + λ(q) + β₄ · NonHFC(f) × Post(q) + ε(f,q)")], { align: AlignmentType.CENTER }));
C.push(P([I("Model 5: AUMgrowth(s,q) = μ(s) + λ(q) + β₅ · Affected(s) × Post(q) + ε(s,q)")], { align: AlignmentType.CENTER }));
C.push(P("In Model 3 the coefficient of interest is β₃: whether bank dependence raised the cost of funds more for NBFCs that faced the higher bank risk weight than for the exempt HFCs. β₂ alone is the simple exemption-based test."));
C.push(caption("Table 5.3: Summary of hypotheses"));
C.push(table(["Model", "H0", "H1 (alternative)", "Test", "Decision rule"], [
  ["1", "δ₁ = 0", "δ₁ < 0", "t-test, Newey–West SE (3 lags), t-distribution; checked against Table 9.5", "Reject H0 if p < 0.05"],
  ["2", "δ₂ = 0", "δ₂ < 0", "As Model 1", "Reject H0 if p < 0.05"],
  ["3", "β₃ = 0 (and β₂ = 0)", "β₃ > 0 (β₂ > 0)", "Wild cluster bootstrap-t", "Reject H0 if p < 0.05"],
  ["4", "β₄ = 0", "β₄ < 0", "Wild cluster bootstrap-t", "Reject H0 if p < 0.05"],
  ["5", "β₅ = 0", "β₅ < 0", "t-test, Driscoll–Kraay SE", "Reject H0 if p < 0.05"],
], [900, 1500, 1500, 3460, 2000]));
C.push(note("All p-values reported in this dissertation are two-sided, which is conservative relative to the one-sided alternatives."));

// ================= Chapter 6 (Methodology)
C.push(Hd("6. Econometric Methodology", HeadingLevel.HEADING_1));
C.push(Hd("6.1 Research design", HeadingLevel.HEADING_2));
C.push(P("The study uses a difference-in-differences design. For each hypothesis, credit that the regulation targeted (the treatment group) is compared with closely related credit that was exempt (the control group), before and after the policy. Any change in the growth gap between the two groups after November 2023 is attributed to the policy, under the assumption that, without the policy, the gap would have stayed at its pre-policy level. Interest-rate changes, economic growth and other common shocks affect both groups and cancel out in the comparison; shocks that affect only one group do not, and Chapter 10 discusses the most important of these."));
C.push(P(`The main sample runs from April 2021 to January 2025, so that the post-policy period ends before the February 2025 rollback was announced. Because RBI observations refer to the last reporting Friday of each month, the November 2023 observation is dated 17 November 2023 — one day after the circular — and is treated as pre-policy. The first post-policy observation is that of 29 December 2023. The main sample therefore has ${H1.its.n_pre} pre-policy and ${H1.its.n_post} post-policy months for Model 1 (${H2.its.n_pre} and ${H2.its.n_post} for Model 2, which loses the merger month).`));
C.push(Hd("6.2 Aggregate tests", HeadingLevel.HEADING_2));
C.push(P([B("Interrupted time series on the growth gap. "), "For each month t, the growth gap is the average monthly growth of the treated categories minus that of the control categories. The gap is regressed on a constant and a post-policy indicator:"]));
C.push(P([I("Gap(t) = α + δ · Post(t) + ε(t)")], { align: AlignmentType.CENTER }));
C.push(P("α is the average pre-policy gap and δ is the change in the gap after the policy — the difference-in-differences estimate. Standard errors are Newey–West with three lags, the conventional rule-of-thumb lag length for about 46 observations (floor of 4(T/100)^(2/9)), and p-values use the t-distribution with n − 2 degrees of freedom. The first draft of this study used six lags; because longer lags give smaller p-values in these data (Chapter 9), the rule-of-thumb choice is used as the main specification and all alternatives are reported."));
C.push(P([B("Panel difference-in-differences. "), "For H1, each category is also treated as a separate series in a panel, with category fixed effects, month fixed effects and a Treated × Post interaction, estimated with the linearmodels PanelOLS estimator and Driscoll–Kraay standard errors (bandwidth three):"]));
C.push(P([I("g(i,t) = μ(i) + λ(t) + δ · Treated(i) × Post(t) + ε(i,t)")], { align: AlignmentType.CENTER }));
C.push(P("With the same series and months, this panel estimator gives essentially the same point estimate as the interrupted time series on the averaged gap; it differs only in how the standard error is computed. It is reported as an alternative calculation, not as independent confirmation."));
C.push(P([B("Placebo tests. "), "Both tests are repeated using only pre-policy data with a fictitious policy date of November 2022. The placebo window starts in July 2021, after the April–June 2021 COVID-19 wave, which produced an extreme negative gap that would otherwise dominate the placebo (results starting in April 2021 are shown for comparison)."]));
C.push(P([B("Pre-trend and event-study tests. "), "The pre-policy gap is regressed on a linear time trend, and an event study replaces the single post-policy indicator with indicators for three-month blocks before and after the circular (Chapter 9)."]));
C.push(Hd("6.3 Treatment and control groups", HeadingLevel.HEADING_2));
C.push(caption("Table 6.1: Treatment and control groups"));
C.push(table(["Hypothesis", "Treated (higher risk weight)", "Control (exempt)", "Expected sign of δ"], [
  ["H1 consumer credit", "Credit cards; other personal loans", "Housing; vehicle loans; education", "Negative"],
  ["H2 bank funding", "Bank credit to NBFCs excluding HFCs", "Total non-food bank credit", "Negative"],
  ["H2 alternatives", "Bank credit to NBFCs ex-HFC / all NBFCs", "Non-food credit excluding NBFCs, cards and other personal loans / bank credit to HFCs / total non-food credit", "Negative"],
], [1900, 2600, 3034, 1826]));
C.push(Hd("6.4 Identifying assumptions", HeadingLevel.HEADING_2));
C.push(P("The key assumption is that, without the policy, the growth gap between treated and control credit would have stayed at its pre-policy average. The placebo, pre-trend and event-study tests check whether the gap was already moving before the policy. The interest-rate cycle and economic growth affect treated and control credit alike and are differenced out; the repo rate was unchanged at 6.5 per cent throughout the main post-policy window. If banks redirected lending from targeted to exempt categories, the control group would grow faster and the estimate would overstate the absolute fall in targeted credit, though it would still measure the change in the relative allocation of credit. The HDFC merger is handled with the RBI's merger-adjusted series (Section 4.1). Threats that the design cannot remove — the RBI acting because targeted credit was growing fast, the bundled nature of the circular, and rising delinquencies in the treated categories — are discussed in Section 10.2."));
C.push(Hd("6.5 Firm-level tests", HeadingLevel.HEADING_2));
C.push(P("Firm-level panels are estimated with two-way fixed effects (company and quarter) using PanelOLS. With 16 firms, inference uses the wild cluster bootstrap-t clustered by company (Webb six-point weights, 19,999 draws, null imposed). Model 3 regresses cost of funds on the interactions of Post with verified pre-policy bank share and with a non-HFC indicator, and on their triple interaction, so that bank dependence counts as exposure only for NBFCs that faced the higher risk weight. Model 4 regresses the bank share of borrowings on Post × non-HFC, with subsamples excluding firms whose bank loans were largely priority-sector-eligible (NBFC-MFIs) or that were under RBI business restrictions. Model 5 compares the growth of Bajaj Finance's affected and exempt segments, with segment and quarter fixed effects."));
C.push(Hd("6.6 How significance is described", HeadingLevel.HEADING_2));
C.push(P("Because each aggregate test rests on 45–46 monthly observations, p-values depend on the inference method. Chapter 9 reports every main estimate under classical, heteroskedasticity-robust, Newey–West (1 to 12 lags), fixed-b, seasonal-dummy, quarterly-average and moving-block-bootstrap inference. A result is described as significant only when it is significant at 5 per cent in the main specification and robust across most of these methods; otherwise it is described as marginal or as not robust."));

// ================= Chapter 7 (H1 results)
C.push(Hd("7. Empirical Findings and Results: H1 — Consumer-Credit Channel", HeadingLevel.HEADING_1));
C.push(...fig("fig1_rbi_indices.png", 600, 234, "Figure 7.1: Bank credit indices, targeted versus exempt categories (October 2023 = 100)",
  "Source: RBI DBIE Table 15; author's calculations. Indices chain monthly growth; merger-adjusted series used where available, merger month set to zero growth."));
C.push(P("Figure 7.1 (left panel) shows the raw data. Before November 2023, bank credit to cards and personal loans was growing faster than exempt household credit. After the circular the targeted series grew more slowly than the exempt series."));
C.push(caption("Table 7.1: H1 — change in targeted bank consumer credit growth relative to exempt household credit"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "Annualised (pp)", "N"], [
  ["Interrupted time series (Newey–West, 3 lags)", f2(H1.its.change), f2(H1.its.se), p3(H1.its.p), f1(H1.its.annualised_change), H1.its.n],
  ["Same comparison as a panel (Driscoll–Kraay)", f2(H1.panel.beta), f2(H1.panel.se), p3(H1.panel.p), f1(H1.panel.annualised), H1.panel.n],
  ["Placebo Nov-2022, time series (Jul-21 to Nov-23)", f2(H1.placebo.change), f2(H1.placebo.se), p3(H1.placebo.p), f1(H1.placebo.annualised_change), H1.placebo.n],
  ["Placebo Nov-2022, panel (Jul-21 to Nov-23)", f2(H1.placebo_panel.beta), f2(H1.placebo_panel.se), p3(H1.placebo_panel.p), f1(H1.placebo_panel.annualised), H1.placebo_panel.n],
  ["Placebo Nov-2022, time series (Apr-21 to Nov-23)", f2(H1.placebo_from_apr21.change), f2(H1.placebo_from_apr21.se), p3(H1.placebo_from_apr21.p), f1(H1.placebo_from_apr21.annualised_change), H1.placebo_from_apr21.n],
], [3560, 1300, 1100, 1100, 1300, 1000]));
C.push(note("Sample April 2021 – January 2025 (placebo: pre-policy observations only). δ is the change in the monthly growth gap between treated and control categories after the policy date. p-values from the t-distribution. The panel uses the same series and months as the time series and is not independent evidence."));
C.push(P(`After the circular, monthly growth of the targeted categories fell by ${f2(-H1.its.change)} percentage points relative to exempt categories (p = ${p3(H1.its.p)}), or about ${f1(-h1a)} percentage points a year — large compared with pre-policy growth rates of 20–23 per cent. The panel calculation gives the same point estimate (p = ${p3(H1.panel.p)}). The placebo starting in July 2021 shows no break at the fictitious November 2022 date (${f2(H1.placebo.change)}, p = ${p3(H1.placebo.p)}). Starting the placebo in April 2021 gives a positive but insignificant estimate (${f2(H1.placebo_from_apr21.change)}, p = ${p3(H1.placebo_from_apr21.p)}); this is driven by the April–June 2021 COVID-19 wave, when the gap was strongly negative, and does not indicate that targeted credit was accelerating before the policy. Excluding that quarter, the quarterly average gap was ${f1(A.gap_quarter.H1["2021Q3"])} and ${f1(A.gap_quarter.H1["2021Q4"])} in the last two quarters of 2021, between ${f1(Math.min(...["2022Q1","2022Q2","2022Q3","2022Q4"].map((q) => A.gap_quarter.H1[q])))} and ${f1(Math.max(...["2022Q1","2022Q2","2022Q3","2022Q4"].map((q) => A.gap_quarter.H1[q])))} in 2022, ${f1(A.gap_quarter.H1["2023Q1"])} in early 2023 and ${f1(A.gap_quarter.H1["2023Q2"])} and ${f1(A.gap_quarter.H1["2023Q3"])} in the two quarters before the policy, so the gap was already narrowing somewhat. Section 9.4 tests this pre-trend and shows how much it matters. The significance of the main estimate depends on the inference method (Section 9.3).`));
C.push(...fig("fig2_h1_gap.png", 560, 238, "Figure 7.2: Monthly growth gap, targeted minus exempt consumer credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters. 2023Q4 (October, 17 November and 29 December observations) is shown grey. Dashed line: pre-policy average gap."));
C.push(Hd("7.1 Robustness", HeadingLevel.HEADING_2));
C.push(caption("Table 7.2: H1 robustness (panel DiD, Driscoll–Kraay standard errors)"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "N"],
  H1.robustness.map((r) => [r.Specification, f2(r.Coefficient), f2(r.SE), p3(r.p), r.N]), [4360, 1300, 1200, 1200, 1300]));
const rb = (l) => H1.robustness.find((r) => r.Specification.startsWith(l));
C.push(P(`The estimate is negative in every specification. It remains significant at the 5 per cent level when credit cards, the housing series or vehicle loans are dropped, and with shorter pre- or post-periods. Dropping education from the control group gives p = ${p3(rb("Drop education").p)} (significant only at 10 per cent), and dropping other personal loans — leaving only credit cards as treated — gives p = ${p3(rb("Drop other personal loans").p)}, which is not significant. The result therefore rests mainly on other personal loans, the larger of the two treated series.`));
C.push(P(`An earlier draft also added bank gold loans to the control group. That is not a valid check: bank gold loans grew very rapidly after 2024 for reasons unrelated to the circular (higher gold prices and the reclassification of agricultural gold loans as retail gold loans), so adding them mechanically enlarges the estimated gap (δ = ${f2(H1.gold_note.beta)}). It is not used as evidence.`));
C.push(Hd("7.2 Economic magnitude", HeadingLevel.HEADING_2));
C.push(P(`Suppose targeted consumer credit had continued to grow at its pre-policy pace relative to exempt credit. Over the ${MG.months} post-policy months from December 2023 to January 2025, a monthly shortfall of ${f2(-H1.its.change)} percentage points accumulates to about ${pct(H1.its.change, MG.months)} per cent. With ₹${lakh(MG.h1_outstanding_jan25)} lakh crore of bank credit outstanding to cards and other personal loans in January 2025, this implies roughly ₹${lakh(MG.h1_shortfall)} lakh crore less targeted credit than would otherwise have been extended. This is an upper-end figure: it assumes the pre-policy relative growth would have continued, whereas booms tend to slow on their own (Section 10.2). If the mild pre-policy decline in the gap since July 2021 is projected forward instead, the monthly effect falls to ${f2(-TC.H1.from_jul21.delta)} percentage points and the implied shortfall to about ₹${lakh(MG.h1_outstanding_jan25 * (Math.exp(-TC.H1.from_jul21.delta * MG.months / 100) - 1))} lakh crore.`));

// ================= Chapter 8 (H2 and firm-level)
C.push(Hd("8. Empirical Findings and Results: H2 — Bank-Funding Channel and Firm-Level Evidence", HeadingLevel.HEADING_1));
C.push(Hd("8.1 Aggregate evidence", HeadingLevel.HEADING_2));
C.push(P("The right panel of Figure 7.1 compares bank credit to NBFCs (excluding HFCs) with total non-food bank credit. Before the policy, bank credit to NBFCs grew much faster than overall bank credit. After November 2023 it grew more slowly, and it accelerated again after the April 2025 rollback."));
C.push(caption("Table 8.1: H2 — change in bank credit growth to NBFCs relative to benchmark"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "Annualised (pp)", "N"], [
  ["Main: NBFCs ex-HFC minus total non-food credit", f2(H2.its.change), f2(H2.its.se), p3(H2.its.p), f1(H2.its.annualised_change), H2.its.n],
  ["Benchmark excl. NBFCs, cards and other personal loans", f2(H2.alt_core.change), f2(H2.alt_core.se), p3(H2.alt_core.p), f1(H2.alt_core.annualised_change), H2.alt_core.n],
  ["Placebo Nov-2022 (Jul-21 to Nov-23)", f2(H2.placebo.change), f2(H2.placebo.se), p3(H2.placebo.p), f1(H2.placebo.annualised_change), H2.placebo.n],
  ["Placebo Nov-2022 (Apr-21 to Nov-23)", f2(H2.placebo_from_apr21.change), f2(H2.placebo_from_apr21.se), p3(H2.placebo_from_apr21.p), f1(H2.placebo_from_apr21.annualised_change), H2.placebo_from_apr21.n],
  ["Alternative: NBFCs ex-HFC minus HFCs", f2(H2.alt_vs_hfc.change), f2(H2.alt_vs_hfc.se), p3(H2.alt_vs_hfc.p), f1(H2.alt_vs_hfc.annualised_change), H2.alt_vs_hfc.n],
  ["Alternative: all NBFCs (incl. HFCs) minus non-food", f2(H2.alt_all_nbfc.change), f2(H2.alt_all_nbfc.se), p3(H2.alt_all_nbfc.p), f1(H2.alt_all_nbfc.annualised_change), H2.alt_all_nbfc.n],
  ["Rollback: change after Apr-2025 (Dec-23 to Jul-26)", f2(H2.rollback.change), f2(H2.rollback.se), p3(H2.rollback.p), f1(H2.rollback.annualised_change), H2.rollback.n],
], [3560, 1250, 1050, 1050, 1250, 1200]));
C.push(note("Interrupted time series on the monthly growth gap; Newey–West standard errors (3 lags) and t-distribution p-values. Main sample April 2021 – January 2025."));
C.push(P(`After the circular, monthly growth of bank credit to NBFCs (excluding HFCs) fell by ${f2(-H2.its.change)} percentage points relative to total non-food credit, about ${f1(-h2a)} percentage points a year. In the main specification the estimate is significant only at the 10 per cent level (p = ${p3(H2.its.p)}). Total non-food credit includes bank credit to NBFCs itself and the H1-treated consumer categories; against a benchmark that excludes them the estimate is slightly larger (${f2(H2.alt_core.change)}, p = ${p3(H2.alt_core.p)}). The placebo shows no break at November 2022 (p = ${p3(H2.placebo.p)}). Chapter 9 shows that the estimate is significant at 5 per cent once calendar-month seasonality (quarter-end lending surges) is controlled for, but not under several conservative methods without seasonal controls. The H2 result is therefore best described as marginal.`));
C.push(P(`Taken at face value, a monthly shortfall of ${f2(-H2.its.change)} percentage points over ${MG.months} months cumulates to about ${pct(H2.its.change, MG.months)} per cent. With ₹${lakh(MG.h2_outstanding_jan25)} lakh crore of bank credit to NBFCs (excluding HFCs) outstanding in January 2025, this implies roughly ₹${lakh(MG.h2_shortfall)} lakh crore less bank credit to NBFCs than if the pre-policy relationship had continued. The figure is approximate and inherits the uncertainty of the estimate.`));
C.push(P(`Two alternative comparisons point in the same direction but are less precise. Against bank credit to the exempt HFCs the slowdown is larger (${f1(-H2.alt_vs_hfc.annualised_change)} percentage points a year) but noisier (p = ${p3(H2.alt_vs_hfc.p)}), because the HFC series is small and was disrupted by the HDFC merger. Including HFCs in the treated group dilutes the effect (p = ${p3(H2.alt_all_nbfc.p)}). After the April 2025 rollback the gap moved back in favour of NBFCs by ${f2(H2.rollback.change)} percentage points a month (p = ${p3(H2.rollback.p)}); the rebound is similar in size to the original slowdown but not statistically significant.`));
C.push(P("A fall in bank lending to NBFCs is consistent with lower bank supply, but it is equally consistent with lower NBFC demand for bank funds, since NBFCs' own consumer and microfinance lending also slowed in 2024. The firm-level cost-of-funds tests below, which would show a price response if supply had tightened, are insignificant. Distinguishing supply from demand would require data on NBFCs' bond and commercial-paper issuance (whether market borrowing rose as bank borrowing fell), which this study did not collect."));
C.push(...fig("fig3_h2_gap.png", 560, 238, "Figure 8.1: Monthly growth gap, bank credit to NBFCs (ex-HFC) minus total non-food credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters (including after the April 2025 rollback)."));
C.push(Hd("8.2 How did listed NBFCs adjust? Firm-level evidence", HeadingLevel.HEADING_2));
C.push(P([B("Company disclosures. "), "Two consumer lenders quantified the immediate capital impact of the rule in their December 2023 presentations, and one disclosed a pricing response (Table 8.2)."]));
C.push(caption("Table 8.2: Company statements on the effect of the November 2023 measures"));
C.push(table(["Company", "Statement (verbatim)", "Implication"], [
  ["Bajaj Finance (Q3 FY24)", "“RBI increased risk weights on consumer credit exposure from 100% to 125% which had an impact of 290 bps on the Company's CRAR. Adjusted for this change CRAR would have been 26.77%.”", "Capital ratio −2.9 pp"],
  ["Bajaj Finance (Q3 FY24)", "“Given the increase in risk weights and higher incremental cost of funds, the Company has increased rates across all portfolios by 20–30 bps.”", "Lending rates +0.20–0.30 pp (attributed to both risk weights and funding costs)"],
  ["SBI Cards (Q3 FY24)", "“CRAR impacted by ~400 bps due to increase in risk weight by RBI.”", "Capital ratio about −4 pp"],
], [2000, 5160, 2200]));
C.push(note("Source: company investor presentations for the quarter ended 31 December 2023."));
C.push(P("These statements show that the higher risk weights bit on NBFCs' own capital ratios. This is a separate channel from H1, which measures bank credit; both companies are NBFCs. The Bajaj rate increase is attributed by the company to both the risk weights and higher incremental funding costs, so it cannot be assigned to the policy alone."));
C.push(P([B("Cost of funds (Model 3). "), `Using ${fc.n} company-quarters for ${fc.firms} firms, the original specification — cost of funds on Post × pre-policy bank share for all firms — gave a positive but insignificant coefficient (${f2(fc.beta * 30)} percentage points for 30 points more bank funding, p = ${p3(fc.p)}). That specification treats the exempt HFCs as exposed. In the triple-difference specification, where bank dependence counts as exposure only for non-HFC NBFCs, the extra effect for non-HFC NBFCs is negative (${f2(A.cof_triple.beta * 30)} pp per 30 points, p = ${p3(A.cof_triple.p)}), and the bank-share gradient comes mainly from the exempt HFCs (${f2(A.cof_triple.beta_exp * 30)} pp per 30 points, p = ${p3(A.cof_triple.p_exp)}). The simple exemption-based test also points the wrong way: relative to HFCs, non-HFC NBFCs' cost of funds rose by ${f2(fch.beta)} percentage points (p = ${p3(fch.p)}), consistent with the raw data in Table 4.3, where HFCs' cost of funds rose more. None of these is significant at 5 per cent, and none supports a policy effect on funding costs.`]));
C.push(...fig("fig4_firm_cof_event.png", 560, 245, "Figure 8.2: Cost of funds and pre-policy bank share, by half-year (all firms)",
  "Coefficients on bank share × half-year, scaled to a 30 pp difference in bank share; reference period FY23 H2; 95% wild-cluster-bootstrap intervals."));
C.push(P("Figure 8.2 shows no clear break at the policy date. The coefficients before the policy (about 0.33 to 0.40 percentage points in FY21–FY22, falling to 0.07–0.13 in FY23) are as large as those after it (0.27 to 0.50). The apparent rise after the policy exists only because the reference period, FY23 H2, is the trough of a V-shape. That V is what the repricing lag of bank loans predicts: market borrowing repriced first when rates rose in 2022–23, so bank-dependent lenders' costs rose more slowly at first, and bank loans linked to the marginal cost of funds based lending rate (MCLR) caught up in FY2024."));
C.push(P([B("Bank share of borrowings (Model 4). "), `Relative to HFCs, the bank share of non-HFC NBFCs' borrowings fell by ${f1(-fb.beta)} percentage points before the rollback (p = ${p3(fb.p)}). The placebo (fictitious November 2022 date) is ${f2(R.firm_bankshare_placebo.beta)} points (p = ${p3(R.firm_bankshare_placebo.p)}), about half the main estimate, so part of the difference was already present before the policy. The estimate is stable when IIFL Finance (under an RBI embargo on its gold-loan business from March to September 2024) is excluded (${f2(sub("Excluding IIFL Finance").beta)}, p = ${p3(sub("Excluding IIFL Finance").p)}), and when the NBFC-MFIs CreditAccess Grameen and Arman Financial — whose bank borrowings were largely priority-sector-eligible and hence exempt — are excluded (${f2(sub("Excluding NBFC-MFIs").beta)}, p = ${p3(sub("Excluding NBFC-MFIs").p)}). Muthoot Finance's bank borrowing may also be partly priority-sector-eligible [VERIFY: share of Muthoot's bank borrowings classified as priority-sector on-lending is not disclosed in its presentations]; excluding it as well gives ${f2(sub("Excluding IIFL, NBFC-MFIs and Muthoot").beta)} (p = ${p3(sub("Excluding IIFL, NBFC-MFIs and Muthoot").p)}).`]));
C.push(P(`The raw pattern does not show non-HFC NBFCs moving away from banks. Their average bank share rose slightly after the policy (Table 4.3), and in a balanced panel of ${A.balanced_firms.length} firms observed in every half-year both groups' bank shares rose in parallel before the policy (to ${f1(A.bankshare_halfyear_bal.FY23H2.nonhfc)} and ${f1(A.bankshare_halfyear_bal.FY23H2.hfc)} per cent in FY23 H2), after which HFCs' share kept rising (${f1(A.bankshare_halfyear_bal.FY25H2.hfc)} per cent in FY25 H2) while non-HFC NBFCs' share levelled off (${f1(A.bankshare_halfyear_bal.FY25H2.nonhfc)} per cent). In the unbalanced raw averages, the jump in HFCs' bank share from ${f1(A.bankshare_halfyear_raw.FY23H1.hfc)} to ${f1(A.bankshare_halfyear_raw.FY23H2.hfc)} per cent in FY23 H2 occurred before the policy. Across firms the picture is mixed: SBI Cards, the most directly targeted firm, raised its bank share from ${f1(chgF("SBI Cards and Payment Services Ltd.").pre)} to ${f1(chgF("SBI Cards and Payment Services Ltd.").post)} per cent, as did Mahindra & Mahindra Financial and Muthoot, while Bajaj Finance, Cholamandalam and Shriram reduced theirs. The difference between the groups therefore reflects HFCs moving toward banks at least as much as NBFCs moving away, and aggregate bank credit to HFCs did not accelerate after the policy (Table 4.1), so a reallocation of bank lending from NBFCs to HFCs is not evident either.`));
C.push(...fig("fig5_bank_share.png", 560, 252, "Figure 8.3: Average bank share of borrowings, non-HFC NBFCs versus HFCs",
  "Extracted from borrowing-mix slides of quarterly investor presentations; definitions follow each company's disclosure. Solid lines: all firms (unbalanced: Shriram enters in FY23, AAVAS is missing in FY26, IIFL is annual only). Dashed lines: balanced panel of firms observed in every half-year FY22–FY25."));
C.push(P([B("Within-firm lending: Bajaj Finance (Model 5). "), `Bajaj Finance is the largest NBFC consumer lender in the sample and reports assets by business segment every quarter. Table 8.3 shows how each segment is classified against the circular's definition of consumer credit. Comparing the growth of affected and exempt segments within the firm, there is no significant relative slowdown (${f2(bw.beta)} pp per quarter, p = ${p3(bw.p)}). This test is confounded: on 15 November 2023, one day before the circular, the RBI barred Bajaj Finance from sanctioning and disbursing loans under its “eCOM” and “Insta EMI Card” products, and lifted the restriction on 2 May 2024. Bajaj's affected consumer share dipped from ${A.bajaj_share_series.FY24Q2} per cent in September 2023 to ${A.bajaj_share_series.FY24Q4} per cent in March 2024 and recovered to ${A.bajaj_share_series.FY25Q4} per cent by March 2025, tracking the embargo. Excluding the embargo quarters (Q3 FY24 to Q1 FY25) leaves the result insignificant (${f2(A.bajaj_ex_embargo.beta)}, p = ${p3(A.bajaj_ex_embargo.p)}). The Bajaj data therefore cannot show whether the risk weight affected NBFC consumer lending.`]));
C.push(caption("Table 8.3: Classification of Bajaj Finance segments"));
C.push(table(["Segment", "Classification", "Reason"], [
  ["Urban Sales Finance", "Affected", "Consumer-durable and lifestyle EMI loans to individuals (consumer credit); includes the embargoed eCOM / Insta EMI Card products"],
  ["Urban B2C", "Affected", "Unsecured personal loans to individuals"],
  ["Rural Sales Finance", "Affected", "Consumer-durable loans to individuals"],
  ["Rural B2C", "Affected (mixed)", "Personal loans; includes gold loans (exempt), which Bajaj reports separately only from Q4 FY24 and which are added back for a consistent series"],
  ["Two & Three-wheeler Finance", "Exempt", "Vehicle loans"],
  ["Mortgages", "Exempt", "Housing loans and loans to developers/businesses, not consumer credit"],
  ["SME Lending", "Exempt", "Business loans; car loans (vehicle loans), reported separately from Q4 FY24, are included"],
  ["Commercial Lending", "Exempt", "Loans to businesses"],
  ["Loan Against Securities", "Exempt", "Capital-market exposure, risk-weighted under separate rules [VERIFY: whether loans against securities to individuals fall within the circular's consumer-credit definition]"],
], [2400, 1600, 5360]));
C.push(note("Segments reported for at least 18 quarters are used; later segments (e.g. MFI lending, CV & tractor finance) are excluded. Source: Bajaj Finance investor presentations."));
C.push(caption("Table 8.4: Firm-level tests"));
C.push(table(["Test", "Estimate", "p-value", "Obs.", "Firms"], [
  ["Cost of funds ~ Post × bank share (all firms)", `${f2(fc.beta * 30)} pp per 30 pp share`, p3(fc.p), fc.n, fc.firms],
  ["Cost of funds ~ Post × bank share × non-HFC (triple difference)", `${f2(A.cof_triple.beta * 30)} pp per 30 pp share`, p3(A.cof_triple.p), A.cof_triple.n, A.cof_triple.firms],
  ["Cost of funds ~ Post × non-HFC", `${f2(fch.beta)} pp`, p3(fch.p), fch.n, fch.firms],
  ["Bank share ~ Post × non-HFC (to Mar-2025)", `${f2(fb.beta)} pp`, p3(fb.p), fb.n, fb.firms],
  ["  excluding IIFL and NBFC-MFIs", `${f2(sub("Excluding IIFL and NBFC-MFIs").beta)} pp`, p3(sub("Excluding IIFL and NBFC-MFIs").p), sub("Excluding IIFL and NBFC-MFIs").n, sub("Excluding IIFL and NBFC-MFIs").firms],
  ["Bank share, placebo Nov-2022", `${f2(R.firm_bankshare_placebo.beta)} pp`, p3(R.firm_bankshare_placebo.p), "—", fb.firms],
  ["Bajaj: affected vs exempt segment growth", `${f2(bw.beta)} pp per quarter`, p3(bw.p), bw.n, `${bw.segments} segments`],
  ["  excluding embargo quarters", `${f2(A.bajaj_ex_embargo.beta)} pp per quarter`, p3(A.bajaj_ex_embargo.p), A.bajaj_ex_embargo.n, `${bw.segments} segments`],
], [4060, 2200, 1100, 1000, 1000]));
C.push(note("Two-way fixed effects (firm/segment and quarter). Firm tests: wild cluster bootstrap-t by firm. Bajaj: Driscoll–Kraay standard errors."));
C.push(P(`In short, the firm-level evidence on how NBFCs adjusted is descriptive and inconclusive. The disclosures confirm that the risk weights cut consumer NBFCs' capital ratios, but none of the cost-of-funds, borrowing-mix or segment-growth tests shows a significant policy effect, and the strongest-looking pattern (Model 4) owes as much to HFCs moving toward banks as to NBFCs moving away. Excluding SBI Cards from Model 4 enlarges the estimate (${f2(sub("Excluding SBI Cards").beta)}, p = ${p3(sub("Excluding SBI Cards").p)}), but this is a post-hoc exclusion of the most directly targeted firm and is not relied on.`));

// ================= Chapter 9 (Diagnostics, inference, identification)
C.push(Hd("9. Model Diagnostics, Inference and Identification Checks", HeadingLevel.HEADING_1));
C.push(Hd("9.1 Model summary, ANOVA and model equations", HeadingLevel.HEADING_2));
C.push(caption("Table 9.1: Model summary"));
C.push(table(["Model", "R", "R²", "Adjusted R²", "Std. error of estimate", "Durbin–Watson"], [MS1, MS2].map((M, k) =>
  [String(k + 1), f2(Math.sqrt(M.r2)), f2(M.r2), f2(M.adj_r2), f2(Math.sqrt(M.ss_resid / M.df_resid)), f2(M.dw)]), [1200, 1300, 1300, 1600, 2160, 1800]));
C.push(note("Predictor: (Constant), Post. Dependent variable: Model 1, Gap1; Model 2, Gap2."));
C.push(P([B("ANOVA. Null hypothesis: "), "Post has no effect on the growth gap (δ = 0). ", B("Alternative hypothesis: "), "δ ≠ 0."]));
C.push(caption("Table 9.2: ANOVA"));
const ms = (x, d) => f2(Number(x) / Number(d));
C.push(table(["Model", "Source", "Sum of squares", "df", "Mean square", "F", "Sig."], [MS1, MS2].flatMap((M, k) => [
  [String(k + 1), "Regression", f2(M.ss_model), M.df_model, ms(M.ss_model, M.df_model), f2(M.f), p3(M.f_p)],
  ["", "Residual", f2(M.ss_resid), M.df_resid, ms(M.ss_resid, M.df_resid), "", ""],
  ["", "Total", f2(M.ss_model + M.ss_resid), M.df_model + M.df_resid, "", "", ""]]), [900, 1400, 1660, 800, 1600, 1400, 1600]));
C.push(note("Classical OLS F-test, which assumes serially uncorrelated, homoskedastic errors."));
C.push(caption("Table 9.3: Coefficients"));
C.push(table(["Model", "", "B", "Std. error (Newey–West, 3 lags)", "t", "Sig."], [MS1, MS2].flatMap((M, k) => [
  [String(k + 1), "(Constant)", f2(M.const), f2(M.const_se), f2(M.const / M.const_se), p3(M.const_p)],
  ["", "Post", f2(M.delta), f2(M.delta_se), f2(M.delta / M.delta_se), p3(M.delta_p)]]), [900, 1400, 1300, 2560, 1400, 1800]));
C.push(P([B("Model equations: "), I(`Gap1(t) = ${f2(MS1.const)} − ${f2(-MS1.delta)} · Post(t);   Gap2(t) = ${f2(MS2.const)} − ${f2(-MS2.delta)} · Post(t)`)]));
C.push(P(`Both models explain only about 5 per cent of month-to-month variation (R² = ${f2(MS1.r2)} and ${f2(MS2.r2)}), and the classical F-tests do not reject the null at conventional levels (p = ${p3(MS1.f_p)} and ${p3(MS2.f_p)}). This is expected for monthly growth data dominated by short-term and seasonal noise that the policy dummy is not meant to explain; the question is whether the average gap shifted. The classical F-test assumes serially independent errors, which the diagnostics below do not fully support for Model 1. With Newey–West standard errors the post-policy coefficient is significant for Model 1 (p = ${p3(MS1.delta_p)}) and significant only at 10 per cent for Model 2 (p = ${p3(MS2.delta_p)}).`));
C.push(Hd("9.2 Diagnostic tests", HeadingLevel.HEADING_2));
C.push(caption("Table 9.4: Diagnostic tests on the residuals of Models 1 and 2"));
C.push(table(["Test", "Null hypothesis", "Model 1 statistic (p)", "Model 2 statistic (p)"], [
  ["Jarque–Bera (normality)", "Residuals are normally distributed", `${f2(MS1.jb)} (${p3(MS1.jb_p)})`, `${f2(MS2.jb)} (${p3(MS2.jb_p)})`],
  ["Skewness / kurtosis", "—", `${f2(MS1.skew)} / ${f2(MS1.kurt)}`, `${f2(MS2.skew)} / ${f2(MS2.kurt)}`],
  ["Durbin–Watson", "No first-order autocorrelation (≈ 2)", f2(MS1.dw), f2(MS2.dw)],
  ["Ljung–Box Q (12 lags)", "No autocorrelation up to lag 12", `${f2(MS1.lb12)} (${p3(MS1.lb12_p)})`, `${f2(MS2.lb12)} (${p3(MS2.lb12_p)})`],
  ["Breusch–Pagan", "Homoskedastic residuals", `${f2(MS1.bp)} (${p3(MS1.bp_p)})`, `${f2(MS2.bp)} (${p3(MS2.bp_p)})`],
  ["VIF (Treated × Post, H1 panel with fixed effects)", "No multicollinearity (VIF < 10)", f2(R.vif_h1_panel), "—"],
], [2900, 2660, 1900, 1900]));
C.push(P(`The Jarque–Bera test does not reject normality (p = ${p3(MS1.jb_p)} and ${p3(MS2.jb_p)}), but with 45–46 observations it has little power, so this is weak evidence. The Durbin–Watson statistics are close to 2, indicating no first-order autocorrelation. For Model 1, the Ljung–Box test points to autocorrelation at longer lags (p = ${p3(MS1.lb12_p)}, driven by a negative correlation at lag 2) and the Breusch–Pagan test to some heteroskedasticity (p = ${p3(MS1.bp_p)}), both at the 10 per cent level. Model 2 shows neither. Multicollinearity is not a concern in the panel specification. The negative lag-2 autocorrelation in Model 1 has a consequence: Newey–West standard errors shrink as the number of lags rises, so the choice of lag length matters, which is why the rule-of-thumb lag length is used and the alternatives are reported.`));
C.push(Hd("9.3 Sensitivity of inference to the choice of standard errors", HeadingLevel.HEADING_2));
C.push(P("Table 9.5 re-estimates the post-policy shift for both models under alternative methods: classical OLS; White heteroskedasticity-robust errors; Newey–West errors with 1, 3 (main), 6 and 12 lags; fixed-b inference (Kiefer–Vogelsang), which corrects the size distortion of Newey–West tests in short samples; the same model with calendar-month dummies to remove seasonality; quarterly averages of the monthly gap; and a moving-block bootstrap (blocks of three months, 10,000 replications) that does not rely on asymptotic formulas. All parametric p-values use the t-distribution."));
C.push(caption("Table 9.5: Post-policy coefficient and p-value under alternative inference methods"));
C.push(table(["Method", "Model 1 δ", "Model 1 p", "Model 2 δ", "Model 2 p"], I1.map((r, i) => [r.Method, f2(r.delta), p3(r.p), f2(I2[i].delta), p3(I2[i].p)]),
  [4560, 1200, 1200, 1200, 1200]));
C.push(note("Sample April 2021 – January 2025. Month dummies: eleven calendar-month indicators. Fixed-b: Bartlett kernel with bandwidth equal to the sample size; null distribution simulated with the actual design. Quarterly averages: calendar quarters, excluding the transition quarter 2023Q4 and the partial quarter 2025Q1. Block bootstrap: residuals resampled in blocks of three months under the null of no effect."));
C.push(P(`Three conclusions follow. First, the sign and approximate size of the estimates are stable: every method gives a negative post-policy shift of about ${f2(-im(I1, "Quarterly").delta)}–${f2(-im(I1, "Month dummies, classical").delta)} percentage points a month for Model 1 and ${f2(-H2.its.change)}–${f2(-im(I2, "Month dummies, classical").delta)} for Model 2. Second, Model 1 is significant at 5 per cent with Newey–West errors of three or more lags, with fixed-b inference (p = ${p3(im(I1, "Fixed-b HAC").p)}) and with the seasonal block bootstrap (p = ${p3(im(I1, "Month dummies, moving-block").p)}); at about 6–7 per cent with White errors and the plain block bootstrap (p = ${p3(im(I1, "White").p)} and ${p3(im(I1, "Moving-block bootstrap (block 3)").p)}); and not significant with classical OLS (p = ${p3(im(I1, "Classical").p)}) or quarterly averages (p = ${p3(im(I1, "Quarterly").p)}). Third, Model 2 is significant at 5 per cent only when calendar-month seasonality is controlled for (p = ${p3(im(I2, "Month dummies, classical").p)} to ${p3(im(I2, "Month dummies, Newey").p)}; ${p3(im(I2, "Month dummies, fixed-b").p)} with fixed-b); without seasonal controls it ranges from p = ${p3(H2.its.p)} (main) to ${p3(im(I2, "Fixed-b HAC").p)} (fixed-b) and ${p3(im(I2, "Quarterly").p)} (quarterly averages).`));
C.push(P("The H2 result is therefore marginal: at or near the 5 per cent line and dependent on controlling for seasonality. The H1 result is significant under autocorrelation-robust methods, including fixed-b inference, but not under classical methods, and it is sensitive to assumptions about the pre-policy trend (Section 9.4). Neither result should be described as strongly significant."));
C.push(Hd("9.4 Pre-trends and event study", HeadingLevel.HEADING_2));
C.push(caption("Table 9.6: Pre-trend tests and trend-adjusted estimates"));
C.push(table(["Test", "Model 1 (H1)", "Model 2 (H2)"], [
  ["Pre-policy linear trend in the gap, Apr-21 to Nov-23 (pp per month, per year)", `${f2(PT.H1.from_apr21.slope_per_year)} (p = ${p3(PT.H1.from_apr21.p)})`, `${f2(PT.H2.from_apr21.slope_per_year)} (p = ${p3(PT.H2.from_apr21.p)})`],
  ["Pre-policy linear trend in the gap, Jul-21 to Nov-23", `${f2(PT.H1.from_jul21.slope_per_year)} (p = ${p3(PT.H1.from_jul21.p)})`, `${f2(PT.H2.from_jul21.slope_per_year)} (p = ${p3(PT.H2.from_jul21.p)})`],
  ["Main estimate δ (constant pre-policy gap)", f2(H1.its.change), f2(H2.its.change)],
  ["δ if the Apr-21 pre-trend had continued", f2(TC.H1.from_apr21.delta), f2(TC.H2.from_apr21.delta)],
  ["δ if the Jul-21 pre-trend had continued", f2(TC.H1.from_jul21.delta), f2(TC.H2.from_jul21.delta)],
  ["Event study: joint test of all leads = 0, HAC Wald", `F = ${f1(EV.H1.leads_F.F)}, p ${p3(EV.H1.leads_F.p)}`, `F = ${f1(EV.H2.leads_F.F)}, p ${p3(EV.H2.leads_F.p)}`],
  ["Event study: joint test of all leads = 0, classical F", `F = ${f2(EV.H1.leads_F.F_classical)}, p = ${p3(EV.H1.leads_F.p_classical)}`, `F = ${f2(EV.H2.leads_F.F_classical)}, p = ${p3(EV.H2.leads_F.p_classical)}`],
  ["Event study: same, excluding the Apr-2021 block", `F = ${f2(EV.H1.leads_F.F_classical_ex_apr21)}, p = ${p3(EV.H1.leads_F.p_classical_ex_apr21)}`, `F = ${f2(EV.H2.leads_F.F_classical_ex_apr21)}, p = ${p3(EV.H2.leads_F.p_classical_ex_apr21)}`],
], [4960, 2200, 2200]));
C.push(note("Trend tests: Newey–West (3 lags), t-distribution. Trend-adjusted δ: average post-policy deviation of the gap from a linear trend fitted to the pre-policy period only. Event study: three-month blocks, reference block September to 17 November 2023."));
C.push(P(`Neither model shows a statistically significant linear pre-trend. For Model 1, the trend is slightly positive from April 2021 and slightly negative (${f2(PT.H1.from_jul21.slope_per_year)} pp per year) from July 2021. The choice matters for the counterfactual: if the July-2021 trend had continued, the H1 effect would be ${f2(TC.H1.from_jul21.delta)} rather than ${f2(H1.its.change)} percentage points a month, less than half the main estimate. For Model 2 the pre-trend is positive, so projecting it forward makes the estimated effect larger, not smaller.`));
C.push(...fig("fig6_event_study.png", 600, 216, "Figure 9.1: Event study — growth gap by three-month block relative to the circular",
  "Coefficients on three-month blocks relative to the reference block (September to 17 November 2023), with 95% confidence intervals (Newey–West, 3 lags). Block 0 = December 2023 – February 2024."));
C.push(P(`Figure 9.1 shows the event study. For H1, the pre-policy blocks are scattered around zero with no individually significant coefficient apart from the April 2021 COVID-19 block, while all five post-policy blocks are negative. For H2, the pre-policy blocks are much noisier, with several individually significant positive and negative coefficients driven by quarter-end lending. A joint HAC Wald test rejects that all leads are zero for both models, but such tests are known to over-reject badly with ten restrictions and about 45 observations; the classical F-test does not reject (p = ${p3(EV.H1.leads_F.p_classical)} and ${p3(EV.H2.leads_F.p_classical)}), nor does it once the April 2021 block is excluded (p = ${p3(EV.H1.leads_F.p_classical_ex_apr21)} and ${p3(EV.H2.leads_F.p_classical_ex_apr21)}). The event study therefore does not reveal a pre-existing divergence, but for H2 the pre-policy noise is as large as the post-policy effect.`));

// ================= Chapter 10 (Discussion)
C.push(Hd("10. Discussion", HeadingLevel.HEADING_1));
C.push(Hd("10.1 Interpretation", HeadingLevel.HEADING_2));
C.push(P(`Bank credit to NBFCs (excluding HFCs) slowed by about ${f1(-h2a)} percentage points a year relative to total bank credit after the November 2023 circular. The estimate is marginally significant — at the 5 per cent line once seasonality is controlled for, and weaker without — and the gap moved back by a similar amount after the April 2025 rollback, although that rebound is not statistically significant. The timing of the slowdown and the rebound, which match the introduction and withdrawal of the measure, is the strongest piece of evidence that the bank-to-NBFC risk weight affected bank lending to NBFCs.`));
C.push(P(`Targeted bank consumer credit slowed by about ${f1(-h1a)} percentage points a year relative to exempt household credit. The result is consistent in sign across specifications and significant with autocorrelation-robust inference, but not with classical standard errors; it rests mainly on other personal loans; and it is sensitive to whether the pre-policy gap is assumed to be stable or gently declining. It should be read as evidence consistent with the circular slowing bank consumer credit, not as a precise measure of the effect.`));
C.push(P("The firm-level evidence is descriptive and inconclusive. The disclosures show that the higher risk weights immediately cut the capital ratios of two consumer NBFCs, by about 3–4 percentage points. Beyond that, the tests do not show significant effects on NBFCs' funding costs, borrowing mix or lending mix, and the main within-firm case is confounded by an RBI embargo. The study therefore cannot say whether listed NBFCs absorbed the shock or cut their own lending."));
C.push(Hd("10.2 Threats to validity", HeadingLevel.HEADING_2));
C.push(P([B("Policy endogeneity and mean reversion. "), `The RBI acted because unsecured credit and bank lending to NBFCs were growing much faster than other credit. The difference-in-differences design assumes that, without the circular, this excess growth would have continued at its pre-policy average. Booms tend to slow on their own, so this assumption is the main identification caveat. The pre-trend tests (Section 9.4) do not find a significant pre-policy trend, but for H1 a mild, insignificant decline since mid-2021 would, if projected forward, cut the estimated effect by more than half. For H2 the pre-trend runs the other way.`]));
C.push(P([B("A package, not only a capital shock. "), "The circular combined higher risk weights with a requirement to review sectoral exposure limits and set board-approved limits on all unsecured consumer credit by 29 February 2024. Banks were well capitalised (system CRAR of 16.8 per cent in September 2023), so for most banks the risk-weight increase is unlikely to have made capital requirements bind. Part of the measured effect may therefore reflect supervisory signalling and the exposure-limit requirement rather than the cost of capital. The title's reference to capital requirements should be read as referring to the risk-weight package. The decisive test — whether banks with less capital headroom cut targeted lending more — requires bank-level data and is left for future work."]));
C.push(P([B("Concurrent deterioration in credit quality. "), "Delinquencies on credit cards and small-ticket personal loans rose through 2024, and the RBI's Financial Stability Reports flagged rising stress and write-offs in unsecured retail loans (RBI, 2024a, 2024b) [VERIFY: specific delinquency ratios in the June and December 2024 Financial Stability Reports]. Banks would have tightened unsecured lending on that ground alone, so part of the H1 effect, and the continued weakness of card credit after 2025, may reflect credit-quality concerns rather than the risk weight."]));
C.push(P([B("Firm-specific regulatory actions. "), "Two sample firms were under RBI business restrictions during the post-policy period: Bajaj Finance (eCOM and Insta EMI Card, 15 November 2023 to 2 May 2024) and IIFL Finance (gold loans, 4 March to 19 September 2024). Both are disclosed in Chapter 8 and handled by exclusion tests."]));
C.push(P([B("Treatment intensity. "), "The bank-to-NBFC risk weight applied only where the rating-based risk weight was below 100 per cent, and excluded HFCs and priority-sector-eligible on-lending. The RBI's NBFC series includes infrastructure and public-sector NBFCs and NBFC-MFIs, for many of which the increase did not apply or was small. The aggregate H2 estimate is therefore a diluted, intent-to-treat effect, and the true effect on fully treated borrowers would be larger."]));
C.push(P([B("Supply versus demand. "), "Lower bank lending to NBFCs is consistent with both lower bank supply and lower NBFC demand for bank funds. The insignificant cost-of-funds response does not support a strong supply shock, and the study did not collect the bond and commercial-paper issuance data needed to test substitution."]));
C.push(P([B("Control-group contamination. "), "Total non-food credit includes both the treated NBFC credit and the H1-treated consumer categories; a benchmark excluding them gives a similar, slightly larger estimate. The HDFC merger changes the composition of housing, other personal loans and non-food credit; the RBI's merger-adjusted series are used to remove it within the estimation window."]));
C.push(Hd("10.3 Relation to the synopsis hypotheses", HeadingLevel.HEADING_2));
C.push(P("The synopsis proposed five hypotheses. Following the guide's advice the formal tests focus on two; the evidence on the others is limited:"));
C.push(bullet([B("H1 (unsecured exposure and loan growth): "), "bank credit to the targeted unsecured categories slowed relative to exempt credit; the result is consistent in sign but its significance depends on the inference method and pre-trend assumptions. NBFCs' own consumer lending is observed only for Bajaj Finance, where the test is confounded by an RBI embargo."]));
C.push(bullet([B("H2 (bank dependence and borrowing costs): "), "bank lending to NBFCs slowed (marginally significant) and recovered after the rollback. Firm-level funding costs did not rise more for bank-dependent, non-exempt NBFCs; the bank-share gradient in funding costs comes mainly from HFCs and is consistent with the lag in repricing bank loans."]));
C.push(bullet([B("H3 (capital): "), "company disclosures show the policy reduced the capital ratios of two consumer NBFCs by about 3–4 percentage points. The two large equity raises in the sample were set in motion before the circular — Bajaj Finance's board approved its ₹10,000 crore raise on 5 October 2023 and opened the QIP on 6 November 2023, and Cholamandalam's QIP opened on 28 September 2023 — so they cannot be read as responses to the policy, although the new capital will have cushioned its impact."]));
C.push(bullet([B("H4 (shift to secured lending): "), `bank gold loans grew by ${f1(gtG[POST])} per cent a year after the policy (${f1(gtG[PRE])} per cent before), but this coincides with a sharp rise in gold prices and the reclassification of agricultural gold loans as retail gold loans, so the evidence for a policy-driven shift to secured credit is weak.`]));
C.push(bullet([B("H5 (recovery after February 2025): "), `bank credit to NBFCs recovered in aggregate after the April 2025 rollback (${f1(gtNB[ROLL])} per cent a year), although the change in the gap is not statistically significant. At the firm level, the bank share of non-HFC NBFCs drifted down after FY2025 while HFCs' continued to rise; this is a descriptive pattern, not a tested effect.`]));
C.push(Hd("10.4 Policy implications", HeadingLevel.HEADING_2));
C.push(bullet("The evidence is consistent with targeted risk weights slowing bank credit to the targeted segments while exempt categories continued to grow, but the size of the effect is uncertain."));
C.push(bullet("The bank-to-NBFC lever appears to transmit and reverse quickly: bank credit to NBFCs slowed after November 2023 and recovered after April 2025."));
C.push(bullet("Because the circular combined risk weights with exposure limits at a time of ample bank capital, its effect cannot be attributed to the capital charge alone; regulators and analysts should treat such packages as supervisory signals as well as capital measures."));
C.push(bullet("For NBFC managers, dependence on a single funding source is a regulatory risk; diversified funding (bonds, external commercial borrowing, securitisation) reduces exposure to this kind of measure."));
C.push(Hd("10.5 Limitations", HeadingLevel.HEADING_2));
C.push(bullet("The RBI data measure bank credit, not NBFCs' own lending; the aggregate tests capture the banking system's response."));
C.push(bullet("Each aggregate test rests on 45–46 monthly observations, so significance depends on the inference method (Table 9.5)."));
C.push(bullet("The firm-level sample has only 16 firms, which limits statistical power; firm definitions of cost of funds and bank borrowing differ, which is why all firm-level comparisons are within-firm over time."));
C.push(bullet("Some firm-level series have breaks: Shriram Finance is included from FY2023 (after the merger of Shriram Transport Finance and Shriram City Union Finance); Poonawalla Fincorp from FY2022 (after the change of control of Magma Fincorp); Arman Financial's early borrowing figures are partly interpolated; IIFL's bank share is annual only."));
C.push(bullet("Following the guide's advice, the study focuses on two hypotheses; synopsis hypotheses H3–H5 are addressed only descriptively."));

// ================= Chapter 11
C.push(Hd("11. Conclusion", HeadingLevel.HEADING_1));
C.push(P(`This dissertation evaluated the RBI's November 2023 increase in risk weights on consumer credit and on bank lending to NBFCs, using RBI sectoral credit data and a difference-in-differences design. Bank credit to NBFCs slowed by about ${f1(-h2a)} percentage points a year relative to total bank credit, an effect that is marginally significant (p = ${p3(H2.its.p)} in the main specification; at or just below 0.05 with seasonal controls) and that reversed after the April 2025 rollback, although the rebound is not itself significant. Targeted bank consumer credit slowed by about ${f1(-h1a)} percentage points a year relative to exempt household credit, a result that is consistent in sign but significant only with autocorrelation-robust inference and sensitive to assumptions about pre-policy trends. Because the circular combined higher risk weights with required exposure limits at a time of ample bank capital, high credit growth and rising delinquencies, these estimates measure the effect of the whole package rather than of capital requirements alone.`));
C.push(P(`Firm-level evidence from ${cnt.decks} investor presentations confirms that the risk weights cut two consumer NBFCs' capital ratios by about 3–4 percentage points, but tests of NBFCs' funding costs, borrowing mix and lending mix are insignificant and partly confounded, so how NBFCs adjusted remains an open question. Future work could use bank-level data to test whether banks with less capital headroom cut targeted lending more — the decisive test of a capital channel — extend the firm-level data to smaller, unlisted NBFCs, and use bond and commercial-paper issuance data to test whether NBFCs replaced bank funding with market funding.`));

// ================= Bibliography
C.push(Hd("Bibliography", HeadingLevel.HEADING_1));
for (const r of [
  "Acharya, V. V., Khandwala, H., & Öncü, T. S. (2013). The growth of a shadow banking system in emerging markets: Evidence from India. Journal of International Money and Finance, 39, 207–230.",
  "Aiyar, S., Calomiris, C. W., & Wieladek, T. (2014). Does macro-prudential regulation leak? Evidence from a UK policy experiment. Journal of Money, Credit and Banking, 46(s1), 181–214.",
  "Basten, C. (2020). Higher bank capital requirements and mortgage pricing: Evidence from the counter-cyclical capital buffer. Review of Finance, 24(2), 453–495.",
  "Behn, M., Haselmann, R., & Wachtel, P. (2016). Procyclical capital regulation and lending. Journal of Finance, 71(2), 919–956.",
  "Bernanke, B. S., & Blinder, A. S. (1988). Credit, money, and aggregate demand. American Economic Review, 78(2), 435–439.",
  "Business Standard (2025, 4 February). Gold loan portfolio of banks jumps 71.3% to Rs 1.72 trillion till Dec 2024.",
  "Callaway, B., Goodman-Bacon, A., & Sant'Anna, P. H. C. (2024). Difference-in-differences with a continuous treatment. NBER Working Paper No. 32117.",
  "Cameron, A. C., Gelbach, J. B., & Miller, D. L. (2008). Bootstrap-based improvements for inference with clustered errors. Review of Economics and Statistics, 90(3), 414–427.",
  "Driscoll, J. C., & Kraay, A. C. (1998). Consistent covariance matrix estimation with spatially dependent panel data. Review of Economics and Statistics, 80(4), 549–560.",
  "Irani, R. M., Iyer, R., Meisenzahl, R. R., & Peydró, J.-L. (2021). The rise of shadow banking: Evidence from capital regulation. Review of Financial Studies, 34(5), 2181–2235.",
  "Jiménez, G., Ongena, S., Peydró, J.-L., & Saurina, J. (2017). Macroprudential policy, countercyclical bank capital buffers, and credit supply: Evidence from the Spanish dynamic provisioning experiments. Journal of Political Economy, 125(6), 2126–2177.",
  "Kashyap, A. K., & Stein, J. C. (2004). Cyclical implications of the Basel II capital standards. Federal Reserve Bank of Chicago Economic Perspectives, 28(1), 18–31.",
  "Khwaja, A. I., & Mian, A. (2008). Tracing the impact of bank liquidity shocks: Evidence from an emerging market. American Economic Review, 98(4), 1413–1442.",
  "Kiefer, N. M., & Vogelsang, T. J. (2005). A new asymptotic theory for heteroskedasticity-autocorrelation robust tests. Econometric Theory, 21(6), 1130–1164.",
  "Newey, W. K., & West, K. D. (1987). A simple, positive semi-definite, heteroskedasticity and autocorrelation consistent covariance matrix. Econometrica, 55(3), 703–708.",
  "Reserve Bank of India (2005). Prudential norms on capital adequacy — risk weight on commercial real estate exposures. Circular, 26 July 2005.",
  "Reserve Bank of India (2023a). Regulatory measures towards consumer credit and bank credit to NBFCs. Circular RBI/2023-24/85, 16 November 2023.",
  "Reserve Bank of India (2023b). Financial Stability Report, December 2023.",
  "Reserve Bank of India (2024a). Financial Stability Report, June 2024.",
  "Reserve Bank of India (2024b). Financial Stability Report, December 2024.",
  "Reserve Bank of India (2025a). Review of risk weights on microfinance loans. Circular RBI/2024-25/119, 25 February 2025.",
  "Reserve Bank of India (2025b). Exposures of scheduled commercial banks to NBFCs — review of risk weights. Circular RBI/2024-25/120, 25 February 2025.",
  "Reserve Bank of India. Database on Indian Economy (DBIE), Table 15: Deployment of Gross Bank Credit by Major Sectors (accessed 2026).",
  "Webb, M. D. (2023). Reworking wild bootstrap-based inference for clustered errors. Canadian Journal of Economics, 56(3), 839–858.",
  "Company investor presentations, Q1 FY2021 – Q4 FY2026: AAVAS Financiers, Arman Financial Services, Bajaj Finance, Can Fin Homes, Cholamandalam Investment & Finance, CreditAccess Grameen, Home First Finance, IIFL Finance, LIC Housing Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, PNB Housing Finance, Poonawalla Fincorp, Repco Home Finance, SBI Cards and Payment Services, Shriram Finance. Company announcements on the RBI restrictions on Bajaj Finance (15 November 2023; lifted 2 May 2024) and IIFL Finance (4 March 2024; lifted 19 September 2024), and on the Bajaj Finance (5 October 2023) and Cholamandalam (28 September 2023) capital raises.",
]) C.push(new Paragraph({ spacing: { after: 120, line: 300 }, indent: { left: 720, hanging: 720 }, children: [new TextRun(r)] }));

// ================= Appendices
C.push(Hd("Appendices", HeadingLevel.HEADING_1));
C.push(Hd("Appendix A: Data Files and Reproducibility", HeadingLevel.HEADING_2));
C.push(P("All data used in this dissertation are provided in the accompanying workbook Dissertation_Data.xlsx, which contains the following sheets:"));
for (const t of [
  "README — definitions and sources of all variables.",
  "RBI_Monthly — outstanding bank credit by category and monthly growth rates (merger-adjusted where available).",
  "RBI_Growth_Table — the figures in Table 4.1.",
  "COF_Quarterly and BankShare_Quarterly — firm data, with the presentation and page each value came from.",
  "Extraction_Recheck — automated re-check of every extracted value against its cited source.",
  "Bank_Exposure_Sep2023 — verified pre-policy bank shares and the values in the original dataset.",
  "Bajaj_Segments — segment-wise assets of Bajaj Finance.",
  "Company_Disclosures — verbatim company statements in Table 8.2.",
  "Results sheets — all regression results reported in Chapters 7 to 9, including the event study and firm-level subsamples.",
]) C.push(bullet(t));
C.push(P("The analysis was carried out in Python (pandas, statsmodels and linearmodels). The scripts are 05_final_analysis.py (all results, tables and figures) and did_tools.py (two-way fixed-effects estimation with wild cluster bootstrap-t inference)."));
C.push(Hd("Appendix B: Firm-Level Data Notes", HeadingLevel.HEADING_2));
for (const t of [
  "Cost of funds is the company's reported cost of borrowing where available. For Mahindra & Mahindra Financial Services, Shriram Finance and Arman Financial Services it is computed as annualised quarterly finance cost divided by average borrowings from the same presentation; for Cholamandalam it is finance cost as a share of average assets (the only consistently reported measure).",
  "Where a later presentation restates an earlier quarter, the value as first reported is used.",
  "Muthoot Finance's FY2021–22 figures are derived from reported year-to-date averages.",
  "Bank share of borrowings follows each company's own borrowing-mix definition (for example, banks and financial institutions for Mahindra & Mahindra; term loans for Shriram and AAVAS). Cholamandalam's institutional financial investors (IFI), reported separately from FY2024, are added back to banks for consistency.",
  "IIFL Finance's bank share is available only annually on a consistent definition; PNB Housing's earlier figures are rescaled to exclude direct assignment.",
]) C.push(bullet(t));
C.push(Hd("Appendix C: Data Tables", HeadingLevel.HEADING_2));
const shortCo = (c) => String(c).replace(/ Ltd\.?| Limited/g, "").replace("Investment & Finance Company", "Inv. & Fin.").replace(" Company India", "").replace("Financial Services", "Fin. Services");
const appTab = (T, cap, src, fmt) => { C.push(caption(cap));
  C.push(table(T.columns, T.data.map((r) => r.map((v, i) => (i === 0 ? shortCo(v) : v === "" ? "—" : fmt(v)))),
    [2760, ...Array(T.columns.length - 1).fill(Math.floor(6600 / (T.columns.length - 1)))])); C.push(note(src)); };
appTab(R.app_cof, "Table C.1: Cost of funds by company, annual average of quarterly values (per cent)", "Financial years ending March. Source: quarterly investor presentations (see Appendix B for definitions).", f2);
appTab(R.app_bsh, "Table C.2: Bank share of borrowings by company, annual average of quarterly values (per cent)", "Financial years ending March. Poonawalla Fincorp does not disclose its borrowing mix. Source: quarterly investor presentations.", f1);
C.push(caption("Table C.3: Outstanding bank credit by category at end-March (₹ lakh crore)"));
C.push(table(R.app_rbi.columns.map((c) => (c === "Date" ? "Month" : SHORT[c] || c)), R.app_rbi.data.map((r) => r.map((v, i) => (i === 0 ? v : f2(v)))),
  [1200, ...Array(R.app_rbi.columns.length - 1).fill(Math.floor(8160 / (R.app_rbi.columns.length - 1)))]));
C.push(note("Source: RBI DBIE Table 15 (reported series). Figures from July 2023 include the effect of the HDFC merger (housing up, HFC and NBFC credit down)."));


// ---------------------------------------------------------------- document
const PB = { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 24 };
const SECT = { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 600, footer: 600 },
  borders: { pageBorders: { display: PageBorderDisplay.ALL_PAGES, offsetFrom: PageBorderOffsetFrom.PAGE, zOrder: PageBorderZOrder.FRONT },
    pageBorderTop: PB, pageBorderBottom: PB, pageBorderLeft: PB, pageBorderRight: PB } } };
const FOOT = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.LEFT,
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9", space: 4 } },
  children: [new TextRun({ children: [PageNumber.CURRENT], size: 20, bold: true }), new TextRun({ text: " | ", size: 20 }),
    new TextRun({ text: "Page", size: 20, color: "7F7F7F", characterSpacing: 60 })] })] });
const doc = new Document({
  creator: "Nityam Gupta", title: "Capital Requirements and Credit Supply",
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: FONT, color: "000000" }, paragraph: { spacing: { before: 480, after: 360 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: FONT, color: "2E4A7A" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  sections: [
    { properties: { ...SECT, page: { ...SECT.page, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } } }, children: C.slice(0, COVER_END) },
    { properties: { ...SECT, page: { ...SECT.page, pageNumbers: { formatType: NumberFormat.LOWER_ROMAN } } }, footers: { default: FOOT() }, children: C.slice(COVER_END, FRONT_END) },
    { properties: { ...SECT, page: { ...SECT.page, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } }, footers: { default: FOOT() }, children: C.slice(FRONT_END) },
  ],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync("Dissertation.docx", buf); console.log("Dissertation.docx written", buf.length, "bytes"); });
