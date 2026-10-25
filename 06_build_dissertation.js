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
const CAPS = [];
const caption = (t) => { CAPS.push(t); return new Paragraph({ keepNext: true, keepLines: true, alignment: AlignmentType.LEFT, spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: t, bold: true, size: 22 })] }); };
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
const S = JSON.parse(fs.readFileSync("outputs/final/simple_results.json", "utf8"));
const XR = JSON.parse(fs.readFileSync("outputs/final/extra_results.json", "utf8"));
const E = JSON.parse(fs.readFileSync("event_study/results.json", "utf8"));
const GTc = (k) => GT[k];
const gtCC = GT["Credit cards"], gtPL = GT["Other personal loans"], gtNB = GT["NBFCs ex-HFC"], gtNF = GT["Non-food credit"];
const gtH = GT["Housing"], gtV = GT["Vehicle loans"], gtE = GT["Education"], gtG = GT["Gold loans"], gtHF = GT["HFCs"];
const gtCore = GT["Non-food ex NBFC, cards, other PL"];
const PRE = "Pre-policy (Apr-21 to mid-Nov-23)", POST = "Post-policy (Dec-23 to Mar-25)", ROLL = "After rollback (Apr-25 to Jul-26)";
const MG = R.magnitudes, FD = R.firm_desc, cnt = R.counts;
const lakh = (x) => (Number(x) / 100000).toFixed(1);
const fdv = (v, g, p) => FD.find((r) => r.Variable.startsWith(v) && r.Group === g && r.Period.startsWith(p));
const pe = (x) => (Number(x) < 0.001 ? "p < 0.001" : `p = ${Number(x).toFixed(3)}`);
const pc = (x) => (100 * Number(x)).toFixed(2);              // event-study returns in per cent
const pcs = (x) => (Number(x) >= 0 ? "+" : "") + pc(x);
const W1 = S.H1.welch, W2 = S.H2.welch, M1 = S.H1.reg_months, M2 = S.H2.reg_months, R1 = S.H1.reg_simple, R2s = S.H2.reg_simple;
const ann = (d) => f1(12 * Number(d));
const EP = (ek, port) => E[ek].NIFTY50.portfolios[port];
const NH = "Non-HFC NBFCs (10)", HF = "HFCs (6, exempt)", CL = "Consumer lenders (Bajaj, SBI Cards, Poonawalla)";
const DNH = "Non-HFC NBFCs minus HFCs", DCL = "Consumer lenders minus HFCs";
const CS = (ek) => E[ek].NIFTY50.cross_section["[0]"];
const FN = { AAVAS: "AAVAS Financiers", ARMANFIN: "Arman Financial", BAJFINANCE: "Bajaj Finance", CANFINHOME: "Can Fin Homes", CHOLAFIN: "Cholamandalam Inv. & Fin.",
  CREDITACC: "CreditAccess Grameen", HOMEFIRST: "Home First Finance", IIFL: "IIFL Finance", LICHSGFIN: "LIC Housing Finance", MMFIN: "M&M Financial Services",
  MUTHOOTFIN: "Muthoot Finance", PNBHOUSING: "PNB Housing Finance", POONAWALLA: "Poonawalla Fincorp", REPCOHOME: "Repco Home Finance", SBICARD: "SBI Cards & Payment", SHRIRAMFIN: "Shriram Finance" };
const HFCS = ["AAVAS", "CANFINHOME", "HOMEFIRST", "LICHSGFIN", "PNBHOUSING", "REPCOHOME"];
const FC = S.firm.cof, FB = S.firm.bank_share;
const rob = (s) => S.H1.robustness.find((r) => r.spec.startsWith(s));

// ---------------------------------------------------------------- content
const C = [];

const ST = { name: "Nityam Gupta", prog: "MBA (Full-Time)", batch: "2025–27", roll: "FT-25-235", date: "03/10/2026" };
const GD = { name: "Dr. Dezy Kumari", desig: "Assistant Professor, Finance", fac: "Faculty of Management Studies", uni: "University of Delhi" };
const TITLE = "Capital Requirements and Credit Supply: The Effect of the RBI's November 2023 Risk-Weight Measures on Bank Consumer Credit and Bank Funding of NBFCs in India, 2021–2026";

// Cover page
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
C.push(P("AI tools were used to assist with collecting data from company investor presentations and with the Python-based data analysis."));
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
C.push(P(`On 16 November 2023 the Reserve Bank of India (RBI) raised the risk weights on unsecured consumer credit and on bank loans to non-banking financial companies (NBFCs), leaving housing finance companies (HFCs) exempt. In February 2025 it withdrew the higher risk weight on bank loans to NBFCs. This study asks whether these measures slowed the targeted flows of bank credit, and how the stock market judged them. Using monthly RBI data on bank credit by sector, it compares the growth of targeted credit with similar credit that was not targeted, before and after the circular, using t-tests and regressions with dummy variables. Bank credit to NBFCs (excluding HFCs) grew about ${ann(-W2.diff)} percentage points a year more slowly than total bank credit after the circular; the difference is significant at the 5 per cent level once seasonal patterns are controlled for (${pe(M2.p)}). Targeted bank consumer credit (credit cards and other personal loans) grew about ${ann(-W1.diff)} percentage points a year more slowly than exempt household credit, which is significant at the 10 per cent level (${pe(W1.p)}). An event study of 16 listed NBFCs and HFCs gives sharper results. On the first trading day after the circular, the shares of the ten NBFCs fell ${pc(-EP("E1", NH)["[0]"].car)} per cent more than the market predicted (${pe(EP("E1", NH)["[0]"].p)}) while HFC shares did not move; on the first trading day after the rollback, NBFC shares rose ${pc(EP("E2", NH)["[0]"].car)} per cent (${pe(EP("E2", NH)["[0]"].p)}), again with no significant move in HFC shares. Company disclosures show that the higher risk weights cut two consumer lenders' capital ratios by 3 to 4 percentage points, but comparisons of the firms' funding costs and borrowing mix show no clear effect. The evidence suggests that the measures slowed bank credit to the targeted segments and that investors saw them as costly for the lenders they targeted. Because the circular also required lenders to set limits on unsecured lending, the results reflect the package as a whole rather than the capital charge alone.`));
C.push(P([B("Keywords: "), "risk weights, capital requirements, credit supply, NBFCs, consumer credit, event study, macroprudential policy, India"]));

// Table of contents
C.push(new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 480 }, children: [new TextRun({ text: "Table of Contents", bold: true, size: 26 })] }));
C.push(new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }));
let FRONT_END = C.length;
const LISTS_AT = C.length;

// Executive summary
C.push(Hd("Executive Summary", HeadingLevel.HEADING_1, true));
C.push(P([B("Context. "), "On 16 November 2023 the RBI raised the risk weights on unsecured consumer credit (personal loans and credit cards) and on banks' loans to NBFCs, with housing finance companies exempt. The same circular asked lenders to set board-approved limits on unsecured consumer lending by 29 February 2024. The higher risk weight on bank loans to NBFCs was withdrawn from 1 April 2025; the consumer-credit risk weight was not."]));
C.push(P([B("Objective. "), "To test whether the measures slowed bank credit to the targeted segments, how the share prices of the affected lenders reacted, and what listed NBFCs reported about the effect on their business."]));
C.push(P([B("Data. "), `Monthly RBI data on bank credit by sector (January 2019 to July 2026); daily share prices of 16 listed NBFCs and HFCs and the Nifty indices (September 2022 to March 2025); and quarterly data on cost of funds, bank borrowing and capital from ${cnt.decks} company investor presentations.`]));
C.push(P([B("Method. "), "Growth of targeted credit is compared with similar exempt credit before and after the circular using Welch t-tests and regressions with a post-policy dummy and calendar-month dummies. An event study measures abnormal share-price returns around the circular and around its rollback, using the market model. Firm-level changes are compared between NBFCs and the exempt HFCs with t-tests."]));
C.push(P([B("Key findings.")]));
C.push(bullet(`Bank credit to NBFCs (excluding HFCs) grew about ${ann(-W2.diff)} percentage points a year more slowly than total bank credit after the circular (${pe(W2.p)} in a simple t-test; ${pe(M2.p)} after controlling for seasonal patterns), and the gap narrowed again after the April 2025 rollback.`));
C.push(bullet(`Targeted bank consumer credit grew about ${ann(-W1.diff)} percentage points a year more slowly than exempt household credit (${pe(W1.p)}; ${pe(M1.p)} with seasonal controls). The result comes mainly from other personal loans.`));
C.push(bullet(`On the first trading day after the circular, the ten NBFCs' shares fell ${pc(-EP("E1", NH)["[0]"].car)} per cent relative to the market (${pe(EP("E1", NH)["[0]"].p)}), and the three consumer lenders fell ${pc(-EP("E1", CL)["[0,+2]"].car)} per cent over three days (${pe(EP("E1", CL)["[0,+2]"].p)}). HFC shares did not move. On the first trading day after the rollback, NBFC shares rose ${pc(EP("E2", NH)["[0]"].car)} per cent (${pe(EP("E2", NH)["[0]"].p)}).`));
C.push(bullet("Two consumer lenders reported that the rule cut their capital ratios by 3 to 4 percentage points. Funding costs and borrowing mix did not change differently for NBFCs and HFCs."));
C.push(bullet(`If the post-policy gaps had not opened, there would have been roughly ₹${lakh(S.H1.shortfall.amount)} lakh crore more bank credit to cards and personal loans and ₹${lakh(S.H2.shortfall.amount)} lakh crore more bank credit to NBFCs by January 2025. These are rough upper-end figures.`));
C.push(P([B("Implications. "), "The measures appear to have done what the RBI intended: bank credit to the targeted segments slowed while exempt credit kept growing, and the market priced the measures as a cost for the targeted lenders. The bank-to-NBFC lever acted quickly in both directions. Because the circular combined higher risk weights with exposure limits at a time when banks had ample capital, the effect cannot be attributed to the capital charge alone."]));

// Rationale
C.push(Hd("Rationale of the Study", HeadingLevel.HEADING_1));
C.push(P("NBFCs now supply a large share of India's retail credit and depend heavily on banks for funding. A regulatory change that raises the cost of both unsecured consumer lending and bank lending to NBFCs therefore matters to the regulator, to banks, to NBFC managers and to investors. The study is worthwhile for four reasons."));
C.push(num([B("Policy relevance. "), "The November 2023 measures were one of the RBI's most significant interventions in retail credit since the IL&FS crisis, and they were partly reversed in 2025. Whether they worked, and through which channel, is relevant to how such tools are used in future."]));
C.push(num([B("A clear before-and-after setting. "), "The circular changed risk weights for clearly defined types of credit on a known date, left similar types untouched and was later partly withdrawn. This makes it possible to compare targeted credit with untargeted credit that faced the same interest rates and economic conditions."]));
C.push(num([B("Managerial relevance. "), "For NBFC managers the episode shows how dependence on bank funding becomes a regulatory risk, and how capital, pricing and funding mix can be used to absorb such a shock. For investors it shows how quickly such a risk is priced."]));
C.push(num([B("Gap in evidence. "), "Existing discussion of the measures is largely descriptive. There is little formal evidence that separates the consumer-credit and bank-funding channels, or that looks at the market's reaction to both the tightening and its rollback."]));

// ================= Chapter 1
C.push(Hd("1. Introduction", HeadingLevel.HEADING_1));
C.push(Hd("1.1 Background and motivation", HeadingLevel.HEADING_2));
C.push(P("Non-banking financial companies (NBFCs) have become one of the main channels of retail credit in India. Between 2021 and 2023, unsecured consumer lending (personal loans and credit cards) grew far faster than overall credit, and NBFCs both lent directly to households and borrowed heavily from banks to fund that lending. The Reserve Bank of India (RBI) saw this combination as a build-up of risk: rapid unsecured lending on one side, and growing dependence between banks and NBFCs on the other."));
C.push(P("On 16 November 2023 the RBI responded with a single circular. First, it raised the risk weight on consumer credit from 100 per cent to 125 per cent, excluding housing, education, vehicle and gold-backed loans (and, for NBFCs, microfinance loans), and raised the risk weights on credit card receivables. Second, it raised by 25 percentage points the risk weight that banks attach to their loans to most NBFCs, while exempting housing finance companies (HFCs) and loans that qualify as priority-sector lending. Third, it asked banks and NBFCs to review their exposure limits and to set board-approved limits on unsecured consumer credit by 29 February 2024 (RBI, 2023a). In February 2025 the RBI withdrew the higher risk weight on bank loans to NBFCs from 1 April 2025, but kept the consumer-credit risk weights (RBI, 2025b)."));
C.push(P("This sequence gives a useful setting for research. The risk weights changed for clearly defined types of credit on a known date, similar types were left alone, and one of the two measures was later reversed. Comparing targeted with untargeted credit separates the effect of the measures from general conditions such as the interest-rate cycle. The share prices of listed lenders give a second, independent view: if investors expected the rules to hurt targeted lenders, their shares should have fallen when the circular was announced and risen when the bank-to-NBFC measure was withdrawn."));
C.push(Hd("1.2 Research problem", HeadingLevel.HEADING_2));
C.push(P("Market commentary since 2023 suggests that unsecured lending slowed after the RBI's measures, but it is not clear how much of the slowdown was caused by the measures, which channel they worked through, and how NBFCs were affected. This study measures the change in bank credit through each targeted channel relative to exempt credit, tests how the share prices of NBFCs and HFCs reacted to the circular and to its rollback, and documents what listed NBFCs reported about their capital, funding and lending."));
C.push(Hd("1.3 Objectives and hypotheses", HeadingLevel.HEADING_2));
C.push(P("Following the guide's direction to focus on one or two hypotheses, the study tests the two measures in the November 2023 circular:"));
C.push(bullet([B("H1 (consumer-credit channel): "), "The November 2023 measures slowed the growth of bank credit to the targeted consumer categories (credit cards and other personal loans) relative to exempt household credit."]));
C.push(bullet([B("H2 (bank-funding channel): "), "The higher risk weight on bank loans to NBFCs slowed the growth of bank credit to NBFCs (excluding the exempt HFCs) relative to banks' overall lending."]));
C.push(P("Each hypothesis is tested in two ways: in RBI data on bank credit, and in the share prices of listed lenders. If H1 holds, the shares of NBFCs focused on unsecured consumer lending should have fallen on the announcement. If H2 holds, the shares of NBFCs, but not of the exempt HFCs, should have fallen on the announcement and risen when the bank-to-NBFC measure was withdrawn. H1 measures bank consumer credit, not NBFCs' own consumer lending; the title reflects this. The firm-level evidence on NBFCs' funding costs and borrowing mix is reported as supporting, descriptive evidence (Chapter 9)."));
C.push(Hd("1.4 Contribution", HeadingLevel.HEADING_2));
C.push(P(`The study makes three contributions. First, it compares targeted and exempt credit in official RBI data, using the RBI's figures that exclude the effect of the HDFC merger, with placebo and trend checks. Second, it measures the stock market's reaction to both the tightening and its partial reversal, which gives a second test of each channel that does not depend on monthly credit data. Third, it compiles a firm-level dataset from ${cnt.decks} quarterly investor presentations of 16 listed NBFCs and HFCs, covering cost of funds, bank borrowing and the companies' own statements about the policy's effect on their capital.`));
C.push(Hd("1.5 Structure", HeadingLevel.HEADING_2));
C.push(P("Chapter 2 describes the regulatory background and Chapter 3 reviews the literature. Chapter 4 describes the data. Chapter 5 sets out the theoretical framework and hypotheses, and Chapter 6 the research methodology. Chapter 7 presents the tests on bank credit data, Chapter 8 the share-price event study and Chapter 9 the firm-level evidence. Chapter 10 discusses the findings and their limits, and Chapter 11 concludes."));

// ================= Chapter 2
C.push(Hd("2. Regulatory Background", HeadingLevel.HEADING_1));
C.push(Hd("2.1 Risk weights and capital requirements", HeadingLevel.HEADING_2));
C.push(P("Under the Basel framework applied in India, banks and NBFCs must hold regulatory capital equal to a minimum percentage of their risk-weighted assets (RWA). Each loan is multiplied by a risk weight; a 125 per cent risk weight means that every ₹100 lent counts as ₹125 of RWA. Raising the risk weight therefore increases the capital a lender must hold against the same loan or, if capital does not change, lowers its capital-to-risk-weighted-assets ratio (CRAR). For an NBFC that must keep its CRAR above 15 per cent, or a bank that must meet its own capital requirements, a higher risk weight raises the cost of supplying that type of credit."));
C.push(...fig("f_capital_example.png", 560, 224, "Figure 2.1: Minimum capital needed per ₹100 of loans before and after the circular",
  "Banks: minimum total capital of 9 per cent of risk-weighted assets plus a 2.5 per cent capital conservation buffer (11.5 per cent). NBFCs: minimum CRAR of 15 per cent. The NBFC-loan case uses an illustrative rating-based risk weight of 50 per cent."));
C.push(P("Figure 2.1 shows what the circular meant in rupees. A bank lending ₹100 as a personal loan had to hold ₹11.50 of capital before the circular and ₹14.38 after it, a 25 per cent increase in the capital tied up in each loan. For credit card balances the requirement rose from ₹14.38 to ₹17.25. For a bank loan to an NBFC with a 50 per cent rating-based weight, the extra 25 percentage points raised the capital needed from ₹5.75 to ₹8.62, an increase of half. An NBFC making a personal loan, which must keep a CRAR of at least 15 per cent, needed ₹18.75 of capital per ₹100 instead of ₹15. Since equity is the most expensive form of funding, each of these increases raises the cost of supplying the targeted loan."));
C.push(Hd("2.2 The November 2023 circular", HeadingLevel.HEADING_2));
C.push(P("The RBI's circular of 16 November 2023 (“Regulatory measures towards consumer credit and bank credit to NBFCs”, RBI/2023-24/85) contained the measures summarised in Table 2.1."));
C.push(caption("Table 2.1: Main measures, November 2023 and February 2025"));
C.push(table(["Measure", "Change", "Exemptions / notes"], [
  ["Consumer credit (banks)", "Risk weight 100% → 125%", "Excludes housing, education, vehicle and gold-secured loans"],
  ["Retail loans (NBFCs)", "Risk weight 100% → 125%", "Also excludes microfinance and SHG loans"],
  ["Credit card receivables", "Banks 125% → 150%; NBFCs 100% → 125%", "None"],
  ["Bank exposures to NBFCs", "+25 percentage points where the rating-based risk weight was below 100%", "Excludes HFCs, core investment companies and loans eligible for priority-sector classification"],
  ["Credit standards", "Review sectoral exposure limits; board-approved limits for all unsecured consumer credit", "To be implemented by 29 February 2024"],
  ["February 2025 circulars", "Higher risk weight on bank exposures to NBFCs withdrawn from 1 April 2025; bank microfinance loans in the nature of consumer credit moved back to 100%", "Other consumer-credit risk weights unchanged"],
], [2300, 3600, 3460]));
C.push(note("Source: RBI circulars RBI/2023-24/85 (16 November 2023), RBI/2024-25/119 and RBI/2024-25/120 (25 February 2025)."));
C.push(P("Three features of the circular matter for this study. The consumer-credit measure targeted unsecured personal loans and cards but left housing, vehicle and education loans untouched, so those exempt categories provide a comparison group for H1. The bank-to-NBFC measure applied to NBFCs but not to HFCs or to priority-sector on-lending (which covers much bank lending to microfinance NBFCs), so total bank credit and the HFCs provide comparison groups for H2. And the circular was a package: along with the risk weights it required limits on unsecured lending, so any effect found in the data is the effect of the package, not of the capital charge alone."));
C.push(...fig("f_timeline.png", 600, 372, "Figure 2.2: Timeline of policy events, firm-specific events and the analysis windows",
  "Sources: RBI circulars; company announcements; author's analysis windows. Dashed lines mark the circular (16 November 2023) and the rollback announcement (25 February 2025)."));
C.push(P(`Figure 2.2 places the main events and the windows used in this study on one timeline. The circular came about nine months after the last repo rate increase, at a time when policy rates were steady, and the rollback was announced in February 2025, shortly before it took effect on 1 April 2025. Two firm-specific RBI actions fall inside the post-policy period and are dealt with separately in the analysis: the embargo on two of Bajaj Finance's lending products, imposed one day before the circular, and the embargo on IIFL Finance's gold-loan business in 2024. The lower half of the figure shows how the data are used. The bank credit tests compare 32 months before the circular with 14 months after it, ending before the rollback was announced, and a later window after the rollback. The event study estimates each stock's normal behaviour over 220 trading days that end about six weeks before each announcement, so that the estimate is not affected by the news itself.`));
C.push(Hd("2.3 Macroeconomic and banking setting", HeadingLevel.HEADING_2));
C.push(P("The policy came after the RBI had raised the repo rate by 250 basis points between May 2022 and February 2023, to 6.5 per cent, where it stayed until February 2025. Funding costs across the financial system were therefore already rising during 2022–23. Banks were well capitalised: the CRAR of scheduled commercial banks was 16.8 per cent in September 2023 (RBI, 2023b), well above the regulatory minimum. A 25-percentage-point rise in the risk weight on a few loan categories was therefore unlikely to push most banks against their capital limits. Part of its effect may have come from the signal it sent about the RBI's concerns (Section 10.2)."));
C.push(Hd("2.4 The NBFC sector before the policy", HeadingLevel.HEADING_2));
C.push(P("NBFCs are regulated by the RBI but, unlike banks, cannot accept demand deposits. They fund themselves through bank loans, bonds (non-convertible debentures), commercial paper, securitisation and, for some, public deposits. The default of IL&FS in 2018 showed how quickly funding can dry up for NBFCs that depend on short-term market borrowing, and banks subsequently became an even more important source of NBFC funding. In October 2021 the RBI introduced a Scale-Based Regulation framework, placing NBFCs in base, middle, upper and top layers with progressively stricter requirements; most of the firms in this study fall in the middle or upper layers."));
C.push(P(`By October 2023 the two targeted channels were large. Outstanding bank credit was ₹${lakh(MG.cards_oct23)} lakh crore to credit cards and ₹${lakh(MG.pl_oct23)} lakh crore in other personal loans, and bank credit to NBFCs excluding HFCs stood at ₹${lakh(MG.nbfc_oct23)} lakh crore, about ${f1(100 * MG.nbfc_oct23 / MG.nonfood_oct23)} per cent of banks' total non-food credit. All three had been growing at more than 20 per cent a year (Table 4.1). The RBI series for bank credit to NBFCs also includes public financial institutions and infrastructure lenders, many of which are rated AAA or are government-owned; for these the rating-based risk weight is often low, and loans that qualify as priority-sector on-lending were exempt. The aggregate test of H2 therefore measures a diluted effect.`));
C.push(...fig("f_credit_size.png", 560, 238, "Figure 2.3: Outstanding bank credit by category, October 2023",
  "Source: RBI DBIE Table 15 (reported series, observation of 20 October 2023). Orange: categories targeted by the circular; blue: exempt categories."));
C.push(P(`Figure 2.3 compares the size of the categories used in the tests. Housing is the largest exempt category (₹${f1(XR.oct23["Housing"])} lakh crore after the HDFC merger), while other personal loans (₹${f1(XR.oct23["Other personal loans"])} lakh crore) and bank credit to NBFCs (₹${f1(XR.oct23["NBFCs ex-HFC"])} lakh crore) are the two largest targeted categories. Credit cards are much smaller (₹${f1(XR.oct23["Credit cards"])} lakh crore), which partly explains why the card series is noisier and why the H1 results depend mainly on personal loans (Section 7.4).`));

// ================= Chapter 3
C.push(Hd("3. Literature Review", HeadingLevel.HEADING_1));
C.push(Hd("3.1 Capital requirements and credit supply", HeadingLevel.HEADING_2));
C.push(P("A large literature shows that capital requirements affect the supply of bank credit. Kashyap and Stein (2004) explain how risk-sensitive capital rules can amplify credit cycles. Aiyar, Calomiris and Wieladek (2014) use changes in bank-specific capital requirements in the United Kingdom to show that higher requirements reduce lending, while part of the gap is filled by foreign branches outside the regulator's reach. Jiménez, Ongena, Peydró and Saurina (2017) find that Spain's dynamic provisioning, a form of capital buffer, dampened credit cycles, with stronger effects at more constrained banks. Behn, Haselmann and Wachtel (2016) show that when model-based risk weights rose during the 2008 crisis, German banks cut lending more for loans whose risk weights increased. Basten (2020) studies Switzerland's sectoral countercyclical capital buffer on mortgages and finds that banks raised mortgage rates, and that banks with less spare capital raised them more. Together these studies suggest that raising the risk weight on a category of loans should reduce credit supply to that category, which is the logic behind H1, and that the size of the effect depends on how close banks are to their capital limits."));
C.push(Hd("3.2 Bank funding and non-bank lenders", HeadingLevel.HEADING_2));
C.push(P("The bank lending channel literature (Bernanke and Blinder, 1988; Khwaja and Mian, 2008) shows that shocks to banks' ability or willingness to lend are passed on to borrowers who depend on bank credit. Khwaja and Mian (2008) use Pakistani loan-level data to show that firms borrowing from affected banks could not fully replace the lost funding. Irani, Iyer, Meisenzahl and Peydró (2021) show that tighter capital regulation of banks shifts lending activity toward non-bank lenders. In India, Acharya, Khandwala and Öncü (2013) document the close funding links between banks and NBFCs. Raising the risk weight on bank loans to NBFCs is a direct shock to this channel, and H2 tests whether it reduced bank credit to NBFCs."));
C.push(Hd("3.3 Targeted risk weights as a policy tool", HeadingLevel.HEADING_2));
C.push(P("Higher risk weights on specific loan categories are a macroprudential tool: instead of tightening credit across the economy through interest rates, the regulator raises the cost of supplying the type of credit it considers risky. India has a domestic precedent. During the 2004–08 credit boom the RBI raised sectoral risk weights several times: in December 2004 it increased the risk weight on housing loans to individuals from 50 to 75 per cent and on consumer credit, including personal loans and credit card receivables, from 100 to 125 per cent as a temporary counter-cyclical measure (RBI, 2004), the same consumer-credit increase it repeated in 2023; and in July 2005 it raised the risk weight on commercial real estate from 100 to 125 per cent (RBI, 2005). The studies above suggest three predictions for such a measure. Credit growth in the targeted category should slow relative to similar untargeted categories. Lenders may respond not only by lending less but also by raising prices, raising capital or shifting toward exempt lending. And part of the targeted activity may move to lenders or funding sources outside the measure's reach (Aiyar, Calomiris and Wieladek, 2014; Irani et al., 2021)."));
C.push(Hd("3.4 Event studies of regulatory announcements", HeadingLevel.HEADING_2));
C.push(P("Event studies measure how a security's price reacts to new information by comparing its actual return around an announcement with the return that would have been expected without the news. The method goes back to Fama, Fisher, Jensen and Roll (1969). Brown and Warner (1985) show that with daily data the simple market model, which predicts a stock's return from the market's return, performs as well as more complex alternatives, and MacKinlay (1997) sets out the standard procedure: an estimation window to fit the model, an event window to measure abnormal returns, and t-tests of whether those abnormal returns differ from zero. When all firms share the same event date, their abnormal returns are correlated, and grouping firms into portfolios before testing is the usual remedy. Event studies are widely used to assess regulatory announcements because share prices reflect investors' expectations of the effect on future profits as soon as the news is public."));
C.push(Hd("3.5 Research gap", HeadingLevel.HEADING_2));
C.push(P("Existing evidence on the November 2023 measures is largely descriptive, drawn from RBI Financial Stability Reports and market commentary. There is little formal evidence that separates the consumer-credit and bank-funding channels, compares targeted with exempt credit, or measures the market's reaction to both the tightening and its reversal. This study addresses that gap."));

// ================= Chapter 4
C.push(Hd("4. Data", HeadingLevel.HEADING_1));
C.push(Hd("4.1 Bank credit data", HeadingLevel.HEADING_2));
C.push(P(`The main dataset is the RBI's table “Deployment of Gross Bank Credit by Major Sectors” (Database on Indian Economy, Table 15), with ${cnt.rbi_months} monthly observations from January 2019 to July 2026. Each observation refers to the last reporting Friday of the month (for example 17 November 2023 and 29 December 2023). It reports outstanding bank credit (₹ crore) to credit cards, other personal loans, housing, vehicle loans, education, loans against gold jewellery, NBFCs, HFCs and total non-food credit. Bank credit to NBFCs excluding HFCs is NBFC credit minus HFC credit. Growth is measured as the monthly change in the natural logarithm of outstanding credit, multiplied by 100 (approximately per cent per month). The data cover banks accounting for about 95 per cent of non-food credit.`));
C.push(P("One adjustment is needed. On 1 July 2023 HDFC Ltd, then India's largest HFC, merged into HDFC Bank. Its loan book moved onto a bank balance sheet and banks' loans to HDFC Ltd disappeared. For bank credit to NBFCs and HFCs this was a one-off drop in the level, so only the merger month (28 July 2023) is left out. For housing, other personal loans and total non-food credit, growth after July 2023 would also include growth of the former HDFC Ltd book. The RBI publishes these three series excluding the effect of the merger from July 2023 to June 2025, and this study uses those adjusted figures for growth over that period, which covers the whole post-policy window used in the tests."));
C.push(caption("Table 4.1: Annualised growth of bank credit by category (per cent per year)"));
const gtRow = (name, o) => [name, f1(o[PRE]), f1(o[POST]), f1(o[ROLL])];
C.push(table(["Category", "Pre-policy (Apr-21 to mid-Nov-23)", "Post-policy (Dec-23 to Mar-25)", "After rollback (Apr-25 to Jul-26)"], [
  gtRow("Credit cards (targeted, H1)", gtCC), gtRow("Other personal loans (targeted, H1)", gtPL),
  gtRow("Housing (exempt)", gtH), gtRow("Vehicle loans (exempt)", gtV), gtRow("Education (exempt)", gtE), gtRow("Gold loans (exempt)", gtG),
  gtRow("NBFCs ex-HFC (targeted, H2)", gtNB), gtRow("HFCs (exempt)", gtHF), gtRow("Total non-food credit (benchmark, H2)", gtNF),
  gtRow("Non-food excl. NBFCs, cards, other PL (alt. benchmark)", gtCore),
], [3560, 1900, 1950, 1950]));
C.push(note("Average monthly log growth × 12. Merger month excluded; RBI figures excluding the merger used for housing, other personal loans and non-food credit from August 2023 to June 2025. The pre-policy period ends with the observation of 17 November 2023, one day after the circular. Source: RBI DBIE Table 15; author's calculations."));
C.push(P(`Table 4.1 shows the raw pattern. Bank credit to credit cards grew ${f1(gtCC[PRE])} per cent a year before the policy and ${f1(gtCC[POST])} per cent after; other personal loans slowed from ${f1(gtPL[PRE])} to ${f1(gtPL[POST])} per cent. Exempt categories slowed much less (vehicle loans ${f1(gtV[PRE])} to ${f1(gtV[POST])}) or grew faster (housing ${f1(gtH[PRE])} to ${f1(gtH[POST])}; education ${f1(gtE[PRE])} to ${f1(gtE[POST])}). Bank credit to NBFCs excluding HFCs slowed from ${f1(gtNB[PRE])} to ${f1(gtNB[POST])} per cent, while total non-food credit eased only from ${f1(gtNF[PRE])} to ${f1(gtNF[POST])} per cent. After the April 2025 rollback, bank credit to NBFCs recovered to ${f1(gtNB[ROLL])} per cent. Card credit, whose higher risk weight was not reversed, slowed further to ${f1(gtCC[ROLL])} per cent, but rising card delinquencies over the same period (Section 10.2) mean this cannot be read as proof that the risk weight still binds. Bank gold loans grew very fast after the policy (${f1(gtG[POST])} per cent a year), but this mostly reflects higher gold prices and banks reclassifying agricultural gold loans as retail gold loans (Business Standard, 2025b), so it says little about a policy-driven shift to secured lending.`));
C.push(...fig("f_yoy_growth.png", 600, 222, "Figure 4.1: Growth of bank credit over the previous 12 months, by category",
  "Sum of the previous 12 monthly log growth rates (merger-adjusted). Orange lines: targeted categories; blue lines: exempt categories. Dashed lines: the circular (November 2023) and the rollback announcement (February 2025)."));
C.push(P("Figure 4.1 shows the same pattern over time. Year-on-year growth of card and personal-loan credit peaked at around 22 to 32 per cent in 2023 and fell steadily after the circular, while housing, vehicle and education credit kept growing at broadly similar rates. Bank credit to NBFCs had been growing at more than 30 per cent a year in 2023; it slowed sharply after the circular, fell below the growth of total bank credit through 2024, and picked up again from the middle of 2025, after the rollback. The figure also shows why simple before-and-after comparisons need care: growth was already off its peak in some series before November 2023, and the monthly data behind these lines are noisy."));
C.push(Hd("4.2 Share price data", HeadingLevel.HEADING_2));
C.push(P(`Daily closing prices of the 16 sample companies and of the Nifty 50 and Nifty 500 indices were downloaded from Investing.com for 1 September 2022 to 28 March 2025 (${E.meta.firms} stocks, 641 trading days each). All prices are adjusted for share splits and bonus issues; this was checked by looking for one-day price changes of more than 15 per cent. There are only six such days, none on or near the event dates, and the largest (IIFL Finance, March 2024) coincides with the RBI's restriction on its gold-loan business. Daily returns are simple percentage changes in the closing price.`));
C.push(caption("Table 4.2: Descriptive statistics of daily share returns, September 2022 to March 2025 (per cent)"));
const RD = XR.return_desc;
C.push(table(["Company / index", "Type", "Mean", "SD", "Min", "Max", "Days"], [...Object.keys(FN).map((f) => [FN[f], HFCS.includes(f) ? "HFC" : "NBFC", f2(RD[f].mean_pct), f2(RD[f].sd_pct), f2(RD[f].min_pct), f2(RD[f].max_pct), RD[f].n]),
  ["Nifty 50", "Index", f2(RD.NIFTY50.mean_pct), f2(RD.NIFTY50.sd_pct), f2(RD.NIFTY50.min_pct), f2(RD.NIFTY50.max_pct), RD.NIFTY50.n],
  ["Nifty 500", "Index", f2(RD.NIFTY500.mean_pct), f2(RD.NIFTY500.sd_pct), f2(RD.NIFTY500.min_pct), f2(RD.NIFTY500.max_pct), RD.NIFTY500.n]], [2960, 1000, 1000, 1000, 1100, 1100, 1200]));
C.push(note("Simple daily returns. Source: Investing.com; author's calculations."));
C.push(P(`Table 4.2 summarises the daily returns. The stocks are far more volatile than the market: the daily standard deviation of the individual stocks ranges from ${f2(Math.min(...Object.keys(FN).map((f) => RD[f].sd_pct)))} to ${f2(Math.max(...Object.keys(FN).map((f) => RD[f].sd_pct)))} per cent, against ${f2(RD.NIFTY50.sd_pct)} per cent for the Nifty 50. A single day's abnormal return of 2 to 3 per cent for one stock is therefore not unusual, which is why the event study tests portfolios of firms rather than relying on individual stocks.`));
C.push(Hd("4.3 Firm-level data from investor presentations", HeadingLevel.HEADING_2));
C.push(P(`The firm-level sample consists of 16 listed NBFCs and HFCs: six HFCs (AAVAS Financiers, Can Fin Homes, Home First Finance, LIC Housing Finance, PNB Housing Finance and Repco Home Finance) and ten NBFCs (Arman Financial Services, Bajaj Finance, Cholamandalam Investment & Finance, CreditAccess Grameen, IIFL Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, Poonawalla Fincorp, SBI Cards and Payment Services, and Shriram Finance). For each company, all quarterly investor presentations from Q1 FY2021 to Q4 FY2026 were collected, ${cnt.decks} documents in total. The following were compiled from them, each value recorded with the presentation and page it came from:`));
C.push(bullet([B("Cost of funds"), ` (${cnt.cof_obs} company-quarters, all 16 firms): the company's reported cost of borrowing; where not reported, annualised finance cost divided by average borrowings from the same presentation.`]));
C.push(bullet([B("Bank share of borrowings"), ` (${cnt.bankshare_obs} company-quarters, ${cnt.bankshare_firms} firms): the share of borrowings from banks shown in each company's borrowing-mix slide (Poonawalla Fincorp does not disclose its mix).`]));
C.push(bullet([B("Pre-policy bank dependence"), ": the bank share of borrowings at 30 September 2023, taken from each company's Q2 FY2024 presentation."]));
C.push(bullet([B("Bajaj Finance segment-wise assets under management"), " (24 quarters), separating consumer segments affected by the rule from exempt segments."]));
C.push(bullet([B("Company statements"), " on the effect of the RBI's measures on capital ratios and pricing."]));
C.push(Hd("4.4 Data verification", HeadingLevel.HEADING_2));
C.push(P("The project began with a pre-compiled firm-level dataset. Checking it against the companies' own filings showed that several key fields were not reported data. For five lenders, “unsecured” assets were a fixed percentage of total assets in every quarter; for Shriram Finance, the series mixed an assumed 9 per cent share with two quarters of actual personal-loan figures. Bank borrowing shares differed materially from the filings for several firms (for example Arman Financial 71 per cent in the dataset against 32 per cent in its filing; Muthoot Finance 41 against 65 per cent). Headline figures such as total assets and gross NPA ratios were generally accurate. All firm-level variables used in this dissertation were therefore re-collected from the companies' presentations, and the original dataset is not used in any result reported here."));
C.push(caption("Table 4.3: Firm sample and pre-policy bank dependence (30 September 2023)"));
const ver = [
  ["AAVAS Financiers", "HFC", "44.5", "49.6"], ["Arman Financial Services", "NBFC (MFI)", "71.0", "32.3"], ["Bajaj Finance", "NBFC", "23.0", "23.0"],
  ["Can Fin Homes", "HFC", "54.0", "57.0"], ["Cholamandalam Inv. & Fin.", "NBFC", "53.5", "60.0"], ["CreditAccess Grameen", "NBFC (MFI)", "58.0", "52.9"],
  ["Home First Finance", "HFC", "57.5", "54.0"], ["IIFL Finance", "NBFC", "56.0", "55.3"], ["LIC Housing Finance", "HFC", "24.5", "33.0"],
  ["M&M Financial Services", "NBFC", "56.5", "57.2"], ["Muthoot Finance", "NBFC (gold)", "41.0", "65.4"], ["PNB Housing Finance", "HFC", "40.2", "40.2"],
  ["Poonawalla Fincorp", "NBFC", "64.0", "not disclosed"], ["Repco Home Finance", "HFC", "74.0", "75.5"], ["SBI Cards & Payment", "NBFC (cards)", "68.5", "77.0"],
  ["Shriram Finance", "NBFC", "30.5", "25.6"]];
C.push(table(["Company", "Type", "Bank share in original dataset (%)", "Bank share in company filing (%)"], ver, [3226, 1600, 2267, 2267]));
C.push(note("Definitions follow each company's borrowing-mix slide (some include financial institutions). Source: each company's Q2 FY2024 investor presentation (30 September 2023)."));
C.push(caption("Table 4.4: Descriptive statistics of firm-level variables"));
C.push(table(["Variable", "Group", "Period", "N", "Mean", "SD", "Min", "Max"], FD.map((r) => [r.Variable, r.Group, r.Period, r.N, f2(r.Mean), f2(r.SD), f2(r.Min), f2(r.Max)]), [2126, 1500, 1500, 600, 900, 900, 900, 934]));
C.push(note("Company-quarters, Q1 FY2021 to Q4 FY2026. Pre = up to Q2 FY2024 (September 2023); Post = Q3 FY2024 onward."));
C.push(Hd("4.5 Descriptive statistics of the bank credit series", HeadingLevel.HEADING_2));
const SHORT = { "Credit cards": "Cards", "Other personal loans": "Oth. PL", "Housing": "Housing", "Vehicle loans": "Vehicle", "Education": "Educ.", "NBFCs ex-HFC": "NBFC ex-HFC", "HFCs": "HFCs", "Non-food credit": "Non-food" };
C.push(caption("Table 4.5: Descriptive statistics of monthly bank credit growth (per cent per month)"));
C.push(table(["Series", "Period", "N", "Mean", "SD", "Min", "Max"], R.rbi_desc.map((r) => [r.Series, r.Period, r.N, f2(r.Mean), f2(r.SD), f2(r.Min), f2(r.Max)]),
  [2660, 1600, 700, 1100, 1100, 1100, 1100]));
C.push(note("Monthly log growth × 100. Pre-policy: April 2021 to 17 November 2023; post-policy: 29 December 2023 to January 2025. Source: RBI DBIE Table 15."));
C.push(P("Monthly growth of the targeted series is volatile: the standard deviation of monthly credit-card growth is larger than its mean. Much of this volatility is seasonal (festive-season card spending, quarter-end lending to NBFCs), which is why the regressions in Chapter 7 include calendar-month dummies."));
C.push(caption("Table 4.6: Correlation matrix of monthly bank credit growth, April 2021 to January 2025"));
C.push(table(["", ...R.rbi_corr.cols.map((c) => SHORT[c])], R.rbi_corr.rows.map((row, i) => [SHORT[R.rbi_corr.cols[i]], ...row.map((v) => f2(v))]),
  [1360, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000]));
C.push(note("Pearson correlations of monthly log growth rates."));
C.push(P("Correlations between categories are low to moderate. The comparison groups are therefore chosen on economic grounds (similar borrowers, not covered by the rule) rather than because they move together month by month, and the tests compare average growth before and after the policy."));

// ================= Chapter 5
C.push(Hd("5. Theoretical Framework", HeadingLevel.HEADING_1));
C.push(Hd("5.1 Conceptual framework", HeadingLevel.HEADING_2));
C.push(P("A risk weight determines how much capital a lender must hold against each rupee of a loan. When the risk weight on a category of loans rises, the lender either holds more capital, which is costly, or lends less in that category. If it already has plenty of capital, the change may matter mainly as a signal of the regulator's concern. The November 2023 circular raised risk weights on two links in the credit chain at once and, in the same circular, required limits on unsecured consumer credit. Table 5.1 traces each link from the regulatory change to what can be observed in the data."));
C.push(caption("Table 5.1: Transmission channels and observable effects"));
C.push(table(["Channel", "Mechanism", "Observable effect", "Test"], [
  ["Consumer credit (H1)", "Higher capital per rupee of unsecured consumer loans, plus exposure limits → banks slow or reprice such lending", "Slower growth of bank credit to cards and personal loans than to exempt household credit; lower share prices of consumer lenders", "t-test, regression (Ch. 7); event study (Ch. 8)"],
  ["Bank funding of NBFCs (H2)", "Higher capital per rupee lent to NBFCs (HFCs exempt) → banks slow lending to NBFCs", "Slower growth of bank credit to NBFCs than to the banking system; NBFC shares fall on the circular and rise on the rollback, HFC shares do not", "t-test, regression (Ch. 7); event study (Ch. 8)"],
  ["NBFC funding and capital", "Bank-dependent NBFCs face dearer bank loans; consumer NBFCs hold more capital", "Change in cost of funds and bank share of borrowings; company statements on capital", "Group comparisons (Ch. 9)"],
], [1700, 2900, 3060, 1700]));
C.push(...fig("f_framework.png", 600, 324, "Figure 5.1: Conceptual framework linking the circular to the observable effects tested in this study",
  "Author's framework. Each box on the right shows what should be observed if the corresponding channel operated, and the chapter in which it is tested."));
C.push(P("Figure 5.1 draws the same framework as a diagram. The circular changed two risk weights. Each change raised the capital that lenders need for one type of credit, and each should leave a trace in two places: in the volume of bank credit, which is tested with RBI data, and in the share prices of the lenders affected, which are tested with the event study. The middle of the diagram shows the firm-level channel, through NBFCs' own capital ratios, funding costs and bank borrowing, which is examined descriptively in Chapter 9."));
C.push(Hd("5.2 Variables", HeadingLevel.HEADING_2));
C.push(caption("Table 5.2: Definition of variables"));
C.push(table(["Variable", "Type", "Definition", "Source"], [
  ["Gap1(t)", "Dependent (H1)", "Average monthly log growth (×100) of bank credit to credit cards and other personal loans minus that of housing, vehicle and education loans", "RBI DBIE Table 15"],
  ["Gap2(t)", "Dependent (H2)", "Monthly log growth (×100) of bank credit to NBFCs excluding HFCs minus that of total non-food bank credit", "RBI DBIE Table 15"],
  ["Post(t)", "Independent (dummy)", "1 from the observation of 29 December 2023 onward; 0 before (the observation of 17 November 2023 is pre-policy)", "RBI circular"],
  ["Month dummies", "Control", "Eleven dummies for calendar months, capturing seasonal patterns", "Calendar"],
  ["R(i,t)", "Dependent (event study)", "Daily return of stock i", "Investing.com"],
  ["R(m,t)", "Independent (event study)", "Daily return of the Nifty 50 (Nifty 500 as a check)", "Investing.com"],
  ["AR, CAR", "Outcome (event study)", "Abnormal return (actual minus market-model expected return) and its sum over the event window", "Calculated"],
  ["Cost of funds, bank share", "Firm-level", "Average before (to Q2 FY2024) and after (Q3 FY2024 to Q4 FY2025) the policy, for each firm", "Investor presentations"],
], [2000, 1800, 4060, 1500]));
C.push(Hd("5.3 Hypotheses for the bank credit tests", HeadingLevel.HEADING_2));
C.push(P([I("Gap(t) = α + δ · Post(t) + Σ γ(m) · Month(m) + ε(t)")], { align: AlignmentType.CENTER }));
C.push(P([B("H1. Null hypothesis H0: "), "δ₁ = 0, the growth gap between targeted and exempt consumer credit did not change after the circular. ", B("Alternative hypothesis: "), "δ₁ < 0, the gap narrowed (targeted credit slowed relative to exempt credit)."]));
C.push(P([B("H2. Null hypothesis H0: "), "δ₂ = 0, the growth gap between bank credit to NBFCs and total bank credit did not change after the circular. ", B("Alternative hypothesis: "), "δ₂ < 0, bank credit to NBFCs slowed relative to total bank credit."]));
C.push(Hd("5.4 Hypotheses for the event study", HeadingLevel.HEADING_2));
C.push(P([I("Expected return: E[R(i,t)] = a(i) + b(i) · R(m,t);   AR(i,t) = R(i,t) − E[R(i,t)];   CAR = Σ AR over the event window")], { align: AlignmentType.CENTER }));
C.push(P([B("Event 1 (circular). H0: "), "CAR = 0 for NBFCs and for consumer lenders, and their CAR equals that of the exempt HFCs. ", B("Alternative: "), "CAR < 0 for NBFCs and consumer lenders, and lower than for HFCs."]));
C.push(P([B("Event 2 (rollback of the bank-to-NBFC risk weight). H0: "), "CAR = 0 for NBFCs, and equal to that of HFCs. ", B("Alternative: "), "CAR > 0 for NBFCs, and higher than for HFCs."]));
C.push(caption("Table 5.3: Summary of tests"));
C.push(table(["Hypothesis", "Test", "Statistic", "Decision rule"], [
  ["H1, H2 (credit data)", "Welch two-sample t-test of the gap, before vs after", "t", "Reject H0 if p < 0.05 (0.10 reported as weak evidence)"],
  ["H1, H2 (credit data)", "OLS regression with Post and month dummies", "t on δ; F for the model", "As above"],
  ["H1, H2 (share prices)", "Event study, market model, portfolio t-test; sign test", "t on CAR; binomial", "As above"],
  ["Supporting", "Firm-level change, NBFCs vs HFCs (Welch t); within group (paired t)", "t", "As above"],
], [1900, 3660, 1700, 2100]));
C.push(note("All p-values are two-sided, which is conservative relative to the one-sided alternatives."));

// ================= Chapter 6
C.push(Hd("6. Research Methodology", HeadingLevel.HEADING_1));
C.push(Hd("6.1 Research design", HeadingLevel.HEADING_2));
C.push(P("The study compares credit that the regulation targeted with closely related credit that it left alone, before and after the circular. Working with the growth gap between the two (targeted growth minus exempt growth) removes influences that affect both, such as interest rates and overall economic growth. If the circular slowed targeted credit, the gap should be smaller after November 2023 than before. The share-price event study tests the same two channels from a different angle, using the reaction of investors on the days the news became public."));
C.push(P(`The bank credit tests use April 2021 to January 2025, so that the post-policy period ends before the February 2025 rollback was announced. Because RBI observations refer to the last reporting Friday of each month, the November 2023 observation is dated 17 November 2023, one day after the circular, and is treated as pre-policy. The first post-policy observation is that of 29 December 2023. This gives ${W1.n_pre} pre-policy and ${W1.n_post} post-policy months for H1 (${W2.n_pre} and ${W2.n_post} for H2, which loses the merger month).`));
C.push(...fig("f_design.png", 600, 300, "Figure 6.1: Research design: data, methods and outputs",
  "Author's design. Each row is one strand of the analysis; the right-hand column gives the tables in which the results are reported."));
C.push(P("Figure 6.1 summarises the research design. The study has three strands, each using a different source of data and a method suited to it. The bank credit tests use monthly aggregate data and compare growth rates before and after the circular. The event study uses daily share prices and measures the market's reaction on the days the news became public. The firm-level comparisons use quarterly data from company presentations. Because the three strands rely on different data, a finding that appears in more than one of them is harder to explain away as a quirk of any single dataset."));
C.push(Hd("6.2 Welch t-test", HeadingLevel.HEADING_2));
C.push(P("The first test compares the average monthly growth gap after the circular with the average before it, using a two-sample t-test that does not assume equal variances in the two periods (Welch's t-test). The difference in means is the estimated change in relative growth, in percentage points per month; multiplying by 12 gives an approximate annual figure."));
C.push(Hd("6.3 Regression with dummy variables", HeadingLevel.HEADING_2));
C.push(P("The second test regresses the monthly growth gap on a post-policy dummy and eleven calendar-month dummies, by ordinary least squares (OLS). The month dummies absorb seasonal patterns, such as quarter-end lending surges and festive-season card spending, which otherwise add a lot of noise to monthly data. The coefficient on the post-policy dummy (δ) is the change in the gap after the circular, holding the calendar month constant. The model summary, ANOVA table and coefficient table are reported, along with the Durbin–Watson statistic, a Jarque–Bera test of normal residuals, a Breusch–Pagan test of constant variance and variance inflation factors (VIF). As a check, the same regression is also estimated with Newey–West standard errors (Newey and West, 1987), which allow for correlation between neighbouring months."));
C.push(Hd("6.4 Event study", HeadingLevel.HEADING_2));
C.push(P("The event study follows the standard procedure in MacKinlay (1997). For each stock, the market model is estimated by OLS over an estimation window of trading days −250 to −31 before the event, regressing the stock's daily return on the Nifty 50 return. The fitted model gives each day's expected return; the abnormal return (AR) is the actual return minus the expected return, and the cumulative abnormal return (CAR) is the sum of ARs over an event window. Four windows are reported: day 0, days 0 to +1, days 0 to +2 and days −1 to +1."));
C.push(P("The event day (day 0) is the first trading session after the news became public. Both circulars were released after the market closed: the earliest dated news reports of the November 2023 circular appeared at 5:26 pm on 16 November (Business Today, 2023), after the 3:30 pm close, so day 0 is Friday 17 November 2023. The rollback was reported from 9:31 pm on 25 February 2025 (Business Standard, 2025a), and the market was closed on 26 February for Mahashivratri, so day 0 is Thursday 27 February 2025."));
C.push(P("Because all firms share the same event date, their abnormal returns are not independent. The main tests therefore use equal-weighted portfolios: the ten non-HFC NBFCs, the six HFCs, the three consumer lenders (Bajaj Finance, SBI Cards and Poonawalla Fincorp), and the difference between the NBFC and HFC portfolios. For each portfolio the market model is estimated on the portfolio return, and the t-statistic is CAR divided by the standard deviation of the estimation-window residuals times the square root of the window length. Two further checks are reported: a sign test of how many NBFCs moved in the predicted direction, and a comparison of the day-0 difference between NBFCs and HFCs with the same difference on each of the 220 ordinary days in the estimation window."));
C.push(Hd("6.5 Firm-level comparisons", HeadingLevel.HEADING_2));
C.push(P("For each firm, the average cost of funds and bank share of borrowings are calculated before the policy (up to Q2 FY2024) and after it (Q3 FY2024 to Q4 FY2025, before the rollback). The change for the ten NBFCs is compared with the change for the six exempt HFCs using Welch t-tests, and the before and after averages within each group are compared with paired t-tests."));
C.push(Hd("6.6 Checks", HeadingLevel.HEADING_2));
C.push(P("Three checks test whether the bank credit results could reflect something other than the policy. A placebo test repeats the t-test and the regression on pre-policy data only (July 2021 to November 2023), with a fictitious policy date of November 2022; a significant result would suggest the gap was already moving before the policy. A trend test regresses the pre-policy gap on time. And the H1 test is repeated after dropping each treated and control series in turn, while the H2 test is repeated against a benchmark that excludes bank credit to NBFCs, cards and personal loans from total credit. The July 2021 start of the placebo window avoids the April to June 2021 COVID-19 wave, which produced an extreme negative gap."));

// ================= Chapter 7
C.push(Hd("7. Results: Bank Credit Data", HeadingLevel.HEADING_1));
C.push(...fig("fig1_rbi_indices.png", 600, 234, "Figure 7.1: Bank credit indices, targeted versus exempt categories (October 2023 = 100)",
  "Source: RBI DBIE Table 15; author's calculations. Indices chain monthly growth; merger-adjusted figures used where available."));
C.push(P("Figure 7.1 shows the raw data. Before November 2023, bank credit to cards and personal loans (left panel) and to NBFCs (right panel) was growing faster than the comparison series. After the circular both targeted series grew more slowly than their comparison series, and bank credit to NBFCs picked up again after the April 2025 rollback."));
C.push(Hd("7.1 Welch t-tests", HeadingLevel.HEADING_2));
C.push(caption("Table 7.1: Monthly growth gap before and after the circular (Welch t-test)"));
const wrow = (lbl, w) => [lbl, f2(w.mean_pre), f2(w.mean_post), f2(w.diff), ann(w.diff), f2(w.t), f1(w.df), p3(w.p)];
C.push(table(["Gap", "Mean before", "Mean after", "Difference (pp/month)", "Per year (pp)", "t", "df", "p-value"], [
  wrow("H1: targeted minus exempt consumer credit", W1), wrow("H2: NBFCs ex-HFC minus non-food credit", W2),
  wrow("H2: NBFCs minus non-food excl. NBFCs, cards, PL", S.H2_core.welch)], [2960, 900, 900, 1100, 1000, 800, 700, 1000]));
C.push(note(`Monthly growth gaps in percentage points. Before: April 2021 to 17 November 2023 (H1 n = ${W1.n_pre}; H2 n = ${W2.n_pre}); after: 29 December 2023 to January 2025 (n = ${W1.n_post}).`));
C.push(P(`For H1, the monthly growth gap fell from ${f2(W1.mean_pre)} to ${f2(W1.mean_post)} percentage points after the circular, a fall of ${f2(-W1.diff)} points a month or about ${ann(-W1.diff)} points a year (t = ${f2(W1.t)}, ${pe(W1.p)}). This is significant at the 10 per cent level but not at 5 per cent. For H2, the gap fell from ${f2(W2.mean_pre)} to ${f2(W2.mean_post)}, a fall of ${f2(-W2.diff)} points a month or about ${ann(-W2.diff)} points a year (${pe(W2.p)}). Against the narrower benchmark that excludes NBFCs, cards and personal loans from total credit, the fall is slightly larger (${f2(-S.H2_core.welch.diff)}, ${pe(S.H2_core.welch.p)}). Both tests point in the predicted direction, but monthly credit growth is noisy and the simple t-test does not separate the policy from seasonal swings.`));
C.push(...fig("f_gap_boxplot.png", 560, 224, "Figure 7.2: Distribution of the monthly growth gaps before and after the circular",
  "Boxes show the middle half of the months and the line inside each box the median; whiskers extend to 1.5 times the box height; dots are individual months; the coloured horizontal line is the mean used in the t-test."));
C.push(P("Figure 7.2 shows the monthly gaps behind these averages. In both cases the whole distribution moved down after the circular: the median, the mean and most individual months are lower. The months before the circular are also much more spread out than the months after it, which is why the t-test that allows for unequal variances (Welch's test) is used rather than the standard version. The large overlap between the two boxes is the reason the simple t-tests are only marginally significant: monthly growth swings a lot from one month to the next."));
C.push(...fig("fig2_h1_gap.png", 560, 238, "Figure 7.3: Monthly growth gap, targeted minus exempt consumer credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters. 2023Q4 (October, 17 November and 29 December observations) is shown grey. Dashed line: pre-policy average gap."));
C.push(Hd("7.2 Regression with dummy variables", HeadingLevel.HEADING_2));
C.push(...fig("f_seasonal.png", 560, 194, "Figure 7.4: Average monthly growth gap by calendar month, April 2021 to January 2025",
  "Simple averages of the monthly gaps for each calendar month."));
C.push(P(`Much of the month-to-month noise in the gaps is seasonal (Figure 7.4). For H2, the gap is strongly positive in October to December and negative in January, April, May and August, reflecting the timing of banks' lending to NBFCs around quarter-ends and the festive season. For H1, the gap is high in March, July and October and low in September. If the before and after periods contain a different mix of calendar months, these patterns can hide or exaggerate a policy effect. The regressions below therefore include a dummy for each calendar month.`));
C.push(caption("Table 7.2: Model summary"));
C.push(table(["Model", "Dependent variable", "R", "R²", "Adjusted R²", "Std. error of estimate", "Durbin–Watson"], [[M1, "Gap1 (H1)"], [M2, "Gap2 (H2)"]].map(([M, d], k) =>
  [String(k + 1), d, f2(Math.sqrt(M.r2)), f2(M.r2), f2(M.adj_r2), f2(Math.sqrt(M.ss_res / M.df_res)), f2(M.dw)]), [800, 1700, 900, 900, 1300, 2160, 1600]));
C.push(note("Predictors: (Constant), Post, eleven calendar-month dummies."));
C.push(P([B("ANOVA. Null hypothesis: "), "the predictors jointly have no effect on the growth gap. ", B("Alternative hypothesis: "), "at least one predictor has an effect."]));
C.push(caption("Table 7.3: ANOVA"));
const msq = (x, d) => f2(Number(x) / Number(d));
C.push(table(["Model", "Source", "Sum of squares", "df", "Mean square", "F", "Sig."], [M1, M2].flatMap((M, k) => [
  [String(k + 1), "Regression", f2(M.ss_reg), M.df_reg, msq(M.ss_reg, M.df_reg), f2(M.F), p3(M.F_p)],
  ["", "Residual", f2(M.ss_res), M.df_res, msq(M.ss_res, M.df_res), "", ""],
  ["", "Total", f2(M.ss_reg + M.ss_res), M.df_reg + M.df_res, "", "", ""]]), [900, 1400, 1660, 800, 1600, 1400, 1600]));
C.push(caption("Table 7.4: Coefficient on the post-policy dummy"));
C.push(table(["Model", "B (Post)", "Std. error", "t", "Sig.", "Newey–West std. error", "Newey–West sig."], [M1, M2].map((M, k) =>
  [String(k + 1), f2(M.b), f2(M.se), f2(M.t), p3(M.p), f2(M.se_nw), p3(M.p_nw)]), [900, 1300, 1300, 1100, 1100, 1960, 1700]));
C.push(note("Month-dummy coefficients are not shown; they are jointly significant in both models (F-test p = " + p3(M1.months_joint_p) + " and " + p3(M2.months_joint_p) + "). Newey–West standard errors with three lags."));
C.push(P([B("Model equations: "), I(`Gap1(t) = ${f2(M1.const)} − ${f2(-M1.b)} · Post(t) + month effects;   Gap2(t) = ${f2(M2.const)} − ${f2(-M2.b)} · Post(t) + month effects`)]));
C.push(P(`Adding the month dummies raises R² from about 0.05 to ${f2(M1.r2)} and ${f2(M2.r2)}, and both models are significant overall (F-test ${pe(M1.F_p)} and ${p3(M2.F_p)}). Seasonal patterns explain about half of the month-to-month variation in the gaps. Holding the calendar month constant, the post-policy dummy is ${f2(M2.b)} for H2 (${pe(M2.p)}), which is significant at the 5 per cent level, and ${f2(M1.b)} for H1 (${pe(M1.p)}), significant at the 10 per cent level. With Newey–West standard errors, which allow for correlation between neighbouring months, both coefficients are significant at the 1 per cent level (${pe(M1.p_nw)} and ${p3(M2.p_nw)}). Against the narrower H2 benchmark the coefficient is ${f2(S.H2_core.reg_months.b)} (${pe(S.H2_core.reg_months.p)}).`));
C.push(Hd("7.3 Diagnostic tests", HeadingLevel.HEADING_2));
C.push(caption("Table 7.5: Diagnostic tests on the regression residuals"));
C.push(table(["Test", "Null hypothesis", "Model 1 statistic (p)", "Model 2 statistic (p)"], [
  ["Durbin–Watson", "No first-order autocorrelation (value near 2)", f2(M1.dw), f2(M2.dw)],
  ["Jarque–Bera", "Residuals are normally distributed", `${f2(M1.jb)} (${p3(M1.jb_p)})`, `${f2(M2.jb)} (${p3(M2.jb_p)})`],
  ["Breusch–Pagan", "Residual variance is constant", `${f2(M1.bp)} (${p3(M1.bp_p)})`, `${f2(M2.bp)} (${p3(M2.bp_p)})`],
  ["Largest VIF", "No serious multicollinearity (VIF < 10)", f2(M1.max_vif), f2(M2.max_vif)],
], [2400, 3160, 1900, 1900]));
C.push(P(`The Durbin–Watson statistics (${f2(M1.dw)} and ${f2(M2.dw)}) are close to 2, so there is little first-order autocorrelation. Normality and constant variance are not rejected at the 5 per cent level, although with 45 to 46 observations these tests have limited power. The largest VIF is below 2, so the post-policy dummy is not confounded with the month dummies.`));
C.push(...fig("f_residuals.png", 560, 307, "Figure 7.5: Residuals of the two regressions over time and their distribution",
  "Left: residuals by month (dashed line: the circular). Right: histogram of residuals with a normal curve of the same standard deviation."));
C.push(P("Figure 7.5 plots the residuals of the two regressions. They move around zero without long runs above or below it and with no visible change in spread after the circular, consistent with the Durbin–Watson and Breusch–Pagan results. The histograms are roughly bell-shaped, with a few large months in the tails, consistent with the Jarque–Bera test."));
C.push(Hd("7.4 Checks", HeadingLevel.HEADING_2));
C.push(caption("Table 7.6: Placebo, trend and rollback checks"));
C.push(table(["Check", "Estimate", "p-value", "Interpretation"], [
  ["H1 placebo (fake date Nov-2022), t-test", f2(S.H1.placebo_welch.diff), p3(S.H1.placebo_welch.p), "No break before the policy"],
  ["H1 placebo, regression with month dummies", f2(S.H1.placebo_reg_months.b), p3(S.H1.placebo_reg_months.p), "No break before the policy"],
  ["H1 pre-policy trend (pp per month, per year)", f2(S.H1.pretrend.slope_per_year), p3(S.H1.pretrend.p), "No significant trend"],
  ["H2 placebo (fake date Nov-2022), t-test", f2(S.H2.placebo_welch.diff), p3(S.H2.placebo_welch.p), "No break before the policy"],
  ["H2 placebo, regression with month dummies", f2(S.H2.placebo_reg_months.b), p3(S.H2.placebo_reg_months.p), "No break before the policy"],
  ["H2 pre-policy trend", f2(S.H2.pretrend.slope_per_year), p3(S.H2.pretrend.p), "No significant trend"],
  ["H2 after the rollback (Apr-25 to Jul-26 vs Dec-23 to Mar-25), t-test", f2(S.H2.rollback_welch.diff), p3(S.H2.rollback_welch.p), "Gap reverses, not significant"],
], [4160, 1300, 1200, 2700]));
C.push(note("Placebo and trend tests use pre-policy data from July 2021 to 17 November 2023."));
C.push(P(`Neither placebo test finds a change at the fictitious November 2022 date (${pe(S.H1.placebo_welch.p)} and ${p3(S.H2.placebo_welch.p)}), and neither gap shows a significant trend before the policy. The changes measured after November 2023 were therefore not already under way. The H1 gap did drift down slightly before the policy (${f2(S.H1.pretrend.slope_per_year)} points a year, not significant); if that drift had continued, part of the measured fall would have happened anyway, so the H1 estimate is best treated as an upper bound. After the April 2025 rollback, the H2 gap moved back by ${f2(S.H2.rollback_welch.diff)} points a month, close to the size of the original fall, though with only 15 months of data this is not statistically significant.`));
C.push(caption("Table 7.7: H1 with one series dropped"));
C.push(table(["Specification", "Difference (t-test)", "p-value", "δ (regression with month dummies)", "p-value"],
  S.H1.robustness.map((r) => [r.spec, f2(r.diff), p3(r.p_welch), f2(r.b_months), p3(r.p_months)]), [3760, 1400, 1100, 2000, 1100]));
C.push(P(`The H1 result comes mainly from other personal loans, the larger of the two targeted series (${f2(rob("Other personal loans only").diff)}, ${pe(rob("Other personal loans only").p_welch)}). Credit cards on their own show a smaller, insignificant fall. The estimate is negative whichever control series is dropped.`));
C.push(...fig("fig3_h2_gap.png", 560, 238, "Figure 7.6: Monthly growth gap, bank credit to NBFCs (ex-HFC) minus total non-food credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters (including after the April 2025 rollback)."));
C.push(Hd("7.5 Economic size", HeadingLevel.HEADING_2));
C.push(P(`Over the ${S.H1.shortfall.months} post-policy months to January 2025, a fall of ${f2(-W1.diff)} points a month in relative growth adds up to about ${f1(S.H1.shortfall.pct)} per cent. With ₹${lakh(S.H1.shortfall.outstanding)} lakh crore of bank credit to cards and other personal loans outstanding in January 2025, this implies roughly ₹${lakh(S.H1.shortfall.amount)} lakh crore less targeted credit than if the pre-policy gap had continued. The same calculation for H2 gives about ${f1(S.H2.shortfall.pct)} per cent of ₹${lakh(S.H2.shortfall.outstanding)} lakh crore, or roughly ₹${lakh(S.H2.shortfall.amount)} lakh crore less bank credit to NBFCs. These are rough upper-end figures, since fast credit growth tends to slow on its own.`));

// ================= Chapter 8
C.push(Hd("8. Results: Share-Price Event Study", HeadingLevel.HEADING_1));
C.push(P(`The event study uses daily returns of the 16 stocks. Day 0 is 17 November 2023 for the circular and 27 February 2025 for the rollback (Section 6.4). On those days the Nifty 50 moved ${pc(E.E1.NIFTY50.market_return_day0)} and ${pc(E.E2.NIFTY50.market_return_day0)} per cent, so neither reaction was driven by a broad market move.`));
const prow = (ek, pn, lbl) => { const d = EP(ek, pn); return [lbl, ...["[0]", "[0,+1]", "[0,+2]", "[-1,+1]"].flatMap((w) => [pcs(d[w].car), p3(d[w].p)])]; };
const pw = [2560, 850, 750, 850, 750, 850, 750, 850, 750];
const phead = ["Portfolio", "CAR [0] %", "p", "CAR [0,+1] %", "p", "CAR [0,+2] %", "p", "CAR [−1,+1] %", "p"];
C.push(Hd("8.1 Market model estimates", HeadingLevel.HEADING_2));
C.push(caption("Table 8.1: Market model estimates for each stock (estimation windows of 220 trading days)"));
const MM = XR.market_model;
C.push(table(["Company", "Type", "Circular: beta", "R²", "Residual SD %", "Rollback: beta", "R²", "Residual SD %"], Object.keys(FN).map((f) =>
  [FN[f], HFCS.includes(f) ? "HFC" : "NBFC", f2(MM.E1[f].beta), f2(MM.E1[f].r2), f2(MM.E1[f].sd_pct), f2(MM.E2[f].beta), f2(MM.E2[f].r2), f2(MM.E2[f].sd_pct)]),
  [2560, 900, 1000, 800, 1100, 1000, 800, 1200]));
C.push(note(`OLS of each stock's daily return on the Nifty 50 return. Estimation windows: ${XR.event_windows.E1.est_start} to ${XR.event_windows.E1.est_end} (circular) and ${XR.event_windows.E2.est_start} to ${XR.event_windows.E2.est_end} (rollback).`));
C.push(P("Table 8.1 reports the market model for each stock. Betas range from about 0.1 to 1.7. Larger and more widely traded stocks, such as Bajaj Finance, LIC Housing Finance and Mahindra & Mahindra Financial, have the highest R², while smaller lenders such as AAVAS Financiers, Arman Financial and Home First Finance have low betas and R² values close to zero, meaning that their returns are driven mostly by company-specific news. The residual standard deviation is the typical size of a stock's daily abnormal return on an ordinary day, and it is the yardstick against which the event-day returns are judged."));
C.push(Hd("8.2 The November 2023 circular", HeadingLevel.HEADING_2));
C.push(caption("Table 8.2: Cumulative abnormal returns around the circular (day 0 = 17 November 2023)"));
C.push(table(phead, [prow("E1", NH, "Non-HFC NBFCs (10)"), prow("E1", HF, "HFCs, exempt (6)"), prow("E1", CL, "Consumer lenders (3)"),
  prow("E1", DNH, "NBFCs minus HFCs"), prow("E1", DCL, "Consumer lenders minus HFCs"),
  prow("E1", "Non-HFC NBFCs excl. Bajaj (9)", "NBFCs excluding Bajaj (9)"), prow("E1", "Non-HFC NBFCs excl. MFIs (8)", "NBFCs excluding microfinance (8)")], pw));
C.push(note("Equal-weighted portfolios; market model on the Nifty 50, estimated over trading days −250 to −31. p-values from the t-distribution."));
C.push(P(`On the first trading day after the circular, the ten NBFCs' shares fell ${pc(-EP("E1", NH)["[0]"].car)} per cent more than the market model predicted (${pe(EP("E1", NH)["[0]"].p)}), while the HFCs' abnormal return was zero (${pcs(EP("E1", HF)["[0]"].car)} per cent, ${pe(EP("E1", HF)["[0]"].p)}). The difference between the two portfolios, ${pc(EP("E1", DNH)["[0]"].car)} per cent, is significant (${pe(EP("E1", DNH)["[0]"].p)}). Nine of the ten NBFCs fell on the day (sign test ${pe(CS("E1").sign_nonhfc.p)}), and the NBFC-minus-HFC return on day 0 was larger than on any of the ${E.E1.placebo_days[DNH].days} ordinary trading days in the estimation window.`));
C.push(P(`The reaction was strongest for the three consumer lenders, which were hit by the consumer-credit measure directly: ${pc(EP("E1", CL)["[0]"].car)} per cent on day 0 (${pe(EP("E1", CL)["[0]"].p)}) and ${pc(EP("E1", CL)["[0,+2]"].car)} per cent over three days (${pe(EP("E1", CL)["[0,+2]"].p)}). Over those three days they fell ${pc(-EP("E1", DCL)["[0,+2]"].car)} per cent more than the HFCs (${pe(EP("E1", DCL)["[0,+2]"].p)}). The results are similar without Bajaj Finance, whose separate RBI embargo was announced the day before, and without the two microfinance lenders, whose bank borrowing was largely exempt. Over the following week the HFCs' shares also drifted down (Figure 8.3), so the clear difference between the groups is in the immediate reaction rather than over longer windows.`));
C.push(...fig("f_daily_ar.png", 600, 216, "Figure 8.1: Daily abnormal returns of the NBFC and HFC portfolios, five days either side of each announcement",
  "Equal-weighted portfolios; market model on the Nifty 50."));
C.push(P("Figure 8.1 shows the day-by-day pattern behind the cumulative figures. Around the circular, the NBFC portfolio's largest abnormal return in the eleven-day window is the fall on day 0; on the days before it, NBFC and HFC returns moved in a similar, small range. Around the rollback, the NBFC portfolio's largest move is the rise on day 0, and the HFC portfolio's largest move comes a day later. In both cases the reaction is concentrated on the first trading session after the news, which is what would be expected in a market that absorbs public information quickly."));
C.push(Hd("8.3 The February 2025 rollback", HeadingLevel.HEADING_2));
C.push(caption("Table 8.3: Cumulative abnormal returns around the rollback (day 0 = 27 February 2025)"));
C.push(table(phead, [prow("E2", NH, "Non-HFC NBFCs (10)"), prow("E2", HF, "HFCs, exempt (6)"), prow("E2", CL, "Consumer lenders (3)"),
  prow("E2", DNH, "NBFCs minus HFCs"), prow("E2", DCL, "Consumer lenders minus HFCs"),
  prow("E2", "Non-HFC NBFCs excl. Bajaj (9)", "NBFCs excluding Bajaj (9)"), prow("E2", "Non-HFC NBFCs excl. MFIs (8)", "NBFCs excluding microfinance (8)")], pw));
C.push(note("As Table 8.1. The rollback was announced on the evening of 25 February 2025; the market was closed on 26 February."));
C.push(P(`When the RBI withdrew the higher risk weight on bank loans to NBFCs, the reaction reversed. On day 0 the NBFCs' shares rose ${pc(EP("E2", NH)["[0]"].car)} per cent more than predicted (${pe(EP("E2", NH)["[0]"].p)}), while the HFCs' abnormal return was ${pcs(EP("E2", HF)["[0]"].car)} per cent (${pe(EP("E2", HF)["[0]"].p)}). The difference, ${pc(EP("E2", DNH)["[0]"].car)} per cent, is significant (${pe(EP("E2", DNH)["[0]"].p)}), nine of the ten NBFCs rose (sign test ${pe(CS("E2").sign_nonhfc.p)}), and only ${E.E2.placebo_days[DNH].days_as_large} of ${E.E2.placebo_days[DNH].days} ordinary days had a difference this large. The same day the RBI also eased the risk weight on banks' microfinance loans, which helps explain the large rise in CreditAccess Grameen (Table 8.4); without the two microfinance lenders the NBFC portfolio still rose ${pc(EP("E2", "Non-HFC NBFCs excl. MFIs (8)")["[0]"].car)} per cent (${pe(EP("E2", "Non-HFC NBFCs excl. MFIs (8)")["[0]"].p)}). The consumer lenders, whose consumer-credit risk weight was not reversed, rose by less relative to the HFCs (${pc(EP("E2", DCL)["[0]"].car)} per cent, ${pe(EP("E2", DCL)["[0]"].p)}). The difference between NBFCs and HFCs disappears from day +1, when HFC shares caught up, partly because of a large one-day rise in Home First Finance.`));
C.push(...fig("f_placebo_days.png", 600, 198, "Figure 8.2: Day-0 difference between NBFC and HFC abnormal returns compared with 220 ordinary trading days",
  "Grey bars: daily abnormal return of the NBFC-minus-HFC portfolio on each day of the estimation window. Orange line: day 0."));
C.push(P(`Figure 8.2 provides a simple check that does not rely on the t-distribution. It compares the day-0 difference between the NBFC and HFC portfolios with the same difference on each of the 220 ordinary days used to estimate the market model. After the circular, the day-0 value (${pc(E.E1.placebo_days[DNH].day0_ar)} per cent) lies outside the whole range of ordinary days. After the rollback, only ${E.E2.placebo_days[DNH].days_as_large} of the 220 ordinary days had a difference as large as on day 0 (${pcs(E.E2.placebo_days[DNH].day0_ar)} per cent), which corresponds to a probability of about ${f2(E.E2.placebo_days[DNH].empirical_p)}.`));
C.push(...fig("fig7_event_car.png", 600, 228, "Figure 8.3: Cumulative abnormal returns of NBFC, HFC and consumer-lender portfolios around the two announcements",
  "Market model on the Nifty 50. Dashed line: between day −1 and day 0."));
C.push(P("Figure 8.3 traces the cumulative abnormal returns of the three portfolios. Before the circular, the NBFC and HFC lines move together; on day 0 the NBFC and consumer-lender lines drop, and the consumer lenders keep falling for two more days. In the rollback panel, the NBFC line jumps on day 0 and stays at the higher level, while the HFC line rises more gradually over the following days."));

C.push(Hd("8.4 Firm-level reactions", HeadingLevel.HEADING_2));
C.push(caption("Table 8.4: Day-0 abnormal returns by company (per cent)"));
const frows = Object.keys(FN).map((f) => { const a = E.E1.NIFTY50.firms[f], b = E.E2.NIFTY50.firms[f];
  return [FN[f], HFCS.includes(f) ? "HFC" : "NBFC", pcs(a["[0]"].car), p3(a["[0]"].p), pcs(b["[0]"].car), p3(b["[0]"].p)]; });
C.push(table(["Company", "Type", "Circular, AR day 0", "p", "Rollback, AR day 0", "p"], frows, [3060, 1100, 1500, 1100, 1500, 1100]));
C.push(...fig("f_firm_ar.png", 600, 264, "Figure 8.4: Day-0 abnormal return of each company, circular and rollback",
  "Market model on the Nifty 50. Purple: the three consumer lenders; orange: other NBFCs; blue: HFCs."));
C.push(P(`Among the NBFCs, the largest falls after the circular were at SBI Cards (${pc(E.E1.NIFTY50.firms.SBICARD["[0]"].car)} per cent), Arman Financial, CreditAccess Grameen, IIFL Finance, Cholamandalam and Poonawalla, and the largest rises after the rollback were at CreditAccess Grameen, Shriram Finance, Cholamandalam and Mahindra & Mahindra Financial. Within the NBFC group, the size of the reaction was not related to how much a firm borrowed from banks before the policy (${pe(E.E1.bankshare_within_nonhfc.p)} and ${p3(E.E2.bankshare_within_nonhfc.p)}): investors appear to have reacted to whether a lender was covered by the rule rather than to its exact funding mix. The results are almost identical with the Nifty 500 as the market index (Appendix C).`));

// ================= Chapter 9
C.push(Hd("9. Firm-Level Evidence", HeadingLevel.HEADING_1));
C.push(Hd("9.1 Company disclosures", HeadingLevel.HEADING_2));
C.push(P("Two consumer lenders quantified the immediate capital impact of the rule in their December 2023 presentations, and one disclosed a pricing response (Table 9.1)."));
C.push(caption("Table 9.1: Company statements on the effect of the November 2023 measures"));
C.push(table(["Company", "Statement (verbatim)", "Implication"], [
  ["Bajaj Finance (Q3 FY24)", "“RBI increased risk weights on consumer credit exposure from 100% to 125% which had an impact of 290 bps on the Company's CRAR. Adjusted for this change CRAR would have been 26.77%.”", "Capital ratio −2.9 pp"],
  ["Bajaj Finance (Q3 FY24)", "“Given the increase in risk weights and higher incremental cost of funds, the Company has increased rates across all portfolios by 20–30 bps.”", "Lending rates +0.20 to 0.30 pp (attributed to both risk weights and funding costs)"],
  ["SBI Cards (Q3 FY24)", "“CRAR impacted by ~400 bps due to increase in risk weight by RBI.”", "Capital ratio about −4 pp"],
], [2000, 5160, 2200]));
C.push(note("Source: company investor presentations for the quarter ended 31 December 2023."));
C.push(P("These statements show that the higher risk weights immediately reduced the capital ratios of NBFCs with large consumer books. This is a separate channel from H1, which measures bank credit, but it is consistent with the sharp fall in the consumer lenders' share prices on the day of the circular."));
C.push(Hd("9.2 Cost of funds and bank borrowing", HeadingLevel.HEADING_2));
C.push(caption("Table 9.2: Change in cost of funds and bank share of borrowings, before vs after the policy"));
const frow = (lbl, F) => [lbl, f2(F.all.mean_change_nonhfc), f2(F.all.mean_change_hfc), f2(F.all.diff), p3(F.all.p), p3(F.excl_mfi.p), p3(F["paired_Non-HFC NBFC"].p), p3(F.paired_HFC.p)];
C.push(table(["Variable", "Change, NBFCs", "Change, HFCs", "Difference", "p (Welch)", "p excl. MFIs", "p paired, NBFCs", "p paired, HFCs"],
  [frow("Cost of funds (pp)", FC), frow("Bank share of borrowings (pp)", FB)], [2160, 1050, 1000, 1000, 950, 1000, 1100, 1100]));
C.push(note(`Firm averages before (to Q2 FY2024) and after (Q3 FY2024 to Q4 FY2025). Cost of funds: ${FC.all.n_nonhfc} NBFCs, ${FC.all.n_hfc} HFCs. Bank share: ${FB.all.n_nonhfc} NBFCs (Poonawalla does not disclose), ${FB.all.n_hfc} HFCs.`));
C.push(...fig("f_firm_dumbbell.png", 600, 276, "Figure 9.1: Each firm's average cost of funds and bank share of borrowings, before and after the circular",
  "Open circles: average up to Q2 FY2024; filled circles: average from Q3 FY2024 to Q4 FY2025. Orange: NBFCs; blue: HFCs. Firms are sorted by group and by the size of the change."));
C.push(P("Figure 9.1 shows the firm-level data behind Table 9.2. Every firm's cost of funds rose, NBFCs and HFCs alike, and the size of the rise does not separate the two groups. The bank share of borrowings is more mixed: it rose for every HFC and for several NBFCs, such as SBI Cards, Mahindra & Mahindra Financial and Muthoot Finance, but fell for Bajaj Finance, Cholamandalam and Shriram Finance. With such differences within each group, a difference between the group averages would need to be large to be statistically significant."));
C.push(P(`Cost of funds rose for both groups after the policy: by ${f2(FC.all.mean_change_nonhfc)} percentage points for the NBFCs and by ${f2(FC.all.mean_change_hfc)} points for the exempt HFCs. Both increases are significant within each group (paired t-test ${pe(FC["paired_Non-HFC NBFC"].p)} and ${p3(FC.paired_HFC.p)}), but the NBFCs' increase was not larger than the HFCs' (${pe(FC.all.p)}). The rise therefore reflects the general increase in interest rates, as bank loans linked to lending benchmarks repriced with a lag, rather than a policy effect specific to NBFCs (Figure 9.2).`));
C.push(...fig("fig4_cof_halfyear.png", 560, 238, "Figure 9.2: Average cost of funds by half-year, NBFCs versus HFCs",
  "Simple averages across firms. Dashed line: between FY24 H1 and FY24 H2 (the circular was issued in November 2023)."));
C.push(P(`The bank share of borrowings also rose for both groups, by ${f2(FB.all.mean_change_nonhfc)} points for NBFCs (paired ${pe(FB["paired_Non-HFC NBFC"].p)}) and ${f2(FB.all.mean_change_hfc)} points for HFCs (paired ${pe(FB.paired_HFC.p)}). The NBFCs' rise was smaller, which is consistent with banks lending relatively less to NBFCs, but the difference is not significant (${pe(FB.all.p)}). The pattern differs a lot between firms: SBI Cards, Mahindra & Mahindra Financial and Muthoot Finance borrowed more from banks after the policy, while Bajaj Finance, Cholamandalam and Shriram Finance borrowed less (Figure 9.3).`));
C.push(...fig("fig5_bank_share.png", 560, 252, "Figure 9.3: Average bank share of borrowings, non-HFC NBFCs versus HFCs",
  "From the borrowing-mix slides of quarterly investor presentations. Solid lines: all firms (Shriram enters in FY23, AAVAS is missing in FY26, IIFL is annual only). Dashed lines: firms observed in every half-year from FY22 to FY25."));
C.push(Hd("9.3 Bajaj Finance", HeadingLevel.HEADING_2));
C.push(P(`Bajaj Finance is the largest NBFC consumer lender in the sample and reports assets by business segment. Its consumer segments (sales finance and personal loans) made up ${R.audit.bajaj_share_series.FY24Q2} per cent of assets in September 2023, ${R.audit.bajaj_share_series.FY24Q4} per cent in March 2024 and ${R.audit.bajaj_share_series.FY25Q4} per cent in March 2025. This dip cannot be attributed to the risk weight: on 15 November 2023, one day before the circular, the RBI barred Bajaj Finance from lending under its “eCOM” and “Insta EMI Card” products, and lifted the restriction on 2 May 2024. The temporary fall in the consumer share matches the period of that restriction.`));
C.push(...fig("f_bajaj_share.png", 560, 224, "Figure 9.4: Bajaj Finance, share of consumer segments covered by the circular in assets under management",
  "Urban and rural sales finance and personal-loan (B2C) segments as a share of total AUM, by quarter. Source: Bajaj Finance investor presentations."));
C.push(P("Figure 9.4 shows that the consumer share of Bajaj Finance's book had moved within a range of about 35 to 39 per cent since 2020. It fell during the embargo period and recovered after the restriction was lifted, returning to about 38 per cent by 2026, even though the higher consumer-credit risk weight stayed in place. This suggests that, for the largest consumer NBFC, the risk weight did not lead to a lasting shift away from consumer lending, although one firm cannot show what happened across the sector."));
C.push(P("Taken together, the firm-level data confirm that the risk weights cut consumer NBFCs' capital ratios, but they do not show a clear effect on NBFCs' funding costs or borrowing mix. With 16 firms and large differences between them, these comparisons have limited power."));

// ================= Chapter 10
C.push(Hd("10. Discussion", HeadingLevel.HEADING_1));
C.push(Hd("10.1 Interpretation", HeadingLevel.HEADING_2));
C.push(P(`The two sources of evidence point the same way. In the bank credit data, bank lending to NBFCs (excluding HFCs) grew about ${ann(-W2.diff)} percentage points a year more slowly than total bank credit after the circular, significant at the 5 per cent level once seasonal patterns are controlled for, and the gap narrowed again after the RBI withdrew this measure in April 2025. In the share-price data, NBFC shares fell on the first trading day after the circular and rose on the first trading day after the rollback, by ${pc(-EP("E1", NH)["[0]"].car)} and ${pc(EP("E2", NH)["[0]"].car)} per cent relative to the market, while the shares of the exempt HFCs did not move on either day. This mirror-image reaction to the tightening and its reversal is the strongest evidence in the study that the bank-to-NBFC measure mattered (H2).`));
C.push(P(`For the consumer-credit channel (H1), targeted bank consumer credit grew about ${ann(-W1.diff)} percentage points a year more slowly than exempt household credit, significant at the 10 per cent level in the simple tests and at the 1 per cent level with Newey–West standard errors. The result comes mainly from personal loans. The share prices of the three consumer lenders fell ${pc(-EP("E1", CL)["[0,+2]"].car)} per cent relative to the market over three days, and two of them reported that the rule cut their capital ratios by 3 to 4 percentage points. Together this is consistent with H1, though the credit-data evidence on its own is weaker than for H2.`));
C.push(P("The firm-level data do not show how NBFCs adjusted beyond the capital impact. Funding costs and the share of bank borrowing changed in similar ways for NBFCs and HFCs, and the clearest case of an NBFC's consumer lending slowing (Bajaj Finance) coincides with a separate RBI restriction."));
C.push(Hd("10.2 Limits to the interpretation", HeadingLevel.HEADING_2));
C.push(P([B("The RBI acted because credit was growing fast. "), "The circular targeted the fastest-growing parts of the credit system. The tests assume that, without the policy, the pre-policy gap in growth rates would have continued. Fast-growing credit tends to slow on its own, so part of the measured slowdown may have happened anyway. The placebo and trend tests find no sign of a slowdown before November 2023, but the H1 gap did drift down slightly, so the H1 estimate is best read as an upper bound."]));
C.push(P([B("A package, not only a capital charge. "), "The circular combined higher risk weights with a requirement to set board-approved limits on unsecured consumer credit by 29 February 2024. Banks had ample capital (CRAR of 16.8 per cent in September 2023), so for most banks the risk weights are unlikely to have made capital requirements bind. Part of the effect may reflect the exposure limits and the signal of the RBI's concern rather than the cost of capital. Testing whether banks with less spare capital cut targeted lending more would need bank-level data and is left for future work."]));
C.push(P([B("Rising credit risk. "), "Delinquencies on credit cards and small personal loans rose during 2024, and the RBI's Financial Stability Reports flagged growing stress and write-offs in unsecured retail loans (RBI, 2024a, 2024b). Banks would have tightened such lending on that ground alone, so part of the H1 effect, and the continued weakness of card credit after 2025, may reflect credit-quality concerns. This affects the credit data more than the event study, since share prices reacted on the day of the announcement."]));
C.push(P([B("Other news around the event dates. "), "Bajaj Finance's product restriction was announced on 15 November 2023, the day before the circular, and the RBI eased risk weights on banks' microfinance loans on the same day as the 2025 rollback. Both are handled by showing results without the firms concerned. The event study measures what investors expected the measures to do, not what they actually did to lending."]));
C.push(P([B("Supply or demand. "), "Lower bank lending to NBFCs is consistent with banks lending less and with NBFCs borrowing less, since NBFCs' own consumer and microfinance lending also slowed in 2024. The share-price reaction on the announcement day, before any change in demand, supports the view that investors expected a cost to NBFCs. Data on NBFCs' bond and commercial-paper issuance would help separate the two."]));
C.push(P([B("Who was covered. "), "The bank-to-NBFC risk weight applied only where the rating-based risk weight was below 100 per cent, and excluded HFCs and priority-sector on-lending. The RBI's NBFC series includes infrastructure, public-sector and microfinance lenders for whom the increase was small or did not apply, so the H2 estimate is diluted and the effect on fully covered borrowers is likely to be larger."]));
C.push(Hd("10.3 Relation to the synopsis hypotheses", HeadingLevel.HEADING_2));
C.push(P("The synopsis proposed five hypotheses. Following the guide's advice the formal tests focus on two; the evidence on the others is limited:"));
C.push(bullet([B("H1 (unsecured exposure and loan growth): "), "bank credit to the targeted unsecured categories slowed relative to exempt credit (significant at 10 per cent in the simple tests), and consumer lenders' share prices fell sharply on the circular. NBFCs' own consumer lending is observed only for Bajaj Finance, where it is affected by a separate RBI restriction."]));
C.push(bullet([B("H2 (bank dependence and borrowing costs): "), "bank lending to NBFCs slowed (significant at 5 per cent with seasonal controls) and recovered after the rollback, and NBFC shares moved in opposite directions on the two announcements. Firm-level funding costs rose for NBFCs and HFCs alike, and within the NBFC group the share-price reaction was not related to bank dependence."]));
C.push(bullet([B("H3 (capital): "), "two consumer NBFCs reported capital-ratio cuts of about 3 to 4 percentage points. The two large equity raises in the sample were set in motion before the circular (Bajaj Finance's board approved its ₹10,000 crore raise on 5 October 2023 and opened the qualified institutional placement on 6 November 2023; Cholamandalam's placement opened on 28 September 2023), so they were not responses to the policy, although the new capital will have softened its impact."]));
C.push(bullet([B("H4 (shift to secured lending): "), `bank gold loans grew by ${f1(gtG[POST])} per cent a year after the policy (${f1(gtG[PRE])} per cent before), but this coincides with a sharp rise in gold prices and the reclassification of agricultural gold loans, so the evidence for a policy-driven shift is weak.`]));
C.push(bullet([B("H5 (recovery after February 2025): "), `bank credit to NBFCs recovered after the April 2025 rollback (${f1(gtNB[ROLL])} per cent a year), and NBFC shares rose ${pc(EP("E2", NH)["[0]"].car)} per cent relative to the market on the announcement, which supports a quick recovery in the bank-funding channel.`]));
C.push(Hd("10.4 Policy implications", HeadingLevel.HEADING_2));
C.push(bullet("Targeted risk weights appear to slow bank credit to the targeted segments while exempt categories keep growing, so they can address a specific risk without a general tightening."));
C.push(bullet("The bank-to-NBFC measure acted quickly in both directions, in the credit data and in share prices, so it can be used and withdrawn as conditions change."));
C.push(bullet("Because the circular combined risk weights with exposure limits at a time of ample bank capital, regulators and analysts should treat such packages as signals as well as capital measures."));
C.push(bullet("For NBFC managers, dependence on bank funding is a regulatory risk that investors price immediately; a diversified funding base (bonds, external commercial borrowing, securitisation) reduces it."));
C.push(Hd("10.5 Limitations", HeadingLevel.HEADING_2));
C.push(bullet("The RBI data measure bank credit, not NBFCs' own lending."));
C.push(bullet("Each credit test rests on 45 to 46 monthly observations, so results that are significant at 5 or 10 per cent should not be over-stated."));
C.push(bullet("The event study covers 16 listed firms; the market reaction of smaller, unlisted NBFCs, which may depend more on banks, is not observed."));
C.push(bullet("Firm definitions of cost of funds and bank borrowing differ, so firm-level comparisons are based on each firm's own change over time."));
C.push(bullet("Some firm series have breaks: Shriram Finance is included from FY2023, after the merger of Shriram Transport Finance and Shriram City Union Finance; Poonawalla Fincorp from FY2022, after the change of control of Magma Fincorp; IIFL's bank share is annual only."));
C.push(bullet("Following the guide's advice, synopsis hypotheses H3 to H5 are addressed only descriptively."));

// ================= Chapter 11
C.push(Hd("11. Conclusion", HeadingLevel.HEADING_1));
C.push(P(`This dissertation examined the RBI's November 2023 increase in risk weights on consumer credit and on bank lending to NBFCs, and the partial rollback of February 2025. In RBI data on bank credit, lending to NBFCs (excluding HFCs) grew about ${ann(-W2.diff)} percentage points a year more slowly than total bank credit after the circular, significant at the 5 per cent level once seasonal patterns are controlled for, and recovered after the rollback. Targeted bank consumer credit grew about ${ann(-W1.diff)} percentage points a year more slowly than exempt household credit, significant at the 10 per cent level. In share prices, the ten NBFCs fell ${pc(-EP("E1", NH)["[0]"].car)} per cent relative to the market on the first trading day after the circular and rose ${pc(EP("E2", NH)["[0]"].car)} per cent after the rollback, while the exempt HFCs did not react on either day; both reactions are significant at the 1 per cent level.`));
C.push(P("The evidence therefore supports both hypotheses: the measures slowed bank credit to the targeted segments, and investors saw them as a cost for the lenders they targeted. Company disclosures confirm that the higher risk weights cut consumer NBFCs' capital ratios, but the firm-level data do not show a clear effect on funding costs or borrowing mix. Because the circular combined higher risk weights with exposure limits at a time of ample bank capital and rising credit risk, the results describe the effect of the package as a whole. Future work could use bank-level data to test whether banks with less spare capital cut targeted lending more, extend the analysis to smaller, unlisted NBFCs, and use bond and commercial-paper issuance data to see whether NBFCs replaced bank funding with market funding."));

// ================= Bibliography
C.push(Hd("Bibliography", HeadingLevel.HEADING_1));
for (const r of [
  "Acharya, V. V., Khandwala, H., & Öncü, T. S. (2013). The growth of a shadow banking system in emerging markets: Evidence from India. Journal of International Money and Finance, 39, 207–230.",
  "Aiyar, S., Calomiris, C. W., & Wieladek, T. (2014). Does macro-prudential regulation leak? Evidence from a UK policy experiment. Journal of Money, Credit and Banking, 46(s1), 181–214.",
  "Basten, C. (2020). Higher bank capital requirements and mortgage pricing: Evidence from the counter-cyclical capital buffer. Review of Finance, 24(2), 453–495.",
  "Behn, M., Haselmann, R., & Wachtel, P. (2016). Procyclical capital regulation and lending. Journal of Finance, 71(2), 919–956.",
  "Bernanke, B. S., & Blinder, A. S. (1988). Credit, money, and aggregate demand. American Economic Review, 78(2), 435–439.",
  "Brown, S. J., & Warner, J. B. (1985). Using daily stock returns: The case of event studies. Journal of Financial Economics, 14(1), 3–31.",
  "Business Standard (2025a, 25 February). NBFCs get breather as RBI restores risk weights on loans from banks.",
  "Business Standard (2025b, 4 February). Gold loan portfolio of banks jumps 71.3% to Rs 1.72 trillion till Dec 2024.",
  "Business Today (2023, 16 November). RBI increases risk weight on consumer credit exposure of banks, NBFCs.",
  "Fama, E. F., Fisher, L., Jensen, M. C., & Roll, R. (1969). The adjustment of stock prices to new information. International Economic Review, 10(1), 1–21.",
  "Irani, R. M., Iyer, R., Meisenzahl, R. R., & Peydró, J.-L. (2021). The rise of shadow banking: Evidence from capital regulation. Review of Financial Studies, 34(5), 2181–2235.",
  "Jiménez, G., Ongena, S., Peydró, J.-L., & Saurina, J. (2017). Macroprudential policy, countercyclical bank capital buffers, and credit supply: Evidence from the Spanish dynamic provisioning experiments. Journal of Political Economy, 125(6), 2126–2177.",
  "Kashyap, A. K., & Stein, J. C. (2004). Cyclical implications of the Basel II capital standards. Federal Reserve Bank of Chicago Economic Perspectives, 28(1), 18–31.",
  "Khwaja, A. I., & Mian, A. (2008). Tracing the impact of bank liquidity shocks: Evidence from an emerging market. American Economic Review, 98(4), 1413–1442.",
  "MacKinlay, A. C. (1997). Event studies in economics and finance. Journal of Economic Literature, 35(1), 13–39.",
  "Newey, W. K., & West, K. D. (1987). A simple, positive semi-definite, heteroskedasticity and autocorrelation consistent covariance matrix. Econometrica, 55(3), 703–708.",
  "Reserve Bank of India (2004). Risk weight on housing loans and consumer credit. Circular RBI/2004/309, 23 December 2004.",
  "Reserve Bank of India (2005). Prudential norms on capital adequacy: Risk weight on commercial real estate exposures. Circular, 26 July 2005.",
  "Reserve Bank of India (2023a). Regulatory measures towards consumer credit and bank credit to NBFCs. Circular RBI/2023-24/85, 16 November 2023.",
  "Reserve Bank of India (2023b). Financial Stability Report, December 2023.",
  "Reserve Bank of India (2024a). Financial Stability Report, June 2024.",
  "Reserve Bank of India (2024b). Financial Stability Report, December 2024.",
  "Reserve Bank of India (2025a). Review of risk weights on microfinance loans. Circular RBI/2024-25/119, 25 February 2025.",
  "Reserve Bank of India (2025b). Exposures of scheduled commercial banks to NBFCs: Review of risk weights. Circular RBI/2024-25/120, 25 February 2025.",
  "Reserve Bank of India. Database on Indian Economy (DBIE), Table 15: Deployment of Gross Bank Credit by Major Sectors (accessed 2026).",
  "Investing.com. Historical daily prices of NSE-listed equities and the Nifty 50 and Nifty 500 indices (accessed 2026).",
  "Company investor presentations, Q1 FY2021 to Q4 FY2026: AAVAS Financiers, Arman Financial Services, Bajaj Finance, Can Fin Homes, Cholamandalam Investment & Finance, CreditAccess Grameen, Home First Finance, IIFL Finance, LIC Housing Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, PNB Housing Finance, Poonawalla Fincorp, Repco Home Finance, SBI Cards and Payment Services, Shriram Finance. Company announcements on the RBI restrictions on Bajaj Finance (15 November 2023; lifted 2 May 2024) and IIFL Finance (4 March 2024; lifted 19 September 2024), and on the Bajaj Finance (5 October 2023) and Cholamandalam (28 September 2023) capital raises.",
]) C.push(new Paragraph({ spacing: { after: 120, line: 300 }, indent: { left: 720, hanging: 720 }, children: [new TextRun(r)] }));

// ================= Appendices
C.push(Hd("Appendices", HeadingLevel.HEADING_1));
C.push(Hd("Appendix A: Data Files", HeadingLevel.HEADING_2));
C.push(P("All data used in this dissertation are provided in the accompanying workbook, and every test is reproduced in a second workbook, with formulas where the test can be done with standard spreadsheet functions and with the full output otherwise. The analysis was carried out in Python (pandas, statsmodels and SciPy)."));
C.push(Hd("Appendix B: Firm-Level Data Notes", HeadingLevel.HEADING_2));
for (const t of [
  "Cost of funds is the company's reported cost of borrowing where available. For Mahindra & Mahindra Financial Services, Shriram Finance and Arman Financial Services it is computed as annualised quarterly finance cost divided by average borrowings from the same presentation; for Cholamandalam it is finance cost as a share of average assets (the only consistently reported measure).",
  "Where a later presentation restates an earlier quarter, the value as first reported is used.",
  "Muthoot Finance's FY2021–22 figures are derived from reported year-to-date averages.",
  "Bank share of borrowings follows each company's own borrowing-mix definition (for example, banks and financial institutions for Mahindra & Mahindra; term loans for Shriram and AAVAS). Cholamandalam's institutional financial investors, reported separately from FY2024, are added back to banks for consistency.",
  "IIFL Finance's bank share is available only annually on a consistent definition; PNB Housing's earlier figures are rescaled to exclude direct assignment.",
]) C.push(bullet(t));
C.push(Hd("Appendix C: Additional Tables", HeadingLevel.HEADING_2));
const prow5 = (ek, pn, lbl) => { const d = E[ek].NIFTY500.portfolios[pn]; return [lbl, ...["[0]", "[0,+1]", "[0,+2]", "[-1,+1]"].flatMap((w) => [pcs(d[w].car), p3(d[w].p)])]; };
C.push(caption("Table C.1: Event study with the Nifty 500 as the market index"));
C.push(table(["Event / portfolio", ...phead.slice(1)], [
  prow5("E1", NH, "Circular: NBFCs"), prow5("E1", HF, "Circular: HFCs"), prow5("E1", DNH, "Circular: NBFCs minus HFCs"), prow5("E1", DCL, "Circular: consumer minus HFCs"),
  prow5("E2", NH, "Rollback: NBFCs"), prow5("E2", HF, "Rollback: HFCs"), prow5("E2", DNH, "Rollback: NBFCs minus HFCs")], pw));
const shortCo = (c) => String(c).replace(/ Ltd\.?| Limited/g, "").replace("Investment & Finance Company", "Inv. & Fin.").replace(" Company India", "").replace("Financial Services", "Fin. Services");
const appTab = (T, cap, src, fmt) => { C.push(caption(cap));
  C.push(table(T.columns, T.data.map((r) => r.map((v, i) => (i === 0 ? shortCo(v) : v === "" ? "n.a." : fmt(v)))),
    [2760, ...Array(T.columns.length - 1).fill(Math.floor(6600 / (T.columns.length - 1)))])); C.push(note(src)); };
appTab(R.app_cof, "Table C.2: Cost of funds by company, annual average of quarterly values (per cent)", "Financial years ending March. Source: quarterly investor presentations (see Appendix B for definitions).", f2);
appTab(R.app_bsh, "Table C.3: Bank share of borrowings by company, annual average of quarterly values (per cent)", "Financial years ending March. Poonawalla Fincorp does not disclose its borrowing mix. Source: quarterly investor presentations.", f1);
C.push(caption("Table C.4: Outstanding bank credit by category at end-March (₹ lakh crore)"));
C.push(table(R.app_rbi.columns.map((c) => (c === "Date" ? "Month" : SHORT[c] || c)), R.app_rbi.data.map((r) => r.map((v, i) => (i === 0 ? v : f2(v)))),
  [1200, ...Array(R.app_rbi.columns.length - 1).fill(Math.floor(8160 / (R.app_rbi.columns.length - 1)))]));
C.push(note("Source: RBI DBIE Table 15 (reported series). Figures from July 2023 include the effect of the HDFC merger (housing up, HFC and NBFC credit down)."));


// ---------------------------------------------------------------- lists of figures, tables and abbreviations (front matter)
const PAGES = fs.existsSync("outputs/final/caption_pages.json") ? JSON.parse(fs.readFileSync("outputs/final/caption_pages.json", "utf8")) : {};
const lp = (t) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9350, leader: "dot" }], spacing: { after: 60, line: 276 },
  indent: { left: 1150, hanging: 1150 }, children: [new TextRun({ text: `${t}\t${PAGES[t.split(":")[0]] || ""}`, size: 21 })] });
const LISTS = [Hd("List of Figures", HeadingLevel.HEADING_1), ...CAPS.filter((t) => t.startsWith("Figure")).map(lp),
  Hd("List of Tables", HeadingLevel.HEADING_1), ...CAPS.filter((t) => t.startsWith("Table")).map(lp),
  Hd("List of Abbreviations", HeadingLevel.HEADING_1),
  table(["Abbreviation", "Meaning"], [
    ["AR", "Abnormal return"], ["AUM", "Assets under management"], ["bps", "Basis points (one hundredth of a percentage point)"],
    ["CAR", "Cumulative abnormal return"], ["CRAR", "Capital to risk-weighted assets ratio"], ["DBIE", "Database on Indian Economy (RBI)"],
    ["FY", "Financial year (April to March)"], ["HFC", "Housing finance company"], ["IL&FS", "Infrastructure Leasing & Financial Services"],
    ["MCLR", "Marginal cost of funds based lending rate"], ["MFI", "Microfinance institution"], ["NBFC", "Non-banking financial company"],
    ["NSE", "National Stock Exchange of India"], ["OLS", "Ordinary least squares"], ["PL", "Personal loans"], ["pp", "Percentage points"],
    ["PSL", "Priority-sector lending"], ["QIP", "Qualified institutional placement"], ["RBI", "Reserve Bank of India"],
    ["RWA", "Risk-weighted assets"], ["SD", "Standard deviation"], ["SHG", "Self-help group"], ["VIF", "Variance inflation factor"],
  ], [2400, 6960])];
C.splice(LISTS_AT, 0, ...LISTS);
FRONT_END += LISTS.length;

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
