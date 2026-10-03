// Builds Dissertation.docx from outputs/final/results.json and figures. Run: node 06_build_dissertation.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, ImageRun, TableOfContents, Footer, PageNumber, LevelFormat, PageBreak,
} = require("docx");

const R = JSON.parse(fs.readFileSync("outputs/final/results.json", "utf8"));
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
const Hd = (text, level) => new Paragraph({ heading: level, spacing: { before: 240, after: 120 }, children: [new TextRun(text)] });
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
const PRE = "Pre-policy (Apr-21 to Oct-23)", POST = "Post-policy (Dec-23 to Mar-25)", ROLL = "After rollback (Apr-25 to Jul-26)";
const fc = R.firm_cof, fb = R.firm_bankshare, bw = R.bajaj_within, cnt = R.counts;
const MG = R.magnitudes, FD = R.firm_desc;
const lakh = (x) => (Number(x) / 100000).toFixed(1);
const fdv = (v, g, p) => FD.find((r) => r.Variable.startsWith(v) && r.Group === g && r.Period.startsWith(p));

// ---------------------------------------------------------------- content
const C = [];

// Title page
C.push(new Paragraph({ spacing: { before: 1800, after: 240 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Capital Requirements and Credit Supply:", bold: true, size: 36 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 },
  children: [new TextRun({ text: "The Effect of the RBI's November 2023 Risk-Weight Measures on Consumer Credit and Bank Funding of Indian NBFCs, FY2021–FY2026", bold: true, size: 28 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "A dissertation submitted in partial fulfilment of the requirements for the degree of", size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Master of Business Administration", bold: true, size: 26 })] }));
for (const t of ["Submitted by: [Student Name]", "Roll / Registration No.: [Number]", "Under the guidance of: [Guide Name, Designation]", "[Department / School], [Institution]", "[Month Year]"])
  C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: t, size: 24 })] }));
C.push(pageBreak());

// Declaration & acknowledgement placeholders
C.push(Hd("Declaration", HeadingLevel.HEADING_1));
C.push(P("I hereby declare that this dissertation titled “Capital Requirements and Credit Supply: The Effect of the RBI's November 2023 Risk-Weight Measures on Consumer Credit and Bank Funding of Indian NBFCs, FY2021–FY2026” is my original work carried out under the guidance of [Guide Name]. It has not been submitted elsewhere for any degree or diploma. [Add any statement on the use of software tools or AI assistance required by your institution's policy.]"));
C.push(P("[Signature]                                                                                     [Date]", { align: AlignmentType.LEFT }));
C.push(Hd("Acknowledgements", HeadingLevel.HEADING_1));
C.push(P("[Student to write: thanks to the guide, faculty, family and others who supported the work.]"));
C.push(pageBreak());

// Abstract
C.push(Hd("Abstract", HeadingLevel.HEADING_1));
C.push(P(`On 16 November 2023 the Reserve Bank of India (RBI) raised capital requirements on two channels of credit at once: the risk weight on unsecured consumer credit (personal loans and credit cards) and the risk weight on banks' lending to non-banking financial companies (NBFCs), with housing finance companies (HFCs) exempt. This dissertation asks whether these targeted increases in capital requirements reduced the supply of credit through each channel, and how listed NBFCs adjusted. Using monthly RBI data on the sectoral deployment of bank credit (January 2019 to July 2026), I compare credit categories hit by the higher risk weights with exempt categories before and after the policy. Growth of bank credit to the targeted consumer categories fell by ${f2(-H1.its.change)} percentage points per month relative to exempt household credit (about ${f1(-h1a)} percentage points a year; p = ${p3(H1.its.p)}), and bank credit to NBFCs (excluding HFCs) slowed by ${f2(-H2.its.change)} percentage points per month relative to banks' total non-food credit (about ${f1(-h2a)} percentage points a year; p = ${p3(H2.its.p)}). Placebo tests using a fictitious November 2022 policy date show no effect. To trace how NBFCs adjusted, I hand-collected quarterly data from ${cnt.decks} investor presentations of ${cnt.firms} listed NBFCs and HFCs. Two lenders disclose that the rule cut their capital ratios by roughly 3 to 4 percentage points, and one reports raising lending rates. Firm-level tests show bank-dependent NBFCs facing somewhat higher funding costs and non-HFC NBFCs drifting away from bank borrowing relative to exempt HFCs, but these differences are not statistically significant in a sample of 16 firms. The evidence indicates that the RBI's risk-weight measures worked mainly through the banking system: banks cut back targeted consumer credit and lending to NBFCs, while listed NBFCs absorbed the shock through capital, pricing and gradual funding substitution rather than by shrinking their own lending.`));
C.push(P([B("Keywords: "), "risk weights, capital requirements, credit supply, NBFCs, consumer credit, macroprudential policy, difference-in-differences, India"]));
C.push(pageBreak());

C.push(Hd("Table of Contents", HeadingLevel.HEADING_1));
C.push(new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }));
C.push(note("Right-click the table above and choose “Update Field” in Word to refresh page numbers."));
C.push(pageBreak());

// ================= Chapter 1
C.push(Hd("1. Introduction", HeadingLevel.HEADING_1));
C.push(Hd("1.1 Background and motivation", HeadingLevel.HEADING_2));
C.push(P("Non-banking financial companies (NBFCs) have become one of the main channels of retail credit in India. Between 2021 and 2023, unsecured consumer lending — personal loans and credit cards — grew far faster than overall credit, and NBFCs both lent directly to households and borrowed heavily from banks to fund that lending. The Reserve Bank of India (RBI) viewed this combination as a build-up of risk: rapid unsecured lending on the asset side, and growing interdependence between banks and NBFCs on the funding side."));
C.push(P("On 16 November 2023 the RBI responded with a single circular that raised capital requirements on both channels. First, it increased the risk weight on consumer credit — excluding housing, education, vehicle, gold-backed and microfinance loans — from 100 per cent to 125 per cent for banks and NBFCs, and raised risk weights on credit card receivables. Second, it increased by 25 percentage points the risk weight that banks attach to their loans to most NBFCs, while exempting housing finance companies (HFCs). Because a higher risk weight requires a lender to hold more capital against the same loan, both measures made the targeted credit more expensive to supply. In February 2025 the RBI partially reversed course, withdrawing the higher risk weight on bank loans to NBFCs from 1 April 2025, but leaving the consumer-credit measures in place."));
C.push(P("This policy is a useful natural experiment. It changed capital requirements for clearly defined categories of credit at a known date, left closely comparable categories untouched, and was later partly reversed. That design makes it possible to separate the effect of the rule from general credit conditions such as the interest-rate cycle."));
C.push(Hd("1.2 Research problem", HeadingLevel.HEADING_2));
C.push(P("Market commentary since 2023 suggests that unsecured lending slowed after the RBI's measures, but it is not clear how much of the slowdown was caused by the rule itself, which channel it worked through, and how NBFCs responded. The higher risk weights could have reduced NBFCs' own lending, raised their borrowing costs, pushed them away from bank funding, or been absorbed through lower capital buffers and higher lending rates. This study measures the effect on each channel and documents how listed NBFCs adjusted."));
C.push(Hd("1.3 Objectives and hypotheses", HeadingLevel.HEADING_2));
C.push(P("Following the guide's direction to focus on one or two hypotheses, the study tests the two levers of the November 2023 circular:"));
C.push(bullet([B("H1 (consumer-credit channel): "), "Higher risk weights on consumer credit reduced the growth of credit to the targeted consumer categories (credit cards and other personal loans) relative to exempt household credit categories."]));
C.push(bullet([B("H2 (bank-funding channel): "), "Higher risk weights on bank exposures to NBFCs reduced the growth of bank credit to NBFCs (excluding the exempt HFCs) relative to banks' overall lending."]));
C.push(P("These hypotheses correspond to H1 and H2 of the original synopsis, re-specified so that each one is tested on the group that the regulation actually targeted. The synopsis versions — firm-level effects of NBFCs' pre-policy unsecured share and bank dependence — are tested as supporting evidence using a hand-collected firm-level dataset (Chapter 7)."));
C.push(Hd("1.4 Contribution", HeadingLevel.HEADING_2));
C.push(P("The study makes three contributions. First, it provides a clean before-and-after comparison of targeted and exempt credit categories using official RBI data, with placebo tests. Second, it assembles a new firm-level dataset from 382 quarterly investor presentations of 16 listed NBFCs and HFCs, including cost of funds, bank share of borrowings and company disclosures of the policy's capital impact. Third, it shows that an initial firm-level dataset obtained for the project was materially inaccurate, and documents how it was verified and replaced — a practical lesson for empirical work on Indian NBFCs."));
C.push(Hd("1.5 Structure", HeadingLevel.HEADING_2));
C.push(P("Chapter 2 describes the regulatory background. Chapter 3 reviews the literature. Chapter 4 describes the data and the verification process, and Chapter 5 the methodology. Chapters 6 and 7 present the results for H1 and H2 respectively, together with the firm-level evidence. Chapter 8 discusses the findings and limitations, and Chapter 9 concludes."));

// ================= Chapter 2
C.push(Hd("2. Regulatory Background", HeadingLevel.HEADING_1));
C.push(Hd("2.1 Risk weights and capital requirements", HeadingLevel.HEADING_2));
C.push(P("Under the Basel framework applied in India, banks and NBFCs must hold regulatory capital equal to a minimum percentage of their risk-weighted assets (RWA). Each exposure is multiplied by a risk weight; a 125 per cent risk weight means that every ₹100 lent counts as ₹125 of RWA. Raising the risk weight therefore increases the capital a lender must hold against the same loan or, equivalently, lowers its capital-to-risk-weighted-assets ratio (CRAR) if capital does not change. For an NBFC that must keep CRAR above 15 per cent, or a bank that must meet its own capital requirements, a higher risk weight raises the effective cost of supplying that type of credit."));
C.push(Hd("2.2 The November 2023 circular", HeadingLevel.HEADING_2));
C.push(P("The RBI's circular of 16 November 2023 (“Regulatory measures towards consumer credit and bank credit to NBFCs”) contained the measures summarised in Table 2.1."));
C.push(caption("Table 2.1: Main risk-weight changes, November 2023 and February 2025"));
C.push(table(["Measure", "Change", "Exemptions / notes"], [
  ["Consumer credit (banks and NBFCs)", "Risk weight 100% → 125%", "Excludes housing, education, vehicle, gold-backed and microfinance/SHG loans"],
  ["Credit card receivables", "Banks 125% → 150%; NBFCs 100% → 125%", "—"],
  ["Bank exposures to NBFCs", "+25 percentage points where the rating-based risk weight was below 100%", "Excludes HFCs and loans eligible for priority-sector classification"],
  ["February 2025 circular", "Additional risk weight on bank exposures to NBFCs withdrawn from 1 April 2025", "Consumer-credit risk weights unchanged; microfinance treatment eased"],
], [2600, 3000, 3426]));
C.push(note("Source: RBI circulars of 16 November 2023 and 25 February 2025 (see References)."));
C.push(P("Two features of the circular matter for the empirical design. The consumer-credit measure targeted unsecured personal loans and cards but left housing, vehicle and education loans untouched, so exempt household categories provide a natural comparison group for H1. The bank-to-NBFC measure applied to NBFCs but not to HFCs, so bank credit to HFCs and to the banking system as a whole provide comparison groups for H2. The February 2025 rollback of the NBFC measure, with no rollback of the consumer measure, provides a further check on the targeted nature of the effects."));
C.push(Hd("2.3 Macroeconomic setting", HeadingLevel.HEADING_2));
C.push(P("The policy came after the RBI had raised the repo rate by 250 basis points between May 2022 and February 2023, to 6.5 per cent, where it stayed until February 2025. Funding costs across the financial system were therefore already rising during 2022–23 as loans repriced. This is the main reason why a comparison with untargeted credit categories, rather than a simple before-and-after comparison, is necessary."));

// ================= Chapter 3
C.push(Hd("2.4 The NBFC sector before the policy", HeadingLevel.HEADING_2));
C.push(P("NBFCs are regulated by the RBI but, unlike banks, cannot accept demand deposits and do not have access to the RBI's liquidity facilities in the same way. They fund themselves through bank loans, bonds (non-convertible debentures), commercial paper, securitisation and, for some, public deposits. The default of IL&FS in 2018 showed how quickly funding can dry up for NBFCs that depend on short-term market borrowing, and banks subsequently became an even more important source of NBFC funding. In October 2021 the RBI introduced a Scale-Based Regulation framework, placing NBFCs in base, middle, upper and top layers with progressively stricter requirements; most of the firms in this study fall in the middle or upper layers."));
C.push(P(`By October 2023 the two channels targeted by the RBI were large. Outstanding bank credit was ₹${lakh(MG.cards_oct23)} lakh crore to credit cards and ₹${lakh(MG.pl_oct23)} lakh crore in other personal loans, and bank credit to NBFCs excluding HFCs stood at ₹${lakh(MG.nbfc_oct23)} lakh crore — about ${f1(100 * MG.nbfc_oct23 / MG.nonfood_oct23)} per cent of banks' total non-food credit. All three had been growing at more than 20 per cent a year (Table 4.1), well above overall credit growth. The circular was therefore aimed at the fastest-growing parts of the credit system, and at the link between banks and NBFCs that had become central to NBFC funding after 2018.`));
C.push(Hd("3. Literature Review", HeadingLevel.HEADING_1));
C.push(Hd("3.1 Capital requirements and credit supply", HeadingLevel.HEADING_2));
C.push(P("A large literature shows that capital requirements affect the supply of bank credit. Kashyap and Stein (2004) explain how risk-sensitive capital rules can amplify credit cycles. Aiyar, Calomiris and Wieladek (2014) use changes in bank-specific capital requirements in the United Kingdom to show that higher requirements reduce lending, while part of the gap is filled by foreign branches outside the regulator's reach. Jiménez, Ongena, Peydró and Saurina (2017) find that Spain's dynamic provisioning — a capital buffer — dampened credit cycles, with effects concentrated in more constrained banks. Behn, Haselmann and Wachtel (2016) show that when model-based risk weights rose during the 2008 crisis, German banks cut lending more for loans whose risk weights increased. Together these studies suggest that raising risk weights on a category of loans should reduce credit supply to that category, which is the logic behind H1."));
C.push(Hd("3.2 Bank funding and shadow banks", HeadingLevel.HEADING_2));
C.push(P("The bank lending channel literature (Bernanke and Blinder, 1988; Khwaja and Mian, 2008) shows that shocks to banks' ability or willingness to lend are passed on to borrowers who depend on bank credit. Khwaja and Mian (2008) use Pakistani loan-level data to show that firms borrowing from affected banks could not fully replace the lost funding. Irani, Iyer, Meisenzahl and Peydró (2021) show that tighter capital regulation of banks shifts lending activity toward non-bank lenders. In India, Acharya, Khandwala and Öncü (2013) document the close funding links between banks and NBFCs. Raising the risk weight on bank loans to NBFCs is a direct shock to this channel; H2 tests whether it reduced bank credit to NBFCs."));
C.push(Hd("3.3 Macroprudential policy and targeted risk weights", HeadingLevel.HEADING_2));
C.push(P("Higher risk weights on specific loan categories are a macroprudential tool: instead of tightening credit across the economy through interest rates, the regulator raises the cost of supplying the type of credit it considers risky. The studies above suggest three predictions for such a measure. First, credit growth in the targeted category should slow relative to similar untargeted categories. Second, lenders may respond not only by lending less but by raising prices, raising capital or shifting toward exempt lending. Third, part of the targeted activity may move to lenders or funding sources outside the measure's reach (Aiyar, Calomiris and Wieladek, 2014; Irani et al., 2021). The November 2023 circular is unusual because it targeted both a type of borrower (unsecured consumers) and a type of lender-funding link (banks to NBFCs) at the same time, allowing both channels to be examined within one policy."));
C.push(Hd("3.4 Methodological literature", HeadingLevel.HEADING_2));
C.push(P("Policy evaluations of this kind typically compare affected and unaffected groups before and after the change (difference-in-differences). Two inference issues are relevant here. With aggregate monthly series, errors are serially correlated, so standard errors robust to autocorrelation (Newey and West, 1987) or to cross-sectional and serial dependence in panels (Driscoll and Kraay, 1998) are needed. With a small number of firms, conventional cluster-robust standard errors over-reject; the wild cluster bootstrap-t (Cameron, Gelbach and Miller, 2008), with six-point weights for very few clusters (Webb, 2023), is the recommended remedy. Callaway, Goodman-Bacon and Sant'Anna (2024) discuss the interpretation of designs with continuous treatment intensity, relevant to the firm-level exposure tests."));
C.push(Hd("3.5 Research gap", HeadingLevel.HEADING_2));
C.push(P("Existing evidence on the November 2023 measures is largely descriptive (RBI Financial Stability Reports and market commentary). There is little formal evidence that separates the consumer-credit and bank-funding channels, uses exempt categories as controls, or documents how listed NBFCs adjusted using their own disclosures. This study addresses that gap."));

// ================= Chapter 4
C.push(Hd("4. Data", HeadingLevel.HEADING_1));
C.push(Hd("4.1 Aggregate data: RBI sectoral deployment of bank credit", HeadingLevel.HEADING_2));
C.push(P(`The main dataset is the RBI's monthly table “Deployment of Gross Bank Credit by Major Sectors” (Database on Indian Economy, Table 15), covering ${cnt.rbi_months} monthly observations from ${cnt.rbi_first} to ${cnt.rbi_last}. It reports outstanding bank credit (₹ crore) to credit cards, other personal loans, housing, vehicle loans, education, loans against gold jewellery, NBFCs, HFCs and total non-food credit. Bank credit to NBFCs excluding HFCs is computed as NBFC credit minus HFC credit. Growth is measured as the monthly change in the natural logarithm of outstanding credit (×100, approximately per cent per month).`));
C.push(P("One adjustment is required. On 1 July 2023 HDFC Ltd, then India's largest HFC, merged into HDFC Bank. Its home loans moved onto a bank balance sheet and its bank borrowings disappeared, producing a one-off jump in bank housing credit and a fall in bank credit to HFCs and NBFCs in July 2023. The RBI notes that data from 28 July 2023 include the impact of the merger. The July 2023 growth observation is therefore excluded for housing, HFC, NBFC and non-food credit. Because all tests use month-on-month growth, the level shift does not affect other months."));
C.push(caption("Table 4.1: Annualised growth of bank credit by category (per cent per year)"));
const gtRow = (name, o) => [name, f1(o[PRE]), f1(o[POST]), f1(o[ROLL])];
C.push(table(["Category", "Pre-policy (Apr-21 to Oct-23)", "Post-policy (Dec-23 to Mar-25)", "After rollback (Apr-25 to Jul-26)"], [
  gtRow("Credit cards (targeted, H1)", gtCC), gtRow("Other personal loans (targeted, H1)", gtPL),
  gtRow("Housing (exempt)", gtH), gtRow("Vehicle loans (exempt)", gtV), gtRow("Education (exempt)", gtE), gtRow("Gold loans (exempt)", gtG),
  gtRow("NBFCs ex-HFC (targeted, H2)", gtNB), gtRow("HFCs (exempt)", gtHF), gtRow("Total non-food credit (benchmark, H2)", gtNF),
], [3626, 1800, 1800, 1800]));
C.push(note("Average monthly log growth × 12. July 2023 excluded for merger-affected series. Source: RBI DBIE Table 15; author's calculations."));
C.push(P(`Table 4.1 already shows the pattern the hypotheses predict. Bank credit to credit cards grew ${f1(gtCC[PRE])} per cent a year before the policy and ${f1(gtCC[POST])} per cent after; other personal loans slowed from ${f1(gtPL[PRE])} to ${f1(gtPL[POST])} per cent. Exempt categories slowed much less (housing ${f1(gtH[PRE])} → ${f1(gtH[POST])}, vehicle ${f1(gtV[PRE])} → ${f1(gtV[POST])}) or accelerated (education). Bank credit to NBFCs excluding HFCs slowed from ${f1(gtNB[PRE])} to ${f1(gtNB[POST])} per cent, while total non-food credit only eased from ${f1(gtNF[PRE])} to ${f1(gtNF[POST])} per cent. After the April 2025 rollback, bank credit to NBFCs recovered to ${f1(gtNB[ROLL])} per cent, whereas card credit — whose higher risk weight was not reversed — slowed further to ${f1(gtCC[ROLL])} per cent. Bank gold loans, which were exempt, grew very rapidly after the policy (${f1(gtG[POST])} per cent a year), consistent with lenders shifting toward secured credit.`));
C.push(Hd("4.2 Firm-level data: hand-collected from investor presentations", HeadingLevel.HEADING_2));
C.push(P(`The firm-level sample consists of 16 listed NBFCs and HFCs: six HFCs (AAVAS Financiers, Can Fin Homes, Home First Finance, LIC Housing Finance, PNB Housing Finance and Repco Home Finance) and ten NBFCs (Arman Financial Services, Bajaj Finance, Cholamandalam Investment & Finance, CreditAccess Grameen, IIFL Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, Poonawalla Fincorp, SBI Cards and Payment Services, and Shriram Finance). For each company, all quarterly investor presentations from Q1 FY2021 to Q4 FY2026 were collected — ${cnt.decks} documents in total. From these, the following were hand-collected, each value recorded with the presentation and page it came from:`));
C.push(bullet([B("Cost of funds"), ` (${cnt.cof_obs} company-quarters, all 16 firms): the company's reported cost of borrowing; where not reported, annualised finance cost divided by average borrowings from the same presentation.`]));
C.push(bullet([B("Bank share of borrowings"), ` (${cnt.bankshare_obs} company-quarters, ${cnt.bankshare_firms} firms): the share of borrowings from banks shown in each company's borrowing-mix slide (Poonawalla Fincorp does not disclose its mix).`]));
C.push(bullet([B("Pre-policy bank dependence"), ": the bank share of borrowings at 30 September 2023, verified against each Q2 FY2024 presentation."]));
C.push(bullet([B("Bajaj Finance segment-wise assets under management"), " (24 quarters), separating consumer segments affected by the rule from exempt segments."]));
C.push(bullet([B("Company statements"), " quantifying the effect of the RBI's measures on capital ratios and pricing."]));
C.push(Hd("4.3 Data verification", HeadingLevel.HEADING_2));
C.push(P("The project began with a pre-compiled firm-level dataset. Checking it against the companies' own filings showed that several key fields were not reported data. For five lenders, “unsecured” assets were a fixed percentage of total assets in every quarter; for Shriram Finance, the series mixed an assumed 9 per cent share with two quarters of actual personal-loan figures, producing artificial spikes. Bank borrowing shares differed materially from the filings for several firms (for example Arman Financial 71 per cent in the dataset versus 32 per cent in its filing; Muthoot Finance 41 versus 65 per cent). Headline figures such as total assets and gross NPA ratios were generally accurate."));
C.push(P("All firm-level variables used in this dissertation were therefore re-collected from primary sources. An early test using the unverified bank shares appeared to show a significant effect on borrowing costs; with verified shares the effect disappeared. This illustrates why primary-source verification was essential, and the original dataset is not used in any result reported here."));
C.push(caption("Table 4.2: Firm sample and pre-policy bank dependence (30 September 2023)"));
const ver = [
  ["AAVAS Financiers", "HFC", "44.5", "49.6"], ["Arman Financial Services", "NBFC", "71.0", "32.3"], ["Bajaj Finance", "NBFC", "23.0", "23.0"],
  ["Can Fin Homes", "HFC", "54.0", "57.0"], ["Cholamandalam Inv. & Fin.", "NBFC", "53.5", "60.0"], ["CreditAccess Grameen", "NBFC (MFI)", "58.0", "52.9"],
  ["Home First Finance", "HFC", "57.5", "54.0"], ["IIFL Finance", "NBFC", "56.0", "55.3"], ["LIC Housing Finance", "HFC", "24.5", "33.0"],
  ["M&M Financial Services", "NBFC", "56.5", "57.2"], ["Muthoot Finance", "NBFC (gold)", "41.0", "65.4"], ["PNB Housing Finance", "HFC", "40.2", "40.2"],
  ["Poonawalla Fincorp", "NBFC", "64.0", "not disclosed"], ["Repco Home Finance", "HFC", "74.0", "75.5"], ["SBI Cards & Payment", "NBFC (cards)", "68.5", "77.0"],
  ["Shriram Finance", "NBFC", "30.5", "25.6"]];
C.push(table(["Company", "Type", "Bank share in original dataset (%)", "Bank share verified from filing (%)"], ver, [3226, 1600, 2100, 2100]));
C.push(note("Definitions follow each company's borrowing-mix slide (e.g. some include financial institutions). Source: Q2 FY2024 investor presentations; Shriram from its Q2 FY2025 presentation."));
C.push(caption("Table 4.3: Descriptive statistics of hand-collected firm-level variables"));
C.push(table(["Variable", "Group", "Period", "N", "Mean", "SD", "Min", "Max"], FD.map((r) => [r.Variable, r.Group, r.Period, r.N, f2(r.Mean), f2(r.SD), f2(r.Min), f2(r.Max)]), [2126, 1500, 1500, 600, 800, 800, 800, 900]));
C.push(note("Company-quarters, Q1 FY2021 – Q4 FY2026. Pre = up to Q2 FY2024 (September 2023); Post = Q3 FY2024 onward."));
C.push(P(`Table 4.3 summarises the firm-level data. Average cost of funds rose for both groups after the policy — from ${f2(fdv("Cost","Non-HFC NBFCs","Pre").Mean)} to ${f2(fdv("Cost","Non-HFC NBFCs","Post").Mean)} per cent for non-HFC NBFCs and from ${f2(fdv("Cost","HFCs","Pre").Mean)} to ${f2(fdv("Cost","HFCs","Post").Mean)} per cent for HFCs — reflecting the general rise in interest rates. The bank share of borrowings, by contrast, rose by about ${f1(fdv("Bank","HFCs","Post").Mean - fdv("Bank","HFCs","Pre").Mean)} percentage points for HFCs but only ${f1(fdv("Bank","Non-HFC NBFCs","Post").Mean - fdv("Bank","Non-HFC NBFCs","Pre").Mean)} for non-HFC NBFCs. These raw differences motivate the formal tests in Chapter 7.`));

// ================= Chapter 5
C.push(Hd("5. Methodology", HeadingLevel.HEADING_1));
C.push(Hd("5.1 Research design", HeadingLevel.HEADING_2));
C.push(P("The study uses a difference-in-differences design. For each hypothesis, credit that the regulation targeted (the treatment group) is compared with closely related credit that was exempt (the control group), before and after the policy. Any change in the growth gap between the two groups after November 2023 is attributed to the policy, under the assumption that, without the policy, the gap would have stayed at its pre-policy level. Interest-rate changes, economic growth and other common shocks affect both groups and cancel out in the comparison."));
C.push(P("The main sample runs from April 2021 to January 2025, so that the post-policy period ends before the February 2025 rollback was announced. The first post-policy observation is the end of November 2023."));
C.push(Hd("5.2 Aggregate tests", HeadingLevel.HEADING_2));
C.push(P([B("Interrupted time series on the growth gap. "), "For each month t, the growth gap is the average monthly growth of the treated categories minus that of the control categories. The gap is regressed on a constant and a post-policy indicator:"]));
C.push(P([I("Gap(t) = α + δ · Post(t) + ε(t)")], { align: AlignmentType.CENTER }));
C.push(P("α is the average pre-policy gap and δ is the change in the gap after the policy — the difference-in-differences estimate. Standard errors are Newey–West with six lags to allow for autocorrelation and seasonality."));
C.push(P([B("Panel difference-in-differences. "), "As a second specification for H1, each category is treated as a separate series in a panel, with category fixed effects, month fixed effects and a Treated × Post interaction, estimated with the linearmodels PanelOLS estimator and Driscoll–Kraay standard errors (bandwidth six):"]));
C.push(P([I("g(i,t) = μ(i) + λ(t) + δ · Treated(i) × Post(t) + ε(i,t)")], { align: AlignmentType.CENTER }));
C.push(P([B("Placebo tests. "), "Both tests are repeated using only pre-policy data (April 2021 to October 2023) with a fictitious policy date of November 2022. A significant placebo effect would indicate that the groups were already diverging before the policy."]));
C.push(Hd("5.3 Hypothesis definitions", HeadingLevel.HEADING_2));
C.push(caption("Table 5.1: Treatment and control groups"));
C.push(table(["Hypothesis", "Treated (higher risk weight)", "Control (exempt)", "Expected sign of δ"], [
  ["H1 consumer credit", "Credit cards; other personal loans", "Housing; vehicle loans; education", "Negative"],
  ["H2 bank funding", "Bank credit to NBFCs excluding HFCs", "Total non-food bank credit", "Negative"],
  ["H2 alternatives", "Bank credit to NBFCs ex-HFC / all NBFCs", "Bank credit to HFCs / total non-food credit", "Negative"],
], [1900, 2600, 2700, 1826]));
C.push(Hd("5.4 Identifying assumptions and threats", HeadingLevel.HEADING_2));
C.push(P("The key assumption is that, without the policy, the growth gap between treated and control credit would have stayed at its pre-policy average. Four threats are considered. (i) Pre-existing trends: if targeted credit was already slowing relative to controls, the estimate would be biased; the placebo tests and Figures 6.2 and 7.1 check this. (ii) Other shocks at the same time: the interest-rate cycle and economic growth affect treated and control credit alike and are differenced out; the repo rate was unchanged at 6.5 per cent throughout the main post-policy window. (iii) Spillovers: if banks redirected lending from targeted to exempt categories, the control group would grow faster and the estimate would overstate the absolute fall in targeted credit, though it would still measure the policy's effect on the relative allocation of credit. (iv) The HDFC merger in July 2023 changes the level of some series; using month-on-month growth and excluding the merger month removes its effect, and a robustness check drops the housing series entirely."));
C.push(Hd("5.5 Firm-level tests", HeadingLevel.HEADING_2));
C.push(P("Firm-level panels are estimated with two-way fixed effects (company and quarter) using PanelOLS. With 16 firms, inference uses the wild cluster bootstrap-t clustered by company (Webb six-point weights, 19,999 draws, null imposed). Three specifications are estimated: (i) cost of funds on Post × verified pre-policy bank share; (ii) bank share of borrowings on Post × non-HFC indicator, using the regulatory exemption of HFCs as the source of variation; and (iii) for Bajaj Finance, segment-level asset growth on Post × affected-segment indicator, with segment and quarter fixed effects."));

// ================= Chapter 6
C.push(Hd("6. Results: H1 — Consumer-Credit Channel", HeadingLevel.HEADING_1));
C.push(...fig("fig1_rbi_indices.png", 600, 234, "Figure 6.1: Bank credit indices, targeted versus exempt categories (October 2023 = 100)",
  "Source: RBI DBIE Table 15; author's calculations. Indices chain monthly growth; July 2023 merger month set to zero growth for affected series."));
C.push(P("Figure 6.1 (left panel) shows the raw data. Before November 2023, bank credit to cards and personal loans was growing much faster than exempt household credit and was catching up rapidly. After the policy, the two series moved together, and from 2025 targeted credit fell behind."));
C.push(caption("Table 6.1: H1 — effect of the consumer-credit risk weight on targeted credit growth"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "Annualised (pp)", "N"], [
  ["Interrupted time series (Newey–West)", f2(H1.its.change), f2(H1.its.se), p3(H1.its.p), f1(H1.its.annualised_change), H1.its.n],
  ["Panel DiD (Driscoll–Kraay)", f2(H1.panel.beta), f2(H1.panel.se), p3(H1.panel.p), f1(H1.panel.annualised), H1.panel.n],
  ["Placebo Nov-2022, time series", f2(H1.placebo.change), f2(H1.placebo.se), p3(H1.placebo.p), f1(H1.placebo.annualised_change), H1.placebo.n],
  ["Placebo Nov-2022, panel", f2(H1.placebo_panel.beta), f2(H1.placebo_panel.se), p3(H1.placebo_panel.p), f1(H1.placebo_panel.annualised), H1.placebo_panel.n],
], [3226, 1300, 1100, 1100, 1300, 1000]));
C.push(note("Sample April 2021 – January 2025 (placebo: April 2021 – October 2023). δ is the change in the monthly growth gap between treated and control categories after the policy date."));
C.push(P(`The estimates support H1. After the policy, monthly growth of the targeted categories fell by ${f2(-H1.its.change)} percentage points relative to exempt categories (p = ${p3(H1.its.p)}); the panel specification gives the same point estimate with p = ${p3(H1.panel.p)}. Annualised, this is a slowdown of about ${f1(-h1a)} percentage points a year in targeted credit growth relative to exempt credit — large compared with the pre-policy growth rates of 20–23 per cent. The placebo estimates are positive and insignificant (p = ${p3(H1.placebo.p)} and ${p3(H1.placebo_panel.p)}): before the policy, if anything, the targeted categories were accelerating relative to the controls.`));
C.push(...fig("fig2_h1_gap.png", 560, 238, "Figure 6.2: Monthly growth gap, targeted minus exempt consumer credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters. Dashed line: pre-policy average gap."));
C.push(Hd("6.1 Robustness", HeadingLevel.HEADING_2));
C.push(caption("Table 6.2: H1 robustness (panel DiD, Driscoll–Kraay standard errors)"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "N"],
  H1.robustness.map((r) => [r.Specification, f2(r.Coefficient), f2(r.SE), p3(r.p), r.N]), [4226, 1300, 1200, 1200, 1100]));
C.push(P("The effect is negative in every specification. It is significant at the 5 per cent level in most, including when credit cards or the merger-affected housing series are dropped, when gold loans are added to the control group, and with shorter pre- or post-periods. Dropping other personal loans (leaving only credit cards as treated) or dropping education from the control group weakens significance to about the 10 per cent level, reflecting the smaller and more volatile series that remain."));
C.push(Hd("6.2 Economic magnitude", HeadingLevel.HEADING_2));
C.push(P(`To put the estimate in rupee terms, suppose targeted consumer credit had continued to grow at its pre-policy pace relative to exempt credit. Over the ${MG.months} months from November 2023 to January 2025, a monthly shortfall of ${f2(-H1.its.change)} percentage points accumulates to about ${f1(100 * (Math.exp(-H1.its.change * MG.months / 100) - 1))} per cent. With ₹${lakh(MG.h1_outstanding_jan25)} lakh crore of bank credit outstanding to cards and other personal loans in January 2025, this implies roughly ₹${lakh(MG.h1_shortfall)} lakh crore less targeted credit than would otherwise have been extended. This is an approximate figure: it assumes the pre-policy relative growth would have continued, which may overstate the counterfactual given how fast unsecured credit had been growing.`));

// ================= Chapter 7
C.push(Hd("7. Results: H2 — Bank-Funding Channel and Firm-Level Evidence", HeadingLevel.HEADING_1));
C.push(Hd("7.1 Aggregate evidence", HeadingLevel.HEADING_2));
C.push(P("The right panel of Figure 6.1 compares bank credit to NBFCs (excluding HFCs) with total non-food bank credit. Before the policy, bank credit to NBFCs grew much faster than overall bank credit. After November 2023 it grew more slowly, and it accelerated sharply again after the April 2025 rollback."));
C.push(caption("Table 7.1: H2 — effect of the risk weight on bank exposures to NBFCs"));
C.push(table(["Specification", "δ (pp per month)", "Std. error", "p-value", "Annualised (pp)", "N"], [
  ["Main: NBFCs ex-HFC minus total non-food credit", f2(H2.its.change), f2(H2.its.se), p3(H2.its.p), f1(H2.its.annualised_change), H2.its.n],
  ["Placebo Nov-2022", f2(H2.placebo.change), f2(H2.placebo.se), p3(H2.placebo.p), f1(H2.placebo.annualised_change), H2.placebo.n],
  ["Alternative: NBFCs ex-HFC minus HFCs", f2(H2.alt_vs_hfc.change), f2(H2.alt_vs_hfc.se), p3(H2.alt_vs_hfc.p), f1(H2.alt_vs_hfc.annualised_change), H2.alt_vs_hfc.n],
  ["Alternative: all NBFCs (incl. HFCs) minus non-food", f2(H2.alt_all_nbfc.change), f2(H2.alt_all_nbfc.se), p3(H2.alt_all_nbfc.p), f1(H2.alt_all_nbfc.annualised_change), H2.alt_all_nbfc.n],
  ["Rollback: change after Apr-2025 (Dec-23 to Jul-26)", f2(H2.rollback.change), f2(H2.rollback.se), p3(H2.rollback.p), f1(H2.rollback.annualised_change), H2.rollback.n],
], [3426, 1250, 1050, 1050, 1250, 1000]));
C.push(note("Interrupted time series on the monthly growth gap with Newey–West standard errors (6 lags). Main sample April 2021 – January 2025."));
C.push(P(`The main estimate supports H2: after the policy, monthly growth of bank credit to NBFCs (excluding HFCs) fell by ${f2(-H2.its.change)} percentage points relative to total non-food credit (p = ${p3(H2.its.p)}), about ${f1(-h2a)} percentage points a year. The placebo estimate is small, positive and insignificant (p = ${p3(H2.placebo.p)}).`));
C.push(P(`In rupee terms, a monthly shortfall of ${f2(-H2.its.change)} percentage points over ${MG.months} months cumulates to about ${f1(100 * (Math.exp(-H2.its.change * MG.months / 100) - 1))} per cent. With ₹${lakh(MG.h2_outstanding_jan25)} lakh crore of bank credit to NBFCs (excluding HFCs) outstanding in January 2025, this implies roughly ₹${lakh(MG.h2_shortfall)} lakh crore less bank funding to NBFCs than if the pre-policy relationship had continued — a substantial amount relative to the sector's funding needs, which NBFCs had to raise elsewhere or forgo.`));
C.push(P(`Two alternative comparisons point in the same direction but are less precise. Against bank credit to the exempt HFCs, the slowdown is larger (${f1(-H2.alt_vs_hfc.annualised_change)} percentage points a year) but noisier (p = ${p3(H2.alt_vs_hfc.p)}), because the HFC series is small and was disrupted by the HDFC merger. Including HFCs in the treated group dilutes the effect (p = ${p3(H2.alt_all_nbfc.p)}), as expected when exempt borrowers are mixed into the treated group. After the April 2025 rollback the gap widened again in favour of NBFCs by ${f2(H2.rollback.change)} percentage points a month (p = ${p3(H2.rollback.p)}); the rebound is economically large but, with only 16 months of post-rollback data, not yet statistically significant. Bank credit to NBFCs is volatile from month to month because banks lend heavily at quarter-ends, which limits the precision of all H2 estimates.`));
C.push(...fig("fig3_h2_gap.png", 560, 238, "Figure 7.1: Monthly growth gap, bank credit to NBFCs (ex-HFC) minus total non-food credit (quarterly averages)",
  "Grey bars: pre-policy quarters; orange bars: post-policy quarters (including after the April 2025 rollback)."));
C.push(Hd("7.2 How did listed NBFCs adjust? Firm-level evidence", HeadingLevel.HEADING_2));
C.push(P([B("Company disclosures. "), "Two lenders quantified the immediate capital impact of the rule in their December 2023 presentations, and one disclosed a pricing response (Table 7.2)."]));
C.push(caption("Table 7.2: Company statements on the effect of the November 2023 measures"));
C.push(table(["Company", "Statement (verbatim)", "Implication"], [
  ["Bajaj Finance (Q3 FY24)", "“RBI increased risk weights on consumer credit exposure from 100% to 125% which had an impact of 290 bps on the Company's CRAR. Adjusted for this change CRAR would have been 26.77%.”", "Capital ratio −2.9 pp"],
  ["Bajaj Finance (Q3 FY24)", "“Given the increase in risk weights and higher incremental cost of funds, the Company has increased rates across all portfolios by 20–30 bps.”", "Lending rates +0.20–0.30 pp"],
  ["SBI Cards (Q3 FY24)", "“CRAR impacted by ~400 bps due to increase in risk weight by RBI.”", "Capital ratio about −4 pp"],
], [2000, 5226, 1800]));
C.push(note("Source: company investor presentations for the quarter ended 31 December 2023."));
C.push(P("These statements confirm the mechanism behind H1: the higher risk weights immediately reduced the capital ratios of consumer lenders, and at least one passed part of the cost on to borrowers."));
C.push(P([B("Cost of funds. "), `Using ${fc.n} company-quarters for ${fc.firms} firms, NBFCs that were more dependent on bank borrowing before the policy saw their cost of funds rise relative to others: the estimate implies about ${f2(fc.beta * 30)} percentage points higher cost of funds for a firm with 30 percentage points more bank funding. The effect builds up after the policy (Figure 7.2), but it is not statistically significant (wild-bootstrap p = ${p3(fc.p)}), and part of it may reflect the slower repricing of bank loans after the 2022–23 rate rises.`]));
C.push(...fig("fig4_firm_cof_event.png", 560, 245, "Figure 7.2: Cost of funds of bank-dependent versus other NBFCs, by half-year",
  "Coefficients on bank share × half-year, scaled to a 30 pp difference in bank share; reference period FY23 H2; 95% wild-cluster-bootstrap intervals."));
C.push(P([B("Bank share of borrowings. "), `Exempt HFCs increased their reliance on bank borrowing after the policy, while non-HFC NBFCs did not (Figure 7.3). Relative to HFCs, the bank share of non-HFC NBFCs' borrowings fell by ${f1(-fb.beta)} percentage points before the rollback (p = ${p3(fb.p)}) and by about ${f1(-R.firm_bankshare_last.beta)} percentage points by the second half of FY2026 (p = ${p3(R.firm_bankshare_last.p)}). The placebo is insignificant (p = ${p3(R.firm_bankshare_placebo.p)}). The direction is consistent with the aggregate H2 result — banks lent less to NBFCs, so NBFCs relied more on bonds, commercial paper and external borrowing — but with 15 firms the firm-level difference is not statistically significant.`]));
C.push(...fig("fig5_bank_share.png", 560, 245, "Figure 7.3: Average bank share of borrowings, non-HFC NBFCs versus HFCs",
  "Hand-collected from borrowing-mix slides of quarterly investor presentations; definitions follow each company's disclosure."));
C.push(P([B("Within-firm lending: Bajaj Finance. "), `Bajaj Finance is the largest NBFC consumer lender in the sample and reports assets by business segment every quarter. Its consumer segments (sales finance and B2C personal loans) made up ${R.bajaj_share.FY24Q2} per cent of assets in September 2023, dipped to ${R.bajaj_share.FY24Q4} per cent by March 2024 and recovered to ${R.bajaj_share.FY25Q4} per cent by March 2025. Comparing the growth of affected and exempt segments within the firm, there is no significant relative slowdown (${f2(bw.beta)} pp per quarter, p = ${p3(bw.p)}). Bajaj absorbed the capital impact, raised its rates and continued to grow its consumer book.`]));
C.push(caption("Table 7.3: Firm-level tests (hand-collected data)"));
C.push(table(["Test", "Estimate", "p-value", "Obs.", "Firms"], [
  ["Cost of funds ~ Post × pre-policy bank share", `${f2(fc.beta * 30)} pp per 30 pp share`, p3(fc.p), fc.n, fc.firms],
  ["Bank share ~ Post × non-HFC (to Mar-2025)", `${f2(fb.beta)} pp`, p3(fb.p), fb.n, fb.firms],
  ["Bank share, placebo Nov-2022", `${f2(R.firm_bankshare_placebo.beta)} pp`, p3(R.firm_bankshare_placebo.p), "—", fb.firms],
  ["Bajaj: affected vs exempt segment growth", `${f2(bw.beta)} pp per quarter`, p3(bw.p), bw.n, `${bw.segments} segments`],
], [3726, 2100, 1100, 1000, 1100]));
C.push(note("Two-way fixed effects (firm/segment and quarter). Firm tests: wild cluster bootstrap-t by firm. Bajaj: Driscoll–Kraay standard errors."));

// ================= Chapter 8
C.push(Hd("8. Discussion", HeadingLevel.HEADING_1));
C.push(Hd("8.1 Interpretation", HeadingLevel.HEADING_2));
C.push(P("Both hypotheses are supported at the system level. Raising the capital requirement on unsecured consumer credit slowed bank credit to cards and personal loans by about eight percentage points a year relative to exempt household credit. Raising the capital requirement on bank loans to NBFCs slowed bank lending to NBFCs by about twelve percentage points a year relative to total bank credit, and the partial reversal in April 2025 was followed by a rapid recovery in NBFC funding while card credit — still subject to the higher risk weight — stayed weak. These patterns are what a targeted capital-requirement shock would produce, and they are not visible before the policy."));
C.push(P("At the level of listed NBFCs, the adjustment looks different. The rule cut capital ratios immediately (by about 3 to 4 percentage points at the two lenders that disclosed it), and at least one lender raised its lending rates. Bank-dependent NBFCs saw somewhat higher funding costs, and non-HFC NBFCs gradually reduced their reliance on bank borrowing. But large listed NBFCs did not shrink their consumer lending relative to their other lending. These firms are well capitalised, rated AAA or AA, and can raise equity and bond finance: Bajaj Finance and Cholamandalam both raised fresh equity around the time of the policy. The policy therefore appears to have bitten mainly through banks' own lending decisions, while large NBFCs absorbed the shock through capital, pricing and funding substitution."));
C.push(P("This answers the question posed in the synopsis — whether NBFCs reduced overall lending or adjusted in other ways. For the large listed NBFCs studied here, the answer is the latter. The very rapid growth of exempt, secured gold loans in the banking data after the policy is also consistent with lenders shifting toward secured credit."));
C.push(Hd("8.2 Relation to the synopsis hypotheses", HeadingLevel.HEADING_2));
C.push(P("The synopsis proposed five hypotheses. Following the guide's advice the formal tests focus on two, but the evidence speaks to all five:"));
C.push(bullet([B("H1 (unsecured exposure and loan growth): "), "supported at the system level — bank credit to the targeted unsecured categories slowed significantly. Within Bajaj Finance, the largest NBFC consumer lender, consumer segments did not slow relative to other segments."]));
C.push(bullet([B("H2 (bank dependence and borrowing costs): "), "the banking system cut its lending to NBFCs significantly. Firm-level funding costs of bank-dependent NBFCs rose somewhat more than others, but not significantly; the 2022–23 rate cycle dominated funding costs."]));
C.push(bullet([B("H3 (capital): "), "company disclosures show the policy reduced capital ratios of consumer lenders by about 3–4 percentage points; the largest affected lenders responded by raising equity, so the slowdown did not need to come from their lending."]));
C.push(bullet([B("H4 (shift to secured lending): "), `in the banking data, exempt and secured gold loans grew by ${f1(gtG[POST])} per cent a year after the policy, far faster than before (${f1(gtG[PRE])} per cent), consistent with a shift toward secured credit.`]));
C.push(bullet([B("H5 (recovery after February 2025): "), `bank credit to NBFCs recovered quickly in aggregate after the April 2025 rollback (${f1(gtNB[ROLL])} per cent a year), but at the firm level non-HFC NBFCs continued to reduce their reliance on bank borrowing, consistent with the synopsis expectation that the relaxation would not bring an immediate full return to bank funding.`]));
C.push(Hd("8.3 Policy implications", HeadingLevel.HEADING_2));
C.push(bullet("Targeted risk weights are an effective tool for slowing specific segments of bank credit without a broad tightening: exempt categories continued to grow."));
C.push(bullet("The bank-to-NBFC lever transmits quickly and reverses quickly: bank credit to NBFCs slowed after November 2023 and recovered after April 2025."));
C.push(bullet("For NBFC managers, diversified funding (bonds, external commercial borrowing, securitisation) reduced exposure to the bank-funding channel; firms with concentrated bank funding were most exposed to this kind of regulatory change."));
C.push(bullet("Well-capitalised NBFCs can offset part of the intended effect through capital raising and alternative funding; regulators assessing the policy should look at NBFC balance sheets as well as bank credit data."));
C.push(Hd("8.4 Limitations", HeadingLevel.HEADING_2));
C.push(bullet("The RBI data measure bank credit, not NBFCs' own lending; the aggregate tests capture the banking system's response."));
C.push(bullet("The difference-in-differences design assumes that, without the policy, the growth gap between treated and control categories would have remained at its pre-policy average. The placebo tests support this, but it cannot be proven."));
C.push(bullet("Bank credit to NBFCs is volatile from month to month, so H2 is estimated less precisely than H1; alternative comparison groups give p-values of about 0.11–0.13."));
C.push(bullet("The firm-level sample has only 16 firms, which limits statistical power; firm definitions of cost of funds and bank borrowing differ, which is why all firm-level comparisons are within-firm over time."));
C.push(bullet("Some firm-level series have breaks: Shriram Finance is included from FY2023 (after the merger of Shriram Transport Finance and Shriram City Union Finance); Poonawalla Fincorp from FY2022 (after the change of control of Magma Fincorp); Arman Financial's early borrowing figures are partly interpolated."));
C.push(bullet("Following the guide's advice, the study focuses on two hypotheses. Synopsis hypotheses on capital buffers (H3), secured-loan substitution (H4) and the speed of recovery after the 2025 relaxation (H5) are addressed only descriptively."));

// ================= Chapter 9
C.push(Hd("9. Conclusion", HeadingLevel.HEADING_1));
C.push(P(`This dissertation evaluated the RBI's November 2023 increase in risk weights on consumer credit and on bank lending to NBFCs. Using RBI sectoral credit data and a difference-in-differences design, it finds that both measures reduced the targeted flows of bank credit: targeted consumer credit grew about ${f1(-h1a)} percentage points a year more slowly than exempt household credit (p = ${p3(H1.its.p)}), and bank credit to NBFCs grew about ${f1(-h2a)} percentage points a year more slowly than total bank credit (p = ${p3(H2.its.p)}). Placebo tests show no such effects before the policy. A firm-level dataset hand-collected from ${cnt.decks} investor presentations shows how listed NBFCs absorbed the shock — through lower capital ratios, higher lending rates, somewhat higher funding costs and a gradual shift away from bank borrowing — rather than by cutting their own lending. The results suggest that targeted capital requirements are a precise and reversible tool for managing credit risk in specific segments, and that their effect on the NBFC sector depends on how easily NBFCs can replace bank funding.`));
C.push(P("Future work could extend the firm-level dataset to unlisted and smaller NBFCs, which are more likely to be funding-constrained, and use loan-level or bureau data to measure NBFCs' consumer lending directly."));

// ================= References
C.push(pageBreak());
C.push(Hd("References", HeadingLevel.HEADING_1));
for (const r of [
  "Acharya, V. V., Khandwala, H., & Öncü, T. S. (2013). The growth of a shadow banking system in emerging markets: Evidence from India. Journal of International Money and Finance, 39, 207–230.",
  "Aiyar, S., Calomiris, C. W., & Wieladek, T. (2014). Does macro-prudential regulation leak? Evidence from a UK policy experiment. Journal of Money, Credit and Banking, 46(s1), 181–214.",
  "Behn, M., Haselmann, R., & Wachtel, P. (2016). Procyclical capital regulation and lending. Journal of Finance, 71(2), 919–956.",
  "Bernanke, B. S., & Blinder, A. S. (1988). Credit, money, and aggregate demand. American Economic Review, 78(2), 435–439.",
  "Callaway, B., Goodman-Bacon, A., & Sant'Anna, P. H. C. (2024). Difference-in-differences with a continuous treatment. NBER Working Paper No. 32117.",
  "Cameron, A. C., Gelbach, J. B., & Miller, D. L. (2008). Bootstrap-based improvements for inference with clustered errors. Review of Economics and Statistics, 90(3), 414–427.",
  "Driscoll, J. C., & Kraay, A. C. (1998). Consistent covariance matrix estimation with spatially dependent panel data. Review of Economics and Statistics, 80(4), 549–560.",
  "Irani, R. M., Iyer, R., Meisenzahl, R. R., & Peydró, J.-L. (2021). The rise of shadow banking: Evidence from capital regulation. Review of Financial Studies, 34(5), 2181–2235.",
  "Jiménez, G., Ongena, S., Peydró, J.-L., & Saurina, J. (2017). Macroprudential policy, countercyclical bank capital buffers, and credit supply: Evidence from the Spanish dynamic provisioning experiments. Journal of Political Economy, 125(6), 2126–2177.",
  "Kashyap, A. K., & Stein, J. C. (2004). Cyclical implications of the Basel II capital standards. Federal Reserve Bank of Chicago Economic Perspectives, 28(1), 18–31.",
  "Khwaja, A. I., & Mian, A. (2008). Tracing the impact of bank liquidity shocks: Evidence from an emerging market. American Economic Review, 98(4), 1413–1442.",
  "Newey, W. K., & West, K. D. (1987). A simple, positive semi-definite, heteroskedasticity and autocorrelation consistent covariance matrix. Econometrica, 55(3), 703–708.",
  "Reserve Bank of India (2023). Regulatory measures towards consumer credit and bank credit to NBFCs. Circular RBI/2023-24/85, 16 November 2023.",
  "Reserve Bank of India (2025). Risk weights on exposures of scheduled commercial banks to NBFCs and microfinance loans. Circular RBI/2024-25/119, 25 February 2025.",
  "Reserve Bank of India. Database on Indian Economy (DBIE), Table 15: Deployment of Gross Bank Credit by Major Sectors (accessed 2026).",
  "Webb, M. D. (2023). Reworking wild bootstrap-based inference for clustered errors. Canadian Journal of Economics, 56(3), 839–858.",
  "Company investor presentations, Q1 FY2021 – Q4 FY2026: AAVAS Financiers, Arman Financial Services, Bajaj Finance, Can Fin Homes, Cholamandalam Investment & Finance, CreditAccess Grameen, Home First Finance, IIFL Finance, LIC Housing Finance, Mahindra & Mahindra Financial Services, Muthoot Finance, PNB Housing Finance, Poonawalla Fincorp, Repco Home Finance, SBI Cards and Payment Services, Shriram Finance.",
]) C.push(new Paragraph({ spacing: { after: 120, line: 300 }, indent: { left: 720, hanging: 720 }, children: [new TextRun(r)] }));

// ================= Appendices
C.push(pageBreak());
C.push(Hd("Appendix A: Data Files and Reproducibility", HeadingLevel.HEADING_1));
C.push(P("All data used in this dissertation are provided in the accompanying workbook Dissertation_Data.xlsx, which contains the following sheets:"));
for (const t of [
  "README — definitions and sources of all variables.",
  "RBI_Monthly — outstanding bank credit by category and monthly growth rates.",
  "RBI_Growth_Table — the figures in Table 4.1.",
  "COF_Quarterly and BankShare_Quarterly — hand-collected firm data, with the presentation and page each value came from.",
  "Bank_Exposure_Sep2023 — verified pre-policy bank shares and the values in the original dataset.",
  "Bajaj_Segments — segment-wise assets of Bajaj Finance.",
  "Company_Disclosures — verbatim company statements in Table 7.2.",
  "Results_* — all regression results reported in Chapters 6 and 7.",
]) C.push(bullet(t));
C.push(P("The analysis was carried out in Python (pandas, statsmodels and linearmodels). The scripts are 05_final_analysis.py (all results, tables and figures) and did_tools.py (two-way fixed-effects estimation with wild cluster bootstrap-t inference)."));
C.push(Hd("Appendix B: Firm-Level Data Notes", HeadingLevel.HEADING_1));
for (const t of [
  "Cost of funds is the company's reported cost of borrowing where available. For Mahindra & Mahindra Financial Services, Shriram Finance and Arman Financial Services it is computed as annualised quarterly finance cost divided by average borrowings from the same presentation; for Cholamandalam it is finance cost as a share of average assets (the only consistently reported measure).",
  "Where a later presentation restates an earlier quarter, the value as first reported is used.",
  "Muthoot Finance's FY2021–22 figures are derived from reported year-to-date averages.",
  "Bank share of borrowings follows each company's own borrowing-mix definition (for example, banks and financial institutions for Mahindra & Mahindra; term loans for Shriram and AAVAS). Cholamandalam's institutional financial investors (IFI), reported separately from FY2024, are added back to banks for consistency.",
  "IIFL Finance's bank share is available only annually on a consistent definition; PNB Housing's earlier figures are rescaled to exclude direct assignment.",
]) C.push(bullet(t));

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: "[Student Name]", title: "Capital Requirements and Credit Supply",
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: FONT, color: "1F3864" }, paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: FONT, color: "2E4A7A" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync("Dissertation.docx", buf); console.log("Dissertation.docx written", buf.length, "bytes"); });
