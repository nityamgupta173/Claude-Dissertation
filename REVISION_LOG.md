# Revision log: response to the audit

This log covers the revised dissertation (`Dissertation.pdf` / `Dissertation.docx`), produced by the updated `05_final_analysis.py` and `06_build_dissertation.js`. The pre-revision numbers are kept in `outputs/final/results_v1_before_audit.json`.

## 0. Two corrections the audit did not ask for, which change the headline numbers

| # | Finding | Effect |
|---|---|---|
| N1 | The RBI Excel already contains the RBI's **merger-adjusted ("excluding merger") series** for non-food credit, housing and other personal loans, for Jul-2023 to Jun-2025. They sit in unlabelled rows under each labelled row; RBI note 5 says "figures in parentheses exclude the impact of the merger". These series now replace the reported growth rates from Aug-2023 to Jun-2025, which covers the whole post-policy estimation window. | This fixes B8 properly. All aggregate estimates change. |
| N2 | **Other personal loans (an H1 *treated* series) were also hit by the merger**: +2.4 pp growth in Jul-2023, against −0.3 pp in the ex-merger series. The old code excluded the merger month only for housing, NBFC, HFC and non-food credit. | The merger month is now excluded for other personal loans as well. |
| N3 | **The Bibliography cited the wrong Feb-2025 circular.** It cited RBI/2024-25/119, titled "…NBFCs and microfinance loans". RBI actually issued two circulars that day: **/119** on microfinance loans and **/120** on bank exposures to NBFCs. | Both are now cited correctly, and Table 2.1 is corrected. |

Because of N1 and N2, several numbers in the audit (e.g. ₹1.74 / ₹1.93 lakh crore, p ≈ 0.055) are superseded by the re-estimated values below. The audit's methods were applied, not its numbers.

## 1. Verdict on each audit item

| Item | Verdict | What was done |
|---|---|---|
| A1 Post dating | **Valid (text error only).** The Nov-2023 observation is dated 17-Nov-2023, and `POST = "2023-11-30"` already coded it as pre-policy. ANOVA arithmetic confirmed: 0.662² × 32·14/46 = 4.27. | Fixed the text in 5.2/6.1 and Table 5.2. Shortfalls now use 14 post-policy months. |
| A2 t vs normal p-values | **Valid.** | Every HAC, HC1 and panel test now uses t-distribution p-values (`use_t=True`; `debiased=True` for the panel). |
| A3 Model 2 coefficient | **Valid.** | One value from a single source, everywhere: now −1.07. |
| A4 Section 7.1 | **Valid.** | Rewritten with new p-values: drop other PL p = 0.200, drop education p = 0.055. |
| A5 Figure 8.2 | **Valid.** Coefficients ×30 confirmed: pre 0.36/0.40/0.33/0.13/0.07; post 0.27–0.50. | Rewritten: no break at the policy date; consistent with the repricing lag. |
| A6 COVID placebo | **Valid.** | Placebo now starts Jul-2021 (Apr-2021 shown for comparison). Pre-trend test, event study and trend-projection counterfactual added (Section 9.4, Fig 9.1). |
| A7 Shriram source | **Partly wrong.** The value is correct: 25.58% term loans at 30-Sep-2023 appears in Shriram's own **Q2 FY24** deck, not only as a comparative in Q2 FY25. | Citation fixed. No re-estimation needed. |
| A8 "Hand-collected" | **Valid.** | Wording changed throughout. An automated re-check of every value against its cited deck was added (Section 4.3). The manual spot-check is set up but marked [VERIFY] for the student. |
| A9 Title | **Valid.** | New title: "…on **Bank Consumer Credit and Bank Funding of NBFCs in India, 2021–2026**". |
| A10 Gold-loan control | **Valid.** Gold reclassification and price surge verified. | Removed from Table 7.2; kept as a note and no longer cited as support. |
| A11 "Robust once seasonality controlled" | **Valid.** | Wording softened everywhere. |
| A12 Panel not independent | **Valid.** Same point estimate by construction. | Now described as "the same comparison as a panel". |
| B1 Endogeneity / mean reversion | **Valid and material.** Projecting the Jul-21 pre-trend forward cuts H1 from −0.75 to −0.34. | Now the main caveat (10.2, 9.4, 7.2, abstract). |
| B2 Bundle / ample capital | **Valid.** Circular wording and CRAR 16.8% verified. | Added to 2.2, 2.3, 10.2, abstract, conclusion. |
| B3 Delinquencies | **Valid** (qualitative fact verified; specific ratios marked [VERIFY]). | Added to 4.1 and 10.2. |
| B4 Bajaj embargo | **Valid.** Dates verified: 15-Nov-2023 to 2-May-2024. | Disclosed. Model 5 re-run without the embargo quarters (still n.s.). "Continued to grow its consumer book" claim dropped. |
| B5 IIFL embargo | **Valid.** Dates verified: 4-Mar-2024 to 19-Sep-2024. | Model 4 re-run without IIFL (unchanged). |
| B6 Equity raises | **Valid.** Bajaj board approval 5-Oct-2023, QIP opened 6-Nov-2023; Chola QIP opened 28-Sep-2023. Both precede the circular. | H3 text rewritten. |
| B7 Gold loans / H4 | **Valid** (verified). | H4 qualified as weak evidence in 4.1 and 10.3. |
| B8 Contamination / merger | **Valid.** | Core benchmark added (Table 8.1). Merger handled with RBI ex-merger series (N1). |
| B9 Misclassification | **Valid.** PSL exemption verified in the circular. | Model 3 re-specified as a triple difference. Model 4 run without MFIs, and also without Muthoot. Dilution noted. |
| B10 Supply vs demand | **Valid.** | Discussed (8.1, 10.2). Issuance data not collected (see §6). |
| B11 Model 4 narrative | **Valid, with one correction.** The claim "HFC credit slowed 7.1→4.8" was an artefact of the old pre-window ending in Oct-2023. On the consistent window it is 4.1→4.8, flat. | Narrative rewritten (8.2, 10.3 H5, abstract). |
| B12 COF ~ NonHFC×Post | **Valid.** It was already estimated but never reported (−0.18, p = 0.292). | Now reported in Table 8.4. |
| B13 Unbalanced panel | **Valid.** | Balanced-panel lines added to Fig 8.3, with a note. |
| B14 Bajaj segments | **Valid.** Gold loans are folded into Rural B2C and LAS classification is uncertain. | Classification table added (Table 8.3). |
| B15 Table 8.2 framing | **Valid.** | Reworded. |
| C1 Lag choice | **Valid.** Six lags were not pre-registered. | Main specification switched to the 3-lag rule of thumb; 6 and 12 lags kept in the sensitivity table. |
| C2 Small-sample inference | **Valid.** | Added fixed-b (KVB, b = 1), quarterly averages, month dummies and block bootstrap. X-13 not run (see §6). |
| C3 Consistent language | **Valid in principle; the audit's wording is superseded.** After N1/N2, H1 became *stronger* (fixed-b p = 0.028), so "suggestive" understates it. H2 became *weaker* (main p = 0.072). | Wording follows the new evidence (see §2). |
| P4 Claims | **Valid.** | Abstract, executive summary, 10.1, 10.4 and 11 rewritten. |
| P5 Minor | **Valid.** | Chapters reordered (framework → methodology → results → common diagnostics chapter 9). VIF(Post) removed. Title years aligned. Source of the original dataset flagged [VERIFY]. Dataset correction removed from 1.4. Literature additions made (Basten verified; 2004–05 precedent partly verified). Jarque–Bera caveat added. |

## 2. Re-estimated results (old → new)

| Quantity | Old | New |
|---|---|---|
| H1 δ (pp/month), ITS | −0.66 | **−0.75** |
| H1 annualised | −7.9 pp | **−9.0 pp** |
| H1 p (main) | 0.002 (NW6, normal) | **0.005** (NW3, t) |
| H1 panel δ / p | −0.66 / 0.008 | −0.75 / 0.011 |
| H1 placebo TS (Jul-21 window) | +0.45, p 0.175 (Apr-21 window) | **−0.09, p 0.820** (Apr-21 window: +0.31, p 0.455) |
| H1 placebo panel | +0.46, p 0.231 | −0.09, p 0.844 |
| H1 robustness: drop cards / drop other PL / drop housing / drop vehicle / drop education | −0.79 (<0.001) / −0.53 (0.164) / −0.71 (0.006) / −0.81 (0.008) / −0.47 (0.078) | −0.86 (<0.001) / −0.64 (0.200) / −0.69 (0.028) / −0.95 (0.006) / −0.61 (0.055) |
| H1 shorter pre / shorter post | −0.78 (<0.001) / −0.67 (0.013) | −0.85 (0.002) / −0.77 (0.016) |
| H1 + gold control | −1.75 (in table) | removed (note only: −1.81) |
| H1 shortfall | 15 months, ₹1.9 lakh crore | **14 months, 11.1%, ₹2.0 lakh crore** (₹0.9 lakh crore if the Jul-21 pre-trend continued) |
| H2 δ | −0.995 (shown as −1.00 / −0.99) | **−1.07** |
| H2 annualised | −11.9 pp | **−12.8 pp** |
| H2 p (main) | 0.049 | **0.072** |
| H2 core benchmark (new) | — | −1.23, p 0.057 |
| H2 placebo | +0.27, p 0.659 | −0.13, p 0.844 (Jul-21 window) |
| H2 vs HFC / all NBFC | −1.41 (0.109) / −0.63 (0.130) | −1.41 (0.171) / −0.70 (0.134) |
| H2 rollback | +0.95, p 0.186 | +0.98, p 0.228 |
| H2 shortfall | 15 months, ₹2.1 lakh crore | **14 months, 16.1%, ₹2.1 lakh crore** |
| Model 1 R² / F p / SS_reg | 0.04 / 0.181 / 4.27 | 0.05 / 0.128 / 5.51 |
| Model 2 R² / F p / SS_reg | 0.04 / 0.180 / 9.55 | 0.05 / 0.150 / 11.03 |
| Model 2 constant | 0.89 | 0.88 |

The full inference table (Table 9.5) gives the post-policy p-value under each method:

| Method | H1 p | H2 p |
|---|---|---|
| Classical OLS | 0.128 | 0.150 |
| White HC1 | 0.057 | 0.105 |
| NW 1 / 3 (main) / 6 / 12 lags | 0.046 / 0.005 / 0.001 / <0.001 | 0.089 / 0.072 / 0.041 / 0.020 |
| Fixed-b (KVB) | 0.028 | 0.190 |
| Month dummies: OLS / HC1 / NW3 / fixed-b | 0.069 / 0.071 / 0.004 / 0.025 | 0.047 / 0.033 / 0.009 / 0.071 |
| Quarterly averages (n = 14) | 0.157 | 0.189 |
| Block bootstrap / with month dummies | 0.072 / 0.041 | 0.150 / 0.036 |

The new identification checks (Section 9.4) found:
- **Pre-trend slope (pp per month, per year):**
  - H1: +0.23 (p 0.54) from Apr-21; −0.37 (p 0.28) from Jul-21.
  - H2: +0.75 (p 0.16) from Apr-21; +0.22 (p 0.67) from Jul-21.
- **δ if the pre-trend had continued:**
  - H1: −1.19 (Apr-21 trend) or **−0.34** (Jul-21 trend).
  - H2: −2.48 (Apr-21) or −1.73 (Jul-21).
- **Event-study joint test of the leads:**
  - HAC Wald rejects (p < 0.001 for both). This test is unreliable with 10 restrictions and about 45 observations.
  - Classical F does not reject: p 0.51 (H1) and 0.44 (H2).
  - Excluding the Apr-21 block: p 0.96 (H1) and 0.70 (H2).

The firm-level re-estimates:
- **Model 3:** the triple difference is −0.48 pp per 30-pp share (p 0.145). The HFC gradient is +0.57 pp per 30 pp (p 0.057). COF ~ NonHFC×Post is −0.18 (p 0.292).
- **Model 4:** −3.29 (p 0.280) in the main sample. It is −3.41 without IIFL, −3.35 without MFIs, −3.50 without IIFL and MFIs, and −4.04 also without Muthoot (p 0.28–0.40). It is −5.15 (p 0.056) without SBI Cards, a post-hoc exclusion and not relied on.
- **Model 5:** 0.58 (p 0.749), and 1.22 (p 0.430) without the embargo quarters.

The source re-check found:
- 652 values in total.
- 88 computed by formula and 90 read from charts (not text-searchable).
- 434 of the 474 checkable values (91.6%) found verbatim in the cited deck.

## 3. Change log (by location)

| # | Location | Old | New | Reason |
|---|---|---|---|---|
| 1 | Title (cover, certificate, declaration) | "…on Consumer Credit and Bank Funding of Indian NBFCs, FY2021–FY2026" | "…on Bank Consumer Credit and Bank Funding of NBFCs in India, 2021–2026" | A9, P5 |
| 2 | Abstract | Led with p = 0.002 / 0.049; "both measures reduced…"; "NBFCs absorbed the shock…" | Rewritten: H2 marginal and reverses after rollback; H1 significant only with autocorrelation-robust errors and sensitive to pre-trend; package caveat; firm evidence inconclusive | P4, C3, B1, B2 |
| 3 | Executive summary | 1.9 / 2.1 lakh crore; "result holds with seasonality"; Bajaj "absorbed" | New numbers (2.0 / 2.1), new significance language, package caveat, Bajaj embargo | A1, A11, C3, B4 |
| 4 | 1.1 | Two measures | Adds the board-approved exposure-limit requirement (deadline 29 Feb 2024) and PSL exemption | B2 |
| 5 | 1.3 | H1 "credit to the targeted consumer categories" | "bank credit…"; note that NBFC consumer lending is observed only for Bajaj | A9 |
| 6 | 1.4 | Three contributions incl. dataset correction | Two contributions; correction moved to data note (4.3) | P5 |
| 7 | 1.5 | Old chapter order | Framework (5) → Methodology (6) → Results (7, 8) → Diagnostics/inference (9) → Discussion (10) → Conclusion (11) | P5 |
| 8 | Table 2.1 | Feb-2025: "microfinance treatment eased"; NBFC exemptions "HFCs and PSL" | Separate NBFC retail row; exposure-limit row; core investment companies added; two Feb-2025 circulars | B2, N3 |
| 9 | 2.3 | No capital context | System CRAR 16.8% (Sep 2023) and signalling caveat | B2 |
| 10 | 2.4 | — | Dilution / intent-to-treat note (rating < 100% condition, PSUs, MFIs) | B9 |
| 11 | 3.1 / 3.3 / 3.4 | — | Basten (2020); India 2004–05 risk-weight precedent ([VERIFY] for consumer-credit dates); Kiefer–Vogelsang (2005) | P5, C2 |
| 12 | 4.1 | "merger … level shift does not affect other months"; Nov obs not discussed | Reporting-Friday dating; ex-merger series used; OPL also affected | A1, B8, N1, N2 |
| 13 | Table 4.1 | Pre period Apr-21 to Oct-23; housing 13.9→11.2; HFC 7.1→4.8; NBFC 23.5→9.1, etc. | Pre period to 17-Nov-23; merger-adjusted: housing 14.6→15.8; HFC 4.1→4.8; NBFC 24.0→9.1; cards 23.2→11.3; OPL 21.2→10.5; non-food 13.5→11.8; new core-benchmark row | A1, B8, N1 |
| 14 | 4.1 text | Card slowdown "still subject to higher risk weight"; gold growth = "shift toward secured credit" | Delinquency caveat; gold growth attributed mainly to prices and reclassification | B3, B7 |
| 15 | 4.2 / 4.3 | "hand-collected" | "AI-assisted extraction … logged to source page"; automated re-check results; manual spot-check [VERIFY]; original dataset source [VERIFY] | A8, P5 |
| 16 | Table 4.2 note | "Shriram from its Q2 FY2025 presentation" | Q2 FY2024 presentation (25.58%), repeated in Q2 FY2025 | A7 |
| 17 | Table 4.4 | Pre period to Oct-23 | Pre period to 17-Nov-23 (matches the regressions) | A1 |
| 18 | Table 5.2 (was 6.2) | Post = 1 "from end-November 2023" | From the 29-Dec-2023 observation; 17-Nov obs pre-policy | A1 |
| 19 | Model 3 (5.5, Tables 5.2, 5.3) | COF ~ Post × bank share (all firms) | Triple difference with NonHFC; β₂ exemption test | B9, B12 |
| 20 | 6.1 / 6.2 | "first post-policy observation is the end of November 2023"; NW 6 lags | 17-Nov obs pre; 29-Dec first post; n_pre/n_post stated; NW 3 lags, t-distribution; panel "not independent"; placebo from Jul-21 | A1, A2, A6, A12, C1 |
| 21 | 6.6 (new) | — | Rule for describing significance | C3 |
| 22 | Table 7.1 | ITS −0.66 (p 0.002); panel p 0.008; placebo +0.45 (0.175) / +0.46 (0.231) | ITS −0.75 (0.005); panel −0.75 (0.011) labelled same comparison; placebo −0.09 (0.820) / −0.09 (0.844); Apr-21 placebo +0.31 (0.455) | N1, A2, A6, A12, C1 |
| 23 | Ch 7 text | "estimates support H1"; "targeted categories were accelerating" | COVID explanation; quarterly gap path; pointer to pre-trend and inference sections | A6, P4 |
| 24 | Table 7.2 / 7.1 text | Gold row (−1.75); "drop other PL … about 10 per cent level" | Gold row removed (note only); drop other PL p 0.200 not significant; drop education p 0.055 | A4, A10 |
| 25 | 7.2 | 15 months; 10.4%; ₹1.9 lakh crore | 14 months; 11.1%; ₹2.0 lakh crore; trend-projected alternative ₹0.9 lakh crore | A1, B1 |
| 26 | Table 8.1 | Main −1.00 (0.049); placebo +0.27 (0.659); vs HFC p 0.109; rollback +0.95 (0.186) | Main −1.07 (0.072); core benchmark −1.23 (0.057); placebo −0.13 (0.844) + Apr-21 version; vs HFC p 0.171; rollback +0.98 (0.228) | N1, A2, A3, B8, C1 |
| 27 | 8.1 text | "supports H2 … p = 0.049"; ₹2.1 lakh crore over 15 months | "marginal"; 14 months, 16.1%, ₹2.1 lakh crore; supply-vs-demand paragraph | A1, A2, B10, C3 |
| 28 | 8.2 disclosures | "confirm the mechanism behind H1" | Bit on NBFC capital ratios, separate channel; Bajaj rate rise also attributed to funding costs | B15, B4 |
| 29 | 8.2 COF + Fig 8.2 | "effect builds up after the policy" | Triple difference and exemption test reported; no break; V-shape and repricing lag | A5, B9, B12 |
| 30 | 8.2 bank share + Fig 8.3 | "non-HFC NBFCs drifting away" | Placebo half of the estimate; subsamples; balanced panel; HFCs moving toward banks; SBI Cards counter-example | B5, B9, B11, B13 |
| 31 | 8.2 Bajaj + new Table 8.3 | "continued to grow its consumer book" | Embargo disclosed and excluded; segment classification table | B4, B14 |
| 32 | Table 8.4 (was 8.3) | 4 tests | 8 tests incl. triple difference, exemption test, subsamples, embargo exclusion | B5, B9, B12, B4 |
| 33 | Ch 9 (new, from old 7.3 / 8.3–8.5) | Separate model blocks per chapter; VIF(Post) row; NW6 | Combined model summary / ANOVA / coefficients; VIF(Post) removed; JB low-power caveat; Table 9.5 with fixed-b, quarterly, t-based p; new 9.4 pre-trends and event study (Table 9.6, Fig 9.1) | P5, C1, C2, A6, B1 |
| 34 | 10.1 (was 9.1) | "Both hypotheses supported"; NBFCs "absorbed the shock" | Suggested framing applied with new numbers; firm evidence inconclusive | P4, C3 |
| 35 | 10.2 (new) | — | Threats: endogeneity/mean reversion, package/capital headroom, delinquencies, embargoes, dilution, supply vs demand, contamination | B1–B10 |
| 36 | 10.3 (was 9.2) | H3 equity raises as policy responses; H4 gold = shift; H5 "continued to reduce reliance" | Raises predate circular; H4 weak; H5 descriptive | B6, B7, B11 |
| 37 | 10.4 (was 9.3) | "effective tool" | "consistent with…; size uncertain"; signalling caveat | P4 |
| 38 | Ch 11 (was 10) | p = 0.002 / 0.049; "both measures reduced…" | Rewritten per P4; future work: bank-level capital-headroom test, issuance data | P4, B2, B10 |
| 39 | Bibliography | RBI/2024-25/119 mis-titled | /119 and /120; added Basten 2020, Kiefer–Vogelsang 2005, RBI 2005 CRE circular, FSRs Dec-2023/Jun-2024/Dec-2024, Business Standard 2025, company announcements | N3, P5 |
| 40 | Appendix A | — | Extraction_Recheck, Manual_Spot_Check sheets | A8 |
| 41 | Captions Fig 7.1, 7.2, 8.3 | "July 2023 merger month set to zero"; "Hand-collected" | Merger-adjusted note; 2023Q4 colouring explained; balanced/unbalanced note | N1, A8, B13 |

## 4. Items marked [VERIFY] and what was found

| Item | Status | Source |
|---|---|---|
| Circular wording: sectoral limits, 29-Feb-2024 deadline, PSL/HFC/CIC exemptions | **Verified, stated as fact** | [RBI/2023-24/85](https://www.rbi.org.in/Scripts/NotificationUser.aspx?Id=12567&Mode=0) |
| Feb-2025 rollback and microfinance circulars | **Verified.** The bibliography error was corrected. | [RBI/2024-25/120](https://rbi.org.in/Scripts/NotificationUser.aspx?Id=12787); [RBI/2024-25/119](https://www.rbi.org.in/Scripts/NotificationUser.aspx?Id=12786) |
| System CRAR ~16–17% | **Verified:** 16.8% in Sep-2023 | [RBI FSR Dec-2023 press release](https://rbi.org.in/scripts/BS_PressReleaseDisplay.aspx?prid=57005) |
| Card / personal-loan delinquencies rising in 2024 | **Qualitative claim verified.** The specific ratios are still [VERIFY] in the text, because they came only from secondary summaries. | [Moneylife](https://moneylife.in/article/rbi-flags-sharp-increase-in-delinquencies-borrower-indebtedness-in-consumer-credit/75997.html); [RBI FSR Jun-2024 page](https://www.rbi.org.in/Scripts/PublicationReportDetails.aspx?UrlPage=&ID=1284) |
| Bajaj eCOM / Insta EMI Card embargo | **Verified:** 15-Nov-2023 to 2-May-2024 | [Business Standard](https://www.business-standard.com/amp/companies/news/bajaj-finance-to-resume-lending-via-ecom-insta-emi-card-as-rbi-lifts-curbs-124050201221_1.html) |
| IIFL gold-loan embargo | **Verified:** 4-Mar-2024 to 19-Sep-2024 | [Business Today](https://www.businesstoday.in/amp/latest/corporate/story/reserve-bank-lifts-ban-on-iifl-finances-gold-loan-business-ahead-of-festive-season-446661-2024-09-19) |
| Bajaj ₹10,000 cr raise timing | **Verified:** board approval 5-Oct-2023; QIP opened 6-Nov-2023 | [Business Standard](https://www.business-standard.com/amp/companies/news/bajaj-finance-to-raise-rs-10-000-crore-through-qip-preferential-issue-123100500770_1.html); [Goodreturns](https://www.goodreturns.in/news/2-bajaj-stocks-in-focus-bajaj-finance-launches-rs-8-800-crore-qip-fixes-floor-price-details-here-1310881.html) |
| Cholamandalam raise timing | **Verified:** QIP opened 28-Sep-2023 (announced Jul-2023) | [BusinessWorld](https://businessworld.in/article/cholamandalam-investment-initiates-qip-at-rs-120051-per-share-floor-price-493227) |
| Gold-loan surge: prices and agri→retail reclassification | **Verified** | [Business Standard, Feb-2025](https://www.business-standard.com/industry/banking/gold-loan-portfolio-of-banks-jumps-71-3-to-rs-1-72-trillion-till-dec-2024-125020401498_1.html) |
| Bank loans to NBFC-MFIs / on-lending are PSL-eligible | **Verified** | [RBI PSL Master Directions](https://www.rbi.org.in/Scripts/BS_ViewMasDirections.aspx?id=12799) |
| Muthoot's PSL share of bank borrowing | **Not found** in its decks; left as [VERIFY] in the text | — |
| Basten (2020) | **Verified:** RoF 24(2), 453–495 | [Review of Finance](https://revfin.org/higher-bank-capital-requirements-and-mortgage-pricing-evidence-from-the-counter-cyclical-capital-buffer/) |
| India 2004–05 risk weights | **Housing 50→75% (Dec-2004) and CRE 100→125% (Jul-2005) verified.** Consumer-credit and capital-market dates left as [VERIFY]. | [RBI CRE circular](https://www.rbi.org.in/Scripts/NotificationUser.aspx?Id=2391) |
| LAS inside or outside "consumer credit" | **Not verified**; left as [VERIFY] in Table 8.3 | — |
| Manual spot-check of 40 extracted values | **For the student to do** ([VERIFY] in 4.3) | `Dissertation_Data.xlsx` → Manual_Spot_Check |
| Source of the original `nbfc_quarterly_data-v3.csv` | **Only the student knows**; left as [VERIFY] in 4.3 | — |

## 5. Numbers that changed (complete list)

**H1:**
- δ −0.66 → −0.75; annualised −7.9 → −9.0; SE 0.22 → 0.26; p 0.002 → 0.005.
- Panel δ −0.66 → −0.75 (SE 0.25 → 0.29, p 0.008 → 0.011, N 229 → 228).
- Placebos: +0.45 / +0.46 → −0.09 / −0.09 (p 0.175 / 0.231 → 0.820 / 0.844; N 31/154 → 29/143).
- New Apr-21 placebo: +0.31 (0.455).
- Robustness rows: all re-estimated (§2); gold row removed.
- Shortfall: 15 → 14 months, 10.4% → 11.1%, ₹1.9 → ₹2.0 lakh crore.

**H2:**
- δ −0.995 (−1.00 / −0.99) → −1.07; annualised −11.9 → −12.8; SE 0.50 → 0.58; p 0.049 → 0.072.
- Placebo +0.27 (0.659) → −0.13 (0.844).
- vs HFC p 0.109 → 0.171; all NBFC −0.63 (0.130) → −0.70 (0.134).
- Rollback +0.95 (0.186) → +0.98 (0.228).
- Shortfall: 15 → 14 months, 16.1% → 16.1% (larger δ, fewer months), ₹2.1 → ₹2.1 lakh crore.

**ANOVA / model summary:**
- Model 1: R² 0.04 → 0.05, SS_reg 4.27 → 5.51, F 1.85 → 2.40, p 0.181 → 0.128, constant 0.67 → 0.66.
- Model 2: R² 0.04 → 0.05, SS_reg 9.55 → 11.03, F 1.86 → 2.15, p 0.180 → 0.150, constant 0.89 → 0.88, SE of constant 0.43 → 0.45.
- Diagnostics: changed in the second decimal (Table 9.4).

**Inference table:** every entry changed (old table in `results_v1_before_audit.json`); four rows added.

**Table 4.1:**
- Cards 23.3 → 23.2 (pre).
- OPL 21.6 → 21.2 (pre), 10.0 → 10.5 (post).
- Housing 13.9/11.2/9.8 → 14.6/15.8/10.7.
- Vehicle 14.6 → 14.9 (pre).
- Gold 6.8 → 6.7 (pre).
- NBFC ex-HFC 23.5 → 24.0 (pre).
- HFC 7.1 → 4.1 (pre).
- Non-food 13.2/11.4/13.4 → 13.5/11.8/14.2.

**Table 4.4:** pre-period rows re-computed (window to 17-Nov-23).

**Firm level:** Model 3 is re-specified (new numbers above). Models 4 and 5 main estimates are unchanged; new subsample rows added.

## 6. Not done, and why

- **NBFC NCD / CP issuance substitution test (B10):** no issuance data was collected. It needs SEBI / CCIL / RBI issuance data; this is listed as future work.
- **X-13 seasonal adjustment (C2):** the X-13ARIMA-SEATS binary is not available in this environment. Seasonality is handled with calendar-month dummies instead. A year-on-year version isn't informative with 14 post-policy months, because almost every YoY window would mix pre- and post-policy months.
- **Cross-bank capital-headroom test (B2):** needs bank-level data. Listed as future work.
- **Manual spot-check and the source of the original dataset (A8, P5):** these need the student; both are marked [VERIFY] in Section 4.3.
- **Specific delinquency ratios, 2004–06 consumer-credit risk-weight dates, LAS classification, Muthoot PSL share:** could not be confirmed from primary sources; marked [VERIFY].
- **Headline significance language:** the audit's suggested wording ("H2 at the 5% line, H1 suggestive") was **not** adopted verbatim. After the merger correction, H1 is significant under fixed-b inference (p = 0.028) and H2's main p-value is 0.072. The text describes H2 as marginal and H1 as significant only with autocorrelation-robust inference and sensitive to pre-trends.


## 7. Revised text in full

### Abstract

On 16 November 2023 the Reserve Bank of India (RBI) raised the risk weights on two channels of credit: banks' and NBFCs' unsecured consumer credit (personal loans and credit cards), and banks' lending to non-banking financial companies (NBFCs), with housing finance companies (HFCs) exempt. This dissertation asks whether these measures slowed the targeted flows of bank credit. Using monthly RBI data on the sectoral deployment of bank credit, it compares targeted with exempt credit categories before and after the circular (a difference-in-differences design), using the RBI's merger-adjusted series to remove the effect of the HDFC merger. Bank credit to NBFCs (excluding HFCs) grew about 12.8 percentage points a year more slowly than other bank credit after the circular; the estimate is marginally significant (p = 0.072 in the main specification and between 0.033 and 0.071 once calendar-month seasonality is controlled for), and the gap reversed after the RBI withdrew this measure in April 2025, although the rebound is not statistically significant. Targeted bank consumer credit (cards and other personal loans) grew about 9.0 percentage points a year more slowly than exempt household credit. This estimate is consistent in sign across specifications and significant with autocorrelation-robust inference (p = 0.005; fixed-b p = 0.028), but not with classical standard errors (p = 0.128), and it shrinks to 0.34 percentage points a month if a mild pre-policy decline in the gap is assumed to continue. Placebo tests on a fictitious November 2022 date show no effect. Because the circular also required board-approved exposure limits, and because unsecured credit was growing unusually fast and delinquencies were rising, the estimates measure the combined effect of the RBI's November 2023 package rather than a pure capital-cost effect. Firm-level data extracted from 382 investor presentations of 16 listed NBFCs and HFCs show that the higher risk weights cut two consumer lenders' capital ratios by about 3–4 percentage points; evidence on how NBFCs adjusted their funding costs, funding mix and lending is descriptive and inconclusive.

### Executive Summary

Context. On 16 November 2023 the Reserve Bank of India raised the risk weights on unsecured consumer credit (personal loans and credit cards) and on banks' loans to NBFCs (housing finance companies exempt), and required lenders to set board-approved limits on unsecured consumer exposures by 29 February 2024. The bank-to-NBFC risk weight was withdrawn from 1 April 2025; the consumer-credit risk weight was not.

Objective. To measure whether the November 2023 measures slowed bank credit through each targeted channel, and to document how listed NBFCs adjusted.

Data. Monthly RBI data on sectoral deployment of bank credit (January 2019 to July 2026), using the RBI's merger-adjusted series for the period affected by the HDFC merger; and firm-level data on 16 listed NBFCs and HFCs (cost of funds, bank share of borrowings, Bajaj Finance segment assets and company statements), extracted with AI assistance from 382 quarterly investor presentations with every value logged to its source page.

Method. Difference-in-differences: credit categories that faced the higher risk weight are compared with closely related exempt categories before and after the circular. Inference uses Newey–West standard errors with the rule-of-thumb lag length and t-distribution p-values, checked against fixed-b, bootstrap, seasonal and quarterly alternatives; placebo, pre-trend and event-study tests check the comparison groups.

Key findings. • Bank credit to NBFCs (excluding HFCs) grew about 12.8 percentage points a year more slowly than total bank credit after the circular. The effect is marginally significant (p = 0.072; 0.033–0.071 with seasonal controls; up to 0.190 without them) and reversed after the April 2025 rollback, though the rebound is not significant (p = 0.228). • Targeted bank consumer credit grew about 9.0 percentage points a year more slowly than exempt household credit. The sign is consistent across specifications and the effect is significant with autocorrelation-robust inference (p = 0.005; fixed-b p = 0.028), but not with classical OLS (p = 0.128) or quarterly averages (p = 0.157), and it is sensitive to assumptions about pre-policy trends. • Taken at face value, the estimates imply roughly ₹2.0 lakh crore less targeted consumer credit and ₹2.1 lakh crore less bank funding to NBFCs by January 2025 than if pre- policy relative growth had continued; both figures inherit the uncertainty above. • The estimates capture the whole November 2023 package (higher risk weights plus required exposure limits) at a time when banks were well capitalised and delinquencies

1|Page on unsecured loans were rising, so they should not be read as a pure capital-requirement effect. • Firm-level evidence is descriptive and inconclusive: two consumer lenders reported capital-ratio cuts of about 3–4 percentage points, but cost-of-funds, borrowing-mix and segment-growth tests show no significant effect, and the main within-firm case (Bajaj Finance) is confounded by an RBI embargo on two of its lending products. Implications. The evidence is consistent with targeted risk weights slowing bank credit to the targeted segments, most clearly for bank lending to NBFCs, which slowed and then recovered when the measure was withdrawn. The size of the effects is uncertain, and the study cannot separate the capital cost from the supervisory signal of the package.

2|Page

### Section 10.1 (formerly 9.1) Interpretation

Bank credit to NBFCs (excluding HFCs) slowed by about 12.8 percentage points a year relative to total bank credit after the November 2023 circular. The estimate is marginally significant — at the 5 per cent line once seasonality is controlled for, and weaker without — and the gap moved back by a similar amount after the April 2025 rollback, although that rebound is not statistically significant. The timing of the slowdown and the rebound, which match the introduction and withdrawal of the measure, is the strongest piece of evidence that the bank-to-NBFC risk weight affected bank lending to NBFCs.

Targeted bank consumer credit slowed by about 9.0 percentage points a year relative to exempt household credit. The result is consistent in sign across specifications and significant with autocorrelation-robust inference, but not with classical standard errors; it rests mainly on other personal loans; and it is sensitive to whether the pre-policy gap is assumed to be stable or gently declining. It should be read as evidence consistent with the circular slowing bank consumer credit, not as a precise measure of the effect.

The firm-level evidence is descriptive and inconclusive. The disclosures show that the higher risk weights immediately cut the capital ratios of two consumer NBFCs, by about 3–4 percentage points. Beyond that, the tests do not show significant effects on NBFCs' funding costs, borrowing mix or lending mix, and the main within-firm case is confounded by an RBI embargo. The study therefore cannot say whether listed NBFCs absorbed the shock or cut their own lending.

### Chapter 11 (formerly Chapter 10) Conclusion

This dissertation evaluated the RBI's November 2023 increase in risk weights on consumer credit and on bank lending to NBFCs, using RBI sectoral credit data and a difference-in-differences design. Bank credit to NBFCs slowed by about 12.8 percentage points a year relative to total bank credit, an effect that is marginally significant (p = 0.072 in the main specification; at or just below 0.05 with seasonal controls) and that reversed after the April 2025 rollback, although the rebound is not itself significant. Targeted bank consumer credit slowed by about 9.0 percentage points a year relative to exempt household credit, a result that is consistent in sign but significant only with autocorrelation-robust inference and sensitive to assumptions about pre-policy trends. Because the circular combined higher risk weights with required exposure limits at a time of ample bank capital, high credit growth and rising delinquencies, these estimates measure the effect of the whole package rather than of capital requirements alone.

Firm-level evidence from 382 investor presentations confirms that the risk weights cut two consumer NBFCs' capital ratios by about 3–4 percentage points, but tests of NBFCs' funding costs, borrowing mix and lending mix are insignificant and partly confounded, so how NBFCs adjusted remains an open question. Future work could use bank-level data to test whether banks with less capital headroom cut targeted lending more — the decisive test of a capital channel — extend the firm-level data to smaller, unlisted NBFCs, and use bond and commercial-paper issuance data to test whether NBFCs replaced bank funding with market funding.
