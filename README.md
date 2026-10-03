# NBFC risk-weight impact analysis (FY21–FY26)

Run order: `python 01_build_master_panel.py` → `python 02_analysis.py` (needs pandas, linearmodels, scipy, matplotlib, openpyxl).
Everything lands in `master_panel_data_final.csv` and `outputs/`.

## What was done
| Step | Status |
|---|---|
| Screener.in patch | Done, with limits. The Screener quarterly block only covers Mar-2024 to Mar-2026; annual sheets cover FY17–FY26. Filled NA cells only: Net profit +47, Interest expense +128 (these two match the CSV's definitions). **Not patched: Interest income and Operating expenses** – Screener "Sales"/"Expenses" use different definitions (e.g. Bajaj Mar-24: 14,926 vs 13,230) so mixing them would corrupt each firm's series; they are stored as separate `Screener_*` columns. Neither is used in any regression. |
| Macro merge | Repo (RBI, day-weighted quarterly mean), 5Y yield (quarterly mean of monthly), GDP YoY (computed from FRED levels; matches the supplied YoY file exactly). |
| Frozen exposure | `Pre_Policy_Bank_Exposure` = Q2 FY24 `Bank_Borrowing_Share_Pct`, broadcast per firm. |
| Balanced panel | Home First dropped (only 22 of 24 quarters) → 15 firms × 24 quarters. |
| Model | `linearmodels.PanelOLS`, firm + quarter FE, DV = 100 × log difference of AUM. Inference = own implementation of **wild cluster bootstrap-t** (Webb 6-pt weights, null imposed, 19,999 draws), validated against PanelOLS point estimates. |
| Placebo / MFI / PDF Steps 2–7 | Done; see `outputs/table1…`, `table2…`, `table3…` (.md/.tex/.csv), `event_study.png`, `exposure_histogram.png`, `summary_statistics*.csv`, `diagnostics.txt`. |

## ⚠ Critical data finding: the "unsecured AUM" series is mostly assumed, not reported
`outputs/unsecured_series_quality.csv`: for **Bajaj (42%), Cholamandalam (7%), IIFL (8%), M&M (6%), Poonawalla (48%)** the CSV's unsecured AUM is a **constant percentage of total AUM** (correlation of unsecured growth with total growth = 1.000). For SBI Cards it is 100% by definition (cards only). **Shriram** (checked against the four investor presentations you supplied, `data_corrections/shriram_filings.csv`): the CSV uses 9.00% of AUM in 6 of 8 verifiable quarters and actual Personal Loans in the other two (Q3 FY24, Q3 FY25 – these are the two "spikes"); real Personal Loans are only 3.4–4.5% of AUM. Total AUM matches filings (max error 0.7%, Q4 FY25). I therefore set Shriram's Unsecured/Secured AUM to NA (filed Personal Loans kept in `Verified_PersonalLoan_Cr`) and replaced its Total AUM with filed figures.
**Consequence:** the unsecured-growth DV is, for 5 of the 7 remaining firms, simply total-AUM growth. Only Arman has a varying (probably reported) series. **The core DiD below is therefore not a valid test of unsecured lending** until real unsecured/personal-loan AUM is collected for those firms. Treat Tables 1–3 as a mechanical pipeline demonstration, not as findings about H1/H3/H4.

## Results (bank-share exposure; 7 firms with an unsecured series after the Shriram fix)
* **Core DiD**: β = 0.016 (CR1 SE 0.061, wild-bootstrap p = 0.74). Null (but see warning above).
* **Placebo (Nov-2022)**: full-sample β = −0.029, p = 0.55 (null). Pre-policy-only version β = −0.080, wild-bootstrap p = 0.19 (t(6) p = 0.048) – no longer significant on the bootstrap now that the bad Shriram quarters are gone.
* **Event study**: 12 pre-period coefficients jointly zero: F = 0.50, wild-bootstrap p = 0.52 → parallel trends not rejected (very low power).
* **MFI / Feb-2025**: β = −0.42, p = 0.42 – only **3 firms** (Arman, CreditAccess, IIFL); not interpretable. See "MFI data" below.
* **Model 2 (H2, borrowing cost; annual Screener data, real figures, 16 firms)**: β = 0.030, wild-bootstrap p = 0.015 – the one result built on genuine data; supports H2.
* Secured (12 firms) β = 0.009, p = 0.76; total (15 firms) β = −0.038, p = 0.60 (this one uses real total AUM). Capital triple interaction (M3): β₂ = −0.047, p = 0.048 – fragile, 7 clusters.

## MFI data – not collected
The network proxy blocks company/exchange sites (shriramfinance.in, muthootmicrofin.com, bseindia.com, …); only web-search snippets are reachable, which cannot give a reliable 24-quarter AUM series per firm, and I will not invent numbers. To test the Feb-2025 reversal you need genuine quarterly MFI AUM and Q2 FY24 bank-borrowing share for further listed NBFC-MFIs, e.g. **Spandana Sphoorty, Satin Creditcare, Fusion Finance, Muthoot Microfin** (listed Dec-2023). Either upload their investor presentations/results PDFs as you did for Shriram, or allow those domains in the environment's network settings.

## Data issues found while patching
* In the original CSV, Q4 `Net_Profit_PAT_Cr` is the **full-year** profit for several firms (e.g. AAVAS Mar-24: 490.8 vs Screener quarterly 142.6) – 34 overlapping cells disagree. Not used in regressions; fix before using PAT anywhere.
* Shriram `Interest_Expense_Cr` in the CSV appears shifted one quarter versus Screener (Dec-24 shows the Sep-24 figure).
* Several other CSV series look like round-number estimates (e.g. Shriram Q2 FY25 total AUM 243,000 vs filed 243,042.6); audit the CSV before relying on it.

## Things you must know (limitations)
1. **8 clusters, not 16.** Log growth is undefined for the 8 firms with zero unsecured AUM (AAVAS, Can Fin, CreditAccess, LIC HF, Muthoot, PNB HF, Repco, + Home First dropped), so the core model has 8 firms. Power is very low.
2. **Macro controls are perfectly collinear with quarter fixed effects** (they vary only over time). PanelOLS drops them (`drop_absorbed`); their coefficients cannot be estimated under TWFE. Tables state this.
3. **"5Y AAA corporate bond yield" was not supplied.** The file is the **5-year Indian government bond yield**; used as a proxy. Replace if you obtain AAA data.
4. **Model 2 (H2) is estimated on an annual panel** (`screener_annual_panel.csv`, 16 firms × FY21–FY26; interest / average borrowings) because quarterly interest expense before Mar-2024 does not exist in any supplied file. Result (Table 2b): β = 0.030 (CR1 SE 0.012, wild-bootstrap p = 0.015) → a 10pp higher pre-policy bank share is associated with ~0.3pp higher borrowing cost after the policy; same with FY24 dropped (p = 0.017). Supports H2. FY24 is a partly-treated year.
5. **Data quality** (see critical finding above): `Bank_Borrowing_Share_Pct` is constant over time for 15/16 firms (a snapshot, so "freezing" is harmless but it suggests the series is not truly quarterly). Shriram unsecured AUM has implausible spikes (e.g. −93%/+101% in FY25 Q3/Q4; list in `outputs/flagged_outlier_quarters.csv`) – please verify against filings. Results are unchanged when these are removed.
6. Bank exposure (update) replaces the PDF's unsecured-share exposure for the main models; the PDF's version is in Table 3.
7. Placebo: both "exact model, full sample" (update) and "pre-policy sample only" (PDF) are reported.
8. Bootstrap-t gives p-values, not SEs; tables show CR1 SEs in parentheses with stars from the bootstrap p-value in brackets.
