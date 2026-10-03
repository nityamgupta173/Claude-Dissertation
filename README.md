# NBFC risk-weight impact analysis (FY21–FY26)

Run order: `python 01_build_master_panel.py` → `python 02_analysis.py` (needs pandas, linearmodels, scipy, matplotlib, openpyxl).
Everything lands in `master_panel_data_final.csv` and `outputs/`.

## What was done
| Step | Status |
|---|---|
| Screener.in P&L patch | **Not done – the 16 Screener Excel files are not in the repo.** The script has a best-effort (untested) patcher: put the files in `./screener/` and re-run. Unpatched NAs: Interest income 325, Opex 320, PAT 112 (Net worth: none). None of these feed the regressions. |
| Macro merge | Repo (RBI, day-weighted quarterly mean), 5Y yield (quarterly mean of monthly), GDP YoY (computed from FRED levels; matches the supplied YoY file exactly). |
| Frozen exposure | `Pre_Policy_Bank_Exposure` = Q2 FY24 `Bank_Borrowing_Share_Pct`, broadcast per firm. |
| Balanced panel | Home First dropped (only 22 of 24 quarters) → 15 firms × 24 quarters. |
| Model | `linearmodels.PanelOLS`, firm + quarter FE, DV = 100 × log difference of AUM. Inference = own implementation of **wild cluster bootstrap-t** (Webb 6-pt weights, null imposed, 19,999 draws), validated against PanelOLS point estimates. |
| Placebo / MFI / PDF Steps 2–7 | Done; see `outputs/table1…`, `table2…`, `table3…` (.md/.tex/.csv), `event_study.png`, `exposure_histogram.png`, `summary_statistics*.csv`, `diagnostics.txt`. |

## Results (bank-share exposure, 8 firms with unsecured AUM)
* **Core DiD**: β = 0.021 (CR1 SE 0.059, wild-bootstrap p = 0.74). **No detectable effect** of bank dependence on unsecured growth. Robust to dropping outlier quarters, winsorising, and leave-one-firm-out (p never < 0.14).
* **Placebo (Nov-2022)**: full-sample version β = −0.056, p = 0.34 (null, as expected). **The pre-policy-only version (as in the PDF) is NOT null**: β = −0.140, wild-bootstrap p = 0.032 (but t(7) p = 0.107). Treat as a warning about pre-trends/low power, not a clean pass.
* **Event study**: 12 pre-period coefficients jointly zero: F = 0.32, wild-bootstrap p = 0.37 → parallel trends not rejected (but CIs after the policy are very wide).
* **MFI / Feb-2025**: β = −0.42, p = 0.42 – **not the expected significant positive**, and uninterpretable: only **3 firms** (Arman, CreditAccess, IIFL) have MFI data.
* Secured growth (13 firms) β = 0.015, p = 0.65; total growth (15 firms) β = −0.038, p = 0.60. Alternative exposure (unsecured share) also null. Capital triple interaction (M3) is significant (β₂ = −0.038, p = 0.018) but with the *opposite* sign to H3 and only 8 clusters – fragile.
* I report what the data show; none of the "expected" outputs were forced.

## Things you must know (limitations)
1. **8 clusters, not 16.** Log growth is undefined for the 8 firms with zero unsecured AUM (AAVAS, Can Fin, CreditAccess, LIC HF, Muthoot, PNB HF, Repco, + Home First dropped), so the core model has 8 firms. Power is very low.
2. **Macro controls are perfectly collinear with quarter fixed effects** (they vary only over time). PanelOLS drops them (`drop_absorbed`); their coefficients cannot be estimated under TWFE. Tables state this.
3. **"5Y AAA corporate bond yield" was not supplied.** The file is the **5-year Indian government bond yield**; used as a proxy. Replace if you obtain AAA data.
4. **Model 2 (borrowing cost, H2) not estimable**: interest expense exists for only 2 firms with ≥ 8 quarters.
5. **Data quality**: `Bank_Borrowing_Share_Pct` is constant over time for 15/16 firms (a snapshot, so "freezing" is harmless but it suggests the series is not truly quarterly). Shriram unsecured AUM has implausible spikes (e.g. −93%/+101% in FY25 Q3/Q4; list in `outputs/flagged_outlier_quarters.csv`) – please verify against filings. Results are unchanged when these are removed.
6. Bank exposure (update) replaces the PDF's unsecured-share exposure for the main models; the PDF's version is in Table 3.
7. Placebo: both "exact model, full sample" (update) and "pre-policy sample only" (PDF) are reported.
8. Bootstrap-t gives p-values, not SEs; tables show CR1 SEs in parentheses with stars from the bootstrap p-value in brackets.
