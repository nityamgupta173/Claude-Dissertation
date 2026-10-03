**Table 2. Models 1-5 (bank-funding exposure, firm-level controls)**

|  | (M1) Unsecured | (M3) Unsecured x capital | (M4) Secured | (M4b) Total | (M5) MFI | (M5') Non-MFI consumer |
|---|---|---|---|---|---|---|
| Post-Nov2023 x Bank exposure | 0.011 | -0.185* | 0.015 | -0.038 | -0.665 | 0.215 |
|  | (0.059) | (0.061) | (0.032) | (0.055) | (0.389) | (0.246) |
|  | [0.833] | [0.052] | [0.651] | [0.597] | [0.311] | [0.511] |
| Post x Exposure x Tier-1 (lag, centred) |  | -0.038** |  |  |  |  |
|  |  | (0.015) |  |  |  |  |
|  |  | [0.018] |  |  |  |  |
| Post-Feb2025 x Bank exposure |  |  |  |  | -0.227 | -0.465 |
|  |  |  |  |  | (0.402) | (0.486) |
|  |  |  |  |  | [0.649] | [0.464] |
| Firm FE | Yes | Yes | Yes | Yes | Yes | Yes |
| Quarter FE | Yes | Yes | Yes | Yes | Yes | Yes |
| Macro controls (repo, 5Y yield, GDP) | Absorbed by Quarter FE | Absorbed by Quarter FE | Absorbed by Quarter FE | Absorbed by Quarter FE | Absorbed by Quarter FE | Absorbed by Quarter FE |
| Firm controls (lag RoA, lag GNPA, ln AUM) | Yes | Yes | Yes | Yes | Yes | Yes |
| Observations | 184 | 184 | 299 | 345 | 69 | 184 |
| Firms (clusters) | 8 | 8 | 13 | 15 | 3 | 8 |
| Adj. R-squared (LSDV) | -0.006 | -0.021 | -0.021 | 0.205 | 0.165 | 0.021 |
| Within R-squared | 0.012 | 0.024 | 0.002 | 0.030 | 0.101 | 0.045 |

*Dependent variables are 100 x log difference of AUM (approx. QoQ growth in percentage points). Exposure is the Q2 FY24 (Sep-2023) bank-borrowing share in % points. Estimated with linearmodels PanelOLS (firm + quarter fixed effects). Parentheses: CR1 cluster-robust SE. Square brackets: wild cluster bootstrap-t p-value (Webb weights, 19999 draws, null imposed); stars from that p-value: * p<0.1, ** p<0.05, *** p<0.01. Model 2 (borrowing cost) not estimable with available data - see README.*