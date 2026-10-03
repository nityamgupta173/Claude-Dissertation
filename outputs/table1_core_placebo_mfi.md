**Table 1. Core continuous DiD, placebo and MFI reversal tests (update specification)**

|  | (1) Core: unsecured | (2) Placebo Nov-22 (full) | (3) Placebo Nov-22 (pre-policy) | (4) MFI, Feb-25 reversal |
|---|---|---|---|---|
| Post-Nov2023 x Bank exposure | 0.021 |  |  |  |
|  | (0.059) |  |  |  |
|  | [0.737] |  |  |  |
| Post-Nov2022 x Bank exposure |  | -0.056 | -0.140** |  |
|  |  | (0.054) | (0.076) |  |
|  |  | [0.339] | [0.032] |  |
| Post-Feb2025 x Bank exposure |  |  |  | -0.421 |
|  |  |  |  | (0.310) |
|  |  |  |  | [0.421] |
| Firm FE | Yes | Yes | Yes | Yes |
| Quarter FE | Yes | Yes | Yes | Yes |
| Macro controls (repo, 5Y yield, GDP) | n/a (absorbed by time FE) | n/a (absorbed by time FE) | n/a (absorbed by time FE) | n/a (absorbed by time FE) |
| Firm controls | No | No | No | No |
| Observations | 184 | 184 | 104 | 69 |
| Firms (clusters) | 8 | 8 | 8 | 3 |
| Adj. R-squared (LSDV) | 0.002 | 0.003 | 0.129 | 0.172 |
| Within R-squared | 0.000 | 0.001 | 0.028 | 0.017 |

*Dependent variables are 100 x log difference of AUM (approx. QoQ growth in percentage points). Exposure is the Q2 FY24 (Sep-2023) bank-borrowing share in % points. Estimated with linearmodels PanelOLS (firm + quarter fixed effects). Parentheses: CR1 cluster-robust SE. Square brackets: wild cluster bootstrap-t p-value (Webb weights, 19999 draws, null imposed); stars from that p-value: * p<0.1, ** p<0.05, *** p<0.01. Models (1)-(3): 8 firms with non-zero unsecured AUM; model (4): 3 firms with MFI AUM only - bootstrap inference is not reliable with 3 clusters.*