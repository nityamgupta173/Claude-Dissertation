# Dissertation: RBI November 2023 risk-weight measures, consumer credit and NBFC bank funding

## Deliverables
* **Dissertation.pdf** – final printable version (TOC filled in). Rebuild with `node 06_build_dissertation.js && python3 tools_pdf.py`.
* **Dissertation.docx** – editable Word version (Word asks to update fields on opening; click Yes to fill the TOC).
* **Dissertation_Data.xlsx** – every dataset used, with sources (README sheet explains each sheet), plus all regression results.

## Hypotheses and headline results
| | Test | Result |
|---|---|---|
| H1 | Bank credit to cards + personal loans vs exempt household credit (RBI monthly data) | −0.66 pp/month (≈ −7.9 pp a year), Newey–West p = 0.002, panel p = 0.008; placebo null. Moderate: p ≈ 0.06–0.11 with White SE / block bootstrap, 0.18 with classical OLS |
| H2 | Bank credit to NBFCs ex-HFC vs total non-food credit | −1.00 pp/month (≈ −11.9 pp a year), p = 0.049; placebo null. With month-of-year dummies p = 0.01–0.05 under White, Newey–West and block bootstrap |
| Firm level | 382 hand-read investor presentations, 16 firms | Disclosed CRAR hits (Bajaj −290 bps, SBI Cards ~−400 bps); cost-of-funds and bank-share effects in expected direction but not significant |

## How to reproduce
```
python 05_final_analysis.py        # all results, tables, figures -> outputs/final/, Dissertation_Data.xlsx
node 06_build_dissertation.js      # builds Dissertation.docx from outputs/final/results.json
```
`did_tools.py` – TWFE (linearmodels PanelOLS) with wild cluster bootstrap-t.

## Data
* `RBI_Sectoral_Bank_Credit.xlsx` – RBI DBIE Table 15 (Jan-2019 to Jul-2026).
* `hand_collected/` – cost of funds, bank share of borrowings, Bajaj segment AUM; every value has its source deck/page.
* `data_corrections/` – verified pre-policy bank shares and Shriram filings.
* Quarterly PDFs (382 files, Google Drive "Dissertation Quarterly") are not stored in git (`quarterly_pdfs/`, `quarterly_text/` ignored).

## Earlier work (kept for the record, not used in the dissertation)
`01_`–`04_` scripts, `outputs/` (other than `outputs/final/`), `master_panel_data_final.csv`, `screener_annual_panel.csv` and `extract/` document the exploratory phase, including the discovery that the original `nbfc_quarterly_data-v3.csv` contained assumed (not reported) values for several fields.
