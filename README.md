# Dissertation: RBI November 2023 risk-weight measures, consumer credit and NBFC bank funding

## Deliverables
* **Dissertation.pdf** – final printable version (TOC filled in). Rebuild with `node 06_build_dissertation.js && python3 tools_pdf.py`.
* **Dissertation.docx** – editable Word version (saved by tools_pdf.py with the table of contents and page numbers already filled in).
* **Dissertation_Data.xlsx** – every dataset used, with sources (README sheet explains each sheet), plus all regression results.

## Hypotheses and headline results (after the audit revision; see REVISION_LOG.md)
| | Test | Result |
|---|---|---|
| H2 | Bank credit to NBFCs ex-HFC vs total non-food credit | −1.07 pp/month (≈ −12.8 pp a year); marginal: p = 0.072 (NW 3 lags, t), 0.033–0.071 with seasonal controls; reverses after Apr-2025 rollback (n.s.) |
| H1 | Bank credit to cards + other personal loans vs exempt household credit | −0.75 pp/month (≈ −9.0 pp a year); p = 0.005, fixed-b 0.028, classical 0.128; shrinks to −0.34 if the Jul-21 pre-trend is projected |
| Firm level | 382 investor presentations, 16 firms (AI-assisted extraction, logged to source page) | Disclosed CRAR hits (Bajaj −290 bps, SBI Cards ~−400 bps); funding-cost, borrowing-mix and segment tests insignificant / confounded |

## How to reproduce
```
python 05_final_analysis.py        # all results, tables, figures -> outputs/final/, Dissertation_Data.xlsx
node 06_build_dissertation.js      # builds Dissertation.docx from outputs/final/results.json
python3 tools_pdf.py               # Dissertation.pdf (+ re-saves the .docx with TOC and page numbers filled)
```
`did_tools.py` – TWFE (linearmodels PanelOLS) with wild cluster bootstrap-t.

## Data
* `RBI_Sectoral_Bank_Credit.xlsx` – RBI DBIE Table 15 (Jan-2019 to Jul-2026).
* `hand_collected/` – cost of funds, bank share of borrowings, Bajaj segment AUM; every value has its source deck/page.
* `data_corrections/` – verified pre-policy bank shares and Shriram filings.
* Quarterly PDFs (382 files, Google Drive "Dissertation Quarterly") are not stored in git (`quarterly_pdfs/`, `quarterly_text/` ignored).

## Earlier work (kept for the record, not used in the dissertation)
`01_`–`04_` scripts, `outputs/` (other than `outputs/final/`), `master_panel_data_final.csv`, `screener_annual_panel.csv` and `extract/` document the exploratory phase, including the discovery that the original `nbfc_quarterly_data-v3.csv` contained assumed (not reported) values for several fields.
