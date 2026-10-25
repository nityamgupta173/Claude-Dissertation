# Dissertation: RBI November 2023 risk-weight measures, consumer credit and NBFC bank funding

## Deliverables
* **Dissertation.pdf** – final printable version. **Dissertation.docx** – editable Word version.
* **Dissertation_Data_Compilation.xlsx** – every dataset used, with sources. Raw inputs in blue, derived series as live Excel formulas in black (growth rates, gaps, test-window and post-policy flags, share returns, pre/post firm periods).

## Method (final version)
* Bank credit (RBI DBIE Table 15, merger-adjusted, Apr-2021 to Jan-2025): Welch t-test of the monthly growth gap before vs after the circular, and OLS on a post-policy dummy with calendar-month dummies; placebo date, pre-trend, rollback and drop-one checks.
* Share prices (Investing.com, `event_study/raw/`): market-model event study around 17-Nov-2023 and 27-Feb-2025 (`event_study/event_study.py`).
* Firm level (investor presentations, `hand_collected/`): change in cost of funds and bank share, non-HFC NBFCs vs HFCs.

## How to reproduce
```
python 05_final_analysis.py          # data preparation, descriptives, base figures -> outputs/final/
python 07_simple_tests.py            # t-tests and dummy regressions -> outputs/final/simple_results.json
python event_study/event_study.py    # event study -> event_study/results.json
python 08_figures.py                 # remaining figures -> outputs/final/f_*.png
node 06_build_dissertation.js && python3 tools_pdf.py && python3 tools_caption_pages.py && node 06_build_dissertation.js && python3 tools_pdf.py
python 09_build_data_workbook.py && python3 tools_xlsx_recalc.py Dissertation_Data_Compilation.xlsx
```
`tools_xlsx_recalc.py` needs LibreOffice Calc (`apt-get install libreoffice-calc`).

## Data
* `RBI_Sectoral_Bank_Credit.xlsx` – RBI DBIE Table 15 (Jan-2019 to Jul-2026).
* `hand_collected/` – cost of funds, bank share of borrowings, Bajaj segment AUM; every value has its source deck/page.
* `data_corrections/` – verified pre-policy bank shares and Shriram filings.
* Quarterly PDFs (382 files, Google Drive "Dissertation Quarterly") are not stored in git (`quarterly_pdfs/`, `quarterly_text/` ignored).

## Earlier work (kept for the record, not used in the dissertation)
`01_`–`04_` scripts, `outputs/` (other than `outputs/final/`), `master_panel_data_final.csv`, `screener_annual_panel.csv` and `extract/` document the exploratory phase, including the discovery that the original `nbfc_quarterly_data-v3.csv` contained assumed (not reported) values for several fields.
