"""Builds Dissertation_Tests.xlsx: every statistical test in the dissertation, done in Excel.
Each test has its own sheet. Where Excel has the functions (T.TEST, LINEST, SLOPE, INTERCEPT, STEYX, BINOM.DIST,
F.DIST.RT, CHISQ.DIST.RT, MINVERSE) the test is written as live formulas on the data. Newey-West standard errors
cannot be computed with standard Excel functions, so that sheet holds the output of the Python estimation.
Inputs are in blue, formulas in black; the column 'Dissertation' shows the value reported in the dissertation.
Run: python 10_build_tests_workbook.py && python3 tools_xlsx_recalc.py Dissertation_Tests.xlsx
"""
import glob
import json
import os
import warnings

import numpy as np
import pandas as pd
import statsmodels.api as sm
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter as L

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUTF = os.path.join(HERE, "Dissertation_Tests.xlsx")
S = json.load(open(os.path.join(HERE, "outputs", "final", "simple_results.json")))
E = json.load(open(os.path.join(HERE, "event_study", "results.json")))

# ------------------------------------------------------------------ data (same preparation as the analysis scripts)
src = open(os.path.join(HERE, "05_final_analysis.py")).read()
ns = {"__file__": os.path.join(HERE, "05_final_analysis.py")}
exec(src[:src.index("def its(")], ns)
g, lev, WIN, TREAT, CTRL = ns["g"], ns["lev"], ns["WIN"], ns["TREAT"], ns["CTRL"]

# ------------------------------------------------------------------ styles
FONT = "Arial"
F_IN = Font(name=FONT, size=10, color="0000FF")
F_FX = Font(name=FONT, size=10, color="000000")
F_LK = Font(name=FONT, size=10, color="008000")           # links to another sheet
F_HD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TX = Font(name=FONT, size=10)
F_B = Font(name=FONT, size=10, bold=True)
F_TI = Font(name=FONT, size=12, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="555555")
F_REF = Font(name=FONT, size=10, color="0000FF", italic=True)
FILL_HD = PatternFill("solid", fgColor="1F3864")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
P3, P4, PCT = "0.000", "0.0000", "0.00%"

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, note):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = F_TI
    ws["A2"] = note; ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False
    return ws


def put(ws, r, c, v, font=F_FX, fmt=None, fill=None, bold=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=10, bold=True, color=font.color) if bold else font
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    return cell


def hrow(ws, r, c0, labels, height=30):
    for j, t in enumerate(labels):
        cell = ws.cell(row=r, column=c0 + j, value=t)
        cell.font = F_HD; cell.fill = FILL_HD; cell.border = BOX
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[r].height = height


def section(ws, r, c, text, span=3):
    for j in range(span):
        ws.cell(row=r, column=c + j).fill = FILL_SUB
    put(ws, r, c, text, F_B, fill=FILL_SUB)


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# =================================================================== Data_Monthly
DM = g[WIN[0][:7]:].copy()
DM_COLS = [("Credit cards", "B"), ("Other personal loans", "C"), ("Housing", "D"), ("Vehicle loans", "E"), ("Education", "F"),
           ("NBFCs ex-HFC", "G"), ("Non-food credit", "H"), ("Non-food ex NBFC, cards, other PL", "I")]
ws = sheet("Data_Monthly", "Monthly growth of bank credit by sector (per cent, log change) and the growth gaps tested",
           "Source: RBI DBIE Table 15, merger-adjusted where the RBI publishes it (see Dissertation_Data_Compilation.xlsx, sheet RBI_Growth). "
           "Blank cells: July 2023 for series affected by the HDFC merger. Gaps are formulas.")
hrow(ws, 4, 1, ["Date (last reporting Friday)", "Credit cards", "Other personal loans", "Housing", "Vehicle loans", "Education",
                "NBFCs excl. HFCs", "Non-food credit", "Non-food excl. NBFCs, cards, other PL",
                "Gap1 (H1): avg(cards, other PL) minus avg(housing, vehicle, education)", "Gap2 (H2): NBFCs excl. HFCs minus non-food",
                "Gap2 alt.: NBFCs excl. HFCs minus non-food excl. NBFCs, cards, PL", "Calendar month"], 66)
DROW = {}
for i, (d, row) in enumerate(DM.iterrows()):
    r = 5 + i; DROW[d] = r
    put(ws, r, 1, d.to_pydatetime(), F_IN, "dd-mmm-yyyy")
    for k, c in DM_COLS:
        v = row[k]
        put(ws, r, ord(c) - 64, None if pd.isna(v) else float(v), F_IN, P3)
    put(ws, r, 10, f'=IF(COUNT(B{r}:C{r})=0,"",AVERAGE(B{r}:C{r})-AVERAGE(D{r}:F{r}))', F_FX, P3)
    put(ws, r, 11, f'=IF(COUNT(G{r},H{r})<2,"",G{r}-H{r})', F_FX, P3)
    put(ws, r, 12, f'=IF(COUNT(G{r},I{r})<2,"",G{r}-I{r})', F_FX, P3)
    put(ws, r, 13, f"=MONTH(A{r})", F_FX, "0")
widths(ws, {"A": 14, **{L(j): 11 for j in range(2, 10)}, "J": 20, "K": 18, "L": 20, "M": 9})
ws.freeze_panes = "B5"


# =================================================================== generic gap-test block
def gap_block(ws, r0, title, series, brk, ref, yfun, *, simple=True, months=True, diag=False, trend=False, short=None):
    """Write the data rows (columns A to R) and the tests (columns T onward) for one growth-gap series.
    series: pandas Series of the gap (dates -> values, NaN rows dropped); yfun(data_row) -> Excel formula of the gap."""
    series = series.dropna()
    rows = [(d, DROW[d]) for d in series.index]
    n = len(rows); npre = int((series.index < pd.Timestamp(brk)).sum())
    rc, rh, rd = r0 + 1, r0 + 2, r0 + 3
    rl = rd + n - 1
    put(ws, r0, 1, title, F_TI)
    hdr = ["Date", "Gap (pp)", "Post"] + [f"M{m}" for m in range(2, 13)]
    if months: hdr += ["Fitted", "Residual", "Residual squared"]
    if trend: hdr += ["Time (years from start)"]
    hrow(ws, rh, 1, hdr, 30)
    BRK = f"$U${r0 + 1}"
    for i, (d, dr) in enumerate(rows):
        r = rd + i
        put(ws, r, 1, f"=Data_Monthly!A{dr}", F_LK, "mmm-yyyy")
        put(ws, r, 2, yfun(dr), F_LK, P3)
        put(ws, r, 3, f"=IF($A{r}>={BRK},1,0)", F_FX, "0")
        for m in range(2, 13):
            put(ws, r, 2 + m, f"=IF(MONTH($A{r})={m},1,0)", F_FX, "0")
        if months:
            put(ws, r, 15, f"=$B${rc}+SUMPRODUCT(C{r}:N{r},$C${rc}:$N${rc})", F_FX, P3)
            put(ws, r, 16, f"=B{r}-O{r}", F_FX, P3)
            put(ws, r, 17, f"=P{r}^2", F_FX, P3)
        if trend:
            put(ws, r, 18, f"=(ROW()-ROW($A${rd}))/12", F_FX, P3)
    Y, X1, XA = f"B{rd}:B{rl}", f"C{rd}:C{rl}", f"C{rd}:N{rl}"
    YPRE, YPOST = f"B{rd}:B{rd + npre - 1}", f"B{rd + npre}:B{rl}"
    A = {}
    cur = [r0]

    def item(key, label, formula, fmt=P3, inp=False):
        r = cur[0]
        put(ws, r, 20, label, F_TX)
        put(ws, r, 21, formula, F_IN if inp else F_FX, fmt)
        if key in ref and ref[key] is not None:
            put(ws, r, 22, float(ref[key]), F_REF, fmt)
        A[key] = f"$U${r}"
        cur[0] += 1

    def sec(text):
        cur[0] += 1
        section(ws, cur[0], 20, text, 3)
        for cc, tx in [(21, "Excel"), (22, "Dissertation")]:
            put(ws, cur[0], cc, tx, F_B, fill=FILL_SUB).alignment = Alignment(horizontal="right")
        cur[0] += 1

    put(ws, r0, 20, "Test results", F_B)
    cur[0] = r0 + 1
    item("brk", "Post = 1 for months dated on or after", pd.Timestamp(brk).to_pydatetime(), "dd-mmm-yyyy", inp=True)
    put(ws, r0 + 2, 20, f"Months before: rows {rd} to {rd + npre - 1}; after: rows {rd + npre} to {rl}.", F_NOTE)
    cur[0] = r0 + 2
    # ---- Welch t-test
    sec("Welch two-sample t-test (unequal variances)")
    item("w.n_pre", "Months before (n1)", f"=COUNT({YPRE})", "0")
    item("w.n_post", "Months after (n2)", f"=COUNT({YPOST})", "0")
    item("w.mean_pre", "Mean gap before", f"=AVERAGE({YPRE})")
    item("w.mean_post", "Mean gap after", f"=AVERAGE({YPOST})")
    item("w.sd_pre", "Standard deviation before", f"=_xlfn.STDEV.S({YPRE})")
    item("w.sd_post", "Standard deviation after", f"=_xlfn.STDEV.S({YPOST})")
    item("w.diff", "Difference (after minus before)", f"={A['w.mean_post']}-{A['w.mean_pre']}")
    item("w.year", "Difference per year (x 12)", f"=12*{A['w.diff']}")
    item("w.se", "Standard error of the difference", f"=SQRT({A['w.sd_post']}^2/{A['w.n_post']}+{A['w.sd_pre']}^2/{A['w.n_pre']})")
    item("w.t", "t statistic", f"={A['w.diff']}/{A['w.se']}")
    v1, v2 = f"({A['w.sd_post']}^2/{A['w.n_post']})", f"({A['w.sd_pre']}^2/{A['w.n_pre']})"
    item("w.df", "Welch degrees of freedom", f"=({v1}+{v2})^2/({v1}^2/({A['w.n_post']}-1)+{v2}^2/({A['w.n_pre']}-1))", "0.0")
    item("w.p", "p-value, two-tailed (T.TEST type 3)", f"=_xlfn.T.TEST({YPOST},{YPRE},2,3)")
    if simple:
        sec("Regression 1 (Post dummy only)")
        L1 = f"LINEST({Y},{X1},TRUE,TRUE)"
        item("s.const", "Constant", f"=INDEX({L1},1,2)")
        item("s.b", "Coefficient on Post", f"=INDEX({L1},1,1)")
        item("s.se", "Standard error", f"=INDEX({L1},2,1)")
        item("s.t", "t statistic", f"={A['s.b']}/{A['s.se']}")
        item("s.dfres", "Residual degrees of freedom", f"=INDEX({L1},4,2)", "0")
        item("s.p", "p-value", f"=_xlfn.T.DIST.2T(ABS({A['s.t']}),{A['s.dfres']})")
        item("s.r2", "R squared", f"=INDEX({L1},3,1)")
        item("s.ssres", "Residual sum of squares", f"=INDEX({L1},5,2)")
    if months:
        L2 = f"LINEST({Y},{XA},TRUE,TRUE)"
        sec("Regression 2 (Post and 11 month dummies): model summary")
        item("m.n", "Observations", f"=COUNT({Y})", "0")
        item("m.r2", "R squared", f"=INDEX({L2},3,1)")
        item("m.R", "R", f"=SQRT({A['m.r2']})")
        item("m.dfres", "Residual degrees of freedom", f"=INDEX({L2},4,2)", "0")
        item("m.adj_r2", "Adjusted R squared", f"=1-(1-{A['m.r2']})*({A['m.n']}-1)/{A['m.dfres']}")
        item("m.se_est", "Standard error of the estimate", f"=INDEX({L2},3,2)")
        sec("Regression 2 (Post and 11 month dummies): ANOVA")
        item("m.ss_reg", "Regression sum of squares", f"=INDEX({L2},5,1)")
        item("m.ss_res", "Residual sum of squares", f"=INDEX({L2},5,2)")
        item("m.df_reg", "Regression degrees of freedom", f"=COLUMNS({XA})", "0")
        item("m.df_res", "Residual degrees of freedom", f"={A['m.dfres']}", "0")
        item("m.ms_reg", "Regression mean square", f"={A['m.ss_reg']}/{A['m.df_reg']}")
        item("m.ms_res", "Residual mean square", f"={A['m.ss_res']}/{A['m.df_res']}")
        item("m.F", "F statistic", f"=INDEX({L2},4,1)")
        item("m.F_p", "Significance of F", f"=_xlfn.F.DIST.RT({A['m.F']},{A['m.df_reg']},{A['m.df_res']})")
        sec("Regression 2 (Post and 11 month dummies): Post")
        gr = r0 + 1                                       # coefficient grid: columns X to AC, beside the results list
        item("m.b", "Coefficient on Post", f"=Y{gr + 2}")
        item("m.se", "Standard error", f"=Z{gr + 2}")
        item("m.t", "t statistic", f"=AA{gr + 2}")
        item("m.p", "p-value", f"=AB{gr + 2}")
        if simple:
            item("m.months_joint_p", "Month dummies jointly zero: F-test p-value",
                 f"=_xlfn.F.DIST.RT((({A['s.ssres']}-{A['m.ss_res']})/11)/({A['m.ss_res']}/{A['m.df_res']}),11,{A['m.df_res']})")
        else:
            cur[0] += 1
        if diag:
            P, Q = f"P{rd}:P{rl}", f"Q{rd}:Q{rl}"
            sec("Diagnostic tests on the residuals of Regression 2")
            item("m.dw", "Durbin-Watson", f"=SUMXMY2(P{rd + 1}:P{rl},P{rd}:P{rl - 1})/SUMSQ({P})")
            item("m.m2", "Second moment of residuals", f"=SUMSQ({P})/{A['m.n']}", P4)
            item("m.skew", "Skewness", f"=(SUMPRODUCT({P}^3)/{A['m.n']})/{A['m.m2']}^1.5")
            item("m.kurt", "Kurtosis", f"=(SUMPRODUCT({P}^4)/{A['m.n']})/{A['m.m2']}^2")
            item("m.jb", "Jarque-Bera statistic", f"={A['m.n']}/6*({A['m.skew']}^2+({A['m.kurt']}-3)^2/4)")
            item("m.jb_p", "Jarque-Bera p-value (chi-squared, 2 df)", f"=_xlfn.CHISQ.DIST.RT({A['m.jb']},2)")
            item("m.bp_r2", "Breusch-Pagan: R squared, squared residuals on regressors", f"=INDEX(LINEST({Q},{XA},TRUE,TRUE),3,1)")
            item("m.bp", "Breusch-Pagan statistic (n x R squared)", f"={A['m.n']}*{A['m.bp_r2']}")
            item("m.bp_p", "Breusch-Pagan p-value (chi-squared, 12 df)", f"=_xlfn.CHISQ.DIST.RT({A['m.bp']},{A['m.df_reg']})")
            item("m.max_vif", "Largest variance inflation factor", f"=MAX(AC{gr + 2}:AC{gr + 13})")
    if trend:
        L3 = f"LINEST({Y},R{rd}:R{rl},TRUE,TRUE)"
        sec("Pre-policy trend: gap on time in years (LINEST)")
        item("t.slope_per_year", "Slope (pp per year)", f"=INDEX({L3},1,1)")
        item("t.se", "Standard error", f"=INDEX({L3},2,1)")
        item("t.tstat", "t statistic", f"={A['t.slope_per_year']}/{A['t.se']}")
        item("t.p", "p-value", f"=_xlfn.T.DIST.2T(ABS({A['t.tstat']}),INDEX({L3},4,2))")
        item("t.n", "Observations", f"=COUNT({Y})", "0")
    if short:
        sec("Economic size over the post-policy months")
        item("sf.outstanding", short[0], float(short[1]), "#,##0", inp=True)
        item("sf.months", "Post-policy months", f"={A['w.n_post']}", "0")
        item("sf.pct", "Cumulative shortfall in growth (per cent)", f"=100*(EXP(-{A['w.diff']}*{A['sf.months']}/100)-1)", "0.00")
        item("sf.amount", "Implied shortfall (₹ crore)", f"={A['sf.outstanding']}*(EXP(-{A['w.diff']}*{A['sf.months']}/100)-1)", "#,##0")
        item("sf.lakh", "Implied shortfall (₹ lakh crore)", f"={A['sf.amount']}/100000", "0.00")
    if months:
        section(ws, gr, 24, "Regression 2: all coefficients", 6)
        hrow(ws, gr + 1, 24, ["Variable", "B", "Std. error", "t", "p-value", "VIF"], 18)
        names = ["Post"] + [f"M{m}" for m in range(2, 13)] + ["Constant"]
        for k, nm in enumerate(names):
            r = gr + 2 + k
            col = 13 if nm == "Constant" else 12 - k   # LINEST lists coefficients in reverse order of the X columns
            put(ws, r, 24, nm, F_TX)
            put(ws, r, 25, f"=INDEX({L2},1,{col})", F_FX, P3)
            put(ws, r, 26, f"=INDEX({L2},2,{col})", F_FX, P3)
            put(ws, r, 27, f"=Y{r}/Z{r}", F_FX, P3)
            put(ws, r, 28, f"=_xlfn.T.DIST.2T(ABS(AA{r}),{A['m.dfres']})", F_FX, P3)
            if nm != "Constant":
                put(ws, r, 29, f"=INDEX(MINVERSE($Y${gr + 18}:$AJ${gr + 29}),{k + 1},{k + 1})", F_FX, P3)
        put(ws, rc, 1, "Coefficients", F_NOTE)
        put(ws, rc, 2, f"=Y{gr + 14}", F_FX, P3)
        for k in range(12):
            put(ws, rc, 3 + k, f"=Y{gr + 2 + k}", F_FX, P3)
        put(ws, gr + 15, 24, "Row " + str(rc) + " repeats these coefficients above the data columns; Fitted = Constant + sum of coefficient x regressor.", F_NOTE)
        put(ws, gr + 16, 24, "VIF = diagonal of the inverse of the correlation matrix of the regressors:", F_NOTE)
        for i_ in range(12):
            put(ws, gr + 17, 25 + i_, names[i_], F_B).alignment = Alignment(horizontal="right")
            put(ws, gr + 18 + i_, 24, names[i_], F_B)
            for j_ in range(12):
                ci, cj = L(3 + i_), L(3 + j_)
                put(ws, gr + 18 + j_, 25 + i_, f"=CORREL(${ci}${rd}:${ci}${rl},${cj}${rd}:${cj}${rl})", F_FX, "0.00")
        cur[0] = max(cur[0], gr + 30)
    widths(ws, {"A": 10, "B": 9, "C": 6, **{L(j): 5 for j in range(4, 15)}, "O": 8, "P": 8, "Q": 9, "R": 9, "S": 2,
                "T": 50, "U": 13, "V": 13, "W": 2, "X": 10, **{L(j): 8 for j in range(25, 37)}})
    return dict(A=A, rd=rd, rl=rl, end=max(rl, cur[0]))


def ref_from(d, prefix):
    return {f"{prefix}.{k}": v for k, v in d.items() if not isinstance(v, (list, dict))}


W = g[WIN[0]:WIN[1]]
GAPS = {"H1": W[TREAT].mean(axis=1) - W[CTRL].mean(axis=1), "H2": W["NBFCs ex-HFC"] - W["Non-food credit"],
        "H2_core": W["NBFCs ex-HFC"] - W["Non-food ex NBFC, cards, other PL"]}
end = lev.loc["2025-01"]
END_DATE = end.index[-1].strftime("%d-%b-%Y")

# ---- H1, H2, H2 alternative
specs = [("H1_Tests", "H1", "J", "Tables 7.1 to 7.5 (H1)", "Gap1: bank credit to credit cards and other personal loans minus housing, vehicle and education loans",
          True, ("Credit to cards and other PL, " + END_DATE + " (₹ crore)", end[TREAT].iloc[-1].sum())),
         ("H2_Tests", "H2", "K", "Tables 7.1 to 7.5 (H2)", "Gap2: bank credit to NBFCs (excl. HFCs) minus total non-food credit",
          True, ("Credit to NBFCs excl. HFCs, " + END_DATE + " (₹ crore)", end["NBFCs ex-HFC"].iloc[-1])),
         ("H2_Alt_Tests", "H2_core", "L", "Table 7.1, third row; Section 7.2", "Gap2 alternative: NBFCs (excl. HFCs) minus non-food credit excl. NBFCs, cards and other PL",
          False, None)]
INFO = []
for name, key, col, tbl, desc, diag, short in specs:
    ws = sheet(name, f"{desc}: tests, April 2021 to January 2025",
               f"Reproduces {tbl}. Data rows link to Data_Monthly; every result is an Excel formula. Months before the circular: April 2021 to 17 November 2023; after: 29 December 2023 to January 2025.")
    ref = {**ref_from(S[key]["welch"], "w"), **ref_from(S[key]["reg_simple"], "s"), **ref_from(S[key]["reg_months"], "m")}
    ref["w.year"] = 12 * S[key]["welch"]["diff"]
    ref["m.se_est"] = np.sqrt(S[key]["reg_months"]["ss_res"] / S[key]["reg_months"]["df_res"])
    if "shortfall" in S[key]:
        ref.update(ref_from(S[key]["shortfall"], "sf")); ref["sf.lakh"] = S[key]["shortfall"]["amount"] / 1e5
    gap_block(ws, 4, desc.split(":")[0], GAPS[key], "2023-11-30", ref, lambda dr, c=col: f"=Data_Monthly!{c}{dr}",
              diag=diag, short=short)
    ws.freeze_panes = "B7"
    INFO.append((name, tbl, "Welch t-test; regression on a Post dummy, without and with calendar-month dummies (model summary, ANOVA, coefficients)" +
                 ("; Durbin-Watson, Jarque-Bera, Breusch-Pagan and VIF; economic size" if diag else ""), "Excel formulas"))

# ---- placebo and pre-trend
PB = g["2021-07":"2023-11-17"]
for name, key, col, gap in [("H1_Placebo", "H1", "J", PB[TREAT].mean(axis=1) - PB[CTRL].mean(axis=1)),
                            ("H2_Placebo", "H2", "K", PB["NBFCs ex-HFC"] - PB["Non-food credit"])]:
    ws = sheet(name, f"{key}: placebo test with a fictitious policy date (November 2022) and pre-policy trend",
               "Reproduces Table 7.6. Pre-policy data only: July 2021 to 17 November 2023. 'Post' switches on at the fictitious date. "
               "The trend regression uses the same months, with time measured in years from July 2021.")
    ref = {**ref_from(S[key]["placebo_welch"], "w"), **ref_from(S[key]["placebo_reg_months"], "m"), **ref_from(S[key]["pretrend"], "t")}
    ref["w.year"] = 12 * S[key]["placebo_welch"]["diff"]
    ref["m.se_est"] = np.sqrt(S[key]["placebo_reg_months"]["ss_res"] / S[key]["placebo_reg_months"]["df_res"])
    gap_block(ws, 4, f"{key} gap, July 2021 to November 2023", gap, "2022-11-30", ref, lambda dr, c=col: f"=Data_Monthly!{c}{dr}",
              simple=False, trend=True)
    ws.freeze_panes = "B7"
    INFO.append((name, "Table 7.6", "Placebo Welch t-test and month-dummy regression at a fictitious date; linear pre-policy trend", "Excel formulas"))

# ---- rollback
RB = g["2023-12":"2026-07"]
ws = sheet("H2_Rollback", "H2 after the April 2025 rollback: December 2023 to March 2025 versus April 2025 to July 2026",
           "Reproduces the rollback row of Table 7.6. Welch t-test of the H2 gap; 'Post' marks the months after the rollback took effect (1 April 2025).")
ref = ref_from(S["H2"]["rollback_welch"], "w"); ref["w.year"] = 12 * S["H2"]["rollback_welch"]["diff"]
gap_block(ws, 4, "H2 gap, December 2023 to July 2026", RB["NBFCs ex-HFC"] - RB["Non-food credit"], "2025-04-30", ref,
          lambda dr: f"=Data_Monthly!K{dr}", simple=False, months=False)
ws.freeze_panes = "B7"
INFO.append(("H2_Rollback", "Table 7.6", "Welch t-test of the H2 gap before and after the rollback", "Excel formulas"))

# ---- H1 drop-one checks
ws = sheet("H1_DropOne", "H1 with one series dropped",
           "Reproduces Table 7.7. Each block recomputes the H1 gap without one series and repeats the Welch t-test and the month-dummy regression.")
DROP = [("Credit cards only (drop other personal loans)", ["Credit cards"], CTRL, "=Data_Monthly!B{r}-AVERAGE(Data_Monthly!D{r}:F{r})"),
        ("Other personal loans only (drop credit cards)", ["Other personal loans"], CTRL, "=Data_Monthly!C{r}-AVERAGE(Data_Monthly!D{r}:F{r})"),
        ("Control without housing", TREAT, ["Vehicle loans", "Education"], "=AVERAGE(Data_Monthly!B{r}:C{r})-AVERAGE(Data_Monthly!E{r}:F{r})"),
        ("Control without vehicle loans", TREAT, ["Housing", "Education"], "=AVERAGE(Data_Monthly!B{r}:C{r})-AVERAGE(Data_Monthly!D{r},Data_Monthly!F{r})"),
        ("Control without education", TREAT, ["Housing", "Vehicle loans"], "=AVERAGE(Data_Monthly!B{r}:C{r})-AVERAGE(Data_Monthly!D{r}:E{r})")]
hrow(ws, 4, 1, ["Specification", "Difference (t-test)", "p-value", "Coefficient on Post (month dummies)", "p-value"], 42)
r0 = 12
for i, (lbl, tr, ct, fx) in enumerate(DROP):
    rob = S["H1"]["robustness"][i]
    ref = {"w.diff": rob["diff"], "w.p": rob["p_welch"], "m.b": rob["b_months"], "m.p": rob["p_months"]}
    res = gap_block(ws, r0, lbl, W[tr].mean(axis=1) - W[ct].mean(axis=1), "2023-11-30", ref, lambda dr, f=fx: f.format(r=dr), simple=False)
    put(ws, 5 + i, 1, lbl, F_TX)
    for j, k in enumerate(["w.diff", "w.p", "m.b", "m.p"]):
        put(ws, 5 + i, 2 + j, "=" + res["A"][k].replace("$", ""), F_FX, P3)
    r0 = res["end"] + 3
widths(ws, {"A": 44, "B": 12, "C": 9, "D": 14, "E": 9})
INFO.append(("H1_DropOne", "Table 7.7", "Welch t-test and month-dummy regression with one H1 series dropped (summary at the top)", "Excel formulas"))

# =================================================================== Newey-West (output of the Python estimation)
ws = sheet("NeweyWest_Output", "Newey-West (HAC) standard errors: output of the Python estimation",
           "Excel has no built-in function for heteroskedasticity- and autocorrelation-consistent standard errors, so this sheet reports the output "
           "of statsmodels OLS (cov_type='HAC', Bartlett kernel, 3 lags, t-distribution with the residual degrees of freedom). Coefficients are identical to the LINEST results.")
r = 4


def nw_fit(gap, brk, months):
    gap = gap.dropna()
    X = pd.DataFrame({"Post": (gap.index >= brk).astype(float)}, index=gap.index)
    if months:
        md = pd.get_dummies(gap.index.month, prefix="M", drop_first=True).astype(float); md.index = gap.index
        X = pd.concat([X, md], axis=1)
    X = sm.add_constant(X)
    return sm.OLS(gap, X).fit(), sm.OLS(gap, X).fit(cov_type="HAC", cov_kwds={"maxlags": 3}, use_t=True)


models = [("H1, Post dummy only", GAPS["H1"], "2023-11-30", False), ("H1, Post and month dummies (Table 7.4, model 1)", GAPS["H1"], "2023-11-30", True),
          ("H2, Post dummy only", GAPS["H2"], "2023-11-30", False), ("H2, Post and month dummies (Table 7.4, model 2)", GAPS["H2"], "2023-11-30", True),
          ("H2 alternative, Post and month dummies", GAPS["H2_core"], "2023-11-30", True),
          ("H1 placebo (Nov-2022), Post and month dummies", PB[TREAT].mean(axis=1) - PB[CTRL].mean(axis=1), "2022-11-30", True),
          ("H2 placebo (Nov-2022), Post and month dummies", PB["NBFCs ex-HFC"] - PB["Non-food credit"], "2022-11-30", True)]
section(ws, r, 1, "Summary: coefficient on Post", 8); r += 1
hrow(ws, r, 1, ["Model", "n", "B (Post)", "Classical std. error", "Classical p-value", "Newey-West std. error", "Newey-West t", "Newey-West p-value"], 30); r += 1
for lbl, gp, brk, mo in models:
    f, nw = nw_fit(gp, brk, mo)
    vals = [int(f.nobs), f.params["Post"], f.bse["Post"], f.pvalues["Post"], nw.bse["Post"], nw.tvalues["Post"], nw.pvalues["Post"]]
    put(ws, r, 1, lbl, F_TX)
    for j, v in enumerate(vals):
        put(ws, r, 2 + j, float(v), F_IN, "0" if j == 0 else P3)
    r += 1
for lbl, gp, brk, mo in [models[1], models[3]]:
    r += 1
    f, nw = nw_fit(gp, brk, mo)
    section(ws, r, 1, f"Full output: {lbl}", 8); r += 1
    hrow(ws, r, 1, ["Variable", "B", "Classical std. error", "Newey-West std. error", "Newey-West t", "Newey-West p-value", "Residual df", "Lags"], 30); r += 1
    for v in f.params.index:
        put(ws, r, 1, "Constant" if v == "const" else v.replace("M_", "M"), F_TX)
        for j, x in enumerate([f.params[v], f.bse[v], nw.bse[v], nw.tvalues[v], nw.pvalues[v]]):
            put(ws, r, 2 + j, float(x), F_IN, P3)
        put(ws, r, 7, int(f.df_resid), F_IN, "0"); put(ws, r, 8, 3, F_IN, "0")
        r += 1
widths(ws, {"A": 48, **{L(j): 14 for j in range(2, 9)}})
INFO.append(("NeweyWest_Output", "Table 7.4 (Newey-West columns)", "Newey-West standard errors, 3 lags", "Python output (not possible with Excel functions)"))

# =================================================================== event study data
P = {}
for f in sorted(glob.glob(os.path.join(HERE, "event_study", "raw", "*.csv"))):
    d = pd.read_csv(f, encoding="utf-8-sig")
    d["Date"] = pd.to_datetime(d.Date, format="%d-%m-%Y")
    d["Price"] = pd.to_numeric(d.Price.astype(str).str.replace(",", ""))
    P[os.path.basename(f)[:-4]] = d.set_index("Date").Price.sort_index()
P = pd.DataFrame(P)
HFC = ["AAVAS", "CANFINHOME", "HOMEFIRST", "LICHSGFIN", "PNBHOUSING", "REPCOHOME"]
NONHFC = sorted(c for c in P.columns if c not in HFC + ["NIFTY50", "NIFTY500"])
ORDER = ["NIFTY50", "NIFTY500"] + HFC + NONHFC
COL = {s: L(2 + i) for i, s in enumerate(ORDER)}                         # Data_Prices / Data_Returns column of each series
NAMES = {"AAVAS": "AAVAS Financiers", "ARMANFIN": "Arman Financial", "BAJFINANCE": "Bajaj Finance", "CANFINHOME": "Can Fin Homes",
         "CHOLAFIN": "Cholamandalam Inv. & Fin.", "CREDITACC": "CreditAccess Grameen", "HOMEFIRST": "Home First Finance",
         "IIFL": "IIFL Finance", "LICHSGFIN": "LIC Housing Finance", "MMFIN": "M&M Financial Services", "MUTHOOTFIN": "Muthoot Finance",
         "PNBHOUSING": "PNB Housing Finance", "POONAWALLA": "Poonawalla Fincorp", "REPCOHOME": "Repco Home Finance",
         "SBICARD": "SBI Cards & Payment", "SHRIRAMFIN": "Shriram Finance", "NIFTY50": "Nifty 50", "NIFTY500": "Nifty 500"}
CONSUMER, MFI = ["BAJFINANCE", "SBICARD", "POONAWALLA"], ["CREDITACC", "ARMANFIN"]
BANKSHARE = {"AAVAS": 49.6, "ARMANFIN": 32.3, "BAJFINANCE": 23.0, "CANFINHOME": 57.0, "CHOLAFIN": 60.0, "CREDITACC": 52.9,
             "HOMEFIRST": 54.0, "IIFL": 55.3, "LICHSGFIN": 33.0, "MMFIN": 57.2, "MUTHOOTFIN": 65.4, "PNBHOUSING": 40.2,
             "REPCOHOME": 75.5, "SBICARD": 77.0, "SHRIRAMFIN": 25.6}

ws = sheet("Data_Prices", "Daily closing prices, 1 September 2022 to 28 March 2025",
           "Source: Investing.com historical data (NSE closing prices; index points for the Nifty indices).")
hrow(ws, 4, 1, ["Date"] + [f"{NAMES[s]} ({s})" for s in ORDER], 42)
for i, (d, row) in enumerate(P[ORDER].iterrows()):
    put(ws, 5 + i, 1, d.to_pydatetime(), F_IN, "dd-mmm-yyyy")
    for s in ORDER:
        put(ws, 5 + i, ord(COL[s]) - 64, float(row[s]), F_IN, "#,##0.00")
widths(ws, {"A": 12, **{COL[s]: 12 for s in ORDER}})
ws.freeze_panes = "B5"
NP = len(P); LAST = 4 + NP

PORTS = [("All 16 firms", "T", HFC + NONHFC), ("Non-HFC NBFCs (10)", "U", NONHFC), ("HFCs (6, exempt)", "V", HFC),
         ("Consumer lenders (Bajaj, SBI Cards, Poonawalla)", "W", CONSUMER),
         ("Non-HFC NBFCs excl. MFIs (8)", "X", [f for f in NONHFC if f not in MFI]),
         ("Non-HFC NBFCs excl. Bajaj (9)", "Y", [f for f in NONHFC if f != "BAJFINANCE"])]
DIFFS = [("Non-HFC NBFCs minus HFCs", "Z", "U", "V"), ("Consumer lenders minus HFCs", "AA", "W", "V")]
PCOL = {p: c for p, c, _ in PORTS} | {p: c for p, c, _, _ in DIFFS}


def avg_expr(members, r):
    cols = sorted(ord(COL[m]) - 64 for m in members)
    if cols == list(range(cols[0], cols[-1] + 1)):
        return f"=AVERAGE({L(cols[0])}{r}:{L(cols[-1])}{r})"
    return "=AVERAGE(" + ",".join(f"{L(c)}{r}" for c in cols) + ")"


ws = sheet("Data_Returns", "Daily simple returns and equal-weighted portfolio returns",
           "Return = today's close / previous close - 1 (formulas on Data_Prices). Portfolio returns are the simple average of member returns; "
           "difference portfolios are the return of one portfolio minus the other.")
hrow(ws, 4, 1, ["Date"] + [s for s in ORDER] + [p for p, _, _ in PORTS] + [p for p, _, _, _ in DIFFS], 54)
for i in range(NP):
    r = 5 + i
    put(ws, r, 1, f"=Data_Prices!A{r}", F_LK, "dd-mmm-yyyy")
    if i == 0:
        continue
    for s in ORDER:
        c = COL[s]; put(ws, r, ord(c) - 64, f"=Data_Prices!{c}{r}/Data_Prices!{c}{r - 1}-1", F_FX, PCT)
    for p, c, mem in PORTS:
        put(ws, r, column_index_from_string(c), avg_expr(mem, r), F_FX, PCT)
    for p, c, a, b in DIFFS:
        put(ws, r, column_index_from_string(c), f"={a}{r}-{b}{r}", F_FX, PCT)
widths(ws, {"A": 12, **{L(j): 10 for j in range(2, 28)}})
ws.freeze_panes = "B5"

RET = P.pct_change().iloc[1:]
EV = {"E1": ("2023-11-17", "the circular (released after market close on 16 November 2023)", -1),
      "E2": ("2025-02-27", "the rollback (released after market close on 25 February 2025; 26 February was a market holiday)", +1)}
ROW0 = {ek: 6 + RET.index.get_loc(pd.Timestamp(d0)) for ek, (d0, _, _) in EV.items()}   # Data_Returns row of day 0
EST = (-250, -31)
WINS = [("[0]", 0, 0), ("[0,+1]", 0, 1), ("[0,+2]", 0, 2), ("[-1,+1]", -1, 1)]


def est_rng(c, ek):
    r0 = ROW0[ek]; return f"Data_Returns!${c}${r0 + EST[0]}:${c}${r0 + EST[1]}"


# ---- market model per firm (Table 8.1)
ws = sheet("ES_MarketModel", "Market model for each stock: R(firm) = alpha + beta x R(Nifty 50), estimated on trading days -250 to -31",
           "Reproduces Table 8.1. INTERCEPT, SLOPE, RSQ and STEYX on the 220-day estimation window of each event (Data_Returns).")
hrow(ws, 4, 1, ["Company", "Symbol", "Group"] + [f"{ek}: {h}" for ek in EV for h in ["alpha (% per day)", "beta", "R squared", "Residual std. dev. (%)", "Days"]], 42)
r = 5
put(ws, r, 1, "Estimation window", F_B)
for k, ek in enumerate(EV):
    put(ws, r, 4 + 5 * k, f"=Data_Returns!A{ROW0[ek] + EST[0]}", F_LK, "dd-mmm-yyyy")
    put(ws, r, 5 + 5 * k, "to", F_TX)
    put(ws, r, 6 + 5 * k, f"=Data_Returns!A{ROW0[ek] + EST[1]}", F_LK, "dd-mmm-yyyy")
MM = {}
FIRMS_ORD = NONHFC + HFC
for i, s in enumerate(FIRMS_ORD):
    r = 6 + i
    put(ws, r, 1, NAMES[s], F_TX); put(ws, r, 2, s, F_TX); put(ws, r, 3, "HFC" if s in HFC else "Non-HFC NBFC", F_TX)
    for k, ek in enumerate(EV):
        y, m = est_rng(COL[s], ek), est_rng("B", ek)
        c0 = 4 + 5 * k
        put(ws, r, c0, f"=INTERCEPT({y},{m})", F_FX, "0.000%")
        put(ws, r, c0 + 1, f"=SLOPE({y},{m})", F_FX, P3)
        put(ws, r, c0 + 2, f"=RSQ({y},{m})", F_FX, P3)
        put(ws, r, c0 + 3, f"=STEYX({y},{m})", F_FX, "0.000%")
        put(ws, r, c0 + 4, f"=COUNT({y})", F_FX, "0")
        MM[(ek, s)] = dict(a=f"ES_MarketModel!${L(c0)}${r}", b=f"ES_MarketModel!${L(c0 + 1)}${r}",
                           sd=f"ES_MarketModel!${L(c0 + 3)}${r}", n=f"ES_MarketModel!${L(c0 + 4)}${r}")
widths(ws, {"A": 28, "B": 13, "C": 13, **{L(j): 11 for j in range(4, 14)}})
ws.freeze_panes = "D6"
INFO.append(("ES_MarketModel", "Table 8.1", "Market model (alpha, beta, R squared, residual standard deviation) for every stock and both events", "Excel formulas"))

ES_HDR = ["Portfolio / company", "Group", "alpha", "beta", "Residual std. dev.", "Days", "AR day -1", "AR day 0", "AR day +1", "AR day +2"] + \
         [f"CAR {w}" for w, _, _ in WINS] + [f"t {w}" for w, _, _ in WINS] + [f"p {w}" for w, _, _ in WINS]


def es_rows(ws, r, ek, mcol, entities, use_mm=False):
    """entities: (label, group, return column, symbol or None). Returns {label: {window: (car cell, p cell)}}."""
    r0 = ROW0[ek]; out = {}
    for lbl, grp, c, sym in entities:
        put(ws, r, 1, lbl, F_TX); put(ws, r, 2, grp, F_TX)
        if use_mm:
            mm = MM[(ek, sym)]
            put(ws, r, 3, "=" + mm["a"], F_LK, "0.000%"); put(ws, r, 4, "=" + mm["b"], F_LK, P3)
            put(ws, r, 5, "=" + mm["sd"], F_LK, "0.000%"); put(ws, r, 6, "=" + mm["n"], F_LK, "0")
        else:
            y, m = est_rng(c, ek), est_rng(mcol, ek)
            put(ws, r, 3, f"=INTERCEPT({y},{m})", F_FX, "0.000%"); put(ws, r, 4, f"=SLOPE({y},{m})", F_FX, P3)
            put(ws, r, 5, f"=STEYX({y},{m})", F_FX, "0.000%"); put(ws, r, 6, f"=COUNT({y})", F_FX, "0")
        for j, dday in enumerate([-1, 0, 1, 2]):
            rr = r0 + dday
            put(ws, r, 7 + j, f"=Data_Returns!{c}{rr}-($C{r}+$D{r}*Data_Returns!{mcol}{rr})", F_FX, PCT)
        ar = {-1: f"G{r}", 0: f"H{r}", 1: f"I{r}", 2: f"J{r}"}
        out[lbl] = {}
        for j, (w, a, b) in enumerate(WINS):
            Lw = b - a + 1
            put(ws, r, 11 + j, "=" + "+".join(ar[d] for d in range(a, b + 1)), F_FX, PCT)
            put(ws, r, 15 + j, f"={L(11 + j)}{r}/($E{r}*SQRT({Lw}))", F_FX, "0.00")
            put(ws, r, 19 + j, f"=_xlfn.T.DIST.2T(ABS({L(15 + j)}{r}),$F{r}-2)", F_FX, P3)
            out[lbl][w] = (f"{L(11 + j)}{r}", f"{L(19 + j)}{r}")
        r += 1
    return out, r


ES_CELLS = {}
for ek, (d0, lbl, sign) in EV.items():
    ws = sheet(f"ES_{ek}", f"Event study around {lbl}: day 0 = {pd.Timestamp(d0).strftime('%d %B %Y').lstrip('0')}",
               ("Reproduces Table 8.2 and Table 8.4 (circular)." if ek == "E1" else "Reproduces Table 8.3 and Table 8.4 (rollback).") +
               " Benchmark: Nifty 50. AR = actual return minus (alpha + beta x market return). t = CAR / (residual std. dev. x square root of days); "
               "p from the t-distribution with (estimation days - 2) degrees of freedom.")
    put(ws, 4, 1, "Event days", F_B)
    for j, dday in enumerate([-1, 0, 1, 2]):
        put(ws, 4, 7 + j, f"=Data_Returns!A{ROW0[ek] + dday}", F_LK, "dd-mmm-yy")
    put(ws, 5, 1, "Nifty 50 return on day 0", F_TX); put(ws, 5, 8, f"=Data_Returns!B{ROW0[ek]}", F_LK, PCT)
    section(ws, 7, 1, "A. Equal-weighted portfolios", 22)
    hrow(ws, 8, 1, ES_HDR, 30)
    ents = [(p, "Portfolio", c, None) for p, c, _ in PORTS] + [(p, "Difference", c, None) for p, c, _, _ in DIFFS]
    pc, r = es_rows(ws, 9, ek, "B", ents)
    r += 1
    section(ws, r, 1, "B. Individual companies (Table 8.4 shows the day-0 AR)", 22); r += 1
    hrow(ws, r, 1, ES_HDR, 30); r += 1
    fstart = r
    fc, r = es_rows(ws, r, ek, "B", [(NAMES[s], "HFC" if s in HFC else "Non-HFC NBFC", COL[s], s) for s in FIRMS_ORD], use_mm=True)
    nb0, nb1 = fstart, fstart + len(NONHFC) - 1
    r += 1
    # sign test
    section(ws, r, 1, f"C. Sign test: how many of the 10 NBFCs had a {'negative' if sign < 0 else 'positive'} CAR", 8); r += 1
    hrow(ws, r, 1, ["Window", "NBFCs", f"Number {'negative' if sign < 0 else 'positive'}", "p-value (two-sided binomial, 0.5)", "Dissertation p"], 54); r += 1
    signrow = r
    for j, (w, _, _) in enumerate(WINS):
        cc = L(11 + j)
        put(ws, r, 1, w, F_TX)
        put(ws, r, 2, f"=COUNT({cc}{nb0}:{cc}{nb1})", F_FX, "0")
        put(ws, r, 3, f'=COUNTIF({cc}{nb0}:{cc}{nb1},"{"<0" if sign < 0 else ">0"}")', F_FX, "0")
        put(ws, r, 4, f"=MIN(1,2*MIN(_xlfn.BINOM.DIST(C{r},B{r},0.5,TRUE),1-_xlfn.BINOM.DIST(C{r}-1,B{r},0.5,TRUE)))", F_FX, P3)
        put(ws, r, 5, float(E[ek]["NIFTY50"]["cross_section"][w]["sign_nonhfc"]["p"]), F_REF, P3)
        r += 1
    r += 1
    # bank-share regression within NBFCs
    section(ws, r, 1, "D. Within the NBFCs: day-0 CAR regressed on the bank share of borrowings (30 Sep 2023)", 8); r += 1
    hrow(ws, r, 1, ["Company", "Bank share of borrowings (%)", "CAR [0]"], 30); r += 1
    b0 = r
    for s in NONHFC:
        if s not in BANKSHARE:
            continue
        put(ws, r, 1, NAMES[s], F_TX); put(ws, r, 2, BANKSHARE[s], F_IN, "0.0")
        put(ws, r, 3, f"=K{fstart + NONHFC.index(s)}", F_FX, PCT); r += 1
    b1 = r - 1
    put(ws, r, 1, "Poonawalla Fincorp does not disclose its borrowing mix and is left out. Source of bank shares: Table 4.3 (company presentations).", F_NOTE); r += 1
    Lb = f"LINEST(C{b0}:C{b1},B{b0}:B{b1},TRUE,TRUE)"
    for lbl_, fx, fmt, refv in [("Slope (CAR per percentage point of bank share)", f"=INDEX({Lb},1,1)", "0.0000%", E[ek]["bankshare_within_nonhfc"]["slope"]),
                                ("Standard error", f"=INDEX({Lb},2,1)", "0.0000%", None),
                                ("t statistic", f"=B{r}/B{r + 1}", "0.00", None),
                                ("p-value", f"=_xlfn.T.DIST.2T(ABS(B{r + 2}),INDEX({Lb},4,2))", P3, E[ek]["bankshare_within_nonhfc"]["p"]),
                                ("Companies", f"=COUNT(B{b0}:B{b1})", "0", E[ek]["bankshare_within_nonhfc"]["n"])]:
        put(ws, r, 1, lbl_, F_TX); put(ws, r, 2, fx, F_FX, fmt)
        if refv is not None: put(ws, r, 3, float(refv), F_REF, fmt)
        r += 1
    put(ws, r, 1, "Column C in blue italics: value reported in the dissertation.", F_NOTE)
    widths(ws, {"A": 44, "B": 14, "C": 11, "D": 11, "E": 11, "F": 7, **{L(j): 10 for j in range(7, 23)}})
    ws.freeze_panes = "C9"
    ES_CELLS[ek] = dict(ports=pc, firms=fc, sign=signrow, bs=b1 + 2)
    INFO.append((f"ES_{ek}", "Tables 8.2 and 8.4" if ek == "E1" else "Tables 8.3 and 8.4",
                 "Abnormal and cumulative abnormal returns of portfolios and companies, t-tests, sign test, bank-share regression", "Excel formulas"))

# ---- Nifty 500 robustness (Table C.1)
ws = sheet("ES_Nifty500", "Event study with the Nifty 500 as the market index",
           "Reproduces Table C.1. Same method as ES_E1 and ES_E2, with the Nifty 500 (column C of Data_Returns) as the benchmark.")
r = 4; N5 = {}
for ek, (d0, lbl, _) in EV.items():
    section(ws, r, 1, f"{ek}: {lbl}, day 0 = {pd.Timestamp(d0).strftime('%d %B %Y').lstrip('0')}", 22); r += 1
    hrow(ws, r, 1, ES_HDR, 30); r += 1
    ents = [(p, "Portfolio", c, None) for p, c, _ in PORTS] + [(p, "Difference", c, None) for p, c, _, _ in DIFFS]
    N5[ek], r = es_rows(ws, r, ek, "C", ents)
    r += 1
widths(ws, {"A": 44, "B": 11, "C": 11, "D": 11, "E": 11, "F": 7, **{L(j): 10 for j in range(7, 23)}})
INFO.append(("ES_Nifty500", "Table C.1", "Event study with the Nifty 500 as benchmark", "Excel formulas"))

# ---- placebo days (Figure 8.2)
ws = sheet("ES_PlaceboDays", "Day-0 abnormal return of the difference portfolios compared with every day of the estimation window",
           "Reproduces Figure 8.2 and the related figures in Chapter 8. For each ordinary day the abnormal return is the market-model residual; "
           "the empirical p-value is the share of the 220 days whose absolute value is at least as large as on day 0.")
c0 = 1; PD = {}
for ek in EV:
    for dn, dc, _, _ in DIFFS:
        r0 = ROW0[ek]; e1, e2 = r0 + EST[0], r0 + EST[1]
        cA, cB, cC, cD = L(c0), L(c0 + 1), L(c0 + 2), L(c0 + 3)
        section(ws, 4, c0, f"{ek}: {dn}", 4)
        lab = [("alpha", f"=INTERCEPT({cB}14:{cB}233,{cC}14:{cC}233)", "0.000%"), ("beta", f"=SLOPE({cB}14:{cB}233,{cC}14:{cC}233)", P3),
               ("Day-0 abnormal return", f"=Data_Returns!{dc}{r0}-({cB}5+{cB}6*Data_Returns!B{r0})", PCT),
               ("Days in estimation window", f"=COUNT({cD}14:{cD}233)", "0"),
               ("Days with |AR| at least as large as day 0", f"=SUMPRODUCT(--(ABS({cD}14:{cD}233)>=ABS({cB}7)))", "0"),
               ("Empirical p-value", f"={cB}9/{cB}8", P3)]
        refs = [None, None, E[ek]["placebo_days"][dn]["day0_ar"], E[ek]["placebo_days"][dn]["days"],
                E[ek]["placebo_days"][dn]["days_as_large"], E[ek]["placebo_days"][dn]["empirical_p"]]
        for i, ((t, fx, fmt), rv) in enumerate(zip(lab, refs)):
            put(ws, 5 + i, c0, t, F_TX); put(ws, 5 + i, c0 + 1, fx, F_FX, fmt)
            if rv is not None: put(ws, 5 + i, c0 + 2, float(rv), F_REF, fmt)
        put(ws, 11, c0 + 2, "Blue italics: dissertation", F_NOTE)
        hrow(ws, 13, c0, ["Date", "Difference return", "Nifty 50 return", "Abnormal return"], 30)
        for i, rr in enumerate(range(e1, e2 + 1)):
            r = 14 + i
            put(ws, r, c0, f"=Data_Returns!A{rr}", F_LK, "dd-mmm-yy")
            put(ws, r, c0 + 1, f"=Data_Returns!{dc}{rr}", F_LK, PCT)
            put(ws, r, c0 + 2, f"=Data_Returns!B{rr}", F_LK, PCT)
            put(ws, r, c0 + 3, f"={cB}{r}-(${cB}$5+${cB}$6*{cC}{r})", F_FX, PCT)
        widths(ws, {cA: 30, cB: 12, cC: 12, cD: 12, L(c0 + 4): 3})
        PD[(ek, dn)] = c0
        c0 += 5
ws.freeze_panes = "A14"
INFO.append(("ES_PlaceboDays", "Figure 8.2", "Day-0 abnormal return of the difference portfolios against the 220 estimation-window days", "Excel formulas"))

# =================================================================== firm-level tests
HC = os.path.join(HERE, "hand_collected")
HFCN = ["AAVAS", "Can Fin", "Home First", "LIC Housing", "PNB Housing", "Repco"]
MFIN = ["CreditAccess", "Arman"]
FIRM = {}
for nm, fn, col, title, srccol in [("cof", "cof_quarterly.csv", "COF_Pct", "Cost of funds (%)", "Note"),
                                   ("bank_share", "bank_share_quarterly.csv", "Bank_Share_Pct", "Bank share of borrowings (%)", "Definition")]:
    d = pd.read_csv(os.path.join(HC, fn))
    sh = "Data_Firm_COF" if nm == "cof" else "Data_Firm_BankShare"
    ws = sheet(sh, f"{title} by company and quarter",
               "Source: company investor presentations (deck and page in column H). Period: Pre = up to Q2 FY2024; Post = Q3 FY2024 to Q4 FY2025 "
               "(before the April 2025 rollback); Later = after Q4 FY2025, not used in the tests.")
    hrow(ws, 4, 1, ["Company", "Group", "Financial year", "Quarter", title, "Period (formula)", "Microfinance lender", "Source (presentation, page)", "Note"], 30)
    for i, row in d.iterrows():
        r = 5 + i
        grp = "HFC" if any(h in row.Company for h in HFCN) else "Non-HFC NBFC"
        put(ws, r, 1, row.Company, F_IN); put(ws, r, 2, grp, F_IN); put(ws, r, 3, row.FY, F_IN); put(ws, r, 4, row.Q, F_IN)
        put(ws, r, 5, None if pd.isna(row[col]) else float(row[col]), F_IN, "0.00")
        put(ws, r, 6, f'=IF(VALUE(MID(C{r},3,2))*10+VALUE(MID(D{r},2,1))<=242,"Pre",IF(VALUE(MID(C{r},3,2))*10+VALUE(MID(D{r},2,1))<=254,"Post","Later"))', F_FX)
        put(ws, r, 7, "Yes" if any(m in row.Company for m in MFIN) else "No", F_IN)
        put(ws, r, 8, row.Source_deck, F_IN); put(ws, r, 9, None if pd.isna(row[srccol]) else row[srccol], F_IN)
    widths(ws, {"A": 40, "B": 13, "C": 9, "D": 8, "E": 11, "F": 10, "G": 10, "H": 22, "I": 60})
    ws.freeze_panes = "B5"
    FIRM[nm] = (sh, 4 + len(d), title)

ws = sheet("Firm_Tests", "Change in cost of funds and bank share of borrowings: non-HFC NBFCs versus exempt HFCs",
           "Reproduces Table 9.2. For each company: average before (to Q2 FY2024) and after (Q3 FY2024 to Q4 FY2025), from the Data_Firm sheets. "
           "Welch t-test compares the changes of the two groups; paired t-tests compare before and after within each group.")
r = 4
for nm in ["cof", "bank_share"]:
    sh, last, title = FIRM[nm]
    res = S["firm"][nm]
    tab = pd.DataFrame(res["table"])
    tab["mfi"] = tab.company.str.contains("|".join(MFIN))
    tab["o"] = np.where(tab.group == "HFC", 2, np.where(tab.mfi, 1, 0))
    tab = tab.sort_values(["o", "company"]).reset_index(drop=True)
    section(ws, r, 1, title, 8); r += 1
    hrow(ws, r, 1, ["Company", "Group", "Microfinance", "Average before", "Average after", "Change (after minus before)"], 30); r += 1
    t0 = r
    for _, row in tab.iterrows():
        put(ws, r, 1, row.company, F_TX); put(ws, r, 2, row.group, F_TX); put(ws, r, 3, "Yes" if row.mfi else "No", F_TX)
        for j, per in enumerate(["Pre", "Post"]):
            put(ws, r, 4 + j, f'=AVERAGEIFS({sh}!$E$5:$E${last},{sh}!$A$5:$A${last},$A{r},{sh}!$F$5:$F${last},"{per}")', F_FX, "0.00")
        put(ws, r, 6, f"=E{r}-D{r}", F_FX, "0.00")
        r += 1
    nb = int((tab.group != "HFC").sum()); nmfi = int(((tab.group != "HFC") & tab.mfi).sum())
    NB = (t0, t0 + nb - 1); NBX = (t0, t0 + nb - nmfi - 1); HF = (t0 + nb, r - 1)
    rng = lambda c, a: f"{c}{a[0]}:{c}{a[1]}"
    r += 1
    hrow(ws, r, 1, ["Test", "Non-HFC NBFCs", "HFCs", "Difference", "t", "df", "p-value", "Dissertation p"], 30); r += 1
    for lbl, a, key in [("Welch t-test of the change: all NBFCs vs HFCs", NB, "all"), ("Welch t-test of the change: NBFCs excl. microfinance vs HFCs", NBX, "excl_mfi")]:
        A_, B_ = rng("F", a), rng("F", HF)
        put(ws, r, 1, lbl, F_TX)
        put(ws, r, 2, f"=AVERAGE({A_})", F_FX, "0.000"); put(ws, r, 3, f"=AVERAGE({B_})", F_FX, "0.000")
        put(ws, r, 4, f"=B{r}-C{r}", F_FX, "0.000")
        va, vb = f"(_xlfn.VAR.S({A_})/COUNT({A_}))", f"(_xlfn.VAR.S({B_})/COUNT({B_}))"
        put(ws, r, 5, f"=D{r}/SQRT({va}+{vb})", F_FX, "0.00")
        put(ws, r, 6, f"=({va}+{vb})^2/({va}^2/(COUNT({A_})-1)+{vb}^2/(COUNT({B_})-1))", F_FX, "0.0")
        put(ws, r, 7, f"=_xlfn.T.TEST({A_},{B_},2,3)", F_FX, P3)
        put(ws, r, 8, float(res[key]["p"]), F_REF, P3)
        r += 1
    r += 1
    hrow(ws, r, 1, ["Paired t-test, before vs after", "Companies", "Mean before", "Mean after", "Mean change", "t", "p-value", "Dissertation p"], 30); r += 1
    for lbl, a, key in [("Non-HFC NBFCs", NB, "paired_Non-HFC NBFC"), ("HFCs", HF, "paired_HFC")]:
        put(ws, r, 1, lbl, F_TX)
        put(ws, r, 2, f"=COUNT({rng('F', a)})", F_FX, "0")
        put(ws, r, 3, f"=AVERAGE({rng('D', a)})", F_FX, "0.000"); put(ws, r, 4, f"=AVERAGE({rng('E', a)})", F_FX, "0.000")
        put(ws, r, 5, f"=AVERAGE({rng('F', a)})", F_FX, "0.000")
        put(ws, r, 6, f"=E{r}/(_xlfn.STDEV.S({rng('F', a)})/SQRT(B{r}))", F_FX, "0.00")
        put(ws, r, 7, f"=_xlfn.T.TEST({rng('E', a)},{rng('D', a)},2,1)", F_FX, P3)
        put(ws, r, 8, float(res[key]["p"]), F_REF, P3)
        r += 1
    r += 2
widths(ws, {"A": 52, "B": 14, "C": 12, "D": 12, "E": 12, "F": 14, "G": 10, "H": 12})
INFO.append(("Firm_Tests", "Table 9.2", "Welch t-tests between groups and paired t-tests within groups, cost of funds and bank share", "Excel formulas"))
INFO.append(("Data_Firm_COF, Data_Firm_BankShare", "Inputs to Tables 4.3, 4.4 and 9.2", "Quarterly company data used in Firm_Tests", "Data"))
INFO.append(("Data_Monthly, Data_Prices, Data_Returns", "Inputs to Chapters 4, 7 and 8", "Monthly credit growth and gaps; daily prices, returns and portfolio returns", "Data and formulas"))

# =================================================================== README (first sheet)
ws = wb.create_sheet("README", 0)
ws.sheet_view.showGridLines = False
put(ws, 1, 1, "Capital Requirements and Credit Supply: statistical tests", F_TI)
put(ws, 2, 1, "Workbook 2 of 2 accompanying the dissertation (Nityam Gupta, FMS Delhi, MBA 2025-27). The first workbook, Dissertation_Data_Compilation.xlsx, holds the full data set.", F_NOTE)
hrow(ws, 4, 1, ["Sheet", "Reproduces", "Tests", "How computed"], 24)
for i, row in enumerate(INFO):
    for j, v in enumerate(row):
        c = put(ws, 5 + i, 1 + j, v, F_TX); c.alignment = Alignment(wrap_text=True, vertical="top")
r = 6 + len(INFO)
for t in ["Colour code: blue = input data; black = Excel formula; green = link to another sheet; blue italics = value reported in the dissertation (for comparison).",
          "Excel functions used: T.TEST, LINEST, INDEX, SLOPE, INTERCEPT, RSQ, STEYX, CORREL, MINVERSE, SUMXMY2, SUMPRODUCT, T.DIST.2T, F.DIST.RT, CHISQ.DIST.RT, BINOM.DIST, AVERAGEIFS.",
          "LINEST returns the coefficients in reverse order of the X columns; the coefficient tables put them back in the order Post, M2 to M12, then the constant.",
          "The Welch p-value uses T.TEST (type 3), which keeps the fractional Welch degrees of freedom.",
          "Newey-West standard errors cannot be computed with standard Excel functions; NeweyWest_Output holds the Python (statsmodels) output.",
          "The analysis itself was run in Python (pandas, statsmodels, SciPy); every value in this workbook matches the value reported in the dissertation."]:
    put(ws, r, 1, t, F_TX); r += 1
widths(ws, {"A": 34, "B": 30, "C": 80, "D": 34})
ORDER_SHEETS = ["README", "H1_Tests", "H2_Tests", "H2_Alt_Tests", "H1_Placebo", "H2_Placebo", "H2_Rollback", "H1_DropOne", "NeweyWest_Output",
                "ES_MarketModel", "ES_E1", "ES_E2", "ES_Nifty500", "ES_PlaceboDays", "Firm_Tests",
                "Data_Monthly", "Data_Prices", "Data_Returns", "Data_Firm_COF", "Data_Firm_BankShare"]
assert sorted(ORDER_SHEETS) == sorted(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER_SHEETS]
wb.active = 0
wb.save(OUTF)
print("saved", OUTF)
