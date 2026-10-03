"""Dissertation analysis: H2 (bank-dependent NBFCs faced higher borrowing costs after Nov-2023)
plus the lending extension (did lending growth fall?). Annual panel FY21-FY26 from Screener.in.

Run:  python 01_build_master_panel.py && python 03_h2_dissertation.py
Outputs: outputs/dissertation/*  and  Dissertation_Data.xlsx
"""
import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from did_tools import run, stars, B_REPS

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "outputs", "dissertation")
os.makedirs(OUT, exist_ok=True)
Y = "Interest_Expense_Ratio_Pct"
C = ["lag_RoA", "ln_Borr_lag"]

# ------------------------------------------------------------------ data
ann = pd.read_csv(os.path.join(HERE, "screener_annual_panel.csv"), parse_dates=["FY_End"])
orig = pd.read_csv(os.path.join(HERE, "nbfc_quarterly_data-v3.csv"))
base = orig[(orig.Fiscal_Year == "FY24") & (orig.Quarter == "Q2")].set_index("Company")
master = pd.read_csv(os.path.join(HERE, "master_panel_data_final.csv"), parse_dates=["Date_End"])

ann["Date_End"] = ann.FY_End
vb = pd.read_csv(os.path.join(HERE, "data_corrections", "bank_exposure_verified.csv")).set_index("Company")
ann["Bank_Exposure"] = ann.Company.map(vb.Verified_Bank_Share_Sep2023_Pct)
ann["Bank_Exposure_CSV"] = ann.Company.map(base.Bank_Borrowing_Share_Pct)
ann = ann.sort_values(["Company", "FY"])
ann["lag_RoA"] = ann.groupby("Company").RoA_annual_Pct.shift(1)
ann["ln_Borr_lag"] = np.log(ann.groupby("Company").Borrowings.shift(1))
# total loan book (AUM) at fiscal-year end from the quarterly panel (Q4); verified for Shriram
q4 = orig[orig.Quarter == "Q4"].assign(FY=lambda d: 2000 + d.Fiscal_Year.str[2:].astype(int))
q4 = q4.set_index(["Company", "FY"]).Total_AUM_Cr
m4 = master[master.Quarter == "Q4"].assign(FY=lambda d: 2000 + d.Fiscal_Year.str[2:].astype(int))
q4.update(m4.set_index(["Company", "FY"]).Total_AUM_Cr)
ann["Total_AUM_Cr"] = [q4.get((c, f), np.nan) for c, f in zip(ann.Company, ann.FY)]
ann["AUM_Growth_Pct"] = ann.groupby("Company").Total_AUM_Cr.transform(lambda s: 100 * np.log(s).diff())
ann = ann[ann.FY >= 2021].copy()                     # FY20 only used for lags
ann["Post"] = (ann.FY >= 2024).astype(int)
ann["Post_x_Exposure"] = ann.Post * ann.Bank_Exposure
rep = master.groupby("Fiscal_Year").Repo_Rate_Pct.mean()
rep.index = [2000 + int(i[2:]) for i in rep.index]
ann["Repo_Avg_Pct"] = ann.FY.map(rep)
ann["Exposure_x_Repo"] = ann.Bank_Exposure * ann.Repo_Avg_Pct
for y in [2021, 2022, 2024, 2025, 2026]:
    ann[f"Exposure_x_FY{y % 100}"] = (ann.FY == y) * ann.Bank_Exposure
med = ann.drop_duplicates("Company").Bank_Exposure.median()
ann["High_Exposure"] = (ann.Bank_Exposure > med).astype(int)
ann["Post_x_HighExposure"] = ann.Post * ann.High_Exposure
pw21 = ann.Company.str.contains("Poonawalla") & (ann.FY == 2021)

# ------------------------------------------------------------------ models
M = {}
M["H2-1 Main"] = run("", ann, Y, ["Post_x_Exposure"], C, min_obs=3)
M["H2-2 No controls"] = run("", ann, Y, ["Post_x_Exposure"], [], min_obs=3)
M["H2-3 Drop Poonawalla FY21 (Magma)"] = run("", ann[~pw21], Y, ["Post_x_Exposure"], C, min_obs=3)
M["H2-4 Drop COVID year FY21"] = run("", ann[ann.FY >= 2022], Y, ["Post_x_Exposure"], C, min_obs=3)
M["H2-5 Drop transition year FY24"] = run("", ann[ann.FY != 2024], Y, ["Post_x_Exposure"], C, min_obs=3)
M["H2-6 + Exposure x Repo"] = run("", ann, Y, ["Post_x_Exposure"], C + ["Exposure_x_Repo"], min_obs=3)
M["H2-7 High/low exposure dummy"] = run("", ann, Y, ["Post_x_HighExposure"], C, min_obs=3)
ES = run("", ann, Y, [f"Exposure_x_FY{y % 100}" for y in [2021, 2022, 2024, 2025, 2026]], C, min_obs=3)
pre = ann[ann.FY <= 2023].copy()
pre["Placebo_x_Exposure"] = (pre.FY >= 2023) * pre.Bank_Exposure
PL = run("", pre, Y, ["Placebo_x_Exposure"], [], min_obs=2)
LEND = run("", ann, "AUM_Growth_Pct", ["Post_x_Exposure"], C, min_obs=3)
loo = []
for f in sorted(ann.Company.unique()):
    r = run("", ann[ann.Company != f], Y, ["Post_x_Exposure"], C, min_obs=3)["res"][0]
    loo.append(dict(Firm_dropped=f, Coefficient=r["beta"], CR1_SE=r["se_cr1"], WCB_p=r["p_wcb"]))
loo = pd.DataFrame(loo)


def row(name, m, j=0):
    r = m["res"][j]
    return dict(Specification=name, Term=m["treat"][j], Coefficient=round(r["beta"], 4),
                CR1_SE=round(r["se_cr1"], 4), WCB_p=round(r["p_wcb"], 3), Stars=stars(r["p_wcb"]),
                CI95_low=round(r["ci_lo"], 4), CI95_high=round(r["ci_hi"], 4),
                N=m["n"], Firms=m["firms"], Adj_R2=round(m["adj_r2"], 3), Within_R2=round(m["within_r2"], 3))


res = pd.DataFrame([row(k, v) for k, v in M.items()] +
                   [row("Placebo: fake policy FY23 (FY21-23 only)", PL),
                    row("Extension: total loan (AUM) growth", LEND)])
es = pd.DataFrame([row(f"Event year FY{t[-2:]}", ES, j) for j, t in enumerate(ES["treat"])])
res.to_csv(os.path.join(OUT, "h2_results.csv"), index=False)
es.to_csv(os.path.join(OUT, "h2_event_years.csv"), index=False)
loo.to_csv(os.path.join(OUT, "h2_leave_one_out.csv"), index=False)
print(res.to_string()); print(es.to_string())
print("LOO p range:", loo.WCB_p.min(), loo.WCB_p.max(), " coef range:", loo.Coefficient.min(), loo.Coefficient.max())

# descriptives
desc_vars = [(Y, "Interest expense / avg borrowings (%)"), ("Bank_Exposure", "Bank borrowing share, Sep-2023 (%)"),
             ("RoA_annual_Pct", "Return on assets (%)"), ("Borrowings", "Borrowings (Rs cr)"),
             ("Total_AUM_Cr", "Total AUM (Rs cr)"), ("AUM_Growth_Pct", "AUM growth (100 x log diff)"),
             ("Repo_Avg_Pct", "Repo rate, FY average (%)")]
drows = []
for v, nm in desc_vars:
    for per, m in [("Pre (FY21-FY23)", ann.Post == 0), ("Post (FY24-FY26)", ann.Post == 1)]:
        s = ann.loc[m, v].dropna()
        drows.append(dict(Variable=nm, Period=per, N=len(s), Mean=s.mean(), SD=s.std(), Min=s.min(), Max=s.max()))
desc = pd.DataFrame(drows).round(2)
desc.to_csv(os.path.join(OUT, "h2_descriptives.csv"), index=False)

# ------------------------------------------------------------------ figures
INK, MUTED, BLUE, ORANGE, GRID = "#1f2430", "#6b7280", "#2b6cb0", "#c2410c", "#e5e7eb"


def tidy(ax):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color=GRID, lw=.6)


g = ann.assign(Group=np.where(ann.High_Exposure == 1, "High bank dependence", "Low bank dependence"))
tr = g.groupby(["FY", "Group"])[Y].mean().unstack()
fig, ax = plt.subplots(figsize=(7.5, 4))
for col, c in [("High bank dependence", ORANGE), ("Low bank dependence", BLUE)]:
    ax.plot(tr.index, tr[col], color=c, lw=2, marker="o", ms=5)
    ax.annotate(col, (tr.index[-1], tr[col].iloc[-1]), xytext=(6, 0), textcoords="offset points", color=INK, fontsize=8, va="center")
ax.axvline(2023.6, color=MUTED, lw=1, ls="--")
ax.text(2023.65, ax.get_ylim()[0] + .05, "Nov-2023 rule", color=MUTED, fontsize=8)
ax.set_xticks(tr.index); ax.set_xticklabels([f"FY{y % 100}" for y in tr.index])
ax.set_ylabel("Interest expense / avg borrowings (%)")
ax.set_title(f"Average borrowing cost: firms above vs below median bank share ({med:.1f}%)", loc="left", fontsize=10, color=INK)
ax.set_xlim(2020.7, 2026.9)
tidy(ax); fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig1_cost_trends.png"), dpi=200); plt.close(fig)

e = es.copy()
e["FY"] = [2021, 2022, 2024, 2025, 2026]
e = pd.concat([e, pd.DataFrame({"FY": [2023], "Coefficient": [0.], "CI95_low": [0.], "CI95_high": [0.]})]).sort_values("FY")
fig, ax = plt.subplots(figsize=(7.5, 4))
ax.axhline(0, color=MUTED, lw=1)
ax.axvline(2023.6, color=MUTED, lw=1, ls="--")
ax.vlines(e.FY, e.CI95_low, e.CI95_high, color=BLUE, lw=2)
ax.plot(e.FY, e.Coefficient, "o", color=BLUE, ms=7)
ax.annotate("reference year", (2023, 0), xytext=(-30, 12), textcoords="offset points", fontsize=8, color=INK)
ax.set_xticks(e.FY); ax.set_xticklabels([f"FY{y % 100}" for y in e.FY])
ax.set_ylabel("Effect of +1pp bank share on\nborrowing cost (pp), vs FY23")
ax.set_title("Year-by-year effect with 95% wild-bootstrap intervals", loc="left", fontsize=10, color=INK)
tidy(ax); fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig2_event_years.png"), dpi=200); plt.close(fig)

fig, ax = plt.subplots(figsize=(7.5, 3.6))
ex = ann.drop_duplicates("Company").sort_values("Bank_Exposure")
ax.barh(ex.Company.str.replace(r" (Ltd\.|Limited)$", "", regex=True), ex.Bank_Exposure, color=BLUE, height=.7)
ax.set_xlabel("Bank borrowing share of total borrowings, Sep-2023 (%)")
ax.set_title("Treatment intensity across the 16 NBFCs", loc="left", fontsize=10, color=INK)
for s in ["top", "right"]:
    ax.spines[s].set_visible(False)
ax.tick_params(axis="y", labelsize=7)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig3_exposure.png"), dpi=200); plt.close(fig)

# ------------------------------------------------------------------ consolidated Excel (only data used)
panel_cols = ["Company", "FY", "FY_End", "Interest", "Borrowings", "Borrowings_avg", Y, "NetProfit", "TotalAssets",
              "Assets_avg", "RoA_annual_Pct", "lag_RoA", "ln_Borr_lag", "Total_AUM_Cr", "AUM_Growth_Pct",
              "Bank_Exposure", "High_Exposure", "Post", "Post_x_Exposure", "Repo_Avg_Pct"]
panel = ann[panel_cols].rename(columns={"Interest": "Interest_Expense_Cr", "Borrowings": "Borrowings_Cr",
                                        "Borrowings_avg": "Avg_Borrowings_Cr", "NetProfit": "Net_Profit_Cr",
                                        "TotalAssets": "Total_Assets_Cr", "Assets_avg": "Avg_Total_Assets_Cr"})
panel["FY_End"] = panel.FY_End.dt.date
expo = base[["Bank_Borrowing_Share_Pct"]].reset_index().rename(columns={"Bank_Borrowing_Share_Pct": "Bank_Borrowing_Share_Sep2023_Pct"})
expo["Source"] = "nbfc_quarterly_data-v3.csv, Q2 FY24 (not yet verified against company filings)"
macro = master.drop_duplicates(["Fiscal_Year", "Quarter"])[["Fiscal_Year", "Quarter", "Date_End", "Repo_Rate_Pct",
                                                            "Bond_Yield_5Y_Pct", "GDP_Growth_YoY_Pct"]]
macro["Date_End"] = macro.Date_End.dt.date
dictionary = pd.DataFrame([
    ("Interest_Expense_Cr", "Annual finance cost (Screener.in 'Interest'), Rs crore", "Screener.in annual P&L"),
    ("Borrowings_Cr", "Total borrowings at fiscal year-end, Rs crore", "Screener.in balance sheet"),
    ("Avg_Borrowings_Cr", "Average of opening and closing borrowings", "Computed"),
    (Y, "Borrowing cost = Interest expense / Avg borrowings x 100 (dependent variable)", "Computed"),
    ("RoA_annual_Pct", "Net profit / average total assets x 100", "Computed from Screener.in"),
    ("lag_RoA", "Previous year's RoA (control)", "Computed"),
    ("ln_Borr_lag", "Natural log of previous year's borrowings (size control)", "Computed"),
    ("Total_AUM_Cr", "Assets under management at March quarter-end", "nbfc_quarterly_data-v3.csv; Shriram from filings"),
    ("AUM_Growth_Pct", "100 x log change in year-end AUM (extension test)", "Computed"),
    ("Bank_Exposure", "Share of borrowings from banks at 30-Sep-2023 (frozen pre-policy)", "nbfc_quarterly_data-v3.csv"),
    ("High_Exposure", f"1 if Bank_Exposure above median ({med:.1f}%)", "Computed"),
    ("Post", "1 for FY24-FY26 (policy announced 16-Nov-2023, in FY24)", "Computed"),
    ("Post_x_Exposure", "Treatment variable = Post x Bank_Exposure", "Computed"),
    ("Repo_Avg_Pct", "RBI policy repo rate, day-weighted fiscal-year average", "RBI DBIE"),
], columns=["Variable", "Definition", "Source"])
xl = os.path.join(HERE, "Dissertation_Data.xlsx")
with pd.ExcelWriter(xl, engine="openpyxl") as w:
    dictionary.to_excel(w, sheet_name="Variable_Dictionary", index=False)
    panel.round(4).to_excel(w, sheet_name="Annual_Panel_FY21_FY26", index=False)
    expo.to_excel(w, sheet_name="Bank_Exposure_Sep2023", index=False)
    macro.round(3).to_excel(w, sheet_name="Macro_Quarterly", index=False)
    desc.to_excel(w, sheet_name="Descriptive_Stats", index=False)
    res.to_excel(w, sheet_name="Regression_Results", index=False)
    es.to_excel(w, sheet_name="Event_Year_Results", index=False)
    loo.round(4).to_excel(w, sheet_name="Leave_One_Out", index=False)
print("saved", xl)
