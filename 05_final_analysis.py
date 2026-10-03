"""Final analysis for the dissertation.

H1  RBI consumer-credit risk weights (Nov-2023) -> slower bank credit to the targeted consumer categories
    (credit cards, other personal loans) relative to exempt household categories (housing, vehicle, education).
H2  RBI risk weights on bank exposures to NBFCs (Nov-2023, HFCs exempt) -> slower bank credit to NBFCs (ex-HFC)
    relative to banks' total non-food credit.
Supporting firm-level evidence from 382 hand-read investor presentations (16 listed NBFCs/HFCs).

Run: python 05_final_analysis.py   -> outputs/final/*  (results.json, CSVs, PNG figures) + Dissertation_Data.xlsx
"""
import json
import os
import re
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from linearmodels.panel import PanelOLS

from did_tools import run

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "outputs", "final")
os.makedirs(OUT, exist_ok=True)
R = {}                       # every number quoted in the dissertation lives here
INK, MUTED, GRID = "#1f2430", "#6b7280", "#e5e7eb"
C_T, C_C = "#c2410c", "#2b6cb0"          # treated / control colours
POLICY, ROLLBACK = pd.Timestamp("2023-11-16"), pd.Timestamp("2025-04-01")


def tidy(ax):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color=GRID, lw=.6)


# =========================================================== 1. RBI sectoral bank credit
raw = pd.read_excel(os.path.join(HERE, "RBI_Sectoral_Bank_Credit.xlsx"), header=None)
dates = pd.to_datetime(raw.iloc[6, 3:].astype(str).str.replace(r"\s+", " ", regex=True), format="%B %d, %Y")
ROWS = {"Non-food credit": 9, "NBFCs (incl. HFCs)": 11, "HFCs": 12, "Housing": 13, "Credit cards": 15,
        "Education": 16, "Vehicle loans": 17, "Gold loans": 18, "Other personal loans": 19}
lev = pd.DataFrame({k: pd.to_numeric(raw.iloc[r, 3:], errors="coerce").values for k, r in ROWS.items()}, index=dates)
lev["NBFCs ex-HFC"] = lev["NBFCs (incl. HFCs)"] - lev["HFCs"]
lev.index.name = "Date"
g = 100 * np.log(lev).diff()                       # monthly log growth, pp
MERGER = pd.Timestamp("2023-07-28")                # HDFC Ltd -> HDFC Bank merger month
for c in ["Housing", "NBFCs (incl. HFCs)", "HFCs", "NBFCs ex-HFC", "Non-food credit"]:
    g.loc[MERGER, c] = np.nan
WIN = ("2021-04-01", "2025-01-31")                 # main window: ends before the Feb-2025 announcement
POST = "2023-11-30"                                # first monthly observation after 16-Nov-2023
TREAT = ["Credit cards", "Other personal loans"]
CTRL = ["Housing", "Vehicle loans", "Education"]


def its(gap, brk, lags=6):
    gap = gap.dropna()
    X = sm.add_constant((gap.index >= brk).astype(float))
    r = sm.OLS(gap.values, X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
    return dict(pre_mean=r.params[0], change=r.params[1], se=r.bse[1], p=r.pvalues[1], n=int(len(gap)),
                annualised_change=12 * r.params[1])


def panel(gg, tr, ct, brk):
    L = gg[tr + ct].stack().rename("g").reset_index()
    L.columns = ["date", "series", "g"]
    L["Treat_x_Post"] = (L.series.isin(tr) & (L.date >= brk)).astype(float)
    p = L.set_index(["series", "date"])
    r = PanelOLS(p.g, p[["Treat_x_Post"]], entity_effects=True, time_effects=True).fit(
        cov_type="kernel", kernel="bartlett", bandwidth=6)
    return dict(beta=float(r.params.iloc[0]), se=float(r.std_errors.iloc[0]), p=float(r.pvalues.iloc[0]),
                n=int(r.nobs), annualised=12 * float(r.params.iloc[0]))


W = g[WIN[0]:WIN[1]]
gap1 = W[TREAT].mean(1) - W[CTRL].mean(1)
R["H1"] = dict(its=its(gap1, POST), panel=panel(W, TREAT, CTRL, POST),
               placebo=its((g["2021-04":"2023-10"][TREAT].mean(1) - g["2021-04":"2023-10"][CTRL].mean(1)), "2022-11-30"),
               placebo_panel=panel(g["2021-04":"2023-10"], TREAT, CTRL, "2022-11-30"))
rob = []
for lbl, tr, ct, w in [("Drop credit cards", ["Other personal loans"], CTRL, WIN),
                       ("Drop other personal loans", ["Credit cards"], CTRL, WIN),
                       ("Drop housing (merger-affected)", TREAT, ["Vehicle loans", "Education"], WIN),
                       ("Drop vehicle loans", TREAT, ["Housing", "Education"], WIN),
                       ("Drop education", TREAT, ["Housing", "Vehicle loans"], WIN),
                       ("Add gold loans to control", TREAT, CTRL + ["Gold loans"], WIN),
                       ("Shorter pre-period (from Apr-2022)", TREAT, CTRL, ("2022-04-01", WIN[1])),
                       ("Shorter post-period (to Oct-2024)", TREAT, CTRL, (WIN[0], "2024-10-31"))]:
    rr = panel(g[w[0]:w[1]], tr, ct, POST)
    rob.append(dict(Specification=lbl, Coefficient=round(rr["beta"], 2), SE=round(rr["se"], 2), p=round(rr["p"], 3), N=rr["n"]))
R["H1"]["robustness"] = rob
pd.DataFrame(rob).to_csv(os.path.join(OUT, "h1_robustness.csv"), index=False)

# H2 aggregate
gap2 = W["NBFCs ex-HFC"] - W["Non-food credit"]
R["H2"] = dict(its=its(gap2, POST),
               placebo=its((g["2021-04":"2023-10"]["NBFCs ex-HFC"] - g["2021-04":"2023-10"]["Non-food credit"]), "2022-11-30"),
               alt_vs_hfc=its(W["NBFCs ex-HFC"] - W["HFCs"], POST),
               alt_all_nbfc=its(W["NBFCs (incl. HFCs)"] - W["Non-food credit"], POST),
               rollback=its((g["2023-12":]["NBFCs ex-HFC"] - g["2023-12":]["Non-food credit"]), "2025-04-30"))
# growth tables (annualised average monthly log growth)
PER = [("Pre-policy (Apr-21 to Oct-23)", "2021-04-01", "2023-10-31"), ("Post-policy (Dec-23 to Mar-25)", "2023-12-01", "2025-03-31"),
       ("After rollback (Apr-25 to Jul-26)", "2025-04-01", "2026-07-31")]
gt = pd.DataFrame({p: (g[a:b].mean() * 12).round(1) for p, a, b in PER})
gt = gt.loc[TREAT + CTRL + ["Gold loans", "NBFCs ex-HFC", "HFCs", "NBFCs (incl. HFCs)", "Non-food credit"]]
gt.to_csv(os.path.join(OUT, "rbi_growth_table.csv"))
R["growth_table"] = {k: v for k, v in gt.to_dict(orient="index").items()}

# Figures from RBI data: chained indices (merger month set to zero growth), Oct-2023 = 100
idx = g.fillna(0).cumsum()
base = idx.loc["2023-10"].iloc[-1]
idx = 100 * np.exp((idx - base) / 100)
idx = idx["2021-04":]
fig, axs = plt.subplots(1, 2, figsize=(10, 3.9), sharey=False)
for ax, cols, title in [(axs[0], [(TREAT, "Targeted: cards + other personal loans", C_T), (CTRL, "Exempt: housing, vehicle, education", C_C)], "H1: consumer credit (index, Oct-2023 = 100)"),
                        (axs[1], [(["NBFCs ex-HFC"], "Bank credit to NBFCs (ex-HFC)", C_T), (["Non-food credit"], "Total non-food bank credit", C_C)], "H2: bank credit to NBFCs (index, Oct-2023 = 100)")]:
    for cl, lab, col in cols:
        ser = 100 * np.exp(np.log(idx[cl] / 100).mean(1))
        ax.plot(ser.index, ser, color=col, lw=2, label=lab)
    ax.axvline(POLICY, color=MUTED, ls="--", lw=1); ax.axvline(ROLLBACK, color=MUTED, ls=":", lw=1)
    ax.text(POLICY, ax.get_ylim()[0], " Nov-23 rule", fontsize=7.5, color=MUTED, va="bottom")
    ax.text(ROLLBACK, ax.get_ylim()[0], " Apr-25 rollback", fontsize=7.5, color=MUTED, va="bottom")
    ax.set_title(title, loc="left", fontsize=9.5, color=INK); ax.legend(frameon=False, fontsize=7.5, loc="upper left"); tidy(ax)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig1_rbi_indices.png"), dpi=200); plt.close(fig)

# H1 dynamics: quarterly averages of the monthly gap, relative to pre-policy mean
for name, gap, fn in [("H1", g[TREAT].mean(1) - g[CTRL].mean(1), "fig2_h1_gap.png"),
                      ("H2", g["NBFCs ex-HFC"] - g["Non-food credit"], "fig3_h2_gap.png")]:
    q = gap["2021-04":].groupby(gap["2021-04":].index.to_period("Q")).mean()
    fig, ax = plt.subplots(figsize=(8, 3.4))
    ax.bar(q.index.to_timestamp(), q.values, width=60, color=[C_T if p.start_time >= pd.Timestamp("2023-12-01") else "#9ca3af" for p in q.index])
    ax.axhline(0, color=MUTED, lw=.8); ax.axvline(POLICY, color=MUTED, ls="--", lw=1)
    pm = gap[WIN[0]:"2023-10-31"].mean(); ax.axhline(pm, color=C_C, lw=1, ls="-.")
    ax.text(pd.Timestamp("2021-04-01"), pm, " pre-policy average", color=C_C, fontsize=7.5, va="bottom")
    ax.set_ylabel("Growth gap, pp per month"); tidy(ax)
    ax.set_title(("Targeted minus exempt consumer credit" if name == "H1" else "Bank credit to NBFCs (ex-HFC) minus total non-food credit") + ": quarterly average of monthly growth gap", loc="left", fontsize=9, color=INK)
    fig.tight_layout(); fig.savefig(os.path.join(OUT, fn), dpi=200); plt.close(fig)

# =========================================================== 2. Firm-level (hand-collected)
HC = os.path.join(HERE, "hand_collected")
cof = pd.read_csv(os.path.join(HC, "cof_quarterly.csv"))
bsh = pd.read_csv(os.path.join(HC, "bank_share_quarterly.csv"))
ver = pd.read_csv(os.path.join(HERE, "data_corrections", "bank_exposure_verified.csv"))
HFC = ["AAVAS", "Can Fin", "Home First", "LIC Housing", "PNB Housing", "Repco"]
QEND = pd.period_range("2020Q2", periods=24, freq="Q").to_timestamp(how="end").normalize()


def prep(d):
    d = d.copy()
    d["t"] = (d.FY.str[2:].astype(int) - 21) * 4 + d.Q.str[1].astype(int) - 1
    d["Date_End"] = QEND[d.t]
    d["NonHFC"] = (~d.Company.str.contains("|".join(HFC))).astype(int)
    d["Exp"] = d.Company.map(ver.set_index("Company").Verified_Bank_Share_Sep2023_Pct)
    d["Post"] = (d.t >= 14).astype(int); d["Rev"] = (d.t >= 20).astype(int)
    return d


cof, bsh = prep(cof), prep(bsh)
cof["Post_x_Exp"] = cof.Post * cof.Exp
m = run("", cof, "COF_Pct", ["Post_x_Exp"], [], min_obs=8)
R["firm_cof"] = dict(beta=m["res"][0]["beta"], se=m["res"][0]["se_cr1"], p=m["res"][0]["p_wcb"], n=m["n"], firms=m["firms"])
cof["NxPost"] = cof.NonHFC * cof.Post
m = run("", cof, "COF_Pct", ["NxPost"], [], min_obs=8)
R["firm_cof_hfc"] = dict(beta=m["res"][0]["beta"], se=m["res"][0]["se_cr1"], p=m["res"][0]["p_wcb"], n=m["n"], firms=m["firms"])
# COF event time (half-years, ref FY23H2)
cof["h"] = cof.t // 2
cols = []
for h in range(12):
    if h == 5: continue
    c = f"E{h}"; cof[c] = (cof.h == h) * cof.Exp; cols.append(c)
m = run("", cof, "COF_Pct", cols, [], min_obs=8)
ev = pd.DataFrame([dict(half=f"FY{21 + int(c[1:]) // 2}H{int(c[1:]) % 2 + 1}", h=int(c[1:]), beta=r["beta"], lo=r["ci_lo"], hi=r["ci_hi"], p=r["p_wcb"]) for c, r in zip(cols, m["res"])])
ev = pd.concat([ev, pd.DataFrame([dict(half="FY23H2", h=5, beta=0, lo=0, hi=0, p=np.nan)])]).sort_values("h")
ev.to_csv(os.path.join(OUT, "firm_cof_event.csv"), index=False)
fig, ax = plt.subplots(figsize=(8, 3.5))
x = np.arange(len(ev)); ax.axhline(0, color=MUTED, lw=1)
ax.vlines(x, ev.lo * 30, ev.hi * 30, color=C_C, lw=2); ax.plot(x, ev.beta * 30, "o", color=C_C)
ax.axvline(5.5, color=MUTED, ls="--", lw=1); ax.set_xticks(x); ax.set_xticklabels(ev.half, fontsize=7.5)
ax.set_ylabel("Cost-of-funds difference (pp)\nfor +30pp bank share"); tidy(ax)
ax.set_title("Cost of funds: bank-dependent vs other NBFCs, by half-year (ref. FY23 H2; 95% wild-bootstrap CI)", loc="left", fontsize=9, color=INK)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig4_firm_cof_event.png"), dpi=200); plt.close(fig)

# Bank share: non-HFC NBFCs vs HFCs
bsh["NxPost"] = bsh.NonHFC * bsh.Post; bsh["NxRev"] = bsh.NonHFC * bsh.Rev
m = run("", bsh[bsh.t < 20], "Bank_Share_Pct", ["NxPost"], [], min_obs=4)
R["firm_bankshare"] = dict(beta=m["res"][0]["beta"], se=m["res"][0]["se_cr1"], p=m["res"][0]["p_wcb"], n=m["n"], firms=m["firms"])
m = run("", bsh, "Bank_Share_Pct", ["NxPost", "NxRev"], [], min_obs=4)
R["firm_bankshare_rollback"] = dict(beta_post=m["res"][0]["beta"], p_post=m["res"][0]["p_wcb"], beta_rev=m["res"][1]["beta"], p_rev=m["res"][1]["p_wcb"], n=m["n"], firms=m["firms"])
bsh["h"] = bsh.t // 2
cols = []
for h in range(12):
    if h == 5: continue
    c = f"B{h}"; bsh[c] = bsh.NonHFC * (bsh.h == h); cols.append(c)
m = run("", bsh, "Bank_Share_Pct", cols, [], min_obs=4)
evb = pd.DataFrame([dict(half=f"FY{21 + int(c[1:]) // 2}H{int(c[1:]) % 2 + 1}", h=int(c[1:]), beta=r["beta"], lo=r["ci_lo"], hi=r["ci_hi"], p=r["p_wcb"]) for c, r in zip(cols, m["res"])])
evb = pd.concat([evb, pd.DataFrame([dict(half="FY23H2", h=5, beta=0, lo=0, hi=0, p=np.nan)])]).sort_values("h")
evb.to_csv(os.path.join(OUT, "firm_bankshare_event.csv"), index=False)
R["firm_bankshare_last"] = dict(beta=float(evb.iloc[-1].beta), p=float(evb.iloc[-1].p))
pre = bsh[bsh.t < 14].copy(); pre["Fake"] = pre.NonHFC * (pre.t >= 10)
m = run("", pre, "Bank_Share_Pct", ["Fake"], [], min_obs=3)
R["firm_bankshare_placebo"] = dict(beta=m["res"][0]["beta"], p=m["res"][0]["p_wcb"])
grp = bsh.groupby([bsh.t // 2, "NonHFC"]).Bank_Share_Pct.mean().unstack()
fig, ax = plt.subplots(figsize=(8, 3.5))
lab = [f"FY{21 + h // 2}H{h % 2 + 1}" for h in grp.index]
ax.plot(lab, grp[1], color=C_T, lw=2, marker="o", ms=4, label="Non-HFC NBFCs (higher bank risk weight)")
ax.plot(lab, grp[0], color=C_C, lw=2, marker="o", ms=4, label="HFCs (exempt)")
ax.axvline(5.5, color=MUTED, ls="--", lw=1); ax.legend(frameon=False, fontsize=8); tidy(ax)
ax.set_ylabel("Bank share of borrowings (%)"); ax.tick_params(axis="x", labelsize=7.5)
ax.set_title("Average bank share of borrowings, hand-collected from investor presentations", loc="left", fontsize=9, color=INK)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig5_bank_share.png"), dpi=200); plt.close(fig)

# Bajaj within-firm (segments), FY21-FY25 consistent labels
seg = pd.read_csv(os.path.join(HC, "segments_bajaj.csv"))
rows = []
MAP = {"Auto Finance": "Two & Three wheeler Finance", "Two & Three-wheeler Finance": "Two & Three wheeler Finance",
       "Sales Finance": "Urban Sales Finance", "Consumer B2C": "Urban B2C", "Securities Lending": "Loan Against Securities",
       "Mortgages Business": "Mortgages", "SME lending": "SME Lending", "Gold Loans": "Rural B2C", "Car Loans": "SME Lending",
       "Urban B2C Loans": "Urban B2C", "Rural B2C Loans": "Rural B2C", "Two and Three-wheeler Finance": "Two & Three wheeler Finance"}
for _, r in seg.iterrows():
    for kv in r.Segments.split("; "):
        k, v = kv.split("=")
        k = re.sub(r"\s*\(.*?\)|\*|Business$", "", k.strip()).strip()
        rows.append((r.FY, r.Q, MAP.get(k, k), float(v.replace(",", ""))))
s = pd.DataFrame(rows, columns=["FY", "Q", "Seg", "AUM"]).groupby(["FY", "Q", "Seg"]).AUM.sum().reset_index()
s["t"] = (s.FY.str[2:].astype(int) - 21) * 4 + s.Q.str[1].astype(int) - 1
s = s[s.t <= 19].sort_values(["Seg", "t"])
s["g"] = 100 * s.groupby("Seg").AUM.transform(lambda x: np.log(x).diff())
s = s[s.Seg.map(s.groupby("Seg").size()) >= 18].dropna(subset=["g"])
s["Treat"] = s.Seg.isin(["Urban Sales Finance", "Urban B2C", "Rural Sales Finance", "Rural B2C"]).astype(int)
s["TxP"] = s.Treat * (s.t >= 14)
p = s.set_index(["Seg", "t"])
r = PanelOLS(p.g, p[["TxP"]], entity_effects=True, time_effects=True).fit(cov_type="kernel", kernel="bartlett", bandwidth=4)
R["bajaj_within"] = dict(beta=float(r.params.iloc[0]), se=float(r.std_errors.iloc[0]), p=float(r.pvalues.iloc[0]), n=int(r.nobs), segments=int(s.Seg.nunique()))
share = seg.assign(sh=100 * seg.Affected_Consumer_AUM_Cr / seg.Total_AUM_Cr)
R["bajaj_share"] = {f"{a}{b}": round(x, 1) for a, b, x in zip(share.FY, share.Q, share.sh)}

# Company disclosures (verbatim quotes, hand-read)
disc = pd.DataFrame([
    ("Bajaj Finance Ltd.", "Q3 FY24 investor presentation (24-3), item 25",
     "RBI increased risk weights on consumer credit exposure from 100% to 125% which had an impact of 290 bps on the Company's CRAR. Adjusted for this change CRAR would have been 26.77%.", "Capital ratio cut by 2.9 pp"),
    ("Bajaj Finance Ltd.", "Q3 FY24 investor presentation (24-3), item 11",
     "Given the increase in risk weights and higher incremental cost of funds, the Company has increased rates across all portfolios by 20-30 bps.", "Lending rates raised 20-30 bps"),
    ("SBI Cards and Payment Services Ltd.", "Q3 FY24 investor presentation (24-3), capital adequacy slide",
     "CRAR impacted by ~400 bps due to increase in risk weight by RBI", "Capital ratio cut by ~4 pp"),
], columns=["Company", "Source", "Quote", "Implication"])
disc.to_csv(os.path.join(OUT, "company_disclosures.csv"), index=False)

# Sample table
sample = ver[["Company", "CSV_Value_Pct", "Verified_Bank_Share_Sep2023_Pct", "Definition_in_filing"]].copy()
sample["Type"] = np.where(sample.Company.str.contains("|".join(HFC)), "HFC (exempt from bank-NBFC risk weight)", "NBFC")
sample["COF quarters"] = sample.Company.map(cof.groupby("Company").size()).fillna(0).astype(int)
sample["Bank-share obs"] = sample.Company.map(bsh.groupby("Company").size()).fillna(0).astype(int)
sample.to_csv(os.path.join(OUT, "firm_sample.csv"), index=False)

R["counts"] = dict(decks=382, firms=16, cof_obs=int(len(cof)), bankshare_obs=int(len(bsh)), bankshare_firms=int(bsh.Company.nunique()),
                   rbi_months=int(lev.shape[0]), rbi_first=str(lev.index.min().date()), rbi_last=str(lev.index.max().date()))

# =========================================================== 3. Consolidated Excel (only data used)
gl = g.copy(); gl.columns = [c + " (growth, pp/month)" for c in gl.columns]
rbi_sheet = pd.concat([lev, gl], axis=1).reset_index()
rbi_sheet["Date"] = rbi_sheet.Date.dt.date
dict_rows = [
    ("RBI_Monthly", "Levels: outstanding bank credit, Rs crore (RBI DBIE Table 15 'Deployment of Gross Bank Credit by Major Sectors'). Growth: 100 x monthly log change; July-2023 set to missing for housing, NBFC, HFC and non-food credit (HDFC merger).", "RBI DBIE"),
    ("NBFCs ex-HFC", "NBFCs (incl. HFCs) minus HFCs (row 3.9 minus 3.9.1)", "Computed"),
    ("COF_Quarterly", "Cost of funds/borrowing (%), per company-quarter, with source deck and page; 'COMPUTED' rows are annualised finance cost / average borrowings from the same deck", "Company investor presentations"),
    ("BankShare_Quarterly", "Bank share of borrowings (%), per company-quarter, as defined in each company's borrowing-mix slide", "Company investor presentations"),
    ("Bank_Exposure_Sep2023", "Bank share of borrowings at 30-Sep-2023 verified against each Q2 FY24 presentation (pre-policy exposure)", "Company investor presentations"),
    ("Bajaj_Segments", "Bajaj Finance consolidated AUM by business segment (Rs crore) per quarter", "Bajaj Finance investor presentations"),
    ("Company_Disclosures", "Verbatim statements quantifying the effect of the RBI risk-weight change", "Company investor presentations"),
    ("Results_*", "All regression results reported in the dissertation", "This study"),
]
xl = os.path.join(HERE, "Dissertation_Data.xlsx")
with pd.ExcelWriter(xl, engine="openpyxl") as w:
    pd.DataFrame(dict_rows, columns=["Sheet / variable", "Definition", "Source"]).to_excel(w, sheet_name="README", index=False)
    rbi_sheet.to_excel(w, sheet_name="RBI_Monthly", index=False)
    gt.reset_index().rename(columns={"index": "Series"}).to_excel(w, sheet_name="RBI_Growth_Table", index=False)
    cof[["Company", "FY", "Q", "COF_Pct", "Source_deck", "Note"]].to_excel(w, sheet_name="COF_Quarterly", index=False)
    bsh[["Company", "FY", "Q", "Bank_Share_Pct", "Source_deck", "Definition"]].to_excel(w, sheet_name="BankShare_Quarterly", index=False)
    ver.to_excel(w, sheet_name="Bank_Exposure_Sep2023", index=False)
    seg.to_excel(w, sheet_name="Bajaj_Segments", index=False)
    disc.to_excel(w, sheet_name="Company_Disclosures", index=False)
    sample.to_excel(w, sheet_name="Firm_Sample", index=False)
    pd.DataFrame([dict(Test="H1 interrupted time series (gap, Newey-West)", **{k: round(v, 4) for k, v in R["H1"]["its"].items()}),
                  dict(Test="H1 panel DiD (Driscoll-Kraay)", **{k: round(v, 4) for k, v in R["H1"]["panel"].items()}),
                  dict(Test="H1 placebo Nov-2022 (ITS)", **{k: round(v, 4) for k, v in R["H1"]["placebo"].items()}),
                  dict(Test="H2 NBFC ex-HFC minus non-food (ITS)", **{k: round(v, 4) for k, v in R["H2"]["its"].items()}),
                  dict(Test="H2 placebo Nov-2022", **{k: round(v, 4) for k, v in R["H2"]["placebo"].items()}),
                  dict(Test="H2 alt: NBFC ex-HFC minus HFC", **{k: round(v, 4) for k, v in R["H2"]["alt_vs_hfc"].items()}),
                  dict(Test="H2 alt: NBFC incl. HFC minus non-food", **{k: round(v, 4) for k, v in R["H2"]["alt_all_nbfc"].items()}),
                  dict(Test="H2 rollback Apr-2025", **{k: round(v, 4) for k, v in R["H2"]["rollback"].items()})]
                 ).to_excel(w, sheet_name="Results_Aggregate", index=False)
    pd.DataFrame(rob).to_excel(w, sheet_name="Results_H1_Robustness", index=False)
    pd.DataFrame([dict(Test="COF ~ Post x verified bank share (16 firms)", **R["firm_cof"]),
                  dict(Test="COF ~ Post x non-HFC", **R["firm_cof_hfc"]),
                  dict(Test="Bank share ~ Post x non-HFC (pre-rollback)", **R["firm_bankshare"]),
                  dict(Test="Bank share placebo Nov-2022", beta=R["firm_bankshare_placebo"]["beta"], p=R["firm_bankshare_placebo"]["p"]),
                  dict(Test="Bajaj within-firm segments (affected vs exempt)", **R["bajaj_within"])]).to_excel(w, sheet_name="Results_Firm", index=False)
    ev.to_excel(w, sheet_name="Results_COF_Event", index=False)
    evb.to_excel(w, sheet_name="Results_BankShare_Event", index=False)

json.dump(R, open(os.path.join(OUT, "results.json"), "w"), indent=1, default=float)
print(json.dumps({k: R[k] for k in ["H1", "H2", "firm_cof", "firm_bankshare", "firm_bankshare_rollback", "firm_bankshare_last", "firm_bankshare_placebo", "bajaj_within", "counts"]}, indent=1, default=lambda x: round(float(x), 4)))

# =========================================================== 4. Magnitudes and firm descriptives (appended)
end = lev.loc["2025-01"].iloc[-1]; m_post = 15          # Nov-2023 .. Jan-2025 monthly observations
tr_out = end[TREAT].sum(); nb_out = end["NBFCs ex-HFC"]
d1, d2 = -R["H1"]["its"]["change"], -R["H2"]["its"]["change"]
R["magnitudes"] = dict(
    h1_outstanding_jan25=float(tr_out), h1_shortfall=float(tr_out * (np.exp(d1 * m_post / 100) - 1)),
    h2_outstanding_jan25=float(nb_out), h2_shortfall=float(nb_out * (np.exp(d2 * m_post / 100) - 1)), months=m_post,
    cards_oct23=float(lev.loc["2023-10"].iloc[-1]["Credit cards"]), pl_oct23=float(lev.loc["2023-10"].iloc[-1]["Other personal loans"]),
    nbfc_oct23=float(lev.loc["2023-10"].iloc[-1]["NBFCs ex-HFC"]), nonfood_oct23=float(lev.loc["2023-10"].iloc[-1]["Non-food credit"]))
desc = []
for nm, d, col in [("Cost of funds (%)", cof, "COF_Pct"), ("Bank share of borrowings (%)", bsh, "Bank_Share_Pct")]:
    for grp_name, mask in [("Non-HFC NBFCs", d.NonHFC == 1), ("HFCs", d.NonHFC == 0)]:
        for per, pm in [("Pre (to Q2 FY24)", d.Post == 0), ("Post (Q3 FY24 on)", d.Post == 1)]:
            x = d.loc[mask & pm, col]
            desc.append(dict(Variable=nm, Group=grp_name, Period=per, N=int(x.size), Mean=round(x.mean(), 2), SD=round(x.std(), 2), Min=round(x.min(), 2), Max=round(x.max(), 2)))
R["firm_desc"] = desc
pd.DataFrame(desc).to_csv(os.path.join(OUT, "firm_descriptives.csv"), index=False)
json.dump(R, open(os.path.join(OUT, "results.json"), "w"), indent=1, default=float)
print("MAGNITUDES", {k: round(v, 1) for k, v in R["magnitudes"].items()})
print(pd.DataFrame(desc).to_string())
