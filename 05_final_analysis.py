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
# H2 benchmark excluding the treated group and the H1-treated categories (audit B8)
lev["Non-food ex NBFC, cards, other PL"] = lev["Non-food credit"] - lev["NBFCs (incl. HFCs)"] - lev["Credit cards"] - lev["Other personal loans"]
lev.index.name = "Date"
# RBI publishes figures excluding the HDFC merger (the numbers in parentheses, unlabelled rows of the sheet) for
# 28-Jul-2023 to 27-Jun-2025. The merger is not only a level shift for housing, other personal loans and non-food
# credit: afterwards their growth includes growth of the former HDFC Ltd book. We therefore use the ex-merger series
# for growth in that period. Bank credit to NBFCs/HFCs has no ex-merger series: there the merger is a one-off
# level drop (bank loans to HDFC Ltd extinguished), so only the merger month is excluded.
EXR = {"Non-food credit": 10, "Housing": 14, "Other personal loans": 20}
exm = pd.DataFrame({k: pd.to_numeric(raw.iloc[r, 3:], errors="coerce").values for k, r in EXR.items()}, index=dates)
exm["Non-food ex NBFC, cards, other PL"] = exm["Non-food credit"] - lev["NBFCs (incl. HFCs)"] - lev["Credit cards"] - exm["Other personal loans"]
g = 100 * np.log(lev).diff()                       # monthly log growth, pp (reported series)
g_ex = 100 * np.log(exm).diff()
MERGER = pd.Timestamp("2023-07-28")                # HDFC Ltd -> HDFC Bank merger (first affected observation)
EX_END = exm["Housing"].last_valid_index()         # last ex-merger observation (27-Jun-2025)
for c in exm.columns:
    ok = (g.index > MERGER) & (g.index <= EX_END)
    g.loc[ok, c] = g_ex.loc[ok, c]
for c in ["Housing", "Other personal loans", "NBFCs (incl. HFCs)", "HFCs", "NBFCs ex-HFC", "Non-food credit", "Non-food ex NBFC, cards, other PL"]:
    g.loc[MERGER, c] = np.nan
WIN = ("2021-04-01", "2025-01-31")                 # main window: ends before the Feb-2025 announcement
POST = "2023-11-30"                                # first post-policy observation = 29-Dec-2023 (the Nov obs is 17-Nov-2023, one day after the circular)
LAGS = 3                                           # Newey-West lags: rule of thumb floor(4(T/100)^(2/9)) = 3 for T ~ 46 (audit C1)
TREAT = ["Credit cards", "Other personal loans"]
CTRL = ["Housing", "Vehicle loans", "Education"]


def its(gap, brk, lags=LAGS):
    gap = gap.dropna()
    X = sm.add_constant((gap.index >= brk).astype(float))
    r = sm.OLS(gap.values, X).fit(cov_type="HAC", cov_kwds={"maxlags": lags}, use_t=True)   # t(n-2) p-values (audit A2)
    return dict(pre_mean=r.params[0], change=r.params[1], se=r.bse[1], p=r.pvalues[1], n=int(len(gap)),
                n_pre=int((X[:, 1] == 0).sum()), n_post=int((X[:, 1] == 1).sum()), annualised_change=12 * r.params[1])


def panel(gg, tr, ct, brk):
    L = gg[tr + ct].stack().rename("g").reset_index()
    L.columns = ["date", "series", "g"]
    L["Treat_x_Post"] = (L.series.isin(tr) & (L.date >= brk)).astype(float)
    p = L.set_index(["series", "date"])
    r = PanelOLS(p.g, p[["Treat_x_Post"]], entity_effects=True, time_effects=True).fit(
        cov_type="kernel", kernel="bartlett", bandwidth=LAGS, debiased=True)
    return dict(beta=float(r.params.iloc[0]), se=float(r.std_errors.iloc[0]), p=float(r.pvalues.iloc[0]),
                n=int(r.nobs), annualised=12 * float(r.params.iloc[0]))


W = g[WIN[0]:WIN[1]]
gap1 = W[TREAT].mean(1) - W[CTRL].mean(1)
PRE_END = "2023-11-17"
PB = g["2021-07":PRE_END]; PA = g["2021-04":PRE_END]
R["H1"] = dict(its=its(gap1, POST), panel=panel(W, TREAT, CTRL, POST),
               placebo=its(PB[TREAT].mean(1) - PB[CTRL].mean(1), "2022-11-30"),
               placebo_panel=panel(PB, TREAT, CTRL, "2022-11-30"),
               placebo_from_apr21=its(PA[TREAT].mean(1) - PA[CTRL].mean(1), "2022-11-30"),
               placebo_panel_from_apr21=panel(PA, TREAT, CTRL, "2022-11-30"))
rob = []
for lbl, tr, ct, w in [("Drop credit cards", ["Other personal loans"], CTRL, WIN),
                       ("Drop other personal loans", ["Credit cards"], CTRL, WIN),
                       ("Drop housing (merger-affected)", TREAT, ["Vehicle loans", "Education"], WIN),
                       ("Drop vehicle loans", TREAT, ["Housing", "Education"], WIN),
                       ("Drop education", TREAT, ["Housing", "Vehicle loans"], WIN),
                       ("Shorter pre-period (from Apr-2022)", TREAT, CTRL, ("2022-04-01", WIN[1])),
                       ("Shorter post-period (to Oct-2024)", TREAT, CTRL, (WIN[0], "2024-10-31"))]:
    rr = panel(g[w[0]:w[1]], tr, ct, POST)
    rob.append(dict(Specification=lbl, Coefficient=round(rr["beta"], 2), SE=round(rr["se"], 2), p=round(rr["p"], 3), N=rr["n"]))
R["H1"]["robustness"] = rob
gold = panel(W, TREAT, CTRL + ["Gold loans"], POST)   # not a valid control (gold-price surge, agri->retail reclassification): note only
R["H1"]["gold_note"] = dict(beta=gold["beta"], p=gold["p"])
pd.DataFrame(rob).to_csv(os.path.join(OUT, "h1_robustness.csv"), index=False)

# H2 aggregate
gap2 = W["NBFCs ex-HFC"] - W["Non-food credit"]
R["H2"] = dict(its=its(gap2, POST),
               placebo=its(PB["NBFCs ex-HFC"] - PB["Non-food credit"], "2022-11-30"),
               placebo_from_apr21=its(PA["NBFCs ex-HFC"] - PA["Non-food credit"], "2022-11-30"),
               alt_core=its(W["NBFCs ex-HFC"] - W["Non-food ex NBFC, cards, other PL"], POST),
               alt_vs_hfc=its(W["NBFCs ex-HFC"] - W["HFCs"], POST),
               alt_all_nbfc=its(W["NBFCs (incl. HFCs)"] - W["Non-food credit"], POST),
               rollback=its((g["2023-12":]["NBFCs ex-HFC"] - g["2023-12":]["Non-food credit"]), "2025-04-30"))
# growth tables (annualised average monthly log growth)
PER = [("Pre-policy (Apr-21 to mid-Nov-23)", "2021-04-01", "2023-11-17"), ("Post-policy (Dec-23 to Mar-25)", "2023-12-01", "2025-03-31"),
       ("After rollback (Apr-25 to Jul-26)", "2025-04-01", "2026-07-31")]
gt = pd.DataFrame({p: (g[a:b].mean() * 12).round(1) for p, a, b in PER})
gt = gt.loc[TREAT + CTRL + ["Gold loans", "NBFCs ex-HFC", "HFCs", "NBFCs (incl. HFCs)", "Non-food credit", "Non-food ex NBFC, cards, other PL"]]
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
    pm = gap[WIN[0]:"2023-11-17"].mean(); ax.axhline(pm, color=C_C, lw=1, ls="-.")
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
    ("Results_*", "All regression results reported in the dissertation (Results_ModelSummary: R2, ANOVA, coefficients and diagnostics of Models 1-2; Results_Inference: p-values under alternative standard errors)", "This study"),
    ("RBI_Descriptives / RBI_Correlation", "Descriptive statistics and correlation matrix of monthly growth rates (Tables 4.4-4.5)", "This study"),
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
end = lev.loc["2025-01"].iloc[-1]; m_post = R["H1"]["its"]["n_post"]   # 14 post-policy observations: 29-Dec-2023 .. 31-Jan-2025
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

# =========================================================== 5. Template-required statistics (appended)
from statsmodels.stats.diagnostic import het_breuschpagan, acorr_ljungbox
from statsmodels.stats.stattools import jarque_bera, durbin_watson

SER = TREAT + CTRL + ["NBFCs ex-HFC", "HFCs", "Non-food credit"]
ds = []
for c in SER:
    for per, a, b in [("Pre-policy", "2021-04-01", "2023-11-17"), ("Post-policy", "2023-12-01", "2025-01-31")]:
        x = g.loc[a:b, c].dropna()
        ds.append(dict(Series=c, Period=per, N=int(x.size), Mean=round(x.mean(), 3), SD=round(x.std(), 3), Min=round(x.min(), 3), Max=round(x.max(), 3)))
R["rbi_desc"] = ds
corr = g.loc[WIN[0]:WIN[1], SER].corr().round(2)
R["rbi_corr"] = {"cols": SER, "rows": corr.values.tolist()}


def its_full(gap, brk, name, lags=LAGS):
    gap = gap.dropna()
    post = (gap.index >= brk).astype(float)
    X = sm.add_constant(post)
    ols = sm.OLS(gap.values, X).fit()
    hac = sm.OLS(gap.values, X).fit(cov_type="HAC", cov_kwds={"maxlags": lags}, use_t=True)
    jb, jbp, skew, kurt = jarque_bera(ols.resid)
    bp = het_breuschpagan(ols.resid, X)
    lb = acorr_ljungbox(ols.resid, lags=[12], return_df=True)
    return dict(name=name, n=int(ols.nobs), r2=ols.rsquared, adj_r2=ols.rsquared_adj, f=ols.fvalue, f_p=ols.f_pvalue,
                ss_model=ols.ess, ss_resid=ols.ssr, df_model=ols.df_model, df_resid=ols.df_resid,
                const=hac.params[0], const_se=hac.bse[0], const_p=hac.pvalues[0], delta=hac.params[1], delta_se=hac.bse[1], delta_p=hac.pvalues[1],
                jb=jb, jb_p=jbp, skew=skew, kurt=kurt, dw=durbin_watson(ols.resid), lb12=float(lb.lb_stat.iloc[0]), lb12_p=float(lb.lb_pvalue.iloc[0]),
                bp=bp[0], bp_p=bp[1], vif=1.0)


R["model_summary"] = [its_full(gap1, POST, "H1"), its_full(gap2, POST, "H2")]
# multicollinearity for the H1 panel design: VIF of Treated x Post given series and month fixed effects
L = W[TREAT + CTRL].stack().rename("g").reset_index(); L.columns = ["date", "series", "g"]
L["TxP"] = (L.series.isin(TREAT) & (L.date >= POST)).astype(float)
Z = pd.get_dummies(L[["series"]].astype(str), drop_first=True).join(pd.get_dummies(L.date.astype(str), drop_first=True)).astype(float)
aux = sm.OLS(L.TxP, sm.add_constant(Z)).fit()
R["vif_h1_panel"] = float(1 / (1 - aux.rsquared))
# appendix data: annual averages of hand-collected series
cof["FYn"] = cof.FY; bsh["FYn"] = bsh.FY
R["app_cof"] = cof.pivot_table(index="Company", columns="FYn", values="COF_Pct", aggfunc="mean").round(2).reset_index().fillna("").to_dict(orient="split")
R["app_bsh"] = bsh.pivot_table(index="Company", columns="FYn", values="Bank_Share_Pct", aggfunc="mean").round(1).reset_index().fillna("").to_dict(orient="split")
R["app_rbi"] = (lev.loc[lev.index.month == 3, SER] / 100000).round(2).reset_index().assign(Date=lambda x: x.Date.dt.strftime("%b-%Y")).to_dict(orient="split")
json.dump(R, open(os.path.join(OUT, "results.json"), "w"), indent=1, default=float)
for m_ in R["model_summary"]:
    print({k: (round(v, 3) if isinstance(v, float) else v) for k, v in m_.items()})
print("VIF H1 panel", round(R["vif_h1_panel"], 2))

# =========================================================== 6. Inference robustness (how sensitive is the p-value to the SE choice?)
def fixed_b_p(gap_y, X, j, B=10000, seed=11):
    """Kiefer-Vogelsang-Bunzel fixed-b test: Bartlett kernel, bandwidth = T (b = 1). The null distribution of the
    t-ratio is simulated with the actual design and iid N(0,1) errors (pivotal under fixed-b asymptotics)."""
    Xv = np.asarray(X, float); T = len(gap_y)
    XtXi = np.linalg.inv(Xv.T @ Xv)
    def tstat(y):
        b = XtXi @ Xv.T @ y; e = y - Xv @ b
        v = Xv * e[:, None]; S = v.T @ v
        for l in range(1, T):
            w = 1 - l / T; G = v[l:].T @ v[:-l]; S += w * (G + G.T)
        V = XtXi @ S @ XtXi
        return b[j] / np.sqrt(V[j, j])
    t0 = tstat(np.asarray(gap_y, float))
    rng = np.random.default_rng(seed)
    ts = np.array([tstat(rng.standard_normal(T)) for _ in range(B)])
    return float((np.abs(ts) >= abs(t0)).mean())


def infer_table(gap, brk, B=10000, seed=7):
    gap = gap.dropna()
    post = pd.Series((gap.index >= brk).astype(float), index=gap.index, name="post")
    md = pd.get_dummies(gap.index.month, prefix="m", drop_first=True).astype(float); md.index = gap.index
    X0, XS = sm.add_constant(post), sm.add_constant(pd.concat([post, md], axis=1))
    rows = []
    def add(lbl, X, kw):
        f = sm.OLS(gap, X).fit(**kw)
        rows.append(dict(Method=lbl, delta=round(float(f.params["post"]), 3), p=round(float(f.pvalues["post"]), 3)))
    add("Classical OLS", X0, {})
    add("White (HC1)", X0, dict(cov_type="HC1", use_t=True))
    for L_ in (1, 3, 6, 12):
        add(f"Newey-West, {L_} lag{'s' if L_ > 1 else ''}" + (" (main: rule of thumb)" if L_ == 3 else ""), X0, dict(cov_type="HAC", cov_kwds={"maxlags": L_}, use_t=True))
    rows.append(dict(Method="Fixed-b HAC (Kiefer-Vogelsang, Bartlett, b = 1)", delta=rows[0]["delta"], p=round(fixed_b_p(gap.values, X0.values, 1), 3)))
    add("Month dummies, classical OLS", XS, {})
    add("Month dummies, White (HC1)", XS, dict(cov_type="HC1", use_t=True))
    add("Month dummies, Newey-West 3 lags", XS, dict(cov_type="HAC", cov_kwds={"maxlags": 3}, use_t=True))
    rows.append(dict(Method="Month dummies, fixed-b HAC (b = 1)", delta=round(float(sm.OLS(gap, XS).fit().params["post"]), 3),
                     p=round(fixed_b_p(gap.values, XS.values, 1), 3)))
    # quarterly averages: calendar quarters; 2023Q4 (Oct, 17-Nov pre; Dec post) dropped as a transition quarter;
    # 2025Q1 has only January in the window and is dropped
    q = gap.groupby(gap.index.to_period("Q")).mean()
    q = q[(q.index != pd.Period("2023Q4")) & (q.index <= pd.Period("2024Q4"))]
    Xq = sm.add_constant(pd.Series((q.index >= pd.Period("2024Q1")).astype(float), index=q.index, name="post"))
    fq = sm.OLS(q.values, Xq.values).fit(cov_type="HC1", use_t=True)
    rows.append(dict(Method=f"Quarterly averages (n = {len(q)}), White (HC1)", delta=round(float(fq.params[1]), 3), p=round(float(fq.pvalues[1]), 3)))
    rng = np.random.default_rng(seed)
    for lbl, X in [("Moving-block bootstrap (block 3)", X0), ("Month dummies, moving-block bootstrap", XS)]:
        d = sm.OLS(gap, X).fit().params["post"]
        f0 = sm.OLS(gap, X.drop(columns="post")).fit(); fit0, r0 = f0.fittedvalues.values, f0.resid.values; n = len(r0)
        Xv, j = X.values, list(X.columns).index("post"); cnt = 0
        for _ in range(B):
            idx = np.concatenate([np.arange(s, s + 3) for s in rng.integers(0, n - 2, size=n // 3 + 1)])[:n]
            cnt += abs(np.linalg.lstsq(Xv, fit0 + r0[idx], rcond=None)[0][j]) >= abs(d)
        rows.append(dict(Method=lbl, delta=round(float(d), 3), p=round(cnt / B, 3)))
    return rows


R["inference"] = dict(H1=infer_table(gap1, POST), H2=infer_table(gap2, POST))
json.dump(R, open(os.path.join(OUT, "results.json"), "w"), indent=1, default=float)
print(pd.DataFrame(R["inference"]["H1"]).merge(pd.DataFrame(R["inference"]["H2"]), on="Method", suffixes=("_H1", "_H2")).to_string())
# add model statistics to the data workbook
with pd.ExcelWriter(xl, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
    pd.DataFrame(R["model_summary"]).to_excel(w, sheet_name="Results_ModelSummary", index=False)
    pd.DataFrame(R["inference"]["H1"]).merge(pd.DataFrame(R["inference"]["H2"]), on="Method", suffixes=("_H1", "_H2")).to_excel(w, sheet_name="Results_Inference", index=False)
    pd.DataFrame(R["rbi_desc"]).to_excel(w, sheet_name="RBI_Descriptives", index=False)
    pd.DataFrame(R["rbi_corr"]["rows"], index=SER, columns=SER).to_excel(w, sheet_name="RBI_Correlation")

# =========================================================== 7. Audit additions: pre-trends, event study, firm-level re-specifications
from scipy import stats as sstats
AUD = {}
# --- pre-trend tests on the pre-policy gap (linear time trend, NW 3 lags, t-based)
def pretrend(gap, start):
    y = gap[start:PRE_END].dropna()
    t = np.arange(len(y)) / 12.0                      # years
    f = sm.OLS(y.values, sm.add_constant(t)).fit(cov_type="HAC", cov_kwds={"maxlags": LAGS}, use_t=True)
    return dict(slope_per_year=float(f.params[1]), se=float(f.bse[1]), p=float(f.pvalues[1]), n=int(len(y)),
                first=str(y.index[0].date()), last=str(y.index[-1].date()))
G1 = g[TREAT].mean(1) - g[CTRL].mean(1); G2 = g["NBFCs ex-HFC"] - g["Non-food credit"]
AUD["pretrend"] = {k: dict(from_apr21=pretrend(G, "2021-04"), from_jul21=pretrend(G, "2021-07")) for k, G in [("H1", G1), ("H2", G2)]}
# half-year means of the gap, for the text
AUD["gap_quarter"] = {k: {str(p): round(float(v), 2) for p, v in G["2021-04":"2025-01"].groupby(G["2021-04":"2025-01"].index.to_period("Q")).mean().items()}
                      for k, G in [("H1", G1), ("H2", G2)]}
# counterfactual = pre-period linear trend extrapolated into the post period (mean reversion check, audit B1)
def trend_cf(gap, start):
    y = gap[start:WIN[1]].dropna(); t = np.arange(len(y)) / 12.0; post = y.index >= POST
    f = sm.OLS(y.values[~post], sm.add_constant(t[~post])).fit()
    dev = y.values[post] - (f.params[0] + f.params[1] * t[post])
    return dict(delta=float(dev.mean()), slope_per_year=float(f.params[1]), n_pre=int((~post).sum()), n_post=int(post.sum()))
AUD["trend_cf"] = {k: dict(from_apr21=trend_cf(G, "2021-04"), from_jul21=trend_cf(G, "2021-07")) for k, G in [("H1", gap1), ("H2", gap2)]}
# --- trend-adjusted ITS: gap = a + b*t + delta*Post (pre-trend assumed to continue; audit B1)
def its_trend(gap, brk):
    y = gap.dropna(); t = np.arange(len(y)) / 12.0; post = (y.index >= brk).astype(float)
    f = sm.OLS(y.values, sm.add_constant(np.column_stack([t, post]))).fit(cov_type="HAC", cov_kwds={"maxlags": LAGS}, use_t=True)
    return dict(delta=float(f.params[2]), se=float(f.bse[2]), p=float(f.pvalues[2]), trend=float(f.params[1]))
AUD["its_trend"] = dict(H1=its_trend(gap1, POST), H2=its_trend(gap2, POST))
# --- event study: 3-month event-time blocks counted from the policy; block -1 (Sep/Oct/17-Nov 2023) = reference
def event_study(gap):
    y = gap.dropna()
    k = np.array([(d.year - 2023) * 12 + d.month - 11 for d in y.index])   # months relative to Nov-2023 (Nov = 0)
    blk = np.where(k <= 0, -((-k) // 3) - 1, (k - 1) // 3)                 # Sep,Oct,Nov23 -> -1 ; Dec23-Feb24 -> 0
    D = pd.get_dummies(pd.Series(blk, index=y.index)).astype(float)
    D = D.drop(columns=-1)
    f = sm.OLS(y.values, sm.add_constant(D.values)).fit(cov_type="HAC", cov_kwds={"maxlags": LAGS}, use_t=True)
    cols = list(D.columns)
    out = pd.DataFrame(dict(block=cols, beta=f.params[1:], lo=f.conf_int()[1:, 0], hi=f.conf_int()[1:, 1], p=f.pvalues[1:]))
    out = pd.concat([out, pd.DataFrame([dict(block=-1, beta=0, lo=0, hi=0, p=np.nan)])]).sort_values("block")
    leads = [i + 1 for i, c in enumerate(cols) if c < -1]
    Rm = np.zeros((len(leads), len(cols) + 1)); Rm[np.arange(len(leads)), leads] = 1
    wt = f.f_test(Rm)
    fc = sm.OLS(y.values, sm.add_constant(D.values)).fit(); wc = fc.f_test(Rm)            # classical F
    lj = [i for i in leads if cols[i - 1] > -11]; Rj = np.zeros((len(lj), len(cols) + 1)); Rj[np.arange(len(lj)), lj] = 1
    wj = fc.f_test(Rj)                                                                     # classical F, excluding Apr-2021 block
    return out, dict(F=float(np.squeeze(wt.fvalue)), p=float(wt.pvalue), df=(int(wt.df_num), int(wt.df_denom)),
                     F_classical=float(np.squeeze(wc.fvalue)), p_classical=float(wc.pvalue),
                     F_classical_ex_apr21=float(np.squeeze(wj.fvalue)), p_classical_ex_apr21=float(wj.pvalue))
ev1, ft1 = event_study(gap1); ev2, ft2 = event_study(gap2)
AUD["event"] = dict(H1=dict(coef=ev1.round(3).to_dict(orient="records"), leads_F=ft1), H2=dict(coef=ev2.round(3).to_dict(orient="records"), leads_F=ft2))
fig, axs = plt.subplots(1, 2, figsize=(10, 3.6))
for ax, ev, ttl in [(axs[0], ev1, "H1: targeted minus exempt consumer credit"), (axs[1], ev2, "H2: NBFCs (ex-HFC) minus non-food credit")]:
    x = ev.block.values
    ax.axhline(0, color=MUTED, lw=.8); ax.axvline(-0.5, color=MUTED, ls="--", lw=1)
    ax.vlines(x, ev.lo, ev.hi, color=C_C, lw=1.6); ax.plot(x, ev.beta, "o", color=C_C, ms=4)
    ax.set_xlabel("3-month blocks relative to the policy (-1 = Sep to mid-Nov 2023, reference)", fontsize=7.5)
    ax.set_ylabel("Gap relative to reference, pp/month", fontsize=8); ax.set_title(ttl, loc="left", fontsize=9, color=INK); tidy(ax)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig6_event_study.png"), dpi=200); plt.close(fig)

# --- firm level: Model 3 as triple difference (HFCs exempt), exemption-based COF test (B12), Model 4 subsamples (B5, B9)
cof["Post_x_NonHFC"] = cof.Post * cof.NonHFC
cof["Post_x_Exp_x_NonHFC"] = cof.Post * cof.Exp * cof.NonHFC
m = run("", cof, "COF_Pct", ["Post_x_Exp", "Post_x_NonHFC", "Post_x_Exp_x_NonHFC"], [], min_obs=8)
AUD["cof_triple"] = dict(beta=m["res"][2]["beta"], se=m["res"][2]["se_cr1"], p=m["res"][2]["p_wcb"], n=m["n"], firms=m["firms"],
                         beta_exp=m["res"][0]["beta"], p_exp=m["res"][0]["p_wcb"])
MFI = ["CreditAccess", "Arman"]
def bs_model(d, label):
    m = run("", d[d.t < 20], "Bank_Share_Pct", ["NxPost"], [], min_obs=4)
    return dict(label=label, beta=m["res"][0]["beta"], p=m["res"][0]["p_wcb"], n=m["n"], firms=m["firms"])
AUD["bankshare_subsamples"] = [
    bs_model(bsh, "All firms (main)"),
    bs_model(bsh[~bsh.Company.str.contains("IIFL")], "Excluding IIFL Finance (gold-loan embargo Mar-Sep 2024)"),
    bs_model(bsh[~bsh.Company.str.contains("|".join(MFI))], "Excluding NBFC-MFIs (CreditAccess, Arman: bank loans largely priority-sector, exempt)"),
    bs_model(bsh[~bsh.Company.str.contains("|".join(MFI + ["IIFL"]))], "Excluding IIFL and NBFC-MFIs"),
    bs_model(bsh[~bsh.Company.str.contains("|".join(MFI + ["IIFL", "Muthoot"]))], "Excluding IIFL, NBFC-MFIs and Muthoot"),
    bs_model(bsh[~bsh.Company.str.contains("SBI Cards")], "Excluding SBI Cards"),
]
# raw firm-level changes in bank share (pre vs post-to-Mar-2025 means) for the narrative
chg = bsh[bsh.t < 20].groupby(["Company", "Post"]).Bank_Share_Pct.mean().unstack()
AUD["bankshare_firm_change"] = {c: dict(pre=round(r[0], 1), post=round(r[1], 1)) for c, r in chg.iterrows()}
# balanced panel: firms observed in every half-year FY22H1 .. FY25H2 (h = 2..9)
hh = bsh[(bsh.h >= 2) & (bsh.h <= 9)]
full = hh.groupby("Company").h.nunique(); bal = full[full == 8].index
grp_bal = bsh[bsh.Company.isin(bal)].groupby([bsh.t // 2, "NonHFC"]).Bank_Share_Pct.mean().unstack()
AUD["balanced_firms"] = list(bal)
AUD["bankshare_halfyear_raw"] = {f"FY{21 + h // 2}H{h % 2 + 1}": dict(nonhfc=round(r[1], 1), hfc=round(r[0], 1)) for h, r in grp.iterrows()}
AUD["bankshare_halfyear_bal"] = {f"FY{21 + h // 2}H{h % 2 + 1}": dict(nonhfc=round(r[1], 1), hfc=round(r[0], 1)) for h, r in grp_bal.iterrows()}
fig, ax = plt.subplots(figsize=(8, 3.6))
lab = [f"FY{21 + h // 2}H{h % 2 + 1}" for h in grp.index]
ax.plot(lab, grp[1], color=C_T, lw=2, marker="o", ms=4, label="Non-HFC NBFCs, all firms")
ax.plot(lab, grp[0], color=C_C, lw=2, marker="o", ms=4, label="HFCs (exempt), all firms")
lb = [f"FY{21 + h // 2}H{h % 2 + 1}" for h in grp_bal.index]
ax.plot(lb, grp_bal[1], color=C_T, lw=1.4, ls="--", label=f"Non-HFC NBFCs, balanced panel ({int((bsh[bsh.Company.isin(bal)].groupby('Company').NonHFC.first() == 1).sum())} firms)")
ax.plot(lb, grp_bal[0], color=C_C, lw=1.4, ls="--", label=f"HFCs, balanced panel ({int((bsh[bsh.Company.isin(bal)].groupby('Company').NonHFC.first() == 0).sum())} firms)")
ax.axvline(5.5, color=MUTED, ls="--", lw=1); ax.legend(frameon=False, fontsize=7); tidy(ax)
ax.set_ylabel("Bank share of borrowings (%)"); ax.tick_params(axis="x", labelsize=7.5)
ax.set_title("Average bank share of borrowings (from investor presentations)", loc="left", fontsize=9, color=INK)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig5_bank_share.png"), dpi=200); plt.close(fig)
# --- Bajaj: segment classification table, and Model 5 excluding the eCOM/Insta EMI Card embargo quarters (Q3 FY24 - Q1 FY25)
AUD["bajaj_segments"] = sorted(s.Seg.unique().tolist())
s2 = s[~s.t.isin([14, 15, 16])].copy()
p2 = s2.set_index(["Seg", "t"])
r2 = PanelOLS(p2.g, p2[["TxP"]], entity_effects=True, time_effects=True).fit(cov_type="kernel", kernel="bartlett", bandwidth=4)
AUD["bajaj_ex_embargo"] = dict(beta=float(r2.params.iloc[0]), p=float(r2.pvalues.iloc[0]), n=int(r2.nobs))
AUD["bajaj_share_series"] = R["bajaj_share"]
R["audit"] = AUD
json.dump(R, open(os.path.join(OUT, "results.json"), "w"), indent=1, default=float)
with pd.ExcelWriter(xl, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
    pd.DataFrame(AUD["bankshare_subsamples"]).to_excel(w, sheet_name="Results_BankShare_Subsamples", index=False)
    ev1.to_excel(w, sheet_name="Results_EventStudy_H1", index=False); ev2.to_excel(w, sheet_name="Results_EventStudy_H2", index=False)
print(json.dumps({k: v for k, v in AUD.items() if k not in ("event", "bankshare_firm_change", "bajaj_share_series")}, indent=1, default=lambda x: round(float(x), 4)))
print("EVENT H1", ev1.round(2).to_string()); print("EVENT H2", ev2.round(2).to_string()); print(ft1, ft2)
