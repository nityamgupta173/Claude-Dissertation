"""Tests used in the final dissertation (Option A + event study).

Bank credit data (RBI DBIE Table 15, merger-adjusted, prepared exactly as in 05_final_analysis.py):
  1. Welch two-sample t-test of the monthly growth gap, before vs after the circular.
  2. OLS regression of the gap on a post-policy dummy, with and without calendar-month dummies
     (classical standard errors; Newey-West standard errors reported as a robustness check).
  3. Placebo date (Nov-2022), pre-policy linear trend, H2 rollback comparison, H1 drop-one-series checks.
Firm data: change in each firm's average cost of funds and bank share (pre vs post), compared between
  non-HFC NBFCs and exempt HFCs with Welch t-tests, plus paired t-tests within each group.
Output: outputs/final/simple_results.json and a cost-of-funds chart.
"""
import json
import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson, jarque_bera
from statsmodels.stats.outliers_influence import variance_inflation_factor

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "outputs", "final")
src = open(os.path.join(HERE, "05_final_analysis.py")).read()
ns = {"__file__": os.path.join(HERE, "05_final_analysis.py")}
exec(src[:src.index("def its(")], ns)                       # data preparation only (section 1 of 05)
g, lev, WIN, TREAT, CTRL = ns["g"], ns["lev"], ns["WIN"], ns["TREAT"], ns["CTRL"]
POST, PRE_END = "2023-11-30", "2023-11-17"
W = g[WIN[0]:WIN[1]]
GAPS = {"H1": W[TREAT].mean(axis=1) - W[CTRL].mean(axis=1),
        "H2": W["NBFCs ex-HFC"] - W["Non-food credit"],
        "H2_core": W["NBFCs ex-HFC"] - W["Non-food ex NBFC, cards, other PL"]}
S = {}


def welch(gap, brk):
    gap = gap.dropna(); pre, post = gap[gap.index < brk], gap[gap.index >= brk]
    t, p = stats.ttest_ind(post, pre, equal_var=False)
    v1, v2 = post.var() / len(post), pre.var() / len(pre)
    df = (v1 + v2) ** 2 / (v1 ** 2 / (len(post) - 1) + v2 ** 2 / (len(pre) - 1))
    return dict(n_pre=len(pre), n_post=len(post), mean_pre=pre.mean(), mean_post=post.mean(), sd_pre=pre.std(), sd_post=post.std(),
                diff=post.mean() - pre.mean(), t=float(t), df=float(df), p=float(p))


def dummy_reg(gap, brk, months=True):
    gap = gap.dropna()
    X = pd.DataFrame({"Post": (gap.index >= brk).astype(float)}, index=gap.index)
    if months:
        md = pd.get_dummies(gap.index.month, prefix="M", drop_first=True).astype(float); md.index = gap.index
        X = pd.concat([X, md], axis=1)
    X = sm.add_constant(X)
    f = sm.OLS(gap, X).fit()
    nw = sm.OLS(gap, X).fit(cov_type="HAC", cov_kwds={"maxlags": 3}, use_t=True)
    jb = jarque_bera(f.resid); bp = het_breuschpagan(f.resid, X)
    vif = max(variance_inflation_factor(X.values, i) for i in range(1, X.shape[1])) if X.shape[1] > 2 else 1.0
    return dict(n=int(f.nobs), k=int(X.shape[1] - 1), const=f.params["const"], const_se=f.bse["const"], const_p=f.pvalues["const"],
                b=f.params["Post"], se=f.bse["Post"], t=f.tvalues["Post"], p=f.pvalues["Post"],
                se_nw=nw.bse["Post"], p_nw=nw.pvalues["Post"],
                r2=f.rsquared, adj_r2=f.rsquared_adj, F=f.fvalue, F_p=f.f_pvalue, ss_reg=f.ess, ss_res=f.ssr,
                df_reg=f.df_model, df_res=f.df_resid, dw=durbin_watson(f.resid), jb=jb[0], jb_p=jb[1], bp=bp[0], bp_p=bp[1], max_vif=vif,
                months_joint_p=(float(f.f_test(np.eye(X.shape[1])[2:]).pvalue) if months else None))


for k, gap in GAPS.items():
    S[k] = dict(welch=welch(gap, POST), reg_simple=dummy_reg(gap, POST, months=False), reg_months=dummy_reg(gap, POST, months=True))
# placebo: pre-policy data only (Jul-2021 to 17-Nov-2023), fake policy date Nov-2022
PB = g["2021-07":PRE_END]
for k, gap in [("H1", PB[TREAT].mean(axis=1) - PB[CTRL].mean(axis=1)), ("H2", PB["NBFCs ex-HFC"] - PB["Non-food credit"])]:
    S[k]["placebo_welch"] = welch(gap, "2022-11-30"); S[k]["placebo_reg_months"] = dummy_reg(gap, "2022-11-30", months=True)
    y = gap.dropna(); t = np.arange(len(y)) / 12.0
    f = sm.OLS(y.values, sm.add_constant(t)).fit()
    S[k]["pretrend"] = dict(slope_per_year=f.params[1], se=f.bse[1], p=f.pvalues[1], n=int(f.nobs))
# H2 after the April-2025 rollback: Dec-2023..Mar-2025 vs Apr-2025..Jul-2026
RB = g["2023-12":"2026-07"]
S["H2"]["rollback_welch"] = welch(RB["NBFCs ex-HFC"] - RB["Non-food credit"], "2025-04-30")
# H1 drop-one checks (Welch and month-dummy regression)
rob = []
for lbl, tr, ct in [("Credit cards only (drop other personal loans)", ["Credit cards"], CTRL),
                    ("Other personal loans only (drop credit cards)", ["Other personal loans"], CTRL),
                    ("Control without housing", TREAT, ["Vehicle loans", "Education"]),
                    ("Control without vehicle loans", TREAT, ["Housing", "Education"]),
                    ("Control without education", TREAT, ["Housing", "Vehicle loans"])]:
    gp = W[tr].mean(axis=1) - W[ct].mean(axis=1)
    a, b = welch(gp, POST), dummy_reg(gp, POST)
    rob.append(dict(spec=lbl, diff=a["diff"], p_welch=a["p"], b_months=b["b"], p_months=b["p"]))
S["H1"]["robustness"] = rob
# magnitudes (difference in mean gaps over the 14 post-policy months)
end = lev.loc["2025-01"].iloc[-1]
for k, base in [("H1", end[TREAT].sum()), ("H2", end["NBFCs ex-HFC"])]:
    d = S[k]["welch"]["diff"]; n = S[k]["welch"]["n_post"]
    S[k]["shortfall"] = dict(months=n, pct=100 * (np.exp(-d * n / 100) - 1), outstanding=float(base), amount=float(base * (np.exp(-d * n / 100) - 1)))

# ------------------------------------------------------------------ firm level: pre vs post firm averages
HC = os.path.join(HERE, "hand_collected")
HFC = ["AAVAS", "Can Fin", "Home First", "LIC Housing", "PNB Housing", "Repco"]
MFI = ["CreditAccess", "Arman"]
firm = {}
for nm, fn, col in [("cof", "cof_quarterly.csv", "COF_Pct"), ("bank_share", "bank_share_quarterly.csv", "Bank_Share_Pct")]:
    d = pd.read_csv(os.path.join(HC, fn))
    d["t"] = (d.FY.str[2:].astype(int) - 21) * 4 + d.Q.str[1].astype(int) - 1
    d = d[d.t < 20]                                            # up to Q4 FY2025 (before the April-2025 rollback)
    d["Post"] = (d.t >= 14).astype(int)                        # Q3 FY2024 onward
    m = d.groupby(["Company", "Post"])[col].mean().unstack().dropna()
    m["change"] = m[1] - m[0]
    m["group"] = np.where(m.index.str.contains("|".join(HFC)), "HFC", "Non-HFC NBFC")
    res = dict(table=[dict(company=c, group=r.group, pre=r[0], post=r[1], change=r.change) for c, r in m.iterrows()])
    for lbl, mm in [("all", m), ("excl_mfi", m[~m.index.str.contains("|".join(MFI))])]:
        a, b = mm[mm.group == "Non-HFC NBFC"].change, mm[mm.group == "HFC"].change
        t, p = stats.ttest_ind(a, b, equal_var=False)
        res[lbl] = dict(n_nonhfc=len(a), n_hfc=len(b), mean_change_nonhfc=a.mean(), mean_change_hfc=b.mean(), diff=a.mean() - b.mean(), t=float(t), p=float(p))
    for grp in ["Non-HFC NBFC", "HFC"]:
        mm = m[m.group == grp]
        t, p = stats.ttest_rel(mm[1], mm[0])
        res[f"paired_{grp}"] = dict(n=len(mm), mean_pre=mm[0].mean(), mean_post=mm[1].mean(), t=float(t), p=float(p))
    firm[nm] = res
    if nm == "cof":
        d["half"] = d.t // 2
        hy = d.groupby(["half", np.where(d.Company.str.contains("|".join(HFC)), "HFC", "NBFC")])[col].mean().unstack()
        fig, ax = plt.subplots(figsize=(8, 3.4))
        lab = [f"FY{21 + h // 2}H{h % 2 + 1}" for h in hy.index]
        ax.plot(lab, hy["NBFC"], color="#c2410c", lw=2, marker="o", ms=4, label="Non-HFC NBFCs")
        ax.plot(lab, hy["HFC"], color="#2b6cb0", lw=2, marker="o", ms=4, label="HFCs (exempt)")
        ax.axvline(5.5, color="#6b7280", ls="--", lw=1); ax.legend(frameon=False, fontsize=8)
        for s_ in ["top", "right"]: ax.spines[s_].set_visible(False)
        ax.grid(axis="y", color="#e5e7eb", lw=.6); ax.set_ylabel("Average cost of funds (%)"); ax.tick_params(axis="x", labelsize=7.5)
        ax.set_title("Average cost of funds by half-year, from company investor presentations", loc="left", fontsize=9)
        fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig4_cof_halfyear.png"), dpi=200); plt.close(fig)
        S["cof_halfyear"] = {l: dict(nbfc=float(a), hfc=float(b)) for l, a, b in zip(lab, hy["NBFC"], hy["HFC"])}
S["firm"] = firm
json.dump(S, open(os.path.join(OUT, "simple_results.json"), "w"), indent=1, default=float)

r = lambda x: round(float(x), 3)
for k in ["H1", "H2", "H2_core"]:
    print(k, "welch", {a: r(b) for a, b in S[k]["welch"].items()})
    for m_ in ["reg_simple", "reg_months"]:
        print("  ", m_, {a: r(b) for a, b in S[k][m_].items() if b is not None})
for k in ["H1", "H2"]:
    print(k, "placebo", r(S[k]["placebo_welch"]["diff"]), r(S[k]["placebo_welch"]["p"]), "reg", r(S[k]["placebo_reg_months"]["b"]), r(S[k]["placebo_reg_months"]["p"]),
          "pretrend", {a: r(b) for a, b in S[k]["pretrend"].items()}, "shortfall", {a: r(b) for a, b in S[k]["shortfall"].items()})
print("rollback", {a: r(b) for a, b in S["H2"]["rollback_welch"].items()})
for x in S["H1"]["robustness"]: print("  rob", x["spec"], r(x["diff"]), r(x["p_welch"]), r(x["b_months"]), r(x["p_months"]))
for nm in firm:
    print(nm, {k: {a: r(b) for a, b in v.items()} for k, v in firm[nm].items() if k != "table"})
