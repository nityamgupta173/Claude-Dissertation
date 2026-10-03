"""Phases 2-3 (+ PDF Steps 2-7): Continuous DiD, TWFE (linearmodels.PanelOLS),
Wild Cluster Bootstrap-t inference, placebo, MFI reversal, event study.

Run:  python 01_build_master_panel.py && python 02_analysis.py
Outputs go to ./outputs
"""
import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats
from linearmodels.panel import PanelOLS

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "outputs")
os.makedirs(OUT, exist_ok=True)
B_REPS = 19999
SEED = 20231116
EXPO = "Pre_Policy_Bank_Exposure"
MACRO = ["Repo_Rate_Pct", "Bond_Yield_5Y_Pct", "GDP_Growth_YoY_Pct"]
FIRM_CTRL = ["lag_RoA_Pct", "lag_GNPA_Pct", "ln_Total_AUM"]

raw = pd.read_csv(os.path.join(HERE, "master_panel_data_final.csv"), parse_dates=["Date_End"])
raw["Firm"] = raw.Company.str.replace(r" (Ltd\.|Limited)$", "", regex=True)
for c in ["Unsecured", "MFI", "Secured", "Total"]:
    raw[f"g_{c}"] = 100 * raw[f"dlog_{c}"]            # 100 x log difference ~ QoQ growth, pp
raw["Post23_x_Exp"] = raw.Post_Nov2023 * raw[EXPO]
raw["Post22_x_Exp"] = raw.Post_Nov2022 * raw[EXPO]
raw["Post25_x_Exp"] = raw.Post_Feb2025 * raw[EXPO]
raw["Post23_x_ExpU"] = raw.Post_Nov2023 * raw.Exposure_Unsecured_Baseline * 100
raw["Post_x_Exp_x_T1"] = raw.Post23_x_Exp * (raw.lag_Tier1_CRAR - raw.lag_Tier1_CRAR.mean())
raw["Post_x_T1"] = raw.Post_Nov2023 * (raw.lag_Tier1_CRAR - raw.lag_Tier1_CRAR.mean())
raw["Exp_x_T1"] = raw[EXPO] * (raw.lag_Tier1_CRAR - raw.lag_Tier1_CRAR.mean())
raw["lag_Tier1_c"] = raw.lag_Tier1_CRAR - raw.lag_Tier1_CRAR.mean()
raw["Post22_x_ExpU"] = raw.Post_Nov2022 * raw.Exposure_Unsecured_Baseline * 100
raw["Post25_x_ExpU"] = raw.Post_Feb2025 * raw.Exposure_Unsecured_Baseline * 100

from did_tools import _basis, wcb_t, wcb_F, design, run, stars  # noqa: E402

ALL = {}
summary_rows = []


def keep(tag, m):
    ALL[tag] = m
    for j, c in enumerate(m["treat"]):
        r = m["res"][j]
        summary_rows.append(dict(model=tag, label=m["label"], dv=m["dv"], term=c, coef=r["beta"],
                                 se_cr1=r["se_cr1"], t=r["t"], p_wild_cluster_boot_t=r["p_wcb"],
                                 p_cr1_t_dist=r["p_cr1_t"], ci95_boot_lo=r["ci_lo"], ci95_boot_hi=r["ci_hi"],
                                 n_obs=m["n"], n_firms=m["firms"], adj_r2=m["adj_r2"],
                                 within_r2=m["within_r2"], absorbed_controls=";".join(m["absorbed"])))
    print(f"[{tag}] {m['label']}: N={m['n']} firms={m['firms']} absorbed={m['absorbed']}")
    for j, c in enumerate(m["treat"]):
        r = m["res"][j]
        print(f"    {c:18s} b={r['beta']:9.4f} seCR1={r['se_cr1']:8.4f} p_WCB={r['p_wcb']:.4f} p_t(G-1)={r['p_cr1_t']:.4f}")


D = raw
# ================= PHASE 2 / 3 : specifications requested in the update =================
keep("P2_core", run("Core: Unsecured growth x Post-Nov2023 x Bank exposure", D, "g_Unsecured", ["Post23_x_Exp"]))
keep("P3a_placebo_full", run("Placebo Nov-2022, full sample (exact Phase-2 model)", D, "g_Unsecured", ["Post22_x_Exp"]))
pre = D[D.Date_End <= "2023-09-30"]
keep("P3a_placebo_pre", run("Placebo Nov-2022, pre-policy sample only (<= Q2 FY24)", pre, "g_Unsecured", ["Post22_x_Exp"]))
keep("P3b_MFI", run("MFI growth x Post-Feb2025 x Bank exposure", D, "g_MFI", ["Post25_x_Exp"]))

# ================= PDF Models 1-5 (bank-share exposure, firm-level controls) =================
C = MACRO + FIRM_CTRL
keep("M1", run("M1 Unsecured growth", D, "g_Unsecured", ["Post23_x_Exp"], C))
D = D.assign(_t1=D.lag_Tier1_c)
keep("M3", run("M3 Unsecured growth, capital triple interaction", D, "g_Unsecured",
               ["Post23_x_Exp", "Post_x_Exp_x_T1"], C + ["lag_Tier1_c", "Post_x_T1", "Exp_x_T1"]))
keep("M4a", run("M4 Secured growth", D, "g_Secured", ["Post23_x_Exp"], C))
keep("M4b", run("M4b Total AUM growth (H4)", D, "g_Total", ["Post23_x_Exp"], C))
keep("M5_MFI", run("M5 MFI growth (Nov-23 + Feb-25)", D, "g_MFI", ["Post23_x_Exp", "Post25_x_Exp"], C))
keep("M5_nonMFI", run("M5' Non-MFI consumer (unsecured) growth (Nov-23 + Feb-25)", D, "g_Unsecured",
                      ["Post23_x_Exp", "Post25_x_Exp"], C))
# alternative exposure = pre-policy unsecured share (synopsis H1 wording; pp)
keep("A1_unsecShare", run("Alt exposure (unsecured share): Unsecured growth", D, "g_Unsecured", ["Post23_x_ExpU"], C))
keep("A4_unsecShare", run("Alt exposure (unsecured share): Secured growth", D, "g_Secured", ["Post23_x_ExpU"], C))
keep("A5_total_unsecShare", run("Alt exposure (unsecured share): Total growth", D, "g_Total", ["Post23_x_ExpU"], C))

# Model 2 (H2): borrowing cost. Quarterly interest expense is only available from Mar-2024 (Screener), so the
# DiD is run on the ANNUAL Screener panel (FY21-FY26): interest / average borrowings.
ann = pd.read_csv(os.path.join(HERE, "screener_annual_panel.csv"), parse_dates=["FY_End"])
orig = pd.read_csv(os.path.join(HERE, "nbfc_quarterly_data-v3.csv"))
expo_all = orig[(orig.Fiscal_Year == "FY24") & (orig.Quarter == "Q2")].set_index("Company").Bank_Borrowing_Share_Pct
ann["Date_End"] = ann.FY_End
ann["Exp"] = ann.Company.map(expo_all)
ann["PostFY24_x_Exp"] = (ann.FY >= 2024) * ann.Exp
ann["PostFY25_x_Exp"] = (ann.FY >= 2025) * ann.Exp
ann = ann.sort_values(["Company", "FY"])
ann["lag_RoA"] = ann.groupby("Company").RoA_annual_Pct.shift(1)
ann["ln_Borr_lag"] = np.log(ann.groupby("Company").Borrowings.shift(1))
ann = ann[ann.FY >= 2021]          # FY20 only used for lags
ann["Interest_Expense_Ratio_Pct"] = ann["Interest_Expense_Ratio_Pct"]
keep("M2_annual", run("M2 Interest-expense ratio, annual, FY24 = post", ann, "Interest_Expense_Ratio_Pct",
                      ["PostFY24_x_Exp"], ["lag_RoA", "ln_Borr_lag"], min_obs=4))
keep("M2_annual_noFY24", run("M2 Interest-expense ratio, annual, FY24 dropped, FY25+ = post",
                             ann[ann.FY != 2024], "Interest_Expense_Ratio_Pct", ["PostFY25_x_Exp"],
                             ["lag_RoA", "ln_Borr_lag"], min_obs=4))
m2_note = ("Model 2 (H2) estimated on the ANNUAL Screener panel (16 firms, FY21-FY26) because quarterly interest "
           "expense exists only from Mar-2024.")
print(m2_note)

# Unsecured-series quality: for 'assumed share' firms, unsecured growth == total growth by construction
qrows = []
for f, g in raw.groupby("Company"):
    x = g[["g_Unsecured", "g_Total"]].dropna()
    if len(x) > 5:
        qrows.append(dict(Company=f, flag=g.Unsecured_Series_Flag.iloc[0], n=len(x),
                          corr_unsecured_vs_total_growth=x.corr().iloc[0, 1],
                          max_abs_diff_pp=(x.g_Unsecured - x.g_Total).abs().max()))
pd.DataFrame(qrows).to_csv(os.path.join(OUT, "unsecured_series_quality.csv"), index=False)

# ================= Robustness =================
core = ALL["P2_core"]["sample"]
loo = []
for f in sorted(core.Company.unique()):
    r = run(f"drop {f}", D[D.Company != f], "g_Unsecured", ["Post23_x_Exp"], MACRO)
    loo.append(dict(dropped=f, coef=r["res"][0]["beta"], se_cr1=r["res"][0]["se_cr1"],
                    p_wcb=r["res"][0]["p_wcb"], n=r["n"], firms=r["firms"]))
loo = pd.DataFrame(loo)
loo.to_csv(os.path.join(OUT, "robustness_leave_one_firm_out.csv"), index=False)

out_mask = (D.g_Unsecured.abs() > 25)
outl = D[out_mask & D.g_Unsecured.notna()][["Company", "Fiscal_Year", "Quarter", "g_Unsecured"]]
outl.to_csv(os.path.join(OUT, "flagged_outlier_quarters.csv"), index=False)
r_out = run("Core, dropping |QoQ unsecured growth|>25pp quarters", D[~out_mask], "g_Unsecured", ["Post23_x_Exp"], MACRO)
keep("R1_core_no_outliers", r_out)
D2 = D.copy(); lo, hi = D2.g_Unsecured.quantile([.05, .95]); D2["g_Unsecured"] = D2.g_Unsecured.clip(lo, hi)
keep("R2_core_winsor", run("Core, DV winsorised 5/95", D2, "g_Unsecured", ["Post23_x_Exp"], MACRO))

# ================= Event study (PDF Step 3) =================
es_sample = D.dropna(subset=["g_Unsecured"] + FIRM_CTRL).copy()
es_sample = es_sample[es_sample.Company.isin(ALL["P2_core"]["sample"].Company.unique())]
quarters = sorted(es_sample.Date_End.unique())
ref = pd.Timestamp("2023-09-30")
es_cols = []
for q in quarters:
    if q == ref:
        continue
    col = f"ES_{pd.Timestamp(q):%Y%m}"
    es_sample[col] = (es_sample.Date_End == q) * es_sample[EXPO]
    es_cols.append((col, q))
cols = [c for c, _ in es_cols]
es = run("Event study", es_sample, "g_Unsecured", cols, C)
pre_cols = [c for c, q in es_cols if q < ref]
d = es["sample"]
ent = pd.get_dummies(d.Company).values.astype(float)
tim = pd.get_dummies(d.Date_End).values.astype(float)[:, 1:]
Wfe = np.column_stack([ent, tim] + [d[c].values for c in es["firm_ctrl"]])
ftest = wcb_F(d.g_Unsecured.values, d[pre_cols].values,
              np.column_stack([Wfe, d[[c for c in cols if c not in pre_cols]].values]), d.Company.values)
ftest_all_pre_n = len(pre_cols)
print("Event-study pre-trend F-test:", ftest)
est = pd.DataFrame([{**r, "quarter": f"{pd.Timestamp(q):%b-%Y}", "date": q,
                     "k": (pd.Timestamp(q).year - 2023) * 4 + (pd.Timestamp(q).month // 3) - 3}
                    for r, (c, q) in zip(es["res"], es_cols)])
tcrit = stats.t.ppf(.975, es["firms"] - 1)
est["lo"] = est.beta - tcrit * est.se_cr1
est["hi"] = est.beta + tcrit * est.se_cr1
est.to_csv(os.path.join(OUT, "event_study_coefficients.csv"), index=False)

fig, ax = plt.subplots(figsize=(9, 4.6))
INK, MUTED, BLUE, RED = "#1f2430", "#6b7280", "#2b6cb0", "#c2410c"
ax.axhline(0, color=MUTED, lw=1)
ax.axvline(ref + pd.Timedelta(days=45), color=RED, lw=1, ls="--")
ax.text(ref + pd.Timedelta(days=60), -3.6, "RBI risk-weight\nrise (Nov-2023)", color=RED, fontsize=8, va="bottom")
pl = pd.concat([est[["date", "beta", "lo", "hi"]], pd.DataFrame({"date": [ref], "beta": [0.], "lo": [0.], "hi": [0.]})]).sort_values("date")
ax.fill_between(pl.date, pl.lo, pl.hi, color=BLUE, alpha=.15, lw=0)
ax.plot(pl.date, pl.beta, color=BLUE, lw=2, marker="o", ms=4)
ax.scatter([ref], [0], color=INK, zorder=3, s=22)
ax.annotate("reference: Q2 FY24", (ref, 0), xytext=(-95, -22), textcoords="offset points", fontsize=8, color=INK)
ax.set_ylabel("Coefficient on Bank exposure x quarter\n(pp of QoQ growth per 1pp exposure)")
ax.set_title("Event study: unsecured AUM growth vs pre-policy bank-funding share", loc="left", fontsize=11, color=INK)
for s in ["top", "right"]:
    ax.spines[s].set_visible(False)
ax.grid(axis="y", color="#e5e7eb", lw=.6)
ax.text(0.01, 0.02, f"95% CI: cluster-robust SE, t({es['firms']-1}). Joint F-test of {ftest_all_pre_n} pre-period coefs = 0: "
        f"F={ftest['F']:.2f}, wild-bootstrap p={ftest['p_wcb']:.3f}", transform=ax.transAxes, fontsize=7.5, color=MUTED)
fig.tight_layout()
fig.savefig(os.path.join(OUT, "event_study.png"), dpi=200)
plt.close(fig)

# ================= Descriptives (PDF Step 2) =================
dv_list = [("Total_AUM_Cr", "Total AUM (Rs cr)"), ("Unsecured_Consumer_AUM_Cr", "Unsecured AUM (Rs cr)"),
           ("Secured_AUM_Cr", "Secured AUM (Rs cr)"), ("g_Unsecured", "Unsecured growth (100 x dlog)"),
           ("g_Secured", "Secured growth (100 x dlog)"), ("g_Total", "Total AUM growth (100 x dlog)"),
           ("g_MFI", "MFI growth (100 x dlog)"), ("Unsecured_AUM_Ratio", "Unsecured share of AUM"),
           ("Bank_Borrowing_Share_Pct", "Bank borrowing share (%)"), ("Tier_1_CRAR_Pct", "Tier-1 CRAR (%)"),
           ("Gross_NPA_GS3_Pct", "GNPA (%)"), ("RoA_Pct", "RoA (%)"), ("Repo_Rate_Pct", "Repo rate (%)"),
           ("Bond_Yield_5Y_Pct", "5Y G-sec yield (%)"), ("GDP_Growth_YoY_Pct", "GDP growth YoY (%)")]
rows = []
for c, nm in dv_list:
    for per, m in [("Pre-Nov 2023", raw.Date_End < "2023-12-31"), ("Post-Nov 2023", raw.Date_End >= "2023-12-31")]:
        s = raw.loc[m, c].dropna()
        rows.append(dict(Variable=nm, Period=per, N=len(s), Mean=s.mean(), Median=s.median(), SD=s.std(),
                         Min=s.min(), Max=s.max()))
desc = pd.DataFrame(rows)
desc.to_csv(os.path.join(OUT, "summary_statistics.csv"), index=False)
wide = desc.pivot(index="Variable", columns="Period", values=["N", "Mean", "Median", "SD", "Min", "Max"])
wide = wide.reindex([nm for _, nm in dv_list])
wide.round(2).to_csv(os.path.join(OUT, "summary_statistics_wide.csv"))

fig, axs = plt.subplots(1, 2, figsize=(10, 3.8))
firms = raw.drop_duplicates("Company")
est_firms = set(ALL["P2_core"]["sample"].Company)
axs[0].hist(firms[EXPO], bins=np.arange(20, 80, 5), color="#9db8d6", edgecolor="white")
axs[0].hist(firms[firms.Company.isin(est_firms)][EXPO], bins=np.arange(20, 80, 5), color=BLUE, edgecolor="white")
axs[0].set_xlabel("Bank borrowing share, Q2 FY24 (%)  [Pre_Policy_Bank_Exposure]")
axs[0].set_ylabel("Number of firms")
axs[0].set_title("Treatment intensity (dark = 8 firms in unsecured model)", loc="left", fontsize=9, color=INK)
axs[1].hist(firms.Exposure_Unsecured_Baseline * 100, bins=np.arange(0, 110, 10), color="#9db8d6", edgecolor="white")
axs[1].set_xlabel("Unsecured share of AUM, Q2 FY24 (%)")
axs[1].set_title("Alternative exposure (synopsis H1)", loc="left", fontsize=9, color=INK)
for a in axs:
    for s in ["top", "right"]:
        a.spines[s].set_visible(False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, "exposure_histogram.png"), dpi=200)
plt.close(fig)

# ================= Tables =================
pd.DataFrame(summary_rows).to_csv(os.path.join(OUT, "all_regression_results.csv"), index=False)

LABELS = {"Post23_x_Exp": "Post-Nov2023 x Bank exposure", "Post22_x_Exp": "Post-Nov2022 x Bank exposure",
          "Post25_x_Exp": "Post-Feb2025 x Bank exposure", "Post_x_Exp_x_T1": "Post x Exposure x Tier-1 (lag, centred)",
          "Post23_x_ExpU": "Post-Nov2023 x Unsecured share",
          "PostFY24_x_Exp": "Post (FY24+) x Bank exposure", "PostFY25_x_Exp": "Post (FY25+) x Bank exposure"}


def table(tags, names, title, fname, extra_note="", annual=False):
    terms = []
    for t in tags:
        for c in ALL[t]["treat"]:
            if c not in terms:
                terms.append(c)
    body = []
    for c in terms:
        r1, r2, r3 = [LABELS.get(c, c)], [""], [""]
        for t in tags:
            m = ALL[t]
            if c in m["treat"]:
                r = m["res"][m["treat"].index(c)]
                r1.append(f"{r['beta']:.3f}{stars(r['p_wcb'])}")
                r2.append(f"({r['se_cr1']:.3f})")
                r3.append(f"[{r['p_wcb']:.3f}]")
            else:
                r1.append(""); r2.append(""); r3.append("")
        body += [r1, r2, r3]
    stat = [["Firm FE"] + ["Yes"] * len(tags), ["Year FE" if annual else "Quarter FE"] + ["Yes"] * len(tags),
            ["Macro controls (repo, 5Y yield, GDP)"] + ["n/a (absorbed by time FE)"] * len(tags),
            ["Firm controls"] + ["Yes (lag RoA, ln lag borrowings)" if annual else ("Yes (lag RoA, lag GNPA, ln AUM)" if "lag_RoA_Pct" in ALL[t]["ctrl"] else "No") for t in tags],
            ["Observations"] + [str(ALL[t]["n"]) for t in tags], ["Firms (clusters)"] + [str(ALL[t]["firms"]) for t in tags],
            ["Adj. R-squared (LSDV)"] + [f"{ALL[t]['adj_r2']:.3f}" for t in tags],
            ["Within R-squared"] + [f"{ALL[t]['within_r2']:.3f}" for t in tags]]
    hdr = [""] + names
    note = ("Dependent variables are 100 x log difference of AUM (approx. QoQ growth in percentage points). Exposure is the "
            "Q2 FY24 (Sep-2023) bank-borrowing share in % points. Estimated with linearmodels PanelOLS (firm + quarter "
            "fixed effects). Parentheses: CR1 cluster-robust SE. Square brackets: wild cluster bootstrap-t p-value "
            f"(Webb weights, {B_REPS} draws, null imposed); stars from that p-value: * p<0.1, ** p<0.05, *** p<0.01. " + extra_note)
    # markdown
    md = [f"**{title}**", "", "| " + " | ".join(hdr) + " |", "|" + "---|" * len(hdr)]
    md += ["| " + " | ".join(r) + " |" for r in body + stat]
    md += ["", f"*{note}*"]
    open(os.path.join(OUT, fname + ".md"), "w").write("\n".join(md))
    pd.DataFrame(body + stat, columns=hdr).to_csv(os.path.join(OUT, fname + ".csv"), index=False)
    # latex
    esc = lambda s: s.replace("&", r"\&").replace("%", r"\%").replace("_", r"\_")
    tex = [r"\begin{table}[htbp]\centering\small", rf"\caption{{{esc(title)}}}",
           r"\begin{tabular}{l" + "c" * len(tags) + "}", r"\hline\hline",
           " & ".join(esc(h) for h in hdr) + r" \\", r"\hline"]
    for r in body:
        tex.append(" & ".join(esc(x).replace("***", "$^{***}$").replace("**", "$^{**}$").replace("*", "$^{*}$")
                              if i else esc(x) for i, x in enumerate(r)) + r" \\")
    tex.append(r"\hline")
    for r in stat:
        tex.append(" & ".join(esc(x) for x in r) + r" \\")
    tex += [r"\hline\hline", r"\end{tabular}", rf"\par\footnotesize {esc(note)}", r"\end{table}"]
    open(os.path.join(OUT, fname + ".tex"), "w").write("\n".join(tex))


table(["P2_core", "P3a_placebo_full", "P3a_placebo_pre", "P3b_MFI"],
      ["(1) Core: unsecured", "(2) Placebo Nov-22 (full)", "(3) Placebo Nov-22 (pre-policy)", "(4) MFI, Feb-25 reversal"],
      "Table 1. Core continuous DiD, placebo and MFI reversal tests (update specification)", "table1_core_placebo_mfi",
      "Models (1)-(3): 8 firms with non-zero unsecured AUM; model (4): 3 firms with MFI AUM only - bootstrap inference is not reliable with 3 clusters.")
table(["M1", "M3", "M4a", "M4b", "M5_MFI", "M5_nonMFI"],
      ["(M1) Unsecured", "(M3) Unsecured x capital", "(M4) Secured", "(M4b) Total", "(M5) MFI", "(M5') Non-MFI consumer"],
      "Table 2. Models 1-5 (bank-funding exposure, firm-level controls)", "table2_models1to5",
      "Model 2 is in Table 2b (annual data).")
table(["M2_annual", "M2_annual_noFY24"], ["(M2) FY24+ = post", "(M2b) FY24 dropped, FY25+ = post"],
      "Table 2b. Model 2: interest-expense ratio (annual Screener panel)", "table2b_model2_annual",
      "DV: interest expense / average borrowings (%), annual; coefficient in pp of cost per 1pp bank-funding share. Annual frequency, controls: lagged annual RoA and ln lagged borrowings.", annual=True)
table(["A1_unsecShare", "A4_unsecShare", "A5_total_unsecShare"],
      ["Unsecured growth", "Secured growth", "Total growth"],
      "Table 3. Alternative exposure: pre-policy unsecured share of AUM (pp)", "table3_alt_exposure_unsecured_share")

with open(os.path.join(OUT, "diagnostics.txt"), "w") as fh:
    fh.write(m2_note + "\n\n")
    fh.write(f"Event-study joint pre-trend test ({ftest_all_pre_n} coefficients, k<-1): F={ftest['F']:.3f}, "
             f"df=({ftest['q']},{ftest['df2']}), classical p={ftest['p_classical']:.4f}, wild-cluster-bootstrap p={ftest['p_wcb']:.4f}\n\n")
    fh.write("Firms in unsecured model: " + ", ".join(sorted(ALL['P2_core']['sample'].Firm.unique())) + "\n")
    fh.write("Firms in MFI model: " + ", ".join(sorted(ALL['P3b_MFI']['sample'].Firm.unique())) + "\n")
    fh.write("Firms in secured model: " + ", ".join(sorted(ALL['M4a']['sample'].Firm.unique())) + "\n\n")
    fh.write("Absorbed controls in core model: " + str(ALL['P2_core']['absorbed']) + "\n")
    fh.write("\nFull PanelOLS output (core):\n" + str(ALL["P2_core"]["fit"].summary) + "\n")
    fh.write("\nLeave-one-firm-out:\n" + loo.to_string() + "\n\nFlagged outlier quarters (|g|>25pp):\n" + outl.to_string() + "\n")
print("done")
