"""Additional figures and tables for the dissertation (timeline, diagrams, distributions, event-study detail).
Run after 05_final_analysis.py, 07_simple_tests.py and event_study/event_study.py.
Output: outputs/final/f*.png and outputs/final/extra_results.json
"""
import json
import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from matplotlib.patches import FancyBboxPatch

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "outputs", "final")
INK, MUTED, GRID, SURF = "#1f2430", "#6b7280", "#e5e7eb", "#ffffff"
C_T, C_C, C_P = "#c2410c", "#2b6cb0", "#7c3aed"          # targeted / exempt / consumer lenders
plt.rcParams.update({"font.size": 8.5, "axes.titlesize": 9.5, "axes.titleweight": "bold", "axes.edgecolor": MUTED,
                     "axes.labelcolor": INK, "xtick.color": INK, "ytick.color": INK, "font.family": "DejaVu Sans"})


def tidy(ax, grid="y"):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    if grid:
        ax.grid(axis=grid, color=GRID, lw=.6); ax.set_axisbelow(True)


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=200, bbox_inches="tight", facecolor=SURF); plt.close(fig)


# ------------------------------------------------------------------ data
src = open(os.path.join(HERE, "05_final_analysis.py")).read()
ns = {"__file__": os.path.join(HERE, "05_final_analysis.py")}
exec(src[:src.index("def its(")], ns)
g, lev, WIN, TREAT, CTRL = ns["g"], ns["lev"], ns["WIN"], ns["TREAT"], ns["CTRL"]
POST, PRE_END = "2023-11-30", "2023-11-17"
W = g[WIN[0]:WIN[1]]
GAP = {"H1": W[TREAT].mean(axis=1) - W[CTRL].mean(axis=1), "H2": W["NBFCs ex-HFC"] - W["Non-food credit"]}
S = json.load(open(os.path.join(OUT, "simple_results.json")))
E = json.load(open(os.path.join(HERE, "event_study", "results.json")))
ens = {"__file__": os.path.join(HERE, "event_study", "event_study.py")}
es = open(os.path.join(HERE, "event_study", "event_study.py")).read()
exec(es[:es.index("R = {\"meta\"")], ens)
RET, FIRMS, HFC, NONHFC, CONSUMER, NAMES = ens["RET"], ens["FIRMS"], ens["HFC"], ens["NONHFC"], ens["CONSUMER"], ens["NAMES"]
X = {}

# ------------------------------------------------------------------ F1 timeline of events and analysis windows
ev_days = {k: RET.index.get_loc(pd.Timestamp(v)) for k, v in [("E1", "2023-11-17"), ("E2", "2025-02-27")]}
est = {k: (RET.index[i - 250], RET.index[i - 31]) for k, i in ev_days.items()}
X["event_windows"] = {k: dict(est_start=str(a.date()), est_end=str(b.date())) for k, (a, b) in est.items()}
D = pd.Timestamp
dn = lambda d: mdates.date2num(D(d))
fig, ax = plt.subplots(figsize=(10, 6.2))
def bar(a, b, y, col, txt, tc="white", h=0.5, inside=True):
    ax.barh(y, dn(b) - dn(a), left=dn(a), height=h, color=col, lw=0)
    if inside: ax.text(dn(a) + (dn(b) - dn(a)) / 2, y, txt, ha="center", va="center", fontsize=7, color=tc)
    else: ax.text(dn(b) + 12, y, txt, ha="left", va="center", fontsize=7, color=INK)
def mark(d, y, txt, col, side="right", dy=0):
    ax.plot([dn(d)], [y], "o", color=col, ms=7, mec="white", mew=1.2, zorder=5)
    ax.text(dn(d) + (14 if side == "right" else -14), y + dy, txt, ha="left" if side == "right" else "right", va="center", fontsize=7, color=INK)
sections = [(11.4, "Policy and market events"), (8.6, "Firm-specific events"), (6.3, "Bank credit tests"), (3.9, "Event study")]
for y, t in sections:
    ax.text(dn("2020-10-01"), y + 0.75, t, fontsize=8.5, fontweight="bold", color=INK)
    ax.axhline(y + 1.25, color=GRID, lw=.8)
bar("2022-05-04", "2023-02-08", 11.4, "#9ca3af", "", tc=INK)
ax.text(dn("2022-05-04") - 14, 11.4, "Repo rate raised by 250 bps (May 2022 to Feb 2023)", ha="right", va="center", fontsize=7, color=INK)
mark("2025-02-25", 11.4, "25 Feb 2025: rollback announced", C_C, side="left")
mark("2023-11-16", 10.8, "16 Nov 2023: RBI circular", C_T, side="left")
mark("2025-04-01", 10.8, "1 Apr 2025: rollback effective", C_C)
mark("2024-02-29", 10.2, "29 Feb 2024: deadline for board-approved exposure limits", MUTED)
mark("2023-10-05", 8.6, "5 Oct 2023: Bajaj Finance board approves ₹10,000 cr raise", MUTED, side="left")
bar("2023-11-15", "2024-05-02", 8.6, "#d97706", "Bajaj Finance product embargo (15 Nov 2023 to 2 May 2024)", inside=False)
bar("2024-03-04", "2024-09-19", 7.9, "#a16207", "IIFL Finance gold-loan embargo (4 Mar to 19 Sep 2024)", inside=False)
bar("2021-04-01", PRE_END, 6.3, "#9ca3af", "Before: Apr 2021 to 17 Nov 2023 (32 months)", tc=INK)
bar("2023-12-29", "2025-01-31", 6.3, C_T, "After: 14 months")
bar("2021-07-01", PRE_END, 5.6, "#e5e7eb", "Placebo window (fake policy date Nov 2022)", tc=INK)
bar("2025-04-25", "2026-07-31", 5.6, C_C, "After rollback (H2)")
a1, b1 = est["E1"]; a2, b2 = est["E2"]
bar(str(a1.date()), str(b1.date()), 3.9, "#e5e7eb", "Estimation window", tc=INK)
mark("2023-11-17", 3.9, "Day 0: 17 Nov 2023", C_T)
bar(str(a2.date()), str(b2.date()), 3.2, "#e5e7eb", "Estimation window", tc=INK)
mark("2025-02-27", 3.2, "Day 0: 27 Feb 2025", C_C)
ax.axvline(dn("2023-11-16"), color=C_T, lw=.8, ls="--", alpha=.5); ax.axvline(dn("2025-02-25"), color=C_C, lw=.8, ls="--", alpha=.5)
ax.set_xlim(dn("2020-09-15"), dn("2026-10-31")); ax.set_ylim(2.6, 12.9); ax.set_yticks([])
ax.xaxis_date(); ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7])); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))
for s_ in ["top", "right", "left"]: ax.spines[s_].set_visible(False)
save(fig, "f_timeline.png")

# ------------------------------------------------------------------ F2 capital required per ₹100 of loans
fig, ax = plt.subplots(figsize=(8, 3.2))
cases = [("Bank: personal loan", 11.5, 100, 125), ("Bank: credit card", 11.5, 125, 150), ("Bank: loan to an NBFC\n(illustrative 50% weight)", 11.5, 50, 75),
         ("NBFC: personal loan", 15.0, 100, 125)]
yy = np.arange(len(cases))[::-1]
for yv, (lab, ratio, a, b) in zip(yy, cases):
    ca, cb = ratio * a / 100, ratio * b / 100
    ax.barh(yv + 0.18, ca, height=0.32, color="#9ca3af"); ax.barh(yv - 0.18, cb, height=0.32, color=C_T)
    ax.text(ca + 0.2, yv + 0.18, f"₹{ca:.2f} (risk weight {a}%)", va="center", fontsize=7.5)
    ax.text(cb + 0.2, yv - 0.18, f"₹{cb:.2f} (risk weight {b}%)", va="center", fontsize=7.5, color=C_T)
ax.set_yticks(yy); ax.set_yticklabels([c[0] for c in cases]); ax.set_xlim(0, 25)
ax.set_xlabel("Minimum capital required per ₹100 lent (₹)"); tidy(ax, "x")
ax.set_title("Capital needed per ₹100 of loans before (grey) and after (orange) the circular", loc="left")
save(fig, "f_capital_example.png")

# ------------------------------------------------------------------ F3 size of credit categories, Oct 2023
oct23 = lev.loc["2023-10"].iloc[-1] / 100000
cats = [("Housing", "Exempt (H1 control)"), ("Other personal loans", "Targeted (H1)"), ("NBFCs ex-HFC", "Targeted (H2)"), ("Vehicle loans", "Exempt (H1 control)"),
        ("HFCs", "Exempt (H2)"), ("Credit cards", "Targeted (H1)"), ("Education", "Exempt (H1 control)"), ("Gold loans", "Exempt")]
fig, ax = plt.subplots(figsize=(8, 3.4))
for i, (c, t) in enumerate(cats[::-1]):
    col = C_T if t.startswith("Targeted") else C_C
    ax.barh(i, oct23[c], color=col, height=0.62)
    ax.text(oct23[c] + 0.2, i, f"{oct23[c]:.1f}  ({t})", va="center", fontsize=7.5)
ax.set_yticks(range(len(cats))); ax.set_yticklabels([c for c, _ in cats[::-1]]); ax.set_xlim(0, 32)
ax.set_xlabel("Outstanding bank credit, ₹ lakh crore (October 2023)"); tidy(ax, "x")
ax.set_title("Size of the targeted and exempt credit categories just before the circular", loc="left")
save(fig, "f_credit_size.png")
X["oct23"] = {c: float(oct23[c]) for c, _ in cats}

# ------------------------------------------------------------------ F4 year-on-year growth by category (12-month sum of monthly log growth)
yoy = g.rolling(12, min_periods=11).sum()["2021-06":"2026-07"]
fig, axs = plt.subplots(1, 2, figsize=(10, 3.7), sharey=True)
for ax, series, ttl in [(axs[0], [("Credit cards", C_T, "-"), ("Other personal loans", C_T, "--"), ("Housing", C_C, "-"), ("Vehicle loans", C_C, "--"), ("Education", C_C, ":")], "Consumer categories (H1)"),
                        (axs[1], [("NBFCs ex-HFC", C_T, "-"), ("HFCs", C_C, "--"), ("Non-food credit", C_C, "-")], "Bank credit to NBFCs (H2)")]:
    for c, col, ls in series:
        ax.plot(yoy.index, yoy[c], color=col, ls=ls, lw=1.8, label=c)
    for d, col in [("2023-11-16", C_T), ("2025-02-25", C_C)]:
        ax.axvline(pd.Timestamp(d), color=col, lw=1, ls="--", alpha=.7)
    ax.axhline(0, color=MUTED, lw=.6); ax.set_title(ttl, loc="left"); tidy(ax); ax.legend(frameon=False, fontsize=7, loc="upper right")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
axs[0].set_ylabel("Growth over previous 12 months (log points × 100)")
save(fig, "f_yoy_growth.png")

# ------------------------------------------------------------------ F5 box plots of monthly gaps, before vs after
fig, axs = plt.subplots(1, 2, figsize=(9, 3.6))
for ax, k, ttl in [(axs[0], "H1", "H1: targeted minus exempt consumer credit"), (axs[1], "H2", "H2: NBFCs (ex-HFC) minus non-food credit")]:
    gp = GAP[k].dropna(); pre, post = gp[gp.index < POST], gp[gp.index >= POST]
    bp = ax.boxplot([pre, post], widths=0.5, patch_artist=True, medianprops=dict(color=INK, lw=1.5), whiskerprops=dict(color=MUTED), capprops=dict(color=MUTED), flierprops=dict(marker="o", ms=4, mfc=MUTED, mec="white"))
    for p_, col in zip(bp["boxes"], ["#d1d5db", "#f3b08a"]): p_.set_facecolor(col); p_.set_edgecolor(MUTED)
    rng = np.random.default_rng(1)
    for i, s_ in enumerate([pre, post], start=1):
        ax.scatter(i + rng.uniform(-0.12, 0.12, len(s_)), s_, s=10, color=INK, alpha=.5, zorder=3)
        ax.plot([i - 0.3, i + 0.3], [s_.mean()] * 2, color=C_T if i == 2 else C_C, lw=2)
    ax.set_xticks([1, 2]); ax.set_xticklabels([f"Before\n(n = {len(pre)})", f"After\n(n = {len(post)})"])
    ax.axhline(0, color=MUTED, lw=.6); ax.set_title(ttl, loc="left"); tidy(ax)
axs[0].set_ylabel("Monthly growth gap (pp)")
save(fig, "f_gap_boxplot.png")

# ------------------------------------------------------------------ F6 seasonal pattern: average gap by calendar month (all months Apr-21..Jan-25)
fig, axs = plt.subplots(1, 2, figsize=(9.5, 3.3), sharey=False)
mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
X["seasonal"] = {}
for ax, k, ttl in [(axs[0], "H1", "H1 gap"), (axs[1], "H2", "H2 gap")]:
    gp = GAP[k].dropna(); m = gp.groupby(gp.index.month).mean()
    X["seasonal"][k] = {mon[i - 1]: float(v) for i, v in m.items()}
    ax.bar([mon[i - 1] for i in m.index], m.values, color=C_C, width=0.7)
    ax.axhline(0, color=MUTED, lw=.6); ax.set_title(f"{ttl}: average by calendar month", loc="left"); tidy(ax)
    ax.tick_params(axis="x", labelsize=7)
axs[0].set_ylabel("Average monthly growth gap (pp)")
save(fig, "f_seasonal.png")

# ------------------------------------------------------------------ F7 residuals of the month-dummy regressions
fig, axs = plt.subplots(2, 2, figsize=(9.5, 5.2))
for row, k in enumerate(["H1", "H2"]):
    gp = GAP[k].dropna()
    Xm = pd.DataFrame({"Post": (gp.index >= POST).astype(float)}, index=gp.index)
    md = pd.get_dummies(gp.index.month, prefix="M", drop_first=True).astype(float); md.index = gp.index
    f = sm.OLS(gp, sm.add_constant(pd.concat([Xm, md], axis=1))).fit()
    ax = axs[row, 0]; ax.plot(f.resid.index, f.resid.values, color=INK, lw=1.2, marker="o", ms=3)
    ax.axhline(0, color=MUTED, lw=.6); ax.axvline(pd.Timestamp("2023-11-16"), color=C_T, lw=.8, ls="--")
    ax.set_title(f"Model {row + 1} ({k}): residuals over time", loc="left"); tidy(ax); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
    ax = axs[row, 1]; ax.hist(f.resid.values, bins=12, color=C_C, edgecolor="white")
    xs = np.linspace(f.resid.min(), f.resid.max(), 100); sd = f.resid.std()
    ax.plot(xs, len(f.resid) * (f.resid.max() - f.resid.min()) / 12 * np.exp(-xs ** 2 / (2 * sd ** 2)) / (sd * np.sqrt(2 * np.pi)), color=C_T, lw=1.5)
    ax.set_title(f"Model {row + 1}: distribution of residuals (normal curve)", loc="left"); tidy(ax)
save(fig, "f_residuals.png")

# ------------------------------------------------------------------ event study: market-model estimates, daily ARs, placebo distribution
mm = {}
for ek, i0 in ev_days.items():
    sl = slice(i0 - 250, i0 - 30)
    rows = {}
    for f_ in FIRMS:
        y_, x_ = RET[f_].iloc[sl], RET["NIFTY50"].iloc[sl]
        fit = sm.OLS(y_.values, sm.add_constant(x_.values)).fit()
        rows[f_] = dict(alpha_pct=100 * fit.params[0], beta=fit.params[1], r2=fit.rsquared, sd_pct=100 * np.sqrt(fit.mse_resid), n=int(fit.nobs))
    mm[ek] = rows
X["market_model"] = mm
desc = {}
for f_ in FIRMS + ["NIFTY50", "NIFTY500"]:
    r = RET[f_]
    desc[f_] = dict(mean_pct=100 * r.mean(), sd_pct=100 * r.std(), min_pct=100 * r.min(), max_pct=100 * r.max(), n=int(r.count()))
X["return_desc"] = desc

def port_ar(members, i0, diff_with=None):
    pr = RET[members].mean(axis=1)
    if diff_with is not None: pr = pr - RET[diff_with].mean(axis=1)
    m = RET["NIFTY50"]; sl = slice(i0 - 250, i0 - 30)
    fit = sm.OLS(pr.iloc[sl].values, sm.add_constant(m.iloc[sl].values)).fit()
    ar = pr - fit.params[0] - fit.params[1] * m
    return ar, ar.iloc[sl]

fig, axs = plt.subplots(1, 2, figsize=(10, 3.6), sharey=True)
for ax, (ek, i0), ttl in zip(axs, ev_days.items(), ["Circular (day 0 = 17 Nov 2023)", "Rollback (day 0 = 27 Feb 2025)"]):
    days = np.arange(-5, 6)
    a_n, _ = port_ar(NONHFC, i0); a_h, _ = port_ar(HFC, i0)
    ax.bar(days - 0.2, 100 * a_n.iloc[i0 - 5:i0 + 6].values, width=0.38, color=C_T, label="Non-HFC NBFCs")
    ax.bar(days + 0.2, 100 * a_h.iloc[i0 - 5:i0 + 6].values, width=0.38, color=C_C, label="HFCs (exempt)")
    ax.axhline(0, color=MUTED, lw=.6); ax.set_xticks(days); ax.set_xlabel("Trading day relative to day 0")
    ax.set_title(ttl, loc="left"); tidy(ax)
axs[0].set_ylabel("Daily abnormal return (%)"); axs[0].legend(frameon=False, fontsize=7.5)
save(fig, "f_daily_ar.png")

fig, axs = plt.subplots(1, 2, figsize=(10, 3.3))
for ax, (ek, i0), ttl in zip(axs, ev_days.items(), ["Circular", "Rollback"]):
    ar, est_ar = port_ar(NONHFC, i0, diff_with=HFC)
    ax.hist(100 * est_ar.values, bins=30, color="#9ca3af", edgecolor="white")
    v = 100 * ar.iloc[i0]; ax.axvline(v, color=C_T, lw=2)
    ax.annotate(f"Day 0: {v:+.2f}%", (v, ax.get_ylim()[1] * 0.85), xytext=(6 if v < 0 else -6, 0), textcoords="offset points",
                ha="left" if v < 0 else "right", fontsize=8, color=C_T)
    ax.set_xlabel("Daily abnormal return, NBFCs minus HFCs (%)"); ax.set_title(f"{ttl}: day 0 against 220 ordinary days", loc="left"); tidy(ax)
axs[0].set_ylabel("Number of days")
save(fig, "f_placebo_days.png")

fig, axs = plt.subplots(1, 2, figsize=(10, 4.4))
for ax, ek, ttl in [(axs[0], "E1", "Circular: day-0 abnormal return"), (axs[1], "E2", "Rollback: day-0 abnormal return")]:
    v = pd.Series({NAMES[f_]: 100 * E[ek]["NIFTY50"]["firms"][f_]["[0]"]["car"] for f_ in FIRMS}).sort_values()
    cols = [C_C if any(NAMES[h] == n for h in HFC) else (C_P if any(NAMES[c] == n for c in CONSUMER) else C_T) for n in v.index]
    ax.barh(v.index, v.values, color=cols, height=0.65); ax.axvline(0, color=MUTED, lw=.6)
    ax.set_title(ttl, loc="left"); tidy(ax, "x"); ax.tick_params(axis="y", labelsize=7); ax.set_xlabel("%")
from matplotlib.patches import Patch
axs[1].legend(handles=[Patch(color=C_T, label="Non-HFC NBFC"), Patch(color=C_P, label="Consumer lender (NBFC)"), Patch(color=C_C, label="HFC (exempt)")], frameon=False, fontsize=7, loc="lower right")
fig.tight_layout(); save(fig, "f_firm_ar.png")

# ------------------------------------------------------------------ firm level: before/after dumbbells
fig, axs = plt.subplots(1, 2, figsize=(10, 4.6))
for ax, nm, ttl, unit in [(axs[0], "cof", "Cost of funds (%)", "%"), (axs[1], "bank_share", "Bank share of borrowings (%)", "%")]:
    t = pd.DataFrame(S["firm"][nm]["table"]).sort_values(["group", "change"])
    t["label"] = t.company.str.replace(r" Ltd\.?| Limited", "", regex=True).str.replace("Investment & Finance Company", "Inv. & Fin.").str.replace(" Company India", "").str.replace("Financial Services", "Fin. Services")
    yv = np.arange(len(t))
    for i, r in zip(yv, t.itertuples()):
        col = C_C if r.group == "HFC" else C_T
        ax.plot([r.pre, r.post], [i, i], color=col, lw=1.8, alpha=.7)
        ax.plot(r.pre, i, "o", color="white", mec=col, mew=1.5, ms=5); ax.plot(r.post, i, "o", color=col, ms=5)
    ax.set_yticks(yv); ax.set_yticklabels(t.label, fontsize=7); ax.set_title(f"{ttl}: before (open) and after (filled)", loc="left"); tidy(ax, "x")
axs[1].legend(handles=[Patch(color=C_T, label="Non-HFC NBFC"), Patch(color=C_C, label="HFC (exempt)")], frameon=False, fontsize=7, loc="lower right")
fig.tight_layout(); save(fig, "f_firm_dumbbell.png")

# ------------------------------------------------------------------ Bajaj consumer share with embargo shading
seg = pd.read_csv(os.path.join(HERE, "hand_collected", "segments_bajaj.csv"))
seg["sh"] = 100 * seg.Affected_Consumer_AUM_Cr / seg.Total_AUM_Cr
seg["date"] = pd.PeriodIndex([f"20{int(fy[2:]) - (1 if q in ('Q1', 'Q2', 'Q3') else 0)}Q{ {'Q1': 2, 'Q2': 3, 'Q3': 4, 'Q4': 1}[q] }" for fy, q in zip(seg.FY, seg.Q)], freq="Q").to_timestamp(how="end")
fig, ax = plt.subplots(figsize=(8, 3.2))
ax.axvspan(pd.Timestamp("2023-11-15"), pd.Timestamp("2024-05-02"), color="#fde68a", alpha=.6, lw=0)
ax.text(pd.Timestamp("2024-05-20"), seg.sh.min() + 0.1, "Shaded: RBI embargo on eCOM and\nInsta EMI Card (15 Nov 2023 to 2 May 2024)", fontsize=7, ha="left", va="bottom")
ax.plot(seg.date, seg.sh, color=C_P, lw=2, marker="o", ms=4)
ax.axvline(pd.Timestamp("2023-11-16"), color=C_T, lw=.8, ls="--")
ax.set_ylabel("Affected consumer segments,\n% of AUM"); tidy(ax); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
ax.set_title("Bajaj Finance: share of consumer segments covered by the circular", loc="left")
save(fig, "f_bajaj_share.png")

# ------------------------------------------------------------------ diagrams: conceptual framework and research design
def box(ax, x, y, w, h, txt, fc, ec, fs=7.6, bold=False, tc=INK):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.01,rounding_size=0.02", fc=fc, ec=ec, lw=1.2))
    ax.text(x + w / 2, y + h / 2, txt, ha="center", va="center", fontsize=fs, fontweight="bold" if bold else "normal", color=tc, wrap=True)
def arrow(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle="-|>", color=MUTED, lw=1.2, shrinkA=2, shrinkB=2))

fig, ax = plt.subplots(figsize=(10, 5.4)); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
box(ax, 0.02, 0.42, 0.17, 0.16, "RBI circular\n16 Nov 2023\n(risk weights +\nexposure limits)", "#fff7ed", C_T, bold=True)
box(ax, 0.26, 0.66, 0.2, 0.15, "Consumer-credit risk weight\n100% → 125%\n(cards 125% → 150%)", "#fff7ed", C_T)
box(ax, 0.26, 0.19, 0.2, 0.15, "Bank-to-NBFC risk weight\n+25 percentage points\n(HFCs exempt)", "#fff7ed", C_T)
box(ax, 0.53, 0.66, 0.19, 0.15, "Banks and NBFCs need more\ncapital per ₹ of unsecured\nconsumer loans", "#f3f4f6", MUTED)
box(ax, 0.53, 0.19, 0.19, 0.15, "Banks need more capital\nper ₹ lent to NBFCs", "#f3f4f6", MUTED)
box(ax, 0.79, 0.74, 0.19, 0.17, "H1: bank credit to cards and\npersonal loans slows vs\nexempt household credit\n(Ch. 7)", "#eff6ff", C_C)
box(ax, 0.79, 0.53, 0.19, 0.15, "Consumer lenders'\nshares fall on the\nannouncement (Ch. 8)", "#eff6ff", C_C)
box(ax, 0.79, 0.30, 0.19, 0.17, "H2: bank credit to NBFCs\nslows vs total bank credit;\nrecovers after rollback\n(Ch. 7)", "#eff6ff", C_C)
box(ax, 0.79, 0.07, 0.19, 0.17, "NBFC shares fall on the\ncircular and rise on the\nrollback; HFC shares do not\n(Ch. 8)", "#eff6ff", C_C)
box(ax, 0.53, 0.42, 0.19, 0.16, "NBFC capital ratios,\nfunding costs and\nbank borrowing (Ch. 9)", "#f3f4f6", MUTED)
for a in [(0.19, 0.53, 0.26, 0.73), (0.19, 0.47, 0.26, 0.27)]: arrow(ax, *a)
arrow(ax, 0.46, 0.735, 0.53, 0.735); arrow(ax, 0.46, 0.265, 0.53, 0.265)
for (y1, y2) in [(0.735, 0.825), (0.735, 0.605)]: arrow(ax, 0.72, y1, 0.79, y2)
for (y1, y2) in [(0.265, 0.385), (0.265, 0.155)]: arrow(ax, 0.72, y1, 0.79, y2)
arrow(ax, 0.625, 0.66, 0.625, 0.58); arrow(ax, 0.625, 0.34, 0.625, 0.42)
ax.text(0.36, 0.92, "Policy change", ha="center", fontsize=9, fontweight="bold", color=INK)
ax.text(0.625, 0.92, "Mechanism", ha="center", fontsize=9, fontweight="bold", color=INK)
ax.text(0.885, 0.95, "What should be observed", ha="center", fontsize=9, fontweight="bold", color=INK)
save(fig, "f_framework.png")

fig, ax = plt.subplots(figsize=(10, 5.0)); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
cols_ = [(0.02, "Data"), (0.36, "Method"), (0.70, "Output")]
for x, t in cols_: ax.text(x + 0.14, 0.95, t, ha="center", fontsize=10, fontweight="bold", color=INK)
rows_ = [("RBI sectoral bank credit\nmonthly, Apr 2021 to Jul 2026\n(merger-adjusted)", "Growth gap: targeted minus exempt\nWelch t-test (before vs after)\nOLS with Post + month dummies", "H1 and H2 in bank credit\nTables 7.1 to 7.7"),
         ("Daily share prices, 16 NBFCs\nand HFCs + Nifty 50/500\nSep 2022 to Mar 2025", "Market model (days -250 to -31)\nAbnormal returns, CAR\nPortfolio t-test, sign test", "H1 and H2 in share prices\nTables 8.1 to 8.4"),
         ("382 investor presentations\n16 firms, FY2021 to FY2026", "Firm averages before vs after\nWelch and paired t-tests\nCompany statements", "Capital, funding costs,\nbank borrowing\nTables 9.1 and 9.2")]
for i, (a, b, c) in enumerate(rows_):
    y = 0.68 - i * 0.3
    box(ax, 0.02, y, 0.28, 0.2, a, "#f3f4f6", MUTED); box(ax, 0.36, y, 0.28, 0.2, b, "#fff7ed", C_T); box(ax, 0.70, y, 0.28, 0.2, c, "#eff6ff", C_C)
    arrow(ax, 0.30, y + 0.1, 0.36, y + 0.1); arrow(ax, 0.64, y + 0.1, 0.70, y + 0.1)
save(fig, "f_design.png")

json.dump(X, open(os.path.join(OUT, "extra_results.json"), "w"), indent=1, default=float)
print("figures done", X["event_windows"])
