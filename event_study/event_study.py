"""Stock-market event study: NBFC/HFC share-price reaction to the RBI's risk-weight circular (16-Nov-2023, released
after market close -> day 0 = 17-Nov-2023) and its partial rollback (25-Feb-2025, released after close; 26-Feb-2025
was a market holiday -> day 0 = 27-Feb-2025).

Method (Brown & Warner, 1985; MacKinlay, 1997): market model on daily simple returns, estimation window trading days
[-250, -31], benchmark Nifty 50 (Nifty 500 as robustness). Because every firm shares the same event date, inference
uses equal-weighted portfolios (abnormal returns of the portfolio, variance from its own estimation-window residuals),
plus a sign test across firms and a cross-sectional regression of firm CARs on pre-policy bank dependence.
Run: python event_study/event_study.py  -> event_study/results.json, event_study/*.csv, event_study/fig_*.png
"""
import glob
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# ------------------------------------------------------------------ data
P = {}
for f in sorted(glob.glob(os.path.join(HERE, "raw", "*.csv"))):
    d = pd.read_csv(f, encoding="utf-8-sig")
    d["Date"] = pd.to_datetime(d.Date, format="%d-%m-%Y")
    d["Price"] = pd.to_numeric(d.Price.astype(str).str.replace(",", ""))
    P[os.path.basename(f)[:-4]] = d.set_index("Date").Price.sort_index()
P = pd.DataFrame(P)
RET = P.pct_change().iloc[1:]
MKT = ["NIFTY50", "NIFTY500"]
FIRMS = [c for c in RET.columns if c not in MKT]
HFC = ["AAVAS", "CANFINHOME", "HOMEFIRST", "LICHSGFIN", "PNBHOUSING", "REPCOHOME"]
NONHFC = [f for f in FIRMS if f not in HFC]
CONSUMER = ["BAJFINANCE", "SBICARD", "POONAWALLA"]          # unsecured consumer lending is a core business
MFI = ["CREDITACC", "ARMANFIN"]                              # bank borrowing largely priority-sector (exempt)
NAMES = {"AAVAS": "AAVAS Financiers", "ARMANFIN": "Arman Financial", "BAJFINANCE": "Bajaj Finance", "CANFINHOME": "Can Fin Homes",
         "CHOLAFIN": "Cholamandalam Inv. & Fin.", "CREDITACC": "CreditAccess Grameen", "HOMEFIRST": "Home First Finance",
         "IIFL": "IIFL Finance", "LICHSGFIN": "LIC Housing Finance", "MMFIN": "M&M Financial Services", "MUTHOOTFIN": "Muthoot Finance",
         "PNBHOUSING": "PNB Housing Finance", "POONAWALLA": "Poonawalla Fincorp", "REPCOHOME": "Repco Home Finance",
         "SBICARD": "SBI Cards & Payment", "SHRIRAMFIN": "Shriram Finance"}
# verified bank share of borrowings at 30-Sep-2023 (Table 4.2); Poonawalla does not disclose its borrowing mix
BANKSHARE = {"AAVAS": 49.6, "ARMANFIN": 32.3, "BAJFINANCE": 23.0, "CANFINHOME": 57.0, "CHOLAFIN": 60.0, "CREDITACC": 52.9,
             "HOMEFIRST": 54.0, "IIFL": 55.3, "LICHSGFIN": 33.0, "MMFIN": 57.2, "MUTHOOTFIN": 65.4, "PNBHOUSING": 40.2,
             "REPCOHOME": 75.5, "SBICARD": 77.0, "SHRIRAMFIN": 25.6}
EVENTS = {"E1": dict(label="RBI circular (16-Nov-2023)", day0="2023-11-17", sign=-1),
          "E2": dict(label="Rollback of bank-to-NBFC risk weight (25-Feb-2025)", day0="2025-02-27", sign=+1)}
EST = (-250, -31)
WINDOWS = {"[0]": (0, 0), "[0,+1]": (0, 1), "[0,+2]": (0, 2), "[-1,+1]": (-1, 1)}
PLOT = (-10, 10)


def event_index(day0):
    i = RET.index.get_loc(pd.Timestamp(day0))
    return i


def market_model(y, m, i0):
    """Fit y = a + b*m on [i0+EST[0], i0+EST[1]]; return params, residual sd, n."""
    sl = slice(i0 + EST[0], i0 + EST[1] + 1)
    yy, mm = y.iloc[sl], m.iloc[sl]
    f = sm.OLS(yy.values, sm.add_constant(mm.values)).fit()
    return f.params, float(np.sqrt(f.mse_resid)), int(f.nobs)


def abnormal(series, mkt, i0):
    a, sd, n = market_model(series, mkt, i0)
    idx = range(i0 + PLOT[0], i0 + PLOT[1] + 1)
    ar = pd.Series([series.iloc[k] - a[0] - a[1] * mkt.iloc[k] for k in idx], index=range(PLOT[0], PLOT[1] + 1))
    return ar, sd, n


def car_test(ar, sd, n, w):
    L = w[1] - w[0] + 1
    car = float(ar.loc[w[0]:w[1]].sum())
    t = car / (sd * np.sqrt(L))
    return dict(car=car, t=t, p=float(2 * stats.t.sf(abs(t), n - 2)))


R = {"meta": dict(estimation_window=EST, windows=list(WINDOWS), firms=len(FIRMS), sample=[str(RET.index.min().date()), str(RET.index.max().date())])}
figdata = {}
for ek, ev in EVENTS.items():
    i0 = event_index(ev["day0"])
    R[ek] = dict(label=ev["label"], day0=ev["day0"], dates={str(k): str(RET.index[i0 + k].date()) for k in (-1, 0, 1, 2)})
    for mk in MKT:
        mkt = RET[mk]
        # firm-level ARs / CARs
        firm = {}
        for f in FIRMS:
            ar, sd, n = abnormal(RET[f], mkt, i0)
            firm[f] = dict(ar=ar, sd=sd, n=n, **{w: car_test(ar, sd, n, WINDOWS[w]) for w in WINDOWS})
        # portfolios (equal weighted)
        ports = {"All 16 firms": FIRMS, "Non-HFC NBFCs (10)": NONHFC, "HFCs (6, exempt)": HFC,
                 "Consumer lenders (Bajaj, SBI Cards, Poonawalla)": CONSUMER,
                 "Non-HFC NBFCs excl. MFIs (8)": [f for f in NONHFC if f not in MFI],
                 "Non-HFC NBFCs excl. Bajaj (9)": [f for f in NONHFC if f != "BAJFINANCE"]}
        pres = {}
        pr_series = {}
        for pn, members in ports.items():
            pr = RET[members].mean(axis=1); pr_series[pn] = pr
            ar, sd, n = abnormal(pr, mkt, i0)
            pres[pn] = dict(members=members, **{w: car_test(ar, sd, n, WINDOWS[w]) for w in WINDOWS})
            if mk == "NIFTY50": figdata[(ek, pn)] = ar
        # difference portfolios (long treated, short exempt HFCs): market model on the return difference
        for dn, a_, b_ in [("Non-HFC NBFCs minus HFCs", "Non-HFC NBFCs (10)", "HFCs (6, exempt)"),
                           ("Consumer lenders minus HFCs", "Consumer lenders (Bajaj, SBI Cards, Poonawalla)", "HFCs (6, exempt)")]:
            ar, sd, n = abnormal(pr_series[a_] - pr_series[b_], mkt, i0)
            pres[dn] = {w: car_test(ar, sd, n, WINDOWS[w]) for w in WINDOWS}
        # sign test (binomial) and cross-sectional regression on bank share
        cs = {}
        for w in WINDOWS:
            cars = pd.Series({f: firm[f][w]["car"] for f in FIRMS})
            nonh = cars[NONHFC]
            k = int((np.sign(nonh) == ev["sign"]).sum())
            sign_p = float(stats.binomtest(k, len(nonh), 0.5).pvalue)
            X = pd.DataFrame({"bank_share": pd.Series(BANKSHARE), "nonhfc": pd.Series({f: int(f in NONHFC) for f in FIRMS})}).dropna()
            X["bank_x_nonhfc"] = X.bank_share * X.nonhfc
            y = cars[X.index]
            fit = sm.OLS(y.values, sm.add_constant(X[["bank_share", "nonhfc", "bank_x_nonhfc"]].values)).fit(cov_type="HC1", use_t=True)
            fit2 = sm.OLS(cars.values, sm.add_constant(pd.Series({f: int(f in NONHFC) for f in FIRMS})[cars.index].values)).fit(cov_type="HC1", use_t=True)
            cs[w] = dict(sign_nonhfc=dict(k=k, n=len(nonh), p=sign_p),
                         reg_bankshare=dict(n=int(fit.nobs), b_bank=float(fit.params[1]), p_bank=float(fit.pvalues[1]),
                                            b_nonhfc=float(fit.params[2]), p_nonhfc=float(fit.pvalues[2]),
                                            b_inter=float(fit.params[3]), p_inter=float(fit.pvalues[3])),
                         reg_nonhfc_dummy=dict(b=float(fit2.params[1]), p=float(fit2.pvalues[1])))
        R[ek][mk] = dict(firms={f: {**{w: firm[f][w] for w in WINDOWS}, "ar0": float(firm[f]["ar"].loc[0]),
                                    "ar_m1": float(firm[f]["ar"].loc[-1])} for f in FIRMS},
                         portfolios=pres, cross_section=cs,
                         market_return_day0=float(mkt.iloc[i0]))

json.dump(R, open(os.path.join(HERE, "results.json"), "w"), indent=1)

# ------------------------------------------------------------------ tables (CSV)
rows = []
for ek in EVENTS:
    for f in FIRMS:
        d = R[ek]["NIFTY50"]["firms"][f]
        rows.append(dict(Event=ek, Firm=NAMES[f], Group="HFC" if f in HFC else "Non-HFC NBFC",
                         **{f"CAR{w} %": round(100 * d[w]["car"], 2) for w in WINDOWS}, **{f"p{w}": round(d[w]["p"], 3) for w in WINDOWS}))
pd.DataFrame(rows).to_csv(os.path.join(HERE, "firm_cars.csv"), index=False)
prow = []
for ek in EVENTS:
    for mk in MKT:
        for pn, d in R[ek][mk]["portfolios"].items():
            prow.append(dict(Event=ek, Benchmark=mk, Portfolio=pn, **{f"CAR{w} %": round(100 * d[w]["car"], 2) for w in WINDOWS},
                             **{f"p{w}": round(d[w]["p"], 3) for w in WINDOWS}))
pd.DataFrame(prow).to_csv(os.path.join(HERE, "portfolio_cars.csv"), index=False)

# ------------------------------------------------------------------ figure: cumulative abnormal returns around each event
fig, axs = plt.subplots(1, 2, figsize=(10, 3.8))
for ax, ek in zip(axs, EVENTS):
    for pn, col in [("Non-HFC NBFCs (10)", "#c2410c"), ("HFCs (6, exempt)", "#2b6cb0"), ("Consumer lenders (Bajaj, SBI Cards, Poonawalla)", "#7c3aed")]:
        car = figdata[(ek, pn)].cumsum() * 100
        ax.plot(car.index, car.values, color=col, lw=2, label=pn)
    ax.axvline(-0.5, color="#6b7280", ls="--", lw=1); ax.axhline(0, color="#6b7280", lw=.8)
    ax.set_title(EVENTS[ek]["label"], loc="left", fontsize=9)
    ax.set_xlabel("Trading days relative to day 0", fontsize=8); ax.set_ylabel("Cumulative abnormal return (%)", fontsize=8)
    for s in ["top", "right"]: ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#e5e7eb", lw=.6)
axs[0].legend(frameon=False, fontsize=7, loc="lower left")
fig.tight_layout(); fig.savefig(os.path.join(HERE, "fig_event_car.png"), dpi=200); plt.close(fig)
print("done")
