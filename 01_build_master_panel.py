"""Phase 1: clean, merge and freeze exposure -> master_panel_data_final.csv

Run:  python 01_build_master_panel.py
"""
import glob
import os
import warnings

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(HERE, "outputs"), exist_ok=True)
LOG = []


def log(msg):
    print(msg)
    LOG.append(msg)


df = pd.read_csv(os.path.join(HERE, "nbfc_quarterly_data-v3.csv"))
df["Date_End"] = pd.to_datetime(df["Date_End"], format="%d-%m-%Y")
df["FQ"] = df["Fiscal_Year"] + "-" + df["Quarter"]
log(f"Loaded raw panel: {df.Company.nunique()} firms, {len(df)} rows")

# ---------------------------------------------------------------- 0. Verified corrections (company filings)
# Shriram Finance: the CSV 'unsecured' series is 9.00% x total AUM in most quarters (an assumed share) and
# the actual Personal Loans figure in two quarters (Q3 FY24, Q3 FY25). Filings (investor presentations,
# see data_corrections/shriram_filings.csv) give real Personal Loans for Q2 FY24-Q1 FY26 only. A consistent
# series is therefore not available: Unsecured and Secured AUM are set to NA for Shriram (kept in
# Verified_PersonalLoan_Cr where filed), and Total AUM is replaced by the filed figure where available.
# Flag series that are a constant % of total AUM (assumed split, not reported data)
_r = (df.Unsecured_Consumer_AUM_Cr / df.Total_AUM_Cr).round(3)
qual = _r.groupby(df.Company).agg(lambda x: x.dropna().nunique())
nonzero = df.groupby("Company").Unsecured_Consumer_AUM_Cr.max() > 0
df["Unsecured_Series_Flag"] = df.Company.map(lambda c: "no unsecured book" if not nonzero[c] else
                                            ("by definition (cards = 100%)" if "SBI Cards" in c else
                                             ("MIXED (9% assumed + 2 actual quarters) -> set to NA" if "Shriram" in c else
                                             ("ASSUMED constant share of AUM" if qual[c] <= 2 else "varies (reported?)"))))
log("Unsecured series flags: " + str(df.drop_duplicates("Company").set_index("Company").Unsecured_Series_Flag.to_dict()))

df["Verified_PersonalLoan_Cr"] = np.nan
corr = pd.read_csv(os.path.join(HERE, "data_corrections", "shriram_filings.csv"))
sh = df.Company.str.contains("Shriram")
for _, r in corr.iterrows():
    m = sh & (df.Fiscal_Year == r.Fiscal_Year) & (df.Quarter == r.Quarter)
    df.loc[m, "Total_AUM_Cr"] = round(r.Total_AUM_Rs_mn / 10, 1)
    df.loc[m, "Verified_PersonalLoan_Cr"] = round(r.Personal_Loans_Rs_mn / 10, 1)
df.loc[sh, ["Unsecured_Consumer_AUM_Cr", "Secured_AUM_Cr"]] = np.nan
log("Shriram corrected from filings: Total AUM replaced for 8 quarters; Unsecured/Secured set to NA (assumed-share series).")

# ---------------------------------------------------------------- 1. Screener patch
# Screener.in 'Data Sheet' exports: quarterly block covers only the latest ~9-10 quarters (Mar-2024 on);
# annual P&L / balance sheet cover FY17-FY26.
# Definition check against the existing CSV (overlapping non-NA cells):
#   Screener 'Interest'   == CSV Interest_Expense_Cr  (exact)   -> used to patch NAs
#   Screener 'Net profit' ~= CSV Net_Profit_PAT_Cr     (<2%)    -> used to patch NAs
#   Screener 'Sales'      != CSV Interest_Income_Cr    (Sales includes fee/other income, ~10-15% higher)
#   Screener 'Expenses'   != CSV Operating_Expenses_Cr (different definition)
# so Sales/Expenses are NOT written into the CSV's columns (mixing definitions inside one firm's series would
# be wrong); they are added as separate Screener_* columns instead.
SCREENER_FILES = {"AAVAS": "AAVAS Financiers", "Arman": "Arman Financial", "Bajaj": "Bajaj Finance",
                  "Can Fin": "Can Fin Homes", "Cholamandalam": "Cholaman.Inv.&Fn", "CreditAccess": "CreditAcc. Gram",
                  "Home First": "Home First Finan", "IIFL": "IIFL Finance", "LIC": "LIC Housing Fin",
                  "Mahindra": "M & M Fin. Serv", "Muthoot": "Muthoot Finance", "PNB": "PNB Housing",
                  "Poonawalla": "Poonawalla Fin", "Repco": "Repco Home Fin", "SBI": "SBI Cards",
                  "Shriram": "Shriram Finance"}
PATCH_COLS = ["Interest_Income_Cr", "Operating_Expenses_Cr", "Net_Profit_PAT_Cr", "Net_Worth_Cr", "Interest_Expense_Cr"]
log("NA counts before patch: " + str({c: int(df[c].isna().sum()) for c in PATCH_COLS}))
for c in ["Screener_Revenue_Cr", "Screener_Expenses_Cr", "Screener_Interest_Cr", "Screener_NetProfit_Cr"]:
    df[c] = np.nan


def block(raw, title, nrows):
    i = raw.index[raw[0].astype(str).str.strip() == title][0]
    b = raw.iloc[i + 1: i + 1 + nrows].set_index(0)
    dates = pd.to_datetime(b.loc["Report Date"], errors="coerce").dropna()
    return b.drop(index="Report Date"), dates


annual_rows = []
filled = {"Net_Profit_PAT_Cr": 0, "Interest_Expense_Cr": 0}
mism = []
for key, fn in SCREENER_FILES.items():
    firm = next(c for c in df.Company.unique() if key.lower() in c.lower())
    raw = pd.read_excel(os.path.join(HERE, fn + ".xlsx"), sheet_name="Data Sheet", header=None)
    qb, qd = block(raw, "Quarters", 9)
    q = lambda r: pd.to_numeric(qb.loc[r, qd.index], errors="coerce").values
    sq = pd.DataFrame({"rev": q("Sales"), "exp": q("Expenses"), "int": q("Interest"), "np": q("Net profit")},
                      index=qd.values)
    for dt, r in sq.iterrows():
        m = (df.Company == firm) & (df.Date_End == dt)
        if not m.any():
            continue
        df.loc[m, ["Screener_Revenue_Cr", "Screener_Expenses_Cr", "Screener_Interest_Cr", "Screener_NetProfit_Cr"]] = \
            [r["rev"], r["exp"], r["int"], r["np"]]
        for col, v in [("Net_Profit_PAT_Cr", r["np"]), ("Interest_Expense_Cr", r["int"])]:
            idx = df.index[m & df[col].isna()]
            if pd.notna(v) and len(idx):
                df.loc[idx, col] = v
                filled[col] += len(idx)
            elif pd.notna(v) and abs(df.loc[m, col].iloc[0] - v) > 0.03 * abs(v) + 1:
                mism.append((firm, str(dt.date()), col, float(df.loc[m, col].iloc[0]), float(v)))
    # annual panel (for Model 2)
    ab, ad = block(raw, "PROFIT & LOSS", 16)
    bb, bd = block(raw, "BALANCE SHEET", 12)
    for dt in ad.values:
        yr = pd.Timestamp(dt)
        g = lambda blk, r, d_: pd.to_numeric(blk.loc[r, d_ == d_], errors="coerce") if False else None
    A = pd.DataFrame({"Sales": pd.to_numeric(ab.loc["Sales", ad.index], errors="coerce").values,
                      "Interest": pd.to_numeric(ab.loc["Interest", ad.index], errors="coerce").values,
                      "NetProfit": pd.to_numeric(ab.loc["Net profit", ad.index], errors="coerce").values},
                     index=pd.DatetimeIndex(ad.values))
    Bs = pd.DataFrame({"Borrowings": pd.to_numeric(bb.loc["Borrowings", bd.index], errors="coerce").values,
                       "Equity": (pd.to_numeric(bb.loc["Equity Share Capital", bd.index], errors="coerce") +
                                  pd.to_numeric(bb.loc["Reserves", bd.index], errors="coerce")).values,
                       "TotalAssets": pd.to_numeric(bb.loc["Total", bd.index].iloc[0:len(bd)] if False else
                                                    bb.loc[bb.index == "Total", bd.index].iloc[0], errors="coerce").values},
                      index=pd.DatetimeIndex(bd.values))
    A = A.join(Bs)
    A["Company"] = firm
    annual_rows.append(A.reset_index().rename(columns={"index": "FY_End"}))
log(f"Screener patch filled (NA cells only): {filled}")
if mism:
    log(f"  {len(mism)} overlapping cells differ >3% between CSV and Screener (CSV kept): {mism[:6]}")
log("NA counts after patch:  " + str({c: int(df[c].isna().sum()) for c in PATCH_COLS}))
ann = pd.concat(annual_rows)
ann = ann.sort_values(["Company", "FY_End"])
ann["Borrowings_avg"] = ann.groupby("Company").Borrowings.transform(lambda x: (x + x.shift(1)) / 2)
ann["Assets_avg"] = ann.groupby("Company").TotalAssets.transform(lambda x: (x + x.shift(1)) / 2)
ann["Interest_Expense_Ratio_Pct"] = ann.Interest / ann.Borrowings_avg * 100
ann["RoA_annual_Pct"] = ann.NetProfit / ann.Assets_avg * 100
ann["FY"] = ann.FY_End.dt.year
ann = ann[ann.FY.between(2021, 2026)]
ann.to_csv(os.path.join(HERE, "screener_annual_panel.csv"), index=False)
log(f"Saved screener_annual_panel.csv ({len(ann)} firm-years)")

# ---------------------------------------------------------------- 2. Macro controls
# quarter label -> calendar quarter start
fy_start_year = 2000 + df["Fiscal_Year"].str[2:].astype(int) - 1
qn = df["Quarter"].str[1].astype(int)
df["q_start"] = pd.to_datetime(dict(year=fy_start_year + (qn == 4).astype(int),
                                    month=np.select([qn == 1, qn == 2, qn == 3, qn == 4], [4, 7, 10, 1]),
                                    day=1))

# (a) GDP growth YoY from the FRED real-GDP levels (NSA) -> percent
g = pd.read_excel(os.path.join(HERE, "NGDPRNSAXDCINQ (1).xlsx"), sheet_name="Quarterly")
g["observation_date"] = pd.to_datetime(g["observation_date"])
g = g.set_index("observation_date")["NGDPRNSAXDCINQ"]
gdp = (g / g.shift(4) - 1) * 100
gdp.name = "GDP_Growth_YoY_Pct"
chk = pd.read_excel(os.path.join(HERE, "India_GDP_Quarterly_YoY_FY21_FY27.xlsx"))
chk_v = (chk["GDP_Growth_Pct"] * 100).values[:24]
mine = gdp.loc["2020-04-01":"2026-01-01"].values
log(f"GDP YoY cross-check, FRED-computed vs supplied YoY file: max abs diff = "
    f"{np.abs(mine - chk_v).max():.3f} pp")

# (b) 5-year yield: quarterly mean of monthly closes
by = pd.read_csv(os.path.join(HERE, "India 5-Year Bond Yield Historical Data.csv"))
by["Date"] = pd.to_datetime(by["Date"], format="%m/%d/%Y")
by = by.set_index("Date")["Price"].sort_index()
bond = by.resample("QS-JAN").mean()          # calendar quarters
bond.name = "Bond_Yield_5Y_Pct"
# shift calendar quarters onto the Apr-start grid used by q_start (already month 4/7/10/1)

# (c) RBI policy repo rate: day-weighted quarterly average of the step series
r = pd.read_excel(os.path.join(HERE, "Major Monetary Policy Rates and Reserve Requirements - Bank Rate, LAF (Repo, Reverse Repo, SDF and MSF) Rates, CRR & SLR.xlsx"),
                  header=None, sheet_name=0)
r = r.iloc[8:, [1, 2]]
r.columns = ["date", "repo"]
r["date"] = pd.to_datetime(r["date"], errors="coerce")
r["repo"] = pd.to_numeric(r["repo"], errors="coerce")          # '-' = no change
r = r.dropna().drop_duplicates("date").set_index("date").sort_index()["repo"]
daily = r.reindex(pd.date_range("2019-01-01", "2026-06-30")).ffill()
repo = daily.resample("QS-JAN").mean()
repo.name = "Repo_Rate_Pct"
log(f"Repo series ends {r.index.max().date()} at {r.iloc[-1]}%; carried forward to Mar-2026 (no later change in file).")

df = df.merge(gdp.rename_axis("q_start").reset_index(), on="q_start", how="left") \
       .merge(bond.rename_axis("q_start").reset_index(), on="q_start", how="left") \
       .merge(repo.rename_axis("q_start").reset_index(), on="q_start", how="left")
assert df[["GDP_Growth_YoY_Pct", "Bond_Yield_5Y_Pct", "Repo_Rate_Pct"]].notna().all().all()
log("Macro merge complete, no missing macro values.")

# ---------------------------------------------------------------- 3. Balance the panel
n = df.groupby("Company").size()
full = n.max()
drop = n[n < full].index.tolist()
if drop:
    log(f"Balancing: dropped {drop} (has {n[n < full].tolist()} of {full} quarters).")
df = df[~df.Company.isin(drop)].copy()
df = df.sort_values(["Company", "Date_End"]).reset_index(drop=True)
assert df.groupby("Company").size().nunique() == 1

# ---------------------------------------------------------------- 4. Freeze baseline exposure (Q2 FY24)
base = df[(df.Fiscal_Year == "FY24") & (df.Quarter == "Q2")].set_index("Company")
df["Pre_Policy_Bank_Exposure"] = df.Company.map(base["Bank_Borrowing_Share_Pct"])
df["Unsecured_AUM_Ratio"] = df["Unsecured_Consumer_AUM_Cr"] / df["Total_AUM_Cr"]
df["Exposure_Unsecured_Baseline"] = df.Company.map(base["Unsecured_Consumer_AUM_Cr"] / base["Total_AUM_Cr"])
df["Pre_Policy_Tier1_CRAR"] = df.Company.map(base["Tier_1_CRAR_Pct"])

# ---------------------------------------------------------------- 5. Dummies, logs, growth, lags
df["t"] = df.groupby("Company").cumcount()                       # 0 = Q1 FY21
df["Post_Nov2023"] = (df.Date_End >= "2023-12-31").astype(int)   # Q3 FY24 onwards
df["Post_Nov2022"] = (df.Date_End >= "2022-12-31").astype(int)   # Q3 FY23 onwards (placebo)
df["Post_Feb2025"] = (df.Date_End >= "2025-03-31").astype(int)   # Q4 FY25 onwards
df["ln_Total_AUM"] = np.log(df.Total_AUM_Cr)
for src, dst in [("Unsecured_Consumer_AUM_Cr", "dlog_Unsecured"), ("MFI_AUM_Cr", "dlog_MFI"),
                 ("Secured_AUM_Cr", "dlog_Secured"), ("Total_AUM_Cr", "dlog_Total")]:
    s = df[src].where(df[src] > 0)           # log undefined for zero/missing AUM -> NaN
    df[dst] = df.assign(_l=np.log(s)).groupby("Company")["_l"].diff()
df["lag_RoA_Pct"] = df.groupby("Company")["RoA_Pct"].shift(1)
df["lag_GNPA_Pct"] = df.groupby("Company")["Gross_NPA_GS3_Pct"].shift(1)
df["lag_Tier1_CRAR"] = df.groupby("Company")["Tier_1_CRAR_Pct"].shift(1)
df["Interest_Expense_Ratio_Pct"] = df.Interest_Expense_Cr / df.Total_Borrowings_Cr * 100

df = df.drop(columns=["q_start"])
out = os.path.join(HERE, "master_panel_data_final.csv")
df.to_csv(out, index=False, date_format="%Y-%m-%d")
log(f"Saved {out}: {df.Company.nunique()} firms x {df.groupby('Company').size().iloc[0]} quarters = {len(df)} rows")
with open(os.path.join(HERE, "outputs", "phase1_log.txt"), "w") as fh:
    fh.write("\n".join(LOG))
