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

# ---------------------------------------------------------------- 1. Screener patch
PATCH_COLS = ["Interest_Income_Cr", "Operating_Expenses_Cr", "Net_Profit_PAT_Cr", "Net_Worth_Cr"]
log("NA counts before patch: " + str({c: int(df[c].isna().sum()) for c in PATCH_COLS}))
SCREENER_MAP = {"Sales": "Interest_Income_Cr", "Expenses": "Operating_Expenses_Cr",
                "Net profit": "Net_Profit_PAT_Cr"}


def read_screener_quarters(path):
    """Best-effort parser for a Screener.in export ('Data Sheet', quarterly block).
    Returns DataFrame indexed by quarter-end date with the mapped columns."""
    raw = pd.read_excel(path, sheet_name="Data Sheet", header=None)
    start = raw.index[raw[0].astype(str).str.strip().eq("Quarters")]
    if len(start) == 0:
        return None
    blk = raw.iloc[start[0] + 1: start[0] + 16].set_index(0)
    dates = pd.to_datetime(blk.loc["Report Date"].dropna(), errors="coerce")
    out = {}
    for src, dst in SCREENER_MAP.items():
        if src in blk.index:
            out[dst] = pd.to_numeric(blk.loc[src].reindex(dates.index), errors="coerce").values
    res = pd.DataFrame(out, index=dates.values)
    return res


screener_files = glob.glob(os.path.join(HERE, "screener", "*.xls*")) + \
    glob.glob(os.path.join(HERE, "*Screener*.xls*")) + glob.glob(os.path.join(HERE, "*screener*.xls*"))
if not screener_files:
    log("!! No Screener.in Excel files found (looked in ./screener/ and repo root). "
        "P&L NAs were NOT patched. Drop the 16 files in ./screener/ and re-run.")
else:
    filled = 0
    for f in screener_files:
        name = os.path.basename(f).lower()
        firm = next((c for c in df.Company.unique() if c.lower().split()[0] in name), None)
        scr = read_screener_quarters(f)
        if firm is None or scr is None:
            log(f"   could not match/parse {f}")
            continue
        for dt, row in scr.iterrows():
            m = (df.Company == firm) & (df.Date_End == dt)
            for col, val in row.items():
                if pd.notna(val):
                    idx = df.index[m & df[col].isna()]
                    df.loc[idx, col] = val
                    filled += len(idx)
    log(f"Screener patch filled {filled} cells (only NA cells; nothing overwritten).")
log("NA counts after patch:  " + str({c: int(df[c].isna().sum()) for c in PATCH_COLS}))

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
