import pandas as pd, numpy as np, statsmodels.api as sm, warnings; warnings.filterwarnings("ignore")
d=pd.read_excel("/home/user/Claude-Dissertation/RBI_Sectoral_Bank_Credit.xlsx",header=None)
dates=pd.to_datetime(d.iloc[6,3:].astype(str).str.replace(r"\s+"," ",regex=True),format="%B %d, %Y")
num=lambda r: pd.to_numeric(d.iloc[r,3:],errors="coerce").values
s=pd.DataFrame({"NBFC_all":num(11),"HFC":num(12)},index=dates); s["NBFC_exHFC"]=s.NBFC_all-s.HFC
print(s.iloc[[0,30,53,54,55,-1]].round(0))
g=100*np.log(s).diff(); g.loc["2023-07-28"]=np.nan
def its(gap,brk,lbl,lags=6):
    gap=gap.dropna(); X=sm.add_constant((gap.index>=brk).astype(float)); r=sm.OLS(gap.values,X).fit(cov_type="HAC",cov_kwds={"maxlags":lags})
    print(f"{lbl:55s} pre-mean={r.params[0]:6.2f} change={r.params[1]:6.2f} pp/month p={r.pvalues[1]:.4f} n={len(gap)}")
gap=g.NBFC_exHFC-g.HFC
its(gap["2021-04":"2025-01"],"2023-11-30","Nov-23: NBFC(ex-HFC) minus HFC growth")
its(gap["2021-04":"2023-10"],"2022-11-30","PLACEBO Nov-22")
its(gap["2023-12":],"2025-04-30","Rollback Apr-25 (sample Dec-23 on)")
its(gap["2023-12":],"2025-02-28","Rollback Feb-25 announcement")
for per,a,b in [("pre Apr21-Oct23","2021-04","2023-10"),("post Dec23-Mar25","2023-12","2025-03"),("rollback Apr25-","2025-04","2026-12")]:
    print(per, (g[a:b][["NBFC_exHFC","HFC"]].mean()*12).round(1).to_dict())
print("--- quarter-end frequency")
q=s[s.index.month.isin([3,6,9,12])]
q=q.groupby(q.index.to_period("Q")).last()
gq=100*np.log(q).diff(); gq.loc[pd.Period("2023Q3")]=np.nan   # quarter containing HDFC merger
gq.index=gq.index.to_timestamp(how="end")
gapq=gq.NBFC_exHFC-gq.HFC
its(gapq["2021-04":"2025-03"],"2023-12-01","Q: Nov-23 (post = Q4 2023 onward)",lags=2)
its(gapq["2021-04":"2023-09"],"2022-12-01","Q: PLACEBO Nov-22",lags=2)
its(gapq["2024-01":],"2025-06-01","Q: rollback (post = Q2 2025 onward)",lags=2)
print(gapq.round(2).to_string())
