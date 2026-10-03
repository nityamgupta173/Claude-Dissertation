import pandas as pd, numpy as np, statsmodels.api as sm, warnings; warnings.filterwarnings("ignore")
d=pd.read_excel("/home/user/Claude-Dissertation/RBI_Sectoral_Bank_Credit.xlsx",header=None)
dates=pd.to_datetime(d.iloc[6,3:].astype(str).str.replace(r"\s+"," ",regex=True),format="%B %d, %Y")
num=lambda r: pd.to_numeric(d.iloc[r,3:],errors="coerce").values
s=pd.DataFrame({"NF":num(9),"NBFC":num(11),"HFC":num(12)},index=dates); s["NBFCx"]=s.NBFC-s.HFC
g=100*np.log(s).diff(); g.loc["2023-07-28"]=np.nan
def its(gap,brk,lbl,lags=6):
    gap=gap.dropna(); X=sm.add_constant((gap.index>=brk).astype(float)); r=sm.OLS(gap.values,X).fit(cov_type="HAC",cov_kwds={"maxlags":lags})
    print(f"{lbl:60s} change={r.params[1]:6.2f} pp/month p={r.pvalues[1]:.4f} n={len(gap)}")
its((g.NBFC-g.NF)["2021-04":"2025-01"],"2023-11-30","NBFC (all) minus non-food credit, Nov-23")
its((g.NBFCx-g.NF)["2021-04":"2025-01"],"2023-11-30","NBFC ex-HFC minus non-food credit, Nov-23")
its((g.NBFCx-g.NF)["2021-04":"2023-10"],"2022-11-30","PLACEBO Nov-22")
its((g.NBFCx-g.NF)["2023-12":],"2025-04-30","Rollback (Apr-25), NBFC ex-HFC minus non-food")
its(g.NBFCx["2023-12":],"2025-04-30","Rollback, NBFC ex-HFC growth level")
print({k:(g[a:b][["NBFCx","NF"]].mean()*12).round(1).to_dict() for k,a,b in [("pre","2021-04","2023-10"),("post","2023-12","2025-03"),("rollback","2025-04","2026-12")]})
