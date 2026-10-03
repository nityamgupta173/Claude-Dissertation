"""H1 (RBI sectoral bank credit): treated vs exempt credit categories around the Nov-2023 risk-weight rise.
Exploratory checks; a final version with tables/figures follows."""
import pandas as pd, numpy as np, statsmodels.api as sm, warnings; warnings.filterwarnings("ignore")
from linearmodels.panel import PanelOLS
d=pd.read_excel("RBI_Sectoral_Bank_Credit.xlsx",header=None)
dates=pd.to_datetime(d.iloc[6,3:].astype(str).str.replace(r"\s+"," ",regex=True),format="%B %d, %Y")
rows={"NBFCs":11,"Credit cards":15,"Other personal":19,"Housing":13,"Vehicle":17,"Education":16,"Gold":18}
s=pd.DataFrame({k:pd.to_numeric(d.iloc[r,3:],errors="coerce").values for k,r in rows.items()},index=dates)
g=100*np.log(s).diff(); g.loc["2023-07-28",["NBFCs","Housing"]]=np.nan
T=["NBFCs","Credit cards","Other personal"]; C=["Housing","Vehicle","Education"]
def its(gg,tc,cc,brk,lbl):
    gap=(gg[tc].mean(1)-gg[cc].mean(1)).dropna()
    X=sm.add_constant((gap.index>=brk).astype(float)); r=sm.OLS(gap.values,X).fit(cov_type="HAC",cov_kwds={"maxlags":6})
    print(f"{lbl:50s} change={r.params[1]:6.2f} p={r.pvalues[1]:.4f} n={len(gap)}")
def panel(gg,tc,cc,brk,lbl,post_col=None):
    L=gg[tc+cc].stack().rename("g").reset_index(); L.columns=["date","series","g"]
    L["Treat_x_Post"]=(L.series.isin(tc)&(L.date>=brk)).astype(float)
    p=L.set_index(["series","date"])
    r=PanelOLS(p.g,p[["Treat_x_Post"]],entity_effects=True,time_effects=True).fit(cov_type="kernel",kernel="bartlett",bandwidth=6)
    print(f"{lbl:50s} b={r.params.iloc[0]:6.2f} p={r.pvalues.iloc[0]:.4f} N={r.nobs}")
W=g["2021-04-01":"2025-01-31"]
its(W,T,C,"2023-11-30","ITS main"); panel(W,T,C,"2023-11-30","Panel DiD (Driscoll-Kraay) main")
P=g["2021-04-01":"2023-10-31"]; its(P,T,C,"2022-11-30","PLACEBO Nov-22 ITS"); panel(P,T,C,"2022-11-30","PLACEBO Nov-22 panel")
for drop in T: panel(W,[x for x in T if x!=drop],C,"2023-11-30",f"drop treated {drop}")
for drop in C: panel(W,T,[x for x in C if x!=drop],"2023-11-30",f"drop control {drop}")
panel(W,T,C+["Gold"],"2023-11-30","+gold control")
panel(g["2022-04-01":"2025-01-31"],T,C,"2023-11-30","shorter pre window (from Apr-22)")
panel(g["2021-04-01":"2024-10-31"],T,C,"2023-11-30","shorter post window (to Oct-24)")
for t in T: panel(W,[t],C,"2023-11-30",f"single treated: {t}")
# reversal: Feb-25 announcement, NBFC risk weight restored from 1-Apr-25
R=g["2023-12-01":]
for brk in ["2025-02-28","2025-04-30"]:
    panel(R,["NBFCs"],C,brk,f"REVERSAL NBFC vs controls, break {brk}"); its(R,["NBFCs"],C,brk,f"REVERSAL ITS NBFC {brk}")
    panel(R,["Credit cards","Other personal"],C,brk,f"REVERSAL consumer (not relaxed) {brk}")
print((g.loc["2023-12-01":"2025-02-28",T+C].mean()*12).round(1).to_dict()); print((g.loc["2025-04-01":,T+C].mean()*12).round(1).to_dict())
