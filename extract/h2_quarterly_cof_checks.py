"""Exploratory: H2 on hand-collected quarterly cost of funds (16 firms)."""
import pandas as pd, numpy as np, sys, warnings; warnings.filterwarnings("ignore")
sys.path.insert(0,"/home/user/Claude-Dissertation"); from did_tools import run
d=pd.read_csv("/home/user/Claude-Dissertation/hand_collected/cof_quarterly.csv")
v=pd.read_csv("/home/user/Claude-Dissertation/data_corrections/bank_exposure_verified.csv").set_index("Company")
d["fy"]=d.FY.str[2:].astype(int); d["q"]=d.Q.str[1].astype(int); d["t"]=(d.fy-21)*4+d.q-1
d["Date_End"]=pd.to_datetime("2020-04-01")+pd.to_timedelta(d.t*91.3+90,unit="D")
d["Date_End"]=d.Date_End.dt.to_period("Q").dt.end_time.dt.normalize()
d["Exp"]=d.Company.map(v.Verified_Bank_Share_Sep2023_Pct)
post=(d.fy>24)|((d.fy==24)&(d.q>=3))
d["PxE"]=post*d.Exp
for lbl,dd in [("all",d),("excl low-quality Arman",d[~d.Note.str.contains("LOW QUALITY")]),("excl proxies (Chola)",d[~d.Note.str.contains("PROXY")]),("FY23+",d[d.fy>=23])]:
    m=run("",dd,"COF_Pct",["PxE"],[],min_obs=8); r=m["res"][0]
    print(f"{lbl:28s} b={r['beta']:.4f} se={r['se_cr1']:.4f} pWCB={r['p_wcb']:.3f} N={m['n']} G={m['firms']}")
for y in range(21,27):
    pass
rep=pd.read_csv("/home/user/Claude-Dissertation/master_panel_data_final.csv").drop_duplicates(["Fiscal_Year","Quarter"])
rep["key"]=rep.Fiscal_Year+rep.Quarter; d["key"]=d.FY+d.Q
d["Repo"]=d.key.map(rep.set_index("key").Repo_Rate_Pct)
d=d.sort_values(["Company","t"]); d["Repo_l2"]=d.groupby("Company").Repo.shift(2)
d["ExR"]=d.Exp*d.Repo; d["ExR2"]=d.Exp*d.Repo_l2
m=run("",d,"COF_Pct",["PxE"],["ExR"],min_obs=8); r=m["res"][0]; print("ctrl Exp x Repo", round(r['beta'],4), round(r['p_wcb'],3))
m=run("",d.dropna(subset=["ExR2"]),"COF_Pct",["PxE"],["ExR2"],min_obs=8); r=m["res"][0]; print("ctrl Exp x Repo(lag2)", round(r['beta'],4), round(r['p_wcb'],3))
# half-year event bins relative to H1FY24 (Q1-Q2 FY24) reference
d["h"]=(d.fy-21)*2+(d.q>2).astype(int)   # 0..11 ; FY24H1 = 6
cols=[]
for h in range(12):
    if h==5: continue  # reference FY23H2 (last pre half before policy H2FY24)
    c=f"E{h}"; d[c]=(d.h==h)*d.Exp; cols.append(c)
m=run("",d,"COF_Pct",cols,[],min_obs=8)
for c,r in zip(cols,m["res"]): print(c, "FY%dH%d"%(21+int(c[1:])//2, int(c[1:])%2+1), round(r['beta'],4), round(r['p_wcb'],3))
