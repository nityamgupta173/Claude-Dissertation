import pandas as pd, numpy as np, sys, warnings; warnings.filterwarnings("ignore")
sys.path.insert(0,"/home/user/Claude-Dissertation"); from did_tools import run
d=pd.read_csv("/home/user/Claude-Dissertation/hand_collected/cof_quarterly.csv")
d["fy"]=d.FY.str[2:].astype(int); d["q"]=d.Q.str[1].astype(int); d["t"]=(d.fy-21)*4+d.q-1
d["Date_End"]=pd.period_range("2020Q2",periods=24,freq="Q").to_timestamp(how="end").normalize()[d.t]
HFC=["AAVAS","Can Fin","Home First","LIC Housing","PNB Housing","Repco"]
d["NonHFC"]=(~d.Company.str.contains("|".join(HFC))).astype(int)
post=(d.t>=14).astype(int); rev=(d.t>=20).astype(int)  # Q3FY24 ; Q1FY26 (rollback effective 1-Apr-2025)
d["NxPost"]=d.NonHFC*post; d["NxRev"]=d.NonHFC*rev
for lbl,dd,tr in [("all",d,["NxPost"]),("pre-rollback sample",d[d.t<20],["NxPost"]),("with rollback",d,["NxPost","NxRev"]),("excl Arman lowq",d[~d.Note.str.contains("LOW Q")],["NxPost","NxRev"])]:
    m=run("",dd,"COF_Pct",tr,[],min_obs=8)
    print(lbl, [(t,round(r['beta'],3),round(r['p_wcb'],3)) for t,r in zip(tr,m['res'])], m['n'], m['firms'])
