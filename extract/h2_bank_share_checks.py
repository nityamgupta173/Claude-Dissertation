import pandas as pd, numpy as np, sys, warnings; warnings.filterwarnings("ignore")
sys.path.insert(0,"/home/user/Claude-Dissertation"); from did_tools import run, wcb_t
d=pd.read_csv("/home/user/Claude-Dissertation/hand_collected/bank_share_quarterly.csv")
d["fy"]=d.FY.str[2:].astype(int); d["q"]=d.Q.str[1].astype(int); d["t"]=(d.fy-21)*4+d.q-1
d["Date_End"]=pd.period_range("2020Q2",periods=24,freq="Q").to_timestamp(how="end").normalize()[d.t]
HFC=["AAVAS","Can Fin","Home First","LIC Housing","PNB Housing","Repco"]
d["NonHFC"]=(~d.Company.str.contains("|".join(HFC))).astype(int)
d["Post"]=(d.t>=14).astype(int); d["Rev"]=(d.t>=20).astype(int)
d["NxPost"]=d.NonHFC*d.Post; d["NxRev"]=d.NonHFC*d.Rev
print(d.groupby("Company").size().to_string()); print("N",len(d))
print(d.groupby(["NonHFC","Post"]).Bank_Share_Pct.mean().unstack().round(1))
for lbl,dd,tr in [("Nov-23, pre-rollback sample",d[d.t<20],["NxPost"]),("Nov-23 + rollback",d,["NxPost","NxRev"]),("excl IIFL (annual only)",d[~d.Company.str.contains("IIFL")],["NxPost","NxRev"]),("semiannual Q2/Q4 only",d[d.q.isin([2,4])],["NxPost","NxRev"])]:
    m=run("",dd,"Bank_Share_Pct",tr,[],min_obs=4)
    print(f"{lbl:32s}", [(t,round(r['beta'],2),round(r['se_cr1'],2),round(r['p_wcb'],3)) for t,r in zip(tr,m['res'])], m['n'], m['firms'])
# placebo: fake Nov-22 using pre-policy data
pre=d[d.t<14].copy(); pre["NxFake"]=pre.NonHFC*(pre.t>=10)
m=run("",pre,"Bank_Share_Pct",["NxFake"],[],min_obs=3); r=m["res"][0]; print("PLACEBO Nov-22", round(r['beta'],2), round(r['p_wcb'],3), m['n'], m['firms'])
print("--- dynamics (ref: FY23H2 = t 10-11)")
d["h"]=(d.t//2)
cols=[]
for h in range(12):
    if h==5: continue
    c=f"H{h}"; d[c]=d.NonHFC*(d.h==h); cols.append(c)
m=run("",d,"Bank_Share_Pct",cols,[],min_obs=4)
for c,r in zip(cols,m["res"]): print(f"FY{21+int(c[1:])//2}H{int(c[1:])%2+1}", round(r['beta'],2), round(r['p_wcb'],3))
d["Post2"]=(d.t>=16).astype(int); d["NxPost2"]=d.NonHFC*d.Post2   # from FY25 (allowing adjustment lag)
m=run("",d[d.t<20],"Bank_Share_Pct",["NxPost2"],[],min_obs=4); r=m['res'][0]; print("FY25 onward vs pre (pre-rollback)", round(r['beta'],2), round(r['p_wcb'],3))
