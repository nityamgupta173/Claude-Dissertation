"""Exploratory: within-firm DiD, Bajaj affected vs exempt segments (FY21-FY25)."""
import pandas as pd, numpy as np, re, warnings; warnings.filterwarnings("ignore")
from linearmodels.panel import PanelOLS
d=pd.read_csv("/home/user/Claude-Dissertation/hand_collected/segments_bajaj.csv")
rows=[]
for _,r in d.iterrows():
    for kv in r.Segments.split("; "):
        k,v=kv.split("="); k=k.strip(); v=float(v.replace(",",""))
        k=re.sub(r"\s*\(.*?\)|\*|Business$","",k).strip()
        k={"Auto Finance":"Two & Three wheeler Finance","Two & Three-wheeler Finance":"Two & Three wheeler Finance","Sales Finance":"Urban Sales Finance","Consumer B2C":"Urban B2C","Securities Lending":"Loan Against Securities","Mortgages Business":"Mortgages","SME lending":"SME Lending","Rural B2C":"Rural B2C","Urban B2C Loans":"Urban B2C","Rural B2C Loans":"Rural B2C","Two and Three-wheeler Finance":"Two & Three wheeler Finance"}.get(k,k)
        if k in ("Gold Loans",): k="Rural B2C"
        if k in ("Car Loans",): k="SME Lending"
        rows.append((r.FY,r.Q,k,v))
s=pd.DataFrame(rows,columns=["FY","Q","Seg","AUM"]).groupby(["FY","Q","Seg"]).AUM.sum().reset_index()
s["t"]=(s.FY.str[2:].astype(int)-21)*4+s.Q.str[1].astype(int)-1
s=s.sort_values(["Seg","t"]); s["g"]=100*s.groupby("Seg").AUM.transform(lambda x: np.log(x).diff())
print(s.groupby("Seg").size())
aff=["Urban Sales Finance","Urban B2C","Rural Sales Finance","Rural B2C"]
s["Treat"]=s.Seg.isin(aff).astype(int); s["Post"]=(s.t>=14).astype(int)  # Q3FY24 = t 14
s["TxP"]=s.Treat*s.Post
s=s[s.t<=19].dropna(subset=["g"]); s=s[s.Seg.map(s.groupby("Seg").size())>=18]
p=s.set_index(["Seg","t"])
for cov in [dict(cov_type="clustered",cluster_entity=True),dict(cov_type="kernel",kernel="bartlett",bandwidth=4)]:
    r=PanelOLS(p.g,p[["TxP"]],entity_effects=True,time_effects=True).fit(**cov)
    print(cov["cov_type"], round(r.params.iloc[0],2), round(r.std_errors.iloc[0],2), round(r.pvalues.iloc[0],4), r.nobs)
pre=s[s.Post==0].groupby("Treat").g.mean(); post_=s[s.Post==1].groupby("Treat").g.mean(); print("mean QoQ growth pre/post (0=exempt,1=affected):",pre.round(2).to_dict(),post_.round(2).to_dict())
