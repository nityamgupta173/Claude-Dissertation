"""Bajaj Finance: consolidated business-segment AUM from each quarterly investor presentation."""
import re, subprocess, glob, os, pandas as pd
D = "/home/user/Claude-Dissertation/quarterly_pdfs/raw/Bajaj Finance Ltd"
AFFECTED = ["sales finance", "b2c", "gold loans"]   # consumer durable + personal loans; gold kept in (it sat inside Rural B2C until FY24-Q4)
rows = []
for pdf in sorted(glob.glob(D + "/*.pdf")):
    q = re.search(r"(\d\d)-(\d)\.pdf", pdf); fy, qn = int(q.group(1)), int(q.group(2))
    pages = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout.split("\f")
    pg = next(i for i, p in enumerate(pages) if "business segment wise aum" in p.lower() and "b2c" in p.lower() and re.search(r"^\s*Total\s+[\d,]", p, re.M))
    idx = 2 if (fy, qn) <= (22, 1) else 3          # column of current-quarter consolidated AUM
    seg = {}
    for l in pages[pg].splitlines():
        m = re.match(r"\s*([A-Za-z][A-Za-z0-9&\-\(\) ]*?[A-Za-z0-9\)])\*?\s{2,}(.*)", l)
        if not m: continue
        toks = re.findall(r"\(?-?[\d,]+(?:\.\d+)?%?\)?|(?<=\s)-(?=\s)", " " + m.group(2) + " ")
        nums = [t for t in toks if "%" not in t and "(" not in t]
        if len(nums) <= idx: continue
        v = nums[idx]
        seg[m.group(1).strip()] = 0.0 if v == "-" else float(v.replace(",", ""))
    total = next(v for k, v in seg.items() if k.lower() == "total")
    lab = {k: v for k, v in seg.items() if k.lower() != "total" and not k.lower().startswith("assets")}
    aff = sum(v for k, v in lab.items() if any(a in k.lower() for a in AFFECTED))
    rows.append(dict(Company="Bajaj Finance Ltd.", FY=f"FY{fy}", Q=f"Q{qn}", Total_AUM_Cr=total,
                     Affected_Consumer_AUM_Cr=aff, Sum_Segments=sum(lab.values()), Source=f"{fy}-{qn} p.{pg+1}",
                     Segments="; ".join(f"{k}={v:,.0f}" for k, v in lab.items())))
df = pd.DataFrame(rows)
df["Check_pct"] = (df.Sum_Segments / df.Total_AUM_Cr - 1) * 100
print(df[["FY", "Q", "Total_AUM_Cr", "Affected_Consumer_AUM_Cr", "Check_pct", "Source"]].to_string())
df.drop(columns="Sum_Segments").to_csv("/home/user/Claude-Dissertation/hand_collected/segments_bajaj.csv", index=False)
