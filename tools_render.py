"""Render the first page of a deck whose text contains all keywords: python tools_render.py <firm-prefix> <deck e.g. 24-2> <out-tag> kw1 [kw2 ...] [--dpi N] [--nth K]"""
import glob, subprocess, sys
args = sys.argv[1:]; dpi = 55; nth = 0
if "--dpi" in args: i = args.index("--dpi"); dpi = int(args[i+1]); del args[i:i+2]
if "--nth" in args: i = args.index("--nth"); nth = int(args[i+1]); del args[i:i+2]
firm, deck, tag, kws = args[0], args[1], args[2], [k.lower() for k in args[3:]]
pdf = [p for p in glob.glob("/home/user/Claude-Dissertation/quarterly_pdfs/raw/*/*.pdf") if p.split("/")[-1].startswith(firm) and p.endswith(f" {deck}.pdf")][0]
pages = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout.split("\f")
hits = [i+1 for i, p in enumerate(pages) if all(any(a in p.lower() for a in k.split("|")) for k in kws)]
if len(hits) <= nth: print("NOPAGE", deck); sys.exit()
p = hits[nth]
subprocess.run(["pdftoppm", "-f", str(p), "-l", str(p), "-r", str(dpi), "-png", "-singlefile", pdf, f"/tmp/claude-0/img/{tag}"])
print(f"/tmp/claude-0/img/{tag}.png  (page {p})")
