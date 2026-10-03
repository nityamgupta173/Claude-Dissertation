"""Print text of PDF pages containing ALL keywords (| separates alternatives within a keyword).
usage: python tools_pages.py <pdf> "<kw1>" ["<kw2>" ...] [--max N] [--grep REGEX]"""
import re, subprocess, sys
args = sys.argv[1:]; mx = 1; grep = None
if "--max" in args: i = args.index("--max"); mx = int(args[i+1]); del args[i:i+2]
if "--grep" in args: i = args.index("--grep"); grep = re.compile(args[i+1], re.I); del args[i:i+2]
pdf, kws = args[0], [k.lower().split("|") for k in args[1:]]
pages = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout.split("\f")
n = 0
for i, p in enumerate(pages):
    pl = p.lower()
    if all(any(a in pl for a in alts) for alts in kws):
        print(f"----- page {i+1}")
        for l in p.splitlines():
            if l.strip() and (grep is None or grep.search(l)):
                print(re.sub(r" {3,}", "   ", l).rstrip())
        n += 1
        if n >= mx: break
