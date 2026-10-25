"""Find the printed page number of every figure and table caption in Dissertation.pdf (arabic-numbered pages only)
and write outputs/final/caption_pages.json for the List of Figures / List of Tables."""
import json, re, subprocess
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", "Dissertation.pdf"], capture_output=True, text=True).stdout).group(1))
pages = {}
for p in range(1, n + 1):
    t = subprocess.run(["pdftotext", "-layout", "-f", str(p), "-l", str(p), "Dissertation.pdf", "-"], capture_output=True, text=True).stdout
    m = re.search(r"(\d+)\s*\|\s*P\s*a\s*g\s*e", t)
    if not m:
        continue                                   # cover or roman-numbered page
    for k in re.findall(r"^\s*((?:Figure|Table) [0-9A-Z]+\.\d+):", t, flags=re.M):
        pages.setdefault(k, m.group(1))
json.dump(pages, open("outputs/final/caption_pages.json", "w"), indent=1)
print(len(pages), "captions located")
