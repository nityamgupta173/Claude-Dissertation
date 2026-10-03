"""Decode Google Drive connector downloads (saved tool-result JSON) into quarterly_pdfs/<file title>."""
import base64, glob, json, os
SRC = "/root/.claude/projects/-home-user-Claude-Dissertation/e85a83e9-5009-54f3-a45e-c0492ff33206/tool-results"
DST = "/home/user/Claude-Dissertation/quarterly_pdfs"
os.makedirs(DST, exist_ok=True)
n = 0
for f in glob.glob(os.path.join(SRC, "mcp-Google_Drive-download_file_content-*.txt")):
    try:
        j = json.load(open(f))
        b = base64.b64decode(j["content"])
        assert b[:4] == b"%PDF"
        open(os.path.join(DST, j["title"]), "wb").write(b)
        os.remove(f); n += 1
    except Exception as e:
        print("FAILED", f, e)
print("decoded", n, "| total pdfs:", len(glob.glob(DST + "/*.pdf")))
