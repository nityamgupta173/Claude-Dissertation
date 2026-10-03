"""Dissertation.docx -> Dissertation.pdf, and Dissertation.docx re-saved with the table of contents and page numbers filled in (LibreOffice UNO).
Run: python3 tools_pdf.py"""
import os, subprocess, time, uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
SRC, OUT = os.path.join(HERE, "Dissertation.docx"), os.path.join(HERE, "Dissertation.pdf")


def prop(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p


proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                         "--accept=socket,host=localhost,port=2002;urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = uno.getComponentContext()
resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
for _ in range(60):
    try:
        rctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(1)
desktop = rctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", rctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(SRC), "_blank", 0, (prop("Hidden", True),))
for _ in range(2):  # second pass picks up page shifts caused by the first
    idx = doc.getDocumentIndexes()
    for i in range(idx.getCount()):
        idx.getByIndex(i).update()
    doc.refresh()
doc.storeToURL(uno.systemPathToFileUrl(OUT), (prop("FilterName", "writer_pdf_Export"),))
# re-save the Word file so it carries the generated TOC entries and cached page numbers
doc.storeToURL(uno.systemPathToFileUrl(SRC), (prop("FilterName", "MS Word 2007 XML"), prop("Overwrite", True)))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print("written", OUT)
