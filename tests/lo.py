"""Ask a running LibreOffice what it sees in a workbook (independent of xlpolish)."""
import subprocess, time, os, uno
from com.sun.star.beans import PropertyValue

_proc = None
def start():
    global _proc
    if _proc: return
    _proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                              '--accept=socket,host=localhost,port=2002;urp;'],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(60):
        try:
            ctx = uno.getComponentContext()
            r = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
            r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); return
        except Exception: time.sleep(1)
    raise RuntimeError("LibreOffice did not start")

def stop():
    global _proc
    if _proc: _proc.terminate(); _proc.wait(); _proc = None

def open_doc(path):
    start()
    ctx = uno.getComponentContext()
    r = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
    rctx = r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desk = rctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", rctx)
    p = PropertyValue(); p.Name = "Hidden"; p.Value = True
    return desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, (p,))
