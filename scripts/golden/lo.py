"""LibreOffice UNO helper: run soffice headless and open a workbook with macros disabled (MacroExecutionMode 0)."""
import subprocess, time, uno, os, sys
from com.sun.star.beans import PropertyValue

def prop(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

class LO:
    def __init__(self, port=2002):
        self.proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--nologo', '--norestore',
            f'--accept=socket,host=localhost,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(60):
            try:
                self.ctx = resolver.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext'); break
            except Exception: time.sleep(0.5)
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
    def open(self, path):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        # MacroExecutionMode 0 = NEVER_EXECUTE
        self.doc = self.desktop.loadComponentFromURL(url, '_blank', 0, (prop('Hidden', True), prop('MacroExecutionMode', 0)))
        return self.doc
    def sheet(self, name): return self.doc.Sheets.getByName(name)
    def close(self):
        try: self.desktop.terminate()
        except Exception: pass
        self.proc.wait(timeout=30)
