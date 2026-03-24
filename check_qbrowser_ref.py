# -*- coding: utf-8 -*-
import os, time, win32com.client
FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")
os.system("taskkill /F /IM MSACCESS.EXE 2>nul"); time.sleep(2)
app = win32com.client.DispatchEx("Access.Application")
app.Visible = False
app.OpenCurrentDatabase(FE, False); time.sleep(2)
app.DoCmd.OpenForm("F_Main", 1); time.sleep(0.3)
frm = app.Forms("F_Main")
comp = app.VBE.VBProjects(1).VBComponents("Form_F_Main")
code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
print("[F_Main VBA - QueryBrowser関連]")
for i, line in enumerate(code.split('\n'), 1):
    if 'QueryBrowser' in line or 'QueryBr' in line or 'qbrowser' in line.lower():
        print(f"  L{i}: {line}")
app.DoCmd.Close(2, "F_Main"); time.sleep(0.2)
app.CloseCurrentDatabase(); app.Quit(); del app
