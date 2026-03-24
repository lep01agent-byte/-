# -*- coding: utf-8 -*-
"""STEP 2: VBA v3 注入（cleanup_and_rebuild.pyのget_all_vba()を使用）"""
import os, sys, time
sys.path.insert(0, os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr"))
from cleanup_and_rebuild import get_all_vba
import win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
time.sleep(2)

app = win32com.client.Dispatch("Access.Application")
app.Visible = False
app.UserControl = False
app.OpenCurrentDatabase(BE)
time.sleep(2)

print("[STEP 2] VBA v3 inject")
FORM_VBA = get_all_vba()

for form_name, code in FORM_VBA.items():
    try:
        app.DoCmd.OpenForm(form_name, 1)  # Design
        time.sleep(0.4)
        frm = app.Forms(form_name)
        frm.HasModule = True
        time.sleep(0.2)

        comp = app.VBE.VBProjects(1).VBComponents("Form_" + form_name)
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)

        lines = code.strip().split("\n")
        for i, line in enumerate(lines, 1):
            cm.InsertLines(i, line)

        app.DoCmd.Save(2, form_name)
        app.DoCmd.Close(2, form_name)
        time.sleep(0.3)
        print(f"  {form_name}: {len(lines)} lines OK")
    except Exception as e:
        print(f"  {form_name}: ERROR {e}")
        try: app.DoCmd.Close(2, form_name)
        except: pass

time.sleep(1)
app.CloseCurrentDatabase()
time.sleep(1)
app.Quit()
time.sleep(1)
print("Done")
