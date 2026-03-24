# -*- coding: utf-8 -*-
"""STEP 1: 不要オブジェクト削除"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
time.sleep(2)

app = win32com.client.Dispatch("Access.Application")
app.Visible = False
app.UserControl = False
app.OpenCurrentDatabase(BE)
time.sleep(2)

print("[STEP 1] Cleanup")

# Delete junk forms (Japanese names)
for name in ["\u30d5\u30a9\u30fc\u30e01", "\u30d5\u30a9\u30fc\u30e02", "\u30d5\u30a9\u30fc\u30e03", "F_Test"]:
    try:
        app.DoCmd.DeleteObject(2, name)
        print(f"  Deleted form: {name}")
    except:
        pass

# Delete unused queries
for name in ["Q_AllMembers", "Q_Records_List", "Q_Targets_Monthly", "Q_Referrals_Monthly", "Q_List_Summary"]:
    try:
        app.DoCmd.DeleteObject(5, name)  # 5 = acQuery
        print(f"  Deleted query: {name}")
    except:
        pass

time.sleep(0.5)
app.CloseCurrentDatabase()
time.sleep(1)
app.Quit()
time.sleep(1)
print("Done")
