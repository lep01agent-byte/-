# -*- coding: utf-8 -*-
"""診断スクリプト: Accessがどこで止まるかを確認"""
import os, time, win32com.client, sys

FE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_FE.accdb")

print("Step 1: DispatchEx...", flush=True)
app = win32com.client.DispatchEx('Access.Application')
app.Visible = True   # ダイアログが見えるよう True にする
app.UserControl = False
print("Step 2: OpenCurrentDatabase...", flush=True)
app.OpenCurrentDatabase(FE)
print("Step 3: 待機...", flush=True)
time.sleep(5)
print("Step 4: F_Main デザインビューで開く...", flush=True)
try:
    app.DoCmd.OpenForm('F_Main', 1)  # acDesign = 1
    time.sleep(2)
    print("Step 5: フォーム取得...", flush=True)
    frm = app.Forms('F_Main')
    print(f"  F_Main 取得成功。コントロール数: {frm.Controls.Count}", flush=True)
    for ctrl in frm.Controls:
        print(f"    [{ctrl.ControlType}] {ctrl.Name}", flush=True)
    app.DoCmd.Close(2, 'F_Main')
    print("Step 6: フォームを閉じた", flush=True)
except Exception as e:
    print(f"ERROR: {e}", flush=True)

print("Step 7: Quit...", flush=True)
try:
    app.CloseCurrentDatabase()
    app.Quit()
except: pass
print("DONE", flush=True)
