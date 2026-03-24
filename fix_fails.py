# -*- coding: utf-8 -*-
"""残りFAIL修正: 新見追加 + junkクエリ削除"""
import os, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")

os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
time.sleep(2)

engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(BE)

# ── 1. T_MEMBERSに「新見」追加 (active=False) ──────────────
print("[1] T_MEMBERSに新見追加")
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE member_name='新見'")
exists = rs.Fields(0).Value > 0
rs.Close()
if exists:
    print("  既に存在 → スキップ")
else:
    db.Execute("INSERT INTO T_MEMBERS (member_name, active) VALUES ('新見', False)")
    print("  新見 (active=False) 追加完了")

# 確認
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
cnt = rs.Fields(0).Value; rs.Close()
print(f"  T_RECORDS orphans残: {cnt}件")

rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
cnt = rs.Fields(0).Value; rs.Close()
print(f"  T_REFERRALS orphans残: {cnt}件")

# ── 2. junkクエリ削除 ─────────────────────────────────────
print("\n[2] junkクエリ削除")
JUNK = ["Q_AllMembers","Q_Records_List","Q_Targets_Monthly","Q_Referrals_Monthly","Q_List_Summary"]
for q in JUNK:
    try:
        db.QueryDefs.Delete(q)
        print(f"  削除: {q}")
    except Exception as e:
        print(f"  スキップ {q}: {e}")

db.Close()

# ── 3. FEからもjunkクエリ削除 ────────────────────────────
print("\n[3] FEからjunkクエリ削除")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")
db2 = engine.OpenDatabase(FE)
for q in JUNK:
    try:
        db2.QueryDefs.Delete(q)
        print(f"  FE削除: {q}")
    except Exception as e:
        print(f"  FEスキップ {q}: {e}")
db2.Close()

print("\n完了")
