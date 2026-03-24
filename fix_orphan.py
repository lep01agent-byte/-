# -*- coding: utf-8 -*-
import os, win32com.client
FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(BE)

# T_RECORDSの孤立member_nameを正確に取得
rs = db.OpenRecordset("SELECT DISTINCT R.member_name FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
orphan_names = []
while not rs.EOF:
    orphan_names.append(rs.Fields('member_name').Value)
    rs.MoveNext()
rs.Close()

print(f"orphan names: {orphan_names}")

# 既存の誤挿入を削除
rs2 = db.OpenRecordset("SELECT member_name FROM T_MEMBERS WHERE active=False")
inactive = []
while not rs2.EOF:
    inactive.append(rs2.Fields('member_name').Value)
    rs2.MoveNext()
rs2.Close()
print(f"inactive members: {inactive}")

for nm in inactive:
    db.Execute(f"DELETE FROM T_MEMBERS WHERE member_name='{nm}' AND active=False")
    print(f"  削除: {repr(nm)}")

# orphan namesをactive=Falseで追加
for nm in orphan_names:
    safe = nm.replace("'", "''")
    db.Execute(f"INSERT INTO T_MEMBERS (member_name, active) VALUES ('{safe}', False)")
    print(f"  追加: {repr(nm)}")

# 確認
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
print(f"T_RECORDS orphans: {rs.Fields(0).Value}"); rs.Close()
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
print(f"T_REFERRALS orphans: {rs.Fields(0).Value}"); rs.Close()
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS")
print(f"T_MEMBERS count: {rs.Fields(0).Value}"); rs.Close()

db.Close()
