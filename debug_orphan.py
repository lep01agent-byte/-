# -*- coding: utf-8 -*-
import os, win32com.client
FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(BE)

# 孤立member_namesを直接確認
rs = db.OpenRecordset("SELECT DISTINCT R.member_name FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
print("T_RECORDS orphan member_names (raw bytes):")
while not rs.EOF:
    v = rs.Fields('member_name').Value
    print(f"  repr={repr(v)}  bytes={v.encode('utf-8').hex()}")
    rs.MoveNext()
rs.Close()

# T_MEMBERSの新見の確認
rs = db.OpenRecordset("SELECT member_name FROM T_MEMBERS WHERE active=False")
print("\nT_MEMBERS inactive members (raw bytes):")
while not rs.EOF:
    v = rs.Fields('member_name').Value
    print(f"  repr={repr(v)}  bytes={v.encode('utf-8').hex()}")
    rs.MoveNext()
rs.Close()

db.Close()
