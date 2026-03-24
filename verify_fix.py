# -*- coding: utf-8 -*-
import os, win32com.client
FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(BE)
rs = db.OpenRecordset("SELECT member_name, active FROM T_MEMBERS ORDER BY member_name")
print("T_MEMBERS:")
while not rs.EOF:
    print(f"  {rs.Fields('member_name').Value}: {rs.Fields('active').Value}")
    rs.MoveNext()
rs.Close()
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
print(f"T_RECORDS orphans: {rs.Fields(0).Value}"); rs.Close()
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
print(f"T_REFERRALS orphans: {rs.Fields(0).Value}"); rs.Close()
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS")
print(f"T_MEMBERS count: {rs.Fields(0).Value}"); rs.Close()
db.Close()
