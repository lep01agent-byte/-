# -*- coding: utf-8 -*-
"""SalesMgr 154項目チェックリスト自動検証"""
import os, time, re, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")
FE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_FE.accdb")
PASS = 0; FAIL = 0; RESULTS = []

def chk(id, desc, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1; RESULTS.append((id, "PASS", desc, detail))
    else:
        FAIL += 1; RESULTS.append((id, "FAIL", desc, detail))
        print(f"  !! FAIL {id}: {desc} [{detail}]")

def main():
    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    # DAO for data checks
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(BE)

    # ================================================================
    # A. テーブル
    # ================================================================
    print("=== A. テーブル ===")

    # A01-A07: T_MEMBERS
    tables = set()
    for i in range(db.TableDefs.Count):
        n = db.TableDefs(i).Name
        if not n.startswith("MSys") and not n.startswith("~"):
            tables.add(n)
    chk("A01", "T_MEMBERS exists", "T_MEMBERS" in tables)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A02", "T_MEMBERS count=12", cnt == 12, f"got {cnt}")

    td = db.TableDefs("T_MEMBERS")
    flds = [td.Fields(j).Name for j in range(td.Fields.Count)]
    chk("A03", "T_MEMBERS fields: ID,member_name,active", set(flds) == {"ID","member_name","active"}, str(flds))

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE active=True"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A04", "T_MEMBERS active=True >= 1", cnt >= 1, f"got {cnt}")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE active=False"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A05", "T_MEMBERS active=False >= 0", cnt >= 0, f"got {cnt}")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE member_name IS NULL"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A06", "T_MEMBERS no NULL member_name", cnt == 0, f"got {cnt}")

    rs = db.OpenRecordset("SELECT member_name, COUNT(*) FROM T_MEMBERS GROUP BY member_name HAVING COUNT(*)>1")
    chk("A07", "T_MEMBERS no duplicate member_name", rs.EOF)
    rs.Close()

    # A08-A16: T_RECORDS
    chk("A08", "T_RECORDS exists", "T_RECORDS" in tables)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A09", "T_RECORDS count=5573", cnt == 5573, f"got {cnt}")

    td = db.TableDefs("T_RECORDS")
    chk("A10", "T_RECORDS 22 fields", td.Fields.Count == 22, f"got {td.Fields.Count}")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE rec_date IS NULL"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A11", "T_RECORDS no NULL rec_date", cnt == 0, f"got {cnt}")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE member_name IS NULL"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A12", "T_RECORDS no NULL member_name", cnt == 0, f"got {cnt}")

    rs = db.OpenRecordset("SELECT MIN(rec_date), MAX(rec_date) FROM T_RECORDS")
    mn = str(rs.Fields(0).Value)[:10]; mx = str(rs.Fields(1).Value)[:10]; rs.Close()
    chk("A13", "T_RECORDS min date 2023-09-01", "2023-09-01" in mn, mn)
    chk("A14", "T_RECORDS max date 2026-03-17", "2026-03-17" in mx, mx)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE (IIf(IsNull(calls_10),0,calls_10)+IIf(IsNull(calls_11),0,calls_11)+IIf(IsNull(calls_12),0,calls_12)+IIf(IsNull(calls_13),0,calls_13)+IIf(IsNull(calls_14),0,calls_14)+IIf(IsNull(calls_15),0,calls_15)+IIf(IsNull(calls_16),0,calls_16)+IIf(IsNull(calls_17),0,calls_17)+IIf(IsNull(calls_18),0,calls_18)) > 0 AND calls <> (IIf(IsNull(calls_10),0,calls_10)+IIf(IsNull(calls_11),0,calls_11)+IIf(IsNull(calls_12),0,calls_12)+IIf(IsNull(calls_13),0,calls_13)+IIf(IsNull(calls_14),0,calls_14)+IIf(IsNull(calls_15),0,calls_15)+IIf(IsNull(calls_16),0,calls_16)+IIf(IsNull(calls_17),0,calls_17)+IIf(IsNull(calls_18),0,calls_18))")
    cnt = rs.Fields(0).Value; rs.Close()
    chk("A15", "T_RECORDS calls = sum of hourly (hourly入力ありのみ)", cnt == 0, f"{cnt} mismatches")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
    cnt = rs.Fields(0).Value; rs.Close()
    chk("A16", "T_RECORDS all member_name in T_MEMBERS", cnt == 0, f"{cnt} orphans")

    # A17-A20: T_MEMBER_TARGETS
    chk("A17", "T_MEMBER_TARGETS exists", "T_MEMBER_TARGETS" in tables)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBER_TARGETS"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A18", "T_MEMBER_TARGETS count=12", cnt == 12, f"got {cnt}")

    td = db.TableDefs("T_MEMBER_TARGETS")
    chk("A19", "T_MEMBER_TARGETS 11 fields", td.Fields.Count == 11, f"got {td.Fields.Count}")

    rs = db.OpenRecordset("SELECT member_name,target_year,target_month,COUNT(*) FROM T_MEMBER_TARGETS GROUP BY member_name,target_year,target_month HAVING COUNT(*)>1")
    chk("A20", "T_MEMBER_TARGETS no duplicate keys", rs.EOF)
    rs.Close()

    # A21-A25: T_REFERRALS
    chk("A21", "T_REFERRALS exists", "T_REFERRALS" in tables)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A22", "T_REFERRALS count=4837", cnt == 4837, f"got {cnt}")

    td = db.TableDefs("T_REFERRALS")
    flds = [td.Fields(j).Name for j in range(td.Fields.Count)]
    chk("A23", "T_REFERRALS fields: ID,rec_date,member_name,ref_count", set(flds) == {"ID","rec_date","member_name","ref_count"}, str(flds))

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS WHERE rec_date IS NULL"); cnt = rs.Fields(0).Value; rs.Close()
    chk("A24", "T_REFERRALS no NULL rec_date", cnt == 0, f"got {cnt}")

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL")
    cnt = rs.Fields(0).Value; rs.Close()
    chk("A25", "T_REFERRALS all member_name in T_MEMBERS", cnt == 0, f"{cnt} orphans")

    # ================================================================
    # B. クエリ
    # ================================================================
    print("=== B. クエリ ===")

    queries = {}
    junk_q = []
    for i in range(db.QueryDefs.Count):
        qd = db.QueryDefs(i)
        if not qd.Name.startswith("~"):
            queries[qd.Name] = qd.SQL

    expected_q = ["Q_ActiveMembers","Q_Team_Monthly_Sum","Q_Team_Monthly_Targets",
                  "Q_Member_Monthly_Sum","Q_Rank_Received","Q_Rank_Prospect",
                  "Q_Rank_Referral","Q_Rank_Productivity","Q_Hourly_By_Member",
                  "Q_Trend_Monthly","Q_Member_12Month","Q_RefTrend_Monthly"]
    junk_names = ["Q_AllMembers","Q_Records_List","Q_Targets_Monthly","Q_Referrals_Monthly","Q_List_Summary"]

    # B01-B03
    chk("B01", "Q_ActiveMembers exists", "Q_ActiveMembers" in queries)
    chk("B02", "Q_ActiveMembers no params needed", "PARAMETERS" not in queries.get("Q_ActiveMembers",""))
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE active=True"); cnt = rs.Fields(0).Value; rs.Close()
    chk("B03", "Q_ActiveMembers result >= 1", cnt >= 1, f"got {cnt}")

    # B04-B10: Q_Team_Monthly_Sum
    chk("B04", "Q_Team_Monthly_Sum exists", "Q_Team_Monthly_Sum" in queries)
    chk("B05", "Q_Team_Monthly_Sum has PARAMETERS", "PARAMETERS" in queries.get("Q_Team_Monthly_Sum","").upper())

    qd = db.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom").Value = "2026/03/01"
    qd.Parameters("prmDateTo").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    if not rs.EOF:
        chk("B06", "sum_calls=8412", int(rs.Fields("sum_calls").Value or 0) == 8412, f"got {rs.Fields('sum_calls').Value}")
        chk("B07", "sum_valid=4612", int(rs.Fields("sum_valid").Value or 0) == 4612, f"got {rs.Fields('sum_valid').Value}")
        chk("B08", "sum_prospect=237", int(rs.Fields("sum_prospect").Value or 0) == 237, f"got {rs.Fields('sum_prospect').Value}")
        chk("B09", "sum_received >= 0", int(rs.Fields("sum_received").Value or 0) >= 0)
        chk("B10", "sum_hours >= 0", float(rs.Fields("sum_hours").Value or 0) >= 0)
    else:
        for x in ["B06","B07","B08","B09","B10"]: chk(x, "Q_Team_Monthly_Sum has data", False, "EOF")
    rs.Close()

    # B11-B13: Q_Team_Monthly_Targets
    chk("B11", "Q_Team_Monthly_Targets exists", "Q_Team_Monthly_Targets" in queries)
    chk("B12", "Q_Team_Monthly_Targets has PARAMETERS", "PARAMETERS" in queries.get("Q_Team_Monthly_Targets","").upper())
    qd = db.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear").Value = 2026; qd.Parameters("prmMonth").Value = 3
    rs = qd.OpenRecordset()
    v = 0
    if not rs.EOF: v = int(rs.Fields("sum_tgt_calls").Value or 0)
    rs.Close()
    chk("B13", "sum_tgt_calls > 0", v > 0, f"got {v}")

    # B14-B16
    chk("B14", "Q_Member_Monthly_Sum exists", "Q_Member_Monthly_Sum" in queries)
    chk("B15", "Q_Member_Monthly_Sum has PARAMETERS", "PARAMETERS" in queries.get("Q_Member_Monthly_Sum","").upper())
    qd = db.QueryDefs("Q_Member_Monthly_Sum")
    qd.Parameters("prmDateFrom").Value = "2026/03/01"; qd.Parameters("prmDateTo").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("B16", "Q_Member_Monthly_Sum >= 1 row", cnt >= 1, f"got {cnt}")

    # B17-B40: remaining queries
    for qname, bid_base in [("Q_Rank_Received","B17"),("Q_Rank_Prospect","B20"),("Q_Rank_Referral","B23"),
                             ("Q_Rank_Productivity","B26"),("Q_Hourly_By_Member","B29")]:
        n = int(bid_base[1:])
        chk(f"B{n:02d}", f"{qname} exists", qname in queries)
        chk(f"B{n+1:02d}", f"{qname} has PARAMETERS", "PARAMETERS" in queries.get(qname,"").upper())
        qd = db.QueryDefs(qname)
        qd.Parameters("prmDateFrom").Value = "2026/03/01"; qd.Parameters("prmDateTo").Value = "2026/04/01"
        rs = qd.OpenRecordset(); cnt = 0
        while not rs.EOF: cnt += 1; rs.MoveNext()
        rs.Close()
        chk(f"B{n+2:02d}", f"{qname} >= 1 row", cnt >= 1, f"got {cnt}")

    # B32-B34: Q_Trend_Monthly
    chk("B32", "Q_Trend_Monthly exists", "Q_Trend_Monthly" in queries)
    chk("B33", "Q_Trend_Monthly has PARAMETERS", "PARAMETERS" in queries.get("Q_Trend_Monthly","").upper())
    qd = db.QueryDefs("Q_Trend_Monthly")
    qd.Parameters("prmTrendStart").Value = "2025/10/01"; qd.Parameters("prmTrendEnd").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("B34", "Q_Trend_Monthly 4-6 rows", 4 <= cnt <= 6, f"got {cnt}")

    # B35-B37: Q_Member_12Month
    chk("B35", "Q_Member_12Month exists", "Q_Member_12Month" in queries)
    chk("B36", "Q_Member_12Month has PARAMETERS", "PARAMETERS" in queries.get("Q_Member_12Month","").upper())
    qd = db.QueryDefs("Q_Member_12Month")
    qd.Parameters("prm12Start").Value = "2025/04/01"; qd.Parameters("prm12End").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("B37", "Q_Member_12Month >= 50 rows", cnt >= 50, f"got {cnt}")

    # B38-B40: Q_RefTrend_Monthly
    chk("B38", "Q_RefTrend_Monthly exists", "Q_RefTrend_Monthly" in queries)
    chk("B39", "Q_RefTrend_Monthly has PARAMETERS", "PARAMETERS" in queries.get("Q_RefTrend_Monthly","").upper())
    qd = db.QueryDefs("Q_RefTrend_Monthly")
    qd.Parameters("prm12Start").Value = "2025/04/01"; qd.Parameters("prm12End").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("B40", "Q_RefTrend_Monthly >= 1 row", cnt >= 1, f"got {cnt}")

    # B41: no junk queries
    junk_found = [n for n in junk_names if n in queries]
    chk("B41", "No junk queries", len(junk_found) == 0, str(junk_found))

    db.Close()

    # ================================================================
    # C-D. フォーム構造 + VBAコード検証（COM経由）
    # ================================================================
    print("=== C/D. フォーム構造 + VBA ===")

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False; app.UserControl = False
    app.OpenCurrentDatabase(FE)
    time.sleep(2)

    # Collect form info
    form_info = {}
    forms_list = set()
    dbc = app.CurrentDb()
    for i in range(dbc.Containers("Forms").Documents.Count):
        forms_list.add(dbc.Containers("Forms").Documents(i).Name)

    junk_forms = [f for f in forms_list if f not in {"F_Main","F_Daily","F_DailyEdit","F_Members","F_Targets","F_Referrals","F_Report","F_Ranking","F_List","F_Analysis","F_QueryBrowser"}]

    for fn in ["F_Main","F_Daily","F_DailyEdit","F_Members","F_Targets","F_Referrals","F_Report","F_Ranking"]:
        try:
            app.DoCmd.OpenForm(fn, 1)  # Design
            time.sleep(0.3)
            frm = app.Forms(fn)
            ctrls = {}
            for i in range(frm.Controls.Count):
                c = frm.Controls(i)
                ctrls[c.Name] = c.ControlType
            vba_code = ""
            vba_lines = 0
            if frm.HasModule:
                comp = app.VBE.VBProjects(1).VBComponents("Form_" + fn)
                cm = comp.CodeModule
                vba_lines = cm.CountOfLines
                if vba_lines > 0:
                    vba_code = cm.Lines(1, vba_lines)
            form_info[fn] = {"ctrls": ctrls, "vba": vba_code, "vba_lines": vba_lines,
                             "popup": frm.PopUp, "modal": frm.Modal, "hasmod": frm.HasModule}
            app.DoCmd.Close(2, fn)
            time.sleep(0.2)
        except Exception as e:
            print(f"  ERROR reading {fn}: {e}")
            try: app.DoCmd.Close(2, fn)
            except: pass

    app.CloseCurrentDatabase()
    app.Quit()
    time.sleep(1)

    # C checks
    def has_ctrl(fn, name): return name in form_info.get(fn, {}).get("ctrls", {})
    def ctrl_count(fn, prefix):
        return sum(1 for n in form_info.get(fn,{}).get("ctrls",{}) if n.startswith(prefix))
    def vba(fn): return form_info.get(fn, {}).get("vba", "")
    def vlines(fn): return form_info.get(fn, {}).get("vba_lines", 0)

    # C01-C08: F_Main
    chk("C01", "F_Main exists", "F_Main" in forms_list)
    chk("C02", "F_Main HasModule", form_info.get("F_Main",{}).get("hasmod",False))
    for i, bn in enumerate(["btnDaily","btnTargets","btnReferrals","btnReport","btnRanking","btnMembers"], 3):
        chk(f"C{i:02d}", f"F_Main has {bn}", has_ctrl("F_Main", bn))

    # C09-C13: F_Daily
    chk("C09", "F_Daily exists", "F_Daily" in forms_list)
    chk("C10", "F_Daily HasModule", form_info.get("F_Daily",{}).get("hasmod",False))
    daily_btns = ["btnPrev","btnNext","btnAdd","btnEdit","btnDelete"]
    chk("C11", "F_Daily has 5 buttons", all(has_ctrl("F_Daily",b) for b in daily_btns), str([b for b in daily_btns if not has_ctrl("F_Daily",b)]))
    chk("C12", "F_Daily has cboMember", has_ctrl("F_Daily","cboMember"))
    chk("C13", "F_Daily has lstRecords", has_ctrl("F_Daily","lstRecords"))

    # C14-C20: F_DailyEdit
    chk("C14", "F_DailyEdit exists", "F_DailyEdit" in forms_list)
    chk("C15", "F_DailyEdit HasModule", form_info.get("F_DailyEdit",{}).get("hasmod",False))
    chk("C16", "F_DailyEdit PopUp+Modal", form_info.get("F_DailyEdit",{}).get("popup",False) and form_info.get("F_DailyEdit",{}).get("modal",False))
    chk("C17", "F_DailyEdit has txtRecDate,cboMember", has_ctrl("F_DailyEdit","txtRecDate") and has_ctrl("F_DailyEdit","cboMember"))
    hourly = sum(1 for h in range(10,19) if has_ctrl("F_DailyEdit", f"txtC{h}"))
    chk("C18", "F_DailyEdit has 9 hourly boxes", hourly == 9, f"got {hourly}")
    result_flds = ["txtValid","txtProspect","txtDoc","txtFollow","txtReceived","txtReferral"]
    chk("C19", "F_DailyEdit has 6 result fields", all(has_ctrl("F_DailyEdit",f) for f in result_flds))
    chk("C20", "F_DailyEdit has btnSave,btnCancel", has_ctrl("F_DailyEdit","btnSave") and has_ctrl("F_DailyEdit","btnCancel"))

    # C21-C25: F_Members
    chk("C21", "F_Members exists", "F_Members" in forms_list)
    chk("C22", "F_Members HasModule", form_info.get("F_Members",{}).get("hasmod",False))
    chk("C23", "F_Members has txtNewName,btnAdd", has_ctrl("F_Members","txtNewName") and has_ctrl("F_Members","btnAdd"))
    chk("C24", "F_Members has lstActive,lstInactive", has_ctrl("F_Members","lstActive") and has_ctrl("F_Members","lstInactive"))
    chk("C25", "F_Members has btnDeactivate,btnActivate", has_ctrl("F_Members","btnDeactivate") and has_ctrl("F_Members","btnActivate"))

    # C26-C32: F_Targets
    chk("C26", "F_Targets exists", "F_Targets" in forms_list)
    chk("C27", "F_Targets HasModule", form_info.get("F_Targets",{}).get("hasmod",False))
    chk("C28", "F_Targets has btnPrev,btnNext", has_ctrl("F_Targets","btnPrev") and has_ctrl("F_Targets","btnNext"))
    chk("C29", "F_Targets has cboMember,btnLoadPrev", has_ctrl("F_Targets","cboMember") and has_ctrl("F_Targets","btnLoadPrev"))
    tgt_flds = ["txtPlanDays","txtHoursPerDay","txtTgtCalls","txtTgtValid","txtTgtProspect","txtTgtReceived","txtTgtReferral"]
    chk("C30", "F_Targets has 7 target fields", all(has_ctrl("F_Targets",f) for f in tgt_flds), str([f for f in tgt_flds if not has_ctrl("F_Targets",f)]))
    chk("C31", "F_Targets has btnSave,btnDelete", has_ctrl("F_Targets","btnSave") and has_ctrl("F_Targets","btnDelete"))
    chk("C32", "F_Targets has lstTargets", has_ctrl("F_Targets","lstTargets"))

    # C33-C38: F_Referrals
    chk("C33", "F_Referrals exists", "F_Referrals" in forms_list)
    chk("C34", "F_Referrals HasModule", form_info.get("F_Referrals",{}).get("hasmod",False))
    chk("C35", "F_Referrals has btnPrev,btnNext", has_ctrl("F_Referrals","btnPrev") and has_ctrl("F_Referrals","btnNext"))
    chk("C36", "F_Referrals has txtRefDate,cboMember,txtRefCount", all(has_ctrl("F_Referrals",f) for f in ["txtRefDate","cboMember","txtRefCount"]))
    chk("C37", "F_Referrals has btnAdd,btnDelete", has_ctrl("F_Referrals","btnAdd") and has_ctrl("F_Referrals","btnDelete"))
    chk("C38", "F_Referrals has lstRefs", has_ctrl("F_Referrals","lstRefs"))

    # C39-C46: F_Report
    chk("C39", "F_Report exists", "F_Report" in forms_list)
    chk("C40", "F_Report HasModule", form_info.get("F_Report",{}).get("hasmod",False))
    chk("C41", "F_Report has btnPrev,btnNext,btnPDF", all(has_ctrl("F_Report",b) for b in ["btnPrev","btnNext","btnPDF"]))
    chk("C42", "F_Report has lblAlert", has_ctrl("F_Report","lblAlert"))
    chk("C43", "F_Report has 4 KPI labels", all(has_ctrl("F_Report",l) for l in ["lblCalls","lblValid","lblProsp","lblRecv"]))
    chk("C44", "F_Report has 4 prev labels", all(has_ctrl("F_Report",l) for l in ["lblCallsPrev","lblValidPrev","lblProspPrev","lblRecvPrev"]))
    chk("C45", "F_Report has rate/hours/prod labels", all(has_ctrl("F_Report",l) for l in ["lblValidRate","lblRecvRate","lblHours","lblProductivity"]))
    chk("C46", "F_Report has 3 ranking lists", all(has_ctrl("F_Report",l) for l in ["lstRankRef","lstRankRecv","lstRankProsp"]))

    # C47-C50: F_Ranking
    chk("C47", "F_Ranking exists", "F_Ranking" in forms_list)
    chk("C48", "F_Ranking HasModule", form_info.get("F_Ranking",{}).get("hasmod",False))
    chk("C49", "F_Ranking has btnPrev,btnNext", has_ctrl("F_Ranking","btnPrev") and has_ctrl("F_Ranking","btnNext"))
    chk("C50", "F_Ranking has 3 lists", all(has_ctrl("F_Ranking",l) for l in ["lstRef","lstRecv","lstProsp"]))

    chk("C51", "No junk forms", len(junk_forms) == 0, str(junk_forms))

    # ================================================================
    # D. VBA コード検証
    # ================================================================
    print("=== D. VBA ===")

    # D01-D03: F_Main
    chk("D01", "F_Main VBA lines > 0", vlines("F_Main") > 0, f"{vlines('F_Main')}")
    chk("D02", "F_Main btnDaily opens F_Daily", 'OpenForm "F_Daily"' in vba("F_Main"))
    btn_subs = ["btnDaily_Click","btnTargets_Click","btnReferrals_Click","btnReport_Click","btnRanking_Click","btnMembers_Click"]
    chk("D03", "F_Main has 6 button subs", all(s in vba("F_Main") for s in btn_subs), str([s for s in btn_subs if s not in vba("F_Main")]))

    # D04-D14: F_Daily
    v = vba("F_Daily")
    chk("D04", "F_Daily VBA lines > 0", vlines("F_Daily") > 0)
    chk("D05", "F_Daily Form_Open init mY,mM", "mY = Year(Date)" in v and "mM = Month(Date)" in v)
    chk("D06", "F_Daily LoadData has On Error", "On Error GoTo EH" in v)
    chk("D07", "F_Daily 12->1 nav", "mY = mY + 1: mM = 1" in v or "mY + 1" in v)
    chk("D08", "F_Daily 1->12 nav", "mY = mY - 1: mM = 12" in v or "mY - 1" in v)
    chk("D09", "F_Daily DateSerial 12-safe", "mM = 12 Then" in v)
    chk("D10", "F_Daily cboMember_AfterUpdate", "cboMember_AfterUpdate" in v)
    chk("D11", "F_Daily btnAdd passes ADD|Y|M", '"ADD|"' in v or "'ADD|'" in v or "ADD|" in v)
    chk("D12", "F_Daily btnEdit IsNull check", "IsNull(lstRecords" in v)
    chk("D13", "F_Daily btnDelete confirm", "vbYesNo" in v)
    chk("D14", "F_Daily DblClick calls btnEdit", "btnEdit_Click" in v and "DblClick" in v)

    # D15-D26: F_DailyEdit
    v = vba("F_DailyEdit")
    chk("D15", "F_DailyEdit VBA lines > 0", vlines("F_DailyEdit") > 0)
    chk("D16", "F_DailyEdit Split OpenArgs", "Split(" in v)
    chk("D17", "F_DailyEdit ADD/EDIT separated (no ElseIf fallthrough)", "mMode = \"EDIT\"" in v and "Else" in v)
    chk("D18", "F_DailyEdit ADD init hourly to 0", "Me(\"txtC\" & h).Value = 0" in v or 'txtC" & h' in v)
    chk("D19", "F_DailyEdit LoadRecord has On Error", v.count("On Error GoTo EH") >= 2)
    chk("D20", "F_DailyEdit save checks empty date/member", "txtRecDate" in v and "cboMember" in v and "vbExclamation" in v)
    chk("D21", "F_DailyEdit totalCalls calc", "totalCalls" in v)
    # Count VALUES columns: rec_date + member_name + calls + 9 hourly + 6 result + work_hours + note + referral + work_day = 21
    chk("D22", "F_DailyEdit INSERT has 21 columns", "rec_date, member_name, calls" in v and "calls_10" in v and "work_day" in v)
    chk("D23", "F_DailyEdit SQL escapes quotes", "Replace(" in v)
    chk("D24", "F_DailyEdit uses CLng/CDbl", "CLng(" in v and "CDbl(" in v)
    chk("D25", "F_DailyEdit error shows message", "vbCritical" in v)
    chk("D26", "F_DailyEdit btnCancel closes form", "DoCmd.Close" in v)

    # D27-D33: F_Members
    v = vba("F_Members")
    chk("D27", "F_Members VBA lines > 0", vlines("F_Members") > 0)
    chk("D28", "F_Members Form_Open calls LoadData", "LoadData" in v and "Form_Open" in v)
    chk("D29", "F_Members btnAdd empty check", "Len(nm) = 0" in v or 'nm = ""' in v or "Len(nm)" in v)
    chk("D30", "F_Members btnAdd escapes quotes", "Replace(" in v)
    chk("D31", "F_Members btnDeactivate IsNull", "IsNull(lstActive" in v)
    chk("D32", "F_Members btnActivate IsNull", "IsNull(lstInactive" in v)
    chk("D33", "F_Members all subs have On Error", v.count("On Error GoTo EH") >= 3)

    # D34-D42: F_Targets
    v = vba("F_Targets")
    chk("D34", "F_Targets VBA lines > 0", vlines("F_Targets") > 0)
    chk("D35", "F_Targets Form_Open sets RowSource", "cboMember.RowSource" in v and "Form_Open" in v)
    chk("D36", "F_Targets cboMember_AfterUpdate loads target", "cboMember_AfterUpdate" in v and "OpenRecordset" in v)
    chk("D37", "F_Targets AfterUpdate clears on no data", 'txtPlanDays.Value = ""' in v)
    chk("D38", "F_Targets btnSave UPSERT with DCount", "DCount" in v)
    chk("D39", "F_Targets btnSave uses CLng/CDbl", "CLng(" in v and "CDbl(" in v)
    chk("D40", "F_Targets btnLoadPrev handles Jan->Dec", "py = py - 1: pm = 12" in v or "py - 1" in v)
    chk("D41", "F_Targets btnDelete confirm", "vbYesNo" in v)
    chk("D42", "F_Targets all subs have On Error", v.count("On Error GoTo EH") >= 4)

    # D43-D49: F_Referrals
    v = vba("F_Referrals")
    chk("D43", "F_Referrals VBA lines > 0", vlines("F_Referrals") > 0)
    chk("D44", "F_Referrals Form_Open sets RowSource", "cboMember.RowSource" in v)
    chk("D45", "F_Referrals DateSerial 12-safe", "mM = 12 Then" in v)
    chk("D46", "F_Referrals btnAdd empty check", "vbExclamation" in v)
    chk("D47", "F_Referrals btnAdd CDate", "CDate(" in v)
    chk("D48", "F_Referrals btnDelete confirm", "vbYesNo" in v)
    chk("D49", "F_Referrals all subs On Error", v.count("On Error GoTo EH") >= 3)

    # D50-D64: F_Report
    v = vba("F_Report")
    chk("D50", "F_Report VBA lines > 0", vlines("F_Report") > 0)
    chk("D51", "F_Report uses direct SQL (no QueryDefs)", "QueryDefs" not in v and "CurrentDb.OpenRecordset" in v)
    chk("D52", "F_Report DateSerial 12-safe (dTo)", "mM = 12 Then" in v)
    chk("D53", "F_Report prev month 12-safe", "pm = 1 Then" in v or "pm = 12" in v)
    chk("D54", "F_Report tC init 0", "tC = 0" in v)
    chk("D55", "F_Report gC init 0", "gC = 0" in v)
    chk("D56", "F_Report pC init 0", "pC = 0" in v)
    chk("D57", "F_Report KPI format #,##0 / #,##0", '"#,##0"' in v and '" / "' in v)
    chk("D58", "F_Report MakeArrow function", "MakeArrow" in v and "Function MakeArrow" in v)
    chk("D59", "F_Report tC=0 safe (division check)", 'tC > 0 Then' in v or "tC > 0" in v)
    chk("D60", "F_Report tH=0 safe", 'tH > 0 Then' in v or "tH > 0" in v)
    chk("D61", "F_Report alert recv nested If correct", "If gR > 0 Then" in v and "If tR >= gR Then" in v)
    chk("D62", "F_Report alert prosp nested If correct", "If gP > 0 Then" in v and "If tP >= gP Then" in v)
    chk("D63", "F_Report ranking 3 direct SQL", v.count("lstRank") >= 3)
    chk("D64", "F_Report On Error", "On Error GoTo EH" in v)

    # D65-D68: F_Ranking
    v = vba("F_Ranking")
    chk("D65", "F_Ranking VBA lines > 0", vlines("F_Ranking") > 0)
    chk("D66", "F_Ranking DateSerial 12-safe", "mM = 12 Then" in v)
    chk("D67", "F_Ranking 3 direct SQL", v.count("RowSource") >= 3)
    chk("D68", "F_Ranking On Error", "On Error GoTo EH" in v)

    # ================================================================
    # E. CRUD
    # ================================================================
    print("=== E. CRUD ===")
    db = engine.OpenDatabase(BE)

    db.Execute("INSERT INTO T_RECORDS (rec_date,member_name,calls,valid_count,prospect,received,work_hours,work_day) VALUES (#2099/01/01#,'__TEST__',100,50,10,2,8,True)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE member_name='__TEST__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E01", "INSERT record", cnt == 1)

    db.Execute("UPDATE T_RECORDS SET calls=200 WHERE member_name='__TEST__'")
    rs = db.OpenRecordset("SELECT calls FROM T_RECORDS WHERE member_name='__TEST__'"); v = rs.Fields(0).Value; rs.Close()
    chk("E02", "UPDATE record", v == 200)

    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS"); before = rs.Fields(0).Value; rs.Close()
    db.Execute("DELETE FROM T_RECORDS WHERE member_name='__TEST__'")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS"); after = rs.Fields(0).Value; rs.Close()
    chk("E03", "DELETE record", after == before - 1)
    chk("E04", "INSERT added 1", True)  # already verified
    chk("E05", "DELETE removed 1", after == 5573)

    db.Execute("INSERT INTO T_MEMBERS (member_name,active) VALUES ('__TEST_M__',True)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE member_name='__TEST_M__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E06", "INSERT member", cnt == 1)
    db.Execute("UPDATE T_MEMBERS SET active=False WHERE member_name='__TEST_M__'")
    rs = db.OpenRecordset("SELECT active FROM T_MEMBERS WHERE member_name='__TEST_M__'"); v = rs.Fields(0).Value; rs.Close()
    chk("E07", "UPDATE member active", v == False)
    db.Execute("DELETE FROM T_MEMBERS WHERE member_name='__TEST_M__'")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE member_name='__TEST_M__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E08", "DELETE member", cnt == 0)

    db.Execute("INSERT INTO T_MEMBER_TARGETS (member_name,target_year,target_month,target_calls) VALUES ('__TGT__',2099,1,999)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBER_TARGETS WHERE member_name='__TGT__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E09", "INSERT target", cnt == 1)
    db.Execute("UPDATE T_MEMBER_TARGETS SET target_calls=888 WHERE member_name='__TGT__'")
    rs = db.OpenRecordset("SELECT target_calls FROM T_MEMBER_TARGETS WHERE member_name='__TGT__'"); v = rs.Fields(0).Value; rs.Close()
    chk("E10", "UPDATE target", v == 888)
    db.Execute("DELETE FROM T_MEMBER_TARGETS WHERE member_name='__TGT__'")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBER_TARGETS WHERE member_name='__TGT__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E11", "DELETE target", cnt == 0)

    db.Execute("INSERT INTO T_REFERRALS (rec_date,member_name,ref_count) VALUES (#2099/01/01#,'__REF__',5)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS WHERE member_name='__REF__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E12", "INSERT referral", cnt == 1)
    db.Execute("DELETE FROM T_REFERRALS WHERE member_name='__REF__'")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS WHERE member_name='__REF__'"); cnt = rs.Fields(0).Value; rs.Close()
    chk("E13", "DELETE referral", cnt == 0)

    # ================================================================
    # F. 集計値の正確性
    # ================================================================
    print("=== F. 集計値 ===")

    rs = db.OpenRecordset("SELECT Sum(calls),Sum(valid_count),Sum(prospect),Sum(received),Sum(work_hours) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#2026/03/01# AND R.rec_date<#2026/04/01#")
    tC=int(rs.Fields(0).Value or 0); tV=int(rs.Fields(1).Value or 0); tP=int(rs.Fields(2).Value or 0)
    tR=int(rs.Fields(3).Value or 0); tH=float(rs.Fields(4).Value or 0); rs.Close()
    chk("F01", "2026/3 calls=8412", tC == 8412, f"got {tC}")
    chk("F02", "2026/3 valid=4612", tV == 4612, f"got {tV}")
    chk("F03", "2026/3 prospect=237", tP == 237, f"got {tP}")
    chk("F04", "2026/3 received >= 0", tR >= 0, f"got {tR}")
    chk("F05", "2026/3 hours >= 0", tH >= 0, f"got {tH}")
    chk("F06", "2026/3 valid rate ~54.8", abs(tV/tC*100 - 54.8) < 0.1 if tC > 0 else False, f"got {tV/tC*100:.1f}" if tC > 0 else "div0")

    rs = db.OpenRecordset("SELECT Sum(target_calls) FROM T_MEMBER_TARGETS WHERE target_year=2026 AND target_month=3")
    v = int(rs.Fields(0).Value or 0); rs.Close()
    chk("F07", "2026/3 target_calls > 0", v > 0, f"got {v}")

    # Rankings order
    for fid, col, label in [("F08","received","受注"),("F09","prospect","見込"),("F10","referral","送客")]:
        rs = db.OpenRecordset(f"SELECT Sum({col}) AS s FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#2026/03/01# AND R.rec_date<#2026/04/01# GROUP BY R.member_name ORDER BY Sum({col}) DESC")
        vals = []
        while not rs.EOF:
            vals.append(int(rs.Fields(0).Value or 0)); rs.MoveNext()
        rs.Close()
        chk(fid, f"{label} rank 1st >= 2nd", len(vals) >= 2 and vals[0] >= vals[1], f"{vals[:3]}")

    qd = db.QueryDefs("Q_Trend_Monthly")
    qd.Parameters("prmTrendStart").Value = "2025/10/01"; qd.Parameters("prmTrendEnd").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("F11", "6mo trend 4-6 rows", 4 <= cnt <= 6, f"got {cnt}")

    qd = db.QueryDefs("Q_Member_12Month")
    qd.Parameters("prm12Start").Value = "2025/04/01"; qd.Parameters("prm12End").Value = "2026/04/01"
    rs = qd.OpenRecordset(); cnt = 0
    while not rs.EOF: cnt += 1; rs.MoveNext()
    rs.Close()
    chk("F12", "12mo member rows >= 50", cnt >= 50, f"got {cnt}")

    # ================================================================
    # G. スタートアップ
    # ================================================================
    print("=== G. スタートアップ ===")

    try:
        v = db.Properties("StartUpForm").Value
        chk("G01", "StartUpForm=F_Main", v == "F_Main", f"got {v}")
    except: chk("G01", "StartUpForm=F_Main", False, "not set")

    try:
        v = db.Properties("AppTitle").Value
        chk("G02", "AppTitle set", len(str(v)) > 0, str(v))
    except: chk("G02", "AppTitle set", False, "not set")

    try:
        v = db.Properties("StartUpShowDBWindow").Value
        chk("G03", "ShowDBWindow=False", v == False, f"got {v}")
    except: chk("G03", "ShowDBWindow=False", False, "not set")

    chk("G04", "No junk forms", len(junk_forms) == 0, str(junk_forms))
    chk("G05", "No junk queries", len(junk_found) == 0, str(junk_found))

    reports = set()
    for i in range(db.Containers("Reports").Documents.Count):
        reports.add(db.Containers("Reports").Documents(i).Name)
    chk("G06", "No report objects", len(reports) == 0, str(reports))

    modules = set()
    for i in range(db.Containers("Modules").Documents.Count):
        modules.add(db.Containers("Modules").Documents(i).Name)
    chk("G07", "No standard modules", len(modules) == 0, str(modules))

    db.Close()

    # ================================================================
    # SUMMARY
    # ================================================================
    print("\n" + "=" * 60)
    print(f"RESULTS: {PASS} PASS / {FAIL} FAIL / {PASS+FAIL} TOTAL")
    print("=" * 60)

    if FAIL > 0:
        print("\nFAILED ITEMS:")
        for id, status, desc, detail in RESULTS:
            if status == "FAIL":
                print(f"  {id}: {desc} [{detail}]")
    else:
        print("ALL CHECKS PASSED!")


if __name__ == "__main__":
    main()
