# -*- coding: utf-8 -*-
"""SalesMgr Access テスト（4カテゴリ）"""
import os, time, win32com.client

BE_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")
PASS = 0
FAIL = 0

def check(desc, cond):
    global PASS, FAIL
    if cond:
        print(f"  [PASS] {desc}")
        PASS += 1
    else:
        print(f"  [FAIL] {desc}")
        FAIL += 1

def main():
    global PASS, FAIL
    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(BE_PATH)

    # ========================================
    print("=" * 50)
    print("TEST 1: Data Integrity")
    print("=" * 50)

    # T_MEMBERS count
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS")
    cnt = rs.Fields(0).Value; rs.Close()
    check("T_MEMBERS count = 11", cnt == 11)

    # T_RECORDS count
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS")
    cnt = rs.Fields(0).Value; rs.Close()
    check("T_RECORDS count = 5573", cnt == 5573)

    # T_MEMBER_TARGETS count
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBER_TARGETS")
    cnt = rs.Fields(0).Value; rs.Close()
    check("T_MEMBER_TARGETS count = 12", cnt == 12)

    # T_REFERRALS count
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS")
    cnt = rs.Fields(0).Value; rs.Close()
    check("T_REFERRALS count = 4837", cnt == 4837)

    # Active members
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE active=True")
    cnt = rs.Fields(0).Value; rs.Close()
    check("Active members > 0", cnt > 0)

    # Date range
    rs = db.OpenRecordset("SELECT MIN(rec_date), MAX(rec_date) FROM T_RECORDS")
    mn = rs.Fields(0).Value; mx = rs.Fields(1).Value; rs.Close()
    check(f"Records date range valid ({mn} to {mx})", mn is not None and mx is not None)

    # ========================================
    print("\n" + "=" * 50)
    print("TEST 2: CRUD Operations")
    print("=" * 50)

    # INSERT test record
    db.Execute("INSERT INTO T_RECORDS (rec_date, member_name, calls, valid_count, prospect, received, work_hours, work_day) VALUES (#2099/01/01#, 'TEST_USER', 100, 50, 10, 2, 8, True)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE member_name='TEST_USER'")
    cnt = rs.Fields(0).Value; rs.Close()
    check("INSERT record works", cnt == 1)

    # UPDATE test record
    db.Execute("UPDATE T_RECORDS SET calls=200 WHERE member_name='TEST_USER'")
    rs = db.OpenRecordset("SELECT calls FROM T_RECORDS WHERE member_name='TEST_USER'")
    val = rs.Fields(0).Value; rs.Close()
    check("UPDATE record works", val == 200)

    # DELETE test record
    db.Execute("DELETE FROM T_RECORDS WHERE member_name='TEST_USER'")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE member_name='TEST_USER'")
    cnt = rs.Fields(0).Value; rs.Close()
    check("DELETE record works", cnt == 0)

    # INSERT/DELETE member
    db.Execute("INSERT INTO T_MEMBERS (member_name, active) VALUES ('TEST_MEMBER', True)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBERS WHERE member_name='TEST_MEMBER'")
    cnt = rs.Fields(0).Value; rs.Close()
    check("INSERT member works", cnt == 1)
    db.Execute("DELETE FROM T_MEMBERS WHERE member_name='TEST_MEMBER'")

    # INSERT/DELETE target
    db.Execute("INSERT INTO T_MEMBER_TARGETS (member_name, target_year, target_month, target_calls) VALUES ('TEST_TGT', 2099, 1, 1000)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_MEMBER_TARGETS WHERE member_name='TEST_TGT'")
    cnt = rs.Fields(0).Value; rs.Close()
    check("INSERT target works", cnt == 1)
    db.Execute("DELETE FROM T_MEMBER_TARGETS WHERE member_name='TEST_TGT'")

    # INSERT/DELETE referral
    db.Execute("INSERT INTO T_REFERRALS (rec_date, member_name, ref_count) VALUES (#2099/01/01#, 'TEST_REF', 5)")
    rs = db.OpenRecordset("SELECT COUNT(*) FROM T_REFERRALS WHERE member_name='TEST_REF'")
    cnt = rs.Fields(0).Value; rs.Close()
    check("INSERT referral works", cnt == 1)
    db.Execute("DELETE FROM T_REFERRALS WHERE member_name='TEST_REF'")

    # ========================================
    print("\n" + "=" * 50)
    print("TEST 3: Queries & Aggregation")
    print("=" * 50)

    # Q_Team_Monthly_Sum for 2026/03
    qd = db.QueryDefs("Q_Team_Monthly_Sum")
    from datetime import date
    qd.Parameters("prmDateFrom").Value = "2026/03/01"
    qd.Parameters("prmDateTo").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    if not rs.EOF:
        calls = rs.Fields("sum_calls").Value or 0
        check(f"Q_Team_Monthly_Sum 2026/03 calls = 8412 (got {calls})", calls == 8412)
        valid = rs.Fields("sum_valid").Value or 0
        check(f"Q_Team_Monthly_Sum valid = 4612 (got {valid})", valid == 4612)
        prosp = rs.Fields("sum_prospect").Value or 0
        check(f"Q_Team_Monthly_Sum prospect = 237 (got {prosp})", prosp == 237)
    else:
        check("Q_Team_Monthly_Sum has data", False)
    rs.Close()

    # Q_Team_Monthly_Targets
    qd = db.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear").Value = 2026
    qd.Parameters("prmMonth").Value = 3
    rs = qd.OpenRecordset()
    has_targets = not rs.EOF
    check("Q_Team_Monthly_Targets returns data", has_targets)
    rs.Close()

    # Q_Rank_Received
    qd = db.QueryDefs("Q_Rank_Received")
    qd.Parameters("prmDateFrom").Value = "2026/03/01"
    qd.Parameters("prmDateTo").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    if not rs.EOF:
        first = rs.Fields("member_name").Value
        check(f"Q_Rank_Received top member exists: {first}", first is not None and len(first) > 0)
    rs.Close()

    # Q_Trend_Monthly (6 months)
    qd = db.QueryDefs("Q_Trend_Monthly")
    qd.Parameters("prmTrendStart").Value = "2025/10/01"
    qd.Parameters("prmTrendEnd").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    cnt = 0
    while not rs.EOF:
        cnt += 1; rs.MoveNext()
    rs.Close()
    check(f"Q_Trend_Monthly returns {cnt} months (expect 4-6)", cnt >= 4)

    # Q_Member_12Month
    qd = db.QueryDefs("Q_Member_12Month")
    qd.Parameters("prm12Start").Value = "2025/04/01"
    qd.Parameters("prm12End").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    cnt = 0
    while not rs.EOF:
        cnt += 1; rs.MoveNext()
    rs.Close()
    check(f"Q_Member_12Month returns {cnt} rows (expect >50)", cnt > 50)

    # ========================================
    print("\n" + "=" * 50)
    print("TEST 4: UI / Forms")
    print("=" * 50)

    # Check forms exist
    form_names = ["F_Main", "F_Daily", "F_DailyEdit", "F_Members", "F_Targets", "F_Referrals", "F_Report", "F_Ranking"]
    existing_forms = set()
    for i in range(db.Containers("Forms").Documents.Count):
        existing_forms.add(db.Containers("Forms").Documents(i).Name)

    for fn in form_names:
        check(f"Form {fn} exists", fn in existing_forms)

    # Check queries exist
    query_names = ["Q_ActiveMembers", "Q_Team_Monthly_Sum", "Q_Team_Monthly_Targets",
                   "Q_Member_Monthly_Sum", "Q_Rank_Received", "Q_Rank_Prospect",
                   "Q_Rank_Referral", "Q_Rank_Productivity", "Q_Hourly_By_Member",
                   "Q_Trend_Monthly", "Q_Member_12Month", "Q_RefTrend_Monthly"]
    existing_queries = set()
    for i in range(db.QueryDefs.Count):
        existing_queries.add(db.QueryDefs(i).Name)

    for qn in query_names:
        check(f"Query {qn} exists", qn in existing_queries)

    # Startup form
    try:
        startup = db.Properties("StartUpForm").Value
        check(f"Startup form = F_Main (got {startup})", startup == "F_Main")
    except:
        check("Startup form set", False)

    try:
        title = db.Properties("AppTitle").Value
        check(f"App title set: {title}", len(title) > 0)
    except:
        check("App title set", False)

    db.Close()

    # ========================================
    print("\n" + "=" * 50)
    print(f"RESULTS: {PASS} PASS / {FAIL} FAIL / {PASS+FAIL} TOTAL")
    print("=" * 50)
    if FAIL == 0:
        print("ALL TESTS PASSED!")
    else:
        print(f"{FAIL} test(s) failed.")


if __name__ == "__main__":
    main()
