# -*- coding: utf-8 -*-
"""全クエリにPARAMETERS句を追加 + F_Report VBA修正"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

# パラメータなしクエリはそのまま
# パラメータありクエリにPARAMETERS句を追加

QUERIES = {
    "Q_ActiveMembers": """
SELECT ID, member_name, active
FROM T_MEMBERS
WHERE active = True
ORDER BY member_name;
""",

    "Q_Team_Monthly_Sum": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT
    Sum(R.calls) AS sum_calls,
    Sum(R.valid_count) AS sum_valid,
    Sum(R.prospect) AS sum_prospect,
    Sum(R.doc) AS sum_doc,
    Sum(R.follow_up) AS sum_follow,
    Sum(R.received) AS sum_received,
    Sum(R.work_hours) AS sum_hours,
    Sum(R.referral) AS sum_referral,
    Sum(IIf(R.work_day, 1, 0)) AS work_days
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo];
""",

    "Q_Team_Monthly_Targets": """
PARAMETERS prmYear Long, prmMonth Long;
SELECT
    Sum(target_calls) AS sum_tgt_calls,
    Sum(target_valid) AS sum_tgt_valid,
    Sum(target_prospect) AS sum_tgt_prospect,
    Sum(target_received) AS sum_tgt_received,
    Sum(target_referral) AS sum_tgt_referral
FROM T_MEMBER_TARGETS
WHERE target_year = [prmYear]
    AND target_month = [prmMonth];
""",

    "Q_Member_Monthly_Sum": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT
    R.member_name,
    Sum(R.calls) AS sum_calls,
    Sum(R.valid_count) AS sum_valid,
    Sum(R.prospect) AS sum_prospect,
    Sum(R.doc) AS sum_doc,
    Sum(R.follow_up) AS sum_follow,
    Sum(R.received) AS sum_received,
    Sum(R.work_hours) AS sum_hours,
    Sum(R.referral) AS sum_referral,
    Sum(IIf(R.work_day, 1, 0)) AS work_days,
    Sum(R.calls_10) AS s10, Sum(R.calls_11) AS s11, Sum(R.calls_12) AS s12,
    Sum(R.calls_13) AS s13, Sum(R.calls_14) AS s14, Sum(R.calls_15) AS s15,
    Sum(R.calls_16) AS s16, Sum(R.calls_17) AS s17, Sum(R.calls_18) AS s18
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
ORDER BY R.member_name;
""",

    "Q_Rank_Received": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT R.member_name, Sum(R.received) AS sum_received
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
ORDER BY Sum(R.received) DESC;
""",

    "Q_Rank_Prospect": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT R.member_name, Sum(R.prospect) AS sum_prospect
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
ORDER BY Sum(R.prospect) DESC;
""",

    "Q_Rank_Referral": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT R.member_name, Sum(R.referral) AS sum_ref
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
ORDER BY Sum(R.referral) DESC;
""",

    "Q_Rank_Productivity": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT R.member_name,
    Sum(R.prospect) AS sum_prospect,
    Sum(R.follow_up) AS sum_follow,
    Sum(R.work_hours) AS sum_hours,
    IIf(Sum(R.work_hours)>0,
        (Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours),
        0) AS productivity
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
HAVING Sum(R.work_hours) > 0
ORDER BY IIf(Sum(R.work_hours)>0,
    (Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours),
    0) DESC;
""",

    "Q_Hourly_By_Member": """
PARAMETERS prmDateFrom DateTime, prmDateTo DateTime;
SELECT R.member_name,
    Sum(R.calls_10) AS s10, Sum(R.calls_11) AS s11, Sum(R.calls_12) AS s12,
    Sum(R.calls_13) AS s13, Sum(R.calls_14) AS s14, Sum(R.calls_15) AS s15,
    Sum(R.calls_16) AS s16, Sum(R.calls_17) AS s17, Sum(R.calls_18) AS s18,
    Sum(IIf(R.work_day, 1, 0)) AS work_days
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmDateFrom]
    AND R.rec_date < [prmDateTo]
GROUP BY R.member_name
ORDER BY R.member_name;
""",

    "Q_Trend_Monthly": """
PARAMETERS prmTrendStart DateTime, prmTrendEnd DateTime;
SELECT Year(R.rec_date) AS y, Month(R.rec_date) AS m,
    Sum(R.calls) AS sum_calls,
    Sum(R.valid_count) AS sum_valid,
    Sum(R.prospect) AS sum_prospect,
    Sum(R.received) AS sum_received,
    Sum(R.work_hours) AS sum_hours,
    Sum(R.referral) AS sum_referral,
    Sum(R.doc) AS sum_doc,
    Sum(R.follow_up) AS sum_follow
FROM T_RECORDS AS R
    INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
WHERE M.active = True
    AND R.rec_date >= [prmTrendStart]
    AND R.rec_date < [prmTrendEnd]
GROUP BY Year(R.rec_date), Month(R.rec_date)
ORDER BY Year(R.rec_date), Month(R.rec_date);
""",

    "Q_Member_12Month": """
PARAMETERS prm12Start DateTime, prm12End DateTime;
SELECT R.member_name, Year(R.rec_date) AS y, Month(R.rec_date) AS m,
    Sum(R.calls) AS sum_calls,
    Sum(R.valid_count) AS sum_valid,
    Sum(R.prospect) AS sum_prospect,
    Sum(R.received) AS sum_received,
    Sum(R.work_hours) AS sum_hours,
    Sum(R.referral) AS sum_referral
FROM T_RECORDS AS R
WHERE R.rec_date >= [prm12Start]
    AND R.rec_date < [prm12End]
GROUP BY R.member_name, Year(R.rec_date), Month(R.rec_date)
ORDER BY R.member_name, Year(R.rec_date), Month(R.rec_date);
""",

    "Q_RefTrend_Monthly": """
PARAMETERS prm12Start DateTime, prm12End DateTime;
SELECT Year(rec_date) AS y, Month(rec_date) AS m, Sum(ref_count) AS sum_ref
FROM T_REFERRALS
WHERE rec_date >= [prm12Start]
    AND rec_date < [prm12End]
GROUP BY Year(rec_date), Month(rec_date)
ORDER BY Year(rec_date), Month(rec_date);
""",
}


def main():
    print("全クエリにPARAMETERS句追加")
    print("=" * 50)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(BE)

    for name, sql in QUERIES.items():
        try:
            qd = db.QueryDefs(name)
            qd.SQL = sql.strip()
            print(f"  {name} OK")
        except Exception as e:
            print(f"  {name} ERROR: {e}")

    db.Close()

    # テスト: パラメータ付きクエリを実行
    print("\nTest queries...")
    db = engine.OpenDatabase(BE)

    # Q_Team_Monthly_Sum
    qd = db.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom").Value = "2026/03/01"
    qd.Parameters("prmDateTo").Value = "2026/04/01"
    rs = qd.OpenRecordset()
    if not rs.EOF:
        print(f"  Q_Team_Monthly_Sum: calls={rs.Fields('sum_calls').Value}")
    rs.Close()

    # Q_Team_Monthly_Targets
    qd = db.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear").Value = 2026
    qd.Parameters("prmMonth").Value = 3
    rs = qd.OpenRecordset()
    if not rs.EOF:
        print(f"  Q_Team_Monthly_Targets: tgt_calls={rs.Fields('sum_tgt_calls').Value}")
    rs.Close()

    db.Close()
    print("\nDone!")


if __name__ == "__main__":
    main()
