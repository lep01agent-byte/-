# -*- coding: utf-8 -*-
"""
SalesMgr Access クエリ一括作成スクリプト
既存の SalesMgr_BE.accdb にクエリを追加する。
ダイアログなし・COM不要（DAO直接操作）
"""
import os, sys, time

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BE_PATH = os.path.join(DESKTOP, "SalesMgr_BE.accdb")

def main():
    if not os.path.exists(BE_PATH):
        print(f"ERROR: {BE_PATH} が見つかりません。先に build_access_app.py を実行してください。")
        sys.exit(1)

    import win32com.client
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(BE_PATH)

    queries = get_queries()

    # 既存クエリ削除（再実行対応）
    existing = set()
    for i in range(db.QueryDefs.Count):
        existing.add(db.QueryDefs(i).Name)

    for qname, sql in queries.items():
        if qname in existing:
            db.QueryDefs.Delete(qname)
        db.CreateQueryDef(qname, sql)
        print(f"  {qname}")

    db.Close()
    print(f"\nクエリ {len(queries)}個 作成完了: {BE_PATH}")


def get_queries():
    q = {}

    # ==========================================
    # 基本
    # ==========================================
    q['Q_ActiveMembers'] = """
        SELECT ID, member_name, active
        FROM T_MEMBERS
        WHERE active = True
        ORDER BY member_name
    """

    q['Q_AllMembers'] = """
        SELECT ID, member_name, active
        FROM T_MEMBERS
        ORDER BY member_name
    """

    # ==========================================
    # チーム月次集計
    # ==========================================
    q['Q_Team_Monthly_Sum'] = """
        SELECT
            Sum(R.calls) AS sum_calls,
            Sum(R.valid_count) AS sum_valid,
            Sum(R.prospect) AS sum_prospect,
            Sum(R.doc) AS sum_doc,
            Sum(R.follow_up) AS sum_follow,
            Sum(R.received) AS sum_received,
            Sum(R.work_hours) AS sum_hours,
            Sum(R.referral) AS sum_referral,
            Sum(IIf(R.work_day,1,0)) AS work_days
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
    """

    q['Q_Team_Monthly_Targets'] = """
        SELECT
            Sum(target_calls) AS sum_tgt_calls,
            Sum(target_valid) AS sum_tgt_valid,
            Sum(target_prospect) AS sum_tgt_prospect,
            Sum(target_received) AS sum_tgt_received,
            Sum(target_referral) AS sum_tgt_referral
        FROM T_MEMBER_TARGETS
        WHERE target_year = [prmYear]
          AND target_month = [prmMonth]
    """

    # ==========================================
    # メンバー別月次集計
    # ==========================================
    q['Q_Member_Monthly_Sum'] = """
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
            Sum(IIf(R.work_day,1,0)) AS work_days,
            Sum(R.calls_10) AS s10, Sum(R.calls_11) AS s11,
            Sum(R.calls_12) AS s12, Sum(R.calls_13) AS s13,
            Sum(R.calls_14) AS s14, Sum(R.calls_15) AS s15,
            Sum(R.calls_16) AS s16, Sum(R.calls_17) AS s17,
            Sum(R.calls_18) AS s18
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY R.member_name
    """

    # ==========================================
    # ランキング
    # ==========================================
    q['Q_Rank_Referral'] = """
        SELECT R.member_name, Sum(R.referral) AS sum_ref
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.referral) DESC
    """

    q['Q_Rank_Received'] = """
        SELECT R.member_name, Sum(R.received) AS sum_received
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.received) DESC
    """

    q['Q_Rank_Prospect'] = """
        SELECT R.member_name, Sum(R.prospect) AS sum_prospect
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.prospect) DESC
    """

    q['Q_Rank_Productivity'] = """
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
            0) DESC
    """

    # ==========================================
    # 時間帯別
    # ==========================================
    q['Q_Hourly_By_Member'] = """
        SELECT R.member_name,
            Sum(R.calls_10) AS s10, Sum(R.calls_11) AS s11,
            Sum(R.calls_12) AS s12, Sum(R.calls_13) AS s13,
            Sum(R.calls_14) AS s14, Sum(R.calls_15) AS s15,
            Sum(R.calls_16) AS s16, Sum(R.calls_17) AS s17,
            Sum(R.calls_18) AS s18,
            Sum(IIf(R.work_day,1,0)) AS work_days
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY R.member_name
    """

    # ==========================================
    # 月別推移（過去6ヶ月）
    # ==========================================
    q['Q_Trend_Monthly'] = """
        SELECT
            Year(R.rec_date) AS y,
            Month(R.rec_date) AS m,
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
        ORDER BY Year(R.rec_date), Month(R.rec_date)
    """

    # ==========================================
    # 12ヶ月メンバー別履歴
    # ==========================================
    q['Q_Member_12Month'] = """
        SELECT
            R.member_name,
            Year(R.rec_date) AS y,
            Month(R.rec_date) AS m,
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
        ORDER BY R.member_name, Year(R.rec_date), Month(R.rec_date)
    """

    # ==========================================
    # 送客推移
    # ==========================================
    q['Q_RefTrend_Monthly'] = """
        SELECT
            Year(rec_date) AS y,
            Month(rec_date) AS m,
            Sum(ref_count) AS sum_ref
        FROM T_REFERRALS
        WHERE rec_date >= [prm12Start]
          AND rec_date < [prm12End]
        GROUP BY Year(rec_date), Month(rec_date)
        ORDER BY Year(rec_date), Month(rec_date)
    """

    # ==========================================
    # 日次レコード一覧
    # ==========================================
    q['Q_Records_List'] = """
        SELECT
            R.ID,
            Format(R.rec_date,'yyyy/mm/dd') AS 日付,
            R.member_name AS 担当者,
            R.calls AS 架電,
            R.valid_count AS 有効,
            R.prospect AS 見込,
            R.doc AS 資料,
            R.follow_up AS 追客,
            R.received AS 受注,
            R.referral AS 送客,
            R.work_hours AS 稼働h,
            R.[note] AS 備考,
            R.work_day,
            R.calls_10, R.calls_11, R.calls_12,
            R.calls_13, R.calls_14, R.calls_15,
            R.calls_16, R.calls_17, R.calls_18
        FROM T_RECORDS AS R
        WHERE R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        ORDER BY R.rec_date DESC, R.member_name, R.ID DESC
    """

    # ==========================================
    # 目標一覧
    # ==========================================
    q['Q_Targets_Monthly'] = """
        SELECT
            T.ID,
            T.member_name AS 担当者,
            T.target_calls AS 架電目標,
            T.target_valid AS 有効目標,
            T.target_prospect AS 見込目標,
            T.target_received AS 受注目標,
            T.target_referral AS 送客目標,
            T.plan_days AS 稼働日数,
            T.work_hours_per_day AS 日稼働h
        FROM T_MEMBER_TARGETS AS T
        WHERE T.target_year = [prmYear]
          AND T.target_month = [prmMonth]
        ORDER BY T.member_name
    """

    # ==========================================
    # 送客一覧
    # ==========================================
    q['Q_Referrals_Monthly'] = """
        SELECT
            R.ID,
            Format(R.rec_date,'yyyy/mm/dd') AS 日付,
            R.member_name AS 担当者,
            R.ref_count AS 件数
        FROM T_REFERRALS AS R
        WHERE R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        ORDER BY R.rec_date DESC, R.member_name
    """

    # ==========================================
    # 月次メンバー別サマリー（履歴画面用）
    # ==========================================
    q['Q_List_Summary'] = """
        SELECT
            R.member_name AS 担当者,
            Sum(IIf(R.work_day,1,0)) AS 稼働日,
            Sum(R.calls) AS 架電,
            Sum(R.valid_count) AS 有効,
            Sum(R.prospect) AS 見込,
            Sum(R.received) AS 受注,
            Sum(R.work_hours) AS 稼働h,
            IIf(Sum(R.calls)>0,
                Format(Sum(R.valid_count)/Sum(R.calls)*100,'0.0') & '%',
                '-') AS 有効率,
            IIf(Sum(R.calls)>0,
                Format(Sum(R.received)/Sum(R.calls)*100,'0.0') & '%',
                '-') AS 受注率
        FROM T_RECORDS AS R
        INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom]
          AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY R.member_name
    """

    return q


if __name__ == '__main__':
    print("SalesMgr クエリ作成")
    print("=" * 40)

    # Accessプロセスが開いていたら閉じる
    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(1)

    main()
