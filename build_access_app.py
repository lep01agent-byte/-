# -*- coding: utf-8 -*-
"""
SalesMgr Access アプリケーション ビルドスクリプト
Flask+SQLiteアプリをMicrosoft Access (.accdb) に完全移行する。

使い方: python build_access_app.py
出力:
  - Desktop\SalesMgr_BE.accdb (バックエンドDB)
  - Desktop\SalesMgr_FE.accdb (フロントエンドUI)
"""
import os, sys, time, sqlite3, pyodbc, datetime, shutil

# ============================================================
# 設定
# ============================================================
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BE_PATH = os.path.join(DESKTOP, "SalesMgr_BE.accdb")
FE_PATH = os.path.join(DESKTOP, "SalesMgr_FE.accdb")
SQLITE_PATH = os.path.join(os.path.expanduser("~"), "salesmgr_web", "dbfile.db")

ODBC_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"

# ============================================================
# Phase 1: バックエンドDB作成 + データ移行
# ============================================================
def create_backend():
    print("=" * 60)
    print("Phase 1: バックエンドDB作成")
    print("=" * 60)

    # 既存ファイル削除
    if os.path.exists(BE_PATH):
        os.remove(BE_PATH)
        print(f"  既存ファイル削除: {BE_PATH}")

    # Access DBを作成 (pyodbc経由)
    conn_str = (
        f"DRIVER={{{ODBC_DRIVER}}};"
        f"DBQ={BE_PATH};"
        "NEWDB=1;"
    )
    # pyodbc cannot create Access DB with NEWDB, use win32com instead
    import win32com.client
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.CreateDatabase(BE_PATH, ";LANGID=0x0411;CP=1252;COUNTRY=0", 64)
    db.Close()
    print(f"  作成: {BE_PATH}")

    # テーブル作成 (DAO経由 - Access DDLの制限を回避)
    print("  テーブル作成中...")

    import win32com.client as w32
    engine = w32.Dispatch("DAO.DBEngine.120")
    daodb = engine.OpenDatabase(BE_PATH)

    # DAO定数
    dbAutoIncrField = 16
    dbLong = 4
    dbText = 10
    dbDate = 8
    dbDouble = 7
    dbBoolean = 1
    dbMemo = 12

    def create_table(daodb, tbl_name, fields, indexes=None):
        td = daodb.CreateTableDef(tbl_name)
        for fname, ftype, fsize, fattribs in fields:
            fld = td.CreateField(fname, ftype, fsize)
            if fattribs:
                fld.Attributes = fattribs
            if ftype == dbText:
                fld.AllowZeroLength = True
            td.Fields.Append(fld)
        # Primary key on ID
        idx = td.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Unique = True
        idf = idx.CreateField("ID")
        idx.Fields.Append(idf)
        td.Indexes.Append(idx)
        daodb.TableDefs.Append(td)
        # Additional indexes
        if indexes:
            td2 = daodb.TableDefs(tbl_name)
            for ix_name, ix_fields, ix_unique in indexes:
                ix = td2.CreateIndex(ix_name)
                ix.Unique = ix_unique
                for ixf in ix_fields:
                    ix.Fields.Append(ix.CreateField(ixf))
                td2.Indexes.Append(ix)

    create_table(daodb, "T_MEMBERS", [
        ("ID", dbLong, 0, dbAutoIncrField),
        ("member_name", dbText, 50, 0),
        ("active", dbBoolean, 0, 0),
    ], [("UQ_member_name", ["member_name"], True)])

    create_table(daodb, "T_RECORDS", [
        ("ID", dbLong, 0, dbAutoIncrField),
        ("rec_date", dbDate, 0, 0),
        ("member_name", dbText, 50, 0),
        ("calls", dbLong, 0, 0),
        ("calls_10", dbLong, 0, 0), ("calls_11", dbLong, 0, 0),
        ("calls_12", dbLong, 0, 0), ("calls_13", dbLong, 0, 0),
        ("calls_14", dbLong, 0, 0), ("calls_15", dbLong, 0, 0),
        ("calls_16", dbLong, 0, 0), ("calls_17", dbLong, 0, 0),
        ("calls_18", dbLong, 0, 0),
        ("valid_count", dbLong, 0, 0),
        ("prospect", dbLong, 0, 0),
        ("doc", dbLong, 0, 0),
        ("follow_up", dbLong, 0, 0),
        ("received", dbLong, 0, 0),
        ("work_hours", dbDouble, 0, 0),
        ("note", dbText, 255, 0),
        ("referral", dbLong, 0, 0),
        ("work_day", dbBoolean, 0, 0),
    ], [("IX_rec_date", ["rec_date"], False), ("IX_rec_member", ["member_name"], False)])

    create_table(daodb, "T_MEMBER_TARGETS", [
        ("ID", dbLong, 0, dbAutoIncrField),
        ("member_name", dbText, 50, 0),
        ("target_year", dbLong, 0, 0),
        ("target_month", dbLong, 0, 0),
        ("plan_days", dbLong, 0, 0),
        ("work_hours_per_day", dbDouble, 0, 0),
        ("target_calls", dbLong, 0, 0),
        ("target_valid", dbLong, 0, 0),
        ("target_prospect", dbLong, 0, 0),
        ("target_received", dbLong, 0, 0),
        ("target_referral", dbLong, 0, 0),
    ], [("UQ_target", ["member_name", "target_year", "target_month"], True)])

    create_table(daodb, "T_REFERRALS", [
        ("ID", dbLong, 0, dbAutoIncrField),
        ("rec_date", dbDate, 0, 0),
        ("member_name", dbText, 50, 0),
        ("ref_count", dbLong, 0, 0),
    ], [("IX_ref_date", ["rec_date"], False), ("IX_ref_member", ["member_name"], False)])

    daodb.Close()
    print("  テーブル作成完了")

    # ODBC接続 (データ移行用)
    conn_str = f"DRIVER={{{ODBC_DRIVER}}};DBQ={BE_PATH};"
    conn = pyodbc.connect(conn_str)
    cur = conn.cursor()

    # --- データ移行 ---
    print(f"\n  SQLiteからデータ移行: {SQLITE_PATH}")
    scon = sqlite3.connect(SQLITE_PATH)
    scur = scon.cursor()

    # Members
    scur.execute("SELECT member_name, active FROM members ORDER BY id")
    rows = scur.fetchall()
    for row in rows:
        cur.execute("INSERT INTO T_MEMBERS (member_name, active) VALUES (?, ?)",
                    row[0], bool(row[1]) if row[1] is not None else True)
    conn.commit()
    print(f"  T_MEMBERS: {len(rows)} 件")

    # Records  (note -> [note] for Access reserved word)
    scur.execute("""SELECT rec_date, member_name, calls,
                    calls_10, calls_11, calls_12, calls_13, calls_14,
                    calls_15, calls_16, calls_17, calls_18,
                    valid, prospect, doc, follow, received,
                    work_hours, note, referral, work_day
                    FROM records ORDER BY id""")
    rows = scur.fetchall()
    insert_sql = """INSERT INTO T_RECORDS (rec_date, member_name, calls,
                      calls_10, calls_11, calls_12, calls_13, calls_14,
                      calls_15, calls_16, calls_17, calls_18,
                      valid_count, prospect, doc, follow_up, received,
                      work_hours, [note], referral, work_day)
                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
    for row in rows:
        try:
            parts = row[0].replace('-', '/').split('/')
            dt = datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
        except:
            dt = datetime.date.today()
        cur.execute(insert_sql,
                    dt, row[1], row[2] or 0,
                    row[3] or 0, row[4] or 0, row[5] or 0, row[6] or 0, row[7] or 0,
                    row[8] or 0, row[9] or 0, row[10] or 0, row[11] or 0,
                    row[12] or 0, row[13] or 0, row[14] or 0, row[15] or 0, row[16] or 0,
                    float(row[17]) if row[17] else 8.0,
                    row[18] if row[18] else None,
                    row[19] or 0,
                    bool(row[20]) if row[20] is not None else True)
    conn.commit()
    print(f"  T_RECORDS: {len(rows)} 件")

    # Member Targets
    scur.execute("""SELECT member_name, year, month, plan_days,
                    work_hours_per_day, target_calls, target_valid,
                    target_prospect, target_received, target_referral
                    FROM member_targets ORDER BY id""")
    rows = scur.fetchall()
    for row in rows:
        cur.execute("""INSERT INTO T_MEMBER_TARGETS (member_name, target_year, target_month,
                      plan_days, work_hours_per_day, target_calls, target_valid,
                      target_prospect, target_received, target_referral)
                      VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    row[0], row[1], row[2],
                    row[3] or 20, float(row[4]) if row[4] else 8.0,
                    row[5] or 0, row[6] or 0, row[7] or 0, row[8] or 0, row[9] or 0)
    conn.commit()
    print(f"  T_MEMBER_TARGETS: {len(rows)} 件")

    # Referrals
    scur.execute("SELECT rec_date, member_name, count FROM referrals ORDER BY id")
    rows = scur.fetchall()
    for row in rows:
        try:
            parts = row[0].replace('-', '/').split('/')
            dt = datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
        except:
            dt = datetime.date.today()
        cur.execute("INSERT INTO T_REFERRALS (rec_date, member_name, ref_count) VALUES (?,?,?)",
                    dt, row[1], row[2] or 0)
    conn.commit()
    print(f"  T_REFERRALS: {len(rows)} 件")

    scon.close()
    cur.close()
    conn.close()
    print("\nPhase 1 完了!")


# ============================================================
# Phase 2: フロントエンドDB作成
# ============================================================
def create_frontend():
    print("\n" + "=" * 60)
    print("Phase 2: フロントエンドDB作成")
    print("=" * 60)

    # 既存ファイル削除
    if os.path.exists(FE_PATH):
        os.remove(FE_PATH)

    import win32com.client
    from win32com.client import constants as c

    # Access アプリケーション起動
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False

    # 新規DB作成
    app.NewCurrentDatabase(FE_PATH)
    time.sleep(2)
    db = app.CurrentDb()

    # --- リンクテーブル ---
    print("  リンクテーブル作成中...")
    for tbl_name in ['T_MEMBERS', 'T_RECORDS', 'T_MEMBER_TARGETS', 'T_REFERRALS']:
        td = db.CreateTableDef(tbl_name)
        td.Connect = f";DATABASE={BE_PATH}"
        td.SourceTableName = tbl_name
        db.TableDefs.Append(td)
    print("  リンクテーブル 4個 完了")

    # --- クエリ作成 ---
    print("  クエリ作成中...")
    queries = _get_all_queries()
    for qname, sql in queries.items():
        try:
            qd = db.CreateQueryDef(qname, sql)
            # パラメータクエリの場合、最大レコード数を0に
        except Exception as e:
            print(f"    WARNING: {qname}: {e}")
    print(f"  クエリ {len(queries)}個 完了")

    # --- VBAモジュール (ファイルに書き出し、後でインポート) ---
    print("  VBAモジュール .bas ファイル出力中...")
    _export_vba_files()
    print("  VBAモジュール .bas 出力完了")

    # COM経由のVBA注入は試みる（ダイアログが出る場合はスキップ）
    try:
        _create_vba_modules(app)
    except Exception as e:
        print(f"    VBA COM注入スキップ: {e}")
        print("    → .bas ファイルを手動でインポートしてください")

    # --- フォーム作成 ---
    print("  フォーム作成中...")
    _create_all_forms(app, db)
    print("  フォーム 完了")

    # --- レポート作成 ---
    print("  レポート作成中...")
    _create_reports(app, db)
    print("  レポート 完了")

    # --- スタートアップ設定 ---
    print("  スタートアップ設定中...")
    _set_startup(app, db)

    # 閉じる
    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(2)
    print("\nPhase 2 完了!")


# ============================================================
# クエリ定義
# ============================================================
def _get_all_queries():
    q = {}

    # 基本
    q['Q_ActiveMembers'] = """
        SELECT ID, member_name, active
        FROM T_MEMBERS WHERE active = True ORDER BY member_name
    """
    q['Q_AllMembers'] = """
        SELECT ID, member_name, active FROM T_MEMBERS ORDER BY member_name
    """

    # チーム月次集計 (日付パラメータはVBAから動的に設定)
    q['Q_Team_Monthly_Sum'] = """
        SELECT
            Sum(R.calls) AS sum_calls,
            Sum(R.valid_count) AS sum_valid,
            Sum(R.prospect) AS sum_prospect,
            Sum(R.doc) AS sum_doc,
            Sum(R.follow_up) AS sum_follow,
            Sum(R.received) AS sum_received,
            Sum(R.work_hours) AS sum_hours,
            Sum(R.referral) AS sum_referral
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
    """

    q['Q_Team_Monthly_Targets'] = """
        SELECT
            Sum(target_calls) AS sum_tgt_calls,
            Sum(target_valid) AS sum_tgt_valid,
            Sum(target_prospect) AS sum_tgt_prospect,
            Sum(target_received) AS sum_tgt_received,
            Sum(target_referral) AS sum_tgt_referral
        FROM T_MEMBER_TARGETS
        WHERE target_year = [prmYear] AND target_month = [prmMonth]
    """

    # メンバー別月次集計
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
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY R.member_name
    """

    # ランキング
    q['Q_Rank_Referral'] = """
        SELECT R.member_name, Sum(R.referral) AS sum_ref
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.referral) DESC
    """

    q['Q_Rank_Received'] = """
        SELECT R.member_name, Sum(R.received) AS sum_received
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.received) DESC
    """

    q['Q_Rank_Prospect'] = """
        SELECT R.member_name, Sum(R.prospect) AS sum_prospect
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        ORDER BY Sum(R.prospect) DESC
    """

    q['Q_Rank_Productivity'] = """
        SELECT R.member_name,
            Sum(R.prospect) AS sum_prospect,
            Sum(R.follow_up) AS sum_follow,
            Sum(R.work_hours) AS sum_hours,
            IIf(Sum(R.work_hours)>0, (Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours), 0) AS productivity
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name
        HAVING Sum(R.work_hours) > 0
        ORDER BY IIf(Sum(R.work_hours)>0, (Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours), 0) DESC
    """

    # 時間帯別
    q['Q_Hourly_By_Member'] = """
        SELECT R.member_name,
            Sum(R.calls_10) AS s10, Sum(R.calls_11) AS s11,
            Sum(R.calls_12) AS s12, Sum(R.calls_13) AS s13,
            Sum(R.calls_14) AS s14, Sum(R.calls_15) AS s15,
            Sum(R.calls_16) AS s16, Sum(R.calls_17) AS s17,
            Sum(R.calls_18) AS s18,
            Sum(IIf(R.work_day,1,0)) AS work_days
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        GROUP BY R.member_name ORDER BY R.member_name
    """

    # 6ヶ月推移
    q['Q_Trend_Monthly'] = """
        SELECT Year(R.rec_date) AS y, Month(R.rec_date) AS m,
            Sum(R.calls) AS sum_calls, Sum(R.valid_count) AS sum_valid,
            Sum(R.prospect) AS sum_prospect, Sum(R.received) AS sum_received,
            Sum(R.work_hours) AS sum_hours, Sum(R.referral) AS sum_referral
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name
        WHERE M.active = True
          AND R.rec_date >= [prmTrendStart] AND R.rec_date < [prmTrendEnd]
        GROUP BY Year(R.rec_date), Month(R.rec_date)
        ORDER BY Year(R.rec_date), Month(R.rec_date)
    """

    # 12ヶ月メンバー別履歴
    q['Q_Member_12Month'] = """
        SELECT R.member_name, Year(R.rec_date) AS y, Month(R.rec_date) AS m,
            Sum(R.calls) AS sum_calls, Sum(R.valid_count) AS sum_valid,
            Sum(R.prospect) AS sum_prospect, Sum(R.received) AS sum_received,
            Sum(R.work_hours) AS sum_hours, Sum(R.referral) AS sum_referral
        FROM T_RECORDS AS R
        WHERE R.rec_date >= [prm12Start] AND R.rec_date < [prm12End]
        GROUP BY R.member_name, Year(R.rec_date), Month(R.rec_date)
        ORDER BY R.member_name, Year(R.rec_date), Month(R.rec_date)
    """

    # 送客推移
    q['Q_RefTrend_Monthly'] = """
        SELECT Year(rec_date) AS y, Month(rec_date) AS m, Sum(ref_count) AS sum_ref
        FROM T_REFERRALS
        WHERE rec_date >= [prm12Start] AND rec_date < [prm12End]
        GROUP BY Year(rec_date), Month(rec_date)
        ORDER BY Year(rec_date), Month(rec_date)
    """

    # 日次一覧
    q['Q_Records_List'] = """
        SELECT R.ID, R.rec_date, R.member_name, R.calls,
            R.valid_count, R.prospect, R.doc, R.follow_up, R.received,
            R.work_hours, R.note, R.referral, R.work_day,
            R.calls_10, R.calls_11, R.calls_12, R.calls_13,
            R.calls_14, R.calls_15, R.calls_16, R.calls_17, R.calls_18
        FROM T_RECORDS AS R
        WHERE R.rec_date >= [prmDateFrom] AND R.rec_date < [prmDateTo]
        ORDER BY R.rec_date DESC, R.member_name, R.ID DESC
    """

    # 目標一覧
    q['Q_Targets_Monthly'] = """
        SELECT * FROM T_MEMBER_TARGETS
        WHERE target_year = [prmYear] AND target_month = [prmMonth]
        ORDER BY member_name
    """

    # 送客一覧
    q['Q_Referrals_Monthly'] = """
        SELECT * FROM T_REFERRALS
        WHERE rec_date >= [prmDateFrom] AND rec_date < [prmDateTo]
        ORDER BY rec_date DESC, member_name
    """

    return q


# ============================================================
# VBAコード
# ============================================================
def _get_vba_modGlobal():
    return """
Option Compare Database
Option Explicit

' 定数
Public Const APP_TITLE As String = "SalesMgr 営業管理"
Public Const FONT_NAME As String = "Yu Gothic UI"

' 色定数
Public Const CLR_HEADER As Long = 3877662   ' RGB(30,41,59) = #1E293B
Public Const CLR_WHITE As Long = 16777215
Public Const CLR_BLUE As Long = 16236344    ' RGB(56,189,248)
Public Const CLR_GREEN As Long = 6333218    ' RGB(34,197,94)
Public Const CLR_RED As Long = 4473071      ' RGB(239,68,68)
Public Const CLR_AMBER As Long = 2473979    ' RGB(251,191,36)
Public Const CLR_LIGHT_BG As Long = 16579832 ' RGB(248,250,252)
Public Const CLR_UP As Long = 5087510       ' RGB(22,163,74)
Public Const CLR_DOWN As Long = 2498780     ' RGB(220,38,38)

' グローバル年月
Public g_Year As Integer
Public g_Month As Integer

Public Function DateFrom(y As Integer, m As Integer) As Date
    DateFrom = DateSerial(y, m, 1)
End Function

Public Function DateTo(y As Integer, m As Integer) As Date
    DateTo = DateSerial(y, m + 1, 1)
End Function

Public Sub PrevYM(ByRef y As Integer, ByRef m As Integer)
    If m = 1 Then
        y = y - 1: m = 12
    Else
        m = m - 1
    End If
End Sub

Public Sub NextYM(ByRef y As Integer, ByRef m As Integer)
    If m = 12 Then
        y = y + 1: m = 1
    Else
        m = m + 1
    End If
End Sub

Public Function WorkDaysInMonth(y As Integer, m As Integer) As Integer
    Dim d As Date, cnt As Integer
    d = DateSerial(y, m, 1)
    Do While Month(d) = m
        If Weekday(d, vbMonday) <= 5 Then cnt = cnt + 1
        d = d + 1
    Loop
    WorkDaysInMonth = cnt
End Function

Public Function Seika(prospect As Long, follow_up As Long) As Double
    Seika = prospect + follow_up * 0.5
End Function

Public Function FmtComma(val As Variant) As String
    If IsNull(val) Or val = 0 Then
        FmtComma = "0"
    Else
        FmtComma = Format(val, "#,##0")
    End If
End Function

Public Function FmtPct(num As Variant, denom As Variant) As String
    If IsNull(denom) Or denom = 0 Then
        FmtPct = "-"
    Else
        FmtPct = Format(num / denom * 100, "0.0") & "%"
    End If
End Function

Public Function Arrow(cur As Variant, prev As Variant) As String
    Dim diff As Long
    If IsNull(cur) Then cur = 0
    If IsNull(prev) Then prev = 0
    diff = CLng(cur) - CLng(prev)
    If diff > 0 Then
        Arrow = Chr(9650) & FmtComma(diff)
    ElseIf diff < 0 Then
        Arrow = Chr(9660) & FmtComma(Abs(diff))
    Else
        Arrow = "-"
    End If
End Function

Public Function RunParamQuery(qName As String, ParamArray params() As Variant) As DAO.Recordset
    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs(qName)
    Dim i As Integer
    For i = 0 To UBound(params) Step 2
        qd.Parameters(CStr(params(i))) = params(i + 1)
    Next i
    Set RunParamQuery = qd.OpenRecordset(dbOpenSnapshot)
End Function

Public Sub ExportCSV(y As Integer, m As Integer)
    Dim fd As Object
    Set fd = Application.FileDialog(2) ' msoFileDialogSaveAs
    fd.Title = "CSV出力先を選択"
    fd.InitialFileName = "records_" & y & Format(m, "00") & ".csv"
    If fd.Show = -1 Then
        Dim path As String: path = fd.SelectedItems(1)
        Dim rs As DAO.Recordset
        Set rs = RunParamQuery("Q_Records_List", "prmDateFrom", DateFrom(y, m), "prmDateTo", DateTo(y, m))
        Dim f As Integer: f = FreeFile
        Open path For Output As #f
        ' Header
        Print #f, "日付,担当者,架電,有効,見込,資料,追客,受注,稼働時間,備考"
        Do While Not rs.EOF
            Print #f, Format(rs("rec_date"), "yyyy/mm/dd") & "," & _
                rs("member_name") & "," & Nz(rs("calls"), 0) & "," & _
                Nz(rs("valid_count"), 0) & "," & Nz(rs("prospect"), 0) & "," & _
                Nz(rs("doc"), 0) & "," & Nz(rs("follow_up"), 0) & "," & _
                Nz(rs("received"), 0) & "," & Nz(rs("work_hours"), 0) & "," & _
                Chr(34) & Nz(rs("note"), "") & Chr(34)
            rs.MoveNext
        Loop
        Close #f
        rs.Close
        MsgBox "CSV出力完了: " & path, vbInformation
    End If
End Sub

Public Sub ExportPDF(y As Integer, m As Integer)
    Dim fd As Object
    Set fd = Application.FileDialog(2)
    fd.Title = "PDF出力先を選択"
    fd.InitialFileName = "report_" & y & Format(m, "00") & ".pdf"
    If fd.Show = -1 Then
        Dim path As String: path = fd.SelectedItems(1)
        ' OpenArgsに年月を渡す
        DoCmd.OpenReport "R_MonthlyReport", acViewPreview, , , , y & "|" & m
        DoCmd.OutputTo acOutputReport, "R_MonthlyReport", acFormatPDF, path
        DoCmd.Close acReport, "R_MonthlyReport"
        MsgBox "PDF出力完了: " & path, vbInformation
    End If
End Sub
"""


def _get_vba_form_main():
    return """
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    Me.Caption = APP_TITLE
End Sub

Private Sub btnDaily_Click()
    DoCmd.OpenForm "F_Daily"
End Sub

Private Sub btnTargets_Click()
    DoCmd.OpenForm "F_Targets"
End Sub

Private Sub btnReferrals_Click()
    DoCmd.OpenForm "F_Referrals"
End Sub

Private Sub btnReport_Click()
    DoCmd.OpenForm "F_Report"
End Sub

Private Sub btnProgress_Click()
    DoCmd.OpenForm "F_Progress"
End Sub

Private Sub btnList_Click()
    DoCmd.OpenForm "F_List"
End Sub

Private Sub btnAnalysis_Click()
    DoCmd.OpenForm "F_Analysis"
End Sub

Private Sub btnRanking_Click()
    DoCmd.OpenForm "F_Ranking"
End Sub

Private Sub btnRefTrend_Click()
    DoCmd.OpenForm "F_RefTrend"
End Sub

Private Sub btnMembers_Click()
    DoCmd.OpenForm "F_Members"
End Sub
"""


def _get_vba_form_daily():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date)
    m_Month = Month(Date)
    LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & Format(m_Month, "00") & "月 日次一覧"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    ' メンバーコンボ
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"

    ' リスト更新
    Dim sql As String
    sql = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd') AS 日付, R.member_name AS 担当者, " & _
          "R.calls AS 架電, R.valid_count AS 有効, R.prospect AS 見込, " & _
          "R.doc AS 資料, R.follow_up AS 追客, R.received AS 受注, " & _
          "R.work_hours AS 稼働h, R.note AS 備考 " & _
          "FROM T_RECORDS AS R " & _
          "WHERE R.rec_date >= #" & Format(DateSerial(m_Year, m_Month, 1), "yyyy/mm/dd") & "# " & _
          "AND R.rec_date < #" & Format(DateSerial(m_Year, m_Month + 1, 1), "yyyy/mm/dd") & "# "

    If Nz(cboMember.Value, "") <> "" Then
        sql = sql & "AND R.member_name = '" & cboMember.Value & "' "
    End If
    sql = sql & "ORDER BY R.rec_date DESC, R.member_name"

    lstRecords.RowSource = sql
    lstRecords.Requery
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month
    LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month
    LoadData
End Sub

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub btnAdd_Click()
    DoCmd.OpenForm "F_DailyEdit", , , , acFormAdd, acDialog, "ADD|" & m_Year & "|" & m_Month
    LoadData
End Sub

Private Sub btnEdit_Click()
    If IsNull(lstRecords.Value) Then
        MsgBox "編集するレコードを選択してください", vbExclamation
        Exit Sub
    End If
    DoCmd.OpenForm "F_DailyEdit", , , , , acDialog, "EDIT|" & lstRecords.Value
    LoadData
End Sub

Private Sub btnDelete_Click()
    If IsNull(lstRecords.Value) Then
        MsgBox "削除するレコードを選択してください", vbExclamation
        Exit Sub
    End If
    If MsgBox("このレコードを削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_RECORDS WHERE ID = " & lstRecords.Value, dbFailOnError
        LoadData
    End If
End Sub

Private Sub lstRecords_DblClick(Cancel As Integer)
    btnEdit_Click
End Sub
"""


def _get_vba_form_daily_edit():
    return """
Option Compare Database
Option Explicit

Private m_Mode As String
Private m_RecID As Long

Private Sub Form_Open(Cancel As Integer)
    ' OpenArgs: "ADD|year|month" or "EDIT|recordID"
    Dim parts() As String
    parts = Split(Nz(Me.OpenArgs, "ADD"), "|")
    m_Mode = parts(0)

    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"

    If m_Mode = "EDIT" And UBound(parts) >= 1 Then
        m_RecID = CLng(parts(1))
        Me.Caption = "日次レコード編集"
        LoadRecord
    Else
        m_RecID = 0
        Me.Caption = "日次レコード新規登録"
        If UBound(parts) >= 2 Then
            txtRecDate.Value = Format(DateSerial(CInt(parts(1)), CInt(parts(2)), Day(Date)), "yyyy/mm/dd")
        Else
            txtRecDate.Value = Format(Date, "yyyy/mm/dd")
        End If
        txtWorkHours.Value = 8
        chkWorkDay.Value = True
    End If
End Sub

Private Sub LoadRecord()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_RECORDS WHERE ID = " & m_RecID)
    If Not rs.EOF Then
        txtRecDate.Value = Format(rs("rec_date"), "yyyy/mm/dd")
        cboMember.Value = rs("member_name")
        txtC10.Value = Nz(rs("calls_10"), 0)
        txtC11.Value = Nz(rs("calls_11"), 0)
        txtC12.Value = Nz(rs("calls_12"), 0)
        txtC13.Value = Nz(rs("calls_13"), 0)
        txtC14.Value = Nz(rs("calls_14"), 0)
        txtC15.Value = Nz(rs("calls_15"), 0)
        txtC16.Value = Nz(rs("calls_16"), 0)
        txtC17.Value = Nz(rs("calls_17"), 0)
        txtC18.Value = Nz(rs("calls_18"), 0)
        txtValid.Value = Nz(rs("valid_count"), 0)
        txtProspect.Value = Nz(rs("prospect"), 0)
        txtDoc.Value = Nz(rs("doc"), 0)
        txtFollow.Value = Nz(rs("follow_up"), 0)
        txtReceived.Value = Nz(rs("received"), 0)
        txtReferral.Value = Nz(rs("referral"), 0)
        txtWorkHours.Value = Nz(rs("work_hours"), 8)
        chkWorkDay.Value = rs("work_day")
        txtNote.Value = Nz(rs("note"), "")
    End If
    rs.Close
End Sub

Private Sub btnSave_Click()
    ' バリデーション
    If Nz(txtRecDate.Value, "") = "" Then
        MsgBox "日付を入力してください", vbExclamation: Exit Sub
    End If
    If Nz(cboMember.Value, "") = "" Then
        MsgBox "担当者を選択してください", vbExclamation: Exit Sub
    End If

    Dim dt As Date: dt = CDate(txtRecDate.Value)
    Dim calls As Long
    calls = Nz(txtC10, 0) + Nz(txtC11, 0) + Nz(txtC12, 0) + _
            Nz(txtC13, 0) + Nz(txtC14, 0) + Nz(txtC15, 0) + _
            Nz(txtC16, 0) + Nz(txtC17, 0) + Nz(txtC18, 0)

    Dim sql As String
    If m_RecID = 0 Then
        sql = "INSERT INTO T_RECORDS (rec_date, member_name, calls, " & _
              "calls_10,calls_11,calls_12,calls_13,calls_14,calls_15,calls_16,calls_17,calls_18, " & _
              "valid_count,prospect,doc,follow_up,received,work_hours,note,referral,work_day) " & _
              "VALUES (#" & Format(dt, "yyyy/mm/dd") & "#, " & _
              "'" & cboMember.Value & "', " & calls & ", " & _
              Nz(txtC10, 0) & "," & Nz(txtC11, 0) & "," & Nz(txtC12, 0) & "," & _
              Nz(txtC13, 0) & "," & Nz(txtC14, 0) & "," & Nz(txtC15, 0) & "," & _
              Nz(txtC16, 0) & "," & Nz(txtC17, 0) & "," & Nz(txtC18, 0) & "," & _
              Nz(txtValid, 0) & "," & Nz(txtProspect, 0) & "," & Nz(txtDoc, 0) & "," & _
              Nz(txtFollow, 0) & "," & Nz(txtReceived, 0) & "," & _
              Nz(txtWorkHours, 8) & ",'" & Replace(Nz(txtNote, ""), "'", "''") & "'," & _
              Nz(txtReferral, 0) & "," & IIf(chkWorkDay.Value, "True", "False") & ")"
    Else
        sql = "UPDATE T_RECORDS SET " & _
              "rec_date = #" & Format(dt, "yyyy/mm/dd") & "#, " & _
              "member_name = '" & cboMember.Value & "', " & _
              "calls = " & calls & ", " & _
              "calls_10=" & Nz(txtC10, 0) & ",calls_11=" & Nz(txtC11, 0) & "," & _
              "calls_12=" & Nz(txtC12, 0) & ",calls_13=" & Nz(txtC13, 0) & "," & _
              "calls_14=" & Nz(txtC14, 0) & ",calls_15=" & Nz(txtC15, 0) & "," & _
              "calls_16=" & Nz(txtC16, 0) & ",calls_17=" & Nz(txtC17, 0) & "," & _
              "calls_18=" & Nz(txtC18, 0) & "," & _
              "valid_count=" & Nz(txtValid, 0) & ",prospect=" & Nz(txtProspect, 0) & "," & _
              "doc=" & Nz(txtDoc, 0) & ",follow_up=" & Nz(txtFollow, 0) & "," & _
              "received=" & Nz(txtReceived, 0) & ",work_hours=" & Nz(txtWorkHours, 8) & "," & _
              "note='" & Replace(Nz(txtNote, ""), "'", "''") & "'," & _
              "referral=" & Nz(txtReferral, 0) & "," & _
              "work_day=" & IIf(chkWorkDay.Value, "True", "False") & " " & _
              "WHERE ID = " & m_RecID
    End If

    CurrentDb.Execute sql, dbFailOnError
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub
"""


def _get_vba_form_members():
    return """
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    Me.Caption = "担当者管理"
    LoadData
End Sub

Private Sub LoadData()
    lstActive.RowSource = "SELECT ID, member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    lstInactive.RowSource = "SELECT ID, member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name"
    lstActive.Requery
    lstInactive.Requery
End Sub

Private Sub btnAdd_Click()
    Dim nm As String: nm = Trim(Nz(txtNewName.Value, ""))
    If nm = "" Then MsgBox "名前を入力してください", vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_MEMBERS (member_name, active) VALUES ('" & _
        Replace(nm, "'", "''") & "', True)", dbFailOnError
    txtNewName.Value = ""
    LoadData
End Sub

Private Sub btnDeactivate_Click()
    If IsNull(lstActive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active = False WHERE ID = " & lstActive.Value, dbFailOnError
    LoadData
End Sub

Private Sub btnActivate_Click()
    If IsNull(lstInactive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active = True WHERE ID = " & lstInactive.Value, dbFailOnError
    LoadData
End Sub
"""


def _get_vba_form_targets():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 目標設定"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    lstTargets.RowSource = "SELECT T.ID, T.member_name AS 担当者, T.target_calls AS 架電, " & _
        "T.target_valid AS 有効, T.target_prospect AS 見込, T.target_received AS 受注, " & _
        "T.target_referral AS 送客, T.plan_days AS 日数 " & _
        "FROM T_MEMBER_TARGETS AS T " & _
        "WHERE T.target_year=" & m_Year & " AND T.target_month=" & m_Month & " " & _
        "ORDER BY T.member_name"
    lstTargets.Requery
    ClearInputs
End Sub

Private Sub ClearInputs()
    txtPlanDays.Value = WorkDaysInMonth(m_Year, m_Month)
    txtHoursPerDay.Value = 8
    txtTgtCalls.Value = 3000
    txtTgtValid.Value = 2000
    txtTgtProspect.Value = 100
    txtTgtReceived.Value = 10
    txtTgtReferral.Value = 120
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub btnSave_Click()
    If Nz(cboMember.Value, "") = "" Then
        MsgBox "担当者を選択してください", vbExclamation: Exit Sub
    End If

    Dim nm As String: nm = cboMember.Value
    ' UPSERT
    Dim cnt As Long
    cnt = DCount("*", "T_MEMBER_TARGETS", "member_name='" & nm & "' AND target_year=" & m_Year & " AND target_month=" & m_Month)

    If cnt > 0 Then
        CurrentDb.Execute "UPDATE T_MEMBER_TARGETS SET " & _
            "plan_days=" & Nz(txtPlanDays, 20) & "," & _
            "work_hours_per_day=" & Nz(txtHoursPerDay, 8) & "," & _
            "target_calls=" & Nz(txtTgtCalls, 0) & "," & _
            "target_valid=" & Nz(txtTgtValid, 0) & "," & _
            "target_prospect=" & Nz(txtTgtProspect, 0) & "," & _
            "target_received=" & Nz(txtTgtReceived, 0) & "," & _
            "target_referral=" & Nz(txtTgtReferral, 0) & " " & _
            "WHERE member_name='" & nm & "' AND target_year=" & m_Year & " AND target_month=" & m_Month, dbFailOnError
    Else
        CurrentDb.Execute "INSERT INTO T_MEMBER_TARGETS " & _
            "(member_name,target_year,target_month,plan_days,work_hours_per_day," & _
            "target_calls,target_valid,target_prospect,target_received,target_referral) VALUES (" & _
            "'" & nm & "'," & m_Year & "," & m_Month & "," & _
            Nz(txtPlanDays, 20) & "," & Nz(txtHoursPerDay, 8) & "," & _
            Nz(txtTgtCalls, 0) & "," & Nz(txtTgtValid, 0) & "," & _
            Nz(txtTgtProspect, 0) & "," & Nz(txtTgtReceived, 0) & "," & _
            Nz(txtTgtReferral, 0) & ")", dbFailOnError
    End If
    LoadData
    MsgBox nm & " の目標を保存しました", vbInformation
End Sub

Private Sub btnLoadPrev_Click()
    If Nz(cboMember.Value, "") = "" Then Exit Sub
    Dim py As Integer, pm As Integer
    py = m_Year: pm = m_Month
    PrevYM py, pm
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE " & _
        "member_name='" & cboMember.Value & "' AND target_year=" & py & " AND target_month=" & pm)
    If Not rs.EOF Then
        txtPlanDays.Value = rs("plan_days")
        txtHoursPerDay.Value = rs("work_hours_per_day")
        txtTgtCalls.Value = rs("target_calls")
        txtTgtValid.Value = rs("target_valid")
        txtTgtProspect.Value = rs("target_prospect")
        txtTgtReceived.Value = rs("target_received")
        txtTgtReferral.Value = rs("target_referral")
    Else
        MsgBox "前月の目標が見つかりません", vbInformation
    End If
    rs.Close
End Sub

Private Sub btnDelete_Click()
    If IsNull(lstTargets.Value) Then Exit Sub
    If MsgBox("この目標を削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & lstTargets.Value, dbFailOnError
        LoadData
    End If
End Sub
"""


def _get_vba_form_referrals():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 送客登録"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"
    txtRefDate.Value = Format(Date, "yyyy/mm/dd")

    lstRefs.RowSource = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd') AS 日付, " & _
        "R.member_name AS 担当者, R.ref_count AS 件数 " & _
        "FROM T_REFERRALS AS R " & _
        "WHERE R.rec_date >= #" & Format(DateSerial(m_Year, m_Month, 1), "yyyy/mm/dd") & "# " & _
        "AND R.rec_date < #" & Format(DateSerial(m_Year, m_Month + 1, 1), "yyyy/mm/dd") & "# " & _
        "ORDER BY R.rec_date DESC, R.member_name"
    lstRefs.Requery
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub btnAdd_Click()
    If Nz(cboMember.Value, "") = "" Or Nz(txtRefDate.Value, "") = "" Then
        MsgBox "日付と担当者を入力してください", vbExclamation: Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO T_REFERRALS (rec_date, member_name, ref_count) VALUES (" & _
        "#" & Format(CDate(txtRefDate.Value), "yyyy/mm/dd") & "#, " & _
        "'" & cboMember.Value & "', " & Nz(txtRefCount, 0) & ")", dbFailOnError
    LoadData
End Sub

Private Sub btnDelete_Click()
    If IsNull(lstRefs.Value) Then Exit Sub
    If MsgBox("この送客記録を削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_REFERRALS WHERE ID=" & lstRefs.Value, dbFailOnError
        LoadData
    End If
End Sub
"""


def _get_vba_form_report():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    LoadReport
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadReport
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadReport
End Sub

Private Sub btnPDF_Click()
    ExportPDF m_Year, m_Month
End Sub

Private Sub LoadReport()
    Me.Caption = m_Year & "年" & m_Month & "月 レポート"
    lblMonth.Caption = m_Year & "年" & m_Month & "月 営業月次レポート"

    Dim dtFrom As Date, dtTo As Date
    dtFrom = DateFrom(m_Year, m_Month)
    dtTo = DateTo(m_Year, m_Month)

    ' チーム集計
    Dim rs As DAO.Recordset
    Set rs = RunParamQuery("Q_Team_Monthly_Sum", "prmDateFrom", dtFrom, "prmDateTo", dtTo)

    Dim tCalls As Long, tValid As Long, tProsp As Long, tRecv As Long, tHours As Double
    If Not rs.EOF Then
        tCalls = Nz(rs("sum_calls"), 0)
        tValid = Nz(rs("sum_valid"), 0)
        tProsp = Nz(rs("sum_prospect"), 0)
        tRecv = Nz(rs("sum_received"), 0)
        tHours = Nz(rs("sum_hours"), 0)
    End If
    rs.Close

    ' チーム目標
    Set rs = RunParamQuery("Q_Team_Monthly_Targets", "prmYear", m_Year, "prmMonth", m_Month)
    Dim gCalls As Long, gValid As Long, gProsp As Long, gRecv As Long
    If Not rs.EOF Then
        gCalls = Nz(rs("sum_tgt_calls"), 0)
        gValid = Nz(rs("sum_tgt_valid"), 0)
        gProsp = Nz(rs("sum_tgt_prospect"), 0)
        gRecv = Nz(rs("sum_tgt_received"), 0)
    End If
    rs.Close

    ' 前月
    Dim py As Integer, pm As Integer: py = m_Year: pm = m_Month
    PrevYM py, pm
    Set rs = RunParamQuery("Q_Team_Monthly_Sum", "prmDateFrom", DateFrom(py, pm), "prmDateTo", DateTo(py, pm))
    Dim pCalls As Long, pValid As Long, pProsp As Long, pRecv As Long
    If Not rs.EOF Then
        pCalls = Nz(rs("sum_calls"), 0): pValid = Nz(rs("sum_valid"), 0)
        pProsp = Nz(rs("sum_prospect"), 0): pRecv = Nz(rs("sum_received"), 0)
    End If
    rs.Close

    ' KPI表示
    lblCalls.Caption = FmtComma(tCalls) & " / " & FmtComma(gCalls)
    lblCallsPrev.Caption = "前月" & FmtComma(pCalls) & " " & Arrow(tCalls, pCalls)
    lblValid.Caption = FmtComma(tValid) & " / " & FmtComma(gValid)
    lblValidPrev.Caption = "前月" & FmtComma(pValid) & " " & Arrow(tValid, pValid)
    lblProsp.Caption = FmtComma(tProsp) & " / " & FmtComma(gProsp)
    lblProspPrev.Caption = "前月" & FmtComma(pProsp) & " " & Arrow(tProsp, pProsp)
    lblRecv.Caption = FmtComma(tRecv) & " / " & FmtComma(gRecv)
    lblRecvPrev.Caption = "前月" & FmtComma(pRecv) & " " & Arrow(tRecv, pRecv)

    ' 有効率・受注率
    lblValidRate.Caption = FmtPct(tValid, tCalls)
    lblRecvRate.Caption = FmtPct(tRecv, tCalls)
    lblHours.Caption = Format(tHours, "#,##0") & "h"

    Dim prod As Double
    If tHours > 0 Then prod = Seika(tProsp, 0) / tHours Else prod = 0
    lblProductivity.Caption = Format(prod, "0.000")

    ' アラート
    Dim alertText As String: alertText = ""
    If gRecv > 0 Then
        Dim rr As Long: rr = gRecv - tRecv
        If rr > 0 Then
            alertText = alertText & Chr(9651) & " 受注 残り" & rr & "件 " & _
                FmtComma(tRecv) & "/" & FmtComma(gRecv) & vbCrLf
        Else
            alertText = alertText & Chr(9675) & " 受注 目標達成！" & vbCrLf
        End If
    End If
    If gProsp > 0 Then
        Dim pr As Long: pr = gProsp - tProsp
        If pr > 0 Then
            alertText = alertText & Chr(9651) & " 見込 残り" & pr & "件 " & _
                FmtComma(tProsp) & "/" & FmtComma(gProsp)
        Else
            alertText = alertText & Chr(9675) & " 見込 目標達成！"
        End If
    End If
    lblAlert.Caption = alertText

    ' ランキング (送客)
    Set rs = RunParamQuery("Q_Rank_Referral", "prmDateFrom", dtFrom, "prmDateTo", dtTo)
    lstRankRef.RowSource = ""
    lstRankRef.RowSource = "SELECT R.member_name, Sum(R.referral) AS sum_ref " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name " & _
        "WHERE M.active = True AND R.rec_date >= #" & Format(dtFrom, "yyyy/mm/dd") & "# " & _
        "AND R.rec_date < #" & Format(dtTo, "yyyy/mm/dd") & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    rs.Close

    ' ランキング (受注)
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received) AS sum_recv " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name " & _
        "WHERE M.active = True AND R.rec_date >= #" & Format(dtFrom, "yyyy/mm/dd") & "# " & _
        "AND R.rec_date < #" & Format(dtTo, "yyyy/mm/dd") & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.received) DESC"

    ' ランキング (見込)
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) AS sum_prosp " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name = M.member_name " & _
        "WHERE M.active = True AND R.rec_date >= #" & Format(dtFrom, "yyyy/mm/dd") & "# " & _
        "AND R.rec_date < #" & Format(dtTo, "yyyy/mm/dd") & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"

    lstRankRef.Requery
    lstRankRecv.Requery
    lstRankProsp.Requery
End Sub
"""


def _get_vba_form_progress():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    If cboMember.ListCount > 0 Then cboMember.Value = cboMember.ItemData(0)
    LoadData
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub LoadData()
    If Nz(cboMember.Value, "") = "" Then Exit Sub
    Me.Caption = m_Year & "年" & m_Month & "月 " & cboMember.Value & " 進捗"
    lblMonth.Caption = m_Year & "年" & m_Month & "月 進捗"

    Dim dtFrom As Date, dtTo As Date
    dtFrom = DateFrom(m_Year, m_Month): dtTo = DateTo(m_Year, m_Month)
    Dim nm As String: nm = cboMember.Value

    ' 実績
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(calls) AS sc, Sum(valid_count) AS sv, Sum(prospect) AS sp, " & _
        "Sum(received) AS sr, Sum(work_hours) AS sh, Sum(referral) AS srf, " & _
        "Sum(IIf(work_day,1,0)) AS wd " & _
        "FROM T_RECORDS WHERE member_name='" & nm & "' " & _
        "AND rec_date >= #" & Format(dtFrom, "yyyy/mm/dd") & "# " & _
        "AND rec_date < #" & Format(dtTo, "yyyy/mm/dd") & "#")

    Dim aCalls As Long, aValid As Long, aProsp As Long, aRecv As Long, aHours As Double, aRef As Long, aWD As Long
    If Not rs.EOF Then
        aCalls = Nz(rs("sc"), 0): aValid = Nz(rs("sv"), 0)
        aProsp = Nz(rs("sp"), 0): aRecv = Nz(rs("sr"), 0)
        aHours = Nz(rs("sh"), 0): aRef = Nz(rs("srf"), 0)
        aWD = Nz(rs("wd"), 0)
    End If
    rs.Close

    ' 目標
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" & nm & "' " & _
        "AND target_year=" & m_Year & " AND target_month=" & m_Month)
    Dim gCalls As Long, gValid As Long, gProsp As Long, gRecv As Long, gRef As Long, planD As Long
    If Not rs.EOF Then
        gCalls = Nz(rs("target_calls"), 0): gValid = Nz(rs("target_valid"), 0)
        gProsp = Nz(rs("target_prospect"), 0): gRecv = Nz(rs("target_received"), 0)
        gRef = Nz(rs("target_referral"), 0): planD = Nz(rs("plan_days"), 20)
    Else
        planD = WorkDaysInMonth(m_Year, m_Month)
    End If
    rs.Close

    ' KPI表示
    lblCalls.Caption = FmtComma(aCalls) & " / " & FmtComma(gCalls)
    lblCallsPct.Caption = FmtPct(aCalls, gCalls)
    lblValid.Caption = FmtComma(aValid) & " / " & FmtComma(gValid)
    lblValidPct.Caption = FmtPct(aValid, gValid)
    lblProsp.Caption = FmtComma(aProsp) & " / " & FmtComma(gProsp)
    lblProspPct.Caption = FmtPct(aProsp, gProsp)
    lblRecv.Caption = FmtComma(aRecv) & " / " & FmtComma(gRecv)
    lblRecvPct.Caption = FmtPct(aRecv, gRecv)
    lblRef.Caption = FmtComma(aRef) & " / " & FmtComma(gRef)
    lblRefPct.Caption = FmtPct(aRef, gRef)
    lblHours.Caption = Format(aHours, "0.0") & "h"
    lblDays.Caption = aWD & " / " & planD & "日"

    Dim prod As Double
    If aHours > 0 Then prod = Seika(aProsp, 0) / aHours Else prod = 0
    lblProductivity.Caption = Format(prod, "0.000")
End Sub
"""


def _get_vba_form_ranking():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    LoadData
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 ランキング"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    Dim dtF As String, dtT As String
    dtF = Format(DateFrom(m_Year, m_Month), "yyyy/mm/dd")
    dtT = Format(DateTo(m_Year, m_Month), "yyyy/mm/dd")

    lstRef.RowSource = "SELECT R.member_name, Sum(R.referral) " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received) " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"

    lstRef.Requery: lstRecv.Requery: lstProsp.Requery
End Sub
"""


def _get_vba_form_list():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub btnCSV_Click()
    ExportCSV m_Year, m_Month
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 履歴一覧"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    Dim dtF As String, dtT As String
    dtF = Format(DateFrom(m_Year, m_Month), "yyyy/mm/dd")
    dtT = Format(DateTo(m_Year, m_Month), "yyyy/mm/dd")

    Dim sql As String
    sql = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd') AS 日付, R.member_name AS 担当者, " & _
        "R.calls AS 架電, R.valid_count AS 有効, R.prospect AS 見込, " & _
        "R.received AS 受注, R.work_hours AS 稼働h " & _
        "FROM T_RECORDS AS R " & _
        "WHERE R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# "
    If Nz(cboMember.Value, "") <> "" Then
        sql = sql & "AND R.member_name='" & cboMember.Value & "' "
    End If
    sql = sql & "ORDER BY R.rec_date DESC, R.member_name"
    lstRecords.RowSource = sql
    lstRecords.Requery
End Sub
"""


def _get_vba_form_analysis():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    LoadData
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 分析"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    Dim dtF As String, dtT As String
    dtF = Format(DateFrom(m_Year, m_Month), "yyyy/mm/dd")
    dtT = Format(DateTo(m_Year, m_Month), "yyyy/mm/dd")

    ' 受注ランキング
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received), " & _
        "IIf(Sum(R.calls)>0, Format(Sum(R.received)/Sum(R.calls)*100,'0.0') & '%', '-') " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.received) DESC"

    ' 見込ランキング
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect), " & _
        "IIf(Sum(R.valid_count)>0, Format(Sum(R.prospect)/Sum(R.valid_count)*100,'0.0') & '%', '-') " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"

    ' 生産性ランキング
    lstProd.RowSource = "SELECT R.member_name, " & _
        "Format(IIf(Sum(R.work_hours)>0,(Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours),0),'0.000') " & _
        "FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name " & _
        "WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# " & _
        "GROUP BY R.member_name HAVING Sum(R.work_hours)>0 " & _
        "ORDER BY IIf(Sum(R.work_hours)>0,(Sum(R.prospect)+Sum(R.follow_up)*0.5)/Sum(R.work_hours),0) DESC"

    lstRecv.Requery: lstProsp.Requery: lstProd.Requery
End Sub
"""


def _get_vba_form_reftrend():
    return """
Option Compare Database
Option Explicit

Private m_Year As Integer
Private m_Month As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Year = Year(Date): m_Month = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub btnPrev_Click()
    PrevYM m_Year, m_Month: LoadData
End Sub

Private Sub btnNext_Click()
    NextYM m_Year, m_Month: LoadData
End Sub

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Year & "年" & m_Month & "月 送客推移"
    lblMonth.Caption = m_Year & "年" & m_Month & "月"

    ' 12ヶ月推移
    Dim y12 As Integer, m12 As Integer
    y12 = m_Year: m12 = m_Month
    Dim i As Integer
    For i = 1 To 11
        PrevYM y12, m12
    Next i
    Dim dt12 As String: dt12 = Format(DateFrom(y12, m12), "yyyy/mm/dd")
    Dim dtEnd As String: dtEnd = Format(DateTo(m_Year, m_Month), "yyyy/mm/dd")

    Dim sql As String
    sql = "SELECT Year(rec_date) & '/' & Format(Month(rec_date),'00') AS 月, " & _
          "Sum(ref_count) AS 件数 FROM T_REFERRALS " & _
          "WHERE rec_date>=#" & dt12 & "# AND rec_date<#" & dtEnd & "# "
    If Nz(cboMember.Value, "") <> "" Then
        sql = sql & "AND member_name='" & cboMember.Value & "' "
    End If
    sql = sql & "GROUP BY Year(rec_date), Month(rec_date) " & _
          "ORDER BY Year(rec_date), Month(rec_date)"
    lstTrend.RowSource = sql
    lstTrend.Requery
End Sub
"""


# ============================================================
# フォーム作成 (COM自動化)
# ============================================================
# Access Form/Control constants (twips: 1 inch = 1440 twips)
acDetail = 0
acHeader = 1
acFooter = 2
acLabel = 100
acTextBox = 109
acComboBox = 111
acCommandButton = 104
acListBox = 110
acCheckBox = 106
acTabCtl = 123
acPage = 124
acSubForm = 112

TWIP = 1440  # 1 inch
CM = 567     # 1 cm


def _create_label(app, frm_name, section, text, left, top, width, height,
                  font_size=9, bold=False, fore_color=0, back_color=-1, name=None):
    ctrl = app.CreateControl(frm_name, acLabel, section, "", "", left, top, width, height)
    ctrl.Caption = text
    ctrl.FontName = "Yu Gothic UI"
    ctrl.FontSize = font_size
    ctrl.FontBold = bold
    if fore_color != 0:
        ctrl.ForeColor = fore_color
    if back_color >= 0:
        ctrl.BackColor = back_color
        ctrl.BackStyle = 1  # Normal (opaque)
    if name:
        ctrl.Name = name
    return ctrl


def _create_button(app, frm_name, section, caption, left, top, width, height, name=None):
    ctrl = app.CreateControl(frm_name, acCommandButton, section, "", "", left, top, width, height)
    ctrl.Caption = caption
    ctrl.FontName = "Yu Gothic UI"
    ctrl.FontSize = 9
    if name:
        ctrl.Name = name
    return ctrl


def _create_textbox(app, frm_name, section, left, top, width, height, name=None, default_val=None):
    ctrl = app.CreateControl(frm_name, acTextBox, section, "", "", left, top, width, height)
    ctrl.FontName = "Yu Gothic UI"
    ctrl.FontSize = 9
    if name:
        ctrl.Name = name
    if default_val is not None:
        ctrl.DefaultValue = str(default_val)
    return ctrl


def _create_combo(app, frm_name, section, left, top, width, height, name=None, row_source=""):
    ctrl = app.CreateControl(frm_name, acComboBox, section, "", "", left, top, width, height)
    ctrl.FontName = "Yu Gothic UI"
    ctrl.FontSize = 9
    if name:
        ctrl.Name = name
    if row_source:
        ctrl.RowSource = row_source
    return ctrl


def _create_listbox(app, frm_name, section, left, top, width, height, name=None, col_count=2, col_widths=""):
    ctrl = app.CreateControl(frm_name, acListBox, section, "", "", left, top, width, height)
    ctrl.FontName = "Yu Gothic UI"
    ctrl.FontSize = 9
    ctrl.ColumnCount = col_count
    if col_widths:
        ctrl.ColumnWidths = col_widths
    if name:
        ctrl.Name = name
    return ctrl


def _create_checkbox(app, frm_name, section, left, top, width, height, name=None):
    ctrl = app.CreateControl(frm_name, acCheckBox, section, "", "", left, top, width, height)
    if name:
        ctrl.Name = name
    return ctrl


def _month_nav_controls(app, frm_name, top_y):
    """共通の月ナビゲーションコントロール（◀ YYYY年MM月 ▶）"""
    _create_button(app, frm_name, acDetail, "◀", CM*0.5, top_y, CM*1.5, CM*0.7, "btnPrev")
    _create_label(app, frm_name, acDetail, "YYYY年MM月", CM*2.2, top_y, CM*6, CM*0.7,
                  font_size=12, bold=True, name="lblMonth")
    _create_button(app, frm_name, acDetail, "▶", CM*8.5, top_y, CM*1.5, CM*0.7, "btnNext")


def _save_form(app, frm, form_name):
    """フォームを保存"""
    app.DoCmd.Save(2, form_name)  # 2 = acForm
    app.DoCmd.Close(2, form_name)  # 2 = acForm
    time.sleep(0.3)


def _create_form_main(app):
    """F_Main: ナビゲーション"""
    frm = app.CreateForm()
    frm.Caption = "SalesMgr 営業管理"
    frm.DefaultView = 0  # Single Form
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0  # Neither

    frm_name = frm.Name

    # ヘッダーラベル
    _create_label(app, frm_name, acDetail, "SalesMgr 営業管理", CM*1, CM*0.5, CM*14, CM*1.5,
                  font_size=18, bold=True, fore_color=3877662)

    # ボタン 2列×5行
    buttons = [
        ("btnDaily", "📋 日次登録"),
        ("btnTargets", "🎯 目標登録"),
        ("btnReferrals", "📤 送客登録"),
        ("btnReport", "📊 レポート"),
        ("btnProgress", "📈 進捗"),
        ("btnList", "📝 履歴一覧"),
        ("btnAnalysis", "🔍 分析"),
        ("btnRanking", "🏆 ランキング"),
        ("btnRefTrend", "📉 送客推移"),
        ("btnMembers", "👥 担当者管理"),
    ]

    for i, (bname, bcaption) in enumerate(buttons):
        col = i % 2
        row = i // 2
        x = CM * 1 + col * CM * 7.5
        y = CM * 2.5 + row * CM * 1.8
        btn = _create_button(app, frm_name, acDetail, bcaption, x, y, CM*6.5, CM*1.3, bname)
        btn.FontSize = 11

    # Rename form
    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Main", 2, frm_name)


def _create_form_daily(app):
    """F_Daily: 日次一覧"""
    frm = app.CreateForm()
    frm.Caption = "日次一覧"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 2  # Vertical

    frm_name = frm.Name

    # 月ナビ
    _month_nav_controls(app, frm_name, CM*0.5)

    # メンバーフィルタ
    _create_label(app, frm_name, acDetail, "担当者:", CM*10.5, CM*0.5, CM*2, CM*0.7)
    _create_combo(app, frm_name, acDetail, CM*12.5, CM*0.5, CM*4, CM*0.7, "cboMember")

    # ボタン
    _create_button(app, frm_name, acDetail, "新規登録", CM*0.5, CM*1.5, CM*3, CM*0.8, "btnAdd")
    _create_button(app, frm_name, acDetail, "編集", CM*4, CM*1.5, CM*2.5, CM*0.8, "btnEdit")
    _create_button(app, frm_name, acDetail, "削除", CM*7, CM*1.5, CM*2.5, CM*0.8, "btnDelete")

    # レコードリスト
    lst = _create_listbox(app, frm_name, acDetail, CM*0.5, CM*2.8, CM*18, CM*12,
                          "lstRecords", col_count=11,
                          col_widths="0cm;2.5cm;2.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;3cm")
    lst.ColumnHeads = True

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Daily", 2, frm_name)


def _create_form_daily_edit(app):
    """F_DailyEdit: 日次登録/編集ダイアログ"""
    frm = app.CreateForm()
    frm.Caption = "日次レコード"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0
    frm.PopUp = True
    frm.Modal = True

    frm_name = frm.Name

    y = CM * 0.5

    # 日付
    _create_label(app, frm_name, acDetail, "日付:", CM*0.5, y, CM*2, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*3, y, CM*4, CM*0.6, "txtRecDate")
    y += CM * 0.9

    # 担当者
    _create_label(app, frm_name, acDetail, "担当者:", CM*0.5, y, CM*2, CM*0.6)
    _create_combo(app, frm_name, acDetail, CM*3, y, CM*4, CM*0.6, "cboMember")
    y += CM * 0.9

    # 稼働
    _create_label(app, frm_name, acDetail, "稼働時間:", CM*0.5, y, CM*2, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*3, y, CM*2, CM*0.6, "txtWorkHours")
    _create_label(app, frm_name, acDetail, "出勤:", CM*6, y, CM*1.5, CM*0.6)
    _create_checkbox(app, frm_name, acDetail, CM*7.5, y, CM*0.5, CM*0.5, "chkWorkDay")
    y += CM * 1.2

    # 時間帯別架電
    _create_label(app, frm_name, acDetail, "時間帯別架電:", CM*0.5, y, CM*4, CM*0.6,
                  font_size=9, bold=True)
    y += CM * 0.7

    for hour in range(10, 19):
        _create_label(app, frm_name, acDetail, f"{hour}時:", CM*0.5, y, CM*1.5, CM*0.6)
        _create_textbox(app, frm_name, acDetail, CM*2.2, y, CM*2, CM*0.6, f"txtC{hour}", 0)
        y += CM * 0.7

    y += CM * 0.3

    # 成果
    for lbl, nm in [("有効:", "txtValid"), ("見込:", "txtProspect"), ("資料:", "txtDoc"),
                    ("追客:", "txtFollow"), ("受注:", "txtReceived"), ("送客:", "txtReferral")]:
        _create_label(app, frm_name, acDetail, lbl, CM*0.5, y, CM*2, CM*0.6)
        _create_textbox(app, frm_name, acDetail, CM*3, y, CM*2, CM*0.6, nm, 0)
        y += CM * 0.7

    # 備考
    _create_label(app, frm_name, acDetail, "備考:", CM*0.5, y, CM*2, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*3, y, CM*5, CM*1, "txtNote")
    y += CM * 1.5

    # ボタン
    _create_button(app, frm_name, acDetail, "保存", CM*1, y, CM*3, CM*0.8, "btnSave")
    _create_button(app, frm_name, acDetail, "キャンセル", CM*5, y, CM*3, CM*0.8, "btnCancel")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_DailyEdit", 2, frm_name)


def _create_form_members(app):
    """F_Members: 担当者管理"""
    frm = app.CreateForm()
    frm.Caption = "担当者管理"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _create_label(app, frm_name, acDetail, "担当者管理", CM*0.5, CM*0.3, CM*8, CM*1,
                  font_size=14, bold=True)

    # 追加
    _create_label(app, frm_name, acDetail, "新規担当者:", CM*0.5, CM*1.5, CM*3, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*3.5, CM*1.5, CM*4, CM*0.6, "txtNewName")
    _create_button(app, frm_name, acDetail, "追加", CM*8, CM*1.5, CM*2, CM*0.6, "btnAdd")

    # アクティブ一覧
    _create_label(app, frm_name, acDetail, "有効な担当者:", CM*0.5, CM*2.8, CM*4, CM*0.6,
                  bold=True)
    _create_listbox(app, frm_name, acDetail, CM*0.5, CM*3.5, CM*6, CM*6, "lstActive", 2, "0cm;4cm")
    _create_button(app, frm_name, acDetail, "無効にする ▶", CM*7, CM*5, CM*3, CM*0.8, "btnDeactivate")

    # 無効一覧
    _create_label(app, frm_name, acDetail, "無効な担当者:", CM*10.5, CM*2.8, CM*4, CM*0.6,
                  bold=True)
    _create_listbox(app, frm_name, acDetail, CM*10.5, CM*3.5, CM*6, CM*6, "lstInactive", 2, "0cm;4cm")
    _create_button(app, frm_name, acDetail, "◀ 有効にする", CM*7, CM*6.5, CM*3, CM*0.8, "btnActivate")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Members", 2, frm_name)


def _create_form_targets(app):
    """F_Targets: 月次目標設定"""
    frm = app.CreateForm()
    frm.Caption = "目標設定"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.5)

    y = CM * 1.8
    _create_label(app, frm_name, acDetail, "担当者:", CM*0.5, y, CM*2, CM*0.6)
    _create_combo(app, frm_name, acDetail, CM*3, y, CM*4, CM*0.6, "cboMember")
    _create_button(app, frm_name, acDetail, "前月コピー", CM*8, y, CM*3, CM*0.6, "btnLoadPrev")
    y += CM * 1

    for lbl, nm in [("稼働日数:", "txtPlanDays"), ("日時間:", "txtHoursPerDay"),
                    ("架電目標:", "txtTgtCalls"), ("有効目標:", "txtTgtValid"),
                    ("見込目標:", "txtTgtProspect"), ("受注目標:", "txtTgtReceived"),
                    ("送客目標:", "txtTgtReferral")]:
        _create_label(app, frm_name, acDetail, lbl, CM*0.5, y, CM*2.5, CM*0.6)
        _create_textbox(app, frm_name, acDetail, CM*3, y, CM*3, CM*0.6, nm)
        y += CM * 0.8

    _create_button(app, frm_name, acDetail, "保存", CM*0.5, y + CM*0.3, CM*3, CM*0.8, "btnSave")
    _create_button(app, frm_name, acDetail, "削除", CM*4, y + CM*0.3, CM*3, CM*0.8, "btnDelete")

    # 一覧
    _create_listbox(app, frm_name, acDetail, CM*0.5, y + CM*1.5, CM*18, CM*5,
                    "lstTargets", 8, "0cm;2.5cm;2cm;2cm;2cm;2cm;2cm;2cm")
    _create_label(app, frm_name, acDetail, "", CM*0.5, y + CM*1.5, CM*1, CM*0.01)  # spacer

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Targets", 2, frm_name)


def _create_form_referrals(app):
    """F_Referrals: 送客登録"""
    frm = app.CreateForm()
    frm.Caption = "送客登録"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.5)

    y = CM * 1.8
    _create_label(app, frm_name, acDetail, "日付:", CM*0.5, y, CM*1.5, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*2, y, CM*3, CM*0.6, "txtRefDate")
    _create_label(app, frm_name, acDetail, "担当者:", CM*5.5, y, CM*2, CM*0.6)
    _create_combo(app, frm_name, acDetail, CM*7.5, y, CM*3.5, CM*0.6, "cboMember")
    _create_label(app, frm_name, acDetail, "件数:", CM*11.5, y, CM*1.5, CM*0.6)
    _create_textbox(app, frm_name, acDetail, CM*13, y, CM*2, CM*0.6, "txtRefCount", 1)
    _create_button(app, frm_name, acDetail, "追加", CM*15.5, y, CM*2, CM*0.6, "btnAdd")
    y += CM * 1

    _create_button(app, frm_name, acDetail, "削除", CM*0.5, y, CM*2.5, CM*0.7, "btnDelete")
    y += CM * 1

    _create_listbox(app, frm_name, acDetail, CM*0.5, y, CM*17, CM*8,
                    "lstRefs", 4, "0cm;3cm;3cm;2cm")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Referrals", 2, frm_name)


def _create_form_report(app):
    """F_Report: 月次レポート"""
    frm = app.CreateForm()
    frm.Caption = "レポート"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 2

    frm_name = frm.Name

    # 月ナビ
    _month_nav_controls(app, frm_name, CM*0.3)
    _create_button(app, frm_name, acDetail, "PDF出力", CM*11, CM*0.3, CM*3, CM*0.7, "btnPDF")

    y = CM * 1.3

    # アラート
    _create_label(app, frm_name, acDetail, "", CM*0.5, y, CM*18, CM*1.2,
                  font_size=9, name="lblAlert", fore_color=2498780)
    y += CM * 1.5

    # KPI カード - 1行目
    kpi_labels_row1 = [("架電件数", "lblCalls", "lblCallsPrev"),
                       ("有効件数", "lblValid", "lblValidPrev"),
                       ("見込件数", "lblProsp", "lblProspPrev"),
                       ("受注件数", "lblRecv", "lblRecvPrev")]
    for i, (title, nm_val, nm_sub) in enumerate(kpi_labels_row1):
        x = CM*0.5 + i * CM*4.7
        _create_label(app, frm_name, acDetail, title, x, y, CM*4.2, CM*0.5,
                      font_size=7, fore_color=6710886)
        _create_label(app, frm_name, acDetail, "-", x, y + CM*0.5, CM*4.2, CM*0.7,
                      font_size=11, bold=True, name=nm_val)
        _create_label(app, frm_name, acDetail, "", x, y + CM*1.2, CM*4.2, CM*0.5,
                      font_size=7, name=nm_sub, fore_color=10066329)
    y += CM * 2

    # KPI カード - 2行目
    kpi_labels_row2 = [("有効率", "lblValidRate"), ("受注率", "lblRecvRate"),
                       ("稼働時間", "lblHours"), ("生産性", "lblProductivity")]
    for i, (title, nm_val) in enumerate(kpi_labels_row2):
        x = CM*0.5 + i * CM*4.7
        _create_label(app, frm_name, acDetail, title, x, y, CM*4.2, CM*0.5,
                      font_size=7, fore_color=6710886)
        _create_label(app, frm_name, acDetail, "-", x, y + CM*0.5, CM*4.2, CM*0.7,
                      font_size=11, bold=True, name=nm_val)
    y += CM * 1.8

    # ランキング 3列
    _create_label(app, frm_name, acDetail, "送客ランキング", CM*0.5, y, CM*5.5, CM*0.6,
                  bold=True)
    _create_label(app, frm_name, acDetail, "受注ランキング", CM*6.5, y, CM*5.5, CM*0.6,
                  bold=True)
    _create_label(app, frm_name, acDetail, "見込ランキング", CM*12.5, y, CM*5.5, CM*0.6,
                  bold=True)
    y += CM * 0.7

    _create_listbox(app, frm_name, acDetail, CM*0.5, y, CM*5.5, CM*5,
                    "lstRankRef", 2, "3cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*6.5, y, CM*5.5, CM*5,
                    "lstRankRecv", 2, "3cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*12.5, y, CM*5.5, CM*5,
                    "lstRankProsp", 2, "3cm;2cm")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Report", 2, frm_name)


def _create_form_progress(app):
    """F_Progress: 個人進捗"""
    frm = app.CreateForm()
    frm.Caption = "進捗"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.3)

    _create_label(app, frm_name, acDetail, "担当者:", CM*11, CM*0.3, CM*2, CM*0.7)
    _create_combo(app, frm_name, acDetail, CM*13, CM*0.3, CM*4, CM*0.7, "cboMember")

    y = CM * 1.5
    kpis = [("架電", "lblCalls", "lblCallsPct"),
            ("有効", "lblValid", "lblValidPct"),
            ("見込", "lblProsp", "lblProspPct"),
            ("受注", "lblRecv", "lblRecvPct"),
            ("送客", "lblRef", "lblRefPct")]
    for i, (title, nm_val, nm_pct) in enumerate(kpis):
        x = CM*0.5 + i * CM*3.6
        _create_label(app, frm_name, acDetail, title, x, y, CM*3.2, CM*0.5,
                      font_size=8, bold=True)
        _create_label(app, frm_name, acDetail, "-", x, y + CM*0.5, CM*3.2, CM*0.7,
                      font_size=10, name=nm_val)
        _create_label(app, frm_name, acDetail, "-", x, y + CM*1.2, CM*3.2, CM*0.5,
                      font_size=8, name=nm_pct, fore_color=10066329)

    y += CM * 2.2
    # 稼働情報
    _create_label(app, frm_name, acDetail, "稼働時間:", CM*0.5, y, CM*2.5, CM*0.6)
    _create_label(app, frm_name, acDetail, "-", CM*3, y, CM*3, CM*0.6, name="lblHours")
    _create_label(app, frm_name, acDetail, "稼働日数:", CM*7, y, CM*2.5, CM*0.6)
    _create_label(app, frm_name, acDetail, "-", CM*9.5, y, CM*3, CM*0.6, name="lblDays")
    _create_label(app, frm_name, acDetail, "生産性:", CM*13, y, CM*2, CM*0.6)
    _create_label(app, frm_name, acDetail, "-", CM*15, y, CM*3, CM*0.6, name="lblProductivity")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Progress", 2, frm_name)


def _create_form_ranking(app):
    """F_Ranking: ランキング"""
    frm = app.CreateForm()
    frm.Caption = "ランキング"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.3)

    y = CM * 1.3
    _create_label(app, frm_name, acDetail, "送客", CM*0.5, y, CM*5.5, CM*0.6, bold=True)
    _create_label(app, frm_name, acDetail, "受注", CM*6.5, y, CM*5.5, CM*0.6, bold=True)
    _create_label(app, frm_name, acDetail, "見込", CM*12.5, y, CM*5.5, CM*0.6, bold=True)
    y += CM * 0.7

    _create_listbox(app, frm_name, acDetail, CM*0.5, y, CM*5.5, CM*7,
                    "lstRef", 2, "3cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*6.5, y, CM*5.5, CM*7,
                    "lstRecv", 2, "3cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*12.5, y, CM*5.5, CM*7,
                    "lstProsp", 2, "3cm;2cm")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Ranking", 2, frm_name)


def _create_form_list(app):
    """F_List: 履歴一覧"""
    frm = app.CreateForm()
    frm.Caption = "履歴一覧"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 2

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.3)
    _create_label(app, frm_name, acDetail, "担当者:", CM*10.5, CM*0.3, CM*2, CM*0.7)
    _create_combo(app, frm_name, acDetail, CM*12.5, CM*0.3, CM*4, CM*0.7, "cboMember")
    _create_button(app, frm_name, acDetail, "CSV出力", CM*17, CM*0.3, CM*2.5, CM*0.7, "btnCSV")

    _create_listbox(app, frm_name, acDetail, CM*0.5, CM*1.5, CM*19, CM*12,
                    "lstRecords", 8, "0cm;2.5cm;2.5cm;2cm;2cm;2cm;2cm;2cm")
    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_List", 2, frm_name)


def _create_form_analysis(app):
    """F_Analysis: 分析"""
    frm = app.CreateForm()
    frm.Caption = "分析"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.3)

    y = CM * 1.3
    _create_label(app, frm_name, acDetail, "受注ランキング", CM*0.5, y, CM*5.5, CM*0.6, bold=True)
    _create_label(app, frm_name, acDetail, "見込ランキング", CM*6.5, y, CM*5.5, CM*0.6, bold=True)
    _create_label(app, frm_name, acDetail, "生産性ランキング", CM*12.5, y, CM*5.5, CM*0.6, bold=True)
    y += CM * 0.7

    _create_listbox(app, frm_name, acDetail, CM*0.5, y, CM*5.5, CM*8,
                    "lstRecv", 3, "3cm;2cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*6.5, y, CM*5.5, CM*8,
                    "lstProsp", 3, "3cm;2cm;2cm")
    _create_listbox(app, frm_name, acDetail, CM*12.5, y, CM*5.5, CM*8,
                    "lstProd", 2, "3cm;3cm")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_Analysis", 2, frm_name)


def _create_form_reftrend(app):
    """F_RefTrend: 送客推移"""
    frm = app.CreateForm()
    frm.Caption = "送客推移"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False

    frm_name = frm.Name

    _month_nav_controls(app, frm_name, CM*0.3)
    _create_label(app, frm_name, acDetail, "担当者:", CM*10.5, CM*0.3, CM*2, CM*0.7)
    _create_combo(app, frm_name, acDetail, CM*12.5, CM*0.3, CM*4, CM*0.7, "cboMember")

    _create_label(app, frm_name, acDetail, "12ヶ月推移", CM*0.5, CM*1.3, CM*5, CM*0.6, bold=True)
    _create_listbox(app, frm_name, acDetail, CM*0.5, CM*2, CM*18, CM*8,
                    "lstTrend", 2, "4cm;3cm")

    _save_form(app, frm, frm_name)
    app.DoCmd.Rename("F_RefTrend", 2, frm_name)


def _create_all_forms(app, db):
    """全フォーム作成"""
    forms = [
        ("F_DailyEdit", _create_form_daily_edit),
        ("F_Members", _create_form_members),
        ("F_Targets", _create_form_targets),
        ("F_Referrals", _create_form_referrals),
        ("F_Daily", _create_form_daily),
        ("F_Progress", _create_form_progress),
        ("F_Analysis", _create_form_analysis),
        ("F_Ranking", _create_form_ranking),
        ("F_List", _create_form_list),
        ("F_RefTrend", _create_form_reftrend),
        ("F_Report", _create_form_report),
        ("F_Main", _create_form_main),
    ]
    for name, func in forms:
        print(f"    {name}...")
        try:
            func(app)
        except Exception as e:
            print(f"    ERROR creating {name}: {e}")
        time.sleep(0.5)


# ============================================================
# VBAモジュール注入
# ============================================================
def _export_vba_files():
    """VBAコードを .bas ファイルとして出力"""
    vba_dir = os.path.join(DESKTOP, "SalesMgr_VBA")
    os.makedirs(vba_dir, exist_ok=True)

    # modGlobal
    with open(os.path.join(vba_dir, "modGlobal.bas"), 'w', encoding='utf-8') as f:
        f.write("Attribute VB_Name = \"modGlobal\"\n")
        f.write(_get_vba_modGlobal())
    print("    modGlobal.bas")

    # フォーム用VBA
    form_vba_map = {
        "Form_F_Main": _get_vba_form_main(),
        "Form_F_Daily": _get_vba_form_daily(),
        "Form_F_DailyEdit": _get_vba_form_daily_edit(),
        "Form_F_Members": _get_vba_form_members(),
        "Form_F_Targets": _get_vba_form_targets(),
        "Form_F_Referrals": _get_vba_form_referrals(),
        "Form_F_Report": _get_vba_form_report(),
        "Form_F_Progress": _get_vba_form_progress(),
        "Form_F_Ranking": _get_vba_form_ranking(),
        "Form_F_List": _get_vba_form_list(),
        "Form_F_Analysis": _get_vba_form_analysis(),
        "Form_F_RefTrend": _get_vba_form_reftrend(),
    }
    for comp_name, code in form_vba_map.items():
        fname = comp_name + ".bas"
        with open(os.path.join(vba_dir, fname), 'w', encoding='utf-8') as f:
            f.write(f'Attribute VB_Name = "{comp_name}"\n')
            f.write(code)
        print(f"    {fname}")

    print(f"    出力先: {vba_dir}")


def _create_vba_modules(app):
    """VBAモジュールを追加 (COM経由)"""
    try:
        vbe = app.VBE
        proj = vbe.VBProjects(1)
    except Exception as e:
        print(f"    WARNING: VBE access denied: {e}")
        print("    → Access Trust Center で「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を有効にしてください")
        return

    # 標準モジュール: modGlobal (ファイルに書き出してインポート)
    import tempfile
    tmp_path = os.path.join(tempfile.gettempdir(), "modGlobal.bas")
    with open(tmp_path, 'w', encoding='utf-8') as f:
        f.write("Attribute VB_Name = \"modGlobal\"\n")
        # Option Compare/Explicit は _get_vba_modGlobal() に含まれている
        code = _get_vba_modGlobal()
        f.write(code)
    proj.VBComponents.Import(tmp_path)
    os.remove(tmp_path)
    print("    modGlobal")


def _inject_form_vba(app):
    """フォーム作成後にVBAコードを注入"""
    try:
        vbe = app.VBE
        proj = vbe.VBProjects(1)
    except:
        print("    WARNING: VBE access denied")
        return

    form_vba_map = {
        "Form_F_Main": _get_vba_form_main(),
        "Form_F_Daily": _get_vba_form_daily(),
        "Form_F_DailyEdit": _get_vba_form_daily_edit(),
        "Form_F_Members": _get_vba_form_members(),
        "Form_F_Targets": _get_vba_form_targets(),
        "Form_F_Referrals": _get_vba_form_referrals(),
        "Form_F_Report": _get_vba_form_report(),
        "Form_F_Progress": _get_vba_form_progress(),
        "Form_F_Ranking": _get_vba_form_ranking(),
        "Form_F_List": _get_vba_form_list(),
        "Form_F_Analysis": _get_vba_form_analysis(),
        "Form_F_RefTrend": _get_vba_form_reftrend(),
    }

    for comp_name, code in form_vba_map.items():
        try:
            comp = proj.VBComponents(comp_name)
            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            # AddFromString の代わりに InsertLines を使用
            lines = code.strip().split('\n')
            for i, line in enumerate(lines, 1):
                cm.InsertLines(i, line)
            print(f"    {comp_name}")
        except Exception as e:
            print(f"    WARNING: {comp_name}: {e}")


# ============================================================
# レポート作成
# ============================================================
def _create_reports(app, db):
    """R_MonthlyReport: PDF出力用レポート"""
    # レポートはVBAのExportPDFから呼ばれる
    # シンプルなレポートを作成 (詳細はVBAで制御)
    try:
        rpt = app.CreateReport()
        rpt.Caption = "営業月次レポート"
        rpt_name = rpt.Name

        # タイトルラベル
        ctrl = app.CreateReportControl(rpt_name, acLabel, acDetail, "", "",
                                       CM*0.5, CM*0.5, CM*18, CM*1)
        ctrl.Caption = "営業月次レポート"
        ctrl.FontName = "Yu Gothic UI"
        ctrl.FontSize = 16
        ctrl.FontBold = True

        app.DoCmd.Save(3, rpt_name)  # 3 = acReport
        app.DoCmd.Close(3, rpt_name)
        app.DoCmd.Rename("R_MonthlyReport", 3, rpt_name)
        print("    R_MonthlyReport")
    except Exception as e:
        print(f"    WARNING: Report creation: {e}")


# ============================================================
# スタートアップ設定
# ============================================================
def _set_startup(app, db):
    """スタートアップフォームを設定"""
    try:
        # StartUpForm property
        props = db.Properties
        try:
            props("StartUpForm").Value = "F_Main"
        except:
            prop = db.CreateProperty("StartUpForm", 10, "F_Main")  # 10 = dbText
            props.Append(prop)

        try:
            props("AppTitle").Value = "SalesMgr 営業管理"
        except:
            prop = db.CreateProperty("AppTitle", 10, "SalesMgr 営業管理")
            props.Append(prop)

        try:
            props("StartUpShowDBWindow").Value = False
        except:
            prop = db.CreateProperty("StartUpShowDBWindow", 1, False)  # 1 = dbBoolean
            props.Append(prop)

        print("  スタートアップ: F_Main, タイトル: SalesMgr 営業管理")
    except Exception as e:
        print(f"  WARNING: Startup settings: {e}")


# ============================================================
# メイン
# ============================================================
if __name__ == '__main__':
    print("SalesMgr Access アプリケーション ビルダー")
    print("=" * 60)
    print(f"SQLite: {SQLITE_PATH}")
    print(f"Backend: {BE_PATH}")
    print(f"Frontend: {FE_PATH}")
    print()

    if not os.path.exists(SQLITE_PATH):
        print(f"ERROR: SQLiteファイルが見つかりません: {SQLITE_PATH}")
        sys.exit(1)

    # まず Access プロセスを全て終了
    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(1)

    create_backend()
    create_frontend()

    # VBAコード注入 (フォーム作成後に再オープンして注入)
    print("\n" + "=" * 60)
    print("Phase 2b: VBAコード注入")
    print("=" * 60)

    try:
        import win32com.client
        app = win32com.client.Dispatch("Access.Application")
        app.Visible = False
        app.UserControl = False
        app.OpenCurrentDatabase(FE_PATH)
        time.sleep(2)

        _inject_form_vba(app)

        time.sleep(0.5)
        app.CloseCurrentDatabase()
        time.sleep(0.5)
        app.Quit()
        time.sleep(1)
    except Exception as e:
        print(f"  VBA注入スキップ: {e}")
        print(f"  → SalesMgr_VBA フォルダの .bas ファイルを手動でインポートしてください")
        try:
            app.Quit()
        except:
            pass

    print("\n" + "=" * 60)
    print("ビルド完了!")
    print("=" * 60)
    print(f"\nバックエンド: {BE_PATH}")
    print(f"フロントエンド: {FE_PATH}")
    print(f"\nSalesMgr_FE.accdb をダブルクリックで起動してください")
