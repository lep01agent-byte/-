# -*- coding: utf-8 -*-
"""VBA v2: 全バグ修正版 再注入"""
import os, time, win32com.client

BE_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

# ============================================================
# 日付ヘルパー関数（全フォーム共通で使う安全な月末計算）
# DateSerial(y,m+1,1) は m=12 で月13→エラーになる
# 修正: NextMonth 関数で安全に計算
# ============================================================

# F_Main: 変更なし（バグなし）
VBA_MAIN = """Option Compare Database
Option Explicit

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

Private Sub btnRanking_Click()
    DoCmd.OpenForm "F_Ranking"
End Sub

Private Sub btnMembers_Click()
    DoCmd.OpenForm "F_Members"
End Sub"""

# F_Members: 変更なし（バグなし）
VBA_MEMBERS = """Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    LoadData
End Sub

Private Sub LoadData()
    lstActive.RowSource = "SELECT ID, member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    lstInactive.RowSource = "SELECT ID, member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name"
    lstActive.Requery
    lstInactive.Requery
End Sub

Private Sub btnAdd_Click()
On Error GoTo EH
    Dim nm As String
    nm = Trim(Nz(txtNewName.Value, ""))
    If nm = "" Then
        MsgBox "名前を入力してください", vbExclamation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO T_MEMBERS (member_name, active) VALUES ('" & Replace(nm, "'", "''") & "', True)", dbFailOnError
    txtNewName.Value = ""
    LoadData
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnDeactivate_Click()
On Error GoTo EH
    If IsNull(lstActive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=False WHERE ID=" & lstActive.Value, dbFailOnError
    LoadData
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnActivate_Click()
On Error GoTo EH
    If IsNull(lstInactive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=True WHERE ID=" & lstInactive.Value, dbFailOnError
    LoadData
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub"""

# F_Daily: DateSerial修正 + エラーハンドリング
VBA_DAILY = """Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer

Private Function NextMonthFirst() As Date
    If mM = 12 Then
        NextMonthFirst = DateSerial(mY + 1, 1, 1)
    Else
        NextMonthFirst = DateSerial(mY, mM + 1, 1)
    End If
End Function

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date)
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    Me.Caption = mY & "年" & Format(mM, "00") & "月 日次一覧"
    lblMonth.Caption = mY & "年" & mM & "月"
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    dT = Format(NextMonthFirst(), "yyyy/mm/dd")

    Dim s As String
    s = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd'), R.member_name, R.calls, R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.[note]" & _
        " FROM T_RECORDS AS R" & _
        " WHERE R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#"

    If Nz(cboMember.Value, "") <> "" Then
        s = s & " AND R.member_name='" & cboMember.Value & "'"
    End If
    s = s & " ORDER BY R.rec_date DESC, R.member_name"

    lstRecords.RowSource = s
    lstRecords.Requery
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadData
End Sub

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub btnAdd_Click()
    DoCmd.OpenForm "F_DailyEdit", , , , , acDialog, "ADD|" & mY & "|" & mM
    LoadData
End Sub

Private Sub btnEdit_Click()
    If IsNull(lstRecords.Value) Then
        MsgBox "行を選択してください", vbExclamation
        Exit Sub
    End If
    DoCmd.OpenForm "F_DailyEdit", , , , , acDialog, "EDIT|" & lstRecords.Value
    LoadData
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstRecords.Value) Then
        MsgBox "行を選択してください", vbExclamation
        Exit Sub
    End If
    If MsgBox("削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_RECORDS WHERE ID=" & lstRecords.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub lstRecords_DblClick(Cancel As Integer)
    btnEdit_Click
End Sub"""

# F_DailyEdit: EDIT/ADDフロー修正 + INSERT文カンマ修正 + エラーハンドリング
VBA_DAILY_EDIT = """Option Compare Database
Option Explicit
Private mMode As String, mID As Long

Private Sub Form_Open(Cancel As Integer)
On Error GoTo EH
    Dim p() As String
    p = Split(Nz(Me.OpenArgs, "ADD"), "|")
    mMode = p(0)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"

    If mMode = "EDIT" And UBound(p) >= 1 Then
        mID = CLng(p(1))
        Me.Caption = "編集"
        LoadRecord
    Else
        mID = 0
        Me.Caption = "新規登録"
        If UBound(p) >= 2 Then
            txtRecDate.Value = Format(DateSerial(CInt(p(1)), CInt(p(2)), Day(Date)), "yyyy/mm/dd")
        Else
            txtRecDate.Value = Format(Date, "yyyy/mm/dd")
        End If
        txtWorkHours.Value = 8
        chkWorkDay.Value = True
    End If
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub LoadRecord()
On Error GoTo EH
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_RECORDS WHERE ID=" & mID)
    If Not rs.EOF Then
        txtRecDate.Value = Format(rs("rec_date"), "yyyy/mm/dd")
        cboMember.Value = rs("member_name")
        Dim h As Integer
        For h = 10 To 18
            Me("txtC" & h).Value = Nz(rs("calls_" & h), 0)
        Next h
        txtValid.Value = Nz(rs("valid_count"), 0)
        txtProspect.Value = Nz(rs("prospect"), 0)
        txtDoc.Value = Nz(rs("doc"), 0)
        txtFollow.Value = Nz(rs("follow_up"), 0)
        txtReceived.Value = Nz(rs("received"), 0)
        txtReferral.Value = Nz(rs("referral"), 0)
        txtWorkHours.Value = Nz(rs("work_hours"), 8)
        chkWorkDay.Value = Nz(rs("work_day"), False)
        txtNote.Value = Nz(rs("note"), "")
    End If
    rs.Close
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnSave_Click()
On Error GoTo EH
    If Nz(txtRecDate.Value, "") = "" Then MsgBox "日付を入力", vbExclamation: Exit Sub
    If Nz(cboMember.Value, "") = "" Then MsgBox "担当者を選択", vbExclamation: Exit Sub

    Dim dt As Date: dt = CDate(txtRecDate.Value)
    Dim totalCalls As Long, h As Integer
    totalCalls = 0
    For h = 10 To 18: totalCalls = totalCalls + Nz(Me("txtC" & h), 0): Next h

    Dim s As String
    If mID = 0 Then
        s = "INSERT INTO T_RECORDS (rec_date, member_name, calls" & _
            ", calls_10, calls_11, calls_12, calls_13, calls_14, calls_15, calls_16, calls_17, calls_18" & _
            ", valid_count, prospect, doc, follow_up, received, work_hours, [note], referral, work_day" & _
            ") VALUES (#" & Format(dt, "yyyy/mm/dd") & "#" & _
            ", '" & cboMember.Value & "'" & _
            ", " & totalCalls
        For h = 10 To 18: s = s & ", " & Nz(Me("txtC" & h), 0): Next h
        s = s & ", " & Nz(txtValid, 0) & _
            ", " & Nz(txtProspect, 0) & _
            ", " & Nz(txtDoc, 0) & _
            ", " & Nz(txtFollow, 0) & _
            ", " & Nz(txtReceived, 0) & _
            ", " & Nz(txtWorkHours, 8) & _
            ", '" & Replace(Nz(txtNote, ""), "'", "''") & "'" & _
            ", " & Nz(txtReferral, 0) & _
            ", " & IIf(chkWorkDay.Value, "True", "False") & _
            ")"
    Else
        s = "UPDATE T_RECORDS SET" & _
            " rec_date=#" & Format(dt, "yyyy/mm/dd") & "#" & _
            ", member_name='" & cboMember.Value & "'" & _
            ", calls=" & totalCalls
        For h = 10 To 18: s = s & ", calls_" & h & "=" & Nz(Me("txtC" & h), 0): Next h
        s = s & ", valid_count=" & Nz(txtValid, 0) & _
            ", prospect=" & Nz(txtProspect, 0) & _
            ", doc=" & Nz(txtDoc, 0) & _
            ", follow_up=" & Nz(txtFollow, 0) & _
            ", received=" & Nz(txtReceived, 0) & _
            ", work_hours=" & Nz(txtWorkHours, 8) & _
            ", [note]='" & Replace(Nz(txtNote, ""), "'", "''") & "'" & _
            ", referral=" & Nz(txtReferral, 0) & _
            ", work_day=" & IIf(chkWorkDay.Value, "True", "False") & _
            " WHERE ID=" & mID
    End If
    CurrentDb.Execute s, dbFailOnError
    DoCmd.Close acForm, Me.Name
    Exit Sub
EH: MsgBox "保存エラー: " & Err.Description & vbCrLf & vbCrLf & "SQL: " & Left(s, 500), vbCritical
End Sub

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub"""

# F_Targets: DateSerial修正 + エラーハンドリング
VBA_TARGETS = """Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    lstTargets.RowSource = "SELECT T.ID, T.member_name, T.target_calls, T.target_valid" & _
        ", T.target_prospect, T.target_received, T.target_referral, T.plan_days, T.work_hours_per_day" & _
        " FROM T_MEMBER_TARGETS AS T WHERE T.target_year=" & mY & " AND T.target_month=" & mM & _
        " ORDER BY T.member_name"
    lstTargets.Requery
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadData
End Sub

Private Sub cboMember_AfterUpdate()
On Error GoTo EH
    If Nz(cboMember.Value, "") = "" Then Exit Sub
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" & cboMember.Value & "' AND target_year=" & mY & " AND target_month=" & mM)
    If Not rs.EOF Then
        txtPlanDays.Value = rs("plan_days")
        txtHoursPerDay.Value = rs("work_hours_per_day")
        txtTgtCalls.Value = rs("target_calls")
        txtTgtValid.Value = rs("target_valid")
        txtTgtProspect.Value = rs("target_prospect")
        txtTgtReceived.Value = rs("target_received")
        txtTgtReferral.Value = rs("target_referral")
    Else
        txtPlanDays.Value = "": txtHoursPerDay.Value = ""
        txtTgtCalls.Value = "": txtTgtValid.Value = ""
        txtTgtProspect.Value = "": txtTgtReceived.Value = ""
        txtTgtReferral.Value = ""
    End If
    rs.Close
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnSave_Click()
On Error GoTo EH
    If Nz(cboMember.Value, "") = "" Then MsgBox "担当者を選択", vbExclamation: Exit Sub
    Dim nm As String: nm = cboMember.Value
    Dim n As Long: n = DCount("*", "T_MEMBER_TARGETS", "member_name='" & nm & "' AND target_year=" & mY & " AND target_month=" & mM)
    If n > 0 Then
        CurrentDb.Execute "UPDATE T_MEMBER_TARGETS SET" & _
            " plan_days=" & Nz(txtPlanDays, 20) & _
            ", work_hours_per_day=" & Nz(txtHoursPerDay, 8) & _
            ", target_calls=" & Nz(txtTgtCalls, 0) & _
            ", target_valid=" & Nz(txtTgtValid, 0) & _
            ", target_prospect=" & Nz(txtTgtProspect, 0) & _
            ", target_received=" & Nz(txtTgtReceived, 0) & _
            ", target_referral=" & Nz(txtTgtReferral, 0) & _
            " WHERE member_name='" & nm & "' AND target_year=" & mY & " AND target_month=" & mM, dbFailOnError
    Else
        CurrentDb.Execute "INSERT INTO T_MEMBER_TARGETS (member_name, target_year, target_month" & _
            ", plan_days, work_hours_per_day, target_calls, target_valid, target_prospect, target_received, target_referral" & _
            ") VALUES ('" & nm & "', " & mY & ", " & mM & _
            ", " & Nz(txtPlanDays, 20) & ", " & Nz(txtHoursPerDay, 8) & _
            ", " & Nz(txtTgtCalls, 0) & ", " & Nz(txtTgtValid, 0) & _
            ", " & Nz(txtTgtProspect, 0) & ", " & Nz(txtTgtReceived, 0) & _
            ", " & Nz(txtTgtReferral, 0) & ")", dbFailOnError
    End If
    LoadData
    MsgBox nm & " 保存完了", vbInformation
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnLoadPrev_Click()
On Error GoTo EH
    If Nz(cboMember.Value, "") = "" Then Exit Sub
    Dim py As Integer, pm As Integer
    py = mY: pm = mM
    If pm = 1 Then py = py - 1: pm = 12 Else pm = pm - 1
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" & cboMember.Value & "' AND target_year=" & py & " AND target_month=" & pm)
    If Not rs.EOF Then
        txtPlanDays.Value = rs("plan_days")
        txtHoursPerDay.Value = rs("work_hours_per_day")
        txtTgtCalls.Value = rs("target_calls")
        txtTgtValid.Value = rs("target_valid")
        txtTgtProspect.Value = rs("target_prospect")
        txtTgtReceived.Value = rs("target_received")
        txtTgtReferral.Value = rs("target_referral")
    Else
        MsgBox "前月目標なし", vbInformation
    End If
    rs.Close
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstTargets.Value) Then Exit Sub
    If MsgBox("削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & lstTargets.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub"""

# F_Referrals: DateSerial修正 + エラーハンドリング
VBA_REFERRALS = """Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer

Private Function NextMonthFirst() As Date
    If mM = 12 Then
        NextMonthFirst = DateSerial(mY + 1, 1, 1)
    Else
        NextMonthFirst = DateSerial(mY, mM + 1, 1)
    End If
End Function

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    txtRefDate.Value = Format(Date, "yyyy/mm/dd")
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    dT = Format(NextMonthFirst(), "yyyy/mm/dd")
    lstRefs.RowSource = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd'), R.member_name, R.ref_count" & _
        " FROM T_REFERRALS AS R WHERE R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " ORDER BY R.rec_date DESC, R.member_name"
    lstRefs.Requery
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadData
End Sub

Private Sub btnAdd_Click()
On Error GoTo EH
    If Nz(cboMember.Value, "") = "" Or Nz(txtRefDate.Value, "") = "" Then
        MsgBox "入力してください", vbExclamation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO T_REFERRALS (rec_date, member_name, ref_count) VALUES (#" & _
        Format(CDate(txtRefDate.Value), "yyyy/mm/dd") & "#, '" & cboMember.Value & "', " & Nz(txtRefCount, 0) & ")", dbFailOnError
    LoadData
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstRefs.Value) Then Exit Sub
    If MsgBox("削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_REFERRALS WHERE ID=" & lstRefs.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub"""

# F_Report: アラート構文修正 + DateSerial修正 + エラーハンドリング
VBA_REPORT = """Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer

Private Function NextMonthFirst() As Date
    If mM = 12 Then
        NextMonthFirst = DateSerial(mY + 1, 1, 1)
    Else
        NextMonthFirst = DateSerial(mY, mM + 1, 1)
    End If
End Function

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date)
    LoadReport
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadReport
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadReport
End Sub

Private Sub btnPDF_Click()
    MsgBox "PDF出力は今後対応予定です", vbInformation
End Sub

Private Sub LoadReport()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月 レポート"

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    dT = Format(NextMonthFirst(), "yyyy/mm/dd")

    ' --- チーム実績 ---
    Dim qd As DAO.QueryDef, rs As DAO.Recordset
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(mY, mM, 1)
    qd.Parameters("prmDateTo") = NextMonthFirst()
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim tC As Long, tV As Long, tP As Long, tR As Long, tH As Double
    tC = 0: tV = 0: tP = 0: tR = 0: tH = 0
    If Not rs.EOF Then
        tC = Nz(rs("sum_calls"), 0)
        tV = Nz(rs("sum_valid"), 0)
        tP = Nz(rs("sum_prospect"), 0)
        tR = Nz(rs("sum_received"), 0)
        tH = Nz(rs("sum_hours"), 0)
    End If
    rs.Close

    ' --- チーム目標 ---
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear") = mY
    qd.Parameters("prmMonth") = mM
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim gC As Long, gV As Long, gP As Long, gR As Long
    gC = 0: gV = 0: gP = 0: gR = 0
    If Not rs.EOF Then
        gC = Nz(rs("sum_tgt_calls"), 0)
        gV = Nz(rs("sum_tgt_valid"), 0)
        gP = Nz(rs("sum_tgt_prospect"), 0)
        gR = Nz(rs("sum_tgt_received"), 0)
    End If
    rs.Close

    ' --- 前月実績 ---
    Dim py As Integer, pm As Integer
    py = mY: pm = mM
    If pm = 1 Then py = py - 1: pm = 12 Else pm = pm - 1

    Dim pNextY As Integer, pNextM As Integer
    pNextY = py: pNextM = pm
    If pNextM = 12 Then pNextY = pNextY + 1: pNextM = 1 Else pNextM = pNextM + 1

    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(py, pm, 1)
    qd.Parameters("prmDateTo") = DateSerial(pNextY, pNextM, 1)
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim pC As Long, pV As Long, pP As Long, pR As Long
    pC = 0: pV = 0: pP = 0: pR = 0
    If Not rs.EOF Then
        pC = Nz(rs("sum_calls"), 0)
        pV = Nz(rs("sum_valid"), 0)
        pP = Nz(rs("sum_prospect"), 0)
        pR = Nz(rs("sum_received"), 0)
    End If
    rs.Close

    ' --- KPI表示 ---
    lblCalls.Caption = Format(tC, "#,##0") & " / " & Format(gC, "#,##0")
    lblValid.Caption = Format(tV, "#,##0") & " / " & Format(gV, "#,##0")
    lblProsp.Caption = Format(tP, "#,##0") & " / " & Format(gP, "#,##0")
    lblRecv.Caption = Format(tR, "#,##0") & " / " & Format(gR, "#,##0")

    lblCallsPrev.Caption = "前月" & Format(pC, "#,##0") & " " & ArrowStr(tC, pC)
    lblValidPrev.Caption = "前月" & Format(pV, "#,##0") & " " & ArrowStr(tV, pV)
    lblProspPrev.Caption = "前月" & Format(pP, "#,##0") & " " & ArrowStr(tP, pP)
    lblRecvPrev.Caption = "前月" & Format(pR, "#,##0") & " " & ArrowStr(tR, pR)

    lblValidRate.Caption = IIf(tC > 0, Format(tV / tC * 100, "0.0") & "%", "-")
    lblRecvRate.Caption = IIf(tC > 0, Format(tR / tC * 100, "0.0") & "%", "-")
    lblHours.Caption = Format(tH, "#,##0") & "h"
    lblProductivity.Caption = IIf(tH > 0, Format(tP / tH, "0.000"), "-")

    ' --- アラート ---
    Dim al As String: al = ""
    If gR > 0 Then
        If tR < gR Then
            al = al & Chr(9651) & " 受注 残り" & (gR - tR) & "件 " & tR & "/" & gR & vbCrLf
        Else
            al = al & Chr(9675) & " 受注 目標達成！" & vbCrLf
        End If
    End If
    If gP > 0 Then
        If tP < gP Then
            al = al & Chr(9651) & " 見込 残り" & (gP - tP) & "件 " & tP & "/" & gP
        Else
            al = al & Chr(9675) & " 見込 目標達成！"
        End If
    End If
    lblAlert.Caption = al

    ' --- ランキング ---
    lstRankRef.RowSource = "SELECT R.member_name, Sum(R.referral)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRankRef.Requery
    lstRankRecv.Requery
    lstRankProsp.Requery
    Exit Sub
EH: MsgBox "レポートエラー: " & Err.Description, vbCritical
End Sub

Private Function ArrowStr(cur As Long, prev As Long) As String
    If cur > prev Then
        ArrowStr = Chr(9650) & Format(cur - prev, "#,##0")
    ElseIf cur < prev Then
        ArrowStr = Chr(9660) & Format(prev - cur, "#,##0")
    Else
        ArrowStr = "-"
    End If
End Function"""

# F_Ranking: DateSerial修正 + エラーハンドリング
VBA_RANKING = """Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer

Private Function NextMonthFirst() As Date
    If mM = 12 Then
        NextMonthFirst = DateSerial(mY + 1, 1, 1)
    Else
        NextMonthFirst = DateSerial(mY, mM + 1, 1)
    End If
End Function

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date)
    LoadData
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    dT = Format(NextMonthFirst(), "yyyy/mm/dd")

    lstRef.RowSource = "SELECT R.member_name, Sum(R.referral)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect)" & _
        " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" & _
        " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" & _
        " GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRef.Requery
    lstRecv.Requery
    lstProsp.Requery
    Exit Sub
EH: MsgBox "エラー: " & Err.Description, vbCritical
End Sub"""


FORM_VBA = {
    "F_Main": VBA_MAIN,
    "F_Members": VBA_MEMBERS,
    "F_Daily": VBA_DAILY,
    "F_DailyEdit": VBA_DAILY_EDIT,
    "F_Targets": VBA_TARGETS,
    "F_Referrals": VBA_REFERRALS,
    "F_Report": VBA_REPORT,
    "F_Ranking": VBA_RANKING,
}


def main():
    print("VBA v2: Bug-fixed injection")
    print("=" * 50)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE_PATH)
    time.sleep(2)

    for form_name, code in FORM_VBA.items():
        try:
            app.DoCmd.OpenForm(form_name, 1)  # Design view
            time.sleep(0.5)
            frm = app.Forms(form_name)
            frm.HasModule = True
            time.sleep(0.3)

            comp = app.VBE.VBProjects(1).VBComponents("Form_" + form_name)
            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)

            lines = code.strip().split("\n")
            for i, line in enumerate(lines, 1):
                cm.InsertLines(i, line)

            app.DoCmd.Save(2, form_name)
            app.DoCmd.Close(2, form_name)
            time.sleep(0.3)
            print(f"  {form_name} OK ({len(lines)} lines)")
        except Exception as e:
            print(f"  {form_name} ERROR: {e}")
            try:
                app.DoCmd.Close(2, form_name)
            except:
                pass

    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)
    print("\nDone!")


if __name__ == "__main__":
    main()
