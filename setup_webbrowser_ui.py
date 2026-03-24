# -*- coding: utf-8 -*-
"""
setup_webbrowser_ui.py  --  SalesMgr_FE.accdb 対象

【実行方法】
1. SalesMgr_FE.accdb を手動でダブルクリックして Access を開く
2. マクロセキュリティを有効化し、ナビゲーションウィンドウが見える状態にする
3. このスクリプトを実行する（GetActiveObject で接続）

F_Main を WebBrowser ベースの HTML UI に切り替える。
"""
import os, time, win32com.client

FOLDER   = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
FE       = os.path.join(FOLDER, "SalesMgr_FE.accdb")
HTML_WIN = os.path.join(FOLDER, "ui", "index.html")
HTML_CONST = HTML_WIN.replace("\\", "/")

CM = 567
def c(n): return int(CM * n)

acForm  = 2
acDesign = 1
acCustomControl = 119
acDetail = 0

FULL_VBA = r'''Option Compare Database
Option Explicit

Private Const HTML_FILE As String = "''' + HTML_CONST + r'''"
Private mYear  As Integer
Private mMonth As Integer
Private mBusy  As Boolean

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo EH
    mYear = Year(Date): mMonth = Month(Date): mBusy = False
    Me.TimerInterval = 300
    DoCmd.Maximize
    Me.WebBrowser0.Object.Navigate "file:///" & HTML_FILE
    Exit Sub
EH: MsgBox Err.Description, vbCritical, "Form_Open"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim w As Long: w = Me.InsideWidth
    Dim h As Long: h = Me.InsideHeight
    If w > 0 Then Me.WebBrowser0.Width = w
    If h > 0 Then Me.WebBrowser0.Height = h
End Sub

Private Sub Form_Close(): Me.TimerInterval = 0: End Sub

Private Sub Form_Timer()
    On Error Resume Next
    If mBusy Then Exit Sub
    Dim doc As Object: Set doc = Me.WebBrowser0.Object.Document
    If doc Is Nothing Then Exit Sub
    Dim title As String: title = doc.Title
    If Left(title, 4) <> "CMD:" Then Set doc = Nothing: Exit Sub
    mBusy = True
    Dim cmd As String: cmd = Mid(title, 5)
    doc.Title = "SalesMgr": Set doc = Nothing
    ProcessCmd cmd: mBusy = False
End Sub

Private Sub ProcessCmd(ByVal cmd As String)
    On Error GoTo EH
    Dim p() As String: p = Split(cmd, "|")
    Select Case LCase(p(0))
        Case "reqdata"
            Select Case p(1)
                Case "init":      PushInit
                Case "dashboard": PushDashboard
                Case "daily"
                    If UBound(p) >= 3 Then mYear = CInt(p(2)): mMonth = CInt(p(3))
                    Dim mf As String: If UBound(p) >= 4 Then mf = p(4) Else mf = ""
                    PushDailyList mf
                Case "report"
                    If UBound(p) >= 3 Then mYear = CInt(p(2)): mMonth = CInt(p(3))
                    PushReport
                Case "targets"
                    If UBound(p) >= 3 Then mYear = CInt(p(2)): mMonth = CInt(p(3))
                    PushTargets
                Case "members":  PushMembers
                Case "ranking"
                    If UBound(p) >= 3 Then mYear = CInt(p(2)): mMonth = CInt(p(3))
                    PushRanking
                Case "referrals"
                    If UBound(p) >= 3 Then mYear = CInt(p(2)): mMonth = CInt(p(3))
                    PushReferrals
            End Select
        Case "savedaily":    SaveDailyRec p
        Case "deldaily":     DelDailyRec CLng(p(1))
        Case "savetarget":   SaveTarget p
        Case "deltarget":    DelTarget CLng(p(1))
        Case "addmember":    AddMember p(1)
        Case "togglemember": ToggleMember p(1), (p(2) = "1")
        Case "saveref":      SaveRef p
        Case "delref":       DelRef CLng(p(1))
    End Select
    ExecJS "vbaReady()"
    Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": ExecJS "vbaReady()"
End Sub

Private Sub ExecJS(ByVal code As String)
    On Error Resume Next
    Me.WebBrowser0.Object.Document.parentWindow.execScript code, "JavaScript"
End Sub

Private Function JQ(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, Chr(34), "\" & Chr(34))
    s = Replace(s, Chr(10), "\n"): s = Replace(s, Chr(13), "")
    JQ = Chr(34) & s & Chr(34)
End Function
Private Function JN(ByVal n As Long) As String:   JN = CStr(n):             End Function
Private Function JD(ByVal d As Double) As String: JD = Format(d, "0.00"):   End Function
Private Function JB(ByVal b As Boolean) As String
    If b Then JB = "true" Else JB = "false"
End Function
Private Function JSEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, "'", "\'"): JSEsc = Replace(s, Chr(10), " ")
End Function
Private Function DFrom(yr As Integer, mo As Integer) As Date: DFrom = DateSerial(yr, mo, 1): End Function
Private Function DTo(yr As Integer, mo As Integer) As Date
    If mo = 12 Then DTo = DateSerial(yr + 1, 1, 1) Else DTo = DateSerial(yr, mo + 1, 1)
End Function
Private Function DF(d As Date) As String: DF = Format(d, "yyyy/mm/dd"): End Function

Private Sub PushInit()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name", dbOpenSnapshot)
    Dim aMem As String: aMem = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then aMem = aMem & ","
        first = False: aMem = aMem & JQ(rs!member_name): rs.MoveNext
    Loop: aMem = aMem & "]": rs.Close
    Dim cntMem As Long, cntRec As Long, cntRef As Long
    Set rs = db.OpenRecordset("SELECT Count(*) AS n FROM T_MEMBERS WHERE active=True", dbOpenSnapshot): cntMem = Nz(rs!n, 0): rs.Close
    Set rs = db.OpenRecordset("SELECT Count(*) AS n FROM T_RECORDS", dbOpenSnapshot): cntRec = Nz(rs!n, 0): rs.Close
    Set rs = db.OpenRecordset("SELECT Count(*) AS n FROM T_REFERRALS", dbOpenSnapshot): cntRef = Nz(rs!n, 0): rs.Close
    Dim json As String
    json = "{""year"":" & mYear & ",""month"":" & mMonth & "," & _
           """activeMembers"":" & aMem & ",""allMembers"":" & aMem & "," & _
           """memberCount"":" & cntMem & ",""recordCount"":" & cntRec & ",""referralCount"":" & cntRef & "}"
    ExecJS "initApp(" & json & ")"
    Set db = Nothing: PushDashboard: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushDashboard()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim dF As Date: dF = DFrom(mYear, mMonth): Dim dT As Date: dT = DTo(mYear, mMonth)
    Dim sql As String
    sql = "SELECT Nz(Sum(calls),0) AS tc, Nz(Sum(valid_count),0) AS tv, Nz(Sum(prospect),0) AS tp, Nz(Sum(received),0) AS tr_, Nz(Sum(referral),0) AS tref, Nz(Sum(work_hours),0) AS th FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "#"
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim tc As Long: tc=CLng(Nz(rs!tc,0)): Dim tv As Long: tv=CLng(Nz(rs!tv,0))
    Dim tp As Long: tp=CLng(Nz(rs!tp,0)): Dim tr_ As Long: tr_=CLng(Nz(rs!tr_,0))
    Dim tref As Long: tref=CLng(Nz(rs!tref,0)): Dim th As Double: th=CDbl(Nz(rs!th,0)): rs.Close
    sql = "SELECT Nz(Sum(target_calls),0) AS gc, Nz(Sum(target_valid),0) AS gv, Nz(Sum(target_prospect),0) AS gp, Nz(Sum(target_received),0) AS gr, Nz(Sum(target_referral),0) AS gref FROM T_MEMBER_TARGETS WHERE target_year=" & mYear & " AND target_month=" & mMonth
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim gc As Long: gc=CLng(Nz(rs!gc,0)): Dim gv As Long: gv=CLng(Nz(rs!gv,0))
    Dim gp As Long: gp=CLng(Nz(rs!gp,0)): Dim gr As Long: gr=CLng(Nz(rs!gr,0))
    Dim gref As Long: gref=CLng(Nz(rs!gref,0)): rs.Close
    sql = "SELECT member_name, Nz(Sum(calls),0) AS mc, Nz(Sum(valid_count),0) AS mv, Nz(Sum(prospect),0) AS mp, Nz(Sum(received),0) AS mr, Nz(Sum(referral),0) AS mref FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "# GROUP BY member_name ORDER BY Nz(Sum(calls),0) DESC"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim memJs As String: memJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then memJs = memJs & ",": first = False
        memJs = memJs & "{""name"":" & JQ(rs!member_name) & ",""calls"":" & JN(CLng(Nz(rs!mc,0))) & ",""valid"":" & JN(CLng(Nz(rs!mv,0))) & ",""prospect"":" & JN(CLng(Nz(rs!mp,0))) & ",""received"":" & JN(CLng(Nz(rs!mr,0))) & ",""referral"":" & JN(CLng(Nz(rs!mref,0))) & "}"
        rs.MoveNext
    Loop: memJs = memJs & "]": rs.Close
    Dim trendJs As String: trendJs = "[": first = True: Dim i As Integer
    For i = 5 To 0 Step -1
        Dim tY As Integer: tY = mYear: Dim tM As Integer: tM = mMonth - i
        Do While tM <= 0: tM = tM + 12: tY = tY - 1: Loop
        Dim tdF As Date: tdF = DFrom(tY, tM): Dim tdT As Date: tdT = DTo(tY, tM)
        sql = "SELECT Nz(Sum(calls),0) AS tc2 FROM T_RECORDS WHERE rec_date>=#" & DF(tdF) & "# AND rec_date<#" & DF(tdT) & "#"
        Dim trs As DAO.Recordset: Set trs = db.OpenRecordset(sql, dbOpenSnapshot)
        If Not first Then trendJs = trendJs & ",": first = False
        trendJs = trendJs & "{""label"":" & JQ(CStr(tM) & "月") & ",""calls"":" & JN(CLng(Nz(trs!tc2,0))) & "}"
        trs.Close: Set trs = Nothing
    Next i: trendJs = trendJs & "]"
    Dim json As String
    json = "{""year"":" & mYear & ",""month"":" & mMonth & ",""calls"":" & tc & ",""callsTarget"":" & gc & ",""valid"":" & tv & ",""validTarget"":" & gv & ",""prospect"":" & tp & ",""prospectTarget"":" & gp & ",""received"":" & tr_ & ",""receivedTarget"":" & gr & ",""referral"":" & tref & ",""referralTarget"":" & gref & ",""workHours"":" & JD(th) & ",""members"":" & memJs & ",""trend"":" & trendJs & "}"
    ExecJS "loadDashboard(" & json & ")"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushDailyList(Optional ByVal mf As String = "")
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim dF As Date: dF = DFrom(mYear, mMonth): Dim dT As Date: dT = DTo(mYear, mMonth)
    Dim sql As String
    sql = "SELECT ID, rec_date, member_name, calls, calls_10, calls_11, calls_12, calls_13, calls_14, calls_15, calls_16, calls_17, calls_18, valid_count, prospect, doc, follow_up, received, work_hours, referral, note, work_day FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "#"
    If mf <> "" Then sql = sql & " AND member_name=" & JQ(mf)
    sql = sql & " ORDER BY rec_date DESC, member_name"
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim rJs As String: rJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then rJs = rJs & ",": first = False
        Dim dt As Date: dt = rs!rec_date
        rJs = rJs & "{""id"":" & JN(rs!ID) & ",""date"":" & JQ(Format(dt,"yyyy/mm/dd")) & ",""dateISO"":" & JQ(Format(dt,"yyyy-mm-dd")) & ",""member"":" & JQ(rs!member_name) & ",""calls"":" & JN(CLng(Nz(rs!calls,0))) & ",""c10"":" & JN(CLng(Nz(rs!calls_10,0))) & ",""c11"":" & JN(CLng(Nz(rs!calls_11,0))) & ",""c12"":" & JN(CLng(Nz(rs!calls_12,0))) & ",""c13"":" & JN(CLng(Nz(rs!calls_13,0))) & ",""c14"":" & JN(CLng(Nz(rs!calls_14,0))) & ",""c15"":" & JN(CLng(Nz(rs!calls_15,0))) & ",""c16"":" & JN(CLng(Nz(rs!calls_16,0))) & ",""c17"":" & JN(CLng(Nz(rs!calls_17,0))) & ",""c18"":" & JN(CLng(Nz(rs!calls_18,0))) & ",""valid"":" & JN(CLng(Nz(rs!valid_count,0))) & ",""prospect"":" & JN(CLng(Nz(rs!prospect,0))) & ",""doc"":" & JN(CLng(Nz(rs!doc,0))) & ",""follow"":" & JN(CLng(Nz(rs!follow_up,0))) & ",""received"":" & JN(CLng(Nz(rs!received,0))) & ",""referral"":" & JN(CLng(Nz(rs!referral,0))) & ",""workHours"":" & JQ(Format(CDbl(Nz(rs!work_hours,0)),"0.0")) & ",""workHoursRaw"":" & JD(CDbl(Nz(rs!work_hours,0))) & ",""workDay"":" & JB(CBool(Nz(rs!work_day,True))) & ",""note"":" & JQ(Nz(rs!note,"")) & "}"
        rs.MoveNext
    Loop: rJs = rJs & "]": rs.Close
    ExecJS "loadDailyList({""year"":" & mYear & ",""month"":" & mMonth & ",""member"":" & JQ(mf) & ",""records"":" & rJs & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushReport()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim dF As Date: dF = DFrom(mYear, mMonth): Dim dT As Date: dT = DTo(mYear, mMonth)
    Dim sql As String
    sql = "SELECT Nz(Sum(calls),0) AS tc, Nz(Sum(valid_count),0) AS tv, Nz(Sum(prospect),0) AS tp, Nz(Sum(received),0) AS tr_, Nz(Sum(work_hours),0) AS th FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "#"
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim tc As Long: tc=CLng(Nz(rs!tc,0)): Dim tv As Long: tv=CLng(Nz(rs!tv,0))
    Dim tp As Long: tp=CLng(Nz(rs!tp,0)): Dim tr_ As Long: tr_=CLng(Nz(rs!tr_,0))
    Dim th As Double: th=CDbl(Nz(rs!th,0)): rs.Close
    Dim tgtCalls As Long: tgtCalls=0
    sql = "SELECT Nz(Sum(target_calls),0) AS gc FROM T_MEMBER_TARGETS WHERE target_year=" & mYear & " AND target_month=" & mMonth
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot): tgtCalls=CLng(Nz(rs!gc,0)): rs.Close
    sql = "SELECT member_name, Nz(Sum(calls),0) AS mc, Nz(Sum(valid_count),0) AS mv, Nz(Sum(prospect),0) AS mp, Nz(Sum(doc),0) AS md_, Nz(Sum(follow_up),0) AS mf2, Nz(Sum(received),0) AS mr_, Nz(Sum(referral),0) AS mref, Nz(Sum(work_hours),0) AS mh FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "# GROUP BY member_name ORDER BY Nz(Sum(calls),0) DESC"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim mJs As String: mJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then mJs = mJs & ",": first = False
        mJs = mJs & "{""name"":" & JQ(rs!member_name) & ",""calls"":" & JN(CLng(Nz(rs!mc,0))) & ",""valid"":" & JN(CLng(Nz(rs!mv,0))) & ",""prospect"":" & JN(CLng(Nz(rs!mp,0))) & ",""doc"":" & JN(CLng(Nz(rs!md_,0))) & ",""follow"":" & JN(CLng(Nz(rs!mf2,0))) & ",""received"":" & JN(CLng(Nz(rs!mr_,0))) & ",""referral"":" & JN(CLng(Nz(rs!mref,0))) & ",""workHours"":" & JQ(Format(CDbl(Nz(rs!mh,0)),"0.0")) & ",""workHoursRaw"":" & JD(CDbl(Nz(rs!mh,0))) & "}"
        rs.MoveNext
    Loop: mJs = mJs & "]": rs.Close
    ExecJS "loadReport({""year"":" & mYear & ",""month"":" & mMonth & ",""calls"":" & tc & ",""valid"":" & tv & ",""prospect"":" & tp & ",""received"":" & tr_ & ",""workHours"":" & JQ(Format(th,"0.0")) & ",""callsTarget"":" & tgtCalls & ",""members"":" & mJs & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushTargets()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim sql As String
    sql = "SELECT ID, member_name, plan_days, work_hours_per_day, target_calls, target_valid, target_prospect, target_received, target_referral FROM T_MEMBER_TARGETS WHERE target_year=" & mYear & " AND target_month=" & mMonth & " ORDER BY member_name"
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim tJs As String: tJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then tJs = tJs & ",": first = False
        tJs = tJs & "{""id"":" & JN(rs!ID) & ",""member"":" & JQ(rs!member_name) & ",""planDays"":" & JN(CLng(Nz(rs!plan_days,0))) & ",""hrsPerDay"":" & JD(CDbl(Nz(rs!work_hours_per_day,8))) & ",""calls"":" & JN(CLng(Nz(rs!target_calls,0))) & ",""valid"":" & JN(CLng(Nz(rs!target_valid,0))) & ",""prospect"":" & JN(CLng(Nz(rs!target_prospect,0))) & ",""received"":" & JN(CLng(Nz(rs!target_received,0))) & ",""referral"":" & JN(CLng(Nz(rs!target_referral,0))) & "}"
        rs.MoveNext
    Loop: tJs = tJs & "]": rs.Close
    ExecJS "loadTargets({""year"":" & mYear & ",""month"":" & mMonth & ",""targets"":" & tJs & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushMembers()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ID, member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name", dbOpenSnapshot)
    Dim aJs As String: aJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then aJs = aJs & ",": first = False: aJs = aJs & "{""id"":" & JN(rs!ID) & ",""name"":" & JQ(rs!member_name) & "}": rs.MoveNext
    Loop: aJs = aJs & "]": rs.Close
    Set rs = db.OpenRecordset("SELECT ID, member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name", dbOpenSnapshot)
    Dim iJs As String: iJs = "[": first = True
    Do While Not rs.EOF
        If Not first Then iJs = iJs & ",": first = False: iJs = iJs & "{""id"":" & JN(rs!ID) & ",""name"":" & JQ(rs!member_name) & "}": rs.MoveNext
    Loop: iJs = iJs & "]": rs.Close
    ExecJS "loadMembers({""active"":" & aJs & ",""inactive"":" & iJs & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub PushRanking()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim dF As Date: dF = DFrom(mYear, mMonth): Dim dT As Date: dT = DTo(mYear, mMonth)
    Dim wc As String: wc = "FROM T_RECORDS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "# GROUP BY member_name"
    ExecJS "loadRanking({""year"":" & mYear & ",""month"":" & mMonth & ",""rankCalls"":" & BRJ(db,"Nz(Sum(calls),0)",wc) & ",""rankReceived"":" & BRJ(db,"Nz(Sum(received),0)",wc) & ",""rankProspect"":" & BRJ(db,"Nz(Sum(prospect),0)",wc) & ",""rankReferral"":" & BRJ(db,"Nz(Sum(referral),0)",wc) & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Function BRJ(db As DAO.Database, ByVal ae As String, ByVal wc As String) As String
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("SELECT member_name, " & ae & " AS rv " & wc & " ORDER BY rv DESC", dbOpenSnapshot)
    Dim js As String: js = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then js = js & ",": first = False: js = js & "{""name"":" & JQ(rs!member_name) & ",""val"":" & JN(CLng(Nz(rs!rv,0))) & "}": rs.MoveNext
    Loop: js = js & "]": rs.Close: BRJ = js
End Function

Private Sub PushReferrals()
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim dF As Date: dF = DFrom(mYear, mMonth): Dim dT As Date: dT = DTo(mYear, mMonth)
    Dim sql As String
    sql = "SELECT ID, rec_date, member_name, ref_count FROM T_REFERRALS WHERE rec_date>=#" & DF(dF) & "# AND rec_date<#" & DF(dT) & "# ORDER BY rec_date DESC, member_name"
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim rJs As String: rJs = "[": Dim first As Boolean: first = True
    Do While Not rs.EOF
        If Not first Then rJs = rJs & ",": first = False
        rJs = rJs & "{""id"":" & JN(rs!ID) & ",""date"":" & JQ(Format(rs!rec_date,"yyyy/mm/dd")) & ",""member"":" & JQ(rs!member_name) & ",""count"":" & JN(CLng(Nz(rs!ref_count,0))) & "}"
        rs.MoveNext
    Loop: rJs = rJs & "]": rs.Close
    ExecJS "loadReferrals({""year"":" & mYear & ",""month"":" & mMonth & ",""records"":" & rJs & "})"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub SaveDailyRec(ByVal p() As String)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim id As String: id = Trim(p(1))
    Dim recDate As Date: recDate = CDate(p(2))
    Dim wh As Double: wh = CDbl(p(4)): Dim wd As Boolean: wd = (p(5) = "1")
    Dim c10 As Long: c10=CLng(p(6)): Dim c11 As Long: c11=CLng(p(7)): Dim c12 As Long: c12=CLng(p(8))
    Dim c13 As Long: c13=CLng(p(9)): Dim c14 As Long: c14=CLng(p(10)): Dim c15 As Long: c15=CLng(p(11))
    Dim c16 As Long: c16=CLng(p(12)): Dim c17 As Long: c17=CLng(p(13)): Dim c18 As Long: c18=CLng(p(14))
    Dim tot As Long: tot=c10+c11+c12+c13+c14+c15+c16+c17+c18
    Dim v_ As Long: v_=CLng(p(15)): Dim pr As Long: pr=CLng(p(16)): Dim dc As Long: dc=CLng(p(17))
    Dim fw As Long: fw=CLng(p(18)): Dim rc As Long: rc=CLng(p(19)): Dim rf As Long: rf=CLng(p(20))
    Dim ms As String: ms=Replace(p(3),"'","''"): Dim ns As String: ns=Replace(p(21),"'","''")
    Dim dfs As String: dfs=Format(recDate,"yyyy/mm/dd"): Dim sql As String
    If id = "" Then
        sql = "INSERT INTO T_RECORDS (rec_date,member_name,calls,calls_10,calls_11,calls_12,calls_13,calls_14,calls_15,calls_16,calls_17,calls_18,valid_count,prospect,doc,follow_up,received,work_hours,referral,note,work_day) VALUES (#" & dfs & "#,'" & ms & "'," & tot & "," & c10 & "," & c11 & "," & c12 & "," & c13 & "," & c14 & "," & c15 & "," & c16 & "," & c17 & "," & c18 & "," & v_ & "," & pr & "," & dc & "," & fw & "," & rc & "," & wh & "," & rf & ",'" & ns & "'," & IIf(wd,-1,0) & ")"
    Else
        sql = "UPDATE T_RECORDS SET rec_date=#" & dfs & "#,member_name='" & ms & "',calls=" & tot & ",calls_10=" & c10 & ",calls_11=" & c11 & ",calls_12=" & c12 & ",calls_13=" & c13 & ",calls_14=" & c14 & ",calls_15=" & c15 & ",calls_16=" & c16 & ",calls_17=" & c17 & ",calls_18=" & c18 & ",valid_count=" & v_ & ",prospect=" & pr & ",doc=" & dc & ",follow_up=" & fw & ",received=" & rc & ",work_hours=" & wh & ",referral=" & rf & ",note='" & ns & "',work_day=" & IIf(wd,-1,0) & " WHERE ID=" & id
    End If
    db.Execute sql, dbFailOnError: ExecJS "showMsg('保存しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub DelDailyRec(ByVal id As Long)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    db.Execute "DELETE FROM T_RECORDS WHERE ID=" & id, dbFailOnError: ExecJS "showMsg('削除しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub SaveTarget(ByVal p() As String)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim id As String: id = Trim(p(1)): Dim ms As String: ms = Replace(p(2),"'","''")
    Dim yr As Long: yr=CLng(p(3)): Dim mo As Long: mo=CLng(p(4)): Dim pd As Long: pd=CLng(p(5))
    Dim hd As Double: hd=CDbl(p(6)): Dim tc As Long: tc=CLng(p(7)): Dim tv_ As Long: tv_=CLng(p(8))
    Dim tp As Long: tp=CLng(p(9)): Dim tr_ As Long: tr_=CLng(p(10)): Dim tref As Long: tref=CLng(p(11))
    Dim sql As String
    If id = "" Then
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("SELECT ID FROM T_MEMBER_TARGETS WHERE member_name='" & ms & "' AND target_year=" & yr & " AND target_month=" & mo, dbOpenSnapshot)
        If Not rs.EOF Then id = CStr(rs!ID): rs.Close
    End If
    If id = "" Then
        sql = "INSERT INTO T_MEMBER_TARGETS (member_name,target_year,target_month,plan_days,work_hours_per_day,target_calls,target_valid,target_prospect,target_received,target_referral) VALUES ('" & ms & "'," & yr & "," & mo & "," & pd & "," & hd & "," & tc & "," & tv_ & "," & tp & "," & tr_ & "," & tref & ")"
    Else
        sql = "UPDATE T_MEMBER_TARGETS SET plan_days=" & pd & ",work_hours_per_day=" & hd & ",target_calls=" & tc & ",target_valid=" & tv_ & ",target_prospect=" & tp & ",target_received=" & tr_ & ",target_referral=" & tref & " WHERE ID=" & id
    End If
    db.Execute sql, dbFailOnError: ExecJS "showMsg('目標を保存しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub DelTarget(ByVal id As Long)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    db.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & id, dbFailOnError: ExecJS "showMsg('削除しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub AddMember(ByVal name As String)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    db.Execute "INSERT INTO T_MEMBERS (member_name,active) VALUES ('" & Replace(name,"'","''") & "',True)", dbFailOnError
    ExecJS "showMsg('メンバーを追加しました')": Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub ToggleMember(ByVal name As String, ByVal makeActive As Boolean)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim v As Long: If makeActive Then v = -1 Else v = 0
    db.Execute "UPDATE T_MEMBERS SET active=" & v & " WHERE member_name='" & Replace(name,"'","''") & "'", dbFailOnError
    ExecJS "showMsg('" & IIf(makeActive,"有効にしました","無効にしました") & "')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub SaveRef(ByVal p() As String)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    Dim id As String: id = Trim(p(1)): Dim refDate As Date: refDate = CDate(p(2))
    Dim ms As String: ms = Replace(p(3),"'","''"): Dim cnt As Long: cnt = CLng(p(4))
    Dim dfs As String: dfs = Format(refDate,"yyyy/mm/dd"): Dim sql As String
    If id = "" Then
        sql = "INSERT INTO T_REFERRALS (rec_date,member_name,ref_count) VALUES (#" & dfs & "#,'" & ms & "'," & cnt & ")"
    Else
        sql = "UPDATE T_REFERRALS SET rec_date=#" & dfs & "#,member_name='" & ms & "',ref_count=" & cnt & " WHERE ID=" & id
    End If
    db.Execute sql, dbFailOnError: ExecJS "showMsg('送客を登録しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub

Private Sub DelRef(ByVal id As Long)
    On Error GoTo EH
    Dim db As DAO.Database: Set db = CurrentDb()
    db.Execute "DELETE FROM T_REFERRALS WHERE ID=" & id, dbFailOnError: ExecJS "showMsg('削除しました')"
    Set db = Nothing: Exit Sub
EH: ExecJS "showMsg('" & JSEsc(Err.Description) & "',true)": Set db = Nothing
End Sub
'''


def inject_vba(app, form_name, code):
    try:
        comp = app.VBE.VBProjects(1).VBComponents('Form_' + form_name)
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        lines = code.replace('\r\n', '\n').strip().split('\n')
        for i, line in enumerate(lines, 1):
            cm.InsertLines(i, line)
        return len(lines)
    except Exception as e:
        print(f'  VBA注入失敗: {e}')
        return 0


def main():
    print('=' * 65)
    print('SalesMgr WebBrowser UI セットアップ（GetActiveObject方式）')
    print('=' * 65)

    # ── 既に開いている Access に接続 ──────────────────────────────────
    print('\n[接続] 起動中の Access.Application に接続...')
    app = None
    try:
        app = win32com.client.GetActiveObject('Access.Application')
        print(f'  接続成功: {app.CurrentDb().Name}')
    except Exception as e:
        print(f'  GetActiveObject 失敗: {e}')
        print('  → Access が開いていない場合は新規起動を試みます...')
        import subprocess
        msaccess = r'C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE'
        if not os.path.exists(msaccess):
            msaccess = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE'
        if not os.path.exists(msaccess):
            print('  MSACCESS.EXE が見つかりません。手動で Access を開いて再実行してください。')
            return
        subprocess.Popen([msaccess, FE])
        print('  Access を起動しました。30秒待機 (ダイアログがあれば手動で閉じてください)...')
        for i in range(30, 0, -5):
            time.sleep(5)
            print(f'  あと {i-5} 秒...')
            try:
                app = win32com.client.GetActiveObject('Access.Application')
                print('  接続成功！')
                break
            except:
                pass

    if app is None:
        print('ERROR: Access に接続できませんでした')
        return

    # ── F_Main をデザインビューで開く ──────────────────────────────────
    print('\n[Phase 1] F_Main をデザインビューで開く...')
    # まず既に開いていれば閉じる
    try:
        app.DoCmd.Close(acForm, 'F_Main')
        time.sleep(1)
        print('  既存 F_Main を閉じました')
    except:
        pass
    try:
        app.DoCmd.SetWarnings(False)
        app.DoCmd.OpenForm('F_Main', acDesign)
        time.sleep(2)
        print('  F_Main 開きました')
    except Exception as e:
        # 排他モード警告は例外として来るが、フォームは開いている場合がある
        print(f'  OpenForm 警告/エラー: {e}')
        print('  フォームが開いているか確認...')
        time.sleep(2)
        try:
            frm_check = app.Forms('F_Main')
            print(f'  F_Main は開いています (コントロール数={frm_check.Controls.Count})')
        except Exception as e2:
            print(f'  F_Main が開いていません: {e2}')
            print('  → RunCommand を使ってデザインビューに切り替えます...')
            try:
                # acCmdFormDesignView = 183
                app.DoCmd.OpenForm('F_Main', acDesign)
                time.sleep(2)
            except Exception as e3:
                print(f'  再試行も失敗: {e3}')
                return

    # ── 既存コントロール削除 ──────────────────────────────────────────
    try:
        frm = app.Forms('F_Main')
        names = [c.Name for c in frm.Controls]
        print(f'  既存コントロール {len(names)} 個: {names}')
        for nm in names:
            try:
                app.DeleteControl('F_Main', nm)
                print(f'  削除: {nm}')
            except Exception as e:
                print(f'  削除スキップ ({nm}): {e}')

        # フォームプロパティ
        frm.Caption = 'SalesMgr 営業管理'
        frm.RecordSelectors = False
        frm.NavigationButtons = False
        frm.DividingLines = False
        frm.ScrollBars = 0
        frm.Width = 18000
        frm.Section(0).Height = 12000
        print('  フォームプロパティ設定完了')
    except Exception as e:
        print(f'  コントロール整理エラー: {e}')

    # ── WebBrowser ActiveX コントロール作成 ───────────────────────────
    print('\n[Phase 2] WebBrowser コントロール追加...')
    wb_ok = False
    try:
        wb = app.CreateControl('F_Main', acCustomControl, acDetail, '', '', 0, 0, 18000, 12000)
        print(f'  CreateControl 成功: ControlType={wb.ControlType}')
        try:
            wb.OLEClass = 'Shell.Explorer.2'
            print(f'  OLEClass 設定: {wb.OLEClass}')
        except Exception as e:
            print(f'  OLEClass 設定失敗: {e} (後でName設定を試みます)')
        try:
            wb.Name = 'WebBrowser0'
            print(f'  Name 設定: {wb.Name}')
        except Exception as e:
            print(f'  Name 設定失敗: {e}')
        wb_ok = True
    except Exception as e:
        print(f'  CreateControl 失敗: {e}')
        print('  → WebBrowser なしで VBA のみ注入します（手動での追加が必要）')

    # ── VBA 注入 ──────────────────────────────────────────────────────
    print('\n[Phase 3] VBA 注入...')
    try:
        frm = app.Forms('F_Main')
        frm.HasModule = True
        frm.TimerInterval = 300
        frm.OnTimer = '[Event Procedure]'
        time.sleep(0.3)
        n = inject_vba(app, 'F_Main', FULL_VBA)
        print(f'  VBA 注入: {n} 行')
    except Exception as e:
        print(f'  VBA 注入エラー: {e}')

    # ── 保存 ──────────────────────────────────────────────────────────
    print('\n[Phase 4] 保存...')
    try:
        app.DoCmd.SetWarnings(False)
        app.DoCmd.Save(acForm, 'F_Main')
        app.DoCmd.Close(acForm, 'F_Main')
        app.DoCmd.SetWarnings(True)
        print('  F_Main 保存完了')
    except Exception as e:
        print(f'  保存エラー: {e}')
        try: app.DoCmd.SetWarnings(True)
        except: pass

    print('\n' + '=' * 65)
    if wb_ok:
        print('完了！ WebBrowser コントロール + VBA 注入完了')
    else:
        print('部分完了: VBA は注入済み。WebBrowser コントロールは手動追加が必要')
        print()
        print('【手動追加手順】')
        print('  1. F_Main をデザインビューで開く')
        print('  2. デザイン → コントロール → ActiveX コントロールの挿入')
        print('  3. "Microsoft Web Browser" を選択してOK')
        print('  4. フォーム全体に広げてプロパティでName="WebBrowser0"に設定')
        print('  5. 保存')
    print('=' * 65)


if __name__ == '__main__':
    main()
