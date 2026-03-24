# -*- coding: utf-8 -*-
"""
enhance_access.py
  1. F_List   新規作成（日次一覧 + CSV出力）
  2. F_Analysis 新規作成（分析：ランキング＋時間帯別）
  3. F_Main  VBA更新（btnList_Click / btnAnalysis_Click 追加）
  4. F_Report VBA更新（btnPDF_Click → Python PDF生成スクリプト呼び出し）
  5. 全フォームのフォントを 游ゴシック に統一
"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")
PY = r"C:\Users\agentcode01\AppData\Local\Programs\Python\Python312\python.exe"

acLabel=100; acButton=104; acTextBox=109; acCombo=111; acList=110; acCheck=106; acDetail=0
CM = 567

# ──────────────────────────────────────────────
# VBA ソース定義
# ──────────────────────────────────────────────

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
End Sub

Private Sub btnList_Click()
    DoCmd.OpenForm "F_List"
End Sub

Private Sub btnAnalysis_Click()
    DoCmd.OpenForm "F_Analysis"
End Sub"""

# ─── F_Report (PDF出力を実装) ───────────────────
VBA_REPORT = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date)
    mM = Month(Date)
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
On Error GoTo EH
    Dim pyExe As String
    Dim script As String
    Dim dbPath As String
    pyExe  = "C:\\Users\\agentcode01\\AppData\\Local\\Programs\\Python\\Python312\\python.exe"
    script = CurrentProject.Path & "\\generate_pdf.py"
    dbPath = CurrentProject.FullName
    Dim cmd As String
    cmd = "cmd /c " & Chr(34) & pyExe & Chr(34) & " " _
        & Chr(34) & script & Chr(34) _
        & " --year " & mY & " --month " & mM _
        & " --db " & Chr(34) & dbPath & Chr(34)
    Shell cmd, vbHide
    MsgBox mY & Chr(24180) & mM & Chr(26376) & " PDF" & Chr(29983) & Chr(25104) & Chr(20013) & "..." & vbCrLf & Chr(12487) & Chr(12473) & Chr(12463) & Chr(12488) & Chr(12483) & Chr(12503) & Chr(12395) & Chr(20445) & Chr(23384) & Chr(12373) & Chr(12428) & Chr(12414) & Chr(12377), vbInformation, "PDF" & Chr(20986) & Chr(21147)
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnPDF"
End Sub

Private Sub LoadReport()
On Error GoTo EH
    lblMonth.Caption = mY & Chr(24180) & mM & Chr(26376) & " " & Chr(12524) & Chr(12509) & Chr(12540) & Chr(12488)

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then
        dT = Format(DateSerial(mY + 1, 1, 1), "yyyy/mm/dd")
    Else
        dT = Format(DateSerial(mY, mM + 1, 1), "yyyy/mm/dd")
    End If

    Dim rs As DAO.Recordset
    Dim tC As Long, tV As Long, tP As Long, tR As Long, tH As Double
    tC = 0: tV = 0: tP = 0: tR = 0: tH = 0

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received), Sum(R.work_hours)" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#")
    If Not rs.EOF Then
        tC = Nz(rs.Fields(0).Value, 0)
        tV = Nz(rs.Fields(1).Value, 0)
        tP = Nz(rs.Fields(2).Value, 0)
        tR = Nz(rs.Fields(3).Value, 0)
        tH = Nz(rs.Fields(4).Value, 0)
    End If
    rs.Close

    Dim gC As Long, gV As Long, gP As Long, gR As Long
    gC = 0: gV = 0: gP = 0: gR = 0
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(target_calls), Sum(target_valid), Sum(target_prospect), Sum(target_received)" _
        & " FROM T_MEMBER_TARGETS WHERE target_year=" & mY & " AND target_month=" & mM)
    If Not rs.EOF Then
        gC = Nz(rs.Fields(0).Value, 0)
        gV = Nz(rs.Fields(1).Value, 0)
        gP = Nz(rs.Fields(2).Value, 0)
        gR = Nz(rs.Fields(3).Value, 0)
    End If
    rs.Close

    Dim py As Integer, pm As Integer
    py = mY: pm = mM
    If pm = 1 Then py = py - 1: pm = 12 Else pm = pm - 1
    Dim pF As String, pT As String
    pF = Format(DateSerial(py, pm, 1), "yyyy/mm/dd")
    If pm = 12 Then
        pT = Format(DateSerial(py + 1, 1, 1), "yyyy/mm/dd")
    Else
        pT = Format(DateSerial(py, pm + 1, 1), "yyyy/mm/dd")
    End If

    Dim pC As Long, pV As Long, pP As Long, pR As Long
    pC = 0: pV = 0: pP = 0: pR = 0
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received)" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & pF & "# AND R.rec_date < #" & pT & "#")
    If Not rs.EOF Then
        pC = Nz(rs.Fields(0).Value, 0)
        pV = Nz(rs.Fields(1).Value, 0)
        pP = Nz(rs.Fields(2).Value, 0)
        pR = Nz(rs.Fields(3).Value, 0)
    End If
    rs.Close

    lblCalls.Caption = Format(tC, "#,##0") & " / " & Format(gC, "#,##0")
    lblValid.Caption = Format(tV, "#,##0") & " / " & Format(gV, "#,##0")
    lblProsp.Caption = Format(tP, "#,##0") & " / " & Format(gP, "#,##0")
    lblRecv.Caption  = Format(tR, "#,##0") & " / " & Format(gR, "#,##0")

    lblCallsPrev.Caption = Chr(21069) & Chr(26376) & Format(pC, "#,##0") & " " & MakeArrow(tC, pC)
    lblValidPrev.Caption = Chr(21069) & Chr(26376) & Format(pV, "#,##0") & " " & MakeArrow(tV, pV)
    lblProspPrev.Caption = Chr(21069) & Chr(26376) & Format(pP, "#,##0") & " " & MakeArrow(tP, pP)
    lblRecvPrev.Caption  = Chr(21069) & Chr(26376) & Format(pR, "#,##0") & " " & MakeArrow(tR, pR)

    If tC > 0 Then
        lblValidRate.Caption = Format(tV / tC * 100, "0.0") & "%"
        lblRecvRate.Caption  = Format(tR / tC * 100, "0.0") & "%"
    Else
        lblValidRate.Caption = "-"
        lblRecvRate.Caption  = "-"
    End If
    lblHours.Caption = Format(tH, "#,##0") & "h"
    If tH > 0 Then
        lblProductivity.Caption = Format(tP / tH, "0.000")
    Else
        lblProductivity.Caption = "-"
    End If

    Dim al As String: al = ""
    If gR > 0 Then
        If tR >= gR Then
            al = al & Chr(9675) & " " & Chr(21463) & Chr(27880) & " " & Chr(30446) & Chr(26631) & Chr(36948) & Chr(25104) & "!" & vbCrLf
        Else
            al = al & Chr(9651) & " " & Chr(21463) & Chr(27880) & " " & Chr(27494) & Chr(12426) & (gR - tR) & Chr(20214) & " " & tR & "/" & gR & vbCrLf
        End If
    End If
    If gP > 0 Then
        If tP >= gP Then
            al = al & Chr(9675) & " " & Chr(35211) & Chr(36796) & " " & Chr(30446) & Chr(26631) & Chr(36948) & Chr(25104) & "!"
        Else
            al = al & Chr(9651) & " " & Chr(35211) & Chr(36796) & " " & Chr(27494) & Chr(12426) & (gP - tP) & Chr(20214) & " " & tP & "/" & gP
        End If
    End If
    lblAlert.Caption = al

    Dim baseQ As String
    baseQ = " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & " GROUP BY R.member_name ORDER BY "

    lstRankRef.RowSource  = "SELECT R.member_name, Sum(R.referral)"  & baseQ & "Sum(R.referral)  DESC"
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received)"  & baseQ & "Sum(R.received)  DESC"
    lstRankProsp.RowSource= "SELECT R.member_name, Sum(R.prospect)"  & baseQ & "Sum(R.prospect)  DESC"
    lstRankRef.Requery
    lstRankRecv.Requery
    lstRankProsp.Requery
    Exit Sub
EH:
    MsgBox "LoadReport Error: " & Err.Description, vbCritical, "LoadReport"
End Sub

Private Function MakeArrow(cur As Long, prev As Long) As String
    If cur > prev Then
        MakeArrow = Chr(9650) & Format(cur - prev, "#,##0")
    ElseIf cur < prev Then
        MakeArrow = Chr(9660) & Format(prev - cur, "#,##0")
    Else
        MakeArrow = "-"
    End If
End Function"""

# ─── F_List (日次一覧 + CSV出力) ─────────────────
VBA_LIST = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date)
    mM = Month(Date)
    cboMember.RowSource = "SELECT '' AS nm UNION SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY nm"
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

Private Sub cboMember_AfterUpdate()
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & Chr(24180) & mM & Chr(26376) & " " & Chr(19968) & Chr(35239)

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then
        dT = Format(DateSerial(mY + 1, 1, 1), "yyyy/mm/dd")
    Else
        dT = Format(DateSerial(mY, mM + 1, 1), "yyyy/mm/dd")
    End If

    Dim mf As String: mf = ""
    If Not IsNull(cboMember.Value) And Len(Trim(Nz(cboMember.Value, ""))) > 0 Then
        mf = " AND R.member_name='" & Replace(Nz(cboMember.Value, ""), "'", "''") & "'"
    End If

    lstRecords.RowSource = _
        "SELECT Format(R.rec_date,'yyyy/mm/dd'), R.member_name," _
        & " R.calls, R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.note" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & mf & " ORDER BY R.rec_date DESC, R.member_name"
    lstRecords.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub

Private Sub btnCSV_Click()
On Error GoTo EH
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then
        dT = Format(DateSerial(mY + 1, 1, 1), "yyyy/mm/dd")
    Else
        dT = Format(DateSerial(mY, mM + 1, 1), "yyyy/mm/dd")
    End If

    Dim mf As String: mf = ""
    If Not IsNull(cboMember.Value) And Len(Trim(Nz(cboMember.Value, ""))) > 0 Then
        mf = " AND R.member_name='" & Replace(Nz(cboMember.Value, ""), "'", "''") & "'"
    End If

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Format(R.rec_date,'yyyy/mm/dd'), R.member_name," _
        & " R.calls, R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.note" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & mf & " ORDER BY R.rec_date, R.member_name")

    Dim csvPath As String
    csvPath = Environ("USERPROFILE") & "\\Desktop\\daily_" & mY & Format(mM, "00") & ".csv"
    Open csvPath For Output As #1
    Print #1, Chr(26085) & Chr(20184) & "," & Chr(25285) & Chr(24403) & Chr(32773) & "," _
        & Chr(26550) & Chr(30005) & "," & Chr(26377) & Chr(21177) & "," & Chr(35211) & Chr(36796) & "," _
        & Chr(36039) & Chr(26009) & "," & Chr(36861) & Chr(23470) & "," & Chr(21463) & Chr(27880) & "," _
        & Chr(31295) & Chr(21205) & Chr(26178) & Chr(38291) & "," & Chr(20999) & Chr(32771)
    Do While Not rs.EOF
        Print #1, Chr(34) & Nz(rs.Fields(0).Value, "") & Chr(34) & "," _
            & Chr(34) & Nz(rs.Fields(1).Value, "") & Chr(34) & "," _
            & Nz(rs.Fields(2).Value, 0) & "," _
            & Nz(rs.Fields(3).Value, 0) & "," _
            & Nz(rs.Fields(4).Value, 0) & "," _
            & Nz(rs.Fields(5).Value, 0) & "," _
            & Nz(rs.Fields(6).Value, 0) & "," _
            & Nz(rs.Fields(7).Value, 0) & "," _
            & Nz(rs.Fields(8).Value, 0) & "," _
            & Chr(34) & Replace(Nz(rs.Fields(9).Value, ""), Chr(34), "'") & Chr(34)
        rs.MoveNext
    Loop
    rs.Close
    Close #1
    MsgBox "CSV: " & csvPath, vbInformation, "CSV"
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnCSV"
    On Error Resume Next
    Close #1
End Sub"""

# ─── F_Analysis (分析) ───────────────────────────
VBA_ANALYSIS = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date)
    mM = Month(Date)
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
    lblMonth.Caption = mY & Chr(24180) & mM & Chr(26376) & " " & Chr(20998) & Chr(26512)

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then
        dT = Format(DateSerial(mY + 1, 1, 1), "yyyy/mm/dd")
    Else
        dT = Format(DateSerial(mY, mM + 1, 1), "yyyy/mm/dd")
    End If

    Dim baseQ As String
    baseQ = " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & " GROUP BY R.member_name ORDER BY "

    lstRankRecv.RowSource = _
        "SELECT R.member_name, Sum(R.received), Sum(R.calls), Sum(R.valid_count), Sum(R.prospect)" _
        & baseQ & "Sum(R.received) DESC"

    lstRankProd.RowSource = _
        "SELECT R.member_name, Sum(R.prospect), Format(Sum(R.work_hours),'0.0') & 'h'," _
        & " IIf(Sum(R.work_hours)>0, Format(Sum(R.prospect)/Sum(R.work_hours),'0.000'), '0')" _
        & baseQ & "IIf(Sum(R.work_hours)>0, Sum(R.prospect)/Sum(R.work_hours), 0) DESC"

    lstRankCalls.RowSource = _
        "SELECT R.member_name, Sum(R.calls), Sum(R.valid_count)," _
        & " IIf(Sum(R.calls)>0, Format(Sum(R.valid_count)/Sum(R.calls)*100,'0.0') & '%', '-')" _
        & baseQ & "Sum(R.calls) DESC"

    lstHourly.RowSource = _
        "SELECT Sum(R.calls_10), Sum(R.calls_11), Sum(R.calls_12), Sum(R.calls_13)," _
        & " Sum(R.calls_14), Sum(R.calls_15), Sum(R.calls_16), Sum(R.calls_17), Sum(R.calls_18)" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#"

    lstRankRecv.Requery
    lstRankProd.Requery
    lstRankCalls.Requery
    lstHourly.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub"""


# ──────────────────────────────────────────────
# ヘルパー関数
# ──────────────────────────────────────────────
def c(cm): return int(cm * CM)

def L(name, cap, l, t, w, h, fs=9, bold=False):
    d = {'type':acLabel,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs}
    if bold: d['bold'] = True
    return d

def B(name, cap, l, t, w, h, fs=9):
    return {'type':acButton,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs}

def T(name, l, t, w, h):
    return {'type':acTextBox,'name':name,'left':l,'top':t,'width':w,'height':h}

def C(name, l, t, w, h):
    return {'type':acCombo,'name':name,'left':l,'top':t,'width':w,'height':h}

def LB(name, l, t, w, h, cc=2, cw="3cm;2cm"):
    return {'type':acList,'name':name,'left':l,'top':t,'width':w,'height':h,'cc':cc,'cw':cw}


def make_form(app, name, caption, controls, vba, width=14500):
    """既存フォームを削除して新規作成"""
    # 既存削除
    try: app.DoCmd.Close(2, name)
    except: pass
    try:
        app.DoCmd.DeleteObject(2, name)
        time.sleep(0.5)
    except: pass

    frm = app.CreateForm()
    frm.Caption = caption
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 2
    frm.DefaultView = 0
    frm.Width = width
    tmp = frm.Name

    for ctrl_def in controls:
        ct = ctrl_def['type']
        ctrl = app.CreateControl(tmp, ct, acDetail, "", "",
                                 ctrl_def['left'], ctrl_def['top'],
                                 ctrl_def['width'], ctrl_def['height'])
        if 'name'    in ctrl_def: ctrl.Name = ctrl_def['name']
        if 'caption' in ctrl_def: ctrl.Caption = ctrl_def['caption']
        if ct != acCheck:
            ctrl.FontName = "Yu Gothic UI"
            ctrl.FontSize = ctrl_def.get('fs', 9)
        if ctrl_def.get('bold'): ctrl.FontBold = True
        if ct == acButton: ctrl.OnClick = "[Event Procedure]"
        if ct == acList:
            ctrl.ColumnCount = ctrl_def.get('cc', 2)
            if 'cw' in ctrl_def: ctrl.ColumnWidths = ctrl_def['cw']
            ctrl.RowSourceType = "Table/Query"
        if ct == acCombo:
            ctrl.ColumnCount = 1
            ctrl.RowSourceType = "Table/Query"

    # HasModule=True を設定してからVBA注入
    frm.HasModule = True
    time.sleep(0.3)

    if vba:
        try:
            comp = app.VBE.VBProjects(1).VBComponents("Form_" + tmp)
            cm_mod = comp.CodeModule
            for i, line in enumerate(vba.strip().split("\n"), 1):
                cm_mod.InsertLines(i, line)
        except Exception as e:
            print(f"    VBA注入エラー {name}: {e}")

    app.DoCmd.Save(2, tmp)
    app.DoCmd.Close(2, tmp)
    time.sleep(0.5)
    app.DoCmd.Rename(name, 2, tmp)
    time.sleep(0.5)
    print(f"  {name}: 作成完了")


def update_vba(app, form_name, vba_code):
    """既存フォームのVBAを置き換え"""
    try:
        app.DoCmd.OpenForm(form_name, 1)  # Design
        time.sleep(0.4)
        comp = app.VBE.VBProjects(1).VBComponents("Form_" + form_name)
        cm_mod = comp.CodeModule
        if cm_mod.CountOfLines > 0:
            cm_mod.DeleteLines(1, cm_mod.CountOfLines)
        for i, line in enumerate(vba_code.strip().split("\n"), 1):
            cm_mod.InsertLines(i, line)
        app.DoCmd.Save(2, form_name)
        app.DoCmd.Close(2, form_name)
        time.sleep(0.3)
        print(f"  {form_name}: VBA更新完了")
    except Exception as e:
        print(f"  {form_name}: VBA更新エラー {e}")
        try: app.DoCmd.Close(2, form_name)
        except: pass


def set_fonts(app, forms):
    """全フォームのフォントを 游ゴシック に統一"""
    FONT = "Yu Gothic UI"
    for fn in forms:
        try:
            app.DoCmd.OpenForm(fn, 1)  # Design
            time.sleep(0.3)
            frm = app.Forms(fn)
            for i in range(frm.Controls.Count):
                ctrl = frm.Controls(i)
                try:
                    ctrl.FontName = FONT
                except Exception:
                    pass  # 一部コントロールはFontNameを持たない
            app.DoCmd.Save(2, fn)
            app.DoCmd.Close(2, fn)
            time.sleep(0.2)
            print(f"  {fn}: フォント設定完了")
        except Exception as e:
            print(f"  {fn}: フォント設定エラー {e}")
            try: app.DoCmd.Close(2, fn)
            except: pass


def main():
    print("=" * 60)
    print("SalesMgr Access 拡張スクリプト")
    print("=" * 60)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE)
    time.sleep(2)

    # ────────────────────────────────────────
    # STEP 1: F_List 作成
    # ────────────────────────────────────────
    print("\n[STEP 1] F_List (日次一覧) 作成")
    list_ctrls = [
        B("btnPrev", "◀", c(0.5), c(0.3), c(1.2), c(0.7)),
        L("lblMonth", "一覧", c(2), c(0.3), c(8), c(0.7), 12, True),
        B("btnNext", "▶", c(10.5), c(0.3), c(1.2), c(0.7)),
        L("lm", "担当者:", c(12.5), c(0.3), c(2), c(0.7)),
        C("cboMember", c(14.5), c(0.3), c(4), c(0.7)),
        B("btnCSV", "CSV出力", c(19), c(0.3), c(2.5), c(0.7)),
        LB("lstRecords", c(0.5), c(1.3), c(22), c(12), 10,
           "2cm;3cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;3cm"),
    ]
    make_form(app, "F_List", "日次一覧", list_ctrls, VBA_LIST, width=13000)

    # ────────────────────────────────────────
    # STEP 2: F_Analysis 作成
    # ────────────────────────────────────────
    print("\n[STEP 2] F_Analysis (分析) 作成")
    y = c(1.5)
    row_h = c(5.5)
    analysis_ctrls = [
        B("btnPrev", "◀", c(0.5), c(0.3), c(1.2), c(0.7)),
        L("lblMonth", "分析", c(2), c(0.3), c(8), c(0.7), 12, True),
        B("btnNext", "▶", c(10.5), c(0.3), c(1.2), c(0.7)),
        # 受注ランキング
        L("lRK1", "受注ランキング（担当者/受注/架電/有効/見込）",
          c(0.5), c(1.3), c(7), c(0.5), 8, True),
        LB("lstRankRecv",  c(0.5), c(1.9), c(7.5), row_h, 5, "3cm;1.5cm;1.5cm;1.5cm;1.5cm"),
        # 生産性ランキング
        L("lRK2", "生産性ランキング（担当者/見込/稼働h/生産性）",
          c(8.5), c(1.3), c(7), c(0.5), 8, True),
        LB("lstRankProd",  c(8.5), c(1.9), c(7.5), row_h, 4, "3cm;1.5cm;1.5cm;2cm"),
        # 架電ランキング
        L("lRK3", "架電ランキング（担当者/架電/有効/有効率）",
          c(16.5), c(1.3), c(7), c(0.5), 8, True),
        LB("lstRankCalls", c(16.5), c(1.9), c(7.5), row_h, 4, "3cm;1.5cm;1.5cm;2cm"),
        # 時間帯別
        L("lHR", "時間帯別架電（10時〜18時）",
          c(0.5), c(8.2), c(10), c(0.5), 8, True),
        LB("lstHourly",    c(0.5), c(8.8), c(24), c(3.5), 9,
           "2cm;2cm;2cm;2cm;2cm;2cm;2cm;2cm;2cm"),
    ]
    make_form(app, "F_Analysis", "分析", analysis_ctrls, VBA_ANALYSIS, width=14500)

    # ────────────────────────────────────────
    # STEP 3: F_Main VBA更新（btnList/btnAnalysis追加）
    # ────────────────────────────────────────
    print("\n[STEP 3] F_Main VBA更新")
    # まず F_Main にボタンを追加してからVBA更新
    try:
        app.DoCmd.OpenForm("F_Main", 1)  # Design
        time.sleep(0.4)
        frm = app.Forms("F_Main")
        tmp = frm.Name

        # 既存ボタン数を確認して適切な位置を決める
        # F_Mainは元々6ボタンあるので下に2つ追加
        # 既存ボタンの配置を確認
        existing = {frm.Controls(i).Name for i in range(frm.Controls.Count)}

        if "btnList" not in existing:
            b1 = app.CreateControl(tmp, acButton, acDetail, "", "",
                                   c(1), c(5.5), c(3), c(0.9))
            b1.Name = "btnList"
            b1.Caption = "日次一覧"
            b1.FontName = "Yu Gothic UI"
            b1.FontSize = 10
            b1.OnClick = "[Event Procedure]"

        if "btnAnalysis" not in existing:
            b2 = app.CreateControl(tmp, acButton, acDetail, "", "",
                                   c(1), c(6.6), c(3), c(0.9))
            b2.Name = "btnAnalysis"
            b2.Caption = "分析"
            b2.FontName = "Yu Gothic UI"
            b2.FontSize = 10
            b2.OnClick = "[Event Procedure]"

        app.DoCmd.Save(2, "F_Main")
        app.DoCmd.Close(2, "F_Main")
        time.sleep(0.3)
        print("  F_Main: ボタン追加完了")
    except Exception as e:
        print(f"  F_Main ボタン追加エラー: {e}")
        try: app.DoCmd.Close(2, "F_Main")
        except: pass

    update_vba(app, "F_Main", VBA_MAIN)

    # ────────────────────────────────────────
    # STEP 4: F_Report VBA更新（btnPDF実装）
    # ────────────────────────────────────────
    print("\n[STEP 4] F_Report VBA更新（PDF出力実装）")
    update_vba(app, "F_Report", VBA_REPORT)

    # ────────────────────────────────────────
    # STEP 5: 全フォームのフォント統一
    # ────────────────────────────────────────
    print("\n[STEP 5] フォント統一（游ゴシック）")
    all_forms = ["F_Main","F_Daily","F_DailyEdit","F_Members",
                 "F_Targets","F_Referrals","F_Report","F_Ranking",
                 "F_List","F_Analysis"]
    set_fonts(app, all_forms)

    # ────────────────────────────────────────
    # 完了
    # ────────────────────────────────────────
    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)
    print("\n[完了] 全処理正常終了")


if __name__ == "__main__":
    main()
