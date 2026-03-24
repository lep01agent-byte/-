# -*- coding: utf-8 -*-
"""Phase 3-6: フォーム8個 + VBA + スタートアップ設定"""
import os, sys, time
import win32com.client

BE_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

acLabel=100; acButton=104; acTextBox=109; acCombo=111; acList=110; acCheck=106; acDetail=0
CM = 567

def make_form(app, name, caption, controls, vba, popup=False, modal=False, width=14000):
    try:
        app.DoCmd.Close(2, name)
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
    if popup: frm.PopUp = True
    if modal: frm.Modal = True
    tmp = frm.Name

    for c in controls:
        ct = c['type']
        ctrl = app.CreateControl(tmp, ct, acDetail, "", "", c['left'], c['top'], c['width'], c['height'])
        if 'name' in c: ctrl.Name = c['name']
        if 'caption' in c: ctrl.Caption = c['caption']
        if ct != acCheck:
            ctrl.FontName = "Yu Gothic UI"
            ctrl.FontSize = c.get('fs', 9)
        if c.get('bold'): ctrl.FontBold = True
        if 'fc' in c: ctrl.ForeColor = c['fc']
        if ct == acButton and c.get('click'): ctrl.OnClick = "[Event Procedure]"
        if ct == acList:
            ctrl.ColumnCount = c.get('cc', 2)
            if 'cw' in c: ctrl.ColumnWidths = c['cw']
            ctrl.RowSourceType = "Table/Query"
        if ct == acCombo:
            ctrl.ColumnCount = 1; ctrl.RowSourceType = "Table/Query"

    # VBA注入（リネーム前に行う。一時名でVBComponentにアクセス）
    if vba:
        try:
            comp = app.VBE.VBProjects(1).VBComponents("Form_" + tmp)
            cm = comp.CodeModule
            for i, line in enumerate(vba.strip().split("\n"), 1):
                cm.InsertLines(i, line)
        except Exception as e:
            print(f"    VBA: {e}")

    app.DoCmd.Save(2, tmp)
    app.DoCmd.Close(2, tmp)
    time.sleep(0.3)
    app.DoCmd.Rename(name, 2, tmp)
    time.sleep(0.3)

    print(f"  {name} OK")

def L(name, cap, l, t, w, h, fs=9, bold=False, fc=0):
    d = {'type':acLabel,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs}
    if bold: d['bold']=True
    if fc: d['fc']=fc
    return d

def B(name, cap, l, t, w, h, fs=9):
    return {'type':acButton,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs,'click':True}

def T(name, l, t, w, h):
    return {'type':acTextBox,'name':name,'left':l,'top':t,'width':w,'height':h}

def C(name, l, t, w, h):
    return {'type':acCombo,'name':name,'left':l,'top':t,'width':w,'height':h}

def LB(name, l, t, w, h, cc=2, cw=""):
    return {'type':acList,'name':name,'left':l,'top':t,'width':w,'height':h,'cc':cc,'cw':cw}

def CK(name, l, t, w, h):
    return {'type':acCheck,'name':name,'left':l,'top':t,'width':w,'height':h}

def c(n): return int(CM*n)

# ======================== VBA CODES ========================

VBA_MAIN = '''Option Compare Database
Option Explicit
Private Sub btnDaily_Click(): DoCmd.OpenForm "F_Daily": End Sub
Private Sub btnTargets_Click(): DoCmd.OpenForm "F_Targets": End Sub
Private Sub btnReferrals_Click(): DoCmd.OpenForm "F_Referrals": End Sub
Private Sub btnReport_Click(): DoCmd.OpenForm "F_Report": End Sub
Private Sub btnRanking_Click(): DoCmd.OpenForm "F_Ranking": End Sub
Private Sub btnMembers_Click(): DoCmd.OpenForm "F_Members": End Sub'''

VBA_MEMBERS = '''Option Compare Database
Option Explicit
Private Sub Form_Open(Cancel As Integer): LD: End Sub
Private Sub LD()
    lstActive.RowSource = "SELECT ID,member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    lstInactive.RowSource = "SELECT ID,member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name"
    lstActive.Requery: lstInactive.Requery
End Sub
Private Sub btnAdd_Click()
    Dim nm As String: nm = Trim(Nz(txtNewName.Value,""))
    If nm="" Then MsgBox "Name required",vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_MEMBERS (member_name,active) VALUES ('" & Replace(nm,"'","''") & "',True)", dbFailOnError
    txtNewName.Value = "": LD
End Sub
Private Sub btnDeactivate_Click()
    If IsNull(lstActive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=False WHERE ID=" & lstActive.Value, dbFailOnError: LD
End Sub
Private Sub btnActivate_Click()
    If IsNull(lstInactive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=True WHERE ID=" & lstInactive.Value, dbFailOnError: LD
End Sub'''

VBA_DAILY = '''Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer
Private Sub Form_Open(Cancel As Integer): mY=Year(Date): mM=Month(Date): LD: End Sub
Private Sub LD()
    Me.Caption = mY & "/" & Format(mM,"00") & " Daily"
    lblMonth.Caption = mY & "/" & mM
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    Dim s As String
    s = "SELECT R.ID, Format(R.rec_date,''yyyy/mm/dd''), R.member_name, R.calls, R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.[note] FROM T_RECORDS AS R WHERE R.rec_date >= #" & Format(DateSerial(mY,mM,1),"yyyy/mm/dd") & "# AND R.rec_date < #" & Format(DateSerial(mY,mM+1,1),"yyyy/mm/dd") & "# "
    If Nz(cboMember.Value,"") <> "" Then s = s & "AND R.member_name=''" & cboMember.Value & "'' "
    s = s & "ORDER BY R.rec_date DESC, R.member_name"
    lstRecords.RowSource = s: lstRecords.Requery
End Sub
Private Sub btnPrev_Click(): If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1: End If: LD: End Sub
Private Sub btnNext_Click(): If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1: End If: LD: End Sub
Private Sub cboMember_AfterUpdate(): LD: End Sub
Private Sub btnAdd_Click(): DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"ADD|" & mY & "|" & mM: LD: End Sub
Private Sub btnEdit_Click()
    If IsNull(lstRecords.Value) Then MsgBox "Select record",vbExclamation: Exit Sub
    DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"EDIT|" & lstRecords.Value: LD
End Sub
Private Sub btnDelete_Click()
    If IsNull(lstRecords.Value) Then MsgBox "Select record",vbExclamation: Exit Sub
    If MsgBox("Delete?",vbYesNo+vbQuestion)=vbYes Then CurrentDb.Execute "DELETE FROM T_RECORDS WHERE ID=" & lstRecords.Value, dbFailOnError: LD
End Sub
Private Sub lstRecords_DblClick(Cancel As Integer): btnEdit_Click: End Sub'''

VBA_DAILY_EDIT = '''Option Compare Database
Option Explicit
Private mMode As String, mID As Long
Private Sub Form_Open(Cancel As Integer)
    Dim p() As String: p = Split(Nz(Me.OpenArgs,"ADD"),"|"): mMode = p(0)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    If mMode="EDIT" And UBound(p)>=1 Then mID = CLng(p(1)): Me.Caption = "Edit": LR Else mID = 0: Me.Caption = "New"
    If mID=0 Then
        If UBound(p)>=2 Then txtRecDate.Value = Format(DateSerial(CInt(p(1)),CInt(p(2)),Day(Date)),"yyyy/mm/dd") Else txtRecDate.Value = Format(Date,"yyyy/mm/dd")
        txtWorkHours.Value = 8: chkWorkDay.Value = True
    End If
End Sub
Private Sub LR()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_RECORDS WHERE ID=" & mID)
    If Not rs.EOF Then
        txtRecDate.Value = Format(rs("rec_date"),"yyyy/mm/dd"): cboMember.Value = rs("member_name")
        Dim h As Integer: For h = 10 To 18: Me("txtC" & h).Value = Nz(rs("calls_" & h),0): Next h
        txtValid.Value = Nz(rs("valid_count"),0): txtProspect.Value = Nz(rs("prospect"),0)
        txtDoc.Value = Nz(rs("doc"),0): txtFollow.Value = Nz(rs("follow_up"),0)
        txtReceived.Value = Nz(rs("received"),0): txtReferral.Value = Nz(rs("referral"),0)
        txtWorkHours.Value = Nz(rs("work_hours"),8): chkWorkDay.Value = rs("work_day")
        txtNote.Value = Nz(rs("note"),"")
    End If: rs.Close
End Sub
Private Sub btnSave_Click()
    If Nz(txtRecDate.Value,"")="" Or Nz(cboMember.Value,"")="" Then MsgBox "Required",vbExclamation: Exit Sub
    Dim dt As Date: dt = CDate(txtRecDate.Value)
    Dim cl As Long, h As Integer: cl = 0: For h = 10 To 18: cl = cl + Nz(Me("txtC" & h),0): Next h
    Dim s As String
    If mID = 0 Then
        s = "INSERT INTO T_RECORDS (rec_date,member_name,calls,calls_10,calls_11,calls_12,calls_13,calls_14,calls_15,calls_16,calls_17,calls_18,valid_count,prospect,doc,follow_up,received,work_hours,[note],referral,work_day) VALUES (#" & Format(dt,"yyyy/mm/dd") & "#,''" & cboMember.Value & "''," & cl
        For h = 10 To 18: s = s & "," & Nz(Me("txtC" & h),0): Next h
        s = s & "," & Nz(txtValid,0) & "," & Nz(txtProspect,0) & "," & Nz(txtDoc,0) & "," & Nz(txtFollow,0) & "," & Nz(txtReceived,0) & "," & Nz(txtWorkHours,8) & ",''" & Replace(Nz(txtNote,""),"'","''") & "''," & Nz(txtReferral,0) & "," & IIf(chkWorkDay.Value,"True","False") & ")"
    Else
        s = "UPDATE T_RECORDS SET rec_date=#" & Format(dt,"yyyy/mm/dd") & "#,member_name=''" & cboMember.Value & "'',calls=" & cl
        For h = 10 To 18: s = s & ",calls_" & h & "=" & Nz(Me("txtC" & h),0): Next h
        s = s & ",valid_count=" & Nz(txtValid,0) & ",prospect=" & Nz(txtProspect,0) & ",doc=" & Nz(txtDoc,0) & ",follow_up=" & Nz(txtFollow,0) & ",received=" & Nz(txtReceived,0) & ",work_hours=" & Nz(txtWorkHours,8) & ",[note]=''" & Replace(Nz(txtNote,""),"'","''") & "'',referral=" & Nz(txtReferral,0) & ",work_day=" & IIf(chkWorkDay.Value,"True","False") & " WHERE ID=" & mID
    End If
    CurrentDb.Execute s, dbFailOnError: DoCmd.Close acForm, Me.Name
End Sub
Private Sub btnCancel_Click(): DoCmd.Close acForm, Me.Name: End Sub'''

VBA_TARGETS = '''Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer
Private Sub Form_Open(Cancel As Integer): mY=Year(Date): mM=Month(Date): cboMember.RowSource="SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name": LD: End Sub
Private Sub LD()
    lblMonth.Caption = mY & "/" & mM
    lstTargets.RowSource = "SELECT T.ID,T.member_name,T.target_calls,T.target_valid,T.target_prospect,T.target_received,T.target_referral,T.plan_days,T.work_hours_per_day FROM T_MEMBER_TARGETS AS T WHERE T.target_year=" & mY & " AND T.target_month=" & mM & " ORDER BY T.member_name"
    lstTargets.Requery
End Sub
Private Sub btnPrev_Click(): If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1: End If: LD: End Sub
Private Sub btnNext_Click(): If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1: End If: LD: End Sub
Private Sub btnSave_Click()
    If Nz(cboMember.Value,"")="" Then MsgBox "Select member",vbExclamation: Exit Sub
    Dim nm As String: nm = cboMember.Value
    Dim n As Long: n = DCount("*","T_MEMBER_TARGETS","member_name=''" & nm & "'' AND target_year=" & mY & " AND target_month=" & mM)
    If n > 0 Then
        CurrentDb.Execute "UPDATE T_MEMBER_TARGETS SET plan_days=" & Nz(txtPlanDays,20) & ",work_hours_per_day=" & Nz(txtHoursPerDay,8) & ",target_calls=" & Nz(txtTgtCalls,0) & ",target_valid=" & Nz(txtTgtValid,0) & ",target_prospect=" & Nz(txtTgtProspect,0) & ",target_received=" & Nz(txtTgtReceived,0) & ",target_referral=" & Nz(txtTgtReferral,0) & " WHERE member_name=''" & nm & "'' AND target_year=" & mY & " AND target_month=" & mM, dbFailOnError
    Else
        CurrentDb.Execute "INSERT INTO T_MEMBER_TARGETS (member_name,target_year,target_month,plan_days,work_hours_per_day,target_calls,target_valid,target_prospect,target_received,target_referral) VALUES (''" & nm & "''," & mY & "," & mM & "," & Nz(txtPlanDays,20) & "," & Nz(txtHoursPerDay,8) & "," & Nz(txtTgtCalls,0) & "," & Nz(txtTgtValid,0) & "," & Nz(txtTgtProspect,0) & "," & Nz(txtTgtReceived,0) & "," & Nz(txtTgtReferral,0) & ")", dbFailOnError
    End If
    LD: MsgBox nm & " saved", vbInformation
End Sub
Private Sub btnLoadPrev_Click()
    If Nz(cboMember.Value,"")="" Then Exit Sub
    Dim py As Integer, pm As Integer: py=mY: pm=mM: If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name=''" & cboMember.Value & "'' AND target_year=" & py & " AND target_month=" & pm)
    If Not rs.EOF Then txtPlanDays.Value=rs("plan_days"): txtHoursPerDay.Value=rs("work_hours_per_day"): txtTgtCalls.Value=rs("target_calls"): txtTgtValid.Value=rs("target_valid"): txtTgtProspect.Value=rs("target_prospect"): txtTgtReceived.Value=rs("target_received"): txtTgtReferral.Value=rs("target_referral") Else MsgBox "No prev data",vbInformation
    rs.Close
End Sub
Private Sub btnDelete_Click()
    If IsNull(lstTargets.Value) Then Exit Sub
    If MsgBox("Delete?",vbYesNo+vbQuestion)=vbYes Then CurrentDb.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & lstTargets.Value, dbFailOnError: LD
End Sub'''

VBA_REFERRALS = '''Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer
Private Sub Form_Open(Cancel As Integer): mY=Year(Date): mM=Month(Date): cboMember.RowSource="SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name": LD: End Sub
Private Sub LD()
    lblMonth.Caption = mY & "/" & mM: txtRefDate.Value = Format(Date,"yyyy/mm/dd")
    lstRefs.RowSource = "SELECT R.ID,Format(R.rec_date,''yyyy/mm/dd''),R.member_name,R.ref_count FROM T_REFERRALS AS R WHERE R.rec_date>=#" & Format(DateSerial(mY,mM,1),"yyyy/mm/dd") & "# AND R.rec_date<#" & Format(DateSerial(mY,mM+1,1),"yyyy/mm/dd") & "# ORDER BY R.rec_date DESC,R.member_name"
    lstRefs.Requery
End Sub
Private Sub btnPrev_Click(): If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1: End If: LD: End Sub
Private Sub btnNext_Click(): If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1: End If: LD: End Sub
Private Sub btnAdd_Click()
    If Nz(cboMember.Value,"")="" Or Nz(txtRefDate.Value,"")="" Then MsgBox "Required",vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_REFERRALS (rec_date,member_name,ref_count) VALUES (#" & Format(CDate(txtRefDate.Value),"yyyy/mm/dd") & "#,''" & cboMember.Value & "''," & Nz(txtRefCount,0) & ")", dbFailOnError: LD
End Sub
Private Sub btnDelete_Click()
    If IsNull(lstRefs.Value) Then Exit Sub
    If MsgBox("Delete?",vbYesNo+vbQuestion)=vbYes Then CurrentDb.Execute "DELETE FROM T_REFERRALS WHERE ID=" & lstRefs.Value, dbFailOnError: LD
End Sub'''

VBA_REPORT = '''Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer
Private Sub Form_Open(Cancel As Integer): mY=Year(Date): mM=Month(Date): LR: End Sub
Private Sub btnPrev_Click(): If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1: End If: LR: End Sub
Private Sub btnNext_Click(): If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1: End If: LR: End Sub
Private Sub btnPDF_Click(): MsgBox "PDF export - future", vbInformation: End Sub
Private Sub LR()
    lblMonth.Caption = mY & "/" & mM & " Report"
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY,mM,1),"yyyy/mm/dd"): dT = Format(DateSerial(mY,mM+1,1),"yyyy/mm/dd")
    Dim qd As DAO.QueryDef, rs As DAO.Recordset
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(mY,mM,1): qd.Parameters("prmDateTo") = DateSerial(mY,mM+1,1)
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim tC As Long, tV As Long, tP As Long, tR As Long, tH As Double
    If Not rs.EOF Then tC=Nz(rs("sum_calls"),0): tV=Nz(rs("sum_valid"),0): tP=Nz(rs("sum_prospect"),0): tR=Nz(rs("sum_received"),0): tH=Nz(rs("sum_hours"),0)
    rs.Close
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear") = mY: qd.Parameters("prmMonth") = mM
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim gC As Long, gV As Long, gP As Long, gR As Long
    If Not rs.EOF Then gC=Nz(rs("sum_tgt_calls"),0): gV=Nz(rs("sum_tgt_valid"),0): gP=Nz(rs("sum_tgt_prospect"),0): gR=Nz(rs("sum_tgt_received"),0)
    rs.Close
    Dim py As Integer, pm As Integer: py=mY: pm=mM: If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(py,pm,1): qd.Parameters("prmDateTo") = DateSerial(py,pm+1,1)
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim pC As Long, pV As Long, pP As Long, pR As Long
    If Not rs.EOF Then pC=Nz(rs("sum_calls"),0): pV=Nz(rs("sum_valid"),0): pP=Nz(rs("sum_prospect"),0): pR=Nz(rs("sum_received"),0)
    rs.Close
    lblCalls.Caption = Format(tC,"#,##0") & " / " & Format(gC,"#,##0")
    lblValid.Caption = Format(tV,"#,##0") & " / " & Format(gV,"#,##0")
    lblProsp.Caption = Format(tP,"#,##0") & " / " & Format(gP,"#,##0")
    lblRecv.Caption = Format(tR,"#,##0") & " / " & Format(gR,"#,##0")
    lblCallsPrev.Caption = "prev " & Format(pC,"#,##0") & " " & IIf(tC>pC,Chr(9650) & Format(tC-pC,"#,##0"),IIf(tC<pC,Chr(9660) & Format(pC-tC,"#,##0"),"-"))
    lblValidPrev.Caption = "prev " & Format(pV,"#,##0") & " " & IIf(tV>pV,Chr(9650) & Format(tV-pV,"#,##0"),IIf(tV<pV,Chr(9660) & Format(pV-tV,"#,##0"),"-"))
    lblProspPrev.Caption = "prev " & Format(pP,"#,##0") & " " & IIf(tP>pP,Chr(9650) & Format(tP-pP,"#,##0"),IIf(tP<pP,Chr(9660) & Format(pP-tP,"#,##0"),"-"))
    lblRecvPrev.Caption = "prev " & Format(pR,"#,##0") & " " & IIf(tR>pR,Chr(9650) & Format(tR-pR,"#,##0"),IIf(tR<pR,Chr(9660) & Format(pR-tR,"#,##0"),"-"))
    lblValidRate.Caption = IIf(tC>0, Format(tV/tC*100,"0.0") & "%", "-")
    lblRecvRate.Caption = IIf(tC>0, Format(tR/tC*100,"0.0") & "%", "-")
    lblHours.Caption = Format(tH,"#,##0") & "h"
    lblProductivity.Caption = IIf(tH>0, Format(tP/tH,"0.000"), "-")
    Dim al As String: al = ""
    If gR > 0 Then: If tR < gR Then al = al & "Recv: " & (gR-tR) & " remaining" & vbCrLf Else al = al & "Recv: Goal!" & vbCrLf
    If gP > 0 Then: If tP < gP Then al = al & "Prosp: " & (gP-tP) & " remaining" Else al = al & "Prosp: Goal!"
    lblAlert.Caption = al
    lstRankRef.RowSource = "SELECT R.member_name, Sum(R.referral) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRankRef.Requery: lstRankRecv.Requery: lstRankProsp.Requery
End Sub'''

VBA_RANKING = '''Option Compare Database
Option Explicit
Private mY As Integer, mM As Integer
Private Sub Form_Open(Cancel As Integer): mY=Year(Date): mM=Month(Date): LD: End Sub
Private Sub btnPrev_Click(): If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1: End If: LD: End Sub
Private Sub btnNext_Click(): If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1: End If: LD: End Sub
Private Sub LD()
    lblMonth.Caption = mY & "/" & mM
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY,mM,1),"yyyy/mm/dd"): dT = Format(DateSerial(mY,mM+1,1),"yyyy/mm/dd")
    lstRef.RowSource = "SELECT R.member_name, Sum(R.referral) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dF & "# AND R.rec_date<#" & dT & "# GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRef.Requery: lstRecv.Requery: lstProsp.Requery
End Sub'''

# ======================== MAIN ========================

def main():
    print("Phase 3-6: Forms + VBA + Startup")
    print("=" * 50)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE_PATH)
    time.sleep(2)

    # 1. F_Main
    print("Phase 3: Forms")
    ctrls = [L("lblTitle","SalesMgr 営業管理",c(1),0,c(12),c(1.5),18,True,3877662)]
    for i,(bn,bc) in enumerate([("btnDaily","日次登録"),("btnTargets","目標設定"),("btnReferrals","送客登録"),("btnReport","レポート"),("btnRanking","ランキング"),("btnMembers","担当者管理")]):
        ctrls.append(B(bn,bc,c(0.5)+(i%2)*c(7),c(2)+(i//2)*c(1.5),c(6),c(1.2),11))
    make_form(app,"F_Main","SalesMgr 営業管理",ctrls,VBA_MAIN,width=c(15))

    # 2. F_Members
    make_form(app,"F_Members","担当者管理",[
        L("lblTitle","担当者管理",c(0.5),c(0.3),c(6),c(1),14,True),
        L("lblAdd","新規:",c(0.5),c(1.5),c(1.5),c(0.6)),T("txtNewName",c(2),c(1.5),c(4),c(0.6)),B("btnAdd","追加",c(6.5),c(1.5),c(2),c(0.6)),
        L("lblA","有効",c(0.5),c(2.8),c(4),c(0.6),9,True),LB("lstActive",c(0.5),c(3.5),c(5),c(5),2,"0cm;4cm"),
        B("btnDeactivate","無効▶",c(6),c(5),c(3),c(0.7)),
        L("lblI","無効",c(10),c(2.8),c(4),c(0.6),9,True),LB("lstInactive",c(10),c(3.5),c(5),c(5),2,"0cm;4cm"),
        B("btnActivate","◀有効",c(6),c(6.5),c(3),c(0.7)),
    ],VBA_MEMBERS)

    # 3. F_DailyEdit
    ctrls_de = [
        L("lbl1","日付:",c(0.3),c(0.5),c(1.5),c(0.6)),T("txtRecDate",c(2),c(0.5),c(3),c(0.6)),
        L("lbl2","担当者:",c(0.3),c(1.3),c(1.5),c(0.6)),C("cboMember",c(2),c(1.3),c(3),c(0.6)),
        L("lbl3","稼働h:",c(0.3),c(2.1),c(1.5),c(0.6)),T("txtWorkHours",c(2),c(2.1),c(1.5),c(0.6)),
        L("lbl3b","出勤:",c(4),c(2.1),c(1),c(0.6)),CK("chkWorkDay",c(5),c(2.1),c(0.5),c(0.5)),
        L("lblHr","時間帯別架電:",c(0.3),c(3),c(3),c(0.6),9,True),
    ]
    y = c(3.7)
    for h in range(10,19):
        ctrls_de.append(L(f"lH{h}",f"{h}時:",c(0.3),y,c(1.2),c(0.5)))
        ctrls_de.append(T(f"txtC{h}",c(1.5),y,c(1.5),c(0.5)))
        y += c(0.6)
    y += c(0.3)
    for lb,nm in [("有効:","txtValid"),("見込:","txtProspect"),("資料:","txtDoc"),("追客:","txtFollow"),("受注:","txtReceived"),("送客:","txtReferral")]:
        ctrls_de.append(L(f"l{nm}",lb,c(0.3),y,c(1.5),c(0.5)))
        ctrls_de.append(T(nm,c(2),y,c(1.5),c(0.5)))
        y += c(0.6)
    ctrls_de.append(L("lNote","備考:",c(0.3),y,c(1.5),c(0.5)))
    ctrls_de.append(T("txtNote",c(2),y,c(4),c(0.8)))
    y += c(1.2)
    ctrls_de.append(B("btnSave","保存",c(0.5),y,c(2.5),c(0.7)))
    ctrls_de.append(B("btnCancel","キャンセル",c(3.5),y,c(2.5),c(0.7)))
    make_form(app,"F_DailyEdit","日次レコード",ctrls_de,VBA_DAILY_EDIT,popup=True,modal=True,width=c(7))

    # 4. F_Daily
    make_form(app,"F_Daily","日次一覧",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),L("lblMonth","YM",c(2),c(0.3),c(5),c(0.7),12,True),B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lblF","担当者:",c(10),c(0.3),c(2),c(0.7)),C("cboMember",c(12),c(0.3),c(3.5),c(0.7)),
        B("btnAdd","新規登録",c(0.5),c(1.3),c(2.5),c(0.7)),B("btnEdit","編集",c(3.5),c(1.3),c(2),c(0.7)),B("btnDelete","削除",c(6),c(1.3),c(2),c(0.7)),
        LB("lstRecords",c(0.5),c(2.5),c(22),c(12),11,"0cm;2.5cm;2.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;3cm"),
    ],VBA_DAILY)

    # 5. F_Targets
    ctrls_tgt = [
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),L("lblMonth","YM",c(2),c(0.3),c(5),c(0.7),12,True),B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lm","担当者:",c(0.5),c(1.5),c(2),c(0.6)),C("cboMember",c(2.5),c(1.5),c(3.5),c(0.6)),B("btnLoadPrev","前月コピー",c(6.5),c(1.5),c(2.5),c(0.6)),
    ]
    y = c(2.4)
    for lb,nm in [("稼働日数:","txtPlanDays"),("日稼働h:","txtHoursPerDay"),("架電:","txtTgtCalls"),("有効:","txtTgtValid"),("見込:","txtTgtProspect"),("受注:","txtTgtReceived"),("送客:","txtTgtReferral")]:
        ctrls_tgt.append(L(f"l{nm}",lb,c(0.5),y,c(2),c(0.6)))
        ctrls_tgt.append(T(nm,c(2.5),y,c(2.5),c(0.6)))
        y += c(0.7)
    ctrls_tgt.append(B("btnSave","保存",c(0.5),y+c(0.2),c(2.5),c(0.7)))
    ctrls_tgt.append(B("btnDelete","削除",c(3.5),y+c(0.2),c(2.5),c(0.7)))
    ctrls_tgt.append(LB("lstTargets",c(0.5),y+c(1.3),c(20),c(4),9,"0cm;2.5cm;2cm;2cm;2cm;2cm;2cm;2cm;2cm"))
    make_form(app,"F_Targets","目標設定",ctrls_tgt,VBA_TARGETS)

    # 6. F_Referrals
    make_form(app,"F_Referrals","送客登録",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),L("lblMonth","YM",c(2),c(0.3),c(5),c(0.7),12,True),B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("ld","日付:",c(0.5),c(1.5),c(1.2),c(0.6)),T("txtRefDate",c(1.7),c(1.5),c(2.5),c(0.6)),
        L("lm","担当者:",c(4.5),c(1.5),c(1.5),c(0.6)),C("cboMember",c(6),c(1.5),c(3),c(0.6)),
        L("lc","件数:",c(9.5),c(1.5),c(1),c(0.6)),T("txtRefCount",c(10.5),c(1.5),c(1.5),c(0.6)),
        B("btnAdd","追加",c(12.5),c(1.5),c(2),c(0.6)),B("btnDelete","削除",c(0.5),c(2.5),c(2),c(0.6)),
        LB("lstRefs",c(0.5),c(3.3),c(15),c(7),4,"0cm;3cm;3cm;2cm"),
    ],VBA_REFERRALS)

    # 7. F_Report
    ctrls_rpt = [
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),L("lblMonth","Report",c(2),c(0.3),c(8),c(0.7),12,True),B("btnNext","▶",c(10.5),c(0.3),c(1.2),c(0.7)),
        B("btnPDF","PDF出力",c(12.5),c(0.3),c(3),c(0.7)),
        L("lblAlert","",c(0.5),c(1.3),c(20),c(1),9,False,2498780),
    ]
    y = c(2.6)
    for i,(title,nv,ns) in enumerate([("架電件数","lblCalls","lblCallsPrev"),("有効件数","lblValid","lblValidPrev"),("見込件数","lblProsp","lblProspPrev"),("受注件数","lblRecv","lblRecvPrev")]):
        x = c(0.5)+i*c(5)
        ctrls_rpt.append(L(f"lK{i}",title,x,y,c(4.5),c(0.4),7,False,6710886))
        ctrls_rpt.append(L(nv,"-",x,y+c(0.4),c(4.5),c(0.7),11,True))
        ctrls_rpt.append(L(ns,"",x,y+c(1.1),c(4.5),c(0.4),7,False,10066329))
    y += c(1.8)
    for i,(title,nv) in enumerate([("有効率","lblValidRate"),("受注率","lblRecvRate"),("稼働時間","lblHours"),("生産性","lblProductivity")]):
        x = c(0.5)+i*c(5)
        ctrls_rpt.append(L(f"lK2{i}",title,x,y,c(4.5),c(0.4),7,False,6710886))
        ctrls_rpt.append(L(nv,"-",x,y+c(0.4),c(4.5),c(0.7),11,True))
    y += c(1.5)
    ctrls_rpt.append(L("lR1","送客ランキング",c(0.5),y,c(6),c(0.6),9,True))
    ctrls_rpt.append(L("lR2","受注ランキング",c(7),y,c(6),c(0.6),9,True))
    ctrls_rpt.append(L("lR3","見込ランキング",c(13.5),y,c(6),c(0.6),9,True))
    y += c(0.7)
    ctrls_rpt.append(LB("lstRankRef",c(0.5),y,c(6),c(5),2,"3cm;2cm"))
    ctrls_rpt.append(LB("lstRankRecv",c(7),y,c(6),c(5),2,"3cm;2cm"))
    ctrls_rpt.append(LB("lstRankProsp",c(13.5),y,c(6),c(5),2,"3cm;2cm"))
    make_form(app,"F_Report","レポート",ctrls_rpt,VBA_REPORT)

    # 8. F_Ranking
    make_form(app,"F_Ranking","ランキング",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),L("lblMonth","YM",c(2),c(0.3),c(5),c(0.7),12,True),B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lR1","送客",c(0.5),c(1.3),c(6),c(0.6),9,True),L("lR2","受注",c(7),c(1.3),c(6),c(0.6),9,True),L("lR3","見込",c(13.5),c(1.3),c(6),c(0.6),9,True),
        LB("lstRef",c(0.5),c(2),c(6),c(7),2,"3cm;2cm"),LB("lstRecv",c(7),c(2),c(6),c(7),2,"3cm;2cm"),LB("lstProsp",c(13.5),c(2),c(6),c(7),2,"3cm;2cm"),
    ],VBA_RANKING)

    print("\nPhase 6: Startup settings")
    try:
        db = app.CurrentDb()
        props = db.Properties
        for pn,pt,pv in [("StartUpForm",10,"F_Main"),("AppTitle",10,"SalesMgr 営業管理"),("StartUpShowDBWindow",1,False)]:
            try: props(pn).Value = pv
            except:
                prop = db.CreateProperty(pn, pt, pv)
                props.Append(prop)
        print("  Startup: F_Main, Title: SalesMgr")
    except Exception as e:
        print(f"  Startup error: {e}")

    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)
    print("\nDone! Open SalesMgr_BE.accdb")


if __name__ == "__main__":
    main()
