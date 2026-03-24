# -*- coding: utf-8 -*-
"""
enhance_design.py
  SalesMgr Access フォーム UI/UXデザイン全面刷新

  方式: 全8フォームを削除→新規作成（CreateForm/CreateControl）
        コントロール作成時にデザインプロパティを一括設定。
        外部COM経由で既存フォームのコントロールへのアクセスは
        制限があるため、再作成方式を採用。

  デザインコンセプト:
    - フラットデザイン (SpecialEffect=Flat)
    - ダークブルー（インディゴ）＋ ホワイト / ライトグレー
    - 游ゴシック (Yu Gothic UI) 統一
    - ゼブラストライプ（リストボックス交互色）
    - セマンティックカラーボタン

使い方:
  C:\\Users\\agentcode01\\AppData\\Local\\Programs\\Python\\Python312\\python.exe enhance_design.py
"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

# ════════════════════════════════════════════════════════════════
# カラーパレット (R + G*256 + B*65536)
# ════════════════════════════════════════════════════════════════
def rgb(r, g, b):
    return r + g * 256 + b * 65536

C_BG        = rgb(248, 250, 252)   # フォーム背景: 極薄グレー
C_WHITE     = rgb(255, 255, 255)   # 入力欄・リスト背景
C_NAVY      = rgb(30,  41,  59 )   # 主テキスト: ダークネイビー
C_MUTED     = rgb(100, 116, 139)   # サブテキスト
C_BORDER    = rgb(203, 213, 225)   # 入力ボーダー
C_ALT       = rgb(239, 246, 255)   # リスト交互行: 極薄ブルー
C_PRIMARY   = rgb(37,  99,  235)   # プライマリ青
C_DANGER    = rgb(185, 28,  28 )   # 削除赤
C_SECONDARY = rgb(71,  85,  105)   # ナビ/キャンセル グレー
C_SUCCESS   = rgb(4,   120, 87 )   # 追加/有効化 緑
C_WARNING   = rgb(180, 83,  9  )   # 無効化 アンバー
C_INFO      = rgb(14,  116, 144)   # PDF シアン
C_MAIN_BTN  = rgb(30,  58,  138)   # F_Main ボタン 濃藍
C_ACCENT    = rgb(37,  99,  235)   # KPI値テキスト
C_TITLE     = rgb(15,  23,  42 )   # タイトル最暗
C_HEADER_TXT= rgb(102, 102, 102)   # 列ヘッダー系ラベル

# ════════════════════════════════════════════════════════════════
# Access 制御タイプ定数
# ════════════════════════════════════════════════════════════════
acLabel=100; acButton=104; acTextBox=109; acCombo=111; acList=110; acCheck=106; acDetail=0
CM = 567   # 1cm = 567 twips

# ════════════════════════════════════════════════════════════════
# ヘルパー: コントロール定義辞書
# ════════════════════════════════════════════════════════════════
def c(n): return int(CM * n)

def L(name, cap, l, t, w, h, fs=9, bold=False, fc=None):
    d = {'type':acLabel,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs}
    if bold: d['bold'] = True
    if fc is not None: d['fc'] = fc
    return d

def B(name, cap, l, t, w, h, fs=9):
    return {'type':acButton,'name':name,'caption':cap,'left':l,'top':t,'width':w,'height':h,'fs':fs}

def T(name, l, t, w, h):
    return {'type':acTextBox,'name':name,'left':l,'top':t,'width':w,'height':h}

def C(name, l, t, w, h):
    return {'type':acCombo,'name':name,'left':l,'top':t,'width':w,'height':h}

def LB(name, l, t, w, h, cc=2, cw=''):
    return {'type':acList,'name':name,'left':l,'top':t,'width':w,'height':h,'cc':cc,'cw':cw}

def CK(name, l, t, w, h):
    return {'type':acCheck,'name':name,'left':l,'top':t,'width':w,'height':h}

# ════════════════════════════════════════════════════════════════
# ボタンのセマンティックカラー決定
# ════════════════════════════════════════════════════════════════
def get_btn_color(name, form_name):
    n = (name or '').lower()
    if form_name == 'F_Main':
        return C_MAIN_BTN
    if 'deactivate' in n: return C_WARNING
    if 'delete' in n:     return C_DANGER
    if 'prev' in n or 'next' in n or 'cancel' in n: return C_SECONDARY
    if 'pdf' in n:        return C_INFO
    if 'activate' in n:   return C_SUCCESS
    return C_PRIMARY

# ════════════════════════════════════════════════════════════════
# コントロール作成後にデザインプロパティを適用
# ════════════════════════════════════════════════════════════════
def sp(ctrl, prop, val):
    """1プロパティを安全に設定（失敗しても次に進む）"""
    try:
        setattr(ctrl, prop, val)
    except Exception:
        pass

def apply_design(ctrl, ct, name, form_name):
    n = (name or '').lower()

    if ct == acLabel:
        sp(ctrl, 'BackStyle',   0)   # Transparent
        sp(ctrl, 'BorderStyle', 0)
        # ForeColor: 背景が薄いのでダーク系で統一（白文字NG）
        if n in ('lblcalls','lblvalid','lblprosp','lblrecv'):
            sp(ctrl, 'ForeColor', C_ACCENT)
            sp(ctrl, 'FontSize',  10)
            sp(ctrl, 'FontBold',  True)
        elif 'month' in n and n.startswith('lbl'):
            sp(ctrl, 'ForeColor', C_TITLE)
            sp(ctrl, 'FontSize',  11)
            sp(ctrl, 'FontBold',  True)
        elif 'alert' in n:
            sp(ctrl, 'ForeColor', C_DANGER)
            sp(ctrl, 'FontBold',  True)
        elif 'prev' in n and n.startswith('lbl'):
            sp(ctrl, 'ForeColor', C_MUTED)
            sp(ctrl, 'FontSize',  8)
        elif n in ('lblvalidrate','lblrecvrate','lblhours','lblproductivity'):
            sp(ctrl, 'ForeColor', C_MUTED)
        elif n.startswith('lk') or n.startswith('lk2'):
            sp(ctrl, 'ForeColor', C_HEADER_TXT)
        else:
            sp(ctrl, 'ForeColor', C_NAVY)   # デフォルト: ダークネイビー

    elif ct == acButton:
        color = get_btn_color(name, form_name)
        sp(ctrl, 'BackColor',    color)
        sp(ctrl, 'ForeColor',    C_WHITE)   # ボタンは白文字（有色背景のため）
        sp(ctrl, 'BackStyle',    1)
        sp(ctrl, 'SpecialEffect',0)
        sp(ctrl, 'BorderStyle',  0)
        sp(ctrl, 'FontBold',     True)
        if form_name == 'F_Main':
            sp(ctrl, 'FontSize', 10)

    elif ct == acTextBox:
        sp(ctrl, 'BackColor',    C_WHITE)
        sp(ctrl, 'ForeColor',    0)          # 黒 (読みやすさ最優先)
        sp(ctrl, 'BorderStyle',  1)
        sp(ctrl, 'BorderColor',  C_BORDER)
        sp(ctrl, 'SpecialEffect',0)

    elif ct == acList:
        sp(ctrl, 'BackColor',    C_WHITE)
        sp(ctrl, 'ForeColor',    0)          # 黒
        sp(ctrl, 'BorderStyle',  1)
        sp(ctrl, 'BorderColor',  C_BORDER)
        sp(ctrl, 'SpecialEffect',0)
        sp(ctrl, 'ColumnHeads',  True)
        sp(ctrl, 'AlternateBackColor', C_ALT)

    elif ct == acCombo:
        sp(ctrl, 'BackColor',    C_WHITE)
        sp(ctrl, 'ForeColor',    0)          # 黒
        sp(ctrl, 'BorderStyle',  1)
        sp(ctrl, 'BorderColor',  C_BORDER)
        sp(ctrl, 'SpecialEffect',0)

    elif ct == acCheck:
        sp(ctrl, 'ForeColor', C_NAVY)

# ════════════════════════════════════════════════════════════════
# フォーム作成（デザイン付き）
# ════════════════════════════════════════════════════════════════
def make_form(app, name, caption, controls, vba,
              popup=False, modal=False, width=c(25)):
    # 既存削除
    try: app.DoCmd.Close(2, name)
    except: pass
    try:
        app.DoCmd.DeleteObject(2, name)
        time.sleep(0.5)
    except: pass

    frm = app.CreateForm()
    frm.Caption          = caption
    frm.RecordSelectors  = False
    frm.NavigationButtons= False
    frm.DividingLines    = False
    frm.ScrollBars       = 2
    frm.DefaultView      = 0
    frm.Width            = width
    if popup: frm.PopUp = True
    if modal: frm.Modal = True
    tmp = frm.Name

    # Detail セクション背景色
    try:
        frm.Section(acDetail).BackColor = C_BG
    except Exception:
        pass

    for cd in controls:
        ct   = cd['type']
        ctrl = app.CreateControl(tmp, ct, acDetail, '', '',
                                 cd['left'], cd['top'], cd['width'], cd['height'])
        cname = cd.get('name', '')
        if cname: ctrl.Name = cname
        if 'caption' in cd: ctrl.Caption = cd['caption']
        if ct != acCheck:
            ctrl.FontName = 'Yu Gothic UI'
            ctrl.FontSize = cd.get('fs', 9)
        if cd.get('bold'): ctrl.FontBold = True
        if ct == acButton: ctrl.OnClick = '[Event Procedure]'
        if ct == acList:
            ctrl.ColumnCount   = cd.get('cc', 2)
            if cd.get('cw'):   ctrl.ColumnWidths = cd['cw']
            ctrl.RowSourceType = 'Table/Query'
        if ct == acCombo:
            ctrl.ColumnCount   = 1
            ctrl.RowSourceType = 'Table/Query'
        # デザイン適用
        apply_design(ctrl, ct, cname, name)

    # HasModule → VBA注入
    frm.HasModule = True
    time.sleep(0.2)
    if vba:
        try:
            comp = app.VBE.VBProjects(1).VBComponents('Form_' + tmp)
            cm   = comp.CodeModule
            for i, line in enumerate(vba.strip().split('\n'), 1):
                cm.InsertLines(i, line)
        except Exception as e:
            print(f'    VBA注入エラー {name}: {e}')

    app.DoCmd.Save(2, tmp)
    app.DoCmd.Close(2, tmp)
    time.sleep(0.5)
    app.DoCmd.Rename(name, 2, tmp)
    time.sleep(0.5)
    print(f'  {name}: OK')

# ════════════════════════════════════════════════════════════════
# VBA: F_Main
# ════════════════════════════════════════════════════════════════
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

# ════════════════════════════════════════════════════════════════
# VBA: F_Members
# ════════════════════════════════════════════════════════════════
VBA_MEMBERS = """Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lstActive.RowSource   = "SELECT ID, member_name FROM T_MEMBERS WHERE active=True  ORDER BY member_name"
    lstInactive.RowSource = "SELECT ID, member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name"
    lstActive.Requery
    lstInactive.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub

Private Sub btnAdd_Click()
On Error GoTo EH
    Dim nm As String
    nm = Trim(Nz(txtNewName.Value, ""))
    If Len(nm) = 0 Then MsgBox "名前を入力してください", vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_MEMBERS (member_name, active) VALUES ('" & Replace(nm,"'","''") & "', True)", dbFailOnError
    txtNewName.Value = ""
    LoadData
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnAdd"
End Sub

Private Sub btnDeactivate_Click()
On Error GoTo EH
    If IsNull(lstActive.Value) Then MsgBox "選択してください", vbExclamation: Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=False WHERE ID=" & lstActive.Value, dbFailOnError
    LoadData
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnDeactivate"
End Sub

Private Sub btnActivate_Click()
On Error GoTo EH
    If IsNull(lstInactive.Value) Then MsgBox "選択してください", vbExclamation: Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=True WHERE ID=" & lstInactive.Value, dbFailOnError
    LoadData
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnActivate"
End Sub"""

# ════════════════════════════════════════════════════════════════
# VBA: F_Daily
# ════════════════════════════════════════════════════════════════
VBA_DAILY = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date): LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    Dim dFrom As Date, dTo As Date
    dFrom = DateSerial(mY, mM, 1)
    If mM = 12 Then dTo = DateSerial(mY + 1, 1, 1) Else dTo = DateSerial(mY, mM + 1, 1)
    Dim sql As String
    sql = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd'), R.member_name, R.calls," _
        & " R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.[note]" _
        & " FROM T_RECORDS AS R" _
        & " WHERE R.rec_date >= #" & Format(dFrom,"yyyy/mm/dd") & "#" _
        & " AND R.rec_date < #" & Format(dTo,"yyyy/mm/dd") & "#"
    If Nz(cboMember.Value,"") <> "" Then
        sql = sql & " AND R.member_name='" & Replace(cboMember.Value,"'","''") & "'"
    End If
    sql = sql & " ORDER BY R.rec_date DESC, R.member_name"
    lstRecords.RowSource = sql: lstRecords.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub

Private Sub btnPrev_Click()
    If mM = 1 Then mY = mY - 1: mM = 12 Else mM = mM - 1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM = 12 Then mY = mY + 1: mM = 1 Else mM = mM + 1
    LoadData
End Sub

Private Sub cboMember_AfterUpdate(): LoadData: End Sub

Private Sub btnAdd_Click()
    DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"ADD|" & mY & "|" & mM
    LoadData
End Sub

Private Sub btnEdit_Click()
    If IsNull(lstRecords.Value) Then MsgBox "行を選択してください", vbExclamation: Exit Sub
    DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"EDIT|" & lstRecords.Value
    LoadData
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstRecords.Value) Then MsgBox "行を選択してください", vbExclamation: Exit Sub
    If MsgBox("削除しますか？", vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "DELETE FROM T_RECORDS WHERE ID=" & lstRecords.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnDelete"
End Sub

Private Sub lstRecords_DblClick(Cancel As Integer): btnEdit_Click: End Sub"""

# ════════════════════════════════════════════════════════════════
# VBA: F_DailyEdit
# ════════════════════════════════════════════════════════════════
VBA_DAILY_EDIT = """Option Compare Database
Option Explicit

Private mMode As String
Private mID As Long

Private Sub Form_Open(Cancel As Integer)
On Error GoTo EH
    Dim args() As String
    args = Split(Nz(Me.OpenArgs,"ADD|0|0"),"|")
    mMode = args(0)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    If mMode = "EDIT" Then
        mID = CLng(args(1)): Me.Caption = "編集": LoadRecord
    Else
        mID = 0: Me.Caption = "新規登録"
        Dim bY As Integer, bM As Integer
        bY = CInt(args(1)): bM = CInt(args(2))
        txtRecDate.Value = Format(DateSerial(bY, bM, Day(Date)),"yyyy/mm/dd")
        txtWorkHours.Value = 8: chkWorkDay.Value = True
        Dim h As Integer
        For h = 10 To 18: Me("txtC" & h).Value = 0: Next h
        txtValid.Value=0: txtProspect.Value=0: txtDoc.Value=0
        txtFollow.Value=0: txtReceived.Value=0: txtReferral.Value=0
        txtNote.Value=""
    End If
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "Form_Open"
End Sub

Private Sub LoadRecord()
On Error GoTo EH
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_RECORDS WHERE ID=" & mID)
    If rs.EOF Then MsgBox "レコードが見つかりません", vbExclamation: rs.Close: Exit Sub
    txtRecDate.Value = Format(rs("rec_date"),"yyyy/mm/dd")
    cboMember.Value  = rs("member_name")
    Dim h As Integer
    For h = 10 To 18: Me("txtC" & h).Value = Nz(rs("calls_" & h),0): Next h
    txtValid.Value=Nz(rs("valid_count"),0): txtProspect.Value=Nz(rs("prospect"),0)
    txtDoc.Value=Nz(rs("doc"),0):           txtFollow.Value=Nz(rs("follow_up"),0)
    txtReceived.Value=Nz(rs("received"),0): txtReferral.Value=Nz(rs("referral"),0)
    txtWorkHours.Value=Nz(rs("work_hours"),8): chkWorkDay.Value=Nz(rs("work_day"),False)
    txtNote.Value=Nz(rs("note"),"")
    rs.Close
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadRecord"
End Sub

Private Sub btnSave_Click()
On Error GoTo EH
    If Nz(txtRecDate.Value,"")="" Then MsgBox "日付を入力してください",vbExclamation: Exit Sub
    If Nz(cboMember.Value,"")=""  Then MsgBox "担当者を選択してください",vbExclamation: Exit Sub
    Dim dt As Date: dt = CDate(txtRecDate.Value)
    Dim totalCalls As Long, h As Integer: totalCalls = 0
    For h = 10 To 18: totalCalls = totalCalls + CLng(Nz(Me("txtC" & h).Value,0)): Next h
    Dim mn As String: mn = Replace(cboMember.Value,"'","''")
    Dim nt As String: nt = Replace(Nz(txtNote.Value,""),"'","''")
    Dim wd As String: wd = IIf(Nz(chkWorkDay.Value,False),"True","False")
    Dim sql As String
    If mID = 0 Then
        sql = "INSERT INTO T_RECORDS (rec_date, member_name, calls" _
            & ", calls_10, calls_11, calls_12, calls_13, calls_14, calls_15, calls_16, calls_17, calls_18" _
            & ", valid_count, prospect, doc, follow_up, received, work_hours, [note], referral, work_day)" _
            & " VALUES (#" & Format(dt,"yyyy/mm/dd") & "#,'" & mn & "'," & totalCalls
        For h = 10 To 18: sql = sql & ", " & CLng(Nz(Me("txtC" & h).Value,0)): Next h
        sql = sql & "," & CLng(Nz(txtValid,0)) & "," & CLng(Nz(txtProspect,0)) _
            & "," & CLng(Nz(txtDoc,0)) & "," & CLng(Nz(txtFollow,0)) _
            & "," & CLng(Nz(txtReceived,0)) & "," & CDbl(Nz(txtWorkHours,8)) _
            & ",'" & nt & "'," & CLng(Nz(txtReferral,0)) & "," & wd & ")"
    Else
        sql = "UPDATE T_RECORDS SET rec_date=#" & Format(dt,"yyyy/mm/dd") & "#,member_name='" & mn & "',calls=" & cl
        For h = 10 To 18: sql = sql & ",calls_" & h & "=" & CLng(Nz(Me("txtC" & h).Value,0)): Next h
        sql = sql & ",valid_count=" & CLng(Nz(txtValid,0)) & ",prospect=" & CLng(Nz(txtProspect,0)) _
            & ",doc=" & CLng(Nz(txtDoc,0)) & ",follow_up=" & CLng(Nz(txtFollow,0)) _
            & ",received=" & CLng(Nz(txtReceived,0)) & ",work_hours=" & CDbl(Nz(txtWorkHours,8)) _
            & ",[note]='" & nt & "',referral=" & CLng(Nz(txtReferral,0)) _
            & ",work_day=" & wd & " WHERE ID=" & mID
    End If
    CurrentDb.Execute sql, dbFailOnError
    DoCmd.Close acForm, Me.Name
    Exit Sub
EH:
    MsgBox "保存エラー: " & Err.Description & vbCrLf & sql, vbCritical, "btnSave"
End Sub

Private Sub btnCancel_Click(): DoCmd.Close acForm, Me.Name: End Sub"""

# ════════════════════════════════════════════════════════════════
# VBA: F_Targets
# ════════════════════════════════════════════════════════════════
VBA_TARGETS = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY=Year(Date): mM=Month(Date)
    cboMember.RowSource="SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    lstTargets.RowSource = "SELECT T.ID,T.member_name,T.target_calls,T.target_valid," _
        & "T.target_prospect,T.target_received,T.target_referral,T.plan_days,T.work_hours_per_day" _
        & " FROM T_MEMBER_TARGETS AS T WHERE T.target_year=" & mY & " AND T.target_month=" & mM _
        & " ORDER BY T.member_name"
    lstTargets.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub

Private Sub btnPrev_Click()
    If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1
    LoadData
End Sub

Private Sub cboMember_AfterUpdate()
On Error GoTo EH
    If Nz(cboMember.Value,"")="" Then Exit Sub
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" _
        & Replace(cboMember.Value,"'","''") & "' AND target_year=" & mY & " AND target_month=" & mM)
    If Not rs.EOF Then
        txtPlanDays.Value=rs("plan_days"): txtHoursPerDay.Value=rs("work_hours_per_day")
        txtTgtCalls.Value=rs("target_calls"): txtTgtValid.Value=rs("target_valid")
        txtTgtProspect.Value=rs("target_prospect"): txtTgtReceived.Value=rs("target_received")
        txtTgtReferral.Value=rs("target_referral")
    Else
        txtPlanDays.Value="": txtHoursPerDay.Value=""
        txtTgtCalls.Value="": txtTgtValid.Value="": txtTgtProspect.Value=""
        txtTgtReceived.Value="": txtTgtReferral.Value=""
    End If
    rs.Close
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "cboMember_AfterUpdate"
End Sub

Private Sub btnLoadPrev_Click()
On Error GoTo EH
    If Nz(cboMember.Value,"")="" Then Exit Sub
    Dim py As Integer, pm As Integer
    py=mY: pm=mM: If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" _
        & Replace(cboMember.Value,"'","''") & "' AND target_year=" & py & " AND target_month=" & pm)
    If Not rs.EOF Then
        txtPlanDays.Value=rs("plan_days"): txtHoursPerDay.Value=rs("work_hours_per_day")
        txtTgtCalls.Value=rs("target_calls"): txtTgtValid.Value=rs("target_valid")
        txtTgtProspect.Value=rs("target_prospect"): txtTgtReceived.Value=rs("target_received")
        txtTgtReferral.Value=rs("target_referral")
    Else
        MsgBox "前月データなし", vbInformation
    End If
    rs.Close
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnLoadPrev"
End Sub

Private Sub btnSave_Click()
On Error GoTo EH
    If Nz(cboMember.Value,"")="" Then MsgBox "担当者を選択してください", vbExclamation: Exit Sub
    Dim nm As String: nm = Replace(cboMember.Value,"'","''")
    Dim n As Long: n = DCount("*","T_MEMBER_TARGETS","member_name='" & nm & "' AND target_year=" & mY & " AND target_month=" & mM)
    If n > 0 Then
        CurrentDb.Execute "UPDATE T_MEMBER_TARGETS SET plan_days=" & CLng(Nz(txtPlanDays,20)) _
            & ",work_hours_per_day=" & CDbl(Nz(txtHoursPerDay,8)) _
            & ",target_calls=" & CLng(Nz(txtTgtCalls,0)) & ",target_valid=" & CLng(Nz(txtTgtValid,0)) _
            & ",target_prospect=" & CLng(Nz(txtTgtProspect,0)) & ",target_received=" & CLng(Nz(txtTgtReceived,0)) _
            & ",target_referral=" & CLng(Nz(txtTgtReferral,0)) _
            & " WHERE member_name='" & nm & "' AND target_year=" & mY & " AND target_month=" & mM, dbFailOnError
    Else
        CurrentDb.Execute "INSERT INTO T_MEMBER_TARGETS (member_name,target_year,target_month,plan_days,work_hours_per_day,target_calls,target_valid,target_prospect,target_received,target_referral) VALUES ('" _
            & nm & "'," & mY & "," & mM & "," & CLng(Nz(txtPlanDays,20)) & "," & CDbl(Nz(txtHoursPerDay,8)) _
            & "," & CLng(Nz(txtTgtCalls,0)) & "," & CLng(Nz(txtTgtValid,0)) & "," & CLng(Nz(txtTgtProspect,0)) _
            & "," & CLng(Nz(txtTgtReceived,0)) & "," & CLng(Nz(txtTgtReferral,0)) & ")", dbFailOnError
    End If
    LoadData
    MsgBox "保存しました", vbInformation
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnSave"
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstTargets.Value) Then Exit Sub
    If MsgBox("削除しますか？",vbYesNo+vbQuestion)=vbYes Then
        CurrentDb.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & lstTargets.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnDelete"
End Sub"""

# ════════════════════════════════════════════════════════════════
# VBA: F_Referrals
# ════════════════════════════════════════════════════════════════
VBA_REFERRALS = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY=Year(Date): mM=Month(Date)
    cboMember.RowSource="SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月"
    txtRefDate.Value = Format(Date,"yyyy/mm/dd")
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY,mM,1),"yyyy/mm/dd")
    If mM = 12 Then dT = Format(DateSerial(mY+1,1,1),"yyyy/mm/dd") Else dT = Format(DateSerial(mY,mM+1,1),"yyyy/mm/dd")
    lstRefs.RowSource = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd'), R.member_name, R.ref_count" _
        & " FROM T_REFERRALS AS R" _
        & " WHERE R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & " ORDER BY R.rec_date DESC, R.member_name"
    lstRefs.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub

Private Sub btnPrev_Click()
    If mM=1 Then mY=mY-1: mM=12 Else mM=mM-1
    LoadData
End Sub

Private Sub btnNext_Click()
    If mM=12 Then mY=mY+1: mM=1 Else mM=mM+1
    LoadData
End Sub

Private Sub btnAdd_Click()
On Error GoTo EH
    If Nz(cboMember.Value,"")="" Or Nz(txtRefDate.Value,"")="" Then MsgBox "必須項目を入力してください", vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_REFERRALS (rec_date,member_name,ref_count) VALUES (#" _
        & Format(CDate(txtRefDate.Value),"yyyy/mm/dd") & "#,'" & Replace(cboMember.Value,"'","''") _
        & "'," & Nz(txtRefCount,0) & ")", dbFailOnError
    LoadData
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnAdd"
End Sub

Private Sub btnDelete_Click()
On Error GoTo EH
    If IsNull(lstRefs.Value) Then Exit Sub
    If MsgBox("削除しますか？",vbYesNo+vbQuestion)=vbYes Then
        CurrentDb.Execute "DELETE FROM T_REFERRALS WHERE ID=" & lstRefs.Value, dbFailOnError
        LoadData
    End If
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnDelete"
End Sub"""

# ════════════════════════════════════════════════════════════════
# VBA: F_Report（PDF出力実装済み・直接SQL方式）
# ════════════════════════════════════════════════════════════════
VBA_REPORT = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date): LoadReport
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
    Dim pyExe  As String
    Dim script As String
    Dim dbPath As String
    Dim cmd    As String
    pyExe  = "C:\\Users\\agentcode01\\AppData\\Local\\Programs\\Python\\Python312\\python.exe"
    script = CurrentProject.Path & "\\generate_pdf.py"
    dbPath = CurrentProject.FullName
    cmd = "cmd /c " & Chr(34) & pyExe & Chr(34) _
        & " " & Chr(34) & script & Chr(34) _
        & " --year " & mY & " --month " & mM _
        & " --db " & Chr(34) & dbPath & Chr(34)
    Shell cmd, vbHide
    MsgBox mY & "年" & mM & "月 PDF生成中..." & vbCrLf & "デスクトップに保存されます。", vbInformation, "PDF出力"
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "btnPDF"
End Sub

Private Sub LoadReport()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月 レポート"
    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then dT = Format(DateSerial(mY+1, 1, 1), "yyyy/mm/dd") Else dT = Format(DateSerial(mY, mM+1, 1), "yyyy/mm/dd")

    Dim rs As DAO.Recordset
    Dim tC As Long, tV As Long, tP As Long, tR As Long, tH As Double
    tC=0: tV=0: tP=0: tR=0: tH=0
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received), Sum(R.work_hours)" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#")
    If Not rs.EOF Then
        tC=Nz(rs.Fields(0).Value,0): tV=Nz(rs.Fields(1).Value,0)
        tP=Nz(rs.Fields(2).Value,0): tR=Nz(rs.Fields(3).Value,0): tH=Nz(rs.Fields(4).Value,0)
    End If: rs.Close

    Dim gC As Long, gV As Long, gP As Long, gR As Long
    gC=0: gV=0: gP=0: gR=0
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(target_calls), Sum(target_valid), Sum(target_prospect), Sum(target_received)" _
        & " FROM T_MEMBER_TARGETS WHERE target_year=" & mY & " AND target_month=" & mM)
    If Not rs.EOF Then
        gC=Nz(rs.Fields(0).Value,0): gV=Nz(rs.Fields(1).Value,0)
        gP=Nz(rs.Fields(2).Value,0): gR=Nz(rs.Fields(3).Value,0)
    End If: rs.Close

    Dim py As Integer, pm As Integer
    py=mY: pm=mM: If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Dim pF As String, pT As String
    pF = Format(DateSerial(py,pm,1),"yyyy/mm/dd")
    If pm=12 Then pT=Format(DateSerial(py+1,1,1),"yyyy/mm/dd") Else pT=Format(DateSerial(py,pm+1,1),"yyyy/mm/dd")
    Dim pC As Long, pV As Long, pP As Long, pR As Long
    pC=0: pV=0: pP=0: pR=0
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received)" _
        & " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & pF & "# AND R.rec_date < #" & pT & "#")
    If Not rs.EOF Then
        pC=Nz(rs.Fields(0).Value,0): pV=Nz(rs.Fields(1).Value,0)
        pP=Nz(rs.Fields(2).Value,0): pR=Nz(rs.Fields(3).Value,0)
    End If: rs.Close

    lblCalls.Caption = Format(tC,"#,##0") & " / " & Format(gC,"#,##0")
    lblValid.Caption = Format(tV,"#,##0") & " / " & Format(gV,"#,##0")
    lblProsp.Caption = Format(tP,"#,##0") & " / " & Format(gP,"#,##0")
    lblRecv.Caption  = Format(tR,"#,##0") & " / " & Format(gR,"#,##0")
    lblCallsPrev.Caption = "前月" & Format(pC,"#,##0") & " " & MakeArrow(tC,pC)
    lblValidPrev.Caption = "前月" & Format(pV,"#,##0") & " " & MakeArrow(tV,pV)
    lblProspPrev.Caption = "前月" & Format(pP,"#,##0") & " " & MakeArrow(tP,pP)
    lblRecvPrev.Caption  = "前月" & Format(pR,"#,##0") & " " & MakeArrow(tR,pR)
    If tC > 0 Then
        lblValidRate.Caption = Format(tV/tC*100,"0.0") & "%"
        lblRecvRate.Caption  = Format(tR/tC*100,"0.0") & "%"
    Else: lblValidRate.Caption="-": lblRecvRate.Caption="-"
    End If
    lblHours.Caption = Format(tH,"#,##0") & "h"
    If tH > 0 Then lblProductivity.Caption = Format(tP/tH,"0.000") Else lblProductivity.Caption="-"

    Dim al As String: al=""
    If gR > 0 Then
        If tR >= gR Then al=al & Chr(9675) & " 受注 目標達成!" & vbCrLf Else al=al & Chr(9651) & " 受注 残り" & (gR-tR) & "件  " & tR & "/" & gR & vbCrLf
    End If
    If gP > 0 Then
        If tP >= gP Then al=al & Chr(9675) & " 見込 目標達成!" Else al=al & Chr(9651) & " 見込 残り" & (gP-tP) & "件  " & tP & "/" & gP
    End If
    lblAlert.Caption = al

    Dim baseQ As String
    baseQ = " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
          & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
          & " GROUP BY R.member_name ORDER BY "
    lstRankRef.RowSource   = "SELECT R.member_name, Sum(R.referral)"  & baseQ & "Sum(R.referral)  DESC"
    lstRankRecv.RowSource  = "SELECT R.member_name, Sum(R.received)"  & baseQ & "Sum(R.received)  DESC"
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect)"  & baseQ & "Sum(R.prospect)  DESC"
    lstRankRef.Requery: lstRankRecv.Requery: lstRankProsp.Requery
    Exit Sub
EH:
    MsgBox "レポート読込エラー: " & Err.Description, vbCritical, "LoadReport"
End Sub

Private Function MakeArrow(cur As Long, prev As Long) As String
    If cur > prev Then MakeArrow = Chr(9650) & Format(cur-prev,"#,##0")
    ElseIf cur < prev Then MakeArrow = Chr(9660) & Format(prev-cur,"#,##0")
    Else MakeArrow = "-"
    End If
End Function"""

# ════════════════════════════════════════════════════════════════
# VBA: F_Ranking（直接SQL方式）
# ════════════════════════════════════════════════════════════════
VBA_RANKING = """Option Compare Database
Option Explicit

Private mY As Integer
Private mM As Integer

Private Sub Form_Open(Cancel As Integer)
    mY = Year(Date): mM = Month(Date): LoadData
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
    If mM = 12 Then dT = Format(DateSerial(mY+1, 1, 1), "yyyy/mm/dd") Else dT = Format(DateSerial(mY, mM+1, 1), "yyyy/mm/dd")
    Dim baseQ As String
    baseQ = " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
          & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
          & " GROUP BY R.member_name ORDER BY "
    lstRef.RowSource  = "SELECT R.member_name, Sum(R.referral)"  & baseQ & "Sum(R.referral)  DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received)"  & baseQ & "Sum(R.received)  DESC"
    lstProsp.RowSource= "SELECT R.member_name, Sum(R.prospect)"  & baseQ & "Sum(R.prospect)  DESC"
    lstRef.Requery: lstRecv.Requery: lstProsp.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub"""

# ════════════════════════════════════════════════════════════════
# メイン処理
# ════════════════════════════════════════════════════════════════
def main():
    print("=" * 65)
    print("SalesMgr フォームデザイン刷新（CreateControl方式）")
    print("  フラット × ダークブルー+ホワイト × 游ゴシック × ゼブラ")
    print("=" * 65)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE)
    time.sleep(2)

    print("\n[フォーム再作成] 全8フォーム")

    # ── F_Main ───────────────────────────────────────────────────
    btns_main = [("btnDaily","日次登録"),("btnTargets","目標設定"),("btnReferrals","送客登録"),
                 ("btnReport","レポート"),("btnRanking","ランキング"),("btnMembers","担当者管理")]
    ctrls_main = [L("lblTitle","SalesMgr 営業管理",c(1),c(0.2),c(13),c(1.2),16,True,C_TITLE)]
    for i,(bn,bc) in enumerate(btns_main):
        ctrls_main.append(B(bn,bc, c(0.5)+(i%2)*c(7), c(1.8)+(i//2)*c(1.5), c(6),c(1.2),11))
    make_form(app,"F_Main","SalesMgr 営業管理",ctrls_main,VBA_MAIN,width=c(15))

    # ── F_Members ────────────────────────────────────────────────
    make_form(app,"F_Members","担当者管理",[
        L("lblTitle","担当者管理",c(0.5),c(0.3),c(6),c(0.9),13,True,C_TITLE),
        L("lblAdd","新規:",c(0.5),c(1.5),c(1.5),c(0.6)),
        T("txtNewName",c(2),c(1.5),c(4),c(0.6)),
        B("btnAdd","追加",c(6.5),c(1.5),c(2),c(0.6)),
        L("lblA","有効",c(0.5),c(2.8),c(4),c(0.6),9,True,C_NAVY),
        LB("lstActive",c(0.5),c(3.5),c(5),c(5),2,"0cm;4cm"),
        B("btnDeactivate","無効▶",c(6),c(5),c(3),c(0.7)),
        L("lblI","無効",c(10),c(2.8),c(4),c(0.6),9,True,C_NAVY),
        LB("lstInactive",c(10),c(3.5),c(5),c(5),2,"0cm;4cm"),
        B("btnActivate","◀有効",c(6),c(6.5),c(3),c(0.7)),
    ],VBA_MEMBERS)

    # ── F_DailyEdit（ポップアップ）────────────────────────────────
    ctrls_de = [
        L("lbl1","日付:",c(0.3),c(0.5),c(1.5),c(0.6)),  T("txtRecDate",c(2),c(0.5),c(3),c(0.6)),
        L("lbl2","担当者:",c(0.3),c(1.3),c(1.5),c(0.6)), C("cboMember",c(2),c(1.3),c(3),c(0.6)),
        L("lbl3","稼働h:",c(0.3),c(2.1),c(1.5),c(0.6)),  T("txtWorkHours",c(2),c(2.1),c(1.5),c(0.6)),
        L("lbl3b","出勤:",c(4),c(2.1),c(1),c(0.6)),       CK("chkWorkDay",c(5),c(2.1),c(0.5),c(0.5)),
        L("lblHr","時間帯別架電:",c(0.3),c(3),c(3),c(0.6),9,True,C_NAVY),
    ]
    y = c(3.7)
    for h in range(10,19):
        ctrls_de.append(L(f"lH{h}",f"{h}時:",c(0.3),y,c(1.2),c(0.5)))
        ctrls_de.append(T(f"txtC{h}",c(1.5),y,c(1.5),c(0.5)))
        y += c(0.6)
    y += c(0.3)
    for lb,nm in [("有効:","txtValid"),("見込:","txtProspect"),("資料:","txtDoc"),
                  ("追客:","txtFollow"),("受注:","txtReceived"),("送客:","txtReferral")]:
        ctrls_de.append(L(f"l{nm}",lb,c(0.3),y,c(1.5),c(0.5)))
        ctrls_de.append(T(nm,c(2),y,c(1.5),c(0.5)))
        y += c(0.6)
    ctrls_de += [
        L("lNote","備考:",c(0.3),y,c(1.5),c(0.5)),
        T("txtNote",c(2),y,c(4),c(0.8)),
        B("btnSave","保存",c(0.5),y+c(1.2),c(2.5),c(0.7)),
        B("btnCancel","キャンセル",c(3.5),y+c(1.2),c(2.5),c(0.7)),
    ]
    make_form(app,"F_DailyEdit","日次レコード",ctrls_de,VBA_DAILY_EDIT,popup=True,modal=True,width=c(7))

    # ── F_Daily ──────────────────────────────────────────────────
    make_form(app,"F_Daily","日次一覧",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),
        L("lblMonth","",c(2),c(0.3),c(5),c(0.7),12,True,C_TITLE),
        B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lblF","担当者:",c(10),c(0.3),c(2),c(0.7)),
        C("cboMember",c(12),c(0.3),c(3.5),c(0.7)),
        B("btnAdd","新規登録",c(0.5),c(1.3),c(2.5),c(0.7)),
        B("btnEdit","編集",c(3.5),c(1.3),c(2),c(0.7)),
        B("btnDelete","削除",c(6),c(1.3),c(2),c(0.7)),
        LB("lstRecords",c(0.5),c(2.5),c(22),c(12),11,
           "0cm;2.5cm;2.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;3cm"),
    ],VBA_DAILY)

    # ── F_Targets ────────────────────────────────────────────────
    ctrls_tgt = [
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),
        L("lblMonth","",c(2),c(0.3),c(5),c(0.7),12,True,C_TITLE),
        B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lm","担当者:",c(0.5),c(1.5),c(2),c(0.6)),
        C("cboMember",c(2.5),c(1.5),c(3.5),c(0.6)),
        B("btnLoadPrev","前月コピー",c(6.5),c(1.5),c(2.5),c(0.6)),
    ]
    y = c(2.4)
    for lb,nm in [("稼働日数:","txtPlanDays"),("日稼働h:","txtHoursPerDay"),
                  ("架電目標:","txtTgtCalls"),("有効目標:","txtTgtValid"),
                  ("見込目標:","txtTgtProspect"),("受注目標:","txtTgtReceived"),
                  ("送客目標:","txtTgtReferral")]:
        ctrls_tgt.append(L(f"l{nm}",lb,c(0.5),y,c(2),c(0.6)))
        ctrls_tgt.append(T(nm,c(2.5),y,c(2.5),c(0.6)))
        y += c(0.7)
    ctrls_tgt += [
        B("btnSave","保存",c(0.5),y+c(0.2),c(2.5),c(0.7)),
        B("btnDelete","削除",c(3.5),y+c(0.2),c(2.5),c(0.7)),
        LB("lstTargets",c(0.5),y+c(1.3),c(20),c(4),9,"0cm;2.5cm;2cm;2cm;2cm;2cm;2cm;2cm;2cm"),
    ]
    make_form(app,"F_Targets","目標設定",ctrls_tgt,VBA_TARGETS)

    # ── F_Referrals ──────────────────────────────────────────────
    make_form(app,"F_Referrals","送客登録",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),
        L("lblMonth","",c(2),c(0.3),c(5),c(0.7),12,True,C_TITLE),
        B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("ld","日付:",c(0.5),c(1.5),c(1.2),c(0.6)),
        T("txtRefDate",c(1.7),c(1.5),c(2.5),c(0.6)),
        L("lm","担当者:",c(4.5),c(1.5),c(1.5),c(0.6)),
        C("cboMember",c(6),c(1.5),c(3),c(0.6)),
        L("lc","件数:",c(9.5),c(1.5),c(1),c(0.6)),
        T("txtRefCount",c(10.5),c(1.5),c(1.5),c(0.6)),
        B("btnAdd","追加",c(12.5),c(1.5),c(2),c(0.6)),
        B("btnDelete","削除",c(0.5),c(2.5),c(2),c(0.6)),
        LB("lstRefs",c(0.5),c(3.3),c(15),c(7),4,"0cm;3cm;3cm;2cm"),
    ],VBA_REFERRALS)

    # ── F_Report ─────────────────────────────────────────────────
    ctrls_rpt = [
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),
        L("lblMonth","",c(2),c(0.3),c(8),c(0.7),12,True,C_TITLE),
        B("btnNext","▶",c(10.5),c(0.3),c(1.2),c(0.7)),
        B("btnPDF","PDF出力",c(12.5),c(0.3),c(3),c(0.7)),
        L("lblAlert","",c(0.5),c(1.3),c(20),c(1),9,False,C_DANGER),
    ]
    y = c(2.6)
    for i,(title,nv,ns) in enumerate([
        ("架電件数","lblCalls","lblCallsPrev"),
        ("有効件数","lblValid","lblValidPrev"),
        ("見込件数","lblProsp","lblProspPrev"),
        ("受注件数","lblRecv","lblRecvPrev")]):
        x = c(0.5)+i*c(5)
        ctrls_rpt.append(L(f"lK{i}",title,x,y,c(4.5),c(0.4),7,False,C_MUTED))
        ctrls_rpt.append(L(nv,"-",x,y+c(0.4),c(4.5),c(0.7),11,True))
        ctrls_rpt.append(L(ns,"",x,y+c(1.1),c(4.5),c(0.4),7,False,C_MUTED))
    y += c(1.8)
    for i,(title,nv) in enumerate([
        ("有効率","lblValidRate"),("受注率","lblRecvRate"),
        ("稼働時間","lblHours"),("生産性","lblProductivity")]):
        x = c(0.5)+i*c(5)
        ctrls_rpt.append(L(f"lK2{i}",title,x,y,c(4.5),c(0.4),7,False,C_MUTED))
        ctrls_rpt.append(L(nv,"-",x,y+c(0.4),c(4.5),c(0.7),11,True))
    y += c(1.5)
    ctrls_rpt += [
        L("lR1","送客ランキング",c(0.5),y,c(6),c(0.6),9,True,C_NAVY),
        L("lR2","受注ランキング",c(7),y,c(6),c(0.6),9,True,C_NAVY),
        L("lR3","見込ランキング",c(13.5),y,c(6),c(0.6),9,True,C_NAVY),
        LB("lstRankRef",c(0.5),y+c(0.7),c(6),c(5),2,"3cm;2cm"),
        LB("lstRankRecv",c(7),y+c(0.7),c(6),c(5),2,"3cm;2cm"),
        LB("lstRankProsp",c(13.5),y+c(0.7),c(6),c(5),2,"3cm;2cm"),
    ]
    make_form(app,"F_Report","レポート",ctrls_rpt,VBA_REPORT)

    # ── F_Ranking ────────────────────────────────────────────────
    make_form(app,"F_Ranking","ランキング",[
        B("btnPrev","◀",c(0.5),c(0.3),c(1.2),c(0.7)),
        L("lblMonth","",c(2),c(0.3),c(5),c(0.7),12,True,C_TITLE),
        B("btnNext","▶",c(7.5),c(0.3),c(1.2),c(0.7)),
        L("lR1","送客",c(0.5),c(1.3),c(6),c(0.6),9,True,C_NAVY),
        L("lR2","受注",c(7),c(1.3),c(6),c(0.6),9,True,C_NAVY),
        L("lR3","見込",c(13.5),c(1.3),c(6),c(0.6),9,True,C_NAVY),
        LB("lstRef",c(0.5),c(2),c(6),c(7),2,"3cm;2cm"),
        LB("lstRecv",c(7),c(2),c(6),c(7),2,"3cm;2cm"),
        LB("lstProsp",c(13.5),c(2),c(6),c(7),2,"3cm;2cm"),
    ],VBA_RANKING)

    # ── スタートアップ設定 ──────────────────────────────────────
    print("\n[スタートアップ設定]")
    try:
        db    = app.CurrentDb()
        props = db.Properties
        for pn, pt, pv in [
            ("StartUpForm",         10, "F_Main"),
            ("AppTitle",            10, "SalesMgr 営業管理"),
            ("StartUpShowDBWindow",  1, False),
        ]:
            try:
                props(pn).Value = pv
            except Exception:
                prop = db.CreateProperty(pn, pt, pv)
                props.Append(prop)
        print("  StartUpForm=F_Main, AppTitle=SalesMgr 営業管理, ShowDBWindow=False")
    except Exception as e:
        print(f"  スタートアップ設定エラー: {e}")

    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)

    print("\n" + "=" * 65)
    print("デザイン刷新完了!")
    print("  次のステップ: python full_check.py でテスト実行")
    print("=" * 65)


if __name__ == "__main__":
    main()
