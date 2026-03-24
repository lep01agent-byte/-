# -*- coding: utf-8 -*-
"""
rename_and_browser.py  --  SalesMgr_FE.accdb 対象

作業1: 12クエリを日本語名に変更 + 全フォームVBAの旧名参照を更新
作業2: F_QueryBrowser フォーム作成 + F_Main にボタン追加
"""
import os, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")

CM = 567
def c(n): return int(CM * n)
acLabel=100; acButton=104; acTextBox=109; acCombo=111; acDetail=0

acQuery = 1
acForm  = 2

# ── クエリ名マッピング ────────────────────────────────────────
RENAME_MAP = {
    "Q_ActiveMembers":        "Q_アクティブメンバー一覧",
    "Q_Hourly_By_Member":     "Q_メンバー別時間別実績",
    "Q_Member_12Month":       "Q_メンバー12ヶ月推移",
    "Q_Member_Monthly_Sum":   "Q_メンバー月次集計",
    "Q_Rank_Productivity":    "Q_生産性ランキング",
    "Q_Rank_Prospect":        "Q_見込みランキング",
    "Q_Rank_Received":        "Q_受電ランキング",
    "Q_Rank_Referral":        "Q_紹介ランキング",
    "Q_RefTrend_Monthly":     "Q_紹介トレンド月次",
    "Q_Team_Monthly_Sum":     "Q_チーム月次集計",
    "Q_Team_Monthly_Targets": "Q_チーム月次目標",
    "Q_Trend_Monthly":        "Q_月次トレンド",
}

# ── F_QueryBrowser VBA ───────────────────────────────────────
VBA_QB = '''Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
On Error GoTo EH
    cboYear.RowSourceType  = "Value List"
    Dim ys As String, y As Integer
    For y = 2020 To 2035: ys = ys & y & ";": Next y
    cboYear.RowSource = Left(ys, Len(ys) - 1)
    cboYear.Value = Year(Date)

    cboMonth.RowSourceType = "Value List"
    cboMonth.RowSource = "1;2;3;4;5;6;7;8;9;10;11;12"
    cboMonth.Value = Month(Date)

    cboQuery.RowSourceType = "Value List"
    cboQuery.RowSource = "Q_アクティブメンバー一覧;Q_チーム月次集計;Q_チーム月次目標;" _
        & "Q_メンバー月次集計;Q_受電ランキング;Q_見込みランキング;Q_紹介ランキング;" _
        & "Q_生産性ランキング;Q_メンバー別時間別実績;Q_月次トレンド;" _
        & "Q_メンバー12ヶ月推移;Q_紹介トレンド月次"
    cboQuery.Value = "Q_チーム月次集計"
    Exit Sub
EH: MsgBox Err.Description, vbCritical, "Form_Open"
End Sub

Private Sub btnShow_Click()
On Error GoTo EH
    Dim qName As String: qName = Nz(cboQuery.Value, "")
    If qName = "" Then MsgBox "クエリを選択してください", vbExclamation: Exit Sub

    Dim yr As Long: yr = CLng(Nz(cboYear.Value, Year(Date)))
    Dim mo As Long: mo = CLng(Nz(cboMonth.Value, Month(Date)))

    ' 当月の開始/終了
    Dim dF As Date, dT As Date
    dF = DateSerial(yr, mo, 1)
    If mo = 12 Then dT = DateSerial(yr + 1, 1, 1) Else dT = DateSerial(yr, mo + 1, 1)

    ' 12ヶ月前の開始（トレンド用）
    Dim y12 As Integer, m12 As Integer: y12 = yr: m12 = mo
    Dim i As Integer
    For i = 1 To 11
        m12 = m12 - 1
        If m12 = 0 Then m12 = 12: y12 = y12 - 1
    Next i
    Dim d12S As Date: d12S = DateSerial(y12, m12, 1)

    ' クエリSQLを取得してPARAMETERS行を除去
    Dim sql As String: sql = Trim(CurrentDb.QueryDefs(qName).SQL)
    If UCase(Left(sql, 10)) = "PARAMETERS" Then
        Dim sp As Long: sp = InStr(sql, ";")
        If sp > 0 Then sql = Trim(Mid(sql, sp + 1))
    End If

    ' パラメータを値で置換
    Dim sf As String: sf = "yyyy/mm/dd"
    sql = Replace(sql, "[prmDateFrom]",   "#" & Format(dF,  sf) & "#")
    sql = Replace(sql, "[prmDateTo]",     "#" & Format(dT,  sf) & "#")
    sql = Replace(sql, "[prmYear]",       CStr(yr))
    sql = Replace(sql, "[prmMonth]",      CStr(mo))
    sql = Replace(sql, "[prmTrendStart]", "#" & Format(d12S, sf) & "#")
    sql = Replace(sql, "[prmTrendEnd]",   "#" & Format(dT,   sf) & "#")
    sql = Replace(sql, "[prm12Start]",    "#" & Format(d12S, sf) & "#")
    sql = Replace(sql, "[prm12End]",      "#" & Format(dT,   sf) & "#")

    ' 一時クエリを作成して開く
    Const TMP As String = "Q_TmpBrowserView"
    On Error Resume Next: CurrentDb.QueryDefs.Delete TMP: On Error GoTo EH
    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.CreateQueryDef(TMP, sql)
    qd.Close: Set qd = Nothing

    DoCmd.OpenQuery TMP, acViewNormal, acReadOnly
    lblStatus.Caption = qName & "  (" & yr & "年" & mo & "月)"
    Exit Sub
EH: MsgBox Err.Description, vbCritical, "btnShow"
End Sub

Private Sub btnClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name
End Sub'''


def inject_vba(app, form_name, code):
    """フォームのVBAモジュールを書き換える"""
    try:
        proj_idx = 1  # FE のプロジェクト
        comp = app.VBE.VBProjects(proj_idx).VBComponents("Form_" + form_name)
        cm   = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        lines = code.strip().replace('\r\n', '\n').split('\n')
        for i, line in enumerate(lines, 1):
            cm.InsertLines(i, line)
        return True
    except Exception as e:
        print(f"    VBA注入失敗 ({form_name}): {e}")
        return False


def make_ctrl(app, form_tmp, ct, name=None, cap=None, l=0, t=0, w=0, h=0, fs=9, bold=False):
    ctrl = app.CreateControl(form_tmp, ct, acDetail, '', '', l, t, w, h)
    if name: ctrl.Name    = name
    if cap:  ctrl.Caption = cap
    if ct   != 106:  # not checkbox
        ctrl.FontName = 'Yu Gothic UI'
        ctrl.FontSize = fs
    if bold: ctrl.FontBold = True
    if ct == acButton:
        ctrl.OnClick = '[Event Procedure]'
    if ct == acCombo:
        ctrl.ColumnCount   = 1
        ctrl.RowSourceType = 'Value List'
    return ctrl


def main():
    print('=' * 65)
    print('SalesMgr クエリ日本語名変更 + クエリブラウザ追加')
    print(f'対象: {FE}')
    print('=' * 65)

    app = win32com.client.DispatchEx('Access.Application')
    app.Visible    = False
    app.UserControl= False
    app.OpenCurrentDatabase(FE)
    time.sleep(2)

    # ════════════════════════════════════════════════════════
    # Phase 1: クエリ名変更
    # ════════════════════════════════════════════════════════
    print('\n[Phase 1] クエリ名変更')
    renamed_ok = {}
    for old, new in RENAME_MAP.items():
        try:
            app.DoCmd.Rename(new, acQuery, old)
            print(f'  OK  {old}')
            print(f'      → {new}')
            renamed_ok[old] = new
        except Exception as e:
            print(f'  NG  {old}: {e}')

    # ════════════════════════════════════════════════════════
    # Phase 2: フォームVBAの旧クエリ名参照を更新
    # ════════════════════════════════════════════════════════
    print('\n[Phase 2] フォームVBA旧名参照チェック・更新')
    FORMS = ['F_Main','F_Members','F_Daily','F_DailyEdit',
             'F_Targets','F_Referrals','F_Report','F_Ranking']

    for frm in FORMS:
        try:
            comp = app.VBE.VBProjects(1).VBComponents('Form_' + frm)
            cm   = comp.CodeModule
            if cm.CountOfLines == 0:
                print(f'  {frm}: コードなし')
                continue
            code = cm.Lines(1, cm.CountOfLines)
            hits = [o for o in renamed_ok if o in code]
            if not hits:
                print(f'  {frm}: 旧名参照なし ✓')
                continue
            # 旧名を新名に置換してモジュール再書き込み
            new_code = code
            for old in hits:
                new_code = new_code.replace(old, renamed_ok[old])
            cm.DeleteLines(1, cm.CountOfLines)
            lines = new_code.replace('\r\n', '\n').split('\n')
            for i, line in enumerate(lines, 1):
                cm.InsertLines(i, line)
            print(f'  {frm}: 更新 ({", ".join(hits)})')
        except Exception as e:
            print(f'  {frm}: エラー {e}')

    # ════════════════════════════════════════════════════════
    # Phase 3: F_Main に「クエリブラウザ」ボタンを追加
    # ════════════════════════════════════════════════════════
    print('\n[Phase 3] F_Main にクエリブラウザボタン追加')
    try:
        app.DoCmd.OpenForm('F_Main', 1)   # Design view
        time.sleep(0.8)

        # 既存ボタン6個の次の行 (i=6相当: row=3, col=both)
        ctrl = app.CreateControl(
            'F_Main', acButton, acDetail, '', '',
            c(0.5), c(6.5), c(13), c(1.2)
        )
        ctrl.Name     = 'btnQueryBrowser'
        ctrl.Caption  = 'クエリブラウザ'
        ctrl.FontName = 'Yu Gothic UI'
        ctrl.FontSize = 11
        ctrl.OnClick  = '[Event Procedure]'

        # F_Main VBAに新Subを追記
        comp = app.VBE.VBProjects(1).VBComponents('Form_F_Main')
        cm   = comp.CodeModule
        n    = cm.CountOfLines
        new_sub = 'Private Sub btnQueryBrowser_Click(): DoCmd.OpenForm "F_QueryBrowser": End Sub'
        cm.InsertLines(n + 1, new_sub)

        app.DoCmd.Save(acForm, 'F_Main')
        app.DoCmd.Close(acForm, 'F_Main')
        time.sleep(0.5)
        print('  F_Main: ボタン追加 + VBA更新完了')
    except Exception as e:
        print(f'  F_Main更新失敗: {e}')

    # ════════════════════════════════════════════════════════
    # Phase 4: F_QueryBrowser フォーム作成
    # ════════════════════════════════════════════════════════
    print('\n[Phase 4] F_QueryBrowser 作成')
    try:
        app.DoCmd.Close(acForm, 'F_QueryBrowser')
    except: pass
    try:
        app.DoCmd.DeleteObject(acForm, 'F_QueryBrowser')
        time.sleep(0.3)
    except: pass

    frm = app.CreateForm()
    frm.Caption          = 'クエリブラウザ'
    frm.RecordSelectors  = False
    frm.NavigationButtons= False
    frm.DividingLines    = False
    frm.ScrollBars       = 0
    frm.DefaultView      = 0
    frm.Width            = c(20)
    tmp = frm.Name

    # コントロール配置
    make_ctrl(app, tmp, acLabel,  name='lblTitle',  cap='クエリブラウザ',
              l=c(0.5), t=c(0.3), w=c(15), h=c(0.8), fs=14, bold=True)

    make_ctrl(app, tmp, acLabel,  name='lY',  cap='年:',
              l=c(0.5), t=c(1.5), w=c(0.8), h=c(0.6))
    make_ctrl(app, tmp, acCombo,  name='cboYear',
              l=c(1.4), t=c(1.5), w=c(2.2), h=c(0.6))

    make_ctrl(app, tmp, acLabel,  name='lM',  cap='月:',
              l=c(4.0), t=c(1.5), w=c(0.8), h=c(0.6))
    make_ctrl(app, tmp, acCombo,  name='cboMonth',
              l=c(4.9), t=c(1.5), w=c(1.8), h=c(0.6))

    make_ctrl(app, tmp, acLabel,  name='lQ',  cap='クエリ:',
              l=c(0.5), t=c(2.5), w=c(1.6), h=c(0.6))
    make_ctrl(app, tmp, acCombo,  name='cboQuery',
              l=c(2.2), t=c(2.5), w=c(14), h=c(0.6))

    make_ctrl(app, tmp, acButton, name='btnShow',  cap='表示',
              l=c(0.5), t=c(3.5), w=c(3.5), h=c(0.7), fs=10)
    make_ctrl(app, tmp, acButton, name='btnClose', cap='閉じる',
              l=c(4.5), t=c(3.5), w=c(2.5), h=c(0.7))

    make_ctrl(app, tmp, acLabel,  name='lblStatus', cap='',
              l=c(0.5), t=c(4.5), w=c(18), h=c(0.6), fs=8)

    # VBA注入
    frm.HasModule = True
    time.sleep(0.2)
    ok = inject_vba(app, tmp, VBA_QB)

    app.DoCmd.Save(acForm, tmp)
    app.DoCmd.Close(acForm, tmp)
    time.sleep(0.3)
    app.DoCmd.Rename('F_QueryBrowser', acForm, tmp)
    time.sleep(0.3)
    print(f'  F_QueryBrowser 作成完了  (VBA: {"OK" if ok else "要確認"})')

    # ════════════════════════════════════════════════════════
    # 保存・終了
    # ════════════════════════════════════════════════════════
    app.CloseCurrentDatabase()
    app.Quit()
    del app
    time.sleep(1)

    print('\n' + '=' * 65)
    print('完了！')
    print(f'  クエリ {len(renamed_ok)}/12 件を日本語名に変更')
    print('  F_Main: クエリブラウザボタン追加')
    print('  F_QueryBrowser: 新規作成（年月指定 + クエリ選択 + 表示）')
    print('=' * 65)


if __name__ == '__main__':
    main()
