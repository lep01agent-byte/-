# -*- coding: utf-8 -*-
"""F_Report と F_Ranking のVBAを直接SQL方式に修正"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

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
    MsgBox "PDF出力は今後対応予定です", vbInformation
End Sub

Private Sub LoadReport()
On Error GoTo EH
    lblMonth.Caption = mY & "年" & mM & "月 レポート"

    Dim dF As String, dT As String
    dF = Format(DateSerial(mY, mM, 1), "yyyy/mm/dd")
    If mM = 12 Then
        dT = Format(DateSerial(mY + 1, 1, 1), "yyyy/mm/dd")
    Else
        dT = Format(DateSerial(mY, mM + 1, 1), "yyyy/mm/dd")
    End If

    ' --- チーム実績（直接SQL） ---
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

    ' --- チーム目標（直接SQL） ---
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

    ' --- 前月実績 ---
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

    ' --- KPI表示 ---
    lblCalls.Caption = Format(tC, "#,##0") & " / " & Format(gC, "#,##0")
    lblValid.Caption = Format(tV, "#,##0") & " / " & Format(gV, "#,##0")
    lblProsp.Caption = Format(tP, "#,##0") & " / " & Format(gP, "#,##0")
    lblRecv.Caption = Format(tR, "#,##0") & " / " & Format(gR, "#,##0")

    lblCallsPrev.Caption = "前月" & Format(pC, "#,##0") & " " & MakeArrow(tC, pC)
    lblValidPrev.Caption = "前月" & Format(pV, "#,##0") & " " & MakeArrow(tV, pV)
    lblProspPrev.Caption = "前月" & Format(pP, "#,##0") & " " & MakeArrow(tP, pP)
    lblRecvPrev.Caption = "前月" & Format(pR, "#,##0") & " " & MakeArrow(tR, pR)

    If tC > 0 Then
        lblValidRate.Caption = Format(tV / tC * 100, "0.0") & "%"
        lblRecvRate.Caption = Format(tR / tC * 100, "0.0") & "%"
    Else
        lblValidRate.Caption = "-"
        lblRecvRate.Caption = "-"
    End If
    lblHours.Caption = Format(tH, "#,##0") & "h"
    If tH > 0 Then
        lblProductivity.Caption = Format(tP / tH, "0.000")
    Else
        lblProductivity.Caption = "-"
    End If

    ' --- アラート ---
    Dim al As String
    al = ""
    If gR > 0 Then
        If tR >= gR Then
            al = al & Chr(9675) & " 受注 目標達成！" & vbCrLf
        Else
            al = al & Chr(9651) & " 受注 残り" & (gR - tR) & "件 " & tR & "/" & gR & vbCrLf
        End If
    End If
    If gP > 0 Then
        If tP >= gP Then
            al = al & Chr(9675) & " 見込 目標達成！"
        Else
            al = al & Chr(9651) & " 見込 残り" & (gP - tP) & "件 " & tP & "/" & gP
        End If
    End If
    lblAlert.Caption = al

    ' --- ランキング（直接SQL） ---
    Dim baseQ As String
    baseQ = " FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name" _
        & " WHERE M.active=True AND R.rec_date >= #" & dF & "# AND R.rec_date < #" & dT & "#" _
        & " GROUP BY R.member_name ORDER BY "

    lstRankRef.RowSource = "SELECT R.member_name, Sum(R.referral)" & baseQ & "Sum(R.referral) DESC"
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received)" & baseQ & "Sum(R.received) DESC"
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect)" & baseQ & "Sum(R.prospect) DESC"
    lstRankRef.Requery
    lstRankRecv.Requery
    lstRankProsp.Requery
    Exit Sub
EH:
    MsgBox "レポート読込エラー: " & Err.Description, vbCritical, "LoadReport"
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

VBA_RANKING = """Option Compare Database
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
    lblMonth.Caption = mY & "年" & mM & "月"

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

    lstRef.RowSource = "SELECT R.member_name, Sum(R.referral)" & baseQ & "Sum(R.referral) DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received)" & baseQ & "Sum(R.received) DESC"
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect)" & baseQ & "Sum(R.prospect) DESC"
    lstRef.Requery
    lstRecv.Requery
    lstProsp.Requery
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "LoadData"
End Sub"""


def main():
    print("F_Report + F_Ranking VBA修正（直接SQL方式）")
    print("=" * 50)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE)
    time.sleep(2)

    for form_name, code in [("F_Report", VBA_REPORT), ("F_Ranking", VBA_RANKING)]:
        try:
            app.DoCmd.OpenForm(form_name, 1)
            time.sleep(0.4)
            frm = app.Forms(form_name)
            frm.HasModule = True
            time.sleep(0.2)
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
            print(f"  {form_name}: {len(lines)} lines OK")
        except Exception as e:
            print(f"  {form_name}: ERROR {e}")
            try: app.DoCmd.Close(2, form_name)
            except: pass

    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)
    print("Done!")


if __name__ == "__main__":
    main()
