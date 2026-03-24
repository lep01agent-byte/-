# -*- coding: utf-8 -*-
"""
SalesMgr Access フォーム・VBA・レポート 一括作成スクリプト
Application.LoadFromText で .txt 形式のフォーム定義をインポートする。
Access COM は使うが、ダイアログなしで完了する。
"""
import os, sys, time, tempfile

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BE_PATH = os.path.join(DESKTOP, "SalesMgr_BE.accdb")

# ============================================================
# フォーム定義テンプレート生成
# ============================================================

def _header(form_name, caption, width=14000, has_module=True, popup=False, modal=False):
    """Access フォーム定義のヘッダー部分"""
    lines = []
    lines.append(f'Version =21')
    lines.append(f'VersionRequired =20')
    lines.append(f'Begin Form')
    lines.append(f'    RecordSelectors = 0')
    lines.append(f'    NavigationButtons = 0')
    lines.append(f'    DividingLines = 0')
    lines.append(f'    AllowDesignChanges = 1')
    lines.append(f'    DefaultView =0')
    lines.append(f'    ScrollBars =2')
    if popup:
        lines.append(f'    PopUp = -1')
    if modal:
        lines.append(f'    Modal = -1')
    lines.append(f'    Width ={width}')
    lines.append(f'    Caption ="{caption}"')
    if has_module:
        lines.append(f'    HasModule =1')
    lines.append(f'    Begin')
    lines.append(f'        Begin')  # Detail section
    lines.append(f'            Height =10000')
    return '\n'.join(lines)


def _label(name, caption, left, top, width, height, font_size=9, bold=False, fore_color=0):
    lines = []
    lines.append(f'            Begin Label')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'                Caption ="{caption}"')
    lines.append(f'                FontName ="Yu Gothic UI"')
    lines.append(f'                FontSize ={font_size}')
    if bold:
        lines.append(f'                FontBold =-1')
    if fore_color != 0:
        lines.append(f'                ForeColor ={fore_color}')
    lines.append(f'            End')
    return '\n'.join(lines)


def _button(name, caption, left, top, width, height, font_size=9, on_click=""):
    lines = []
    lines.append(f'            Begin CommandButton')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'                Caption ="{caption}"')
    lines.append(f'                FontName ="Yu Gothic UI"')
    lines.append(f'                FontSize ={font_size}')
    if on_click:
        lines.append(f'                OnClick ="[Event Procedure]"')
    lines.append(f'            End')
    return '\n'.join(lines)


def _textbox(name, left, top, width, height, default_val=None):
    lines = []
    lines.append(f'            Begin TextBox')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'                FontName ="Yu Gothic UI"')
    lines.append(f'                FontSize =9')
    if default_val is not None:
        lines.append(f'                DefaultValue ="{default_val}"')
    lines.append(f'            End')
    return '\n'.join(lines)


def _combo(name, left, top, width, height, row_source=""):
    lines = []
    lines.append(f'            Begin ComboBox')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'                FontName ="Yu Gothic UI"')
    lines.append(f'                FontSize =9')
    lines.append(f'                ColumnCount =1')
    if row_source:
        lines.append(f'                RowSourceType ="Table/Query"')
        lines.append(f'                RowSource ="{row_source}"')
    lines.append(f'            End')
    return '\n'.join(lines)


def _listbox(name, left, top, width, height, col_count=2, col_widths="", column_heads=True):
    lines = []
    lines.append(f'            Begin ListBox')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'                FontName ="Yu Gothic UI"')
    lines.append(f'                FontSize =9')
    lines.append(f'                ColumnCount ={col_count}')
    if col_widths:
        lines.append(f'                ColumnWidths ="{col_widths}"')
    if column_heads:
        lines.append(f'                ColumnHeads =-1')
    lines.append(f'                RowSourceType ="Table/Query"')
    lines.append(f'            End')
    return '\n'.join(lines)


def _checkbox(name, left, top, width, height):
    lines = []
    lines.append(f'            Begin CheckBox')
    lines.append(f'                Left ={left}')
    lines.append(f'                Top ={top}')
    lines.append(f'                Width ={width}')
    lines.append(f'                Height ={height}')
    lines.append(f'                Name ="{name}"')
    lines.append(f'            End')
    return '\n'.join(lines)


def _footer(form_name, vba_code=""):
    lines = []
    lines.append(f'        End')  # End Detail
    lines.append(f'    End')
    lines.append(f'End')
    if vba_code:
        lines.append(f'CodeBehindForm')
        lines.append(vba_code)
    return '\n'.join(lines)


CM = 567  # 1cm in twips


# ============================================================
# 各フォーム定義
# ============================================================

def form_F_Main():
    parts = [_header("F_Main", "SalesMgr 営業管理", width=10000)]

    # タイトル
    parts.append(_label("lblTitle", "SalesMgr 営業管理", CM*1, CM*0, CM*14, int(CM*1.5),
                        font_size=18, bold=True, fore_color=3877662))

    # ボタン
    buttons = [
        ("btnDaily", "日次登録"), ("btnTargets", "目標設定"),
        ("btnReferrals", "送客登録"), ("btnReport", "レポート"),
        ("btnRanking", "ランキング"), ("btnMembers", "担当者管理"),
    ]
    for i, (bname, bcaption) in enumerate(buttons):
        col = i % 2
        row = i // 2
        x = int(CM*0.5) + col * int(CM*7)
        y = int(CM*2) + row * int(CM*1.5)
        parts.append(_button(bname, bcaption, x, y, int(CM*6), int(CM*1.2), font_size=11, on_click="=1"))

    vba = _get_vba_main()
    parts.append(_footer("F_Main", vba))
    return '\n'.join(parts)


def form_F_Daily():
    parts = [_header("F_Daily", "日次一覧", width=14000)]

    # 月ナビ
    parts.append(_button("btnPrev", "◀", int(CM*0.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_label("lblMonth", "YYYY年MM月", int(CM*2), int(CM*0.3), int(CM*5), int(CM*0.7), font_size=12, bold=True))
    parts.append(_button("btnNext", "▶", int(CM*7.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))

    # メンバーフィルタ
    parts.append(_label("lblFilter", "担当者:", int(CM*10), int(CM*0.3), int(CM*2), int(CM*0.7)))
    parts.append(_combo("cboMember", int(CM*12), int(CM*0.3), int(CM*3.5), int(CM*0.7)))

    # ボタン
    parts.append(_button("btnAdd", "新規登録", int(CM*0.5), int(CM*1.3), int(CM*2.5), int(CM*0.7), on_click="=1"))
    parts.append(_button("btnEdit", "編集", int(CM*3.5), int(CM*1.3), int(CM*2), int(CM*0.7), on_click="=1"))
    parts.append(_button("btnDelete", "削除", int(CM*6), int(CM*1.3), int(CM*2), int(CM*0.7), on_click="=1"))

    # リスト
    parts.append(_listbox("lstRecords", int(CM*0.5), int(CM*2.5), int(CM*22), int(CM*12),
                           col_count=11, col_widths="0cm;2.5cm;2.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;1.5cm;3cm"))

    vba = _get_vba_daily()
    parts.append(_footer("F_Daily", vba))
    return '\n'.join(parts)


def form_F_DailyEdit():
    parts = [_header("F_DailyEdit", "日次レコード", width=6000, popup=True, modal=True)]

    y = int(CM*0.5)

    # 日付 + 担当者
    parts.append(_label("lbl1", "日付:", int(CM*0.3), y, int(CM*1.5), int(CM*0.6)))
    parts.append(_textbox("txtRecDate", int(CM*2), y, int(CM*3), int(CM*0.6)))
    y += int(CM*0.8)
    parts.append(_label("lbl2", "担当者:", int(CM*0.3), y, int(CM*1.5), int(CM*0.6)))
    parts.append(_combo("cboMember", int(CM*2), y, int(CM*3), int(CM*0.6)))
    y += int(CM*0.8)
    parts.append(_label("lbl3", "稼働時間:", int(CM*0.3), y, int(CM*1.5), int(CM*0.6)))
    parts.append(_textbox("txtWorkHours", int(CM*2), y, int(CM*1.5), int(CM*0.6), "8"))
    parts.append(_label("lbl3b", "出勤:", int(CM*4), y, int(CM*1), int(CM*0.6)))
    parts.append(_checkbox("chkWorkDay", int(CM*5), y, int(CM*0.5), int(CM*0.5)))
    y += int(CM*1)

    # 時間帯別
    parts.append(_label("lblHourly", "時間帯別架電:", int(CM*0.3), y, int(CM*3), int(CM*0.6), bold=True))
    y += int(CM*0.7)
    for h in range(10, 19):
        parts.append(_label(f"lblH{h}", f"{h}時:", int(CM*0.3), y, int(CM*1.2), int(CM*0.5)))
        parts.append(_textbox(f"txtC{h}", int(CM*1.5), y, int(CM*1.5), int(CM*0.5), "0"))
        y += int(CM*0.6)

    y += int(CM*0.3)
    # 成果項目
    for lbl, nm in [("有効:", "txtValid"), ("見込:", "txtProspect"), ("資料:", "txtDoc"),
                    ("追客:", "txtFollow"), ("受注:", "txtReceived"), ("送客:", "txtReferral")]:
        parts.append(_label(f"lbl_{nm}", lbl, int(CM*0.3), y, int(CM*1.5), int(CM*0.5)))
        parts.append(_textbox(nm, int(CM*2), y, int(CM*1.5), int(CM*0.5), "0"))
        y += int(CM*0.6)

    # 備考
    parts.append(_label("lblNote", "備考:", int(CM*0.3), y, int(CM*1.5), int(CM*0.5)))
    parts.append(_textbox("txtNote", int(CM*2), y, int(CM*4), int(CM*0.8)))
    y += int(CM*1.2)

    # ボタン
    parts.append(_button("btnSave", "保存", int(CM*0.5), y, int(CM*2.5), int(CM*0.7), on_click="=1"))
    parts.append(_button("btnCancel", "キャンセル", int(CM*3.5), y, int(CM*2.5), int(CM*0.7), on_click="=1"))

    vba = _get_vba_daily_edit()
    parts.append(_footer("F_DailyEdit", vba))
    return '\n'.join(parts)


def form_F_Members():
    parts = [_header("F_Members", "担当者管理", width=12000)]

    parts.append(_label("lblTitle", "担当者管理", int(CM*0.5), int(CM*0.3), int(CM*6), int(CM*0.8), font_size=14, bold=True))

    # 追加
    parts.append(_label("lblAdd", "新規担当者:", int(CM*0.5), int(CM*1.5), int(CM*2.5), int(CM*0.6)))
    parts.append(_textbox("txtNewName", int(CM*3), int(CM*1.5), int(CM*3.5), int(CM*0.6)))
    parts.append(_button("btnAdd", "追加", int(CM*7), int(CM*1.5), int(CM*2), int(CM*0.6), on_click="=1"))

    # 有効一覧
    parts.append(_label("lblActive", "有効な担当者:", int(CM*0.5), int(CM*2.8), int(CM*4), int(CM*0.6), bold=True))
    parts.append(_listbox("lstActive", int(CM*0.5), int(CM*3.5), int(CM*5), int(CM*5), 2, "0cm;4cm", False))
    parts.append(_button("btnDeactivate", "無効にする ▶", int(CM*6), int(CM*5), int(CM*3), int(CM*0.7), on_click="=1"))

    # 無効一覧
    parts.append(_label("lblInactive", "無効な担当者:", int(CM*9.5), int(CM*2.8), int(CM*4), int(CM*0.6), bold=True))
    parts.append(_listbox("lstInactive", int(CM*9.5), int(CM*3.5), int(CM*5), int(CM*5), 2, "0cm;4cm", False))
    parts.append(_button("btnActivate", "◀ 有効にする", int(CM*6), int(CM*6.5), int(CM*3), int(CM*0.7), on_click="=1"))

    vba = _get_vba_members()
    parts.append(_footer("F_Members", vba))
    return '\n'.join(parts)


def form_F_Targets():
    parts = [_header("F_Targets", "目標設定", width=14000)]

    parts.append(_button("btnPrev", "◀", int(CM*0.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_label("lblMonth", "YYYY年MM月", int(CM*2), int(CM*0.3), int(CM*5), int(CM*0.7), font_size=12, bold=True))
    parts.append(_button("btnNext", "▶", int(CM*7.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))

    y = int(CM*1.5)
    parts.append(_label("lbl_m", "担当者:", int(CM*0.5), y, int(CM*2), int(CM*0.6)))
    parts.append(_combo("cboMember", int(CM*2.5), y, int(CM*3.5), int(CM*0.6)))
    parts.append(_button("btnLoadPrev", "前月コピー", int(CM*6.5), y, int(CM*2.5), int(CM*0.6), on_click="=1"))
    y += int(CM*0.9)

    for lbl, nm in [("稼働日数:", "txtPlanDays"), ("日稼働h:", "txtHoursPerDay"),
                    ("架電目標:", "txtTgtCalls"), ("有効目標:", "txtTgtValid"),
                    ("見込目標:", "txtTgtProspect"), ("受注目標:", "txtTgtReceived"),
                    ("送客目標:", "txtTgtReferral")]:
        parts.append(_label(f"lbl_{nm}", lbl, int(CM*0.5), y, int(CM*2), int(CM*0.6)))
        parts.append(_textbox(nm, int(CM*2.5), y, int(CM*2.5), int(CM*0.6)))
        y += int(CM*0.7)

    parts.append(_button("btnSave", "保存", int(CM*0.5), y + int(CM*0.2), int(CM*2.5), int(CM*0.7), on_click="=1"))
    parts.append(_button("btnDelete", "削除", int(CM*3.5), y + int(CM*0.2), int(CM*2.5), int(CM*0.7), on_click="=1"))

    parts.append(_listbox("lstTargets", int(CM*0.5), y + int(CM*1.3), int(CM*20), int(CM*4),
                           col_count=9, col_widths="0cm;2.5cm;2cm;2cm;2cm;2cm;2cm;2cm;2cm"))

    vba = _get_vba_targets()
    parts.append(_footer("F_Targets", vba))
    return '\n'.join(parts)


def form_F_Referrals():
    parts = [_header("F_Referrals", "送客登録", width=14000)]

    parts.append(_button("btnPrev", "◀", int(CM*0.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_label("lblMonth", "YYYY年MM月", int(CM*2), int(CM*0.3), int(CM*5), int(CM*0.7), font_size=12, bold=True))
    parts.append(_button("btnNext", "▶", int(CM*7.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))

    y = int(CM*1.5)
    parts.append(_label("lbl_d", "日付:", int(CM*0.5), y, int(CM*1.2), int(CM*0.6)))
    parts.append(_textbox("txtRefDate", int(CM*1.7), y, int(CM*2.5), int(CM*0.6)))
    parts.append(_label("lbl_m", "担当者:", int(CM*4.5), y, int(CM*1.5), int(CM*0.6)))
    parts.append(_combo("cboMember", int(CM*6), y, int(CM*3), int(CM*0.6)))
    parts.append(_label("lbl_c", "件数:", int(CM*9.5), y, int(CM*1), int(CM*0.6)))
    parts.append(_textbox("txtRefCount", int(CM*10.5), y, int(CM*1.5), int(CM*0.6), "1"))
    parts.append(_button("btnAdd", "追加", int(CM*12.5), y, int(CM*2), int(CM*0.6), on_click="=1"))

    y += int(CM*1)
    parts.append(_button("btnDelete", "削除", int(CM*0.5), y, int(CM*2), int(CM*0.6), on_click="=1"))
    y += int(CM*0.8)

    parts.append(_listbox("lstRefs", int(CM*0.5), y, int(CM*15), int(CM*7),
                           col_count=4, col_widths="0cm;3cm;3cm;2cm"))

    vba = _get_vba_referrals()
    parts.append(_footer("F_Referrals", vba))
    return '\n'.join(parts)


def form_F_Report():
    parts = [_header("F_Report", "レポート", width=14000)]

    parts.append(_button("btnPrev", "◀", int(CM*0.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_label("lblMonth", "YYYY年MM月 レポート", int(CM*2), int(CM*0.3), int(CM*8), int(CM*0.7), font_size=12, bold=True))
    parts.append(_button("btnNext", "▶", int(CM*10.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_button("btnPDF", "PDF出力", int(CM*12.5), int(CM*0.3), int(CM*3), int(CM*0.7), on_click="=1"))

    y = int(CM*1.3)
    # アラート
    parts.append(_label("lblAlert", "", int(CM*0.5), y, int(CM*20), int(CM*1), fore_color=2498780))
    y += int(CM*1.3)

    # KPIカード 1行目
    kpis1 = [("架電件数", "lblCalls", "lblCallsPrev"),
             ("有効件数", "lblValid", "lblValidPrev"),
             ("見込件数", "lblProsp", "lblProspPrev"),
             ("受注件数", "lblRecv", "lblRecvPrev")]
    for i, (title, nm_val, nm_sub) in enumerate(kpis1):
        x = int(CM*0.5) + i * int(CM*5)
        parts.append(_label(f"lblK{i}", title, x, y, int(CM*4.5), int(CM*0.4), font_size=7, fore_color=6710886))
        parts.append(_label(nm_val, "-", x, y+int(CM*0.4), int(CM*4.5), int(CM*0.7), font_size=11, bold=True))
        parts.append(_label(nm_sub, "", x, y+int(CM*1.1), int(CM*4.5), int(CM*0.4), font_size=7, fore_color=10066329))
    y += int(CM*1.8)

    # KPIカード 2行目
    kpis2 = [("有効率", "lblValidRate"), ("受注率", "lblRecvRate"),
             ("稼働時間", "lblHours"), ("生産性", "lblProductivity")]
    for i, (title, nm_val) in enumerate(kpis2):
        x = int(CM*0.5) + i * int(CM*5)
        parts.append(_label(f"lblK2{i}", title, x, y, int(CM*4.5), int(CM*0.4), font_size=7, fore_color=6710886))
        parts.append(_label(nm_val, "-", x, y+int(CM*0.4), int(CM*4.5), int(CM*0.7), font_size=11, bold=True))
    y += int(CM*1.5)

    # ランキング
    parts.append(_label("lblR1", "送客ランキング", int(CM*0.5), y, int(CM*6), int(CM*0.6), bold=True))
    parts.append(_label("lblR2", "受注ランキング", int(CM*7), y, int(CM*6), int(CM*0.6), bold=True))
    parts.append(_label("lblR3", "見込ランキング", int(CM*13.5), y, int(CM*6), int(CM*0.6), bold=True))
    y += int(CM*0.7)

    parts.append(_listbox("lstRankRef", int(CM*0.5), y, int(CM*6), int(CM*5), 2, "3cm;2cm"))
    parts.append(_listbox("lstRankRecv", int(CM*7), y, int(CM*6), int(CM*5), 2, "3cm;2cm"))
    parts.append(_listbox("lstRankProsp", int(CM*13.5), y, int(CM*6), int(CM*5), 2, "3cm;2cm"))

    vba = _get_vba_report()
    parts.append(_footer("F_Report", vba))
    return '\n'.join(parts)


def form_F_Ranking():
    parts = [_header("F_Ranking", "ランキング", width=14000)]

    parts.append(_button("btnPrev", "◀", int(CM*0.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))
    parts.append(_label("lblMonth", "YYYY年MM月", int(CM*2), int(CM*0.3), int(CM*5), int(CM*0.7), font_size=12, bold=True))
    parts.append(_button("btnNext", "▶", int(CM*7.5), int(CM*0.3), int(CM*1.2), int(CM*0.7), on_click="=1"))

    y = int(CM*1.3)
    parts.append(_label("lblR1", "送客", int(CM*0.5), y, int(CM*6), int(CM*0.6), bold=True))
    parts.append(_label("lblR2", "受注", int(CM*7), y, int(CM*6), int(CM*0.6), bold=True))
    parts.append(_label("lblR3", "見込", int(CM*13.5), y, int(CM*6), int(CM*0.6), bold=True))
    y += int(CM*0.7)

    parts.append(_listbox("lstRef", int(CM*0.5), y, int(CM*6), int(CM*7), 2, "3cm;2cm"))
    parts.append(_listbox("lstRecv", int(CM*7), y, int(CM*6), int(CM*7), 2, "3cm;2cm"))
    parts.append(_listbox("lstProsp", int(CM*13.5), y, int(CM*6), int(CM*7), 2, "3cm;2cm"))

    vba = _get_vba_ranking()
    parts.append(_footer("F_Ranking", vba))
    return '\n'.join(parts)


# ============================================================
# VBAコード（各フォーム用）
# ============================================================

def _get_vba_main():
    return """
Attribute VB_Name = "Form_F_Main"
Option Compare Database
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
"""

def _get_vba_daily():
    return """
Attribute VB_Name = "Form_F_Daily"
Option Compare Database
Option Explicit
Private m_Y As Integer, m_M As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Y = Year(Date): m_M = Month(Date): LoadData
End Sub

Private Sub LoadData()
    Me.Caption = m_Y & "年" & Format(m_M, "00") & "月 日次一覧"
    lblMonth.Caption = m_Y & "年" & m_M & "月"
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    Dim sql As String
    sql = "SELECT R.ID, Format(R.rec_date,'yyyy/mm/dd'), R.member_name, R.calls, R.valid_count, R.prospect, R.doc, R.follow_up, R.received, R.work_hours, R.[note] FROM T_RECORDS AS R WHERE R.rec_date >= #" & Format(DateSerial(m_Y,m_M,1),"yyyy/mm/dd") & "# AND R.rec_date < #" & Format(DateSerial(m_Y,m_M+1,1),"yyyy/mm/dd") & "# "
    If Nz(cboMember.Value,"") <> "" Then sql = sql & "AND R.member_name='" & cboMember.Value & "' "
    sql = sql & "ORDER BY R.rec_date DESC, R.member_name"
    lstRecords.RowSource = sql: lstRecords.Requery
End Sub

Private Sub btnPrev_Click()
    If m_M=1 Then m_Y=m_Y-1: m_M=12 Else m_M=m_M-1
    LoadData
End Sub
Private Sub btnNext_Click()
    If m_M=12 Then m_Y=m_Y+1: m_M=1 Else m_M=m_M+1
    LoadData
End Sub
Private Sub cboMember_AfterUpdate()
    LoadData
End Sub
Private Sub btnAdd_Click()
    DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"ADD|" & m_Y & "|" & m_M: LoadData
End Sub
Private Sub btnEdit_Click()
    If IsNull(lstRecords.Value) Then MsgBox "選択してください",vbExclamation: Exit Sub
    DoCmd.OpenForm "F_DailyEdit",,,,,acDialog,"EDIT|" & lstRecords.Value: LoadData
End Sub
Private Sub btnDelete_Click()
    If IsNull(lstRecords.Value) Then MsgBox "選択してください",vbExclamation: Exit Sub
    If MsgBox("削除しますか？",vbYesNo+vbQuestion)=vbYes Then
        CurrentDb.Execute "DELETE FROM T_RECORDS WHERE ID=" & lstRecords.Value, dbFailOnError: LoadData
    End If
End Sub
Private Sub lstRecords_DblClick(Cancel As Integer)
    btnEdit_Click
End Sub
"""

def _get_vba_daily_edit():
    return """
Attribute VB_Name = "Form_F_DailyEdit"
Option Compare Database
Option Explicit
Private m_Mode As String, m_ID As Long

Private Sub Form_Open(Cancel As Integer)
    Dim p() As String: p = Split(Nz(Me.OpenArgs,"ADD"),"|"): m_Mode = p(0)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    If m_Mode="EDIT" And UBound(p)>=1 Then
        m_ID = CLng(p(1)): Me.Caption = "編集": LoadRec
    Else
        m_ID = 0: Me.Caption = "新規登録"
        If UBound(p)>=2 Then txtRecDate.Value = Format(DateSerial(CInt(p(1)),CInt(p(2)),Day(Date)),"yyyy/mm/dd") Else txtRecDate.Value = Format(Date,"yyyy/mm/dd")
        txtWorkHours.Value = 8: chkWorkDay.Value = True
    End If
End Sub

Private Sub LoadRec()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_RECORDS WHERE ID=" & m_ID)
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
    If Nz(txtRecDate.Value,"")="" Then MsgBox "日付を入力",vbExclamation: Exit Sub
    If Nz(cboMember.Value,"")="" Then MsgBox "担当者を選択",vbExclamation: Exit Sub
    Dim dt As Date: dt = CDate(txtRecDate.Value)
    Dim c As Long: Dim h As Integer: c = 0: For h = 10 To 18: c = c + Nz(Me("txtC" & h),0): Next h
    Dim sql As String
    If m_ID = 0 Then
        sql = "INSERT INTO T_RECORDS (rec_date,member_name,calls,calls_10,calls_11,calls_12,calls_13,calls_14,calls_15,calls_16,calls_17,calls_18,valid_count,prospect,doc,follow_up,received,work_hours,[note],referral,work_day) VALUES (#" & Format(dt,"yyyy/mm/dd") & "#,'" & cboMember.Value & "'," & c
        For h = 10 To 18: sql = sql & "," & Nz(Me("txtC" & h),0): Next h
        sql = sql & "," & Nz(txtValid,0) & "," & Nz(txtProspect,0) & "," & Nz(txtDoc,0) & "," & Nz(txtFollow,0) & "," & Nz(txtReceived,0) & "," & Nz(txtWorkHours,8) & ",'" & Replace(Nz(txtNote,""),"'","''") & "'," & Nz(txtReferral,0) & "," & IIf(chkWorkDay.Value,"True","False") & ")"
    Else
        sql = "UPDATE T_RECORDS SET rec_date=#" & Format(dt,"yyyy/mm/dd") & "#,member_name='" & cboMember.Value & "',calls=" & c
        For h = 10 To 18: sql = sql & ",calls_" & h & "=" & Nz(Me("txtC" & h),0): Next h
        sql = sql & ",valid_count=" & Nz(txtValid,0) & ",prospect=" & Nz(txtProspect,0) & ",doc=" & Nz(txtDoc,0) & ",follow_up=" & Nz(txtFollow,0) & ",received=" & Nz(txtReceived,0) & ",work_hours=" & Nz(txtWorkHours,8) & ",[note]='" & Replace(Nz(txtNote,""),"'","''") & "',referral=" & Nz(txtReferral,0) & ",work_day=" & IIf(chkWorkDay.Value,"True","False") & " WHERE ID=" & m_ID
    End If
    CurrentDb.Execute sql, dbFailOnError: DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub
"""

def _get_vba_members():
    return """
Attribute VB_Name = "Form_F_Members"
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    LoadData
End Sub

Private Sub LoadData()
    lstActive.RowSource = "SELECT ID,member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    lstInactive.RowSource = "SELECT ID,member_name FROM T_MEMBERS WHERE active=False ORDER BY member_name"
    lstActive.Requery: lstInactive.Requery
End Sub

Private Sub btnAdd_Click()
    Dim nm As String: nm = Trim(Nz(txtNewName.Value,""))
    If nm="" Then MsgBox "名前を入力",vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_MEMBERS (member_name,active) VALUES ('" & Replace(nm,"'","''") & "',True)", dbFailOnError
    txtNewName.Value = "": LoadData
End Sub
Private Sub btnDeactivate_Click()
    If IsNull(lstActive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=False WHERE ID=" & lstActive.Value, dbFailOnError: LoadData
End Sub
Private Sub btnActivate_Click()
    If IsNull(lstInactive.Value) Then Exit Sub
    CurrentDb.Execute "UPDATE T_MEMBERS SET active=True WHERE ID=" & lstInactive.Value, dbFailOnError: LoadData
End Sub
"""

def _get_vba_targets():
    return """
Attribute VB_Name = "Form_F_Targets"
Option Compare Database
Option Explicit
Private m_Y As Integer, m_M As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Y = Year(Date): m_M = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
    lblMonth.Caption = m_Y & "年" & m_M & "月"
    lstTargets.RowSource = "SELECT T.ID,T.member_name,T.target_calls,T.target_valid,T.target_prospect,T.target_received,T.target_referral,T.plan_days,T.work_hours_per_day FROM T_MEMBER_TARGETS AS T WHERE T.target_year=" & m_Y & " AND T.target_month=" & m_M & " ORDER BY T.member_name"
    lstTargets.Requery
End Sub

Private Sub btnPrev_Click()
    If m_M=1 Then m_Y=m_Y-1: m_M=12 Else m_M=m_M-1: LoadData
End Sub
Private Sub btnNext_Click()
    If m_M=12 Then m_Y=m_Y+1: m_M=1 Else m_M=m_M+1: LoadData
End Sub

Private Sub btnSave_Click()
    If Nz(cboMember.Value,"")="" Then MsgBox "担当者を選択",vbExclamation: Exit Sub
    Dim nm As String: nm = cboMember.Value
    Dim cnt As Long: cnt = DCount("*","T_MEMBER_TARGETS","member_name='" & nm & "' AND target_year=" & m_Y & " AND target_month=" & m_M)
    If cnt > 0 Then
        CurrentDb.Execute "UPDATE T_MEMBER_TARGETS SET plan_days=" & Nz(txtPlanDays,20) & ",work_hours_per_day=" & Nz(txtHoursPerDay,8) & ",target_calls=" & Nz(txtTgtCalls,0) & ",target_valid=" & Nz(txtTgtValid,0) & ",target_prospect=" & Nz(txtTgtProspect,0) & ",target_received=" & Nz(txtTgtReceived,0) & ",target_referral=" & Nz(txtTgtReferral,0) & " WHERE member_name='" & nm & "' AND target_year=" & m_Y & " AND target_month=" & m_M, dbFailOnError
    Else
        CurrentDb.Execute "INSERT INTO T_MEMBER_TARGETS (member_name,target_year,target_month,plan_days,work_hours_per_day,target_calls,target_valid,target_prospect,target_received,target_referral) VALUES ('" & nm & "'," & m_Y & "," & m_M & "," & Nz(txtPlanDays,20) & "," & Nz(txtHoursPerDay,8) & "," & Nz(txtTgtCalls,0) & "," & Nz(txtTgtValid,0) & "," & Nz(txtTgtProspect,0) & "," & Nz(txtTgtReceived,0) & "," & Nz(txtTgtReferral,0) & ")", dbFailOnError
    End If
    LoadData: MsgBox nm & " 保存完了", vbInformation
End Sub

Private Sub btnLoadPrev_Click()
    If Nz(cboMember.Value,"")="" Then Exit Sub
    Dim py As Integer, pm As Integer: py = m_Y: pm = m_M
    If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_MEMBER_TARGETS WHERE member_name='" & cboMember.Value & "' AND target_year=" & py & " AND target_month=" & pm)
    If Not rs.EOF Then
        txtPlanDays.Value = rs("plan_days"): txtHoursPerDay.Value = rs("work_hours_per_day")
        txtTgtCalls.Value = rs("target_calls"): txtTgtValid.Value = rs("target_valid")
        txtTgtProspect.Value = rs("target_prospect"): txtTgtReceived.Value = rs("target_received")
        txtTgtReferral.Value = rs("target_referral")
    Else: MsgBox "前月の目標なし",vbInformation
    End If: rs.Close
End Sub

Private Sub btnDelete_Click()
    If IsNull(lstTargets.Value) Then Exit Sub
    If MsgBox("削除しますか？",vbYesNo+vbQuestion)=vbYes Then
        CurrentDb.Execute "DELETE FROM T_MEMBER_TARGETS WHERE ID=" & lstTargets.Value, dbFailOnError: LoadData
    End If
End Sub
"""

def _get_vba_referrals():
    return """
Attribute VB_Name = "Form_F_Referrals"
Option Compare Database
Option Explicit
Private m_Y As Integer, m_M As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Y = Year(Date): m_M = Month(Date)
    cboMember.RowSource = "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name"
    LoadData
End Sub

Private Sub LoadData()
    lblMonth.Caption = m_Y & "年" & m_M & "月": txtRefDate.Value = Format(Date,"yyyy/mm/dd")
    lstRefs.RowSource = "SELECT R.ID,Format(R.rec_date,'yyyy/mm/dd'),R.member_name,R.ref_count FROM T_REFERRALS AS R WHERE R.rec_date>=#" & Format(DateSerial(m_Y,m_M,1),"yyyy/mm/dd") & "# AND R.rec_date<#" & Format(DateSerial(m_Y,m_M+1,1),"yyyy/mm/dd") & "# ORDER BY R.rec_date DESC,R.member_name"
    lstRefs.Requery
End Sub

Private Sub btnPrev_Click()
    If m_M=1 Then m_Y=m_Y-1: m_M=12 Else m_M=m_M-1: LoadData
End Sub
Private Sub btnNext_Click()
    If m_M=12 Then m_Y=m_Y+1: m_M=1 Else m_M=m_M+1: LoadData
End Sub
Private Sub btnAdd_Click()
    If Nz(cboMember.Value,"")="" Or Nz(txtRefDate.Value,"")="" Then MsgBox "入力してください",vbExclamation: Exit Sub
    CurrentDb.Execute "INSERT INTO T_REFERRALS (rec_date,member_name,ref_count) VALUES (#" & Format(CDate(txtRefDate.Value),"yyyy/mm/dd") & "#,'" & cboMember.Value & "'," & Nz(txtRefCount,0) & ")", dbFailOnError: LoadData
End Sub
Private Sub btnDelete_Click()
    If IsNull(lstRefs.Value) Then Exit Sub
    If MsgBox("削除しますか？",vbYesNo+vbQuestion)=vbYes Then
        CurrentDb.Execute "DELETE FROM T_REFERRALS WHERE ID=" & lstRefs.Value, dbFailOnError: LoadData
    End If
End Sub
"""

def _get_vba_report():
    return """
Attribute VB_Name = "Form_F_Report"
Option Compare Database
Option Explicit
Private m_Y As Integer, m_M As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Y = Year(Date): m_M = Month(Date): LoadReport
End Sub

Private Sub btnPrev_Click()
    If m_M=1 Then m_Y=m_Y-1: m_M=12 Else m_M=m_M-1: LoadReport
End Sub
Private Sub btnNext_Click()
    If m_M=12 Then m_Y=m_Y+1: m_M=1 Else m_M=m_M+1: LoadReport
End Sub
Private Sub btnPDF_Click()
    MsgBox "PDF出力は今後対応予定です", vbInformation
End Sub

Private Sub LoadReport()
    lblMonth.Caption = m_Y & "年" & m_M & "月 レポート"
    Dim dtF As String, dtT As String
    dtF = Format(DateSerial(m_Y,m_M,1),"yyyy/mm/dd"): dtT = Format(DateSerial(m_Y,m_M+1,1),"yyyy/mm/dd")

    ' チーム集計
    Dim rs As DAO.Recordset
    Dim qd As DAO.QueryDef: Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(m_Y,m_M,1): qd.Parameters("prmDateTo") = DateSerial(m_Y,m_M+1,1)
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim tC As Long, tV As Long, tP As Long, tR As Long, tH As Double
    If Not rs.EOF Then tC = Nz(rs("sum_calls"),0): tV = Nz(rs("sum_valid"),0): tP = Nz(rs("sum_prospect"),0): tR = Nz(rs("sum_received"),0): tH = Nz(rs("sum_hours"),0)
    rs.Close

    ' 目標
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Targets")
    qd.Parameters("prmYear") = m_Y: qd.Parameters("prmMonth") = m_M
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim gC As Long, gV As Long, gP As Long, gR As Long
    If Not rs.EOF Then gC = Nz(rs("sum_tgt_calls"),0): gV = Nz(rs("sum_tgt_valid"),0): gP = Nz(rs("sum_tgt_prospect"),0): gR = Nz(rs("sum_tgt_received"),0)
    rs.Close

    ' 前月
    Dim py As Integer, pm As Integer: py = m_Y: pm = m_M
    If pm=1 Then py=py-1: pm=12 Else pm=pm-1
    Set qd = CurrentDb.QueryDefs("Q_Team_Monthly_Sum")
    qd.Parameters("prmDateFrom") = DateSerial(py,pm,1): qd.Parameters("prmDateTo") = DateSerial(py,pm+1,1)
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Dim pC As Long, pV As Long, pP As Long, pR As Long
    If Not rs.EOF Then pC = Nz(rs("sum_calls"),0): pV = Nz(rs("sum_valid"),0): pP = Nz(rs("sum_prospect"),0): pR = Nz(rs("sum_received"),0)
    rs.Close

    ' KPI表示
    lblCalls.Caption = Format(tC,"#,##0") & " / " & Format(gC,"#,##0")
    lblValid.Caption = Format(tV,"#,##0") & " / " & Format(gV,"#,##0")
    lblProsp.Caption = Format(tP,"#,##0") & " / " & Format(gP,"#,##0")
    lblRecv.Caption = Format(tR,"#,##0") & " / " & Format(gR,"#,##0")

    Dim arrF As String
    arrF = IIf(tC>pC, Chr(9650) & Format(tC-pC,"#,##0"), IIf(tC<pC, Chr(9660) & Format(pC-tC,"#,##0"), "-"))
    lblCallsPrev.Caption = "前月" & Format(pC,"#,##0") & " " & arrF
    arrF = IIf(tV>pV, Chr(9650) & Format(tV-pV,"#,##0"), IIf(tV<pV, Chr(9660) & Format(pV-tV,"#,##0"), "-"))
    lblValidPrev.Caption = "前月" & Format(pV,"#,##0") & " " & arrF
    arrF = IIf(tP>pP, Chr(9650) & Format(tP-pP,"#,##0"), IIf(tP<pP, Chr(9660) & Format(pP-tP,"#,##0"), "-"))
    lblProspPrev.Caption = "前月" & Format(pP,"#,##0") & " " & arrF
    arrF = IIf(tR>pR, Chr(9650) & Format(tR-pR,"#,##0"), IIf(tR<pR, Chr(9660) & Format(pR-tR,"#,##0"), "-"))
    lblRecvPrev.Caption = "前月" & Format(pR,"#,##0") & " " & arrF

    lblValidRate.Caption = IIf(tC>0, Format(tV/tC*100,"0.0") & "%", "-")
    lblRecvRate.Caption = IIf(tC>0, Format(tR/tC*100,"0.0") & "%", "-")
    lblHours.Caption = Format(tH,"#,##0") & "h"
    lblProductivity.Caption = IIf(tH>0, Format(tP/tH,"0.000"), "-")

    ' アラート
    Dim al As String: al = ""
    If gR > 0 Then
        If tR < gR Then al = al & Chr(9651) & " 受注 残り" & (gR-tR) & "件 " & tR & "/" & gR & vbCrLf Else al = al & Chr(9675) & " 受注 目標達成！" & vbCrLf
    End If
    If gP > 0 Then
        If tP < gP Then al = al & Chr(9651) & " 見込 残り" & (gP-tP) & "件 " & tP & "/" & gP Else al = al & Chr(9675) & " 見込 目標達成！"
    End If
    lblAlert.Caption = al

    ' ランキング
    lstRankRef.RowSource = "SELECT R.member_name, Sum(R.referral) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRankRecv.RowSource = "SELECT R.member_name, Sum(R.received) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstRankProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRankRef.Requery: lstRankRecv.Requery: lstRankProsp.Requery
End Sub
"""

def _get_vba_ranking():
    return """
Attribute VB_Name = "Form_F_Ranking"
Option Compare Database
Option Explicit
Private m_Y As Integer, m_M As Integer

Private Sub Form_Open(Cancel As Integer)
    m_Y = Year(Date): m_M = Month(Date): LoadData
End Sub

Private Sub btnPrev_Click()
    If m_M=1 Then m_Y=m_Y-1: m_M=12 Else m_M=m_M-1: LoadData
End Sub
Private Sub btnNext_Click()
    If m_M=12 Then m_Y=m_Y+1: m_M=1 Else m_M=m_M+1: LoadData
End Sub

Private Sub LoadData()
    lblMonth.Caption = m_Y & "年" & m_M & "月"
    Dim dtF As String, dtT As String
    dtF = Format(DateSerial(m_Y,m_M,1),"yyyy/mm/dd"): dtT = Format(DateSerial(m_Y,m_M+1,1),"yyyy/mm/dd")
    lstRef.RowSource = "SELECT R.member_name, Sum(R.referral) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.referral) DESC"
    lstRecv.RowSource = "SELECT R.member_name, Sum(R.received) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.received) DESC"
    lstProsp.RowSource = "SELECT R.member_name, Sum(R.prospect) FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.active=True AND R.rec_date>=#" & dtF & "# AND R.rec_date<#" & dtT & "# GROUP BY R.member_name ORDER BY Sum(R.prospect) DESC"
    lstRef.Requery: lstRecv.Requery: lstProsp.Requery
End Sub
"""


# ============================================================
# メイン: LoadFromText でインポート
# ============================================================
def main():
    print("SalesMgr フォーム・VBA 一括作成")
    print("=" * 50)

    os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
    time.sleep(2)

    import win32com.client
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(BE_PATH)
    time.sleep(2)

    # フォーム定義
    forms = {
        "F_Main": form_F_Main(),
        "F_DailyEdit": form_F_DailyEdit(),
        "F_Members": form_F_Members(),
        "F_Targets": form_F_Targets(),
        "F_Referrals": form_F_Referrals(),
        "F_Daily": form_F_Daily(),
        "F_Report": form_F_Report(),
        "F_Ranking": form_F_Ranking(),
    }

    tmp_dir = tempfile.mkdtemp()

    for form_name, form_def in forms.items():
        txt_path = os.path.join(tmp_dir, f"{form_name}.txt")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(form_def)

        # 既存フォーム削除
        try:
            app.DoCmd.DeleteObject(2, form_name)  # 2 = acForm
        except:
            pass

        try:
            app.LoadFromText(2, form_name, txt_path)  # 2 = acForm
            print(f"  {form_name} OK")
        except Exception as e:
            print(f"  {form_name} ERROR: {e}")

        time.sleep(0.3)

    # スタートアップ設定
    try:
        db = app.CurrentDb()
        props = db.Properties
        try:
            props("StartUpForm").Value = "F_Main"
        except:
            prop = db.CreateProperty("StartUpForm", 10, "F_Main")
            props.Append(prop)
        try:
            props("AppTitle").Value = "SalesMgr 営業管理"
        except:
            prop = db.CreateProperty("AppTitle", 10, "SalesMgr 営業管理")
            props.Append(prop)
        try:
            props("StartUpShowDBWindow").Value = False
        except:
            prop = db.CreateProperty("StartUpShowDBWindow", 1, False)
            props.Append(prop)
        print("  スタートアップ設定 OK")
    except Exception as e:
        print(f"  スタートアップ設定 ERROR: {e}")

    time.sleep(1)
    app.CloseCurrentDatabase()
    time.sleep(1)
    app.Quit()
    time.sleep(1)

    # 一時ファイル削除
    import shutil
    shutil.rmtree(tmp_dir, ignore_errors=True)

    print(f"\n完了: {BE_PATH}")
    print("SalesMgr_BE.accdb をダブルクリックで起動してください")


if __name__ == '__main__':
    main()
