# -*- coding: utf-8 -*-
"""
generate_pdf.py - SalesMgr Access DB から月次レポートPDFを生成
Usage: python generate_pdf.py --year YYYY --month MM --db "path/to/SalesMgr_BE.accdb"
"""
import argparse, os, sys, datetime, math

def main():
    parser = argparse.ArgumentParser(description="SalesMgr 月次レポートPDF生成")
    parser.add_argument('--year',  type=int, required=True)
    parser.add_argument('--month', type=int, required=True)
    parser.add_argument('--db',    type=str, required=True)
    args = parser.parse_args()

    year, month, db_path = args.year, args.month, args.db

    if not os.path.exists(db_path):
        print(f"ERROR: DB not found: {db_path}", file=sys.stderr)
        sys.exit(1)

    print(f"データ読込中... {year}年{month}月")
    data = read_access_data(db_path, year, month)

    print("PDF生成中...")
    pdf_path = build_pdf(year, month, data)

    print(f"完了: {pdf_path}")
    os.startfile(pdf_path)


# ──────────────────────────────────────────────
# データ取得（DAO経由 - Accessが開いていても共有アクセス可）
# ──────────────────────────────────────────────
def read_access_data(db_path, year, month):
    import win32com.client

    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    db = engine.OpenDatabase(db_path, False, True)   # read-only shared

    def date_range(y, m):
        dF = f"{y}/{m:02d}/01"
        dT = f"{y+1}/01/01" if m == 12 else f"{y}/{m+1:02d}/01"
        return dF, dT

    def query(sql):
        rs = db.OpenRecordset(sql)
        rows = []
        while not rs.EOF:
            rows.append([rs.Fields(i).Value for i in range(rs.Fields.Count)])
            rs.MoveNext()
        rs.Close()
        return rows

    def q1(sql):
        rows = query(sql)
        return rows[0] if rows else None

    dF, dT = date_range(year, month)

    # 有効担当者リスト
    members = [r[0] for r in query(
        "SELECT member_name FROM T_MEMBERS WHERE active=True ORDER BY member_name")]

    # チーム当月実績
    row = q1(f"""
        SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect),
               Sum(R.received), Sum(R.work_hours), Sum(R.referral),
               Sum(R.doc), Sum(R.follow_up)
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name
        WHERE M.active=True AND R.rec_date >= #{dF}# AND R.rec_date < #{dT}#
    """) or [0]*8
    team = dict(calls=int(row[0] or 0), valid=int(row[1] or 0),
                prospect=int(row[2] or 0), received=int(row[3] or 0),
                hours=float(row[4] or 0), referral=int(row[5] or 0),
                doc=int(row[6] or 0), follow=int(row[7] or 0))

    # チーム目標
    row = q1(f"""
        SELECT Sum(target_calls), Sum(target_valid), Sum(target_prospect),
               Sum(target_received), Sum(target_referral)
        FROM T_MEMBER_TARGETS WHERE target_year={year} AND target_month={month}
    """) or [0]*5
    tgt_team = dict(calls=int(row[0] or 0), valid=int(row[1] or 0),
                    prospect=int(row[2] or 0), received=int(row[3] or 0),
                    referral=int(row[4] or 0))

    # 担当者別当月実績
    member_data = {}
    for r in query(f"""
        SELECT R.member_name,
               Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received),
               Sum(R.work_hours), Sum(R.referral), Sum(R.doc), Sum(R.follow_up),
               Sum(R.calls_10), Sum(R.calls_11), Sum(R.calls_12), Sum(R.calls_13),
               Sum(R.calls_14), Sum(R.calls_15), Sum(R.calls_16), Sum(R.calls_17), Sum(R.calls_18)
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name
        WHERE M.active=True AND R.rec_date >= #{dF}# AND R.rec_date < #{dT}#
        GROUP BY R.member_name ORDER BY R.member_name
    """):
        m = r[0]
        member_data[m] = {
            'calls': int(r[1] or 0), 'valid': int(r[2] or 0),
            'prospect': int(r[3] or 0), 'received': int(r[4] or 0),
            'hours': float(r[5] or 0), 'referral': int(r[6] or 0),
            'doc': int(r[7] or 0), 'follow': int(r[8] or 0),
            **{f'c{h}': int(r[9+i] or 0) for i, h in enumerate(range(10, 19))}
        }

    # 担当者別目標
    tgt_map = {}
    for r in query(f"""
        SELECT member_name, plan_days, work_hours_per_day,
               target_calls, target_valid, target_prospect, target_received, target_referral
        FROM T_MEMBER_TARGETS WHERE target_year={year} AND target_month={month}
    """):
        tgt_map[r[0]] = dict(plan_days=int(r[1] or 20), hours_per_day=float(r[2] or 8),
                             calls=int(r[3] or 0), valid=int(r[4] or 0),
                             prospect=int(r[5] or 0), received=int(r[6] or 0),
                             referral=int(r[7] or 0))

    # 前月実績
    if month == 1: py, pm = year-1, 12
    else: py, pm = year, month-1
    pF, pT = date_range(py, pm)
    row = q1(f"""
        SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received), Sum(R.referral)
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name
        WHERE M.active=True AND R.rec_date >= #{pF}# AND R.rec_date < #{pT}#
    """) or [0]*5
    prev_team = dict(calls=int(row[0] or 0), valid=int(row[1] or 0),
                     prospect=int(row[2] or 0), received=int(row[3] or 0),
                     referral=int(row[4] or 0))

    # 過去6ヶ月推移
    trend6 = []
    cy, cm = year, month
    for _ in range(6):
        tF, tT = date_range(cy, cm)
        row = q1(f"""
            SELECT Sum(R.calls), Sum(R.valid_count), Sum(R.prospect), Sum(R.received)
            FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name
            WHERE M.active=True AND R.rec_date >= #{tF}# AND R.rec_date < #{tT}#
        """) or [0]*4
        trend6.insert(0, dict(year=cy, month=cm,
                               calls=int(row[0] or 0), valid=int(row[1] or 0),
                               prospect=int(row[2] or 0), received=int(row[3] or 0)))
        if cm == 1: cy, cm = cy-1, 12
        else: cm -= 1

    # 時間帯別（チーム合計）
    row = q1(f"""
        SELECT Sum(R.calls_10), Sum(R.calls_11), Sum(R.calls_12), Sum(R.calls_13),
               Sum(R.calls_14), Sum(R.calls_15), Sum(R.calls_16), Sum(R.calls_17), Sum(R.calls_18)
        FROM T_RECORDS AS R INNER JOIN T_MEMBERS AS M ON R.member_name=M.member_name
        WHERE M.active=True AND R.rec_date >= #{dF}# AND R.rec_date < #{dT}#
    """) or [0]*9
    hourly = {h: int(row[i] or 0) for i, h in enumerate(range(10, 19))}

    db.Close()

    return dict(members=members, team=team, tgt_team=tgt_team,
                prev_team=prev_team, member_data=member_data,
                tgt_map=tgt_map, trend6=trend6, hourly=hourly)


# ──────────────────────────────────────────────
# PDF生成
# ──────────────────────────────────────────────
def build_pdf(year, month, data):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # フォント登録（游ゴシック優先）
    FN = 'Helvetica'
    for fp in [r"C:\Windows\Fonts\YuGothM.ttc",
               r"C:\Windows\Fonts\yugothm.ttf",
               r"C:\Windows\Fonts\meiryo.ttc",
               r"C:\Windows\Fonts\msgothic.ttc"]:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('JP', fp))
                FN = 'JP'
                break
            except Exception:
                pass

    team      = data['team']
    tgt_team  = data['tgt_team']
    prev_team = data['prev_team']
    members   = data['members']
    mdata     = data['member_data']
    tgt_map   = data['tgt_map']
    trend6    = data['trend6']
    hourly    = data['hourly']

    def c(v):
        try: return f"{int(v):,}"
        except: return str(v)

    def pct(a, b):
        if not b: return '-'
        return f"{a/b*100:.1f}%"

    def arrow(cur, prev):
        if cur > prev: return f"▲{c(cur-prev)}"
        if cur < prev: return f"▼{c(prev-cur)}"
        return "→"

    def ach(act, tgt):
        if not tgt: return '-'
        return f"{act/tgt*100:.0f}%"

    # カラーパレット
    DK   = colors.HexColor('#1e293b')
    LT   = colors.HexColor('#f8fafc')
    GD   = colors.HexColor('#cbd5e1')
    GOLD = colors.HexColor('#fef9c3')
    SLV  = colors.HexColor('#f1f5f9')
    BRZ  = colors.HexColor('#fed7aa')
    OK_C = colors.HexColor('#f0fdf4')
    NG_C = colors.HexColor('#fffbeb')

    def ts(extra=None):
        s = [('FONTNAME',(0,0),(-1,-1),FN), ('FONTSIZE',(0,0),(-1,-1),7),
             ('BACKGROUND',(0,0),(-1,0),DK), ('TEXTCOLOR',(0,0),(-1,0),colors.white),
             ('ALIGN',(1,0),(-1,-1),'RIGHT'), ('ALIGN',(0,0),(0,-1),'LEFT'),
             ('GRID',(0,0),(-1,-1),0.4,GD),
             ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,LT]),
             ('TOPPADDING',(0,0),(-1,-1),2), ('BOTTOMPADDING',(0,0),(-1,-1),2)]
        if extra: s.extend(extra)
        return TableStyle(s)

    def p(text, fs=7, bold=False, align=0, color=DK):
        st = ParagraphStyle('_', fontName=FN, fontSize=fs, alignment=align, textColor=color)
        return Paragraph(text, st)

    def sec(text):
        return p(text, fs=11, color=colors.HexColor('#0e7490'))

    # 出力先（デスクトップ）
    pdf_path = os.path.join(os.path.expanduser("~"), "Desktop",
                            f"営業月次レポート_{year}年{month:02d}月.pdf")

    doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4),
                            leftMargin=12*mm, rightMargin=12*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    story = []

    # ═════════════════════════════════════════
    # Page 1: 表紙
    # ═════════════════════════════════════════
    s_cov = ParagraphStyle('cv', fontName=FN, fontSize=22, alignment=1)
    s_sub = ParagraphStyle('cs', fontName=FN, fontSize=14, alignment=1,
                           textColor=colors.HexColor('#475569'))
    s_inf = ParagraphStyle('ci', fontName=FN, fontSize=9,  alignment=1,
                           textColor=colors.HexColor('#64748b'))
    s_toc = ParagraphStyle('ct', fontName=FN, fontSize=9,
                           leftIndent=20*mm, textColor=colors.HexColor('#334155'))

    story += [Spacer(1,30*mm),
              Paragraph("月次営業レポート", s_cov), Spacer(1,6*mm),
              Paragraph(f"{year}年{month}月", s_sub), Spacer(1,4*mm),
              Paragraph(f"作成日: {datetime.date.today().strftime('%Y年%m月%d日')}", s_inf),
              Paragraph(f"対象担当者: {len(members)}名 ｜ 架電合計: {c(team['calls'])} ｜ "
                        f"受注合計: {team['received']}", s_inf),
              Spacer(1,12*mm),
              Paragraph("1. 全体月次サマリー（当月実績・目標・前月比）", s_toc),
              Paragraph("2. 月別推移（過去6ヶ月）",   s_toc),
              Paragraph("3. 個人別成績ランキング",     s_toc),
              Paragraph("4. 時間帯別架電分析",         s_toc),
              Paragraph(f"5. 担当者別詳細（全指標 ×{len(members)}名）", s_toc),
              PageBreak()]

    # ═════════════════════════════════════════
    # Page 2: 全体月次サマリー
    # ═════════════════════════════════════════
    story.append(sec(f"1. 全体月次サマリー　{year}年{month}月"))
    story.append(Spacer(1,3*mm))

    # アラートライン
    for key, label in [('received','受注'), ('prospect','見込')]:
        tgt = tgt_team.get(key, 0)
        act = team[key]
        if tgt > 0:
            if act >= tgt:
                txt = f"◎ {label} 目標達成！　{c(act)} / {c(tgt)}"
                bg = OK_C; fc = colors.HexColor('#166534')
            else:
                txt = f"△ {label} 目標まで残り{c(tgt-act)}件　{c(act)} / {c(tgt)}　（{act/tgt*100:.0f}%）"
                bg = NG_C; fc = colors.HexColor('#92400e')
            s_al = ParagraphStyle('al', fontName=FN, fontSize=8, textColor=fc,
                                  backColor=bg, borderWidth=1, borderPadding=4,
                                  borderColor=colors.HexColor('#f59e0b') if act < tgt else colors.HexColor('#22c55e'),
                                  spaceBefore=1*mm, spaceAfter=1*mm)
            story.append(Paragraph(txt, s_al))

    story.append(Spacer(1,2*mm))

    kpi = [['指標','実績','目標','達成率','前月','前月比'],
           ['架電',    c(team['calls']),    c(tgt_team.get('calls',0)),    ach(team['calls'],    tgt_team.get('calls',0)),    c(prev_team['calls']),    arrow(team['calls'],    prev_team['calls'])],
           ['有効',    c(team['valid']),    c(tgt_team.get('valid',0)),    ach(team['valid'],    tgt_team.get('valid',0)),    c(prev_team['valid']),    arrow(team['valid'],    prev_team['valid'])],
           ['見込',    c(team['prospect']), c(tgt_team.get('prospect',0)), ach(team['prospect'], tgt_team.get('prospect',0)), c(prev_team['prospect']), arrow(team['prospect'], prev_team['prospect'])],
           ['受注',    c(team['received']), c(tgt_team.get('received',0)), ach(team['received'], tgt_team.get('received',0)), c(prev_team['received']), arrow(team['received'], prev_team['received'])],
           ['送客',    c(team['referral']), c(tgt_team.get('referral',0)), ach(team['referral'], tgt_team.get('referral',0)), c(prev_team['referral']), arrow(team['referral'], prev_team['referral'])],
           ['有効率',  pct(team['valid'],    team['calls']),   '-', '-', pct(prev_team['valid'],    prev_team['calls']),   '-'],
           ['受注率',  pct(team['received'], team['calls']),   '-', '-', pct(prev_team['received'], prev_team['calls']),   '-'],
           ['稼働時間',f"{team['hours']:.0f}h", '-', '-', '-', '-'],
           ['生産性',  f"{team['prospect']/team['hours']:.3f}" if team['hours'] else '-', '-', '-', '-', '-']]
    tbl = Table(kpi, colWidths=[32*mm,28*mm,28*mm,22*mm,28*mm,22*mm])
    tbl.setStyle(ts())
    story += [tbl, PageBreak()]

    # ═════════════════════════════════════════
    # Page 3: 月別推移
    # ═════════════════════════════════════════
    story.append(sec("2. 月別推移（過去6ヶ月）"))
    story.append(Spacer(1,3*mm))
    tr = [['年月','架電','有効','有効率','見込','受注','受注率']]
    for t in trend6:
        tr.append([f"{t['year']}/{t['month']:02d}",
                   c(t['calls']), c(t['valid']),  pct(t['valid'],    t['calls']),
                   c(t['prospect']), c(t['received']), pct(t['received'], t['calls'])])
    tbl = Table(tr, colWidths=[32*mm]*7)
    tbl.setStyle(ts())
    story += [tbl, PageBreak()]

    # ═════════════════════════════════════════
    # Page 4: ランキング
    # ═════════════════════════════════════════
    story.append(sec("3. 個人別成績ランキング"))
    story.append(Spacer(1,3*mm))

    def rank_tbl(title, key, lbl, top=5):
        ranked = sorted([(m, mdata.get(m,{})) for m in members],
                        key=lambda x: -x[1].get(key,0))
        data = [['順位','担当者', lbl,'架電','有効','見込']]
        for i, (m, d) in enumerate(ranked[:top], 1):
            data.append([str(i), m, c(d.get(key,0)),
                        c(d.get('calls',0)), c(d.get('valid',0)), c(d.get('prospect',0))])
        extra = [('BACKGROUND',(0,1),(0,1),GOLD),
                 ('BACKGROUND',(0,2),(0,2),SLV),
                 ('BACKGROUND',(0,3),(0,3),BRZ)]
        tbl = Table(data, colWidths=[14*mm,38*mm,24*mm,24*mm,24*mm,24*mm])
        tbl.setStyle(ts(extra))
        return [p(title, fs=9), tbl, Spacer(1,3*mm)]

    # 生産性ランキング（見込÷稼働時間）
    def prod_rank_tbl(top=5):
        ranked = sorted(
            [(m, mdata.get(m,{})) for m in members],
            key=lambda x: -(x[1].get('prospect',0)/x[1].get('hours',1) if x[1].get('hours',0)>0 else 0)
        )
        data = [['順位','担当者','生産性(見込/h)','見込','稼働h','受注']]
        for i, (m, d) in enumerate(ranked[:top], 1):
            prod = f"{d.get('prospect',0)/d.get('hours',1):.3f}" if d.get('hours',0)>0 else '-'
            data.append([str(i), m, prod,
                        c(d.get('prospect',0)), f"{d.get('hours',0):.0f}h", c(d.get('received',0))])
        extra = [('BACKGROUND',(0,1),(0,1),GOLD),
                 ('BACKGROUND',(0,2),(0,2),SLV),
                 ('BACKGROUND',(0,3),(0,3),BRZ)]
        tbl = Table(data, colWidths=[14*mm,38*mm,30*mm,24*mm,24*mm,24*mm])
        tbl.setStyle(ts(extra))
        return [p("生産性ランキング（上位5名）", fs=9), tbl, Spacer(1,3*mm)]

    for item in rank_tbl("受注ランキング（上位5名）", 'received', '受注'):   story.append(item)
    for item in rank_tbl("見込ランキング（上位5名）",  'prospect',  '見込'):  story.append(item)
    for item in rank_tbl("送客ランキング（上位5名）",  'referral',  '送客'):  story.append(item)
    for item in prod_rank_tbl(): story.append(item)
    story.append(PageBreak())

    # ═════════════════════════════════════════
    # Page 5: 時間帯別分析
    # ═════════════════════════════════════════
    story.append(sec("4. 時間帯別架電分析"))
    story.append(Spacer(1,3*mm))

    # チーム合計
    ht = [[''] + [f"{h}時" for h in range(10,19)],
          ['チーム合計'] + [c(hourly.get(h,0)) for h in range(10,19)]]
    tbl = Table(ht, colWidths=[30*mm]+[24*mm]*9)
    tbl.setStyle(ts())
    story += [tbl, Spacer(1,4*mm)]

    # 担当者別時間帯
    mh = [['担当者']+[f"{h}時" for h in range(10,19)]+['合計']]
    for m in members:
        d = mdata.get(m,{})
        mh.append([m]+[c(d.get(f'c{h}',0)) for h in range(10,19)]+[c(d.get('calls',0))])
    cws = [34*mm]+[21*mm]*9+[24*mm]
    # 列幅が広すぎる場合は縮める
    tbl = Table(mh, colWidths=cws)
    tbl.setStyle(ts())
    story += [tbl, PageBreak()]

    # ═════════════════════════════════════════
    # Pages 6+: 担当者別詳細
    # ═════════════════════════════════════════
    story.append(sec("5. 担当者別詳細"))

    for idx, m in enumerate(members):
        d   = mdata.get(m, {})
        tgt = tgt_map.get(m, {})

        story.append(Spacer(1,2*mm))
        story.append(p(f"■ {m}", fs=10))
        story.append(Spacer(1,2*mm))

        kpi_m = [['指標','実績','目標','達成率'],
                 ['架電',   c(d.get('calls',0)),    c(tgt.get('calls',0)),    ach(d.get('calls',0),    tgt.get('calls',0))],
                 ['有効',   c(d.get('valid',0)),    c(tgt.get('valid',0)),    ach(d.get('valid',0),    tgt.get('valid',0))],
                 ['見込',   c(d.get('prospect',0)), c(tgt.get('prospect',0)), ach(d.get('prospect',0), tgt.get('prospect',0))],
                 ['受注',   c(d.get('received',0)), c(tgt.get('received',0)), ach(d.get('received',0), tgt.get('received',0))],
                 ['送客',   c(d.get('referral',0)), c(tgt.get('referral',0)), ach(d.get('referral',0), tgt.get('referral',0))],
                 ['稼働時間', f"{d.get('hours',0):.1f}h", '-', '-'],
                 ['有効率', pct(d.get('valid',0), d.get('calls',0)), '-', '-'],
                 ['受注率', pct(d.get('received',0), d.get('calls',0)), '-', '-']]
        tbl = Table(kpi_m, colWidths=[34*mm,30*mm,30*mm,30*mm])
        tbl.setStyle(ts())
        story += [tbl, Spacer(1,3*mm)]

        hd = [['10時','11時','12時','13時','14時','15時','16時','17時','18時']]
        hd.append([c(d.get(f'c{h}',0)) for h in range(10,19)])
        tbl = Table(hd, colWidths=[27*mm]*9)
        tbl.setStyle(ts())
        story.append(tbl)

        if idx < len(members)-1:
            story.append(PageBreak())

    doc.build(story)
    return pdf_path


if __name__ == "__main__":
    main()
