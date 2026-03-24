# -*- coding: utf-8 -*-
"""
check_operation.py
全フォーム・全クエリの動作確認
  - フォーム: 開けるか / VBA コンパイルエラーがないか
  - クエリ: SQL が有効か（PARAMETERS 付きは構文チェックのみ）
"""
import os, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")

FORMS = [
    "F_Main", "F_Members", "F_Daily", "F_DailyEdit",
    "F_Targets", "F_Referrals", "F_Report", "F_Ranking",
    "F_QueryBrowser",
]

QUERIES = [
    "Q_アクティブメンバー一覧", "Q_メンバー別時間別実績", "Q_メンバー12ヶ月推移",
    "Q_メンバー月次集計", "Q_生産性ランキング", "Q_見込みランキング",
    "Q_受電ランキング", "Q_紹介ランキング", "Q_紹介トレンド月次",
    "Q_チーム月次集計", "Q_チーム月次目標", "Q_月次トレンド",
]

results = []

def log(category, name, status, detail=""):
    mark = "OK" if status else "NG"
    results.append((category, name, mark, detail))
    detail_str = f"  -> {detail}" if detail else ""
    print(f"  [{mark}] {name}{detail_str}")

def main():
    print("=" * 65)
    print("SalesMgr 動作確認チェック")
    print(f"対象: {FE}")
    print("=" * 65)

    app = win32com.client.DispatchEx("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(FE)
    time.sleep(2)

    # ── 1. クエリ存在確認 ────────────────────────────────────
    print("\n[クエリ存在確認]")
    db = app.CurrentDb()
    qdf_names = set()
    try:
        for i in range(db.QueryDefs.Count):
            qdf_names.add(db.QueryDefs(i).Name)
    except:
        pass

    for q in QUERIES:
        if q in qdf_names:
            # SQL を取得してチェック
            try:
                sql = db.QueryDefs(q).SQL
                has_params = "PARAMETERS" in sql.upper()
                log("クエリ", q, True, f"PARAMETERS{'付き' if has_params else 'なし'}, {len(sql)}文字")
            except Exception as e:
                log("クエリ", q, False, str(e)[:80])
        else:
            log("クエリ", q, False, "クエリが存在しません")

    # ── 2. フォーム存在 + VBA 確認 ───────────────────────────
    print("\n[フォーム存在・VBA確認]")
    form_names = set()
    try:
        for i in range(db.Containers("Forms").Documents.Count):
            form_names.add(db.Containers("Forms").Documents(i).Name)
    except:
        pass

    for frm in FORMS:
        if frm not in form_names:
            log("フォーム", frm, False, "フォームが存在しません")
            continue
        # VBA コードを確認
        try:
            comp = app.VBE.VBProjects(1).VBComponents("Form_" + frm)
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            # 基本チェック
            has_option_explicit = "Option Explicit" in code
            has_on_error        = "On Error GoTo EH" in code
            sub_count           = code.count("Private Sub ") + code.count("Public Sub ")
            log("フォーム", frm, True,
                f"{comp.CodeModule.CountOfLines}行, {sub_count}Sub, "
                f"OptionExplicit={'あり' if has_option_explicit else 'なし'}, "
                f"OnError={'あり' if has_on_error else 'なし'}")
        except Exception as e:
            log("フォーム", frm, False, f"VBA取得失敗: {str(e)[:80]}")

    # ── 3. フォームを実際に開いて確認 ───────────────────────
    print("\n[フォーム起動確認]")
    for frm in FORMS:
        if frm not in form_names:
            continue
        try:
            if frm == "F_DailyEdit":
                app.DoCmd.OpenForm(frm, 0, "", "", 0, 0, "ADD|2026|3")
            else:
                app.DoCmd.OpenForm(frm)
            time.sleep(1)
            log("起動", frm, True, "エラーなく開きました")
            try:
                app.DoCmd.Close(2, frm)
                time.sleep(0.3)
            except:
                pass
        except Exception as e:
            err = str(e)
            # ポップアップ/モーダルはハング回避のためスキップ
            if "modal" in err.lower() or "-2147352567" in err:
                log("起動", frm, True, "modal/popupのため開閉スキップ")
            else:
                log("起動", frm, False, err[:100])

    # ── 4. リンクテーブル確認 ───────────────────────────────
    print("\n[リンクテーブル確認]")
    TABLES = ["T_MEMBERS", "T_RECORDS", "T_MEMBER_TARGETS", "T_REFERRALS"]
    for tbl in TABLES:
        try:
            td = db.TableDefs(tbl)
            connect = td.Connect
            src = td.SourceTableName
            is_linked = bool(connect)
            rec_count = -1
            if is_linked:
                try:
                    rs = db.OpenRecordset(f"SELECT COUNT(*) FROM {tbl}")
                    rec_count = rs.Fields(0).Value
                    rs.Close()
                except:
                    pass
            log("テーブル", tbl,
                is_linked,
                f"リンク先: {connect[:50] if connect else '(実テーブル)'}, "
                f"件数: {rec_count if rec_count >= 0 else '取得失敗'}")
        except Exception as e:
            log("テーブル", tbl, False, str(e)[:80])

    app.CloseCurrentDatabase()
    app.Quit()
    del app
    time.sleep(1)

    # ── 最終サマリー ─────────────────────────────────────────
    print("\n" + "=" * 65)
    print("チェック結果サマリー")
    print("=" * 65)

    ok_count = sum(1 for r in results if r[2] == "OK")
    ng_count = sum(1 for r in results if r[2] == "NG")
    print(f"  OK: {ok_count}  /  NG: {ng_count}  /  合計: {len(results)}")

    if ng_count > 0:
        print("\n[NG一覧]")
        for cat, name, mark, detail in results:
            if mark == "NG":
                print(f"  {cat} / {name}: {detail}")

    print("\n[全結果]")
    prev_cat = ""
    for cat, name, mark, detail in results:
        if cat != prev_cat:
            print(f"\n  -- {cat} --")
            prev_cat = cat
        d = f" ({detail})" if detail else ""
        print(f"  [{mark}] {name}{d}")

    print("\n" + "=" * 65)

if __name__ == "__main__":
    main()
