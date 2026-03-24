# -*- coding: utf-8 -*-
"""
split_fe_be.py
SalesMgr_BE.accdb (テーブル+フォーム+クエリ) を
  SalesMgr_BE.accdb (テーブルのみ) と
  SalesMgr_FE.accdb (フォーム+クエリ+VBA、リンクテーブル) に分離する

手順:
  1. SalesMgr_FE.accdb を作成 (BE をコピー)
  2. FE の実テーブルをリンクテーブルに置換
  3. BE からフォームとクエリを削除・スタートアップをクリア
  4. FE にスタートアップ設定 (F_Main)
"""
import os, shutil, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")

TABLES = ["T_MEMBERS", "T_RECORDS", "T_MEMBER_TARGETS", "T_REFERRALS"]

# 全フォーム（enhance_access.py で追加されたものも含め安全に試みる）
FORMS = [
    "F_Main","F_Members","F_Daily","F_DailyEdit",
    "F_Targets","F_Referrals","F_Report","F_Ranking",
    "F_List","F_Analysis",
]

# 全クエリ（setup_queries.py + fix_queries.py で作成されたもの）
QUERIES = [
    "Q_DailyByMember","Q_MonthlySummary","Q_MonthlyByMember",
    "Q_RankingCalls","Q_RankingValid","Q_RankingReceived",
    "Q_HourlyCalls","Q_Referrals","Q_TargetVsActual",
    "Q_WorkDays","Q_MemberActivity","Q_RefTrend",
]

acForm  = 2
acQuery = 1

def main():
    print("=" * 60)
    print("SalesMgr FE/BE 分離スクリプト")
    print("  BE = テーブルのみ")
    print("  FE = フォーム+クエリ+VBA+リンクテーブル")
    print("=" * 60)

    # ── Step 1: FE を BE からコピーして作成 ──────────────────
    print("\n[Step 1] SalesMgr_FE.accdb を作成 (BE のコピー)...")
    if os.path.exists(FE):
        os.remove(FE)
        time.sleep(0.5)
    shutil.copy2(BE, FE)
    print(f"  コピー完了: {FE}")

    # ── Step 2: FE の実テーブル → リンクテーブルに置換 ──────
    print("\n[Step 2] FE のテーブルをリンクテーブルに置換...")
    app = win32com.client.DispatchEx("Access.Application")
    app.Visible = False
    app.UserControl = False
    app.OpenCurrentDatabase(FE)
    db = app.CurrentDb()

    # 実テーブルを削除
    for tbl in TABLES:
        try:
            db.TableDefs.Delete(tbl)
            print(f"  削除: {tbl}")
        except Exception as e:
            print(f"  削除スキップ ({tbl}): {e}")
    time.sleep(0.3)

    # リンクテーブルを追加（同フォルダの BE を参照）
    for tbl in TABLES:
        try:
            td = db.CreateTableDef(tbl)
            td.Connect = ";DATABASE=" + BE
            td.SourceTableName = tbl
            db.TableDefs.Append(td)
            print(f"  リンク追加: {tbl} -> BE")
        except Exception as e:
            print(f"  リンク追加失敗 ({tbl}): {e}")

    # FE のスタートアップ設定（CurrentDb().Properties 方式）
    try:
        props = app.CurrentDb().Properties
        for pn, pt, pv in [
            ("StartUpForm",        10, "F_Main"),
            ("AppTitle",           10, "SalesMgr 営業管理"),
            ("StartUpShowDBWindow", 1, False),
        ]:
            try:
                props(pn).Value = pv
            except Exception:
                props.Append(app.CurrentDb().CreateProperty(pn, pt, pv))
        print("  FE スタートアップ設定完了 (F_Main)")
    except Exception as e:
        print(f"  スタートアップ設定警告: {e}")

    app.CloseCurrentDatabase()
    app.Quit()
    del app
    time.sleep(2)

    # ── Step 3: BE からフォームとクエリを削除 ────────────────
    print("\n[Step 3] BE からフォームとクエリを削除...")
    app2 = win32com.client.DispatchEx("Access.Application")
    app2.Visible = False
    app2.UserControl = False
    app2.OpenCurrentDatabase(BE)

    for frm in FORMS:
        try:
            app2.DoCmd.Close(acForm, frm)
        except:
            pass
        try:
            app2.DoCmd.DeleteObject(acForm, frm)
            print(f"  フォーム削除: {frm}")
        except Exception as e:
            if "見つかりません" in str(e) or "not found" in str(e).lower() or "-2147352567" in str(e):
                pass  # 存在しないフォームはスキップ
            else:
                print(f"  フォーム削除スキップ ({frm}): {e}")

    for qry in QUERIES:
        try:
            app2.DoCmd.DeleteObject(acQuery, qry)
            print(f"  クエリ削除: {qry}")
        except Exception as e:
            if "見つかりません" in str(e) or "not found" in str(e).lower() or "-2147352567" in str(e):
                pass
            else:
                print(f"  クエリ削除スキップ ({qry}): {e}")

    # BE のスタートアップをクリア（データのみなので起動フォームなし）
    try:
        props2 = app2.CurrentDb().Properties
        for pn, pt, pv in [
            ("StartUpForm",        10, ""),
            ("AppTitle",           10, "SalesMgr BE (Data)"),
            ("StartUpShowDBWindow", 1, True),
        ]:
            try:
                props2(pn).Value = pv
            except Exception:
                props2.Append(app2.CurrentDb().CreateProperty(pn, pt, pv))
        print("  BE スタートアップをクリア")
    except Exception as e:
        print(f"  BE スタートアップクリア警告: {e}")

    app2.CloseCurrentDatabase()
    app2.Quit()
    del app2
    time.sleep(1)

    print("\n" + "=" * 60)
    print("分離完了！")
    print(f"  SalesMgr_FE.accdb  ... フォーム+クエリ+VBA（ユーザーが開く）")
    print(f"  SalesMgr_BE.accdb  ... テーブルのみ（データ格納）")
    print()
    print("  ※ 同じフォルダに置いて FE を開けばリンクテーブルが自動接続")
    print("  ※ ネットワーク共有後は FE でリンクテーブルマネージャーから再リンク")
    print("=" * 60)

if __name__ == "__main__":
    main()
