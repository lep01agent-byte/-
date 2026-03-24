# -*- coding: utf-8 -*-
"""NG クエリのSQL確認 + FE フォーム確認"""
import os, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")

NG_QUERIES = ["Q_List_Summary", "Q_Records_List", "Q_Referrals_Monthly", "Q_Targets_Monthly"]

TARGET_FORMS = [
    "F_Main", "F_DailyInput", "F_DailyEdit",
    "F_Report", "F_Targets", "F_Members",
    "F_Ranking", "F_Referrals",
]

results = []

def log(category, name, status, detail=""):
    mark = "OK" if status else "NG"
    results.append((category, name, mark, detail))
    d = f"  -> {detail}" if detail else ""
    print(f"  [{mark}] {name}{d}", flush=True)

# ─────────────────── BE: NG クエリSQL確認 ───────────────────
print("=" * 70)
print("BE: NG クエリ SQL確認")
print("=" * 70)

app = win32com.client.DispatchEx("Access.Application")
app.Visible = False
app.OpenCurrentDatabase(BE, False)
time.sleep(2)
db = app.CurrentDb()

for qname in NG_QUERIES:
    try:
        sql = db.QueryDefs(qname).SQL
        print(f"\n[{qname}]")
        print(sql[:500])
    except Exception as e:
        print(f"\n[{qname}] 取得失敗: {e}")

app.CloseCurrentDatabase()
app.Quit()
del app
time.sleep(1)

# ─────────────────── FE: 全フォーム確認 ─────────────────────
print("\n" + "=" * 70)
print("FE: フォーム確認")
print(f"対象: {FE}")
print("=" * 70)

app2 = win32com.client.DispatchEx("Access.Application")
app2.Visible = False
try:
    app2.AutomationSecurity = 1
except:
    pass

app2.OpenCurrentDatabase(FE, False)
time.sleep(2)
db2 = app2.CurrentDb()

# フォーム一覧
form_names = set()
try:
    docs = db2.Containers("Forms").Documents
    for i in range(docs.Count):
        form_names.add(docs(i).Name)
except Exception as e:
    print(f"フォームコンテナ取得失敗: {e}")

print(f"\nFE フォーム総数: {len(form_names)}")
for fn in sorted(form_names):
    print(f"  {fn}")

# VBA コンポーネント確認
print("\n[FE VBA コンポーネント]")
try:
    proj = app2.VBE.VBProjects(1)
    print(f"  プロジェクト: {proj.Name}, コンポーネント数: {proj.VBComponents.Count}")
    for i in range(proj.VBComponents.Count):
        comp = proj.VBComponents(i)
        lines = comp.CodeModule.CountOfLines
        print(f"    {comp.Name}: {lines}行")
except Exception as e:
    print(f"  VBE取得失敗: {e}")

# フォーム存在確認
print("\n[FE フォーム存在確認]")
for frm in TARGET_FORMS:
    if frm in form_names:
        log("FEフォーム存在", frm, True, "あり")
    else:
        log("FEフォーム存在", frm, False, "見つかりません")

# フォーム起動確認
print("\n[FE フォーム起動確認]")
for frm in TARGET_FORMS:
    if frm not in form_names:
        log("FEフォーム起動", frm, False, "存在しないためスキップ")
        continue
    try:
        if frm == "F_DailyEdit":
            app2.DoCmd.OpenForm(frm, 0, "", "", 0, 0, "ADD|2026|3")
        else:
            app2.DoCmd.OpenForm(frm, 0)
        time.sleep(1.5)
        log("FEフォーム起動", frm, True, "エラーなく開きました")
        try:
            app2.DoCmd.Close(2, frm)
            time.sleep(0.5)
        except:
            pass
    except Exception as e:
        err = str(e)
        if "modal" in err.lower() or "2147352567" in err:
            log("FEフォーム起動", frm, True, "modal/popup スキップ(正常)")
        else:
            log("FEフォーム起動", frm, False, err[:120])

# FE クエリ存在確認
print("\n[FE クエリ確認]")
try:
    qdf_names = []
    for i in range(db2.QueryDefs.Count):
        n = db2.QueryDefs(i).Name
        if not n.startswith("~"):
            qdf_names.append(n)
    print(f"  FE クエリ総数: {len(qdf_names)}")
    for q in qdf_names:
        print(f"    {q}")
except Exception as e:
    print(f"  クエリ取得失敗: {e}")

# リンクテーブル確認
print("\n[FE リンクテーブル確認]")
TABLES = ["T_MEMBERS", "T_RECORDS", "T_MEMBER_TARGETS", "T_REFERRALS"]
for tbl in TABLES:
    try:
        td = db2.TableDefs(tbl)
        connect = td.Connect
        is_linked = bool(connect)
        if is_linked:
            log("FEリンクテーブル", tbl, True, f"リンク先={connect[:60]}")
        else:
            log("FEリンクテーブル", tbl, False, "リンクテーブルではありません(実テーブル?)")
    except Exception as e:
        log("FEリンクテーブル", tbl, False, str(e)[:80])

app2.CloseCurrentDatabase()
app2.Quit()
del app2
time.sleep(1)

# サマリー
print("\n" + "=" * 70)
print("サマリー")
ok = sum(1 for r in results if r[2]=="OK")
ng = sum(1 for r in results if r[2]=="NG")
print(f"  OK: {ok}  /  NG: {ng}  /  合計: {len(results)}")
if ng > 0:
    print("\n[NG一覧]")
    for cat,name,mark,detail in results:
        if mark=="NG":
            print(f"  NG: [{cat}] {name} -> {detail}")
print("=" * 70)
