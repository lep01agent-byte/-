# -*- coding: utf-8 -*-
"""残りFAIL調査"""
import os, time, win32com.client

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
BE = os.path.join(FOLDER, "SalesMgr_BE.accdb")
FE = os.path.join(FOLDER, "SalesMgr_FE.accdb")

os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
time.sleep(2)

# ── A15/A16/A25: データ調査 ──────────────────────────────
print("=" * 70)
print("A. データ整合性調査")
print("=" * 70)

engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(BE)

# A15: calls vs hourly sum の不一致サンプル
print("\n[A15] calls ≠ hourly合計 (最初の10件)")
sql = """SELECT rec_date, member_name, calls,
  IIf(IsNull(calls_10),0,calls_10)+IIf(IsNull(calls_11),0,calls_11)+
  IIf(IsNull(calls_12),0,calls_12)+IIf(IsNull(calls_13),0,calls_13)+
  IIf(IsNull(calls_14),0,calls_14)+IIf(IsNull(calls_15),0,calls_15)+
  IIf(IsNull(calls_16),0,calls_16)+IIf(IsNull(calls_17),0,calls_17)+
  IIf(IsNull(calls_18),0,calls_18) AS hourly_sum,
  calls_10, calls_11, calls_12, calls_13
FROM T_RECORDS
WHERE calls <> (IIf(IsNull(calls_10),0,calls_10)+IIf(IsNull(calls_11),0,calls_11)+
  IIf(IsNull(calls_12),0,calls_12)+IIf(IsNull(calls_13),0,calls_13)+
  IIf(IsNull(calls_14),0,calls_14)+IIf(IsNull(calls_15),0,calls_15)+
  IIf(IsNull(calls_16),0,calls_16)+IIf(IsNull(calls_17),0,calls_17)+
  IIf(IsNull(calls_18),0,calls_18))
ORDER BY rec_date DESC"""
rs = db.OpenRecordset(sql)
cnt = 0
while not rs.EOF and cnt < 10:
    print(f"  {str(rs.Fields('rec_date').Value)[:10]} {rs.Fields('member_name').Value}: calls={rs.Fields('calls').Value}, hourly_sum={rs.Fields('hourly_sum').Value}, c10={rs.Fields('calls_10').Value}")
    rs.MoveNext(); cnt += 1
rs.Close()

# A15: 時間別が全部NULLの件数
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE calls_10 IS NULL AND calls_11 IS NULL AND calls_12 IS NULL AND calls_13 IS NULL AND calls_14 IS NULL AND calls_15 IS NULL AND calls_16 IS NULL AND calls_17 IS NULL AND calls_18 IS NULL")
null_hourly = rs.Fields(0).Value; rs.Close()
print(f"  時間別全NULL件数: {null_hourly}")
rs = db.OpenRecordset("SELECT COUNT(*) FROM T_RECORDS WHERE calls_10 IS NULL AND calls_11 IS NULL AND calls_12 IS NULL AND calls_13 IS NULL AND calls_14 IS NULL AND calls_15 IS NULL AND calls_16 IS NULL AND calls_17 IS NULL AND calls_18 IS NULL AND calls > 0")
null_hourly_nonzero = rs.Fields(0).Value; rs.Close()
print(f"  時間別全NULL かつ calls>0: {null_hourly_nonzero}")

# A16: orphan member_names in T_RECORDS
print("\n[A16] T_RECORDS orphan member_names")
rs = db.OpenRecordset("SELECT DISTINCT R.member_name, COUNT(*) AS cnt FROM T_RECORDS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL GROUP BY R.member_name")
while not rs.EOF:
    print(f"  '{rs.Fields('member_name').Value}': {rs.Fields('cnt').Value}件")
    rs.MoveNext()
rs.Close()

# A25: orphan member_names in T_REFERRALS
print("\n[A25] T_REFERRALS orphan member_names")
rs = db.OpenRecordset("SELECT DISTINCT R.member_name, COUNT(*) AS cnt FROM T_REFERRALS AS R LEFT JOIN T_MEMBERS AS M ON R.member_name=M.member_name WHERE M.ID IS NULL GROUP BY R.member_name")
while not rs.EOF:
    print(f"  '{rs.Fields('member_name').Value}': {rs.Fields('cnt').Value}件")
    rs.MoveNext()
rs.Close()

# T_MEMBERSの全メンバー
print("\n[T_MEMBERS 全メンバー]")
rs = db.OpenRecordset("SELECT member_name, active FROM T_MEMBERS ORDER BY member_name")
while not rs.EOF:
    print(f"  {rs.Fields('member_name').Value}: active={rs.Fields('active').Value}")
    rs.MoveNext()
rs.Close()

db.Close()

# ── B41/C51: junkオブジェクトのVBA参照確認 ───────────────
print("\n" + "=" * 70)
print("B. junkオブジェクトのVBA参照確認")
print("=" * 70)

app = win32com.client.DispatchEx("Access.Application")
app.Visible = False
app.OpenCurrentDatabase(FE, False)
time.sleep(2)

JUNK_QUERIES = ["Q_AllMembers","Q_Records_List","Q_Targets_Monthly","Q_Referrals_Monthly","Q_List_Summary"]
JUNK_FORMS = ["F_QueryBrowser"]

# 全フォームのVBAを取得
vba_all = {}
db2 = app.CurrentDb()
form_names = set()
try:
    docs = db2.Containers("Forms").Documents
    for i in range(docs.Count):
        form_names.add(docs(i).Name)
except: pass

print(f"\nFEフォーム: {sorted(form_names)}")

for fn in form_names:
    try:
        app.DoCmd.OpenForm(fn, 1)  # Design view
        time.sleep(0.3)
        frm = app.Forms(fn)
        if frm.HasModule:
            try:
                comp = app.VBE.VBProjects(1).VBComponents("Form_" + fn)
                lines = comp.CodeModule.CountOfLines
                if lines > 0:
                    vba_all[fn] = comp.CodeModule.Lines(1, lines)
            except Exception as ve:
                print(f"  VBE取得失敗 {fn}: {ve}")
        app.DoCmd.Close(2, fn)
        time.sleep(0.2)
    except Exception as e:
        print(f"  フォームオープン失敗 {fn}: {e}")

print(f"\nVBA取得済みフォーム: {list(vba_all.keys())}")

# junkクエリの参照確認
print("\n[junkクエリのVBA参照]")
for jq in JUNK_QUERIES:
    refs = [fn for fn, vba in vba_all.items() if jq in vba]
    print(f"  {jq}: 参照={refs if refs else '参照なし'}")

# junkフォームの参照確認
print("\n[junkフォームのVBA参照]")
for jf in JUNK_FORMS:
    refs = [fn for fn, vba in vba_all.items() if jf in vba]
    print(f"  {jf}: 参照={refs if refs else '参照なし'}")

app.CloseCurrentDatabase()
app.Quit()
del app
