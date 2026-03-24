# -*- coding: utf-8 -*-
"""DAO参照設定を追加 + VBAからDAO型宣言を削除（Late Binding方式に変更）"""
import os, time, win32com.client

BE = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr", "SalesMgr_BE.accdb")

os.system("taskkill /F /IM MSACCESS.EXE 2>nul")
time.sleep(2)

app = win32com.client.Dispatch("Access.Application")
app.Visible = False
app.UserControl = False
app.OpenCurrentDatabase(BE)
time.sleep(2)

# 方法1: DAO参照を追加
proj = app.VBE.VBProjects(1)
print("Current references:")
for i in range(1, proj.References.Count + 1):
    ref = proj.References(i)
    print(f"  {ref.Name}: {ref.FullPath}")

# DAO 3.6 を追加
dao_path = r"C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
dao_path2 = r"C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll"
# Access 2016+ uses Microsoft Office xx.0 Access database engine
acedao_path = r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL"

added = False
for path in [acedao_path, dao_path, dao_path2]:
    if os.path.exists(path):
        try:
            proj.References.AddFromFile(path)
            print(f"\nAdded DAO ref: {path}")
            added = True
            break
        except Exception as e:
            if "already" in str(e).lower() or "1032813" in str(e):
                print(f"\nDAO ref already exists: {path}")
                added = True
                break
            print(f"  Failed: {e}")

if not added:
    # Try by GUID
    try:
        # Microsoft DAO 3.6 GUID
        proj.References.AddFromGuid("{00025E01-0000-0000-C000-000000000046}", 5, 0)
        print("\nAdded DAO ref by GUID")
        added = True
    except Exception as e:
        if "already" in str(e).lower():
            print("\nDAO ref already exists (GUID)")
            added = True
        else:
            print(f"  GUID failed: {e}")

if not added:
    # Try Microsoft Office Access database engine
    try:
        proj.References.AddFromGuid("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", 0, 0)
        print("\nAdded ACEDAO ref by GUID")
        added = True
    except Exception as e:
        if "already" in str(e).lower():
            print("\nACEDAO ref already exists")
            added = True
        else:
            print(f"  ACEDAO GUID failed: {e}")

print("\nReferences after fix:")
for i in range(1, proj.References.Count + 1):
    ref = proj.References(i)
    print(f"  {ref.Name}: {ref.FullPath}")

time.sleep(1)
app.CloseCurrentDatabase()
time.sleep(1)
app.Quit()
time.sleep(1)
print("\nDone!")
