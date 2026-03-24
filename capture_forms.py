# -*- coding: utf-8 -*-
"""
capture_forms.py  --  Access フォームをスクリーンショット撮影
PrintWindow API を使用（非表示ウィンドウも撮影可能）
"""
import os, time, ctypes, win32com.client, win32gui, win32con, win32ui
from PIL import Image

FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "SalesMgr")
FE     = os.path.join(FOLDER, "SalesMgr_FE.accdb")
OUT    = os.path.join(FOLDER, "screenshots")
os.makedirs(OUT, exist_ok=True)

FORMS = [
    "F_Main",
    "F_Report",
    "F_Targets",
    "F_Daily",
    "F_Members",
    "F_Ranking",
    "F_Referrals",
    "F_DailyEdit",
]

def find_access_hwnd():
    """Access メインウィンドウの HWND を取得"""
    result = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if "OMain" in cls or "OForm" in cls or cls == "OMain" \
               or "Microsoft Access" in win32gui.GetWindowText(hwnd):
                result.append(hwnd)
    win32gui.EnumWindows(cb, None)
    # Access のメインウィンドウクラスは "OMain"
    for hwnd in result:
        if win32gui.GetClassName(hwnd) == "OMain":
            return hwnd
    return result[0] if result else None

def capture_hwnd(hwnd, path):
    """PrintWindow API でウィンドウをキャプチャ"""
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bottom - top
        if w <= 0 or h <= 0:
            return False

        hwnd_dc  = win32gui.GetWindowDC(hwnd)
        mfc_dc   = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc  = mfc_dc.CreateCompatibleDC()
        bmp      = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)

        # PrintWindow: PW_RENDERFULLCONTENT=2, PW_CLIENTONLY=1
        ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

        bmp_info = bmp.GetInfo()
        bmp_str  = bmp.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmp_info['bmWidth'], bmp_info['bmHeight']),
            bmp_str, 'raw', 'BGRX', 0, 1
        )
        img.save(path)

        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        win32gui.DeleteObject(bmp.GetHandle())
        return True
    except Exception as e:
        print(f"    PrintWindow失敗: {e}")
        return False

def main():
    print("=" * 60)
    print("SalesMgr フォームスクリーンショット撮影")
    print(f"保存先: {OUT}")
    print("=" * 60)

    app = win32com.client.DispatchEx("Access.Application")
    app.Visible = True
    app.UserControl = True
    app.OpenCurrentDatabase(FE)
    time.sleep(3)

    # Access ウィンドウを最大化・前面表示
    hwnd = find_access_hwnd()
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(1)
        print(f"  Accessウィンドウ HWND: {hwnd}")
    else:
        print("  警告: Accessウィンドウが見つかりません（撮影は続行）")

    saved = []

    for form_name in FORMS:
        print(f"\n  [{form_name}]...")
        try:
            # F_DailyEdit は popup/modal なので引数付きで開く
            if form_name == "F_DailyEdit":
                app.DoCmd.OpenForm(form_name, 0, "", "", "", 0, "ADD|2026|3")
            else:
                app.DoCmd.OpenForm(form_name)
            time.sleep(2)

            # 最新の Access ウィンドウハンドルを取得（フォームで変わる場合あり）
            hwnd2 = find_access_hwnd() or hwnd

            if hwnd2:
                win32gui.SetForegroundWindow(hwnd2)
                time.sleep(0.5)
                png_path = os.path.join(OUT, f"{form_name}.png")
                if capture_hwnd(hwnd2, png_path):
                    print(f"  保存: {png_path}")
                    saved.append(png_path)
                else:
                    print(f"  キャプチャ失敗")
            else:
                print(f"  ウィンドウなし、スキップ")

            # フォームを閉じる
            try:
                app.DoCmd.Close(2, form_name)
                time.sleep(0.5)
            except:
                pass

        except Exception as e:
            print(f"  フォームオープンエラー: {e}")

    app.Quit()

    print("\n" + "=" * 60)
    print(f"完了: {len(saved)}/{len(FORMS)} 枚保存")
    for p in saved:
        print(f"  {p}")
    print("=" * 60)

if __name__ == "__main__":
    main()
