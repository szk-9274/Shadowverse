# pip install pyautogui pywin32
import time
import random
import ctypes
import pyautogui
import win32gui
import win32con

GAME_KEYWORD = "Shadowverse"  # ウィンドウタイトルに含まれる文字列

# --- クリック座標（スクリーン座標） ---
POS_HOME              = (956, 991)   # ホーム画面
POS_TO_RANKED         = (1499, 675)  # ホーム→ランクマッチ画面
POS_BATTLE_BUTTON     = (1716, 673)  # バトルボタン

def set_dpi_awareness():
    """DPIスケーリングのズレを防ぐ"""
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # Per-monitor v2
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def find_shadowverse_window():
    """タイトルにGAME_KEYWORDを含む最大面積の可視ウィンドウを返す"""
    cands = []
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if title and GAME_KEYWORD.lower() in title.lower():
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            cands.append(((r - l) * (b - t), hwnd))
    win32gui.EnumWindows(enum_handler, None)
    if not cands:
        return None
    cands.sort(reverse=True)
    return cands[0][1]

def bring_to_front(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    for _ in range(4):
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.05)
        except Exception:
            time.sleep(0.05)

def click_screen(x, y, jitter=2, move_duration=(0.05, 0.12), post_sleep=(0.25, 0.45)):
    """スクリーン座標をそのままクリック。微小ジッタとランダム待機を挟む"""
    jx = x + random.randint(-jitter, jitter)
    jy = y + random.randint(-jitter, jitter)
    pyautogui.moveTo(jx, jy, duration=random.uniform(*move_duration))
    pyautogui.click()
    time.sleep(random.uniform(*post_sleep))

def start_ranked_match():
    """
    ホーム→ランクマッチ→バトル開始 まで自動でクリックする
    事前にShadowverseを起動して、該当画面が見える状態で実行してください
    """
    set_dpi_awareness()
    hwnd = find_shadowverse_window()
    if not hwnd:
        raise RuntimeError("Shadowverseのウィンドウが見つかりません。起動中か確認してください。")

    bring_to_front(hwnd)
    time.sleep(0.2)

    # 1) ホーム画面のボタン
    click_screen(*POS_HOME)

    # 2) ランクマッチ画面へ
    click_screen(*POS_TO_RANKED)

    # 3) バトル開始ボタン
    click_screen(*POS_BATTLE_BUTTON)

    # 必要ならマッチング確定用の追クリックを入れる（例：OKなど）
    # click_screen(XXXXX, YYYYY)

if __name__ == "__main__":
    start_ranked_match()
    print("ランクマッチ開始までのクリックを実行しました。")
