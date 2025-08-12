# pip install pynput pywin32
import time, json, ctypes, sys, threading
from datetime import datetime
from pynput import mouse, keyboard
import win32gui, win32con

GAME_KEYWORD = "Shadowverse"
OUTPUT_FILE = "sv_clicks.json"

# --- DPI対策 ---
def set_dpi_awareness():
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # Per-monitor v2
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def list_candidate_windows(keyword):
    cands = []
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if title and keyword.lower() in title.lower():
            try:
                l,t,r,b = win32gui.GetWindowRect(hwnd)
                area = max(0, r-l) * max(0, b-t)
                cands.append((area, hwnd))
            except Exception:
                pass
    win32gui.EnumWindows(enum_handler, None)
    cands.sort(reverse=True, key=lambda x: x[0])
    return [h for _,h in cands]

def bring_to_front(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    for _ in range(5):
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.05)
        except Exception:
            time.sleep(0.05)

def get_client_origin_screen(hwnd):
    # クライアント(0,0)のスクリーン座標
    return win32gui.ClientToScreen(hwnd, (0, 0))

def get_client_size(hwnd):
    l,t,r,b = win32gui.GetClientRect(hwnd)
    return (r - l, b - t)

def point_in_client(hwnd, sx, sy):
    ox, oy = get_client_origin_screen(hwnd)
    w, h = get_client_size(hwnd)
    return (ox <= sx < ox + w) and (oy <= sy < oy + h)

def screen_to_client(hwnd, sx, sy):
    ox, oy = get_client_origin_screen(hwnd)
    return (sx - ox, sy - oy)

def print_target_info(hwnd):
    l,t,r,b = win32gui.GetWindowRect(hwnd)
    ox, oy = get_client_origin_screen(hwnd)
    w,h = get_client_size(hwnd)
    print(f"[Target] hwnd={hwnd} title='{win32gui.GetWindowText(hwnd)}'")
    print(f"  WindowRect: ({l},{t},{r},{b})")
    print(f"  ClientOrigin(screen): ({ox},{oy}), ClientSize: {w}x{h}")

def save_json(records):
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"[Saved] {OUTPUT_FILE} に {len(records)} 件保存しました。")

def main():
    set_dpi_awareness()
    cands = list_candidate_windows(GAME_KEYWORD)
    if not cands:
        print("Shadowverseウィンドウが見つかりません。起動中か、タイトルに'Shadowverse'が含まれているか確認してください。")
        sys.exit(1)

    hwnd = cands[0]  # 面積最大を採用
    bring_to_front(hwnd)
    time.sleep(0.2)
    print_target_info(hwnd)
    print("操作方法: 左クリック=記録 / F10=保存 / Esc=終了\n")

    records = []
    step_counter = 1
    lock = threading.Lock()

    def on_click(sx, sy, button, pressed):
        nonlocal step_counter
        if not pressed:
            return  # 押した瞬間だけ拾う
        if button.name != "left":
            return
        try:
            if point_in_client(hwnd, sx, sy):
                cx, cy = screen_to_client(hwnd, sx, sy)
                with lock:
                    rec = {
                        "step": step_counter,
                        "client_x": cx,
                        "client_y": cy,
                        "screen_x": sx,
                        "screen_y": sy,
                        "ts": datetime.now().isoformat(timespec="seconds")
                    }
                    records.append(rec)
                    step_counter += 1
                print(f"[REC] step={rec['step']} client=({cx},{cy}) screen=({sx},{sy})")
            else:
                print(f"[SKIP] ウィンドウ外のクリック screen=({sx},{sy})")
        except Exception as e:
            print("on_click error:", e)

    def on_press(key):
        try:
            if key == keyboard.Key.esc:
                # 終了時に自動保存
                with lock:
                    if records:
                        save_json(records)
                print("終了します。")
                return False
            elif key == keyboard.Key.f10:
                with lock:
                    save_json(records)
        except Exception as e:
            print("on_press error:", e)

    with mouse.Listener(on_click=on_click) as ml, keyboard.Listener(on_press=on_press) as kl:
        ml.join()
        kl.stop()

if __name__ == "__main__":
    main()
