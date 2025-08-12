# pip install pyautogui pillow pytesseract
import time, os, ctypes, random, re, unicodedata
from datetime import datetime
from pathlib import Path

import pyautogui
from PIL import Image, ImageOps, ImageFilter
import pytesseract
from pytesseract import Output

# ====== Tesseract のパス直指定（あなたの環境に合わせて）======
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ====== クリック座標 ======
CLICK_POS     = (1678, 95)   # フォーカス用に最初に1回クリック
PAGE_NEXT_POS = (1866, 543)  # ページ送り

# ====== 保存先 ======
OUTPUT_DIR = Path(r"C:\00_mycode\Shadowverse\assets\cards")

# ====== 画像保存用の矩形（4隅：スクリーン座標）=====
IMG_PTS = {
    "rb": (831, 1010),  # 右下
    "lb": (202, 1010),  # 左下
    "lt": (202, 185),   # 左上
    "rt": (831, 188),   # 右上
}

# ====== OCR（名前帯）用の矩形（4隅：スクリーン座標）=====
NAME_PTS = {
    "ru": (1814, 219),  # 右上
    "rb": (1814, 279),  # 右下
    "lu": (904,  224),  # 左上
    "lb": (904,  283),  # 左下
}

# ====== 動作パラメータ ======
PAGE_WAIT = 0.40           # ページ送り後の待機
MOVE_DUR = (0.05, 0.12)    # マウス移動の所要時間
JITTER   = 2               # クリック時の微小ジッタ
MAX_LOOPS = 500            # 念のための上限

# ================= DPI対策 =================
def set_dpi_awareness():
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # Per-monitor v2
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# ================= 座標/クリック系 =================
def rect_from_points(points_dict):
    xs = [p[0] for p in points_dict.values()]
    ys = [p[1] for p in points_dict.values()]
    left, top = min(xs), min(ys)
    right, bottom = max(xs), max(ys)
    return left, top, right - left, bottom - top  # (x, y, w, h)

def click_screen(pt, jitter=JITTER, move_dur=MOVE_DUR):
    x, y = pt
    if jitter:
        x += random.randint(-jitter, jitter)
        y += random.randint(-jitter, jitter)
    pyautogui.moveTo(x, y, duration=random.uniform(*move_dur))
    pyautogui.click()

# ================= 日本語ファイル名整形 =================
def sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"\s+", " ", name)
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:80] if len(name) > 80 else name

def normalize_japanese_name(s: str) -> str:
    # 全角/半角・合成文字の正規化
    s = unicodedata.normalize("NFKC", s)
    # 全角スペース→半角、連続空白圧縮
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # 日本語文字（ひら/カナ/漢字/長音符）に挟まれた空白を削除
    jp = r"\u3040-\u30FF\u4E00-\u9FFF\u30FC"
    s = re.sub(fr"(?<=[{jp}])\s+(?=[{jp}])", "", s)
    # 「・」の周囲の余計な空白を整理
    s = re.sub(r"\s*・\s*", "・", s)
    return s

def load_existing_names(directory: Path) -> set:
    names = set()
    if directory.exists():
        for p in directory.glob("*.png"):
            names.add(p.stem.lower())
    return names

# ================= 日本語OCR 強化版 =================
def _preprocess_variants(img: Image.Image):
    """名前帯向けの前処理候補"""
    g = ImageOps.grayscale(img)

    # 1) 自動コントラスト
    yield ImageOps.autocontrast(g)

    # 2) オートコントラスト + 軽シャープ
    yield ImageOps.autocontrast(g).filter(ImageFilter.UnsharpMask(radius=1, percent=160, threshold=3))

    # 3) メディアン平滑→二値化（背景模様の除去）
    s = g.filter(ImageFilter.MedianFilter(size=3))
    yield s.point(lambda p: 255 if p > 155 else 0, mode="1")

    # 4) 2倍拡大（小さい文字に有効）
    yield ImageOps.autocontrast(g).resize((g.width*2, g.height*2), Image.LANCZOS)

def _ocr_once(pil, config, lang="jpn"):
    data = pytesseract.image_to_data(pil, lang=lang, config=config, output_type=Output.DICT)
    words = [w for w in data["text"] if w.strip()]
    confs = [int(c) for c in data["conf"] if c not in (-1, "-1")]
    text = " ".join(words).strip()
    avg_conf = (sum(confs)/len(confs)) if confs else 0
    return text, avg_conf

def ocr_card_name_smart(name_img: Image.Image) -> str:
    """
    日本語（漢字・カタカナ）特化の強化OCR。
    単一行想定（psm=7）を主、必要に応じて psm=6 も試す。
    preserve_interword_spaces=0 で空白を抑制。
    """
    cfgs = [
        "--oem 1 --psm 7 -c preserve_interword_spaces=0",
        "--oem 3 --psm 7 -c preserve_interword_spaces=0",
        "--oem 1 --psm 6 -c preserve_interword_spaces=0",
    ]
    best_text, best_conf = "", 0.0
    for v in _preprocess_variants(name_img):
        for cfg in cfgs:
            text, conf = _ocr_once(v, cfg, lang="jpn")
            text = text.replace("\n", " ")
            text = re.sub(r"\s+", " ", text).strip()
            text = re.sub(r"^[・\-\.\,()\[\]{}＿_　]+|[・\-\.\,()\[\]{}＿_　]+$", "", text)
            if text and conf > best_conf:
                best_text, best_conf = text, conf
    return best_text

# ================= メイン =================
def main():
    set_dpi_awareness()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # フォーカスクリック
    click_screen(CLICK_POS)
    time.sleep(0.2)

    seen = load_existing_names(OUTPUT_DIR)
    img_region  = rect_from_points(IMG_PTS)    # (x, y, w, h)
    name_region = rect_from_points(NAME_PTS)   # (x, y, w, h)

    for i in range(1, MAX_LOOPS + 1):
        # 画像領域をスクショ
        card_img = pyautogui.screenshot(region=img_region)

        # 名前帯をスクショ → 強化OCR → 日本語名の空白除去正規化
        name_img = pyautogui.screenshot(region=name_region)
        card_name_raw = ocr_card_name_smart(name_img)
        card_name = normalize_japanese_name(card_name_raw)

        # フォールバック（OCR失敗時）
        if not card_name:
            card_name = f"card_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            dup_key = None  # 重複停止には使わない
        else:
            dup_key = sanitize_filename(card_name).lower()

        # 重複名で停止
        if dup_key and dup_key in seen:
            print(f"[STOP] 同名検出: '{card_name}' → 終了します。")
            break

        # 保存（ファイル名衝突回避）
        safe = sanitize_filename(card_name)
        out_path = OUTPUT_DIR / f"{safe}.png"
        base = safe; idx = 1
        while out_path.exists():
            safe = f"{base}_{idx}"
            out_path = OUTPUT_DIR / f"{safe}.png"
            idx += 1
        card_img.save(out_path)
        print(f"[{i}] Saved: {out_path}   OCR(raw)='{card_name_raw}'  ->  name='{card_name}'")

        if dup_key:
            seen.add(dup_key)

        # ページ送り
        click_screen(PAGE_NEXT_POS)
        time.sleep(PAGE_WAIT)
    else:
        print(f"[END] MAX_LOOPS({MAX_LOOPS}) に到達 → 終了")

if __name__ == "__main__":
    main()
