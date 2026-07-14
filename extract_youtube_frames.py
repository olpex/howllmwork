# -*- coding: utf-8 -*-
"""
================================================================================
  Екстрактор кадрів з YouTube з перекладом російського тексту на українську
--------------------------------------------------------------------------------
  Конвеєр (для ЛОКАЛЬНОГО запуску — YouTube заблоковано в хмарному середовищі):

      1. yt-dlp   — завантажує відео один раз
      2. ffmpeg   — вирізає кадри за таймкодами з frames_timecodes.json
      3. EasyOCR  — знаходить російський текст і його розташування на кадрі
      4. Claude   — перекладає знайдений текст RU -> UA (модель claude-haiku-4-5)
      5. Pillow   — замальовує оригінальний напис і накладає український переклад

  Результат кладеться в images/slide_NN.png — саме туди, звідки
  generate_ai_presentation.py вбудовує зображення у плейсхолдери.

--------------------------------------------------------------------------------
  ВСТАНОВЛЕННЯ:
      # системні:
      #   ffmpeg  (apt install ffmpeg / brew install ffmpeg)
      # python:
      pip install yt-dlp easyocr pillow anthropic

  КЛЮЧ ДО API (для перекладу):
      export ANTHROPIC_API_KEY=...        # anthropic.Anthropic() читає його сам
      # або: ant auth login  (SDK підхопить профіль автоматично)

  ЗАПУСК:
      python extract_youtube_frames.py                 # весь конвеєр
      python extract_youtube_frames.py --no-translate  # без перекладу (сирі кадри)
      python extract_youtube_frames.py --slides 5,9,28 # лише вибрані слайди
================================================================================
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from typing import Dict, List, Optional, Tuple

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "frames_timecodes.json")
IMAGES_DIR = os.path.join(HERE, "images")
WORK_DIR = os.path.join(HERE, ".yt_work")
VIDEO_FILE = os.path.join(WORK_DIR, "source_video.mp4")

TRANSLATION_MODEL = "claude-haiku-4-5"  # дешева й швидка модель для перекладу


# ------------------------------------------------------------------ #
#  Допоміжне
# ------------------------------------------------------------------ #
def _have(cmd: str) -> bool:
    """Чи доступна зовнішня утиліта в PATH."""
    from shutil import which
    return which(cmd) is not None


def _to_seconds(t: str) -> float:
    """'MM:SS' або 'HH:MM:SS' -> секунди."""
    parts = [float(p) for p in t.split(":")]
    if len(parts) == 2:
        return parts[0] * 60 + parts[1]
    if len(parts) == 3:
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    return float(parts[0])


def load_manifest() -> Dict:
    with open(MANIFEST, "r", encoding="utf-8") as fh:
        return json.load(fh)


# ------------------------------------------------------------------ #
#  Крок 1: завантаження відео (yt-dlp)
# ------------------------------------------------------------------ #
def download_video(url: str) -> bool:
    """Завантажує відео один раз у WORK_DIR через yt-dlp."""
    if os.path.isfile(VIDEO_FILE) and os.path.getsize(VIDEO_FILE) > 0:
        print(f"[=] Відео вже завантажено: {VIDEO_FILE}")
        return True
    if not _have("yt-dlp"):
        print("[X] Не знайдено yt-dlp. Встановіть: pip install yt-dlp")
        return False
    os.makedirs(WORK_DIR, exist_ok=True)
    print(f"[↓] Завантаження відео: {url}")
    cmd = [
        "yt-dlp", "-f", "bestvideo[height<=1080]+bestaudio/best[height<=1080]",
        "--merge-output-format", "mp4", "-o", VIDEO_FILE, url,
    ]
    try:
        subprocess.run(cmd, check=True)
        return os.path.isfile(VIDEO_FILE)
    except subprocess.CalledProcessError as err:
        print(f"[X] yt-dlp помилка: {err}")
        return False


# ------------------------------------------------------------------ #
#  Крок 2: вирізання кадру (ffmpeg)
# ------------------------------------------------------------------ #
def extract_frame(time_str: str, dest: str) -> bool:
    """Вирізає один кадр на заданому таймкоді через ffmpeg."""
    if not _have("ffmpeg"):
        print("[X] Не знайдено ffmpeg. Встановіть: apt install ffmpeg / brew install ffmpeg")
        return False
    sec = _to_seconds(time_str)
    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-ss", str(sec), "-i", VIDEO_FILE,
        "-frames:v", "1", "-q:v", "2", dest,
    ]
    try:
        subprocess.run(cmd, check=True)
        return os.path.isfile(dest)
    except subprocess.CalledProcessError as err:
        print(f"    [!] ffmpeg помилка на {time_str}: {err}")
        return False


# ------------------------------------------------------------------ #
#  Крок 3: OCR російського тексту (EasyOCR)
# ------------------------------------------------------------------ #
_READER = None


def _get_reader():
    """Ліниво ініціалізує EasyOCR-читач (ru+en)."""
    global _READER
    if _READER is None:
        import easyocr  # важкий імпорт — тільки коли треба
        print("[i] Ініціалізація EasyOCR (ru+en)... (перший запуск завантажує моделі)")
        _READER = easyocr.Reader(["ru", "en"], gpu=False)
    return _READER


def ocr_boxes(image_path: str) -> List[Tuple[List[Tuple[int, int]], str, float]]:
    """Повертає [(bbox[4 точки], text, confidence), ...] для кадру."""
    reader = _get_reader()
    return reader.readtext(image_path)


# ------------------------------------------------------------------ #
#  Крок 4: переклад RU -> UA (Claude API, модель claude-haiku-4-5)
# ------------------------------------------------------------------ #
_CLIENT = None


def _get_client():
    """Ліниво створює Anthropic-клієнт. Ключ береться з ANTHROPIC_API_KEY / профілю."""
    global _CLIENT
    if _CLIENT is None:
        import anthropic
        # Порожній конструктор: SDK сам читає ANTHROPIC_API_KEY,
        # потім ANTHROPIC_AUTH_TOKEN, потім активний профіль `ant auth login`.
        _CLIENT = anthropic.Anthropic()
    return _CLIENT


def translate_batch(texts: List[str]) -> List[str]:
    """Перекладає список російських рядків українською одним запитом.

    Використовує дешеву швидку модель claude-haiku-4-5. Повертає список
    перекладів у тому ж порядку; за помилки повертає оригінали.
    """
    texts = [t.strip() for t in texts]
    if not any(texts):
        return texts
    try:
        client = _get_client()
        numbered = "\n".join(f"{i + 1}. {t}" for i, t in enumerate(texts))
        prompt = (
            "Переклади ці написи з російської на українську. Це підписи з "
            "навчального відео про штучний інтелект — зберігай технічні терміни. "
            "Поверни РІВНО стільки ж рядків, кожен у форматі '<номер>. <переклад>', "
            "без пояснень.\n\n" + numbered
        )
        resp = client.messages.create(
            model=TRANSLATION_MODEL,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        out_text = next((b.text for b in resp.content if b.type == "text"), "")
        # розбираємо пронумеровані рядки назад у список
        result = list(texts)
        for line in out_text.splitlines():
            line = line.strip()
            if not line or "." not in line:
                continue
            num, _, val = line.partition(".")
            if num.strip().isdigit():
                idx = int(num.strip()) - 1
                if 0 <= idx < len(result):
                    result[idx] = val.strip()
        return result
    except Exception as err:  # noqa: BLE001
        print(f"    [!] Помилка перекладу ({err}); лишаю оригінали.")
        return texts


# ------------------------------------------------------------------ #
#  Крок 5: замалювати оригінал і накласти переклад (Pillow)
# ------------------------------------------------------------------ #
def _load_font(size: int):
    from PIL import ImageFont
    for name in ("DejaVuSans.ttf", "Arial.ttf", "Verdana.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:  # noqa: BLE001
            continue
    return ImageFont.load_default()


def overlay_translation(image_path: str, dest: str,
                        boxes, translations: List[str]) -> None:
    """Замальовує кожен bbox фоном і малює українських переклад поверх."""
    from PIL import Image, ImageDraw
    img = Image.open(image_path).convert("RGB")
    draw = ImageDraw.Draw(img)

    for (bbox, _orig, _conf), ua in zip(boxes, translations):
        xs = [int(p[0]) for p in bbox]
        ys = [int(p[1]) for p in bbox]
        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
        w, h = x1 - x0, y1 - y0
        if w <= 0 or h <= 0:
            continue
        # колір фону беремо як середній піксель трохи вище рамки
        sample_y = max(0, y0 - 3)
        try:
            bg = img.getpixel((min(x0 + 2, img.width - 1), sample_y))
        except Exception:  # noqa: BLE001
            bg = (16, 36, 59)
        # заливаємо оригінальний напис
        draw.rectangle([x0, y0, x1, y1], fill=bg)
        # добираємо кегль під висоту рамки
        font = _load_font(max(10, int(h * 0.8)))
        # контрастний колір тексту
        luma = 0.299 * bg[0] + 0.587 * bg[1] + 0.114 * bg[2]
        fg = (0, 0, 0) if luma > 140 else (255, 255, 255)
        draw.text((x0 + 2, y0), ua, font=font, fill=fg)

    img.save(dest)


# ------------------------------------------------------------------ #
#  Оркестрація
# ------------------------------------------------------------------ #
def process(slides: Optional[List[str]], translate: bool) -> int:
    manifest = load_manifest()
    frames = manifest["frames"]
    url = manifest["video_url"]

    if slides:
        frames = {k: v for k, v in frames.items() if k in slides}
        if not frames:
            print("[X] Жоден зі вказаних слайдів не знайдено в маніфесті.")
            return 1

    if not download_video(url):
        return 1

    os.makedirs(IMAGES_DIR, exist_ok=True)
    ok, fail = 0, 0

    for slide_no, meta in sorted(frames.items(), key=lambda kv: int(kv[0])):
        raw = os.path.join(WORK_DIR, f"raw_{meta['file']}")
        dest = os.path.join(IMAGES_DIR, meta["file"])
        print(f"[slide {slide_no}] {meta['time']} — {meta['topic']}")

        if not extract_frame(meta["time"], raw):
            fail += 1
            continue

        if not translate:
            # без перекладу — просто копіюємо кадр
            from shutil import copyfile
            copyfile(raw, dest)
            ok += 1
            continue

        try:
            boxes = ocr_boxes(raw)
            ru_texts = [b[1] for b in boxes]
            if ru_texts:
                ua_texts = translate_batch(ru_texts)
                overlay_translation(raw, dest, boxes, ua_texts)
                print(f"    [✓] Перекладено написів: {len(ru_texts)}")
            else:
                from shutil import copyfile
                copyfile(raw, dest)
                print("    [i] Тексту не знайдено — кадр без змін.")
            ok += 1
        except Exception as err:  # noqa: BLE001
            print(f"    [!] Помилка обробки: {err}")
            fail += 1

    print("\n" + "=" * 52)
    print(f"[OK] Готово: {ok} | Помилок: {fail}")
    print(f"[OK] Зображення у: {IMAGES_DIR}")
    print("[OK] Далі: python generate_ai_presentation.py")
    return 0 if fail == 0 else 2


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Витягти кадри з YouTube і перекласти RU->UA написи.")
    ap.add_argument("--no-translate", action="store_true",
                    help="Не перекладати — зберегти сирі кадри.")
    ap.add_argument("--slides", type=str, default="",
                    help="Кома-розділений список номерів слайдів (напр. 5,9,28).")
    args = ap.parse_args()

    slides = [s.strip() for s in args.slides.split(",") if s.strip()] or None
    return process(slides, translate=not args.no_translate)


if __name__ == "__main__":
    sys.exit(main())
