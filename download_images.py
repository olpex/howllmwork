# -*- coding: utf-8 -*-
"""
================================================================================
  Завантажувач технічних діаграм для презентації
--------------------------------------------------------------------------------
  Зображення згенеровано через Higgsfield (nano_banana_pro, 1792x2400, 3:4) під
  плейсхолдери презентації "Як працює штучний інтелект".

  У хмарному середовищі Claude Code хост зображень (cloudfront) заблоковано
  мережевою політикою, тому завантаження треба виконати ЛОКАЛЬНО, де інтернет
  доступний без обмежень.

  ЗАПУСК (локально):
      python download_images.py

  Скрипт покладе файли у папку images/ як slide_03.png ... slide_36.png.
  Після цього просто запустіть генератор — він автоматично вбудує їх:
      python generate_ai_presentation.py
================================================================================
"""

from __future__ import annotations

import json
import os
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(HERE, "images")
MANIFEST = os.path.join(IMAGES_DIR, "images_manifest.json")


def main() -> int:
    """Завантажує всі зображення з маніфесту у папку images/."""
    if not os.path.isfile(MANIFEST):
        print(f"[X] Не знайдено маніфест: {MANIFEST}")
        return 1

    with open(MANIFEST, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    os.makedirs(IMAGES_DIR, exist_ok=True)
    items = data.get("images", {})
    ok, fail = 0, 0

    for slide_no, meta in sorted(items.items(), key=lambda kv: int(kv[0])):
        dest = os.path.join(IMAGES_DIR, meta["file"])
        url = meta["url"]
        if os.path.isfile(dest) and os.path.getsize(dest) > 0:
            print(f"[=] Слайд {slide_no}: вже є ({meta['file']})")
            ok += 1
            continue
        try:
            print(f"[↓] Слайд {slide_no}: {meta['file']}  —  {meta['desc']}")
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as out:
                out.write(resp.read())
            if os.path.getsize(dest) == 0:
                raise IOError("порожній файл")
            ok += 1
        except Exception as err:  # noqa: BLE001
            print(f"    [!] Помилка: {err}")
            fail += 1

    print("\n" + "=" * 50)
    print(f"[OK] Завантажено/наявно: {ok} | Помилок: {fail}")
    if fail == 0:
        print("[OK] Тепер запустіть: python generate_ai_presentation.py")
    return 0 if fail == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
