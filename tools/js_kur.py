"""cdrcift - JavaScript surumunu CorelDRAW Scripts klasorune kurar.

Kaynak dosyalar UTF-8'dir (gelistirme icin). CorelDRAW'in JavaScript
host'u, mevcut Summa scriptleriyle ayni sekilde ANSI (Turkce sistemde
cp1254) bekler; bu arac cevrimi yapar ve kurulumu dogrular.

Kullanim:
    python tools/js_kur.py                       # 2026
    python tools/js_kur.py --surum "Graphics Suite 2025"
    python tools/js_kur.py --kuru                # sadece kopyala, dogrulama yok
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

KAYNAK = Path(__file__).resolve().parent.parent / "js"
DOSYALAR = ["CiftCizgi.js"]
KODLAMA = "cp1254"


def hedef_klasor(surum: str) -> Path:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise SystemExit("APPDATA bulunamadi")
    return Path(appdata) / "Corel" / ("CorelDRAW " + surum) / "Draw" / "Scripts"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--surum", default="Graphics Suite 2026",
                    help="CorelDRAW klasor adi (varsayilan: Graphics Suite 2026)")
    ap.add_argument("--kuru", action="store_true", help="yalnizca kopyala")
    a = ap.parse_args()

    hedef = hedef_klasor(a.surum)
    hedef.mkdir(parents=True, exist_ok=True)
    print(f"Kaynak : {KAYNAK}")
    print(f"Hedef  : {hedef}")

    hata = 0
    for ad in DOSYALAR:
        p = KAYNAK / ad
        if not p.exists():
            print(f"  [HATA] kaynak yok: {p}")
            hata += 1
            continue

        metin = p.read_text(encoding="utf-8")
        try:
            veri = metin.encode(KODLAMA)
        except UnicodeEncodeError as e:
            kotu = metin[e.start:e.end]
            print(f"  [HATA] {ad}: {KODLAMA} icinde olmayan karakter {kotu!r}")
            hata += 1
            continue

        q = hedef / ad
        q.write_bytes(veri)

        # dogrulama: geri oku ve karsilastir
        geri = q.read_bytes().decode(KODLAMA)
        if geri != metin:
            print(f"  [HATA] {ad}: geri okuma farki var")
            hata += 1
            continue

        print(f"  [OK] {ad:24s} {len(veri):6d} bayt  ({KODLAMA})")

    if hata:
        return 1
    print("\nKurulum tamam. CorelDRAW'da: Araclar > Makrolar > JavaScript")
    print("  veya Scripts panelinden CiftCizgi.js")
    return 0


if __name__ == "__main__":
    sys.exit(main())
