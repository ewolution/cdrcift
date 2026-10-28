"""cdrcift - Faz 3: PLT / HPGL mukerrer hat denetleyicisi.

CorelDRAW'dan sonra araya giren nesting / kesim yaziliminin cikardigi
PLT dosyalarinda ayni hatti iki kez cizen komutlari bulur.

Kullanim
--------
    python tools/plt_dedup.py dosya.plt
    python tools/plt_dedup.py dosya.plt --csv rapor.csv
    python tools/plt_dedup.py dosya.plt --temizle temiz.plt
    python tools/plt_dedup.py --test

Notlar
------
* Yalnizca **metin** HPGL desteklenir (CorelDRAW ve cogu nesting yazilimi
  metin PLT uretir). Ikili HPGL icin once metne cevirin.
* Varsayilan yalnizca **rapor** uretir; `--temizle` vermeden dosya yazilmaz.
* Pen numaralari (`SP`), hiz/kalite (`VS`, `FS`) ve diger komutlar korunur.
  Yalnizca mukerrer cizim hareketleri atilir; atlanan hareket `PU` olarak yazilir.
* Ayni hattin ters yonde cizilmesi de mukerrer sayilir (uca gore anahtar).

Cikis kodu: 0 = mukerrer yok, 1 = mukerrer var, 2 = hata.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

KOMUT = re.compile(r"([A-Za-z]{2})([^A-Za-z;]*)(;?)")


@dataclass
class Hareket:
    x: float
    y: float
    cizim: bool          # True = PD (kesim), False = PU (bos hareket)
    pen: int = 1


@dataclass
class Sonuc:
    komut_sayisi: int = 0
    cizim_segmenti: int = 0
    mukerrer: int = 0
    yakin: int = 0
    bos_uzunluk: float = 0.0
    ornekler: list[tuple[float, float, float, float]] = field(default_factory=list)


# ---------------------------------------------------------------------
# Ayristirma
# ---------------------------------------------------------------------

def komutlara_ayir(metin: str) -> list[tuple[str, list[float], str]]:
    """HPGL metnini (komut, parametreler, ham metin) listesine cevirir."""
    out: list[tuple[str, list[float], str]] = []
    for m in KOMUT.finditer(metin):
        ad = m.group(1).upper()
        govde = m.group(2).strip()
        ham = m.group(0)
        if not govde:
            out.append((ad, [], ham))
            continue
        parcalar = [p for p in re.split(r"[,\s]+", govde) if p]
        degerler: list[float] = []
        sayisal = True
        for p in parcalar:
            try:
                degerler.append(float(p))
            except ValueError:
                sayisal = False
                break
        if sayisal:
            out.append((ad, degerler, ham))
        else:
            out.append((ad, [], ham))      # bilinmeyen parametre: aynen koru
    return out


def hareketleri_uret(komutlar: list[tuple[str, list[float], str]]) -> list[Hareket]:
    """Pen konumunu izleyerek cizim hareketlerini uretir."""
    hareketler: list[Hareket] = []
    x = y = 0.0
    mutlak = True
    cizim = False
    pen = 1

    for ad, p, _ in komutlar:
        if ad == "IN":
            x = y = 0.0
            mutlak = True
            cizim = False
        elif ad == "PA":
            mutlak = True
            _noktalar(p, x, y, mutlak, cizim, pen, hareketler)
            if len(p) >= 2:
                x, y = p[-2], p[-1]
        elif ad == "PR":
            mutlak = False
            _noktalar(p, x, y, mutlak, cizim, pen, hareketler)
            if len(p) >= 2:
                x += p[-2]
                y += p[-1]
        elif ad == "PU":
            cizim = False
            _noktalar(p, x, y, mutlak, cizim, pen, hareketler)
            if len(p) >= 2:
                if mutlak:
                    x, y = p[-2], p[-1]
                else:
                    x += p[-2]
                    y += p[-1]
        elif ad == "PD":
            cizim = True
            _noktalar(p, x, y, mutlak, cizim, pen, hareketler)
            if len(p) >= 2:
                if mutlak:
                    x, y = p[-2], p[-1]
                else:
                    x += p[-2]
                    y += p[-1]
        elif ad == "SP" and p:
            pen = int(p[0])
    return hareketler


def _noktalar(p: list[float], x: float, y: float, mutlak: bool, cizim: bool,
              pen: int, out: list[Hareket]) -> None:
    for i in range(0, len(p) - 1, 2):
        px, py = p[i], p[i + 1]
        if mutlak:
            nx, ny = px, py
        else:
            nx, ny = x + px, y + py
        out.append(Hareket(nx, ny, cizim, pen))
        x, y = nx, ny


# ---------------------------------------------------------------------
# Mukerrer tespiti
# ---------------------------------------------------------------------

def _anahtar(a: Hareket, b: Hareket, yuvarla: float) -> tuple:
    p1 = (round(a.x / yuvarla), round(a.y / yuvarla))
    p2 = (round(b.x / yuvarla), round(b.y / yuvarla))
    return (p1, p2) if p1 <= p2 else (p2, p1)


def _hucre(x: float, y: float, hucre: float) -> tuple[int, int]:
    return int(math.floor(x / hucre)), int(math.floor(y / hucre))


def analiz(hareketler: list[Hareket], yuvarla: float = 0.01,
           yakin: float = 0.0) -> tuple[Sonuc, list[tuple[int, str, Hareket, Hareket]]]:
    """(ozet, [(indeks, tur, a, b), ...]) dondurur."""
    s = Sonuc()
    onceki = hareketler[0] if hareketler else None
    gorulen: dict[tuple, int] = {}
    izgara: dict[tuple[int, int], list[tuple[float, float, float, float]]] = {}
    isaretler: list[tuple[int, str, Hareket, Hareket]] = []

    for i in range(1, len(hareketler)):
        a, b = hareketler[i - 1], hareketler[i]
        if not b.cizim:
            onceki = b
            continue
        boy = math.hypot(b.x - a.x, b.y - a.y)
        s.cizim_segmenti += 1
        if boy <= yuvarla:
            continue

        k = _anahtar(a, b, yuvarla)
        if k in gorulen:
            s.mukerrer += 1
            s.bos_uzunluk += boy
            isaretler.append((i, "TAM", a, b))
            if len(s.ornekler) < 20:
                s.ornekler.append((a.x, a.y, b.x, b.y))
        elif yakin > 0:
            yak = _yakinda_var(izgara, a, b, yakin, yuvarla)
            if yak:
                s.yakin += 1
                isaretler.append((i, "YAKIN", a, b))
        else:
            gorulen[k] = i

        if yakin > 0:
            _izgaraya_ekle(izgara, a, b, yakin, yuvarla)
        onceki = b

    return s, isaretler


def _izgaraya_ekle(izgara, a: Hareket, b: Hareket, hucre: float, yuvarla: float) -> None:
    cx, cy = _hucre((a.x + b.x) / 2, (a.y + b.y) / 2, hucre)
    izgara.setdefault((cx, cy), []).append(
        (round(a.x / yuvarla), round(a.y / yuvarla), round(b.x / yuvarla), round(b.y / yuvarla))
    )


def _yakinda_var(izgara, a: Hareket, b: Hareket, hucre: float, yuvarla: float) -> bool:
    cx, cy = _hucre((a.x + b.x) / 2, (a.y + b.y) / 2, hucre)
    ax1, ay1 = a.x / yuvarla, a.y / yuvarla
    bx1, by1 = b.x / yuvarla, b.y / yuvarla
    sinir = hucre / yuvarla
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for (x1, y1, x2, y2) in izgara.get((cx + dx, cy + dy), ()):
                if (abs(x1 - ax1) <= sinir and abs(y1 - ay1) <= sinir and
                        abs(x2 - bx1) <= sinir and abs(y2 - by1) <= sinir) or \
                   (abs(x1 - bx1) <= sinir and abs(y1 - by1) <= sinir and
                        abs(x2 - ax1) <= sinir and abs(y2 - ay1) <= sinir):
                    return True
    return False


# ---------------------------------------------------------------------
# Temiz PLT yazimi
# ---------------------------------------------------------------------

def temizle(komutlar: list[tuple[str, list[float], str]], yuvarla: float) -> tuple[str, int]:
    """Mukerrer cizim hareketlerini atlayarak yeni HPGL metni uretir."""
    gorulen: dict[tuple, int] = {}
    x = y = 0.0
    mutlak = True
    cizim = False
    cikti: list[str] = []
    atilan = 0

    for ad, p, ham in komutlar:
        if ad in ("PA", "PR", "PU", "PD") and len(p) >= 2:
            yeni_mutlak = mutlak if ad in ("PU", "PD") else (ad == "PA")
            yeni_cizim = cizim if ad in ("PA", "PR") else (ad == "PD")
            parcalar: list[str] = []

            for i in range(0, len(p) - 1, 2):
                if yeni_mutlak:
                    nx, ny = p[i], p[i + 1]
                else:
                    nx, ny = x + p[i], y + p[i + 1]

                if yeni_cizim:
                    k = _anahtar(Hareket(x, y, False), Hareket(nx, ny, True), yuvarla)
                    if k in gorulen:
                        atilan += 1
                        if parcalar:
                            cikti.append(f"{ad}{','.join(parcalar)};")
                            parcalar = []
                        # Kalemi kaldirarak konumu degistir: cizim olmaz
                        cikti.append(f"PU{_f(nx)},{_f(ny)};")
                        x, y = nx, ny
                        continue
                    gorulen[k] = 1

                parcalar.append(f"{_f(nx)},{_f(ny)}")
                x, y = nx, ny

            if parcalar:
                cikti.append(f"{ad}{','.join(parcalar)};")
            mutlak = yeni_mutlak
            cizim = yeni_cizim
        else:
            cikti.append(ham if ham.endswith(";") else ham + ";")
            if ad == "PA":
                mutlak = True
            elif ad == "PR":
                mutlak = False
            elif ad == "PU":
                cizim = False
            elif ad == "PD":
                cizim = True

    return "".join(cikti), atilan


def _f(v: float) -> str:
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return f"{v:.3f}".rstrip("0").rstrip(".")


# ---------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------

def rapor_yaz(s: Sonuc, kaynak: Path) -> None:
    birim = 40.0     # HPGL varsayilani: 1016 birim/inc ~ 40 birim/mm
    print(f"Dosya            : {kaynak}")
    print(f"Komut            : {s.komut_sayisi}")
    print(f"Cizim segmenti   : {s.cizim_segmenti}")
    print(f"Mukerrer (TAM)   : {s.mukerrer}")
    if s.yakin:
        print(f"Yakin kopya      : {s.yakin}")
    print(f"Bosuna cizim     : {s.bos_uzunluk:.0f} birim  (~{s.bos_uzunluk / birim:.1f} mm)")
    if s.ornekler:
        print("Ornekler (x1,y1 -> x2,y2):")
        for o in s.ornekler[:10]:
            print(f"  {o[0]:.0f},{o[1]:.0f} -> {o[2]:.0f},{o[3]:.0f}")


def test() -> int:
    ornek = ("IN;SP1;PU0,0;PD1000,0;PD1000,1000;PD0,1000;PD0,0;"
             "PU0,0;PD1000,0;"          # mukerrer: ilk kenar
             "PU2000,0;PD3000,0;PU3000,0;PD2000,0;")   # mukerrer: ters yon
    komutlar = komutlara_ayir(ornek)
    hareketler = hareketleri_uret(komutlar)
    s, isaretler = analiz(hareketler, 1.0)
    hatalar = []
    if s.cizim_segmenti != 7:
        hatalar.append(f"segment sayisi 7 bekleniyordu, {s.cizim_segmenti}")
    if s.mukerrer != 2:
        hatalar.append(f"mukerrer 2 bekleniyordu, {s.mukerrer}")
    if not isaretler:
        hatalar.append("isaret listesi bos")

    yeni, atilan = temizle(komutlar, 1.0)
    s2, _ = analiz(hareketleri_uret(komutlara_ayir(yeni)), 1.0)
    if atilan != 2:
        hatalar.append(f"temizlemede 2 atilma bekleniyordu, {atilan}")
    if s2.mukerrer != 0:
        hatalar.append(f"temizlenmis dosyada mukerrer kaldi: {s2.mukerrer}")
    if s2.cizim_segmenti != 5:
        hatalar.append(f"temizlenmis segment 5 bekleniyordu, {s2.cizim_segmenti}")

    if hatalar:
        for h in hatalar:
            print("HATA " + h)
        return 2
    print("plt_dedup testleri gecti.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="PLT/HPGL mukerrer hat denetleyicisi")
    ap.add_argument("dosya", nargs="?", help="girdi .plt")
    ap.add_argument("--csv", help="rapor CSV yolu")
    ap.add_argument("--temizle", metavar="CIKTI", help="temizlenmis PLT yaz")
    ap.add_argument("--yuvarla", type=float, default=0.01,
                    help="koordinat yuvarlama (plotter birimi, varsayilan 0.01)")
    ap.add_argument("--yakin", type=float, default=0.0,
                    help="yakin kopya hucre boyu (plotter birimi; 0 = kapali)")
    ap.add_argument("--test", action="store_true", help="dahili testleri calistir")
    a = ap.parse_args()

    if a.test:
        return test()
    if not a.dosya:
        ap.print_help()
        return 2

    yol = Path(a.dosya)
    if not yol.exists():
        print(f"Dosya yok: {yol}", file=sys.stderr)
        return 2

    metin = yol.read_text(encoding="latin-1", errors="replace")
    komutlar = komutlara_ayir(metin)
    hareketler = hareketleri_uret(komutlar)
    s, isaretler = analiz(hareketler, a.yuvarla, a.yakin)
    s.komut_sayisi = len(komutlar)

    rapor_yaz(s, yol)

    if a.csv:
        with open(a.csv, "w", encoding="utf-8-sig", newline="") as f:
            f.write("no;tur;x1;y1;x2;y2\n")
            for no, tur, p1, p2 in isaretler:
                f.write(f"{no};{tur};{p1.x:.0f};{p1.y:.0f};{p2.x:.0f};{p2.y:.0f}\n")
        print(f"CSV              : {a.csv}")

    if a.temizle:
        yeni, atilan = temizle(komutlar, a.yuvarla)
        Path(a.temizle).write_text(yeni, encoding="latin-1", errors="replace")
        print(f"Temiz dosya      : {a.temizle}  ({atilan} hareket atildi)")

    return 1 if s.mukerrer else 0


if __name__ == "__main__":
    raise SystemExit(main())
