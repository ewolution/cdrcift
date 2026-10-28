"""cdrcift - referans algoritma testleri.

Calistirma:  python tools/test_geom.py
"""

from __future__ import annotations

import math
import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from geom_referans import (  # noqa: E402
    Finding, Seg, Shape, SubPath, analyse, classify, signature,
)

GECEN = 0
KALAN = 0
RAPOR: list[str] = []


def kontrol(kosul: bool, aciklama: str) -> None:
    global GECEN, KALAN
    if kosul:
        GECEN += 1
        RAPOR.append(f"[OK]   {aciklama}")
    else:
        KALAN += 1
        RAPOR.append(f"[HATA] {aciklama}")


def cizgi(x1, y1, x2, y2, idx, z=None) -> Shape:
    sp = SubPath([Seg(x1, y1, x2, y2)], closed=False)
    return Shape(idx, z=idx if z is None else z, subpaths=[sp])


def dikdortgen(l, b, r, t, idx) -> Shape:
    segs = [
        Seg(l, b, l, t),
        Seg(l, t, r, t),
        Seg(r, t, r, b),
        Seg(r, b, l, b),
    ]
    return Shape(idx, z=idx, subpaths=[SubPath(segs, closed=True)])


def bul(findings: list[Finding], tur: str, x: float, y: float, pay: float = 1.0):
    for f in findings:
        if f.tur == tur and abs(f.x - x) <= pay and abs(f.y - y) <= pay:
            return f
    return None


def fixture() -> list[Shape]:
    shapes: list[Shape] = []
    shapes.append(cizgi(10, 10, 10, 60, 0))          # A
    shapes.append(cizgi(10, 10, 10, 60, 1))          # B: birebir kopya
    shapes.append(cizgi(110, 10, 110, 60, 2))        # C1
    shapes.append(cizgi(110, 60, 110, 10, 3))        # C2: ters yon
    shapes.append(dikdortgen(200, 10, 240, 50, 4))   # D1
    shapes.append(dikdortgen(240, 10, 280, 50, 5))   # D2: ortak kenar x=240
    shapes.append(cizgi(300, 10, 300, 60, 6))        # E1
    shapes.append(cizgi(300.3, 10, 300.3, 60, 7))    # E2: 0,3 mm kaymis
    shapes.append(cizgi(320, 10, 320, 60, 8))        # F1
    shapes.append(cizgi(320.1, 10, 320.1, 60, 9))    # F2: 0,1 mm kaymis
    return shapes


def test_senaryo() -> None:
    f = analyse(fixture())

    tam = bul(f, "TAM", 10, 35)
    kontrol(tam is not None, "Birebir kopya (10,35) TAM bulundu")
    kontrol(tam is not None and tam.eylem == "sil", "Birebir kopya otomatik silinebilir")

    kontrol(bul(f, "TAM", 110, 35) is not None, "Ters yondeki kopya (110,35) TAM bulundu")

    kismi = bul(f, "KISMI", 240, 30, pay=2)
    kontrol(kismi is not None, "Ortak kenar (240,30) KISMI bulundu")
    kontrol(kismi is not None and kismi.eylem == "elle", "KISMI asla otomatik silinmez")

    yakin = bul(f, "YAKIN", 300.15, 35)
    kontrol(yakin is not None, "0,3 mm kaymis cizgi YAKIN bulundu")
    kontrol(yakin is not None and yakin.eylem == "elle", "YAKIN asla otomatik silinmez")

    kontrol(bul(f, "TAM", 320.05, 35, pay=2) is None, "0,1 mm kaymis cizgi TAM sayilmadi")

    kontrol(all(x.eylem != "sil" for x in f if x.tur in ("KISMI", "YAKIN")),
            "KISMI/YAKIN icinde 'sil' eylemi yok")

    # ozet
    sayim: dict[str, int] = {}
    for x in f:
        sayim[x.tur] = sayim.get(x.tur, 0) + 1
    kontrol(sayim.get("TAM", 0) == 2, f"Tam olarak 2 TAM beklendi, gelen {sayim.get('TAM', 0)}")
    kontrol(sayim.get("YAKIN", 0) >= 2, f"En az 2 YAKIN beklendi, gelen {sayim.get('YAKIN', 0)}")
    RAPOR.append("       sayim: " + ", ".join(f"{k}={v}" for k, v in sorted(sayim.items())))


def test_imza() -> None:
    a = dikdortgen(0, 0, 100, 50, 0).subpaths[0]
    # ayni dikdortgen, ters yon ve farkli baslangic dugumu
    segs = [
        Seg(100, 0, 100, 50),
        Seg(100, 50, 0, 50),
        Seg(0, 50, 0, 0),
        Seg(0, 0, 100, 0),
    ]
    b = SubPath(segs, closed=True)
    kontrol(signature(a) == signature(b),
            "Kapali yol: ters yon + farkli baslangic dugumu ayni imza")

    # 0,01 mm kaymis kopya ayni imza olmamali (izgara 0,005 mm)
    c = dikdortgen(0.01, 0, 100.01, 50, 0).subpaths[0]
    kontrol(signature(a) != signature(c), "0,01 mm kaymis dikdortgen farkli imza")

    # cizgi: ters yon
    d1 = SubPath([Seg(0, 0, 50, 0)])
    d2 = SubPath([Seg(50, 0, 0, 0)])
    kontrol(signature(d1) == signature(d2), "Acik cizgi: ters yon ayni imza")


def test_izgara_esdegerligi() -> None:
    """Izgara, kaba kuvvet (O(n^2)) ile ayni ciftleri bulmali."""
    random.seed(20261028)
    segs: list[Seg] = []
    for _ in range(120):
        x = random.uniform(0, 200)
        y = random.uniform(0, 200)
        aci = random.choice([0, 90, 45, 30, 12.5, random.uniform(0, 360)])
        boy = random.choice([5, 20, 60, 150])
        dx = math.cos(math.radians(aci)) * boy
        dy = math.sin(math.radians(aci)) * boy
        segs.append(Seg(x, y, x + dx, y + dy))

    # bazi kopyalar ve kaymis kopyalar ekle (gercekci senaryo)
    for _ in range(30):
        s = random.choice(segs)
        segs.append(Seg(s.x1, s.y1, s.x2, s.y2))
    for _ in range(15):
        s = random.choice(segs)
        segs.append(Seg(s.x1 + 0.2, s.y1 + 0.2, s.x2 + 0.2, s.y2 + 0.2))

    tol, yakin, min_ov, hucre = 0.05, 0.5, 0.5, 2.0

    kaba: set[tuple[int, int]] = set()
    for i in range(len(segs)):
        for j in range(i + 1, len(segs)):
            if classify(i, j, segs, tol, yakin, min_ov):
                kaba.add((i, j))

    from geom_referans import Grid
    g = Grid(segs, hucre)
    izgara: set[tuple[int, int]] = set()
    for i in range(len(segs)):
        for j in g.candidates(i):
            if j > i and classify(i, j, segs, tol, yakin, min_ov):
                izgara.add((i, j))

    eksik = kaba - izgara
    kontrol(not eksik, f"Izgara kaba kuvvetle ayni ciftleri buldu (eksik: {len(eksik)})")
    if eksik:
        ornek = sorted(eksik)[:5]
        RAPOR.append("       eksik ornek: " + str(ornek))
        for i, j in ornek:
            RAPOR.append(f"         {i}:{segs[i]}  {j}:{segs[j]}")


def main() -> int:
    test_imza()
    test_senaryo()
    test_izgara_esdegerligi()

    print("\n".join(RAPOR))
    print()
    print(f"Gecen: {GECEN}   Kalan: {KALAN}")
    return 0 if KALAN == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
