"""cdrcift - geometri cekirdeginin Python referans uygulamasi.

Amac
----
VBA kodu CorelDRAW icinde calisir; burada calistirilamaz. Bu modul
`src/CiftCizgiGeom.bas` ve `src/CiftCizgiTarama.bas` icindeki algoritmayi
birebir taklit eder, boylece siniflandirma ve izgara mantigi sayisal
olarak dogrulanabilir.

Eslesme noktalari (VBA <-> Python):
  Karsilastir        -> classify()
  AltYolImzasi       -> signature()
  IzgarayaEkle       -> Grid.insert()
  AdaylariTopla      -> Grid.candidates()
  YapisalTamTespit   -> structural_pass()
  SegmentTarama      -> segment_pass()
  BulguTamamla       -> finalize()

Toleranslar mm cinsindendir.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

QZ = 0.005          # imza izgarasi (mm)
PARALEL_SIN = math.sin(math.radians(1.0))


# ---------------------------------------------------------------------
# Veri modeli
# ---------------------------------------------------------------------

@dataclass
class Seg:
    """Bir alt yolun duz (veya duzlestirilmis) parcasi."""
    x1: float
    y1: float
    x2: float
    y2: float
    typ: int = 0                      # 0 = cizgi, 1 = egri (kontrol noktali)
    c1: tuple[float, float] | None = None
    c2: tuple[float, float] | None = None

    @property
    def length(self) -> float:
        return math.hypot(self.x2 - self.x1, self.y2 - self.y1)


@dataclass
class SubPath:
    segs: list[Seg]
    closed: bool = False
    # VBA'da gAltYolX/Y: noktalarin agirlik merkezi
    cx: float = 0.0
    cy: float = 0.0

    @property
    def length(self) -> float:
        return sum(s.length for s in self.segs)


@dataclass
class Shape:
    idx: int
    layer: str = "Layer 1"
    color: str = "#000000"
    width: float = 0.0
    z: int = 0            # derinlik: buyuk = altta (CorelDRAW ZOrder ile ayni)
    subpaths: list[SubPath] = field(default_factory=list)
    name: str = ""

    def __post_init__(self) -> None:
        if not self.name:
            self.name = f"Nesne {self.idx + 1}"


@dataclass
class Finding:
    tur: str
    n1: int
    a1: int
    n2: int
    a2: int
    length: float
    ratio: float
    x: float
    y: float
    eylem: str = "elle"
    note: str = ""


# ---------------------------------------------------------------------
# Imza (yapisal TAM tespiti)
# ---------------------------------------------------------------------

def _tok(s: Seg, ters: bool = False) -> str:
    if not ters:
        t = f"{s.typ}|{qz(s.x1)}|{qz(s.y1)}|{qz(s.x2)}|{qz(s.y2)}"
        if s.typ == 1:
            t += f"|{qz(s.c1[0])}|{qz(s.c1[1])}|{qz(s.c2[0])}|{qz(s.c2[1])}"
    else:
        t = f"{s.typ}|{qz(s.x2)}|{qz(s.y2)}|{qz(s.x1)}|{qz(s.y1)}"
        if s.typ == 1:
            t += f"|{qz(s.c2[0])}|{qz(s.c2[1])}|{qz(s.c1[0])}|{qz(s.c1[1])}"
    return t


def qz(v: float) -> int:
    return int(math.floor(v / QZ + 0.5))


def _rotate_min(tokens: list[str]) -> str:
    n = len(tokens)
    start = min(range(n), key=lambda i: tokens[i])
    return ";".join(tokens[(start + i) % n] for i in range(n))


def signature(sp: SubPath) -> str:
    if not sp.segs:
        return ""
    fwd = [_tok(s) for s in sp.segs]
    rev = [_tok(s, ters=True) for s in reversed(sp.segs)]
    if sp.closed:
        a, b = _rotate_min(fwd), _rotate_min(rev)
    else:
        a, b = ";".join(fwd), ";".join(rev)
    s = min(a, b)
    return ("C" if sp.closed else "O") + str(len(sp.segs)) + "|" + s


# ---------------------------------------------------------------------
# Siniflandirma (VBA: Karsilastir)
# ---------------------------------------------------------------------

def classify(i: int, j: int, segs: list[Seg], tol: float, yakin: float,
             min_ov: float) -> tuple[str, float, float, float] | None:
    a, b = segs[i], segs[j]
    l1, l2 = a.length, b.length
    if l1 < 0.001 or l2 < 0.001:
        return None

    ux, uy = (a.x2 - a.x1) / l1, (a.y2 - a.y1) / l1
    vx, vy = (b.x2 - b.x1) / l2, (b.y2 - b.y1) / l2

    if abs(ux * vy - uy * vx) > PARALEL_SIN:
        return None

    d1 = abs((b.x1 - a.x1) * -uy + (b.y1 - a.y1) * ux)
    d2 = abs((b.x2 - a.x1) * -uy + (b.y2 - a.y1) * ux)
    off = max(d1, d2)

    t1 = (b.x1 - a.x1) * ux + (b.y1 - a.y1) * uy
    t2 = (b.x2 - a.x1) * ux + (b.y2 - a.y1) * uy
    lo = max(0.0, min(t1, t2))
    hi = min(l1, max(t1, t2))
    ov = hi - lo
    if ov < 0.001:
        return None

    mn = min(l1, l2)

    if off <= tol:
        if ov >= mn - 0.001:
            tur = "TAM"
        elif ov >= min_ov:
            tur = "KISMI"
        else:
            return None
    elif off <= yakin and ov >= 0.8 * mn:
        tur = "YAKIN"
    else:
        return None

    return tur, ov, a.x1 + ux * (lo + ov / 2), a.y1 + uy * (lo + ov / 2)


# ---------------------------------------------------------------------
# Uzamsal izgara (VBA: mIzgara)
# ---------------------------------------------------------------------

class Grid:
    def __init__(self, segs: list[Seg], hucre: float = 2.0) -> None:
        self.segs = segs
        self.hucre = hucre
        self.cells: dict[tuple[int, int], list[int]] = {}
        for i in range(len(segs)):
            self.insert(i)

    def _samples(self, i: int, adim: float):
        s = self.segs[i]
        adim = max(adim, 0.25)
        n = int(s.length / adim) + 1
        for k in range(n + 1):
            t = 0.0 if n == 0 else k / n
            yield (s.x1 + (s.x2 - s.x1) * t, s.y1 + (s.y2 - s.y1) * t)

    def _cell(self, x: float, y: float) -> tuple[int, int]:
        return (int(math.floor(x / self.hucre)), int(math.floor(y / self.hucre)))

    def insert(self, i: int) -> None:
        last = None
        for x, y in self._samples(i, self.hucre / 2):
            c = self._cell(x, y)
            if c != last:
                self.cells.setdefault(c, []).append(i)
                last = c

    def candidates(self, i: int) -> set[int]:
        out: set[int] = set()
        seen: set[tuple[int, int]] = set()
        for x, y in self._samples(i, self.hucre):
            cx, cy = self._cell(x, y)
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    c = (cx + dx, cy + dy)
                    if c in seen:
                        continue
                    seen.add(c)
                    out.update(self.cells.get(c, ()))
        return out


# ---------------------------------------------------------------------
# Tum boru hatti (VBA: TaramaYap)
# ---------------------------------------------------------------------

def _pair_key(n1: int, a1: int, n2: int, a2: int) -> str:
    if n1 < n2 or (n1 == n2 and a1 <= a2):
        return f"{n1}|{a1}|{n2}|{a2}"
    return f"{n2}|{a2}|{n1}|{a1}"


def analyse(shapes: list[Shape], tol: float = 0.05, yakin: float = 0.5,
            min_ov: float = 0.5, hucre: float = 2.0,
            otomatik_sil: bool = True) -> list[Finding]:
    # Alt yol slotlari
    slots: list[tuple[int, int, SubPath]] = []
    for sh in shapes:
        for k, sp in enumerate(sh.subpaths, start=1):
            sp.cx, sp.cy = body_centroid(sp)
            slots.append((sh.idx, k, sp))

    by_idx = {sh.idx: sh for sh in shapes}

    findings: list[Finding] = []
    listed: dict[str, int] = {}
    tam_cift: set[str] = set()

    # --- 1) Yapisal TAM ---
    groups: dict[str, list[int]] = {}
    for slot, (_, _, sp) in enumerate(slots):
        sig = signature(sp)
        if sig:
            groups.setdefault(sig, []).append(slot)

    for members in groups.values():
        if len(members) < 2:
            continue
        keeper = max(members, key=lambda s: by_idx[slots[s][0]].z)
        for s in members:
            if s == keeper:
                continue
            n_d, a_d, sp_d = slots[s]
            n_k, a_k, sp_k = slots[keeper]
            if n_d == n_k:
                eylem, note = "sil", "ayni nesne icinde tekrar"
            else:
                ok, note = same_attrs(by_idx[n_d], by_idx[n_k])
                eylem = "sil" if (ok and otomatik_sil) else "elle"
            findings.append(Finding("TAM", n_d, a_d, n_k, a_k, sp_d.length,
                                    1.0, sp_d.cx, sp_d.cy, eylem, note))
            tam_cift.add(_pair_key(n_d, a_d, n_k, a_k))

    # --- 2) Segment taramasi ---
    segs: list[Seg] = []
    seg_slot: list[int] = []
    for slot, (_, _, sp) in enumerate(slots):
        for s in sp.segs:
            segs.append(s)
            seg_slot.append(slot)

    grid = Grid(segs, hucre)
    for i in range(len(segs)):
        for j in sorted(grid.candidates(i)):
            if j <= i:
                continue
            sa, sb = seg_slot[i], seg_slot[j]
            if sa == sb and (j - i) == 1:
                continue
            r = classify(i, j, segs, tol, yakin, min_ov)
            if not r:
                continue
            tur, ov, x, y = r
            n_a, a_a, _ = slots[sa]
            n_b, a_b, _ = slots[sb]
            if _pair_key(n_a, a_a, n_b, a_b) in tam_cift:
                continue
            key = _pair_key(n_a, a_a, n_b, a_b) + "|" + tur
            if key in listed:
                findings[listed[key]].length += ov
            else:
                listed[key] = len(findings)
                findings.append(Finding(tur, n_a, a_a, n_b, a_b, ov, 0.0, x, y))

    # --- 3) BulguTamamla ---
    lengths: dict[str, float] = {}
    for n, a, sp in slots:
        lengths.setdefault(f"{n}|{a}", sp.length)

    for f in findings:
        if f.tur != "NOKTAS":
            b1 = lengths.get(f"{f.n1}|{f.a1}", 0.0)
            b2 = lengths.get(f"{f.n2}|{f.a2}", b1) if f.n2 >= 0 else b1
            mn = min(b1, b2)
            if mn > 0:
                f.ratio = min(1.0, f.length / mn)

        if f.tur == "TAM" and f.ratio < 0.98:
            f.tur = "KISMI"

        if f.tur == "TAM":
            if f.eylem != "sil" and otomatik_sil and f.n1 != f.n2:
                ok, _ = same_attrs(by_idx[f.n1], by_idx[f.n2])
                if ok:
                    f.eylem = "sil"
            if f.eylem == "sil" and not otomatik_sil:
                f.eylem = "elle"
        else:
            f.eylem = "elle"

    return findings


def _centroid(sp: SubPath) -> float:
    return body_centroid(sp)[0]


def same_attrs(a: Shape, b: Shape) -> tuple[bool, str]:
    if a.layer != b.layer:
        return False, "farkli katman"
    if abs(a.width - b.width) > 0.001:
        return False, "farkli cizgi kalinligi"
    if a.color.lower() != b.color.lower():
        return False, "farkli cizgi rengi"
    return True, ""


def body_centroid(sp: SubPath) -> tuple[float, float]:
    """VBA Duzlestir ile ayni: tum segment uc noktalarinin ortalamasi."""
    xs: list[float] = []
    ys: list[float] = []
    for s in sp.segs:
        xs.extend((s.x1, s.x2))
        ys.extend((s.y1, s.y2))
    if not xs:
        return 0.0, 0.0
    return sum(xs) / len(xs), sum(ys) / len(ys)
