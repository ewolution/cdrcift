"""cdrcift - VBA kaynaklari icin hafif yapisal denetleyici.

VBA derleyicisi yok; bu arac VBE'ye import etmeden once yakalanabilen
hatalari yakalar:

  * blok dengesi        (Sub/If/With/For/Do/Select/Type/Enum)
  * ayni prosedurde tanimsiz/tekrarlanan Dim
  * proje genelinde tekrarlanan prosedur adi
  * 1023 karakterden uzun satir
  * End Sub/End Function eslesmesi

Kullanim:  python tools/vba_lint.py
Cikis kodu 0 = temiz, 1 = hata var.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"

KAPALI_BLOK = {
    "sub": "End Sub",
    "function": "End Function",
    "property": "End Property",
    "if": "End If",
    "with": "End With",
    "for": "Next",
    "do": "Loop",
    "select": "End Select",
    "type": "End Type",
    "enum": "End Enum",
}

ACICI = re.compile(r"^(?:(?:public|private|friend|static)\s+)*(sub|function|property)\s+([a-z_][a-z0-9_]*)", re.I)
BITIRICI = re.compile(r"^end\s+(sub|function|property)\b", re.I)
KAPATICI = re.compile(r"^(end\s+if|end\s+with|next|loop|end\s+select|end\s+type|end\s+enum)\b", re.I)


def yorumlari_ayikla(satir: str) -> str:
    cikti = []
    tirnak = False
    i = 0
    while i < len(satir):
        c = satir[i]
        if c == '"':
            tirnak = not tirnak
            cikti.append(c)
        elif c == "'" and not tirnak:
            break
        else:
            cikti.append(c)
        i += 1
    s = "".join(cikti)
    if re.match(r"^\s*rem\b", s, re.I):
        return ""
    return s


def ifadeler(satir: str) -> list[str]:
    """':' ile ayrilmis ifadeleri dondurur (string icindeki ':' haric)."""
    out, buf, tirnak = [], [], False
    for c in satir:
        if c == '"':
            tirnak = not tirnak
            buf.append(c)
        elif c == ":" and not tirnak:
            out.append("".join(buf))
            buf = []
        else:
            buf.append(c)
    out.append("".join(buf))
    return [s.strip() for s in out if s.strip()]


def denetle(dosya: Path) -> list[str]:
    hatalar: list[str] = []
    metin = dosya.read_text(encoding="utf-8", errors="replace")
    satirlar = metin.splitlines()

    # Mantiksal satirlar: "_" ile devam eden satirlari birlestir
    mantiksal: list[tuple[int, str]] = []
    tampon = ""
    tampon_no = 0
    for no, ham in enumerate(satirlar, start=1):
        if len(ham) > 1023:
            hatalar.append(f"{dosya.name}:{no}: satir 1023 karakterden uzun")
        s = yorumlari_ayikla(ham).rstrip()
        if not tampon:
            tampon_no = no
        if s.endswith("_") and not s.endswith('"_'):
            tampon += s[:-1] + " "
            continue
        tampon += s
        mantiksal.append((tampon_no, tampon))
        tampon = ""
    if tampon:
        mantiksal.append((tampon_no, tampon))

    yigin: list[tuple[str, int]] = []
    prosedur: str | None = None
    bildirimler: dict[str, int] = {}
    adlar: dict[str, int] = {}

    for no, mantik in mantiksal:
        if not mantik.strip():
            continue

        for ifade in ifadeler(mantik):
            d = ifade.strip()
            dl = d.lower()

            m = ACICI.match(d)
            if m and not dl.startswith("declare"):
                tur = m.group(1).lower()
                ad = m.group(2)
                prosedur = no
                bildirimler = {}
                prosedur_adi = (ad.lower(), no)
                if ad.lower() in adlar:
                    hatalar.append(
                        f"{dosya.name}:{no}: '{ad}' proseduru {adlar[ad.lower()]} satirinda da tanimli")
                adlar[ad.lower()] = no
                yigin.append((tur, no))
                continue

            if BITIRICI.match(d):
                tur = BITIRICI.match(d).group(1).lower()
                if not yigin or yigin[-1][0] != tur:
                    hatalar.append(f"{dosya.name}:{no}: 'End {tur}' beklenmeyen yerde "
                                   f"(yigin: {yigin[-3:]})")
                    if yigin:
                        yigin.pop()
                else:
                    yigin.pop()
                prosedur = None
                continue

            if re.match(r"^else\b", dl) or dl.startswith("elseif "):
                if not any(t == "if" for t, _ in yigin):
                    hatalar.append(f"{dosya.name}:{no}: '{d}' If blogu disinda")
                continue

            m = KAPATICI.match(d)
            if m:
                kelime = m.group(1).lower()
                beklenen = {
                    "end if": "if",
                    "end with": "with",
                    "next": "for",
                    "loop": "do",
                    "end select": "select",
                    "end type": "type",
                    "end enum": "enum",
                }[re.sub(r"\s+", " ", kelime)]
                if not yigin or yigin[-1][0] != beklenen:
                    hatalar.append(f"{dosya.name}:{no}: '{d}' beklenmeyen yerde "
                                   f"(beklenen kapanis: {yigin[-1][0] if yigin else 'yok'})")
                    continue
                if beklenen == "for" and d.count(",") >= 1:
                    # Next i, j  -> birden fazla For kapatir
                    for _ in range(d.count(",")):
                        if yigin and yigin[-1][0] == "for":
                            yigin.pop()
                yigin.pop()
                continue

            if re.match(r"^(for|for\s+each)\b", dl):
                yigin.append(("for", no))
                continue
            if re.match(r"^do\b", dl):
                yigin.append(("do", no))
                continue
            if re.match(r"^with\b", dl):
                yigin.append(("with", no))
                continue
            if re.match(r"^select\s+case\b", dl):
                yigin.append(("select", no))
                continue
            if re.match(r"^(public\s+|private\s+)?type\s+[a-z_]", dl) and not dl.startswith("type "):
                yigin.append(("type", no))
                continue
            if re.match(r"^(public\s+|private\s+)?enum\b", dl):
                yigin.append(("enum", no))
                continue

            if re.match(r"^if\b", dl):
                # tek satirlik If ... Then x = 1  -> blok acmaz
                if re.search(r"\bthen\b\s*\S", dl):
                    continue
                yigin.append(("if", no))
                continue

            # Dim / Static / Const bildirimleri (prosedur icinde)
            m = re.match(r"^(?:dim|static)\s+(.+)$", dl)
            if m and prosedur is not None:
                for parca in re.split(r",(?![^(]*\))", m.group(1)):
                    ad_m = re.match(r"([a-z_][a-z0-9_]*)", parca.strip())
                    if not ad_m:
                        continue
                    ad = ad_m.group(1)
                    if ad in bildirimler:
                        hatalar.append(
                            f"{dosya.name}:{no}: '{ad}' ayni prosedurde tekrar Dim "
                            f"(once {bildirimler[ad]}. satir)")
                    else:
                        bildirimler[ad] = no

    if yigin:
        hatalar.append(f"{dosya.name}: kapatilmayan blok(lar): {yigin}")

    return hatalar


def main() -> int:
    dosyalar = sorted(SRC.glob("*.bas"))
    if not dosyalar:
        print(f"Kaynak bulunamadi: {SRC}")
        return 1

    tum: list[str] = []
    print(f"Denetlenen dosya: {len(dosyalar)}")
    for d in dosyalar:
        h = denetle(d)
        bina = sum(1 for s in d.read_text(encoding="utf-8").splitlines())
        print(f"  {d.name:28s} {bina:5d} satir   {'OK' if not h else str(len(h)) + ' hata'}")
        tum.extend(h)

    print()
    if tum:
        for h in tum:
            print("HATA " + h)
        return 1
    print("Temiz: blok dengesi, tekrar Dim ve prosedur adlari sorunsuz.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
