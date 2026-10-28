# Test Planı

İki katman: (A) burada çalıştırılabilen otomatik testler, (B) CorelDRAW içinde
çalıştırılan doğrulamalar.

## A. Otomatik (bu makinede)

```powershell
node js\test_ciftcizgi.js     # 41 test — JS sürümünün TAM boru hattı
python tools\test_geom.py     # 15 test — bağımsız Python referansı
python tools\plt_dedup.py --test
```

### A1. `js/test_ciftcizgi.js` (41 test)

CorelDRAW nesne modelini taklit eden bir **mock** kurar ve gerçek `CiftCizgi.js`
kodunu çalıştırır: dolaşma → düzleştirme → imza → ızgara → sınıflandırma →
oran/eylem. Kapsam:

| # | Test | Ne kanıtlar |
|---|---|---|
| 1–2 | Tarama hatasız, 10 nesne | dolaşma |
| 3–4 | Birebir kopya TAM + `sil` | mutlu yol |
| 5 | Ters yönlü kopya TAM | kanonikleştirme |
| 6–7 | Ortak kenar KISMI + `elle` | bitişik parça korunur |
| 8 | 0,3 mm kaymış YAKIN | tolerans bandı |
| 9 | 0,1 mm kaymış TAM değil | eşik |
| 10–11 | KISMI/YAKIN içinde `sil` yok; sayım 2 TAM | güvenlik değişmezi |
| 12 | Kapalı yol: ters yön + farklı başlangıç TAM | imza |
| 13 | Açık çizgi ters yön TAM | imza |
| 14 | 0,01 mm kaymış dikdörtgen TAM değil | ızgara 0,005 mm |
| 15–16 | İnç belgesi: 1 inç = 25,4 mm | birim tuzağı |
| 17–18 | Farklı katman: TAM bulunur, `elle` kalır | silme güvenliği |
| 19–20 | 75 nesneli rastgele set: tam 15 kopya | ızgara + gürültü |
| 21–22 | (imza/döngü yardımcıları) | — |

`(Geçen: 41, Kalan: 0)`

### A2. `coredraw_smoke.ps1` — 29/29

Gerçek CorelDRAW 2026'da (görünmez örnek) birim, parametrik şekil geometrisi,
`SubPath.Delete`, `ZOrder`, grup dönüşümü, `Printable`, `Undo` doğrulandı.

## B. CorelDRAW içinde (JS)

| # | Senaryo | Ölçüt |
|---|---|---|
| 1 | `AYAR.asama = "tara"` | Yalnız tarama + özet; çökme yok |
| 2 | Demo: aynı çizgi iki kez | 1 TAM, işaretli, raporda |
| 3 | Bitişik iki dikdörtgen | 1 KISMI, `elle` |
| 4 | `AYAR.otomatikSil=false` | hiçbir şey silinmez |
| 5 | `otomatikSil=true` + onay | yalnız TAM silinir, Ctrl+Z geri alır |
| 6 | Farklı katmanlı kopya | silinmez, `elle` |
| 7 | Dönüştürülmemiş dikdörtgen/elips | taranır (atlama yok) |
| 8 | 2–3 gerçek dosya | yanlış pozitif < %5, süre < 60 sn |
| 9 | 5.000 nesneli dosya | kilitlenme yok |
| 10 | 2024 / 2025 / 2026 | kurulum + çalıştırma |

## C. Kabul kriterleri

- [x] `js/test_ciftcizgi.js` 41/41
- [x] `test_geom.py` 15/15
- [x] CorelDRAW API 29/29
- [ ] JS: CorelDRAW içinde ilk çalıştırma (kullanıcı)
- [ ] KISMI/YAKIN hiçbir koşulda otomatik silinmiyor (kod + test)
- [ ] TAM silme tek Ctrl+Z ile geri alınıyor
- [ ] DUP_CHECK katmanı çıktıya gitmiyor
- [ ] Üç sürümde kurulum

## D. Regresyon

Her değişiklikten sonra:

```
node js\test_ciftcizgi.js && python tools\test_geom.py
```

Algoritma değiştiyse **üç** uygulama (JS, Python, VBA) birlikte güncellenmeli;
referans testleri ayrışırsa yanıltıcı olur.


## B. CorelDRAW içinde (kabul testleri)

Çalıştırma: `CiftCizgiTest.TestleriCalistir`
Sonuç: MsgBox + `%TEMP%\cdrcift\test-sonuc.txt`

Fixture (yeni belge, birim = mm):

| # | Geometri | Beklenen |
|---|---|---|
| 1 | Aynı çizgi iki kez `(10,10)-(10,60)` | 1 TAM @ (10,35), `sil` |
| 2 | `(110,10)-(110,60)` + tersi | 1 TAM @ (110,35) |
| 3 | Bitişik iki dikdörtgen, ortak kenar x=240 | 1 KISMI @ (240,30), `elle` |
| 4 | `x=300` ve `x=300,3` çizgileri | 1 YAKIN |
| 5 | `x=320` ve `x=320,1` çizgileri | TAM **yok** (yanlış pozitif testi) |
| 6 | Sıfır uzunluklu curve segment | varsa NOKTAS, yoksa sessiz geçer |
| 7 | `Sil` | ≥2 nesne silinir, nesne sayısı azalır |
| 8 | `Ctrl+Z` (Undo 1) | nesne sayısı geri gelir |

Fixture ayrıca **dönüştürülmemiş dikdörtgen** içerir: `Shape.Curve` boş,
`DisplayCurve` dolu. Tarama bunları atlamamalı (2026'da ölçüldü).

### B2. Gerçek dosya testleri (kullanıcı)

| # | Senaryo | Ölçüt |
|---|---|---|
| 9 | Sorun yaşanan gerçek dosya (2–3 adet) | Tespitler gözle doğrulanır; yanlış pozitif oranı < %5 |
| 10 | 5.000 nesneli dosya | Tarama < 60 sn, CorelDRAW kilitlenmez |
| 11 | İnç birimli dosya | Rapor `1 birim = 25,4 mm` yazar; tolerans mm olarak doğru |
| 12 | Gruplu + döndürülmüş nesneler | Uyarı listesinde "atlandı" yok veya açıklamalı |
| 13 | Metin + bitmap + blend karışık | Atlanan sayıları raporda, çökme yok |
| 14 | PowerClip içinde kopya | Tespit edilir veya "PowerClip atlandı" sayılır |
| 15 | `IsaretleriKaldir` | DUP_CHECK katmanı ve içeriği tamamen gider |
| 16 | Plotter çıktısı | DUP_CHECK katmanı çıktıda görünmez |
| 17 | Türkçe Excel'de CSV | Tek sütuna düşmez, ondalık doğru |
| 18 | 2024 / 2025 / 2026 | Kurulum + `Denetle` + `TestleriCalistir` geçer |

## C. Kabul kriterleri (özet)

- [ ] `test_geom.py` 15/15
- [ ] CorelDRAW `TestleriCalistir` 8/8
- [ ] KISMI/YAKIN hiçbir koşulda otomatik silinmiyor
- [ ] TAM silme tek Ctrl+Z ile geri alınıyor
- [ ] Farklı katman/renk/kalınlıkta kopya silinmiyor (`elle` kalıyor)
- [ ] 5.000 nesne < 60 sn
- [ ] Rapor TXT/CSV Türkçe Excel'de bozulmuyor
- [ ] DUP_CHECK katmanı çıktıya gitmiyor
- [ ] Üç sürümde kurulum doğrulandı

## D. Regresyon

Her değişiklikten sonra:

```
node js\test_ciftcizgi.js && python tools\test_geom.py
```

CorelDRAW tarafında `TestleriCalistir`. Algoritma değiştiyse **hem** Python
referansı **hem** VBA güncellenmeli; ikisi ayrışırsa referans testleri yanıltıcı olur.
