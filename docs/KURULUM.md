# Kurulum — JavaScript sürümü (önerilen)

> Yönetici hakkı, GMS kaydı veya makro güvenliği ayarı gerekmez.

## 1. Otomatik kurulum (bu makinede yapıldı)

```powershell
python tools\js_kur.py
```

Scriptleri şuraya yazar (mevcut Summa scriptlerinin yanına):

```
%APPDATA%\Corel\CorelDRAW Graphics Suite 2026\Draw\Scripts\
    CiftCizgi.js
```

Kodlama **cp1254** (ANSI): CorelDRAW'ın JavaScript host'u, mevcut Summa
scriptlerinizle aynı kodlamayı bekliyor. Araç UTF-8 kaynağı çevirip geri
okuyarak doğrular.

Başka bir sürüm için:

```powershell
python tools\js_kur.py --surum "Graphics Suite 2025"
```

## 2. Elle kurulum (gerekirse)

`js\CiftCizgi.js` dosyasını yukarıdaki klasöre kopyalayın. **Not Defteri ile açıp
"ANSI" olarak kaydedin** (UTF-8 kaydederseniz Türkçe karakterler bozulur).

## 3. Çalıştırma

CorelDRAW'da:

- `Araçlar > Makrolar > JavaScript` → `CiftCizgi` (veya dosya seç)
- ya da Scripts paneli / `Araçlar > Scriptler`

İlk çalıştırmada `AYAR.otomatikSil = false` yapmanız önerilir: hiçbir şey
silinmez, yalnız rapor + işaret gelir. Sonuçları görüp onayladıktan sonra
`true` yapıp yeniden kurun.

Sonra `CiftCizgi.js`:

1. Rapor + işaretleme yapar, özet gösterir.
2. `AYAR.otomatikSil = true` ise "silinsin mi?" diye sorar.
3. Onaylarsanız TAM kopyalar tek command group'ta silinir → **tek Ctrl+Z**.

## 4. Ayarlar

`CiftCizgi.js` dosyasının başındaki `AYAR` nesnesini düzenleyin, sonra
`python tools\js_kur.py` ile yeniden kurun.

```js
var AYAR = {
  tolerans: 0.05,        // birebir eşiği (mm)
  yakinTolerans: 0.5,    // kaymış kopya eşiği (mm)
  minOrtusme: 0.5,       // raporlanacak en kısa örtüşme (mm)
  kapsam: "sayfa",       // "sayfa" | "secim"
  otomatikSil: false,    // true: onay sonrası TAM sil
  boya: false,           // silinecekleri kırmızıya boya
  isaretKatmani: "DUP_CHECK",
  isaretSiniri: 300,
  yoksayKatmanlar: [],   // örn. ["Regmark"]
  hucreMm: 2.0,
  ondalik: ","           // CSV ondalık ayırıcı (TR Excel)
};
```

## Sorun giderme

| Belirti | Neden / çözüm |
|---|---|
| Script listede görünmüyor | Yanlış sürüm klasörü. `--surum "Graphics Suite 2025"` ile kurun veya CorelDRAW'ı yeniden başlatın. |
| `alert` içinde `Ã§`, `ÅŸ` | Dosya UTF-8 kaydedilmiş. Not Defteri'nde ANSI olarak kaydedin. |
| "Cift Cizgi hata verdi: ..." | `AYAR.asama = "tara"` yapıp yeniden kurun; hata mesajındaki satırı bana iletin. |
| Rapor dosyası yok | JS host'unda dosya erişimi yok. Özet `alert` içinde gelir; ayrıntılı metin için `AYAR.raporMetni = true` (parçalı yazılır). |
| İşaretler plotter'a gidiyor | `DUP_CHECK` katmanını elle "yazdırılamaz" yapın (script bunu dener ve raporda uyarır). |
| Silme sonrası çizgi kaybı | `Ctrl+Z` (tek adım). Kalıcıysa dosyayı kaydetmeyin. |

## Kaldırma

`Draw\Scripts\CiftCizgi.js` dosyasını silin.

---

# Ek: VBA (GMS) sürümü — deneysel

`legacy/vba/src/*.bas` dosyaları aynı algoritmanın VBA karşılığıdır (bakımda değil).
Kurulum denemesi:

```powershell
powershell -ExecutionPolicy Bypass -File legacy\vba\gms_kur.ps1
```

**Durum:** VBA projesi derlemede modal hata veriyor (çözülmedi). Bu nedenle
JavaScript sürümü birincil ve tek desteklenen sürümdür.

CorelDRAW 2026'da doğrulanan davranışlar (her iki sürüm için geçerli) için
`docs/API-NOTLARI.md`.
