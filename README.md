# cdrcift — CorelDRAW Çift Çizgi Denetleyicisi

Plotter/kesim dosyalarında **üst üste binen ve ortak kullanılan hatları** bulan,
işaretleyen ve güvenli olanları tek geri-alma adımında temizleyen CorelDRAW aracı.

Amaç: plotterin aynı hattı iki kez kesip parçayı ayırmasını önlemek; ızgara/karosel
tasarımlarında ortak kenarları **tek çizgiye** indirmek.

```
CorelDRAW 2024 / 2025 / 2026   ·   JavaScript (V8)   ·   kurulum gerektirmez
```

## Ne bulur, ne yapar?

| Tür | Anlamı | Otomatik işlem |
|---|---|---|
| **TAM** | Birebir kopya (kopyala-yapıştır, ters yön, farklı başlangıç düğümü) | ✅ Tek kopya kalır |
| **KISMI** | İki hattın bir bölümü ortak (bitişik parçaların ortak kenarı) | ✅ `ortakKenarSil` açıksa tek çizgiye iner |
| **YAKIN** | Tolerans kadar kaymış kopya | ❌ Yalnız raporlanır |
| **NOKTAS** | Sıfır uzunluklu segment / üst üste düğüm | ❌ Yalnız raporlanır |
| **Ortak kenar** | Izgara/karosel hücrelerinin paylaştığı kenar | ✅ Kalır: bir kez |

## Kurulum

### Tek tıkla (Windows, önerilen)

Depoyu indirip **`tools\KUR.cmd`** dosyasına çift tıklayın.

- Kurulu **tüm** CorelDRAW sürümlerini (2024/2025/2026) bulur,
- `Draw\Scripts` klasörüne **cp1254** kodlamasıyla kopyalar,
- Yönetici hakkı, Python veya kayıt defteri değişikliği **gerektirmez**.

Komut satırından:

```powershell
powershell -ExecutionPolicy Bypass -File tools\kur.ps1
powershell -ExecutionPolicy Bypass -File tools\kur.ps1 -Surum 2026
powershell -ExecutionPolicy Bypass -File tools\kur.ps1 -Kaldir
```

Kaldırma: **`tools\KALDIR.cmd`** (veya `kur.ps1 -Kaldir`).

### Python ile (geliştirici)

```powershell
python tools\js_kur.py
```

Scriptleri aynı klasöre yazar:

```
%APPDATA%\Corel\CorelDRAW Graphics Suite 2026\Draw\Scripts\CiftCizgi.js
```

### Elle

`js/CiftCizgi.js` dosyasını yukarıdaki klasöre kopyalayın.
**Not Defteri ile açıp "ANSI" olarak kaydedin** (UTF-8 kaydederseniz Türkçe bozulur).

## Kullanım

1. CorelDRAW'da dosyayı açın.
2. `Araçlar > Makrolar > JavaScript` → **`CiftCizgi`** (veya Scripts paneli).
3. Özet penceresi gelir; bulgular `DUP_CHECK` katmanında numaralı işaretlenir
   (katman yazdırılamaz → plotter'a gitmez).
4. Silme/iyileştirme aynı komut grubunda yapılır → **tek `Ctrl+Z`** hepsini geri alır.

> Bu host'ta dosya yazma ve `confirm` penceresi yok. Ayrıntılı rapor istenirse
> `AYAR.raporMetni = true` ile belgeye, sayfanın sağına parçalı metin olarak yazılır.

### Ayarlar

`js/CiftCizgi.js` başındaki `AYAR` nesnesi; düzenledikten sonra yeniden kurun.

```js
var AYAR = {
  tolerans: 0.05,          // birebir eşiği (mm)
  yakinTolerans: 0.5,      // kaymış kopya eşiği (mm)
  minOrtusme: 0.5,         // raporlanacak en kısa çakışma (mm)
  kapsam: "sayfa",         // "sayfa" | "secim"
  otomatikSil: true,       // TAM kopyaları sil
  ortakKenarSil: true,     // ortak kenarları tek çizgiye indir
  isaretSiniri: 150,       // en fazla kaç bulgu işaretlenir
  yoksayKatmanlar: [],     // örn. ["Regmark"]
  ondalik: ","             // CSV/rapor ondalık ayırıcı
};
```

| Ayar | Öneri |
|---|---|
| `otomatikSil` | İlk denemede `false`; sonucu görüp `true` yapın |
| `ortakKenarSil` | Izgara/karosel dosyalarında `true`; **logo/harf ağırlıklı dosyalarda `false`** |
| `ortakKenarAzamiKenar` | Kare hücre = 4, altıgen = 6 |
| `ortakKenarTekAltYol` | `true` → birleşik şekiller (harfler) korunur |

## Nasıl çalışır?

1. **Belge dolaşma** — grup ve PowerClip içine iner; bitmap/metin/efekt nesnelerini atlar.
2. **Geometri** — `DisplayCurve` üzerinden alt yollar okunur (parametrik şekillerde
   `Curve` boş döner), Bézier eğrileri düzleştirilir, tüm hesap **mm** cinsindendir.
3. **TAM** — düğüm düzeyinde **kanonik imza**: ters yön ve farklı başlangıç düğümü aynı
   imzayı verir; 0,005 mm ızgara sayesinde 0,01 mm kayma yakalanmaz.
4. **Segment taraması** — uzamsal ızgara (O(n²) yerine ~O(n)); paralellik, dik uzaklık ve
   1B örtüşme ile TAM/KISMI/YAKIN sınıflandırması.
5. **Ortak kenar** — tam çakışan düz segmentler; **yalnız** farklı nesnelerin, kapalı,
   4 kenarlı, tek alt yollu hücrelerinde uygulanır. Kenarı kaybeden hücre, kalan kenarlarla
   yeniden kurulur (renk/kalınlık kopyalanır).
6. **Rapor + işaret** — `DUP_CHECK` katmanı, numaralı daireler; istenirse metin raporu.

## Güvenlik

- TAM silme yalnız **aynı katman + aynı renk + aynı kalınlık** ise yapılır.
- `YAKIN` / `NOKTAS` hiçbir koşulda otomatik silinmez.
- Ortak kenar yeniden kurma: beklenen = üretilen segment sayısı doğrulanır; tutmazsa
  nesneye **dokunulmaz**.
- Harf, yıldız, birleşik şekiller (çok kenarlı / çok alt yollu / eğrili) korunur.
- Tüm işlemler tek `BeginCommandGroup` içinde → **tek `Ctrl+Z`**.
- `DUP_CHECK` katmanı `Printable = false` yapılmaya çalışılır.

## Testler

```powershell
node js\test_ciftcizgi.js       # 41 test — tüm hat, CorelDRAW mock'u ile
python tools\test_geom.py       # 15 test — bağımsız Python referansı
python tools\plt_dedup.py --test
```

`js/test_ciftcizgi.js` gerçek bir CorelDRAW nesne modeli taklidi kurar ve
dolaşma → düzleştirme → imza → ızgara → sınıflandırma → yeniden kurma hattının
tamamını koşturur (2×2 ızgara, üst üste kareler, harf koruma senaryoları dahil).

`tools/coredraw_smoke.ps1` ise **gerçek CorelDRAW** üzerinde 29 API varsayımını
doğrular (görünmez örnek; açık belgelere dokunmaz).

## Depo yapısı

```
js/
  CiftCizgi.js           Tek dosyalık araç
  test_ciftcizgi.js      Node testleri (mock ile 41 test)
tools/
  kur.ps1                Tek tıkla kurulum / kaldırma (Windows)
  KUR.cmd                Çift tıklama kısayolu
  KALDIR.cmd             Kaldırma kısayolu
  js_kur.py              Alternatif kurulum (UTF-8 -> cp1254)
  geom_referans.py       Algoritmanın Python referansı
  test_geom.py           Referans testleri (15)
  coredraw_smoke.ps1     CorelDRAW API doğrulaması (29 kontrol)
  plt_dedup.py           PLT/HPGL mükerrer hat denetleyicisi
docs/
  KURULUM.md  KULLANIM.md  TEST-PLANI.md  API-NOTLARI.md
legacy/vba/              Deneysel VBA sürümü (bakımda değil, referans)
```

## Bilinen sınırlar

- Dosya yazma ve `confirm` bu JavaScript host'unda yok; rapor belge içine yazılır.
- Eğri içeren hücreler ortak kenar temizliğinde `elle` bırakılır.
- Metin nesneleri taranmaz (`Ctrl+Q` ile eğriye çevrilmiş metin taranır).
- 5.000 nesne / ~60.000 segment hedefi < 60 sn (`kapsam: "secim"` ile daha hızlı).
- CorelDRAW 2024/2025 üzerinde otomatik test edilmedi; aynı V8 host'u kullanır.

## Katkı

Issues ve pull request'ler açıktır. Algoritma değişecekse **üç uygulama**
(`js/CiftCizgi.js`, `tools/geom_referans.py`) ve testler birlikte güncellenmelidir.

## Lisans

MIT — bkz. [LICENSE](LICENSE).
