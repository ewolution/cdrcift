# Kullanım — JavaScript sürümü

## İş akışı

```
Dosyayı aç  →  CiftCizgi.js  →  özet + işaretler  →  (onay) sil  →  Ctrl+Z
```

### 1. Çalıştırma

`Araçlar > Makrolar > JavaScript` → `CiftCizgi` (veya Scripts paneli).

Script önce sayfayı tarar, sonra tek bir özet penceresi gösterir:

```
TAM (birebir)      : 7
KISMI (ortak kenar): 12
YAKIN (kaymis)     : 3
NOKTAS             : 0
TOPLAM             : 22
Silinebilir TAM    : 5

Taranan: 1840 nesne / 2130 alt yol / 58211 segment
Isaretlenen: 22 (DUP_CHECK katmani)
Rapor: C:\...\Temp\cdrcift\rapor-20261028-104512.txt
```

### 2. Rapor

`%TEMP%\cdrcift\rapor-YYYYMMDD-HHNNSS.txt` (+ `.csv`)

TXT kolonları:

| Kolon | Anlamı |
|---|---|
| `Tur` | TAM / KISMI / YAKIN / NOKTAS |
| `Katman` | Nesnenin katmanı |
| `Nesne` | Nesne adı (yoksa `Nesne <StaticID>`) |
| `AltYol` | Curve içindeki alt yol numarası |
| `Uzunluk mm` | Çakışan toplam uzunluk |
| `Oran %` | Çakışma / kısa alt yolun uzunluğu |
| `Konum X,Y` | Çakışma merkezi (mm) |
| `Eylem` | `sil` = onayla silinebilir, `elle` = kullanıcı kararı |
| `Eslesen` | Korunan (referans) nesne |

Numara = `DUP_CHECK` katmanındaki işaret numarası.

> Dosya yazma bu CorelDRAW JavaScript host'unda **yok** (`host.CreateObject` ve
> `ActiveXObject` tanımlı değil). Bu yüzden tam rapor, sayfanın sağına
> `DUP_CHECK_RAPOR` katmanına metin nesnesi olarak yazılır; `%TEMP%` dosyası
> oluşmaz. Özet + ilk 10 bulgu `alert` penceresinde gelir.

### 3. Temizlik

- `otomatikSil: true` → TAM (birebir) kopyalar silinir.
- `ortakKenarSil: true` → ızgara/karosel tasarımlarında **ortak kenarlar tek çizgiye** iner:
  kenarı paylaşan karelerden biri o kenarı kaybeder, kalan kenarlar yeni bir eğri
  olarak yeniden kurulur. İç hücrelerin tüm kenarları paylaşılıyorsa hücre tamamen kalkar.
- **Güvenlik kuralları** (logo/harf bozulmasını önler): yalnız
  farklı nesnelerin, **kapalı**, **4 kenarlı** (ayarlanabilir), **tek alt yollu**
  hücreleri işlenir; eğri içeren nesnelere dokunulmaz. Harf, yıldız, birleşik
  şekiller bu yüzden korunur.
- Logo/harf ağırlıklı dosyalarda bu özelliği kapatın: `ortakKenarSil: false`.
- Yalnızca `TAM` + aynı katman + aynı renk + aynı kalınlık silinir.
- Tümü tek `BeginCommandGroup` içinde → **tek Ctrl+Z**.
- `KISMI`/`YAKIN`/`NOKTAS` raporlanır; otomatik silinmez.

### 4. İşaretler ve rapor nesnesi

- Her çalıştırmada eski `DUP_CHECK` katmanı silinip yenisi oluşturulur.
- **Dosya yazma bu host'ta yok** (`host.CreateObject` / `ActiveXObject` yok).
  Bu yüzden tam rapor, sayfanın sağına `DUP_CHECK_RAPOR` katmanına **metin nesnesi**
  olarak yazılır (yazdırılamaz). Metni seçip kopyalayabilir veya dışa aktarabilirsiniz.
- Özet + ilk 10 bulgu `alert` penceresinde gelir; `alert` 3000+ karakter gösterir.
- Katmanlar `Printable = false` yapılmaya çalışılır; yapılamazsa özette uyarı çıkar.

## Tolerans rehberi

| Ayar | Varsayılan | Ne yapar | Öneri |
|---|---|---|---|
| `tolerans` | 0,05 mm | Birebir eşiği | Plotter 0,1 mm altını tek çizgi kabul eder; 0,05 iyi |
| `yakinTolerans` | 0,5 mm | Kaymış kopya eşiği | Kasıtlı çift çerçeve varsa 0,2'ye indirin |
| `minOrtusme` | 0,5 mm | Raporlanacak en kısa çakışma | 1,0 yaparsanız mikro çakışmalar susar |
| `kapsam` | `sayfa` | `secim` = yalnız seçili nesneler | Büyük dosyada önce seçip `secim` kullanın |
| `otomatikSil` | `true` | TAM kopyaları sil | İlk günlerde kapalı tutun |
| `ortakKenarSil` | `true` | Ortak kenarları tek çizgiye indir (ızgara/karosel) | Logo/harf dosyalarında **kapatın** |
| `ortakKenarAzamiKenar` | 4 | Yalnız bu kadar kenarlı kapalı hücreler (kare=4) | Altıgen için 6 |
| `ortakKenarTekAltYol` | `true` | Yalnız tek alt yollu nesneler | Birleşik şekilleri korur |
| `ortakKenarMax` | 200 | Tek çalıştırmada yeniden kurulacak en fazla nesne | Büyük işlerde artırın |
| `hucreMm` | 2,0 | Izgara hücresi | Küçültmek yavaşlatır; büyütmek 0,5 mm üstü kaymaları kaçırabilir |
| `yoksayKatmanlar` | `[]` | Atlanacak katmanlar | `["Regmark"]` gibi |
| `boya` | `false` | Silinecekleri kırmızıya boya | Görsel kontrol için |
| `ondalik` | `,` | CSV ondalık ayırıcı | İngilizce Excel: `.` |

## Sık sorulanlar

**Kısmi çakışmayı neden silmiyor?**
Ortak kenar, bitişik iki parçanın gerçek kesim hattıdır. Silmek parçayı bozar.
Araç işaretler; karar sizde.

**Dikdörtgen/elips dönüştürmem gerekiyor mu?**
Hayır. Dönüştürülmemiş şekiller de taranır (`DisplayCurve` üzerinden).

**Hangi nesneler atlanır?**
Bitmap, metin, blend/contour/extrude/bevel/gölge, mesh fill. Raporda sayısı yazar.
Metni `Ctrl+Q` ile eğriye çevirirseniz taranır.

**Performans**
5.000 nesne / ~60.000 segment hedefi < 60 sn. Daha hızlısı: `kapsam: "secim"`,
`hucreMm: 4`, veya gerekiyorsa `isaretSiniri: 100`.

**Türkçe karakterler**
Script cp1254 (ANSI) kurulur; mevcut Summa scriptlerinizle aynı. UTF-8
kaydederseniz bozulur.

## Komut özeti

| Dosya | İş |
|---|---|
| `CiftCizgi.js` | Tara + raporla + işaretle + (onayla) temizle |

## VBA sürümü

`legacy/vba/src/*.bas` aynı algoritmanın VBA karşılığıdır; derleme hatası nedeniyle
bakımda değildir (yalnız referans).
