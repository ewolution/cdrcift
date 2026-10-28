# API Notları — doğrulanmış CorelDRAW VBA yüzeyi

Bu dosya tahmin değil **kanıt** içerir. Yöntem: bu makinede kurulu
`CorelDRAW.tlb` tip kütüphanesi `comtypes.client.GetModule` ile açıldı, üye
imzaları ve `cdr*` sabitleri doğrudan okundu.

```
C:\Program Files\Corel\CorelDRAW Graphics Suite\27\Programs64\TypeLibs\CorelDRAW.tlb
CorelDRAW Graphics Suite 2024 - VBA (x64)   25.1     (kayıtlı)
CorelDRAW Graphics Suite 2025 - VBA (x64)   26.1     (kayıtlı)
CorelDRAW Graphics Suite 2026 - VBA (x64)   27.0     (kayıtlı)
VGCoreAuto.tlb / CorelDRAW.tlb               mevcut
```

## 1. Ortam

| | |
|---|---|
| GMS klasörleri | `%APPDATA%\Corel\CorelDRAW Graphics Suite 2024\Draw\GMS` (boş) |
| | `%APPDATA%\Corel\CorelDRAW Graphics Suite 2025\Draw\GMS` (boş) |
| | `%APPDATA%\Corel\CorelDRAW Graphics Suite 2026\Draw\GMS` (boş) |
| Not | Yol **sürüm adıyla** klasörlenir, `...\27\...` değil. |
| Kullanıcı profili | `C:\Users\Evrim İş\...` — **Türkçe karakterli yol**. Rapor/ini için `%TEMP%` ve `%APPDATA%` kullanıldı (ASCII alt klasör `cdrcift`). |

## 2. Temizlik için kritik üyeler (hepsi doğrulandı)

| Üye | İmza / not | Kullanım |
|---|---|---|
| `SubPath.Delete` | `()` | Alt yol silme **var** → yeniden inşa gerekmez |
| `Node.Delete` | `()` | (kullanılmadı, yedek) |
| `Segment.Type` | `propget Long` — 0 = çizgi, 1 = eğri | düzleştirme |
| `Segment.StartNode` / `EndNode` | `Node` | uç noktalar |
| `Segment.StartingControlPointX/Y`, `EndingControlPointX/Y` | `Double` | Bezier |
| `Shape.Delete` | `()` | nesne silme |
| `Shape.StaticID` | `Long` | kimlik |
| `Shape.ZOrder` | `Long` | alt/üst sıra (korunan kopya seçimi) |
| `Shape.Layer` | `Layer` | güvenlik kuralı |
| `Shape.Outline.Width` / `.Color` | `Double` / `Color` | güvenlik kuralı |
| `Color.IsSame(other)` | `Boolean` | renk eşitliği |
| `Color.HexValue` | `BSTR` | rapor |
| `Document.BeginCommandGroup(name)` / `EndCommandGroup` | | tek Ctrl+Z |
| `Application.Optimization` / `EventsEnabled` | `Boolean` | hız |
| `Application.Documents.Count`, `ActiveDocument` | | belge kontrolü |
| `Document.Unit` | `cdrUnit` | **birim tuzağı** (aşağıda) |

## 3. Birim tuzağı (bu projede önemli)

CorelDRAW object model koordinatları **belge biriminde** döner. Plotter dosyaları
sıklıkla inç veya punto olarak açılır. 0,05 mm tolerans, birim çevrilmezse
inç belgede 0,05 inç = 1,27 mm olur (25× yanlış).

`Document.Unit` → `cdrUnit` sabitleri (tlb'den):

```
cdrTenthMicron=0  cdrInch=1      cdrFoot=2       cdrMillimeter=3
cdrCentimeter=4   cdrPixel=5     cdrMile=6       cdrMeter=7
cdrKilometer=8    cdrDidots=9    cdrAgate=10     cdrYard=11
cdrPica=12        cdrCicero=13   cdrPoint=14     cdrUnitQ=15
cdrUnitH=16
```

`CiftCizgiGeom.BirimKatsayi()` bu tabloyu mm katsayısına çevirir. `cdrPixel`
çözünürlüğe bağlı olduğundan 1 kabul edilir ve rapora **uyarı** yazılır.

## 4. Nesne tipleri (`cdrShapeType`)

```
cdrNoShape=0            cdrRectangleShape=1   cdrEllipseShape=2
cdrCurveShape=3         cdrPolygonShape=4     cdrBitmapShape=5
cdrTextShape=6          cdrGroupShape=7       cdrSelectionShape=8
cdrGuidelineShape=9     cdrBlendGroupShape=10 cdrExtrudeGroupShape=11
cdrOLEObjectShape=12    cdrContourGroupShape=13
cdrLinearDimensionShape=14  cdrBevelGroupShape=15
cdrDropShadowGroupShape=16  cdr3DObjectShape=17
cdrArtisticMediaGroupShape=18  cdrConnectorShape=19
cdrMeshFillShape=20     cdrCustomShape=21     cdrCustomEffectGroupShape=22
```

Kararlar:

- **7** (grup) → özyinelemeli inilir.
- **10, 11, 13, 15, 16, 18, 22** (efekt grupları) → atlanır (üretilmiş geometri,
  gerçek kesim hattı değil), sayılır.
- **5, 9, 12, 17, 20** (bitmap, kılavuz, OLE, 3D, mesh) → atlanır.
- **6** (metin) → atlanır. `Ctrl+Q` ile eğriye çevrilmiş metin **3** olur ve taranır.

## 5. Koordinat uzayı ve grup içi nesneler

- `Shape.Curve` çoğu durumda belge uzayında döner; ancak grup içinde dönüşüm
  uygulanmış nesnelerde beklenenden farklı olabilir.
- `Shape.DisplayCurve` (doğrulandı) görüntü uzayındaki eğriyi verir.
- Bu proje: grup içinde **DisplayCurve** birincil, `Curve` yedek; ardından
  `Curve.BoundingBox` ile `Shape.BoundingBox` karşılaştırılır. Uyuşmazsa nesne
  **atlanır ve uyarı yazılır** (sessiz yanlış sonuç yok).
- `Rect` üyeleri: `Left`, `Top`, `Bottom`, `Right`, `Width`, `Height`,
  `CenterX`, `CenterY` (y = yukarı).

## 6. PowerClip

`Shape.PowerClip` → `PowerClip.Shapes`, `EnterEditMode`, `LeaveEditMode`,
`ExtractShapes`. Kap ve içerik ayrı ayrı taranır; erişim hatasında
`gPowerClipAtlanan` sayacı artar.

## 7. İşaret katmanı

| Üye | Not |
|---|---|
| `Page.CreateLayer(name)` | katman oluşturur |
| `Layer.Printable` | **False** → işaretler plotter'a gitmez |
| `Layer.Editable`, `Visible`, `Delete` | kilit açma / silme |
| `Layer.CreateEllipse(l, t, r, b)` | daire işareti |
| `Layer.CreateArtisticText(x, y, text)` | numara |
| `Shape.Fill.ApplyNoFill` | içi boş daire |
| `Text.FontProperties.Size` | punto |

## 8. Gerçek CorelDRAW 2026 duman testi — 29/29 geçti

`tools/coredraw_smoke.ps1` çalıştırıldı (yeni, görünmez CorelDRAW örneği;
kullanıcının açık belgesine dokunulmadı). Doğrulanan davranışlar:

| Doğrulama | Sonuç | Kod üzerindeki etkisi |
|---|---|---|
| Birim tuzağı | 1 inç'lik çizgi `Unit=1`'de 1, `Unit=3`'te **25,4** | `BirimKatsayi` şart |
| **Parametrik şekil** | Dikdörtgen/elips/poligonda `Shape.Curve.SubPaths = 0`, `DisplayCurve.SubPaths = 1` | **Kod düzeltildi**: geometri `DisplayCurve`'den |
| Segment tipleri | Dikdörtgen → 0 (çizgi), elips → 1 (eğri) | `Duzlestir` doğru |
| `SubPath.Delete` | 2 → 1 alt yol | alt-yol temizliği mümkün |
| Son alt yolu silmek | Nesne boşalıyor/ölüyor | `SubPaths.Count > 1` koruması |
| `ZOrder` | `Shapes(1)` en üstte, **Z = 1**; büyük Z = altta | tutucu = **max Z** (en alttaki korunur) |
| Grup + `Rotate(30)` | `Curve` = `DisplayCurve` = nesne bbox (sapma **0,000**) | yedek mantığı zararsız |
| Bbox farkı (düz şekil) | 0,0002 mm | 1 mm tolerans yeterli |
| `Layer.Printable` | Yeni katmanda varsayılan **True** | her seferinde açıkça `False` |
| `Color.IsSame` / `HexValue` / `RGBAssign` / `Fill.ApplyNoFill` | Çalışıyor (#FF0000) | güvenlik kuralı + işaretler |
| `BeginCommandGroup` + `Undo(1)` | 8 → 7 nesne | tek Ctrl+Z |
| `FontProperties.Size` | PowerShell'den okunamadı (PSParameterizedProperty) | VBA'da hata korumalı; başarısız olursa varsayılan 12 pt kalır |

Tekrar çalıştırma:

```powershell
powershell -ExecutionPolicy Bypass -File tools\coredraw_smoke.ps1
```

## 9. Bilinen belirsizlikler (kodda savunmalı)

| Konu | Durum | Savunma |
|---|---|---|
| `SubPath.Delete` son alt yolda ne yapar? | 2026'da **nesne boşalıyor/ölüyor** (ölçüldü) | `SubPaths.Count > 1` kontrolü; tek alt yollu nesne tamamen silinir |
| `Shape.Curve` parametrik şekillerde | **boş** (ölçüldü) | `DisplayCurve` birincil, `Curve` yedek, ikisi de boşsa uyarı + atla |
| `Shape.Curve` grup içinde uzay | 2026'da doğru (sapma 0) | yine de bbox doğrulaması + DisplayCurve yedeği |
| `Application.StatusBar.Text` | `StatusBar` bir `ICUICommandBar` döner; `.Text` garanti değil | `On Error Resume Next` ile sessiz geçer |
| `Layer.Printable` varsayılanı | **True** (ölçüldü) | her işaretlemede açıkça `False` |
| `Text.FontProperties.Size` | VBA'da denenmeli (PowerShell erişemedi) | hata korumalı; başarısızsa varsayılan punto |
| `Curve.GetPolyline(precision, out)` | Var; ama alt yol başına nesne ayırır | kullanılmadı, manuel Bezier düzleştirme (hızlı) |
| `Curve.RemoveOverlaps(out)` | Var; davranışı belgelenmemiş | kullanılmadı (kendi tespitimiz denetlenebilir) |

## 9. Teşhis makrosu

`CiftCizgiTest.AltYolDok` — seçili nesneler için birim, bbox (curve vs nesne),
alt yol/segment/düğüm sayısı, ilk nokta ve üretilen imzayı **Hemen Penceresi**'ne
yazar. Yeni bir sürümde tuhaflık görürseniz ilk çalıştırılacak komut budur.

## 10. Yeniden doğrulama

Bir sürüm yükseltmesinden sonra:

```powershell
python -c "import comtypes.client as cc; cc.GetModule(r'C:\Program Files\Corel\CorelDRAW Graphics Suite\27\Programs64\TypeLibs\CorelDRAW.tlb')"
```

Ardından `CiftCizgiTest.TestleriCalistir` + `AltYolDok`.
