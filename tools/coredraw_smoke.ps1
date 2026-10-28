# cdrcift - CorelDRAW API duman testi (COM ile, VBA olmadan)
#
# Amac: VBA kodunun dayandigi API varsayimlarini gercek CorelDRAW
# uzerinde dogrulamak. YENI bir bos belge acar, islemleri yapar ve
# belgeyi kaydetmeden kapatir. Acik belgelere ve uygulamaya dokunmaz.
#
# Kullanim:  powershell -ExecutionPolicy Bypass -File tools\coredraw_smoke.ps1

$ErrorActionPreference = "Stop"
$script:gecen = 0
$script:kalan = 0

function Sonuc([string]$ad, [object]$deger) {
    Write-Host ("[OK]   {0,-52} {1}" -f $ad, $deger)
    $script:gecen++
}

function Hata([string]$ad, [string]$mesaj) {
    Write-Host ("[HATA] {0,-52} {1}" -f $ad, $mesaj)
    $script:kalan++
}

function Dene([string]$ad, [scriptblock]$blok) {
    try {
        $r = & $blok
        Sonuc $ad $r
        return $r
    } catch {
        Hata $ad $_.Exception.Message
        return $null
    }
}

Write-Host "=== cdrcift CorelDRAW API duman testi ==="

# 1) Calisan ornege baglan; yoksa YENI ornek ac (acik belgelere dokunmamak icin)
$app = $null
$yeniOrnek = $false
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application")
    Sonuc "GetActiveObject(CorelDRAW.Application)" "calisan ornege baglandi"
} catch {
    Write-Host "[bilgi] ROT'ta CorelDRAW yok; yeni ve gorunmez bir ornek aciliyor"
    try {
        $app = New-Object -ComObject CorelDRAW.Application
        $yeniOrnek = $true
        $app.Visible = $false
        Sonuc "New-Object CorelDRAW.Application" "yeni ornek (Visible=False)"
    } catch {
        Hata "Yeni CorelDRAW ornegi" $_.Exception.Message
        exit 2
    }
}

$doc = $null
$page = $null
$lyr = $null
$olusturulan = @()

try {
    # 2) Yeni belge
    $doc = $app.CreateDocument()
    if ($null -eq $doc) { throw "CreateDocument null dondu" }
    Sonuc "CreateDocument" "yeni belge"
    $doc.Unit = 3      # cdrMillimeter
    $page = $doc.ActivePage
    $lyr = $page.ActiveLayer
    Sonuc "Document.Unit = 3 (mm)" $doc.Unit

    # 3) Temel cizim + bbox birimi
    $l1 = $lyr.CreateLineSegment(10, 10, 10, 60)
    $l2 = $lyr.CreateLineSegment(10, 10, 10, 60)
    $olusturulan += $l1, $l2
    $bb = $l1.BoundingBox
    Sonuc "CreateLineSegment + BoundingBox (mm)" ("L={0} B={1} R={2} T={3}" -f $bb.Left, $bb.Bottom, $bb.Right, $bb.Top)

    $cv = $l1.Curve
    $cb = $cv.BoundingBox
    Sonuc "Shape.Curve bbox (mm)" ("L={0} B={1} R={2} T={3}" -f $cb.Left, $cb.Bottom, $cb.Right, $cb.Top)
    Sonuc "SubPaths/Segments/Nodes" ("{0}/{1}/{2}" -f $cv.SubPaths.Count, $cv.Segments.Count, $cv.Nodes.Count)

    # 4) Birim tuzagi: inc'e gecince degerler inc olarak mi geliyor?
    $doc.Unit = 1      # cdrInch
    $inc = $lyr.CreateLineSegment(1, 1, 1, 2)
    $olusturulan += $inc
    $ibb = $inc.BoundingBox
    Sonuc "Unit=1 (inc) iken bbox yuksekligi" ([math]::Round($ibb.Top - $ibb.Bottom, 4))
    $doc.Unit = 3
    $mmb = $inc.BoundingBox
    Sonuc "Unit=3 (mm) iken ayni cizgi yuksekligi" ([math]::Round($mmb.Top - $mmb.Bottom, 4))

    # 5) Z-order semantigi (kucuk mu buyuk mu ustte?)
    Sonuc "ZOrder: l1 / l2" ("{0} / {1}" -f $l1.ZOrder, $l2.ZOrder)

    # 6) Katman: Printable varsayilani + create
    $testLyr = $page.CreateLayer("CDRCIFT_TEST")
    Sonuc "Yeni katman Printable varsayilani" $testLyr.Printable
    $testLyr.Printable = $false
    Sonuc "Printable=False atandi" $testLyr.Printable

    # 7) Isaret olusturma (daire + metin) + Fill.ApplyNoFill
    $mark = $testLyr.CreateEllipse(50 - 1.5, 50 + 1.5, 50 + 1.5, 50 - 1.5)
    $mark.Outline.Width = 0.5
    $mark.Outline.Color.RGBAssign(255, 0, 0)
    $mark.Fill.ApplyNoFill()
    Sonuc "CreateEllipse + RGBAssign + ApplyNoFill" ("genislik={0} renk={1}" -f $mark.Outline.Width, $mark.Outline.Color.HexValue)

    $txt = $testLyr.CreateArtisticText(52, 51, "7")
    try { $txt.Text.FontProperties.Size = 6 } catch {}
    Sonuc "CreateArtisticText + FontProperties.Size" $txt.Text.FontProperties.Size

    # 8) Combine -> 2 alt yol -> SubPath.Delete
    $a = $lyr.CreateLineSegment(200, 10, 200, 60)
    $b = $lyr.CreateLineSegment(220, 10, 220, 60)
    $a.AddToSelection()
    $b.AddToSelection()
    $sel = $doc.SelectionRange
    $comb = $sel.Combine()
    Sonuc "Combine sonrasi alt yol sayisi" $comb.Curve.SubPaths.Count
    try {
        $comb.Curve.SubPaths(1).Delete()
        Sonuc "SubPath.Delete sonrasi alt yol sayisi" $comb.Curve.SubPaths.Count
    } catch {
        Hata "SubPath.Delete" $_.Exception.Message
    }

    # 9) Grup + donusum: Curve mi DisplayCurve mi belge uzayinda?
    $doc.ClearSelection()
    $g1 = $lyr.CreateLineSegment(300, 10, 300, 60)
    $g2 = $lyr.CreateLineSegment(340, 10, 340, 60)
    $g1.AddToSelection()
    $g2.AddToSelection()
    $grp = $doc.SelectionRange.Group()
    $dondu = $false
    try { $grp.Rotate(30); $dondu = $true } catch { }
    Sonuc "Grup olustu / donduruldu" ("{0} / {1}" -f $grp.Shapes.Count, $dondu)

    $cocuk = $grp.Shapes(1)
    $cbb = $cocuk.BoundingBox
    $ccv = $cocuk.Curve.BoundingBox
    $cdc = $cocuk.DisplayCurve.BoundingBox
    Sonuc "Grup cocugu: nesne bbox" ("L={0:N3} R={1:N3}" -f $cbb.Left, $cbb.Right)
    Sonuc "Grup cocugu: Curve bbox" ("L={0:N3} R={1:N3}" -f $ccv.Left, $ccv.Right)
    Sonuc "Grup cocugu: DisplayCurve bbox" ("L={0:N3} R={1:N3}" -f $cdc.Left, $cdc.Right)
    $farkC = [math]::Abs($ccv.Left - $cbb.Left)
    $farkD = [math]::Abs($cdc.Left - $cbb.Left)
    Sonuc "Curve sapmasi / DisplayCurve sapmasi (mm)" ("{0:N3} / {1:N3}" -f $farkC, $farkD)

    # 10) Renk karsilastirma + HexValue
    Sonuc "Color.IsSame(l1,l2)" $l1.Outline.Color.IsSame($l2.Outline.Color)

    # 11) CommandGroup + Undo
    $doc.BeginCommandGroup("cdrcift test")
    $gecici = $lyr.CreateLineSegment(400, 10, 400, 60)
    $doc.EndCommandGroup()
    $onceki = $page.Shapes.Count
    $doc.Undo(1)
    Sonuc "CommandGroup + Undo(1) nesne sayisi" ("{0} -> {1}" -f $onceki, $page.Shapes.Count)

    # 12) Parametrik sekiller: Curve bos, DisplayCurve dolu mu?
    $rect = $lyr.CreateRectangle(500, 60, 540, 10)
    $ell = $lyr.CreateEllipse(560, 60, 600, 10)
    Sonuc "Dikdortgen: Curve / DisplayCurve alt yol" ("{0} / {1}" -f $rect.Curve.SubPaths.Count, $rect.DisplayCurve.SubPaths.Count)
    Sonuc "Elips: Curve / DisplayCurve alt yol" ("{0} / {1}" -f $ell.Curve.SubPaths.Count, $ell.DisplayCurve.SubPaths.Count)
    $segK = $rect.DisplayCurve.Segments
    Sonuc "Dikdortgen DisplayCurve segment sayisi" $segK.Count
    if ($segK.Count -gt 0) {
        Sonuc "Dikdortgen ilk segment Type (0=cizgi)" $segK.Item(1).Type
    }
    $segE = $ell.DisplayCurve.Segments
    if ($segE.Count -gt 0) {
        Sonuc "Elips ilk segment Type (1=egri)" $segE.Item(1).Type
    }

    # 13) Duz sekilde Curve/Shape bbox farki
    $fark = [math]::Abs($cv.BoundingBox.Left - $l1.BoundingBox.Left)
    Sonuc "Duz sekilde Curve/Shape bbox farki (mm)" ("{0:N4}" -f $fark)

    # 14) SubPath.Delete son alt yolda nesneyi de siler mi? (kodda koruma var)
    $doc.ClearSelection()
    $q1 = $lyr.CreateLineSegment(700, 10, 700, 60)
    $q2 = $lyr.CreateLineSegment(720, 10, 720, 60)
    $q1.AddToSelection()
    $q2.AddToSelection()
    $qc = $doc.SelectionRange.Combine()
    $qc.Curve.SubPaths(1).Delete()
    $qc.Curve.SubPaths(1).Delete()
    $hayatta = $true
    try { $x = $qc.StaticID } catch { $hayatta = $false }
    Sonuc "Son alt yol silinince nesne hayatta mi (beklenen: hayir)" $hayatta
}
catch {
    Hata "Genel" $_.Exception.Message
}
finally {
    # Temizlik: yalnizca kendi belgemizi kapat, kullanicinin uygulamasini KAPATMA
    if ($null -ne $doc) {
        try {
            $doc.Dirty = $false
            $doc.Close()
            Write-Host "[bilgi] test belgesi kaydedilmeden kapatildi"
        } catch {
            Write-Host ("[uyari] belge kapatilamadi: " + $_.Exception.Message)
        }
    }
    if ($yeniOrnek) {
        try {
            $app.Quit()
            Write-Host "[bilgi] test icin acilan CorelDRAW ornegi kapatildi"
        } catch {
            Write-Host ("[uyari] ornek kapatilamadi: " + $_.Exception.Message)
        }
    }
}

Write-Host ""
Write-Host ("Gecen: {0}   Kalan: {1}" -f $script:gecen, $script:kalan)
if ($script:kalan -gt 0) { exit 1 }
exit 0
