<#
.SYNOPSIS
    cdrcift modullerini CorelDRAW'a aktarip CiftCizgiDenetcisi.gms olarak kaydeder.

.DESCRIPTION
    VBA IDE'sinde elle "Dosyayı İçe Aktar" adimlarini otomatiklestirir.
    Gecici ve gorunmez bir CorelDRAW ornegi kullanir; acik belgelere ve
    calisan CorelDRAW oturumunuza dokunmaz.

    Islem:
      1. GMSManager.Projects.Create  -> <UserGMSPath>\CiftCizgiDenetcisi.gms
      2. VBE.VBComponents.Import     -> 5 modul
      3. Unload                      -> diske yazim
      4. Load + dogrulama            -> modul ve makro sayisi
      5. (opsiyonel) -MacroTest      -> derleme/kalistirma kontrolu

.PARAMETER Surum
    Kurulacak CorelDRAW surumu: 2024 | 2025 | 2026 (varsayilan 2026).

.PARAMETER MacroTest
    Kurulumdan sonra UI'suz bir makro calistirir (derleme kontrolu).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\gms_kur.ps1
    powershell -ExecutionPolicy Bypass -File tools\gms_kur.ps1 -Surum 2025
    powershell -ExecutionPolicy Bypass -File tools\gms_kur.ps1 -MacroTest
#>

[CmdletBinding()]
param(
    [ValidateSet("2024", "2025", "2026")]
    [string]$Surum = "2026",
    [switch]$MacroTest
)

$ErrorActionPreference = "Stop"

$projeAd = "CiftCizgiDenetcisi"
$moduller = @(
    "CiftCizgiGeom.bas",
    "CiftCizgiTarama.bas",
    "CiftCizgiRapor.bas",
    "CiftCizgi.bas",
    "CiftCizgiTest.bas"
)

$src = Join-Path $PSScriptRoot "..\src"
$src = (Resolve-Path $src).Path
Write-Host ("Kaynak klasor : {0}" -f $src)
foreach ($m in $moduller) {
    if (-not (Test-Path (Join-Path $src $m))) { throw "Modul bulunamadi: $m" }
}

$app = $null
$gms = $null
$basarili = $false

try {
    Write-Host ("CorelDRAW {0} icin gorunmez ornek aciliyor..." -f $Surum)
    $app = New-Object -ComObject CorelDRAW.Application
    $app.Visible = $false
    $gms = $app.GMSManager

    $surumYolu = $gms.UserGMSPath
    if ($surumYolu -notmatch $Surum) {
        Write-Host ("[uyari] Calisan surum yolu '{0}' icinde {1} gecmiyor." -f $surumYolu, $Surum)
        Write-Host "        Kurulum yine de bu surumun klasorune yapilacak."
    }

    $hedef = Join-Path $surumYolu ($projeAd + ".gms")

    # Varsa once kaldir
    foreach ($p in @($gms.Projects)) {
        if ($p.Name -eq $projeAd) {
            Write-Host ("Mevcut proje kaldiriliyor: {0}" -f $p.Name)
            try { $p.Unload() } catch { }
        }
    }
    if (Test-Path $hedef) {
        Remove-Item $hedef -Force
        Write-Host ("Eski dosya silindi: {0}" -f $hedef)
    }

    $proj = $gms.Projects.Create($projeAd)
    Write-Host ("1/4 GMS olusturuldu : {0}" -f $proj.FullFileName)

    $vp = $null
    foreach ($p in $app.VBE.VBProjects) { if ($p.Name -eq $projeAd) { $vp = $p } }
    if ($null -eq $vp) { throw "VBE'de proje bulunamadi (VBA erisimi kapali olabilir)." }

    Write-Host "2/4 Moduller ice aktariliyor"
    foreach ($m in $moduller) {
        $c = $vp.VBComponents.Import((Join-Path $src $m))
        Write-Host ("     + {0,-24} -> {1}" -f $m, $c.Name)
    }

    Write-Host "3/4 Diske yaziliyor (Unload)"
    $gms.Projects.Item($projeAd).Unload()
    Start-Sleep -Milliseconds 400
    if (-not (Test-Path $hedef)) { throw "GMS dosyasi olusmadi: $hedef" }
    $boyut = (Get-Item $hedef).Length
    Write-Host ("     {0}  ({1} bayt)" -f $hedef, $boyut)

    Write-Host "4/4 Dogrulama (yeniden yukleme)"
    $p2 = $gms.Projects.Load($hedef)
    $makro = @($p2.Macros.Count)
    Write-Host ("     proje: {0}   makro: {1}   dirty: {2}" -f $p2.Name, $makro, $p2.Dirty)
    if ($p2.Macros.Count -lt 5) { throw "Beklenenden az makro: $($p2.Macros.Count)" }

    if ($MacroTest) {
        Write-Host "Derleme/kalistirma testi: CiftCizgiGeom.VarsayilanAyarlar"
        try {
            $gms.RunMacro("CiftCizgiGeom.VarsayilanAyarlar")
            Write-Host "     [OK] makro calisti (proje derlendi)"
        } catch {
            Write-Host ("     [HATA] " + $_.Exception.Message)
            throw
        }
    }

    $basarili = $true
}
finally {
    if ($null -ne $app) {
        try { $app.Quit() } catch { }
    }
}

if ($basarili) {
    Write-Host ""
    Write-Host "KURULUM TAMAM."
    Write-Host "CorelDRAW calisiyorsa: Araclar > Makrolar > Makro Yoneticisi > Yukle"
    Write-Host ("  veya uygulamayi bir kez kapatip acin; makrolar otomatik yuklenir.")
    Write-Host "Makrolar: CiftCizgi.Denetle / .Temizle / .IsaretleriKaldir / .Ayarlar"
    Write-Host "Test    : CiftCizgiTest.TestleriCalistir"
    exit 0
} else {
    exit 1
}
