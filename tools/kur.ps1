<#
.SYNOPSIS
    cdrcift - CorelDRAW Cift Cizgi Denetleyicisi: otomatik kurulum (Windows)

.DESCRIPTION
    CiftCizgi.js dosyasini, kurulu her CorelDRAW surumunun
    "Draw\Scripts" klasorune kopyalar ve Turkce karakterler icin
    ANSI (cp1254) kodlamasina cevirir.

    * Yonetici hakki GEREKMEZ
    * Kayit defterine dokunmaz
    * Var olan dosyanin uzerine yazar (once .bak yedegi alir)

.PARAMETER Surum
    Yalniz belirli bir surum: 2024 | 2025 | 2026
    Bos birakilirsa bulunan tum surumlere kurar.

.PARAMETER Kaldir
    Kurulumu kaldirir (CiftCizgi.js dosyalarini siler).

.PARAMETER Kaynak
    CiftCizgi.js dosyasinin tam yolu. Verilmezse script'in yaninda
    veya ..\js\ altinda aranir; bulunamazsa GitHub'dan indirilir.

.PARAMETER Sessiz
    Ekrana ozet yazmaz (otomasyon icin).

.PARAMETER KlasorAc
    Kurulumdan sonra Scripts klasorunu Explorer'da acar.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File kur.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File kur.ps1 -Surum 2026

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File kur.ps1 -Kaldir
#>

[CmdletBinding()]
param(
    [string]$Surum = "",
    [switch]$Kaldir,
    [switch]$Sessiz,
    [string]$Kaynak = "",
    [switch]$KlasorAc
)

$ErrorActionPreference = "Stop"
$UzakAdres = "https://raw.githubusercontent.com/ewolution/cdrcift/main/js/CiftCizgi.js"
$DosyaAdi = "CiftCizgi.js"
$Kodlama = [System.Text.Encoding]::GetEncoding(1254)

function Yaz([string]$metin) {
    if (-not $Sessiz) { Write-Host $metin }
}

function Hata([string]$metin) {
    Write-Host ("[HATA] " + $metin) -ForegroundColor Red
}

# ---------------------------------------------------------------------
# 1) Kurulu CorelDRAW surumlerini bul
# ---------------------------------------------------------------------

$corelKok = Join-Path $env:APPDATA "Corel"
if (-not (Test-Path $corelKok)) {
    Hata "CorelDRAW ayar klasoru bulunamadi: $corelKok"
    exit 1
}

$surumler = @()
foreach ($d in Get-ChildItem $corelKok -Directory -ErrorAction SilentlyContinue) {
    if ($d.Name -notmatch "^CorelDRAW Graphics Suite") { continue }
    $draw = Join-Path $d.FullName "Draw"
    if (-not (Test-Path $draw)) { continue }
    if ($Surum -and ($d.Name -notlike "*$Surum*")) { continue }
    $surumler += [pscustomobject]@{
        Ad      = $d.Name
        Scripts = Join-Path $draw "Scripts"
    }
}

if ($surumler.Count -eq 0) {
    Hata "Kurulu CorelDRAW surumu bulunamadi" + $(if ($Surum) { " (filtre: $Surum)" } else { "" })
    Yaz  "  Beklenen konum: $corelKok\CorelDRAW Graphics Suite <yil>\Draw"
    exit 1
}

Yaz ("cdrcift kurulum - " + $surumler.Count + " CorelDRAW surumu bulundu")

# ---------------------------------------------------------------------
# 2) Kaldirma
# ---------------------------------------------------------------------

if ($Kaldir) {
    $silinen = 0
    foreach ($s in $surumler) {
        $hedef = Join-Path $s.Scripts $DosyaAdi
        if (Test-Path $hedef) {
            Remove-Item $hedef -Force
            Yaz ("  [kaldirildi] " + $s.Ad)
            $silinen++
        }
    }
    if ($silinen -eq 0) { Yaz "  Silinecek dosya bulunamadi." }
    else { Yaz ("TAMAM: " + $silinen + " dosya silindi.") }
    exit 0
}

# ---------------------------------------------------------------------
# 3) Kaynak dosyayi bul (yerel -> GitHub)
# ---------------------------------------------------------------------

function KaynakBul {
    param([string]$Verilen)

    $adaylar = @()
    if ($Verilen) { $adaylar += $Verilen }
    if ($PSScriptRoot) {
        $adaylar += (Join-Path $PSScriptRoot $DosyaAdi)
        $adaylar += (Join-Path (Split-Path $PSScriptRoot -Parent) ("js\" + $DosyaAdi))
        $adaylar += (Join-Path (Split-Path $PSScriptRoot -Parent) ("js/" + $DosyaAdi))
    }
    $adaylar += (Join-Path (Get-Location) $DosyaAdi)

    foreach ($a in $adaylar) {
        if ($a -and (Test-Path $a)) { return (Resolve-Path $a).Path }
    }

    # Yerelde yok: GitHub'dan indir
    $gecici = Join-Path $env:TEMP ("cdrcift-" + $DosyaAdi)
    Yaz "  Yerel kopya yok; GitHub'dan indiriliyor..."
    try {
        Invoke-WebRequest -Uri $UzakAdres -OutFile $gecici -UseBasicParsing
    } catch {
        Hata "Indirme basarisiz: $($_.Exception.Message)"
        exit 2
    }
    return $gecici
}

$kaynakYol = KaynakBul -Verilen $Kaynak
Yaz ("  Kaynak: " + $kaynakYol)

$utf8 = New-Object System.Text.UTF8Encoding($false)
$metin = [System.IO.File]::ReadAllText($kaynakYol, $utf8)

try {
    $veri = $Kodlama.GetBytes($metin)
} catch {
    Hata "Dosya cp1254 kodlamasina cevrilemedi: $($_.Exception.Message)"
    exit 2
}

# ---------------------------------------------------------------------
# 4) Kopyala + dogrula
# ---------------------------------------------------------------------

$kurulan = 0
$hatali = 0

foreach ($s in $surumler) {
    try {
        if (-not (Test-Path $s.Scripts)) {
            New-Item -ItemType Directory -Force -Path $s.Scripts | Out-Null
            Yaz ("  [olusturuldu] " + $s.Scripts)
        }

        $hedef = Join-Path $s.Scripts $DosyaAdi

        if (Test-Path $hedef) {
            Copy-Item $hedef ($hedef + ".bak") -Force
        }

        [System.IO.File]::WriteAllBytes($hedef, $veri)

        # Dogrulama: geri oku, metin ayni mi?
        $geri = [System.IO.File]::ReadAllText($hedef, $Kodlama)
        if ($geri -ne $metin) { throw "Geri okuma farki var" }

        $boyut = (Get-Item $hedef).Length
        Yaz ("  [OK] " + $s.Ad + "  ->  " + $hedef + "  (" + $boyut + " bayt)")
        $kurulan++
    } catch {
        Hata ($s.Ad + ": " + $_.Exception.Message)
        $hatali++
    }
}

# ---------------------------------------------------------------------
# 5) Ozet
# ---------------------------------------------------------------------

Yaz ""
if ($hatali -eq 0) {
    Yaz ("KURULUM TAMAM: " + $kurulan + " surume kuruldu.")
} else {
    Yaz ("KURULUM: " + $kurulan + " basarili, " + $hatali + " hatali.")
}

if ($KlasorAc -and $surumler.Count -gt 0) {
    try { Start-Process explorer.exe $surumler[0].Scripts } catch { }
}

Yaz ""
Yaz "Kullanim:"
Yaz "  1) CorelDRAW'da bir dosya acin"
Yaz "  2) Araclar > Makrolar > JavaScript  (veya Scripts paneli)"
Yaz "  3) CiftCizgi"
Yaz ""
Yaz "Kaldirmak icin:  kur.ps1 -Kaldir"

if ($hatali -gt 0) { exit 2 }
exit 0
