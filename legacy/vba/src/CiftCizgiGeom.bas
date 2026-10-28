Attribute VB_Name = "CiftCizgiGeom"
Option Explicit

' =====================================================================
'  cdrcift - Cift cizgi denetleyicisi
'  Modul: Geometri cekirdegi (tip tanimlari, ayarlar, matematik, imza)
'  Surum: 1.0.0        Hedef: CorelDRAW 2024 / 2025 / 2026 (VBA)
'
'  NOT: Bu dosya bilincli olarak ASCII tutulmustur. VBE (VBA IDE)
'  .bas dosyalarini yerel kod sayfasi ile okudugu icin Turkce
'  karakterler bozulabilir; bu yuzden metinler sade Turkce yazildi.
' =====================================================================

'-[ Sabitler ]--------------------------------------------------------

Public Const SURUM As String = "1.0.0"
Public Const ISARET_KATMAN As String = "DUP_CHECK"
Public Const UYGULAMA_ADI As String = "Cift Cizgi Denetleyicisi"

'-[ Tipler ]----------------------------------------------------------

Public Type TNesne
    Ad As String            ' Gorunen ad / otomatik ad
    Katman As String
    Renk As String          ' HexValue (rapor icin)
    Kalinlik As Double      ' mm
    StatikID As Long
    Z As Long               ' Z-order: 1 = en ustte, buyuk deger = altta
    AltYolSayisi As Long
    Uzunluk As Double       ' mm, tum alt yollar
End Type

Public Type TBulgu
    Tur As String           ' TAM | KISMI | YAKIN | NOKTAS
    Nesne1 As Long          ' gNesne indeksi
    AltYol1 As Long
    Nesne2 As Long          ' eslesen (korunan) nesne; -1 = yok
    AltYol2 As Long
    Uzunluk As Double       ' mm
    Oran As Double          ' 0..1
    X As Double
    Y As Double
    Eylem As String         ' sil | elle
    Aciklama As String
End Type

'-[ Ayarlar ]---------------------------------------------------------

Public gTolerans As Double          ' mm, birebir sayilma esigi
Public gYakinTol As Double          ' mm, kaymis kopya esigi
Public gMinOrtusme As Double        ' mm, en kisa raporlanacak ortusme
Public gKapsam As String            ' "sayfa" | "secim"
Public gKuruCalisma As Boolean
Public gOtomatikSil As Boolean
Public gYoksayKatmanlar As String   ' "Katman1,Katman2"
Public gCsvAyirici As String        ' ";" | ","
Public gOndalik As String           ' "," | "."
Public gHucreMm As Double           ' uzamsal izgara hucre boyu
Public gMaxBulgu As Long
Public gBoyama As Boolean           ' bulunan nesneleri kirmiziya boya

'-[ Birim ]-----------------------------------------------------------

Public gBirimKatsayi As Double      ' belge birimi -> mm
Public gBirimAdi As String

'-[ Tarama durumu ]---------------------------------------------------

Public gNesne() As TNesne
Public gNesneRef() As Object
Public gNesneSayisi As Long

Public gAltYolImza() As String      ' her alt yol icin yapisal imza
Public gAltYolNesne() As Long
Public gAltYolNo() As Long
Public gAltYolUzunluk() As Double
Public gAltYolKapali() As Boolean
Public gAltYolX() As Double
Public gAltYolY() As Double
Public gAltYolSayisi As Long

Public gSegX1() As Double
Public gSegY1() As Double
Public gSegX2() As Double
Public gSegY2() As Double
Public gSegNesne() As Long
Public gSegAltYol() As Long
Public gSegSayisi As Long

'-[ Sonuclar ]--------------------------------------------------------

Public gBulgu() As TBulgu
Public gBulguSayisi As Long
Public gBulguKesildi As Boolean

Public gAtlanan As Long             ' curve alinamayan nesne sayisi
Public gEfektAtlanan As Long
Public gPowerClipAtlanan As Long
Public gUyari As String

'-[ Karsilastirma ciktisi ]-------------------------------------------

Public gKarsOrtusme As Double
Public gKarsOran As Double
Public gKarsX As Double
Public gKarsY As Double

Private Const PI2 As Double = 6.28318530717959
Public gParalelSin As Double        ' 1 derecenin sinusu

'-[ Kurulum ]---------------------------------------------------------

Public Sub VarsayilanAyarlar()
    gTolerans = 0.05
    gYakinTol = 0.5
    gMinOrtusme = 0.5
    gKapsam = "sayfa"
    gKuruCalisma = False
    gOtomatikSil = True
    gYoksayKatmanlar = ""
    gCsvAyirici = ";"
    gOndalik = ","
    gHucreMm = 2
    gMaxBulgu = 5000
    gBoyama = False
    gParalelSin = Sin(1 * PI2 / 360)    ' 1 derece
    gBirimKatsayi = 1
    gBirimAdi = "mm"
End Sub

Public Sub DurumSifirla()
    gNesneSayisi = 0
    gAltYolSayisi = 0
    gSegSayisi = 0
    gBulguSayisi = 0
    gBulguKesildi = False
    gAtlanan = 0
    gEfektAtlanan = 0
    gPowerClipAtlanan = 0
    gUyari = ""
    Erase gNesne
    Erase gNesneRef
    Erase gAltYolImza
    Erase gAltYolNesne
    Erase gAltYolNo
    Erase gAltYolUzunluk
    Erase gAltYolKapali
    Erase gAltYolX
    Erase gAltYolY
    Erase gSegX1
    Erase gSegY1
    Erase gSegX2
    Erase gSegY2
    Erase gSegNesne
    Erase gSegAltYol
    Erase gBulgu
End Sub

'-[ Dizi buyutme ]----------------------------------------------------

Public Sub NesneYerAc()
    If gNesneSayisi = 0 Then
        ReDim gNesne(0 To 255)
        ReDim gNesneRef(0 To 255)
    ElseIf gNesneSayisi > UBound(gNesne) Then
        ReDim Preserve gNesne(0 To UBound(gNesne) * 2 + 1)
        ReDim Preserve gNesneRef(0 To UBound(gNesneRef) * 2 + 1)
    End If
End Sub

Public Sub AltYolYerAc()
    If gAltYolSayisi = 0 Then
        ReDim gAltYolImza(0 To 511)
        ReDim gAltYolNesne(0 To 511)
        ReDim gAltYolNo(0 To 511)
        ReDim gAltYolUzunluk(0 To 511)
        ReDim gAltYolKapali(0 To 511)
        ReDim gAltYolX(0 To 511)
        ReDim gAltYolY(0 To 511)
    ElseIf gAltYolSayisi > UBound(gAltYolImza) Then
        ReDim Preserve gAltYolImza(0 To UBound(gAltYolImza) * 2 + 1)
        ReDim Preserve gAltYolNesne(0 To UBound(gAltYolNesne) * 2 + 1)
        ReDim Preserve gAltYolNo(0 To UBound(gAltYolNo) * 2 + 1)
        ReDim Preserve gAltYolUzunluk(0 To UBound(gAltYolUzunluk) * 2 + 1)
        ReDim Preserve gAltYolKapali(0 To UBound(gAltYolKapali) * 2 + 1)
        ReDim Preserve gAltYolX(0 To UBound(gAltYolX) * 2 + 1)
        ReDim Preserve gAltYolY(0 To UBound(gAltYolY) * 2 + 1)
    End If
End Sub

Public Sub SegmentYerAc()
    If gSegSayisi = 0 Then
        ReDim gSegX1(0 To 2047)
        ReDim gSegY1(0 To 2047)
        ReDim gSegX2(0 To 2047)
        ReDim gSegY2(0 To 2047)
        ReDim gSegNesne(0 To 2047)
        ReDim gSegAltYol(0 To 2047)
    ElseIf gSegSayisi > UBound(gSegX1) Then
        ReDim Preserve gSegX1(0 To UBound(gSegX1) * 2 + 1)
        ReDim Preserve gSegY1(0 To UBound(gSegY1) * 2 + 1)
        ReDim Preserve gSegX2(0 To UBound(gSegX2) * 2 + 1)
        ReDim Preserve gSegY2(0 To UBound(gSegY2) * 2 + 1)
        ReDim Preserve gSegNesne(0 To UBound(gSegNesne) * 2 + 1)
        ReDim Preserve gSegAltYol(0 To UBound(gSegAltYol) * 2 + 1)
    End If
End Sub

Public Sub BulguYerAc()
    If gBulguSayisi = 0 Then
        ReDim gBulgu(0 To 127)
    ElseIf gBulguSayisi > UBound(gBulgu) Then
        ReDim Preserve gBulgu(0 To UBound(gBulgu) * 2 + 1)
    End If
End Sub

'-[ Matematik ]-------------------------------------------------------

Public Function Qz(ByVal v As Double) As Long
    ' 0,005 mm izgarasina yuvarlar (negatiflerde de dogru)
    Qz = CLng(Int(v / 0.005 + 0.5))
End Function

Public Function Kisalt(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then
        Kisalt = s
    Else
        Kisalt = Left$(s, n - 3) & "..."
    End If
End Function

Public Function Kucuk(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Kucuk = a Else Kucuk = b
End Function

Public Function Buyuk(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Buyuk = a Else Buyuk = b
End Function

'-[ Birim donusumu ]--------------------------------------------------

Public Function BirimKatsayi(ByVal belgeBirimi As Long, ByRef ad As String) As Double
    ' Application.ConvertUnits icin cdrUnit degerleri (CorelDRAW tip kutuphanesi):
    ' 0 TenthMicron, 1 Inch, 2 Foot, 3 Millimeter, 4 Centimeter, 5 Pixel,
    ' 6 Mile, 7 Meter, 8 Kilometer, 9 Didots, 10 Agate, 11 Yard, 12 Pica,
    ' 13 Cicero, 14 Point, 15 Q, 16 H
    Select Case belgeBirimi
        Case 0:  ad = "0.1 um":   BirimKatsayi = 0.0001
        Case 1:  ad = "inc":      BirimKatsayi = 25.4
        Case 2:  ad = "foot":     BirimKatsayi = 304.8
        Case 3:  ad = "mm":       BirimKatsayi = 1
        Case 4:  ad = "cm":       BirimKatsayi = 10
        Case 5:  ad = "piksel":   BirimKatsayi = 1      ' cozunurluge bagli; uyari uretilir
        Case 6:  ad = "mil":      BirimKatsayi = 1609344
        Case 7:  ad = "m":        BirimKatsayi = 1000
        Case 8:  ad = "km":       BirimKatsayi = 1000000
        Case 9:  ad = "didot":    BirimKatsayi = 0.375972
        Case 10: ad = "agate":    BirimKatsayi = 4.233333
        Case 11: ad = "yarda":    BirimKatsayi = 914.4
        Case 12: ad = "pika":     BirimKatsayi = 4.233333
        Case 13: ad = "cicero":   BirimKatsayi = 4.511278
        Case 14: ad = "pt":       BirimKatsayi = 0.352778
        Case 15: ad = "Q":        BirimKatsayi = 0.25
        Case 16: ad = "H":        BirimKatsayi = 0.1
        Case Else: ad = "bilinmiyor": BirimKatsayi = 1
    End Select
End Function

'-[ Yapisal imza ]----------------------------------------------------
'  Amac: birebir kopya (TAM) tespitini poligon yaklastirmasindan
'  bagimsiz, dugum duzeyinde ve O(1) karsilastirilabilir yapmak.
'  - Kapali yollarda baslangic dugumu dondurulur (kanonik)
'  - Ters yonde cizilmis kopyalar ayni imzayi verir
'  - Koordinatlar 0,005 mm izgarasina yuvarlanir

Public Function AltYolImzasi(ByVal sp As Object) As String
    Dim n As Long, i As Long
    Dim segs As Object, seg As Object
    Dim tokF() As String, tokR() As String
    Dim tf As String, tr As String
    Dim typ As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double
    Dim kapali As Boolean

    On Error GoTo Hata
    Set segs = sp.Segments
    n = segs.Count
    If n < 1 Then Exit Function

    ReDim tokF(0 To n - 1)
    ReDim tokR(0 To n - 1)

    For i = 1 To n
        Set seg = segs(i)
        typ = seg.Type
        x1 = seg.StartNode.PositionX * gBirimKatsayi
        y1 = seg.StartNode.PositionY * gBirimKatsayi
        x2 = seg.EndNode.PositionX * gBirimKatsayi
        y2 = seg.EndNode.PositionY * gBirimKatsayi

        tf = CStr(typ) & "|" & Qz(x1) & "|" & Qz(y1) & "|" & Qz(x2) & "|" & Qz(y2)
        tr = CStr(typ) & "|" & Qz(x2) & "|" & Qz(y2) & "|" & Qz(x1) & "|" & Qz(y1)

        If typ = 1 Then
            c1x = seg.StartingControlPointX * gBirimKatsayi
            c1y = seg.StartingControlPointY * gBirimKatsayi
            c2x = seg.EndingControlPointX * gBirimKatsayi
            c2y = seg.EndingControlPointY * gBirimKatsayi
            tf = tf & "|" & Qz(c1x) & "|" & Qz(c1y) & "|" & Qz(c2x) & "|" & Qz(c2y)
            tr = tr & "|" & Qz(c2x) & "|" & Qz(c2y) & "|" & Qz(c1x) & "|" & Qz(c1y)
        End If

        tokF(i - 1) = tf
        tokR(i - 1) = tr
    Next i

    kapali = sp.Closed

    If kapali Then
        tf = KapaliKanonik(tokF)
        tr = KapaliKanonik(tokR)
    Else
        tf = Join(tokF, ";")
        tr = Join(tokR, ";")
    End If

    If StrComp(tr, tf, vbBinaryCompare) < 0 Then tf = tr
    AltYolImzasi = IIf(kapali, "C", "O") & CStr(n) & "|" & tf
    Exit Function
Hata:
    AltYolImzasi = ""
End Function

Private Function KapaliKanonik(ByRef tok() As String) As String
    ' Dairesel diziyi sozluk sirasina gore en kucuk token'a dondurur
    Dim i As Long, n As Long, en As Long
    Dim s As String
    n = UBound(tok) + 1
    en = 0
    For i = 1 To n - 1
        If StrComp(tok(i), tok(en), vbBinaryCompare) < 0 Then en = i
    Next i
    For i = 0 To n - 1
        If i > 0 Then s = s & ";"
        s = s & tok((en + i) Mod n)
    Next i
    KapaliKanonik = s
End Function

'-[ Siniflandirma ]---------------------------------------------------
'  Iki segmenti (mm cinsinden) karsilastirir.
'  Donus: "" | "TAM" | "KISMI" | "YAKIN"
'  Sonuc ayrintilari gKarsOrtusme / gKarsOran / gKarsX / gKarsY'de.

Public Function Karsilastir(ByVal i As Long, ByVal j As Long) As String
    Dim ax1 As Double, ay1 As Double, ax2 As Double, ay2 As Double
    Dim bx1 As Double, by1 As Double, bx2 As Double, by2 As Double
    Dim l1 As Double, l2 As Double, mn As Double
    Dim ux As Double, uy As Double, vx As Double, vy As Double
    Dim cr As Double
    Dim d1 As Double, d2 As Double, off As Double
    Dim t1 As Double, t2 As Double, lo As Double, hi As Double, ov As Double

    Karsilastir = ""
    gKarsOrtusme = 0
    gKarsOran = 0

    ax1 = gSegX1(i): ay1 = gSegY1(i): ax2 = gSegX2(i): ay2 = gSegY2(i)
    bx1 = gSegX1(j): by1 = gSegY1(j): bx2 = gSegX2(j): by2 = gSegY2(j)

    l1 = Sqr((ax2 - ax1) * (ax2 - ax1) + (ay2 - ay1) * (ay2 - ay1))
    l2 = Sqr((bx2 - bx1) * (bx2 - bx1) + (by2 - by1) * (by2 - by1))
    If l1 < 0.001 Or l2 < 0.001 Then Exit Function

    ux = (ax2 - ax1) / l1: uy = (ay2 - ay1) / l1
    vx = (bx2 - bx1) / l2: vy = (by2 - by1) / l2

    cr = Abs(ux * vy - uy * vx)
    If cr > gParalelSin Then Exit Function          ' paralel degil

    ' B uc noktalarinin A dogrusuna dik uzakligi
    d1 = Abs((bx1 - ax1) * (-uy) + (by1 - ay1) * ux)
    d2 = Abs((bx2 - ax1) * (-uy) + (by2 - ay1) * ux)
    off = Buyuk(d1, d2)

    ' A uzerindeki 1B ortusme
    t1 = (bx1 - ax1) * ux + (by1 - ay1) * uy
    t2 = (bx2 - ax1) * ux + (by2 - ay1) * uy
    lo = Buyuk(0, Kucuk(t1, t2))
    hi = Kucuk(l1, Buyuk(t1, t2))
    ov = hi - lo
    If ov < 0.001 Then Exit Function

    mn = Kucuk(l1, l2)

    If off <= gTolerans Then
        If ov >= mn - 0.001 Then
            Karsilastir = "TAM"
        ElseIf ov >= gMinOrtusme Then
            Karsilastir = "KISMI"
        Else
            Exit Function
        End If
    ElseIf off <= gYakinTol And ov >= 0.8 * mn Then
        Karsilastir = "YAKIN"
    Else
        Exit Function
    End If

    gKarsOrtusme = ov
    gKarsOran = ov / mn
    gKarsX = ax1 + ux * (lo + ov / 2)
    gKarsY = ay1 + uy * (lo + ov / 2)
End Function

'-[ Bulgu kaydi ]-----------------------------------------------------

Public Function BulguEkle(ByVal tur As String, ByVal n1 As Long, ByVal a1 As Long, _
                          ByVal n2 As Long, ByVal a2 As Long, ByVal uzunluk As Double, _
                          ByVal oran As Double, ByVal x As Double, ByVal y As Double, _
                          ByVal eylem As String, ByVal aciklama As String) As Long

    If gBulguSayisi >= gMaxBulgu Then
        gBulguKesildi = True
        BulguEkle = -1
        Exit Function
    End If

    BulguYerAc()
    With gBulgu(gBulguSayisi)
        .Tur = tur
        .Nesne1 = n1
        .AltYol1 = a1
        .Nesne2 = n2
        .AltYol2 = a2
        .Uzunluk = uzunluk
        .Oran = oran
        .X = x
        .Y = y
        .Eylem = eylem
        .Aciklama = aciklama
    End With
    BulguEkle = gBulguSayisi
    gBulguSayisi = gBulguSayisi + 1
End Function

Public Function BulguNoBul(ByVal tur As String, ByVal n1 As Long, ByVal a1 As Long, _
                           ByVal n2 As Long, ByVal a2 As Long) As Long
    Dim i As Long
    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            If .Tur = tur Then
                If (.Nesne1 = n1 And .AltYol1 = a1 And .Nesne2 = n2 And .AltYol2 = a2) Or _
                   (.Nesne1 = n2 And .AltYol1 = a2 And .Nesne2 = n1 And .AltYol2 = a1) Then
                    BulguNoBul = i
                    Exit Function
                End If
            End If
        End With
    Next i
    BulguNoBul = -1
End Function

'-[ Nesne ozellik karsilastirmasi (silme guvenligi) ]----------------

Public Function AyniOzellik(ByVal n1 As Long, ByVal n2 As Long, ByRef neden As String) As Boolean
    ' Otomatik silme yalnizca katman + cizgi rengi + kalinlik ayniysa yapilir.
    Dim s1 As Object, s2 As Object
    On Error Resume Next
    Set s1 = gNesneRef(n1)
    Set s2 = gNesneRef(n2)
    If s1 Is Nothing Or s2 Is Nothing Then
        neden = "nesne referansi yok"
        AyniOzellik = False
        Exit Function
    End If

    If gNesne(n1).Katman <> gNesne(n2).Katman Then
        neden = "farkli katman"
        AyniOzellik = False
        Exit Function
    End If

    If Abs(gNesne(n1).Kalinlik - gNesne(n2).Kalinlik) > 0.001 Then
        neden = "farkli cizgi kalinligi"
        AyniOzellik = False
        Exit Function
    End If

    Dim ayni As Boolean
    ayni = False
    ayni = s1.Outline.Color.IsSame(s2.Outline.Color)
    If Err.Number <> 0 Then
        Err.Clear
        neden = "renk karsilastirilamadi"
        AyniOzellik = False
        Exit Function
    End If
    If Not ayni Then
        neden = "farkli cizgi rengi"
        AyniOzellik = False
        Exit Function
    End If

    neden = ""
    AyniOzellik = True
End Function

Public Function KatmanYoksayilir(ByVal katman As String) As Boolean
    If Len(Trim$(gYoksayKatmanlar)) = 0 Then Exit Function
    Dim p() As String, i As Long
    p = Split(gYoksayKatmanlar, ",")
    For i = 0 To UBound(p)
        If StrComp(Trim$(p(i)), katman, vbTextCompare) = 0 Then
            KatmanYoksayilir = True
            Exit Function
        End If
    Next i
End Function

'-[ Rapor bicimlendirme ]---------------------------------------------

Public Function Sayi(ByVal v As Double, ByVal basamak As Long) As String
    ' Format$ yerel ondalik ayiricisini kullanir; istenen ayiriciya cevrilir.
    Dim s As String, yerel As String
    If basamak <= 0 Then
        s = Format$(v, "0")
    Else
        s = Format$(v, "0." & String$(basamak, "0"))
    End If
    yerel = Mid$(Format$(1.5, "0.0"), 2, 1)
    If yerel <> gOndalik Then s = Replace(s, yerel, gOndalik)
    Sayi = s
End Function
