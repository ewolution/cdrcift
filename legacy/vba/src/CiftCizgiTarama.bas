Attribute VB_Name = "CiftCizgiTarama"
Option Explicit

' =====================================================================
'  cdrcift - Tarama modulu
'  Belge dolasimi, alt yol duzlestirme, uzamsal izgara, cift tarama
'  Surum: 1.0.0
' =====================================================================

Private mIzgara As Object             ' hucre -> "seg1;seg2;..."
Private mTamCift As Object            ' yapisal TAM ile cozulen alt yol ciftleri
Private mBulguDedup As Object         ' birlesik bulgu anahtari -> gBulgu indeksi
Private mGorulen As Object            ' AdaylariTopla icinde tekrar kullanim

'-[ Ana giris ]-------------------------------------------------------

Public Function TaramaYap(ByRef hata As String) As Boolean
    Dim doc As Object
    Dim kok As Object
    Dim s As Object
    Dim i As Long
    Dim bas As Double

    On Error GoTo Hata

    hata = ""
    If Application.Documents.Count = 0 Then
        hata = "Acik belge yok. Once bir dosya acin."
        Exit Function
    End If
    Set doc = Application.ActiveDocument

    DurumSifirla
    gBirimKatsayi = BirimKatsayi(doc.Unit, gBirimAdi)
    If gBirimAdi = "piksel" Or gBirimAdi = "bilinmiyor" Then
        gUyari = gUyari & "Belge birimi (" & gBirimAdi & ") mm'ye kesin cevrilemiyor; tolerans mm kabul edildi." & vbCrLf
    End If

    bas = Timer

    ' Kok nesneler
    If LCase$(gKapsam) = "secim" Then
        If Application.ActiveSelectionRange Is Nothing Then
            hata = "Secim bos. Once nesneleri secin (veya kapsami 'sayfa' yapin)."
            Exit Function
        End If
        If Application.ActiveSelectionRange.Count = 0 Then
            hata = "Secim bos. Once nesneleri secin (veya kapsami 'sayfa' yapin)."
            Exit Function
        End If
        Set kok = Application.ActiveSelectionRange
    Else
        Set kok = doc.ActivePage.Shapes
    End If

    Set mTamCift = CreateObject("Scripting.Dictionary")
    Set mBulguDedup = CreateObject("Scripting.Dictionary")
    mTamCift.CompareMode = 0
    mBulguDedup.CompareMode = 0

    For Each s In kok
        NesneTara s, 0
    Next s

    YapisalTamTespit
    SegmentTarama
    BulguTamamla
    BulguSirala

    TaramaYap = True

    Debug.Print "Tarama: " & Format$(Timer - bas, "0.00") & " sn, " & _
                gNesneSayisi & " nesne, " & gAltYolSayisi & " alt yol, " & _
                gSegSayisi & " segment, " & gBulguSayisi & " bulgu"
    Exit Function

Hata:
    hata = "Tarama hatasi: " & Err.Number & " - " & Err.Description
    TaramaYap = False
End Function

'-[ Nesne dolasimi ]--------------------------------------------------

Private Sub NesneTara(ByVal s As Object, ByVal derinlik As Long)
    Dim tur As Long
    Dim i As Long
    Dim pc As Object
    Dim katman As String

    On Error GoTo Hata

    If derinlik > 12 Then Exit Sub

    tur = s.Type

    ' Grup: icindeki nesnelere in
    If tur = 7 Then                                  ' cdrGroupShape
        For i = 1 To s.Shapes.Count
            NesneTara s.Shapes(i), derinlik + 1
        Next i
        Exit Sub
    End If

    ' Efekt gruplari (blend, extrude, contour, bevel, drop shadow, artistic media, custom effect)
    If tur = 10 Or tur = 11 Or tur = 13 Or tur = 15 Or tur = 16 Or tur = 18 Or tur = 22 Then
        gEfektAtlanan = gEfektAtlanan + 1
        Exit Sub
    End If

    ' Vektor olmayan / kesimle ilgisiz nesneler
    If tur = 5 Or tur = 9 Or tur = 12 Or tur = 17 Or tur = 20 Then
        gAtlanan = gAtlanan + 1
        Exit Sub
    End If

    ' Metin: kesim hatti degil (plotter dosyalarinda nadir, genelde donusmus olur)
    If tur = 6 Then
        gAtlanan = gAtlanan + 1
        Exit Sub
    End If

    katman = ""
    katman = s.Layer.Name
    If StrComp(katman, ISARET_KATMAN, vbTextCompare) = 0 Then Exit Sub
    If KatmanYoksayilir(katman) Then Exit Sub

    ' PowerClip: hem kap hem icerik taranir
    Set pc = Nothing
    On Error Resume Next
    Set pc = s.PowerClip
    If Err.Number <> 0 Then
        Err.Clear
        Set pc = Nothing
    End If
    On Error GoTo Hata
    If Not pc Is Nothing Then
        On Error Resume Next
        For i = 1 To pc.Shapes.Count
            NesneTara pc.Shapes(i), derinlik + 1
        Next i
        If Err.Number <> 0 Then
            gPowerClipAtlanan = gPowerClipAtlanan + 1
            Err.Clear
        End If
        On Error GoTo Hata
    End If

    CurveTara s
    Exit Sub

Hata:
    gAtlanan = gAtlanan + 1
    Err.Clear
End Sub

'-[ Tek nesne: curve -> alt yollar -> segmentler ]-------------------

Private Sub CurveTara(ByVal s As Object)
    Dim crv As Object
    Dim sp As Object
    Dim idx As Long
    Dim i As Long
    Dim uyum As Boolean

    On Error GoTo Hata

    ' 1) Geometriyi al (parametrik sekillerde DisplayCurve sart)
    Set crv = CurveAl(s)
    If crv Is Nothing Then
        gAtlanan = gAtlanan + 1
        Exit Sub
    End If

    uyum = BboxUyumlu(s, crv)

    If crv.SubPaths.Count = 0 Then
        ' Bazi nesnelerde ilk kaynak bos olabilir: digerini dene
        Set crv = CurveYedek(s)
        If crv Is Nothing Then
            UyariEkle "Nesne " & s.StaticID & ": geometri alinamadi (tip " & s.Type & "), atlandi."
            gAtlanan = gAtlanan + 1
            Exit Sub
        End If
        If crv.SubPaths.Count = 0 Then
            UyariEkle "Nesne " & s.StaticID & ": bos geometri (tip " & s.Type & "), atlandi."
            gAtlanan = gAtlanan + 1
            Exit Sub
        End If
        uyum = BboxUyumlu(s, crv)
    End If

    If Not uyum Then
        Set crv = CurveYedek(s)
        If crv Is Nothing Then
            UyariEkle "Nesne " & s.StaticID & ": koordinat uzayi dogrulanamadi, atlandi."
            gAtlanan = gAtlanan + 1
            Exit Sub
        End If
        If Not BboxUyumlu(s, crv) Then
            UyariEkle "Nesne " & s.StaticID & ": donusum uyumsuz, atlandi."
            gAtlanan = gAtlanan + 1
            Exit Sub
        End If
    End If

    NesneYerAc()
    idx = gNesneSayisi
    With gNesne(idx)
        .Ad = NesneAdi(s)
        .Katman = s.Layer.Name
        .Renk = RenkMetni(s)
        .Kalinlik = s.Outline.Width * gBirimKatsayi
        .StatikID = s.StaticID
        .Z = s.ZOrder
        .AltYolSayisi = crv.SubPaths.Count
        .Uzunluk = 0
    End With
    Set gNesneRef(idx) = s
    gNesneSayisi = gNesneSayisi + 1

    For i = 1 To crv.SubPaths.Count
        Set sp = crv.SubPaths(i)
        AltYolIsle idx, i, sp
        If i Mod 25 = 0 Then Durum "Taranıyor: nesne " & (idx + 1) & " / alt yol " & i
    Next i
    Exit Sub

Hata:
    gAtlanan = gAtlanan + 1
    Err.Clear
End Sub

Private Function CurveAl(ByVal s As Object) As Object
    ' DisplayCurve: parametrik sekiller (dikdortgen/elips/poligon), grup icindeki
    ' donusmus nesneler ve normal egriler icin geometriyi verir.
    ' (CorelDRAW 2026'da dogrulandi: Shape.Curve parametrik sekillerde bos doner.)
    On Error Resume Next
    Set CurveAl = s.DisplayCurve
    If CurveAl Is Nothing Then Set CurveAl = s.Curve
    If Not CurveAl Is Nothing Then
        If CurveAl.SubPaths.Count = 0 Then Set CurveAl = s.Curve
    End If
    Err.Clear
End Function

Private Function CurveYedek(ByVal s As Object) As Object
    On Error Resume Next
    Set CurveYedek = s.Curve
    If CurveYedek Is Nothing Then Set CurveYedek = s.DisplayCurve
    Err.Clear
End Function

Private Function BboxUyumlu(ByVal s As Object, ByVal crv As Object) As Boolean
    Dim a As Object, b As Object
    Dim pay As Double
    On Error GoTo Hata

    Set a = crv.BoundingBox
    Set b = s.BoundingBox

    pay = 1# + KalinlikPay(s)

    BboxUyumlu = (Abs((a.Right - a.Left) * gBirimKatsayi - (b.Right - b.Left) * gBirimKatsayi) <= pay) And _
                 (Abs((a.Top - a.Bottom) * gBirimKatsayi - (b.Top - b.Bottom) * gBirimKatsayi) <= pay) And _
                 (Abs((a.Left - b.Left) * gBirimKatsayi) <= pay) And _
                 (Abs((a.Top - b.Top) * gBirimKatsayi) <= pay)
    Exit Function
Hata:
    BboxUyumlu = True      ' olculemiyorsa engellemeyelim
End Function

Private Function KalinlikPay(ByVal s As Object) As Double    On Error Resume Next
    KalinlikPay = Abs(s.Outline.Width * gBirimKatsayi) * 2
End Function

Private Sub AltYolIsle(ByVal nesneIdx As Long, ByVal spNo As Long, ByVal sp As Object)
    Dim slot As Long
    Dim nokta As Long

    On Error GoTo Hata

    AltYolYerAc()
    slot = gAltYolSayisi
    gAltYolNesne(slot) = nesneIdx
    gAltYolNo(slot) = spNo
    gAltYolKapali(slot) = sp.Closed
    gAltYolUzunluk(slot) = 0
    gAltYolX(slot) = 0
    gAltYolY(slot) = 0
    gAltYolImza(slot) = ""
    gAltYolSayisi = gAltYolSayisi + 1

    nokta = Duzlestir(slot, sp)

    gAltYolImza(slot) = AltYolImzasi(sp)
    gNesne(nesneIdx).Uzunluk = gNesne(nesneIdx).Uzunluk + gAltYolUzunluk(slot)

    ' Sifir uzunluklu segment / ust uste dugum
    If nokta > 0 Then
        BulguEkle "NOKTAS", nesneIdx, spNo, -1, 0, 0, 0, gAltYolX(slot), gAltYolY(slot), _
                  "elle", nokta & " adet sifir uzunluklu segment"
    End If
    Exit Sub

Hata:
    gAtlanan = gAtlanan + 1
    Err.Clear
End Sub

'-[ Duzlestirme: segment -> coklu nokta ]-----------------------------
'  Donus: sifir uzunluklu segment sayisi

Private Function Duzlestir(ByVal slot As Long, ByVal sp As Object) As Long
    Dim segs As Object, seg As Object
    Dim n As Long, i As Long, j As Long
    Dim typ As Long
    Dim px As Double, py As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double
    Dim adim As Long
    Dim t As Double, qx As Double, qy As Double
    Dim k As Double
    Dim sifir As Long
    Dim sx As Double, sy As Double, sn As Long

    On Error GoTo Hata

    Set segs = sp.Segments
    n = segs.Count
    If n < 1 Then Exit Function

    k = gBirimKatsayi

    Set seg = segs(1)
    px = seg.StartNode.PositionX * k
    py = seg.StartNode.PositionY * k

    For i = 1 To n
        Set seg = segs(i)
        typ = seg.Type
        x2 = seg.EndNode.PositionX * k
        y2 = seg.EndNode.PositionY * k

        If typ = 0 Then                          ' cdrLineSegment
            If SegmentEkle(slot, px, py, x2, y2) = False Then sifir = sifir + 1
            sx = sx + px + x2: sy = sy + py + y2: sn = sn + 2
            px = x2: py = y2
        Else                                     ' cdrCurveSegment (kubik Bezier)
            c1x = seg.StartingControlPointX * k
            c1y = seg.StartingControlPointY * k
            c2x = seg.EndingControlPointX * k
            c2y = seg.EndingControlPointY * k

            adim = CLng(Sqr((x2 - px) * (x2 - px) + (y2 - py) * (y2 - py)) / 0.7) + 4
            If adim < 6 Then adim = 6
            If adim > 48 Then adim = 48

            For j = 1 To adim
                t = j / adim
                BezierNokta px, py, c1x, c1y, c2x, c2y, x2, y2, t, qx, qy
                If SegmentEkle(slot, px, py, qx, qy) = False Then sifir = sifir + 1
                sx = sx + px + qx: sy = sy + py + qy: sn = sn + 2
                px = qx: py = qy
            Next j
        End If
    Next i

    ' Isaret / rapor konumu: noktalarin agirlik merkezi
    If sn > 0 Then
        gAltYolX(slot) = sx / sn
        gAltYolY(slot) = sy / sn
    End If

    Duzlestir = sifir
    Exit Function

Hata:
    Duzlestir = sifir
End Function

Private Sub BezierNokta(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, _
                         ByVal t As Double, ByRef ox As Double, ByRef oy As Double)
    Dim u As Double
    Dim a As Double, b As Double, c As Double, d As Double
    u = 1 - t
    a = u * u * u
    b = 3 * u * u * t
    c = 3 * u * t * t
    d = t * t * t
    ox = a * x0 + b * x1 + c * x2 + d * x3
    oy = a * y0 + b * y1 + c * y2 + d * y3
End Sub

Private Function SegmentEkle(ByVal slot As Long, ByVal x1 As Double, ByVal y1 As Double, _
                             ByVal x2 As Double, ByVal y2 As Double) As Boolean
    Dim boy As Double
    boy = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))

    If boy < 0.005 Then
        SegmentEkle = False
        Exit Function
    End If

    SegmentYerAc()
    gSegX1(gSegSayisi) = x1
    gSegY1(gSegSayisi) = y1
    gSegX2(gSegSayisi) = x2
    gSegY2(gSegSayisi) = y2
    gSegNesne(gSegSayisi) = gAltYolNesne(slot)
    gSegAltYol(gSegSayisi) = slot

    gAltYolUzunluk(slot) = gAltYolUzunluk(slot) + boy

    gSegSayisi = gSegSayisi + 1
    SegmentEkle = True
End Function

'-[ Yapisal TAM: dugum duzeyinde birebir kopya ]---------------------

Public Sub YapisalTamTespit()
    Dim d As Object, i As Long, k As String
    Dim p() As String, j As Long
    Dim tutucu As Long, aday As Long
    Dim neden As String
    Dim eylem As String
    Dim ilkZ As Long

    If gAltYolSayisi = 0 Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0

    For i = 0 To gAltYolSayisi - 1
        k = gAltYolImza(i)
        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = d(k) & "," & CStr(i)
            Else
                d.Add k, CStr(i)
            End If
        End If
    Next i

    Dim anahtar As Variant
    For Each anahtar In d.Keys
        p = Split(d(anahtar), ",")
        If UBound(p) >= 1 Then
            ' Tutucu: en altta kalan kopya korunur.
            ' CorelDRAW'da ZOrder 1 = en ustte; buyuk Z = altta (2026'da dogrulandi).
            tutucu = CLng(p(0))
            ilkZ = gNesne(gAltYolNesne(tutucu)).Z
            For j = 1 To UBound(p)
                aday = CLng(p(j))
                If gNesne(gAltYolNesne(aday)).Z > ilkZ Then
                    ilkZ = gNesne(gAltYolNesne(aday)).Z
                    tutucu = aday
                End If
            Next j

            For j = 0 To UBound(p)
                aday = CLng(p(j))
                If aday <> tutucu Then
                    If gAltYolNesne(aday) = gAltYolNesne(tutucu) Then
                        eylem = "sil"
                        neden = "ayni nesne icinde tekrar"
                    ElseIf AyniOzellik(gAltYolNesne(aday), gAltYolNesne(tutucu), neden) Then
                        eylem = IIf(gOtomatikSil, "sil", "elle")
                    Else
                        eylem = "elle"
                    End If

                    BulguEkle "TAM", gAltYolNesne(aday), gAltYolNo(aday), _
                              gAltYolNesne(tutucu), gAltYolNo(tutucu), _
                              gAltYolUzunluk(aday), 1, gAltYolX(aday), gAltYolY(aday), _
                              eylem, IIf(Len(neden) > 0, "birebir kopya (" & neden & ")", "birebir kopya")

                    TamCiftEkle gAltYolNesne(aday), gAltYolNo(aday), _
                                gAltYolNesne(tutucu), gAltYolNo(tutucu)
                    neden = ""
                End If
            Next j
        End If
    Next anahtar
End Sub

Private Sub TamCiftEkle(ByVal n1 As Long, ByVal a1 As Long, ByVal n2 As Long, ByVal a2 As Long)
    Dim k As String
    k = CStr(n1) & "|" & CStr(a1) & "|" & CStr(n2) & "|" & CStr(a2)
    If Not mTamCift.Exists(k) Then mTamCift.Add k, 1
    k = CStr(n2) & "|" & CStr(a2) & "|" & CStr(n1) & "|" & CStr(a1)
    If Not mTamCift.Exists(k) Then mTamCift.Add k, 1
End Sub

Private Function TamCiftMi(ByVal n1 As Long, ByVal a1 As Long, ByVal n2 As Long, ByVal a2 As Long) As Boolean
    TamCiftMi = mTamCift.Exists(CStr(n1) & "|" & CStr(a1) & "|" & CStr(n2) & "|" & CStr(a2))
End Function

'-[ Uzamsal izgara + segment taramasi ]------------------------------

Public Sub SegmentTarama()
    Dim i As Long, j As Long
    Dim k As Variant
    Dim aday As Object
    Dim tur As String
    Dim slotA As Long, slotB As Long
    Dim nA As Long, aA As Long, nB As Long, aB As Long
    Dim anahtar As String, idx As Long
    Dim esik As Long

    If gSegSayisi = 0 Then Exit Sub

    Set mIzgara = CreateObject("Scripting.Dictionary")
    mIzgara.CompareMode = 0

    For i = 0 To gSegSayisi - 1
        IzgarayaEkle i
        If i Mod 2000 = 0 Then Durum "Izgara: " & i & " / " & gSegSayisi
    Next i

    Set aday = CreateObject("Scripting.Dictionary")
    aday.CompareMode = 0
    Set mGorulen = CreateObject("Scripting.Dictionary")
    mGorulen.CompareMode = 0

    esik = 0
    For i = 0 To gSegSayisi - 1
        aday.RemoveAll
        AdaylariTopla i, aday
        If aday.Count > 0 Then
            For Each k In aday.Keys
                j = CLng(k)
                If j > i Then
                    slotA = gSegAltYol(i)
                    slotB = gSegAltYol(j)

                    ' Ayni alt yolda ardisik segmentler karsilastirilmaz
                    If Not (slotA = slotB And (j - i) = 1) Then
                        tur = Karsilastir(i, j)
                        If Len(tur) > 0 Then
                            nA = gAltYolNesne(slotA): aA = gAltYolNo(slotA)
                            nB = gAltYolNesne(slotB): aB = gAltYolNo(slotB)
                            If Not TamCiftMi(nA, aA, nB, aB) Then
                                anahtar = CiftAnahtar(nA, aA, nB, aB) & "|" & tur
                                If mBulguDedup.Exists(anahtar) Then
                                    idx = mBulguDedup(anahtar)
                                    gBulgu(idx).Uzunluk = gBulgu(idx).Uzunluk + gKarsOrtusme
                                Else
                                    idx = BulguEkle(tur, nA, aA, nB, aB, gKarsOrtusme, gKarsOran, _
                                                    gKarsX, gKarsY, "elle", "")
                                    If idx >= 0 Then mBulguDedup.Add anahtar, idx
                                End If
                            End If
                        End If
                    End If
                End If
            Next k
        End If

        esik = esik + 1
        If esik Mod 2000 = 0 Then Durum "Karsilastirma: " & esik & " / " & gSegSayisi
    Next i

    Durum ""
End Sub

Private Function CiftAnahtar(ByVal n1 As Long, ByVal a1 As Long, ByVal n2 As Long, ByVal a2 As Long) As String
    If (n1 < n2) Or (n1 = n2 And a1 <= a2) Then
        CiftAnahtar = CStr(n1) & "|" & CStr(a1) & "|" & CStr(n2) & "|" & CStr(a2)
    Else
        CiftAnahtar = CStr(n2) & "|" & CStr(a2) & "|" & CStr(n1) & "|" & CStr(a1)
    End If
End Function

Private Sub IzgarayaEkle(ByVal idx As Long)
    Dim x As Double, y As Double
    Dim onceki As String, k As String
    Dim adim As Double, boy As Double, t As Double
    Dim n As Long, i As Long

    boy = Sqr((gSegX2(idx) - gSegX1(idx)) ^ 2 + (gSegY2(idx) - gSegY1(idx)) ^ 2)
    adim = gHucreMm / 2
    If adim < 0.25 Then adim = 0.25
    n = CLng(boy / adim) + 1

    onceki = ""
    For i = 0 To n
        If n = 0 Then
            t = 0
        Else
            t = i / n
        End If
        x = gSegX1(idx) + (gSegX2(idx) - gSegX1(idx)) * t
        y = gSegY1(idx) + (gSegY2(idx) - gSegY1(idx)) * t
        k = HucreAnahtar(x, y)
        If k <> onceki Then
            If mIzgara.Exists(k) Then
                mIzgara(k) = mIzgara(k) & CStr(idx) & ";"
            Else
                mIzgara.Add k, CStr(idx) & ";"
            End If
            onceki = k
        End If
    Next i
End Sub

Private Sub AdaylariTopla(ByVal idx As Long, ByVal aday As Object)
    Dim boy As Double, adim As Double, n As Long, i As Long
    Dim t As Double, x As Double, y As Double
    Dim cx As Long, cy As Long, dx As Long, dy As Long
    Dim k As String, p() As String, q As Long

    boy = Sqr((gSegX2(idx) - gSegX1(idx)) ^ 2 + (gSegY2(idx) - gSegY1(idx)) ^ 2)
    ' Sorgu adimi = hucre boyu. 3x3 komsuluk 6 mm'lik alani taradigi icin
    ' 2 mm'de bir ornek yeterli (ekleme adimi hucre/2, daha yogun).
    adim = gHucreMm
    If adim < 0.5 Then adim = 0.5
    n = CLng(boy / adim) + 1

    mGorulen.RemoveAll

    For i = 0 To n
        If n = 0 Then t = 0 Else t = i / n
        x = gSegX1(idx) + (gSegX2(idx) - gSegX1(idx)) * t
        y = gSegY1(idx) + (gSegY2(idx) - gSegY1(idx)) * t
        cx = Int(x / gHucreMm)
        cy = Int(y / gHucreMm)

        For dx = -1 To 1
            For dy = -1 To 1
                k = CStr(cx + dx) & "|" & CStr(cy + dy)
                If Not mGorulen.Exists(k) Then
                    mGorulen.Add k, 1
                    If mIzgara.Exists(k) Then
                        p = Split(mIzgara(k), ";")
                        For q = 0 To UBound(p)
                            If Len(p(q)) > 0 Then aday(p(q)) = 1
                        Next q
                    End If
                End If
            Next dy
        Next dx
    Next i
End Sub

Private Function HucreAnahtar(ByVal x As Double, ByVal y As Double) As String
    HucreAnahtar = CStr(Int(x / gHucreMm)) & "|" & CStr(Int(y / gHucreMm))
End Function

'-[ Bulgu bitirme: oran, eylem, siralama ]---------------------------

Public Sub BulguTamamla()
    Dim i As Long
    Dim boy1 As Double, boy2 As Double, mn As Double
    Dim d As Object, k As String
    Dim neden As String

    ' Alt yol uzunluklari: tek seferlik sozluk (dogrusal arama yerine)
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    For i = 0 To gAltYolSayisi - 1
        k = CStr(gAltYolNesne(i)) & "|" & CStr(gAltYolNo(i))
        If Not d.Exists(k) Then d.Add k, gAltYolUzunluk(i)
    Next i

    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            ' 1) Ortusme orani
            If .Tur <> "NOKTAS" Then
                boy1 = 0: boy2 = 0
                k = CStr(.Nesne1) & "|" & CStr(.AltYol1)
                If d.Exists(k) Then boy1 = d(k)
                If .Nesne2 >= 0 Then
                    k = CStr(.Nesne2) & "|" & CStr(.AltYol2)
                    If d.Exists(k) Then boy2 = d(k)
                Else
                    boy2 = boy1
                End If
                mn = Kucuk(boy1, boy2)
                If mn > 0 Then
                    .Oran = .Uzunluk / mn
                    If .Oran > 1 Then .Oran = 1
                End If
            End If

            ' 2) Segment duzeyinde "TAM" yalnizca o segmentlerin cakistigini soyler.
            '    Alt yolun tamami ortusmuyorsa dogru sinif KISMI'dir.
            If .Tur = "TAM" And .Oran < 0.98 Then .Tur = "KISMI"

            ' 3) Eylem: yalnizca TAM otomatik silinebilir
            If .Tur = "TAM" Then
                If .Eylem <> "sil" Then
                    neden = ""
                    If .Nesne1 <> .Nesne2 Then
                        If AyniOzellik(.Nesne1, .Nesne2, neden) Then .Eylem = "sil"
                    End If
                End If
                If .Eylem = "sil" And Not gOtomatikSil Then .Eylem = "elle"
            Else
                .Eylem = "elle"          ' KISMI / YAKIN / NOKTAS asla otomatik silinmez
            End If
        End With
    Next i
End Sub

Public Sub BulguSirala()
    ' Basit ekleme sirasi: tur onceligi, sonra uzunluk (buyuk -> kucuk)
    Dim i As Long, j As Long
    Dim tmp As TBulgu
    If gBulguSayisi < 2 Then Exit Sub
    If gBulguSayisi > 2000 Then Exit Sub     ' cok buyuk listelerde siralama maliyeti gereksiz

    For i = 1 To gBulguSayisi - 1
        tmp = gBulgu(i)
        j = i - 1
        Do While j >= 0
            If BulguOnce(gBulgu(j), tmp) Then Exit Do
            gBulgu(j + 1) = gBulgu(j)
            j = j - 1
        Loop
        gBulgu(j + 1) = tmp
    Next i
End Sub

Private Function BulguOnce(ByRef a As TBulgu, ByRef b As TBulgu) As Boolean
    Dim pa As Long, pb As Long
    pa = TurPuan(a.Tur): pb = TurPuan(b.Tur)
    If pa <> pb Then
        BulguOnce = (pa < pb)
    Else
        BulguOnce = (a.Uzunluk >= b.Uzunluk)
    End If
End Function

Private Function TurPuan(ByVal tur As String) As Long
    Select Case tur
        Case "TAM": TurPuan = 1
        Case "KISMI": TurPuan = 2
        Case "YAKIN": TurPuan = 3
        Case Else: TurPuan = 4
    End Select
End Function

'-[ Kucuk yardimcilar ]----------------------------------------------

Private Sub UyariEkle(ByVal metin As String)
    If Len(gUyari) < 1000 Then gUyari = gUyari & metin & vbCrLf
End Sub

Private Function NesneAdi(ByVal s As Object) As String
    Dim ad As String
    On Error Resume Next
    ad = s.Name
    If Err.Number <> 0 Or Len(ad) = 0 Then
        Err.Clear
        NesneAdi = "Nesne " & s.StaticID
    Else
        NesneAdi = ad
    End If
End Function

Private Function RenkMetni(ByVal s As Object) As String
    Dim h As String
    On Error Resume Next
    h = s.Outline.Color.HexValue
    If Err.Number <> 0 Then
        Err.Clear
        RenkMetni = "?"
    Else
        RenkMetni = h
    End If
End Function

Private Sub Durum(ByVal metin As String)
    On Error Resume Next
    Application.StatusBar.Text = metin
End Sub
