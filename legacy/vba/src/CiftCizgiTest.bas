Attribute VB_Name = "CiftCizgiTest"
Option Explicit

' =====================================================================
'  cdrcift - Kendi kendini test (fixture uretici + dogrulama)
'  Surum: 1.0.0
'
'  Calistirma: TestleriCalistir makrosu (yeni gecici belge acar).
'  Amac: CorelDRAW disinda birim test yazilamadigi icin regresyon
'  guvencesi. Her faz sonunda bu makro yesil olmalidir.
' =====================================================================

Private mGecen As Long
Private mKalan As Long
Private mRapor As String

'-[ Test calistirici ]-----------------------------------------------

Public Sub TestleriCalistir()
    Dim doc As Object
    Dim lyr As Object
    Dim oncekiAdet As Long, sonrakiAdet As Long
    Dim silinen As Long
    Dim ozet As String
    Dim testDosya As String
    Dim f As Integer

    mGecen = 0: mKalan = 0: mRapor = ""

    ' Test icin sabit ayarlar
    VarsayilanAyarlar
    gKapsam = "sayfa"
    gTolerans = 0.05
    gYakinTol = 0.5
    gMinOrtusme = 0.5
    gOtomatikSil = True
    gBoyama = False
    gMaxBulgu = 5000

    On Error GoTo Hata

    Set doc = Application.CreateDocument()
    doc.Unit = 3                                   ' cdrMillimeter
    Set lyr = doc.ActivePage.ActiveLayer

    FixtureCiz lyr

    If Not TaramaYap(ozet) Then
        Kontrol False, "Tarama calistirilamadi: " & ozet
        GoTo Rapor
    End If

    ' 1) Birebir kopya (ayni koordinat, iki ayri nesne)
    Kontrol Say("TAM") >= 1, "Birebir kopya TAM olarak bulunmali (beklenen >=1, gelen " & Say("TAM") & ")"
    Kontrol VarMi("TAM", 10, 35, 1), "A/B kopyasi (10,35) TAM olarak bulunmali"
    Kontrol Eylem("TAM", 10, 35, 1) = "sil", "A/B kopyasi otomatik silinebilir olmali"

    ' 2) Ters yonde cizilmis kopya
    Kontrol VarMi("TAM", 110, 35, 1), "Ters yondeki kopya (110,35) TAM olarak bulunmali"

    ' 3) Ortak kenar: iki komsu dikdortgen
    Kontrol VarMi("KISMI", 240, 30, 2), "Ortak kenar (240,30) KISMI olarak bulunmali"
    Kontrol Eylem("KISMI", 240, 30, 2) = "elle", "KISMI bulgu asla otomatik silinmemeli"

    ' 4) 0,3 mm kaymis kopya -> YAKIN
    Kontrol VarMi("YAKIN", 300.15, 35, 1), "0,3 mm kaymis cizgi YAKIN olarak bulunmali"

    ' 5) 0,1 mm kaymis cizgi TAM sayilmamali (yanlis pozitif testi)
    Kontrol (Not VarMi("TAM", 320.05, 35, 2)), "0,1 mm kaymis cizgi TAM sayilmamali"

    ' 6) Kismi/yakin bulgularin hicbiri silinemez olmamali
    Kontrol SilinebilirKismiYok(), "KISMI/YAKIN bulgularda 'sil' eylemi olmamali"

    ' 7) Silme: en az 2 kopya nesne kalkmali
    oncekiAdet = doc.ActivePage.Shapes.Count
    silinen = Sil(ozet)
    sonrakiAdet = doc.ActivePage.Shapes.Count
    Kontrol silinen >= 2, "En az 2 kopya silinmeli (silinen: " & silinen & ")"
    Kontrol sonrakiAdet < oncekiAdet, "Nesne sayisi azalmali (" & oncekiAdet & " -> " & sonrakiAdet & ")"

    ' 8) Geri alma: silme tek adimda geri alinmali
    doc.Undo 1
    Kontrol doc.ActivePage.Shapes.Count = oncekiAdet, "Ctrl+Z sonrasi nesne sayisi geri gelmeli"

Rapor:
    doc.Dirty = False
    On Error Resume Next
    doc.Close
    Err.Clear
    On Error GoTo 0

    testDosya = RaporKlasoru() & "\test-sonuc.txt"
    If KlasorOlustur(RaporKlasoru()) Then
        f = FreeFile
        Open testDosya For Output As #f
        Print #f, "cdrcift test sonucu - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Print #f, "Gecen: " & mGecen & "  Kalan: " & mKalan
        Print #f, ""
        Print #f, mRapor
        Close #f
    End If

    MsgBox "TEST SONUCU" & vbCrLf & vbCrLf & _
           "Gecen: " & mGecen & "   Kalan: " & mKalan & vbCrLf & vbCrLf & _
           IIf(mKalan = 0, "Tum testler basarili.", "Basarisiz testler icin rapor:") & vbCrLf & _
           IIf(mKalan = 0, "", testDosya & vbCrLf & vbCrLf & Kisalt(mRapor, 700)), _
           IIf(mKalan = 0, vbInformation, vbExclamation), UYGULAMA_ADI
    Exit Sub

Hata:
    MsgBox "Test hatasi: " & Err.Number & " - " & Err.Description, vbCritical, UYGULAMA_ADI
    On Error Resume Next
    doc.Dirty = False
    doc.Close
    Err.Clear
End Sub

'-[ Fixture uret ve incele (kullaniciya demo) ]----------------------

Public Sub FixtureUret()
    Dim doc As Object, lyr As Object
    Dim hata As String

    On Error GoTo Hata
    Set doc = Application.CreateDocument()
    doc.Unit = 3
    Set lyr = doc.ActivePage.ActiveLayer
    FixtureCiz lyr

    VarsayilanAyarlar
    gKapsam = "sayfa"
    gOtomatikSil = True
    TaramaYap hata
    If gBulguSayisi > 0 Then Isaretle

    MsgBox "Demo belge olusturuldu." & vbCrLf & _
           "Bulgu: " & gBulguSayisi & vbCrLf & vbCrLf & _
           "Denetle / Temizle makrolarini bu belgede deneyebilirsiniz.", _
           vbInformation, UYGULAMA_ADI
    Exit Sub
Hata:
    MsgBox "Fixture hatasi: " & Err.Description, vbCritical, UYGULAMA_ADI
End Sub

Private Sub FixtureCiz(ByVal lyr As Object)
    ' 1) Birebir kopya: ayni cizgi iki kez
    lyr.CreateLineSegment 10, 10, 10, 60
    lyr.CreateLineSegment 10, 10, 10, 60

    ' 2) Ters yonde cizilmis kopya
    lyr.CreateLineSegment 110, 10, 110, 60
    lyr.CreateLineSegment 110, 60, 110, 10

    ' 3) Ortak kenarli iki dikdortgen (x=240 kenari ortak)
    lyr.CreateRectangle 200, 50, 240, 10
    lyr.CreateRectangle 240, 50, 280, 10

    ' 4) 0,3 mm kaymis kopya -> YAKIN
    lyr.CreateLineSegment 300, 10, 300, 60
    lyr.CreateLineSegment 300.3, 10, 300.3, 60

    ' 5) 0,1 mm kaymis cizgi -> TAM sayilmamali
    lyr.CreateLineSegment 320, 10, 320, 60
    lyr.CreateLineSegment 320.1, 10, 320.1, 60

    ' 6) Sifir uzunluklu segment (desteklenmiyorsa sessizce atlanir)
    On Error Resume Next
    lyr.CreateCurveSegment 350, 10, 350, 10, 350, 10, 350, 10
    Err.Clear
    On Error GoTo 0
End Sub

'-[ Teshis: secili nesnenin alt yol dokumu ]-------------------------
'  Faz 0 spike: koordinat uzayi, birim ve alt yol yapisini dogrular.
'  Cikti: VBA IDE "Hemen Penceresi" (Ctrl+G).

Public Sub AltYolDok()
    Dim doc As Object, s As Object, crv As Object, sp As Object
    Dim i As Long, j As Long
    Dim n As Long
    Dim birimKod As Long

    On Error GoTo Hata

    If Application.Documents.Count = 0 Then
        MsgBox "Acik belge yok.", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If
    If Application.ActiveSelectionRange Is Nothing Then
        MsgBox "Once nesne secin.", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Set doc = Application.ActiveDocument
    birimKod = doc.Unit
    VarsayilanAyarlar
    gBirimKatsayi = BirimKatsayi(doc.Unit, gBirimAdi)

    Debug.Print "=== cdrcift AltYolDok ==="
    Debug.Print "Belge birimi: " & gBirimAdi & " (kod " & birimKod & "), 1 birim = " & gBirimKatsayi & " mm"
    Debug.Print "Secim adedi: " & Application.ActiveSelectionRange.Count

    For i = 1 To Application.ActiveSelectionRange.Count
        Set s = Application.ActiveSelectionRange(i)
        Debug.Print "---"
        Debug.Print "Nesne " & i & " : StaticID=" & s.StaticID & "  Tip=" & s.Type & _
                    "  Katman=" & s.Layer.Name & "  Z=" & s.ZOrder
        Debug.Print "  Ad=" & NesneAdiTest(s) & "  Kalinlik=" & (s.Outline.Width * gBirimKatsayi) & " mm" & _
                    "  Renk=" & HexMetin(s)

        Set crv = Nothing
        On Error Resume Next
        Set crv = s.DisplayCurve          ' parametrik sekillerde tek kaynak
        If crv Is Nothing Then Set crv = s.Curve
        If Err.Number <> 0 Then
            Err.Clear
            Debug.Print "  HATA: Curve alinamadi"
        End If
        On Error GoTo Hata

        If Not crv Is Nothing Then
            n = crv.SubPaths.Count
            Debug.Print "  Alt yol sayisi: " & n & "  Segment: " & crv.Segments.Count & _
                        "  Dugum: " & crv.Nodes.Count
            Debug.Print "  DisplayCurve bbox: " & BboxMetni(crv.BoundingBox)
            Debug.Print "  Nesne bbox       : " & BboxMetni(s.BoundingBox)
            On Error Resume Next
            Debug.Print "  Shape.Curve alt yol: " & s.Curve.SubPaths.Count & _
                        " (parametrik sekilde 0 olabilir)"
            Err.Clear
            On Error GoTo Hata

            For j = 1 To n
                Set sp = crv.SubPaths(j)
                Debug.Print "   alt yol " & j & ": kapali=" & sp.Closed & _
                            " segment=" & sp.Segments.Count & " dugum=" & sp.Nodes.Count & _
                            " uzunluk=" & (sp.Length * gBirimKatsayi) & " mm"
                If j <= 3 Then
                    Debug.Print "     ilk nokta: " & _
                        (sp.StartNode.PositionX * gBirimKatsayi) & ", " & _
                        (sp.StartNode.PositionY * gBirimKatsayi) & " mm"
                    Debug.Print "     imza     : " & Kisalt(AltYolImzasi(sp), 120)
                End If
            Next j
        End If
    Next i

    Debug.Print "=== bitti ==="
    MsgBox "Dokum VBA IDE 'Hemen Penceresi'ne yazildi (Ctrl+G).", vbInformation, UYGULAMA_ADI
    Exit Sub

Hata:
    MsgBox "Teshis hatasi: " & Err.Description, vbCritical, UYGULAMA_ADI
    Err.Clear
End Sub

Private Function NesneAdiTest(ByVal s As Object) As String
    Dim ad As String
    On Error Resume Next
    ad = s.Name
    If Err.Number <> 0 Or Len(ad) = 0 Then
        Err.Clear
        NesneAdiTest = "Nesne " & s.StaticID
    Else
        NesneAdiTest = ad
    End If
End Function

Private Function HexMetin(ByVal s As Object) As String
    On Error Resume Next
    HexMetin = s.Outline.Color.HexValue
    If Err.Number <> 0 Then
        Err.Clear
        HexMetin = "?"
    End If
End Function

Private Function BboxMetni(ByVal bb As Object) As String
    On Error Resume Next
    BboxMetni = "L=" & (bb.Left * gBirimKatsayi) & " B=" & (bb.Bottom * gBirimKatsayi) & _
               " R=" & (bb.Right * gBirimKatsayi) & " T=" & (bb.Top * gBirimKatsayi) & " mm"
    If Err.Number <> 0 Then
        Err.Clear
        BboxMetni = "?"
    End If
End Function

'-[ Kucuk test yardimcilari ]----------------------------------------

Private Sub Kontrol(ByVal kosul As Boolean, ByVal aciklama As String)
    If kosul Then
        mGecen = mGecen + 1
        mRapor = mRapor & "[OK]   " & aciklama & vbCrLf
    Else
        mKalan = mKalan + 1
        mRapor = mRapor & "[HATA] " & aciklama & vbCrLf
    End If
End Sub

Private Function Say(ByVal tur As String) As Long
    Dim i As Long, n As Long
    For i = 0 To gBulguSayisi - 1
        If gBulgu(i).Tur = tur Then n = n + 1
    Next i
    Say = n
End Function

Private Function VarMi(ByVal tur As String, ByVal x As Double, ByVal y As Double, ByVal pay As Double) As Boolean
    Dim i As Long
    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            If .Tur = tur Then
                If Abs(.X - x) <= pay And Abs(.Y - y) <= pay Then
                    VarMi = True
                    Exit Function
                End If
            End If
        End With
    Next i
End Function

Private Function Eylem(ByVal tur As String, ByVal x As Double, ByVal y As Double, ByVal pay As Double) As String
    Dim i As Long
    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            If .Tur = tur Then
                If Abs(.X - x) <= pay And Abs(.Y - y) <= pay Then
                    Eylem = .Eylem
                    Exit Function
                End If
            End If
        End With
    Next i
    Eylem = "(yok)"
End Function

Private Function SilinebilirKismiYok() As Boolean
    Dim i As Long
    For i = 0 To gBulguSayisi - 1
        If gBulgu(i).Tur = "KISMI" Or gBulgu(i).Tur = "YAKIN" Then
            If gBulgu(i).Eylem = "sil" Then Exit Function
        End If
    Next i
    SilinebilirKismiYok = True
End Function
