Attribute VB_Name = "CiftCizgi"
Option Explicit

' =====================================================================
'  cdrcift - Cift cizgi denetleyicisi / giris noktalari
'  Surum: 1.0.0        CorelDRAW 2024 / 2025 / 2026 (VBA)
'
'  Makrolar:
'    Denetle            : Tara, raporla, isaretle (varsa temizleme onerisi)
'    Temizle            : Tara ve yalnizca birebir (TAM) kopyalari sil
'    IsaretleriKaldir   : DUP_CHECK katmanini ve isaretleri kaldir
'    Ayarlar            : Tolerans / kapsam ayarlarini duzenle
'    Yardim             : Kullanim ozeti
' =====================================================================

'-[ Ana makro: Denetle ]---------------------------------------------

Public Sub Denetle()
    Dim hata As String
    Dim ok As Boolean
    Dim raporTxt As String, raporCsv As String
    Dim cevap As VbMsgBoxResult
    Dim silSayisi As Long
    Dim ozet As String
    Dim eskiOpt As Boolean, eskiEv As Boolean

    AyarlariYukle
    If Not BelgeVarMi(hata) Then
        MsgBox hata, vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    eskiOpt = Application.Optimization
    eskiEv = Application.EventsEnabled
    Application.Optimization = True
    Application.EventsEnabled = False
    On Error GoTo Son

    ok = TaramaYap(hata)
    If Not ok Then
        MsgBox hata, vbCritical, UYGULAMA_ADI
        GoTo Son
    End If

    If RaporYaz(raporTxt, raporCsv, hata) = False Then
        MsgBox "Rapor yazilamadi: " & hata & vbCrLf & _
               "Sonuclar ekranda ozetlendi.", vbExclamation, UYGULAMA_ADI
    End If

    If gBulguSayisi = 0 Then
        MsgBox "Sonuc: cift cizgi bulunmadi." & vbCrLf & vbCrLf & _
               "Taranan: " & gNesneSayisi & " nesne, " & gAltYolSayisi & " alt yol, " & _
               gSegSayisi & " segment." & vbCrLf & _
               IIf(gAtlanan > 0, "Atlanan nesne: " & gAtlanan & vbCrLf, "") & _
               IIf(gEfektAtlanan > 0, "Efekt nesnesi (blend/contour vb.): " & gEfektAtlanan & vbCrLf, ""), _
               vbInformation, UYGULAMA_ADI
        GoTo Son
    End If

    Isaretle
    cevap = MsgBox(OzetMetni(raporTxt, raporCsv) & vbCrLf & _
                   "Bulunan yerler DUP_CHECK katmaninda isaretlendi." & vbCrLf & vbCrLf & _
                   SilinebilirOzet() & vbCrLf & vbCrLf & _
                   "Simdi otomatik silme yapilsin mi?" & vbCrLf & _
                   "(Tek Ctrl+Z ile geri alinabilir.)", _
                   vbYesNo + vbQuestion, UYGULAMA_ADI)

    If cevap = vbYes Then
        silSayisi = Sil(ozet)
        If silSayisi > 0 Then
            MsgBox ozet, vbInformation, UYGULAMA_ADI
        Else
            MsgBox "Otomatik silinecek uygun kayit yok." & vbCrLf & _
                   "Kismi/yakin cakismalar guvenlik geregi silinmez; " & _
                   "rapor numaralari ile elle duzeltin.", vbInformation, UYGULAMA_ADI
        End If
    End If

Son:
    Application.EventsEnabled = eskiEv
    Application.Optimization = eskiOpt
    If Err.Number <> 0 Then
        MsgBox "Beklenmeyen hata: " & Err.Number & " - " & Err.Description, vbCritical, UYGULAMA_ADI
        Err.Clear
    End If
End Sub

'-[ Temizle ]---------------------------------------------------------

Public Sub Temizle()
    Dim hata As String, ozet As String
    Dim ok As Boolean
    Dim n As Long
    Dim eskiOpt As Boolean, eskiEv As Boolean

    AyarlariYukle
    If Not BelgeVarMi(hata) Then
        MsgBox hata, vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    If Application.ActiveDocument.Dirty Then
        If MsgBox("Belgede kaydedilmemis degisiklikler var." & vbCrLf & _
                  "Temizlik oncesi dosyayi kaydetmeniz onerilir." & vbCrLf & vbCrLf & _
                  "Devam edilsin mi?", vbYesNo + vbExclamation, UYGULAMA_ADI) <> vbYes Then Exit Sub
    End If

    eskiOpt = Application.Optimization
    eskiEv = Application.EventsEnabled
    Application.Optimization = True
    Application.EventsEnabled = False
    On Error GoTo Son

    ok = TaramaYap(hata)
    If Not ok Then
        MsgBox hata, vbCritical, UYGULAMA_ADI
        GoTo Son
    End If

    n = Sil(ozet)
    If n = 0 Then
        MsgBox "Silinecek birebir kopya bulunamadi." & vbCrLf & vbCrLf & SilinebilirOzet(), _
               vbInformation, UYGULAMA_ADI
    Else
        MsgBox ozet, vbInformation, UYGULAMA_ADI
    End If
    GoTo Son

Son:
    Application.EventsEnabled = eskiEv
    Application.Optimization = eskiOpt
    If Err.Number <> 0 Then
        MsgBox "Beklenmeyen hata: " & Err.Number & " - " & Err.Description, vbCritical, UYGULAMA_ADI
        Err.Clear
    End If
End Sub

'-[ Silme cekirdegi ]-------------------------------------------------

Public Function Sil(ByRef ozet As String) As Long
    Dim doc As Object
    Dim silAltYol As Object
    Dim silNesne As Object
    Dim i As Long, k As String
    Dim altSil As Long, nesneSil As Long
    Dim liste() As Long, adet As Long
    Dim j As Long, gecici As Long
    Dim s As Object
    Dim sp As Object
    Dim altListe() As Long, altAdet As Long
    Dim sayim As Object, kendKopya As Object
    Dim anahtar As Variant, anahtar2 As Variant
    Dim parca() As String

    ozet = ""
    Sil = 0
    If gBulguSayisi = 0 Then Exit Function
    If Not gOtomatikSil Then Exit Function

    Set doc = Application.ActiveDocument
    Set silAltYol = CreateObject("Scripting.Dictionary")
    Set silNesne = CreateObject("Scripting.Dictionary")
    silAltYol.CompareMode = 0

    ' 1) Silinecek alt yollari topla (yalniz TAM + Eylem=sil)
    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            If .Tur = "TAM" And .Eylem = "sil" Then
                k = CStr(.Nesne1) & "|" & CStr(.AltYol1)
                If Not silAltYol.Exists(k) Then silAltYol.Add k, 1
            End If
        End With
    Next i
    If silAltYol.Count = 0 Then Exit Function

    ' 2) Butun alt yollari silinen nesneler: nesneyi tumden sil
    Set sayim = CreateObject("Scripting.Dictionary")
    Set kendKopya = CreateObject("Scripting.Dictionary")

    For Each anahtar2 In silAltYol.Keys
        parca = Split(CStr(anahtar2), "|")
        k = parca(0)
        If sayim.Exists(k) Then sayim(k) = sayim(k) + 1 Else sayim.Add k, 1
    Next anahtar2

    For i = 0 To gBulguSayisi - 1
        With gBulgu(i)
            If .Tur = "TAM" And .Eylem = "sil" And .Nesne1 = .Nesne2 Then
                k = CStr(.Nesne1)
                If Not kendKopya.Exists(k) Then kendKopya.Add k, 1
            End If
        End With
    Next i

    For i = 0 To gNesneSayisi - 1
        k = CStr(i)
        If sayim.Exists(k) Then
            If sayim(k) >= gNesne(i).AltYolSayisi Then
                If gNesne(i).AltYolSayisi > 0 Then
                    If Not kendKopya.Exists(k) Then silNesne.Add k, i
                End If
            End If
        End If
    Next i

    doc.BeginCommandGroup "Cift cizgi temizligi"

    ' 3) Nesneleri Z sirasina gore (ustten alta) sil
    adet = silNesne.Count
    If adet > 0 Then
        ReDim liste(0 To adet - 1)
        j = 0
        For Each anahtar In silNesne.Keys
            liste(j) = CLng(anahtar)
            j = j + 1
        Next anahtar
        ' Z sirasina gore (ondeki once): ZOrder 1 = en ustte
        For i = 0 To adet - 2
            For j = 0 To adet - 2 - i
                If gNesne(liste(j)).Z > gNesne(liste(j + 1)).Z Then
                    gecici = liste(j): liste(j) = liste(j + 1): liste(j + 1) = gecici
                End If
            Next j
        Next i

        For i = 0 To adet - 1
            Set s = Nothing
            Set s = gNesneRef(liste(i))
            If Not s Is Nothing Then
                On Error Resume Next
                s.Delete
                If Err.Number = 0 Then
                    nesneSil = nesneSil + 1
                Else
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next i
    End If

    ' 4) Kismi kopyalarda yalniz ilgili alt yolu sil (buyuk indeksten kucuge)
    For i = 0 To gNesneSayisi - 1
        If Not silNesne.Exists(CStr(i)) Then
            altAdet = 0
            Erase altListe
            For j = 0 To gBulguSayisi - 1
                If gBulgu(j).Tur = "TAM" And gBulgu(j).Eylem = "sil" And gBulgu(j).Nesne1 = i Then
                    ReDim Preserve altListe(0 To altAdet)
                    altListe(altAdet) = gBulgu(j).AltYol1
                    altAdet = altAdet + 1
                End If
            Next j

            If altAdet > 0 Then
                ' ayni numarayi iki kez silmeyi engelle + azalan sira
                BenzersizAzalan altListe, altAdet

                Set s = Nothing
                Set s = gNesneRef(i)
                If Not s Is Nothing Then
                    For j = 0 To altAdet - 1
                        On Error Resume Next
                        Set sp = s.Curve.SubPaths(altListe(j))
                        If Not sp Is Nothing Then
                            If s.Curve.SubPaths.Count > 1 Then
                                sp.Delete
                                If Err.Number = 0 Then altSil = altSil + 1
                                Err.Clear
                            End If
                        End If
                        Set sp = Nothing
                        On Error GoTo 0
                    Next j
                End If
            End If
        End If
    Next i

    doc.EndCommandGroup

    ' 5) Ozet
    Sil = nesneSil + altSil
    ozet = "Temizlik tamamlandi." & vbCrLf & vbCrLf & _
           "Silinen nesne: " & nesneSil & vbCrLf & _
           "Silinen alt yol: " & altSil & vbCrLf & vbCrLf & _
           "Geri almak icin Ctrl+Z (tek adim)." & vbCrLf & _
           "Isaretleri kaldirmak icin: IsaretleriKaldir makrosu."
End Function

Private Sub BenzersizAzalan(ByRef dizi() As Long, ByRef adet As Long)
    Dim i As Long, j As Long, gecici As Long
    Dim sonuc() As Long
    Dim n As Long

    ' once sirala (buyukten kucuge)
    For i = 0 To adet - 2
        For j = 0 To adet - 2 - i
            If dizi(j) < dizi(j + 1) Then
                gecici = dizi(j): dizi(j) = dizi(j + 1): dizi(j + 1) = gecici
            End If
        Next j
    Next i

    ' tekrarlari at
    ReDim sonuc(0 To adet - 1)
    n = 0
    For i = 0 To adet - 1
        If i = 0 Then
            sonuc(n) = dizi(i): n = n + 1
        ElseIf dizi(i) <> dizi(i - 1) Then
            sonuc(n) = dizi(i): n = n + 1
        End If
    Next i
    If n = 0 Then n = 1
    ReDim Preserve sonuc(0 To n - 1)

    Erase dizi
    dizi = sonuc
    adet = n
End Sub

'-[ Isaretleme ]------------------------------------------------------

Public Function Isaretle() As Long
    Dim doc As Object, page As Object, lyr As Object
    Dim i As Long, n As Long
    Dim r As Double
    Dim m As Object, t As Object
    Dim kalan As Long
    Dim bx As Double, by As Double

    On Error GoTo Hata

    Set doc = Application.ActiveDocument
    Set page = doc.ActivePage

    IsaretleriKaldirSessiz
    Set lyr = KatmanBulVeyaOlustur(page)
    If lyr Is Nothing Then
        gUyari = gUyari & "DUP_CHECK katmani olusturulamadi." & vbCrLf
        Exit Function
    End If
    lyr.Printable = False          ' isaretler plotter'a asla gitmez

    doc.BeginCommandGroup "Cift cizgi isaretleri"

    r = 1.5 / gBirimKatsayi
    For i = 0 To gBulguSayisi - 1
        If n >= 300 Then
            kalan = gBulguSayisi - i
            Exit For
        End If
        With gBulgu(i)
            bx = .X / gBirimKatsayi
            by = .Y / gBirimKatsayi
            Set m = Nothing
            On Error Resume Next
            Set m = lyr.CreateEllipse(bx - r, by + r, bx + r, by - r)
            If Not m Is Nothing Then
                m.Outline.Width = 0.5 / gBirimKatsayi
                TurRengiUygula m.Outline.Color, .Tur
                m.Fill.ApplyNoFill
                If Err.Number <> 0 Then Err.Clear
                Set t = Nothing
                Set t = lyr.CreateArtisticText(bx + r * 1.6, by + r, CStr(i + 1))
                If Not t Is Nothing Then
                    t.Text.FontProperties.Size = 6
                    TurRengiUygula t.Outline.Color, .Tur
                End If
            End If
            Err.Clear
            On Error GoTo Hata
        End With
        n = n + 1
    Next i

    ' Istege bagli: kopya olarak silinecek nesneleri kirmiziya boya
    If gBoyama Then
        For i = 0 To gBulguSayisi - 1
            With gBulgu(i)
                If .Tur = "TAM" And .Eylem = "sil" Then
                    On Error Resume Next
                    TurRengiUygula gNesneRef(.Nesne1).Outline.Color, "TAM"
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo Hata
                End If
            End With
        Next i
    End If

    doc.EndCommandGroup
    Isaretle = n
    If kalan > 0 Then
        gUyari = gUyari & "Yalniz ilk 300 bulgu isaretlendi; " & kalan & " bulgu yalniz raporda." & vbCrLf
    End If
    Exit Function

Hata:
    On Error Resume Next
    doc.EndCommandGroup
    gUyari = gUyari & "Isaretleme hatasi: " & Err.Description & vbCrLf
    Err.Clear
End Function

Private Sub TurRengiUygula(ByVal c As Object, ByVal tur As String)
    On Error Resume Next
    Select Case tur
        Case "TAM": c.RGBAssign 255, 0, 0            ' kirmizi
        Case "KISMI": c.RGBAssign 255, 128, 0        ' turuncu
        Case "YAKIN": c.RGBAssign 0, 112, 255        ' mavi
        Case Else: c.RGBAssign 128, 0, 160           ' mor
    End Select
    Err.Clear
End Sub

Public Function KatmanBulVeyaOlustur(ByVal page As Object) As Object
    Dim lyr As Object
    Set lyr = KatmanBul(page)
    If lyr Is Nothing Then
        On Error Resume Next
        Set lyr = page.CreateLayer(ISARET_KATMAN)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
    End If
    On Error Resume Next
    lyr.Visible = True
    lyr.Editable = True
    lyr.Printable = False
    Err.Clear
    Set KatmanBulVeyaOlustur = lyr
End Function

Private Function KatmanBul(ByVal page As Object) As Object
    Dim lyr As Object
    On Error Resume Next
    For Each lyr In page.Layers
        If StrComp(lyr.Name, ISARET_KATMAN, vbTextCompare) = 0 Then
            Set KatmanBul = lyr
            Exit Function
        End If
    Next lyr
    Err.Clear
    Set KatmanBul = Nothing
End Function

Public Sub IsaretleriKaldir()
    If IsaretleriKaldirSessiz() Then
        MsgBox "Isaret katmani kaldirildi.", vbInformation, UYGULAMA_ADI
    Else
        MsgBox "Isaret katmani (" & ISARET_KATMAN & ") bulunamadi; belge temiz.", _
               vbInformation, UYGULAMA_ADI
    End If
End Sub

Public Function IsaretleriKaldirSessiz() As Boolean
    Dim doc As Object, lyr As Object
    On Error GoTo Hata
    Set doc = Application.ActiveDocument
    Set lyr = KatmanBul(doc.ActivePage)
    If lyr Is Nothing Then Exit Function

    doc.BeginCommandGroup "Isaretleri kaldir"
    On Error Resume Next
    lyr.Delete
    If Err.Number <> 0 Then
        ' kilitli olabilir: once ac, sonra sil
        Err.Clear
        lyr.Editable = True
        lyr.Delete
        Err.Clear
    End If
    doc.EndCommandGroup
    IsaretleriKaldirSessiz = True
    Exit Function
Hata:
    Err.Clear
End Function

'-[ Ayarlar ]---------------------------------------------------------

Public Sub Ayarlar()
    Dim s As String
    Dim v As String

    AyarlariYukle

    s = "Mevcut ayarlar:" & vbCrLf & vbCrLf & _
        "1) Birebir tolerans (mm) : " & Sayi(gTolerans, 3) & vbCrLf & _
        "2) Yakin kopya toleransi (mm) : " & Sayi(gYakinTol, 3) & vbCrLf & _
        "3) En kisa ortusme (mm) : " & Sayi(gMinOrtusme, 2) & vbCrLf & _
        "4) Kapsam : " & gKapsam & "  (sayfa | secim)" & vbCrLf & _
        "5) Otomatik silme : " & IIf(gOtomatikSil, "acik", "kapali") & vbCrLf & _
        "6) Izgara hucre (mm) : " & Sayi(gHucreMm, 1) & vbCrLf & _
        "7) Yok sayilan katmanlar : " & IIf(Len(gYoksayKatmanlar) = 0, "(yok)", gYoksayKatmanlar) & vbCrLf & _
        "8) CSV ayirici / ondalik : " & gCsvAyirici & " / " & gOndalik & vbCrLf & _
        "9) Bulunanlari kirmiziya boya : " & IIf(gBoyama, "acik", "kapali") & vbCrLf & _
        "10) Kuru calisma (hic silme) : " & IIf(gKuruCalisma, "acik", "kapali") & vbCrLf & vbCrLf & _
        "Degistirmek istediginiz numarayi girin (bos = cikis):"

    v = InputBox(s, UYGULAMA_ADI & " - Ayarlar")
    If Len(Trim$(v)) = 0 Then Exit Sub

    Select Case Val(v)
        Case 1
            v = InputBox("Birebir tolerans (mm). Oneri: 0,05", UYGULAMA_ADI, Sayi(gTolerans, 3))
            If SayiGecerli(v, 0.0001, 5) Then gTolerans = Cozumle(v)
        Case 2
            v = InputBox("Yakin kopya toleransi (mm). Oneri: 0,5", UYGULAMA_ADI, Sayi(gYakinTol, 3))
            If SayiGecerli(v, 0.001, 10) Then
                If Cozumle(v) > gTolerans Then gYakinTol = Cozumle(v)
            End If
        Case 3
            v = InputBox("En kisa ortusme (mm). Oneri: 0,5", UYGULAMA_ADI, Sayi(gMinOrtusme, 2))
            If SayiGecerli(v, 0, 100) Then gMinOrtusme = Cozumle(v)
        Case 4
            v = InputBox("Kapsam: sayfa veya secim", UYGULAMA_ADI, gKapsam)
            If StrComp(Trim$(v), "secim", vbTextCompare) = 0 Then
                gKapsam = "secim"
            Else
                gKapsam = "sayfa"
            End If
        Case 5
            gOtomatikSil = Not gOtomatikSil
        Case 6
            v = InputBox("Izgara hucre boyu (mm). Oneri: 2", UYGULAMA_ADI, Sayi(gHucreMm, 1))
            If SayiGecerli(v, 0.5, 20) Then gHucreMm = Cozumle(v)
        Case 7
            v = InputBox("Yok sayilacak katman adlari (virgulle). Bos = yok", UYGULAMA_ADI, gYoksayKatmanlar)
            gYoksayKatmanlar = v
        Case 8
            v = InputBox("CSV ayirici (; veya ,) / ondalik (, veya .) - ornek: ;,", UYGULAMA_ADI, gCsvAyirici & gOndalik)
            If Len(v) >= 2 Then
                gCsvAyirici = Left$(v, 1)
                gOndalik = Mid$(v, 2, 1)
            End If
        Case 9
            gBoyama = Not gBoyama
        Case 10
            gKuruCalisma = Not gKuruCalisma
        Case Else
            Exit Sub
    End Select

    AyarlariKaydet
    MsgBox "Ayarlar kaydedildi." & vbCrLf & IniYolu(), vbInformation, UYGULAMA_ADI
End Sub

Private Function Cozumle(ByVal v As String) As Double
    ' Val daima "." ondaligini kabul eder; Turkce virgul once cevrilir.
    Cozumle = Val(Replace(Trim$(v), ",", "."))
End Function

Private Function SayiGecerli(ByVal v As String, ByVal enAz As Double, ByVal enCok As Double) As Boolean
    Dim d As Double
    If Len(Trim$(v)) = 0 Then Exit Function
    d = Cozumle(v)
    SayiGecerli = (d >= enAz And d <= enCok)
End Function

'-[ Yardim ]----------------------------------------------------------

Public Sub Yardim()
    Dim s As String
    s = "CIFT CIZGI DENETLEYICISI " & SURUM & vbCrLf & _
        "CorelDRAW 2024 / 2025 / 2026" & vbCrLf & vbCrLf & _
        "NE YAPAR" & vbCrLf & _
        "Plotter kesim dosyalarinda ust uste binen hatlari bulur:" & vbCrLf & _
        "  TAM    : birebir kopya (kopyala-yapistir, ice aktarma artigi)" & vbCrLf & _
        "  KISMI  : iki hattin bir bolumu ortak (bitisik parca kenari)" & vbCrLf & _
        "  YAKIN  : tolerans kadar kaymis kopya (cift kesim riski)" & vbCrLf & _
        "  NOKTAS : sifir uzunluklu segment / ust uste dugum" & vbCrLf & vbCrLf & _
        "KULLANIM" & vbCrLf & _
        "1) Dosyayi acin. 2) Denetle makrosunu calistirin." & vbCrLf & _
        "3) Rapor: %TEMP%\cdrcift klasorunde TXT + CSV." & vbCrLf & _
        "4) Bulgular DUP_CHECK katmaninda numarali olarak isaretlenir." & vbCrLf & _
        "5) Isterseniz TAM kopyalar tek komutla silinir (Ctrl+Z ile geri alinir)." & vbCrLf & vbCrLf & _
        "GUVENLIK" & vbCrLf & _
        "- KISMI ve YAKIN bulgular ASLA otomatik silinmez." & vbCrLf & _
        "- TAM silme yalniz ayni katman + ayni renk + ayni kalinlikta yapilir." & vbCrLf & _
        "- DUP_CHECK katmani yazdirilamaz (plotter'a gitmez)." & vbCrLf & vbCrLf & _
        "AYARLAR: Ayarlar makrosu. Dosya: " & IniYolu()
    MsgBox s, vbInformation, UYGULAMA_ADI
End Sub

'-[ Yardimcilar ]-----------------------------------------------------

Private Function BelgeVarMi(ByRef hata As String) As Boolean
    On Error Resume Next
    If Application.Documents.Count = 0 Then
        hata = "Acik belge yok. Once bir CorelDRAW dosyasi acin."
        Exit Function
    End If
    If Application.ActiveDocument Is Nothing Then
        hata = "Aktif belge yok."
        Exit Function
    End If
    Err.Clear
    BelgeVarMi = True
End Function

Private Function OzetMetni(ByVal raporTxt As String, ByVal raporCsv As String) As String
    Dim tam As Long, kismi As Long, yakin As Long, nokta As Long
    Dim i As Long
    For i = 0 To gBulguSayisi - 1
        Select Case gBulgu(i).Tur
            Case "TAM": tam = tam + 1
            Case "KISMI": kismi = kismi + 1
            Case "YAKIN": yakin = yakin + 1
            Case "NOKTAS": nokta = nokta + 1
        End Select
    Next i

    OzetMetni = "Sonuc: " & gBulguSayisi & " bulgu" & vbCrLf & _
                "  TAM (birebir) : " & tam & vbCrLf & _
                "  KISMI ortak   : " & kismi & vbCrLf & _
                "  YAKIN kopya   : " & yakin & vbCrLf & _
                "  NOKTAS        : " & nokta & vbCrLf & vbCrLf & _
                "Taranan: " & gNesneSayisi & " nesne / " & gAltYolSayisi & " alt yol / " & gSegSayisi & " segment" & vbCrLf & _
                "Rapor: " & raporTxt & vbCrLf & _
                "       " & raporCsv & vbCrLf & _
                IIf(Len(gUyari) > 0, vbCrLf & "UYARI:" & vbCrLf & Kisalt(gUyari, 400), "")
End Function

Private Function SilinebilirOzet() As String
    Dim i As Long, n As Long
    For i = 0 To gBulguSayisi - 1
        If gBulgu(i).Tur = "TAM" And gBulgu(i).Eylem = "sil" Then n = n + 1
    Next i
    If n = 0 Then
        SilinebilirOzet = "Otomatik silinebilir birebir kopya yok."
    Else
        SilinebilirOzet = "Otomatik silinebilir birebir kopya: " & n & " adet."
    End If
End Function
