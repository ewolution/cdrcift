Attribute VB_Name = "CiftCizgiRapor"
Option Explicit

' =====================================================================
'  cdrcift - Rapor ve ayar dosyasi modulu
'  Surum: 1.0.0
' =====================================================================

Private Const INI_BOLUM As String = "cdrcift"

'-[ Yollar ]----------------------------------------------------------

Public Function IniYolu() As String
    IniYolu = OrtamKlasoru("APPDATA") & "\cdrcift\CiftCizgi.ini"
End Function

Public Function RaporKlasoru() As String
    RaporKlasoru = OrtamKlasoru("TEMP") & "\cdrcift"
End Function

Private Function OrtamKlasoru(ByVal ad As String) As String
    Dim s As String
    s = Environ$(ad)
    If Len(s) = 0 Then s = ProfilYolu()
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    OrtamKlasoru = s
End Function

Private Function ProfilYolu() As String
    ' Environ bos ise (nadir) kullanici profilini turet
    Dim s As String
    s = Environ$("USERPROFILE")
    If Len(s) = 0 Then s = "C:\Temp"
    ProfilYolu = s
End Function

Public Function KlasorOlustur(ByVal yol As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(yol) Then
        KlasorOlustur = True
        Exit Function
    End If
    fso.CreateFolder yol
    KlasorOlustur = (Err.Number = 0)
    Err.Clear
End Function

'-[ Ayar dosyasi ]----------------------------------------------------

Public Sub AyarlariYukle()
    Dim f As Integer, satir As String
    Dim p As Long, anahtar As String, deger As String
    Dim yol As String

    VarsayilanAyarlar

    yol = IniYolu()
    On Error Resume Next
    If Dir$(yol) = "" Then Exit Sub

    f = FreeFile
    Open yol For Input As #f
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    Do While Not EOF(f)
        Line Input #f, satir
        satir = Trim$(satir)
        If Len(satir) > 0 Then
            If Left$(satir, 1) <> "[" And Left$(satir, 1) <> ";" Then
                p = InStr(satir, "=")
                If p > 1 Then
                    anahtar = LCase$(Trim$(Left$(satir, p - 1)))
                    deger = Trim$(Mid$(satir, p + 1))
                    AyarUygula anahtar, deger
                End If
            End If
        End If
    Loop
    Close #f
    Err.Clear

    If gYakinTol <= gTolerans Then gYakinTol = gTolerans * 10
    If gHucreMm < 0.5 Then gHucreMm = 0.5

    ' Kuru calisma: hicbir sey silinmez (sadece rapor + isaret)
    If gKuruCalisma Then gOtomatikSil = False
    On Error GoTo 0
End Sub

Private Sub AyarUygula(ByVal anahtar As String, ByVal deger As String)
    Dim d As Double
    On Error Resume Next

    Select Case anahtar
        Case "tolerans"
            d = SayiCoz(deger)
            If d > 0 And d <= 5 Then gTolerans = d
        Case "yakintolerans"
            d = SayiCoz(deger)
            If d > 0 And d <= 10 Then gYakinTol = d
        Case "minortusme"
            d = SayiCoz(deger)
            If d >= 0 And d <= 100 Then gMinOrtusme = d
        Case "kapsam"
            If StrComp(deger, "secim", vbTextCompare) = 0 Then gKapsam = "secim" Else gKapsam = "sayfa"
        Case "otomatiksil"
            gOtomatikSil = (deger = "1")
        Case "kurucalisma"
            gKuruCalisma = (deger = "1")
        Case "yoksaykatmanlar"
            gYoksayKatmanlar = deger
        Case "csvayirici"
            If deger = "," Or deger = ";" Then gCsvAyirici = deger
        Case "ondalik"
            If deger = "," Or deger = "." Then gOndalik = deger
        Case "hucremm"
            d = SayiCoz(deger)
            If d >= 0.5 And d <= 20 Then gHucreMm = d
        Case "maxbulgu"
            d = SayiCoz(deger)
            If d >= 100 And d <= 100000 Then gMaxBulgu = CLng(d)
        Case "boyama"
            gBoyama = (deger = "1")
    End Select
    Err.Clear
End Sub

Private Function SayiCoz(ByVal s As String) As Double
    ' Val daima "." ondaligini kabul eder (yerel ayardan bagimsiz).
    SayiCoz = Val(Replace(Trim$(s), ",", "."))
End Function

Public Sub AyarlariKaydet()
    Dim f As Integer
    Dim yol As String

    yol = IniYolu()
    If Not KlasorOlustur(OrtamKlasoru("APPDATA") & "\cdrcift") Then Exit Sub

    On Error GoTo Hata
    f = FreeFile
    Open yol For Output As #f
    Print #f, "[" & INI_BOLUM & "]"
    Print #f, "; cdrcift " & SURUM & " - ayarlar"
    Print #f, "; tolerans/yakinTolerans/minOrtusme/hucreMm: mm"
    Print #f, "tolerans=" & NoktaliSayi(gTolerans)
    Print #f, "yakinTolerans=" & NoktaliSayi(gYakinTol)
    Print #f, "minOrtusme=" & NoktaliSayi(gMinOrtusme)
    Print #f, "kapsam=" & gKapsam
    Print #f, "otomatikSil=" & IIf(gOtomatikSil, "1", "0")
    Print #f, "kuruCalisma=" & IIf(gKuruCalisma, "1", "0")
    Print #f, "yoksayKatmanlar=" & gYoksayKatmanlar
    Print #f, "csvAyirici=" & gCsvAyirici
    Print #f, "ondalik=" & gOndalik
    Print #f, "hucreMm=" & NoktaliSayi(gHucreMm)
    Print #f, "maxBulgu=" & gMaxBulgu
    Print #f, "boyama=" & IIf(gBoyama, "1", "0")
    Close #f
    Exit Sub
Hata:
    On Error Resume Next
    Close #f
    Err.Clear
End Sub

Private Function NoktaliSayi(ByVal v As Double) As String
    NoktaliSayi = Replace(CStr(v), ",", ".")
End Function

'-[ Rapor ]-----------------------------------------------------------

Public Function RaporYaz(ByRef txt As String, ByRef csv As String, ByRef hata As String) As Boolean
    Dim klasor As String
    Dim zaman As String
    Dim f As Integer
    Dim i As Long
    Dim tam As Long, kismi As Long, yakin As Long, nokta As Long
    Dim eylemSil As Long

    hata = ""
    klasor = RaporKlasoru()
    If Not KlasorOlustur(klasor) Then
        hata = "Klasor olusturulamadi: " & klasor
        Exit Function
    End If

    zaman = Format$(Now, "yyyymmdd-hhnnss")
    txt = klasor & "\rapor-" & zaman & ".txt"
    csv = klasor & "\rapor-" & zaman & ".csv"

    For i = 0 To gBulguSayisi - 1
        Select Case gBulgu(i).Tur
            Case "TAM": tam = tam + 1
            Case "KISMI": kismi = kismi + 1
            Case "YAKIN": yakin = yakin + 1
            Case "NOKTAS": nokta = nokta + 1
        End Select
        If gBulgu(i).Tur = "TAM" And gBulgu(i).Eylem = "sil" Then eylemSil = eylemSil + 1
    Next i

    On Error GoTo Hata

    f = FreeFile
    Open txt For Output As #f
    Print #f, "CIFT CIZGI DENETLEYICISI - RAPOR"
    Print #f, "Surum      : " & SURUM
    Print #f, "Tarih      : " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "Belge      : " & BelgeAdi()
    Print #f, "Belge birimi: " & gBirimAdi & " (1 birim = " & Sayi(gBirimKatsayi, 4) & " mm)"
    Print #f, "Kapsam     : " & gKapsam
    Print #f, "Tolerans   : " & Sayi(gTolerans, 3) & " mm"
    Print #f, "Yakin tol. : " & Sayi(gYakinTol, 3) & " mm"
    Print #f, "Min ortusme: " & Sayi(gMinOrtusme, 2) & " mm"
    Print #f, "Otomatik silme: " & IIf(gOtomatikSil, "acik", "kapali")
    Print #f, ""
    Print #f, "OZET"
    Print #f, "  TAM (birebir)      : " & tam
    Print #f, "  KISMI (ortak kenar): " & kismi
    Print #f, "  YAKIN (kaymis)     : " & yakin
    Print #f, "  NOKTAS             : " & nokta
    Print #f, "  TOPLAM             : " & gBulguSayisi
    Print #f, "  Otomatik silinecek : " & eylemSil
    Print #f, ""
    Print #f, "TARANAN"
    Print #f, "  Nesne    : " & gNesneSayisi
    Print #f, "  Alt yol  : " & gAltYolSayisi
    Print #f, "  Segment  : " & gSegSayisi
    Print #f, "  Atlanan nesne: " & gAtlanan & "   Efekt: " & gEfektAtlanan & _
                 "   PowerClip: " & gPowerClipAtlanan
    If gBulguKesildi Then
        Print #f, "  NOT: Bulgu listesi " & gMaxBulgu & " kayitla sinirlandi."
    End If
    If Len(gUyari) > 0 Then
        Print #f, ""
        Print #f, "UYARILAR"
        Print #f, gUyari
    End If
    Print #f, ""
    Print #f, "BULGULAR"
    Print #f, "No | Tur | Katman | Nesne | AltYol | Uzunluk mm | Oran % | Konum X,Y | Eylem | Eslesen | Not"
    Print #f, String$(120, "-")

    For i = 0 To gBulguSayisi - 1
        Print #f, SatirMetni(i, False)
    Next i

    Close #f

    f = FreeFile
    Open csv For Output As #f
    Print #f, CsvSatir(Array("No", "Tur", "Katman", "Nesne", "AltYol", "Uzunluk_mm", "Oran_yuzde", _
                             "X", "Y", "Eylem", "EslesenNesne", "EslesenAltYol", "Not"))
    For i = 0 To gBulguSayisi - 1
        Print #f, CsvSatir(Array(CStr(i + 1), gBulgu(i).Tur, gNesne(gBulgu(i).Nesne1).Katman, _
                                 gNesne(gBulgu(i).Nesne1).Ad, CStr(gBulgu(i).AltYol1), _
                                 Sayi(gBulgu(i).Uzunluk, 2), Sayi(gBulgu(i).Oran * 100, 0), _
                                 Sayi(gBulgu(i).X, 2), Sayi(gBulgu(i).Y, 2), gBulgu(i).Eylem, _
                                 EslesenAd(i), EslesenAltYol(i), gBulgu(i).Aciklama))
    Next i
    Close #f

    RaporYaz = True
    Exit Function

Hata:
    On Error Resume Next
    Close #f
    hata = "Dosya yazma hatasi (" & Err.Number & "): " & Err.Description
    Err.Clear
End Function

Private Function SatirMetni(ByVal i As Long) As String
    With gBulgu(i)
        SatirMetni = Right$("     " & CStr(i + 1), 5) & " | " & _
                     Left$(.Tur & "     ", 6) & " | " & _
                     Left$(gNesne(.Nesne1).Katman & "          ", 10) & " | " & _
                     Left$(gNesne(.Nesne1).Ad & "                    ", 20) & " | " & _
                     Right$("   " & CStr(.AltYol1), 3) & " | " & _
                     Right$("        " & Sayi(.Uzunluk, 2), 8) & " | " & _
                     Right$("    " & Sayi(.Oran * 100, 0), 4) & " | " & _
                     Right$("        " & Sayi(.X, 2), 8) & "," & Right$("        " & Sayi(.Y, 2), 8) & " | " & _
                     Left$(.Eylem & "    ", 4) & " | " & _
                     Left$(EslesenAd(i) & "          ", 10) & " | " & _
                     .Aciklama
    End With
End Function

Private Function EslesenAd(ByVal i As Long) As String
    ' IIf her iki dali da degerlendirdigi icin gNesne(-1) hatasini onler
    If gBulgu(i).Nesne2 >= 0 Then
        EslesenAd = gNesne(gBulgu(i).Nesne2).Ad
    Else
        EslesenAd = "-"
    End If
End Function

Private Function EslesenAltYol(ByVal i As Long) As String
    If gBulgu(i).Nesne2 >= 0 Then
        EslesenAltYol = CStr(gBulgu(i).AltYol2)
    Else
        EslesenAltYol = ""
    End If
End Function

Private Function CsvSatir(ByVal alan As Variant) As String
    Dim i As Long
    Dim s As String, v As String
    For i = LBound(alan) To UBound(alan)
        v = CStr(alan(i))
        If InStr(v, gCsvAyirici) > 0 Or InStr(v, """") > 0 Or InStr(v, vbCr) > 0 Or InStr(v, vbLf) > 0 Then
            v = """" & Replace(v, """", """""") & """"
        End If
        If i > 0 Then s = s & gCsvAyirici
        s = s & v
    Next i
    CsvSatir = s
End Function

Private Function BelgeAdi() As String
    Dim ad As String
    On Error Resume Next
    ad = Application.ActiveDocument.FileName
    If Err.Number <> 0 Or Len(ad) = 0 Then
        Err.Clear
        ad = "(kaydedilmemis belge)"
    End If
    BelgeAdi = ad
End Function
