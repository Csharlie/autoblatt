Attribute VB_Name = "modBeallitas"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.21 - 2026.10.15.
'   9.21: "Mentes Drive-ra" gomb (K3, fejlesztes alatt, mint az e-mail gomb):
'         Mentes maskent kozvetlenul a honap mappajaba: OneDrive\Autoblatt\<Ev>\<Honap>\
'         <datum> <tipus> <megnevezes> <preseles> <hibakep> (<id>-foto).xlsm,
'         ekezet nelkul; az id a Data!User tablabol (A2), kisbetuvel.
'         Utkozeskor v2, v3... Az e-mail targya: Cikkszam | Tipus | Megnevezes |
'         Preseles | Hibakep | Gyarto berendezes | Fotoblatt (ekezetekkel, ahogy a cellakban).
'         A level szovege: "Sziasztok! / Mellekelten kuldok egy fotolapot!"
'         Mentes utan nincs kulon Intezo-ablak (MENTES_MAPPA_NYITAS = True: van).
'         A fejlesztoi gombok lathatosaga egy helyen (GombIgazitas).
'   9.20: verzio-igazitas (uj mappaszerkezet, l. modTermekadatok), kodvaltozas nincs.
'   9.19: E-mail gomb a Fotoblatt lapon (K1, a nyomtatasi teruleten kivul):
'         uj Outlook-level, targy: Cikkszam | Megnevezes | Tipus | Preseles |
'         Hibakep | Gyarto berendezes, csatolmany: a munkafuzet masolata
'         eredeti neven (csak megjeleniti, nem kuldi el). FEJLESZTES ALATT:
'         csak akkor lathato es mukodik, ha az A2 es a Windows-felhasznalo
'         is EMAIL_FEJLESZTO; kiadas: EMAIL_MINDENKINEK = True.
'   9.18: verzio-igazitas a modFotolap 9.18-hoz, ebben a modulban nincs
'         kodvaltozas.
'   9.17: lablec "Verzio" cella = FOTOBLATT_VERZIO (LablecVerzio, a beallitas
'         es az ujraigazitas is irja). Egyszeri takaritas (RegiMaradvanyokTorlese):
'         a regi rejtett A6ListaMod nev torlese. Hibakep-legordulo a
'         Hibafeltaras cimsorban (Data!tblHibakep -> HibakepLista), a vegen
'         a hibamegnevezesek beolvasasat is felajanlja (HibakepekFrissitese).
'         A horgony-ellenorzes a H_CIM / H_FEJLEC allandokat hasznalja.
'   9.16: a beallitas vegen a TERMEKADATOK utan a cikkszamtorzs (A/N-szamok)
'         beolvasasat is felajanlja (CikkszamtorzsFrissitese).
'
' A kod Sardy Peter munkaja. Modositas elott egyeztess vele.

Option Explicit

'--- E-mail gomb (v9.19) --------------------------------------------------
' Fejlesztes alatt csak EMAIL_FEJLESZTO latja es hasznalhatja (A2 ES a
' Windows-felhasznalo is). Kiadaskor: EMAIL_MINDENKINEK = True.
Private Const EMAIL_MINDENKINEK As Boolean = False
Private Const EMAIL_FEJLESZTO   As String = "PESARDY"
Private Const BTN_EMAIL         As String = "btnEmail"
Private Const EMAIL_CIMZETT     As String = ""        ' ures = kezzel adod meg

'--- Mentes Drive-ra gomb (v9.21) ---------------------------------------------
' Fejlesztes alatt ugyanugy csak a fejleszto latja (EMAIL_FEJLESZTO), az
' e-mail gombtol fuggetlenul adhato ki: MENTES_MINDENKINEK = True.
Private Const MENTES_MINDENKINEK As Boolean = False
Private Const BTN_MENTES         As String = "btnMentesDrive"
Private Const MENTES_GYOKER      As String = "Autoblatt"   ' a OneDrive mappa alatt
Private Const MENTES_ESETMAPPA   As Boolean = False        ' True = minden fotolap sajat (fajlnev szerinti) mappaba
Private Const MENTES_MAPPA_NYITAS As Boolean = False        ' True = mentes utan Intezo-ablak a mappaval

'==========================================================================
' Fotolap - egyszer futtatando beallitas
'--------------------------------------------------------------------------
' Alt+F8 -> SetupFotolap -> Futtatas.
' Letrehozza a hianyzo neveket es tablat, helyreallitja a legordoloket,
' es egyszer vegigfuttatja a szerkezeti normalizalast.
'
' FONTOS: ez a fajl szandekosan CSAK ASCII karaktereket tartalmaz, mert a
' VBE a .bas fajlokat a rendszer kodlapjaval olvassa be, nem UTF-8-kent.
'==========================================================================

Public Sub SetupFotolap()

    Dim ws As Worksheet, wsData As Worksheet
    Dim uzenet As String

    On Error GoTo Hiba

    Set ws = ThisWorkbook.Worksheets(SH_FOTO)
    Set wsData = ThisWorkbook.Worksheets(SH_DATA)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--- 1. igen/nem tabla a Data lapon, ha meg nincs ---------------------
    If Not TablaLetezik(wsData, "IgenNem") Then
        With wsData
            .Range("O1").Value = "IgenNem"
            .Range("O2").Value = "igen / ja"
            .Range("O3").Value = "nem / nein"
            .ListObjects.Add(xlSrcRange, .Range("O1:O3"), , xlYes).Name = "IgenNem"
            .ListObjects("IgenNem").ListColumns(1).Name = "IgenNem"
        End With
        uzenet = uzenet & "- IgenNem tabla letrehozva (Data!O1:O3)" & vbLf
    End If

    '--- 2. Definialt nevek ----------------------------------------------
    NevBeallit "UserIDLista", "=User[ID]"
    NevBeallit "ProdLista", "=Prod[Production]"
    NevBeallit "CikkszamLista", "=Article[PartNum]"
    NevBeallit "TipusLista", "=Type[Type]"
    NevBeallit "IgenNemLista", "=IgenNem[IgenNem]"
    NevBeallit "ErtekelesLista", "=Table13[Column1]"
    NevBeallit "PrioLista", "=Prio[Prio]"
    NevBeallit "OPLista", "=OP[OP]"
    uzenet = uzenet & "- 8 definialt nev beallitva" & vbLf

    '--- 2b. v9.2: Article oszlopok, tipuskodok, A6 lista ---------------
    ArticleOszlopokBiztositasa
    uzenet = uzenet & "- Article oszlopok (Type, Forras, Statusz, ...) rendben" & vbLf
    uzenet = uzenet & TipusKodok92()
    A6ListaFrissites ws
    uzenet = uzenet & "- A6Lista nev beallitva (Parts!J)" & vbLf

    '--- 3. Fejlec legordulok --------------------------------------------
    LegorduloBeallit ws.Range("A2"), "=UserIDLista"        ' felhasznalo
    LegorduloBeallit ws.Range("A6:B7"), "=" & NEV_A6LISTA  ' cikkszam (v9.2: szurt lista)
    LegorduloBeallit ws.Range("G6:G7"), "=TipusLista"      ' tipus
    LegorduloBeallit ws.Range("H6:I7"), "=ProdLista"       ' gyartoberendezes
    uzenet = uzenet & "- A2 / A6 / G6 / H6 legordulo helyreallitva" & vbLf

    '--- 3b. v9.17: hibakep-legordulo a Hibafeltaras cimsorban ------------
    HibakepTablaBiztositasa
    If HibakepLegorduloBeallitas(ws) Then
        uzenet = uzenet & "- Hibakep legordulo beallitva (Data!tblHibakep)" & vbLf
    Else
        uzenet = uzenet & "- Hibakep legordulo: a cimsorban nincs 'Hibakep' felirat, kihagyva" & vbLf
    End If

    '--- 4. Datummezo ----------------------------------------------------
    On Error Resume Next
    ws.Range("I3").Validation.Delete       ' regi, #REF!-re mutato ervenyesites
    On Error GoTo Hiba
    ws.Range("I3").NumberFormat = "yyyy.mm.dd;@"
    uzenet = uzenet & "- I3 datummezo rendbe teve" & vbLf

    '--- 5. Oldalbeallitas -----------------------------------------------
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
    End With
    On Error GoTo Hiba
    ElerhetoMagassagUjraszamol

    '--- 5b. v9.17: regi maradvanyok (A6ListaMod) -------------------------
    uzenet = uzenet & RegiMaradvanyokTorlese(ws)

    '--- 6. Szerkezeti normalizalas + A4 igazitas -------------------------
    HibafeltarasKitoltes ws
    uzenet = uzenet & "- Hibafeltaras blokk es kepterulet igazitva" & vbLf

    '--- 7. v9.17: lablec verzio -----------------------------------------
    If LablecVerzio(ws) Then
        uzenet = uzenet & "- Lablec verzio: " & FOTOBLATT_VERZIO & vbLf
    Else
        uzenet = uzenet & "- Lablec: 'Verzio' cella nem talalhato, kihagyva" & vbLf
    End If

    '--- 8. v9.19: e-mail gomb (fejlesztes alatt) ------------------------
    EmailGombLathatosag ws
    If EmailJogosult(ws) Then
        uzenet = uzenet & "- E-mail gomb: lathato (K1)" & vbLf
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "A Fotolap beallitasa kesz:" & vbLf & vbLf & uzenet, _
           vbInformation, "Fotolap beallitas"

    If MsgBox("Beolvassam most a TERMEKADATOK fajlt (R-szamok)?", _
              vbQuestion + vbYesNo, "Fotolap beallitas") = vbYes Then
        TermekadatokFrissitese
    End If

    ' v9.16: A- es N-szamok a cikkszamtorzsbol
    If MsgBox("Beolvassam most a cikksz" & ChrW(225) & "mt" & ChrW(246) & "rzset (A/N-sz" & _
              ChrW(225) & "mok)?", vbQuestion + vbYesNo, "Fotolap beallitas") = vbYes Then
        CikkszamtorzsFrissitese
    End If

    ' v9.17: hibamegnevezesek (magyar nevek) a hibakep-legordulohoz
    If MsgBox("Beolvassam most a hibamegnevez" & ChrW(233) & "seket (Hibamegnevez" & _
              ChrW(233) & "sek.xls)?", vbQuestion + vbYesNo, "Fotolap beallitas") = vbYes Then
        HibakepekFrissitese
    End If
    Exit Sub

Hiba:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "A beallitas megszakadt: " & Err.Description, vbExclamation, "Fotolap beallitas"

End Sub


'==========================================================================
' Segedeljarasok
'==========================================================================
Private Sub NevBeallit(ByVal nev As String, ByVal hivatkozas As String)

    On Error Resume Next
    ThisWorkbook.Names(nev).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nev, RefersTo:=hivatkozas

End Sub


Private Sub LegorduloBeallit(cel As Range, ByVal keplet As String)

    With cel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=keplet
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False              ' listan kivuli ertek is beirhato
    End With

End Sub


'--------------------------------------------------------------------------
' v9.17: a lablec "Verzio" cellaja = "Verzio: " & FOTOBLATT_VERZIO.
' A "Keszitette" (H_VEG) sorban, a B:I oszlopokban keresi azt a cellat,
' amelynek szovege (ekezet nelkul) "Verzio"-val kezdodik (most: D59).
' Ha a sor vagy a cella nem talalhato, csendben kihagyja (hibat nem dob).
' True, ha beirta (a SetupFotolap osszegzesehez).
'--------------------------------------------------------------------------
Private Function LablecVerzio(ws As Worksheet) As Boolean

    Dim sor As Long, c As Long

    On Error GoTo Vege
    sor = SorKereses(ws, H_VEG, 2, 200)
    If sor = 0 Then Exit Function

    For c = 2 To 9
        If LCase$(Left$(Ekezetlen(Trim$(CStr(ws.Cells(sor, c).Value))), 6)) = "verzio" Then
            ws.Cells(sor, c).Value = "Verzio: " & FOTOBLATT_VERZIO
            LablecVerzio = True
            Exit Function
        End If
    Next c

Vege:
End Function


'--------------------------------------------------------------------------
' v9.17: egyszeri takaritas, a SetupFotolap hivja. Visszaad: osszegzo sorok.
' A regi, rejtett A6ListaMod nev torlese (a 9.7-9.10-es kereso maradvanya).
' (A Hibafeltaras cimsorhoz NEM nyul: az E oszlopban a "Hibakep /
' Fehlerbild" felirat, mellette a hibakep-legordulo all.)
'--------------------------------------------------------------------------
Private Function RegiMaradvanyokTorlese(ws As Worksheet) As String

    Dim nev As Name
    Dim ki As String

    On Error Resume Next
    Set nev = ThisWorkbook.Names("A6ListaMod")
    On Error GoTo 0
    If Not nev Is Nothing Then
        nev.Delete
        ki = ki & "- Regi A6ListaMod nev torolve" & vbLf
    End If

    RegiMaradvanyokTorlese = ki

End Function


Private Function TablaLetezik(ws As Worksheet, ByVal nev As String) As Boolean

    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If StrComp(lo.Name, nev, vbTextCompare) = 0 Then
            TablaLetezik = True
            Exit Function
        End If
    Next lo

End Function


'==========================================================================
' Kezi ujraigazitas, ha valaki utolag modositotta a lapot
'==========================================================================
Public Sub FotolapUjraigazitas()

    On Error GoTo Hiba
    Application.EnableEvents = False
    ElerhetoMagassagUjraszamol
    HibafeltarasKitoltes ThisWorkbook.Worksheets(SH_FOTO)
    LablecVerzio ThisWorkbook.Worksheets(SH_FOTO)          ' v9.17

Hiba:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "Az ujraigazitas megszakadt: " & Err.Description, vbExclamation
    Else
        MsgBox "A Fotolap ujraigazitva.", vbInformation
    End If

End Sub


'==========================================================================
' Diagnosztika: kiirja, hol talalja meg a szerkezeti horgonyokat
'==========================================================================
Public Sub HorgonyokEllenorzese()

    Dim ws As Worksheet
    Dim jelentes As String
    Dim horgonyok As Variant
    Dim i As Long, sor As Long

    Set ws = ThisWorkbook.Worksheets(SH_FOTO)

    horgonyok = Array(H_CIM, H_FEJLEC, "WZG IH", _
                      "Kepek/Bilder", "Intezkedesek WIH", "Keszitette")

    For i = LBound(horgonyok) To UBound(horgonyok)
        sor = SorKereses(ws, CStr(horgonyok(i)), 2, 200)
        If sor > 0 Then
            jelentes = jelentes & "OK    " & horgonyok(i) & " -> sor " & sor
            If ws.Rows(sor).Hidden Then jelentes = jelentes & "  (rejtett)"
            jelentes = jelentes & vbLf
        Else
            jelentes = jelentes & "HIANY " & horgonyok(i) & vbLf
        End If
    Next i

    MsgBox jelentes, vbInformation, "Szerkezeti horgonyok"

End Sub


'==========================================================================
' A Hibafeltaras blokk kezi megmutatasa, ha valaki rejtve ragadna
'==========================================================================
Public Sub HibafeltarasMegmutatasa()

    Dim ws As Worksheet
    Dim cimSor As Long, hdrSor As Long

    Set ws = ThisWorkbook.Worksheets(SH_FOTO)

    cimSor = SorKereses(ws, "Hibafeltaras", 2, 200)
    hdrSor = SorKereses(ws, H_FEJLEC, 2, 200)

    If cimSor = 0 Or hdrSor = 0 Or hdrSor - 1 < cimSor Then
        MsgBox "A blokk hatarai nem allapithatok meg.", vbExclamation
        Exit Sub
    End If

    ws.Rows(cimSor & ":" & CStr(hdrSor - 1)).Hidden = False
    MsgBox "A Hibafeltaras blokk (" & cimSor & "-" & (hdrSor - 1) & ". sor) lathato.", _
           vbInformation

End Sub


'==========================================================================
' v9.19: E-mail gomb - a munkafuzet uj Outlook-levelhez csatolva
'--------------------------------------------------------------------------
' Urlapgomb a Fotoblatt lapon (K1, a nyomtatasi teruleten kivul, nem
' nyomtatodik). Kattintasra a munkafuzet masolata (eredeti neven, a TEMP
' mappaban) egy uj Outlook-levelhez csatolodik, targy (v9.21): Cikkszam |
' Tipus | Megnevezes | Preseles | Hibakep | Gyarto berendezes | Fotoblatt. A level csak
' megjelenik, NEM megy el magatol. Klasszikus Outlook kell (COM).
'
' Fejlesztes alatt (EMAIL_MINDENKINEK = False) a gomb csak akkor LATHATO es
' csak akkor MUKODIK, ha az A2 es a Windows-felhasznalo is EMAIL_FEJLESZTO.
' A lapkod minden valtozaskor, lapra lepeskor es kijeloleskor igazitja a
' lathatosagot, a makro maga is ellenoriz (ha valaki megis elinditana).
'==========================================================================

' True, ha az aktualis felhasznalo latja/hasznalhatja az e-mail gombot.
' Hibat nem dob (hiba eseten False = rejtve).
Public Function EmailJogosult(ws As Worksheet) As Boolean
    EmailJogosult = FejlesztoiJogosult(ws, EMAIL_MINDENKINEK)
End Function

' v9.21: ugyanez a "Mentes Drive-ra" gombra
Public Function MentesJogosult(ws As Worksheet) As Boolean
    MentesJogosult = FejlesztoiJogosult(ws, MENTES_MINDENKINEK)
End Function

' v9.21: kozos jogosultsag-ellenorzes a fejlesztes alatti gombokhoz: kiadva
' mindenkinek, kulonben csak ha az A2 ES a Windows-felhasznalo is a fejleszto.
Private Function FejlesztoiJogosult(ws As Worksheet, ByVal mindenkinek As Boolean) As Boolean
    On Error GoTo Vege
    If mindenkinek Then
        FejlesztoiJogosult = True
    Else
        FejlesztoiJogosult = _
            (StrComp(Trim$(CStr(ws.Range("A2").Value)), EMAIL_FEJLESZTO, vbTextCompare) = 0) And _
            (StrComp(Trim$(Environ$("USERNAME")), EMAIL_FEJLESZTO, vbTextCompare) = 0)
    End If
Vege:
End Function

' A fejlesztoi gombok lathatosaganak igazitasa (a lapkod hivja; a nev a
' korabbi, v9.21 ota mindket gombot kezeli). Hibat nem dob.
Public Sub EmailGombLathatosag(ws As Worksheet)
    On Error Resume Next
    GombIgazitas ws, BTN_EMAIL, EmailJogosult(ws), _
        ChrW(9993) & "  E-mail k" & ChrW(252) & "ld" & ChrW(233) & "se...", "FotolapEmail", 1
    GombIgazitas ws, BTN_MENTES, MentesJogosult(ws), _
        ChrW(9729) & "  Ment" & ChrW(233) & "s Drive-ra...", "FotolapMentesDrive", 3
End Sub

' v9.21: egy gomb lathatosaga; jogosultnal, ha meg nincs, letrehozza (urlapgomb
' a K oszlopban, az A:I oldalon kivul, nem nyomtatodik). Csak akkor ir, ha
' valtozas kell. Hibat nem dob.
Private Sub GombIgazitas(ws As Worksheet, ByVal nev As String, ByVal kell As Boolean, _
                         ByVal felirat As String, ByVal makro As String, ByVal sor As Long)

    Dim gomb As Shape, hely As Range, b As Object

    On Error Resume Next
    Set gomb = ws.Shapes(nev)
    If gomb Is Nothing And kell Then
        Set hely = ws.Cells(sor, 11)
        Set b = ws.Buttons.Add(hely.Left + 6, hely.Top + 3, 170, 24)
        b.Name = nev
        b.Caption = felirat
        b.OnAction = makro
        b.Placement = xlFreeFloating
        b.PrintObject = False
        Set gomb = ws.Shapes(nev)
    End If
    If Not gomb Is Nothing Then
        If CBool(gomb.Visible) <> kell Then gomb.Visible = kell
    End If

End Sub

' v9.21: a fejlesztoi gombok elrejtese (pl. mentes vagy masolat elott), es
' az elozo allapot visszaadasa (vesszovel elvalasztott "nev=0/1" lista).
Private Function GombokElrejtese(ws As Worksheet) As String
    Dim nev As Variant, gomb As Shape, ki As String
    For Each nev In Array(BTN_EMAIL, BTN_MENTES)
        Set gomb = Nothing
        On Error Resume Next
        Set gomb = ws.Shapes(CStr(nev))
        On Error GoTo 0
        If Not gomb Is Nothing Then
            ki = ki & nev & "=" & IIf(CBool(gomb.Visible), "1", "0") & ","
            gomb.Visible = False
        End If
    Next nev
    GombokElrejtese = ki
End Function

Private Sub GombokVisszaallitasa(ws As Worksheet, ByVal allapot As String)
    Dim t As Variant, p As Long
    On Error Resume Next
    For Each t In Split(allapot, ",")
        p = InStr(t, "=")
        If p > 0 Then ws.Shapes(Left$(t, p - 1)).Visible = (Mid$(t, p + 1) = "1")
    Next t
End Sub

' Az e-mail gomb makroja. Az opcionalis parameter miatt NEM jelenik meg az
' Alt+F8 makrolistaban; a gomb (OnAction) igy is meghivja.
Public Sub FotolapEmail(Optional ByVal rejtett As Variant)

    Dim ws As Worksheet, ol As Object, level As Object
    Dim targy As String, masolat As String, html As String, torzs As String
    Dim p As Long, lepes As String, gombAllapot As String

    On Error GoTo Hiba
    Set ws = ThisWorkbook.Worksheets(SH_FOTO)

    lepes = "jogosultsag"
    If Not EmailJogosult(ws) Then
        MsgBox "Ez a funkci" & ChrW(243) & " m" & ChrW(233) & "g fejleszt" & ChrW(233) & _
               "s alatt " & ChrW(225) & "ll.", vbInformation, "E-mail"
        EmailGombLathatosag ws
        Exit Sub
    End If

    targy = EmailTargy(ws)

    '--- A munkafuzet masolata a TEMP mappaba, eredeti neven es formatumban.
    '    (A OneDrive-os munkafuzet utvonala URL lehet, kozvetlenul nem
    '    csatolhato; a masolat a nem mentett modositasokat is tartalmazza.)
    '    A masolat elott a fejlesztoi gombokat elrejtjuk, hogy a cimzettnel ne
    '    latszodjanak, utana visszaallitjuk.
    lepes = "munkafuzet masolasa"
    masolat = Environ$("TEMP") & "\" & ThisWorkbook.Name
    If Len(Dir$(masolat)) > 0 Then Kill masolat
    gombAllapot = GombokElrejtese(ws)
    ThisWorkbook.SaveCopyAs masolat
    GombokVisszaallitasa ws, gombAllapot
    gombAllapot = ""

    '--- Outlook-level -----------------------------------------------------
    lepes = "Outlook inditasa"
    Set ol = CreateObject("Outlook.Application")
    Set level = ol.CreateItem(0)                     ' olMailItem

    lepes = "level osszeallitasa"
    If Len(EMAIL_CIMZETT) > 0 Then level.To = EMAIL_CIMZETT
    level.Subject = targy
    level.Attachments.Add masolat
    level.Display                                    ' itt kerul bele az alairas

    ' Rovid szoveg az alairas ele (a <body> nyito tag utan)
    ' v9.21: koszontes + uj mondat
    torzs = "<p>Sziasztok!</p><p>" & HtmlSzoveg("Mell" & ChrW(233) & "kelten k" & ChrW(252) & _
            "ld" & ChrW(246) & "k egy fot" & ChrW(243) & "lapot!") & "</p>"
    html = level.HTMLBody
    p = InStr(1, html, "<body", vbTextCompare)
    If p > 0 Then p = InStr(p, html, ">")
    If p > 0 Then
        level.HTMLBody = Left$(html, p) & torzs & Mid$(html, p + 1)
    Else
        level.HTMLBody = torzs & html
    End If
    Exit Sub

Hiba:
    Dim hSzam As Long, hSzoveg As String
    hSzam = Err.Number: hSzoveg = Err.Description
    ' a gombok ne maradjanak rejtve, ha a masolas kozben tortent a hiba
    On Error Resume Next
    If Len(gombAllapot) > 0 Then GombokVisszaallitasa ws, gombAllapot
    MsgBox "Az e-mail nem k" & ChrW(233) & "sz" & ChrW(252) & "lt el [" & lepes & "]:" & vbLf & _
           hSzoveg & " (" & hSzam & ")", vbExclamation, "E-mail"
End Sub

'==========================================================================
' v9.21: Mentes Drive-ra (fejlesztes alatt)
'--------------------------------------------------------------------------
' Mentes maskent (a megnyitott munkafuzet maga lesz a mentett fotolap, az
' autoblatt.xlsm sablon a lemezen erintetlen marad):
'   OneDrive\Autoblatt\<Ev>\<Honap>\[<esetmappa>\]
'     <datum> <tipus> <megnevezes> <preseles> <hibakep> (<id>-foto).xlsm
' A datum az I3-bol (ures: mai nap), ebbol az ev es a honap is. Minden resz
' ekezet nelkul; az ures mezo kimarad. Az ID a Data!User tablabol (az A2-ben
' kivalasztott felhasznalo sora). Ha a fajl mar letezik: v2, v3... (kerdesre).
'==========================================================================
Public Sub FotolapMentesDrive(Optional ByVal rejtett As Variant)

    Dim ws As Worksheet, datum As Date, id As String, alap As String
    Dim od As String, mappa As String, ut As String, v As Long
    Dim lepes As String, gombAllapot As String, valasz As VbMsgBoxResult, CIM As String

    On Error GoTo Hiba
    CIM = "Ment" & ChrW(233) & "s Drive-ra"
    Set ws = ThisWorkbook.Worksheets(SH_FOTO)

    lepes = "jogosultsag"
    If Not MentesJogosult(ws) Then
        MsgBox "Ez a funkci" & ChrW(243) & " m" & ChrW(233) & "g fejleszt" & ChrW(233) & _
               "s alatt " & ChrW(225) & "ll.", vbInformation, CIM
        EmailGombLathatosag ws
        Exit Sub
    End If

    '--- Nev es hely ------------------------------------------------------
    lepes = "fajlnev"
    If IsDate(ws.Range("I3").Value) Then datum = CDate(ws.Range("I3").Value) Else datum = Date
    id = LCase$(FelhasznaloId(ws))               ' a fajlnevben kisbetuvel
    alap = MentesNevAlap(ws, datum)
    If Len(alap) <= 8 Then
        MsgBox "A f" & ChrW(225) & "jln" & ChrW(233) & "vhez nincs el" & ChrW(233) & "g adat: " & _
               "t" & ChrW(246) & "ltsd ki legal" & ChrW(225) & "bb a cikksz" & ChrW(225) & "mot " & _
               "(t" & ChrW(237) & "pus, megnevez" & ChrW(233) & "s) " & ChrW(233) & "s a hibak" & _
               ChrW(233) & "pet.", vbExclamation, CIM
        Exit Sub
    End If

    lepes = "OneDrive mappa"
    od = Environ$("OneDriveCommercial")
    If Len(od) = 0 Then od = Environ$("OneDrive")
    If Len(od) = 0 Then Err.Raise vbObjectError + 570, , "A OneDrive mappa nem talalhato (OneDriveCommercial)."
    mappa = od & "\" & MENTES_GYOKER & "\" & Format$(datum, "yyyy") & "\" & Format$(datum, "mm")
    If MENTES_ESETMAPPA Then mappa = mappa & "\" & alap
    ut = mappa & "\" & alap & " (" & id & "-foto).xlsm"

    If Len(Dir$(ut)) > 0 Then
        valasz = MsgBox("A f" & ChrW(225) & "jl m" & ChrW(225) & "r l" & ChrW(233) & "tezik:" & vbLf & vbLf & ut & _
                        vbLf & vbLf & "Mentsem " & ChrW(250) & "j verzi" & ChrW(243) & "k" & ChrW(233) & "nt " & _
                        "(v2, v3 ...)?", vbQuestion + vbYesNo, CIM)
        If valasz <> vbYes Then Exit Sub
        v = 1
        Do
            v = v + 1
            ut = mappa & "\" & alap & " (" & id & "-foto) v" & v & ".xlsm"
        Loop While Len(Dir$(ut)) > 0
    End If

    valasz = MsgBox("Ment" & ChrW(233) & "s ide:" & vbLf & vbLf & ut & vbLf & vbLf & _
                    "A megnyitott munkaf" & ChrW(252) & "zet ezut" & ChrW(225) & "n ez a f" & ChrW(225) & _
                    "jl lesz (az autoblatt.xlsm sablon nem v" & ChrW(225) & "ltozik).", _
                    vbQuestion + vbOKCancel, CIM)
    If valasz <> vbOK Then Exit Sub

    '--- Mentes -----------------------------------------------------------
    lepes = "mappa letrehozasa"
    MappaLetrehozas mappa

    lepes = "mentes"
    gombAllapot = GombokElrejtese(ws)          ' a mentett fajlban ne latszodjanak
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ut, FileFormat:=52      ' xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    GombokVisszaallitasa ws, gombAllapot
    gombAllapot = ""
    ThisWorkbook.Saved = True                  ' a gombok visszaallitasa nem modositas

    If MENTES_MAPPA_NYITAS Then Shell "explorer.exe """ & mappa & """", vbNormalFocus
    MsgBox "Mentve:" & vbLf & ut, vbInformation, CIM
    Exit Sub

Hiba:
    Dim hSzam As Long, hSzoveg As String
    hSzam = Err.Number: hSzoveg = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    If Len(gombAllapot) > 0 Then GombokVisszaallitasa ws, gombAllapot
    MsgBox "A ment" & ChrW(233) & "s nem siker" & ChrW(252) & "lt [" & lepes & "]:" & vbLf & _
           hSzoveg & " (" & hSzam & ")", vbExclamation, CIM
End Sub

' v9.21: "<datum> <tipus> <megnevezes> <preseles> <hibakep>" ekezet nelkul,
' az ures mezok nelkul (G6, C6, E6, hibakep-cella).
Private Function MentesNevAlap(ws As Worksheet, ByVal datum As Date) As String
    Dim reszek As Variant, resz As Variant, s As String, hibakep As String, cel As Range
    Set cel = HibakepCella(ws)
    If Not cel Is Nothing Then hibakep = CellaSzoveg(ws, cel.Cells(1, 1).Address)
    reszek = Array(CellaSzoveg(ws, "G6"), CellaSzoveg(ws, "C6"), CellaSzoveg(ws, "E6"), hibakep)
    s = Format$(datum, "yyyymmdd")
    For Each resz In reszek
        resz = FajlnevResz(CStr(resz))
        If Len(resz) > 0 Then s = s & " " & resz
    Next resz
    MentesNevAlap = s
End Function

' v9.21: a kitolto azonositoja a Data!User tablabol (az A2-ben kivalasztott
' felhasznalo sora); tartalek: az A2 erteke, vegul a Windows-felhasznalo.
Private Function FelhasznaloId(ws As Worksheet) As String
    Dim a2 As String, lo As ListObject, c As Range
    a2 = CellaSzoveg(ws, "A2")
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(SH_DATA).ListObjects("User")
    If Not lo Is Nothing And Len(a2) > 0 Then
        For Each c In lo.ListColumns("ID").DataBodyRange.Cells
            If StrComp(Trim$(CStr(c.Value)), a2, vbTextCompare) = 0 Then a2 = Trim$(CStr(c.Value)): Exit For
        Next c
    End If
    On Error GoTo 0
    If Len(a2) = 0 Then a2 = UCase$(Environ$("USERNAME"))
    FelhasznaloId = FajlnevResz(a2)
End Function

' v9.21: fajlnevbe valo resz: ekezet nelkul (Ekezetlen + ss), csak kiirhato
' ASCII, a tiltott karakterek helyett kotojel, egyszeres szokozok.
Private Function FajlnevResz(ByVal s As String) As String
    Dim i As Long, ch As String, k As Long, ki As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    s = Ekezetlen(s)
    s = Replace(s, ChrW(223), "ss")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): k = AscW(ch)
        If InStr(1, "\/:*?""<>|", ch) > 0 Then
            ki = ki & "-"
        ElseIf k >= 32 And k <= 126 Then
            ki = ki & ch
        End If
    Next i
    Do While InStr(ki, "  ") > 0: ki = Replace(ki, "  ", " "): Loop
    FajlnevResz = Trim$(ki)
End Function

' v9.21: mappa letrehozasa a teljes utvonalon (a hianyzo szinteket is)
Private Sub MappaLetrehozas(ByVal ut As String)
    Dim reszek As Variant, i As Long, aktualis As String
    reszek = Split(ut, "\")
    aktualis = reszek(0)
    For i = 1 To UBound(reszek)
        aktualis = aktualis & "\" & reszek(i)
        If Len(Dir$(aktualis, vbDirectory)) = 0 Then MkDir aktualis
    Next i
End Sub

' Targy (v9.21): Cikkszam | Tipus | Megnevezes | Preseles | Hibakep | Gyarto berendezes | Fotoblatt
' (A6 | G6 | C6 | E6 | hibakep-cella | H6 | "Fotoblatt"), ekezetekkel, ahogy a
' cellakban allnak. Az ures mezo kimarad, igy nem
' lesz a targyban "| |".
Private Function EmailTargy(ws As Worksheet) As String

    Dim reszek As Variant, resz As Variant, s As String, hibakep As String
    Dim cel As Range

    Set cel = HibakepCella(ws)
    If Not cel Is Nothing Then hibakep = CellaSzoveg(ws, cel.Cells(1, 1).Address)

    reszek = Array(CellaSzoveg(ws, "A6"), CellaSzoveg(ws, "G6"), CellaSzoveg(ws, "C6"), _
                   CellaSzoveg(ws, "E6"), hibakep, CellaSzoveg(ws, "H6"), "Fotoblatt")
    For Each resz In reszek
        If Len(resz) > 0 Then s = s & IIf(Len(s) > 0, " | ", "") & resz
    Next resz
    EmailTargy = s

End Function

' Cella szovege egy sorban, hibaertek- es Empty-biztosan
Private Function CellaSzoveg(ws As Worksheet, ByVal cim As String) As String
    Dim v As Variant
    v = ws.Range(cim).Value
    If IsError(v) Or IsEmpty(v) Then Exit Function
    CellaSzoveg = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
End Function

' Szoveg HTML-be: & < > " escape, minden nem ASCII karakter &#kod; alakban
' (igy a forras ASCII marad, es az ekezetek is helyesen jelennek meg)
Private Function HtmlSzoveg(ByVal s As String) As String
    Dim i As Long, k As Long, ch As String, ki As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        k = AscW(ch): If k < 0 Then k = k + 65536
        Select Case ch
            Case "&": ki = ki & "&amp;"
            Case "<": ki = ki & "&lt;"
            Case ">": ki = ki & "&gt;"
            Case """": ki = ki & "&quot;"
            Case Else
                If k > 127 Then ki = ki & "&#" & k & ";" Else ki = ki & ch
        End Select
    Next i
    HtmlSzoveg = ki
End Function
