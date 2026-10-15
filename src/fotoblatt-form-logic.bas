Attribute VB_Name = "modFotolap"
' Fejlesztes es karbantartas : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary Kft. - Kecskemet, PW
' Verzio 9.19 - 2026.10.14.
'   9.19: verzio-igazitas (e-mail gomb a modBeallitas-ban), ebben a
'         modulban csak a FOTOBLATT_VERZIO valtozott.
'   9.18c: Utomunka (NA) is: alapertek (eszlelve gyartas soran, mintadarab,
'          elozo RK2 darabok = nem) es ures = szurke / kitoltott = feher.
'          Ezzel mindharom kerdesblokk egyforman viselkedik.
'   9.18b: 7.50 (LBL) is: alapertekek (RK2 = nem, palettak megjelolese =
'          igen, zarolas = nem) es ures = szurke / kitoltott = feher.
'   9.18: Hibafeltaras valaszok (G:I):
'         - Alapertek (KerdesAlapertek): ServoTPS-nel a 3 "megtalalhato"
'           kerdes valasza a Data!IgenNem tabla "nem" kezdetu erteke
'           ("nem / nein"). Nincs hozza kulon oszlop a Data lapon.
'         - Szurkites (SZURKE_TABLAK, most csak ServoTPS = 7.41 / 7.22):
'           ures valasz szurke hatteru, kitoltott feher. Feltetelesen
'           formazva, igy kezi bevitelre, beillesztesre, torlesre is
'           azonnal valt, esemenykezelo nelkul.
'         - A tobbi berendezesnel a valaszcellak hattere torlodik (nincs
'           kitoltes, nincs felteteles formazas), igy berendezesvaltaskor
'           nem marad ott a regi szurkites.
'         ValidaciosLista: az igen/nem mintak kulcsszavasak lettek ("-e" /
'         " e" nelkul), igy a kerdes atfogalmazasa nem rontja el. Uj igen/nem
'         legordulot kap: ServoTPS/LBL "el lett mentve?", LBL "Palettak
'         megjelolesre kerultek?", "Zarolas tortent?", "Okozott allasidot?",
'         "Koveto gyartast erintheti?", NA "megtalalhato a ..." (2 kerdes).
'   9.17: FOTOBLATT_VERZIO - a kiadasi verzio egyetlen helye (a lablec
'         "Verzio" cellajat a modBeallitas.LablecVerzio irja). A H_CIM,
'         H_FEJLEC es H_VEG horgony Public lett (a modBeallitas is hasznalja).
'         A kerdesblokk alatti fejlecsor horgonya: "Hibaleiras /
'         Fehlerbeschreibung" (a lapon atnevezve; a "Hibakep / Fehlerbild"
'         felirat mostantol csak a cimsorban, az E oszlopban all).
'
' Belso fejlesztes. Modositas kizarolag a karbantartoval egyeztetve.

Option Explicit
Option Compare Text

'==========================================================================
' Fotolap (Fotoblatt) - futasideju logika
'--------------------------------------------------------------------------
' A lapmodul csak esemenyt fog el, a munkat ezek az eljarasok vegzik.
' Kiadasi verzio: l. FOTOBLATT_VERZIO lent.
'
' FONTOS: ez a fajl szandekosan CSAK ASCII karaktereket tartalmaz.
' A VBE a .bas fajlokat a rendszer kodlapjaval olvassa be, nem UTF-8-kent,
' ezert az ekezetes betuk importalaskor elromlananak. A lapon levo ekezetes
' szovegeket az Ekezetlen() fuggveny normalizalja, igy az osszehasonlitas
' ekezet-fuggetlen es a kodlaptol fuggetlenul mukodik.
'
' A szerkezetet horgonyszovegek hatarozzak meg, nem beegetett sorszamok.
' A horgonyokat az A:I oszlopokban keressuk, igy mindegy, hogy a szoveg az
' A vagy az E oszlopban ul, es hogy a sor hogyan van egyesitve.
'==========================================================================

'--- Lapnevek egy helyen -------------------------------------------------
Public Const SH_FOTO  As String = "Fotoblatt"
Public Const SH_DATA  As String = "Data"
Public Const SH_PARTS As String = "Parts"

'--- Kiadasi verzio (v9.17): EZ az egyetlen hely, ahol szerepel -----------
Public Const FOTOBLATT_VERZIO As String = "9.19"

'--- Viselkedes ----------------------------------------------------------
' True  = minden szerkesztes az aktualis gepfelhasznalora allitja az A2-t
' False = az A2 csak akkor toltodik, ha meg ures (kezi valasztas megmarad)
Public Const A2_MINDIG_FELULIR As Boolean = True

'--- Az urlap szerkezete -------------------------------------------------
Public Const ELSO_KERDES_SOR As Long = 9     ' tartalek, ha a cimsor nem talalhato
Private Const MAX_KERES_SOR  As Long = 200   ' meddig keressuk a horgonyszovegeket
Private Const UTOLSO_OSZLOP  As Long = 9     ' A:I

'--- Hibafeltaras valaszok (v9.18) ---------------------------------------
' Azok a Data-tablak, ahol az ures valasz szurke, a kitoltott feher.
' Bovites: "|ServoTPS|LBL|" stb.
Private Const SZURKE_TABLAK    As String = "|ServoTPS|LBL|NA|"

'--- Horgonyszovegek, ekezet nelkuli alakban -----------------------------
Public Const H_CIM      As String = "Hibafeltaras"     ' v9.17: Public (modBeallitas)
Public Const H_FEJLEC   As String = "Hibaleiras / Fehlerbeschreibung"   ' v9.17
Private Const H_WZG     As String = "WZG IH"
Private Const H_KEPEK   As String = "Kepek/Bilder"
Private Const H_INTEZK  As String = "Intezkedesek WIH"
Public Const H_VEG      As String = "Keszitette"       ' v9.17: Public (modBeallitas)

'--- A4 igazitas ---------------------------------------------------------
Private Const A4_MAGASSAG_PT As Double = 841.89   ' A4 allo, pont
Private Const A4_TARTALEK_PT As Double = 6        ' biztonsagi tartalek
Private Const KEP_MIN_SORMAGASSAG As Double = 6
Private Const KEP_MAX_SORMAGASSAG As Double = 40

Private mElerhetoPt As Double    ' 0 = meg nincs kiszamolva (munkamenetre gyorsitotarazva)


'==========================================================================
' 1. Szerkesztesbelyeg: datum + felhasznalo
'==========================================================================
Public Sub SzerkesztesBelyeg(ws As Worksheet, Target As Range)

    Dim uid As String

    ' Datum: minden valodi szerkesztesre, kiveve ha maga az I3 valtozott
    If Intersect(Target, ws.Range("I3")) Is Nothing Then
        ws.Range("I3").Value = Date
    End If

    ' Felhasznalo: ne irjuk felul, ha epp az A2-t szerkesztik kezzel
    If Not Intersect(Target, ws.Range("A2")) Is Nothing Then Exit Sub

    If Not A2_MINDIG_FELULIR Then
        If Len(Trim$(CStr(ws.Range("A2").Value))) > 0 Then Exit Sub
    End If

    uid = UCase$(Trim$(Environ$("USERNAME")))
    If Len(uid) = 0 Then Exit Sub

    ' Csak akkor irunk, ha tenyleg mas van bent
    If StrComp(Trim$(CStr(ws.Range("A2").Value)), uid, vbTextCompare) = 0 Then Exit Sub

    ws.Range("A2").Value = uid
    FelhasznaloKitoltes ws

End Sub


'==========================================================================
' 2. A2 -> B3 (nev) es F3 (muszak) a Data!User tablabol
'==========================================================================
Public Sub FelhasznaloKitoltes(ws As Worksheet)

    Dim lo As ListObject
    Dim m As Variant

    Set lo = ThisWorkbook.Worksheets(SH_DATA).ListObjects("User")
    m = Application.Match(ws.Range("A2").Value, lo.ListColumns("ID").DataBodyRange, 0)

    If IsError(m) Then
        ws.Range("B3").Value = ""
        ws.Range("F3").Value = ""
    Else
        ws.Range("B3").Value = lo.ListColumns("Name").DataBodyRange.Cells(m, 1).Value
        ws.Range("F3").Value = lo.ListColumns("Shift").DataBodyRange.Cells(m, 1).Value
    End If

End Sub


'==========================================================================
' 3. A6 -> C6 (megnevezes), G6 (tipus) a Parts!Article tablabol
'==========================================================================
Public Sub MegnevezesKitoltes(ws As Worksheet)

    ' v9.2: a legordulo "Cikkszam . Megnevezes . Tipus" szoveget ad,
    ' ezt a modTermekadatok bontja szet (A6 = szam, C6 = nev, G6 = tipus).
    CikkszamValasztas ws

End Sub


'==========================================================================
' 4. H6 -> Hibafeltaras blokk a Data lap tablajabol
'--------------------------------------------------------------------------
' Lekepezett berendezes  -> a blokk lathato, a kerdesek betoltodnek
' Nem lekepezett ertek   -> a blokk (cimsorral egyutt) elrejtve, uresen
'==========================================================================
Public Sub HibafeltarasKitoltes(ws As Worksheet)

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim sVal As String
    Dim cimSor As Long, elsoKerdes As Long, hdrSor As Long, ujHdr As Long
    Dim jelenlegi As Long, kell As Long, diff As Long
    Dim i As Long, r As Long
    Dim kerdes As String
    Dim vanTabla As Boolean
    Dim lepes As String
    Dim hibaSzam As Long, hibaSzoveg As String
    Dim szurkites As Boolean                           ' v9.18
    Dim valasz As Range                                ' v9.18
    Dim alapertek As String                            ' v9.18

    On Error GoTo Hiba
    lepes = "indulas"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsData = ThisWorkbook.Worksheets(SH_DATA)
    sVal = Ekezetlen(Trim$(Replace(CStr(ws.Range("H6").Value), ",", ".")))

    '--- Gyartoberendezes -> tabla ---------------------------------------
    Select Case sVal
        Case "7.41", "7.22":  Set tbl = wsData.ListObjects("ServoTPS")
        Case "7.50":          Set tbl = wsData.ListObjects("LBL")
        Case "Utomunka":      Set tbl = wsData.ListObjects("NA")
        Case Else:            Set tbl = Nothing
    End Select

    vanTabla = Not (tbl Is Nothing)
    If vanTabla Then
        If tbl.DataBodyRange Is Nothing Then vanTabla = False
    End If

    '--- Szerkezeti horgonyok --------------------------------------------
    lepes = "'" & H_CIM & "' cimsor keresese"
    cimSor = SorKereses(ws, H_CIM, 2, MAX_KERES_SOR)
    If cimSor = 0 Then cimSor = ELSO_KERDES_SOR - 1
    elsoKerdes = cimSor + 1

    lepes = "'" & H_FEJLEC & "' fejlecsor keresese"
    hdrSor = SorKereses(ws, H_FEJLEC, elsoKerdes, MAX_KERES_SOR)
    If hdrSor = 0 Then
        Err.Raise vbObjectError + 513, , _
            "Nem talalhato a '" & H_FEJLEC & "' sor az A:I oszlopokban."
    End If

    '--- A blokkot mindig lathatova tesszuk, hogy kiszamithato legyen ----
    lepes = "blokk lathatova tetele"
    BlokkLathatosag ws, cimSor, hdrSor - 1, True

    '--- Meretezes: csak ha van hozza tabla ------------------------------
    If vanTabla Then
        jelenlegi = hdrSor - elsoKerdes
        kell = tbl.DataBodyRange.Rows.Count
        If kell < 1 Then kell = 1

        diff = kell - jelenlegi
        lepes = "sorok beszurasa/torlese (diff=" & diff & ")"
        If diff > 0 Then
            If hdrSor - 1 >= elsoKerdes Then
                ws.Rows(hdrSor - 1).Copy                 ' kesz kerdessor a minta
                ws.Rows(hdrSor).Resize(diff).Insert Shift:=xlDown
                Application.CutCopyMode = False
            Else
                ws.Rows(hdrSor).Resize(diff).Insert Shift:=xlDown
            End If
        ElseIf diff < 0 Then
            ws.Rows(hdrSor + diff).Resize(-diff).Delete Shift:=xlUp
        End If

        ' A fejlecsort ujra megkeressuk, nem szamoljuk
        lepes = "fejlecsor ujrakeresese"
        ujHdr = SorKereses(ws, H_FEJLEC, elsoKerdes, MAX_KERES_SOR)
        If ujHdr > 0 Then hdrSor = ujHdr Else hdrSor = hdrSor + diff
    End If

    '--- Kerdessorok ujraepitese es uritese -------------------------------
    For r = elsoKerdes To hdrSor - 1
        lepes = "kerdessor ujraepitese (sor " & r & ")"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, UTOLSO_OSZLOP)).UnMerge
        ws.Cells(r, 1).Value = ""
        ws.Cells(r, 7).Value = ""
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Merge     ' kerdes: A:F
        ws.Range(ws.Cells(r, 7), ws.Cells(r, UTOLSO_OSZLOP)).Merge   ' valasz: G:I
        ValaszLegordulo ws.Range(ws.Cells(r, 7), ws.Cells(r, UTOLSO_OSZLOP)), ""
        ' v9.18: a beszurt/masolt sorok formazasa se maradjon itt
        ValaszHatter ws.Range(ws.Cells(r, 7), ws.Cells(r, UTOLSO_OSZLOP)), False
    Next r

    lepes = "fejlecsor rendezese (sor " & hdrSor & ")"
    FejlecSorRendezes ws, hdrSor

    '--- Kitoltes vagy elrejtes ------------------------------------------
    If vanTabla Then
        szurkites = (InStr(1, SZURKE_TABLAK, "|" & tbl.Name & "|", vbTextCompare) > 0)

        For i = 1 To hdrSor - elsoKerdes
            r = elsoKerdes + i - 1
            lepes = "kerdes betoltese (sor " & r & ")"
            kerdes = CStr(tbl.DataBodyRange.Cells(i, 1).Value)
            ws.Cells(r, 1).Value = kerdes

            Set valasz = ws.Range(ws.Cells(r, 7), ws.Cells(r, UTOLSO_OSZLOP))
            ValaszLegordulo valasz, kerdes

            ' v9.18: alapertek (l. KerdesAlapertek)
            lepes = "alapertek beirasa (sor " & r & ")"
            alapertek = KerdesAlapertek(tbl.Name, kerdes)
            If Len(alapertek) > 0 Then valasz.Cells(1, 1).Value = alapertek

            ' v9.18: ures = szurke, kitoltott = feher
            lepes = "valasz hattere (sor " & r & ")"
            ValaszHatter valasz, szurkites
        Next i
    Else
        lepes = "blokk elrejtese"
        BlokkLathatosag ws, cimSor, hdrSor - 1, False
    End If

    '--- Kepteruletek igazitasa, hogy egy A4 oldalra ferjen ---------------
    lepes = "A4 igazitas"
    A4Igazitas ws

    Exit Sub

Hiba:
    hibaSzam = Err.Number
    hibaSzoveg = Err.Description
    Application.CutCopyMode = False
    Err.Raise hibaSzam, , "[" & lepes & "] " & hibaSzoveg

End Sub


'--------------------------------------------------------------------------
' A Hibafeltaras blokk (cimsor + kerdessorok) elrejtese / megmutatasa
'--------------------------------------------------------------------------
Private Sub BlokkLathatosag(ws As Worksheet, ByVal cimSor As Long, _
                            ByVal utolsoSor As Long, ByVal lathato As Boolean)

    If utolsoSor < cimSor Then Exit Sub
    ws.Rows(cimSor & ":" & CStr(utolsoSor)).Hidden = Not lathato

End Sub


'--------------------------------------------------------------------------
' A fejlecsor rendezese.
' Ha a fejlec szovege mar az A oszlopban van (teljes szelessegu sor), akkor
' a sor rendben van, nem nyulunk hozza. Csak a regi elrendezest javitjuk,
' ahol az A:D reszben meg egy kerdes ul, a fejlec pedig az E:I reszben.
'--------------------------------------------------------------------------
Private Sub FejlecSorRendezes(ws As Worksheet, ByVal hdrSor As Long)

    Dim hdrSzoveg As String

    If InStr(1, Ekezetlen(CStr(ws.Cells(hdrSor, 1).Value)), H_FEJLEC, vbTextCompare) > 0 Then
        Exit Sub                     ' mar rendben van, nem bantjuk
    End If

    hdrSzoveg = CStr(ws.Cells(hdrSor, 5).Value)
    If Len(Trim$(hdrSzoveg)) = 0 Then Exit Sub    ' nem az ismert regi elrendezes

    ws.Range(ws.Cells(hdrSor, 1), ws.Cells(hdrSor, UTOLSO_OSZLOP)).UnMerge
    ws.Cells(hdrSor, 1).Value = ""                             ' a regi utolso kerdes
    ws.Cells(hdrSor, 5).Value = hdrSzoveg
    ws.Range(ws.Cells(hdrSor, 1), ws.Cells(hdrSor, 4)).Merge   ' ures: A:D
    ws.Range(ws.Cells(hdrSor, 5), ws.Cells(hdrSor, UTOLSO_OSZLOP)).Merge   ' fejlec: E:I

End Sub


'--------------------------------------------------------------------------
' v9.18: valaszcella (G:I) hattere.
' szurkites = True : alap kitoltes szurke (Hatter 1, 25%-kal sotetebb), es
'                    felteteles formazas: ha a G cellaban van szoveg -> feher.
'                    Soronkent abszolut hivatkozas ($G$12), igy nem fugg az
'                    aktiv cellatol (a FormatConditions.Add relativ
'                    hivatkozast az aktiv cellahoz kepest ertelmezne).
' szurkites = False: nincs kitoltes, nincs felteteles formazas.
'--------------------------------------------------------------------------
Private Sub ValaszHatter(cel As Range, ByVal szurkites As Boolean)

    Dim fc As FormatCondition

    cel.FormatConditions.Delete

    If Not szurkites Then
        cel.Interior.Pattern = xlNone
        Exit Sub
    End If

    With cel.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893          ' "25%-kal sotetebb"
    End With

    Set fc = cel.FormatConditions.Add(Type:=xlExpression, _
             Formula1:="=LEN(TRIM(" & cel.Cells(1, 1).Address(True, True) & "))>0")
    With fc.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0                           ' feher
    End With
    fc.StopIfTrue = False

End Sub


'--------------------------------------------------------------------------
' v9.18: a kerdes valaszanak alaperteke. Ures = nincs alapertek.
' ITT LEHET BOVITENI. A kerdes szoveget ekezet nelkul hasonlitjuk (mint a
' ValidaciosLista-ban), igy a Data-tabla sorrendjetol fuggetlen.
' A valasz szovege nincs beegetve: az IgenNem tablabol jon.
'--------------------------------------------------------------------------
Private Function KerdesAlapertek(ByVal tablaNev As String, ByVal kerdes As String) As String

    Dim k As String
    k = Ekezetlen(Trim$(kerdes))
    KerdesAlapertek = ""
    If Len(k) = 0 Then Exit Function

    Select Case LCase$(tablaNev)

        Case "servotps"                                  ' 7.41 / 7.22
            ' elozo gyartas utolso darabja, mintadarab, elozo RK2 darabok
            If Tartalmaz(k, "megtalalhato") Then
                KerdesAlapertek = IgenNemErtek("nem")
            End If

        Case "lbl"                                       ' 7.50
            Select Case True
                Case Tartalmaz(k, "megtalalhato")        ' elozo RK2 darabok
                    KerdesAlapertek = IgenNemErtek("nem")
                Case Tartalmaz(k, "megjelolesre")        ' palettak megjelolve
                    KerdesAlapertek = IgenNemErtek("igen")
                Case Tartalmaz(k, "zarolas tortent"), _
                     Tartalmaz(k, "tortent-e zarolas"), _
                     Tartalmaz(k, "tortent e zarolas")   ' zarolas
                    KerdesAlapertek = IgenNemErtek("nem")
            End Select

        Case "na"                                        ' Utomunka
            ' eszlelve gyartas soran, mintadarab, elozo RK2 darabok
            If Tartalmaz(k, "eszlelve") Or Tartalmaz(k, "megtalalhato") Then
                KerdesAlapertek = IgenNemErtek("nem")
            End If

    End Select

End Function


'--------------------------------------------------------------------------
' v9.18: a Data!IgenNem tabla elso olyan erteke, amely a megadott szoval
' kezdodik ("nem" -> "nem / nein", "igen" -> "igen / ja"). Sorrend- es
' szovegvaltozas-fuggetlen. Ures, ha a tabla vagy az ertek nem talalhato.
'--------------------------------------------------------------------------
Private Function IgenNemErtek(ByVal kezdet As String) As String

    Dim lo As ListObject
    Dim c As Range
    Dim ertek As String

    IgenNemErtek = ""

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(SH_DATA).ListObjects("IgenNem")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For Each c In lo.DataBodyRange.Columns(1).Cells
        If Not IsError(c.Value) Then
            ertek = Trim$(CStr(c.Value))
            If Ekezetlen(ertek) Like kezdet & "*" Then     ' Option Compare Text
                IgenNemErtek = ertek
                Exit Function
            End If
        End If
    Next c

End Function


'==========================================================================
' 5. Melyik kerdeshez melyik legordulo
'--------------------------------------------------------------------------
' ITT LEHET BOVITENI. A sorrend szamit: a fajlagosabb szabaly elol.
' A kerdes szoveget ekezet nelkul hasonlitjuk, ezert itt is ekezet nelkul
' kell megadni a kulcsszavakat. Ures visszateres = szabad szoveges mezo.
'==========================================================================
Private Function ValidaciosLista(ByVal kerdes As String) As String

    Dim k As String
    k = Ekezetlen(Trim$(kerdes))
    ValidaciosLista = ""
    If Len(k) = 0 Then Exit Function

    Select Case True

        Case Tartalmaz(k, "milyen etekelest kapot"), _
             Tartalmaz(k, "milyen ertekelest kapott")
            ValidaciosLista = "PrioLista"

        Case Tartalmaz(k, "szerszam operacio megnevezese"), _
             Tartalmaz(k, "hiba keletkezesi helye")
            ValidaciosLista = "OPLista"

        ' v9.18: kulcsszavak "-e" / " e" nelkul, hogy az atfogalmazas
        ' ("Tortent-e zarolas?" -> "Zarolas tortent?") ne rontsa el.
        ' Vigyazat: a "Zarolas, ha volt, akkor mennyi ..." szabad szoveg
        ' marad, ezert nem eleg a puszta "zarolas".
        Case Tartalmaz(k, "megtalalhato"), _
             Tartalmaz(k, "lett mentve"), _
             Tartalmaz(k, "lett-e mentve"), _
             Tartalmaz(k, "eszlelve"), _
             Tartalmaz(k, "megjelolesre"), _
             Tartalmaz(k, "zarolas tortent"), _
             Tartalmaz(k, "tortent-e zarolas"), _
             Tartalmaz(k, "tortent e zarolas"), _
             Tartalmaz(k, "lehajlas"), _
             Tartalmaz(k, "okozott"), _
             Tartalmaz(k, "erintheti")
            ValidaciosLista = "IgenNemLista"

    End Select

End Function


Private Function Tartalmaz(ByVal Szoveg As String, ByVal minta As String) As Boolean
    Tartalmaz = (InStr(1, Szoveg, minta, vbTextCompare) > 0)
End Function


Private Sub ValaszLegordulo(cel As Range, ByVal kerdes As String)

    Dim nev As String
    nev = ValidaciosLista(kerdes)

    On Error Resume Next
    cel.Validation.Delete
    On Error GoTo 0

    If Len(nev) = 0 Then Exit Sub

    On Error Resume Next
    With cel.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & nev
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False              ' listan kivuli ertek is beirhato
    End With
    On Error GoTo 0

End Sub


'==========================================================================
' 6. A4 igazitas: a ket kepterulet veszi fel a maradek helyet
'--------------------------------------------------------------------------
' A rejtett sorok nem szamitanak bele, igy ha a Hibafeltaras blokk el van
' rejtve, a kepteruletek automatikusan nagyobbak lesznek.
'==========================================================================
Public Sub A4Igazitas(ws As Worksheet)

    Dim hdrSor As Long, wzgSor As Long, kepekSor As Long
    Dim intezkSor As Long, vegSor As Long
    Dim kep1Tol As Long, kep1Ig As Long, kep2Tol As Long, kep2Ig As Long
    Dim fixPt As Double, szabadPt As Double, sorMagassag As Double
    Dim kepSorok As Long
    Dim r As Long

    hdrSor = SorKereses(ws, H_FEJLEC, 2, MAX_KERES_SOR)
    wzgSor = SorKereses(ws, H_WZG, 2, MAX_KERES_SOR)
    kepekSor = SorKereses(ws, H_KEPEK, 2, MAX_KERES_SOR)
    intezkSor = SorKereses(ws, H_INTEZK, 2, MAX_KERES_SOR)
    vegSor = SorKereses(ws, H_VEG, 2, MAX_KERES_SOR)

    ' Ha barmelyik horgony hianyzik, inkabb nem nyulunk a magassagokhoz
    If hdrSor = 0 Or wzgSor = 0 Or kepekSor = 0 Or intezkSor = 0 Or vegSor = 0 Then Exit Sub

    kep1Tol = hdrSor + 1:   kep1Ig = wzgSor - 1        ' Hibaleiras terulet
    kep2Tol = kepekSor + 1: kep2Ig = intezkSor - 1     ' Kepek/Bilder terulet
    If kep1Ig < kep1Tol Or kep2Ig < kep2Tol Then Exit Sub

    kepSorok = (kep1Ig - kep1Tol + 1) + (kep2Ig - kep2Tol + 1)

    '--- Rendelkezesre allo magassag (munkamenetenkent egyszer szamoljuk) --
    If mElerhetoPt <= 0 Then mElerhetoPt = ElerhetoMagassagPt(ws)
    If mElerhetoPt <= 0 Then Exit Sub

    '--- A nem kepteruleti, LATHATO sorok osszmagassaga --------------------
    fixPt = 0
    For r = 1 To vegSor
        If (r < kep1Tol Or r > kep1Ig) And (r < kep2Tol Or r > kep2Ig) Then
            If Not ws.Rows(r).Hidden Then fixPt = fixPt + ws.Rows(r).RowHeight
        End If
    Next r

    szabadPt = mElerhetoPt - fixPt
    sorMagassag = szabadPt / kepSorok

    If sorMagassag < KEP_MIN_SORMAGASSAG Then sorMagassag = KEP_MIN_SORMAGASSAG
    If sorMagassag > KEP_MAX_SORMAGASSAG Then sorMagassag = KEP_MAX_SORMAGASSAG

    ws.Rows(kep1Tol & ":" & CStr(kep1Ig)).RowHeight = sorMagassag
    ws.Rows(kep2Tol & ":" & CStr(kep2Ig)).RowHeight = sorMagassag

    '--- Nyomtatasi terulet a tenyleges utolso sorig ----------------------
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$I$" & vegSor
    On Error GoTo 0

End Sub


Private Function ElerhetoMagassagPt(ws As Worksheet) As Double

    Dim zoomSzazalek As Double

    On Error GoTo Hiba
    With ws.PageSetup
        If VarType(.Zoom) = vbBoolean Then
            zoomSzazalek = 100                  ' FitToPage aktiv, nincs zoom
        Else
            zoomSzazalek = CDbl(.Zoom)
        End If
        If zoomSzazalek <= 0 Then zoomSzazalek = 100

        ElerhetoMagassagPt = (A4_MAGASSAG_PT - .TopMargin - .BottomMargin) _
                             * 100# / zoomSzazalek - A4_TARTALEK_PT
    End With
    Exit Function

Hiba:
    ElerhetoMagassagPt = 0

End Function


'==========================================================================
' 7. Segedek
'==========================================================================

' Sor keresese horgonyszoveg alapjan az A:I oszlopokban, ekezet-fuggetlenul.
' Eloszor az A oszlopot nezzuk vegig (a horgonyok tobbsege ott van), utana a
' tobbit. Igy mindegy, hogy a szoveg az A vagy az E oszlopba keruit, es hogy
' a sor hogyan van egyesitve.
Public Function SorKereses(ws As Worksheet, ByVal minta As String, _
                           ByVal tol As Long, ByVal ig As Long) As Long

    Dim r As Long, c As Long

    SorKereses = 0
    If tol < 1 Then tol = 1

    For r = tol To ig
        If InStr(1, Ekezetlen(CStr(ws.Cells(r, 1).Value)), minta, vbTextCompare) > 0 Then
            SorKereses = r
            Exit Function
        End If
    Next r

    For r = tol To ig
        For c = 2 To UTOLSO_OSZLOP
            If InStr(1, Ekezetlen(CStr(ws.Cells(r, c).Value)), minta, vbTextCompare) > 0 Then
                SorKereses = r
                Exit Function
            End If
        Next c
    Next r

End Function


' Ekezetek levalasztasa. A lapon vegyesen szerepel o / o / o es u / u,
' ezert a normalizalas a forrasadatok kovetkezetlensegeit is elsimitja.
Public Function Ekezetlen(ByVal s As String) As String

    Dim i As Long, ch As String, kod As Long
    Dim ki As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        kod = AscW(ch)
        Select Case kod
            Case 224 To 229, 259, 261:           ch = "a"
            Case 192 To 197, 258, 260:           ch = "A"
            Case 232 To 235, 283, 281, 277:      ch = "e"
            Case 200 To 203, 282, 280:           ch = "E"
            Case 236 To 239:                     ch = "i"
            Case 204 To 207:                     ch = "I"
            Case 242 To 246, 248, 337, 335:      ch = "o"
            Case 210 To 214, 216, 336, 334:      ch = "O"
            Case 249 To 252, 369, 367, 365:      ch = "u"
            Case 217 To 220, 368, 366, 364:      ch = "U"
            Case 253, 255:                       ch = "y"
            Case 221:                            ch = "Y"
            Case 231, 263, 269:                  ch = "c"
            Case 199, 262, 268:                  ch = "C"
            Case 351, 353:                       ch = "s"
            Case 350, 352:                       ch = "S"
            Case 378, 380, 382:                  ch = "z"
            Case 377, 379, 381:                  ch = "Z"
            Case 355, 357:                       ch = "t"
            Case 354, 356:                       ch = "T"
            Case 241, 324, 328:                  ch = "n"
            Case 209, 323, 327:                  ch = "N"
            Case 271, 273:                       ch = "d"
            Case 270, 272:                       ch = "D"
            Case 341, 345:                       ch = "r"
            Case 340, 344:                       ch = "R"
            Case 314, 318:                       ch = "l"
            Case 313, 317:                       ch = "L"
        End Select
        ki = ki & ch
    Next i

    Ekezetlen = ki

End Function


'==========================================================================
' 8. Gyorsitotar uritese (ha az oldalbeallitas kozben valtozik)
'==========================================================================
Public Sub ElerhetoMagassagUjraszamol()
    mElerhetoPt = 0
End Sub
