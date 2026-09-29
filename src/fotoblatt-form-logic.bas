Attribute VB_Name = "modFotolap"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.2 - 2026.09.22.
'
' A kod Sardy Peter munkaja. Modositas elott egyeztess vele.

Option Explicit
Option Compare Text

'==========================================================================
' Fotolap (Fotoblatt) - futasideju logika
'--------------------------------------------------------------------------
' A lapmodul csak esemenyt fog el, a munkat ezek az eljarasok vegzik.
' Verzio: 9.2
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

'--- Viselkedes ----------------------------------------------------------
' True  = minden szerkesztes az aktualis gepfelhasznalora allitja az A2-t
' False = az A2 csak akkor toltodik, ha meg ures (kezi valasztas megmarad)
Public Const A2_MINDIG_FELULIR As Boolean = True

'--- Az urlap szerkezete -------------------------------------------------
Public Const ELSO_KERDES_SOR As Long = 9     ' tartalek, ha a cimsor nem talalhato
Private Const MAX_KERES_SOR  As Long = 200   ' meddig keressuk a horgonyszovegeket
Private Const UTOLSO_OSZLOP  As Long = 9     ' A:I

'--- Horgonyszovegek, ekezet nelkuli alakban -----------------------------
Private Const H_CIM     As String = "Hibafeltaras"
Private Const H_FEJLEC  As String = "Hibakep / Fehlerbild"
Private Const H_WZG     As String = "WZG IH"
Private Const H_KEPEK   As String = "Kepek/Bilder"
Private Const H_INTEZK  As String = "Intezkedesek WIH"
Private Const H_VEG     As String = "Keszitette"

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
    Next r

    lepes = "fejlecsor rendezese (sor " & hdrSor & ")"
    FejlecSorRendezes ws, hdrSor

    '--- Kitoltes vagy elrejtes ------------------------------------------
    If vanTabla Then
        For i = 1 To hdrSor - elsoKerdes
            r = elsoKerdes + i - 1
            lepes = "kerdes betoltese (sor " & r & ")"
            kerdes = CStr(tbl.DataBodyRange.Cells(i, 1).Value)
            ws.Cells(r, 1).Value = kerdes
            ValaszLegordulo ws.Range(ws.Cells(r, 7), ws.Cells(r, UTOLSO_OSZLOP)), kerdes
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

        Case Tartalmaz(k, "megtalalhato-e"), _
             Tartalmaz(k, "el lett-e mentve"), _
             Tartalmaz(k, "eszlelve"), _
             Tartalmaz(k, "kerultek e"), _
             Tartalmaz(k, "tortent e zarolas"), _
             Tartalmaz(k, "lehajlas"), _
             Tartalmaz(k, "okozott e"), _
             Tartalmaz(k, "erintheti e")
            ValidaciosLista = "IgenNemLista"

    End Select

End Function


Private Function Tartalmaz(ByVal szoveg As String, ByVal minta As String) As Boolean
    Tartalmaz = (InStr(1, szoveg, minta, vbTextCompare) > 0)
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

    kep1Tol = hdrSor + 1:   kep1Ig = wzgSor - 1        ' Hibakep terulet
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
