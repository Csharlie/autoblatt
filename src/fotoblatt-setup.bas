Attribute VB_Name = "modBeallitas"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.18 - 2026.10.14.
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
