Attribute VB_Name = "modBeallitas"
' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.2 - 2026.09.22.
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

    '--- 6. Szerkezeti normalizalas + A4 igazitas -------------------------
    HibafeltarasKitoltes ws
    uzenet = uzenet & "- Hibafeltaras blokk es kepterulet igazitva" & vbLf

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "A Fotolap beallitasa kesz:" & vbLf & vbLf & uzenet, _
           vbInformation, "Fotolap beallitas"

    If MsgBox("Beolvassam most a TERMEKADATOK fajlt (R-szamok)?", _
              vbQuestion + vbYesNo, "Fotolap beallitas") = vbYes Then
        TermekadatokFrissitese
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

    horgonyok = Array("Hibafeltaras", "Hibakep / Fehlerbild", "WZG IH", _
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
    hdrSor = SorKereses(ws, "Hibakep / Fehlerbild", 2, 200)

    If cimSor = 0 Or hdrSor = 0 Or hdrSor - 1 < cimSor Then
        MsgBox "A blokk hatarai nem allapithatok meg.", vbExclamation
        Exit Sub
    End If

    ws.Rows(cimSor & ":" & CStr(hdrSor - 1)).Hidden = False
    MsgBox "A Hibafeltaras blokk (" & cimSor & "-" & (hdrSor - 1) & ". sor) lathato.", _
           vbInformation

End Sub
