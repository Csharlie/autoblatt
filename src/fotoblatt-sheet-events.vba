' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.15 - 2026.10.08.
'   9.15: egyseges cikkszam-formatum (pl. "R244 637 0100"), csak a komment
'         peldaja frissult, a logika nem valtozott.
'   9.14: - E4/E6 kozvetlenul az A6 alapjan valt, MINDEN mas lepes elott
'           (R-szam -> "Alapanyag/Material" + "Platine"; A-szam vagy ures
'           A6 -> "Preseles/Abpressung", az automatikus "Platine" torolve).
'         - Torles (ures ertek beirasa) nem ir be datumot es felhasznalot,
'           igy a mezok teljesen uresre torolhetok.
'         - Csak ASCII karakterek (ekezet ChrW-vel), a regi hibas
'           ekezetek javitva.
'
' A kod Sardy Peter munkaja. Modositas elott egyeztess vele.
'
' HELYE: a Fotoblatt lap kodmodulja (VBE: shFotoblatt). A teljes tartalmat
' cserold erre.

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    ' Datum (I3) es felhasznalo (A2) csak valodi bevitelre, torlesre nem
    If VanErtek(Target) Then SzerkesztesBelyeg Me, Target

    If Not Intersect(Target, Me.Range("A2")) Is Nothing Then FelhasznaloKitoltes Me

    If Not Intersect(Target, Me.Range("A6")) Is Nothing Then
        AnyagFelirat                     ' elobb, hogy semmi ne akassza meg
        MegnevezesKitoltes Me
    End If

    If Not Intersect(Target, Me.Range("H6")) Is Nothing Then
        HibafeltarasKitoltes Me
        A6ListaFrissites Me
    End If

    If Not Intersect(Target, Me.Range("G6")) Is Nothing Then A6ListaFrissites Me

CleanExit:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Hiba a Fot" & ChrW(243) & "lap kit" & ChrW(246) & "lt" & ChrW(233) & _
           "sekor: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub


' Lapra lepeskor frissul az A6 lista, igy a Parts lapon kezzel vegzett
' modositasok is azonnal latszanak a legordulokben.
Private Sub Worksheet_Activate()

    On Error Resume Next
    Application.EnableEvents = False
    A6ListaFrissites Me
    Application.EnableEvents = True

End Sub


'--------------------------------------------------------------------------
' E4 / E6 az A6 elso betuje alapjan.
' A legordulo szovege ("R244 637 0100 . Megnevezes . Tipus") es a beirt
' cikkszam is a szammal kezdodik, ezert eleg az elso karakter.
'--------------------------------------------------------------------------
Private Sub AnyagFelirat()

    Dim elso As String

    elso = UCase$(Left$(Trim$(CStr(Me.Range("A6").Value)), 1))

    If elso = "R" Then
        Me.Range("E4").Value = "Alapanyag/Material"
        Me.Range("E6").Value = "Platine"
    Else
        Me.Range("E4").Value = "Pr" & ChrW(233) & "sel" & ChrW(233) & "s/Abpressung"
        If StrComp(Trim$(CStr(Me.Range("E6").Value)), "Platine", vbTextCompare) = 0 Then
            Me.Range("E6").Value = ""
        End If
    End If

End Sub


'--------------------------------------------------------------------------
' True, ha a modositott cellak kozul legalabb egyben van ertek.
' Torles (Delete) utan False -> nincs datum/felhasznalo belyeg.
'--------------------------------------------------------------------------
Private Function VanErtek(ByVal Target As Range) As Boolean

    On Error Resume Next
    VanErtek = (Application.WorksheetFunction.CountA(Target) > 0)

End Function
