' Keszitette es karbantartja : Sardy Peter (PESARDY)
' Mercedes-Benz Manufacturing Hungary, Kecskemet - PW
' Verzio 9.20 - 2026.10.15.
'   9.20: verzio-igazitas (uj mappaszerkezet), a lapkod nem valtozott.
'   9.19: E-mail gomb lathatosaga (modBeallitas.EmailGombLathatosag):
'         valtozaskor (A2), lapra lepeskor es kijeloleskor igazodik, igy egy
'         lathato gombbal mentett fajlban is az elso kattintasra eltunik
'         annal, aki nem jogosult.
'   9.18: verzio-igazitas a modFotolap 9.18-hoz, a lapkod nem valtozott.
'   9.17: - Az E4/E6 logika egy helyen: AnyagFeliratBeallitas (modTermekadatok),
'           az AnyagFelirat csak azt hivja (tovabbra is elsokent).
'         - H6/G6 valtozaskor nincs A6 lista ujraepites (a lista nem fugg
'           toluk, v9.5). Lapra lepeskor csak akkor epul ujra, ha az Article
'           valtozott (A6ListaElavult).
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

    ' v9.19: az e-mail gomb az A2 (es a Windows-felhasznalo) szerint
    EmailGombLathatosag Me

    If Not Intersect(Target, Me.Range("A6")) Is Nothing Then
        AnyagFelirat                     ' elobb, hogy semmi ne akassza meg
        MegnevezesKitoltes Me
    End If

    ' v9.17: H6/G6 valtozasa nem epiti ujra az A6 listat (nem fugg toluk)
    If Not Intersect(Target, Me.Range("H6")) Is Nothing Then
        HibafeltarasKitoltes Me
    End If

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
' v9.17: csak akkor epul ujra, ha az Article valtozott (gyors atlepes).
Private Sub Worksheet_Activate()

    On Error Resume Next
    Application.EnableEvents = False
    If A6ListaElavult() Then A6ListaFrissites Me
    EmailGombLathatosag Me                       ' v9.19
    Application.EnableEvents = True

End Sub


' v9.19: az e-mail gomb lathatosaga mar az elso kattintasnal helyreall
' (pl. ha a fajl lathato gombbal mentodott, es mas nyitja meg).
' Csak akkor nyul a laphoz, ha valtozas kell, igy a kijeloles gyors marad.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    EmailGombLathatosag Me

End Sub


'--------------------------------------------------------------------------
' E4 / E6 az A6 elso betuje alapjan.
' A legordulo szovege ("R244 637 0100 . Megnevezes . Tipus") es a beirt
' cikkszam is a szammal kezdodik, ezert eleg az elso karakter.
' v9.17: a logika a modTermekadatok.AnyagFeliratBeallitas-ban van.
'--------------------------------------------------------------------------
Private Sub AnyagFelirat()

    AnyagFeliratBeallitas Me, CStr(Me.Range("A6").Value)

End Sub


'--------------------------------------------------------------------------
' True, ha a modositott cellak kozul legalabb egyben van ertek.
' Torles (Delete) utan False -> nincs datum/felhasznalo belyeg.
'--------------------------------------------------------------------------
Private Function VanErtek(ByVal Target As Range) As Boolean

    On Error Resume Next
    VanErtek = (Application.WorksheetFunction.CountA(Target) > 0)

End Function
