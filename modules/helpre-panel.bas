Attribute VB_Name = "helprePanel"
Option Explicit

' ============================================================
' AUTOBLATT - HELPRE PANEL (FOTOLAP KITOLTESI SEGEDLET)
' ============================================================
' Letrehozza a 'Helpre Panel' munkalapot 7 szekcios statikus
' instrukcios tartalommal: bevezeto, kitoltesi sorrend tabla,
' jo/rossz peldak, mondatvaz, sajat vazlat blokk, ellenorzolista.
'
' A SaveToDrive nem hivja - kulon makro (AB_HelprePanel).
' ============================================================

' --- Publikus belepesi pont ---
Public Sub RunHelprePanel()
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Nincs aktiv munkafuzet.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, SHEET_HELPRE)

    Call BuildHelperLayout(ws)
    Call FormatHelperSheet(ws)

    Application.ScreenUpdating = True
    ws.Activate
End Sub

' ------------------------------------------------------------
' STATIKUS TARTALOM EPITESE
' ------------------------------------------------------------
' A felhasznaloi beviteli cellak (B28:D31 vazlatpontok, B34:D38
' vegleges szoveg) erintetlenek maradnak.
Private Sub BuildHelperLayout(ByVal ws As Worksheet)
    ' --- Oszlopszelessegek ---
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 34

    ' --- 1. FOCIM ---
    ws.Rows("1:1").RowHeight = 30
    ws.Range("B1").Value = "Fotolap kitoltesi segedlet"

    ' --- 2. ROVID LEIRAS ---
    ws.Rows("2:3").RowHeight = 18
    ws.Range("B2").Value = "Ez a lap segit abban, hogy kezzel, ertheton es kovetkezetesen fogalmazd meg a hibaleirast."
    ws.Range("B3").Value = "Nem automatikus mondatgenerator " & ChrW(8211) & " a cel a jobb minosegu kezi megfogalmazas tamogatasa."
    ws.Rows("4:4").RowHeight = 10

    ' --- 3. KITOLTESI SORREND ---
    ws.Range("B5").Value = "Kitoltesi sorrend"
    ws.Rows("5:5").RowHeight = 22

    ws.Range("A6").Value = "#"
    ws.Range("B6").Value = "Szempont"
    ws.Range("C6").Value = "Mire figyelj?"
    ws.Range("D6").Value = "Pelda"
    ws.Rows("6:6").RowHeight = 18

    Dim szempontok(1 To 6, 1 To 3) As String
    szempontok(1, 1) = "Hol keszult?"
    szempontok(1, 2) = "sor, allomas, zona, oldal"
    szempontok(1, 3) = "7.41 jobb oldal"

    szempontok(2, 1) = "Hol lett eszrevve?"
    szempontok(2, 2) = "RK, utomunka, vegellenorzes, gyartas"
    szempontok(2, 3) = "RK2"

    szempontok(3, 1) = "Mi a konkret hiba?"
    szempontok(3, 2) = "lathato jelenseg, ne altalanos legyen"
    szempontok(3, 3) = "karc / horpadas / szennyezodes"

    szempontok(4, 1) = "Hol van pontosan?"
    szempontok(4, 2) = "melyik elem, melyik resz, melyik oldal"
    szempontok(4, 3) = "jobb elso ajto kilincs alatt"

    szempontok(5, 1) = "Mekkora / mennyire lathato?"
    szempontok(5, 2) = "cm, mm, kis/nagy, kozelrol/fenyben lathato"
    szempontok(5, 3) = "kb. 3 cm-es"

    szempontok(6, 1) = "Mi a teendo?"
    szempontok(6, 2) = "ellenorzes, javitas, polirozas, csere, megfigyeles"
    szempontok(6, 3) = "polirozas elott ellenorizni"

    Dim r As Long
    For r = 1 To 6
        ws.Range("A" & (6 + r)).Value = r
        ws.Range("B" & (6 + r)).Value = szempontok(r, 1)
        ws.Range("C" & (6 + r)).Value = szempontok(r, 2)
        ws.Range("D" & (6 + r)).Value = szempontok(r, 3)
        ws.Rows(6 + r).RowHeight = 18
    Next r

    ' --- 4. JO / ROSSZ PELDA BLOKK ---
    ws.Rows("14:14").RowHeight = 10
    ws.Range("B15").Value = "Peldak " & ChrW(8211) & " jo es rossz megfogalmazas"
    ws.Rows("15:15").RowHeight = 22

    ws.Range("B16").Value = "Nem eleg jo:"
    ws.Range("C16").Value = "hiba az ajton"
    ws.Rows("16:16").RowHeight = 18

    ws.Range("B17").Value = "Nem eleg jo:"
    ws.Range("C17").Value = "karcolas lathato"
    ws.Rows("17:17").RowHeight = 18

    ws.Range("B18").Value = "Jo:"
    ws.Range("C18").Value = "RK2-n eszlelve, a jobb elso ajto kulso feluleten, a kilincs alatt kb. 3 cm-es karc lathato."
    ws.Rows("18:18").RowHeight = 28

    ws.Range("B19").Value = "Jo:"
    ws.Range("C19").Value = "Gyartas kozben, 7.41 allomason, bal elso ajto belso panelje teherereosen nyomott " & ChrW(8211) & " polirozas szukseges."
    ws.Rows("19:19").RowHeight = 28

    ' --- 5. MONDATVAZ ---
    ws.Rows("21:21").RowHeight = 10
    ws.Range("B22").Value = "Ajanlott mondatvaz (kovetendo sablon)"
    ws.Rows("22:22").RowHeight = 22
    ws.Range("B23").Value = "[Hol keszult], [hol lett eszreveve], [mi a hiba], [pontos hely], [meret/mertek], [teendo]."
    ws.Rows("23:23").RowHeight = 20

    ' --- 6. SAJAT VAZLAT ---
    ws.Rows("25:25").RowHeight = 10
    ws.Range("B26").Value = "Sajat vazlat es szovegezes"
    ws.Rows("26:26").RowHeight = 22

    ws.Range("B27").Value = "Vazlatpontok (kezzel toltsd ki):"
    ws.Rows("27:27").RowHeight = 18
    ' B28:D31 - vazlatpontok: felhasznaloi terulet, nem toroljuk

    ws.Range("B33").Value = "Vegleges sajat leiras (kezzel):"
    ws.Rows("33:33").RowHeight = 18
    ' B34:D38 - vegleges szoveg: felhasznaloi terulet, nem toroljuk

    Dim rr As Long
    For rr = 28 To 38
        ws.Rows(rr).RowHeight = 20
    Next rr

    ' --- 7. ELLENORZOLISTA ---
    ws.Rows("40:40").RowHeight = 10
    ws.Range("B41").Value = "Ellenorzolista a leiras elott"
    ws.Rows("41:41").RowHeight = 22

    Dim ellenorzok(1 To 4) As String
    ellenorzok(1) = "[ ]  Szerepel a pontos helyszin (sor, allomas, elem)?"
    ellenorzok(2) = "[ ]  Szerepel, hol lett eszreveve (RK, gyartas, stb.)?"
    ellenorzok(3) = "[ ]  Egyertelmu a hiba tipusa (nem csak altalanos)?"
    ellenorzok(4) = "[ ]  Szerepel a teendo vagy kovetkezo lepes?"
    For r = 1 To 4
        ws.Range("B" & (41 + r)).Value = ellenorzok(r)
        ws.Rows(41 + r).RowHeight = 18
    Next r
End Sub

' ------------------------------------------------------------
' FORMAZAS
' ------------------------------------------------------------
Private Sub FormatHelperSheet(ByVal ws As Worksheet)
    ' --- Focim ---
    With ws.Range("B1")
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(68, 114, 196)
        .VerticalAlignment = xlCenter
    End With

    ' --- Leiras ---
    With ws.Range("B2:D3")
        .Font.Size = 10
        .Font.Italic = True
        .WrapText = True
    End With

    ' --- Kitoltesi sorrend - alcim ---
    With ws.Range("B5")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(68, 114, 196)
    End With

    ' Tablafejlec
    With ws.Range("A6:D6")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Adatsorok - valtakozo hatter
    Dim r As Long
    For r = 7 To 12
        With ws.Range("A" & r & ":D" & r)
            .Font.Size = 10
            .WrapText = True
            .VerticalAlignment = xlCenter
            If r Mod 2 = 1 Then
                .Interior.Color = RGB(255, 255, 255)
            Else
                .Interior.Color = RGB(242, 242, 242)
            End If
        End With
        ws.Range("A" & r).HorizontalAlignment = xlCenter
        ws.Range("B" & r).Font.Bold = True
    Next r

    ws.Range("A6:D12").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    With ws.Range("A6:D12").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
    With ws.Range("A6:D12").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With

    ' --- Peldak - alcim ---
    With ws.Range("B15")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(68, 114, 196)
    End With

    ' "Nem jo" sorok
    With ws.Range("B16:D17")
        .Interior.Color = RGB(255, 235, 220)
        .Font.Size = 10
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B16").Font.Bold = True
    ws.Range("B17").Font.Bold = True

    ' "Jo" sorok
    With ws.Range("B18:D19")
        .Interior.Color = RGB(220, 240, 220)
        .Font.Size = 10
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B18").Font.Bold = True
    ws.Range("B19").Font.Bold = True

    ws.Range("B16:D19").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    With ws.Range("B16:D19").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With

    ' --- Mondatvaz - alcim ---
    With ws.Range("B22")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(68, 114, 196)
    End With
    With ws.Range("B23:D23")
        .Interior.Color = RGB(255, 255, 204)
        .Font.Size = 10
        .Font.Italic = True
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B23:D23").BorderAround LineStyle:=xlContinuous, Weight:=xlThin

    ' --- Sajat vazlat - alcim ---
    With ws.Range("B26")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(68, 114, 196)
    End With
    With ws.Range("B27")
        .Font.Bold = True
        .Font.Size = 10
    End With

    ' Vazlatpontok beviteli terulet
    With ws.Range("B28:D31")
        .Interior.Color = RGB(235, 241, 255)
        .Font.Size = 10
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Range("B28:D31").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

    With ws.Range("B33")
        .Font.Bold = True
        .Font.Size = 10
    End With

    ' Vegleges szoveg beviteli terulet
    With ws.Range("B34:D38")
        .Interior.Color = RGB(235, 241, 255)
        .Font.Size = 11
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Range("B34:D38").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

    ' --- Ellenorzolista - alcim ---
    With ws.Range("B41")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(68, 114, 196)
    End With
    With ws.Range("B42:D45")
        .Font.Size = 10
        .Interior.Color = RGB(242, 242, 242)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B42:D45").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    With ws.Range("B42:D45").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With

    ' A oszlop - szurke margo
    ws.Columns("A").Interior.Color = RGB(242, 242, 242)

    ' Munkalap fultabla szin
    ws.Tab.Color = RGB(68, 114, 196)
End Sub
