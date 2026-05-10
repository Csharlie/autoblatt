Attribute VB_Name = "fillPanel"
Option Explicit

' ============================================================
' AUTOBLATT - FOTOLAP KITOLTOPANEL
' ============================================================
' Letrehoz vagy frissit egy 'Fill Panel' nevu munkalapot egy
' 6 mezos beviteli urlappal es egy 5x5 osszevont osszefoglalo
' blokkal, amely a beirt ertekekbol epit egy mondatot.
'
' A SaveToDrive nem hivja - kulon makro (AB_FillPanel).
' ============================================================

' --- Publikus belepesi pont ---
Public Sub RunFillPanel()
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Nincs aktiv munkafuzet.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, SHEET_FILLPANEL)

    Call BuildEntryPanel(ws)

    Application.ScreenUpdating = True
    ws.Activate
End Sub

' --- Panel epites: bal urlap + jobb osszefoglalo ---
Private Sub BuildEntryPanel(ByVal ws As Worksheet)
    Const baseRow As Long = 2
    Const baseCol As Long = 2

    ' A oszlop: szuk margo
    ws.Columns(1).ColumnWidth = 2.5

    ' Fejlec - osszevont B-C cella
    On Error Resume Next
    ws.Range(ws.Cells(baseRow, baseCol), ws.Cells(baseRow, baseCol + 1)).Merge
    On Error GoTo 0
    ws.Cells(baseRow, baseCol).Value = "Fotolap kitoltopanel"

    ' Cimkek
    ws.Cells(baseRow + 1, baseCol).Value = "Mikor keszult?"
    ws.Cells(baseRow + 2, baseCol).Value = "Ki vette eszre?"
    ws.Cells(baseRow + 3, baseCol).Value = "Zarolas tortent?"
    ws.Cells(baseRow + 4, baseCol).Value = "Mennyiseg"
    ws.Cells(baseRow + 5, baseCol).Value = "Egyseg"

    ' Lenyilo listak
    ApplyDropdown ws.Cells(baseRow + 1, baseCol + 1), _
        "Szeriajovahagyo,Gyartas elejen,Gyartas kozepen,Gyartas vegen,Muszakatadas elott,Muszakatadas utan,Koteg elejen,Koteg vegen,Kotegcsere utan"
    ApplyDropdown ws.Cells(baseRow + 2, baseCol + 1), "RK1,RK2,NA"
    ApplyDropdown ws.Cells(baseRow + 3, baseCol + 1), "Igen,Nem"
    ApplyDropdown ws.Cells(baseRow + 5, baseCol + 1), "db,LT"

    ' Formazas
    Call FormatEntryPanel(ws, baseRow, baseCol)

    ' Osszefoglalo blokk: baseCol + 3-tol (egy ures res utan)
    Call BuildSummaryBlock(ws, baseRow, baseCol + 3)

    ' Nevvel ellatott tartomany - "FillPanel" (idempotens)
    Dim panelAddr As String
    panelAddr = ws.Range(ws.Cells(baseRow, baseCol), _
                         ws.Cells(baseRow + 5, baseCol + 1)).Address(External:=True)
    On Error Resume Next
    ws.Parent.Names("FillPanel").Delete
    On Error GoTo 0
    ws.Parent.Names.Add Name:="FillPanel", RefersTo:="=" & panelAddr
End Sub

' --- Kitoltopanel formazas ---
Private Sub FormatEntryPanel(ByVal ws As Worksheet, ByVal baseRow As Long, ByVal baseCol As Long)
    Dim panelRange As Range
    Set panelRange = ws.Range(ws.Cells(baseRow, baseCol), _
                              ws.Cells(baseRow + 5, baseCol + 1))

    ' Fejlec sor
    With ws.Range(ws.Cells(baseRow, baseCol), ws.Cells(baseRow, baseCol + 1))
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

    ' Mezo sorok 1-5
    Dim r As Long
    For r = 1 To 5
        ' Cimke oszlop
        With ws.Cells(baseRow + r, baseCol)
            .Font.Bold = True
            .Font.Size = 10
            .Interior.Color = RGB(242, 242, 242)
            .VerticalAlignment = xlCenter
            .RowHeight = 20
        End With
        ' Ertek cella
        With ws.Cells(baseRow + r, baseCol + 1)
            .Font.Size = 10
            .VerticalAlignment = xlCenter
            ' Mennyiseg (r=4) es Egyseg (r=5): sarga = csak zarolas eseten relevans
            If r >= 4 Then
                .Interior.Color = RGB(255, 255, 204)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If
        End With
    Next r

    ' Belso racsvonalak
    With panelRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
    With panelRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
    panelRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

    ws.Columns(baseCol).ColumnWidth = 22
    ws.Columns(baseCol + 1).ColumnWidth = 16
End Sub

' --- Osszefoglalo blokk: 5 oszlop x 6 sor merged + automatikus mondat ---
Private Sub BuildSummaryBlock(ByVal ws As Worksheet, ByVal baseRow As Long, ByVal sumCol As Long)
    ' Fejlec: 5 cella merged
    On Error Resume Next
    ws.Range(ws.Cells(baseRow, sumCol), ws.Cells(baseRow, sumCol + 4)).Merge
    On Error GoTo 0
    With ws.Cells(baseRow, sumCol)
        .Value = "Osszefoglalas"
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

    ' Szoveg blokk: 5 col x 5 sor merged
    On Error Resume Next
    ws.Range(ws.Cells(baseRow + 1, sumCol), _
             ws.Cells(baseRow + 5, sumCol + 4)).Merge
    On Error GoTo 0
    With ws.Range(ws.Cells(baseRow + 1, sumCol), _
                  ws.Cells(baseRow + 5, sumCol + 4))
        .Font.Size = 11
        .Font.Italic = True
        .Interior.Color = RGB(235, 241, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    ' Cellacimek a kitolto urlapon (3. oszlopban vannak az ertekek)
    Dim addrWhen As String, addrWho As String, addrLock As String
    Dim addrQty As String, addrUnit As String
    addrWhen = ws.Cells(baseRow + 1, 3).Address
    addrWho = ws.Cells(baseRow + 2, 3).Address
    addrLock = ws.Cells(baseRow + 3, 3).Address
    addrQty = ws.Cells(baseRow + 4, 3).Address
    addrUnit = ws.Cells(baseRow + 5, 3).Address

    ' Mondatgenerator formula
    Dim f As String
    f = "=IF(AND(" & addrWhen & "=""""," & addrWho & "=""""),""""," & _
        "TRIM(" & _
            "IF(" & addrWhen & "=""Szeriajovahagyo""," & _
                """szeriajovahagyo alkatresz ellenorzese soran""," & _
                "IF(" & addrWhen & "="""",""""," & addrWhen & "&"" "")&" & _
                "IF(" & addrWho & "="""",""""," & _
                    "IF(" & addrWho & "=""NA"",""nem azonositott""," & addrWho & "&""-es"")&"" ellenorzes soran""))&" & _
            "IF(" & addrLock & "=""Igen"","", zarolas: ""&" & _
                "IF(" & addrQty & "="""",""?""," & addrQty & "&"" ""&" & addrUnit & "),"""")" & _
        "))"
    ws.Cells(baseRow + 1, sumCol).Formula = f

    ' Keret
    ws.Range(ws.Cells(baseRow, sumCol), _
             ws.Cells(baseRow + 5, sumCol + 4)).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium

    Dim c As Long
    For c = sumCol To sumCol + 4
        ws.Columns(c).ColumnWidth = 12
    Next c
End Sub
