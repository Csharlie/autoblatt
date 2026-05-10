Attribute VB_Name = "buttonsInstaller"
Option Explicit

' ============================================================
' AUTOBLATT - WORKSHEET BUTTON UI (RIBBON HELYETTESITO)
' ============================================================
' Letrehoz egy "AB Indito" nevu munkalapot a megnyitott
' munkafuzetben, kattinhato Form gombokkal, amik az AB_*
' makrokat hivjak. Ez a primer UI, mivel ceges policy nem engedi
' a custom ribbont.
'
' Hasznalat: Alt+F8 -> AB_InstallButtons -> Run
' (Egyszer kell futtatni, vagy ha a gombok elveszettek/torolve.)
' ============================================================

' --- Publikus belepesi pont ---
Public Sub AB_InstallButtons()
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Nincs aktiv munkafuzet.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(wb, "AB Indito")

    ' Tartalom torlese, alapformazas
    ws.Cells.Clear
    ws.Cells.RowHeight = 18
    ws.Columns("A:E").ColumnWidth = 22
    ws.Tab.Color = RGB(68, 114, 196)

    ' Cim
    With ws.Range("A1:E1")
        .Merge
        .Value = APP_NAME & " v" & APP_VERSION & " - Inditopanel"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With

    ' Leiras
    With ws.Range("A3:E3")
        .Merge
        .Value = "Kattints egy gombra a megfelelo makro futtatasahoz. (Alternativ: Alt+F8 -> AB_*)"
        .Font.Italic = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Gombok torlese
    Call DeleteShapesByPrefix(ws, "AB_Btn_")

    ' Gombok elhelyezese 2 oszlopban
    Const BTN_TOP_START As Double = 80
    Const BTN_LEFT_COL1 As Double = 30
    Const BTN_LEFT_COL2 As Double = 250
    Const BTN_WIDTH    As Double = 200
    Const BTN_HEIGHT   As Double = 36
    Const BTN_VSPACE   As Double = 14

    Dim defs As Variant
    defs = Array( _
        Array("Drive Mentes", "AB_SaveToDrive"), _
        Array("Kep Import", "AB_ImportImages"), _
        Array("Email Kuldes", "AB_SendEmails"), _
        Array("Fill Panel", "AB_FillPanel"), _
        Array("Helpre Panel", "AB_HelprePanel"), _
        Array("Beallitasok", "AB_Settings"), _
        Array("Sugo", "AB_Help"), _
        Array("Modulok listazasa", "AB_ListAllModules") _
    )

    Dim i As Long
    For i = LBound(defs) To UBound(defs)
        Dim col As Long, rowIdx As Long
        col = i Mod 2
        rowIdx = i \ 2
        Dim leftPos As Double, topPos As Double
        If col = 0 Then leftPos = BTN_LEFT_COL1 Else leftPos = BTN_LEFT_COL2
        topPos = BTN_TOP_START + rowIdx * (BTN_HEIGHT + BTN_VSPACE)

        Call AddButton(ws, defs(i)(0), defs(i)(1), leftPos, topPos, BTN_WIDTH, BTN_HEIGHT)
    Next i

    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Az 'AB Indito' lap elkeszult, a gombok hasznalhatok.", vbInformation, APP_NAME
End Sub

' --- Egy gomb hozzaadasa, makrohoz csatolva ---
Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, _
                      ByVal leftPos As Double, ByVal topPos As Double, _
                      ByVal width As Double, ByVal height As Double)
    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPos, topPos, width, height)
    btn.Name = "AB_Btn_" & macroName
    btn.Caption = caption
    btn.OnAction = macroName
    btn.Font.Size = 11
    btn.Font.Bold = True
End Sub

' Shape torles prefix alapjan (ismetli az image-importer-beli logikat,
' de nem fugg attol a moduletol mert a buttons-installer onaltallhat).
Private Sub DeleteShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim shp As Shape
    Dim toDelete As Collection
    Set toDelete = New Collection
    For Each shp In ws.Shapes
        If InStr(1, shp.Name, prefix) > 0 Then toDelete.Add shp
    Next shp
    Dim i As Long
    For i = 1 To toDelete.Count
        On Error Resume Next
        toDelete(i).Delete
        On Error GoTo 0
    Next i
End Sub
