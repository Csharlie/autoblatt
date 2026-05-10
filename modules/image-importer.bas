Attribute VB_Name = "imageImporter"
Option Explicit

' ============================================================
' AUTOBLATT - KEPIMPORTALO
' ============================================================
' Beolvas .jpg/.jpeg/.png kepeket az aktiv munkafuzet mappajabol,
' es elhelyezi oket egy 2 oszlopos rácsban az aktiv lapon az
' IMAGE_START_CELL alapjan. EXIF orientation-bol detektalja
' az allo/fekvo tajolast (WIA COM-on keresztul).
'
' Felbontasfuggetlen pozicionalas: a startCell tenyleges Top/Left
' ertekeit hasznaljuk, nem onkenyes pixel-kuszoboket.
' ============================================================

' --- Layout konstansok (cm-ben) ---
Private Const COL_LEFT_OFFSET_CM    As Double = 0.05    ' Bal oszlop X-eltolas
Private Const COL_RIGHT_OFFSET_CM   As Double = 9.6     ' Jobb oszlop X-eltolas
Private Const ROW_STEP_CM           As Double = 5.38    ' Sortavolsag (top->top)
Private Const IMAGE_WIDTH_CM        As Double = 9.5     ' Kep szelesseg fekvo eseten
Private Const IMAGE_HEIGHT_PORTRAIT_CM As Double = 4.75 ' Kep magassag allo eseten

' --- Publikus belepesi pont ---
Public Sub RunImageImporter()
    Call ImportImages
End Sub

' --- Diagnosztika: aktiv kezdocella poziciojanak megjelenitese ---
' Hasznald ha a kepek nem oda kerulnek ahova varod, hogy lasd a startCell
' aktualis Top/Left erteket az adott monitor felbontason.
Public Sub AB_ImportImagesDebugPosition()
    Dim startCell As Range
    Set startCell = ActiveSheet.Range(IMAGE_START_CELL)
    MsgBox IMAGE_START_CELL & " cella poziciója:" & vbCrLf & _
           "Left: " & Format(startCell.Left, "0.00") & " pt (" & Format(startCell.Left / 28.35, "0.00") & " cm)" & vbCrLf & _
           "Top:  " & Format(startCell.Top, "0.00") & " pt (" & Format(startCell.Top / 28.35, "0.00") & " cm)" & vbCrLf & _
           "Width:  " & Format(startCell.Width, "0.00") & " pt" & vbCrLf & _
           "Height: " & Format(startCell.Height, "0.00") & " pt" & vbCrLf & _
           "Application.UsableHeight: " & Application.UsableHeight & " pt", _
           vbInformation, APP_NAME
End Sub

' --- Karbantartasi segedmakro: csak az 'Import_' nevu kepek torlese ---
Public Sub AB_CleanImportImages()
    Dim count As Long
    count = DeleteShapesByPrefix(ActiveSheet, "Import_")
    MsgBox count & " db Import_ kep torolve.", vbInformation, APP_NAME
End Sub

' ------------------------------------------------------------
' FOORCHESTRATOR
' ------------------------------------------------------------
Private Sub ImportImages()
    On Error GoTo Fail

    ' 1. Kepek torlese (kiveve SHAPE_KEEP - pl. logo, sablon)
    Call DeleteAllShapesExcept(ActiveSheet, SHAPE_KEEP)

    ' 2. Kepek importalasa a dokumentum mappajabol
    Call ImportImagesFromDocumentFolder

    MsgBox "Kepimportalas befejezve!", vbInformation, APP_NAME
    Exit Sub

Fail:
    MsgBox "Hiba a kepimportalas soran: " & Err.Description, vbCritical, APP_NAME
    WriteLog "ERROR ImportImages: " & Err.Description
End Sub

' ------------------------------------------------------------
' SHAPE TAKARITAS
' ------------------------------------------------------------
' Minden shape torlese az adott lapon, kiveve azokat amelyek neve egyezik keepName-mel.
Private Sub DeleteAllShapesExcept(ByVal ws As Worksheet, ByVal keepName As String)
    Dim shp As Shape
    Dim toDelete As Collection
    Set toDelete = New Collection

    For Each shp In ws.Shapes
        If shp.Name <> keepName Then toDelete.Add shp
    Next shp

    Dim i As Long
    For i = 1 To toDelete.Count
        On Error Resume Next
        toDelete(i).Delete
        On Error GoTo 0
    Next i
End Sub

' Csak az adott prefix-szel kezdodo nevu shape-ek torlese. Visszaadja a torolt darabszamot.
Private Function DeleteShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String) As Long
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

    DeleteShapesByPrefix = toDelete.Count
End Function

' ------------------------------------------------------------
' KEPIMPORTALAS
' ------------------------------------------------------------
Private Sub ImportImagesFromDocumentFolder()
    Dim folderPath As String
    folderPath = ResolveLocalImageFolder()

    If Len(folderPath) = 0 Then
        MsgBox "Nem talalhato helyi mappa! Gyozodj meg rola, hogy az Excel fajl olyan mappaban van, ami el van erve helyileg.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    Dim startCell As Range
    Set startCell = ActiveSheet.Range(IMAGE_START_CELL)

    Dim files As Collection
    Set files = GetImageFiles(folderPath)

    If files.Count = 0 Then
        MsgBox "Nem talalhatok kepfajlok a mappaban:" & vbCrLf & folderPath, _
               vbInformation, APP_NAME
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To files.Count
        Call PlaceImageInGrid(CStr(files(i)), startCell, i)
    Next i
End Sub

' Visszaadja az aktuális munkafuzet helyi mappajat (kepek forrasa).
' SharePoint/OneDrive URL eseten visszafejti az utolso mappa nevet,
' es a OneDrive helyi mappabol epiti az utvonalat.
Private Function ResolveLocalImageFolder() As String
    Dim wbPath As String
    wbPath = ActiveWorkbook.Path

    ' Helyi utvonal eseten azonnal hasznalhato
    If Len(wbPath) > 0 And InStr(wbPath, "https://") = 0 Then
        ResolveLocalImageFolder = wbPath & "\"
        Exit Function
    End If

    ' SharePoint URL eseten: az utolso szegmenst kivesszuk
    Dim folderName As String
    Dim parts() As String
    parts = Split(wbPath, "/")
    If UBound(parts) >= 0 Then folderName = parts(UBound(parts))
    folderName = UrlDecode(folderName)

    If Len(folderName) = 0 Then
        ResolveLocalImageFolder = ""
        Exit Function
    End If

    Dim onedrive As String
    onedrive = GetOneDrivePath()
    If Len(onedrive) = 0 Then
        ResolveLocalImageFolder = ""
        Exit Function
    End If

    ' Mentett fajlok a OneDrive\Sheets\<ev>\<folderName> alatt vannak
    Dim yearFolder As String
    yearFolder = CStr(Year(Date))
    ResolveLocalImageFolder = onedrive & "\" & SAVE_ROOT_FOLDER & "\" & yearFolder & "\" & folderName & "\"
End Function

' Osszegyujti a kepfajlok abszolut utvonalait egy collection-be.
Private Function GetImageFiles(ByVal folderPath As String) As Collection
    Dim files As New Collection
    Dim fname As String
    Dim ext As Variant

    On Error GoTo Fail
    For Each ext In Array("*.jpg", "*.jpeg", "*.png")
        fname = Dir(folderPath & ext)
        Do While Len(fname) > 0
            files.Add folderPath & fname
            fname = Dir
        Loop
    Next ext

    Set GetImageFiles = files
    Exit Function

Fail:
    MsgBox "Hiba a kepfajlok keresese soran: " & Err.Description & vbCrLf & "Mappa: " & folderPath, _
           vbCritical, APP_NAME
    Set GetImageFiles = files
End Function

' ------------------------------------------------------------
' KEP ELHELYEZES (felbontasfuggetlen)
' ------------------------------------------------------------
' A startCell tenyleges Top/Left erteket hasznalja kiindulasként,
' igy automatikusan idomul a sor magassagahoz, az ablak meretehez.
Private Sub PlaceImageInGrid(ByVal imagePath As String, _
                              ByVal startCell As Range, _
                              ByVal imageIndex As Long)
    Dim newShape As Shape
    Dim gridRow As Long, gridCol As Long
    gridRow = (imageIndex - 1) \ 2
    gridCol = (imageIndex - 1) Mod 2

    ' Pozicio a startCell-bol szamitva (felbontasfuggetlen)
    Dim leftPt As Double, topPt As Double
    Dim colOffsetCm As Double
    If gridCol = 0 Then
        colOffsetCm = COL_LEFT_OFFSET_CM
    Else
        colOffsetCm = COL_RIGHT_OFFSET_CM
    End If

    leftPt = startCell.Left + Application.CentimetersToPoints(colOffsetCm)
    topPt = startCell.Top + Application.CentimetersToPoints(gridRow * ROW_STEP_CM)

    ' Kep beszurasa
    On Error GoTo InsertFail
    Set newShape = ActiveSheet.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=100, Top:=100, Width:=-1, Height:=-1)
    newShape.Placement = xlFreeFloating
    newShape.LockAspectRatio = msoTrue

    If IsImagePortrait(imagePath) Then
        newShape.Height = Application.CentimetersToPoints(IMAGE_HEIGHT_PORTRAIT_CM)
    Else
        newShape.Width = Application.CentimetersToPoints(IMAGE_WIDTH_CM)
    End If

    Application.ScreenUpdating = False
    newShape.Left = leftPt
    newShape.Top = topPt
    Application.ScreenUpdating = True

    newShape.Name = "Import_" & GetFileNameWithoutExtension(imagePath)
    Exit Sub

InsertFail:
    MsgBox "Nem sikerult betolteni a kepet:" & vbCrLf & imagePath & vbCrLf & _
           "Hiba: " & Err.Description, vbExclamation, APP_NAME
    WriteLog "ERROR PlaceImageInGrid: " & imagePath & " - " & Err.Description
End Sub

' --- Fajlnev kiterjesztes nelkul ---
Private Function GetFileNameWithoutExtension(ByVal fullPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    fileName = Dir(fullPath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetFileNameWithoutExtension = Left(fileName, dotPos - 1)
    Else
        GetFileNameWithoutExtension = fileName
    End If
End Function

' --- Allo (portrait) kep detektalasa WIA-val EXIF Orientation alapjan ---
Private Function IsImagePortrait(ByVal imagePath As String) As Boolean
    On Error GoTo Fail
    Dim img As Object
    Dim w As Long, h As Long, orient As Long

    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile imagePath
    w = img.Width
    h = img.Height

    orient = 1
    On Error Resume Next
    orient = img.Properties("274").Value
    On Error GoTo Fail

    Select Case orient
        Case 1, 2, 3, 4
            IsImagePortrait = (h > w)
        Case 5, 6, 7, 8
            IsImagePortrait = (w > h)
        Case Else
            IsImagePortrait = (h > w)
    End Select
    Exit Function

Fail:
    IsImagePortrait = False
End Function
