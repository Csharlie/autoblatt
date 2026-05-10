Attribute VB_Name = "driveManager"
Option Explicit

' ============================================================
' AUTOBLATT - DRIVE MANAGER
' ============================================================
' Felelosseg: szemelyzeti adatok lekerese a data.xlsm-bol,
' Fotoblatt cellak feltoltese es formazasa, Data lap masolas,
' szovegnormalizalas, verziozott mentes a OneDrive Sheets mappaba.
'
' NEM hivja a Fill Panelt es a Helpre Panelt - ezek kulon
' makrok (AB_FillPanel, AB_HelprePanel).
' ============================================================

' --- Publikus belepesi pont ---
Public Sub RunDriveManager()
    Call SaveToDrive
End Sub

' --- Foorchestrator ---
Private Sub SaveToDrive()
    Dim activeWb As Workbook
    Dim activeWs As Worksheet

    Set activeWb = Application.ActiveWorkbook
    Set activeWs = activeWb.ActiveSheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo Fail

    ' 1. Szemelyzeti adatok lekerese es kitoltese
    If Not FetchPersonalData(activeWb, activeWs) Then GoTo Cleanup

    ' 2. Data lap feltoltese (a forrasbol masolas)
    Call PopulateDataSheet(activeWb)

    ' 3. Fotoblatt formazas
    Call ApplyFotoblattFormatting(activeWs)

    ' 4. Szovegnormalizalas (REPLACE_RULES alapjan)
    Call NormalizeWorkbookText(activeWb)

    ' 5. Fotoblatt aktivalas + nezet beallitas
    Call ActivateFotoblatt(activeWb)

    ' 6. Mentes verziozott utvonalra
    If Not SaveVersionedWorkbook(activeWb, activeWs) Then GoTo Cleanup

    MsgBox "Az adatok sikeresen mentve!", vbInformation, APP_NAME
    GoTo Cleanup

Fail:
    MsgBox "Hiba tortent: " & Err.Description, vbCritical, APP_NAME
    WriteLog "ERROR SaveToDrive: " & Err.Description

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' ------------------------------------------------------------
' 1. SZEMELYZETI ADATOK
' ------------------------------------------------------------
Private Function FetchPersonalData(ByVal activeWb As Workbook, _
                                    ByVal activeWs As Worksheet) As Boolean
    Dim sourcePath As String
    sourcePath = ResolveDataXlsmPath()

    If Dir(sourcePath) = "" Then
        MsgBox "A 'data.xlsm' nem talalhato a varhato helyen:" & vbCrLf & sourcePath, _
               vbExclamation, APP_NAME
        FetchPersonalData = False
        Exit Function
    End If

    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(sourcePath, ReadOnly:=True)

    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = wbSource.Sheets(SHEET_PERSONAL)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "A '" & SHEET_PERSONAL & "' munkalap nem talalhato a data.xlsm-ben.", _
               vbExclamation, APP_NAME
        wbSource.Close False
        FetchPersonalData = False
        Exit Function
    End If

    Dim personalTable As ListObject
    On Error Resume Next
    Set personalTable = wsSource.ListObjects(TABLE_PERSONAL)
    On Error GoTo 0
    If personalTable Is Nothing Then
        MsgBox "A '" & TABLE_PERSONAL & "' tabla nem talalhato.", vbExclamation, APP_NAME
        wbSource.Close False
        FetchPersonalData = False
        Exit Function
    End If

    Dim idCol As Range, nameCol As Range, shiftCol As Range
    Set idCol = personalTable.ListColumns(COL_PERSONAL_ID).DataBodyRange
    Set nameCol = personalTable.ListColumns(COL_PERSONAL_NAME).DataBodyRange
    Set shiftCol = personalTable.ListColumns(COL_PERSONAL_SHIFT).DataBodyRange

    Dim userName As String
    userName = Environ("USERNAME")
    activeWs.Range(CELL_USERNAME).Value = userName

    Dim foundCell As Range
    Set foundCell = idCol.Find(What:=userName, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        activeWb.Activate
        activeWs.Range(CELL_FULLNAME).Value = foundCell.Offset(0, nameCol.Column - idCol.Column).Value
        activeWs.Range(CELL_SHIFT).Value = foundCell.Offset(0, shiftCol.Column - idCol.Column).Value
        activeWs.Range(CELL_DATE).Value = Date
    End If

    ' A masolashoz a tablat globalisan elerhetove tesszuk a kovetkezo lepesen at:
    ' a PopulateDataSheet ujra megnyitja, vagy itt elotte masolunk. Egyszerubb itt:
    Call CopyPersonalTableToDataSheet(activeWb, personalTable)

    wbSource.Close False
    FetchPersonalData = True
End Function

' Visszaadja a data.xlsm utvonalat: elobb a config-bol, kulonben a beepitett app/data/ mappabol.
Private Function ResolveDataXlsmPath() As String
    Dim configured As String
    configured = GetConfigValue("dataXlsmPath", "")
    If Len(configured) > 0 And Dir(configured) <> "" Then
        ResolveDataXlsmPath = configured
        Exit Function
    End If

    Dim local As String
    local = GetAppRootPath() & "\data\" & FALLBACK_DATA_FILENAME
    If Dir(local) <> "" Then
        ResolveDataXlsmPath = local
        Exit Function
    End If

    ' Vegso visszaeses: OneDrive gyokere
    ResolveDataXlsmPath = GetOneDrivePath() & "\" & FALLBACK_DATA_FILENAME
End Function

' ------------------------------------------------------------
' 2. DATA LAP FELTOLTES
' ------------------------------------------------------------
Private Sub PopulateDataSheet(ByVal activeWb As Workbook)
    ' A Data lapot a CopyPersonalTableToDataSheet mar feltoltotte.
    ' Itt csak biztositjuk hogy letezzen es az aktiv lap a Fotoblatt maradjon.
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(activeWb, SHEET_DATA)
    Application.CutCopyMode = False
End Sub

Private Sub CopyPersonalTableToDataSheet(ByVal activeWb As Workbook, _
                                          ByVal personalTable As ListObject)
    Dim dataSheet As Worksheet
    Set dataSheet = GetOrCreateWorksheet(activeWb, SHEET_DATA)
    dataSheet.UsedRange.ClearContents

    personalTable.Range.Copy
    activeWb.Activate
    Dim dest As Range
    Set dest = dataSheet.Range("A1")
    dest.PasteSpecial xlPasteValues
    dest.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' ------------------------------------------------------------
' 3. FOTOBLATT FORMAZAS
' ------------------------------------------------------------
Private Sub ApplyFotoblattFormatting(ByVal ws As Worksheet)
    ' Felhasznalonev cella - kis betuvel, vastagon, kozeppe, nagybetus
    With ws.Range(CELL_USERNAME)
        .Value = UCase(.Value)
        .Font.Size = 9
        .Font.Bold = True
        .Font.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Kozos formazasok (10pt, kozeppe)
    With ws.Range(CELL_FULLNAME & "," & CELL_SHIFT & "," & CELL_DATE & "," & _
                  CELL_SUBJECT_1 & "," & CELL_SUBJECT_2 & "," & CELL_SUBJECT_3 & "," & _
                  CELL_HIGHLIGHT & "," & CELL_SUBJECT_4)
        .Font.Size = 10
        .Font.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Vastag betutipus
    With ws.Range(CELL_FULLNAME & "," & CELL_SUBJECT_1 & "," & CELL_SUBJECT_2 & "," & CELL_SUBJECT_4)
        .Font.Bold = True
    End With

    ' Sima betutipus
    With ws.Range(CELL_SHIFT & "," & CELL_DATE & "," & CELL_SUBJECT_3 & "," & CELL_HIGHLIGHT)
        .Font.Bold = False
    End With

    ' Piros kiemeles
    With ws.Range(CELL_SUBJECT_3 & "," & CELL_HIGHLIGHT)
        .Font.Color = vbRed
    End With

    ' G8 cella elso betuje legyen nagy
    With ws.Range(CELL_SUBJECT_4)
        If Len(.Value) > 0 Then
            .Value = UCase(Left(.Value, 1)) & Mid(.Value, 2)
        End If
    End With

    ' Datum jobbra igazitasa
    ws.Range(CELL_DATE).HorizontalAlignment = xlRight

    ' A9 footer cella
    With ws.Range(CELL_FOOTER)
        .Font.Size = 10
        .Font.Bold = False
    End With

    ' C15 cella torlese (ha nem ures, kezeli az osszevont cellakat is)
    With ws.Range(CELL_NOTE)
        If .MergeCells Then
            .MergeArea.ClearContents
        ElseIf Not IsEmpty(.Value) Then
            .ClearContents
        End If
    End With

    ' Object 1 nevu shape torlese ha letezik
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("Object 1")
    If Not shp Is Nothing Then shp.Delete
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' 4. SZOVEG NORMALIZALAS
' ------------------------------------------------------------
Private Sub NormalizeWorkbookText(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim cell As Range
    For Each ws In wb.Worksheets
        For Each cell In ws.UsedRange
            If VarType(cell.Value) = vbString Then
                Dim newVal As String
                newVal = ApplyReplaceRules(CStr(cell.Value), REPLACE_RULES)
                If newVal <> cell.Value Then cell.Value = newVal
            End If
        Next cell
    Next ws
End Sub

' ------------------------------------------------------------
' 5. FOTOBLATT AKTIVALAS
' ------------------------------------------------------------
Private Sub ActivateFotoblatt(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(SHEET_FOTOBLATT)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "A '" & SHEET_FOTOBLATT & "' munkalap nem talalhato.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ws.Activate
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
End Sub

' ------------------------------------------------------------
' 6. VERZIOZOTT MENTES
' ------------------------------------------------------------
Private Function SaveVersionedWorkbook(ByVal activeWb As Workbook, _
                                        ByVal activeWs As Worksheet) As Boolean
    Dim onedrivePath As String
    onedrivePath = GetOneDrivePath()

    ' Mappa nev epitese: A6 (szokoz nelkul) + E6 + G8
    Dim folderA6 As String, folderE6 As String, folderG8 As String
    folderA6 = SanitizeForFilename(Replace(CStr(activeWs.Range(CELL_SUBJECT_1).Value), " ", ""))
    folderE6 = SanitizeForFilename(CStr(activeWs.Range(CELL_SUBJECT_3).Value))
    folderG8 = SanitizeForFilename(CapitalizeWords(CStr(activeWs.Range(CELL_SUBJECT_4).Value)))
    Dim folderName As String
    folderName = Trim(folderA6 & " " & folderE6 & " " & folderG8)

    ' Fajlnev epitese: I3 (datum yyyymmdd) + C6 + E6 + G8 + (PS-foto|PS-info)
    Dim partC6 As String, partE6 As String, partG8 As String
    partC6 = SanitizeForFilename(CStr(activeWs.Range(CELL_SUBJECT_2).Value))
    partE6 = SanitizeForFilename(CStr(activeWs.Range(CELL_SUBJECT_3).Value))
    partG8 = SanitizeForFilename(CStr(activeWs.Range(CELL_SUBJECT_4).Value))

    Dim suffix As String
    Dim sheetType As String
    sheetType = LCase(CStr(activeWs.Range(CELL_SHEETTYPE).Value))
    If InStr(sheetType, "fotoblatt") > 0 Then
        suffix = "foto"
    Else
        suffix = "info"
    End If

    Dim datePart As String
    datePart = Format(activeWs.Range(CELL_DATE).Value, "yyyymmdd")

    Dim fileName As String
    fileName = datePart & " " & partC6 & " " & partE6 & " " & partG8 & " (PS-" & suffix & ")"

    ' Teljes utvonal: OneDrive\Sheets\<ev>\<folderName>
    Dim yearFolder As String
    yearFolder = CStr(Year(Date))
    Dim fullFolder As String
    fullFolder = onedrivePath & "\" & SAVE_ROOT_FOLDER & "\" & yearFolder & "\" & folderName

    EnsureFolder fullFolder

    ' Verziozas
    Dim baseName As String, fullPath As String
    Dim versionNum As Integer
    baseName = fileName
    versionNum = 1
    fullPath = fullFolder & "\" & baseName & ".xlsm"

    If Dir(fullPath) <> "" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("A fajl mar letezik:" & vbCrLf & vbCrLf & fullPath & vbCrLf & vbCrLf & _
                      "Felulirashoz csak ments ra a fajlra, vagy modositsd az adatokat!" & vbCrLf & vbCrLf & _
                      "Akarsz uj verziot? (a fajl neve ki lesz egeszitve v2, v3, ...)", _
                      vbYesNo + vbQuestion, "Fajl mar letezik")
        If resp = vbYes Then
            Do
                versionNum = versionNum + 1
                fullPath = fullFolder & "\" & baseName & " v" & versionNum & ".xlsm"
            Loop While Dir(fullPath) <> ""
        Else
            SaveVersionedWorkbook = False
            Exit Function
        End If
    End If

    On Error GoTo SaveFail
    Application.DisplayAlerts = False
    activeWb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    Shell "explorer.exe """ & fullFolder & """", vbNormalFocus
    WriteLog "INFO SaveToDrive: saved to " & fullPath
    SaveVersionedWorkbook = True
    Exit Function

SaveFail:
    Application.DisplayAlerts = True
    MsgBox "A fajl nem menthetto: " & fullPath & vbCrLf & "Hiba: " & Err.Description, _
           vbCritical, APP_NAME
    WriteLog "ERROR SaveToDrive: " & Err.Description
    SaveVersionedWorkbook = False
End Function

' Egy szoveg minden szavanak elso betujet nagybetussel allitja.
Private Function CapitalizeWords(ByVal text As String) As String
    If Len(text) = 0 Then
        CapitalizeWords = text
        Exit Function
    End If
    Dim words() As String
    Dim i As Long
    words = Split(text, " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    CapitalizeWords = Join(words, " ")
End Function
