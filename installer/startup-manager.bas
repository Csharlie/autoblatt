Attribute VB_Name = "excelStartupManager"
Option Explicit

' ============================================================
' AUTOBLATT - EXCEL STARTUP MANAGER (TELEPITESI MODUL)
' ============================================================
' A telepito VBS importalja ezt egy uj Excel workbook-ba, majd
' meghivja SetupPersonalWorkbook-ot. Ez letrehozza a kiosztott
' autoblatt.xlsb-t a Documents\autoblatt mappaban (vagy a config
' altal megadott helyen), es importalja a core modulokat.
' ============================================================

Public installBasePath As String
Public targetWorkbookPath As String
Public targetWorkbookName As String

' --- Fo telepitesi belepesi pont (VBS hivja) ---
Public Sub SetupPersonalWorkbook(Optional basePath As Variant)
    On Error GoTo Fail

    If IsMissing(basePath) Or IsEmpty(basePath) Then
        installBasePath = ""
    Else
        installBasePath = CStr(basePath)
    End If

    Call LoadConfiguration
    Call LogLine("=== Autoblatt setup elkezdve ===")
    Call LogLine("Install path: " & installBasePath)
    Call LogLine("Target workbook: " & targetWorkbookPath & targetWorkbookName)

    If Not EnsureTargetFolder() Then
        Call LogLine("HIBA: Cel mappa setup sikertelen!")
        Err.Raise 1001, "SetupPersonalWorkbook", "Cel mappa setup sikertelen"
    End If

    If Not CreateOrUpdateWorkbook() Then
        Call LogLine("HIBA: autoblatt.xlsb letrehozas/frissites sikertelen!")
        Err.Raise 1002, "SetupPersonalWorkbook", "autoblatt.xlsb letrehozas sikertelen"
    End If

    Call LogLine("=== Autoblatt setup befejezve ===")
    Exit Sub

Fail:
    Call LogLine("HIBA: Err.Number=" & Err.Number & " Description=" & Err.Description)
End Sub

' --- Config betoltese (vagy alapertelmezesek) ---
Private Sub LoadConfiguration()
    targetWorkbookPath = Environ("USERPROFILE") & "\Documents\autoblatt\"
    targetWorkbookName = "autoblatt.xlsb"

    Dim configPath As String
    configPath = installBasePath & "\settings\system-config.json"

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(configPath) Then
        Dim file As Object
        Set file = fso.OpenTextFile(configPath, 1)
        Dim content As String
        content = file.ReadAll
        file.Close

        ' Egyszeru kulcs-keresesi parsing - elegendo lapos kulcs-ertek parokhoz
        Dim val As String
        val = ExtractJsonValue(content, "targetWorkbookPath")
        If Len(val) > 0 Then
            targetWorkbookPath = Replace(val, "\\", "\")
            If Right(targetWorkbookPath, 1) <> "\" Then targetWorkbookPath = targetWorkbookPath & "\"
        End If

        val = ExtractJsonValue(content, "targetWorkbookName")
        If Len(val) > 0 Then targetWorkbookName = val
    End If
    On Error GoTo 0
End Sub

Private Function ExtractJsonValue(ByVal content As String, ByVal key As String) As String
    Dim pattern As String, startPos As Long, valStart As Long, valEnd As Long
    pattern = """" & key & """"
    startPos = InStr(1, content, pattern, vbBinaryCompare)
    If startPos = 0 Then Exit Function

    valStart = InStr(startPos + Len(pattern), content, """")
    If valStart = 0 Then Exit Function

    ' Az ertek vege a kovetkezo nem-escapelt "
    ' (paros szamu backslash elotte = nem escape, paratlan = escape)
    valEnd = valStart + 1
    Do While valEnd <= Len(content)
        If Mid(content, valEnd, 1) = """" Then
            Dim slashes As Long
            slashes = 0
            Do While valEnd - 1 - slashes >= valStart + 1
                If Mid(content, valEnd - 1 - slashes, 1) = "\" Then
                    slashes = slashes + 1
                Else
                    Exit Do
                End If
            Loop
            If slashes Mod 2 = 0 Then Exit Do
        End If
        valEnd = valEnd + 1
    Loop

    If valEnd > valStart Then
        ExtractJsonValue = Mid(content, valStart + 1, valEnd - valStart - 1)
    End If
End Function

' --- Cel mappa biztosit�sa (rekurzivan) ---
Private Function EnsureTargetFolder() As Boolean
    EnsureTargetFolder = True
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(targetWorkbookPath) Then Exit Function

    On Error Resume Next
    Call CreateFolderRecursive(fso, targetWorkbookPath)
    If Err.Number <> 0 Then
        Call LogLine("HIBA: Cel mappa letrehozas sikertelen: " & Err.Description)
        EnsureTargetFolder = False
    End If
    On Error GoTo 0
End Function

Private Sub CreateFolderRecursive(ByVal fso As Object, ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        Call CreateFolderRecursive(fso, parent)
    End If
    fso.CreateFolder folderPath
End Sub

' --- autoblatt.xlsb letrehozasa vagy frissitese ---
Private Function CreateOrUpdateWorkbook() As Boolean
    CreateOrUpdateWorkbook = True
    Dim wbPath As String
    wbPath = targetWorkbookPath & targetWorkbookName
    Call LogLine("Cel xlsb path: " & wbPath)

    Dim wb As Object
    Dim wasOpen As Boolean

    On Error Resume Next
    Set wb = Workbooks(targetWorkbookName)
    If Not wb Is Nothing Then wasOpen = True
    On Error GoTo 0

    If Not wasOpen Then
        On Error Resume Next
        Set wb = Workbooks.Open(wbPath)
        If Err.Number <> 0 Then
            Call LogLine("xlsb nem letezik, uj letrehozasa...")
            Err.Clear
            Set wb = Workbooks.Add()
            wb.SaveAs wbPath, 50 ' xlOpenXMLWorkbookMacroEnabled = 52, xlExcel12 = 50
            If Err.Number <> 0 Then
                Call LogLine("HIBA: xlsb letrehozas sikertelen: " & Err.Description)
                CreateOrUpdateWorkbook = False
                Exit Function
            End If
            Call LogLine("xlsb letrehozva")
        Else
            Call LogLine("xlsb megnyitva")
        End If
        On Error GoTo 0
    End If

    If Not ImportCoreModules(wb) Then
        CreateOrUpdateWorkbook = False
        Exit Function
    End If

    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
        Call LogLine("HIBA: xlsb mentes sikertelen: " & Err.Description)
        CreateOrUpdateWorkbook = False
        Exit Function
    End If
    If Not wasOpen Then wb.Close
    On Error GoTo 0
End Function

' --- Core modulok importalasa ---
Private Function ImportCoreModules(ByVal wb As Object) As Boolean
    ImportCoreModules = True
    Call LogLine("Core modulok importalasa...")

    Dim moduleFiles As Variant
    moduleFiles = Array( _
        "core\config.bas", _
        "core\utils.bas", _
        "core\module-loader.bas", _
        "core\cleanup-helper.bas", _
        "core\ribbon-callbacks.bas", _
        "core\simple-macros.bas", _
        "core\settings-ui.bas", _
        "ui\buttons-installer.bas" _
    )

    Dim i As Long
    For i = LBound(moduleFiles) To UBound(moduleFiles)
        Dim relativePath As String, fullPath As String, modName As String
        relativePath = CStr(moduleFiles(i))
        fullPath = installBasePath & "\" & relativePath
        modName = ModuleNameFromPath(relativePath)

        ' Letezo modul eltavolitasa
        On Error Resume Next
        Dim existing As Object
        Set existing = Nothing
        Set existing = wb.VBProject.VBComponents(modName)
        If Not existing Is Nothing Then
            wb.VBProject.VBComponents.Remove existing
            Call LogLine("Eltavolitva: " & modName)
        End If
        Err.Clear
        On Error GoTo 0

        ' Uj importalas
        On Error Resume Next
        wb.VBProject.VBComponents.Import fullPath
        If Err.Number <> 0 Then
            Call LogLine("HIBA: " & relativePath & " import sikertelen: " & Err.Description)
            ImportCoreModules = False
            Err.Clear
        Else
            Call LogLine("Importalva: " & relativePath)
        End If
        On Error GoTo 0
    Next i
End Function

' --- Modulnev konvencio: kebab-case fajl -> camelCase modulnev ---
Private Function ModuleNameFromPath(ByVal filePath As String) As String
    Dim fileName As String, baseName As String
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    baseName = Left(fileName, InStrRev(fileName, ".") - 1)

    Dim parts() As String
    parts = Split(baseName, "-")
    Dim result As String
    result = parts(0)
    Dim i As Long
    For i = 1 To UBound(parts)
        result = result & UCase(Left(parts(i), 1)) & Mid(parts(i), 2)
    Next i
    ModuleNameFromPath = result
End Function

' --- Egyszeru log - %LOCALAPPDATA%\Autoblatt\install-USER.log ---
Private Sub LogLine(ByVal message As String)
    On Error Resume Next
    Dim logDir As String, logFile As String
    logDir = Environ("LOCALAPPDATA") & "\Autoblatt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logDir) Then fso.CreateFolder logDir

    logFile = logDir & "\install-" & Environ("USERNAME") & ".log"

    Dim fnum As Integer
    fnum = FreeFile
    Open logFile For Append As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fnum
End Sub
