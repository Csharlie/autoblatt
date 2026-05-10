Option Explicit

' Autoblatt - Lokalis telepito script
' Importalja az excelStartupManager modult egy uj Excel workbook-ba,
' es meghivja a SetupPersonalWorkbook-ot, ami letrehozza a kiosztott
' autoblatt.xlsb-t a Documents\autoblatt mappaban.

Dim fso, currentPath, installPath, logFile, userName, logDir
Set fso = CreateObject("Scripting.FileSystemObject")

currentPath = fso.GetAbsolutePathName(".")
installPath = currentPath

userName = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERNAME%")
logDir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Autoblatt"

If Not fso.FolderExists(logDir) Then fso.CreateFolder(logDir)
logFile = logDir & "\install-" & userName & ".log"

Call Main()

Sub Main()
    On Error Resume Next
    Call WriteLog("=== AUTOBLATT lokalis telepites inditasa ===")
    Call WriteLog("Telepitesi konyvtar: " & installPath)

    If Not CheckPrerequisites() Then
        WScript.Echo "Telepites megszakitva: nem teljesulnek az elofeltetelek"
        WScript.Quit(1)
    End If

    If Not InstallExcelMacros() Then
        WScript.Echo "Telepites megszakitva: makro telepitesi hiba"
        WScript.Quit(1)
    End If

    If Not WriteLocalConfig() Then
        WScript.Echo "Telepites megszakitva: konfiguracios hiba"
        WScript.Quit(1)
    End If

    Call WriteLog("=== Lokalis telepites befejezve ===")
    WScript.Echo "Lokalis telepites sikeresen befejezve!"
    WScript.Quit(0)
End Sub

Function CheckPrerequisites()
    CheckPrerequisites = True
    Call WriteLog("Elofeltetelek ellenorzese...")

    If Not fso.FolderExists(installPath & "\core") Then
        Call WriteLog("HIBA: core mappa nem talalhato")
        CheckPrerequisites = False
    End If

    If Not fso.FolderExists(installPath & "\modules") Then
        Call WriteLog("HIBA: modules mappa nem talalhato")
        CheckPrerequisites = False
    End If

    If Not fso.FileExists(installPath & "\core\module-loader.bas") Then
        Call WriteLog("HIBA: module-loader.bas nem talalhato")
        CheckPrerequisites = False
    End If

    If CheckPrerequisites Then Call WriteLog("Elofeltetelek rendben")
End Function

Function InstallExcelMacros()
    InstallExcelMacros = True
    Call WriteLog("Excel makrok telepitese...")

    Dim excelApp, workbook, importPath
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Call WriteLog("HIBA: Excel nem elerheto: " & Err.Description)
        InstallExcelMacros = False
        Exit Function
    End If
    On Error GoTo 0

    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set workbook = excelApp.Workbooks.Add()

    importPath = installPath & "\installer\startup-manager.bas"
    If Not fso.FileExists(importPath) Then
        Call WriteLog("HIBA: Nem talalhato a modul: " & importPath)
        workbook.Close False
        excelApp.Quit
        InstallExcelMacros = False
        Exit Function
    End If

    On Error Resume Next
    workbook.VBProject.VBComponents.Import(importPath)
    If Err.Number <> 0 Then
        Call WriteLog("HIBA: startup-manager import sikertelen: " & Err.Description)
        Call WriteLog("Tipp: Az Excel Trust Center-ben engedelyezni kell a 'Trust access to the VBA project object model' opciot.")
        InstallExcelMacros = False
        workbook.Close False
        excelApp.Quit
        Exit Function
    End If
    Call WriteLog("startup-manager importalva")

    excelApp.Run "excelStartupManager.SetupPersonalWorkbook", installPath
    If Err.Number <> 0 Then
        Call WriteLog("HIBA: SetupPersonalWorkbook sikertelen: " & Err.Description)
        InstallExcelMacros = False
    Else
        Call WriteLog("autoblatt.xlsb sikeresen beallitva")
    End If
    On Error GoTo 0

    workbook.Close False
    excelApp.Quit
End Function

Function WriteLocalConfig()
    WriteLocalConfig = True
    Call WriteLog("Lokalis konfig irasa...")

    Dim configFile, configContent
    configFile = installPath & "\settings\system-config.json"

    If Not fso.FolderExists(installPath & "\settings") Then
        fso.CreateFolder(installPath & "\settings")
    End If

    Dim documentsPath
    documentsPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\autoblatt\"

    configContent = "{" & vbCrLf & _
                    "  ""appName"": ""Autoblatt""," & vbCrLf & _
                    "  ""appVersion"": ""1.0.0""," & vbCrLf & _
                    "  ""installType"": ""local""," & vbCrLf & _
                    "  ""installPath"": """ & Replace(installPath, "\", "\\") & """," & vbCrLf & _
                    "  ""installDate"": """ & Now() & """," & vbCrLf & _
                    "  ""userName"": """ & userName & """," & vbCrLf & _
                    "  ""targetWorkbookPath"": """ & Replace(documentsPath, "\", "\\") & """," & vbCrLf & _
                    "  ""targetWorkbookName"": ""autoblatt.xlsb""," & vbCrLf & _
                    "  ""dataXlsmPath"": """ & Replace(installPath & "\data\data.xlsm", "\", "\\") & """," & vbCrLf & _
                    "  ""saveRootFolder"": ""Sheets""," & vbCrLf & _
                    "  ""replaceRules"": ""bpl=BPL;sw=SW;li=Li""," & vbCrLf & _
                    "  ""emailSignatureMode"": ""outlook""" & vbCrLf & _
                    "}"

    On Error Resume Next
    Dim file
    Set file = fso.CreateTextFile(configFile, True)
    file.WriteLine configContent
    file.Close

    If Err.Number <> 0 Then
        Call WriteLog("HIBA: Konfig iras sikertelen: " & Err.Description)
        WriteLocalConfig = False
    Else
        Call WriteLog("Konfig leirva: " & configFile)
    End If
    On Error GoTo 0
End Function

Sub WriteLog(message)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"

    Dim existingContent : existingContent = ""
    If fso.FileExists(logFile) Then
        stream.Open
        stream.LoadFromFile logFile
        existingContent = stream.ReadText
        stream.Close
    End If

    stream.Open
    stream.WriteText existingContent & Now() & " - " & message & vbCrLf
    stream.SaveToFile logFile, 2
    stream.Close
    On Error GoTo 0
End Sub
