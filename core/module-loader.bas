Attribute VB_Name = "moduleLoader"
Option Explicit

' ============================================================
' AUTOBLATT - DINAMIKUS MODUL BETOLTO (FEJLESZTOI MOD)
' ============================================================
' A fejlesztoi xlsm-ben a modulok futasidoben importalodnak a
' fajlrendszerrol (core/, modules/), igy a .bas fajlok git-ben
' szerkesztesek.
'
' A standalone dist/autoblatt.xlsb-ben a modulok be vannak egetve,
' ott ez a loader nem kerul lefuttatasra (a simple-macros direkt
' meghivja a modulok fuggvenyeit).
' ============================================================

' Bejegyzett modul-meghatarozas
Private Type ModuleSpec
    relativePath As String
    componentName As String
    mainFunction As String
End Type

' --- Publikus belepesi pont ---
' Egy modul kulcsanak megfelelo .bas fajl betoltese es futtatasa.
Public Sub LoadAndRunModule(ByVal moduleKey As String)
    Dim spec As ModuleSpec
    spec = GetModuleSpec(moduleKey)

    If Len(spec.componentName) = 0 Then
        MsgBox "Ismeretlen modul kulcs: " & moduleKey, vbCritical, "Autoblatt"
        Exit Sub
    End If

    Dim fullPath As String
    fullPath = GetAppRootPath() & "\" & spec.relativePath

    If Dir(fullPath) = "" Then
        MsgBox "Modul fajl nem talalhato:" & vbCrLf & fullPath, vbCritical, "Autoblatt"
        WriteLog "ERROR LoadAndRunModule: file not found - " & fullPath
        Exit Sub
    End If

    ' Korabbi importalt valtozat eltavolitasa (klon felhalmozodas megelozese)
    On Error Resume Next
    Dim existing As Object
    Set existing = ThisWorkbook.VBProject.VBComponents(spec.componentName)
    If Not existing Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove existing
    End If
    On Error GoTo 0

    ' Uj komponens importalasa
    On Error GoTo ImportFail
    Dim imported As Object
    Set imported = ThisWorkbook.VBProject.VBComponents.Import(fullPath)
    imported.Name = spec.componentName

    ' Fofuggveny futtatasa
    Application.Run spec.componentName & "." & spec.mainFunction
    Exit Sub

ImportFail:
    MsgBox "Hiba a modul importalasa kozben:" & vbCrLf & Err.Description, _
           vbCritical, "Autoblatt"
    WriteLog "ERROR LoadAndRunModule import: " & Err.Description
End Sub

' --- Modul-specifikaciok kulcs alapjan ---
Private Function GetModuleSpec(ByVal key As String) As ModuleSpec
    Dim spec As ModuleSpec

    Select Case key
        Case MODULE_KEY_DRIVE
            spec.relativePath = "modules\drive-manager.bas"
            spec.componentName = VBMODULE_DRIVE
            spec.mainFunction = "RunDriveManager"

        Case MODULE_KEY_IMAGE
            spec.relativePath = "modules\image-importer.bas"
            spec.componentName = VBMODULE_IMAGE
            spec.mainFunction = "RunImageImporter"

        Case MODULE_KEY_EMAIL
            spec.relativePath = "modules\email-sender.bas"
            spec.componentName = VBMODULE_EMAIL
            spec.mainFunction = "RunEmailSender"

        Case MODULE_KEY_FILL
            spec.relativePath = "modules\fill-panel.bas"
            spec.componentName = VBMODULE_FILL
            spec.mainFunction = "RunFillPanel"

        Case MODULE_KEY_HELPRE
            spec.relativePath = "modules\helpre-panel.bas"
            spec.componentName = VBMODULE_HELPRE
            spec.mainFunction = "RunHelprePanel"
    End Select

    GetModuleSpec = spec
End Function

' --- Elerheto modulok listaja ---
Public Function GetAvailableModules() As String()
    Dim modules(0 To 4) As String
    modules(0) = MODULE_KEY_DRIVE
    modules(1) = MODULE_KEY_IMAGE
    modules(2) = MODULE_KEY_EMAIL
    modules(3) = MODULE_KEY_FILL
    modules(4) = MODULE_KEY_HELPRE
    GetAvailableModules = modules
End Function
