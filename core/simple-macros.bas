Attribute VB_Name = "simpleMacros"
Option Explicit

' ============================================================
' AUTOBLATT FELHASZNALOI MAKROK
' Alt+F8 -> AB_ prefixu makrok -> Futtatas
' ============================================================

' --- AB_SaveToDrive: drive mentes ---
Public Sub AB_SaveToDrive()
    On Error GoTo Fail
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_DRIVE)
    Exit Sub
Fail:
    MsgBox "Hiba a SaveToDrive futtatasakor: " & Err.Description, vbCritical, APP_NAME
End Sub

' --- AB_ImportImages: kepek importalasa ---
Public Sub AB_ImportImages()
    On Error GoTo Fail
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_IMAGE)
    Exit Sub
Fail:
    MsgBox "Hiba az ImportImages futtatasakor: " & Err.Description, vbCritical, APP_NAME
End Sub

' --- AB_SendEmails: Outlook draft email ---
Public Sub AB_SendEmails()
    On Error GoTo Fail
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_EMAIL)
    Exit Sub
Fail:
    MsgBox "Hiba a SendEmails futtatasakor: " & Err.Description, vbCritical, APP_NAME
End Sub

' --- AB_FillPanel: Fotolap kitoltopanel kulon makro ---
Public Sub AB_FillPanel()
    On Error GoTo Fail
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_FILL)
    Exit Sub
Fail:
    MsgBox "Hiba a FillPanel futtatasakor: " & Err.Description, vbCritical, APP_NAME
End Sub

' --- AB_HelprePanel: kitoltesi segedlet kulon makro ---
Public Sub AB_HelprePanel()
    On Error GoTo Fail
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_HELPRE)
    Exit Sub
Fail:
    MsgBox "Hiba a HelprePanel futtatasakor: " & Err.Description, vbCritical, APP_NAME
End Sub

' --- AB_SystemTest: minden modul ellenorzese ---
Public Sub AB_SystemTest()
    Call RunAllModuleTests
End Sub

Private Sub RunAllModuleTests()
    On Error Resume Next
    Dim modules() As String
    modules = moduleLoader.GetAvailableModules()

    Dim i As Long
    Dim okCount As Long, errCount As Long
    For i = LBound(modules) To UBound(modules)
        Call moduleLoader.LoadAndRunModule(modules(i))
        If Err.Number = 0 Then
            okCount = okCount + 1
        Else
            errCount = errCount + 1
            Err.Clear
        End If
    Next i

    Dim msg As String
    msg = "Rendszer teszt eredmenye:" & vbCrLf & vbCrLf & _
          "Sikeres modulok: " & okCount & vbCrLf & _
          "Hibas modulok: " & errCount & vbCrLf & vbCrLf
    If errCount = 0 Then
        msg = msg & "Minden modul mukodik!"
    Else
        msg = msg & "Nehany modul javitasra szorul."
    End If
    MsgBox msg, vbInformation, APP_NAME & " Rendszer Teszt"
End Sub

' --- AB_Settings: beallitasok panel megnyitasa ---
Public Sub AB_Settings()
    On Error GoTo Fail
    Call settingsUI.OpenSettingsDialog
    Exit Sub
Fail:
    MsgBox "A beallitasok panel meg nem erheto el." & vbCrLf & _
           "Hiba: " & Err.Description, vbExclamation, APP_NAME
End Sub

' --- AB_Help: rovid sugo ---
Public Sub AB_Help()
    Dim h As String
    h = APP_NAME & " v" & APP_VERSION & " makrok:" & vbCrLf & vbCrLf & _
        "AB_SaveToDrive    - Adatok lekerese, formazas, mentes a OneDrive-ra" & vbCrLf & _
        "AB_ImportImages   - Kepek beillesztese 2 oszlopos racsban" & vbCrLf & _
        "AB_SendEmails     - Outlook draft email a munkafuzettel mellekletkent" & vbCrLf & _
        "AB_FillPanel      - Fotolap kitoltopanel letrehozasa" & vbCrLf & _
        "AB_HelprePanel    - Kitoltesi segedlet letrehozasa" & vbCrLf & _
        "AB_InstallButtons - 'AB Indito' lap letrehozasa kattinthato gombokkal" & vbCrLf & _
        "AB_Settings       - Beallitasok panel" & vbCrLf & _
        "AB_SystemTest     - Osszes modul ellenorzese" & vbCrLf & _
        "AB_CleanupSystem  - Dinamikusan toltott modul-klonok torlese" & vbCrLf & _
        "AB_ListAllModules - VBA komponensek listazasa" & vbCrLf & vbCrLf & _
        "Futtatas: Alt+F8 -> AB_ tipus -> Kivalasztas -> Run" & vbCrLf & _
        "Vagy: 'AB Indito' lap gombjai (futtasd egyszer az AB_InstallButtons-t)"
    MsgBox h, vbInformation, APP_NAME & " Sugo"
End Sub
