Attribute VB_Name = "ribbonCallbacks"
Option Explicit

' ============================================================
' AUTOBLATT - SZALAG (RIBBON) CALLBACK-EK
' ============================================================
' A ribbon-config.xml gombjai ezeket hivjak. Cegen belul a custom
' ribbon le van tiltva policy-vel - ezert az elsodleges UI a
' munkalapra agyazott gombok lesznek (lasd ui/ mappa).
' Ezek a callback-ek megmaradnak arra az esetre, ha egyszer
' engedelyezi a policy a custom ribbont.
' ============================================================

Public Sub OnRunDriveManager(control As IRibbonControl)
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_DRIVE)
End Sub

Public Sub OnRunImageImporter(control As IRibbonControl)
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_IMAGE)
End Sub

Public Sub OnRunEmailSender(control As IRibbonControl)
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_EMAIL)
End Sub

Public Sub OnRunFillPanel(control As IRibbonControl)
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_FILL)
End Sub

Public Sub OnRunHelprePanel(control As IRibbonControl)
    Call moduleLoader.LoadAndRunModule(MODULE_KEY_HELPRE)
End Sub

Public Sub OnOpenSettings(control As IRibbonControl)
    On Error GoTo Fail
    Call settingsUI.OpenSettingsDialog
    Exit Sub
Fail:
    MsgBox "A beallitasok panel meg nem erheto el." & vbCrLf & _
           "Hiba: " & Err.Description, vbExclamation, APP_NAME
End Sub
