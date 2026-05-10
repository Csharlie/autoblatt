Attribute VB_Name = "cleanupHelper"
Option Explicit

' ============================================================
' AUTOBLATT - DUPLIKACIO CLEANUP (FEJLESZTOI ESZKOZ)
' ============================================================
' A dinamikus modul-loader minden futaskor importalja a .bas
' fajlokat. Ha nem takaritunk, a VBA editor felhalmozza a klonokat
' (driveManager1, driveManager2, ...). Ez az eszkoz eltavolitja
' azokat.
'
' A standalone dist/autoblatt.xlsb-ben a modulok be vannak egetve,
' ott ez nem szukseges.
' ============================================================

' --- AB_CleanupSystem: dinamikusan toltott modul klonok torlese ---
Public Sub AB_CleanupSystem()
    On Error Resume Next

    Dim modulesToClean As Variant
    modulesToClean = Array(VBMODULE_DRIVE, VBMODULE_IMAGE, VBMODULE_EMAIL, _
                           VBMODULE_FILL, VBMODULE_HELPRE)

    Dim totalBefore As Long
    totalBefore = ThisWorkbook.VBProject.VBComponents.Count

    Dim cleanedCount As Long
    Dim i As Long
    For i = LBound(modulesToClean) To UBound(modulesToClean)
        Dim vbComp As Object
        Set vbComp = Nothing
        Set vbComp = ThisWorkbook.VBProject.VBComponents(CStr(modulesToClean(i)))
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            cleanedCount = cleanedCount + 1
        End If
    Next i

    Dim totalAfter As Long
    totalAfter = ThisWorkbook.VBProject.VBComponents.Count

    MsgBox APP_NAME & " Cleanup eredmeny:" & vbCrLf & vbCrLf & _
           "Modulok elotte: " & totalBefore & vbCrLf & _
           "Eltavolitott klonok: " & cleanedCount & vbCrLf & _
           "Modulok utana: " & totalAfter & vbCrLf & vbCrLf & _
           "A makro lista most tiszta!", _
           vbInformation, APP_NAME

    On Error GoTo 0
End Sub

' --- AB_ListAllModules: jelenleg betoltott VBA komponensek listazasa ---
Public Sub AB_ListAllModules()
    Dim list As String
    list = APP_NAME & " - betoltott VBA komponensek:" & vbCrLf & String(40, "=") & vbCrLf

    Dim i As Long
    Dim vbComp As Object
    For i = 1 To ThisWorkbook.VBProject.VBComponents.Count
        Set vbComp = ThisWorkbook.VBProject.VBComponents(i)
        list = list & i & ". " & vbComp.Name & " (" & GetComponentTypeName(vbComp.Type) & ")" & vbCrLf
    Next i

    MsgBox list, vbInformation, APP_NAME & " Modul Lista"
End Sub

Private Function GetComponentTypeName(ByVal compType As Long) As String
    Select Case compType
        Case 1: GetComponentTypeName = "Modul"
        Case 2: GetComponentTypeName = "Class"
        Case 3: GetComponentTypeName = "UserForm"
        Case 100: GetComponentTypeName = "Document"
        Case Else: GetComponentTypeName = "Ismeretlen"
    End Select
End Function
