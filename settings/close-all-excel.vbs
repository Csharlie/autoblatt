' Autoblatt - Minden megnyitott Excel munkafuzet bezarasa mentes nelkul

On Error Resume Next
Dim excelApp, wb

Set excelApp = GetObject(, "Excel.Application")
If Not excelApp Is Nothing Then
    For Each wb In excelApp.Workbooks
        wb.Close False
    Next
    excelApp.Quit
    Set excelApp = Nothing
End If
On Error GoTo 0
