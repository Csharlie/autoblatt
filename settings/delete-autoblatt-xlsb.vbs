' Autoblatt - Telepitett autoblatt.xlsb torlese (uj telepiteshez)
'
' Bezarja az osszes Excel folyamatot, majd torli az autoblatt.xlsb-t
' a Documents\autoblatt mappabol es az Excel XLSTART mappabol.

On Error Resume Next

Dim objShell, fso, userProfile, appData, targets, t

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Excel folyamat killolesa
objShell.Run "taskkill /f /im EXCEL.EXE", 0, True

WScript.Echo "autoblatt.xlsb torlese..."

userProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
appData = objShell.ExpandEnvironmentStrings("%APPDATA%")

targets = Array( _
    userProfile & "\Documents\autoblatt\autoblatt.xlsb", _
    appData & "\Microsoft\Excel\XLSTART\autoblatt.xlsb", _
    appData & "\Microsoft\Excel\XLSTART\personal.xlsb" _
)

For Each t In targets
    If fso.FileExists(t) Then
        fso.DeleteFile t, True
        WScript.Echo "Torolve: " & t
    End If
Next

WScript.Echo "Kesz!"
