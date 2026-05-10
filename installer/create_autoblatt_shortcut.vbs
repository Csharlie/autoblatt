Option Explicit

' Autoblatt - Asztali parancsikon letrehozasa
' A bat sajat helyzetebol szamol az icon es a launcher utvonalat,
' nincs hardkodolt P:\ utvonal.

Dim objFSO, objShell, scriptDir, icoSource, lnkPath, lnk, desktop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' A script sajat mappaja (installer) -> egy szinttel feljebb a project root
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
Dim projectRoot
projectRoot = objFSO.GetParentFolderName(scriptDir)

icoSource = projectRoot & "\launcher_icon.ico"
desktop = objShell.SpecialFolders("Desktop")
lnkPath = desktop & "\AutoBlatt.lnk"

If Not objFSO.FileExists(icoSource) Then
    ' Az icon nem kotelezo - hianya nem akadalyozza a parancsikon letrejotttet
    icoSource = ""
End If

Set lnk = objShell.CreateShortcut(lnkPath)
lnk.TargetPath = projectRoot & "\setup-launcher.bat"
If Len(icoSource) > 0 Then lnk.IconLocation = icoSource
lnk.WorkingDirectory = projectRoot
lnk.Save

MsgBox "Az Autoblatt parancsikon elkeszult az asztalon!" & vbCrLf & "Cel: " & projectRoot & "\setup-launcher.bat", 64, "Kesz"
