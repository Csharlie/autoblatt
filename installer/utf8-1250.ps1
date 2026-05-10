# Autoblatt - UTF-8 -> Windows-1250 konverter
# A .bas fajlokat a VBA editor cp1250-ben varja (magyar Windows alapertelmezese).
# Ez a script bejarja a project mappat, es minden szoveges fajlt
# UTF-8-bol Windows-1250-be konvertalt.
#
# Hasznalat:
#   PowerShell -ExecutionPolicy Bypass -File installer\utf8-1250.ps1
#
# Csak a .bas, .frm, .vbs, .xml, .json, .bat fajlokat erinti.

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot

$extensionsToConvert = @('.bas', '.frm', '.cls')

Write-Host "Autoblatt encoding converter (UTF-8 -> Windows-1250)" -ForegroundColor Cyan
Write-Host "Project root: $projectRoot"
Write-Host ""

$count = 0
Get-ChildItem -Path $projectRoot -File -Recurse | Where-Object {
    $extensionsToConvert -contains $_.Extension.ToLower()
} | ForEach-Object {
    try {
        $utf8Content = Get-Content -Path $_.FullName -Encoding UTF8
        [System.IO.File]::WriteAllLines($_.FullName, $utf8Content, [System.Text.Encoding]::GetEncoding(1250))
        Write-Host "OK:  $($_.FullName)" -ForegroundColor Green
        $count++
    }
    catch {
        Write-Host "ERR: $($_.FullName) - $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Konvertalt: $count fajl" -ForegroundColor Cyan
