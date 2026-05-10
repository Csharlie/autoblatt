# Autoblatt - Standalone autoblatt.xlsb build script
#
# Letrehoz egy uj Excel workbook-ot, importalja az osszes
# core/, modules/, ui/ alatti .bas fajlt, majd lementi
# dist\autoblatt.xlsb-be (be vannak egetve a kodok, nincs loader).
#
# Hasznalat:
#   PowerShell -ExecutionPolicy Bypass -File installer\build-xlsb.ps1
#
# Elofeltetel: Excel telepitve, es a Trust Center-ben engedelyezve
# a "Trust access to the VBA project object model" opcio.

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$distFolder = Join-Path $projectRoot 'dist'
$outputPath = Join-Path $distFolder 'autoblatt.xlsb'

Write-Host "Autoblatt build (standalone xlsb)" -ForegroundColor Cyan
Write-Host "Project root: $projectRoot"
Write-Host "Output:       $outputPath"
Write-Host ""

# Dist mappa biztositasa
if (-not (Test-Path $distFolder)) {
    New-Item -ItemType Directory -Path $distFolder | Out-Null
}

# Korabbi build torlese
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
    Write-Host "Korabbi build torolve" -ForegroundColor Yellow
}

# Excel inditas
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # Importalando modulok (sorrend szamit: config -> utils -> tobbi)
    $modules = @(
        'core\config.bas',
        'core\utils.bas',
        'core\module-loader.bas',
        'core\cleanup-helper.bas',
        'core\ribbon-callbacks.bas',
        'core\simple-macros.bas',
        'core\settings-ui.bas',
        'modules\drive-manager.bas',
        'modules\fill-panel.bas',
        'modules\helpre-panel.bas',
        'modules\image-importer.bas',
        'modules\email-sender.bas',
        'ui\buttons-installer.bas'
    )

    foreach ($rel in $modules) {
        $full = Join-Path $projectRoot $rel
        if (-not (Test-Path $full)) {
            Write-Host "MISS: $rel" -ForegroundColor Yellow
            continue
        }
        try {
            $wb.VBProject.VBComponents.Import($full) | Out-Null
            Write-Host "IMP:  $rel" -ForegroundColor Green
        }
        catch {
            Write-Host "ERR:  $rel - $_" -ForegroundColor Red
        }
    }

    # Mentes xlsb formatumban (xlExcel12 = 50)
    $wb.SaveAs($outputPath, 50)
    Write-Host "`nMentve: $outputPath" -ForegroundColor Cyan
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`nBuild kesz." -ForegroundColor Green
