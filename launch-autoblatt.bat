@echo off
chcp 1250 >nul
title Autoblatt Launcher

REM Az autoblatt.xlsb (kiosztott verzio) inditasa
REM A bat sajat helyzetebol (dist mellett) szamol az xlsb utvonalat.

set "APP_ROOT=%~dp0"
set "XLSB_PATH=%APP_ROOT%dist\autoblatt.xlsb"
set "DEV_PATH=%APP_ROOT%autoblatt.xlsm"

if exist "%XLSB_PATH%" (
    echo [OK] Autoblatt fajl megtalalhato: %XLSB_PATH%
    start "" "%XLSB_PATH%"
    goto END
)

if exist "%DEV_PATH%" (
    echo [INFO] Standalone xlsb nem talalhato, fejlesztoi xlsm inditasa.
    start "" "%DEV_PATH%"
    goto END
)

echo [HIBA] Sem dist\autoblatt.xlsb, sem autoblatt.xlsm nem talalhato.
echo Ellenorizd a setup-launcher.bat-tal a telepitest.
pause
exit /b 1

:END
exit /b 0
