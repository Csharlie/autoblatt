@echo off
chcp 1250 >nul
title Autoblatt Telepito
color 0A
cls
echo ******************************************************************************
echo *                                                                            *
echo *                                  AUTOBLATT                                 *
echo *                           Peter Sardy Production                           *
echo *                                                                            *
echo ******************************************************************************
echo *                                                                            *
echo *        A kovetkezokben egy Modult fogsz importalni Excelhez, amely         *
echo *        megkonnyiti a munkadat a Foto- es Infolapok feldolgozasaban.        *
echo *                                                                            *
echo ******************************************************************************

set "APP_ROOT=%~dp0"

:MENU
echo.
echo MENU:
echo.
echo [0] Autoblatt inditasa
echo [1] Lokalis telepites
echo [2] Halozati telepites (ha P:\ elerheto)
echo [3] Excel bezarasa
echo [4] Telepitett autoblatt.xlsb torlese
echo [x] Kilep
echo.
set /p choice=Valasztas (0-4):

if "%choice%"=="0" goto LAUNCH
if "%choice%"=="1" goto LOCAL_INSTALL
if "%choice%"=="2" goto NETWORK_INSTALL
if "%choice%"=="3" goto CLOSE_EXCEL
if "%choice%"=="4" goto DELETE_XLSB
if "%choice%"=="x" goto EXIT
echo Ervenytelen valasztas!
goto MENU

:LAUNCH
call "%APP_ROOT%launch-autoblatt.bat"
goto MENU

:LOCAL_INSTALL
echo.
echo Lokalis telepites inditasa...
cscript //nologo "%APP_ROOT%installer\local-installer.vbs"
if %errorlevel% neq 0 (
    echo HIBA: Lokalis telepites sikertelen!
    goto MENU
)
echo Lokalis telepites befejezve!
goto MENU

:NETWORK_INSTALL
echo.
echo Halozati telepites inditasa...
if not exist "P:\" (
    echo [HIBA] P:\ meghajto nem elerheto.
    goto MENU
)
cscript //nologo "%APP_ROOT%installer\network-installer.vbs"
if %errorlevel% neq 0 (
    echo HIBA: Halozati telepites sikertelen!
    goto MENU
)
echo Halozati telepites befejezve!
goto MENU

:CLOSE_EXCEL
cscript //nologo "%APP_ROOT%settings\close-all-excel.vbs"
goto MENU

:DELETE_XLSB
cscript //nologo "%APP_ROOT%settings\delete-autoblatt-xlsb.vbs"
goto MENU

:EXIT
exit /b 0
