@echo off
REM Llanca el generador d'informes. Doble clic per executar.
REM El .bat es trasllada amb el .ps1: %~dp0 resol al directori del propi .bat.

title Generador d'informes Cornella

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0GenerarInforme.ps1"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo  S'ha produit un error al script. Premeu una tecla per tancar.
    echo ============================================================
    pause >nul
)
