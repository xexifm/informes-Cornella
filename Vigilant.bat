@echo off
REM Vigilant: genera sol els informes que arriben del mobil (via Google Drive).
REM Deixa aquesta finestra oberta en segon pla mentre treballes al PC.
REM Doble clic per executar.

title Vigilant d'informes Cornella
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\Vigilant.ps1"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo  S'ha produit un error al vigilant. Premeu una tecla per tancar.
    echo ============================================================
    pause >nul
)
