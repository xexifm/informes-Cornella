@echo off
REM Llanca el generador d'informes. Doble clic per executar.
REM El .ps1 viu a suport\ ; %~dp0 resol al directori d'aquest .bat.

title Generador d'informes Cornella

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\GenerarInforme.ps1"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo  S'ha produit un error al script. Premeu una tecla per tancar.
    echo ============================================================
    pause >nul
)
