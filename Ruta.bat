@echo off
REM Planificador de rutes d'inspeccio. Doble clic per executar.
REM Demana un llistat d'ID Activitat, els geolocalitza al mapa i en calcula
REM la ruta circular mes rapida. El .ps1 viu a suport\ ; %~dp0 resol al
REM directori d'aquest .bat.

title Planificador de rutes - Cornella

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\Ruta.ps1"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo  S'ha produit un error al script. Premeu una tecla per tancar.
    echo ============================================================
    pause >nul
)
