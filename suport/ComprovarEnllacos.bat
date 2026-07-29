@echo off
REM Comprova els enllacos (URLs) dels catalegs. Doble clic per executar.
REM Per defecte comprova TOTS els catalegs .json d'ESTRUCTURALS. Necessita internet.
REM Aquest .bat viu a suport\ ; el .ps1 es al seu costat (%~dp0).

title Comprova enllacos Cornella

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Comprova-Enllacos.ps1" %*

echo.
echo ============================================================
echo  Comprovacio acabada. Premeu una tecla per tancar.
echo ============================================================
pause >nul
