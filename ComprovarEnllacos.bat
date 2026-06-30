@echo off
REM Comprova els enllacos (URLs) dels catalegs. Doble clic per executar.
REM Per defecte comprova ESTRUCTURALS\REQ1.docx. Necessita internet.
REM El .ps1 viu a suport\ ; %~dp0 resol al directori d'aquest .bat.

title Comprova enllacos Cornella

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\Comprova-Enllacos.ps1" %*

echo.
echo ============================================================
echo  Comprovacio acabada. Premeu una tecla per tancar.
echo ============================================================
pause >nul
