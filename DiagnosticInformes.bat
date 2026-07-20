@echo off
REM Diagnostic de la carpeta d'informes (per afinar l'escaner).
REM Escaneja la carpeta al TEU PC i escriu un resum a l'Escriptori
REM (informes-diagnostic.txt). NO puja cap contingut: nomes estructura.
REM Doble clic per executar. Es veu una consola amb el progres.

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\DiagnosticInformes.ps1"

echo.
pause
