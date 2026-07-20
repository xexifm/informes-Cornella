@echo off
REM Obre el programa AMB consola visible, per veure qualsevol error d'arrencada.
REM (El GenerarInforme.bat normal l'amaga; per aixo, si peta en arrencar, sembla
REM que "no fa res". Aquest .bat mostra l'error.)
REM Doble clic per executar. Si surt text en VERMELL, fes-ne una captura.

cd /d "%~dp0"
echo ============================================================
echo  Obrint el programa amb consola visible (diagnostic)...
echo  Si hi ha un error, es veura mes avall EN VERMELL.
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\GenerarInforme.ps1"

echo.
echo ============================================================
echo  Fi. Si hi ha hagut un error a dalt (vermell), fes captura
echo  i envia-la a Claude. Si s'ha obert el programa, tot be.
echo ============================================================
pause
