@echo off
REM Recupera els catalegs (ESTRUCTURALS) que es van quedar "amagats" al stash
REM amb la versio antiga d'Actualitzar.bat.
REM
REM NO toca res del programa: nomes EXTREU aquells fitxers a
REM %LOCALAPPDATA%\InformesCornella\recuperats\ perque els puguis mirar i,
REM si vols, copiar-los tu mateix a ESTRUCTURALS.
REM
REM Doble clic per executar.

setlocal
title Recuperar catalegs perduts
cd /d "%~dp0"

echo === Recuperar catalegs dels stashes ===
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\SincronitzaCatalegs.ps1" -Fase Recuperar

echo.
echo Prem qualsevol tecla per obrir la carpeta amb el que s'hagi recuperat...
pause >nul
if exist "%LOCALAPPDATA%\InformesCornella\recuperats" start "" "%LOCALAPPDATA%\InformesCornella\recuperats"
exit /b 0
