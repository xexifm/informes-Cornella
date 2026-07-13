@echo off
REM Llanca el generador d'informes Cornella. Doble clic per executar.
REM
REM S'obre SENSE finestra de consola (el CMD queda amagat): nomes es veuen les
REM finestres del programa. Els errors es mostren en finestres (MessageBox), no
REM a la consola. Per aixo llancem PowerShell amb -WindowStyle Hidden i amb
REM 'start' (el .bat surt de seguida i no deixa cap consola oberta).
REM
REM %~dp0 resol al directori d'aquest .bat; el .ps1 viu a suport\.

cd /d "%~dp0"

start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0suport\GenerarInforme.ps1"
