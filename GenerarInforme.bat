@echo off
REM Llanca el generador d'informes Cornella. Doble clic per executar.
REM
REM S'obre SENSE cap finestra de consola: el llancament el fa un .vbs
REM (suport\GenerarInforme.vbs) via wscript, que NO mostra finestra i arrenca
REM PowerShell del tot amagat. Aixi ja no es veu la llampada blava de la
REM consola de PowerShell. Els errors i missatges es mostren en finestres
REM (MessageBox) del propi programa, no a cap consola.
REM
REM Si el programa ja esta obert, no se n'obre un segon: es porta al davant la
REM finestra que ja hi ha (una sola instancia).
REM
REM %~dp0 resol al directori d'aquest .bat; el .vbs i el .ps1 viuen a suport\.

cd /d "%~dp0"

start "" wscript.exe "%~dp0suport\GenerarInforme.vbs"
