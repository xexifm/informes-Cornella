@echo off
REM ===========================================================================
REM  Crea l'acces directe del Generador d'informes per poder-lo ANCORAR A LA
REM  BARRA DE TASQUES. Doble clic i llestos.
REM ===========================================================================
REM  Windows NO deixa ancorar un .bat a la barra de tasques: nomes hi admet
REM  accessos directes que apuntin a un EXECUTABLE. Aquest en fa un que apunta a
REM  wscript.exe amb el llancador del programa (suport\GenerarInforme.vbs), que
REM  es exactament el que ja fa GenerarInforme.bat, i li posa l'escut.
REM
REM  El deixa a l'ESCRIPTORI i al MENU INICI. Despres, per ancorar-lo:
REM     clic dret damunt de l'acces directe  ->  "Ancorar a la barra de tasques"
REM  (al Windows 11 pot caldre triar abans "Mostra mes opcions").
REM
REM  ASCII PUR: els .bat amb accents es trenquen segons la codepage.
REM  La feina la fa suport\AccesDirecte.ps1: una ordre de PowerShell llarga dins
REM  d'un .bat, amb cometes i canonades, es un niu d'errors.
REM ===========================================================================

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ". '%~dp0suport\AccesDirecte.ps1'; Invoke-CrearAccesDirecte '%~dp0' | Out-Null"

echo.
echo ---------------------------------------------------------------------------
echo  Per ANCORAR-LO A LA BARRA DE TASQUES:
echo    clic dret damunt de l'acces directe (escriptori o menu Inici)
echo    i tria "Ancorar a la barra de tasques".
echo.
echo  Al Windows 11 pot ser que primer hagis de triar "Mostra mes opcions".
echo.
echo  SI JA EL TENIES ANCORAT: treu-lo de la barra i torna-hi a posar. Windows
echo  es queda la copia del dia que el vas ancorar, o sigui que fins que no el
echo  tornis a ancorar seguira sense icona i obrira un segon boto.
echo ---------------------------------------------------------------------------
echo.
pause
