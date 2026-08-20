@echo off
REM ---------------------------------------------------------------------------
REM  Provar-Cadastre.bat - comprova que el servei d'adreces del Cadastre respon
REM  i que el programa enten la resposta.
REM
REM  Fa UNA consulta i ensenya els portals que n'ha tret. NO obre cap finestra,
REM  no toca l'Excel i no modifica res: nomes pregunta.
REM
REM  Es pot fer DOBLE CLIC (fa servir una parcel.la d'exemple de Cornella, la
REM  illa dels carrers Cadis i Huelva) o passar-li una referencia cadastral:
REM      Provar-Cadastre.bat 4091106DF2749A
REM
REM  Existeix perque la comanda equivalent escrita a ma es un camp de mines de
REM  cometes: llancada des d'un PowerShell, el shell de FORA expandeix el
REM  $env:COORDENADES_TEST abans de passar-lo i al de dins li arriba "=1; ...".
REM  Aqui la variable la posa el cmd i a la linia de PowerShell no hi ha cap $.
REM
REM  ASCII pur i sense accents a proposit: els .bat amb accents es trenquen
REM  segons la codepage de la consola.
REM ---------------------------------------------------------------------------
setlocal

set "REFCAT=%~1"
if "%REFCAT%"=="" set "REFCAT=2295827DF2729E"

REM Mode headless: Coordenades.ps1 nomes defineix funcions i NO obre la finestra
REM de l'eina (ni carrega el planificador de rutes).
set "COORDENADES_TEST=1"

REM L'arrel del clone es dos nivells amunt d'aquest fitxer (suport\rutes\).
pushd "%~dp0..\.."

echo Provant el servei d'adreces del Cadastre amb la referencia %REFCAT%...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ". suport\rutes\Coordenades.ps1; Test-Geocodificador '%REFCAT%'"

popd
echo.
pause
endlocal
