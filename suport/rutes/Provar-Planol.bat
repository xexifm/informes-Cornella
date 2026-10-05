@echo off
REM ---------------------------------------------------------------------------
REM  Provar-Planol.bat - comprova que els dos serveis del Cadastre que fa servir
REM  el "Planol activitats" responen i que el programa n'enten la resposta:
REM    1. la geometria d'una parcel.la (INSPIRE, parcel.les cadastrals)
REM    2. l'escala, planta i porta d'una unitat (consulta descriptiva)
REM
REM  NO obre cap finestra, no toca cap Excel i no modifica res: nomes pregunta.
REM  Desa les respostes senceres a local\geocodificacio\ (resposta-*.xml).
REM
REM  Doble clic (fa servir una unitat d'exemple de Cornella, al C/ Cadis 19) o
REM  passar-li una referencia cadastral de 20 caracters:
REM      Provar-Planol.bat 4091106DF2749A0006XJ
REM
REM  Un .bat i no una comanda per escriure a ma: llancada des d'un PowerShell,
REM  el shell de FORA expandeix el $env: abans de passar-lo (vegeu
REM  Provar-Cadastre.bat). Aqui la variable la posa el cmd.
REM
REM  ASCII pur i sense accents a proposit.
REM ---------------------------------------------------------------------------
setlocal

set "REFCAT=%~1"
if "%REFCAT%"=="" set "REFCAT=2295827DF2729E0011RQ"

REM Mode headless: Planol.ps1 nomes defineix funcions i NO obre l'eina.
set "PLANOL_TEST=1"

pushd "%~dp0..\.."

echo Provant els serveis del Cadastre del Planol activitats amb %REFCAT%...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ". suport\rutes\Planol.ps1; Test-Planol '%REFCAT%'"

popd
echo.
pause
endlocal
