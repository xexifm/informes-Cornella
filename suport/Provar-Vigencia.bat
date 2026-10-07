@echo off
REM ---------------------------------------------------------------------------
REM  Provar-Vigencia.bat - diagnostic de la vigencia de la normativa.
REM
REM  Desa el que responen el Portal Juridic i el BOE per a unes quantes normes
REM  (la pagina, les metadades, els scripts i la pagina dibuixada per l'Edge) a
REM  local\revisions\diagnostic-vigencia-*.zip, perque es pugui veure com saber
REM  si una norma es vigent sense que l'Edge es pengi.
REM
REM  NO modifica res. Triga uns minuts. Doble clic i, en acabar, passa el .zip
REM  a Claude.
REM
REM  ASCII pur i sense accents a proposit.
REM ---------------------------------------------------------------------------
setlocal
pushd "%~dp0"
echo Provant el Portal Juridic i el BOE (uns minuts)...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ProvarVigencia.ps1"
popd
echo.
pause
endlocal
