@echo off
REM Actualitza el programa a l'ultima versio de la branca estable "main".
REM Doble clic per executar.

setlocal
title Actualitzar generador d'informes Cornella
cd /d "%~dp0"

echo === Estat ABANS ===
git branch --show-current
git log -1 --format="%%h  %%s"
echo.

echo Fetching...
git fetch origin
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut connectar a GitHub. Comprova la xarxa.
    pause
    exit /b 1
)

REM Si hi ha canvis locals, els guardem en un stash de seguretat per
REM no perdre'ls i poder fer pull sense conflictes.
set "STASHED=0"
git diff --quiet
if errorlevel 1 set "STASHED=1"
git diff --cached --quiet
if errorlevel 1 set "STASHED=1"

if "%STASHED%"=="1" (
    echo Detectats canvis locals: els guardo com a copia de seguretat (stash)...
    git stash push -u -m "Auto-stash per Actualitzar.bat"
    if errorlevel 1 (
        echo ERROR: no s'ha pogut fer stash. Avorto.
        pause
        exit /b 1
    )
)

git checkout main
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut canviar a la branca 'main'.
    git status
    pause
    exit /b 1
)

git pull --ff-only origin main
if errorlevel 1 (
    echo.
    echo ERROR: 'git pull' ha fallat. Mira el missatge anterior.
    pause
    exit /b 1
)

echo.
echo === Estat DESPRES ===
git branch --show-current
git log -1 --format="%%h  %%s"

if "%STASHED%"=="1" (
    echo.
    echo NOTA: els teus canvis locals han quedat guardats al stash.
    echo   - Per recuperar-los:    git stash pop
    echo   - Per veure'ls:         git stash list
    echo   - Per descartar-los:    git stash drop
    echo   En molts casos no cal fer res; el stash es una xarxa de seguretat.
)

echo.
echo Fet. Ja tens l'ultima versio.
pause
