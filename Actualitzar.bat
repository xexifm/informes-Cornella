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

git checkout main
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut canviar a la branca 'main'.
    echo    Causa probable: tens canvis locals al codi que bloquegen el canvi.
    echo    Solucio: descarta-los amb 'git reset --hard' (es perden) o committeja'ls.
    echo.
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
echo.
echo Fet. Ja tens l'ultima versio.
pause
