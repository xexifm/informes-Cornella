@echo off
REM Actualitza el programa a l'ultima versio de la branca estable "main".
REM Doble clic per executar. Sempre et deixa a "main" encara que el clone
REM s'hagi quedat en una altra branca.

title Actualitzar generador d'informes Cornella
cd /d "%~dp0"

echo Baixant l'ultima versio (branca main)...
git fetch origin
git checkout main
git pull --ff-only origin main

echo.
echo Fet. Ja tens l'ultima versio.
pause
