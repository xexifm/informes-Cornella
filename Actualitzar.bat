@echo off
REM Actualitza el generador d'informes Cornella.
REM - Si tens canvis locals a ESTRUCTURALS\*.docx (plantilles), els puja a GitHub.
REM - Si tens canvis a codi (.ps1, .bat) els guarda al stash com a copia de
REM   seguretat i continua. No es perden mai.
REM - Sempre acaba a 'main' amb l'ultim commit i mostra resultat.
REM
REM Doble clic per executar.

setlocal EnableDelayedExpansion
title Actualitzar generador d'informes Cornella
cd /d "%~dp0"

echo === Estat ABANS ===
git branch --show-current
git log -1 --format="%%h  %%s"
echo.

REM --- 0. Detectar fitxers de bloqueig de Word/Excel ---
dir /b "ESTRUCTURALS\~$*" >nul 2>&1
if not errorlevel 1 (
    echo.
    echo  Hi ha plantilles obertes al Word ESTRUCTURALS\~$*  .
    echo    Tanca el Word i torna a executar Actualitzar.bat.
    echo.
    pause
    exit /b 1
)

REM --- 1. Fetch ---
echo Connectant a GitHub...
git fetch origin
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut connectar a GitHub. Comprova la xarxa.
    pause
    exit /b 1
)

REM --- 2. Hi ha canvis a ESTRUCTURALS\*.docx ? Si si, els pugem ---
REM    (codi ps1/bat NO el pujem mai des d'aqui: el toca Claude o tu via PR)
set "PLANTILLES_CANVIADES=0"
for /f "delims=" %%f in ('git status --porcelain ESTRUCTURALS ^| findstr /R "\.docx$"') do (
    set "PLANTILLES_CANVIADES=1"
)

if "%PLANTILLES_CANVIADES%"=="1" (
    echo Detectats canvis locals a ESTRUCTURALS\*.docx. Els pujo a GitHub...
    git add "ESTRUCTURALS/*.docx"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Plantilles ESTRUCTURALS actualitzades des de Actualitzar.bat"
    if errorlevel 1 (
        echo  No hi havia res a commitejar a ESTRUCTURALS o el commit ha fallat. Continuo.
    )
)

REM --- 3. Si queden ALTRES canvis (codi), els guardem al stash ---
set "STASHED=0"
git diff --quiet
if errorlevel 1 set "STASHED=1"
git diff --cached --quiet
if errorlevel 1 set "STASHED=1"

if "%STASHED%"=="1" (
    echo Hi ha canvis locals a fitxers de codi. Els guardo al stash com a copia de seguretat...
    git stash push -u -m "Auto-stash per Actualitzar.bat"
    if errorlevel 1 (
        echo ERROR: no s'ha pogut fer stash. Avorto.
        pause
        exit /b 1
    )
)

REM --- 4. Anar a main i actualitzar ---
git checkout main
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut canviar a la branca 'main'.
    git status
    pause
    exit /b 1
)

REM Primer intentem fast-forward. Si tenim un commit local de plantilles
REM i el remot tambe ha avancat, fem rebase per integrar-ho ordenadament.
git pull --ff-only origin main
if errorlevel 1 (
    echo Fast-forward no possible (la branca local i la remota han divergit^).
    echo Faig rebase per integrar els canvis...
    git pull --rebase origin main
    if errorlevel 1 (
        echo.
        echo ERROR: rebase ha fallat. Probablement hi ha un conflicte.
        echo   Executa  git status  per veure-ho. Si vols avortar:  git rebase --abort
        pause
        exit /b 1
    )
)

REM --- 5. Si abans hem commitejat plantilles, pugem-les ---
if "%PLANTILLES_CANVIADES%"=="1" (
    echo Pujant les plantilles a GitHub...
    git push origin main
    if errorlevel 1 (
        echo  No s'han pogut pujar les plantilles. Es queden en local fins la propera.
    )
)

echo.
echo === Estat DESPRES ===
git branch --show-current
git log -1 --format="%%h  %%s"

if "%STASHED%"=="1" (
    echo.
    echo NOTA: hi havia canvis locals a fitxers de codi: estan al stash.
    echo   Per recuperar:    git stash pop
    echo   Per veure:        git stash list
    echo   Per descartar:    git stash drop
)

echo.
echo Fet.
pause
