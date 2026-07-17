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

REM --- Tancar el programa si esta obert ---
REM Abans d'actualitzar tanquem el generador si s'esta executant (si no, podria
REM quedar amb la versio antiga carregada o bloquejar fitxers). El programa desa
REM el PID del seu proces viu a running.pid; el llegim i, si es realment un
REM powershell.exe amb aquest PID, el tanquem.
set "PIDFILE=%LOCALAPPDATA%\InformesCornella\running.pid"
if exist "%PIDFILE%" (
    set "GENPID="
    set /p GENPID=<"%PIDFILE%"
    if defined GENPID (
        tasklist /FI "PID eq !GENPID!" /FI "IMAGENAME eq powershell.exe" 2>nul | find "!GENPID!" >nul
        if not errorlevel 1 (
            echo Tancant el generador que esta obert ^(PID !GENPID!^)...
            taskkill /PID !GENPID! /T /F >nul 2>&1
        )
        del "%PIDFILE%" >nul 2>&1
    )
)

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
REM ESTRUCTURALS nomes conte plantilles .docx, aixi que QUALSEVOL canvi aqui es
REM un canvi de plantilla. Evitem 'findstr' amb ancora "$": el git treu linies
REM acabades en LF i el findstr de Windows falla amb l'ancora de final de linia
REM sobre LF (per aixo abans no detectava els .docx i anaven a parar al stash).
set "PLANTILLES_CANVIADES=0"
for /f "delims=" %%f in ('git status --porcelain -- ESTRUCTURALS') do set "PLANTILLES_CANVIADES=1"

if "%PLANTILLES_CANVIADES%"=="1" (
    echo Detectats canvis locals a ESTRUCTURALS\*.docx.
    echo Regenerant les dades del mobil ^(docs\dades^) des de les plantilles...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\mobil\ExportaDades.ps1" -Plantilles
    if errorlevel 1 (
        echo  Avis: no s'han pogut regenerar les dades del mobil ^(Word obert o no instal·lat?^). Continuo igualment.
    )
    echo Pujant plantilles i dades del mobil a GitHub...
    git add "ESTRUCTURALS/*.docx" "docs/dades/*.json"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Plantilles + dades del mobil actualitzades des de Actualitzar.bat"
    if errorlevel 1 (
        echo  No hi havia res a commitejar o el commit ha fallat. Continuo.
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

REM --- 6. Refrescar la base de dades d'activitats al Drive (per al mobil) ---
REM    Agafa l'Excel mes recent (xarxa I: si hi ets, o la carpeta local
REM    BASE DE DADES ACTIVITATS) i el puja a Drive. Necessita haver autoritzat
REM    el PC (Authorize-Drive.bat) i tenir Excel. Es fail-safe: si no pot, avisa.
echo.
echo Refrescant la base de dades d'activitats al Drive (per al mobil)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\mobil\ExportaDades.ps1" -Activitats
if errorlevel 1 (
    echo  Avis: no s'han pogut refrescar les activitats al Drive. Continuo.
)

REM --- 7. Refrescar el mapa public d'activitats precintades ---
REM    Regenera docs\dades\precintades.json des de l'ultim Excel i, si ha
REM    canviat, el puja a 'main' (GitHub Pages el serveix public). El JSON NO
REM    conte dades personals (nomes activitat generica, adreca i coordenades).
REM    Fail-safe: si no pot (sense Excel, etc.), avisa i continua.
echo.
echo Refrescant el mapa public d'activitats precintades ^(docs\dades^)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\rutes\Precintades.ps1"
if errorlevel 1 (
    echo  Avis: no s'ha pogut regenerar el mapa de precintades. Continuo.
) else (
    git add "docs/dades/precintades.json"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Mapa d'activitats precintades actualitzat des de Actualitzar.bat"
    if errorlevel 1 (
        echo  El mapa de precintades ja estava al dia ^(res a commitejar^).
    ) else (
        echo Pujant el mapa de precintades a GitHub...
        git push origin main
        if errorlevel 1 echo  No s'ha pogut pujar el mapa. Es queda en local fins la propera.
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
