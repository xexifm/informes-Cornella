@echo off
REM Actualitza el generador d'informes Cornella.
REM - Si tens canvis locals a ESTRUCTURALS\*.docx (plantilles), els puja a GitHub.
REM - Si tens canvis a codi (.ps1, .bat) els guarda al stash com a copia de
REM   seguretat i continua. No es perden mai.
REM - Sempre acaba a 'main' amb l'ultim commit i mostra resultat.
REM - En acabar, torna a obrir el programa (actualitzat).
REM
REM Doble clic per executar.

setlocal EnableDelayedExpansion
title Actualitzar generador d'informes Cornella
cd /d "%~dp0"

REM --- Tancar el programa si esta obert i ESPERAR que es tanqui del tot ---
REM El programa desa el PID del seu proces viu a running.pid. IMPORTANT: quan
REM s'obre des del boto "Actualitzar" del programa, aquest .bat s'executa com a
REM proces FILL del programa, per aixo el taskkill va SENSE /T: amb /T ens
REM matariem a nosaltres mateixos (a tota la branca de processos) i
REM l'actualitzacio no arribava a executar-se (el programa es tancava pero no
REM s'actualitzava res). Despres esperem que el proces desaparegui del tot per no
REM trobar fitxers bloquejats en fer git.
set "PIDFILE=%LOCALAPPDATA%\InformesCornella\running.pid"
set "GENPID="
if exist "%PIDFILE%" set /p GENPID=<"%PIDFILE%"
if defined GENPID call :TancaGenerador "%GENPID%"
if exist "%PIDFILE%" del "%PIDFILE%" >nul 2>&1

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
    goto :ERROR
)

REM --- 1. Fetch ---
echo Connectant a GitHub...
git fetch origin
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut connectar a GitHub. Comprova la xarxa.
    goto :ERROR
)

REM --- 2. Hi ha canvis als CATALEGS/plantilles d'ESTRUCTURALS ? Els pugem ---
REM    (codi ps1/bat NO el pujem mai des d'aqui: el toca Claude o tu via PR)
REM ESTRUCTURALS conte les plantilles .docx I ELS CATALEGS .json (els que escriu
REM l'editor "Editar catalegs"). ATENCIO: abans aqui nomes es feia
REM 'git add ESTRUCTURALS/*.docx'; els .json quedaven bruts, se'ls enduia el
REM 'git stash' del pas 3 i la feina de l'usuari DESAPAREIXIA del programa. Ara
REM s'afegeixen TAMBE els .json i, a mes, es fa una COPIA DE SEGURETAT abans de
REM tocar res de git (i es torna a aplicar despres del pull: la versio de
REM l'usuari PREVAL). Evitem 'findstr' amb ancora "$": el git treu linies
REM acabades en LF i el findstr de Windows falla amb l'ancora de final de linia.
set "PLANTILLES_CANVIADES=0"
for /f "delims=" %%f in ('git status --porcelain -- ESTRUCTURALS') do set "PLANTILLES_CANVIADES=1"

REM COPIA DE SEGURETAT de TOT el que l'usuari pot editar des del programa
REM (ESTRUCTURALS + docs\dades). Es fa SEMPRE i ABANS de tocar res de git, aixi
REM passi el que passi (stash, conflicte, rebase) la feina queda desada en disc.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\SincronitzaCatalegs.ps1" -Fase Backup

if "%PLANTILLES_CANVIADES%"=="1" (
    echo Detectats canvis locals als catalegs/plantilles ESTRUCTURALS.
    echo Regenerant les dades del mobil ^(docs\dades^) des dels catalegs...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\mobil\ExportaDades.ps1" -Plantilles
    if errorlevel 1 (
        echo  Avis: no s'han pogut regenerar les dades del mobil ^(Word obert o no instal·lat?^). Continuo igualment.
    )
    echo Pujant catalegs, plantilles i dades del mobil a GitHub...
    git add "ESTRUCTURALS/*.docx" "ESTRUCTURALS/*.json" "docs/dades/*.json"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Catalegs, plantilles i dades del mobil actualitzats des de Actualitzar.bat"
    if errorlevel 1 (
        echo  No hi havia res a commitejar o el commit ha fallat. Continuo.
    )
)

REM --- 2b. Canvis als textos del correu del mobil (editats des de l'app)? ---
REM    docs\dades\email-textos.json es tracked; si l'usuari l'ha editat des de
REM    l'eina "Textos del correu", el commitegem AQUI (abans del stash; si no,
REM    aniria a parar al stash i no es publicaria).
set "TEXTOS_CANVIATS=0"
for /f "delims=" %%f in ('git status --porcelain -- docs/dades/email-textos.json') do set "TEXTOS_CANVIATS=1"
if "%TEXTOS_CANVIATS%"=="1" (
    echo Detectats canvis als textos del correu del mobil.
    git add "docs/dades/email-textos.json"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Textos del correu del mobil actualitzats des de l'app"
    if errorlevel 1 (
        echo  No hi havia res a commitejar o el commit ha fallat. Continuo.
    )
)

REM Hi havia canvis en alguna cosa PROTEGIDA (catalegs o textos del correu)? Si
REM es aixi, despres del pull els tornarem a aplicar des de la copia (pas 4b).
set "PROTEGITS_CANVIATS=0"
if "%PLANTILLES_CANVIADES%"=="1" set "PROTEGITS_CANVIATS=1"
if "%TEXTOS_CANVIATS%"=="1" set "PROTEGITS_CANVIATS=1"

REM --- 3. Si queden ALTRES canvis (codi), els guardem al stash ---
REM    Abans, pero, comprovem que NO quedi res brut a ESTRUCTURALS: si en
REM    quedes, el stash se l'enduria i l'usuari el perdria de vista (aixo es
REM    exactament el que passava abans). Si passa, avisem ben clar; la copia de
REM    seguretat del pas 2 ja el te desat igualment.
set "CATALEGS_BRUTS=0"
for /f "delims=" %%f in ('git status --porcelain -- ESTRUCTURALS') do set "CATALEGS_BRUTS=1"
if "%CATALEGS_BRUTS%"=="1" (
    echo.
    echo  ATENCIO: encara hi ha canvis a ESTRUCTURALS que no s'han pogut commitejar.
    echo    Tens una copia de seguretat a:
    echo      %LOCALAPPDATA%\InformesCornella\backups
    echo    Es tornaran a aplicar despres d'actualitzar.
    echo.
)

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
        goto :ERROR
    )
)

REM --- 4. Anar a main i actualitzar ---
git checkout main
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut canviar a la branca 'main'.
    git status
    goto :ERROR
)

REM Primer intentem fast-forward. Si tenim un commit local de plantilles
REM i el remot tambe ha avancat, fem rebase per integrar-ho ordenadament.
git pull --ff-only origin main
if errorlevel 1 (
    echo Fast-forward no possible ^(la branca local i la remota han divergit^).
    echo Faig rebase per integrar els canvis...
    git pull --rebase origin main
    if errorlevel 1 (
        echo.
        echo  El rebase ha trobat un conflicte. Els TEUS catalegs manen:
        echo    avorto el rebase, em poso al dia i els torno a aplicar.
        git rebase --abort
        git reset --hard origin/main
    )
)

REM --- 4b. Tornar a aplicar els catalegs de l'usuari (PREVALEN) ---
REM    Despres del pull, els catalegs que hagin baixat es substitueixen pels que
REM    tenia l'usuari (copia del pas 2). Aixi la seva feina mai no es perd, ni
REM    quan el repositori ha tocat el mateix fitxer. Si el resultat difereix del
REM    que hi ha al repositori, es committeja i es puja.
if "%PROTEGITS_CANVIATS%"=="1" (
    echo.
    echo Tornant a aplicar els teus catalegs ^(la teva versio preval^)...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0suport\SincronitzaCatalegs.ps1" -Fase Restore
    git add "ESTRUCTURALS/*.docx" "ESTRUCTURALS/*.json" "docs/dades/*.json"
    git -c user.name="Generador d'informes" -c user.email="generador@local" commit -m "Catalegs de l'usuari mantinguts des de Actualitzar.bat"
    if errorlevel 1 (
        echo  Els catalegs ja estaven al dia ^(res a commitejar^).
    )
)

REM --- 5. Pujar a GitHub si el 'main' local va per davant de l'origin ---
REM    (commits fets per aquest .bat: plantilles, textos del correu, o algun de
REM    pendent d'una execucio anterior que no es va poder pujar).
set "PENDENTS=1"
for /f %%c in ('git rev-list --count origin/main..HEAD') do set "PENDENTS=%%c"
if not "%PENDENTS%"=="0" (
    echo Pujant els canvis locals a GitHub...
    git push origin main
    if errorlevel 1 (
        echo  No s'han pogut pujar els canvis. Es queden en local fins la propera.
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
echo.
echo Prem qualsevol tecla per tancar aquesta finestra i obrir el programa actualitzat...
pause >nul
call :ReobrePrograma
exit

REM ==========================================================================
REM  Sortida amb ERROR: no s'ha actualitzat. Tornem a obrir el programa (amb la
REM  versio actual, sense canvis) perque l'usuari no es quedi sense el programa.
REM ==========================================================================
:ERROR
echo.
echo Prem qualsevol tecla per tancar aquesta finestra i tornar a obrir el programa...
pause >nul
call :ReobrePrograma
exit /b 1

REM ==========================================================================
REM  Subrutines
REM ==========================================================================

REM Tanca el generador (PID a %1) SENSE /T i espera fins a ~20s que desaparegui.
:TancaGenerador
set "_p=%~1"
tasklist /FI "PID eq %_p%" /FI "IMAGENAME eq powershell.exe" 2>nul | find "%_p%" >nul
if errorlevel 1 goto :eof
echo Tancant el generador que esta obert ^(PID %_p%^)...
taskkill /PID %_p% /F >nul 2>&1
set /a _n=0
:tg_wait
tasklist /FI "PID eq %_p%" /FI "IMAGENAME eq powershell.exe" 2>nul | find "%_p%" >nul
if errorlevel 1 goto :eof
set /a _n+=1
if %_n% geq 20 goto :eof
ping -n 2 127.0.0.1 >nul
goto :tg_wait

REM Torna a obrir el programa (via el .vbs, sense finestra de consola).
:ReobrePrograma
start "" wscript.exe "%~dp0suport\GenerarInforme.vbs"
goto :eof
