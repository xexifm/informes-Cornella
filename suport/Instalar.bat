@echo off
REM ============================================================
REM  INSTAL-LADOR del generador d'informes Cornella
REM ------------------------------------------------------------
REM  Per a un ordinador NOU i net. Deixa el programa operatiu:
REM    1. Instal-la Git si no hi es (winget o descarrega directa).
REM    2. Baixa el programa de GitHub a la branca 'main'.
REM    3. A punt per fer servir GenerarInforme.bat / Actualitzar.bat.
REM
REM  Funciona en 3 situacions:
REM    a) Has extret el ZIP de GitHub -> converteix la carpeta en clone.
REM    b) Algu t'ha enviat NOMES aquest .bat -> clona tot a
REM       'informes-Cornella'.
REM    c) Ja tens un clone -> l'actualitza a 'main'.
REM
REM  Doble clic per executar.
REM ============================================================

setlocal EnableDelayedExpansion
title Instal-lar generador d'informes Cornella
cd /d "%~dp0"

set "REPO_URL=https://github.com/xexifm/informes-cornella"

echo ============================================================
echo  Instal-lador del generador d'informes Cornella
echo ============================================================
echo.

REM ------------------------------------------------------------
REM  1. Localitzar / instal-lar Git
REM ------------------------------------------------------------
call :FIND_GIT
if defined GIT goto GIT_OK

echo [Git] No s'ha trobat Git. Provant d'instal-lar-lo...
echo.

where winget >nul 2>&1
if errorlevel 1 goto GIT_TRY_DOWNLOAD
echo [Git] Instal-lant amb winget...
winget install -e --id Git.Git --accept-package-agreements --accept-source-agreements
call :FIND_GIT
if defined GIT goto GIT_OK

:GIT_TRY_DOWNLOAD
echo [Git] Provant descarrega directa de Git per a Windows...
set "GIT_SETUP=%TEMP%\git-setup-cornella.exe"
if exist "%GIT_SETUP%" del "%GIT_SETUP%" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; try { $out = Join-Path $env:TEMP 'git-setup-cornella.exe'; $rel = Invoke-RestMethod 'https://api.github.com/repos/git-for-windows/git/releases/latest' -Headers @{'User-Agent'='cornella'}; $a = $rel.assets | Where-Object { $_.name -match '64-bit\.exe$' -and $_.name -notmatch 'Portable|busybox|rc|arm' } | Select-Object -First 1; if (-not $a) { exit 2 }; Invoke-WebRequest $a.browser_download_url -OutFile $out; exit 0 } catch { exit 3 }"
if not exist "%GIT_SETUP%" goto GIT_MANUAL
echo [Git] Instal-lant en silenci...
"%GIT_SETUP%" /VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS
del "%GIT_SETUP%" >nul 2>&1
call :FIND_GIT
if defined GIT goto GIT_OK

:GIT_MANUAL
echo.
echo ============================================================
echo  No s'ha pogut instal-lar Git automaticament.
echo  S'obrira la pagina de descarrega. Instal-la Git i torna a
echo  executar aquest fitxer (Instalar.bat).
echo ============================================================
start "" "https://git-scm.com/download/win"
pause
exit /b 1

:GIT_OK
echo [Git] Disponible: !GIT!
echo.

REM ------------------------------------------------------------
REM  2. Detectar l'escenari
REM     ROOT = arrel del repo (on hi ha GenerarInforme.bat).
REM     Pot ser aquesta carpeta o la carpeta pare (si el .bat es a suport\).
REM ------------------------------------------------------------
set "ROOT="
if exist "%~dp0GenerarInforme.bat"   set "ROOT=%~dp0"
if not defined ROOT if exist "%~dp0..\GenerarInforme.bat" set "ROOT=%~dp0.."

if not defined ROOT goto CLONE_FRESH

pushd "%ROOT%"
if exist ".git\." goto UPDATE_CLONE
goto ADOPT_ZIP

:UPDATE_CLONE
echo [Repo] Ja es un clone. Actualitzant a 'main'...
call :NO_MAINTENANCE
"%GIT%" fetch origin
if errorlevel 1 goto NET_ERR
"%GIT%" checkout main
"%GIT%" pull --ff-only origin main
if errorlevel 1 (
    echo Fast-forward no possible. Provant rebase...
    "%GIT%" pull --rebase origin main
)
set "FINALDIR=%CD%"
popd
goto OFFICE_CHECK

:ADOPT_ZIP
echo [Repo] Carpeta del programa sense Git. La converteixo en clone...
"%GIT%" init
call :NO_MAINTENANCE
"%GIT%" remote remove origin >nul 2>&1
"%GIT%" remote add origin "%REPO_URL%"
"%GIT%" fetch origin
if errorlevel 1 goto NET_ERR
"%GIT%" checkout -f -B main origin/main
"%GIT%" branch --set-upstream-to=origin/main main >nul 2>&1
set "FINALDIR=%CD%"
popd
goto OFFICE_CHECK

:CLONE_FRESH
echo [Repo] Clonant el programa des de GitHub...
"%GIT%" clone "%REPO_URL%" informes-Cornella
if errorlevel 1 (
    echo ERROR: el clone ha fallat. Comprova la xarxa.
    pause
    exit /b 1
)
set "FINALDIR=%~dp0informes-Cornella"
pushd "%FINALDIR%"
call :NO_MAINTENANCE
popd
goto OFFICE_CHECK

:NET_ERR
echo.
echo ERROR: no s'ha pogut connectar a GitHub. Comprova la xarxa.
popd
pause
exit /b 1

REM ------------------------------------------------------------
REM  3. Comprovacio suau d'Office (no bloqueja)
REM ------------------------------------------------------------
:OFFICE_CHECK
echo.
reg query "HKCR\Word.Application" >nul 2>&1
if errorlevel 1 (
    echo [Avis] No s'ha detectat Microsoft Word en aquest equip.
    echo        El programa NECESSITA Word i Excel instal-lats. No es
    echo        poden instal-lar automaticament per llicencia.
    echo.
)

REM ------------------------------------------------------------
REM  4. Acces directe (per poder ancorar el programa a la barra de tasques)
REM ------------------------------------------------------------
REM  Windows no deixa ancorar un .bat: cal un acces directe que apunti a un
REM  executable. El deixa a l'escriptori i al menu Inici, amb l'escut.
if defined FINALDIR (
  echo [Acces directe] Creant l'acces directe...
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ". '%FINALDIR%\suport\AccesDirecte.ps1'; Invoke-CrearAccesDirecte '%FINALDIR%' | Out-Null"
)

REM ------------------------------------------------------------
REM  5. Final
REM ------------------------------------------------------------
echo ============================================================
echo  Fet. El programa esta a:
echo    %FINALDIR%
echo.
echo  A partir d'ara, dins d'aquesta carpeta:
echo    - GenerarInforme.bat      =^> generar un informe
echo    - Actualitzar.bat         =^> baixar l'ultima versio
echo    - Crear-acces-directe.bat =^> refer l'acces directe
echo.
echo  Tens un acces directe a l'escriptori i al menu Inici. Per ancorar-lo
echo  a la barra de tasques: clic dret damunt seu =^> "Ancorar a la barra
echo  de tasques" (al Windows 11, potser abans "Mostra mes opcions").
echo ============================================================
echo.

choice /c SN /n /m "Vols obrir la carpeta ara? (S/N): "
if not errorlevel 2 if defined FINALDIR start "" "%FINALDIR%"

echo.
pause
exit /b 0

REM ============================================================
REM  Subrutina: localitza git.exe i el desa a la variable GIT
REM ============================================================
:FIND_GIT
set "GIT="
for /f "delims=" %%g in ('where git 2^>nul') do if not defined GIT set "GIT=%%g"
if not defined GIT if exist "%ProgramFiles%\Git\cmd\git.exe" set "GIT=%ProgramFiles%\Git\cmd\git.exe"
if not defined GIT if exist "%ProgramFiles(x86)%\Git\cmd\git.exe" set "GIT=%ProgramFiles(x86)%\Git\cmd\git.exe"
if not defined GIT if exist "%LocalAppData%\Programs\Git\cmd\git.exe" set "GIT=%LocalAppData%\Programs\Git\cmd\git.exe"
goto :eof

REM ============================================================
REM  Subrutina: desactiva el manteniment automatic del git al clone
REM ------------------------------------------------------------
REM  Quan el clone viu en una unitat de XARXA, el "geometric-repack" que el git
REM  llanca sol despres d'un fetch/push no pot reanomenar el fitxer .idx (SMB el
REM  te bloquejat) i deixa errors "Permission denied". A sobre PREGUNTA
REM  "Should I try again? (y/n)", cosa que pot deixar l'instal-lador o
REM  l'Actualitzar.bat ATURATS esperant una tecla. El repositori es petit i el
REM  manteniment no cal, aixi que el desactivem (nomes en aquest clone).
REM ============================================================
:NO_MAINTENANCE
"%GIT%" config maintenance.auto false >nul 2>&1
"%GIT%" config gc.auto 0 >nul 2>&1
goto :eof
