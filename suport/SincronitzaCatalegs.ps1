#requires -Version 5.1
<#
.SYNOPSIS
  Protegeix els CATALEGS (i les dades editables des del programa) quan
  Actualitzar.bat fa git: copia de seguretat, restauracio i recuperacio.

.DESCRIPTION
  PER QUE EXISTEIX: Actualitzar.bat detectava canvis a ESTRUCTURALS pero nomes
  feia 'git add ESTRUCTURALS/*.docx'. Els .json dels catalegs (els que escriu
  l'editor "Editar catalegs") no es commitejaven mai, quedaven bruts, i el
  'git stash push -u' del pas seguent se'ls enduia. El pull restaurava la versio
  del repositori i la feina de l'usuari DESAPAREIXIA del programa (encara que
  quedava al stash, ningu no la treia mai d'alli).

  Regla nova: els catalegs editats a l'ordinador de l'usuari SON L'AUTORITAT.
  Es copien abans de tocar res de git, es tornen a aplicar despres del pull
  (per tant PREVALEN sobre el que baixi) i es committegen i pugen a GitHub.

  Fases:
    -Fase Backup     copia els fitxers modificats/nous de les rutes protegides a
                     %LOCALAPPDATA%\InformesCornella\backups\<data-hora>\ i deixa
                     la ruta a 'ultima-copia-catalegs.txt'.
    -Fase Restore    torna a copiar aquells fitxers al clone (la versio de
                     l'usuari preval sobre la que hagi baixat del repositori).
    -Fase Recuperar  busca als STASHES antics els catalegs que es van perdre amb
                     la versio anterior d'Actualitzar.bat i els extreu a
                     %LOCALAPPDATA%\InformesCornella\recuperats\ (NO fa 'pop':
                     no toca res del clone).

  Les funcions de text son PURES i es proven en headless; la resta nomes fa
  copies de fitxers i crides a git (cap dependencia del motor d'informes).
#>

param(
    [ValidateSet('Backup', 'Restore', 'Recuperar')]
    [string]$Fase = 'Backup'
)

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables)
# ----------------------------------------------------------------------------

# Rutes PROTEGIDES: tot el que l'usuari pot editar des del programa i que no s'ha
# de perdre mai en actualitzar.
function _CatalegsRutesProtegides {
    return @('ESTRUCTURALS', 'docs/dades')
}

# Un fitxer s'ha de protegir? Els .bak els crea l'editor en desar (copia de
# seguretat local) i no s'han de pujar mai.
function _CatalegEsProtegible([string]$path) {
    $p = ([string]$path).Replace('\', '/').Trim()
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }
    if ($p -like '*.bak') { return $false }
    if ($p -like '*~$*')  { return $false }   # temporals de Word
    foreach ($r in (_CatalegsRutesProtegides)) {
        if ($p -like ($r + '/*')) { return $true }
    }
    return $false
}

# Converteix la sortida de 'git status --porcelain' en la llista de rutes
# relatives protegides. Format de cada linia: XY<espai>RUTA (o 'R  vell -> nou').
# Funcio PURA.
# IMPORTANT: retorna un ARRAY PLA (no ',$ArrayList'): amb el ',' l'@() del crider
# no l'enumera i rep la llista sencera com a unic element (ja ens va passar amb
# _AutoFirmaCandidatePaths).
function _ParseGitStatusPaths($lines) {
    $out = New-Object System.Collections.ArrayList
    foreach ($ln in @($lines)) {
        $l = [string]$ln
        if ($l.Length -lt 4) { continue }
        $path = $l.Substring(3).Trim()
        # Canvi de nom: 'vell -> nou'; ens quedem amb el desti.
        $idx = $path.IndexOf(' -> ')
        if ($idx -ge 0) { $path = $path.Substring($idx + 4).Trim() }
        # git enquota les rutes amb caracters especials.
        if ($path.StartsWith('"') -and $path.EndsWith('"') -and $path.Length -ge 2) {
            $path = $path.Substring(1, $path.Length - 2)
        }
        $path = $path.Replace('\', '/')
        if (_CatalegEsProtegible $path) { [void]$out.Add($path) }
    }
    return $out.ToArray()
}

# Nom de la carpeta de copia per a un moment donat. Funcio PURA.
function _CatalegsBackupName([datetime]$now) {
    return $now.ToString('yyyyMMdd-HHmmss')
}

# ----------------------------------------------------------------------------
# Rutes de treball
# ----------------------------------------------------------------------------
function _CatalegsBaseDir {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path $base 'InformesCornella')
}
function _CatalegsBackupRoot   { return (Join-Path (_CatalegsBaseDir) 'backups') }
function _CatalegsRecuperatDir { return (Join-Path (_CatalegsBaseDir) 'recuperats') }
function _CatalegsMarkerPath   { return (Join-Path (_CatalegsBaseDir) 'ultima-copia-catalegs.txt') }

# Arrel del clone (aquest script viu a <clone>\suport\).
function _CatalegsRepoRoot {
    if ($PSScriptRoot) { return (Split-Path -Parent $PSScriptRoot) }
    return (Get-Location).Path
}

# ----------------------------------------------------------------------------
# FASES
# ----------------------------------------------------------------------------

# Copia els fitxers protegits que estiguin modificats/nous. Retorna la ruta de la
# copia (o '' si no hi havia res a copiar).
function Invoke-CatalegsBackup {
    $root = _CatalegsRepoRoot
    Push-Location $root
    try {
        $status = @(git status --porcelain -- ESTRUCTURALS docs/dades 2>$null)
        $fitxers = @(_ParseGitStatusPaths $status)
        if ($fitxers.Count -eq 0) {
            # Cap canvi: esborrem el marcador perque el Restore no faci res.
            $mk = _CatalegsMarkerPath
            if (Test-Path -LiteralPath $mk) { Remove-Item -LiteralPath $mk -Force -ErrorAction SilentlyContinue }
            Write-Host "  (cap canvi local als catalegs)"
            return ''
        }

        $dest = Join-Path (_CatalegsBackupRoot) (_CatalegsBackupName (Get-Date))
        if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }

        $copiats = New-Object System.Collections.ArrayList
        foreach ($rel in $fitxers) {
            $src = Join-Path $root ($rel -replace '/', '\')
            if (-not (Test-Path -LiteralPath $src)) { continue }   # esborrat: no hi ha res a copiar
            $dst = Join-Path $dest ($rel -replace '/', '\')
            $dstDir = Split-Path -Parent $dst
            if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
            Copy-Item -LiteralPath $src -Destination $dst -Force
            [void]$copiats.Add($rel)
            Write-Host ("  copiat: " + $rel)
        }

        if ($copiats.Count -eq 0) {
            Write-Host "  (cap fitxer per copiar)"
            return ''
        }
        # Manifest: la llista exacta que s'ha de tornar a aplicar.
        ($copiats -join "`r`n") | Set-Content -LiteralPath (Join-Path $dest 'manifest.txt') -Encoding UTF8
        $dest | Set-Content -LiteralPath (_CatalegsMarkerPath) -Encoding UTF8
        Write-Host ""
        Write-Host ("  COPIA DE SEGURETAT: " + $dest)
        return $dest
    } finally {
        Pop-Location
    }
}

# Torna a aplicar la copia al clone: la versio de l'usuari PREVAL sobre la que
# hagi baixat del repositori. Retorna el nombre de fitxers restaurats.
function Invoke-CatalegsRestore {
    $root = _CatalegsRepoRoot
    $mk = _CatalegsMarkerPath
    if (-not (Test-Path -LiteralPath $mk)) { return 0 }
    $dest = ''
    try { $dest = (Get-Content -LiteralPath $mk -Raw -Encoding UTF8).Trim() } catch { $dest = '' }
    if ([string]::IsNullOrWhiteSpace($dest) -or -not (Test-Path -LiteralPath $dest)) { return 0 }
    $man = Join-Path $dest 'manifest.txt'
    if (-not (Test-Path -LiteralPath $man)) { return 0 }

    $n = 0
    foreach ($rel in @(Get-Content -LiteralPath $man -Encoding UTF8)) {
        $r = ([string]$rel).Trim()
        if ([string]::IsNullOrWhiteSpace($r)) { continue }
        $src = Join-Path $dest ($r -replace '/', '\')
        $dst = Join-Path $root ($r -replace '/', '\')
        if (-not (Test-Path -LiteralPath $src)) { continue }
        $dstDir = Split-Path -Parent $dst
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $src -Destination $dst -Force
        $n++
        Write-Host ("  restaurat: " + $r)
    }
    return $n
}

# Busca als stashes els catalegs que es van perdre amb la versio anterior
# d'Actualitzar.bat i els extreu (sense tocar el clone ni els stashes).
function Invoke-CatalegsRecuperar {
    $root = _CatalegsRepoRoot
    Push-Location $root
    try {
        $stashes = @(git stash list --format=%gd 2>$null)
        if ($stashes.Count -eq 0) {
            Write-Host "No hi ha cap stash: no hi ha res a recuperar."
            return
        }
        $outRoot = _CatalegsRecuperatDir
        $trobats = 0
        foreach ($ref in $stashes) {
            $r = ([string]$ref).Trim()
            if ([string]::IsNullOrWhiteSpace($r)) { continue }
            $noms = @(git stash show --name-only $r 2>$null)
            $prot = @()
            foreach ($n in $noms) { if (_CatalegEsProtegible ([string]$n)) { $prot += ([string]$n).Trim() } }
            if ($prot.Count -eq 0) { continue }

            $safe = ($r -replace '[^A-Za-z0-9]', '_')
            $dir = Join-Path $outRoot $safe
            foreach ($rel in $prot) {
                $dst = Join-Path $dir ($rel -replace '/', '\')
                $dstDir = Split-Path -Parent $dst
                if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
                # El contingut del fitxer TAL COM estava al stash.
                $txt = (git show ("{0}:{1}" -f $r, $rel) 2>$null | Out-String)
                if ([string]::IsNullOrWhiteSpace($txt)) { continue }
                [System.IO.File]::WriteAllText($dst, $txt, (New-Object System.Text.UTF8Encoding($false)))
                $trobats++
                Write-Host ("  {0}  ->  {1}" -f $r, $rel)
            }
        }
        Write-Host ""
        if ($trobats -eq 0) {
            Write-Host "No s'ha trobat cap cataleg als stashes."
        } else {
            Write-Host ("S'han extret {0} fitxers a:" -f $trobats)
            Write-Host ("  " + $outRoot)
            Write-Host ""
            Write-Host "Obre'ls per veure els teus canvis. Si en vols recuperar algun,"
            Write-Host "copia'l a la carpeta ESTRUCTURALS del programa i executa Actualitzar.bat."
        }
    } finally {
        Pop-Location
    }
}

# ----------------------------------------------------------------------------
# Execucio (en mode proves nomes es volien les definicions)
# ----------------------------------------------------------------------------
if ($env:GENINFORME_TEST -eq '1') { return }

switch ($Fase) {
    'Backup'    { [void](Invoke-CatalegsBackup) }
    'Restore'   { [void](Invoke-CatalegsRestore) }
    'Recuperar' { Invoke-CatalegsRecuperar }
}
