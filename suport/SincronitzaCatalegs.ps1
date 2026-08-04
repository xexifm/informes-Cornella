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

  Les funcions de text son PURES i es proven en headless; la resta nomes fa
  copies de fitxers i crides a git (cap dependencia del motor d'informes).
#>

param(
    [ValidateSet('Backup', 'Restore')]
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

# Un fitxer protegit es BINARI (no es pot fusionar)? Funcio PURA.
#
# Els catalegs son .json: text. Si el repositori i l'usuari toquen el mateix
# .json, tornar a aplicar el de l'usuari es el que ha de passar (ell mana) i com
# a molt es perd un canvi de text que es pot tornar a fer.
#
# '0 CAPCALERA.docx' NO: es un ZIP amb XML a dins. Alli "guanya l'usuari" vol dir
# LLENCAR el fitxer sencer de l'altra banda, i el que hi pot haver a l'altra
# banda es una peca que el programa NECESSITA (un bloc [[CAP:...]] nou, un
# marcador <<...>>). Vegeu _CatalegHiHaColisio.
function _CatalegEsBinari([string]$path) {
    $p = ([string]$path).Replace('\', '/').Trim().ToLower()
    return ($p -like '*.docx' -or $p -like '*.doc')
}

# Hi ha COL·LISIO en un fitxer? Funcio PURA.
#
# $shaBase = el fitxer al commit on era el clone ABANS del pull (d'on va sortir
#            la copia de l'usuari)
# $shaAra  = el fitxer que hi ha al clone DESPRES del pull (el del repositori)
#
# Si son iguals, el repositori no l'ha tocat: la copia de l'usuari es pot tornar
# a aplicar tranquil·lament. Si son diferents, l'han tocat TOTS DOS i, en un
# fitxer binari, no es poden tenir les dues coses: cal aturar-se i dir-ho.
#
# Si no se sap el sha de base (cadena buida), es diu que NO hi ha col·lisio: val
# mes tornar a aplicar el de l'usuari (el comportament de sempre, que mai no li
# perd res) que arriscar-se a descartar-lo per un dubte.
function _CatalegHiHaColisio([string]$shaBase, [string]$shaAra, [bool]$esBinari) {
    if (-not $esBinari) { return $false }
    if ([string]::IsNullOrWhiteSpace($shaBase) -or [string]::IsNullOrWhiteSpace($shaAra)) { return $false }
    return ($shaBase.Trim() -ne $shaAra.Trim())
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
        # BASE: el commit on era el clone ABANS del pull, que es d'on va sortir
        # el que l'usuari te al disc. El Restore el necessita per saber si el
        # repositori ha tocat el mateix fitxer (vegeu _CatalegHiHaColisio).
        # S'apunta AQUI perque aquesta fase corre abans del commit i del pull.
        $base = ''
        try { $base = ([string](git rev-parse HEAD 2>$null)).Trim() } catch { $base = '' }
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            $base | Set-Content -LiteralPath (Join-Path $dest 'base.txt') -Encoding UTF8
        }
        $dest | Set-Content -LiteralPath (_CatalegsMarkerPath) -Encoding UTF8
        Write-Host ""
        Write-Host ("  COPIA DE SEGURETAT: " + $dest)
        return $dest
    } finally {
        Pop-Location
    }
}

# Torna a aplicar la copia al clone: la versio de l'usuari PREVAL sobre la que
# hagi baixat del repositori.
#
# EXCEPCIO: els fitxers BINARIS que hagin canviat a les DUES bandes. Vegeu
# _CatalegHiHaColisio: alli "l'usuari mana" vol dir llencar el fitxer sencer del
# repositori, i aixo ja va passar de veritat amb '0 CAPCALERA.docx' — el bloc
# [[CAP:LLIC]] que hi acabava d'entrar va desapareixer i, a sobre, la versio
# sense el bloc es va PUJAR a main. En aquest cas no es toca res: es deixa la
# del repositori (el programa queda sencer), la de l'usuari es queda a la copia
# de seguretat i s'avisa. Com que el fitxer no es modifica, el clone queda net i
# l'Actualitzar.bat no puja res: ningu no decideix per l'usuari en silenci.
#
# Retorna @{ Restaurats; Colisions } (llista de rutes).
function Invoke-CatalegsRestore {
    $buit = @{ Restaurats = 0; Colisions = @() }
    $root = _CatalegsRepoRoot
    $mk = _CatalegsMarkerPath
    if (-not (Test-Path -LiteralPath $mk)) { return $buit }
    $dest = ''
    try { $dest = (Get-Content -LiteralPath $mk -Raw -Encoding UTF8).Trim() } catch { $dest = '' }
    if ([string]::IsNullOrWhiteSpace($dest) -or -not (Test-Path -LiteralPath $dest)) { return $buit }
    $man = Join-Path $dest 'manifest.txt'
    if (-not (Test-Path -LiteralPath $man)) { return $buit }

    # El commit d'on venia el que l'usuari te al disc (l'escriu el Backup).
    $base = ''
    $baseFile = Join-Path $dest 'base.txt'
    if (Test-Path -LiteralPath $baseFile) {
        try { $base = (Get-Content -LiteralPath $baseFile -Raw -Encoding UTF8).Trim() } catch { $base = '' }
    }

    $n = 0
    $colisions = New-Object System.Collections.ArrayList
    Push-Location $root
    try {
        foreach ($rel in @(Get-Content -LiteralPath $man -Encoding UTF8)) {
            $r = ([string]$rel).Trim()
            if ([string]::IsNullOrWhiteSpace($r)) { continue }
            $src = Join-Path $dest ($r -replace '/', '\')
            $dst = Join-Path $root ($r -replace '/', '\')
            if (-not (Test-Path -LiteralPath $src)) { continue }

            # El repositori ha tocat el mateix fitxer binari? Es compara pel sha
            # del blob (res de llegir binaris amb PowerShell): el de la base
            # contra el que hi ha ARA al clone, que ja ve del pull.
            $esBin = _CatalegEsBinari $r
            if ($esBin -and -not [string]::IsNullOrWhiteSpace($base)) {
                $shaBase = ''
                $shaAra  = ''
                try { $shaBase = ([string](git rev-parse ("{0}:{1}" -f $base, $r) 2>$null)).Trim() } catch { $shaBase = '' }
                try { $shaAra  = ([string](git hash-object -- $dst 2>$null)).Trim() } catch { $shaAra = '' }
                if (_CatalegHiHaColisio $shaBase $shaAra $true) {
                    [void]$colisions.Add($r)
                    Write-Host ""
                    Write-Host ("  ATENCIO: '" + $r + "' ha canviat a les DUES bandes.")
                    Write-Host  "    Es queda la versio del repositori (el programa la necessita sencera)."
                    Write-Host ("    La TEVA es guarda a: " + $src)
                    Write-Host  "    No es puja res: passa-li la teva copia a en Claude i te l'ajuntara."
                    Write-Host ""
                    continue
                }
            }

            $dstDir = Split-Path -Parent $dst
            if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
            Copy-Item -LiteralPath $src -Destination $dst -Force
            $n++
            Write-Host ("  restaurat: " + $r)
        }
    } finally {
        Pop-Location
    }
    return @{ Restaurats = $n; Colisions = $colisions.ToArray() }
}

# ----------------------------------------------------------------------------
# Execucio (en mode proves nomes es volien les definicions)
# ----------------------------------------------------------------------------
if ($env:GENINFORME_TEST -eq '1') { return }

switch ($Fase) {
    'Backup'  { [void](Invoke-CatalegsBackup) }
    'Restore' {
        $res = Invoke-CatalegsRestore
        # Codi 2 = hi ha hagut col·lisio. L'Actualitzar.bat el mira per tornar a
        # dir-ho al FINAL, que es on l'usuari mira: aqui enmig es perd amunt.
        if (@($res.Colisions).Count -gt 0) { exit 2 }
    }
}
