#requires -Version 5.1
<#
.SYNOPSIS
  Vigilant: genera automaticament els informes que arriben del mobil via Drive.

.DESCRIPTION
  Vigila la carpeta $DriveEntradaDir (sincronitzada amb el Google Drive
  d'escriptori). Quan hi apareix un paquet *.json (preparat al mobil), crida
  GenerarInforme.ps1 -DesDePaquet per generar el .docx complet, i mou el paquet
  a $DriveProcessatsDir. Aixi, quan arribes al PC, l'informe ja esta fet a
  'Informes generats'.

  Pensat per deixar-lo obert en segon pla al PC (Vigilant.bat). Fa polling cada
  pocs segons (mes robust que FileSystemWatcher sobre carpetes sincronitzades).

.PARAMETER IntervalSec
  Segons entre comprovacions (per defecte 10).

.PARAMETER Once
  Processa els paquets pendents un sol cop i surt (per a proves o per a una
  tasca programada de Windows).
#>
param(
    [int]$IntervalSec = 10,
    [switch]$Once
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem GenerarInforme.ps1 en headless NOMES per obtenir les rutes de Drive
# ($DriveEntradaDir, $DriveProcessatsDir...). La generacio real es fa cridant el
# .ps1 com a proces independent (amb -DesDePaquet), per aïllar cada informe.
$env:GENINFORME_TEST = '1'
. (Join-Path $ScriptRoot 'GenerarInforme.ps1')
Remove-Item Env:\GENINFORME_TEST -ErrorAction SilentlyContinue

$GenerarPs1 = Join-Path $ScriptRoot 'GenerarInforme.ps1'

function _EnsureDir($d) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

function Process-Pending {
    _EnsureDir $DriveEntradaDir
    _EnsureDir $DriveProcessatsDir
    $pending = Get-ChildItem -LiteralPath $DriveEntradaDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime
    foreach ($f in $pending) {
        Write-Host ("[{0}] Paquet detectat: {1}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name)
        try {
            # Generem en un proces a part amb el mode paquet. -ExecutionPolicy
            # Bypass perque funcioni sense tocar politiques.
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $GenerarPs1 -DesDePaquet $f.FullName
            if ($LASTEXITCODE -ne 0) { throw "GenerarInforme.ps1 ha retornat codi $LASTEXITCODE" }

            $dest = Join-Path $DriveProcessatsDir $f.Name
            if (Test-Path -LiteralPath $dest) {
                $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
                $dest = Join-Path $DriveProcessatsDir ("{0}_{1}{2}" -f $f.BaseName, $stamp, $f.Extension)
            }
            Move-Item -LiteralPath $f.FullName -Destination $dest -Force
            Write-Host ("[{0}] OK. Paquet mogut a Processats." -f (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Green
        } catch {
            Write-Host ("[{0}] ERROR amb {1}: {2}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name, $_.Exception.Message) -ForegroundColor Red
            # Movem el paquet problematic a Processats amb sufix .error perque no
            # es reintenti en bucle. El pots revisar a ma.
            try {
                $errDest = Join-Path $DriveProcessatsDir ($f.BaseName + '.error' + $f.Extension)
                if (Test-Path -LiteralPath $errDest) { Remove-Item -LiteralPath $errDest -Force }
                Move-Item -LiteralPath $f.FullName -Destination $errDest -Force
            } catch { }
        }
    }
}

# Mode API (mobil SENSE Drive d'escriptori): recull els paquets directament de
# la carpeta Entrada de Drive, els genera i els mou a Processats.
function Process-PendingApi {
    if (-not $DriveEntradaId -or -not $DriveProcessatsId) {
        Write-Host "Falten \$DriveEntradaId/\$DriveProcessatsId a config.ps1." -ForegroundColor Red
        return
    }
    $pending = Get-DriveChildren $DriveEntradaId '.json'
    foreach ($f in $pending) {
        Write-Host ("[{0}] Paquet detectat (Drive): {1}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name)
        $tmp = Join-Path $env:TEMP ("paquet_" + [guid]::NewGuid().ToString() + ".json")
        try {
            $content = Get-DriveFileText $f.Id
            Set-Content -LiteralPath $tmp -Value $content -Encoding UTF8
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $GenerarPs1 -DesDePaquet $tmp
            if ($LASTEXITCODE -ne 0) { throw "GenerarInforme.ps1 ha retornat codi $LASTEXITCODE" }
            Move-DriveFile $f.Id $DriveProcessatsId $DriveEntradaId
            Write-Host ("[{0}] OK. Paquet mogut a Processats (Drive)." -f (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Green
        } catch {
            Write-Host ("[{0}] ERROR amb {1}: {2}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name, $_.Exception.Message) -ForegroundColor Red
            # El movem a Processats igualment perque no es reintenti en bucle.
            try { Move-DriveFile $f.Id $DriveProcessatsId $DriveEntradaId } catch { }
        } finally {
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        }
    }
}

$UsaApi = Test-DriveApiConfigured

Write-Host "Vigilant d'informes Cornella"
if ($UsaApi) {
    Write-Host "  Mode:       Google Drive per API (sense Drive d'escriptori)"
    Write-Host "  Entrada:    carpeta Drive $DriveEntradaId"
    Write-Host "  Processats: carpeta Drive $DriveProcessatsId"
} else {
    Write-Host "  Mode:       carpeta local de Drive d'escriptori"
    Write-Host "  Entrada:    $DriveEntradaDir"
    Write-Host "  Processats: $DriveProcessatsDir"
}

function Process-All {
    if ($UsaApi) { Process-PendingApi } else { Process-Pending }
}

if ($Once) {
    Process-All
    Write-Host "Fet (mode -Once)."
    return
}
Write-Host ("Vigilant cada {0}s. Tanca la finestra per aturar." -f $IntervalSec)
while ($true) {
    try { Process-All } catch { Write-Host "Error al cicle: $($_.Exception.Message)" -ForegroundColor Red }
    Start-Sleep -Seconds $IntervalSec
}
