#requires -Version 5.1
<#
.SYNOPSIS
  Revisa UNA vegada si han arribat informes del mobil (via Drive) i els genera.

.DESCRIPTION
  Mira la carpeta d'entrada ($DriveEntradaDir, sincronitzada amb el Google Drive
  d'escriptori) o la carpeta de Drive per API. Per cada paquet *.json pendent
  (preparat al mobil), crida GenerarInforme.ps1 -DesDePaquet per generar el .docx
  complet i mou el paquet a Processats. Fa NOMES una passada i surt: no es queda
  vigilant en segon pla ni fa polling. El programa el llanca (una vegada) des del
  boto "Revisar entrades del mobil" del menu.

.PARAMETER ResultFile
  Ruta (opcional) on escriure el resum en JSON ({ ok, err, pending }) perque el
  programa que l'ha llancat el pugui mostrar.
#>
param(
    [string]$ResultFile
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
# Aquest script viu a suport/mobil/. El motor compartit (GenerarInforme.ps1)
# viu a suport/, que es el directori pare.
$SuportDir  = Split-Path -Parent $ScriptRoot

# Carreguem GenerarInforme.ps1 en headless NOMES per obtenir les rutes de Drive
# ($DriveEntradaDir, $DriveProcessatsDir...). La generacio real es fa cridant el
# .ps1 com a proces independent (amb -DesDePaquet), per aïllar cada informe.
$env:GENINFORME_TEST = '1'
. (Join-Path $SuportDir 'GenerarInforme.ps1')
Remove-Item Env:\GENINFORME_TEST -ErrorAction SilentlyContinue

$GenerarPs1 = Join-Path $SuportDir 'GenerarInforme.ps1'

# Comptadors del resum d'aquesta passada.
$Script:GeneratsOk  = 0
$Script:GeneratsErr = 0
$Script:Pendents    = 0

function _EnsureDir($d) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

function Process-Pending {
    _EnsureDir $DriveEntradaDir
    _EnsureDir $DriveProcessatsDir
    $pending = @(Get-ChildItem -LiteralPath $DriveEntradaDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime)
    $Script:Pendents += $pending.Count
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
            $Script:GeneratsOk++
            Write-Host ("[{0}] OK. Paquet mogut a Processats." -f (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Green
        } catch {
            $Script:GeneratsErr++
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
    $pending = @(Get-DriveChildren $DriveEntradaId '.json')
    $Script:Pendents += $pending.Count
    foreach ($f in $pending) {
        Write-Host ("[{0}] Paquet detectat (Drive): {1}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name)
        $tmp = Join-Path $env:TEMP ("paquet_" + [guid]::NewGuid().ToString() + ".json")
        try {
            Write-Host "   1/3 Baixant el paquet de Drive..."
            $content = Get-DriveFileText $f.Id
            Set-Content -LiteralPath $tmp -Value $content -Encoding UTF8

            Write-Host "   2/3 Generant l'informe..."
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $GenerarPs1 -DesDePaquet $tmp
            if ($LASTEXITCODE -ne 0) { throw "GenerarInforme.ps1 ha retornat codi $LASTEXITCODE (mira la finestra del generador)" }

            Write-Host "   3/3 Movent el paquet a Processats..."
            Move-DriveFile $f.Id $DriveProcessatsId $DriveEntradaId
            $Script:GeneratsOk++
            Write-Host ("[{0}] OK. Informe generat i paquet mogut a Processats." -f (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Green
        } catch {
            $Script:GeneratsErr++
            Write-Host ("[{0}] ERROR amb {1}: {2}" -f (Get-Date -Format 'HH:mm:ss'), $f.Name, $_.Exception.Message) -ForegroundColor Red
            # El movem a Processats igualment perque no es reintenti en bucle.
            try { Move-DriveFile $f.Id $DriveProcessatsId $DriveEntradaId } catch { }
        } finally {
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        }
    }
}

# Comprovacio inicial: si hi ha carpetes de Drive configurades (mode API) pero
# aquest PC encara NO esta autoritzat, llancem l'autorització abans de comencar.
# Aixi, en un PC nou, el Vigilant et demana el Secret un sol cop i continua.
if ($DriveEntradaId -and -not (Test-DriveApiConfigured)) {
    Write-Host ""
    Write-Host "Aquest PC encara NO esta autoritzat a Google Drive." -ForegroundColor Yellow
    Write-Host "Iniciem l'autoritzacio (et demanara el Secret un sol cop i s'obrira el navegador)..." -ForegroundColor Yellow
    Write-Host ""
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptRoot 'Authorize-Drive.ps1')
    Write-Host ""
    if (Test-DriveApiConfigured) {
        Write-Host "Autoritzacio correcta. Continuem amb la revisio." -ForegroundColor Green
    } else {
        Write-Host "No s'ha completat l'autoritzacio. Es treballara en mode carpeta local." -ForegroundColor Yellow
    }
    Write-Host ""
}

$UsaApi = Test-DriveApiConfigured

Write-Host "Revisar entrades del mobil (una passada)"
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

# UNA sola passada i sortim (no hi ha bucle ni polling).
try { Process-All } catch { Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red; $Script:GeneratsErr++ }

Write-Host ""
Write-Host ("Fet. Pendents: {0}  -  Generats: {1}  -  Errors: {2}" -f $Script:Pendents, $Script:GeneratsOk, $Script:GeneratsErr) -ForegroundColor Cyan

# Resum en JSON per al programa que ens ha llancat (si ens ha passat -ResultFile).
if (-not [string]::IsNullOrWhiteSpace($ResultFile)) {
    try {
        ([pscustomobject]@{ ok = $Script:GeneratsOk; err = $Script:GeneratsErr; pending = $Script:Pendents } |
            ConvertTo-Json -Compress) | Set-Content -LiteralPath $ResultFile -Encoding UTF8
    } catch { }
}
