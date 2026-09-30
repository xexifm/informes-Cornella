#requires -Version 5.1
<#
.SYNOPSIS
  Comprova els enllacos (URLs) dels catalegs i avisa dels que estan CAIGUTS.

.DESCRIPTION
  Llegeix els catalegs .json d'ESTRUCTURALS, n'extreu tots els enllacos i fa una
  peticio a cadascun per saber si responen. Marca en VERD els que van be i en
  VERMELL els que estan CAIGUTS (codi 4xx/5xx o sense resposta), i al final en
  fa un resum.

  Llegeix el JSON, que es la FONT DE VERITAT dels catalegs. Abans obria el .docx
  com un ZIP i en treia els hipervincles i el text visible; els .docx ja no son
  catalegs (son vistes generades), i al JSON els enllacos ja venen marcats amb
  "url": true, o sigui que no cal endevinar res.

.PARAMETER Cataleg
  Nom o ruta del cataleg a comprovar. Sense valor, es comproven TOTS els
  catalegs d'ESTRUCTURALS (els .json que no comencen per "0 ").

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1 REQ1
#>

param([string]$Cataleg)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot   = Split-Path -Parent $ScriptRoot
$EstructDir = Join-Path $RepoRoot 'ESTRUCTURALS'

# Decideix quins catalegs comprovar.
$targets = @()
if ([string]::IsNullOrWhiteSpace($Cataleg)) {
    $targets = @(Get-ChildItem -LiteralPath $EstructDir -Filter '*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '0 *' -and -not $_.Name.StartsWith('~$') } |
        Sort-Object Name | ForEach-Object { $_.FullName })
} elseif (Test-Path -LiteralPath $Cataleg) {
    $targets = @((Resolve-Path -LiteralPath $Cataleg).Path)
} else {
    # Nom curt ('REQ1'): el busquem a ESTRUCTURALS.
    $n = if ([System.IO.Path]::GetExtension($Cataleg) -ieq '.json') { $Cataleg } else { "$Cataleg.json" }
    $targets = @(Join-Path $EstructDir $n)
}
if ($targets.Count -eq 0) { Write-Host "No s'ha trobat cap cataleg a comprovar." -ForegroundColor Red; exit 1 }

# El recorregut i la comprovacio son a Enllacos.ps1 (la mateixa que fa servir
# l'eina "Revisar requeriments"). Abans eren aqui i cridaven Read-JsonFile sense
# carregar Json.ps1: l'script petava a cada cataleg.
. (Join-Path $ScriptRoot 'Json.ps1')
. (Join-Path $ScriptRoot 'Enllacos.ps1')

$totalCaiguts = @()

foreach ($t in $targets) {
    if (-not (Test-Path -LiteralPath $t)) { Write-Host "No trobo: $t" -ForegroundColor Red; continue }
    $urls = Get-CatalegUrls $t
    Write-Host ("`n===== {0}  ({1} enllacos) =====" -f (Split-Path -Leaf $t), $urls.Count) -ForegroundColor Cyan
    foreach ($u in $urls) {
        $prova = Test-EnllacViu $u
        $code = $prova.Codi; $ok = [bool]$prova.Ok
        if ($ok) {
            Write-Host ("  OK     [{0}] {1}" -f $code, $u) -ForegroundColor Green
        } else {
            Write-Host ("  CAIGUT [{0}] {1}" -f $code, $u) -ForegroundColor Red
            $totalCaiguts += [pscustomobject]@{ Fitxer=(Split-Path -Leaf $t); Codi=$code; Url=$u }
        }
    }
}

Write-Host ("`n========================================") -ForegroundColor Cyan
if ($totalCaiguts.Count -eq 0) {
    Write-Host "Tots els enllacos responen correctament." -ForegroundColor Green
} else {
    Write-Host ("ENLLACOS CAIGUTS: {0}" -f $totalCaiguts.Count) -ForegroundColor Red
    $totalCaiguts | ForEach-Object { Write-Host ("  [{0}] {1}  ({2})" -f $_.Codi, $_.Url, $_.Fitxer) -ForegroundColor Yellow }
}
Write-Host ("========================================")
