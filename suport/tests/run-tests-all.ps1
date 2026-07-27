#requires -Version 5.1
<#
.SYNOPSIS
  Executa TOTES les suites de proves i retorna un sol codi de sortida.

.DESCRIPTION
  Cada suite corre en un PROCES A PART: aixi els dobles de prova, les
  variables d'entorn (GENINFORME_TEST, RUTA_TEST...) i les funcions
  redefinides d'una suite no poden contaminar la seguent.

  Codi de sortida 0 nomes si TOTES les suites passen; 1 si en falla alguna
  (aixi es pot encadenar: run-tests-all.ps1 && git commit ...).

.PARAMETER Suite
  Nom(s) de suite a executar (sense el prefix 'run-tests-'). Per defecte,
  totes. Exemple: -Suite ruta,actextr
#>
param(
    [string[]]$Suite
)

$ErrorActionPreference = 'Stop'

$totes = [ordered]@{
    'motor'       = 'run-tests.ps1'
    'ruta'        = 'run-tests-ruta.ps1'
    'actextr'     = 'run-tests-actextr.ps1'
    'precintades' = 'run-tests-precintades.ps1'
}

$aExecutar = if ($Suite) { $Suite } else { @($totes.Keys) }

# Mateix host de PowerShell que ens executa (pwsh o powershell 5.1), perque
# les proves corrin exactament on l'usuari les ha llancat.
$hostExe = (Get-Process -Id $PID).Path
if ([string]::IsNullOrWhiteSpace($hostExe)) { $hostExe = 'powershell.exe' }

$resultats = @()
$fallades  = 0

foreach ($nom in $aExecutar) {
    if (-not $totes.Contains($nom)) {
        Write-Host "Suite desconeguda: '$nom' (n'hi ha: $($totes.Keys -join ', '))" -ForegroundColor Red
        $fallades++
        continue
    }
    $fitxer = Join-Path $PSScriptRoot $totes[$nom]
    Write-Host ""
    Write-Host "########## $nom ($($totes[$nom])) ##########" -ForegroundColor Cyan

    $sortida = & $hostExe -NoProfile -ExecutionPolicy Bypass -File $fitxer 2>&1
    $codi = $LASTEXITCODE
    $sortida | ForEach-Object { Write-Host $_ }

    # Recomptes de la linia de resum, per al total final.
    $ok = 0; $ko = 0
    $linia = @($sortida | Where-Object { "$_" -match '^RESULTAT.*: (\d+) OK, (\d+) FAIL' }) | Select-Object -Last 1
    if ($linia -and ("$linia" -match ': (\d+) OK, (\d+) FAIL')) { $ok = [int]$Matches[1]; $ko = [int]$Matches[2] }

    $resultats += [pscustomobject]@{ Suite = $nom; OK = $ok; FAIL = $ko; Codi = $codi }
    if ($codi -ne 0) { $fallades++ }
}

Write-Host ""
Write-Host "========================================"
Write-Host "RESUM DE TOTES LES SUITES" -ForegroundColor Cyan
Write-Host "========================================"
foreach ($r in $resultats) {
    $estat = if ($r.Codi -eq 0) { 'OK   ' } else { 'FALLA' }
    $color = if ($r.Codi -eq 0) { 'Green' } else { 'Red' }
    Write-Host ("  {0}  {1,-12} {2,4} proves, {3} fallades" -f $estat, $r.Suite, $r.OK, $r.FAIL) -ForegroundColor $color
}
$totalOk = ($resultats | Measure-Object -Property OK -Sum).Sum
$totalKo = ($resultats | Measure-Object -Property FAIL -Sum).Sum
Write-Host "----------------------------------------"
Write-Host ("  TOTAL: {0} OK, {1} FAIL" -f $totalOk, $totalKo) -ForegroundColor $(if ($fallades -eq 0) { 'Green' } else { 'Red' })
Write-Host "========================================"

if ($fallades -gt 0) { exit 1 } else { exit 0 }
