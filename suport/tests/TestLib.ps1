#requires -Version 5.1
<#
.SYNOPSIS
  Utillatge compartit de les proves (asserts + resum).

.DESCRIPTION
  Els quatre runners (run-tests, -ruta, -actextr, -precintades) tenien cadascun
  la seva copia de Assert/AssertEq i del bloc de resum. Aqui hi ha una sola
  versio.

  US (des de qualsevol runner, al principi):
      . (Join-Path $PSScriptRoot 'TestLib.ps1')
  i al final:
      exit (Write-TestSummary 'RESULTAT')

  Cal dot-source (no cridar-lo com a script): aixi les funcions i els
  comptadors $script:pass/$script:fail queden a l'abast del runner.
#>

$script:pass = 0
$script:fail = 0

function Assert($cond, $name) {
    if ($cond) { $script:pass++; Write-Host "  OK   $name" -ForegroundColor Green }
    else       { $script:fail++; Write-Host "  FAIL $name" -ForegroundColor Red }
}

function AssertEq($actual, $expected, $name) {
    Assert ([string]$actual -eq [string]$expected) "$name (esperat '$expected', obtingut '$actual')"
}

# Comparacio numerica amb tolerancia (distancies, coordenades...).
function AssertNear($actual, $expected, $tol, $name) {
    Assert ([math]::Abs([double]$actual - [double]$expected) -le $tol) "$name (esperat ~$expected, obtingut $actual)"
}

# Escriu el resum i retorna el codi de sortida (0 = tot OK, 1 = hi ha fallades).
function Write-TestSummary([string]$label = 'RESULTAT') {
    $color = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
    Write-Host "`n========================================"
    Write-Host ("{0}: {1} OK, {2} FAIL" -f $label, $script:pass, $script:fail) -ForegroundColor $color
    Write-Host "========================================"
    if ($script:fail -gt 0) { return 1 } else { return 0 }
}
