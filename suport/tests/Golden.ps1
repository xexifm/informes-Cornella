#requires -Version 5.1
<#
.SYNOPSIS
  Fitxers d'or: la seqüencia EXACTA de crides de format d'un informe.

.DESCRIPTION
  Un fitxer d'or es la llista sencera de crides Format-* que fa una familia
  d'informes, una per linia, tal com les enregistra FormatDoubles.ps1.

  PER QUE. El motor de composicio s'ha d'anar refent, i aquest projecte te
  l'historial de defectes que NO fallen sino que EMPITJOREN EN SILENCI: un
  espai que desapareix, un sub-punt que passa de 12 a 6 pt, un enllac que
  canvia de lloc. Cap prova puntual no els veu tots; una comparacio linia a
  linia de tot el document, si.

  US:
      Assert-Golden 'llicencia-requeriment' $global:emitCalls

  Per REFER-LOS despres d'un canvi VOLGUT:
      GENINFORME_GOLDEN=1 pwsh -File suport/tests/run-tests-golden.ps1
  ...i despres MIRA'T EL `git diff`: es tota la gracia. Si el diff no es
  exactament el que esperaves, el canvi no era el que et pensaves.
#>

$Script:GoldenDir = Join-Path $PSScriptRoot 'dades'

function _GoldenPath([string]$nom) { return (Join-Path $Script:GoldenDir ("emit-$nom.txt")) }

function _GoldenEsRefer { return (-not [string]::IsNullOrWhiteSpace($env:GENINFORME_GOLDEN)) }

# Compara la seqüencia amb el fitxer d'or. Si no hi es (o s'esta refent), l'escriu.
function Assert-Golden([string]$nom, $crides) {
    $linies = @(@($crides) | ForEach-Object { [string]$_ })
    $path = _GoldenPath $nom

    # PARENTESIS OBLIGATORIS a la crida: sense ells, PowerShell llegeix
    # "_GoldenEsRefer -or -not (...)" com una CRIDA A L'ORDRE amb '-or' i '-not'
    # com a ARGUMENTS, i la condicio sempre es certa. (Mateixa familia que la
    # trampa del '+' solt que hi ha documentada a CLAUDE.md.)
    if ((_GoldenEsRefer) -or (-not (Test-Path -LiteralPath $path))) {
        if (-not (Test-Path -LiteralPath $Script:GoldenDir)) {
            New-Item -ItemType Directory -Path $Script:GoldenDir -Force | Out-Null
        }
        # Sense BOM i amb salt de linia \n: aixi el git diff es net a totes dues
        # plataformes i el fitxer no balla segons qui l'escriu.
        $txt = ($linies -join "`n")
        if ($linies.Count -gt 0) { $txt += "`n" }
        [System.IO.File]::WriteAllText($path, $txt, (New-Object System.Text.UTF8Encoding($false)))
        Assert $true ("or [$nom]: escrit (" + $linies.Count + ' linies)')
        return
    }

    $esperat = @()
    $raw = [System.IO.File]::ReadAllText($path)
    if (-not [string]::IsNullOrEmpty($raw)) {
        $esperat = @(($raw -replace "`r`n", "`n").TrimEnd("`n") -split "`n")
        if ($esperat.Count -eq 1 -and [string]::IsNullOrEmpty($esperat[0])) { $esperat = @() }
    }

    # LA PRIMERA DIFERENCIA es el que interessa: la resta acostuma a ser el
    # mateix desplacat una posicio, i abocar-ho tot amaga la linia que importa.
    $n = [Math]::Max($esperat.Count, $linies.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $e = if ($i -lt $esperat.Count) { $esperat[$i] } else { '<res>' }
        $a = if ($i -lt $linies.Count)  { $linies[$i] }  else { '<res>' }
        if ($e -ne $a) {
            Assert $false ("or [$nom]: la linia " + ($i + 1) + " ha canviat`n         esperat: " + $e + "`n         obtingut: " + $a)
            return
        }
    }
    AssertEq $linies.Count $esperat.Count "or [$nom]: la seqüencia sencera es igual ($($linies.Count) linies)"
}
