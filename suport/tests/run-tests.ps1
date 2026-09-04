# Proves automatiques de les funcions PURES del motor (Motor.ps1).
#
# NO prova la part de Word/Excel (COM) ni les finestres (WinForms): aixo
# nomes es pot provar a Windows amb Office. Aqui es validen les funcions de
# logica (parseig de camps, normalitzacio, cerca de columnes, dates, claus...).
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File tests/run-tests.ps1
#
# Carrega Motor.ps1 (nomes definicions) en mode "headless" (GENINFORME_TEST=1)
# perque no carregui WinForms. El motor no arrenca res per si sol: qui executa
# el programa es GenerarInforme.ps1, que aqui no toquem.

$ErrorActionPreference = 'Stop'
$env:GENINFORME_TEST = '1'
# A Linux no existeix LOCALAPPDATA; el donem perque el dot-source no falli.
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Motor.ps1'
. $scriptPath   # dot-source: defineix les funcions del motor

. (Join-Path $PSScriptRoot 'TestLib.ps1')   # Assert / AssertEq / Write-TestSummary

# ----------------------------------------------------------------------------
# LES PROVES, PER AREES
# ----------------------------------------------------------------------------
# Aixo era UN fitxer de 5.854 linies amb 128 seccions d'una trentena de moduls:
# mes gros que qualsevol fitxer de produccio. La regla que ens vam escriure -si
# per dir que fa un fitxer necessites la paraula "i", mira-t'ho- hi valia igual.
#
# Van amb DOT-SOURCE i no com a suites a part a posta: aixi comparteixen ambit
# -les variables de muntatge d'un tros les pot fer servir el seguent, com fins
# ara- i el comptador d'asserts es UN, o sigui que partir-ho no en pot perdre
# cap pel cami. L'ORDRE importa: es el mateix que tenien.
# Els trossos fan servir $TestsDir i no $PSScriptRoot: dot-sourcejats des de
# proves\, $PSScriptRoot els apuntaria ALLA i no trobarien ni FormatDoubles.ps1
# ni l'arrel del repositori.
$TestsDir = $PSScriptRoot

foreach ($area in @('01-motor', '02-eines', '03-llicencia', '04-correu', '05-composicio', '06-guards')) {
    . (Join-Path $PSScriptRoot (Join-Path 'proves' ($area + '.ps1')))
}

exit (Write-TestSummary 'RESULTAT')
