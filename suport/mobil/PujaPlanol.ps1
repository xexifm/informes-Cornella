#requires -Version 5.1
<#
  PujaPlanol.ps1 - Puja l'ultim "Planol activitats" a la carpeta PRIVADA Dades
  del Drive (Dades/planol.html), perque el mobil el pugui consultar.

  L'usuari (octubre 2026): "l'eina s'ha de poder usar des del mobil [...]
  Potser s'haura de fer una copia de la base de dades d'activitats al drive".
  Es puja el MATEIX HTML que s'obre al PC (una sola eina, no dues): el mobil el
  llegeix amb el compte de Google de l'usuari (docs/planol.html) i el mostra
  sencer. Porta requeriments i noms d'activitats: va NOMES al Drive privat, mai
  al GitHub public (docs/), com activitats.json.

  El llanca Planol.ps1 en SEGON PLA (Start-ScriptSegonPla) en acabar el planol:
  el planol corre en el proces de rutes/, que no pot carregar Motor.ps1, i aqui
  si (com ExportaDades.ps1), o sigui que la tria "API o carpeta de Drive" es la
  de sempre (Save-ADadesDrive, Activitats.ps1). Deixa el resultat a
  pujada-mobil.log, al costat del planol.
#>
param([Parameter(Mandatory = $true)][string]$Fitxer)

$ErrorActionPreference = 'Stop'
$SuportDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$log = Join-Path (Split-Path -Parent $Fitxer) 'pujada-mobil.log'
$escriu = {
    param($t)
    try { Add-Content -LiteralPath $log -Value ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $t) -Encoding UTF8 } catch { }
}
try {
    $MotorSenseGui = $true
    . (Join-Path $SuportDir 'Motor.ps1')
    $text = [System.IO.File]::ReadAllText($Fitxer, [System.Text.Encoding]::UTF8)
    Save-ADadesDrive 'planol.html' $text 'text/html; charset=UTF-8'
    & $escriu ("OK: $Fitxer -> Dades/planol.html (" + [math]::Round($text.Length / 1KB) + ' KB)')
} catch {
    & $escriu ("ERROR: $($_.Exception.Message)")
    exit 1
}
