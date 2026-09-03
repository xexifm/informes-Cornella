#requires -Version 5.1
<#
.SYNOPSIS
  Envia UNA tanda de recordatoris sense interfície (mode automàtic).

.DESCRIPTION
  El llança la tasca programada del Windows que crea l'eina "Recordatoris"
  (botó "Automàtic..."). Fa NOMÉS una passada i surt: no es queda en segon pla.
  Mateix patró que suport/mobil/Vigilant.ps1.

  Només envia les campanyes que estiguin ACTIVES i en mode 'auto'. Tot el que
  passa queda al registre (%LOCALAPPDATA%\InformesCornella\recordatoris-log.txt).

  PROTECCIÓ CRÍTICA: si la base d'informes és més vella que
  $Script:RecMaxAntiguitatDbDies, NO envia res. Amb una base desfasada
  s'escriuria a titulars que ja han complert, i això no es pot desfer.
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem el motor NOMÉS com a biblioteca: som un procés de consola, no volem
# WinForms ni obrir cap finestra.
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

try {
    _RecLog '--- Execució automàtica ---'

    $db = _RecCarregaDb
    if ($null -eq $db) {
        _RecLog "ATURAT: no hi ha base d'informes (executa 'Actualitzar base')."
        exit 1
    }
    $antig = _RecAntiguitatDb $db (Get-Date)
    if ($antig -lt 0) {
        _RecLog "ATURAT: no es pot saber la data de la base d'informes."
        exit 1
    }
    if ($antig -gt $Script:RecMaxAntiguitatDbDies) {
        _RecLog ("ATURAT: la base d'informes té $antig dies (màxim $($Script:RecMaxAntiguitatDbDies)). " +
                 "Amb una base desfasada s'escriuria a qui ja ha complert.")
        exit 1
    }

    $estat = _RecLlegeix
    $quota = _QuotaLlegeix
    _RecLog "Base de fa $antig dies. Quota: $($quota.enviats)/$($quota.limit)."

    $totalEnviats = 0
    foreach ($camp in @(_RecCampanyes)) {
        $clau = [string]$camp.Clau
        $cfg = $estat.campanyes[$clau]
        if (-not [bool]$cfg['actiu']) { _RecLog "$clau: apagada, no es fa res."; continue }
        if ([string]$cfg['mode'] -ne 'auto') { _RecLog "$clau: en mode manual, no es fa res."; continue }

        $r = _RecDueActivitats $db $camp $cfg $estat.historial[$clau] (Get-Date)
        $toca = @(@($r.Files) | Where-Object { $_.Toca -and -not $_.Excloure })
        _RecLog "$clau: $($toca.Count) activitats toquen avui."
        if ($toca.Count -eq 0) { continue }

        $res = Invoke-RecordatorisTanda $clau $toca $true
        $totalEnviats += [int]$res.Enviats
        _RecLog ("$clau: enviats=$($res.Enviats) fallats=$($res.Fallats) " +
                 "sense_correu=$($res.SenseCorreu)" + $(if ($res.Aturat) { " ATURAT: $($res.Motiu)" } else { '' }))
        # L'historial l'acaba d'escriure la tanda: el rellegim perquè la campanya
        # següent no treballi amb una còpia vella.
        $estat = _RecLlegeix
        if ($res.Aturat) { break }
    }

    $q = _QuotaLlegeix
    _RecLog "Final: $totalEnviats correus enviats. Quota: $($q.enviats)/$($q.limit)."
    exit 0
} catch {
    _RecLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
             $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
