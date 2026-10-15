#requires -Version 5.1
<#
.SYNOPSIS
  Fa els "Controls periodics" sense interficie (mode automatic): prepara a
  Esborranys de l'Outlook els correus dels controls periodics vençuts.

.DESCRIPTION
  El llanca el MENU (l'interruptor A/M de sota la rajola "Controls periodics",
  ControlsPeriodicsAutomatic.ps1) quan toca: per defecte cada dilluns a les
  13:00 (es canvia a Configuracio) i, si l'ultima vegada que tocava no es va
  fer, en obrir el programa. MAI envia res: nomes esborranys.
  Tot el que passa va a %LOCALAPPDATA%\InformesCornella\controls-periodics-log.txt.
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

try {
    $r = Invoke-ControlsPeriodicsAuto
    if ([bool]$r.Ok) { _MarcaEinaUsada 'controlsperiodics'; exit 0 }
    exit 1
} catch {
    try { [void](_CpAutoDesaEstat @{ auto_el = (Get-Date).ToString('o') }) } catch { }
    _CpAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
