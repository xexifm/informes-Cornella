#requires -Version 5.1
<#
.SYNOPSIS
  Fa els tres Excels d'errors del GIA sense interficie (mode automatic) i els
  envia per correu.

.DESCRIPTION
  El llanca el MENU (l'interruptor A/M de sota la rajola "Excel Errors",
  ExcelErrorsAutomatic.ps1) quan toca: per defecte cada dues setmanes, el
  dilluns a les 13:00 (es canvia a Configuracio) i, si l'ultima vegada que
  tocava no es va fer, en obrir el programa.
  Tot el que passa va a %LOCALAPPDATA%\InformesCornella\excel-errors-log.txt.
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

try {
    $r = Invoke-ExcelErrorsAuto
    if ([bool]$r.Ok) { _MarcaEinaUsada 'excelerrors'; exit 0 }
    exit 1
} catch {
    try { [void](_ExAutoDesaEstat @{ auto_el = (Get-Date).ToString('o') }) } catch { }
    _ExAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
