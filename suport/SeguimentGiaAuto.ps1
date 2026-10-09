#requires -Version 5.1
<#
.SYNOPSIS
  Fa el "Seguiment" sense interficie (mode automatic, cada dues setmanes) i
  l'envia per correu.

.DESCRIPTION
  El llanca el MENU (l'interruptor A/M de sota la rajola "Seguiment",
  SeguimentGiaAutomatic.ps1) quan toca: per defecte dilluns alternats a les
  13:00 (es canvia a Configuracio) i, si l'ultima vegada que tocava no es va
  fer, en obrir el programa. Fa el PDF dels llistats (tots menys la fulla
  "Estes") i l'envia als destinataris de Configuracio -> Correus de cada eina.
  Tot el que passa va a %LOCALAPPDATA%\InformesCornella\seguiment-log.txt.
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

try {
    $r = Invoke-SeguimentAuto
    if ([bool]$r.Ok) { _MarcaEinaUsada 'seguimentgia'; exit 0 }
    exit 1
} catch {
    try { [void](_SgAutoDesaEstat @{ auto_el = (Get-Date).ToString('o') }) } catch { }
    _SgAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
