#requires -Version 5.1
<#
.SYNOPSIS
  Fa el "Planol activitats" sense interficie (mode automatic setmanal).

.DESCRIPTION
  El llanca el MENU (l'interruptor A/M de sota la rajola "Planol activitats",
  PlanolAutomatic.ps1) quan toca: per defecte cada dilluns a les 13:00 (es
  canvia a Configuracio) i, si l'ultima vegada que tocava no es va fer, en obrir
  el programa.

  Fa el planol amb la MATEIXA funcio que el boto (Invoke-PlanolGenera, en
  silenci: cap finestra ni pregunta), el desa a local\planol-activitats\ i el
  puja al Drive per al mobil (Dades/planol.html, Save-ADadesDrive). Tot el que
  passa va a %LOCALAPPDATA%\InformesCornella\planol-log.txt.

  Planol.ps1 es carrega DINS D'UN BLOC (& { . ... }): defineix $ScriptRoot,
  $RepoRoot... pel seu compte, i aixi no trepitja els del Motor.
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

$ini = @{ auto_el = (Get-Date).ToString('o') }
try {
    $caixa = @{ Res = $null }
    $planolPs1 = Join-Path $ScriptRoot (Join-Path 'rutes' 'Planol.ps1')
    $lliure = Invoke-AmbMutexUnic $Script:PlanolMutexNom {
        $caixa.Res = & {
            $PlanolNomesFuncions = $true
            . $planolPs1
            Invoke-PlanolGenera $true
        }
    }
    [void](_PlanolAutoDesaEstat $ini)
    if (-not $lliure) { _PlanolAutoLog 'Passada automatica: ja se n''esta fent un, no es fa res.'; exit 0 }
    $r = @($caixa.Res)[-1]
    if ($null -eq $r -or -not [bool]$r.Ok) {
        _PlanolAutoLog ('ATURAT: ' + ([string]$r.Error -replace "`r?`n", ' '))
        exit 1
    }
    [void](_PlanolAutoDesaEstat @{ mode = 'auto' })
    _MarcaEinaUsada 'planol'
    try {
        $text = [System.IO.File]::ReadAllText($r.OutPath, [System.Text.Encoding]::UTF8)
        Save-ADadesDrive 'planol.html' $text 'text/html; charset=UTF-8'
        _PlanolAutoLog ("Passada automatica: $($r.OutPath) (i al Drive, Dades/planol.html)")
    } catch {
        _PlanolAutoLog ("Passada automatica: $($r.OutPath). NO s'ha pogut pujar al Drive: " + $_.Exception.Message)
    }
    exit 0
} catch {
    [void](_PlanolAutoDesaEstat $ini)
    _PlanolAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                    $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
