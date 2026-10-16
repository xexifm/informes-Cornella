#requires -Version 5.1
<#
  ExcelErrorsAutomatic.ps1 - El MODE AUTOMATIC de l'"Excel Errors" (l'interruptor
  A/M de sota la rajola).

  L'usuari (octubre 2026): "Vull automatitzar tambe els errors. Consistira en
  enviar un correu amb 3 Excels" (Innocues, Bars i Restaurants, Annexos II i
  III; vegeu ExcelErrorsLlistats.ps1). Per defecte CADA DUES SETMANES, el
  dilluns a l'hora dels automatismes, com el Seguiment (es canvia a
  Configuracio -> Automatismes); si l'ultima vegada que tocava no es va fer, es
  fa en obrir el programa.

  LA PASSADA (EN SEGON PLA, ExcelErrorsAuto.ps1) fa el mateix que el boto
  "Enviar per correu" de la finestra: llegeix l'Excel, en treu els errors dels
  tipus MARCATS a la finestra, sense les activitats tretes dels llistats, fa
  els tres Excels i els envia a les adreces fixes de l'eina. Mai la CCO "a
  triar". Sense destinataris l'interruptor no s'engega.

  L'estat (excel-errors-auto.json) i l'enviament son a ExcelErrorsLlistats.ps1.
  NOMES DEFINEIX FUNCIONS (i s'apunta als registres).
#>

$Script:ExMutexNom = 'Global\InformesCornella.ExcelErrors'

function _ExAutoLog([string]$msg) { Write-AutoLog 'excel-errors-log.txt' $msg }

# LA PASSADA AUTOMATICA. Apunta 'auto_el' SEMPRE (si no, el menu la tornaria a
# llancar cada minut). Torna @{ Ok; Error }.
function Invoke-ExcelErrorsAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $caixa = @{ Gen = $null; Env = $null }
    $lliure = Invoke-AmbMutexUnic $Script:ExMutexNom {
        $caixa.Gen = Invoke-ExcelErrorsLlistats
        if ([bool]$caixa.Gen.Ok) { $caixa.Env = Send-ExcelErrorsCorreu $caixa.Gen $true }
    }
    [void](_ExAutoDesaEstat $ini)
    if (-not $lliure) { _ExAutoLog 'Passada automatica: ja se n''esta fent una, no es fa res.'; return @{ Ok = $false; Error = 'ocupat' } }
    if (-not [bool]$caixa.Gen.Ok) {
        _ExAutoLog ('ATURAT: ' + ([string]$caixa.Gen.Error -replace "`r?`n", ' '))
        return @{ Ok = $false; Error = [string]$caixa.Gen.Error }
    }
    $resum = (@($caixa.Gen.Resum.Keys) | ForEach-Object { [string]$_ + ': ' + [string]$caixa.Gen.Resum[$_] }) -join '; '
    $sense = @($caixa.Gen.SenseLlistat | ForEach-Object { [string]$_.Id })
    $txt = $resum + $(if ($sense.Count -gt 0) { '. Amb errors i sense llistat (no s''envien): ' + ($sense -join ', ') } else { '' })
    _ExAutoLog ($txt + ' ' + [char]0x00B7 + ' ' + [string]$caixa.Env.Text)
    return @{ Ok = [bool]$caixa.Env.Ok; Error = $(if ($caixa.Env.Ok) { '' } else { [string]$caixa.Env.Text }) }
}

function Start-ExcelErrorsAuto { return (Start-ProcesAutoUnic 'excelerrors' 'ExcelErrorsAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-ExcelErrorsAutoSiToca {
    $est = _ExAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (Test-ProgramacioToca 'excelerrors' (Get-Date) $est['auto_el'])) { return $false }
    return (Start-ExcelErrorsAuto)
}

Register-ProgramacioAuto 'excelerrors' 'Excel Errors' 'quinzena' 1
$Script:ModesAuto['excelerrors'] = @{
    Titol     = 'Excel Errors'
    Actiu     = { $e = _ExAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _ExAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _ExAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-ExcelErrorsAutoSiToca }
    # Sense destinataris faria els Excels i no els enviaria a ningu.
    Requisit  = {
        if ([string](Get-CorreuEina 'excelerrors').Fixes) { return '' }
        return ("Per enviar els errors sols cal dir a qui van.`n`nVes a Configuraci" + [char]0x00F3 + " -> Configuraci" + [char]0x00F3 + " correu electr" + [char]0x00F2 + "nic -> Excel Errors i posa-hi les adreces (Adreces fixes).")
    }
    TipA      = { Get-AutoTipText "es fan els tres Excels d'errors i s'envien per correu" 'excelerrors' }
    TipM      = ("Mode MANUAL: nom" + [char]0x00E9 + "s es fa quan cliques la rajola. Clica per posar-ho en autom" + [char]0x00E0 + "tic.")
}
