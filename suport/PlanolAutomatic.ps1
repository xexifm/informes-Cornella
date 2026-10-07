#requires -Version 5.1
<#
  PlanolAutomatic.ps1 - El MODE AUTOMATIC del "Planol activitats" (l'interruptor
  A/M de sota la rajola).

  L'usuari (octubre 2026): "aplica l'automatisme de Copiar informes i
  Actualitzar base a Planol activitats, pero un cop a la setmana en comptes d'un
  cop al dia". Per defecte, cada DILLUNS a l'hora dels modes automatics (es
  canvia a Configuracio, com tots). Mateixa regla: si l'ultima vegada que tocava
  no es va fer, es fa en obrir el programa.

  La passada es fa EN SEGON PLA (PlanolAuto.ps1): llegeix els Excel, demana al
  Cadastre el que falti (memoria cau), desa el planol a local\planol-activitats\
  i el puja al Drive per al mobil (Dades/planol.html). No obre res.

  L'ESTAT, a local\planol-activitats\planol-auto.json:
    auto     l'interruptor
    auto_el  l'ultima PASSADA automatica (encara que no hagi pogut fer res)
    mode     'auto' | 'manual': qui va fer l'ultim planol

  NOMES DEFINEIX FUNCIONS (i s'apunta al registre de ModeAutomatic.ps1).
#>

$Script:PlanolEstatPlantilla = [ordered]@{ auto = $false; auto_el = ''; mode = '' }
$Script:PlanolMutexNom = 'Global\InformesCornella.PlanolActivitats'

function _PlanolAutoStatePath {
    try { return [string](Join-Path (Get-LocalSubdir $RepoRoot 'Planol') 'planol-auto.json') } catch { return '' }
}

function _PlanolAutoEstat { return (Read-EstatAuto (_PlanolAutoStatePath) $Script:PlanolEstatPlantilla) }

function _PlanolAutoDesaEstat($canvis) { return (Save-EstatAuto (_PlanolAutoStatePath) $canvis $Script:PlanolEstatPlantilla) }

function _PlanolAutoLog([string]$msg) { Write-AutoLog 'planol-log.txt' $msg }

function _PlanolAutoToca([datetime]$ara, $ultimAuto) {
    return (Test-ProgramacioToca 'planol' $ara $ultimAuto)
}

function Start-PlanolAuto { return (Start-ProcesAutoUnic 'planol' 'PlanolAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-PlanolAutoSiToca {
    $est = _PlanolAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (_PlanolAutoToca (Get-Date) $est['auto_el'])) { return $false }
    return (Start-PlanolAuto)
}

Register-ProgramacioAuto 'planol' ('Pl' + [char]0x00E0 + 'nol activitats') 'setmana' 1
$Script:ModesAuto['planol'] = @{
    Titol     = ('Pl' + [char]0x00E0 + 'nol activitats')
    Actiu     = { $e = _PlanolAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _PlanolAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _PlanolAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-PlanolAutoSiToca }
    # Sense l'Excel d'activitats no hi ha planol.
    Requisit  = {
        if (-not [string]::IsNullOrWhiteSpace($ActivitatsDir) -or -not [string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
        return ("Per fer el pl" + [char]0x00E0 + "nol sol cal dir on " + [char]0x00E9 + "s l'Excel d'activitats.`n`nVes a Configuraci" + [char]0x00F3 + " i indica 'Carpeta de l'Excel d'activitats'.")
    }
    TipA      = { Get-AutoTipText ('el pl' + [char]0x00E0 + 'nol es fa sol (i es puja al m' + [char]0x00F2 + 'bil)') 'planol' }
    TipM      = ("Mode MANUAL: nom" + [char]0x00E9 + "s es fa quan cliques la rajola. Clica per posar-ho en autom" + [char]0x00E0 + "tic.")
}
