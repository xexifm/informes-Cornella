#requires -Version 5.1
<#
  ControlsPeriodicsAutomatic.ps1 - El MODE AUTOMATIC dels "Controls periodics"
  (l'interruptor A/M de sota la rajola).

  L'usuari (octubre 2026): "aplica l'automatitzacio de les eines a l'eina
  Controls periodics". La feina que l'usuari feia a ma era: obrir la llista,
  marcar les activitats que ja han passat el "Proper CP previst" i preparar-ne
  els correus. Aixo es el que fa l'automatic, per defecte CADA DILLUNS a l'hora
  dels automatismes (es canvia a Configuracio -> Automatismes, com tots; si
  l'ultima vegada que tocava no es va fer, es fa en obrir el programa).

  QUE FA LA PASSADA (EN SEGON PLA, ControlsPeriodicsAuto.ps1):
    - llegeix l'Excel d'activitats (la mateixa lectura que la finestra,
      _ReadControlsPeriodics);
    - tria les activitats amb el control VENÇUT ("Proper CP previst" anterior a
      avui: el text del correu diu "havia de passar ... i no ens consta") i que
      encara no s'han avisat d'aquella data (l'historial, ControlsCpEmail.ps1);
    - en prepara l'esborrany a l'Outlook, com el boto (_CpPreparaEsborranys).
  NOMES ESBORRANYS, MAI ENVIA: es la regla de l'eina (hi ha guard), i aqui
  encara mes, que no hi ha ningu davant. El menu avisa que n'hi ha de pendents
  (Add-CorreuEsborranysPendents), com amb els recordatoris.

  COM A MOLT $Script:CpAutoMaxPerTanda PER PASSADA, i primer els que han vençut
  FA MENYS: la primera vegada hi pot haver centenars de controls vençuts de fa
  anys, i el que val la pena que surti sol es el que acaba de vencer. La resta
  van sortint les setmanes seguents (o es fan a ma des de la finestra).
  La CCO "a triar" no hi va mai (no hi ha ningu per triar-la): nomes la de
  "sempre".

  L'ESTAT, a local\base-dades-activitats\controls-periodics-auto.json:
    auto       l'interruptor
    auto_el    l'ultima PASSADA automatica (encara que no hagi pogut fer res)
    mode       'auto' | 'manual': qui va preparar els ultims esborranys
    fet_el     quan

  NOMES DEFINEIX FUNCIONS (i s'apunta als registres). Es carrega DESPRES de
  ControlsCpEmail.ps1.
#>

$Script:CpEstatPlantilla = [ordered]@{ auto = $false; auto_el = ''; mode = ''; fet_el = '' }
$Script:CpMutexNom = 'Global\InformesCornella.ControlsPeriodics'
$Script:CpAutoMaxPerTanda = 20

function _CpAutoStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'controls-periodics-auto.json')
}
function _CpAutoEstat { return (Read-EstatAuto (_CpAutoStatePath) $Script:CpEstatPlantilla) }
function _CpAutoDesaEstat($canvis) { return (Save-EstatAuto (_CpAutoStatePath) $canvis $Script:CpEstatPlantilla) }
function _CpAutoLog([string]$msg) { Write-AutoLog 'controls-periodics-log.txt' $msg }

# QUINES activitats toquen: control vençut (KProperCP abans d'avui), encara no
# avisat d'aquella data, les que han vençut FA MENYS primer, com a molt $max.
# Torna @{ Triades; Vençuts (tots els vençuts); Pendents (els que no hi caben) }.
# PURA.
function _CpAutoTriades($rows, $avisats, [datetime]$avui, [int]$max) {
    $venc = New-Object System.Collections.ArrayList
    $nous = New-Object System.Collections.ArrayList
    foreach ($r in @($rows)) {
        if ($null -eq $r -or $null -eq $r.KProperCP) { continue }
        if (([datetime]$r.KProperCP).Date -ge $avui.Date) { continue }
        [void]$venc.Add($r)
        if (-not (_CpJaAvisat $r $avisats)) { [void]$nous.Add($r) }
    }
    $ord = @($nous | Sort-Object -Property @{ Expression = { [datetime]$_.KProperCP }; Descending = $true }, @{ Expression = { _GiaNumeric $_.Id } })
    $tria = @($ord | Select-Object -First ([Math]::Max(0, $max)))
    return @{ Triades = $tria; Vencuts = $venc.Count; Pendents = ($ord.Count - $tria.Count) }
}

# LA PASSADA AUTOMATICA (la crida ControlsPeriodicsAuto.ps1). Cap finestra: tot
# al registre. Apunta 'auto_el' SEMPRE (si no, el menu la tornaria a llancar
# cada minut). Torna @{ Ok; Error; Text }.
function Invoke-ControlsPeriodicsAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $caixa = @{ R = $null }
    $lliure = Invoke-AmbMutexUnic $Script:CpMutexNom {
        $llista = _ReadControlsPeriodics
        if (-not [bool]$llista.Ok) { $caixa.R = @{ Ok = $false; Error = [string]$llista.Error }; return }
        $t = _CpAutoTriades @($llista.Rows) (Read-CpAvisats) (Get-Date) $Script:CpAutoMaxPerTanda
        if (@($t.Triades).Count -eq 0) {
            $caixa.R = @{ Ok = $true; Error = ''; Text = ("Cap control peri" + [char]0x00F2 + "dic ven" + [char]0x00E7 + "ut per avisar (" + [string]$t.Vencuts + " de ven" + [char]0x00E7 + "uts, tots avisats).") }
            return
        }
        $cfgE = Get-CorreuEina 'controls'
        # NOMES la CCO de "sempre": la "a triar" demana algu que la tri.
        $bcc = if ($null -ne $cfgE.Cco) { [string]$cfgE.Cco } else { '' }
        $p = _CpPreparaEsborranys @($t.Triades) $cfgE $bcc 'auto'
        if ($p.Fets.Count -gt 0) {
            Add-CorreuEsborranysPendents ('Controls peri' + [char]0x00F2 + 'dics') @($p.Fets | ForEach-Object { "GIA $($_.Id) - $($_.RaoSocial) (Proper CP previst $($_.ProperCP))" })
            [void](_CpAutoDesaEstat @{ mode = 'auto'; fet_el = (Get-Date).ToString('o') })
        }
        $txt = "Esborranys: " + $p.Fets.Count + " de " + @($t.Triades).Count
        if ($t.Pendents -gt 0) { $txt += " (" + $t.Pendents + " per a les properes setmanes)" }
        if ($p.SenseCorreu.Count -gt 0) { $txt += ". Sense correu: " + ($p.SenseCorreu -join '; ') }
        if ($p.AvisTecnic.Count -gt 0) { $txt += ". Correu del tecnic segons els documents: " + ($p.AvisTecnic -join '; ') }
        if ($p.Errors.Count -gt 0) { $txt += ". Errors: " + ($p.Errors -join '; ') }
        $caixa.R = @{ Ok = ($p.Errors.Count -eq 0); Error = ($p.Errors -join '; '); Text = $txt }
    }
    [void](_CpAutoDesaEstat $ini)
    if (-not $lliure) { _CpAutoLog 'Passada automatica: ja se n''esta fent una, no es fa res.'; return @{ Ok = $false; Error = 'ocupat' } }
    $r = $caixa.R
    if (-not [bool]$r.Ok -and -not [string]$r.Text) {
        _CpAutoLog ('ATURAT: ' + ([string]$r.Error -replace "`r?`n", ' '))
        return $r
    }
    _CpAutoLog ([string]$r.Text -replace "`r?`n", ' ')
    return $r
}

function Start-ControlsPeriodicsAuto { return (Start-ProcesAutoUnic 'controlsperiodics' 'ControlsPeriodicsAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-ControlsPeriodicsAutoSiToca {
    $est = _CpAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (Test-ProgramacioToca 'controlsperiodics' (Get-Date) $est['auto_el'])) { return $false }
    return (Start-ControlsPeriodicsAuto)
}

# ----------------------------------------------------------------------------
# ELS REGISTRES
# ----------------------------------------------------------------------------
Register-ProgramacioAuto 'controlsperiodics' ('Controls peri' + [char]0x00F2 + 'dics') 'setmana' 1
$Script:ModesAuto['controlsperiodics'] = @{
    Titol     = ('Controls peri' + [char]0x00F2 + 'dics')
    Actiu     = { $e = _CpAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _CpAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _CpAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-ControlsPeriodicsAutoSiToca }
    # Sense l'Excel d'activitats no hi ha llista (i l'estat va al seu costat).
    Requisit  = {
        if (-not [string]::IsNullOrWhiteSpace($ActivitatsDir) -or -not [string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
        return ("Per preparar sols els correus dels controls peri" + [char]0x00F2 + "dics cal dir on " + [char]0x00E9 + "s l'Excel d'activitats.`n`nVes a Configuraci" + [char]0x00F3 + " i indica 'Carpeta de l'Excel d'activitats'.")
    }
    TipA      = { Get-AutoTipText ("es preparen sols, a Esborranys de l'Outlook, els correus dels controls peri" + [char]0x00F2 + "dics ven" + [char]0x00E7 + "uts (mai s'envien sols)") 'controlsperiodics' }
    TipM      = ("Mode MANUAL: nom" + [char]0x00E9 + "s es fa quan cliques la rajola. Clica per posar-ho en autom" + [char]0x00E0 + "tic.")
}
