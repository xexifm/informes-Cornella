#requires -Version 5.1
<#
  SeguimentGiaAutomatic.ps1 - El MODE AUTOMATIC del "Seguiment" (l'interruptor
  A/M de sota la rajola) i el CORREU que porta els llistats.

  L'usuari (octubre 2026): "implementar l'opcio d'automatitzar l'eina Seguiment
  cada dues setmanes + enviament del document adjunt per correu". Per defecte,
  dilluns alternats a les 13:00 (es canvia a Configuracio -> Automatismes, com
  tots, ara amb l'opcio "cada dues setmanes"). Mateixa regla que la resta: si
  l'ultima vegada que tocava no es va fer, es fa en obrir el programa.

  LA PASSADA es fa EN SEGON PLA (SeguimentGiaAuto.ps1): genera el PDF dels
  llistats (TOTS menys la copia de la fulla "Estes": l'usuari, "nomes els PDF,
  tot menys la fulla Estes"), amb la MATEIXA funcio que el boto
  (_SgGeneraFitxer), i l'envia als destinataris de l'eina 'seguiment'
  (Configuracio -> Correus de cada eina: adreces fixes i CCO; nomes Outlook,
  perque porta l'adjunt). Si va a Esborranys, el menu ho avisa com amb els
  recordatoris.

  EL TEXT DEL CORREU es edita a "Textos del correu" (CorreuEinesPantalla.ps1) i
  es desa a %LOCALAPPDATA%\InformesCornella\seguiment-correu.json.

  L'ESTAT, a local\seguiment-gia\seguiment-auto.json:
    auto       l'interruptor
    auto_el    l'ultima PASSADA automatica (encara que no hagi pogut fer res)
    mode       'auto' | 'manual': qui va fer l'ultim enviament
    enviat_el  l'ultim enviament

  El correu (el text, l'enviament) i l'estat son a SeguimentGia.ps1: els fa
  servir tambe el boto "PDF i enviar" de la finestra. Aqui, l'automatisme.
  Es carrega DESPRES de CorreuEines.ps1 (s'hi registra la prova).
  NOMES DEFINEIX FUNCIONS (i s'apunta als registres).
#>

$Script:SgMutexNom = 'Global\InformesCornella.Seguiment'

function _SgAutoLog([string]$msg) { Write-AutoLog 'seguiment-log.txt' $msg }

# LA PASSADA AUTOMATICA (la crida SeguimentGiaAuto.ps1). Cap finestra: tot al
# registre. Apunta 'auto_el' SEMPRE (si no, el menu la tornaria a llancar cada
# minut). Torna @{ Ok; Error }.
function Invoke-SeguimentAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $caixa = @{ Gen = $null; Env = $null }
    $lliure = Invoke-AmbMutexUnic $Script:SgMutexNom {
        $caixa.Gen = _SgGeneraFitxer 'pdf' @(_SgSeleccioAuto)
        if ([bool]$caixa.Gen.Ok) { $caixa.Env = Send-SeguimentCorreu $caixa.Gen $true }
    }
    [void](_SgAutoDesaEstat $ini)
    if (-not $lliure) { _SgAutoLog 'Passada automatica: ja se n''esta fent un, no es fa res.'; return @{ Ok = $false; Error = 'ocupat' } }
    if (-not [bool]$caixa.Gen.Ok) {
        _SgAutoLog ('ATURAT: ' + ([string]$caixa.Gen.Error -replace "`r?`n", ' '))
        return @{ Ok = $false; Error = [string]$caixa.Gen.Error }
    }
    _SgAutoLog ('PDF: ' + [string]$caixa.Gen.Path + ' ' + [char]0x00B7 + ' ' + [string]$caixa.Env.Text)
    return @{ Ok = [bool]$caixa.Env.Ok; Error = $(if ($caixa.Env.Ok) { '' } else { [string]$caixa.Env.Text }) }
}

function Start-SeguimentAuto { return (Start-ProcesAutoUnic 'seguiment' 'SeguimentGiaAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-SeguimentAutoSiToca {
    $est = _SgAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (Test-ProgramacioToca 'seguimentgia' (Get-Date) $est['auto_el'])) { return $false }
    return (Start-SeguimentAuto)
}

# ----------------------------------------------------------------------------
# ELS REGISTRES
# ----------------------------------------------------------------------------
Register-ProgramacioAuto 'seguimentgia' 'Seguiment' 'quinzena' 1
$Script:ModesAuto['seguimentgia'] = @{
    Titol     = 'Seguiment'
    Actiu     = { $e = _SgAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _SgAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _SgAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-SeguimentAutoSiToca }
    # Sense destinataris l'automatic faria el PDF i no l'enviaria a ningu.
    Requisit  = {
        if ([string](Get-CorreuEina 'seguiment').Fixes) { return '' }
        return ("Per enviar el seguiment sol cal dir a qui va.`n`nVes a Configuraci" + [char]0x00F3 + " -> Configuraci" + [char]0x00F3 + " correu electr" + [char]0x00F2 + "nic -> Seguiment i posa-hi les adreces (Adreces fixes).")
    }
    TipA      = { Get-AutoTipText "es fa el PDF dels llistats i s'envia per correu" 'seguimentgia' }
    TipM      = ("Mode MANUAL: nom" + [char]0x00E9 + "s es fa quan cliques la rajola. Clica per posar-ho en autom" + [char]0x00E0 + "tic.")
}

# El correu de prova: el text d'aquest correu i, adjunt, el darrer PDF de
# seguiment que hi hagi (sense, va sense adjunt i ho diu).
$Script:CorreuProves['seguiment'] = {
    param($gia, $cache, $cfgE)
    $t = _LoadSgCorreu
    # EL PDF ES FA ARA, com el de l'automatic: les dades del dia (l'usuari,
    # octubre 2026: "s'envia l'ultim arxiu que jo he creat a ma, pero la
    # gracia es que aquell document es faci abans de ser enviat i que estigui
    # actualitzat del mateix dia"). Abans s'hi adjuntava el darrer PDF que hi
    # hagues a la carpeta. Si no es pot fer, la prova surt sense adjunt i diu
    # per que (mai un PDF vell).
    $gen = $null
    try { $gen = _SgGeneraFitxer 'pdf' @(_SgSeleccioAuto) } catch { $gen = @{ Ok = $false; Error = [string]$_.Exception.Message } }
    $ok = ($null -ne $gen -and [bool]$gen.Ok)
    $ara = Get-Date
    $cos = if ($ok) { [string]$t['cos'] } else { _SgCosSenseDades ([string]$t['cos']) }
    $resum = if ($ok) { $gen.Resum } else { [ordered]@{} }
    $html = _CosAHtml (_SgFillPh $cos $resum $(if ($ok) { [string]$gen.Origen } else { '' }) $ara)
    if (-not $ok) { $html = '<p><i>(Sense adjunt: no s''ha pogut fer el PDF. ' + [System.Net.WebUtility]::HtmlEncode(([string]$gen.Error -replace "`r?`n", ' ')) + ')</i></p>' + $html }
    return @{
        Assumpte = (_SgFillPh ([string]$t['assumpte']) $resum '' $ara)
        Html = $html
        Destinataris = @(_CorreuLlistaAdreces ([string]$cfgE.Fixes))
        Adjunts = $(if ($ok) { @([string]$gen.Path) } else { @() })
        CcoAbans = ''
    }
}
$Script:CorreuCcoAbans['seguiment'] = { '' }
$Script:CorreuTextosEditors['seguiment'] = @{ Desc = "El correu amb el PDF dels llistats de seguiment (cada dues setmanes)."; Obre = { Invoke-SeguimentCorreuTextos } }
