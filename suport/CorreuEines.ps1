#requires -Version 5.1
<#
.SYNOPSIS
  Els correus de CADA EINA: per on surten, a qui van, la CCO i el correu de
  prova (octubre 2026).

.DESCRIPTION
  L'usuari: "A Configuracio s'ha de poder triar entre les opcions (Outlook
  esborrany, Outlook enviament, EmailJS) segons quina eina s'estigui usant [...]
  totes aquelles opcions que permeten enviar un correu. Tambe s'ha de veure clar
  a qui s'envien els correus: titulars, tecnics, representants legals... i CCO.
  [...] un boto a cadascuna de les opcions per enviar un correu de prova, i
  poder determinar el ID GIA per agafar les dades de la prova i el destinatari
  que rebra aquest correu de prova".

  Abans hi havia UNA via per a tot el PC ('CorreuVia' a settings.json) i cada
  eina decidia pel seu compte a qui escrivia (el titular i el representant de
  l'Excel, sempre). Ara:

    $Script:CorreuEines   el registre: les eines, en UN sol lloc. La pantalla de
                          Configuracio, la dels textos i cada enviament hi beuen.
                          Cada eina diu QUINES VIES admet (Controls periodics
                          nomes desa esborranys -l'usuari els revisa-; el
                          Seguiment porta adjunts i EmailJS no en sap).
    correus.json          el que l'usuari hi ha triat, a %LOCALAPPDATA%\
                          InformesCornella (no a settings.json: el "Desar" de
                          Configuracio el reescriu sencer, i les adreces no han
                          d'anar mai al repositori, que es public).
    $Script:CorreuProves  el correu de PROVA de cada eina: el registra cada
                          eina al seu fitxer (aqui no se'n pot cridar cap
                          funcio sense fer un cicle de dependencies: totes
                          criden Get-CorreuEina).

  La via global ('CorreuVia', CorreuVia.ps1) es queda com a valor per defecte
  d'una eina que encara no s'ha configurat: qui ja l'havia triada no veu cap
  canvi.

  NOMES DEFINEIX FUNCIONS.
#>

# Les eines que envien correus. Dest = a qui pot anar (clau de
# $Script:CorreuDestNoms); DestDefecte = el que hi havia fins ara.
$Script:CorreuEines = [ordered]@{
    'mobil' = [ordered]@{
        Nom = ("Enviar correu (informe de visita d'inspecci" + [char]0x00F3 + ' / m' + [char]0x00F2 + 'bil)')
        Vies = @('emailjs', 'outlook-esborrany', 'outlook')
        Dest = @('titular', 'representant', 'autoritzats'); DestDefecte = @('titular', 'representant')
    }
    'rec-requeriments' = [ordered]@{
        Nom = ('Recordatoris ' + [char]0x00B7 + ' Requeriments')
        Vies = @('emailjs', 'outlook-esborrany', 'outlook')
        Dest = @('titular', 'representant', 'autoritzats'); DestDefecte = @('titular', 'representant')
    }
    'rec-precintes' = [ordered]@{
        Nom = ('Recordatoris ' + [char]0x00B7 + ' Precintes')
        Vies = @('emailjs', 'outlook-esborrany', 'outlook')
        Dest = @('titular', 'representant', 'autoritzats'); DestDefecte = @('titular', 'representant')
    }
    'controls' = [ordered]@{
        Nom = ('Controls peri' + [char]0x00F2 + 'dics')
        # NOMES ESBORRANYS: l'usuari els revisa i els envia ell (hi ha guard).
        Vies = @('outlook-esborrany')
        Dest = @('titular', 'representant', 'autoritzats'); DestDefecte = @('titular', 'representant')
    }
    'seguiment' = [ordered]@{
        Nom = 'Seguiment (llistats del GIA, cada dues setmanes)'
        # Porta els PDF adjunts: EmailJS no en sap.
        Vies = @('outlook-esborrany', 'outlook')
        # No es de cap activitat: nomes adreces fixes.
        Dest = @(); DestDefecte = @()
    }
}

$Script:CorreuDestNoms = [ordered]@{
    'titular'      = 'Titular'
    'representant' = 'Representant legal'
    'autoritzats'  = ('Persones autoritzades / t' + [char]0x00E8 + 'cnic')
}

# El correu de prova de cada eina: clau -> { param($gia, $cache) @{ Assumpte;
# Html; Adjunts; Destinataris (els que tindria de debo) } }. L'omple cada eina.
if ($null -eq $Script:CorreuProves) { $Script:CorreuProves = @{} }
# I la CCO que l'eina feia servir fins ara (abans de correus.json), per
# ensenyar-la a la casella de Configuracio: clau -> { '' o "a; b" }.
if ($null -eq $Script:CorreuCcoAbans) { $Script:CorreuCcoAbans = @{} }

# Per a una llista d'adreces: la primera a "Per a" i la resta a "CC" (Controls
# periodics: el titular a Per a, el representant a CC). PURA.
function _CorreuParteixToCc($adreces) {
    $l = @(@($adreces) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -like '*@*' })
    if ($l.Count -eq 0) { return @{ To = ''; Cc = ''; Ok = $false } }
    return @{ To = [string]$l[0]; Cc = ((@($l | Select-Object -Skip 1)) -join '; '); Ok = $true }
}

# ----------------------------------------------------------------------------
# EL FITXER (correus.json)
# ----------------------------------------------------------------------------
function _CorreuEinesPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'correus.json'))
}

# @{ eines = @{ clau = @{ via; dest; fixes; cco } }; prova = @{ gia; desti } }.
# Mai peta: sense fitxer (o trencat), tot buit.
function Read-CorreuEines {
    $o = $null
    try { $o = Read-JsonFile (_CorreuEinesPath) } catch { $o = $null }
    $out = @{ eines = @{}; prova = @{ gia = ''; desti = '' } }
    if ($null -eq $o) { return $out }
    $e = _PropInf $o 'eines'
    if ($null -ne $e) { foreach ($p in @($e.PSObject.Properties)) { $out.eines[[string]$p.Name] = ConvertTo-Mapa $p.Value } }
    $pr = _PropInf $o 'prova'
    if ($null -ne $pr) { $out.prova = @{ gia = [string](_PropInf $pr 'gia'); desti = [string](_PropInf $pr 'desti') } }
    return $out
}

function Save-CorreuEines($cfg) {
    $eines = [ordered]@{}
    foreach ($k in @($Script:CorreuEines.Keys)) {
        if (-not $cfg.eines.ContainsKey($k)) { continue }
        $c = $cfg.eines[$k]
        $o = [ordered]@{ via = [string]$c['via']; dest = [string[]]@(@($c['dest']) | Where-Object { $_ }); fixes = [string]$c['fixes'] }
        # 'cco' nomes si s'ha configurat: sense, l'eina fa servir la d'abans.
        if ($c.ContainsKey('cco') -and $null -ne $c['cco']) { $o['cco'] = [string]$c['cco'] }
        $eines[$k] = $o
    }
    Write-JsonFile (_CorreuEinesPath) ([ordered]@{ eines = $eines; prova = [ordered]@{ gia = [string]$cfg.prova.gia; desti = [string]$cfg.prova.desti } }) 6
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# La configuracio d'una eina, normalitzada contra el registre. $raw: el que hi
# ha a correus.json (o $null); $viaGlobal: la via de tot el PC (la d'abans).
# Torna @{ Clau; Via; Dest; Fixes; Cco ($null = no configurada) }. Una via que
# l'eina no admet cau a la primera que si (Controls periodics: esborrany).
function _CorreuEinaNormalitza($raw, [string]$clau, [string]$viaGlobal) {
    $def = $Script:CorreuEines[$clau]
    if ($null -eq $def) { throw "Eina de correu desconeguda: $clau" }
    $r = ConvertTo-Mapa $raw
    $vies = @($def.Vies)
    $via = (_CorreuViaValida ([string]$r['via']))
    if ([string]::IsNullOrWhiteSpace([string]$r['via'])) { $via = _CorreuViaValida $viaGlobal }
    if ($vies -notcontains $via) { $via = [string]$vies[0] }
    $dest = if ($r.ContainsKey('dest') -and $null -ne $r['dest']) { @(@($r['dest']) | ForEach-Object { [string]$_ } | Where-Object { @($def.Dest) -contains $_ }) } else { @($def.DestDefecte) }
    $cco = if ($r.ContainsKey('cco') -and $null -ne $r['cco']) { _CorreuNetejaAdreces ([string]$r['cco']) } else { $null }
    return @{ Clau = $clau; Via = $via; Dest = $dest; Fixes = (_CorreuNetejaAdreces ([string]$r['fixes'])); Cco = $cco }
}

# Adreces separades per ';' o ',' -> "a@b; c@d", sense repetides ni buides.
# PURA.
function _CorreuNetejaAdreces([string]$s) {
    return ((@(_CorreuLlistaAdreces $s)) -join '; ')
}

function _CorreuLlistaAdreces([string]$s) {
    $vistes = @{}
    $out = New-Object System.Collections.ArrayList
    foreach ($a in @(([string]$s) -split '[;,\s]+')) {
        $a = $a.Trim()
        if ($a -notlike '*@*') { continue }
        $k = $a.ToLowerInvariant()
        if ($vistes.ContainsKey($k)) { continue }
        $vistes[$k] = $true
        [void]$out.Add($a)
    }
    # Array PLA: el cridador hi posa @() (una funcio que torna amb coma no es
    # pot embolcallar amb @(), hi ha guard).
    return $out.ToArray()
}

# A QUI VA un correu d'una eina, per a una activitat. $emails: @{ titular;
# representant } (de l'Excel); $autoritzats: les adreces de les persones
# autoritzades (base de contactes). Torna les adreces EN ORDRE (titular,
# representant, autoritzats, fixes), sense repetides. PURA.
function _CorreuDestinataris($cfgEina, $emails, $autoritzats = @()) {
    $tots = New-Object System.Collections.ArrayList
    $e = ConvertTo-Mapa $emails
    foreach ($d in @($cfgEina.Dest)) {
        switch ([string]$d) {
            'titular'      { [void]$tots.Add([string]$e['titular']) }
            'representant' { [void]$tots.Add([string]$e['representant']) }
            'autoritzats'  { foreach ($a in @($autoritzats)) { [void]$tots.Add([string]$a) } }
        }
    }
    [void]$tots.Add([string]$cfgEina.Fixes)
    # Array PLA (el cridador hi posa @()).
    return @(_CorreuLlistaAdreces ($tots -join ';'))
}

# El text que diu a qui va (Configuracio i les confirmacions). PURA.
function _CorreuDestText($cfgEina) {
    $parts = New-Object System.Collections.ArrayList
    foreach ($d in @($cfgEina.Dest)) { [void]$parts.Add([string]$Script:CorreuDestNoms[[string]$d]) }
    if ([string]$cfgEina.Fixes) { [void]$parts.Add([string]$cfgEina.Fixes) }
    if ($parts.Count -eq 0) { return '(ning' + [char]0x00FA + ')' }
    return ($parts -join ', ')
}

# El correu de PROVA: va NOMES a $desti, i al davant diu a qui hauria anat de
# debo. PURA.
function _CorreuProvaHtml([string]$html, [string]$nomEina, $destinatarisReals, [string]$cco) {
    $reals = (@($destinatarisReals) | Where-Object { $_ }) -join '; '
    if (-not $reals) { $reals = '(cap adre' + [char]0x00E7 + 'a)' }
    $esc = { param($t) [System.Net.WebUtility]::HtmlEncode([string]$t) }
    $cap = '<div style="border:2px solid #b00020;padding:8px;margin-bottom:12px;font-family:Arial,sans-serif;font-size:10pt">' +
           '<b>CORREU DE PROVA</b> &middot; ' + (& $esc $nomEina) + '<br>' +
           'De debo hauria anat a: ' + (& $esc $reals) +
           $(if ($cco) { '<br>CCO: ' + (& $esc $cco) } else { '' }) + '</div>'
    return ($cap + [string]$html)
}

# ----------------------------------------------------------------------------
# EL QUE FAN SERVIR LES EINES
# ----------------------------------------------------------------------------
function Get-CorreuEina([string]$clau) {
    $cfg = Read-CorreuEines
    $raw = if ($cfg.eines.ContainsKey($clau)) { $cfg.eines[$clau] } else { $null }
    return (_CorreuEinaNormalitza $raw $clau (Get-CorreuVia))
}

# Desa nomes la via d'una eina (el desplegable d'"Enviar correu").
function Set-CorreuEinaVia([string]$clau, [string]$via) {
    $cfg = Read-CorreuEines
    $c = if ($cfg.eines.ContainsKey($clau)) { $cfg.eines[$clau] } else { @{} }
    $c['via'] = _CorreuViaValida $via
    $cfg.eines[$clau] = $c
    try { Save-CorreuEines $cfg; return $true } catch { return $false }
}

# Les adreces de les persones autoritzades d'una activitat (la base de
# contactes, ContactesDb.ps1). Sense base, cap.
function Get-CorreuAutoritzats([string]$gia) {
    if (-not (Get-Command Get-ContactesAutoritzatsEmails -ErrorAction SilentlyContinue)) { return @() }
    try { return @(Get-ContactesAutoritzatsEmails $gia) } catch { return @() }
}

# Les adreces de l'Excel d'una activitat (la fila de la cache).
function _CorreuEmailsDeAct($act) {
    $t = ''; $r = ''
    if ($null -ne $act) {
        try { if ($act.ContainsKey('EMAIL'))     { $t = [string]$act['EMAIL'] } } catch { }
        try { if ($act.ContainsKey('EMAIL_REP')) { $r = [string]$act['EMAIL_REP'] } } catch { }
    }
    return @{ titular = $t; representant = $r }
}

# ENVIA EL CORREU DE PROVA d'una eina, amb les dades de l'activitat de prova, a
# l'adreca de prova. Mai a cap destinatari de debo ni amb CCO: la CCO real es
# diu al cap del correu. Torna @{ Ok; Text }.
function Send-CorreuProva([string]$clau, [string]$gia, [string]$desti) {
    $def = $Script:CorreuEines[$clau]
    if ($null -eq $def) { return @{ Ok = $false; Text = "Eina desconeguda: $clau" } }
    $desti = _CorreuNetejaAdreces $desti
    if (-not $desti) { return @{ Ok = $false; Text = "Escriu l'adre" + [char]0x00E7 + 'a que ha de rebre la prova.' } }
    $prova = $Script:CorreuProves[$clau]
    if ($null -eq $prova) { return @{ Ok = $false; Text = "Aquesta eina encara no t" + [char]0x00E9 + " correu de prova." } }
    $cfgE = Get-CorreuEina $clau
    $cache = $null
    try {
        $xls = Find-LatestActivitatsExcel
        if ($null -ne $xls) { $cache = Initialize-ActivitatsCache $xls.File }
    } catch { $cache = $null }
    $m = $null
    try { $m = & $prova ([string]$gia).Trim() $cache $cfgE } catch { return @{ Ok = $false; Text = [string]$_.Exception.Message } }
    if ($null -eq $m) { return @{ Ok = $false; Text = "No s'ha pogut muntar el correu de prova." } }
    $cco = if ($null -ne $cfgE.Cco) { [string]$cfgE.Cco } else { [string]$m.CcoAbans }
    $html = _CorreuProvaHtml ([string]$m.Html) ([string]$def.Nom) @($m.Destinataris) $cco
    $motiu = Test-CorreuViaLlest $cfgE.Via
    if ($motiu) { return @{ Ok = $false; Text = $motiu } }
    $ses = $null
    try {
        $ses = Open-CorreuSessio $cfgE.Via
        Send-CorreuSessio $ses $desti '' ('[PROVA] ' + [string]$m.Assumpte) $html '' @($m.Adjunts)
        $com = if ($cfgE.Via -eq 'outlook-esborrany') { "desat a Esborranys de l'Outlook" } else { 'enviat' }
        return @{ Ok = $true; Text = ("Correu de prova " + $com + " per a " + $desti + " (" + (_CorreuViaText $cfgE.Via) + ").") }
    } catch {
        $txt = if ($null -eq $ses) { [string]$_.Exception.Message } else { _CorreuSessioError $ses $_ }
        return @{ Ok = $false; Text = $txt }
    } finally {
        Close-CorreuSessio $ses
    }
}
