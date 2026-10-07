#requires -Version 5.1
<#
.SYNOPSIS
  El que tenen en comu les eines que es poden fer SOLES (interruptor A/M del
  menu): "Copiar informes" i "Actualitzar base".

.DESCRIPTION
  Fins a l'octubre de 2026 nomes "Copiar informes" tenia mode automatic i tot
  aixo vivia a CopiaInformes.ps1, escrit per a ella. Quan "Actualitzar base" en
  va voler un (l'usuari: "sera tan important per fer el Planol activitats"),
  copiar-ho era la manera segura que les dues versions divergissin.

  Aqui hi ha nomes el que es IGUAL; cada eina posa la seva hora, el seu fitxer
  d'estat i la seva feina:
    _AutoVenciment / _AutoToca   quan toca (una sola pregunta: "des de l'ultim
                                 venciment, s'ha fet cap passada?")
    Read-EstatAuto / Save-EstatAuto  l'estat en JSON (NOMES les claus que es
                                 donen, damunt del que ja hi ha)
    Write-AutoLog                el registre (un mode que no ensenya res nomes
                                 es pot seguir aixi)
    Start-ProcesAutoUnic         llanca la passada EN SEGON PLA, una a la vegada
    Invoke-AmbMutexUnic          dins del proces: si ja n'hi ha una en marxa,
                                 no fa res (no espera)

  I EL REGISTRE ($Script:ModesAuto): cada eina amb interruptor s'hi apunta en
  carregar-se, i el menu nomes el recorre (no sap res de cap eina).

  NOMES DEFINEIX FUNCIONS (i el registre, buit).
#>

# Per a cada accio del menu (la clau), un hashtable amb:
#   Titol      per als avisos
#   Actiu      { } -> [bool]: l'interruptor
#   DesaActiu  { param($on) }
#   UltimMode  { } -> 'auto' | 'manual' | '': qui va fer l'ultima passada (el
#              menu pinta la data en verd si va ser l'automatic)
#   SiToca     { }: si l'interruptor esta ences i toca, llanca la passada
#   Requisit   { } -> '' si es pot engegar; si no, el text de l'avis
#   TipA, TipM l'ajuda de l'interruptor en automatic i en manual
# Els scriptblocks NO son closures: criden funcions, i un $Script: dins d'una
# closure no es el de l'script (vegeu CLAUDE.md).
$Script:ModesAuto = [ordered]@{}

# L'HORA DE TOTS ELS MODES AUTOMATICS, en un sol lloc (octubre 2026, l'usuari:
# "a les 13.00, a totes les eines que es llancen automaticament"). La fan servir
# Copiar informes, Actualitzar base i la tasca del Windows dels Recordatoris.
# Abans cada una tenia la seva (14:30, 14:00, 09:00).
#
# LA REGLA ES LA MATEIXA PER A TOTES: es fa a aquesta hora i, si l'ULTIMA
# VEGADA QUE TOCAVA no es va poder fer (el programa tancat, el PC apagat), es
# fa tan aviat com es pot: en obrir el programa (Copiar informes, Actualitzar
# base) o en engegar el PC (la tasca dels recordatoris, StartWhenAvailable).
$Script:AutoHora  = 13
$Script:AutoMinut = 0

function Get-AutoHoraText { return ('{0:00}:{1:00}' -f [int]$Script:AutoHora, [int]$Script:AutoMinut) }

# ----------------------------------------------------------------------------
# LA PROGRAMACIO DE CADA AUTOMATISME (octubre 2026, l'usuari: "aquests
# automatismes, com son ja uns quants, haurien de ser configurables des de la
# configuracio", i el Planol activitats "un cop a la setmana")
# ----------------------------------------------------------------------------
# Cada automatisme s'hi apunta en carregar-se amb la seva programacio PER
# DEFECTE; la pantalla de Configuracio en desa una de propia per a aquest PC
# (settings.json, clau "Automatismes") i es llegeix EN VIU (cada minut, el menu):
# canviar-la no demana reiniciar.
#   Freq  'dia' (cada dia) | 'setmana' (un dia de la setmana)
#   Dia   1 dilluns ... 7 diumenge (nomes 'setmana')
#   Hora  'HH:mm'
# La regla de sempre no canvia: si l'ULTIMA VEGADA QUE TOCAVA no es va fer, es
# fa tan aviat com es pot.
$Script:ProgramacionsAuto = [ordered]@{}
$Script:DiesSetmana = @('', 'dilluns', 'dimarts', 'dimecres', 'dijous', 'divendres', 'dissabte', 'diumenge')

function Register-ProgramacioAuto([string]$clau, [string]$titol, [string]$freq = 'dia', [int]$dia = 1, [string]$hora = '') {
    if ($hora -eq '') { $hora = Get-AutoHoraText }
    $Script:ProgramacionsAuto[$clau] = [pscustomobject]@{ Clau = $clau; Titol = $titol; Freq = $freq; Dia = $dia; Hora = $hora }
}

# Una programacio es valida? (el que ve de settings.json es d'un fitxer que es
# pot tocar a ma). PURA.
function Test-ProgramacioValida($p) {
    if ($null -eq $p) { return $false }
    if (@('dia', 'setmana') -notcontains [string]$p.Freq) { return $false }
    if ([string]$p.Hora -notmatch '^([01]?\d|2[0-3]):[0-5]\d$') { return $false }
    if ([string]$p.Freq -eq 'setmana') { $d = 0; if (-not [int]::TryParse([string]$p.Dia, [ref]$d) -or $d -lt 1 -or $d -gt 7) { return $false } }
    return $true
}

# La programacio EFECTIVA: la d'aquest PC (settings.json) si n'hi ha una de
# valida, si no la per defecte. $settings: per a les proves (si no, es llegeix).
function Get-ProgramacioAuto([string]$clau, $settings = $null) {
    $def = $Script:ProgramacionsAuto[$clau]
    if ($null -eq $def) { $def = [pscustomobject]@{ Clau = $clau; Titol = $clau; Freq = 'dia'; Dia = 1; Hora = (Get-AutoHoraText) } }
    if ($null -eq $settings) { try { $settings = Load-AppSettings } catch { $settings = $null } }
    $o = $null
    try { if ($null -ne $settings -and $null -ne $settings.Automatismes) { $o = $settings.Automatismes.$clau } } catch { $o = $null }
    if (Test-ProgramacioValida $o) {
        $h = ([string]$o.Hora).Split(':')
        return [pscustomobject]@{ Clau = $clau; Titol = $def.Titol; Freq = [string]$o.Freq; Dia = [int]$o.Dia
                                  Hora = ('{0:00}:{1:00}' -f [int]$h[0], [int]$h[1]); Propia = $true }
    }
    return [pscustomobject]@{ Clau = $clau; Titol = $def.Titol; Freq = $def.Freq; Dia = [int]$def.Dia; Hora = $def.Hora; Propia = $false }
}

# "cada dia a les 13:00" / "cada dilluns a les 13:00". PURA (amb $p donada).
function Get-ProgramacioText($p) {
    if ($p -is [string]) { $p = Get-ProgramacioAuto $p }
    if ([string]$p.Freq -eq 'setmana') { return ('cada ' + $Script:DiesSetmana[[int]$p.Dia] + ' a les ' + $p.Hora) }
    return ('cada dia a les ' + $p.Hora)
}

# L'hora, el minut i el dia de la setmana (0 = cada dia) d'una programacio. PURA.
function _ProgramacioParts($p) {
    $h = ([string]$p.Hora).Split(':')
    $dia = if ([string]$p.Freq -eq 'setmana') { [int]$p.Dia } else { 0 }
    return @([int]$h[0], [int]$h[1], $dia)
}

# Toca fer la passada d'aquest automatisme, amb la seva programacio?
function Test-ProgramacioToca([string]$clau, [datetime]$ara, $ultim, $settings = $null) {
    $x = _ProgramacioParts (Get-ProgramacioAuto $clau $settings)
    return (_AutoToca $ara $ultim $x[0] $x[1] $x[2])
}

# L'ultim venciment d'aquest automatisme a l'hora $ara.
function Get-ProgramacioVenciment([string]$clau, [datetime]$ara, $settings = $null) {
    $x = _ProgramacioParts (Get-ProgramacioAuto $clau $settings)
    return (_AutoVenciment $ara $x[0] $x[1] $x[2])
}

# El que es desa a settings.json: nomes les programacions DIFERENTS de la per
# defecte (com les carpetes, _BuildSettingsOverrides). $valors: clau ->
# { Freq; Dia; Hora }. PURA (amb el registre).
function ConvertTo-AutomatismesSettings($valors) {
    $out = [ordered]@{}
    foreach ($k in @($valors.Keys)) {
        $v = $valors[$k]; $d = $Script:ProgramacionsAuto[$k]
        if (-not (Test-ProgramacioValida $v)) { continue }
        $igual = ($null -ne $d -and [string]$v.Freq -eq [string]$d.Freq -and [string]$v.Hora -eq [string]$d.Hora -and
                  ([string]$v.Freq -ne 'setmana' -or [int]$v.Dia -eq [int]$d.Dia))
        if (-not $igual) { $out[$k] = [ordered]@{ Freq = [string]$v.Freq; Dia = [int]$v.Dia; Hora = [string]$v.Hora } }
    }
    return $out
}

# L'ajuda de l'interruptor en AUTOMATIC. $que: el que es fa ("es copia sol",
# "la base s'actualitza sola"); $clau: l'automatisme (la seva programacio; sense,
# la de per defecte, cada dia).
function Get-AutoTipText([string]$que, [string]$clau = '') {
    $quan = if ($clau -ne '') { Get-ProgramacioText $clau } else { 'cada dia a les ' + (Get-AutoHoraText) }
    return ("Mode AUTOMATIC: " + $que + " " + $quan + ". Si l'ultima vegada que tocava " +
            "no es va poder fer (el programa estava tancat), es fa en obrir-lo. Clica per passar a manual. " +
            "Quan es fa, es canvia a Configuracio.")
}

# L'ultim venciment que ja hauria d'estar servit a l'hora $ara. $diaSetmana: 0
# cada dia; 1..7 (dilluns..diumenge) un cop a la setmana. PURA.
function _AutoVenciment([datetime]$ara, [int]$hora, [int]$minut, [int]$diaSetmana = 0) {
    $avui = New-Object datetime($ara.Year, $ara.Month, $ara.Day, $hora, $minut, 0)
    if ($diaSetmana -lt 1) {
        if ($ara -lt $avui) { return $avui.AddDays(-1) }
        return $avui
    }
    # DayOfWeek: diumenge = 0. En ISO, diumenge = 7.
    $iso = [int]$ara.DayOfWeek; if ($iso -eq 0) { $iso = 7 }
    $venc = $avui.AddDays( - (($iso - $diaSetmana + 7) % 7))
    if ($venc -gt $ara) { $venc = $venc.AddDays(-7) }
    return $venc
}

# Toca fer la passada? $ultim: la marca de l'ultima passada automatica (text
# ISO; buida si no s'ha fet mai). PURA.
#   menu obert a l'hora      -> el venciment passa a ser el d'avui i toca
#   ahir el PC estava apagat -> el venciment d'ahir no es va servir, toca
#   obres a la tarda i el d'avui no s'ha fet -> toca (no s'espera a dema)
# Amb $diaSetmana, el mateix per setmanes.
function _AutoToca([datetime]$ara, $ultim, [int]$hora, [int]$minut, [int]$diaSetmana = 0) {
    $venc = _AutoVenciment $ara $hora $minut $diaSetmana
    $t = [string]$ultim
    if ([string]::IsNullOrWhiteSpace($t)) { return $true }
    try { $fet = [datetime]::Parse($t) } catch { return $true }
    return ($fet -lt $venc)
}

# L'estat d'una eina: SEMPRE el diccionari sencer de la plantilla (els valors
# per defecte si no hi ha fitxer o esta corrupte), aixi cap crider ha de mirar si
# una clau hi es. Les claus amb valor per defecte [bool] es llegeixen com a bool;
# la resta, com a text (les dates, normalitzades amb Read-JsonIso).
function Read-EstatAuto([string]$path, $plantilla) {
    $out = [ordered]@{}
    foreach ($k in @($plantilla.Keys)) { $out[$k] = $plantilla[$k] }
    if ([string]::IsNullOrWhiteSpace($path)) { return $out }
    $o = Read-JsonFile $path
    if ($null -eq $o) { return $out }
    foreach ($k in @($plantilla.Keys)) {
        if (-not $o.PSObject.Properties[$k]) { continue }
        if ($plantilla[$k] -is [bool]) { $out[$k] = [bool]$o.$k } else { $out[$k] = [string](Read-JsonIso $o.$k) }
    }
    return $out
}

# Desa NOMES les claus que es donen, damunt del que ja hi ha: encendre
# l'interruptor no pot esborrar la data de l'ultima passada, ni al reves. Si la
# plantilla te 'generat_el' i es buit, s'hi posa ara. Mai llanca: si la carpeta
# no hi es (unitat de xarxa fora de servei), el programa ha de seguir igual.
function Save-EstatAuto([string]$path, $canvis, $plantilla) {
    try {
        if ([string]::IsNullOrWhiteSpace($path)) { return $false }
        $est = Read-EstatAuto $path $plantilla
        if ($null -ne $canvis) { foreach ($k in @($canvis.Keys)) { $est[$k] = $canvis[$k] } }
        if ($est.Contains('generat_el') -and [string]::IsNullOrWhiteSpace([string]$est['generat_el'])) { $est['generat_el'] = (Get-Date).ToString('o') }
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile $path ([pscustomobject]$est) 5
        return $true
    } catch { return $false }
}

# El registre d'un mode automatic, a %LOCALAPPDATA%\InformesCornella\<fitxer>.
function Get-AutoLogPath([string]$fitxer) {
    return [string](Join-Path (Join-Path $env:LOCALAPPDATA 'InformesCornella') $fitxer)
}

function Write-AutoLog([string]$fitxer, [string]$msg) {
    try {
        $p = Get-AutoLogPath $fitxer
        $dir = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $p -Value ('[' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '] ' + [string]$msg) -Encoding UTF8
    } catch { }
}

# Llanca la passada EN SEGON PLA i torna de seguida: el menu no es pot quedar
# congelat mentre es recorre la carpeta d'informes, ni fer la feina dins d'un
# tick del rellotge (amb els DoEvents, l'usuari podria obrir una eina a mitges).
# NOMES UNA PER EINA: si l'anterior encara corre (el menu es torna a obrir a cada
# volta de Main), no se'n llanca una altra. $clau separa les eines.
$Script:ProcAuto = @{}

function Start-ProcesAutoUnic([string]$clau, [string]$script) {
    try {
        $p = $Script:ProcAuto[$clau]
        if ($null -ne $p -and -not $p.HasExited) { return $false }
    } catch { }
    $Script:ProcAuto[$clau] = Start-ScriptSegonPla $script
    return ($null -ne $Script:ProcAuto[$clau])
}

# ESPERA que un altre proces acabi la feina d'un mutex (fins a $maxSegons) i
# torna: no la fa ni el reté. Torna $true si es lliure (o ho ha quedat) i
# $false si s'ha cansat d'esperar. Els Recordatoris automatics ho fan amb el de
# la base d'informes: si "Actualitzar base" corre a la mateixa hora, que llegeixin
# la base acabada de fer i no la d'ahir.
function Wait-MutexLliure([string]$nom, [int]$maxSegons) {
    $mutex = $null
    try {
        $mutex = New-Object System.Threading.Mutex($false, $nom)
        $tinc = $false
        try { $tinc = $mutex.WaitOne([int]([Math]::Max(0, $maxSegons) * 1000)) } catch [System.Threading.AbandonedMutexException] { $tinc = $true }
        if ($tinc) { try { $mutex.ReleaseMutex() } catch { } }
        return [bool]$tinc
    } catch { return $true } finally {
        if ($null -ne $mutex) { try { $mutex.Dispose() } catch { } }
    }
}

# Fa $feina NOMES si ningu mes no la fa (un mutex amb nom, per a tot l'ordinador):
# el menu ja mira el seu proces, pero el handle es d'aquella instancia del
# programa. NO espera: torna $false si ja n'hi ha un altre en marxa (i crida
# $siOcupat, si n'hi ha). Les excepcions de $feina es propaguen; el mutex
# s'allibera igualment.
function Invoke-AmbMutexUnic([string]$nom, [scriptblock]$feina, [scriptblock]$siOcupat = $null) {
    $mutex = $null
    $tinc = $false
    try {
        $mutex = New-Object System.Threading.Mutex($false, $nom)
        try { $tinc = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $tinc = $true }
    } catch { $mutex = $null; $tinc = $true }
    if (-not $tinc) {
        if ($null -ne $siOcupat) { & $siOcupat }
        return $false
    }
    try {
        $null = & $feina
        return $true
    } finally {
        if ($null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { }; try { $mutex.Dispose() } catch { } }
    }
}
