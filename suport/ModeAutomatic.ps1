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

# L'ultim venciment que ja hauria d'estar servit a l'hora $ara. PURA.
function _AutoVenciment([datetime]$ara, [int]$hora, [int]$minut) {
    $avui = New-Object datetime($ara.Year, $ara.Month, $ara.Day, $hora, $minut, 0)
    if ($ara -lt $avui) { return $avui.AddDays(-1) }
    return $avui
}

# Toca fer la passada? $ultim: la marca de l'ultima passada automatica (text
# ISO; buida si no s'ha fet mai). PURA.
#   menu obert a l'hora      -> el venciment passa a ser el d'avui i toca
#   ahir el PC estava apagat -> el venciment d'ahir no es va servir, toca
#   obres a la tarda i el d'avui no s'ha fet -> toca (no s'espera a dema)
function _AutoToca([datetime]$ara, $ultim, [int]$hora, [int]$minut) {
    $venc = _AutoVenciment $ara $hora $minut
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
