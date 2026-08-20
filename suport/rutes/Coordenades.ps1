<#
  Coordenades.ps1 - Eina per REPASSAR i CORREGIR la geolocalitzacio dels
  establiments.

  Que fa:
    1. Localitza el fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' mes recent (xarxa
       de la feina o carpeta local de fallback), fulla "Estes"/"Estes".
    2. Detecta les activitats APILADES: les que comparteixen exactament la
       mateixa coordenada amb alguna altra. Es el problema que venim a
       resoldre -- el Cadastre georeferencia la PARCEL.LA i no el local, aixi
       que els quinze locals d'un edifici cauen tots al mateix punt.
    3. Demana al Cadastre la coordenada de FACANA (el portal) de cada
       activitat, per parcel.la i amb memoria cau (Geocodificador.ps1).
    4. Genera un mapa HTML on es veuen les DUES coordenades alhora:
         VERMELL = la que hi ha ara a l'Excel (centre de la parcel.la)
         VERD    = la de facana, i es pot ARROSSEGAR per corregir-la a ma
    5. Des del mapa et pots baixar un .xlsx amb l'ID GIA, la coordenada vella
       i la nova. El fitxer es genera al mateix navegador.

  Aixo NO toca ni Ruta.ps1 ni Precintades.ps1: aquells segueixen fent servir
  la coordenada original de l'Excel. Aquesta eina nomes MIRA i genera un
  fitxer; no reescriu res de la base de dades.

  Es un programa INDEPENDENT, com Ruta.ps1: es carrega des del menu
  (Motor.ps1 -> Start-CoordenadesTool) pero corre en el seu propi ambit.

  Reutilitza les funcions ja provades de Ruta.ps1 (cerca de l'Excel, cerca de
  la fulla, columnes per nom, conversio UTM -> lat/lon, format d'adreca)
  carregant-lo en mode headless.

  Mode "headless" per a proves: si $env:COORDENADES_TEST o $env:GENINFORME_TEST
  estan definides, NOMES es defineixen les funcions (no s'obre cap finestra ni
  es llegeix cap Excel). Aixi es proven les funcions pures a Linux sense
  Office.

  ATENCIÓ AL BOM: aquest fitxer s'ha de desar en UTF-8 AMB BOM, com Ruta.ps1.
  El Windows PowerShell 5.1 llegeix els .ps1 sense BOM com a ANSI i corromp
  els literals accentuats, i aquí n'hi ha molts: tot el text que l'usuari veu
  al mapa (llegenda, popups, capçaleres de l'Excel que es baixa) viu dins de
  l'HTML que hi ha més avall. Si algun dia surten "Ã§" pel mapa, el primer que
  s'ha de mirar és si el fitxer ha perdut el BOM.
#>

$ErrorActionPreference = 'Stop'

# Headless: nomes definir funcions (proves). Compartim el flag amb Ruta/Motor.
$Script:CoordHeadless = [bool]$env:COORDENADES_TEST -or [bool]$env:GENINFORME_TEST

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
# Aquest script viu a suport/rutes/. Ruta.ps1 i Geocodificador.ps1 son al
# costat; l'arrel del clone es dos nivells amunt (suport/rutes/../..).
$SuportDir  = Split-Path -Parent $ScriptRoot          # suport/
$RepoRoot   = Split-Path -Parent $SuportDir           # informes-Cornella/

# WinForms el carreguem AQUI i no ho deixem en mans de Ruta.ps1: el carregarem
# en mode headless (RUTA_TEST) expressament perque no obri la seva finestra, i
# en aquest mode ell no fa cap Add-Type.
if (-not $Script:CoordHeadless) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

# ----------------------------------------------------------------------------
# Modul de facanes. Es carrega ABANS de Ruta.ps1 EXPRESSAMENT: Ruta.ps1 es qui
# carrega config.ps1, i volem que config.ps1 pugui sobreescriure tambe les
# variables del geocodificador ($GeoDistanciaMaximaM, $GeoCatastroUrlTemplate...).
# Si el carreguessim despres, els seus valors per defecte trepitjarien el que
# l'usuari hagues posat a config.ps1.
# ----------------------------------------------------------------------------
. (Join-Path $ScriptRoot 'Geocodificador.ps1')

# ----------------------------------------------------------------------------
# Reutilitzem les funcions de Ruta.ps1 (Find-LatestRutaExcel, Find-HeaderColumn,
# _RutaFindEstesSheet, ConvertTo-UtmNumber, Format-EmpAddress,
# Convert-UtmToLatLon, _HtmlEncode). El carreguem en mode headless perque NOMES
# defineixi funcions i no obri la seva finestra ni executi la seva Main.
# Restaurem la variable d'entorn despres per no afectar la resta del proces.
# ----------------------------------------------------------------------------
$Script:_prevRutaTestCoord = $env:RUTA_TEST
$env:RUTA_TEST = '1'
try {
    . (Join-Path $ScriptRoot 'Ruta.ps1')
} finally {
    if ($null -eq $Script:_prevRutaTestCoord) {
        Remove-Item Env:\RUTA_TEST -ErrorAction SilentlyContinue
    } else {
        $env:RUTA_TEST = $Script:_prevRutaTestCoord
    }
}

# Icona corporativa per a les finestres d'aquesta eina. Ruta.ps1 nomes la
# carrega quan NO va en headless, i nosaltres l'hi hem fet anar, aixi que ens
# la carreguem pel nostre compte.
$Script:CoordIcon = $null
if (-not $Script:CoordHeadless) {
    try {
        $coordIconPath = Join-Path $SuportDir 'cornella.ico'
        if (Test-Path -LiteralPath $coordIconPath) {
            $Script:CoordIcon = New-Object System.Drawing.Icon($coordIconPath)
        }
    } catch { $Script:CoordIcon = $null }
}

# Carpeta de sortida: local\geocodificacio\ (la mateixa on viu la memoria cau
# dels portals). Dins del clone pero fora del repositori.
$CoordOutputDir = Get-LocalSubdir $RepoRoot 'Geocodificacio'

# ============================================================================
# FUNCIONS PURES (provables en mode headless, sense Office)
# ============================================================================

# Clau d'agrupacio per coordenada. Arrodonim a 2 decimals (centimetres): l'Excel
# porta les coordenades amb 2 decimals i comparar doubles "a pel" es una manera
# excel.lent de no trobar mai dos punts iguals.
function Get-ClauCoord([double]$x, [double]$y) {
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    return ([math]::Round($x, 2).ToString('F2', $inv) + '|' + [math]::Round($y, 2).ToString('F2', $inv))
}

# ----------------------------------------------------------------------------
# ZONES
# ----------------------------------------------------------------------------
# El repas de 700 activitats no es pot fer d'una tirada, aixi que el municipi es
# parteix en una GRAELLA de quadres de 400 m i es treballa zona a zona.
#
# L'ancoratge es CONSTANT i no surt de les dades. Si la graella s'ancores al
# minim de l'Excel, n'hi hauria prou que una activitat nova caigues mes a
# l'oest perque TOTES les zones es desplacessin i "la zona C6" volgues dir una
# altra cosa que la setmana passada. Amb un origen fix, un nom de zona sempre
# es el mateix rectangle.
#
# 400 m calculat amb la base del 2026-08-18: 40 zones amb activitats apilades,
# la mes gran de 40 i la mediana de 17 -- una tanda raonable.
$CoordZonaMetres = 400
$CoordZonaX0     = 421200      # cantonada SO de la graella (UTM 31N)
$CoordZonaY0     = 4577200

# Nom de la zona d'una coordenada: lletra per la FILA (de sud a nord) i numero
# per la COLUMNA (d'oest a est) -- 'C6', 'F2'... Les coordenades que cauen fora
# de la graella tambe en reben un (la graella no te limit): el que les descarta
# es Test-CoordPlausible, no aixo.
function Get-ZonaDeCoord([double]$x, [double]$y) {
    $col = [int][math]::Floor(($x - $CoordZonaX0) / $CoordZonaMetres)
    $fil = [int][math]::Floor(($y - $CoordZonaY0) / $CoordZonaMetres)
    return ((_CoordLletraFila $fil) + [string]($col + 1))
}

# Lletra de la fila. Passada la Z es continua amb AA, AB... (no hi arribarem
# mai amb Cornella, pero una funcio que retorna escombraries fora de rang es
# una trampa esperant algu).
function _CoordLletraFila([int]$fil) {
    if ($fil -lt 0) { return 'z' + [string](-$fil) }   # al sud de l'origen
    $n = $fil
    $s = ''
    do {
        $s = [string][char](65 + ($n % 26)) + $s
        $n = [int][math]::Floor($n / 26) - 1
    } while ($n -ge 0)
    return $s
}

# Una coordenada UTM 31N pot ser d'aquest mon? Es una comprovacio GENEROSA (tot
# el fus, no nomes Cornella): nomes ha de caçar el que es impossible.
#
# Cal perque la base en porta: el GIA 1009 (Quintana i Millas 9) te
# X=423,37 Y=4578,81 -- li falten tres xifres. Sense aixo es pinta al golf de
# Guinea i estira el mapa sencer, de manera que la resta de punts queden
# amuntegats en un pixel.
function Test-CoordPlausible([double]$x, [double]$y) {
    return ($x -ge 100000 -and $x -le 900000 -and $y -ge 4000000 -and $y -le 4900000)
}

# Agrupa els registres per zona i retorna, per cada una, el nom, quantes
# activitats hi ha i els carrers mes repetits (per poder-la reconeixer). PURA.
#
# Els noms dels carrers es calculen AQUI, de les dades: al codi no hi ha escrit
# cap nom de cap carrer de Cornella, aixi que aixo no es pot desfasar.
function Get-ZonesAmbActivitats($records) {
    $z = @{}
    foreach ($r in @($records)) {
        $x = [double]$r.UtmX; $y = [double]$r.UtmY
        if (-not (Test-CoordPlausible $x $y)) { continue }
        $nom = Get-ZonaDeCoord $x $y
        if (-not $z.ContainsKey($nom)) { $z[$nom] = New-Object System.Collections.ArrayList }
        [void]$z[$nom].Add($r)
    }
    $out = @()
    foreach ($nom in @($z.Keys)) {
        $regs = @($z[$nom])
        $out += [pscustomobject]@{
            Nom        = $nom
            Comptador  = $regs.Count
            Carrers    = (Get-CarrersDominants $regs)
            Registres  = $regs
        }
    }
    # De mes gran a mes petita: les que fan mes nosa, primer.
    return @($out | Sort-Object -Property @{Expression='Comptador';Descending=$true}, @{Expression='Nom'})
}

# Els carrers mes repetits d'un grup de registres, per posar nom a una zona.
function Get-CarrersDominants($records, [int]$quants = 2) {
    $c = @{}
    foreach ($r in @($records)) {
        $v = ([string]$r.Carrer).Trim()
        if ($v -eq '') { continue }
        if ($c.ContainsKey($v)) { $c[$v] = $c[$v] + 1 } else { $c[$v] = 1 }
    }
    if ($c.Count -eq 0) { return '' }
    $top = @($c.GetEnumerator() |
        Sort-Object -Property @{Expression='Value';Descending=$true}, @{Expression='Key'} |
        Select-Object -First $quants | ForEach-Object { $_.Key })
    return ($top -join ' / ')
}

# Els registres APILATS: els que comparteixen coordenada amb algun altre.
# Aquests son exactament els que fan nosa al mapa. Retorna un subconjunt de
# $records, conservant l'ordre d'entrada.
# Retorna un array PLA (sense la coma protectora) i s'ha de consumir amb @().
# Vegeu la nota de ConvertFrom-CatastroAdXml a Geocodificador.ps1: barrejar les
# dues convencions embolcalla l'array dues vegades.
function Get-RegistresApilats($records) {
    $arr = @($records)
    if ($arr.Count -eq 0) { return @() }
    $comptes = @{}
    foreach ($r in $arr) {
        $k = Get-ClauCoord ([double]$r.UtmX) ([double]$r.UtmY)
        if ($comptes.ContainsKey($k)) { $comptes[$k] = $comptes[$k] + 1 } else { $comptes[$k] = 1 }
    }
    $out = @()
    foreach ($r in $arr) {
        $k = Get-ClauCoord ([double]$r.UtmX) ([double]$r.UtmY)
        if ($comptes[$k] -gt 1) { $out += $r }
    }
    return @($out)
}

# Les referencies cadastrals de PARCEL.LA (14 car.) que caldra consultar per a
# un conjunt de registres, sense repetits i ordenades. Es la unitat de consulta
# al Cadastre: una crida per parcel.la, no per activitat.
function Get-RefcatsAConsultar($records) {
    $set = @{}
    foreach ($r in @($records)) {
        $rc = Get-RefcatParcel $r.Rc
        if ($rc -ne '') { $set[$rc] = $true }
    }
    return @(@($set.Keys) | Sort-Object)
}

# Combina un registre de l'Excel amb els portals de la seva parcel.la i en
# treu l'objecte que anira al mapa: les DUES coordenades (la de l'Excel i la
# de facana) en UTM i en lat/lon, mes d'on surt la verda.
#
# Si no s'ha trobat portal, la verda es COL.LOCA A SOBRE de la vermella i es
# marca 'cadastre': aixi l'usuari la pot arrossegar igualment on toqui.
function New-ItemCoordenades($record, $portals) {
    $x = [double]$record.UtmX
    $y = [double]$record.UtmY
    $coord = Resolve-CoordEstabliment $portals $record.Carrer $record.Numero $x $y
    $llExcel  = Convert-UtmToLatLon $x $y 31 $true
    $llFacana = Convert-UtmToLatLon ([double]$coord.X) ([double]$coord.Y) 31 $true
    return [pscustomobject]@{
        Id        = [string]$record.Id
        Zona      = (Get-ZonaDeCoord $x $y)
        Rc        = [string]$record.Rc
        Adreca    = [string]$record.Adreca
        Activitat = [string]$record.Activitat
        XExcel    = $x
        YExcel    = $y
        LatExcel  = $llExcel.Lat
        LonExcel  = $llExcel.Lon
        XFacana   = [double]$coord.X
        YFacana   = [double]$coord.Y
        LatFacana = $llFacana.Lat
        LonFacana = $llFacana.Lon
        Precisio  = [string]$coord.Precisio
    }
}

# Recompte per a la finestra de tria i per al resum final.
function Get-ResumPrecisio($items) {
    $r = [ordered]@{ facana = 0; 'facana-dubtosa' = 0; 'facana-aprox' = 0; cadastre = 0 }
    foreach ($it in @($items)) {
        $p = [string]$it.Precisio
        if ($r.Contains($p)) { $r[$p] = $r[$p] + 1 } else { $r[$p] = 1 }
    }
    return $r
}

# ============================================================================
# EL MAPA (HTML)
# ============================================================================

# Genera el document HTML del mapa de coordenades. $items es la sortida de
# New-ItemCoordenades. Retorna l'HTML com a cadena.
#
# Dins de l'HTML hi ha tres peces que val la pena tenir localitzades:
#   latLonToUtm31()  la projeccio DIRECTA (lat/lon -> UTM 31N). Cal perque el
#                    Leaflet ens dona graus quan s'arrossega un punt i nosaltres
#                    hem d'exportar metres. Es la inversa exacta de
#                    Convert-UtmToLatLon (Ruta.ps1); verificada d'anada i
#                    tornada sobre tot el terme municipal amb un error maxim de
#                    0,07 mm.
#   buildXlsx()      escriu un .xlsx de veritat sense cap biblioteca: un .xlsx
#                    es un ZIP amb cinc XML a dins, i amb el metode "sense
#                    compressio" nomes cal el CRC-32 i les capçaleres del ZIP.
#   desaCorreccions() els punts que mous a ma van al localStorage del navegador,
#                    amb clau del fitxer d'origen. Si tanques la pagina i la
#                    tornes a obrir, hi son.
function Build-CoordenadesHtml($items, [string]$dbLabel, [string]$abast, [string]$fontName, $portals) {
    $arr = @($items)
    $itemsJson = ConvertTo-Json @($arr | ForEach-Object {
        # [ordered]: sense aixo, ConvertTo-Json treu les propietats en un ordre
        # diferent a cada execucio i l'HTML generat canvia sense que hagin
        # canviat les dades (la mateixa trampa que ja hi havia a Ruta.ps1).
        [ordered]@{
            id        = [string]$_.Id
            zona      = [string]$_.Zona
            rc        = [string]$_.Rc
            adreca    = [string]$_.Adreca
            activitat = [string]$_.Activitat
            xe        = [double]$_.XExcel
            ye        = [double]$_.YExcel
            late      = [double]$_.LatExcel
            lone      = [double]$_.LonExcel
            xf        = [double]$_.XFacana
            yf        = [double]$_.YFacana
            latf      = [double]$_.LatFacana
            lonf      = [double]$_.LonFacana
            prec      = [string]$_.Precisio
        }
    }) -Depth 5 -Compress
    # El JSON ha de ser una LLISTA sempre. Aixo abans es decidia pel nombre
    # d'elements, donant per fet que ConvertTo-Json desembolcalla quan n'hi ha
    # un de sol -- i en el PowerShell de l'usuari NO ho fa: el mapa d'una sola
    # activitat sortia amb [[{...}]] i no arrencava. Ara es mira la SORTIDA, que
    # es el que compta, i tant se val com es comporti cada versio.
    if ([string]::IsNullOrWhiteSpace($itemsJson) -or $itemsJson -eq 'null') { $itemsJson = '[]' }
    elseif (-not $itemsJson.TrimStart().StartsWith('[')) { $itemsJson = "[$itemsJson]" }

    # Els portals de les parcel.les consultades, per pintar-los amb el seu
    # numero com al planol del Cadastre. Poden ser cap.
    $portalsJson = ConvertTo-Json @(@($portals) | Where-Object { $null -ne $_ } | ForEach-Object {
        [ordered]@{ n = [string]$_.Numero; v = [string]$_.Via; lat = [double]$_.Lat; lon = [double]$_.Lon }
    }) -Depth 5 -Compress
    if ([string]::IsNullOrWhiteSpace($portalsJson) -or $portalsJson -eq 'null') { $portalsJson = '[]' }
    elseif (-not $portalsJson.TrimStart().StartsWith('[')) { $portalsJson = "[$portalsJson]" }

    $resum = Get-ResumPrecisio $arr
    $today = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    $dbEnc = _HtmlEncode $dbLabel
    # El nom del fitxer d'origen va al JavaScript com a literal JSON (i no
    # HTML-escapat): es la clau amb que es desen les correccions al navegador,
    # i ha de ser el nom EXACTE perque les correccions d'una base es quedin
    # amb aquella base.
    $fontJson = ConvertTo-Json ([string]$fontName) -Compress
    $abastEnc = _HtmlEncode $abast
    $nTot = $arr.Count
    $nFac = [int]$resum['facana']
    $nDub = [int]$resum['facana-dubtosa']
    $nApr = [int]$resum['facana-aprox']
    $nCad = [int]$resum['cadastre']

    $html = @"
<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Coordenades dels establiments - Cornella</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
      integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=" crossorigin=""/>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; font-family: Segoe UI, Arial, sans-serif; color: #1a1a1a; }
  #top { background: #14365c; color: #fff; padding: 10px 16px; display: flex; align-items: baseline; gap: 14px; flex-wrap: wrap; }
  #top h1 { font-size: 18px; margin: 0; }
  #top .meta { font-size: 13px; opacity: .9; }
  #wrap { display: flex; height: calc(100vh - 116px); }
  #map { flex: 1 1 auto; }
  #side { width: 380px; overflow: auto; border-left: 1px solid #ddd; padding: 0 0 30px 0; }
  #side h2 { font-size: 14px; margin: 12px 14px 6px; }
  #cerca { width: calc(100% - 28px); margin: 0 14px 8px; padding: 6px 8px; font-size: 13px;
           border: 1px solid #ccd2da; border-radius: 4px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th, td { text-align: left; padding: 5px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
  th { background: #f3f5f8; position: sticky; top: 0; z-index: 2; }
  tbody tr { cursor: pointer; }
  tbody tr:hover { background: #f7fafd; }
  tr.moguda td.id { font-weight: bold; color: #a0560b; }
  tr.revisada { background: #f2f7f4; }
  tr.revisada td.id::before { content: '\2713 '; color: #14365c; font-weight: bold; }
  td.id { font-family: Consolas, monospace; color: #14365c; white-space: nowrap; }
  td.dist { text-align: right; white-space: nowrap; color: #555; }
  .pin { display: inline-block; width: 9px; height: 9px; border-radius: 50%; margin-right: 5px; vertical-align: -1px; }
  .pin-fac { background: #2ecc71; border: 1px solid #145a32; }
  .pin-dub { background: #f1c40f; border: 1px solid #7d6608; }
  .pin-apr { background: #a9dfbf; border: 1px solid #145a32; }
  .pin-cad { background: #fff;    border: 2px solid #145a32; width: 7px; height: 7px; }
  .pin-man { background: #e67e22; border: 1px solid #7e5109; }
  .pin-rev { background: #fff; border: 2px solid #14365c; }
  .pin-red { background: #c0392b; }
  #llegenda b { color: #14365c; }
  #llegenda { font-size: 12px; color: #444; margin: 10px 14px; line-height: 1.7; }
  #bar { padding: 8px 16px; border-top: 1px solid #ddd; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
  button { background: #14365c; color: #fff; border: 0; padding: 8px 16px; border-radius: 5px; cursor: pointer; font-size: 14px; }
  button:hover { background: #1d4d82; }
  button.sec { background: #fff; color: #14365c; border: 1px solid #c3ccd8; }
  button.sec:hover { background: #eef3f9; }
  #estat { font-size: 12px; color: #555; }
  /* El COLOR diu d'on surt el punt; el CONTORN GRUIXUT, que ja l'has repassat. */
  .marker-verd { width: 14px; height: 14px; border-radius: 50%; background: #2ecc71;
                 border: 2px solid #145a32; box-shadow: 0 1px 3px rgba(0,0,0,.45); cursor: move; }
  .marker-verd.dubtosa     { background: #f1c40f; border-color: #7d6608; }
  .marker-verd.aprox       { background: #a9dfbf; border-color: #145a32; }
  .marker-verd.sensefacana { background: #fff;    border-color: #145a32; }
  .marker-verd.moguda      { background: #e67e22; border-color: #7e5109; }
  .marker-verd.revisada    { box-shadow: 0 0 0 3px #14365c, 0 1px 3px rgba(0,0,0,.45); }
  /* Numeros de portal del Cadastre: nomes a partir del zoom de carrer. */
  .portal-num { white-space: nowrap; font-size: 11px; color: #34495e; font-weight: bold;
                text-shadow: -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff, 1px 1px 0 #fff;
                pointer-events: none; }
  .portal-num i { display: inline-block; width: 5px; height: 5px; border-radius: 50%;
                  background: #34495e; margin-right: 3px; vertical-align: 1px; }
  .pop h3 { margin: 0 0 4px; font-size: 14px; color: #14365c; }
  .pop .rc { font-family: Consolas, monospace; font-size: 11px; color: #666; }
  .pop table { font-size: 12px; margin-top: 6px; }
  .pop td { border: 0; padding: 1px 6px 1px 0; }
  .pop .org { margin-top: 5px; font-size: 12px; }
</style>
</head>
<body>
<div id="top">
  <h1>Coordenades dels establiments &mdash; Cornella de Llobregat</h1>
  <span class="meta">$nTot activitats &middot; $abastEnc</span>
  <span class="meta">Portal: $nFac &middot; dubtosos: $nDub &middot; aprox.: $nApr &middot; sense portal: $nCad</span>
  <span class="meta">Generat: $today</span>
</div>
<div id="wrap">
  <div id="map"></div>
  <div id="side">
    <h2>Activitats</h2>
    <input id="cerca" type="search" placeholder="Filtra per ID o adreca...">
    <table>
      <thead><tr><th>ID GIA</th><th>Adreca</th><th>Desplac.</th></tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div id="llegenda">
      <span class="pin pin-red"></span> coordenada actual de l'Excel (fixa)<br>
      <span class="pin pin-fac"></span> coordenada de facana del Cadastre<br>
      <span class="pin pin-cad"></span> sense facana: comenca sobre la vermella<br>
      <span class="pin pin-man"></span> moguda per tu<br>
      Arrossega els punts verds per posar-los on toca.
    </div>
  </div>
</div>
<div id="bar">
  <button onclick="baixaExcel()">Baixar Excel (.xlsx)</button>
  <button class="sec" onclick="validaVisibles()">Validar tot el que es veu</button>
  <button class="sec" onclick="esborraCorreccions()">Esborrar el meu repàs</button>
  <label style="font-size:12px;color:#333;"><input type="checkbox" id="chkNums" checked> números dels portals</label>
  <span id="estat"></span>
  <span style="font-size:12px;color:#777;">Base de dades: $dbEnc</span>
</div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
        integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=" crossorigin=""></script>
<script>
var ITEMS   = $itemsJson;
var PORTALS = $portalsJson;
var FONT    = $fontJson;
var CLAU  = 'coordenades:' + FONT;
</script>
<script>
// ---------------------------------------------------------------------------
// PROJECCIO: lat/lon (WGS84) -> UTM fus 31N (ETRS89, EPSG:25831).
// El Leaflet ens dona graus quan s'arrossega un punt, i nosaltres hem
// d'exportar metres. Es la inversa exacta de Convert-UtmToLatLon (Ruta.ps1);
// comprovada d'anada i tornada sobre tot el terme municipal, error < 0,1 mm.
// ---------------------------------------------------------------------------
function latLonToUtm31(lat, lon) {
  var a = 6378137.0, f = 1.0 / 298.257223563, k0 = 0.9996;
  var e2 = f * (2 - f), ep2 = e2 / (1 - e2);
  var rad = Math.PI / 180;
  var phi = lat * rad, lam = lon * rad;
  var lam0 = (31 * 6 - 183) * rad;
  var sinP = Math.sin(phi), cosP = Math.cos(phi), tanP = Math.tan(phi);
  var N = a / Math.sqrt(1 - e2 * sinP * sinP);
  var T = tanP * tanP;
  var C = ep2 * cosP * cosP;
  var A = (lam - lam0) * cosP;
  var M = a * ((1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256) * phi
    - (3 * e2 / 8 + 3 * e2 * e2 / 32 + 45 * e2 * e2 * e2 / 1024) * Math.sin(2 * phi)
    + (15 * e2 * e2 / 256 + 45 * e2 * e2 * e2 / 1024) * Math.sin(4 * phi)
    - (35 * e2 * e2 * e2 / 3072) * Math.sin(6 * phi));
  var x = k0 * N * (A + (1 - T + C) * Math.pow(A, 3) / 6
    + (5 - 18 * T + T * T + 72 * C - 58 * ep2) * Math.pow(A, 5) / 120) + 500000.0;
  var y = k0 * (M + N * tanP * (A * A / 2
    + (5 - T + 9 * C + 4 * C * C) * Math.pow(A, 4) / 24
    + (61 - 58 * T + T * T + 600 * C - 330 * ep2) * Math.pow(A, 6) / 720));
  return [x, y];
}

function distanciaUtm(x1, y1, x2, y2) {
  var dx = x1 - x2, dy = y1 - y2;
  return Math.sqrt(dx * dx + dy * dy);
}

function escHtml(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ---------------------------------------------------------------------------
// ESCRIPTOR DE .xlsx (sense cap biblioteca).
// Un .xlsx es un ZIP amb cinc XML a dins. Amb el metode "sense compressio"
// nomes cal el CRC-32 i les capçaleres del ZIP, i surt un fitxer que l'Excel
// obre amb doble clic i sense cap avis de format.
// ---------------------------------------------------------------------------
var CRC_TABLE = (function () {
  var t = new Uint32Array(256);
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) { c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); }
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  var c = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) { c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function utf8(str) { return new TextEncoder().encode(str); }

function escXml(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    // Els caracters de control no son legals en XML 1.0 i farien il-legible el
    // fitxer. No n'hi hauria d'haver, pero les dades venen del GIA.
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
}

// Nom de columna d'Excel: 1 -> A, 27 -> AA.
function colName(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ZIP amb metode 0 (sense comprimir). Retorna Uint8Array.
function zipStore(files) {
  var chunks = [], central = [], offset = 0;
  function u16(v) { return [v & 0xFF, (v >>> 8) & 0xFF]; }
  function u32(v) { return [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF]; }
  for (var i = 0; i < files.length; i++) {
    var name = utf8(files[i].name);
    var data = files[i].data;
    var crc = crc32(data);
    // Data/hora fixes (1980-01-01): el mateix contingut dona sempre el mateix
    // fitxer, cosa que fa que es pugui comparar byte a byte a les proves.
    var lfh = [].concat(
      u32(0x04034B50), u16(20), u16(0x0800), u16(0), u16(0), u16(33),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0));
    chunks.push(new Uint8Array(lfh), name, data);
    central.push({ name: name, crc: crc, size: data.length, offset: offset });
    offset += lfh.length + name.length + data.length;
  }
  var cdChunks = [], cdSize = 0;
  for (var j = 0; j < central.length; j++) {
    var e = central[j];
    var cdh = [].concat(
      u32(0x02014B50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(33),
      u32(e.crc), u32(e.size), u32(e.size), u16(e.name.length),
      u16(0), u16(0), u16(0), u16(0), u32(0), u32(e.offset));
    cdChunks.push(new Uint8Array(cdh), e.name);
    cdSize += cdh.length + e.name.length;
  }
  cdChunks.push(new Uint8Array([].concat(
    u32(0x06054B50), u16(0), u16(0), u16(central.length), u16(central.length),
    u32(cdSize), u32(offset), u16(0))));
  var all = chunks.concat(cdChunks), total = 0;
  for (var k = 0; k < all.length; k++) { total += all[k].length; }
  var out = new Uint8Array(total), p = 0;
  for (var m = 0; m < all.length; m++) { out.set(all[m], p); p += all[m].length; }
  return out;
}

// header: array de textos. rows: array d'arrays; els numbers van com a numero
// i la resta com a text inline (aixi no cal sharedStrings.xml).
function buildXlsx(sheetName, header, rows) {
  function cellsOf(vals, rowNum) {
    var s = '';
    for (var c = 0; c < vals.length; c++) {
      var ref = colName(c + 1) + rowNum, v = vals[c];
      if (typeof v === 'number' && isFinite(v)) {
        s += '<c r="' + ref + '"><v>' + v + '</v></c>';
      } else if (v !== null && v !== undefined && v !== '') {
        s += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + escXml(v) + '</t></is></c>';
      }
    }
    return '<row r="' + rowNum + '">' + s + '</row>';
  }
  var sd = cellsOf(header, 1);
  for (var r = 0; r < rows.length; r++) { sd += cellsOf(rows[r], r + 2); }

  // Amplades de columna, perque el fitxer s'obri llegible i no s'hagin
  // d'eixamplar a ma cada vegada.
  var widths = [10, 24, 40, 15, 15, 15, 15, 26, 14];
  var cols = '<cols>';
  for (var w = 0; w < header.length; w++) {
    cols += '<col min="' + (w + 1) + '" max="' + (w + 1) + '" width="' + (widths[w] || 14) + '" customWidth="1"/>';
  }
  cols += '</cols>';

  var decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  var parts = [
    { name: '[Content_Types].xml', text: decl +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>' },
    { name: '_rels/.rels', text: decl +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>' },
    { name: 'xl/workbook.xml', text: decl +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + escXml(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', text: decl +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>' },
    { name: 'xl/styles.xml', text: decl +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>' },
    { name: 'xl/worksheets/sheet1.xml', text: decl +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      cols + '<sheetData>' + sd + '</sheetData></worksheet>' }
  ];
  return zipStore(parts.map(function (p) { return { name: p.name, data: utf8(p.text) }; }));
}

// ---------------------------------------------------------------------------
// ESTAT DE CADA ACTIVITAT
//   origen:   d'on surt el punt verd -- 'facana' | 'facana-dubtosa' |
//             'facana-aprox' | 'cadastre' | 'manual' (l'has mogut tu)
//   revisada: ja l'has donada per bona
//
// estat[] va per POSICIO dins d'ITEMS, no per ID: si algun dia la base portes
// dos cops el mateix ID Activitat, dues fitxes es trepitjarien. Al navegador,
// en canvi, es desa per ID -- que es el que ha de sobreviure quan es torni a
// generar el mapa d'una altra zona.
// ---------------------------------------------------------------------------
var estat = [];

function llegeixDesat() {
  try {
    var cru = window.localStorage.getItem(CLAU);
    return cru ? (JSON.parse(cru) || {}) : {};
  } catch (e) { return {}; }   // localStorage desactivat: no es cap drama
}

function escriuDesat(d) {
  try { window.localStorage.setItem(CLAU, JSON.stringify(d)); } catch (e) { }
}

function r2(v) { return Math.round(v * 100) / 100; }

function carregaCorreccions() {
  var desat = llegeixDesat();
  for (var i = 0; i < ITEMS.length; i++) {
    var it = ITEMS[i];
    var d = desat[it.id];
    if (d && isFinite(d.lat) && isFinite(d.lon)) {
      // Les versions velles nomes desaven lat/lon: es donen per mogudes a ma.
      estat[i] = { lat: d.lat, lon: d.lon, origen: d.origen || 'manual', revisada: true };
    } else {
      estat[i] = { lat: it.latf, lon: it.lonf, origen: it.prec, revisada: false };
    }
  }
}

// Desa (o treu) una activitat del navegador. Es desa la FILA SENCERA i no
// nomes la posicio, perque l'Excel ha de poder portar tot el que has repassat
// d'aquesta base -- tambe el de les zones que avui no tens obertes.
function desaItem(i) {
  var it = ITEMS[i], e = estat[i];
  var desat = llegeixDesat();
  if (e.revisada) {
    var u = utmActual(i);
    desat[it.id] = {
      lat: e.lat, lon: e.lon, origen: e.origen,
      x: r2(u[0]), y: r2(u[1]), xe: it.xe, ye: it.ye,
      rc: it.rc, adreca: it.adreca, zona: it.zona
    };
  } else {
    delete desat[it.id];
  }
  escriuDesat(desat);
}

function comptaRevisades() {
  var n = 0;
  for (var i = 0; i < estat.length; i++) { if (estat[i].revisada) n++; }
  return n;
}

function textOrigen(o) {
  if (o === 'facana')         return 'portal amb el número exacte';
  if (o === 'facana-dubtosa') return 'portal DUBTÓS: n\'hi havia més d\'un amb aquell número';
  if (o === 'facana-aprox')   return 'aquell número no hi era: el portal més proper de la parcel·la';
  if (o === 'manual')         return 'mogut per tu';
  return 'sense portal: es queda al centre de la parcel·la';
}

// Coordenada UTM actual del punt verd. Si no s'ha mogut a ma, tornem els metres
// TAL COM van arribar (del Cadastre o de l'Excel) en lloc de reprojectar-los:
// aixi no s'hi acumula l'error d'anar i tornar de graus.
function utmActual(i) {
  var it = ITEMS[i], e = estat[i];
  if (e.origen === 'manual') { return latLonToUtm31(e.lat, e.lon); }
  return [it.xf, it.yf];
}

// ---------------------------------------------------------------------------
// MAPA
// ---------------------------------------------------------------------------
// preferCanvas: amb centenars d'activitats, dibuixar els cercles i les linies
// al canvas en lloc de fer-ne SVG es la diferencia entre un mapa fluid i un
// mapa que va a batzegades.
var map = L.map('map', { preferCanvas: true, zoomSnap: 0.25, zoomDelta: 0.5, wheelPxPerZoomLevel: 120 });
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  maxZoom: 19, attribution: '&copy; OpenStreetMap'
}).addTo(map);

var capes = [];   // per posicio: { verd, vermell, linia, fila }
var bounds = [];

function classeVerd(i) {
  var e = estat[i];
  var c = 'marker-verd';
  if (e.origen === 'manual')             { c += ' moguda'; }
  else if (e.origen === 'facana-dubtosa'){ c += ' dubtosa'; }
  else if (e.origen === 'facana-aprox')  { c += ' aprox'; }
  else if (e.origen === 'cadastre')      { c += ' sensefacana'; }
  if (e.revisada) { c += ' revisada'; }
  return c;
}

function iconaVerd(i) {
  return L.divIcon({ className: '', html: '<div class="' + classeVerd(i) + '"></div>',
                     iconSize: [14, 14], iconAnchor: [7, 7] });
}

function popupHtml(i) {
  var it = ITEMS[i], e = estat[i];
  var utm = utmActual(i);
  var d = distanciaUtm(it.xe, it.ye, utm[0], utm[1]);
  return '<div class="pop">' +
    '<h3>ID ' + escHtml(it.id) + (it.zona ? ' &middot; zona ' + escHtml(it.zona) : '') + '</h3>' +
    '<div>' + escHtml(it.activitat || '(activitat no especificada)') + '</div>' +
    '<div>' + escHtml(it.adreca || '(sense adreça)') + '</div>' +
    '<div class="rc">' + escHtml(it.rc) + '</div>' +
    '<table>' +
    '<tr><td><span class="pin pin-red"></span>Excel</td><td>' + it.xe.toFixed(2) + '</td><td>' + it.ye.toFixed(2) + '</td></tr>' +
    '<tr><td>Nova</td><td>' + utm[0].toFixed(2) + '</td><td>' + utm[1].toFixed(2) + '</td></tr>' +
    '</table>' +
    '<div class="org">' + escHtml(textOrigen(e.origen)) + '<br>' +
    'Desplaçament: ' + d.toFixed(1) + ' m<br>' +
    (e.revisada ? '<b>Validat.</b> Torna-hi a fer clic per desfer-ho.'
                : 'Fes clic al punt per validar-lo, o arrossega\'l.') +
    '</div></div>';
}

function refrescaItem(i) {
  var it = ITEMS[i], c = capes[i], e = estat[i];
  var utm = utmActual(i);
  var d = distanciaUtm(it.xe, it.ye, utm[0], utm[1]);
  var html = popupHtml(i);
  c.linia.setLatLngs([[it.late, it.lone], [e.lat, e.lon]]);
  c.verd.setIcon(iconaVerd(i));
  c.verd.setPopupContent(html);
  c.vermell.setPopupContent(html);
  c.fila.className = (e.revisada ? 'revisada' : '') + (e.origen === 'manual' ? ' moguda' : '');
  c.fila.cells[2].textContent = d.toFixed(1) + ' m';
}

function actualitzaEstatBarra() {
  var n = comptaRevisades();
  var total = Object.keys(llegeixDesat()).length;
  document.getElementById('estat').textContent =
    'Validades ' + n + ' de ' + ITEMS.length + ' en aquest mapa  ·  ' +
    total + ' en total en aquesta base de dades';
}

// Validar / desfer. Desfer un punt que havies mogut el torna on el Cadastre
// deia, que es l'unica manera de fer marxa enrere sense refer el mapa.
function commutaValidada(i) {
  var it = ITEMS[i], e = estat[i];
  if (e.revisada) {
    estat[i] = { lat: it.latf, lon: it.lonf, origen: it.prec, revisada: false };
    capes[i].verd.setLatLng([it.latf, it.lonf]);
  } else {
    e.revisada = true;
  }
  refrescaItem(i);
  desaItem(i);
  actualitzaEstatBarra();
}

function validaVisibles() {
  var n = 0;
  for (var i = 0; i < ITEMS.length; i++) {
    if (capes[i].fila.style.display === 'none') { continue; }
    if (estat[i].revisada) { continue; }
    estat[i].revisada = true;
    refrescaItem(i);
    desaItem(i);
    n++;
  }
  actualitzaEstatBarra();
  alert(n === 0 ? 'Ja les tenies totes validades.' : 'Validades ' + n + ' activitats.');
}

var tbody = document.getElementById('tbody');

function pinta() {
  carregaCorreccions();
  for (var i = 0; i < ITEMS.length; i++) {
    (function (idx) {
      var it = ITEMS[idx], e = estat[idx];
      var html = popupHtml(idx);

      // Línia fina entre la coordenada de l'Excel i la nova, perquè es vegi
      // quin verd correspon a quin vermell.
      var linia = L.polyline([[it.late, it.lone], [e.lat, e.lon]],
        { color: '#888', weight: 1, opacity: .8, dashArray: '3,4', interactive: false }).addTo(map);

      // Vermell: la coordenada que hi ha ara a l'Excel. NO es pot moure.
      var vermell = L.circleMarker([it.late, it.lone],
        { radius: 5, color: '#8e2b21', weight: 1, fillColor: '#c0392b', fillOpacity: .95 })
        .addTo(map).bindPopup(html);

      // Verd: la coordenada de façana. Aquest sí que es pot arrossegar, i per
      // això ha de ser un L.marker (els circleMarker no són arrossegables).
      var verd = L.marker([e.lat, e.lon], { draggable: true, icon: iconaVerd(idx) })
        .addTo(map).bindPopup(html);

      verd.on('dragend', function () {
        var p = verd.getLatLng();
        estat[idx] = { lat: p.lat, lon: p.lng, origen: 'manual', revisada: true };
        refrescaItem(idx);
        desaItem(idx);
        actualitzaEstatBarra();
      });
      // Un clic al punt el valida (o desfà la validació). Leaflet dispara
      // 'click' després d'un arrossegament curt, així que ens assegurem que no
      // ve d'un drag mirant si la posició ha canviat.
      verd.on('click', function () { commutaValidada(idx); });

      var fila = document.createElement('tr');
      var utm = utmActual(idx);
      fila.innerHTML =
        '<td class="id">' + escHtml(it.id) + '</td>' +
        '<td>' + escHtml(it.adreca || '—') + '</td>' +
        '<td class="dist">' + distanciaUtm(it.xe, it.ye, utm[0], utm[1]).toFixed(1) + ' m</td>';
      fila.className = (e.revisada ? 'revisada' : '') + (e.origen === 'manual' ? ' moguda' : '');
      fila.addEventListener('click', function () {
        map.setView([estat[idx].lat, estat[idx].lon], 19);
        verd.openPopup();
      });
      tbody.appendChild(fila);

      capes[idx] = { verd: verd, vermell: vermell, linia: linia, fila: fila };
      bounds.push([it.late, it.lone]);
      bounds.push([e.lat, e.lon]);
    })(i);
  }

  muntaPortals();

  if (bounds.length > 0) { map.fitBounds(L.latLngBounds(bounds).pad(0.08)); }
  else { map.setView([41.355, 2.073], 14); }
  actualitzaEstatBarra();
}

// ---------------------------------------------------------------------------
// ELS NUMEROS DELS PORTALS (com al planol del Cadastre)
// ---------------------------------------------------------------------------
// Son TOTS els portals de les parcel.les consultades, tambe els que no tenen
// cap activitat: sense ells no hi ha manera de dir si un punt esta ben posat.
// Nomes surten a partir del zoom de carrer, perque de lluny una illa amb vint
// numeros tapa el mapa.
var ZOOM_NUMS = 18;
var capaPortals = L.layerGroup();

function muntaPortals() {
  for (var i = 0; i < PORTALS.length; i++) {
    var p = PORTALS[i];
    var etiqueta = (p.n || '?') + (p.v ? '' : '');
    capaPortals.addLayer(L.marker([p.lat, p.lon], {
      interactive: false, keyboard: false,
      icon: L.divIcon({ className: '', iconSize: [0, 0], iconAnchor: [0, 6],
                        html: '<div class="portal-num"><i></i>' + escHtml(etiqueta) + '</div>' })
    }));
  }
  map.on('zoomend', refrescaPortals);
  document.getElementById('chkNums').addEventListener('change', refrescaPortals);
  refrescaPortals();
}

function refrescaPortals() {
  var vol = document.getElementById('chkNums').checked && map.getZoom() >= ZOOM_NUMS;
  if (vol && !map.hasLayer(capaPortals)) { map.addLayer(capaPortals); }
  else if (!vol && map.hasLayer(capaPortals)) { map.removeLayer(capaPortals); }
}

// Filtre del panell lateral: per ID o per adreça.
document.getElementById('cerca').addEventListener('input', function (ev) {
  var q = ev.target.value.trim().toLowerCase();
  for (var i = 0; i < ITEMS.length; i++) {
    var it = ITEMS[i];
    var visible = q === '' ||
      String(it.id).toLowerCase().indexOf(q) >= 0 ||
      String(it.adreca || '').toLowerCase().indexOf(q) >= 0 ||
      String(it.zona || '').toLowerCase() === q;
    capes[i].fila.style.display = visible ? '' : 'none';
  }
});

// ---------------------------------------------------------------------------
// BAIXAR L'EXCEL
// ---------------------------------------------------------------------------
// Hi surt TOT el que hagis validat d'aquesta base de dades, encara que sigui
// d'una altra zona i d'un altre dia: la idea es acabar amb UN sol fitxer.
// Quan actualitzis l'Excel d'activitats i en generis un de nou, aixo es buida
// sol, perque la clau del navegador es el nom del fitxer d'origen.
function baixaExcel() {
  var desat = llegeixDesat();
  var ids = Object.keys(desat);
  if (ids.length === 0) {
    alert('Encara no has validat cap activitat.\n\nAmplia el mapa, comprova els punts i fes-hi clic per validar-los (o arrossega\'ls). Després torna a provar.');
    return;
  }
  var header = ['ID GIA', 'Ref. cadastral', 'Adreça', 'Zona',
                'UTM X (Excel)', 'UTM Y (Excel)', 'UTM X (nova)', 'UTM Y (nova)',
                'Origen', 'Desplaçament (m)'];
  var rows = [];
  for (var k = 0; k < ids.length; k++) {
    var id = ids[k], d = desat[id];
    var xe = isFinite(d.xe) ? d.xe : '';
    var ye = isFinite(d.ye) ? d.ye : '';
    var x = isFinite(d.x) ? d.x : '';
    var y = isFinite(d.y) ? d.y : '';
    var desp = '';
    if (isFinite(d.xe) && isFinite(d.x)) { desp = Math.round(distanciaUtm(d.xe, d.ye, d.x, d.y) * 10) / 10; }
    rows.push([String(id), String(d.rc || ''), String(d.adreca || ''), String(d.zona || ''),
               xe, ye, x, y, textOrigen(d.origen || 'manual'), desp]);
  }
  rows.sort(function (a, b) { return (a[3] + '').localeCompare(b[3] + '') || (a[0] - b[0]); });

  var bytes = buildXlsx('Coordenades', header, rows);
  var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  var ara = new Date();
  function dosX(n) { return (n < 10 ? '0' : '') + n; }
  a.href = url;
  a.download = 'Coordenades_' + ara.getFullYear() + dosX(ara.getMonth() + 1) + dosX(ara.getDate()) +
               '_' + dosX(ara.getHours()) + dosX(ara.getMinutes()) + '.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
}

function esborraCorreccions() {
  var total = Object.keys(llegeixDesat()).length;
  if (total === 0) { alert('Encara no has validat res.'); return; }
  if (!confirm('Segur que vols descartar el repàs sencer d\'aquesta base de dades?\n\nSón ' + total + ' activitats validades, també les d\'altres zones i altres dies.')) { return; }
  try { window.localStorage.removeItem(CLAU); } catch (e) { }
  for (var i = 0; i < ITEMS.length; i++) {
    var it = ITEMS[i];
    estat[i] = { lat: it.latf, lon: it.lonf, origen: it.prec, revisada: false };
    capes[i].verd.setLatLng([it.latf, it.lonf]);
    refrescaItem(i);
  }
  actualitzaEstatBarra();
}

pinta();
</script>
</body>
</html>
"@
    return $html
}

# ============================================================================
# LECTURA D'EXCEL (COM) - nomes a Windows amb Excel; no es prova en headless.
# ============================================================================

# Llegeix la fulla "Estes" i retorna un registre per activitat amb tot el que
# necessitem: { Id; Rc; Adreca; Carrer; Numero; Activitat; UtmX; UtmY }.
# Les activitats SENSE coordenades s'ometen (no es poden situar al mapa) i es
# compten a part.
function Read-CoordenadesFromExcel($excelFile) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            $sh = _RutaFindEstesSheet $wb
            if ($null -eq $sh) { throw "No s'ha trobat la fulla 'Estes'/'Estes' al fitxer Excel." }
            $data = $sh.UsedRange.Value2
            if ($null -eq $data) { return [pscustomobject]@{ Registres = @(); SenseCoord = 0; Impossibles = @() } }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

            # Capcalera (fila 1) -> array 0-based de noms de columna.
            $headers = @()
            for ($c = 1; $c -le $cols; $c++) {
                $hv = $data[1, $c]
                $headers += $(if ($null -eq $hv) { '' } else { ([string]$hv).Trim() })
            }

            # Columnes per NOM (mes robust que per index: si el GIA n'afegeix
            # una al mig, res no es trenca). Els noms de cerca s'escriuen en
            # ASCII SENSE accents: Find-HeaderColumn normalitza sense
            # diacritics, aixi que 'Emp. Numero' encaixa amb 'Emp. Numero' real.
            $colId   = Find-HeaderColumn $headers 'ID Activitat'
            $colRc   = Find-HeaderColumn $headers 'Ref. cadastral'
            $colUtmX = Find-HeaderColumn $headers 'UTM X'
            $colUtmY = Find-HeaderColumn $headers 'UTM Y'
            $colVia  = Find-HeaderColumn $headers 'Emp. Tipus via'
            $colCarr = Find-HeaderColumn $headers 'Emp. Carrer'
            $colNum  = Find-HeaderColumn $headers 'Emp. Numero'
            $colLlet = Find-HeaderColumn $headers 'Emp. Lletra'
            $colAct  = Find-HeaderColumn $headers 'Activitat principal'
            $colNom  = Find-HeaderColumn $headers 'Nom comercial activitat'

            if ($colUtmX -lt 1 -or $colUtmY -lt 1) {
                throw "La fulla 'Estes' no te les columnes 'UTM X' i 'UTM Y'."
            }

            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }

            $registres = @()
            $senseCoord = 0
            $impossibles = @()
            for ($r = 2; $r -le $rows; $r++) {
                # ID Activitat (numero -> enter sense decimals, com a Ruta).
                $idCell = if ($colId -ge 1 -and $colId -le $cols) { $data[$r, $colId] } else { $null }
                $id = if ($idCell -is [double]) {
                    if ([math]::Floor($idCell) -eq $idCell) { [string][int]$idCell } else { [string]$idCell }
                } elseif ($null -ne $idCell) { ([string]$idCell).Trim() } else { '' }
                if ($id -eq '') { continue }

                $carrerRaw = & $get $r $colCarr
                $numeroRaw = & $get $r $colNum

                $x = ConvertTo-UtmNumber (& $get $r $colUtmX)
                $y = ConvertTo-UtmNumber (& $get $r $colUtmY)
                if ($null -eq $x -or $null -eq $y -or $x -eq 0 -or $y -eq 0) { $senseCoord++; continue }
                # Coordenades que no poden ser d'aquest mon (al GIA n'hi ha: una
                # activitat amb X=423,37, que son les xifres bones dividides per
                # mil). Si es colessin, el mapa s'estiraria fins a l'Atlantic i
                # la resta de punts quedarien tots en un pixel.
                if (-not (Test-CoordPlausible $x $y)) {
                    $impossibles += [pscustomobject]@{ Id = $id; X = $x; Y = $y; Adreca = (Format-EmpAddress (& $get $r $colVia) $carrerRaw $numeroRaw '') }
                    continue
                }

                $nomCom = & $get $r $colNom
                $actPri = & $get $r $colAct
                $activitat = if ($nomCom -ne '' -and $actPri -ne '') { "$nomCom - $actPri" }
                             elseif ($nomCom -ne '') { $nomCom }
                             else { $actPri }

                $registres += [pscustomobject]@{
                    Id        = $id
                    Rc        = (& $get $r $colRc)
                    Adreca    = (Format-EmpAddress (& $get $r $colVia) $carrerRaw $numeroRaw (& $get $r $colLlet))
                    Carrer    = $carrerRaw
                    Numero    = $numeroRaw
                    Activitat = $activitat
                    UtmX      = $x
                    UtmY      = $y
                }
            }
            return [pscustomobject]@{ Registres = @($registres); SenseCoord = $senseCoord; Impossibles = @($impossibles) }
        } finally {
            $wb.Close($false)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# ============================================================================
# INTERFICIE (WinForms) - nomes en us normal.
# ============================================================================

function Show-CoordInfo([string]$msg, [string]$title = 'Coordenades', [string]$icon = 'Information') {
    [System.Windows.Forms.MessageBox]::Show($msg, $title, 'OK', $icon) | Out-Null
}

# Finestra de tria: quines ZONES es repassen.
#
# Retorna { NomsZones; NomesApilades } o $null si es cancel.la.
#
# NOTA sobre "quantes en portes de repassades": aqui NO es pot saber. El que has
# validat viu al localStorage del NAVEGADOR, i el PowerShell no hi te acces. El
# progres, per tant, el mostra el mapa (que si que hi te acces) i aquesta
# finestra nomes diu quantes activitats hi ha a cada zona.
function Show-CoordenadesForm([string]$dbLabel, $zonesApilades, $zonesTotes,
                              [int]$nSenseCoord, [int]$nImpossibles) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Coordenades dels establiments'
    $form.Size = New-Object System.Drawing.Size(600, 610)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $true; $form.MaximizeBox = $false
    if ($null -ne $Script:CoordIcon) { $form.Icon = $Script:CoordIcon }

    $lblDb = New-Object System.Windows.Forms.Label
    $lblDb.Text = $dbLabel
    $lblDb.AutoSize = $false
    $lblDb.Size = New-Object System.Drawing.Size(550, 20)
    $lblDb.Location = New-Object System.Drawing.Point(15, 12)
    $lblDb.ForeColor = [System.Drawing.Color]::FromArgb(20, 54, 92)
    $form.Controls.Add($lblDb)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "El Cadastre situa cada activitat al centre de la seva PARCEL·LA, no al local, " +
                "i per això totes les d'un mateix edifici cauen al mateix punt.`r`n" +
                "El mapa et mostra les dues coordenades: la de l'Excel (vermell) i la del " +
                "PORTAL segons l'adreça (verd), que pots arrossegar."
    $lbl.AutoSize = $false
    $lbl.Size = New-Object System.Drawing.Size(550, 56)
    $lbl.Location = New-Object System.Drawing.Point(15, 36)
    $form.Controls.Add($lbl)

    $chkApil = New-Object System.Windows.Forms.CheckBox
    $chkApil.Text = 'Només les APILADES (les que comparteixen punt amb una altra)'
    $chkApil.AutoSize = $false
    $chkApil.Size = New-Object System.Drawing.Size(550, 22)
    $chkApil.Location = New-Object System.Drawing.Point(15, 96)
    $chkApil.Checked = $true
    $form.Controls.Add($chkApil)

    $lblZ = New-Object System.Windows.Forms.Label
    $lblZ.Text = 'Tria les zones que vols repassar ara (el municipi va per quadres de 400 m):'
    $lblZ.AutoSize = $false
    $lblZ.Size = New-Object System.Drawing.Size(550, 20)
    $lblZ.Location = New-Object System.Drawing.Point(15, 122)
    $form.Controls.Add($lblZ)

    $llista = New-Object System.Windows.Forms.CheckedListBox
    $llista.Size = New-Object System.Drawing.Size(550, 300)
    $llista.Location = New-Object System.Drawing.Point(15, 144)
    $llista.CheckOnClick = $true
    $llista.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($llista)

    $lblTotal = New-Object System.Windows.Forms.Label
    $lblTotal.AutoSize = $false
    $lblTotal.Size = New-Object System.Drawing.Size(320, 20)
    $lblTotal.Location = New-Object System.Drawing.Point(15, 452)
    $lblTotal.ForeColor = [System.Drawing.Color]::FromArgb(20, 54, 92)
    $form.Controls.Add($lblTotal)

    # $estat es un hashtable a proposit: les closures capturen els VALORS, i una
    # variable normal reassignada aqui dins no arribaria als handlers (la trampa
    # del .GetNewClosure() que ja ha costat una ronda en aquest projecte).
    $estat = @{ Zones = @() }

    $refrescaTotal = {
        $n = 0
        foreach ($i in $llista.CheckedIndices) { $n += [int]$estat.Zones[$i].Comptador }
        $lblTotal.Text = "$($llista.CheckedIndices.Count) zones triades  ·  $n activitats"
    }.GetNewClosure()

    $omple = {
        $zones = if ($chkApil.Checked) { @($zonesApilades) } else { @($zonesTotes) }
        $estat.Zones = $zones
        $llista.BeginUpdate()
        $llista.Items.Clear()
        foreach ($z in $zones) {
            [void]$llista.Items.Add(("{0,-4} {1,4} act.  {2}" -f $z.Nom, $z.Comptador, $z.Carrers))
        }
        $llista.EndUpdate()
        & $refrescaTotal
    }.GetNewClosure()

    $llista.add_ItemCheck({
        param($sender, $e)
        # ItemCheck salta ABANS que l'estat canvii: el total es calcula amb el
        # valor NOU d'aquest item i els que ja estaven marcats.
        $n = 0
        foreach ($i in $llista.CheckedIndices) { if ($i -ne $e.Index) { $n += [int]$estat.Zones[$i].Comptador } }
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { $n += [int]$estat.Zones[$e.Index].Comptador }
        $marcades = $llista.CheckedIndices.Count
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { $marcades++ } else { $marcades-- }
        $lblTotal.Text = "$marcades zones triades  ·  $n activitats"
    }.GetNewClosure())

    $chkApil.add_CheckedChanged($omple)

    $btnTot = New-Object System.Windows.Forms.Button
    $btnTot.Text = 'Marcar-ho tot'
    $btnTot.Size = New-Object System.Drawing.Size(120, 26)
    $btnTot.Location = New-Object System.Drawing.Point(340, 449)
    $btnTot.add_Click({
        for ($i = 0; $i -lt $llista.Items.Count; $i++) { $llista.SetItemChecked($i, $true) }
        & $refrescaTotal
    }.GetNewClosure())
    $form.Controls.Add($btnTot)

    $btnCap = New-Object System.Windows.Forms.Button
    $btnCap.Text = 'Desmarcar-ho tot'
    $btnCap.Size = New-Object System.Drawing.Size(120, 26)
    $btnCap.Location = New-Object System.Drawing.Point(445, 449)
    $btnCap.add_Click({
        for ($i = 0; $i -lt $llista.Items.Count; $i++) { $llista.SetItemChecked($i, $false) }
        & $refrescaTotal
    }.GetNewClosure())
    $form.Controls.Add($btnCap)

    $y = 478
    if ($nSenseCoord -gt 0 -or $nImpossibles -gt 0) {
        $avisos = @()
        if ($nSenseCoord -gt 0)   { $avisos += "$nSenseCoord sense coordenades a l'Excel" }
        if ($nImpossibles -gt 0)  { $avisos += "$nImpossibles amb coordenades impossibles (mira l'avís del final)" }
        $lblAvis = New-Object System.Windows.Forms.Label
        $lblAvis.Text = '(' + ($avisos -join '; ') + ": no es poden situar.)"
        $lblAvis.AutoSize = $false
        $lblAvis.Size = New-Object System.Drawing.Size(550, 20)
        $lblAvis.Location = New-Object System.Drawing.Point(15, $y)
        $lblAvis.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 20)
        $form.Controls.Add($lblAvis)
        $y += 22
    }

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Generar mapa'
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    $btnOk.Location = New-Object System.Drawing.Point(295, $y)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Enrere'
    $btnCancel.Size = New-Object System.Drawing.Size(130, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(435, $y)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $form.ClientSize = New-Object System.Drawing.Size(580, ($y + 46))

    & $omple

    $res = $form.ShowDialog()
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $noms = @()
    foreach ($i in $llista.CheckedIndices) { $noms += [string]$estat.Zones[$i].Nom }
    return [pscustomobject]@{ NomsZones = @($noms); NomesApilades = [bool]$chkApil.Checked }
}

# Finestra de progres de les consultes al Cadastre, amb Cancel.lar de veritat.
# Retorna un objecte amb el formulari i els seus controls; qui la crida ha de
# fer .Form.Close() al final.
function New-CoordProgress([int]$total) {
    $Script:CoordCancelat = $false
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Consultant el Cadastre'
    $form.Size = New-Object System.Drawing.Size(560, 190)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false; $form.MaximizeBox = $false
    $form.ControlBox = $false
    if ($null -ne $Script:CoordIcon) { $form.Icon = $Script:CoordIcon }

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(510, 42)
    $lbl.Text = "Demanant els portals de $total parcel·les..."
    $form.Controls.Add($lbl)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 66)
    $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Minimum = 0
    $bar.Maximum = [math]::Max($total, 1)
    $form.Controls.Add($bar)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Cancel·lar'
    $btn.Size = New-Object System.Drawing.Size(120, 30)
    $btn.Location = New-Object System.Drawing.Point(410, 100)
    $btn.add_Click({ $Script:CoordCancelat = $true })
    $form.Controls.Add($btn)

    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    return [pscustomobject]@{ Form = $form; Label = $lbl; Bar = $bar }
}

# ============================================================================
# MAIN
# ============================================================================
function Invoke-CoordenadesMain {
    # 1. Localitzar l'Excel.
    $xls = Find-LatestRutaExcel
    if ($null -eq $xls) {
        Show-CoordInfo ("No s'ha trobat cap base de dades d'activitats.`n`n" +
            "Busco un fitxer 'YYYY-MM-DD ACTIVITATS.xlsx' a:`n" +
            "  1. $ActivitatsDir`n" +
            "  2. $LocalActivitatsDir`n`n" +
            "Copia'n un a la carpeta local i torna a provar.") 'Coordenades' 'Warning'
        return
    }
    $dbLabel = if ($xls.Source -eq 'fallback') {
        "[FALLBACK LOCAL] Base de dades: $($xls.File.Name)"
    } else {
        "Base de dades: $($xls.File.Name)"
    }

    # 2. Llegir l'Excel.
    $espera = New-CoordProgress 1
    $espera.Label.Text = "Llegint $($xls.File.Name)..."
    $espera.Bar.Style = 'Marquee'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $lectura = Read-CoordenadesFromExcel $xls.File
    } catch {
        $espera.Form.Close()
        Show-CoordInfo "Error llegint l'Excel:`n$($_.Exception.Message)" 'Coordenades' 'Error'
        return
    } finally {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
    }

    $tots = @($lectura.Registres)
    if ($tots.Count -eq 0) {
        Show-CoordInfo "La fulla 'Estes' no te cap activitat amb coordenades." 'Coordenades' 'Warning'
        return
    }
    $apilats = @(Get-RegistresApilats $tots)
    if ($apilats.Count -eq 0) {
        Show-CoordInfo "No hi ha cap activitat apilada: totes tenen ja un punt propi." 'Coordenades' 'Information'
        return
    }

    # 3. Triar les ZONES. El repas de centenars d'activitats no es fa d'una
    # tirada: es va per quadres de 400 m, i cada tanda es la que caben en una
    # estona.
    $zonesApil = @(Get-ZonesAmbActivitats $apilats)
    $zonesTot  = @(Get-ZonesAmbActivitats $tots)
    $tria = Show-CoordenadesForm $dbLabel $zonesApil $zonesTot `
                                 ([int]$lectura.SenseCoord) (@($lectura.Impossibles).Count)
    if ($null -eq $tria) { return }
    if (@($tria.NomsZones).Count -eq 0) {
        Show-CoordInfo "No has triat cap zona." 'Coordenades' 'Warning'
        return
    }

    $base    = if ($tria.NomesApilades) { $apilats } else { $tots }
    $zonesOk = @{}
    foreach ($n in @($tria.NomsZones)) { $zonesOk[$n] = $true }
    $triats = @($base | Where-Object {
        (Test-CoordPlausible ([double]$_.UtmX) ([double]$_.UtmY)) -and
        $zonesOk.ContainsKey((Get-ZonaDeCoord ([double]$_.UtmX) ([double]$_.UtmY)))
    })
    $abast = ("{0} {1}" -f (@($tria.NomsZones) -join ', '),
                            $(if ($tria.NomesApilades) { '(apilades)' } else { '(totes)' }))
    if ($triats.Count -eq 0) {
        Show-CoordInfo "Les zones triades no tenen cap activitat." 'Coordenades' 'Warning'
        return
    }

    # 4. Portals del Cadastre, amb barra de progres i Cancel.lar.
    $refcats = @(Get-RefcatsAConsultar $triats)
    $prog = New-CoordProgress $refcats.Count
    $onProgress = {
        param($fetes, $total, $rc)
        $prog.Bar.Value = [math]::Min($fetes, $prog.Bar.Maximum)
        $prog.Label.Text = "Parcel·la $fetes de $total  ($rc)"
        [System.Windows.Forms.Application]::DoEvents()
        return (-not $Script:CoordCancelat)
    }.GetNewClosure()
    try {
        $portalsPerRc = Get-PortalsPerParcelles $refcats $onProgress
    } catch {
        $portalsPerRc = @{}
    } finally {
        if (-not $prog.Form.IsDisposed) { $prog.Form.Close() }
    }
    if ($Script:CoordCancelat) {
        Show-CoordInfo ("S'ha cancel·lat. El que ja s'havia demanat queda desat, aixi que si ho " +
                        "tornes a provar continuarà des d'on era.") 'Coordenades' 'Information'
        return
    }

    # 5. Muntar els punts del mapa.
    $items = @()
    foreach ($r in $triats) {
        $rc = Get-RefcatParcel $r.Rc
        $portals = @()
        if ($rc -ne '' -and $portalsPerRc.ContainsKey($rc)) { $portals = @($portalsPerRc[$rc]) }
        $items += New-ItemCoordenades $r $portals
    }

    # 5b. TOTS els portals de les parcel.les consultades, per pintar-los al mapa
    # amb el seu numero (com al planol del Cadastre). Ja els tenim demanats: fins
    # ara se'n feia servir un i la resta es llencaven.
    $portalsMapa = @()
    foreach ($rc in @($portalsPerRc.Keys)) {
        foreach ($p in @($portalsPerRc[$rc])) {
            if ($null -eq $p) { continue }
            $px = [double]$p.X; $py = [double]$p.Y
            if (-not (Test-CoordPlausible $px $py)) { continue }
            $ll = Convert-UtmToLatLon $px $py 31 $true
            $portalsMapa += [pscustomobject]@{
                Numero = [string]$p.Numero
                Via    = [string]$p.Via
                Lat    = $ll.Lat
                Lon    = $ll.Lon
            }
        }
    }

    # 6. Generar l'HTML i obrir-lo.
    $html = Build-CoordenadesHtml $items $dbLabel $abast $xls.File.Name $portalsMapa
    if (-not (Test-Path -LiteralPath $CoordOutputDir)) {
        New-Item -ItemType Directory -Path $CoordOutputDir -Force | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
    $outPath = Join-Path $CoordOutputDir "Coordenades_$stamp.html"
    [System.IO.File]::WriteAllText($outPath, $html, (New-Object System.Text.UTF8Encoding($false)))
    Start-Process $outPath

    # 7. Resum.
    $resum = Get-ResumPrecisio $items
    $nFac = [int]$resum['facana']
    $nDub = [int]$resum['facana-dubtosa']
    $nApr = [int]$resum['facana-aprox']
    $nCad = [int]$resum['cadastre']
    $msg  = "Mapa generat amb $(@($items).Count) activitats.`n"
    $msg += "Zones: $abast`n`n"
    $msg += "Portal exacte:                 $nFac`n"
    $msg += "Portal DUBTOS (mira'ls):       $nDub`n"
    $msg += "Portal mes proper:             $nApr`n"
    $msg += "Sense portal (es queden on eren): $nCad`n`n"
    $msg += "Al mapa: amplia fins que surtin els NUMEROS dels portals, valida`n"
    $msg += "amb un clic els punts que ja son bons i arrossega els que no.`n"
    $msg += "Despres, 'Baixar Excel (.xlsx)': hi surt tot el que hagis validat`n"
    $msg += "d'aquesta base de dades, tambe el d'altres zones i altres dies.`n`n"
    if (@($lectura.Impossibles).Count -gt 0) {
        $msg += "ATENCIO: aquestes activitats tenen unes coordenades IMPOSSIBLES a`n"
        $msg += "l'Excel i no es poden situar enlloc (sembla que els falten xifres):`n"
        foreach ($im in @($lectura.Impossibles)) {
            $msg += ("  GIA {0}  X={1}  Y={2}   {3}`n" -f $im.Id, $im.X, $im.Y, $im.Adreca)
        }
        $msg += "`n"
    }
    $msg += "Fitxer: $outPath"
    Show-CoordInfo $msg 'Coordenades'
}

if (-not $Script:CoordHeadless) {
    Invoke-CoordenadesMain
}
