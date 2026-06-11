<#
  Ruta.ps1 - Planificador de rutes d'inspeccio per a Cornella.

  Donat un llistat d'ID Activitat (els escrius/enganxes a una finestra), el
  programa:
    1. Els busca a la base de dades d'activitats (Excel, fulla "Estes"/"Estès").
    2. Els geolocalitza al mapa amb les columnes "UTM X" i "UTM Y"
       (coordenades ETRS89 / UTM fus 31N, EPSG:25831).
    3. Els identifica amb l'ID Activitat i l'adreca de l'emplacament
       ("Emp. Tipus via" + "Emp. Carrer" + "Emp. Numero" + "Emp. Lletra").
    4. Calcula la ruta CIRCULAR mes rapida que les visita totes i torna a
       l'inici (servei OSRM per carretera; si no hi ha xarxa, una aproximacio
       en linia recta).
    5. Genera un mapa HTML (Leaflet) que pots IMPRIMIR a PDF des del navegador.

  Es un programa INDEPENDENT: l'unica cosa que comparteix amb la resta es el
  fitxer opcional config.ps1 (rutes i, si vols, el servidor de rutes). No
  dot-sourceja GenerarInforme.ps1 (no necessita Word ni l'assistent de passos).

  Mode "headless" per a proves: si $env:RUTA_TEST o $env:GENINFORME_TEST estan
  definides, NOMES es defineixen les funcions (no s'obre cap finestra ni
  s'executa Main). Aixi es poden provar les funcions pures a Linux sense Office.
#>

$ErrorActionPreference = 'Stop'

$Script:HeadlessTest = [bool]$env:RUTA_TEST -or [bool]$env:GENINFORME_TEST
if (-not $Script:HeadlessTest) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
# Arrel del clone (un nivell amunt: suport/.. = informes-Cornella/).
$RepoRoot   = Split-Path -Parent $ScriptRoot

# ----------------------------------------------------------------------------
# Configuracio per defecte (sobreescriptible des de suport/config.ps1).
# ----------------------------------------------------------------------------
# Carpeta on viu l'Excel d'activitats a la xarxa de la feina.
$ActivitatsDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
# Carpeta on es desen els mapes de ruta generats (a l'arrel del clone, fora
# de suport/, perque l'usuari els trobi facilment). Ignorada per git. Es fa
# servir un nom propi ($RutesOutputDir, no $OutputDir) per NO barrejar-se amb
# la carpeta d'informes .docx que comparteix config.ps1 amb GenerarInforme.
$RutesOutputDir = Join-Path $RepoRoot 'Rutes generades'
# Servidor de rutes OSRM. Per defecte el servidor public de demostracio
# (router.project-osrm.org). NOMES s'hi envien coordenades (mai noms ni
# adreces). Si tens un OSRM propi, posa la seva URL base a config.ps1:
#   $OsrmBaseUrl = 'http://localhost:5000'
# Si el deixes buit ('') o no hi ha xarxa, s'usa una ruta aproximada local.
$OsrmBaseUrl   = 'https://router.project-osrm.org'

# Punt de SORTIDA de la ruta (la base des d'on surts). La ruta comencara
# SEMPRE per l'activitat mes propera a aquest punt (i hi tornara al final).
# Per defecte: Carrer de l'Energia, 97 (Cornella de Llobregat), en coordenades
# UTM (mateix sistema que la base de dades: ETRS89/31N). Per canviar la base,
# posa unes altres coordenades a config.ps1. Si les deixes a 0 (o buides), la
# ruta comenca per la primera activitat que escriguis a la llista.
$RutaOrigenUtmX = 424456.0   # Carrer Energia 97
$RutaOrigenUtmY = 4578205.0  # Carrer Energia 97

# config.ps1 viu al costat d'aquest .ps1 (dins de suport/). Opcional.
$configPath = Join-Path $ScriptRoot 'config.ps1'
if (Test-Path -LiteralPath $configPath) {
    . $configPath
}

# Carpeta local de fallback (mateixa que fa servir GenerarInforme): si no hi
# ha xarxa de la feina, s'agafa l'Excel mes recent d'aqui.
$LocalActivitatsDir = Join-Path $RepoRoot 'BASE DE DADES ACTIVITATS'

# Mapeig de columnes (1-based) de la fulla "Estes"/"Estès" que necessitem.
$Script:RutaColumns = @{
    ID        = 1    # ID Activitat
    UTMX      = 3    # UTM X
    UTMY      = 4    # UTM Y
    TIPUS_VIA = 48   # Emp. Tipus via
    CARRER    = 49   # Emp. Carrer
    NUMERO    = 50   # Emp. Numero
    LLETRA    = 52   # Emp. Lletra
}

# ============================================================================
# FUNCIONS PURES (provables en mode headless, sense Office)
# ============================================================================

# Escapa text per a HTML (sense dependre de System.Web, que no esta a totes
# les plataformes de PowerShell). Cobreix els caracters perillosos habituals.
function _HtmlEncode($s) {
    if ($null -eq $s) { return '' }
    $t = [string]$s
    $t = $t -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    $t = $t -replace "'", '&#39;'
    return $t
}

# Normalitza text Unicode (sense diacritics, minuscules) per a comparacions.
function _RutaNormalize($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC).Trim().ToLowerInvariant()
}

# Parseja la llista d'IDs que escriu l'usuari. Accepta separadors: comes,
# punts i comes, espais, tabuladors i salts de linia. Treu duplicats
# conservant l'ordre d'aparicio. Retorna un array de cadenes.
function ConvertFrom-IdList([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $parts = $text -split '[\s,;]+'
    $seen = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($p in $parts) {
        $id = $p.Trim()
        if ($id -eq '') { continue }
        if (-not $seen.Contains($id)) { [void]$seen.Add($id, $true) }
    }
    return @($seen.Keys)
}

# Converteix un valor de cel·la (numero o text, amb coma o punt decimal) a
# double. Retorna $null si es buit o no es un numero valid.
function ConvertTo-UtmNumber($value) {
    if ($null -eq $value) { return $null }
    if ($value -is [double]) { return [double]$value }
    $s = ([string]$value).Trim()
    if ($s -eq '') { return $null }
    $s = $s -replace ',', '.'
    $n = 0.0
    $ok = [double]::TryParse($s, [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)
    if ($ok) { return $n }
    return $null
}

# Munta l'adreca de l'emplacament a partir de les parts, descartant buits i
# valors com '   '. No hi posa ciutat (ja se sap que es Cornella).
function Format-EmpAddress([string]$tipusVia, [string]$carrer, [string]$numero, [string]$lletra) {
    $parts = @($tipusVia, $carrer, $numero, $lletra) |
        ForEach-Object { if ($null -ne $_) { ([string]$_).Trim() } else { '' } } |
        Where-Object { $_ -ne '' }
    return ($parts -join ' ')
}

# Converteix coordenades UTM (fus 31N nord, ETRS89/WGS84) a latitud/longitud
# en graus decimals (WGS84). Inversa de la projeccio Transversa de Mercator
# (series de Snyder). Validat contra pyproj (EPSG:25831 -> EPSG:4326).
function Convert-UtmToLatLon([double]$easting, [double]$northing, [int]$zone = 31, [bool]$north = $true) {
    $a  = 6378137.0
    $f  = 1.0 / 298.257223563
    $k0 = 0.9996
    $e2  = $f * (2 - $f)
    $ep2 = $e2 / (1 - $e2)
    $E0 = 500000.0
    $N0 = if ($north) { 0.0 } else { 10000000.0 }

    $x = $easting  - $E0
    $y = $northing - $N0
    $M  = $y / $k0
    $mu = $M / ($a * (1 - $e2/4 - 3*[math]::Pow($e2,2)/64 - 5*[math]::Pow($e2,3)/256))
    $e1 = (1 - [math]::Sqrt(1 - $e2)) / (1 + [math]::Sqrt(1 - $e2))

    $phi1 = $mu `
        + (3*$e1/2 - 27*[math]::Pow($e1,3)/32) * [math]::Sin(2*$mu) `
        + (21*[math]::Pow($e1,2)/16 - 55*[math]::Pow($e1,4)/32) * [math]::Sin(4*$mu) `
        + (151*[math]::Pow($e1,3)/96) * [math]::Sin(6*$mu) `
        + (1097*[math]::Pow($e1,4)/512) * [math]::Sin(8*$mu)

    $sinPhi = [math]::Sin($phi1)
    $cosPhi = [math]::Cos($phi1)
    $tanPhi = [math]::Tan($phi1)
    $C1 = $ep2 * $cosPhi * $cosPhi
    $T1 = $tanPhi * $tanPhi
    $N1 = $a / [math]::Sqrt(1 - $e2 * $sinPhi * $sinPhi)
    $R1 = $a * (1 - $e2) / [math]::Pow(1 - $e2 * $sinPhi * $sinPhi, 1.5)
    $D  = $x / ($N1 * $k0)

    $lat = $phi1 - ($N1 * $tanPhi / $R1) * ( `
        [math]::Pow($D,2)/2 `
        - (5 + 3*$T1 + 10*$C1 - 4*$C1*$C1 - 9*$ep2) * [math]::Pow($D,4)/24 `
        + (61 + 90*$T1 + 298*$C1 + 45*$T1*$T1 - 252*$ep2 - 3*$C1*$C1) * [math]::Pow($D,6)/720)

    $lon0 = ($zone * 6 - 183) * [math]::PI / 180.0
    $lon = $lon0 + ( `
        $D `
        - (1 + 2*$T1 + $C1) * [math]::Pow($D,3)/6 `
        + (5 - 2*$C1 + 28*$T1 - 3*$C1*$C1 + 8*$ep2 + 24*$T1*$T1) * [math]::Pow($D,5)/120) / $cosPhi

    return [pscustomobject]@{
        Lat = $lat * 180.0 / [math]::PI
        Lon = $lon * 180.0 / [math]::PI
    }
}

# Distancia en metres entre dos punts (lat/lon) per la formula de l'haversine.
function Get-HaversineMeters([double]$lat1, [double]$lon1, [double]$lat2, [double]$lon2) {
    $R = 6371000.0
    $toRad = [math]::PI / 180.0
    $dLat = ($lat2 - $lat1) * $toRad
    $dLon = ($lon2 - $lon1) * $toRad
    $s = [math]::Sin($dLat/2)*[math]::Sin($dLat/2) +
         [math]::Cos($lat1*$toRad)*[math]::Cos($lat2*$toRad)*[math]::Sin($dLon/2)*[math]::Sin($dLon/2)
    return $R * 2 * [math]::Atan2([math]::Sqrt($s), [math]::Sqrt(1-$s))
}

# Retorna l'index del punt (llista d'objectes {Lat,Lon}) mes proper a una
# coordenada origen (lat/lon). -1 si la llista es buida.
function Get-NearestPointIndex($points, [double]$originLat, [double]$originLon) {
    $n = @($points).Count
    if ($n -eq 0) { return -1 }
    $best = -1; $bestD = [double]::MaxValue
    for ($i = 0; $i -lt $n; $i++) {
        $d = Get-HaversineMeters $originLat $originLon $points[$i].Lat $points[$i].Lon
        if ($d -lt $bestD) { $bestD = $d; $best = $i }
    }
    return $best
}

# Reordena la llista de punts perque el mes proper a l'origen quedi el PRIMER
# (la resta conserva l'ordre relatiu). Aixi la ruta circular comenca i acaba
# a prop de la base. Retorna una nova llista (array). Si no hi ha origen valid
# o la llista te menys de 2 punts, la retorna sense canvis.
function Set-StartNearest($points, $originLat, $originLon) {
    $arr = @($points)
    if ($arr.Count -lt 2 -or $null -eq $originLat -or $null -eq $originLon) { return $arr }
    $idx = Get-NearestPointIndex $arr ([double]$originLat) ([double]$originLon)
    if ($idx -le 0) { return $arr }   # ja es el primer o no trobat
    $new = @($arr[$idx])
    for ($i = 0; $i -lt $arr.Count; $i++) {
        if ($i -ne $idx) { $new += $arr[$i] }
    }
    return $new
}

# Resol una aproximacio del TSP circular (veí mes proper + millora 2-opt)
# sobre una matriu de punts {Lat,Lon}, comencant SEMPRE pel punt 0. Retorna
# un array d'indexs (l'ordre de visita; no repeteix el 0 al final). S'usa com
# a fallback quan no hi ha servidor de rutes.
function Get-TspOrder($points) {
    $n = @($points).Count
    if ($n -le 2) { return @(0..([math]::Max($n-1,0))) }

    # Matriu de distancies (haversine).
    $dist = New-Object 'double[,]' $n, $n
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $n; $j++) {
            if ($i -eq $j) { $dist[$i,$j] = 0 }
            else { $dist[$i,$j] = Get-HaversineMeters $points[$i].Lat $points[$i].Lon $points[$j].Lat $points[$j].Lon }
        }
    }

    # Veí mes proper des de 0.
    $visited = New-Object 'bool[]' $n
    $order = New-Object System.Collections.ArrayList
    $cur = 0
    $visited[0] = $true
    [void]$order.Add(0)
    for ($step = 1; $step -lt $n; $step++) {
        $best = -1; $bestD = [double]::MaxValue
        for ($j = 0; $j -lt $n; $j++) {
            if (-not $visited[$j] -and $dist[$cur,$j] -lt $bestD) { $bestD = $dist[$cur,$j]; $best = $j }
        }
        $visited[$best] = $true
        [void]$order.Add($best)
        $cur = $best
    }

    # Millora 2-opt (cicle tancat: torna a l'inici). Mantenim el punt 0 fix.
    $route = @($order)
    $improved = $true
    $guard = 0
    while ($improved -and $guard -lt 100) {
        $improved = $false
        $guard++
        for ($i = 1; $i -lt $route.Count - 1; $i++) {
            for ($k = $i + 1; $k -lt $route.Count; $k++) {
                $a = $route[$i-1]; $b = $route[$i]
                $c = $route[$k]; $d = $route[($k+1) % $route.Count]
                if ($a -eq $c -or $b -eq $d) { continue }
                $delta = ($dist[$a,$c] + $dist[$b,$d]) - ($dist[$a,$b] + $dist[$c,$d])
                if ($delta -lt -0.01) {
                    # Inverteix el segment [i..k].
                    $seg = @($route[$i..$k])
                    [array]::Reverse($seg)
                    $new = @()
                    if ($i -gt 0) { $new += $route[0..($i-1)] }
                    $new += $seg
                    if ($k + 1 -lt $route.Count) { $new += $route[($k+1)..($route.Count-1)] }
                    $route = $new
                    $improved = $true
                }
            }
        }
    }
    return @($route)
}

# Interpreta la resposta del servei /trip d'OSRM (ConvertFrom-Json) i retorna
# l'ordre de visita, la geometria (llista de [lat,lon]) i els totals. Retorna
# $null si la resposta no es valida.
function ConvertFrom-OsrmTrip($resp, [int]$count) {
    if ($null -eq $resp -or $resp.code -ne 'Ok') { return $null }
    if ($null -eq $resp.trips -or @($resp.trips).Count -eq 0) { return $null }
    $trip = $resp.trips[0]

    # Ordre de visita: cada waypoint d'entrada porta el seu 'waypoint_index'
    # (posicio dins de la ruta optimitzada). Ordenem els indexs d'entrada.
    $pairs = @()
    for ($i = 0; $i -lt $count; $i++) {
        $wi = $resp.waypoints[$i].waypoint_index
        $pairs += [pscustomobject]@{ Input = $i; Visit = [int]$wi }
    }
    $order = @($pairs | Sort-Object Visit | ForEach-Object { $_.Input })

    # Geometria GeoJSON: coordenades [lon,lat] -> [lat,lon] per a Leaflet.
    $geom = @()
    if ($trip.geometry -and $trip.geometry.coordinates) {
        foreach ($c in $trip.geometry.coordinates) {
            $geom += ,@([double]$c[1], [double]$c[0])
        }
    }

    return [pscustomobject]@{
        Order        = $order
        Geometry     = $geom
        DistanceM    = [double]$trip.distance
        DurationS    = [double]$trip.duration
    }
}

# Genera el document HTML del mapa. $stops es una llista ordenada (ordre de
# visita) d'objectes amb: Order (1..n), Id, Address, Lat, Lon. $geometry es la
# llista de [lat,lon] de la linia de la ruta (pot ser $null). $mode es un text
# descriptiu ('xarxa' o 'aproximada'). Retorna l'HTML com a cadena.
function Build-RouteHtml($stops, $geometry, [double]$distanceM, [double]$durationS, [string]$mode, [string]$dbLabel) {
    $stopsJson = ConvertTo-Json @($stops | ForEach-Object {
        @{ order = $_.Order; id = $_.Id; address = $_.Address; lat = $_.Lat; lon = $_.Lon }
    }) -Depth 5 -Compress
    if (@($stops).Count -eq 1) { $stopsJson = "[$stopsJson]" }  # ConvertTo-Json no embolcalla 1 element

    $geomJson = if ($geometry -and @($geometry).Count -gt 0) {
        ConvertTo-Json @($geometry) -Depth 5 -Compress
    } else { 'null' }

    $km    = [math]::Round($distanceM / 1000.0, 1)
    $mins  = [math]::Round($durationS / 60.0)
    $modeText = if ($mode -eq 'xarxa') {
        "Ruta per carretera (OSRM) &middot; $km km &middot; aprox. $mins min"
    } elseif ($mode -eq 'aproximada') {
        "Ruta aproximada en linia recta (sense xarxa) &middot; $km km en total"
    } else {
        "$([int](@($stops).Count)) parada(es)"
    }

    $rows = ''
    foreach ($s in $stops) {
        $addr = _HtmlEncode $s.Address
        if ([string]::IsNullOrWhiteSpace($addr)) { $addr = '<i>(sense adreca)</i>' }
        $rows += "<tr><td class='num'>$($s.Order)</td><td class='id'>$(_HtmlEncode ([string]$s.Id))</td><td>$addr</td></tr>`n"
    }

    $today = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    $dbEnc = _HtmlEncode $dbLabel

    $html = @"
<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ruta d'inspeccio - Cornella</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
      integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=" crossorigin=""/>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; font-family: Segoe UI, Arial, sans-serif; color: #1a1a1a; }
  #top { background: #14365c; color: #fff; padding: 10px 16px; display: flex; align-items: baseline; gap: 14px; flex-wrap: wrap; }
  #top h1 { font-size: 18px; margin: 0; }
  #top .meta { font-size: 13px; opacity: .9; }
  #wrap { display: flex; height: calc(100vh - 92px); }
  #map { flex: 1 1 auto; }
  #side { width: 340px; overflow: auto; border-left: 1px solid #ddd; padding: 0 0 40px 0; }
  #side h2 { font-size: 14px; margin: 14px 14px 6px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid #eee; vertical-align: top; }
  th { background: #f3f5f8; position: sticky; top: 0; }
  td.num { font-weight: bold; color: #fff; }
  tr td.num { background: #c0392b; text-align: center; width: 30px; border-radius: 0; }
  td.id { font-family: Consolas, monospace; color: #14365c; white-space: nowrap; }
  #bar { padding: 8px 16px; border-top: 1px solid #ddd; display: flex; gap: 10px; align-items: center; }
  button { background: #14365c; color: #fff; border: 0; padding: 8px 16px; border-radius: 5px; cursor: pointer; font-size: 14px; }
  button:hover { background: #1d4d82; }
  .marker-num { background: #c0392b; color: #fff; border: 2px solid #fff; border-radius: 50%;
                width: 26px; height: 26px; line-height: 22px; text-align: center; font-weight: bold;
                box-shadow: 0 1px 4px rgba(0,0,0,.4); font-size: 13px; }
  .marker-start { background: #1e8449; }
  @media print {
    #bar, .leaflet-control-zoom, .leaflet-control-attribution { display: none !important; }
    #wrap { height: 540px; }
    #side { width: 300px; }
    @page { size: A4 landscape; margin: 10mm; }
  }
</style>
</head>
<body>
<div id="top">
  <h1>Ruta d'inspeccio &mdash; Cornella de Llobregat</h1>
  <span class="meta">$modeText</span>
  <span class="meta">Generat: $today</span>
</div>
<div id="wrap">
  <div id="map"></div>
  <div id="side">
    <h2>Ordre de visita</h2>
    <table>
      <thead><tr><th>#</th><th>ID</th><th>Adreca</th></tr></thead>
      <tbody>
$rows
      </tbody>
    </table>
    <p style="font-size:11px;color:#777;margin:10px 14px;">Base de dades: $dbEnc</p>
  </div>
</div>
<div id="bar">
  <button onclick="window.print()">&#128424;&#65039; Imprimir / Desar com a PDF</button>
  <span style="font-size:12px;color:#555;">Al dialeg d'impressio tria "Desar com a PDF" / "Microsoft Print to PDF".</span>
</div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
        integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=" crossorigin=""></script>
<script>
  var stops = $stopsJson;
  var routeGeom = $geomJson;
  var map = L.map('map');
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19, attribution: '&copy; OpenStreetMap'
  }).addTo(map);

  var bounds = [];
  stops.forEach(function (s) {
    var isStart = (s.order === 1);
    var icon = L.divIcon({
      className: '', html: '<div class="marker-num' + (isStart ? ' marker-start' : '') + '">' + s.order + '</div>',
      iconSize: [26, 26], iconAnchor: [13, 13]
    });
    var addr = s.address && s.address.length ? s.address : '(sense adreca)';
    L.marker([s.lat, s.lon], { icon: icon })
     .addTo(map)
     .bindPopup('<b>Parada ' + s.order + '</b><br>ID ' + s.id + '<br>' + addr);
    bounds.push([s.lat, s.lon]);
  });

  if (routeGeom && routeGeom.length > 1) {
    var line = L.polyline(routeGeom, { color: '#c0392b', weight: 4, opacity: .8 }).addTo(map);
    map.fitBounds(line.getBounds().pad(0.15));
  } else if (bounds.length === 1) {
    map.setView(bounds[0], 16);
  } else if (bounds.length > 1) {
    map.fitBounds(L.latLngBounds(bounds).pad(0.15));
  } else {
    map.setView([41.355, 2.073], 14);
  }
</script>
</body>
</html>
"@
    return $html
}

# ============================================================================
# LECTURA D'EXCEL (COM) - nomes a Windows amb Excel; no es prova en headless.
# ============================================================================

# Cerca el fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' mes recent en una carpeta.
function _RutaFindLatestIn($dir) {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $regex = '^(\d{4}-\d{2}-\d{2})\s+ACTIVITATS\.(xls|xlsx)$'
    $cands = Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $regex } |
        ForEach-Object {
            if ($_.Name -match $regex) {
                [pscustomobject]@{ File = $_; Date = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $null) }
            }
        } | Sort-Object Date -Descending
    if (@($cands).Count -eq 0) { return $null }
    return $cands[0]
}

# Xarxa de la feina primer; despres fallback local del clone.
function Find-LatestRutaExcel {
    $r = _RutaFindLatestIn $ActivitatsDir
    if ($null -ne $r) { Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'primary' -Force; return $r }
    $r = _RutaFindLatestIn $LocalActivitatsDir
    if ($null -ne $r) { Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'fallback' -Force; return $r }
    return $null
}

# Localitza la fulla "Estes"/"Estès" acceptant variants Unicode.
function _RutaFindEstesSheet($wb) {
    foreach ($s in $wb.Sheets) {
        if ((_RutaNormalize $s.Name) -eq 'estes') { return $s }
    }
    return $null
}

# Llegeix TOTA la fulla "Estes" i retorna una hashtable [string ID] ->
# objecte amb UtmX, UtmY (cadenes crues) i les parts de l'adreca.
function Read-ActivitatsForRoute($excelFile) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            $sh = _RutaFindEstesSheet $wb
            if ($null -eq $sh) { throw "No s'ha trobat la fulla 'Estes'/'Estès' al fitxer Excel." }
            $data = $sh.UsedRange.Value2
            if ($null -eq $data) { return @{} }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)
            $C = $Script:RutaColumns
            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }
            $byId = @{}
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, $C.ID]
                if ($null -eq $cell) { continue }
                $id = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                $id = $id.Trim()
                if ($id -eq '') { continue }
                $byId[$id] = [pscustomobject]@{
                    Id       = $id
                    UtmX     = & $get $r $C.UTMX
                    UtmY     = & $get $r $C.UTMY
                    TipusVia = & $get $r $C.TIPUS_VIA
                    Carrer   = & $get $r $C.CARRER
                    Numero   = & $get $r $C.NUMERO
                    Lletra   = & $get $r $C.LLETRA
                }
            }
            return $byId
        } finally {
            $wb.Close($false)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# Demana la ruta optimitzada a OSRM (servei /trip). Retorna el resultat de
# ConvertFrom-OsrmTrip o $null si falla (xarxa, servidor, etc.).
function Invoke-OsrmTrip($points) {
    if ([string]::IsNullOrWhiteSpace($OsrmBaseUrl)) { return $null }
    $coords = (@($points | ForEach-Object {
        "{0},{1}" -f ([string]$_.Lon), ([string]$_.Lat)
    }) -join ';')
    $url = "$($OsrmBaseUrl.TrimEnd('/'))/trip/v1/driving/$coords`?source=first&roundtrip=true&geometries=geojson&overview=full"
    try {
        $resp = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 25
        return ConvertFrom-OsrmTrip $resp (@($points).Count)
    } catch {
        Write-Host "Avis: no s'ha pogut contactar amb el servidor de rutes ($($_.Exception.Message)). Uso ruta aproximada."
        return $null
    }
}

# ============================================================================
# INTERFICIE (WinForms) - nomes en us normal.
# ============================================================================

# Finestra per escriure/enganxar la llista d'IDs. Retorna l'array d'IDs o
# $null si l'usuari cancel·la.
function Show-IdInputForm([string]$dbLabel) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Planificador de rutes d'inspeccio"
    $form.Size = New-Object System.Drawing.Size(520, 420)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false; $form.MinimizeBox = $false

    $lblDb = New-Object System.Windows.Forms.Label
    $lblDb.Text = $dbLabel
    $lblDb.AutoSize = $false
    $lblDb.Size = New-Object System.Drawing.Size(480, 20)
    $lblDb.Location = New-Object System.Drawing.Point(15, 12)
    $lblDb.ForeColor = [System.Drawing.Color]::FromArgb(20, 54, 92)
    $form.Controls.Add($lblDb)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Escriu o enganxa els ID Activitat a visitar (separats per espais, comes o salts de linia):"
    $lbl.AutoSize = $false
    $lbl.Size = New-Object System.Drawing.Size(480, 36)
    $lbl.Location = New-Object System.Drawing.Point(15, 38)
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ScrollBars = 'Vertical'
    $txt.Size = New-Object System.Drawing.Size(480, 250)
    $txt.Location = New-Object System.Drawing.Point(15, 78)
    $txt.Font = New-Object System.Drawing.Font('Consolas', 11)
    $form.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Generar ruta'
    $btnOk.Size = New-Object System.Drawing.Size(120, 32)
    $btnOk.Location = New-Object System.Drawing.Point(255, 338)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'
    $btnCancel.Size = New-Object System.Drawing.Size(120, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(380, 338)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return ConvertFrom-IdList $txt.Text
}

function Show-Info([string]$msg, [string]$title = 'Rutes', [string]$icon = 'Information') {
    [System.Windows.Forms.MessageBox]::Show($msg, $title, 'OK', $icon) | Out-Null
}

# ============================================================================
# MAIN
# ============================================================================
function Invoke-RutaMain {
    # 1. Localitzar l'Excel.
    $xls = Find-LatestRutaExcel
    if ($null -eq $xls) {
        Show-Info ("No s'ha trobat cap base de dades d'activitats.`n`n" +
            "Busco un fitxer 'YYYY-MM-DD ACTIVITATS.xlsx' a:`n" +
            "  1. $ActivitatsDir`n" +
            "  2. $LocalActivitatsDir`n`n" +
            "Copia'n un a la carpeta local i torna a provar.") 'Rutes' 'Warning'
        return
    }
    $dbLabel = if ($xls.Source -eq 'fallback') {
        "[FALLBACK LOCAL] Base de dades: $($xls.File.Name)"
    } else {
        "Base de dades: $($xls.File.Name)"
    }

    # 2. Demanar la llista d'IDs.
    $ids = Show-IdInputForm $dbLabel
    if ($null -eq $ids) { return }
    if (@($ids).Count -eq 0) {
        Show-Info "No has indicat cap ID Activitat." 'Rutes' 'Warning'
        return
    }

    # 3. Llegir l'Excel.
    try {
        $byId = Read-ActivitatsForRoute $xls.File
    } catch {
        Show-Info "Error llegint l'Excel:`n$($_.Exception.Message)" 'Rutes' 'Error'
        return
    }

    # 4. Resoldre cada ID -> punt amb coordenades.
    $points = New-Object System.Collections.ArrayList
    $missing = @()
    $noCoords = @()
    foreach ($id in $ids) {
        if (-not $byId.ContainsKey($id)) { $missing += $id; continue }
        $rec = $byId[$id]
        $x = ConvertTo-UtmNumber $rec.UtmX
        $y = ConvertTo-UtmNumber $rec.UtmY
        if ($null -eq $x -or $null -eq $y) { $noCoords += $id; continue }
        $ll = Convert-UtmToLatLon $x $y 31 $true
        [void]$points.Add([pscustomobject]@{
            Id      = $id
            Address = (Format-EmpAddress $rec.TipusVia $rec.Carrer $rec.Numero $rec.Lletra)
            Lat     = $ll.Lat
            Lon     = $ll.Lon
        })
    }

    # 5. Avisar de problemes.
    $warns = @()
    if (@($missing).Count -gt 0)  { $warns += "No trobats a la base de dades: $($missing -join ', ')" }
    if (@($noCoords).Count -gt 0) { $warns += "Sense coordenades UTM (no es poden situar): $($noCoords -join ', ')" }
    if (@($points).Count -eq 0) {
        Show-Info ("No s'ha pogut situar cap activitat al mapa.`n`n" + ($warns -join "`n")) 'Rutes' 'Warning'
        return
    }

    # 5b. La ruta ha de comencar per l'activitat mes propera a la base
    # (per defecte Carrer Energia 97). Reordenem perque quedi la primera; com
    # que tant OSRM (source=first) com el TSP local comencen pel punt 0, amb
    # aixo n'hi ha prou.
    $oX = ConvertTo-UtmNumber $RutaOrigenUtmX
    $oY = ConvertTo-UtmNumber $RutaOrigenUtmY
    if ($null -ne $oX -and $null -ne $oY -and $oX -ne 0 -and $oY -ne 0) {
        $origin = Convert-UtmToLatLon $oX $oY 31 $true
        $reordered = Set-StartNearest $points $origin.Lat $origin.Lon
        $points = New-Object System.Collections.ArrayList
        foreach ($p in $reordered) { [void]$points.Add($p) }
    }

    # 6. Calcular la ruta circular.
    $mode = 'cap'
    $geometry = $null
    $distanceM = 0.0
    $durationS = 0.0
    $ordered = @(0..(@($points).Count - 1))

    if (@($points).Count -ge 2) {
        $trip = Invoke-OsrmTrip $points
        if ($null -ne $trip -and @($trip.Order).Count -eq @($points).Count) {
            $ordered = $trip.Order
            $geometry = $trip.Geometry
            $distanceM = $trip.DistanceM
            $durationS = $trip.DurationS
            $mode = 'xarxa'
        } else {
            # Fallback: TSP local + linies rectes.
            $ordered = Get-TspOrder $points
            $mode = 'aproximada'
            # Geometria recta tancada + distancia haversine.
            $geometry = @()
            $prev = $null
            foreach ($idx in $ordered) {
                $p = $points[$idx]
                $geometry += ,@($p.Lat, $p.Lon)
                if ($null -ne $prev) { $distanceM += Get-HaversineMeters $prev.Lat $prev.Lon $p.Lat $p.Lon }
                $prev = $p
            }
            # Tancar el cicle.
            $first = $points[$ordered[0]]
            $last  = $points[$ordered[-1]]
            $distanceM += Get-HaversineMeters $last.Lat $last.Lon $first.Lat $first.Lon
            $geometry += ,@($first.Lat, $first.Lon)
        }
    }

    # 7. Muntar la llista ordenada de parades (Order 1..n).
    $stops = New-Object System.Collections.ArrayList
    $n = 1
    foreach ($idx in $ordered) {
        $p = $points[$idx]
        [void]$stops.Add([pscustomobject]@{
            Order   = $n
            Id      = $p.Id
            Address = $p.Address
            Lat     = $p.Lat
            Lon     = $p.Lon
        })
        $n++
    }

    # 8. Generar l'HTML.
    $html = Build-RouteHtml $stops $geometry $distanceM $durationS $mode $xls.File.Name

    if (-not (Test-Path -LiteralPath $RutesOutputDir)) {
        New-Item -ItemType Directory -Path $RutesOutputDir -Force | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
    $outPath = Join-Path $RutesOutputDir "Ruta_$stamp.html"
    [System.IO.File]::WriteAllText($outPath, $html, (New-Object System.Text.UTF8Encoding($false)))

    # 9. Obrir al navegador.
    Start-Process $outPath

    # 10. Resum.
    $summary = "Ruta generada amb $(@($stops).Count) parada(es)."
    if ($mode -eq 'xarxa') {
        $summary += "`nDistancia: $([math]::Round($distanceM/1000.0,1)) km, temps aprox. $([math]::Round($durationS/60.0)) min."
    } elseif ($mode -eq 'aproximada') {
        $summary += "`n(Ruta aproximada en linia recta: sense connexio al servidor de rutes.)"
    }
    if (@($warns).Count -gt 0) { $summary += "`n`nAvisos:`n" + ($warns -join "`n") }
    $summary += "`n`nFitxer: $outPath`nObre'l i fes 'Imprimir / Desar com a PDF' per tenir-la en PDF."
    Show-Info $summary 'Ruta generada' 'Information'
}

if (-not $Script:HeadlessTest) {
    Invoke-RutaMain
}
