# Proves automatiques de les funcions PURES de Ruta.ps1 (planificador de rutes).
#
# NO prova la part d'Excel (COM), la crida a OSRM (xarxa) ni les finestres
# (WinForms): aixo nomes funciona a Windows / amb connexio. Aqui es validen
# les funcions de logica: parseig d'IDs i numeros, construccio d'adreces,
# conversio UTM->lat/lon, haversine, TSP i interpretacio de la resposta OSRM.
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File tests/run-tests-ruta.ps1

$ErrorActionPreference = 'Stop'
$env:RUTA_TEST = '1'   # mode headless: nomes defineix funcions, no obre res

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Ruta.ps1'
. $scriptPath

$script:pass = 0
$script:fail = 0
function Assert($cond, $name) {
    if ($cond) { $script:pass++; Write-Host "  OK   $name" -ForegroundColor Green }
    else       { $script:fail++; Write-Host "  FAIL $name" -ForegroundColor Red }
}
function AssertEq($actual, $expected, $name) {
    Assert ([string]$actual -eq [string]$expected) "$name (esperat '$expected', obtingut '$actual')"
}
function AssertNear($actual, $expected, $tol, $name) {
    Assert ([math]::Abs([double]$actual - [double]$expected) -le $tol) "$name (esperat ~$expected, obtingut $actual)"
}

Write-Host "`n--- ConvertFrom-IdList ---"
AssertEq ((ConvertFrom-IdList '1429 1428,1427') -join '|') '1429|1428|1427' 'separadors espai i coma'
AssertEq ((ConvertFrom-IdList "1429`n1428;1427`t1426") -join '|') '1429|1428|1427|1426' 'salts, punt i coma, tab'
AssertEq ((ConvertFrom-IdList '1429 1429 1428') -join '|') '1429|1428' 'treu duplicats conservant ordre'
AssertEq (@(ConvertFrom-IdList '   ').Count) 0 'nomes espais -> buit'
AssertEq (@(ConvertFrom-IdList $null).Count) 0 'null -> buit'

Write-Host "`n--- ConvertTo-UtmNumber ---"
AssertNear (ConvertTo-UtmNumber '422843.04') 422843.04 0.001 'punt decimal'
AssertNear (ConvertTo-UtmNumber '422843,04') 422843.04 0.001 'coma decimal'
AssertNear (ConvertTo-UtmNumber ([double]4577731.73)) 4577731.73 0.001 'double directe'
Assert ($null -eq (ConvertTo-UtmNumber ' '))   'espai -> null'
Assert ($null -eq (ConvertTo-UtmNumber ''))    'buit -> null'
Assert ($null -eq (ConvertTo-UtmNumber 'abc')) 'no numeric -> null'

Write-Host "`n--- Format-EmpAddress ---"
AssertEq (Format-EmpAddress 'AV' 'BAIX LLOBREGAT' 'S/N' '   ') 'AV BAIX LLOBREGAT S/N' 'descarta lletra buida'
AssertEq (Format-EmpAddress 'C' 'QUINTANA I MILLAS' '7' '9')   'C QUINTANA I MILLAS 7 9' 'totes les parts'
AssertEq (Format-EmpAddress '   ' '   ' '   ' '   ')           ''                       'tot buit -> cadena buida'

Write-Host "`n--- Convert-UtmToLatLon (validat contra pyproj EPSG:25831) ---"
$ll = Convert-UtmToLatLon 422843.04 4577731.73 31 $true
AssertNear $ll.Lat 41.347387 0.00001 'latitud Cornella'
AssertNear $ll.Lon 2.077719  0.00001 'longitud Cornella'
$ll2 = Convert-UtmToLatLon 422753.21 4579729.52 31 $true
AssertNear $ll2.Lat 41.365371 0.00001 'latitud 2n punt'
AssertNear $ll2.Lon 2.076391  0.00001 'longitud 2n punt'

Write-Host "`n--- Get-HaversineMeters ---"
# Dos punts a ~1 km: comprovem ordre de magnitud.
$d = Get-HaversineMeters 41.34739 2.07772 41.35603 2.08174
AssertNear $d 1020 80 'distancia ~1 km entre 2 activitats'
AssertNear (Get-HaversineMeters 41.3 2.0 41.3 2.0) 0 0.001 'mateix punt -> 0'

Write-Host "`n--- Get-TspOrder ---"
# 4 punts en quadrat: l'ordre optim ha de fer el perimetre, no les diagonals.
$sq = @(
    [pscustomobject]@{ Lat=41.30; Lon=2.00 }   # 0 (inici)
    [pscustomobject]@{ Lat=41.30; Lon=2.05 }   # 1
    [pscustomobject]@{ Lat=41.35; Lon=2.05 }   # 2
    [pscustomobject]@{ Lat=41.35; Lon=2.00 }   # 3
)
$ord = Get-TspOrder $sq
AssertEq @($ord).Count 4 'retorna tots els punts'
AssertEq $ord[0] 0 'comenca sempre pel punt 0'
Assert (@($ord | Sort-Object) -join ',' -eq '0,1,2,3') 'es una permutacio (cap repetit)'
# El cicle optim del quadrat es 0-1-2-3 (o 0-3-2-1). En tots dos, els veins
# de 0 son 1 i 3 (no el 2, que es la diagonal).
$pos = @{}; for ($i=0;$i -lt $ord.Count;$i++){ $pos[$ord[$i]] = $i }
$neigh = @($ord[($pos[0]+1)%4], $ord[($pos[0]+3)%4]) | Sort-Object
Assert (($neigh -join ',') -eq '1,3') 'el punt 0 limita amb 1 i 3 (perimetre, no diagonal)'

Write-Host "`n--- Get-NearestPointIndex / Set-StartNearest (sortida des de la base) ---"
$pts = @(
    [pscustomobject]@{ Id='A'; Lat=41.300; Lon=2.000 }
    [pscustomobject]@{ Id='B'; Lat=41.352; Lon=2.097 }   # ~Carrer Energia 97
    [pscustomobject]@{ Id='C'; Lat=41.380; Lon=2.050 }
)
# Origen = Carrer Energia 97 (lat/lon derivat de l'UTM 424456,4578205).
$oLat = 41.351804; $oLon = 2.096944
AssertEq (Get-NearestPointIndex $pts $oLat $oLon) 1 'el mes proper a Energia 97 es B (index 1)'
AssertEq (Get-NearestPointIndex @() $oLat $oLon) -1 'llista buida -> -1'
$re = Set-StartNearest $pts $oLat $oLon
AssertEq $re[0].Id 'B' 'Set-StartNearest posa B (mes proper) el primer'
AssertEq @($re).Count 3 'Set-StartNearest no perd punts'
AssertEq (($re | ForEach-Object { $_.Id }) -join '') 'BAC' 'conserva l ordre relatiu de la resta'
$one = @([pscustomobject]@{ Id='X'; Lat=41.3; Lon=2.0 })
AssertEq (Set-StartNearest $one $oLat $oLon)[0].Id 'X' 'amb 1 punt el deixa igual'

Write-Host "`n--- ConvertFrom-OsrmTrip ---"
$resp = [pscustomobject]@{
    code = 'Ok'
    waypoints = @(
        [pscustomobject]@{ waypoint_index = 2 }
        [pscustomobject]@{ waypoint_index = 0 }
        [pscustomobject]@{ waypoint_index = 1 }
    )
    trips = @(
        [pscustomobject]@{
            distance = 5000.0; duration = 600.0
            geometry = [pscustomobject]@{ coordinates = @( @(2.07,41.34), @(2.08,41.35) ) }
        }
    )
}
$t = ConvertFrom-OsrmTrip $resp 3
AssertEq ($t.Order -join ',') '1,2,0' 'ordena els indexs d entrada per waypoint_index'
AssertEq $t.DistanceM 5000 'distancia del trip'
AssertEq $t.DurationS 600  'durada del trip'
AssertEq @($t.Geometry).Count 2 'geometria amb 2 punts'
AssertEq ($t.Geometry[0] -join ',') '41.34,2.07' 'geometria convertida a [lat,lon]'
Assert ($null -eq (ConvertFrom-OsrmTrip ([pscustomobject]@{ code='NoRoute' }) 3)) 'code != Ok -> null'
Assert ($null -eq (ConvertFrom-OsrmTrip $null 3)) 'null -> null'

Write-Host "`n--- Build-RouteHtml ---"
$stops = @(
    [pscustomobject]@{ Order=1; Id='1429'; Address='AV BAIX LLOBREGAT S/N'; Lat=41.347; Lon=2.077 }
    [pscustomobject]@{ Order=2; Id='1427'; Address='C QUINTANA I MILLAS 7 9'; Lat=41.356; Lon=2.081 }
)
$geom = @( ,@(41.347,2.077), ,@(41.356,2.081) )
$html = Build-RouteHtml $stops $geom 5000 600 'xarxa' '2026-06-11 ACTIVITATS.xlsx'
Assert ($html -match 'leaflet') 'inclou Leaflet'
Assert ($html -match '1429') 'inclou l ID de la parada'
Assert ($html -match 'QUINTANA I MILLAS') 'inclou l adreca'
Assert ($html -match 'window.print') 'inclou el boto d imprimir'
Assert ($html -match '5 km') 'mostra els km de la ruta'

$summaryColor = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "`n========================================"
Write-Host ("RESULTAT: {0} OK, {1} FAIL" -f $script:pass, $script:fail) -ForegroundColor $summaryColor
Write-Host "========================================"
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
