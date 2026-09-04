# Proves automatiques de les funcions PURES de l'eina Coordenades
# (rutes/Coordenades.ps1 + rutes/Geocodificador.ps1).
#
# NO prova la part d'Excel (COM), la crida al Cadastre (xarxa) ni les finestres
# (WinForms): aixo nomes funciona a Windows / amb connexio. Aqui es valida la
# logica: parseig de la resposta INSPIRE, tria del portal, la guardia de
# distancia, la memoria cau, la deteccio d'activitats apilades i l'HTML del mapa.
#
# La resposta d'exemple del Cadastre es a dades/wfsAD-exemple.xml. ATENCIO:
# NO es una resposta gravada del servei real (des de l'entorn on es va escriure
# el codi el host del Cadastre estava bloquejat), sino un fitxer muntat seguint
# l'esquema INSPIRE Addresses. Per validar-lo contra el servei de veritat, fes
# doble clic a suport\rutes\Provar-Cadastre.bat
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File tests/run-tests-coordenades.ps1

$ErrorActionPreference = 'Stop'
# A Linux no existeix LOCALAPPDATA; el donem perque el dot-source no falli
# (mateix guard que run-tests.ps1 i run-tests-actextr.ps1).
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }
$env:COORDENADES_TEST = '1'   # mode headless: nomes defineix funcions

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'rutes' 'Coordenades.ps1')
. $scriptPath

. (Join-Path $PSScriptRoot 'TestLib.ps1')   # Assert / AssertEq / AssertNear / Write-TestSummary

$fixture = Get-Content -LiteralPath (Join-Path $PSScriptRoot (Join-Path 'dades' 'wfsAD-exemple.xml')) -Raw -Encoding UTF8

Write-Host "`n--- Get-RefcatParcel ---"
AssertEq (Get-RefcatParcel '2295827DF2729E0011RQ') '2295827DF2729E' 'referencia de 20 -> parcel.la de 14'
AssertEq (Get-RefcatParcel '4091106DF2749A')       '4091106DF2749A' 'referencia de 14 es queda igual'
AssertEq (Get-RefcatParcel '2295827df2729e0011rq') '2295827DF2729E' 'minuscules -> majuscules'
AssertEq (Get-RefcatParcel ' 2295827DF2729E 0011 ') '2295827DF2729E' 'espais i simbols fora'
AssertEq (Get-RefcatParcel '2295827DF272')         ''               'massa curta -> buida'
AssertEq (Get-RefcatParcel '')                     ''               'buida -> buida'
AssertEq (Get-RefcatParcel $null)                  ''               'null -> buida'

Write-Host "`n--- Get-NumeroPortal ---"
AssertEq (Get-NumeroPortal '147')     '147' 'numero simple'
AssertEq (Get-NumeroPortal '147-179') '147' 'rang -> el primer'
AssertEq (Get-NumeroPortal '13 B')    '13'  'numero amb lletra'
AssertEq (Get-NumeroPortal '1-3')     '1'   'rang curt'
AssertEq (Get-NumeroPortal 'S/N')     ''    'sense numero'
AssertEq (Get-NumeroPortal '')        ''    'buit'
AssertEq (Get-NumeroPortal $null)     ''    'null'

Write-Host "`n--- Get-ViaNormalitzada ---"
AssertEq (Get-ViaNormalitzada 'C/ CADIS')     'CADIS' 'treu el tipus de via abreujat'
AssertEq (Get-ViaNormalitzada 'CALLE CADIS')  'CADIS' 'treu el tipus de via sencer'
AssertEq (Get-ViaNormalitzada 'CARRER CADIS') 'CADIS' 'i el catala'
AssertEq (Get-ViaNormalitzada 'CARRETERA HOSPITALET') 'HOSPITALET' 'CARRETERA no es confon amb CARRER'
AssertEq (Get-ViaNormalitzada 'Doctor Ferran') 'DOCTOR FERRAN' 'majuscules'
AssertEq (Get-ViaNormalitzada 'Muntanya  del   Vent') 'MUNTANYA DEL VENT' 'espais collapsats'
AssertEq (Get-ViaNormalitzada '') '' 'buida'
AssertEq (Get-ViaNormalitzada $null) '' 'null'
# LES SIGLES DEL CADASTRE. Aquesta es la prova que hauria evitat que cap adreca
# coincidis mai per carrer: el Cadastre escriu 'CL CADIS' i, sense 'CL' a la
# llista, no es convertia en 'CADIS'. La tria del portal es feia nomes pel
# numero i s'agafava el del carrer del costat.
AssertEq (Get-ViaNormalitzada 'CL CADIS')  'CADIS'  'CL (calle, com ho escriu el Cadastre)'
AssertEq (Get-ViaNormalitzada 'CL CADIS')  (Get-ViaNormalitzada 'C CADIS') 'el Cadastre i l Excel han de coincidir'
AssertEq (Get-ViaNormalitzada 'CL HUELVA') 'HUELVA' 'l altre carrer de la illa'
AssertEq (Get-ViaNormalitzada 'PZ CATALUNYA') 'CATALUNYA' 'PZ (plaza)'
AssertEq (Get-ViaNormalitzada 'TR MIG')       'MIG'       'TR (travesia)'
AssertEq (Get-ViaNormalitzada 'GL ESPANYA')   'ESPANYA'   'GL (glorieta)'
AssertEq (Get-ViaNormalitzada 'CL DE LA PAU') 'DE LA PAU' 'nomes es treu la PRIMERA paraula'

Write-Host "`n--- ConvertFrom-CatastroAdXml (fixture INSPIRE) ---"
$portals = @(ConvertFrom-CatastroAdXml $fixture)
AssertEq $portals.Count 5 'la fixture porta 5 portals'
AssertEq (($portals | ForEach-Object { $_.Numero }) -join ',') '1,3,5,6,8' 'numeros, en ordre del document'
AssertEq $portals[0].Via 'CALLE CADIS'  'el nom de via es resol pel xlink:href'
AssertEq $portals[3].Via 'CALLE HUELVA' 'i el de l altre carrer tambe'
AssertNear $portals[0].X 421960.10 0.001 'X del primer portal'
AssertNear $portals[0].Y 4579500.20 0.001 'Y del primer portal'
# El parseig ha d'aguantar qualsevol cosa sense petar: si no s'enten, cap portal
# (i llavors cada activitat es queda amb la coordenada de l'Excel de sempre).
AssertEq @(ConvertFrom-CatastroAdXml 'aixo no es XML').Count 0 'text que no es XML -> cap portal'
AssertEq @(ConvertFrom-CatastroAdXml '').Count                0 'buit -> cap portal'
AssertEq @(ConvertFrom-CatastroAdXml $null).Count             0 'null -> cap portal'
AssertEq @(ConvertFrom-CatastroAdXml '<a><b/></a>').Count      0 'XML sense adreces -> cap portal'
# Adreca sense <pos>: s'omet, pero les altres s'han de llegir igualment.
$senseP = '<c xmlns:ad="urn:x"><ad:Address><ad:locator><ad:designator>9</ad:designator></ad:locator></ad:Address></c>'
AssertEq @(ConvertFrom-CatastroAdXml $senseP).Count 0 'adreca sense coordenades -> s omet'

Write-Host "`n--- ConvertFrom-CatastroAdXml: guardia d'eixos ---"
# En UTM 31N l'est (~420.000) va molt per sota del nord (~4.578.000). Si algun
# dia el servei els servis a l'inreves, s'han de tornar a posar al seu lloc.
$girat = $fixture -replace '<gml:pos>421960\.10 4579500\.20</gml:pos>', '<gml:pos>4579500.20 421960.10</gml:pos>'
$pg = @(ConvertFrom-CatastroAdXml $girat)
AssertNear $pg[0].X 421960.10  0.001 'eixos girats: X torna al seu lloc'
AssertNear $pg[0].Y 4579500.20 0.001 'eixos girats: Y torna al seu lloc'

Write-Host "`n--- Select-PortalFacana ---"
$t = Select-PortalFacana $portals 'CADIS' '3'
AssertEq $t.Precisio 'facana' 'Cadis 3: numero exacte'
AssertNear $t.X 421971.45 0.001 'Cadis 3: el portal que toca'
# El mateix numero existeix als dos carrers de la illa: el carrer ha de manar.
$t = Select-PortalFacana $portals 'HUELVA' '6'
AssertEq $t.Precisio 'facana' 'Huelva 6: numero exacte'
AssertNear $t.X 421944.20 0.001 'Huelva 6: el carrer desambigua'
# Un numero que no hi es: el mes proper de la MATEIXA BANDA del carrer.
$t = Select-PortalFacana $portals 'CADIS' '7'
AssertEq $t.Precisio 'facana-aprox' 'Cadis 7: no hi es, s aproxima'
AssertNear $t.X 421982.90 0.001 'Cadis 7 -> el 5 (senar, mateixa banda)'
$t = Select-PortalFacana $portals '' '2'
AssertNear $t.X 421944.20 0.001 'sense carrer, 2 -> el 6 (parell mes proper)'
$t = Select-PortalFacana $portals '' '7'
AssertNear $t.X 421982.90 0.001 'sense carrer, 7 -> el 5'
# Un carrer que no hi es no ha de descartar-ho tot: es cau als portals de la parcel.la.
$t = Select-PortalFacana $portals 'AVINGUDA INVENTADA' '3'
AssertEq $t.Precisio 'facana' 'carrer desconegut: encara val el numero de la parcel.la'
Assert ($null -eq (Select-PortalFacana @() 'CADIS' '3')) 'sense portals -> null'
Assert ($null -eq (Select-PortalFacana $null 'CADIS' '3')) 'portals null -> null'
# Un sol portal i cap numero: es pot fer servir, pero marcat com a aproximat.
$unic = @([pscustomobject]@{ Numero=''; Via='CALLE CADIS'; X=421960.0; Y=4579500.0 })
$t = Select-PortalFacana $unic 'CADIS' ''
AssertEq $t.Precisio 'facana-aprox' 'un sol portal sense numero -> aproximat'
# Diversos portals sense numero i sense saber el numero: NO ens la juguem.
$dos = @(
    [pscustomobject]@{ Numero=''; Via=''; X=421960.0; Y=4579500.0 }
    [pscustomobject]@{ Numero=''; Via=''; X=421980.0; Y=4579520.0 }
)
Assert ($null -eq (Select-PortalFacana $dos '' '')) 'ambigu del tot -> null'
# DOS portals amb el MATEIX numero al mateix carrer: la tria es una moneda a
# l'aire i s'ha de dir. Passa de debo a la illa de Cadis, on n'hi ha dos amb
# l'1 i un d'ells cau al centre de la parcel.la.
$repetit = @(
    [pscustomobject]@{ Numero='1'; Via='CL CADIS'; X=422069.174; Y=4579449.952 }
    [pscustomobject]@{ Numero='1'; Via='CL CADIS'; X=421968.090; Y=4579505.550 }
    [pscustomobject]@{ Numero='3'; Via='CL CADIS'; X=422060.714; Y=4579461.182 }
)
$t = Select-PortalFacana $repetit 'CADIS' '1'
AssertEq $t.Precisio 'facana-dubtosa' 'dos portals amb el numero 1 -> dubtos'
$t = Select-PortalFacana $repetit 'CADIS' '3'
AssertEq $t.Precisio 'facana' 'i el que nomes hi es un cop, segueix sent fiable'

Write-Host "`n--- Zones (graella de 400 m ancorada a un origen FIX) ---"
# L'ancoratge es constant a proposit: si la graella sortis del minim de les
# dades, una activitat nova mes a l'oest desplacaria TOTES les zones i 'la zona
# C6' voldria dir una altra cosa que la setmana passada.
AssertEq (Get-ZonaDeCoord 421968.09 4579505.55) 'F2' 'la illa de Cadis/Huelva es la F2'
AssertEq (Get-ZonaDeCoord 423912.16 4578928.25) 'E7' 'Ctra. de l Hospitalet 147 es la E7'
AssertEq (Get-ZonaDeCoord 421200.0  4577200.0)  'A1' 'el canto de la graella es A1'
AssertEq (Get-ZonaDeCoord 421599.9  4577599.9)  'A1' 'i tot el quadre de 400 m tambe'
AssertEq (Get-ZonaDeCoord 421600.0  4577600.0)  'B2' 'el quadre seguent, en diagonal'
AssertEq (Get-ZonaDeCoord 421200.0  4587200.0)  'Z1'  'la fila 25 es la Z'
AssertEq (Get-ZonaDeCoord 421200.0  4587600.0)  'AA1' 'i passada la Z, AA (no es queda sense lletres)'

Assert (Test-CoordPlausible 421968.09 4579505.55) 'una coordenada de Cornella es plausible'
# El GIA 1009 (Quintana i Millas 9) porta X=423,37 Y=4578,81: les xifres bones
# dividides per mil. Si es cola, el mapa s'estira fins a l'Atlantic.
Assert (-not (Test-CoordPlausible 423.37 4578.81)) 'les xifres partides per mil, NO'
Assert (-not (Test-CoordPlausible 0 0))            'l origen, NO'
Assert (-not (Test-CoordPlausible 421968.09 12.5)) 'una Y impossible, NO'

$regsZ = @(
    [pscustomobject]@{ Id='1'; Rc='2295827DF2729E0011RQ'; Carrer='CADIS';  Numero='19'; UtmX=421968.09; UtmY=4579505.55 }
    [pscustomobject]@{ Id='2'; Rc='2295827DF2729E0008RQ'; Carrer='HUELVA'; Numero='6';  UtmX=421968.09; UtmY=4579505.55 }
    [pscustomobject]@{ Id='3'; Rc='2295827DF2729E0003XL'; Carrer='HUELVA'; Numero='16'; UtmX=421970.00; UtmY=4579500.00 }
    [pscustomobject]@{ Id='4'; Rc='4091106DF2749A0006XJ'; Carrer='HOSPITALET'; Numero='147'; UtmX=423912.16; UtmY=4578928.25 }
    [pscustomobject]@{ Id='5'; Rc='XXXX'; Carrer='ENLLOC'; Numero='1'; UtmX=423.37; UtmY=4578.81 }
)
$zones = @(Get-ZonesAmbActivitats $regsZ)
AssertEq $zones.Count 2 'dues zones (la impossible no compta)'
AssertEq $zones[0].Nom 'F2' 'primer la que en te mes'
AssertEq $zones[0].Comptador 3 'i en te tres'
AssertEq $zones[0].Carrers 'HUELVA / CADIS' 'el nom de la zona surt dels carrers, el mes repetit primer'
AssertEq $zones[1].Nom 'E7' 'i la segona zona'
AssertEq $zones[1].Comptador 1 'amb una activitat'
AssertEq @(Get-ZonesAmbActivitats @()).Count 0 'sense registres, cap zona'
AssertEq (Get-CarrersDominants @()) '' 'sense carrers, nom buit'
AssertEq (Get-CarrersDominants $regsZ 1) 'HUELVA' 'el carrer mes repetit de tots'

Write-Host "`n--- Get-UtmDistanceM ---"
AssertNear (Get-UtmDistanceM 0 0 3 4) 5 0.0001 'triangle 3-4-5'
AssertNear (Get-UtmDistanceM 421960 4579500 421960 4579500) 0 0.0001 'mateix punt -> 0'

Write-Host "`n--- Resolve-CoordEstabliment (guardia de distancia) ---"
$r = Resolve-CoordEstabliment $portals 'CADIS' '3' 421968.09 4579505.55
AssertEq $r.Precisio 'facana' 'un portal a prop s accepta'
AssertNear $r.X 421971.45 0.001 'i es fa servir la seva X'
# Un portal a 8 km NO es d'aquesta parcel.la: val mes un marcador imprecis que
# un marcador mentider.
$lluny = @([pscustomobject]@{ Numero='3'; Via='CALLE CADIS'; X=430000.0; Y=4579500.0 })
$r = Resolve-CoordEstabliment $lluny 'CADIS' '3' 421968.09 4579505.55
AssertEq $r.Precisio 'cadastre' 'un portal a 8 km es descarta'
AssertNear $r.X 421968.09 0.001 'i es conserva la coordenada de l Excel'
$r = Resolve-CoordEstabliment @() 'CADIS' '3' 421968.09 4579505.55
AssertEq $r.Precisio 'cadastre' 'sense portals -> coordenada de l Excel'


Write-Host "`n--- Build-CatastroAdUrl ---"
$url = Build-CatastroAdUrl '2295827DF2729E'
Assert ($url -match 'wfsAD')             'la URL apunta al servei d adreces'
Assert ($url -match '2295827DF2729E')    'porta la referencia cadastral'
Assert ($url -match '25831')             'demana les coordenades en EPSG:25831'
Assert ($url -notmatch '\{0\}')          'la plantilla queda substituida'

Write-Host "`n--- Test-GeoCacheEntryValida ---"
$ara = [datetime]'2026-08-19T10:00:00'
$unPortal = @([pscustomobject]@{ Numero='1'; Via='X'; X=1.0; Y=2.0 })
$fresca  = [pscustomobject]@{ Data = '2026-08-01T10:00:00'; Portals = $unPortal }
$vella   = [pscustomobject]@{ Data = '2024-01-01T10:00:00'; Portals = $unPortal }
$buidaOk = [pscustomobject]@{ Data = '2026-08-10T10:00:00'; Portals = @() }
$buidaKo = [pscustomobject]@{ Data = '2026-05-01T10:00:00'; Portals = @() }
Assert (Test-GeoCacheEntryValida $fresca $ara)          'entrada recent amb portals: val'
Assert (-not (Test-GeoCacheEntryValida $vella $ara))    'entrada de fa dos anys: caducada'
Assert (Test-GeoCacheEntryValida $buidaOk $ara)         'entrada buida recent: val'
Assert (-not (Test-GeoCacheEntryValida $buidaKo $ara))  'entrada buida de fa mesos: es torna a provar'
Assert (-not (Test-GeoCacheEntryValida $null $ara))     'null: no val'
Assert (-not (Test-GeoCacheEntryValida ([pscustomobject]@{ Data='ahir'; Portals=@() }) $ara)) 'data illegible: no val'
Assert (-not (Test-GeoCacheEntryValida ([pscustomobject]@{ Data='2027-01-01T00:00:00'; Portals=$unPortal }) $ara)) 'data del futur: no val'

Write-Host "`n--- Get-ClauCoord / Get-RegistresApilats / Get-RefcatsAConsultar ---"
AssertEq (Get-ClauCoord 421968.09 4579505.55) '421968.09|4579505.55' 'clau amb 2 decimals'
AssertEq (Get-ClauCoord 421968.094 4579505.551) (Get-ClauCoord 421968.09 4579505.55) 'arrodoneix al centimetre'
AssertEq (Get-ClauCoord 421968 4579505) '421968.00|4579505.00' 'sempre amb 2 decimals (punt, no coma)'

# Tres activitats al mateix punt, dues a un altre, i una de sola.
$recs = @(
    [pscustomobject]@{ Id='1'; Rc='2295827DF2729E0011RQ'; Carrer='CADIS';  Numero='19'; UtmX=421968.09; UtmY=4579505.55; Adreca='C CADIS 19';  Activitat='A' }
    [pscustomobject]@{ Id='2'; Rc='2295827DF2729E0008RQ'; Carrer='HUELVA'; Numero='6';  UtmX=421968.09; UtmY=4579505.55; Adreca='C HUELVA 6';  Activitat='B' }
    [pscustomobject]@{ Id='3'; Rc='2295827DF2729E0003XL'; Carrer='HUELVA'; Numero='16'; UtmX=421968.09; UtmY=4579505.55; Adreca='C HUELVA 16'; Activitat='C' }
    [pscustomobject]@{ Id='4'; Rc='3085213DF2738E0116ZL'; Carrer='FERROCARRILS CATALANS'; Numero='177'; UtmX=422811.94; UtmY=4578281.76; Adreca='PG FC 177'; Activitat='D' }
    [pscustomobject]@{ Id='5'; Rc='3085213DF2738E0283QX'; Carrer='FERROCARRILS CATALANS'; Numero='117'; UtmX=422811.94; UtmY=4578281.76; Adreca='PG FC 117'; Activitat='E' }
    [pscustomobject]@{ Id='6'; Rc='2895225DF2729F0001XE'; Carrer='DOCTOR FERRAN'; Numero='13'; UtmX=422689.00; UtmY=4579272.54; Adreca='C DR FERRAN 13'; Activitat='F' }
)
$apilats = @(Get-RegistresApilats $recs)
AssertEq $apilats.Count 5 'els apilats son 5 (3 + 2); el que va sol queda fora'
AssertEq (($apilats | ForEach-Object { $_.Id }) -join ',') '1,2,3,4,5' 'i conserven l ordre d entrada'
AssertEq @(Get-RegistresApilats @()).Count 0 'llista buida -> buida'
$soles = @([pscustomobject]@{ Id='9'; Rc='X'; UtmX=1.0; UtmY=2.0 })
AssertEq @(Get-RegistresApilats $soles).Count 0 'una sola activitat mai esta apilada'

$refcats = @(Get-RefcatsAConsultar $recs)
AssertEq $refcats.Count 3 'les 6 activitats son de 3 parcel.les'
AssertEq ($refcats -join ',') '2295827DF2729E,2895225DF2729F,3085213DF2738E' 'parcel.les uniques i ordenades'
$senseRc = @([pscustomobject]@{ Id='9'; Rc=''; UtmX=1.0; UtmY=2.0 })
AssertEq @(Get-RefcatsAConsultar $senseRc).Count 0 'sense referencia cadastral no es consulta res'

Write-Host "`n--- New-ItemCoordenades ---"
$it = New-ItemCoordenades $recs[0] $portals
AssertEq $it.Id '1' 'conserva l ID'
AssertEq $it.Zona 'F2' 'i porta la seva zona (per al mapa i per a l Excel)'
AssertEq $it.Precisio 'facana-aprox' 'Cadis 19 no hi es: agafa el portal senar mes proper'
AssertNear $it.XExcel  421968.09 0.001 'la coordenada de l Excel es conserva intacta'
AssertNear $it.XFacana 421982.90 0.001 'i la nova es la del portal'
# Les dues coordenades han de venir tambe en graus, per al mapa.
AssertNear $it.LatExcel  41.363279 0.000001 'lat de l Excel (contrastada amb la inversa de Convert-UtmToLatLon)'
AssertNear $it.LonExcel  2.067034  0.000001 'lon de l Excel'
Assert ($it.LatFacana -ne $it.LatExcel) 'la lat de facana es diferent de la de l Excel'
# Sense portals, el punt verd ha de coincidir EXACTAMENT amb el vermell: aixi
# al mapa surt a sobre i l'usuari el pot arrossegar on toqui.
$itSense = New-ItemCoordenades $recs[5] @()
AssertEq $itSense.Precisio 'cadastre' 'sense portals -> cadastre'
AssertNear $itSense.XFacana $itSense.XExcel 0.0000001 'el verd comenca sobre el vermell (X)'
AssertNear $itSense.YFacana $itSense.YExcel 0.0000001 'el verd comenca sobre el vermell (Y)'
AssertNear $itSense.LatFacana $itSense.LatExcel 0.0000001 'i tambe en graus'

Write-Host "`n--- Get-ResumPrecisio ---"
$resum = Get-ResumPrecisio @($it, $itSense, $it)
AssertEq $resum['facana-aprox']   2 'compta els aproximats'
AssertEq $resum['cadastre']       1 'compta els que es queden al cadastre'
AssertEq $resum['facana']         0 'i els exactes, encara que siguin zero'
AssertEq $resum['facana-dubtosa'] 0 'i els dubtosos'

Write-Host "`n--- Build-CoordenadesHtml ---"
$items = @($it, $itSense)
$portalsMapa = @(
    [pscustomobject]@{ Numero='19'; Via='CL CADIS';  Lat=41.36537; Lon=2.06765 }
    [pscustomobject]@{ Numero='6';  Via='CL HUELVA'; Lat=41.36452; Lon=2.06780 }
)
$html = Build-CoordenadesHtml $items 'Base de dades: 2026-08-18 ACTIVITATS.xls' 'F2 (apilades)' '2026-08-18 ACTIVITATS.xls' $portalsMapa
Assert ($html -match 'leaflet')                 'inclou Leaflet'
Assert ($html -match '"id":"1"')                'inclou l ID de l activitat'
Assert ($html -match '421968\.09')              'inclou la coordenada de l Excel'
Assert ($html -match '421982\.9')               'inclou la coordenada de facana'
Assert ($html -match 'baixaExcel')              'inclou el boto de baixar l Excel'
Assert ($html -match 'draggable: true')         'el punt verd es pot arrossegar'
Assert ($html -match 'latLonToUtm31')           'inclou la projeccio per als punts moguts'
Assert ($html -match 'localStorage')            'les correccions es recorden al navegador'
Assert ($html -match 'validaVisibles')          'hi ha el boto de validar tot el que es veu'
Assert ($html -match 'ZOOM_NUMS')               'els numeros de portal tenen llindar de zoom'
Assert ($html -match 'chkNums')                 'i una casella per apagar-los'
Assert ($html -match '"zona":"F2"')             'cada activitat porta la seva zona al JSON'
# Els portals han d'arribar a la pagina amb el seu numero: son el que et deixa
# dir si un punt esta ben posat.
Assert ($html -match 'var PORTALS = (\[.*?\]);') 'hi ha la llista de portals'
$jsonP = $Matches[1]
$parsedP = $null
try { $parsedP = $jsonP | ConvertFrom-Json } catch { $parsedP = $null }
Assert ($null -ne $parsedP) 'el JSON dels portals es valid'
AssertEq @($parsedP).Count 2 'hi son els dos portals'
AssertEq $parsedP[0].n '19' 'amb el seu numero'
Assert ($html -match 'ACTIVITATS\.xls')         'inclou el nom de la base de dades'
Assert ($html -match '<meta charset="utf-8">')  'declara la codificacio'
# El JSON que s'incrusta ha de ser valid: si no, la pagina no arrenca.
Assert ($html -match 'var ITEMS   = (\[.*?\]);') 'hi ha la llista d items'
$json = $Matches[1]
$parsed = $null
try { $parsed = $json | ConvertFrom-Json } catch { $parsed = $null }
Assert ($null -ne $parsed) 'el JSON incrustat es valid'
AssertEq @($parsed).Count 2 'hi son les dues activitats'
AssertEq $parsed[0].prec 'facana-aprox' 'el JSON porta la precisio'
# AMB UNA SOLA ACTIVITAT. Aquesta es la prova que fallava de debo: el guard
# mirava el NOMBRE d'elements donant per fet que ConvertTo-Json en desembolcalla
# un de sol, i en el PowerShell de l'usuari NO ho fa -- sortia [[{...}]] i la
# pagina no arrencava. Ara es mira la SORTIDA, que es el que compta, i tant se
# val com es comporti cada versio. Amb zones petites, aixo passara sovint.
$html1 = Build-CoordenadesHtml @($it) 'db' 'abast' 'font.xls' @()
Assert ($html1 -match 'var ITEMS   = \[\{') 'amb un sol item, el JSON segueix sent una llista'
Assert ($html1 -notmatch 'var ITEMS   = \[\[') 'i NO queda embolcallat dues vegades'
# I sense cap activitat, la pagina no ha de petar.
$html0 = Build-CoordenadesHtml @() 'db' 'abast' 'font.xls' @()
Assert ($html0 -match 'var PORTALS = \[\];') 'sense portals, llista buida'
Assert ($html0 -match 'var ITEMS   = \[\];') 'sense activitats, llista buida'

exit (Write-TestSummary 'RESULTAT')
