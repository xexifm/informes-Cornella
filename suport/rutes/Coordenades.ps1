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
  (Motor.ps1 -> Start-EinaRutes) pero corre en el seu propi ambit.

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
# Cadastre.ps1 (el comu a totes les consultes al Cadastre: xarxa, memoria cau,
# bucle amb progres) va al davant: Geocodificador.ps1 s'hi recolza.
. (Join-Path $ScriptRoot 'Cadastre.ps1')
. (Join-Path $ScriptRoot 'CadastreParceles.ps1')   # el punt del Cadastre de cada parcel.la (compartit amb el Planol)
. (Join-Path $ScriptRoot 'Geocodificador.ps1')

# ----------------------------------------------------------------------------
# Reutilitzem les funcions de Ruta.ps1 (Find-LatestRutaExcel, Find-HeaderColumn,
# Read-FullaEstesa, ConvertTo-UtmNumber, Format-EmpAddress,
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

# L'"Excel per importar" (una copia de la base amb les coordenades corregides
# en vermell). Nomes defineix funcions.
. (Join-Path $ScriptRoot 'CoordenadesImportar.ps1')
# Omplir la plantilla del mapa i el JSON dins del <script> (comu als mapes).
. (Join-Path $ScriptRoot 'MapaHtml.ps1')

# Les finestres comunes de les eines de 'rutes/' (missatge, progres amb
# Cancel.lar i la icona). Nomes defineix.
. (Join-Path $ScriptRoot 'EinesUi.ps1')
$Script:EinaTitol = 'Coordenades'
if (-not $Script:CoordHeadless) { $Script:EinaIcon = Get-EinaIcon $SuportDir }

# Carpeta de sortida: local\geocodificacio\ (la mateixa on viu la memoria cau
# dels portals). Dins del clone pero fora del repositori.
$CoordOutputDir = Get-LocalSubdir $RepoRoot 'Geocodificacio'

# La pagina del mapa (HTML + JavaScript), amb marques {{nom}} per a les dades.
# Vegeu Build-CoordenadesHtml.
$Script:CoordPlantillaMapa = Join-Path $ScriptRoot 'CoordenadesMapa.html'

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
#
# $puntCad (opcional): @(x, y), el punt de la parcel.la al Cadastre
# (Get-ParcelesCadastre). L'usuari (octubre 2026): "ha de quedar ben clar
# aquelles coordenades que han sigut modificades respecte les del Cadastre".
# Corregida = l'Excel ja no hi es (a mes de $Script:CoordCorregidaM): el mapa
# la pinta lila i n'hi dibuixa la linia de punts.
$Script:CoordCorregidaM = 1.0
function New-ItemCoordenades($record, $portals, $puntCad = $null) {
    $x = [double]$record.UtmX
    $y = [double]$record.UtmY
    $coord = Resolve-CoordEstabliment $portals $record.Carrer $record.Numero $x $y
    $llExcel  = Convert-UtmToLatLon $x $y 31 $true
    $llFacana = Convert-UtmToLatLon ([double]$coord.X) ([double]$coord.Y) 31 $true
    $cad = $null; $distCad = $null
    if ($null -ne $puntCad -and @($puntCad).Count -eq 2) {
        $cx = [double]@($puntCad)[0]; $cy = [double]@($puntCad)[1]
        $llC = Convert-UtmToLatLon $cx $cy 31 $true
        $cad = [pscustomobject]@{ X = $cx; Y = $cy; Lat = $llC.Lat; Lon = $llC.Lon }
        $distCad = [math]::Sqrt(($x - $cx) * ($x - $cx) + ($y - $cy) * ($y - $cy))
    }
    return [pscustomobject]@{
        Id        = [string]$record.Id
        Zona      = (Get-ZonaDeCoord $x $y)
        Rc        = [string]$record.Rc
        Adreca    = [string]$record.Adreca
        Activitat = [string]$record.Activitat
        Titular   = [string]$record.Titular
        AdrecaTitular = [string]$record.AdrecaTitular
        XExcel    = $x
        YExcel    = $y
        LatExcel  = $llExcel.Lat
        LonExcel  = $llExcel.Lon
        XFacana   = [double]$coord.X
        YFacana   = [double]$coord.Y
        LatFacana = $llFacana.Lat
        LonFacana = $llFacana.Lon
        Precisio  = [string]$coord.Precisio
        Cadastre  = $cad
        DistCadastre = $distCad
        Corregida = ($null -ne $distCad -and $distCad -ge $Script:CoordCorregidaM)
    }
}

# Els registres segons el filtre de la finestra de tria: 'apilades' (les que
# comparteixen punt amb una altra), 'noapilades' (la resta) o 'totes'. PURA.
function Get-RegistresPerAbast($records, [string]$abast) {
    $arr = @($records)
    if ($abast -eq 'totes') { return $arr }
    $apil = @{}
    foreach ($r in @(Get-RegistresApilats $arr)) { $apil[[string]$r.Id + '|' + (Get-ClauCoord ([double]$r.UtmX) ([double]$r.UtmY))] = $true }
    $esApil = { param($r) $apil.ContainsKey([string]$r.Id + '|' + (Get-ClauCoord ([double]$r.UtmX) ([double]$r.UtmY))) }
    if ($abast -eq 'apilades') { return @($arr | Where-Object { & $esApil $_ }) }
    return @($arr | Where-Object { -not (& $esApil $_) })
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
# La pagina es la plantilla CoordenadesMapa.html; aqui nomes es calculen les
# dades que s'hi injecten. Dins de la plantilla hi ha tres peces que val la
# pena tenir localitzades:
#   latLonToUtm31()  la projeccio DIRECTA (lat/lon -> UTM 31N). Cal perque el
#                    Leaflet ens dona graus quan s'arrossega un punt i nosaltres
#                    hem d'exportar metres. Es la inversa exacta de
#                    Convert-UtmToLatLon (Ruta.ps1); verificada d'anada i
#                    tornada sobre tot el terme municipal amb un error maxim de
#                    0,07 mm.
#   buildXlsx()      escriu un .xlsx de veritat sense cap biblioteca: un .xlsx
#                    es un ZIP amb cinc XML a dins, i amb el metode "sense
#                    compressio" nomes cal el CRC-32 i les capçaleres del ZIP.
#   desaItem()       els punts que valides o mous van al localStorage del navegador,
#                    amb clau del fitxer d'origen. Si tanques la pagina i la
#                    tornes a obrir, hi son.
# $filtre: el filtre amb que s'obre el mapa ('tots' o 'avis', les marcades per
# revisar, que nomes les sap el navegador).
# $triaZones: el mapa porta TOTES les zones i les tries al planol (la graella
# de 400 m es dibuixa i es clica); $zonesInicials, les que ja venien marcades a
# la llista de la finestra.
function Build-CoordenadesHtml($items, [string]$dbLabel, [string]$abast, [string]$fontName, $portals, [string]$filtre = 'tots',
                               [bool]$triaZones = $false, $zonesInicials = @()) {
    $arr = @($items)
    $itemsJson = ConvertTo-JsonScript @($arr | ForEach-Object {
        # [ordered]: sense aixo, ConvertTo-Json treu les propietats en un ordre
        # diferent a cada execucio i l'HTML generat canvia sense que hagin
        # canviat les dades (la mateixa trampa que ja hi havia a Ruta.ps1).
        [ordered]@{
            id        = [string]$_.Id
            zona      = [string]$_.Zona
            rc        = [string]$_.Rc
            adreca    = [string]$_.Adreca
            activitat = [string]$_.Activitat
            titular   = [string]$_.Titular
            adt       = [string]$_.AdrecaTitular
            xe        = [double]$_.XExcel
            ye        = [double]$_.YExcel
            late      = [double]$_.LatExcel
            lone      = [double]$_.LonExcel
            xf        = [double]$_.XFacana
            yf        = [double]$_.YFacana
            latf      = [double]$_.LatFacana
            lonf      = [double]$_.LonFacana
            prec      = [string]$_.Precisio
            # El punt de la parcel.la al Cadastre (o null) i si l'Excel ja no hi es.
            latc      = if ($null -ne $_.Cadastre) { [double]$_.Cadastre.Lat } else { $null }
            lonc      = if ($null -ne $_.Cadastre) { [double]$_.Cadastre.Lon } else { $null }
            # I en metres: "torna-la al punt de la parcel.la" ha d'exportar la
            # coordenada del Cadastre TAL QUAL, sense anar i tornar de graus.
            xc        = if ($null -ne $_.Cadastre) { [double]$_.Cadastre.X } else { $null }
            yc        = if ($null -ne $_.Cadastre) { [double]$_.Cadastre.Y } else { $null }
            cor       = [int][bool]$_.Corregida
            dc        = if ($null -ne $_.DistCadastre) { [math]::Round([double]$_.DistCadastre, 1) } else { $null }
        }
    }) -Llista -Fondaria 5
    # (Una LLISTA sempre, i '</' escapat: ConvertTo-JsonScript, MapaHtml.ps1.
    # El guard de la llista mira la SORTIDA: al PC de l'usuari ConvertTo-Json no
    # desembolcalla un array d'un sol element i el mapa sortia amb [[{...}]].)

    # Els portals de les parcel.les consultades, per pintar-los amb el seu
    # numero com al planol del Cadastre. Poden ser cap.
    $portalsJson = ConvertTo-JsonScript @(@($portals) | Where-Object { $null -ne $_ } | ForEach-Object {
        [ordered]@{ n = [string]$_.Numero; v = [string]$_.Via; lat = [double]$_.Lat; lon = [double]$_.Lon }
    }) -Llista -Fondaria 5

    $resum = Get-ResumPrecisio $arr
    $today = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    $dbEnc = _HtmlEncode $dbLabel
    # El nom del fitxer d'origen va al JavaScript com a literal JSON (i no
    # HTML-escapat): es la clau amb que es desen les correccions al navegador,
    # i ha de ser el nom EXACTE perque les correccions d'una base es quedin
    # amb aquella base.
    $fontJson = ConvertTo-JsonScript ([string]$fontName)
    $abastEnc = _HtmlEncode $abast
    $nTot = $arr.Count
    $nFac = [int]$resum['facana']
    $nDub = [int]$resum['facana-dubtosa']
    $nApr = [int]$resum['facana-aprox']
    $nCad = [int]$resum['cadastre']

    # El HTML i el JavaScript del mapa viuen a CoordenadesMapa.html, al costat
    # d'aquest fitxer. Abans eren un here-string de 645 linies aqui dins, amb
    # dues trampes permanents: qualsevol '$' o '`' del JavaScript era una
    # interpolacio de PowerShell, i tot el text catala depenia que aquest .ps1
    # no perdes el BOM. Ara la plantilla es llegeix en UTF-8 EXPLICIT (el 5.1,
    # sense dir-li res, la llegiria com a ANSI) i nomes porta marques {{nom}}.
    $valors = @{
        nTot        = [string]$nTot
        abastEnc    = $abastEnc
        nFac        = [string]$nFac
        nDub        = [string]$nDub
        nApr        = [string]$nApr
        nCad        = [string]$nCad
        today       = $today
        dbEnc       = $dbEnc
        itemsJson   = $itemsJson
        portalsJson = $portalsJson
        fontJson    = $fontJson
        filtreJson  = (ConvertTo-JsonScript ([string]$filtre))
        # La graella (el mateix origen i mida que Get-ZonaDeCoord: una zona del
        # planol es el mateix rectangle que el de la llista), o null.
        triaJson    = if ($triaZones) {
                          ConvertTo-JsonScript ([ordered]@{ x0 = $CoordZonaX0; y0 = $CoordZonaY0; m = $CoordZonaMetres
                                                            inicials = @(@($zonesInicials) | ForEach-Object { [string]$_ }) }) -Fondaria 3
                      } else { 'null' }
    }
    return (Get-PlantillaHtml $Script:CoordPlantillaMapa $valors)
}

# ============================================================================
# LECTURA D'EXCEL (COM) - nomes a Windows amb Excel; no es prova en headless.
# ============================================================================

# EL TITULAR: Get-ColumnaTitular viu a Ruta.ps1, amb Find-HeaderColumn (la
# fa servir tambe el Planol activitats).


# L'ADRECA SENCERA (octubre 2026, l'usuari: "posa'm tota l'adreca, no nomes
# carrer i numero"): la via i el numero (Format-EmpAddress) i, darrere, el
# bloc, l'escala, el pis i la porta, amb les etiquetes del Planol activitats.
# Serveix per a la de l'activitat (Emp.) i per a la del titular (Rao soc.).
# PURA.
function Format-AdrecaSencera([string]$base, [string]$bloc, [string]$escala, [string]$pis, [string]$porta) {
    $parts = @()
    if ($base.Trim() -ne '') { $parts += $base.Trim() }
    foreach ($p in @(@('Bl.', $bloc), @('Esc.', $escala), @('Pl.', $pis), @('Pt.', $porta))) {
        $v = ([string]$p[1]).Trim()
        if ($v -ne '') { $parts += ($p[0] + ' ' + $v) }
    }
    return ($parts -join ', ')
}

# Llegeix la fulla "Estes" i retorna un registre per activitat amb tot el que
# necessitem: { Id; Rc; Adreca (sencera); Carrer; Numero; Activitat; Titular;
# AdrecaTitular; UtmX; UtmY }.
# Les activitats SENSE coordenades s'ometen (no es poden situar al mapa) i es
# compten a part.

function Read-CoordenadesFromExcel($excelFile) {
    $out = Read-FullaEstesa $excelFile {
        param($x)
        $data = $x.Data; $rows = $x.Rows; $cols = $x.Cols; $headers = $x.Headers
        if ($null -eq $data) { return [pscustomobject]@{ Registres = @(); SenseCoord = 0; Impossibles = @() } }

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
        $colTit  = Get-ColumnaTitular $headers
        # La resta de l'adreca de l'activitat, i la del titular (columnes que,
        # si no hi son, surten buides: el lector de cel.la torna '').
        $colEmp = @{}
        foreach ($k in @('Bloc', 'Escala', 'Pis', 'Porta')) { $colEmp[$k] = Find-HeaderColumn $headers ('Emp. ' + $k) }
        $colRao = @{}
        foreach ($k in @('Tipus via', 'Carrer', 'Numero', 'Escala', 'Pis', 'Porta')) { $colRao[$k] = Find-HeaderColumn $headers ('Rao soc. ' + $k) }

        if ($colUtmX -lt 1 -or $colUtmY -lt 1) {
            throw "La fulla 'Estes' no te les columnes 'UTM X' i 'UTM Y'."
        }

        $get = $x.Cel   # el lector de cel·la ve amb el context (Excel.ps1)

        $registres = @()
        $senseCoord = 0
        $impossibles = @()
        for ($r = 2; $r -le $rows; $r++) {
            $idCell = if ($colId -ge 1 -and $colId -le $cols) { $data[$r, $colId] } else { $null }
            $id = Get-IdDeCella $idCell
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
                Adreca    = (Format-AdrecaSencera (Format-EmpAddress (& $get $r $colVia) $carrerRaw $numeroRaw (& $get $r $colLlet)) `
                                (& $get $r $colEmp['Bloc']) (& $get $r $colEmp['Escala']) (& $get $r $colEmp['Pis']) (& $get $r $colEmp['Porta']))
                Carrer    = $carrerRaw
                Numero    = $numeroRaw
                Activitat = $activitat
                Titular   = if ($colTit -ge 1) { [string](& $get $r $colTit) } else { '' }
                AdrecaTitular = (Format-AdrecaSencera (Format-EmpAddress (& $get $r $colRao['Tipus via']) (& $get $r $colRao['Carrer']) (& $get $r $colRao['Numero']) '') `
                                '' (& $get $r $colRao['Escala']) (& $get $r $colRao['Pis']) (& $get $r $colRao['Porta']))
                UtmX      = $x
                UtmY      = $y
            }
        }
        return [pscustomobject]@{ Registres = @($registres); SenseCoord = $senseCoord; Impossibles = @($impossibles) }
    }
    return $out
}

# ============================================================================
# INTERFICIE (WinForms) - nomes en us normal.
# ============================================================================


# Una consulta al Cadastre amb la barra de progres i Cancel.lar. Torna
# { Resultat; Cancelat }; si la consulta peta, Resultat es un hashtable buit.
function _CoordAmbProgres($refcats, [string]$que, [scriptblock]$feina) {
    $prog = New-EinaProgres @($refcats).Count
    $onProgress = {
        param($fetes, $total, $rc)
        $prog.Bar.Value = [math]::Min($fetes, $prog.Bar.Maximum)
        $prog.Label.Text = "$que - parcel" + [char]0x00B7 + "la $fetes de $total  ($rc)"
        [System.Windows.Forms.Application]::DoEvents()
        return (-not $prog.Estat.Cancelat)
    }.GetNewClosure()
    $r = @{}
    try { $r = & $feina @($refcats) $onProgress } catch { $r = @{} } finally {
        if (-not $prog.Form.IsDisposed) { $prog.Form.Close() }
    }
    return [pscustomobject]@{ Resultat = $r; Cancelat = [bool]$prog.Estat.Cancelat }
}

# Finestra de tria: quines ZONES es repassen.
#
# Retorna { NomsZones; Abast ('apilades'|'noapilades'|'totes'); AmagaCorregides;
# NomesAvis; AlPlanol }, { Accio = 'importar' } o $null si es cancel.la.
#
# NOTA sobre "quantes en portes de repassades": aqui NO es pot saber. El que has
# validat viu al localStorage del NAVEGADOR, i el PowerShell no hi te acces. El
# progres, per tant, el mostra el mapa (que si que hi te acces) i aquesta
# finestra nomes diu quantes activitats hi ha a cada zona.
function Show-CoordenadesForm([string]$dbLabel, $zonesApilades, $zonesNoApilades, $zonesTotes,
                              [int]$nSenseCoord, [int]$nImpossibles) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Coordenades dels establiments'
    $form.Size = New-Object System.Drawing.Size(600, 610)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $true; $form.MaximizeBox = $false
    if ($null -ne $Script:EinaIcon) { $form.Icon = $Script:EinaIcon }

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
                "El mapa et mostra la de l'Excel (vermell) i la del PORTAL segons l'adreça (verd), que pots " +
                "arrossegar. Les JA CORREGIDES (lila) es mouen des d'on són, o les tornes a la parcel·la."
    $lbl.AutoSize = $false
    $lbl.Size = New-Object System.Drawing.Size(550, 56)
    $lbl.Location = New-Object System.Drawing.Point(15, 36)
    $form.Controls.Add($lbl)

    # QUINES ACTIVITATS (octubre 2026, l'usuari: "tambe vull poder moure els
    # punts de les activitats que no estan duplicades [...] que et deixi decidir
    # per un filtre abans d'obrir l'eina"). Abans hi havia nomes la casella
    # "Nomes les APILADES".
    $grupAbast = New-Object System.Windows.Forms.Panel
    $grupAbast.Size = New-Object System.Drawing.Size(550, 24)
    $grupAbast.Location = New-Object System.Drawing.Point(15, 94)
    $form.Controls.Add($grupAbast)
    $radios = @{}
    $xr = 0
    foreach ($op in @(@('apilades', 'Només les APILADES', 170), @('noapilades', 'Només les NO apilades', 180), @('totes', 'Totes', 120))) {
        $rb = New-Object System.Windows.Forms.RadioButton
        $rb.Text = $op[1]; $rb.Tag = $op[0]
        $rb.AutoSize = $false; $rb.Size = New-Object System.Drawing.Size($op[2], 22)
        $rb.Location = New-Object System.Drawing.Point($xr, 0)
        $grupAbast.Controls.Add($rb)
        $radios[$op[0]] = $rb
        $xr += $op[2]
    }
    $radios['apilades'].Checked = $true

    $chkCorr = New-Object System.Windows.Forms.CheckBox
    $chkCorr.Text = "Amaga les JA CORREGIDES (l'Excel ja no les té al punt del Cadastre; es mira en generar el mapa)"
    $chkCorr.AutoSize = $false
    $chkCorr.Size = New-Object System.Drawing.Size(550, 22)
    $chkCorr.Location = New-Object System.Drawing.Point(15, 118)
    $form.Controls.Add($chkCorr)

    $chkAvis = New-Object System.Windows.Forms.CheckBox
    # Sense el simbol: en una casella la Segoe UI el pinta com un quadrat (guard).
    $chkAvis.Text = 'Només les marcades PER REVISAR (de totes les zones)'
    $chkAvis.AutoSize = $false
    $chkAvis.Size = New-Object System.Drawing.Size(550, 22)
    $chkAvis.Location = New-Object System.Drawing.Point(15, 140)
    $form.Controls.Add($chkAvis)

    $lblZ = New-Object System.Windows.Forms.Label
    $lblZ.Text = 'Tria les zones que vols repassar (quadres de 400 m):'
    $lblZ.AutoSize = $false
    $lblZ.Size = New-Object System.Drawing.Size(395, 20)
    $lblZ.Location = New-Object System.Drawing.Point(15, 166)
    $form.Controls.Add($lblZ)

    # TRIAR-LES AL PLANOL (octubre 2026, l'usuari: "que te les deixi seleccionar
    # al planol"). El PowerShell no sap pintar un mapa: es fa el mapa amb TOTES
    # les zones (de l'abast triat) i la graella s'hi clica. Les que ja hagis
    # marcat a la llista hi surten triades.
    $btnPlanol = New-Object System.Windows.Forms.Button
    $btnPlanol.Text = 'Triar-les al plànol...'
    $btnPlanol.Size = New-Object System.Drawing.Size(150, 24)
    $btnPlanol.Location = New-Object System.Drawing.Point(415, 162)
    $btnPlanol.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Controls.Add($btnPlanol)
    $tipPl = New-Object System.Windows.Forms.ToolTip
    $tipPl.SetToolTip($btnPlanol, "Obre el mapa amb la graella de zones: clica les que vulguis repassar. La primera vegada ha de demanar al Cadastre totes les parcel·les (després ja queda desat).")

    $llista = New-Object System.Windows.Forms.CheckedListBox
    $llista.Size = New-Object System.Drawing.Size(550, 300)
    $llista.Location = New-Object System.Drawing.Point(15, 188)
    $llista.CheckOnClick = $true
    $llista.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($llista)

    $lblTotal = New-Object System.Windows.Forms.Label
    $lblTotal.AutoSize = $false
    $lblTotal.Size = New-Object System.Drawing.Size(320, 20)
    $lblTotal.Location = New-Object System.Drawing.Point(15, 496)
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
        $zones = if ($radios['apilades'].Checked) { @($zonesApilades) }
                 elseif ($radios['noapilades'].Checked) { @($zonesNoApilades) }
                 else { @($zonesTotes) }
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

    foreach ($rb in @($radios.Values)) { $rb.add_CheckedChanged($omple) }
    # Les marcades per revisar viuen al NAVEGADOR (el PowerShell no les veu): el
    # mapa es fa amb TOTES les zones i s'obre amb el filtre "Per revisar".
    $chkAvis.add_CheckedChanged({
        $on = $chkAvis.Checked
        $llista.Enabled = -not $on; $grupAbast.Enabled = -not $on; $chkCorr.Enabled = -not $on; $btnPlanol.Enabled = -not $on
        $lblTotal.Text = if ($on) { "Totes les zones: el mapa s'obrirà amb el filtre «Per revisar»" } else { $lblTotal.Text }
        if (-not $on) { & $refrescaTotal }
    }.GetNewClosure())

    $btnTot = New-Object System.Windows.Forms.Button
    $btnTot.Text = 'Marcar-ho tot'
    $btnTot.Size = New-Object System.Drawing.Size(120, 26)
    $btnTot.Location = New-Object System.Drawing.Point(340, 493)
    $btnTot.add_Click({
        for ($i = 0; $i -lt $llista.Items.Count; $i++) { $llista.SetItemChecked($i, $true) }
        & $refrescaTotal
    }.GetNewClosure())
    $form.Controls.Add($btnTot)

    $btnCap = New-Object System.Windows.Forms.Button
    $btnCap.Text = 'Desmarcar-ho tot'
    $btnCap.Size = New-Object System.Drawing.Size(120, 26)
    $btnCap.Location = New-Object System.Drawing.Point(445, 493)
    $btnCap.add_Click({
        for ($i = 0; $i -lt $llista.Items.Count; $i++) { $llista.SetItemChecked($i, $false) }
        & $refrescaTotal
    }.GetNewClosure())
    $form.Controls.Add($btnCap)

    $y = 522
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

    # "Excel per importar...": una copia de la base amb les coordenades que has
    # corregit al mapa, en vermell (vegeu CoordenadesImportar.ps1). Va aqui
    # perque es fa amb la MATEIXA base que s'acaba de llegir.
    $peu = _AddPeuBotons $form @(@{ Nom = 'Enrere'; Text = (_TxtEnrere); Resultat = 'Cancel'; Esc = $true }) @(
        @{ Nom = 'Importar'; Text = 'Excel per importar...'; Resultat = 'Yes' }
        @{ Nom = 'Ok'; Text = 'Generar mapa'; Estil = 'primari'; Resultat = 'OK'; Intro = $true }) $y
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($peu['Importar'], "Fa una còpia d'aquesta base de dades amb les coordenades que has corregit al mapa, en vermell, per a qui les hagi d'importar.")

    $form.ClientSize = New-Object System.Drawing.Size(580, ($y + 46))

    & $omple

    # Scroll vertical i ajust a la pantalla (vegeu suport/UiFinestra.ps1).
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })
    $res = $form.ShowDialog()
    if ($res -eq [System.Windows.Forms.DialogResult]::Yes) { return [pscustomobject]@{ Accio = 'importar' } }
    $alPlanol = ($res -eq [System.Windows.Forms.DialogResult]::Retry)
    if ($res -ne [System.Windows.Forms.DialogResult]::OK -and -not $alPlanol) { return $null }
    $noms = @()
    foreach ($i in $llista.CheckedIndices) { $noms += [string]$estat.Zones[$i].Nom }
    $abast = 'totes'
    foreach ($k in @($radios.Keys)) { if ($radios[$k].Checked) { $abast = $k } }
    return [pscustomobject]@{ NomsZones = @($noms); Abast = $abast; AmagaCorregides = [bool]$chkCorr.Checked
                              NomesAvis = ((-not $alPlanol) -and [bool]$chkAvis.Checked); AlPlanol = $alPlanol }
}

# El boto "Excel per importar..." de la finestra de Coordenades.
function Invoke-CoordExcelImportar($baseFile) {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Tria l'Excel del repàs que et vas baixar del mapa (Coordenades_....xlsx)"
    $dlg.Filter = 'Repàs de coordenades (*.xlsx)|*.xlsx'
    $baixades = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
    if (Test-Path -LiteralPath $baixades) { $dlg.InitialDirectory = $baixades }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $files = @(Read-RepasXlsx $dlg.FileName)
        $corr = Get-CorreccionsDelRepas $files $baseFile.Name
    } catch {
        Show-EinaInfo ("No s'ha pogut llegir el repàs:`n`n" + $_.Exception.Message) 'Coordenades' 'Warning'
        return
    }
    if ($corr.PerId.Count -eq 0) {
        $msg = "El repàs no porta cap coordenada canviada per a aquesta base de dades ($($baseFile.Name))."
        if ($corr.AltraBase -gt 0) { $msg += "`n`nTé $($corr.AltraBase) activitats d'una altra base: $(@($corr.Bases) -join ', ')." }
        if ($corr.SenseCanvi -gt 0) { $msg += "`n`n$($corr.SenseCanvi) validades sense moure: tenen la mateixa coordenada que la base." }
        if ($corr.PerRevisar -gt 0) { $msg += "`n`n$($corr.PerRevisar) marcades per revisar (no es toquen; les veus al mapa amb el filtre «Per revisar»)." }
        Show-EinaInfo $msg 'Coordenades' 'Information'
        return
    }

    if (-not (Test-Path -LiteralPath $CoordOutputDir)) { New-Item -ItemType Directory -Path $CoordOutputDir -Force | Out-Null }
    $outPath = Join-Path $CoordOutputDir (Get-NomExcelImportar $baseFile.Name (Get-Date))
    # Una COPIA exacta del fitxer de la base: mateix format, mateixes columnes.
    [System.IO.File]::Copy($baseFile.FullName, $outPath, $true)

    $espera = New-EinaProgres 1
    $espera.Label.Text = "Escrivint $($corr.PerId.Count) coordenades a la còpia de la base..."
    $espera.Bar.Style = 'Marquee'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $res = Set-CoordenadesALaBase (Get-Item -LiteralPath $outPath) $corr.PerId
    } catch {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
        try { Remove-Item -LiteralPath $outPath -Force -ErrorAction SilentlyContinue } catch { }
        Show-EinaInfo ("No s'ha pogut escriure l'Excel per importar:`n`n" + $_.Exception.Message) 'Coordenades' 'Error'
        return
    } finally {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
    }

    $msg  = "Fet. És una còpia de la base de dades amb les coordenades corregides EN VERMELL.`n`n"
    $msg += "Coordenades canviades:                 $(@($res.Aplicades).Count)`n"
    if ($corr.SenseCanvi -gt 0) { $msg += "Validades sense moure (igual que abans): $($corr.SenseCanvi)`n" }
    if ($corr.PerRevisar -gt 0) { $msg += "Marcades per revisar (només s'hi escriu la coordenada si també la vas validar): $($corr.PerRevisar)`n" }
    if (@($res.JaCanviades).Count -gt 0) {
        $msg += "`nNO TOCADES perquè a la base ja tenen una altra coordenada (potser ja`n"
        $msg += "s'havien corregit): GIA $(@($res.JaCanviades) -join ', ')`n"
    }
    if (@($res.NoTrobades).Count -gt 0) { $msg += "`nNo són a la base: GIA $(@($res.NoTrobades) -join ', ')`n" }
    if ($corr.AltraBase -gt 0) { $msg += "`nD'una altra base de dades, no aplicades: $($corr.AltraBase) ($(@($corr.Bases) -join ', '))`n" }
    $msg += "`nFitxer: $outPath"
    Show-EinaInfo $msg 'Coordenades'
    Start-Process -FilePath $outPath
}


# ============================================================================
# MAIN
# ============================================================================
function Invoke-CoordenadesMain {
    # 1. Localitzar l'Excel.
    $xls = Find-LatestRutaExcel
    if ($null -eq $xls) {
        Show-EinaInfo ("No s'ha trobat cap base de dades d'activitats.`n`n" +
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
    $espera = New-EinaProgres 1
    $espera.Label.Text = "Llegint $($xls.File.Name)..."
    $espera.Bar.Style = 'Marquee'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $lectura = Read-CoordenadesFromExcel $xls.File
    } catch {
        $espera.Form.Close()
        Show-EinaInfo "Error llegint l'Excel:`n$($_.Exception.Message)" 'Coordenades' 'Error'
        return
    } finally {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
    }

    $tots = @($lectura.Registres)
    if ($tots.Count -eq 0) {
        Show-EinaInfo "La fulla 'Estes' no te cap activitat amb coordenades." 'Coordenades' 'Warning'
        return
    }
    # Abans, sense cap apilada l'eina plegava; ara les NO apilades tambe es
    # poden repassar (l'usuari, octubre 2026).
    $apilats   = @(Get-RegistresPerAbast $tots 'apilades')
    $noApilats = @(Get-RegistresPerAbast $tots 'noapilades')

    # 3. Triar les ZONES. El repas de centenars d'activitats no es fa d'una
    # tirada: es va per quadres de 400 m, i cada tanda es la que caben en una
    # estona.
    $zonesApil   = @(Get-ZonesAmbActivitats $apilats)
    $zonesNoApil = @(Get-ZonesAmbActivitats $noApilats)
    $zonesTot    = @(Get-ZonesAmbActivitats $tots)
    $tria = Show-CoordenadesForm $dbLabel $zonesApil $zonesNoApil $zonesTot `
                                 ([int]$lectura.SenseCoord) (@($lectura.Impossibles).Count)
    if ($null -eq $tria) { return }
    if ($tria.Accio -eq 'importar') { Invoke-CoordExcelImportar $xls.File; return }
    if (-not $tria.NomesAvis -and -not $tria.AlPlanol -and @($tria.NomsZones).Count -eq 0) {
        Show-EinaInfo "No has triat cap zona." 'Coordenades' 'Warning'
        return
    }

    if ($tria.NomesAvis) {
        # Les marcades viuen al navegador: totes les activitats, i el mapa
        # s'obre amb el filtre "Per revisar".
        $triats = @($tots | Where-Object { Test-CoordPlausible ([double]$_.UtmX) ([double]$_.UtmY) })
        $abast = 'per revisar (totes les zones)'
    } elseif ($tria.AlPlanol) {
        # Totes les zones de l'abast: les tries al mapa (la graella).
        $triats = @(Get-RegistresPerAbast $tots $tria.Abast | Where-Object { Test-CoordPlausible ([double]$_.UtmX) ([double]$_.UtmY) })
        $txtAbast = @{ apilades = '(apilades)'; noapilades = '(no apilades)'; totes = '(totes)' }[$tria.Abast]
        $abast = "zones triades al plànol $txtAbast"
    } else {
        $base    = @(Get-RegistresPerAbast $tots $tria.Abast)
        $zonesOk = @{}
        foreach ($n in @($tria.NomsZones)) { $zonesOk[$n] = $true }
        $triats = @($base | Where-Object {
            (Test-CoordPlausible ([double]$_.UtmX) ([double]$_.UtmY)) -and
            $zonesOk.ContainsKey((Get-ZonaDeCoord ([double]$_.UtmX) ([double]$_.UtmY)))
        })
        $txtAbast = @{ apilades = '(apilades)'; noapilades = '(no apilades)'; totes = '(totes)' }[$tria.Abast]
        $abast = ("{0} {1}" -f (@($tria.NomsZones) -join ', '), $txtAbast)
    }
    if ($triats.Count -eq 0) {
        Show-EinaInfo "Les zones triades no tenen cap activitat." 'Coordenades' 'Warning'
        return
    }

    # 4. Portals del Cadastre i el PUNT de cada parcel.la (la mateixa memoria
    # cau que el Planol activitats: el que ja hagi demanat, aqui no es torna a
    # demanar), amb barra de progres i Cancel.lar.
    $refcats = @(Get-RefcatsAConsultar $triats)
    $fetP = _CoordAmbProgres $refcats 'Portals' { param($l, $p) Get-PortalsPerParcelles $l $p }
    $fetC = $null
    if (-not $fetP.Cancelat) { $fetC = _CoordAmbProgres $refcats 'Punt del Cadastre' { param($l, $p) Get-ParcelesCadastre $l $p } }
    if ($fetP.Cancelat -or ($null -ne $fetC -and $fetC.Cancelat)) {
        Show-EinaInfo ("S'ha cancel·lat. El que ja s'havia demanat queda desat, aixi que si ho " +
                        "tornes a provar continuarà des d'on era.") 'Coordenades' 'Information'
        return
    }
    $portalsPerRc = if ($fetP.Resultat -is [hashtable]) { $fetP.Resultat } else { @{} }
    $puntsCad = @{}
    if ($null -ne $fetC.Resultat -and $null -ne $fetC.Resultat.PSObject.Properties['Punts']) { $puntsCad = $fetC.Resultat.Punts }

    # 5. Muntar els punts del mapa.
    $items = @()
    $nAmagades = 0
    foreach ($r in $triats) {
        $rc = Get-RefcatParcel $r.Rc
        $portals = @()
        if ($rc -ne '' -and $portalsPerRc.ContainsKey($rc)) { $portals = @($portalsPerRc[$rc]) }
        $pc = if ($rc -ne '' -and $puntsCad.ContainsKey($rc)) { $puntsCad[$rc] } else { $null }
        $it = New-ItemCoordenades $r $portals $pc
        if ($tria.AmagaCorregides -and $it.Corregida) { $nAmagades++; continue }
        $items += $it
    }
    if ($items.Count -eq 0) {
        Show-EinaInfo "Totes les activitats triades ja estan corregides (no tenen el punt del Cadastre)." 'Coordenades' 'Information'
        return
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
    $html = Build-CoordenadesHtml $items $dbLabel $abast $xls.File.Name $portalsMapa $(if ($tria.NomesAvis) { 'avis' } else { 'tots' }) `
                                  ([bool]$tria.AlPlanol) @($tria.NomsZones)
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
    $msg += "Sense portal (es queden on eren): $nCad`n"
    $msg += "Ja corregides a l'Excel (lila): $(@($items | Where-Object { $_.Corregida }).Count)`n"
    if ($nAmagades -gt 0) { $msg += "Amagades perque ja estan corregides: $nAmagades`n" }
    $msg += "`n"
    if ($tria.AlPlanol) {
        $msg += "Al mapa, clica les ZONES que vulguis repassar (i 'Fet'). Les pots`n"
        $msg += "tornar a triar quan vulguis amb el boto 'Triar zones'.`n`n"
    }
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
    Show-EinaInfo $msg 'Coordenades'
}

if (-not $Script:CoordHeadless) {
    Invoke-CoordenadesMain
}
