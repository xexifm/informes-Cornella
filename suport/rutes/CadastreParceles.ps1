#requires -Version 5.1
<#
  CadastreParceles.ps1 - LA PARCEL.LA DEL CADASTRE (INSPIRE wfsCP): el dibuix
  (poligons amb els seus patis) i el PUNT de la parcel.la (cp:referencePoint),
  amb la memoria cau parceles2.json.

  Per que es un fitxer a part (octubre 2026): ho feia nomes el Planol
  activitats, i l'eina Coordenades tambe necessita el punt del Cadastre de cada
  parcel.la (la linia de punts i saber quines coordenades ja s'han corregit).
  Les dues fan servir la MATEIXA memoria cau: el que ja ha demanat una, l'altra
  no ho torna a demanar.

  NOMES DEFINEIX FUNCIONS (i les variables de sota, que config.ps1 pot
  sobreescriure: es carrega abans de Ruta.ps1). Depen de Cadastre.ps1
  (Get-AmbCacheCadastre). ASCII pur.
#>

# Geometria d'una parcel.la (INSPIRE, parcel.les cadastrals). {0} = refcat de 14.
$PlanolParcelUrlTemplate = 'https://ovc.catastro.meh.es/INSPIRE/wfsCP.aspx?service=wfs&version=2.0.0&request=GetFeature&STOREDQUERIE_ID=GetParcel&refcat={0}&srsname=EPSG::25831'
# Les parcel.les (i les unitats del Planol) canvien poc: un any. Sense
# resultat, 30 dies.
$PlanolCacheDies     = 365
$PlanolCacheDiesBuit = 30

# ----------------------------------------------------------------------------
# CADASTRE: GEOMETRIA DE LA PARCEL.LA (INSPIRE wfsCP)
# ----------------------------------------------------------------------------
# Les coordenades d'un anell (exterior o interior): un array PLA de doubles
# [x1, y1, x2, y2, ...] en UTM 31N, a 2 decimals. Pla i no parelles a posta:
# un array d'arrays en PowerShell es desenrotlla a la minima.
function _GmlCoords($node) {
    $vals = New-Object System.Collections.Generic.List[double]
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $any = [System.Globalization.NumberStyles]::Float
    $llistes = @($node.SelectNodes(".//*[local-name()='posList']"))
    if ($llistes.Count -eq 0) { $llistes = @($node.SelectNodes(".//*[local-name()='pos']")) }
    foreach ($pl in $llistes) {
        $dim = 2
        foreach ($at in $pl.Attributes) { if ($at.LocalName -eq 'srsDimension') { [void][int]::TryParse([string]$at.Value, [ref]$dim) } }
        if ($dim -lt 2) { $dim = 2 }
        $nums = @(([string]$pl.InnerText).Trim() -split '\s+' | Where-Object { $_ -ne '' })
        for ($i = 0; $i + 1 -lt $nums.Count; $i += $dim) {
            $x = 0.0; $y = 0.0
            if (-not [double]::TryParse($nums[$i], $any, $inv, [ref]$x)) { continue }
            if (-not [double]::TryParse($nums[$i + 1], $any, $inv, [ref]$y)) { continue }
            $vals.Add([math]::Round($x, 2)); $vals.Add([math]::Round($y, 2))
        }
    }
    # Guardia d'eixos, com als portals: en UTM 31N l'est (~420.000) va molt per
    # sota del nord (~4.578.000). Si venen a l'inreves, es giren tots.
    if ($vals.Count -ge 2 -and $vals[0] -gt $vals[1]) {
        for ($i = 0; $i + 1 -lt $vals.Count; $i += 2) { $t = $vals[$i]; $vals[$i] = $vals[$i + 1]; $vals[$i + 1] = $t }
    }
    return ,($vals.ToArray())
}

# Els poligons d'una parcel.la: array PLA d'objectes { Anells } on Anells[0] es
# l'exterior i la resta, forats. Cada anell, un array pla de doubles. Es
# consumeix amb @() (la convencio de rutes/). Mai llanca: una resposta que no
# s'enten torna cap poligon i la parcel.la surt com un punt.
function ConvertFrom-CatastroParcelXml($xmlText) {
    if ([string]::IsNullOrWhiteSpace($xmlText)) { return @() }
    $doc = $null
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.LoadXml([string]$xmlText)
    } catch { return @() }
    $polys = @()
    foreach ($pn in @($doc.SelectNodes("//*[local-name()='PolygonPatch' or local-name()='Polygon']"))) {
        $anells = New-Object System.Collections.ArrayList
        $ext = $pn.SelectSingleNode("*[local-name()='exterior']")
        if ($null -eq $ext) { continue }
        $ce = _GmlCoords $ext
        if ($ce.Count -lt 6) { continue }
        [void]$anells.Add($ce)
        foreach ($int in @($pn.SelectNodes("*[local-name()='interior']"))) {
            $ci = _GmlCoords $int
            if ($ci.Count -ge 6) { [void]$anells.Add($ci) }
        }
        $polys += [pscustomobject]@{ Anells = $anells.ToArray() }
    }
    return @($polys)
}

# EL PUNT DE LA PARCEL.LA segons el Cadastre (cp:referencePoint, a la mateixa
# resposta que el dibuix): @(x, y) en UTM 31N, o $null. Es "la UTM de la
# parcel.la cadastral" de l'usuari; al planol, una linia de punts l'uneix amb
# l'etiqueta de l'ID (que va a la UTM de l'Excel d'activitats). Guardia d'eixos
# com als portals. PURA.
function Get-PuntReferenciaParcela($xmlText) {
    if ([string]::IsNullOrWhiteSpace($xmlText)) { return $null }
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.LoadXml([string]$xmlText)
    } catch { return $null }
    $pos = $doc.SelectSingleNode("//*[local-name()='referencePoint']//*[local-name()='pos']")
    if ($null -eq $pos) { return $null }
    $parts = ([string]$pos.InnerText).Trim() -split '\s+'
    if (@($parts).Count -lt 2) { return $null }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $x = 0.0; $y = 0.0
    if (-not [double]::TryParse($parts[0], [System.Globalization.NumberStyles]::Float, $inv, [ref]$x)) { return $null }
    if (-not [double]::TryParse($parts[1], [System.Globalization.NumberStyles]::Float, $inv, [ref]$y)) { return $null }
    if ($x -gt $y) { $t = $x; $x = $y; $y = $t }
    return @($x, $y)
}

# La parcel.la sencera de la resposta del wfsCP: { Poligons; Punt }, o $null si
# no n'hi ha res (aixi la memoria cau la torna a preguntar al cap de 30 dies,
# com abans). PURA.
function ConvertFrom-CatastroParcela($xmlText) {
    $polys = @(ConvertFrom-CatastroParcelXml $xmlText)
    $punt = Get-PuntReferenciaParcela $xmlText
    if ($polys.Count -eq 0 -and $null -eq $punt) { return $null }
    return [pscustomobject]@{ Poligons = $polys; Punt = $punt }
}

# ----------------------------------------------------------------------------
# LA CONSULTA (amb Cadastre.ps1: memoria cau, Cancel.lar)
# ----------------------------------------------------------------------------
# UN FITXER NOU (parceles2.json, octubre 2026): abans s'hi desaven nomes els
# poligons ('Poligons'). Si s'hi hagues afegit el punt amb un altre camp, les
# entrades velles de menys de 30 dies es donarien per bones SENSE geometria
# (Test-CacheCadastreValida les pren per buides) i les parcel.les sortirien com
# un punt. La primera vegada es tornen a demanar totes (uns minuts).
function _ConsultaParcelesCadastre {
    return @{
        Fitxer = 'parceles2.json'; Arrel = 'Parcelles'; Camp = 'Parcela'
        Dies = $PlanolCacheDies; DiesBuit = $PlanolCacheDiesBuit
        Url     = { param($rc) $PlanolParcelUrlTemplate -f $rc }
        Parseja = { param($t) return (ConvertFrom-CatastroParcela $t) }
    }
}

# Les parcel.les del Cadastre: { Geometries = refcat14 -> array de poligons
# (buit si no n'hi ha o ha fallat); Punts = refcat14 -> @(x, y) (nomes les que
# en tenen) }.
function Get-ParcelesCadastre($refcats, [scriptblock]$onProgress = $null) {
    $r = Get-AmbCacheCadastre $refcats (_ConsultaParcelesCadastre) $onProgress
    $geos = @{}; $punts = @{}
    foreach ($k in @($r.Keys)) {
        $v = $r[$k]
        $geos[$k] = @()
        if ($null -eq $v) { continue }
        if ($null -ne $v.PSObject.Properties['Poligons']) { $geos[$k] = @(@($v.Poligons) | Where-Object { $null -ne $_ }) }
        if ($null -ne $v.PSObject.Properties['Punt'] -and @($v.Punt).Count -eq 2) { $punts[$k] = @([double]@($v.Punt)[0], [double]@($v.Punt)[1]) }
    }
    return [pscustomobject]@{ Geometries = $geos; Punts = $punts }
}
