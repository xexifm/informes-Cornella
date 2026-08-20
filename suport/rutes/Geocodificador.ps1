<#
  Geocodificador.ps1 - Coordenades de FACANA (portal) per a cada establiment.

  EL PROBLEMA QUE RESOL
  ---------------------
  Les columnes "UTM X" / "UTM Y" de la fulla "Estes"/"Estes" venen del
  Cadastre, i el Cadastre georeferencia la PARCEL.LA, no el local. Per tant
  totes les subdivisions d'una mateixa referencia cadastral (els 20 caracters
  comparteixen els 14 primers) surten EXACTAMENT al mateix punt del mapa.

  Amb la base de dades del 2026-08-18: 1.380 activitats -> nomes 899 punts
  diferents; 711 files (52%) en comparteixen un amb alguna altra. De les 898
  parcel.les, 227 tenen mes d'una activitat, i son les que concentren aquelles
  711 files. El pitjor cas (4091106DF2749A, Ctra. de l'Hospitalet 147) en te
  19 d'apilades.

  QUE FA AQUEST MODUL
  -------------------
  Per a cada PARCEL.LA (no per a cada activitat) demana al Cadastre la llista
  de PORTALS: el servei INSPIRE d'Adreces (wfsAD), consulta desada
  'GetadByRefcat'. Cada portal es un punt d'ENTRADA a l'edifici -- o sigui, un
  punt de FACANA -- i ve amb el seu numero de carrer. Despres assigna a cada
  activitat el portal que li correspon pel seu "Emp. Numero".

  Aixo NO inventa posicions: son punts oficials del Cadastre. Alli on una
  parcel.la te diversos portals (el cas classic de l'illa amb entrades per dos
  carrers) cada activitat cau al SEU portal. Alli on totes comparteixen el
  mateix portal (locals apilats en un mateix edifici) no hi ha res a fer:
  geograficament SON al mateix lloc.

  Avantatges sobre geocodificar el text de l'adreca: una consulta per
  parcel.la en lloc d'una per activitat, i la resposta ja ve en EPSG:25831,
  el MATEIX sistema que l'Excel (cap reprojeccio, cap error de conversio).

  QUI EL FA SERVIR
  ----------------
  Nomes l'eina 'Coordenades' (Coordenades.ps1). Ruta.ps1 i Precintades.ps1 no
  el toquen: segueixen fent servir la coordenada original de l'Excel.

  SEGURETAT
  ---------
  - Nomes s'hi envia la referencia cadastral (cap nom, cap rao social).
  - Si el servei no respon, si la resposta no s'enten o si el punt obtingut
    cau MES LLUNY de $GeoDistanciaMaximaM metres del punt del Cadastre que ja
    teniem, es descarta i es fa servir el de sempre. Mai pot moure un
    marcador a un lloc absurd.
  - El resultat es desa a local\geocodificacio\portals.json: a la segona
    execucio ja no cal xarxa.

  MODE HEADLESS (proves): les funcions pures (parseig del XML, triar el
  portal, normalitzar numeros) no toquen ni xarxa ni disc i es proven a
  suport/tests/run-tests-coordenades.ps1.

  NOTA: aquest fitxer s'escriu en ASCII SENSE accents a proposit. El Windows
  PowerShell 5.1 llegeix els .ps1 sense BOM com a ANSI i corromp els literals
  accentuats (mateixa convencio que Precintades.ps1).
#>

# ----------------------------------------------------------------------------
# Configuracio (sobreescriptible des de suport/config.ps1)
# ----------------------------------------------------------------------------
# Coordenades.ps1 carrega aquest modul ABANS de config.ps1, aixi que tot el
# que hi ha en aquest bloc es pot ajustar des de config.ps1 com qualsevol
# altra opcio del programa.

# Plantilla de la URL del servei INSPIRE d'Adreces del Cadastre. Es una
# variable (i no una cadena enterrada al codi) expressament: si el Cadastre
# canvia el nom del parametre de la consulta desada, es pot arreglar des de
# config.ps1 sense tocar el programa. {0} = referencia cadastral de 14 car.
$GeoCatastroUrlTemplate = 'https://ovc.catastro.meh.es/INSPIRE/wfsAD.aspx?service=WFS&version=2.0.0&request=GetFeature&STOREDQUERIE_ID=GetadByRefcat&refcat={0}&srsname=EPSG::25831'

# Distancia maxima (metres) que acceptem entre el portal trobat i la
# coordenada de parcel.la que ja teniem. Per damunt d'aixo assumim que alguna
# cosa ha anat malament (resposta d'una altra parcel.la, eixos canviats...) i
# ens quedem amb la coordenada de sempre. Val mes un marcador imprecis que un
# marcador mentider.
$GeoDistanciaMaximaM = 250.0

# Segons d'espera per consulta i nombre d'intents.
$GeoTimeoutSec = 20
$GeoIntents    = 2

# Dies que val una entrada de la memoria cau. Els portals del Cadastre no es
# mouen, pero les parcel.les SENSE resultat es tornen a provar molt abans
# (potser el servei estava caigut): per aixo hi ha dos terminis.
$GeoCacheDies     = 365
$GeoCacheDiesBuit = 30

# ----------------------------------------------------------------------------
# FUNCIONS PURES (provables sense xarxa ni disc)
# ----------------------------------------------------------------------------

# Referencia cadastral de la PARCEL.LA: els 14 primers caracters de la
# referencia de 20 (els 6 ultims identifiquen el local i el control). Retorna
# '' si no hi ha prou caracters.
function Get-RefcatParcel($rc) {
    if ($null -eq $rc) { return '' }
    $t = ([string]$rc).Trim().ToUpperInvariant() -replace '[^A-Z0-9]', ''
    if ($t.Length -lt 14) { return '' }
    return $t.Substring(0, 14)
}

# Numero de portal a partir del camp "Emp. Numero" de l'Excel, que pot venir
# com '147', '147-179', '13 B', '1-3', 'S/N'... Ens quedem amb el PRIMER grup
# de digits. Retorna '' si no n'hi ha cap (S/N, buit...).
function Get-NumeroPortal($numero) {
    if ($null -eq $numero) { return '' }
    $t = ([string]$numero).Trim()
    if ($t -match '(\d+)') { return $Matches[1] }
    return ''
}

# Normalitza un nom de via per comparar-lo (sense accents, sense el tipus de
# via al davant, majuscules, espais collapsats). 'C/ Doctor Ferran' i
# 'CARRER DOCTOR FERRAN' han de coincidir.
function Get-ViaNormalitzada($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $t = $sb.ToString().ToUpperInvariant()
    $t = $t -replace '[^A-Z0-9 ]', ' '
    # Treu el tipus de via del davant: el Cadastre el porta enganxat al nom i
    # l'Excel el te en una columna a part (Emp. Tipus via).
    $t = $t -replace '^(CARRER|CALLE|AVINGUDA|AVENIDA|AVDA|PASSEIG|PASEO|PLACA|PLAZA|RAMBLA|RBLA|CARRETERA|CTRA|TRAVESSERA|CAMINO|CAMI|RONDA|POLIGON|PASSATGE|PASAJE|AV|PG|PS|PL|CR|RD|PJ|PI|CM|C)\s+', ''
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

# Llegeix la resposta XML del servei INSPIRE d'Adreces i en treu la llista de
# portals: array d'objectes { Numero; Via; X; Y }.
#
# El parseig es a proposit TOLERANT (per local-name(), sense lligar-se a cap
# espai de noms ni a cap nivell concret de l'arbre): aixi sobreviu a un canvi
# de versio de l'esquema INSPIRE (ad/3.0 -> ad/4.0) sense tocar codi.
# CONVENCIO DE RETORN (llegeix-ho abans de tocar-hi): aquesta funcio retorna un
# array PLA, sense la coma protectora, i qui la crida l'ha d'embolcallar amb
# @(). Es la combinacio contraria a la de _FindCampInfoPairs (que retorna
# ,@(...) i s'ha de consumir SENSE @()), i barrejar-les costa car: amb
# 'return ,@()' i '@(...)' al lloc de la crida, l'array queda embolcallat DUES
# vegades, .Count val 1 i $p.Numero fa enumeracio de membres i retorna
# 'System.Object[]'. Va passar de debo la primera vegada que es va provar
# contra el Cadastre de veritat.
function ConvertFrom-CatastroAdXml($xmlText) {
    if ([string]::IsNullOrWhiteSpace($xmlText)) { return @() }
    $doc = $null
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null       # mai resoldre entitats externes
        $doc.LoadXml([string]$xmlText)
    } catch { return @() }

    # Diccionari gml:id -> nom de via, muntat amb QUALSEVOL element que sigui
    # un nom de via (ThoroughfareName) i porti un <text> a dins.
    $vies = @{}
    foreach ($tn in $doc.SelectNodes("//*[local-name()='ThoroughfareName']")) {
        $id = ''
        foreach ($at in $tn.Attributes) { if ($at.LocalName -eq 'id') { $id = [string]$at.Value } }
        if ($id -eq '') { continue }
        $txt = $tn.SelectSingleNode(".//*[local-name()='text']")
        if ($null -ne $txt -and -not [string]::IsNullOrWhiteSpace($txt.InnerText)) {
            $vies[$id] = ([string]$txt.InnerText).Trim()
        }
    }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $any = [System.Globalization.NumberStyles]::Any
    $portals = @()
    foreach ($ad in $doc.SelectNodes("//*[local-name()='Address']")) {
        # Coordenades: el primer <pos> que hi hagi a dins ("X Y").
        $posNode = $ad.SelectSingleNode(".//*[local-name()='pos']")
        if ($null -eq $posNode) { continue }
        $parts = ([string]$posNode.InnerText).Trim() -split '\s+'
        if (@($parts).Count -lt 2) { continue }
        $x = 0.0; $y = 0.0
        if (-not [double]::TryParse($parts[0], $any, $inv, [ref]$x)) { continue }
        if (-not [double]::TryParse($parts[1], $any, $inv, [ref]$y)) { continue }
        # Guardia d'eixos: en UTM 31N l'est (~420.000) va molt per sota del
        # nord (~4.578.000). Si venen a l'inreves, els girem.
        if ($x -gt $y) { $tmp = $x; $x = $y; $y = $tmp }

        # Numero de portal: entre tots els elements 'designator', el que porta
        # text i NO te fills (el de dins de LocatorDesignator; el de fora es
        # nomes un embolcall amb el mateix nom).
        $num = ''
        foreach ($d in $ad.SelectNodes(".//*[local-name()='designator']")) {
            if ($null -ne $d.SelectSingleNode("*")) { continue }
            $v = ([string]$d.InnerText).Trim()
            if ($v -ne '') { $num = $v; break }
        }

        # Nom de via: el <component xlink:href="#ES.SDGC.TN.xxx"> ens hi porta.
        $via = ''
        foreach ($c in $ad.SelectNodes(".//*[local-name()='component']")) {
            foreach ($at in $c.Attributes) {
                if ($at.LocalName -ne 'href') { continue }
                $ref = ([string]$at.Value).TrimStart('#')
                if ($vies.ContainsKey($ref)) { $via = $vies[$ref] }
            }
            if ($via -ne '') { break }
        }

        $portals += [pscustomobject]@{
            Numero = (Get-NumeroPortal $num)
            Via    = $via
            X      = $x
            Y      = $y
        }
    }
    return @($portals)
}

# Tria el portal que correspon a una activitat.
#
#   $portals : sortida de ConvertFrom-CatastroAdXml
#   $carrer  : "Emp. Carrer" de l'Excel (pot ser buit)
#   $numero  : "Emp. Numero" de l'Excel (pot ser buit)
#
# Retorna un objecte { X; Y; Precisio } o $null si no hi ha cap portal usable.
# Precisio:
#   'facana'       -> el numero del portal es EXACTAMENT el de l'activitat
#   'facana-aprox' -> no hi havia aquell numero; s'agafa el mes proper de la
#                     mateixa parcel.la, prioritzant la mateixa BANDA del
#                     carrer (o sigui, la mateixa paritat del numero)
function Select-PortalFacana($portals, $carrer, $numero) {
    $arr = @($portals)
    if ($arr.Count -eq 0) { return $null }

    # 1. Si sabem el carrer i algun portal el porta, ens quedem amb aquests.
    $viaBuscada = Get-ViaNormalitzada $carrer
    $cands = $arr
    if ($viaBuscada -ne '') {
        $ambVia = @($arr | Where-Object { (Get-ViaNormalitzada $_.Via) -eq $viaBuscada })
        if ($ambVia.Count -gt 0) { $cands = $ambVia }
    }

    # 2. Numero exacte.
    $numBuscat = Get-NumeroPortal $numero
    if ($numBuscat -ne '') {
        $exacte = @($cands | Where-Object { $_.Numero -eq $numBuscat })
        if ($exacte.Count -gt 0) {
            return [pscustomobject]@{ X = $exacte[0].X; Y = $exacte[0].Y; Precisio = 'facana' }
        }
    }

    # 3. El mes proper. Mateixa paritat primer: als carrers els senars son a
    # una banda i els parells a l'altra, aixi que el 15 s'assembla molt mes al
    # 13 que al 14, encara que numericament siguin igual de lluny.
    $ambNum = @($cands | Where-Object { $_.Numero -ne '' })
    if ($ambNum.Count -eq 0) {
        # Cap portal numerat: nomes serveix si n'hi ha un de sol i sense dubte.
        if ($cands.Count -eq 1) {
            return [pscustomobject]@{ X = $cands[0].X; Y = $cands[0].Y; Precisio = 'facana-aprox' }
        }
        return $null
    }
    if ($numBuscat -eq '') {
        if ($ambNum.Count -eq 1) {
            return [pscustomobject]@{ X = $ambNum[0].X; Y = $ambNum[0].Y; Precisio = 'facana-aprox' }
        }
        return $null
    }

    $n = [int]$numBuscat
    $millor = $null; $millorClau = $null
    foreach ($p in $ambNum) {
        $pn = [int]$p.Numero
        $paritat = if (((($pn - $n) % 2) + 2) % 2 -eq 0) { 0 } else { 1 }
        $clau = ($paritat * 1000000) + [math]::Abs($pn - $n)
        if ($null -eq $millorClau -or $clau -lt $millorClau) { $millorClau = $clau; $millor = $p }
    }
    if ($null -eq $millor) { return $null }
    return [pscustomobject]@{ X = $millor.X; Y = $millor.Y; Precisio = 'facana-aprox' }
}

# Distancia en metres entre dues coordenades UTM del mateix fus. Plana, que a
# aquestes escales (metres dins d'un municipi) es exacta de sobres.
function Get-UtmDistanceM([double]$x1, [double]$y1, [double]$x2, [double]$y2) {
    $dx = $x1 - $x2
    $dy = $y1 - $y2
    return [math]::Sqrt(($dx * $dx) + ($dy * $dy))
}

# Decideix la coordenada FINAL d'una activitat: el portal si es fiable, i si
# no la de sempre. Concentra tota la politica de seguretat, i es PURA (els
# portals se li passen ja resolts).
#
# Retorna { X; Y; Precisio }, amb Precisio 'cadastre' quan es queda amb la
# coordenada original de la parcel.la.
function Resolve-CoordEstabliment($portals, $carrer, $numero, [double]$utmX, [double]$utmY) {
    $original = [pscustomobject]@{ X = $utmX; Y = $utmY; Precisio = 'cadastre' }
    $tria = Select-PortalFacana $portals $carrer $numero
    if ($null -eq $tria) { return $original }
    # Xarxa de seguretat: un portal a 3 km de la parcel.la NO es d'aquesta
    # parcel.la. Val mes un marcador imprecis que un marcador mentider.
    $d = Get-UtmDistanceM $tria.X $tria.Y $utmX $utmY
    if ($d -gt [double]$GeoDistanciaMaximaM) { return $original }
    return $tria
}

# Text curt per ensenyar a l'usuari (popup del mapa, columna de l'Excel).
function Get-PrecisioText([string]$precisio) {
    switch ($precisio) {
        'facana'       { return 'portal (facana)' }
        'facana-aprox' { return 'portal mes proper (facana)' }
        'manual'       { return 'moguda a ma' }
        default        { return 'centre de la parcel.la (cadastre)' }
    }
}

# ----------------------------------------------------------------------------
# MEMORIA CAU EN DISC
# ----------------------------------------------------------------------------
# Un sol JSON amb tots els portals que hem demanat mai:
#   { "Versio": 1, "Parcelles": { "<refcat14>": { "Data": "..."; "Portals": [...] } } }
# Viu a local\geocodificacio\ (dins del clone pero fora del repositori).

function Get-GeoCachePath {
    $dir = Get-LocalSubdir $RepoRoot 'Geocodificacio'
    return (Join-Path $dir 'portals.json')
}

# Diu si una entrada de la memoria cau encara val. PURA (la data d'ara se li
# passa) per poder-la provar sense esperar un any.
function Test-GeoCacheEntryValida($entry, [datetime]$ara) {
    if ($null -eq $entry) { return $false }
    $data = [datetime]::MinValue
    try {
        $data = [datetime]::Parse([string]$entry.Data, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch { return $false }
    $dies = ($ara - $data).TotalDays
    $nPortals = 0
    if ($null -ne $entry.Portals) { $nPortals = @($entry.Portals).Count }
    $limit = if ($nPortals -eq 0) { [double]$GeoCacheDiesBuit } else { [double]$GeoCacheDies }
    return ($dies -ge 0 -and $dies -le $limit)
}

function Import-GeoCache {
    $path = Get-GeoCachePath
    if (-not (Test-Path -LiteralPath $path)) { return @{} }
    try {
        $obj = (Get-Content -LiteralPath $path -Raw -Encoding UTF8) | ConvertFrom-Json
        $out = @{}
        if ($null -ne $obj -and $null -ne $obj.Parcelles) {
            foreach ($p in $obj.Parcelles.PSObject.Properties) { $out[$p.Name] = $p.Value }
        }
        return $out
    } catch {
        # Una memoria cau corrupta no ha de tombar l'eina: es descarta i es
        # torna a preguntar.
        return @{}
    }
}

function Export-GeoCache($cache) {
    $path = Get-GeoCachePath
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $parcelles = [ordered]@{}
    foreach ($k in @($cache.Keys | Sort-Object)) { $parcelles[$k] = $cache[$k] }
    $obj = [ordered]@{ Versio = 1; Parcelles = $parcelles }
    $json = ($obj | ConvertTo-Json -Depth 8)
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

# ----------------------------------------------------------------------------
# XARXA (l'unica part que no es prova en headless)
# ----------------------------------------------------------------------------

function Build-CatastroAdUrl([string]$refcat) {
    return ($GeoCatastroUrlTemplate -f $refcat)
}

# Demana els portals d'una parcel.la. Retorna el XML cru o $null. Deixa el
# motiu de la fallada a $Script:GeoUltimError (per al diagnostic).
function Invoke-CatastroAd([string]$refcat) {
    $Script:GeoUltimError = ''
    $url = Build-CatastroAdUrl $refcat
    for ($attempt = 1; $attempt -le [int]$GeoIntents; $attempt++) {
        try {
            $resp = Invoke-WebRequest -Uri $url -Method Get -TimeoutSec ([int]$GeoTimeoutSec) -UseBasicParsing
            return [string]$resp.Content
        } catch {
            $Script:GeoUltimError = "$($_.Exception.Message)  [$url]"
            if ($attempt -lt [int]$GeoIntents) { Start-Sleep -Milliseconds 700 }
        }
    }
    return $null
}

# Aconsegueix els portals de TOTES les parcel.les demanades, fent servir la
# memoria cau i consultant nomes les que falten. Retorna una hashtable
# refcat -> array de portals.
#
# $onProgress (opcional) es crida com  & $onProgress $fetes $total $refcat  i,
# si retorna $false, s'atura i es torna el que s'hagi aconseguit fins llavors
# (aixi el boto Cancel.lar de la barra de progres funciona de debo).
#
# NO llenca mai: si el servei falla, la parcel.la queda sense portals i cada
# activitat es quedara amb la seva coordenada de sempre.
function Get-PortalsPerParcelles($refcats, [scriptblock]$onProgress = $null) {
    $result = @{}
    $llista = @($refcats | Where-Object { $_ -ne '' -and $null -ne $_ } | Sort-Object -Unique)
    if ($llista.Count -eq 0) { return $result }

    $cache = Import-GeoCache
    $ara = Get-Date
    $nous = 0
    $fetes = 0
    $total = $llista.Count
    foreach ($rc in $llista) {
        $fetes++
        if ($null -ne $onProgress) {
            $seguim = & $onProgress $fetes $total $rc
            if ($seguim -eq $false) { break }
        }

        $entry = $null
        if ($cache.ContainsKey($rc)) { $entry = $cache[$rc] }
        if (Test-GeoCacheEntryValida $entry $ara) {
            $result[$rc] = @()
            if ($null -ne $entry.Portals) { $result[$rc] = @($entry.Portals) }
            continue
        }

        $xml = Invoke-CatastroAd $rc
        $portals = @()
        if ($null -ne $xml) { $portals = @(ConvertFrom-CatastroAdXml $xml) }
        $result[$rc] = $portals
        # Nomes desem al cache el que hem pogut PREGUNTAR. Si la crida ha
        # fallat (xml null) no hi escrivim res: aixi no ens quedem trenta dies
        # amb un buit causat per una caiguda de xarxa d'un moment.
        if ($null -ne $xml) {
            $cache[$rc] = [pscustomobject]@{
                Data    = $ara.ToString('yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
                Portals = $portals
            }
            $nous++
            # Desem cada 25 parcel.les noves: si l'usuari cancel.la o peta
            # l'ordinador a mitja tanda, no es perd tot el que s'ha demanat.
            if (($nous % 25) -eq 0) { try { Export-GeoCache $cache } catch { } }
        }
    }
    if ($nous -gt 0) {
        try { Export-GeoCache $cache } catch { }   # no poder desar-la no es motiu per fallar
    }
    return $result
}

# ----------------------------------------------------------------------------
# DIAGNOSTIC
# ----------------------------------------------------------------------------
# Fa UNA consulta real i explica que ha entes. Serveix per comprovar, des de
# l'ordinador de la feina, que el servei del Cadastre respon i que el parseig
# funciona (a l'entorn on es va escriure el codi el host estava bloquejat).
#
# COM ES CRIDA. La manera bona es fer DOBLE CLIC a:
#
#   suport\rutes\Provar-Cadastre.bat
#
# Si ja tens un PowerShell obert a l'arrel del clone, tambe val:
#
#   $env:COORDENADES_TEST=1; . .\suport\rutes\Coordenades.ps1; Test-Geocodificador '2295827DF2729E'
#
# El que NO funciona es embolcallar-ho en un altre powershell -Command "...":
# el shell de FORA expandeix el $env: abans de passar-ho i al de dins li arriba
# "=1; ...". I tampoc facis '. Ruta.ps1' a pel: executa la seva Main i obre el
# planificador de rutes.
function Test-Geocodificador([string]$refcat = '2295827DF2729E') {
    $rc = Get-RefcatParcel $refcat
    if ($rc -eq '') { Write-Host "Referencia cadastral no valida: '$refcat'" -ForegroundColor Red; return }
    $url = Build-CatastroAdUrl $rc
    Write-Host "URL: $url"
    $xml = Invoke-CatastroAd $rc
    if ($null -eq $xml) {
        Write-Host "SENSE RESPOSTA: $Script:GeoUltimError" -ForegroundColor Red
        return
    }

    # Desem SEMPRE la resposta sencera. Si el parseig falla, aquest fitxer es
    # l'unica cosa que permet arreglar-lo sense anar a les palpentes.
    $desat = ''
    try {
        $dir = Split-Path -Parent (Get-GeoCachePath)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $desat = Join-Path $dir ("resposta-$rc.xml")
        [System.IO.File]::WriteAllText($desat, $xml, (New-Object System.Text.UTF8Encoding($false)))
    } catch { $desat = '' }

    # Quantes adreces porta la resposta CRUA, comptades sense parsejar res. Aixi
    # es distingeix "el servei no ha tornat res" de "no n'he sabut treure res".
    $crues = ([regex]::Matches($xml, '<[A-Za-z0-9]*:?Address[ >]')).Count

    Write-Host ("Resposta rebuda: {0} caracters, amb {1} adreces." -f $xml.Length, $crues) -ForegroundColor Cyan
    if ($desat -ne '') { Write-Host "Resposta sencera desada a: $desat" }

    $portals = @(ConvertFrom-CatastroAdXml $xml)
    Write-Host ("`nPortals entesos: {0} (de {1})" -f $portals.Count, $crues) -ForegroundColor Cyan
    foreach ($p in $portals) {
        Write-Host ("  numero='{0}'  via='{1}'  X={2}  Y={3}" -f $p.Numero, $p.Via, $p.X, $p.Y)
    }

    if ($portals.Count -eq 0 -and $crues -gt 0) {
        Write-Host "`nEl servei ha respost pero no n'he sabut treure cap portal." -ForegroundColor Yellow
        Write-Host "Passa el fitxer desat mes amunt per arreglar el parseig." -ForegroundColor Yellow
    } elseif ($portals.Count -lt $crues) {
        Write-Host ("`nAvis: el servei ha tornat {0} adreces i nomes n'he entes {1}." -f $crues, $portals.Count) -ForegroundColor Yellow
        Write-Host "Pot ser normal (adreces sense coordenades), pero val la pena mirar-s'ho." -ForegroundColor Yellow
    } elseif ($portals.Count -gt 0) {
        Write-Host "`nTot correcte: el servei respon i el parseig l'enten." -ForegroundColor Green
    }
}
