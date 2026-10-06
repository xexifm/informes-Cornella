<#
  PlanolDades.ps1 - Les dades del "Planol activitats": quines parcel.les es
  pinten, de quin color i amb quins ID GIA.

  D'ON SURT CADA COSA
    - Excel d'ESTABLIMENTS (fulla "Establiments"): un establiment per fila, amb
      la referencia cadastral SENCERA (20 car.), local/pis/porta, "Local buit"
      i l'ID de l'activitat que hi ha. Una activitat pot tenir MES D'UN
      establiment (naus, locals contigus): per aixo el planol surt d'aqui i no
      de l'Excel d'activitats, que nomes en porta una referencia.
    - Excel d'ACTIVITATS (fulla "Estes"): el nom, l'activitat i si es
      PRECINTADA (el mateix camp que el planol public, Test-IsPrecintada).
    - Base d'informes (informes-db.json): l'estat del darrer informe fiable de
      cada activitat (estat_actual, el calcula "Actualitzar base").
    - Cadastre: la GEOMETRIA de cada parcel.la (INSPIRE wfsCP) i, quan l'Excel
      no diu local/planta/porta, la planta i porta de la UNITAT (Consulta_DNPRC).

  ELS COLORS (decidit amb l'usuari, octubre 2026)
    vermell  precintada (Excel) o darrer informe 'Precinte / Cessament'
    groc     darrer informe 'Requeriment' o 'Ampliacio termini'
    verd     'Favorable', 'FI Requeriment', 'FI Precinte / Cessament'
    blau     sense cap informe, 'Revisar', 'Altres', 'Sense efecte'...: no
             sabem si esta legalitzada
  Una parcel.la amb activitats en estats diferents agafa el PITJOR (vermell >
  groc > blau > verd); ho decideix el mapa, perque depen dels filtres.

  NOMES DEFINEIX FUNCIONS. ASCII pur. Les funcions pures es proven a
  tests/run-tests-planol.ps1.
#>

# ----------------------------------------------------------------------------
# Configuracio (sobreescriptible des de suport/config.ps1)
# ----------------------------------------------------------------------------
# Geometria d'una parcel.la (INSPIRE, parcel.les cadastrals). {0} = refcat de 14.
$PlanolParcelUrlTemplate = 'https://ovc.catastro.meh.es/INSPIRE/wfsCP.aspx?service=wfs&version=2.0.0&request=GetFeature&STOREDQUERIE_ID=GetParcel&refcat={0}&srsname=EPSG::25831'
# Dades d'una unitat (escala, planta, porta, us). {0} = refcat de 20.
$PlanolUnitatUrlTemplate = 'https://ovc.catastro.meh.es/ovcservweb/OVCSWLocalizacionRC/OVCCallejero.asmx/Consulta_DNPRC?Provincia=&Municipio=&RC={0}'
# Les parcel.les i les unitats canvien poc: un any. Sense resultat, 30 dies.
$PlanolCacheDies     = 365
$PlanolCacheDiesBuit = 30

# Ordre de gravetat dels colors (el pitjor guanya).
$Script:PlanolOrdreEstat = @{ verd = 0; blau = 1; groc = 2; vermell = 3 }

# ----------------------------------------------------------------------------
# REFERENCIES CADASTRALS
# ----------------------------------------------------------------------------
function _PlanolRcNeta($rc) {
    if ($null -eq $rc) { return '' }
    return ([string]$rc).Trim().ToUpperInvariant() -replace '\s', ''
}

# La parcel.la: els 14 primers caracters, si tenen la forma d'una refcat.
function Get-PlanolParcela($rc) {
    $r = _PlanolRcNeta $rc
    if ($r.Length -lt 14) { return '' }
    $p = $r.Substring(0, 14)
    if ($p -notmatch '^[0-9A-Z]{14}$') { return '' }
    return $p
}

# El numero d'UNITAT dins de la parcel.la (caracters 15-18: '0011'), o ''.
# Cada nau, local o pis en te un de propi al Cadastre.
function Get-PlanolUnitat($rc) {
    $r = _PlanolRcNeta $rc
    if ($r.Length -lt 18) { return '' }
    $u = $r.Substring(14, 4)
    if ($u -notmatch '^[0-9]{4}$') { return '' }
    return $u
}

# ----------------------------------------------------------------------------
# L'ESTAT (COLOR) D'UNA ACTIVITAT
# ----------------------------------------------------------------------------
function Get-EstatPlanol([bool]$precinte, [string]$estatInforme) {
    $n = _NormalitzaText $estatInforme
    if ($precinte -or $n -eq 'precinte / cessament') { return 'vermell' }
    if ($n -eq 'requeriment' -or $n -eq 'ampliacio termini') { return 'groc' }
    if ($n -eq 'favorable' -or $n -eq 'fi requeriment' -or $n -eq 'fi precinte / cessament') { return 'verd' }
    return 'blau'
}

# El pitjor de diversos estats (per a la llegenda i el resum; el mapa fa el
# mateix amb els que passen el filtre).
function Get-PitjorEstatPlanol($estats) {
    $pitjor = ''; $n = -1
    foreach ($e in @($estats)) {
        $k = [string]$e
        if ($Script:PlanolOrdreEstat.ContainsKey($k) -and $Script:PlanolOrdreEstat[$k] -gt $n) { $pitjor = $k; $n = $Script:PlanolOrdreEstat[$k] }
    }
    return $pitjor
}

# ----------------------------------------------------------------------------
# EL "SUB-ESTABLIMENT": local, planta, porta
# ----------------------------------------------------------------------------
function _PlanolValor($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) { if ([math]::Floor($v) -eq $v) { return [string][long]$v } return [string]$v }
    return (([string]$v).Trim() -replace '\s+', ' ')
}

# Una part amb la seva etiqueta, i els numerics sense zeros a l'esquerra (el
# Cadastre diu '02'). -NomesNumeric (el local): l'etiqueta nomes si el valor
# comenca per un numero; 'NAU 6' o 'EDIFICI BERLIN' ja diuen que son. El bloc,
# l'escala, la planta i la porta la porten sempre: un 'C' sol no diu res.
function _PlanolPart([string]$etiqueta, [string]$v, [switch]$NomesNumeric) {
    if ($v -eq '') { return '' }
    if ($v -match '^-?\d+$') { return ($etiqueta + ' ' + [string][int]$v) }
    if ($NomesNumeric -and $v -notmatch '^\d') { return $v }
    return ($etiqueta + ' ' + $v)
}

# Com es diu l'establiment dins de la parcel.la. Torna { Text; Font }:
#   Font 'gia'      el que posa l'Excel d'establiments (local, bloc, escala,
#                   planta, porta)
#   Font 'cadastre' l'Excel no diu res i el Cadastre si (la seva unitat)
#   Font 'unitat'   ningu no diu res: el numero d'unitat de la refcat
#   Font ''         res de res
# $unitat: la sortida de ConvertFrom-CatastroDnprcXml per a la seva refcat, o $null.
function Get-SubEstabliment($est, $unitat) {
    $parts = @()
    foreach ($p in @(
        (_PlanolPart 'Local' (_PlanolValor $est.Local) -NomesNumeric),
        (_PlanolPart 'Bl.'   (_PlanolValor $est.Bloc)),
        (_PlanolPart 'Esc.'  (_PlanolValor $est.Escala)),
        (_PlanolPart 'Pl.'   (_PlanolValor $est.Pis)),
        (_PlanolPart 'Pt.'   (_PlanolValor $est.Porta)))) {
        if ($p -ne '') { $parts += $p }
    }
    if ($parts.Count -gt 0) { return [pscustomobject]@{ Text = ($parts -join ' - '); Font = 'gia' } }

    if ($null -ne $unitat) {
        $pl = _PlanolValor $unitat.Planta
        $plText = if ($pl -match '^0+$') { 'Pl. baixa' } else { _PlanolPart 'Pl.' $pl }
        foreach ($p in @(
            (_PlanolPart 'Bl.'  (_PlanolValor $unitat.Bloc)),
            (_PlanolPart 'Esc.' (_PlanolValor $unitat.Escala)),
            $plText,
            (_PlanolPart 'Pt.'  (_PlanolValor $unitat.Porta)))) {
            if ($p -ne '') { $parts += $p }
        }
        if ($parts.Count -gt 0) { return [pscustomobject]@{ Text = ($parts -join ' - '); Font = 'cadastre' } }
    }

    $u = Get-PlanolUnitat $est.Rc
    if ($u -ne '') { return [pscustomobject]@{ Text = ('unitat ' + $u); Font = 'unitat' } }
    return [pscustomobject]@{ Text = ''; Font = '' }
}

# ----------------------------------------------------------------------------
# LECTURA DE LES FULLES (PURA: rep la matriu 1-based de Value2)
# ----------------------------------------------------------------------------
function _PlanolCel($data, [int]$r, [int]$c) {
    if ($c -lt 1) { return $null }
    return $data[$r, $c]
}

# La fulla "Establiments". Un registre per fila amb ID d'establiment.
function ConvertFrom-FullaEstabliments($data, [int]$rows, $headers) {
    $c = @{
        Id = Find-HeaderColumn $headers 'ID Establiment GIA'; Rc = Find-HeaderColumn $headers 'Ref. cadastral'
        X = Find-HeaderColumn $headers 'UTM X'; Y = Find-HeaderColumn $headers 'UTM Y'
        Via = Find-HeaderColumn $headers 'Emp. Tipus via'; Carrer = Find-HeaderColumn $headers 'Emp. Carrer'
        Num = Find-HeaderColumn $headers 'Emp. Numero'; Lletra = Find-HeaderColumn $headers 'Emp. Lletra'
        Bloc = Find-HeaderColumn $headers 'Emp. Bloc'; Local = Find-HeaderColumn $headers 'Emp. N Local'
        Escala = Find-HeaderColumn $headers 'Emp. Escala'; Pis = Find-HeaderColumn $headers 'Emp. Pis'
        Porta = Find-HeaderColumn $headers 'Emp. Porta'; Buit = Find-HeaderColumn $headers 'Local buit'
        Act = Find-HeaderColumn $headers 'ID Activitat'
    }
    if ($c.Rc -lt 1 -or $c.Act -lt 1) {
        throw "La fulla d'establiments no te les columnes 'Ref. cadastral' i 'ID Activitat'."
    }
    $out = @()
    for ($r = 2; $r -le $rows; $r++) {
        $id = Get-IdDeCella (_PlanolCel $data $r $c.Id)
        $rc = _PlanolRcNeta (_PlanolCel $data $r $c.Rc)
        if ($id -eq '' -and $rc -eq '') { continue }
        $buitTxt = _NormalitzaText (_PlanolValor (_PlanolCel $data $r $c.Buit))
        $out += [pscustomobject]@{
            IdEst       = $id
            Rc          = $rc
            IdActivitat = Get-IdDeCella (_PlanolCel $data $r $c.Act)
            Local  = _PlanolValor (_PlanolCel $data $r $c.Local)
            Bloc   = _PlanolValor (_PlanolCel $data $r $c.Bloc)
            Escala = _PlanolValor (_PlanolCel $data $r $c.Escala)
            Pis    = _PlanolValor (_PlanolCel $data $r $c.Pis)
            Porta  = _PlanolValor (_PlanolCel $data $r $c.Porta)
            Buit   = ($buitTxt -eq 'si' -or $buitTxt -eq 's')
            UtmX   = ConvertTo-UtmNumber (_PlanolCel $data $r $c.X)
            UtmY   = ConvertTo-UtmNumber (_PlanolCel $data $r $c.Y)
            # En brut, per buscar el portal del Cadastre (Get-PortalsExactes).
            Carrer = _PlanolValor (_PlanolCel $data $r $c.Carrer)
            Numero = _PlanolValor (_PlanolCel $data $r $c.Num)
            Adreca = Format-EmpAddress (_PlanolValor (_PlanolCel $data $r $c.Via)) (_PlanolValor (_PlanolCel $data $r $c.Carrer)) (_PlanolValor (_PlanolCel $data $r $c.Num)) (_PlanolValor (_PlanolCel $data $r $c.Lletra))
        }
    }
    return @($out)
}

# La fulla "Estes" (activitats): hashtable ID -> { Id; Nom; Activitat; Rc;
# Precinte; Adreca; UtmX; UtmY }.
function ConvertFrom-FullaActivitatsPlanol($data, [int]$rows, $headers) {
    $c = @{
        Id = Find-HeaderColumn $headers 'ID Activitat'; Rc = Find-HeaderColumn $headers 'Ref. cadastral'
        X = Find-HeaderColumn $headers 'UTM X'; Y = Find-HeaderColumn $headers 'UTM Y'
        Via = Find-HeaderColumn $headers 'Emp. Tipus via'; Carrer = Find-HeaderColumn $headers 'Emp. Carrer'
        Num = Find-HeaderColumn $headers 'Emp. Numero'; Lletra = Find-HeaderColumn $headers 'Emp. Lletra'
        Act = Find-HeaderColumn $headers 'Activitat principal'; Nom = Find-HeaderColumn $headers 'Nom comercial activitat'
        Classif = Find-HeaderColumn $headers 'Classificacio general annex'
    }
    if ($c.Id -lt 1) { throw "La fulla d'activitats no te la columna 'ID Activitat'." }
    $pairs = _FindCampInfoPairs $headers
    $colsCcae = @(Get-ColumnesCcae $headers)
    $out = @{}
    for ($r = 2; $r -le $rows; $r++) {
        $id = Get-IdDeCella (_PlanolCel $data $r $c.Id)
        if ($id -eq '') { continue }
        $prec = $false
        foreach ($p in $pairs) {
            if (Test-IsPrecintada (_PlanolValor (_PlanolCel $data $r $p.NomCol)) (_PlanolValor (_PlanolCel $data $r $p.ValorCol))) { $prec = $true; break }
        }
        $out[$id] = [pscustomobject]@{
            Id        = $id
            Nom       = _PlanolValor (_PlanolCel $data $r $c.Nom)
            Activitat = _PlanolValor (_PlanolCel $data $r $c.Act)
            Rc        = _PlanolRcNeta (_PlanolCel $data $r $c.Rc)
            Precinte  = $prec
            Adreca    = Format-EmpAddress (_PlanolValor (_PlanolCel $data $r $c.Via)) (_PlanolValor (_PlanolCel $data $r $c.Carrer)) (_PlanolValor (_PlanolCel $data $r $c.Num)) (_PlanolValor (_PlanolCel $data $r $c.Lletra))
            UtmX      = ConvertTo-UtmNumber (_PlanolCel $data $r $c.X)
            UtmY      = ConvertTo-UtmNumber (_PlanolCel $data $r $c.Y)
            Carrer    = _PlanolValor (_PlanolCel $data $r $c.Carrer)
            Numero    = _PlanolValor (_PlanolCel $data $r $c.Num)
            Turistic  = Test-CcaeAllotjamentTuristic @($colsCcae | ForEach-Object { _PlanolValor (_PlanolCel $data $r $_) })
            # "Classificacio general annex" (II, III, L18 Cert...): filtre del mapa.
            Classificacio = _PlanolValor (_PlanolCel $data $r $c.Classif)
        }
    }
    return $out
}

# ALLOTJAMENTS TURISTICS (octubre 2026, l'usuari: "en principi no m'interessen.
# Filtra-ho per CCAE 552/5520"). La columna es "CCAE Codi" de la fulla Estes
# (l'usuari). Si un dia canvia de nom: les que porten "CCAE" al nom, i si alguna
# diu "principal", nomes aquella (un CCAE secundari d'allotjament no fa que
# l'activitat ho sigui). Array PLA d'indexs (1-based); buit si no n'hi ha cap.
# PURA.
function Get-ColumnesCcae($headers) {
    $codi = Find-HeaderColumn $headers 'CCAE Codi'
    if ($codi -gt 0) { return @($codi) }
    $arr = @($headers); $totes = @(); $princ = @()
    for ($i = 0; $i -lt $arr.Count; $i++) {
        $h = _NormalitzaText $arr[$i]
        if ($h -notmatch 'ccae') { continue }
        $totes += ($i + 1)
        if ($h -match 'principal') { $princ += ($i + 1) }
    }
    if ($princ.Count -gt 0) { return $princ }
    return $totes
}

# Es un allotjament turistic? Algun dels valors comenca per 552 (el grup 55.2 de
# la CCAE-2009: 552, 5520, '55.20', '5520 - Allotjaments turistics...'). PURA.
function Test-CcaeAllotjamentTuristic($valors) {
    foreach ($v in @($valors)) {
        $t = ([string]$v).Trim()
        if ($t -eq '') { continue }
        $m = [regex]::Match($t, '^\D*(\d[\d.\s]*)')
        if (-not $m.Success) { continue }
        $digits = $m.Groups[1].Value -replace '[^0-9]', ''
        if ($digits.StartsWith('552')) { return $true }
    }
    return $false
}

# La base d'informes (ja llegida): hashtable ID GIA -> { Estat; NInformes }.
# Fa servir estat_actual, que el calculen "Actualitzar base" i "Editar base"
# (aquell proces te _EstatActualActivitat; aqui no es torna a calcular).
function ConvertFrom-InformesDbPlanol($db) {
    $out = @{}
    if ($null -eq $db -or $null -eq $db.PSObject.Properties['activitats']) { return $out }
    foreach ($a in @($db.activitats)) {
        if ($null -eq $a) { continue }
        $id = ([string]$a.id_gia).Trim()
        if ($id -eq '' -or $id -eq '-') { continue }
        $n = 0
        if ($null -ne $a.PSObject.Properties['informes'] -and $null -ne $a.informes) { $n = @($a.informes).Count }
        $out[$id] = [pscustomobject]@{ Estat = [string]$a.estat_actual; NInformes = $n }
    }
    return $out
}

# ----------------------------------------------------------------------------
# QUINES UNITATS S'HAN DE PREGUNTAR AL CADASTRE
# ----------------------------------------------------------------------------
# Nomes les que fan falta: refcat sencera (20), l'Excel no diu ni local ni
# planta ni porta, i la parcel.la te mes d'un establiment (si n'hi ha un de
# sol, no cal distingir-lo de res).
function Get-UnitatsAConsultar($establiments) {
    $perParcela = @{}
    foreach ($e in @($establiments)) {
        $p = Get-PlanolParcela $e.Rc
        if ($p -eq '') { continue }
        if ($perParcela.ContainsKey($p)) { $perParcela[$p]++ } else { $perParcela[$p] = 1 }
    }
    $set = @{}
    foreach ($e in @($establiments)) {
        $p = Get-PlanolParcela $e.Rc
        if ($p -eq '' -or $perParcela[$p] -lt 2) { continue }
        if ((_PlanolRcNeta $e.Rc).Length -ne 20) { continue }
        if (((Get-SubEstabliment $e $null).Font) -eq 'gia') { continue }
        $set[(_PlanolRcNeta $e.Rc)] = $true
    }
    return @(@($set.Keys) | Sort-Object)
}

# ----------------------------------------------------------------------------
# EL MODEL: parcel.les amb les seves activitats i locals buits
# ----------------------------------------------------------------------------
# Torna { Parceles; Resum }. Cada parcel.la: { Clau; Rc; X; Y; Entrades }.
# La clau es la refcat de 14; si un establiment no en te una de valida, es la
# coordenada ('xy:...') i al mapa surt com un punt.
# Cada entrada: { Tipus ('activitat'|'buit'); Gia; Nom; Activitat; Sub; SubFont;
#   Estat; EstatText; Precinte; MarcatBuit; SenseEstabliment; NoBase; NInformes;
#   Adreca; Rc }.
function Build-PlanolModel($establiments, $activitats, $estats, $unitats) {
    if ($null -eq $activitats) { $activitats = @{} }
    if ($null -eq $estats) { $estats = @{} }
    # Un establiment es un registre, mai una llista: si n'arriba una de
    # llistes (la crida l'havia embolcallat dues vegades, vegeu Planol.ps1), es
    # desplega. Si no, TOTES les activitats sortien "sense establiment".
    $establiments = @(@($establiments) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
    if ($null -eq $unitats) { $unitats = @{} }
    $parceles = [ordered]@{}
    $res = [ordered]@{ Establiments = 0; Activitats = 0; Buits = 0; SenseEstabliment = 0; NoBase = 0; MarcatsBuit = 0; SensePosicio = 0 }

    $afegeix = {
        param($rc, $x, $y, $entrada)
        $p = Get-PlanolParcela $rc
        $clau = $p
        if ($clau -eq '') {
            if ($null -eq $x -or $null -eq $y -or -not (Test-CoordPlausible ([double]$x) ([double]$y))) { $res.SensePosicio++; return }
            $clau = 'xy:' + [math]::Round([double]$x) + '_' + [math]::Round([double]$y)
        }
        if (-not $parceles.Contains($clau)) {
            $parceles[$clau] = [pscustomobject]@{ Clau = $clau; Rc = $p; X = $x; Y = $y; Entrades = (New-Object System.Collections.ArrayList) }
        }
        $pc = $parceles[$clau]
        if (($null -eq $pc.X -or $null -eq $pc.Y) -and $null -ne $x -and $null -ne $y) { $pc.X = $x; $pc.Y = $y }
        [void]$pc.Entrades.Add($entrada)
    }

    $giaAmbEstabliment = @{}
    foreach ($e in @($establiments)) {
        $res.Establiments++
        $unitat = $null
        $rcN = _PlanolRcNeta $e.Rc
        if ($unitats.ContainsKey($rcN)) { $unitat = $unitats[$rcN] }
        $sub = Get-SubEstabliment $e $unitat
        if ([string]$e.IdActivitat -eq '') {
            $res.Buits++
            & $afegeix $e.Rc $e.UtmX $e.UtmY ([pscustomobject]@{
                Tipus = 'buit'; Gia = ''; Nom = ''; Activitat = ''; Sub = $sub.Text; SubFont = $sub.Font
                Estat = ''; EstatText = ''; Precinte = $false; MarcatBuit = $true; SenseEstabliment = $false
                NoBase = $false; NInformes = 0; Adreca = $e.Adreca; Rc = $rcN
                Carrer = [string]$e.Carrer; Numero = [string]$e.Numero; Turistic = $false; Classificacio = '' })
            continue
        }
        $gia = [string]$e.IdActivitat
        $giaAmbEstabliment[$gia] = $true
        $act = $null
        if ($activitats.ContainsKey($gia)) { $act = $activitats[$gia] }
        $inf = $null
        if ($estats.ContainsKey($gia)) { $inf = $estats[$gia] }
        $estatText = if ($null -ne $inf) { [string]$inf.Estat } else { '' }
        $prec = ($null -ne $act -and [bool]$act.Precinte)
        if ($null -eq $act) { $res.NoBase++ }
        if ($e.Buit) { $res.MarcatsBuit++ }
        & $afegeix $e.Rc $e.UtmX $e.UtmY ([pscustomobject]@{
            Tipus = 'activitat'; Gia = $gia
            Nom = if ($null -ne $act) { $act.Nom } else { '' }
            Activitat = if ($null -ne $act) { $act.Activitat } else { '' }
            Sub = $sub.Text; SubFont = $sub.Font
            Estat = (Get-EstatPlanol $prec $estatText); EstatText = $estatText; Precinte = $prec
            MarcatBuit = [bool]$e.Buit; SenseEstabliment = $false; NoBase = ($null -eq $act)
            NInformes = if ($null -ne $inf) { [int]$inf.NInformes } else { 0 }
            Adreca = $e.Adreca; Rc = $rcN
            Carrer = [string]$e.Carrer; Numero = [string]$e.Numero
            Turistic = if ($null -ne $act) { [bool]$act.Turistic } else { $false }
            Classificacio = if ($null -ne $act) { [string]$act.Classificacio } else { '' } })
    }

    # Activitats que no surten a cap establiment: amb la refcat de l'Excel
    # d'activitats (una de sola), marcades perque es vegi.
    foreach ($gia in @($activitats.Keys | Sort-Object { [long]($_ -replace '\D', '0') })) {
        if ($giaAmbEstabliment.ContainsKey($gia)) { continue }
        $act = $activitats[$gia]
        $inf = $null
        if ($estats.ContainsKey($gia)) { $inf = $estats[$gia] }
        $estatText = if ($null -ne $inf) { [string]$inf.Estat } else { '' }
        $res.SenseEstabliment++
        & $afegeix $act.Rc $act.UtmX $act.UtmY ([pscustomobject]@{
            Tipus = 'activitat'; Gia = $gia; Nom = $act.Nom; Activitat = $act.Activitat
            Sub = ''; SubFont = ''; Estat = (Get-EstatPlanol ([bool]$act.Precinte) $estatText); EstatText = $estatText
            Precinte = [bool]$act.Precinte; MarcatBuit = $false; SenseEstabliment = $true; NoBase = $false
            NInformes = if ($null -ne $inf) { [int]$inf.NInformes } else { 0 }
            Adreca = $act.Adreca; Rc = $act.Rc
            Carrer = [string]$act.Carrer; Numero = [string]$act.Numero; Turistic = [bool]$act.Turistic
            Classificacio = [string]$act.Classificacio })
    }

    # Dins de cada parcel.la: primer les activitats (per sub-establiment i ID
    # GIA NUMERIC: 9 abans que 10), despres els locals buits.
    $gias = @{}
    foreach ($pc in $parceles.Values) {
        $ord = @($pc.Entrades | Sort-Object @{ Expression = { if ($_.Tipus -eq 'activitat') { 0 } else { 1 } } },
                                          @{ Expression = { $_.Sub } },
                                          @{ Expression = { $n = 0L; if ([long]::TryParse([string]$_.Gia, [ref]$n)) { $n } else { [long]::MaxValue } } })
        $pc.Entrades = $ord
        foreach ($en in $ord) { if ($en.Tipus -eq 'activitat') { $gias[$en.Gia] = $true } }
    }
    $res.Activitats = $gias.Count
    return [pscustomobject]@{ Parceles = @($parceles.Values); Resum = [pscustomobject]$res }
}

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

# ----------------------------------------------------------------------------
# CADASTRE: LA UNITAT (Consulta_DNPRC)
# ----------------------------------------------------------------------------
# { Escala; Planta; Porta; Bloc; Us; Superficie; Text } o $null (error del
# servei, o la resposta no diu res de la unitat).
function ConvertFrom-CatastroDnprcXml($xmlText) {
    if ([string]::IsNullOrWhiteSpace($xmlText)) { return $null }
    $doc = $null
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.LoadXml([string]$xmlText)
    } catch { return $null }
    if ($null -ne $doc.SelectSingleNode("//*[local-name()='lerr']//*[local-name()='err']")) { return $null }
    $t = {
        param($nom, $ctx)
        $n = $ctx.SelectSingleNode(".//*[local-name()='" + $nom + "']")
        if ($null -eq $n) { return '' }
        return ([string]$n.InnerText).Trim()
    }
    $loint = $doc.SelectSingleNode("//*[local-name()='loint']")
    $es = ''; $pt = ''; $pu = ''; $bq = ''
    if ($null -ne $loint) {
        $es = & $t 'es' $loint; $pt = & $t 'pt' $loint; $pu = & $t 'pu' $loint; $bq = & $t 'bq' $loint
    }
    $us = & $t 'luso' $doc.DocumentElement
    $sfc = & $t 'sfc' $doc.DocumentElement
    $ldt = & $t 'ldt' $doc.DocumentElement
    if ($es -eq '' -and $pt -eq '' -and $pu -eq '' -and $bq -eq '' -and $us -eq '' -and $ldt -eq '') { return $null }
    return [pscustomobject]@{ Escala = $es; Planta = $pt; Porta = $pu; Bloc = $bq; Us = $us; Superficie = $sfc; Text = $ldt }
}

# ----------------------------------------------------------------------------
# CADASTRE: LES CONSULTES (amb Cadastre.ps1: memoria cau, Cancel.lar)
# ----------------------------------------------------------------------------
# Muntades en el moment de fer-les: aixi val el que config.ps1 hagi posat.
function _PlanolConsultaParceles {
    return @{
        Fitxer = 'parceles.json'; Arrel = 'Parcelles'; Camp = 'Poligons'
        Dies = $PlanolCacheDies; DiesBuit = $PlanolCacheDiesBuit
        Url     = { param($rc) $PlanolParcelUrlTemplate -f $rc }
        Parseja = { param($t) return ,@(ConvertFrom-CatastroParcelXml $t) }
    }
}
function _PlanolConsultaUnitats {
    return @{
        Fitxer = 'unitats.json'; Arrel = 'Unitats'; Camp = 'Unitat'
        Dies = $PlanolCacheDies; DiesBuit = $PlanolCacheDiesBuit
        Url     = { param($rc) $PlanolUnitatUrlTemplate -f $rc }
        Parseja = { param($t) return (ConvertFrom-CatastroDnprcXml $t) }
    }
}

# Hashtable refcat14 -> array de poligons (buit si no n'hi ha o ha fallat).
function Get-GeometriesParceles($refcats, [scriptblock]$onProgress = $null) {
    $r = Get-AmbCacheCadastre $refcats (_PlanolConsultaParceles) $onProgress
    $out = @{}
    foreach ($k in @($r.Keys)) {
        $v = $r[$k]
        $out[$k] = if ($null -eq $v) { @() } else { @($v | Where-Object { $null -ne $_ }) }
    }
    return $out
}

# Hashtable refcat20 -> unitat (o $null).
function Get-UnitatsCadastre($refcats, [scriptblock]$onProgress = $null) {
    return (Get-AmbCacheCadastre $refcats (_PlanolConsultaUnitats) $onProgress)
}

# ----------------------------------------------------------------------------
# DEL MODEL A LES DADES DEL MAPA (graus i noms curts per al JSON)
# ----------------------------------------------------------------------------
function _PlanolAnellAGraus($anell) {
    $a = @($anell)
    $out = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i + 1 -lt $a.Count; $i += 2) {
        $ll = Convert-UtmToLatLon ([double]$a[$i]) ([double]$a[$i + 1]) 31 $true
        $out.Add([math]::Round($ll.Lat, 6)); $out.Add([math]::Round($ll.Lon, 6))
    }
    return ,($out.ToArray())
}

# ----------------------------------------------------------------------------
# L'ENTRADA DE CADA ESTABLIMENT I LES PARCEL.LES JUNTADES (octubre 2026)
# ----------------------------------------------------------------------------
# Els portals d'una parcel.la amb EXACTAMENT el numero de l'establiment (i el
# carrer, si algun portal el porta: el mateix criteri que Coordenades,
# Select-PortalFacana). Si no hi es exacte, cap: l'usuari ho va triar aixi ("el
# 73 no es el 71"; l'ID surt en vermell al centre). Array PLA. PURA.
function Get-PortalsExactes($portals, $carrer, $numero) {
    $arr = @(@($portals) | Where-Object { $null -ne $_ })
    $num = Get-NumeroPortal $numero
    if ($arr.Count -eq 0 -or $num -eq '') { return @() }
    $via = Get-ViaNormalitzada $carrer
    $cands = $arr
    if ($via -ne '') {
        $ambVia = @($arr | Where-Object { (Get-ViaNormalitzada $_.Via) -eq $via })
        if ($ambVia.Count -gt 0) { $cands = $ambVia }
    }
    return @($cands | Where-Object { [string]$_.Numero -eq $num })
}

# On va l'etiqueta d'un establiment: { X; Y; Dx; Dy } (Resolve-AncoraEntrada) o
# $null si no hi ha cap portal exacte que caigui dins (o prou a prop) dels
# poligons. Amb dos portals amb el mateix numero, el mes proper a la parcel.la.
function Resolve-EntradaEstabliment($portals, $carrer, $numero, $polys) {
    $millor = $null; $dm = $null
    foreach ($p in @(Get-PortalsExactes $portals $carrer $numero)) {
        $a = Resolve-AncoraEntrada ([double]$p.X) ([double]$p.Y) $polys
        if ($null -eq $a) { continue }
        $d = [math]::Sqrt(([double]$p.X - $a.X) * ([double]$p.X - $a.X) + ([double]$p.Y - $a.Y) * ([double]$p.Y - $a.Y))
        if ($null -eq $dm -or $d -lt $dm) { $dm = $d; $millor = $a }
    }
    return $millor
}

# Les activitats d'una parcel.la, com a clau: les parcel.les que es toquen i
# tenen EXACTAMENT les mateixes activitats es juntes (una sola forma i una
# sola etiqueta). Nomes les mateixes: si una porta una activitat mes, juntar-les
# hi barrejaria el color de l'altra. '' si no n'hi ha cap.
function _PlanolClauActivitats($pc) {
    $g = @(@($pc.Entrades) | Where-Object { $_.Tipus -eq 'activitat' -and [string]$_.Gia -ne '' } | ForEach-Object { [string]$_.Gia } | Sort-Object -Unique)
    return ($g -join ',')
}

# ELS GRUPS: les parcel.les del model, juntant les que es toquen i tenen les
# mateixes activitats. Cada grup: { Membres (parcel.les); Polys (UTM) }. Les que
# no es junten, un grup cada una (amb els seus poligons, o cap si no en te).
function Get-GrupsParceles($parceles, $geometries) {
    if ($null -eq $geometries) { $geometries = @{} }
    $llista = @($parceles)
    $polysDe = @{}
    for ($i = 0; $i -lt $llista.Count; $i++) {
        $pc = $llista[$i]
        $polysDe[$i] = if ($pc.Rc -ne '' -and $geometries.ContainsKey($pc.Rc)) { @($geometries[$pc.Rc]) } else { @() }
    }
    # Union-find entre les que tenen la mateixa clau d'activitats i es toquen.
    $pare = @{}; for ($i = 0; $i -lt $llista.Count; $i++) { $pare[$i] = $i }
    $arrel = { param($x) while ($pare[$x] -ne $x) { $pare[$x] = $pare[$pare[$x]]; $x = $pare[$x] }; return $x }
    $perClau = @{}
    for ($i = 0; $i -lt $llista.Count; $i++) {
        if (@($polysDe[$i]).Count -eq 0) { continue }
        $k = _PlanolClauActivitats $llista[$i]
        if ($k -eq '') { continue }
        if (-not $perClau.ContainsKey($k)) { $perClau[$k] = New-Object System.Collections.Generic.List[int] }
        $perClau[$k].Add($i)
    }
    foreach ($k in @($perClau.Keys)) {
        $ix = $perClau[$k]
        for ($a = 0; $a -lt $ix.Count; $a++) {
            for ($b = $a + 1; $b -lt $ix.Count; $b++) {
                $j = Join-Poligons (@($polysDe[$ix[$a]]) + @($polysDe[$ix[$b]]))
                if ($j.Toquen) { $ra = & $arrel $ix[$a]; $rb = & $arrel $ix[$b]; if ($ra -ne $rb) { $pare[$rb] = $ra } }
            }
        }
    }
    # Claus de TEXT: en un [ordered], una clau entera es pren per una POSICIO.
    $grups = [ordered]@{}
    for ($i = 0; $i -lt $llista.Count; $i++) {
        $r = [string](& $arrel $i)
        if (-not $grups.Contains($r)) { $grups[$r] = New-Object System.Collections.Generic.List[int] }
        $grups[$r].Add($i)
    }
    $out = @()
    foreach ($r in @($grups.Keys)) {
        $ix = $grups[$r]
        $membres = @($ix | ForEach-Object { $llista[$_] })
        $polys = @($ix | ForEach-Object { @($polysDe[$_]) })
        if ($ix.Count -gt 1) { $polys = @((Join-Poligons $polys).Polys) }
        $out += [pscustomobject]@{ Membres = $membres; Polys = $polys }
    }
    return $out
}

# LES ETIQUETES D'UN GRUP. Cada activitat a la seva entrada (si es troba) i, si
# no, en VERMELL al punt de sempre. Es junten a la mateixa etiqueta: les
# entrades a menys de $Script:PlanolMateixaPortaM (la mateixa porta) i les de la
# MATEIXA activitat a menys de $Script:PlanolMateixaActivitatM.
# $portals: hashtable refcat14 -> portals; sense la parcel.la (no s'han pogut
# demanar), l'etiqueta va al centre com sempre i NO en vermell: no se sap si
# l'entrada hi es o no.
# Torna { Etiquetes = [{ X; Y; Dir; Gias; Vermell; Centre }]; Entrada = hashtable
# entrada -> $true/$false (trobada) }.
function Get-EtiquetesGrup($grup, $portals, [double]$cx, [double]$cy) {
    $punts = New-Object System.Collections.Generic.List[object]
    $alCentre = New-Object System.Collections.Generic.List[string]
    $vermells = New-Object System.Collections.Generic.List[string]
    $trobada = @{}
    foreach ($pc in @($grup.Membres)) {
        foreach ($en in @($pc.Entrades)) {
            if ($en.Tipus -ne 'activitat') { continue }
            $g = [string]$en.Gia
            # Sense els portals de la parcel.la (no s'han pogut demanar): no se
            # sap on es l'entrada, va al centre com sempre i sense vermell.
            if ($null -eq $portals -or -not $portals.ContainsKey([string]$pc.Rc)) {
                if (-not $alCentre.Contains($g)) { $alCentre.Add($g) }
                continue
            }
            # Sense dibuix de la parcel.la (un punt): no hi ha "dins" on posar-la.
            if (@($grup.Polys).Count -eq 0) {
                $trobada[$en] = $false; if (-not $vermells.Contains($g)) { $vermells.Add($g) }
                continue
            }
            $a = Resolve-EntradaEstabliment $portals[[string]$pc.Rc] $en.Carrer $en.Numero $grup.Polys
            if ($null -eq $a) { $trobada[$en] = $false; if (-not $vermells.Contains($g)) { $vermells.Add($g) }; continue }
            $trobada[$en] = $true
            $punts.Add([pscustomobject]@{ X = $a.X; Y = $a.Y; Dx = $a.Dx; Dy = $a.Dy; Gia = $g })
        }
    }
    # Agrupar les entrades (union-find petit).
    $n = $punts.Count
    $pare = New-Object int[] $n; for ($i = 0; $i -lt $n; $i++) { $pare[$i] = $i }
    $arrel = { param($x) while ($pare[$x] -ne $x) { $x = $pare[$x] }; return $x }
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = $i + 1; $j -lt $n; $j++) {
            $d = [math]::Sqrt(($punts[$i].X - $punts[$j].X) * ($punts[$i].X - $punts[$j].X) + ($punts[$i].Y - $punts[$j].Y) * ($punts[$i].Y - $punts[$j].Y))
            $junta = ($d -le $Script:PlanolMateixaPortaM) -or ($punts[$i].Gia -eq $punts[$j].Gia -and $d -le $Script:PlanolMateixaActivitatM)
            if ($junta) { $ri = & $arrel $i; $rj = & $arrel $j; if ($ri -ne $rj) { $pare[$rj] = $ri } }
        }
    }
    $clusters = [ordered]@{}
    for ($i = 0; $i -lt $n; $i++) { $r = [string](& $arrel $i); if (-not $clusters.Contains($r)) { $clusters[$r] = New-Object System.Collections.Generic.List[object] }; $clusters[$r].Add($punts[$i]) }
    $etiq = @()
    foreach ($r in @($clusters.Keys)) {
        $m = $clusters[$r]
        $sx = 0.0; $sy = 0.0; $sdx = 0.0; $sdy = 0.0; $gias = New-Object System.Collections.Generic.List[string]
        foreach ($p in $m) { $sx += $p.X; $sy += $p.Y; $sdx += $p.Dx; $sdy += $p.Dy; if (-not $gias.Contains($p.Gia)) { $gias.Add($p.Gia) } }
        $x = $sx / $m.Count; $y = $sy / $m.Count
        # La mitjana de dues entrades d'una L pot caure fora: llavors, la primera.
        if (-not (Test-PuntDinsPoligons $x $y $grup.Polys)) { $x = $m[0].X; $y = $m[0].Y }
        $etiq += [pscustomobject]@{ X = $x; Y = $y; Dir = (Get-DireccioEtiqueta $sdx $sdy); Gias = @($gias); Vermell = $false; Centre = $false }
    }
    if ($vermells.Count -gt 0) { $etiq += [pscustomobject]@{ X = $cx; Y = $cy; Dir = 'c'; Gias = @($vermells); Vermell = $true; Centre = $true } }
    if ($alCentre.Count -gt 0) { $etiq += [pscustomobject]@{ X = $cx; Y = $cy; Dir = 'c'; Gias = @($alCentre); Vermell = $false; Centre = $true } }
    return [pscustomobject]@{ Etiquetes = $etiq; Entrada = $trobada }
}

# Una llista d'objectes per al JSON del mapa, UN PER GRUP (una parcel.la o
# diverses de juntades). Noms curts: el fitxer porta milers d'entrades.
#   k clau, rc parcel.la (la primera), rcs totes (si n'hi ha mes d'una), c [lat,
#   lon] del centre (un punt DINS), p poligons [[anell...]] (cada anell pla
#   [lat, lon, ...]), l etiquetes [{ c [lat, lon], d 'r'|'l'|'t'|'b'|'c', g [ID
#   GIA], v 1 si va en vermell (entrada no trobada) }], e entrades:
#   { t 'a'|'b', g gia, n nom, ac activitat, s sub, sf font del sub, e estat,
#     et text de l'estat, pr precinte, mb marcat buit, se sense establiment,
#     nb no es a la base, ni informes, ad adreca, rc refcat, at allotjament
#     turistic, x entrada: 1 trobada, 0 no trobada, -1 no se sap, cl
#     "Classificacio general annex" de l'Excel d'activitats }
# $portals (opcional): hashtable refcat14 -> portals (Get-PortalsPerParcelles).
function ConvertTo-PlanolDadesMapa($model, $geometries, $portals = $null) {
    if ($null -eq $geometries) { $geometries = @{} }
    $out = @()
    foreach ($gr in @(Get-GrupsParceles @($model.Parceles) $geometries)) {
        $polys = @($gr.Polys)
        $pc0 = @($gr.Membres)[0]
        $centre = $null
        if ($polys.Count -gt 0) { $ip = Get-PuntInterior $polys; if ($null -ne $ip) { $centre = @([double]$ip.X, [double]$ip.Y) } }
        if ($null -eq $centre -and $null -ne $pc0.X -and $null -ne $pc0.Y) { $centre = @([double]$pc0.X, [double]$pc0.Y) }
        if ($null -eq $centre) { continue }
        $pJson = New-Object System.Collections.ArrayList
        foreach ($poly in $polys) {
            $ag = New-Object System.Collections.ArrayList
            foreach ($an in @($poly.Anells)) { [void]$ag.Add((_PlanolAnellAGraus $an)) }
            if ($ag.Count -gt 0) { [void]$pJson.Add($ag.ToArray()) }
        }
        $et = Get-EtiquetesGrup $gr $portals $centre[0] $centre[1]
        $llc = Convert-UtmToLatLon ([double]$centre[0]) ([double]$centre[1]) 31 $true
        $ents = @()
        foreach ($pc in @($gr.Membres)) {
            foreach ($en in @($pc.Entrades)) {
                $x = -1
                if ($et.Entrada.ContainsKey($en)) { $x = if ($et.Entrada[$en]) { 1 } else { 0 } }
                $ents += [ordered]@{
                    t = if ($en.Tipus -eq 'buit') { 'b' } else { 'a' }
                    g = [string]$en.Gia; n = [string]$en.Nom; ac = [string]$en.Activitat
                    s = [string]$en.Sub; sf = [string]$en.SubFont
                    e = [string]$en.Estat; et = [string]$en.EstatText
                    pr = [bool]$en.Precinte; mb = [bool]$en.MarcatBuit; se = [bool]$en.SenseEstabliment; nb = [bool]$en.NoBase
                    ni = [int]$en.NInformes; ad = [string]$en.Adreca; rc = [string]$en.Rc
                    at = [bool]$en.Turistic; x = $x; cl = ([string]$en.Classificacio).Trim()
                }
            }
        }
        $etJson = @()
        foreach ($l in @($et.Etiquetes)) {
            $ll = Convert-UtmToLatLon ([double]$l.X) ([double]$l.Y) 31 $true
            $etJson += [ordered]@{ c = @([math]::Round($ll.Lat, 7), [math]::Round($ll.Lon, 7)); d = [string]$l.Dir; g = @($l.Gias); v = [int][bool]$l.Vermell }
        }
        $rcs = @(@($gr.Membres) | ForEach-Object { [string]$_.Rc } | Where-Object { $_ -ne '' })
        $o = [ordered]@{
            k = [string]$pc0.Clau; rc = [string]$pc0.Rc
            c = @([math]::Round($llc.Lat, 6), [math]::Round($llc.Lon, 6))
            p = $pJson.ToArray()
            l = $etJson
            e = $ents
        }
        if ($rcs.Count -gt 1) { $o['rcs'] = $rcs }
        $out += $o
    }
    return @($out)
}

# ----------------------------------------------------------------------------
# DIAGNOSTIC (doble clic a suport\rutes\Provar-Planol.bat)
# ----------------------------------------------------------------------------
# Fa UNA consulta real de cada servei i explica que n'ha entes. Desa SEMPRE les
# respostes senceres a local\geocodificacio\ : si el parseig falla, son l'unica
# cosa que permet arreglar-lo sense anar a les palpentes (les fixtures de les
# proves estan muntades a ma, perque des d'on es va escriure el codi el
# Cadastre estava bloquejat).
function Test-Planol([string]$refcat = '2295827DF2729E0011RQ') {
    $rc20 = _PlanolRcNeta $refcat
    $rc14 = Get-PlanolParcela $rc20
    if ($rc14 -eq '') { Write-Host "Referencia cadastral no valida: '$refcat'" -ForegroundColor Red; return }
    $dir = Split-Path -Parent (Get-CacheCadastrePath 'x')
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $utf8 = New-Object System.Text.UTF8Encoding($false)

    Write-Host "`n1. GEOMETRIA DE LA PARCEL.LA $rc14" -ForegroundColor Cyan
    $url = $PlanolParcelUrlTemplate -f $rc14
    Write-Host "URL: $url"
    $xml = Invoke-CadastreGet $url
    if ($null -eq $xml) {
        Write-Host "SENSE RESPOSTA: $Script:CadastreUltimError" -ForegroundColor Red
    } else {
        $f = Join-Path $dir ("resposta-parcela-$rc14.xml")
        [System.IO.File]::WriteAllText($f, $xml, $utf8)
        $crues = ([regex]::Matches($xml, '<[A-Za-z0-9]*:?posList[ >]')).Count
        $polys = @(ConvertFrom-CatastroParcelXml $xml)
        Write-Host ("Resposta: {0} caracters, {1} llistes de coordenades. Desada a {2}" -f $xml.Length, $crues, $f)
        Write-Host ("Poligons entesos: {0}" -f $polys.Count) -ForegroundColor $(if ($polys.Count -gt 0) { 'Green' } else { 'Red' })
        $i = 0
        foreach ($p in $polys) {
            $i++
            $ext = @($p.Anells[0])
            Write-Host ("  poligon {0}: {1} vertexs, {2} forats, primer vertex X={3} Y={4}, area {5} m2" -f $i, ($ext.Count / 2), (@($p.Anells).Count - 1), $ext[0], $ext[1], [math]::Round((Get-AreaAnell $ext)))
        }
    }

    Write-Host "`n2. PLANTA I PORTA DE LA UNITAT $rc20" -ForegroundColor Cyan
    if ($rc20.Length -ne 20) { Write-Host "(cal una referencia de 20 caracters per a la unitat)"; return }
    $url = $PlanolUnitatUrlTemplate -f $rc20
    Write-Host "URL: $url"
    $xml = Invoke-CadastreGet $url
    if ($null -eq $xml) { Write-Host "SENSE RESPOSTA: $Script:CadastreUltimError" -ForegroundColor Red; return }
    $f = Join-Path $dir ("resposta-unitat-$rc20.xml")
    [System.IO.File]::WriteAllText($f, $xml, $utf8)
    Write-Host ("Resposta: {0} caracters. Desada a {1}" -f $xml.Length, $f)
    $u = ConvertFrom-CatastroDnprcXml $xml
    if ($null -eq $u) {
        Write-Host "No n'he tret res (o el Cadastre diu que no existeix)." -ForegroundColor Red
    } else {
        Write-Host ("Escala '{0}'  planta '{1}'  porta '{2}'  bloc '{3}'  us '{4}'  {5} m2" -f $u.Escala, $u.Planta, $u.Porta, $u.Bloc, $u.Us, $u.Superficie) -ForegroundColor Green
        Write-Host ("Descripcio: {0}" -f $u.Text)
        $est = [pscustomobject]@{ Rc = $rc20; Local = ''; Bloc = ''; Escala = ''; Pis = ''; Porta = '' }
        Write-Host ("Al planol hi sortiria: {0}" -f (Get-SubEstabliment $est $u).Text)
    }
}
