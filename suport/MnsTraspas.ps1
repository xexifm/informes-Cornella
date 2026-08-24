#requires -Version 5.1
<#
.SYNOPSIS
  Dos informes curts de Llicencia: MODIFICACIO NO SUBSTANCIAL i TRASPAS.

.DESCRIPTION
  Van dins de "Llicencia (Annex II / LL Prov)" perque comparteixen capcalera
  (la de LLIC, amb la linia "Classificacio:") i son del mateix tramit, pero el
  document no s'assembla gens: no hi ha bloc ABANS ni DESPRES ni deficiencies
  de projecte, nomes tres o quatre paragrafs fixos.

  L'UNICA COSA QUE ES TRIA son ELS PUNTS DE REQ1 que s'hi adjunten, amb la
  MATEIXA pantalla de sempre (Select-Items) i pintats amb la MATEIXA funcio que
  els informes de REQ1 (_WriteCatalegBody). O sigui que el format es identic per
  construccio, i si es canvia un requeriment a REQ1 aqui canvia tambe.

  La tria decideix la frase:
    - amb punts -> "...amb les seguents observacions:" i, a sota, els punts
    - sense     -> "...sense mes observacions en relacio a aquest tramit."

  Abans hi havia una enumeracio buida d'observacions perque l'usuari hi
  escrigues a ma; ara alli hi van els requeriments del cataleg. L'unica llista
  de Word buida que queda es la de les MODIFICACIONS de la MNS, que si que
  s'escriuen a ma.

  EL TEXT NO ES AQUI: viu a ESTRUCTURALS\MNSTRAS.json (un sol cataleg per als
  dos informes, com va demanar l'usuari) i es pot editar des de l'editor de
  catalegs com tota la resta. Cada paragraf hi va com un node:

    tipus 'text' -> paragraf normal (Format-Body)
    tipus 'item' -> paragraf de llista de Word BUIT (Format-ListItem)

  ...i la CLAU diu quan hi entra:

    ''                    -> sempre
    'amb-observacions'    -> nomes si n'hi ha
    'sense-observacions'  -> nomes si no n'hi ha
    'llista-observacions' -> la llista de les observacions (nomes si n'hi ha)

  Al Word que va enviar l'usuari, aquestes dues variants anaven escrites en
  VERMELL, i tambe els titols "MODIFICACIO NO SUBSTANCIAL" i "TRASPAS". Aquell
  color era una marca SEVA per veure que havia de canviar a cada informe, no
  part del document: els titols no s'escriuen i el text va amb el format de
  sempre (Format.ps1).

.NOTES
  CONVENCIO ASCII: el codi no porta accents. El text accentuat que va a la
  pantalla es fa amb [char]0xNN; el de l'informe surt del JSON.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Ruta del cataleg dels dos informes.
function _MnsCatalegPath {
    return [string](Join-Path $EstructuralsDir 'MNSTRAS.json')
}

# Llegeix MNSTRAS.json. $null si no hi es (el programa ho ha de dir, no fer com
# si res).
function Read-MnsCataleg([string]$path = '') {
    if ([string]::IsNullOrWhiteSpace($path)) { $path = _MnsCatalegPath }
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

# Les dues fases que aporta aquest modul, amb el mateix esquema que _LlicFases.
function _MnsFases {
    return @(
        [pscustomobject]@{
            Clau = 'mns'
            Nom  = 'Modificaci' + [char]0x00F3 + ' NO Substancial'
            Sub  = 'S' + [char]0x2019 + 'informa favorablement una modificaci' + [char]0x00F3 + ' que no es substancial'
            Curt = 'LlicMNS'
        }
        [pscustomobject]@{
            Clau = 'traspas'
            Nom  = 'Trasp' + [char]0x00E0 + 's'
            Sub  = 'Canvi de nom del titular de l' + [char]0x2019 + 'activitat'
            Curt = 'LlicTraspas'
        }
    )
}

# Es una fase d'aquest modul? Funcio PURA.
function _MnsEsFase([string]$fase) {
    foreach ($f in @(_MnsFases)) { if ([string]$f.Clau -eq [string]$fase) { return $true } }
    return $false
}

# La seccio del cataleg d'una fase (la clau del node = la clau de la fase).
function _MnsSeccio($cat, [string]$fase) {
    if ($null -eq $cat) { return $null }
    foreach ($s in @($cat.nodes)) {
        if ([string]$s.clau -eq [string]$fase) { return $s }
    }
    return $null
}

# HI ENTRA, aquest node? Funcio PURA. La clau diu de quina variant es; qualsevol
# altra clau (o cap) vol dir que hi va sempre.
function _MnsNodeEntra([string]$clau, [bool]$ambObservacions) {
    switch (([string]$clau).Trim().ToLower()) {
        'amb-observacions'    { return $ambObservacions }
        'sense-observacions'  { return (-not $ambObservacions) }
        'llista-observacions' { return $ambObservacions }
    }
    return $true
}

# ELS PARAGRAFS d'un informe, en ordre i ja decidits. Funcio PURA: retorna
# @{ Tipus; Linies } amb Tipus 'text' (paragraf normal) o 'llista' (paragraf de
# llista de Word). Es el que escriu el document i tambe el que ensenya la vista
# en Word del cataleg: no hi pot haver dues versions del mateix.
function _MnsParagrafs($cat, [string]$fase, [bool]$ambObservacions) {
    $out = New-Object System.Collections.ArrayList
    $sec = _MnsSeccio $cat $fase
    if ($null -eq $sec) { return $out.ToArray() }
    foreach ($nd in @($sec.fills)) {
        if (-not (_MnsNodeEntra ([string]$nd.clau) $ambObservacions)) { continue }
        $tipus = if ([string]$nd.tipus -eq 'item') { 'llista' } else { 'text' }
        $linies = New-Object System.Collections.ArrayList
        foreach ($p in @($nd.cos)) { [void]$linies.Add((_JsonParaToBodyLine $p)) }
        [void]$out.Add(@{ Tipus = $tipus; Linies = $linies.ToArray() })
    }
    return $out.ToArray()
}

# QUINES CONCLUSIONS van a l'informe. Funcio PURA (rep les dues llistes de
# conclusions ja llegides del cataleg).
#
#   MNS       -> SEMPRE l'avis de l'article 59.1.d: val per a qualsevol
#                modificacio no substancial, hi hagi observacions o no.
#   amb punts -> i la conclusio de REQUERIMENT de REQ1, la mateixa que fan
#                servir els requeriments normals (no se'n fa cap copia).
#
# Si no en queda cap -Traspas sense punts-, l'informe no porta bloc de
# CONCLUSIONS: la conclusio ja es dins del text fix.
function _MnsTriaConclusions($selMns, $selReq1, [string]$fase, [bool]$ambObservacions) {
    $out = New-Object System.Collections.ArrayList
    if ([string]$fase -eq 'mns') {
        foreach ($x in @($selMns)) { [void]$out.Add($x) }
    }
    if ($ambObservacions) {
        foreach ($x in @(Build-ConclusionsFromTitles $selReq1 @('Requeriment'))) { [void]$out.Add($x) }
    }
    return $out.ToArray()
}

# ...i la versio que llegeix el cataleg (no es pura, per aixo va a part).
function _MnsConclusions([string]$fase, [bool]$ambObservacions) {
    $mns = $null; $req1 = $null
    try { $mns = Read-Conclusions $ConclusionsPath 'MNS' } catch { }
    try { $req1 = Read-Conclusions $ConclusionsPath 'REQ1' } catch { }
    $sMns = if ($null -ne $mns) { @($mns.Selectable) } else { @() }
    $sReq = if ($null -ne $req1) { @($req1.Selectable) } else { @() }
    return (_MnsTriaConclusions $sMns $sReq $fase $ambObservacions)
}

# Nom del fitxer de sortida (mateix patro que la resta: data al principi).
function _MnsNomFitxer([datetime]$data, [string]$fase, [string]$idGia) {
    $curt = 'LlicMNS'
    foreach ($f in @(_MnsFases)) { if ([string]$f.Clau -eq [string]$fase) { $curt = [string]$f.Curt } }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($data.ToString('yyyy-MM-dd'))
    [void]$parts.Add($curt)
    if (-not [string]::IsNullOrWhiteSpace($idGia)) { [void]$parts.Add('GIA ' + $idGia.Trim()) }
    $nom = ($parts -join '_')
    $nom = [regex]::Replace($nom, '[\\/:*?"<>|]', '-')
    return ($nom + '.docx')
}

# ----------------------------------------------------------------------------
# COMPOSICIO DEL DOCUMENT (Word COM)
# ----------------------------------------------------------------------------
function Build-MnsDocument($word, $model) {
    $header = $model.Header
    $baseName = _MnsNomFitxer (Get-Date) ([string]$model.Fase) ([string]$header['ID_GIA'])
    $cfg = $Script:ReportFormatConfig
    $fields = $model.Fields

    # LA MATEIXA CAPCALERA que la resta d'informes de Llicencia (porta la linia
    # "Classificacio:"), tal com va demanar l'usuari.
    return Write-InformeDocx $word $baseName 'LLIC' $header {
        param($sel)
        # HI HA PUNTS DE REQ1? Es l'unica cosa que decideix la frase i la conclusio.
        $seccions = @($model.Punts)
        $amb = ($seccions.Count -gt 0)

        foreach ($p in @(_MnsParagrafs $model.Cataleg ([string]$model.Fase) $amb)) {
            if ([string]$p.Tipus -eq 'llista') {
                # El paragraf de llista va BUIT: l'omple l'usuari al Word.
                Format-ListItem $sel ''
                continue
            }
            foreach ($l in @(Apply-FieldsToLines $p.Linies $fields)) {
                $pp = _SplitTextAndUrls ([string]$l)
                if (-not [string]::IsNullOrWhiteSpace($pp.Text)) { Format-Body $sel $pp.Text }
                foreach ($u in @($pp.Urls)) { Format-Url $sel $u }
            }
            if ($cfg.SpacerAfterItem) { Format-Spacer $sel }
        }

        # ELS PUNTS DE REQ1, amb la MATEIXA funcio que els informes de REQ1: el
        # format es identic per construccio i no n'hi ha cap copia.
        if ($amb) { _WriteCatalegBody $sel $cfg $seccions $fields '' }

        # CONCLUSIONS. El bloc nomes surt si te alguna linia; el tancament hi va
        # sempre, com a la resta d'informes (son els nodes 'sempre' del cataleg).
        $concl = @(_MnsConclusions ([string]$model.Fase) $amb)
        $cap = if ($concl.Count -gt 0) { 'CONCLUSIONS' } else { '' }
        _WriteConclusionsBlock $sel $cfg $cap $concl (Get-TextTancament) $fields
    }
}

# ----------------------------------------------------------------------------
# NO HI HA PANTALLA PROPIA
# ----------------------------------------------------------------------------
# Abans hi havia Select-MnsObservacions, que nomes preguntava "hi ha
# observacions?". Ara la pregunta es "quins punts de REQ1 hi adjuntes?", i
# aquella pantalla ja existeix: es Select-Items, la mateixa del pas Projecte de
# Llicencia i la de "Requeriment - Nou". No s'hi inventa res.
