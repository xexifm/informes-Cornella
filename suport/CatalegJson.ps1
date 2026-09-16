#requires -Version 5.1
<#
.SYNOPSIS
  Lectura dels ESTRUCTURALS en FORMAT ESTANDARD UNIC (JSON amb "runs").

.DESCRIPTION
  Tots els ESTRUCTURALS (catalegs REQ1/TERMINI, conclusions, plantilles ACT_EXTR)
  es guarden amb la MATEIXA estructura JSON, editable des del mateix programa:

    { "tipus","familia","intro":[<paragraf>],
      "nodes":[ {"tipus","titol","clau"?,"ajuda"?,"cos":[<paragraf>],"fills":[<node>]} ] }

    <paragraf> = { "runs":[ {"t","b","i"} ], "url": bool }
    <ajuda>    = { "norma","criteri","aplica","competencia","revisat" }

  El "tipus" de cada node (mateix vocabulari a tots els catalegs) + la 'familia'
  donen la semantica. Vocabulari:
    - seccio     : contenidor de primer nivell (seccio de cataleg; grup de
                   conclusions -titol=tipus d'informe-; seccio visual d'actextr).
    - subseccio  : sub-contenidor (cataleg) que AGRUPA els seus items (fills).
    - item       : unitat de contingut (deficiencia / conclusio / bloc actextr).
    - subitem    : sub-item amb pic (fill d'un item; ::CHILD:: a actextr).
    - text       : text/introduccio (::TEXT:: a actextr).
    - sempre     : (conclusions) frase que s'inclou sempre.
    - nota/etiqueta/capcalera/paragraf : (actextr) estils ::NOTE::/::LABEL::/
                   ::HEADER::/::CONC:: de l'informe favorable.
  A ACT_EXTR, 'clau' es la [[KEY]] funcional (Decret 112); mai es toca aqui.

  L''AJUDA' es la fitxa de criteri d'un requeriment: que diu la norma, quan s'ha
  d'exigir, a qui aplica i qui es competent per exigir-ho. NO ES TEXT DE
  L'INFORME i no ha de sortir-hi MAI: viu a part del 'cos' justament perque el
  motor de composicio nomes llegeix el cos. Es veu al Pas 3 (boto d'informacio
  de cada requeriment) i a la vista en Word del cataleg, en gris.

  Aquest modul llegeix el JSON i el converteix al MATEIX model en memoria que
  retornava el lector de .docx (ja esborrat) i Build-ActExtrBlocks, de manera que
  la resta del programa (Build-Document, wizard, ACT_EXTR...) NO canvia i la
  generacio es identica a la del .docx. El cos de cada paragraf s'"aplana" a la
  mateixa cadena amb marques (**negreta**, //cursiva//, [[URL]] ...) que ja
  entenen Type-RichText i _SplitTextAndUrls.

  Nomes defineix funcions (cap execucio en carregar-se): segur en headless.
#>

# Aplana una llista de "runs" a la cadena amb marques equivalent.
function _RunsToMarkup($runs) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in @($runs)) {
        $t = [string]$r.t
        if ($r.i) { $t = '//' + $t + '//' }
        if ($r.b) { $t = '**' + $t + '**' }
        [void]$sb.Append($t)
    }
    return $sb.ToString()
}

# Un paragraf de cos {runs,url} -> la BodyLine (cadena) que espera el motor.
function _JsonParaToBodyLine($p) {
    $line = _RunsToMarkup $p.runs
    if ($p.url) { return '[[URL]] ' + $line }
    return $line
}

# ----------------------------------------------------------------------------
# LA FITXA D'AJUDA d'un node
# ----------------------------------------------------------------------------
# Els CINC camps son fixos i sempre els mateixos. No es un text lliure a posta:
# la pregunta que ha de respondre la fitxa ("ho requereixo o no?") te sempre les
# mateixes quatre parts -que diu la norma, quin es el llindar, a quines
# activitats s'aplica i qui es competent per exigir-ho- i un text seguit les
# barreja. El cinque, 'revisat', es el que permet saber quines fitxes s'han
# quedat endarrerides quan canvia una norma.
$Script:AjudaCamps = @('norma', 'criteri', 'aplica', 'competencia', 'revisat')

# La fitxa d'ajuda d'un node del JSON -> objecte normalitzat, o $null si no n'hi
# ha. Torna $null tambe quan hi es pero es BUIDA: aixi qui la consulta nomes ha
# de mirar si es $null, i el Pas 3 no pinta el boto d'un requeriment que encara
# no te fitxa.
function Read-AjudaNode($node) {
    if ($null -eq $node) { return $null }
    $a = $node.ajuda
    if ($null -eq $a) { return $null }
    $vals = [ordered]@{}
    $teRes = $false
    foreach ($c in $Script:AjudaCamps) {
        $v = [string]$a.$c
        if (-not [string]::IsNullOrWhiteSpace($v)) { $teRes = $true } else { $v = '' }
        $vals[$c] = $v.Trim()
    }
    if (-not $teRes) { return $null }
    return [pscustomobject]@{
        Norma       = [string]$vals['norma']
        Criteri     = [string]$vals['criteri']
        Aplica      = [string]$vals['aplica']
        Competencia = [string]$vals['competencia']
        Revisat     = [string]$vals['revisat']
    }
}

# La fitxa d'ajuda -> les linies "Etiqueta: valor" que en veuen l'usuari (Pas 3)
# i la vista en Word. UNA SOLA funcio perque les dues pantalles no puguin
# ensenyar coses diferents de la mateixa fitxa, que es el defecte que ja va
# passar amb la vista dels catalegs i el document.
#
# PURA: no toca ni Word ni WinForms, o sigui que es prova en headless.
function Format-AjudaLinies($ajuda) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $ajuda) { return $out.ToArray() }
    $etiquetes = [ordered]@{
        Norma       = 'Norma'
        Criteri     = 'Criteri'
        Aplica      = "A qui s'aplica"
        Competencia = 'Competencia'
        Revisat     = 'Revisat'
    }
    foreach ($prop in $etiquetes.Keys) {
        $v = [string]$ajuda.$prop
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        [void]$out.Add(($etiquetes[$prop] + ': ' + $v))
    }
    return $out.ToArray()
}

# Llegeix el JSON d'un ESTRUCTURAL.
#
# Accepta una RUTA o un objecte JA PARSEJAT (el que torna ConvertFrom-Json).
# L'editor valida el que acaba de serialitzar cridant el lector: amb la ruta
# calia escriure'l a un temporal i tornar-lo a llegir del disc, i en un cataleg
# de mig mega aquell viatge d'anada i tornada es de segons a cada desat.
function _LoadEstructuralJson($json) {
    if ($json -is [string]) { return (Read-JsonFile $json) }
    return $json
}

# ----------------------------------------------------------------------------
# CATALEG -> model del motor (Sections=[{Title; Items=[pla]}], ...)
# ----------------------------------------------------------------------------
# Aplana un node del cataleg a la LLISTA PLANA d'items que espera el motor:
#   - subseccio -> un Item Kind=subsection i, tot seguit, els seus fills (items)
#                  a la MATEIXA llista, en ordre (reprodueix el model pla d'abans).
#   - text      -> Item Kind=intro.
#   - item      -> Item Kind=item, amb els seus subitems com a Children.
function _EmitCatalegItem($node, $items) {
    $tipus = [string]$node.tipus
    $body = New-Object System.Collections.ArrayList
    foreach ($p in @($node.cos)) { [void]$body.Add((_JsonParaToBodyLine $p)) }

    if ($tipus -eq 'subseccio') {
        [void]$items.Add([pscustomobject]@{
            Kind = 'subsection'; Short = [string]$node.titol
            BodyLines = $body; Children = (New-Object System.Collections.ArrayList)
            Ajuda = $null
        })
        foreach ($ch in @($node.fills)) { _EmitCatalegItem $ch $items }
    }
    elseif ($tipus -eq 'text') {
        [void]$items.Add([pscustomobject]@{
            Kind = 'intro'; Short = [string]$node.titol
            BodyLines = $body; Children = (New-Object System.Collections.ArrayList)
            Ajuda = $null
        })
    }
    else {   # item
        $children = New-Object System.Collections.ArrayList
        foreach ($ch in @($node.fills)) {
            $cbody = New-Object System.Collections.ArrayList
            foreach ($p in @($ch.cos)) { [void]$cbody.Add((_JsonParaToBodyLine $p)) }
            [void]$children.Add([pscustomobject]@{
                Kind = 'child'; Short = [string]$ch.titol
                BodyLines = $cbody; Children = (New-Object System.Collections.ArrayList)
                Ajuda = (Read-AjudaNode $ch)
            })
        }
        [void]$items.Add([pscustomobject]@{
            Kind = 'item'; Short = [string]$node.titol
            BodyLines = $body; Children = $children
            Ajuda = (Read-AjudaNode $node)
        })
    }
}

# Llegeix un cataleg i retorna el model que espera el motor.
function Read-CatalegJson($json) {
    $o = _LoadEstructuralJson $json
    $sections = New-Object System.Collections.ArrayList
    foreach ($sec in @($o.nodes)) {
        $items = New-Object System.Collections.ArrayList
        foreach ($node in @($sec.fills)) { _EmitCatalegItem $node $items }
        [void]$sections.Add([pscustomobject]@{ Title = [string]$sec.titol; Items = $items })
    }
    $fixed = New-Object System.Collections.ArrayList
    foreach ($p in @($o.intro)) { [void]$fixed.Add((_JsonParaToBodyLine $p)) }
    $introText = if ($fixed.Count -gt 0) { [string]$fixed[0] } else { '' }

    return [pscustomobject]@{
        IntroText      = $introText
        Sections       = $sections
        IsFixedBody    = ($sections.Count -eq 0)
        FixedBodyLines = $fixed.ToArray()
    }
}

# ----------------------------------------------------------------------------
# CONCLUSIONS -> model del motor (HeaderText; Selectable; Always)
# ----------------------------------------------------------------------------
function Read-ConclusionsJson($json, $reportType = $null) {
    $o = _LoadEstructuralJson $json
    $want = if ([string]::IsNullOrWhiteSpace($reportType)) { '' } else { _NormalitzaText $reportType }

    $selectable = New-Object System.Collections.ArrayList
    $always = New-Object System.Collections.ArrayList
    foreach ($n in @($o.nodes)) {
        $tp = [string]$n.tipus
        if ($tp -eq 'seccio') {
            $inGroup = ($want -eq '') -or ((_NormalitzaText $n.titol) -eq $want)
            if ($inGroup) {
                foreach ($c in @($n.fills)) {
                    [void]$selectable.Add([pscustomobject]@{
                        Title = [string]$c.titol; Body = (_JsonParaToBodyLine @($c.cos)[0])
                    })
                }
            }
        } elseif ($tp -eq 'sempre') {
            [void]$always.Add((_JsonParaToBodyLine @($n.cos)[0]))
        }
    }
    $header = if (@($o.intro).Count -gt 0) { _RunsToMarkup (@($o.intro)[0]).runs } else { '' }

    return [pscustomobject]@{
        HeaderText = $header
        Selectable = $selectable.ToArray()
        Always     = $always.ToArray()
    }
}

# ----------------------------------------------------------------------------
# ACT_EXTR -> registres de paragraf { Text; Style } per a Build-ActExtrBlocks
# ----------------------------------------------------------------------------
# El 'tipus' del node determina el marcador ::KIND:: que cal reconstruir a la
# capcalera h2 "[[clau]] ::KIND:: titol". Build-ActExtrBlocks nomes fa servir la
# clau i el ::KIND:: (l'etiqueta la ignora), aixi que la generacio queda identica.
function _ActExtrTipusToken([string]$tipus) {
    switch ($tipus) {
        'subitem'   { return '::CHILD::' }
        'text'      { return '::TEXT::' }
        'nota'      { return '::NOTE::' }
        'etiqueta'  { return '::LABEL::' }
        'capcalera' { return '::HEADER::' }
        'paragraf'  { return '::CONC::' }
        default     { return '' }   # item
    }
}

# Emet (recursivament, en profunditat) els records d'un node ACT_EXTR: la seva
# capcalera h2 reconstruida, el seu cos (normal/url) i despres els seus fills.
function _EmitActExtrRecords($node, $records) {
    $tok = _ActExtrTipusToken ([string]$node.tipus)
    $h2 = '[[' + [string]$node.clau + ']]'
    if ($tok) { $h2 += ' ' + $tok }
    if (-not [string]::IsNullOrEmpty([string]$node.titol)) { $h2 += ' ' + [string]$node.titol }
    [void]$records.Add(@{ Text = $h2; Style = 'h2' })
    foreach ($p in @($node.cos)) {
        $st = if ($p.url) { 'url' } else { 'normal' }
        [void]$records.Add(@{ Text = (_RunsToMarkup $p.runs); Style = $st })
    }
    foreach ($ch in @($node.fills)) { _EmitActExtrRecords $ch $records }
}

function Read-ActExtrRecordsJson($json) {
    $o = _LoadEstructuralJson $json
    $records = New-Object System.Collections.ArrayList
    foreach ($sec in @($o.nodes)) {
        [void]$records.Add(@{ Text = [string]$sec.titol; Style = 'h1' })
        foreach ($node in @($sec.fills)) { _EmitActExtrRecords $node $records }
    }
    return $records
}
