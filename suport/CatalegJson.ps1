#requires -Version 5.1
<#
.SYNOPSIS
  Lectura dels ESTRUCTURALS en FORMAT ESTANDARD UNIC (JSON amb "runs").

.DESCRIPTION
  Tots els ESTRUCTURALS (catalegs REQ1/TERMINI, conclusions, plantilles ACT_EXTR)
  es guarden amb la MATEIXA estructura JSON, editable des del mateix programa:

    { "tipus","familia","intro":[<paragraf>],
      "nodes":[ {"nivell","marca","titol","cos":[<paragraf>],"fills":[<node>]} ] }

    <paragraf> = { "runs":[ {"t","b","i"} ], "url": bool }

  La 'familia' i la 'marca' de cada node donen la semantica (l'estructura es
  sempre la mateixa; nomes canvia com s'interpreta):
    - cataleg     : nivell1 marca=seccio; nivell2 marca=item|subseccio|intro;
                    nivell3 marca=fill (imbricat). 'intro' = cos fix (TERMINI).
    - conclusions : nivell1 marca=grup (titol=tipus) amb fills marca=conclusio;
                    nivell1 marca=sempre (cos). 'intro' = [capcalera].
    - actextr     : nivell1 marca=seccio; nivell2 marca=bloc (titol="[[KEY]] ...").

  Aquest modul llegeix el JSON i el converteix al MATEIX model en memoria que
  retornaven Parse-Cataleg, Read-Conclusions i Build-ActExtrBlocks, de manera que
  la resta del programa (Build-Document, wizard, ACT_EXTR...) no canvia. El cos de
  cada paragraf s'"aplana" a la mateixa cadena amb marques (**negreta**,
  //cursiva//, [[URL]] ...) que ja entenen Type-RichText i _SplitTextAndUrls. Aixi
  la generacio des de JSON es identica a la del .docx.

  Nomes defineix funcions (cap execucio en carregar-se): segur en headless.
#>

# Aplana una llista de "runs" a la cadena amb marques equivalent. Es l'invers de
# la conversio del convertidor .docx->json (regex de Type-RichText).
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

# Llegeix i cacheja el JSON d'un ESTRUCTURAL (per no tornar-lo a parsejar cada cop).
function _LoadEstructuralJson($jsonPath) {
    return (Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)
}

# Llegeix un cataleg (familia 'cataleg') i retorna el MATEIX objecte que
# Parse-Cataleg:
#   IntroText, Sections=[{Title; Items=[{Kind; Short; BodyLines; Children}]}],
#   IsFixedBody, FixedBodyLines.
# Nivells: nivell1 marca=seccio -> Section; nivell2 (item/subseccio/intro) -> Item
# (Kind item/subsection/intro); nivell3 marca=fill (imbricat dins l'item) -> Child.
function Read-CatalegJson($jsonPath) {
    $o = _LoadEstructuralJson $jsonPath

    $sections = New-Object System.Collections.ArrayList
    foreach ($sec in @($o.nodes)) {
        $items = New-Object System.Collections.ArrayList
        foreach ($it in @($sec.fills)) {
            $kind = switch ([string]$it.marca) {
                'subseccio' { 'subsection' }
                'intro'     { 'intro' }
                default     { 'item' }
            }
            $body = New-Object System.Collections.ArrayList
            foreach ($p in @($it.cos)) { [void]$body.Add((_JsonParaToBodyLine $p)) }

            $children = New-Object System.Collections.ArrayList
            foreach ($ch in @($it.fills)) {
                $cbody = New-Object System.Collections.ArrayList
                foreach ($p in @($ch.cos)) { [void]$cbody.Add((_JsonParaToBodyLine $p)) }
                [void]$children.Add([pscustomobject]@{
                    Kind = 'child'; Short = [string]$ch.titol
                    BodyLines = $cbody; Children = (New-Object System.Collections.ArrayList)
                })
            }
            [void]$items.Add([pscustomobject]@{
                Kind = $kind; Short = [string]$it.titol
                BodyLines = $body; Children = $children
            })
        }
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

# Llegeix les conclusions (familia 'conclusions') i retorna el MATEIX objecte que
# Read-Conclusions per al $reportType demanat: HeaderText, Selectable=[{Title;
# Body}], Always. Nivell1 marca=grup (titol=tipus) amb fills marca=conclusio;
# nivell1 marca=sempre (cos). La capcalera es el primer paragraf d''intro'.
function Read-ConclusionsJson($jsonPath, $reportType = $null) {
    $o = _LoadEstructuralJson $jsonPath
    $want = if ([string]::IsNullOrWhiteSpace($reportType)) { '' } else { _NormalizeText $reportType }

    $selectable = New-Object System.Collections.ArrayList
    $always = New-Object System.Collections.ArrayList
    foreach ($n in @($o.nodes)) {
        $marca = [string]$n.marca
        if ($marca -eq 'grup') {
            $inGroup = ($want -eq '') -or ((_NormalizeText $n.titol) -eq $want)
            if ($inGroup) {
                foreach ($c in @($n.fills)) {
                    [void]$selectable.Add([pscustomobject]@{
                        Title = [string]$c.titol; Body = (_JsonParaToBodyLine @($c.cos)[0])
                    })
                }
            }
        } elseif ($marca -eq 'sempre') {
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

# Recorre l'arbre d'un ESTRUCTURAL 'actextr' i reconstrueix la llista ORDENADA de
# registres de paragraf @{ Text; Style } (Style 'h1'|'h2'|'normal'|'url') que
# espera Build-ActExtrBlocks. Nivell1 -> h1 (titol de seccio); nivell2 -> h2
# (titol "[[KEY]] ..."); cada paragraf de 'cos' -> normal/url segons 'url'.
function Read-ActExtrRecordsJson($jsonPath) {
    $o = _LoadEstructuralJson $jsonPath
    $records = New-Object System.Collections.ArrayList
    foreach ($sec in @($o.nodes)) {
        [void]$records.Add(@{ Text = [string]$sec.titol; Style = 'h1' })
        foreach ($blk in @($sec.fills)) {
            [void]$records.Add(@{ Text = [string]$blk.titol; Style = 'h2' })
            foreach ($p in @($blk.cos)) {
                $st = if ($p.url) { 'url' } else { 'normal' }
                [void]$records.Add(@{ Text = (_RunsToMarkup $p.runs); Style = $st })
            }
        }
    }
    return $records
}
