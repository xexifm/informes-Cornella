#requires -Version 5.1
<#
.SYNOPSIS
  Lectura de catalegs/conclusions en format JSON (Opcio B: cos en "runs").

.DESCRIPTION
  Els catalegs (REQ1) i les conclusions (0 CONCLUSIONS) es poden guardar com a
  JSON estructurat en lloc del .docx. Aquest modul els llegeix i els converteix
  al MATEIX model en memoria que retornen Parse-Cataleg i Read-Conclusions, de
  manera que la resta del programa (Build-Document, wizard, etc.) no canvia: el
  cos de cada paragraf es "aplana" a la mateixa cadena amb marques
  (**negreta**, //cursiva//, [[URL]] ...) que ja entenen Type-RichText i
  _SplitTextAndUrls. Aixi la generacio des de JSON es identica a la del .docx.

  Format JSON (per paragraf de cos): { "runs": [ {"t","b","i"} ], "url": bool }.
  Un "run" es un fragment de text amb negreta (b) i/o cursiva (i). 'url'=true
  marca un paragraf d'enllac (estil Cita).

  Nomes defineix funcions (cap execucio en carregar-se): segur en headless i
  testejable sense Word.
#>

# Aplana una llista de "runs" a la cadena amb marques equivalent. Es l'invers de
# la conversio que fa el convertidor .docx->json (regex de Type-RichText).
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

# Llegeix un cataleg JSON i retorna el MATEIX objecte que Parse-Cataleg:
#   IntroText, Sections=[{Title; Items=[{Kind; Short; BodyLines; Children}]}],
#   IsFixedBody, FixedBodyLines.
function Read-CatalegJson($jsonPath) {
    $o = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $sections = New-Object System.Collections.ArrayList
    foreach ($sec in @($o.seccions)) {
        $items = New-Object System.Collections.ArrayList
        foreach ($it in @($sec.items)) {
            $kind = switch ([string]$it.nivell) {
                'subseccio' { 'subsection' }
                'intro'     { 'intro' }
                'fill'      { 'child' }
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

# Llegeix les conclusions JSON i retorna el MATEIX objecte que Read-Conclusions
# per al $reportType demanat: HeaderText, Selectable=[{Title; Body}], Always.
function Read-ConclusionsJson($jsonPath, $reportType = $null) {
    $o = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $want = if ([string]::IsNullOrWhiteSpace($reportType)) { '' } else { _NormalizeText $reportType }

    $selectable = New-Object System.Collections.ArrayList
    foreach ($g in @($o.grups)) {
        $inGroup = ($want -eq '') -or ((_NormalizeText $g.tipus) -eq $want)
        if (-not $inGroup) { continue }
        foreach ($c in @($g.conclusions)) {
            [void]$selectable.Add([pscustomobject]@{ Title = [string]$c.titol; Body = (_JsonParaToBodyLine $c.cos) })
        }
    }
    $always = New-Object System.Collections.ArrayList
    foreach ($p in @($o.sempre)) { [void]$always.Add((_JsonParaToBodyLine $p)) }

    return [pscustomobject]@{
        HeaderText = [string]$o.capcalera
        Selectable = $selectable.ToArray()
        Always     = $always.ToArray()
    }
}
