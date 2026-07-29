#requires -Version 5.1
<#
.SYNOPSIS
  Els camps que has d'omplir: [CAMP: nom], [OPCIO: nom | a | b] i el text ric.

.DESCRIPTION
  Detecta els marcadors dins dels textos triats, en munta el registre de camps,
  els substitueix al text final (Apply-Fields) i els dibuixa INCRUSTATS dins del
  paragraf a les pantalles de tria (_RenderRichInto), de manera que omples el
  camp alli on despres sortira. Tambe hi ha la interpretacio de **negreta** i
  //cursiva// i la separacio text/enllacos (_SplitTextAndUrls).

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Step 4 - Field placeholders
#   [CAMP: nom]                  -> camp de text lliure
#   [CAMP: nom (hint)]           -> camp de text amb ajuda
#   [OPCIO: nom | A | B | C]     -> desplegable; l'usuari tria A, B o C i el
#                                   text triat substitueix el placeholder
# ----------------------------------------------------------------------------
$Script:CampRegex  = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'
$Script:OpcioRegex = [regex]'\[OPCIO:\s*([^\]]+?)\s*\]'

# Analitza el contingut d'un [OPCIO: ...]: "nom | A | B" -> nom + opcions.
function _ParseOpcio($raw) {
    $segs = $raw -split '\|'
    $name = $segs[0].Trim()
    $opts = @()
    for ($i = 1; $i -lt $segs.Count; $i++) {
        $o = $segs[$i].Trim()
        if ($o -ne '') { $opts += $o }
    }
    return @{ Name = $name; Options = $opts }
}

# Detecta [CAMP: ...] i [OPCIO: ...] dins $allText i els afegeix a $fields
# (sense duplicar). Modifica $fields in-place.
function _AddFieldsFromText($fields, $allText) {
    foreach ($m in $Script:CampRegex.Matches($allText)) {
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $hint = ''
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) {
            $name = $raw.Substring(0, $parenIdx).Trim()
            $hint = $raw.Substring($parenIdx).Trim().TrimStart('(').TrimEnd(')')
        }
        if (-not $fields.Contains($name)) {
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='text'; Hint=$hint; Options=@(); Value='' }
        }
    }
    foreach ($m in $Script:OpcioRegex.Matches($allText)) {
        $parsed = _ParseOpcio $m.Groups[1].Value
        $name = $parsed.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $fields.Contains($name)) {
            $val = if ($parsed.Options.Count -gt 0) { [string]$parsed.Options[0] } else { '' }
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='choice'; Hint=''; Options=$parsed.Options; Value=$val }
        }
    }
}

function Get-FieldsFromSelection($selectedSections) {
    $fields = [ordered]@{}
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            $allText = ($it.BodyLines -join ' ')
            foreach ($ch in $it.Children) {
                $allText += ' ' + ($ch.BodyLines -join ' ')
            }
            _AddFieldsFromText $fields $allText
        }
    }
    return $fields
}

# Afegeix els camps detectats a les conclusions TRIADES i les SEMPRE al
# diccionari $fields existent. Aixi al Pas 4 surten alhora els del REQ1
# i els del CONCLUSIONS.
function Add-FieldsFromConclusions($fields, $selectedConcl, $alwaysConcl) {
    foreach ($c in $selectedConcl) {
        $body = if ($c -is [string]) { $c } else { [string]$c.Body }
        _AddFieldsFromText $fields $body
    }
    foreach ($a in $alwaysConcl) {
        _AddFieldsFromText $fields ([string]$a)
    }
}

function Apply-Fields($text, $fields) {
    # Primer els desplegables [OPCIO: nom | ...] i despres els [CAMP: ...].
    $out = $Script:OpcioRegex.Replace($text, {
        param($m)
        $name = (_ParseOpcio $m.Groups[1].Value).Name
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    $out = $Script:CampRegex.Replace($out, {
        param($m)
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) { $name = $raw.Substring(0, $parenIdx).Trim() }
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    return $out
}

# Extreu els valors dels camps en un hashtable simple per a la sessio.
function Get-FieldValuesForSession($fields) {
    $h = @{}
    foreach ($name in $fields.Keys) { $h[$name] = $fields[$name].Value }
    return $h
}

# ----------------------------------------------------------------------------
# Renderitzat "ric" amb camps inline (Pas 3 i Pas de conclusions)
# ----------------------------------------------------------------------------
# Treu els marcadors **negreta** i //cursiva// d'un text per mostrar-lo net a
# la pantalla (al .docx final SI s'apliquen via Type-RichText). Es "loose":
# elimina TOTS els ** i //, encara que un parell quedi partit per un
# [OPCIO:]/[CAMP:] (p.ex. "**ampliar el termini [OPCIO]**"). Aixo nomes afecta
# la PREVISUALITZACIO; el text original (amb marcadors) es el que es desa i
# s'emet al document.
function _StripMarkers([string]$t) {
    if ([string]::IsNullOrEmpty($t)) { return '' }
    $t = $t -replace '\*\*', ''
    $t = $t -replace '//', ''
    return $t
}

# Segmenta un text amb [OPCIO:]/[CAMP:] en trossos ORDENATS, per renderitzar-lo
# amb controls inline alla on toca. Retorna una llista de hashtables:
#   @{ Kind='text';  Text='...' }                  (marcadors ** // ja retirats)
#   @{ Kind='opcio'; Name='...'; Options=@(...) }
#   @{ Kind='camp';  Name='...'; Hint='...' }
# Es una funcio PURA (provable sense Word/WinForms).
function _SegmentRichText([string]$text) {
    $segments = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrEmpty($text)) { return $segments.ToArray() }
    # CAMP i OPCIO en una sola passada, en ordre d'aparicio.
    $rx = [regex]'\[OPCIO:\s*([^\]]+?)\s*\]|\[CAMP:\s*([^\]]+?)\s*\]'
    $pos = 0
    foreach ($m in $rx.Matches($text)) {
        if ($m.Index -gt $pos) {
            $plain = _StripMarkers $text.Substring($pos, $m.Index - $pos)
            if ($plain.Length -gt 0) { [void]$segments.Add(@{ Kind='text'; Text=$plain }) }
        }
        if ($m.Groups[1].Success) {
            $p = _ParseOpcio $m.Groups[1].Value
            [void]$segments.Add(@{ Kind='opcio'; Name=$p.Name; Options=$p.Options })
        } else {
            $raw = $m.Groups[2].Value.Trim(); $name = $raw; $hint = ''
            $pi = $raw.IndexOf('(')
            if ($pi -ge 0) { $name = $raw.Substring(0, $pi).Trim(); $hint = $raw.Substring($pi).Trim().TrimStart('(').TrimEnd(')') }
            [void]$segments.Add(@{ Kind='camp'; Name=$name; Hint=$hint })
        }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $text.Length) {
        $plain = _StripMarkers $text.Substring($pos)
        if ($plain.Length -gt 0) { [void]$segments.Add(@{ Kind='text'; Text=$plain }) }
    }
    return $segments.ToArray()
}

# Llegeix un valor precarregat (sessio anterior) per nom de camp. $preload pot
# ser un hashtable o un PSCustomObject. Retorna $null si no hi es.
function _GetPreloadValue($preload, $name) {
    if ($null -eq $preload) { return $null }
    if ($preload -is [System.Collections.IDictionary]) {
        if ($preload.Contains($name)) { return $preload[$name] }
        return $null
    }
    if ($preload.PSObject.Properties.Name -contains $name) { return $preload.$name }
    return $null
}

# El "registre" relaciona nom de camp -> llista de controls (per sincronitzar
# els duplicats: un mateix nom pot sortir a diversos llocs de la mateixa
# pantalla i tots han de mostrar el mateix valor).
function _NewFieldRegistry { return @{} }
function _RegisterFieldControl($registry, $name, $ctrl) {
    if (-not $registry.ContainsKey($name)) { $registry[$name] = New-Object System.Collections.ArrayList }
    [void]$registry[$name].Add($ctrl)
}

# Renderitza $text dins d'un FlowLayoutPanel ($flow) barrejant etiquetes de
# text (paraula a paraula, per poder fer salt de linia) i controls inline per
# als [OPCIO:]/[CAMP:]. Crea/actualitza les entrades a $fields (mateixa forma
# que _AddFieldsFromText) i registra els controls a $registry per sincronitzar.
function _RenderRichInto($flow, [string]$text, $fields, $preload, $registry) {
    $segs = _SegmentRichText $text
    foreach ($seg in $segs) {
        if ($seg.Kind -eq 'text') {
            foreach ($w in ($seg.Text -split '\s+')) {
                if ($w -eq '') { continue }
                $lbl = New-Object System.Windows.Forms.Label
                $lbl.AutoSize = $true
                $lbl.Text = $w
                $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
                [void]$flow.Controls.Add($lbl)
            }
            continue
        }

        $name = $seg.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        if ($seg.Kind -eq 'opcio') {
            if (-not $fields.Contains($name)) {
                $val = if ($seg.Options.Count -gt 0) { [string]$seg.Options[0] } else { '' }
                $pv = _GetPreloadValue $preload $name
                if ($null -ne $pv) { $val = [string]$pv }
                $fields[$name] = [pscustomobject]@{ Name=$name; Type='choice'; Hint=''; Options=$seg.Options; Value=$val }
            }
            $cb = New-Object System.Windows.Forms.ComboBox
            $cb.DropDownStyle = 'DropDownList'
            $cb.Margin = New-Object System.Windows.Forms.Padding(0, 1, 4, 0)
            foreach ($o in $seg.Options) { [void]$cb.Items.Add($o) }
            # Amplada segons l'opcio mes llarga (limitada), per llegir-la be.
            $maxLen = 0; foreach ($o in $seg.Options) { if ($o.Length -gt $maxLen) { $maxLen = $o.Length } }
            $cb.Width = [Math]::Min(520, [Math]::Max(90, ($maxLen * 7) + 30))
            $idx = $cb.Items.IndexOf([string]$fields[$name].Value)
            if ($idx -lt 0 -and $cb.Items.Count -gt 0) { $idx = 0 }
            if ($idx -ge 0) { $cb.SelectedIndex = $idx }
            $cb.Tag = $name
            _RegisterFieldControl $registry $name $cb
            $cb.add_SelectedIndexChanged({
                $v = if ($null -ne $cb.SelectedItem) { [string]$cb.SelectedItem } else { '' }
                $fields[$name].Value = $v
                foreach ($other in $registry[$name]) {
                    if ($other -ne $cb -and ($other -is [System.Windows.Forms.ComboBox]) -and ([string]$other.SelectedItem -ne $v)) {
                        $other.SelectedItem = $v
                    }
                }
            }.GetNewClosure())
            [void]$flow.Controls.Add($cb)
        } else {
            if (-not $fields.Contains($name)) {
                $val = ''
                $pv = _GetPreloadValue $preload $name
                if ($null -ne $pv) { $val = [string]$pv }
                $fields[$name] = [pscustomobject]@{ Name=$name; Type='text'; Hint=$seg.Hint; Options=@(); Value=$val }
            }
            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Margin = New-Object System.Windows.Forms.Padding(0, 1, 4, 0)
            $tb.Width = 150
            $tb.Text = [string]$fields[$name].Value
            if ($seg.Hint) { $tb.AccessibleDescription = $seg.Hint }
            $tb.Tag = $name
            _RegisterFieldControl $registry $name $tb
            $tb.add_TextChanged({
                $fields[$name].Value = $tb.Text
                foreach ($other in $registry[$name]) {
                    if ($other -ne $tb -and ($other -is [System.Windows.Forms.TextBox]) -and ($other.Text -ne $tb.Text)) {
                        $other.Text = $tb.Text
                    }
                }
            }.GetNewClosure())
            [void]$flow.Controls.Add($tb)
            if ($seg.Hint) {
                $hl = New-Object System.Windows.Forms.Label
                $hl.AutoSize = $true
                $hl.Text = "($($seg.Hint))"
                $hl.ForeColor = [System.Drawing.Color]::DimGray
                $hl.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
                [void]$flow.Controls.Add($hl)
            }
        }
    }
}

# Construeix el text "ric" d'un element (item/fill): nomes la part de TEXT de
# cada BodyLine (descartem els URLs, que aqui no s'editen), unit amb espais.
function _RichTextOfBodyLines($bodyLines) {
    $parts = New-Object System.Collections.ArrayList
    foreach ($ln in $bodyLines) {
        $sp = _SplitTextAndUrls $ln
        if (-not [string]::IsNullOrWhiteSpace($sp.Text)) { [void]$parts.Add($sp.Text) }
    }
    return ($parts -join ' ')
}

# Separa el text d'una linia dels URLs que pugui contenir. Retorna:
#   @{ Text = '<tot el que hi ha abans del primer URL>'; Urls = @(url1, url2...) }
#
# Hi ha dues fonts d'URLs reconegudes:
#   1. Prefix intern '[[URL]] ': l'ha posat el lector del cataleg quan el paragraf
#      del .docx te estil 'Cita' (manera explicita, recomanada al cataleg
#      modern). En aquest cas tota la linia es l'URL.
#   2. Deteccio per contingut: qualsevol token que comenci per 'http://' o
#      'https://' (retrocompatible amb cataleg vell).
function _SplitTextAndUrls($line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return @{ Text=''; Urls=@() } }
    # Cas 1: paragraf marcat com a enllac pel lector del cataleg.
    if ($line.StartsWith('[[URL]] ')) {
        $url = $line.Substring('[[URL]] '.Length).Trim()
        return @{ Text=''; Urls=@($url) }
    }
    # Cas 2: deteccio per contingut.
    $m = [regex]::Match($line, 'https?://')
    if (-not $m.Success) { return @{ Text = $line.Trim(); Urls=@() } }
    $text = $line.Substring(0, $m.Index).Trim()
    $rest = $line.Substring($m.Index)
    $urls = @()
    foreach ($tok in ($rest -split '\s+')) {
        if ($tok -match '^https?://') { $urls += $tok }
    }
    return @{ Text = $text; Urls = $urls }
}
