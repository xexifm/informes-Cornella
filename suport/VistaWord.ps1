#requires -Version 5.1
<#
.SYNOPSIS
  Genera una VISTA en Word (.docx) de cada cataleg, a partir del seu JSON.

.DESCRIPTION
  La FONT DE VERITAT del programa son els JSON d'ESTRUCTURALS. Els .docx ja no
  serveixen per generar res (l'unica excepcio es '0 CAPCALERA.docx', que SI que
  es una plantilla de veritat: la generacio en copia el fitxer i hi substitueix
  els <<PLACEHOLDERS>>).

  Aquest modul escriu, per a cada cataleg, un .docx amb TOT el contingut possible
  (tots els requeriments, totes les conclusions...) perque es pugui consultar
  comodament sense obrir el programa. Es una VISTA de sortida (el text tal com
  sortiria a l'informe) pero amb els TITOLS DE WORD posats, de manera que el
  panell de navegacio de Word mostri l'estructura:

    Titol 1  -> seccio del cataleg (o grup de conclusions = tipus d'informe)
    Titol 2  -> subseccio, o be l'item quan no hi ha subseccio
    Titol 3  -> item dins d'una subseccio
    Normal   -> el cos, amb la negreta/cursiva i els enllacos

  Es regenera automaticament en desar des de l'editor de catalegs i des de
  Actualitzar.bat, sobreescrivint el .docx antic del mateix nom.

  Les funcions de text son PURES (testejables en headless); nomes l'escriptura
  necessita Word (COM), i per tant nomes va a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables)
# ----------------------------------------------------------------------------

# Ruta del .docx de la vista d'un cataleg (mateixa carpeta i nom, .docx).
function _VistaWordPathFor([string]$jsonPath) {
    return [System.IO.Path]::ChangeExtension($jsonPath, '.docx')
}

# '0 CAPCALERA' es una plantilla de VERITAT: no se n'ha de generar mai cap vista
# (la sobreescriuriem i perdriem la carta amb l'escut i la taula).
function _VistaEsProtegit([string]$jsonPath) {
    $b = [System.IO.Path]::GetFileNameWithoutExtension([string]$jsonPath)
    return ([string]$b -like '0 CAPCALERA*')
}

# Parteix una linia de cos en trossos amb negreta/cursiva. Les marques son les
# mateixes que entenen Type-RichText i el motor: **negreta** i //cursiva//.
# Retorna un ARRAY PLA de @{ Text; Bold; Italic }. Funcio PURA.
function _VistaSegments([string]$line) {
    $out = New-Object System.Collections.ArrayList
    $s = [string]$line
    if ([string]::IsNullOrEmpty($s)) { return @() }
    $rx = [regex]'\*\*(.+?)\*\*|//(.+?)//'
    $pos = 0
    foreach ($m in $rx.Matches($s)) {
        if ($m.Index -gt $pos) {
            [void]$out.Add(@{ Text = $s.Substring($pos, $m.Index - $pos); Bold = $false; Italic = $false })
        }
        if ($m.Groups[1].Success) { [void]$out.Add(@{ Text = $m.Groups[1].Value; Bold = $true;  Italic = $false }) }
        else                      { [void]$out.Add(@{ Text = $m.Groups[2].Value; Bold = $false; Italic = $true  }) }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $s.Length) {
        [void]$out.Add(@{ Text = $s.Substring($pos); Bold = $false; Italic = $false })
    }
    return $out.ToArray()
}

# Cal tornar a generar la vista? Funcio PURA (mateixa forma que _PdfShouldConvert).
# NOMES es regenera si el JSON s'ha tocat despres del .docx: si es regenerava
# sempre, cada Actualitzar.bat faria un commit d'un .docx "nou" (Word hi posa
# dates internes) i el repositori s'ompliria de canvis inutils.
function _VistaCalRegenerar([bool]$docxExists, [datetime]$jsonUtc, [datetime]$docxUtc, [bool]$force) {
    if ($force) { return $true }
    if (-not $docxExists) { return $true }
    return ($jsonUtc -gt $docxUtc)
}

# D'una capcalera h2 d'ACT_EXTR ("[[CLAU]] ::TOKEN:: etiqueta") en treu una
# etiqueta llegible per al titol de la vista. Funcio PURA.
function _VistaActExtrTitol([string]$h2) {
    $t = [string]$h2
    $clau = ''
    $m = [regex]::Match($t, '^\s*\[\[([^\]]*)\]\]')
    if ($m.Success) { $clau = $m.Groups[1].Value; $t = $t.Substring($m.Index + $m.Length) }
    $t = [regex]::Replace($t, '::[A-Z]+::', '')
    $t = $t.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return $clau }
    if ([string]::IsNullOrWhiteSpace($clau)) { return $t }
    return ($t + '  [' + $clau + ']')
}

# ----------------------------------------------------------------------------
# ESCRIPTURA A WORD (COM) - nomes Windows
# ----------------------------------------------------------------------------
# Constants d'estil integrades (independents de l'idioma del Word instal·lat:
# per nom serien "Ttulo 1" en castella i "Heading 1" en angles).
$Script:WdNormal = -1
$Script:WdH1     = -2
$Script:WdH2     = -3
$Script:WdH3     = -4

# Escriu un paragraf amb l'estil indicat i el text ja segmentat (negreta/cursiva).
function _VistaEscriuPara($sel, [int]$style, [string]$text, [bool]$bullet = $false) {
    try { $sel.Style = $style } catch { }
    $sel.Font.Bold = $false
    $sel.Font.Italic = $false
    $prefix = if ($bullet) { [char]0x2022 + ' ' } else { '' }
    if ($prefix) { $sel.TypeText($prefix) }
    foreach ($seg in @(_VistaSegments $text)) {
        $sel.Font.Bold = [bool]$seg.Bold
        $sel.Font.Italic = [bool]$seg.Italic
        $sel.TypeText([string]$seg.Text)
    }
    $sel.Font.Bold = $false
    $sel.Font.Italic = $false
    $sel.TypeParagraph()
}

# Escriu les linies de cos d'un item (les que venen del lector: poden dur el
# prefix '[[URL]] ' quan la linia es un enllac).
function _VistaEscriuCos($sel, $bodyLines, [bool]$bullet = $false) {
    foreach ($ln in @($bodyLines)) {
        $l = [string]$ln
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        if ($l.StartsWith('[[URL]] ')) {
            $url = $l.Substring('[[URL]] '.Length).Trim()
            try { $sel.Style = $Script:WdNormal } catch { }
            $sel.Font.Bold = $false
            $sel.Font.Italic = $true
            $sel.TypeText($url)
            $sel.Font.Italic = $false
            $sel.TypeParagraph()
            continue
        }
        _VistaEscriuPara $sel $Script:WdNormal $l $bullet
    }
}

# ---- Vista d'un CATALEG (REQ1, TERMINI...) ---------------------------------
function _VistaCataleg($sel, [string]$jsonPath, [string]$nom) {
    $parsed = Read-CatalegJson $jsonPath
    _VistaEscriuPara $sel $Script:WdH1 ("Cat" + [char]0x00E0 + "leg " + $nom)
    if (-not [string]::IsNullOrWhiteSpace([string]$parsed.IntroText)) {
        _VistaEscriuPara $sel $Script:WdNormal ([string]$parsed.IntroText)
    }
    if ($parsed.IsFixedBody) {
        _VistaEscriuCos $sel @($parsed.FixedBodyLines)
        return
    }
    foreach ($sec in @($parsed.Sections)) {
        _VistaEscriuPara $sel $Script:WdH1 ([string]$sec.Title)
        # Els items pengen de Titol 2; quan apareix una subseccio, ella es el
        # Titol 2 i els items que la segueixen baixen a Titol 3.
        $hItem = $Script:WdH2
        foreach ($it in @($sec.Items)) {
            switch ([string]$it.Kind) {
                'subsection' {
                    _VistaEscriuPara $sel $Script:WdH2 ([string]$it.Short)
                    _VistaEscriuCos $sel @($it.BodyLines)
                    $hItem = $Script:WdH3
                }
                'intro' {
                    _VistaEscriuCos $sel @($it.BodyLines)
                }
                default {
                    _VistaEscriuPara $sel $hItem ([string]$it.Short)
                    _VistaEscriuCos $sel @($it.BodyLines)
                    foreach ($ch in @($it.Children)) {
                        _VistaEscriuCos $sel @($ch.BodyLines) $true
                    }
                }
            }
        }
    }
}

# ---- Vista de les CONCLUSIONS ----------------------------------------------
function _VistaConclusions($sel, [string]$jsonPath) {
    $o = _LoadEstructuralJson $jsonPath
    _VistaEscriuPara $sel $Script:WdH1 'Conclusions'
    foreach ($p in @($o.intro)) { _VistaEscriuPara $sel $Script:WdNormal (_JsonParaToBodyLine $p) }
    $sempre = New-Object System.Collections.ArrayList
    foreach ($n in @($o.nodes)) {
        if ([string]$n.tipus -eq 'sempre') {
            foreach ($p in @($n.cos)) { [void]$sempre.Add((_JsonParaToBodyLine $p)) }
            continue
        }
        # Grup = tipus d'informe (REQ1, SEGUIMENT, TERMINI...).
        _VistaEscriuPara $sel $Script:WdH1 ('Conclusions de: ' + [string]$n.titol)
        foreach ($c in @($n.fills)) {
            _VistaEscriuPara $sel $Script:WdH2 ([string]$c.titol)
            foreach ($p in @($c.cos)) { _VistaEscriuPara $sel $Script:WdNormal (_JsonParaToBodyLine $p) }
        }
    }
    if ($sempre.Count -gt 0) {
        _VistaEscriuPara $sel $Script:WdH1 ('Frases que surten SEMPRE')
        foreach ($l in $sempre) { _VistaEscriuPara $sel $Script:WdNormal ([string]$l) }
    }
}

# ---- Vista d'una plantilla ACT_EXTR ----------------------------------------
function _VistaActExtr($sel, [string]$jsonPath, [string]$nom) {
    $records = @(Read-ActExtrRecordsJson $jsonPath)
    _VistaEscriuPara $sel $Script:WdH1 ('Activitats extraordin' + [char]0x00E0 + 'ries ' + [char]0x00B7 + ' ' + $nom)
    foreach ($r in $records) {
        $txt = [string]$r.Text
        switch ([string]$r.Style) {
            'h1'     { _VistaEscriuPara $sel $Script:WdH1 $txt }
            'h2'     { _VistaEscriuPara $sel $Script:WdH2 (_VistaActExtrTitol $txt) }
            'url'    { _VistaEscriuCos $sel @('[[URL]] ' + $txt) }
            default  { if (-not [string]::IsNullOrWhiteSpace($txt)) { _VistaEscriuPara $sel $Script:WdNormal $txt } }
        }
    }
}

# ---- Genera la vista d'UN cataleg ------------------------------------------
function Export-VistaWord($word, [string]$jsonPath) {
    if (_VistaEsProtegit $jsonPath) { return $false }
    if (-not (Test-Path -LiteralPath $jsonPath)) { return $false }
    $o = _LoadEstructuralJson $jsonPath
    $familia = [string]$o.familia
    $nom = [System.IO.Path]::GetFileNameWithoutExtension($jsonPath)
    $out = _VistaWordPathFor $jsonPath

    $doc = $word.Documents.Add()
    try {
        $sel = $word.Selection
        switch ($familia) {
            'cataleg'     { _VistaCataleg $sel $jsonPath $nom }
            'conclusions' { _VistaConclusions $sel $jsonPath }
            'actextr'     { _VistaActExtr $sel $jsonPath $nom }
            default       { _VistaCataleg $sel $jsonPath $nom }
        }
        # Nota final: que quedi clar que es una vista generada i que no s'edita.
        _VistaEscriuPara $sel $Script:WdNormal ''
        _VistaEscriuPara $sel $Script:WdNormal ("//Vista generada autom" + [char]0x00E0 + "ticament des de " + [System.IO.Path]::GetFileName($jsonPath) + " el " + (Get-Date).ToString('dd/MM/yyyy HH:mm') + ". No l'editis: els canvis es fan des de l'editor de cat" + [char]0x00E0 + "legs del programa.//")
        $doc.SaveAs([ref]$out, [ref]16)   # 16 = wdFormatDocumentDefault (.docx)
        return $true
    } finally {
        try { $doc.Close($false) } catch { }
    }
}

# ---- Genera TOTES les vistes ------------------------------------------------
# Retorna el nombre de vistes generades. Fail-safe: si no hi ha Word, no peta.
function Invoke-ExportarVistesWord([switch]$Force) {
    if (-not (Test-Path -LiteralPath $EstructuralsDir)) { return 0 }
    $tots = @(Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' -ErrorAction SilentlyContinue |
              Where-Object { -not (_VistaEsProtegit $_.FullName) } | Sort-Object Name)
    # Nomes els que tinguin el JSON mes nou que la vista (o cap vista encara).
    $jsons = @()
    foreach ($j in $tots) {
        $out = _VistaWordPathFor $j.FullName
        $ex = Test-Path -LiteralPath $out
        $docxUtc = if ($ex) { (Get-Item -LiteralPath $out).LastWriteTimeUtc } else { [datetime]::MinValue }
        if (_VistaCalRegenerar $ex $j.LastWriteTimeUtc $docxUtc ([bool]$Force)) { $jsons += $j }
    }
    if ($jsons.Count -eq 0) { return 0 }

    $word = $null
    try { $word = New-Object -ComObject Word.Application } catch { $word = $null }
    if ($null -eq $word) {
        Write-Host "  Avis: no s'ha pogut obrir el Word; no s'han generat les vistes."
        return 0
    }
    $word.Visible = $false
    try { $word.DisplayAlerts = 0 } catch { }
    $n = 0
    try {
        foreach ($j in $jsons) {
            try {
                if (Export-VistaWord $word $j.FullName) {
                    $n++
                    Write-Host ("  vista: " + [System.IO.Path]::GetFileNameWithoutExtension($j.Name) + ".docx")
                }
            } catch {
                Write-Host ("  Avis: no s'ha pogut generar la vista de " + $j.Name + " (" + $_.Exception.Message + ")")
            }
        }
    } finally {
        try { $word.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    return $n
}
