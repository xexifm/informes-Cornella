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

# VERSIO del generador de vistes. Puja-la SEMPRE que canviï com es veuen les
# vistes: si no, les que ja existeixen es queden amb el format antic per sempre
# (la regla de sota nomes regenera quan el JSON es mes nou que el .docx, i just
# despres de generar-les el .docx sempre es el mes nou). En canviar de versio es
# regeneren totes una vegada.
#   1 -> primera versio (format propi, amb estils de titol)
#   2 -> format de l'informe (Format.ps1) + nivells d'esquema
#   3 -> tipografia base de la plantilla (Bookman Old Style, justificat,
#        interlineat i marges) via Format-ApplyBaseStyle
#   4 -> separacio entre l'item i el seu PRIMER sub-punt (Format-Bullet -First)
#   5 -> negreta del numero de l'item aplicada pel RANG (no s'encomana al cos)
#        i sangria dels fills a 1 cm amb francesa de 0,5 cm
$Script:VistaWordVersio = 5

function _VistaVersioPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path $base (Join-Path 'InformesCornella' 'vistes-versio.txt'))
}

# La versio amb que es van generar les vistes d'aquest ordinador (0 si no consta).
function _VistaVersioDesada {
    $p = _VistaVersioPath
    if (-not (Test-Path -LiteralPath $p)) { return 0 }
    try { return [int](Get-Content -LiteralPath $p -Raw -ErrorAction Stop).Trim() } catch { return 0 }
}

function _VistaDesaVersio([int]$v) {
    try {
        $p = _VistaVersioPath
        $d = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        [string]$v | Set-Content -LiteralPath $p -Encoding UTF8
    } catch { }
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
# El format es EXACTAMENT el de l'informe: totes les funcions de sota criden les
# Format-* de Format.ps1 (les mateixes que fa servir Build-Document), aixi la
# vista es veu igual que sortiria el document generat.
#
# A mes, cada titol rep un NIVELL D'ESQUEMA (OutlineLevel) perque surti al panell
# de navegacio del Word. L'OutlineLevel NO canvia com es veu el paragraf: nomes
# el fa navegable. Compte: el Word HERETA el nivell al paragraf seguent, per aixo
# el cos el torna sempre a 10 (wdOutlineLevelBodyText).
$Script:WdOutlineBody = 10

function _VistaNivell($sel, [int]$n) {
    try { $sel.ParagraphFormat.OutlineLevel = $n } catch { }
}

# --- Embolcalls: format de l'informe + nivell d'esquema ---------------------
function _VSection($sel, [string]$t)  { Format-Section $sel $t;    _VistaNivell $sel 1 }
function _VSubsection($sel, [string]$t) { Format-Subsection $sel $t; _VistaNivell $sel 2 }
function _VItem($sel, [string]$num, [string]$t) { Format-Item $sel $num $t; _VistaNivell $sel 3 }
function _VBody($sel, [string]$t, [bool]$isChild = $false) {
    if ($isChild) { Format-Body $sel $t -IsChild } else { Format-Body $sel $t }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VUrl($sel, [string]$u, [bool]$isChild = $false) {
    if ($isChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VBullet($sel, [string]$t, [bool]$isChild = $true, [bool]$first = $false) {
    if ($isChild) { Format-Bullet $sel $t -IsChild -First:$first } else { Format-Bullet $sel $t -First:$first }
    _VistaNivell $sel $Script:WdOutlineBody
}
function _VSpacer($sel) { Format-Spacer $sel; _VistaNivell $sel $Script:WdOutlineBody }

# Escriu una linia de cos separant text i URLs, com fa el motor (_SplitTextAndUrls).
function _VLine($sel, [string]$line, [bool]$isChild = $false) {
    if ([string]::IsNullOrWhiteSpace($line)) { return }
    $parts = _SplitTextAndUrls $line
    if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { _VBody $sel $parts.Text $isChild }
    foreach ($u in $parts.Urls) { _VUrl $sel $u $isChild }
}

# ---- Vista d'un CATALEG (REQ1, TERMINI...) ---------------------------------
# Reprodueix el que faria _WriteCatalegBody amb TOTS els items triats: seccio en
# MAJUSCULES, subseccio subratllada, items numerats amb el numero en negreta i
# fills com a punts de llista. Els [CAMP:]/[OPCIO:] es deixen tal qual (es una
# vista del cataleg, no un informe d'una activitat concreta).
function _VistaCataleg($sel, [string]$jsonPath, [string]$nom) {
    $parsed = Read-CatalegJson $jsonPath
    $cfg = $Script:ReportFormatConfig

    if (-not [string]::IsNullOrWhiteSpace([string]$parsed.IntroText)) {
        _VBody $sel ([string]$parsed.IntroText)
        if ($cfg.SpacerAfterIntroParagraph) { _VSpacer $sel }
    }
    if ($parsed.IsFixedBody) {
        $lines = @($parsed.FixedBodyLines)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            _VLine $sel ([string]$lines[$i])
            if ($i -lt ($lines.Count - 1)) { _VSpacer $sel }
        }
        return
    }

    $num = 0
    foreach ($sec in @($parsed.Sections)) {
        _VSection $sel ([string]$sec.Title)
        if ($cfg.SpacerAfterSection) { _VSpacer $sel }
        foreach ($el in @($sec.Items)) {
            switch ([string]$el.Kind) {
                'subsection' {
                    _VSubsection $sel ([string]$el.Short)
                    if ($cfg.SpacerAfterSubsection) { _VSpacer $sel }
                }
                'intro' {
                    foreach ($ln in @($el.BodyLines)) { _VLine $sel ([string]$ln) }
                    if ($cfg.SpacerAfterIntro) { _VSpacer $sel }
                }
                default {
                    $lines = @($el.BodyLines)
                    $escrit = $false
                    if ($lines.Count -gt 0) {
                        $num++
                        $p0 = _SplitTextAndUrls ([string]$lines[0])
                        _VItem $sel ("$num.") ([string]$p0.Text)
                        foreach ($u in $p0.Urls) { _VUrl $sel $u }
                        for ($i = 1; $i -lt $lines.Count; $i++) { _VLine $sel ([string]$lines[$i]) }
                        $escrit = $true
                    }
                    $primerFill = $true
                    foreach ($ch in @($el.Children)) {
                        $cl = @($ch.BodyLines)
                        if ($cl.Count -eq 0) { continue }
                        if (-not $escrit) { $num++; $escrit = $true }
                        # Els fills NO es numeren: van amb pic, com a l'informe.
                        # El PRIMER se separa mes de l'item (com fa Motor.ps1).
                        $pc = _SplitTextAndUrls ([string]$cl[0])
                        if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                            _VBullet $sel ([string]$pc.Text) $true $primerFill
                            $primerFill = $false
                        }
                        foreach ($u in $pc.Urls) { _VUrl $sel $u $true }
                        for ($i = 1; $i -lt $cl.Count; $i++) { _VLine $sel ([string]$cl[$i]) $true }
                    }
                    if ($escrit -and $cfg.SpacerAfterItem) { _VSpacer $sel }
                }
            }
        }
    }
}

# ---- Vista de les CONCLUSIONS ----------------------------------------------
function _VistaConclusions($sel, [string]$jsonPath) {
    $o = _LoadEstructuralJson $jsonPath
    $sempre = New-Object System.Collections.ArrayList
    foreach ($p in @($o.intro)) {
        $t = _JsonParaToBodyLine $p
        if (-not [string]::IsNullOrWhiteSpace($t)) { _VSection $sel $t }
    }
    foreach ($n in @($o.nodes)) {
        if ([string]$n.tipus -eq 'sempre') {
            foreach ($p in @($n.cos)) { [void]$sempre.Add((_JsonParaToBodyLine $p)) }
            continue
        }
        # Grup = tipus d'informe (REQ1, SEGUIMENT, TERMINI...).
        _VSpacer $sel
        _VSection $sel ('Conclusions de ' + [string]$n.titol)
        _VSpacer $sel
        $num = 0
        foreach ($c in @($n.fills)) {
            $num++
            $cos = @($c.cos)
            $primera = if ($cos.Count -gt 0) { _JsonParaToBodyLine $cos[0] } else { '' }
            _VItem $sel ("$num.") ([string]$primera)
            for ($i = 1; $i -lt $cos.Count; $i++) { _VLine $sel (_JsonParaToBodyLine $cos[$i]) }
            _VSpacer $sel
        }
    }
    if ($sempre.Count -gt 0) {
        _VSpacer $sel
        _VSection $sel 'Frases que surten sempre'
        _VSpacer $sel
        foreach ($l in $sempre) { _VLine $sel ([string]$l) }
    }
}

# ---- Vista d'una plantilla ACT_EXTR ----------------------------------------
function _VistaActExtr($sel, [string]$jsonPath, [string]$nom) {
    $records = @(Read-ActExtrRecordsJson $jsonPath)
    $num = 0
    foreach ($r in $records) {
        $txt = [string]$r.Text
        switch ([string]$r.Style) {
            'h1' {
                _VSpacer $sel
                _VSection $sel $txt
                _VSpacer $sel
                $num = 0
            }
            'h2' {
                # La capcalera del bloc ("[[CLAU]] ::TOKEN:: etiqueta") no surt a
                # l'informe: al document nomes hi va el CONTINGUT. A la vista si
                # que la posem (subratllada) per saber quin bloc es cadascun.
                _VSubsection $sel (_VistaActExtrTitol $txt)
            }
            'url' { _VUrl $sel $txt }
            default {
                if (-not [string]::IsNullOrWhiteSpace($txt)) {
                    $num++
                    _VItem $sel ("$num.") $txt
                }
            }
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
        # Un document NOU de Word surt en Calibri, alineat a l'esquerra i amb
        # uns altres marges. Li posem la MATEIXA base que la plantilla de
        # l'informe (Bookman Old Style, justificat, interlineat i marges), que
        # es on esta declarada: Format.ps1.
        Format-ApplyBaseStyle $doc
        $sel = $word.Selection
        switch ($familia) {
            'cataleg'     { _VistaCataleg $sel $jsonPath $nom }
            'conclusions' { _VistaConclusions $sel $jsonPath }
            'actextr'     { _VistaActExtr $sel $jsonPath $nom }
            default       { _VistaCataleg $sel $jsonPath $nom }
        }
        # Nota final: que quedi clar que es una vista generada i que no s'edita.
        _VSpacer $sel
        _VBody $sel ("//Vista generada autom" + [char]0x00E0 + "ticament des de " + [System.IO.Path]::GetFileName($jsonPath) + " el " + (Get-Date).ToString('dd/MM/yyyy HH:mm') + ". No l'editis: els canvis es fan des de l'editor de cat" + [char]0x00E0 + "legs del programa.//")
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
    # Si ha canviat el format de les vistes, es regeneren TOTES una vegada.
    $canviDeFormat = ((_VistaVersioDesada) -ne $Script:VistaWordVersio)
    $forcar = ([bool]$Force -or $canviDeFormat)
    if ($canviDeFormat) { Write-Host "  (el format de les vistes ha canviat: es regeneren totes)" }

    # Nomes els que tinguin el JSON mes nou que la vista (o cap vista encara).
    $jsons = @()
    foreach ($j in $tots) {
        $out = _VistaWordPathFor $j.FullName
        $ex = Test-Path -LiteralPath $out
        $docxUtc = if ($ex) { (Get-Item -LiteralPath $out).LastWriteTimeUtc } else { [datetime]::MinValue }
        if (_VistaCalRegenerar $ex $j.LastWriteTimeUtc $docxUtc $forcar) { $jsons += $j }
    }
    if ($jsons.Count -eq 0) {
        if ($canviDeFormat) { _VistaDesaVersio $Script:VistaWordVersio }
        return 0
    }

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
    # Nomes donem la versio per bona si s'han pogut generar (si Word ha fallat,
    # la propera vegada ho tornara a intentar).
    if ($n -gt 0) { _VistaDesaVersio $Script:VistaWordVersio }
    return $n
}
