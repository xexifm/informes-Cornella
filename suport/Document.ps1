#requires -Version 5.1
<#
.SYNOPSIS
  Composicio del document Word final (Pas 6).

.DESCRIPTION
  Copia '0 CAPCALERA.docx', hi substitueix els <<PLACEHOLDERS>>, hi escriu el
  cos (seccions, items, sub-punts, enllacos) i el bloc de conclusions, i el desa
  amb el nom que toca a la carpeta de sortida. Tot el FORMAT (lletra, sangries,
  espaiats) el posa Format.ps1: aqui nomes es decideix QUE s'escriu i EN QUIN
  ORDRE.

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Step 6 - Compose final document
# ----------------------------------------------------------------------------
# Retalla 0 CAPCALERA.docx per quedar-se nomes amb el bloc de capcalera demanat.
# El document pot contenir DUES capcaleres: la de REQ1 (a dalt, l'original) i la
# d'ACT_EXTR (a sota), separades per un paragraf marcador "[[CAP:ACT_EXTR]]".
#   $which = 'REQ1'     -> esborra des del marcador fins al final (i el marcador).
#   $which = 'ACT_EXTR' -> esborra des de l'inici fins al marcador (inclos).
# Si el marcador no existeix (capcalera antiga, nomes REQ1), no fa res. Aixi es
# retrocompatible amb una 0 CAPCALERA.docx que encara no tingui el bloc ACT_EXTR.
function Select-CapcaleraBlock($doc, [string]$which) {
    $marker = $null
    foreach ($p in $doc.Paragraphs) {
        $t = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
        if ($t.Trim() -eq '[[CAP:ACT_EXTR]]') { $marker = $p; break }
    }
    if ($null -eq $marker) { return }   # nomes hi ha la capcalera REQ1: res a fer
    if ($which -eq 'ACT_EXTR') {
        # Esborra tot el que hi ha ABANS del marcador (bloc REQ1 + taula) i el
        # propi marcador.
        $rng = $doc.Range(0, $marker.Range.End)
        $rng.Delete() | Out-Null
    } else {
        # REQ1: esborra des del marcador (inclos) fins al final del document.
        $rng = $doc.Range($marker.Range.Start, $doc.Content.End)
        $rng.Delete() | Out-Null
    }
}

# Munta el text de la linia "Objecte:" (placeholder <<ORIGEN>>) de la capcalera
# generica a partir de l'origen triat al Pas 2. Funcio PURA (nomes llegeix el
# hashtable/objecte $header), testejable en headless. Accents amb codepoint
# per no dependre de l'encoding amb que PowerShell 5.1 llegeix aquest fitxer.
#   'doc'  -> "Doc. aportada amb Num. d'anotacio <NUM_ANOTACIO> del <DATA_ANOTACIO>"
#   'insp' -> "Visita inspeccio <DATA_INSPECCIO>"
#   'cap'  -> '' (sense Objecte; p.ex. requeriments de control periodic)
function _BuildOrigenText($header) {
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] }; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    $tipus = (& $get 'ORIGEN_TIPUS'); if ([string]::IsNullOrWhiteSpace($tipus)) { $tipus = 'doc' }
    if ($tipus -eq 'cap')  { return '' }
    if ($tipus -eq 'insp') {
        return 'Visita inspecci' + [char]0x00F3 + ' ' + (& $get 'DATA_INSPECCIO')
    }
    return 'Doc. aportada amb N' + [char]0x00FA + 'm. d' + [char]0x2019 + 'anotaci' + [char]0x00F3 +
           ' ' + (& $get 'NUM_ANOTACIO') + ' del ' + (& $get 'DATA_ANOTACIO')
}

function Apply-HeaderReplacements($doc, $header) {
    # Substituim els placeholders <<NOM>> de la capcalera pels valors del Pas 2.
    # Inclou els d'ACT_EXTR (<<DATES>>, <<AFORAMENT>>); si no apareixen a la
    # capcalera triada, simplement no es substitueix res.
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] }; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    # <<ORIGEN>>: la linia "Objecte:" de la capcalera generica (REQ1) es
    # construeix segons el que s'ha triat al Pas 2 (documentacio aportada o
    # visita d'inspeccio). Vegeu _BuildOrigenText (funcio pura, testejable).
    $origen = _BuildOrigenText $header
    $map = @{
        '<<ID_GIA>>'         = (& $get 'ID_GIA')
        '<<EXP_NUM>>'        = (& $get 'EXP_NUM')
        '<<ADRECA>>'         = (& $get 'ADRECA')
        '<<ACTIVITAT>>'      = (& $get 'ACTIVITAT')
        '<<TITULAR>>'        = (& $get 'TITULAR')
        '<<ORIGEN>>'         = $origen
        '<<NUM_ANOTACIO>>'   = (& $get 'NUM_ANOTACIO')
        '<<DATA_ANOTACIO>>'  = (& $get 'DATA_ANOTACIO')
        '<<DATA_INSPECCIO>>' = (& $get 'DATA_INSPECCIO')
        '<<DATES>>'          = (& $get 'DATES')
        '<<AFORAMENT>>'      = (& $get 'AFORAMENT')
    }
    foreach ($k in $map.Keys) {
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $k
        $find.Replacement.Text = [string]$map[$k]
        $find.Forward = $true
        $find.Wrap = 1
        $find.MatchCase = $false
        $find.Execute([ref]$k, $false, $false, $false, $false, $false, $true, 1, $false, [string]$map[$k], 2) | Out-Null
    }
}

# Calcula el nom de fitxer de sortida: YYYY-MM-DD_<TipusCataleg>_GIA <id>.docx
function _GetOutputFileName($catalegName, $gia) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $cat   = $catalegName
    if ($cat) { $cat = $cat.Substring(0,1).ToUpper() + $cat.Substring(1).ToLower() }
    else      { $cat = 'Informe' }
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    return ("{0}_{1}_GIA {2}.docx" -f $today, $cat, $gia)
}

# A partir del nom base, retorna la ruta a $targetDir que no col·lisioni amb
# cap fitxer existent. Si el primer ja existeix, prova "_2", "_3"... fins
# trobar-ne un de lliure. Aixi pots fer diversos informes del mateix GIA el
# mateix dia sense haver de tancar Word ni renombrar res manualment.
#
# Ex.: "2026-05-29_Req1_GIA 1379.docx" existeix
#      -> torna "2026-05-29_Req1_GIA 1379_2.docx"
function _GetUniqueOutputPath($targetDir, $baseFileName) {
    $candidate = Join-Path $targetDir $baseFileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($baseFileName)
    $ext  = [System.IO.Path]::GetExtension($baseFileName)
    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $targetDir ("{0}_{1}{2}" -f $stem, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    # Si arribem aqui es que hi ha mes de 999 informes pel mateix GIA i dia;
    # cas extrem, retornem un nom amb timestamp.
    return Join-Path $targetDir ("{0}_{1}{2}" -f $stem, (Get-Date -Format 'HHmmss'), $ext)
}

# Determina el directori de sortida: l'$OutputDir si es accessible, en cas
# contrari local\informes-generats\ (al
# costat dels .bat).
function _ResolveOutputDir {
    $targetDir = $OutputDir
    try {
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
        return $targetDir
    } catch {
        $local = Get-LocalSubdir $RepoRoot 'Informes'
        if (-not (Test-Path -LiteralPath $local)) {
            New-Item -ItemType Directory -Path $local -Force | Out-Null
        }
        return $local
    }
}

# Obre el document Word a partir d'una copia LOCAL de la capcalera (per
# evitar la "Vista protegida" en unitats de xarxa). Retorna el doc obert i
# la ruta temporal.
function _OpenOutputDocument($word, $tempPath) {
    Copy-Item -LiteralPath $HeaderPath -Destination $tempPath -Force
    $doc = $word.Documents.Open($tempPath, $false, $false)
    try {
        if ($doc.ProtectedViewWindow -ne $null) {
            $doc = $doc.ProtectedViewWindow.Edit()
        }
    } catch { }
    return $doc
}

# Escriu el cos del document (intro del cataleg + seccions amb items numerats).
# Retorna el comptador global utilitzat per a la numeracio.
function _WriteCatalegBody($sel, $cfg, $selectedSections, $fields, $introText, $isFixedBody = $false, $fixedBodyLines = @()) {
    # Informe de cos fix (p.ex. TERMINI.docx): no hi ha seccions ni items a
    # numerar; el cos son directament els paragrafs del document, amb els camps
    # [CAMP:]/[OPCIO:] resolts i separant text/URLs com a la resta del motor.
    if ($isFixedBody) {
        $lines = @($fixedBodyLines)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $resolved = [string](Apply-Fields -text $lines[$i] -fields $fields)
            if ([string]::IsNullOrWhiteSpace($resolved)) { continue }
            $parts = _SplitTextAndUrls $resolved
            if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { Format-Body $sel $parts.Text }
            foreach ($u in $parts.Urls) { Format-Url $sel $u }
            if ($i -lt ($lines.Count - 1)) { Format-Spacer $sel }
        }
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($introText)) {
        Format-Body $sel $introText
        if ($cfg.SpacerAfterIntroParagraph) { Format-Spacer $sel }
    }

    # Resol [CAMP: ...] a cada linia i EMET cada linia resolta al pipeline.
    # Els cridadors fan servir @(& $resolveLines ...) per recollir un array
    # PLA de cadenes. (No retornem ,@(...): combinat amb el @() del cridador
    # provocava un doble embolcall on $itemLines[0] era TOT l'array de linies
    # en comptes de la primera linia -> trencava la separacio text/URL i feia
    # petar Substring.)
    $resolveLines = {
        param($lines)
        foreach ($ln in $lines) {
            [string](Apply-Fields -text $ln -fields $fields)
        }
    }

    # Emet una linia separant text i URLs: el text (si n'hi ha) va com a cos i
    # cada URL com a hipervincle en paragraf propi.
    $emitLine = {
        param($line, $isChild)
        if ([string]::IsNullOrWhiteSpace($line)) { return }
        $parts = _SplitTextAndUrls $line
        if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
            if ($isChild) { Format-Body $sel $parts.Text -IsChild } else { Format-Body $sel $parts.Text }
        }
        foreach ($u in $parts.Urls) {
            if ($isChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
        }
    }

    $emitExtras = {
        param($lines, $isChild)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            & $emitLine $lines[$i] $isChild
        }
    }

    $emitIntro = {
        param($introEl)
        $lines = @(& $resolveLines $introEl.BodyLines)
        foreach ($bp in $lines) { & $emitLine $bp $false }
        if ($cfg.SpacerAfterIntro) { Format-Spacer $sel }
    }

    $emitItem = {
        param($it)
        $itemLines = @(& $resolveLines $it.BodyLines)
        $hasChildren = ($it.Children.Count -gt 0)
        $itemWritten = $false

        if ($it.Selected -or $hasChildren) {
            if ($itemLines.Count -gt 0) {
                $script:_buildGlobal++
                # Separem un possible URL enganxat al text principal de l'item.
                $p0 = _SplitTextAndUrls $itemLines[0]
                Format-Item $sel "$($script:_buildGlobal)." $p0.Text
                foreach ($u in $p0.Urls) { Format-Url $sel $u }
                & $emitExtras $itemLines $false
                $itemWritten = $true
            }
        }
        if ($hasChildren) {
            $firstChild = $true
            foreach ($ch in $it.Children) {
                $childLines = @(& $resolveLines $ch.BodyLines)
                if ($childLines.Count -eq 0) { continue }
                if (-not $itemWritten) {
                    $script:_buildGlobal++
                    $itemWritten = $true
                }
                # Els fills (::CHILD::) NO es numeren: s'emeten com a PUNT DE
                # LLISTA amb Format-Bullet (-IsChild = sangria de sub-nivell).
                # Les linies extra del fill (p.ex. un URL) segueixen com a
                # enllac/cos de fill, sense un nou pic.
                $pc = _SplitTextAndUrls $childLines[0]
                if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                    # -First: el primer sub-punt se separa de l'item numerat amb
                    # ItemSpaceAfterPt (els seguents, amb BulletSpaceBeforePt).
                    Format-Bullet $sel $pc.Text -IsChild -First:$firstChild
                    $firstChild = $false
                }
                foreach ($u in $pc.Urls) { Format-Url $sel $u -IsChild }
                & $emitExtras $childLines $true
            }
        }
        if ($itemWritten -and $cfg.SpacerAfterItem) { Format-Spacer $sel }
    }

    $script:_buildGlobal = 0
    $lastSectionName = $null

    foreach ($sec in $selectedSections) {
        $parts = $sec.Title -split ' - ', 2
        if ($parts.Count -eq 2) {
            $secName = $parts[0].Trim()
            $subName = $parts[1].Trim()
            if ($secName -ne $lastSectionName) {
                Format-Section $sel $secName
                if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
                $lastSectionName = $secName
            }
            Format-Subsection $sel $subName
            if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
        } else {
            Format-Section $sel $sec.Title
            if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
            $lastSectionName = $sec.Title
        }

        # Les subseccions i les intros s'emeten "tard": nomes quan ve un
        # item REAL que les segueix. Si una secció conté 3 ::SUB:: pero
        # nomes s'ha triat un ítem que viu a la 3a subsecció, només
        # surten el títol de la secció i la 3a subsecció (no les 2
        # anteriors buides). Una nova subsecció sobreescriu la pendent.
        $pendingSubsection = $null
        $pendingIntro = $null
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'subsection') {
                $pendingSubsection = $el
                $pendingIntro = $null   # una nova subseccio invalida l'intro pendent
                continue
            }
            if ($el.Kind -eq 'intro') {
                $pendingIntro = $el
                continue
            }
            # Item real: emetem primer la subseccio pendent (si en hi ha),
            # despres l'intro pendent, i finalment l'item.
            if ($null -ne $pendingSubsection) {
                Format-Subsection $sel $pendingSubsection.Short
                if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
                $pendingSubsection = $null
            }
            if ($null -ne $pendingIntro) {
                & $emitIntro $pendingIntro
                $pendingIntro = $null
            }
            & $emitItem $el
        }
    }
}

# Escriu el bloc de conclusions:
#   - $headerText : titol del bloc (sol ser 'CONCLUSIONS'), centrat-negreta.
#                   '' = no s'emet.
#   - $conclusions : array d'objectes {Title; Body} de les conclusions
#                    TRIADES al Pas 5. Es emet el seu Body.
#   - $alwaysConclusions : array de cadenes ja sense el prefix '::SEMPRE::'
#                          que s'emeten sempre, despres de les triades.
#   - $fields : per resoldre [CAMP:] i [OPCIO:] dins els textos.
# La separacio entre conclusions la posa Format-Conclusion via SpaceAfter
# (ConclusionSpaceAfterPt), aixi que aqui no hi fa falta un Spacer entre.
function _WriteConclusionsBlock($sel, $cfg, $headerText, $conclusions, $alwaysConclusions, $fields) {
    $hasBody = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    $hasHead = -not [string]::IsNullOrWhiteSpace($headerText)
    if (-not $hasBody -and -not $hasHead) { return }

    if ($cfg.SpacerBeforeConclusionsBlock) { Format-Spacer $sel }

    if ($hasHead) {
        Format-ConclusionHeader $sel $headerText
    }

    foreach ($c in $conclusions) {
        $txt = if ($c -is [string]) { $c } else { [string]$c.Body }
        $resolved = Apply-Fields -text $txt -fields $fields
        Format-Conclusion $sel $resolved
    }
    foreach ($a in $alwaysConclusions) {
        $resolved = Apply-Fields -text ([string]$a) -fields $fields
        Format-Conclusion $sel $resolved
    }
}

function Build-Document($word, $header, $selectedSections, $fields, $conclusions, $alwaysConclusions, $catalegName, $introText, $conclusionsHeaderText, $isFixedBody = $false, $fixedBodyLines = @()) {
    $baseName  = _GetOutputFileName $catalegName $header['ID_GIA']
    $targetDir = _ResolveOutputDir
    # Triem el primer nom lliure al directori de sortida (afegim _2, _3...
    # si ja existeix). Aixi pots generar diversos informes del mateix dia/GIA
    # sense que cap es sobreescrigui.
    $outPath  = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)

    # Treballem amb una copia LOCAL (a %TEMP%) per evitar que Word obri el
    # fitxer en "Vista protegida" quan el desti es una unitat de xarxa.
    # El temp porta el mateix nom (ja unic) que el desti final.
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    # 0 CAPCALERA.docx pot portar tambe el bloc d'ACT_EXTR a sota; ens quedem
    # nomes amb el bloc de REQ1 (no fa res si el marcador no hi es).
    Select-CapcaleraBlock $doc 'REQ1'
    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)  # wdStory = 6

    $cfg = $Script:ReportFormatConfig
    _WriteCatalegBody $sel $cfg $selectedSections $fields $introText $isFixedBody $fixedBodyLines
    _WriteConclusionsBlock $sel $cfg $conclusionsHeaderText $conclusions $alwaysConclusions $fields

    $doc.Save()
    $doc.Close($false)

    # Movem el fitxer al desti final (xarxa o local segons disponibilitat).
    try {
        Move-Item -LiteralPath $tempPath -Destination $outPath -Force
    } catch {
        return $tempPath
    }
    return $outPath
}
