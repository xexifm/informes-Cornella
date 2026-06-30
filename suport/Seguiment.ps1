#requires -Version 5.1
<#
.SYNOPSIS
  Mode "Informe de seguiment": pren un informe anterior amb requeriments
  enumerats i hi afegeix, sota cada punt, una anotacio datada indicant si la
  documentacio presentada l'ha resolt o no.

.DESCRIPTION
  Es un mode alternatiu del programa (es tria a la pantalla inicial). NO
  regenera l'informe: edita una COPIA del .docx i nomes
  insereix/esborra/formata, de manera que es preserva exactament la capcalera,
  el text dels requeriments i el format (tambe d'informes fets a ma).

  L'EDICIO es fa manipulant directament el XML intern del .docx (un .docx es un
  ZIP amb word/document.xml). NO es fa servir Word per editar: aixi s'eviten
  tots els problemes de Word COM (mode lectura, nomes-lectura, vista protegida,
  document actiu, marca final de paragraf...). Word nomes s'invoca, opcionalment,
  al final per OBRIR el resultat (via Invoke-Item, l'app per defecte del SO).

  Flux (Invoke-SeguimentFlow):
    1. Triar l'informe anterior (.docx).
    2. Llegir-lo (XML) i modelar requeriments + anotacions existents.
    3. Triar el primer paragraf de conclusions a esborrar (manual).
    4. Introduir la data de la ronda (per defecte, avui).
    5. Per cada requeriment: comentari nou + checkbox "Resolt".
    6. Triar conclusions (de 0 CONCLUSIONS.docx) i omplir camps.
    7. Editar el XML: esborrar conclusions -> inserir anotacions -> recalcular
       negreta -> afegir conclusions -> desar a un .docx nou.

  Iteratiu: tornar a passar-lo sobre un informe de seguiment AFEGEIX una linia
  nova sota cada requeriment (no duplica). La negreta es dinamica: mentre un
  requeriment NO estigui resolt, el requeriment i les seves anotacions van en
  negreta; quan es resol, res en negreta.

.NOTES
  Aquest fitxer es carrega via dot-source des de GenerarInforme.ps1 (tambe en
  mode headless de proves). Les funcions pures i les de manipulacio XML son
  testejables a Linux (sense Word); nomes els dialegs WinForms necessiten Windows.

  Reutilitza de GenerarInforme.ps1: _NormalizeText, Test-StyleMatch,
  Apply-Fields, Select-Conclusions (amb camps inline), Get-FieldValuesForSession,
  _ResolveOutputDir, _GetUniqueOutputPath.
#>

# Espai de noms WordprocessingML.
$Script:WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Format del text afegit pel seguiment (anotacions i conclusions): ha de ser
# com l'estil Normal de REQ1 -> Bookman Old Style 11, justificat.
$Script:SeguimentFontName     = 'Bookman Old Style'
$Script:SeguimentFontHalfPt   = '22'   # 22 mig-punts = 11 pt

# Frases per defecte del comentari segons la casella "Resolt" (editables).
$Script:SeguimentPhraseResolt   = "S'aporta."
$Script:SeguimentPhraseNoResolt = "No s'aporta."

# ----------------------------------------------------------------------------
# Expressions regulars (a nivell de script: definides en carregar el fitxer).
# ----------------------------------------------------------------------------
# Requeriment numerat: comenca per "N." seguit d'espai. Ex: "1. Baixa tensio."
$Script:SeguimentReqRegex      = [regex]'^\s*(\d+)\.\s'
# Anotacio datada: "dd/MM/aaaa:" al principi. Ex: "01/06/2026: No s'entrega."
$Script:SeguimentAnnotRegex    = [regex]'^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*:'
# Numero de llista auto-numerada del Word (ListFormat.ListString). Ex: "1." o "1"
$Script:SeguimentListNumRegex  = [regex]'^\d+\.?$'

# ----------------------------------------------------------------------------
# FUNCIONS PURES (sense COM/WinForms) - testejables en headless
# ----------------------------------------------------------------------------

# Classifica un paragraf a partir del seu text i del numero de llista (si en
# te). Retorna { Kind; Number; Date; ViaList } on Kind es:
#   'requirement' (requeriment numerat) | 'annotation' (anotacio datada) | 'other'
function _ClassifyParagraph([string]$text, [string]$listString) {
    $t  = if ($null -eq $text)       { '' } else { ([string]$text).Trim() }
    $ls = if ($null -eq $listString) { '' } else { ([string]$listString).Trim() }

    # 1) Anotacio datada (es comprova primer; "dd/MM/aaaa" no casa amb "N.").
    $ma = $Script:SeguimentAnnotRegex.Match($t)
    if ($ma.Success) {
        $d = '{0:D2}/{1:D2}/{2}' -f [int]$ma.Groups[1].Value, [int]$ma.Groups[2].Value, $ma.Groups[3].Value
        return [pscustomobject]@{ Kind='annotation'; Number=$null; Date=$d; ViaList=$false }
    }

    # 2) Requeriment amb numero literal al text ("1. ...").
    $mr = $Script:SeguimentReqRegex.Match($t)
    if ($mr.Success) {
        return [pscustomobject]@{ Kind='requirement'; Number=[int]$mr.Groups[1].Value; Date=$null; ViaList=$false }
    }

    # 3) Requeriment via llista auto-numerada del Word (informes fets a ma).
    if ($ls -ne '' -and $Script:SeguimentListNumRegex.IsMatch($ls)) {
        $num = ($ls.TrimEnd('.') -as [int])
        return [pscustomobject]@{ Kind='requirement'; Number=$num; Date=$null; ViaList=$true }
    }

    return [pscustomobject]@{ Kind='other'; Number=$null; Date=$null; ViaList=$false }
}

# Decideix l'estat "resolt" a partir d'un valor de negreta estil Word.Font.Bold:
#   -1       = tot en negreta  -> pendent  (NO resolt)
#    0       = res en negreta   -> resolt
#    9999999 = mixt (p.ex. el "N." en negreta i el text no)
#              -> el tractem com a pendent (NO resolt), que es el cas de
#                 l'informe original acabat de generar.
function _InferResolvedFromBold($boldValue) {
    return ((($boldValue -as [int]) -eq 0))
}

# Mentre un requeriment no estigui resolt, ha d'anar en negreta.
function _ShouldBeBold($resolved) {
    return (-not $resolved)
}

# Munta el model ordenat de requeriments-amb-anotacions a partir d'un array de
# registres de paragraf { Index; Text; ListString; Bold }. Funcio PURA.
function _BuildSeguimentModel($paraRecords) {
    $reqs       = New-Object System.Collections.ArrayList
    $current    = $null
    $lastReqIdx = 0
    foreach ($r in $paraRecords) {
        $c = _ClassifyParagraph $r.Text $r.ListString
        if ($c.Kind -eq 'requirement') {
            $current = [pscustomobject]@{
                Number         = $c.Number
                ParaIndex      = [int]$r.Index
                Text           = ([string]$r.Text).Trim()
                IsAutoNumbered = [bool]$c.ViaList
                Bold           = $r.Bold
                WasResolved    = (_InferResolvedFromBold $r.Bold)
                Annotations    = (New-Object System.Collections.ArrayList)
            }
            [void]$reqs.Add($current)
            $lastReqIdx = [int]$r.Index
        }
        elseif ($c.Kind -eq 'annotation' -and $null -ne $current) {
            [void]$current.Annotations.Add([pscustomobject]@{
                ParaIndex = [int]$r.Index
                Date      = $c.Date
                Text      = ([string]$r.Text).Trim()
                Bold      = $r.Bold
            })
        }
        # 'other' (cos, URL, intro, capcalera, conclusions...) s'ignora al model.
    }
    # Estat "resolt" (per defecte de la casella i per saber si cal reescriure):
    #   - Sense anotacions previes -> primera ronda -> PENDENT.
    #   - Amb anotacions -> ho inferim de l'ULTIMA anotacio: si el seu comentari
    #     esta en negreta, el punt encara estava PENDENT; si no, RESOLT. (El
    #     seguiment nomes posa en negreta el comentari de l'ultima entrega quan
    #     queda pendent; aixi l'estat queda "guardat" dins el propi .docx.)
    foreach ($req in $reqs) {
        if ($req.Annotations.Count -eq 0) {
            $req.WasResolved = $false
        } else {
            $last = $req.Annotations[$req.Annotations.Count - 1]
            $req.WasResolved = (_InferResolvedFromBold $last.Bold)
        }
    }
    return [pscustomobject]@{
        Requirements     = $reqs.ToArray()
        LastReqParaIndex = $lastReqIdx
    }
}

# Localitza l'index (1-based) del primer paragraf del bloc de conclusions:
# busca NOMES despres de l'ultim requeriment (perque una frase dins un cos de
# requeriment no dispari) i casa qualsevol de les frases conegudes de manera
# insensible a accents/majuscules. Retorna -1 si no en troba cap.
#   $paraTexts : array de textos de paragraf (posicio i => paragraf i+1).
function _FindConclusionStartIndex($paraTexts, [int]$lastReqEndIndex, $phrases) {
    if ($null -eq $paraTexts) { return -1 }
    $normPhrases = @()
    foreach ($ph in $phrases) {
        $np = _NormalizeText $ph
        if (-not [string]::IsNullOrWhiteSpace($np)) { $normPhrases += $np }
    }
    if ($normPhrases.Count -eq 0) { return -1 }

    for ($i = 0; $i -lt $paraTexts.Count; $i++) {
        $paraIndex = $i + 1
        if ($paraIndex -le $lastReqEndIndex) { continue }
        $nt = _NormalizeText $paraTexts[$i]
        if ([string]::IsNullOrWhiteSpace($nt)) { continue }
        foreach ($np in $normPhrases) {
            if ($nt.Contains($np)) { return $paraIndex }
        }
    }
    return -1
}

# Valida/normalitza la data de la ronda. Buit -> avui. Retorna { Ok; Normalized }
# sempre en dd/MM/yyyy. Rebutja dates impossibles (32/13/2026...).
function _ValidateRoundDate($text) {
    $s = if ($null -eq $text) { '' } else { ([string]$text).Trim() }
    if ($s -eq '') {
        return [pscustomobject]@{ Ok=$true; Normalized=(Get-Date).ToString('dd/MM/yyyy') }
    }
    $dt   = [datetime]::MinValue
    $fmts = [string[]]@('dd/MM/yyyy','d/M/yyyy','dd/MM/yy','d/M/yy')
    $ci   = [System.Globalization.CultureInfo]::InvariantCulture
    if ([datetime]::TryParseExact($s, $fmts, $ci, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return [pscustomobject]@{ Ok=$true; Normalized=$dt.ToString('dd/MM/yyyy') }
    }
    return [pscustomobject]@{ Ok=$false; Normalized=$s }
}

# Format unic de la linia d'anotacio: "dd/MM/aaaa: comentari".
function _FormatAnnotationLine($dateStr, $comment) {
    $c = if ($null -eq $comment) { '' } else { ([string]$comment).Trim() }
    return ('{0}: {1}' -f $dateStr, $c)
}

# Nom del fitxer de sortida del seguiment. La DATA es la d'AVUI (el dia que es
# genera l'informe), no la de l'anotacio. S'incrementa el numero de ronda "Req":
#   - "...Req1..."           -> "<avui>_...Req2...", "Req2"->"Req3"...
#   - "...Req..." sense num   -> "<avui>_...Req2..." (requeriments antics fets
#                                sense el programa, que nomes posen "Req").
#   - sense "Req" enlloc      -> "<avui>_Seguiment_<nom>.docx".
# Es treu la data inicial del nom origen i es preserva la resta (GIA, titular...).
# Sempre s'ha de passar el resultat per _GetUniqueOutputPath (afegeix _2, _3...).
function _SeguimentOutputName([string]$sourceBaseName, [datetime]$today) {
    $day  = $today.ToString('yyyy-MM-dd')
    # Treu una data inicial "YYYY-MM-DD" (amb "_" o espai darrere) del nom origen.
    $rest = [regex]::Replace([string]$sourceBaseName, '^\s*\d{4}-\d{2}-\d{2}[_ ]+', '')
    # Busca "Req" amb numero opcional i l'incrementa (sense numero -> 2).
    $m = [regex]::Match($rest, '(?i)Req(\d*)')
    if ($m.Success) {
        $n = if ($m.Groups[1].Value -ne '') { [int]$m.Groups[1].Value + 1 } else { 2 }
        $rest = $rest.Remove($m.Index, $m.Length).Insert($m.Index, "Req$n")
        $name = '{0}_{1}' -f $day, $rest
    } else {
        $safe = $rest
        if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'informe' }
        $name = '{0}_Seguiment_{1}' -f $day, $safe
    }
    $name = ($name -replace '[\\/:*?"<>|]','_').Trim()
    return ($name + '.docx')
}

# ----------------------------------------------------------------------------
# CAPA XML - lectura i edicio del .docx SENSE Word (ZIP + WordprocessingML)
# Testejable a Linux.
# ----------------------------------------------------------------------------

function _NewWordNsMgr($xml) {
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $Script:WNS)
    # ,$ns: l'XmlNamespaceManager es IEnumerable; sense la coma, PowerShell
    # l'enumeraria i retornaria un array de prefixos en lloc del gestor.
    return ,$ns
}

# Llegeix word/document.xml d'un .docx i retorna el seu text (UTF-8).
function _ReadDocxPartText($docxPath, $partName) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
    $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $entry = $zip.GetEntry($partName)
        if ($null -eq $entry) { return $null }
        $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try { return $sr.ReadToEnd() } finally { $sr.Close() }
    } finally { $zip.Dispose() }
}

# numbering.xml: mapa numId -> format del nivell 0 (decimal, bullet, lowerLetter...).
function _BuildNumFmtMap($docxPath) {
    $map = @{}
    $txt = _ReadDocxPartText $docxPath 'word/numbering.xml'
    if ($null -eq $txt) { return $map }
    try {
        $xml = New-Object System.Xml.XmlDocument; $xml.PreserveWhitespace = $true; $xml.LoadXml($txt)
        $ns = _NewWordNsMgr $xml
        $absFmt = @{}
        foreach ($abs in $xml.SelectNodes('//w:abstractNum', $ns)) {
            $aid = $abs.GetAttribute('abstractNumId', $Script:WNS)
            foreach ($lvl in $abs.SelectNodes('w:lvl', $ns)) {
                if ($lvl.GetAttribute('ilvl', $Script:WNS) -eq '0') {
                    $fmtEl = $lvl.SelectSingleNode('w:numFmt', $ns)
                    if ($null -ne $fmtEl) { $absFmt[$aid] = $fmtEl.GetAttribute('val', $Script:WNS) }
                    break
                }
            }
        }
        foreach ($num in $xml.SelectNodes('//w:num', $ns)) {
            $nid = $num.GetAttribute('numId', $Script:WNS)
            $aidEl = $num.SelectSingleNode('w:abstractNumId', $ns)
            if ($null -ne $aidEl) {
                $aid = $aidEl.GetAttribute('val', $Script:WNS)
                if ($absFmt.ContainsKey($aid)) { $map[$nid] = $absFmt[$aid] }
            }
        }
    } catch { }
    return $map
}

# styles.xml: mapa styleId -> { NumId; Ilvl } efectiu (numeracio que aporta
# l'estil, seguint la cadena basedOn). Es el cas dels informes on el numero del
# requeriment ve de l'estil "Prrafodelista" (List Paragraph) i no del paragraf.
function _BuildStyleNumMap($docxPath) {
    $map = @{}
    $txt = _ReadDocxPartText $docxPath 'word/styles.xml'
    if ($null -eq $txt) { return $map }
    try {
        $xml = New-Object System.Xml.XmlDocument; $xml.PreserveWhitespace = $true; $xml.LoadXml($txt)
        $ns = _NewWordNsMgr $xml
        $direct = @{}; $basedOn = @{}
        foreach ($st in $xml.SelectNodes('//w:style', $ns)) {
            $sid = $st.GetAttribute('styleId', $Script:WNS)
            if ([string]::IsNullOrEmpty($sid)) { continue }
            $np = $st.SelectSingleNode('w:pPr/w:numPr', $ns)
            if ($null -ne $np) {
                $nidEl  = $np.SelectSingleNode('w:numId', $ns)
                $ilvlEl = $np.SelectSingleNode('w:ilvl', $ns)
                $nid  = if ($null -ne $nidEl)  { $nidEl.GetAttribute('val', $Script:WNS) } else { '0' }
                $ilvl = if ($null -ne $ilvlEl) { [int]$ilvlEl.GetAttribute('val', $Script:WNS) } else { 0 }
                $direct[$sid] = @{ NumId = $nid; Ilvl = $ilvl }
            }
            $boEl = $st.SelectSingleNode('w:basedOn', $ns)
            if ($null -ne $boEl) { $basedOn[$sid] = $boEl.GetAttribute('val', $Script:WNS) }
        }
        $allIds = New-Object System.Collections.Generic.HashSet[string]
        foreach ($k in $direct.Keys)  { [void]$allIds.Add($k) }
        foreach ($k in $basedOn.Keys) { [void]$allIds.Add($k) }
        foreach ($sid in $allIds) {
            $cur = $sid; $depth = 0
            while ($null -ne $cur -and $cur -ne '' -and $depth -lt 10) {
                if ($direct.ContainsKey($cur)) { $map[$sid] = $direct[$cur]; break }
                $cur = if ($basedOn.ContainsKey($cur)) { $basedOn[$cur] } else { $null }
                $depth++
            }
        }
    } catch { }
    return $map
}

# Carrega word/document.xml com a XmlDocument. Retorna { Path; Xml; Ns; Body;
# NumFmt; StyleNum } (els dos darrers per resoldre la numeracio automatica).
function _LoadDocxXml($docxPath) {
    $text = _ReadDocxPartText $docxPath 'word/document.xml'
    if ($null -eq $text) { throw "El fitxer no sembla un .docx valid (falta word/document.xml)." }
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($text)
    $ns = _NewWordNsMgr $xml
    $body = $xml.SelectSingleNode('//w:body', $ns)
    if ($null -eq $body) { throw "document.xml sense <w:body>." }
    return [pscustomobject]@{
        Path     = $docxPath
        Xml      = $xml
        Ns       = $ns
        Body     = $body
        NumFmt   = (_BuildNumFmtMap $docxPath)
        StyleNum = (_BuildStyleNumMap $docxPath)
    }
}

# Desa: copia el .docx origen a $outPath i hi reescriu word/document.xml.
# La resta de parts (estils, capcalera, rels...) queden intactes.
function _SaveDocxXml($xmlInfo, $srcPath, $outPath) {
    Copy-Item -LiteralPath $srcPath -Destination $outPath -Force
    try { (Get-Item -LiteralPath $outPath).IsReadOnly = $false } catch { }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
    $zip = [System.IO.Compression.ZipFile]::Open($outPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) { $entry = $zip.CreateEntry('word/document.xml') }
        $stream = $entry.Open()
        try {
            $stream.SetLength(0)
            $xmlInfo.Xml.Save($stream)
        } finally { $stream.Close() }
    } finally { $zip.Dispose() }
}

# Paragrafs <w:p> que son fills DIRECTES del <w:body> (en ordre). Els paragrafs
# dins de taules (capcalera) NO hi son: queden fora de la deteccio, com volem.
function _BodyParagraphsXml($xmlInfo) {
    $list = New-Object System.Collections.ArrayList
    foreach ($child in $xmlInfo.Body.ChildNodes) {
        if ($child.LocalName -eq 'p' -and $child.NamespaceURI -eq $Script:WNS) {
            [void]$list.Add($child)
        }
    }
    return $list.ToArray()
}

# Text d'un paragraf: concatena w:t, tabula w:tab i tracta w:br com a espai,
# en ORDRE de document (important per a "1.<tab>text" -> "1.\ttext").
function _ParagraphTextXml($p, $ns) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in $p.SelectNodes('.//w:t | .//w:tab | .//w:br', $ns)) {
        switch ($n.LocalName) {
            't'   { [void]$sb.Append($n.InnerText) }
            'tab' { [void]$sb.Append("`t") }
            'br'  { [void]$sb.Append(' ') }
        }
    }
    return $sb.ToString()
}

# Un run <w:r> esta en negreta? (w:rPr/w:b sense val o amb val != 0/false).
function _RunIsBoldXml($r, $ns) {
    $b = $r.SelectSingleNode('w:rPr/w:b', $ns)
    if ($null -eq $b) { return $false }
    $val = $b.GetAttribute('val', $Script:WNS)
    if ([string]::IsNullOrEmpty($val)) { return $true }
    return ($val -ne '0' -and $val -ne 'false')
}

# Estat de negreta del paragraf, en l'estil Word.Font.Bold:
#   0 (cap run de text en negreta) / -1 (tots) / 9999999 (mixt).
function _ParagraphBoldStateXml($p, $ns) {
    $withText = New-Object System.Collections.ArrayList
    foreach ($r in $p.SelectNodes('w:r', $ns)) {
        $t = $r.SelectSingleNode('w:t', $ns)
        if ($null -ne $t -and -not [string]::IsNullOrEmpty($t.InnerText)) { [void]$withText.Add($r) }
    }
    if ($withText.Count -eq 0) { return 0 }
    $bold = 0
    foreach ($r in $withText) { if (_RunIsBoldXml $r $ns) { $bold++ } }
    if ($bold -eq 0) { return 0 }
    if ($bold -eq $withText.Count) { return -1 }
    return 9999999
}

# Numeracio EFECTIVA d'un paragraf: { NumId; Ilvl }. Prioritat: numPr en linia
# al paragraf; si no n'hi ha, el numPr que aporti el seu estil (pStyle). numId='0'
# vol dir "sense numeracio". $styleNumMap pot ser $null (documents sintetics).
function _EffectiveListInfo($p, $ns, $styleNumMap) {
    $inline = $p.SelectSingleNode('w:pPr/w:numPr', $ns)
    if ($null -ne $inline) {
        $nidEl  = $inline.SelectSingleNode('w:numId', $ns)
        $ilvlEl = $inline.SelectSingleNode('w:ilvl', $ns)
        $nid  = if ($null -ne $nidEl)  { $nidEl.GetAttribute('val', $Script:WNS) } else { '0' }
        $ilvl = if ($null -ne $ilvlEl) { [int]$ilvlEl.GetAttribute('val', $Script:WNS) } else { 0 }
        return @{ NumId = $nid; Ilvl = $ilvl }
    }
    $pStyleEl = $p.SelectSingleNode('w:pPr/w:pStyle', $ns)
    if ($null -ne $pStyleEl -and $null -ne $styleNumMap) {
        $sid = $pStyleEl.GetAttribute('val', $Script:WNS)
        if ($styleNumMap.ContainsKey($sid)) { return $styleNumMap[$sid] }
    }
    return @{ NumId = '0'; Ilvl = 0 }
}

# Registres { Index; Text; ListString; Bold } a partir dels paragrafs del body.
# Per a la numeracio AUTOMATICA del Word (numero no escrit al text, sino aportat
# pel numPr del paragraf o del seu estil), es detecta una llista DECIMAL de
# nivell 0 i s'omple ListString amb el numero corresponent ("1.", "2."...), de
# manera que _ClassifyParagraph el tracti com a requeriment igual que si el
# numero fos text literal.
function _CollectParaRecordsXml($xmlInfo, $bodyParas) {
    $ns        = $xmlInfo.Ns
    $numFmt    = $xmlInfo.NumFmt
    $styleNum  = $xmlInfo.StyleNum
    $records = New-Object System.Collections.ArrayList
    $autoCounter = 0
    for ($i = 0; $i -lt $bodyParas.Count; $i++) {
        $p = $bodyParas[$i]
        $listStr = ''
        $li = _EffectiveListInfo $p $ns $styleNum
        if ($li.NumId -ne '0' -and $li.Ilvl -eq 0) {
            $fmt = if ($null -ne $numFmt -and $numFmt.ContainsKey($li.NumId)) { $numFmt[$li.NumId] } else { 'decimal' }
            if ($fmt -eq 'decimal') {
                $autoCounter++
                $listStr = "$autoCounter."
            }
        }
        [void]$records.Add([pscustomobject]@{
            Index      = ($i + 1)
            Text       = (_ParagraphTextXml $p $ns)
            ListString = $listStr
            Bold       = (_ParagraphBoldStateXml $p $ns)
        })
    }
    return $records.ToArray()
}

# Afegeix la font del cos (Bookman Old Style 11) a un <w:rPr>. L'ordre dins rPr
# (rFonts abans de sz) es respecta perque despres _SetParagraphBoldXml insereix
# b/bCs JUST despres de rFonts (ordre valid de CT_RPr).
function _ApplyBodyFontXml($xml, $rPr) {
    $w = $Script:WNS
    $rf = $xml.CreateElement('w','rFonts',$w)
    [void]$rf.SetAttribute('ascii',$w,$Script:SeguimentFontName)
    [void]$rf.SetAttribute('hAnsi',$w,$Script:SeguimentFontName)
    [void]$rf.SetAttribute('cs',$w,$Script:SeguimentFontName)
    [void]$rPr.AppendChild($rf)
    $sz   = $xml.CreateElement('w','sz',$w);   [void]$sz.SetAttribute('val',$w,$Script:SeguimentFontHalfPt);   [void]$rPr.AppendChild($sz)
    $szCs = $xml.CreateElement('w','szCs',$w); [void]$szCs.SetAttribute('val',$w,$Script:SeguimentFontHalfPt); [void]$rPr.AppendChild($szCs)
}

# Un paragraf es d'enllac (URL) si conte un <w:hyperlink> o el text es una URL.
function _IsUrlParagraphXml($p, $ns) {
    if ($null -ne $p.SelectSingleNode('.//w:hyperlink', $ns)) { return $true }
    return ((_ParagraphTextXml $p $ns).Trim() -match '^https?://')
}

# Posa/treu negreta a TOTS els runs amb text d'un paragraf, mantenint l'ordre
# valid de CT_RPr (b/bCs just despres de rStyle/rFonts).
#   $boldOn = $true  -> <w:b/> i <w:bCs/>
#   $boldOn = $false -> <w:b w:val="false"/> i <w:bCs w:val="false"/> (anul·la
#                       fins i tot la negreta que ja portava, p.ex. el "N.").
function _SetParagraphBoldXml($xmlInfo, $p, [bool]$boldOn) {
    $xml = $xmlInfo.Xml; $ns = $xmlInfo.Ns; $w = $Script:WNS
    foreach ($r in $p.SelectNodes('w:r', $ns)) {
        if ($null -eq $r.SelectSingleNode('w:t', $ns)) { continue }
        $rPr = $r.SelectSingleNode('w:rPr', $ns)
        if ($null -eq $rPr) {
            $rPr = $xml.CreateElement('w','rPr',$w)
            [void]$r.PrependChild($rPr)
        }
        # Treure b/bCs existents per reinserir-los en la posicio correcta.
        foreach ($old in @($rPr.SelectNodes('w:b', $ns)) + @($rPr.SelectNodes('w:bCs', $ns))) {
            [void]$rPr.RemoveChild($old)
        }
        $b   = $xml.CreateElement('w','b',$w)
        $bCs = $xml.CreateElement('w','bCs',$w)
        if (-not $boldOn) {
            [void]$b.SetAttribute('val', $w, 'false')
            [void]$bCs.SetAttribute('val', $w, 'false')
        }
        $anchor = $null
        foreach ($name in 'w:rFonts','w:rStyle') {
            $x = $rPr.SelectSingleNode($name, $ns)
            if ($null -ne $x) { $anchor = $x; break }
        }
        if ($null -ne $anchor) {
            [void]$rPr.InsertAfter($bCs, $anchor)
            [void]$rPr.InsertAfter($b, $anchor)   # queda: anchor, b, bCs
        } else {
            [void]$rPr.PrependChild($bCs)
            [void]$rPr.PrependChild($b)
        }
    }
}

# Crea un run <w:r> en Bookman Old Style 11, opcionalment en negreta, amb l'ordre
# valid de CT_RPr (rFonts, b, bCs, sz, szCs).
function _MakeBodyRunXml($xml, $text, [bool]$bold) {
    $w = $Script:WNS
    $r = $xml.CreateElement('w','r',$w)
    $rPr = $xml.CreateElement('w','rPr',$w)
    $rf = $xml.CreateElement('w','rFonts',$w)
    [void]$rf.SetAttribute('ascii',$w,$Script:SeguimentFontName)
    [void]$rf.SetAttribute('hAnsi',$w,$Script:SeguimentFontName)
    [void]$rf.SetAttribute('cs',$w,$Script:SeguimentFontName)
    [void]$rPr.AppendChild($rf)
    if ($bold) {
        [void]$rPr.AppendChild($xml.CreateElement('w','b',$w))
        [void]$rPr.AppendChild($xml.CreateElement('w','bCs',$w))
    }
    $sz   = $xml.CreateElement('w','sz',$w);   [void]$sz.SetAttribute('val',$w,$Script:SeguimentFontHalfPt);   [void]$rPr.AppendChild($sz)
    $szCs = $xml.CreateElement('w','szCs',$w); [void]$szCs.SetAttribute('val',$w,$Script:SeguimentFontHalfPt); [void]$rPr.AppendChild($szCs)
    [void]$r.AppendChild($rPr)
    $t = $xml.CreateElement('w','t',$w)
    $xsp=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace'); $xsp.Value='preserve'; [void]$t.Attributes.Append($xsp)
    $t.InnerText = [string]$text
    [void]$r.AppendChild($t)
    return ,$r
}

# Crea un <w:p> d'anotacio "dd/MM/aaaa: comentari". El PREFIX de la data va sense
# negreta; NOMES el comentari va en negreta si $commentBold (= punt encara no
# resolt). Clona el pPr del requeriment (estil/sagnat) forcant numId=0.
function _MakeAnnotationParagraphXml($xmlInfo, $reqNode, $dateStr, $comment, [bool]$commentBold, [bool]$spaceBefore = $false) {
    $xml = $xmlInfo.Xml; $ns = $xmlInfo.Ns; $w = $Script:WNS
    $p = $xml.CreateElement('w','p',$w)
    $pPr = $null
    $reqPPr = $reqNode.SelectSingleNode('w:pPr', $ns)
    if ($null -ne $reqPPr) {
        $pPr = $reqPPr.CloneNode($true)
        $pmRPr = $pPr.SelectSingleNode('w:rPr', $ns)   # format de la marca de paragraf
        if ($null -ne $pmRPr) { [void]$pPr.RemoveChild($pmRPr) }
        # FORCAR numId=0 perque NO s'enumeri (mantenint l'estil de llista, que
        # aporta la sagnia per alinear l'anotacio sota el requeriment).
        $numPr = $pPr.SelectSingleNode('w:numPr', $ns)
        if ($null -eq $numPr) {
            $numPr = $xml.CreateElement('w','numPr',$w)
            $ilvl = $xml.CreateElement('w','ilvl',$w);  [void]$ilvl.SetAttribute('val',$w,'0'); [void]$numPr.AppendChild($ilvl)
            $nid  = $xml.CreateElement('w','numId',$w); [void]$nid.SetAttribute('val',$w,'0');  [void]$numPr.AppendChild($nid)
            $pStyle = $pPr.SelectSingleNode('w:pStyle', $ns)
            if ($null -ne $pStyle) { [void]$pPr.InsertAfter($numPr, $pStyle) } else { [void]$pPr.PrependChild($numPr) }
        } else {
            $nid = $numPr.SelectSingleNode('w:numId', $ns)
            if ($null -eq $nid) { $nid = $xml.CreateElement('w','numId',$w); [void]$numPr.AppendChild($nid) }
            [void]$nid.SetAttribute('val', $w, '0')
        }
    } else {
        $pPr = $xml.CreateElement('w','pPr',$w)
    }
    # Espai a sobre (separacio visual amb el cos de l'item). Es fa amb spacing
    # (no amb un paragraf buit) per no trencar la deteccio de fi de cos.
    if ($spaceBefore) {
        $sp = $pPr.SelectSingleNode('w:spacing', $ns)
        if ($null -eq $sp) {
            $sp = $xml.CreateElement('w','spacing',$w)
            $numPr2 = $pPr.SelectSingleNode('w:numPr', $ns)
            if ($null -ne $numPr2) { [void]$pPr.InsertAfter($sp, $numPr2) } else { [void]$pPr.AppendChild($sp) }
        }
        [void]$sp.SetAttribute('before', $w, '200')
    }
    [void]$p.AppendChild($pPr)
    # Run 1: "dd/MM/aaaa: " (mai negreta). Run 2: comentari (negreta si pendent).
    [void]$p.AppendChild((_MakeBodyRunXml $xml ('{0}: ' -f $dateStr) $false))
    [void]$p.AppendChild((_MakeBodyRunXml $xml ([string]$comment) $commentBold))
    return ,$p   # ,: el node <w:p> es IEnumerable; evitem que s'enumeri
}

# Un paragraf es una SUBSECCIO si te algun run subratllat (Format-Subsection).
# El subratllat es un senyal estable: el seguiment no l'afegeix ni el treu mai.
function _IsSubsectionXml($p, $ns) {
    $u = $p.SelectSingleNode('.//w:rPr/w:u', $ns)
    if ($null -eq $u) { return $false }
    return ($u.GetAttribute('val', $Script:WNS) -ne 'none')
}

# Calcula, per cada requeriment, el seu BLOC = el requeriment + el seu cos
# (sub-linies, enllac, anotacions previes) que el segueixen de manera CONSECUTIVA
# sense paragraf buit. El cos d'un item acaba al PRIMER paragraf buit (l'espaiador
# que el generador posa despres de cada item), o a una subseccio (subratllat) o al
# requeriment seguent. Aixi NO s'inclouen seccions/subseccions (que venen despres
# de l'espaiador). Retorna, per requeriment:
#   ReqNode            : el paragraf del requeriment.
#   AnchorNode         : ultim paragraf NO buit del cos (on inserir la nova anotacio,
#                        "a sota de tot": despres del cos, l'enllac i anotacions previes).
#   AnchorIsAnnotation : si l'ancora ja es una anotacio datada (re-seguiment).
#   ContentNodes       : paragrafs a posar en negreta si pendent (requeriment +
#                        sub-linies + anotacions). MAI l'enllac.
function _SeguimentBlocksXml($bodyParas, $model, [int]$conclusionStartIndex, $ns) {
    $reqs = $model.Requirements
    $hardEnd = if ($conclusionStartIndex -ge 1) { $conclusionStartIndex - 1 } else { $bodyParas.Count }
    $blocks = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $reqs.Count; $k++) {
        $startIdx = [int]$reqs[$k].ParaIndex                      # 1-based
        $nextReq  = if ($k -lt $reqs.Count - 1) { [int]$reqs[$k+1].ParaIndex } else { [int]::MaxValue }

        $reqNode = $bodyParas[$startIdx - 1]
        $anchor  = $reqNode
        $content = New-Object System.Collections.ArrayList
        [void]$content.Add($reqNode)                              # el requeriment sempre (per la negreta)

        for ($i = $startIdx + 1; ($i -le $hardEnd) -and ($i -lt $nextReq); $i++) {
            $node = $bodyParas[$i - 1]
            $txt  = (_ParagraphTextXml $node $ns)
            if ([string]::IsNullOrWhiteSpace($txt))           { break }   # espaiador -> fi del cos
            if ($Script:SeguimentReqRegex.IsMatch($txt.Trim())) { break } # seguent requeriment
            if (_IsSubsectionXml $node $ns)                   { break }   # subseccio
            $anchor = $node                                                # ultim no buit del cos
            if (-not (_IsUrlParagraphXml $node $ns)) { [void]$content.Add($node) }
        }

        $anchorIsAnnot = $Script:SeguimentAnnotRegex.IsMatch((_ParagraphTextXml $anchor $ns).Trim())
        [void]$blocks.Add([pscustomobject]@{
            ReqIndex           = $k
            ReqNode            = $reqNode
            AnchorNode         = $anchor
            AnchorIsAnnotation = $anchorIsAnnot
            ContentNodes       = $content.ToArray()
        })
    }
    return $blocks.ToArray()
}

# Runs <w:r> a partir d'un text interpretant **negreta** i //cursiva//.
function _RichTextRunsXml($xmlInfo, $text) {
    $xml = $xmlInfo.Xml; $w = $Script:WNS
    $runs = New-Object System.Collections.ArrayList
    $make = {
        param($segment, $bold, $italic)
        if ([string]::IsNullOrEmpty($segment)) { return }
        $r = $xml.CreateElement('w','r',$w)
        $rPr = $xml.CreateElement('w','rPr',$w)
        _ApplyBodyFontXml $xml $rPr
        if ($bold)   { [void]$rPr.AppendChild($xml.CreateElement('w','b',$w)) }
        if ($italic) { [void]$rPr.AppendChild($xml.CreateElement('w','i',$w)) }
        [void]$r.AppendChild($rPr)
        $t = $xml.CreateElement('w','t',$w)
        $xsp=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace'); $xsp.Value='preserve'; [void]$t.Attributes.Append($xsp)
        $t.InnerText = $segment
        [void]$r.AppendChild($t)
        [void]$runs.Add($r)
    }
    $rx = [regex]'\*\*(.+?)\*\*|//(.+?)//'
    $pos = 0
    foreach ($m in $rx.Matches([string]$text)) {
        if ($m.Index -gt $pos) { & $make ([string]$text).Substring($pos, $m.Index - $pos) $false $false }
        if     ($m.Groups[1].Success) { & $make $m.Groups[1].Value $true  $false }
        elseif ($m.Groups[2].Success) { & $make $m.Groups[2].Value $false $true  }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt ([string]$text).Length) { & $make ([string]$text).Substring($pos) $false $false }
    return ,($runs.ToArray())
}

# Crea un <w:p> per a una conclusio. Si $centeredBold, centrat i en negreta
# (titol del bloc); si no, text normal amb **negreta**/​//cursiva// inline.
function _MakeConclusionParagraphXml($xmlInfo, $text, [bool]$centeredBold) {
    $xml = $xmlInfo.Xml; $w = $Script:WNS
    $p = $xml.CreateElement('w','p',$w)
    if ($centeredBold) {
        $pPr = $xml.CreateElement('w','pPr',$w)
        $jc = $xml.CreateElement('w','jc',$w); [void]$jc.SetAttribute('val',$w,'center'); [void]$pPr.AppendChild($jc)
        [void]$p.AppendChild($pPr)
        $r = $xml.CreateElement('w','r',$w)
        $rPr = $xml.CreateElement('w','rPr',$w)
        _ApplyBodyFontXml $xml $rPr
        [void]$rPr.AppendChild($xml.CreateElement('w','b',$w))
        [void]$r.AppendChild($rPr)
        $t = $xml.CreateElement('w','t',$w); $xsp=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace'); $xsp.Value='preserve'; [void]$t.Attributes.Append($xsp); $t.InnerText = [string]$text
        [void]$r.AppendChild($t); [void]$p.AppendChild($r)
    } else {
        # Paragraf de conclusio: justificat (com l'estil Normal de REQ1).
        $pPr = $xml.CreateElement('w','pPr',$w)
        $jc = $xml.CreateElement('w','jc',$w); [void]$jc.SetAttribute('val',$w,'both'); [void]$pPr.AppendChild($jc)
        [void]$p.AppendChild($pPr)
        foreach ($r in (_RichTextRunsXml $xmlInfo $text)) { [void]$p.AppendChild($r) }
    }
    return ,$p
}

# Afegeix els paragrafs de conclusions al final del cos, JUST ABANS del
# <w:sectPr> (si existeix), resolent [CAMP:]/[OPCIO:] amb Apply-Fields.
function _AppendConclusionParagraphsXml($xmlInfo, $headerText, $conclusions, $always, $fields) {
    $ns = $xmlInfo.Ns; $body = $xmlInfo.Body
    $sectPr = $body.SelectSingleNode('w:sectPr', $ns)
    $append = {
        param($node)
        if ($null -ne $sectPr) { [void]$body.InsertBefore($node, $sectPr) }
        else                   { [void]$body.AppendChild($node) }
    }
    if (-not [string]::IsNullOrWhiteSpace($headerText)) {
        & $append (_MakeConclusionParagraphXml $xmlInfo $headerText $true)
    }
    foreach ($c in $conclusions) {
        $txt = if ($c -is [string]) { $c } else { [string]$c.Body }
        $resolved = Apply-Fields -text $txt -fields $fields
        & $append (_MakeConclusionParagraphXml $xmlInfo $resolved $false)
    }
    foreach ($a in $always) {
        $resolved = Apply-Fields -text ([string]$a) -fields $fields
        & $append (_MakeConclusionParagraphXml $xmlInfo $resolved $false)
    }
}

# Llegeix 0 CONCLUSIONS.docx via XML (equivalent a Read-Conclusions pero sense
# Word). Retorna { HeaderText; Selectable=@({Title;Body}); Always }.
# Dedueix el tipus d'informe (REQ1, TERMINI...) a partir del nom d'un informe
# generat. El motor anomena els fitxers "YYYY-MM-DD_<Tipus>_GIA <id>.docx"
# (vegeu _GetOutputFileName), aixi que el tipus es el 2n segment separat per '_'.
# Retorna '' si no es pot deduir (llavors s'ofereixen totes les conclusions).
function _ReportTypeFromFileName($fileName) {
    if ([string]::IsNullOrWhiteSpace($fileName)) { return '' }
    $stem  = [System.IO.Path]::GetFileNameWithoutExtension([string]$fileName)
    $parts = $stem -split '_'
    if ($parts.Count -ge 2) { return $parts[1].Trim() }
    return ''
}

function Read-ConclusionsXml($path, $reportType = $null) {
    # Versio XML (sense Word) de Read-Conclusions. Mateixa semantica, incloent
    # els grups per tipus d'informe: Ttulo1 (Heading 1) = tipus d'informe,
    # Ttulo2 (Heading 2) = titol de conclusio triable. $reportType buit/null
    # retorna TOTES les conclusions (de tots els grups).
    $empty = [pscustomobject]@{ HeaderText=''; Selectable=@(); Always=@() }
    if (-not (Test-Path -LiteralPath $path)) { return $empty }

    $xmlInfo   = _LoadDocxXml $path
    $ns        = $xmlInfo.Ns
    $bodyParas = @(_BodyParagraphsXml $xmlInfo)

    $wantType = if ([string]::IsNullOrWhiteSpace($reportType)) { '' } else { _NormalizeText $reportType }

    $headerText   = ''
    $selectable   = New-Object System.Collections.ArrayList
    $always       = New-Object System.Collections.ArrayList
    $pendingTitle = $null
    $inGroup      = ($wantType -eq '')
    $isFirst      = $true

    foreach ($p in $bodyParas) {
        $text = (_ParagraphTextXml $p $ns).TrimEnd("`r","`n","`t"," ")
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $styleVal = ''
        $pStyle = $p.SelectSingleNode('w:pPr/w:pStyle', $ns)
        if ($null -ne $pStyle) { $styleVal = $pStyle.GetAttribute('val', $Script:WNS) }
        $isH1 = Test-StyleMatch $styleVal 1
        $isH2 = Test-StyleMatch $styleVal 2

        if ($isFirst -and -not $isH1 -and -not $isH2) {
            $isFirst = $false
            $jc = $p.SelectSingleNode('w:pPr/w:jc', $ns)
            $centered = ($null -ne $jc -and $jc.GetAttribute('val', $Script:WNS) -eq 'center')
            if ($centered) { $headerText = $text; continue }
        }
        $isFirst = $false

        if ($isH1) {
            # Nou grup (tipus d'informe).
            $inGroup = ($wantType -eq '') -or ((_NormalizeText $text) -eq $wantType)
            $pendingTitle = $null
            continue
        }
        if ($isH2) { $pendingTitle = $text; continue }

        if ($text.StartsWith('::SEMPRE::')) {
            [void]$always.Add($text.Substring('::SEMPRE::'.Length).Trim())
            $pendingTitle = $null
            continue
        }
        if ($null -ne $pendingTitle) {
            if ($inGroup) {
                [void]$selectable.Add([pscustomobject]@{ Title=$pendingTitle; Body=$text })
            }
            $pendingTitle = $null
        }
    }

    return [pscustomobject]@{
        HeaderText = $headerText
        Selectable = $selectable.ToArray()
        Always     = $always.ToArray()
    }
}

# Transforma el XML carregat (en memoria, sense E/S): esborra conclusions,
# insereix anotacions, recalcula negreta i afegeix les conclusions noves.
# Treballa amb REFERENCIES A NODES (no indexs): inserir/esborrar no invalida res.
#   $bodyParas : array de nodes <w:p> (1-based via index) del moment de la lectura.
#   $decisions : array alineat amb $model.Requirements; { Resolved; NewComment }.
#   $conclusionStartIndex : index 1-based del primer paragraf a esborrar, o -1.
function _ApplySeguimentTransform {
    param(
        $xmlInfo, $bodyParas, $model, [int]$conclusionStartIndex,
        $decisions, $dateStr,
        $conclHeaderText, $selectedConclusions, $alwaysConclusions, $fields
    )

    # Calculem els blocs ABANS de tocar res (anclatge i nodes de contingut).
    $blocks = @(_SeguimentBlocksXml $bodyParas $model $conclusionStartIndex $xmlInfo.Ns)

    # (1) Esborrar el bloc de conclusions (nodes <w:p> des del cut fins al final).
    # El <w:sectPr> (fill del body, no es <w:p>) es preserva automaticament.
    if ($conclusionStartIndex -ge 1) {
        for ($i = $bodyParas.Count - 1; $i -ge ($conclusionStartIndex - 1); $i--) {
            $node = $bodyParas[$i]
            if ($null -ne $node.ParentNode) { [void]$node.ParentNode.RemoveChild($node) }
        }
    }

    # (2) Per cada requeriment: afegir la nova anotacio A SOTA DE TOT. NOMES el
    # comentari va en negreta i nomes si el punt queda PENDENT (no resolt). El
    # requeriment, la data i el text NO es toquen. Les anotacions ANTERIORS es
    # des-negreten (historic): nomes l'ultima entrega pendent queda destacada.
    for ($k = 0; $k -lt $blocks.Count; $k++) {
        $b   = $blocks[$k]
        $dec = $decisions[$k]
        $req = $model.Requirements[$k]

        # Des-negretar les anotacions previes (deixen de ser "l'ultima pendent").
        foreach ($a in $req.Annotations) {
            $annNode = $bodyParas[$a.ParaIndex - 1]
            if ($null -ne $annNode) { _SetParagraphBoldXml $xmlInfo $annNode $false }
        }

        $wasResolved = [bool]$req.WasResolved
        $nowResolved = [bool]$dec.Resolved
        $comment     = [string]$dec.NewComment

        # Si ja estava resolt I segueix resolt -> NO cal reescriure res (evita
        # "S'aporta." repetit en rondes posteriors).
        $addLine = (-not ($wasResolved -and $nowResolved)) -and (-not [string]::IsNullOrWhiteSpace($comment))
        if ($addLine) {
            $commentBold = (-not $nowResolved)                 # negreta nomes si pendent
            $spaceBefore = (-not $b.AnchorIsAnnotation)         # espai si es la 1a anotacio
            $newP = _MakeAnnotationParagraphXml $xmlInfo $b.ReqNode $dateStr $comment $commentBold $spaceBefore
            [void]$xmlInfo.Body.InsertAfter($newP, $b.AnchorNode)
        }
    }

    # (3) Afegir les conclusions noves al final.
    _AppendConclusionParagraphsXml $xmlInfo $conclHeaderText $selectedConclusions $alwaysConclusions $fields
}

# Aplica tot el seguiment sobre el XML carregat i el desa a un fitxer nou.
# Retorna la ruta final.
function Apply-SeguimentXml {
    param(
        $xmlInfo, $bodyParas, $model, [int]$conclusionStartIndex,
        $decisions, $dateStr,
        $conclHeaderText, $selectedConclusions, $alwaysConclusions, $fields,
        $srcPath, [string]$sourceBaseName, [datetime]$roundDate
    )

    _ApplySeguimentTransform -xmlInfo $xmlInfo -bodyParas $bodyParas -model $model `
        -conclusionStartIndex $conclusionStartIndex -decisions $decisions -dateStr $dateStr `
        -conclHeaderText $conclHeaderText -selectedConclusions $selectedConclusions `
        -alwaysConclusions $alwaysConclusions -fields $fields

    # La data del nom es la d'AVUI (dia de generacio), no la de l'anotacio.
    $outName   = _SeguimentOutputName $sourceBaseName (Get-Date)
    $targetDir = _ResolveOutputDir
    $outPath   = _GetUniqueOutputPath $targetDir $outName
    _SaveDocxXml $xmlInfo $srcPath $outPath
    return $outPath
}

# ----------------------------------------------------------------------------
# CAPA WINFORMS - dialegs del flux de seguiment (nomes Windows)
# ----------------------------------------------------------------------------

# Pantalla inicial: tria entre generar un informe nou o fer un seguiment.
# Retorna 'nou' | 'seguiment'. Tancar la finestra (X) avorta (exit 0).
function Select-Mode {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Informes Cornella'
    $form.Size = New-Object System.Drawing.Size(440, 290)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Que vols fer?'
    $lbl.Location = New-Object System.Drawing.Point(20, 15)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $btnNou = New-Object System.Windows.Forms.Button
    $btnNou.Text = 'Generar informe nou'
    $btnNou.Location = New-Object System.Drawing.Point(20, 50)
    $btnNou.Size = New-Object System.Drawing.Size(390, 45)
    $btnNou.DialogResult = 'Yes'
    $form.Controls.Add($btnNou)

    $btnSeg = New-Object System.Windows.Forms.Button
    $btnSeg.Text = "Fer seguiment d'un informe existent"
    $btnSeg.Location = New-Object System.Drawing.Point(20, 105)
    $btnSeg.Size = New-Object System.Drawing.Size(390, 45)
    $btnSeg.DialogResult = 'No'
    $form.Controls.Add($btnSeg)

    # Tercer mode: activitats extraordinaries (Decret 112/2010). Retorna
    # DialogResult 'Retry' perque Main el dirigeixi a Invoke-ActExtrFlow.
    $btnActExtr = New-Object System.Windows.Forms.Button
    $btnActExtr.Text = "Activitats extraordinaries (ACT_EXTR)"
    $btnActExtr.Location = New-Object System.Drawing.Point(20, 160)
    $btnActExtr.Size = New-Object System.Drawing.Size(390, 45)
    $btnActExtr.DialogResult = 'Retry'
    $form.Controls.Add($btnActExtr)

    $form.AcceptButton = $btnNou
    $res = $form.ShowDialog()
    if ($res -eq 'Yes')   { return 'nou' }
    if ($res -eq 'No')    { return 'seguiment' }
    if ($res -eq 'Retry') { return 'actextr' }
    exit 0
}

# Tria de l'informe anterior (.docx). Retorna la ruta o $null si es cancel·la.
function Select-PreviousReport {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Documents Word (*.docx)|*.docx'
    $dlg.Title  = "Tria l'informe anterior amb els requeriments"
    try {
        $od = _ResolveOutputDir
        if (Test-Path -LiteralPath $od) { $dlg.InitialDirectory = $od }
    } catch { }
    if ($dlg.ShowDialog() -ne 'OK') { return $null }
    return $dlg.FileName
}

# Tria del bloc de conclusions a esborrar. SEMPRE mostra el selector MANUAL
# (preferencia de l'usuari): es llisten els paragrafs a partir de l'ultim
# requeriment enumerat i l'usuari tria el primer a esborrar. Si la deteccio
# automatica ha trobat un punt d'inici, es preselecciona com a ajuda.
# Retorna { StartIndex } (1-based, o -1 per no esborrar); $null si cancel·la.
function Confirm-ConclusionDeletion {
    param($paraTexts, [int]$lastReqEndIndex, [int]$detectedStart)
    return (Select-ConclusionCutManually -paraTexts $paraTexts -lastReqEndIndex $lastReqEndIndex -preselectIndex $detectedStart)
}

# Selector manual del primer paragraf a esborrar. Llista els paragrafs a partir
# de l'ultim requeriment enumerat. $preselectIndex (1-based) es el paragraf
# preseleccionat (de la deteccio automatica), o -1 si no n'hi ha.
# Retorna { StartIndex } o { StartIndex = -1 } (no esborrar res); $null si cancel·la.
function Select-ConclusionCutManually {
    param($paraTexts, [int]$lastReqEndIndex, [int]$preselectIndex = -1)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Tria el primer paragraf a esborrar'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.MinimumSize = New-Object System.Drawing.Size(420, 300)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Tria el PRIMER paragraf del bloc de conclusions a esborrar (s'esborrara fins al final):"
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(680, 20)
    $lbl.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lbl)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(15, 35)
    $list.Size = New-Object System.Drawing.Size(680, 360)
    $list.Anchor = 'Top, Bottom, Left, Right'
    # Mapatge posicio-de-la-llista -> index 1-based de paragraf.
    $map = @()
    [void]$list.Items.Add('(No esborrar res)')
    $map += -1
    for ($i = $lastReqEndIndex + 1; $i -le $paraTexts.Count; $i++) {
        $tx = [string]$paraTexts[$i - 1]
        if ([string]::IsNullOrWhiteSpace($tx)) { continue }
        $disp = if ($tx.Length -gt 90) { $tx.Substring(0, 90) + '...' } else { $tx }
        [void]$list.Items.Add(('#{0}: {1}' -f $i, $disp))
        $map += $i
    }
    # Per defecte, el primer paragraf real; pero si hi ha un punt detectat, el
    # preseleccionem.
    $selPos = if ($list.Items.Count -gt 1) { 1 } else { 0 }
    if ($preselectIndex -ge 1) {
        for ($j = 0; $j -lt $map.Count; $j++) {
            if ($map[$j] -eq $preselectIndex) { $selPos = $j; break }
        }
    }
    $list.SelectedIndex = $selPos
    $form.Controls.Add($list)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(600, 405)
    $btnOk.Size = New-Object System.Drawing.Size(95, 30)
    $btnOk.Anchor = 'Bottom, Right'
    $btnOk.DialogResult = 'OK'
    $form.AcceptButton = $btnOk
    $form.Controls.Add($btnOk)

    $res = $form.ShowDialog()
    if ($res -ne 'OK') { return $null }
    return [pscustomobject]@{ StartIndex = [int]$map[$list.SelectedIndex] }
}

# Demana la data de la ronda (per defecte avui). Retorna { Nav; Data } amb Data
# en dd/MM/yyyy.
function Prompt-RoundDate {
    param($preset = $null)
    $default = if ($preset) { [string]$preset } else { (Get-Date).ToString('dd/MM/yyyy') }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data del seguiment'
    $form.Size = New-Object System.Drawing.Size(420, 200)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Data d'aquesta entrega/seguiment (dd/MM/aaaa):"
    $lbl.Location = New-Object System.Drawing.Point(15, 20)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(15, 50)
    $tb.Size = New-Object System.Drawing.Size(200, 24)
    $tb.Text = $default
    $form.Controls.Add($tb)

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 110)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(295, 110)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    while ($true) {
        $res = $form.ShowDialog()
        if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
        if ($res -ne 'OK')    { exit 0 }
        $v = _ValidateRoundDate $tb.Text
        if ($v.Ok) { return [pscustomobject]@{ Nav='next'; Data=$v.Normalized } }
        [System.Windows.Forms.MessageBox]::Show('Data no valida. Format esperat: dd/MM/aaaa.', 'Seguiment', 'OK', 'Warning') | Out-Null
    }
}

# Pas de comentaris: per cada requeriment, mostra el text + historial, un quadre
# de comentari nou i un checkbox "Resolt". Retorna { Nav; Data } amb Data un
# array alineat amb $requirements: { Resolved; NewComment }.
function Prompt-SeguimentComments {
    param($requirements, $dateStr, $preload = $null)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Seguiment dels requeriments'
    $form.Size = New-Object System.Drawing.Size(860, 640)
    $form.MinimumSize = New-Object System.Drawing.Size(560, 400)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = ("Anotacio del {0}. Escriu el comentari de cada requeriment i marca 'Resolt' si ha quedat resolt." -f $dateStr)
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(810, 20)
    $lbl.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 35)
    $panel.Size = New-Object System.Drawing.Size(815, 510)
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panel)

    $innerW = 760   # amplada dels controls dins el panell (deixa espai per la barra)
    $rows = @()
    $y = 8
    for ($i = 0; $i -lt $requirements.Count; $i++) {
        $req = $requirements[$i]

        # Etiqueta del requeriment: AutoSize amb amplada maxima -> ajusta l'alcada
        # al text complet (encara que ocupi diverses linies).
        $reqLbl = New-Object System.Windows.Forms.Label
        $reqLbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $reqLbl.AutoSize = $true
        $reqLbl.MaximumSize = New-Object System.Drawing.Size($innerW, 0)
        $reqLbl.Location = New-Object System.Drawing.Point(8, $y)
        $reqLbl.Text = [string]$req.Text
        $panel.Controls.Add($reqLbl)
        $y += $reqLbl.PreferredSize.Height + 4

        if ($req.Annotations.Count -gt 0) {
            # L'historial: cada anotacio ja inclou "data: comentari" (NO repetir la data).
            $hist = ($req.Annotations | ForEach-Object { [string]$_.Text }) -join "`r`n"
            $histLbl = New-Object System.Windows.Forms.Label
            $histLbl.AutoSize = $true
            $histLbl.MaximumSize = New-Object System.Drawing.Size($innerW, 0)
            $histLbl.Location = New-Object System.Drawing.Point(20, $y)
            $histLbl.ForeColor = [System.Drawing.Color]::DimGray
            $histLbl.Text = $hist
            $panel.Controls.Add($histLbl)
            $y += $histLbl.PreferredSize.Height + 4
        }

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.ScrollBars = 'Vertical'
        $tb.Location = New-Object System.Drawing.Point(20, $y)
        $tb.Size = New-Object System.Drawing.Size(($innerW - 12), 46)
        $panel.Controls.Add($tb)
        $y += 52

        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = 'Resolt'
        $cb.Location = New-Object System.Drawing.Point(20, $y)
        $cb.Size = New-Object System.Drawing.Size(120, 22)
        $cb.Checked = [bool]$req.WasResolved
        $panel.Controls.Add($cb)
        $y += 30

        # Comentari per defecte segons "Resolt": "S'aporta." / "No s'aporta.".
        $tb.Text = if ($cb.Checked) { $Script:SeguimentPhraseResolt } else { $Script:SeguimentPhraseNoResolt }
        # En canviar la casella, si el comentari encara es una frase automatica
        # (o buit), el commutem; si l'usuari l'ha editat, el respectem.
        $cb.Tag = $tb
        $cb.Add_CheckedChanged({
            $box = $this.Tag
            $cur = ([string]$box.Text).Trim()
            if ($cur -eq '' -or $cur -eq $Script:SeguimentPhraseResolt -or $cur -eq $Script:SeguimentPhraseNoResolt) {
                $box.Text = if ($this.Checked) { $Script:SeguimentPhraseResolt } else { $Script:SeguimentPhraseNoResolt }
            }
        })

        # Separador visual
        $sep = New-Object System.Windows.Forms.Label
        $sep.BorderStyle = 'Fixed3D'
        $sep.Location = New-Object System.Drawing.Point(8, $y)
        $sep.Size = New-Object System.Drawing.Size($innerW, 2)
        $panel.Controls.Add($sep)
        $y += 12

        # Precarrega (tornar enrere): restaura casella i, despres, el text guardat.
        if ($null -ne $preload -and $i -lt $preload.Count) {
            $cb.Checked = [bool]$preload[$i].Resolved
            $tb.Text    = [string]$preload[$i].NewComment
        }

        $rows += [pscustomobject]@{ Comment=$tb; Resolved=$cb }
    }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 555)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.Anchor = 'Bottom, Left'
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(735, 555)
    $ok.Size = New-Object System.Drawing.Size(95, 30)
    $ok.Anchor = 'Bottom, Right'
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    $decisions = New-Object System.Collections.ArrayList
    foreach ($r in $rows) {
        [void]$decisions.Add([pscustomobject]@{
            Resolved   = [bool]$r.Resolved.Checked
            NewComment = [string]$r.Comment.Text
        })
    }
    return [pscustomobject]@{ Nav='next'; Data=$decisions.ToArray() }
}

# ----------------------------------------------------------------------------
# Orquestrador del flux de seguiment (SENSE Word per editar).
# ----------------------------------------------------------------------------
function Invoke-SeguimentFlow {
    $sourcePath = Select-PreviousReport
    if (-not $sourcePath) { return }

    # Llista de frases que marquen l'inici del bloc de conclusions. Es pot
    # sobreescriure des de config.ps1 ($SeguimentConclusionPhrases).
    $phrases = if ($null -ne $SeguimentConclusionPhrases) { $SeguimentConclusionPhrases } else {
        @("Vist l'anterior", 'Ho poso al seu coneixement', 'Cornella de Llobregat,')
    }

    try {
        $xmlInfo   = _LoadDocxXml $sourcePath
        $bodyParas = @(_BodyParagraphsXml $xmlInfo)
        $records   = @(_CollectParaRecordsXml $xmlInfo $bodyParas)
        $model     = _BuildSeguimentModel $records
        $paraTexts = @($records | ForEach-Object { $_.Text })

        if ($model.Requirements.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No s'han trobat requeriments numerats (1., 2., 3...) en aquest document.",
                'Seguiment', 'OK', 'Warning') | Out-Null
            return
        }

        $detected = _FindConclusionStartIndex $paraTexts $model.LastReqParaIndex $phrases
        $cut = Confirm-ConclusionDeletion -paraTexts $paraTexts -lastReqEndIndex $model.LastReqParaIndex -detectedStart $detected
        if ($null -eq $cut) { return }
        $conclusionStartIndex = [int]$cut.StartIndex

        # Les conclusions triables depenen del tipus d'informe original, que
        # deduim del nom del fitxer de l'informe anterior (REQ1, TERMINI...).
        $reportType = _ReportTypeFromFileName ([System.IO.Path]::GetFileName($sourcePath))
        $conclAll = Read-ConclusionsXml $ConclusionsPath $reportType

        # Maquina de passos: 1=data, 2=comentaris, 3=conclusions (amb camps
        # inline). Les opcions/camps de les conclusions s'omplen dins del propi
        # text al Pas 3, aixi que ja no hi ha un pas separat de "camps".
        $st  = @{ Date=$null; Decisions=$null; Conclusions=@(); Fields=[ordered]@{} }
        $pre = @{ Date=$null; ConclTitles=$null }
        $step = 1
        while ($step -ge 1 -and $step -le 3) {
            switch ($step) {
                1 {
                    $r = Prompt-RoundDate -preset $pre.Date
                    if ($r.Nav -eq 'back') { return }   # enrere des del primer pas = sortir
                    $st.Date = $r.Data; $pre.Date = $r.Data; $step = 2
                }
                2 {
                    $r = Prompt-SeguimentComments -requirements $model.Requirements -dateStr $st.Date -preload $st.Decisions
                    if ($r.Nav -eq 'back') { $step = 1 }
                    else { $st.Decisions = $r.Data; $step = 3 }
                }
                3 {
                    if ($conclAll.Selectable.Count -eq 0) {
                        $st.Conclusions = @()
                        $step = 4   # surt del bucle: no hi ha res mes a fer
                    } else {
                        $r = Select-Conclusions -conclusions $conclAll.Selectable -always $conclAll.Always -fields $st.Fields -preloadTitles $pre.ConclTitles -preloadValues (Get-FieldValuesForSession $st.Fields)
                        if ($r.Nav -eq 'back') { $step = 2 }
                        else {
                            $st.Conclusions = $r.Data
                            $pre.ConclTitles = @($st.Conclusions | ForEach-Object { $_.Title })
                            $step = 4   # surt del bucle
                        }
                    }
                }
            }
        }

        $roundDt    = [datetime]::ParseExact($st.Date, 'dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
        $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)

        $outPath = Apply-SeguimentXml -xmlInfo $xmlInfo -bodyParas $bodyParas -model $model `
                       -conclusionStartIndex $conclusionStartIndex `
                       -decisions $st.Decisions -dateStr $st.Date `
                       -conclHeaderText $conclAll.HeaderText -selectedConclusions $st.Conclusions `
                       -alwaysConclusions $conclAll.Always -fields $st.Fields `
                       -srcPath $sourcePath -sourceBaseName $sourceBase -roundDate $roundDt

        [System.Windows.Forms.MessageBox]::Show(
            "Informe de seguiment generat:`n$outPath", 'Finalitzat', 'OK', 'Information') | Out-Null

        # Obrir el resultat amb l'app per defecte (Word), sense COM.
        try { Invoke-Item -LiteralPath $outPath } catch { }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        throw
    }
}
