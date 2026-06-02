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
  Add-FieldsFromConclusions, Apply-Fields, Prompt-Fields, Select-Conclusions,
  _ResolveOutputDir, _GetUniqueOutputPath.
#>

# Espai de noms WordprocessingML.
$Script:WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

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
            })
        }
        # 'other' (cos, URL, intro, capcalera, conclusions...) s'ignora al model.
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

# Nom del fitxer de sortida del seguiment.
#   - Font amb l'esquema del programa "YYYY-MM-DD_<Cat>_GIA <id>" ->
#     "<data>_<Cat>_GIA <id>_SEG.docx" (preservant Cat i GIA; sense duplicar _SEG).
#   - Font feta a ma -> "<data>_Seguiment_<nom>.docx".
# Sempre s'ha de passar el resultat per _GetUniqueOutputPath (afegeix _2, _3...).
function _SeguimentOutputName([string]$sourceBaseName, [datetime]$roundDate) {
    $day = $roundDate.ToString('yyyy-MM-dd')
    $rx  = [regex]'^\d{4}-\d{2}-\d{2}_(.+?)_GIA\s+(\d+|s_n)'
    $m   = $rx.Match([string]$sourceBaseName)
    if ($m.Success) {
        $cat = $m.Groups[1].Value
        $gia = $m.Groups[2].Value
        return ('{0}_{1}_GIA {2}_SEG.docx' -f $day, $cat, $gia)
    }
    $safe = (([string]$sourceBaseName) -replace '[\\/:*?"<>|]','_').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'informe' }
    return ('{0}_Seguiment_{1}.docx' -f $day, $safe)
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

# Carrega word/document.xml com a XmlDocument. Retorna { Path; Xml; Ns; Body }.
function _LoadDocxXml($docxPath) {
    $text = _ReadDocxPartText $docxPath 'word/document.xml'
    if ($null -eq $text) { throw "El fitxer no sembla un .docx valid (falta word/document.xml)." }
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($text)
    $ns = _NewWordNsMgr $xml
    $body = $xml.SelectSingleNode('//w:body', $ns)
    if ($null -eq $body) { throw "document.xml sense <w:body>." }
    return [pscustomobject]@{ Path=$docxPath; Xml=$xml; Ns=$ns; Body=$body }
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

# Registres { Index; Text; ListString; Bold } a partir dels paragrafs del body.
function _CollectParaRecordsXml($bodyParas, $ns) {
    $records = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $bodyParas.Count; $i++) {
        $p = $bodyParas[$i]
        [void]$records.Add([pscustomobject]@{
            Index      = ($i + 1)
            Text       = (_ParagraphTextXml $p $ns)
            ListString = ''
            Bold       = (_ParagraphBoldStateXml $p $ns)
        })
    }
    return $records.ToArray()
}

# Posa/treu negreta a TOTS els runs amb text d'un paragraf.
#   $boldOn = $true  -> afegeix <w:b/> i <w:bCs/>
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
        $b   = $rPr.SelectSingleNode('w:b', $ns)
        $bCs = $rPr.SelectSingleNode('w:bCs', $ns)
        if ($boldOn) {
            if ($null -eq $b)   { [void]$rPr.AppendChild($xml.CreateElement('w','b',$w)) }   else { $b.RemoveAttribute('val', $w) }
            if ($null -eq $bCs) { [void]$rPr.AppendChild($xml.CreateElement('w','bCs',$w)) } else { $bCs.RemoveAttribute('val', $w) }
        } else {
            if ($null -eq $b)   { $b   = $xml.CreateElement('w','b',$w);   [void]$rPr.AppendChild($b) }
            if ($null -eq $bCs) { $bCs = $xml.CreateElement('w','bCs',$w); [void]$rPr.AppendChild($bCs) }
            [void]$b.SetAttribute('val', $w, 'false')
            [void]$bCs.SetAttribute('val', $w, 'false')
        }
    }
}

# Crea un <w:p> d'anotacio clonant el pPr del requeriment (estil/sagnat) pero
# sense numeracio ni rPr de marca, amb un run amb el text.
function _MakeAnnotationParagraphXml($xmlInfo, $reqNode, $line) {
    $xml = $xmlInfo.Xml; $ns = $xmlInfo.Ns; $w = $Script:WNS
    $p = $xml.CreateElement('w','p',$w)
    $reqPPr = $reqNode.SelectSingleNode('w:pPr', $ns)
    if ($null -ne $reqPPr) {
        $pPr = $reqPPr.CloneNode($true)
        $numPr = $pPr.SelectSingleNode('w:numPr', $ns)
        if ($null -ne $numPr) { [void]$pPr.RemoveChild($numPr) }
        $pmRPr = $pPr.SelectSingleNode('w:rPr', $ns)   # format de la marca de paragraf
        if ($null -ne $pmRPr) { [void]$pPr.RemoveChild($pmRPr) }
        [void]$p.AppendChild($pPr)
    }
    $r = $xml.CreateElement('w','r',$w)
    $t = $xml.CreateElement('w','t',$w)
    $xsp=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace'); $xsp.Value='preserve'; [void]$t.Attributes.Append($xsp)
    $t.InnerText = [string]$line
    [void]$r.AppendChild($t)
    [void]$p.AppendChild($r)
    return ,$p   # ,: el node <w:p> es IEnumerable; evitem que s'enumeri
}

# Runs <w:r> a partir d'un text interpretant **negreta** i //cursiva//.
function _RichTextRunsXml($xmlInfo, $text) {
    $xml = $xmlInfo.Xml; $w = $Script:WNS
    $runs = New-Object System.Collections.ArrayList
    $make = {
        param($segment, $bold, $italic)
        if ([string]::IsNullOrEmpty($segment)) { return }
        $r = $xml.CreateElement('w','r',$w)
        if ($bold -or $italic) {
            $rPr = $xml.CreateElement('w','rPr',$w)
            if ($bold)   { [void]$rPr.AppendChild($xml.CreateElement('w','b',$w)) }
            if ($italic) { [void]$rPr.AppendChild($xml.CreateElement('w','i',$w)) }
            [void]$r.AppendChild($rPr)
        }
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
        $rPr = $xml.CreateElement('w','rPr',$w); [void]$rPr.AppendChild($xml.CreateElement('w','b',$w)); [void]$r.AppendChild($rPr)
        $t = $xml.CreateElement('w','t',$w); $xsp=$xml.CreateAttribute('xml','space','http://www.w3.org/XML/1998/namespace'); $xsp.Value='preserve'; [void]$t.Attributes.Append($xsp); $t.InnerText = [string]$text
        [void]$r.AppendChild($t); [void]$p.AppendChild($r)
    } else {
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
function Read-ConclusionsXml($path) {
    $empty = [pscustomobject]@{ HeaderText=''; Selectable=@(); Always=@() }
    if (-not (Test-Path -LiteralPath $path)) { return $empty }

    $xmlInfo   = _LoadDocxXml $path
    $ns        = $xmlInfo.Ns
    $bodyParas = @(_BodyParagraphsXml $xmlInfo)

    $headerText   = ''
    $selectable   = New-Object System.Collections.ArrayList
    $always       = New-Object System.Collections.ArrayList
    $pendingTitle = $null
    $isFirst      = $true

    foreach ($p in $bodyParas) {
        $text = (_ParagraphTextXml $p $ns).TrimEnd("`r","`n","`t"," ")
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $styleVal = ''
        $pStyle = $p.SelectSingleNode('w:pPr/w:pStyle', $ns)
        if ($null -ne $pStyle) { $styleVal = $pStyle.GetAttribute('val', $Script:WNS) }
        $isH1 = Test-StyleMatch $styleVal 1

        if ($isFirst -and -not $isH1) {
            $isFirst = $false
            $jc = $p.SelectSingleNode('w:pPr/w:jc', $ns)
            $centered = ($null -ne $jc -and $jc.GetAttribute('val', $Script:WNS) -eq 'center')
            if ($centered) { $headerText = $text; continue }
        }
        $isFirst = $false

        if ($isH1) { $pendingTitle = $text; continue }

        if ($text.StartsWith('::SEMPRE::')) {
            [void]$always.Add($text.Substring('::SEMPRE::'.Length).Trim())
            $pendingTitle = $null
            continue
        }
        if ($null -ne $pendingTitle) {
            [void]$selectable.Add([pscustomobject]@{ Title=$pendingTitle; Body=$text })
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

    # (1) Esborrar el bloc de conclusions (nodes <w:p> des del cut fins al final).
    # El <w:sectPr> (fill del body, no es <w:p>) es preserva automaticament.
    if ($conclusionStartIndex -ge 1) {
        for ($i = $bodyParas.Count - 1; $i -ge ($conclusionStartIndex - 1); $i--) {
            $node = $bodyParas[$i]
            if ($null -ne $node.ParentNode) { [void]$node.ParentNode.RemoveChild($node) }
        }
    }

    # (2) Per cada requeriment: inserir la nova anotacio (si hi ha comentari) i
    # recalcular la negreta de tota la columna (requeriment + anotacions).
    for ($k = 0; $k -lt $model.Requirements.Count; $k++) {
        $req = $model.Requirements[$k]
        $dec = $decisions[$k]
        $reqNode = $bodyParas[$req.ParaIndex - 1]

        $annotNodes = New-Object System.Collections.ArrayList
        foreach ($a in $req.Annotations) { [void]$annotNodes.Add($bodyParas[$a.ParaIndex - 1]) }

        $comment = [string]$dec.NewComment
        if (-not [string]::IsNullOrWhiteSpace($comment)) {
            $line = _FormatAnnotationLine $dateStr $comment
            $newP = _MakeAnnotationParagraphXml $xmlInfo $reqNode $line
            $anchor = if ($annotNodes.Count -gt 0) { $annotNodes[$annotNodes.Count - 1] } else { $reqNode }
            [void]$xmlInfo.Body.InsertAfter($newP, $anchor)
            [void]$annotNodes.Add($newP)
        }

        $boldOn = (_ShouldBeBold $dec.Resolved)
        _SetParagraphBoldXml $xmlInfo $reqNode $boldOn
        foreach ($n in $annotNodes) { _SetParagraphBoldXml $xmlInfo $n $boldOn }
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

    $outName   = _SeguimentOutputName $sourceBaseName $roundDate
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
    $form.Size = New-Object System.Drawing.Size(440, 235)
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

    $form.AcceptButton = $btnNou
    $res = $form.ShowDialog()
    if ($res -eq 'Yes') { return 'nou' }
    if ($res -eq 'No')  { return 'seguiment' }
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
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Tria el PRIMER paragraf del bloc de conclusions a esborrar (s'esborrara fins al final):"
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(680, 20)
    $form.Controls.Add($lbl)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(15, 35)
    $list.Size = New-Object System.Drawing.Size(680, 360)
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
    $form.Size = New-Object System.Drawing.Size(820, 620)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = ("Anotacio del {0}. Escriu el comentari de cada requeriment i marca 'Resolt' si ha quedat resolt." -f $dateStr)
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(770, 20)
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 35)
    $panel.Size = New-Object System.Drawing.Size(775, 490)
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panel)

    $rows = @()
    $y = 8
    for ($i = 0; $i -lt $requirements.Count; $i++) {
        $req = $requirements[$i]

        $reqLbl = New-Object System.Windows.Forms.Label
        $reqLbl.Text = [string]$req.Text
        $reqLbl.Location = New-Object System.Drawing.Point(8, $y)
        $reqLbl.Size = New-Object System.Drawing.Size(740, 38)
        $reqLbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $panel.Controls.Add($reqLbl)
        $y += 40

        if ($req.Annotations.Count -gt 0) {
            $hist = ($req.Annotations | ForEach-Object { ('{0}: {1}' -f $_.Date, $_.Text) }) -join "`r`n"
            $histLbl = New-Object System.Windows.Forms.Label
            $histLbl.Text = $hist
            $histLbl.Location = New-Object System.Drawing.Point(20, $y)
            $histLbl.Size = New-Object System.Drawing.Size(728, (16 * $req.Annotations.Count + 2))
            $histLbl.ForeColor = [System.Drawing.Color]::DimGray
            $panel.Controls.Add($histLbl)
            $y += (16 * $req.Annotations.Count + 6)
        }

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.ScrollBars = 'Vertical'
        $tb.Location = New-Object System.Drawing.Point(20, $y)
        $tb.Size = New-Object System.Drawing.Size(728, 44)
        $panel.Controls.Add($tb)
        $y += 50

        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = 'Resolt'
        $cb.Location = New-Object System.Drawing.Point(20, $y)
        $cb.Size = New-Object System.Drawing.Size(120, 22)
        $cb.Checked = [bool]$req.WasResolved
        $panel.Controls.Add($cb)
        $y += 30

        # Separador visual
        $sep = New-Object System.Windows.Forms.Label
        $sep.BorderStyle = 'Fixed3D'
        $sep.Location = New-Object System.Drawing.Point(8, $y)
        $sep.Size = New-Object System.Drawing.Size(740, 2)
        $panel.Controls.Add($sep)
        $y += 12

        # Precarrega (tornar enrere)
        if ($null -ne $preload -and $i -lt $preload.Count) {
            $tb.Text    = [string]$preload[$i].NewComment
            $cb.Checked = [bool]$preload[$i].Resolved
        }

        $rows += [pscustomobject]@{ Comment=$tb; Resolved=$cb }
    }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 535)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(695, 535)
    $ok.Size = New-Object System.Drawing.Size(95, 30)
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
        $records   = @(_CollectParaRecordsXml $bodyParas $xmlInfo.Ns)
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

        $conclAll = Read-ConclusionsXml $ConclusionsPath

        # Maquina de passos: 1=data, 2=comentaris, 3=conclusions, 4=camps. 5=fi.
        $st  = @{ Date=$null; Decisions=$null; Conclusions=@(); Fields=$null }
        $pre = @{ Date=$null; ConclTitles=$null }
        $step = 1
        while ($step -ge 1 -and $step -le 4) {
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
                        $step = 4
                    } else {
                        $r = Select-Conclusions -conclusions $conclAll.Selectable -preloadTitles $pre.ConclTitles
                        if ($r.Nav -eq 'back') { $step = 2 }
                        else {
                            $st.Conclusions = $r.Data
                            $pre.ConclTitles = @($st.Conclusions | ForEach-Object { $_.Title })
                            $step = 4
                        }
                    }
                }
                4 {
                    $fields = [ordered]@{}
                    Add-FieldsFromConclusions $fields $st.Conclusions $conclAll.Always
                    if ($fields.Count -eq 0) {
                        $st.Fields = $fields; $step = 5
                    } else {
                        $r = Prompt-Fields -fields $fields
                        if ($r.Nav -eq 'back') { $step = 3 }
                        else { $st.Fields = $r.Data; $step = 5 }
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
