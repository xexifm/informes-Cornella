#requires -Version 5.1
<#
.SYNOPSIS
  Llegir i editar un .docx SENSE Word: ZIP + WordprocessingML.

.DESCRIPTION
  Aixo vivia dins de Seguiment.ps1, que es el fitxer d'UNA eina. Pero
  Informes.ps1 ja en feia servir tres funcions (_LoadDocxXml, _BodyParagraphsXml
  i _ParagraphTextXml) per llegir els informes sense obrir el Word: una eina
  depenia d'una altra eina per a una cosa que no es de cap de les dues.

  Aqui hi ha NOMES el que sap d'OOXML i no sap res del seguiment. El que porta
  "Seguiment" a dins -encara que toqui XML- s'hi ha quedat: les anotacions
  datades, els blocs, i les dues que escriuen amb la LLETRA del seguiment
  (_ApplyBodyFontXml i _MakeBodyRunXml criden _SeguimentFontName). Tampoc no ha
  vingut _CollectParaRecordsXml, que decideix IsBulletChild, vocabulari del
  model de seguiment i de ningu mes.

  Es testejable a Linux: no hi ha Word pel mig, nomes System.IO.Compression i
  System.Xml.
#>

# Espai de noms WordprocessingML.
$Script:WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

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
