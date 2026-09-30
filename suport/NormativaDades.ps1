#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Normativa": baixa a local\normativa el text de totes les normes que fa
  servir el programa (les que cita REQ1 i les dels marcadors de l'usuari).

.DESCRIPTION
  EL CATALEG es suport\normativa.json (al repositori: no porta cap dada
  personal). Cada norma hi te l'ambit i el tema (la classificacio dels marcadors
  de Chrome de l'usuari: "Vector ambiental" > "Residus"...), el tipus, el
  numero, l'any, un titol curt i l'URL d'on es baixa.

  TOT EN UNA SOLA CARPETA, classificat pel NOM del fitxer (peticio de
  l'usuari, setembre 2026):

      Vector ambiental_Residus_2016_Decret 197-2016 Comunicació prèvia....pdf

  TEXT CONSOLIDAT (la versio vigent, amb les modificacions):
    - BOE: la pagina de la norma dona l'identificador (BOE-A-...) i la data de
      l'ultima actualitzacio; el PDF consolidat es a /buscar/pdf/.../
      <id>-consolidado.pdf. Si la norma no te text consolidat (algunes ITC),
      l'original que enllaca la mateixa pagina.
    - Portal Juridic de Catalunya, CIDO i la resta de pagines web: no donen
      cap PDF del consolidat, o sigui que la pagina s'IMPRIMEIX A PDF amb l'Edge
      sense finestra (--headless --print-to-pdf), que ja hi es a tots els Windows.
    - Un URL que ja es un PDF (EUR-Lex, CTE, notes aclaridores...): tal qual.
    - Sense URL: l'index diu amb quin nom s'ha de desar a ma.

  ACTUALITZAR SOLA: una norma del BOE es torna a baixar quan la data de la seva
  ultima actualitzacio canvia; la resta, quan fa mes de $Script:NormativaDiesRefresc
  dies. La versio anterior NO es perd: va a la subcarpeta 'anteriors' amb la data
  fins a la qual va ser la bona.

  L'INDEX es un Excel ('0 Index normativa.xlsx') a la mateixa carpeta: cada norma
  amb el seu fitxer (enllac), l'enllac web, la versio, quan es va baixar, si ha
  anat be i a quins punts de REQ1 surt. Es fa sense Excel (OpenXML a ma).

  PARTIT EN DOS: aquest fitxer (el cataleg, els noms, l'index i la cerca
  d'una norma en un text; pur, sense interficie) i Normativa.ps1 (les baixades
  i la finestra). La fitxa d'ajuda (UiComuns.ps1) nomes necessita el primer, i
  si fos un sol fitxer UiComuns i Normativa dependrien l'un de l'altre.
#>

$Script:NormativaCatalegPath = Join-Path $PSScriptRoot 'normativa.json'
$Script:NormativaIndexNom    = '0 Index normativa.xlsx'
$Script:NormativaEstatNom    = '_estat.json'
$Script:NormativaDiesRefresc = 180
$Script:NormativaTitolMax    = 70
# El cataleg ja llegit, per a la fitxa d'ajuda (Get-NormativaPdfDeText).
$Script:NormativaCache       = $null
# Un navegador de debo: el BOE i alguns servidors de la Generalitat tornen un
# error o una pagina buida a un client que no s'identifica.
$Script:NormativaUA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36 Edg/126.0'

# ----------------------------------------------------------------------------
# EL CATALEG
# ----------------------------------------------------------------------------
function Get-NormativaCataleg([string]$path = $Script:NormativaCatalegPath) {
    $o = Read-JsonFile $path
    if ($null -eq $o) { return @() }
    return @($o.Normes)
}

function Get-NormativaDir {
    if (-not $RepoRoot) { return '' }
    return [string](Get-LocalSubdir $RepoRoot 'Normativa')
}

# ----------------------------------------------------------------------------
# COM ES RECONEIX UNA NORMA DINS D'UN TEXT. Funcions PURES.
# ----------------------------------------------------------------------------
# Els textos del cataleg la citen de moltes maneres ("Real Decreto 842/2002",
# "RD 842/2002", "Reial Decret 842/2002"; "Ley"/"Llei"...). Tot es porta a una
# forma comuna abans de comparar: minuscules, sense accents i els tipus abreujats.
function _NormativaNormText([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $t = $s.Replace([char]0x2019, "'").Replace([char]0x00B7, '.')
    $t = $t.Normalize([System.Text.NormalizationForm]::FormD)
    $t = ($t -replace '\p{Mn}', '').ToLowerInvariant()
    $t = $t -replace '\s+', ' '
    $subs = @(
        @('\breal decreto legislativo\b', 'rdleg'), @('\breial decret legislatiu\b', 'rdleg'), @('\brdlg\b', 'rdleg'),
        @('\breal decreto[- ]ley\b', 'rdl'), @('\breial decret[- ]llei\b', 'rdl'),
        @('\breal decreto\b', 'rd'), @('\breial decret\b', 'rd'),
        @('\bdecreto legislativo\b', 'dleg'), @('\bdecret legislatiu\b', 'dleg'),
        @('\bdecret[- ]llei\b', 'dl'), @('\bdecreto[- ]ley\b', 'dl'),
        @('\bdecreto\b', 'decret'),
        @('\bley organica\b', 'lo'), @('\bllei organica\b', 'lo'),
        @('\bley\b', 'llei'), @('\borden\b', 'ordre'),
        @('\breglamento\b', 'reglament'),
        @('\breglament \((ce|ue)\)( n[o\.\u00BA\u00B0]*)?', 'reglament')
    )
    foreach ($p in $subs) { $t = $t -replace $p[0], $p[1] }
    return ($t -replace '\s+', ' ').Trim()
}

# El prefix normalitzat de cada tipus del cataleg ('' = no es busca pel numero).
function _NormativaPrefix([string]$tipus) {
    switch -Regex ($tipus) {
        '^(Llei|Ley)$'            { return 'llei' }
        '^Decret$'                { return 'decret' }
        '^Decret legislatiu$'     { return 'dleg' }
        '^Decret llei$'           { return 'dl' }
        '^RD$'                    { return 'rd' }
        '^RDL$'                   { return 'rdl' }
        '^RDLeg$'                 { return 'rdleg' }
        '^(Ordre|Orden)$'         { return 'ordre' }
        '^Reglament \((CE|UE)\)$' { return 'reglament' }
        '^Directiva$'             { return 'directiva' }
        '^LO$'                    { return 'lo' }
    }
    return ''
}

# Les claus (ja normalitzades) amb que una norma es reconeix en un text: la del
# tipus i el numero ("rd 842/2002") mes les 'Claus' que porti el cataleg (una
# ordenanca no te numero).
function _NormativaClaus($e) {
    $out = New-Object System.Collections.ArrayList
    $p = _NormativaPrefix ([string]$e.Tipus)
    if ($p) { [void]$out.Add((_NormativaNormText ($p + ' ' + [string]$e.Num))) }
    foreach ($c in @($e.Claus)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$c)) { [void]$out.Add((_NormativaNormText ([string]$c))) }
    }
    return [string[]]@($out)
}

# La norma que un text cita PRIMER (la fitxa d'ajuda posa la principal davant:
# "Llei 16/2002 ... Desplegament: Decret 176/2009"). $null si no en cita cap.
# El numero ha d'anar sencer: "rd 9/2005" no es troba dins "rd 19/2005" ni
# "decret 30/2015" dins "decret 130/2015".
function _NormativaBuscaEnText($normes, [string]$text) {
    $t = _NormativaNormText $text
    if (-not $t) { return $null }
    $millor = $null; $pos = [int]::MaxValue
    foreach ($e in @($normes)) {
        foreach ($k in @(_NormativaClaus $e)) {
            $m = [regex]::Match($t, '(?<![\w/])' + [regex]::Escape($k) + '(?![\d/])')
            if ($m.Success -and $m.Index -lt $pos) { $pos = $m.Index; $millor = $e }
        }
    }
    return $millor
}

# ----------------------------------------------------------------------------
# EL NOM DEL FITXER. Funcio PURA.
# ----------------------------------------------------------------------------
# Ambit_Tema_Any_Tipus Num Titol.pdf (el tema i l'any, si en te). La barra del
# numero no pot anar en un nom de fitxer: "Decret 197-2016". El titol es talla
# perque la ruta sencera no passi dels 260 caracters del Windows.
function _NormativaNetejaNom([string]$s) {
    $t = ([string]$s).Replace([char]0x2019, "'")
    $t = $t -replace '[\\/:*?"<>|]', '-'
    return ($t -replace '\s+', ' ').Trim().TrimEnd('.')
}

function _NormativaNomFitxer($e) {
    $titol = [string]$e.Titol
    if ($titol.Length -gt $Script:NormativaTitolMax) { $titol = $titol.Substring(0, $Script:NormativaTitolMax).TrimEnd() }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add([string]$e.Ambit)
    if (-not [string]::IsNullOrWhiteSpace([string]$e.Tema)) { [void]$parts.Add([string]$e.Tema) }
    if (-not [string]::IsNullOrWhiteSpace([string]$e.Any)) { [void]$parts.Add([string]$e.Any) }
    [void]$parts.Add((@([string]$e.Tipus, [string]$e.Num, $titol) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ')
    return ((_NormativaNetejaNom ($parts -join '_')) + '.pdf')
}

# ----------------------------------------------------------------------------
# D'ON I COM ES BAIXA. Funcions PURES.
# ----------------------------------------------------------------------------
#   boe    -> pagina del BOE: l'identificador i la versio, i el PDF consolidat
#   pdf    -> l'URL ja es el PDF
#   web    -> qualsevol altra pagina: s'imprimeix a PDF amb l'Edge
#   manual -> sense URL
function _NormativaFont([string]$url) {
    if ([string]::IsNullOrWhiteSpace($url)) { return 'manual' }
    $u = $url.Trim()
    if ($u -match '^https?://(www\.)?boe\.es/(buscar/(act|doc)\.php\?id=BOE-A-|eli/es/)') { return 'boe' }
    if ($u -match '(?i)\.pdf($|[?#/])' -or $u -match '(?i)/TXT/PDF/') { return 'pdf' }
    return 'web'
}

# El que interessa de la pagina d'una norma del BOE. PURA.
#   Id       BOE-A-aaaa-nnn
#   Versio   la data de l'ultima actualitzacio del text consolidat ('' si no
#            n'hi ha: la norma no te text consolidat)
#   Original l'URL del PDF de la publicacio original ('' si no hi es)
function _NormativaBoeInfo([string]$html) {
    $id = ''
    foreach ($re in @('(BOE-A-\d{4}-\d+)-consolidado\.pdf', '[?&]id=(BOE-A-\d{4}-\d+)', '(BOE-A-\d{4}-\d+)')) {
        $m = [regex]::Match([string]$html, $re)
        if ($m.Success) { $id = $m.Groups[1].Value; break }
    }
    $versio = ''
    $m = [regex]::Match([string]$html, '(?is)ltima actualizaci.{0,80}?(\d{1,2}/\d{1,2}/\d{4})')
    if ($m.Success) { $versio = $m.Groups[1].Value }
    $orig = ''
    $m = [regex]::Match([string]$html, '(?i)href="([^"]*/boe/dias/[^"]+\.pdf)"')
    if ($m.Success) {
        $orig = $m.Groups[1].Value
        if ($orig.StartsWith('/')) { $orig = 'https://www.boe.es' + $orig }
    }
    return @{ Id = $id; Versio = $versio; Original = $orig }
}

# ELS ENLLACOS A UN PDF D'UNA PAGINA (el boto "PDF" del Portal Juridic, el de
# la BOPB...), del mes probable al menys. PURA.
#   3  el text de l'enllac es "PDF" (el boto de "Descarrega  PDF  RDF  XML")
#   2  l'adreca acaba en .pdf o porta format=pdf
#   1  hi surt "pdf" en algun atribut
# Fora: RDF/TTL/XML (les altres descarregues del mateix bloc) i el RESUM fet
# amb IA del Portal Juridic ("Descarrega (CA)"), que no es la norma.
function _NormativaPdfsDeHtml([string]$html, [string]$base) {
    if ([string]::IsNullOrWhiteSpace($html)) { return [string[]]@() }
    $cands = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($html, '(?is)<a\b([^>]*)>(.*?)</a>')) {
        $attrs = $m.Groups[1].Value
        $mh = [regex]::Match($attrs, '(?i)\bhref\s*=\s*["'']([^"'']+)["'']')
        if (-not $mh.Success) { continue }
        $href = [System.Net.WebUtility]::HtmlDecode($mh.Groups[1].Value).Trim()
        if ($href -match '^(?i)(javascript:|#|mailto:)') { continue }
        $text = [System.Net.WebUtility]::HtmlDecode(([regex]::Replace($m.Groups[2].Value, '<[^>]+>', ' '))) -replace '\s+', ' '
        $text = $text.Trim()
        $tot = ($attrs + ' ' + $text + ' ' + $href)
        if ($href -match '(?i)\.(rdf|ttl|xml|docx?|xlsx?)([?#]|$)' -or $href -match '(?i)[?&](format|output|tipus)=(rdf|ttl|xml)') { continue }
        if ($tot -match '(?i)resum|resumen|\(ca\)|\(es\)') { continue }
        $punts = 0
        if ($text -match '^(?i)(descarrega\s+)?pdf$' -or $attrs -match '(?i)(title|aria-label)\s*=\s*["''](descarrega\s+)?pdf["'']') { $punts = 3 }
        elseif ($href -match '(?i)\.pdf([?#/]|$)' -or $href -match '(?i)[?&](format|output|tipus|type)=pdf') { $punts = 2 }
        elseif ($tot -match '(?i)pdf') { $punts = 1 }
        if ($punts -eq 0) { continue }
        $abs = $href
        try { $abs = (New-Object System.Uri((New-Object System.Uri($base)), $href)).AbsoluteUri } catch { }
        [void]$cands.Add([pscustomobject]@{ Url = $abs; Punts = $punts; Ordre = $cands.Count })
    }
    $vist = @{}
    $out = New-Object System.Collections.ArrayList
    foreach ($c in @($cands | Sort-Object -Property @{ Expression = 'Punts'; Descending = $true }, @{ Expression = 'Ordre'; Descending = $false })) {
        if ($vist.ContainsKey($c.Url)) { continue }
        $vist[$c.Url] = $true
        [void]$out.Add([string]$c.Url)
    }
    return [string[]]@($out)
}

function _NormativaBoePdfConsolidat([string]$id) {
    $m = [regex]::Match([string]$id, '^BOE-A-(\d{4})-\d+$')
    if (-not $m.Success) { return '' }
    return ('https://www.boe.es/buscar/pdf/' + $m.Groups[1].Value + '/' + $id + '-consolidado.pdf')
}

function _NormativaEsPdf([byte[]]$b) {
    return ($null -ne $b -and $b.Length -ge 5 -and $b[0] -eq 0x25 -and $b[1] -eq 0x50 -and $b[2] -eq 0x44 -and $b[3] -eq 0x46)
}

# CAL BAIXAR-LA? PURA. $est = el que es va apuntar l'ultima vegada (o $null),
# $existeix = el fitxer hi es, $versio = la versio d'ara ('' si no se sap).
# Torna '' (no cal) o el motiu: 'nova', 'versio', 'antiga', 'forcat'.
function _NormativaCalBaixar($est, [bool]$existeix, [string]$versio, [datetime]$ara, [bool]$forca) {
    if (-not $existeix) { return 'nova' }
    if ($forca) { return 'forcat' }
    $vella = if ($null -ne $est) { [string]$est.Versio } else { '' }
    if ($versio) { if ($versio -ne $vella) { return 'versio' } else { return '' } }
    $d = [datetime]::MinValue
    $baixat = if ($null -ne $est) { [string]$est.Baixat } else { '' }
    if (-not [datetime]::TryParse($baixat, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$d)) { return 'antiga' }
    if (($ara - $d).TotalDays -gt $Script:NormativaDiesRefresc) { return 'antiga' }
    return ''
}

# On va la versio que es substitueix: anteriors\<nom> (fins dd-MM-aaaa).pdf. PURA.
function _NormativaNomAnterior([string]$nom, [datetime]$fins) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($nom)
    return ($base + ' (fins ' + $fins.ToString('dd-MM-yyyy') + ').pdf')
}

# ELS PUNTS DE REQ1 QUE CITEN CADA NORMA: Id -> llista de titols. PURA (rep el
# cataleg ja parsejat, el que torna Get-ParsedCataleg).
function _NormativaPuntsReq1($normes, $req1) {
    $out = @{}
    if ($null -eq $req1) { return $out }
    $visita = New-Object System.Collections.ArrayList
    foreach ($s in @($req1.Sections)) {
        foreach ($el in @($s.Items)) {
            [void]$visita.Add($el)
            foreach ($c in @($el.Children)) { [void]$visita.Add($c) }
        }
    }
    foreach ($el in $visita) {
        if ($null -eq $el) { continue }
        $textos = New-Object System.Collections.ArrayList
        if ($null -ne $el.Ajuda) { [void]$textos.Add([string]$el.Ajuda.Norma) }
        foreach ($l in @($el.BodyLines)) { [void]$textos.Add([string]$l) }
        $t = _NormativaNormText ($textos -join ' ')
        if (-not $t) { continue }
        foreach ($e in @($normes)) {
            $hi = $false
            foreach ($k in @(_NormativaClaus $e)) {
                if ([regex]::IsMatch($t, '(?<![\w/])' + [regex]::Escape($k) + '(?![\d/])')) { $hi = $true; break }
            }
            if (-not $hi) { continue }
            $id = [string]$e.Id
            if (-not $out.ContainsKey($id)) { $out[$id] = New-Object System.Collections.ArrayList }
            $nomPunt = [string]$el.Short
            if ($nomPunt -and -not $out[$id].Contains($nomPunt)) { [void]$out[$id].Add($nomPunt) }
        }
    }
    return $out
}

# ----------------------------------------------------------------------------
# L'INDEX EN EXCEL, sense Excel. Funcio PURA: torna els bytes del .xlsx.
# ----------------------------------------------------------------------------
# $files: llista de files; cada cel·la es un text o @{ Text; Link } (surt com a
# formula HYPERLINK, que l'Excel calcula en obrir-lo: un enllac relatiu obre el
# fitxer de la mateixa carpeta que l'index).
function _NormativaXlsxCol([int]$i) {
    $s = ''
    $n = $i + 1
    while ($n -gt 0) { $r = ($n - 1) % 26; $s = [string][char](65 + $r) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
    return $s
}

function _NormativaXmlEsc([string]$s) {
    $t = [string]$s
    $t = [regex]::Replace($t, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function _NormativaXlsxBytes([string[]]$capcalera, $files, [int[]]$amples) {
    Add-Type -AssemblyName System.IO.Compression
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $capcalera.Count; $i++) {
        $w = if ($null -ne $amples -and $i -lt $amples.Count) { $amples[$i] } else { 20 }
        [void]$sb.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $w))
    }
    [void]$sb.Append('</cols><sheetData>')
    $totes = New-Object System.Collections.ArrayList
    [void]$totes.Add(@($capcalera))
    foreach ($f in @($files)) { [void]$totes.Add(@($f)) }
    for ($r = 0; $r -lt $totes.Count; $r++) {
        [void]$sb.Append(('<row r="{0}">' -f ($r + 1)))
        $cel = @($totes[$r])
        for ($c = 0; $c -lt $cel.Count; $c++) {
            $ref = (_NormativaXlsxCol $c) + ($r + 1)
            $v = $cel[$c]
            $estil = if ($r -eq 0) { ' s="1"' } else { '' }
            if ($v -is [System.Collections.IDictionary] -and -not [string]::IsNullOrWhiteSpace([string]$v.Link)) {
                $f = 'HYPERLINK("' + ([string]$v.Link).Replace('"', '""') + '","' + ([string]$v.Text).Replace('"', '""') + '")'
                [void]$sb.Append(('<c r="{0}" t="str" s="2"><f>{1}</f><v>{2}</v></c>' -f $ref, (_NormativaXmlEsc $f), (_NormativaXmlEsc ([string]$v.Text))))
            } else {
                $txt = if ($v -is [System.Collections.IDictionary]) { [string]$v.Text } else { [string]$v }
                [void]$sb.Append(('<c r="{0}" t="inlineStr"{1}><is><t xml:space="preserve">{2}</t></is></c>' -f $ref, $estil, (_NormativaXmlEsc $txt)))
            }
        }
        [void]$sb.Append('</row>')
    }
    $ultima = (_NormativaXlsxCol ($capcalera.Count - 1)) + $totes.Count
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append(('<autoFilter ref="A1:{0}"/>' -f $ultima))
    [void]$sb.Append('</worksheet>')

    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Normativa" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Normativa!$A$1:$' + (_NormativaXlsxCol ($capcalera.Count - 1)) + '$' + $totes.Count + '</definedName></definedNames><calcPr fullCalcOnLoad="1"/></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/styles.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf fontId="1" applyFont="1"/><xf fontId="2" applyFont="1"/></cellXfs></styleSheet>'
        'xl/worksheets/sheet1.xml' = $sb.ToString()
    }
    $ms = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        foreach ($k in $parts.Keys) {
            $en = $zip.CreateEntry($k)
            $st = $en.Open()
            try { $b = $utf8.GetBytes([string]$parts[$k]); $st.Write($b, 0, $b.Length) } finally { $st.Dispose() }
        }
    } finally { $zip.Dispose() }
    return ,$ms.ToArray()
}

# Les files de l'index. PURA: $estat = Id -> el que se'n va apuntar; $punts =
# Id -> titols de REQ1; $existeix = scriptblock (nom del fitxer) -> bool.
function _NormativaFilesIndex($normes, $estat, $punts, $existeix) {
    $out = New-Object System.Collections.ArrayList
    foreach ($e in @($normes)) {
        $id = [string]$e.Id
        $nom = _NormativaNomFitxer $e
        $est = if ($null -ne $estat -and $estat.ContainsKey($id)) { $estat[$id] } else { $null }
        $hi = [bool](& $existeix $nom)
        $resultat = if ($hi -and $null -ne $est -and [string]$est.Via -eq 'pàgina impresa') { 'Baixada (pàgina desada com a PDF: no s''ha trobat el PDF de la norma)' }
                    elseif ($hi) { 'Baixada' }
                    elseif ((_NormativaFont ([string]$e.Url)) -eq 'manual') { "Sense enllaç: desa-la a mà amb aquest nom" }
                    elseif ($null -ne $est -and $est.Error) { 'Error: ' + [string]$est.Error }
                    else { 'Pendent' }
        $fitxer = if ($hi) { @{ Text = $nom; Link = $nom } } else { $nom }
        $web = if ([string]$e.Url) { @{ Text = 'Obrir'; Link = [string]$e.Url } } else { '' }
        $pp = if ($null -ne $punts -and $punts.ContainsKey($id)) { (@($punts[$id]) -join '; ') } else { '' }
        [void]$out.Add(@(
            [string]$e.Ambit, [string]$e.Tema, ([string]$e.Tipus + ' ' + [string]$e.Num), [string]$e.Titol,
            $(if ($e.Guia) { $(if ($e.Derogada) { 'Guia (antiga)' } else { 'Guia' }) } elseif ($e.Derogada) { 'Derogada' } else { 'Vigent' }),
            $fitxer, $web,
            $(if ($null -ne $est) { [string]$est.Versio } else { '' }),
            $(if ($null -ne $est -and $est.Baixat) { try { ([datetime]::Parse([string]$est.Baixat, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)).ToString('dd/MM/yyyy') } catch { '' } } else { '' }),
            $resultat, $pp))
    }
    # Amb la coma: una sola norma tornaria la fila desfeta en cel·les.
    return ,$out.ToArray()
}

$Script:NormativaCapcaleraIndex = @('Àmbit', 'Tema', 'Norma', 'Títol', 'Estat', 'Fitxer', 'Web', 'Versió (BOE)', 'Baixada el', 'Resultat', 'Punts de REQ1')
$Script:NormativaAmplesIndex    = @(18, 18, 24, 50, 10, 60, 8, 13, 12, 40, 60)

# ----------------------------------------------------------------------------
# L'ESTAT i els fitxers de la carpeta
# ----------------------------------------------------------------------------
function _NormativaLlegeixEstat([string]$dir) {
    $h = @{}
    $o = Read-JsonFile (Join-Path $dir $Script:NormativaEstatNom)
    if ($null -eq $o) { return $h }
    foreach ($p in @($o.PSObject.Properties)) {
        $v = $p.Value
        $h[[string]$p.Name] = @{ Versio = [string]$v.Versio; Baixat = [string]$v.Baixat; Error = [string]$v.Error; Mida = [long]$v.Mida; Via = [string]$v.Via }
    }
    return $h
}

function _NormativaDesaEstat([string]$dir, $estat) {
    $o = [ordered]@{}
    foreach ($k in @($estat.Keys | Sort-Object)) { $o[$k] = [pscustomobject]$estat[$k] }
    Write-JsonFile (Join-Path $dir $Script:NormativaEstatNom) ([pscustomobject]$o) 5
}

# El PDF desat d'una norma que cita un text ('' si no n'hi ha). Per a la fitxa
# d'ajuda (Show-Ajuda): "Obre el PDF desat".
function Get-NormativaPdfDeText([string]$text) {
    try {
        $dir = Get-NormativaDir
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return '' }
        if ($null -eq $Script:NormativaCache) { $Script:NormativaCache = @(Get-NormativaCataleg) }
        $e = _NormativaBuscaEnText $Script:NormativaCache $text
        if ($null -eq $e) { return '' }
        $p = Join-Path $dir (_NormativaNomFitxer $e)
        if (Test-Path -LiteralPath $p) { return [string]$p }
    } catch { }
    return ''
}

# Escriu l'index. Torna '' o el missatge d'error (l'Excel obert el bloqueja).
function _NormativaEscriuIndex([string]$dir, $normes, $estat) {
    try {
        $punts = @{}
        try { $punts = _NormativaPuntsReq1 $normes (Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')) } catch { }
        $existeix = { param($nom) Test-Path -LiteralPath (Join-Path $dir $nom) }.GetNewClosure()
        $files = _NormativaFilesIndex $normes $estat $punts $existeix
        $bytes = _NormativaXlsxBytes $Script:NormativaCapcaleraIndex $files $Script:NormativaAmplesIndex
        [System.IO.File]::WriteAllBytes((Join-Path $dir $Script:NormativaIndexNom), $bytes)
        return ''
    } catch {
        return $_.Exception.Message
    }
}

