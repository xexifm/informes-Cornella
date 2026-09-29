#requires -Version 5.1
<#
.SYNOPSIS
  El FORMAT del correu de requeriments: el mateix que l'informe de REQ1.

.DESCRIPTION
  El correu surt de dos llocs -l'eina "Enviar correu" del PC, que llegeix el
  .docx ja generat, i l'app del mobil, que el munta de la seleccio- i abans
  cadascun el pintava a la seva manera: al PC totes les linies de cos anaven
  sagnades, al mobil els enllacos anaven sagnats i no hi havia cap linia en
  blanc entre punts, i cap dels dos s'assemblava a l'informe (seccions en
  negreta, subseccions subratllades...). Peticio de l'usuari (setembre 2026):
  el format ha de ser el de REQ1 i el del seguiment, i els dos correus han de
  coincidir.

  Com es garanteix:
    - Les MIDES (sagnies, espais, lletra) surten de $ReportFormatConfig
      (Format.ps1), que es d'on les treu el Word. _CorreuFormat les tradueix a
      px/pt i ExportaDades les publica a docs\dades\correu-format.json, que es
      el que llegeix el mobil (docs\correu.js). Hi ha guard que el fitxer
      publicat es igual al que surt d'aqui.
    - El MOBIL munta els mateixos blocs que el motor (Build-CatalegBlocs) i els
      pinta amb el mateix contracte de paragraf que _CorreuBlocsAHtml. Hi ha
      una prova que passa la mateixa seleccio de REQ1 pels dos (PowerShell i
      Node) i exigeix el MATEIX HTML.
    - El PC llegeix el .docx SENSE Word (XML) i en respecta la sagnia, els
      espais i la negreta de cada paragraf: un informe de seguiment porta les
      anotacions datades i la negreta dels punts pendents, que no son a cap
      cataleg.

  La capcalera del correu son les linies amb etiqueta de '0 CAPCALERA'
  (ID GIA, Exp. Num, Adreca, Activitat, Titular, Objecte): les mateixes que
  l'informe i en el mateix ordre (peticio de l'usuari: "copiem 0 CAPCALERA
  aixi ho tenim estandarditzat").

  NOMES DEFINEIX FUNCIONS i no toca el Word: tot es prova a Linux.
#>

# L'etiqueta de la capcalera ocupa 1560 twips (2,75 cm): es el w:hanging de les
# linies de '0 CAPCALERA.docx'. Hi ha guard que ho compara amb la plantilla.
$Script:CorreuCapcaleraTwips = 1560
$Script:CorreuLletra = 'Calibri, Arial, sans-serif'

# --- Utils de text -> HTML ----------------------------------------------------
function _EscHtml($s) {
    return ([string]$s).Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}
# Escapa, **negreta** -> <b>, //cursiva// -> <i>, i enllaça els URLs.
#
# ELS URLs S'APARTEN ABANS DE MIRAR LA CURSIVA, i no es un detall d'estil: la
# cursiva es "//...//" i un "https://" en porta un de "//" a dins. Amb DUES
# adreces a la mateixa linia, l'expressio de la cursiva es menjava tot el tros
# d'una a l'altra:
#
#   Mira https://a.cat i tambe https://b.cat
#   -> Mira https:<i>a.cat i tambe https:</i>b.cat        (les dues destrossades)
#
# No era hipotetic: aquesta funcio ja la fan servir els recordatoris, i el text
# el pot editar l'usuari. Ara cada URL es substitueix per una marca amb caracters
# de control -que no poden sortir en un text escrit a ma i que cap de les dues
# expressions toca- i es torna a posar, ja com a enllac, al final.
function _TextToHtml($s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $h = _EscHtml $s

    $urls = New-Object System.Collections.ArrayList
    $h = [regex]::Replace($h, '(https?://[^\s<]+)', {
        param($m)
        $i = $urls.Add($m.Groups[1].Value)
        return ([char]1 + [string]$i + [char]1)
    })

    $h = [regex]::Replace($h, '\*\*(.+?)\*\*', '<b>$1</b>')
    $h = [regex]::Replace($h, '//(.+?)//', '<i>$1</i>')
    # !!vermell!!: peticio de l'usuari (setembre 2026) per al "Com presentar la
    # documentacio" del correu. El mateix vermell a docs\correu.js.
    $h = [regex]::Replace($h, '!!(.+?)!!', '<span style="color:#C00000">$1</span>')

    for ($i = 0; $i -lt $urls.Count; $i++) {
        $u = [string]$urls[$i]
        $h = $h.Replace(([char]1 + [string]$i + [char]1), ('<a href="' + $u + '">' + $u + '</a>'))
    }
    return $h.Replace("`r`n","`n").Replace("`n",'<br>')
}

# --- Unitats -------------------------------------------------------------------
# 96 px per polzada: 1 cm = 37,8 px i 1 twip = 1/15 px. Els NUMEROS van sempre
# amb punt decimal (cultura invariant): en un Windows en catala 6.5 s'escriuria
# "6,5" i el CSS el descartaria.
function _CmAPx([double]$cm)   { return [int][Math]::Round($cm * 96 / 2.54) }
function _TwipsAPx([double]$tw) { return [int][Math]::Round($tw / 15) }
function _TwipsAPt([double]$tw) { return [Math]::Round($tw / 20, 1) }
function _CorreuNum($x) { return ([double]$x).ToString([Globalization.CultureInfo]::InvariantCulture) }

# --- Les mides, des de $ReportFormatConfig --------------------------------------
# Torna un objecte que es serialitza TAL QUAL a docs\dades\correu-format.json.
# Totes les entrades de Blocs porten totes les claus, perque el mobil no hagi de
# tenir valors per defecte propis (serien una segona copia).
function _CorreuFormat {
    $c = $Script:ReportFormatConfig
    $alinea = if ([int]$c.BodyAlignment -eq 3) { 'justify' } else { 'left' }
    $m = {
        param($esqCm, $penjatCm, $abansPt, $despresPt, [string]$al, $midaPt)
        return [ordered]@{
            Esq     = (_CmAPx ([double]$esqCm))
            Penjat  = (_CmAPx ([double]$penjatCm))
            Abans   = [double]$abansPt
            Despres = [double]$despresPt
            Alinea  = $(if ($al) { $al } else { $alinea })
            MidaPt  = [double]$midaPt
        }
    }
    $url = [double]$c.UrlFontSize
    $blocs = [ordered]@{
        seccio        = (& $m $c.SectionIndentCm 0 0 0 '' 0)
        subseccio     = (& $m $c.SubsectionIndentCm 0 0 0 '' 0)
        item          = (& $m $c.ItemIndentCm 0 0 0 '' 0)
        cos           = (& $m $c.ItemIndentCm 0 0 0 '' 0)
        cosFill       = (& $m $c.ChildIndentCm 0 0 0 '' 0)
        enllac        = (& $m $c.ItemIndentCm 0 0 0 '' $url)
        enllacFill    = (& $m $c.ChildIndentCm 0 0 0 '' $url)
        pic           = (& $m $c.BulletIndentCm $c.BulletHangCm $c.BulletSpaceBeforePt 0 '' 0)
        picPrimer     = (& $m $c.BulletIndentCm $c.BulletHangCm $c.PrimerSubpuntSpaceBeforePt 0 '' 0)
        picFill       = (& $m $c.BulletChildIndentCm $c.BulletChildHangCm $c.BulletSpaceBeforePt 0 '' 0)
        picFillPrimer = (& $m $c.BulletChildIndentCm $c.BulletChildHangCm $c.PrimerSubpuntSpaceBeforePt 0 '' 0)
        conclusiocap  = (& $m 0 0 0 $c.ConclusionHeaderSpaceAfterPt 'center' 0)
        conclusio     = (& $m $c.ConclusionIndentCm 0 0 $c.ConclusionSpaceAfterPt '' 0)
    }
    $aire = [ordered]@{}
    foreach ($k in @($Script:AireFlagPerClau.Keys | Sort-Object)) { $aire[$k] = [bool](Test-FormatAire $k) }
    return [ordered]@{
        Lletra         = $Script:CorreuLletra
        MidaPt         = [double]$c.BodyFontSize
        InterlineatPct = [int][Math]::Round([double]$c.BaseLineSpacing * 100)
        Alinea         = $alinea
        EtiquetaPx     = (_TwipsAPx $Script:CorreuCapcaleraTwips)
        Aire           = $aire
        Blocs          = $blocs
    }
}

# --- El contracte d'un paragraf (el mateix a docs\correu.js) ------------------
# Tot el que es pinta passa per aqui, el del mobil i el del PC. Cada <p> i cada
# <td> porta la lletra: l'Outlook d'escriptori no l'hereta dins de les taules.
function _CorreuBase($fmt) {
    return ('font-family:' + [string]$fmt.Lletra + ';font-size:' + (_CorreuNum $fmt.MidaPt) +
            'pt;line-height:' + [string]$fmt.InterlineatPct + '%;color:#000000')
}

function _CorreuBuitHtml($fmt) {
    return ('<p style="' + (_CorreuBase $fmt) + ';margin:0">&nbsp;</p>')
}

# $m = @{ Esq; Penjat; Abans; Despres; Alinea } (px, px, pt, pt).
# Amb $abansTab (el que va abans del tabulador: el pic, l'etiqueta) el paragraf
# es pinta com una fila de taula: es l'unica manera que l'Outlook respecti una
# sagnia francesa (no fa cas del display:inline-block ni de les tabulacions).
function _CorreuParagrafHtml([string]$contingut, $m, $fmt, $abansTab = $null) {
    $base = _CorreuBase $fmt
    $esq = [int]$m.Esq; $pen = [int]$m.Penjat
    $marge = (_CorreuNum $m.Abans) + 'pt 0 ' + (_CorreuNum $m.Despres) + 'pt '
    if ($null -ne $abansTab -and $pen -gt 0) {
        return ('<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:' +
                $marge + [string]($esq - $pen) + 'px"><tr><td style="' + $base + ';width:' + [string]$pen +
                'px;padding:0;vertical-align:top">' + [string]$abansTab + '</td><td style="' + $base +
                ';padding:0;vertical-align:top;text-align:' + [string]$m.Alinea + '">' + $contingut + '</td></tr></table>')
    }
    if ($null -ne $abansTab) { $contingut = [string]$abansTab + ' ' + $contingut }
    $st = $base + ';margin:' + $marge + [string]$esq + 'px'
    if ($pen -ne 0) { $st += ';text-indent:' + [string](-$pen) + 'px' }
    $st += ';text-align:' + [string]$m.Alinea
    return ('<p style="' + $st + '">' + $contingut + '</p>')
}

function _CorreuEnllacHtml([string]$url, $midaPt) {
    $u = _EscHtml $url
    # word-break: un enllac de 150 caracters no es parteix sol i sortia del
    # marge (al mobil, fora de la pantalla).
    return ('<a href="' + $u + '" style="font-size:' + (_CorreuNum $midaPt) + 'pt;word-break:break-all">' + $u + '</a>')
}

# --- Blocs -> HTML (el mateix bucle que Write-Informe, MotorInforme.ps1) ------
# Nomes els tipus de bloc que poden sortir en un correu de requeriments. Un de
# desconegut PETA, com a Write-Informe: val mes aixo que un correu al qual li
# falta un tros sense que ho digui ningu.
function _CorreuBlocsAHtml($blocs, $fmt) {
    $estat = @{ PrimerFill = $false; Escrits = 0; UltimBuit = $true; Sb = (New-Object System.Text.StringBuilder) }
    _CorreuBlocsRec $blocs $fmt $estat
    return $estat.Sb.ToString()
}

function _CorreuBlocsRec($blocs, $fmt, $estat) {
    $sb = $estat.Sb
    $buit = { [void]$sb.Append((_CorreuBuitHtml $fmt)); $estat.UltimBuit = $true }
    foreach ($b in @($blocs)) {
        if ($null -eq $b) { continue }
        $t = [string]$b.T
        $fill = [bool]$b.Fill
        if ($t -eq 'unitat') {
            $abans = $estat.Escrits
            $estat.PrimerFill = $true
            _CorreuBlocsRec $b.Blocs $fmt $estat
            if ($estat.Escrits -gt $abans -and [bool]$fmt.Aire['item']) { & $buit }
            continue
        }
        if ($t -eq 'aire') {
            if ([bool]$fmt.Aire[([string]$b.Clau).ToLower()]) { & $buit }
            continue
        }
        if ($t -eq 'espai') { & $buit; continue }
        if ($t -eq 'separa') { if (-not $estat.UltimBuit) { & $buit }; continue }

        # $mb i no $B: en PowerShell $B i $b son la MATEIXA variable, i es
        # trepitjava el bloc que s'esta pintant.
        $mb = $fmt.Blocs
        $h = switch ($t) {
            'seccio'       { _CorreuParagrafHtml (_EscHtml ([string]$b.Text).ToUpper()) $mb.seccio $fmt }
            'subseccio'    { _CorreuParagrafHtml (_EscHtml ([string]$b.Text)) $mb.subseccio $fmt }
            'item'         { _CorreuParagrafHtml ('<b>' + (_EscHtml ([string]$b.Num)) + '</b> ' + (_TextToHtml ([string]$b.Text))) $mb.item $fmt }
            'cos' {
                $x = _TextToHtml ([string]$b.Text)
                if ([bool]$b.Negreta) { $x = '<b>' + $x + '</b>' }
                _CorreuParagrafHtml $x $(if ($fill) { $mb.cosFill } else { $mb.cos }) $fmt
            }
            'enllac' {
                $mm = if ($fill) { $mb.enllacFill } else { $mb.enllac }
                _CorreuParagrafHtml (_CorreuEnllacHtml ([string]$b.Url) $mm.MidaPt) $mm $fmt
            }
            'pic' {
                $primer = if ($b.ContainsKey('First')) { [bool]$b.First } else { [bool]$estat.PrimerFill }
                $mm = if ($fill) { if ($primer) { $mb.picFillPrimer } else { $mb.picFill } }
                      else       { if ($primer) { $mb.picPrimer } else { $mb.pic } }
                $estat.PrimerFill = $false
                _CorreuParagrafHtml (_TextToHtml ([string]$b.Text)) $mm $fmt ([string][char]0x2022)
            }
            'conclusiocap' { _CorreuParagrafHtml ('<b>' + (_EscHtml ([string]$b.Text)) + '</b>') $mb.conclusiocap $fmt }
            'conclusio'    { _CorreuParagrafHtml (_TextToHtml ([string]$b.Text)) $mb.conclusio $fmt }
            default { throw ("_CorreuBlocsAHtml: tipus de bloc desconegut '" + $t + "'") }
        }
        [void]$sb.Append([string]$h)
        $estat.Escrits++
        $estat.UltimBuit = $false
    }
}

# --- La capcalera del correu: les linies de '0 CAPCALERA' ----------------------
# Les linies AMB ETIQUETA del bloc generic (el primer, el de REQ1): ID GIA,
# Exp. Num, Adreca, Activitat, Titular i Objecte. Torna @(@{Etiqueta; Plantilla})
# amb la plantilla tal com es a la capcalera ("<<ID_GIA>>"). PURA.
function _CorreuCapcaleraLinies($capJson) {
    $out = @()
    if ($null -eq $capJson) { return $out }
    $gen = @(@($capJson.nodes) | Where-Object { [string]$_.tipus -eq 'seccio' }) | Select-Object -First 1
    if ($null -eq $gen) { return $out }
    foreach ($f in @($gen.fills)) {
        if ([string]$f.tipus -ne 'etiqueta') { continue }
        $pl = ''
        foreach ($p in @($f.cos)) { foreach ($r in @($p.runs)) { $pl += [string]$r.t } }
        $out += [pscustomobject]@{ Etiqueta = ([string]$f.titol).Trim(); Plantilla = $pl.Trim() }
    }
    return $out
}

# Omple <<CLAU>> amb els valors d'un hashtable/objecte. PURA.
function _CorreuOmplePlantilla([string]$plantilla, $valors) {
    return [regex]::Replace([string]$plantilla, '<<\s*([A-Za-z0-9_]+)\s*>>', {
        param($mt)
        $k = $mt.Groups[1].Value
        if ($null -eq $valors) { return '' }
        if ($valors -is [System.Collections.IDictionary]) { if ($valors.Contains($k)) { return [string]$valors[$k] }; return '' }
        if ($valors.PSObject.Properties.Name -contains $k) { return [string]$valors.$k }
        return ''
    })
}

# El TITOL que va sota les linies de la capcalera ("INFORME"): el primer text
# sense etiqueta DESPRES de l'ultima linia amb etiqueta del bloc generic. PURA.
function _CorreuCapcaleraTitol($capJson) {
    if ($null -eq $capJson) { return '' }
    $gen = @(@($capJson.nodes) | Where-Object { [string]$_.tipus -eq 'seccio' }) | Select-Object -First 1
    if ($null -eq $gen) { return '' }
    $fills = @($gen.fills)
    $ultima = -1
    for ($i = 0; $i -lt $fills.Count; $i++) { if ([string]$fills[$i].tipus -eq 'etiqueta') { $ultima = $i } }
    for ($i = $ultima + 1; $i -lt $fills.Count; $i++) {
        if ([string]$fills[$i].tipus -ne 'text') { continue }
        $t = ''
        foreach ($p in @($fills[$i].cos)) { foreach ($r in @($p.runs)) { $t += [string]$r.t } }
        if (-not [string]::IsNullOrWhiteSpace($t)) { return $t.Trim() }
    }
    return ''
}

# Les linies amb valor -> la taula de la capcalera. Una linia SENSE valor no
# surt (un "Objecte:" buit al correu no diu res). $linies = @(@{Etiqueta; Valor}).
# Amb $titol, a sota, una linia en blanc i el titol ("INFORME") centrat i en
# negreta, com a l'informe (peticio de l'usuari, setembre 2026).
function _CorreuCapcaleraHtml($linies, $fmt, [string]$titol = '') {
    $sb = New-Object System.Text.StringBuilder
    $px = [int]$fmt.EtiquetaPx
    $m = @{ Esq = $px; Penjat = $px; Abans = 0; Despres = 0; Alinea = 'left' }
    foreach ($l in @($linies)) {
        if ([string]::IsNullOrWhiteSpace([string]$l.Valor)) { continue }
        [void]$sb.Append((_CorreuParagrafHtml (_EscHtml ([string]$l.Valor).Trim()) $m $fmt ('<b>' + (_EscHtml ([string]$l.Etiqueta)) + '</b>')))
    }
    if (-not [string]::IsNullOrWhiteSpace($titol)) {
        [void]$sb.Append((_CorreuBuitHtml $fmt))
        [void]$sb.Append((_CorreuParagrafHtml ('<b>' + (_EscHtml $titol.Trim()) + '</b>') @{ Esq = 0; Penjat = 0; Abans = 0; Despres = 0; Alinea = 'center' } $fmt))
    }
    return $sb.ToString()
}

# --- L'origen: "Doc. aportada..." / "Visita inspeccio..." al reves -------------
# Llegeix el text de la linia "Objecte:" d'un informe i en treu el tipus i les
# dades, amb les MATEIXES plantilles que el munten ($Script:OrigenPlantilles,
# MotorInforme.ps1). Un text que no encaixa amb cap torna ORIGEN_TIPUS = ''.
function _OrigenDesDeText([string]$text) {
    $res = @{ ORIGEN_TIPUS = ''; NUM_ANOTACIO = ''; DATA_ANOTACIO = ''; DATA_INSPECCIO = '' }
    $t = ([string]$text).Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return $res }
    foreach ($tipus in @('doc', 'insp')) {
        $pl = [string]$Script:OrigenPlantilles[$tipus]
        $rx = '^' + [regex]::Replace([regex]::Escape($pl), '<<([A-Za-z0-9_]+)>>', '(?<$1>.*?)') + '$'
        # L'apostrof: el Word pot haver-lo canviat (tipografic o recte).
        $rx = $rx.Replace([string][char]0x2019, "[" + [char]0x2019 + "']")
        $mt = [regex]::Match($t, $rx)
        # "Visita inspeccio " sense data: el valor ve retallat i l'espai que
        # precedeix la data buida ja no hi es.
        if (-not $mt.Success) { $mt = [regex]::Match($t + ' ', $rx) }
        if (-not $mt.Success) { continue }
        $res.ORIGEN_TIPUS = $tipus
        foreach ($g in @('NUM_ANOTACIO', 'DATA_ANOTACIO', 'DATA_INSPECCIO')) {
            if ($mt.Groups[$g].Success) { $res[$g] = $mt.Groups[$g].Value.Trim() }
        }
        return $res
    }
    return $res
}

# La frase que introdueix els requeriments, segons l'origen (peticio de
# l'usuari, setembre 2026): un requeriment sobre documentacio aportada NO pot
# dir "s'han detectat a la visita". Les frases son a email-textos.json.
#   doc  amb num. i data d'anotacio -> introDoc
#   doc  sense                      -> introDocSenseAnotacio
#   insp                            -> introInsp (sense data: la d'avui)
# PURA: $avui arriba des de fora.
function _CorreuIntro($textos, $origen, [string]$avui) {
    $o = $origen
    $tipus = [string]$o.ORIGEN_TIPUS
    $clau = if ($tipus -eq 'insp') { 'introInsp' }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$o.NUM_ANOTACIO) -and -not [string]::IsNullOrWhiteSpace([string]$o.DATA_ANOTACIO)) { 'introDoc' }
            else { 'introDocSenseAnotacio' }
    $frase = ''
    if ($textos -is [System.Collections.IDictionary]) { if ($textos.Contains($clau)) { $frase = [string]$textos[$clau] } }
    elseif ($null -ne $textos -and $textos.PSObject.Properties[$clau]) { $frase = [string]$textos.$clau }
    $dataInsp = [string]$o.DATA_INSPECCIO
    if ([string]::IsNullOrWhiteSpace($dataInsp)) { $dataInsp = $avui }
    return $frase.Replace('{NUM_ANOTACIO}', [string]$o.NUM_ANOTACIO).Replace('{DATA_ANOTACIO}', [string]$o.DATA_ANOTACIO).Replace('{DATA_INSPECCIO}', $dataInsp)
}

# --- El cos del correu (email-textos.json) -> HTML -----------------------------
# Una linia = un paragraf i una linia buida = un paragraf buit, com a l'informe.
# {CAPCALERA} i {REQUERIMENTS} han d'anar SOLS a la seva linia: son blocs, no
# text. La resta de variables ({INTRO}, {ID_GIA}...) les omple $omple.
function _CorreuCosAHtml([string]$cos, [scriptblock]$omple, [string]$capHtml, [string]$reqHtml, $fmt) {
    $sb = New-Object System.Text.StringBuilder
    $m = $fmt.Blocs.cos
    foreach ($ln in ((([string]$cos) -replace "`r`n", "`n") -split "`n")) {
        $t = $ln.Trim()
        if ($t -eq '{CAPCALERA}')        { [void]$sb.Append($capHtml); continue }
        if ($t -eq '{REQUERIMENTS}')     { [void]$sb.Append($reqHtml); continue }
        if ([string]::IsNullOrWhiteSpace($t)) { [void]$sb.Append((_CorreuBuitHtml $fmt)); continue }
        [void]$sb.Append((_CorreuParagrafHtml (_TextToHtml ([string](& $omple $ln))) $m $fmt))
    }
    return ('<div style="' + (_CorreuBase $fmt) + '">' + $sb.ToString() + '</div>')
}

# --- El .docx d'un informe -> capcalera + cos en HTML (SENSE Word) -------------
# Clau per comparar etiquetes ("Exp. Num: " = "exp. num"). PURA.
function _CorreuClauEtiqueta([string]$e) {
    return ((_NormalitzaText ([string]$e)).Trim().TrimEnd(':').Trim())
}

# Els paragrafs FILLS DIRECTES del <w:body> (les taules, fora), amb el format
# que ens cal: sagnia, espais, alineat i els runs amb negreta/cursiva/subratllat
# i mida. PURA (rep el text de word/document.xml).
function _CorreuParagrafsDocx([string]$xml) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($xml)
    $wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('w', $wns)
    $body = $doc.SelectSingleNode('//w:body', $ns)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $body) { return $out.ToArray() }
    $attr = { param($n, $a) if ($null -eq $n) { return '' }; return [string]$n.GetAttribute($a, $wns) }
    $on = {
        param($n)
        if ($null -eq $n) { return $false }
        $v = [string]$n.GetAttribute('val', $wns)
        return -not ($v -eq '0' -or $v -eq 'false' -or $v -eq 'off' -or $v -eq 'none')
    }
    foreach ($p in $body.ChildNodes) {
        if ($p.LocalName -ne 'p' -or $p.NamespaceURI -ne $wns) { continue }
        $ind = $p.SelectSingleNode('w:pPr/w:ind', $ns)
        $sp  = $p.SelectSingleNode('w:pPr/w:spacing', $ns)
        $jc  = $p.SelectSingleNode('w:pPr/w:jc', $ns)
        $esq = & $attr $ind 'left'; if (-not $esq) { $esq = & $attr $ind 'start' }
        $pen = & $attr $ind 'hanging'
        $pri = & $attr $ind 'firstLine'
        $runs = New-Object System.Collections.ArrayList
        foreach ($r in $p.SelectNodes('.//w:r', $ns)) {
            $rPr = $r.SelectSingleNode('w:rPr', $ns)
            $sz = & $attr ($r.SelectSingleNode('w:rPr/w:sz', $ns)) 'val'
            $fmtR = @{
                B  = [bool](& $on ($r.SelectSingleNode('w:rPr/w:b', $ns)))
                I  = [bool](& $on ($r.SelectSingleNode('w:rPr/w:i', $ns)))
                U  = [bool](& $on ($r.SelectSingleNode('w:rPr/w:u', $ns)))
                Sz = $(if ($sz) { [double]$sz / 2 } else { 0 })
            }
            foreach ($n in $r.ChildNodes) {
                $txt = switch ($n.LocalName) {
                    't'              { $n.InnerText }
                    'tab'            { "`t" }
                    'br'             { "`n" }
                    'cr'             { "`n" }
                    'noBreakHyphen'  { '-' }
                    default          { $null }
                }
                if ($null -eq $txt) { continue }
                [void]$runs.Add(@{ Text = [string]$txt; B = $fmtR.B; I = $fmtR.I; U = $fmtR.U; Sz = $fmtR.Sz })
            }
        }
        $text = (@($runs | ForEach-Object { $_.Text }) -join '')
        $jv = & $attr $jc 'val'
        $alinea = switch ($jv) {
            'center' { 'center' } 'right' { 'right' } 'end' { 'right' }
            'left' { 'left' } 'start' { 'left' }
            default { $null }
        }
        [void]$out.Add(@{
            Text    = $text
            Runs    = $runs.ToArray()
            Esq     = $(if ($esq) { _TwipsAPx ([double]$esq) } else { 0 })
            Penjat  = $(if ($pen) { _TwipsAPx ([double]$pen) } elseif ($pri) { -(_TwipsAPx ([double]$pri)) } else { 0 })
            Abans   = $(if (& $attr $sp 'before') { _TwipsAPt ([double](& $attr $sp 'before')) } else { 0 })
            Despres = $(if (& $attr $sp 'after')  { _TwipsAPt ([double](& $attr $sp 'after'))  } else { 0 })
            Alinea  = $alinea
        })
    }
    return $out.ToArray()
}

# Escapa i enllaca els URLs d'un tros de text del .docx (SENSE **/ //: el
# text del Word no porta marques, i un "//" hi es literal).
function _CorreuTextDocxHtml([string]$t) {
    $h = _EscHtml $t
    return [regex]::Replace($h, '(https?://[^\s<]+)', { param($mt) '<a href="' + $mt.Value + '">' + $mt.Value + '</a>' })
}

# Els runs d'un paragraf -> HTML. Els runs seguits amb el mateix format
# s'ajunten ABANS d'escapar: el Word parteix el text on vol, i un URL partit en
# dos runs sortiria com dos enllacos trencats.
function _CorreuRunsHtml($runs, $fmt) {
    $sb = New-Object System.Text.StringBuilder
    $grups = New-Object System.Collections.ArrayList
    foreach ($r in @($runs)) {
        $n = $grups.Count
        if ($n -gt 0) {
            $g = $grups[$n - 1]
            if ($g.B -eq $r.B -and $g.I -eq $r.I -and $g.U -eq $r.U -and $g.Sz -eq $r.Sz) { $g.Text += [string]$r.Text; continue }
        }
        [void]$grups.Add(@{ Text = [string]$r.Text; B = $r.B; I = $r.I; U = $r.U; Sz = $r.Sz })
    }
    foreach ($g in $grups) {
        $h = (_CorreuTextDocxHtml ([string]$g.Text).Replace("`t", ' ')).Replace("`n", '<br>')
        if ($g.U) { $h = '<u>' + $h + '</u>' }
        if ($g.I) { $h = '<i>' + $h + '</i>' }
        if ($g.B) { $h = '<b>' + $h + '</b>' }
        if ($g.Sz -gt 0 -and [double]$g.Sz -ne [double]$fmt.MidaPt) { $h = '<span style="font-size:' + (_CorreuNum $g.Sz) + 'pt">' + $h + '</span>' }
        [void]$sb.Append($h)
    }
    return $sb.ToString()
}

# Un paragraf del .docx -> HTML, amb el MATEIX contracte que els blocs.
# UN REQUERIMENT (o un sub-punt) TOT EN NEGRETA ES LA MARCA DEL SEGUIMENT, no
# format de l'informe. Els seguiments d'abans posaven en negreta el punt sencer
# mentre era pendent (_InferResolvedFromBold en llegeix encara la marca); al
# correu nomes hi ha d'anar el comentari ("No s'aporta.") (peticio de l'usuari,
# setembre 2026). Es treu la negreta i es torna a posar NOMES al numero, que es
# com surt a REQ1. Un punt amb una part en negreta (**...** del cataleg) no es
# toca. PURA.
function _CorreuSenseNegretaSeguiment($p) {
    $t = [string]$p.Text
    $mNum = [regex]::Match($t, '^\s*\d+\.\s')
    $esPunt = $mNum.Success -or $t.TrimStart().StartsWith([string][char]0x2022)
    if (-not $esPunt) { return $p.Runs }
    $ambText = @(@($p.Runs) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Text) })
    if ($ambText.Count -eq 0 -or @($ambText | Where-Object { -not $_.B }).Count -gt 0) { return $p.Runs }
    $out = New-Object System.Collections.ArrayList
    $queda = if ($mNum.Success) { $mNum.Length } else { 0 }
    foreach ($r in @($p.Runs)) {
        $txt = [string]$r.Text
        if ($queda -gt 0) {
            $n = [Math]::Min($queda, $txt.Length)
            [void]$out.Add(@{ Text = $txt.Substring(0, $n); B = $true; I = $r.I; U = $r.U; Sz = $r.Sz })
            $queda -= $n
            $txt = $txt.Substring($n)
            if ($txt.Length -eq 0) { continue }
        }
        [void]$out.Add(@{ Text = $txt; B = $false; I = $r.I; U = $r.U; Sz = $r.Sz })
    }
    return $out.ToArray()
}

function _CorreuParagrafDocxHtml($p, $fmt) {
    if ([string]::IsNullOrWhiteSpace([string]$p.Text)) { return (_CorreuBuitHtml $fmt) }
    $p = @{ Text = $p.Text; Runs = @(_CorreuSenseNegretaSeguiment $p); Esq = $p.Esq; Penjat = $p.Penjat
            Abans = $p.Abans; Despres = $p.Despres; Alinea = $p.Alinea }
    $m = @{ Esq = $p.Esq; Penjat = $p.Penjat; Abans = $p.Abans; Despres = $p.Despres
            Alinea = $(if ($p.Alinea) { $p.Alinea } else { [string]$fmt.Alinea }) }
    $t = ([string]$p.Text).Trim()
    if ($t -match '^https?://\S+$') {
        $mida = [double]$fmt.MidaPt
        foreach ($r in @($p.Runs)) { if (-not [string]::IsNullOrWhiteSpace([string]$r.Text)) { if ($r.Sz -gt 0) { $mida = [double]$r.Sz }; break } }
        return (_CorreuParagrafHtml (_CorreuEnllacHtml $t $mida) $m $fmt)
    }
    # Sagnia francesa amb tabulador (el pic d'un sub-punt): el que va abans del
    # tabulador a la primera columna, com fan els blocs 'pic'.
    if ([int]$p.Penjat -gt 0 -and ([string]$p.Text).Contains("`t")) {
        $abans = New-Object System.Collections.ArrayList
        $despres = New-Object System.Collections.ArrayList
        $vist = $false
        foreach ($r in @($p.Runs)) {
            if (-not $vist -and [string]$r.Text -eq "`t") { $vist = $true; continue }
            if ($vist) { [void]$despres.Add($r) } else { [void]$abans.Add($r) }
        }
        if ($vist) {
            $a = (_CorreuRunsHtml $abans.ToArray() $fmt).Trim()
            return (_CorreuParagrafHtml (_CorreuRunsHtml $despres.ToArray() $fmt) $m $fmt $a)
        }
    }
    return (_CorreuParagrafHtml (_CorreuRunsHtml $p.Runs $fmt) $m $fmt)
}

# EL .DOCX D'UN INFORME -> @{ Capcalera; Html }. PURA (rep el document.xml).
#   Capcalera: etiqueta normalitzada -> valor, de les linies de dalt de tot
#              (fins a "INFORME"): "ID GIA:<tab>1398" -> 'id gia' = '1398'.
#   Html:      el cos, DES DE DESPRES de la frase "...deficiencies... esmenar"
#              (la del cataleg: el correu ja porta la seva, {INTRO}) i fins a la
#              frase de tancament ("Ho poso al seu coneixement", "Cornella de
#              Llobregat,") o fins al titol CONCLUSIONS: les conclusions NO van
#              al correu (ni al del mobil). Sense la frase, des de despres
#              d'"INFORME".
# Abans es llegia amb el Word i es comencava al primer paragraf en MAJUSCULES:
# "ID GIA: 1398" ho es, i el correu repetia tota la capcalera, "INFORME" i la
# nota de l'Ordenanca.
function _CorreuDocxLlegeix([string]$xml, $fmt) {
    $ps = @(_CorreuParagrafsDocx $xml)
    $cap = [ordered]@{}
    $iInforme = -1
    for ($i = 0; $i -lt $ps.Count; $i++) {
        $t = ([string]$ps[$i].Text).Trim()
        if ($t -ceq 'INFORME') { $iInforme = $i; break }
        $k = ([string]$ps[$i].Text).IndexOf("`t")
        if ($k -gt 0) {
            $e = _CorreuClauEtiqueta (([string]$ps[$i].Text).Substring(0, $k))
            if ($e -and -not $cap.Contains($e)) { $cap[$e] = ([string]$ps[$i].Text).Substring($k + 1).Replace("`t", ' ').Trim() }
        }
    }
    $ini = $iInforme + 1
    for ($i = $ini; $i -lt $ps.Count; $i++) {
        $n = _NormalitzaText ([string]$ps[$i].Text)
        if ($n -match 'defici' -and $n -match 'esmenar') { $ini = $i + 1; break }
    }
    $fi = $ps.Count
    for ($i = $ini; $i -lt $ps.Count; $i++) {
        $t = [string]$ps[$i].Text
        $n = (_NormalitzaText $t).Trim()
        if ((_EsFraseTancament $t) -or $n.StartsWith('cornella de llobregat,') -or $n -match '^conclusi(o|ons)$') { $fi = $i; break }
    }
    # Fora els paragrafs buits dels extrems.
    $buit = { param($j) [string]::IsNullOrWhiteSpace([string]$ps[$j].Text) }
    while ($ini -lt $fi -and (& $buit $ini)) { $ini++ }
    while ($fi -gt $ini -and (& $buit ($fi - 1))) { $fi-- }
    $sb = New-Object System.Text.StringBuilder
    for ($i = $ini; $i -lt $fi; $i++) { [void]$sb.Append((_CorreuParagrafDocxHtml $ps[$i] $fmt)) }
    return @{ Capcalera = $cap; Html = $sb.ToString() }
}
