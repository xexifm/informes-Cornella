#requires -Version 5.1
<#
  ContactesExtraccio.ps1 - QUE DIU CADA DOCUMENT de qui es l'activitat.

  Per al repas de contactes d'"Actualitzar base" (vegeu ContactesDb.ps1 i la
  seccio "Contactes de les activitats" de documentacio/eines.md). L'usuari ho
  va mesurar: gairebe totes les dades surten de TRES tipus de document, petits,
  i es trien PEL NOM abans d'obrir res (_CtTipusDocument):

    etram         *etram-tramit*.xml (namespace http://www.aoc.cat/etram)
    instancia     la instancia generica / l'esmena de la seu (PDF amb text)
    autoritzacio  l'autoritzacio de representacio (PDF): qui es el
                  representant legal DE VERITAT i a qui autoritza
    xmlantic      els XML antics (*XML_TRAMIT*, *sicres3*), nomes si no hi ha res mes

  Els certificats, les memories i els projectes NO es llegeixen: pesen molt i
  gairebe nomes donen el tecnic.

  Cada lector torna el mateix registre (_CtDocNou): interessat, representant,
  establiment, autoritzacio i el que diu de la data, el titol, el GIA i
  l'expedient. Les persones (_CtPersona) porten nom, nif, email, telefon i
  mobil; els telefons es reparteixen PEL PREFIX (6/7 mobil, 8/9 fix), que es
  el mateix criteri amb que es compara l'Excel (ContactesRegles.ps1).

  ON VA EL TEXT: tot aixo es pur (rep el text, no obre res) i es prova amb
  documents INVENTATS. Llegir els fitxers ho fa ContactesDb.ps1. Res de dades
  reals ni aqui ni a les proves.
  NOMES DEFINEIX FUNCIONS.
#>

# ----------------------------------------------------------------------------
# PECES COMUNES (PURES)
# ----------------------------------------------------------------------------

# Llegeix un camp d'un hashtable O d'un objecte (el que torna del JSON). PURA.
function _CtV($o, [string]$k) {
    if ($null -eq $o) { return $null }
    if ($o -is [System.Collections.IDictionary]) { if ($o.Contains($k)) { return $o[$k] } return $null }
    $p = $o.PSObject.Properties[$k]
    if ($null -eq $p) { return $null }
    return $p.Value
}

# El text "pla" per buscar-hi: minuscules, sense accents i amb els apostrofs
# rectes, CARACTER A CARACTER (la mateixa llargada que l'original). Aixi una
# posicio trobada al text pla serveix per tallar l'original, amb les seves
# majuscules i accents. PURA.
function _CtPla([string]$s) {
    if ($null -eq $s) { return '' }
    # El cas de quasi sempre (nomes ASCII sense apostrofs rars): de cop.
    if ($s -notmatch '[^\x20-\x7E\t\r\n]') { return $s.ToLowerInvariant() }
    $sb = New-Object System.Text.StringBuilder ($s.Length)
    foreach ($ch in $s.ToCharArray()) {
        if ([int]$ch -lt 128) { [void]$sb.Append([char]::ToLowerInvariant($ch)); continue }
        if ($ch -eq [char]0x2019 -or $ch -eq [char]0x2018 -or $ch -eq [char]0x00B4 -or $ch -eq [char]0x0060) { [void]$sb.Append("'"); continue }
        if ($ch -eq [char]0x00A0) { [void]$sb.Append(' '); continue }
        if ($ch -eq [char]0x00B7 -or $ch -eq [char]0x2027) { [void]$sb.Append('.'); continue }
        $d = ([string]$ch).Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
        if ($d.Length -ne 1) { $d = [string]$ch }
        [void]$sb.Append($d.ToLowerInvariant())
    }
    return $sb.ToString()
}

# Per comparar noms: pla, sense signes i amb els espais junts. PURA (amb
# memoria: les regles comparen cada persona amb els ~250 tecnics coneguts, i
# normalitzar-los cada vegada feia el repas lent; mesurat).
$Script:CtMemoNom = @{}
function _CtNomNorm([string]$s) {
    if ($null -eq $s) { $s = '' }
    if ($Script:CtMemoNom.ContainsKey($s)) { return $Script:CtMemoNom[$s] }
    $t = (_CtPla $s) -replace "[^a-z0-9 ]", ' '
    $r = (($t -replace '\s+', ' ').Trim())
    if ($Script:CtMemoNom.Count -gt 50000) { $Script:CtMemoNom.Clear() }
    $Script:CtMemoNom[$s] = $r
    return $r
}

# Les paraules d'un nom que compten (sense "de", "la", "i", la forma juridica).
$Script:CtNomBuides = @('de', 'del', 'la', 'les', 'el', 'els', 'i', 'y', 'sl', 'slu', 'sa', 'sau', 'scp', 'sccl', 'cb', 'sll', 'slne', 'sc')
function _CtNomParaules([string]$s) {
    return @((_CtNomNorm $s) -split ' ' | Where-Object { $_.Length -ge 2 -and $Script:CtNomBuides -notcontains $_ })
}

# Dos noms son de la mateixa persona? Les paraules del curt son totes al llarg
# (i n'hi ha com a minim dues, o es el mateix nom sencer): "Maria Exemple" i
# "MARIA EXEMPLE PROVA" si; "Maria" sola, no. PURA.
function _CtMateixNom([string]$a, [string]$b) {
    $pa = @(_CtNomParaules $a); $pb = @(_CtNomParaules $b)
    if ($pa.Count -eq 0 -or $pb.Count -eq 0) { return $false }
    if ((_CtNomNorm $a) -eq (_CtNomNorm $b)) { return $true }
    $curt = $pa; $llarg = $pb
    if ($pa.Count -gt $pb.Count) { $curt = $pb; $llarg = $pa }
    if ($curt.Count -lt 2) { return $false }
    foreach ($w in $curt) { if ($llarg -notcontains $w) { return $false } }
    return $true
}

function _CtEmailNet([string]$e) {
    $t = ([string]$e).Trim().Trim('.', ';', ',', '<', '>', '(', ')').ToLowerInvariant()
    if ($t -match '^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$') { return $t }
    return ''
}

# Nomes les xifres, sense el prefix d'Espanya. PURA.
function _CtTelefonNet([string]$t) {
    $d = ([string]$t) -replace '[^\d]', ''
    if ($d.Length -eq 13 -and $d.StartsWith('0034')) { $d = $d.Substring(4) }
    elseif ($d.Length -eq 11 -and $d.StartsWith('34')) { $d = $d.Substring(2) }
    return $d
}

# 'mobil' (6/7), 'fix' (8/9) o '' (no es un telefon de 9 xifres). PURA.
function _CtTipusTelefon([string]$t) {
    $d = _CtTelefonNet $t
    if ($d -match '^[67]\d{8}$') { return 'mobil' }
    if ($d -match '^[89]\d{8}$') { return 'fix' }
    return ''
}

function _CtNifNet([string]$n) { return ((([string]$n).ToUpperInvariant()) -replace '[\s.\-/]', '') }

# 'fisica' (DNI: 8 xifres + lletra; NIE: X/Y/Z + 7 xifres + lletra),
# 'juridica' (NIF d'entitat) o ''. PURA.
function _CtTipusNif([string]$n) {
    $t = _CtNifNet $n
    if ($t -match '^\d{8}[A-Z]$' -or $t -match '^[XYZ]\d{7}[A-Z]$') { return 'fisica' }
    if ($t -match '^[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]$') { return 'juridica' }
    return ''
}

# Totes les adreces, NIF i telefons d'un text, amb la posicio. PURES.
$Script:CtReEmail = '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}'
$Script:CtReNif   = '(?<![A-Za-z0-9])(?:\d{8}\s?-?\s?[A-Za-z]|[XYZxyz]\s?-?\d{7}\s?-?[A-Za-z]|[ABCDEFGHJNPQRSUVWabcdefghjnpqrsuvw]\s?-?\d{7}[0-9A-Ja-j]?)(?![A-Za-z0-9])'
$Script:CtReTel   = '(?<![\d])(?:\+34\s?|0034\s?)?[6789](?:[\s.]?\d){8}(?![\d])'
function _CtTroba([string]$text, [string]$re) {
    $out = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches([string]$text, $re)) { [void]$out.Add(@{ Pos = $m.Index; Fi = $m.Index + $m.Length; Val = $m.Value }) }
    return $out.ToArray()
}

# Una persona: sempre els mateixos camps (i nets). PURA.
function _CtPersona([string]$nom = '', [string]$nif = '', [string]$email = '', [string]$telefon = '', [string]$mobil = '') {
    $p = [ordered]@{ nom = (_CtNetejaNom $nom); nif = (_CtNifNet $nif); email = (_CtEmailNet $email); telefon = ''; mobil = '' }
    if ((_CtTipusNif $p.nif) -eq '') { $p.nif = '' }
    foreach ($t in @($mobil, $telefon)) { _CtPosaTelefon $p $t }
    return $p
}

# Posa un telefon al camp que li toca pel prefix; si aquell ja n'te un, a
# l'altre (dos mobils: el segon va a "telefon"). PURA (toca $p).
function _CtPosaTelefon($p, [string]$t) {
    $d = _CtTelefonNet $t
    $tip = _CtTipusTelefon $d
    if ($tip -eq '') { return }
    if ([string]$p.mobil -eq $d -or [string]$p.telefon -eq $d) { return }
    if ($tip -eq 'mobil' -and [string]$p.mobil -eq '') { $p.mobil = $d; return }
    if ([string]$p.telefon -eq '') { $p.telefon = $d; return }
    if ([string]$p.mobil -eq '') { $p.mobil = $d }
}

function _CtPersonaBuida($p) {
    if ($null -eq $p) { return $true }
    foreach ($k in 'nom', 'nif', 'email', 'telefon', 'mobil') { if ([string](_CtV $p $k) -ne '') { return $false } }
    return $true
}

# Un nom tret d'una frase: sense tractaments, sense "major d'edat" ni el DNI
# enganxat, i amb els espais junts. PURA.
function _CtNetejaNom([string]$s) {
    $t = ([string]$s) -replace '\s+', ' '
    $t = $t.Trim().Trim(',', ';', ':', '.', '-', ' ')
    $pla = _CtPla $t
    foreach ($re in @('\s+major d.edat.*$', '\s+mayor de edad.*$', ',?\s+amb (dni|nif|nie|d\.n\.i).*$', ',?\s+con (dni|nif|nie|d\.n\.i).*$', '\s+(dni|nif|nie)\s*:?\s*[0-9xyz].*$')) {
        $m = [regex]::Match($pla, $re)
        if ($m.Success) { $t = $t.Substring(0, $m.Index); $pla = $pla.Substring(0, $m.Index) }
    }
    $m = [regex]::Match($pla, '^(sr\.?|sra\.?|senyor|senyora|en|na|d\.?\s*/\s*dna\.?|d\.|dna\.|don|dona)\s+')
    if ($m.Success) { $t = $t.Substring($m.Length) }
    return $t.Trim().Trim(',', ';', ':', '.', '-', ' ')
}

# ----------------------------------------------------------------------------
# QUIN DOCUMENT ES (pel NOM del fitxer, abans d'obrir-lo). PURA.
# ----------------------------------------------------------------------------
function _CtTipusDocument([string]$nom) {
    $n = _CtPla $nom
    $ext = [System.IO.Path]::GetExtension($n)
    if ($n -match 'index electronic') { return '' }
    if ($ext -eq '.xml') {
        if ($n -match 'etram-tramit') { return 'etram' }
        if ($n -match 'xml_tramit' -or $n -match 'sicres3') { return 'xmlantic' }
        return ''
    }
    if ($ext -ne '.pdf') { return '' }
    if ($n -match 'autoritz' -or $n -match 'autoriz' -or $n -match 'representaci') { return 'autoritzacio' }
    if ($n -match 'instancia[ _]generica' -or $n -match 'esmena sol\.?licitud' -or $n -match 'seu\s+oac') { return 'instancia' }
    return ''
}

# El registre que torna cada lector (el mateix per a tots). PURA.
function _CtDocNou([string]$tipus) {
    return [ordered]@{
        tipus = $tipus; data = ''; titol = ''; gia = ''; expedient = ''
        interessat = $null; representant = $null; establiment = $null; autoritzacio = $null
    }
}

# ----------------------------------------------------------------------------
# DATES, GIA I EXPEDIENT DINS D'UN TEXT (PURES)
# ----------------------------------------------------------------------------
$Script:CtMesos = @{ gener = 1; febrer = 2; marc = 3; abril = 4; maig = 5; juny = 6; juliol = 7; agost = 8; setembre = 9; octubre = 10; novembre = 11; desembre = 12
                     enero = 1; febrero = 2; marzo = 3; mayo = 5; junio = 6; julio = 7; agosto = 8; septiembre = 9; setiembre = 9; noviembre = 11; diciembre = 12 }

function _CtDataIso([int]$a, [int]$m, [int]$d) {
    if ($a -lt 1990 -or $a -gt 2100 -or $m -lt 1 -or $m -gt 12 -or $d -lt 1 -or $d -gt 31) { return '' }
    try { return ([datetime]::new($a, $m, $d)).ToString('yyyy-MM-dd') } catch { return '' }
}

# Una data escrita "12 de marc de 2025" / "12 d'abril de 2025" / "12/03/2025" /
# "2025-03-12". $preferent: el que va primer ("Cornella de Llobregat, ..."). PURA.
function _CtDataDeText([string]$text) {
    $p = _CtPla $text
    $m = [regex]::Match($p, 'cornella de llobregat,?\s+(?:a\s+)?(\d{1,2})\s+(?:de\s+|d.)\s*([a-z]+)\s+(?:de|del)\s+(\d{4})')
    if (-not $m.Success) { $m = [regex]::Match($p, '(?<!\d)(\d{1,2})\s+(?:de\s+|d.)\s*([a-z]+)\s+(?:de|del)\s+(\d{4})') }
    while ($m.Success) {
        $mes = [string]$m.Groups[2].Value
        if ($Script:CtMesos.ContainsKey($mes)) {
            $r = _CtDataIso ([int]$m.Groups[3].Value) $Script:CtMesos[$mes] ([int]$m.Groups[1].Value)
            if ($r -ne '') { return $r }
        }
        $m = $m.NextMatch()
    }
    $m = [regex]::Match($p, '(?<!\d)(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})(?!\d)')
    if ($m.Success) { $r = _CtDataIso ([int]$m.Groups[3].Value) ([int]$m.Groups[2].Value) ([int]$m.Groups[1].Value); if ($r -ne '') { return $r } }
    $m = [regex]::Match($p, '(?<!\d)(\d{4})-(\d{2})-(\d{2})')
    if ($m.Success) { return (_CtDataIso ([int]$m.Groups[1].Value) ([int]$m.Groups[2].Value) ([int]$m.Groups[3].Value)) }
    return ''
}

function _CtGiaDeText([string]$text) {
    $m = [regex]::Match((_CtPla $text), '\bgia\s*(?:n\.?|num\.?|:)?\s*(\d{1,6})\b')
    if ($m.Success) { return [string][int]$m.Groups[1].Value }
    return ''
}

# "2579/2024/12", "2025/1/2563"... el primer que hi hagi. PURA.
function _CtExpedientDeText([string]$text) {
    $m = [regex]::Match([string]$text, '(?<![\d/])(\d{1,4}/\d{1,4}/\d{1,6})(?![\d/])')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

# La serie 2579 es la de les QUEIXES: els interessats son els veins. PURA.
function _CtEsQueixa([string]$expedient, [string]$nomFitxer = '') {
    if (([string]$expedient) -match '^\s*2579\s*/') { return $true }
    return (([string]$nomFitxer) -match '(?<!\d)2579[/_\-]\d{4}')
}

# ----------------------------------------------------------------------------
# L'XML DE L'e-TRAM
# ----------------------------------------------------------------------------
# Tolerant amb el namespace (local-name()) i amb la forma de les <dada>
# (atribut clau/valor, fills <clau>/<valor> o el text).
function _CtXmlFill($el, [string]$nom) {
    if ($null -eq $el) { return '' }
    foreach ($c in @($el.ChildNodes)) {
        if ($c.NodeType -eq 'Element' -and $c.LocalName -ieq $nom) { return ([string]$c.InnerText).Trim() }
    }
    return ''
}

function _CtXmlPrimer($xml, [string]$nom) {
    $n = $xml.SelectSingleNode("//*[local-name()='" + $nom + "']")
    return $n
}

function _CtPersonaEtram($el) {
    if ($null -eq $el) { return $null }
    $nom = _CtXmlFill $el 'raoSocial'
    if ($nom -eq '') { $nom = (@((_CtXmlFill $el 'nom'), (_CtXmlFill $el 'cognom1'), (_CtXmlFill $el 'cognom2')) | Where-Object { $_ -ne '' }) -join ' ' }
    $p = _CtPersona $nom (_CtXmlFill $el 'numeroDocument') (_CtXmlFill $el 'correuElectronic') (_CtXmlFill $el 'telefon') (_CtXmlFill $el 'telefonMobil')
    if (_CtPersonaBuida $p) { return $null }
    return $p
}

function Read-EtramXml([string]$xmlText) {
    $x = New-Object System.Xml.XmlDocument
    $x.XmlResolver = $null
    try { $x.LoadXml($xmlText) } catch { return $null }
    $r = _CtDocNou 'etram'
    $r.interessat = _CtPersonaEtram (_CtXmlPrimer $x 'solicitant')
    $r.representant = _CtPersonaEtram (_CtXmlPrimer $x 'representant')
    $dc = _CtXmlPrimer $x 'dataCreacio'
    if ($null -ne $dc) { $r.data = _CtDataDeText ([string]$dc.InnerText) }
    $nt = _CtXmlPrimer $x 'nomTramit'
    if ($null -ne $nt) { $r.titol = ([string]$nt.InnerText).Trim() }
    # Els canals de notificacio: de qui presenta (el representant si n'hi ha).
    # Nomes omplen el que falti.
    $qui = if ($null -ne $r.representant) { $r.representant } else { $r.interessat }
    if ($null -ne $qui) {
        $sms = _CtXmlPrimer $x 'canalSms'
        if ($null -ne $sms) { foreach ($t in @(_CtTroba ([string]$sms.InnerText) $Script:CtReTel)) { _CtPosaTelefon $qui $t.Val } }
        $ce = _CtXmlPrimer $x 'canalCorreuElectronic'
        if ($null -ne $ce -and [string]$qui.email -eq '') {
            $em = @(_CtTroba ([string]$ce.InnerText) $Script:CtReEmail)
            if ($em.Count -gt 0) { $qui.email = _CtEmailNet $em[0].Val }
        }
    }
    $est = [ordered]@{ nom_comercial = ''; telefon = ''; email = '' }
    foreach ($d in @($x.SelectNodes("//*[local-name()='dada']"))) {
        $clau = ''
        if ($null -ne $d.Attributes -and $null -ne $d.Attributes['clau']) { $clau = [string]$d.Attributes['clau'].Value }
        if ($clau -eq '') { $clau = _CtXmlFill $d 'clau' }
        if ($clau -eq '') { $clau = _CtXmlFill $d 'nom' }
        $val = ''
        if ($null -ne $d.Attributes -and $null -ne $d.Attributes['valor']) { $val = [string]$d.Attributes['valor'].Value }
        if ($val -eq '') { $val = _CtXmlFill $d 'valor' }
        if ($val -eq '' -and @($d.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }).Count -eq 0) { $val = ([string]$d.InnerText).Trim() }
        $c = (_CtPla $clau) -replace '[^a-z]', ''
        if ($c -eq 'telefonestabliment') { $est.telefon = _CtTelefonNet $val; if ((_CtTipusTelefon $est.telefon) -eq '') { $est.telefon = '' } }
        elseif ($c -eq 'email') { $est.email = _CtEmailNet $val }
        elseif ($c -eq 'comercialname') { $est.nom_comercial = $val.Trim() }
    }
    if ($est.nom_comercial -ne '' -or $est.telefon -ne '' -or $est.email -ne '') { $r.establiment = $est }
    $r.gia = _CtGiaDeText $r.titol
    return $r
}

# Els XML antics (XML_TRAMIT, sicres3): Nombre_Interesado, Documento_
# Identificacion_Interesado, Telefono_Contacto_Interesado... i els mateixos amb
# _Representante.
function Read-XmlAntic([string]$xmlText) {
    $x = New-Object System.Xml.XmlDocument
    $x.XmlResolver = $null
    try { $x.LoadXml($xmlText) } catch { return $null }
    $r = _CtDocNou 'xmlantic'
    foreach ($qui in @(@{ Suf = 'interesado'; Camp = 'interessat' }, @{ Suf = 'representante'; Camp = 'representant' })) {
        $v = @{ nom = @(); nif = ''; email = ''; tel = @() }
        foreach ($n in @($x.SelectNodes("//*[contains(translate(local-name(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'_" + $qui.Suf + "')]"))) {
            $ln = ([string]$n.LocalName).ToLowerInvariant()
            $t = ([string]$n.InnerText).Trim()
            if ($t -eq '') { continue }
            if ($ln -match '^(nombre|razon|apellido)') { $v.nom += $t }
            elseif ($ln -match '^documento') { $v.nif = $t }
            elseif ($ln -match '^(correo|email|mail)') { $v.email = $t }
            elseif ($ln -match '^(telefono|movil)') { $v.tel += $t }
        }
        $p = _CtPersona ($v.nom -join ' ') $v.nif $v.email
        foreach ($t in $v.tel) { _CtPosaTelefon $p $t }
        if (-not (_CtPersonaBuida $p)) { $r[$qui.Camp] = $p }
    }
    foreach ($n in @($x.SelectNodes("//*[contains(translate(local-name(),'FECHA','fecha'),'fecha')]"))) {
        $d = _CtDataDeText ([string]$n.InnerText)
        if ($d -ne '') { $r.data = $d; break }
    }
    return $r
}

# ----------------------------------------------------------------------------
# LA INSTANCIA GENERICA / L'ESMENA (el text del PDF)
# ----------------------------------------------------------------------------
# Les seccions "DADES DE LA PERSONA O ENTITAT INTERESSADA" i "DADES DE LA
# PERSONA REPRESENTANT", i a dins etiqueta + valor (a la mateixa linia, darrere
# dels dos punts o no, o a la linia de sota). Com que no s'han pogut veure
# instancies de debo, es TOLERANT: si una etiqueta no hi es, les adreces, els
# NIF i els telefons de la seccio es busquen pel seu format. El diagnostic
# (DiagnosticPdf.ps1) es el que ho ha de confirmar amb les de veritat.
$Script:CtSeccioInteressat = @('dades de la persona o entitat interessada', 'dades de la persona interessada', 'dades de l.interessat', 'dades del sol.?licitant', 'dades de la persona sol.?licitant', 'persona o entitat interessada')
$Script:CtSeccioRepresentant = @('dades de la persona representant', 'dades del representant', 'persona representant')
# El que tanca una seccio (qualsevol altre titol conegut).
$Script:CtSeccioFi = @('dades ', 'expos', 'sol.?licit', 'documentacio', 'notificaci', 'mitja ', 'adreca a efectes', 'signatura', 'informacio basica', 'proteccio de dades', 'declaro', 'declaracio', 'observacions', 'dades de l.establiment', 'dades de l.activitat', 'dades de la sol.?licitud')

# Les etiquetes de cada camp, de la mes llarga a la mes curta (la primera que
# casa guanya: "telefon mobil" abans que "telefon").
$Script:CtEtiquetes = [ordered]@{
    nom     = @('nom i cognoms o rao social', 'nom o rao social', 'nom/rao social', 'rao social', 'nom i cognoms', 'denominacio', 'nom')
    cognom1 = @('primer cognom', 'cognom 1', 'cognom1')
    cognom2 = @('segon cognom', 'cognom 2', 'cognom2')
    nif     = @('dni/nif/nie', 'nif/nie/cif', 'nif/nie', 'dni/nie', 'dni/nif', 'numero de document', 'num. document', 'document identificatiu', 'nif', 'dni', 'nie', 'cif')
    mobil   = @('telefon mobil', 'mobil')
    telefon = @('telefon fix', 'telefon')
    email   = @('adreca de correu electronic', 'adreca electronica', 'correu electronic', 'e-mail', 'email', 'correu')
}

# Totes les etiquetes, per tallar un valor on comenca la seguent ("Telefon fix:
# 93... Mobil: 6..." a la mateixa linia).
function _CtTotesEtiquetes {
    $t = New-Object System.Collections.ArrayList
    foreach ($k in $Script:CtEtiquetes.Keys) { foreach ($e in $Script:CtEtiquetes[$k]) { [void]$t.Add($e) } }
    foreach ($e in @('adreca', 'domicili', 'codi postal', 'municipi', 'poblacio', 'provincia', 'pais', 'tipus de persona', 'tipus de document')) { [void]$t.Add($e) }
    return $t.ToArray()
}

# Les linies d'un text, sense buides.
function _CtLinies([string]$text) {
    return @(([string]$text -split "[`r`n`f]+") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
}

function _CtIndexSeccio($pla, [string[]]$titols, [int]$des = 0) {
    for ($i = $des; $i -lt $pla.Count; $i++) {
        foreach ($t in $titols) { if ($pla[$i] -match ('^\W*' + $t)) { return $i } }
    }
    return -1
}

function _CtFiSeccio($pla, [int]$ini) {
    for ($i = $ini + 1; $i -lt $pla.Count; $i++) {
        if (($i - $ini) -gt 30) { return $i }
        foreach ($t in $Script:CtSeccioFi) { if ($pla[$i] -match ('^\W*' + $t) -and $pla[$i] -notmatch ':\s*\S') { return $i } }
    }
    return $pla.Count
}

# El valor d'una etiqueta dins de les linies [ini, fi). '' si no hi es.
function _CtValorEtiqueta($lin, $pla, [int]$ini, [int]$fi, [string[]]$etiq, [string[]]$totes) {
    for ($i = $ini; $i -lt $fi; $i++) {
        foreach ($e in $etiq) {
            $m = [regex]::Match($pla[$i], '(?:^|\s|\()' + [regex]::Escape($e) + '\s*(?:\([^)]*\))?\s*[:\.]?\s*')
            if (-not $m.Success) { continue }
            $resta = $pla[$i].Substring($m.Index + $m.Length)
            $orig = $lin[$i].Substring($m.Index + $m.Length)
            # Talla on comenca la seguent etiqueta de la mateixa linia.
            $tall = $resta.Length
            foreach ($o in $totes) {
                if ($o -eq $e) { continue }
                $mo = [regex]::Match($resta, '(?:^|\s)' + [regex]::Escape($o) + '\s*[:]')
                if ($mo.Success -and $mo.Index -lt $tall) { $tall = $mo.Index }
            }
            $v = $orig.Substring(0, $tall).Trim().Trim(':', ' ')
            if ($v -ne '') { return $v }
            # Res darrere: el valor es a la linia de sota (si no es una altra etiqueta).
            if ($i + 1 -lt $fi) {
                $seg = $pla[$i + 1]
                $esEtiq = $false
                foreach ($o in $totes) { if ($seg -match ('^' + [regex]::Escape($o) + '\s*[:\.]?\s*$')) { $esEtiq = $true; break } }
                if (-not $esEtiq) { return $lin[$i + 1].Trim() }
            }
            return ''
        }
    }
    return ''
}

function _CtPersonaDeSeccio($lin, $pla, [int]$ini, [int]$fi) {
    if ($ini -lt 0) { return $null }
    $totes = @(_CtTotesEtiquetes)
    $v = @{}
    foreach ($k in $Script:CtEtiquetes.Keys) { $v[$k] = _CtValorEtiqueta $lin $pla ($ini + 1) $fi $Script:CtEtiquetes[$k] $totes }
    $nom = (@($v.nom, $v.cognom1, $v.cognom2) | Where-Object { $_ -ne '' }) -join ' '
    $text = ($lin[($ini + 1)..([Math]::Max($ini + 1, $fi - 1))] -join "`n")
    if ($ini + 1 -ge $fi) { $text = '' }
    $nif = $v.nif
    if ((_CtTipusNif $nif) -eq '') { $t = @(_CtTroba $text $Script:CtReNif | Where-Object { (_CtTipusNif $_.Val) -ne '' }); $nif = if ($t.Count -gt 0) { $t[0].Val } else { '' } }
    $email = _CtEmailNet $v.email
    if ($email -eq '') { $t = @(_CtTroba $text $Script:CtReEmail); if ($t.Count -gt 0) { $email = $t[0].Val } }
    $p = _CtPersona $nom $nif $email
    _CtPosaTelefon $p $v.mobil
    _CtPosaTelefon $p $v.telefon
    foreach ($t in @(_CtTroba $text $Script:CtReTel)) { _CtPosaTelefon $p $t.Val }
    if (_CtPersonaBuida $p) { return $null }
    return $p
}

function Read-InstanciaText([string]$text) {
    $lin = @(_CtLinies $text)
    if ($lin.Count -eq 0) { return $null }
    $pla = @($lin | ForEach-Object { _CtPla $_ })
    $r = _CtDocNou 'instancia'
    $iI = _CtIndexSeccio $pla $Script:CtSeccioInteressat
    $iR = _CtIndexSeccio $pla $Script:CtSeccioRepresentant
    if ($iI -ge 0) { $r.interessat = _CtPersonaDeSeccio $lin $pla $iI (_CtFiSeccio $pla $iI) }
    if ($iR -ge 0) { $r.representant = _CtPersonaDeSeccio $lin $pla $iR (_CtFiSeccio $pla $iR) }
    $r.data = _CtDataDeText $text
    # El titol: les primeres linies (hi sol haver l'expedient i el GIA).
    $cap = ($lin[0..([Math]::Min(3, $lin.Count - 1))] -join ' ')
    $r.titol = $lin[0]
    $r.gia = _CtGiaDeText $cap
    if ($r.gia -eq '') { $r.gia = _CtGiaDeText $text }
    $r.expedient = _CtExpedientDeText $cap
    if ($r.expedient -eq '') { $r.expedient = _CtExpedientDeText $text }
    return $r
}

# ----------------------------------------------------------------------------
# L'AUTORITZACIO DE REPRESENTACIO
# ----------------------------------------------------------------------------
#   "Jo, X, ... en nom i representacio de l'empresa Y ..., AUTORITZO a Z"
#   "D./Dna. X, como administrador/apoderado de Y, autoriza a Z"
# X es el REPRESENTANT LEGAL (signa per l'empresa) i Z la persona autoritzada
# (normalment el tecnic). Sense "en nom de" (Jo, X, AUTORITZO a Z), X signa en
# nom propi: es el titular, persona fisica.
$Script:CtReAutCa = "\bjo,?\s+(?<x>[^,]{3,90}?)\s*(?:,|\s+amb\s+(?:dni|nif|nie|d\.n\.i))"
$Script:CtReAutEs = "(?:^|[\s(])(?:d\.?\s*/\s*dna\.?|d\.\s*/\s*d\.?a\.?|dna\.|d\.|don|dona|sr\.|sra\.)\s*(?<x>[^,]{3,90}?)\s*(?:,|\s+con\s+(?:dni|nif|nie|d\.n\.i))"
$Script:CtReEnNom = "(?:en\s+nom\s+i\s+representacio\s+de|en\s+representacio\s+de|en\s+qualitat\s+d.(?:administrador|administradora|apoderat|apoderada|representant\s+legal|gerent)[a-z]*\s+(?:unic\s+|solidari\s+|mancomunat\s+)?de|como\s+(?:administrador|administradora|apoderado|apoderada|representante\s+legal|gerente|consejero\s+delegado|presidente|presidenta)(?:\s+(?:unico|unica|solidario|solidaria|mancomunado|mancomunada))?\s+de|en\s+nombre\s+y\s+representacion\s+de)\s+(?:(?:l.empresa|la\s+societat|l.entitat|la\s+mercantil|la\s+empresa|la\s+sociedad|la\s+entidad|la\s+companyia|la\s+compania)\s+)?(?<y>[^,]{2,110}?)\s*(?:,|\s+amb\s+(?:nif|cif)|\s+con\s+(?:nif|cif)|\s+amb\s+domicili|\s+con\s+domicilio|\.\s)"
$Script:CtReAutoritza = "\b(?:autoritzo|autoritza|autorizo|autoriza)\s+(?:a\s+|al\s+|a\s+la\s+)?(?:(?:senyor|senyora|sr\.|sra\.|en|na|d\.?\s*/\s*dna\.?|d\.|dna\.|don|dona)\s+)?(?<z>[^,]{3,110}?)\s*(?:,|\s+amb\s+(?:dni|nif|nie|d\.n\.i)|\s+con\s+(?:dni|nif|nie|d\.n\.i)|\s+per\s|\s+perque\s|\s+para\s|\s+a\s+fi\s|\.\s|$)"

function Read-AutoritzacioText([string]$text) {
    $orig = ([string]$text -replace '\s+', ' ').Trim()
    if ($orig -eq '') { return $null }
    $pla = _CtPla $orig
    $mz = [regex]::Match($pla, $Script:CtReAutoritza)
    if (-not $mz.Success) { return $null }
    $mx = [regex]::Match($pla.Substring(0, $mz.Index), $Script:CtReAutCa)
    if (-not $mx.Success) { $mx = [regex]::Match($pla.Substring(0, $mz.Index), $Script:CtReAutEs) }
    $my = [regex]::Match($pla.Substring(0, $mz.Index), $Script:CtReEnNom)
    # Sense qui signa ni en nom de qui, un "autoritza" qualsevol del text no fa
    # una autoritzacio.
    if (-not $mx.Success -and -not $my.Success) { return $null }
    $r = _CtDocNou 'autoritzacio'
    $nifs = @(_CtTroba $orig $Script:CtReNif | Where-Object { (_CtTipusNif $_.Val) -ne '' })
    $emails = @(_CtTroba $orig $Script:CtReEmail)
    $tels = @(_CtTroba $orig $Script:CtReTel)
    $primer = { param($llista, [int]$des, [int]$fins) foreach ($t in $llista) { if ($t.Pos -ge $des -and $t.Pos -lt $fins) { return $t.Val } } return '' }

    $iniY = if ($my.Success) { $my.Index } else { $mz.Index }
    $signant = $null
    if ($mx.Success) {
        $gx = $mx.Groups['x']
        $nomX = $orig.Substring($gx.Index, $gx.Length)
        $signant = _CtPersona $nomX (& $primer $nifs ($gx.Index) $iniY)
    }
    $empresa = $null
    if ($my.Success) {
        $gy = $my.Groups['y']
        $empresa = _CtPersona ($orig.Substring($gy.Index, $gy.Length)) (& $primer $nifs ($gy.Index) $mz.Index)
    }
    $gz = $mz.Groups['z']
    $aut = _CtPersona ($orig.Substring($gz.Index, $gz.Length)) (& $primer $nifs ($gz.Index) $orig.Length) (& $primer $emails ($gz.Index) $orig.Length)
    foreach ($t in $tels) { if ($t.Pos -ge $gz.Index) { _CtPosaTelefon $aut $t.Val } }
    $r.autoritzacio = [ordered]@{
        signant      = $signant
        empresa      = $empresa
        autoritzats  = @($aut)
        en_nom_propi = (-not $my.Success)
    }
    # El titular: l'empresa si signa en nom d'una; si no, qui signa.
    $r.interessat = if ($null -ne $empresa) { $empresa } else { $signant }
    $r.data = _CtDataDeText $orig
    $r.gia = _CtGiaDeText $orig
    $r.expedient = _CtExpedientDeText $orig
    return $r
}
