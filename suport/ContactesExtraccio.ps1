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

# Els noms de pila corrents: sols no diuen qui es ningu (l'usuari, 2a ronda: la
# representant legal del GIA 161 sortia com a "llista de tecnics coneguts"
# nomes per compartir "MARIA" amb una tecnica).
$Script:CtNomsPila = @{}
foreach ($w in @('maria', 'mari', 'jose', 'josep', 'joan', 'juan', 'jordi', 'antonio', 'antoni', 'francisco', 'francesc', 'manuel', 'david', 'javier',
    'xavier', 'carlos', 'carles', 'ana', 'anna', 'laura', 'marta', 'carmen', 'pere', 'pedro', 'jesus', 'luis', 'lluis', 'miguel', 'miquel', 'angel',
    'rosa', 'isabel', 'montserrat', 'pilar', 'dolores', 'mar', 'mercedes', 'cristina', 'elena', 'sergio', 'sergi', 'daniel', 'alejandro', 'alex',
    'pablo', 'pau', 'marc', 'albert', 'alberto', 'rafael', 'raul', 'fernando', 'ramon', 'enric', 'enrique', 'jaume', 'jaime', 'josefa', 'teresa',
    'nuria', 'silvia', 'sara', 'paula', 'lucia', 'eva', 'raquel', 'mireia', 'sonia', 'alba', 'julia', 'jorge', 'oscar', 'ivan', 'ruben', 'adrian',
    'victor', 'roberto', 'robert', 'andres', 'andreu', 'joaquin', 'joaquim', 'ignacio', 'ignasi', 'eduardo', 'eduard', 'mohamed', 'muhammad',
    'ahmed', 'ali', 'li', 'wang', 'carmen', 'concepcion', 'josefina', 'antonia', 'francisca', 'manuela', 'angeles', 'lourdes', 'beatriz', 'patricia',
    'monica', 'susana', 'yolanda', 'noelia', 'gemma', 'judit', 'meritxell', 'nerea', 'irene', 'andrea', 'clara', 'jose maria', 'juana', 'luisa')) { $Script:CtNomsPila[$w] = $true }

# Dos noms son de la mateixa persona? Les paraules del curt son totes al llarg,
# n'hi ha com a minim dues (o es el mateix nom sencer) i com a minim UNA no es
# un nom de pila corrent (un cognom; l'usuari, 2a ronda): "Maria Exemple" i
# "MARIA EXEMPLE PROVA" si; "Maria" sola, o "Maria del Mar" amb "Maria del Mar
# Garcia", no. PURA.
function _CtMateixNom([string]$a, [string]$b) {
    $pa = @(_CtNomParaules $a); $pb = @(_CtNomParaules $b)
    if ($pa.Count -eq 0 -or $pb.Count -eq 0) { return $false }
    $curt = $pa; $llarg = $pb
    if ($pa.Count -gt $pb.Count) { $curt = $pb; $llarg = $pa }
    if ($curt.Count -lt 2 -and ($pa -join ' ') -ne ($pb -join ' ')) { return $false }
    foreach ($w in $curt) { if ($llarg -notcontains $w) { return $false } }
    foreach ($w in $curt) { if (-not $Script:CtNomsPila.ContainsKey($w)) { return $true } }
    return $false
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
    # Una etiqueta no es mai un nom ("Nombre y apellidos o razon social"), ni un
    # tros de frase ("N.I. y N.I.F. numero ... INTERVIENE en su calidad de").
    if ((_CtNomEsEtiqueta $p.nom) -or (_CtNomBrossa $p.nom $p.nif)) { $p.nom = '' }
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

# LES ETIQUETES DELS FORMULARIS (en catala i en castella). La instancia porta
# cada etiqueta en les dues llengues, i el text del PDF les talla en linies
# diferents ("Nom i cognoms o rao social Nombre y apellidos o razon social").
# Una linia que, sense la puntuacio, nomes te paraules d'aquestes es una
# etiqueta, no un valor. "Forts": les que no poden ser el principi d'un nom.
$Script:CtParaulesEtiquetaForts = @('nom', 'cognoms', 'cognom', 'rao', 'social', 'nombre', 'apellidos', 'apellido', 'razon', 'dni', 'nif', 'nie', 'cif', 'pas',
    'passaport', 'pasaporte', 'adreca', 'postal', 'direccion', 'carrer', 'placa', 'av', 'avda', 'num', 'esc', 'pis', 'porta', 'calle', 'plaza', 'piso', 'puerta',
    'municipi', 'municipio', 'cp', 'codi', 'codigo', 'provincia', 'poblacio', 'poblacion', 'pais', 'telefon', 'fix', 'telefono', 'fijo', 'mobil', 'movil',
    'correu', 'electronic', 'correo', 'electronico', 'mail', 'email', 'representat', 'representado', 'segon', 'segundo', 'document', 'documento', 'numero',
    'identificatiu', 'identificativo', 'denominacio', 'denominacion', 'domicili', 'domicilio', 'fax', 'tipus', 'tipo', 'nombres',
    'avinguda', 'avenida', 'escala', 'escalera', 'bloc', 'bloque', 'planta', 'via', 'km', 'lletra', 'letra', 'portal', 'localitat', 'localidad',
    'comarca', 'notificacio', 'notificacion', 'mitja', 'medio', 'canal', 'electronica', 'paper', 'papel', 'preferent', 'preferente',
    'urbanitzacio', 'urbanizacion', 'apartat', 'apartado', 'nacionalitat', 'nacionalidad', 'naixement', 'nacimiento', 'postals', 'telefons', 'telefonos')
$Script:CtParaulesEtiqueta = @{}
foreach ($w in @($Script:CtParaulesEtiquetaForts + @('i', 'y', 'o', 'e', 'per', 'por', 'de', 'del', 'la', 'el', 'primer', 'a', 'en', 'persona', 'fisica', 'juridica', 'efectes', 'efectos', 'data', 'fecha', 'correus', 'correos'))) { $Script:CtParaulesEtiqueta[$w] = $true }
$Script:CtParaulesEtiquetaFort = @{}
foreach ($w in $Script:CtParaulesEtiquetaForts) { $Script:CtParaulesEtiquetaFort[$w] = $true }

# Les paraules (lletres i xifres) del text pla, amb la posicio. PURA.
function _CtParaulesAmbPos([string]$pla) {
    return @([regex]::Matches([string]$pla, '[a-z0-9]+') | ForEach-Object { @{ W = $_.Value; Pos = $_.Index } })
}

# Una linia que nomes te paraules d'etiqueta (i com a minim una forta). O
# gairebe nomes (l'usuari, 2a ronda: "Adreca postal (carrer, placa, ...)
# Direccion postal (...)" sortia de nom, i deia que "presentava per a 78
# titulars"): amb tres paraules fortes o mes, n'hi ha prou amb el 80%. PURA.
function _CtEsLiniaEtiqueta([string]$s) {
    $ws = @(_CtParaulesAmbPos (_CtPla $s))
    if ($ws.Count -eq 0) { return $false }
    $nFort = 0; $nEtiq = 0
    foreach ($w in $ws) {
        if (-not $Script:CtParaulesEtiqueta.ContainsKey($w.W)) { continue }
        $nEtiq++
        if ($Script:CtParaulesEtiquetaFort.ContainsKey($w.W)) { $nFort++ }
    }
    if ($nFort -eq 0) { return $false }
    if ($nEtiq -eq $ws.Count) { return $true }
    return ($nFort -ge 3 -and ($nEtiq / $ws.Count) -ge 0.8)
}

# El valor d'una linia que comenca amb etiquetes ("Nom i cognoms o rao social
# ... DNI / NIF / NIE / Pas. EXEMPLE SL B00000000", la instancia telematica
# antiga): el que hi ha darrere. Si no en comenca amb cap etiqueta forta (un nom
# com "DE LA FONT"), la linia sencera; si es tota etiqueta, ''. PURA.
function _CtSenseEtiquetes([string]$s) {
    $pla = _CtPla $s
    $fort = $false
    foreach ($w in @(_CtParaulesAmbPos $pla)) {
        if ($Script:CtParaulesEtiqueta.ContainsKey($w.W)) { if ($Script:CtParaulesEtiquetaFort.ContainsKey($w.W)) { $fort = $true }; continue }
        if (-not $fort) { return $s.Trim() }
        return $s.Substring($w.Pos).Trim()
    }
    if ($fort) { return '' }
    return $s.Trim()
}

# Un "nom" que es una etiqueta (l'usuari: 132 titulars amb el nom "Nombre y
# apellidos o razon social DNI / NIF / NIE / Pas"). PURA.
function _CtNomEsEtiqueta([string]$nom) {
    if ([string]::IsNullOrWhiteSpace($nom)) { return $false }
    $p = _CtPla $nom
    if ($p -match 'razon social|rao social|\brepresentado\b|\brepresentat\b|dni\s*/\s*nif|nombre y apellidos|nom i cognoms|adreca postal|direccion postal|codi postal|codigo postal|mitja de notificacio|medio de notificacion|tipus de via|tipo de via|correu electronic|correo electronico|telefon mobil|telefono movil') { return $true }
    return (_CtEsLiniaEtiqueta $nom)
}

# Un "nom" que es un tros de frase (l'usuari, 2a ronda: "N.I. 478981881<",
# "N.I. y N. I. F. numero ... INTERVIENE en su calidad de APODERADO"): comenca
# per N.I, porta "interviene", "calidad", "qualitat", "numero", o xifres (les
# xifres, nomes en una persona: "BAR 2000 SL" si que es un nom). PURA.
function _CtNomBrossa([string]$nom, [string]$nif = '') {
    if ([string]::IsNullOrWhiteSpace($nom)) { return $false }
    $p = _CtPla $nom
    if ($p -match '^\W*n\s*\.\s*i\b' -or $p -match '\b(interviene|interve|calidad|qualitat|numero)\b') { return $true }
    if ($p -match '\d') {
        $entitat = ((_CtTipusNif $nif) -eq 'juridica') -or ($p -match '\b(s\.?\s?l\.?u?|s\.?\s?a\.?u?|scp|sccl|c\.?\s?b|slne|sll)\W*$')
        if (-not $entitat) { return $true }
    }
    return $false
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
    # Els tractaments del davant ("Sr.", "D./Dna.", "Dona", "En", "/D>>"), amb la
    # puntuacio del principi.
    for ($k = 0; $k -lt 4; $k++) {
        $m = [regex]::Match($pla, '^[^\p{L}\d]+')
        if ($m.Success) { $t = $t.Substring($m.Length); $pla = $pla.Substring($m.Length) }
        $m = [regex]::Match($pla, '^(?:sr|sra|senyor|senyora|en|na|d|dna|don|dona)(?:\s*[.\u00BB"''/]+\s*|\s+)(?=\S)')
        if (-not $m.Success) { break }
        $t = $t.Substring($m.Length); $pla = $pla.Substring($m.Length)
    }
    # Mai un DNI dins d'un nom ("N.I. 00000000T en representacio de..."): el
    # nom s'acaba on comenca el document, i el que hi ha davant ("amb N.I.")
    # tampoc es del nom.
    $m = [regex]::Match($pla, $Script:CtReNif)
    if ($m.Success -and (_CtTipusNif $m.Value) -ne '') { $t = $t.Substring(0, $m.Index); $pla = $pla.Substring(0, $m.Index) }
    $m = [regex]::Match($pla, '(?:^|[\s,])(?:amb|con)?\s*(?:d\.?\s?n\.?\s?i\.?|n\.?\s?i\.?\s?f\.?|n\.?\s?i\.?\s?e\.?|n\.\s?i\.?)\s*[:.]?\s*$')
    if ($m.Success) { $t = $t.Substring(0, $m.Index) }
    return $t.Trim().Trim(',', ';', ':', '.', '-', ' ')
}

# El nom del que signa (o del que s'autoritza) tret d'una frase: s'atura on
# comenca el que ja no es el nom (l'usuari: sortia "<Nom> actuando como
# representante legal de la empresa..."). PURA (sobre el text pla).
$Script:CtReFiNom = '(?:,|;|\s+(?:amb|con)\s+(?:d\.?\s?n\.?(?:\s?\(\s?i\s?\)|\s?i)?(?![a-z])|n\.?\s?i\.?\s?f|n\.?\s?i\.?\s?e|n\.\s?i\.|dni|nif|nie)|\s+actuant|\s+actuando|\s+en\s+representaci|\s+en\s+nom(?:bre)?\b|\s+com\s+a\s|\s+como\s|\s+en\s+qualitat|\s+en\s+calidad|\s+major\s+d.edat|\s+mayor\s+de\s+edad)'
function _CtTallaNom([string]$nom) {
    $t = [string]$nom
    $m = [regex]::Match((_CtPla $t), $Script:CtReFiNom)
    if ($m.Success) { $t = $t.Substring(0, $m.Index) }
    return (_CtNetejaNom $t)
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
        signatures = @()
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

# "2024/3165/33", "2579/2024/12"... el primer que hi hagi. NOMES AAAA/NNNN/N:
# abans s'hi colaven les dates ("25/03/2025", l'usuari: 262 documents), i una
# data AAAA/MM/DD tampoc no ho es. PURA.
function _CtExpedientDeText([string]$text) {
    $m = [regex]::Match([string]$text, '(?<![\d/])(\d{4})/(\d{1,5})/(\d{1,6})(?![\d/])')
    while ($m.Success) {
        $b = $m.Groups[2].Value; $c = $m.Groups[3].Value
        $esData = ($b.Length -eq 2 -and $c.Length -eq 2 -and (_CtDataIso ([int]$m.Groups[1].Value) ([int]$b) ([int]$c)) -ne '')
        if (-not $esData) { return $m.Value }
        $m = $m.NextMatch()
    }
    return ''
}

# Les carpetes d'una clau (ruta relativa), de la de mes a prop a la d'activitat:
# @{ Nom; Rel } (Rel: la ruta relativa de la carpeta). PURA.
function _CtCarpetesDe([string]$clau) {
    $p = @(([string]$clau -replace '/', '\').Trim('\').Split('\') | Where-Object { $_ -ne '' })
    $out = New-Object System.Collections.ArrayList
    for ($i = $p.Count - 2; $i -ge 0; $i--) { [void]$out.Add(@{ Nom = [string]$p[$i]; Rel = ($p[0..$i] -join '\') }) }
    return $out.ToArray()
}

# L'expedient del nom d'una carpeta ("2025-16-3165 GIA 595 - ..."), o ''. PURA.
function _CtExpedientDeCarpeta([string]$nom) {
    $m = [regex]::Match([string]$nom, '(?<![\d\-])(\d{4})-(\d{1,5})-(\d{1,6})(?![\d\-])')
    if (-not $m.Success) { return '' }
    return ([string][int]$m.Groups[1].Value + '-' + [string][int]$m.Groups[2].Value + '-' + [string][int]$m.Groups[3].Value)
}

# Una carpeta diu que es d'una queixa? 'si' (serie 2579 al nom, "AAAA-N-2579",
# o les paraules), 'no' (es d'un altre expedient) o '' (no ho diu). PURA.
function _CtCarpetaQueixa([string]$nom) {
    if ((_CtPla $nom) -match '\b(queixa|queixes|quexia|queja|quejas|denuncia)') { return 'si' }
    $e = _CtExpedientDeCarpeta $nom
    if ($e -eq '') { return '' }
    $p = $e.Split('-')
    if ($p[2] -eq '2579' -or $p[0] -eq '2579') { return 'si' }
    return 'no'
}

# La serie 2579 es la de les QUEIXES: els interessats son els veins. Ho diu
# l'expedient del document ("2579/2025/12", "2025/2579/12"), el nom del fitxer o
# LA CARPETA D'EXPEDIENT MES PROPERA (l'usuari: el document era a "...2579 GIA
# 36 - QUEIXA OLORS\"). La carpeta del titular DINS de la d'una queixa no es
# queixa. $ruta: la ruta relativa (o nomes el nom del fitxer). PURA.
function _CtEsQueixa([string]$expedient, [string]$ruta = '') {
    $propera = ''; $sobre = $false
    foreach ($c in @(_CtCarpetesDe $ruta)) {
        $q = _CtCarpetaQueixa $c.Nom
        if ($q -eq '') { continue }
        if ($propera -eq '') { $propera = $q } elseif ($q -eq 'si') { $sobre = $true }
    }
    if ($propera -eq 'si') { return $true }
    if ($propera -eq 'no' -and $sobre) { return $false }
    if (([string]$expedient) -match '^\s*2579\s*/' -or ([string]$expedient) -match '^\s*\d{4}\s*/\s*2579\s*/') { return $true }
    $nom = @(([string]$ruta -replace '/', '\').Split('\'))[-1]
    return (([string]$nom) -match '(?<!\d)2579[/_\-]\d{4}')
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

# Un raoSocial buit de contingut: l'e-TRAM antic hi posa "rao" (i "0" al
# telefon) per a una persona fisica (l'usuari: el representant legal "rao"). PURA.
function _CtRaoBuida([string]$s) {
    $p = (_CtPla $s).Trim()
    if (@('rao', 'raon', 'razon', 'nom', 'nombre') -contains $p) { return $true }
    return (([regex]::Matches($p, '[a-z]')).Count -lt 3)
}

function _CtPersonaEtram($el) {
    if ($null -eq $el) { return $null }
    $nom = _CtXmlFill $el 'raoSocial'
    if (_CtRaoBuida $nom) { $nom = '' }
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
$Script:CtSeccioInteressat = @('dades de la persona o entitat interessada', 'dades de la persona interessada', 'dades de l.interessat', 'dades del sol.?licitant', 'dades de la persona sol.?licitant', 'persona o entitat interessada', 'datos de la persona o entidad interesada', 'datos de la persona interesada', 'datos del solicitante')
$Script:CtSeccioRepresentant = @('dades de la persona representant', 'dades del representant', 'persona representant', 'datos de la persona representante', 'datos del representante')
# El que tanca una seccio (qualsevol altre titol conegut).
$Script:CtSeccioFi = @('dades ', 'datos ', 'expos', 'sol.?licit', 'documentacio', 'notificaci', 'mitja ', 'adreca a efectes', 'signatura', 'informacio basica', 'proteccio de dades', 'declaro', 'declaracio', 'observacions', 'dades de l.establiment', 'dades de l.activitat', 'dades de la sol.?licitud')

# Les etiquetes de cada camp, de la mes llarga a la mes curta (la primera que
# casa guanya: "telefon mobil" abans que "telefon").
$Script:CtEtiquetes = [ordered]@{
    nom     = @('nombre y apellidos o razon social', 'nom i cognoms o rao social', 'representado por', 'representat per', 'nombre o razon social', 'nom o rao social', 'nom/rao social', 'rao social', 'razon social', 'nombre y apellidos', 'nom i cognoms', 'denominacion', 'denominacio', 'nombre', 'nom')
    cognom1 = @('primer cognom', 'primer apellido', 'cognom 1', 'cognom1')
    cognom2 = @('segon cognom', 'segundo apellido', 'cognom 2', 'cognom2')
    nif     = @('dni / nif / nie / pas', 'dni / nif / nie', 'dni/nif/nie/pas', 'dni/nif/nie', 'nif/nie/cif', 'dni / nif', 'nif / nie', 'nif/nie', 'dni/nie', 'dni/nif', 'numero de documento', 'numero de document', 'num. document', 'document identificatiu', 'nif', 'dni', 'nie', 'cif')
    mobil   = @('telefono movil', 'telefon mobil', 'movil', 'mobil')
    telefon = @('telefono fijo', 'telefon fix', 'telefono', 'telefon')
    email   = @('adreca de correu electronic', 'adreca electronica', 'correo electronico', 'correu electronic', 'e-mail', 'email', 'correo', 'correu')
}

# Totes les etiquetes, per tallar un valor on comenca la seguent ("Telefon fix:
# 93... Mobil: 6..." a la mateixa linia).
function _CtTotesEtiquetes {
    $t = New-Object System.Collections.ArrayList
    foreach ($k in $Script:CtEtiquetes.Keys) { foreach ($e in $Script:CtEtiquetes[$k]) { [void]$t.Add($e) } }
    foreach ($e in @('adreca postal', 'direccion postal', 'adreca', 'direccion', 'domicili', 'domicilio', 'codi postal', 'codigo postal', 'municipi', 'municipio', 'poblacio', 'poblacion', 'provincia', 'pais', 'tipus de persona', 'tipus de document', 'cp')) { [void]$t.Add($e) }
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
            $m = [regex]::Match($pla[$i], '(?:^|\s|\()' + [regex]::Escape($e) + '(?![a-z0-9])\s*(?:\([^)]*\))?\s*[:\.]?\s*')
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
            # Darrere hi ha la mateixa etiqueta en castella: no es cap valor.
            if ($v -ne '' -and (_CtEsLiniaEtiqueta $v)) { $v = '' }
            if ($v -ne '') { return $v }
            # Res darrere: el valor es a la linia de sota (si no es una altra etiqueta).
            if ($i + 1 -lt $fi) {
                $seg = $pla[$i + 1]
                $esEtiq = (_CtEsLiniaEtiqueta $lin[$i + 1])
                foreach ($o in $totes) { if ($esEtiq) { break }; if ($seg -match ('^' + [regex]::Escape($o) + '\s*[:\.]?\s*$')) { $esEtiq = $true } }
                if (-not $esEtiq) { return $lin[$i + 1].Trim() }
            }
            return ''
        }
    }
    return ''
}

# El primer valor d'una seccio que pot ser un nom: el de la primera linia que
# no es etiqueta (sense les etiquetes del davant), fins al NIF, el telefon o el
# correu que porti enganxats. Ni una linia nomes de xifres (el CP) ni de
# signes. PURA.
function _CtPrimerValorNom($lin, [int]$ini, [int]$fi) {
    for ($i = $ini; $i -lt $fi; $i++) {
        $v = _CtSenseEtiquetes $lin[$i]
        if ($v -eq '' -or $v -notmatch '\p{L}{2}') { continue }
        $tall = $v.Length
        foreach ($re in @($Script:CtReNif, $Script:CtReEmail, $Script:CtReTel)) {
            $m = [regex]::Match($v, $re)
            if ($m.Success -and $m.Index -lt $tall) { $tall = $m.Index }
        }
        $v = $v.Substring(0, $tall).Trim().Trim(',', ';', ':', '-', ' ')
        if ($v -notmatch '\p{L}{2}' -or (_CtNomEsEtiqueta $v)) { continue }
        return $v
    }
    return ''
}

function _CtPersonaDeSeccio($lin, $pla, [int]$ini, [int]$fi) {
    if ($ini -lt 0) { return $null }
    $totes = @(_CtTotesEtiquetes)
    $v = @{}
    foreach ($k in $Script:CtEtiquetes.Keys) { $v[$k] = _CtValorEtiqueta $lin $pla ($ini + 1) $fi $Script:CtEtiquetes[$k] $totes }
    $nom = (@($v.nom, $v.cognom1, $v.cognom2) | Where-Object { $_ -ne '' }) -join ' '
    if (_CtNomEsEtiqueta $nom) { $nom = '' }
    # La instancia d'ara: totes les etiquetes primer i els valors a sota. El
    # nom es la primera linia que no es etiqueta (i el NIF, la seguent).
    if ($nom -eq '') { $nom = _CtPrimerValorNom $lin ($ini + 1) $fi }
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
    $r.signatures = @(_CtSignaturesRep $text)
    return $r
}

# ----------------------------------------------------------------------------
# LA SIGNATURA ELECTRONICA DE REPRESENTANT: "(R:<CIF>)"
# ----------------------------------------------------------------------------
# L'usuari (octubre 2026): "es la font mes segura". Els certificats de
# representant d'entitat signen aixi:
#   Firmado digitalmente por 00000000T MARIA EXEMPLE (R:B00000000)
#   MARIA EXEMPLE (R: B00000000) Firmado digitalmente por ...
# La 2a ronda: la forma mes corrent porta un ESPAI despres dels dos punts, i al
# caixeti el nom sol surt SENSE el DNI (abans es demanava el DNI just davant
# del nom i en trobava 81 de 183). El text del PDF ho pot partir en linies i
# tallar el CIF ("(R:B000000" / "00)"): s'ajunten abans de buscar. Per cada
# "(R:CIF)" es mira el tros de davant (fins a "Firmado digitalmente por", un
# altre ")" o 150 caracters): el DNI, si n'hi ha, i el nom (les paraules en
# MAJUSCULES de just davant). Qui signa amb (R:<CIF>) es el representant
# legal de l'entitat d'aquell CIF. Torna @{ nom; nif; cif } per cada signant
# (sense repetir: la mateixa persona amb DNI i sense, una sola vegada). PURA.
$Script:CtReSigR = '\(\s*R\s*:\s*([A-Za-z](?:\s*[0-9A-Za-z]){7,8})\s*\)'
$Script:CtReSigFrontera = '(?i)(?:firmad[oa]\s+digitalmente\s+por|signat\s+digitalment\s+per|firmat\s+digitalment\s+per|firmado\s+por|signat\s+per|digitally\s+signed\s+by|por\s*:|per\s*:|\)|:)'
function _CtSignaturesRep([string]$text) {
    $t = ([string]$text) -replace '\s+', ' '
    if ($t -notmatch '\(\s*R\s*:') { return @() }
    $t = [regex]::Replace($t, $Script:CtReSigR, { param($m) '(R:' + ($m.Groups[1].Value -replace ' ', '').ToUpperInvariant() + ')' })
    $perClau = [ordered]@{}
    foreach ($m in [regex]::Matches($t, '\(R:([A-Z]\d{7}[0-9A-J])\)')) {
        $cif = _CtNifNet $m.Groups[1].Value
        if ((_CtTipusNif $cif) -ne 'juridica') { continue }
        $ini = [Math]::Max(0, $m.Index - 150)
        $tros = $t.Substring($ini, $m.Index - $ini)
        $fr = @([regex]::Matches($tros, $Script:CtReSigFrontera))
        if ($fr.Count -gt 0) { $u = $fr[$fr.Count - 1]; $tros = $tros.Substring($u.Index + $u.Length) }
        $nif = ''
        $mn = [regex]::Match($tros, '(?<![A-Za-z0-9])(\d{8}\s?-?\s?[A-Z]|[XYZ]\s?\d{7}\s?[A-Z])(?![A-Za-z0-9])')
        if ($mn.Success -and (_CtTipusNif $mn.Groups[1].Value) -eq 'fisica') { $nif = _CtNifNet $mn.Groups[1].Value; $tros = $tros.Remove($mn.Index, $mn.Length) }
        # El nom: les paraules en majuscules del final del tros.
        $ws = @(($tros -replace '\b(DNI|NIF|NIE)\b', ' ' -replace '[-,;]', ' ').Trim() -split '\s+' | Where-Object { $_ -ne '' })
        $nomW = New-Object System.Collections.ArrayList
        for ($k = $ws.Count - 1; $k -ge 0 -and $nomW.Count -lt 6; $k--) {
            if ($ws[$k] -cmatch "^[\p{Lu}][\p{Lu}\p{M}'\u00B7.]*$") { [void]$nomW.Insert(0, $ws[$k]) } else { break }
        }
        $nom = if ($nomW.Count -ge 2) { _CtNetejaNom ($nomW -join ' ') } else { '' }
        if ($nom -eq '' -and $nif -eq '') { continue }
        $clau = $(if ($nom -ne '') { _CtNomNorm $nom } else { $nif }) + '|' + $cif
        if ($perClau.Contains($clau)) { if ([string]$perClau[$clau].nif -eq '') { $perClau[$clau].nif = $nif }; continue }
        # El mateix DNI amb el nom escrit d'una altra manera, o sense nom.
        $repetit = $false
        foreach ($o in $perClau.Values) {
            if ([string]$o.cif -ne $cif) { continue }
            if (($nif -ne '' -and [string]$o.nif -eq $nif) -or ($nom -ne '' -and (_CtMateixNom $nom ([string]$o.nom)))) {
                if ([string]$o.nif -eq '') { $o.nif = $nif }
                if ([string]$o.nom -eq '') { $o.nom = $nom }
                $repetit = $true; break
            }
        }
        if (-not $repetit) { $perClau[$clau] = [ordered]@{ nom = $nom; nif = $nif; cif = $cif } }
    }
    return @($perClau.Values)
}

# ----------------------------------------------------------------------------
# L'AUTORITZACIO DE REPRESENTACIO
# ----------------------------------------------------------------------------
#   "Jo, X, ... en nom i representacio de l'empresa Y ..., AUTORITZO a Z"
#   "D./Dna. X, como administrador/apoderado de Y, autoriza a Z"
#   "X, amb DNI ..., en nom i representacio de Y, DESIGNA com a tecnica ... a Z"
# X es el REPRESENTANT LEGAL (signa per l'empresa) i Z la persona autoritzada
# (normalment el tecnic). Sense "en nom de" (Jo, X, AUTORITZO a Z), X signa en
# nom propi: es el titular, persona fisica.
$Script:CtReAutCa = "\bjo,?\s+(?<x>[^,]{3,90}?)\s*" + $Script:CtReFiNom
$Script:CtReAutEs = "(?:^|[\s(])(?:d\.?\s*/\s*dna\.?|d\.\s*/\s*d\.?a\.?|dna\.|d\.|don|dona|sr\.|sra\.)\s*(?<x>[^,]{3,90}?)\s*" + $Script:CtReFiNom
# Sense "Jo," ni tractament: el nom que va davant de "amb DNI" (l'ultim abans
# de "en nom de"). Del que hi ha davant (el titol del document) se'n treu.
$Script:CtReAutNomDni = "(?<x>[^,.;:()]{3,120}?)\s*,?\s+(?:amb|con)\s+(?:d\.?\s?n\.?(?:\s?\(\s?i\s?\)|\s?i)?(?![a-z])|n\.?\s?i\.?\s?f|n\.?\s?i\.?\s?e|n\.\s?i\.|dni|nif|nie)"
$Script:CtReTitolDoc = "^.*\b(?:autoritzacio|autorizacion|representacio|representacion|designacio|designacion|nomenament|nombramiento|sotasignat|sotasignada|subscriu|suscribe|firmante|declaro|declara)\b\s*"
$Script:CtReEnNom = "(?:en\s+nom\s+i\s+representacio\s+de|en\s+representacio\s+de|en\s+qualitat\s+d.(?:administrador|administradora|apoderat|apoderada|representant\s+legal|gerent)[a-z]*\s+(?:unic\s+|solidari\s+|mancomunat\s+)?de|como\s+(?:administrador|administradora|apoderado|apoderada|representante\s+legal|gerente|consejero\s+delegado|presidente|presidenta)(?:\s+(?:unico|unica|solidario|solidaria|mancomunado|mancomunada))?\s+de|en\s+nombre\s+y\s+representacion\s+de)\s+(?:(?:l.empresa|la\s+societat|l.entitat|la\s+mercantil|la\s+empresa|la\s+sociedad|la\s+entidad|la\s+companyia|la\s+compania)\s+)?(?<y>[^,]{2,110}?)\s*(?:,|\s+amb\s+(?:nif|cif)|\s+con\s+(?:nif|cif)|\s+amb\s+domicili|\s+con\s+domicilio|\.\s)"
# El verb: "autoritzo", i tambe (l'usuari) "designa", "nomena", "atorga la seva
# representacio a", "otorga su representacion a", "apodera", "faculta". Amb "com
# a tecnica responsable ... a Z", Z es el que va darrere de la "a".
$Script:CtReVerbAut = "(?:autoritzo|autoritza|autorizo|autoriza|designo|designa|nomeno|nomena|nombra|apodero|apodera|faculto|faculta|(?:atorgo|atorga)\s+(?:la\s+)?(?:seva\s+|meva\s+)?representacio|(?:otorgo|otorga)\s+(?:(?:su|mi)\s+)?representacion)"
$Script:CtReAutoritza = "\b" + $Script:CtReVerbAut + "\s+(?:(?:com\s+a|como)\s+[^,]{0,80}?\s+(?:a|al|a\s+la)\s+|(?:a\s+|al\s+|a\s+la\s+|a\s+favor\s+de\s+|en\s+favor\s+de\s+))?(?:(?:senyor|senyora|sr\.|sra\.|en|na|d\.?\s*/\s*dna\.?|d\.|dna\.|don|dona)\s+)?(?<z>[^,]{3,250}?)\s*(?:,|\s+amb\s+(?:dni|nif|nie|d\.n\.i)|\s+con\s+(?:dni|nif|nie|d\.n\.i)|\s+per\s|\s+perque\s|\s+para\s|\s+a\s+fi\s|\s+com\s+a\s|\s+como\s|\.\s|$)"

# El text amb els espais junts, i on comencava cada linia (per saber on comenca
# el nom del que signa quan no hi ha "Jo," ni tractament: al principi de la
# seva linia). PURA.
function _CtAjuntaLinies([string]$text) {
    $sb = New-Object System.Text.StringBuilder
    $salts = New-Object System.Collections.ArrayList
    foreach ($tros in [regex]::Split(([string]$text).Trim(), '(\s+)')) {
        if ($tros -eq '') { continue }
        if ($tros -match '^\s+$') {
            [void]$sb.Append(' ')
            if ($tros -match '[\r\n]') { [void]$salts.Add($sb.Length) }
        } else { [void]$sb.Append($tros) }
    }
    return @{ Text = $sb.ToString(); Salts = $salts.ToArray() }
}

# El nom del que signa sense "Jo," (2a ronda, GIA 5: "X amb DN(I): ..., en nom
# i representacio de Y, ... designa ..."): del principi de la seva linia, sense
# el titol del document; i si davant hi ha text en minuscules ("La sotasignada
# MARIA EXEMPLE"), nomes les paraules en MAJUSCULES del final. PURA.
function _CtNomSignantSenseJo([string]$orig, [string]$pla, [int]$ini, [int]$fi, $salts) {
    foreach ($sl in @($salts)) { if ([int]$sl -gt $ini -and [int]$sl -lt $fi) { $ini = [int]$sl } }
    $nom = $orig.Substring($ini, $fi - $ini)
    $mt = [regex]::Match($pla.Substring($ini, $fi - $ini), $Script:CtReTitolDoc)
    if ($mt.Success) { $nom = $nom.Substring($mt.Length) }
    $ws = @($nom.Trim() -split '\s+' | Where-Object { $_ -ne '' })
    $maj = New-Object System.Collections.ArrayList
    for ($k = $ws.Count - 1; $k -ge 0; $k--) { if ($ws[$k] -cmatch "^[\p{Lu}][\p{Lu}\p{M}'\u00B7.\-]*$") { [void]$maj.Insert(0, $ws[$k]) } else { break } }
    if ($maj.Count -ge 2 -and $maj.Count -lt $ws.Count) { $nom = $maj -join ' ' }
    return $nom
}

# "designa com a Tecnica Responsable del projecte ... a Z, Enginyer": Z es el
# que va darrere de l'ULTIMA " a " abans de la coma (o de "per", "amb DNI"...).
# Abans la "a" s'havia de trobar en 80 caracters i, si no, l'autoritzat era
# "com a Tecnica Responsable del projecte..." (2a ronda, GIA 5). Torna
# @{ Ini; Fi } (posicions al text pla) o $null. PURA.
function _CtZDespresDeCarrec([string]$pla, [int]$des) {
    $m = [regex]::Match($pla.Substring($des), '^\s*(?:com\s+a|como)\s+')
    if (-not $m.Success) { return $null }
    $ini = $des + $m.Length
    $resta = $pla.Substring($ini)
    $mf = [regex]::Match($resta, ',|;|\.\s|\s+(?:per|perque|para|amb|con|a\s+fi)\s')
    $seg = if ($mf.Success) { $resta.Substring(0, $mf.Index) } else { $resta }
    $ma = @([regex]::Matches($seg, '\s(?:a|al)\s+'))
    if ($ma.Count -eq 0) { return $null }
    $u = $ma[$ma.Count - 1]
    $zi = $u.Index + $u.Length
    $mt = [regex]::Match($seg.Substring($zi), '^(?:la|el)\s+(?=(?:senyor|senyora|sra?\.?|d\.|dna|dona|don)\b)')
    if ($mt.Success) { $zi += $mt.Length }
    if ($seg.Length - $zi -lt 3) { return $null }
    return @{ Ini = $ini + $zi; Fi = $ini + $seg.Length }
}

function Read-AutoritzacioText([string]$text) {
    $aj = _CtAjuntaLinies $text
    $orig = [string]$aj.Text
    if ($orig -eq '') { return $null }
    $pla = _CtPla $orig
    $mz = [regex]::Match($pla, $Script:CtReAutoritza)
    if (-not $mz.Success) { return $null }
    $abans = $pla.Substring(0, $mz.Index)
    $mx = [regex]::Match($abans, $Script:CtReAutCa)
    if (-not $mx.Success) { $mx = [regex]::Match($abans, $Script:CtReAutEs) }
    $my = [regex]::Match($abans, $Script:CtReEnNom)
    $senseJo = $false
    if (-not $mx.Success -and $my.Success) {
        # L'ultim "<nom>, amb DNI" abans de "en nom de".
        foreach ($c in [regex]::Matches($abans.Substring(0, $my.Index), $Script:CtReAutNomDni)) { $mx = $c; $senseJo = $true }
    }
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
        if ($senseJo) { $nomX = _CtNomSignantSenseJo $orig $pla $gx.Index ($gx.Index + $gx.Length) $aj.Salts }
        else {
            $nomX = $orig.Substring($gx.Index, $gx.Length)
            $mt = [regex]::Match($pla.Substring($gx.Index, $gx.Length), $Script:CtReTitolDoc)
            if ($mt.Success) { $nomX = $nomX.Substring($mt.Length) }
        }
        $signant = _CtPersona (_CtTallaNom $nomX) (& $primer $nifs ($gx.Index) $iniY)
    }
    $empresa = $null
    if ($my.Success) {
        $gy = $my.Groups['y']
        $empresa = _CtPersona ($orig.Substring($gy.Index, $gy.Length)) (& $primer $nifs ($gy.Index) $mz.Index)
    }
    $gz = $mz.Groups['z']
    $zIni = $gz.Index; $zFi = $gz.Index + $gz.Length
    $mv = [regex]::Match($pla.Substring($mz.Index), '^\b' + $Script:CtReVerbAut)
    $zc = if ($mv.Success) { _CtZDespresDeCarrec $pla ($mz.Index + $mv.Length) } else { $null }
    if ($null -ne $zc) { $zIni = $zc.Ini; $zFi = $zc.Fi }
    $aut = _CtPersona (_CtTallaNom ($orig.Substring($zIni, $zFi - $zIni))) (& $primer $nifs $zIni $orig.Length) (& $primer $emails $zIni $orig.Length)
    foreach ($t in $tels) { if ($t.Pos -ge $zIni) { _CtPosaTelefon $aut $t.Val } }
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
    $r.signatures = @(_CtSignaturesRep $text)
    return $r
}
