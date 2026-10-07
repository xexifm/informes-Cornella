#requires -Version 5.1
<#
  Eina "Revisar requeriments": les decisions (PURES; es proven a Linux). La
  finestra i la xarxa son a Revisio.ps1.

  Peticio de l'usuari (setembre 2026): una eina per mantenir el programa al dia
  que miri si la normativa dels requeriments encara es vigent, si els enllacos
  funcionen, si cal baixar normativa nova que substitueixi l'anterior i si tots
  els requeriments tenen la fitxa d'informacio (la i): n'hi va afegint de nous i
  no sempre la porten.

  QUE ES DERIVA D'ON:
    - la vigencia, de la pagina de cada norma (BOE: el text de la pagina;
      Portal Juridic: la etiqueta VIGENT / DEROGAT del costat del titol, que es
      munta amb JavaScript i per aixo es llegeix del DOM de l'Edge);
    - la norma que la substitueix, del mateix BOE ("... por el Real Decreto X
      (Ref. BOE-A-...)");
    - els punts que la citen, de _NormativaPuntsReq1 (NormativaDades.ps1).
  Cap d'aquestes coses canvia REQ1 tota sola: l'informe diu que cal fer i on.
#>

# El text d'una pagina, sense etiquetes ni scripts. PURA.
function _RevTextDeHtml([string]$html) {
    if ([string]::IsNullOrWhiteSpace($html)) { return '' }
    $t = [regex]::Replace($html, '(?is)<(script|style)\b.*?</\1>', ' ')
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ($t -replace '\s+', ' ').Trim()
}

# VIGENT O DEROGADA, segons la pagina del BOE. PURA. Torna @{ Estat; Detall;
# Substituta; SubstitutaId }. Estat: 'vigent' | 'derogada' | '?' (no es veu clar:
# val mes dir-ho que endevinar-ho). Una derogacio PARCIAL ("se deroga el art. 5")
# no fa derogada la norma.
function _RevEstatBoe([string]$html) {
    $t = _RevTextDeHtml $html
    $out = @{ Estat = '?'; Detall = ''; Substituta = ''; SubstitutaId = '' }
    if (-not $t) { return $out }
    # NOMES les marques d'ESTAT de la pagina: la capcalera "Norma derogada, con
    # efectos de..." i l'etiqueta "[Disposicion derogada]" del costat del titol.
    # Abans es buscava "norma anulada" o "fue derogada por" a TOT el text, i el
    # preambul en parla sovint d'ALTRES normes ("la version de la norma anulada",
    # "la Directiva 95/16/CE fue derogada por..."): el REBT, el reglament del gas,
    # el d'alta tensio i el d'ascensors sortien derogats (octubre 2026).
    $derog = [regex]::Match($t, '(?i)\bnorma\s+(derogada|anulada)\s*,?\s+con\s+efectos\b|\[\s*disposici[oó]n\s+(derogada|anulada)\s*\]|\bestado\s*:?\s*derogad[ao]\b')
    if ($derog.Success) {
        $out.Estat = 'derogada'
        $ini = [Math]::Max(0, $derog.Index - 40)
        $out.Detall = $t.Substring($ini, [Math]::Min(260, $t.Length - $ini)).Trim()
        $sub = [regex]::Match($t.Substring($derog.Index), '(?i)\bpor\s+(?:el\s+|la\s+)?(.{5,200}?)\s*\(Ref\.\s*(BOE-A-\d{4}-\d+)\)')
        if ($sub.Success) { $out.Substituta = $sub.Groups[1].Value.Trim().TrimEnd(','); $out.SubstitutaId = $sub.Groups[2].Value }
        return $out
    }
    if ($t -match '(?i)\b(estado|situaci[oó]n)\s*:?\s*vigente\b' -or $t -match '(?i)\bnorma\s+vigente\b' -or $t -match '(?i)texto consolidado') { $out.Estat = 'vigent' }
    else { $out.Detall = "La pàgina del BOE no diu l'estat (potser l'enllaç és al text original i no al consolidat)." }
    return $out
}

# PER QUE NO ES SAP SI UNA NORMA DEL PORTAL JURIDIC ES VIGENT. PURA. $dom: el
# que ha tornat l'Edge; $errEdge: el seu error ('' si no n'hi ha); $nTextos:
# quantes respostes del servidor s'han mirat.
function _RevMotiuPjur([string]$dom, [string]$errEdge, [int]$nTextos) {
    if ($errEdge) { return ("L'Edge no ha pogut obrir la pàgina: " + $errEdge) }
    if (-not $dom) { return "L'Edge no ha tornat res." }
    $base = "La pàgina no mostra l'etiqueta VIGENT ni NO VIGENT"
    if ($nTextos -eq 0) { return ($base + " (i el servidor no ha respost sense l'Edge).") }
    return ($base + '.')
}

# VIGENT O DEROGADA, segons el Portal Juridic (el DOM ja dibuixat). PURA.
# L'etiqueta va entre la barra de descarregues ("Copia la URI ELI") i el titol,
# en MAJUSCULES; es mira NOMES alli, perque dins del text consolidat hi pot
# haver "(Derogat)" al costat d'un article, i aixo no vol dir que la norma ho
# sigui. Sense la barra, la primera etiqueta en majuscules que surti.
function _RevEstatPjur([string]$html) {
    $t = _RevTextDeHtml $html
    $out = @{ Estat = '?'; Detall = ''; Substituta = ''; SubstitutaId = '' }
    if (-not $t) { return $out }
    $zona = $t
    $m = [regex]::Match($t, '(?i)Copia la URI ELI')
    if ($m.Success) { $zona = $t.Substring($m.Index + $m.Length, [Math]::Min(160, $t.Length - $m.Index - $m.Length)) }
    $e = [regex]::Match($zona, '(?-i)\b(NO VIGENT|VIGENT|DEROGADA|DEROGAT|ANUL.LAT|ANUL.LADA)\b')
    if ($e.Success) {
        $out.Estat = if ($e.Groups[1].Value -eq 'VIGENT') { 'vigent' } else { 'derogada' }
        $out.Detall = $e.Groups[1].Value
    }
    return $out
}

# VIGENT O DEROGADA, segons les metadades ELI (RDF/TTL/XML del Portal Juridic).
# PURA. L'ontologia ELI ho diu amb un valor fix (eli:in_force):
#   ...#InForce-inForce | #InForce-notInForce | #InForce-partiallyInForce
# Parcialment vigent vol dir que la norma segueix: es 'vigent'.
function _RevEstatEli([string]$text) {
    $out = @{ Estat = '?'; Detall = ''; Substituta = ''; SubstitutaId = '' }
    $m = [regex]::Match([string]$text, '(?i)InForce-(notInForce|partiallyInForce|inForce)')
    if ($m.Success) {
        $v = $m.Groups[1].Value
        $out.Estat = if ($v -ieq 'notInForce') { 'derogada' } else { 'vigent' }
        $out.Detall = 'ELI: ' + $v
    }
    return $out
}

# ELS PUNTS SENSE FITXA D'INFORMACIO (la i del Pas 3). PURA, sobre el JSON del
# cataleg (el que torna Read-JsonFile). Nomes punts i sub-punts: les seccions i
# els textos no en porten.
function _RevPuntsSenseFitxa($o) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $o) { return $out.ToArray() }
    $visita = $null
    $visita = {
        param($nodes, [string]$seccio)
        foreach ($n in @($nodes)) {
            if ($null -eq $n) { continue }
            $sec = if ([string]$n.tipus -eq 'seccio') { [string]$n.titol } else { $seccio }
            if (@('item', 'subitem') -contains [string]$n.tipus) {
                $a = $n.ajuda
                $buida = ($null -eq $a) -or ([string]::IsNullOrWhiteSpace([string]$a.norma) -and [string]::IsNullOrWhiteSpace([string]$a.criteri))
                if ($buida) { [void]$out.Add([pscustomobject]@{ Seccio = $sec; Punt = [string]$n.titol; Tipus = [string]$n.tipus }) }
            }
            if ($n.fills) { & $visita $n.fills $sec }
        }
    }
    & $visita $o.nodes ''
    return $out.ToArray()
}

# Les columnes de l'informe de la revisio.
$Script:RevCapcalera = @('Revisió', 'Catàleg', 'Punt', 'Què passa', 'Enllaç', 'Què cal fer')
$Script:RevAmples    = @(22, 10, 45, 70, 12, 60)

# L'Excel de l'informe de revisio. En una funcio perque qui el desa ho fa des
# d'una closure (el boto de la pantalla), on $Script:RevCapcalera i
# $Script:RevAmples valien buit: l'informe sortia sense capcalera ni amples.
function _RevInformeXlsxBytes($files) {
    return (_NormativaXlsxBytes $Script:RevCapcalera @($files) $Script:RevAmples 'Revisio')
}

function _RevFila([string]$revisio, [string]$cataleg, [string]$punt, [string]$que, [string]$url, [string]$fer) {
    $enllac = if ($url) { @{ Text = 'Obrir'; Link = $url } } else { '' }
    return , @($revisio, $cataleg, $punt, $que, $enllac, $fer)
}
