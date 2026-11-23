#requires -Version 5.1
<#
.SYNOPSIS
  "Requeriment - Seguiment": les regles PURES que miren el document sencer per
  decidir com acaba el seguiment: on comencen les conclusions que s'esborren,
  quines s'hi tornen a posar (el 59.1.d d'una MNS), de quina familia es
  l'informe i com es diu el fitxer de sortida.

  Es van partir de Seguiment.ps1 (novembre 2026) quan aquell va passar de les
  1.200 linies; com tot el programa, va amb dot-source al mateix ambit
  (Motor.ps1 el carrega darrere de Seguiment.ps1). Nomes defineix funcions.
#>

# El paragraf que es NOMES el titol "CONCLUSIONS" (sense accents ni
# majuscules, i sense signes: "CONCLUSIONS:" tambe), despres de l'ultim
# requeriment. Index 1-based, o -1. PURA.
function _TitolConclusionsIndex($paraTexts, [int]$lastReqEndIndex) {
    if ($null -eq $paraTexts) { return -1 }
    $l = @($paraTexts)
    for ($i = [math]::Max(0, $lastReqEndIndex); $i -lt $l.Count; $i++) {
        if (((_NormalitzaText ([string]$l[$i])) -replace '[^a-z]', '') -eq 'conclusions') { return ($i + 1) }
    }
    return -1
}

# LES CONCLUSIONS QUE ES MANTENEN EN UN SEGUIMENT. El seguiment esborra el
# bloc de conclusions de l'informe i hi posa les seves; pero l'avis de l'article
# 59.1.d d'una MNS ("aquestes Modificacions NO Substancials han de figurar a
# les posteriors actes dels controls periodics") val igual despres (l'usuari,
# novembre 2026: "haura de continuar constant"). Torna, en l'ordre del
# document, els paragrafs del bloc esborrat que son alguna de $frases (sense
# accents, majuscules ni signes). PURA.
function _ConclusionsQueEsMantenen($paraTexts, [int]$startIndex, $frases) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $paraTexts -or $startIndex -lt 1) { return @() }
    $netes = @(@($frases) | ForEach-Object { (_NormalitzaText ([string]$_)) -replace '[^a-z0-9]', '' } | Where-Object { $_ -ne '' })
    if ($netes.Count -eq 0) { return @() }
    $l = @($paraTexts)
    for ($i = $startIndex - 1; $i -lt $l.Count; $i++) {
        $t = [string]$l[$i]
        if ($netes -contains ((_NormalitzaText $t) -replace '[^a-z0-9]', '')) { [void]$out.Add($t.Trim()) }
    }
    return $out.ToArray()
}

# Es una MNS? Mira el text que hi ha ABANS del primer requeriment (la capcalera
# "ASSUMPTE ... Modificacio NO Substancial" i el "En relacio a la sol.licitud de
# Modificacio NO Substancial..."): una observacio que parli d'una MNS no compta.
# Fa falta perque les MNS fetes abans del programa d'ara no porten el 59.1.d, i
# el seguiment l'ha de posar igualment (l'usuari, novembre 2026, amb la MNS del
# GIA 341 de juliol). $primerReq: index 1-based del primer requeriment (0 =
# cap). PURA.
function _EsInformeMns($paraTexts, [int]$primerReq) {
    if ($null -eq $paraTexts) { return $false }
    $l = @($paraTexts)
    $fins = if ($primerReq -ge 1) { [math]::Min($primerReq - 1, $l.Count) } else { $l.Count }
    for ($i = 0; $i -lt $fins; $i++) {
        if ((_NormalitzaText ([string]$l[$i])) -match 'modificacio no substancial') { return $true }
    }
    return $false
}

# De quina familia es l'informe, per al nom del seguiment quan l'original no
# porta "Req" (vegeu _SeguimentOutputName): 'MNS', 'Trans' o '' (un requeriment
# normal). Pel text d'abans del primer requeriment, com _EsInformeMns. PURA.
function _FamiliaSeguiment($paraTexts, [int]$primerReq) {
    if (_EsInformeMns $paraTexts $primerReq) { return 'MNS' }
    $l = @($paraTexts)
    $fins = if ($primerReq -ge 1) { [math]::Min($primerReq - 1, $l.Count) } else { $l.Count }
    for ($i = 0; $i -lt $fins; $i++) {
        if ((_NormalitzaText ([string]$l[$i])) -match 'transmissio de la titularitat') { return 'Trans' }
    }
    return ''
}

# Les frases que es mantenen: les conclusions de la MNS del cataleg de
# conclusions (avui, nomes el 59.1.d). No llanca.
function _FrasesConclusioPermanents {
    try {
        $mns = Read-Conclusions $ConclusionsPath 'MNS'
        return @(@($mns.Selectable) | ForEach-Object { [string]$_.Body } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } catch { return @() }
}

# Nom del fitxer de sortida del seguiment. La DATA es la d'AVUI (el dia que es
# genera l'informe), no la de l'anotacio. S'incrementa el numero de ronda "Req":
#   - "...Req1..."           -> "<avui>_...Req2...", "Req2"->"Req3"...
#   - "...Req..." sense num   -> "<avui>_...Req2..." (requeriments antics fets
#                                sense el programa, que nomes posen "Req").
#   - amb el GIA ($gia), el que va DARRERE del Req (el titular dels noms antics)
#     es torna "_GIA n": "Req_SERVICIOS Y ..." -> "Req2_GIA n".
#   - sense "Req" enlloc, amb el GIA ($gia, de la capcalera del document) ->
#     "<avui>_[<familia>_]Req2_GIA <n>.docx" (novembre 2026: el seguiment d'una
#     MNS antiga "2026-07-01_D112_MNS_OK_JOSE_PEREZ" es "..._MNS_Req2_GIA 341").
#   - i sense GIA             -> "<avui>_Seguiment_<nom>.docx", com abans.
# Es treu la data inicial del nom origen i es preserva la resta (GIA, titular...).
# Sempre s'ha de passar el resultat per _GetUniqueOutputPath (afegeix _2, _3...).
function _SeguimentOutputName([string]$sourceBaseName, [datetime]$today, [string]$familia = '', [string]$gia = '') {
    $day  = $today.ToString('yyyy-MM-dd')
    # Treu una data inicial "YYYY-MM-DD" (amb "_" o espai darrere) del nom origen.
    $rest = [regex]::Replace([string]$sourceBaseName, '^\s*\d{4}-\d{2}-\d{2}[_ ]+', '')
    # Busca "Req" amb numero opcional i l'incrementa (sense numero -> 2).
    # "Req" NOMES si no el segueix cap lletra: el "REQ" de "REQUERIMENT" d'un nom
    # antic donava "Req2UERIMENT" (novembre 2026). "LlicReq", "Req1", "Req_" si.
    $m = [regex]::Match($rest, '(?i)Req(\d*)(?![a-z])')
    if ($m.Success) {
        $n = if ($m.Groups[1].Value -ne '') { [int]$m.Groups[1].Value + 1 } else { 2 }
        if (-not [string]::IsNullOrWhiteSpace($gia)) {
            # AMB EL GIA DEL DOCUMENT, el que va darrere del Req es substitueix
            # per "_GIA n" (l'usuari, novembre 2026: els noms antics acaben en el
            # titular, "2026-04-01_Req_SERVICIOS Y DISTRIBUCIONES TECNICAS", i el
            # seguiment ha de ser "..._Req2_GIA n"). El que hi ha DAVANT del Req
            # es conserva ("MNS_", "Llic_"); si no n'hi ha, el tipus del document.
            $davant = $rest.Substring(0, $m.Index)
            if ($davant.Trim() -eq '' -and -not [string]::IsNullOrWhiteSpace($familia)) { $davant = $familia.Trim() + '_' }
            $name = '{0}_{1}Req{2}_GIA {3}' -f $day, $davant, $n, $gia.Trim()
        } else {
            $rest = $rest.Remove($m.Index, $m.Length).Insert($m.Index, "Req$n")
            $name = '{0}_{1}' -f $day, $rest
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($gia)) {
        $pref = if ([string]::IsNullOrWhiteSpace($familia)) { '' } else { $familia.Trim() + '_' }
        $name = '{0}_{1}Req2_GIA {2}' -f $day, $pref, $gia.Trim()
    } else {
        $safe = $rest
        if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'informe' }
        $name = '{0}_Seguiment_{1}' -f $day, $safe
    }
    $name = ($name -replace '[\\/:*?"<>|]','_').Trim()
    return ($name + '.docx')
}
