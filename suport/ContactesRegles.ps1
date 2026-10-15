#requires -Version 5.1
<#
  ContactesRegles.ps1 - DE QUI ES CADA DADA (les regles) i QUE NO QUADRA amb
  l'Excel (els avisos). Tot PUR: rep els documents ja llegits
  (ContactesExtraccio.ps1), la fila de l'Excel i el context de tota la base, i
  torna l'activitat tal com es desa a contactes-db.json.

  LES REGLES, de l'usuari (octubre 2026), i per que cadascuna:
   1. EL TECNIC NO ES MAI EL REPRESENTANT LEGAL. Tecnics, enginyers,
      arquitectes, gestories i tramitadors son PERSONES AUTORITZADES. Es
      reconeixen (_CtMotiuTecnic) perque son a tecnics-coneguts_*.json, perque
      el seu correu o telefon surt a activitats de TITULARS DIFERENTS (el senyal
      mes fort: es compta sobre tota la base, _CtContext), perque el correu es
      professional (engin, arquitec, gestor...), perque una autoritzacio els
      autoritza, o perque presenten en nom d'un titular diferent amb un NIF
      d'entitat (una gestoria).
   2. REPRESENTANT LEGAL: persona fisica -> ella mateixa, es deixa EN BLANC.
      Persona juridica -> qui signa en nom de l'empresa (l'autoritzacio,
      "segur"), o el representant d'un tramit que no sigui tecnic ("probable").
      Si no es troba, en blanc. Mai el tecnic.
   3. EL TITULAR = l'interessat, TRET del que sigui del tecnic: un correu o un
      telefon d'una persona autoritzada no pot ser del titular.
   4. L'ESTABLIMENT: el telefon i el correu de l'e-TRAM, si no son del tecnic.
   5. FORA: els interessats de les QUEIXES (serie 2579: son els veins), tret
      que el NIF sigui el del titular; les administracions; les entitats de
      control (OCA).
   6. RECINTES amb organitzadors d'actes ($Script:CtRecintes): els documents
      son dels organitzadors. No es proposa res del titular ni del
      representant.
   7. EL MES RECENT MANA (per la data del document). La resta, a l'historic.
   8. Cada dada porta font (ruta relativa), data i confianca.

  El format de cada activitat es el de la REFERENCIA (contactes-referencia_
  *.json), perque ValidarContactes.ps1 la pugui comparar camp a camp.
  NOMES DEFINEIX FUNCIONS.
#>

# Els recintes on els documents son dels organitzadors (l'usuari: "ara GIA 28 i
# GIA 1324"). Es canvia a contactes-config.json de local\base-dades-activitats
# (Read-ContactesConfig, ContactesDb.ps1).
$Script:CtRecintesDefecte = @('28', '1324')

# Dominis que no son mai del titular: administracions i entitats de control.
$Script:CtDominisAdmin = @('cornella.cat', 'aj-cornella.cat', 'amb.cat', 'gencat.cat', 'aoc.cat', 'mossos.cat', 'bombers.cat', 'diba.cat', 'interior.gencat.cat')
# Sobre el DOMINI del correu: un tros acabat en "-" ("eca-") nomes al principi
# del domini o d'una part; un amb punt ("tuv.com"), el domini o el final.
$Script:CtEntitatsControl = @('aucatel', 'tuvsud', 'tuv-sud', 'tuv.com', 'sgs.com', 'dekra', 'bureauveritas', 'bureau-veritas', 'applus', 'ocacert', 'atisae', 'eurocontrol', 'ecacert', 'idiada', 'oca-', 'ocaglobal', 'eca-', 'ecaglobal', 'eicon', 'addient', 'noverca', 'cualicontrol', 'kiwa')
# El correu d'un professional (regla 1). Sobre el text pla de l'adreca.
$Script:CtPatronsProfessional = @('engin', 'ingen', 'arquitec', 'projec', 'proyec', 'consult', 'gestor', 'gestoria', 'gestio', 'asesor', 'assessor', 'asesores', 'assessors', 'tramit', 'licenc', 'llicen', 'tecnic', 'estudi', 'oficina', 'enginyers.net', 'apabcn', 'coac', 'caateeb', 'aparellador', 'aparejador', 'industrial', 'enginy', 'advocat', 'abogad', 'icab.cat', 'legal', 'lawyer', 'ebcn')

function _CtDomini([string]$email) {
    $e = _CtEmailNet $email
    $i = $e.LastIndexOf('@')
    if ($i -lt 0) { return '' }
    return $e.Substring($i + 1)
}

function _CtEsDominiDe([string]$email, [string[]]$dominis) {
    $d = _CtDomini $email
    if ($d -eq '') { return $false }
    foreach ($x in $dominis) { if ($d -eq $x -or $d.EndsWith('.' + $x)) { return $true } }
    return $false
}

# Un correu que no pot ser mai del titular (regla 5). PURA.
function _CtEmailExclos([string]$email) {
    $e = _CtEmailNet $email
    if ($e -eq '') { return $false }
    if (_CtEsDominiDe $e $Script:CtDominisAdmin) { return $true }
    if ($e -match 'mossos|bombers') { return $true }
    $d = _CtDomini $e
    foreach ($x in $Script:CtEntitatsControl) {
        if ($x.EndsWith('-')) { if ($d.StartsWith($x) -or $d.Contains('.' + $x)) { return $true } }
        elseif ($x.Contains('.')) { if ($d -eq $x -or $d.EndsWith('.' + $x)) { return $true } }
        elseif ($d.Contains($x)) { return $true }
    }
    return $false
}

# Un correu d'assessoria o de gestoria PEL NOM (l'usuari, 2a ronda:
# "asesoriarovira2000@hotmail.es" sortia de correu del titular). La part d'abans
# de l'@ ho diu i el nom del titular no. PURA.
$Script:CtPatronsGestoriaNom = @('asesor', 'assessor', 'gestor', 'gestoria', 'advocat', 'abogad', 'ingenier', 'enginy', 'arquitect')
function _CtCorreuGestoriaPelNom([string]$email, [string]$titular) {
    $e = _CtEmailNet $email
    if ($e -eq '') { return $false }
    $local = _CtPla ($e.Substring(0, $e.IndexOf('@')))
    $tn = _CtPla $titular
    foreach ($x in $Script:CtPatronsGestoriaNom) { if ($local.Contains($x) -and -not $tn.Contains($x)) { return $true } }
    return $false
}

# El domini es el d'un tecnic de la llista, o s'hi assembla (un error de tecleig,
# distancia <= 2: "inaargesa.com" / "inargesa.com", GIA 32). Els dominis dels
# proveidors corrents (gmail...) no compten. Torna el domini de la llista o ''.
# PURA (amb memoria al context: la distancia es cara).
function _CtDominiTecnic([string]$email, $ctx) {
    $d = _CtDomini $email
    if ($d -eq '' -or $null -eq $ctx.Tecnics.Dominis -or $Script:CtDominisComuns -contains $d) { return '' }
    if ($ctx.Tecnics.Dominis.ContainsKey($d)) { return $d }
    if ($null -eq $ctx.MemoDomini) { $ctx.MemoDomini = @{} }
    if ($ctx.MemoDomini.ContainsKey($d)) { return [string]$ctx.MemoDomini[$d] }
    $r = ''
    if ($d.Length -ge 8) {
        foreach ($x in @($ctx.Tecnics.Dominis.Keys)) {
            if ([Math]::Abs($x.Length - $d.Length) -gt 2) { continue }
            if ((_CtDistancia $d $x) -le 2) { $r = [string]$x; break }
        }
    }
    $ctx.MemoDomini[$d] = $r
    return $r
}

# El correu es a la llista de tecnics, pero la persona que el porta NO es cap
# dels noms d'aquella entrada (l'usuari, GIA 87: l'administrador presenta amb el
# correu de la seva gestoria). Llavors la persona conserva el seu paper i el
# correu es de la gestoria. Torna l'entrada de la llista o $null. PURA.
function _CtCorreuTecnicAltri($p, $ctx) {
    $e = _CtEmailNet ([string](_CtV $p 'email'))
    if ($e -eq '' -or -not $ctx.Tecnics.Emails.ContainsKey($e)) { return $null }
    $ent = $ctx.Tecnics.Emails[$e]
    $noms = @(@($ent.noms) | Where-Object { @(_CtNomParaules ([string]$_)).Count -ge 2 })
    $nom = [string](_CtV $p 'nom')
    if ($noms.Count -eq 0 -or @(_CtNomParaules $nom).Count -lt 2) { return $null }
    foreach ($n in $noms) { if (_CtMateixNom $nom ([string]$n)) { return $null } }
    return $ent
}

function _CtEmailProfessional([string]$email) {
    $e = _CtPla (_CtEmailNet $email)
    if ($e -eq '') { return $false }
    foreach ($x in $Script:CtPatronsProfessional) { if ($e.Contains($x)) { return $true } }
    return $false
}

# ----------------------------------------------------------------------------
# EL CONTEXT DE TOTA LA BASE (regla 1: el correu que surt a titulars diferents)
# ----------------------------------------------------------------------------
# $docsPerGia: gia -> llista de documents (amb .extret); $excel: gia -> fila
# de contactes de l'Excel; $tecnics: el de Read-TecnicsConeguts. PURA.
function _CtIdentitatTitular($docs, $xl) {
    $nif = _CtNifNet ([string](_CtV $xl 'NIF'))
    if ($nif -ne '') { return $nif }
    foreach ($d in @($docs)) {
        $i = _CtV (_CtV $d 'extret') 'interessat'
        $n = _CtNifNet ([string](_CtV $i 'nif'))
        if ($n -ne '') { return $n }
    }
    $t = _CtNomNorm ([string](_CtV $xl 'TITULAR'))
    if ($t -ne '') { return $t }
    return ''
}

# Els titulars VINCULATS no son "diferents" (l'usuari, 2a ronda: empreses
# germanes, el mateix propietari): comparteixen el representant legal de
# l'Excel (el NIF o el nom) o el domini del correu. Nomes si el comparteixen
# com a molt 3 titulars i no te res de tecnic: el cas mes frequent dels errors
# de l'Excel es el TECNIC posat de representant legal a moltes activitats, i
# aquell no ha de fer de pont. Torna titular -> arrel del seu grup. PURA.
function _CtArrelGrup($pare, [string]$x) {
    $r = $x
    while ($pare.ContainsKey($r) -and [string]$pare[$r] -ne $r) { $r = [string]$pare[$r] }
    return $r
}

function _CtGrupsTitulars($titPerGia, $excel, $tecnics) {
    $pare = @{}
    foreach ($t in @($titPerGia.Values)) { $pare[[string]$t] = [string]$t }
    $perClau = @{}
    $afegeix = { param([string]$k, [string]$t) if ($k -eq '') { return }; if (-not $perClau.ContainsKey($k)) { $perClau[$k] = @{} }; $perClau[$k][$t] = $true }
    foreach ($g in @($titPerGia.Keys)) {
        $xl = if ($null -ne $excel -and $excel.ContainsKey($g)) { $excel[$g] } else { $null }
        if ($null -eq $xl) { continue }
        $t = [string]$titPerGia[$g]
        $rn = _CtNifNet ([string](_CtV $xl 'REP_NIF'))
        if ((_CtTipusNif $rn) -eq 'fisica') { & $afegeix ('n:' + $rn) $t }
        $rnom = [string](_CtV $xl 'REP_NOM')
        $pw = @(_CtNomParaules $rnom)
        if ($pw.Count -ge 2 -and @($pw | Where-Object { -not $Script:CtNomsPila.ContainsKey($_) }).Count -gt 0) {
            $tecNom = $false
            foreach ($n in @($tecnics.Noms)) { if (_CtMateixNom $rnom $n) { $tecNom = $true; break } }
            if (-not $tecNom) { & $afegeix ('m:' + (_CtNomNorm $rnom)) $t }
        }
        foreach ($k in 'EMAIL', 'REP_EMAIL') {
            $e = _CtEmailNet ([string](_CtV $xl $k))
            $d = _CtDomini $e
            if ($d -eq '' -or $Script:CtDominisComuns -contains $d -or (_CtEmailExclos $e) -or (_CtEmailProfessional $e) -or $tecnics.Emails.ContainsKey($e)) { continue }
            if ($null -ne $tecnics.Dominis -and $tecnics.Dominis.ContainsKey($d)) { continue }
            & $afegeix ('d:' + $d) $t
        }
    }
    foreach ($k in @($perClau.Keys)) {
        $ts = @($perClau[$k].Keys)
        if ($ts.Count -lt 2 -or $ts.Count -gt 3) { continue }
        $a = _CtArrelGrup $pare $ts[0]
        for ($i = 1; $i -lt $ts.Count; $i++) { $b = _CtArrelGrup $pare $ts[$i]; if ($a -ne $b) { $pare[$b] = $a } }
    }
    $out = @{}
    foreach ($t in @($pare.Keys)) { $out[$t] = _CtArrelGrup $pare $t }
    return $out
}

function _CtContext($docsPerGia, $excel, $tecnics, $recintes = $null) {
    $ctx = @{
        Tecnics = if ($null -ne $tecnics) { $tecnics } else { @{ Emails = @{}; Noms = @(); Dominis = @{} } }
        EmailTitulars = @{}; TelTitulars = @{}; NifTitulars = @{}; NomTitulars = @{}
        EmailRol = @{}; TelRol = @{}; MemoDomini = @{}
        Recintes = @(if ($null -ne $recintes) { @($recintes | ForEach-Object { [string]$_ }) } else { $Script:CtRecintesDefecte })
    }
    $apunta = {
        param($mapa, [string]$k, [string]$tit)
        if ($k -eq '' -or $tit -eq '') { return }
        if (-not $mapa.ContainsKey($k)) { $mapa[$k] = @{} }
        $mapa[$k][$tit] = $true
    }
    $gies = @{}
    foreach ($g in @($docsPerGia.Keys)) { $gies[[string]$g] = $true }
    if ($null -ne $excel) { foreach ($g in @($excel.Keys)) { $gies[[string]$g] = $true } }
    $titPerGia = @{}
    foreach ($g in @($gies.Keys)) {
        if ($ctx.Recintes -contains $g) { continue }
        $docs = if ($docsPerGia.ContainsKey($g)) { @($docsPerGia[$g]) } else { @() }
        $xl = if ($null -ne $excel -and $excel.ContainsKey($g)) { $excel[$g] } else { $null }
        $tit = _CtIdentitatTitular $docs $xl
        if ($tit -eq '') { $tit = 'GIA ' + $g }
        $titPerGia[$g] = $tit
    }
    $grups = _CtGrupsTitulars $titPerGia $excel $ctx.Tecnics
    foreach ($g in @($titPerGia.Keys)) {
        $docs = if ($docsPerGia.ContainsKey($g)) { @($docsPerGia[$g]) } else { @() }
        $xl = if ($null -ne $excel -and $excel.ContainsKey($g)) { $excel[$g] } else { $null }
        $tit = [string]$grups[[string]$titPerGia[$g]]
        $persones = New-Object System.Collections.ArrayList
        $presenten = New-Object System.Collections.ArrayList
        foreach ($d in $docs) {
            $e = _CtV $d 'extret'
            $p = _CtV $e 'interessat'; if ($null -ne $p) { [void]$persones.Add($p) }
            $p = _CtV $e 'representant'; if ($null -ne $p) { [void]$persones.Add($p); [void]$presenten.Add($p) }
            $a = _CtV $e 'autoritzacio'
            if ($null -ne $a) { foreach ($p in @(_CtV $a 'autoritzats')) { if ($null -ne $p) { [void]$persones.Add($p); [void]$presenten.Add($p) } } }
        }
        foreach ($p in $persones) {
            & $apunta $ctx.EmailTitulars (_CtEmailNet ([string](_CtV $p 'email'))) $tit
            foreach ($k in 'telefon', 'mobil') { & $apunta $ctx.TelTitulars (_CtTelefonNet ([string](_CtV $p $k))) $tit }
        }
        # El tecnic sovint posa el correu del client (l'usuari): de qui presenta
        # pel titular (el representant del tramit i l'autoritzat) es compta
        # tambe el NIF i el nom.
        foreach ($p in $presenten) {
            & $apunta $ctx.EmailRol (_CtEmailNet ([string](_CtV $p 'email'))) $tit
            foreach ($k in 'telefon', 'mobil') { & $apunta $ctx.TelRol (_CtTelefonNet ([string](_CtV $p $k))) $tit }
            & $apunta $ctx.NifTitulars (_CtNifNet ([string](_CtV $p 'nif'))) $tit
            $nom = [string](_CtV $p 'nom')
            if (@(_CtNomParaules $nom).Count -ge 2) { & $apunta $ctx.NomTitulars (_CtNomNorm $nom) $tit }
        }
        if ($null -ne $xl) {
            foreach ($k in 'EMAIL', 'REP_EMAIL') { & $apunta $ctx.EmailTitulars (_CtEmailNet ([string](_CtV $xl $k))) $tit }
            foreach ($k in 'TELEFON', 'MOBIL', 'REP_TELEFON', 'REP_MOBIL') { & $apunta $ctx.TelTitulars (_CtTelefonNet ([string](_CtV $xl $k))) $tit }
        }
    }
    return $ctx
}

# Per que una persona es tecnic (regla 1), o '' si no ho sembla. -Fort: nomes
# els senyals que no fallen (la llista i els titulars diferents). El correu
# "professional" i el nom no serveixen per treure un correu al TITULAR: un
# "estudi de dansa" o una "oficina" tambe poden ser el titular. PURA.
# "Titulars diferents" (l'usuari, 2a ronda: "surt a 2 activitats es massa
# poc"): 3 o mes, o 2 si hi ha un altre senyal. PURA.
function _CtProuTitulars($mapa, [string]$k, [bool]$senyal) {
    if ($k -eq '' -or $null -eq $mapa -or -not $mapa.ContainsKey($k)) { return 0 }
    $n = $mapa[$k].Count
    if ($n -ge 3 -or ($n -eq 2 -and $senyal)) { return $n }
    return 0
}

function _CtMotiuTecnic($p, $ctx, [switch]$Fort) {
    if ($null -eq $p) { return '' }
    $e = _CtEmailNet ([string](_CtV $p 'email'))
    # El correu de la llista que porta UNA ALTRA persona (la de la gestoria):
    # no fa tecnic qui el porta.
    $altri = ($null -ne (_CtCorreuTecnicAltri $p $ctx))
    # (Ni el correu professional: es el de la gestoria, no el seu.)
    $prof = (-not $altri) -and (_CtEmailProfessional $e)
    if ($e -ne '' -and -not $altri) {
        if ($ctx.Tecnics.Emails.ContainsKey($e)) { return "es a la llista de t$([char]0x00E8)cnics coneguts" }
        $dt = _CtDominiTecnic $e $ctx
        if ($dt -ne '') { return ("el domini del correu " + $(if ($dt -eq (_CtDomini $e)) { "es" } else { "s'assembla al" }) + " d'un t$([char]0x00E8)cnic de la llista (" + $dt + ")") }
        $rol = ($null -ne $ctx.EmailRol -and $ctx.EmailRol.ContainsKey($e) -and $ctx.EmailRol[$e].Count -ge 2)
        $n = _CtProuTitulars $ctx.EmailTitulars $e ($prof -or $rol)
        if ($n -gt 0) { return ("el seu correu surt a " + $n + " activitats de titulars diferents") }
    }
    foreach ($k in 'mobil', 'telefon') {
        $t = _CtTelefonNet ([string](_CtV $p $k))
        if ($t -eq '') { continue }
        $rol = ($null -ne $ctx.TelRol -and $ctx.TelRol.ContainsKey($t) -and $ctx.TelRol[$t].Count -ge 2)
        $n = _CtProuTitulars $ctx.TelTitulars $t ($prof -or $rol)
        if ($n -gt 0) { return ("el seu tel$([char]0x00E8)fon surt a " + $n + " activitats de titulars diferents") }
    }
    $mi = _CtMotiuTecnicIdentitat $p $ctx
    if ($mi -ne '') { return $mi }
    if ($Fort) { return '' }
    if ($prof) { return 'correu de professional' }
    $nom = [string](_CtV $p 'nom')
    foreach ($n in @($ctx.Tecnics.Noms)) { if (_CtMateixNom $nom $n) { return "es a la llista de t$([char]0x00E8)cnics coneguts" } }
    return ''
}

# El mateix NIF o el mateix nom presentant per a titulars diferents: es tecnic
# (l'usuari: "compta tambe el NIF/DNI i el nom normalitzat del representant del
# tramit"). '' si no. PURA.
# 3 titulars o mes, o 2 amb un correu professional. Un nom ha de tenir un
# cognom (B1: "MARIA JOSE" sol no es ningu).
function _CtMotiuTecnicIdentitat($p, $ctx) {
    if ($null -eq $p -or $null -eq $ctx.NifTitulars) { return '' }
    $prof = _CtEmailProfessional ([string](_CtV $p 'email'))
    $n = _CtNifNet ([string](_CtV $p 'nif'))
    $c = _CtProuTitulars $ctx.NifTitulars $n $prof
    if ($c -gt 0) { return ("el seu DNI presenta per a " + $c + " titulars diferents") }
    $nom = [string](_CtV $p 'nom')
    $pw = @(_CtNomParaules $nom)
    if ($pw.Count -ge 2 -and @($pw | Where-Object { -not $Script:CtNomsPila.ContainsKey($_) }).Count -gt 0) {
        $c = _CtProuTitulars $ctx.NomTitulars (_CtNomNorm $nom) $prof
        if ($c -gt 0) { return ("el seu nom presenta per a " + $c + " titulars diferents") }
    }
    return ''
}

# ----------------------------------------------------------------------------
# UNA ACTIVITAT
# ----------------------------------------------------------------------------
# Qui signa amb (R:CIF), amb el correu i els telefons del mateix document si
# hi surt (de representant del tramit o d'interessat). PURA.
function _CtPersonaSignant($sg, $e) {
    $p = _CtPersona ([string](_CtV $sg 'nom')) ([string](_CtV $sg 'nif'))
    foreach ($k in 'representant', 'interessat') {
        $q = _CtV $e $k
        if ($null -eq $q) { continue }
        $nq = _CtNifNet ([string](_CtV $q 'nif'))
        if (-not (($nq -ne '' -and $nq -eq [string]$p.nif) -or (_CtMateixNom ([string](_CtV $q 'nom')) ([string]$p.nom)))) { continue }
        if ([string]$p.nom -eq '') { $p.nom = [string](_CtV $q 'nom') }
        foreach ($c in 'email', 'telefon', 'mobil') { if ([string]$p[$c] -eq '') { $p[$c] = [string](_CtV $q $c) } }
        break
    }
    return $p
}

# L'interessat d'una instancia o d'un e-TRAM es una PERSONA FISICA i el titular
# una ENTITAT (l'usuari, 2a ronda: GIA 17, GIA 293): candidat a representant
# legal ("probable"; "segur" si ho confirma la signatura (R:CIF) o
# l'autoritzacio, que ja passen al davant). PURA.
function _CtInteressatRepCandidat($e, [string]$tipusTit) {
    if ($tipusTit -ne 'juridica' -or [string](_CtV $e 'tipus') -eq 'autoritzacio') { return $false }
    return ((_CtTipusNif ([string](_CtV (_CtV $e 'interessat') 'nif'))) -eq 'fisica')
}

# ...i si a mes presenta en nom propi amb el correu de l'empresa, no es un
# canvi de titular (GIA 293). PURA.
function _CtInteressatDeLEmpresa($e, [string]$tipusTit, $xl) {
    if ($tipusTit -ne 'juridica' -or $null -ne (_CtV $e 'representant') -or [string](_CtV $e 'tipus') -eq 'autoritzacio') { return $false }
    $i = _CtV $e 'interessat'
    if ((_CtTipusNif ([string](_CtV $i 'nif'))) -ne 'fisica') { return $false }
    $em = _CtEmailNet ([string](_CtV $i 'email'))
    if ($em -eq '' -or (_CtEmailExclos $em)) { return $false }
    if ($em -eq (_CtEmailNet ([string](_CtV $xl 'EMAIL')))) { return $true }
    return ($Script:CtDominisComuns -notcontains (_CtDomini $em))
}
# ELS DOMINIS DEL TITULAR (l'usuari, 2a ronda: GIA 92, "@siemens.com"; GIA 154,
# "@antala.com"): una dada amb un d'aquests dominis NO es del tecnic encara que
# ell l'hagi posada. Compten: el domini dels correus de l'interessat-titular
# als documents, el del correu de l'Excel (si no es d'un tecnic: es l'error mes
# corrent de l'Excel) i qualsevol domini que porti una paraula del nom de
# l'empresa. Mai un proveidor corrent (gmail...) ni un professional. PURA.
function _CtDominisTitular($xl, $ord, [string]$nifTit, $ctx) {
    $out = @{}
    $bo = { param([string]$e)
        $d = _CtDomini $e
        if ($d -eq '' -or $Script:CtDominisComuns -contains $d -or (_CtEmailExclos $e) -or (_CtEmailProfessional $e)) { return '' }
        if ($ctx.Tecnics.Emails.ContainsKey($e) -or (_CtDominiTecnic $e $ctx) -ne '') { return '' }
        return $d
    }
    $paraules = @(_CtNomParaules ([string](_CtV $xl 'TITULAR')) | Where-Object { $_.Length -ge 4 })
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'
        if (_CtEsQueixa ([string](_CtV $e 'expedient')) ([string](_CtV $d 'ruta'))) { continue }
        foreach ($k in 'interessat', 'representant') {
            $q = _CtV $e $k
            $em = _CtEmailNet ([string](_CtV $q 'email'))
            $dd = & $bo $em
            if ($dd -eq '') { continue }
            $nq = _CtNifNet ([string](_CtV $q 'nif'))
            $etiq = ($dd.Split('.')[0])
            $porta = $false
            foreach ($w in $paraules) { if ($etiq.Contains($w)) { $porta = $true; break } }
            if ($porta -or ($k -eq 'interessat' -and $nifTit -ne '' -and $nq -eq $nifTit -and $null -eq (_CtV $e 'representant'))) { $out[$dd] = $true }
        }
    }
    $ex = _CtEmailNet ([string](_CtV $xl 'EMAIL'))
    $dx = & $bo $ex
    if ($dx -ne '' -and (_CtProuTitulars $ctx.EmailTitulars $ex $true) -eq 0) { $out[$dx] = $true }
    return $out
}

# Una dada amb el seu origen (regla 8).
function _CtAmbOrigen($p, [string]$font, [string]$data, [string]$conf) {
    $o = [ordered]@{}
    foreach ($k in 'nom', 'nif', 'email', 'telefon', 'mobil') { $o[$k] = [string](_CtV $p $k) }
    $o.font = $font; $o.data = $data; $o.confianca = $conf
    return $o
}

# La clau d'una persona per no repetir-la: el correu, el NIF o el nom.
function _CtClauPersona($p) {
    $e = _CtEmailNet ([string](_CtV $p 'email')); if ($e -ne '') { return 'e:' + $e }
    $n = _CtNifNet ([string](_CtV $p 'nif')); if ($n -ne '') { return 'n:' + $n }
    $m = _CtNomNorm ([string](_CtV $p 'nom')); if ($m -ne '') { return 'm:' + $m }
    return ''
}

# La data per ordenar un document: la seva o, si no en diu cap, la del fitxer.
function _CtDataDoc($d) {
    $x = [string](_CtV (_CtV $d 'extret') 'data')
    if ($x -ne '') { return $x }
    $m = Read-JsonIso (_CtV $d 'modificat')
    if ($m.Length -ge 10) { return $m.Substring(0, 10) }
    return ''
}

# Totes les claus d'una persona (correu, NIF i nom): una correccio a ma casa
# amb la persona encara que un document en porti el correu i l'altre no. PURA.
function _CtClausPersona($p) {
    $out = New-Object System.Collections.ArrayList
    $e = _CtEmailNet ([string](_CtV $p 'email')); if ($e -ne '') { [void]$out.Add('e:' + $e) }
    $n = _CtNifNet ([string](_CtV $p 'nif')); if ($n -ne '') { [void]$out.Add('n:' + $n) }
    $m = _CtNomNorm ([string](_CtV $p 'nom')); if ($m -ne '') { [void]$out.Add('m:' + $m) }
    return $out.ToArray()
}

function _CtCasaClaus($p, $claus) {
    if ($null -eq $p -or $null -eq $claus) { return $false }
    foreach ($k in @(_CtClausPersona $p)) { if (@($claus) -contains $k) { return $true } }
    return $false
}

# LES CORRECCIONS A MA (la finestra Contactes) passades a l'entrada de les
# regles. Com editat_a_ma a la base d'informes: MANEN sobre l'automatic, i el
# proxim "Actualitzar base" no les desfa perque es tornen a aplicar a cada
# passada. $corrs: la llista desada per al GIA. PURA.
#   es_tecnic     { claus }            aquesta persona es tecnic (autoritzada)
#   es_rep_legal  { claus }            aquesta persona es el representant legal
#   descarta      { avis }             aquest avis no torna a sortir
#   edita         { camp, valor }      camp = 'titular.email', 'representant_legal.nom'...
function _CtManDeCorreccions($corrs) {
    $man = @{ Tecnics = New-Object System.Collections.ArrayList; Rep = @(); RepPersona = $null; Descartats = @{}; Edits = New-Object System.Collections.ArrayList }
    foreach ($c in @($corrs)) {
        if ($null -eq $c) { continue }
        switch ([string](_CtV $c 'tipus')) {
            'es_tecnic'    { [void]$man.Tecnics.Add(@{ Claus = @(_CtV $c 'claus'); Persona = (_CtV $c 'persona') }) }
            'es_rep_legal' { $man.Rep = @(_CtV $c 'claus'); $man.RepPersona = (_CtV $c 'persona') }
            'descarta'     { $man.Descartats[[string](_CtV $c 'avis')] = $true }
            'edita'        { [void]$man.Edits.Add($c) }
        }
    }
    return $man
}

function _CtManEsTecnic($man, $p) {
    if ($null -eq $man) { return $false }
    foreach ($t in $man.Tecnics) { if (_CtCasaClaus $p $t.Claus) { return $true } }
    return $false
}

# $gia, els documents de l'activitat ($docs: @{ ruta; modificat; extret }), la
# fila de l'Excel ($xl, o $null), el context i les correccions a ma ($man, de
# _CtManDeCorreccions, o $null). Torna l'activitat. PURA.
function Get-ContactesActivitat([string]$gia, $docs, $xl, $ctx, $man = $null) {
    $act = [ordered]@{
        titular_tipus = ''; titular = $null; representant_legal = $null
        persones_autoritzades = @(); establiment = $null
        avisos = @(); documents = @(); historic = @(); canvi_titular = $null; notes = ''
        editat_a_ma = $false; correccions = @(); avisos_descartats = 0
    }
    $aplicades = New-Object System.Collections.ArrayList
    $ord = @(@($docs) | Where-Object { $null -ne (_CtV $_ 'extret') } | Sort-Object { _CtDataDoc $_ } -Descending)
    $act.documents = @($ord | ForEach-Object { [ordered]@{ ruta = [string](_CtV $_ 'ruta'); tipus = [string](_CtV (_CtV $_ 'extret') 'tipus'); data = (_CtDataDoc $_) } })
    if ($ctx.Recintes -contains $gia) {
        $act.notes = "Recinte amb organitzadors d'actes: els documents s" + [char]0x00F3 + "n dels organitzadors. No es proposa cap canvi de titular ni de representant legal."
        return $act
    }
    $nifXl = _CtNifNet ([string](_CtV $xl 'NIF'))
    # El NIF del titular segons els documents que NO son de queixa (per saber si
    # l'interessat d'una queixa es el titular).
    $nifTit = $nifXl
    if ($nifTit -eq '') {
        foreach ($d in $ord) {
            $e = _CtV $d 'extret'
            if (_CtEsQueixa ([string](_CtV $e 'expedient')) ([string](_CtV $d 'ruta'))) { continue }
            $n = _CtNifNet ([string](_CtV (_CtV $e 'interessat') 'nif'))
            if ($n -ne '') { $nifTit = $n; break }
        }
    }
    $tipusTit = _CtTipusNif $nifTit
    $titXl = [string](_CtV $xl 'TITULAR')
    # Les dades que NO son del tecnic encara que les porti (B3 de la 2a ronda):
    # el domini del titular, i les que l'interessat dona en un document SENSE
    # tecnic (cap representant).
    $dominisTit = _CtDominisTitular $xl $ord $nifTit $ctx
    $act.dominis_titular = @($dominisTit.Keys | Sort-Object)
    $dadesInt = @{}
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'
        if ($null -ne (_CtV $e 'representant') -or [string](_CtV $e 'tipus') -eq 'autoritzacio') { continue }
        if (_CtEsQueixa ([string](_CtV $e 'expedient')) ([string](_CtV $d 'ruta'))) { continue }
        $q = _CtV $e 'interessat'
        foreach ($k in 'email', 'telefon', 'mobil') { $v = [string](_CtV $q $k); if ($v -ne '') { $dadesInt[$v] = $true } }
    }

    # 1. LES PERSONES AUTORITZADES i els candidats a representant legal.
    # El correu que el tecnic ha posat i que no es seu (es del client): va al
    # titular si no en te cap altre.
    $emailsClient = New-Object System.Collections.ArrayList
    $aut = [ordered]@{}
    $candRep = New-Object System.Collections.ArrayList
    $afegeixAut = {
        param($p, [string]$rol, [string]$font, [string]$data, [string]$conf, [string]$motiu)
        # Una entitat de control (ECA, OCA) o una administracio no es mai
        # persona autoritzada de l'activitat.
        if (_CtEmailExclos ([string](_CtV $p 'email'))) { return }
        # El correu del domini del titular (i els telefons que porta amb ell) o
        # el que l'interessat dona sense tecnic no son seus: van al titular.
        $em0 = _CtEmailNet ([string](_CtV $p 'email'))
        $delTit = ($em0 -ne '' -and $dominisTit.ContainsKey((_CtDomini $em0)))
        $treu = @{}
        if ($delTit -or ($em0 -ne '' -and $dadesInt.ContainsKey($em0))) { $treu['email'] = $true }
        foreach ($c in 'telefon', 'mobil') { $v = [string](_CtV $p $c); if ($v -ne '' -and ($delTit -or $dadesInt.ContainsKey($v))) { $treu[$c] = $true } }
        if ($treu.Count -gt 0) {
            if ($treu.ContainsKey('email')) { [void]$emailsClient.Add(@{ Email = $em0; Font = $font; Data = $data }) }
            $p = _CtPersona ([string](_CtV $p 'nom')) ([string](_CtV $p 'nif')) $(if ($treu.ContainsKey('email')) { '' } else { $em0 }) $(if ($treu.ContainsKey('telefon')) { '' } else { [string](_CtV $p 'telefon') }) $(if ($treu.ContainsKey('mobil')) { '' } else { [string](_CtV $p 'mobil') })
        }
        $k = _CtClauPersona $p
        if ($k -eq '') { return }
        if ($aut.Contains($k)) {
            $o = $aut[$k]
            foreach ($c in 'nom', 'nif', 'email', 'telefon', 'mobil') { if ([string]$o[$c] -eq '' -and [string](_CtV $p $c) -ne '') { $o[$c] = [string](_CtV $p $c) } }
            if ($conf -eq 'segur') { $o.confianca = 'segur' }
            return
        }
        $o = _CtAmbOrigen $p $font $data $conf
        $o.Insert(1, 'rol', $rol)
        $emp = ''
        $e = _CtEmailNet ([string](_CtV $p 'email'))
        if ($e -ne '' -and $ctx.Tecnics.Emails.ContainsKey($e)) { $emp = (@($ctx.Tecnics.Emails[$e].empresa) | Select-Object -First 1) }
        $o.Insert(2, 'empresa', [string]$emp)
        $o.motiu = $motiu
        $aut[$k] = $o
    }
    # B2: el correu de la llista que porta una altra persona es de la gestoria
    # (persona autoritzada); la persona, sense aquell correu, conserva el seu
    # paper. Torna la persona sense el correu (o la mateixa).
    $treuCorreuAltri = {
        param($p, [string]$font, [string]$data)
        $ent = _CtCorreuTecnicAltri $p $ctx
        if ($null -eq $ent) { return $p }
        $em = _CtEmailNet ([string](_CtV $p 'email'))
        $nomG = [string](@(@($ent.noms) | Where-Object { [string]$_ -ne '' }) | Select-Object -First 1)
        $rolG = if ([string]$ent.rol -ne '') { [string]$ent.rol } else { 'gestoria' }
        & $afegeixAut (_CtPersona $nomG '' $em) $rolG $font $data 'segur' ("el correu es de la llista de t" + [char]0x00E8 + "cnics coneguts (" + $(if (@($ent.empresa).Count -gt 0) { [string]@($ent.empresa)[0] } else { $nomG }) + "), no d'aquesta persona")
        return (_CtPersona ([string](_CtV $p 'nom')) ([string](_CtV $p 'nif')) '' ([string](_CtV $p 'telefon')) ([string](_CtV $p 'mobil')))
    }
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'; $font = [string](_CtV $d 'ruta'); $data = _CtDataDoc $d
        $queixa = _CtEsQueixa ([string](_CtV $e 'expedient')) $font
        # Un document atribuit pel NIF (C2) i no per la carpeta: tot "probable".
        $docProb = [bool](_CtV $d 'probable')
        # LA SIGNATURA DE REPRESENTANT (R:CIF), la font mes segura: del CIF del
        # titular, el representant legal; d'una altra entitat (una enginyeria,
        # una gestoria), persona autoritzada. Les de les administracions
        # (P, Q, S: l'Ajuntament que signa un requeriment) no son de ningu.
        foreach ($sg in @(_CtV $e 'signatures')) {
            if ($null -eq $sg) { continue }
            $cif = _CtNifNet ([string](_CtV $sg 'cif'))
            $pS = _CtPersonaSignant $sg $e
            if ($nifTit -ne '' -and $cif -eq $nifTit) { $pS = & $treuCorreuAltri $pS $font $data; [void]$candRep.Add(@{ P = $pS; Font = $font; Data = $data; Conf = $(if ($docProb) { 'probable' } else { 'segur' }); Via = 'signatura' }); continue }
            if ($queixa -or $cif -match '^[PQS]' -or $nifTit -eq '') { continue }
            & $afegeixAut $pS ("t" + [char]0x00E8 + "cnic") $font $data 'segur' ("signa com a representant d'una altra entitat (R:" + $cif + ")")
        }
        $a = _CtV $e 'autoritzacio'
        if ($null -ne $a -and -not $queixa) {
            foreach ($p in @(_CtV $a 'autoritzats')) { if ($null -ne $p) { & $afegeixAut $p 'autoritzat' $font $data 'segur' "l'autoritzaci$([char]0x00F3) l'autoritza" } }
            $sig = _CtV $a 'signant'
            if ($null -ne $sig -and $null -ne (_CtV $a 'empresa')) { $sig = & $treuCorreuAltri $sig $font $data; [void]$candRep.Add(@{ P = $sig; Font = $font; Data = $data; Conf = $(if ($docProb) { 'probable' } else { 'segur' }); Via = 'autoritzacio' }) }
        }
        # L'interessat persona fisica d'una entitat (A2 de la 2a ronda).
        if (-not $queixa -and (_CtInteressatRepCandidat $e $tipusTit)) {
            $int = & $treuCorreuAltri (_CtV $e 'interessat') $font $data
            if ((_CtMotiuTecnic $int $ctx) -eq '' -and -not (_CtManEsTecnic $man $int)) { [void]$candRep.Add(@{ P = $int; Font = $font; Data = $data; Conf = 'probable'; Via = 'interessat' }) }
        }
        $rep = _CtV $e 'representant'
        if ($null -ne $rep -and -not $queixa) { $rep = & $treuCorreuAltri $rep $font $data }
        if ($null -ne $rep -and -not $queixa) {
            $motiu = _CtMotiuTecnic $rep $ctx
            $nifRep = _CtNifNet ([string](_CtV $rep 'nif'))
            if ($motiu -eq '' -and (_CtTipusNif $nifRep) -eq 'juridica' -and $nifRep -ne $nifTit) { $motiu = "presenta amb un NIF d'entitat (gestoria)" }
            if (_CtManEsTecnic $man $rep) { $motiu = "marcat a m$([char]0x00E0) com a t$([char]0x00E8)cnic" }
            if ($null -ne $man -and (_CtCasaClaus $rep $man.Rep)) { $motiu = ''; [void]$candRep.Add(@{ P = $rep; Font = $font; Data = $data; Conf = 'segur'; Via = 'tramit' }); continue }
            # Tecnic pel DNI o pel nom: el correu que ha posat, si no te res de
            # tecnic, es del client (l'usuari, GIA 366).
            $pAut = $rep
            if ($motiu -ne '' -and $motiu -eq (_CtMotiuTecnicIdentitat $rep $ctx)) {
                $em = _CtEmailNet ([string](_CtV $rep 'email'))
                if ($em -ne '' -and (_CtMotiuTecnic @{ email = $em } $ctx) -eq '') {
                    $pAut = _CtPersona ([string](_CtV $rep 'nom')) $nifRep '' ([string](_CtV $rep 'telefon')) ([string](_CtV $rep 'mobil'))
                    [void]$emailsClient.Add(@{ Email = $em; Font = $font; Data = $data })
                }
            }
            if ($motiu -ne '') { & $afegeixAut $pAut $(if ($motiu -match 'gestoria') { 'gestoria' } else { "t$([char]0x00E8)cnic" }) $font $data 'segur' $motiu }
            elseif ($tipusTit -ne 'juridica') {
                # Persona fisica: qui presenta per ella es una persona autoritzada
                # (el representant legal d'una fisica es ella mateixa). Si es
                # ella mateixa, no.
                $int = _CtV $e 'interessat'
                $mateix = ($nifRep -ne '' -and $nifRep -eq $nifTit) -or (_CtMateixNom ([string](_CtV $rep 'nom')) ([string](_CtV $int 'nom')))
                if (-not $mateix) { & $afegeixAut $rep ('representant del tr' + [char]0x00E0 + 'mit') $font $data 'probable' 'presenta en nom del titular' }
            }
            else { [void]$candRep.Add(@{ P = $rep; Font = $font; Data = $data; Conf = 'probable'; Via = 'tramit' }) }
        }
    }
    # Les correccions a ma de persones: el marcat com a tecnic passa a
    # autoritzat (encara que fos el signant d'una autoritzacio) i el marcat com
    # a representant legal surt dels autoritzats.
    $repMa = $null
    if ($null -ne $man) {
        foreach ($c in @($candRep)) {
            if (_CtManEsTecnic $man $c.P) {
                & $afegeixAut $c.P ("t" + [char]0x00E8 + "cnic") $c.Font $c.Data 'segur' ("marcat a m" + [char]0x00E0 + " com a t" + [char]0x00E8 + "cnic")
                $candRep.Remove($c)
                [void]$aplicades.Add([ordered]@{ tipus = 'es_tecnic'; text = ([string](_CtV $c.P 'nom') + " es t" + [char]0x00E8 + "cnic"); auto = 'candidat a representant legal' })
            }
        }
        foreach ($k in @($aut.Keys)) {
            if (_CtManEsTecnic $man $aut[$k]) { [void]$aplicades.Add([ordered]@{ tipus = 'es_tecnic'; text = ([string]$aut[$k].nom + " es t" + [char]0x00E8 + "cnic"); auto = [string]$aut[$k].rol }) }
        }
        # Marcat com a tecnic pero no surt a cap document (p.ex. el
        # representant legal que diu l'Excel): s'afegeix tal com es va marcar,
        # i aixi l'avis es_el_tecnic de l'Excel el troba.
        foreach ($t in $man.Tecnics) {
            $hi = $false
            foreach ($o in $aut.Values) { if (_CtCasaClaus $o $t.Claus) { $hi = $true; break } }
            if ($hi -or $null -eq $t.Persona) { continue }
            & $afegeixAut $t.Persona ("t" + [char]0x00E8 + "cnic") '' '' 'segur' ("marcat a m" + [char]0x00E0 + " com a t" + [char]0x00E8 + "cnic")
            [void]$aplicades.Add([ordered]@{ tipus = 'es_tecnic'; text = ([string](_CtV $t.Persona 'nom') + " es t" + [char]0x00E8 + "cnic"); auto = '(no surt als documents)' })
        }
        if (@($man.Rep).Count -gt 0) {
            foreach ($k in @($aut.Keys)) {
                if (_CtCasaClaus $aut[$k] $man.Rep) {
                    $repMa = @{ P = $aut[$k]; Font = [string]$aut[$k].font; Data = [string]$aut[$k].data; Auto = ('persona autoritzada (' + [string]$aut[$k].rol + ')') }
                    $aut.Remove($k)
                }
            }
            if ($null -eq $repMa) {
                foreach ($c in @($candRep)) { if (_CtCasaClaus $c.P $man.Rep) { $repMa = @{ P = $c.P; Font = $c.Font; Data = $c.Data; Auto = 'candidat a representant legal' }; break } }
            }
            if ($null -eq $repMa) {
                # L'interessat d'algun document (p.ex. la persona fisica que
                # presenta per una empresa).
                foreach ($d in $ord) {
                    $p = _CtV (_CtV $d 'extret') 'interessat'
                    if ($null -ne $p -and (_CtCasaClaus $p $man.Rep)) { $repMa = @{ P = $p; Font = [string](_CtV $d 'ruta'); Data = (_CtDataDoc $d); Auto = 'interessat' }; break }
                }
            }
            # No surt a cap document (el que diu l'Excel): tal com es va marcar.
            if ($null -eq $repMa -and $null -ne $man.RepPersona) { $repMa = @{ P = $man.RepPersona; Font = ''; Data = ''; Auto = '(no surt als documents)' } }
        }
    }
    # EL REPRESENTANT DE LA SIGNATURA (R:CIF del titular) mana sobre les
    # regles: nomes el treu una correccio a ma (es_tecnic, mes amunt). Si
    # tambe sortia com a autoritzat (el DNI repetit), en surt.
    $repSig = $null
    if ($null -eq $repMa) {
        foreach ($c in @($candRep | Where-Object { [string]$_.Via -eq 'signatura' } | Sort-Object { [string]$_.Data } -Descending)) { $repSig = $c; break }
        if ($null -ne $repSig) {
            foreach ($k in @($aut.Keys)) {
                $o = $aut[$k]
                if (([string]$repSig.P.nif -ne '' -and [string]$o.nif -eq [string]$repSig.P.nif) -or (_CtMateixNom ([string]$o.nom) ([string]$repSig.P.nom))) { $aut.Remove($k) }
            }
        }
    }
    # Un candidat a representant que resulta que es tecnic (surt com a
    # autoritzat en un altre document, o pel context) passa a autoritzat.
    $tecEmails = @{}; $tecTels = @{}; $tecNifs = @{}
    $refesTec = {
        $tecEmails.Clear(); $tecTels.Clear(); $tecNifs.Clear()
        foreach ($o in $aut.Values) {
            if ([string]$o.email -ne '') { $tecEmails[[string]$o.email] = $o }
            foreach ($c in 'telefon', 'mobil') { if ([string]$o[$c] -ne '') { $tecTels[[string]$o[$c]] = $o } }
            if ([string]$o.nif -ne '') { $tecNifs[[string]$o.nif] = $o }
        }
    }
    & $refesTec
    $esDeTecnic = {
        param($p)
        $e = _CtEmailNet ([string](_CtV $p 'email')); if ($e -ne '' -and $tecEmails.ContainsKey($e)) { return $true }
        $n = _CtNifNet ([string](_CtV $p 'nif')); if ($n -ne '' -and $tecNifs.ContainsKey($n)) { return $true }
        foreach ($o in $aut.Values) { if (_CtMateixNom ([string](_CtV $p 'nom')) ([string]$o.nom)) { return $true } }
        return $false
    }
    # Un interessat que es el tecnic (pel NIF: el de l'autoritzat, el que
    # presenta per a titulars diferents o el marcat a ma).
    $interessatTecnic = {
        param($p)
        $n = _CtNifNet ([string](_CtV $p 'nif'))
        if ($n -ne '' -and $tecNifs.ContainsKey($n)) { return $true }
        if ((_CtMotiuTecnicIdentitat @{ nif = $n } $ctx) -ne '') { return $true }
        return (_CtManEsTecnic $man $p)
    }

    # EL CANVI DE TITULAR (nomes avis: l'Excel pot ser mes nou): el NIF de
    # l'interessat del document no-queixa MES RECENT es d'una altra entitat o
    # persona que no es el tecnic. No ho es: el mateix nom de l'Excel (l'Excel
    # te el NIF malament), qui signa amb el certificat del titular, ni la
    # persona fisica que presenta per l'empresa.
    $act.canvi_titular = $null
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'
        if (_CtEsQueixa ([string](_CtV $e 'expedient')) ([string](_CtV $d 'ruta'))) { continue }
        if ([string](_CtV $e 'tipus') -eq 'autoritzacio') { continue }
        $i = _CtV $e 'interessat'
        $n = _CtNifNet ([string](_CtV $i 'nif'))
        if ($n -eq '') { continue }
        if ($nifXl -eq '' -or $n -eq $nifXl) { break }
        if (& $interessatTecnic $i) { continue }
        if (_CtMateixNom ([string](_CtV $i 'nom')) $titXl) { break }
        if ($null -ne $repSig -and $n -eq [string]$repSig.P.nif) { break }
        if (_CtInteressatDeLEmpresa $e $tipusTit $xl) { break }
        $act.canvi_titular = [ordered]@{ nom = [string](_CtV $i 'nom'); nif = $n; font = [string](_CtV $d 'ruta'); data = (_CtDataDoc $d); confianca = 'segur' }
        break
    }

    # 2. EL TITULAR (regles 3, 5 i 7): camp a camp, el mes recent que no sigui
    #    del tecnic ni d'una administracio.
    $tit = [ordered]@{ nom = ''; nif = ''; email = ''; telefon = ''; mobil = ''; font = ''; data = ''; confianca = '' }
    $hist = New-Object System.Collections.ArrayList
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'; $font = [string](_CtV $d 'ruta'); $data = _CtDataDoc $d
        $p = _CtV $e 'interessat'
        if ($null -eq $p) { continue }
        $nifP = _CtNifNet ([string](_CtV $p 'nif'))
        if ((_CtEsQueixa ([string](_CtV $e 'expedient')) $font) -and ($nifP -eq '' -or $nifP -ne $nifTit)) { continue }
        # Un "interessat" que es el tecnic (el tecnic que es posa ell mateix).
        if ($nifP -ne '' -and $nifP -ne $nifXl -and (& $interessatTecnic $p)) { continue }
        # Un interessat amb un NIF que no es el de l'Excel (i un altre nom): el
        # nom i el NIF no hi van (es el canvi_titular dels avisos), pero el
        # correu i el telefon SI son del titular, "probable" (l'usuari: "el
        # contacte de l'interessat nomes es feia servir quan el NIF casava").
        $altre = ($nifXl -ne '' -and $nifP -ne '' -and $nifP -ne $nifXl -and -not (_CtMateixNom ([string](_CtV $p 'nom')) $titXl))
        $conf = if ([string](_CtV $e 'tipus') -eq 'autoritzacio' -or $altre -or [bool](_CtV $d 'probable')) { 'probable' } else { 'segur' }
        $posat = $false
        foreach ($c in 'nom', 'nif', 'email', 'telefon', 'mobil') {
            $v = [string](_CtV $p $c)
            if ($v -eq '') { continue }
            if ($c -eq 'email' -and ((_CtEmailExclos $v) -or $tecEmails.ContainsKey($v) -or (_CtMotiuTecnic @{ email = $v } $ctx -Fort) -ne '')) { continue }
            # B5: el correu d'una assessoria pel nom (i el titular no ho es).
            if ($c -eq 'email' -and (_CtCorreuGestoriaPelNom $v ($titXl + ' ' + [string](_CtV $p 'nom')))) { continue }
            if (($c -eq 'telefon' -or $c -eq 'mobil') -and ($tecTels.ContainsKey($v) -or (_CtMotiuTecnic @{ $c = $v } $ctx -Fort) -ne '')) { continue }
            if ($altre -and ($c -eq 'nom' -or $c -eq 'nif')) {
                [void]$hist.Add([ordered]@{ camp = $c; valor = $v; font = $font; data = $data; nota = 'un altre titular' })
                continue
            }
            if ([string]$tit[$c] -eq '') {
                $tit[$c] = $v; $posat = $true
            } elseif ([string]$tit[$c] -ne $v) {
                [void]$hist.Add([ordered]@{ camp = $c; valor = $v; font = $font; data = $data })
            }
        }
        if ($posat) {
            if ($tit.font -eq '') { $tit.font = $font; $tit.data = $data; $tit.confianca = $conf }
            elseif ($altre) { $tit.confianca = 'probable' }
        }
    }
    # El correu del client que havia posat el tecnic.
    if ([string]$tit.email -eq '') {
        foreach ($ce in $emailsClient) {
            if ((_CtEmailExclos $ce.Email) -or $tecEmails.ContainsKey($ce.Email) -or (_CtCorreuGestoriaPelNom $ce.Email $titXl)) { continue }
            $tit.email = $ce.Email
            if ($tit.font -eq '') { $tit.font = $ce.Font; $tit.data = $ce.Data }
            $tit.confianca = 'probable'
            break
        }
    }
    if ($tit.font -ne '') { $act.titular = $tit }
    # Sense nom (o amb una etiqueta per nom): la Rao social de l'Excel.
    if ($null -ne $act.titular -and $titXl -ne '' -and ([string]$act.titular.nom -eq '' -or (_CtNomEsEtiqueta ([string]$act.titular.nom)))) { $act.titular.nom = $titXl }
    $tipusFinal = _CtTipusNif $(if ($nifXl -ne '') { $nifXl } else { [string]$tit.nif })
    $act.titular_tipus = $tipusFinal
    $act.historic = $hist.ToArray()

    # 3. EL REPRESENTANT LEGAL (regla 2). El de la correccio a ma, mana.
    if ($null -ne $repMa) {
        $act.representant_legal = _CtAmbOrigen $repMa.P $repMa.Font $repMa.Data ("a m" + [char]0x00E0)
        $act.representant_legal.via = ("a m" + [char]0x00E0)
        [void]$aplicades.Add([ordered]@{ tipus = 'es_rep_legal'; text = ([string](_CtV $repMa.P 'nom') + ' es el representant legal'); auto = [string]$repMa.Auto })
    } elseif ($null -ne $repSig) {
        $r = _CtAmbOrigen $repSig.P $repSig.Font $repSig.Data 'segur'
        if ($tecEmails.ContainsKey([string]$r.email) -or (_CtEmailExclos ([string]$r.email))) { $r.email = '' }
        $r.via = 'signatura'
        $act.representant_legal = $r
    } elseif ($tipusFinal -eq 'juridica') {
        foreach ($c in @($candRep | Sort-Object @{ Expression = { if ($_.Conf -eq 'segur') { 0 } else { 1 } } }, @{ Expression = { $_.Data }; Descending = $true })) {
            if ((_CtMotiuTecnic $c.P $ctx) -ne '' -or (& $esDeTecnic $c.P)) { continue }
            if (_CtEmailExclos ([string](_CtV $c.P 'email'))) { continue }
            $r = _CtAmbOrigen $c.P $c.Font $c.Data $c.Conf
            $r.via = [string]$c.Via
            # Ni el correu ni el telefon del tecnic.
            if ($tecEmails.ContainsKey([string]$r.email)) { $r.email = '' }
            foreach ($k in 'telefon', 'mobil') { if ($tecTels.ContainsKey([string]$r[$k])) { $r[$k] = '' } }
            $act.representant_legal = $r
            break
        }
    }

    # 4. L'ESTABLIMENT (regla 4): el de l'e-TRAM mes recent.
    foreach ($d in $ord) {
        $es = _CtV (_CtV $d 'extret') 'establiment'
        if ($null -eq $es) { continue }
        $o = [ordered]@{ nom_comercial = [string](_CtV $es 'nom_comercial'); telefon = [string](_CtV $es 'telefon'); email = [string](_CtV $es 'email')
                         font = [string](_CtV $d 'ruta'); data = (_CtDataDoc $d); confianca = 'segur' }
        if ($tecEmails.ContainsKey($o.email) -or (_CtEmailExclos $o.email)) { $o.email = '' }
        if ($tecTels.ContainsKey($o.telefon)) { $o.telefon = '' }
        if ($o.nom_comercial -ne '' -or $o.email -ne '' -or $o.telefon -ne '') { $act.establiment = $o; break }
    }

    $act.persones_autoritzades = @($aut.Values)

    # Les dades corregides a ma (Edita), abans dels avisos: l'Excel es compara
    # amb el valor bo. Es guarda l'automatic per poder-ho desfer.
    if ($null -ne $man) {
        foreach ($ed in $man.Edits) {
            $parts = ([string](_CtV $ed 'camp')).Split('.')
            if ($parts.Count -ne 2) { continue }
            $qui = $parts[0]; $dada = $parts[1]
            if (@('titular', 'representant_legal', 'establiment') -notcontains $qui) { continue }
            if ($null -eq $act[$qui]) {
                $act[$qui] = if ($qui -eq 'establiment') { [ordered]@{ nom_comercial = ''; telefon = ''; email = ''; font = ''; data = ''; confianca = '' } }
                             else { [ordered]@{ nom = ''; nif = ''; email = ''; telefon = ''; mobil = ''; font = ''; data = ''; confianca = '' } }
            }
            $auto = [string]$act[$qui][$dada]
            $act[$qui][$dada] = [string](_CtV $ed 'valor')
            $act[$qui].confianca = "a m" + [char]0x00E0
            [void]$aplicades.Add([ordered]@{ tipus = 'edita'; camp = [string](_CtV $ed 'camp'); text = ([string](_CtV $ed 'camp') + ' = ' + [string](_CtV $ed 'valor')); auto = $auto })
        }
        if ($null -ne $act.titular -and [string]$act.titular.nif -ne '') { $act.titular_tipus = _CtTipusNif ([string]$act.titular.nif) }
    }

    $avisos = @(_CtAvisos $act $xl $ctx)
    if ($null -ne $man -and $man.Descartats.Count -gt 0) {
        $act.avisos_descartats = @($avisos | Where-Object { $man.Descartats.ContainsKey([string]$_.id) }).Count
        $avisos = @($avisos | Where-Object { -not $man.Descartats.ContainsKey([string]$_.id) })
    }
    $act.avisos = $avisos
    $act.correccions = $aplicades.ToArray()
    $act.editat_a_ma = ($aplicades.Count -gt 0 -or [int]$act.avisos_descartats -gt 0)
    return $act
}

# ----------------------------------------------------------------------------
# ELS AVISOS CONTRA L'EXCEL
# ----------------------------------------------------------------------------
# Els camps de l'Excel que es comparen: el nom de la columna (com surt a la
# correccio), la clau de la fila de contactes, de qui es i quin camp.
$Script:CtCampsExcel = @(
    @{ Camp = ('Ra' + [char]0x00F3 + ' social');            Xl = 'TITULAR';     Qui = 'titular';            Dada = 'nom' }
    @{ Camp = 'NIF';                                          Xl = 'NIF';         Qui = 'titular';            Dada = 'nif' }
    @{ Camp = ('Tel' + [char]0x00E8 + 'fon');               Xl = 'TELEFON';     Qui = 'titular';            Dada = 'telefon' }
    @{ Camp = ('M' + [char]0x00F2 + 'bil');                 Xl = 'MOBIL';       Qui = 'titular';            Dada = 'mobil' }
    @{ Camp = 'E-mail';                                       Xl = 'EMAIL';       Qui = 'titular';            Dada = 'email' }
    @{ Camp = 'Representant legal';                           Xl = 'REP_NOM';     Qui = 'representant_legal'; Dada = 'nom' }
    @{ Camp = 'Rep. Leg. NIF';                                Xl = 'REP_NIF';     Qui = 'representant_legal'; Dada = 'nif' }
    @{ Camp = ('Rep. Leg. Tel' + [char]0x00E8 + 'fon');     Xl = 'REP_TELEFON'; Qui = 'representant_legal'; Dada = 'telefon' }
    @{ Camp = ('Rep. Leg. M' + [char]0x00F2 + 'bil');       Xl = 'REP_MOBIL';   Qui = 'representant_legal'; Dada = 'mobil' }
    @{ Camp = 'Rep. Leg. E-mail';                             Xl = 'REP_EMAIL';   Qui = 'representant_legal'; Dada = 'email' }
)

# Els proveidors de correu mes corrents (per detectar-ne les errades).
$Script:CtDominisComuns = @('gmail.com', 'hotmail.com', 'hotmail.es', 'yahoo.es', 'yahoo.com', 'outlook.com', 'outlook.es', 'live.com', 'icloud.com', 'telefonica.net', 'msn.com')
$Script:CtDominisErronis = @('gmailo.com', 'gmial.com', 'gmal.com', 'gmai.com', 'gmail.co', 'gmail.es', 'gmail.cat', 'hotmal.com', 'holmail.com', 'hotmial.com', 'hotmai.com', 'hotmil.com', 'hotmail.co', 'hotamil.com', 'yahooo.es', 'yaho.es', 'outlok.com', 'outlook.co')

# Distancia d'edicio (Levenshtein). PURA.
function _CtDistancia([string]$a, [string]$b) {
    $n = $a.Length; $m = $b.Length
    if ($n -eq 0) { return $m }; if ($m -eq 0) { return $n }
    $ant = New-Object 'int[]' ($m + 1); $act = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $ant[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $act[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $c = if ($a[$i - 1] -eq $b[$j - 1]) { 0 } else { 1 }
            $act[$j] = [Math]::Min([Math]::Min($act[$j - 1] + 1, $ant[$j] + 1), $ant[$j - 1] + $c)
        }
        $t = $ant; $ant = $act; $act = $t
    }
    return $ant[$m]
}

# La lletra del DNI/NIE quadra? PURA.
function _CtLletraDniBona([string]$nif) {
    $t = _CtNifNet $nif
    if ((_CtTipusNif $t) -ne 'fisica') { return $true }
    $num = $t.Substring(0, 8) -replace '^X', '0' -replace '^Y', '1' -replace '^Z', '2'
    $lletres = 'TRWAGMYFPDXBNJZSQVHLCKE'
    return ($lletres[[int64]$num % 23] -eq $t[8])
}

# Una errada de tecleig o de columna en un valor de l'Excel, o ''. $doc: el
# valor dels documents (per comparar-hi el domini). PURA.
function _CtErrada([string]$dada, [string]$valor, [string]$doc, [string]$nifTitular = '') {
    $v = ([string]$valor).Trim()
    if ($v -eq '') { return '' }
    switch ($dada) {
        'email' {
            $e = _CtEmailNet $v
            if ($e -eq '') { return "no t$([char]0x00E9) forma d'adre" + [char]0x00E7 + "a de correu" }
            $d = _CtDomini $e
            if ($Script:CtDominisErronis -contains $d) { return "domini mal escrit ($d)" }
            $dd = _CtDomini $doc
            $local = $e.Substring(0, $e.IndexOf('@'))
            if ($dd -ne '' -and $dd -ne $d -and $local -eq (_CtEmailNet $doc).Split('@')[0] -and (_CtDistancia $d $dd) -le 2) { return "domini mal escrit ($d en lloc de $dd)" }
            foreach ($c in $Script:CtDominisComuns) { if ($d -ne $c -and (_CtDistancia $d $c) -eq 1 -and $Script:CtDominisComuns -notcontains $d) { return "domini mal escrit ($d en lloc de $c)" } }
            return ''
        }
        { $_ -eq 'telefon' -or $_ -eq 'mobil' } {
            if ($v.Contains('@')) { return "hi ha un correu a la columna del tel$([char]0x00E8)fon" }
            $t = _CtTelefonNet $v
            if ($t.Length -eq 10) { return "el tel$([char]0x00E8)fon t$([char]0x00E9) 10 xifres" }
            if ($t.Length -ne 9) { return "el tel$([char]0x00E8)fon no t$([char]0x00E9) 9 xifres" }
            if ($dada -eq 'mobil' -and (_CtTipusTelefon $t) -eq 'fix') { return "un fix a la columna del m$([char]0x00F2)bil" }
            return ''
        }
        'nif' {
            $t = _CtNifNet $v
            if ((_CtTipusNif $t) -eq '') { return "no t$([char]0x00E9) forma de NIF" }
            if (-not (_CtLletraDniBona $t)) { return 'la lletra del DNI no quadra' }
            return ''
        }
    }
    return ''
}

function _CtIgual([string]$dada, [string]$a, [string]$b) {
    switch ($dada) {
        'email'   { return ((_CtEmailNet $a) -eq (_CtEmailNet $b)) }
        'nif'     { return ((_CtNifNet $a) -eq (_CtNifNet $b)) }
        'nom'     { return (_CtMateixNom $a $b) }
        default   { return ((_CtTelefonNet $a) -eq (_CtTelefonNet $b)) }
    }
}

# L'identificador estable d'un avis (per poder-lo descartar i que no torni).
function _CtAvisId([string]$tipus, [string]$camp, [string]$valor) {
    return ($tipus + '|' + (_CtPla $camp) + '|' + (_CtPla ([string]$valor).Trim()))
}

function _CtAvis([string]$tipus, [string]$camp, [string]$valor, [string]$problema, [string]$proposta, $origen) {
    return [ordered]@{
        id = (_CtAvisId $tipus $camp $valor); tipus = $tipus; camp = $camp; valor_excel = $valor
        problema = $problema; proposta = $proposta
        font = [string](_CtV $origen 'font'); data = [string](_CtV $origen 'data'); confianca = [string](_CtV $origen 'confianca')
    }
}

# A quina persona autoritzada correspon un valor de l'Excel, o $null. PURA.
function _CtAutoritzatDe($act, [string]$dada, [string]$valor) {
    foreach ($o in @($act.persones_autoritzades)) {
        if ($dada -eq 'nom') { if (_CtMateixNom $valor ([string]$o.nom)) { return $o } continue }
        if ($dada -eq 'telefon' -or $dada -eq 'mobil') {
            $t = _CtTelefonNet $valor
            if ($t -ne '' -and ($t -eq [string]$o.telefon -or $t -eq [string]$o.mobil)) { return $o }
            continue
        }
        if ($valor -ne '' -and (_CtIgual $dada $valor ([string]$o[$dada]))) { return $o }
    }
    return $null
}

function _CtAvisos($act, $xl, $ctx) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $xl) { return $out.ToArray() }
    $nifTitXl = _CtNifNet ([string](_CtV $xl 'NIF'))
    $tit = $act.titular
    # canvi_titular: els documents son d'un altre NIF (nomes avis: l'Excel pot
    # ser mes nou que la carpeta).
    $ct = $act.canvi_titular
    $canvi = ($null -ne $ct)
    if ($canvi) {
        $nouNom = [string]$ct.nom
        [void]$out.Add((_CtAvis 'canvi_titular' 'NIF' $nifTitXl ("els documents m$([char]0x00E9)s recents s$([char]0x00F3)n d'un altre titular: " + $(if ($nouNom -ne '') { $nouNom + ' (' + [string]$ct.nif + ')' } else { [string]$ct.nif })) ([string]$ct.nif) $ct))
    }
    foreach ($c in $Script:CtCampsExcel) {
        $vXl = ([string](_CtV $xl $c.Xl)).Trim()
        $p = $act[$c.Qui]
        $vDoc = [string](_CtV $p $c.Dada)
        # 1. es_el_tecnic
        if ($vXl -ne '') {
            $o = _CtAutoritzatDe $act $c.Dada $vXl
            $mot = ''
            # (B3: el correu del domini del titular no es mai del tecnic.)
            if ($null -eq $o -and $c.Dada -eq 'email' -and @($act.dominis_titular) -notcontains (_CtDomini $vXl)) { $mot = _CtMotiuTecnic @{ email = $vXl } $ctx -Fort }
            if ($null -eq $o -and ($c.Dada -eq 'telefon' -or $c.Dada -eq 'mobil')) { $mot = _CtMotiuTecnic @{ mobil = (_CtTelefonNet $vXl) } $ctx -Fort }
            if ($null -eq $o -and $c.Dada -eq 'nom' -and $c.Qui -eq 'representant_legal') {
                foreach ($n in @($ctx.Tecnics.Noms)) { if (_CtMateixNom $vXl $n) { $mot = "es a la llista de t$([char]0x00E8)cnics coneguts"; break } }
            }
            if ($null -ne $o -or $mot -ne '') {
                $qui = if ($null -ne $o) { [string]$o.nom + ' (' + [string]$o.rol + ')' } else { $mot }
                if ($qui -match '^\s*\(') { $qui = [string]$o.email + ' (' + [string]$o.rol + ')' }
                $prop = $vDoc
                if ($c.Qui -eq 'representant_legal' -and $act.titular_tipus -eq 'fisica') { $prop = '' }
                $txt = "$([char]0x00E9)s del t$([char]0x00E8)cnic o la gestoria: $qui"
                if ($prop -eq '') { $txt += $(if ($c.Qui -eq 'representant_legal' -and $act.titular_tipus -eq 'fisica') { ". Persona f$([char]0x00ED)sica: el representant legal es deixa en blanc" } else { '. Cal deixar-lo en blanc o posar el del titular' }) }
                $orig = if ($null -ne $o) { $o } else { $p }
                [void]$out.Add((_CtAvis 'es_el_tecnic' $c.Camp $vXl $txt $prop $orig))
                continue
            }
        }
        # Un representant legal a l'Excel d'una persona fisica.
        if ($c.Qui -eq 'representant_legal' -and $act.titular_tipus -eq 'fisica') { continue }
        # 2. error
        $err = ''
        if ($vXl -ne '') { $err = _CtErrada $c.Dada $vXl $vDoc $nifTitXl }
        if ($err -eq '' -and $c.Xl -eq 'REP_NIF' -and $vXl -ne '') {
            $rn = _CtNifNet $vXl
            if ($rn -eq $nifTitXl) { $err = "$([char]0x00E9)s el NIF de l'empresa, no el del representant" }
            elseif ((_CtTipusNif $rn) -eq 'juridica') { $err = "$([char]0x00E9)s un NIF d'entitat, no el d'una persona" }
        }
        if ($err -ne '') { [void]$out.Add((_CtAvis 'error' $c.Camp $vXl $err $vDoc $p)); continue }
        if ($vDoc -eq '') { continue }
        if ($c.Qui -eq 'titular' -and ($c.Dada -eq 'nom' -or $c.Dada -eq 'nif') -and $canvi) { continue }
        # 3. falta / 4. diferent
        if ($vXl -eq '') { [void]$out.Add((_CtAvis 'falta' $c.Camp '' ("l'Excel " + [char]0x00E9 + "s buit i els documents donen la dada") $vDoc $p)); continue }
        if (-not (_CtIgual $c.Dada $vXl $vDoc)) {
            [void]$out.Add((_CtAvis 'diferent' $c.Camp $vXl "els documents m$([char]0x00E9)s recents diuen una altra cosa" $vDoc $p))
        }
    }
    return $out.ToArray()
}
