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
$Script:CtEntitatsControl = @('aucatel', 'tuvsud', 'tuv-sud', 'sgs.com', 'dekra', 'bureauveritas', 'bureau-veritas', 'applus', 'ocacert', 'atisae', 'eurocontrol', 'ecacert', 'idiada', 'oca-', 'ocaglobal', 'noverca', 'cualicontrol', 'kiwa')
# El correu d'un professional (regla 1). Sobre el text pla de l'adreca.
$Script:CtPatronsProfessional = @('engin', 'ingen', 'arquitec', 'projec', 'proyec', 'consult', 'gestor', 'gestoria', 'asesor', 'assessor', 'tramit', 'licenc', 'llicen', 'tecnic', 'estudi', 'oficina', 'enginyers.net', 'apabcn', 'coac', 'caateeb', 'aparellador', 'aparejador', 'industrial', 'enginy')

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
    foreach ($x in $Script:CtEntitatsControl) { if ($e.Contains($x)) { return $true } }
    return $false
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

function _CtContext($docsPerGia, $excel, $tecnics, $recintes = $null) {
    $ctx = @{
        Tecnics = if ($null -ne $tecnics) { $tecnics } else { @{ Emails = @{}; Noms = @() } }
        EmailTitulars = @{}; TelTitulars = @{}
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
    foreach ($g in @($gies.Keys)) {
        if ($ctx.Recintes -contains $g) { continue }
        $docs = if ($docsPerGia.ContainsKey($g)) { @($docsPerGia[$g]) } else { @() }
        $xl = if ($null -ne $excel -and $excel.ContainsKey($g)) { $excel[$g] } else { $null }
        $tit = _CtIdentitatTitular $docs $xl
        if ($tit -eq '') { $tit = 'GIA ' + $g }
        $persones = New-Object System.Collections.ArrayList
        foreach ($d in $docs) {
            $e = _CtV $d 'extret'
            foreach ($k in 'interessat', 'representant') { $p = _CtV $e $k; if ($null -ne $p) { [void]$persones.Add($p) } }
            $a = _CtV $e 'autoritzacio'
            if ($null -ne $a) { foreach ($p in @(_CtV $a 'autoritzats')) { if ($null -ne $p) { [void]$persones.Add($p) } } }
        }
        foreach ($p in $persones) {
            & $apunta $ctx.EmailTitulars (_CtEmailNet ([string](_CtV $p 'email'))) $tit
            foreach ($k in 'telefon', 'mobil') { & $apunta $ctx.TelTitulars (_CtTelefonNet ([string](_CtV $p $k))) $tit }
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
function _CtMotiuTecnic($p, $ctx, [switch]$Fort) {
    if ($null -eq $p) { return '' }
    $e = _CtEmailNet ([string](_CtV $p 'email'))
    if ($e -ne '' -and $ctx.Tecnics.Emails.ContainsKey($e)) { return "es a la llista de t$([char]0x00E8)cnics coneguts" }
    if ($e -ne '' -and $ctx.EmailTitulars.ContainsKey($e) -and $ctx.EmailTitulars[$e].Count -ge 2) {
        return ("el seu correu surt a " + $ctx.EmailTitulars[$e].Count + " activitats de titulars diferents")
    }
    foreach ($k in 'mobil', 'telefon') {
        $t = _CtTelefonNet ([string](_CtV $p $k))
        if ($t -ne '' -and $ctx.TelTitulars.ContainsKey($t) -and $ctx.TelTitulars[$t].Count -ge 2) {
            return ("el seu tel$([char]0x00E8)fon surt a " + $ctx.TelTitulars[$t].Count + " activitats de titulars diferents")
        }
    }
    if ($Fort) { return '' }
    if (_CtEmailProfessional $e) { return 'correu de professional' }
    $nom = [string](_CtV $p 'nom')
    foreach ($n in @($ctx.Tecnics.Noms)) { if (_CtMateixNom $nom $n) { return "es a la llista de t$([char]0x00E8)cnics coneguts" } }
    return ''
}

# ----------------------------------------------------------------------------
# UNA ACTIVITAT
# ----------------------------------------------------------------------------
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
    # El NIF de l'interessat del document MES RECENT (no de queixa): si no es el
    # de l'Excel, avis canvi_titular (nomes avis: l'Excel pot ser mes nou).
    $act.canvi_titular = $null
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'
        if (_CtEsQueixa ([string](_CtV $e 'expedient')) ([string](_CtV $d 'ruta'))) { continue }
        if ([string](_CtV $e 'tipus') -eq 'autoritzacio') { continue }
        $i = _CtV $e 'interessat'
        $n = _CtNifNet ([string](_CtV $i 'nif'))
        if ($n -eq '') { continue }
        if ($nifXl -ne '' -and $n -ne $nifXl) {
            $act.canvi_titular = [ordered]@{ nom = [string](_CtV $i 'nom'); nif = $n; font = [string](_CtV $d 'ruta'); data = (_CtDataDoc $d); confianca = 'segur' }
        }
        break
    }

    # 1. LES PERSONES AUTORITZADES i els candidats a representant legal.
    $aut = [ordered]@{}
    $candRep = New-Object System.Collections.ArrayList
    $afegeixAut = {
        param($p, [string]$rol, [string]$font, [string]$data, [string]$conf, [string]$motiu)
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
    foreach ($d in $ord) {
        $e = _CtV $d 'extret'; $font = [string](_CtV $d 'ruta'); $data = _CtDataDoc $d
        $queixa = _CtEsQueixa ([string](_CtV $e 'expedient')) $font
        $a = _CtV $e 'autoritzacio'
        if ($null -ne $a -and -not $queixa) {
            foreach ($p in @(_CtV $a 'autoritzats')) { if ($null -ne $p) { & $afegeixAut $p 'autoritzat' $font $data 'segur' "l'autoritzaci$([char]0x00F3) l'autoritza" } }
            $sig = _CtV $a 'signant'
            if ($null -ne $sig -and $null -ne (_CtV $a 'empresa')) { [void]$candRep.Add(@{ P = $sig; Font = $font; Data = $data; Conf = 'segur' }) }
        }
        $rep = _CtV $e 'representant'
        if ($null -ne $rep -and -not $queixa) {
            $motiu = _CtMotiuTecnic $rep $ctx
            $nifRep = _CtNifNet ([string](_CtV $rep 'nif'))
            if ($motiu -eq '' -and (_CtTipusNif $nifRep) -eq 'juridica' -and $nifRep -ne $nifTit) { $motiu = "presenta amb un NIF d'entitat (gestoria)" }
            if (_CtManEsTecnic $man $rep) { $motiu = "marcat a m$([char]0x00E0) com a t$([char]0x00E8)cnic" }
            if ($null -ne $man -and (_CtCasaClaus $rep $man.Rep)) { $motiu = ''; [void]$candRep.Add(@{ P = $rep; Font = $font; Data = $data; Conf = 'segur' }); continue }
            if ($motiu -ne '') { & $afegeixAut $rep $(if ($motiu -match 'gestoria') { 'gestoria' } else { "t$([char]0x00E8)cnic" }) $font $data 'segur' $motiu }
            elseif ($tipusTit -ne 'juridica') {
                # Persona fisica: qui presenta per ella es una persona autoritzada
                # (el representant legal d'una fisica es ella mateixa). Si es
                # ella mateixa, no.
                $int = _CtV $e 'interessat'
                $mateix = ($nifRep -ne '' -and $nifRep -eq $nifTit) -or (_CtMateixNom ([string](_CtV $rep 'nom')) ([string](_CtV $int 'nom')))
                if (-not $mateix) { & $afegeixAut $rep ('representant del tr' + [char]0x00E0 + 'mit') $font $data 'probable' 'presenta en nom del titular' }
            }
            else { [void]$candRep.Add(@{ P = $rep; Font = $font; Data = $data; Conf = 'probable' }) }
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
        if ($nifP -ne '' -and $tecNifs.ContainsKey($nifP) -and $nifP -ne $nifXl) { continue }
        # Un interessat d'UN ALTRE titular (NIF diferent del de l'Excel) no
        # omple res: es el canvi_titular dels avisos.
        $altre = ($nifXl -ne '' -and $nifP -ne '' -and $nifP -ne $nifXl)
        $conf = if ([string](_CtV $e 'tipus') -eq 'autoritzacio') { 'probable' } else { 'segur' }
        $posat = $false
        foreach ($c in 'nom', 'nif', 'email', 'telefon', 'mobil') {
            $v = [string](_CtV $p $c)
            if ($v -eq '') { continue }
            if ($c -eq 'email' -and ((_CtEmailExclos $v) -or $tecEmails.ContainsKey($v) -or (_CtMotiuTecnic @{ email = $v } $ctx -Fort) -ne '')) { continue }
            if (($c -eq 'telefon' -or $c -eq 'mobil') -and ($tecTels.ContainsKey($v) -or (_CtMotiuTecnic @{ $c = $v } $ctx -Fort) -ne '')) { continue }
            if ($altre) {
                if ($c -eq 'nom' -or $c -eq 'nif') { [void]$hist.Add([ordered]@{ camp = $c; valor = $v; font = $font; data = $data; nota = 'un altre titular' }) }
                continue
            }
            if ([string]$tit[$c] -eq '') {
                $tit[$c] = $v; $posat = $true
            } elseif ([string]$tit[$c] -ne $v) {
                [void]$hist.Add([ordered]@{ camp = $c; valor = $v; font = $font; data = $data })
            }
        }
        if ($posat -and $tit.font -eq '') { $tit.font = $font; $tit.data = $data; $tit.confianca = $conf }
    }
    if ($tit.font -ne '') { $act.titular = $tit }
    $tipusFinal = _CtTipusNif $(if ($nifXl -ne '') { $nifXl } else { [string]$tit.nif })
    $act.titular_tipus = $tipusFinal
    $act.historic = $hist.ToArray()

    # 3. EL REPRESENTANT LEGAL (regla 2). El de la correccio a ma, mana.
    if ($null -ne $repMa) {
        $act.representant_legal = _CtAmbOrigen $repMa.P $repMa.Font $repMa.Data ("a m" + [char]0x00E0)
        [void]$aplicades.Add([ordered]@{ tipus = 'es_rep_legal'; text = ([string](_CtV $repMa.P 'nom') + ' es el representant legal'); auto = [string]$repMa.Auto })
    } elseif ($tipusFinal -eq 'juridica') {
        foreach ($c in @($candRep | Sort-Object @{ Expression = { if ($_.Conf -eq 'segur') { 0 } else { 1 } } }, @{ Expression = { $_.Data }; Descending = $true })) {
            if ((_CtMotiuTecnic $c.P $ctx) -ne '' -or (& $esDeTecnic $c.P)) { continue }
            $r = _CtAmbOrigen $c.P $c.Font $c.Data $c.Conf
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
            if ($null -eq $o -and $c.Dada -eq 'email') { $mot = _CtMotiuTecnic @{ email = $vXl } $ctx -Fort }
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
