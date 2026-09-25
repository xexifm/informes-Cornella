#requires -Version 5.1
<#
.SYNOPSIS
  Informe de LLICENCIA d'activitat (Annex II de la Llei 20/2009 i llicencia
  provisional). Tres informes encadenats sobre el mateix expedient.

.DESCRIPTION
  El full de ruta d'una llicencia son TRES informes:

    1. REQUERIMENT           "Cal requerir l'esmena de les deficiencies..."
    2. FAVORABLE PRE         "S'informa favorablement a l'espera de rebre la
                              citada documentacio..." (+ condicions, opcional)
    3. FAVORABLE POST        "S'informa favorablement l'activitat i es dona per
                              tancat l'expedient." (+ condicions, opcional)

  El primer es opcional, pero es fa gairebe sempre.

  DIFERENCIA AMB UN REQUERIMENT NORMAL: aqui els punts no son deficiencies sino
  DOCUMENTACIO, i surten TANT si es te com si no. Per cada punt s'hi tria:
    - "No es disposa..."  -> NEGRETA (falta)
    - "Es disposa... (Id Firmadoc: ...)" -> sense negreta (ja hi es)
  Al Word que feia servir l'usuari sortien en verd, pero el color era una MARCA
  SEVA per veure que havia de canviar a cada informe; al document generat van
  amb el color de sempre (Format.ps1).

  D'ON SURT EL TEXT: el cos de cada punt es de REQ1, EN VIU. LLIC.json nomes hi
  afegeix el que es propi de Llicencia (els dos comentaris i el "Quan:") i una
  CLAU que apunta a l'item de REQ1 ("Seccio::Titol"). Aixi, canviar un text a
  REQ1 el canvia tambe aqui i no hi ha dues copies per mantenir. Els punts que
  no tenen equivalent a REQ1 porten el text a LLIC (no duen clau).

  ESTRUCTURA DE L'INFORME:
    DOCUMENTACIO NECESSARIA ABANS DE LA RESOLUCIO...
      (punt condicional segons Annex II / llicencia provisional)
      Autoritzacions / Informes preceptius   <- bloc ABANS de LLIC
      Projecte                               <- requeriments normals de REQ1
      Documentacio                           <- tecnic redactor + Id Firmadoc
    DOCUMENTACIO NECESSARIA DESPRES DE LA RESOLUCIO... (amb "Quan:")
    Conclusio de la fase
    ANNEX 1   <- nomes si REQUERIMENT i llicencia provisional

  Les funcions de dades son PURES (es proven en headless, sense Word); nomes
  l'assistent i la composicio del document fan servir WinForms i Word COM.
#>

# ----------------------------------------------------------------------------
# DEFINICIONS (un sol lloc)
# ----------------------------------------------------------------------------
# Les tres fases. 'Clau' es el que es desa i es compara; 'Nom' el que es
# veu; 'Conclusio' el text que tanca l'informe.
function _LlicFases {
    return @(
        [pscustomobject]@{
            Clau = 'requeriment'
            Nom  = 'Requeriment'
            Sub  = 'Es demana la documentaci' + [char]0x00F3 + ' que falta'
        }
        [pscustomobject]@{
            Clau = 'favorable-pre'
            Nom  = 'Favorable pre-llic' + [char]0x00E8 + 'ncia'
            Sub  = 'Ja hi ha tota la documentaci' + [char]0x00F3 + ' d' + [char]0x2019 + 'abans de la resoluci' + [char]0x00F3
        }
        [pscustomobject]@{
            Clau = 'favorable-post'
            Nom  = 'Favorable post-llic' + [char]0x00E8 + 'ncia'
            Sub  = 'Ja s' + [char]0x2019 + 'ha comprovat tota la documentaci' + [char]0x00F3
        }
    )
}

# Cos de lletra del full de signatures de l'ANNEX 1 (a la plantilla, sz=18
# mig-punts = 9 pt). La resta de l'informe va a 11.
$Script:LlicAnnexSignaturaCos = 9

# El titol que obre el full de signatures de l'ANNEX 1. Es mira pel PRINCIPI del
# text -es una frase llarga amb citacions legals- i sense accents, per no
# dependre de com s'hagi escrit al cataleg. Funcio PURA.
function _LlicEsTitolAcceptacio([string]$text) {
    $t = ([string]$text).Trim().ToLower()
    if ($t.Length -lt 12) { return $false }
    $t = $t.Replace([char]0x2019, "'").Replace([char]0x00F3, 'o').Replace([char]0x00E9, 'e')
    return $t.StartsWith("document d'acceptacio")
}

# Titols dels dos grans blocs de l'informe.
#
# SENSE PUNT FINAL: a l'informe fet a ma cap titol de seccio no en porta (ho
# vaig comprovar al requeriment i al favorable pre del GIA 1457). Son els dos
# unics titols de seccio escrits al codi; els de REQ1 venen del cataleg.
function _LlicTitolAbans {
    return ('DOCUMENTACI' + [char]0x00D3 + ' NECESS' + [char]0x00C0 + 'RIA ABANS DE LA RESOLUCI' + [char]0x00D3 +
            ' DE L' + [char]0x2019 + [char]0x00D2 + 'RGAN T' + [char]0x00C8 + 'CNIC AMBIENTAL')
}
function _LlicTitolProjecte {
    return 'REQUERIMENTS PROJECTE'
}
function _LlicTitolDespres {
    return ('DOCUMENTACI' + [char]0x00D3 + ' NECESS' + [char]0x00C0 + 'RIA DESPR' + [char]0x00C9 + 'S DE LA RESOLUCI' +
            [char]0x00D3 + ' DE L' + [char]0x2019 + [char]0x00D2 + 'RGAN T' + [char]0x00C8 + 'CNIC AMBIENTAL EN ELS ' +
            'TERMINIS DE TEMPS ESPECIFICATS')
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Ruta del cataleg de Llicencia.
function _LlicCatalegPath {
    return [string](Join-Path $EstructuralsDir 'LLIC.json')
}

# Llegeix LLIC.json. Retorna $null si no hi es (el programa ha de dir-ho, no
# fer com si res).
function Read-LlicCataleg([string]$path = '') {
    if ([string]::IsNullOrWhiteSpace($path)) { $path = _LlicCatalegPath }
    return (Read-JsonFile $path)
}

# Aplana el cos d'un node del JSON a les mateixes linies amb marques que fa
# servir tot el programa (**negreta**, //cursiva//, [[URL]]...). Reaprofita
# _JsonParaToBodyLine de CatalegJson.ps1: el format es el mateix.
function _LlicCos($node) {
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($node.cos)) { [void]$out.Add((_JsonParaToBodyLine $p)) }
    return $out.ToArray()
}

# El fill d'un item de LLIC amb aquell tipus ('nodisposa', 'sidisposa', 'quan',
# 'subitem'), o $null. Els tipus son els del format estandard, ampliat.
function _LlicFill($item, [string]$tipus) {
    foreach ($f in @($item.fills)) {
        if ([string]$f.tipus -eq $tipus) { return $f }
    }
    return $null
}

# Tots els fills d'un tipus (per als sub-punts, que poden ser-ne uns quants).
function _LlicFills($item, [string]$tipus) {
    $out = New-Object System.Collections.ArrayList
    foreach ($f in @($item.fills)) {
        if ([string]$f.tipus -eq $tipus) { [void]$out.Add($f) }
    }
    return $out.ToArray()
}

# Index dels items de REQ1 per clau ("Seccio::Titol"), per poder-hi anar de
# pressa. $parsed es el que retorna Get-ParsedCataleg / Read-CatalegJson.
function _LlicIndexReq1($parsed) {
    $idx = @{}
    if ($null -eq $parsed) { return $idx }
    foreach ($sec in @($parsed.Sections)) {
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -ne 'item') { continue }
            if ([string]::IsNullOrWhiteSpace([string]$el.Short)) { continue }
            $idx[(_ItemKey $sec.Title $el.Short)] = $el
        }
    }
    return $idx
}

# Les seccions de REQ1 que son DOCUMENTACIO (no deficiencies del projecte) i
# que, per tant, es demanen al bloc ABANS de la resolucio. Funcio PURA.
#
# PER QUE UNA LLISTA I NO LA DE LLIC.json: aixi un requeriment NOU d'aquestes
# seccions surt sol a la pantalla, sense haver de recordar-se d'apuntar-lo
# tambe a LLIC. LLIC.json hi aporta el "No es disposa / Es disposa" de cada un.
function _LlicSeccionsAbans {
    return @(
        'Autoritzacions / Informes preceptius'
        'Registres'
    )
}

# Una seccio de REQ1 es de les que van al bloc ABANS? Funcio PURA. Es compara
# sense accents ni apostrofs: el cataleg els escriu amb l'apostrof tipografic i
# es facil que algun dia no coincideixin caracter a caracter.
function _LlicEsSeccioAbans([string]$titol) {
    $norm = {
        param($x)
        $t = ([string]$x).Trim().ToLower()
        $t = $t.Replace([char]0x2019, "'").Replace([char]0x00F3, 'o').Replace([char]0x00E8, 'e')
        $t = $t.Replace([char]0x00E9, 'e').Replace([char]0x00E0, 'a').Replace([char]0x00ED, 'i')
        return ($t -replace '\s+', ' ')
    }
    $n = & $norm $titol
    foreach ($s in (_LlicSeccionsAbans)) { if ((& $norm $s) -eq $n) { return $true } }
    return $false
}

# On viu un element de REQ1: @{ Clau; Seccio; Subseccio }. Funcio PURA.
# El model pla no porta la seccio a dins de l'element, i la SUBSECCIO nomes es
# sap recorrent la llista en ordre (Kind='subsection' i despres els seus items).
function _LlicUbicacioDeItem($req1, $el) {
    foreach ($sec in @($req1.Sections)) {
        $sub = ''
        foreach ($x in @($sec.Items)) {
            if ([string]$x.Kind -eq 'subsection') { $sub = [string]$x.Short; continue }
            if ([object]::ReferenceEquals($x, $el)) {
                return @{ Clau = (_ItemKey ([string]$sec.Title) ([string]$el.Short))
                          Seccio = [string]$sec.Title; Subseccio = $sub }
            }
        }
    }
    return @{ Clau = ''; Seccio = ''; Subseccio = '' }
}

# Els ITEMS d'una SECCIO o SUBSECCIO de REQ1. Funcio PURA.
#
# La clau pot ser "Seccio" (tota la seccio) o "Seccio::Subseccio" (nomes
# aquella part). El lector aplana les subseccions -Kind='subsection' seguit dels
# seus items a la MATEIXA llista-, o sigui que els items d'una subseccio son els
# que van despres del seu marcador i abans del marcador seguent.
function _LlicItemsDeSubseccio($req1, [string]$clau) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $req1 -or [string]::IsNullOrWhiteSpace($clau)) { return $out.ToArray() }
    $i = $clau.IndexOf('::')
    $secDemanada = if ($i -gt 0) { $clau.Substring(0, $i) } else { $clau }
    $subDemanada = if ($i -gt 0) { $clau.Substring($i + 2) } else { '' }
    foreach ($sec in @($req1.Sections)) {
        if ([string]$sec.Title -ne $secDemanada) { continue }
        $dins = [string]::IsNullOrWhiteSpace($subDemanada)   # tota la seccio: des del principi
        foreach ($el in @($sec.Items)) {
            $kind = [string]$el.Kind
            if ($kind -eq 'subsection') {
                if (-not [string]::IsNullOrWhiteSpace($subDemanada)) { $dins = ([string]$el.Short -eq $subDemanada) }
                continue
            }
            if (-not $dins) { continue }
            if ($kind -ne 'item') { continue }
            [void]$out.Add($el)
        }
    }
    return $out.ToArray()
}

# Les SECCIONS i SUBSECCIONS de REQ1 que un bloc de LLIC expandeix senceres.
# Funcio PURA. Surten del PROPI cataleg (una entrada amb clau que NO es un
# item), no d'una llista al codi: aixi l'usuari pot moure una seccio de bloc
# des de l'editor sense tocar el programa.
function _LlicSeccionsExpandides($llic, $idxReq1) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $llic) { return $out.ToArray() }
    foreach ($sec in @($llic.nodes)) {
        foreach ($it in @($sec.fills)) {
            $c = [string]$it.clau
            if ([string]::IsNullOrWhiteSpace($c)) { continue }
            if ($null -ne $idxReq1 -and $idxReq1.ContainsKey($c)) { continue }   # es un item
            if (-not $out.Contains($c)) { [void]$out.Add($c) }
        }
    }
    return $out.ToArray()
}

function _LlicSeccionsSenseSubseccions($sections, $claus) {
    $fora = @{}
    foreach ($c in @($claus)) { $fora[[string]$c] = $true }
    $out = New-Object System.Collections.ArrayList
    foreach ($sec in @($sections)) {
        # Una clau sense '::' treu la SECCIO sencera.
        if ($fora.ContainsKey([string]$sec.Title)) { continue }
        $items = New-Object System.Collections.ArrayList
        $saltant = $false
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection') {
                $saltant = $fora.ContainsKey((_ItemKey ([string]$sec.Title) ([string]$el.Short)))
                if ($saltant) { continue }
            }
            if ($saltant -and [string]$el.Kind -ne 'subsection') { continue }
            [void]$items.Add($el)
        }
        $teItem = $false
        foreach ($el in $items) { if ([string]$el.Kind -eq 'item') { $teItem = $true; break } }
        if (-not $teItem) { continue }
        [void]$out.Add([pscustomobject]@{ Title = [string]$sec.Title; Items = $items.ToArray() })
    }
    return $out.ToArray()
}

# ELS TEXTOS PER DEFECTE d'un punt d'ABANS que no consta a LLIC.json. Funcio
# PURA. Retorna @{ NoDisposa; SiDisposa }.
#
# PER QUE: la llista d'ABANS surt de les 4 seccions de REQ1 (43 punts) i LLIC
# nomes en descriu 15. Als altres 28, triar "Es disposa del document" no
# ensenyava res i a l'informe no s'hi escrivia res. Com a MINIM tots han de
# poder dir que es tenen, amb el seu Id Firmadoc; qui necessiti una redaccio
# propia (l'expedient, la referencia, el NIMA...) la posa a LLIC.json i mana
# aquella. Aixi un requeriment NOU de REQ1 ja surt utilitzable sense tocar res.
function _LlicTextosPerDefecte {
    return @{
        NoDisposa = @('No es disposa del document.')
        SiDisposa = @('Es disposa del document (Id Firmadoc: [CAMP: Id Firmadoc])')
    }
}

# EL COS DE L'EINA, i es PUR: resol els punts d'un bloc de LLIC ('ABANS',
# 'DESPRES' o 'PROPIS') ajuntant-los amb el text de REQ1.
#
# Retorna @{ Punts; Orfes }:
#   Punts : llista de @{ Clau; Titol; Cos; NoDisposa; SiDisposa; Quan; Subs;
#                        Condicio } en l'ordre del cataleg.
#   Orfes : claus que son a LLIC pero JA NO a REQ1. NO s'amaguen: si el lligam
#           s'ha trencat (perque algu ha reanomenat un requeriment), el
#           programa ho ha de dir en lloc de deixar-se un punt en silenci.
# ON VIU CADA ITEM d'una llista de seccions (la de REQ1 o la que torna
# Select-Items) i QUIN TEXT FIX l'encapcala. Funcio PURA.
#
# Retorna, per cada item: @{ Seccio; Subseccio; Intro; El }.
#
# PER QUE: als informes de Llicencia hi han de sortir la SECCIO i la SUBSECCIO
# de REQ1 de cada punt, i els TEXTOS FIXOS que encapcalen una subseccio (p.ex.
# "Segons l'article 3.1 de l'Ordenanca ... adjuntant la documentacio
# complementaria:"). Abans els punts s'aplanaven i tot aixo es perdia: sortien
# tots seguits, sense saber de quina part venien.
#
# CADA ITEM ES PORTA L'INTRO QUE LI TOCA, no nomes el primer de la llista.
#
# Aixo es important i va fallar: si l'intro s'enganxa NOMES al primer item del
# cataleg i l'usuari no tria aquell item, el text fix no surt. Va passar de debo
# a Instal-lacions / Legalitzacions: l'intro ("Segons l'article 4 de
# l'Ordenanca...") penja de "RITSIC - fotovoltaica", i en un informe que nomes
# demanava la baixa tensio i el PCI no sortia enlloc.
#
# A REQ1 aixo no passa perque alli l'intro queda PENDENT fins que surt un item,
# sigui quin sigui. Aqui es reprodueix igual: cada item es queda l'ultim intro
# vist DINS DEL SEU GRUP, i qui l'escriu nomes el treu quan CANVIA (vegeu
# _LlicEscriuPunts). Amb aixo:
#   - si el primer item del grup no es tria, l'intro surt igualment amb el que
#     surti primer;
#   - i si un grup tingues dos intros, cada un sortiria davant dels seus items,
#     exactament com faria REQ1.
# Una subseccio nova (o una seccio nova) buida l'intro pendent.
#
# El titol pot venir com a "Seccio - Subseccio" (el que munta
# Build-SelectionFromKeys) o com a titol sol (el cataleg sencer).
function _LlicItemsAmbUbicacio($seccions) {
    $out = New-Object System.Collections.ArrayList
    foreach ($sec in @($seccions)) {
        $parts = ([string]$sec.Title) -split ' - ', 2
        $nomSec = if ($parts.Count -eq 2) { $parts[0].Trim() } else { [string]$sec.Title }
        $sub    = if ($parts.Count -eq 2) { $parts[1].Trim() } else { '' }
        # Un text fix d'ABANS de la primera subseccio es de la SECCIO i sobreviu
        # als canvis de subseccio; el de DINS d'una subseccio mor amb ella.
        # Mateixa regla que Build-CatalegBlocs (MotorInforme.ps1), i pel mateix
        # motiu: sense aixo, un text posat a la seccio no sortiria MAI.
        $intro = @()
        $introSec = @()
        $dinsSub = $false
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection') {
                $sub = [string]$el.Short; $dinsSub = $true; $intro = @(); continue
            }
            if ([string]$el.Kind -eq 'intro') {
                if ($dinsSub) { $intro = @($el.BodyLines) } else { $introSec = @($el.BodyLines) }
                continue
            }
            if ([string]$el.Kind -ne 'item') { continue }
            $deSec = (@($intro).Count -eq 0 -and @($introSec).Count -gt 0)
            $lines = if (@($intro).Count -gt 0) { @($intro) } else { @($introSec) }
            [void]$out.Add(@{ Seccio = $nomSec; Subseccio = $sub; Intro = $lines
                              IntroDeSeccio = $deSec; El = $el })
        }
    }
    return $out.ToArray()
}

function _LlicPuntsPerBloc($llic, $idxReq1, [string]$bloc, $req1 = $null) {
    $punts = New-Object System.Collections.ArrayList
    $orfes = New-Object System.Collections.ArrayList
    if ($null -eq $llic) { return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() } }

    # EL BLOC 'ABANS' surt de REQ1, no de la llista de LLIC: son TOTS els items
    # de les seccions de documentacio (_LlicSeccionsAbans). LLIC nomes hi posa
    # el "No es disposa / Es disposa" de cada un, per clau. Un requeriment nou
    # d'aquelles seccions surt sol, encara que ningu l'hagi apuntat a LLIC.
    if ($bloc -eq 'ABANS' -and $null -ne $req1) {
        $perClau = @{}
        foreach ($s in @($llic.nodes)) {
            if ([string]$s.titol -ne 'ABANS') { continue }
            foreach ($it in @($s.fills)) {
                $c = [string]$it.clau
                if (-not [string]::IsNullOrWhiteSpace($c)) { $perClau[$c] = $it }
            }
        }
        $seccionsAbans = @(@($req1.Sections) | Where-Object { _LlicEsSeccioAbans ([string]$_.Title) })
        foreach ($u in @(_LlicItemsAmbUbicacio $seccionsAbans)) {
            $el = $u.El
            if (-not [string]::IsNullOrWhiteSpace([string]$el.Short)) {
                $clau = _ItemKey $u.Seccio $el.Short
                $it = if ($perClau.ContainsKey($clau)) { $perClau[$clau] } else { $null }
                $nod = if ($null -ne $it) { _LlicFill $it 'nodisposa' } else { $null }
                $sid = if ($null -ne $it) { _LlicFill $it 'sidisposa' } else { $null }
                # Els que no consten a LLIC (o hi consten sense text) agafen els
                # textos per defecte: tots han de poder dir que es tenen.
                $def = _LlicTextosPerDefecte
                $lNod = if ($null -ne $nod) { @(_LlicCos $nod) } else { @() }
                $lSid = if ($null -ne $sid) { @(_LlicCos $sid) } else { @() }
                if (@($lNod).Count -eq 0) { $lNod = @($def.NoDisposa) }
                if (@($lSid).Count -eq 0) { $lSid = @($def.SiDisposa) }
                [void]$punts.Add([pscustomobject]@{
                    Clau      = $clau
                    Seccio    = [string]$u.Seccio
                    Subseccio = [string]$u.Subseccio
                    Intro     = @($u.Intro)
                    IntroDeSeccio = [bool]$u.IntroDeSeccio
                    Titol     = [string]$el.Short
                    Condicio  = ''
                    Cos       = @($el.BodyLines)
                    NoDisposa = $lNod
                    SiDisposa = $lSid
                    Quan      = @()
                    Subs      = @(@($el.Children) | ForEach-Object { @($_.BodyLines) })
                })
            }
        }
        return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() }
    }

    $sec = $null
    foreach ($s in @($llic.nodes)) {
        if ([string]$s.titol -eq $bloc) { $sec = $s; break }
    }
    if ($null -eq $sec) { return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() } }

    foreach ($it in @($sec.fills)) {
        $clau = [string]$it.clau
        $nod = _LlicFill $it 'nodisposa'
        $sid = _LlicFill $it 'sidisposa'
        $qua = _LlicFill $it 'quan'
        $lNod = if ($null -ne $nod) { @(_LlicCos $nod) } else { @() }
        $lSid = if ($null -ne $sid) { @(_LlicCos $sid) } else { @() }
        $lQua = if ($null -ne $qua) { @(_LlicCos $qua) } else { @() }

        # UNA CLAU POT SER UNA SECCIO O UNA SUBSECCIO SENCERA de REQ1, i llavors
        # l'entrada s'EXPANDEIX: un punt per cada item d'aquella part, amb el
        # text LITERAL de REQ1 i el mateix "Quan:" per a tots. Aixi el bloc
        # DESPRES es porta seccions senceres (Instal-lacions, Controls
        # inicials...) sense mantenir-ne cap copia, i un requeriment nou d'aquella
        # seccio hi surt sol.
        $esItem = ($null -ne $idxReq1 -and $idxReq1.ContainsKey($clau))
        if (-not [string]::IsNullOrWhiteSpace($clau) -and -not $esItem) {
            $delsSubs = @(_LlicItemsDeSubseccio $req1 $clau)
            if (@($delsSubs).Count -eq 0) {
                [void]$orfes.Add($clau)
                continue
            }
            # On viu cada item i quin text fix l'encapcala: aixi el bloc es porta
            # la SECCIO, la SUBSECCIO i l'INTRO de REQ1, no nomes els punts.
            $ubis = @{}
            foreach ($u in @(_LlicItemsAmbUbicacio $req1.Sections)) {
                $k = _ItemKey ([string]$u.Seccio) ([string]$u.El.Short)
                if (-not $ubis.ContainsKey($k)) { $ubis[$k] = $u }
            }
            foreach ($el in $delsSubs) {
                $ub = _LlicUbicacioDeItem $req1 $el
                $u = $ubis[[string]$ub.Clau]
                [void]$punts.Add([pscustomobject]@{
                    Clau      = [string]$ub.Clau
                    Seccio    = [string]$ub.Seccio
                    Intro     = @($(if ($null -ne $u) { $u.Intro } else { @() }))
                    IntroDeSeccio = [bool]$(if ($null -ne $u) { $u.IntroDeSeccio } else { $false })
                    Subseccio = [string]$ub.Subseccio
                    Titol     = [string]$el.Short
                    Condicio  = [string]$it.condicio
                    Cos       = @($el.BodyLines)
                    NoDisposa = $lNod
                    SiDisposa = $lSid
                    Quan      = $lQua
                    Subs      = @(@($el.Children) | ForEach-Object { ,@($_.BodyLines) })
                })
            }
            continue
        }

        $cos = if ($esItem) { @($idxReq1[$clau].BodyLines) } else { @(_LlicCos $it) }
        # La SECCIO surt de la clau ("Seccio::Item"). Els punts PROPIS de LLIC no
        # en tenen cap i van al principi del bloc, sense capcalera.
        $ubIt = if ($esItem -and $null -ne $req1) { _LlicUbicacioDeItem $req1 $idxReq1[$clau] } else { $null }
        [void]$punts.Add([pscustomobject]@{
            Clau      = $clau
            Seccio    = [string]$(if ($null -ne $ubIt) { $ubIt.Seccio } else { '' })
            Subseccio = [string]$(if ($null -ne $ubIt) { $ubIt.Subseccio } else { '' })
            Intro     = @()
            Titol     = [string]$it.titol
            Condicio  = [string]$it.condicio
            Cos       = $cos
            NoDisposa = $lNod
            SiDisposa = $lSid
            Quan      = $lQua
            Subs      = @(@(_LlicFills $it 'subitem') | ForEach-Object { @(_LlicCos $_) })
        })
    }
    return @{ Punts = $punts.ToArray(); Orfes = $orfes.ToArray() }
}

# Els documents que pot haver signat el tecnic redactor. Funcio PURA (i per aixo
# es aqui i no dins del dialeg: aixi es pot COMPTAR en una prova).
#
# ATENCIO als PARENTESIS: dins d'un @(...) la coma lliga MES FORT que el '+', o
# sigui que @('Pl' + [char]0x00E0 + 'nols') son TRES elements, no un. Aixo va
# passar de debo: a la pantalla hi sortien cinc caselles -Projecte, Pl, a, nols,
# Annexos- en lloc de tres. Esta avisat a CLAUDE.md i hi vaig caure igualment.
function _LlicDocsSignats {
    return @('Projecte', ('Pl' + [char]0x00E0 + 'nols'), 'Annexos')
}

# Els noms dels camps [CAMP: ...] que hi ha en unes linies de text. Funcio PURA.
# Serveix per saber QUINES dades ha d'omplir l'usuari quan diu que ja disposa
# d'un document (Id Firmadoc, Expedient, Referencia... segons el punt).
function _LlicCampsDelText($linies) {
    $out = New-Object System.Collections.ArrayList
    foreach ($l in @($linies)) {
        foreach ($m in [regex]::Matches([string]$l, '\[CAMP:\s*([^\]]+)\]')) {
            $nom = ([string]$m.Groups[1].Value).Trim()
            if ($nom -and -not $out.Contains($nom)) { [void]$out.Add($nom) }
        }
    }
    return $out.ToArray()
}

# Substitueix els [CAMP: nom] d'unes linies pels valors donats. Funcio PURA.
#
# PER QUE NO ES FA SERVIR EL DICCIONARI DE CAMPS COMPARTIT: alli les claus son
# el NOM del camp, i aqui "Id Firmadoc" te un valor DIFERENT a cada punt (cada
# document te el seu). Per aixo el valor es resol punt a punt i s'hi deixa el
# text ja resolt.
function _LlicAplicaCamps($linies, $valors) {
    $out = New-Object System.Collections.ArrayList
    foreach ($l in @($linies)) {
        $t = [string]$l
        foreach ($m in [regex]::Matches($t, '\[CAMP:\s*([^\]]+)\]')) {
            $nom = ([string]$m.Groups[1].Value).Trim()
            $v = ''
            if ($null -ne $valors -and $valors.Contains($nom)) { $v = [string]$valors[$nom] }
            $t = $t.Replace([string]$m.Value, $v)
        }
        [void]$out.Add($t)
    }
    return $out.ToArray()
}

# La clau amb que es recorda que havia triat l'usuari a la pantalla de
# documentacio (per poder-ho tornar a pintar si torna ENRERE). Funcio PURA: la
# clau de REQ1 si en te, i si no el titol -que es l'unic que distingeix els
# punts propis-.
function _LlicClauPunt($punt) {
    $c = [string]$punt.Clau
    if (-not [string]::IsNullOrWhiteSpace($c)) { return $c }
    return ('#' + [string]$punt.Titol)
}

# La SECCIO d'un punt, per agrupar-lo a la pantalla de documentacio. Funcio
# PURA i sense esquema nou: la clau d'un punt que ve de REQ1 ja es
# "Seccio::Item" (_ItemKey, Motor.ps1), o sigui que la seccio es el tros
# d'abans del "::". Els punts PROPIS (i els que es llegeixen d'un informe
# anterior) no tenen clau: retornen '' i van al primer nivell de l'arbre.
function _LlicSeccioDePunt($punt) {
    $c = [string]$punt.Clau
    if ([string]::IsNullOrWhiteSpace($c)) { return '' }
    $i = $c.IndexOf('::')
    if ($i -le 0) { return '' }
    return $c.Substring(0, $i)
}

# El text d'un punt a l'arbre de la pantalla de documentacio. Funcio PURA.
#
# El TITOL primer: als punts que venen de REQ1 es el nom curt del cataleg
# ("Sanitat", "Incendis"), que es exactament el que surt al Pas 3. Nomes es cau
# al cos quan no n'hi ha (els punts trets d'un informe ja emes).
function _LlicEtiquetaPunt($punt, [int]$max = 110) {
    $t = [string]$punt.Titol
    if ([string]::IsNullOrWhiteSpace($t)) {
        foreach ($l in @($punt.Cos)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$l)) { $t = [string]$l; break }
        }
    }
    $t = ([string]$t -replace '\s+', ' ').Trim()
    if ($max -gt 0 -and $t.Length -gt $max) { $t = $t.Substring(0, $max).TrimEnd() + [char]0x2026 }
    return $t
}

# El cos d'un punt, en text pla per ensenyar-lo a la pantalla. Funcio PURA.
#
# Treu el marcador intern '[[URL]] ' -que el posa el lector del cataleg als
# paragrafs d'enllac (CatalegJson.ps1)- i deixa l'adreca. Sortia TAL QUAL al
# panell de detall i al tooltip de l'arbre.
function _LlicTextPlaDelCos($cos) {
    $t = (@($cos) -join ' ')
    $t = $t -replace '\[\[URL\]\]\s*', ''
    return (($t -replace '\s+', ' ').Trim())
}

# Agrupa els punts per SECCIO i SUBSECCIO per pintar-los en ARBRE. Funcio PURA.
#
# Retorna els grups en ORDRE DE PRIMERA APARICIO, cada un amb els INDEXS dels
# seus punts dins de $punts:
#     @( @{ Titol=''; Sub=''; Idx=@(0,1) },
#        @{ Titol='Instal-lacions'; Sub='Legalitzacions'; Idx=@(5,6) } )
#
# DOS NIVELLS, com el Pas 3: a REQ1 les seccions grans (Instal-lacions,
# Registres, Incendis) tenen subseccions, i sense elles surten trenta punts
# seguits a la mateixa alcada i no es poden llegir.
#
# Els punts sense seccio van al grup de titol '' -el primer nivell de l'arbre,
# sense capcalera-.
#
# AIXO NOMES ES DE PANTALLA: l'informe es munta recorrent $punts en l'ordre del
# cataleg, no l'arbre, o sigui que agrupar no reordena res del document.
function _LlicAgrupaPunts($punts) {
    $punts = @($punts)
    $ordre = New-Object System.Collections.ArrayList
    $perGrup = @{}
    for ($i = 0; $i -lt $punts.Count; $i++) {
        $sec = _LlicSeccioDePunt $punts[$i]
        $sub = [string]$punts[$i].Subseccio
        if ([string]::IsNullOrWhiteSpace($sec)) { $sub = '' }   # al primer nivell no hi ha subseccio
        $clau = $sec + [char]0x0001 + $sub
        if (-not $perGrup.ContainsKey($clau)) {
            $perGrup[$clau] = New-Object System.Collections.ArrayList
            [void]$ordre.Add(@{ Clau = $clau; Titol = $sec; Sub = $sub })
        }
        [void]$perGrup[$clau].Add($i)
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($g in $ordre) {
        [void]$out.Add(@{ Titol = [string]$g.Titol; Sub = [string]$g.Sub; Idx = $perGrup[$g.Clau].ToArray() })
    }
    return $out.ToArray()
}

# La CLASSIFICACIO de l'activitat ("Llei 20/2009; Annex II; Epigraf 12.25" o
# "Llei 18/2020; Epigraf ..."), que nomes surt als informes de Llicencia.
#
# SURT SOLA, no es pregunta: es llegeix de l'Excel per ID GIA (_ClassificacioText
# la munta a Activitats.ps1 a partir de "Classificacio general annex" i
# "... Apartat"). Si no se'n troba cap, es deixa BUIDA i el crider ho avisa en
# acabar: aturar l'assistent per aixo seria pitjor que generar l'informe.
#
# ATENCIO: la fitxa de la cache es un HASHTABLE (Activitats.ps1 hi desa @{...}),
# no un PSCustomObject. Amb $act.PSObject.Properties['CLASSIFICACIO'] sempre
# sortia buit i per aixo es preguntava sempre.
function _LlicClassificacio($header, $cache = $null) {
    if ($null -eq $header) { return '' }
    if ($header.Contains('CLASSIFICACIO')) {
        $ja = [string]$header['CLASSIFICACIO']
        if (-not [string]::IsNullOrWhiteSpace($ja)) { return $ja }
    }
    if ($null -eq $cache) { $cache = $script:_sessionActCache }
    try {
        $idGia = [string]$header['ID_GIA']
        if ([string]::IsNullOrWhiteSpace($idGia)) { return '' }
        $act = Get-ActivitatFromCache $cache $idGia
        if ($null -eq $act) { return '' }
        if ($act -is [System.Collections.IDictionary]) {
            if ($act.Contains('CLASSIFICACIO')) { return [string]$act['CLASSIFICACIO'] }
            return ''
        }
        if ($act.PSObject.Properties.Name -contains 'CLASSIFICACIO') { return [string]$act.CLASSIFICACIO }
    } catch { }
    return ''
}

# Quins punts condicionals entren, segons si es llicencia provisional.
#   'annexii'     -> nomes si NO ho es
#   'provisional' -> nomes si SI ho es
# Un punt sense condicio entra sempre. Funcio PURA.
function _LlicCondicioEntra([string]$condicio, [bool]$esProvisional) {
    $c = ([string]$condicio).Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($c)) { return $true }
    if ($c -eq 'provisional') { return $esProvisional }
    if ($c -eq 'annexii') { return (-not $esProvisional) }
    return $true
}

# Avisa d'on ha quedat l'informe i el deixa obert al Word, com la resta de
# fluxos del programa. Es un sol lloc perque els dos camins de l'assistent
# -l'informe llarg i els dos curts- acabin exactament igual.
function _LlicObreIAvisa($word, [string]$out) {
    [System.Windows.Forms.MessageBox]::Show(
        "Informe generat:`n$out", 'Finalitzat', 'OK', 'Information') | Out-Null
    $word.Visible = $true
    $word.Documents.Open($out) | Out-Null
}

# TOTES LES FASES que ofereix el pas 1 de Llicencia: les tres de l'informe
# llarg i les dues curtes (Modificacio NO Substancial i Traspas). Funcio PURA.
#
# Van juntes al mateix menu perque per a l'usuari son "l'informe de la
# llicencia" i comparteixen capcalera; el que canvia es el document que en
# surt, i d'aixo ja se n'ocupa cada modul.
# QUINA FASE SURT MARCADA. Funcio PURA.
#
# Retorna $preFase si es de la llista, i si no la PRIMERA de la llista.
#
# ATENCIO al motiu: aqui hi havia un 'requeriment' escrit al codi com a
# respatller. Des que MNS/Traspas te entrada propia, la seva llista NO en te
# cap, i la pantalla petava amb "La propiedad 'Checked' no se encuentra en este
# objeto" -perque $radios['requeriment'] era $null. Cap llista de fases pot
# donar per fet quines fases porta.
function _LlicFasePerDefecte($fases, [string]$preFase) {
    $llista = @($fases)
    if ($llista.Count -eq 0) { return '' }
    foreach ($f in $llista) { if ([string]$f.Clau -eq [string]$preFase) { return [string]$preFase } }
    return [string]$llista[0].Clau
}

function _LlicTotesLesFases {
    $out = New-Object System.Collections.ArrayList
    foreach ($f in @(_LlicFases)) { [void]$out.Add($f) }
    foreach ($f in @(_MnsFases))  { [void]$out.Add($f) }
    return $out.ToArray()
}

# EL BLOC DESPRES SEGONS LA FASE. Funcio PURA.
#
# ELS TRES INFORMES SON EL MATEIX DOCUMENT. Aixo es el que va costar de veure:
# el favorable POST no es un informe curt que llegeix l'anterior -aixi estava
# fet i no es el que fa l'usuari a ma-, sino EL MATEIX informe sencer
# (documentacio del projecte, bloc ABANS i bloc DESPRES amb els seus "Quan:").
# L'unica cosa que canvia entre les tres fases es QUE ES DIU DE CADA PUNT DEL
# BLOC DESPRES:
#
#   requeriment    -> res (encara no toca dir si es te o no)
#   favorable-pre  -> res (TAMPOC: vegeu a sota)
#   favorable-post -> "Es disposa del document (Id Firmadoc: ...)" o, si algun
#                     encara falta, "No es disposa de la documentacio."
#
# ...i la conclusio, que ja la decidia _LlicConclusioText.
#
# EL PRE NO DIU "No es disposa de la documentacio." Hi era, en negreta, a tots
# els punts del bloc, i l'usuari ho va treure (setembre 2026): al pre-llicencia
# la documentacio de DESPRES de la resolucio encara no toca tenir-la -el bloc ja
# diu "en els terminis de temps especificats" i cada punt porta el seu "Quan:"-,
# o sigui que dir de cada punt que falta no aporta res. Quedar-se o no amb el
# document nomes es diu al POST, que es quan es comprova.
function _LlicEstatDespres([string]$fase) {
    $sid = @('Es disposa del document (Id Firmadoc: [CAMP: Id Firmadoc])')
    $nod = @('No es disposa de la documentaci' + [char]0x00F3 + '.')
    switch ([string]$fase) {
        'favorable-post' { return @{ Estat = 'si'; NoDisposa = $nod; SiDisposa = $sid; AmbEstat = $true;  AmbDades = $true } }
    }
    return @{ Estat = ''; NoDisposa = @(); SiDisposa = @(); AmbEstat = $false; AmbDades = $false }
}

# Els punts del bloc DESPRES amb els textos de la fase. Funcio PURA: retorna
# copies, mai toca els punts que li arriben. Un punt que ja porti text propi al
# cataleg el conserva -alli hi ha la redaccio bona i aqui nomes hi ha el text
# generic-.
function _LlicPuntsAmbEstatFase($punts, [string]$fase) {
    $ef = _LlicEstatDespres $fase
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($punts)) {
        $nod = @($p.NoDisposa); $sid = @($p.SiDisposa)
        if ($nod.Count -eq 0) { $nod = @($ef.NoDisposa) }
        if ($sid.Count -eq 0) { $sid = @($ef.SiDisposa) }
        # LA COPIA ES 'Select-Object *', NO una llista de camps a ma. Aqui hi
        # havia els nou camps enumerats i, en afegir-hi Seccio i Intro, es van
        # PERDRE en silenci: el bloc DESPRES -el mes gros- sortia amb les
        # subseccions pero sense les seccions ni els textos fixos. No petava:
        # simplement faltaven. Ho va enxampar un fitxer d'or.
        $c = $p | Select-Object *
        $c.NoDisposa = $nod
        $c.SiDisposa = $sid
        [void]$out.Add($c)
    }
    return $out.ToArray()
}

# CAL L'ANNEX 1? Funcio PURA.
#
# L'ANNEX 1 diu QUINA DOCUMENTACIO s'ha d'enviar per demanar l'autoritzacio
# d'usos i obres provisionals. Si l'usuari ha marcat que d'aquella autoritzacio
# JA SE'N DISPOSA, l'annex no te cap sentit: ja no s'ha de demanar res.
#
# El punt es reconeix per la CONDICIO 'provisional' (la que el fa entrar nomes a
# les llicencies provisionals), no pel titol: el titol es pot reescriure des de
# l'editor de catalegs i el lligam es trencaria en silenci.
function _LlicCalAnnex1($punts, [bool]$esProvisional) {
    if (-not $esProvisional) { return $false }
    foreach ($p in @($punts)) {
        if (([string]$p.Condicio).Trim().ToLower() -ne 'provisional') { continue }
        if ([string]$p.Estat -eq 'si') { return $false }
    }
    return $true
}

# QUINES FASES PODEN PORTAR CONDICIONS: els dos favorables. Funcio PURA, i
# l'unic lloc que ho diu: el pas de les condicions i la conclusio ho pregunten
# aqui.
function _LlicAdmetCondicions([string]$fase) {
    return ([string]$fase -eq 'favorable-pre' -or [string]$fase -eq 'favorable-post')
}

# ELS ACTORS QUE POSEN CONDICIONS. Funcio PURA.
#
# Les condicions d'una llicencia no les escriu l'Ajuntament: les posen els
# organismes que emeten els informes preceptius (OGAU, Agencia de Residus de
# Catalunya, Direccio General de Canvi Climatic...), i l'informe nomes diu QUINS
# son -els seus informes van adjunts a continuacio-. Per aixo el pas de les
# condicions es una LLISTA d'actors per marcar, i n'hi ha prou amb un de marcat
# perque l'informe porti condicions.
#
# LA LLISTA VIU AL CATALEG (seccio CONDICIONS de LLIC.json), no al codi: cada
# item es un actor (el titol, tal com surt a l'informe) i la seva CLAU apunta al
# punt de REQ1 que el fa intervenir. Un actor pot sortir diverses vegades amb
# claus diferents (l'ACA, per exemple, en te dues): aqui es fonen en un de sol.
#
# Retorna, en l'ordre del cataleg: @{ Nom; Claus[] }.
function _LlicActorsCondicions($llic) {
    $out = New-Object System.Collections.ArrayList
    $perNom = @{}
    if ($null -eq $llic) { return $out.ToArray() }
    foreach ($sec in @($llic.nodes)) {
        if (([string]$sec.titol).Trim().ToUpper() -ne 'CONDICIONS') { continue }
        foreach ($it in @($sec.fills)) {
            $nom = ([string]$it.titol).Trim()
            if ([string]::IsNullOrWhiteSpace($nom)) { continue }
            $k = $nom.ToLowerInvariant()
            if (-not $perNom.ContainsKey($k)) {
                $a = @{ Nom = $nom; Claus = (New-Object System.Collections.ArrayList) }
                $perNom[$k] = $a
                [void]$out.Add($a)
            }
            $c = [string]$it.clau
            if (-not [string]::IsNullOrWhiteSpace($c) -and -not $perNom[$k].Claus.Contains($c)) {
                [void]$perNom[$k].Claus.Add($c)
            }
        }
    }
    return $out.ToArray()
}

# QUINS ACTORS SURTEN MARCATS la primera vegada. Funcio PURA.
#
# Els que tenen algun dels seus punts marcat al bloc ABANS amb "Es disposa":
# si ja hi ha l'informe preceptiu, es que l'organisme l'ha emes, i es aquell
# informe el que porta les condicions. Es nomes el punt de partida: l'usuari
# ho pot canviar tot.
function _LlicActorsPerDefecte($actors, $abans) {
    $ambInforme = @{}
    foreach ($p in @($abans)) {
        if ($null -eq $p) { continue }
        if ([string]$p.Estat -ne 'si') { continue }
        $c = [string]$p.Clau
        if (-not [string]::IsNullOrWhiteSpace($c)) { $ambInforme[$c] = $true }
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($a in @($actors)) {
        foreach ($c in @($a.Claus)) {
            if ($ambInforme.ContainsKey([string]$c)) { [void]$out.Add([string]$a.Nom); break }
        }
    }
    return $out.ToArray()
}

# Text de la conclusio d'una fase. Funcio PURA.
#
# ELS DOS FAVORABLES PODEN PORTAR CONDICIONS, i llavors la frase ho anuncia
# ("...sota les condicions que es determinen en els seguents informes (adjunts a
# continuacio):") i a sota hi van els actors que les posen. Hi ha condicions
# quan s'ha marcat com a minim un actor al pas de les condicions.
function _LlicConclusioText([string]$fase, [bool]$ambCondicions) {
    # EL TEXT VE DEL CATALEG, del grup 'LLIC' de '0 CONCLUSIONS.json', i el titol
    # de cada entrada es la CLAU DE LA FASE. Abans era al codi (_LlicFases), o
    # sigui que canviar una conclusio de llicencia volia dir tocar el programa
    # mentre que les de REQ1 s'editaven des de l'editor de catalegs.
    #
    # Amb condicions, l'entrada es '<fase>-condicions'. Si el cataleg no la te
    # (un 0 CONCLUSIONS.json de l'usuari encara sense actualitzar), es fa servir
    # la de la fase: val mes una conclusio sense la coda que cap conclusio.
    $c = $null
    try { $c = Read-Conclusions $ConclusionsPath 'LLIC' } catch { $c = $null }
    if ($null -eq $c) { return '' }
    $titols = @([string]$fase)
    if ($ambCondicions -and (_LlicAdmetCondicions $fase)) { $titols = @(([string]$fase + '-condicions'), [string]$fase) }
    foreach ($t in $titols) {
        foreach ($x in @($c.Selectable)) {
            if ([string]$x.Title -eq $t) { return [string]$x.Body }
        }
    }
    return ''
}

# El paragraf "Documentacio signada digitalment pel tecnic redactor..." Funcio
# PURA. $data ja ve formatada com la vol l'usuari ("20 de febrer de 2024").
function _LlicTextDocumentacio([string]$tecnic, [string]$numCol, [string]$collegi, [string]$data) {
    if ([string]::IsNullOrWhiteSpace($tecnic)) { return '' }
    $t = ('Documentaci' + [char]0x00F3 + ' signada digitalment pel t' + [char]0x00E8 + 'cnic redactor ' + $tecnic.Trim())
    if (-not [string]::IsNullOrWhiteSpace($numCol)) {
        $t += (', col' + [char]0x00B7 + 'legiat n' + [char]0x00FA + 'mero ' + $numCol.Trim())
    }
    if (-not [string]::IsNullOrWhiteSpace($collegi)) { $t += (' del ' + $collegi.Trim()) }
    if (-not [string]::IsNullOrWhiteSpace($data)) { $t += (', en data ' + $data.Trim()) }
    return ($t + '.')
}

# Nom del fitxer de sortida. Segueix el mateix patro que la resta d'informes:
# data al principi (aixi "Actualitzar base d'informes" el reconeix).
# El nom NO porta el titular (l'usuari no el vol): data_fase_GIA. El titular ja
# surt a la capcalera del document.
function _LlicNomFitxer([datetime]$data, [string]$fase, [string]$idGia) {
    $curt = switch ($fase) {
        'favorable-pre'  { 'LlicFavPre' }
        'favorable-post' { 'LlicFavPost' }
        default          { 'LlicReq' }
    }
    return (_NomInformeFitxer $data $curt $idGia)
}

# ----------------------------------------------------------------------------
# COMPOSICIO DEL DOCUMENT (Word COM)
# ----------------------------------------------------------------------------
# Escriu un punt de Llicencia: el cos (de REQ1 o propi) com un item numerat, els
# seus sub-punts amb pic, i despres el comentari triat en VERD -negreta si falta
# la documentacio, sense negreta si ja hi es- i, al bloc DESPRES, el "Quan:".
#
# Tot el format surt de Format.ps1: aqui no s'hi inventa res. L'unic afegit es
# el color, que Format-Body ja sap aplicar.
# EL TEXT DEL "Quan:" segons la fase. Funcio PURA.
#
# Al favorable PRE va ENTRE PARENTESIS i sense el punt final:
# "(Quan: Abans d'iniciar l'activitat)" -decisio de l'usuari, setembre 2026-.
# Alla el punt no diu si es te o no el document (vegeu _LlicEstatDespres): el
# termini es nomes un aclariment del punt. A la resta de fases, com sempre.
function _LlicTextQuan([string]$quan, [string]$fase) {
    $q = ([string]$quan).Trim()
    if ([string]$fase -eq 'favorable-pre') { return ('(Quan: ' + $q.TrimEnd('.').TrimEnd() + ')') }
    return ('Quan: ' + $q)
}

function _LlicEscriuPunt($sel, $punt, [string]$marca, $fields, [string]$estat, [bool]$ambQuan, [string]$fase = '') {
    # ON VA L'ENLLAC. El comentari acaba dient "...en el seguent enllac:", o
    # sigui que l'enllac ha d'anar JUST DESPRES d'aquella frase. Pero el cos de
    # l'item (que ve de REQ1) sol portar EL MATEIX enllac, i sortia abans -amb
    # la frase penjada sense res al darrere-.
    #
    # Per aixo es miren PRIMER els enllacos del comentari: els que tambe son al
    # cos de l'item NO s'emeten amb l'item; s'esperen i surten despres del
    # comentari. Aixi no se'n repeteix cap i cada un queda on el text l'anuncia.
    #
    # ELS CAMPS ES RESOLEN PER BLOC (Apply-FieldsToLines, Camps.ps1) i NO linia a
    # linia: un [OPCIO:]/[CAMP:] pot ocupar dos paragrafs del cataleg, i llavors
    # cap de les dues linies en te un de sencer i el marcador sortia TAL QUAL al
    # Word. D'aqui avall les linies ja venen resoltes.
    $comLinies = if ($estat -eq 'si') { @(Apply-FieldsToLines $punt.SiDisposa $fields) }
                 elseif ($estat -eq 'no') { @(Apply-FieldsToLines $punt.NoDisposa $fields) }
                 else { @() }
    $urlsComentari = New-Object System.Collections.ArrayList
    foreach ($l in $comLinies) {
        foreach ($u in @((_SplitTextAndUrls ([string]$l)).Urls)) {
            $c = ([string]$u).Trim()
            if (-not $urlsComentari.Contains($c)) { [void]$urlsComentari.Add($c) }
        }
    }
    # Els URLs ja emesos en AQUEST punt (per no repetir-ne cap).
    $vistos = New-Object System.Collections.ArrayList
    foreach ($c in $urlsComentari) { [void]$vistos.Add($c) }
    $emesos = New-Object System.Collections.ArrayList
    $linies = @(Apply-FieldsToLines $punt.Cos $fields)
    $primera = if ($linies.Count -gt 0) { [string]$linies[0] } else { [string](Apply-Fields -text $punt.Titol -fields $fields) }
    # El numero i el text van junts a Format-Item; l'URL que porti la PRIMERA
    # linia s'emet a part, com fa REQ1 (_WriteCatalegBody).
    $p0 = _SplitTextAndUrls ([string]$primera)
    Format-Item $sel ([string]$marca) $p0.Text
    foreach ($u in @($p0.Urls)) {
        $c = ([string]$u).Trim()
        if ($vistos.Contains($c)) { continue }
        [void]$vistos.Add($c); [void]$emesos.Add($c); Format-Url $sel $u
    }
    for ($i = 1; $i -lt $linies.Count; $i++) {
        Write-Linia $sel ([string]$linies[$i]) $vistos $emesos
    }
    # Sub-punts (per exemple, quines instal·lacions s'han de legalitzar).
    #
    # L'ENLLAC D'UN SUB-PUNT SENSE TEXT NO VA SAGNAT. Un sub-punt que nomes
    # porta un enllac no penja de cap pic visible -no se n'ha arribat a emetre
    # cap-, o sigui que sagnar-lo el deixava despenjat un centimetre a la dreta
    # i sense res a sobre. Ho vaig veure comparant l'informe generat amb el fet
    # a ma: era l'unic dels vuit hiperenllacos que sortia sagnat.
    $primerSub = $true
    foreach ($sub in @($punt.Subs)) {
        $ambPic = $false
        foreach ($l in @(Apply-FieldsToLines $sub $fields)) {
            $pc = _SplitTextAndUrls ([string]$l)
            if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                if ($primerSub) { Format-Bullet $sel $pc.Text -IsChild -First; $primerSub = $false }
                else { Format-Bullet $sel $pc.Text -IsChild }
                $ambPic = $true
            }
            foreach ($u in @($pc.Urls)) {
                $c = ([string]$u).Trim()
                if ($vistos.Contains($c)) { continue }
                [void]$vistos.Add($c); [void]$emesos.Add($c)
                if ($ambPic) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
            }
        }
    }
    # EL "Quan:" VA ABANS DEL COMENTARI. A l'informe fet a ma, cada punt del
    # bloc DESPRES diu primer QUAN s'ha de tenir i despres SI ES TE o no; al
    # reves quedava el termini penjat al final del punt.
    if ($ambQuan) {
        foreach ($l in @(Apply-FieldsToLines $punt.Quan $fields)) {
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            Format-Body $sel (_LlicTextQuan ([string]$l) $fase)
        }
    }

    # El comentari. 'no' = falta (negreta); 'si' = ja hi es (normal). Al Word de
    # l'usuari anaven en verd, pero aquell color era una MARCA SEVA per saber que
    # havia de canviar a cada informe, no part del document: aqui van amb el
    # color de sempre (Format.ps1).
    #
    # I VA SEPARAT del cos del punt: -Separat hi posa al davant els mateixos
    # 12 pt que separen un item del seu primer sub-punt. A l'informe fet a ma
    # aquesta linia va separada als CINC punts, tant si al davant hi ha l'item
    # com si hi ha un enllac; nomes la PRIMERA linia del comentari.
    $primerCom = $true
    foreach ($l in $comLinies) {
        $pp = _SplitTextAndUrls ([string]$l)
        if (-not [string]::IsNullOrWhiteSpace($pp.Text)) {
            if ($primerCom -and $estat -eq 'no') { Format-Body $sel $pp.Text -Bold -Separat:$primerCom }
            else { Format-Body $sel $pp.Text -Separat:$primerCom }
            $primerCom = $false
        }
        # Aqui SI que s'emeten: es el lloc que la frase anuncia.
        foreach ($u in @($pp.Urls)) {
            $c = ([string]$u).Trim()
            if ($emesos.Contains($c)) { continue }
            [void]$emesos.Add($c); Format-Url $sel $u
        }
    }
}

# EL BLOC PROJECTE I LA DOCUMENTACIO DEL PROJECTE SON COMPLEMENTARIS, i les dues
# regles van juntes perque son la mateixa idea vista de les dues bandes:
#
#   REQUERIMENT   -> hi ha requeriments de projecte (bloc PROJECTE, amb lletres)
#                    i per tant el projecte i els planols encara s'han de
#                    modificar: la documentacio NO es definitiva i no hi surt.
#   FAVORABLES    -> ja no queden requeriments de projecte (el bloc no hi surt)
#                    i la documentacio ja es la bona: hi va, i dalt de tot.
#
# Funcions PURES, i la regla escrita en un sol lloc.
function _LlicPortaDocProjecte([string]$fase) {
    return ([string]$fase -eq 'favorable-pre' -or [string]$fase -eq 'favorable-post')
}

function _LlicPortaProjecte([string]$fase) {
    return ([string]$fase -eq 'requeriment')
}

# LA MARCA D'UN PUNT: "1." o "A.". Funcions PURES.
#
# EL BLOC PROJECTE VA AMB LLETRES i la resta amb numeros, i el motiu no es
# estetic: quan els requeriments de projecte queden resolts han de desapareixer
# de l'informe SENSE que la resta de la documentacio es renumeri. Amb tot
# numerat, el dia que el bloc PROJECTE marxa, l'"1." passa a ser una altra cosa
# i el titular no pot comparar-ho amb el que ja tenia.
#
# Passades les 26, segueix com les columnes de l'Excel: AA, AB... Aixi no hi ha
# cap topall amagat.
function _LlicLletra([int]$i) {
    if ($i -le 0) { return '' }
    $s = ''
    $n = $i
    while ($n -gt 0) {
        $n--
        $s = ([string][char](65 + ($n % 26))) + $s
        $n = [int][Math]::Floor($n / 26)
    }
    return $s
}

function _LlicMarca([int]$i, [string]$estil) {
    if ([string]$estil -eq 'lletra') { return ((_LlicLletra $i) + '.') }
    return ([string]$i + '.')
}

# ESCRIU UNA LLISTA DE PUNTS AMB LA SEVA ESTRUCTURA DE REQ1: la SECCIO en
# majuscules, la SUBSECCIO i el TEXT FIX que l'encapcala, i despres els punts.
#
# Abans els punts sortien tots seguits, sense saber de quina part de REQ1
# venien, i els textos fixos d'una subseccio (p.ex. "Segons l'article 3.1 de
# l'Ordenanca ... adjuntant la documentacio complementaria:") no sortien.
#
# La capcalera nomes s'escriu quan CANVIA, i nomes si hi ha algun punt a sota:
# com que penja de cada punt, una seccio sense punts no pot sortir.
#
# $n va per REFERENCIA: la numeracio de l'informe es SEGUIDA de cap a peus, no
# una llista nova per bloc.
function _LlicEscriuPunts($sel, $punts, [ref]$n, $fields, [bool]$ambQuan, [string]$estil = 'numero', [string]$fase = '') {
    $secAra = $null
    $subAra = $null
    $introAra = $null      # intro d'una SUBSECCIO: es reinicia a cada grup
    $introSecAra = $null   # intro de la SECCIO: nomes es reinicia amb la seccio
    foreach ($p in @($punts)) {
        $s  = [string]$p.Seccio
        $sb = [string]$p.Subseccio
        # L'INTRO SURT QUAN CANVIA, no quan canvia la subseccio: aixi surt amb el
        # PRIMER punt que s'escrigui del grup encara que no sigui el primer del
        # cataleg, i un grup amb dos intros els treu tots dos. Es la regla de
        # REQ1 (alli l'intro queda pendent fins que surt un item).
        #
        # NOMES si te text de debo: Apply-FieldsToLines pot tornar una linia
        # buida i llavors sortiria l'aire sense res al davant.
        $intro = @(@(Apply-FieldsToLines @($p.Intro) $fields) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $clauIntro = ($intro -join "`n")
        $deSeccio = [bool]$p.IntroDeSeccio

        if ($s -ne $secAra) {
            if (-not [string]::IsNullOrWhiteSpace($s)) {
                Format-Section $sel $s
                Format-Aire $sel 'seccio'
            }
            $secAra = $s
            $subAra = $null      # una seccio nova reinicia la subseccio...
            $introAra = $null    # ...i les dues intros pendents
            $introSecAra = $null
        }
        # L'intro de la SECCIO va ABANS del titol de la subseccio -com a REQ1- i
        # NO es reescriu a cada subseccio: introdueix tota la seccio. Sense el
        # tracker propi sortia tres vegades a Instal-lacions.
        if ($deSeccio -and $intro.Count -gt 0 -and $clauIntro -ne $introSecAra) {
            foreach ($l in $intro) { Write-Linia $sel ([string]$l) }
            Format-Aire $sel 'intro'
            $introSecAra = $clauIntro
        }
        if ($sb -ne $subAra) {
            if (-not [string]::IsNullOrWhiteSpace($sb)) {
                Format-Subsection $sel $sb
                Format-Aire $sel 'subseccio'
            }
            $subAra = $sb
            $introAra = $null    # grup nou: l'intro s'ha de tornar a escriure
        }
        if (-not $deSeccio -and $intro.Count -gt 0 -and $clauIntro -ne $introAra) {
            foreach ($l in $intro) { Write-Linia $sel ([string]$l) }
            Format-Aire $sel 'intro'
            $introAra = $clauIntro
        }
        $n.Value++
        _LlicEscriuPunt $sel $p (_LlicMarca $n.Value $estil) $fields ([string]$p.Estat) $ambQuan $fase
        Format-Aire $sel 'item'
    }
}

# Composa l'informe sencer i el desa. Retorna la ruta.
#
# $model porta tot el que ha triat l'usuari a l'assistent:
#   Fase, EsProvisional, Header, Fields, Abans, Projecte, Despres, Doc,
#   CondicionsActors (noms dels actors marcats), Orfes.
function Build-LlicenciaDocument($word, $model) {
    $header = $model.Header
    $baseName = _LlicNomFitxer (Get-Date) ([string]$model.Fase) ([string]$header['ID_GIA'])
    $cfg = $Script:ReportFormatConfig
    $fields = $model.Fields

    # La capcalera de LLICENCIA (porta la linia "Classificacio:"). Si el bloc no
    # hi es (0 CAPCALERA.docx encara sense actualitzar), Select-CapcaleraBlock es
    # queda amb el generic i l'informe surt igualment, sense la classificacio.
    return Write-InformeDocx $word $baseName 'LLIC' $header {
        param($sel)
        # ---- DOCUMENTACIO DEL PROJECTE ----
        # NOMES ALS FAVORABLES, i alli va al principi de tot.
        #
        # Al REQUERIMENT no hi va: si s'estan demanant modificacions al projecte
        # i als planols, aquella documentacio ENCARA NO ES DEFINITIVA i no te
        # sentit donar-la per bona. Quan ho sera -al favorable pre i al post- hi
        # surt, i dalt de tot. (Decisio de l'usuari, agost 2026.)
        if (_LlicPortaDocProjecte ([string]$model.Fase)) {
            $doc1 = [string]$model.Doc.Text
            if (-not [string]::IsNullOrWhiteSpace($doc1)) {
                Format-BlockTitle $sel ('DOCUMENTACI' + [char]0x00D3 + ' PROJECTE')
                Format-Aire $sel 'seccio'
                Format-Body $sel $doc1
                $primer = $true
                foreach ($d in @($model.Doc.Items)) {
                    if ($primer) { Format-Bullet $sel ([string]$d) -IsChild -First; $primer = $false }
                    else { Format-Bullet $sel ([string]$d) -IsChild }
                }
                Format-Aire $sel 'item'
            }
        }

        # ---- PROJECTE ----
        # NOMES AL REQUERIMENT (vegeu _LlicPortaProjecte): als favorables aquests
        # requeriments ja estan resolts i el bloc no te cap sentit.
        #
        # Va EL PRIMER de tot i amb LLETRES (A, B, C...). El motiu de les lletres
        # no es estetic: quan aquests requeriments quedin resolts, el bloc
        # desapareix i la resta de la documentacio ha de conservar la MATEIXA
        # numeracio. Amb tot numerat, l'"1." passaria a ser una altra cosa i el
        # titular no ho podria comparar amb el que ja tenia.
        #
        # Els espais els mana Format.ps1 (Format-Aire 'seccio' / 'subseccio' /
        # 'item'), exactament com _WriteCatalegBody de REQ1: aqui no s'hi inventa
        # cap separacio. I ELS TITOLS DE BLOC van amb Format-BlockTitle
        # (MAJUSCULES i subratllat): son el nivell de mes amunt de l'informe, per
        # sobre de les seccions de REQ1 que hi van a dins.
        $proj = if (_LlicPortaProjecte ([string]$model.Fase)) { @($model.Projecte) } else { @() }
        if ($proj.Count -gt 0) {
            Format-BlockTitle $sel (_LlicTitolProjecte)
            Format-Aire $sel 'seccio'
            $lletra = 0
            _LlicEscriuPunts $sel $proj ([ref]$lletra) $fields $false 'lletra'
        }

        # ---- ABANS ----
        Format-BlockTitle $sel (_LlicTitolAbans)
        Format-Aire $sel 'seccio'
        $n = 0
        _LlicEscriuPunts $sel @($model.Abans) ([ref]$n) $fields $false

        # ---- DESPRES ----
        # LA NUMERACIO CONTINUA la del bloc ABANS ($n NO es reinicia): els punts
        # numerats van seguits de cap a peus. El bloc PROJECTE no hi compta: te
        # el seu comptador de lletres.
        $desp = @($model.Despres)
        if ($desp.Count -gt 0) {
            Format-BlockTitle $sel (_LlicTitolDespres)
            Format-Aire $sel 'seccio'
            _LlicEscriuPunts $sel $desp ([ref]$n) $fields $true 'numero' ([string]$model.Fase)
        }

        # ---- CONCLUSIO ----
        $actorsCond = @(@($model.CondicionsActors) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $ambCond = ($actorsCond.Count -gt 0 -and (_LlicAdmetCondicions ([string]$model.Fase)))
        # Mateix bloc que REQ1 (_WriteConclusionsBlock): capcalera CONCLUSIONS
        # centrada i en negreta, i la conclusio en negreta -que aqui ve del **...**
        # del cataleg, exactament com a REQ1: el text ja no es del codi.
        Format-Aire $sel 'conclusions'
        Format-ConclusionHeader $sel 'CONCLUSIONS'
        Format-Conclusion $sel (_LlicConclusioText ([string]$model.Fase) $ambCond)
        # ELS ACTORS, sota la conclusio que els anuncia ("...en els seguents
        # informes (adjunts a continuacio):"). Amb LLETRES minuscules com al Word
        # de l'usuari (a., b., c.), escrites com a TEXT -com tota la numeracio
        # d'aquests informes-. Les condicions no s'escriuen: son als informes
        # d'aquells organismes, que van adjunts darrere.
        if ($ambCond) {
            for ($i = 0; $i -lt $actorsCond.Count; $i++) {
                Format-Item $sel ((_LlicLletra ($i + 1)).ToLower() + '.') ([string]$actorsCond[$i])
                Format-Aire $sel 'item'
            }
        }
        # El tancament: del cataleg, com tots els altres informes (Write-Tancament).
        Write-Tancament $sel $fields

        # ---- ANNEX 1: nomes al REQUERIMENT d'una llicencia PROVISIONAL ----
        if ([string]$model.Fase -eq 'requeriment' -and (_LlicCalAnnex1 $model.Abans ([bool]$model.EsProvisional))) {
            _LlicEscriuAnnex1 $sel $model.Cataleg
        }
    }
}

# L'ANNEX 1, que nomes va al REQUERIMENT d'una llicencia provisional. El text es
# FIX i viu al cataleg (seccio que comenca per "ANNEX 1"), no encastat aqui:
# aixi l'usuari el pot editar des de l'editor de catalegs com tota la resta.
function _LlicSeccioAnnex1($llic) {
    foreach ($s in @($llic.nodes)) {
        if ([string]$s.titol -like 'ANNEX 1*') { return $s }
    }
    return $null
}

# L'ANNEX 1 va en TEXT PLA: sense sagnies, sense pics i sense numeracio (a la
# plantilla de l'usuari es text corrent, encara que alli el Word hi tingues una
# llista). NOMES van en negreta els dos TITOLS:
#   - "ANNEX 1. Documentacio per demanar..."
#   - "Document d'acceptacio del cessament dels usos..."
# I des del "Document d'acceptacio..." fins al final: PAGINA NOVA i cos 9 (a la
# plantilla, sz=18 mig-punts).
# Afegeix text AL PARAGRAF QUE S'ACABA D'ESCRIURE, separat per un espai. Els
# Format-* comencen sempre amb TypeParagraph, o sigui que n'obren un de nou;
# aqui volem continuar el mateix (a la plantilla, l'aclariment d'un punt de
# l'ANNEX 1 va dins del punt, no en un paragraf a part).
function _LlicAfegeixAlParagraf($sel, [string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    Format-Append $sel (' ' + $text)
}

function _LlicEscriuAnnex1($sel, $llic) {
    $sec = _LlicSeccioAnnex1 $llic
    if ($null -eq $sec) { return }
    # Salt de pagina: l'annex es un document a part dins de l'informe.
    Format-SaltPagina $sel
    Format-Plain $sel ([string]$sec.titol) -Bold

    $cos9 = $false          # ja som al full de signatures?
    $num = 0                # el comptador dels punts numerats
    $primerItem = $true     # el primer punt no porta linia en blanc al davant
    $obreBloc = $false      # al full de signatures, comenca una declaracio nova
    foreach ($nd in @($sec.fills)) {
        # LA MARCA VA COM A TEXT, no com a llista del Word: la plantilla la porta
        # amb numeracio automatica i sagnia, i l'usuari la vol PLANA (nomes el
        # numero o el guio escrits al davant). El comptador NOMES avanca amb els
        # 'item': els 'text' del mig no es numeren.
        $marca = ''
        $tip = [string]$nd.tipus
        if ($tip -eq 'item') { $num++; $marca = [string]$num + '. ' }
        elseif ($tip -eq 'subitem') { $marca = '- ' }
        # UNA LINIA EN BLANC entre punts numerats (a l'informe fet a ma n'hi ha
        # una, i sense ella els quatre punts sortien arrapats).
        if ($tip -eq 'item' -and -not $primerItem -and -not $cos9) { Format-Spacer $sel }
        if ($tip -eq 'item') { $primerItem = $false }
        $primeraLinia = $true
        foreach ($l in @(_LlicCos $nd)) {
            $pp = _SplitTextAndUrls ([string]$l)
            $t = [string]$pp.Text
            if (-not [string]::IsNullOrWhiteSpace($t)) {
                $esTitolAcceptacio = (_LlicEsTitolAcceptacio $t)
                if ($esTitolAcceptacio -and -not $cos9) {
                    # A partir d'aqui, full a part i lletra mes petita.
                    Format-SaltPagina $sel
                    $cos9 = $true
                    $obreBloc = $false
                }
                if ($cos9) {
                    # EL FULL DE SIGNATURES: dues linies en blanc davant de cada
                    # declaracio (despres del titol i despres de cada "Signat"),
                    # que es com esta a la plantilla.
                    if ($obreBloc) { Format-Spacer $sel; Format-Spacer $sel; $obreBloc = $false }
                    if ($esTitolAcceptacio) { Format-Plain $sel $t -Bold -Size $Script:LlicAnnexSignaturaCos }
                    else                    { Format-Plain $sel $t -Size $Script:LlicAnnexSignaturaCos }
                    if ($esTitolAcceptacio -or $t.Trim() -eq 'Signat') { $obreBloc = $true }
                    $primeraLinia = $false
                }
                elseif ($tip -eq 'text' -and -not $primerItem) {
                    # UN NODE 'text' ES LA CONTINUACIO del punt de sobre, no un
                    # paragraf nou: a la plantilla va DINS del mateix paragraf,
                    # separat per un espai. Anava a part i partia el punt en dos.
                    _LlicAfegeixAlParagraf $sel $t
                }
                else {
                    # La marca nomes a la PRIMERA linia del punt (un punt pot
                    # tenir mes d'un paragraf de cos).
                    $txt = if ($primeraLinia) { $marca + $t } else { $t }
                    Format-Plain $sel $txt
                    $primeraLinia = $false
                }
            }
            foreach ($u in @($pp.Urls)) { Format-Url $sel $u }
        }
    }
}

# ----------------------------------------------------------------------------
# ASSISTENT (WinForms, nomes Windows)
# ----------------------------------------------------------------------------
# Pas 1: la FASE i si es llicencia provisional. Retorna @{ Nav; Fase; Prov }.
# $fases: quines fases s'ofereixen. Des que la Modificacio NO Substancial i el
# Traspas tenen entrada propia al menu, cada familia ensenya NOMES les seves:
# _LlicFases per a Llicencia i _MnsFases per a MNS/Traspas. Amb $null les
# ensenya totes (compatibilitat).
function Select-LlicFase($preFase, $preProv, $fases = $null, [string]$titol = '') {
    $llista = if ($null -ne $fases) { @($fases) } else { @(_LlicTotesLesFases) }
    $form = _NewForm
    $form.Text = if ($titol) { $titol } else { 'Llic' + [char]0x00E8 + 'ncia - Pas 1' }
    $form.ClientSize = New-Object System.Drawing.Size(520, 330)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 72)
    $lbl.Size = New-Object System.Drawing.Size(480, 20)
    $lbl.Text = 'Quin informe vols fer?'
    [void]$form.Controls.Add($lbl)

    $y = 98
    $radios = @{}
    foreach ($f in @($llista)) {
        $rb = New-Object System.Windows.Forms.RadioButton
        $rb.Location = New-Object System.Drawing.Point(30, $y)
        $rb.Size = New-Object System.Drawing.Size(460, 22)
        $rb.Text = [string]$f.Nom
        $rb.Checked = ([string]$preFase -eq [string]$f.Clau)
        [void]$form.Controls.Add($rb)
        $sub = New-Object System.Windows.Forms.Label
        $sub.Location = New-Object System.Drawing.Point(50, ($y + 21))
        $sub.Size = New-Object System.Drawing.Size(440, 18)
        $sub.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
        $sub.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $sub.Text = [string]$f.Sub
        [void]$form.Controls.Add($sub)
        $radios[[string]$f.Clau] = $rb
        $y += 46
    }
    if (-not ($radios.Values | Where-Object { $_.Checked })) {
        $perDefecte = _LlicFasePerDefecte $llista $preFase
        if ($radios.ContainsKey($perDefecte)) { $radios[$perDefecte].Checked = $true }
    }

    $cbProv = New-Object System.Windows.Forms.CheckBox
    $cbProv.Location = New-Object System.Drawing.Point(30, ($y + 6))
    $cbProv.AutoSize = $true
    $cbProv.Text = 'Llic' + [char]0x00E8 + 'ncia provisional'
    $cbProv.Checked = [bool]$preProv
    [void]$form.Controls.Add($cbProv)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Location = New-Object System.Drawing.Point(50, ($y + 28))
    $lbl2.Size = New-Object System.Drawing.Size(450, 32)
    $lbl2.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lbl2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbl2.Text = ('Canvia el punt de compatibilitat (AMB en lloc d' + [char]0x2019 + 'Annex II) i, al requeriment, ' +
                  'hi afegeix l' + [char]0x2019 + 'ANNEX 1.')
    [void]$form.Controls.Add($lbl2)

    # La casella "Llicencia provisional" nomes te sentit a l'informe llarg: als
    # dos curts no canvia res del document, i deixar-la activa nomes despista.
    #
    # VA AQUI I NO MES AMUNT: .GetNewClosure() copia els VALORS del moment, o
    # sigui que una closure creada abans de $cbProv i $lbl2 se'ls quedaria a
    # $null (vegeu CLAUDE.md). I un clic en un radio dispara DOS esdeveniments
    # -el que es marca i el germa que es desmarca-, pero aqui es idempotent.
    $fnFase = @{}
    $fnFase.Refresca = {
        $curta = $false
        foreach ($k in @($radios.Keys)) { if ($radios[$k].Checked -and (_MnsEsFase $k)) { $curta = $true } }
        $cbProv.Enabled = (-not $curta)
        $lbl2.Visible = (-not $curta)
    }.GetNewClosure()
    foreach ($k in @($radios.Keys)) {
        $radios[$k].add_CheckedChanged({ & $fnFase.Refresca }.GetNewClosure())
    }
    & $fnFase.Refresca

    # ELS BOTONS, SOTA L'ULTIMA ETIQUETA. Estaven clavats a y=286 i la nota de
    # la llicencia provisional (y=264, alt 32) els trepitjava. Ara surten del
    # peu real de $lbl2, o sigui que si hi afegim una fase o una linia de text
    # baixen sols i la finestra creix amb ells.
    $yBotons = $lbl2.Bottom + 14
    $form.ClientSize = New-Object System.Drawing.Size(520, ($yBotons + 32 + 16))

    $res = @{ Nav = 'back' }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(370, $yBotons)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        foreach ($k in $radios.Keys) { if ($radios[$k].Checked) { $res.Fase = $k } }
        $res.Prov = [bool]$cbProv.Checked
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, $yBotons)
    $btnBack.Size = New-Object System.Drawing.Size(110, 32)
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form ('Llic' + [char]0x00E8 + 'ncia (Annex II / LL Prov)') 'Tria quin informe vols fer' 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de tria de DOCUMENTACIO (blocs ABANS i DESPRES).
#
# ARBRE a l'esquerra + DETALL a la dreta, el mateix aspecte que el Pas 3
# (Select-Items, SeleccioItems.ps1): les seccions en negreta i els punts a
# dins. Abans era una llista plana amb 40 punts a la mateixa alcada, i abans
# encara una graella amb un boto "Omplir..." que obria un dialeg -que no
# s'assemblava a com s'omplen els camps a la resta del programa-.
#
# L'AGRUPACIO ES NOMES DE PANTALLA (_LlicAgrupaPunts): els punts es recorren en
# l'ordre del cataleg per muntar l'informe, o sigui que agrupar no en canvia
# l'ordre. Els punts sense seccio (els PROPIS, i els que es llegeixen d'un
# informe ja emes) van al primer nivell, sense capcalera.
#
# Al detall hi ha, segons el bloc:
#   - la tria "No es disposa / Es disposa" ($ambEstat);
#   - la frase del cataleg amb els [CAMP: ...] INLINE ($ambDades, nomes ABANS),
#     renderitzada amb _RenderRichInto (Camps.ps1) -la MATEIXA funcio que fa
#     servir REQ1-;
#   - les caselles dels SUB-PUNTS ($ambSubs, nomes DESPRES): els certificats
#     d'inscripcio i les inspeccions inicials no els te tothom.
#
# ELS CAMPS VAN PER PUNT, no al diccionari compartit: "Id Firmadoc" val una cosa
# diferent a cada document.
#
# $marcatPerDefecte: si els punts surten ja marcats. Al bloc DESPRES si (el Word
# de l'usuari els portava tots i ell hi anava esborrant el que no tocava; picar
# quinze caselles cada vegada era feina de mes), i al bloc ABANS no, perque alli
# cada punt demana a mes decidir si es te la documentacio o no.
#
# Retorna @{ Nav; Punts; Memoria }.
function Select-LlicDocumentacio($punts, [string]$titol, [string]$subtitol, [bool]$ambEstat,
                                 [bool]$marcatPerDefecte = $false, [bool]$ambDades = $false,
                                 $preSel = $null, [bool]$ambSubs = $false,
                                 [string]$estatPerDefecte = 'no') {
    $punts = @($punts)
    $grups = @(_LlicAgrupaPunts $punts)

    # Estat de cada punt (viu tota la pantalla i es el que es retorna).
    $st = @{}
    for ($i = 0; $i -lt $punts.Count; $i++) {
        $p = $punts[$i]
        $clau = _LlicClauPunt $p
        # Camps  = els objectes de camp VIUS de la pantalla (els fa _RenderRichInto).
        # Valors = el mapa pla nom -> valor, que es el que es RECORDA i es desa a
        #          la base de dades. Els objectes de camp no sobreviuen un pas per
        #          JSON; el mapa pla si, i _RenderRichInto ja el sap llegir com a
        #          $preload (_GetPreloadValue, Camps.ps1).
        $ini = if ([string]::IsNullOrWhiteSpace($estatPerDefecte)) { 'no' } else { [string]$estatPerDefecte }
        # EstatPrevi: el que deia la memoria, tal qual. Una pantalla que NO
        # pregunta l'estat (requeriment, favorable pre) el torna a desar com
        # l'ha trobat: si hi desava el seu 'no' per defecte, el POST sortia amb
        # "No es disposa" marcat a tots els punts en lloc del seu 'si'.
        $e = @{ Marcat = $marcatPerDefecte; Estat = $ini; EstatPrevi = ''; Camps = [ordered]@{}; Valors = @{}; Subs = @{} }
        if ($null -ne $preSel -and $preSel.Contains($clau)) {
            $e.Marcat = [bool]$preSel[$clau].Marcat
            $e.EstatPrevi = [string]$preSel[$clau].Estat
            # Buit = la pantalla d'on ve no el preguntava: mana el de la fase.
            if (-not [string]::IsNullOrWhiteSpace($e.EstatPrevi)) { $e.Estat = $e.EstatPrevi }
            if ($null -ne $preSel[$clau].Valors) { $e.Valors = $preSel[$clau].Valors }
            if ($null -ne $preSel[$clau].Subs)   { $e.Subs   = $preSel[$clau].Subs }
        }
        # Per defecte, TOTS els sub-punts d'un punt marcat entren.
        foreach ($k in 0..([Math]::Max(0, @($p.Subs).Count - 1))) {
            if (-not $e.Subs.Contains($k)) { $e.Subs[$k] = $true }
        }
        $st[$i] = $e
    }

    # LES FUNCIONS DE LA PANTALLA, TOTES DINS D'UN HASHTABLE.
    #
    # PER QUE: .GetNewClosure() copia el VALOR de les variables en el moment de
    # crear el scriptblock. Un scriptblock que es cridi a si mateix (o que
    # cridi un que encara no existeix) es quedaria amb $null i peta amb
    #   "L'expressio que segueix a & ... no es un nom d'ordre ni un scriptblock".
    # El hashtable, en canvi, es captura per REFERENCIA: $fn.Pinta es resol en
    # cridar-lo i l'ordre de definicio deixa d'importar.
    # Hi ha una prova que ho vigila a run-tests.ps1 ("cap closure es refereix a
    # si mateixa"); no tornis a fer $x = { ... & $x ... }.GetNewClosure().
    $fn = @{}
    $estatUi = @{ Busy = $false }

    $form = _NewForm
    $form.Text = $titol
    $form.ClientSize = New-Object System.Drawing.Size(1080, 660)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(820, 520)

    # ---- Esquerra: cercador + ARBRE amb caselles ---------------------------
    $panEsq = New-Object System.Windows.Forms.Panel
    $panEsq.Location = New-Object System.Drawing.Point(14, 66)
    $panEsq.Size = New-Object System.Drawing.Size(500, 520)
    $panEsq.Anchor = 'Top,Bottom,Left'
    [void]$form.Controls.Add($panEsq)

    # Mateix aspecte que el Pas 3 (Select-Items): seccions en negreta i punts a
    # dins. El font BASE es la negreta mes ampla, si no WinForms retalla els
    # nodes que tenen un NodeFont mes ample que el del control.
    $arbre = New-Object System.Windows.Forms.TreeView
    $arbre.Location = New-Object System.Drawing.Point(0, 28)
    $arbre.Size = New-Object System.Drawing.Size(500, 492)
    $arbre.Anchor = 'Top,Bottom,Left,Right'
    $arbre.CheckBoxes = $true
    $arbre.HideSelection = $false
    $arbre.ShowNodeToolTips = $true
    $arbre.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    [void]$panEsq.Controls.Add($arbre)

    # ---- Dreta: detall del punt seleccionat --------------------------------
    $panDret = New-Object System.Windows.Forms.Panel
    $panDret.Location = New-Object System.Drawing.Point(526, 66)
    $panDret.Size = New-Object System.Drawing.Size(540, 520)
    $panDret.Anchor = 'Top,Bottom,Left,Right'
    $panDret.AutoScroll = $true
    $panDret.BorderStyle = 'FixedSingle'
    $panDret.BackColor = [System.Drawing.Color]::White
    [void]$form.Controls.Add($panDret)

    # Reconstrueix l'arbre segons el filtre. L'estat de les caselles NO viu a
    # l'arbre sino a $st: aixi el filtre no en pot perdre cap.
    $fn.Omple = {
        param($filtre)
        $estatUi.Busy = $true
        $arbre.BeginUpdate()
        try {
            $arbre.Nodes.Clear()
            $f = ([string]$filtre).Trim()
            # DOS NIVELLS, com el Pas 3: seccio en negreta i, si en te, la
            # subseccio subratllada a dins. El node de seccio es REAPROFITA
            # entre subseccions consecutives de la mateixa seccio.
            $secAra = [char]0x0001   # cap seccio encara (no pot coincidir amb res)
            $nodeSec = $null
            foreach ($g in $grups) {
                $secTit = [string]$g.Titol
                $subTit = [string]$g.Sub
                $secMatch = ((_TextMatches $secTit $f) -or (_TextMatches $subTit $f))
                # Els punts del grup que passen el filtre.
                $visibles = New-Object System.Collections.ArrayList
                foreach ($i in @($g.Idx)) {
                    $et = _LlicEtiquetaPunt $punts[$i]
                    if ($secMatch -or (_TextMatches $et $f)) { [void]$visibles.Add($i) }
                }
                if ($visibles.Count -eq 0) { continue }

                # Grup sense titol = primer nivell, sense capcalera.
                $pare = $null
                if (-not [string]::IsNullOrWhiteSpace($secTit)) {
                    if ($secTit -ne $secAra) {
                        $secAra = $secTit
                        $nodeSec = New-Object System.Windows.Forms.TreeNode($secTit)
                        $nodeSec.Tag = @{ Kind = 'Section' }
                        $nodeSec.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
                        [void]$arbre.Nodes.Add($nodeSec)
                    }
                    $pare = $nodeSec
                    if (-not [string]::IsNullOrWhiteSpace($subTit)) {
                        $nodeSub = New-Object System.Windows.Forms.TreeNode($subTit)
                        $nodeSub.Tag = @{ Kind = 'Section' }
                        $nodeSub.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
                        [void]$nodeSec.Nodes.Add($nodeSub)
                        $pare = $nodeSub
                    }
                }
                $totsMarcats = $true
                foreach ($i in $visibles) {
                    $nd = New-Object System.Windows.Forms.TreeNode((_LlicEtiquetaPunt $punts[$i]))
                    $nd.Tag = @{ Kind = 'Item'; Idx = $i }
                    $nd.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                    $nd.ToolTipText = (_LlicTextPlaDelCos $punts[$i].Cos)
                    if ($nd.ToolTipText.Length -gt 600) { $nd.ToolTipText = $nd.ToolTipText.Substring(0, 600) + '...' }
                    $nd.Checked = [bool]$st[$i].Marcat
                    if (-not $nd.Checked) { $totsMarcats = $false }
                    if ($null -eq $pare) { [void]$arbre.Nodes.Add($nd) } else { [void]$pare.Nodes.Add($nd) }
                }
                if ($null -ne $pare) {
                    $pare.Checked = $totsMarcats
                    $pare.Expand()
                    if ($null -ne $pare.Parent) { $pare.Parent.Expand() }
                }
            }
        } finally {
            $arbre.EndUpdate()
            $estatUi.Busy = $false
        }
    }.GetNewClosure()

    $cerca = _AddSearchBox $panEsq 0 2 380 'Cerca:' {
        param($sender, $ev)
        & $fn.Omple $sender.Text
    }.GetNewClosure()

    $fn.Pinta = {
        param($idx)
        $panDret.Controls.Clear()
        if ($null -eq $idx -or $idx -lt 0) { return }
        $p = $punts[$idx]
        $e = $st[$idx]
        $y = 10
        # UN REGISTRE DE CAMPS NOU A CADA PINTADA, mai un de tota la pantalla.
        #
        # El registre (Camps.ps1) SINCRONITZA els controls que porten el MATEIX
        # nom de camp: escriure en un "Id Firmadoc" copia el text a tots els
        # altres "Id Firmadoc" que hi hagi registrats. Es el que vol REQ1, on un
        # camp val el mateix a tot l'informe. Aqui NO: cada punt es un document
        # diferent. Amb un registre de tota la pantalla, els quadres dels punts
        # ja visitats (trets del panell pero vius, amb el seu handler) rebien
        # el text del punt nou i l'escrivien al SEU punt: l'informe del GIA 924
        # va sortir amb el mateix Id Firmadoc (9887463) als cinc punts d'ABANS.
        # Hi ha un guard que ho vigila (06-guards.ps1).
        $fldRegistry = _NewFieldRegistry

        # El text sencer del punt.
        $lbT = New-Object System.Windows.Forms.Label
        $lbT.Location = New-Object System.Drawing.Point(10, $y)
        $lbT.MaximumSize = New-Object System.Drawing.Size(495, 0)
        $lbT.AutoSize = $true
        $lbT.Text = (_LlicTextPlaDelCos $p.Cos)
        if ([string]::IsNullOrWhiteSpace($lbT.Text)) { $lbT.Text = (_LlicEtiquetaPunt $p 0) }
        [void]$panDret.Controls.Add($lbT)
        $y += [Math]::Max(24, $lbT.PreferredHeight + 10)

        if ($ambEstat) {
            $rbNo = New-Object System.Windows.Forms.RadioButton
            $rbNo.Location = New-Object System.Drawing.Point(10, $y)
            $rbNo.AutoSize = $true
            $rbNo.Text = 'No es disposa del document'
            $rbNo.Checked = ([string]$e.Estat -ne 'si')
            [void]$panDret.Controls.Add($rbNo)
            $y += 24
            $rbSi = New-Object System.Windows.Forms.RadioButton
            $rbSi.Location = New-Object System.Drawing.Point(10, $y)
            $rbSi.AutoSize = $true
            $rbSi.Text = 'Es disposa del document'
            $rbSi.Checked = ([string]$e.Estat -eq 'si')
            [void]$panDret.Controls.Add($rbSi)
            $y += 30
            # UNA COPIA LOCAL DE $fn.
            #
            # .GetNewClosure() nomes copia els LOCALS del context que la crida.
            # Aqui dins, $idx, $e i els dos radios SI que ho son, pero $fn ve del
            # modul de la closure de fora i arribaria als handlers com a $null
            # (-> "& $null.Pinta", el quadre d'error en triar "Es disposa").
            # Hi ha una prova que ho vigila; vegeu CLAUDE.md.
            $fnAquest = $fn
            # UN HANDLER PER RADIO, i nomes actua el que s'acaba de marcar: un
            # sol clic dispara DOS esdeveniments -el que es marca i el germa que
            # es desmarca- i amb un handler compartit la pantalla es repintava
            # dues vegades, la segona llegint uns controls que Controls.Clear()
            # acabava de treure del panell.
            $rbSi.add_CheckedChanged({
                if (-not $rbSi.Checked) { return }
                $e.Estat = 'si'
                & $fnAquest.Pinta $idx
            }.GetNewClosure())
            $rbNo.add_CheckedChanged({
                if (-not $rbNo.Checked) { return }
                $e.Estat = 'no'
                & $fnAquest.Pinta $idx
            }.GetNewClosure())
        }

        # La frase del cataleg amb els camps INLINE (nomes al bloc ABANS).
        if ($ambDades) {
            # TOT EL BLOC JUNT, no linia a linia: un [CAMP:]/[OPCIO:] pot ocupar
            # dos paragrafs del cataleg, i la pantalla ha de veure el mateix
            # text que el generador (que resol per bloc, Apply-FieldsToLines).
            $linies = if ([string]$e.Estat -eq 'si') { @($p.SiDisposa) } else { @($p.NoDisposa) }
            $linies = @(($linies -join [char]10))
            foreach ($l in $linies) {
                if ([string]::IsNullOrWhiteSpace($l)) { continue }
                $flow = New-Object System.Windows.Forms.FlowLayoutPanel
                $flow.Location = New-Object System.Drawing.Point(10, $y)
                $flow.Size = New-Object System.Drawing.Size(500, 10)
                $flow.AutoSize = $true
                $flow.AutoSizeMode = 'GrowAndShrink'
                $flow.MaximumSize = New-Object System.Drawing.Size(500, 0)
                $flow.WrapContents = $true
                $flow.FlowDirection = 'LeftToRight'
                [void]$panDret.Controls.Add($flow)
                # LA MATEIXA funcio que REQ1, amb un diccionari PER PUNT i amb
                # els valors recordats com a $preload: aixi els Id Firmadoc i
                # els expedients de l'informe anterior ja surten escrits.
                _RenderRichInto $flow ([string]$l) $e.Camps $e.Valors $fldRegistry
                $y += [Math]::Max(26, $flow.PreferredSize.Height + 8)
            }
        }

        # Els SUB-PUNTS (nomes al bloc DESPRES): no tothom els te tots.
        if ($ambSubs -and @($p.Subs).Count -gt 0) {
            $lbS = New-Object System.Windows.Forms.Label
            $lbS.Location = New-Object System.Drawing.Point(10, $y)
            $lbS.AutoSize = $true
            $lbS.Text = 'Quins hi entren:'
            [void]$panDret.Controls.Add($lbS)
            $y += 22
            for ($k = 0; $k -lt @($p.Subs).Count; $k++) {
                $sub = @($p.Subs)[$k]
                $txtSub = (@($sub) -join ' ').Trim()
                if ([string]::IsNullOrWhiteSpace($txtSub)) { continue }
                $cb = New-Object System.Windows.Forms.CheckBox
                $cb.Location = New-Object System.Drawing.Point(24, $y)
                $cb.MaximumSize = New-Object System.Drawing.Size(470, 0)
                $cb.AutoSize = $true
                $cb.Text = $txtSub
                $cb.Checked = [bool]$e.Subs[$k]
                $kk = $k
                $cb.add_CheckedChanged({ $e.Subs[$kk] = [bool]$cb.Checked }.GetNewClosure())
                [void]$panDret.Controls.Add($cb)
                $y += [Math]::Max(24, $cb.PreferredHeight + 4)
            }
        }
    }.GetNewClosure()

    # Marcar una SECCIO marca tots els seus punts (com al Pas 3).
    $arbre.add_AfterCheck({
        param($sender, $ev)
        if ($estatUi.Busy) { return }
        $estatUi.Busy = $true
        try {
            $tag = $ev.Node.Tag
            if ($null -ne $tag -and [string]$tag.Kind -eq 'Section') {
                # Baixa per tot l'arbre: una seccio pot tenir subseccions.
                $pila = New-Object System.Collections.ArrayList
                [void]$pila.Add($ev.Node)
                while ($pila.Count -gt 0) {
                    $nd = $pila[0]; [void]$pila.RemoveAt(0)
                    foreach ($fill in $nd.Nodes) {
                        $fill.Checked = $ev.Node.Checked
                        if ($null -ne $fill.Tag -and [string]$fill.Tag.Kind -eq 'Item') {
                            $st[[int]$fill.Tag.Idx].Marcat = [bool]$ev.Node.Checked
                        } else { [void]$pila.Add($fill) }
                    }
                }
            } elseif ($null -ne $tag -and [string]$tag.Kind -eq 'Item') {
                $st[[int]$tag.Idx].Marcat = [bool]$ev.Node.Checked
                # La casella de la seccio segueix els seus fills.
                $pare = $ev.Node.Parent
                while ($null -ne $pare) {
                    $tots = $true
                    foreach ($fill in $pare.Nodes) { if (-not $fill.Checked) { $tots = $false; break } }
                    $pare.Checked = $tots
                    $pare = $pare.Parent
                }
            }
        } finally { $estatUi.Busy = $false }
    }.GetNewClosure())

    $arbre.add_AfterSelect({
        param($sender, $ev)
        if ($estatUi.Busy) { return }
        $tag = $ev.Node.Tag
        if ($null -eq $tag -or [string]$tag.Kind -ne 'Item') { $panDret.Controls.Clear(); return }
        & $fn.Pinta ([int]$tag.Idx)
    }.GetNewClosure())

    & $fn.Omple ''
    if ($arbre.Nodes.Count -gt 0) {
        $primer = $arbre.Nodes[0]
        if ($null -ne $primer.Tag -and [string]$primer.Tag.Kind -ne 'Item' -and $primer.Nodes.Count -gt 0) {
            $primer = $primer.Nodes[0]
        }
        $arbre.SelectedNode = $primer
    }

    # ---- Botons -----------------------------------------------------------
    $res = @{ Nav = 'back'; Punts = @(); Memoria = $null }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(941, 606)
    $btnOk.Size = New-Object System.Drawing.Size(125, 34)
    $btnOk.Anchor = 'Bottom,Right'
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        $sel = New-Object System.Collections.ArrayList
        $mem = @{}
        # EN L'ORDRE DEL CATALEG, no el de l'arbre: agrupar es NOMES de pantalla.
        for ($i = 0; $i -lt $punts.Count; $i++) {
            $p = $punts[$i]
            $e = $st[$i]
            $clau = _LlicClauPunt $p
            # El registre de Camps.ps1 desa objectes amb .Value; aqui en volem un
            # mapa nom -> valor, que es el que es recorda i el que es desa.
            # S'HI CONSERVA el que ja hi havia: si un punt no s'ha arribat a
            # pintar (no s'hi ha clicat mai), $e.Camps es buit i els valors
            # recuperats de la base es perdrien.
            $vals = @{}
            foreach ($k in @($e.Valors.Keys)) { $vals[[string]$k] = [string]$e.Valors[$k] }
            foreach ($k in @($e.Camps.Keys))  { $vals[[string]$k] = [string]$e.Camps[$k].Value }
            $e.Valors = $vals
            # Si aquesta pantalla no pregunta l'estat, ni el recorda ni el diu:
            # es desa el que hi havia i el punt surt SENSE estat (cap "No es
            # disposa..." ni "Es disposa..." a l'informe, digui el que digui el
            # cataleg).
            $estatMem = if ($ambEstat) { [string]$e.Estat } else { [string]$e.EstatPrevi }
            $estatPunt = if ($ambEstat) { [string]$e.Estat } else { '' }
            $mem[$clau] = @{ Marcat = $e.Marcat; Estat = $estatMem; Valors = $vals; Subs = $e.Subs }
            if (-not $e.Marcat) { continue }
            $si = @($p.SiDisposa); $no = @($p.NoDisposa)
            if ($estatPunt -eq 'si') { $si = @(_LlicAplicaCamps $p.SiDisposa $vals) }
            else                     { $no = @(_LlicAplicaCamps $p.NoDisposa $vals) }
            # Nomes els sub-punts triats.
            $subs = New-Object System.Collections.ArrayList
            for ($k = 0; $k -lt @($p.Subs).Count; $k++) {
                if ($ambSubs -and -not [bool]$e.Subs[$k]) { continue }
                [void]$subs.Add(@($p.Subs)[$k])
            }
            # 'Select-Object *' i no una llista de camps: el punt porta tambe la
            # SECCIO, la SUBSECCIO i l'INTRO de REQ1, i enumerar-los aqui vol dir
            # que el dia que se n'afegeixi un es perdi en silenci. Ja va passar.
            $c = $p | Select-Object *
            $c.NoDisposa = $no
            $c.SiDisposa = $si
            $c.Subs = $subs.ToArray()
            $c | Add-Member NoteProperty Estat $estatPunt -Force
            [void]$sel.Add($c)
        }
        $res.Punts = $sel.ToArray()
        $res.Memoria = $mem
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(14, 606)
    $btnBack.Size = New-Object System.Drawing.Size(115, 34)
    $btnBack.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    $fn.MarcaTot = {
        param($valor)
        for ($i = 0; $i -lt $punts.Count; $i++) { $st[$i].Marcat = $valor }
        & $fn.Omple $cerca.Text
    }.GetNewClosure()
    $btnTot = New-Object System.Windows.Forms.Button
    $btnTot.Text = 'Marcar-ho tot'
    $btnTot.Location = New-Object System.Drawing.Point(139, 606)
    $btnTot.Size = New-Object System.Drawing.Size(125, 34)
    $btnTot.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnTot
    $btnTot.add_Click({ & $fn.MarcaTot $true }.GetNewClosure())
    [void]$form.Controls.Add($btnTot)

    $btnCap = New-Object System.Windows.Forms.Button
    $btnCap.Text = 'Desmarcar-ho tot'
    $btnCap.Location = New-Object System.Drawing.Point(272, 606)
    $btnCap.Size = New-Object System.Drawing.Size(140, 34)
    $btnCap.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnCap
    $btnCap.add_Click({ & $fn.MarcaTot $false }.GetNewClosure())
    [void]$form.Controls.Add($btnCap)

    [void](_AddBrandHeader $form $titol $subtitol 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# Pas de la DOCUMENTACIO del tecnic redactor.
# Retorna @{ Nav; Text; Items }.
# ELS DOCUMENTS SIGNATS, en la forma que es DESA: nom -> @{ Marcat; Id }.
# Funcio PURA. Es la mateixa forma que torna la pantalla i que va a la base de
# dades, o sigui que no hi ha cap conversio pel mig.
function _LlicDocsBuits {
    $out = [ordered]@{}
    foreach ($d in @(_LlicDocsSignats)) { $out[[string]$d] = @{ Marcat = $false; Id = '' } }
    return $out
}

# Les linies que van a l'informe a partir d'aquell mapa. Funcio PURA.
function _LlicItemsDocsSignats($docs) {
    $out = New-Object System.Collections.ArrayList
    $m = @{}
    if ($docs -is [System.Collections.IDictionary]) {
        foreach ($k in @($docs.Keys)) { $m[[string]$k] = $docs[$k] }
    } elseif ($null -ne $docs) {
        foreach ($pr in @($docs.PSObject.Properties)) { $m[[string]$pr.Name] = $pr.Value }
    }
    # EN L'ORDRE DEL CATALEG, no el del mapa: un hashtable no en te, i a
    # l'informe els documents han de sortir sempre igual.
    foreach ($d in @(_LlicDocsSignats)) {
        if (-not $m.ContainsKey([string]$d)) { continue }
        $e = $m[[string]$d]
        $marcat = $false; $id = ''
        if ($e -is [System.Collections.IDictionary]) { $marcat = [bool]$e['Marcat']; $id = [string]$e['Id'] }
        elseif ($null -ne $e) { $marcat = [bool]$e.Marcat; $id = [string]$e.Id }
        if (-not $marcat) { continue }
        $id = $id.Trim()
        if ([string]::IsNullOrWhiteSpace($id)) { [void]$out.Add([string]$d) }
        else { [void]$out.Add([string]$d + ' (Id Firmadoc: ' + $id + ')') }
    }
    return $out.ToArray()
}

function Select-LlicTecnic($pre, $preDocs = $null) {
    $form = _NewForm
    $form.Text = 'Llic' + [char]0x00E8 + 'ncia - Documentaci' + [char]0x00F3
    $form.ClientSize = New-Object System.Drawing.Size(620, 430)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $camps = @(
        @{ K = 'Tecnic';  L = 'T' + [char]0x00E8 + 'cnic redactor:' }
        @{ K = 'NumCol';  L = 'N' + [char]0x00FA + 'm. col' + [char]0x00B7 + 'legiat:' }
        @{ K = 'Collegi'; L = 'Col' + [char]0x00B7 + 'legi:' }
        @{ K = 'Data';    L = 'Data de signatura:' }
    )
    $tb = @{}
    $y = 76
    foreach ($c in $camps) {
        $l = New-Object System.Windows.Forms.Label
        $l.Location = New-Object System.Drawing.Point(20, ($y + 3))
        $l.Size = New-Object System.Drawing.Size(150, 20)
        $l.Text = [string]$c.L
        [void]$form.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(175, $y)
        $t.Size = New-Object System.Drawing.Size(420, 22)
        if ($null -ne $pre -and $pre.Contains([string]$c.K)) { $t.Text = [string]$pre[[string]$c.K] }
        [void]$form.Controls.Add($t)
        $tb[[string]$c.K] = $t
        $y += 32
    }

    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Location = New-Object System.Drawing.Point(20, ($y + 8))
    $lblD.Size = New-Object System.Drawing.Size(560, 20)
    $lblD.Text = 'Quins documents s' + [char]0x2019 + 'han signat, i el seu Id Firmadoc:'
    [void]$form.Controls.Add($lblD)
    $y += 32

    $docs = @(_LlicDocsSignats)
    # EL QUE JA S'HAVIA TRIAT. Sense aixo, tornar Enrere o fer el segon informe
    # de la mateixa llicencia obligava a tornar a marcar-ho i a reescriure els
    # Id Firmadoc.
    $preD = @{}
    if ($preDocs -is [System.Collections.IDictionary]) {
        foreach ($k in @($preDocs.Keys)) { $preD[[string]$k] = $preDocs[$k] }
    } elseif ($null -ne $preDocs) {
        foreach ($pr in @($preDocs.PSObject.Properties)) { $preD[[string]$pr.Name] = $pr.Value }
    }
    $cbDoc = @{}
    $tbDoc = @{}
    foreach ($d in $docs) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Location = New-Object System.Drawing.Point(30, ($y + 2))
        $cb.Size = New-Object System.Drawing.Size(110, 22)
        $cb.Text = [string]$d
        [void]$form.Controls.Add($cb)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(150, $y)
        $t.Size = New-Object System.Drawing.Size(300, 22)
        [void]$form.Controls.Add($t)
        $lid = New-Object System.Windows.Forms.Label
        $lid.Location = New-Object System.Drawing.Point(458, ($y + 3))
        $lid.Size = New-Object System.Drawing.Size(140, 20)
        $lid.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
        $lid.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $lid.Text = 'Id Firmadoc'
        [void]$form.Controls.Add($lid)
        if ($preD.ContainsKey([string]$d)) {
            $e = $preD[[string]$d]
            if ($e -is [System.Collections.IDictionary]) { $cb.Checked = [bool]$e['Marcat']; $t.Text = [string]$e['Id'] }
            elseif ($null -ne $e) { $cb.Checked = [bool]$e.Marcat; $t.Text = [string]$e.Id }
        }
        $cbDoc[[string]$d] = $cb
        $tbDoc[[string]$d] = $t
        $y += 30
    }

    $res = @{ Nav = 'back'; Text = ''; Items = @(); Camps = @{}; Docs = (_LlicDocsBuits) }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(465, 380)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        $res.Text = _LlicTextDocumentacio $tb['Tecnic'].Text $tb['NumCol'].Text $tb['Collegi'].Text $tb['Data'].Text
        $tria = [ordered]@{}
        foreach ($d in $docs) {
            $tria[[string]$d] = @{
                Marcat = [bool]$cbDoc[[string]$d].Checked
                Id     = ([string]$tbDoc[[string]$d].Text).Trim()
            }
        }
        $res.Docs = $tria
        $res.Items = @(_LlicItemsDocsSignats $tria)
        foreach ($k in $tb.Keys) { $res.Camps[$k] = [string]$tb[$k].Text }
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, 380)
    $btnBack.Size = New-Object System.Drawing.Size(115, 32)
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form ('Documentaci' + [char]0x00F3) ('Qui ha signat el projecte i amb quin Id Firmadoc') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# ELS ADJUNTS: els informes dels organismes que posen les condicions, que van
# darrere del nostre quan es passa a PDF (PdfUnio.ps1).
#
# Quan l'usuari en tria un, se'n fa una COPIA LOCAL a la carpeta de la llicencia
# -local\base-dades-llicencies\GIA <id>\a.OGAU.pdf- (decisio de l'usuari): els
# originals solen ser a la unitat de xarxa o a Descarregues, i quan es passi
# l'informe a PDF, potser dies despres, han de ser-hi encara. A la carpeta hi
# queden, a mes, els originals SIGNATS i valids: al PDF ajuntat les seves
# signatures nomes hi son com a imatge.

# El nom de la copia: "a.OGAU.pdf". La lletra es la de l'informe (a., b., c...).
# Funcio PURA. Es treuen els caracters que el Windows no admet en un nom.
function _LlicNomAdjunt([int]$i, [string]$nom) {
    $n = ([string]$nom).Trim()
    foreach ($c in @('\', '/', ':', '*', '?', '"', '<', '>', '|')) { $n = $n.Replace($c, '-') }
    $n = [regex]::Replace($n, '[\x00-\x1F]', '')
    $n = $n.TrimEnd('.', ' ')
    return ((_LlicLletra $i).ToLower() + '.' + $n + '.pdf')
}

# La carpeta dels adjunts d'una llicencia: "GIA 924", al costat de la base de
# dades de llicencies. Funcio PURA.
function _LlicCarpetaAdjunts([string]$idGia) {
    $id = ([string]$idGia).Trim()
    foreach ($c in @('\', '/', ':', '*', '?', '"', '<', '>', '|')) { $id = $id.Replace($c, '-') }
    if ([string]::IsNullOrWhiteSpace($id)) { $id = 'sense GIA' }
    return [string](Join-Path $Script:LlicDbDir ('GIA ' + $id))
}

# Es un PDF? Nomes s'hi mira la capcalera ("%PDF-" al primer KB): triar un
# .docx amb extensio canviada no s'ha de descobrir el dia de passar-ho a PDF.
function _LlicEsPdf([string]$path) {
    try {
        $fs = [System.IO.File]::OpenRead($path)
        try {
            $buf = New-Object byte[] 1024
            $n = $fs.Read($buf, 0, $buf.Length)
            return ([System.Text.Encoding]::ASCII.GetString($buf, 0, $n).Contains('%PDF-'))
        } finally { $fs.Dispose() }
    } catch { return $false }
}

# Copia els PDF dels actors MARCATS a la carpeta de la llicencia, amb la lletra
# que tindran a l'informe. $actors: noms marcats EN ORDRE; $fonts: nom -> ruta
# triada. Retorna @{ Pdfs (nom -> copia local); Llista (les copies, en ordre);
# Errors }.
#
# Si la ruta triada JA ES la copia (ve de la memoria) i la lletra no ha canviat,
# no es torna a copiar. Si la lletra ha canviat, es copia amb el nom nou i la
# vella es deixa: un informe anterior hi pot apuntar.
function Copy-LlicAdjunts([string]$idGia, $actors, $fonts) {
    $pdfs = @{}
    if ($null -ne $fonts) { foreach ($k in @($fonts.Keys)) { $pdfs[[string]$k] = [string]$fonts[$k] } }
    $llista = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $carpeta = _LlicCarpetaAdjunts $idGia
    $i = 0
    foreach ($nom in @($actors)) {
        $i++
        $font = if ($pdfs.ContainsKey([string]$nom)) { [string]$pdfs[[string]$nom] } else { '' }
        if ([string]::IsNullOrWhiteSpace($font)) { continue }
        if (-not (Test-Path -LiteralPath $font -PathType Leaf)) { [void]$errors.Add(([string]$nom + ': no trobo el fitxer ' + $font)); continue }
        if (-not (_LlicEsPdf $font)) { [void]$errors.Add(([string]$nom + ': no es un PDF (' + (Split-Path -Leaf $font) + ')')); continue }
        $desti = [string](Join-Path $carpeta (_LlicNomAdjunt $i ([string]$nom)))
        $mateix = $false
        try { $mateix = ([System.IO.Path]::GetFullPath($font) -ieq [System.IO.Path]::GetFullPath($desti)) } catch { }
        if (-not $mateix) {
            try {
                if (-not (Test-Path -LiteralPath $carpeta)) { [void](New-Item -ItemType Directory -Path $carpeta -Force) }
                Copy-Item -LiteralPath $font -Destination $desti -Force -ErrorAction Stop
            } catch {
                [void]$errors.Add(([string]$nom + ': no s''ha pogut copiar -> ' + $_.Exception.Message)); continue
            }
        }
        $pdfs[[string]$nom] = $desti
        [void]$llista.Add($desti)
    }
    return @{ Pdfs = $pdfs; Llista = $llista.ToArray(); Errors = $errors.ToArray() }
}

# Pas de les CONDICIONS (nomes als favorables): QUINS ACTORS les posen i, de
# cada un, el PDF del seu informe.
#
# Abans era un quadre de text lliure (i despres, una casella al pas 1). L'usuari
# va explicar que les condicions les posen els organismes que informen els
# punts d'Autoritzacions / Informes preceptius, i que el que ha de dir l'informe
# es QUINS: els seus informes van adjunts darrere. Una llista per marcar, i amb
# un de marcat ja hi ha condicions.
#
# $pdfs: nom -> ruta ja triada (la copia local, si ve de la memoria).
# Retorna @{ Nav; Actors (els noms marcats, en l'ordre de la llista); Pdfs }.
function Select-LlicCondicions($actors, $marcats, $pdfs = $null) {
    $actors = @($actors)
    $form = _NewForm
    $form.Text = 'Condicions de la llic' + [char]0x00E8 + 'ncia'
    $form.ClientSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 70)
    $lbl.Size = New-Object System.Drawing.Size(660, 52)
    $lbl.Text = ('Marca qui posa condicions i tria el PDF del seu informe: en passar l' + [char]0x2019 + 'informe a PDF ' +
                 's' + [char]0x2019 + 'hi afegira darrere. Si no en marques cap, la conclusi' + [char]0x00F3 +
                 ' no parla de condicions. Surten marcats els que tenen l' + [char]0x2019 + 'informe preceptiu com a "Es disposa".')
    [void]$form.Controls.Add($lbl)

    $pan = New-Object System.Windows.Forms.Panel
    $pan.Location = New-Object System.Drawing.Point(20, 128)
    $pan.Size = New-Object System.Drawing.Size(660, 302)
    $pan.Anchor = 'Top,Bottom,Left,Right'
    $pan.AutoScroll = $true
    $pan.BorderStyle = 'FixedSingle'
    $pan.BackColor = [System.Drawing.Color]::White
    [void]$form.Controls.Add($pan)

    $marcatsSet = @{}
    foreach ($m in @($marcats)) { $marcatsSet[([string]$m).Trim().ToLowerInvariant()] = $true }
    $files = New-Object System.Collections.ArrayList
    $y = 8
    foreach ($a in $actors) {
        $nom = [string]$a.Nom
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Location = New-Object System.Drawing.Point(8, ($y + 2))
        $cb.Size = New-Object System.Drawing.Size(290, 22)
        $cb.AutoEllipsis = $true
        $cb.Text = $nom
        $cb.Checked = $marcatsSet.ContainsKey($nom.Trim().ToLowerInvariant())
        [void]$pan.Controls.Add($cb)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(304, $y)
        $tb.Size = New-Object System.Drawing.Size(250, 24)
        if ($null -ne $pdfs -and $pdfs.Contains($nom)) { $tb.Text = [string]$pdfs[$nom] }
        [void]$pan.Controls.Add($tb)
        $bt = New-Object System.Windows.Forms.Button
        $bt.Location = New-Object System.Drawing.Point(560, ($y - 1))
        $bt.Size = New-Object System.Drawing.Size(70, 26)
        $bt.Text = 'PDF' + [char]0x2026
        _StyleSecondaryButton $bt
        # Triar un PDF marca l'actor: si t'hi has molestat, es que hi va.
        $bt.add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = 'PDF (*.pdf)|*.pdf'
            $dlg.Title = 'Informe de: ' + $cb.Text
            try {
                $dir = Split-Path -Parent $tb.Text
                if ($dir -and (Test-Path -LiteralPath $dir)) { $dlg.InitialDirectory = $dir }
            } catch { }
            if ($dlg.ShowDialog() -eq 'OK') { $tb.Text = $dlg.FileName; $cb.Checked = $true }
            $dlg.Dispose()
        }.GetNewClosure())
        [void]$pan.Controls.Add($bt)
        [void]$files.Add(@{ Nom = $nom; Cb = $cb; Tb = $tb })
        $y += 32
    }

    $res = @{ Nav = 'back'; Actors = @(); Pdfs = @{} }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(550, 448)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    $btnOk.Anchor = 'Bottom,Right'
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        # En l'ordre de la LLISTA (el del cataleg), no en el que s'han clicat.
        $tri = New-Object System.Collections.ArrayList
        $rutes = @{}
        $sensePdf = New-Object System.Collections.ArrayList
        foreach ($f in $files) {
            $ruta = ([string]$f.Tb.Text).Trim().Trim('"')
            # Es recorden TOTES les rutes, tambe les dels no marcats: desmarcar
            # un actor un moment no ha de fer perdre el seu PDF.
            if ($ruta) { $rutes[[string]$f.Nom] = $ruta }
            if (-not $f.Cb.Checked) { continue }
            [void]$tri.Add([string]$f.Nom)
            if (-not $ruta) { [void]$sensePdf.Add([string]$f.Nom) }
        }
        if ($sensePdf.Count -gt 0) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                ("D'aquests no has triat el PDF:`n`n  " + ($sensePdf -join "`n  ") + "`n`n" +
                 "L'informe els anomenara, pero quan el passis a PDF no s'hi adjuntara el seu informe.`n`nContinuar igualment?"),
                'Condicions', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }
        $res.Actors = $tri.ToArray()
        $res.Pdfs = $rutes
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, 448)
    $btnBack.Size = New-Object System.Drawing.Size(115, 32)
    $btnBack.Anchor = 'Bottom,Left'
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form 'Condicions' ('Qui les posa i el seu informe (nom' + [char]0x00E9 + 's als favorables)') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}

# ----------------------------------------------------------------------------
# PUNT D'ENTRADA (des del menu)
# ----------------------------------------------------------------------------
# $fases: quines fases ofereix aquest assistent. El menu principal en te DUES
# entrades -Llicencia (_LlicFases) i Modificacio NO Substancial / Traspas
# (_MnsFases)- i totes dues passen per aqui: comparteixen capcalera, tramit i
# base de dades, i el que canvia es nomes el document que en surt.
function Invoke-LlicenciaWizard($fases = $null, [string]$titol = '') {
    $llic = Read-LlicCataleg
    if ($null -eq $llic) {
        [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha trobat ESTRUCTURALS\LLIC.json.`n`nAquest fitxer es la base de dades de Llicencia " +
             "(que aporta cada requeriment i que no). Fes 'Actualitzar.bat' per baixar-lo."),
            'Llicencia', 'OK', 'Error') | Out-Null
        return
    }

    $word = $null
    # $st.Fields es el diccionari de camps COMPARTIT de tot l'assistent (el
    # mateix paper que a Invoke-NouWizard): els [CAMP:]/[OPCIO:] s'hi omplen
    # alla on surten i despres la composicio els hi busca.
    # La fase inicial surt de la llista d'AQUEST assistent, no d'un literal: la
    # de MNS/Traspas no te cap 'requeriment'.
    $st = @{ Fase = (_LlicFasePerDefecte $(if ($null -ne $fases) { $fases } else { _LlicTotesLesFases }) 'requeriment')
             Prov = $false; Tecnic = @{}
             # Els actors marcats al pas de les condicions. $null = encara no
             # s'hi ha passat (i llavors se'n proposen segons el bloc ABANS).
             CondActors = $null
             Fields = [ordered]@{}
             # El que s'havia triat a cada pantalla de documentacio, per no
             # perdre-ho quan l'usuari torna ENRERE (era exactament el que
             # passava: tornaves i havies de tornar a marcar-ho tot).
             MemAbans = $null; MemDespres = $null
             # La base de dades nomes es llegeix un cop per sessio (vegeu pas 2).
             DbCarregat = $false }
    $step = 1
    try {
        while ($true) {
            switch ($step) {
                1 {
                    $r = Select-LlicFase $st.Fase $st.Prov $fases $titol
                    if ($r.Nav -ne 'fwd') { return }
                    $st.Fase = [string]$r.Fase
                    $st.Prov = [bool]$r.Prov
                    $step = 2
                }
                2 {
                    $r = Get-HeaderData -preload $st.HeaderPre
                    if ($r.Nav -eq 'back') { $step = 1; break }
                    $st.Header = $r.Data
                    $st.HeaderPre = $r.Data
                    # LA CLASSIFICACIO. La capcalera generica no en te camp (es
                    # NOMES de Llicencia), o sigui que s'omple aqui des de
                    # l'Excel, per ID GIA. Si l'Excel no en te, es demana: sortia
                    # una linia "Classificacio:" BUIDA a l'informe.
                    $st.Header['CLASSIFICACIO'] = _LlicClassificacio $st.Header
                    # LA MEMORIA D'AQUESTA LLICENCIA. Un informe de llicencia
                    # gairebe mai va sol (requeriment -> favorable pre -> post) i
                    # fins ara el segon tornava a demanar-ho TOT, Id Firmadoc i
                    # expedients inclosos. Es carrega UNA sola vegada per sessio:
                    # si l'usuari torna Enrere, el que acaba d'editar mana.
                    if (-not $st.DbCarregat) {
                        $st.DbCarregat = $true
                        $rec = Get-LlicenciaRecord (Load-LlicenciaDb) ([string]$st.Header['ID_GIA'])
                        if ($null -ne $rec) {
                            [void](Restore-LlicenciaState $rec $st)
                            $quan = Get-LlicenciaDataText $rec
                            [System.Windows.Forms.MessageBox]::Show(
                                ("S'han recuperat les dades de l'informe de llic" + [char]0x00E8 + 'ncia del ' + $quan + ".`n`n" +
                                 "Ho trobaras ja marcat i omplert als passos seguents; canvia el que calgui."),
                                'Llicencia', 'OK', 'Information') | Out-Null
                        }
                    }
                    # ELS DOS INFORMES CURTS (Modificacio NO Substancial i
                    # Traspas) no tenen ni blocs de documentacio ni deficiencies
                    # de projecte: nomes cal saber si hi ha observacions.
                    $step = if (_MnsEsFase ([string]$st.Fase)) { 20 } else { 3 }
                }
                20 {
                    if ($null -eq $st.MnsCataleg) { $st.MnsCataleg = Read-MnsCataleg }
                    if ($null -eq $st.MnsCataleg) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("No trobo el cataleg MNSTRAS.json a ESTRUCTURALS.`n`n" +
                             "Sense el text no es pot fer aquest informe."),
                            'Llicencia', 'OK', 'Error') | Out-Null
                        return
                    }
                    # QUINS PUNTS DE REQ1 S'HI ADJUNTEN. Amb la pantalla de
                    # sempre (Select-Items) i el cataleg SENCER de REQ1: aqui no
                    # hi ha cap seccio que sobri, perque no s'ha demanat res
                    # abans. -permetreBuit: no marcar-ne cap vol dir "sense mes
                    # observacions", que es un cas ben normal.
                    if ($null -eq $st.Req1) {
                        $st.Req1 = Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')
                        $st.IdxReq1 = _LlicIndexReq1 $st.Req1
                    }
                    $r = Select-Items -sections @($st.Req1.Sections) -preloadSelectedKeys $st.ProjKeys `
                            -fields $st.Fields -preloadValues $st.ProjVals -permetreBuit $true
                    if ($r.Nav -eq 'back') { $step = 2; break }
                    if ($r.Nav -eq 'stay') { break }
                    $st.ProjSel = $r.Data
                    $st.ProjKeys = Get-SelectedKeysFromResult $st.ProjSel
                    $st.ProjVals = Get-FieldValuesForSession $st.Fields
                    if ($null -eq $word) { $word = New-WordApp }
                    $out = Build-MnsDocument $word @{
                        Fase = [string]$st.Fase
                        Header = $st.Header
                        Fields = $st.Fields
                        Punts = @($st.ProjSel)
                        Cataleg = $st.MnsCataleg
                    }
                    _LlicObreIAvisa $word $out
                    return
                }
                3 {
                    # Aqui nomes cal REQ1 (el JSON d'on surt el text). El Word
                    # s'arrenca DIFERIT al pas 9, quan es genera de debo (mateix
                    # motiu que a Invoke-NouWizard: arrencar-lo en fred es lent
                    # i si l'usuari tira enrere no ha de quedar obert per res).
                    if ($null -eq $st.Req1) {
                        $req1Path = Join-Path $EstructuralsDir 'REQ1.json'
                        $st.Req1 = Get-ParsedCataleg -path $req1Path
                        $st.IdxReq1 = _LlicIndexReq1 $st.Req1
                    }
                    # Els punts, resolts amb el text de REQ1. Les claus ORFES
                    # s'avisen: si algu ha reanomenat un requeriment a REQ1, el
                    # punt desapareixeria de l'informe sense dir res.
                    $bAbans  = _LlicPuntsPerBloc $llic $st.IdxReq1 'ABANS' $st.Req1
                    $bDesp   = _LlicPuntsPerBloc $llic $st.IdxReq1 'DESPRES' $st.Req1
                    $bPropis = _LlicPuntsPerBloc $llic $st.IdxReq1 'PROPIS'
                    $orfes = @($bAbans.Orfes) + @($bDesp.Orfes) + @($bPropis.Orfes)
                    if ($orfes.Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("Hi ha " + $orfes.Count + " punt(s) de Llicencia que apunten a un requeriment que JA NO " +
                             "existeix a REQ1:`n`n  " + ($orfes -join "`n  ") +
                             "`n`nNo sortiran a l'informe. Arregla-ho des de l'editor de catalegs."),
                            'Llicencia', 'OK', 'Warning') | Out-Null
                    }
                    # Els condicionals entren segons el tipus de llicencia.
                    $cond = @(@($bPropis.Punts) | Where-Object { _LlicCondicioEntra ([string]$_.Condicio) ([bool]$st.Prov) })
                    $st.AbansTots = @($cond) + @($bAbans.Punts)
                    # ELS TEXTOS DE LA FASE al bloc DESPRES (nomes al post: "Es
                    # disposa..." o "No es disposa de la documentacio.").
                    # S'apliquen AQUI perque la pantalla del pas 7 i el document
                    # facin servir EXACTAMENT els mateixos punts.
                    $st.DespresTots = @(_LlicPuntsAmbEstatFase $bDesp.Punts ([string]$st.Fase))
                    $step = 4
                }
                4 {
                    $r = Select-LlicDocumentacio $st.AbansTots ('Documentaci' + [char]0x00F3 + ' ABANS de la resoluci' + [char]0x00F3) `
                            ('Marca la que aplica, si ja es t' + [char]0x00E9 + ' i les seves dades') $true $false $true $st.MemAbans
                    if ($r.Nav -ne 'fwd') { $step = 2; break }
                    $st.Abans = $r.Punts
                    $st.MemAbans = $r.Memoria
                    $step = 5
                }
                5 {
                    # Projecte: els requeriments NORMALS de REQ1, amb la mateixa
                    # pantalla de sempre (no s'hi inventa res).
                    # El PROJECTE es la resta de REQ1: les seccions de
                    # documentacio ja s'han demanat al pas d'ABANS i no s'han de
                    # poder demanar dues vegades.
                    if ($null -eq $st.SeccionsProjecte) {
                        $senseAbans = @(@($st.Req1.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
                        # ...i fora tambe tot el que un altre bloc ja expandeix
                        # sencer (no es pot demanar dues vegades). La llista surt
                        # del PROPI cataleg, no d'aqui.
                        $st.SeccionsProjecte = @(_LlicSeccionsSenseSubseccions $senseAbans (_LlicSeccionsExpandides $llic $st.IdxReq1))
                    }
                    # -permetreBuit: pot ser que l'activitat no tingui cap
                    # deficiencia de projecte, i llavors no s'ha d'aturar res.
                    $r = Select-Items -sections $st.SeccionsProjecte -preloadSelectedKeys $st.ProjKeys -fields $st.Fields -preloadValues $st.ProjVals -permetreBuit $true
                    if ($r.Nav -eq 'back') { $step = 4; break }
                    if ($r.Nav -eq 'stay') { break }
                    $st.ProjSel = $r.Data
                    $st.ProjKeys = Get-SelectedKeysFromResult $st.ProjSel
                    $st.ProjVals = Get-FieldValuesForSession $st.Fields
                    $step = 6
                }
                6 {
                    $r = Select-LlicTecnic $st.Tecnic $st.TecnicDocs
                    if ($r.Nav -ne 'fwd') { $step = 5; break }
                    $st.Doc = @{ Text = [string]$r.Text; Items = @($r.Items) }
                    $st.Tecnic = $r.Camps
                    $st.TecnicDocs = $r.Docs
                    $step = 7
                }
                7 {
                    # LA MATEIXA PANTALLA PER A LES TRES FASES. Nomes canvia si
                    # es demana l'estat de cada punt (i les seves dades), que ho
                    # decideix _LlicEstatDespres: al requeriment i al favorable
                    # pre encara no toca dir si es te o no; al post, si.
                    $ef = _LlicEstatDespres ([string]$st.Fase)
                    $titol = 'Documentaci' + [char]0x00F3 + ' DESPR' + [char]0x00C9 + 'S de la resoluci' + [char]0x00F3
                    $sub = "Marca la que entra a l'informe"
                    # Tot marcat de sortida: el Word de l'usuari els portava tots
                    # i ell hi anava esborrant el que no tocava.
                    $r = Select-LlicDocumentacio $st.DespresTots $titol $sub ([bool]$ef.AmbEstat) $true ([bool]$ef.AmbDades) `
                            $st.MemDespres $true ([string]$ef.Estat)
                    if ($r.Nav -ne 'fwd') { $step = 6; break }
                    $st.Despres = $r.Punts
                    $st.MemDespres = $r.Memoria
                    $step = 8
                }
                8 {
                    # LES CONDICIONS, despres de tota la documentacio: la tria
                    # per defecte surt del bloc ABANS (els informes preceptius
                    # que ja es tenen), o sigui que ha d'anar darrere.
                    if (-not (_LlicAdmetCondicions ([string]$st.Fase))) { $step = 9; break }
                    $actors = @(_LlicActorsCondicions $llic)
                    $pre = if ($null -ne $st.CondActors) { @($st.CondActors) } else { @(_LlicActorsPerDefecte $actors $st.Abans) }
                    $r = Select-LlicCondicions $actors $pre $st.CondPdfs
                    if ($r.Nav -ne 'fwd') { $step = 7; break }
                    # LA COPIA LOCAL dels PDF triats (local\base-dades-llicencies\
                    # GIA <id>\a.OGAU.pdf). Si en falla alguna, es diu i es torna
                    # a la pantalla: val mes saber-ho ara que el dia de passar-ho
                    # a PDF.
                    $cp = Copy-LlicAdjunts ([string]$st.Header['ID_GIA']) @($r.Actors) $r.Pdfs
                    if (@($cp.Errors).Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("No s'han pogut guardar aquests PDF:`n`n  " + (@($cp.Errors) -join "`n  ")),
                            'Condicions', 'OK', 'Warning') | Out-Null
                        $st.CondActors = @($r.Actors)
                        $st.CondPdfs = $r.Pdfs
                        break
                    }
                    $st.CondActors = @($r.Actors)
                    $st.CondPdfs = $cp.Pdfs
                    $st.CondAdjunts = @($cp.Llista)
                    $step = 9
                }
                9 {
                    if ($null -eq $word) { $word = New-WordApp }
                    # Fora dels favorables, cap actor ni cap adjunt (encara que la
                    # memoria en porti).
                    $actorsModel = if (_LlicAdmetCondicions ([string]$st.Fase)) { @($st.CondActors) } else { @() }
                    $adjuntsModel = if (_LlicAdmetCondicions ([string]$st.Fase)) { @($st.CondAdjunts) } else { @() }
                    # Els camps [CAMP: ...] dels textos triats.
                    $model = @{
                        Fase = [string]$st.Fase
                        EsProvisional = [bool]$st.Prov
                        Header = $st.Header
                        Fields = $st.Fields
                        Abans = @($st.Abans)
                        Projecte = @(_LlicPuntsDeSeleccio $st.ProjSel)
                        Despres = @($st.Despres)
                        Doc = $(if ($null -ne $st.Doc) { $st.Doc } else { @{ Text = ''; Items = @() } })
                        CondicionsActors = $actorsModel
                        Cataleg = $llic
                    }
                    $out = Build-LlicenciaDocument $word $model
                    # Es desa igual que a "Requeriment - Nou" perque el boto
                    # "Recuperar dades ultim informe" del Pas 2 hi arribi. Les
                    # claus desades son les de REQ1 (el bloc Projecte es el
                    # mateix cataleg), o sigui que serveixen als dos fluxos.
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = 'LLIC'
                        Header          = $st.Header
                        SelectedKeys    = @($st.ProjKeys)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @()
                    })
                    # ...i a la BASE DE DADES DE LLICENCIES, que es el que fa que
                    # el proper informe d'aquesta activitat surti ja omplert.
                    # Un error aqui no pot fer perdre l'informe, que ja esta fet.
                    try {
                        $db = Load-LlicenciaDb
                        $vell = Get-LlicenciaRecord $db ([string]$st.Header['ID_GIA'])
                        $hist = New-Object System.Collections.ArrayList
                        if ($null -ne $vell) { foreach ($x in @($vell.Historial)) { [void]$hist.Add($x) } }
                        # Amb els ADJUNTS d'aquest informe: "Word a PDF" els hi
                        # buscara per saber que ha d'ajuntar-hi (PdfSignar.ps1).
                        [void]$hist.Add((New-LlicenciaHistorial ([string]$st.Fase) ([string]$out) $adjuntsModel))
                        [void](Set-LlicenciaRecord $db (ConvertTo-LlicenciaRecord $st $hist.ToArray()))
                        Save-LlicenciaDb $db
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("L'informe s'ha generat be, pero no s'han pogut desar les dades a la base de " +
                             "llicencies:`n`n" + $_.Exception.Message),
                            'Llicencia', 'OK', 'Warning') | Out-Null
                    }
                    _LlicObreIAvisa $word $out
                    return
                }
                default { return }
            }
        }
    } catch {
        # SENSE AIXO, qualsevol error aqui dins matava el programa EN SILENCI:
        # "es tanca i no passa res, tampoc es genera cap informe". Ara es diu
        # que ha passat i ON, i es torna al menu en lloc de tancar-ho tot.
        $on = ''
        try { $on = "`n`n(" + [System.IO.Path]::GetFileName([string]$_.InvocationInfo.ScriptName) + ', linia ' + [string]$_.InvocationInfo.ScriptLineNumber + ')' } catch { }
        [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha pogut acabar l'informe de Llicencia:`n`n" + $_.Exception.Message + $on),
            'Llicencia', 'OK', 'Error') | Out-Null
    } finally {
        # Si l'informe s'ha generat, el Word s'ha fet visible i es deixa obert
        # per a l'usuari; si es va cancel·lar pel cami, es tanca.
        if ($null -ne $word -and -not $word.Visible) { try { Close-WordApp $word } catch { } }
    }
}

# Converteix el que retorna Select-Items (les deficiencies normals de REQ1) als
# mateixos punts que fa servir la composicio, per no tenir dos camins.
function _LlicPuntsDeSeleccio($seleccio) {
    $out = New-Object System.Collections.ArrayList
    # Amb la SECCIO, la SUBSECCIO i l'INTRO de cada punt: al bloc PROJECTE hi han
    # de sortir igual que a REQ1, no una llista plana de punts.
    foreach ($u in @(_LlicItemsAmbUbicacio $seleccio)) {
        $it = $u.El
        $subs = New-Object System.Collections.ArrayList
        foreach ($ch in @($it.Children)) { [void]$subs.Add(@($ch.BodyLines)) }
        [void]$out.Add([pscustomobject]@{
            Clau = ''; Titol = [string]$it.Short; Condicio = ''
            Seccio = [string]$u.Seccio; Subseccio = [string]$u.Subseccio; Intro = @($u.Intro)
            IntroDeSeccio = [bool]$u.IntroDeSeccio
            Cos = @($it.BodyLines); NoDisposa = @(); SiDisposa = @()
            Quan = @(); Subs = $subs.ToArray(); Estat = ''
        })
    }
    return $out.ToArray()
}
