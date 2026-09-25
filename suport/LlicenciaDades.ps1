#requires -Version 5.1
<#
.SYNOPSIS
  Llicencia: les DADES. Funcions PURES (es proven a Linux, sense Word ni
  WinForms): el cataleg LLIC.json resolt contra REQ1, les fases, els textos de
  cada fase, la conclusio, els actors de les condicions, els noms de fitxer i
  la copia local dels PDF adjunts.

  El modul de Llicencia son QUATRE fitxers, un per cosa (revisio d'arquitectura,
  setembre 2026; abans era un de sol de 2.550 linies):
    LlicenciaDades.ps1      les dades (aquest)
    LlicenciaBlocs.ps1      que s'escriu a l'informe (blocs purs + Write-Informe)
    LlicenciaPantalles.ps1  les pantalles de l'assistent (WinForms)
    Llicencia.ps1           l'assistent: el recorregut dels passos
  Tots van amb dot-source al mateix ambit (Motor.ps1): partir-lo no en canvia
  el comportament. Llegeix suport/documentacio/llicencia.md abans de tocar-lo.
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
# _LlicBlocsPunts). Amb aixo:
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
    $m = ConvertTo-Mapa $docs
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

# LA LLETRA D'UN PUNT ("A", "B"... "AA"). PURA. La fan servir la marca dels
# punts (_LlicMarca, LlicenciaBlocs.ps1) i el nom dels adjunts (a.OGAU.pdf).
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
