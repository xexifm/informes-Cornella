#requires -Version 5.1
<#
.SYNOPSIS
  Mode ACT_EXTR: la logica del Decret 112/2010 (aplicabilitat i valors
  calculats), els punts, el parseig de plantilles i el registre local.
  Funcions PURES, provades a Linux. Calcula les rutes del registre i de les
  plantilles EN CARREGAR-SE: ha d'anar despres del bloc de rutes de Motor.ps1.

  Part del modul ACT_EXTR, partit per responsabilitats (com Llicencia):
    ActExtrDades.ps1     logica del Decret, punts, plantilla i registre (PURES)
    ActExtrBlocs.ps1     composicio del document en blocs (PURA) + el .docx
    ActExtrPantalles.ps1 les finestres (WinForms)
    ActExtr.ps1          l'orquestrador (Invoke-ActExtrFlow) i la descripcio
  Tot va amb dot-source al mateix ambit (Motor.ps1 els carrega en aquest ordre).
#>


# ----------------------------------------------------------------------------
# Localitzacio del registre local (JSON, ignorat per git)
# ----------------------------------------------------------------------------
# Per defecte, local\base-dades-actextr\ (dins del clone pero fora del
# repositori: 'local' s'ignora sencera).
# Es pot sobreescriure $script:ActExtrRegistryDir des de config.ps1 o des dels
# tests. $RepoRoot el defineix GenerarInforme.ps1 abans de carregar aquest fitxer.
if (-not $script:ActExtrRegistryDir) {
    $script:ActExtrRegistryDir = if ($RepoRoot) { Get-LocalSubdir $RepoRoot 'ActExtr' } else { 'base-dades-actextr' }
}
$script:ActExtrRegistryFile = 'activitats-extraordinaries.json'

# Rutes de les plantilles ACT_EXTR (a ESTRUCTURALS). $EstructuralsDir el
# defineix GenerarInforme.ps1.
if ($EstructuralsDir) {
    $script:ActExtrReqTemplate = Join-Path $EstructuralsDir 'ACT_EXTR_REQ.json'
    $script:ActExtrFavTemplate = Join-Path $EstructuralsDir 'ACT_EXTR_FAV.json'
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES - logica del Decret 112/2010 (testejables en headless)
# ----------------------------------------------------------------------------

# VLOOKUP aproximat (coincidencia per defecte, com VLOOKUP(...,TRUE)): donat un
# valor numeric i una llista de parells @(@(llindar, valor), ...) ORDENADA de
# menys a mes per llindar, retorna el 'valor' del llindar mes gran que sigui
# <= $n. Si $n < primer llindar, retorna el valor del primer parell.
function _ActExtrVlookup([double]$n, $pairs) {
    $res = $pairs[0][1]
    foreach ($p in $pairs) {
        if ($n -ge $p[0]) { $res = $p[1] } else { break }
    }
    return $res
}

# Formata un enter amb punt com a separador de milers (2000000 -> "2.000.000").
function _ActExtrThousands([int64]$n) {
    $neg = $n -lt 0
    $s = [string][math]::Abs($n)
    $out = ''
    $c = 0
    for ($i = $s.Length - 1; $i -ge 0; $i--) {
        $out = $s[$i] + $out
        $c++
        if (($c % 3) -eq 0 -and $i -gt 0) { $out = '.' + $out }
    }
    if ($neg) { $out = '-' + $out }
    return $out
}

# Nombre de vigilants de seguretat privada (Decret 112/2010, Article 43) segons
# l'aforament. Taula: 0->0, 501->1, 1001->2, 2001->3, +1 cada +1000.
function Get-ActExtrVigilants([double]$aforament) {
    if ($aforament -lt 501) { return 0 }
    return ([int][math]::Floor(($aforament - 1) / 1000) + 1)
}

# Nombre de personal de control d'acces (Article 58) segons l'aforament. Taula:
# 0->0, 150->2, 501->3, 1001->4, 2001->5, +1 cada +1000.
function Get-ActExtrControladors([double]$aforament) {
    if ($aforament -lt 150) { return 0 }
    if ($aforament -lt 501) { return 2 }
    if ($aforament -lt 1001) { return 3 }
    return ([int][math]::Floor(($aforament - 1) / 1000) + 3)
}

# Lavabos i cabines de vater (Article 47) segons l'aforament. Retorna
# @{ Lavabos; Cabines }.
function Get-ActExtrHigiene([double]$aforament) {
    $pairsL = @(@(0,1),@(51,2),@(151,2),@(301,4),@(501,4),@(1001,8))
    $pairsC = @(@(0,2),@(51,4),@(151,6),@(301,8),@(501,12),@(1001,24))
    if ($aforament -ge 1001) {
        # A partir de 1001, +500 d'aforament = +4 lavabos i +12 cabines.
        $steps  = [int][math]::Floor(($aforament - 1001) / 500)
        return @{ Lavabos = (8 + 4 * $steps); Cabines = (24 + 12 * $steps) }
    }
    return @{
        Lavabos = [int](_ActExtrVlookup $aforament $pairsL)
        Cabines = [int](_ActExtrVlookup $aforament $pairsC)
    }
}

# Quantia minima de la polissa de responsabilitat civil (Articles 80-81) segons
# l'aforament i si l'activitat es du a terme sota rasant (parcialment *1,25 /
# totalment *1,30; s'aplica el factor mes alt). Retorna l'import en euros (int).
function Get-ActExtrPolissaRC([double]$aforament, [bool]$parcialSotaRasant, [bool]$totalSotaRasant) {
    $base = if     ($aforament -lt 101)  { 300000 }
            elseif ($aforament -lt 151)  { 400000 }
            elseif ($aforament -lt 301)  { 600000 }
            elseif ($aforament -lt 501)  { 750000 }
            elseif ($aforament -lt 1001) { 900000 }
            elseif ($aforament -lt 1501) { 1200000 }
            elseif ($aforament -lt 2501) { 1600000 }
            elseif ($aforament -lt 5001) { 2000000 }
            else {
                $v = 2000000 + ([int][math]::Floor(($aforament - 5001) / 1000) + 1) * 60000
                [math]::Min($v, 6000000)
            }
    $factor = 1.0
    if ($parcialSotaRasant -and 1.25 -gt $factor) { $factor = 1.25 }
    if ($totalSotaRasant   -and 1.30 -gt $factor) { $factor = 1.30 }
    return [int64][math]::Round($base * $factor)
}

# Normalitza una resposta Si/No a 'Si' o 'No' (accepta variants: si, s, yes,
# true, compleix...). Per defecte 'No'.
function _ActExtrYesNo($v) {
    $n = _NormalitzaText ([string]$v)
    if ($n -in @('si','s','yes','y','true','1','compleix')) { return 'Si' }
    return 'No'
}

# Construeix l'objecte "decret" normalitzat a partir d'un hashtable/PSObject de
# respostes (les que recull el Pas 3 i es desen al registre). Claus esperades:
#   Aforament, Incendis, Mobilitat, ControlAccessos, PauCatalunya, PauLocal,
#   EstablimentDotat, ParcialSotaRasant, TotalSotaRasant.
function Build-ActExtrDecret($answers) {
    $get = {
        param($k)
        if ($null -eq $answers) { return $null }
        if ($answers -is [System.Collections.IDictionary]) {
            if ($answers.Contains($k)) { return $answers[$k] }
            return $null
        }
        if ($answers.PSObject.Properties.Name -contains $k) { return $answers.$k }
        return $null
    }
    $af = 0.0
    $rawAf = & $get 'Aforament'
    [double]::TryParse(([string]$rawAf -replace '[^\d]', ''), [ref]$af) | Out-Null
    return [pscustomobject]@{
        Aforament         = [int]$af
        Incendis          = _ActExtrYesNo (& $get 'Incendis')
        Mobilitat         = _ActExtrYesNo (& $get 'Mobilitat')
        ControlAccessos   = _ActExtrYesNo (& $get 'ControlAccessos')
        PauCatalunya      = _ActExtrYesNo (& $get 'PauCatalunya')
        PauLocal          = _ActExtrYesNo (& $get 'PauLocal')
        EstablimentDotat  = _ActExtrYesNo (& $get 'EstablimentDotat')
        ParcialSotaRasant = _ActExtrYesNo (& $get 'ParcialSotaRasant')
        TotalSotaRasant   = _ActExtrYesNo (& $get 'TotalSotaRasant')
        HiHaLasers        = _ActExtrYesNo (& $get 'HiHaLasers')
    }
}

# Calcula tots els valors derivats de l'aforament/respostes. Retorna un
# hashtable amb els tokens numerics (per substituir {{...}} a les plantilles) i
# metadades (organ del Pla d'Autoproteccio, etc.).
function Get-ActExtrComputed($decret) {
    $af = [double]$decret.Aforament
    $hig = Get-ActExtrHigiene $af
    $rc = Get-ActExtrPolissaRC $af ($decret.ParcialSotaRasant -eq 'Si') ($decret.TotalSotaRasant -eq 'Si')
    # Organ del Pla d'Autoproteccio: Catalunya te prioritat si esta a tots dos.
    $organKey = if ($decret.PauCatalunya -eq 'Si') { 'CAT' }
                elseif ($decret.PauLocal -eq 'Si')  { 'LOCAL' }
                else { '' }
    $pauObligat = ($organKey -ne '')

    # Assistencia sanitaria ({{ASSISTENCIA}}): text que completa la frase
    # "...En aquest cas {{ASSISTENCIA}}." de la plantilla.
    #  - Si cal Pla d'Autoproteccio  -> els dispositius els determina el PAU
    #    homologat, d'acord amb l'Annex III del Decret 30/2015.
    #  - Si NO cal PAU               -> Article 48 del Decret 112/2010:
    #      aforament < 1000 -> farmaciola; aforament >= 1000 -> infermeria.
    $assistencia =
        if ($pauObligat) {
            "els dispositius d'assist" + [char]0x00E8 + "ncia sanit" + [char]0x00E0 + "ria seran els que estableixi el Pla d'Autoprotecci" + [char]0x00F3 + " homologat, d'acord amb l'Annex III del Decret 30/2015, de 3 de mar" + [char]0x00E7 + ", pel qual s'aprova el cat" + [char]0x00E0 + "leg d'activitats i centres obligats a adoptar mesures d'autoprotecci" + [char]0x00F3
        } elseif ($af -lt 1000) {
            "s'ha de disposar d'una farmaciola amb els materials i els equips adequats per facilitar primeres cures en cas d'accident, malaltia o crisi sobtada"
        } else {
            "s'ha de disposar d'una infermeria amb instal" + [char]0x00B7 + "lacions, materials i equips adequats per prestar els primers auxilis en cas d'accident, malaltia o crisi sobtada. La infermeria pot ser substitu" + [char]0x00EF + "da per una farmaciola i la pres" + [char]0x00E8 + "ncia de vehicles medicalitzats mentre l'establiment estigui obert al p" + [char]0x00FA + "blic o l'activitat recreativa s'estigui duent a terme"
        }

    return @{
        VIGILANTS    = (Get-ActExtrVigilants $af)
        CONTROLADORS = (Get-ActExtrControladors $af)
        LAVABOS      = $hig.Lavabos
        CABINES      = $hig.Cabines
        RC_IMPORT    = (_ActExtrThousands $rc)
        RC_RAW       = $rc
        PAU_ORGAN_KEY= $organKey
        PAU_OBLIGAT  = $pauObligat
        HAS_LASERS   = ($decret.HiHaLasers -eq 'Si')
        ASSISTENCIA  = $assistencia
    }
}

# Substitueix els tokens numerics {{VIGILANTS}}, {{CONTROLADORS}}, {{LAVABOS}},
# {{CABINES}}, {{RC_IMPORT}} d'un text pels valors calculats.
function Resolve-ActExtrTokens([string]$text, $computed) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    $out = $text
    foreach ($k in @('VIGILANTS','CONTROLADORS','LAVABOS','CABINES','RC_IMPORT','ASSISTENCIA')) {
        $out = $out.Replace('{{' + $k + '}}', [string]$computed[$k])
    }
    return $out
}

# ----------------------------------------------------------------------------
# Punts del Decret: aplicabilitat + motiu ("aplica i per que")
# ----------------------------------------------------------------------------
# Ordre dels punts tal com es mostren al Pas 3. Cada punt te:
#   Key          : clau interna (tambe clau de l'estat "lliurat" al registre)
#   Title        : titol curt llegible
#   NeedsDelivery: si l'usuari ha de marcar "lliurat/pendent" per aquest punt
$script:ActExtrPoints = @(
    @{ Key='INCENDIS';             Title='Incendis (informe de prevencio)';        NeedsDelivery=$true }
    @{ Key='MOBILITAT';            Title='Mobilitat (Guardia Urbana)';             NeedsDelivery=$true }
    @{ Key='PAU';                  Title="Pla d'Autoproteccio";                    NeedsDelivery=$true }
    @{ Key='RC';                   Title='Responsabilitat civil (polissa)';        NeedsDelivery=$true }
    @{ Key='ASSIST_SANITARIA';     Title='Assistencia sanitaria';                  NeedsDelivery=$true }
    @{ Key='CONTROLADORS';         Title="Controladors d'acces";                   NeedsDelivery=$true }
    @{ Key='VIGILANTS';            Title='Seguretat privada (vigilants)';          NeedsDelivery=$true }
    @{ Key='SERVEIS_HIGIENE';      Title="Serveis d'higiene";                      NeedsDelivery=$true }
    @{ Key='IMPACTE_ACUSTIC';      Title='Impacte acustic';                        NeedsDelivery=$true }
    @{ Key='DISPONIBILITAT_ESPAI'; Title="Disponibilitat de l'espai";              NeedsDelivery=$true }
    @{ Key='LASERS';               Title='Lasers (autoritzacio)';                  NeedsDelivery=$true }
    @{ Key='MEMORIA_A';            Title='Memoria a) Identificacio';               NeedsDelivery=$true }
    @{ Key='MEMORIA_B';            Title='Memoria b) Data i horari';               NeedsDelivery=$true }
    @{ Key='MEMORIA_C';            Title='Memoria c) Responsables';                NeedsDelivery=$true }
    @{ Key='MEMORIA_D';            Title='Memoria d) Descripcio i aforament';      NeedsDelivery=$true }
    @{ Key='MEMORIA_E';            Title='Memoria e) Mesures adoptades';           NeedsDelivery=$true }
    @{ Key='MEMORIA_F';            Title='Memoria f) Declaracio polissa';          NeedsDelivery=$true }
    @{ Key='MEMORIA_G';            Title='Memoria g) Titulars disponibilitat';     NeedsDelivery=$true }
)

# Decideix, per a una clau de punt, si APLICA a aquesta activitat i PER QUE.
# Retorna @{ Applies=bool; Reason=string }. Funcio PURA.
function Get-ActExtrPointApplicability([string]$key, $decret, $computed) {
    $af = [int]$decret.Aforament
    switch ($key) {
        'INCENDIS' {
            if ($decret.Incendis -eq 'Si') {
                return @{ Applies=$true; Reason="Art. 23 Llei 3/2010: l'activitat esta inclosa en algun suposit (acte esporadic >500 persones en establiment tancat o >1.000 en espai obert; estructures desmuntables/itinerants >1.000). Cal l'informe de prevencio i seguretat en materia d'incendis." }
            }
            return @{ Applies=$false; Reason="Art. 23 Llei 3/2010: no inclosa en cap suposit. Nomes cal complir la SP-144." }
        }
        'MOBILITAT' {
            if ($decret.Mobilitat -eq 'Si') {
                return @{ Applies=$true; Reason="Art. 111.b Decret 112/2010 (Decret 344/2006, art. 3.4): cal estudi d'avaluacio de la mobilitat generada / informe de la Guardia Urbana." }
            }
            return @{ Applies=$false; Reason="Art. 111.b: no obligatori l'estudi d'avaluacio de la mobilitat." }
        }
        'PAU' {
            if ($computed.PAU_ORGAN_KEY -eq 'CAT') {
                return @{ Applies=$true; Reason="Art. 57.d / Decret 30/2015 (Annex I, Cat. A): obligat a disposar de Pla d'Autoproteccio homologat per Proteccio civil de Catalunya." }
            }
            if ($computed.PAU_ORGAN_KEY -eq 'LOCAL') {
                return @{ Applies=$true; Reason="Art. 57.d / Decret 30/2015 (Annex I, Cat. B): obligat a disposar de Pla d'Autoproteccio homologat per Proteccio civil local." }
            }
            return @{ Applies=$false; Reason="Decret 30/2015: l'activitat no esta recollida al Cataleg (Annex I) -> no obligat a Pla d'Autoproteccio." }
        }
        'RC' {
            return @{ Applies=$true; Reason=("Art. 80 Decret 112/2010: tota activitat extraordinaria ha de disposar de polissa de RC. Quantia minima segons aforament ({0}): {1} EUR." -f $af, $computed.RC_IMPORT) }
        }
        'ASSIST_SANITARIA' {
            if ($computed.PAU_OBLIGAT) {
                return @{ Applies=$true; Reason="Annex III Decret 30/2015: en disposar de Pla d'Autoproteccio, els dispositius d'assistencia sanitaria son els que hi constin (segons l'Annex III del Decret 30/2015)." }
            }
            $tipus = if ($af -lt 1000) { 'farmaciola' } else { 'infermeria' }
            return @{ Applies=$true; Reason=("Art. 48 Decret 112/2010 (sense Pla d'Autoproteccio): dispositius d'assistencia sanitaria (aforament {0} -> {1})." -f $af, $tipus) }
        }
        'LASERS' {
            if ($decret.HiHaLasers -eq 'Si') {
                return @{ Applies=$true; Reason="L'activitat preveu l'us de lasers: cal acreditar l'autoritzacio del Departamento de Coordinacion Operativa del Espacio Aereo (proteccio de la navegacio aeria)." }
            }
            return @{ Applies=$false; Reason="L'activitat NO preveu l'us de lasers: es prohibeix l'emissio de qualsevol senyal luminica que pugui destorbar la navegacio aeria (no cal documentacio)." }
        }
        'CONTROLADORS' {
            if ($decret.ControlAccessos -eq 'Si') {
                return @{ Applies=$true; Reason=("Art. 57/58 Decret 112/2010: activitat musical a partir de 150 persones d'aforament -> cal personal de control d'acces ({0})." -f $computed.CONTROLADORS) }
            }
            return @{ Applies=$false; Reason="Art. 57/58: l'activitat no te l'obligacio de disposar de personal de control d'acces." }
        }
        'VIGILANTS' {
            if ($computed.VIGILANTS -ge 1) {
                return @{ Applies=$true; Reason=("Art. 43 Decret 112/2010: aforament {0} -> {1} vigilant(s) de seguretat privada." -f $af, $computed.VIGILANTS) }
            }
            return @{ Applies=$false; Reason="Art. 43: aforament inferior a 501 -> no calen vigilants de seguretat privada." }
        }
        'SERVEIS_HIGIENE' {
            $extra = if ($decret.EstablimentDotat -eq 'No') { " L'establiment NO esta dotat d'aquests equipaments -> cal instal-lar-ne de temporals." } else { '' }
            return @{ Applies=$true; Reason=("Art. 47 Decret 112/2010: proporcio minima de {0} lavabos i {1} cabines de vater.{2}" -f $computed.LAVABOS, $computed.CABINES, $extra) }
        }
        'IMPACTE_ACUSTIC' {
            return @{ Applies=$true; Reason="Art. 111.g Decret 112/2010: cal presentar una valoracio de l'impacte acustic de l'espectacle/activitat." }
        }
        'DISPONIBILITAT_ESPAI' {
            return @{ Applies=$true; Reason="Art. 111.i Decret 112/2010: cal acreditar la disponibilitat de l'establiment o de l'espai." }
        }
        default {
            if ($key -like 'MEMORIA_*') {
                $desc = switch ($key) {
                    'MEMORIA_A' { "Identificacio de l'espectacle public o activitat recreativa." }
                    'MEMORIA_B' { 'Data o dates i horari previst per a la realitzacio.' }
                    'MEMORIA_C' { 'Nom, cognoms, adreca i telefons de, com a minim, dues persones responsables.' }
                    'MEMORIA_D' { "Descripcio breu i nombre maxim de persones que assistiran/participaran." }
                    'MEMORIA_E' { 'Mesures adoptades (seguretat privada, control d acces, serveis municipals...).' }
                    'MEMORIA_F' { "Declaracio responsable de disposar de la polissa d'assegurances de RC." }
                    'MEMORIA_G' { "Identificacio dels titulars de la disponibilitat de l'espai." }
                    default     { '' }
                }
                return @{ Applies=$true; Reason=("Art. 113 Decret 112/2010 (memoria): {0}" -f $desc) }
            }
            return @{ Applies=$false; Reason='' }
        }
    }
}

# Construeix l'estat complet de la documentacio: per cada punt, aplicabilitat,
# motiu i si esta lliurat. $delivered es un hashtable/PSObject key->bool.
# Retorna una llista ordenada de PSCustomObject. Funcio PURA.
function Get-ActExtrStatus($decret, $computed, $delivered) {
    $isDelivered = {
        param($k)
        if ($null -eq $delivered) { return $false }
        if ($delivered -is [System.Collections.IDictionary]) {
            if ($delivered.Contains($k)) { return [bool]$delivered[$k] }
            return $false
        }
        if ($delivered.PSObject.Properties.Name -contains $k) { return [bool]$delivered.$k }
        return $false
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $script:ActExtrPoints) {
        $ap = Get-ActExtrPointApplicability $p.Key $decret $computed
        [void]$out.Add([pscustomobject]@{
            Key           = $p.Key
            Title         = $p.Title
            Applies       = [bool]$ap.Applies
            Reason        = [string]$ap.Reason
            NeedsDelivery = [bool]$p.NeedsDelivery
            Delivered     = [bool](& $isDelivered $p.Key)
        })
    }
    return $out
}

# Llista de claus de punt que generarien una deficiencia (apliquen i NO estan
# lliurades). Funcio PURA. Serveix per saber si cal emetre la intro del
# requeriment i si l'informe favorable es pot fer (sense pendents).
function Get-ActExtrDeficiencies($decret, $computed, $delivered) {
    $status = Get-ActExtrStatus $decret $computed $delivered
    $keys = New-Object System.Collections.ArrayList
    foreach ($s in $status) {
        if ($s.Applies -and -not $s.Delivered) { [void]$keys.Add($s.Key) }
    }
    return $keys.ToArray()
}

# ----------------------------------------------------------------------------
# Parseig de plantilla ACT_EXTR (mateix format que REQ1)
# ----------------------------------------------------------------------------
# Les plantilles ACT_EXTR_*.docx segueixen les MATEIXES convencions d'estil que
# REQ1 perque siguin igual de comodes d'editar a l'editor de catalegs:
#   - Titol 1 (Heading 1): titol de SECCIO. Nomes organitza el document; NO surt
#     a l'informe (a l'informe no hi ha titols de seccio).
#   - Titol 2 (Heading 2): obre un BLOC. El text es "[[KEY]] <tipus?> <titol>".
#       [[KEY]]    -> clau interna (aplicabilitat / valors).
#       ::TEXT::   -> el bloc es un paragraf de cos (intro, encapcalaments,
#                     tancament): surt SENSE pic ni numero.
#       ::CHILD::  -> el bloc es un sub-apartat: surt amb pic i sagnat.
#       ::NOTE::   -> sub-paragraf sagnat SENSE pic (nota).
#       ::LABEL::  -> etiqueta de subseccio (text normal amb espai a sota).
#       ::HEADER:: -> capcalera de conclusions (centrada i en negreta).
#       ::CONC::   -> paragraf de conclusio (justificat).
#       (cap)      -> el bloc es un item: numerat al requeriment, amb pic al
#                     favorable.
#     El <titol> es nomes una etiqueta per editar; NO surt a l'informe.
#   - Normal: el CONTINGUT del bloc (el text que surt a l'informe). Cada
#     paragraf Normal del bloc es un item/sub-item/paragraf segons el tipus.
#   - Cita (Quote): enllac (URL) del bloc.
$script:ActExtrKeyRegex = [regex]'^\s*\[\[([A-Z0-9_]+)\]\]\s*(.*)$'

# Analitza el text d'un marcador Titol 2: clau, tipus de render i etiqueta.
# Funcio PURA. Retorna @{ Key; Kind } o $null si no hi ha [[KEY]].
function _ParseActExtrMarker([string]$text) {
    $m = $script:ActExtrKeyRegex.Match([string]$text)
    if (-not $m.Success) { return $null }
    $key  = $m.Groups[1].Value
    $rest = $m.Groups[2].Value
    $kind = 'item'
    if     ($rest -match '::TEXT::')   { $kind = 'text' }
    elseif ($rest -match '::CHILD::')  { $kind = 'child' }
    elseif ($rest -match '::NOTE::')   { $kind = 'note' }
    elseif ($rest -match '::LABEL::')  { $kind = 'label' }
    elseif ($rest -match '::HEADER::') { $kind = 'header' }
    elseif ($rest -match '::CONC::')   { $kind = 'conc' }
    return @{ Key = $key; Kind = $kind }
}

# Construeix els blocs a partir de registres de paragraf @{ Text; Style } on
# Style es 'h1' | 'h2' | 'normal' | 'url'. Funcio PURA (sense Word). Retorna una
# llista de blocs @{ Key; Kind; Contents=@(@{Text;IsUrl},...) } en ordre.
function Build-ActExtrBlocks($paraRecords) {
    $blocks = New-Object System.Collections.ArrayList
    $current = $null
    # Index de seccio: s'incrementa a cada Titol 1. Serveix per saber, a
    # l'emissio de l'informe favorable, quan cal una linia en blanc (nomes al
    # canvi de seccio; dins d'una seccio els punts se separen amb SpaceBefore).
    $section = 0
    foreach ($r in $paraRecords) {
        $text  = [string]$r.Text
        $style = if ($r.Style) { [string]$r.Style } else { 'normal' }

        if ($style -eq 'h1') { $section++ }

        if ($style -eq 'h1' -or $style -eq 'h2') {
            $marker = _ParseActExtrMarker $text
            if ($null -eq $marker) { $current = $null; continue }  # titol visual sense clau
            $current = [pscustomobject]@{
                Key      = $marker.Key
                Kind     = $marker.Kind
                Section  = $section
                Contents = (New-Object System.Collections.ArrayList)
            }
            [void]$blocks.Add($current)
            continue
        }

        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($null -eq $current) { continue }   # contingut abans de cap [[KEY]]: s'ignora
        [void]$current.Contents.Add(@{ Text = $text.Trim(); IsUrl = ($style -eq 'url') })
    }
    return $blocks.ToArray()
}

# Mapeja una clau de BLOC de plantilla a la clau de PUNT (per consultar
# aplicabilitat i estat "lliurat"). Funcio PURA.
function _ActExtrBlockPoint([string]$blockKey) {
    switch ($blockKey) {
        'PAU_CAT'            { return 'PAU' }
        'PAU_LOCAL'          { return 'PAU' }
        'ASSIST'             { return 'ASSIST_SANITARIA' }
        'ASSIST_SANITARIA'   { return 'ASSIST_SANITARIA' }
        default              { return $blockKey }
    }
}

# Decideix si un bloc de plantilla s'ha d'incloure. $mode = 'req' | 'fav'.
# Funcio PURA. $ctx = @{ Decret; Computed; Delivered; StatusByKey; DefKeys }.
function Test-ActExtrIncludeBlock([string]$blockKey, [string]$mode, $ctx) {
    $isDelivered = {
        param($k)
        if ($null -eq $ctx.Delivered) { return $false }
        if ($ctx.Delivered -is [System.Collections.IDictionary]) {
            if ($ctx.Delivered.Contains($k)) { return [bool]$ctx.Delivered[$k] }
            return $false
        }
        if ($ctx.Delivered.PSObject.Properties.Name -contains $k) { return [bool]$ctx.Delivered.$k }
        return $false
    }
    $defKeys = @($ctx.DefKeys)

    # Lasers a l'informe favorable: dues variants segons si l'activitat en
    # preveu l'us. Cal decidir-ho ABANS del catch-all FAV_* (FAV_LASERS hi
    # entraria). Si hi ha lasers -> text d'autoritzacio (FAV_LASERS); si no
    # -> text de prohibicio (NO_LASERS).
    if ($blockKey -eq 'FAV_LASERS') { return ($mode -eq 'fav') -and [bool]$ctx.Computed.HAS_LASERS }
    if ($blockKey -eq 'NO_LASERS')  { return ($mode -eq 'fav') -and (-not [bool]$ctx.Computed.HAS_LASERS) }

    # Blocs estructurals/fixos del favorable (normativa, encapcalaments,
    # retols, conclusions...): hi son sempre que es generi l'informe favorable.
    if ($blockKey -like 'FAV_*') { return ($mode -eq 'fav') }

    switch ($blockKey) {
        'REQ_INTRO'      { return ($mode -eq 'req') -and ($defKeys.Count -gt 0) }
        'REQ_CLOSING'    { return ($mode -eq 'req') }
        'MEMORIA_HEADER' {
            if ($mode -ne 'req') { return $false }
            foreach ($k in $defKeys) { if ($k -like 'MEMORIA_*') { return $true } }
            return $false
        }
        'ASSIST_SANITARIA' {
            # Bloc del REQUERIMENT: nomes si encara no s'ha lliurat.
            if ($mode -ne 'req') { return $false }
            return (-not (& $isDelivered 'ASSIST_SANITARIA'))
        }
        'ASSIST' {
            # Bloc de l'informe FAVORABLE: l'assistencia sanitaria sempre hi es
            # (el text concret {{ASSISTENCIA}} el resol Get-ActExtrComputed).
            return ($mode -eq 'fav')
        }
        'PAU_CAT' {
            $applies = ($ctx.Computed.PAU_ORGAN_KEY -eq 'CAT')
            if (-not $applies) { return $false }
            if ($mode -eq 'req') { return (-not (& $isDelivered 'PAU')) }
            return $true
        }
        'PAU_LOCAL' {
            $applies = ($ctx.Computed.PAU_ORGAN_KEY -eq 'LOCAL')
            if (-not $applies) { return $false }
            if ($mode -eq 'req') { return (-not (& $isDelivered 'PAU')) }
            return $true
        }
        default {
            $pointKey = _ActExtrBlockPoint $blockKey
            $st = $ctx.StatusByKey[$pointKey]
            if ($null -eq $st) { return $false }
            if (-not $st.Applies) { return $false }
            # MOBILITAT nomes te text a l'informe favorable (no al requeriment).
            if ($blockKey -eq 'MOBILITAT' -and $mode -eq 'req') { return $false }
            # Els punts de la memoria (Art. 113) nomes surten al requeriment.
            if ($blockKey -like 'MEMORIA_*' -and $mode -eq 'fav') { return $false }
            if ($mode -eq 'req') { return (-not (& $isDelivered $pointKey)) }
            return $true
        }
    }
}

# ----------------------------------------------------------------------------
# Registre local (JSON, ignorat per git)
# ----------------------------------------------------------------------------
function Get-ActExtrRegistryPath {
    return (Join-Path $script:ActExtrRegistryDir $script:ActExtrRegistryFile)
}

function Load-ActExtrRegistry {
    $obj = Read-JsonFile (Get-ActExtrRegistryPath)
    if ($null -eq $obj) { return [pscustomobject]@{ Version = 1; Activitats = @() } }
    if ($null -eq $obj.Activitats) {
        Add-Member -InputObject $obj -NotePropertyName Activitats -NotePropertyValue @() -Force
    }
    # Forcem que Activitats sigui sempre un array (ConvertFrom-Json
    # desempaqueta els arrays d'1 element).
    $obj.Activitats = @($obj.Activitats)
    return $obj
}

function Save-ActExtrRegistry($registry) {
    $registry.Version = 1
    Write-JsonFile (Get-ActExtrRegistryPath) $registry 12
}

# Cerca una activitat al registre per ID GIA. Retorna el PSObject o $null.
function Get-ActExtrActivity($registry, [string]$id) {
    if ($null -eq $registry -or $null -eq $registry.Activitats) { return $null }
    foreach ($a in @($registry.Activitats)) {
        if ([string]$a.IdGia -eq [string]$id) { return $a }
    }
    return $null
}

# Insereix o actualitza una activitat al registre (per ID GIA). Retorna el
# registre modificat. Funcio (gairebe) PURA: nomes toca l'objecte en memoria.
function Set-ActExtrActivity($registry, $activity) {
    if ($null -eq $registry.Activitats) {
        Add-Member -InputObject $registry -NotePropertyName Activitats -NotePropertyValue @() -Force
    }
    $list = New-Object System.Collections.ArrayList
    $replaced = $false
    foreach ($a in @($registry.Activitats)) {
        if ([string]$a.IdGia -eq [string]$activity.IdGia) {
            [void]$list.Add($activity); $replaced = $true
        } else {
            [void]$list.Add($a)
        }
    }
    if (-not $replaced) { [void]$list.Add($activity) }
    $registry.Activitats = $list.ToArray()
    return $registry
}

# Recalcula l'estat global d'una activitat: 'tancat' si cap punt aplicable
# queda pendent (s'ha pogut fer / es pot fer l'informe favorable); 'pendent'
# altrament. Funcio PURA.
function Get-ActExtrActivityEstat($decret, $computed, $delivered) {
    $defs = Get-ActExtrDeficiencies $decret $computed $delivered
    if (@($defs).Count -eq 0) { return 'tancat' } else { return 'pendent' }
}

# ----------------------------------------------------------------------------
# Lectura de la plantilla amb Word (COM) -> registres de paragraf
# ----------------------------------------------------------------------------
# Llegeix una plantilla ACT_EXTR i en treu els blocs keyed.
#
# Nomes JSON (format estandard unic: nodes seccio/bloc amb el cos en "runs").
# Aqui hi havia un respatller que obria el .docx amb el Word i el parsejava pels
# estils; es va treure perque ACT_EXTR_REQ.docx / ACT_EXTR_FAV.docx ja no son
# fonts, son VISTES generades des dels JSON (VistaWord.ps1). El respatller no
# hauria fallat: hauria llegit la vista i hauria compost un informe
# silenciosament equivocat.
function Parse-ActExtrTemplate($path) {
    $jsonPath = if ([System.IO.Path]::GetExtension($path) -ieq '.json') { $path }
                else { [System.IO.Path]::ChangeExtension($path, '.json') }
    if (-not (Test-Path -LiteralPath $jsonPath)) {
        throw "No s'ha trobat la plantilla ACT_EXTR: $jsonPath"
    }
    $records = Read-ActExtrRecordsJson $jsonPath
    return (Build-ActExtrBlocks $records)
}
