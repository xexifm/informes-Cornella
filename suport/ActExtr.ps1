#requires -Version 5.1
<#
.SYNOPSIS
  Mode "ACT_EXTR": seguiment d'activitats extraordinaries de caracter
  esporadic (Decret 112/2010). Comprova la documentacio segons el Decret,
  decideix QUE cal i, segons si s'ha entregat o no, genera un REQUERIMENT
  (del que falta) o un INFORME FAVORABLE.

.DESCRIPTION
  Es un mode alternatiu del programa (es tria a la pantalla inicial, com el
  "seguiment"). A diferencia del flux REQ1 (on l'usuari tria deficiencies d'un
  cataleg), aqui el PROGRAMA sap que cal segons el Decret 112/2010 a partir de
  les respostes de classificacio i de l'aforament; l'usuari nomes marca, per
  cada punt aplicable, si la documentacio esta lliurada o pendent.

  Substitueix l'Excel "DATA_Act_Extr_NOM" (basat en el Decret 112/2010):
  la logica d'aplicabilitat i el calcul dels valors (vigilants, controladors,
  lavabos/cabines, p0lissa RC, organ del Pla d'Autoproteccio...) viuen aqui;
  els TEXTOS (requeriment i informe favorable) viuen a plantilles editables
  ESTRUCTURALS\ACT_EXTR_REQ.docx i ACT_EXTR_FAV.docx.

  Flux (Invoke-ActExtrFlow):
    1. Llistat local d'activitats extraordinaries amb el seu estat (pendent /
       tancat). Es pot crear-ne una de nova o triar-ne una d'existent.
    2. (Nova) Dades de capcalera ACT_EXTR (capcalera propia, diferent de REQ1):
       ID GIA, Exp., Adreca, Activitat, Titular, Dates i AFORAMENT.
    3. Comprovacio de la documentacio: respostes de classificacio del Decret +
       aforament. El programa mostra, per cada punt, si APLICA i PER QUE, i
       l'usuari marca si esta lliurat. Des d'aqui es pot generar el requeriment
       (del que falta) o l'informe favorable (si tot esta lliurat).

  Persistencia: un registre local (JSON, ignorat per git) guarda les dades de
  cada activitat (titular, adreca, aforament, respostes, estat de cada punt i
  historial de documents generats). Aixi cada vegada que es fa un requeriment o
  l'informe favorable es te a ma la "memoria de l'estat de l'activitat".

.NOTES
  Es carrega via dot-source des de GenerarInforme.ps1 (tambe en mode headless
  de proves). Les funcions PURES (logica del Decret, model, parseig de
  plantilla, inclusio de blocs) son testejables a Linux sense Word.

  Reutilitza de GenerarInforme.ps1 / Format.ps1: _NormalizeText, Test-StyleMatch,
  _SplitTextAndUrls, _ResolveOutputDir, _GetUniqueOutputPath, Apply-HeaderReplacements,
  _OpenOutputDocument, les funcions Format-* i $ReportFormatConfig.

  CONVENCIO ASCII: per evitar problemes d'encoding de PowerShell 5.1, el codi
  no porta accents. Tot el text accentuat que va als documents viu a les
  plantilles .docx (no al codi).
#>

# ----------------------------------------------------------------------------
# Localitzacio del registre local (JSON, ignorat per git)
# ----------------------------------------------------------------------------
# Per defecte, una carpeta 'BASE DE DADES ACT_EXTR' a l'arrel del clone (al
# costat de 'BASE DE DADES ACTIVITATS'). El .json de dins NO es puja (gitignore).
# Es pot sobreescriure $script:ActExtrRegistryDir des de config.ps1 o des dels
# tests. $RepoRoot el defineix GenerarInforme.ps1 abans de carregar aquest fitxer.
if (-not $script:ActExtrRegistryDir) {
    $script:ActExtrRegistryDir = if ($RepoRoot) { Join-Path $RepoRoot 'BASE DE DADES ACT_EXTR' } else { 'BASE DE DADES ACT_EXTR' }
}
$script:ActExtrRegistryFile = 'activitats-extraordinaries.json'

# Rutes de les plantilles ACT_EXTR (a ESTRUCTURALS). $EstructuralsDir el
# defineix GenerarInforme.ps1.
if ($EstructuralsDir) {
    $script:ActExtrReqTemplate = Join-Path $EstructuralsDir 'ACT_EXTR_REQ.docx'
    $script:ActExtrFavTemplate = Join-Path $EstructuralsDir 'ACT_EXTR_FAV.docx'
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
    $n = _NormalizeText ([string]$v)
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
# REQ1.docx perque siguin igual de comodes d'editar al Word:
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
    $path = Get-ActExtrRegistryPath
    if (-not (Test-Path -LiteralPath $path)) {
        return [pscustomobject]@{ Version = 1; Activitats = @() }
    }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{ Version=1; Activitats=@() } }
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj.Activitats) {
            Add-Member -InputObject $obj -NotePropertyName Activitats -NotePropertyValue @() -Force
        }
        # Forcem que Activitats sigui sempre un array (ConvertFrom-Json
        # desempaqueta els arrays d'1 element).
        $obj.Activitats = @($obj.Activitats)
        return $obj
    } catch {
        return [pscustomobject]@{ Version = 1; Activitats = @() }
    }
}

function Save-ActExtrRegistry($registry) {
    if (-not (Test-Path -LiteralPath $script:ActExtrRegistryDir)) {
        New-Item -ItemType Directory -Path $script:ActExtrRegistryDir -Force | Out-Null
    }
    $registry.Version = 1
    ($registry | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath (Get-ActExtrRegistryPath) -Encoding UTF8
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
# Llegeix una plantilla ACT_EXTR (.docx) i en treu els blocs keyed. Detecta
# l'estil de render de cada paragraf:
#   - 'Cita'/'Quote' o contingut http -> 'url'
#   - 'List Paragraph'/'Parrafo de lista'/'Llista...' -> 'list'
#   - altrament -> 'normal'
function Parse-ActExtrTemplate($word, $path) {
    # Si hi ha un JSON al costat (llista plana de paragrafs: estil + runs), el
    # fem servir (no cal Word): reconstruim els records i els passem al mateix
    # Build-ActExtrBlocks. Fallback segur al .docx si falla.
    $jsonPath = [System.IO.Path]::ChangeExtension($path, '.json')
    if (Test-Path -LiteralPath $jsonPath) {
        try {
            $o = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $records = New-Object System.Collections.ArrayList
            foreach ($p in @($o.paragrafs)) {
                [void]$records.Add(@{ Text = (_RunsToMarkup $p.runs); Style = [string]$p.estil })
            }
            return (Build-ActExtrBlocks $records)
        } catch {
            Write-Host "Avis: no s'ha pogut llegir '$jsonPath' ($($_.Exception.Message)); es fa servir el .docx."
        }
    }
    if (-not (Test-Path -LiteralPath $path)) {
        throw "No s'ha trobat la plantilla ACT_EXTR: $path"
    }
    $doc = $word.Documents.Open($path, $false, $true)  # ReadOnly
    try {
        $records = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            $text = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if ([string]::IsNullOrWhiteSpace($text)) {
                [void]$records.Add(@{ Text=''; Style='normal' })
                continue
            }
            $styleName = ''
            try { $styleName = $p.Style.NameLocal } catch { }
            $style = 'normal'
            if     (Test-StyleMatch $styleName 1) { $style = 'h1' }
            elseif (Test-StyleMatch $styleName 2) { $style = 'h2' }
            elseif ($styleName -match '^(Cita|Cite|Quote|Cita destacada|Quote intense)$') { $style = 'url' }
            # Fallback: una linia que es nomes un URL es tracta com a 'url'.
            if ($style -eq 'normal' -and $text.Trim() -match '^https?://') { $style = 'url' }
            [void]$records.Add(@{ Text=$text; Style=$style })
        }
        return (Build-ActExtrBlocks $records)
    } finally {
        $doc.Close($false)
    }
}

# ----------------------------------------------------------------------------
# Composicio del document (requeriment o informe favorable)
# ----------------------------------------------------------------------------
# Capcalera ACT_EXTR: hashtable amb les claus que espera Apply-HeaderReplacements
# mes <<DATES>> i <<AFORAMENT>>.
function _ActExtrHeaderMap($header) {
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] } ; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    return @{
        ID_GIA    = (& $get 'ID_GIA')
        EXP_NUM   = (& $get 'EXP_NUM')
        ADRECA    = (& $get 'ADRECA')
        ACTIVITAT = (& $get 'ACTIVITAT')
        TITULAR   = (& $get 'TITULAR')
        DATES     = (& $get 'DATES')
        AFORAMENT = (& $get 'AFORAMENT')
        # Claus de REQ1 que la capcalera ACT_EXTR no usa (per si de cas, buides).
        NUM_ANOTACIO  = ''
        DATA_ANOTACIO = ''
    }
}

# Nom del fitxer de sortida: YYYY-MM-DD_ActExtr-<Tipus>_GIA <id>_<Activitat>.docx
# S'hi inclou el nom de l'ACTIVITAT perque en un mateix establiment (mateix GIA)
# es poden fer diferents activitats extraordinaries i s'han de poder distingir.
function _GetActExtrOutputFileName([string]$tipus, [string]$gia, [string]$activitat) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    $name = "{0}_ActExtr-{1}_GIA {2}" -f $today, $tipus, $gia
    $act = ([string]$activitat -replace '[\\/:*?"<>|]','_').Trim()
    if (-not [string]::IsNullOrWhiteSpace($act)) { $name += "_$act" }
    return ($name + '.docx')
}

# Emet el cos del document a partir dels blocs de la plantilla. Segons el
# marcador Titol 2 (Kind), cada paragraf de contingut es renderitza diferent:
#   'item'   -> item: NUMERAT al requeriment; amb PIC (vinyeta) al favorable.
#   'child'  -> sub-apartat amb PIC i sagnat.
#   'note'   -> sub-paragraf sagnat SENSE pic (nota).
#   'text'   -> paragraf de cos (sense pic ni numero).
#   'header' -> capcalera de conclusions (centrada i en negreta).
#   'conc'   -> paragraf de conclusio (justificat).
#
# ESPAIAT (diferent en cada mode, per reproduir el format de referencia):
#   - REQUERIMENT: cada unitat (item o paragraf) va separada per una linia en
#     blanc (com REQ1). Sub-items i URLs s'enganxen a la unitat anterior.
#   - FAVORABLE: els punts d'una mateixa seccio se separen amb SpaceBefore (no
#     linies en blanc); nomes s'insereix una linia en blanc al CANVI de seccio
#     (Titol 1 de la plantilla). Aixi la llista surt compacta, com el document
#     de referencia.
function _WriteActExtrBody($sel, $blocks, $mode, $ctx, $computed) {
    if ($mode -eq 'fav') { _WriteActExtrBodyFav $sel $blocks $ctx $computed; return }

    # ---- REQUERIMENT ----
    $num0 = 0   # comptador d'items de 1r nivell (continu)
    $first = $true
    foreach ($block in $blocks) {
        if (-not (Test-ActExtrIncludeBlock $block.Key $mode $ctx)) { continue }
        $kind = [string]$block.Kind
        foreach ($c in $block.Contents) {
            $resolved = Resolve-ActExtrTokens ([string]$c.Text) $computed

            if ($c.IsUrl) {
                $u = $resolved
                if ($u.StartsWith('[[URL]] ')) { $u = $u.Substring('[[URL]] '.Length).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($u)) {
                    if ($kind -eq 'child') { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
                }
                continue
            }

            $parts = _SplitTextAndUrls $resolved
            if ([string]::IsNullOrWhiteSpace($parts.Text) -and @($parts.Urls).Count -eq 0) { continue }

            if ($kind -eq 'child') {
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { Format-Bullet $sel $parts.Text -IsChild }
                foreach ($x in $parts.Urls) { Format-Url $sel $x -IsChild }
                $first = $false
                continue
            }

            # Unitat nova (item o paragraf de cos): linia en blanc al davant si
            # no es la primera unitat del document.
            if (-not $first) { Format-Spacer $sel }
            if ($kind -eq 'item') {
                $num0++
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { Format-Item $sel "$num0." $parts.Text }
                foreach ($x in $parts.Urls) { Format-Url $sel $x }
            } else {
                if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { Format-Body $sel $parts.Text }
                foreach ($x in $parts.Urls) { Format-Url $sel $x }
            }
            $first = $false
        }
    }
}

# Emissio del cos de l'INFORME FAVORABLE (espaiat per seccions; vegeu
# _WriteActExtrBody). Els punts van amb PIC i separats per SpaceBefore; entre
# seccions (canvi de Titol 1 de la plantilla) s'hi posa una linia en blanc.
function _WriteActExtrBodyFav($sel, $blocks, $ctx, $computed) {
    $firstBlock  = $true
    $prevSection = -1
    foreach ($block in $blocks) {
        if (-not (Test-ActExtrIncludeBlock $block.Key 'fav' $ctx)) { continue }
        $kind = [string]$block.Kind
        # Linia en blanc nomes al canvi de seccio (no dins d'una seccio).
        if ((-not $firstBlock) -and ([int]$block.Section -ne [int]$prevSection)) { Format-Spacer $sel }

        foreach ($c in $block.Contents) {
            $resolved = Resolve-ActExtrTokens ([string]$c.Text) $computed

            if ($c.IsUrl) {
                $u = $resolved
                if ($u.StartsWith('[[URL]] ')) { $u = $u.Substring('[[URL]] '.Length).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($u)) {
                    if ($kind -eq 'child') { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
                }
                continue
            }

            $parts = _SplitTextAndUrls $resolved
            if ([string]::IsNullOrWhiteSpace($parts.Text) -and @($parts.Urls).Count -eq 0) { continue }
            $txt = $parts.Text

            switch ($kind) {
                'child'  { if ($txt) { Format-Bullet $sel $txt -IsChild }; foreach ($x in $parts.Urls) { Format-Url $sel $x -IsChild } }
                'note'   { if ($txt) { Format-Note $sel $txt };            foreach ($x in $parts.Urls) { Format-Url $sel $x -IsChild } }
                'label'  { if ($txt) { Format-Label $sel $txt };           foreach ($x in $parts.Urls) { Format-Url $sel $x } }
                'header' { if ($txt) { Format-ConclusionHeader $sel $txt } }
                'conc'   { if ($txt) { Format-Conclusion $sel $txt };      foreach ($x in $parts.Urls) { Format-Url $sel $x } }
                'text'   { if ($txt) { Format-Body $sel $txt };            foreach ($x in $parts.Urls) { Format-Url $sel $x } }
                default  { if ($txt) { Format-Bullet $sel $txt };          foreach ($x in $parts.Urls) { Format-Url $sel $x } }  # 'item'
            }
        }
        $prevSection = [int]$block.Section
        $firstBlock  = $false
    }
}

# Genera el document (requeriment o informe favorable) i retorna la ruta del
# .docx. $mode = 'req' | 'fav'. Reutilitza la capcalera 0 CAPCALERA.docx
# (bloc ACT_EXTR) i el motor de format.
function Build-ActExtrDocument($word, $header, $decret, $delivered, $mode) {
    $computed = Get-ActExtrComputed $decret
    $status   = Get-ActExtrStatus $decret $computed $delivered
    $statusByKey = @{}
    foreach ($s in $status) { $statusByKey[$s.Key] = $s }
    $defKeys = Get-ActExtrDeficiencies $decret $computed $delivered
    $ctx = @{ Decret=$decret; Computed=$computed; Delivered=$delivered; StatusByKey=$statusByKey; DefKeys=$defKeys }

    $tplPath = if ($mode -eq 'fav') { $script:ActExtrFavTemplate } else { $script:ActExtrReqTemplate }
    $blocks  = Parse-ActExtrTemplate $word $tplPath

    $tipus = if ($mode -eq 'fav') { 'Fav' } else { 'Req' }
    $gia = if ($header -is [System.Collections.IDictionary]) { [string]$header['ID_GIA'] } else { [string]$header.ID_GIA }
    $act = if ($header -is [System.Collections.IDictionary]) { [string]$header['ACTIVITAT'] } else { [string]$header.ACTIVITAT }
    $baseName  = _GetActExtrOutputFileName $tipus $gia $act
    $targetDir = _ResolveOutputDir
    $outPath   = _GetUniqueOutputPath $targetDir $baseName
    $fileName  = [System.IO.Path]::GetFileName($outPath)

    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath
    try {
        # Retallem la capcalera per quedar-nos nomes amb el bloc ACT_EXTR.
        Select-CapcaleraBlock $doc 'ACT_EXTR'
        Apply-HeaderReplacements -doc $doc -header (_ActExtrHeaderMap $header)

        $doc.Activate()
        $sel = $word.Selection
        [void]$sel.EndKey(6)  # wdStory = 6
        # El cos hereta l'estil 'List Paragraph' del darrer paragraf de la
        # capcalera ACT_EXTR (igual que REQ1), que ja resol a Bookman Old Style.
        # Aixi el format surt del motor Format.ps1 i de l'estil, no d'un override.

        _WriteActExtrBody $sel $blocks $mode $ctx $computed

        $doc.Save()
        $doc.Close($false)
    } catch {
        try { $doc.Close($false) } catch { }
        throw
    }
    try { Move-Item -LiteralPath $tempPath -Destination $outPath -Force } catch { return $tempPath }
    return $outPath
}

# ----------------------------------------------------------------------------
# UI (WinForms) - nomes en mode no-headless
# ----------------------------------------------------------------------------
# Llistat d'activitats extraordinaries amb el seu estat. Retorna:
#   @{ Action='new' } | @{ Action='open'; Id=<id> } | @{ Action='exit' }
function Show-ActExtrList($registry) {
    $form = _NewForm
    $form.Text = 'Activitats extraordinaries (ACT_EXTR)'
    $form.Size = New-Object System.Drawing.Size(820, 540)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Activitats extraordinaries. A la dreta hi ha l''estat (PENDENT si falta documentacio, TANCAT si tot esta lliurat).'
    $lbl.Location = New-Object System.Drawing.Point(15, 12)
    $lbl.Size = New-Object System.Drawing.Size(780, 20)
    $form.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(15, 40)
    $lv.Size = New-Object System.Drawing.Size(780, 400)
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.MultiSelect = $false
    $lv.Anchor = 'Top, Bottom, Left, Right'
    [void]$lv.Columns.Add('ID GIA', 90)
    [void]$lv.Columns.Add('Titular', 240)
    [void]$lv.Columns.Add('Activitat', 230)
    [void]$lv.Columns.Add('Dates', 120)
    [void]$lv.Columns.Add('Estat', 90)
    foreach ($a in @($registry.Activitats)) {
        $h = $a.Header
        $it = New-Object System.Windows.Forms.ListViewItem([string]$a.IdGia)
        [void]$it.SubItems.Add([string]$h.TITULAR)
        [void]$it.SubItems.Add([string]$h.ACTIVITAT)
        [void]$it.SubItems.Add([string]$h.DATES)
        $estat = if ($a.Estat) { [string]$a.Estat } else { 'pendent' }
        [void]$it.SubItems.Add($estat.ToUpper())
        $it.Tag = [string]$a.IdGia
        if ($estat -eq 'pendent') { $it.ForeColor = [System.Drawing.Color]::Firebrick }
        else                      { $it.ForeColor = [System.Drawing.Color]::ForestGreen }
        [void]$lv.Items.Add($it)
    }
    $form.Controls.Add($lv)

    # Enrere (torna al menu inicial) SEMPRE a baix a l'esquerra.
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(15, 455)
    $btnBack.Size = New-Object System.Drawing.Size(90, 32)
    $btnBack.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnBack)

    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = 'Nova activitat'
    $btnNew.Location = New-Object System.Drawing.Point(115, 455)
    $btnNew.Size = New-Object System.Drawing.Size(150, 32)
    $btnNew.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnNew)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = 'Obrir / continuar'
    $btnOpen.Location = New-Object System.Drawing.Point(275, 455)
    $btnOpen.Size = New-Object System.Drawing.Size(160, 32)
    $btnOpen.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnOpen)

    # 'exit' (Enrere o tancar la finestra) fa que Invoke-ActExtrFlow torni al
    # menu inicial (el programa no es tanca; nomes es tanca des del Pas 1).
    $result = @{ Action = 'exit'; Id = $null }
    $btnNew.add_Click({ $result.Action = 'new'; $form.DialogResult = 'OK'; $form.Close() })
    $openAction = {
        if ($lv.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Tria una activitat de la llista.','Sense seleccio','OK','Information') | Out-Null
            return
        }
        $result.Action = 'open'; $result.Id = [string]$lv.SelectedItems[0].Tag
        $form.DialogResult = 'OK'; $form.Close()
    }
    $btnOpen.add_Click($openAction)
    $lv.add_DoubleClick($openAction)
    $btnBack.add_Click({ $result.Action = 'exit'; $form.DialogResult = 'Cancel'; $form.Close() })

    [void]$form.ShowDialog()
    return $result
}

# Capcalera ACT_EXTR (entrada manual). $preload pot precarregar valors.
# Retorna @{ Nav='next'|'back'; Data=@{ ID_GIA;EXP_NUM;ADRECA;ACTIVITAT;TITULAR;DATES;AFORAMENT } }
function Get-ActExtrHeader {
    param($preload = $null, [bool]$lockId = $false)
    $form = _NewForm
    $form.Text = 'Activitat extraordinaria - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(700, 420)
    $form.StartPosition = 'CenterScreen'

    $controls = @{}
    # $addRow afegeix una etiqueta + caixa de text a la posicio $y i retorna
    # la $y de la fila seguent (patro explicit, sense estat compartit).
    $addRow = {
        param($label, $key, $width, $yPos)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $label
        $l.Location = New-Object System.Drawing.Point(15, $yPos)
        $l.Size = New-Object System.Drawing.Size(180, 22)
        [void]$form.Controls.Add($l)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(200, ($yPos - 2))
        $tb.Size = New-Object System.Drawing.Size($width, 22)
        [void]$form.Controls.Add($tb)
        $controls[$key] = $tb
        return ($yPos + 36)
    }
    $y = 20
    $y = & $addRow 'ID GIA' 'ID_GIA' 300 $y
    $y = & $addRow "Num. d'expedient" 'EXP_NUM' 450 $y
    $y = & $addRow 'Titular' 'TITULAR' 450 $y
    $y = & $addRow 'Adreca (carrer i numero)' 'ADRECA' 450 $y
    $y = & $addRow 'Activitat' 'ACTIVITAT' 450 $y
    $y = & $addRow 'Dates' 'DATES' 450 $y
    $y = & $addRow 'Aforament autoritzat' 'AFORAMENT' 150 $y

    if ($preload) {
        foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','DATES','AFORAMENT') {
            $v = $null
            if ($preload -is [System.Collections.IDictionary]) { if ($preload.Contains($k)) { $v = $preload[$k] } }
            elseif ($preload.PSObject.Properties.Name -contains $k) { $v = $preload.$k }
            if ($null -ne $v) { $controls[$k].Text = [string]$v }
        }
    }
    if ($lockId) { $controls['ID_GIA'].ReadOnly = $true }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    [void]$form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(575, $y)
    $ok.Size = New-Object System.Drawing.Size(95, 30)
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    $data = $null
    $ok.add_Click({
        if ([string]::IsNullOrWhiteSpace($controls['ID_GIA'].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return
        }
        $af = 0
        if (-not [int]::TryParse((($controls['AFORAMENT'].Text) -replace '[^\d]',''), [ref]$af) -or $af -le 0) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un aforament autoritzat (nombre).",'Falta aforament','OK','Warning') | Out-Null
            return
        }
        $script:_actExtrHeaderData = @{
            ID_GIA    = $controls['ID_GIA'].Text.Trim()
            EXP_NUM   = $controls['EXP_NUM'].Text.Trim()
            TITULAR   = $controls['TITULAR'].Text.Trim()
            ADRECA    = $controls['ADRECA'].Text.Trim()
            ACTIVITAT = $controls['ACTIVITAT'].Text.Trim()
            DATES     = $controls['DATES'].Text.Trim()
            AFORAMENT = [string]$af
        }
        $form.DialogResult = 'OK'; $form.Close()
    })

    $script:_actExtrHeaderData = $null
    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { return [pscustomobject]@{ Nav='back' } }
    return [pscustomobject]@{ Nav='next'; Data=$script:_actExtrHeaderData }
}

# Pas 3: comprovacio de la documentacio. Mostra les preguntes de classificacio
# (esquerra) i, per cada punt, si APLICA i PER QUE + casella "lliurat" (dreta).
# Retorna @{ Action='req'|'fav'|'save'|'back'; Answers=@{...}; Delivered=@{...} }
function Edit-ActExtrDocumentacio {
    param($header, $answers = $null, $delivered = $null)

    $form = _NewForm
    $form.Text = 'Comprovacio de la documentacio - ' + ([string]$header.ID_GIA)
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(1160, 680)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)

    # Estat compartit en $script: (sempre accessible des de qualsevol handler,
    # independentment de l'abast d'invocacio). Aixi s'eviten els problemes
    # d'abast dels scriptblocks niats de WinForms.
    $script:_actExtrDelivered = @{}
    foreach ($p in $script:ActExtrPoints) {
        $script:_actExtrDelivered[$p.Key] = [bool](_GetActExtrDelivered $delivered $p.Key)
    }
    $script:_actExtrResult = @{ Action='back'; Answers=$null; Delivered=$null }

    # --- Barra inferior de botons (ancorada a baix; sempre visible) ---
    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = 'Bottom'
    $bottom.Height = 50
    $form.Controls.Add($bottom)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(12, 9)
    $btnBack.Size = New-Object System.Drawing.Size(90, 32)
    $bottom.Controls.Add($btnBack)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(110, 9)
    $btnSave.Size = New-Object System.Drawing.Size(110, 32)
    $bottom.Controls.Add($btnSave)

    $btnReq = New-Object System.Windows.Forms.Button
    $btnReq.Text = 'Generar requeriment'
    $btnReq.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 390), 9)
    $btnReq.Size = New-Object System.Drawing.Size(185, 32)
    $btnReq.Anchor = 'Top, Right'
    $bottom.Controls.Add($btnReq)

    $btnFav = New-Object System.Windows.Forms.Button
    $btnFav.Text = 'Generar informe favorable'
    $btnFav.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 200), 9)
    $btnFav.Size = New-Object System.Drawing.Size(190, 32)
    $btnFav.Anchor = 'Top, Right'
    $bottom.Controls.Add($btnFav)

    # --- Esquerra: preguntes de classificacio (amb scroll propi) ---
    $gbQ = New-Object System.Windows.Forms.GroupBox
    $gbQ.Text = 'Classificacio (Decret 112/2010)'
    $gbQ.Location = New-Object System.Drawing.Point(12, 12)
    $gbQ.Size = New-Object System.Drawing.Size(430, ($form.ClientSize.Height - 50 - 24))
    $gbQ.Anchor = 'Top, Bottom, Left'
    $form.Controls.Add($gbQ)

    # Panell intern desplacable: hi caben totes les preguntes encara que la
    # finestra sigui curta.
    $qPanel = New-Object System.Windows.Forms.Panel
    $qPanel.Location = New-Object System.Drawing.Point(8, 20)
    $qPanel.Size = New-Object System.Drawing.Size(414, ($gbQ.Height - 28))
    $qPanel.AutoScroll = $true
    $qPanel.Anchor = 'Top, Bottom, Left, Right'
    $gbQ.Controls.Add($qPanel)

    $aforamentInit = if ($answers) { [string](_GetActExtrAnswer $answers 'Aforament') } else { '' }
    if ([string]::IsNullOrWhiteSpace($aforamentInit)) { $aforamentInit = [string]$header.AFORAMENT }

    $lblAf = New-Object System.Windows.Forms.Label
    $lblAf.Text = 'Aforament autoritzat:'
    $lblAf.Location = New-Object System.Drawing.Point(8, 10)
    $lblAf.Size = New-Object System.Drawing.Size(160, 22)
    $qPanel.Controls.Add($lblAf)
    $tbAf = New-Object System.Windows.Forms.TextBox
    $tbAf.Location = New-Object System.Drawing.Point(175, 7)
    $tbAf.Size = New-Object System.Drawing.Size(110, 22)
    $tbAf.Text = $aforamentInit
    $qPanel.Controls.Add($tbAf)
    $script:_actExtrAfBox = $tbAf

    # Preguntes Si/No (clau -> etiqueta)
    $questions = @(
        @{ Key='Incendis';          Label='Incendis: inclosa a l''Art. 23 Llei 3/2010?' }
        @{ Key='Mobilitat';         Label='Mobilitat: cal estudi (Decret 344/2006)?' }
        @{ Key='ControlAccessos';   Label='Control d''acces: musical >=150 aforament?' }
        @{ Key='PauCatalunya';      Label='PAU: al Cataleg de Catalunya (Annex I, Cat. A)?' }
        @{ Key='PauLocal';          Label='PAU: al Cataleg local (Annex I, Cat. B)?' }
        @{ Key='EstablimentDotat';  Label='Higiene: l''establiment ja esta dotat de lavabos?' }
        @{ Key='ParcialSotaRasant'; Label='RC: l''activitat es du PARCIALMENT sota rasant?' }
        @{ Key='TotalSotaRasant';   Label='RC: l''activitat es du TOTALMENT sota rasant?' }
        @{ Key='HiHaLasers';        Label='Lasers: l''activitat preveu l''us de lasers?' }
    )
    $qRadios = @{}
    $qy = 44
    foreach ($q in $questions) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $q.Label
        $l.Location = New-Object System.Drawing.Point(8, $qy)
        $l.Size = New-Object System.Drawing.Size(395, 30)
        $qPanel.Controls.Add($l)
        $qy += 30
        # Cada parella Si/No va al SEU PROPI panell perque els RadioButton nomes
        # siguin exclusius DINS de la pregunta (si no, tot el contenidor formaria
        # un sol grup i nomes es podria triar una resposta a tota la columna).
        $rp = New-Object System.Windows.Forms.Panel
        $rp.Location = New-Object System.Drawing.Point(20, $qy)
        $rp.Size = New-Object System.Drawing.Size(200, 26)
        $rbSi = New-Object System.Windows.Forms.RadioButton
        $rbSi.Text = 'Si'
        $rbSi.Location = New-Object System.Drawing.Point(0, 2)
        $rbSi.Size = New-Object System.Drawing.Size(60, 22)
        $rbNo = New-Object System.Windows.Forms.RadioButton
        $rbNo.Text = 'No'
        $rbNo.Location = New-Object System.Drawing.Point(70, 2)
        $rbNo.Size = New-Object System.Drawing.Size(60, 22)
        $cur = if ($answers) { _ActExtrYesNo (_GetActExtrAnswer $answers $q.Key) } else { 'No' }
        if ($cur -eq 'Si') { $rbSi.Checked = $true } else { $rbNo.Checked = $true }
        $rp.Controls.Add($rbSi); $rp.Controls.Add($rbNo)
        $qPanel.Controls.Add($rp)
        $qRadios[$q.Key] = @{ Si=$rbSi; No=$rbNo }
        $qy += 34
    }
    $script:_actExtrRadios = $qRadios

    # --- Dreta: estat per punt (aplica + per que + lliurat) ---
    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Text = 'Documentacio per punt (APLICA / per que / lliurat):'
    $lblD.Location = New-Object System.Drawing.Point(455, 12)
    $lblD.Size = New-Object System.Drawing.Size(680, 20)
    $lblD.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lblD)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(455, 36)
    $panel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 455 - 12), ($form.ClientSize.Height - 50 - 48))
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($panel)

    # Llegeix les respostes actuals dels controls (via $script:, robust).
    $readAnswers = {
        $h = @{ Aforament = ($script:_actExtrAfBox.Text -replace '[^\d]','') }
        foreach ($k in $script:_actExtrRadios.Keys) {
            $h[$k] = if ($script:_actExtrRadios[$k].Si.Checked) { 'Si' } else { 'No' }
        }
        return $h
    }
    $script:_actExtrReadAnswers = $readAnswers

    # Refresca el panell de la dreta segons les respostes actuals.
    $refresh = {
        $ans = & $script:_actExtrReadAnswers
        $decret = Build-ActExtrDecret $ans
        $computed = Get-ActExtrComputed $decret
        $status = Get-ActExtrStatus $decret $computed $script:_actExtrDelivered
        $panel.SuspendLayout()
        $panel.Controls.Clear()
        $yy = 8
        foreach ($s in $status) {
            $title = New-Object System.Windows.Forms.Label
            $applyTxt = if ($s.Applies) { 'APLICA' } else { 'no aplica' }
            $title.Text = ('{0}  -  {1}' -f $s.Title, $applyTxt)
            $title.Location = New-Object System.Drawing.Point(8, $yy)
            $title.Size = New-Object System.Drawing.Size(640, 20)
            $title.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $title.ForeColor = if ($s.Applies) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::Gray }
            [void]$panel.Controls.Add($title)
            $yy += 22

            $reason = New-Object System.Windows.Forms.Label
            $reason.Text = $s.Reason
            $reason.Location = New-Object System.Drawing.Point(20, $yy)
            $reason.Size = New-Object System.Drawing.Size(645, 42)
            $reason.ForeColor = [System.Drawing.Color]::DimGray
            [void]$panel.Controls.Add($reason)
            $yy += 44

            if ($s.Applies -and $s.NeedsDelivery) {
                $cb = New-Object System.Windows.Forms.CheckBox
                $cb.Text = 'Lliurat correctament'
                $cb.Location = New-Object System.Drawing.Point(20, $yy)
                $cb.Size = New-Object System.Drawing.Size(300, 22)
                $cb.Checked = [bool]$script:_actExtrDelivered[$s.Key]
                $cb.Tag = $s.Key
                # Handler robust: el sender ve com a parametre i l'estat es
                # $script:_actExtrDelivered (no depen de l'abast d'invocacio).
                $cb.add_CheckedChanged({
                    param($snd, $e)
                    $script:_actExtrDelivered[[string]$snd.Tag] = [bool]$snd.Checked
                })
                [void]$panel.Controls.Add($cb)
                $yy += 28
            }
            $sepY = $yy + 2
            $sep = New-Object System.Windows.Forms.Label
            $sep.BorderStyle = 'Fixed3D'
            $sep.Location = New-Object System.Drawing.Point(8, $sepY)
            $sep.Size = New-Object System.Drawing.Size(650, 2)
            [void]$panel.Controls.Add($sep)
            $yy = $sepY + 10
        }
        $panel.ResumeLayout()
    }
    $script:_actExtrRefresh = $refresh

    $tbAf.add_TextChanged({ & $script:_actExtrRefresh })
    foreach ($k in $qRadios.Keys) {
        $qRadios[$k].Si.add_CheckedChanged({ & $script:_actExtrRefresh })
    }

    $finish = {
        param($action)
        $script:_actExtrResult = @{
            Action    = $action
            Answers   = (& $script:_actExtrReadAnswers)
            Delivered = $script:_actExtrDelivered.Clone()
        }
        $form.DialogResult = 'OK'; $form.Close()
    }
    $script:_actExtrFinish = $finish

    $btnSave.add_Click({ & $script:_actExtrFinish 'save' })
    $btnReq.add_Click({ & $script:_actExtrFinish 'req' })
    $btnFav.add_Click({
        $ans = & $script:_actExtrReadAnswers
        $decret = Build-ActExtrDecret $ans
        $computed = Get-ActExtrComputed $decret
        $defs = Get-ActExtrDeficiencies $decret $computed $script:_actExtrDelivered
        if (@($defs).Count -gt 0) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                "Encara hi ha punts aplicables PENDENTS de lliurar.`n`nL'informe favorable dona per fet que tot esta lliurat. Vols generar-lo igualment?",
                'Hi ha pendents', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }
        & $script:_actExtrFinish 'fav'
    })
    $btnBack.add_Click({ $script:_actExtrResult = @{ Action='back' }; $form.DialogResult='Cancel'; $form.Close() })

    & $refresh
    [void]$form.ShowDialog()
    return $script:_actExtrResult
}

# Helpers d'acces a respostes/lliurats precarregats (hashtable o PSObject).
function _GetActExtrAnswer($answers, [string]$key) {
    if ($null -eq $answers) { return $null }
    if ($answers -is [System.Collections.IDictionary]) { if ($answers.Contains($key)) { return $answers[$key] }; return $null }
    if ($answers.PSObject.Properties.Name -contains $key) { return $answers.$key }
    return $null
}
function _GetActExtrDelivered($delivered, [string]$key) {
    if ($null -eq $delivered) { return $false }
    if ($delivered -is [System.Collections.IDictionary]) { if ($delivered.Contains($key)) { return [bool]$delivered[$key] }; return $false }
    if ($delivered.PSObject.Properties.Name -contains $key) { return [bool]$delivered.$key }
    return $false
}

# ----------------------------------------------------------------------------
# Orquestrador del mode ACT_EXTR
# ----------------------------------------------------------------------------
function Invoke-ActExtrFlow {
    $registry = Load-ActExtrRegistry
    $word = $null
    try {
        while ($true) {
            $sel = Show-ActExtrList $registry
            if ($sel.Action -eq 'exit') { break }

            $header = $null; $answers = $null; $delivered = $null
            if ($sel.Action -eq 'new') {
                $h = Get-ActExtrHeader
                if ($h.Nav -ne 'next') { continue }
                $header = $h.Data
            } else {
                $act = Get-ActExtrActivity $registry $sel.Id
                if ($null -eq $act) { continue }
                $header    = $act.Header
                $answers   = $act.Decret
                $delivered = $act.Punts
                # Permet revisar/editar la capcalera abans del Pas 3.
                $h = Get-ActExtrHeader -preload $header -lockId $true
                if ($h.Nav -eq 'next') { $header = $h.Data }
            }

            # Pas 3 (es pot repetir fins que es tanqui amb Enrere)
            $stay = $true
            while ($stay) {
                $r = Edit-ActExtrDocumentacio -header $header -answers $answers -delivered $delivered
                if ($r.Action -eq 'back') { $stay = $false; break }
                $answers   = $r.Answers
                $delivered = $r.Delivered

                $decret   = Build-ActExtrDecret $answers
                $computed = Get-ActExtrComputed $decret
                $estat    = Get-ActExtrActivityEstat $decret $computed $delivered

                # Desem/actualitzem l'activitat al registre. Fem servir un
                # ArrayList per a l'historial: amb '+=' sobre el resultat d'un
                # 'if' de @(...) d'un sol element, PowerShell el desempaqueta a
                # escalar i el '+=' peta ("op_Addition" sobre PSObject).
                $existing = Get-ActExtrActivity $registry ([string]$header.ID_GIA)
                $historial = New-Object System.Collections.ArrayList
                if ($existing -and $existing.Historial) {
                    foreach ($h in @($existing.Historial)) { if ($null -ne $h) { [void]$historial.Add($h) } }
                }
                $creatAt = if ($existing -and $existing.CreatAt) { [string]$existing.CreatAt } else { (Get-Date).ToString('o') }

                $outPath = $null
                if ($r.Action -in @('req','fav')) {
                    if ($null -eq $word) { $word = New-WordApp }
                    $outPath = Build-ActExtrDocument $word $header $decret $delivered $r.Action
                    $tipus = if ($r.Action -eq 'fav') { 'favorable' } else { 'requeriment' }
                    [void]$historial.Add([pscustomobject]@{
                        Data    = (Get-Date).ToString('o')
                        Tipus   = $tipus
                        Fitxer  = $outPath
                    })
                }

                $activity = [pscustomobject]@{
                    IdGia      = [string]$header.ID_GIA
                    Header     = $header
                    Decret     = $answers
                    Punts      = $delivered
                    Estat      = $estat
                    CreatAt    = $creatAt
                    ModificatAt= (Get-Date).ToString('o')
                    Historial  = @($historial.ToArray())
                }
                $registry = Set-ActExtrActivity $registry $activity
                Save-ActExtrRegistry $registry

                if ($r.Action -in @('req','fav')) {
                    [System.Windows.Forms.MessageBox]::Show("Document generat:`n$outPath",'Finalitzat','OK','Information') | Out-Null
                    $word.Visible = $true
                    $word.Documents.Open($outPath) | Out-Null
                    # En obrir Word per a l'usuari, no el tanquem; sortim al llistat.
                    $word = $null
                    $stay = $false
                } else {
                    # 'save': nomes desar; tornem al Pas 3 amb les dades desades.
                    [System.Windows.Forms.MessageBox]::Show('Dades desades al registre.','Desat','OK','Information') | Out-Null
                }
            }
        }
    } finally {
        if ($null -ne $word) { Close-WordApp $word }
    }
}
