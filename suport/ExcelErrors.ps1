#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Excel Errors" (menu, fila GIA): els errors de la base de dades
  d'activitats (la fulla "Estes" de l'Excel del GIA), per corregir-los al GIA.

.DESCRIPTION
  L'usuari (octubre 2026): "una nova eina que comprovi errors de la base de
  dades d'activitats [...] que falta alguna dada de les importants (rao social,
  representant legal, adreca, cadastre, UTM, classificacio, activitat), dades
  del tecnic a representant legal, etc. [...] Aquests errors es podran extreure
  en un Excel indicant el ID Activitat i el tipus d'error (deixa'm marcar el
  tipus d'error que vull que s'exporti)".

  QUE ES COMPROVA I D'ON SURT. Cada tipus d'error es una clau de
  $Script:ExTipus (l'UNICA llista: la fan servir la finestra, la tria desada i
  l'exportacio). La majoria surten de la fila de l'Excel tal qual
  (Get-ExcelErrorsFila, PURA). Dos venen de llocs que el programa ja tenia, i
  NO es tornen a escriure aqui:
    - les ERRADES de format d'un correu, un telefon o un NIF: _CtErrada i
      _CtTipusNif (ContactesRegles / ContactesExtraccio), les mateixes regles
      que el repas de contactes;
    - "DADES DEL TECNIC" i "no quadra amb els documents": els avisos de la base
      de contactes (contactes-db.json, que fa "Actualitzar base"). Sense la
      base, aquells dos tipus no es poden mirar i el peu ho diu. Els avisos
      'error' de la base NO s'hi afegeixen: son els mateixos que es miren aqui
      directament a l'Excel i sortirien dues vegades.
  El control periodic fa servir el criteri de l'eina "Controls periodics"
  (_ControlPeriodicClassify): una activitat que n'ha de passar i no te data
  prevista no surt mai a la seva llista de vençuts, ni a l'automatic.

  UNA COLUMNA QUE NO HI ES no fa saltar cap error: aquella comprovacio no es fa
  i el peu diu quina columna falta (si el GIA reanomena una capcalera, es veu
  de seguida en lloc de sortir 1.300 errors falsos).

  L'Excel del GIA NO ES TOCA MAI. L'exportacio es un .xlsx a part (sense Excel:
  _NormativaXlsxBytes, el mateix escriptor que l'index de la Normativa i les
  correccions de contactes).

  La tria dels tipus es desa en aquest PC (local\base-dades-activitats\
  excel-errors.json): s'hi apunten els DESMARCATS, aixi un tipus nou surt marcat.
#>

# ----------------------------------------------------------------------------
# ELS TIPUS D'ERROR (en l'ordre en que surten a la finestra i a l'Excel)
# ----------------------------------------------------------------------------
$Script:ExTipus = [ordered]@{
    id_duplicat      = 'ID Activitat repetit'
    rao_social       = 'Falta la raó social'
    nif_titular      = 'NIF del titular: falta o és incorrecte'
    rep_legal        = 'Representant legal: falta, incomplet o NIF incorrecte'
    tecnic           = 'Dades del tècnic o la gestoria com a titular o representant'
    adreca           = "Adreça de l'emplaçament incompleta"
    cadastre         = 'Referència cadastral: falta o és incorrecta'
    utm              = 'Coordenades UTM: falten o són fora de Cornellà'
    classificacio    = "Classificació: falta l'annex o l'apartat"
    activitat        = "Falta l'activitat principal"
    expedient        = "Falta el número d'expedient"
    data_llicencia   = 'Falta la data de llicència/comunicació'
    control_periodic = 'Control periòdic sense data prevista o sense periodicitat'
    correu           = "Correu electrònic: no n'hi ha cap o és incorrecte"
    telefon          = 'Telèfon o mòbil incorrecte'
    documents        = "Contactes: l'Excel no quadra amb els documents"
}

# La zona on han de caure les coordenades (UTM 31N, ETRS89). Cornella fa uns 3
# km d'un cap a l'altre; el quadre es generos a posta: el que vol enxampar son
# les coordenades d'un altre lloc, les X i Y girades, un zero de mes o de menys
# o els graus posats on va l'UTM, no una parcel.la del limit del terme.
$Script:ExUtmZona = @{ XMin = 420500; XMax = 427500; YMin = 4575500; YMax = 4582000 }

# ----------------------------------------------------------------------------
# LES COLUMNES (per la capcalera; 0 = no hi es i aquella comprovacio no es fa)
# ----------------------------------------------------------------------------
# PURA. $headers: la fila 1 (0-based), com _HeadersDeFila1.
function _ExColumnes($headers) {
    $h = @($headers)
    $c = Get-ColumnesContacte $h
    $col = @{
        ID       = (_ColumnaAmbTermes $h @('id activitat') $null)
        RAO      = (_ColumnaAmbTermes $h @('rao social') $null)
        VIA      = (_ColumnaAmbTermes $h @('emp', 'tipus via') $null)
        CARRER   = (_ColumnaAmbTermes $h @('emp', 'carrer') $null)
        NUMERO   = (_ColumnaAmbTermes $h @('emp', 'numero') $null)
        RC       = (_ColumnaAmbTermes $h @('ref', 'cadastral') $null)
        UTMX     = (_ColumnaAmbTermes $h @('utm x') $null)
        UTMY     = (_ColumnaAmbTermes $h @('utm y') $null)
        ANNEX    = (_ColumnaAmbTermes $h @('classificacio general annex') $null)
        APARTAT  = (_ColumnaAmbTermes $h @('classificacio general apartat') $null)
        ACT      = (_ColumnaAmbTermes $h @('activitat principal') $null)
        EXP      = (_ColumnaAmbTermes $h @('expedient') $null)
        LLIC     = (_ColumnaAmbTermes $h @('data', 'llicencia') $null)
        PERIOD   = (_ColumnaAmbTermes $h @('periodicitat') $null)
        PROPER   = (_ColumnaAmbTermes $h @('proper', 'previst') $null)
    }
    # Els index historics, com la resta del programa (Activitats.ps1, Ruta.ps1):
    # l'ID a la 1 i la rao social a la 10 si la capcalera no es reconeix.
    if ($col.ID -eq 0 -and $h.Count -ge 1) { $col.ID = 1 }
    if ($col.RAO -eq 0 -and $h.Count -ge 10 -and (_NormalitzaText $h[9]) -match 'rao') { $col.RAO = 10 }
    foreach ($k in @('NIF', 'TELEFON', 'MOBIL', 'EMAIL', 'REP_NOM', 'REP_NIF', 'REP_TELEFON', 'REP_MOBIL', 'REP_EMAIL')) { $col[$k] = [int]$c[$k] }
    # Get-ColumnesContacte dona 23 i 25 de respatller per al mobil i el correu:
    # aqui nomes valen si la capcalera hi es (si no, sortirien errors d'una
    # columna que no es la que es pensa).
    foreach ($k in @('MOBIL', 'EMAIL')) {
        if ($col[$k] -gt $h.Count -or ($col[$k] -gt 0 -and (_NormalitzaText $h[$col[$k] - 1]) -notmatch 'mobil|mail')) { $col[$k] = 0 }
    }
    return $col
}

# Les capcaleres (per al peu) de les columnes que no s'han trobat. PURA.
$Script:ExColumnesNoms = [ordered]@{
    RAO = 'Raó social'; NIF = 'Raó soc. NIF'; REP_NOM = 'Representant legal'; REP_NIF = 'Rep. Leg. NIF'
    VIA = 'Emp. Tipus via'; CARRER = 'Emp. Carrer'; NUMERO = 'Emp. Número'; RC = 'Ref. cadastral'
    UTMX = 'UTM X'; UTMY = 'UTM Y'; ANNEX = 'Classificació general annex'; APARTAT = 'Classificació general Apartat'
    ACT = 'Activitat principal'; EXP = 'Núm. expedient'; LLIC = 'Data llicència/comunicació'
    PERIOD = 'Periodicitat CP'; PROPER = 'Proper CP previst'; EMAIL = 'Raó soc. E-mail'; REP_EMAIL = 'Rep. Leg. E-mail'
}
function _ExColumnesQueFalten($col) {
    return @(@($Script:ExColumnesNoms.Keys) | Where-Object { [int]$col[$_] -le 0 } | ForEach-Object { $Script:ExColumnesNoms[$_] })
}

# ----------------------------------------------------------------------------
# LES COMPROVACIONS D'UNA FILA (PURES)
# ----------------------------------------------------------------------------
# Un numero de l'Excel (amb coma o punt decimal, o ja numeric), o $null. PURA.
function _ExNumero($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) { return [double]$v }
    $t = ([string]$v).Trim() -replace '\s', ''
    if ($t -eq '') { return $null }
    if ($t -match '^-?\d{1,3}(\.\d{3})+(,\d+)?$') { $t = $t.Replace('.', '') }
    $t = $t.Replace(',', '.')
    $d = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}

# La referencia cadastral: 20 caracters (14 de la parcel.la + 4 del carrec + 2
# de control) o, com a minim, els 14 de la parcel.la. Torna '' si es bona o el
# problema. PURA.
function _ExErradaCadastre([string]$rc) {
    $t = (([string]$rc).ToUpperInvariant()) -replace '[\s\-\.]', ''
    if ($t -eq '') { return 'falta' }
    if ($t -notmatch '^[0-9A-Z]+$') { return "porta car" + [char]0x00E0 + "cters que no hi poden ser" }
    if ($t.Length -ne 20 -and $t.Length -ne 14) { return ("t" + [char]0x00E9 + " " + $t.Length + " car" + [char]0x00E0 + "cters (n'ha de tenir 20, o 14 de la parcel" + [char]0x00B7 + "la)") }
    return ''
}

function _ExErr([string]$tipus, [string]$camp, [string]$valor, [string]$detall) {
    return @{ Tipus = $tipus; Camp = $camp; Valor = $valor; Detall = $detall }
}

# Tots els errors d'UNA activitat. $f: els valors de la fila (claus de
# _ExColumnes, ja com a text; UTMX/UTMY tal com venen); $col: les columnes (una
# a 0 vol dir que no hi es i no es mira). Torna una llista de @{ Tipus; Camp;
# Valor; Detall }. PURA.
function Get-ExcelErrorsFila($f, $col) {
    $out = New-Object System.Collections.ArrayList
    $v = { param($k) if ($null -eq $f[$k]) { '' } else { ([string]$f[$k]).Trim() } }
    $hi = { param($k) [int]$col[$k] -gt 0 }

    # Titular: la rao social i el NIF.
    if ((& $hi 'RAO') -and (& $v 'RAO') -eq '') { [void]$out.Add((_ExErr 'rao_social' ('Ra' + [char]0x00F3 + ' social') '' 'buida')) }
    $nifTit = _CtNifNet (& $v 'NIF')
    $tipusTit = _CtTipusNif $nifTit
    if (& $hi 'NIF') {
        if ($nifTit -eq '') { [void]$out.Add((_ExErr 'nif_titular' ('Ra' + [char]0x00F3 + ' soc. NIF') '' 'buit')) }
        else {
            $e = _CtErrada 'nif' (& $v 'NIF') '' ''
            if ($e -ne '') { [void]$out.Add((_ExErr 'nif_titular' ('Ra' + [char]0x00F3 + ' soc. NIF') (& $v 'NIF') $e)) }
        }
    }

    # Representant legal. Una entitat (NIF juridic) n'ha de tenir; d'una persona
    # fisica pot no tenir-ne, pero si n'hi ha, que sigui sencer.
    $repNom = & $v 'REP_NOM'; $repNifTxt = & $v 'REP_NIF'; $repNif = _CtNifNet $repNifTxt
    if (& $hi 'REP_NOM') {
        if ($repNom -eq '' -and $tipusTit -eq 'juridica') {
            [void]$out.Add((_ExErr 'rep_legal' 'Representant legal' '' ("buit, i el titular " + [char]0x00E9 + "s una entitat (" + $nifTit + ")")))
        } elseif ($repNom -eq '' -and $repNif -ne '') {
            [void]$out.Add((_ExErr 'rep_legal' 'Representant legal' '' ("buit, per" + [char]0x00F2 + " hi ha el NIF del representant (" + $repNifTxt + ")")))
        }
    }
    if (& $hi 'REP_NIF') {
        if ($repNom -ne '' -and $repNif -eq '' -and $tipusTit -eq 'juridica') {
            [void]$out.Add((_ExErr 'rep_legal' 'Rep. Leg. NIF' '' ("buit (representant: " + $repNom + ")")))
        } elseif ($repNif -ne '') {
            $e = ''
            if ($repNif -eq $nifTit) { $e = [char]0x00E9 + "s el NIF del titular, no el del representant" }
            elseif ((_CtTipusNif $repNif) -eq 'juridica') { $e = [char]0x00E9 + "s un NIF d'entitat, no el d'una persona" }
            else { $e = _CtErrada 'nif' $repNifTxt '' '' }
            if ($e -ne '') { [void]$out.Add((_ExErr 'rep_legal' 'Rep. Leg. NIF' $repNifTxt $e)) }
        }
    }

    # L'emplacament.
    $falten = @()
    foreach ($k in @('VIA', 'CARRER', 'NUMERO')) { if ((& $hi $k) -and (& $v $k) -eq '') { $falten += $Script:ExColumnesNoms[$k] } }
    if ($falten.Count -gt 0) {
        $adr = (@((& $v 'VIA'), (& $v 'CARRER'), (& $v 'NUMERO')) | Where-Object { $_ -ne '' }) -join ' '
        [void]$out.Add((_ExErr 'adreca' ($falten -join ', ') $adr ('buit: ' + ($falten -join ', '))))
    }
    if (& $hi 'RC') {
        $e = _ExErradaCadastre (& $v 'RC')
        if ($e -ne '') { [void]$out.Add((_ExErr 'cadastre' 'Ref. cadastral' (& $v 'RC') $e)) }
    }
    if ((& $hi 'UTMX') -and (& $hi 'UTMY')) {
        $x = _ExNumero $f['UTMX']; $y = _ExNumero $f['UTMY']
        $txt = (& $v 'UTMX') + ' / ' + (& $v 'UTMY')
        $z = $Script:ExUtmZona
        if ($null -eq $x -or $null -eq $y -or ($x -eq 0 -and $y -eq 0)) {
            $d = if ((& $v 'UTMX') -eq '' -and (& $v 'UTMY') -eq '') { 'falten' } else { "no s'entenen com a n" + [char]0x00FA + "meros" }
            [void]$out.Add((_ExErr 'utm' 'UTM X / UTM Y' $(if ($txt -eq ' / ') { '' } else { $txt }) $d))
        } elseif ($x -ge $z.YMin -and $x -le $z.YMax -and $y -ge $z.XMin -and $y -le $z.XMax) {
            [void]$out.Add((_ExErr 'utm' 'UTM X / UTM Y' $txt 'la X i la Y estan girades'))
        } elseif ($x -lt $z.XMin -or $x -gt $z.XMax -or $y -lt $z.YMin -or $y -gt $z.YMax) {
            [void]$out.Add((_ExErr 'utm' 'UTM X / UTM Y' $txt ("fora de Cornell" + [char]0x00E0 + " (s'esperen X entre " + $z.XMin + " i " + $z.XMax + ", Y entre " + $z.YMin + " i " + $z.YMax + ")")))
        }
    }

    # La classificacio i l'activitat.
    $annex = & $v 'ANNEX'; $apart = & $v 'APARTAT'
    if (& $hi 'ANNEX') {
        if ($annex -eq '') { [void]$out.Add((_ExErr 'classificacio' ('Classificaci' + [char]0x00F3 + ' general annex') '' 'buida')) }
        elseif ((& $hi 'APARTAT') -and $apart -eq '' -and $annex.ToUpperInvariant() -match '\bII\b|\bIII\b|\bI\b') {
            [void]$out.Add((_ExErr 'classificacio' ('Classificaci' + [char]0x00F3 + ' general Apartat') '' ("buit (annex " + $annex + ")")))
        }
    }
    if ((& $hi 'ACT') -and (& $v 'ACT') -eq '') { [void]$out.Add((_ExErr 'activitat' 'Activitat principal' '' 'buida')) }
    if ((& $hi 'EXP') -and (& $v 'EXP') -eq '') { [void]$out.Add((_ExErr 'expedient' ('N' + [char]0x00FA + 'm. expedient') '' 'buit')) }
    if ((& $hi 'LLIC') -and (& $v 'LLIC') -eq '') { [void]$out.Add((_ExErr 'data_llicencia' ('Data llic' + [char]0x00E8 + 'ncia/comunicaci' + [char]0x00F3) '' 'buida')) }

    # El control periodic (el criteri de l'eina Controls periodics).
    $cl = _ControlPeriodicClassify $annex $apart
    if ($cl.Qualifies) {
        $quin = (@($(if ($cl.IsII) { 'annex II' }), $(if ($cl.IsIII) { 'annex III' }), $(if ($cl.Is561) { 'apartat 561' })) | Where-Object { $_ }) -join ', '
        if ((& $hi 'PROPER') -and (& $v 'PROPER') -eq '') { [void]$out.Add((_ExErr 'control_periodic' 'Proper CP previst' '' ("buit, i en t" + [char]0x00E9 + " (" + $quin + ")"))) }
        if ((& $hi 'PERIOD') -and (& $v 'PERIOD') -eq '') { [void]$out.Add((_ExErr 'control_periodic' 'Periodicitat CP' '' ("buida, i en t" + [char]0x00E9 + " (" + $quin + ")"))) }
    }

    # Els correus i els telefons (les regles del repas de contactes).
    if ((& $hi 'EMAIL') -or (& $hi 'REP_EMAIL')) {
        if ((& $v 'EMAIL') -eq '' -and (& $v 'REP_EMAIL') -eq '') {
            [void]$out.Add((_ExErr 'correu' 'E-mail' '' ("no n'hi ha cap, ni del titular ni del representant: no se li pot escriure")))
        }
    }
    foreach ($p in @(@{ K = 'EMAIL'; D = 'email'; C = ('Ra' + [char]0x00F3 + ' soc. E-mail'); T = 'correu' },
                     @{ K = 'REP_EMAIL'; D = 'email'; C = 'Rep. Leg. E-mail'; T = 'correu' },
                     @{ K = 'TELEFON'; D = 'telefon'; C = ('Ra' + [char]0x00F3 + ' soc. Tel' + [char]0x00E8 + 'fon'); T = 'telefon' },
                     @{ K = 'MOBIL'; D = 'mobil'; C = ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil'); T = 'telefon' },
                     @{ K = 'REP_TELEFON'; D = 'telefon'; C = ('Rep. Leg. Tel' + [char]0x00E8 + 'fon'); T = 'telefon' },
                     @{ K = 'REP_MOBIL'; D = 'mobil'; C = ('Rep. Leg. M' + [char]0x00F2 + 'bil'); T = 'telefon' })) {
        if (-not (& $hi $p.K)) { continue }
        $val = & $v $p.K
        if ($val -eq '') { continue }
        # Diversos correus o telefons a la mateixa cel.la: un per un.
        foreach ($tros in @($val -split '\s*[;,/]\s*|\s+-\s+' | Where-Object { $_ -ne '' })) {
            $e = _CtErrada $p.D $tros '' ''
            if ($e -ne '') { [void]$out.Add((_ExErr $p.T $p.C $tros $e)) }
        }
    }
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# TOTS ELS ERRORS (PURA)
# ----------------------------------------------------------------------------
# $files: llista de @{ Id; Valors (hashtable per clau de columna) } en l'ordre
# de l'Excel; $col: _ExColumnes; $contactes: GIA -> activitat de la base de
# contactes (o $null si no n'hi ha). Torna una llista de [pscustomobject]
# @{ Id; Titular; Tipus; TipusNom; Camp; Valor; Detall }.
function Get-ExcelErrors($files, $col, $contactes) {
    $out = New-Object System.Collections.ArrayList
    $vistos = @{}
    foreach ($fl in @($files)) { if ($null -ne $fl) { $k = [string]$fl.Id; $vistos[$k] = 1 + [int]$vistos[$k] } }
    $repetitsDits = @{}
    $afegeix = {
        param($fl, $e)
        [void]$out.Add([pscustomobject]@{
            Id = [string]$fl.Id; Titular = [string]$fl.Valors['RAO']; Tipus = [string]$e.Tipus
            TipusNom = [string]$Script:ExTipus[[string]$e.Tipus]; Camp = [string]$e.Camp; Valor = [string]$e.Valor; Detall = [string]$e.Detall
        })
    }
    foreach ($fl in @($files)) {
        if ($null -eq $fl) { continue }
        $id = [string]$fl.Id
        if ([int]$vistos[$id] -gt 1 -and -not $repetitsDits.ContainsKey($id)) {
            $repetitsDits[$id] = $true
            & $afegeix $fl (_ExErr 'id_duplicat' 'ID Activitat' $id ("surt " + $vistos[$id] + " vegades a la fulla Est" + [char]0x00E8 + "s"))
        }
        foreach ($e in @(Get-ExcelErrorsFila $fl.Valors $col)) { & $afegeix $fl $e }
        if ($null -eq $contactes -or -not $contactes.ContainsKey($id)) { continue }
        foreach ($a in @(_CtV $contactes[$id] 'avisos')) {
            if ($null -eq $a) { continue }
            $t = [string](_CtV $a 'tipus')
            $tipus = switch ($t) { 'es_el_tecnic' { 'tecnic' } 'falta' { 'documents' } 'diferent' { 'documents' } 'canvi_titular' { 'documents' } default { '' } }
            if ($tipus -eq '') { continue }
            $det = [string](_CtV $a 'problema')
            $prop = [string](_CtV $a 'proposta')
            if ($prop -ne '') { $det += ' (proposta: ' + $prop + ')' }
            & $afegeix $fl (_ExErr $tipus ([string](_CtV $a 'camp')) ([string](_CtV $a 'valor_excel')) $det)
        }
    }
    return $out.ToArray()
}

# Quants n'hi ha de cada tipus (totes les claus, tambe les de 0). PURA.
function _ExComptes($errors) {
    $n = [ordered]@{}
    foreach ($k in @($Script:ExTipus.Keys)) { $n[$k] = 0 }
    foreach ($e in @($errors)) { if ($null -ne $e -and $n.Contains([string]$e.Tipus)) { $n[[string]$e.Tipus]++ } }
    return $n
}

# Les files de l'Excel exportat: NOMES els tipus triats, per ID i en l'ordre
# dels tipus. PURA.
$Script:ExExportCapcalera = @('ID Activitat', "Tipus d'error", 'Camp', "Valor a l'Excel", 'Detall', ('Ra' + [char]0x00F3 + ' social'))
$Script:ExExportAmples = @(12, 48, 28, 32, 70, 40)
function _ExFilesExport($errors, $tipusTriats) {
    $ordre = @{}; $i = 0
    foreach ($k in @($Script:ExTipus.Keys)) { $ordre[$k] = $i++ }
    $tria = @($errors | Where-Object { $null -ne $_ -and @($tipusTriats) -contains [string]$_.Tipus })
    $tria = @($tria | Sort-Object -Property @{ Expression = { _GiaNumeric $_.Id } }, @{ Expression = { [int]$ordre[[string]$_.Tipus] } })
    return @($tria | ForEach-Object { , @([string]$_.Id, [string]$_.TipusNom, [string]$_.Camp, [string]$_.Valor, [string]$_.Detall, [string]$_.Titular) })
}

function Export-ExcelErrors([string]$desti, $errors, $tipusTriats) {
    $bytes = _NormativaXlsxBytes $Script:ExExportCapcalera @(_ExFilesExport $errors $tipusTriats) $Script:ExExportAmples 'Errors'
    [System.IO.File]::WriteAllBytes($desti, $bytes)
    return $desti
}

# ----------------------------------------------------------------------------
# LA TRIA DESADA (els tipus DESMARCATS; un tipus nou surt marcat)
# ----------------------------------------------------------------------------
function _ExTriaPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'excel-errors.json')
}

# Els tipus marcats. Mai peta.
function Read-ExTriats([string]$path = '') {
    if (-not $path) { $path = _ExTriaPath }
    $no = @()
    if ($path) { try { $no = @(_PropInf (Read-JsonFile $path) 'desmarcats') | ForEach-Object { [string]$_ } } catch { $no = @() } }
    return @(@($Script:ExTipus.Keys) | Where-Object { $no -notcontains $_ })
}

function Save-ExTriats($triats, [string]$path = '') {
    if (-not $path) { $path = _ExTriaPath }
    if (-not $path) { return }
    try {
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile $path ([ordered]@{ desmarcats = @(@($Script:ExTipus.Keys) | Where-Object { @($triats) -notcontains $_ }) }) 3
    } catch { }
}

# ----------------------------------------------------------------------------
# LA LECTURA (Excel + base de contactes)
# ----------------------------------------------------------------------------
# @{ Ok; Files; Col; File; Error }.
function _ReadExcelErrors {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok = $false; Error = ("No s'ha trobat cap Excel d'activitats (" + (_ExcelBaseNomText) + ").") } }
    try {
        $lectura = Read-FullaEstesa $latest.File {
            param($x)
            $data = $x.Data
            if ($null -eq $data) { return @{ Ok = $true; Files = @(); Col = (_ExColumnes @()); File = $latest.File.Name } }
            $col = _ExColumnes $x.Headers
            $get = $x.Cel
            $files = New-Object System.Collections.ArrayList
            for ($ri = 2; $ri -le $x.Rows; $ri++) {
                $id = Get-IdDeCella $data[$ri, $col.ID]
                if ([string]::IsNullOrWhiteSpace($id)) { continue }
                $vals = @{}
                foreach ($k in @($col.Keys)) {
                    if ($k -eq 'ID' -or [int]$col[$k] -le 0) { continue }
                    # Les coordenades, tal com venen (numero o text): _ExNumero.
                    $vals[$k] = if ($k -eq 'UTMX' -or $k -eq 'UTMY') { $data[$ri, $col[$k]] } else { & $get $ri $col[$k] }
                }
                [void]$files.Add(@{ Id = $id; Valors = $vals })
            }
            return @{ Ok = $true; Files = $files.ToArray(); Col = $col; File = $latest.File.Name }
        }
        return $lectura
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# Les activitats de la base de contactes (GIA -> activitat), o $null si no n'hi ha.
function _ExContactes {
    try {
        $db = Read-ContactesDb
        if ($null -eq $db) { return $null }
        return (_CtMapaDe (_CtV $db 'activitats'))
    } catch { return $null }
}

# ----------------------------------------------------------------------------
# LA FINESTRA
# ----------------------------------------------------------------------------
function Show-ExcelErrorsFinestra($errors, [string]$peu, [string]$fileName) {
    $titol = 'Excel Errors'
    $comptes = _ExComptes $errors
    $claus = @($Script:ExTipus.Keys)
    $st = @{ Loading = $false; Canvi = $null }

    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(1180, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 560)
    $form.StartPosition = 'CenterScreen'

    # ---- A dalt: els tipus d'error, per marcar ---------------------------
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 246
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true; $lbl.Location = New-Object System.Drawing.Point(10, 8)
    $lbl.Text = "Tipus d'error (marca els que vols veure i exportar):"
    $top.Controls.Add($lbl)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(10, 28); $clb.Size = New-Object System.Drawing.Size(1140, 170)
    $clb.Anchor = 'Top, Left, Right'
    $clb.CheckOnClick = $true; $clb.IntegralHeight = $false
    $clb.MultiColumn = $true; $clb.ColumnWidth = 560
    $triats = @(Read-ExTriats)
    foreach ($k in $claus) {
        $i = $clb.Items.Add(([string]$Script:ExTipus[$k] + '  (' + [int]$comptes[$k] + ')'))
        $clb.SetItemChecked($i, ($triats -contains $k))
    }
    $top.Controls.Add($clb)
    $lnkTotes = New-Object System.Windows.Forms.LinkLabel
    $lnkTotes.Text = 'Marca-les totes'; $lnkTotes.AutoSize = $true
    $lnkTotes.Location = New-Object System.Drawing.Point(10, 206)
    $top.Controls.Add($lnkTotes)
    $lnkCap = New-Object System.Windows.Forms.LinkLabel
    $lnkCap.Text = 'Desmarca-les totes'; $lnkCap.AutoSize = $true
    $lnkCap.Location = New-Object System.Drawing.Point(120, 206)
    $top.Controls.Add($lnkCap)
    $lblCompte = New-Object System.Windows.Forms.Label
    $lblCompte.AutoSize = $true
    $lblCompte.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblCompte.Location = New-Object System.Drawing.Point(640, 206)
    $top.Controls.Add($lblCompte)

    # ---- La graella (dins d'un panell: _StyleListGrid fa Dock=Fill) --------
    $mig = New-Object System.Windows.Forms.Panel
    $mig.Dock = 'Fill'
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.ReadOnly = $true
    foreach ($c in @(@{ H = 'ID Activitat'; W = 80 }, @{ H = ('Ra' + [char]0x00F3 + ' social'); W = 210 }, @{ H = "Tipus d'error"; W = 300 },
                     @{ H = 'Camp'; W = 150 }, @{ H = "Valor a l'Excel"; W = 170 }, @{ H = 'Detall'; W = 380 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c.H; $col.Width = $c.W; $col.SortMode = 'NotSortable'
        [void]$grid.Columns.Add($col)
    }
    $mig.Controls.Add($grid)
    $txtCerca = _AddSearchBox $top 300 203 300 'Cerca:' $null

    # Els tipus marcats. Dins de l'ItemCheck el canvi encara no es efectiu
    # (salta ABANS): $st.Canvi porta el que s'esta canviant.
    $triatsAra = {
        $out = @()
        for ($i = 0; $i -lt $clb.Items.Count; $i++) {
            $m = $clb.GetItemChecked($i)
            if ($null -ne $st.Canvi -and [int]$st.Canvi.I -eq $i) { $m = [bool]$st.Canvi.V }
            if ($m) { $out += $claus[$i] }
        }
        return ,$out
    }.GetNewClosure()

    $fill = {
        if ($st.Loading) { return }
        $grid.Rows.Clear()
        $sel = & $triatsAra
        $n = ([string]$txtCerca.Text).Trim().ToLower()
        $nf = 0
        foreach ($fila in @(_ExFilesExport $errors $sel)) {
            if ($n -ne '' -and -not (_TextMatches ($fila -join ' ') $n)) { continue }
            [void]$grid.Rows.Add(@($fila[0], $fila[5], $fila[1], $fila[2], $fila[3], $fila[4]))
            $nf++
        }
        $nAct = @(@($errors) | Where-Object { $sel -contains [string]$_.Tipus } | ForEach-Object { [string]$_.Id } | Sort-Object -Unique).Count
        $lblCompte.Text = ([string]$nf + ' errors ' + [char]0x00B7 + ' ' + $nAct + ' activitats')
    }.GetNewClosure()

    $clb.add_ItemCheck({
        param($s, $e)
        $st.Canvi = @{ I = $e.Index; V = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) }
        try { & $fill } finally { $st.Canvi = $null }
    }.GetNewClosure())
    $txtCerca.add_TextChanged({ & $fill }.GetNewClosure())
    $lnkTotes.add_LinkClicked({ $st.Loading = $true; for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }; $st.Loading = $false; & $fill }.GetNewClosure())
    $lnkCap.add_LinkClicked({ $st.Loading = $true; for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }; $st.Loading = $false; & $fill }.GetNewClosure())

    # ---- El peu: el que no s'ha pogut mirar, i els botons ------------------
    $baix = New-Object System.Windows.Forms.Panel
    $baix.Dock = 'Bottom'; $baix.Height = 112
    $tbPeu = New-Object System.Windows.Forms.TextBox
    $tbPeu.Multiline = $true; $tbPeu.ReadOnly = $true; $tbPeu.ScrollBars = 'Vertical'
    $tbPeu.Location = New-Object System.Drawing.Point(10, 4); $tbPeu.Size = New-Object System.Drawing.Size(1140, 56)
    $tbPeu.Anchor = 'Top, Left, Right'
    $tbPeu.BackColor = [System.Drawing.Color]::White
    $tbPeu.Text = $peu
    $baix.Controls.Add($tbPeu)
    [void](_AddPeuBotons $form @(
        @{ Nom = 'Tancar'; Text = 'Tancar'; Clic = { $form.Close() }.GetNewClosure() }) @(
        @{ Nom = 'Export'; Text = 'Exportar a Excel'; Estil = 'primari'; Clic = {
            $sel = & $triatsAra
            if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Marca almenys un tipus d'error.", $titol, 'OK', 'Information') | Out-Null; return }
            Save-ExTriats $sel
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
            $dlg.FileName = 'Excel errors ' + (Get-Date).ToString('yyyy-MM-dd') + '.xlsx'
            try { $dlg.InitialDirectory = _ResolveOutputDir } catch { }
            if ($dlg.ShowDialog() -ne 'OK') { return }
            try {
                [void](Export-ExcelErrors $dlg.FileName $errors $sel)
                [void](Open-PerUsuari $dlg.FileName)
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut exportar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() }) 64 $baix)

    # L'ordre compta (Dock): el que s'afegeix l'ultim es col·loca el primer.
    $form.Controls.Add($mig)
    $form.Controls.Add($top)
    $form.Controls.Add($baix)
    [void](_AddBrandHeader $form $titol ("Errors de la base de dades d'activitats " + [char]0x00B7 + ' ' + $fileName + ' ' + [char]0x00B7 + " l'Excel no es toca") 56)
    $form.add_FormClosing({ Save-ExTriats (& $triatsAra) }.GetNewClosure())
    _AjustaFinestraAPantalla $form
    & $fill
    [void]$form.ShowDialog()
    $form.Dispose()
}

# El peu: les columnes que no s'han trobat i si s'han pogut mirar els documents.
# PURA.
function _ExPeu($col, [bool]$ambContactes, [int]$nAct) {
    $l = New-Object System.Collections.ArrayList
    [void]$l.Add(([string]$nAct + " activitats a la fulla Est" + [char]0x00E8 + "s."))
    $falten = @(_ExColumnesQueFalten $col)
    if ($falten.Count -gt 0) { [void]$l.Add(("Columnes que no s'han trobat (aquelles comprovacions no s'han fet): " + ($falten -join ', ') + '.')) }
    if (-not $ambContactes) { [void]$l.Add(("Sense la base de contactes (la fa 'Actualitzar base') no s'han pogut mirar: " + $Script:ExTipus['tecnic'] + '; ' + $Script:ExTipus['documents'] + '.')) }
    return ($l -join "`r`n")
}

# ----------------------------------------------------------------------------
# Punt d'entrada (menu, fila GIA)
# ----------------------------------------------------------------------------
function Invoke-ExcelErrors {
    $titol = 'Excel Errors'
    $sp = _NewForm
    $sp.Text = $titol
    $sp.Size = New-Object System.Drawing.Size(420, 130)
    $sp.FormBorderStyle = 'FixedDialog'
    $sp.MaximizeBox = $false; $sp.MinimizeBox = $false
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(20, 18); $l.Size = New-Object System.Drawing.Size(370, 40)
    $l.Text = "Llegint la base d'activitats (Excel)..."
    $sp.Controls.Add($l)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 62); $bar.Size = New-Object System.Drawing.Size(370, 20)
    $bar.Style = 'Marquee'
    $sp.Controls.Add($bar)
    $sp.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $res = $null; $ct = $null
    try { $res = _ReadExcelErrors; $ct = _ExContactes } finally { try { $sp.Close() } catch { } }
    if ($null -eq $res -or -not [bool]$res.Ok) {
        $msg = if ($res) { [string]$res.Error } else { 'Error desconegut.' }
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'activitats:`n$msg", $titol, 'OK', 'Error') | Out-Null
        return
    }
    $errors = @(Get-ExcelErrors @($res.Files) $res.Col $ct)
    Show-ExcelErrorsFinestra $errors (_ExPeu $res.Col ($null -ne $ct) @($res.Files).Count) ([string]$res.File)
}
