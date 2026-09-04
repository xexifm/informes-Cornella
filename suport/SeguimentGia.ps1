#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Seguiment" (fila GIA): genera l'Excel de seguiment d'activitats a partir
  de la base de dades d'activitats, i el pot exportar a Excel o a PDF.

.DESCRIPTION
  Substitueix un Excel amb formules ('0_PLANTILLA.xlsx') que l'usuari es va
  muntar fa temps. Aquell fitxer treia cinc llistats de la base de dades amb
  VLOOKUP/MATCH, i per fer-ho necessitava 15 COLUMNES OCULTES d'ajuda a la fulla
  Estes, rangs fixos fins a la fila 15000 (per aixo la plantilla declara
  A1:EV15000 tot i que nomes ~1.312 files tenen ID Activitat) i 100 files
  pre-omplertes per pestanya. Aqui la feina de les formules la fa el codi, i el
  resultat es el mateix pero sense cap columna oculta ni res a recalcular.

  Pestanyes que genera:
    Estes              copia de la base de dades d'activitats
    PRECINTES          activitats amb el Camp Info 'PRECINTE ACTIVITAT?'
    DENUNCIES          ... amb 'DENUNCIA?'
    REQUERIT DECRET    ... amb 'REQUERIT PER DECRET?'
    SONOMETRIA         ... amb 'SONOMETRIA?'
    ANNEX II           'Classificacio general annex' = II amb 'Descripcio lliure'

  EL CRITERI ES NOMES QUE L'ACTIVITAT TINGUI AQUELL CAMP INFO, digui el que
  digui el valor. NO es demana que comenci per "SI". Aixo es deliberat: es el
  que fa la plantilla de l'usuari, i s'ha comprovat contra les seves dades (a
  REQUERIT DECRET hi ha dues activitats amb valor 'PROCEDIMENT ESMENA' i
  'CONTROL PERIODIC VTO. 27-04-2026...'). Compte, doncs: l'eina "Comprovar
  Excel" (Informes.ps1) fa servir un criteri DIFERENT, alli si que cal el "SI".

  Les columnes es resolen PEL NOM DE LA CAPCALERA, com feia el MATCH de la
  plantilla: aixi no es trenca si el GIA afegeix o mou columnes.

  Les funcions de dades son PURES (es proven en headless, sense Excel ni
  Windows); nomes la construccio del llibre i l'exportacio fan servir Excel COM.
#>

# ----------------------------------------------------------------------------
# DEFINICIO DE LES PESTANYES (un sol lloc)
# ----------------------------------------------------------------------------
# Columnes comunes de l'activitat, en l'ordre de la plantilla. El text es el NOM
# DE LA CAPCALERA a la base de dades (amb els espais finals que hi porta de
# debo: 'Num. expedient ' i 'Nom comercial activitat ' n'acaben amb un).
$Script:SgColsActivitat = @(
    'ID Activitat'
    ('N' + [char]0x00FA + 'm. expedient ')
    ('Ra' + [char]0x00F3 + ' social')
    ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil')
    'Representant legal'
    ('Rep. Leg. M' + [char]0x00F2 + 'bil')
    'Nom comercial activitat '
    'Emp. Tipus via'
    'Emp. Carrer'
    ('Emp. N' + [char]0x00FA + 'mero')
)

# Amplades de columna EXACTES de la plantilla (unitats de l'Excel), tretes del
# seu <cols> descomprimit. N'hi ha d'haver UNA PER COLUMNA de sortida:
#   1a  = la columna 'N' (numero de fila)
#   ...  les 10 de l'activitat
#   2 ultimes = Camp Info Nom / Valor (la del Valor, ampla, amb text ajustat)
#
# El 0 vol dir "no la toquem": a la plantilla, 'Rep. Leg. Mobil' NO porta amplada
# propia i es queda amb la de defecte de la fulla (baseColWidth=10). Ha de ser-hi
# com a forat, perque si no totes les amplades seguents ballen una posicio (hi va
# passar: la columna ampla del Valor es quedava sense amplada i les altres
# s'aplicaven a la columna del costat).
$Script:SgAmpladesCampInfo = @(3.57, 10.14, 12.86, 31.71, 10.86, 31.57, 0, 23.29, 13.14, 25.86, 12.71, 20.86, 58)
$Script:SgAmpladesAnnex    = @(3.57, 10.14, 12.86, 31.71, 10.86, 31.57, 0, 23.29, 7.14, 20, 7.57, 7.14, 9.14, 10.14, 10.43, 39.57, 50.86)

# Les 5 pestanyes de llistat. Per cada una:
#   Nom        nom de la pestanya
#   Titol      text de la fila 1 (se li afegeix la data)
#   Tipus      'campinfo' (busca un Camp Info) o 'annex'
#   Camp       nom del Camp Info a buscar (nomes si Tipus='campinfo')
#   Cols       capceleres de les columnes de dades (sense la 'N')
#   Amplades   amplades, incloent-hi la de la 'N'
#   DataCols   capceleres que s'han de mostrar com a data dd/MM/aaaa
function _SgFullesDef {
    $ci = @($Script:SgColsActivitat) + @('Camp Info 1 - Nom', 'Camp Info 1 - Valor')
    $an = @($Script:SgColsActivitat) + @(
        ('Classificaci' + [char]0x00F3 + ' general annex')
        ('Classificaci' + [char]0x00F3 + ' general Apartat')
        ('Data control inicial/verificaci' + [char]0x00F3)
        ('Data control peri' + [char]0x00F2 + 'dic')
        'Activitat principal'
        ('Descripci' + [char]0x00F3 + ' lliure')
    )
    return @(
        [pscustomobject]@{ Nom='PRECINTES';       Titol='ACTIVITAT PRECINTADA?';  Tipus='campinfo'; Camp='PRECINTE ACTIVITAT?';   Cols=$ci; Amplades=$Script:SgAmpladesCampInfo; DataCols=@() }
        [pscustomobject]@{ Nom=('DEN' + [char]0x00DA + 'NCIES'); Titol=('DEN' + [char]0x00DA + 'NCIA?'); Tipus='campinfo'; Camp=('DEN' + [char]0x00DA + 'NCIA?'); Cols=$ci; Amplades=$Script:SgAmpladesCampInfo; DataCols=@() }
        [pscustomobject]@{ Nom='REQUERIT DECRET'; Titol='REQUERIT PER DECRET?';   Tipus='campinfo'; Camp='REQUERIT PER DECRET?';  Cols=$ci; Amplades=$Script:SgAmpladesCampInfo; DataCols=@() }
        [pscustomobject]@{ Nom='SONOMETRIA';      Titol='SONOMETRIA?';            Tipus='campinfo'; Camp='SONOMETRIA?';           Cols=$ci; Amplades=$Script:SgAmpladesCampInfo; DataCols=@() }
        [pscustomobject]@{ Nom='ANNEX II';        Titol='ANNEXOS II';             Tipus='annex';    Camp='';                      Cols=$an; Amplades=$Script:SgAmpladesAnnex;
                           DataCols=@(('Data control inicial/verificaci' + [char]0x00F3), ('Data control peri' + [char]0x00F2 + 'dic')) }
    )
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Nom de la pestanya amb la copia de la base de dades. Es alhora el nom real de
# la fulla i la clau amb que es tria a la finestra.
function _SgNomEstes { return ('Est' + [char]0x00E8 + 's') }

# TOTES les opcions exportables, en l'ordre en que surten al llibre. Es l'unica
# llista que hi ha: la finestra la fa servir per pintar les caselles i la
# construccio per saber que ha de muntar, o sigui que no es poden desincronitzar.
function _SgOpcionsExport {
    $out = New-Object System.Collections.ArrayList
    [void]$out.Add((_SgNomEstes))
    foreach ($d in @(_SgFullesDef)) { [void]$out.Add([string]$d.Nom) }
    return $out.ToArray()
}

# Cert si la seleccio inclou la copia de la base de dades.
function _SgSeleccioTeEstes($seleccio) {
    $n = _NormalitzaText (_SgNomEstes)
    foreach ($s in @($seleccio)) {
        if ((_NormalitzaText $s) -eq $n) { return $true }
    }
    return $false
}

# Les definicions de llistat que s'han triat, EN L'ORDRE de _SgFullesDef (no en
# l'ordre en que l'usuari les hagi marcat: el llibre sempre surt igual).
function _SgFullesTriades($seleccio) {
    $tri = @{}
    foreach ($s in @($seleccio)) {
        $k = _NormalitzaText $s
        if (-not [string]::IsNullOrWhiteSpace($k)) { $tri[$k] = $true }
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @(_SgFullesDef)) {
        if ($tri.ContainsKey((_NormalitzaText $d.Nom))) { [void]$out.Add($d) }
    }
    return $out.ToArray()
}

# Index (1-based) de la columna amb aquesta capcalera. La PRIMERA que hi
# coincideixi, com feia el MATCH de la plantilla: a la base de dades hi ha
# capceleres repetides ('Classificacio general annex' hi surt dues vegades) i la
# plantilla mostra sempre la primera. 0 si no hi es. $headers es 0-based.
function _SgColIndex($headers, [string]$nom) {
    $t = _NormalitzaText $nom
    $h = @($headers)
    for ($i = 0; $i -lt $h.Count; $i++) {
        if ((_NormalitzaText $h[$i]) -eq $t) { return ($i + 1) }
    }
    return 0
}

# Deixa les parelles de Camp Info en una llista PLANA de @{NomCol;ValorCol}.
#
# Per que cal: _FindCampInfoPairs (Informes.ps1) acaba amb 'return ,@($pairs)'.
# La coma hi es a posta (protegeix el cas d'UNA sola parella: sense ella el
# consumidor rebria el hashtable pelat i .Count li donaria el nombre de CLAUS),
# pero vol dir que s'ha de consumir SENSE @(): un @() al voltant hi torna a
# posar la capa i deixa un array d'UN element que conte l'array de parelles.
# Quan passava aixo, '$p.NomCol' dins del bucle feia enumeracio de membres i
# retornava un Object[], i '[int]$p.NomCol' petava amb
#   "No se puede convertir el valor System.Object[] ... al tipo System.Int32".
# Aixo es aplanar-ho una vegada al punt d'entrada, de manera que la logica
# funcioni amb les dues formes i el crash no es pugui repetir. Funcio PURA.
function _SgAplanaPairs($pairs) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $pairs) { return $out.ToArray() }
    foreach ($p in @($pairs)) {
        if ($null -eq $p) { continue }
        if ($p -is [System.Collections.IDictionary]) { [void]$out.Add($p); continue }
        # No es un hashtable: es una col·leccio (la capa de mes). L'aplanem.
        if ($p -is [System.Collections.IEnumerable] -and $p -isnot [string]) {
            foreach ($q in $p) {
                if ($q -is [System.Collections.IDictionary]) { [void]$out.Add($q) }
            }
            continue
        }
        [void]$out.Add($p)
    }
    return $out.ToArray()
}

# Quina parella 'Camp Info N - Nom/Valor' d'aquesta fila val per al criteri.
# $pairs son les parelles de _FindCampInfoPairs (@{NomCol;ValorCol}, 1-based) i
# $fila es la fila de valors (0-based). Retorna @{NomCol;ValorCol} o $null.
# Es l'equivalent del MATCH(criteri, P:FZ, 0) de la plantilla.
function _SgCampInfoCoincident($pairs, $fila, [string]$camp) {
    if ($null -eq $pairs -or $null -eq $fila) { return $null }
    $t = _NormalitzaText $camp
    foreach ($p in @($pairs)) {
        $i = [int]$p.NomCol - 1
        if ($i -lt 0 -or $i -ge @($fila).Count) { continue }
        if ((_NormalitzaText $fila[$i]) -eq $t) { return $p }
    }
    return $null
}

# Criteri de la pestanya ANNEX II: classificacio 'II' i descripcio lliure amb
# contingut. A la plantilla la classificacio es llegeix de la SEGONA columna
# 'Classificacio general annex' (DG) mentre que la que ES MOSTRA es la primera
# (CZ). A les dades reals les dues son identiques a totes les activitats, o
# sigui que no canvia res, pero es deixa dit.
function _SgAnnexCoincideix([string]$classificacio, [string]$descripcio) {
    if ([string]::IsNullOrWhiteSpace($descripcio)) { return $false }
    return ((_NormalitzaText $classificacio) -eq 'ii')
}

# Titol de la fila 1: "ACTIVITAT PRECINTADA? 30/07/2026".
function _SgTitolFulla($def, [datetime]$data) {
    return ([string]$def.Titol + ' ' + $data.ToString('dd/MM/yyyy'))
}

# Nom del fitxer de sortida (sense carpeta).
function _SgNomFitxer([datetime]$data, [string]$ext = 'xlsx') {
    return ($data.ToString('yyyy-MM-dd') + ' Seguiment GIA.' + $ext)
}

# EL COS DE L'EINA, i es PUR: a partir de la taula ja llegida de l'Excel, en
# treu les files d'una pestanya.
#   $headers  capceleres (0-based)
#   $files    llista de files de valors (0-based cadascuna)
#   $pairs    parelles Camp Info (_FindCampInfoPairs)
# Retorna una llista d'arrays: [numero, ...columnes de $def.Cols...].
# Les files surten en l'ORDRE de la base de dades, com a la plantilla.
function _SgFilesPerFulla($def, $headers, $files, $pairs) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $files) { return $out.ToArray() }
    # Punt d'entrada de les parelles: les aplanem aqui i prou (vegeu
    # _SgAplanaPairs). Aixi tant se val com les hagi consumit qui ens crida.
    $pairs = _SgAplanaPairs $pairs

    # Index de cada columna de sortida, resolt UNA vegada (no per fila).
    $idx = @()
    foreach ($c in @($def.Cols)) { $idx += (_SgColIndex $headers $c) }
    $iGia = _SgColIndex $headers 'ID Activitat'
    $esData = @{}
    foreach ($d in @($def.DataCols)) { $esData[(_NormalitzaText $d)] = $true }

    # ANNEX II: la classificacio del CRITERI es la SEGONA columna amb aquest nom.
    $iClassCrit = 0
    $iDescr = 0
    if ($def.Tipus -eq 'annex') {
        $nomClass = _NormalitzaText ('Classificaci' + [char]0x00F3 + ' general annex')
        $vistes = 0
        $h = @($headers)
        for ($i = 0; $i -lt $h.Count; $i++) {
            if ((_NormalitzaText $h[$i]) -eq $nomClass) {
                $vistes++
                if ($vistes -eq 2) { $iClassCrit = $i + 1; break }
            }
        }
        if ($iClassCrit -eq 0) { $iClassCrit = _SgColIndex $headers ('Classificaci' + [char]0x00F3 + ' general annex') }
        $iDescr = _SgColIndex $headers ('Descripci' + [char]0x00F3 + ' lliure')
    }

    $cel = {
        param($fila, $i)
        if ($i -le 0 -or $i -gt @($fila).Count) { return '' }
        $v = $fila[$i - 1]
        if ($null -eq $v) { return '' }
        return ([string]$v).Trim()
    }

    $n = 0
    foreach ($fila in @($files)) {
        if ((& $cel $fila $iGia) -eq '') { continue }   # fila sense activitat

        $ciPair = $null
        if ($def.Tipus -eq 'annex') {
            if (-not (_SgAnnexCoincideix (& $cel $fila $iClassCrit) (& $cel $fila $iDescr))) { continue }
        } else {
            $ciPair = _SgCampInfoCoincident $pairs $fila $def.Camp
            if ($null -eq $ciPair) { continue }
        }

        $n++
        $rec = New-Object System.Collections.ArrayList
        [void]$rec.Add($n)
        for ($k = 0; $k -lt @($def.Cols).Count; $k++) {
            $nomCol = [string]@($def.Cols)[$k]
            # Les dues ultimes columnes d'una pestanya de Camp Info NO son
            # columnes fixes: son la parella que ha coincidit en AQUESTA fila.
            if ($null -ne $ciPair -and $nomCol -eq 'Camp Info 1 - Nom')   { [void]$rec.Add((& $cel $fila $ciPair.NomCol));   continue }
            if ($null -ne $ciPair -and $nomCol -eq 'Camp Info 1 - Valor') { [void]$rec.Add((& $cel $fila $ciPair.ValorCol)); continue }
            $v = & $cel $fila $idx[$k]
            if ($esData.ContainsKey((_NormalitzaText $nomCol))) { $v = _FormatDateOnly $v }
            [void]$rec.Add($v)
        }
        [void]$out.Add($rec.ToArray())
    }
    # Array PLA (no ,$ArrayList): aixi @() l'enumera fila a fila.
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# CONSTRUCCIO DEL LLIBRE (Excel COM)
# ----------------------------------------------------------------------------
# Constants d'Excel (per no dependre de les enumeracions amb nom).
$Script:SgXl = @{
    A3            = 8      # xlPaperA3
    Landscape     = 2      # xlLandscape
    Thin          = 2      # xlThin
    ContinuousLn  = 1      # xlContinuous
    OpenXMLBook   = 51     # xlOpenXMLWorkbook (.xlsx)
    TypePDF       = 0      # xlTypePDF
    EdgeLeft      = 7; EdgeTop = 8; EdgeBottom = 9; EdgeRight = 10
    InsideVert    = 11; InsideHoriz = 12
}

# COLORS, tal com son a la plantilla de l'usuari (trets del styles.xml).
# L'Excel els vol com a R + G*256 + B*65536 (BGR), no com un #RRGGBB.
#   Capcalera: lletra BLANCA sobre BLAU MARI. Al fitxer son colors indexats de
#   la paleta antiga: 'indexed 9' = FFFFFF i 'indexed 18' = 000080. (Aqui hi
#   havia un #D9E1F2 inventat per mi, que no s'assemblava a l'original.)
#   Files de dades: ratllat de zebra amb el 'theme 2' del tema del fitxer, que
#   val E8E8E8; la 1a fila de dades va ombrejada i despres una si una no.
$Script:SgColors = @{
    CapcaleraFons  = 8388608    # 000080 blau mari
    CapcaleraText  = 16777215   # FFFFFF blanc
    FilaAlterna    = 15263976   # E8E8E8 gris clar
}

# Text d'un error: el missatge, EN QUIN PAS estavem i la LINIA exacta del codi.
#
# Aquesta eina nomes es pot provar amb l'Excel de debo (fora de Windows no hi ha
# COM), o sigui que quan peta ha de dir tot el que sap. Amb el missatge pelat
# ("Unable to cast object of type X to type Y") toca endevinar quina de les
# vint crides a l'Excel ha estat, i cada intent costa una volta sencera.
function _SgTextError($err, [string]$pas) {
    $parts = New-Object System.Collections.ArrayList
    $msg = ''
    try { $msg = [string]$err.Exception.Message } catch { }
    if ([string]::IsNullOrWhiteSpace($msg)) { $msg = [string]$err }
    [void]$parts.Add($msg)
    if (-not [string]::IsNullOrWhiteSpace($pas)) { [void]$parts.Add(('Estava ' + $pas + '.')) }
    try {
        $inv = $err.InvocationInfo
        if ($null -ne $inv) {
            [void]$parts.Add(('SeguimentGia.ps1, linia ' + [string]$inv.ScriptLineNumber + ':' + "`n" + ([string]$inv.Line).Trim()))
        }
    } catch { }
    return ($parts -join "`n`n")
}

# Valor tal com s'ha d'escriure a una cel·la: SEMPRE una cadena.
#
# A l'Excel de l'usuari, escriure a Value2 qualsevol cosa que NO sigui una cadena
# peta. Ho hem vist dues vegades seguides, i la segona ho va deixar clar:
#   - $rang.Value2 = $matriu   -> "Unable to cast object of type
#     'System.Object[,]' to type 'System.String'"
#   - $cel.Value2  = 1         -> "Unable to cast object of type 'System.Int32'
#     to type 'System.String'"
# ...mentre que el titol i les capceleres (cadenes) sempre s'han escrit be. O
# sigui que l'adaptador COM d'aquest PowerShell resol el 'put' de Value2 com si
# demanes una cadena, i tot el que no ho es, peta.
#
# No es cap perdua: de tota la taula, l'UNIC valor que no era text ja era la
# columna 'N' (el numero de fila); tota la resta passa per $cel, que ja retorna
# cadenes. I l'Excel interpreta el text en assignar-lo igual que si l'escrivissis
# a ma, o sigui que "1" segueix sent el numero 1 i les dates, dates.
# Funcio PURA.
function _SgValorCella($v) {
    if ($null -eq $v) { return '' }
    return [string]$v
}

# Escriu una matriu [files x columnes] a la fulla, a partir de $filaInici i de la
# columna 1. Retorna com ho ha aconseguit ('bloc', 'bloc-invoke' o 'cel·la').
#
# Es proven tres maneres perque la primera es la rapida i les altres dues son el
# respatller (vegeu _SgValorCella per al perque). Si totes tres fallen, es llanca
# un error que diu que ha passat a cada una: aquesta eina nomes es pot provar amb
# l'Excel de debo, i un error mut costa una volta sencera amb l'usuari.
#
# El cami cel·la a cel·la aqui es pot permetre: aquests llistats son de desenes
# de files (26/24/48/8/51), no de milers. El "per matriu es de minuts a segons"
# valia per a la LECTURA de la base sencera (1.312 x 152), no per a aixo.
function _SgEscriuMatriu($sh, [int]$filaInici, $matriu) {
    $nf = $matriu.GetLength(0)
    $nc = $matriu.GetLength(1)
    $errors = New-Object System.Collections.ArrayList

    try {
        $dest = $sh.Range($sh.Cells.Item($filaInici, 1), $sh.Cells.Item($filaInici + $nf - 1, $nc))
        $dest.Value2 = $matriu
        return 'bloc'
    } catch { [void]$errors.Add('bloc: ' + $_.Exception.Message) }

    # El @(,$matriu) no es decoratiu: la llista d'arguments ha de portar la
    # matriu com un UNIC element; sense la coma, l'array es desenrotllaria.
    try {
        $dest = $sh.Range($sh.Cells.Item($filaInici, 1), $sh.Cells.Item($filaInici + $nf - 1, $nc))
        [void]$dest.GetType().InvokeMember('Value2',
            [System.Reflection.BindingFlags]::SetProperty, $null, $dest, @(, $matriu))
        return 'bloc-invoke'
    } catch { [void]$errors.Add('bloc-invoke: ' + $_.Exception.Message) }

    try {
        for ($i = 0; $i -lt $nf; $i++) {
            for ($j = 0; $j -lt $nc; $j++) {
                $sh.Cells.Item($filaInici + $i, $j + 1).Value2 = (_SgValorCella $matriu[$i, $j])
            }
        }
        return ('cel' + [char]0x00B7 + 'la')
    } catch { [void]$errors.Add('cel-la: ' + $_.Exception.Message) }

    throw ("No s'han pogut escriure les dades a la fulla de cap de les tres maneres:" + "`n" + ($errors -join "`n"))
}

# Deixa una pestanya de llistat amb el format i la configuracio d'impressio de
# la plantilla. Tot el que es veu (colors, amplades, marges, peu de pagina) surt
# d'aqui: si algun dia canvia, es canvia en un sol lloc.
function _SgFormatarFulla($sh, $def, [int]$nFiles, $excel) {
    $nCols = @($def.Cols).Count + 1
    $ampl  = @($def.Amplades)

    # --- Fila 1: titol en Arial 14 negreta VERMELLA ---
    $sh.Rows.Item(1).RowHeight = 18
    $t = $sh.Cells.Item(1, 1)
    $t.Font.Name = 'Arial'; $t.Font.Size = 14; $t.Font.Bold = $true
    $t.Font.Color = 255                                  # vermell (BGR)
    $t.VerticalAlignment = -4108                         # xlCenter

    # --- Fila 2: capceleres, amb els colors de la plantilla ---
    #     Lletra BLANCA sobre BLAU MARI, i el text AJUSTAT perque no es talli (la
    #     fila creix tota sola mes avall, quan les columnes ja tenen l'amplada
    #     definitiva). A la plantilla, la columna 'N' NO porta fons.
    $hdr = $sh.Range($sh.Cells.Item(2, 1), $sh.Cells.Item(2, $nCols))
    $hdr.Font.Name = 'Arial'; $hdr.Font.Size = 10; $hdr.Font.Bold = $true
    $hdr.HorizontalAlignment = -4108                      # xlCenter
    $hdr.VerticalAlignment = -4108
    $hdr.WrapText = $true
    if ($nCols -ge 2) {
        $hdrFons = $sh.Range($sh.Cells.Item(2, 2), $sh.Cells.Item(2, $nCols))
        $hdrFons.Interior.Color = $Script:SgColors.CapcaleraFons
        $hdrFons.Font.Color = $Script:SgColors.CapcaleraText
    }

    # --- Dades: vores fines i text ajustat a les columnes llargues ---
    $ultima = if ($nFiles -gt 0) { 2 + $nFiles } else { 2 }
    $tot = $sh.Range($sh.Cells.Item(2, 1), $sh.Cells.Item($ultima, $nCols))
    foreach ($e in @($Script:SgXl.EdgeLeft, $Script:SgXl.EdgeTop, $Script:SgXl.EdgeBottom,
                     $Script:SgXl.EdgeRight, $Script:SgXl.InsideVert, $Script:SgXl.InsideHoriz)) {
        try {
            $b = $tot.Borders.Item($e)
            $b.LineStyle = $Script:SgXl.ContinuousLn
            $b.Weight = $Script:SgXl.Thin
        } catch { }
    }

    # --- Amplades EXACTES de la plantilla (0 = deixar la de defecte) ---
    for ($c = 1; $c -le $nCols; $c++) {
        if ($c -gt $ampl.Count) { continue }
        $w = [double]$ampl[$c - 1]
        if ($w -le 0) { continue }
        $sh.Columns.Item($c).ColumnWidth = $w
    }

    # --- Text AJUSTAT a les columnes de text llarg (el que ha demanat l'usuari):
    #     'Camp Info 1 - Valor' i, a ANNEX II, 'Descripcio lliure'. Sempre son
    #     l'ultima columna de la pestanya.
    if ($nFiles -gt 0) {
        $colWrap = $nCols
        $rng = $sh.Range($sh.Cells.Item(3, $colWrap), $sh.Cells.Item($ultima, $colWrap))
        $rng.WrapText = $true
        $rng.VerticalAlignment = -4160                    # xlTop
        # Autoajust d'alcada: amb wrap, si no, el text queda tallat.
        try { $sh.Rows("3:$ultima").AutoFit() | Out-Null } catch { }
    }

    # --- Alcada de la CAPCALERA: ha d'anar DESPRES de les amplades, si no
    #     l'AutoFit calcularia l'alcada amb l'amplada que no toca i el text es
    #     tallaria igualment (que es el que passava).
    try { $sh.Rows.Item(2).AutoFit() | Out-Null } catch { }

    # --- Ratllat de zebra de les dades, com a la plantilla: la 1a fila de dades
    #     ombrejada i despres una si una no.
    if ($nFiles -gt 0) {
        for ($r = 3; $r -le $ultima; $r += 2) {
            try { $sh.Range($sh.Cells.Item($r, 1), $sh.Cells.Item($r, $nCols)).Interior.Color = $Script:SgColors.FilaAlterna } catch { }
        }
    }

    # --- Panells fixats a A3 (les dues primeres files sempre visibles) ---
    try {
        $sh.Activate()
        $sh.Range('A3').Select() | Out-Null
        $excel.ActiveWindow.FreezePanes = $true
    } catch { }

    # --- Autofiltre a la fila 2, sobre les columnes de dades (no la 'N') ---
    if ($nFiles -gt 0) {
        try { $sh.Range($sh.Cells.Item(2, 2), $sh.Cells.Item($ultima, $nCols)).AutoFilter() | Out-Null } catch { }
    }

    # --- IMPRESSIO: exactament el que te la plantilla ---
    #     horitzontal, A3, ajustat a 1 pagina d'AMPLE (i tantes d'alt com
    #     calgui), marges de 0,5 cm, files 1:2 repetides i peu amb el nom de la
    #     pestanya a l'esquerra i "Pagina N" a la dreta.
    try {
        $ps = $sh.PageSetup
        $ps.Orientation      = $Script:SgXl.Landscape
        $ps.PaperSize        = $Script:SgXl.A3
        $ps.Zoom             = $false
        $ps.FitToPagesWide   = 1
        $ps.FitToPagesTall   = $false
        $cm = { param($v) $excel.CentimetersToPoints($v) }
        $ps.LeftMargin   = & $cm 0.5
        $ps.RightMargin  = & $cm 0.5
        $ps.TopMargin    = & $cm 0.5
        $ps.BottomMargin = & $cm 0.5
        $ps.HeaderMargin = & $cm 0.8
        $ps.FooterMargin = & $cm 0.8
        $ps.PrintTitleRows = '$1:$2'
        $ps.CenterFooter = ''
        $ps.LeftFooter   = '&A'

        # --- PEU: la pagina DINS D'AQUESTA pestanya, no dins del PDF sencer ---
        # Per defecte l'Excel numera de correguda per tot el traball d'impressio:
        # si ANNEX II comenca a la pagina 540, el peu hi diria 540. Amb
        # FirstPageNumber = 1 cada pestanya torna a comencar per 1.
        $ps.FirstPageNumber = 1
        # I el TOTAL tampoc no pot ser '&N': en una exportacio de diverses
        # pestanyes, &N es el total del PDF sencer. El total d'aquesta pestanya
        # el sabem nosaltres amb Pages.Count, i s'ha de llegir AQUI, al final,
        # quan la paginacio ja esta decidida (orientacio, paper i ajust a
        # l'ample); si es llegis abans, el numero seria d'una altra pagina.
        $ps.RightFooter = ('P' + [char]0x00E0 + 'gina &P')
        try {
            $nPag = [int]$sh.PageSetup.Pages.Count
            if ($nPag -gt 0) { $ps.RightFooter = ('P' + [char]0x00E0 + 'gina &P de ' + $nPag) }
        } catch { }
    } catch { }
}

# Construeix el llibre amb les pestanyes TRIADES i el retorna (obert, sense
# desar). $seleccio son noms de _SgOpcionsExport; buit vol dir totes.
# Retorna @{ Ok; Error; Workbook; Excel; Resum } — $Resum = files per pestanya.
function _SgConstruirLlibre($seleccio) {
    if ($null -eq $seleccio -or @($seleccio).Count -eq 0) { $seleccio = @(_SgOpcionsExport) }
    $volEstes = _SgSeleccioTeEstes $seleccio
    $defs = @(_SgFullesTriades $seleccio)
    if (-not $volEstes -and $defs.Count -eq 0) {
        return @{ Ok=$false; Error="No s'ha triat cap pestanya per exportar." }
    }

    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok=$false; Error="No s'ha trobat cap Excel d'activitats." } }

    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) { return @{ Ok=$false; Error="No s'ha pogut iniciar Microsoft Excel." } }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # On som, per si peta: aquesta eina nomes es pot provar amb l'Excel de debo,
    # o sigui que quan falla ha de dir EXACTAMENT en quin punt i a quina pestanya.
    # Sense aixo, un error de COM nomes dona el missatge pelat i toca endevinar.
    $pas = "obrint l'Excel d'activitats"
    $wbOrig = $null
    try {
        $wbOrig = $excel.Workbooks.Open($latest.File.FullName, 0, $true)   # ReadOnly
        $pas = ("localitzant la fulla 'Est" + [char]0x00E8 + "s'")
        # NO passa per Read-FullaEstesa i es a posta: aquesta eina no nomes
        # LLEGEIX, tambe COPIA la fulla dins d'un llibre NOU amb la MATEIXA
        # instancia d'Excel, i despres l'exporta a PDF. Read-FullaEstesa tanca
        # l'Excel en sortir del cos, o sigui que aqui no hi cap. El que si que
        # comparteix es com es troba la fulla.
        $found = _TrobaFullaEstesa $wbOrig
        $shOrig = $found.Sheet
        if ($null -eq $shOrig) {
            return @{ Ok=$false; Error=("No s'ha trobat la fulla 'Est" + [char]0x00E8 + "s' a l'Excel d'activitats. Fulles disponibles: " + (@($found.Noms) -join ', ')) }
        }

        # 1) Llegim la taula sencera d'una sola vegada (molt mes rapid que anar
        #    cel·la a cel·la per COM). Nomes cal si s'ha triat algun LLISTAT: si
        #    l'usuari nomes vol la copia de la base, no te sentit llegir-la tota.
        $headers = @(); $pairs = @(); $files = New-Object System.Collections.ArrayList
        if ($defs.Count -gt 0) {
            $pas = "llegint les dades de l'Excel"
            $data = $shOrig.UsedRange.Value2
            if ($null -eq $data) { return @{ Ok=$false; Error="La fulla d'activitats es buida." } }
            $nRows = $data.GetLength(0); $nCols = $data.GetLength(1)
            for ($c = 1; $c -le $nCols; $c++) { $headers += [string]$data[1, $c] }
            # SENSE @(): _FindCampInfoPairs ja protegeix el seu retorn amb una coma
            # i un @() al voltant hi tornaria a posar la capa (vegeu _SgAplanaPairs).
            $pairs = _FindCampInfoPairs $headers
            for ($r = 2; $r -le $nRows; $r++) {
                $fila = New-Object object[] $nCols
                for ($c = 1; $c -le $nCols; $c++) { $fila[$c - 1] = $data[$r, $c] }
                [void]$files.Add($fila)
            }
        }

        # 2) Llibre nou. El full buit que hi ve de fabrica fa de PLACEHOLDER: un
        #    llibre no pot quedar-se sense cap fulla, o sigui que no s'esborra
        #    fins al final, quan ja hi ha les pestanyes de debo.
        $pas = 'creant el llibre nou'
        $wb = $excel.Workbooks.Add()
        while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
        if ($volEstes) {
            $pas = ("copiant la fulla 'Est" + [char]0x00E8 + "s' al llibre nou")
            $shOrig.Copy($wb.Sheets.Item(1))          # la posa ABANS del full buit
            $wb.Sheets.Item(1).Name = (_SgNomEstes)
            $wb.Sheets.Item($wb.Sheets.Count).Delete()   # fora el full buit que sobra
        }

        # 3) Les pestanyes de llistat que s'hagin triat
        $ara = Get-Date
        $resum = [ordered]@{}
        $avisos = New-Object System.Collections.ArrayList
        foreach ($def in $defs) {
            # Sheets.Add(Before, After, ...): per saltar-se 'Before' cal passar
            # Missing.Value, NO $null (amb $null el COM es pensa que li donem un
            # 'Before' buit i afegeix la pestanya al principi).
            $pas = ("creant la pestanya '" + [string]$def.Nom + "'")
            $sh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
            $sh.Name = [string]$def.Nom

            $pas = ("triant les activitats de '" + [string]$def.Nom + "'")
            $rows = @(_SgFilesPerFulla $def $headers $files $pairs)
            $resum[[string]$def.Nom] = $rows.Count

            $pas = ("escrivint el titol i les capceleres de '" + [string]$def.Nom + "'")
            $sh.Cells.Item(1, 1).Value2 = (_SgTitolFulla $def $ara)
            $cols = @('N') + @($def.Cols)
            for ($c = 0; $c -lt $cols.Count; $c++) { $sh.Cells.Item(2, $c + 1).Value2 = [string]$cols[$c] }

            if ($rows.Count -gt 0) {
                $pas = ("muntant la matriu de '" + [string]$def.Nom + "' (" + $rows.Count + " files x " + $cols.Count + " columnes)")
                $m = New-Object 'object[,]' $rows.Count, $cols.Count
                for ($i = 0; $i -lt $rows.Count; $i++) {
                    $f = @($rows[$i])
                    for ($j = 0; $j -lt $cols.Count; $j++) { $m[$i, $j] = if ($j -lt $f.Count) { $f[$j] } else { '' } }
                }
                # Les dades comencen a la fila 3 (1 = titol, 2 = capceleres).
                $pas = ("bolcant les dades a '" + [string]$def.Nom + "'")
                [void](_SgEscriuMatriu $sh 3 $m)
            }
            # El format es NOMES aparenca: si peta, val mes tenir el fitxer amb
            # les dades i un avis que no pas quedar-se sense res. (Mateixa lliço
            # que als reintents de la signatura: el try/catch va DINS del bucle,
            # perque una pestanya no s'endugui les altres.)
            $pas = ("donant format a '" + [string]$def.Nom + "'")
            try {
                _SgFormatarFulla $sh $def $rows.Count $excel
            } catch {
                [void]$avisos.Add((_SgTextError $_ $pas))
            }
        }

        # Si no s'ha triat la copia de la base, el full buit encara hi es: fora.
        if (-not $volEstes -and $wb.Sheets.Count -gt $defs.Count) {
            $pas = 'traient el full buit que sobra'
            $wb.Sheets.Item(1).Delete()
        }

        $wb.Sheets.Item(1).Activate()
        return @{ Ok=$true; Workbook=$wb; Excel=$excel; Resum=$resum; Origen=$latest.File.Name; Avisos=$avisos.ToArray() }
    } catch {
        try { if ($null -ne $excel) { $excel.Quit() } } catch { }
        return @{ Ok=$false; Error=(_SgTextError $_ $pas) }
    } finally {
        try { if ($null -ne $wbOrig) { $wbOrig.Close($false) } } catch { }
        try { if ($null -ne $excel) { $excel.ScreenUpdating = $true } } catch { }
    }
}

# Carpeta de sortida de l'eina: local\seguiment-gia\ (fora del repositori).
function _SgCarpetaSortida {
    $d = Get-LocalSubdir $RepoRoot 'Seguiment'
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    return [string]$d
}

# Tanca l'Excel que hem obert, passi el que passi.
function _SgTancar($excel, $wb) {
    try { if ($null -ne $wb) { $wb.Close($false) } } catch { }
    try { if ($null -ne $excel) { $excel.Quit() } } catch { }
    try { if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } } catch { }
}

# Genera i desa. $mode = 'excel' o 'pdf'; $seleccio, les pestanyes triades.
# Retorna la ruta o '' si ha fallat.
function _SgExportar([string]$mode, $seleccio) {
    $r = _SgConstruirLlibre $seleccio
    if (-not $r.Ok) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut generar el seguiment:`n$($r.Error)", 'Seguiment', 'OK', 'Error') | Out-Null
        return ''
    }
    $wb = $r.Workbook; $excel = $r.Excel
    $ara = Get-Date
    $dir = _SgCarpetaSortida
    $ext = if ($mode -eq 'pdf') { 'pdf' } else { 'xlsx' }
    # [string] a posta: _GetUniqueOutputPath torna la sortida de Join-Path, que
    # es un cmdlet i arriba embolcallada en un PSObject. Al COM li pot fer nosa.
    [string]$path = _GetUniqueOutputPath $dir (_SgNomFitxer $ara $ext)
    try {
        if ($mode -eq 'pdf') {
            # El llibre ja conte NOMES el que s'ha triat, o sigui que s'exporta
            # sencer. ExportAsFixedFormat es del WORKBOOK i respecta el PageSetup
            # de cada pestanya. Compte: $excel.ActiveWindow.SelectedSheets.ExportAsFixedFormat
            # NO existeix (SelectedSheets es una col·leccio 'Sheets'; el metode
            # nomes el tenen Workbook, Worksheet, Chart i Range) — l'Excel deia
            # "[System.__ComObject] no contiene ningun metodo llamado
            # 'ExportAsFixedFormat'".
            $wb.ExportAsFixedFormat($Script:SgXl.TypePDF, $path)
        } else {
            $wb.SaveAs($path, $Script:SgXl.OpenXMLBook)
        }
    } catch {
        $detall = _SgTextError $_ ("desant el fitxer en " + $ext)
        _SgTancar $excel $wb
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar el fitxer:`n`n$detall", 'Seguiment', 'OK', 'Error') | Out-Null
        return ''
    }
    _SgTancar $excel $wb

    $det = (@($r.Resum.Keys) | ForEach-Object { "  $_ : $($r.Resum[$_])" }) -join "`n"
    $msg = "Seguiment generat:`n$path`n`nActivitats per pestanya:`n$det`n`nBase de dades: $($r.Origen)"
    # Si alguna pestanya s'ha quedat sense format, es diu (les dades hi son).
    $avisos = @($r.Avisos)
    if ($avisos.Count -gt 0) {
        $msg += "`n`nAVIS: hi ha " + $avisos.Count + " pestanya(es) sense format. Les dades hi son igualment.`n" + ($avisos -join "`n")
    }
    $msg += "`n`nVols obrir-lo ara?"
    $rc = [System.Windows.Forms.MessageBox]::Show($msg, 'Seguiment', 'YesNo', 'Information')
    if ($rc -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Invoke-Item -LiteralPath $path } catch { }
    }
    return $path
}

# ----------------------------------------------------------------------------
# Finestra de l'eina: les caselles del que es vol exportar + els dos botons
# ----------------------------------------------------------------------------
function Invoke-SeguimentGia {
    $form = _NewForm
    $form.Text = 'Seguiment'
    $form.ClientSize = New-Object System.Drawing.Size(470, 336)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 72)
    $lbl.Size = New-Object System.Drawing.Size(430, 34)
    $lbl.Text = ("Genera els llistats de seguiment a partir de la base de dades del GIA." + "`r`n" +
                 "Marca les pestanyes que vols exportar:")
    [void]$form.Controls.Add($lbl)

    # Caselles: una per pestanya, totes marcades de bon principi. La llista surt
    # de _SgOpcionsExport, que es la MATEIXA que fa servir la construccio del
    # llibre, o sigui que no es poden desincronitzar.
    # CheckOnClick: sense aixo cal clicar exactament el quadradet (el primer clic
    # nomes selecciona la fila), i es un embolic.
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(20, 110)
    $clb.Size = New-Object System.Drawing.Size(430, 112)
    $clb.CheckOnClick = $true
    $clb.IntegralHeight = $false
    $clb.BorderStyle = 'FixedSingle'
    $clb.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    foreach ($op in @(_SgOpcionsExport)) { [void]$clb.Items.Add([string]$op, $true) }
    [void]$form.Controls.Add($clb)

    $lnk = New-Object System.Windows.Forms.LinkLabel
    $lnk.Location = New-Object System.Drawing.Point(20, 226)
    $lnk.Size = New-Object System.Drawing.Size(200, 18)
    $lnk.Text = 'Marcar-les / desmarcar-les totes'
    $lnk.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lnk.add_LinkClicked({
        # Si n'hi ha cap sense marcar, es marquen totes; si ja hi eren totes, es
        # desmarquen. Un sol enllac per als dos casos.
        $totes = $true
        for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $totes = $false; break } }
        for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, (-not $totes)) }
    }.GetNewClosure())
    [void]$form.Controls.Add($lnk)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Location = New-Object System.Drawing.Point(20, 248)
    $lbl2.Size = New-Object System.Drawing.Size(430, 42)
    $lbl2.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lbl2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbl2.Text = ("Els llistats surten en horitzontal i A3, ajustats a l'ample. La pestanya Est" + [char]0x00E8 +
                  "s son 152 columnes: per al PDF val m" + [char]0x00E9 + "s desmarcar-la." + "`r`n" +
                  "Necessita l'Excel treballant una estona: pot trigar.")
    [void]$form.Controls.Add($lbl2)

    $btnXls = New-Object System.Windows.Forms.Button
    $btnXls.Text = 'Exportar a Excel'
    $btnXls.Location = New-Object System.Drawing.Point(20, 292)
    $btnXls.Size = New-Object System.Drawing.Size(150, 32)
    _StylePrimaryButton $btnXls
    [void]$form.Controls.Add($btnXls)

    $btnPdf = New-Object System.Windows.Forms.Button
    $btnPdf.Text = 'Exportar a PDF'
    $btnPdf.Location = New-Object System.Drawing.Point(180, 292)
    $btnPdf.Size = New-Object System.Drawing.Size(150, 32)
    _StyleSecondaryButton $btnPdf
    [void]$form.Controls.Add($btnPdf)

    $btnTanca = New-Object System.Windows.Forms.Button
    $btnTanca.Text = 'Tancar'
    $btnTanca.Location = New-Object System.Drawing.Point(362, 292)
    $btnTanca.Size = New-Object System.Drawing.Size(88, 32)
    _StyleSecondaryButton $btnTanca
    $btnTanca.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnTanca)

    # Mentre l'Excel treballa, els botons queden desactivats i el cursor en
    # espera: si no, es pot clicar dues vegades i s'obren dos Excel.
    $ferExport = {
        param($mode)
        $sel = @($clb.CheckedItems | ForEach-Object { [string]$_ })
        if ($sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Marca almenys una pestanya per exportar.', 'Seguiment', 'OK', 'Information') | Out-Null
            return
        }
        $btnXls.Enabled = $false; $btnPdf.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { [void](_SgExportar $mode $sel) } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnXls.Enabled = $true; $btnPdf.Enabled = $true
        }
    }.GetNewClosure()
    $btnXls.add_Click({ & $ferExport 'excel' }.GetNewClosure())
    $btnPdf.add_Click({ & $ferExport 'pdf' }.GetNewClosure())

    $sub = 'Llistats de seguiment des de la base de dades del GIA'
    [void](_AddBrandHeader $form 'Seguiment' $sub 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}
