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

# Index (1-based) de la columna amb aquesta capcalera. La PRIMERA que hi
# coincideixi, com feia el MATCH de la plantilla: a la base de dades hi ha
# capceleres repetides ('Classificacio general annex' hi surt dues vegades) i la
# plantilla mostra sempre la primera. 0 si no hi es. $headers es 0-based.
function _SgColIndex($headers, [string]$nom) {
    $t = _NormalizeText $nom
    $h = @($headers)
    for ($i = 0; $i -lt $h.Count; $i++) {
        if ((_NormalizeText $h[$i]) -eq $t) { return ($i + 1) }
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
    $t = _NormalizeText $camp
    foreach ($p in @($pairs)) {
        $i = [int]$p.NomCol - 1
        if ($i -lt 0 -or $i -ge @($fila).Count) { continue }
        if ((_NormalizeText $fila[$i]) -eq $t) { return $p }
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
    return ((_NormalizeText $classificacio) -eq 'ii')
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
    foreach ($d in @($def.DataCols)) { $esData[(_NormalizeText $d)] = $true }

    # ANNEX II: la classificacio del CRITERI es la SEGONA columna amb aquest nom.
    $iClassCrit = 0
    $iDescr = 0
    if ($def.Tipus -eq 'annex') {
        $nomClass = _NormalizeText ('Classificaci' + [char]0x00F3 + ' general annex')
        $vistes = 0
        $h = @($headers)
        for ($i = 0; $i -lt $h.Count; $i++) {
            if ((_NormalizeText $h[$i]) -eq $nomClass) {
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
            if ($esData.ContainsKey((_NormalizeText $nomCol))) { $v = _FormatDateOnly $v }
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

    # --- Fila 2: capceleres ---
    $hdr = $sh.Range($sh.Cells.Item(2, 1), $sh.Cells.Item(2, $nCols))
    $hdr.Font.Name = 'Arial'; $hdr.Font.Size = 10; $hdr.Font.Bold = $true
    $hdr.Interior.Color = 15917529                        # gris blavos clar
    $hdr.HorizontalAlignment = -4108

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
        $ps.RightFooter  = ('P' + [char]0x00E1 + 'gina &P')
    } catch { }
}

# Construeix el llibre sencer i el retorna (obert, sense desar).
# Retorna @{ Ok; Error; Workbook; Excel; Resum } — $Resum = files per pestanya.
function _SgConstruirLlibre {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok=$false; Error="No s'ha trobat cap Excel d'activitats." } }

    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) { return @{ Ok=$false; Error="No s'ha pogut iniciar Microsoft Excel." } }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wbOrig = $null
    try {
        $wbOrig = $excel.Workbooks.Open($latest.File.FullName, 0, $true)   # ReadOnly
        $found = _FindEstesSheet $wbOrig
        $shOrig = $found.Sheet
        if ($null -eq $shOrig) {
            return @{ Ok=$false; Error=("No s'ha trobat la fulla 'Est" + [char]0x00E8 + "s' a l'Excel d'activitats.") }
        }

        # 1) Llegim la taula sencera d'una sola vegada (molt mes rapid que anar
        #    cel·la a cel·la per COM).
        $data = $shOrig.UsedRange.Value2
        if ($null -eq $data) { return @{ Ok=$false; Error="La fulla d'activitats es buida." } }
        $nRows = $data.GetLength(0); $nCols = $data.GetLength(1)
        $headers = @()
        for ($c = 1; $c -le $nCols; $c++) { $headers += [string]$data[1, $c] }
        # SENSE @(): _FindCampInfoPairs ja protegeix el seu retorn amb una coma i
        # un @() al voltant hi tornaria a posar la capa (vegeu _SgAplanaPairs).
        $pairs = _FindCampInfoPairs $headers
        $files = New-Object System.Collections.ArrayList
        for ($r = 2; $r -le $nRows; $r++) {
            $fila = New-Object object[] $nCols
            for ($c = 1; $c -le $nCols; $c++) { $fila[$c - 1] = $data[$r, $c] }
            [void]$files.Add($fila)
        }

        # 2) Llibre nou amb la fulla Estes COPIADA (valors i format de l'origen).
        $wb = $excel.Workbooks.Add()
        while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
        $shOrig.Copy($wb.Sheets.Item(1))          # la posa ABANS del full buit
        $wb.Sheets.Item(1).Name = ('Est' + [char]0x00E8 + 's')
        $wb.Sheets.Item($wb.Sheets.Count).Delete()   # fora el full buit que sobra

        # 3) Les 5 pestanyes de llistat
        $ara = Get-Date
        $resum = [ordered]@{}
        foreach ($def in @(_SgFullesDef)) {
            # Sheets.Add(Before, After, ...): per saltar-se 'Before' cal passar
            # Missing.Value, NO $null (amb $null el COM es pensa que li donem un
            # 'Before' buit i afegeix la pestanya al principi).
            $sh = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
            $sh.Name = [string]$def.Nom
            $rows = @(_SgFilesPerFulla $def $headers $files $pairs)
            $resum[[string]$def.Nom] = $rows.Count

            $sh.Cells.Item(1, 1).Value2 = (_SgTitolFulla $def $ara)
            $cols = @('N') + @($def.Cols)
            for ($c = 0; $c -lt $cols.Count; $c++) { $sh.Cells.Item(2, $c + 1).Value2 = [string]$cols[$c] }

            if ($rows.Count -gt 0) {
                # Matriu [files x columnes] i UNA sola assignacio: per COM, la
                # diferencia entre aixo i escriure cel·la a cel·la es de minuts
                # a segons.
                $m = New-Object 'object[,]' $rows.Count, $cols.Count
                for ($i = 0; $i -lt $rows.Count; $i++) {
                    $f = @($rows[$i])
                    for ($j = 0; $j -lt $cols.Count; $j++) { $m[$i, $j] = if ($j -lt $f.Count) { $f[$j] } else { '' } }
                }
                $dest = $sh.Range($sh.Cells.Item(3, 1), $sh.Cells.Item(2 + $rows.Count, $cols.Count))
                $dest.Value2 = $m
            }
            _SgFormatarFulla $sh $def $rows.Count $excel
        }

        $wb.Sheets.Item(1).Activate()
        return @{ Ok=$true; Workbook=$wb; Excel=$excel; Resum=$resum; Origen=$latest.File.Name }
    } catch {
        try { if ($null -ne $excel) { $excel.Quit() } } catch { }
        return @{ Ok=$false; Error=$_.Exception.Message }
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

# Genera i desa. $mode = 'excel' o 'pdf'. Retorna la ruta o '' si ha fallat.
function _SgExportar([string]$mode) {
    $r = _SgConstruirLlibre
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
            # NOMES les 5 pestanyes de llistat: la fulla Estes son 152 columnes
            # i a la plantilla ni tan sols esta preparada per imprimir. Es
            # seleccionen com a grup perque l'exportacio nomes agafi la seleccio.
            $noms = @(@(_SgFullesDef) | ForEach-Object { [string]$_.Nom })
            $wb.Sheets.Item($noms[0]).Select()
            for ($i = 1; $i -lt $noms.Count; $i++) { $wb.Sheets.Item($noms[$i]).Select($false) }
            # SelectedSheets (no ActiveSheet): ActiveSheet nomes exportaria la
            # pestanya activa i el PDF sortiria amb un sol llistat.
            $excel.ActiveWindow.SelectedSheets.ExportAsFixedFormat($Script:SgXl.TypePDF, $path)
        } else {
            $wb.SaveAs($path, $Script:SgXl.OpenXMLBook)
        }
    } catch {
        _SgTancar $excel $wb
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar el fitxer:`n$($_.Exception.Message)", 'Seguiment', 'OK', 'Error') | Out-Null
        return ''
    }
    _SgTancar $excel $wb

    $det = (@($r.Resum.Keys) | ForEach-Object { "  $_ : $($r.Resum[$_])" }) -join "`n"
    $msg = "Seguiment generat:`n$path`n`nActivitats per pestanya:`n$det`n`nBase de dades: $($r.Origen)`n`nVols obrir-lo ara?"
    $rc = [System.Windows.Forms.MessageBox]::Show($msg, 'Seguiment', 'YesNo', 'Information')
    if ($rc -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Invoke-Item -LiteralPath $path } catch { }
    }
    return $path
}

# ----------------------------------------------------------------------------
# Finestra de l'eina: nomes els DOS botons d'exportacio
# ----------------------------------------------------------------------------
function Invoke-SeguimentGia {
    $form = _NewForm
    $form.Text = 'Seguiment'
    $form.ClientSize = New-Object System.Drawing.Size(470, 250)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 76)
    $lbl.Size = New-Object System.Drawing.Size(430, 84)
    $lbl.Text = ("Genera el seguiment d'activitats a partir de la base de dades del GIA, amb una pestanya per cada llistat:" + "`r`n`r`n" +
                 ("   Est" + [char]0x00E8 + "s  " + [char]0x00B7 + "  PRECINTES  " + [char]0x00B7 + "  DEN" + [char]0x00DA + "NCIES  " + [char]0x00B7 + "  REQUERIT DECRET") + "`r`n" +
                 ("   SONOMETRIA  " + [char]0x00B7 + "  ANNEX II"))
    [void]$form.Controls.Add($lbl)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Location = New-Object System.Drawing.Point(20, 164)
    $lbl2.Size = New-Object System.Drawing.Size(430, 30)
    $lbl2.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lbl2.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbl2.Text = ("Al PDF hi van nom" + [char]0x00E9 + "s els cinc llistats (horitzontal, A3, ajustat a l'ample). Necessita l'Excel obert una estona: pot trigar.")
    [void]$form.Controls.Add($lbl2)

    $btnXls = New-Object System.Windows.Forms.Button
    $btnXls.Text = 'Exportar a Excel'
    $btnXls.Location = New-Object System.Drawing.Point(20, 202)
    $btnXls.Size = New-Object System.Drawing.Size(150, 32)
    _StylePrimaryButton $btnXls
    [void]$form.Controls.Add($btnXls)

    $btnPdf = New-Object System.Windows.Forms.Button
    $btnPdf.Text = 'Exportar a PDF'
    $btnPdf.Location = New-Object System.Drawing.Point(180, 202)
    $btnPdf.Size = New-Object System.Drawing.Size(150, 32)
    _StyleSecondaryButton $btnPdf
    [void]$form.Controls.Add($btnPdf)

    $btnTanca = New-Object System.Windows.Forms.Button
    $btnTanca.Text = 'Tancar'
    $btnTanca.Location = New-Object System.Drawing.Point(362, 202)
    $btnTanca.Size = New-Object System.Drawing.Size(88, 32)
    _StyleSecondaryButton $btnTanca
    $btnTanca.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnTanca)

    # Mentre l'Excel treballa, els botons queden desactivats i el cursor en
    # espera: si no, es pot clicar dues vegades i s'obren dos Excel.
    $ferExport = {
        param($mode)
        $btnXls.Enabled = $false; $btnPdf.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { [void](_SgExportar $mode) } finally {
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
