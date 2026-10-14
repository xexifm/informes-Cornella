#requires -Version 5.1
<#
.SYNOPSIS
  Llegir la fulla "Estes" de l'Excel d'activitats. Una sola manera, per a tot.

.DESCRIPTION
  Hi havia SET funcions que feien exactament la mateixa seqüencia -obrir l'Excel
  per COM, buscar la fulla "Estes", llegir UsedRange.Value2, treure'n files i
  columnes i la capcalera, i tancar-ho tot- i nomes es diferenciaven en QUE en
  treien. Set copies de vint linies:

    Activitats.ps1        Initialize-ActivitatsCache
    Informes.ps1          _ReadExcelCampInfoPerGia
    ControlsPeriodics.ps1 _ReadControlsPeriodics
    SeguimentGia.ps1      la lectura de la base (nomes la meitat de lectura)
    rutes/Ruta.ps1        Read-ActivitatsForRoute
    rutes/Precintades.ps1 Read-PrecintadesFromExcel
    rutes/Coordenades.ps1 Read-CoordenadesFromExcel

  Aixo NO era nomes lletgesa: les copies s'havien anat separant, i les
  diferencies eren DEFECTES.

  1. LES TRES DE 'rutes' DEIXAVEN UN EXCEL.EXE ORFE. El seu finally feia
     '$wb.Close($false)' i '$excel.Quit()' SENSE try. Si el Close peta -i un
     llibre obert des d'una unitat de xarxa ho pot fer-, el Quit no s'executa
     mai: l'Excel es queda corrent, invisible, amb el fitxer agafat. Les copies
     de suport/ si que els emboliquen. Aqui cada pas del tancament va dins del
     seu propi try.

  2. LES MATEIXES TRES NO COMPROVAVEN EL $null. New-Object -ComObject pot tornar
     $null SENSE llancar (Excel no instal-lat, primera execucio pendent, COM
     trencat). Llavors la primera linia que hi toca peta amb un "metode sobre
     NULL" quaranta linies mes avall, que no diu res. Aqui es comprova de
     seguida i es llanca un missatge que si que ho diu.

  3. LA FULLA ES BUSCAVA DE DUES MANERES (_FindEstesSheet a Activitats.ps1 i
     _RutaFindEstesSheet a Ruta.ps1) i el missatge d'error de no trobar-la
     nomes deia els noms de les fulles a UNA de les set. Ara el diu sempre: si
     algu reanomena la pestanya, el missatge ja diu com es diu ara.

.NOTES
  NOMES DEFINEIX FUNCIONS. Es el que permet que el carreguin els DOS processos:
  rutes/Ruta.ps1 corre a part i no pot carregar UiComuns.ps1, que si que executa
  coses en carregar-se (AppUserModelID, icona). Mateix motiu i mateix patro que
  UiFinestra.ps1 i Json.ps1.
#>

# Normalitza un text per COMPARAR: sense diacritics, sense espais als extrems i
# en minuscules.
#
# Abans n'hi havia dues, una a cada proces: la d'Activitats.ps1 (45 usos) i la
# de Ruta.ps1 (6). Feien el mateix; l'unica diferencia real era ToLower() contra
# ToLowerInvariant() -el Normalize(FormC) que feia la segona es un no-op un cop
# tretes les marques-.
#
# ES QUEDA ToLowerInvariant. Un normalitzador que serveix per comparar NO POT
# dependre de l'idioma del Windows: amb la cultura turca, ToLower() de 'I' dona
# 'i' sense punt i la comparacio deixa de coincidir. Aqui es fan servir per
# trobar la fulla "Estes", les capcaleres de columna i les sigles de via del
# Cadastre: tot aixo ha de donar el mateix en qualsevol equip.
function _NormalitzaText($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    return (($t -replace '\p{Mn}', '').Trim().ToLowerInvariant())
}

# La fila 1 de la matriu com a array 0-BASED de noms de columna, ja retallats.
#
# El bucle estava copiat a quatre llocs. Es consumeix TAL QUAL, SENSE @():
# el 'return ,[string[]]' hi es a posta perque un Excel d'UNA sola columna no
# torni una cadena pelada -llavors $headers[0] donaria el primer CARACTER, no la
# capcalera-. Hi ha prova d'aquest cas.
function _HeadersDeFila1($data, [int]$cols) {
    if ($null -eq $data -or $cols -lt 1) { return ,([string[]]@()) }
    $out = New-Object 'string[]' $cols
    for ($c = 1; $c -le $cols; $c++) {
        $v = $data[1, $c]
        $out[$c - 1] = if ($null -eq $v) { '' } else { ([string]$v).Trim() }
    }
    return ,$out
}

# La fulla "Estes" d'un llibre ja obert, o $null. Torna tambe els noms de totes
# les fulles, que es el que fa util el missatge quan no la troba.
function _TrobaFullaEstesa($wb, [string]$fulla = 'estes') {
    $cerca = _NormalitzaText $fulla
    $noms = @()
    foreach ($s in $wb.Sheets) {
        $noms += [string]$s.Name
        if ((_NormalitzaText $s.Name) -eq $cerca) { return @{ Sheet = $s; Noms = $noms } }
    }
    return @{ Sheet = $null; Noms = $noms }
}

# LA SEQÜENCIA COMPARTIDA: obre l'Excel en NOMES LECTURA, busca la fulla
# "Estes", llegeix la matriu i li passa tot al cos; despres tanca passi el que
# passi.
#
# El cos arriba com a SCRIPTBLOCK i rep un context:
#
#   @{ Data; Rows; Cols; Headers; Sheet; Noms }
#
# Es el mateix patro que Write-InformeDocx (MotorInforme.ps1), i com alla el
# scriptblock va SENSE .GetNewClosure(): ha de veure els locals del cridador en
# temps d'execucio, i una closure en copiaria els valors del moment de crear-la.
#
# QUE DECIDEIX AQUESTA FUNCIO i que no:
#   - Si l'Excel no arrenca o la fulla no hi es, LLANCA. Els criders que volen
#     un @{ Ok=$false; Error=... } en lloc d'una excepcio s'ho emboliquen amb un
#     try/catch, que es una linia i deixa el missatge a mans de qui el mostra.
#   - Un Excel BUIT (Value2 = $null) NO es un error: el cos es crida igualment
#     amb Data=$null, Rows=0, Cols=0 i Headers buit. Cada eina en torna una cosa
#     diferent (un mapa buit, una llista buida, un objecte amb comptadors) i
#     aquesta funcio no en pot decidir cap.
#
# COM ES CONSUMEIX QUAN EL COS TORNA UNA LLISTA. El crider ASSIGNA primer i
# despres torna amb COMA:
#
#     function Read-XFromExcel($f) {
#         $out = Read-FullaEstesa $f { param($x) ...; return ,@($registres) }
#         return ,@($out)
#     }
#
# NO 'return (Read-FullaEstesa ...)' directament. Aixo es MESURAT, no deduit:
# aixi hi ha DOS 'return' seguits, el pipeline desenrotlla una capa a cada un, i
# una llista d'UN SOL registre arriba al crider com l'objecte PELAT (sortia un
# PSCustomObject alla on toca un Object[]). Amb l'assignacio pel mig no passa,
# perque assignar NO desenrotlla. Els tres criders de rutes/ ho fan aixi.
#
# Aqui dins NO hi ha cap coma al 'return $resultat', i tambe esta comprovat:
# amb la forma d'us de dalt, posar-n'hi una no canvia res (0, 1, 2 i 5
# registres donen el mateix), i una coma que no fa res nomes despista.
#
# -Desa: obre el llibre per ESCRIURE (el cos hi escriu amb $x.Sheet) i el desa
# en acabar. Nomes si el cos acaba be: un llibre a mig escriure no es desa mai.
# El fa servir l'"Excel per importar" de Coordenades, que escriu sobre una COPIA
# de la base (mai sobre l'original). Es aqui i no en una funcio a part perque
# obrir i tancar l'Excel es fa en UN sol lloc (hi ha guard).
#
# -Fulla: una altra fulla que no sigui "Estes" (l'Excel d'ESTABLIMENTS porta la
# fulla "Establiments"). Es compara normalitzada, com "Estes".
function Read-FullaEstesa($excelFile, [scriptblock]$cos, [switch]$Desa, [string]$Fulla = '') {
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) {
        throw "No s'ha pogut iniciar Microsoft Excel. Comprova que estigui instal-lat i que l'hagis obert almenys un cop."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, (-not $Desa))   # ReadOnly, tret de -Desa
        $nomFulla = if ($Fulla -ne '') { $Fulla } else { 'Est' + [char]0x00E8 + 's' }
        $trobada = _TrobaFullaEstesa $wb $nomFulla
        $sh = $trobada.Sheet
        if ($null -eq $sh) {
            throw ("No s'ha trobat la fulla '" + $nomFulla + "' al fitxer Excel. Fulles disponibles: " + (@($trobada.Noms) -join ', '))
        }

        $data = $sh.UsedRange.Value2
        $rows = 0; $cols = 0
        if ($null -ne $data) { $rows = $data.GetLength(0); $cols = $data.GetLength(1) }

        # Cel: llegeix UNA cel·la ja retallada, o '' si la fila/columna no hi
        # es. Aquest bloc estava copiat als CINC cossos -Activitats,
        # ControlsPeriodics, Ruta, Precintades i Coordenades-, identic fins a
        # l'espaiat: es va unificar la carcassa i es va deixar el helper a dins
        # de cada cos. Ara ve amb el context.
        #
        # EL .GetNewClosure() ES OBLIGATORI, i esta MESURAT: sense ell, el bloc
        # resol $data i $cols quan es crida, i els cossos es declaren els SEUS
        # ($data = $x.Data). Amb un cos que faci aixo -i en fan tots- el lector
        # acabaria indexant la variable del cos: a la prova, treure la closure
        # fa que '$get 2 1' torni 'O' en lloc del valor de la cel·la.
        #
        # Compte que aixo va AL REVES que els blocs -Desa/-Restaurar de
        # Show-EditorAssumpteCos, que NO en porten: alla el bloc s'escriu al
        # crider i s'executa a la funcio compartida (i per tant ja veu els
        # locals de qui el va escriure); aqui s'escriu a la funcio compartida i
        # s'executa al crider. La direccio decideix la resposta, i per aixo cada
        # cas te la seva prova.
        $cel = {
            param($r, $c)
            if ($c -lt 1 -or $c -gt $cols) { return '' }
            $v = $data[$r, $c]
            if ($null -eq $v) { return '' }
            return ([string]$v).Trim()
        }.GetNewClosure()

        $resultat = & $cos @{
            Data    = $data
            Rows    = $rows
            Cols    = $cols
            Headers = (_HeadersDeFila1 $data $cols)
            Cel     = $cel
            Sheet   = $sh
            Noms    = @($trobada.Noms)
        }

        if ($Desa) { $wb.Save() }
        return $resultat
    } finally {
        # CADA PAS DINS DEL SEU try. Si el Close peta i s'endu el Quit, l'Excel
        # es queda corrent invisible amb el fitxer agafat: es el defecte que
        # tenien les tres copies de rutes/.
        if ($null -ne $wb)    { try { $wb.Close($false) } catch { } }
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        if ($null -ne $excel) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
    }
}

# ----------------------------------------------------------------------------
# Els "Camp Info" de l'Excel d'activitats (parelles Nom/Valor). El fan servir
# "Comprovar Excel" i "Seguiment"; vivia a Informes.ps1 i una eina depenia de
# l'altra per una cosa que no es de cap de les dues.
# ----------------------------------------------------------------------------
# PURA i testejable. A partir de la fila de capcaleres (array 0-based de textos),
# localitza els parells de columnes "Camp Info N - Nom" / "Camp Info N - Valor"
# i retorna una llista de @{ NomCol; ValorCol } (indexs 1-based, com Excel).
function _FindCampInfoPairs($headers) {
    $h = @($headers)
    $pairs = @()
    for ($i = 0; $i -lt $h.Count; $i++) {
        $n = _NormalitzaText $h[$i]
        if ($n -match '^camp info\s+(\d+)\s*-\s*nom$') {
            $num = $Matches[1]
            $target = _NormalitzaText ("camp info $num - valor")
            $valorCol = 0
            for ($j = 0; $j -lt $h.Count; $j++) {
                if ((_NormalitzaText $h[$j]) -eq $target) { $valorCol = $j + 1; break }
            }
            if ($valorCol -gt 0) { $pairs += @{ NomCol = ($i + 1); ValorCol = $valorCol } }
        }
    }
    return ,@($pairs)
}

# ----------------------------------------------------------------------------
# PRECINTADA? Vivia a rutes/Precintades.ps1; ara la fan servir tambe el
# "Planol activitats" i qui la necessiti, i aquest es el fitxer de l'Excel
# d'activitats que carreguen tots dos processos.
# ----------------------------------------------------------------------------
# Nom EXACTE (normalitzat) del camp lliure que marca una activitat precintada.
$Script:PrecCampNom = 'precinte activitat?'

# Cert si un camp lliure (Nom/Valor) indica una activitat PRECINTADA: el Nom es
# "PRECINTE ACTIVITAT?" i el Valor comenca per "SI" (com a paraula: "SI",
# "SI, ...", "SI ..."; NO "SITUACIO..."). Insensible a accents i majuscules.
function Test-IsPrecintada([string]$nom, [string]$valor) {
    if ((_NormalitzaText $nom) -ne $Script:PrecCampNom) { return $false }
    $v = _NormalitzaText $valor
    if ($v -eq '') { return $false }
    return [bool]($v -match '^si\b')
}

# L'ID Activitat d'una cel.la, com a text: un numero de l'Excel (101.0) -> '101',
# sense decimals, com a Ruta. PURA. La fan servir Coordenades (lectura de la
# base i l'Excel per importar, que ha de casar el repas amb la base pel mateix
# ID) i el Planol activitats.
function Get-IdDeCella($cell) {
    if ($null -eq $cell) { return '' }
    if ($cell -is [double]) {
        if ([math]::Floor($cell) -eq $cell) { return [string][long]$cell }
        return [string]$cell
    }
    return ([string]$cell).Trim()
}

# ----------------------------------------------------------------------------
# QUIN FITXER ES LA BASE: els noms de l'Excel d'activitats i d'establiments
# ----------------------------------------------------------------------------
# L'usuari (octubre 2026): "a partir d'ara, els arxius que s'han de llegir com a
# base de dades d'activitats i d'establiments es diran
# ExportacioEstabliments_20261014.xls / ExportacioActivitats_20261014.xls". Fins
# llavors eren 'AAAA-MM-DD ACTIVITATS.xls' (i '..._ESTABLIMENTS.xls'), i la
# carpeta encara en pot tenir d'aquells: s'accepten tots dos noms i mana la data.
# Abans hi havia DUES cerques (Activitats.ps1 i rutes/Ruta.ps1, una per proces)
# i ja no acceptaven el mateix: la de rutes admetia '_' entre la data i el nom i
# la de suport no. Ara son aqui, que carreguen els dos processos.
#
# El tipus: 'ACTIVITATS' o 'ESTABLIMENTS'.
$Script:ExcelBaseNoms = @{ ACTIVITATS = 'Activitats'; ESTABLIMENTS = 'Establiments' }

# La data que porta el NOM d'un fitxer de la base, o $null si no ho es. Nom nou
# 'Exportacio<Tipus>_AAAAMMDD.xls(x)' (amb o sense accent, i amb el ' (1)' que
# hi posa el navegador si es baixa dues vegades) o l'antic 'AAAA-MM-DD TIPUS'.
# PURA.
function _ExcelBaseDataDelNom([string]$nom, [string]$tipus = 'ACTIVITATS') {
    $t = [string]$Script:ExcelBaseNoms[$tipus.ToUpperInvariant()]
    if (-not $t) { return $null }
    $d = [datetime]::MinValue
    $nou = '(?i)^Exportaci[o' + [char]0x00F3 + ']' + $t + '_(\d{8})(\s*\(\d+\))?\.(xls|xlsx)$'
    if ($nom -match $nou -and [datetime]::TryParseExact($Matches[1], 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    $vell = '(?i)^(\d{4}-\d{2}-\d{2})[\s_]+' + $t + '\.(xls|xlsx)$'
    if ($nom -match $vell -and [datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

# La data (AAAA-MM-DD) que porta un text que diu de quin Excel surt una cosa
# (el 'Source' d'activitats.json, el 'Font' de precintades.json): '' si no en
# porta. Amb tots dos noms. PURA.
function _ExcelBaseDataDeText([string]$text) {
    $d = [datetime]::MinValue
    if ([string]$text -match '(?i)Exportaci\S*?_(\d{8})' -and [datetime]::TryParseExact($Matches[1], 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
    if ([string]$text -match '(\d{4}-\d{2}-\d{2})') { return $Matches[1] }
    return ''
}

# El mes recent d'una carpeta: @{ File; Date } o $null. A la mateixa data, el
# que s'ha desat mes tard (un ' (1)' baixat despres).
function Find-LatestBaseExcelIn($dir, [string]$tipus = 'ACTIVITATS') {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $cands = @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name.StartsWith('~$')) { return }
        $dt = _ExcelBaseDataDelNom $_.Name $tipus
        if ($null -ne $dt) { [pscustomobject]@{ File = $_; Date = $dt } }
    } | Sort-Object -Property @{ Expression = { $_.Date }; Descending = $true }, @{ Expression = { $_.File.LastWriteTime }; Descending = $true })
    if ($cands.Count -eq 0) { return $null }
    return $cands[0]
}

# Com es diu el fitxer que es busca, per als missatges ("no s'ha trobat..."). PURA.
function _ExcelBaseNomText([string]$tipus = 'ACTIVITATS') {
    $t = [string]$Script:ExcelBaseNoms[$tipus.ToUpperInvariant()]
    return ("'Exportacio" + $t + "_AAAAMMDD.xls' (o l'antic 'AAAA-MM-DD " + $tipus.ToUpperInvariant() + ".xls')")
}

# ----------------------------------------------------------------------------
# LES COLUMNES PER TERMES I LES DE CONTACTE (octubre 2026)
# ----------------------------------------------------------------------------
# La columna (1-based) de la capcalera que conte TOTS els termes de $amb i CAP
# dels de $sense (normalitzat: minuscules, sense accents). 0 si no n'hi ha. Abans
# vivia nomes a Activitats.ps1 (_FindColIndex, que ara hi delega) i el proces de
# rutes no la podia fer servir; el Planol activitats necessita les mateixes
# columnes de contacte (l'usuari: "a la fitxa vull el nom de la rao social, el
# del representant legal i els mobils, i des del mobil poder trucar"). PURA.
function _ColumnaAmbTermes($headers, [string[]]$amb, [string[]]$sense) {
    $arr = @($headers)
    for ($i = 0; $i -lt $arr.Count; $i++) {
        $h = _NormalitzaText ([string]$arr[$i])
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $ok = $true
        foreach ($m in @($amb)) { if (-not $h.Contains((_NormalitzaText $m))) { $ok = $false; break } }
        if ($ok -and $null -ne $sense) { foreach ($m in @($sense)) { if ($h.Contains((_NormalitzaText $m))) { $ok = $false; break } } }
        if ($ok) { return $i + 1 }
    }
    return 0
}

# LES COLUMNES DE CONTACTE de la fulla Estes, en UN sol lloc: les fa servir el
# repas de contactes i el correu (Activitats.ps1) i la fitxa del Planol
# activitats (rutes/PlanolDades.ps1). Torna @{ NIF; TELEFON; MOBIL; EMAIL;
# REP_NOM; REP_NIF; REP_TELEFON; REP_MOBIL; REP_EMAIL } (0 = no hi es). El mobil
# i el correu de la rao social tenen el seu index historic (23 i 25) de
# respatller. PURA.
function Get-ColumnesContacte($headers) {
    $mail = _ColumnaAmbTermes $headers @('rao soc', 'mail') $null
    if ($mail -eq 0) { $mail = 25 }
    $mob = _ColumnaAmbTermes $headers @('rao soc', 'mobil') $null
    if ($mob -eq 0) { $mob = 23 }
    $c = @{
        NIF         = (_ColumnaAmbTermes $headers @('rao soc', 'nif') $null)
        TELEFON     = (_ColumnaAmbTermes $headers @('rao soc', 'telefon') @('mobil'))
        MOBIL       = $mob
        EMAIL       = $mail
        REP_NOM     = (_ColumnaAmbTermes $headers @('representant legal') @('nif', 'mobil', 'telefon', 'mail'))
        REP_NIF     = (_ColumnaAmbTermes $headers @('rep', 'leg', 'nif') $null)
        REP_TELEFON = (_ColumnaAmbTermes $headers @('rep', 'leg', 'telefon') @('mobil'))
        REP_MOBIL   = (_ColumnaAmbTermes $headers @('rep', 'leg', 'mobil') $null)
        REP_EMAIL   = (_ColumnaAmbTermes $headers @('rep', 'leg', 'mail') $null)
    }
    # Si el GIA les anomena sense "Rao soc.", la primera que no sigui del
    # representant ni de l'emplacament.
    if ($c.NIF -eq 0)     { $c.NIF = _ColumnaAmbTermes $headers @('nif') @('rep', 'emp') }
    if ($c.TELEFON -eq 0) { $c.TELEFON = _ColumnaAmbTermes $headers @('telefon') @('rep', 'mobil', 'emp', 'establ') }
    return $c
}
