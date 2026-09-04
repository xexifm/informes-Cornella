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
function _TrobaFullaEstesa($wb) {
    $noms = @()
    foreach ($s in $wb.Sheets) {
        $noms += [string]$s.Name
        if ((_NormalitzaText $s.Name) -eq 'estes') { return @{ Sheet = $s; Noms = $noms } }
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
function Read-FullaEstesa($excelFile, [scriptblock]$cos) {
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) {
        throw "No s'ha pogut iniciar Microsoft Excel. Comprova que estigui instal-lat i que l'hagis obert almenys un cop."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)   # ReadOnly
        $trobada = _TrobaFullaEstesa $wb
        $sh = $trobada.Sheet
        if ($null -eq $sh) {
            throw ("No s'ha trobat la fulla 'Est" + [char]0x00E8 + "s' al fitxer Excel. Fulles disponibles: " + (@($trobada.Noms) -join ', '))
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
