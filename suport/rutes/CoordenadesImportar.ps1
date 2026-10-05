<#
  CoordenadesImportar.ps1 - L'"Excel per importar": una COPIA de la base
  d'activitats amb les coordenades que s'han corregit al mapa, en VERMELL.

  Per que (peticio de l'usuari, octubre 2026): les correccions no les entra
  l'usuari, les IMPORTA una altra persona, i les necessita amb EL MATEIX FORMAT
  que la base de dades original. Per aixo no es munta cap Excel nou: es COPIA
  el fitxer de la base tal qual (mateixa extensio, mateixes columnes, mateix
  tot) i nomes s'hi reescriuen les cel.les 'UTM X' i 'UTM Y' de les activitats
  corregides, en vermell perque qui ho importi les vegi d'una ullada.

  D'on surten les correccions: del .xlsx que l'usuari es baixa del mapa (el
  "repas"). Es UN SOL FITXER per a tot: la copia de seguretat del repas (el
  mapa el pot tornar a carregar) i l'entrada d'aquest pas.

  Tres capes, i nomes l'ultima toca l'Excel:
    Read-RepasXlsx             llegeix el .xlsx del repas (ZIP + XML, sense
                               Excel: tal com surt del mapa o desat de nou amb
                               l'Excel, comprimit i amb sharedStrings)
    Get-CorreccionsDelRepas    que s'ha de canviar, i que no (PURA)
    Get-EscripturesCoordenades quines cel.les de la base s'escriuen (PURA)
    Set-CoordenadesALaBase     les escriu, amb Read-FullaEstesa -Desa (COM)

  L'ORIGINAL NO ES TOCA MAI: s'escriu sobre la copia. I una activitat la
  coordenada de la qual a la base JA NO es la que hi havia quan es va repassar
  NO es toca: vol dir que la base ha canviat (potser ja s'ha corregit), i
  sobreescriure-la seria desfer feina d'algu altre. Es llista a part.

  NOMES DEFINEIX FUNCIONS, i NO en crida cap de Coordenades.ps1 (que n'es el
  client; hi ha guard de cicles): la finestra que ho llança,
  Invoke-CoordExcelImportar, es alla. Depen de Ruta.ps1 (ConvertTo-UtmNumber,
  Find-HeaderColumn), Excel.ps1 i Geocodificador.ps1 (Test-CoordPlausible).
#>

# Les columnes del repas que es fan servir, normalitzades (_NormalitzaText).
# Han de coincidir amb CAPCALERA de CoordenadesMapa.html.
$Script:RepasColId     = 'id gia'
$Script:RepasColXNova  = 'utm x (nova)'
$Script:RepasColYNova  = 'utm y (nova)'
$Script:RepasColXVella = 'utm x (excel)'
$Script:RepasColYVella = 'utm y (excel)'
$Script:RepasColBase   = 'base de dades'

# Vermell per a Font.Color. L'Excel el vol en BGR: 0x0000FF = 255.
$Script:CoordColorCorregida = 255

# ----------------------------------------------------------------------------
# LECTOR DEL .xlsx DEL REPAS
# ----------------------------------------------------------------------------
function _XlsxXml($zip, [string]$nom) {
    $e = $zip.GetEntry($nom)
    if ($null -eq $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $text = $sr.ReadToEnd() } finally { $sr.Dispose() }
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($text)
    return $doc
}

# Columna 0-based d'una referencia de cel.la: 'A1' -> 0, 'K12' -> 10, 'AA3' -> 26.
function _XlsxColumna([string]$ref) {
    $n = 0
    foreach ($ch in $ref.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { break }
        $n = $n * 26 + ([int]$ch - 64)
    }
    return ($n - 1)
}

# Les files de la PRIMERA fulla del llibre. Cada fila es un objeto amb .Cells
# (array 0-based: cadenes o [double]); les cel.les buides, $null.
#
# Torna un array PLA d'objectes i s'ha de consumir amb @() (la convencio de
# rutes/: vegeu ConvertFrom-CatastroAdXml). Objectes i no arrays d'arrays a
# posta: una llista d'UNA sola fila desenrotllaria la fila sencera.
#
# Per NOM LOCAL (local-name()), com el parseig del Cadastre: cada programa que
# desa un .xlsx prefixa els espais de noms a la seva manera.
function Read-RepasXlsx([string]$path) {
    try { Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue | Out-Null } catch { }
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null } catch { }
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        # La fulla, la que digui el llibre: l'Excel pot haver canviat el nom
        # del fitxer intern.
        $fulla = 'xl/worksheets/sheet1.xml'
        $wb = _XlsxXml $zip 'xl/workbook.xml'
        $rels = _XlsxXml $zip 'xl/_rels/workbook.xml.rels'
        if ($null -ne $wb -and $null -ne $rels) {
            $s0 = $wb.SelectSingleNode("//*[local-name()='sheet']")
            if ($null -ne $s0) {
                $rid = $s0.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                foreach ($rel in @($rels.SelectNodes("//*[local-name()='Relationship']"))) {
                    if ($rel.GetAttribute('Id') -ne $rid) { continue }
                    $dest = $rel.GetAttribute('Target')
                    $fulla = if ($dest.StartsWith('/')) { $dest.Substring(1) } else { 'xl/' + $dest }
                }
            }
        }
        $compartits = New-Object System.Collections.Generic.List[string]
        $ss = _XlsxXml $zip 'xl/sharedStrings.xml'
        if ($null -ne $ss) {
            foreach ($si in @($ss.SelectNodes("//*[local-name()='si']"))) {
                $t = ''
                foreach ($tn in @($si.SelectNodes(".//*[local-name()='t']"))) { $t += $tn.InnerText }
                $compartits.Add($t)
            }
        }
        $doc = _XlsxXml $zip $fulla
        if ($null -eq $doc) { throw "El fitxer no te cap full de calcul." }

        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        $files = @()
        foreach ($row in @($doc.SelectNodes("//*[local-name()='row']"))) {
            $cells = New-Object System.Collections.ArrayList
            foreach ($c in @($row.SelectNodes("*[local-name()='c']"))) {
                # La columna, de la referencia: l'Excel no escriu les cel.les
                # buides i, sense aixo, tot es desplaçaria.
                $ref = $c.GetAttribute('r')
                $col = if ($ref -ne '') { _XlsxColumna $ref } else { $cells.Count }
                $tipus = $c.GetAttribute('t')
                $v = $c.SelectSingleNode("*[local-name()='v']")
                $val = $null
                if ($tipus -eq 'inlineStr') {
                    $val = ''
                    foreach ($tn in @($c.SelectNodes(".//*[local-name()='t']"))) { $val += $tn.InnerText }
                } elseif ($null -eq $v) {
                    $val = $null
                } elseif ($tipus -eq 's') {
                    $i = [int]$v.InnerText
                    $val = if ($i -ge 0 -and $i -lt $compartits.Count) { $compartits[$i] } else { '' }
                } elseif ($tipus -eq 'str' -or $tipus -eq 'b' -or $tipus -eq 'e') {
                    $val = $v.InnerText
                } else {
                    $d = 0.0
                    $val = if ([double]::TryParse($v.InnerText, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { $d } else { $v.InnerText }
                }
                while ($cells.Count -le $col) { [void]$cells.Add($null) }
                $cells[$col] = $val
            }
            $files += [pscustomobject]@{ Cells = $cells.ToArray() }
        }
        return @($files)
    } finally {
        $zip.Dispose()
    }
}

# ----------------------------------------------------------------------------
# QUE S'HA DE CANVIAR (PURA)
# ----------------------------------------------------------------------------

# L'ID Activitat d'una cel.la, com a text: un numero de l'Excel (101.0) -> '101',
# sense decimals, com a Ruta. PURA. La fan servir la lectura de la base
# (Coordenades.ps1) i aquest fitxer, que ha de casar el repas amb la base pel
# mateix ID: per aixo viu aqui, al modul de mes avall.
function Get-IdDeCella($cell) {
    if ($null -eq $cell) { return '' }
    if ($cell -is [double]) {
        if ([math]::Floor($cell) -eq $cell) { return [string][long]$cell }
        return [string]$cell
    }
    return ([string]$cell).Trim()
}


# Un numero d'una cel.la del repas: [double] tal qual, o text amb coma o punt
# decimal (si algu l'ha tocat a ma). $null si no n'es cap.
function _RepasNumero($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double]) { return [double]$v }
    return (ConvertTo-UtmNumber ([string]$v))
}

# Les correccions d'un repas per a la base $baseNom. Torna:
#   PerId       hashtable  ID GIA -> { XNova; YNova; XVella; YVella }
#   SenseCanvi  validades SENSE moure (la nova es la mateixa): no cal escriure-les
#   Invalides   files sense una coordenada nova que pugui ser d'aquest mon
#   AltraBase   files d'una ALTRA base de dades (la columna 'Base de dades')
#   Bases       quines
# Llança si el fitxer no es un repas (hi falten les columnes que calen).
function Get-CorreccionsDelRepas($files, [string]$baseNom) {
    $arr = @($files)
    if ($arr.Count -eq 0) { throw "El fitxer es buit." }
    $cap = @($arr[0].Cells | ForEach-Object { _NormalitzaText $_ })
    $col = @{}
    foreach ($k in @($Script:RepasColId, $Script:RepasColXNova, $Script:RepasColYNova,
                     $Script:RepasColXVella, $Script:RepasColYVella, $Script:RepasColBase)) {
        $col[$k] = [array]::IndexOf([object[]]$cap, $k)
    }
    if ($col[$Script:RepasColId] -lt 0 -or $col[$Script:RepasColXNova] -lt 0 -or $col[$Script:RepasColYNova] -lt 0) {
        throw "Aquest fitxer no es un repas de Coordenades: hi falten les columnes 'ID GIA', 'UTM X (nova)' i 'UTM Y (nova)'."
    }
    $cel = { param($f, $k) $i = $col[$k]; if ($i -lt 0 -or $i -ge @($f.Cells).Count) { $null } else { $f.Cells[$i] } }

    $perId = @{}
    $senseCanvi = 0; $invalides = 0; $altraBase = 0
    $bases = New-Object System.Collections.ArrayList
    for ($r = 1; $r -lt $arr.Count; $r++) {
        $f = $arr[$r]
        $id = Get-IdDeCella (& $cel $f $Script:RepasColId)
        if ($id -eq '') { continue }
        $base = ([string](& $cel $f $Script:RepasColBase)).Trim()
        if ($base -ne '' -and $baseNom -ne '' -and $base -ne $baseNom) {
            $altraBase++
            if (-not $bases.Contains($base)) { [void]$bases.Add($base) }
            continue
        }
        $xn = _RepasNumero (& $cel $f $Script:RepasColXNova)
        $yn = _RepasNumero (& $cel $f $Script:RepasColYNova)
        if ($null -eq $xn -or $null -eq $yn -or -not (Test-CoordPlausible $xn $yn)) { $invalides++; continue }
        $xv = _RepasNumero (& $cel $f $Script:RepasColXVella)
        $yv = _RepasNumero (& $cel $f $Script:RepasColYVella)
        # Validada sense moure (sense portal, o el portal ja era on deia l'Excel):
        # la coordenada es la mateixa i no s'ha de marcar en vermell.
        if ($null -ne $xv -and $null -ne $yv -and [math]::Abs($xn - $xv) -lt 0.005 -and [math]::Abs($yn - $yv) -lt 0.005) {
            $senseCanvi++; continue
        }
        $perId[$id] = [pscustomobject]@{ XNova = $xn; YNova = $yn; XVella = $xv; YVella = $yv }
    }
    return [pscustomobject]@{
        PerId = $perId; SenseCanvi = $senseCanvi; Invalides = $invalides
        AltraBase = $altraBase; Bases = @($bases)
    }
}

# El valor que s'escriu a la cel.la, AMB EL TIPUS QUE TENIA: si la base hi
# portava un numero, un numero; si hi portava text, text amb el mateix separador
# decimal. Sempre amb 2 decimals (centimetres), que es el que porta la base.
#
# El text va amb l'apostrof davant: si s'escrivis '421982,90' a pel per COM,
# l'Excel en catala/castella el convertiria en NUMERO (com si s'hagues teclejat)
# i la cel.la canviaria de tipus. L'apostrof no queda al valor: es el prefix
# que diu "aixo es text".
function Format-CoordComOriginal($original, [double]$nou) {
    $arrod = [math]::Round($nou, 2)
    if ($original -is [string]) {
        $t = $arrod.ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($original.Contains(',')) { $t = $t.Replace('.', ',') }
        return ("'" + $t)
    }
    return $arrod
}

# Quines cel.les s'han d'escriure, sobre la matriu de la fulla "Estes" (PURA:
# es prova amb una matriu feta a ma). Torna:
#   Escriptures  { Fila; Col; Valor } (dues per activitat: UTM X i UTM Y)
#   Aplicades    ids que s'escriuen
#   JaCanviades  ids la coordenada dels quals a la base JA NO es la que hi havia
#                en repassar-les (no es toquen: potser ja estan corregides)
#   NoTrobades   ids del repas que no son a la base
function Get-EscripturesCoordenades($data, [int]$rows, $headers, $perId) {
    $colId = Find-HeaderColumn $headers 'ID Activitat'
    $colX  = Find-HeaderColumn $headers 'UTM X'
    $colY  = Find-HeaderColumn $headers 'UTM Y'
    if ($colId -lt 1 -or $colX -lt 1 -or $colY -lt 1) {
        throw "La fulla 'Estes' no te les columnes 'ID Activitat', 'UTM X' i 'UTM Y'."
    }
    $escriptures = @(); $aplicades = @(); $jaCanviades = @()
    $vistos = @{}
    for ($r = 2; $r -le $rows; $r++) {
        $id = Get-IdDeCella $data[$r, $colId]
        if ($id -eq '' -or -not $perId.ContainsKey($id)) { continue }
        $vistos[$id] = $true
        $c = $perId[$id]
        $ox = $data[$r, $colX]; $oy = $data[$r, $colY]
        if ($null -ne $c.XVella -and $null -ne $c.YVella) {
            $bx = ConvertTo-UtmNumber $ox; $by = ConvertTo-UtmNumber $oy
            if ($null -eq $bx -or $null -eq $by -or
                [math]::Abs($bx - $c.XVella) -gt 0.01 -or [math]::Abs($by - $c.YVella) -gt 0.01) {
                $jaCanviades += $id; continue
            }
        }
        $escriptures += [pscustomobject]@{ Fila = $r; Col = $colX; Valor = (Format-CoordComOriginal $ox $c.XNova) }
        $escriptures += [pscustomobject]@{ Fila = $r; Col = $colY; Valor = (Format-CoordComOriginal $oy $c.YNova) }
        $aplicades += $id
    }
    $noTrobades = @(@($perId.Keys) | Where-Object { -not $vistos.ContainsKey($_) } | Sort-Object)
    return [pscustomobject]@{
        Escriptures = @($escriptures); Aplicades = @($aplicades)
        JaCanviades = @($jaCanviades); NoTrobades = @($noTrobades)
    }
}

# Nom de l'Excel per importar: el de la base, amb el mateix format (extensio)
# i un afegit que diu que porta les coordenades corregides i quan es va fer.
function Get-NomExcelImportar([string]$baseNom, [datetime]$ara) {
    $ext = [System.IO.Path]::GetExtension($baseNom)
    $arrel = [System.IO.Path]::GetFileNameWithoutExtension($baseNom)
    return ("{0} - coordenades corregides {1}{2}" -f $arrel, $ara.ToString('yyyy-MM-dd HHmm'), $ext)
}

# ----------------------------------------------------------------------------
# ESCRIURE-HO (COM, nomes a Windows amb Excel)
# ----------------------------------------------------------------------------
# $copia ha de ser la COPIA, mai la base original.
function Set-CoordenadesALaBase($copia, $perId) {
    $out = Read-FullaEstesa $copia {
        param($x)
        $pla = Get-EscripturesCoordenades $x.Data $x.Rows $x.Headers $perId
        foreach ($e in @($pla.Escriptures)) {
            $cel = $x.Sheet.Cells.Item($e.Fila, $e.Col)
            $cel.Value2 = $e.Valor
            $cel.Font.Color = $Script:CoordColorCorregida
        }
        return $pla
    } -Desa
    return $out
}
