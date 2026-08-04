#requires -Version 5.1
<#
.SYNOPSIS
  Base de dades d'ACTIVITATS: l'Excel de la feina, la seva cache i l'exportacio
  al Google Drive (per al formulari del mobil).

.DESCRIPTION
  Busca l'Excel "AAAA-MM-DD ACTIVITATS.xls(x)" mes recent, primer a la carpeta
  de la feina ($ActivitatsDir, unitat de xarxa) i, si no hi ha acces, a la copia
  local (local\base-dades-activitats\). En llegeix la fulla "Estes"/"Estès",
  en valida les capceleres i en munta una CACHE per ID GIA que despres omple
  sola la capcalera de l'informe (Pas 2).

  Tambe puja aquestes dades (nomes les que necessita el mobil) a la carpeta
  Dades del Drive, i nomes quan l'Excel local es mes nou que el que hi ha.

  Ve de Motor.ps1: son ~380 linies d'un concepte propi (Excel + Drive) que no
  tenien res a veure amb la resta del motor. Es dot-sourceja des de Motor.ps1,
  o sigui que comparteix ambit i el comportament no canvia.
#>

# ----------------------------------------------------------------------------
# Activitats Excel database - precarrega + validacio
# ----------------------------------------------------------------------------
# Mapeig de columnes Excel (1-based) per la fulla "Estes"/"Estès" del fitxer
# YYYY-MM-DD ACTIVITATS.xls. Es valida pel text de capcalera (fila 1); si no
# es troba el text esperat, es continua amb l'index per defecte pero
# s'afegeix un avis a $script:_activitatsWarnings.
$Script:ActivitatsColumns = @(
    @{ Key='ID';        Col=1;  HeaderHint='ID Activitat' }
    @{ Key='TITULAR';   Col=10; HeaderHint='Rao social' }
    @{ Key='TIPUS_VIA'; Col=48; HeaderHint='Tipus via' }
    @{ Key='CARRER';    Col=49; HeaderHint='Carrer' }
    @{ Key='NUMERO';    Col=50; HeaderHint='Numero' }
    @{ Key='LLETRA';    Col=52; HeaderHint='Lletra' }
    @{ Key='PIS';       Col=55; HeaderHint='Pis' }
    @{ Key='PORTA';     Col=56; HeaderHint='Porta' }
    @{ Key='ACTIVITAT'; Col=94; HeaderHint='Activitat principal' }
)

# Ruta de la base de dades LOCAL al clone. Si l'usuari executa el programa
# fora de la xarxa de la feina, pot copiar una "YYYY-MM-DD ACTIVITATS.xls"
# a aquesta carpeta i el programa la fara servir com a fallback. La carpeta
# es queda al clone (existeix amb un .gitkeep); els .xls/.xlsx de dins NO
# es pugen (estan al .gitignore).
$LocalActivitatsDir = Get-LocalSubdir $RepoRoot 'Activitats'

# Cerca el fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' mes recent en una carpeta.
# Retorna $null si no se'n troba cap.
function _FindLatestActivitatsIn($dir) {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $regex = '^(\d{4}-\d{2}-\d{2})\s+ACTIVITATS\.(xls|xlsx)$'
    $candidates = Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $regex } |
        ForEach-Object {
            if ($_.Name -match $regex) {
                [pscustomobject]@{
                    File = $_
                    Date = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $null)
                }
            }
        } | Sort-Object Date -Descending
    if ($candidates.Count -eq 0) { return $null }
    return $candidates[0]
}

# Cerca la base de dades en dues ubicacions, en ordre:
#   1. $ActivitatsDir  (xarxa de la feina)
#   2. $LocalActivitatsDir  (carpeta local del clone, fallback per a fora feina)
# Retorna un PSCustomObject amb:
#   File   : System.IO.FileInfo
#   Date   : data parsejada del nom
#   Source : 'primary' (xarxa) o 'fallback' (local del clone)
# Si no se'n troba a cap, retorna $null.
function Find-LatestActivitatsExcel {
    $r = _FindLatestActivitatsIn $ActivitatsDir
    if ($null -ne $r) {
        Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'primary' -Force
        return $r
    }
    $r = _FindLatestActivitatsIn $LocalActivitatsDir
    if ($null -ne $r) {
        Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'fallback' -Force
        return $r
    }
    return $null
}

# Normalitza un text Unicode (sense diacritics, minuscules) per a comparacio.
function _NormalizeText($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    return (($t -replace '\p{Mn}','').ToLower().Trim())
}

# Cerca l'index (1-based) de la columna a la fila de capcalera (fila 1) el text
# de la qual conte TOTS els termes de $mustContain i CAP dels de $mustNotContain
# (comparacio normalitzada: minuscules, sense accents). Retorna 0 si no es troba.
function _FindColIndex($data, $cols, [string[]]$mustContain, [string[]]$mustNotContain) {
    for ($c = 1; $c -le $cols; $c++) {
        $h = _NormalizeText $data[1, $c]
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $ok = $true
        foreach ($m in $mustContain) {
            if (-not $h.Contains((_NormalizeText $m))) { $ok = $false; break }
        }
        if ($ok -and $null -ne $mustNotContain) {
            foreach ($m in $mustNotContain) {
                if ($h.Contains((_NormalizeText $m))) { $ok = $false; break }
            }
        }
        if ($ok) { return $c }
    }
    return 0
}

# Converteix un valor de cel·la a text. Els enters d'Excel arriben com a double;
# els mostrem sense decimals ni notacio cientifica.
function _CellToString($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ([math]::Floor($v) -eq $v) { return [string][int64]$v }
        return [string]$v
    }
    return ([string]$v).Trim()
}

# Formata una cel·la de data a "dd/MM/yyyy" descartant l'hora. A l'Excel les
# dates arriben com a double (numero de serie OLE); tambe acceptem text.
function _FormatDateOnly($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        try { return ([DateTime]::FromOADate($v)).ToString('dd/MM/yyyy') } catch { return '' }
    }
    $s = ([string]$v).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt.ToString('dd/MM/yyyy') }
    return $s.Split(' ')[0]  # si no es pot parsejar, agafem la part abans de l'hora
}

# Text de la linia "Classificacio:" de la capcalera de LLICENCIA, a partir de les
# dues columnes de l'Excel: "Classificacio general annex" i "... Apartat".
#
#   ('II', '12.25')  ->  "Llei 20/2009; II; Epigraf 12.25"
#   ('II', '')       ->  "Llei 20/2009; II"
#   ('', '')         ->  ""      (no s'inventa una classificacio que no hi es)
#
# El prefix i el "Epigraf" surten del Word que feia servir l'usuari
# ("Llei 20/2009; Annex II; Epigraf 12.25"); a l'Excel l'annex ja hi consta com
# a "II" o "III", i el text es completa aqui. Funcio PURA.
function _ClassificacioText($annex, $apartat) {
    $a = ([string]$annex).Trim()
    $p = ([string]$apartat).Trim()
    if ([string]::IsNullOrWhiteSpace($a) -and [string]::IsNullOrWhiteSpace($p)) { return '' }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add('Llei 20/2009')
    if (-not [string]::IsNullOrWhiteSpace($a)) {
        # Si a l'Excel ja hi diu "Annex II", no s'ha de repetir la paraula.
        if ($a -match '(?i)^annex\b') { [void]$parts.Add($a) } else { [void]$parts.Add('Annex ' + $a) }
    }
    if (-not [string]::IsNullOrWhiteSpace($p)) {
        if ($p -match '(?i)^ep' + [char]0x00ED + 'graf\b') { [void]$parts.Add($p) }
        else { [void]$parts.Add('Ep' + [char]0x00ED + 'graf ' + $p) }
    }
    return ($parts -join '; ')
}

# Localitza la fulla "Estes"/"Estès" del workbook acceptant variants Unicode.
function _FindEstesSheet($wb) {
    $sheetNames = @()
    foreach ($s in $wb.Sheets) {
        $sheetNames += $s.Name
        if ((_NormalizeText $s.Name) -eq 'estes') { return @{ Sheet=$s; Names=$sheetNames } }
    }
    return @{ Sheet=$null; Names=$sheetNames }
}

# Valida la fila de capcalera comparant els textos esperats. Retorna una
# llista (potser buida) d'avisos en text per mostrar a l'usuari.
function _ValidateActivitatsHeaders($data, $rows, $cols) {
    $warnings = New-Object System.Collections.ArrayList
    if ($rows -lt 1) { return $warnings }
    foreach ($col in $Script:ActivitatsColumns) {
        $idx  = $col.Col
        if ($idx -lt 1 -or $idx -gt $cols) {
            [void]$warnings.Add("Columna $idx fora de rang per a '$($col.Key)' (Excel te $cols columnes).")
            continue
        }
        $cell = $data[1, $idx]
        $cellN = _NormalizeText $cell
        $hintN = _NormalizeText $col.HeaderHint
        if ([string]::IsNullOrWhiteSpace($cellN) -or -not $cellN.Contains($hintN.Split(' ')[0])) {
            [void]$warnings.Add("La columna $idx esperava '$($col.HeaderHint)' pero te '$cell'.")
        }
    }
    return $warnings
}

# Precarrega TOTES les activitats de l'Excel en una hashtable indexada per ID.
# Es crida una sola vegada al comencar el Pas 2; despres les cerques son
# instantanies (no calen mes obertures d'Excel encara que l'usuari premi
# "Cercar" diverses vegades).
#
# Retorna un PSCustomObject amb:
#   File        : System.IO.FileInfo del fitxer Excel
#   Date        : data del fitxer (parsejada del nom)
#   ById        : hashtable [string ID] -> hashtable amb TITULAR, ADRECA,
#                 ACTIVITAT, EXP_NUM, NUM_ANOTACIO, DATA_ANOTACIO
#   Warnings    : llista de cadenes amb avisos de validacio de columnes
function Initialize-ActivitatsCache($excelFile) {
    # Igual que amb Word: si Excel no esta disponible, New-Object pot fallar o
    # retornar $null. Donem un missatge clar (aquest error es propaga a
    # Get-HeaderData, que ja mostra "Error llegint l'Excel").
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) {
        throw "No s'ha pogut iniciar Microsoft Excel. Comprova que estigui instal-lat i obert almenys un cop."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            $found = _FindEstesSheet $wb
            $sh = $found.Sheet
            if ($null -eq $sh) {
                throw "No s'ha trobat la fulla 'Estes'/'Estès' al fitxer Excel. Fulles disponibles: $($found.Names -join ', ')"
            }
            $used = $sh.UsedRange
            $data = $used.Value2
            if ($null -eq $data) {
                return [pscustomobject]@{ ById = @{}; Warnings = @("L'Excel sembla buit.") }
            }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

            # Avisos de validacio. Construim un ArrayList REAL aqui (no podem
            # assignar directament el retorn de _ValidateActivitatsHeaders: en
            # retornar un ArrayList, PowerShell el desempaqueta i, si es buit,
            # $warnings quedaria $null i $warnings.Add()/.ToArray() petarien).
            $warnings = New-Object System.Collections.ArrayList
            foreach ($w in (_ValidateActivitatsHeaders $data $rows $cols)) {
                if ($null -ne $w) { [void]$warnings.Add($w) }
            }

            # Columnes localitzades pel TEXT de la capcalera (mes robust que un
            # index fix). Si l'Excel canvia l'ordre de columnes, segueix
            # funcionant mentre el nom es mantingui.
            $colExp  = _FindColIndex $data $cols @('expedient') $null
            $colNum  = _FindColIndex $data $cols @('registre','entrada') @('data')
            $colData = _FindColIndex $data $cols @('data','registre','entrada') $null
            # Classificacio de l'activitat, per a la capcalera de Llicencia.
            $colAnx  = _FindColIndex $data $cols @('classificacio general annex') $null
            $colApa  = _FindColIndex $data $cols @('classificacio general apartat') $null
            if ($colExp  -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Num. expedient'.") }
            if ($colNum  -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Num. registre entrada'.") }
            if ($colData -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Data registre entrada'.") }

            # Index per ID (columna 1).
            $byId = @{}
            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, 1]
                if ($null -eq $cell) { continue }
                $id = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                if ([string]::IsNullOrWhiteSpace($id)) { continue }

                $tipusVia = & $get $r 48
                $carrer   = & $get $r 49
                $numero   = & $get $r 50
                $lletra   = & $get $r 52
                $pis      = & $get $r 55
                $porta    = & $get $r 56
                $rao      = & $get $r 10
                $raoMobil = & $get $r 23
                $raoEmail = & $get $r 25
                $actPrin  = & $get $r 94
                $parts = @($tipusVia, $carrer, $numero, $lletra, $pis, $porta) |
                    Where-Object { $_ -and $_.Trim() -ne '' }
                # Construim l'accent amb el codepoint Unicode explicit (U+00C0,
                # 'A' amb accent greu) per evitar que la lletra accentuada del
                # literal es corrompi segons l'encoding amb que PowerShell 5.1
                # llegeix aquest fitxer (sortia "CORNELLÃ€").
                $ciutat = "CORNELL$([char]0x00C0) DE LLOBREGAT"
                $adreca = ($parts -join ' ') + ", $ciutat"

                $expNum = if ($colExp  -gt 0) { _CellToString  $data[$r, $colExp] }  else { '' }
                $numAno = if ($colNum  -gt 0) { _CellToString  $data[$r, $colNum] }  else { '' }
                $datAno = if ($colData -gt 0) { _FormatDateOnly $data[$r, $colData] } else { '' }
                $anx    = if ($colAnx  -gt 0) { _CellToString  $data[$r, $colAnx] }  else { '' }
                $apa    = if ($colApa  -gt 0) { _CellToString  $data[$r, $colApa] }  else { '' }

                $byId[$id] = @{
                    TITULAR       = $rao
                    MOBIL         = $raoMobil
                    EMAIL         = $raoEmail
                    ADRECA        = $adreca
                    ACTIVITAT     = $actPrin
                    EXP_NUM       = $expNum
                    NUM_ANOTACIO  = $numAno
                    DATA_ANOTACIO = $datAno
                    CLASSIFICACIO = (_ClassificacioText $anx $apa)
                }
            }
            return [pscustomobject]@{ ById = $byId; Warnings = $warnings.ToArray() }
        } finally {
            $wb.Close($false)
        }
    } finally {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# Cerca una activitat per ID al cache precarregat. Retorna $null si no es troba.
function Get-ActivitatFromCache($cache, $idGia) {
    if ($null -eq $cache -or $null -eq $cache.ById) { return $null }
    $key = [string]$idGia
    if ($cache.ById.ContainsKey($key)) { return $cache.ById[$key] }
    return $null
}

# Treu la data (yyyy-MM-dd) d'un objecte activitats.json ja existent: prefereix
# el camp SourceDate i, si no hi es (versions antigues), la treu del nom Source.
# Retorna [datetime]::MinValue si no en pot deduir cap.
function _ParseActivitatsDate($obj) {
    if ($null -eq $obj) { return [datetime]::MinValue }
    $d = [datetime]::MinValue
    if ($obj.SourceDate -and [datetime]::TryParseExact([string]$obj.SourceDate, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d
    }
    if ($obj.Source -and ([string]$obj.Source) -match '(\d{4}-\d{2}-\d{2})') {
        if ([datetime]::TryParseExact($matches[1], 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    }
    return [datetime]::MinValue
}

# Decideix si cal exportar la base d'activitats al Drive. Nomes cal si la base
# LOCAL es MES NOVA que la que ja hi ha al Drive. Si tenen la mateixa data (o
# la del Drive es mes nova), NO cal: ens estalviem obrir l'Excel i pujar un
# fitxer identic. Funcio pura (provable).
#   $localDate    : data del fitxer Excel local (del nom YYYY-MM-DD).
#   $existingDate : data de l'activitats.json que ja hi ha al Drive.
function Test-ShouldExportActivitats([datetime]$localDate, [datetime]$existingDate) {
    # Si no tenim una data local fiable, exportem (no podem decidir res).
    if ($localDate -le [datetime]::MinValue) { return $true }
    return ($localDate -gt $existingDate)
}

# Llegeix la data (SourceDate) de l'activitats.json que ja hi ha al Drive,
# SENSE obrir l'Excel. Funciona tant en mode API com en mode carpeta local
# sincronitzada. Retorna [datetime]::MinValue si no n'hi ha o no es pot llegir
# (es fail-safe: qualsevol error -> MinValue, que fa que s'exporti).
function Get-DriveActivitatsDate {
    try {
        if (Test-DriveApiConfigured) {
            if (-not $DriveDadesId) { return [datetime]::MinValue }
            $existingId = Find-DriveFileId 'activitats.json' $DriveDadesId
            if (-not $existingId) { return [datetime]::MinValue }
            $existing = (Get-DriveFileText $existingId) | ConvertFrom-Json
            return _ParseActivitatsDate $existing
        }
        $outFile = Join-Path $DriveDadesDir 'activitats.json'
        if (Test-Path -LiteralPath $outFile) {
            $existing = (Get-Content -LiteralPath $outFile -Raw -Encoding UTF8) | ConvertFrom-Json
            return _ParseActivitatsDate $existing
        }
        return [datetime]::MinValue
    } catch {
        return [datetime]::MinValue
    }
}

# Exporta la base de dades d'activitats (nomes els camps de capcalera, per ID
# GIA) a un JSON dins la carpeta PRIVADA de Drive, perque el mobil pugui
# auto-emplenar la capcalera. Aquestes dades son personals i NO van mai al
# GitHub public: nomes a Drive (compte privat de l'usuari). Es fail-safe:
# qualsevol error es registra i es retorna $false sense interrompre el flux.
#
# IMPORTANT: NO sobreescriu la base del Drive si la que ja hi ha es MES NOVA
# (pujada des d'un altre PC). Compara per data del nom del fitxer Excel.
function Export-ActivitatsToDrive($cache, $latest) {
    try {
        if ($null -eq $cache -or $null -eq $cache.ById) { return $false }
        $localDate = if ($latest -and $latest.Date) { $latest.Date } else { [datetime]::MinValue }

        # No tornar a exportar si el Drive ja te una base amb la MATEIXA data
        # (o mes nova, pujada des d'un altre PC). Aixi no es sobreescriu una
        # versio mes nova ni es perd temps pujant una d'identica. La data del
        # Drive es llegeix del propi activitats.json (camp SourceDate).
        $existingDate = Get-DriveActivitatsDate
        if (-not (Test-ShouldExportActivitats $localDate $existingDate)) {
            Write-Host ("  El Drive ja esta al dia ({0}); no cal tornar a exportar (local {1})." -f $existingDate.ToString('yyyy-MM-dd'), $localDate.ToString('yyyy-MM-dd'))
            return $true
        }

        $payload = [ordered]@{
            GeneratedAt = (Get-Date).ToString('o')
            Source      = if ($latest) { $latest.File.Name } else { '' }
            SourceDate  = $localDate.ToString('yyyy-MM-dd')
            Count       = $cache.ById.Count
            ById        = $cache.ById
        }
        $json = ($payload | ConvertTo-Json -Depth 6)

        # Mode API (sense Drive d'escriptori): pugem activitats.json directament
        # a la carpeta Dades de Drive. Si no hi ha credencials, caiem al mode de
        # carpeta local sincronitzada.
        if (Test-DriveApiConfigured) {
            if (-not $DriveDadesId) {
                Write-Host "Avis: hi ha credencials de Drive pero falta \$DriveDadesId a config.ps1. No s'exporten activitats."
                return $false
            }
            Save-DriveJson 'activitats.json' $DriveDadesId $json | Out-Null
            return $true
        }

        if (-not (Test-Path -LiteralPath $DriveDadesDir)) {
            New-Item -ItemType Directory -Path $DriveDadesDir -Force | Out-Null
        }
        $outFile = Join-Path $DriveDadesDir 'activitats.json'
        $json | Set-Content -LiteralPath $outFile -Encoding UTF8
        return $true
    } catch {
        Write-Host "Avis: no s'ha pogut exportar les activitats a Drive ($($_.Exception.Message))."
        return $false
    }
}
