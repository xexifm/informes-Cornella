<#
  Precintades.ps1 - Generador de dades del PLANOL PUBLIC d'activitats precintades.

  Que fa:
    1. Localitza el fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' mes recent (xarxa de
       la feina o carpeta local de fallback), fulla "Estes"/"Estès".
    2. Detecta, de forma dinamica llegint la CAPCALERA, els parells de camps
       lliures "Camp Info N - Nom" / "Camp Info N - Valor" (n'hi pot haver mes
       de 3) i les columnes fixes que necessitem (ID Activitat, UTM X/Y,
       adreca de l'emplacament i activitat principal).
    3. Selecciona les activitats PRECINTADES: aquelles amb un camp lliure amb
       Nom = "PRECINTE ACTIVITAT?" i Valor que comenca per "SI".
    4. De cada una identifica: ID Activitat (col. A), activitat principal
       (col. CP) i adreca (Emp. Tipus via + Carrer + Numero, col. AV+AW+AX).
    5. Les geolocalitza (UTM ETRS89/31N -> lat/lon WGS84) i escriu
       docs/dades/precintades.json, que la pagina publica docs/precintades.html
       (GitHub Pages) llegeix per pintar el mapa.

  El JSON NO conte cap dada personal (ni rao social ni notes internes del Valor):
  nomes activitat generica, adreca de l'establiment, ID intern i coordenades.

  Reutilitza les funcions ja provades de Ruta.ps1 (conversio UTM, format
  d'adreca, cerca de l'Excel...) carregant-lo en mode headless.

  Mode "headless" per a proves: si $env:PRECINTADES_TEST o $env:GENINFORME_TEST
  estan definides, NOMES es defineixen les funcions (no es llegeix cap Excel ni
  s'escriu res). Aixi es poden provar les funcions pures a Linux sense Office.
#>

$ErrorActionPreference = 'Stop'

# Headless: nomes definir funcions (proves). Compartim el flag amb Ruta/Motor.
$Script:PrecHeadless = [bool]$env:PRECINTADES_TEST -or [bool]$env:GENINFORME_TEST

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
# Aquest script viu a suport/rutes/. Ruta.ps1 es al costat; l'arrel del clone
# es dos nivells amunt (suport/rutes/../..).
$SuportDir  = Split-Path -Parent $ScriptRoot          # suport/
$RepoRoot   = Split-Path -Parent $SuportDir           # informes-Cornella/

# ----------------------------------------------------------------------------
# Reutilitzem les funcions de Ruta.ps1 (Convert-UtmToLatLon, Format-EmpAddress,
# ConvertTo-UtmNumber, _RutaNormalize, Find-LatestRutaExcel, _RutaFindEstesSheet).
# El carreguem en mode headless (RUTA_TEST) perque NOMES defineixi funcions i
# no obri la seva finestra ni executi la seva Main. Restaurem la variable
# d'entorn despres de carregar-lo per no afectar la resta del proces.
# ----------------------------------------------------------------------------
$Script:_prevRutaTest = $env:RUTA_TEST
$env:RUTA_TEST = '1'
try {
    . (Join-Path $ScriptRoot 'Ruta.ps1')
} finally {
    if ($null -eq $Script:_prevRutaTest) {
        Remove-Item Env:\RUTA_TEST -ErrorAction SilentlyContinue
    } else {
        $env:RUTA_TEST = $Script:_prevRutaTest
    }
}

# Carpeta de sortida: docs/dades (GitHub Pages la serveix des de /docs).
$WebDadesDir = Join-Path $RepoRoot (Join-Path 'docs' 'dades')

# Nom EXACTE (normalitzat) del camp lliure que marca una activitat precintada.
$Script:PrecCampNom = 'precinte activitat?'

# ============================================================================
# FUNCIONS PURES (provables en mode headless, sense Office)
# ============================================================================

# (Find-HeaderColumn vivia aqui i ara es a Ruta.ps1, amb la resta d'utillatge
# comu de la carpeta 'rutes': la fa servir tambe l'eina Coordenades. Ruta.ps1
# ja es carrega mes amunt, aixi que segueix a l'abast des d'aqui.)

# Detecta tots els parells de camps lliures "Camp Info N - Nom" /
# "Camp Info N - Valor" a la capcalera, de forma dinamica (n'hi pot haver mes
# de 3). Retorna un array de hashtables @{ NomCol = <1-based>; ValorCol = <1-based> }.
function Get-CampInfoPairs($headers) {
    $pairs = @()
    for ($i = 0; $i -lt @($headers).Count; $i++) {
        $h = _RutaNormalize $headers[$i]
        if ($h -match '^camp info\s+(\d+)\s*-\s*nom$') {
            $n = $Matches[1]
            $valorCol = Find-HeaderColumn $headers ("Camp Info $n - Valor")
            if ($valorCol -gt 0) {
                $pairs += @{ NomCol = ($i + 1); ValorCol = $valorCol }
            }
        }
    }
    return ,@($pairs)
}

# Cert si un camp lliure (Nom/Valor) indica una activitat PRECINTADA: el Nom es
# "PRECINTE ACTIVITAT?" i el Valor comenca per "SI" (com a paraula: "SI",
# "SI, ...", "SI ..."; NO "SITUACIO..."). Insensible a accents i majuscules.
function Test-IsPrecintada([string]$nom, [string]$valor) {
    if ((_RutaNormalize $nom) -ne $Script:PrecCampNom) { return $false }
    $v = _RutaNormalize $valor
    if ($v -eq '') { return $false }
    return [bool]($v -match '^si\b')
}

# Construeix l'objecte que es serialitzara a precintades.json a partir dels
# registres (cada un amb Id, ActivitatPrincipal, Adreca, Lat, Lon).
function Build-PrecintadesObject($records, [string]$fontName) {
    $acts = @($records | ForEach-Object {
        [pscustomobject]@{
            id        = [string]$_.Id
            activitat = [string]$_.ActivitatPrincipal
            adreca    = [string]$_.Adreca
            lat       = [double]$_.Lat
            lon       = [double]$_.Lon
        }
    })
    # SourceDate: data (yyyy-MM-dd) de l'Excel origen, treta del seu nom. Serveix
    # per no fer enrere el mapa si despres es genera des d'una base mes antiga.
    $srcDate = ''
    if (([string]$fontName) -match '(\d{4}-\d{2}-\d{2})') { $srcDate = $Matches[1] }
    return [pscustomobject]@{
        GeneratedAt = (Get-Date).ToString('o')
        Font        = [string]$fontName
        SourceDate  = $srcDate
        Comptador   = @($acts).Count
        Activitats  = $acts
    }
}

# Treu la data (yyyy-MM-dd) d'un precintades.json ja existent: prefereix el camp
# SourceDate i, si no hi es (versions antigues), la treu del nom Font. Retorna
# [datetime]::MinValue si no en pot deduir cap.
function _ParsePrecintadesDate($obj) {
    if ($null -eq $obj) { return [datetime]::MinValue }
    $d = [datetime]::MinValue
    if ($obj.SourceDate -and [datetime]::TryParseExact([string]$obj.SourceDate, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d
    }
    if ($obj.Font -and ([string]$obj.Font) -match '(\d{4}-\d{2}-\d{2})') {
        if ([datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    }
    return [datetime]::MinValue
}

# Decideix si cal (re)generar el mapa de precintades. Nomes si la base d'Excel
# trobada es MES NOVA que la del precintades.json que ja hi ha. Si tenen la
# mateixa data (o la publicada es mes nova), NO s'actualitza: aixi, si generem
# des d'una base local antiga (fora de la feina), no fem enrere el mapa public.
# Mateixa logica que Test-ShouldExportActivitats (Drive). Funcio pura.
function Test-ShouldUpdatePrecintades([datetime]$localDate, [datetime]$existingDate) {
    if ($localDate -le [datetime]::MinValue) { return $true }
    return ($localDate -gt $existingDate)
}

# ============================================================================
# LECTURA D'EXCEL (COM) - nomes a Windows amb Excel; no es prova en headless.
# ============================================================================

# Llegeix la fulla "Estes" i retorna els registres de les activitats
# PRECINTADES (array d'objectes {Id, ActivitatPrincipal, Adreca, Lat, Lon}).
function Read-PrecintadesFromExcel($excelFile) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            $sh = _RutaFindEstesSheet $wb
            if ($null -eq $sh) { throw "No s'ha trobat la fulla 'Estes'/'Estès' al fitxer Excel." }
            $data = $sh.UsedRange.Value2
            if ($null -eq $data) { return @() }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

            # Capcalera (fila 1) -> array 0-based de noms de columna.
            $headers = @()
            for ($c = 1; $c -le $cols; $c++) {
                $hv = $data[1, $c]
                $headers += $(if ($null -eq $hv) { '' } else { ([string]$hv).Trim() })
            }

            # Columnes fixes (per NOM, mes robust que per index). IMPORTANT: els
            # noms de cerca s'escriuen en ASCII SENSE accents ('Numero', no
            # 'Número'). El Windows PowerShell 5.1 llegeix els .ps1 sense BOM com
            # a ANSI i corromp els literals accentuats; a mes, Find-HeaderColumn
            # normalitza sense diacritics, aixi que 'Emp. Numero' encaixa amb la
            # capcalera real 'Emp. Número'. (Mateixa convencio que Ruta.ps1, que
            # compara contra 'estes' i no 'Estès'.)
            $colId   = Find-HeaderColumn $headers 'ID Activitat'
            $colUtmX = Find-HeaderColumn $headers 'UTM X'
            $colUtmY = Find-HeaderColumn $headers 'UTM Y'
            $colVia  = Find-HeaderColumn $headers 'Emp. Tipus via'
            $colCarr = Find-HeaderColumn $headers 'Emp. Carrer'
            $colNum  = Find-HeaderColumn $headers 'Emp. Numero'
            $colAct  = Find-HeaderColumn $headers 'Activitat principal'
            $pairs   = Get-CampInfoPairs $headers

            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }

            $records = @()
            for ($r = 2; $r -le $rows; $r++) {
                # Es precintada si ALGUN parell Camp Info ho indica.
                $isPrec = $false
                foreach ($p in $pairs) {
                    $nom   = & $get $r $p.NomCol
                    $valor = & $get $r $p.ValorCol
                    if (Test-IsPrecintada $nom $valor) { $isPrec = $true; break }
                }
                if (-not $isPrec) { continue }

                # ID Activitat (numero -> enter sense decimals, com a Ruta).
                $idCell = if ($colId -ge 1 -and $colId -le $cols) { $data[$r, $colId] } else { $null }
                $id = if ($idCell -is [double]) {
                    if ([math]::Floor($idCell) -eq $idCell) { [string][int]$idCell } else { [string]$idCell }
                } elseif ($null -ne $idCell) { ([string]$idCell).Trim() } else { '' }

                $x = ConvertTo-UtmNumber (& $get $r $colUtmX)
                $y = ConvertTo-UtmNumber (& $get $r $colUtmY)
                if ($null -eq $x -or $null -eq $y) { continue }   # sense coordenades: no es pot situar
                $ll = Convert-UtmToLatLon $x $y 31 $true

                $records += [pscustomobject]@{
                    Id                = $id
                    ActivitatPrincipal = (& $get $r $colAct)
                    Adreca            = (Format-EmpAddress (& $get $r $colVia) (& $get $r $colCarr) (& $get $r $colNum) '')
                    Lat               = $ll.Lat
                    Lon               = $ll.Lon
                }
            }
            return ,@($records)
        } finally {
            $wb.Close($false)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# ============================================================================
# MAIN
# ============================================================================
function Invoke-PrecintadesMain {
    $xls = Find-LatestRutaExcel
    if ($null -eq $xls) {
        Write-Host "No s'ha trobat cap base de dades d'activitats; ometo el mapa de precintades."
        return $false
    }

    # No fer enrere el mapa: si el precintades.json que ja hi ha ve d'una base
    # IGUAL o MES NOVA que la trobada ara, no el regenerem (mateixa idea que amb
    # el Drive). Aixi, executant Actualitzar fora de la feina amb una base local
    # antiga, el mapa public no es degrada.
    $outPath = Join-Path $WebDadesDir 'precintades.json'
    $existingDate = [datetime]::MinValue
    if (Test-Path -LiteralPath $outPath) {
        try { $existingDate = _ParsePrecintadesDate ((Get-Content -LiteralPath $outPath -Raw -Encoding UTF8) | ConvertFrom-Json) } catch { $existingDate = [datetime]::MinValue }
    }
    if (-not (Test-ShouldUpdatePrecintades $xls.Date $existingDate)) {
        Write-Host ("El mapa de precintades ja ve d'una base igual o mes nova ({0:yyyy-MM-dd}); no l'actualitzo." -f $existingDate)
        return $true
    }

    Write-Host "Llegint activitats precintades de: $($xls.File.Name)"
    $records = Read-PrecintadesFromExcel $xls.File
    $obj = Build-PrecintadesObject $records $xls.File.Name

    if (-not (Test-Path -LiteralPath $WebDadesDir)) {
        New-Item -ItemType Directory -Path $WebDadesDir -Force | Out-Null
    }
    $json = ($obj | ConvertTo-Json -Depth 6)
    [System.IO.File]::WriteAllText($outPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "  precintades.json -> $outPath ($($obj.Comptador) activitats precintades)"
    return $true
}

if (-not $Script:PrecHeadless) {
    Invoke-PrecintadesMain | Out-Null
}
