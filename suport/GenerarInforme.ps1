#requires -Version 5.1
<#
.SYNOPSIS
  Generador d'informes de l'Ajuntament de Cornella.

.DESCRIPTION
  Flux:
    1. L'usuari escull un cataleg de defciencies (ESTRUCTURALS\REQ*.docx).
    2. Demana les dades de la capcalera (ID GIA, EXP_NUM, etc.).
    3. Mostra un TreeView amb les seccions/items del cataleg; l'usuari marca
       quines defciencies aplicaran a l'informe.
    4. Per cada placeholder [CAMP: ...] que apareixi en algun item seleccionat,
       demana el valor (1 cop per nom de camp).
    5. Mostra un formulari amb les conclusions disponibles (paragrafs de
       ESTRUCTURALS\0 CONCLUSIONS.docx) i deixa triar-ne una o mes.
    6. Composa el document final:
         - Capcalera (ESTRUCTURALS\0 CAPCALERA.docx) amb valors substituits.
         - Per cada seccio escollida: titol Heading 1 + items numerats
           globalment 1..N.
         - Conclusions seleccionades.

.NOTES
  Configuracio: les rutes i constants es defineixen al fitxer config.ps1
  (opcional) al costat del .ps1; si no existeix, s'usen els valors per
  defecte definits a sota.

  Persistencia: despres de cada pas, l'estat es guarda a
  %LOCALAPPDATA%\InformesCornella\session.json. Si en arrencar es
  detecta una sessio anterior, el script pregunta si es vol recuperar.

  Cache: el resultat del parseig del cataleg .docx es guarda a
  %LOCALAPPDATA%\InformesCornella\cache\<basename>.json amb un hash
  del fitxer com a clau de validesa.

  Convencions del cataleg (REQ1.docx i seguents):
    - Heading 1  -> titol de seccio.
    - Heading 2  -> nom curt de l'item (per al TreeView). Si comenca per
                    "::CHILD:: " es tracta d'un sub-bullet (fill de l'item
                    Heading 2 anterior); el prefix es elimina abans de mostrar.
    - Normal     -> cos de l'item: la primera linia es el text principal, les
                    seguents son URLs o complements (es mantenen tal qual).

  Placeholders al cos:
    [CAMP: nom]                 -> demana 'nom'
    [CAMP: nom (hint d'ajuda)]  -> demana 'nom', el hint apareix sota el camp
    Mateix nom = mateix valor (es demana un sol cop).
#>

$ErrorActionPreference = 'Stop'

# Mode "headless" per a proves automatiques: si la variable d'entorn
# GENINFORME_TEST esta definida, NO carreguem WinForms ni executem el
# programa (Main). Aixo permet carregar (dot-source) el script en un Linux
# sense Windows/Office per provar les funcions pures. En us normal (sense la
# variable) el comportament es identic al d'abans.
$Script:HeadlessTest = [bool]$env:GENINFORME_TEST
if (-not $Script:HeadlessTest) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

$ScriptRoot      = Split-Path -Parent $MyInvocation.MyCommand.Path
# Arrel del clone (un nivell amunt: suport/.. = informes-Cornella/).
# Aixi pots moure la carpeta informes-Cornella on vulguis i tot segueix
# funcionant; nomes la base de dades d'activitats (ActivitatsDir) es una
# ruta absoluta externa que no es mou.
$RepoRoot        = Split-Path -Parent $ScriptRoot

# Carreguem el modul de format (Format.ps1). Conte les funcions Format-Section,
# Format-Item, etc. i $ReportFormatConfig. Reutilitzable per altres tipus
# d'informes.
. (Join-Path $ScriptRoot 'Format.ps1')

# Carreguem el modul de seguiment (Seguiment.ps1). Conte el mode "Informe de
# seguiment" (afegir anotacions de resolucio sobre un informe anterior). Es
# carrega tambe en mode headless perque les seves funcions pures es puguin
# provar des dels tests.
. (Join-Path $ScriptRoot 'Seguiment.ps1')

# ESTRUCTURALS viu a l'arrel del clone (al costat dels .bat), no dins
# de suport/, per facilitar que l'usuari edita les plantilles.
$EstructuralsDir = Join-Path $RepoRoot 'ESTRUCTURALS'
$HeaderPath      = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
$ConclusionsPath = Join-Path $EstructuralsDir '0 CONCLUSIONS.docx'

# ----------------------------------------------------------------------------
# Configuracio per defecte. Es pot sobreescriure des de config.ps1 (opcional)
# al costat del .ps1 (dins de suport/).
# ----------------------------------------------------------------------------
# OutputDir per defecte: 'Informes generats' AL COSTAT DEL CLONE (no dins
# de suport/), per quedar al nivell que l'usuari veu i pot obrir amb el
# Word sense entrar a carpetes internes. El .gitignore l'exclou.
$OutputDir              = Join-Path $RepoRoot 'Informes generats'
$ActivitatsDir          = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
$AlwaysConclusionsCount = 2

# Mode "Informe de seguiment": frases que marquen l'inici del bloc de
# conclusions a esborrar de l'informe anterior. La deteccio es insensible a
# accents/majuscules. Es pot sobreescriure des de config.ps1.
$SeguimentConclusionPhrases = @(
    "Vist l'anterior",
    'Ho poso al seu coneixement',
    'Cornella de Llobregat,'
)

$configPath = Join-Path $ScriptRoot 'config.ps1'
if (Test-Path -LiteralPath $configPath) {
    . $configPath
}

# Estat persistent. Es guarda a %LOCALAPPDATA% per no embrutar el repositori
# i no haver de tocar .gitignore. lastreport.json conte les dades de l'ULTIM
# informe generat amb exit, per poder-lo replicar des del Pas 2.
$AppDataDir     = Join-Path $env:LOCALAPPDATA 'InformesCornella'
$LastReportPath = Join-Path $AppDataDir 'lastreport.json'

function Ensure-AppDataDir {
    if (-not (Test-Path -LiteralPath $AppDataDir)) {
        New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null
    }
}

# ----------------------------------------------------------------------------
# Word COM helpers
# ----------------------------------------------------------------------------
function New-WordApp {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0  # wdAlertsNone
    # Evita que Word obri els fitxers de xarxa en "Vista protegida", que
    # bloqueja InsertParagraphAfter i altres operacions de modificacio.
    try { $w.AutomationSecurity = 1 } catch { }  # msoAutomationSecurityLow
    return $w
}

function Close-WordApp($word) {
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ----------------------------------------------------------------------------
# Persistencia de l'ultim informe (per replicar-lo)
# ----------------------------------------------------------------------------
# Format del lastreport.json (versio 1):
#   {
#     "Version": 1,
#     "Timestamp": "<ISO 8601>",
#     "CatalegBaseName": "REQ1",
#     "Header": { "ID_GIA": "...", ... },
#     "SelectedKeys": [ "SectionTitle::ItemShort", ... ],
#     "FieldValues":  { "nom": "valor", ... },
#     "ConclusionTexts": [ "text1", ... ]
#   }

function Save-LastReport($state) {
    try {
        Ensure-AppDataDir
        $state.Version   = 1
        $state.Timestamp = (Get-Date).ToString('o')
        ($state | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $LastReportPath -Encoding UTF8
    } catch {
        # Si no podem desar, no es un error fatal. Continuem en silenci.
    }
}

function Load-LastReport {
    if (-not (Test-Path -LiteralPath $LastReportPath)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $LastReportPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

# Construeix la clau "Seccio::Item" o "Seccio::Item::Fill" usada per
# identificar de manera unica un element seleccionat al Pas 3.
function _ItemKey($sectionTitle, $itemShort, $childShort = $null) {
    if ($childShort) { return "$sectionTitle::$itemShort::$childShort" }
    return "$sectionTitle::$itemShort"
}

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
$LocalActivitatsDir = Join-Path $RepoRoot 'BASE DE DADES ACTIVITATS'

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
    $excel = New-Object -ComObject Excel.Application
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

                $byId[$id] = @{
                    TITULAR       = $rao
                    ADRECA        = $adreca
                    ACTIVITAT     = $actPrin
                    EXP_NUM       = $expNum
                    NUM_ANOTACIO  = $numAno
                    DATA_ANOTACIO = $datAno
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

# ----------------------------------------------------------------------------
# Step 1 - Cataleg picker
# ----------------------------------------------------------------------------
function Get-Catalegs {
    # Tots els .docx d'ESTRUCTURALS\ que NO comencin amb "0 " son catalegs.
    Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.docx' |
        Where-Object { $_.Name -notlike '0 *' -and $_.Name -notlike '0_*' -and -not $_.Name.StartsWith('~$') } |
        Sort-Object Name
}

function Select-Cataleg {
    param($preloadBaseName = $null)
    $catalegs = @(Get-Catalegs)
    if ($catalegs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap cataleg (REQ*.docx) a $EstructuralsDir.",
            'Error', 'OK', 'Error') | Out-Null
        exit 1
    }
    if ($catalegs.Count -eq 1) { return [pscustomobject]@{ Nav='next'; Data=$catalegs[0] } }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 1 - Tipus d''informe'
    $form.Size = New-Object System.Drawing.Size(500, 320)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Selecciona el cataleg de deficiencies:'
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(15, 40)
    $list.Size = New-Object System.Drawing.Size(450, 180)
    foreach ($c in $catalegs) { [void]$list.Items.Add($c.BaseName) }
    $list.SelectedIndex = 0
    if ($preloadBaseName) {
        for ($i = 0; $i -lt $catalegs.Count; $i++) {
            if ($catalegs[$i].BaseName -eq $preloadBaseName) { $list.SelectedIndex = $i; break }
        }
    }
    $form.Controls.Add($list)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(290, 240)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    # No hi ha boto Cancel ni Enrere: el Pas 1 es el primer. Tancar la
    # finestra (X) avorta el programa.
    if ($form.ShowDialog() -ne 'OK') { exit 0 }
    return [pscustomobject]@{ Nav='next'; Data=$catalegs[$list.SelectedIndex] }
}

# ----------------------------------------------------------------------------
# Step 2 - Header data (formulari + precarrega Excel)
# ----------------------------------------------------------------------------
# Construeix el formulari de capcalera (controls + botons), retorna la
# tupla amb el formulari, el diccionari de controls i el boto Cercar perque
# Get-HeaderData hi puga lligar la logica de cerca i validacio.
function _BuildHeaderForm($excelInfo) {
    # $excelInfo: hashtable amb claus Text (string) i Source ('primary'|'fallback').
    # Quan es 'fallback', el rotul es taronja i porta el text "[FALLBACK LOCAL]"
    # davant perque l'usuari sapiga que NO esta usant la base de dades oficial
    # de la xarxa.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lblBd = New-Object System.Windows.Forms.Label
    $lblBd.Location = New-Object System.Drawing.Point(15, 12)
    $lblBd.Size = New-Object System.Drawing.Size(680, 22)
    if ($excelInfo.Source -eq 'fallback') {
        $lblBd.Text = "[FALLBACK LOCAL]  " + $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblBd.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    } else {
        $lblBd.Text = $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    $form.Controls.Add($lblBd)

    $controls = @{}
    $addRow = {
        param($label, $y, $tbWidth, $key)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $label
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(200, 22)
        [void]$form.Controls.Add($lbl)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(220, ($y - 2))
        $tb.Size = New-Object System.Drawing.Size($tbWidth, 22)
        [void]$form.Controls.Add($tb)
        $controls[$key] = $tb
    }

    $y = 50
    & $addRow 'ID GIA' $y 380 'ID_GIA'
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = 'Cercar'
    $btnSearch.Location = New-Object System.Drawing.Point(605, ($y - 3))
    $btnSearch.Size = New-Object System.Drawing.Size(80, 26)
    [void]$form.Controls.Add($btnSearch)
    $y += 38

    & $addRow "Num. d'expedient (autom., editable)"   $y 460 'EXP_NUM';      $y += 38
    & $addRow 'Titular (autom., editable)'            $y 460 'TITULAR';      $y += 38
    & $addRow 'Adreca (autom., editable)'             $y 460 'ADRECA';       $y += 38
    & $addRow 'Activitat (autom., editable)'          $y 460 'ACTIVITAT';    $y += 38
    & $addRow "Num. d'anotacio (autom., editable)"    $y 460 'NUM_ANOTACIO'; $y += 38
    & $addRow "Data d'anotacio (autom., editable)"    $y 460 'DATA_ANOTACIO';$y += 50

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    [void]$form.Controls.Add($back)

    $recover = New-Object System.Windows.Forms.Button
    $recover.Text = "Recuperar dades ultim informe"
    $recover.Location = New-Object System.Drawing.Point(115, $y)
    $recover.Size = New-Object System.Drawing.Size(250, 28)
    [void]$form.Controls.Add($recover)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(595, $y)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    return @{ Form=$form; Controls=$controls; BtnSearch=$btnSearch; BtnOk=$ok; BtnBack=$back; BtnRecover=$recover }
}

# Llegeix els valors dels controls i retorna un hashtable amb la capcalera.
function _ReadHeaderControls($controls) {
    @{
        ID_GIA        = $controls['ID_GIA'].Text.Trim()
        EXP_NUM       = $controls['EXP_NUM'].Text.Trim()
        TITULAR       = $controls['TITULAR'].Text.Trim()
        ADRECA        = $controls['ADRECA'].Text.Trim()
        ACTIVITAT     = $controls['ACTIVITAT'].Text.Trim()
        NUM_ANOTACIO  = $controls['NUM_ANOTACIO'].Text.Trim()
        DATA_ANOTACIO = $controls['DATA_ANOTACIO'].Text.Trim()
    }
}

# Precarrega valors d'una capcalera anterior als controls del formulari.
# $preload pot ser un hashtable (navegacio en memoria) o un PSCustomObject
# (dades de l'ultim informe llegides de JSON).
function _PreloadHeaderControls($controls, $preload) {
    if ($null -eq $preload) { return }
    foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','NUM_ANOTACIO','DATA_ANOTACIO') {
        $v = $null
        if ($preload -is [System.Collections.IDictionary]) {
            if ($preload.Contains($k)) { $v = $preload[$k] }
        } elseif ($preload.PSObject.Properties.Name -contains $k) {
            $v = $preload.$k
        }
        if ($null -ne $v) { $controls[$k].Text = [string]$v }
    }
}

function Get-HeaderData {
    param($preload = $null)

    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' a cap de les ubicacions:`n`n" +
            "  1. $ActivitatsDir`n" +
            "  2. $LocalActivitatsDir  (fallback local)`n`n" +
            "Si estas fora de la feina, copia una base de dades a la carpeta`n" +
            "'BASE DE DADES ACTIVITATS' dins de la carpeta del programa.",
            'Base de dades no trobada', 'OK', 'Error') | Out-Null
        exit 1
    }
    if ($latest.Source -eq 'fallback') {
        # Avis explicit en obrir el Pas 2 perque l'usuari sapiga que esta
        # treballant amb una copia local i no amb la xarxa de la feina.
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la base de dades a la xarxa.`n`n" +
            "S'usa la copia LOCAL del clone:`n  $($latest.File.FullName)`n`n" +
            "Comprova que sigui prou recent.",
            'Base de dades: copia local (fallback)', 'OK', 'Warning') | Out-Null
    }

    # Precarrega TOTA la base de dades a memoria una sola vegada. A partir
    # d'aqui les cerques son immediates (no cal reobrir Excel).
    try {
        $actCache = Initialize-ActivitatsCache -excelFile $latest.File
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error llegint l'Excel:`n$($_.Exception.Message)",'Error','OK','Error') | Out-Null
        exit 1
    }
    if ($actCache.Warnings -and $actCache.Warnings.Count -gt 0) {
        $msg = "Avisos de validacio de l'Excel (l'auto-fill podria fallar):`n`n" + ($actCache.Warnings -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg,'Avisos','OK','Warning') | Out-Null
    }

    $labelText = "Base de dades d'activitats: $($latest.File.Name)  (data: $($latest.Date.ToString('yyyy-MM-dd'))) - $($actCache.ById.Count) activitats carregades"
    $f = _BuildHeaderForm @{ Text = $labelText; Source = $latest.Source }
    $form       = $f.Form
    $controls   = $f.Controls
    $btnSearch  = $f.BtnSearch
    $ok         = $f.BtnOk
    $btnRecover = $f.BtnRecover

    _PreloadHeaderControls $controls $preload

    # Boto "Recuperar dades ultim informe": carrega les dades de l'ultim
    # informe generat amb exit i les deixa als formularis perque l'usuari
    # les revisi/modifiqui pas per pas. La resta de passos (seleccio, camps,
    # conclusions) es precarreguen via $script:_recoveredReport, que Main llegeix.
    $script:_recoveredReport = $null
    $btnRecover.add_Click({
        $rep = Load-LastReport
        if ($null -eq $rep) {
            [System.Windows.Forms.MessageBox]::Show("No hi ha cap informe anterior desat.",'Sense dades','OK','Information') | Out-Null
            return
        }
        if ($rep.Header) { _PreloadHeaderControls $controls $rep.Header }
        $script:_recoveredReport = $rep
        [System.Windows.Forms.MessageBox]::Show(
            "Dades de l'ultim informe carregades.`n`nRevisa-les i modifica el que calgui (per exemple, canvia l'ID GIA i prem 'Cercar' per a una activitat nova). En continuar, els passos seguents tambe sortiran precarregats.",
            'Recuperat', 'OK', 'Information') | Out-Null
    })

    # Cerca per ID GIA: instantania des del cache. Omple els camps automatics.
    $doSearch = {
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return $false
        }
        $act = Get-ActivitatFromCache $actCache $idGia
        if ($null -eq $act) {
            [System.Windows.Forms.MessageBox]::Show(
                "L'ID GIA '$idGia' no s'ha trobat a la base de dades`n($($latest.File.Name)).",
                'Activitat no trobada', 'OK', 'Error') | Out-Null
            foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
                $controls[$k].Text = ''
            }
            return $false
        }
        foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
            if ($act.ContainsKey($k)) { $controls[$k].Text = [string]$act[$k] }
        }
        return $true
    }

    $btnSearch.add_Click({ [void](& $doSearch) })

    $script:_headerData = $null
    $ok.add_Click({
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falten dades','OK','Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($controls['TITULAR'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ADRECA'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ACTIVITAT'].Text)) {
            if (-not (& $doSearch)) { return }
        }
        $script:_headerData = _ReadHeaderControls $controls
        $form.DialogResult = 'OK'
        $form.Close()
    })

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }
    return [pscustomobject]@{ Nav='next'; Data=$script:_headerData; Recovered=$script:_recoveredReport }
}

# ----------------------------------------------------------------------------
# Step 3 - Parse cataleg
# ----------------------------------------------------------------------------
# NOTA: es va provar una cache en disc del resultat del parseig (JSON), pero
# el round-trip ConvertTo-Json/ConvertFrom-Json no preserva de manera fiable
# l'estructura niada (BodyLines/Children), cosa que trencava el format dels
# enllacos al document final. El parseig d'un .docx triga molt poc, aixi que
# es fa sempre en fresc. Si en el futur es vol cachejar, cal fer-ho amb
# Export-Clixml/Import-Clixml (preserva tipus i arrays), no amb JSON.
# Cert si $styleName encaixa amb "Heading N" (N=1 o 2) en qualsevol de les
# variants que escupen els Word EN/CA/ES, incloent les variants compactades
# 'Ttulo1' / 'Ttulo 1' que apareixen amb el Word castella d'algunes versions.
function Test-StyleMatch([string]$styleName, [int]$level) {
    if ([string]::IsNullOrWhiteSpace($styleName)) { return $false }
    # Normalitzem: minuscules, treiem accents i caracters no alfanumerics
    # (espais, guions, etc.), aixi 'Título 1', 'Titol 1' i 'Ttulo1' col·lapsen.
    $n = (_NormalizeText $styleName) -replace '[^a-z0-9]',''
    $patterns = @(
        "heading$level",
        "titulo$level",
        "titol$level",
        "ttulo$level"     # variant que apareix per a 'Ttulo1' (sense accent ni espai)
    )
    foreach ($pat in $patterns) {
        if ($n -eq $pat) { return $true }
    }
    return $false
}

function Get-ParsedCataleg($word, $path) {
    return (Parse-Cataleg -word $word -path $path)
}

function Parse-Cataleg($word, $path) {
    # Retorna un PSCustomObject amb:
    #   IntroText : la frase introductoria del cataleg (primer paragraf Normal
    #               abans de la primera seccio). Apareix sempre al document.
    #   Sections  : llista de seccions. Cada seccio te:
    #                 Title : titol Heading 1.
    #                 Items : llista plana d'elements del catalog. Cada element
    #                         te un camp Kind:
    #                           'item'       (Heading 2 sense prefix)
    #                           'subsection' (Heading 2 ::SUB::)
    #                           'intro'      (Heading 2 ::INTRO::)
    #                         Els items poden tenir Children (Heading 2 ::CHILD::).
    #
    # Estils Word reconeguts (NameLocal segons l'idioma del Word de l'usuari):
    #   Heading 1, Titol 1, Titulo 1, Tisingleitulo 1 (placeholder per a accents):
    #   en realitat: 'Titulo 1', 'Título 1', 'Titol 1', 'Títol 1' i la variant
    #   compactada 'Ttulo1' / 'Ttulo 1' que apareix amb alguns Word castellans.
    #
    # Estil 'Cita' (o 'Cite'/'Quote' en angles, 'Cita' en castella/catala) es
    # tracta com a paragraf d'URL: el text es l'enllac.
    $doc = $word.Documents.Open($path, $false, $true)  # ReadOnly
    try {
        $sections      = New-Object System.Collections.ArrayList
        $introText     = ''
        $currentSection = $null
        $lastItem      = $null   # darrer Heading 2 'item' (per associar fills)
        $lastH2        = $null   # darrer Heading 2 sigui del tipus que sigui

        foreach ($p in $doc.Paragraphs) {
            $text = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            $styleName = ''
            try { $styleName = $p.Style.NameLocal } catch { }
            # Acceptem: amb/sense espai, amb/sense accent, i diversos idiomes.
            #   "Heading 1"  (en),  "Titol 1"/"Titulo 1" (ca/es),
            #   "Titulo 1" amb i sense espai/accent,
            #   "Ttulo1" i "Ttulo 1" (com els desa el Word castella en alguns casos).
            $isH1 = (Test-StyleMatch $styleName 1)
            $isH2 = (Test-StyleMatch $styleName 2)
            $isCita = ($styleName -match '^(Cita|Cite|Quote|Cita destacada|Quote intense)$')

            if ($isH1) {
                $currentSection = [pscustomobject]@{
                    Title = $text
                    Items = New-Object System.Collections.ArrayList
                }
                [void]$sections.Add($currentSection)
                $lastItem = $null
                $lastH2   = $null
                continue
            }

            if ($isH2) {
                if ($null -eq $currentSection) {
                    $currentSection = [pscustomobject]@{
                        Title = '(Sense seccio)'
                        Items = New-Object System.Collections.ArrayList
                    }
                    [void]$sections.Add($currentSection)
                }
                $kind = 'item'
                $short = $text
                if     ($short -like '::CHILD::*') { $kind = 'child';      $short = $short.Substring('::CHILD::'.Length).Trim() }
                elseif ($short -like '::SUB::*')   { $kind = 'subsection'; $short = $short.Substring('::SUB::'.Length).Trim() }
                elseif ($short -like '::INTRO::*') { $kind = 'intro';      $short = $short.Substring('::INTRO::'.Length).Trim() }

                $newEl = [pscustomobject]@{
                    Kind      = $kind
                    Short     = $short
                    BodyLines = New-Object System.Collections.ArrayList
                    Children  = New-Object System.Collections.ArrayList
                }

                if ($kind -eq 'child' -and $null -ne $lastItem) {
                    [void]$lastItem.Children.Add($newEl)
                } else {
                    [void]$currentSection.Items.Add($newEl)
                    if ($kind -eq 'item')        { $lastItem = $newEl }
                    elseif ($kind -eq 'subsection') { $lastItem = $null }
                }
                $lastH2 = $newEl
                continue
            }

            # Paragraf Normal o Cita: l'afegim al BodyLines de l'element actiu.
            # Si l'estil es Cita, marquem el text amb un prefix intern [[URL]]
            # perque l'emissor (_SplitTextAndUrls) el tracti SEMPRE com a URL
            # (i no com a text), encara que no comenci per "http". Aixo permet
            # a l'usuari marcar enllacos al Word de manera explicita per estil,
            # sense dependre de regex sobre el contingut.
            $textToAdd = if ($isCita) { '[[URL]] ' + $text } else { $text }
            if ($null -eq $lastH2) {
                if ($null -eq $currentSection -and [string]::IsNullOrWhiteSpace($introText)) {
                    $introText = $textToAdd
                }
                continue
            }
            $target = $lastH2
            if ($lastH2.Kind -eq 'item' -and $lastH2.Children.Count -gt 0) {
                $target = $lastH2.Children[$lastH2.Children.Count - 1]
            }
            [void]$target.BodyLines.Add($textToAdd)
        }
        return [pscustomobject]@{ IntroText = $introText; Sections = $sections }
    }
    finally {
        $doc.Close($false)
    }
}

# ----------------------------------------------------------------------------
# Step 3 (UI) - TreeView amb filtre + checkboxes
# ----------------------------------------------------------------------------
# Helpers per al filtre del TreeView. Es manté l'estructura $sections a part i
# es reconstrueix el tree quan canvia el filtre, preservant els check states.

function _TextMatches($text, $needle) {
    if ([string]::IsNullOrEmpty($needle)) { return $true }
    if ($null -eq $text) { return $false }
    return $text.ToLower().Contains($needle.ToLower())
}

# Reconstrueix el TreeView segons el text de filtre. Preserva check states
# (passats en una hashtable [key] -> bool) i els actualitza durant la construccio.
# Bloquegem la propagacio automatica de check durant el rebuild perque
# marcar nodes programmaticament dispara l'event AfterCheck.
function _RebuildTree($tv, $sections, $needle, $checkStates) {
    $tv.BeginUpdate()
    $script:_propagating = $true
    try {
        $tv.Nodes.Clear()
        foreach ($sec in $sections) {
            $secMatches = _TextMatches $sec.Title $needle

            # Recollim items/children que cal mostrar
            $itemNodesToAdd = New-Object System.Collections.ArrayList
            $currentContainer = $null
            foreach ($el in $sec.Items) {
                if ($el.Kind -eq 'subsection') {
                    $subShow = $secMatches -or (_TextMatches $el.Short $needle)
                    $itemNodesToAdd.Add(@{ Kind='Subsection'; El=$el; ChildShows=@(); ShowMe=$subShow }) | Out-Null
                    continue
                }
                if ($el.Kind -eq 'intro') { continue }  # mai al TreeView

                $itemMatches = _TextMatches $el.Short $needle
                $matchedChildren = New-Object System.Collections.ArrayList
                foreach ($ch in $el.Children) {
                    if ($secMatches -or $itemMatches -or (_TextMatches $ch.Short $needle)) {
                        [void]$matchedChildren.Add($ch)
                    }
                }
                $showItem = $secMatches -or $itemMatches -or ($matchedChildren.Count -gt 0)
                if ($showItem) {
                    $itemNodesToAdd.Add(@{ Kind='Item'; El=$el; ChildShows=$matchedChildren; ShowMe=$true }) | Out-Null
                }
            }

            # Si el filtre no es buit i no hi ha cap item/subsection visible,
            # ometem la seccio del tot (tret que el titol de la seccio matchi).
            $anyChild = $false
            foreach ($n in $itemNodesToAdd) { if ($n.ShowMe) { $anyChild = $true; break } }
            if (-not $secMatches -and -not $anyChild) { continue }

            $secNode = New-Object System.Windows.Forms.TreeNode($sec.Title)
            $secNode.Tag = @{ Kind = 'Section'; Ref = $sec; Key = $sec.Title }
            $secNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
            $stKey = "SECT::$($sec.Title)"
            if ($checkStates.ContainsKey($stKey)) { $secNode.Checked = $checkStates[$stKey] }
            [void]$tv.Nodes.Add($secNode)

            $container = $secNode
            foreach ($n in $itemNodesToAdd) {
                if (-not $n.ShowMe) { continue }
                if ($n.Kind -eq 'Subsection') {
                    $subNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                    $subNode.Tag = @{ Kind = 'Subsection'; Ref = $n.El }
                    $subNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
                    [void]$secNode.Nodes.Add($subNode)
                    $container = $subNode
                    continue
                }
                # Item. NodeFont explicit (regular) perque el font base del
                # TreeView es negreta (vegeu nota a Select-Items) i no volem
                # que els items surtin en negreta ni que es retallin.
                $itNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                $itNode.Tag = @{ Kind = 'Item'; Ref = $n.El; SectionTitle = $sec.Title }
                $itNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                $itKey = (_ItemKey $sec.Title $n.El.Short)
                if ($checkStates.ContainsKey($itKey)) { $itNode.Checked = $checkStates[$itKey] }
                [void]$container.Nodes.Add($itNode)
                foreach ($ch in $n.ChildShows) {
                    $chNode = New-Object System.Windows.Forms.TreeNode($ch.Short)
                    $chNode.Tag = @{ Kind = 'Child'; Ref = $ch; SectionTitle = $sec.Title; ParentShort = $n.El.Short }
                    $chNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                    $chKey = (_ItemKey $sec.Title $n.El.Short $ch.Short)
                    if ($checkStates.ContainsKey($chKey)) { $chNode.Checked = $checkStates[$chKey] }
                    [void]$itNode.Nodes.Add($chNode)
                }
            }
            $secNode.ExpandAll()
        }
    } finally {
        $script:_propagating = $false
        $tv.EndUpdate()
    }
}

# Recorre el TreeView i llegeix tots els check states en una hashtable.
function _CollectCheckStates($tv, $checkStates) {
    foreach ($secNode in $tv.Nodes) {
        $secTitle = $secNode.Tag.Ref.Title
        $checkStates["SECT::$secTitle"] = [bool]$secNode.Checked
        foreach ($node in $secNode.Nodes) {
            $kind = $node.Tag.Kind
            if ($kind -eq 'Subsection') {
                foreach ($itNode in $node.Nodes) {
                    $itShort = $itNode.Tag.Ref.Short
                    $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$itNode.Checked
                    foreach ($chNode in $itNode.Nodes) {
                        $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                    }
                }
            } elseif ($kind -eq 'Item') {
                $itShort = $node.Tag.Ref.Short
                $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$node.Checked
                foreach ($chNode in $node.Nodes) {
                    $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                }
            }
        }
    }
}

function Select-Items {
    param($sections, $preloadSelectedKeys = $null)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 3 - Seleccio de deficiencies'
    $form.Size = New-Object System.Drawing.Size(1100, 720)
    $form.StartPosition = 'CenterScreen'

    # Filtre (textbox al capdamunt). Cada vegada que canvia, es reconstrueix
    # el TreeView amb nomes les coincidencies. Els check states es preserven.
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = 'Filtre:'
    $lblFilter.Location = New-Object System.Drawing.Point(10, 14)
    $lblFilter.AutoSize = $true
    $form.Controls.Add($lblFilter)

    $tbFilter = New-Object System.Windows.Forms.TextBox
    $tbFilter.Location = New-Object System.Drawing.Point(60, 10)
    $tbFilter.Size = New-Object System.Drawing.Size(400, 22)
    $tbFilter.Anchor = 'Top, Left'
    $form.Controls.Add($tbFilter)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Esborra'
    $btnClear.Location = New-Object System.Drawing.Point(465, 9)
    $btnClear.Size = New-Object System.Drawing.Size(70, 24)
    $btnClear.add_Click({ $tbFilter.Text = '' })
    $form.Controls.Add($btnClear)

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(10, 40)
    $tv.Size = New-Object System.Drawing.Size(1060, 590)
    $tv.CheckBoxes = $true
    $tv.HideSelection = $false
    $tv.ShowNodeToolTips = $true
    $tv.Anchor = 'Top, Bottom, Left, Right'
    # El font base es la negreta mes ampla que faran servir les seccions. Aixo
    # evita el bug de WinForms en que un node amb NodeFont mes ample que el
    # font del control surt retallat. Items i fills posen NodeFont regular.
    $tv.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($tv)

    # Estats de check persistents entre rebuilds. Inicialitzat des de session.
    $checkStates = @{}
    if ($preloadSelectedKeys) {
        foreach ($k in $preloadSelectedKeys) {
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            $checkStates[[string]$k] = $true
        }
    }

    # Build inicial sense filtre
    _RebuildTree $tv $sections '' $checkStates

    # Propagacio recursiva: marcar un node marca tots els descendents.
    $script:_propagating = $false
    $propagate = {
        param($node)
        foreach ($c in $node.Nodes) {
            $c.Checked = $node.Checked
            & $propagate $c
        }
    }
    $tv.add_AfterCheck({
        param($sender, $e)
        if ($script:_propagating) { return }
        $script:_propagating = $true
        try { & $propagate $e.Node } finally { $script:_propagating = $false }
    })

    # Refilter en temps real (debouncing simple: rebuild a cada keystroke;
    # amb 131 items va fluid)
    $tbFilter.add_TextChanged({
        # Guardem l'estat actual ABANS de reconstruir
        _CollectCheckStates $tv $checkStates
        _RebuildTree $tv $sections $tbFilter.Text $checkStates
    })

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(10, 640)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    $back.Anchor = 'Bottom, Left'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(980, 640)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Recollim l'estat final i el barregem amb el que tenim memoritzat per
    # items que ara mateix no es mostren (perque hi hagi filtre actiu).
    _CollectCheckStates $tv $checkStates

    # Construim el resultat en ordre del data, preservant subseccions/intros.
    $result = New-Object System.Collections.ArrayList
    foreach ($sec in $sections) {
        $chosen = New-Object System.Collections.ArrayList
        foreach ($el in $sec.Items) {
            if ($el.Kind -in 'subsection','intro') {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = $el.Kind
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = @()
                    Selected  = $false
                })
                continue
            }
            $itKey = (_ItemKey $sec.Title $el.Short)
            $isSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
            $chosenChildren = New-Object System.Collections.ArrayList
            foreach ($ch in $el.Children) {
                $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
                if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) {
                    [void]$chosenChildren.Add($ch)
                }
            }
            if ($isSel -or $chosenChildren.Count -gt 0) {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = 'item'
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = $chosenChildren
                    Selected  = [bool]$isSel
                })
            }
        }
        $hasRealItem = $false
        foreach ($x in $chosen) { if ($x.Kind -eq 'item') { $hasRealItem = $true; break } }
        if ($hasRealItem) {
            [void]$result.Add([pscustomobject]@{
                Title = $sec.Title
                Items = $chosen
            })
        }
    }
    if ($result.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No s''ha seleccionat cap deficiencia.','Avis','OK','Warning') | Out-Null
        return [pscustomobject]@{ Nav='stay' }   # es torna a mostrar el Pas 3
    }
    return [pscustomobject]@{ Nav='next'; Data=$result }
}

# Extreu les claus "Seccio::Item[::Fill]" del resultat de Select-Items, per
# desar-les a la sessio.
function Get-SelectedKeysFromResult($selectedSections) {
    $keys = New-Object System.Collections.ArrayList
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            if ($it.Kind -ne 'item') { continue }
            if ($it.Selected) { [void]$keys.Add((_ItemKey $sec.Title $it.Short)) }
            foreach ($ch in $it.Children) {
                [void]$keys.Add((_ItemKey $sec.Title $it.Short $ch.Short))
            }
        }
    }
    return $keys.ToArray()
}

# ----------------------------------------------------------------------------
# Step 4 - Field placeholders
#   [CAMP: nom]                  -> camp de text lliure
#   [CAMP: nom (hint)]           -> camp de text amb ajuda
#   [OPCIO: nom | A | B | C]     -> desplegable; l'usuari tria A, B o C i el
#                                   text triat substitueix el placeholder
# ----------------------------------------------------------------------------
$Script:CampRegex  = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'
$Script:OpcioRegex = [regex]'\[OPCIO:\s*([^\]]+?)\s*\]'

# Analitza el contingut d'un [OPCIO: ...]: "nom | A | B" -> nom + opcions.
function _ParseOpcio($raw) {
    $segs = $raw -split '\|'
    $name = $segs[0].Trim()
    $opts = @()
    for ($i = 1; $i -lt $segs.Count; $i++) {
        $o = $segs[$i].Trim()
        if ($o -ne '') { $opts += $o }
    }
    return @{ Name = $name; Options = $opts }
}

# Detecta [CAMP: ...] i [OPCIO: ...] dins $allText i els afegeix a $fields
# (sense duplicar). Modifica $fields in-place.
function _AddFieldsFromText($fields, $allText) {
    foreach ($m in $Script:CampRegex.Matches($allText)) {
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $hint = ''
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) {
            $name = $raw.Substring(0, $parenIdx).Trim()
            $hint = $raw.Substring($parenIdx).Trim().TrimStart('(').TrimEnd(')')
        }
        if (-not $fields.Contains($name)) {
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='text'; Hint=$hint; Options=@(); Value='' }
        }
    }
    foreach ($m in $Script:OpcioRegex.Matches($allText)) {
        $parsed = _ParseOpcio $m.Groups[1].Value
        $name = $parsed.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $fields.Contains($name)) {
            $val = if ($parsed.Options.Count -gt 0) { [string]$parsed.Options[0] } else { '' }
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='choice'; Hint=''; Options=$parsed.Options; Value=$val }
        }
    }
}

function Get-FieldsFromSelection($selectedSections) {
    $fields = [ordered]@{}
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            $allText = ($it.BodyLines -join ' ')
            foreach ($ch in $it.Children) {
                $allText += ' ' + ($ch.BodyLines -join ' ')
            }
            _AddFieldsFromText $fields $allText
        }
    }
    return $fields
}

# Afegeix els camps detectats a les conclusions TRIADES i les SEMPRE al
# diccionari $fields existent. Aixi al Pas 4 surten alhora els del REQ1
# i els del CONCLUSIONS.
function Add-FieldsFromConclusions($fields, $selectedConcl, $alwaysConcl) {
    foreach ($c in $selectedConcl) {
        $body = if ($c -is [string]) { $c } else { [string]$c.Body }
        _AddFieldsFromText $fields $body
    }
    foreach ($a in $alwaysConcl) {
        _AddFieldsFromText $fields ([string]$a)
    }
}

function Prompt-Fields {
    param($fields, $preloadValues = $null)
    if ($fields.Count -eq 0) { return [pscustomobject]@{ Nav='next'; Data=$fields } }

    # Precarrega valors anteriors (per nom de camp)
    if ($preloadValues) {
        foreach ($name in $fields.Keys) {
            $v = $null
            if ($preloadValues -is [System.Collections.IDictionary] -and $preloadValues.Contains($name)) {
                $v = $preloadValues[$name]
            } elseif ($preloadValues -is [psobject] -and ($preloadValues.PSObject.Properties.Name -contains $name)) {
                $v = $preloadValues.$name
            }
            if ($null -ne $v) { $fields[$name].Value = [string]$v }
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 4 - Omplir camps'
    $form.StartPosition = 'CenterScreen'
    $form.AutoScroll = $true

    $y = 15
    $inputs = @{}   # nom -> control (TextBox o ComboBox)
    foreach ($name in $fields.Keys) {
        $f = $fields[$name]
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $name
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(520, 22)
        $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $form.Controls.Add($lbl)
        $y += 22

        if ($f.Hint) {
            $hintLbl = New-Object System.Windows.Forms.Label
            $hintLbl.Text = $f.Hint
            $hintLbl.Location = New-Object System.Drawing.Point(15, $y)
            $hintLbl.Size = New-Object System.Drawing.Size(520, 18)
            $hintLbl.ForeColor = [System.Drawing.Color]::DimGray
            $form.Controls.Add($hintLbl)
            $y += 18
        }

        if ($f.Type -eq 'choice') {
            # Desplegable (l'usuari nomes pot triar de la llista)
            $cb = New-Object System.Windows.Forms.ComboBox
            $cb.Location = New-Object System.Drawing.Point(15, $y)
            $cb.Size = New-Object System.Drawing.Size(520, 24)
            $cb.DropDownStyle = 'DropDownList'
            foreach ($o in $f.Options) { [void]$cb.Items.Add($o) }
            $idx = if ($f.Value) { $cb.Items.IndexOf([string]$f.Value) } else { -1 }
            if ($idx -lt 0 -and $cb.Items.Count -gt 0) { $idx = 0 }
            if ($idx -ge 0) { $cb.SelectedIndex = $idx }
            $form.Controls.Add($cb)
            $inputs[$name] = $cb
        } else {
            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Location = New-Object System.Drawing.Point(15, $y)
            $tb.Size = New-Object System.Drawing.Size(520, 22)
            $tb.Text = $f.Value
            $form.Controls.Add($tb)
            $inputs[$name] = $tb
        }
        $y += 32
    }

    $form.ClientSize = New-Object System.Drawing.Size(560, [Math]::Min(640, ($y + 70)))

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(450, $y)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }
    foreach ($name in $fields.Keys) {
        $ctrl = $inputs[$name]
        if ($fields[$name].Type -eq 'choice') {
            $fields[$name].Value = if ($null -ne $ctrl.SelectedItem) { [string]$ctrl.SelectedItem } else { '' }
        } else {
            $fields[$name].Value = $ctrl.Text
        }
    }
    return [pscustomobject]@{ Nav='next'; Data=$fields }
}

function Apply-Fields($text, $fields) {
    # Primer els desplegables [OPCIO: nom | ...] i despres els [CAMP: ...].
    $out = $Script:OpcioRegex.Replace($text, {
        param($m)
        $name = (_ParseOpcio $m.Groups[1].Value).Name
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    $out = $Script:CampRegex.Replace($out, {
        param($m)
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) { $name = $raw.Substring(0, $parenIdx).Trim() }
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    return $out
}

# Extreu els valors dels camps en un hashtable simple per a la sessio.
function Get-FieldValuesForSession($fields) {
    $h = @{}
    foreach ($name in $fields.Keys) { $h[$name] = $fields[$name].Value }
    return $h
}

# Separa el text d'una linia dels URLs que pugui contenir. Retorna:
#   @{ Text = '<tot el que hi ha abans del primer URL>'; Urls = @(url1, url2...) }
#
# Hi ha dues fonts d'URLs reconegudes:
#   1. Prefix intern '[[URL]] ': el ha posat Parse-Cataleg quan el paragraf
#      del .docx te estil 'Cita' (manera explicita, recomanada al cataleg
#      modern). En aquest cas tota la linia es l'URL.
#   2. Deteccio per contingut: qualsevol token que comenci per 'http://' o
#      'https://' (retrocompatible amb cataleg vell).
function _SplitTextAndUrls($line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return @{ Text=''; Urls=@() } }
    # Cas 1: estil Cita marcat per Parse-Cataleg.
    if ($line.StartsWith('[[URL]] ')) {
        $url = $line.Substring('[[URL]] '.Length).Trim()
        return @{ Text=''; Urls=@($url) }
    }
    # Cas 2: deteccio per contingut.
    $m = [regex]::Match($line, 'https?://')
    if (-not $m.Success) { return @{ Text = $line.Trim(); Urls=@() } }
    $text = $line.Substring(0, $m.Index).Trim()
    $rest = $line.Substring($m.Index)
    $urls = @()
    foreach ($tok in ($rest -split '\s+')) {
        if ($tok -match '^https?://') { $urls += $tok }
    }
    return @{ Text = $text; Urls = $urls }
}

# ----------------------------------------------------------------------------
# Step 5 - Conclusions
# ----------------------------------------------------------------------------
function Read-Conclusions($word, $path) {
    # Llegeix el fitxer 0 CONCLUSIONS.docx i retorna un PSCustomObject amb:
    #
    #   HeaderText       : text del titol del document (sol ser 'CONCLUSIONS'),
    #                      llegit del primer paragraf centrat-negreta. '' si no n'hi ha.
    #   Selectable       : llista d'objectes triables al Pas 5. Cada element:
    #                        Title : el titol curt (Ttulo1) que es mostra a la
    #                                checkbox del Pas 5.
    #                        Body  : el text complet del cos (paragraf Normal
    #                                que segueix al Ttulo1) que s'imprimeix
    #                                si l'usuari el tria.
    #   Always           : llista de cadenes amb les frases fixes. Son els
    #                      paragrafs Normal que comencen amb '::SEMPRE:: '
    #                      (s'inclouen sempre al final del document, sense
    #                      passar pel Pas 5). El prefix s'elimina.
    #
    # NOTES sobre el format esperat de 0 CONCLUSIONS.docx:
    #   - Primer paragraf (opcional): titol del bloc (centrat-negreta).
    #   - Per cada conclusio triable: un paragraf Ttulo1 (titol curt) + un
    #     paragraf Normal (cos).
    #   - Frases fixes (al final, sempre): paragrafs Normal que comencen amb
    #     '::SEMPRE:: '. S'imprimeixen en l'ordre del fitxer.
    $empty = [pscustomobject]@{ HeaderText=''; Selectable=@(); Always=@() }
    if (-not (Test-Path -LiteralPath $path)) { return $empty }

    $doc = $word.Documents.Open($path, $false, $true)
    try {
        $headerText = ''
        $selectable = New-Object System.Collections.ArrayList
        $always     = New-Object System.Collections.ArrayList
        $pendingTitle = $null   # ultim Ttulo1 vist (esperant cos Normal)
        $isFirstPara  = $true

        foreach ($p in $doc.Paragraphs) {
            $text = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            $styleName = ''
            try { $styleName = $p.Style.NameLocal } catch { }
            $isH1 = Test-StyleMatch $styleName 1

            # Titol del bloc: primer paragraf no buit, centrat (-jc center)
            # i amb estil Normal. Es queda nomes per ser emes per
            # _WriteConclusionsBlock.
            if ($isFirstPara -and -not $isH1) {
                $isFirstPara = $false
                $jc = ''
                try { $jc = [string]$p.Format.Alignment } catch { }
                # wdAlignParagraphCenter = 1
                if ($jc -eq '1' -or $jc -eq 'wdAlignParagraphCenter') {
                    $headerText = $text
                    continue
                }
            }
            $isFirstPara = $false

            if ($isH1) {
                # Nou titol de conclusio. Si l'anterior queda sense cos, l'ignorem.
                $pendingTitle = $text
                continue
            }

            # Paragraf Normal:
            if ($text.StartsWith('::SEMPRE::')) {
                $stripped = $text.Substring('::SEMPRE::'.Length).Trim()
                [void]$always.Add($stripped)
                $pendingTitle = $null
                continue
            }

            if ($null -ne $pendingTitle) {
                # Cos de la conclusio precedida pel Ttulo1.
                [void]$selectable.Add([pscustomobject]@{
                    Title = $pendingTitle
                    Body  = $text
                })
                $pendingTitle = $null
            }
            # Si no hi havia titol pendent ni '::SEMPRE::', ignorem (text
            # de transicio sense rol clar).
        }

        return [pscustomobject]@{
            HeaderText = $headerText
            Selectable = $selectable.ToArray()
            Always     = $always.ToArray()
        }
    } finally {
        $doc.Close($false)
    }
}

function Select-Conclusions {
    # $conclusions : array d'objectes {Title, Body} (de Read-Conclusions).
    # $preloadTitles : array de titols (string) preseleccionats (sessio anterior).
    param($conclusions, $preloadTitles = $null)
    if ($conclusions.Count -eq 0) { return [pscustomobject]@{ Nav='next'; Data=@() } }

    # Convertim preloadTitles a un HashSet per a comparacio rapida.
    $preloadSet = New-Object System.Collections.Generic.HashSet[string]
    if ($preloadTitles) { foreach ($t in $preloadTitles) { [void]$preloadSet.Add([string]$t) } }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 5 - Conclusions'
    $form.Size = New-Object System.Drawing.Size(780, 560)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Selecciona les conclusions a incloure:'
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 35)
    $panel.Size = New-Object System.Drawing.Size(730, 430)
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panel)

    $checks = @()
    $y = 5
    for ($i = 0; $i -lt $conclusions.Count; $i++) {
        $c = $conclusions[$i]
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $c.Title    # nomes el titol curt, mes llegible
        $cb.Location = New-Object System.Drawing.Point(5, $y)
        $cb.Size = New-Object System.Drawing.Size(700, 28)
        $cb.AutoSize = $false
        $cb.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $cb.Tag = $c
        # Tooltip: cos sencer per veure'l sense haver de tornar al Word
        $tt = New-Object System.Windows.Forms.ToolTip
        $tt.AutoPopDelay = 30000
        $tt.SetToolTip($cb, [string]$c.Body)
        if ($preloadSet.Contains([string]$c.Title)) { $cb.Checked = $true }
        $panel.Controls.Add($cb)
        $checks += $cb
        $y += 34
    }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 480)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Generar'
    $ok.Location = New-Object System.Drawing.Point(655, 480)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Retornem els objectes triats (no nomes el text), per preservar
    # Title i Body per al desat de sessio i l'emissio al document.
    $selected = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $checks.Count; $i++) {
        if ($checks[$i].Checked) { [void]$selected.Add($checks[$i].Tag) }
    }
    return [pscustomobject]@{ Nav='next'; Data=(,$selected.ToArray()) }
}

# ----------------------------------------------------------------------------
# Step 6 - Compose final document
# ----------------------------------------------------------------------------
function Apply-HeaderReplacements($doc, $header) {
    # Substituim els placeholders <<NOM>> de la capcalera pels valors del Pas 2.
    $map = @{
        '<<ID_GIA>>'        = $header['ID_GIA']
        '<<EXP_NUM>>'       = $header['EXP_NUM']
        '<<ADRECA>>'        = $header['ADRECA']
        '<<ACTIVITAT>>'     = $header['ACTIVITAT']
        '<<TITULAR>>'       = $header['TITULAR']
        '<<NUM_ANOTACIO>>'  = $header['NUM_ANOTACIO']
        '<<DATA_ANOTACIO>>' = $header['DATA_ANOTACIO']
    }
    foreach ($k in $map.Keys) {
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $k
        $find.Replacement.Text = [string]$map[$k]
        $find.Forward = $true
        $find.Wrap = 1
        $find.MatchCase = $false
        $find.Execute([ref]$k, $false, $false, $false, $false, $false, $true, 1, $false, [string]$map[$k], 2) | Out-Null
    }
}

# Calcula el nom de fitxer de sortida: YYYY-MM-DD_<TipusCataleg>_GIA <id>.docx
function _GetOutputFileName($catalegName, $gia) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $cat   = $catalegName
    if ($cat) { $cat = $cat.Substring(0,1).ToUpper() + $cat.Substring(1).ToLower() }
    else      { $cat = 'Informe' }
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    return ("{0}_{1}_GIA {2}.docx" -f $today, $cat, $gia)
}

# A partir del nom base, retorna la ruta a $targetDir que no col·lisioni amb
# cap fitxer existent. Si el primer ja existeix, prova "_2", "_3"... fins
# trobar-ne un de lliure. Aixi pots fer diversos informes del mateix GIA el
# mateix dia sense haver de tancar Word ni renombrar res manualment.
#
# Ex.: "2026-05-29_Req1_GIA 1379.docx" existeix
#      -> torna "2026-05-29_Req1_GIA 1379_2.docx"
function _GetUniqueOutputPath($targetDir, $baseFileName) {
    $candidate = Join-Path $targetDir $baseFileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($baseFileName)
    $ext  = [System.IO.Path]::GetExtension($baseFileName)
    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $targetDir ("{0}_{1}{2}" -f $stem, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    # Si arribem aqui es que hi ha mes de 999 informes pel mateix GIA i dia;
    # cas extrem, retornem un nom amb timestamp.
    return Join-Path $targetDir ("{0}_{1}{2}" -f $stem, (Get-Date -Format 'HHmmss'), $ext)
}

# Determina el directori de sortida: l'$OutputDir si es accessible, en cas
# contrari una subcarpeta 'Informes generats' a l'arrel del clone (al
# costat dels .bat).
function _ResolveOutputDir {
    $targetDir = $OutputDir
    try {
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
        return $targetDir
    } catch {
        $local = Join-Path $RepoRoot 'Informes generats'
        if (-not (Test-Path -LiteralPath $local)) {
            New-Item -ItemType Directory -Path $local -Force | Out-Null
        }
        return $local
    }
}

# Obre el document Word a partir d'una copia LOCAL de la capcalera (per
# evitar la "Vista protegida" en unitats de xarxa). Retorna el doc obert i
# la ruta temporal.
function _OpenOutputDocument($word, $tempPath) {
    Copy-Item -LiteralPath $HeaderPath -Destination $tempPath -Force
    $doc = $word.Documents.Open($tempPath, $false, $false)
    try {
        if ($doc.ProtectedViewWindow -ne $null) {
            $doc = $doc.ProtectedViewWindow.Edit()
        }
    } catch { }
    return $doc
}

# Escriu el cos del document (intro del cataleg + seccions amb items numerats).
# Retorna el comptador global utilitzat per a la numeracio.
function _WriteCatalegBody($sel, $cfg, $selectedSections, $fields, $introText) {
    if (-not [string]::IsNullOrWhiteSpace($introText)) {
        Format-Body $sel $introText
        if ($cfg.SpacerAfterIntroParagraph) { Format-Spacer $sel }
    }

    # Resol [CAMP: ...] a cada linia i EMET cada linia resolta al pipeline.
    # Els cridadors fan servir @(& $resolveLines ...) per recollir un array
    # PLA de cadenes. (No retornem ,@(...): combinat amb el @() del cridador
    # provocava un doble embolcall on $itemLines[0] era TOT l'array de linies
    # en comptes de la primera linia -> trencava la separacio text/URL i feia
    # petar Substring.)
    $resolveLines = {
        param($lines)
        foreach ($ln in $lines) {
            [string](Apply-Fields -text $ln -fields $fields)
        }
    }

    # Emet una linia separant text i URLs: el text (si n'hi ha) va com a cos i
    # cada URL com a hipervincle en paragraf propi.
    $emitLine = {
        param($line, $isChild)
        if ([string]::IsNullOrWhiteSpace($line)) { return }
        $parts = _SplitTextAndUrls $line
        if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
            if ($isChild) { Format-Body $sel $parts.Text -IsChild } else { Format-Body $sel $parts.Text }
        }
        foreach ($u in $parts.Urls) {
            if ($isChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
        }
    }

    $emitExtras = {
        param($lines, $isChild)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            & $emitLine $lines[$i] $isChild
        }
    }

    $emitIntro = {
        param($introEl)
        $lines = @(& $resolveLines $introEl.BodyLines)
        foreach ($bp in $lines) { & $emitLine $bp $false }
        if ($cfg.SpacerAfterIntro) { Format-Spacer $sel }
    }

    $emitItem = {
        param($it)
        $itemLines = @(& $resolveLines $it.BodyLines)
        $hasChildren = ($it.Children.Count -gt 0)
        $itemWritten = $false

        if ($it.Selected -or $hasChildren) {
            if ($itemLines.Count -gt 0) {
                $script:_buildGlobal++
                # Separem un possible URL enganxat al text principal de l'item.
                $p0 = _SplitTextAndUrls $itemLines[0]
                Format-Item $sel "$($script:_buildGlobal)." $p0.Text
                foreach ($u in $p0.Urls) { Format-Url $sel $u }
                & $emitExtras $itemLines $false
                $itemWritten = $true
            }
        }
        if ($hasChildren) {
            foreach ($ch in $it.Children) {
                $childLines = @(& $resolveLines $ch.BodyLines)
                if ($childLines.Count -eq 0) { continue }
                if (-not $itemWritten) {
                    $script:_buildGlobal++
                    $itemWritten = $true
                }
                # Els fills NO es numeren: s'emeten com a paragraf de cos
                # amb la sagnia de fill (ChildIndentCm).
                $pc = _SplitTextAndUrls $childLines[0]
                if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                    Format-Body $sel $pc.Text -IsChild
                }
                foreach ($u in $pc.Urls) { Format-Url $sel $u -IsChild }
                & $emitExtras $childLines $true
            }
        }
        if ($itemWritten -and $cfg.SpacerAfterItem) { Format-Spacer $sel }
    }

    $script:_buildGlobal = 0
    $lastSectionName = $null

    foreach ($sec in $selectedSections) {
        $parts = $sec.Title -split ' - ', 2
        if ($parts.Count -eq 2) {
            $secName = $parts[0].Trim()
            $subName = $parts[1].Trim()
            if ($secName -ne $lastSectionName) {
                Format-Section $sel $secName
                if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
                $lastSectionName = $secName
            }
            Format-Subsection $sel $subName
            if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
        } else {
            Format-Section $sel $sec.Title
            if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
            $lastSectionName = $sec.Title
        }

        # Les subseccions i les intros s'emeten "tard": nomes quan ve un
        # item REAL que les segueix. Si una secció conté 3 ::SUB:: pero
        # nomes s'ha triat un ítem que viu a la 3a subsecció, només
        # surten el títol de la secció i la 3a subsecció (no les 2
        # anteriors buides). Una nova subsecció sobreescriu la pendent.
        $pendingSubsection = $null
        $pendingIntro = $null
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'subsection') {
                $pendingSubsection = $el
                $pendingIntro = $null   # una nova subseccio invalida l'intro pendent
                continue
            }
            if ($el.Kind -eq 'intro') {
                $pendingIntro = $el
                continue
            }
            # Item real: emetem primer la subseccio pendent (si en hi ha),
            # despres l'intro pendent, i finalment l'item.
            if ($null -ne $pendingSubsection) {
                Format-Subsection $sel $pendingSubsection.Short
                if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
                $pendingSubsection = $null
            }
            if ($null -ne $pendingIntro) {
                & $emitIntro $pendingIntro
                $pendingIntro = $null
            }
            & $emitItem $el
        }
    }
}

# Escriu el bloc de conclusions:
#   - $headerText : titol del bloc (sol ser 'CONCLUSIONS'), centrat-negreta.
#                   '' = no s'emet.
#   - $conclusions : array d'objectes {Title; Body} de les conclusions
#                    TRIADES al Pas 5. Es emet el seu Body.
#   - $alwaysConclusions : array de cadenes ja sense el prefix '::SEMPRE::'
#                          que s'emeten sempre, despres de les triades.
#   - $fields : per resoldre [CAMP:] i [OPCIO:] dins els textos.
# La separacio entre conclusions la posa Format-Conclusion via SpaceAfter
# (ConclusionSpaceAfterPt), aixi que aqui no hi fa falta un Spacer entre.
function _WriteConclusionsBlock($sel, $cfg, $headerText, $conclusions, $alwaysConclusions, $fields) {
    $hasBody = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    $hasHead = -not [string]::IsNullOrWhiteSpace($headerText)
    if (-not $hasBody -and -not $hasHead) { return }

    if ($cfg.SpacerBeforeConclusionsBlock) { Format-Spacer $sel }

    if ($hasHead) {
        Format-ConclusionHeader $sel $headerText
    }

    foreach ($c in $conclusions) {
        $txt = if ($c -is [string]) { $c } else { [string]$c.Body }
        $resolved = Apply-Fields -text $txt -fields $fields
        Format-Conclusion $sel $resolved
    }
    foreach ($a in $alwaysConclusions) {
        $resolved = Apply-Fields -text ([string]$a) -fields $fields
        Format-Conclusion $sel $resolved
    }
}

function Build-Document($word, $header, $selectedSections, $fields, $conclusions, $alwaysConclusions, $catalegName, $introText, $conclusionsHeaderText) {
    $baseName  = _GetOutputFileName $catalegName $header['ID_GIA']
    $targetDir = _ResolveOutputDir
    # Triem el primer nom lliure al directori de sortida (afegim _2, _3...
    # si ja existeix). Aixi pots generar diversos informes del mateix dia/GIA
    # sense que cap es sobreescrigui.
    $outPath  = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)

    # Treballem amb una copia LOCAL (a %TEMP%) per evitar que Word obri el
    # fitxer en "Vista protegida" quan el desti es una unitat de xarxa.
    # El temp porta el mateix nom (ja unic) que el desti final.
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)  # wdStory = 6

    $cfg = $Script:ReportFormatConfig
    _WriteCatalegBody $sel $cfg $selectedSections $fields $introText
    _WriteConclusionsBlock $sel $cfg $conclusionsHeaderText $conclusions $alwaysConclusions $fields

    $doc.Save()
    $doc.Close($false)

    # Movem el fitxer al desti final (xarxa o local segons disponibilitat).
    try {
        Move-Item -LiteralPath $tempPath -Destination $outPath -Force
    } catch {
        return $tempPath
    }
    return $outPath
}

# ----------------------------------------------------------------------------
# Main flow
# ----------------------------------------------------------------------------
function Main {
    if (-not (Test-Path $HeaderPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat $HeaderPath",'Error','OK','Error') | Out-Null
        exit 1
    }

    # Pas 0: tria de mode. "seguiment" es un flux autonom (Seguiment.ps1) que
    # gestiona la seva propia app de Word; el wizard de 6 passos de sota queda
    # intacte per al mode "nou".
    $mode = Select-Mode
    if ($mode -eq 'seguiment') {
        Invoke-SeguimentFlow
        return
    }

    # Assistent navegable. Cada pas (dialeg) retorna un objecte amb:
    #   Nav  = 'next' | 'back' | 'stay'
    #   Data = el resultat del pas (si 'next')
    # Tancar la finestra (X) avorta el programa (exit 0). El boto "Enrere"
    # retorna 'back' i Main torna al pas anterior conservant les dades ja
    # introduides (es passen com a precarrega).
    #
    # $st  : dades confirmades de cada pas (es mantenen en memoria al navegar).
    # $pre : precarregues per a cada pas (de l'estat o de "Recuperar ultim").
    $st  = @{ Cataleg=$null; Parsed=$null; Header=$null; Selected=$null; Fields=$null; ConclAll=$null; Conclusions=$null }
    $pre = @{ Cat=$null; Header=$null; Keys=$null; Fields=$null; Concl=$null }

    $word = New-WordApp
    try {
        $step = 1
        $dir  = 'fwd'
        while ($step -ge 1 -and $step -le 6) {
            switch ($step) {

                1 {
                    $r = Select-Cataleg -preloadBaseName $pre.Cat
                    # El Pas 1 no te 'back'; nomes 'next' o tancar (exit dins la funcio).
                    $st.Cataleg = $r.Data
                    $st.Parsed  = Get-ParsedCataleg -word $word -path $st.Cataleg.FullName
                    $pre.Cat    = $st.Cataleg.BaseName
                    $step = 2; $dir = 'fwd'
                }

                2 {
                    $r = Get-HeaderData -preload $pre.Header
                    if ($r.Nav -eq 'back') { exit 0 }   # enrere des del Pas 2 = sortir
                    else {
                        $st.Header  = $r.Data
                        $pre.Header = $r.Data
                        if ($r.Recovered) {
                            # "Recuperar dades ultim informe": precarreguem la
                            # resta de passos amb les dades de l'ultim informe.
                            $pre.Keys   = $r.Recovered.SelectedKeys
                            $pre.Fields = $r.Recovered.FieldValues
                            $pre.Concl  = $r.Recovered.ConclusionTexts
                        }
                        $step = 3; $dir = 'fwd'
                    }
                }

                3 {
                    $r = Select-Items -sections $st.Parsed.Sections -preloadSelectedKeys $pre.Keys
                    if     ($r.Nav -eq 'back') { $step = 2; $dir = 'back' }
                    elseif ($r.Nav -eq 'stay') { }   # cap seleccio: es torna a mostrar
                    else {
                        $st.Selected = $r.Data
                        $pre.Keys    = Get-SelectedKeysFromResult $st.Selected
                        $step = 4; $dir = 'fwd'
                    }
                }

                # Ordre: Pas 4 = CONCLUSIONS (abans dels camps), Pas 5 = CAMPS.
                # Es necessari perque les conclusions poden contenir [CAMP:] /
                # [OPCIO:]. Cal triar conclusions PRIMER per saber quins camps
                # cal demanar.
                4 {
                    if ($null -eq $st.ConclAll) {
                        $st.ConclAll = Read-Conclusions -word $word -path $ConclusionsPath
                    }
                    if ($st.ConclAll.Selectable.Count -eq 0) {
                        # No hi ha conclusions triables: saltem el pas.
                        $st.Conclusions = @()
                        if ($dir -eq 'back') { $step = 3 } else { $step = 5 }
                    } else {
                        $r = Select-Conclusions -conclusions $st.ConclAll.Selectable -preloadTitles $pre.Concl
                        if ($r.Nav -eq 'back') { $step = 3; $dir = 'back' }
                        else {
                            $st.Conclusions = $r.Data
                            # $pre.Concl guarda nomes els TITOLS per a la
                            # precarrega de la propera vegada.
                            $pre.Concl = @($st.Conclusions | ForEach-Object { $_.Title })
                            $step = 5; $dir = 'fwd'
                        }
                    }
                }

                5 {
                    # Camps: detectem [CAMP:] / [OPCIO:] tant al REQ1 (items
                    # seleccionats) com a les conclusions triades i fixes.
                    $fields = Get-FieldsFromSelection $st.Selected
                    Add-FieldsFromConclusions $fields $st.Conclusions $st.ConclAll.Always
                    if ($fields.Count -eq 0) {
                        $st.Fields = $fields
                        if ($dir -eq 'back') { $step = 4 } else { $step = 6 }
                    } else {
                        $r = Prompt-Fields -fields $fields -preloadValues $pre.Fields
                        if ($r.Nav -eq 'back') { $step = 4; $dir = 'back' }
                        else {
                            $st.Fields  = $r.Data
                            $pre.Fields = Get-FieldValuesForSession $st.Fields
                            $step = 6; $dir = 'fwd'
                        }
                    }
                }

                6 {
                    $outPath = Build-Document -word $word -header $st.Header `
                                              -selectedSections $st.Selected `
                                              -fields $st.Fields `
                                              -conclusions $st.Conclusions `
                                              -alwaysConclusions $st.ConclAll.Always `
                                              -catalegName $st.Cataleg.BaseName `
                                              -introText $st.Parsed.IntroText `
                                              -conclusionsHeaderText $st.ConclAll.HeaderText

                    # Desem les dades per poder replicar aquest informe mes endavant.
                    # Per a 'ConclusionTexts' guardem els TITOLS triats (es el que
                    # fa servir Select-Conclusions per precarregar).
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = $st.Cataleg.BaseName
                        Header          = $st.Header
                        SelectedKeys    = (Get-SelectedKeysFromResult $st.Selected)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @($st.Conclusions | ForEach-Object { $_.Title })
                    })

                    [System.Windows.Forms.MessageBox]::Show(
                        "Informe generat:`n$outPath",
                        'Finalitzat', 'OK', 'Information') | Out-Null

                    # Obrim Word en primer pla per a l'usuari
                    $word.Visible = $true
                    $word.Documents.Open($outPath) | Out-Null
                    $step = 99   # surt del bucle
                }
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)",'Error','OK','Error') | Out-Null
        throw
    }
    finally {
        if (-not $word.Visible) { Close-WordApp $word }
    }
}

if (-not $Script:HeadlessTest) { Main }
