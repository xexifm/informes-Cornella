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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptRoot      = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem el modul de format (Format.ps1). Conte les funcions Format-Section,
# Format-Item, etc. i $ReportFormatConfig. Reutilitzable per altres tipus
# d'informes.
. (Join-Path $ScriptRoot 'Format.ps1')

$EstructuralsDir = Join-Path $ScriptRoot 'ESTRUCTURALS'
$HeaderPath      = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
$ConclusionsPath = Join-Path $EstructuralsDir '0 CONCLUSIONS.docx'

# ----------------------------------------------------------------------------
# Configuracio per defecte. Es pot sobreescriure des de config.ps1 (opcional)
# al costat del .ps1.
# ----------------------------------------------------------------------------
$OutputDir              = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\0_Plantilles\Powershell\Informes generats'
$ActivitatsDir          = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
$AlwaysConclusionsCount = 2

$configPath = Join-Path $ScriptRoot 'config.ps1'
if (Test-Path -LiteralPath $configPath) {
    . $configPath
}

# Estat persistent (sessio). Es guarda a %LOCALAPPDATA% per no embrutar el
# repositori i no haver de tocar .gitignore.
$AppDataDir  = Join-Path $env:LOCALAPPDATA 'InformesCornella'
$SessionPath = Join-Path $AppDataDir 'session.json'

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
# Persistencia de sessio
# ----------------------------------------------------------------------------
# Format del session.json (versio 1):
#   {
#     "Version": 1,
#     "Timestamp": "<ISO 8601>",
#     "CatalegBaseName": "REQ1",
#     "Header": { "ID_GIA": "...", ... },
#     "SelectedKeys": [ "SectionTitle::ItemShort", ... ],
#     "FieldValues":  { "nom": "valor", ... },
#     "ConclusionTexts": [ "text1", ... ]
#   }
# Cada camp es opcional: nomes hi son els passos completats.

function Save-Session($state) {
    try {
        Ensure-AppDataDir
        $state.Version   = 1
        $state.Timestamp = (Get-Date).ToString('o')
        ($state | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $SessionPath -Encoding UTF8
    } catch {
        # Si no podem desar la sessio, no es un error fatal. Continuem en silenci.
    }
}

function Load-Session {
    if (-not (Test-Path -LiteralPath $SessionPath)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $SessionPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Clear-Session {
    if (Test-Path -LiteralPath $SessionPath) {
        try { Remove-Item -LiteralPath $SessionPath -Force } catch { }
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

function Find-LatestActivitatsExcel {
    if (-not (Test-Path -LiteralPath $ActivitatsDir)) { return $null }
    $regex = '^(\d{4}-\d{2}-\d{2})\s+ACTIVITATS\.(xls|xlsx)$'
    $candidates = Get-ChildItem -LiteralPath $ActivitatsDir -File -ErrorAction SilentlyContinue |
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

# Normalitza un text Unicode (sense diacritics, minuscules) per a comparacio.
function _NormalizeText($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    return (($t -replace '\p{Mn}','').ToLower().Trim())
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
#   ById        : hashtable [string ID] -> hashtable @{ TITULAR; ADRECA; ACTIVITAT }
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

            $warnings = _ValidateActivitatsHeaders $data $rows $cols

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
                $byId[$id] = @{
                    TITULAR   = $rao
                    ADRECA    = $adreca
                    ACTIVITAT = $actPrin
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
    if ($catalegs.Count -eq 1) { return $catalegs[0] }

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

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(380, 240)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    if ($form.ShowDialog() -ne 'OK') { exit 0 }
    return $catalegs[$list.SelectedIndex]
}

# ----------------------------------------------------------------------------
# Step 2 - Header data (formulari + precarrega Excel)
# ----------------------------------------------------------------------------
# Construeix el formulari de capcalera (controls + botons), retorna la
# tupla amb el formulari, el diccionari de controls i el boto Cercar perque
# Get-HeaderData hi puga lligar la logica de cerca i validacio.
function _BuildHeaderForm($excelFileLabel) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lblBd = New-Object System.Windows.Forms.Label
    $lblBd.Text = $excelFileLabel
    $lblBd.Location = New-Object System.Drawing.Point(15, 12)
    $lblBd.Size = New-Object System.Drawing.Size(680, 22)
    $lblBd.ForeColor = [System.Drawing.Color]::DarkBlue
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

    & $addRow "Num. d'expedient"             $y 460 'EXP_NUM';      $y += 38
    & $addRow 'Titular (autom., editable)'   $y 460 'TITULAR';      $y += 38
    & $addRow 'Adreca (autom., editable)'    $y 460 'ADRECA';       $y += 38
    & $addRow 'Activitat (autom., editable)' $y 460 'ACTIVITAT';    $y += 38
    & $addRow "Num. d'anotacio (Objecte)"    $y 460 'NUM_ANOTACIO'; $y += 38
    & $addRow "Data d'anotacio (dd/mm/aaaa)" $y 460 'DATA_ANOTACIO';$y += 50

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(505, $y)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(605, $y)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $form.CancelButton = $cancel
    [void]$form.Controls.Add($cancel)

    return @{ Form=$form; Controls=$controls; BtnSearch=$btnSearch; BtnOk=$ok }
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
function _PreloadHeaderControls($controls, $preload) {
    if ($null -eq $preload) { return }
    foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','NUM_ANOTACIO','DATA_ANOTACIO') {
        if ($preload.PSObject.Properties.Name -contains $k -and $null -ne $preload.$k) {
            $controls[$k].Text = [string]$preload.$k
        }
    }
}

function Get-HeaderData {
    param($preload = $null)

    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap fitxer 'YYYY-MM-DD ACTIVITATS.xls' a:`n$ActivitatsDir",
            'Base de dades no trobada', 'OK', 'Error') | Out-Null
        exit 1
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

    $label = "Base de dades d'activitats: $($latest.File.Name)  (data: $($latest.Date.ToString('yyyy-MM-dd'))) - $($actCache.ById.Count) activitats carregades"
    $f = _BuildHeaderForm $label
    $form      = $f.Form
    $controls  = $f.Controls
    $btnSearch = $f.BtnSearch
    $ok        = $f.BtnOk

    _PreloadHeaderControls $controls $preload

    # Cerca per ID GIA: instantania des del cache. Omple els 3 camps autom.
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
            $controls['TITULAR'].Text = ''
            $controls['ADRECA'].Text = ''
            $controls['ACTIVITAT'].Text = ''
            return $false
        }
        $controls['TITULAR'].Text   = $act['TITULAR']
        $controls['ADRECA'].Text    = $act['ADRECA']
        $controls['ACTIVITAT'].Text = $act['ACTIVITAT']
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

    if ($form.ShowDialog() -ne 'OK') { exit 0 }
    return $script:_headerData
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
            $isH1 = ($styleName -match '^(Heading 1|Titol 1|Titulo 1|T.tulo 1)$')
            $isH2 = ($styleName -match '^(Heading 2|Titol 2|Titulo 2|T.tulo 2)$')

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

            # Paragraf Normal: l'afegim al BodyLines de l'element actiu.
            if ($null -eq $lastH2) {
                if ($null -eq $currentSection -and [string]::IsNullOrWhiteSpace($introText)) {
                    $introText = $text
                }
                continue
            }
            $target = $lastH2
            if ($lastH2.Kind -eq 'item' -and $lastH2.Children.Count -gt 0) {
                $target = $lastH2.Children[$lastH2.Children.Count - 1]
            }
            [void]$target.BodyLines.Add($text)
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
                # Item
                $itNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                $itNode.Tag = @{ Kind = 'Item'; Ref = $n.El; SectionTitle = $sec.Title }
                $itKey = (_ItemKey $sec.Title $n.El.Short)
                if ($checkStates.ContainsKey($itKey)) { $itNode.Checked = $checkStates[$itKey] }
                [void]$container.Nodes.Add($itNode)
                foreach ($ch in $n.ChildShows) {
                    $chNode = New-Object System.Windows.Forms.TreeNode($ch.Short)
                    $chNode.Tag = @{ Kind = 'Child'; Ref = $ch; SectionTitle = $sec.Title; ParentShort = $n.El.Short }
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

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(890, 640)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(980, 640)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $cancel.Anchor = 'Bottom, Right'
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    if ($form.ShowDialog() -ne 'OK') { exit 0 }

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
        exit 0
    }
    return $result
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
# Step 4 - Field placeholders [CAMP: nom (hint)]
# ----------------------------------------------------------------------------
$Script:CampRegex = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'

function Get-FieldsFromSelection($selectedSections) {
    $fields = [ordered]@{}
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            $allText = ($it.BodyLines -join ' ')
            foreach ($ch in $it.Children) {
                $allText += ' ' + ($ch.BodyLines -join ' ')
            }
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
                    $fields[$name] = [pscustomobject]@{ Name = $name; Hint = $hint; Value = '' }
                }
            }
        }
    }
    return $fields
}

function Prompt-Fields {
    param($fields, $preloadValues = $null)
    if ($fields.Count -eq 0) { return $fields }

    # Precarrega valors anteriors (per nom de camp)
    if ($preloadValues) {
        foreach ($name in $fields.Keys) {
            $v = $null
            if ($preloadValues -is [hashtable] -and $preloadValues.ContainsKey($name)) {
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
    $textboxes = @{}
    foreach ($name in $fields.Keys) {
        $f = $fields[$name]
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $name
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(220, 22)
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
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(15, $y)
        $tb.Size = New-Object System.Drawing.Size(520, 22)
        $tb.Text = $f.Value
        $form.Controls.Add($tb)
        $textboxes[$name] = $tb
        $y += 32
    }

    $form.ClientSize = New-Object System.Drawing.Size(560, [Math]::Min(640, ($y + 70)))

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(360, $y)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(450, $y)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    if ($form.ShowDialog() -ne 'OK') { exit 0 }
    foreach ($name in $fields.Keys) { $fields[$name].Value = $textboxes[$name].Text }
    return $fields
}

function Apply-Fields($text, $fields) {
    return $Script:CampRegex.Replace($text, {
        param($m)
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) { $name = $raw.Substring(0, $parenIdx).Trim() }
        if ($fields.Contains($name)) { return $fields[$name].Value }
        return ''
    })
}

# Extreu els valors dels camps en un hashtable simple per a la sessio.
function Get-FieldValuesForSession($fields) {
    $h = @{}
    foreach ($name in $fields.Keys) { $h[$name] = $fields[$name].Value }
    return $h
}

# ----------------------------------------------------------------------------
# Step 5 - Conclusions
# ----------------------------------------------------------------------------
function Read-Conclusions($word, $path) {
    # Retorna un PSCustomObject amb dues llistes:
    #   Selectable : els paragrafs que apareixen al Pas 5 com a checkboxes.
    #   Always     : els darrers $AlwaysConclusionsCount paragrafs, que
    #                s'inclouen sempre al document final sense preguntar.
    $empty = [pscustomobject]@{ Selectable = @(); Always = @() }
    if (-not (Test-Path -LiteralPath $path)) { return $empty }
    $doc = $word.Documents.Open($path, $false, $true)
    try {
        $list = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            $t = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$list.Add($t) }
        }
        $n = $list.Count
        $alwaysN = [Math]::Min($AlwaysConclusionsCount, $n)
        $selectable = @()
        $always = @()
        if ($alwaysN -gt 0) {
            $selectable = $list.GetRange(0, $n - $alwaysN).ToArray()
            $always     = $list.GetRange($n - $alwaysN, $alwaysN).ToArray()
        } else {
            $selectable = $list.ToArray()
        }
        return [pscustomobject]@{ Selectable = $selectable; Always = $always }
    } finally {
        $doc.Close($false)
    }
}

function Select-Conclusions {
    param($conclusions, $preloadTexts = $null)
    if ($conclusions.Count -eq 0) { return @() }

    # Convertim preloadTexts a un HashSet per a comparacio rapida.
    $preloadSet = New-Object System.Collections.Generic.HashSet[string]
    if ($preloadTexts) { foreach ($t in $preloadTexts) { [void]$preloadSet.Add([string]$t) } }

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
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $conclusions[$i]
        $cb.Location = New-Object System.Drawing.Point(5, $y)
        $cb.Size = New-Object System.Drawing.Size(700, 60)
        $cb.AutoSize = $false
        if ($preloadSet.Contains([string]$conclusions[$i])) { $cb.Checked = $true }
        $panel.Controls.Add($cb)
        $checks += $cb
        $y += 65
    }

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Generar'
    $ok.Location = New-Object System.Drawing.Point(570, 480)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(660, 480)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    if ($form.ShowDialog() -ne 'OK') { exit 0 }

    $selected = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $checks.Count; $i++) {
        if ($checks[$i].Checked) { [void]$selected.Add($conclusions[$i]) }
    }
    return ,$selected.ToArray()
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

# Determina el directori de sortida: l'$OutputDir si es accessible, en cas
# contrari una subcarpeta 'Informes generats' al costat del .ps1.
function _ResolveOutputDir {
    $targetDir = $OutputDir
    try {
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
        return $targetDir
    } catch {
        $local = Join-Path $ScriptRoot 'Informes generats'
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

    # Resol [CAMP: ...] a cada linia. Retorna SEMPRE un array.
    $resolveLines = {
        param($lines)
        $arr = New-Object System.Collections.ArrayList
        foreach ($ln in $lines) {
            [void]$arr.Add( (Apply-Fields -text $ln -fields $fields) )
        }
        return ,@($arr.ToArray())
    }

    $emitExtras = {
        param($lines, $isChild)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^https?://') {
                if ($isChild) { Format-Url $sel $line -IsChild } else { Format-Url $sel $line }
            } else {
                if ($isChild) { Format-Body $sel $line -IsChild } else { Format-Body $sel $line }
            }
        }
    }

    $emitIntro = {
        param($introEl)
        $lines = @(& $resolveLines $introEl.BodyLines)
        foreach ($bp in $lines) {
            if ([string]::IsNullOrWhiteSpace($bp)) { continue }
            if ($bp -match '^https?://') { Format-Url $sel $bp } else { Format-Body $sel $bp }
        }
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
                Format-Item $sel "$($script:_buildGlobal)." $itemLines[0]
                & $emitExtras $itemLines $false
                $itemWritten = $true
            }
        }
        if ($hasChildren) {
            $subCounter = 0
            foreach ($ch in $it.Children) {
                $childLines = @(& $resolveLines $ch.BodyLines)
                if ($childLines.Count -eq 0) { continue }
                $subCounter++
                if (-not $itemWritten) {
                    $script:_buildGlobal++
                    $itemWritten = $true
                }
                Format-Item $sel "$($script:_buildGlobal).$subCounter." $childLines[0] -IsChild
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

        $pendingIntro = $null
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'subsection') {
                Format-Subsection $sel $el.Short
                if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
                $pendingIntro = $null
                continue
            }
            if ($el.Kind -eq 'intro') {
                $pendingIntro = $el
                continue
            }
            if ($null -ne $pendingIntro) {
                & $emitIntro $pendingIntro
                $pendingIntro = $null
            }
            & $emitItem $el
        }
    }
}

# Escriu el bloc de conclusions (les triades + les "sempre incloses").
function _WriteConclusionsBlock($sel, $cfg, $conclusions, $alwaysConclusions) {
    $hasConcl = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    if (-not $hasConcl) { return }
    if ($cfg.SpacerBeforeConclusionsBlock) { Format-Spacer $sel }
    $allConcl = @($conclusions) + @($alwaysConclusions)
    $totalC = $allConcl.Count
    for ($i = 0; $i -lt $totalC; $i++) {
        Format-Conclusion $sel $allConcl[$i]
        if ($cfg.SpacerBetweenConclusions -and $i -lt ($totalC - 1)) {
            Format-Spacer $sel
        }
    }
}

function Build-Document($word, $header, $selectedSections, $fields, $conclusions, $alwaysConclusions, $catalegName, $introText) {
    $fileName  = _GetOutputFileName $catalegName $header['ID_GIA']
    $targetDir = _ResolveOutputDir
    $outPath   = Join-Path $targetDir $fileName

    # Treballem amb una copia LOCAL (a %TEMP%) per evitar que Word obri el
    # fitxer en "Vista protegida" quan el desti es una unitat de xarxa.
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)  # wdStory = 6

    $cfg = $Script:ReportFormatConfig
    _WriteCatalegBody $sel $cfg $selectedSections $fields $introText
    _WriteConclusionsBlock $sel $cfg $conclusions $alwaysConclusions

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

    # Sessio anterior? Si existeix, ofereix recuperar les dades.
    $preload = $null
    $prevSession = Load-Session
    if ($null -ne $prevSession) {
        $ts = if ($prevSession.Timestamp) { $prevSession.Timestamp } else { '(sense data)' }
        $r = [System.Windows.Forms.MessageBox]::Show(
            "S'ha detectat una sessio anterior ($ts) que potser no es va completar.`n`nVols precarregar les seves dades als formularis?",
            'Recuperar sessio anterior', 'YesNo', 'Question')
        if ($r -eq 'Yes') { $preload = $prevSession }
    }

    # Estat acumulat que es va desant despres de cada pas.
    $sessionState = [ordered]@{
        Version         = 1
        Timestamp       = (Get-Date).ToString('o')
        CatalegBaseName = $null
        Header          = $null
        SelectedKeys    = $null
        FieldValues     = $null
        ConclusionTexts = $null
    }

    $cataleg = Select-Cataleg -preloadBaseName ($preload.CatalegBaseName)
    $sessionState.CatalegBaseName = $cataleg.BaseName
    Save-Session $sessionState

    $header = Get-HeaderData -preload ($preload.Header)
    $sessionState.Header = $header
    Save-Session $sessionState

    $word = New-WordApp
    try {
        $parsed       = Get-ParsedCataleg -word $word -path $cataleg.FullName
        $selected     = Select-Items -sections $parsed.Sections -preloadSelectedKeys ($preload.SelectedKeys)
        $sessionState.SelectedKeys = Get-SelectedKeysFromResult $selected
        Save-Session $sessionState

        $fields       = Get-FieldsFromSelection $selected
        $fields       = Prompt-Fields -fields $fields -preloadValues ($preload.FieldValues)
        $sessionState.FieldValues = Get-FieldValuesForSession $fields
        Save-Session $sessionState

        $conclusionsAll = Read-Conclusions -word $word -path $ConclusionsPath
        $conclusions  = Select-Conclusions -conclusions $conclusionsAll.Selectable -preloadTexts ($preload.ConclusionTexts)
        $sessionState.ConclusionTexts = @($conclusions)
        Save-Session $sessionState

        $outPath      = Build-Document -word $word -header $header `
                                       -selectedSections $selected `
                                       -fields $fields `
                                       -conclusions $conclusions `
                                       -alwaysConclusions $conclusionsAll.Always `
                                       -catalegName $cataleg.BaseName `
                                       -introText $parsed.IntroText

        # Generacio completada: ja no necessitem la sessio anterior.
        Clear-Session

        [System.Windows.Forms.MessageBox]::Show(
            "Informe generat:`n$outPath",
            'Finalitzat', 'OK', 'Information') | Out-Null

        # Obrim Word en primer pla per a l'usuari
        $word.Visible = $true
        $word.Documents.Open($outPath) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)",'Error','OK','Error') | Out-Null
        throw
    }
    finally {
        if (-not $word.Visible) { Close-WordApp $word }
    }
}

Main
