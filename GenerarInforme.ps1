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
# Ruta de sortida principal (xarxa). Si no es accessible, cau a una carpeta
# local 'Informes generats' al costat del .ps1.
$OutputDir       = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\0_Plantilles\Powershell\Informes generats'
# Directori de bases de dades d'activitats (Excel). El nom del fitxer ha de
# seguir el patro "YYYY-MM-DD ACTIVITATS.xls" o ".xlsx".
$ActivitatsDir   = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
# Quantes conclusions del final del fitxer 0 CONCLUSIONS.docx s'inclouen
# sempre al document final (no apareixen al Pas 5).
$AlwaysConclusionsCount = 2

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
# Activitats Excel database
# ----------------------------------------------------------------------------
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

# Cerca una activitat pel seu ID a la fulla "Estes"/"Estès" del fitxer Excel.
# Retorna un hashtable amb TITULAR, ADRECA, ACTIVITAT o $null si no es troba.
function Get-ActivitatByID($excelFile, $idGia) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            # Normalitza el nom de la fulla per acceptar variants com "Estes",
            # "Estès" (precompost), o "Estès" (descompost amb caracter combinant).
            $sh = $null
            $sheetNames = @()
            foreach ($s in $wb.Sheets) {
                $sheetNames += $s.Name
                $n = $s.Name.Normalize([System.Text.NormalizationForm]::FormD)
                $n = ($n -replace '\p{Mn}','').ToLower().Trim()
                if ($n -eq 'estes') { $sh = $s; break }
            }
            if ($null -eq $sh) {
                throw "No s'ha trobat la fulla 'Estes'/'Estès' al fitxer Excel. Fulles disponibles: $($sheetNames -join ', ')"
            }
            $used = $sh.UsedRange
            $data = $used.Value2
            if ($null -eq $data) { return $null }
            # Columnes Excel (1-based) segons la convencio del fitxer:
            #   1  = ID Activitat
            #   10 = Rao social
            #   48 = Emp. Tipus via
            #   49 = Emp. Carrer
            #   50 = Emp. Numero
            #   52 = Emp. Lletra
            #   55 = Emp. Pis
            #   56 = Emp. Porta
            #   94 = Activitat principal
            $rows = $data.GetLength(0)
            $idTarget = [string]$idGia
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, 1]
                if ($null -eq $cell) { continue }
                # ID pot ser numeric o string; normalitzem.
                $id = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                if ($id -eq $idTarget) {
                    $get = {
                        param($c)
                        $v = $data[$r, $c]
                        if ($null -eq $v) { return '' }
                        return ([string]$v).Trim()
                    }
                    $tipusVia = & $get 48
                    $carrer   = & $get 49
                    $numero   = & $get 50
                    $lletra   = & $get 52
                    $pis      = & $get 55
                    $porta    = & $get 56
                    $rao      = & $get 10
                    $actPrin  = & $get 94
                    $parts = @($tipusVia, $carrer, $numero, $lletra, $pis, $porta) |
                        Where-Object { $_ -and $_.Trim() -ne '' }
                    $adreca = ($parts -join ' ') + ', CORNELLÀ DE LLOBREGAT'
                    return @{
                        TITULAR   = $rao
                        ADRECA    = $adreca
                        ACTIVITAT = $actPrin
                    }
                }
            }
            return $null
        } finally {
            $wb.Close($false)
        }
    } finally {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
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
# Step 2 - Header data
# ----------------------------------------------------------------------------
function Get-HeaderData {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap fitxer 'YYYY-MM-DD ACTIVITATS.xls' a:`n$ActivitatsDir",
            'Base de dades no trobada', 'OK', 'Error') | Out-Null
        exit 1
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lblBd = New-Object System.Windows.Forms.Label
    $lblBd.Text = "Base de dades d'activitats: $($latest.File.Name)  (data: $($latest.Date.ToString('yyyy-MM-dd')))"
    $lblBd.Location = New-Object System.Drawing.Point(15, 12)
    $lblBd.Size = New-Object System.Drawing.Size(680, 22)
    $lblBd.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($lblBd)

    # Helper per afegir una fila etiqueta + TextBox. Guarda el TextBox a
    # $controls[$key]. Suprimim qualsevol output al pipeline.
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
    # 1) ID GIA + boto "Cercar"
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

    # Botons
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

    # Cercar a l'Excel pel ID GIA i omplir els 3 camps autom.
    $doSearch = {
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return $false
        }
        try {
            $act = Get-ActivitatByID -excelFile $latest.File -idGia $idGia
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error llegint l'Excel:`n$($_.Exception.Message)",'Error','OK','Error') | Out-Null
            return $false
        }
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

    # Validacio al "Seguent": ID GIA i 3 camps autom. han de tenir valor.
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
            # Cercar automaticament si encara no s'ha fet
            if (-not (& $doSearch)) { return }
        }
        $data = @{}
        $data['ID_GIA']        = $controls['ID_GIA'].Text.Trim()
        $data['EXP_NUM']       = $controls['EXP_NUM'].Text.Trim()
        $data['TITULAR']       = $controls['TITULAR'].Text.Trim()
        $data['ADRECA']        = $controls['ADRECA'].Text.Trim()
        $data['ACTIVITAT']     = $controls['ACTIVITAT'].Text.Trim()
        $data['NUM_ANOTACIO']  = $controls['NUM_ANOTACIO'].Text.Trim()
        $data['DATA_ANOTACIO'] = $controls['DATA_ANOTACIO'].Text.Trim()
        $script:_headerData = $data
        $form.DialogResult = 'OK'
        $form.Close()
    })

    if ($form.ShowDialog() -ne 'OK') { exit 0 }
    return $script:_headerData
}

# ----------------------------------------------------------------------------
# Step 3 - Parse cataleg (Heading 1 / Heading 2 / Normal)
# ----------------------------------------------------------------------------
function Parse-Cataleg($word, $path) {
    $doc = $word.Documents.Open($path, $false, $true)  # ReadOnly
    try {
        $sections = New-Object System.Collections.ArrayList
        $currentSection = $null
        $currentItem    = $null

        foreach ($p in $doc.Paragraphs) {
            $text = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            $styleName = ''
            try { $styleName = $p.Style.NameLocal } catch { }
            # Acceptem tant noms anglesos com locals (Titol 1, Titulo 1).
            $isH1 = ($styleName -match '^(Heading 1|Titol 1|Titulo 1|T.tulo 1)$')
            $isH2 = ($styleName -match '^(Heading 2|Titol 2|Titulo 2|T.tulo 2)$')

            if ($isH1) {
                $currentSection = [pscustomobject]@{
                    Title = $text
                    Items = New-Object System.Collections.ArrayList
                }
                [void]$sections.Add($currentSection)
                $currentItem = $null
                continue
            }

            if ($isH2) {
                if ($null -eq $currentSection) {
                    # Heading 2 sense seccio: creem una de generica.
                    $currentSection = [pscustomobject]@{
                        Title = '(Sense seccio)'
                        Items = New-Object System.Collections.ArrayList
                    }
                    [void]$sections.Add($currentSection)
                }
                $isChild = $false
                $short = $text
                if ($short -like '::CHILD::*') {
                    $isChild = $true
                    $short = $short.Substring('::CHILD::'.Length).Trim()
                }
                $newItem = [pscustomobject]@{
                    Short    = $short
                    BodyLines = New-Object System.Collections.ArrayList
                    Children = New-Object System.Collections.ArrayList
                    IsChild  = $isChild
                }
                if ($isChild -and $null -ne $currentItem) {
                    [void]$currentItem.Children.Add($newItem)
                } else {
                    [void]$currentSection.Items.Add($newItem)
                    $currentItem = $newItem
                }
                continue
            }

            # Normal / body: l'afegim a l'item actiu (pare o ultim fill).
            $target = $null
            if ($null -ne $currentItem) {
                if ($currentItem.Children.Count -gt 0) {
                    $target = $currentItem.Children[$currentItem.Children.Count - 1]
                } else {
                    $target = $currentItem
                }
            }
            if ($null -ne $target) { [void]$target.BodyLines.Add($text) }
        }
        return $sections
    }
    finally {
        $doc.Close($false)
    }
}

# ----------------------------------------------------------------------------
# Step 3 (UI) - TreeView with checkboxes
# ----------------------------------------------------------------------------
function Select-Items($sections) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 3 - Seleccio de deficiencies'
    $form.Size = New-Object System.Drawing.Size(1100, 720)
    $form.StartPosition = 'CenterScreen'

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(10, 10)
    $tv.Size = New-Object System.Drawing.Size(1060, 620)
    $tv.CheckBoxes = $true
    $tv.HideSelection = $false
    $tv.ShowNodeToolTips = $true
    $tv.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($tv)

    foreach ($sec in $sections) {
        $secNode = New-Object System.Windows.Forms.TreeNode($sec.Title)
        $secNode.Tag = @{ Kind = 'Section'; Ref = $sec }
        $secNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        [void]$tv.Nodes.Add($secNode)
        foreach ($it in $sec.Items) {
            $itNode = New-Object System.Windows.Forms.TreeNode($it.Short)
            $itNode.Tag = @{ Kind = 'Item'; Ref = $it }
            [void]$secNode.Nodes.Add($itNode)
            foreach ($ch in $it.Children) {
                $chNode = New-Object System.Windows.Forms.TreeNode($ch.Short)
                $chNode.Tag = @{ Kind = 'Child'; Ref = $ch; Parent = $it }
                [void]$itNode.Nodes.Add($chNode)
            }
        }
        $secNode.Expand()
    }

    # Propagation: marcar una seccio marca tots els seus items; marcar un
    # item marca els seus fills. (Senzill i previsible.)
    $tv.add_AfterCheck({
        param($sender, $e)
        if ($script:_propagating) { return }
        $script:_propagating = $true
        try {
            foreach ($child in $e.Node.Nodes) {
                $child.Checked = $e.Node.Checked
                foreach ($gc in $child.Nodes) { $gc.Checked = $e.Node.Checked }
            }
        } finally {
            $script:_propagating = $false
        }
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

    # Construim la llista de seccions escollides amb nomes els items marcats.
    $result = New-Object System.Collections.ArrayList
    foreach ($secNode in $tv.Nodes) {
        $secData = $secNode.Tag.Ref
        $chosenItems = New-Object System.Collections.ArrayList
        foreach ($itNode in $secNode.Nodes) {
            $it = $itNode.Tag.Ref
            $chosenChildren = New-Object System.Collections.ArrayList
            foreach ($chNode in $itNode.Nodes) {
                if ($chNode.Checked) { [void]$chosenChildren.Add($chNode.Tag.Ref) }
            }
            $itemSelected = $itNode.Checked -or ($chosenChildren.Count -gt 0)
            if ($itemSelected) {
                [void]$chosenItems.Add([pscustomobject]@{
                    Short     = $it.Short
                    BodyLines = $it.BodyLines
                    Children  = $chosenChildren
                    Selected  = [bool]$itNode.Checked
                })
            }
        }
        if ($chosenItems.Count -gt 0) {
            [void]$result.Add([pscustomobject]@{
                Title = $secData.Title
                Items = $chosenItems
            })
        }
    }
    if ($result.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No s''ha seleccionat cap deficiencia.','Avis','OK','Warning') | Out-Null
        exit 0
    }
    return $result
}

# ----------------------------------------------------------------------------
# Step 4 - Field placeholders [CAMP: nom (hint)]
# ----------------------------------------------------------------------------
function Get-FieldsFromSelection($selectedSections) {
    $fields = [ordered]@{}
    $regex = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            $allText = ($it.BodyLines -join ' ')
            foreach ($ch in $it.Children) {
                $allText += ' ' + ($ch.BodyLines -join ' ')
            }
            foreach ($m in $regex.Matches($allText)) {
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

function Prompt-Fields($fields) {
    if ($fields.Count -eq 0) { return $fields }

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
    $regex = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'
    return $regex.Replace($text, {
        param($m)
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) { $name = $raw.Substring(0, $parenIdx).Trim() }
        if ($fields.Contains($name)) { return $fields[$name].Value }
        return ''
    })
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

function Select-Conclusions($conclusions) {
    if ($conclusions.Count -eq 0) { return @() }

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
    # Els placeholders disponibles (segons 0 CAPCALERA.docx) son:
    #   <<ID_GIA>> <<EXP_NUM>> <<ADRECA>> <<ACTIVITAT>> <<TITULAR>>
    #   <<NUM_ANOTACIO>> <<DATA_ANOTACIO>>
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
        $find.Wrap = 1  # wdFindContinue
        $find.MatchCase = $false
        $find.Execute([ref]$k, $false, $false, $false, $false, $false, $true, 1, $false, [string]$map[$k], 2) | Out-Null
    }
}

function Build-Document($word, $header, $selectedSections, $fields, $conclusions, $alwaysConclusions, $catalegName) {
    # Nom del fitxer: YYYY-MM-DD_<TipusCataleg>_GIA <id_gia>.docx
    #   <TipusCataleg> = BaseName del cataleg amb la primera lletra en
    #                    majuscula i la resta en minuscules (REQ1 -> Req1).
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $cat   = $catalegName
    if ($cat) {
        $cat = $cat.Substring(0,1).ToUpper() + $cat.Substring(1).ToLower()
    } else {
        $cat = 'Informe'
    }
    $gia = $header['ID_GIA']
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    $fileName = "{0}_{1}_GIA {2}.docx" -f $today, $cat, $gia

    # Triem el directori: el principal si es accessible, si no la carpeta local.
    $targetDir = $OutputDir
    try {
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        $targetDir = Join-Path $ScriptRoot 'Informes generats'
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
    }
    $outPath = Join-Path $targetDir $fileName

    # Treballem amb una copia LOCAL (a %TEMP%) per evitar que Word obri el
    # fitxer en "Vista protegida" quan el desti es una unitat de xarxa.
    # En acabar, movem el fitxer final al directori de sortida.
    $tempPath = Join-Path $env:TEMP $fileName
    Copy-Item -LiteralPath $HeaderPath -Destination $tempPath -Force
    $doc = $word.Documents.Open($tempPath, $false, $false)

    # Si malgrat tot s'ha obert en mode protegit, sortim-ne.
    try {
        if ($doc.ProtectedViewWindow -ne $null) {
            $doc = $doc.ProtectedViewWindow.Edit()
        }
    } catch { }

    Apply-HeaderReplacements -doc $doc -header $header

    # Activem el document i fem servir Selection per inserir text al final.
    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)  # wdStory = 6

    # Logica d'escriptura. Les funcions Format-X venen de Format.ps1.
    $globalItemCounter = 0
    $lastSectionName = $null

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

    foreach ($sec in $selectedSections) {
        # Si el titol te " - ", el partim en seccio + subseccio. Si la
        # seccio coincideix amb l'anterior, no la repetim.
        $parts = $sec.Title -split ' - ', 2
        if ($parts.Count -eq 2) {
            $secName = $parts[0].Trim()
            $subName = $parts[1].Trim()
            if ($secName -ne $lastSectionName) {
                Format-Section $sel $secName
                $lastSectionName = $secName
            }
            Format-Subsection $sel $subName
        } else {
            Format-Section $sel $sec.Title
            $lastSectionName = $sec.Title
        }

        foreach ($it in $sec.Items) {
            $itemLines = @($it.BodyLines | ForEach-Object { Apply-Fields -text $_ -fields $fields })
            $hasChildren = ($it.Children.Count -gt 0)
            $itemWritten = $false

            # Item pare numerat (sempre que tinguem cos o sigui un atomic
            # seleccionat). Si nomes hi ha fills i el pare no esta seleccionat
            # i no te cos, no escrivim res pel pare.
            if ($it.Selected -or $hasChildren) {
                if ($itemLines.Count -gt 0) {
                    $globalItemCounter++
                    Format-Item $sel "$globalItemCounter." $itemLines[0]
                    & $emitExtras $itemLines $false
                    $itemWritten = $true
                }
            }

            if ($hasChildren) {
                $subCounter = 0
                foreach ($ch in $it.Children) {
                    $childLines = @($ch.BodyLines | ForEach-Object { Apply-Fields -text $_ -fields $fields })
                    if ($childLines.Count -eq 0) { continue }
                    $subCounter++
                    if (-not $itemWritten) {
                        # Cas estrany: fills sense pare escrit. Garantitzem
                        # que el comptador global avanci almenys un cop per
                        # tenir una numeracio coherent.
                        $globalItemCounter++
                        $itemWritten = $true
                    }
                    Format-Item $sel "$globalItemCounter.$subCounter." $childLines[0] -IsChild
                    & $emitExtras $childLines $true
                }
            }

            # Espai despres de l'item complet (incloent els seus fills).
            if ($itemWritten -and $Script:ReportFormatConfig.SpacerAfterItem) {
                Format-Spacer $sel
            }
        }
    }

    $hasConcl = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    if ($hasConcl) {
        if ($Script:ReportFormatConfig.SpacerBeforeConclusionsBlock) {
            Format-Spacer $sel
        }
        $allConcl = @($conclusions) + @($alwaysConclusions)
        $totalC = $allConcl.Count
        for ($i = 0; $i -lt $totalC; $i++) {
            Format-Conclusion $sel $allConcl[$i]
            # Espai entre conclusions (i abans de l'ultima): tret de l'ultima.
            if ($Script:ReportFormatConfig.SpacerBetweenConclusions -and $i -lt ($totalC - 1)) {
                Format-Spacer $sel
            }
        }
    }

    $doc.Save()
    $doc.Close($false)

    # Movem el fitxer al desti final (xarxa o local segons disponibilitat).
    try {
        Move-Item -LiteralPath $tempPath -Destination $outPath -Force
    } catch {
        # Si no podem moure (xarxa caiguda), deixem el fitxer al TEMP i
        # informem-ne tornant aquesta ruta.
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

    $cataleg = Select-Cataleg
    $header  = Get-HeaderData

    $word = New-WordApp
    try {
        $sections     = Parse-Cataleg -word $word -path $cataleg.FullName
        $selected     = Select-Items $sections
        $fields       = Get-FieldsFromSelection $selected
        $fields       = Prompt-Fields $fields
        $conclusionsAll = Read-Conclusions -word $word -path $ConclusionsPath
        $conclusions  = Select-Conclusions $conclusionsAll.Selectable
        $outPath      = Build-Document -word $word -header $header `
                                       -selectedSections $selected `
                                       -fields $fields `
                                       -conclusions $conclusions `
                                       -alwaysConclusions $conclusionsAll.Always `
                                       -catalegName $cataleg.BaseName

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
