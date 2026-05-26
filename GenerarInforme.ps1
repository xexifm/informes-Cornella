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
$EstructuralsDir = Join-Path $ScriptRoot 'ESTRUCTURALS'
$HeaderPath      = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
$ConclusionsPath = Join-Path $EstructuralsDir '0 CONCLUSIONS.docx'

# ----------------------------------------------------------------------------
# Word COM helpers
# ----------------------------------------------------------------------------
function New-WordApp {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0  # wdAlertsNone
    return $w
}

function Close-WordApp($word) {
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
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
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(560, 540)
    $form.StartPosition = 'CenterScreen'

    $fields = [ordered]@{
        'ID_GIA'       = 'ID GIA'
        'EXP_NUM'      = 'Numero d''expedient'
        'ADRECA'       = 'Adreca'
        'ACTIVITAT'    = 'Activitat'
        'PETICIONARI'  = 'Peticionari'
        'DATA'         = 'Data (dd/mm/aaaa)'
        'DECISIO'      = 'Decisio / Resolucio'
        'TECNIC'       = 'Tecnic redactor'
    }

    $controls = @{}
    $y = 20
    foreach ($key in $fields.Keys) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $fields[$key]
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(170, 22)
        $form.Controls.Add($lbl)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(190, ($y - 2))
        $tb.Size = New-Object System.Drawing.Size(340, 22)
        $form.Controls.Add($tb)
        $controls[$key] = $tb
        $y += 38
    }
    # Default the date to today
    $controls['DATA'].Text = (Get-Date).ToString('dd/MM/yyyy')

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(360, ($y + 10))
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(450, ($y + 10))
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    if ($form.ShowDialog() -ne 'OK') { exit 0 }

    $data = @{}
    foreach ($key in $fields.Keys) { $data[$key] = $controls[$key].Text }
    return $data
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
    $form.Size = New-Object System.Drawing.Size(760, 640)
    $form.StartPosition = 'CenterScreen'

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(10, 10)
    $tv.Size = New-Object System.Drawing.Size(720, 540)
    $tv.CheckBoxes = $true
    $tv.HideSelection = $false
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
    $ok.Location = New-Object System.Drawing.Point(560, 560)
    $ok.Size = New-Object System.Drawing.Size(80, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(650, 560)
    $cancel.Size = New-Object System.Drawing.Size(80, 28)
    $cancel.DialogResult = 'Cancel'
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
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $doc = $word.Documents.Open($path, $false, $true)
    try {
        $list = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            $t = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
            if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$list.Add($t) }
        }
        return ,$list.ToArray()
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
    $map = @{
        '<<ID_GIA>>'      = $header['ID_GIA']
        '<<EXP_NUM>>'     = $header['EXP_NUM']
        '<<ADRECA>>'      = $header['ADRECA']
        '<<ACTIVITAT>>'   = $header['ACTIVITAT']
        '<<PETICIONARI>>' = $header['PETICIONARI']
        '<<DATA>>'        = $header['DATA']
        '<<DECISIO>>'     = $header['DECISIO']
        '<<TECNIC>>'      = $header['TECNIC']
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

function Build-Document($word, $header, $selectedSections, $fields, $conclusions) {
    # Determinem el cami de sortida i copiem la CAPCALERA en aquell fitxer,
    # per no tocar mai l'original.
    $safeAct = ($header['ACTIVITAT'] -replace '[\\/:*?"<>|]','_').Trim()
    if (-not $safeAct) { $safeAct = 'Informe' }
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmm')
    $outDir = Join-Path $ScriptRoot 'Informes generats'
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $outPath = Join-Path $outDir ("Informe - {0} - {1}.docx" -f $safeAct, $stamp)
    Copy-Item -LiteralPath $HeaderPath -Destination $outPath -Force

    $doc = $word.Documents.Open($outPath, $false, $false)
    Apply-HeaderReplacements -doc $doc -header $header

    # Inserim el contingut al final.
    $end = $doc.Content
    $end.Collapse(0)  # wdCollapseEnd

    $globalCounter = 0
    foreach ($sec in $selectedSections) {
        # Titol de seccio
        $end.InsertParagraphAfter()
        $end.Collapse(0)
        $end.Text = $sec.Title
        try { $end.set_Style('Heading 1') } catch { $end.Bold = $true }
        $end.Collapse(0)
        $end.InsertParagraphAfter(); $end.Collapse(0)

        foreach ($it in $sec.Items) {
            # Item pare (si l'usuari l'ha marcat com a selected, o si te fills seleccionats)
            $itemTexts = New-Object System.Collections.ArrayList
            if ($it.Selected) {
                $body = ($it.BodyLines -join "`r")
                [void]$itemTexts.Add($body)
            }
            foreach ($ch in $it.Children) {
                $cbody = ($ch.BodyLines -join "`r")
                [void]$itemTexts.Add($cbody)
            }
            foreach ($txt in $itemTexts) {
                $globalCounter++
                $resolved = Apply-Fields -text $txt -fields $fields
                # Si el text ja conte salts de linia interns, el primer es el cos
                # i la resta (URL, p.ex.) van com a paragrafs nous sense numero.
                $parts = $resolved -split "`r"
                $first = $parts[0]
                $end.Text = "{0}. {1}" -f $globalCounter, $first
                try { $end.set_Style('Normal') } catch { }
                $end.Bold = $false
                $end.Collapse(0)
                $end.InsertParagraphAfter(); $end.Collapse(0)
                for ($i = 1; $i -lt $parts.Count; $i++) {
                    if ([string]::IsNullOrWhiteSpace($parts[$i])) { continue }
                    $end.Text = $parts[$i]
                    try { $end.set_Style('Normal') } catch { }
                    $end.Collapse(0)
                    $end.InsertParagraphAfter(); $end.Collapse(0)
                }
            }
        }
    }

    if ($conclusions.Count -gt 0) {
        $end.InsertParagraphAfter(); $end.Collapse(0)
        $end.Text = 'Conclusions'
        try { $end.set_Style('Heading 1') } catch { $end.Bold = $true }
        $end.Collapse(0)
        $end.InsertParagraphAfter(); $end.Collapse(0)
        foreach ($c in $conclusions) {
            $end.Text = $c
            try { $end.set_Style('Normal') } catch { }
            $end.Collapse(0)
            $end.InsertParagraphAfter(); $end.Collapse(0)
        }
    }

    $doc.Save()
    $doc.Close($false)
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
        $conclusions  = Select-Conclusions $conclusionsAll
        $outPath      = Build-Document -word $word -header $header `
                                       -selectedSections $selected `
                                       -fields $fields `
                                       -conclusions $conclusions

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
