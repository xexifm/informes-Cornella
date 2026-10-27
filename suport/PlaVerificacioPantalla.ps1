#requires -Version 5.1
<#
.SYNOPSIS
  La FINESTRA de l'eina "Pla verificacio" (menu, fila BASE D'INFORMES) i el
  seu punt d'entrada. Els calculs son a PlaVerificacio.ps1.
#>

# Una de les dues taules de la Memoria: un grup amb una graella de dues
# columnes (Indicador | $col), l'indicador en cursiva i el valor centrat.
function _PvTaulaGraella($pare, [string]$titol, [string]$col, [int]$x, [int]$y, [int]$w, [int]$nFiles) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $titol
    $g.Location = New-Object System.Drawing.Point($x, $y)
    $g.Size = New-Object System.Drawing.Size($w, (24 + 26 + 26 * $nFiles + 14))
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.ReadOnly = $true
    $grid.ScrollBars = 'None'
    $grid.ColumnHeadersHeight = 26
    $grid.RowTemplate.Height = 26
    $c1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c1.HeaderText = 'Indicador'; $c1.AutoSizeMode = 'Fill'; $c1.SortMode = 'NotSortable'
    $c1.DefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Italic)
    $c2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c2.HeaderText = $col; $c2.Width = 96; $c2.SortMode = 'NotSortable'
    $c2.DefaultCellStyle.Alignment = 'MiddleCenter'
    $c2.HeaderCell.Style.Alignment = 'MiddleCenter'
    [void]$grid.Columns.Add($c1); [void]$grid.Columns.Add($c2)
    [void]$g.Controls.Add($grid)
    [void]$pare.Controls.Add($g)
    return $grid
}

# $dades: Read-PvDades (els programes signats i les fotos dels anys tancats).
function Show-PlaVerificacioFinestra($files, $informes, [string]$fileName, $dades) {
    $titol = ('Pla verificaci' + [char]0x00F3)
    $avui = Get-Date
    $anys = _PvAnys $avui $files
    # Dins d'una closure, $Script: NO es la de l'script (CLAUDE.md): els
    # valors, capturats aqui.
    $tExec = $Script:PvTitolExecucio
    $tAct = $Script:PvTitolActivitats
    $grups = $Script:PvGrups
    $capLl = $Script:PvLlistaCapcalera
    $amplesLl = $Script:PvLlistaAmples
    $st = @{ Ind = $null; Locals = (Read-PvPrevistes); Carregant = $false; Any = 0 }

    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(1180, 800)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = 'CenterScreen'

    # ---- A dalt: l'any, les previstes i les dues taules -------------------
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 300
    $lAny = New-Object System.Windows.Forms.Label
    $lAny.Text = 'Any:'; $lAny.Location = New-Object System.Drawing.Point(14, 12); $lAny.Size = New-Object System.Drawing.Size(36, 20)
    $cbAny = New-Object System.Windows.Forms.ComboBox
    $cbAny.DropDownStyle = 'DropDownList'
    $cbAny.Location = New-Object System.Drawing.Point(52, 9); $cbAny.Size = New-Object System.Drawing.Size(80, 24)
    foreach ($a in $anys) { [void]$cbAny.Items.Add([string]$a) }
    $lPrev = New-Object System.Windows.Forms.Label
    $lPrev.Text = 'Verificacions previstes al Programa anual:'
    $lPrev.Location = New-Object System.Drawing.Point(156, 12); $lPrev.Size = New-Object System.Drawing.Size(250, 20)
    $nuPrev = New-Object System.Windows.Forms.NumericUpDown
    $nuPrev.Minimum = 0; $nuPrev.Maximum = 999
    $nuPrev.Location = New-Object System.Drawing.Point(408, 9); $nuPrev.Size = New-Object System.Drawing.Size(64, 24)
    $lPrevNota = New-Object System.Windows.Forms.Label
    $lPrevNota.Text = ("Les del programa signat. D'un any nou, posa-les aqu" + [char]0x00ED + ": es desen en aquest PC.")
    $lPrevNota.Location = New-Object System.Drawing.Point(484, 12); $lPrevNota.Size = New-Object System.Drawing.Size(660, 20)
    $lPrevNota.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    foreach ($c in @($lAny, $cbAny, $lPrev, $nuPrev, $lPrevNota)) { [void]$top.Controls.Add($c) }
    $gExec = _PvTaulaGraella $top $tExec 'Seguiment' 14 44 640 5
    $gAct = _PvTaulaGraella $top $tAct 'Resultat' 668 44 482 2
    # El SEGUIMENT DEL PROGRAMA (l'usuari: "un boto per poder-ho veure, pero
    # no com a cosa primordial"): sota la taula de les verificacions.
    $btnSeg = New-Object System.Windows.Forms.Button
    $btnSeg.Text = 'Activitats del programa...'
    $btnSeg.Location = New-Object System.Drawing.Point(14, 246); $btnSeg.Size = New-Object System.Drawing.Size(210, 28)
    _StyleSecondaryButton $btnSeg
    # Un any TANCAT surt amb les dades de la seva Memoria (la foto); es pot
    # comparar amb el que diria ara l'Excel.
    $chkExcel = New-Object System.Windows.Forms.CheckBox
    $chkExcel.Text = ("Calcular-ho amb l'Excel d'ara (en lloc de les dades de la Mem" + [char]0x00F2 + "ria)")
    $chkExcel.Location = New-Object System.Drawing.Point(238, 249); $chkExcel.Size = New-Object System.Drawing.Size(416, 22)
    foreach ($c in @($btnSeg, $chkExcel)) { [void]$top.Controls.Add($c) }
    # Com es compta cada indicador (l'usuari ho ha de poder comprovar).
    $lCrit = New-Object System.Windows.Forms.Label
    $lCrit.Location = New-Object System.Drawing.Point(668, 168); $lCrit.Size = New-Object System.Drawing.Size(482, 126)
    $lCrit.Anchor = 'Top, Left, Right'
    $lCrit.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lCrit.Text = ("Com es compta (Excel " + $fileName + "):`n" +
        [char]0x00B7 + " Programa: 'Data control inicial/verificaci" + [char]0x00F3 + "' de l'any i annex que no " + [char]0x00E9 + "s II ni III.`n" +
        [char]0x00B7 + " Den" + [char]0x00FA + "ncies: data d''Inspecci" + [char]0x00F3 + "' de l'any.   " + [char]0x00B7 + " EC-PCA: control de l'any i annex III.`n" +
        [char]0x00B7 + " Greus: del programa, amb una visita del Pla de verificaci" + [char]0x00F3 + " de l'any amb Precinte / Cessament (no les de Queixa sola).`n" +
        "L'Excel nom" + [char]0x00E9 + "s guarda la darrera data de control: un any tancat surt amb les dades de la seva Mem" + [char]0x00F2 + "ria.")
    [void]$top.Controls.Add($lCrit)

    # ---- Al mig: les activitats que compten -------------------------------
    $mig = New-Object System.Windows.Forms.Panel
    $mig.Dock = 'Fill'
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.ReadOnly = $true
    for ($i = 0; $i -lt $capLl.Count; $i++) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = [string]$capLl[$i]; $col.Width = [int]$amplesLl[$i] * 6 + 20; $col.SortMode = 'NotSortable'
        [void]$grid.Columns.Add($col)
    }
    $mig.Controls.Add($grid)
    $lLl = New-Object System.Windows.Forms.Label
    $lLl.Dock = 'Top'; $lLl.Height = 22
    $lLl.Padding = New-Object System.Windows.Forms.Padding(14, 4, 0, 0)
    $mig.Controls.Add($lLl)

    # Torna a calcular i pinta tot per a l'any triat.
    $fill = {
        if ($cbAny.SelectedIndex -lt 0) { return }
        $any = [int]$anys[$cbAny.SelectedIndex]
        $prev = _PvPrevistes $any $st.Locals $dades
        $foto = if ($dades.Fotos.Contains([string]$any)) { $dades.Fotos[[string]$any] } else { $null }
        $st.Carregant = $true
        try {
            $nuPrev.Value = [decimal]$prev
            # La casella nomes te sentit si l'any te foto; en canviar d'any, desmarcada.
            if ($st.Any -ne $any) { $chkExcel.Checked = $false }
            $chkExcel.Visible = ($null -ne $foto)
        } finally { $st.Carregant = $false }
        $st.Any = $any
        $btnSeg.Enabled = $dades.Programes.Contains([string]$any)
        $ind = if ($null -ne $foto -and -not $chkExcel.Checked) { Get-PvFotoIndicadors $foto $any $prev } else { Get-PlaVerificacioIndicadors $files $any $prev $informes }
        $st.Ind = $ind
        $gExec.Rows.Clear(); foreach ($f in @(_PvFilesExecucio $ind)) { [void]$gExec.Rows.Add(@([string]$f[0], [string]$f[1])) }
        $gAct.Rows.Clear(); foreach ($f in @(_PvFilesActivitats $ind)) { [void]$gAct.Rows.Add(@([string]$f[0], [string]$f[1])) }
        $gExec.ClearSelection(); $gAct.ClearSelection()
        $grid.Rows.Clear()
        foreach ($e in @($ind.Llista)) {
            $ix = $grid.Rows.Add([object[]]@(_PvFilaLlista $e))
            if ($null -ne $e.Precinte) { $grid.Rows[$ix].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(160, 40, 40) }
        }
        $perGrup = (@(@($grups.Keys) | ForEach-Object { $k = $_; [string]$grups[$k] + ' ' + @(@($ind.Llista) | Where-Object { [string]$_.Grup -eq $k }).Count })) -join ('  ' + [char]0x00B7 + '  ')
        $lLl.Text = ("Activitats que compten el " + $any + ": " + $perGrup +
            $(if ([bool]$ind.Foto) { "   (dades de la " + [string]$ind.Document + ")" } elseif ($ind.SenseBase) { "   (sense base d'informes: les greus no es poden saber; fes Actualitzar base)" } else { '' }))
    }.GetNewClosure()

    $cbAny.add_SelectedIndexChanged({ & $fill }.GetNewClosure())
    $chkExcel.add_CheckedChanged({ if (-not $st.Carregant) { & $fill } }.GetNewClosure())
    $btnSeg.add_Click({
        $k = [string]$st.Any
        if (-not $dades.Programes.Contains($k)) { return }
        $foto = if ($dades.Fotos.Contains($k) -and -not $chkExcel.Checked) { $dades.Fotos[$k] } else { $null }
        Show-PvSeguimentPrograma $dades.Programes[$k] $files ([int]$st.Any) $foto
    }.GetNewClosure())
    # Les previstes d'un any: es desen en aquest PC en canviar-les.
    $nuPrev.add_ValueChanged({
        if ($st.Carregant -or $cbAny.SelectedIndex -lt 0) { return }
        $any = [string]$anys[$cbAny.SelectedIndex]
        $st.Locals[$any] = [int]$nuPrev.Value
        try { Save-PvPrevistes $st.Locals } catch { }
        & $fill
    }.GetNewClosure())

    # ---- El peu ----------------------------------------------------------
    $baix = New-Object System.Windows.Forms.Panel
    $baix.Dock = 'Bottom'; $baix.Height = 60
    [void](_AddPeuBotons $form @(
        @{ Nom = 'Tancar'; Text = 'Tancar'; Clic = { $form.Close() }.GetNewClosure() }) @(
        # Per enganxar-les a la Memoria: amb el format de les taules del Word.
        @{ Nom = 'Copiar'; Text = 'Copiar les taules'; Clic = {
            if ($null -eq $st.Ind) { return }
            try {
                $do = New-Object System.Windows.Forms.DataObject
                $do.SetData([System.Windows.Forms.DataFormats]::Html, (_PvCfHtml (_PvTaulesHtml $st.Ind)))
                $do.SetData([System.Windows.Forms.DataFormats]::UnicodeText, (_PvTaulesText $st.Ind))
                [System.Windows.Forms.Clipboard]::SetDataObject($do, $true)
                [System.Windows.Forms.MessageBox]::Show("Copiades. Enganxa-les al Word (Ctrl+V): surten com les taules de la Mem" + [char]0x00F2 + "ria.", $titol, 'OK', 'Information') | Out-Null
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'han pogut copiar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() },
        @{ Nom = 'Export'; Text = 'Exportar la llista a Excel'; Estil = 'primari'; Clic = {
            if ($null -eq $st.Ind) { return }
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
            $dlg.FileName = 'Pla verificacio ' + [string]$st.Ind.Any + ' - activitats (' + (Get-Date).ToString('yyyy-MM-dd') + ').xlsx'
            try { $dlg.InitialDirectory = _ResolveOutputDir } catch { }
            if ($dlg.ShowDialog() -ne 'OK') { return }
            try {
                $bytes = _NormativaXlsxBytes $capLl @(@($st.Ind.Llista) | ForEach-Object { , @(_PvFilaLlista $_) }) $amplesLl ('Verificacions ' + [string]$st.Ind.Any)
                [System.IO.File]::WriteAllBytes($dlg.FileName, $bytes)
                [void](Open-PerUsuari $dlg.FileName)
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut exportar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() }) 12 $baix)

    # L'ordre compta (Dock): el que s'afegeix l'ultim es col·loca el primer.
    $form.Controls.Add($mig)
    $form.Controls.Add($top)
    $form.Controls.Add($baix)
    [void](_AddBrandHeader $form $titol ("Indicadors de la Mem" + [char]0x00F2 + "ria anual " + [char]0x00B7 + " Excel d'activitats i base d'informes") 56)
    _AjustaFinestraAPantalla $form
    # L'any d'avui, o el darrer que hi hagi.
    $ixAvui = [array]::IndexOf($anys, $avui.Year)
    $cbAny.SelectedIndex = $(if ($ixAvui -ge 0) { $ixAvui } else { $anys.Count - 1 })
    [void]$form.ShowDialog()
    $form.Dispose()
}

# ----------------------------------------------------------------------------
# EL SEGUIMENT DEL PROGRAMA (boto secundari de la finestra)
# ----------------------------------------------------------------------------
function Show-PvSeguimentPrograma($prog, $files, [int]$any, $foto) {
    $titol = ('Seguiment del programa ' + $any)
    $seg = Get-PvSeguimentPrograma $prog $files $any $foto
    $cap = $Script:PvSegCapcalera
    $amples = $Script:PvSegAmples
    $f = _NewForm
    $f.Text = $titol
    $f.Size = New-Object System.Drawing.Size(1100, 720)
    $f.MinimumSize = New-Object System.Drawing.Size(760, 480)
    $mig = New-Object System.Windows.Forms.Panel
    $mig.Dock = 'Fill'
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.ReadOnly = $true
    for ($i = 0; $i -lt $cap.Count; $i++) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = [string]$cap[$i]; $col.Width = [int]$amples[$i] * 6 + 20; $col.SortMode = 'NotSortable'
        [void]$grid.Columns.Add($col)
    }
    foreach ($r in @($seg.Files)) {
        $ix = $grid.Rows.Add([object[]]@(_PvFilaSeguiment $r))
        if ([bool]$r.Feta) { $grid.Rows[$ix].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 110, 50) }
    }
    $mig.Controls.Add($grid)
    $l = New-Object System.Windows.Forms.Label
    $l.Dock = 'Top'; $l.Height = 40
    $l.Padding = New-Object System.Windows.Forms.Padding(14, 6, 0, 0)
    $l.Text = ([string]$prog.Document + "`n" +
        "Annex I (criticitat m" + [char]0x00E9 + "s alta): " + $seg.FetesI + ' de ' + $seg.TotalI + ' verificades' +
        $(if ([int]$seg.TotalII -gt 0) { "   " + [char]0x00B7 + "   Annex II (aleat" + [char]0x00F2 + "ries): " + $seg.FetesII + ' de ' + $seg.TotalII } else { '' }) +
        $(if ($null -ne $foto) { "   (segons la Mem" + [char]0x00F2 + "ria)" } else { "   (segons la 'Data control inicial/verificaci" + [char]0x00F3 + "' de l'Excel)" }))
    $mig.Controls.Add($l)
    $baix = New-Object System.Windows.Forms.Panel
    $baix.Dock = 'Bottom'; $baix.Height = 60
    [void](_AddPeuBotons $f @(
        @{ Nom = 'Tancar'; Text = 'Tancar'; Clic = { $f.Close() }.GetNewClosure() }) @(
        @{ Nom = 'Export'; Text = 'Exportar a Excel'; Estil = 'primari'; Clic = {
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
            $dlg.FileName = 'Seguiment programa verificacio ' + $any + ' (' + (Get-Date).ToString('yyyy-MM-dd') + ').xlsx'
            try { $dlg.InitialDirectory = _ResolveOutputDir } catch { }
            if ($dlg.ShowDialog() -ne 'OK') { return }
            try {
                $bytes = _NormativaXlsxBytes $cap @(@($seg.Files) | ForEach-Object { , @(_PvFilaSeguiment $_) }) $amples ('Programa ' + $any)
                [System.IO.File]::WriteAllBytes($dlg.FileName, $bytes)
                [void](Open-PerUsuari $dlg.FileName)
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut exportar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() }) 12 $baix)
    $f.Controls.Add($mig)
    $f.Controls.Add($baix)
    [void](_AddBrandHeader $f $titol ("Les activitats previstes i si ja s'han verificat") 56)
    _AjustaFinestraAPantalla $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

# ----------------------------------------------------------------------------
# Punt d'entrada (menu, fila BASE D'INFORMES)
# ----------------------------------------------------------------------------
function Invoke-PlaVerificacio {
    $titol = ('Pla verificaci' + [char]0x00F3)
    $sp = _NewForm
    $sp.Text = $titol
    $sp.Size = New-Object System.Drawing.Size(420, 130)
    $sp.FormBorderStyle = 'FixedDialog'
    $sp.MaximizeBox = $false; $sp.MinimizeBox = $false
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(20, 18); $l.Size = New-Object System.Drawing.Size(370, 40)
    $l.Text = "Llegint la base d'activitats (Excel) i la base d'informes..."
    $sp.Controls.Add($l)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 62); $bar.Size = New-Object System.Drawing.Size(370, 20)
    $bar.Style = 'Marquee'
    $sp.Controls.Add($bar)
    $sp.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $res = $null; $inf = $null
    try { $res = _ReadPlaVerificacio; $inf = _ReadPvInformes } finally { try { $sp.Close() } catch { } }
    $dades = Read-PvDades
    if ($null -eq $res -or -not [bool]$res.Ok) {
        $msg = if ($res) { [string]$res.Error } else { 'Error desconegut.' }
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'activitats:`n$msg", $titol, 'OK', 'Error') | Out-Null
        return
    }
    $falten = @(_PvColumnesQueFalten $res.Col)
    if ($falten.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(("A la fulla Est" + [char]0x00E8 + "s de l'Excel (" + [string]$res.File + ") no hi ha: " + ($falten -join ', ') + ". Els indicadors que en depenen sortiran a 0."), $titol, 'OK', 'Warning') | Out-Null
    }
    Show-PlaVerificacioFinestra @($res.Files) $inf ([string]$res.File) $dades
}
