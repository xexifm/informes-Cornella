#requires -Version 5.1
<#
.SYNOPSIS
  La FINESTRA de l'eina "Excel Errors" (menu, fila GIA) i el seu punt d'entrada.

.DESCRIPTION
  Va a part d'ExcelErrors.ps1 (les comprovacions, la lectura i l'exportacio)
  i d'ExcelErrorsLlistats.ps1 (els tres llistats per correu i els exclosos):
  la finestra fa servir els dos, i l'automatic nomes el segon. Si la finestra
  fos a ExcelErrors.ps1, aquell fitxer i el dels llistats dependrien l'un de
  l'altre (el guard de cicles ho enxampa).
#>

# ----------------------------------------------------------------------------
# LA FINESTRA
# ----------------------------------------------------------------------------
function Show-ExcelErrorsFinestra($errors, [string]$peu, [string]$fileName) {
    $titol = 'Excel Errors'
    $comptes = _ExComptes $errors
    $claus = @($Script:ExTipus.Keys)
    # Dins d'una closure, $Script: NO es la de l'script (CLAUDE.md): el valor,
    # capturat aqui.
    $llistats = $Script:ExLlistats
    # Excl: les activitats TRETES dels llistats (ExcelErrorsLlistats.ps1): no
    # surten ni a la graella, ni a l'Excel, ni al correu.
    $st = @{ Loading = $false; Canvi = $null; Excl = @(Read-ExExclosos) }

    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(1180, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 560)
    $form.StartPosition = 'CenterScreen'

    # ---- A dalt: els tipus d'error, per marcar ---------------------------
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 254
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true; $lbl.Location = New-Object System.Drawing.Point(10, 8)
    $lbl.Text = "Tipus d'error (marca els que vols veure i exportar):"
    $top.Controls.Add($lbl)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(10, 28); $clb.Size = New-Object System.Drawing.Size(1140, 170)
    $clb.Anchor = 'Top, Left, Right'
    $clb.CheckOnClick = $true; $clb.IntegralHeight = $false
    $clb.MultiColumn = $true; $clb.ColumnWidth = 560
    $triats = @(Read-ExTriats)
    foreach ($k in $claus) {
        $i = $clb.Items.Add(([string]$Script:ExTipus[$k] + '  (' + [int]$comptes[$k] + ')'))
        $clb.SetItemChecked($i, ($triats -contains $k))
    }
    $top.Controls.Add($clb)
    $lnkTotes = New-Object System.Windows.Forms.LinkLabel
    $lnkTotes.Text = 'Marca-les totes'; $lnkTotes.AutoSize = $true
    $lnkTotes.Location = New-Object System.Drawing.Point(10, 206)
    $top.Controls.Add($lnkTotes)
    $lnkCap = New-Object System.Windows.Forms.LinkLabel
    $lnkCap.Text = 'Desmarca-les totes'; $lnkCap.AutoSize = $true
    $lnkCap.Location = New-Object System.Drawing.Point(120, 206)
    $top.Controls.Add($lnkCap)
    $lblCompte = New-Object System.Windows.Forms.Label
    $lblCompte.AutoSize = $true
    $lblCompte.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblCompte.Location = New-Object System.Drawing.Point(640, 206)
    $top.Controls.Add($lblCompte)
    # On aniria cada error si s'envies ara (els tres llistats del correu).
    $lblLl = New-Object System.Windows.Forms.Label
    $lblLl.AutoSize = $true
    $lblLl.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblLl.Location = New-Object System.Drawing.Point(10, 230)
    $top.Controls.Add($lblLl)

    # ---- La graella (dins d'un panell: _StyleListGrid fa Dock=Fill) --------
    $mig = New-Object System.Windows.Forms.Panel
    $mig.Dock = 'Fill'
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.ReadOnly = $true
    # Mes d'una fila: "Treure dels llistats" en pot treure diverses de cop.
    $grid.MultiSelect = $true
    foreach ($c in @(@{ H = 'ID Activitat'; W = 80 }, @{ H = 'Classif. annex'; W = 90 }, @{ H = 'Classif. Apartat'; W = 100 }, @{ H = 'Llistat'; W = 120 }, @{ H = ('Ra' + [char]0x00F3 + ' social'); W = 190 }, @{ H = "Tipus d'error"; W = 290 },
                     @{ H = 'Camp'; W = 150 }, @{ H = "Valor a l'Excel"; W = 170 }, @{ H = 'Detall'; W = 380 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c.H; $col.Width = $c.W; $col.SortMode = 'NotSortable'
        [void]$grid.Columns.Add($col)
    }
    $mig.Controls.Add($grid)
    $txtCerca = _AddSearchBox $top 300 203 300 'Cerca:' $null

    # Els tipus marcats. Dins de l'ItemCheck el canvi encara no es efectiu
    # (salta ABANS): $st.Canvi porta el que s'esta canviant.
    $triatsAra = {
        $out = @()
        for ($i = 0; $i -lt $clb.Items.Count; $i++) {
            $m = $clb.GetItemChecked($i)
            if ($null -ne $st.Canvi -and [int]$st.Canvi.I -eq $i) { $m = [bool]$st.Canvi.V }
            if ($m) { $out += $claus[$i] }
        }
        return ,$out
    }.GetNewClosure()

    $fill = {
        if ($st.Loading) { return }
        $grid.Rows.Clear()
        $sel = & $triatsAra
        $n = ([string]$txtCerca.Text).Trim().ToLower()
        $nf = 0
        $ex = @{}; foreach ($x in @($st.Excl)) { $ex[[string]$x.Id] = $true }
        $visibles = @(@($errors) | Where-Object { -not $ex.ContainsKey([string]$_.Id) })
        foreach ($e in @(_ExTriats $visibles $sel)) {
            $ll = _ExLlistatDe ([string]$e.Annex) ([string]$e.Apartat) ([string]$e.Estat)
            $llNom = if ($ll -ne '') { [string]$llistats[$ll] } else { '(cap)' }
            $fila = @([string]$e.Id, [string]$e.Annex, [string]$e.Apartat, $llNom, [string]$e.Titular, [string]$e.TipusNom, [string]$e.Camp, [string]$e.Valor, [string]$e.Detall)
            if ($n -ne '' -and -not (_TextMatches ($fila -join ' ') $n)) { continue }
            $ix = $grid.Rows.Add($fila)
            $grid.Rows[$ix].Tag = $e
            $nf++
        }
        $nAct = @($visibles | Where-Object { $sel -contains [string]$_.Tipus } | ForEach-Object { [string]$_.Id } | Sort-Object -Unique).Count
        $lblCompte.Text = ([string]$nf + ' errors ' + [char]0x00B7 + ' ' + $nAct + ' activitats' + $(if (@($st.Excl).Count -gt 0) { ' ' + [char]0x00B7 + ' ' + @($st.Excl).Count + ' tretes dels llistats' } else { '' }))
        $per = _ExPerLlistat $errors $sel $st.Excl
        $lblLl.Text = ('Per correu: ' + ((@($llistats.Keys) | ForEach-Object { [string]$llistats[$_] + ' ' + (_ExResumLlistat @($per[$_])) }) -join '  ' + [char]0x00B7 + '  ') +
                       '  ' + [char]0x00B7 + '  sense llistat: ' + @($per.SenseLlistat).Count + ' activitats')
    }.GetNewClosure()

    $clb.add_ItemCheck({
        param($s, $e)
        $st.Canvi = @{ I = $e.Index; V = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) }
        try { & $fill } finally { $st.Canvi = $null }
    }.GetNewClosure())
    $txtCerca.add_TextChanged({ & $fill }.GetNewClosure())
    $lnkTotes.add_LinkClicked({ $st.Loading = $true; for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }; $st.Loading = $false; & $fill }.GetNewClosure())
    $lnkCap.add_LinkClicked({ $st.Loading = $true; for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }; $st.Loading = $false; & $fill }.GetNewClosure())

    # ---- El peu: el que no s'ha pogut mirar, i els botons ------------------
    $baix = New-Object System.Windows.Forms.Panel
    $baix.Dock = 'Bottom'; $baix.Height = 112
    $tbPeu = New-Object System.Windows.Forms.TextBox
    $tbPeu.Multiline = $true; $tbPeu.ReadOnly = $true; $tbPeu.ScrollBars = 'Vertical'
    $tbPeu.Location = New-Object System.Drawing.Point(10, 4); $tbPeu.Size = New-Object System.Drawing.Size(1140, 56)
    $tbPeu.Anchor = 'Top, Left, Right'
    $tbPeu.BackColor = [System.Drawing.Color]::White
    $tbPeu.Text = $peu
    $baix.Controls.Add($tbPeu)
    [void](_AddPeuBotons $form @(
        @{ Nom = 'Tancar'; Text = 'Tancar'; Clic = { $form.Close() }.GetNewClosure() },
        # Les tretes dels llistats: es poden tornar a posar.
        @{ Nom = 'Exclosos'; Text = 'Tretes dels llistats...'; Clic = {
            if (@($st.Excl).Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No n'hi ha cap de treta.", $titol, 'OK', 'Information') | Out-Null; return }
            $tornar = @(Select-ExExclososATornar $st.Excl)
            if ($tornar.Count -eq 0) { return }
            $nous = @(_ExExclososAmb $st.Excl $tornar $true)
            try { Save-ExExclosos $nous } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null; return }
            $st.Excl = $nous
            & $fill
        }.GetNewClosure() }) @(
        # L'usuari: "vull poder eliminar activitats d'aquests llistats, no
        # m'importa si tenen errors".
        @{ Nom = 'Treure'; Text = 'Treure dels llistats'; Clic = {
            $sel = @(@($grid.SelectedRows) | ForEach-Object { $_.Tag } | Where-Object { $null -ne $_ })
            if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Tria a la llista les activitats que vols treure.", $titol, 'OK', 'Information') | Out-Null; return }
            $items = @($sel | ForEach-Object { @{ Id = [string]$_.Id; Text = ([string]$_.Titular + $(if ([string]$_.Annex) { ' (' + [string]$_.Annex + ')' } else { '' })) } })
            $ids = @($items | ForEach-Object { $_.Id } | Sort-Object -Unique)
            $rr = [System.Windows.Forms.MessageBox]::Show(("Es trauran dels llistats (i d'aquesta finestra) les activitats: " + ($ids -join ', ') + ".`n`nNo tornaran a sortir, tinguin els errors que tinguin, fins que les tornis a posar (boto 'Tretes dels llistats...').`n`nVols continuar?"), $titol, 'YesNo', 'Question')
            if ($rr -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            $nous = @(_ExExclososAmb $st.Excl $items)
            try { Save-ExExclosos $nous } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null; return }
            $st.Excl = $nous
            & $fill
        }.GetNewClosure() },
        # EL MATEIX QUE L'AUTOMATIC: els tres Excels i el correu.
        @{ Nom = 'Correu'; Text = 'Enviar per correu'; Clic = {
            $sel = & $triatsAra
            if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Marca almenys un tipus d'error.", $titol, 'OK', 'Information') | Out-Null; return }
            Save-ExTriats $sel
            $ccoT = Select-CorreuCcoTria 'excelerrors'
            if (-not $ccoT.Ok) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try {
                $g = New-ExcelErrorsLlistats $errors $sel $st.Excl
                if (-not $g.Ok) { [System.Windows.Forms.MessageBox]::Show("No s'han pogut fer els llistats:`n" + $g.Error, $titol, 'OK', 'Error') | Out-Null; return }
                $e = Send-ExcelErrorsCorreu $g $false @($ccoT.Adreces)
                $txt = ((@($g.Resum.Keys) | ForEach-Object { [string]$_ + ': ' + [string]$g.Resum[$_] }) -join "`n") + "`n`n" + $e.Text
                if (@($g.SenseLlistat).Count -gt 0) { $txt += "`n`nAmb errors i sense llistat (no s'envien): " + ((@($g.SenseLlistat) | ForEach-Object { [string]$_.Id }) -join ', ') }
                [System.Windows.Forms.MessageBox]::Show($txt, $titol, 'OK', $(if ($e.Ok) { 'Information' } else { 'Warning' })) | Out-Null
            } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
        }.GetNewClosure() },
        @{ Nom = 'Export'; Text = 'Exportar a Excel'; Estil = 'primari'; Clic = {
            $sel = & $triatsAra
            if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Marca almenys un tipus d'error.", $titol, 'OK', 'Information') | Out-Null; return }
            Save-ExTriats $sel
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
            $dlg.FileName = 'Excel errors ' + (Get-Date).ToString('yyyy-MM-dd') + '.xlsx'
            try { $dlg.InitialDirectory = _ResolveOutputDir } catch { }
            if ($dlg.ShowDialog() -ne 'OK') { return }
            try {
                $ex = @{}; foreach ($x in @($st.Excl)) { $ex[[string]$x.Id] = $true }
                [void](Export-ExcelErrors $dlg.FileName @(@($errors) | Where-Object { -not $ex.ContainsKey([string]$_.Id) }) $sel)
                [void](Open-PerUsuari $dlg.FileName)
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut exportar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() }) 64 $baix)

    # L'ordre compta (Dock): el que s'afegeix l'ultim es col·loca el primer.
    $form.Controls.Add($mig)
    $form.Controls.Add($top)
    $form.Controls.Add($baix)
    [void](_AddBrandHeader $form $titol ("Errors de la base de dades d'activitats " + [char]0x00B7 + ' ' + $fileName + ' ' + [char]0x00B7 + " l'Excel no es toca") 56)
    $form.add_FormClosing({ Save-ExTriats (& $triatsAra) }.GetNewClosure())
    _AjustaFinestraAPantalla $form
    & $fill
    [void]$form.ShowDialog()
    $form.Dispose()
}

# Les tretes dels llistats, per triar quines es tornen a posar. Torna les
# triades (cap si es tanca).
function Select-ExExclososATornar($exclosos) {
    $f = _NewForm
    $f.Text = 'Tretes dels llistats'
    $f.ClientSize = New-Object System.Drawing.Size(560, 420)
    $f.MinimumSize = New-Object System.Drawing.Size(420, 320)
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(14, 66); $l.Size = New-Object System.Drawing.Size(532, 20)
    $l.Anchor = 'Top, Left, Right'
    $l.Text = "Marca les que vols tornar a posar als llistats:"
    [void]$f.Controls.Add($l)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(14, 90); $clb.Size = New-Object System.Drawing.Size(532, 270)
    $clb.Anchor = 'Top, Bottom, Left, Right'
    $clb.CheckOnClick = $true; $clb.IntegralHeight = $false
    $llista = @($exclosos)
    foreach ($x in $llista) { [void]$clb.Items.Add(('GIA ' + [string]$x.Id + $(if ([string]$x.Text) { '  -  ' + [string]$x.Text } else { '' }))) }
    [void]$f.Controls.Add($clb)
    $res = @{ Triades = @() }
    $peu = _AddPeuBotons $f @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) @(@{ Nom = 'Tornar'; Text = 'Tornar als llistats'; Estil = 'primari' }) 370 -Ancorat
    $peu.Tornar.add_Click({
        $res.Triades = @(@($clb.CheckedIndices | ForEach-Object { [int]$_ }) | ForEach-Object { $llista[$_] })
        $f.Close()
    }.GetNewClosure())
    [void](_AddBrandHeader $f 'Tretes dels llistats' "No surten a l'Excel Errors ni als llistats del correu" 56)
    [void]$f.ShowDialog()
    $f.Dispose()
    return @($res.Triades)
}

# ----------------------------------------------------------------------------
# Punt d'entrada (menu, fila GIA)
# ----------------------------------------------------------------------------
function Invoke-ExcelErrors {
    $titol = 'Excel Errors'
    $sp = _NewForm
    $sp.Text = $titol
    $sp.Size = New-Object System.Drawing.Size(420, 130)
    $sp.FormBorderStyle = 'FixedDialog'
    $sp.MaximizeBox = $false; $sp.MinimizeBox = $false
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(20, 18); $l.Size = New-Object System.Drawing.Size(370, 40)
    $l.Text = "Llegint la base d'activitats (Excel)..."
    $sp.Controls.Add($l)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 62); $bar.Size = New-Object System.Drawing.Size(370, 20)
    $bar.Style = 'Marquee'
    $sp.Controls.Add($bar)
    $sp.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $res = $null; $ct = $null
    try { $res = _ReadExcelErrors; $ct = _ExContactes } finally { try { $sp.Close() } catch { } }
    if ($null -eq $res -or -not [bool]$res.Ok) {
        $msg = if ($res) { [string]$res.Error } else { 'Error desconegut.' }
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'activitats:`n$msg", $titol, 'OK', 'Error') | Out-Null
        return
    }
    $errors = @(Get-ExcelErrors @($res.Files) $res.Col $ct)
    Show-ExcelErrorsFinestra $errors (_ExPeu $res.Col ($null -ne $ct) @($res.Files).Count) ([string]$res.File)
}
