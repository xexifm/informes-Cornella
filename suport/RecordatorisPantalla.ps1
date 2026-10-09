#requires -Version 5.1
<#
  RecordatorisPantalla.ps1 - La FINESTRA de l'eina "Recordatoris" (les dues
  pestanyes, Requeriments i Precintes).

  Vivia al final de Recordatoris.ps1, que va passar de les 1.200 linies en
  afegir-hi la preseleccio pel maxim de la tanda i les (i) de cada camp
  (octubre 2026). Es parteix com Llicencia / LlicenciaPantalles: alla hi queden
  les campanyes, el calcul de qui toca, la tanda i la tasca del Windows (el que
  tambe fa servir RecordatorisAuto.ps1, sense finestres); aqui, nomes la
  pantalla.

  NOMES DEFINEIX FUNCIONS.
#>

# ----------------------------------------------------------------------------
# FINESTRA PRINCIPAL
# ----------------------------------------------------------------------------
function Invoke-Recordatoris {
    $db = _RecCarregaDb
    if ($null -eq $db) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base' (secció INFORMES).",
            'Recordatoris', 'OK', 'Information') | Out-Null
        return
    }
    $antig = _RecAntiguitatDb $db (Get-Date)
    $estat = _RecLlegeix
    $quota = _QuotaLlegeix

    $form = _NewForm
    $form.Text = 'Recordatoris'
    $form.ClientSize = New-Object System.Drawing.Size(1080, 660)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $tabs.Padding = New-Object System.Drawing.Point(14, 5)

    # Panell d'informació: quota + frescor de la base. La data de la base és la
    # protecció principal: amb una base vella s'escriuria a qui ja ha complert.
    $info = New-Object System.Windows.Forms.Panel
    $info.Dock = 'Top'; $info.Height = 50
    $lblQ = New-Object System.Windows.Forms.Label
    $lblQ.Location = New-Object System.Drawing.Point(16, 6)
    $lblQ.Size = New-Object System.Drawing.Size(1040, 18)
    # La via de cada campanya (Configuracio -> Correus de cada eina).
    $viesTxt = (@(_RecCampanyes) | ForEach-Object { [string]$_.Nom + ': ' + (_CorreuViaText (Get-CorreuEina ('rec-' + $_.Clau)).Via) }) -join ('  ' + [char]0x00B7 + '  ')
    $lblQ.Text = ($viesTxt + " " + [char]0x00B7 + " EmailJS aquest mes: $($quota.enviats) / $($quota.limit) correus " +
                  "(reserva de $($Script:QuotaLimitCompte - $quota.limit) sobre els $($Script:QuotaLimitCompte) del compte)")
    $info.Controls.Add($lblQ)
    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Location = New-Object System.Drawing.Point(16, 26)
    $lblD.Size = New-Object System.Drawing.Size(1040, 18)
    if ($antig -lt 0) {
        $lblD.Text = "Base d'informes: data desconeguda. Actualitza-la abans d'enviar res."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(176, 0, 32)
    } elseif ($antig -ge $Script:RecAvisAntiguitatDbDies) {
        $lblD.Text = "ATENCIÓ: la base d'informes té $antig dies. Actualitza-la abans d'enviar: podries escriure a qui ja ha complert."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(176, 0, 32)
    } else {
        $lblD.Text = "Base d'informes: actualitzada fa $antig dies · $(@($db.activitats).Count) activitats."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    }
    $info.Controls.Add($lblD)

    foreach ($camp in @(_RecCampanyes)) {
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $camp.Nom
        $tab.BackColor = [System.Drawing.Color]::White
        [void]$tabs.TabPages.Add($tab)
        _RecMuntaTab $tab $camp $estat $db
    }

    $form.Controls.Add($tabs)
    $form.Controls.Add($info)
    [void](_AddBrandHeader $form 'Recordatoris' 'Avisos periòdics als titulars amb tràmits pendents' 56)
    [void]$form.ShowDialog()
}

# Munta el contingut d'UNA pestanya (una campanya). Les dues pestanyes són
# independents: cada una té la seva configuració, el seu text i el seu historial.
function _RecMuntaTab($tab, $camp, $estat, $db) {
    $clau = [string]$camp.Clau
    $cfg  = $estat.campanyes[$clau]
    # Hashtable de funcions: es captura per REFERÈNCIA, que és l'única manera
    # que els handlers vegin les funcions que es defineixen més avall.
    $fn = @{}
    $ui = @{ Rows = @(); SenseGia = 0 }

    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.AllowUserToResizeRows = $false
    $cSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cSel.HeaderText = 'Enviar'; $cSel.Width = 55
    [void]$grid.Columns.Add($cSel)
    foreach ($c in @(
        @{ H = 'GIA';           W = 70  }
        @{ H = 'Titular';       W = 260 }
        @{ H = 'Estat';         W = 130 }
        @{ H = 'Data informe';  W = 95  }
        @{ H = 'Últim avís';    W = 95  }
        @{ H = 'Avisos';        W = 60  }
        @{ H = 'Situació';      W = 230 }
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c.H; $col.Width = $c.W; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }

    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 84

    # Cada valor porta una (i): l'usuari no recordava que volien dir "Cada" i
    # "Espera" (octubre 2026).
    $mkNum = {
        param($etiqueta, $x, $valor, $min, $max, $ajuda)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $etiqueta
        $l.Location = New-Object System.Drawing.Point($x, 13)
        $l.Size = New-Object System.Drawing.Size(96, 18)
        $top.Controls.Add($l)
        $n = New-Object System.Windows.Forms.NumericUpDown
        $n.Location = New-Object System.Drawing.Point(($x + 98), 10)
        $n.Size = New-Object System.Drawing.Size(58, 22)
        $n.Minimum = $min; $n.Maximum = $max; $n.Value = $valor
        $top.Controls.Add($n)
        [void](_AddInfoIcona $top ($x + 160) 12 $etiqueta.TrimEnd(':') $ajuda)
        return $n
    }
    $numPer  = & $mkNum 'Cada (dies):'      16  ([int]$cfg['periodicitatDies'])  1 365 (_RecAjudaCamp 'cada')
    $numEsp  = & $mkNum 'Espera (dies):'    200 ([int]$cfg['esperaInicialDies']) 0 365 (_RecAjudaCamp 'espera')
    $numMax  = & $mkNum 'Màx. per tanda:'   384 ([int]$cfg['maxPerTanda'])       1 150 (_RecAjudaCamp 'max')

    $rbMan = New-Object System.Windows.Forms.RadioButton
    $rbMan.Text = 'Manual'
    $rbMan.Location = New-Object System.Drawing.Point(690, 10)
    $rbMan.Size = New-Object System.Drawing.Size(80, 22)
    $rbAut = New-Object System.Windows.Forms.RadioButton
    $rbAut.Text = 'Automàtic'
    $rbAut.Location = New-Object System.Drawing.Point(772, 10)
    $rbAut.Size = New-Object System.Drawing.Size(90, 22)
    if ([string]$cfg['mode'] -eq 'auto') { $rbAut.Checked = $true } else { $rbMan.Checked = $true }
    $top.Controls.Add($rbMan); $top.Controls.Add($rbAut)
    [void](_AddInfoIcona $top 866 12 ('Manual / Autom' + [char]0x00E0 + 'tic') (_RecAjudaCamp 'mode'))

    $chkNomes = New-Object System.Windows.Forms.CheckBox
    $chkNomes.Text = 'Només els que toquen avui'
    $chkNomes.Location = New-Object System.Drawing.Point(16, 48)
    $chkNomes.Size = New-Object System.Drawing.Size(200, 22)
    $chkNomes.Checked = $true
    $top.Controls.Add($chkNomes)

    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Location = New-Object System.Drawing.Point(610, 51)
    $lblN.Size = New-Object System.Drawing.Size(560, 18)
    $lblN.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    $top.Controls.Add($lblN)

    $bot = New-Object System.Windows.Forms.Panel
    $bot.Dock = 'Bottom'; $bot.Height = 48
    $peu = _AddPeuBotons $tab @(
        @{ Nom = 'Text'; Text = 'Editar text...' },
        @{ Nom = 'Csv'; Text = 'Exportar CSV' },
        @{ Nom = 'Exc'; Text = 'Excloure / incloure' }) @(
        @{ Nom = 'Send'; Text = 'Enviar tanda'; Estil = 'primari' }) 8 $bot
    $btnText = $peu.Text; $btnCsv = $peu.Csv; $btnExc = $peu.Exc; $btnSend = $peu.Send

    # --- Funcions de la pestanya --------------------------------------------
    $txtCerca = _AddSearchBox $top 240 48 300 'Cerca:' { & $fn.Pinta }

    $fn.Desa = {
        $cfg['periodicitatDies']  = [int]$numPer.Value
        $cfg['esperaInicialDies'] = [int]$numEsp.Value
        $cfg['maxPerTanda']       = [int]$numMax.Value
        $cfg['mode']              = if ($rbAut.Checked) { 'auto' } else { 'manual' }
        $estat.campanyes[$clau]   = $cfg
        _RecDesa $estat
    }.GetNewClosure()

    $fn.Calcula = {
        $r = _RecDueActivitats $db $camp $cfg $estat.historial[$clau] (Get-Date)
        $ui.Rows = @($r.Files)
        $ui.SenseGia = [int]$r.SenseGia
    }.GetNewClosure()

    $fn.Pinta = {
        $cerca = ''
        try { $cerca = ([string]$txtCerca.Text).Trim().ToLower() } catch { }
        $nomes = [bool]$chkNomes.Checked
        $grid.Rows.Clear()
        $nToca = 0; $nSel = 0
        foreach ($row in @($ui.Rows)) {
            if ($row.Toca) { $nToca++ }
            if ($row.Sel) { $nSel++ }
            if ($nomes -and -not $row.Toca) { continue }
            if ($cerca -ne '') {
                $hay = ([string]$row.Id + ' ' + [string]$row.Titular + ' ' + [string]$row.Estat).ToLower()
                if (-not $hay.Contains($cerca)) { continue }
            }
            $i = $grid.Rows.Add(@(
                [bool]$row.Sel, [string]$row.Id, [string]$row.Titular, [string]$row.Estat,
                [string]$row.DataInforme, [string]$row.Ultim, [string]$row.Compte, [string]$row.Motiu
            ))
            $grid.Rows[$i].Tag = $row
            if ($row.Excloure) { $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150) }
        }
        $lblN.Text = "$nToca activitats toquen avui, de $(@($ui.Rows).Count) en aquest estat · $nSel marcades (m" + [char]0x00E0 + "x. per tanda: $([int]$cfg['maxPerTanda']))" +
                     $(if ($ui.SenseGia -gt 0) { " · $($ui.SenseGia) sense GIA (no es poden avisar)" } else { '' })
    }.GetNewClosure()

    $fn.Refresca = { & $fn.Calcula; & $fn.Pinta }.GetNewClosure()

    # La casella "Enviar" es desa a l'objecte fila: així sobreviu a filtres i
    # a repintats (mateix patró que Controls periòdics).
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($e.ColumnIndex -ne 0 -or $e.RowIndex -lt 0) { return }
        $r = $s.Rows[$e.RowIndex].Tag
        if ($null -ne $r) { $r.Sel = [bool]$s.Rows[$e.RowIndex].Cells[0].Value }
    })
    $grid.add_CurrentCellDirtyStateChanged({
        param($s, $e)
        if ($s.IsCurrentCellDirty) { $s.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    })

    $chkNomes.add_CheckedChanged({ & $fn.Pinta }.GetNewClosure())
    foreach ($n in @($numPer, $numEsp, $numMax)) { $n.add_ValueChanged({ & $fn.Desa; & $fn.Refresca }.GetNewClosure()) }
    # Un clic en un radio dispara DOS esdeveniments (el que es marca i el germà
    # que es desmarca): només reaccionem al que queda marcat.
    foreach ($rb in @($rbMan, $rbAut)) {
        $rb.add_CheckedChanged({ param($s, $e) if ($s.Checked) { & $fn.Desa } }.GetNewClosure())
    }

    $btnText.add_Click({ if (Invoke-RecordatorisTextos $clau) { $estat.campanyes[$clau] = (_RecLlegeix).campanyes[$clau] } }.GetNewClosure())

    $btnExc.add_Click({
        if ($null -eq $grid.CurrentRow -or $null -eq $grid.CurrentRow.Tag) { return }
        $r = $grid.CurrentRow.Tag
        $nou = -not [bool]$r.Excloure
        $estat.historial[$clau] = _RecHistorialExclou $estat.historial[$clau] ([string]$r.Id) $nou
        _RecDesa $estat
        & $fn.Refresca
    }.GetNewClosure())

    $btnCsv.add_Click({
        $sel = @($ui.Rows)
        if ($sel.Count -eq 0) { return }
        $cache = $null
        try {
            $xls = Find-LatestActivitatsExcel
            if ($null -ne $xls) { $cache = Initialize-ActivitatsCache $xls.File }
        } catch { }
        $out = New-Object System.Collections.ArrayList
        foreach ($r in $sel) {
            $r2 = _RecOmpleDadesFila $r $cache (Get-CorreuEina ('rec-' + $clau))
            [void]$out.Add([pscustomobject]@{
                'GIA' = $r2.Id; 'Titular' = $r2.Titular; 'Adreça' = $r2.Adreca
                'Estat' = $r2.Estat; 'Data informe' = $r2.DataInforme
                'Últim avís' = $r2.Ultim; 'Avisos' = $r2.Compte
                'Toca avui' = $(if ($r2.Toca) { 'SI' } else { 'NO' })
                'Situació' = $r2.Motiu; 'Correus' = $r2.Correus
            })
        }
        $dir = _ResolveOutputDir
        $path = _GetUniqueOutputPath $dir ('Recordatoris ' + $clau + ' ' + (Get-Date).ToString('yyyy-MM-dd') + '.csv')
        try {
            $out | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            $r = [System.Windows.Forms.MessageBox]::Show("CSV generat:`n$path`n`nVols obrir-lo?", 'Recordatoris', 'YesNo', 'Information')
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { try { Start-Process -FilePath $path | Out-Null } catch { } }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut escriure el CSV:`n$($_.Exception.Message)", 'Recordatoris', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $btnSend.add_Click({
        & $fn.Desa
        $tria = @(@($ui.Rows) | Where-Object { $_.Sel -and $_.Toca -and -not $_.Excloure })
        if ($tria.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No hi ha cap activitat marcada que toqui avui.', 'Recordatoris', 'OK', 'Information') | Out-Null
            return
        }
        $q = _QuotaLlegeix
        $viaAra = [string](Get-CorreuEina ('rec-' + $clau)).Via
        $ambOutlook = _CorreuViaEsOutlook $viaAra
        $rest = if ($ambOutlook) { [int]::MaxValue } else { _QuotaRestant $q }
        $prev = [Math]::Min($tria.Count, [Math]::Min([int]$cfg['maxPerTanda'], $rest))
        $msg = "S'enviaran fins a $prev correus (de $($tria.Count) marcats).`n`n" +
               "Per: " + (_CorreuViaText $viaAra) + " " + (_CorreuRemitentText $viaAra (Get-CorreuRemitent)) + " (es canvia a Configuraci" + [char]0x00F3 + ")`n" +
               "Topall per tanda: $($cfg['maxPerTanda'])`n" +
               $(if ($ambOutlook) { '' } else { "Quota d'aquest mes: $($q.enviats) / $($q.limit) (en queden $rest)`n" }) + "`n" +
               'Vols continuar?'
        if ([System.Windows.Forms.MessageBox]::Show($msg, 'Enviar recordatoris', 'YesNo', 'Question') -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $res = Invoke-RecordatorisTanda $clau $tria $false
        $resum = "Enviats: $($res.Enviats)`nFallats: $($res.Fallats)`nSense correu: $($res.SenseCorreu)"
        if ($res.Esborranys -gt 0) { $resum = "Desats a Esborranys de l'Outlook: $($res.Esborranys). Envia'ls des de l'Outlook: ja consten als recordatoris.`n" + $resum }
        if ($res.Aturat) { $resum += "`n`nATURAT: $($res.Motiu)" }
        $q2 = _QuotaLlegeix
        $resum += "`n`nQuota: $($q2.enviats) / $($q2.limit) aquest mes."
        [System.Windows.Forms.MessageBox]::Show($resum, 'Recordatoris', 'OK', 'Information') | Out-Null
        $estat.historial[$clau] = (_RecLlegeix).historial[$clau]
        & $fn.Refresca
    }.GetNewClosure())

    # Ordre de docking calcat de Controls periòdics: primer la graella (Fill),
    # després els panells de dalt i de baix.
    $tab.Controls.Add($grid)
    $tab.Controls.Add($top)
    $tab.Controls.Add($bot)
    & $fn.Refresca
}
