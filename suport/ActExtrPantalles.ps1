#requires -Version 5.1
<#
.SYNOPSIS
  Mode ACT_EXTR: les finestres (llistat d'activitats, capcalera i
  comprovacio de la documentacio). Nomes Windows.

  Part del modul ACT_EXTR, partit per responsabilitats (com Llicencia):
    ActExtrDades.ps1     logica del Decret, punts, plantilla i registre (PURES)
    ActExtrBlocs.ps1     composicio del document en blocs (PURA) + el .docx
    ActExtrPantalles.ps1 les finestres (WinForms)
    ActExtr.ps1          l'orquestrador (Invoke-ActExtrFlow) i la descripcio
  Tot va amb dot-source al mateix ambit (Motor.ps1 els carrega en aquest ordre).
#>

# ----------------------------------------------------------------------------
# UI (WinForms) - nomes en mode no-headless
# ----------------------------------------------------------------------------
# Llistat d'activitats extraordinaries amb el seu estat. Retorna:
#   @{ Action='new' } | @{ Action='open'; Id=<id> } | @{ Action='exit' }
function Show-ActExtrList($registry) {
    $form = _NewForm
    $form.Text = 'Activitats extraordinaries (ACT_EXTR)'
    $form.Size = New-Object System.Drawing.Size(820, 540)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Activitats extraordinaries. A la dreta hi ha l''estat (PENDENT si falta documentacio, TANCAT si tot esta lliurat).'
    $lbl.Location = New-Object System.Drawing.Point(15, 12)
    $lbl.Size = New-Object System.Drawing.Size(780, 20)
    $form.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(15, 40)
    $lv.Size = New-Object System.Drawing.Size(780, 400)
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.MultiSelect = $false
    $lv.Anchor = 'Top, Bottom, Left, Right'
    [void]$lv.Columns.Add('ID GIA', 90)
    [void]$lv.Columns.Add('Titular', 240)
    [void]$lv.Columns.Add('Activitat', 230)
    [void]$lv.Columns.Add('Dates', 120)
    [void]$lv.Columns.Add('Estat', 90)
    foreach ($a in @($registry.Activitats)) {
        $h = $a.Header
        $it = New-Object System.Windows.Forms.ListViewItem([string]$a.IdGia)
        [void]$it.SubItems.Add([string]$h.TITULAR)
        [void]$it.SubItems.Add([string]$h.ACTIVITAT)
        [void]$it.SubItems.Add([string]$h.DATES)
        $estat = if ($a.Estat) { [string]$a.Estat } else { 'pendent' }
        [void]$it.SubItems.Add($estat.ToUpper())
        $it.Tag = [string]$a.IdGia
        if ($estat -eq 'pendent') { $it.ForeColor = [System.Drawing.Color]::Firebrick }
        else                      { $it.ForeColor = [System.Drawing.Color]::ForestGreen }
        [void]$lv.Items.Add($it)
    }
    $form.Controls.Add($lv)

    # Enrere (torna al menu inicial) SEMPRE a baix a l'esquerra.
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(15, 455)
    $btnBack.Size = New-Object System.Drawing.Size(90, 32)
    $btnBack.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnBack)

    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = 'Nova activitat'
    $btnNew.Location = New-Object System.Drawing.Point(115, 455)
    $btnNew.Size = New-Object System.Drawing.Size(150, 32)
    $btnNew.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnNew)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = 'Obrir / continuar'
    $btnOpen.Location = New-Object System.Drawing.Point(275, 455)
    $btnOpen.Size = New-Object System.Drawing.Size(160, 32)
    $btnOpen.Anchor = 'Bottom, Left'
    $form.Controls.Add($btnOpen)

    # 'exit' (Enrere o tancar la finestra) fa que Invoke-ActExtrFlow torni al
    # menu inicial (el programa no es tanca; nomes es tanca des del Pas 1).
    $result = @{ Action = 'exit'; Id = $null }
    $btnNew.add_Click({ $result.Action = 'new'; $form.DialogResult = 'OK'; $form.Close() })
    $openAction = {
        if ($lv.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Tria una activitat de la llista.','Sense seleccio','OK','Information') | Out-Null
            return
        }
        $result.Action = 'open'; $result.Id = [string]$lv.SelectedItems[0].Tag
        $form.DialogResult = 'OK'; $form.Close()
    }
    $btnOpen.add_Click($openAction)
    $lv.add_DoubleClick($openAction)
    $btnBack.add_Click({ $result.Action = 'exit'; $form.DialogResult = 'Cancel'; $form.Close() })

    [void]$form.ShowDialog()
    return $result
}

# Capcalera ACT_EXTR (entrada manual). $preload pot precarregar valors.
# Retorna @{ Nav='next'|'back'; Data=@{ ID_GIA;EXP_NUM;ADRECA;ACTIVITAT;TITULAR;DATES;AFORAMENT } }
function Get-ActExtrHeader {
    param($preload = $null, [bool]$lockId = $false)
    $form = _NewForm
    $form.Text = 'Activitat extraordinaria - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(700, 420)
    $form.StartPosition = 'CenterScreen'

    $controls = @{}
    # $addRow afegeix una etiqueta + caixa de text a la posicio $y i retorna
    # la $y de la fila seguent (patro explicit, sense estat compartit).
    $addRow = {
        param($label, $key, $width, $yPos)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $label
        $l.Location = New-Object System.Drawing.Point(15, $yPos)
        $l.Size = New-Object System.Drawing.Size(180, 22)
        [void]$form.Controls.Add($l)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(200, ($yPos - 2))
        $tb.Size = New-Object System.Drawing.Size($width, 22)
        [void]$form.Controls.Add($tb)
        $controls[$key] = $tb
        return ($yPos + 36)
    }
    $y = 20
    $y = & $addRow 'ID GIA' 'ID_GIA' 300 $y
    $y = & $addRow "Num. d'expedient" 'EXP_NUM' 450 $y
    $y = & $addRow 'Titular' 'TITULAR' 450 $y
    $y = & $addRow 'Adreca (carrer i numero)' 'ADRECA' 450 $y
    $y = & $addRow 'Activitat' 'ACTIVITAT' 450 $y
    $y = & $addRow 'Dates' 'DATES' 450 $y
    $y = & $addRow 'Aforament autoritzat' 'AFORAMENT' 150 $y

    if ($preload) {
        foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','DATES','AFORAMENT') {
            $v = $null
            if ($preload -is [System.Collections.IDictionary]) { if ($preload.Contains($k)) { $v = $preload[$k] } }
            elseif ($preload.PSObject.Properties.Name -contains $k) { $v = $preload.$k }
            if ($null -ne $v) { $controls[$k].Text = [string]$v }
        }
    }
    if ($lockId) { $controls['ID_GIA'].ReadOnly = $true }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    [void]$form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(575, $y)
    $ok.Size = New-Object System.Drawing.Size(95, 30)
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    $data = $null
    $ok.add_Click({
        if ([string]::IsNullOrWhiteSpace($controls['ID_GIA'].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return
        }
        $af = 0
        if (-not [int]::TryParse((($controls['AFORAMENT'].Text) -replace '[^\d]',''), [ref]$af) -or $af -le 0) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un aforament autoritzat (nombre).",'Falta aforament','OK','Warning') | Out-Null
            return
        }
        $script:_actExtrHeaderData = @{
            ID_GIA    = $controls['ID_GIA'].Text.Trim()
            EXP_NUM   = $controls['EXP_NUM'].Text.Trim()
            TITULAR   = $controls['TITULAR'].Text.Trim()
            ADRECA    = $controls['ADRECA'].Text.Trim()
            ACTIVITAT = $controls['ACTIVITAT'].Text.Trim()
            DATES     = $controls['DATES'].Text.Trim()
            AFORAMENT = [string]$af
        }
        $form.DialogResult = 'OK'; $form.Close()
    })

    $script:_actExtrHeaderData = $null
    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { return [pscustomobject]@{ Nav='back' } }
    return [pscustomobject]@{ Nav='next'; Data=$script:_actExtrHeaderData }
}

# Pas 3: comprovacio de la documentacio. Mostra les preguntes de classificacio
# (esquerra) i, per cada punt, si APLICA i PER QUE + casella "lliurat" (dreta).
# Retorna @{ Action='req'|'fav'|'save'|'back'; Answers=@{...}; Delivered=@{...} }
function Edit-ActExtrDocumentacio {
    param($header, $answers = $null, $delivered = $null)

    $form = _NewForm
    $form.Text = 'Comprovacio de la documentacio - ' + ([string]$header.ID_GIA)
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(1160, 680)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)

    # Estat compartit en $script: (sempre accessible des de qualsevol handler,
    # independentment de l'abast d'invocacio). Aixi s'eviten els problemes
    # d'abast dels scriptblocks niats de WinForms.
    $script:_actExtrDelivered = @{}
    foreach ($p in $script:ActExtrPoints) {
        $script:_actExtrDelivered[$p.Key] = [bool](_GetActExtrDelivered $delivered $p.Key)
    }
    $script:_actExtrResult = @{ Action='back'; Answers=$null; Delivered=$null }

    # --- Barra inferior de botons (ancorada a baix; sempre visible) ---
    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = 'Bottom'
    $bottom.Height = 50
    $form.Controls.Add($bottom)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(12, 9)
    $btnBack.Size = New-Object System.Drawing.Size(90, 32)
    $bottom.Controls.Add($btnBack)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(110, 9)
    $btnSave.Size = New-Object System.Drawing.Size(110, 32)
    $bottom.Controls.Add($btnSave)

    $btnReq = New-Object System.Windows.Forms.Button
    $btnReq.Text = 'Generar requeriment'
    $btnReq.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 390), 9)
    $btnReq.Size = New-Object System.Drawing.Size(185, 32)
    $btnReq.Anchor = 'Top, Right'
    $bottom.Controls.Add($btnReq)

    $btnFav = New-Object System.Windows.Forms.Button
    $btnFav.Text = 'Generar informe favorable'
    $btnFav.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 200), 9)
    $btnFav.Size = New-Object System.Drawing.Size(190, 32)
    $btnFav.Anchor = 'Top, Right'
    $bottom.Controls.Add($btnFav)

    # --- Esquerra: preguntes de classificacio (amb scroll propi) ---
    $gbQ = New-Object System.Windows.Forms.GroupBox
    $gbQ.Text = 'Classificacio (Decret 112/2010)'
    $gbQ.Location = New-Object System.Drawing.Point(12, 12)
    $gbQ.Size = New-Object System.Drawing.Size(430, ($form.ClientSize.Height - 50 - 24))
    $gbQ.Anchor = 'Top, Bottom, Left'
    $form.Controls.Add($gbQ)

    # Panell intern desplacable: hi caben totes les preguntes encara que la
    # finestra sigui curta.
    $qPanel = New-Object System.Windows.Forms.Panel
    $qPanel.Location = New-Object System.Drawing.Point(8, 20)
    $qPanel.Size = New-Object System.Drawing.Size(414, ($gbQ.Height - 28))
    $qPanel.AutoScroll = $true
    $qPanel.Anchor = 'Top, Bottom, Left, Right'
    $gbQ.Controls.Add($qPanel)

    $aforamentInit = if ($answers) { [string](_GetActExtrAnswer $answers 'Aforament') } else { '' }
    if ([string]::IsNullOrWhiteSpace($aforamentInit)) { $aforamentInit = [string]$header.AFORAMENT }

    $lblAf = New-Object System.Windows.Forms.Label
    $lblAf.Text = 'Aforament autoritzat:'
    $lblAf.Location = New-Object System.Drawing.Point(8, 10)
    $lblAf.Size = New-Object System.Drawing.Size(160, 22)
    $qPanel.Controls.Add($lblAf)
    $tbAf = New-Object System.Windows.Forms.TextBox
    $tbAf.Location = New-Object System.Drawing.Point(175, 7)
    $tbAf.Size = New-Object System.Drawing.Size(110, 22)
    $tbAf.Text = $aforamentInit
    $qPanel.Controls.Add($tbAf)
    $script:_actExtrAfBox = $tbAf

    # Preguntes Si/No (clau -> etiqueta)
    $questions = @(
        @{ Key='Incendis';          Label='Incendis: inclosa a l''Art. 23 Llei 3/2010?' }
        @{ Key='Mobilitat';         Label='Mobilitat: cal estudi (Decret 344/2006)?' }
        @{ Key='ControlAccessos';   Label='Control d''acces: musical >=150 aforament?' }
        @{ Key='PauCatalunya';      Label='PAU: al Cataleg de Catalunya (Annex I, Cat. A)?' }
        @{ Key='PauLocal';          Label='PAU: al Cataleg local (Annex I, Cat. B)?' }
        @{ Key='EstablimentDotat';  Label='Higiene: l''establiment ja esta dotat de lavabos?' }
        @{ Key='ParcialSotaRasant'; Label='RC: l''activitat es du PARCIALMENT sota rasant?' }
        @{ Key='TotalSotaRasant';   Label='RC: l''activitat es du TOTALMENT sota rasant?' }
        @{ Key='HiHaLasers';        Label='Lasers: l''activitat preveu l''us de lasers?' }
    )
    $qRadios = @{}
    $qy = 44
    foreach ($q in $questions) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $q.Label
        $l.Location = New-Object System.Drawing.Point(8, $qy)
        $l.Size = New-Object System.Drawing.Size(395, 30)
        $qPanel.Controls.Add($l)
        $qy += 30
        # Cada parella Si/No va al SEU PROPI panell perque els RadioButton nomes
        # siguin exclusius DINS de la pregunta (si no, tot el contenidor formaria
        # un sol grup i nomes es podria triar una resposta a tota la columna).
        $rp = New-Object System.Windows.Forms.Panel
        $rp.Location = New-Object System.Drawing.Point(20, $qy)
        $rp.Size = New-Object System.Drawing.Size(200, 26)
        $rbSi = New-Object System.Windows.Forms.RadioButton
        $rbSi.Text = 'Si'
        $rbSi.Location = New-Object System.Drawing.Point(0, 2)
        $rbSi.Size = New-Object System.Drawing.Size(60, 22)
        $rbNo = New-Object System.Windows.Forms.RadioButton
        $rbNo.Text = 'No'
        $rbNo.Location = New-Object System.Drawing.Point(70, 2)
        $rbNo.Size = New-Object System.Drawing.Size(60, 22)
        $cur = if ($answers) { _ActExtrYesNo (_GetActExtrAnswer $answers $q.Key) } else { 'No' }
        if ($cur -eq 'Si') { $rbSi.Checked = $true } else { $rbNo.Checked = $true }
        $rp.Controls.Add($rbSi); $rp.Controls.Add($rbNo)
        $qPanel.Controls.Add($rp)
        $qRadios[$q.Key] = @{ Si=$rbSi; No=$rbNo }
        $qy += 34
    }
    $script:_actExtrRadios = $qRadios

    # --- Dreta: estat per punt (aplica + per que + lliurat) ---
    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Text = 'Documentacio per punt (APLICA / per que / lliurat):'
    $lblD.Location = New-Object System.Drawing.Point(455, 12)
    $lblD.Size = New-Object System.Drawing.Size(680, 20)
    $lblD.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lblD)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(455, 36)
    $panel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 455 - 12), ($form.ClientSize.Height - 50 - 48))
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($panel)

    # Llegeix les respostes actuals dels controls (via $script:, robust).
    $readAnswers = {
        $h = @{ Aforament = ($script:_actExtrAfBox.Text -replace '[^\d]','') }
        foreach ($k in $script:_actExtrRadios.Keys) {
            $h[$k] = if ($script:_actExtrRadios[$k].Si.Checked) { 'Si' } else { 'No' }
        }
        return $h
    }
    $script:_actExtrReadAnswers = $readAnswers

    # Refresca el panell de la dreta segons les respostes actuals.
    $refresh = {
        $ans = & $script:_actExtrReadAnswers
        $decret = Build-ActExtrDecret $ans
        $computed = Get-ActExtrComputed $decret
        $status = Get-ActExtrStatus $decret $computed $script:_actExtrDelivered
        $panel.SuspendLayout()
        $panel.Controls.Clear()
        $yy = 8
        foreach ($s in $status) {
            $title = New-Object System.Windows.Forms.Label
            $applyTxt = if ($s.Applies) { 'APLICA' } else { 'no aplica' }
            $title.Text = ('{0}  -  {1}' -f $s.Title, $applyTxt)
            $title.Location = New-Object System.Drawing.Point(8, $yy)
            $title.Size = New-Object System.Drawing.Size(640, 20)
            $title.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $title.ForeColor = if ($s.Applies) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::Gray }
            [void]$panel.Controls.Add($title)
            $yy += 22

            $reason = New-Object System.Windows.Forms.Label
            $reason.Text = $s.Reason
            $reason.Location = New-Object System.Drawing.Point(20, $yy)
            $reason.Size = New-Object System.Drawing.Size(645, 42)
            $reason.ForeColor = [System.Drawing.Color]::DimGray
            [void]$panel.Controls.Add($reason)
            $yy += 44

            if ($s.Applies -and $s.NeedsDelivery) {
                $cb = New-Object System.Windows.Forms.CheckBox
                $cb.Text = 'Lliurat correctament'
                $cb.Location = New-Object System.Drawing.Point(20, $yy)
                $cb.Size = New-Object System.Drawing.Size(300, 22)
                $cb.Checked = [bool]$script:_actExtrDelivered[$s.Key]
                $cb.Tag = $s.Key
                # Handler robust: el sender ve com a parametre i l'estat es
                # $script:_actExtrDelivered (no depen de l'abast d'invocacio).
                $cb.add_CheckedChanged({
                    param($snd, $e)
                    $script:_actExtrDelivered[[string]$snd.Tag] = [bool]$snd.Checked
                })
                [void]$panel.Controls.Add($cb)
                $yy += 28
            }
            $sepY = $yy + 2
            $sep = New-Object System.Windows.Forms.Label
            $sep.BorderStyle = 'Fixed3D'
            $sep.Location = New-Object System.Drawing.Point(8, $sepY)
            $sep.Size = New-Object System.Drawing.Size(650, 2)
            [void]$panel.Controls.Add($sep)
            $yy = $sepY + 10
        }
        $panel.ResumeLayout()
    }
    $script:_actExtrRefresh = $refresh

    $tbAf.add_TextChanged({ & $script:_actExtrRefresh })
    foreach ($k in $qRadios.Keys) {
        $qRadios[$k].Si.add_CheckedChanged({ & $script:_actExtrRefresh })
    }

    $finish = {
        param($action)
        $script:_actExtrResult = @{
            Action    = $action
            Answers   = (& $script:_actExtrReadAnswers)
            Delivered = $script:_actExtrDelivered.Clone()
        }
        $form.DialogResult = 'OK'; $form.Close()
    }
    $script:_actExtrFinish = $finish

    $btnSave.add_Click({ & $script:_actExtrFinish 'save' })
    $btnReq.add_Click({ & $script:_actExtrFinish 'req' })
    $btnFav.add_Click({
        $ans = & $script:_actExtrReadAnswers
        $decret = Build-ActExtrDecret $ans
        $computed = Get-ActExtrComputed $decret
        $defs = Get-ActExtrDeficiencies $decret $computed $script:_actExtrDelivered
        if (@($defs).Count -gt 0) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                "Encara hi ha punts aplicables PENDENTS de lliurar.`n`nL'informe favorable dona per fet que tot esta lliurat. Vols generar-lo igualment?",
                'Hi ha pendents', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }
        & $script:_actExtrFinish 'fav'
    })
    $btnBack.add_Click({ $script:_actExtrResult = @{ Action='back' }; $form.DialogResult='Cancel'; $form.Close() })

    & $refresh
    [void]$form.ShowDialog()
    return $script:_actExtrResult
}

# Helpers d'acces a respostes/lliurats precarregats (hashtable o PSObject).
function _GetActExtrAnswer($answers, [string]$key) {
    if ($null -eq $answers) { return $null }
    if ($answers -is [System.Collections.IDictionary]) { if ($answers.Contains($key)) { return $answers[$key] }; return $null }
    if ($answers.PSObject.Properties.Name -contains $key) { return $answers.$key }
    return $null
}
function _GetActExtrDelivered($delivered, [string]$key) {
    if ($null -eq $delivered) { return $false }
    if ($delivered -is [System.Collections.IDictionary]) { if ($delivered.Contains($key)) { return [bool]$delivered[$key] }; return $false }
    if ($delivered.PSObject.Properties.Name -contains $key) { return [bool]$delivered.$key }
    return $false
}
