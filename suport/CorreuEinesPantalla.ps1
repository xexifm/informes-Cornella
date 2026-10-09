#requires -Version 5.1
<#
  CorreuEinesPantalla.ps1 - La finestra "Correus de cada eina" (des de
  Configuracio): per on surt el correu de cada eina, a qui va, les adreces
  fixes, la CCO i el boto de prova. Tot el que decideix es a CorreuEines.ps1;
  aqui nomes la pantalla.

  NOMES DEFINEIX FUNCIONS.
#>

function _CorreuCcoAbansDe([string]$clau) {
    $f = $Script:CorreuCcoAbans[$clau]
    if ($null -eq $f) { return '' }
    try { return (_CorreuNetejaAdreces ([string](& $f))) } catch { return '' }
}

# El que es desa d'una eina a partir del que hi ha a la pantalla. La CCO nomes
# si no es la de sempre (si no, l'eina segueix amb la seva, encara que despres
# canvii). PURA.
function _CorreuEinaDeLaPantalla([string]$via, $dest, [string]$fixes, [string]$cco, [string]$ccoAbans) {
    $o = @{ via = $via; dest = @(@($dest) | Where-Object { $_ }); fixes = (_CorreuNetejaAdreces $fixes) }
    $c = _CorreuNetejaAdreces $cco
    # Com a conjunt: les mateixes adreces en un altre ordre son la de sempre.
    $clau = { param($x) ((@(_CorreuLlistaAdreces $x) | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ';') }
    if ((& $clau $cco) -ne (& $clau $ccoAbans)) { $o['cco'] = $c }
    return $o
}

function Show-CorreuEinesConfig {
    $cfg = Read-CorreuEines
    $eines = $Script:CorreuEines
    $destNoms = $Script:CorreuDestNoms
    $vies = $Script:CorreuVies

    $form = _NewForm
    $form.Text = 'Correus de cada eina'
    $ampl = 980
    $altEina = 100
    $yIni = 150
    $form.ClientSize = New-Object System.Drawing.Size($ampl, ($yIni + $altEina * @($eines.Keys).Count + 60))
    $form.StartPosition = 'CenterScreen'

    # ---- El correu de prova (comu a totes les eines) ----------------------
    $grpP = New-Object System.Windows.Forms.GroupBox
    $grpP.Text = 'Correu de prova'
    $grpP.Location = New-Object System.Drawing.Point(14, 66)
    $grpP.Size = New-Object System.Drawing.Size(($ampl - 28), 74)
    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = 'ID GIA de les dades:'; $l1.Location = New-Object System.Drawing.Point(12, 30); $l1.Size = New-Object System.Drawing.Size(120, 20)
    $tbGia = New-Object System.Windows.Forms.TextBox
    $tbGia.Location = New-Object System.Drawing.Point(134, 27); $tbGia.Size = New-Object System.Drawing.Size(70, 24)
    $tbGia.Text = [string]$cfg.prova.gia
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = 'El rep:'; $l2.Location = New-Object System.Drawing.Point(222, 30); $l2.Size = New-Object System.Drawing.Size(48, 20)
    $tbDesti = New-Object System.Windows.Forms.TextBox
    $tbDesti.Location = New-Object System.Drawing.Point(272, 27); $tbDesti.Size = New-Object System.Drawing.Size(270, 24)
    $tbDesti.Text = [string]$cfg.prova.desti
    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = ("La prova va NOM" + [char]0x00C9 + "S a aquesta adre" + [char]0x00E7 + "a (mai al titular ni a la CCO) i diu al davant a qui hauria anat de debo.")
    $l3.Location = New-Object System.Drawing.Point(556, 22); $l3.Size = New-Object System.Drawing.Size(390, 40)
    $l3.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    foreach ($c in @($l1, $tbGia, $l2, $tbDesti, $l3)) { [void]$grpP.Controls.Add($c) }
    [void]$form.Controls.Add($grpP)

    # ---- Una fila per eina -------------------------------------------------
    $ctl = [ordered]@{}
    $y = $yIni
    foreach ($k in @($eines.Keys)) {
        $def = $eines[$k]
        $raw = if ($cfg.eines.ContainsKey($k)) { $cfg.eines[$k] } else { $null }
        $e = _CorreuEinaNormalitza $raw $k (Get-CorreuVia)
        $ccoAbans = _CorreuCcoAbansDe $k
        $g = New-Object System.Windows.Forms.GroupBox
        $g.Text = [string]$def.Nom
        $g.Location = New-Object System.Drawing.Point(14, $y)
        $g.Size = New-Object System.Drawing.Size(($ampl - 28), ($altEina - 6))
        $lv = New-Object System.Windows.Forms.Label
        $lv.Text = 'Per:'; $lv.Location = New-Object System.Drawing.Point(12, 27); $lv.Size = New-Object System.Drawing.Size(34, 20)
        $cbV = New-Object System.Windows.Forms.ComboBox
        $cbV.DropDownStyle = 'DropDownList'
        $cbV.Location = New-Object System.Drawing.Point(48, 24); $cbV.Size = New-Object System.Drawing.Size(270, 24)
        $vk = @($def.Vies)
        foreach ($v in $vk) { [void]$cbV.Items.Add([string]$vies[$v]) }
        $cbV.SelectedIndex = [math]::Max(0, [array]::IndexOf($vk, [string]$e.Via))
        $cbV.Tag = $vk
        $chks = [ordered]@{}
        $x = 336
        $la = New-Object System.Windows.Forms.Label
        $la.Text = 'A qui:'; $la.Location = New-Object System.Drawing.Point($x, 27); $la.Size = New-Object System.Drawing.Size(42, 20)
        [void]$g.Controls.Add($la)
        $x += 44
        if (@($def.Dest).Count -eq 0) {
            $ln = New-Object System.Windows.Forms.Label
            $ln.Text = ("Nom" + [char]0x00E9 + "s a les adreces fixes (no " + [char]0x00E9 + "s de cap activitat)")
            $ln.Location = New-Object System.Drawing.Point($x, 27); $ln.Size = New-Object System.Drawing.Size(330, 20)
            [void]$g.Controls.Add($ln)
        }
        foreach ($d in @($def.Dest)) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = [string]$destNoms[$d]
            $w = if ($d -eq 'autoritzats') { 200 } else { 128 }
            $cb.Location = New-Object System.Drawing.Point($x, 25); $cb.Size = New-Object System.Drawing.Size($w, 22)
            $cb.Checked = (@($e.Dest) -contains $d)
            [void]$g.Controls.Add($cb)
            $chks[$d] = $cb
            $x += $w + 4
        }
        $btnP = New-Object System.Windows.Forms.Button
        $btnP.Text = 'Enviar prova'
        $btnP.Location = New-Object System.Drawing.Point(($ampl - 28 - 132), 20); $btnP.Size = New-Object System.Drawing.Size(120, 28)
        _StyleSecondaryButton $btnP
        $lf = New-Object System.Windows.Forms.Label
        $lf.Text = 'Adreces fixes:'; $lf.Location = New-Object System.Drawing.Point(12, 61); $lf.Size = New-Object System.Drawing.Size(88, 20)
        $tbF = New-Object System.Windows.Forms.TextBox
        $tbF.Location = New-Object System.Drawing.Point(102, 58); $tbF.Size = New-Object System.Drawing.Size(360, 24)
        $tbF.Text = [string]$e.Fixes
        $lc = New-Object System.Windows.Forms.Label
        $lc.Text = 'CCO:'; $lc.Location = New-Object System.Drawing.Point(476, 61); $lc.Size = New-Object System.Drawing.Size(38, 20)
        $tbC = New-Object System.Windows.Forms.TextBox
        $tbC.Location = New-Object System.Drawing.Point(516, 58); $tbC.Size = New-Object System.Drawing.Size(420, 24)
        $tbC.Text = if ($null -ne $e.Cco) { [string]$e.Cco } else { $ccoAbans }
        foreach ($c in @($lv, $cbV, $btnP, $lf, $tbF, $lc, $tbC)) { [void]$g.Controls.Add($c) }
        [void]$form.Controls.Add($g)
        $ctl[$k] = @{ Via = $cbV; Dest = $chks; Fixes = $tbF; Cco = $tbC; CcoAbans = $ccoAbans; Btn = $btnP }
        $y += $altEina
    }

    # Llegeix la pantalla i ho desa. Torna $true si ha pogut.
    $st = @{ Cfg = $cfg; Ctl = $ctl; Gia = $tbGia; Desti = $tbDesti }
    $desa = {
        $nou = @{ eines = @{}; prova = @{ gia = ([string]$st.Gia.Text).Trim(); desti = (_CorreuNetejaAdreces ([string]$st.Desti.Text)) } }
        foreach ($k in @($st.Ctl.Keys)) {
            $c = $st.Ctl[$k]
            $vk = @($c.Via.Tag)
            $via = if ($c.Via.SelectedIndex -ge 0) { [string]$vk[$c.Via.SelectedIndex] } else { [string]$vk[0] }
            $dest = @(@($c.Dest.Keys) | Where-Object { $c.Dest[$_].Checked })
            $nou.eines[$k] = _CorreuEinaDeLaPantalla $via $dest ([string]$c.Fixes.Text) ([string]$c.Cco.Text) ([string]$c.CcoAbans)
        }
        try { Save-CorreuEines $nou; $st.Cfg = $nou; return $true } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n" + $_.Exception.Message, 'Correus de cada eina', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    foreach ($k in @($ctl.Keys)) {
        $clau = [string]$k
        $ctl[$k].Btn.add_Click({
            # La prova surt amb el que hi ha a la pantalla: primer es desa.
            if (-not (& $desa)) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try { $r = Send-CorreuProva $clau ([string]$st.Gia.Text) ([string]$st.Desti.Text) } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
            [System.Windows.Forms.MessageBox]::Show([string]$r.Text, 'Correu de prova', 'OK', $(if ($r.Ok) { 'Information' } else { 'Warning' })) | Out-Null
        }.GetNewClosure())
    }

    $peu = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) @(@{ Nom = 'Desar'; Text = 'Desar'; Estil = 'primari' }) ($y + 8)
    $peu.Desar.add_Click({
        if (& $desa) {
            [System.Windows.Forms.MessageBox]::Show("Desat en aquest ordinador. Ja ho fan servir totes les eines.", 'Correus de cada eina', 'OK', 'Information') | Out-Null
            $form.Close()
        }
    }.GetNewClosure())
    [void](_AddBrandHeader $form 'Correus de cada eina' ('Per on surten, a qui van i la CCO ' + [char]0x00B7 + ' nom' + [char]0x00E9 + 's en aquest ordinador') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}

# ----------------------------------------------------------------------------
# "TEXTOS DEL CORREU": els de TOTES les eines (octubre 2026)
# ----------------------------------------------------------------------------
# L'usuari: "l'eina Textos del correu no ha de servir nomes per als correus del
# mobil, sino per modificar tots els textos dels correus. Similar a Editar
# catalegs, que pots triar entre informes". Les eines son les del registre
# ($Script:CorreuEines) i cada una obre el seu editor, que ja existia (tots
# passen per Show-EditorAssumpteCos): aqui nomes es tria.
$Script:CorreuTextosEditors = [ordered]@{
    'mobil'            = @{ Desc = "El correu de l'informe al titular (el m" + [char]0x00F2 + "bil i Enviar correu). Es publica amb Actualitzar."; Obre = { Invoke-EmailTextos } }
    'rec-requeriments' = @{ Desc = "El recordatori a qui t" + [char]0x00E9 + " un requeriment pendent."; Obre = { [void](Invoke-RecordatorisTextos 'requeriments') } }
    'rec-precintes'    = @{ Desc = "El recordatori a qui t" + [char]0x00E9 + " l'activitat precintada o suspesa."; Obre = { [void](Invoke-RecordatorisTextos 'precintes') } }
    'controls'         = @{ Desc = "L'av" + [char]0x00ED + "s de control peri" + [char]0x00F2 + "dic pendent (esborranys a l'Outlook)."; Obre = { Invoke-ControlsCpEmailTextos } }
}

# Les eines que tenen text per editar, en l'ordre del registre. PURA.
function _CorreuTextosEines {
    $out = New-Object System.Collections.ArrayList
    foreach ($k in @($Script:CorreuEines.Keys)) {
        if ($Script:CorreuTextosEditors.Contains($k)) { [void]$out.Add([string]$k) }
    }
    return $out.ToArray()
}

function Invoke-TextosCorreu {
    $claus = @(_CorreuTextosEines)
    $eines = $Script:CorreuEines
    $eds = $Script:CorreuTextosEditors
    $form = _NewForm
    $form.Text = 'Textos del correu'
    $form.ClientSize = New-Object System.Drawing.Size(620, 250)
    $form.StartPosition = 'CenterScreen'
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Quin correu vols editar?'
    $l.Location = New-Object System.Drawing.Point(20, 76); $l.Size = New-Object System.Drawing.Size(580, 20)
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.DropDownStyle = 'DropDownList'
    $cb.Location = New-Object System.Drawing.Point(20, 100); $cb.Size = New-Object System.Drawing.Size(580, 24)
    foreach ($k in $claus) { [void]$cb.Items.Add([string]$eines[$k].Nom) }
    $d = New-Object System.Windows.Forms.Label
    $d.Location = New-Object System.Drawing.Point(20, 134); $d.Size = New-Object System.Drawing.Size(580, 40)
    $d.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    foreach ($c in @($l, $cb, $d)) { [void]$form.Controls.Add($c) }
    $cb.add_SelectedIndexChanged({ if ($cb.SelectedIndex -ge 0) { $d.Text = [string]$eds[$claus[$cb.SelectedIndex]].Desc } }.GetNewClosure())
    if ($claus.Count -gt 0) { $cb.SelectedIndex = 0 }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) @(@{ Nom = 'Obre'; Text = 'Editar el text'; Estil = 'primari'; Intro = $true }) 190
    # L'editor s'obre DAMUNT d'aquesta finestra i en tornar s'hi pot triar un
    # altre correu: com Editar catalegs, sense tornar al menu.
    $peu.Obre.add_Click({
        if ($cb.SelectedIndex -lt 0) { return }
        try { & $eds[$claus[$cb.SelectedIndex]].Obre } catch {
            [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Textos del correu', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    [void](_AddBrandHeader $form 'Textos del correu' ('Els textos de tots els correus que envia el programa') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}
