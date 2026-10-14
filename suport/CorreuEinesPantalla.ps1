#requires -Version 5.1
<#
  CorreuEinesPantalla.ps1 - La finestra "Correus de cada eina" (des de
  Configuracio): per on surt el correu de cada eina, a qui va, les adreces
  fixes, la CCO i el boto de prova. Tot el que decideix es a CorreuEines.ps1;
  aqui nomes la pantalla.

  NOMES DEFINEIX FUNCIONS.
#>

# El que es desa d'una eina a partir del que hi ha a la pantalla. Cada CCO
# (la de sempre i la de triar) nomes si no es la d'abans (si no, l'eina segueix
# amb la seva, encara que despres canvii). $ccoTria $null = la pantalla no en
# te (es deixa com estava). PURA.
function _CorreuEinaDeLaPantalla([string]$via, $dest, [string]$fixes, [string]$cco, [string]$ccoAbans, $ccoTria = $null, [string]$ccoTriaAbans = '') {
    $o = @{ via = $via; dest = @(@($dest) | Where-Object { $_ }); fixes = (_CorreuNetejaAdreces $fixes) }
    # Com a conjunt: les mateixes adreces en un altre ordre son la de sempre.
    $clau = { param($x) ((@(_CorreuLlistaAdreces $x) | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ';') }
    if ((& $clau $cco) -ne (& $clau $ccoAbans)) { $o['cco'] = _CorreuNetejaAdreces $cco }
    if ($null -ne $ccoTria -and (& $clau ([string]$ccoTria)) -ne (& $clau $ccoTriaAbans)) { $o['ccoTria'] = _CorreuNetejaAdreces ([string]$ccoTria) }
    return $o
}

# On va cada cosa de la fila "Per / A qui / Enviar prova" d'una eina. PURA (es
# prova a Linux). $amples: l'amplada que necessita cada casella (el text
# mesurat + la casella). Abans les amplades eren fixes (128 i 200) i la de
# "Persones autoritzades / tecnic" acabava a x=844, sota el boto "Enviar
# prova" (x=820): la finestra avisava de "Pantalla mal col·locada" (octubre
# 2026). Ara les caselles van una darrere l'altra amb el que ocupen de debo i,
# si encara no hi cabessin, la finestra s'eixampla: torna @{ Ampl; X (una per
# casella); Boto (x del boto) }.
function _CorreuFilaGeometria([int]$amplMin, [int[]]$amples) {
    $x0 = 316 + 44
    $xs = New-Object System.Collections.ArrayList
    $x = $x0
    foreach ($w in @($amples)) { [void]$xs.Add($x); $x += [int]$w + 10 }
    # El boto (120) va a 28 + 12 de la vora dreta; hi ha d'haver 8 px de marge.
    $ampl = [Math]::Max($amplMin, $x + 8 + 120 + 40)
    return @{ Ampl = $ampl; X = $xs.ToArray(); Boto = ($ampl - 28 - 132); Fi = $x }
}

function Show-CorreuEinesConfig {
    $cfg = Read-CorreuEines
    $eines = $Script:CorreuEines
    $destNoms = $Script:CorreuDestNoms
    $vies = $Script:CorreuVies

    $form = _NewForm
    $form.Text = 'Correus de cada eina'
    # L'amplada de cada casella, mesurada amb la lletra de la finestra.
    $amplesDest = @{}
    foreach ($d in @($destNoms.Keys)) { $amplesDest[[string]$d] = [System.Windows.Forms.TextRenderer]::MeasureText([string]$destNoms[$d], $form.Font).Width + 24 }
    $amplNec = 980
    foreach ($k in @($eines.Keys)) {
        $gm = _CorreuFilaGeometria 980 ([int[]]@(@($eines[$k].Dest) | ForEach-Object { [int]$amplesDest[[string]$_] }))
        $amplNec = [Math]::Max($amplNec, [int]$gm.Ampl)
    }
    $ampl = $amplNec
    # Tres linies per eina: via i a qui; fixes i CCO de sempre; CCO a triar.
    $altEina = 128
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
        $ccoTriaAbans = _CorreuCcoTriaAbansDe $k
        $g = New-Object System.Windows.Forms.GroupBox
        $g.Text = [string]$def.Nom
        $g.Location = New-Object System.Drawing.Point(14, $y)
        $g.Size = New-Object System.Drawing.Size(($ampl - 28), ($altEina - 6))
        $lv = New-Object System.Windows.Forms.Label
        $lv.Text = 'Per:'; $lv.Location = New-Object System.Drawing.Point(12, 27); $lv.Size = New-Object System.Drawing.Size(34, 20)
        $cbV = New-Object System.Windows.Forms.ComboBox
        $cbV.DropDownStyle = 'DropDownList'
        $cbV.Location = New-Object System.Drawing.Point(48, 24); $cbV.Size = New-Object System.Drawing.Size(256, 24)
        $vk = @($def.Vies)
        foreach ($v in $vk) { [void]$cbV.Items.Add([string]$vies[$v]) }
        $cbV.SelectedIndex = [math]::Max(0, [array]::IndexOf($vk, [string]$e.Via))
        $cbV.Tag = $vk
        $chks = [ordered]@{}
        $gm = _CorreuFilaGeometria $ampl ([int[]]@(@($def.Dest) | ForEach-Object { [int]$amplesDest[[string]$_] }))
        $x = 316
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
        $iD = 0
        foreach ($d in @($def.Dest)) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = [string]$destNoms[$d]
            $w = [int]$amplesDest[[string]$d]
            $cb.Location = New-Object System.Drawing.Point([int]$gm.X[$iD], 25); $cb.Size = New-Object System.Drawing.Size($w, 22)
            $cb.Checked = (@($e.Dest) -contains $d)
            [void]$g.Controls.Add($cb)
            $chks[$d] = $cb
            $iD++
        }
        $btnP = New-Object System.Windows.Forms.Button
        $btnP.Text = 'Enviar prova'
        $btnP.Location = New-Object System.Drawing.Point([int]$gm.Boto, 20); $btnP.Size = New-Object System.Drawing.Size(120, 28)
        _StyleSecondaryButton $btnP
        $lf = New-Object System.Windows.Forms.Label
        $lf.Text = 'Adreces fixes:'; $lf.Location = New-Object System.Drawing.Point(12, 61); $lf.Size = New-Object System.Drawing.Size(88, 20)
        $tbF = New-Object System.Windows.Forms.TextBox
        $tbF.Location = New-Object System.Drawing.Point(102, 58); $tbF.Size = New-Object System.Drawing.Size(360, 24)
        $tbF.Text = [string]$e.Fixes
        $lc = New-Object System.Windows.Forms.Label
        $lc.Text = 'CCO sempre:'; $lc.Location = New-Object System.Drawing.Point(476, 61); $lc.Size = New-Object System.Drawing.Size(92, 20)
        $tbC = New-Object System.Windows.Forms.TextBox
        $tbC.Location = New-Object System.Drawing.Point(570, 58); $tbC.Size = New-Object System.Drawing.Size(($ampl - 28 - 570 - 16), 24)
        $tbC.Text = if ($null -ne $e.Cco) { [string]$e.Cco } else { $ccoAbans }
        # La CCO A TRIAR: surt desmarcada cada vegada que s'envia a ma; en
        # automatic no hi va (vegeu _CorreuCcoOpcions, CorreuEines.ps1).
        $lt = New-Object System.Windows.Forms.Label
        $lt.Text = 'CCO a triar:'; $lt.Location = New-Object System.Drawing.Point(476, 95); $lt.Size = New-Object System.Drawing.Size(92, 20)
        $tbT = New-Object System.Windows.Forms.TextBox
        $tbT.Location = New-Object System.Drawing.Point(570, 92); $tbT.Size = New-Object System.Drawing.Size(($ampl - 28 - 570 - 16), 24)
        $tbT.Text = if ($null -ne $e.CcoTria) { [string]$e.CcoTria } else { $ccoTriaAbans }
        $lh = New-Object System.Windows.Forms.Label
        $lh.Text = ("Sempre: a tots els correus, tamb" + [char]0x00E9 + " els autom" + [char]0x00E0 + "tics. A triar: surten per marcar cada vegada que envies a m" + [char]0x00E0 + "; en autom" + [char]0x00E0 + "tic, no.")
        $lh.Location = New-Object System.Drawing.Point(12, 88); $lh.Size = New-Object System.Drawing.Size(450, 32)
        $lh.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
        foreach ($c in @($lv, $cbV, $btnP, $lf, $tbF, $lc, $tbC, $lt, $tbT, $lh)) { [void]$g.Controls.Add($c) }
        [void]$form.Controls.Add($g)
        $ctl[$k] = @{ Via = $cbV; Dest = $chks; Fixes = $tbF; Cco = $tbC; CcoAbans = $ccoAbans; CcoTria = $tbT; CcoTriaAbans = $ccoTriaAbans; Btn = $btnP }
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
            $nou.eines[$k] = _CorreuEinaDeLaPantalla $via $dest ([string]$c.Fixes.Text) ([string]$c.Cco.Text) ([string]$c.CcoAbans) ([string]$c.CcoTria.Text) ([string]$c.CcoTriaAbans)
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
    [void](_AddBrandHeader $form 'Correus de cada eina' ('Per on surten, a qui van i la CCO (sempre o a triar) ' + [char]0x00B7 + ' nom' + [char]0x00E9 + 's en aquest ordinador') 56)
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
