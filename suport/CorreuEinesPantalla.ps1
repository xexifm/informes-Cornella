#requires -Version 5.1
<#
  CorreuEinesPantalla.ps1 - La finestra "Configuracio correu electronic"
  (abans "Correus de cada eina", des de Configuracio): des d'on surten (el
  remitent), per on surt el correu de cada eina, a qui va, les adreces fixes,
  la CCO, el text, el boto de prova i la fila de l'app del mobil. Tot el que decideix es a CorreuEines.ps1;
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
# si encara no hi cabessin, la finestra s'eixampla. Al final, dos botons:
# "Editar el text" (130) i "Enviar prova" (120). Torna @{ Ampl; X (una per
# casella); BotoText; Boto (x de cada boto) }.
function _CorreuFilaGeometria([int]$amplMin, [int[]]$amples) {
    $x0 = 316 + 44
    $xs = New-Object System.Collections.ArrayList
    $x = $x0
    foreach ($w in @($amples)) { [void]$xs.Add($x); $x += [int]$w + 10 }
    # El de prova va a 28 + 12 de la vora dreta, el del text 10 mes a
    # l'esquerra; hi ha d'haver 8 px de marge amb l'ultima casella.
    $ampl = [Math]::Max($amplMin, $x + 8 + 130 + 10 + 120 + 40)
    $boto = $ampl - 28 - 132
    return @{ Ampl = $ampl; X = $xs.ToArray(); BotoText = ($boto - 10 - 130); Boto = $boto; Fi = $x }
}

# LA FILA DE L'APP DEL MOBIL (octubre 2026). L'usuari: "App del mobil (web) no
# se per que te un boto, quan podria ser un apartat mes com la resta". Es EL
# MATEIX CORREU que la fila 'mobil' (l'informe de visita al titular, amb el
# mateix text), pero l'envia l'app web del mobil per EmailJS: no surt d'aquest
# PC, i per aixo no te adreces fixes ni prova. A qui va i la CCO es
# desen amb Save-CorreuMobilConfig (CorreuMobil.ps1). $m: _CorreuMobilLlegeix
# ($null si no s'han pogut llegir els textos; $err diu per que). Torna els
# controls.
function _CorreuFilaAppMobil($form, [int]$y, [int]$ampl, [int]$alt, $amplesDest, $m, [string]$err) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = ('Informe de visita ' + [char]0x00B7 + ' des del m' + [char]0x00F2 + 'bil (app web)')
    $g.Location = New-Object System.Drawing.Point(14, $y)
    $g.Size = New-Object System.Drawing.Size(($ampl - 28), ($alt - 6))
    $mk = {
        param($tipus, $text, $x, $yy, $w, $h)
        $c = New-Object ('System.Windows.Forms.' + $tipus)
        if ($null -ne $text) { $c.Text = $text }
        $c.Location = New-Object System.Drawing.Point($x, $yy); $c.Size = New-Object System.Drawing.Size($w, $h)
        [void]$g.Controls.Add($c)
        return $c
    }
    [void](& $mk 'Label' 'Per:' 12 27 34 20)
    [void](& $mk 'Label' ("EmailJS, des del m" + [char]0x00F2 + "bil") 48 27 256 20)
    [void](& $mk 'Label' 'A qui:' 316 27 42 20)
    $wT = [int]$amplesDest['titular']; $wR = [int]$amplesDest['representant']
    $gm = _CorreuFilaGeometria $ampl ([int[]]@($wT, $wR))
    $cbTi = & $mk 'CheckBox' ([string]$Script:CorreuDestNoms['titular']) ([int]$gm.X[0]) 25 $wT 22
    $cbRe = & $mk 'CheckBox' ([string]$Script:CorreuDestNoms['representant']) ([int]$gm.X[1]) 25 $wR 22
    $btnT = & $mk 'Button' 'Editar el text' ([int]$gm.BotoText) 20 130 28
    $btnP = & $mk 'Button' 'Prompt...' ([int]$gm.Boto) 20 120 28
    _StyleSecondaryButton $btnT; _StyleSecondaryButton $btnP
    $lq = & $mk 'Label' ("El mateix correu i el mateix text que la fila de dalt, per" + [char]0x00F2 + " l'envia l'app del m" + [char]0x00F2 + "bil. A qui va hi arriba amb Actualitzar.") 12 56 450 32
    [void](& $mk 'Label' 'CCO sempre:' 476 61 92 20)
    $tbS = & $mk 'TextBox' '' 570 58 ($ampl - 28 - 570 - 16) 24
    [void](& $mk 'Label' 'CCO a triar:' 476 95 92 20)
    $tbT = & $mk 'TextBox' '' 570 92 ($ampl - 28 - 570 - 16) 24
    $lh = & $mk 'Label' ("La CCO es desa en aquest PC i al teu Drive (mai al GitHub). Sempre: surten marcades al m" + [char]0x00F2 + "bil; a triar: desmarcades.") 12 88 450 32
    foreach ($c in @($lq, $lh)) { $c.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90) }
    if ($null -ne $m) {
        $cbTi.Checked = (@($m.Dest) -contains 'titular'); $cbRe.Checked = (@($m.Dest) -contains 'representant')
        $tbS.Text = [string]$m.Sempre; $tbT.Text = [string]$m.Tria
    } else {
        foreach ($c in @($cbTi, $cbRe, $tbS, $tbT, $btnP)) { $c.Enabled = $false }
        $lq.Text = ("No s'han pogut llegir els textos del m" + [char]0x00F2 + "bil: " + $err)
        $lq.ForeColor = [System.Drawing.Color]::FromArgb(160, 40, 40)
    }
    [void]$form.Controls.Add($g)
    # El que diu la pantalla, com _CorreuMobilLlegeix.
    $llegeix = {
        $d = @(); if ($cbTi.Checked) { $d += 'titular' }; if ($cbRe.Checked) { $d += 'representant' }
        return @{ Sempre = (_CorreuNetejaAdreces ([string]$tbS.Text)); Tria = (_CorreuNetejaAdreces ([string]$tbT.Text)); Dest = $d }
    }.GetNewClosure()
    $btnP.add_Click({ Show-CorreuMobilPrompt (& $llegeix) }.GetNewClosure())
    return @{ Llegeix = $llegeix; BtnText = $btnT }
}

# LA FILA DEL REMITENT i la del correu de prova, l'una al costat de l'altra.
# PURA. Torna l'amplada de cada grup i on comenca el segon.
function _CorreuCapGeometria([int]$ampl) {
    $w = [int][Math]::Floor(($ampl - 28 - 12) / 2)
    return @{ W = $w; X2 = (14 + $w + 12) }
}

# $publica: el que fa Actualitzar, per publicar a qui va el correu del mobil
# (Configuracio hi passa Invoke-ActualitzarPrograma; aqui no es pot cridar
# sense un cicle de dependencies).
function Show-CorreuEinesConfig([scriptblock]$publica = $null) {
    $cfg = Read-CorreuEines
    $eines = $Script:CorreuEines
    $destNoms = $Script:CorreuDestNoms
    $vies = $Script:CorreuVies
    $titol = ('Configuraci' + [char]0x00F3 + ' correu electr' + [char]0x00F2 + 'nic')
    # L'app del mobil: el que hi ha ara (per desar nomes si canvia).
    $mobAbans = $null; $mobErr = ''
    try { $mobAbans = _CorreuMobilLlegeix (_LoadEmailTextos) } catch { $mobErr = [string]$_.Exception.Message }
    $remAbans = Get-CorreuRemitent

    $form = _NewForm
    $form.Text = $titol
    # L'amplada de cada casella, mesurada amb la lletra de la finestra.
    $amplesDest = @{}
    foreach ($d in @($destNoms.Keys)) { $amplesDest[[string]$d] = [System.Windows.Forms.TextRenderer]::MeasureText([string]$destNoms[$d], $form.Font).Width + 24 }
    $amplNec = 1100
    foreach ($k in @($eines.Keys)) {
        $gm = _CorreuFilaGeometria 1100 ([int[]]@(@($eines[$k].Dest) | ForEach-Object { [int]$amplesDest[[string]$_] }))
        $amplNec = [Math]::Max($amplNec, [int]$gm.Ampl)
    }
    $ampl = $amplNec
    # Tres linies per eina: via i a qui; fixes i CCO de sempre; CCO a triar.
    $altEina = 128
    $altCap = 92
    $yIni = 66 + $altCap + 10
    # Una fila per eina + la de l'app del mobil + el peu (60). Si no hi cap,
    # la finestra fa scroll (_NewForm).
    $form.ClientSize = New-Object System.Drawing.Size($ampl, ($yIni + $altEina * (@($eines.Keys).Count + 1) + 60))
    $form.StartPosition = 'CenterScreen'
    $cap = _CorreuCapGeometria $ampl
    $gris = [System.Drawing.Color]::FromArgb(90, 90, 90)

    # ---- DES D'ON surten (el remitent de l'Outlook) -----------------------
    # Abans era a Configuracio (l'usuari, octubre 2026: "allà dins poses
    # l'adreça des d'on s'envien els correus"). L'usuari: "a la feina puc
    # enviar des d'adreces diferents": editable, una bustia compartida no surt
    # als comptes i s'escriu a ma. Es desa a settings.json ('CorreuRemitent').
    $grpR = New-Object System.Windows.Forms.GroupBox
    $grpR.Text = ("Des d'on surten (Outlook)")
    $grpR.Location = New-Object System.Drawing.Point(14, 66)
    $grpR.Size = New-Object System.Drawing.Size($cap.W, $altCap)
    $lblRem = New-Object System.Windows.Forms.Label
    $lblRem.Text = 'Des de:'; $lblRem.Location = New-Object System.Drawing.Point(12, 30); $lblRem.Size = New-Object System.Drawing.Size(50, 20)
    $btnComptes = New-Object System.Windows.Forms.Button
    $btnComptes.Text = "Comptes de l'Outlook"
    _StyleSecondaryButton $btnComptes
    # Mesurat amb la lletra del boto (amb 140 fixos es tallava "Comptes de
    # l'Outlo", octubre 2026), arrimat a la dreta del grup.
    $wComptes = [Math]::Max(150, [System.Windows.Forms.TextRenderer]::MeasureText($btnComptes.Text, $btnComptes.Font).Width + 24)
    $btnComptes.Size = New-Object System.Drawing.Size($wComptes, 28)
    $btnComptes.Location = New-Object System.Drawing.Point(($cap.W - 12 - $wComptes), 25)
    $cbRem = New-Object System.Windows.Forms.ComboBox
    $cbRem.DropDownStyle = 'DropDown'
    $cbRem.Location = New-Object System.Drawing.Point(64, 27)
    $cbRem.Size = New-Object System.Drawing.Size([Math]::Max(140, ($cap.W - 12 - $wComptes - 6 - 64)), 24)
    $cbRem.Text = $remAbans
    $lRemNota = New-Object System.Windows.Forms.Label
    $lRemNota.Text = ("Buit = el compte per defecte de l'Outlook. Una b" + [char]0x00FA + "stia compartida s'escriu a m" + [char]0x00E0 + " (cal tenir-hi perm" + [char]0x00ED + "s). Amb EmailJS no compta.")
    $lRemNota.Location = New-Object System.Drawing.Point(12, 56); $lRemNota.Size = New-Object System.Drawing.Size(($cap.W - 24), 32)
    $lRemNota.ForeColor = $gris
    $btnComptes.add_Click({
        try { $comptes = Get-OutlookComptes } catch {
            [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Comptes de l''Outlook', 'OK', 'Warning') | Out-Null
            return
        }
        $cbRem.Items.Clear()
        foreach ($c in @($comptes)) { [void]$cbRem.Items.Add($c) }
        if ($cbRem.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("L'Outlook no t" + [char]0x00E9 + " cap compte amb adre" + [char]0x00E7 + "a. Si envies des d'una b" + [char]0x00FA + "stia compartida, escriu-ne l'adre" + [char]0x00E7 + "a.", 'Comptes de l''Outlook', 'OK', 'Information') | Out-Null
            return
        }
        $cbRem.DroppedDown = $true
    }.GetNewClosure())
    foreach ($c in @($lblRem, $cbRem, $btnComptes, $lRemNota)) { [void]$grpR.Controls.Add($c) }
    [void]$form.Controls.Add($grpR)

    # ---- El correu de prova (comu a totes les eines) ----------------------
    $grpP = New-Object System.Windows.Forms.GroupBox
    $grpP.Text = 'Correu de prova'
    $grpP.Location = New-Object System.Drawing.Point($cap.X2, 66)
    $grpP.Size = New-Object System.Drawing.Size($cap.W, $altCap)
    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = 'ID GIA de les dades:'; $l1.Location = New-Object System.Drawing.Point(12, 30); $l1.Size = New-Object System.Drawing.Size(120, 20)
    $tbGia = New-Object System.Windows.Forms.TextBox
    $tbGia.Location = New-Object System.Drawing.Point(134, 27); $tbGia.Size = New-Object System.Drawing.Size(70, 24)
    $tbGia.Text = [string]$cfg.prova.gia
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = 'El rep:'; $l2.Location = New-Object System.Drawing.Point(222, 30); $l2.Size = New-Object System.Drawing.Size(48, 20)
    $tbDesti = New-Object System.Windows.Forms.TextBox
    $tbDesti.Location = New-Object System.Drawing.Point(272, 27); $tbDesti.Size = New-Object System.Drawing.Size(($cap.W - 272 - 12), 24)
    $tbDesti.Text = [string]$cfg.prova.desti
    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = ("La prova va NOM" + [char]0x00C9 + "S a aquesta adre" + [char]0x00E7 + "a (mai al titular ni a la CCO) i diu al davant a qui hauria anat de debo.")
    $l3.Location = New-Object System.Drawing.Point(12, 56); $l3.Size = New-Object System.Drawing.Size(($cap.W - 24), 32)
    $l3.ForeColor = $gris
    foreach ($c in @($l1, $tbGia, $l2, $tbDesti, $l3)) { [void]$grpP.Controls.Add($c) }
    [void]$form.Controls.Add($grpP)

    # ---- Una fila per eina -------------------------------------------------
    $ctl = [ordered]@{}
    $mob = $null
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
        # EL TEXT del correu d'aquesta eina (l'usuari: "s'ha de poder configurar
        # tot el correu des d'aquesta pantalla"): el mateix editor que obre
        # Textos del correu.
        $btnT = New-Object System.Windows.Forms.Button
        $btnT.Text = 'Editar el text'
        $btnT.Location = New-Object System.Drawing.Point([int]$gm.BotoText, 20); $btnT.Size = New-Object System.Drawing.Size(130, 28)
        _StyleSecondaryButton $btnT
        $btnT.Enabled = $Script:CorreuTextosEditors.Contains($k)
        [void]$g.Controls.Add($btnT)
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
        $ctl[$k] = @{ Via = $cbV; Dest = $chks; Fixes = $tbF; Cco = $tbC; CcoAbans = $ccoAbans; CcoTria = $tbT; CcoTriaAbans = $ccoTriaAbans; Btn = $btnP; BtnText = $btnT }
        $y += $altEina
        # L'app del mobil just a sota d'"Enviar correu": es el mateix correu.
        if ($k -eq 'mobil') {
            $mob = _CorreuFilaAppMobil $form $y $ampl $altEina $amplesDest $mobAbans $mobErr
            $y += $altEina
        }
    }
    if ($null -eq $mob) {
        $mob = _CorreuFilaAppMobil $form $y $ampl $altEina $amplesDest $mobAbans $mobErr
        $y += $altEina
    }

    # Llegeix la pantalla i ho desa tot: les eines (correus.json), el remitent
    # (settings.json) i l'app del mobil (nomes si ha canviat). Torna $true si
    # ha pogut.
    $st = @{ Cfg = $cfg; Ctl = $ctl; Gia = $tbGia; Desti = $tbDesti; Rem = $cbRem; RemAbans = $remAbans; Mob = $mob; MobAbans = $mobAbans; MobDest = $false; MobText = ''; MobOk = $true; Titol = $titol }
    $desa = {
        $rem = ([string]$st.Rem.Text).Trim()
        if (-not (_CorreuRemitentValid $rem)) {
            [System.Windows.Forms.MessageBox]::Show(("'" + $rem + "' no " + [char]0x00E9 + "s una adre" + [char]0x00E7 + "a de correu. Deixa-ho en blanc per fer servir el compte per defecte de l'Outlook."), $st.Titol, 'OK', 'Warning') | Out-Null
            return $false
        }
        $nou = @{ eines = @{}; prova = @{ gia = ([string]$st.Gia.Text).Trim(); desti = (_CorreuNetejaAdreces ([string]$st.Desti.Text)) } }
        foreach ($k in @($st.Ctl.Keys)) {
            $c = $st.Ctl[$k]
            $vk = @($c.Via.Tag)
            $via = if ($c.Via.SelectedIndex -ge 0) { [string]$vk[$c.Via.SelectedIndex] } else { [string]$vk[0] }
            $dest = @(@($c.Dest.Keys) | Where-Object { $c.Dest[$_].Checked })
            $nou.eines[$k] = _CorreuEinaDeLaPantalla $via $dest ([string]$c.Fixes.Text) ([string]$c.Cco.Text) ([string]$c.CcoAbans) ([string]$c.CcoTria.Text) ([string]$c.CcoTriaAbans)
        }
        try {
            Save-CorreuEines $nou; $st.Cfg = $nou
            if ($rem -ne [string]$st.RemAbans) {
                if (-not (Set-CorreuRemitent $rem)) { throw "No s'ha pogut desar el remitent a settings.json." }
                $st.RemAbans = $rem
            }
            if ($null -ne $st.MobAbans) {
                $ara = & $st.Mob.Llegeix
                $r = Save-CorreuMobilConfig $st.MobAbans $ara
                $st.MobAbans = $ara
                if ($r.Dest) { $st.MobDest = $true }
                if ($r.Text) { $st.MobText = [string]$r.Text; $st.MobOk = [bool]$r.Ok }
            }
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n" + $_.Exception.Message, $st.Titol, 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    $eds = $Script:CorreuTextosEditors
    foreach ($k in @($ctl.Keys)) {
        $clau = [string]$k
        $ctl[$k].BtnText.add_Click({
            if (-not $eds.Contains($clau)) { return }
            try { & $eds[$clau].Obre } catch {
                [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Text del correu', 'OK', 'Error') | Out-Null
            }
        }.GetNewClosure())
        $ctl[$k].Btn.add_Click({
            # La prova surt amb el que hi ha a la pantalla: primer es desa.
            if (-not (& $desa)) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try { $r = Send-CorreuProva $clau ([string]$st.Gia.Text) ([string]$st.Desti.Text) } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
            [System.Windows.Forms.MessageBox]::Show([string]$r.Text, 'Correu de prova', 'OK', $(if ($r.Ok) { 'Information' } else { 'Warning' })) | Out-Null
        }.GetNewClosure())
    }
    # El text de l'app del mobil es el d'"Enviar correu" (el mateix editor).
    $mob.BtnText.add_Click({
        try { & $eds['mobil'].Obre } catch {
            [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Text del correu', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $peu = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) @(@{ Nom = 'Desar'; Text = 'Desar'; Estil = 'primari' }) ($y + 8)
    $peu.Desar.add_Click({
        if (-not (& $desa)) { return }
        $txt = "Desat en aquest ordinador. Ja ho fan servir totes les eines."
        if ($st.MobText) { $txt += "`n`n" + [string]$st.MobText }
        # A QUI VA el correu del mobil arriba a l'app amb Actualitzar (es a
        # email-textos.json, que es publica): s'ofereix fer-ho ara.
        if ($st.MobDest -and $null -ne $publica) {
            $rr = [System.Windows.Forms.MessageBox]::Show(($txt + "`n`nA qui va el correu del m" + [char]0x00F2 + "bil ha canviat i hi arriba amb Actualitzar. Vols actualitzar ara?"), $st.Titol, 'YesNo', 'Question')
            $form.Close()
            if ($rr -eq [System.Windows.Forms.DialogResult]::Yes) { & $publica }
            return
        }
        if ($st.MobDest) { $txt += "`n`nA qui va el correu del m" + [char]0x00F2 + "bil hi arribar" + [char]0x00E0 + " la propera vegada que facis Actualitzar." }
        [System.Windows.Forms.MessageBox]::Show($txt, $st.Titol, 'OK', $(if ($st.MobOk) { 'Information' } else { 'Warning' })) | Out-Null
        $form.Close()
    }.GetNewClosure())
    [void](_AddBrandHeader $form $titol ("Des d'on surten, per on, a qui van, la CCO i el text de cada correu " + [char]0x00B7 + ' nom' + [char]0x00E9 + 's en aquest ordinador') 56)
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
