#requires -Version 5.1
<#
  ContactesPantalla.ps1 - LA FINESTRA "Contactes" (des d'"Editar base"): la
  base de contactes per revisar-la i corregir-la a ma.

  L'usuari (octubre 2026): una llista per GIA amb filtre per tipus d'avis; el
  detall de l'activitat amb l'Excel i els documents de costat i un boto per
  obrir el document font; els botons "Es el tecnic", "Es el representant
  legal", "Descarta l'avis" i "Edita"; i EXPORTAR la llista de correccions per
  entrar-les al GIA (l'Excel no es toca mai).

  Les correccions es desen a contactes-db.json (Set-ContactesCorreccio, amb el
  mutex de la base) i MANEN sobre el que torni a sortir d'"Actualitzar base";
  "Desfer canvi a ma" les treu. Les files de les graelles i el que es desa en
  clicar cada boto es decideix en funcions PURES (_CtFiles*, _CtCorreccioDe*),
  que es proven a Linux; la finestra nomes les pinta.

  S'apunta a "Editar base" pel registre $Script:EditarBaseBotonsExtra
  (Informes.ps1): el client es aquest fitxer, no l'editor.
#>

# Les opcions del filtre: el text i quines activitats hi entren. PURES.
$Script:CtFiltres = [ordered]@{
    'totes'         = 'Totes les activitats'
    'avisos'        = 'Amb avisos'
    'es_el_tecnic'  = ("Dades del t" + [char]0x00E8 + "cnic a l'Excel")
    'error'         = 'Errors de tecleig o de columna'
    'falta'         = "Falten a l'Excel"
    'diferent'      = 'Diferents dels documents'
    'canvi_titular' = 'Canvi de titular'
    'ama'           = ("Corregides a m" + [char]0x00E0)
}

function _CtPassaFiltre($act, [string]$filtre) {
    $av = @(_CtV $act 'avisos')
    switch ($filtre) {
        'totes'  { return $true }
        'avisos' { return ($av.Count -gt 0) }
        'ama'    { return [bool](_CtV $act 'editat_a_ma') }
        default  { foreach ($v in $av) { if ([string](_CtV $v 'tipus') -eq $filtre) { return $true } } return $false }
    }
}

# Les files de la llista: @{ Gia; Titular; Avisos; AMa; Color }. $acts: gia ->
# activitat. PURA.
function _CtFilesLlista($acts, [string]$filtre, [string]$cerca) {
    $out = New-Object System.Collections.ArrayList
    $q = _CtPla ([string]$cerca).Trim()
    foreach ($g in @(@($acts.Keys) | Sort-Object { _GiaNumeric $_ })) {
        $a = $acts[$g]
        if (-not (_CtPassaFiltre $a $filtre)) { continue }
        $tit = [string](_CtV (_CtV $a 'excel') 'TITULAR')
        if ($tit -eq '') { $tit = [string](_CtV (_CtV $a 'titular') 'nom') }
        if ($q -ne '' -and -not ((_CtPla ([string]$g + ' ' + $tit)).Contains($q))) { continue }
        $tipus = @(@(_CtV $a 'avisos') | ForEach-Object { [string](_CtV $_ 'tipus') } | Select-Object -Unique)
        $color = if ($tipus -contains 'es_el_tecnic') { 'tecnic' } elseif ($tipus -contains 'error' -or $tipus -contains 'canvi_titular') { 'error' } elseif ($tipus.Count -gt 0) { 'avis' } else { '' }
        [void]$out.Add(@{
            Gia = [string]$g; Titular = $tit
            Avisos = $(if ($tipus.Count -gt 0) { [string]@(_CtV $a 'avisos').Count + ': ' + ($tipus -join ', ') } else { '' })
            AMa = $(if ([bool](_CtV $a 'editat_a_ma')) { 'si' } else { '' })
            Color = $color
        })
    }
    return $out.ToArray()
}

# Les files del detall: l'Excel i els documents de costat. Cada fila porta el
# que fan servir els botons (Qui, Dada, la persona i la font). PURA.
$Script:CtDetallCamps = @(
    @{ Qui = 'titular'; Etiq = 'Titular'; Camps = @(@{ D = 'nom'; X = 'TITULAR'; N = ('Ra' + [char]0x00F3 + ' social') }, @{ D = 'nif'; X = 'NIF'; N = 'NIF' }, @{ D = 'email'; X = 'EMAIL'; N = 'E-mail' }, @{ D = 'telefon'; X = 'TELEFON'; N = ('Tel' + [char]0x00E8 + 'fon') }, @{ D = 'mobil'; X = 'MOBIL'; N = ('M' + [char]0x00F2 + 'bil') }) }
    @{ Qui = 'representant_legal'; Etiq = 'Representant legal'; Camps = @(@{ D = 'nom'; X = 'REP_NOM'; N = 'Nom' }, @{ D = 'nif'; X = 'REP_NIF'; N = 'NIF' }, @{ D = 'email'; X = 'REP_EMAIL'; N = 'E-mail' }, @{ D = 'telefon'; X = 'REP_TELEFON'; N = ('Tel' + [char]0x00E8 + 'fon') }, @{ D = 'mobil'; X = 'REP_MOBIL'; N = ('M' + [char]0x00F2 + 'bil') }) }
    @{ Qui = 'establiment'; Etiq = 'Establiment'; Camps = @(@{ D = 'nom_comercial'; X = ''; N = 'Nom comercial' }, @{ D = 'telefon'; X = ''; N = ('Tel' + [char]0x00E8 + 'fon') }, @{ D = 'email'; X = ''; N = 'E-mail' }) }
)

function _CtFilesDetall($act) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $act) { return $out.ToArray() }
    $xl = _CtV $act 'excel'
    foreach ($g in $Script:CtDetallCamps) {
        $p = _CtV $act $g.Qui
        foreach ($c in $g.Camps) {
            $vx = if ($c.X -ne '') { [string](_CtV $xl $c.X) } else { '' }
            $vd = [string](_CtV $p $c.D)
            [void]$out.Add(@{
                Qui = $g.Qui; Dada = $c.D; XlClau = $c.X; Grup = $g.Etiq; Camp = $c.N
                Excel = $vx; Documents = $vd
                Confianca = $(if ($vd -ne '') { [string](_CtV $p 'confianca') } else { '' })
                Font = $(if ($vd -ne '') { [string](_CtV $p 'font') } else { '' })
                Persona = $p
            })
        }
    }
    foreach ($o in @(_CtV $act 'persones_autoritzades')) {
        if ($null -eq $o) { continue }
        $txt = (@([string](_CtV $o 'nom'), [string](_CtV $o 'email'), [string](_CtV $o 'mobil'), [string](_CtV $o 'telefon')) | Where-Object { $_ -ne '' }) -join ' ' + [char]0x00B7 + ' '
        [void]$out.Add(@{
            Qui = 'autoritzat'; Dada = ''; XlClau = ''; Grup = ('Autoritzat (' + [string](_CtV $o 'rol') + ')'); Camp = [string](_CtV $o 'motiu')
            Excel = ''; Documents = $txt; Confianca = [string](_CtV $o 'confianca'); Font = [string](_CtV $o 'font'); Persona = $o
        })
    }
    return $out.ToArray()
}

# La persona d'una fila del detall, per marcar-la: la dels documents o, si no
# n'hi ha, la que diu l'Excel (el representant legal de l'Excel que es el
# tecnic). PURA.
function _CtPersonaDeFila($fila, $act) {
    if ($null -eq $fila) { return $null }
    if ($fila.Qui -eq 'autoritzat') { return (_CtPersona ([string](_CtV $fila.Persona 'nom')) ([string](_CtV $fila.Persona 'nif')) ([string](_CtV $fila.Persona 'email')) ([string](_CtV $fila.Persona 'telefon')) ([string](_CtV $fila.Persona 'mobil'))) }
    if ($fila.Qui -eq 'establiment') { return $null }
    $xl = _CtV $act 'excel'
    if ($fila.Qui -eq 'representant_legal') {
        $p = $fila.Persona
        if ($null -ne $p -and -not (_CtPersonaBuida $p)) { return (_CtPersona ([string](_CtV $p 'nom')) ([string](_CtV $p 'nif')) ([string](_CtV $p 'email')) ([string](_CtV $p 'telefon')) ([string](_CtV $p 'mobil'))) }
        $x = _CtPersona ([string](_CtV $xl 'REP_NOM')) ([string](_CtV $xl 'REP_NIF')) ([string](_CtV $xl 'REP_EMAIL')) ([string](_CtV $xl 'REP_TELEFON')) ([string](_CtV $xl 'REP_MOBIL'))
        if (_CtPersonaBuida $x) { return $null }
        return $x
    }
    # Del titular, nomes un correu o un telefon (el del tecnic posat al titular).
    if (@('email', 'telefon', 'mobil') -notcontains $fila.Dada) { return $null }
    $v = if ([string]$fila.Excel -ne '') { [string]$fila.Excel } else { [string]$fila.Documents }
    if ($v -eq '') { return $null }
    if ($fila.Dada -eq 'email') { return (_CtPersona '' '' $v) }
    return (_CtPersona '' '' '' $v)
}

# Les correccions que desen els botons. PURES.
function _CtCorreccioPersona([string]$tipus, $persona) {
    if ($null -eq $persona) { return $null }
    $cl = @(_CtClausPersona $persona)
    if ($cl.Count -eq 0) { return $null }
    return [ordered]@{ tipus = $tipus; claus = $cl; persona = $persona }
}
function _CtCorreccioDescarta($avis) {
    if ($null -eq $avis) { return $null }
    return [ordered]@{ tipus = 'descarta'; avis = [string](_CtV $avis 'id'); text = ([string](_CtV $avis 'tipus') + ': ' + [string](_CtV $avis 'camp') + ' ' + [string](_CtV $avis 'valor_excel')) }
}
function _CtCorreccioEdita($fila, [string]$valor) {
    if ($null -eq $fila -or $fila.Qui -eq 'autoritzat' -or [string]$fila.Dada -eq '') { return $null }
    return [ordered]@{ tipus = 'edita'; camp = ($fila.Qui + '.' + $fila.Dada); valor = $valor.Trim() }
}

# Com es llegeix una correccio a la graella. PURA.
function _CtTextCorreccio($c, $act) {
    $t = [string](_CtV $c 'tipus')
    $nom = [string](_CtV (_CtV $c 'persona') 'nom'); if ($nom -eq '') { $nom = [string](_CtV (_CtV $c 'persona') 'email') }
    if ($nom -eq '') { $nom = [string](_CtV (_CtV $c 'persona') 'mobil') }
    $auto = ''
    foreach ($a in @(_CtV $act 'correccions')) {
        if ([string](_CtV $a 'tipus') -ne $t) { continue }
        if ($t -eq 'edita' -and [string](_CtV $a 'camp') -ne [string](_CtV $c 'camp')) { continue }
        $auto = [string](_CtV $a 'auto'); break
    }
    switch ($t) {
        'es_tecnic'    { return @{ Que = ("$nom $([char]0x00E9)s t$([char]0x00E8)cnic"); Auto = $auto } }
        'es_rep_legal' { return @{ Que = ("$nom $([char]0x00E9)s el representant legal"); Auto = $auto } }
        'descarta'     { return @{ Que = ("Avis descartat: " + [string](_CtV $c 'text')); Auto = '' } }
        'edita'        { return @{ Que = ([string](_CtV $c 'camp') + ' = ' + [string](_CtV $c 'valor')); Auto = $auto } }
    }
    return @{ Que = $t; Auto = '' }
}

# ----------------------------------------------------------------------------
# LA FINESTRA
# ----------------------------------------------------------------------------
function _CtDemanaValor([string]$titol, [string]$etiqueta, [string]$valor) {
    $form = _NewForm
    $form.Text = $titol
    $form.FormBorderStyle = 'FixedDialog'
    $form.ClientSize = New-Object System.Drawing.Size(520, 150)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16, 16)
    $lbl.Size = New-Object System.Drawing.Size(488, 40)
    $lbl.Text = $etiqueta
    $form.Controls.Add($lbl)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(16, 60)
    $tb.Size = New-Object System.Drawing.Size(488, 24)
    $tb.Text = $valor
    $form.Controls.Add($tb)
    [void](_AddPeuBotons $form @(
        @{ Nom = 'No'; Text = 'Cancel' + [char]0x00B7 + 'lar'; Resultat = 'Cancel'; Esc = $true }) @(
        @{ Nom = 'Ok'; Text = 'Desar'; Resultat = 'OK'; Intro = $true; Estil = 'primari' }) 104)
    if ($form.ShowDialog() -ne 'OK') { return $null }
    return [string]$tb.Text
}

function Show-ContactesFinestra {
    $titol = 'Contactes de les activitats'
    $db = $null
    try { $db = Read-ContactesDb } catch { $db = $null }
    if ($null -eq $db) {
        [System.Windows.Forms.MessageBox]::Show("Encara no hi ha cap repas de contactes.`n`nFes 'Actualitzar base': al final repassa les dades de contacte de cada activitat.", $titol, 'OK', 'Information') | Out-Null
        return
    }
    $st = @{ Db = $db; Acts = (_CtMapaDe (_CtV $db 'activitats')); Gia = ''; Filtre = 'avisos'; Cerca = ''; Arrel = [string](_CtV $db 'carpeta_arrel') }

    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(1240, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)
    $ra = _CtResumAvisos $st.Acts
    $st.Loading = $false

    # Barra de dalt: filtre i cerca.
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 40
    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Mostra:'; $lblF.AutoSize = $true; $lblF.Location = New-Object System.Drawing.Point(12, 12)
    $top.Controls.Add($lblF)
    $cmbF = New-Object System.Windows.Forms.ComboBox
    $cmbF.DropDownStyle = 'DropDownList'; $cmbF.Location = New-Object System.Drawing.Point(70, 8); $cmbF.Width = 240
    foreach ($k in $Script:CtFiltres.Keys) { [void]$cmbF.Items.Add($Script:CtFiltres[$k]) }
    $top.Controls.Add($cmbF)
    $clausF = @($Script:CtFiltres.Keys)

    # El cos: la llista a l'esquerra, el detall a la dreta.
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    # El SplitterDistance s'ha de posar quan el SplitContainer ja te la mida
    # de la finestra: abans fa 150 d'ample i un 420 llanca.
    $form.add_Shown({ try { $split.SplitterDistance = 420 } catch { } }.GetNewClosure())
    $gL = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $gL
    $gL.ReadOnly = $true
    foreach ($c in @(@('GIA', 60), @('Titular', 190), @('Avisos', 130), @(('A m' + [char]0x00E0), 40))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.Width = $c[1]; [void]$gL.Columns.Add($col)
    }
    $split.Panel1.Controls.Add($gL)

    $dreta = New-Object System.Windows.Forms.TableLayoutPanel
    $dreta.Dock = 'Fill'; $dreta.ColumnCount = 1; $dreta.RowCount = 6
    foreach ($h in @(22, 55, 22, 30, 22, 15)) {
        $tipus = if ($h -le 22) { [System.Windows.Forms.SizeType]::Absolute } else { [System.Windows.Forms.SizeType]::Percent }
        [void]$dreta.RowStyles.Add((New-Object System.Windows.Forms.RowStyle($tipus, $(if ($h -le 22) { 22 } else { $h }))))
    }
    $nouGrid = {
        param($cols)
        $g = New-Object System.Windows.Forms.DataGridView
        _StyleListGrid $g
        $g.ReadOnly = $true
        foreach ($c in $cols) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.Width = $c[1]; [void]$g.Columns.Add($col) }
        return $g
    }
    $etiq = { param($t) $l = New-Object System.Windows.Forms.Label; $l.Text = $t; $l.Dock = 'Fill'; $l.Font = New-Object System.Drawing.Font($l.Font, [System.Drawing.FontStyle]::Bold); return $l }
    $gD = & $nouGrid @(@('', 120), @('Camp', 90), @("A l'Excel", 170), @('Als documents', 230), @(('Confian' + [char]0x00E7 + 'a'), 70), @('Font', 260))
    $gA = & $nouGrid @(@('Tipus', 90), @('Camp', 110), @("A l'Excel", 150), @('Proposta', 150), @('Problema', 380))
    $gC = & $nouGrid @(@(('Correcci' + [char]0x00F3), 380), @(('Valor autom' + [char]0x00E0 + 'tic'), 300))
    $dreta.Controls.Add((& $etiq "Excel i documents (les dades marcades, d'on surten)"), 0, 0)
    $dreta.Controls.Add($gD, 0, 1)
    $dreta.Controls.Add((& $etiq 'Avisos'), 0, 2)
    $dreta.Controls.Add($gA, 0, 3)
    $dreta.Controls.Add((& $etiq ('Correccions a m' + [char]0x00E0 + ' (manen sobre Actualitzar base)')), 0, 4)
    $dreta.Controls.Add($gC, 0, 5)
    $split.Panel2.Controls.Add($dreta)

    $colors = @{
        tecnic = [System.Drawing.Color]::FromArgb(255, 226, 226)
        error  = [System.Drawing.Color]::FromArgb(255, 238, 214)
        avis   = [System.Drawing.Color]::FromArgb(255, 250, 220)
    }

    $omplDetall = {
        $gD.Rows.Clear(); $gA.Rows.Clear(); $gC.Rows.Clear()
        if ($st.Gia -eq '' -or -not $st.Acts.ContainsKey($st.Gia)) { return }
        $a = $st.Acts[$st.Gia]
        foreach ($f in @(_CtFilesDetall $a)) {
            $i = $gD.Rows.Add(@($f.Grup, $f.Camp, $f.Excel, $f.Documents, $f.Confianca, $f.Font))
            $gD.Rows[$i].Tag = $f
            if ($f.Excel -ne '' -and $f.Documents -ne '' -and -not (_CtIgual $f.Dada $f.Excel $f.Documents)) { $gD.Rows[$i].DefaultCellStyle.BackColor = $colors.avis }
        }
        foreach ($v in @(_CtV $a 'avisos')) {
            if ($null -eq $v) { continue }
            $i = $gA.Rows.Add(@([string](_CtV $v 'tipus'), [string](_CtV $v 'camp'), [string](_CtV $v 'valor_excel'), [string](_CtV $v 'proposta'), [string](_CtV $v 'problema')))
            $gA.Rows[$i].Tag = $v
            $k = if ([string](_CtV $v 'tipus') -eq 'es_el_tecnic') { 'tecnic' } elseif (@('error', 'canvi_titular') -contains [string](_CtV $v 'tipus')) { 'error' } else { 'avis' }
            $gA.Rows[$i].DefaultCellStyle.BackColor = $colors[$k]
        }
        $corrs = _CtCorreccionsDe $st.Db
        $llista = if ($corrs.ContainsKey($st.Gia)) { @($corrs[$st.Gia]) } else { @() }
        for ($k = 0; $k -lt $llista.Count; $k++) {
            $t = _CtTextCorreccio $llista[$k] $a
            $i = $gC.Rows.Add(@($t.Que, $t.Auto))
            $gC.Rows[$i].Tag = $k
        }
        $n = [string](_CtV $a 'notes')
        if ($n -ne '') { $i = $gA.Rows.Add(@('nota', '', '', '', $n)) }
    }.GetNewClosure()

    # En omplir, el Rows.Add ja selecciona la primera fila i salta el
    # SelectionChanged ABANS que la fila porti el Tag: per aixo el 'Loading', i
    # la que estava triada es torna a triar al final.
    $omplLlista = {
        $volgut = $st.Gia
        $st.Loading = $true
        try {
            $gL.Rows.Clear()
            $sel = -1
            foreach ($f in @(_CtFilesLlista $st.Acts $st.Filtre $st.Cerca)) {
                $i = $gL.Rows.Add(@($f.Gia, $f.Titular, $f.Avisos, $f.AMa))
                $gL.Rows[$i].Tag = $f.Gia
                if ($f.Color -ne '') { $gL.Rows[$i].DefaultCellStyle.BackColor = $colors[$f.Color] }
                if ($f.Gia -eq $volgut) { $sel = $i }
            }
            if ($sel -lt 0 -and $gL.Rows.Count -gt 0) { $sel = 0 }
            $gL.ClearSelection()
            if ($sel -ge 0) { $gL.CurrentCell = $gL.Rows[$sel].Cells[0]; $gL.Rows[$sel].Selected = $true; $st.Gia = [string]$gL.Rows[$sel].Tag }
            else { $st.Gia = '' }
        } finally { $st.Loading = $false }
        & $omplDetall
    }.GetNewClosure()

    $gL.add_SelectionChanged({
        if ($st.Loading -or $gL.SelectedRows.Count -eq 0) { return }
        $st.Gia = [string]$gL.SelectedRows[0].Tag
        & $omplDetall
    }.GetNewClosure())
    $cmbF.add_SelectedIndexChanged({ $st.Filtre = $clausF[$cmbF.SelectedIndex]; & $omplLlista }.GetNewClosure())
    $tbCerca = _AddSearchBox $top 330 9 220 'Cerca:' $null
    $tbCerca.add_TextChanged({ $st.Cerca = $tbCerca.Text; & $omplLlista }.GetNewClosure())

    # Desa una correccio (o en treu una) i torna a llegir la base.
    $desa = {
        param($corr, [int]$treu = -1)
        if ($null -eq $corr -and $treu -lt 0) { return }
        $ok = $false
        try { $ok = Set-ContactesCorreccio $st.Gia $corr $treu } catch { $ok = $false }
        if (-not $ok) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar (potser 'Actualitzar base' s'est" + [char]0x00E0 + " fent ara mateix). Torna-hi d'aqu" + [char]0x00ED + " a una estona.", $titol, 'OK', 'Warning') | Out-Null
            return
        }
        $st.Db = Read-ContactesDb
        $st.Acts = _CtMapaDe (_CtV $st.Db 'activitats')
        & $omplLlista
    }.GetNewClosure()
    $filaD = { if ($gD.SelectedRows.Count -eq 0) { return $null } return $gD.SelectedRows[0].Tag }.GetNewClosure()
    $cal = { param($t) [System.Windows.Forms.MessageBox]::Show($t, $titol, 'OK', 'Information') | Out-Null }.GetNewClosure()

    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    [void](_AddPeuBotons $form @(
        @{ Nom = 'Enrere'; Text = (_TxtEnrere); Clic = { $form.Close() }.GetNewClosure() },
        @{ Nom = 'Export'; Text = 'Exportar correccions (Excel)'; Clic = {
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
            $dlg.FileName = 'correccions-contactes_' + (Get-Date).ToString('yyyy-MM-dd') + '.xlsx'
            try { $dlg.InitialDirectory = Split-Path -Parent (Get-ContactesDbPath) } catch { }
            if ($dlg.ShowDialog() -ne 'OK') { return }
            try {
                [void](Export-ContactesCorreccions $dlg.FileName $st.Acts)
                Start-Process -FilePath $dlg.FileName | Out-Null
            } catch { [System.Windows.Forms.MessageBox]::Show("No s'ha pogut exportar:`n" + $_.Exception.Message, $titol, 'OK', 'Error') | Out-Null }
        }.GetNewClosure() },
        @{ Nom = 'Obrir'; Text = 'Obre el document'; Clic = {
            $f = & $filaD
            $font = if ($null -ne $f) { [string]$f.Font } else { '' }
            if ($font -eq '' -and $gA.SelectedRows.Count -gt 0 -and $null -ne $gA.SelectedRows[0].Tag) { $font = [string](_CtV $gA.SelectedRows[0].Tag 'font') }
            if ($font -eq '') { & $cal "Tria una dada o un avis que surti d'un document."; return }
            $ruta = if ([System.IO.Path]::IsPathRooted($font)) { $font } else { Join-Path $st.Arrel $font }
            if (-not (Test-Path -LiteralPath $ruta)) { & $cal ("No s'ha trobat el document:`n" + $ruta); return }
            try { Start-Process -FilePath $ruta | Out-Null } catch { & $cal ("No s'ha pogut obrir:`n" + $_.Exception.Message) }
        }.GetNewClosure() }) @(
        @{ Nom = 'Tecnic'; Text = ([char]0x00C9 + 's el t' + [char]0x00E8 + 'cnic'); Clic = {
            $f = & $filaD
            $p = _CtPersonaDeFila $f $st.Acts[$st.Gia]
            if ($null -eq $p) { & $cal "Tria la persona (el representant legal, una persona autoritzada) o el correu o el tel$([char]0x00E8)fon del titular que s$([char]0x00F3)n del t$([char]0x00E8)cnic."; return }
            & $desa (_CtCorreccioPersona 'es_tecnic' $p)
        }.GetNewClosure() },
        @{ Nom = 'Rep'; Text = ([char]0x00C9 + 's el representant legal'); Clic = {
            $f = & $filaD
            if ($null -eq $f -or @('autoritzat', 'representant_legal') -notcontains $f.Qui) { & $cal 'Tria una persona autoritzada (o el representant legal).'; return }
            & $desa (_CtCorreccioPersona 'es_rep_legal' (_CtPersonaDeFila $f $st.Acts[$st.Gia]))
        }.GetNewClosure() },
        @{ Nom = 'Descarta'; Text = ("Descarta l'av" + [char]0x00ED + 's'); Clic = {
            if ($gA.SelectedRows.Count -eq 0 -or $null -eq $gA.SelectedRows[0].Tag -or $gA.SelectedRows[0].Tag -is [string]) { & $cal 'Tria un avis.'; return }
            & $desa (_CtCorreccioDescarta $gA.SelectedRows[0].Tag)
        }.GetNewClosure() },
        @{ Nom = 'Edita'; Text = 'Edita'; Clic = {
            $f = & $filaD
            if ($null -eq $f -or $f.Qui -eq 'autoritzat') { & $cal "Tria una dada del titular, del representant legal o de l'establiment."; return }
            $v = _CtDemanaValor $titol ($f.Grup + ' ' + [char]0x00B7 + ' ' + $f.Camp + ":`n(el valor bo; manar" + [char]0x00E0 + " sobre el que diguin els documents)") $(if ([string]$f.Documents -ne '') { [string]$f.Documents } else { [string]$f.Excel })
            if ($null -eq $v) { return }
            & $desa (_CtCorreccioEdita $f $v)
        }.GetNewClosure() },
        @{ Nom = 'Desfer'; Text = ('Desfer canvi a m' + [char]0x00E0); Clic = {
            if ($gC.SelectedRows.Count -eq 0) { & $cal ("Tria una correcci" + [char]0x00F3 + " de la llista de baix."); return }
            & $desa $null ([int]$gC.SelectedRows[0].Tag)
        }.GetNewClosure() }) 8 $botPanel)

    # L'ordre compta (Dock): el que s'afegeix l'ultim es col·loca el primer.
    $form.Controls.Add($split)
    $form.Controls.Add($top)
    $form.Controls.Add($botPanel)
    [void](_AddBrandHeader $form $titol ([string]$st.Acts.Count + ' activitats ' + [char]0x00B7 + ' ' + $ra.Total + ' avisos ' + [char]0x00B7 + " l'Excel no es toca: les correccions s'exporten per entrar-les al GIA") 56)
    $cmbF.SelectedIndex = [Math]::Max(0, [array]::IndexOf($clausF, 'avisos'))
    [void]$form.ShowDialog()
}

# S'apunta a "Editar base" (Informes.ps1): el boto i que fa.
[void]$Script:EditarBaseBotonsExtra.Add(@{ Nom = 'Contactes'; Text = 'Contactes...'; Clic = { Show-ContactesFinestra } })
