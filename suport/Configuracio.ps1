#requires -Version 5.1
<#
.SYNOPSIS
  Pantalla "Configuracio": rutes d'aquest PC + actualitzar el programa.

.DESCRIPTION
  Finestra WinForms (boto "Configuracio" del menu principal) on l'usuari pot
  veure i sobreescriure, NOMES per a aquest ordinador, les carpetes que el
  programa fa servir (informes, Excel d'activitats, sortida d'informes,
  sortida de rutes, Drive d'escriptori). Els canvis es desen amb
  Save-AppSettings (Settings.ps1) a %LOCALAPPDATA%\InformesCornella\settings.json
  -- mai a suport/config.ps1 (que es comparteix via git). Aixo permet fer
  servir el mateix clone a diversos ordinadors (p.ex. feina i casa) sense que
  la configuracio d'un trepitgi la de l'altre.

  Tambe inclou una seccio "Manteniment" amb un boto per llancar Actualitzar.bat
  (mateixa logica que el .bat de sempre, nomes que es pot obrir des de dins del
  programa; segueix funcionant igual fent-hi doble clic per fora).

  Nomes defineix funcions (cap execucio en carregar-se): segur en mode headless.
#>

# ($Script:ConfigUiAccent, el granat d'aquesta pantalla, era el CINQUE literal
# de la paleta i no el llegia ningu. Fora: si algun dia cal, es $Script:BrandMaroon.)

# _AddConfigRow (fila per triar una carpeta) viu ara a UiComuns.ps1: el fa
# servir tambe PdfSignar.ps1, i un modul no ha de dependre d'aquesta pantalla
# per dibuixar un selector de carpeta.

# ACTUALITZAR EL PROGRAMA: llanca Actualitzar.bat i tanca el programa. Un sol
# lloc per als dos botons que ho fan -el de Configuracio i el de la banda del
# menu principal (octubre 2026: l'usuari el fa servir molt i era dos clics
# endins)-, perque el missatge i la manera de tancar no es puguin separar.
function Invoke-ActualitzarPrograma {
    $batPath = Join-Path $RepoRoot 'Actualitzar.bat'
    if (-not (Test-Path -LiteralPath $batPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat Actualitzar.bat a:`n$batPath", 'Actualitzar el programa', 'OK', 'Error') | Out-Null
        return
    }
    $rr = [System.Windows.Forms.MessageBox]::Show(
        "S'obrira una finestra per actualitzar el programa des de GitHub. El programa es tancara mentre s'actualitza i, en acabar (quan premis una tecla), es tornara a obrir ja actualitzat.`n`nVols continuar?",
        'Actualitzar el programa', 'YesNo', 'Question')
    if ($rr -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        Start-Process -FilePath $batPath -WorkingDirectory $RepoRoot
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir Actualitzar.bat:`n$($_.Exception.Message)", 'Actualitzar el programa', 'OK', 'Error') | Out-Null
        return
    }
    # Tanquem el programa DES DE DINS del gestor de clic. Fer servir 'exit'
    # aqui llenca una excepcio de PowerShell que la bomba de missatges de
    # WinForms mostra com a "Excepcio no controlada en un component"; en
    # canvi [Environment]::Exit acaba el proces directament, sense l'error.
    [System.Environment]::Exit(0)
}
function Invoke-ConfiguracioScreen {
    # Llegim els overrides ACTUALS d'aquest PC (poden haver canviat des de
    # l'arrencada si l'usuari torna a obrir aquesta pantalla) i en derivem el
    # valor EFECTIU de cada camp per preomplir els textbox.
    $current = Load-AppSettings
    $effInformesDir    = _ResolveEffectiveValue $current.InformesDir    $Script:DefaultInformesDir
    $effActivitatsDir  = _ResolveEffectiveValue $current.ActivitatsDir  $Script:DefaultActivitatsDir
    $effOutputDir      = _ResolveEffectiveValue $current.OutputDir      $Script:DefaultOutputDir
    $effRutesOutputDir = _ResolveEffectiveValue $current.RutesOutputDir $Script:DefaultRutesOutputDir
    $effDriveBaseDir   = _ResolveEffectiveValue $current.DriveBaseDir   $Script:DefaultDriveBaseDir
    $effCopiaInformesDir = _ResolveEffectiveValue $current.CopiaInformesDir $Script:DefaultCopiaInformesDir

    $form = _NewForm
    $form.Text = 'Configuracio'
    # EN HORITZONTAL, com el menu principal (octubre 2026): dues columnes de
    # 514. A l'esquerra les carpetes; a la dreta els automatismes (una fila per
    # cada un) i el manteniment. Les dues columnes acaben a la mateixa alcada:
    # el grup d'automatismes s'estira fins a la de l'esquerra si en te prou.
    $xEsq = 14; $xDre = 542; $amplCol = 514; $yTop = 66
    # Les carpetes acaben a $fiCarp; a sota, el grup dels correus.
    $fiCarp = 566
    $fiEsq = $fiCarp + 8 + 100
    $nAuto = @($Script:ProgramacionsAuto.Keys).Count
    $altAuto = 72 + 30 * $nAuto
    $altGrpAuto = [math]::Max(($altAuto - 8), ($fiEsq - $yTop - 104))
    $fiCols = [math]::Max($fiEsq, ($yTop + $altGrpAuto + 104))
    $form.ClientSize = New-Object System.Drawing.Size(($xDre + $amplCol + 14), ($fiCols + 12 + 52))
    $form.MinimumSize = New-Object System.Drawing.Size(640, 480)
    $form.StartPosition = 'CenterScreen'

    # ---- Carpetes principals -------------------------------------------
    $y = 12
    $grpPrincipals = New-Object System.Windows.Forms.GroupBox
    $grpPrincipals.Text = 'Carpetes principals'
    $grpPrincipals.Location = New-Object System.Drawing.Point($xEsq, $yTop)
    $grpPrincipals.Size = New-Object System.Drawing.Size($amplCol, 172)
    $grpPrincipals.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

    $r = _AddConfigRow $grpPrincipals 24 "Carpeta on hi ha els informes ja generats" $effInformesDir
    $tbInformes = $r.TextBox
    $r = _AddConfigRow $grpPrincipals $r.NextY "Carpeta de l'Excel d'activitats" $effActivitatsDir
    $tbActivitats = $r.TextBox

    # ---- Carpetes addicionals ------------------------------------------
    $grpAddicionals = New-Object System.Windows.Forms.GroupBox
    $grpAddicionals.Text = 'Carpetes addicionals'
    $grpAddicionals.Location = New-Object System.Drawing.Point($xEsq, ($yTop + 180))
    $grpAddicionals.Size = New-Object System.Drawing.Size($amplCol, ($fiCarp - $yTop - 180))
    $grpAddicionals.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

    $r = _AddConfigRow $grpAddicionals 24 "Carpeta on desar els informes que generis" $effOutputDir
    $tbOutput = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta on desar els mapes de ruta" $effRutesOutputDir
    $tbRutes = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta del Drive d'escriptori (per al mobil)" $effDriveBaseDir
    $tbDrive = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta on copiar els informes (copia de seguretat)" $effCopiaInformesDir
    $tbCopia = $r.TextBox

    # ---- Correus d'aquest PC (octubre 2026, CorreuVia.ps1) ------------
    # L'usuari: "poder triar entre EmailJS per no desfer-ho i Outlook per fer
    # proves". La mateixa preferencia que el desplegable d'"Enviar correu".
    # Des de l'octubre de 2026 la via es tria PER EINA (CorreuEines.ps1): aqui
    # hi ha el boto que obre aquella finestra. La via de tot el PC
    # ('CorreuVia') es queda com a valor per defecte de les eines encara no
    # configurades, i el "Desar" d'aqui la conserva tal com era.
    $effCorreuVia = _CorreuViaValida (_PropInf $current 'CorreuVia')
    $grpCorreu = New-Object System.Windows.Forms.GroupBox
    $grpCorreu.Text = "Correus que s'envien des d'aquest PC"
    $grpCorreu.Location = New-Object System.Drawing.Point($xEsq, ($fiCarp + 8))
    $grpCorreu.Size = New-Object System.Drawing.Size($amplCol, 100)
    $grpCorreu.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $btnCorreus = New-Object System.Windows.Forms.Button
    $btnCorreus.Text = 'Correus de cada eina...'
    $btnCorreus.Location = New-Object System.Drawing.Point(12, 20)
    $btnCorreus.Size = New-Object System.Drawing.Size(250, 30)
    _StyleSecondaryButton $btnCorreus
    $btnCorreus.add_Click({ Show-CorreuEinesConfig })
    [void]$grpCorreu.Controls.Add($btnCorreus)
    $lblCorreuNota = New-Object System.Windows.Forms.Label
    $lblCorreuNota.Text = ("Per on surt el correu de cada eina, a qui va, la CCO i el correu de prova.")
    $lblCorreuNota.Location = New-Object System.Drawing.Point(270, 20)
    $lblCorreuNota.Size = New-Object System.Drawing.Size(236, 32)
    $lblCorreuNota.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    [void]$grpCorreu.Controls.Add($lblCorreuNota)
    # EL REMITENT (l'usuari: "a la feina puc enviar des d'adreces diferents;
    # com ho puc seleccionar per no haver de canviar-la a tots els correus?").
    # Editable: una bustia compartida no surt als comptes i s'escriu a ma.
    $lblRem = New-Object System.Windows.Forms.Label
    $lblRem.Text = 'Des de (Outlook):'
    $lblRem.Location = New-Object System.Drawing.Point(12, 63)
    $lblRem.Size = New-Object System.Drawing.Size(104, 20)
    [void]$grpCorreu.Controls.Add($lblRem)
    $cbRem = New-Object System.Windows.Forms.ComboBox
    $cbRem.DropDownStyle = 'DropDown'
    $cbRem.Location = New-Object System.Drawing.Point(118, 60)
    $cbRem.Size = New-Object System.Drawing.Size(238, 24)
    $cbRem.Text = ([string](_PropInf $current 'CorreuRemitent')).Trim()
    [void]$grpCorreu.Controls.Add($cbRem)
    $btnComptes = New-Object System.Windows.Forms.Button
    $btnComptes.Text = "Comptes de l'Outlook"
    $btnComptes.Location = New-Object System.Drawing.Point(362, 58)
    $btnComptes.Size = New-Object System.Drawing.Size(140, 28)
    _StyleSecondaryButton $btnComptes
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
    [void]$grpCorreu.Controls.Add($btnComptes)
    $ttRem = New-Object System.Windows.Forms.ToolTip
    $ttRem.SetToolTip($cbRem, ("Buit = el compte per defecte de l'Outlook. Una b" + [char]0x00FA + "stia compartida s'escriu a m" + [char]0x00E0 + " (cal tenir-hi perm" + [char]0x00ED + "s). Amb EmailJS no compta."))

    # ---- Automatismes (octubre 2026) ----------------------------------
    # L'usuari: "aquests automatismes, com son ja uns quants, haurien de ser
    # configurables des de la configuracio". Una fila per cada automatisme del
    # registre (ModeAutomatic.ps1): engegat o no (el mateix que l'interruptor
    # A/M del menu), cada dia o un dia de la setmana, i l'hora. Es desa a
    # settings.json (nomes el que difereix del per defecte) i s'aplica en viu.
    $grpAuto = New-Object System.Windows.Forms.GroupBox
    $grpAuto.Text = 'Automatismes'
    $grpAuto.Location = New-Object System.Drawing.Point($xDre, $yTop)
    $grpAuto.Size = New-Object System.Drawing.Size($amplCol, $altGrpAuto)
    $grpAuto.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $modesAuto = $Script:ModesAuto
    $progDefs = @{}
    $autoCtl = [ordered]@{}
    $dies = @('dilluns', 'dimarts', 'dimecres', 'dijous', 'divendres', 'dissabte', 'diumenge')
    $ya = 22
    foreach ($k in @($Script:ProgramacionsAuto.Keys)) {
        $def = $Script:ProgramacionsAuto[$k]
        $progDefs[$k] = $def
        $pr = Get-ProgramacioAuto $k $current
        $c = @{ Clau = $k }
        if ($modesAuto.Contains($k)) {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = [string]$def.Titol
            $chk.Location = New-Object System.Drawing.Point(12, $ya)
            $chk.Size = New-Object System.Drawing.Size(190, 24)
            $chk.Checked = [bool](& $modesAuto[$k].Actiu)
            $c.Chk = $chk; $c.Abans = $chk.Checked
            [void]$grpAuto.Controls.Add($chk)
        } else {
            # Els recordatoris s'engeguen a la seva eina (la tasca del Windows):
            # aqui nomes quan.
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = [string]$def.Titol
            $lbl.Location = New-Object System.Drawing.Point(30, ($ya + 4))
            $lbl.Size = New-Object System.Drawing.Size(172, 20)
            [void]$grpAuto.Controls.Add($lbl)
        }
        $cbF = New-Object System.Windows.Forms.ComboBox
        $cbF.DropDownStyle = 'DropDownList'
        [void]$cbF.Items.AddRange(@('cada dia', 'cada setmana'))
        $cbF.SelectedIndex = if ([string]$pr.Freq -eq 'setmana') { 1 } else { 0 }
        $cbF.Location = New-Object System.Drawing.Point(206, $ya); $cbF.Size = New-Object System.Drawing.Size(104, 24)
        $cbD = New-Object System.Windows.Forms.ComboBox
        $cbD.DropDownStyle = 'DropDownList'
        [void]$cbD.Items.AddRange($dies)
        $cbD.SelectedIndex = [math]::Max(0, [math]::Min(6, [int]$pr.Dia - 1))
        $cbD.Location = New-Object System.Drawing.Point(316, $ya); $cbD.Size = New-Object System.Drawing.Size(98, 24)
        $cbD.Enabled = ($cbF.SelectedIndex -eq 1)
        $cbF.add_SelectedIndexChanged({ $cbD.Enabled = ($cbF.SelectedIndex -eq 1) }.GetNewClosure())
        $dt = New-Object System.Windows.Forms.DateTimePicker
        $dt.Format = 'Custom'; $dt.CustomFormat = 'HH:mm'; $dt.ShowUpDown = $true
        $hm = ([string]$pr.Hora).Split(':')
        $dt.Value = (Get-Date).Date.AddHours([int]$hm[0]).AddMinutes([int]$hm[1])
        $dt.Location = New-Object System.Drawing.Point(420, $ya); $dt.Size = New-Object System.Drawing.Size(80, 24)
        $c.Freq = $cbF; $c.Dia = $cbD; $c.Hora = $dt
        [void]$grpAuto.Controls.Add($cbF); [void]$grpAuto.Controls.Add($cbD); [void]$grpAuto.Controls.Add($dt)
        $autoCtl[$k] = $c
        $ya += 30
    }
    $lblAutoNota = New-Object System.Windows.Forms.Label
    $lblAutoNota.Text = ("Si a l'hora que toca el programa (o el PC) estava tancat, es fa en obrir-lo. Els recordatoris s'engeguen a la seva eina.")
    $lblAutoNota.Location = New-Object System.Drawing.Point(12, ($ya + 6))
    $lblAutoNota.Size = New-Object System.Drawing.Size(488, 36)
    $lblAutoNota.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    [void]$grpAuto.Controls.Add($lblAutoNota)

    # ---- Manteniment: info de versio + actualitzar ----------------------
    $grpMant = New-Object System.Windows.Forms.GroupBox
    $grpMant.Text = 'Manteniment'
    $grpMant.Location = New-Object System.Drawing.Point($xDre, ($yTop + $altGrpAuto + 8))
    $grpMant.Size = New-Object System.Drawing.Size($amplCol, 96)
    $grpMant.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

    $branch = ''
    $commit = ''
    try {
        $branch = ((& git -C $RepoRoot branch --show-current 2>$null) | Select-Object -First 1)
        $commit = ((& git -C $RepoRoot log -1 --format='%h  %s' 2>$null) | Select-Object -First 1)
    } catch { }
    $verText = if ([string]::IsNullOrWhiteSpace($branch) -and [string]::IsNullOrWhiteSpace($commit)) {
        "Informacio del programa no disponible (git no trobat)."
    } else {
        "Branca: $branch" + $(if ($commit) { "   .   Ultim commit: $commit" } else { '' })
    }
    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text = $verText
    $lblVer.Location = New-Object System.Drawing.Point(14, 26)
    $lblVer.Size = New-Object System.Drawing.Size(486, 20)
    $lblVer.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblVer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    [void]$grpMant.Controls.Add($lblVer)

    $btnActualitzar = New-Object System.Windows.Forms.Button
    $btnActualitzar.Location = New-Object System.Drawing.Point(14, 52)
    $btnActualitzar.Size = New-Object System.Drawing.Size(260, 32)
    _StyleSecondaryButton $btnActualitzar
    _PosaIcona $btnActualitzar ([string][char]0x21BB) 'Actualitzar el programa'
    $btnActualitzar.add_Click({ Invoke-ActualitzarPrograma })
    [void]$grpMant.Controls.Add($btnActualitzar)

    # ---- Barra inferior: Tancar / Restaura ... Desar ---------------------
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'
    $botPanel.Height = 52

    $peu = _AddPeuBotons $form @(
        @{ Nom = 'Tancar'; Text = 'Tancar' },
        @{ Nom = 'Restaura'; Text = 'Restaura els valors per defecte' }) @(
        @{ Nom = 'Desar'; Text = 'Desar'; Estil = 'primari' }) 10 $botPanel
    $btnTancar = $peu.Tancar; $btnRestaura = $peu.Restaura; $btnDesar = $peu.Desar

    # ELS VALORS PER DEFECTE, CAPTURATS AQUI. Dins d'una closure (.GetNewClosure)
    # $Script: no es el d'aquest script sino el d'un modul nou: els
    # $Script:Default* hi valien buit. "Restaura" deixava totes les caselles en
    # blanc i "Desar" comparava amb uns valors per defecte buits, o sigui que
    # desava TOTES les rutes com si fossin d'aquest PC (vegeu CLAUDE.md).
    $defs = @{
        InformesDir      = $Script:DefaultInformesDir
        ActivitatsDir    = $Script:DefaultActivitatsDir
        OutputDir        = $Script:DefaultOutputDir
        RutesOutputDir   = $Script:DefaultRutesOutputDir
        DriveBaseDir     = $Script:DefaultDriveBaseDir
        CopiaInformesDir = $Script:DefaultCopiaInformesDir
    }

    # Els valors amb que s'ha obert la pantalla: si no se'n toca cap, no cal reiniciar.
    $inicials = @{
        InformesDir = $effInformesDir; ActivitatsDir = $effActivitatsDir; OutputDir = $effOutputDir
        RutesOutputDir = $effRutesOutputDir; DriveBaseDir = $effDriveBaseDir; CopiaInformesDir = $effCopiaInformesDir
    }

    $btnRestaura.add_Click({
        $tbInformes.Text   = $defs.InformesDir
        $tbActivitats.Text = $defs.ActivitatsDir
        $tbOutput.Text     = $defs.OutputDir
        $tbRutes.Text      = $defs.RutesOutputDir
        $tbDrive.Text      = $defs.DriveBaseDir
        $tbCopia.Text      = $defs.CopiaInformesDir
        # La programacio per defecte de cada automatisme (l'interruptor no es toca).
        foreach ($k in @($autoCtl.Keys)) {
            $c = $autoCtl[$k]; $d = $progDefs[$k]
            $c.Freq.SelectedIndex = if ([string]$d.Freq -eq 'setmana') { 1 } else { 0 }
            $c.Dia.SelectedIndex = [math]::Max(0, [int]$d.Dia - 1)
            $hm = ([string]$d.Hora).Split(':')
            $c.Hora.Value = (Get-Date).Date.AddHours([int]$hm[0]).AddMinutes([int]$hm[1])
        }
        $cbRem.Text = ''
    }.GetNewClosure())

    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())

    $btnDesar.add_Click({
        $remitent = $cbRem.Text.Trim()
        if (-not (_CorreuRemitentValid $remitent)) {
            [System.Windows.Forms.MessageBox]::Show(("'" + $remitent + "' no " + [char]0x00E9 + "s una adre" + [char]0x00E7 + "a de correu. Deixa-ho en blanc per fer servir el compte per defecte de l'Outlook."), 'Configuracio', 'OK', 'Warning') | Out-Null
            return
        }
        $values = @{
            InformesDir      = $tbInformes.Text.Trim()
            ActivitatsDir    = $tbActivitats.Text.Trim()
            OutputDir        = $tbOutput.Text.Trim()
            RutesOutputDir   = $tbRutes.Text.Trim()
            DriveBaseDir     = $tbDrive.Text.Trim()
            CopiaInformesDir = $tbCopia.Text.Trim()
        }
        $overrides = _BuildSettingsOverrides $values $defs
        # Els automatismes: la programacio (nomes la que difereix) i els
        # interruptors (el mateix que el menu; si no es pot engegar, ho diu).
        $progs = @{}
        foreach ($k in @($autoCtl.Keys)) {
            $c = $autoCtl[$k]
            $progs[$k] = @{ Freq = $(if ($c.Freq.SelectedIndex -eq 1) { 'setmana' } else { 'dia' }); Dia = ($c.Dia.SelectedIndex + 1); Hora = $c.Hora.Value.ToString('HH:mm') }
        }
        $autoSet = ConvertTo-AutomatismesSettings $progs
        if ($autoSet.Count -gt 0) { $overrides['Automatismes'] = $autoSet }
        # Aquest "Desar" reescriu settings.json sencer: la via dels correus hi
        # ha de ser, si no es perdria (nomes si no es la per defecte).
        $overrides = _SettingsAmbCorreuVia $overrides $effCorreuVia
        $overrides = _SettingsAmbClau $overrides 'CorreuRemitent' $remitent
        if (-not (Save-AppSettings $overrides)) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar la configuracio.", 'Configuracio', 'OK', 'Error') | Out-Null
            return
        }
        foreach ($k in @($autoCtl.Keys)) {
            $c = $autoCtl[$k]
            if ($null -eq $c.Chk -or $c.Chk.Checked -eq $c.Abans) { continue }
            if ($c.Chk.Checked) {
                $req = [string](& $modesAuto[$k].Requisit)
                if ($req -ne '') { [System.Windows.Forms.MessageBox]::Show($req, [string]$modesAuto[$k].Titol, 'OK', 'Warning') | Out-Null; continue }
            }
            [void](& $modesAuto[$k].DesaActiu $c.Chk.Checked)
        }
        try { Update-RecordatorisTascaSiCal -Forca } catch { }
        # Les carpetes demanen reiniciar; els automatismes no (es llegeixen en viu).
        $carpetesIguals = $true
        foreach ($k in @($values.Keys)) { if ([string]$values[$k] -ne [string]$inicials[$k]) { $carpetesIguals = $false } }
        if ($carpetesIguals) {
            [System.Windows.Forms.MessageBox]::Show("Configuracio desada. Els automatismes ja fan servir la programacio nova.", 'Configuracio', 'OK', 'Information') | Out-Null
            $form.Close()
            return
        }
        $rr = [System.Windows.Forms.MessageBox]::Show(
            "Configuracio desada.`n`nCal reiniciar el programa perque els canvis s'apliquin a totes les pantalles. Vols reiniciar ara?",
            'Configuracio', 'YesNo', 'Information')
        if ($rr -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + (Join-Path $ScriptRoot 'GenerarInforme.vbs') + '"')
            } catch { }
            # Mateix motiu que a "Actualitzar el programa": acabem el proces amb
            # [Environment]::Exit i no amb 'exit' (que petaria dins del gestor).
            [System.Environment]::Exit(0)
        } else {
            $form.Close()
        }
    }.GetNewClosure())

    [void]$form.Controls.Add($grpPrincipals)
    [void]$form.Controls.Add($grpAddicionals)
    [void]$form.Controls.Add($grpCorreu)
    [void]$form.Controls.Add($grpAuto)
    [void]$form.Controls.Add($grpMant)
    [void]$form.Controls.Add($botPanel)
    [void](_AddBrandHeader $form ('Configuraci' + [char]0x00F3) ("Nom" + [char]0x00E9 + "s afecta aquest ordinador: no es comparteix ni es puja a GitHub.") 56)

    [void]$form.ShowDialog()
}
