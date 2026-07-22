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

if (-not $Script:HeadlessTest) {
    $Script:ConfigUiAccent = [System.Drawing.Color]::FromArgb(166, 26, 47)   # granat corporatiu
}

# Fila d'una carpeta configurable: etiqueta + textbox + boto "..." (Explora)
# + indicador d'estat en viu (Test-Path). Afegeix els controls a $parent i
# retorna @{ TextBox = ...; NextY = ... } per encadenar files.
function _AddConfigRow($parent, [int]$y, [string]$labelText, [string]$initialValue) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lbl.Location = New-Object System.Drawing.Point(14, $y)
    $lbl.AutoSize = $true
    [void]$parent.Controls.Add($lbl)
    $y += 20

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(14, $y)
    $tb.Size = New-Object System.Drawing.Size(432, 24)
    $tb.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $tb.Text = $initialValue
    [void]$parent.Controls.Add($tb)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = '...'
    $btn.Location = New-Object System.Drawing.Point(452, ($y - 1))
    $btn.Size = New-Object System.Drawing.Size(36, 24)
    $btn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    [void]$parent.Controls.Add($btn)
    $y += 28

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(14, $y)
    $status.AutoSize = $true
    $status.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
    [void]$parent.Controls.Add($status)
    $y += 22

    $refreshStatus = {
        $p = $tb.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($p)) {
            $status.Text = "(buit: es fara servir el valor per defecte)"
            $status.ForeColor = [System.Drawing.Color]::Gray
            return
        }
        $ok = $false
        try { $ok = Test-Path -LiteralPath $p } catch { $ok = $false }
        if ($ok) {
            $status.Text = "$([char]0x2713) Trobada"
            $status.ForeColor = [System.Drawing.Color]::SeaGreen
        } else {
            $status.Text = "$([char]0x26A0) No trobada ara (es pot desar igualment)"
            $status.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    }.GetNewClosure()
    $tb.add_TextChanged($refreshStatus)
    & $refreshStatus

    $btn.add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $labelText
        try {
            if (-not [string]::IsNullOrWhiteSpace($tb.Text) -and (Test-Path -LiteralPath $tb.Text)) {
                $dlg.SelectedPath = $tb.Text
            }
        } catch { }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text = $dlg.SelectedPath }
    }.GetNewClosure())

    return @{ TextBox = $tb; NextY = ($y + 8) }
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
    $form.Size = New-Object System.Drawing.Size(560, 830)
    $form.MinimumSize = New-Object System.Drawing.Size(480, 560)
    $form.StartPosition = 'CenterScreen'

    # ---- Carpetes principals -------------------------------------------
    $y = 12
    $grpPrincipals = New-Object System.Windows.Forms.GroupBox
    $grpPrincipals.Text = 'Carpetes principals'
    $grpPrincipals.Location = New-Object System.Drawing.Point(14, 66)
    $grpPrincipals.Size = New-Object System.Drawing.Size(514, 172)
    $grpPrincipals.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $r = _AddConfigRow $grpPrincipals 24 "Carpeta on hi ha els informes ja generats" $effInformesDir
    $tbInformes = $r.TextBox
    $r = _AddConfigRow $grpPrincipals $r.NextY "Carpeta de l'Excel d'activitats" $effActivitatsDir
    $tbActivitats = $r.TextBox

    # ---- Carpetes addicionals ------------------------------------------
    $grpAddicionals = New-Object System.Windows.Forms.GroupBox
    $grpAddicionals.Text = 'Carpetes addicionals'
    $grpAddicionals.Location = New-Object System.Drawing.Point(14, 246)
    $grpAddicionals.Size = New-Object System.Drawing.Size(514, 320)
    $grpAddicionals.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $r = _AddConfigRow $grpAddicionals 24 "Carpeta on desar els informes que generis" $effOutputDir
    $tbOutput = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta on desar els mapes de ruta" $effRutesOutputDir
    $tbRutes = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta del Drive d'escriptori (per al mobil)" $effDriveBaseDir
    $tbDrive = $r.TextBox
    $r = _AddConfigRow $grpAddicionals $r.NextY "Carpeta on copiar els informes (copia de seguretat)" $effCopiaInformesDir
    $tbCopia = $r.TextBox

    # ---- Manteniment: info de versio + actualitzar ----------------------
    $grpMant = New-Object System.Windows.Forms.GroupBox
    $grpMant.Text = 'Manteniment'
    $grpMant.Location = New-Object System.Drawing.Point(14, 578)
    $grpMant.Size = New-Object System.Drawing.Size(514, 96)
    $grpMant.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

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
    $btnActualitzar.Text = "$([char]0x21BB) Actualitzar el programa"
    $btnActualitzar.Location = New-Object System.Drawing.Point(14, 52)
    $btnActualitzar.Size = New-Object System.Drawing.Size(260, 32)
    _StyleSecondaryButton $btnActualitzar
    $btnActualitzar.add_Click({
        $batPath = Join-Path $RepoRoot 'Actualitzar.bat'
        if (-not (Test-Path -LiteralPath $batPath)) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha trobat Actualitzar.bat a:`n$batPath", 'Actualitzar el programa', 'OK', 'Error') | Out-Null
            return
        }
        $rr = [System.Windows.Forms.MessageBox]::Show(
            "S'obrira una finestra per actualitzar el programa des de GitHub. El programa es tancara mentre s'actualitza.`n`nVols continuar?",
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
    }.GetNewClosure())
    [void]$grpMant.Controls.Add($btnActualitzar)

    # ---- Barra inferior: Desar / Restaura / Tancar -----------------------
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'
    $botPanel.Height = 52

    $btnDesar = New-Object System.Windows.Forms.Button
    $btnDesar.Text = 'Desar'
    $btnDesar.Size = New-Object System.Drawing.Size(110, 32)
    $btnDesar.Location = New-Object System.Drawing.Point(14, 10)
    _StylePrimaryButton $btnDesar
    [void]$botPanel.Controls.Add($btnDesar)

    $btnRestaura = New-Object System.Windows.Forms.Button
    $btnRestaura.Text = 'Restaura els valors per defecte'
    $btnRestaura.Size = New-Object System.Drawing.Size(210, 32)
    $btnRestaura.Location = New-Object System.Drawing.Point(134, 10)
    _StyleSecondaryButton $btnRestaura
    [void]$botPanel.Controls.Add($btnRestaura)

    $btnTancar = New-Object System.Windows.Forms.Button
    $btnTancar.Text = 'Tancar'
    $btnTancar.Size = New-Object System.Drawing.Size(110, 32)
    $btnTancar.Location = New-Object System.Drawing.Point(354, 10)
    _StyleSecondaryButton $btnTancar
    [void]$botPanel.Controls.Add($btnTancar)

    $btnRestaura.add_Click({
        $tbInformes.Text   = $Script:DefaultInformesDir
        $tbActivitats.Text = $Script:DefaultActivitatsDir
        $tbOutput.Text     = $Script:DefaultOutputDir
        $tbRutes.Text      = $Script:DefaultRutesOutputDir
        $tbDrive.Text      = $Script:DefaultDriveBaseDir
        $tbCopia.Text      = $Script:DefaultCopiaInformesDir
    }.GetNewClosure())

    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())

    $btnDesar.add_Click({
        $values = @{
            InformesDir      = $tbInformes.Text.Trim()
            ActivitatsDir    = $tbActivitats.Text.Trim()
            OutputDir        = $tbOutput.Text.Trim()
            RutesOutputDir   = $tbRutes.Text.Trim()
            DriveBaseDir     = $tbDrive.Text.Trim()
            CopiaInformesDir = $tbCopia.Text.Trim()
        }
        $defaults = @{
            InformesDir      = $Script:DefaultInformesDir
            ActivitatsDir    = $Script:DefaultActivitatsDir
            OutputDir        = $Script:DefaultOutputDir
            RutesOutputDir   = $Script:DefaultRutesOutputDir
            DriveBaseDir     = $Script:DefaultDriveBaseDir
            CopiaInformesDir = $Script:DefaultCopiaInformesDir
        }
        $overrides = _BuildSettingsOverrides $values $defaults
        if (-not (Save-AppSettings $overrides)) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar la configuracio.", 'Configuracio', 'OK', 'Error') | Out-Null
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
    [void]$form.Controls.Add($grpMant)
    [void]$form.Controls.Add($botPanel)
    [void](_AddBrandHeader $form ('Configuraci' + [char]0x00F3) ("Nom" + [char]0x00E9 + "s afecta aquest ordinador: no es comparteix ni es puja a GitHub.") 56)

    [void]$form.ShowDialog()
}
