#requires -Version 5.1
<#
.SYNOPSIS
  Helpers d'interficie (WinForms) COMPARTITS per tot el programa.

.DESCRIPTION
  Aquest modul conte la "carcassa" visual comuna: l'escut i la icona de la
  barra de tasques, la finestra estandard (_NewForm), la banda granat de
  capcalera (_AddBrandHeader), la barra de passos (_AddStepBar), els estils de
  boto i els dos controls compostos que es repeteixen a diverses pantalles:

    _MakeMultiFilter  desplegable de seleccio multiple (Editar base d'informes,
                      Controls periodics)
    _AddConfigRow     fila "quadre editable + ... + indicador ✓/⚠" per triar
                      una carpeta (Configuracio, Word a PDF)

  PER QUE UN MODUL A PART: abans _MakeMultiFilter vivia al punt d'entrada
  (GenerarInforme.ps1) i _AddConfigRow a Configuracio.ps1, pero els feien
  servir moduls que es carreguen abans o que no en depenen conceptualment
  (Informes.ps1, ControlsPeriodics.ps1, PdfSignar.ps1). Nomes funcionava
  perque el dot-source ho aboca tot al mateix ambit global. Aixi la dependencia
  va en la direccio correcta: els moduls depenen d'UiComuns, mai del programa.

.NOTES
  Es carrega el PRIMER des de Motor.ps1. Nomes depen de WinForms/Drawing,
  de $ScriptRoot i de $Script:HeadlessTest (en headless no dibuixa res).
  No coneix res del motor d'informes: no hi posis logica de negoci.
#>

# Icona corporativa (escut de Cornella) per a TOTES les finestres i la
# miniatura de la barra de tasques de Windows. Nomes en mode interactiu (en
# headless no hi ha System.Drawing carregat). Es carrega un sol cop.
$Script:AppIcon = $null
if (-not $Script:HeadlessTest) {
    # AppUserModelID propi: sense aixo, la barra de tasques agrupa la finestra
    # sota el proces amfitrio (PowerShell) i mostra la SEVA icona blava. Amb un
    # ID propi, Windows tracta el programa com una app a part i la barra de
    # tasques fa servir la icona de la finestra (l'escut). Cal fer-ho ABANS de
    # crear cap finestra.
    try {
        Add-Type -Namespace CornellaApp -Name Shell -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("shell32.dll", SetLastError=true)]
public static extern void SetCurrentProcessExplicitAppUserModelID([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string AppID);
'@ -ErrorAction Stop
        [CornellaApp.Shell]::SetCurrentProcessExplicitAppUserModelID('Cornella.Informes.Generador')
    } catch { }
    try {
        $iconPath = Join-Path $ScriptRoot 'cornella.ico'
        if (Test-Path -LiteralPath $iconPath) { $Script:AppIcon = New-Object System.Drawing.Icon($iconPath) }
    } catch { $Script:AppIcon = $null }

    # Escut BLANC (per a la banda granat de capcalera de totes les pantalles).
    # Es un PNG amb fons transparent; es carrega un sol cop.
    $Script:EscutBlanc = $null
    try {
        $escutPath = Join-Path $ScriptRoot 'escut-blanc.png'
        if (Test-Path -LiteralPath $escutPath) { $Script:EscutBlanc = [System.Drawing.Image]::FromFile($escutPath) }
    } catch { $Script:EscutBlanc = $null }

    # Helper user32 per portar a primer pla la finestra d'una instancia ja oberta
    # del programa (una sola instancia: si es torna a llancar, s'enfoca la que
    # ja hi ha en lloc d'obrir-ne una segona). Donat el PID del proces de la
    # instancia viva, busca la seva finestra de nivell superior visible i la
    # restaura + la posa al davant.
    try {
        Add-Type -Namespace CornellaApp -Name Win -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, System.IntPtr lParam);
delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool IsWindowVisible(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetForegroundWindow(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll")] static extern System.IntPtr GetWindow(System.IntPtr hWnd, uint uCmd);
const int SW_RESTORE = 9;
static System.IntPtr _found;
static uint _target;
static bool _Cb(System.IntPtr h, System.IntPtr l) {
    uint p; GetWindowThreadProcessId(h, out p);
    // GW_OWNER = 4: nomes finestres de nivell superior (sense propietari).
    if (p == _target && IsWindowVisible(h) && GetWindow(h, 4) == System.IntPtr.Zero) { _found = h; return false; }
    return true;
}
public static bool FocusProcessWindow(int pid) {
    _found = System.IntPtr.Zero; _target = (uint)pid;
    EnumWindows(_Cb, System.IntPtr.Zero);
    if (_found == System.IntPtr.Zero) return false;
    ShowWindow(_found, SW_RESTORE);
    SetForegroundWindow(_found);
    return true;
}
'@ -ErrorAction Stop
    } catch { }
}

# Crea una finestra estandard de l'app: centrada, MINIMITZABLE, MAXIMITZABLE i
# tancable, amb l'escut de Cornella al titol i a la barra de tasques. Totes les
# pantalles del programa (inclos Seguiment i ACT_EXTR) la fan servir, aixi el
# logo i els botons de finestra son consistents a tot arreu.
function _NewForm {
    $f = New-Object System.Windows.Forms.Form
    $f.StartPosition = 'CenterScreen'
    $f.MinimizeBox = $true
    $f.MaximizeBox = $true
    if ($null -ne $Script:AppIcon) { $f.Icon = $Script:AppIcon }
    return $f
}

# Color corporatiu granat (redisseny UX/UI). Es fa servir a la banda de
# capcalera de totes les pantalles. NOMES en interactiu: en headless (Linux,
# proves) System.Drawing no esta carregat i [System.Drawing.Color] petaria en
# carregar el motor.
$Script:BrandMaroon     = $null
$Script:BrandMaroonSoft = $null
if (-not $Script:HeadlessTest) {
    $Script:BrandMaroon     = [System.Drawing.Color]::FromArgb(166, 26, 47)
    $Script:BrandMaroonSoft = [System.Drawing.Color]::FromArgb(247, 231, 234)
}

# Afegeix a $form una BANDA superior granat (Dock=Top) amb l'escut blanc,
# un titol i un subtitol opcional. Retorna el Panel (per si el qui crida hi vol
# afegir botons a la dreta). Es la capcalera comuna del redisseny; totes les
# pantalles l'han de fer servir per unificar l'aspecte.
#   S'ha d'afegir DESPRES dels controls posicionats en absolut (com fan els
#   Dock=Top/Bottom) perque no els desplaci.
function _AddBrandHeader($form, [string]$title, [string]$subtitle, [int]$height = 56) {
    $band = New-Object System.Windows.Forms.Panel
    $band.Dock = 'Top'
    $band.Height = $height
    $band.BackColor = $Script:BrandMaroon
    $xText = 18
    if ($null -ne $Script:EscutBlanc) {
        $escH = [int]($height - 22); if ($escH -lt 18) { $escH = 18 }
        $escW = [int]([double]$Script:EscutBlanc.Width / $Script:EscutBlanc.Height * $escH)
        $pic = New-Object System.Windows.Forms.PictureBox
        $pic.Size = New-Object System.Drawing.Size($escW, $escH)
        $pic.Location = New-Object System.Drawing.Point(16, [int](($height - $escH) / 2))
        $pic.SizeMode = 'Zoom'
        $pic.BackColor = [System.Drawing.Color]::Transparent
        $pic.Image = $Script:EscutBlanc
        [void]$band.Controls.Add($pic)
        $sep = New-Object System.Windows.Forms.Panel
        $sep.Size = New-Object System.Drawing.Size(1, [int]($height - 24))
        $sep.Location = New-Object System.Drawing.Point((16 + $escW + 12), 12)
        $sep.BackColor = [System.Drawing.Color]::FromArgb(206, 138, 148)
        [void]$band.Controls.Add($sep)
        $xText = 16 + $escW + 24
    }
    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Text = $title
    $lblT.ForeColor = [System.Drawing.Color]::White
    $lblT.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $lblT.AutoSize = $true
    $lblT.BackColor = [System.Drawing.Color]::Transparent
    if ([string]::IsNullOrWhiteSpace($subtitle)) {
        $lblT.Location = New-Object System.Drawing.Point($xText, [int](($height - 26) / 2))
        [void]$band.Controls.Add($lblT)
    } else {
        $lblT.Location = New-Object System.Drawing.Point($xText, 7)
        [void]$band.Controls.Add($lblT)
        $lblS = New-Object System.Windows.Forms.Label
        $lblS.Text = $subtitle
        $lblS.ForeColor = [System.Drawing.Color]::FromArgb(232, 210, 215)
        $lblS.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
        $lblS.AutoSize = $true
        $lblS.BackColor = [System.Drawing.Color]::Transparent
        $lblS.Location = New-Object System.Drawing.Point(($xText + 2), 33)
        [void]$band.Controls.Add($lblS)
    }
    [void]$form.Controls.Add($band)
    return $band
}

# Barra de PASSOS del wizard (1 Tipus · 2 Dades · 3 Deficiencies · 4 Conclusions).
# Es un Panel Dock=Top; el pas actiu ($active, 1..4) surt en granat i negreta.
# S'ha d'afegir ABANS de la banda (perque la banda quedi a dalt de tot).
function _AddStepBar($form, [int]$active) {
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = 'Top'
    $bar.Height = 34
    $bar.BackColor = [System.Drawing.Color]::White
    $steps = @('Tipus', 'Dades', 'Deficiències', 'Conclusions')
    $fA = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $fI = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
    $granat = $Script:BrandMaroon
    $idle   = [System.Drawing.Color]::FromArgb(140, 148, 158)
    $bar.Tag = @{ Steps=$steps; Active=$active; FA=$fA; FI=$fI; Granat=$granat; Idle=$idle }
    $bar.add_Paint({
        param($s, $e)
        $t = $s.Tag
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rc = $s.ClientRectangle
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(214, 219, 225))
        $g.DrawLine($pen, 0, ($rc.Height - 1), $rc.Width, ($rc.Height - 1)); $pen.Dispose()
        $fl = [System.Windows.Forms.TextFormatFlags]::NoPadding
        $x = 16
        for ($i = 0; $i -lt 4; $i++) {
            $isA = (($i + 1) -eq $t.Active)
            $f = if ($isA) { $t.FA } else { $t.FI }
            $col = if ($isA) { $t.Granat } else { $t.Idle }
            $txt = [string]($i + 1) + '  ' + $t.Steps[$i]
            $sz = [System.Windows.Forms.TextRenderer]::MeasureText($g, $txt, $f, [System.Drawing.Size]::Empty, $fl)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $txt, $f, (New-Object System.Drawing.Point($x, 9)), $col, $fl)
            $x += $sz.Width + 10
            if ($i -lt 3) {
                [System.Windows.Forms.TextRenderer]::DrawText($g, ([string][char]0x203A), $t.FI, (New-Object System.Drawing.Point($x, 9)), $t.Idle, $fl)
                $x += 18
            }
        }
    })
    [void]$form.Controls.Add($bar)
    return $bar
}

# Estil de boto PRIMARI (granat ple, text blanc) i SECUNDARI (blanc, text/vora
# granat). Reutilitzables a totes les pantalles del redisseny.
function _StylePrimaryButton($btn) {
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $Script:BrandMaroon
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(138, 20, 38)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor = 'Hand'
}
function _StyleSecondaryButton($btn) {
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = [System.Drawing.Color]::White
    $btn.ForeColor = $Script:BrandMaroon
    $btn.FlatAppearance.BorderColor = $Script:BrandMaroon
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(247, 231, 234)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
    $btn.Cursor = 'Hand'
}

# Filtre desplegable de SELECCIO MULTIPLE (WinForms no en te de natiu): un boto
# que sembla un desplegable i, en clicar-lo, obre un menu amb items marcables
# (checkbox). El menu no es tanca en marcar (nomes en clicar fora / Escape).
# Retorna @{ Button; Menu; GetSelected } on GetSelected torna l'array de textos
# marcats (buit = cap filtre = totes les files passen). $onChange es crida a
# cada canvi. S'afegeix a $parent a la posicio i amplada donades.
function _MakeMultiFilter($parent, [int]$x, [int]$y, [int]$width, [string]$allLabel, $options, [scriptblock]$onChange) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $btn.BackColor = [System.Drawing.Color]::White
    $btn.TextAlign = 'MiddleLeft'
    $btn.Padding = New-Object System.Windows.Forms.Padding(6, 0, 4, 0)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $btn.Size = New-Object System.Drawing.Size($width, 24)

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    # No tancar el menu en marcar un item (per poder-ne triar diversos).
    $menu.add_Closing({
        param($s, $e)
        if ($e.CloseReason -eq [System.Windows.Forms.ToolStripDropDownCloseReason]::ItemClicked) { $e.Cancel = $true }
    }.GetNewClosure())

    $arrow = [char]0x25BE   # ▾
    $refresh = {
        $sel = @()
        foreach ($it in $menu.Items) { if ($it.Checked) { $sel += [string]$it.Text } }
        if ($sel.Count -eq 0)      { $btn.Text = $allLabel + '   ' + $arrow }
        elseif ($sel.Count -le 2)  { $btn.Text = ($sel -join ', ') + '   ' + $arrow }
        else                       { $btn.Text = ('' + $sel.Count + ' triats   ' + $arrow) }
    }.GetNewClosure()

    foreach ($opt in $options) {
        $it = New-Object System.Windows.Forms.ToolStripMenuItem
        $it.Text = [string]$opt
        $it.CheckOnClick = $true
        $it.add_CheckedChanged({ & $refresh; if ($onChange) { & $onChange } }.GetNewClosure())
        [void]$menu.Items.Add($it)
    }
    $btn.add_Click({ $menu.Show($btn, (New-Object System.Drawing.Point(0, $btn.Height))) }.GetNewClosure())
    & $refresh
    [void]$parent.Controls.Add($btn)

    $getSel = {
        $s = @()
        foreach ($it in $menu.Items) { if ($it.Checked) { $s += [string]$it.Text } }
        return , ([string[]]$s)
    }.GetNewClosure()

    return @{ Button = $btn; Menu = $menu; GetSelected = $getSel }
}

# Fila d'una carpeta configurable: etiqueta + textbox + boto "..." (Explora)
# + indicador d'estat en viu (Test-Path). Afegeix els controls a $parent i
# retorna @{ TextBox = ...; NextY = ... } per encadenar files.
# TOTS els selectors de carpeta del programa han de fer servir aquest format
# (pantalla Configuracio, Word a PDF...).
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
