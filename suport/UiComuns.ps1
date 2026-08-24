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
# L'AppUserModelID viu a AccesDirecte.ps1 i NOMES alli: es el mateix que es posa
# a la drecera de la barra de tasques, i si els dos no coincideixen Windows els
# tracta com dues aplicacions diferents. AccesDirecte.ps1 no depen de res, o
# sigui que carregar-lo aqui es segur (i Motor.ps1 el torna a carregar despres,
# que es idempotent).
if (-not (Get-Command Get-AccesDirecteObjectiu -ErrorAction SilentlyContinue)) {
    try { . (Join-Path $ScriptRoot 'AccesDirecte.ps1') } catch { }
}
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
        [CornellaApp.Shell]::SetCurrentProcessExplicitAppUserModelID([string]$Script:AppUserModelId)
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
    # Cada finestra es comprova a si mateixa en obrir-se (vegeu mes avall).
    $f.add_Shown({ param($s, $e) _AvisaSolapaments $s }.GetNewClosure())
    return $f
}

# ----------------------------------------------------------------------------
# CONTROLS QUE ES TREPITGEN: detectar-ho SOL, a totes les pantalles
# ----------------------------------------------------------------------------
# Els solapaments son el defecte recurrent d'aquest programa: coordenades
# posades a ma, i un text que creix o un control nou que hi passa per sobre. El
# darrer va ser el xip "Dades" tapant el "LL Prov" del menu.
#
# Les proves no ho poden veure -les finestres nomes es dibuixen a Windows-, o
# sigui que qui ho ha de veure es EL PROGRAMA MATEIX: cada finestra es mira en
# obrir-se i, si hi ha controls germans que es trepitgen, ho diu amb els noms i
# les coordenades. Val mes un avis lleig un cop que una pantalla mig tapada
# durant setmanes.

# La part de decisio es PURA i es prova a Linux: rebre rectangles i dir quins
# parells es trepitgen.
#
# $rects: llista de @{ Nom; X; Y; W; H }.
#
# NO es solapament:
#   - que un CONTINGUI l'altre del tot (un fons a posta);
#   - que nomes es toquin per la vora;
#   - i, sobretot, que es TOQUIN DE POC. Aquesta ultima condicio es la que fa
#     que la comprovacio serveixi de res: un Label amb AutoSize reserva mes
#     alcada de la que pinta (la del descens de les lletres), o sigui que dues
#     linies apilades a la distancia normal es "solapen" 1-5 px SEMPRE. Sense
#     tolerancia, la meitat de les pantalles del programa donaven avis i
#     l'usuari deixava de mirar-los -que es pitjor que no tenir-ne cap-.
$Script:SolapMinPx = 8      # ha de trepitjar com a minim aixo en LES DUES direccions
$Script:SolapMinPct = 0.15  # ...i cobrir aquesta part del control mes petit
function _TrobaSolapaments($rects) {
    $out = New-Object System.Collections.ArrayList
    $l = @($rects)
    for ($i = 0; $i -lt $l.Count; $i++) {
        for ($j = $i + 1; $j -lt $l.Count; $j++) {
            $a = $l[$i]; $b = $l[$j]
            $ax2 = [int]$a.X + [int]$a.W; $ay2 = [int]$a.Y + [int]$a.H
            $bx2 = [int]$b.X + [int]$b.W; $by2 = [int]$b.Y + [int]$b.H
            if (-not ([int]$a.X -lt $bx2 -and [int]$b.X -lt $ax2 -and
                      [int]$a.Y -lt $by2 -and [int]$b.Y -lt $ay2)) { continue }
            # Un dins de l'altre del tot: es un fons, no un error.
            $aDinsB = ([int]$b.X -le [int]$a.X -and [int]$b.Y -le [int]$a.Y -and $bx2 -ge $ax2 -and $by2 -ge $ay2)
            $bDinsA = ([int]$a.X -le [int]$b.X -and [int]$a.Y -le [int]$b.Y -and $ax2 -ge $bx2 -and $ay2 -ge $by2)
            if ($aDinsB -or $bDinsA) { continue }
            # Es toquen de poc: no molesta ningu.
            $ampleTrepitjat = [Math]::Min($ax2, $bx2) - [Math]::Max([int]$a.X, [int]$b.X)
            $altTrepitjada  = [Math]::Min($ay2, $by2) - [Math]::Max([int]$a.Y, [int]$b.Y)
            if ($ampleTrepitjat -lt $Script:SolapMinPx -or $altTrepitjada -lt $Script:SolapMinPx) { continue }
            $areaMinima = [Math]::Min(([int]$a.W * [int]$a.H), ([int]$b.W * [int]$b.H))
            if ($areaMinima -le 0) { continue }
            if ((($ampleTrepitjat * $altTrepitjada) / [double]$areaMinima) -lt $Script:SolapMinPct) { continue }
            [void]$out.Add(('{0} [{1},{2} {3}x{4}] i {5} [{6},{7} {8}x{9}]' -f `
                $a.Nom, $a.X, $a.Y, $a.W, $a.H, $b.Nom, $b.X, $b.Y, $b.W, $b.H))
        }
    }
    return $out.ToArray()
}

# El nom amb que surt un control a l'avis: el text si en te, si no el tipus.
function _NomControl($c) {
    $t = ''
    try { $t = ([string]$c.Text -replace '\s+', ' ').Trim() } catch { $t = '' }
    $tipus = $c.GetType().Name
    if ([string]::IsNullOrWhiteSpace($t)) { return $tipus }
    if ($t.Length -gt 28) { $t = $t.Substring(0, 28) + [char]0x2026 }
    return ($tipus + " '" + $t + "'")
}

# Recorre l'arbre de controls d'un contenidor i compara NOMES ELS GERMANS (dins
# d'un contenidor les coordenades son relatives a ell: comparar-les entre
# contenidors diferents no vol dir res).
function _SolapamentsDeContenidor($cont) {
    $out = New-Object System.Collections.ArrayList
    $fills = New-Object System.Collections.ArrayList
    foreach ($c in $cont.Controls) {
        if (-not $c.Visible) { continue }
        # Els Dock els col-loca WinForms; per definicio no es trepitgen.
        try { if ([string]$c.Dock -ne 'None') { continue } } catch { }
        [void]$fills.Add(@{ Nom = (_NomControl $c); X = $c.Left; Y = $c.Top; W = $c.Width; H = $c.Height })
    }
    foreach ($s in @(_TrobaSolapaments $fills)) { [void]$out.Add($s) }
    foreach ($c in $cont.Controls) {
        if ($c.Controls.Count -gt 0) {
            foreach ($s in @(_SolapamentsDeContenidor $c)) { [void]$out.Add($s) }
        }
    }
    return $out.ToArray()
}

# Nomes s'avisa UNA vegada per pantalla i sessio: si no, un formulari que es
# repinta seria inaguantable.
$Script:SolapamentsAvisats = @{}
function _AvisaSolapaments($form) {
    if ($Script:HeadlessTest) { return }
    try {
        $clau = [string]$form.Text
        if ($Script:SolapamentsAvisats.ContainsKey($clau)) { return }
        $Script:SolapamentsAvisats[$clau] = $true
        $sol = @(_SolapamentsDeContenidor $form)
        if ($sol.Count -eq 0) { return }
        [System.Windows.Forms.MessageBox]::Show(
            ("A la pantalla '" + $clau + "' hi ha " + $sol.Count + " control(s) que es trepitgen:`n`n  " +
             (($sol | Select-Object -First 8) -join "`n  ") +
             "`n`nL'informe es genera igual; digues-ho perque ho arreglin."),
            'Pantalla mal col-locada', 'OK', 'Warning') | Out-Null
    } catch { }   # una comprovacio no pot impedir obrir una finestra
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

# ----------------------------------------------------------------------------
# Graelles de llistat (DataGridView)
# ----------------------------------------------------------------------------
# Les pantalles "Editar base d'informes" (Informes.ps1) i "Controls periodics"
# (ControlsPeriodics.ps1) son totes dues un llistat amb cerca, filtres i ordre
# per capcalera. NO comparteixen un widget unic: les columnes, el criteri de
# filtratge i el d'ordenacio son massa diferents (l'una agrupa SEMPRE per
# activitat, l'altra ordena per data real i te una casella de seleccio), i un
# "constructor de graelles" generic acabaria sent mes complicat que les dues
# pantalles juntes. El que si compartim son les peces que estaven copiades
# literalment: la carcassa, la caixa de cerca i l'ordre per capcalera.

# Carcassa comuna d'una graella de llistat (nomes lectura de files, seleccio
# de fila sencera, sense capcaleres de fila). Les columnes les afegeix qui
# crida, DESPRES.
function _StyleListGrid($grid) {
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'None'
    $grid.BackgroundColor = [System.Drawing.Color]::White
}

# Etiqueta + caixa de cerca. Retorna el TextBox (el qui crida el llegeix des
# del seu $fill). $onChange es crida a cada tecla.
function _AddSearchBox($parent, [int]$x, [int]$y, [int]$width, [string]$labelText, [scriptblock]$onChange) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($x, ($y + 3))
    [void]$parent.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(($x + 52), $y)
    $tb.Width = $width
    [void]$parent.Controls.Add($tb)
    if ($onChange) { $tb.add_TextChanged($onChange) }
    return $tb
}

# Pinta la fletxa d'ordenacio NOMES a la columna activa (WinForms no ho fa sol
# quan l'ordenacio es programatica).
function _SetSortGlyph($grid, [int]$colIdx, [bool]$asc) {
    foreach ($c in $grid.Columns) { $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
    if ($colIdx -lt 0 -or $colIdx -ge $grid.Columns.Count) { return }
    $grid.Columns[$colIdx].HeaderCell.SortGlyphDirection =
        if ($asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
}

# Ordre per clic a la capcalera, en mode PROGRAMATIC: clicar una columna la
# tria com a columna d'ordre; tornar-hi a clicar alterna asc/desc. Desa l'estat
# a $state.SortColIdx / $state.SortAsc i despres crida $onSort (que ha de
# reomplir la graella). Les columnes amb index < $minCol i les de $skipCols
# no s'ordenen (caselles, botons).
#   IMPORTANT: qui cridi aixo ha de posar SortMode='Programmatic' a les
#   columnes ordenables, si no WinForms ordenaria pel seu compte.
function _EnableHeaderSort($grid, $state, [int]$minCol = 0, $skipCols = @(), [scriptblock]$onSort) {
    $grid.add_ColumnHeaderMouseClick({
        param($s, $e)
        if ($e.ColumnIndex -lt $minCol) { return }
        if ($skipCols -contains $e.ColumnIndex) { return }
        if ($state.SortColIdx -eq $e.ColumnIndex) { $state.SortAsc = (-not $state.SortAsc) }
        else { $state.SortColIdx = $e.ColumnIndex; $state.SortAsc = $true }
        _SetSortGlyph $grid $state.SortColIdx ([bool]$state.SortAsc)
        & $onSort
    }.GetNewClosure())
}

# NOTA: per al filtre de text lliure de les graelles NO hi ha cap helper nou
# aqui: ja existeix _TextMatches a Motor.ps1 (el fa servir tambe el cercador
# del TreeView de catalegs) i el reutilitzem. Els qui criden li passen el text
# de cerca JA net (.Trim()), que es el seu contracte.

# Codi C# (COM) del navegador de carpetes MODERN (IFileOpenDialog amb
# FOS_PICKFOLDERS): el diàleg estil Explorer amb barra d'adreça (on es pot
# enganxar la ruta), panell lateral d'unitats/xarxa i cerca. Es compila EN VIU
# el primer cop que es fa servir (mai en headless).
$Script:FolderPickerCSharp = @"
using System;
using System.Runtime.InteropServices;
namespace CornellaShell {
  [ComImport, ClassInterface(ClassInterfaceType.None), Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
  internal class FileOpenDialogRCW { }

  [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  internal interface IShellItem {
    void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
  }

  [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  internal interface IFileDialog {
    [PreserveSig] int Show(IntPtr parent);
    void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
    void SetFileTypeIndex(uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise(IntPtr pfde, out uint pdwCookie);
    void Unadvise(uint dwCookie);
    void SetOptions(uint fos);
    void GetOptions(out uint fos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, uint fdap);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close(int hr);
    void SetClientGuid(ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
  }

  public static class FolderPicker {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void SHCreateItemFromParsingName(
      [MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc,
      ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

    public static string Pick(string initialPath, string title, IntPtr owner) {
      IFileDialog dlg = (IFileDialog)new FileOpenDialogRCW();
      uint opts;
      dlg.GetOptions(out opts);
      // FOS_PICKFOLDERS=0x20 | FOS_FORCEFILESYSTEM=0x40 | FOS_PATHMUSTEXIST=0x800
      dlg.SetOptions(opts | 0x20 | 0x40 | 0x800);
      if (!string.IsNullOrEmpty(title)) { try { dlg.SetTitle(title); } catch {} }
      if (!string.IsNullOrEmpty(initialPath)) {
        try {
          Guid iid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
          IShellItem item;
          SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref iid, out item);
          if (item != null) dlg.SetFolder(item);
        } catch {}
      }
      int hr = dlg.Show(owner);
      if (hr != 0) return "";   // S_OK=0; cancel·lat o error -> buit
      IShellItem res;
      dlg.GetResult(out res);
      string path;
      res.GetDisplayName(0x80058000u, out path);   // SIGDN_FILESYSPATH
      return path == null ? "" : path;
    }
  }
}
"@

# Obre el navegador de carpetes MODERN i retorna la ruta triada (o '' si es
# cancel·la). Compila el codi COM el primer cop. Si res falla (Windows molt
# antic, etc.) recorre al FolderBrowserDialog classic com a xarxa de seguretat.
function _PickFolderModern([string]$initialPath, [string]$title, $ownerForm) {
    try {
        if (-not ([System.Management.Automation.PSTypeName]'CornellaShell.FolderPicker').Type) {
            Add-Type -TypeDefinition $Script:FolderPickerCSharp -ErrorAction Stop
        }
        $owner = [IntPtr]::Zero
        if ($null -ne $ownerForm) { try { $owner = $ownerForm.Handle } catch { $owner = [IntPtr]::Zero } }
        return [string][CornellaShell.FolderPicker]::Pick($initialPath, $title, $owner)
    } catch {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if (-not [string]::IsNullOrWhiteSpace($title)) { $dlg.Description = $title }
        try { if (-not [string]::IsNullOrWhiteSpace($initialPath) -and (Test-Path -LiteralPath $initialPath)) { $dlg.SelectedPath = $initialPath } } catch { }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return [string]$dlg.SelectedPath }
        return ''
    }
}

# Fila d'una carpeta configurable: etiqueta + textbox + boto "..." (Explora, amb
# el navegador MODERN _PickFolderModern) + indicador d'estat en viu (Test-Path).
# Afegeix els controls a $parent i retorna @{ TextBox = ...; NextY = ... } per
# encadenar files. TOTS els selectors de carpeta del programa fan servir aquest
# format (pantalla Configuracio, Word a PDF...).
# $fileFilter (opcional): si s'indica (p. ex. 'Documents Word|*.docx;*.doc'),
# s'afegeix un SEGON boto que obre un dialeg de FITXERS, de manera que al mateix
# quadre s'hi pot posar una CARPETA o un DOCUMENT concret (ho fa servir Word a PDF).
function _AddConfigRow($parent, [int]$y, [string]$labelText, [string]$initialValue, [string]$fileFilter = '') {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lbl.Location = New-Object System.Drawing.Point(14, $y)
    $lbl.AutoSize = $true
    [void]$parent.Controls.Add($lbl)
    $y += 20

    # Amb filtre de fitxer hi ha DOS botons (carpeta + document): el quadre es
    # fa una mica mes estret per encabir-los tots dos.
    $ambFitxer = (-not [string]::IsNullOrWhiteSpace($fileFilter))
    $tbWidth = if ($ambFitxer) { 386 } else { 432 }

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(14, $y)
    $tb.Size = New-Object System.Drawing.Size($tbWidth, 24)
    $tb.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $tb.Text = $initialValue
    [void]$parent.Controls.Add($tb)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    if ($ambFitxer) {
        $btn.Text = 'Carpeta'
        $btn.Location = New-Object System.Drawing.Point(404, ($y - 1))
        $btn.Size = New-Object System.Drawing.Size(58, 24)
    } else {
        $btn.Text = '...'
        $btn.Location = New-Object System.Drawing.Point(452, ($y - 1))
        $btn.Size = New-Object System.Drawing.Size(36, 24)
    }
    [void]$parent.Controls.Add($btn)

    if ($ambFitxer) {
        $btnFile = New-Object System.Windows.Forms.Button
        $btnFile.Text = 'Document'
        $btnFile.Location = New-Object System.Drawing.Point(466, ($y - 1))
        $btnFile.Size = New-Object System.Drawing.Size(70, 24)
        $btnFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        [void]$parent.Controls.Add($btnFile)
        $btnFile.add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = $labelText
            $dlg.Filter = $fileFilter
            $dlg.CheckFileExists = $true
            try {
                $cur = [string]$tb.Text
                if (-not [string]::IsNullOrWhiteSpace($cur)) {
                    if (Test-Path -LiteralPath $cur -PathType Container) { $dlg.InitialDirectory = $cur }
                    elseif (Test-Path -LiteralPath $cur) { $dlg.InitialDirectory = (Split-Path -Parent $cur) }
                }
            } catch { }
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text = $dlg.FileName }
        }.GetNewClosure())
    }
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
            $esFitxer = $false
            try { $esFitxer = Test-Path -LiteralPath $p -PathType Leaf } catch { }
            if ($esFitxer) { $status.Text = "$([char]0x2713) Document trobat" }
            else           { $status.Text = "$([char]0x2713) Trobada" }
            $status.ForeColor = [System.Drawing.Color]::SeaGreen
        } else {
            $status.Text = "$([char]0x26A0) No trobada ara (es pot desar igualment)"
            $status.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    }.GetNewClosure()
    $tb.add_TextChanged($refreshStatus)
    & $refreshStatus

    $btn.add_Click({
        $sel = _PickFolderModern $tb.Text $labelText $parent
        if (-not [string]::IsNullOrWhiteSpace($sel)) { $tb.Text = $sel }
    }.GetNewClosure())

    return @{ TextBox = $tb; NextY = ($y + 8) }
}
