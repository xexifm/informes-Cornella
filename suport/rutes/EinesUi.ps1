<#
  EinesUi.ps1 - Les finestres que comparteixen les eines de 'rutes/' que corren
  al seu propi ambit (Coordenades, Planol activitats): un missatge, una
  pregunta amb botons propis, la barra de progres amb Cancel.lar de debo i la icona.

  Per que no son a UiComuns.ps1: aquestes eines no el poden carregar (executa
  coses en carregar-se: AppUserModelID, icona del proces). Venien de
  Coordenades.ps1; el Planol en necessita exactament les mateixes.

  NOMES DEFINEIX FUNCIONS. Qui les fa servir posa $Script:EinaTitol (el titol
  dels missatges) i $Script:EinaIcon.
#>

$Script:EinaTitol = 'Eina'
$Script:EinaIcon = $null

# Icona corporativa (suport\cornella.ico), o $null.
function Get-EinaIcon([string]$suportDir) {
    try {
        $ruta = Join-Path $suportDir 'cornella.ico'
        if (Test-Path -LiteralPath $ruta) { return (New-Object System.Drawing.Icon($ruta)) }
    } catch { }
    return $null
}

function Show-EinaInfo([string]$msg, [string]$title = '', [string]$icon = 'Information') {
    if ($title -eq '') { $title = $Script:EinaTitol }
    [System.Windows.Forms.MessageBox]::Show($msg, $title, 'OK', $icon) | Out-Null
}

# Una pregunta amb botons propis (el MessageBox nomes en sap de Si/No). Els
# botons son specs de _AddPeuBotons (UiFinestra.ps1): @{ Nom; Text; Estil;
# Intro; Esc }. Torna el Nom del boto premut, o '' si es tanca la finestra.
function Show-EinaTria([string]$msg, $esquerra, $dreta, [string]$title = '') {
    if ($title -eq '') { $title = $Script:EinaTitol }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.ClientSize = New-Object System.Drawing.Size(520, 190)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false; $form.MaximizeBox = $false
    if ($null -ne $Script:EinaIcon) { $form.Icon = $Script:EinaIcon }
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(480, 110)
    $lbl.Text = $msg
    $form.Controls.Add($lbl)
    $tria = @{ Nom = '' }
    $peu = _AddPeuBotons $form $esquerra $dreta 142
    foreach ($k in @($peu.Keys)) {
        $nom = [string]$k
        $peu[$k].add_Click({ $tria.Nom = $nom; $form.Close() }.GetNewClosure())
    }
    [void]$form.ShowDialog()
    $form.Dispose()
    return [string]$tria.Nom
}

# Finestra de progres amb Cancel.lar DE DEBO. Retorna { Form; Label; Bar; Estat }:
# el boto posa Estat.Cancelat, i el bloc de progres que es passa al bucle de
# Cadastre.ps1 ha de tornar (-not $prog.Estat.Cancelat). Qui la crida ha de fer
# .Form.Close() al final.
#
# UN HASHTABLE i no $Script:..., i es MESURAT: el bloc de progres porta
# .GetNewClosure() (ha de veure $prog), i dins d'una closure el $Script: es el
# de la closure, no el de l'eina. Fins a l'octubre de 2026 el boto posava
# $Script:CoordCancelat i el bloc el llegia des de la closure: no el veia mai, i
# el Cancel.lar de Coordenades no aturava res. Un hashtable es una referencia:
# el boto i la closure veuen el mateix.
function New-EinaProgres([int]$total, [string]$titol = 'Consultant el Cadastre', [string]$text = '') {
    $estat = @{ Cancelat = $false }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(560, 190)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false; $form.MaximizeBox = $false
    $form.ControlBox = $false
    if ($null -ne $Script:EinaIcon) { $form.Icon = $Script:EinaIcon }
    # Scroll vertical i ajust a la pantalla (vegeu suport/UiFinestra.ps1).
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(510, 42)
    $lbl.Text = if ($text -ne '') { $text } else { 'Demanant ' + $total + ' consultes al Cadastre...' }
    $form.Controls.Add($lbl)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 66)
    $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Minimum = 0
    $bar.Maximum = [math]::Max($total, 1)
    $form.Controls.Add($bar)

    [void](_AddPeuBotons $form @(@{ Nom = 'Cancel'; Text = ('Cancel' + [char]0x00B7 + 'lar'); Clic = { $estat.Cancelat = $true }.GetNewClosure() }) @() 98)

    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    return [pscustomobject]@{ Form = $form; Label = $lbl; Bar = $bar; Estat = $estat }
}
