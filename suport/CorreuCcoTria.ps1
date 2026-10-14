#requires -Version 5.1
<#
  CorreuCcoTria.ps1 - La finestreta "CCO d'aquest enviament": abans d'una tanda
  enviada A MA (Recordatoris, Controls periodics, Seguiment), les adreces de
  CCO "a triar" de l'eina, desmarcades, perque l'usuari marqui les que vol
  (octubre 2026). Les de sempre ja hi van i es diuen a sobre.

  Per que un fitxer a part i no a CorreuEinesPantalla.ps1: aquella crida les
  eines (l'editor de textos de cadascuna) i les eines criden aquesta; juntes
  farien un cicle de dependencies (hi ha guard). "Enviar correu" no la fa
  servir: la seva finestra ja te les caselles de la CCO (_CorreuCcoOpcions).

  NOMES DEFINEIX FUNCIONS.
#>

# Torna @{ Ok; Adreces } (Ok $false = l'usuari ha cancel·lat). Sense cap adreca
# a triar, no obre res i torna Ok amb cap adreca. Un hashtable i no la llista
# sola: una llista buida tornaria $null i no es distingiria de cancel·lar.
function Select-CorreuCcoTria([string]$clau, $cfgE = $null) {
    $cc = Get-CorreuCcoEina $clau $cfgE
    $tria = @(_CorreuLlistaAdreces ([string]$cc.Tria))
    if ($tria.Count -eq 0) { return @{ Ok = $true; Adreces = @() } }
    $def = $Script:CorreuEines[$clau]

    $form = _NewForm
    $form.Text = "CCO d'aquest enviament"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(16, 72); $l.Size = New-Object System.Drawing.Size(448, 40)
    $l.Text = ('Vols afegir alguna d' + [char]0x2019 + 'aquestes adreces en c' + [char]0x00F2 + 'pia oculta (CCO)?' +
               $(if ([string]$cc.Sempre) { "`nJa hi van sempre: " + [string]$cc.Sempre } else { '' }))
    $l.AutoEllipsis = $true
    [void]$form.Controls.Add($l)
    $checks = New-Object System.Collections.ArrayList
    $y = 118
    foreach ($a in $tria) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = [string]$a
        $cb.Location = New-Object System.Drawing.Point(28, $y); $cb.Size = New-Object System.Drawing.Size(430, 22)
        [void]$form.Controls.Add($cb)
        [void]$checks.Add($cb)
        $y += 26
    }
    $yPeu = $y + 14
    [void](_AddPeuBotons $form @(@{ Nom = 'Cancel'; Text = ('Cancel' + [char]0x00B7 + 'lar'); Resultat = 'Cancel'; Esc = $true }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Resultat = 'OK'; Intro = $true; Estil = 'primari' }) $yPeu)
    $form.ClientSize = New-Object System.Drawing.Size(480, ($yPeu + 46))
    [void](_AddBrandHeader $form "CCO d'aquest enviament" ([string]$def.Nom) 56)
    $r = $form.ShowDialog()
    $out = @(@($checks) | Where-Object { $_.Checked } | ForEach-Object { [string]$_.Text })
    $form.Dispose()
    if ($r -ne [System.Windows.Forms.DialogResult]::OK) { return @{ Ok = $false; Adreces = @() } }
    return @{ Ok = $true; Adreces = $out }
}
