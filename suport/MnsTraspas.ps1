#requires -Version 5.1
<#
.SYNOPSIS
  Dos informes curts de Llicencia: MODIFICACIO NO SUBSTANCIAL i TRASPAS.

.DESCRIPTION
  Van dins de "Llicencia (Annex II / LL Prov)" perque comparteixen capcalera
  (la de LLIC, amb la linia "Classificacio:") i son del mateix tramit, pero el
  document no s'assembla gens: no hi ha bloc ABANS ni DESPRES ni deficiencies
  de projecte, nomes tres o quatre paragrafs fixos.

  L'UNICA COSA QUE ES TRIA es si hi ha observacions:
    - SI  -> "...amb la seguent observacio:" i, a sota, un paragraf de LLISTA
             DE WORD buit, perque l'usuari hi escrigui el que calgui un cop
             generat el document.
    - NO  -> "...sense mes observacions en relacio a aquest tramit." i cap
             llista.

  EL TEXT NO ES AQUI: viu a ESTRUCTURALS\MNSTRAS.json (un sol cataleg per als
  dos informes, com va demanar l'usuari) i es pot editar des de l'editor de
  catalegs com tota la resta. Cada paragraf hi va com un node:

    tipus 'text' -> paragraf normal (Format-Body)
    tipus 'item' -> paragraf de llista de Word BUIT (Format-ListItem)

  ...i la CLAU diu quan hi entra:

    ''                    -> sempre
    'amb-observacions'    -> nomes si n'hi ha
    'sense-observacions'  -> nomes si no n'hi ha
    'llista-observacions' -> la llista de les observacions (nomes si n'hi ha)

  Al Word que va enviar l'usuari, aquestes dues variants anaven escrites en
  VERMELL, i tambe els titols "MODIFICACIO NO SUBSTANCIAL" i "TRASPAS". Aquell
  color era una marca SEVA per veure que havia de canviar a cada informe, no
  part del document: els titols no s'escriuen i el text va amb el format de
  sempre (Format.ps1).

.NOTES
  CONVENCIO ASCII: el codi no porta accents. El text accentuat que va a la
  pantalla es fa amb [char]0xNN; el de l'informe surt del JSON.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Ruta del cataleg dels dos informes.
function _MnsCatalegPath {
    return [string](Join-Path $EstructuralsDir 'MNSTRAS.json')
}

# Llegeix MNSTRAS.json. $null si no hi es (el programa ho ha de dir, no fer com
# si res).
function Read-MnsCataleg([string]$path = '') {
    if ([string]::IsNullOrWhiteSpace($path)) { $path = _MnsCatalegPath }
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

# Les dues fases que aporta aquest modul, amb el mateix esquema que _LlicFases.
function _MnsFases {
    return @(
        [pscustomobject]@{
            Clau = 'mns'
            Nom  = 'Modificaci' + [char]0x00F3 + ' NO Substancial'
            Sub  = 'S' + [char]0x2019 + 'informa favorablement una modificaci' + [char]0x00F3 + ' que no es substancial'
            Curt = 'LlicMNS'
        }
        [pscustomobject]@{
            Clau = 'traspas'
            Nom  = 'Trasp' + [char]0x00E0 + 's'
            Sub  = 'Canvi de nom del titular de l' + [char]0x2019 + 'activitat'
            Curt = 'LlicTraspas'
        }
    )
}

# Es una fase d'aquest modul? Funcio PURA.
function _MnsEsFase([string]$fase) {
    foreach ($f in @(_MnsFases)) { if ([string]$f.Clau -eq [string]$fase) { return $true } }
    return $false
}

# La seccio del cataleg d'una fase (la clau del node = la clau de la fase).
function _MnsSeccio($cat, [string]$fase) {
    if ($null -eq $cat) { return $null }
    foreach ($s in @($cat.nodes)) {
        if ([string]$s.clau -eq [string]$fase) { return $s }
    }
    return $null
}

# HI ENTRA, aquest node? Funcio PURA. La clau diu de quina variant es; qualsevol
# altra clau (o cap) vol dir que hi va sempre.
function _MnsNodeEntra([string]$clau, [bool]$ambObservacions) {
    switch (([string]$clau).Trim().ToLower()) {
        'amb-observacions'    { return $ambObservacions }
        'sense-observacions'  { return (-not $ambObservacions) }
        'llista-observacions' { return $ambObservacions }
    }
    return $true
}

# ELS PARAGRAFS d'un informe, en ordre i ja decidits. Funcio PURA: retorna
# @{ Tipus; Linies } amb Tipus 'text' (paragraf normal) o 'llista' (paragraf de
# llista de Word). Es el que escriu el document i tambe el que ensenya la vista
# en Word del cataleg: no hi pot haver dues versions del mateix.
function _MnsParagrafs($cat, [string]$fase, [bool]$ambObservacions) {
    $out = New-Object System.Collections.ArrayList
    $sec = _MnsSeccio $cat $fase
    if ($null -eq $sec) { return $out.ToArray() }
    foreach ($nd in @($sec.fills)) {
        if (-not (_MnsNodeEntra ([string]$nd.clau) $ambObservacions)) { continue }
        $tipus = if ([string]$nd.tipus -eq 'item') { 'llista' } else { 'text' }
        $linies = New-Object System.Collections.ArrayList
        foreach ($p in @($nd.cos)) { [void]$linies.Add((_JsonParaToBodyLine $p)) }
        [void]$out.Add(@{ Tipus = $tipus; Linies = $linies.ToArray() })
    }
    return $out.ToArray()
}

# Nom del fitxer de sortida (mateix patro que la resta: data al principi).
function _MnsNomFitxer([datetime]$data, [string]$fase, [string]$idGia) {
    $curt = 'LlicMNS'
    foreach ($f in @(_MnsFases)) { if ([string]$f.Clau -eq [string]$fase) { $curt = [string]$f.Curt } }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($data.ToString('yyyy-MM-dd'))
    [void]$parts.Add($curt)
    if (-not [string]::IsNullOrWhiteSpace($idGia)) { [void]$parts.Add('GIA ' + $idGia.Trim()) }
    $nom = ($parts -join '_')
    $nom = [regex]::Replace($nom, '[\\/:*?"<>|]', '-')
    return ($nom + '.docx')
}

# ----------------------------------------------------------------------------
# COMPOSICIO DEL DOCUMENT (Word COM)
# ----------------------------------------------------------------------------
function Build-MnsDocument($word, $model) {
    $header = $model.Header
    $baseName = _MnsNomFitxer (Get-Date) ([string]$model.Fase) ([string]$header['ID_GIA'])
    $targetDir = _ResolveOutputDir
    [string]$outPath = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    # LA MATEIXA CAPCALERA que la resta d'informes de Llicencia (porta la linia
    # "Classificacio:"), tal com va demanar l'usuari.
    Select-CapcaleraBlock $doc 'LLIC'
    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)   # wdStory

    $cfg = $Script:ReportFormatConfig
    foreach ($p in @(_MnsParagrafs $model.Cataleg ([string]$model.Fase) ([bool]$model.AmbObservacions))) {
        if ([string]$p.Tipus -eq 'llista') {
            # El paragraf de llista va BUIT: l'omple l'usuari al Word.
            Format-ListItem $sel ''
            continue
        }
        foreach ($l in @($p.Linies)) {
            $pp = _SplitTextAndUrls ([string]$l)
            if (-not [string]::IsNullOrWhiteSpace($pp.Text)) { Format-Body $sel $pp.Text }
            foreach ($u in @($pp.Urls)) { Format-Url $sel $u }
        }
        if ($cfg.SpacerAfterItem) { Format-Spacer $sel }
    }

    # El tancament: del cataleg, com tots els altres informes (Write-Tancament).
    Write-Tancament $sel

    $doc.Save()
    $doc.Close($false)
    try { Move-Item -LiteralPath $tempPath -Destination $outPath -Force } catch { return $tempPath }
    return $outPath
}

# ----------------------------------------------------------------------------
# LA PANTALLA (l'unica part que toca WinForms)
# ----------------------------------------------------------------------------
# Nomes hi ha una cosa per decidir: si hi ha observacions o no. Ensenya, a sota,
# com quedara la frase, que es l'unica manera de triar-ho sense dubtar.
function Select-MnsObservacions([string]$fase, $cat, $pre = $null) {
    $form = _NewForm
    $nom = ''
    foreach ($f in @(_MnsFases)) { if ([string]$f.Clau -eq [string]$fase) { $nom = [string]$f.Nom } }
    $form.Text = 'Llic' + [char]0x00E8 + 'ncia - ' + $nom
    $form.ClientSize = New-Object System.Drawing.Size(620, 330)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 72)
    $lbl.Size = New-Object System.Drawing.Size(570, 20)
    $lbl.Text = 'Hi ha observacions en relaci' + [char]0x00F3 + ' a aquest tr' + [char]0x00E0 + 'mit?'
    [void]$form.Controls.Add($lbl)

    $rbNo = New-Object System.Windows.Forms.RadioButton
    $rbNo.Location = New-Object System.Drawing.Point(30, 98)
    $rbNo.Size = New-Object System.Drawing.Size(560, 22)
    $rbNo.Text = 'No, cap observaci' + [char]0x00F3
    [void]$form.Controls.Add($rbNo)

    $rbSi = New-Object System.Windows.Forms.RadioButton
    $rbSi.Location = New-Object System.Drawing.Point(30, 124)
    $rbSi.Size = New-Object System.Drawing.Size(560, 22)
    $rbSi.Text = 'S' + [char]0x00ED + ', i les escriur' + [char]0x00E9 + ' al Word (hi surt una llista buida)'
    [void]$form.Controls.Add($rbSi)

    if ($null -ne $pre -and [bool]$pre) { $rbSi.Checked = $true } else { $rbNo.Checked = $true }

    $lblPrev = New-Object System.Windows.Forms.Label
    $lblPrev.Location = New-Object System.Drawing.Point(30, 158)
    $lblPrev.Size = New-Object System.Drawing.Size(560, 90)
    $lblPrev.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lblPrev.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    [void]$form.Controls.Add($lblPrev)

    # LES FUNCIONS DE LA PANTALLA, TOTES DINS D'UN HASHTABLE (vegeu CLAUDE.md:
    # .GetNewClosure() copia VALORS, i una closure que se'n cridi una altra es
    # quedaria amb $null).
    $fn = @{}
    $fn.Previsualitza = {
        $amb = [bool]$rbSi.Checked
        $linies = New-Object System.Collections.ArrayList
        foreach ($p in @(_MnsParagrafs $cat $fase $amb)) {
            if ([string]$p.Tipus -eq 'llista') { [void]$linies.Add('   1. ...'); continue }
            foreach ($l in @($p.Linies)) { [void]$linies.Add(((_SplitTextAndUrls ([string]$l)).Text)) }
        }
        # Nomes la frase que canvia: la resta ja la veura al document.
        $clau = if ($amb) { 'amb la seg' } else { 'sense m' }
        $frase = @($linies | Where-Object { ([string]$_) -like ('*' + $clau + '*') })
        $lblPrev.Text = if ($frase.Count -gt 0) { [string]$frase[0] } else { '' }
    }.GetNewClosure()
    # UN CLIC EN UN RADIO DISPARA DOS ESDEVENIMENTS (el que es marca i el germa
    # que es desmarca): un handler per radio i prou.
    $rbSi.add_CheckedChanged({ & $fn.Previsualitza }.GetNewClosure())
    $rbNo.add_CheckedChanged({ & $fn.Previsualitza }.GetNewClosure())
    & $fn.Previsualitza

    $res = @{ Nav = 'back'; AmbObservacions = $false }
    $yBotons = 262
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Generar'
    $btnOk.Location = New-Object System.Drawing.Point(465, $yBotons)
    $btnOk.Size = New-Object System.Drawing.Size(130, 32)
    _StylePrimaryButton $btnOk
    $btnOk.add_Click({
        $res.AmbObservacions = [bool]$rbSi.Checked
        $res.Nav = 'fwd'
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())
    [void]$form.Controls.Add($btnOk)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = [string][char]0x2190 + ' Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(20, $yBotons)
    $btnBack.Size = New-Object System.Drawing.Size(115, 32)
    _StyleSecondaryButton $btnBack
    $btnBack.add_Click({ $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnBack)

    [void](_AddBrandHeader $form $nom ('Nom' + [char]0x00E9 + 's cal decidir si hi ha observacions') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
    return $res
}
