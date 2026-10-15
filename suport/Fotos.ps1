#requires -Version 5.1
<#
  Fotos.ps1 - EL REPORTATGE FOTOGRAFIC dels informes de VISITA D'INSPECCIO.

  L'usuari (octubre 2026): quan l'origen de l'informe es "Visita d'inspeccio",
  despres de les conclusions el programa ha de deixar triar fotos i posar-les a
  l'informe, com ho feia a ma (la mostra: el Req1 del GIA 1415):
    - a les conclusions, "S'adjunta reportatge fotografic." (abans de "Ho poso
      al seu coneixement...");
    - darrere del tancament, PAGINA NOVA amb el titol REPORTATGE FOTOGRAFIC
      (centrat i en negreta, com CONCLUSIONS);
    - les fotos, centrades: les VERTICALS a 6 cm d'ample i, dues verticals
      seguides, DE COSTAT (12 cm entre les dues); les HORITZONTALS a 12 cm (el
      mateix ample que la parella); a sota de cada una, "Foto n".
  Les tria sempre l'usuari (la carpeta d'entrada es la d'informes de la
  Configuracio) i el programa en fa una copia REDUIDA (1600 px pel costat llarg,
  JPEG): a 6-12 cm es veuen igual i l'informe (i el PDF signat que s'envia per
  correu) no fa 30 MB. Els originals no es toquen.

  Les peces PURES (la disposicio, la mida, el gir EXIF i els blocs) es proven
  a Linux. Llegir i reduir les fotos (System.Drawing) i la pantalla nomes
  corren a Windows. El Word ho escriu Format-Fotos (Format.ps1), com tot el
  format.
#>

$Script:FotosFraseConclusio = 'S' + [char]0x2019 + 'adjunta reportatge fotogr' + [char]0x00E0 + 'fic.'
$Script:FotosTitol = 'REPORTATGE FOTOGR' + [char]0x00C0 + 'FIC'
$Script:FotosAmpleVerticalCm = 6
$Script:FotosAmpleHoritzontalCm = 12
$Script:FotosMaxPx = 1600
$Script:FotosQualitatJpeg = 85

# ----------------------------------------------------------------------------
# PECES PURES
# ----------------------------------------------------------------------------
# L'orientacio EXIF (etiqueta 0x0112) d'una foto del mobil: el gir que cal per
# veure-la dreta (un nom de [System.Drawing.RotateFlipType], '' si cap) i si
# en girar-la s'intercanvien l'ample i l'alt (5-8: una foto vertical que el
# mobil desa "estirada"). PURES.
function _FotoGirExif([int]$o) {
    switch ($o) {
        2 { return 'RotateNoneFlipX' }
        3 { return 'Rotate180FlipNone' }
        4 { return 'Rotate180FlipX' }
        5 { return 'Rotate90FlipX' }
        6 { return 'Rotate90FlipNone' }
        7 { return 'Rotate270FlipX' }
        8 { return 'Rotate270FlipNone' }
    }
    return ''
}
function _FotoGiraEixos([int]$o) { return ($o -ge 5 -and $o -le 8) }

# La mida reduida: el costat llarg a $max com a molt, sense deformar. PURA.
function _FotoMidaReduida([int]$w, [int]$h, [int]$max) {
    if ($w -le 0 -or $h -le 0) { return @($w, $h) }
    $m = [Math]::Max($w, $h)
    if ($m -le $max) { return @($w, $h) }
    $f = $max / $m
    return @([int][Math]::Max(1, [Math]::Round($w * $f)), [int][Math]::Max(1, [Math]::Round($h * $f)))
}

# LA DISPOSICIO: en l'ordre que les ha deixat l'usuari, dues verticals SEGUIDES
# van de costat; la resta, una per fila. $fotos: @{ Path; Vertical }. Torna les
# files: @{ Fotos = @( @{ Path; Num; AmpleCm } ) }. Nomes es fa parella amb la
# SEGUENT (no es reordena res: "Foto n" ha de ser l'ordre que l'usuari ha
# triat, que es el que citara). PURA.
function _FotosFiles($fotos) {
    $llista = @(@($fotos) | Where-Object { $null -ne $_ })
    $files = New-Object System.Collections.ArrayList
    $i = 0
    while ($i -lt $llista.Count) {
        $f = $llista[$i]
        if ([bool]$f.Vertical -and ($i + 1) -lt $llista.Count -and [bool]$llista[$i + 1].Vertical) {
            [void]$files.Add(@{ Fotos = @(
                @{ Path = [string]$f.Path; Num = ($i + 1); AmpleCm = $Script:FotosAmpleVerticalCm },
                @{ Path = [string]$llista[$i + 1].Path; Num = ($i + 2); AmpleCm = $Script:FotosAmpleVerticalCm }) })
            $i += 2
            continue
        }
        $ample = if ([bool]$f.Vertical) { $Script:FotosAmpleVerticalCm } else { $Script:FotosAmpleHoritzontalCm }
        [void]$files.Add(@{ Fotos = @(@{ Path = [string]$f.Path; Num = ($i + 1); AmpleCm = $ample }) })
        $i++
    }
    return $files.ToArray()
}

# Els BLOCS del reportatge (Write-Informe): pagina nova, el titol i una fila de
# fotos per bloc. Sense fotos, cap. PURA.
function _FotosBlocs($fotos) {
    $files = @(_FotosFiles $fotos)
    if ($files.Count -eq 0) { return @() }
    $b = New-Object System.Collections.ArrayList
    [void]$b.Add(@{ T = 'saltpagina' })
    [void]$b.Add(@{ T = 'conclusiocap'; Text = $Script:FotosTitol })
    foreach ($f in $files) { [void]$b.Add(@{ T = 'fotos'; Fila = $f }) }
    return $b.ToArray()
}

# L'informe es d'una VISITA D'INSPECCIO (la tria "Origen de l'informe" del Pas
# 2, ORIGEN_TIPUS = 'insp'): nomes llavors hi ha el pas de les fotos. PURA.
function _EsVisitaInspeccio($header) {
    if ($null -eq $header) { return $false }
    $v = if ($header -is [System.Collections.IDictionary]) { $header['ORIGEN_TIPUS'] } elseif ($null -ne $header.PSObject.Properties['ORIGEN_TIPUS']) { $header.ORIGEN_TIPUS } else { '' }
    return ([string]$v -eq 'insp')
}

# Les conclusions de l'informe, amb la frase del reportatge al final (abans de
# les de sempre, "Ho poso al seu coneixement..."), si n'hi ha. PURA.
function _ConclusionsAmbFotos($conclusions, $fotos) {
    $c = @(@($conclusions) | Where-Object { $null -ne $_ })
    if (@(@($fotos) | Where-Object { $null -ne $_ }).Count -eq 0) { return $c }
    return @($c + @($Script:FotosFraseConclusio))
}

# La carpeta on s'obre el dialeg: la que s'ha fet servir l'ultima vegada (en
# aquesta tria) o, si no, la d'informes de la Configuracio (l'usuari). PURA.
function _FotosCarpetaInicial([string]$ultima, [string]$informes) {
    if (-not [string]::IsNullOrWhiteSpace($ultima)) { return $ultima }
    return [string]$informes
}

# ----------------------------------------------------------------------------
# REDUIR LES FOTOS (nomes Windows: System.Drawing)
# ----------------------------------------------------------------------------
# UNA foto: la llegeix (d'un flux en memoria: el fitxer de la xarxa no queda
# agafat), la gira segons l'EXIF, la redueix i la desa JPEG a $dir. Torna
# @{ Path; Vertical; Original }. El bitmap NOU no porta l'etiqueta d'orientacio:
# si la portes, el Word la tornaria a girar. LLANCA si no la pot llegir (un HEIC
# de l'iPhone, per exemple).
function _FotoPreparaUna([string]$path, [string]$dir, [int]$num) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $ms = New-Object System.IO.MemoryStream (, $bytes)
    $img = $null; $bmp = $null; $g = $null
    try {
        $img = [System.Drawing.Image]::FromStream($ms)
        $o = 1
        if (@($img.PropertyIdList) -contains 274) {
            try { $o = [int][BitConverter]::ToUInt16($img.GetPropertyItem(274).Value, 0) } catch { $o = 1 }
        }
        $gir = _FotoGirExif $o
        if ($gir -ne '') { $img.RotateFlip([System.Drawing.RotateFlipType]$gir) }
        $mida = _FotoMidaReduida $img.Width $img.Height $Script:FotosMaxPx
        $bmp = New-Object System.Drawing.Bitmap ([int]$mida[0]), ([int]$mida[1])
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($img, 0, 0, [int]$mida[0], [int]$mida[1])
        $codec = @([System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' })[0]
        $params = New-Object System.Drawing.Imaging.EncoderParameters 1
        $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), ([long]$Script:FotosQualitatJpeg)
        $desti = Join-Path $dir ('foto-' + $num.ToString('000') + '.jpg')
        $bmp.Save($desti, $codec, $params)
        return @{ Path = [string]$desti; Vertical = ($mida[1] -gt $mida[0]); Original = $path }
    } finally {
        if ($null -ne $g) { $g.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        $ms.Dispose()
    }
}

# TOTES: les copies reduides a una carpeta temporal propia (qui crida l'esborra
# quan l'informe ja es desat: el Word les incrusta). Torna @{ Fotos; Dir;
# Errors } (les que no s'han pogut llegir, amb el motiu).
# System.Drawing es carrega AQUI i no dins de _FotoPreparaUna: un tipus dins
# d'una funcio es resol en compilar-ne el cos (vegeu CLAUDE.md, regla 11).
function Convert-FotosInforme($paths) {
    Add-Type -AssemblyName System.Drawing
    $dir = Join-Path ([System.IO.Path]::GetTempPath()) ('informes-fotos-' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $dir -Force)
    $fotos = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($p in @($paths)) {
        $i++
        try { [void]$fotos.Add((_FotoPreparaUna ([string]$p) $dir $i)) }
        catch { [void]$errors.Add(([System.IO.Path]::GetFileName([string]$p) + ': ' + [string]$_.Exception.Message)) }
    }
    return @{ Fotos = $fotos.ToArray(); Dir = [string]$dir; Errors = $errors.ToArray() }
}

# ----------------------------------------------------------------------------
# LA PANTALLA (Pas 5 de l'assistent, nomes amb "Visita d'inspeccio")
# ----------------------------------------------------------------------------
# Els menys passos possibles (l'usuari): en obrir-se, ja surt el dialeg per
# triar-ne (es poden triar totes de cop, amb Ctrl o Maj); si no se'n vol cap,
# Cancel.lar i "Generar informe". La llista es pot ordenar i se'n poden treure
# o afegir mes. $prev: les que ja s'havien triat (tornant d'Enrere). Torna
# @{ Nav = 'back' | 'next'; Fotos = @(rutes) }.
function Select-FotosInforme([string]$dirInicial, $prev = @()) {
    $st = @{ Fotos = (New-Object System.Collections.ArrayList); Ultima = '' }
    foreach ($p in @($prev)) { if (-not [string]::IsNullOrWhiteSpace([string]$p)) { [void]$st.Fotos.Add([string]$p) } }

    $form = _NewForm
    $form.Text = 'Reportatge fotogr' + [char]0x00E0 + 'fic'
    $form.Size = New-Object System.Drawing.Size(820, 600)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 480)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 102)
    $lbl.Size = New-Object System.Drawing.Size(760, 40)
    $lbl.Anchor = 'Top, Left, Right'
    $lbl.Text = "Les fotos van a l'informe en aquest ordre (Foto 1, Foto 2...). Dues verticals seguides surten de costat; les horitzontals, a l'ample de les dues. Sense cap foto, l'informe surt sense reportatge."
    [void]$form.Controls.Add($lbl)

    $llista = New-Object System.Windows.Forms.ListBox
    $llista.Location = New-Object System.Drawing.Point(20, 146)
    $llista.Size = New-Object System.Drawing.Size(620, 330)
    $llista.Anchor = 'Top, Bottom, Left, Right'
    $llista.HorizontalScrollbar = $true
    $llista.SelectionMode = 'MultiExtended'
    [void]$form.Controls.Add($llista)

    $omple = {
        param([int]$tria = -1)
        $llista.BeginUpdate()
        $llista.Items.Clear()
        for ($k = 0; $k -lt $st.Fotos.Count; $k++) {
            $p = [string]$st.Fotos[$k]
            [void]$llista.Items.Add(('Foto ' + ($k + 1) + '   ' + [System.IO.Path]::GetFileName($p) + '   (' + [System.IO.Path]::GetDirectoryName($p) + ')'))
        }
        $llista.EndUpdate()
        if ($tria -ge 0 -and $tria -lt $llista.Items.Count) { $llista.SelectedIndex = $tria }
    }.GetNewClosure()

    $afegir = {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Tria les fotos del reportatge (totes de cop: Ctrl o Maj)'
        $dlg.Filter = 'Fotos (*.jpg; *.jpeg; *.png)|*.jpg;*.jpeg;*.png|Tots els fitxers (*.*)|*.*'
        $dlg.Multiselect = $true
        $ini = _FotosCarpetaInicial $st.Ultima $dirInicial
        if (-not [string]::IsNullOrWhiteSpace($ini) -and (Test-Path -LiteralPath $ini)) { $dlg.InitialDirectory = $ini }
        if ($dlg.ShowDialog($form) -ne 'OK') { return }
        # Pel NOM: les del mobil (IMG_AAAAMMDD_HHMMSS) queden per ordre d'hora.
        foreach ($f in @($dlg.FileNames | Sort-Object)) { if (-not $st.Fotos.Contains([string]$f)) { [void]$st.Fotos.Add([string]$f) } }
        if ($dlg.FileNames.Count -gt 0) { $st.Ultima = [System.IO.Path]::GetDirectoryName([string]$dlg.FileNames[0]) }
        & $omple ($st.Fotos.Count - 1)
    }.GetNewClosure()

    $mou = {
        param([int]$d)
        $i = $llista.SelectedIndex
        $j = $i + $d
        if ($i -lt 0 -or $j -lt 0 -or $j -ge $st.Fotos.Count) { return }
        $x = $st.Fotos[$i]; $st.Fotos[$i] = $st.Fotos[$j]; $st.Fotos[$j] = $x
        & $omple $j
    }.GetNewClosure()

    $mkBoto = {
        param([string]$text, [int]$y, $clic)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        $b.Location = New-Object System.Drawing.Point(656, $y)
        $b.Size = New-Object System.Drawing.Size(140, 32)
        $b.Anchor = 'Top, Right'
        _StyleSecondaryButton $b
        $b.add_Click($clic)
        [void]$form.Controls.Add($b)
        return $b
    }.GetNewClosure()
    [void](& $mkBoto ('Afegir fotos' + [char]0x2026) 146 $afegir)
    [void](& $mkBoto 'Treure' 186 {
        $sel = @($llista.SelectedIndices | ForEach-Object { [int]$_ } | Sort-Object -Descending)
        foreach ($i in $sel) { $st.Fotos.RemoveAt($i) }
        & $omple ([Math]::Min($(if ($sel.Count -gt 0) { $sel[-1] } else { 0 }), $st.Fotos.Count - 1))
    }.GetNewClosure())
    [void](& $mkBoto ([string][char]0x2191 + ' Amunt') 226 { & $mou (-1) }.GetNewClosure())
    [void](& $mkBoto ([string][char]0x2193 + ' Avall') 266 { & $mou 1 }.GetNewClosure())

    [void](_AddPeuBotons $form @(@{ Nom = 'Enrere'; Text = (_TxtEnrere); Resultat = 'Retry' }) @(
        @{ Nom = 'Ok'; Text = ('Generar informe ' + [char]0x2192); Estil = 'primari'; Resultat = 'OK'; Intro = $true }) 500 -Ancorat)
    [void](_AddStepBar $form 5)
    [void](_AddBrandHeader $form ('Reportatge fotogr' + [char]0x00E0 + 'fic') "Visita d'inspecci$([char]0x00F3): tria les fotos que vols adjuntar a l'informe")

    & $omple 0
    # Un pas menys: si encara no n'hi ha cap, el dialeg s'obre sol.
    $form.add_Shown({ if ($st.Fotos.Count -eq 0) { & $afegir } }.GetNewClosure())

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav = 'back'; Fotos = @($st.Fotos) } }
    if ($res -ne 'OK') { exit 0 }
    return [pscustomobject]@{ Nav = 'next'; Fotos = @($st.Fotos) }
}
