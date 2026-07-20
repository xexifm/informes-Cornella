#requires -Version 5.1
<#
.SYNOPSIS
  Escaner d'informes ja generats -> base de dades JSON.

.DESCRIPTION
  Recorre l'arbre de carpetes dels informes (per defecte $InformesDir, germa de
  la carpeta de l'Excel d'activitats) i, per cada informe (.docx), en treu:
    - la DATA (del principi del nom del fitxer),
    - l'ID GIA (del document; si no hi es, del nom de la carpeta "GIA 361"; si
      tampoc, de l'Excel d'activitats cercant per numero d'expedient),
    - la CONCLUSIO (el paragraf que comenca amb "Vist l'anterior").
  Ho desa AGRUPAT PER ACTIVITAT a BASE DE DADES ACTIVITATS\informes-db.json
  (carpeta ignorada per git). Els informes que no es poden resoldre del tot van
  a un bloc "a_revisar" per poder-los repassar.

  Es un modul del motor: es carrega (dot-source) des de GenerarInforme.ps1, aixi
  reutilitza les funcions de lectura de .docx sense Word (de Seguiment.ps1),
  _NormalizeText i l'acces a l'Excel d'activitats (Find-LatestActivitatsExcel /
  Initialize-ActivitatsCache). Les funcions de logica de text son PURES (operen
  sobre cadenes) perque es puguin provar en headless (Linux, sense Word).

.NOTES
  Es llanca des del menu (Pas 1) amb el boto "Actualitzar base d'informes".
#>

# ----------------------------------------------------------------------------
# Logica de text (funcions PURES, testejables en headless)
# ----------------------------------------------------------------------------

# Treu la data del PRINCIPI del nom del fitxer. Tolerant amb el format:
#   AAAA-MM-DD / AAAAMMDD / AAAA.MM.DD / AAAA_MM_DD  (any de 4 xifres)
#   AA-MM-DD   / AA.MM.DD / AA_MM_DD                 (any de 2 xifres, amb
#                                                     separador obligatori per no
#                                                     confondre-ho amb AAAAMMDD)
# Retorna 'yyyy-MM-dd' o $null si el nom no comenca amb una data valida.
# (El programa considera "informe" NOMES els fitxers que comencen amb data.)
function _ParseDataInformeFromName($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    $n = [string]$name
    # Any de 4 xifres, separadors opcionals.
    if ($n -match '^(\d{4})[-_.]?(\d{2})[-_.]?(\d{2})(?:\D|$)') {
        $y = [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($y -ge 1990 -and $y -le 2100 -and $mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    # Any de 2 xifres, separador OBLIGATORI.
    if ($n -match '^(\d{2})[-_.](\d{2})[-_.](\d{2})(?:\D|$)') {
        $y = 2000 + [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    return $null
}

# ID GIA d'una llista de linies de text del document. Busca la linia que conte
# "ID GIA" i en retorna el valor (despres dels dos punts / espais). '' si no hi es.
function _ExtractIdGia($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        if ($s -imatch 'ID\s*GIA\s*:?\s*(.+)$') {
            $val = $Matches[1].Trim()
            # Ens quedem nomes amb el primer "token" del valor (l'ID; evita
            # arrossegar text si la linia porta res mes al darrere).
            if ($val -match '^([\w./-]+)') { return $Matches[1] }
            return $val
        }
    }
    return ''
}

# Numero d'expedient d'una llista de linies. Busca la linia que comenca per
# "Exp" (p.ex. "Exp. Num: 2025/1/2563") i retorna el valor despres dels dos punts.
function _ExtractExpedient($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        $nt = _NormalizeText $s
        if ($nt -match '^exp') {
            $idx = $s.IndexOf(':')
            if ($idx -ge 0) { return $s.Substring($idx + 1).Trim() }
            if ($s -match '(\S+/\S+/\S+)') { return $Matches[1] }
        }
    }
    return ''
}

# Normalitzacio per comparar frases: sense accents, minuscules i SENSE
# apostrofs. Els informes fan servir l'apostrof TIPOGRAFIC (U+2019), pero les
# nostres frases de referencia el recte (U+0027); traient-los tots dos (i altres
# variants) la comparacio casa igual. Fem servir codepoints [char]0x.... per no
# dependre de l'encoding amb que PowerShell 5.1 llegeix aquest fitxer.
function _ConclNorm($s) {
    $t = _NormalizeText $s
    $apos = @([char]0x0027, [char]0x2018, [char]0x2019, [char]0x02BC, [char]0x00B4, [char]0x0060)
    foreach ($a in $apos) { $t = $t.Replace([string]$a, '') }
    return $t
}

# Conclusio: des del paragraf que conte "Vist l'anterior" fins (exclos) el que
# conte "Ho poso al seu coneixement" o "Cornella de Llobregat,". Uneix els
# paragrafs amb un espai i retorna el text ORIGINAL (no normalitzat). '' si no
# es troba l'inici.
function _ExtractConclusio($lines) {
    $startPhrase = _ConclNorm "Vist l'anterior"
    $endPhrases  = @((_ConclNorm 'Ho poso al seu coneixement'), (_ConclNorm 'Cornella de Llobregat,'))
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ((_ConclNorm $lines[$i]).Contains($startPhrase)) { $start = $i; break }
    }
    if ($start -lt 0) { return '' }
    $parts = New-Object System.Collections.ArrayList
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $ln = [string]$lines[$i]
        $nl = _ConclNorm $ln
        $isEnd = $false
        foreach ($ep in $endPhrases) { if ($nl.Contains($ep)) { $isEnd = $true; break } }
        if ($isEnd) { break }
        if (-not [string]::IsNullOrWhiteSpace($ln)) { [void]$parts.Add($ln.Trim()) }
    }
    return ($parts -join ' ')
}

# ID GIA a partir dels noms de les carpetes pare (p.ex. la carpeta de l'activitat
# "2025-1-2563 GIA 361 - ... KRICHI ..."). Retorna '' si cap carpeta no en porta.
# Parteix la ruta manualment per '\' o '/' (aixi funciona igual a Windows i a
# Linux, i es pot provar en headless).
function _GiaFromFolderName($path) {
    $segs = ([string]$path) -split '[\\/]'
    # L'ultim segment es el nom del fitxer; recorrem les carpetes de dins enfora.
    for ($i = $segs.Count - 2; $i -ge 0; $i--) {
        if ($segs[$i] -match 'GIA\s*(\d+)') { return $Matches[1] }
    }
    return ''
}

# Nom de la carpeta (activitat) on viu l'informe: la carpeta pare immediata.
function _CarpetaActivitat($path) {
    $segs = ([string]$path) -split '[\\/]' | Where-Object { $_ -ne '' }
    if ($segs.Count -ge 2) { return $segs[$segs.Count - 2] }
    return ''
}

# Normalitza un numero d'expedient per comparar-lo (l'informe fa servir "/", la
# carpeta "-", i l'Excel pot portar zeros al davant): parteix en grups i treu els
# zeros inicials de cada grup numeric. "2025/1/2563" i "2025/01/2563" -> "2025-1-2563".
function _NormalitzaExpedient($s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $groups = ([string]$s).Trim() -split '[^\dA-Za-z]+' | Where-Object { $_ -ne '' }
    $norm = $groups | ForEach-Object {
        if ($_ -match '^\d+$') { [string][int]$_ } else { $_.ToUpper() }
    }
    return ($norm -join '-')
}

# Construeix un mapa expedient_normalitzat -> ID GIA a partir de la cache de
# l'Excel d'activitats ($cache.ById[id] = @{ EXP_NUM; TITULAR; ... }).
function Build-ExpedientToGiaMap($cache) {
    $map = @{}
    if ($null -eq $cache -or $null -eq $cache.ById) { return $map }
    foreach ($kv in $cache.ById.GetEnumerator()) {
        $exp = _NormalitzaExpedient $kv.Value.EXP_NUM
        if ($exp -ne '' -and -not $map.ContainsKey($exp)) { $map[$exp] = [string]$kv.Key }
    }
    return $map
}

# ----------------------------------------------------------------------------
# Lectura del .docx (necessita les primitives de Seguiment.ps1; no headless)
# ----------------------------------------------------------------------------
# Retorna un array amb el text de TOTS els paragrafs del document, INCLOENT els
# de dins de taules (la capcalera amb ID GIA / Exp. Num viu en una taula, i
# _BodyParagraphsXml les salta; per aixo seleccionem './/w:p').
function _ReadDocxParagraphs($docxPath) {
    $info = _LoadDocxXml $docxPath
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $info.Body.SelectNodes('.//w:p', $info.Ns)) {
        [void]$out.Add((_ParagraphTextXml $p $info.Ns))
    }
    return $out.ToArray()
}

# Analitza UN informe. Retorna un PSCustomObject amb data, gia, expedient,
# conclusio, fitxer, ruta, carpeta i el motiu (si cal revisar-lo).
function Get-InformeData($file, $expToGia, $cache) {
    $data = _ParseDataInformeFromName $file.Name
    $lines = @()
    try { $lines = _ReadDocxParagraphs $file.FullName } catch { $lines = @() }

    $gia = _ExtractIdGia $lines
    $exp = _ExtractExpedient $lines
    $font = 'document'
    if ([string]::IsNullOrWhiteSpace($gia)) {
        $gia = _GiaFromFolderName $file.FullName
        if (-not [string]::IsNullOrWhiteSpace($gia)) { $font = 'carpeta' }
    }
    if ([string]::IsNullOrWhiteSpace($gia) -and $null -ne $expToGia) {
        $key = _NormalitzaExpedient $exp
        if ($key -ne '' -and $expToGia.ContainsKey($key)) { $gia = $expToGia[$key]; $font = 'excel' }
    }

    $concl = _ExtractConclusio $lines

    $motius = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($gia))   { [void]$motius.Add('sense ID GIA') }
    if ([string]::IsNullOrWhiteSpace($concl)) { [void]$motius.Add('sense conclusio') }

    $titular = ''
    if ($null -ne $cache -and -not [string]::IsNullOrWhiteSpace($gia) -and $cache.ById.ContainsKey([string]$gia)) {
        $titular = [string]$cache.ById[[string]$gia].TITULAR
    }

    return [pscustomobject]@{
        Data       = $data
        Gia        = $gia
        GiaFont    = $font
        Expedient  = $exp
        Titular    = $titular
        Conclusio  = $concl
        Fitxer     = $file.Name
        Ruta       = $file.FullName
        Carpeta    = _CarpetaActivitat $file.FullName
        Motius     = $motius.ToArray()
    }
}

# ----------------------------------------------------------------------------
# Escaneig complet + escriptura del JSON (interactiu, amb finestra de progres)
# ----------------------------------------------------------------------------
function Invoke-InformesDbScan {
    # 1. Resoldre la carpeta d'informes.
    $dir = $InformesDir
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la carpeta d'informes:`n$dir`n`nConfigura `$InformesDir a config.ps1 si la tens en una altra ubicacio.",
            'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }

    # 2. Finestra de progres.
    $form = _NewForm
    $form.Text = "Actualitzant base d'informes"
    $form.Size = New-Object System.Drawing.Size(560, 170)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(510, 60)
    $lbl.Text = "Cercant informes a:`n$dir"
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 90)
    $bar.Size = New-Object System.Drawing.Size(510, 24)
    $bar.Style = 'Marquee'
    $form.Controls.Add($bar)
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # 3. Carregar l'Excel d'activitats (opcional; per la cerca inversa i el
        #    titular). Si no hi ha Excel, es continua sense aquest fallback.
        $cache = $null; $expToGia = $null
        try {
            $excel = Find-LatestActivitatsExcel
            if ($null -ne $excel) {
                $lbl.Text = "Llegint la base d'activitats (Excel)..."
                [System.Windows.Forms.Application]::DoEvents()
                $cache = Initialize-ActivitatsCache $excel.File
                $expToGia = Build-ExpedientToGiaMap $cache
            }
        } catch { $cache = $null; $expToGia = $null }

        # 4. Recollir els fitxers candidats (.docx amb data al principi del nom).
        $lbl.Text = "Cercant informes a:`n$dir"
        [System.Windows.Forms.Application]::DoEvents()
        $allDocx = Get-ChildItem -LiteralPath $dir -Recurse -File -Filter '*.docx' -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -notlike '~$*' -and $null -ne (_ParseDataInformeFromName $_.Name) }
        $files = @($allDocx)
        $total = $files.Count

        $bar.Style = 'Continuous'
        $bar.Minimum = 0
        $bar.Maximum = [Math]::Max(1, $total)

        # 5. Analitzar cada informe.
        $informes = New-Object System.Collections.ArrayList
        $revisar  = New-Object System.Collections.ArrayList
        $i = 0
        foreach ($f in $files) {
            $i++
            if (($i % 5) -eq 0 -or $i -eq $total) {
                $lbl.Text = "Analitzant informes... ($i de $total)"
                $bar.Value = [Math]::Min($bar.Maximum, $i)
                [System.Windows.Forms.Application]::DoEvents()
            }
            $r = Get-InformeData $f $expToGia $cache
            [void]$informes.Add($r)
            if ($r.Motius.Count -gt 0) {
                [void]$revisar.Add([pscustomobject]@{
                    fitxer = $r.Fitxer
                    ruta   = $r.Ruta
                    motiu  = ($r.Motius -join ', ')
                })
            }
        }

        # 6. Agrupar per activitat (clau = ID GIA; si falta, EXP:<expedient>; si
        #    tampoc, la carpeta). Ordenem els informes de cada activitat per data.
        $groups = [ordered]@{}
        foreach ($r in $informes) {
            $key = if (-not [string]::IsNullOrWhiteSpace($r.Gia)) { "GIA:$($r.Gia)" }
                   elseif (-not [string]::IsNullOrWhiteSpace($r.Expedient)) { "EXP:$(_NormalitzaExpedient $r.Expedient)" }
                   else { "DIR:$($r.Carpeta)" }
            if (-not $groups.Contains($key)) {
                $groups[$key] = [pscustomobject]@{
                    id_gia    = $r.Gia
                    expedient = $r.Expedient
                    titular   = $r.Titular
                    carpeta   = $r.Carpeta
                    _informes = (New-Object System.Collections.ArrayList)
                }
            }
            $g = $groups[$key]
            # Emplenem camps de l'activitat si encara estan buits.
            if ([string]::IsNullOrWhiteSpace($g.id_gia)    -and -not [string]::IsNullOrWhiteSpace($r.Gia))       { $g.id_gia = $r.Gia }
            if ([string]::IsNullOrWhiteSpace($g.expedient) -and -not [string]::IsNullOrWhiteSpace($r.Expedient)) { $g.expedient = $r.Expedient }
            if ([string]::IsNullOrWhiteSpace($g.titular)   -and -not [string]::IsNullOrWhiteSpace($r.Titular))   { $g.titular = $r.Titular }
            [void]$g._informes.Add([pscustomobject]@{
                data      = $r.Data
                fitxer    = $r.Fitxer
                conclusio = $r.Conclusio
            })
        }

        $activitats = New-Object System.Collections.ArrayList
        foreach ($g in $groups.Values) {
            $ordered = @($g._informes | Sort-Object { if ($_.data) { $_.data } else { '' } })
            [void]$activitats.Add([pscustomobject]@{
                id_gia    = $g.id_gia
                expedient = $g.expedient
                titular   = $g.titular
                carpeta   = $g.carpeta
                informes  = $ordered
            })
        }
        $activitatsOrd = @($activitats | Sort-Object { [string]$_.id_gia })

        # 7. Escriure el JSON.
        Ensure-AppDataDir | Out-Null
        $outObj = [pscustomobject]@{
            generat_el    = (Get-Date).ToString('o')
            carpeta_arrel = $dir
            n_informes    = $informes.Count
            n_activitats  = $activitatsOrd.Count
            activitats    = $activitatsOrd
            a_revisar     = @($revisar)
        }
        $outPath = Join-Path $LocalActivitatsDir 'informes-db.json'
        if (-not (Test-Path -LiteralPath $LocalActivitatsDir)) {
            New-Item -ItemType Directory -Path $LocalActivitatsDir -Force | Out-Null
        }
        ($outObj | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $outPath -Encoding UTF8

        $form.Close()

        # 8. Resum.
        $msg = "Base d'informes actualitzada.`n`n" +
               "Informes trobats: $($informes.Count)`n" +
               "Activitats: $($activitatsOrd.Count)`n" +
               "A revisar: $($revisar.Count)`n`n" +
               "Fitxer:`n$outPath`n`nVols obrir-lo?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Base d''informes', 'YesNo', 'Information')
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$outPath`"" | Out-Null } catch { }
        }
    }
    catch {
        try { $form.Close() } catch { }
        [System.Windows.Forms.MessageBox]::Show(
            "Error escanejant els informes:`n$($_.Exception.Message)",
            'Base d''informes', 'OK', 'Error') | Out-Null
    }
}
