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
        Modificat  = $file.LastWriteTimeUtc.ToString('o')
        Ignorat    = $false
        Motius     = $motius.ToArray()
    }
}

# Decideix (funcio PURA) si un informe s'ha de tornar a parsejar (obrir el .docx)
# o si es pot reutilitzar l'entrada de l'escaneig anterior:
#   - Si NO en teniem entrada -> cal parsejar (es nou).
#   - Si en teniem i el fitxer s'ha modificat DESPRES de l'ultima actualitzacio
#     -> cal parsejar.
#   - Si en teniem i no s'ha tocat des de l'ultima actualitzacio -> reutilitzar.
# $lwUtc i $prevUtc son [datetime] en UTC.
function _HaDeReprocessar([datetime]$lwUtc, [datetime]$prevUtc, [bool]$teEntrada) {
    if (-not $teEntrada) { return $true }
    return ($lwUtc -gt $prevUtc)
}

# Aplana la base d'informes carregada (objecte de ConvertFrom-Json) en registres
# plans indexats per 'ruta', arrossegant les dades de l'activitat a cada informe.
# Cada registre te la mateixa forma que Get-InformeData (perque el reagrupament
# els tracti igual). Retorna una hashtable [ruta] -> registre.
function _FlattenInformesDb($db) {
    $map = @{}
    if ($null -eq $db -or $null -eq $db.activitats) { return $map }
    foreach ($act in $db.activitats) {
        if ($null -eq $act.informes) { continue }
        foreach ($inf in $act.informes) {
            $ruta = [string]$inf.ruta
            if ([string]::IsNullOrWhiteSpace($ruta)) { continue }
            $motiuStr = if ($null -ne $inf.PSObject.Properties['motiu']) { [string]$inf.motiu } else { '' }
            $motius = if ([string]::IsNullOrWhiteSpace($motiuStr)) { @() } else { @($motiuStr -split ',\s*') }
            $ign = $false
            if ($null -ne $inf.PSObject.Properties['ignorat']) { $ign = [bool]$inf.ignorat }
            $map[$ruta] = [pscustomobject]@{
                Data      = [string]$inf.data
                Gia       = [string]$act.id_gia
                GiaFont   = ''
                Expedient = [string]$act.expedient
                Titular   = [string]$act.titular
                Conclusio = [string]$inf.conclusio
                Fitxer    = [string]$inf.fitxer
                Ruta      = $ruta
                Carpeta   = [string]$act.carpeta
                Modificat = if ($null -ne $inf.PSObject.Properties['modificat']) { [string]$inf.modificat } else { '' }
                Ignorat   = $ign
                Motius    = $motius
            }
        }
    }
    return $map
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

        # 3b. Carregar la base anterior (si existeix) per fer un escaneig
        #     INCREMENTAL: nomes es reobren els .docx modificats DESPRES de
        #     l'ultima actualitzacio; la resta es reutilitzen (conservant el seu
        #     "ignorat"). Els fitxers que ja no existeixen es podaran sols (nomes
        #     reagrupem els que trobem ara). Si no hi ha base previa (o esta
        #     corrupta), es fa un escaneig complet.
        $outPath    = Join-Path $LocalActivitatsDir 'informes-db.json'
        $prevByRuta = @{}
        $prevUtc    = [datetime]::MinValue
        $generatEl  = (Get-Date).ToString('o')
        if (Test-Path -LiteralPath $outPath) {
            try {
                $prevDb     = (Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json)
                $prevByRuta = _FlattenInformesDb $prevDb
                if ($prevDb.PSObject.Properties['actualitzat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.actualitzat_el)) {
                    try { $prevUtc = ([datetime]::Parse([string]$prevDb.actualitzat_el)).ToUniversalTime() } catch { $prevUtc = [datetime]::MinValue }
                }
                if ($prevDb.PSObject.Properties['generat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.generat_el)) {
                    $generatEl = [string]$prevDb.generat_el
                }
            } catch { $prevByRuta = @{}; $prevUtc = [datetime]::MinValue }
        }

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

        # 5. Analitzar cada informe (incremental: reutilitzem els no modificats).
        $informes = New-Object System.Collections.ArrayList
        $revisar  = New-Object System.Collections.ArrayList
        $reprocessats = 0
        $i = 0
        foreach ($f in $files) {
            $i++
            $ruta = $f.FullName
            $teEntrada = $prevByRuta.ContainsKey($ruta)
            if (-not (_HaDeReprocessar $f.LastWriteTimeUtc $prevUtc $teEntrada)) {
                # No s'ha tocat des de l'ultim escaneig: reutilitzem l'entrada.
                $r = $prevByRuta[$ruta]
            } else {
                $r = Get-InformeData $f $expToGia $cache
                # Conservem l'"ignorat" que l'usuari hagi marcat abans.
                if ($teEntrada) { $r.Ignorat = [bool]$prevByRuta[$ruta].Ignorat }
                $reprocessats++
            }
            if (($i % 5) -eq 0 -or $i -eq $total) {
                $lbl.Text = "Analitzant informes... ($i de $total, $reprocessats de nous/modificats)"
                $bar.Value = [Math]::Min($bar.Maximum, $i)
                [System.Windows.Forms.Application]::DoEvents()
            }
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
                ruta      = $r.Ruta
                conclusio = $r.Conclusio
                modificat = $r.Modificat
                ignorat   = [bool]$r.Ignorat
                motiu     = ($r.Motius -join ', ')
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

        # 7. Escriure el JSON (conservem generat_el; actualitzat_el = ara).
        $outObj = [pscustomobject]@{
            generat_el     = $generatEl
            actualitzat_el = (Get-Date).ToString('o')
            carpeta_arrel  = $dir
            n_informes     = $informes.Count
            n_activitats   = $activitatsOrd.Count
            activitats     = $activitatsOrd
            a_revisar      = @($revisar)
        }
        if (-not (Test-Path -LiteralPath $LocalActivitatsDir)) {
            New-Item -ItemType Directory -Path $LocalActivitatsDir -Force | Out-Null
        }
        ($outObj | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $outPath -Encoding UTF8

        $form.Close()

        # 8. Resum.
        $msg = "Base d'informes actualitzada.`n`n" +
               "Informes trobats: $($informes.Count)`n" +
               "Nous o modificats (reprocessats): $reprocessats`n" +
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

# ----------------------------------------------------------------------------
# Editor de la base d'informes (finestra amb taula)
# ----------------------------------------------------------------------------
# Estil d'una fila segons si l'informe esta ignorat: gris + tatxat si ho esta.
function _StyleInformeRow($gridRow, [bool]$ignorat, $fontNormal, $fontStrike) {
    if ($ignorat) {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        $gridRow.DefaultCellStyle.Font      = $fontStrike
    } else {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        $gridRow.DefaultCellStyle.Font      = $fontNormal
    }
}

# Obre una finestra amb la base d'informes en forma de taula: es pot veure cada
# informe (data, GIA, titular, carpeta, conclusio), OBRIR-lo (boto) i marcar-lo
# com a IGNORAT (casella; reversible). Els canvis d'"ignorat" es desen al JSON.
function Invoke-InformesDbEdit {
    $outPath = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base d'informes'.",
            'Editar base d''informes', 'OK', 'Information') | Out-Null
        return
    }
    $db = $null
    try {
        $db = (Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
        return
    }

    # Aplanar en files, normalitzant cada informe (assegurar ignorat/ruta/motiu).
    # Cada fila guarda una REFERENCIA a l'objecte informe del JSON ($inf), aixi
    # marcar/desmarcar "ignorat" modifica directament la base que despres desem.
    $allRows = New-Object System.Collections.ArrayList
    if ($null -ne $db.activitats) {
        foreach ($act in $db.activitats) {
            if ($null -eq $act.informes) { continue }
            foreach ($inf in $act.informes) {
                if ($null -eq $inf.PSObject.Properties['ignorat']) { Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue $false -Force }
                if ($null -eq $inf.PSObject.Properties['ruta'])    { Add-Member -InputObject $inf -NotePropertyName ruta -NotePropertyValue '' -Force }
                if ($null -eq $inf.PSObject.Properties['motiu'])   { Add-Member -InputObject $inf -NotePropertyName motiu -NotePropertyValue '' -Force }
                [void]$allRows.Add([pscustomobject]@{
                    Obj       = $inf
                    Data      = [string]$inf.data
                    Gia       = [string]$act.id_gia
                    Titular   = [string]$act.titular
                    Carpeta   = [string]$act.carpeta
                    Conclusio = [string]$inf.conclusio
                    Motiu     = [string]$inf.motiu
                    Ruta      = [string]$inf.ruta
                })
            }
        }
    }

    # Estat compartit amb els gestors d'esdeveniments (hashtable per referencia).
    # 'Loading' evita que el gestor de la casella reaccioni mentre s'omple la
    # graella (Rows.Add pot disparar CellValueChanged abans d'assignar el Tag).
    $state = @{ Dirty = $false; Db = $db; Path = $outPath; Loading = $false }

    $form = _NewForm
    $form.Text = "Editar base d'informes"
    $form.Size = New-Object System.Drawing.Size(1000, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 420)

    # Graella.
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'None'
    $grid.MultiSelect = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White

    $cData = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cData.HeaderText = 'Data';      $cData.ReadOnly = $true; $cData.Width = 90
    $cGia  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cGia.HeaderText  = 'GIA';       $cGia.ReadOnly  = $true; $cGia.Width  = 60
    $cTit  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cTit.HeaderText  = 'Titular';   $cTit.ReadOnly  = $true; $cTit.Width  = 180
    $cCar  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCar.HeaderText  = 'Carpeta';   $cCar.ReadOnly  = $true; $cCar.Width  = 190
    $cCon  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCon.HeaderText  = 'Conclusio'; $cCon.ReadOnly  = $true; $cCon.Width  = 260
    $cMot  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cMot.HeaderText  = 'Motiu';     $cMot.ReadOnly  = $true; $cMot.Width  = 110
    $cObr  = New-Object System.Windows.Forms.DataGridViewButtonColumn;   $cObr.HeaderText  = '';          $cObr.Text = 'Obrir'; $cObr.UseColumnTextForButtonValue = $true; $cObr.Width = 64
    $cIgn  = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $cIgn.HeaderText  = 'Ignorar';   $cIgn.Width = 60
    [void]$grid.Columns.Add($cData)
    [void]$grid.Columns.Add($cGia)
    [void]$grid.Columns.Add($cTit)
    [void]$grid.Columns.Add($cCar)
    [void]$grid.Columns.Add($cCon)
    [void]$grid.Columns.Add($cMot)
    [void]$grid.Columns.Add($cObr)
    [void]$grid.Columns.Add($cIgn)
    $idxObrir   = 6
    $idxIgnorar = 7

    $fontNormal = $grid.Font
    $fontStrike = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Strikeout)

    # (Re)omple la graella segons el text del filtre.
    $fill = {
        param($needle)
        $state.Loading = $true
        $grid.Rows.Clear()
        $n = if ($needle) { ([string]$needle).ToLower() } else { '' }
        foreach ($row in $allRows) {
            if ($n -ne '') {
                $hay = (($row.Data + ' ' + $row.Gia + ' ' + $row.Titular + ' ' + $row.Carpeta + ' ' + $row.Conclusio + ' ' + $row.Motiu)).ToLower()
                if (-not $hay.Contains($n)) { continue }
            }
            $ign = [bool]$row.Obj.ignorat
            $idx = $grid.Rows.Add(@($row.Data, $row.Gia, $row.Titular, $row.Carpeta, $row.Conclusio, $row.Motiu, 'Obrir', $ign))
            $gr = $grid.Rows[$idx]
            $gr.Tag = $row
            $gr.Cells[4].ToolTipText = $row.Conclusio
            _StyleInformeRow $gr $ign $fontNormal $fontStrike
        }
        $state.Loading = $false
    }.GetNewClosure()

    # Barra superior: cerca.
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'; $topPanel.Height = 40
    $lblCerca = New-Object System.Windows.Forms.Label
    $lblCerca.Text = 'Cerca:'; $lblCerca.AutoSize = $true
    $lblCerca.Location = New-Object System.Drawing.Point(10, 12)
    $txtCerca = New-Object System.Windows.Forms.TextBox
    $txtCerca.Location = New-Object System.Drawing.Point(62, 9); $txtCerca.Width = 320
    $txtCerca.add_TextChanged({ & $fill $txtCerca.Text }.GetNewClosure())
    $topPanel.Controls.Add($lblCerca)
    $topPanel.Controls.Add($txtCerca)

    # Desa la base (retorna $true si va be).
    $doSave = {
        try {
            ($state.Db | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $state.Path -Encoding UTF8
            $state.Dirty = $false
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    # Barra inferior: botons.
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    $btnDesar = New-Object System.Windows.Forms.Button
    $btnDesar.Text = 'Desar'; $btnDesar.Size = New-Object System.Drawing.Size(120, 30)
    $btnDesar.Location = New-Object System.Drawing.Point(10, 9)
    $btnDesar.add_Click({
        if (& $doSave) {
            [System.Windows.Forms.MessageBox]::Show('Canvis desats.', 'Editar base d''informes', 'OK', 'Information') | Out-Null
        }
    }.GetNewClosure())
    $btnTancar = New-Object System.Windows.Forms.Button
    $btnTancar.Text = 'Enrere'; $btnTancar.Size = New-Object System.Drawing.Size(120, 30)
    $btnTancar.Location = New-Object System.Drawing.Point(140, 9)
    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())
    $botPanel.Controls.Add($btnDesar)
    $botPanel.Controls.Add($btnTancar)

    # Obrir l'informe en clicar el boto "Obrir".
    $grid.add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -eq $idxObrir) {
            $row = $s.Rows[$e.RowIndex].Tag
            $ruta = [string]$row.Ruta
            if ([string]::IsNullOrWhiteSpace($ruta) -or -not (Test-Path -LiteralPath $ruta)) {
                [System.Windows.Forms.MessageBox]::Show("No s'ha trobat el fitxer:`n$ruta", 'Obrir informe', 'OK', 'Warning') | Out-Null
                return
            }
            try { Start-Process -FilePath $ruta | Out-Null } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir:`n$($_.Exception.Message)", 'Obrir informe', 'OK', 'Error') | Out-Null
            }
        }
    }.GetNewClosure())

    # Forcem que la casella "Ignorar" es confirmi de seguida (no en sortir de la
    # cel-la), perque CellValueChanged salti al moment del clic.
    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    }.GetNewClosure())
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($state.Loading) { return }
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne $idxIgnorar) { return }
        $gr = $s.Rows[$e.RowIndex]
        $row = $gr.Tag
        if ($null -eq $row) { return }
        $val = [bool]$gr.Cells[$idxIgnorar].Value
        $row.Obj.ignorat = $val
        _StyleInformeRow $gr $val $fontNormal $fontStrike
        $state.Dirty = $true
    }.GetNewClosure())

    # Si es tanca amb canvis sense desar, oferim desar-los.
    $form.add_FormClosing({
        param($s, $e)
        if ($state.Dirty) {
            $r = [System.Windows.Forms.MessageBox]::Show('Hi ha canvis sense desar. Vols desar-los?', 'Editar base d''informes', 'YesNoCancel', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
                if (-not (& $doSave)) { $e.Cancel = $true }
            } elseif ($r -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $e.Cancel = $true
            }
        }
    }.GetNewClosure())

    # Ordre d'afegit: primer el Fill (queda al centre), despres Top i Bottom.
    $form.Controls.Add($grid)
    $form.Controls.Add($topPanel)
    $form.Controls.Add($botPanel)

    & $fill ''
    [void]$form.ShowDialog()
}
