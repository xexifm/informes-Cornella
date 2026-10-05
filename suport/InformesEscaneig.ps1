#requires -Version 5.1
<#
.SYNOPSIS
  "Actualitzar base": recorre la carpeta d'informes i en fa la base
  (local\base-dades-activitats\informes-db.json).

.DESCRIPTION
  Vivia a Informes.ps1, que tambe te l'editor de la base i els classificadors
  del text: amb les edicions a ma i el mode automatic (octubre 2026) passava de
  les 1.200 linies. Aqui hi ha el que LLEGEIX els informes del disc i munta la
  base; a Informes.ps1, el que n'enten el text (conclusio, GIA, expedient) i
  l'editor.

  NOMES DEFINEIX FUNCIONS.
#>

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

# Llegeix tots els paragrafs d'un .doc antic (Word 97-2003) via Word COM, en
# nomes-lectura. Necessita una instancia de Word JA OBERTA ($wordApp, creada
# mandrosament nomes si cal a Invoke-InformesDbScan); aqui nomes s'obre i es
# tanca el DOCUMENT (mai l'aplicacio).
function _ReadDocParagraphsWord($wordApp, $docPath) {
    $doc = $wordApp.Documents.Open($docPath, $false, $true, $false)
    try {
        $out = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            $t = $p.Range.Text
            if ($null -ne $t) { $t = $t.TrimEnd([char]13, [char]7) }
            [void]$out.Add($t)
        }
        return $out.ToArray()
    } finally {
        $doc.Close($false)
    }
}

# Tria com llegir els paragrafs d'un informe segons l'extensio: .docx (sense
# Word, via zip) o .doc antic (via Word COM; retorna buit si no hi ha Word
# disponible, i l'informe queda "a revisar" com si no s'hagues pogut llegir).
function _ReadInformeParagraphs($file, $wordApp) {
    if ($file.Extension -ieq '.doc') {
        if ($null -eq $wordApp) { return @() }
        return _ReadDocParagraphsWord $wordApp $file.FullName
    }
    return _ReadDocxParagraphs $file.FullName
}

# Analitza UN informe. Retorna un PSCustomObject amb data, gia, expedient,
# conclusio, fitxer, ruta, carpeta i el motiu (si cal revisar-lo). $wordApp es
# opcional (nomes cal per als .doc antics; vegeu _ReadInformeParagraphs).
function Get-InformeData($file, $expToGia, $cache, $wordApp = $null) {
    $data = _ParseDataInformeFromName $file.Name
    $lines = @()
    try { $lines = _ReadInformeParagraphs $file $wordApp } catch { $lines = @() }

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

    $conclInfo = _ExtractConclusio $lines
    $concl = $conclInfo.Text

    $motius = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($gia)) { [void]$motius.Add('sense ID GIA') }
    $conclMotiu = _ConclusioMotiu $conclInfo
    if (-not [string]::IsNullOrWhiteSpace($conclMotiu)) { [void]$motius.Add($conclMotiu) }

    $titular = ''
    if ($null -ne $cache -and -not [string]::IsNullOrWhiteSpace($gia) -and $cache.ById.ContainsKey([string]$gia)) {
        $titular = [string]$cache.ById[[string]$gia].TITULAR
    }

    return [pscustomobject]@{
        Data          = $data
        Gia           = $gia
        GiaFont       = $font
        Expedient     = $exp
        Titular       = $titular
        Conclusio     = $concl
        ConclusioBreu = (_ConclusioBreu $concl)
        Fitxer        = $file.Name
        Ruta          = $file.FullName
        Carpeta       = _CarpetaActivitat $file.FullName
        Modificat     = $file.LastWriteTimeUtc.ToString('o')
        Ignorat       = (_ConclusioIgnorarPerDefecte $conclInfo)
        Motius        = $motius.ToArray()
        EditatAMa         = $false
        AutoConclusioBreu = ''
        AutoIgnorat       = $false
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
            $teMarca = ($null -ne $inf.PSObject.Properties['editat_a_ma'])
            _InferEditatAMa $inf
            $motiuStr = if ($null -ne $inf.PSObject.Properties['motiu']) { [string]$inf.motiu } else { '' }
            $motius = if ([string]::IsNullOrWhiteSpace($motiuStr)) { @() } else { @($motiuStr -split ',\s*') }
            $ign = $false
            if ($null -ne $inf.PSObject.Properties['ignorat']) { $ign = [bool]$inf.ignorat }
            $conclusioText = [string]$inf.conclusio
            # Compatibilitat: si la base es d'abans d'aquest camp, la calculem
            # ara mateix (no cal reescanejar per tenir-la la primera vegada).
            $conclusioBreu = if ($null -ne $inf.PSObject.Properties['conclusio_breu']) { [string]$inf.conclusio_breu } else { _ConclusioBreu $conclusioText }
            $map[$ruta] = [pscustomobject]@{
                Data          = [string]$inf.data
                Gia           = [string]$act.id_gia
                GiaFont       = ''
                Expedient     = [string]$act.expedient
                Titular       = [string]$act.titular
                Conclusio     = $conclusioText
                ConclusioBreu = $conclusioBreu
                Fitxer        = [string]$inf.fitxer
                Ruta          = $ruta
                Carpeta       = [string]$act.carpeta
                Modificat     = if ($null -ne $inf.PSObject.Properties['modificat']) { [string]$inf.modificat } else { '' }
                Ignorat       = $ign
                Motius        = $motius
                EditatAMa         = [bool](_PropInf $inf 'editat_a_ma')
                TeMarcaEdicio     = $teMarca
                AutoConclusioBreu = [string](_PropInf $inf 'auto_conclusio_breu')
                AutoIgnorat       = [bool](_PropInf $inf 'auto_ignorat')
            }
        }
    }
    return $map
}

# En reprocessar un informe que ja era a la base: el de l'usuari PREVAL si l'hi
# havia corregit; si no, mana el que acaba de sortir de l'informe (que potser
# s'ha tornat a escriure). $prev: el registre de _FlattenInformesDb; $nou: el
# de Get-InformeData. Retorna $nou, amb les correccions si n'hi ha. PURA.
#
# Una base d'abans de la marca (el registre no la porta) es tracta com
# _InferEditatAMa, i a mes, aqui si que es pot dir de l'ignorat: si el de la
# base no es el per defecte que acaba de sortir, l'havia posat l'usuari.
function _AplicaEdicioPrevia($prev, $nou) {
    $manual = [bool]$prev.EditatAMa
    if (-not [bool]$prev.TeMarcaEdicio) {
        $manual = $manual -or ([bool]$prev.Ignorat -ne [bool]$nou.Ignorat)
    }
    if (-not $manual) { return $nou }
    $nou.AutoConclusioBreu = [string]$nou.ConclusioBreu
    $nou.AutoIgnorat = [bool]$nou.Ignorat
    $nou.ConclusioBreu = [string]$prev.ConclusioBreu
    $nou.Ignorat = [bool]$prev.Ignorat
    $nou.EditatAMa = $true
    return $nou
}

# Un informe tal com es desa a informes-db.json. Els camps de l'edicio a ma
# (auto_*) nomes si ho esta: la base no s'omple de camps buits. PURA.
function _InformeAJson($r) {
    $o = [pscustomobject]@{
        data           = $r.Data
        fitxer         = $r.Fitxer
        ruta           = $r.Ruta
        conclusio      = $r.Conclusio
        conclusio_breu = $r.ConclusioBreu
        modificat      = $r.Modificat
        ignorat        = [bool]$r.Ignorat
        motiu          = (@($r.Motius) -join ', ')
        editat_a_ma    = [bool]$r.EditatAMa
    }
    if ([bool]$r.EditatAMa) {
        Add-Member -InputObject $o -NotePropertyName auto_conclusio_breu -NotePropertyValue ([string]$r.AutoConclusioBreu)
        Add-Member -InputObject $o -NotePropertyName auto_ignorat -NotePropertyValue ([bool]$r.AutoIgnorat)
    }
    return $o
}

# ----------------------------------------------------------------------------
# Escaneig complet + escriptura del JSON (interactiu, amb finestra de progres)
# ----------------------------------------------------------------------------
function Invoke-InformesDbScan {
    # 1. Resoldre la carpeta d'informes. -ErrorAction SilentlyContinue: si la
    #    unitat (p.ex. la I: de la feina) no existeix, Test-Path no ha de petar,
    #    nomes ha de donar 'no trobada' (potser estas fora de la feina).
    $dir = $InformesDir
    $existeix = $false
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        try { $existeix = Test-Path -LiteralPath $dir -ErrorAction SilentlyContinue } catch { $existeix = $false }
    }
    if (-not $existeix) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la carpeta d'informes:`n$dir`n`nSi treballes fora de la feina (sense la unitat I:), obre-la quan hi tinguis accés. Pots canviar la ruta amb `$InformesDir a config.ps1.",
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
                $prevDb     = Read-JsonFile $outPath
                $prevByRuta = _FlattenInformesDb $prevDb
                if ($prevDb.PSObject.Properties['actualitzat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.actualitzat_el)) {
                    try { $prevUtc = ([datetime]::Parse([string]$prevDb.actualitzat_el)).ToUniversalTime() } catch { $prevUtc = [datetime]::MinValue }
                }
                if ($prevDb.PSObject.Properties['generat_el'] -and -not [string]::IsNullOrWhiteSpace([string]$prevDb.generat_el)) {
                    $generatEl = [string]$prevDb.generat_el
                }
            } catch { $prevByRuta = @{}; $prevUtc = [datetime]::MinValue }
        }

        # 4. Recollir els fitxers candidats (.docx o .doc amb data al principi
        #    del nom). Un sol Get-ChildItem recursiu (sense -Filter) i filtrem
        #    per extensio nosaltres: evita el parany de "*.doc" -Filter que a
        #    vegades tambe encerta ".docx" pel nom curt (8.3) de NTFS.
        $lbl.Text = "Cercant informes a:`n$dir"
        [System.Windows.Forms.Application]::DoEvents()
        $allInformes = Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
                       Where-Object {
                           $_.Name -notlike '~$*' -and
                           ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                           $null -ne (_ParseDataInformeFromName $_.Name)
                       }
        $files = @($allInformes)
        $total = $files.Count

        $bar.Style = 'Continuous'
        $bar.Minimum = 0
        $bar.Maximum = [Math]::Max(1, $total)

        # 5. Analitzar cada informe (incremental: reutilitzem els no modificats).
        #    Word només es crea (mandrosament) si cal reprocessar algun .doc
        #    antic; es tanca sempre al 'finally', encara que hi hagi un error.
        $informes = New-Object System.Collections.ArrayList
        $revisar  = New-Object System.Collections.ArrayList
        $reprocessats = 0
        $i = 0
        $wordApp = $null
        try {
            foreach ($f in $files) {
                $i++
                $ruta = $f.FullName
                $teEntrada = $prevByRuta.ContainsKey($ruta)
                if (-not (_HaDeReprocessar $f.LastWriteTimeUtc $prevUtc $teEntrada)) {
                    # No s'ha tocat des de l'ultim escaneig: reutilitzem l'entrada.
                    $r = $prevByRuta[$ruta]
                } else {
                    if ($f.Extension -ieq '.doc' -and $null -eq $wordApp) {
                        # -Opcional: sense Word, _ReadInformeParagraphs torna @()
                        # i l'informe es queda sense conclusio, pero l'escaneig
                        # continua. New-WordApp hi afegeix l'AutomationSecurity,
                        # que aqui compta: aquests .doc son a la unitat de xarxa.
                        $wordApp = New-WordApp -Opcional
                    }
                    $r = Get-InformeData $f $expToGia $cache $wordApp
                    # El que l'usuari hagi corregit a ma ("Editar base") PREVAL;
                    # la resta, la mana el que acaba de sortir de l'informe
                    # (vegeu _AplicaEdicioPrevia).
                    if ($teEntrada) { $r = _AplicaEdicioPrevia $prevByRuta[$ruta] $r }
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
        } finally {
            if ($null -ne $wordApp) { try { $wordApp.Quit() } catch { } }
        }

        # 6. Agrupar per activitat: per ID GIA quan n'hi ha; si NO en tenen, per
        #    CARPETA (tots els informes d'una mateixa carpeta = una activitat).
        #    Ordenem els informes de cada activitat per data.
        $groups = [ordered]@{}
        foreach ($r in $informes) {
            $key = if (-not [string]::IsNullOrWhiteSpace($r.Gia)) { "GIA:$($r.Gia)" }
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
            [void]$g._informes.Add((_InformeAJson $r))
        }

        $activitats = New-Object System.Collections.ArrayList
        foreach ($g in $groups.Values) {
            $ordered = @($g._informes | Sort-Object { if ($_.data) { $_.data } else { '' } })
            [void]$activitats.Add([pscustomobject]@{
                id_gia       = $g.id_gia
                expedient    = $g.expedient
                titular      = $g.titular
                carpeta      = $g.carpeta
                estat_actual = (_EstatActualActivitat $ordered)
                informes     = $ordered
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
        Write-JsonFile $outPath $outObj 8

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
