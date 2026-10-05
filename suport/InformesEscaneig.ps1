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
# EL NUCLI de l'escaneig, SENSE CAP FINESTRA: el fan servir el boto (amb la
# seva finestra de progres) i el mode automatic (en segon pla, sense res).
# $onProgres (opcional): & $onProgres <text> <fets> <total> ($total 0 = encara no
# se sap). Torna @{ Ok; Error; NInformes; Reprocessats; NActivitats; NRevisar;
# OutPath }. Si alguna cosa peta a mitges, llanca (i no s'ha escrit res).
# ----------------------------------------------------------------------------
function _InformesDirAccessible([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { return $false }
    try { return [bool](Test-Path -LiteralPath $dir -ErrorAction SilentlyContinue) } catch { return $false }
}

function Invoke-InformesDbEscaneig([scriptblock]$onProgres = $null) {
    # La carpeta d'informes. Si la unitat (la I: de la feina) no hi es, no es un
    # error del programa: potser s'esta fora de la feina.
    $dir = $InformesDir
    if (-not (_InformesDirAccessible $dir)) {
        return @{ Ok = $false; Error = "No s'ha trobat la carpeta d'informes: $dir" }
    }
    $avisa = { param($t, $i, $n) if ($null -ne $onProgres) { & $onProgres $t $i $n } }

    # 3. Carregar l'Excel d'activitats (opcional; per la cerca inversa i el
    #    titular). Si no hi ha Excel, es continua sense aquest fallback.
    $cache = $null; $expToGia = $null
    try {
        $excel = Find-LatestActivitatsExcel
        if ($null -ne $excel) {
            & $avisa "Llegint la base d'activitats (Excel)..." 0 0
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
    & $avisa "Cercant informes a:`n$dir" 0 0
    $allInformes = Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
                   Where-Object {
                       $_.Name -notlike '~$*' -and
                       ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                       $null -ne (_ParseDataInformeFromName $_.Name)
                   }
    $files = @($allInformes)
    $total = $files.Count


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
                & $avisa "Analitzant informes... ($i de $total, $reprocessats de nous/modificats)" $i $total
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
    # Per ID GIA NUMERIC (com a text, '10' anava abans que '9').
    $activitatsOrd = @($activitats | Sort-Object { _GiaNumeric $_.id_gia }, { [string]$_.carpeta })

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

    return @{ Ok = $true; Error = ''; NInformes = $informes.Count; Reprocessats = $reprocessats
              NActivitats = $activitatsOrd.Count; NRevisar = $revisar.Count; OutPath = $outPath }
}

# Nom del mutex de l'escaneig: el comparteixen el boto i l'automatic, perque dos
# escaneigs alhora escriurien informes-db.json l'un sobre l'altre.
$Script:BaseMutexNom = 'Global\InformesCornella.BaseInformes'

# ----------------------------------------------------------------------------
# El BOTO "Actualitzar base": el nucli amb finestra de progres i un resum.
# ----------------------------------------------------------------------------
function Invoke-InformesDbScan {
    $dir = $InformesDir
    if (-not (_InformesDirAccessible $dir)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat la carpeta d'informes:`n$dir`n`nSi treballes fora de la feina (sense la unitat I:), obre-la quan hi tinguis accés. Pots canviar la ruta a Configuració.",
            'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }

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

    $onProgres = {
        param($text, $fets, $total)
        $lbl.Text = $text
        if ($total -gt 0) {
            if ($bar.Style -ne 'Continuous') { $bar.Style = 'Continuous'; $bar.Minimum = 0 }
            $bar.Maximum = [Math]::Max(1, $total)
            $bar.Value = [Math]::Min($bar.Maximum, $fets)
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()

    $caixa = @{ Res = $null }
    $res = $null
    $err = ''
    try {
        # Si el mode automatic esta escanejant ara mateix, no se'n fa un altre
        # al damunt (escriurien la base l'un sobre l'altre).
        $fet = Invoke-AmbMutexUnic $Script:BaseMutexNom {
            $caixa.Res = Invoke-InformesDbEscaneig $onProgres
        }
        if ($fet) { $res = $caixa.Res } else { $err = 'ocupat' }
    } catch {
        $err = $_.Exception.Message
    } finally {
        try { $form.Close() } catch { }
    }
    if ($err -eq 'ocupat') {
        [System.Windows.Forms.MessageBox]::Show(
            "La base d'informes s'est" + [char]0x00E0 + " actualitzant sola ara mateix (mode autom" + [char]0x00E0 + "tic).`n`nTorna-hi d'aqu" + [char]0x00ED + " a una estona.",
            'Base d''informes', 'OK', 'Information') | Out-Null
        return
    }
    if ($err -ne '' -or $null -eq $res) {
        [System.Windows.Forms.MessageBox]::Show("Error escanejant els informes:`n$err", 'Base d''informes', 'OK', 'Error') | Out-Null
        return
    }
    if (-not $res.Ok) {
        [System.Windows.Forms.MessageBox]::Show([string]$res.Error, 'Base d''informes', 'OK', 'Warning') | Out-Null
        return
    }
    # L'ultima l'has feta tu: la data del menu deixa de sortir en verd.
    [void](_BaseAutoDesaEstat @{ mode = 'manual' })
    $msg = "Base d'informes actualitzada.`n`n" +
           "Informes trobats: $($res.NInformes)`n" +
           "Nous o modificats (reprocessats): $($res.Reprocessats)`n" +
           "Activitats: $($res.NActivitats)`n" +
           "A revisar: $($res.NRevisar)`n`n" +
           "Fitxer:`n$($res.OutPath)`n`nVols obrir-lo?"
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Base d''informes', 'YesNo', 'Information')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$($res.OutPath)`"" | Out-Null } catch { }
    }
}

# ============================================================================
# MODE AUTOMATIC (l'interruptor A/M de sota la rajola "Actualitzar base")
# ============================================================================
# L'usuari: "ja que sera tan important per fer el Planol activitats, que tambe
# tingui l'opcio d'actualitzar-se automaticament (igual que Copiar informes)".
# Mateixa regla i mateixa hora (ModeAutomatic.ps1, $Script:AutoHora: les 13:00)
# amb el programa obert i, si l'ultima vegada que tocava no es va fer, en obrir
# el programa;
# en un proces a part (BaseInformesAuto.ps1), sense res a la pantalla.
#
# Les correccions a ma ("Editar base") PREVALEN igual que amb el boto: la
# passada es el mateix Invoke-InformesDbEscaneig.
#
# L'ESTAT, a informes-db-auto.json (no dins de la base: l'editor la reescriu
# sencera i s'hi perdria):
#   auto     l'interruptor
#   auto_el  l'ultima PASSADA automatica, encara que no hagi pogut fer res: es
#            la marca que diu que el venciment ja s'ha servit
#   mode     'auto' | 'manual': qui va fer l'ultima actualitzacio
$Script:BaseEstatPlantilla = [ordered]@{ auto = $false; auto_el = ''; mode = '' }

function _BaseAutoStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'informes-db-auto.json')
}

function _BaseAutoEstat { return (Read-EstatAuto (_BaseAutoStatePath) $Script:BaseEstatPlantilla) }

function _BaseAutoDesaEstat($canvis) { return (Save-EstatAuto (_BaseAutoStatePath) $canvis $Script:BaseEstatPlantilla) }

function _BaseAutoToca([datetime]$ara, $ultimAuto) {
    return (_AutoToca $ara $ultimAuto $Script:AutoHora $Script:AutoMinut)
}

function _BaseAutoLog([string]$msg) { Write-AutoLog 'informes-db-log.txt' $msg }

# La passada (la crida BaseInformesAuto.ps1). Cap finestra ni cap pregunta: tot
# va al registre. Apunta 'auto_el' SEMPRE (si no, el menu la tornaria a llancar
# cada minut), tambe quan la carpeta no hi es o quan el boto esta escanejant
# ara mateix (llavors la base ja s'esta posant al dia). Torna el resultat de
# l'escaneig, o @{ Ok = $false; Error }.
function Invoke-InformesDbAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $caixa = @{ Res = $null; Err = '' }
    $lliure = $false
    try {
        $lliure = Invoke-AmbMutexUnic $Script:BaseMutexNom {
            $caixa.Res = Invoke-InformesDbEscaneig $null
        }
    } catch { $caixa.Err = [string]$_.Exception.Message }
    if (-not $lliure -and $caixa.Err -eq '') {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog 'Passada automatica: ja s''esta actualitzant, no es fa res.'
        return @{ Ok = $false; Error = 'ocupat' }
    }
    if ($caixa.Err -ne '') {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog ('ERROR: ' + $caixa.Err)
        return @{ Ok = $false; Error = $caixa.Err }
    }
    $res = $caixa.Res
    if (-not [bool]$res.Ok) {
        [void](_BaseAutoDesaEstat $ini)
        _BaseAutoLog ('ATURAT: ' + ([string]$res.Error -replace "`r?`n", ' '))
        return $res
    }
    [void](_BaseAutoDesaEstat @{ auto_el = $ini['auto_el']; mode = 'auto' })
    _BaseAutoLog ("Passada automatica: informes=$($res.NInformes) reprocessats=$($res.Reprocessats) " +
                  "activitats=$($res.NActivitats) a_revisar=$($res.NRevisar)")
    return $res
}

function Start-InformesDbAuto { return (Start-ProcesAutoUnic 'base' 'BaseInformesAuto.ps1') }

# El menu ho crida en obrir-se i a cada minut.
function Invoke-BaseAutoSiToca {
    $est = _BaseAutoEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (_BaseAutoToca (Get-Date) $est['auto_el'])) { return $false }
    return (Start-InformesDbAuto)
}

$Script:ModesAuto['informesdb'] = @{
    Titol     = 'Actualitzar base'
    Actiu     = { $e = _BaseAutoEstat; [bool]$e['auto'] }
    DesaActiu = { param($on) _BaseAutoDesaEstat @{ auto = [bool]$on } }
    UltimMode = { $e = _BaseAutoEstat; [string]$e['mode'] }
    SiToca    = { Invoke-BaseAutoSiToca }
    Requisit  = {
        if (-not [string]::IsNullOrWhiteSpace($InformesDir)) { return '' }
        return ("Per actualitzar la base sola cal dir on s" + [char]0x00F3 + "n els informes.`n`nVes a Configuraci" + [char]0x00F3 + " (el bot" + [char]0x00F3 + " de la roda, a dalt a la dreta) i indica la carpeta d'informes.")
    }
    TipA      = (Get-AutoTipText "la base s'actualitza sola")
    TipM      = "Mode MANUAL: nomes s'actualitza quan cliques la rajola. Clica per posar-ho en automatic."
}
