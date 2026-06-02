#requires -Version 5.1
<#
.SYNOPSIS
  Mode "Informe de seguiment": pren un informe anterior amb requeriments
  enumerats i hi afegeix, sota cada punt, una anotacio datada indicant si la
  documentacio presentada l'ha resolt o no.

.DESCRIPTION
  Es un mode alternatiu del programa (es tria a la pantalla inicial). NO
  regenera l'informe: treballa sobre una COPIA del .docx i nomes
  insereix/esborra/formata, de manera que es preserva exactament la capcalera,
  el text dels requeriments i el format (tambe d'informes fets a ma).

  Flux (Invoke-SeguimentFlow):
    1. Triar l'informe anterior (.docx).
    2. Llegir-lo i modelar requeriments + anotacions existents.
    3. Confirmar l'esborrat del bloc de conclusions antic (amb preview).
    4. Introduir la data de la ronda (per defecte, avui).
    5. Per cada requeriment: comentari nou + checkbox "Resolt".
    6. Triar conclusions (reutilitza 0 CONCLUSIONS.docx) i omplir camps.
    7. Aplicar sobre la copia: esborrar conclusions -> inserir anotacions
       (de baix a dalt) -> recalcular negreta -> afegir conclusions -> desar.

  Iteratiu: tornar a passar-lo sobre un informe de seguiment AFEGEIX una linia
  nova sota cada requeriment (no duplica). La negreta es dinamica: mentre un
  requeriment NO estigui resolt, el requeriment i les seves anotacions van en
  negreta; quan es resol, res en negreta.

.NOTES
  Aquest fitxer es carrega via dot-source des de GenerarInforme.ps1 (tambe en
  mode headless de proves). Les funcions pures (sense COM/WinForms) son
  testejables a Linux; les funcions COM/WinForms nomes s'executen a Windows.

  Reutilitza de GenerarInforme.ps1: New-WordApp, Close-WordApp, _NormalizeText,
  Read-Conclusions, Select-Conclusions, Add-FieldsFromConclusions, Prompt-Fields,
  _WriteConclusionsBlock, _ResolveOutputDir, _GetUniqueOutputPath; i de
  Format.ps1: $ReportFormatConfig.
#>

# ----------------------------------------------------------------------------
# Expressions regulars (a nivell de script: definides en carregar el fitxer).
# ----------------------------------------------------------------------------
# Requeriment numerat: comenca per "N." seguit d'espai. Ex: "1. Baixa tensio."
$Script:SeguimentReqRegex      = [regex]'^\s*(\d+)\.\s'
# Anotacio datada: "dd/MM/aaaa:" al principi. Ex: "01/06/2026: No s'entrega."
$Script:SeguimentAnnotRegex    = [regex]'^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*:'
# Numero de llista auto-numerada del Word (ListFormat.ListString). Ex: "1." o "1"
$Script:SeguimentListNumRegex  = [regex]'^\d+\.?$'

# ----------------------------------------------------------------------------
# FUNCIONS PURES (sense COM/WinForms) - testejables en headless
# ----------------------------------------------------------------------------

# Classifica un paragraf a partir del seu text i del numero de llista (si en
# te). Retorna { Kind; Number; Date; ViaList } on Kind es:
#   'requirement' (requeriment numerat) | 'annotation' (anotacio datada) | 'other'
function _ClassifyParagraph([string]$text, [string]$listString) {
    $t  = if ($null -eq $text)       { '' } else { ([string]$text).Trim() }
    $ls = if ($null -eq $listString) { '' } else { ([string]$listString).Trim() }

    # 1) Anotacio datada (es comprova primer; "dd/MM/aaaa" no casa amb "N.").
    $ma = $Script:SeguimentAnnotRegex.Match($t)
    if ($ma.Success) {
        $d = '{0:D2}/{1:D2}/{2}' -f [int]$ma.Groups[1].Value, [int]$ma.Groups[2].Value, $ma.Groups[3].Value
        return [pscustomobject]@{ Kind='annotation'; Number=$null; Date=$d; ViaList=$false }
    }

    # 2) Requeriment amb numero literal al text ("1. ...").
    $mr = $Script:SeguimentReqRegex.Match($t)
    if ($mr.Success) {
        return [pscustomobject]@{ Kind='requirement'; Number=[int]$mr.Groups[1].Value; Date=$null; ViaList=$false }
    }

    # 3) Requeriment via llista auto-numerada del Word (informes fets a ma).
    if ($ls -ne '' -and $Script:SeguimentListNumRegex.IsMatch($ls)) {
        $num = ($ls.TrimEnd('.') -as [int])
        return [pscustomobject]@{ Kind='requirement'; Number=$num; Date=$null; ViaList=$true }
    }

    return [pscustomobject]@{ Kind='other'; Number=$null; Date=$null; ViaList=$false }
}

# Decideix l'estat "resolt" a partir del valor Font.Bold del Word:
#   -1       = tot en negreta  -> pendent  (NO resolt)
#    0       = res en negreta   -> resolt
#    9999999 = mixt (wdUndefined, p.ex. el "N." en negreta i el text no)
#              -> el tractem com a pendent (NO resolt), que es el cas de
#                 l'informe original acabat de generar.
function _InferResolvedFromBold($boldValue) {
    return ((($boldValue -as [int]) -eq 0))
}

# Mentre un requeriment no estigui resolt, ha d'anar en negreta.
function _ShouldBeBold($resolved) {
    return (-not $resolved)
}

# Munta el model ordenat de requeriments-amb-anotacions a partir d'un array de
# registres de paragraf { Index; Text; ListString; Bold }. Funcio PURA: el
# lector COM nomes recull els registres i crida aqui.
function _BuildSeguimentModel($paraRecords) {
    $reqs       = New-Object System.Collections.ArrayList
    $current    = $null
    $lastReqIdx = 0
    foreach ($r in $paraRecords) {
        $c = _ClassifyParagraph $r.Text $r.ListString
        if ($c.Kind -eq 'requirement') {
            $current = [pscustomobject]@{
                Number         = $c.Number
                ParaIndex      = [int]$r.Index
                Text           = ([string]$r.Text).Trim()
                IsAutoNumbered = [bool]$c.ViaList
                Bold           = $r.Bold
                WasResolved    = (_InferResolvedFromBold $r.Bold)
                Annotations    = (New-Object System.Collections.ArrayList)
            }
            [void]$reqs.Add($current)
            $lastReqIdx = [int]$r.Index
        }
        elseif ($c.Kind -eq 'annotation' -and $null -ne $current) {
            [void]$current.Annotations.Add([pscustomobject]@{
                ParaIndex = [int]$r.Index
                Date      = $c.Date
                Text      = ([string]$r.Text).Trim()
            })
        }
        # 'other' (cos, URL, intro, capcalera, conclusions...) s'ignora al model.
    }
    return [pscustomobject]@{
        Requirements     = $reqs.ToArray()
        LastReqParaIndex = $lastReqIdx
    }
}

# Localitza l'index (1-based) del primer paragraf del bloc de conclusions:
# busca NOMES despres de l'ultim requeriment (perque una frase dins un cos de
# requeriment no dispari) i casa qualsevol de les frases conegudes de manera
# insensible a accents/majuscules. Retorna -1 si no en troba cap.
#   $paraTexts : array de textos de paragraf (posicio i => paragraf i+1).
function _FindConclusionStartIndex($paraTexts, [int]$lastReqEndIndex, $phrases) {
    if ($null -eq $paraTexts) { return -1 }
    $normPhrases = @()
    foreach ($ph in $phrases) {
        $np = _NormalizeText $ph
        if (-not [string]::IsNullOrWhiteSpace($np)) { $normPhrases += $np }
    }
    if ($normPhrases.Count -eq 0) { return -1 }

    for ($i = 0; $i -lt $paraTexts.Count; $i++) {
        $paraIndex = $i + 1
        if ($paraIndex -le $lastReqEndIndex) { continue }
        $nt = _NormalizeText $paraTexts[$i]
        if ([string]::IsNullOrWhiteSpace($nt)) { continue }
        foreach ($np in $normPhrases) {
            if ($nt.Contains($np)) { return $paraIndex }
        }
    }
    return -1
}

# Valida/normalitza la data de la ronda. Buit -> avui. Retorna { Ok; Normalized }
# sempre en dd/MM/yyyy. Rebutja dates impossibles (32/13/2026...).
function _ValidateRoundDate($text) {
    $s = if ($null -eq $text) { '' } else { ([string]$text).Trim() }
    if ($s -eq '') {
        return [pscustomobject]@{ Ok=$true; Normalized=(Get-Date).ToString('dd/MM/yyyy') }
    }
    $dt   = [datetime]::MinValue
    $fmts = [string[]]@('dd/MM/yyyy','d/M/yyyy','dd/MM/yy','d/M/yy')
    $ci   = [System.Globalization.CultureInfo]::InvariantCulture
    if ([datetime]::TryParseExact($s, $fmts, $ci, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return [pscustomobject]@{ Ok=$true; Normalized=$dt.ToString('dd/MM/yyyy') }
    }
    return [pscustomobject]@{ Ok=$false; Normalized=$s }
}

# Format unic de la linia d'anotacio: "dd/MM/aaaa: comentari".
function _FormatAnnotationLine($dateStr, $comment) {
    $c = if ($null -eq $comment) { '' } else { ([string]$comment).Trim() }
    return ('{0}: {1}' -f $dateStr, $c)
}

# Nom del fitxer de sortida del seguiment.
#   - Font amb l'esquema del programa "YYYY-MM-DD_<Cat>_GIA <id>" ->
#     "<data>_<Cat>_GIA <id>_SEG.docx" (preservant Cat i GIA; sense duplicar _SEG).
#   - Font feta a ma -> "<data>_Seguiment_<nom>.docx".
# Sempre s'ha de passar el resultat per _GetUniqueOutputPath (afegeix _2, _3...).
function _SeguimentOutputName([string]$sourceBaseName, [datetime]$roundDate) {
    $day = $roundDate.ToString('yyyy-MM-dd')
    $rx  = [regex]'^\d{4}-\d{2}-\d{2}_(.+?)_GIA\s+(\d+|s_n)'
    $m   = $rx.Match([string]$sourceBaseName)
    if ($m.Success) {
        $cat = $m.Groups[1].Value
        $gia = $m.Groups[2].Value
        return ('{0}_{1}_GIA {2}_SEG.docx' -f $day, $cat, $gia)
    }
    $safe = (([string]$sourceBaseName) -replace '[\\/:*?"<>|]','_').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'informe' }
    return ('{0}_Seguiment_{1}.docx' -f $day, $safe)
}

# ----------------------------------------------------------------------------
# CAPA COM - lectura i edicio del .docx (nomes Windows + Word)
# ----------------------------------------------------------------------------

# Recull els registres de paragraf del document obert (1-based, alineat amb
# $doc.Paragraphs.Item($i)).
function _CollectParaRecords($doc) {
    $records = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($p in $doc.Paragraphs) {
        $i++
        $text = ''
        try { $text = [string]$p.Range.Text } catch { }
        $text = $text.TrimEnd("`r","`n","`a"," ")
        $ls = ''
        try { $ls = [string]$p.Range.ListFormat.ListString } catch { }
        $bold = $null
        try { $bold = $p.Range.Font.Bold } catch { }
        [void]$records.Add([pscustomobject]@{ Index=$i; Text=$text; ListString=$ls; Bold=$bold })
    }
    return $records.ToArray()
}

# Obre una COPIA local (TEMP) de l'informe anterior i en construeix el model.
# Retorna { Doc; TempPath; Model; ParaTexts; Records }. El document queda OBERT
# (s'editara mes tard a Apply-Seguiment).
function Read-PreviousReport($word, $path) {
    $tempPath = Join-Path $env:TEMP ('SEG_' + [System.IO.Path]::GetFileName($path))
    Copy-Item -LiteralPath $path -Destination $tempPath -Force

    # L'informe sovint ve de la unitat de xarxa (I:\...). La copia local pot
    # heretar l'atribut de NOMES-LECTURA i la "marca web" (Mark of the Web),
    # que fan que el Word l'obri en mode protegit / nomes-lectura i, llavors,
    # les insercions fallen amb "este comando no esta disponible para la
    # lectura". Ho neutralitzem ABANS d'obrir.
    try { (Get-Item -LiteralPath $tempPath).IsReadOnly = $false } catch { }
    try { Unblock-File -LiteralPath $tempPath } catch { }

    # Obrim explicitament en mode NO nomes-lectura.
    #   Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ...)
    $doc = $word.Documents.Open($tempPath, $false, $false)

    # Si tot i aixi s'ha obert en vista protegida, passem a edicio.
    try {
        if ($doc.ProtectedViewWindow -ne $null) { $doc = $doc.ProtectedViewWindow.Edit() }
    } catch { }

    # Si el document porta proteccio d'edicio (formularis / nomes-lectura sense
    # contrasenya), la traiem perque es pugui editar. Si te contrasenya,
    # Unprotect falla i ho deixem com esta (es detectara mes avall).
    try { if ($doc.ProtectionType -ne -1) { $doc.Unprotect() } } catch { }

    # Ultim recurs: si encara es nomes-lectura, ho diem clarament en lloc de
    # petar a mitja edicio.
    $isReadOnly = $false
    try { $isReadOnly = [bool]$doc.ReadOnly } catch { }
    if ($isReadOnly) {
        throw ("El document s'ha obert en mode nomes-lectura i no es pot editar. " +
               "Comprova que el fitxer no estigui obert en una altra finestra del Word " +
               "ni marcat com a nomes-lectura, i torna-ho a provar.")
    }

    $records   = _CollectParaRecords $doc
    $model     = _BuildSeguimentModel $records
    $paraTexts = @($records | ForEach-Object { $_.Text })
    return [pscustomobject]@{
        Doc       = $doc
        TempPath  = $tempPath
        Model     = $model
        ParaTexts = $paraTexts
        Records   = $records
    }
}

# Aplica el seguiment sobre el document ja obert i el desa al directori de
# sortida. Retorna la ruta final.
#   $decisions : array alineat amb $model.Requirements; cada element
#                { Resolved; NewComment }.
#   $conclusionStartIndex : index 1-based del primer paragraf de conclusions a
#                esborrar, o -1 / 0 per no esborrar res.
function Apply-Seguiment {
    param(
        $word, $doc, $tempPath, $model, [int]$conclusionStartIndex,
        $decisions, $dateStr,
        $conclHeaderText, $selectedConclusions, $alwaysConclusions, $fields,
        [string]$sourceBaseName, [datetime]$roundDate
    )

    # Cada fase es marca a $phase: si una crida COM falla, reportem en quina
    # fase ha estat (ajuda a diagnosticar errors com "conversio no valida").
    $phase = 'inici'
    try {
        # (1) Esborrar PRIMER el bloc de conclusions antic (de l'inici detectat
        # fins al final de la historia). Aixi els indexs de requeriments i
        # anotacions (tots anteriors) no es veuen afectats. Fem servir
        # l'assignacio de Range.End (mes robusta que $doc.Range(start,end), que
        # en alguns equips llanca "conversio no valida").
        $phase = 'esborrar conclusions antigues'
        if ($conclusionStartIndex -ge 1) {
            $rng = $doc.Paragraphs.Item($conclusionStartIndex).Range
            $rng.End = $doc.Content.End
            [void]$rng.Delete()
        }

        # (2) Inserir anotacions de BAIX A DALT, perque les insercions no
        # invalidin els indexs dels requeriments encara no processats (els que
        # tenen index mes petit). L'ancora es l'ultima anotacio existent del
        # requeriment, o el propi paragraf del requeriment si encara no en te cap.
        $phase = 'inserir anotacions'
        for ($k = $model.Requirements.Count - 1; $k -ge 0; $k--) {
            $req     = $model.Requirements[$k]
            $dec     = $decisions[$k]
            $comment = [string]$dec.NewComment
            if ([string]::IsNullOrWhiteSpace($comment)) { continue }

            $anchorIdx = if ($req.Annotations.Count -gt 0) {
                [int]$req.Annotations[$req.Annotations.Count - 1].ParaIndex
            } else {
                [int]$req.ParaIndex
            }

            # Insercio: afegim un paragraf buit despres de l'ancora i hi
            # escrivim el text. Tot amb metodes de Range (sense $doc.Range(a,b),
            # que en alguns equips llanca "conversio no valida").
            $anchorRange = $doc.Paragraphs.Item($anchorIdx).Range
            [void]$anchorRange.InsertParagraphAfter()

            # El nou paragraf es ara el seguent (anchorIdx + 1). Hi escrivim el
            # text abans de la seva marca de paragraf.
            $np2 = $doc.Paragraphs.Item($anchorIdx + 1)
            $line = _FormatAnnotationLine $dateStr $comment
            $np2.Range.InsertBefore($line)
            $r2  = $np2.Range
            # Evitar que l'anotacio hereti la numeracio de llista del requeriment.
            try { [void]$r2.ListFormat.RemoveNumbers() } catch { }
            try { $r2.Font.Bold      = 0 } catch { }
            try { $r2.Font.Italic    = 0 } catch { }
            try { $r2.Font.Underline = 0 } catch { }
            try { $r2.Font.Size      = $Script:ReportFormatConfig.BodyFontSize } catch { }
            try { $np2.Range.ParagraphFormat.LeftIndent = $doc.Paragraphs.Item($req.ParaIndex).Range.ParagraphFormat.LeftIndent } catch { }
        }

        # (3) Recalcular la negreta de tota la "columna" de cada requeriment
        # segons l'estat ACTUAL (re-escanegem perque els indexs han canviat amb
        # les insercions). Pendents -> negreta; resolts -> sense negreta.
        $phase = 'recalcular negreta'
        $records2 = _CollectParaRecords $doc
        $model2   = _BuildSeguimentModel $records2
        $n = [Math]::Min($model2.Requirements.Count, $decisions.Count)
        for ($k = 0; $k -lt $n; $k++) {
            $req2 = $model2.Requirements[$k]
            $bold = if (_ShouldBeBold $decisions[$k].Resolved) { 1 } else { 0 }
            try { $doc.Paragraphs.Item($req2.ParaIndex).Range.Font.Bold = $bold } catch { }
            foreach ($a in $req2.Annotations) {
                try { $doc.Paragraphs.Item($a.ParaIndex).Range.Font.Bold = $bold } catch { }
            }
        }

        # (4) Afegir les conclusions noves al final de la historia (reutilitza el
        # mateix escriptor que el generador). Cal fer-ho amb Selection, despres
        # de totes les edicions amb Range.
        $phase = 'escriure conclusions'
        $doc.Activate()
        $sel = $word.Selection
        [void]$sel.EndKey(6)  # wdStory = 6
        _WriteConclusionsBlock $sel $Script:ReportFormatConfig $conclHeaderText $selectedConclusions $alwaysConclusions $fields

        # (5) Desar i moure al directori de sortida amb nom unic.
        $phase = 'desar'
        $outName   = _SeguimentOutputName $sourceBaseName $roundDate
        $targetDir = _ResolveOutputDir
        $outPath   = _GetUniqueOutputPath $targetDir $outName

        $doc.Save()
        $doc.Close($false)
        try {
            Move-Item -LiteralPath $tempPath -Destination $outPath -Force
        } catch {
            return $tempPath
        }
        return $outPath
    }
    catch {
        throw ("[fase: {0}] {1}" -f $phase, $_.Exception.Message)
    }
}

# ----------------------------------------------------------------------------
# CAPA WINFORMS - dialegs del flux de seguiment (nomes Windows)
# ----------------------------------------------------------------------------

# Pantalla inicial: tria entre generar un informe nou o fer un seguiment.
# Retorna 'nou' | 'seguiment'. Tancar la finestra (X) avorta (exit 0).
function Select-Mode {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Informes Cornella'
    $form.Size = New-Object System.Drawing.Size(440, 235)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Que vols fer?'
    $lbl.Location = New-Object System.Drawing.Point(20, 15)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $btnNou = New-Object System.Windows.Forms.Button
    $btnNou.Text = 'Generar informe nou'
    $btnNou.Location = New-Object System.Drawing.Point(20, 50)
    $btnNou.Size = New-Object System.Drawing.Size(390, 45)
    $btnNou.DialogResult = 'Yes'
    $form.Controls.Add($btnNou)

    $btnSeg = New-Object System.Windows.Forms.Button
    $btnSeg.Text = "Fer seguiment d'un informe existent"
    $btnSeg.Location = New-Object System.Drawing.Point(20, 105)
    $btnSeg.Size = New-Object System.Drawing.Size(390, 45)
    $btnSeg.DialogResult = 'No'
    $form.Controls.Add($btnSeg)

    $form.AcceptButton = $btnNou
    $res = $form.ShowDialog()
    if ($res -eq 'Yes') { return 'nou' }
    if ($res -eq 'No')  { return 'seguiment' }
    exit 0
}

# Tria de l'informe anterior (.docx). Retorna la ruta o $null si es cancel·la.
function Select-PreviousReport {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Documents Word (*.docx)|*.docx'
    $dlg.Title  = "Tria l'informe anterior amb els requeriments"
    try {
        $od = _ResolveOutputDir
        if (Test-Path -LiteralPath $od) { $dlg.InitialDirectory = $od }
    } catch { }
    if ($dlg.ShowDialog() -ne 'OK') { return $null }
    return $dlg.FileName
}

# Confirma l'esborrat del bloc de conclusions. Mostra el bloc detectat; permet
# ajustar manualment (informes fets a ma) o no esborrar res.
# Retorna { StartIndex } amb StartIndex 1-based, o -1 per no esborrar; $null si
# l'usuari cancel·la tot el flux.
function Confirm-ConclusionDeletion {
    param($paraTexts, [int]$lastReqEndIndex, [int]$detectedStart)

    if ($detectedStart -ge 1) {
        $lines = @()
        for ($i = $detectedStart; $i -le $paraTexts.Count; $i++) {
            $tx = [string]$paraTexts[$i - 1]
            if (-not [string]::IsNullOrWhiteSpace($tx)) { $lines += $tx }
        }
        $preview = ($lines -join "`r`n")

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Conclusions a esborrar'
        $form.Size = New-Object System.Drawing.Size(720, 480)
        $form.StartPosition = 'CenterScreen'

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "S'esborrara aquest bloc de conclusions abans d'afegir-ne de noves:"
        $lbl.Location = New-Object System.Drawing.Point(15, 10)
        $lbl.AutoSize = $true
        $form.Controls.Add($lbl)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.ReadOnly = $true
        $tb.ScrollBars = 'Vertical'
        $tb.Location = New-Object System.Drawing.Point(15, 35)
        $tb.Size = New-Object System.Drawing.Size(680, 360)
        $tb.Text = $preview
        $form.Controls.Add($tb)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = 'Esborrar i continuar'
        $btnOk.Location = New-Object System.Drawing.Point(15, 405)
        $btnOk.Size = New-Object System.Drawing.Size(160, 30)
        $btnOk.DialogResult = 'OK'
        $form.Controls.Add($btnOk)

        $btnAdj = New-Object System.Windows.Forms.Button
        $btnAdj.Text = 'Ajustar manualment'
        $btnAdj.Location = New-Object System.Drawing.Point(185, 405)
        $btnAdj.Size = New-Object System.Drawing.Size(160, 30)
        $btnAdj.DialogResult = 'Retry'
        $form.Controls.Add($btnAdj)

        $form.AcceptButton = $btnOk
        $res = $form.ShowDialog()
        if ($res -eq 'OK')    { return [pscustomobject]@{ StartIndex = $detectedStart } }
        if ($res -eq 'Retry') { return (Select-ConclusionCutManually -paraTexts $paraTexts -lastReqEndIndex $lastReqEndIndex) }
        return $null
    }

    # No s'ha detectat cap bloc conegut: avisem i passem al selector manual.
    [System.Windows.Forms.MessageBox]::Show(
        "No s'ha detectat cap bloc de conclusions conegut. Tria manualment a partir d'on s'ha d'esborrar.",
        'Seguiment', 'OK', 'Information') | Out-Null
    return (Select-ConclusionCutManually -paraTexts $paraTexts -lastReqEndIndex $lastReqEndIndex)
}

# Selector manual del primer paragraf a esborrar (per a informes atipics).
# Retorna { StartIndex } o { StartIndex = -1 } (no esborrar res); $null si cancel·la.
function Select-ConclusionCutManually {
    param($paraTexts, [int]$lastReqEndIndex)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Tria el primer paragraf a esborrar'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Tria el PRIMER paragraf del bloc de conclusions a esborrar (s'esborrara fins al final):"
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(680, 20)
    $form.Controls.Add($lbl)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(15, 35)
    $list.Size = New-Object System.Drawing.Size(680, 360)
    # Mapatge posicio-de-la-llista -> index 1-based de paragraf.
    $map = @()
    [void]$list.Items.Add('(No esborrar res)')
    $map += -1
    for ($i = $lastReqEndIndex + 1; $i -le $paraTexts.Count; $i++) {
        $tx = [string]$paraTexts[$i - 1]
        if ([string]::IsNullOrWhiteSpace($tx)) { continue }
        $disp = if ($tx.Length -gt 90) { $tx.Substring(0, 90) + '...' } else { $tx }
        [void]$list.Items.Add(('#{0}: {1}' -f $i, $disp))
        $map += $i
    }
    $list.SelectedIndex = if ($list.Items.Count -gt 1) { 1 } else { 0 }
    $form.Controls.Add($list)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Continuar'
    $btnOk.Location = New-Object System.Drawing.Point(600, 405)
    $btnOk.Size = New-Object System.Drawing.Size(95, 30)
    $btnOk.DialogResult = 'OK'
    $form.AcceptButton = $btnOk
    $form.Controls.Add($btnOk)

    $res = $form.ShowDialog()
    if ($res -ne 'OK') { return $null }
    return [pscustomobject]@{ StartIndex = [int]$map[$list.SelectedIndex] }
}

# Demana la data de la ronda (per defecte avui). Retorna { Nav; Data } amb Data
# en dd/MM/yyyy.
function Prompt-RoundDate {
    param($preset = $null)
    $default = if ($preset) { [string]$preset } else { (Get-Date).ToString('dd/MM/yyyy') }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data del seguiment'
    $form.Size = New-Object System.Drawing.Size(420, 200)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Data d'aquesta entrega/seguiment (dd/MM/aaaa):"
    $lbl.Location = New-Object System.Drawing.Point(15, 20)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(15, 50)
    $tb.Size = New-Object System.Drawing.Size(200, 24)
    $tb.Text = $default
    $form.Controls.Add($tb)

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 110)
    $back.Size = New-Object System.Drawing.Size(90, 28)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(295, 110)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    while ($true) {
        $res = $form.ShowDialog()
        if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
        if ($res -ne 'OK')    { exit 0 }
        $v = _ValidateRoundDate $tb.Text
        if ($v.Ok) { return [pscustomobject]@{ Nav='next'; Data=$v.Normalized } }
        [System.Windows.Forms.MessageBox]::Show('Data no valida. Format esperat: dd/MM/aaaa.', 'Seguiment', 'OK', 'Warning') | Out-Null
    }
}

# Pas de comentaris: per cada requeriment, mostra el text + historial, un quadre
# de comentari nou i un checkbox "Resolt". Retorna { Nav; Data } amb Data un
# array alineat amb $requirements: { Resolved; NewComment }.
function Prompt-SeguimentComments {
    param($requirements, $dateStr, $preload = $null)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Seguiment dels requeriments'
    $form.Size = New-Object System.Drawing.Size(820, 620)
    $form.StartPosition = 'CenterScreen'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = ("Anotacio del {0}. Escriu el comentari de cada requeriment i marca 'Resolt' si ha quedat resolt." -f $dateStr)
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.Size = New-Object System.Drawing.Size(770, 20)
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 35)
    $panel.Size = New-Object System.Drawing.Size(775, 490)
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panel)

    $rows = @()
    $y = 8
    for ($i = 0; $i -lt $requirements.Count; $i++) {
        $req = $requirements[$i]

        $reqLbl = New-Object System.Windows.Forms.Label
        $reqLbl.Text = [string]$req.Text
        $reqLbl.Location = New-Object System.Drawing.Point(8, $y)
        $reqLbl.Size = New-Object System.Drawing.Size(740, 38)
        $reqLbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $panel.Controls.Add($reqLbl)
        $y += 40

        if ($req.Annotations.Count -gt 0) {
            $hist = ($req.Annotations | ForEach-Object { ('{0}: {1}' -f $_.Date, $_.Text) }) -join "`r`n"
            $histLbl = New-Object System.Windows.Forms.Label
            $histLbl.Text = $hist
            $histLbl.Location = New-Object System.Drawing.Point(20, $y)
            $histLbl.Size = New-Object System.Drawing.Size(728, (16 * $req.Annotations.Count + 2))
            $histLbl.ForeColor = [System.Drawing.Color]::DimGray
            $panel.Controls.Add($histLbl)
            $y += (16 * $req.Annotations.Count + 6)
        }

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.ScrollBars = 'Vertical'
        $tb.Location = New-Object System.Drawing.Point(20, $y)
        $tb.Size = New-Object System.Drawing.Size(728, 44)
        $panel.Controls.Add($tb)
        $y += 50

        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = 'Resolt'
        $cb.Location = New-Object System.Drawing.Point(20, $y)
        $cb.Size = New-Object System.Drawing.Size(120, 22)
        $cb.Checked = [bool]$req.WasResolved
        $panel.Controls.Add($cb)
        $y += 30

        # Separador visual
        $sep = New-Object System.Windows.Forms.Label
        $sep.BorderStyle = 'Fixed3D'
        $sep.Location = New-Object System.Drawing.Point(8, $y)
        $sep.Size = New-Object System.Drawing.Size(740, 2)
        $panel.Controls.Add($sep)
        $y += 12

        # Precarrega (tornar enrere)
        if ($null -ne $preload -and $i -lt $preload.Count) {
            $tb.Text    = [string]$preload[$i].NewComment
            $cb.Checked = [bool]$preload[$i].Resolved
        }

        $rows += [pscustomobject]@{ Comment=$tb; Resolved=$cb }
    }

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, 535)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Seguent'
    $ok.Location = New-Object System.Drawing.Point(695, 535)
    $ok.Size = New-Object System.Drawing.Size(95, 30)
    $ok.DialogResult = 'OK'
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    $decisions = New-Object System.Collections.ArrayList
    foreach ($r in $rows) {
        [void]$decisions.Add([pscustomobject]@{
            Resolved   = [bool]$r.Resolved.Checked
            NewComment = [string]$r.Comment.Text
        })
    }
    return [pscustomobject]@{ Nav='next'; Data=$decisions.ToArray() }
}

# ----------------------------------------------------------------------------
# Orquestrador del flux de seguiment.
# ----------------------------------------------------------------------------
function Invoke-SeguimentFlow {
    $sourcePath = Select-PreviousReport
    if (-not $sourcePath) { return }

    # Llista de frases que marquen l'inici del bloc de conclusions. Es pot
    # sobreescriure des de config.ps1 ($SeguimentConclusionPhrases).
    $phrases = if ($null -ne $SeguimentConclusionPhrases) { $SeguimentConclusionPhrases } else {
        @("Vist l'anterior", 'Ho poso al seu coneixement', 'Cornella de Llobregat,')
    }

    $word = New-WordApp
    try {
        $prev = Read-PreviousReport -word $word -path $sourcePath
        if ($prev.Model.Requirements.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No s'han trobat requeriments numerats (1., 2., 3...) en aquest document.",
                'Seguiment', 'OK', 'Warning') | Out-Null
            return
        }

        $detected = _FindConclusionStartIndex $prev.ParaTexts $prev.Model.LastReqParaIndex $phrases
        $cut = Confirm-ConclusionDeletion -paraTexts $prev.ParaTexts -lastReqEndIndex $prev.Model.LastReqParaIndex -detectedStart $detected
        if ($null -eq $cut) { return }
        $conclusionStartIndex = [int]$cut.StartIndex

        $conclAll = Read-Conclusions -word $word -path $ConclusionsPath

        # Maquina de passos: 1=data, 2=comentaris, 3=conclusions, 4=camps. 5=fi.
        $st  = @{ Date=$null; Decisions=$null; Conclusions=@(); Fields=$null }
        $pre = @{ Date=$null; ConclTitles=$null }
        $step = 1
        while ($step -ge 1 -and $step -le 4) {
            switch ($step) {
                1 {
                    $r = Prompt-RoundDate -preset $pre.Date
                    if ($r.Nav -eq 'back') { return }   # enrere des del primer pas = sortir
                    $st.Date = $r.Data; $pre.Date = $r.Data; $step = 2
                }
                2 {
                    $r = Prompt-SeguimentComments -requirements $prev.Model.Requirements -dateStr $st.Date -preload $st.Decisions
                    if ($r.Nav -eq 'back') { $step = 1 }
                    else { $st.Decisions = $r.Data; $step = 3 }
                }
                3 {
                    if ($conclAll.Selectable.Count -eq 0) {
                        $st.Conclusions = @()
                        $step = 4
                    } else {
                        $r = Select-Conclusions -conclusions $conclAll.Selectable -preloadTitles $pre.ConclTitles
                        if ($r.Nav -eq 'back') { $step = 2 }
                        else {
                            $st.Conclusions = $r.Data
                            $pre.ConclTitles = @($st.Conclusions | ForEach-Object { $_.Title })
                            $step = 4
                        }
                    }
                }
                4 {
                    $fields = [ordered]@{}
                    Add-FieldsFromConclusions $fields $st.Conclusions $conclAll.Always
                    if ($fields.Count -eq 0) {
                        $st.Fields = $fields; $step = 5
                    } else {
                        $r = Prompt-Fields -fields $fields
                        if ($r.Nav -eq 'back') { $step = 3 }
                        else { $st.Fields = $r.Data; $step = 5 }
                    }
                }
            }
        }

        $roundDt    = [datetime]::ParseExact($st.Date, 'dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
        $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)

        $outPath = Apply-Seguiment -word $word -doc $prev.Doc -tempPath $prev.TempPath `
                       -model $prev.Model -conclusionStartIndex $conclusionStartIndex `
                       -decisions $st.Decisions -dateStr $st.Date `
                       -conclHeaderText $conclAll.HeaderText -selectedConclusions $st.Conclusions `
                       -alwaysConclusions $conclAll.Always -fields $st.Fields `
                       -sourceBaseName $sourceBase -roundDate $roundDt

        [System.Windows.Forms.MessageBox]::Show(
            "Informe de seguiment generat:`n$outPath", 'Finalitzat', 'OK', 'Information') | Out-Null

        $word.Visible = $true
        $word.Documents.Open($outPath) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        throw
    }
    finally {
        if (-not $word.Visible) { Close-WordApp $word }
    }
}
