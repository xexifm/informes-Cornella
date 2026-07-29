#requires -Version 5.1
<#
.SYNOPSIS
  Pas 2 de l'assistent: el formulari de la CAPCALERA de l'informe.

.DESCRIPTION
  Munta la pantalla amb les dades de l'activitat (ID GIA, expedient, adreca,
  titular, activitat, dates...), l'omple sola amb la cache de l'Excel
  (Activitats.ps1) quan escrius un ID GIA conegut, i en retorna els valors.
  Tambe hi ha el boto "Recuperar dades ultim informe".

  Ve de Motor.ps1 (Step 2). Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Step 2 - Header data (formulari + precarrega Excel)
# ----------------------------------------------------------------------------
# Construeix el formulari de capcalera (controls + botons), retorna la
# tupla amb el formulari, el diccionari de controls i el boto Cercar perque
# Get-HeaderData hi puga lligar la logica de cerca i validacio.
function _BuildHeaderForm($excelInfo) {
    # $excelInfo: hashtable amb claus Text (string) i Source ('primary'|'fallback').
    # Quan es 'fallback', el rotul es taronja i porta el text "[FALLBACK LOCAL]"
    # davant perque l'usuari sapiga que NO esta usant la base de dades oficial
    # de la xarxa.
    $form = _NewForm
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(720, 620)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda granat (44) + barra de passos (34):
    # tots els controls absoluts es desplacen cap avall aquest offset.
    $topOffset = 74

    $lblBd = New-Object System.Windows.Forms.Label
    $lblBd.Location = New-Object System.Drawing.Point(15, (12 + $topOffset))
    $lblBd.Size = New-Object System.Drawing.Size(680, 22)
    if ($excelInfo.Source -eq 'fallback') {
        $lblBd.Text = "[FALLBACK LOCAL]  " + $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblBd.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    } else {
        $lblBd.Text = $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    $form.Controls.Add($lblBd)

    $controls  = @{}
    $rowLabels = @{}
    $addRow = {
        param($label, $y, $tbWidth, $key)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $label
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(200, 22)
        [void]$form.Controls.Add($lbl)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(220, ($y - 2))
        $tb.Size = New-Object System.Drawing.Size($tbWidth, 22)
        [void]$form.Controls.Add($tb)
        $controls[$key]  = $tb
        $rowLabels[$key] = $lbl
    }

    $y = 50 + $topOffset
    & $addRow 'ID GIA' $y 380 'ID_GIA'
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = 'Cercar'
    $btnSearch.Location = New-Object System.Drawing.Point(605, ($y - 4))
    $btnSearch.Size = New-Object System.Drawing.Size(80, 28)
    _StylePrimaryButton $btnSearch
    [void]$form.Controls.Add($btnSearch)
    $y += 38

    & $addRow "Num. d'expedient (autom., editable)"   $y 460 'EXP_NUM';   $y += 38
    & $addRow 'Titular (autom., editable)'            $y 460 'TITULAR';   $y += 38
    & $addRow 'Adreca (autom., editable)'             $y 460 'ADRECA';    $y += 38
    & $addRow 'Activitat (autom., editable)'          $y 460 'ACTIVITAT'; $y += 38

    # --- Origen de l'informe: documentacio aportada o visita d'inspeccio ---
    # Segons la tria es mostren uns camps o uns altres i canvia la linia
    # "Objecte:" de la capcalera (placeholder <<ORIGEN>>, muntat a
    # Apply-HeaderReplacements).
    $lblOrigen = New-Object System.Windows.Forms.Label
    $lblOrigen.Text = "Origen de l'informe:"
    $lblOrigen.Location = New-Object System.Drawing.Point(15, $y)
    $lblOrigen.Size = New-Object System.Drawing.Size(200, 22)
    [void]$form.Controls.Add($lblOrigen)

    $rbDoc = New-Object System.Windows.Forms.RadioButton
    $rbDoc.Text = 'Documentacio aportada'
    $rbDoc.Location = New-Object System.Drawing.Point(220, ($y - 2))
    $rbDoc.AutoSize = $true
    $rbDoc.Checked = $true
    [void]$form.Controls.Add($rbDoc)

    $rbInsp = New-Object System.Windows.Forms.RadioButton
    $rbInsp.Text = "Visita d'inspeccio"
    $rbInsp.Location = New-Object System.Drawing.Point(440, ($y - 2))
    $rbInsp.AutoSize = $true
    [void]$form.Controls.Add($rbInsp)
    $controls['_ORIGEN_DOC']  = $rbDoc
    $controls['_ORIGEN_INSP'] = $rbInsp
    $y += 34

    # Camps de documentacio aportada (NUM/DATA anotacio) i, superposat al primer,
    # el camp de data d'inspeccio: nomes es veu el joc que toca segons la tria.
    $yAnot = $y
    & $addRow "Num. d'anotacio (autom., editable)" $yAnot 460 'NUM_ANOTACIO'
    & $addRow "Data d'inspeccio"                    $yAnot 460 'DATA_INSPECCIO'
    $y = $yAnot + 38
    & $addRow "Data d'anotacio (autom., editable)"  $y 460 'DATA_ANOTACIO'; $y += 50

    # Mostra/amaga els camps segons l'origen triat.
    $applyOrigen = {
        $isDoc = $rbDoc.Checked
        foreach ($k in 'NUM_ANOTACIO', 'DATA_ANOTACIO') {
            $controls[$k].Visible  = $isDoc
            $rowLabels[$k].Visible = $isDoc
        }
        $controls['DATA_INSPECCIO'].Visible  = -not $isDoc
        $rowLabels['DATA_INSPECCIO'].Visible = -not $isDoc
    }.GetNewClosure()
    $rbDoc.add_CheckedChanged($applyOrigen)
    & $applyOrigen

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    _StyleSecondaryButton $back
    [void]$form.Controls.Add($back)

    $recover = New-Object System.Windows.Forms.Button
    $recover.Text = ([char]0x21BA + " Recuperar dades ultim informe")
    $recover.Location = New-Object System.Drawing.Point(115, $y)
    $recover.Size = New-Object System.Drawing.Size(250, 30)
    _StyleSecondaryButton $recover
    [void]$form.Controls.Add($recover)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Seg' + [char]0x00FC + 'ent ' + [char]0x2192)   # Següent →
    $ok.Location = New-Object System.Drawing.Point(575, $y)
    $ok.Size = New-Object System.Drawing.Size(110, 30)
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    # Barra de passos (pas 2 actiu) + banda granat. La banda s'afegeix DESPRES
    # per quedar a dalt de tot; les dues son Dock=Top.
    [void](_AddStepBar $form 2)
    [void](_AddBrandHeader $form "Dades de l'activitat" $null 44)

    return @{ Form=$form; Controls=$controls; BtnSearch=$btnSearch; BtnOk=$ok; BtnBack=$back; BtnRecover=$recover }
}

# Llegeix els valors dels controls i retorna un hashtable amb la capcalera.
function _ReadHeaderControls($controls) {
    $tipus = if ($controls['_ORIGEN_INSP'] -and $controls['_ORIGEN_INSP'].Checked) { 'insp' } else { 'doc' }
    @{
        ID_GIA         = $controls['ID_GIA'].Text.Trim()
        EXP_NUM        = $controls['EXP_NUM'].Text.Trim()
        TITULAR        = $controls['TITULAR'].Text.Trim()
        ADRECA         = $controls['ADRECA'].Text.Trim()
        ACTIVITAT      = $controls['ACTIVITAT'].Text.Trim()
        ORIGEN_TIPUS   = $tipus
        NUM_ANOTACIO   = $controls['NUM_ANOTACIO'].Text.Trim()
        DATA_ANOTACIO  = $controls['DATA_ANOTACIO'].Text.Trim()
        DATA_INSPECCIO = $controls['DATA_INSPECCIO'].Text.Trim()
    }
}

# Precarrega valors d'una capcalera anterior als controls del formulari.
# $preload pot ser un hashtable (navegacio en memoria) o un PSCustomObject
# (dades de l'ultim informe llegides de JSON).
function _PreloadHeaderControls($controls, $preload) {
    if ($null -eq $preload) { return }
    # Helper per llegir una clau tant si $preload es hashtable com PSCustomObject.
    $getVal = {
        param($k)
        if ($preload -is [System.Collections.IDictionary]) {
            if ($preload.Contains($k)) { return $preload[$k] }
        } elseif ($preload.PSObject.Properties.Name -contains $k) {
            return $preload.$k
        }
        return $null
    }
    foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','NUM_ANOTACIO','DATA_ANOTACIO','DATA_INSPECCIO') {
        $v = & $getVal $k
        if ($null -ne $v) { $controls[$k].Text = [string]$v }
    }
    # Restaura l'origen triat (documentacio aportada / visita d'inspeccio).
    $tipus = & $getVal 'ORIGEN_TIPUS'
    if ($controls['_ORIGEN_DOC'] -and $controls['_ORIGEN_INSP']) {
        if ([string]$tipus -eq 'insp') { $controls['_ORIGEN_INSP'].Checked = $true }
        else                           { $controls['_ORIGEN_DOC'].Checked  = $true }
    }
}

function Get-HeaderData {
    param($preload = $null)

    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' a cap de les ubicacions:`n`n" +
            "  1. $ActivitatsDir`n" +
            "  2. $LocalActivitatsDir  (fallback local)`n`n" +
            "Si estas fora de la feina, copia una base de dades a la carpeta`n" +
            "'local\base-dades-activitats' dins de la carpeta del programa.",
            'Base de dades no trobada', 'OK', 'Error') | Out-Null
        exit 1
    }
    # RENDIMENT: obrir l'Excel i llegir tota la base es la part mes lenta del
    # Pas 2. Ho fem UNA sola vegada per sessio (i per fitxer): si es torna al
    # Pas 2 (p.ex. Enrere des del Pas 3), reaprofitem la cache ja carregada i
    # NO reobrim l'Excel ni refem l'export a Drive. Aixi el pas a pas es immediat.
    $cacheKey = '{0}|{1}|{2}' -f $latest.File.FullName, $latest.File.Length, $latest.File.LastWriteTimeUtc.Ticks
    if ($script:_sessionActKey -eq $cacheKey -and $null -ne $script:_sessionActCache) {
        $actCache = $script:_sessionActCache
    } else {
        if ($latest.Source -eq 'fallback') {
            # Avis explicit (un cop) perque l'usuari sapiga que treballa amb una
            # copia local i no amb la xarxa de la feina.
            [System.Windows.Forms.MessageBox]::Show(
                "No s'ha trobat la base de dades a la xarxa.`n`n" +
                "S'usa la copia LOCAL del clone:`n  $($latest.File.FullName)`n`n" +
                "Comprova que sigui prou recent.",
                'Base de dades: copia local (fallback)', 'OK', 'Warning') | Out-Null
        }
        # Precarrega TOTA la base de dades a memoria. A partir d'aqui les cerques
        # son immediates (no cal reobrir Excel).
        try {
            $actCache = Initialize-ActivitatsCache -excelFile $latest.File
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error llegint l'Excel:`n$($_.Exception.Message)",'Error','OK','Error') | Out-Null
            exit 1
        }
        if ($actCache.Warnings -and $actCache.Warnings.Count -gt 0) {
            $msg = "Avisos de validacio de l'Excel (l'auto-fill podria fallar):`n`n" + ($actCache.Warnings -join "`n")
            [System.Windows.Forms.MessageBox]::Show($msg,'Avisos','OK','Warning') | Out-Null
        }
        $script:_sessionActCache = $actCache
        $script:_sessionActKey   = $cacheKey
    }

    # NOTA: NO refresquem la copia d'activitats al Drive aqui. Comprovar-ho i
    # pujar-ho feia que generar un informe (Pas 2) trigues molt (accedia a la
    # xarxa/Drive cada cop). Aquesta sincronitzacio per al mobil ara nomes es fa
    # amb Actualitzar.bat (ExportaDades.ps1 -> Export-ActivitatsCmd), aixi
    # generar informes es rapid i no depen de la xarxa.

    $labelText = "Base de dades d'activitats: $($latest.File.Name)  (data: $($latest.Date.ToString('yyyy-MM-dd'))) - $($actCache.ById.Count) activitats carregades"
    $f = _BuildHeaderForm @{ Text = $labelText; Source = $latest.Source }
    $form       = $f.Form
    $controls   = $f.Controls
    $btnSearch  = $f.BtnSearch
    $ok         = $f.BtnOk
    $btnRecover = $f.BtnRecover

    _PreloadHeaderControls $controls $preload

    # Boto "Recuperar dades ultim informe": carrega les dades de l'ultim
    # informe generat amb exit i les deixa als formularis perque l'usuari
    # les revisi/modifiqui pas per pas. La resta de passos (seleccio, camps,
    # conclusions) es precarreguen via $script:_recoveredReport, que Main llegeix.
    $script:_recoveredReport = $null
    $btnRecover.add_Click({
        $rep = Load-LastReport
        if ($null -eq $rep) {
            [System.Windows.Forms.MessageBox]::Show("No hi ha cap informe anterior desat.",'Sense dades','OK','Information') | Out-Null
            return
        }
        if ($rep.Header) { _PreloadHeaderControls $controls $rep.Header }
        $script:_recoveredReport = $rep
        [System.Windows.Forms.MessageBox]::Show(
            "Dades de l'ultim informe carregades.`n`nRevisa-les i modifica el que calgui (per exemple, canvia l'ID GIA i prem 'Cercar' per a una activitat nova). En continuar, els passos seguents tambe sortiran precarregats.",
            'Recuperat', 'OK', 'Information') | Out-Null
    })

    # Cerca per ID GIA: instantania des del cache. Omple els camps automatics.
    $doSearch = {
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return $false
        }
        $act = Get-ActivitatFromCache $actCache $idGia
        if ($null -eq $act) {
            [System.Windows.Forms.MessageBox]::Show(
                "L'ID GIA '$idGia' no s'ha trobat a la base de dades`n($($latest.File.Name)).",
                'Activitat no trobada', 'OK', 'Error') | Out-Null
            foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
                $controls[$k].Text = ''
            }
            return $false
        }
        foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
            if ($act.ContainsKey($k)) { $controls[$k].Text = [string]$act[$k] }
        }
        return $true
    }

    $btnSearch.add_Click({ [void](& $doSearch) })

    $script:_headerData = $null
    $ok.add_Click({
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falten dades','OK','Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($controls['TITULAR'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ADRECA'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ACTIVITAT'].Text)) {
            if (-not (& $doSearch)) { return }
        }
        $script:_headerData = _ReadHeaderControls $controls
        $form.DialogResult = 'OK'
        $form.Close()
    })

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }
    return [pscustomobject]@{ Nav='next'; Data=$script:_headerData; Recovered=$script:_recoveredReport }
}
