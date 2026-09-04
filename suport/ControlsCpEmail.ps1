#requires -Version 5.1
<#
.SYNOPSIS
  Avisos de CONTROL PERIODIC per correu (esborranys a Outlook) des de l'eina
  "Controls periodics".

.DESCRIPTION
  Per a les activitats SELECCIONADES a la graella de Controls periodics, crea un
  correu per titular avisant que constava un control periodic a passar (en data
  X) per la seva activitat X situada a X. Els correus i les dades surten de la
  base d'activitats (Excel) que ja llegeix _ReadControlsPeriodics; el programa
  els deixa com a ESBORRANYS a Outlook (mai els envia sol) perque l'usuari els
  revisi i envii.

  El text es EDITABLE (assumpte + cos amb variables), com els "Textos del correu"
  del mobil, pero es guarda LOCALMENT (%LOCALAPPDATA%\InformesCornella\
  controls-cp-email.json), fora del repositori: no conte cap dada personal i
  sobreviu a Actualitzar.bat (git pull) sense conflictes.

  Variables del text:
    {ACTIVITAT} {ADRECA} {ID_GIA} {TITULAR}   dades de l'activitat
    {PROPER_CP}      Proper CP previst (la data en que tocava el control)
    {DATA_CONTROL}   Data de l'ultim control periodic registrat
    {DATA}           data d'avui (dd/MM/yyyy)
  El cos admet **negreta** i els enllacos http(s) es fan clicables.

  Funcions PURES (plantilla, destinataris, substitucio de variables, HTML)
  testejables en headless; Outlook (COM) i les finestres (WinForms) nomes a
  Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# Ruta del fitxer d'overrides local (mai al repositori; sense dades personals).
function _ControlsCpEmailPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path $base (Join-Path 'InformesCornella' 'controls-cp-email.json'))
}

# Valors PER DEFECTE (assumpte + cos, bilingue CA/ES, cordial).
function _DefaultControlsCpEmail {
    $cos = @(
        'Activitat: {ACTIVITAT}'
        'Adreça: {ADRECA}'
        'ID GIA: {ID_GIA}'
        'Titular: {TITULAR}'
        ''
        '**Català**'
        "Benvolgut/da,"
        "Segons les nostres dades, la vostra activitat situada a {ADRECA} havia de passar un **control periòdic** amb data prevista **{PROPER_CP}**, i a hores d'ara no ens consta que s'hagi dut a terme."
        "Us recordem que aquest control és una obligació periòdica de l'activitat. Us demanem que, com abans millor, encarregueu el control a una entitat/organisme de control habilitat i que ens en presenteu l'acta favorable."
        "Podeu presentar la documentació mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat, a l'atenció del **Departament d'Activitats**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        '**Castellano**'
        "Estimado/a,"
        "Según nuestros datos, su actividad situada en {ADRECA} debía pasar un **control periódico** con fecha prevista **{PROPER_CP}**, y a día de hoy no nos consta que se haya realizado."
        "Le recordamos que este control es una obligación periódica de la actividad. Le pedimos que, cuanto antes, encargue el control a una entidad/organismo de control habilitado y nos presente el acta favorable."
        "Puede presentar la documentación mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat, a la atención del **Departamento de Actividades**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        '________________________________________'
        ''
        "Departament d'Activitats · Ajuntament de Cornellà de Llobregat · Carrer de l'Energia, 97 · Tel. 93 377 02 12 (ext. 1227)"
        "IMPORTANT: aquest és un correu automàtic. Per a qualsevol consulta, adreceu-vos al Departament d'Activitats. / IMPORTANTE: este es un correo automático. Para cualquier consulta, diríjase al Departamento de Actividades."
    ) -join "`n"
    $d = [ordered]@{
        assumpte = 'Control periòdic pendent · GIA {ID_GIA}'
        cos      = $cos
    }
    return ,$d
}

# Text d'ajuda amb les variables disponibles.
function _ControlsCpEmailAjuda {
    return ('Variables: {ACTIVITAT} {ADRECA} {ID_GIA} {TITULAR} {PROPER_CP} {DATA_CONTROL} {DATA}   ' + [char]0x00B7 + '   **negreta**   ' + [char]0x00B7 + '   //cursiva//   ' + [char]0x00B7 + '   els enllaços http es fan clicables')
}

# Llegeix el JSON local (si hi es) fusionat sobre els valors per defecte.
function _LoadControlsCpEmail {
    $def = _DefaultControlsCpEmail
    $path = _ControlsCpEmailPath
    $o = Read-JsonFile $path
    if ($null -ne $o) {
        try {
            foreach ($k in @($def.Keys)) {
                if ($o.PSObject.Properties[$k] -and -not [string]::IsNullOrEmpty([string]$o.$k)) {
                    $def[$k] = [string]$o.$k
                }
            }
        } catch { }
    }
    return ,$def
}

# Escriu el JSON local (UTF-8 sense BOM), creant la carpeta si cal.
function _SaveControlsCpEmail($obj) {
    $path = _ControlsCpEmailPath
    Write-JsonFile $path $obj 5
}

# Munta els destinataris: titular a "Per a" (To), representant a "CC". Si en
# falta un, l'altre passa a To. Valida que continguin '@'. Ok=$false si cap.
function _ControlsCpRecipients([string]$raoEmail, [string]$repEmail) {
    $rao = ([string]$raoEmail).Trim()
    $rep = ([string]$repEmail).Trim()
    $raoOk = ($rao -like '*@*')
    $repOk = ($rep -like '*@*')
    $to = ''; $cc = ''
    if ($raoOk -and $repOk) { $to = $rao; $cc = $rep }
    elseif ($raoOk) { $to = $rao }
    elseif ($repOk) { $to = $rep }
    return @{ To = $to; Cc = $cc; Ok = [bool]($to -ne '') }
}

# Substitueix les variables del text amb les dades d'una fila d'activitat.
# El MAPA es d'aqui (les claus i d'on surten els valors son d'aquesta eina); el
# bucle el fa _OmpleVariables (EnviarCorreu.ps1), que estava copiat tres cops.
function _FillControlsCpPh([string]$text, $row) {
    return (_OmpleVariables $text ([ordered]@{
        '{ACTIVITAT}'    = [string]$row.ActPrincipal
        '{ADRECA}'       = [string]$row.Adreca
        '{ID_GIA}'       = [string]$row.Id
        '{TITULAR}'      = [string]$row.RaoSocial
        '{PROPER_CP}'    = [string]$row.ProperCP
        '{DATA_CONTROL}' = [string]$row.DataControlPer
        '{DATA}'         = (Get-Date).ToString('dd/MM/yyyy')
    }))
}

# L'HTML del cos el fa _CosAHtml (EnviarCorreu.ps1). Aqui hi havia
# _ControlsCpEmailHtml, IDENTICA linia a linia a _RecCosHtml dels recordatoris
# -mateix estil inline inclos-, i un _ControlsCpLineHtml propi que feia el mateix
# que _TextToHtml pero SENSE cursiva. Ara aquest correu tambe accepta //cursiva//.

# ----------------------------------------------------------------------------
# OUTLOOK (COM) - nomes a Windows. Crea ESBORRANYS; MAI envia.
# ----------------------------------------------------------------------------
function Invoke-ControlsCpEmailDrafts($rows) {
    $rows = @($rows)
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Marca almenys una activitat (columna 'Generar') per preparar-ne el correu.", 'Enviar correu', 'OK', 'Information') | Out-Null
        return
    }

    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("Es prepararan $($rows.Count) correus (un per activitat triada) com a ESBORRANYS a Outlook.`n`n" +
         "Titular a 'Per a' i representant a 'CC' (quan hi siguin). NO s'envia res: els revisaràs i enviaràs tu des d'Outlook.`n`nVols continuar?"),
        'Enviar correu', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $textos = _LoadControlsCpEmail
    $assTpl = [string]$textos['assumpte']
    $cosTpl = [string]$textos['cos']

    # Finestra de progres amb Cancel·lar (mateix patro que Generar informes).
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Enviar correu'
    $form.Size = New-Object System.Drawing.Size(560, 170)
    $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18); $lbl.Size = New-Object System.Drawing.Size(510, 40)
    $lbl.Text = 'Preparant...'
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 66); $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $rows.Count)
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'; $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(410, 96)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({ param($s, $e) if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true } }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $ok = 0; $senseCorreu = New-Object System.Collections.ArrayList
    $err = 0; $errDetalls = New-Object System.Collections.ArrayList; $cancelled = $false
    $outlook = $null
    try {
        try { $outlook = New-Object -ComObject Outlook.Application } catch { $outlook = $null }
        if ($null -eq $outlook) {
            try { $form.Close() } catch { }
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut iniciar Microsoft Outlook.", 'Enviar correu', 'OK', 'Error') | Out-Null
            return
        }

        $done = 0
        foreach ($r in $rows) {
            if ($cancel.Flag) { $cancelled = $true; break }
            $done++
            $lbl.Text = "Preparant esborranys...  $done de $($rows.Count)`nGIA $($r.Id) - $($r.RaoSocial)"
            if ($bar.Value -lt $bar.Maximum) { $bar.Value = $done }
            [System.Windows.Forms.Application]::DoEvents()

            $rec = _ControlsCpRecipients $r.RaoEmail $r.RepEmail
            if (-not $rec.Ok) { [void]$senseCorreu.Add("GIA $($r.Id) - $($r.RaoSocial)"); continue }

            $mail = $null
            try {
                $mail = $outlook.CreateItem(0)   # olMailItem
                $mail.To = $rec.To
                if ($rec.Cc) { $mail.CC = $rec.Cc }
                $mail.Subject = (_FillControlsCpPh $assTpl $r)
                $mail.HTMLBody = (_CosAHtml (_FillControlsCpPh $cosTpl $r))
                $mail.Save()   # queda a Esborranys; MAI Send()
                $ok++
            } catch {
                $err++; [void]$errDetalls.Add("GIA $($r.Id): $($_.Exception.Message)")
            } finally {
                if ($null -ne $mail) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null } catch { } }
            }
        }
    } finally {
        $cancel.Running = $false
        try { $form.Close() } catch { }
        # NO fem $outlook.Quit() (podria tancar l'Outlook de l'usuari); nomes alliberem.
        if ($null -ne $outlook) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null } catch { } }
    }

    $titol = if ($cancelled) { 'Preparació cancel·lada' } else { 'Esborranys preparats' }
    $msg = "$titol`n`nEsborranys creats a Outlook: $ok"
    if ($senseCorreu.Count -gt 0) {
        $msg += "`n`nActivitats SENSE correu (omeses): $($senseCorreu.Count)`n - " + (($senseCorreu | Select-Object -First 15) -join "`n - ")
    }
    if ($err -gt 0) { $msg += "`n`nErrors: $err`n - " + (($errDetalls | Select-Object -First 10) -join "`n - ") }
    $msg += "`n`nRevisa'ls a Outlook (carpeta Esborranys) abans d'enviar-los."
    [System.Windows.Forms.MessageBox]::Show($msg, 'Enviar correu', 'OK', 'Information') | Out-Null
}

# ----------------------------------------------------------------------------
# EDITOR del text (WinForms) - nomes a Windows
# ----------------------------------------------------------------------------
function Invoke-ControlsCpEmailTextos {
    $textos = _LoadControlsCpEmail

    $form = _NewForm
    $form.Text = 'Text del correu (controls periodics)'
    $form.ClientSize = New-Object System.Drawing.Size(760, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 470)

    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = 'Assumpte del correu:'
    $lblA.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblA.Location = New-Object System.Drawing.Point(16, 70)
    $lblA.AutoSize = $true
    [void]$form.Controls.Add($lblA)

    $tbA = New-Object System.Windows.Forms.TextBox
    $tbA.Location = New-Object System.Drawing.Point(16, 92)
    $tbA.Size = New-Object System.Drawing.Size(728, 24)
    $tbA.Anchor = 'Top, Left, Right'
    $tbA.Text = [string]$textos['assumpte']
    [void]$form.Controls.Add($tbA)

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Cos del correu:'
    $lblC.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblC.Location = New-Object System.Drawing.Point(16, 126)
    $lblC.AutoSize = $true
    [void]$form.Controls.Add($lblC)

    $lblH = New-Object System.Windows.Forms.Label
    $lblH.Text = _ControlsCpEmailAjuda
    $lblH.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
    $lblH.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lblH.Location = New-Object System.Drawing.Point(16, 146)
    $lblH.AutoSize = $false
    $lblH.Size = New-Object System.Drawing.Size(728, 16)
    $lblH.Anchor = 'Top, Left, Right'
    [void]$form.Controls.Add($lblH)

    $tbC = New-Object System.Windows.Forms.TextBox
    $tbC.Location = New-Object System.Drawing.Point(16, 166)
    $tbC.Size = New-Object System.Drawing.Size(728, 396)
    $tbC.Anchor = 'Top, Bottom, Left, Right'
    $tbC.Multiline = $true
    $tbC.ScrollBars = 'Vertical'
    $tbC.WordWrap = $true
    $tbC.AcceptsReturn = $true
    $tbC.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    # El TextBox multilinia nomes mostra CRLF; el cos es guarda amb LF.
    $tbC.Text = [string]$textos['cos'] -replace "`r?`n", "`r`n"
    [void]$form.Controls.Add($tbC)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(16, 578)
    $btnBack.Size = New-Object System.Drawing.Size(110, 30)
    $btnBack.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnBack
    [void]$form.Controls.Add($btnBack)

    $btnDefault = New-Object System.Windows.Forms.Button
    $btnDefault.Text = 'Restaurar original'
    $btnDefault.Location = New-Object System.Drawing.Point(136, 578)
    $btnDefault.Size = New-Object System.Drawing.Size(150, 30)
    $btnDefault.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnDefault
    [void]$form.Controls.Add($btnDefault)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(624, 578)
    $btnSave.Size = New-Object System.Drawing.Size(120, 30)
    $btnSave.Anchor = 'Bottom, Right'
    _StylePrimaryButton $btnSave
    [void]$form.Controls.Add($btnSave)

    [void](_AddBrandHeader $form 'Text del correu' ('Avís de control peri' + [char]0x00F2 + 'dic ' + [char]0x00B7 + ' es desa en aquest ordinador'))

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    $btnDefault.add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show('Vols recuperar el text original? (No es desa fins que premis Desar.)', 'Text del correu', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $def = _DefaultControlsCpEmail
        $tbA.Text = [string]$def['assumpte']
        $tbC.Text = [string]$def['cos'] -replace "`r?`n", "`r`n"
    }.GetNewClosure())

    $btnSave.add_Click({
        $out = [ordered]@{ assumpte = [string]$tbA.Text; cos = ([string]$tbC.Text -replace "`r`n", "`n") }
        try {
            _SaveControlsCpEmail $out
            [System.Windows.Forms.MessageBox]::Show('Text desat en aquest ordinador.', 'Text del correu', 'OK', 'Information') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Text del correu', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    [void]$form.ShowDialog()
    $form.Dispose()
}
