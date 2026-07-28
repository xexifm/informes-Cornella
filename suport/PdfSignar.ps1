#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Convertir informes a PDF (i signar)": passa tots els Word d'una carpeta a
  PDF i, opcionalment, els signa amb AutoFirma (certificat del magatzem de
  Windows).

.DESCRIPTION
  - Converteix cada .doc/.docx d'una carpeta (i subcarpetes) a PDF, al MATEIX
    lloc i amb el mateix nom (informe.docx -> informe.pdf), fent servir Word (COM,
    ExportAsFixedFormat). Salta els que ja tenen un PDF al dia (si no es marca
    "sobreescriure").
  - Si es demana signar, crida AutoFirma per linia de comandes (operacio 'sign',
    format PAdES, magatzem 'windows') per a cada PDF. Un filtre de certificat
    (text del titular, p. ex. el NOM o el NIF) permet que AutoFirma triï el
    certificat sol, sense diàleg. El PDF signat substitueix el sense signar.

  Per què AutoFirma amb el magatzem de Windows? Reutilitza EXACTAMENT el mateix
  certificat que ja fas servir a Windows, és l'eina oficial (la signatura és
  vàlida per a l'administració) i sap fer PAdES correctament. Només cal tenir
  AutoFirma instal·lat (habitual a l'administració).

  Les funcions PURES (rutes, decisió de conversió, arguments d'AutoFirma) son
  testejables en headless; el Word (COM) i AutoFirma només s'executen a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables)
# ----------------------------------------------------------------------------

# Ruta del PDF corresponent a un document (mateixa carpeta, extensió .pdf).
function _PdfPathForDoc([string]$docPath) {
    return [System.IO.Path]::ChangeExtension($docPath, '.pdf')
}

# Decideix si cal (re)generar el PDF. Funció PURA.
#   $overwrite  -> sempre.
#   PDF no existeix -> sí.
#   El Word és més nou que el PDF -> sí (regenerar).
#   Altrament -> no (ja està al dia).
function _PdfShouldConvert([bool]$pdfExists, [datetime]$docTimeUtc, [datetime]$pdfTimeUtc, [bool]$overwrite) {
    if ($overwrite) { return $true }
    if (-not $pdfExists) { return $true }
    return ($docTimeUtc -gt $pdfTimeUtc)
}

# Valor del filtre de certificat d'AutoFirma a partir d'un text del titular.
# Buit -> '' (sense filtre). Altrament 'subject.contains:TEXT' (AutoFirma triarà
# el certificat el subjecte del qual contingui aquest text, p. ex. el NIF/nom).
function _CertFilterValue([string]$text) {
    $t = ([string]$text).Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return '' }
    return 'subject.contains:' + $t
}

# Text per defecte del CAIXETÍ de la signatura visible (reprodueix l'aspecte
# "CERTIFICAT SENSE DNI" de l'usuari: nom / càrrec / organisme / data, SENSE DNI).
# La data la posa AutoFirma amb el marcador $$SIGNDATE=...$$.
function _DefaultCaixeti {
    return (@(
        'Sergi Fadurdo Modesto'
        "Enginyer d'Activitats"
        'Aj.Cornellà de Llobregat'
        '$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$'
    ) -join "`n")
}

# Posició del caixetí a la pàgina (A4 595x842 pt): a DALT A LA DRETA. Tunejable.
$Script:AutoFirmaCaixetiPos = @{ Page = 1; LLX = 360; LLY = 740; URX = 560; URY = 815 }

# Encoding del valor de -config d'AutoFirma. Segons la versió, AutoFirma vol els
# extraParams en Base64 ($true, RECOMANAT: evita problemes d'espais/salts a la
# línia d'ordres) o en text pla ($false). Si el caixetí no surt a Windows, prova
# de posar-ho a $false.
$Script:AutoFirmaConfigBase64 = $true

# Construeix la cadena d'extraParams (TEXT PLA, determinista) per a una signatura
# VISIBLE PAdES amb el caixetí donat. Els parells van separats per salt de línia;
# dins de layer2Text els salts són \n LITERAL (barra+n), com espera AutoFirma.
# Funció PURA. Caixetí buit -> '' (signatura invisible, com abans).
function _AutoFirmaVisibleExtraParams([string]$caixeti) {
    if ([string]::IsNullOrWhiteSpace($caixeti)) { return '' }
    $p = $Script:AutoFirmaCaixetiPos
    $layer = (([string]$caixeti -replace "`r`n", "`n") -replace "`n", '\n')
    return (@(
        "signaturePage=$($p.Page)"
        "signaturePositionOnPageLowerLeftX=$($p.LLX)"
        "signaturePositionOnPageLowerLeftY=$($p.LLY)"
        "signaturePositionOnPageUpperRightX=$($p.URX)"
        "signaturePositionOnPageUpperRightY=$($p.URY)"
        'layer2FontFamily=1'
        'layer2FontSize=8'
        "layer2Text=$layer"
    ) -join "`n")
}

# El fragment ' -config <valor>' per afegir a la línia d'arguments (o '' si no hi
# ha caixetí). Codifica els extraParams en Base64 o text pla segons el commutador.
function _AutoFirmaConfigArg([string]$caixeti) {
    $ep = _AutoFirmaVisibleExtraParams $caixeti
    if ([string]::IsNullOrWhiteSpace($ep)) { return '' }
    if ($Script:AutoFirmaConfigBase64) {
        $b64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ep))
        return ' -config ' + $b64
    }
    return ' -config "' + $ep + '"'
}

# Construeix la LÍNIA d'arguments (string, amb les rutes entre cometes) per signar
# un PDF amb AutoFirma. Funció PURA. Operació 'sign', PAdES, magatzem 'windows'.
# Si $caixeti no és buit, afegeix el -config de la signatura VISIBLE (caixetí a
# dalt a la dreta). Sense $caixeti es comporta EXACTAMENT com abans (invisible).
function _BuildAutoFirmaSignArgs([string]$inPdf, [string]$outPdf, [string]$filter, [string]$algorithm, [string]$caixeti = '') {
    if ([string]::IsNullOrWhiteSpace($algorithm)) { $algorithm = 'SHA256withRSA' }
    $a = 'sign -i "' + $inPdf + '" -o "' + $outPdf + '" -store windows -format pades -algorithm ' + $algorithm
    if (-not [string]::IsNullOrWhiteSpace($filter)) {
        $a += ' -filter "' + $filter + '"'
    }
    $a += (_AutoFirmaConfigArg $caixeti)
    return $a
}

# Rutes candidates on sol estar instal·lat AutoFirma.exe a Windows. Si les
# variables d'entorn no hi son (p. ex. proves a Linux), s'usen els camins
# literals habituals de Windows perque la funció sempre retorni candidats.
# Retorna un array pla de cadenes (NO un ArrayList amb ,$out: aixi @() sempre
# l'enumera bé i el bucle rep cadenes, no la llista sencera).
function _AutoFirmaCandidatePaths {
    $pf  = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles))       { 'C:\Program Files' }       else { $env:ProgramFiles }
    $px  = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }
    $paths = @(
        (Join-Path $pf 'AutoFirma\AutoFirma\AutoFirma.exe')
        (Join-Path $pf 'AutoFirma\AutoFirma.exe')
        (Join-Path $px 'AutoFirma\AutoFirma\AutoFirma.exe')
        (Join-Path $px 'AutoFirma\AutoFirma.exe')
    )
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $paths += (Join-Path $env:LOCALAPPDATA 'AutoFirma\AutoFirma\AutoFirma.exe')
    }
    return $paths
}

# Localitza AutoFirma.exe. Retorna SEMPRE una cadena (la ruta o ''). $preferit
# té prioritat (ruta desada per l'usuari).
function _FindAutoFirmaExe([string]$preferit) {
    if (-not [string]::IsNullOrWhiteSpace($preferit) -and (Test-Path -LiteralPath $preferit)) { return [string]$preferit }
    foreach ($c in @(_AutoFirmaCandidatePaths)) {
        $cs = [string]$c
        if (Test-Path -LiteralPath $cs) { return $cs }
    }
    return ''
}

# Extreu el "nom comú" (CN) d'un subjecte de certificat (DN). Funció PURA.
# "CN=NOM COGNOM - 12345678Z, O=..., C=ES" -> "NOM COGNOM - 12345678Z".
function _CertCommonName([string]$subject) {
    if ([string]::IsNullOrWhiteSpace($subject)) { return '' }
    $m = [regex]::Match($subject, 'CN=([^,]+)')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ([string]$subject).Trim()
}

# ----------------------------------------------------------------------------
# Estat desat (carpeta, opcions de signatura). Sidecar JSON, ignorat per git.
# ----------------------------------------------------------------------------
function _PdfSignarStatePath {
    $dir = if ($LocalActivitatsDir) { $LocalActivitatsDir } else { $env:TEMP }
    return (Join-Path $dir 'pdf-signar-state.json')
}

function _LoadPdfSignarState {
    $p = _PdfSignarStatePath
    $def = @{ folder = ''; sign = $false; certFilter = ''; autofirma = ''; overwrite = $false; visibleSign = $true; caixeti = (_DefaultCaixeti) }
    if (-not (Test-Path -LiteralPath $p)) { return $def }
    try {
        $o = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @('folder','certFilter','autofirma','caixeti')) { if ($o.PSObject.Properties[$k]) { $def[$k] = [string]$o.$k } }
        if ($o.PSObject.Properties['sign'])        { $def['sign'] = [bool]$o.sign }
        if ($o.PSObject.Properties['overwrite'])   { $def['overwrite'] = [bool]$o.overwrite }
        if ($o.PSObject.Properties['visibleSign']) { $def['visibleSign'] = [bool]$o.visibleSign }
        if ([string]::IsNullOrWhiteSpace([string]$def['caixeti'])) { $def['caixeti'] = (_DefaultCaixeti) }
    } catch { }
    return $def
}

function _SavePdfSignarState($state) {
    try {
        $dir = Split-Path -Parent (_PdfSignarStatePath)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        ($state | ConvertTo-Json) | Set-Content -LiteralPath (_PdfSignarStatePath) -Encoding UTF8
    } catch { }
}

# ----------------------------------------------------------------------------
# Diàleg d'opcions (carpeta + signatura). Retorna @{Folder;Sign;CertText;
# Overwrite;AutoFirma} o $null si es cancel·la.
# ----------------------------------------------------------------------------
function _ShowConvertPdfOptions {
    $st = _LoadPdfSignarState
    $folder = if (-not [string]::IsNullOrWhiteSpace($st.folder) -and (Test-Path -LiteralPath $st.folder)) { [string]$st.folder }
              elseif ($InformesDir -and (Test-Path -LiteralPath $InformesDir)) { [string]$InformesDir }
              else { '' }
    $autofirma = _FindAutoFirmaExe $st.autofirma

    # Certificats disponibles al magatzem de Windows (per triar-lo d'una llista,
    # en lloc d'escriure text). Primer element: deixar-ho a AutoFirma.
    $certOpts = New-Object System.Collections.ArrayList
    [void]$certOpts.Add(@{ Display = ('(triar-lo a AutoFirma en signar)'); Filter = '' })
    try {
        Get-ChildItem 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue |
            Where-Object { $_.HasPrivateKey } | Sort-Object NotAfter -Descending | ForEach-Object {
                $cn = _CertCommonName ([string]$_.Subject)
                $exp = ''
                try { $exp = $_.NotAfter.ToString('dd/MM/yyyy') } catch { }
                [void]$certOpts.Add(@{ Display = ("{0}  (fins {1})" -f $cn, $exp); Filter = (_CertFilterValue $cn) })
            }
    } catch { }

    $form = _NewForm
    $form.Text = 'Convertir informes a PDF'
    $form.ClientSize = New-Object System.Drawing.Size(510, 476)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # Carpeta: mateix format que la Configuració (quadre editable + "..." +
    # indicador ✓/⚠ en viu). Helper comú _AddConfigRow (Configuracio.ps1).
    $row = _AddConfigRow $form 70 'Carpeta amb els informes (Word):' $folder
    $tbF = $row.TextBox
    $y = [int]$row.NextY

    $cbOver = New-Object System.Windows.Forms.CheckBox
    $cbOver.Text = 'Tornar a generar els PDF que ja existeixen'
    $cbOver.Location = New-Object System.Drawing.Point(14, $y)
    $cbOver.AutoSize = $true
    $cbOver.Checked = [bool]$st.overwrite
    [void]$form.Controls.Add($cbOver)
    $y += 30

    $cbSign = New-Object System.Windows.Forms.CheckBox
    $cbSign.Text = 'Signar els PDF amb AutoFirma (certificat de Windows)'
    $cbSign.Location = New-Object System.Drawing.Point(14, $y)
    $cbSign.AutoSize = $true
    $cbSign.Checked = [bool]$st.sign
    [void]$form.Controls.Add($cbSign)
    $y += 26

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Certificat amb què signar:'
    $lblC.Location = New-Object System.Drawing.Point(34, $y)
    $lblC.AutoSize = $true
    [void]$form.Controls.Add($lblC)
    $y += 20

    $cbCert = New-Object System.Windows.Forms.ComboBox
    $cbCert.DropDownStyle = 'DropDownList'
    $cbCert.Location = New-Object System.Drawing.Point(34, $y)
    $cbCert.Size = New-Object System.Drawing.Size(454, 24)
    foreach ($o in $certOpts) { [void]$cbCert.Items.Add([string]$o.Display) }
    $selIdx = 0
    for ($i = 0; $i -lt $certOpts.Count; $i++) { if ([string]$certOpts[$i].Filter -eq [string]$st.certFilter -and [string]$st.certFilter -ne '') { $selIdx = $i } }
    if ($cbCert.Items.Count -gt 0) { $cbCert.SelectedIndex = $selIdx }
    [void]$form.Controls.Add($cbCert)
    $y += 30

    $lblAF = New-Object System.Windows.Forms.Label
    $lblAF.Location = New-Object System.Drawing.Point(34, $y)
    $lblAF.MaximumSize = New-Object System.Drawing.Size(454, 0)
    $lblAF.AutoSize = $true
    $lblAF.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
    if ([string]::IsNullOrWhiteSpace($autofirma)) {
        $lblAF.Text = ([char]0x26A0 + ' AutoFirma no trobat. Instal·la''l o desmarca la signatura.')
        $lblAF.ForeColor = [System.Drawing.Color]::Firebrick
    } else {
        $lblAF.Text = ([char]0x2713 + ' AutoFirma: ' + $autofirma)
        $lblAF.ForeColor = [System.Drawing.Color]::SeaGreen
    }
    [void]$form.Controls.Add($lblAF)
    $y += 40

    # Signatura VISIBLE (caixetí a dalt a la dreta) + text editable del caixetí.
    $cbVis = New-Object System.Windows.Forms.CheckBox
    $cbVis.Text = 'Signatura visible (caixetí a dalt a la dreta)'
    $cbVis.Location = New-Object System.Drawing.Point(34, $y)
    $cbVis.AutoSize = $true
    $cbVis.Checked = [bool]$st.visibleSign
    [void]$form.Controls.Add($cbVis)
    $y += 26

    $lblCx = New-Object System.Windows.Forms.Label
    $lblCx.Text = 'Text del caixetí (una línia per fila; $$SIGNDATE=...$$ = data):'
    $lblCx.Location = New-Object System.Drawing.Point(34, $y)
    $lblCx.AutoSize = $true
    $lblCx.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
    [void]$form.Controls.Add($lblCx)
    $y += 20

    $tbCx = New-Object System.Windows.Forms.TextBox
    $tbCx.Location = New-Object System.Drawing.Point(34, $y)
    $tbCx.Size = New-Object System.Drawing.Size(454, 76)
    $tbCx.Multiline = $true
    $tbCx.ScrollBars = 'Vertical'
    $tbCx.AcceptsReturn = $true
    $tbCx.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    # El TextBox multilínia només mostra CRLF; el caixetí es guarda amb LF.
    $tbCx.Text = ([string]$st.caixeti -replace "`r?`n", "`r`n")
    [void]$form.Controls.Add($tbCx)

    $syncSign = {
        $on = $cbSign.Checked
        $lblC.Enabled = $on; $cbCert.Enabled = $on; $lblAF.Enabled = $on
        $cbVis.Enabled = $on
        $onVis = ($on -and $cbVis.Checked)
        $lblCx.Enabled = $onVis; $tbCx.Enabled = $onVis
    }.GetNewClosure()
    $cbSign.add_CheckedChanged($syncSign)
    $cbVis.add_CheckedChanged($syncSign)
    & $syncSign

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Comença'
    $btnGo.Location = New-Object System.Drawing.Point(280, 438)
    $btnGo.Size = New-Object System.Drawing.Size(120, 30)
    $btnGo.Anchor = 'Bottom, Right'
    _StylePrimaryButton $btnGo
    [void]$form.Controls.Add($btnGo)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Tanca'
    $btnCancel.Location = New-Object System.Drawing.Point(406, 438)
    $btnCancel.Size = New-Object System.Drawing.Size(88, 30)
    $btnCancel.Anchor = 'Bottom, Right'
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $form.DialogResult = 'Cancel'; $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnCancel)

    $result = @{ Value = $null }
    $btnGo.add_Click({
        $f = [string]$tbF.Text
        if ([string]::IsNullOrWhiteSpace($f) -or -not (Test-Path -LiteralPath $f)) {
            [System.Windows.Forms.MessageBox]::Show('Tria una carpeta vàlida.', 'Convertir informes a PDF', 'OK', 'Warning') | Out-Null
            return
        }
        if ($cbSign.Checked -and [string]::IsNullOrWhiteSpace($autofirma)) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha trobat AutoFirma. Desmarca la signatura (es faran només els PDF) o instal·la AutoFirma.", 'Convertir informes a PDF', 'OK', 'Warning') | Out-Null
            return
        }
        $certFilter = ''
        if ($cbCert.SelectedIndex -ge 0 -and $cbCert.SelectedIndex -lt $certOpts.Count) { $certFilter = [string]$certOpts[$cbCert.SelectedIndex].Filter }
        $caixeti = ([string]$tbCx.Text -replace "`r`n", "`n")
        $result.Value = @{
            Folder = $f; Sign = [bool]$cbSign.Checked; CertFilter = $certFilter
            Overwrite = [bool]$cbOver.Checked; AutoFirma = [string]$autofirma
            VisibleSign = [bool]$cbVis.Checked; Caixeti = $caixeti
        }
        _SavePdfSignarState @{ folder = $f; sign = [bool]$cbSign.Checked; certFilter = $certFilter; autofirma = [string]$autofirma; overwrite = [bool]$cbOver.Checked; visibleSign = [bool]$cbVis.Checked; caixeti = $caixeti }
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    [void](_AddBrandHeader $form 'Convertir informes a PDF' ('PDF a la mateixa carpeta ' + [char]0x00B7 + ' signatura opcional amb AutoFirma'))
    [void]$form.ShowDialog()
    $v = $result.Value
    $form.Dispose()
    return $v
}

# ----------------------------------------------------------------------------
# Execució: converteix (i signa) amb barra de progrés i cancel·lació.
# ----------------------------------------------------------------------------
function _RunConvertPdf($opts) {
    # Defensa: si $opts arribés embolcallat en un array, agafem l'element real.
    if ($opts -is [System.Array]) { $opts = @($opts)[-1] }
    $folderPath = [string]$opts.Folder
    $afExe      = [string]$opts.AutoFirma
    # Enumeració dels Word de la carpeta (i subcarpetes), saltant temporals ~$.
    $files = New-Object System.Collections.ArrayList
    Get-ChildItem -LiteralPath $folderPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name -notlike '~$*' -and ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc')) {
            [void]$files.Add($_)
        }
    }
    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat cap Word (.doc/.docx) a la carpeta.", 'Convertir informes a PDF', 'OK', 'Information') | Out-Null
        return
    }

    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("S'han trobat {0} documents Word.`n`nEs generaran els PDF a la mateixa carpeta{1}.`n`nVols continuar?" -f $files.Count, $(if ($opts.Sign) { ' i es signaran amb AutoFirma' } else { '' })),
        'Convertir informes a PDF', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # ---- Finestra de progrés amb cancel·lació ----
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Convertir informes a PDF'
    $form.Size = New-Object System.Drawing.Size(580, 200)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(530, 60)
    $lbl.Text = 'Preparant...'
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 92)
    $bar.Size = New-Object System.Drawing.Size(530, 22)
    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $files.Count); $bar.Value = 0
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(430, 124)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({
        param($s, $e)
        if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true }
    }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $filter = [string]$opts.CertFilter
    # Caixetí de signatura visible (buit = signatura invisible, com abans).
    $caixeti = if ([bool]$opts.VisibleSign) { [string]$opts.Caixeti } else { '' }
    $converted = 0; $skipped = 0; $signed = 0; $errors = 0; $done = 0
    $errDetalls = New-Object System.Collections.ArrayList
    $word = $null
    try {
        $word = New-WordApp
        if ($null -eq $word) { return }   # New-WordApp ja avisa

        foreach ($f in $files) {
            if ($cancel.Flag) { break }
            $done++
            $bar.Value = [Math]::Min($bar.Maximum, $done)
            $lbl.Text = ("Processant {0} de {1}...`n{2}" -f $done, $files.Count, $f.Name)
            [System.Windows.Forms.Application]::DoEvents()

            $pdf = _PdfPathForDoc $f.FullName
            $pdfExists = Test-Path -LiteralPath $pdf
            $pdfTime = if ($pdfExists) { (Get-Item -LiteralPath $pdf).LastWriteTimeUtc } else { [datetime]::MinValue }

            # 1. Conversió a PDF (si cal).
            if (_PdfShouldConvert $pdfExists $f.LastWriteTimeUtc $pdfTime ([bool]$opts.Overwrite)) {
                try {
                    $doc = $word.Documents.Open($f.FullName, $false, $true)   # ReadOnly
                    try {
                        $doc.ExportAsFixedFormat($pdf, 17)   # 17 = wdExportFormatPDF
                    } finally {
                        $doc.Close($false)
                    }
                    $converted++
                    $pdfExists = $true
                } catch {
                    $errors++
                    [void]$errDetalls.Add(("PDF: {0} -> {1}" -f $f.Name, $_.Exception.Message))
                    continue
                }
            } else {
                $skipped++
            }

            # 2. Signatura (si es demana i el PDF existeix).
            if ($opts.Sign -and $pdfExists) {
                if ($cancel.Flag) { break }
                $lbl.Text = ("Signant {0} de {1}...`n{2}" -f $done, $files.Count, $f.Name)
                [System.Windows.Forms.Application]::DoEvents()
                $tmpSigned = $pdf + '.signat.pdf'
                try { if (Test-Path -LiteralPath $tmpSigned) { Remove-Item -LiteralPath $tmpSigned -Force } } catch { }
                $argLine = [string](_BuildAutoFirmaSignArgs $pdf $tmpSigned $filter '' $caixeti)
                try {
                    $p = Start-Process -FilePath $afExe -ArgumentList $argLine -Wait -PassThru -WindowStyle Hidden
                    if ($p.ExitCode -eq 0 -and (Test-Path -LiteralPath $tmpSigned)) {
                        Move-Item -LiteralPath $tmpSigned -Destination $pdf -Force
                        $signed++
                    } else {
                        $errors++
                        [void]$errDetalls.Add(("Signatura: {0} (codi {1})" -f $f.Name, $p.ExitCode))
                        try { if (Test-Path -LiteralPath $tmpSigned) { Remove-Item -LiteralPath $tmpSigned -Force } } catch { }
                    }
                } catch {
                    $errors++
                    [void]$errDetalls.Add(("Signatura: {0} -> {1}" -f $f.Name, $_.Exception.Message))
                }
            }
        }
    } finally {
        if ($null -ne $word) {
            try { $word.Quit() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
        }
        $cancel.Running = $false
        $form.Close(); $form.Dispose()
    }

    # ---- Resum ----
    $msg = New-Object System.Text.StringBuilder
    if ($cancel.Flag) { [void]$msg.AppendLine('Cancel·lat.') ; [void]$msg.AppendLine('') }
    [void]$msg.AppendLine(("PDF generats: {0}" -f $converted))
    [void]$msg.AppendLine(("Ja estaven al dia (saltats): {0}" -f $skipped))
    if ($opts.Sign) { [void]$msg.AppendLine(("PDF signats: {0}" -f $signed)) }
    if ($errors -gt 0) {
        [void]$msg.AppendLine(("Errors: {0}" -f $errors))
        [void]$msg.AppendLine('')
        $mostra = @($errDetalls) | Select-Object -First 8
        foreach ($e in $mostra) { [void]$msg.AppendLine('  - ' + $e) }
        if ($errDetalls.Count -gt 8) { [void]$msg.AppendLine(('  ... i {0} més.' -f ($errDetalls.Count - 8))) }
    }
    $icon = if ($errors -gt 0) { 'Warning' } else { 'Information' }
    [System.Windows.Forms.MessageBox]::Show($msg.ToString(), 'Convertir informes a PDF', 'OK', $icon) | Out-Null
}

# Punt d'entrada de l'eina (des del menú principal).
function Invoke-ConvertirPdf {
    $opts = _ShowConvertPdfOptions
    if ($null -eq $opts) { return }
    _RunConvertPdf $opts
}
