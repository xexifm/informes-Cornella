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

# Construeix la LÍNIA d'arguments (string, amb les rutes entre cometes) per signar
# un PDF amb AutoFirma. Funció PURA. Operació 'sign', PAdES, magatzem 'windows'.
function _BuildAutoFirmaSignArgs([string]$inPdf, [string]$outPdf, [string]$filter, [string]$algorithm) {
    if ([string]::IsNullOrWhiteSpace($algorithm)) { $algorithm = 'SHA256withRSA' }
    $a = 'sign -i "' + $inPdf + '" -o "' + $outPdf + '" -store windows -format pades -algorithm ' + $algorithm
    if (-not [string]::IsNullOrWhiteSpace($filter)) {
        $a += ' -filter "' + $filter + '"'
    }
    return $a
}

# Rutes candidates on sol estar instal·lat AutoFirma.exe a Windows. Si les
# variables d'entorn no hi son (p. ex. proves a Linux), s'usen els camins
# literals habituals de Windows perque la funció sempre retorni candidats.
function _AutoFirmaCandidatePaths {
    $pf  = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles))       { 'C:\Program Files' }       else { $env:ProgramFiles }
    $px  = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }
    $out = New-Object System.Collections.ArrayList
    foreach ($base in @($pf, $px)) {
        [void]$out.Add((Join-Path $base 'AutoFirma\AutoFirma\AutoFirma.exe'))
        [void]$out.Add((Join-Path $base 'AutoFirma\AutoFirma.exe'))
    }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        [void]$out.Add((Join-Path $env:LOCALAPPDATA 'AutoFirma\AutoFirma\AutoFirma.exe'))
    }
    return ,$out
}

# Localitza AutoFirma.exe. Retorna la ruta o '' si no es troba. $preferit té
# prioritat (ruta desada per l'usuari).
function _FindAutoFirmaExe([string]$preferit) {
    if (-not [string]::IsNullOrWhiteSpace($preferit) -and (Test-Path -LiteralPath $preferit)) { return $preferit }
    foreach ($c in @(_AutoFirmaCandidatePaths)) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return ''
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
    $def = @{ folder = ''; sign = $false; certText = ''; autofirma = ''; overwrite = $false }
    if (-not (Test-Path -LiteralPath $p)) { return $def }
    try {
        $o = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @('folder','certText','autofirma')) { if ($o.PSObject.Properties[$k]) { $def[$k] = [string]$o.$k } }
        if ($o.PSObject.Properties['sign'])      { $def['sign'] = [bool]$o.sign }
        if ($o.PSObject.Properties['overwrite']) { $def['overwrite'] = [bool]$o.overwrite }
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

    $form = _NewForm
    $form.Text = 'Convertir informes a PDF'
    $form.ClientSize = New-Object System.Drawing.Size(600, 300)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $y = 74
    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Carpeta amb els informes (Word):'
    $lblF.Location = New-Object System.Drawing.Point(18, $y)
    $lblF.AutoSize = $true
    [void]$form.Controls.Add($lblF)

    $tbF = New-Object System.Windows.Forms.TextBox
    $tbF.Location = New-Object System.Drawing.Point(18, ($y + 22))
    $tbF.Size = New-Object System.Drawing.Size(460, 24)
    $tbF.ReadOnly = $true
    $tbF.Text = $folder
    [void]$form.Controls.Add($tbF)

    $btnTria = New-Object System.Windows.Forms.Button
    $btnTria.Text = 'Tria...'
    $btnTria.Location = New-Object System.Drawing.Point(486, ($y + 21))
    $btnTria.Size = New-Object System.Drawing.Size(96, 26)
    _StyleSecondaryButton $btnTria
    $btnTria.add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = 'Tria la carpeta amb els informes en Word'
        if (-not [string]::IsNullOrWhiteSpace($tbF.Text) -and (Test-Path -LiteralPath $tbF.Text)) { $fb.SelectedPath = $tbF.Text }
        if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tbF.Text = $fb.SelectedPath }
    }.GetNewClosure())
    [void]$form.Controls.Add($btnTria)

    $cbOver = New-Object System.Windows.Forms.CheckBox
    $cbOver.Text = 'Tornar a generar els PDF que ja existeixen'
    $cbOver.Location = New-Object System.Drawing.Point(18, ($y + 54))
    $cbOver.AutoSize = $true
    $cbOver.Checked = [bool]$st.overwrite
    [void]$form.Controls.Add($cbOver)

    $cbSign = New-Object System.Windows.Forms.CheckBox
    $cbSign.Text = 'Signar els PDF amb AutoFirma (certificat de Windows)'
    $cbSign.Location = New-Object System.Drawing.Point(18, ($y + 84))
    $cbSign.AutoSize = $true
    $cbSign.Checked = [bool]$st.sign
    [void]$form.Controls.Add($cbSign)

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = ('Certificat (nom o NIF del titular, per triar-lo sol):')
    $lblC.Location = New-Object System.Drawing.Point(38, ($y + 110))
    $lblC.AutoSize = $true
    [void]$form.Controls.Add($lblC)

    $tbC = New-Object System.Windows.Forms.TextBox
    $tbC.Location = New-Object System.Drawing.Point(38, ($y + 132))
    $tbC.Size = New-Object System.Drawing.Size(320, 24)
    $tbC.Text = [string]$st.certText
    [void]$form.Controls.Add($tbC)

    $lblAF = New-Object System.Windows.Forms.Label
    $lblAF.Location = New-Object System.Drawing.Point(38, ($y + 160))
    $lblAF.AutoSize = $true
    if ([string]::IsNullOrWhiteSpace($autofirma)) {
        $lblAF.Text = 'AutoFirma: no s''ha trobat. Instal·la''l o desmarca la signatura.'
        $lblAF.ForeColor = [System.Drawing.Color]::Firebrick
    } else {
        $lblAF.Text = 'AutoFirma: ' + $autofirma
        $lblAF.ForeColor = [System.Drawing.Color]::ForestGreen
    }
    [void]$form.Controls.Add($lblAF)

    $syncSign = {
        $on = $cbSign.Checked
        $tbC.Enabled = $on; $lblC.Enabled = $on; $lblAF.Enabled = $on
    }.GetNewClosure()
    $cbSign.add_CheckedChanged($syncSign)
    & $syncSign

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Comença'
    $btnGo.Location = New-Object System.Drawing.Point(370, 262)
    $btnGo.Size = New-Object System.Drawing.Size(120, 30)
    _StylePrimaryButton $btnGo
    [void]$form.Controls.Add($btnGo)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Tanca'
    $btnCancel.Location = New-Object System.Drawing.Point(496, 262)
    $btnCancel.Size = New-Object System.Drawing.Size(86, 30)
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
        $result.Value = @{
            Folder = $f; Sign = [bool]$cbSign.Checked; CertText = [string]$tbC.Text
            Overwrite = [bool]$cbOver.Checked; AutoFirma = $autofirma
        }
        _SavePdfSignarState @{ folder = $f; sign = [bool]$cbSign.Checked; certText = [string]$tbC.Text; autofirma = $autofirma; overwrite = [bool]$cbOver.Checked }
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
    # Enumeració dels Word de la carpeta (i subcarpetes), saltant temporals ~$.
    $files = New-Object System.Collections.ArrayList
    Get-ChildItem -LiteralPath $opts.Folder -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
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

    $filter = _CertFilterValue $opts.CertText
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
                $argLine = _BuildAutoFirmaSignArgs $pdf $tmpSigned $filter ''
                try {
                    $p = Start-Process -FilePath $opts.AutoFirma -ArgumentList $argLine -Wait -PassThru -WindowStyle Hidden
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
