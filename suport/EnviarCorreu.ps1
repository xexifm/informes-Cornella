#requires -Version 5.1
<#
  EnviarCorreu.ps1 - Eina "Enviar correu" (secció MÒBIL).

  Envia el correu de requeriments des del PC, amb el MATEIX format i remitent que
  el mòbil (mateixa plantilla d'EmailJS). El cos de requeriments es construeix
  llegint el .docx GENERAT (estructura REQ1: seccions, subseccions, numeració i
  qualsevol edició manual), i s'embolcalla amb les condicions del correu (text
  introductori + text final, SENSE conclusions) definides a
  docs/dades/email-textos.json.

  Enviament: API REST d'EmailJS des de PowerShell. La Public key / Service ID /
  Template ID es llegeixen de docs/config.js (no secretes). La Private key es
  desa a la carpeta local/ del repositori (gitignored): local/emailjs.json ->
  { "private_key": "..." }.
#>

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

function _CorreuRepoRoot {
    if ($RepoRoot) { return [string]$RepoRoot }
    return (Split-Path -Parent $PSScriptRoot)
}

# --- Configuració (claus d'EmailJS) ------------------------------------------
function _CorreuConfig {
    $repo = _CorreuRepoRoot
    $cfgJs = Join-Path $repo (Join-Path 'docs' 'config.js')
    $pub = ''; $svc = ''; $tpl = ''; $from = 'Ajuntament de Cornellà de Llobregat - Activitats'
    if (Test-Path -LiteralPath $cfgJs) {
        $t = Get-Content -LiteralPath $cfgJs -Raw -Encoding UTF8
        if ($t -match 'EMAILJS_PUBLIC_KEY\s*:\s*"([^"]*)"')  { $pub = $matches[1] }
        if ($t -match 'EMAILJS_SERVICE_ID\s*:\s*"([^"]*)"')  { $svc = $matches[1] }
        if ($t -match 'EMAILJS_TEMPLATE_ID\s*:\s*"([^"]*)"') { $tpl = $matches[1] }
        if ($t -match 'EMAIL_FROM_NAME\s*:\s*"([^"]*)"')     { $from = $matches[1] }
    }
    # Private key: carpeta local/ (fora del repositori public).
    $priv = ''
    $pkPath = Join-Path $repo (Join-Path 'local' 'emailjs.json')
    if (Test-Path -LiteralPath $pkPath) {
        try {
            $j = Get-Content -LiteralPath $pkPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($j.private_key) { $priv = [string]$j.private_key }
        } catch { }
    }
    return [pscustomobject]@{
        PublicKey = $pub; ServiceId = $svc; TemplateId = $tpl; FromName = $from; PrivateKey = $priv
        PrivatePath = $pkPath
    }
}

# --- Utils de text -> HTML ----------------------------------------------------
function _EscHtml($s) {
    return ([string]$s).Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}
# Escapa, **negreta** -> <b>, //cursiva// -> <i>, i enllaça els URLs.
function _TextToHtml($s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $h = _EscHtml $s
    $h = [regex]::Replace($h, '\*\*(.+?)\*\*', '<b>$1</b>')
    $h = [regex]::Replace($h, '//(.+?)//', '<i>$1</i>')
    $h = [regex]::Replace($h, '(https?://[^\s<]+)', '<a href="$1">$1</a>')
    return $h.Replace("`r`n","`n").Replace("`n",'<br>')
}
function _NormCorreu($s) {
    if ($null -eq $s) { return '' }
    if (Get-Command _NormalizeText -ErrorAction SilentlyContinue) { return (_NormalizeText $s) }
    $t = ([string]$s).Normalize([Text.NormalizationForm]::FormD)
    return (($t -replace '\p{Mn}','').ToLower().Trim())
}

# --- Llegir el cos de requeriments del .docx generat --------------------------
# Recorre els paragrafs: comença despres de la intro de deficiencies (o del
# primer titol/numero) i acaba abans de les conclusions. Torna HTML.
function _DocxRequerimentsHtml($docxPath) {
    $phrases = if ($SeguimentConclusionPhrases) { @($SeguimentConclusionPhrases) } else {
        @("Vist l'anterior", 'Ho poso al seu coneixement', 'Cornella de Llobregat,', 'CONCLUSIONS')
    }
    $phrN = $phrases | ForEach-Object { _NormCorreu $_ }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try { $word.DisplayAlerts = 0 } catch { }
    $out = New-Object System.Collections.ArrayList
    $started = $false
    try {
        $doc = $word.Documents.Open($docxPath, $false, $true)   # ReadOnly
        try {
            foreach ($p in $doc.Paragraphs) {
                $txt = ([string]$p.Range.Text).TrimEnd("`r","`n","`a"," ")
                if ([string]::IsNullOrWhiteSpace($txt)) { continue }
                $n = _NormCorreu $txt

                # Fi: primera frase de conclusions.
                $isConcl = $false
                foreach ($pn in $phrN) { if ($pn -and $n.StartsWith($pn)) { $isConcl = $true; break } }
                if ($isConcl) { break }

                $isCaps = ($txt -cmatch '\p{Lu}') -and -not ($txt -cmatch '\p{Ll}')
                $isNum  = $txt -match '^\s*\d+\.\s'
                $isUrl  = $txt -match '^\s*https?://'

                if (-not $started) {
                    # Comença despres de la intro ("...deficiencies... esmenar...")
                    if ($n -match 'defici' -and $n -match 'esmenar') { $started = $true; continue }
                    if ($isCaps -or $isNum) { $started = $true }  # o al primer titol/numero
                    else { continue }
                }

                # Subratllat -> subseccio.
                $isUnder = $false
                try { $u = $p.Range.Font.Underline; if ($u -ne 0 -and $u -ne 9999999) { $isUnder = $true } } catch { }

                if ($isCaps) {
                    [void]$out.Add('<div style="font-weight:bold;margin-top:12px">' + (_EscHtml $txt) + '</div>')
                } elseif ($isUrl) {
                    [void]$out.Add('<div><a href="' + (_EscHtml $txt) + '">' + (_EscHtml $txt) + '</a></div>')
                } elseif ($isUnder -and -not $isNum) {
                    [void]$out.Add('<div style="text-decoration:underline;margin-top:4px">' + (_EscHtml $txt) + '</div>')
                } else {
                    $mt = if ($isNum) { 'margin-top:6px' } else { 'margin-left:18px' }
                    [void]$out.Add('<div style="' + $mt + '">' + (_TextToHtml $txt) + '</div>')
                }
            }
        } finally { $doc.Close($false) }
    } finally {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    return ($out -join "`n")
}

# --- Embolcall del correu (plantilla email-textos.json) -----------------------
function _CorreuTextos {
    if (Get-Command _LoadEmailTextos -ErrorAction SilentlyContinue) { return (_LoadEmailTextos) }
    $repo = _CorreuRepoRoot
    $p = Join-Path $repo (Join-Path 'docs' (Join-Path 'dades' 'email-textos.json'))
    if (Test-Path -LiteralPath $p) {
        try { return (Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { }
    }
    return ([pscustomobject]@{ assumpte = 'GIA {ID_GIA} Requeriments'; cos = '{REQUERIMENTS}' })
}
function _FillVars([string]$s, $h) {
    $today = (Get-Date).ToString('dd/MM/yyyy')
    $g = { param($k) if ($h -and $h.ContainsKey($k)) { [string]$h[$k] } else { '' } }
    return $s.Replace('{ID_GIA}', (& $g 'ID_GIA')).Replace('{ADRECA}', (& $g 'ADRECA')).
        Replace('{ACTIVITAT}', (& $g 'ACTIVITAT')).Replace('{TITULAR}', (& $g 'TITULAR')).
        Replace('{DATA}', $today)
}
function _BuildCorreu($requerimentsHtml, $header) {
    $tx = _CorreuTextos
    $subject = _FillVars ([string]$tx.assumpte) $header
    $cos = _FillVars ([string]$tx.cos) $header
    $parts = $cos -split '\{REQUERIMENTS\}', 2
    $html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;line-height:1.4">'
    $html += _TextToHtml $parts[0]
    $html += $requerimentsHtml
    if ($parts.Count -gt 1) { $html += _TextToHtml $parts[1] }
    $html += '</div>'
    return [pscustomobject]@{ Subject = $subject; Html = $html }
}

# --- Enviament EmailJS --------------------------------------------------------
# Adreces disponibles per a Còpia Oculta (CCO). La primera va marcada per defecte.
$Script:CorreuBccOpcions = @(
    @{ Addr = 'sfadurdom@aj-cornella.cat'; Default = $true }
    @{ Addr = 'rbaratoc@aj-cornella.cat';  Default = $false }
    @{ Addr = 'misalasj@aj-cornella.cat';  Default = $false }
    @{ Addr = 'hamadorp@aj-cornella.cat';  Default = $false }
)

function Send-EmailJs($cfg, $toEmail, $bcc, $subject, $htmlMessage) {
    $payload = @{
        service_id  = $cfg.ServiceId
        template_id = $cfg.TemplateId
        user_id     = $cfg.PublicKey
        accessToken = $cfg.PrivateKey
        template_params = @{ to_email = $toEmail; bcc = $bcc; subject = $subject; message = $htmlMessage; name = $cfg.FromName }
    }
    $json  = $payload | ConvertTo-Json -Depth 6
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    Invoke-RestMethod -Method Post -Uri 'https://api.emailjs.com/api/v1.0/email/send' -ContentType 'application/json' -Body $bytes | Out-Null
}

# --- Localitzar el .docx mes recent ------------------------------------------
function _LatestDocx {
    $dir = if ($OutputDir) { [string]$OutputDir } else { Join-Path (_CorreuRepoRoot) 'Informes generats' }
    if (Get-Command _ResolveOutputDir -ErrorAction SilentlyContinue) {
        try { $dir = _ResolveOutputDir } catch { }
    }
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    $f = Get-ChildItem -LiteralPath $dir -Filter '*.docx' -File -ErrorAction SilentlyContinue |
         Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    return $null
}

# --- Diàleg d'enviament -------------------------------------------------------
# Diàleg únic: fusiona el missatge "Informe generat" amb l'enviament del correu.
# Torna @{ To = @(...); Bcc = @(...) } o $null si es cancel·la.
function _DialegEnviar($build, $destinatariDefault, $docxPath) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Informe generat - Enviar correu'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.Size = New-Object System.Drawing.Size(600, 420)

    $lblInf = New-Object System.Windows.Forms.Label
    $lblInf.Text = "Informe generat:`n$docxPath"
    $lblInf.Location = New-Object System.Drawing.Point(15, 12)
    $lblInf.Size = New-Object System.Drawing.Size(560, 42)
    $form.Controls.Add($lblInf)

    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = "Assumpte: $($build.Subject)"
    $lblA.Location = New-Object System.Drawing.Point(15, 58)
    $lblA.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblA)

    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Text = "Destinataris (separa'ls amb ; si n'hi ha més d'un):"
    $lblD.Location = New-Object System.Drawing.Point(15, 88)
    $lblD.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblD)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(15, 110)
    $tb.Size = New-Object System.Drawing.Size(560, 24)
    $tb.Text = [string]$destinatariDefault
    $form.Controls.Add($tb)

    $lblB = New-Object System.Windows.Forms.Label
    $lblB.Text = 'Còpia oculta (CCO):'
    $lblB.Location = New-Object System.Drawing.Point(15, 146)
    $lblB.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblB)

    $checks = @()
    $y = 170
    foreach ($opt in $Script:CorreuBccOpcions) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $opt.Addr
        $cb.Checked = [bool]$opt.Default
        $cb.Location = New-Object System.Drawing.Point(25, $y)
        $cb.Size = New-Object System.Drawing.Size(540, 22)
        $form.Controls.Add($cb)
        $checks += $cb
        $y += 26
    }

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Enviar'; $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(390, 335); $ok.Size = New-Object System.Drawing.Size(90, 32)
    $form.AcceptButton = $ok; $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'No enviar'; $cancel.DialogResult = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(485, 335); $cancel.Size = New-Object System.Drawing.Size(90, 32)
    $form.CancelButton = $cancel; $form.Controls.Add($cancel)

    # Scroll vertical i ajust a la pantalla (vegeu suport/UiFinestra.ps1).
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })
    if ($form.ShowDialog() -ne 'OK') { return $null }

    $tos = @($tb.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $bccs = @()
    foreach ($cb in $checks) { if ($cb.Checked) { $bccs += $cb.Text } }
    return @{ To = $tos; Bcc = $bccs }
}

# Envia el correu per a un .docx concret. Reutilitzable (eina i final de generacio).
function Send-CorreuPerDocx($docxPath) {
    $cfg = _CorreuConfig
    if (-not $cfg.PublicKey -or -not $cfg.ServiceId -or -not $cfg.TemplateId) {
        [System.Windows.Forms.MessageBox]::Show("Falten les claus d'EmailJS a docs\config.js.",'Enviar correu','OK','Warning') | Out-Null
        return
    }
    if (-not $cfg.PrivateKey) {
        [System.Windows.Forms.MessageBox]::Show(
            "Falta la Private key d'EmailJS. Crea el fitxer:`n$($cfg.PrivatePath)`n`namb el contingut:`n{ ""private_key"": ""GOCSPX...la teva private key..."" }",
            'Enviar correu', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $docxPath -or -not (Test-Path -LiteralPath $docxPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat cap informe (.docx) per enviar.",'Enviar correu','OK','Information') | Out-Null
        return
    }

    $rep = if (Get-Command Load-LastReport -ErrorAction SilentlyContinue) { Load-LastReport } else { $null }
    $header = @{}
    if ($rep -and $rep.Header) { foreach ($p in $rep.Header.PSObject.Properties) { $header[$p.Name] = $p.Value } }

    try {
        $reqHtml = _DocxRequerimentsHtml $docxPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error llegint l'informe:`n$($_.Exception.Message)",'Enviar correu','OK','Error') | Out-Null
        return
    }
    $build = _BuildCorreu $reqHtml $header

    $res = _DialegEnviar $build ([string]$header['EMAIL']) $docxPath
    if ($null -eq $res) { return }
    if (-not $res.To -or @($res.To).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Indica almenys un destinatari.','Enviar correu','OK','Warning') | Out-Null
        return
    }
    $toStr  = ($res.To -join ',')
    $bccStr = ($res.Bcc -join ',')
    try {
        Send-EmailJs $cfg $toStr $bccStr $build.Subject $build.Html
        $resum = "Correu enviat a: $toStr"
        if ($bccStr) { $resum += "`nCCO: $bccStr" }
        [System.Windows.Forms.MessageBox]::Show($resum,'Enviar correu','OK','Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut enviar:`n$($_.Exception.Message)",'Enviar correu','OK','Error') | Out-Null
    }
}

# Eina de menu: agafa el .docx mes recent (l'ultim generat, potser editat a ma).
function Invoke-EnviarCorreu {
    $docx = _LatestDocx
    Send-CorreuPerDocx $docx
}

# Ofereix enviar el correu al final de generar un informe (normal o seguiment).
# El diàleg d'enviament és la confirmació (ja fusionat amb "Informe generat"):
# no es fa cap pregunta Sí/No prèvia. Si l'usuari no vol enviar, tanca el diàleg.
function Offer-EnviarCorreu($docxPath) {
    if ([string]::IsNullOrWhiteSpace([string]$docxPath) -or -not (Test-Path -LiteralPath $docxPath)) {
        $docxPath = _LatestDocx
    }
    Send-CorreuPerDocx $docxPath
}
