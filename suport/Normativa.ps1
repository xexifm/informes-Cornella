#requires -Version 5.1
<#
  Eina "Normativa": les BAIXADES (xarxa i Edge) i la FINESTRA. Nomes corre a
  Windows. El cataleg, els noms dels fitxers, l'index i com es reconeix una norma
  son a NormativaDades.ps1 (pur, i es el que fa servir la fitxa d'ajuda).
#>

# ----------------------------------------------------------------------------
# LES BAIXADES (xarxa i Edge: nomes a Windows)
# ----------------------------------------------------------------------------
function _NormativaPreparaXarxa {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    # El proxy de l'Ajuntament (si n'hi ha) amb l'usuari del Windows.
    try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch { }
}

function _NormativaGet([string]$url) {
    $r = Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $Script:NormativaUA -TimeoutSec 60 -MaximumRedirection 10 -UseDefaultCredentials -ErrorAction Stop
    return $r
}

function _NormativaGetBytes([string]$url, [string]$desti) {
    Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $Script:NormativaUA -TimeoutSec 180 -MaximumRedirection 10 -UseDefaultCredentials -OutFile $desti -ErrorAction Stop | Out-Null
    return [System.IO.File]::ReadAllBytes($desti)
}

function _NormativaEdgeExe {
    foreach ($p in @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'))) {
        if ($p -and (Test-Path -LiteralPath $p)) { return [string]$p }
    }
    return ''
}

# Imprimeix una pagina web a PDF amb l'Edge sense finestra. Un perfil PROPI
# (--user-data-dir): si l'Edge de l'usuari ja es obert, sense aixo l'ordre se
# n'hi aniria a ell i tornaria de seguida sense fer res.
function _NormativaImprimeix([string]$url, [string]$desti) {
    $edge = _NormativaEdgeExe
    if (-not $edge) { throw "No trobo l'Edge ni el Chrome per desar la pàgina com a PDF." }
    $perfil = Join-Path $env:TEMP 'informes-normativa-edge'
    if (Test-Path -LiteralPath $desti) { Remove-Item -LiteralPath $desti -Force }
    $argv = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--no-pdf-header-footer',
              '--virtual-time-budget=25000', ('--user-data-dir="' + $perfil + '"'), ('--print-to-pdf="' + $desti + '"'), ('"' + $url + '"'))
    $p = Start-Process -FilePath $edge -ArgumentList $argv -WindowStyle Hidden -PassThru
    if (-not $p.WaitForExit(120000)) { try { $p.Kill() } catch { }; throw "L'Edge no ha acabat en 2 minuts." }
    if (-not (Test-Path -LiteralPath $desti)) { throw "L'Edge no ha generat el PDF." }
    $b = [System.IO.File]::ReadAllBytes($desti)
    if (-not (_NormativaEsPdf $b)) { throw "L'Edge no ha generat un PDF vàlid." }
    # Una pagina que no s'ha arribat a carregar (o nomes l'avis de galetes) fa
    # un PDF d'una pagina gairebe buit.
    if ($b.Length -lt 20000) { throw ("El PDF ha sortit gairebé buit (" + $b.Length + " bytes): potser la pàgina no s'ha carregat.") }
    return $b
}

# EL DOM DE LA PAGINA JA DIBUIXADA (--dump-dom), per trobar-hi el boto "PDF":
# el Portal Juridic munta la pagina amb JavaScript, i el que torna el servidor
# sense executar-lo no porta l'enllac.
function _NormativaDomEdge([string]$url) {
    $edge = _NormativaEdgeExe
    if (-not $edge) { return '' }
    $perfil = Join-Path $env:TEMP 'informes-normativa-edge'
    $sortida = Join-Path $env:TEMP ('normativa-dom-' + [guid]::NewGuid().ToString('N') + '.html')
    try {
        $argv = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
                  '--virtual-time-budget=20000', ('--user-data-dir="' + $perfil + '"'), '--dump-dom', ('"' + $url + '"'))
        $p = Start-Process -FilePath $edge -ArgumentList $argv -WindowStyle Hidden -PassThru -RedirectStandardOutput $sortida
        if (-not $p.WaitForExit(90000)) { try { $p.Kill() } catch { }; return '' }
        if (-not (Test-Path -LiteralPath $sortida)) { return '' }
        return [System.IO.File]::ReadAllText($sortida, [System.Text.Encoding]::UTF8)
    } catch { return '' }
    finally { try { if (Test-Path -LiteralPath $sortida) { Remove-Item -LiteralPath $sortida -Force } } catch { } }
}

# UNA NORMA QUE NO ES DEL BOE, per ordre de preferencia:
#   1. l'URL ja es el PDF (EUR-Lex, CTE, guies, la Diputacio...);
#   2. el boto "PDF" de la pagina (Portal Juridic, BOPB, CIDO): primer al que
#      torna el servidor i, si no hi es, a la pagina dibuixada per l'Edge;
#   3. si no n'hi ha cap, la pagina impresa a PDF.
# Torna @{ Bytes; Via } (Via ho diu a l'index: si surt "pagina impresa" es que
# no s'ha trobat el PDF de debo).
function _NormativaBaixaWeb([string]$url, [string]$tmp) {
    $html = ''
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $Script:NormativaUA -TimeoutSec 60 -MaximumRedirection 10 -UseDefaultCredentials -OutFile $tmp -PassThru -ErrorAction Stop
        $b = [System.IO.File]::ReadAllBytes($tmp)
        if (_NormativaEsPdf $b) { return @{ Bytes = $b; Via = 'PDF' } }
        $html = [System.Text.Encoding]::UTF8.GetString($b)
    } catch { $html = '' }
    $provats = New-Object System.Collections.ArrayList
    foreach ($font in @('servidor', 'edge')) {
        if ($font -eq 'edge') { $html = _NormativaDomEdge $url }
        foreach ($c in @(_NormativaPdfsDeHtml $html $url | Select-Object -First 4)) {
            if ($provats.Contains($c)) { continue }
            [void]$provats.Add($c)
            try {
                $b = _NormativaGetBytes $c $tmp
                if (_NormativaEsPdf $b) { return @{ Bytes = $b; Via = 'PDF de la pàgina' } }
            } catch { }
        }
    }
    return @{ Bytes = (_NormativaImprimeix $url $tmp); Via = 'pàgina impresa' }
}

# Baixa UNA norma. Torna @{ Fet (s'ha escrit un fitxer); Versio; Error; Motiu }.
function _NormativaBaixaUna($e, [string]$dir, $est, [bool]$forca) {
    $nom = _NormativaNomFitxer $e
    $desti = Join-Path $dir $nom
    $existeix = Test-Path -LiteralPath $desti
    $font = _NormativaFont ([string]$e.Url)
    if ($font -eq 'manual') { return @{ Fet = $false; Versio = ''; Error = ''; Motiu = 'manual' } }
    $tmp = Join-Path $dir ('~baixant ' + [guid]::NewGuid().ToString('N') + '.pdf')
    try {
        $versio = ''
        $pdfUrl = ''
        $bytes = $null
        $via = 'BOE'
        if ($font -eq 'boe') {
            $pag = _NormativaGet ([string]$e.Url)
            $info = _NormativaBoeInfo ([string]$pag.Content)
            if (-not $info.Id) { throw "La pàgina del BOE no diu l'identificador de la norma." }
            $versio = [string]$info.Versio
            $motiu = _NormativaCalBaixar $est $existeix $versio (Get-Date) $forca
            if (-not $motiu) { return @{ Fet = $false; Versio = $versio; Error = ''; Motiu = '' } }
            $pdfUrl = if ($versio) { _NormativaBoePdfConsolidat $info.Id } else { [string]$info.Original }
            if (-not $pdfUrl) { $pdfUrl = _NormativaBoePdfConsolidat $info.Id }
            $bytes = _NormativaGetBytes $pdfUrl $tmp
            if (-not (_NormativaEsPdf $bytes) -and $info.Original -and $pdfUrl -ne $info.Original) {
                $bytes = _NormativaGetBytes ([string]$info.Original) $tmp
            }
        } else {
            $motiu = _NormativaCalBaixar $est $existeix '' (Get-Date) $forca
            if (-not $motiu) { return @{ Fet = $false; Versio = ''; Error = ''; Motiu = '' } }
            $w = _NormativaBaixaWeb ([string]$e.Url) $tmp
            $bytes = $w.Bytes
            $via = [string]$w.Via
        }
        if (-not (_NormativaEsPdf $bytes)) { throw "El que s'ha baixat no és un PDF." }
        # LA VERSIO ANTERIOR ES GUARDA si de debo es una altra: la del BOE, quan
        # ha canviat la data; la resta, quan la mida canvia mes d'un 2 %
        # (imprimir la mateixa pagina dues vegades no dona el mateix fitxer
        # byte a byte, pero si gairebe la mateixa mida).
        if ($existeix) {
            $vella = (Get-Item -LiteralPath $desti).Length
            $canvia = if ($font -eq 'boe') { $motiu -eq 'versio' } else { [Math]::Abs($vella - $bytes.Length) -gt ($vella * 0.02) }
            if ($canvia) {
                $ant = Join-Path $dir 'anteriors'
                if (-not (Test-Path -LiteralPath $ant)) { New-Item -ItemType Directory -Path $ant -Force | Out-Null }
                Move-Item -LiteralPath $desti -Destination (Join-Path $ant (_NormativaNomAnterior $nom (Get-Date))) -Force
            } else {
                Remove-Item -LiteralPath $desti -Force
            }
        }
        Move-Item -LiteralPath $tmp -Destination $desti -Force
        return @{ Fet = $true; Versio = $versio; Error = ''; Motiu = $motiu; Mida = [long]$bytes.Length; Via = $via }
    } catch {
        return @{ Fet = $false; Versio = ''; Error = [string]$_.Exception.Message; Motiu = 'error' }
    } finally {
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
    }
}

# ----------------------------------------------------------------------------
# LA FINESTRA (eina del menu)
# ----------------------------------------------------------------------------
function Invoke-Normativa {
    $normes = @(Get-NormativaCataleg)
    if ($normes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No trobo el catàleg de normativa (suport\normativa.json). Fes Actualitzar.bat.", 'Normativa', 'OK', 'Error') | Out-Null
        return
    }
    $dir = Get-NormativaDir
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $form = _NewForm
    $form.Text = 'Normativa'
    $form.ClientSize = New-Object System.Drawing.Size(720, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 460)
    $form.StartPosition = 'CenterScreen'
    $ui = @{ Corrent = $false; Cancel = $false }

    $lblResum = New-Object System.Windows.Forms.Label
    $lblResum.Location = New-Object System.Drawing.Point(16, 70)
    $lblResum.Size = New-Object System.Drawing.Size(690, 40)
    $lblResum.Anchor = 'Top,Left,Right'
    $form.Controls.Add($lblResum)

    $lnk = New-Object System.Windows.Forms.LinkLabel
    $lnk.Text = 'Obrir la carpeta'
    $lnk.Location = New-Object System.Drawing.Point(16, 112)
    $lnk.AutoSize = $true
    $lnk.add_LinkClicked({ try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $dir + '"') | Out-Null } catch { } }.GetNewClosure())
    $form.Controls.Add($lnk)

    $lnkIdx = New-Object System.Windows.Forms.LinkLabel
    $lnkIdx.Text = "Obrir l'índex (Excel)"
    $lnkIdx.Location = New-Object System.Drawing.Point(140, 112)
    $lnkIdx.AutoSize = $true
    $lnkIdx.add_LinkClicked({
        $p = Join-Path $dir $Script:NormativaIndexNom
        if (Test-Path -LiteralPath $p) { try { Start-Process -FilePath $p | Out-Null } catch { } }
    }.GetNewClosure())
    $form.Controls.Add($lnkIdx)

    $chkTot = New-Object System.Windows.Forms.CheckBox
    $chkTot.Text = 'Tornar-les a baixar totes, encara que no hagin canviat'
    $chkTot.Location = New-Object System.Drawing.Point(16, 140)
    $chkTot.AutoSize = $true
    $form.Controls.Add($chkTot)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(16, 172)
    $bar.Size = New-Object System.Drawing.Size(688, 20)
    $bar.Anchor = 'Top,Left,Right'
    $bar.Minimum = 0; $bar.Maximum = $normes.Count
    $form.Controls.Add($bar)

    $log = New-Object System.Windows.Forms.TextBox
    $log.Multiline = $true; $log.ReadOnly = $true; $log.ScrollBars = 'Vertical'
    $log.BackColor = [System.Drawing.Color]::White
    $log.Font = New-Object System.Drawing.Font('Consolas', 9)
    $log.Location = New-Object System.Drawing.Point(16, 200)
    $log.Size = New-Object System.Drawing.Size(688, 300)
    $log.Anchor = 'Top,Bottom,Left,Right'
    $form.Controls.Add($log)

    $fn = @{}
    $fn.Resum = {
        $estat = _NormativaLlegeixEstat $dir
        $baix = 0; $man = 0; $err = 0
        foreach ($e in $normes) {
            if (Test-Path -LiteralPath (Join-Path $dir (_NormativaNomFitxer $e))) { $baix++ }
            elseif ((_NormativaFont ([string]$e.Url)) -eq 'manual') { $man++ }
            elseif ($estat.ContainsKey([string]$e.Id) -and $estat[[string]$e.Id].Error) { $err++ }
        }
        $lblResum.Text = ([string]$normes.Count + ' normes al catàleg: ' + $baix + ' baixades, ' + ($normes.Count - $baix - $man) +
                          ' per baixar' + $(if ($err) { ' (' + $err + " amb error l'últim cop)" } else { '' }) + ', ' + $man + ' per desar a mà.' +
                          "`r`n" + $dir)
        $lnkIdx.Enabled = (Test-Path -LiteralPath (Join-Path $dir $Script:NormativaIndexNom))
    }.GetNewClosure()
    $fn.Log = { param($t) $log.AppendText($t + "`r`n"); [System.Windows.Forms.Application]::DoEvents() }.GetNewClosure()

    $peu = _AddPeuBotons $form @(@{ Nom = 'Tanca'; Text = 'Tancar' }) @(
        @{ Nom = 'Baixa'; Text = 'Baixar i actualitzar'; Estil = 'primari' }) 514 -Ancorat
    $btnTanca = $peu.Tanca; $btnBaixa = $peu.Baixa

    $btnTanca.add_Click({
        if ($ui.Corrent) { $ui.Cancel = $true; & $fn.Log 'Aturant després de la norma que s''està baixant...'; return }
        $form.Close()
    }.GetNewClosure())
    $form.add_FormClosing({ param($s, $ev) if ($ui.Corrent) { $ui.Cancel = $true; $ev.Cancel = $true } }.GetNewClosure())

    $btnBaixa.add_Click({
        if ($ui.Corrent) { return }
        $ui.Corrent = $true; $ui.Cancel = $false
        $btnBaixa.Enabled = $false; $chkTot.Enabled = $false; $btnTanca.Text = 'Aturar'
        $log.Clear()
        _NormativaPreparaXarxa
        $estat = _NormativaLlegeixEstat $dir
        $forca = [bool]$chkTot.Checked
        $n = @{ Noves = 0; Act = 0; Igual = 0; Err = 0; Man = 0 }
        $bar.Value = 0
        try {
            foreach ($e in $normes) {
                if ($ui.Cancel) { & $fn.Log 'Aturat.'; break }
                $id = [string]$e.Id
                $est = if ($estat.ContainsKey($id)) { $estat[$id] } else { $null }
                $r = _NormativaBaixaUna $e $dir $est $forca
                $ara = (Get-Date).ToString('o')
                switch ([string]$r.Motiu) {
                    'manual' { $n.Man++ }
                    'error'  {
                        $n.Err++
                        & $fn.Log ('ERROR  ' + $id + ': ' + $r.Error)
                        $estat[$id] = @{ Versio = $(if ($est) { $est.Versio } else { '' }); Baixat = $(if ($est) { $est.Baixat } else { '' }); Error = [string]$r.Error; Mida = $(if ($est) { $est.Mida } else { 0 }) }
                    }
                    ''       {
                        $n.Igual++
                        if ($est) { $est.Error = '' } else { $estat[$id] = @{ Versio = [string]$r.Versio; Baixat = $ara; Error = ''; Mida = 0 } }
                    }
                    default  {
                        if ($r.Motiu -eq 'nova') { $n.Noves++; & $fn.Log ('Nova   ' + $id) } else { $n.Act++; & $fn.Log ('Actualitzada  ' + $id) }
                        $estat[$id] = @{ Versio = [string]$r.Versio; Baixat = $ara; Error = ''; Mida = [long]$r.Mida; Via = [string]$r.Via }
                    }
                }
                $bar.Value = [Math]::Min($bar.Maximum, $bar.Value + 1)
                [System.Windows.Forms.Application]::DoEvents()
            }
        } finally {
            try { _NormativaDesaEstat $dir $estat } catch { & $fn.Log ("No s'ha pogut desar l'estat: " + $_.Exception.Message) }
            $errIdx = _NormativaEscriuIndex $dir $normes $estat
            if ($errIdx) { & $fn.Log ("No s'ha pogut escriure l'índex (és obert a l'Excel?): " + $errIdx) }
            & $fn.Log ('')
            & $fn.Log (('Fet. Noves: {0} · Actualitzades: {1} · Ja al dia: {2} · Errors: {3} · Per desar a mà: {4}' -f $n.Noves, $n.Act, $n.Igual, $n.Err, $n.Man))
            if ($n.Err -gt 0) { & $fn.Log ("Les que han fallat surten a l'índex amb el motiu; es tornaran a provar la propera vegada.") }
            $ui.Corrent = $false
            $btnBaixa.Enabled = $true; $chkTot.Enabled = $true; $btnTanca.Text = 'Tancar'
            & $fn.Resum
        }
    }.GetNewClosure())

    [void](_AddBrandHeader $form 'Normativa' ("Tota la normativa en una carpeta, classificada pel nom del fitxer"))
    & $fn.Resum
    [void]$form.ShowDialog()
    $form.Dispose()
}
