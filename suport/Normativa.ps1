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

# L'EDGE SENSE FINESTRA (imprimir una pagina, o treure'n el DOM dibuixat).
#
# AL PC DE L'USUARI ES PENJAVA (setembre 2026): 2 minuts per norma, i amb
# 60 normes del Portal Juridic la primera passada no s'acabava mai. Tres canvis:
#   - un PERFIL NOU a cada crida (--user-data-dir unic, que s'esborra): un Edge
#     que es quedava penjat bloquejava el perfil compartit i feia penjar tots
#     els que venien darrere;
#   - si no acaba, es mata TOT l'arbre de processos (taskkill /T), no nomes el
#     primer: els fills eren els que es quedaven vius;
#   - si es penja UN cop, en aquella passada ja no es torna a fer servir
#     ($Script:NormativaEdgeKO): val mes una norma amb error que una hora d'espera.
$Script:NormativaEdgeKO = $false

function _NormativaEdge([string[]]$argv, [int]$segons, [string]$sortida = '') {
    if ($Script:NormativaEdgeKO) { throw "L'Edge no respon en aquest ordinador (s'ha deixat de fer servir en aquesta passada)." }
    $edge = _NormativaEdgeExe
    if (-not $edge) { throw "No trobo l'Edge ni el Chrome." }
    $perfil = Join-Path $env:TEMP ('informes-normativa-edge-' + [guid]::NewGuid().ToString('N'))
    $tots = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
              ('--user-data-dir="' + $perfil + '"')) + @($argv)
    try {
        $p = if ($sortida) { Start-Process -FilePath $edge -ArgumentList $tots -WindowStyle Hidden -PassThru -RedirectStandardOutput $sortida }
             else { Start-Process -FilePath $edge -ArgumentList $tots -WindowStyle Hidden -PassThru }
        if (-not $p.WaitForExit($segons * 1000)) {
            try { Start-Process -FilePath 'taskkill.exe' -ArgumentList @('/T', '/F', '/PID', [string]$p.Id) -WindowStyle Hidden -Wait | Out-Null } catch { }
            $Script:NormativaEdgeKO = $true
            throw ("L'Edge no ha acabat en " + $segons + " segons (no es tornarà a fer servir en aquesta passada).")
        }
    } finally {
        Start-Sleep -Milliseconds 300
        try { if (Test-Path -LiteralPath $perfil) { Remove-Item -LiteralPath $perfil -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
    }
}

# Imprimeix una pagina web a PDF.
function _NormativaImprimeix([string]$url, [string]$desti) {
    if (Test-Path -LiteralPath $desti) { Remove-Item -LiteralPath $desti -Force }
    _NormativaEdge @('--no-pdf-header-footer', '--virtual-time-budget=20000', ('--print-to-pdf="' + $desti + '"'), ('"' + $url + '"')) 60
    if (-not (Test-Path -LiteralPath $desti)) { throw "L'Edge no ha generat el PDF." }
    $b = [System.IO.File]::ReadAllBytes($desti)
    if (-not (_NormativaEsPdf $b)) { throw "L'Edge no ha generat un PDF vàlid." }
    # Una pagina que no s'ha arribat a carregar (o nomes l'avis de galetes) fa
    # un PDF d'una pagina gairebe buit.
    if ($b.Length -lt 20000) { throw ("El PDF ha sortit gairebé buit (" + $b.Length + " bytes): potser la pàgina no s'ha carregat.") }
    return $b
}

# EL DOM DE LA PAGINA JA DIBUIXADA (--dump-dom). '' si no s'ha pogut.
function _NormativaDomEdge([string]$url) {
    $sortida = Join-Path $env:TEMP ('normativa-dom-' + [guid]::NewGuid().ToString('N') + '.html')
    try {
        _NormativaEdge @('--virtual-time-budget=15000', '--dump-dom', ('"' + $url + '"')) 45 $sortida
        if (-not (Test-Path -LiteralPath $sortida)) { return '' }
        return [System.IO.File]::ReadAllText($sortida, [System.Text.Encoding]::UTF8)
    } catch { return '' }
    finally { try { if (Test-Path -LiteralPath $sortida) { Remove-Item -LiteralPath $sortida -Force } } catch { } }
}

# EL PORTAL JURIDIC SENSE L'EDGE: la pagina tal com la dona el servidor i les
# metadades ELI (RDF/TTL/XML), buscant-hi el numero de versio del PDF del DOGC
# (_NormativaPdfPjurDeText). Torna @{ Pdf; Textos } (els textos serveixen
# tambe a la revisio, per saber si es vigent). Es desa per URL mentre dura la
# passada: la revisio i la baixada no l'han de demanar dues vegades.
$Script:NormativaPjurCache = @{}
function _NormativaFontsPjur([string]$url) {
    if ($Script:NormativaPjurCache.ContainsKey($url)) { return $Script:NormativaPjurCache[$url] }
    $textos = New-Object System.Collections.ArrayList
    $pdf = ''
    $html = ''
    try { $html = [string](_NormativaGet $url).Content } catch { $html = '' }
    if ($html) { [void]$textos.Add($html) }
    $pdf = _NormativaPdfPjurDeText $html
    $eli = if ($url -match '/eli/es-ct/') { _NormativaEliDeText $url } else { _NormativaEliDeText $html }
    if (-not $pdf) {
        foreach ($u in @(_NormativaUrlsMetaPjur $html $url $eli)) {
            $t = ''
            try { $t = [string](_NormativaGet $u).Content } catch { continue }
            if (-not $t) { continue }
            [void]$textos.Add($t)
            $pdf = _NormativaPdfPjurDeText $t
            if ($pdf) { break }
        }
    }
    $r = @{ Pdf = $pdf; Textos = $textos.ToArray(); Eli = $eli }
    $Script:NormativaPjurCache[$url] = $r
    return $r
}

# UNA NORMA QUE NO ES DEL BOE, per ordre de preferencia:
#   1. l'URL ja es el PDF (EUR-Lex, CTE, guies, la Diputacio...);
#   2. el boto "PDF" de la pagina (Portal Juridic, BOPB, CIDO): primer al que
#      torna el servidor i, si no hi es, a la pagina dibuixada per l'Edge;
#   3. si no n'hi ha cap, la pagina impresa a PDF.
# Torna @{ Bytes; Via } (Via ho diu a l'index: si surt "pagina impresa" es que
# no s'ha trobat el PDF de debo).
function _NormativaBaixaWeb([string]$url, [string]$tmp) {
    if ($url -match '(?i)portaljuridic\.gencat\.cat|dogc\.gencat\.cat') {
        $pj = _NormativaFontsPjur $url
        if ($pj.Pdf) {
            try {
                $b = _NormativaGetBytes ([string]$pj.Pdf) $tmp
                if (_NormativaEsPdf $b) { return @{ Bytes = $b; Via = 'PDF del Portal Jurídic' } }
            } catch { }
        }
    }
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
        _NormativaDesaNou $tmp (Join-Path $dir $nom) $dir $(if ($font -eq 'boe') { $motiu -eq 'versio' } else { $null })
        return @{ Fet = $true; Versio = $versio; Error = ''; Motiu = $motiu; Mida = [long]$bytes.Length; Via = $via }
    } catch {
        return @{ Fet = $false; Versio = ''; Error = [string]$_.Exception.Message; Motiu = 'error' }
    } finally {
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
    }
}

# EL FITXER NOU AL SEU LLOC, i l'anterior a 'anteriors' si de debo es una altra
# versio: la del BOE, quan ha canviat la data ($canvia); la resta ($canvia =
# $null), quan la mida canvia mes d'un 2 % (imprimir la mateixa pagina dues
# vegades no dona el mateix fitxer byte a byte, pero si gairebe la mateixa mida).
function _NormativaDesaNou([string]$tmp, [string]$desti, [string]$dir, $canvia) {
    if (Test-Path -LiteralPath $desti) {
        $vella = (Get-Item -LiteralPath $desti).Length
        $nova = (Get-Item -LiteralPath $tmp).Length
        $guarda = if ($null -ne $canvia) { [bool]$canvia } else { [Math]::Abs($vella - $nova) -gt ($vella * 0.02) }
        if ($guarda) {
            $ant = Join-Path $dir 'anteriors'
            if (-not (Test-Path -LiteralPath $ant)) { New-Item -ItemType Directory -Path $ant -Force | Out-Null }
            Move-Item -LiteralPath $desti -Destination (Join-Path $ant (_NormativaNomAnterior ([System.IO.Path]::GetFileName($desti)) (Get-Date))) -Force
        } else {
            Remove-Item -LiteralPath $desti -Force
        }
    }
    Move-Item -LiteralPath $tmp -Destination $desti -Force
}

# UNA COL·LECCIO (les ITC de Bombers, les TINSCI): la llista de documents es
# treu de la pagina cada vegada, perque Interior en publica de nous. Si la pagina
# no porta els PDF directament, se'n miren les pagines filles (un nivell). Cada
# document s'apunta a l'estat amb 'Pare' = la col·leccio. Torna els comptadors.
function _NormativaBaixaColleccio($e, [string]$dir, $estat, [bool]$forca, $log) {
    $n = @{ Noves = 0; Act = 0; Igual = 0; Err = 0; Docs = 0 }
    $url = [string]$e.Url
    $html = ''
    try { $html = [string](_NormativaGet $url).Content } catch { $html = '' }
    $docs = @(_NormativaDocsDeColleccio $html $url)
    if ($docs.Count -eq 0) { $html = _NormativaDomEdge $url; $docs = @(_NormativaDocsDeColleccio $html $url) }
    if ($docs.Count -eq 0) {
        $llista = New-Object System.Collections.ArrayList
        foreach ($sp in @(_NormativaSubpagines $html $url | Select-Object -First 80)) {
            $h2 = ''
            try { $h2 = [string](_NormativaGet ([string]$sp.Url)).Content } catch { }
            $d2 = @(_NormativaDocsDeColleccio $h2 ([string]$sp.Url))
            if ($d2.Count -eq 0) { $h2 = _NormativaDomEdge ([string]$sp.Url); $d2 = @(_NormativaDocsDeColleccio $h2 ([string]$sp.Url)) }
            foreach ($d in $d2) {
                # Un sol PDF a la fitxa: el nom bo es el de l'enllac de la llista.
                $t = if ($d2.Count -eq 1 -and $sp.Text) { [string]$sp.Text } else { [string]$d.Text }
                [void]$llista.Add([pscustomobject]@{ Url = [string]$d.Url; Text = $t })
            }
        }
        $docs = @($llista)
    }
    $id = [string]$e.Id
    if ($docs.Count -eq 0) {
        $estat[$id] = @{ Error = "no s'hi ha trobat cap document"; Baixat = ''; Versio = ''; Mida = 0 }
        $n.Err++
        & $log ('ERROR  ' + $id + ": no s'hi ha trobat cap document")
        return $n
    }
    $estat[$id] = @{ Error = ''; Baixat = (Get-Date).ToString('o'); Versio = ''; Mida = 0 }
    $n.Docs = $docs.Count
    $usats = @{}
    foreach ($d in $docs) {
        $nom = _NormativaNomDocColleccio $e ([string]$d.Text) ([string]$d.Url)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($nom); $k = 2
        while ($usats.ContainsKey($nom)) { $nom = $base + ' (' + $k + ').pdf'; $k++ }
        $usats[$nom] = $true
        $clau = $id + ' | ' + $nom
        $est = if ($estat.ContainsKey($clau)) { $estat[$clau] } else { $null }
        $desti = Join-Path $dir $nom
        $motiu = _NormativaCalBaixar $est (Test-Path -LiteralPath $desti) '' (Get-Date) $forca
        if (-not $motiu) { $n.Igual++; continue }
        $tmp = Join-Path $dir ('~baixant ' + [guid]::NewGuid().ToString('N') + '.pdf')
        try {
            $b = _NormativaGetBytes ([string]$d.Url) $tmp
            if (-not (_NormativaEsPdf $b)) { throw "no és un PDF" }
            _NormativaDesaNou $tmp $desti $dir $null
            $estat[$clau] = @{ Pare = $id; Titol = [string]$d.Text; Url = [string]$d.Url; Nom = $nom; Baixat = (Get-Date).ToString('o'); Versio = ''; Error = ''; Mida = [long]$b.Length; Via = 'PDF' }
            if ($motiu -eq 'nova') { $n.Noves++; & $log ('Nova   ' + $nom) } else { $n.Act++; & $log ('Actualitzada  ' + $nom) }
        } catch {
            $n.Err++
            & $log ('ERROR  ' + $nom + ': ' + $_.Exception.Message)
            $estat[$clau] = @{ Pare = $id; Titol = [string]$d.Text; Url = [string]$d.Url; Nom = $nom; Baixat = $(if ($est) { $est.Baixat } else { '' }); Versio = ''; Error = [string]$_.Exception.Message; Mida = 0 }
        } finally {
            try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
        }
    }
    return $n
}

# TOTA LA BAIXADA, sense finestra: la fan servir l'eina Normativa i la revisio
# del programa (Revisio.ps1). $log rep cada linia; $pas es crida despres de cada
# norma (la barra); $cancel diu si s'ha d'aturar. Desa l'estat i l'index.
function Invoke-NormativaBaixada($normes, [string]$dir, [bool]$forca, $log, $pas = $null, $cancel = $null) {
    _NormativaPreparaXarxa
    $Script:NormativaEdgeKO = $false
    $Script:NormativaPjurCache = @{}
    $estat = _NormativaLlegeixEstat $dir
    $n = @{ Noves = 0; Act = 0; Igual = 0; Err = 0; Man = 0 }
    try {
        foreach ($e in @($normes)) {
            if ($null -ne $cancel -and (& $cancel)) { & $log 'Aturat.'; break }
            $id = [string]$e.Id
            if ($e.Colleccio) {
                & $log ('Col·lecció  ' + $id + '...')
                $c = _NormativaBaixaColleccio $e $dir $estat $forca $log
                foreach ($k in @('Noves', 'Act', 'Igual', 'Err')) { $n[$k] += [int]$c[$k] }
                if ($null -ne $pas) { & $pas }
                continue
            }
            $est = if ($estat.ContainsKey($id)) { $estat[$id] } else { $null }
            $r = _NormativaBaixaUna $e $dir $est $forca
            $ara = (Get-Date).ToString('o')
            switch ([string]$r.Motiu) {
                'manual' { $n.Man++ }
                'error'  {
                    $n.Err++
                    & $log ('ERROR  ' + $id + ': ' + $r.Error)
                    $estat[$id] = @{ Versio = $(if ($est) { $est.Versio } else { '' }); Baixat = $(if ($est) { $est.Baixat } else { '' }); Error = [string]$r.Error; Mida = $(if ($est) { $est.Mida } else { 0 }) }
                }
                ''       {
                    $n.Igual++
                    if ($est) { $est.Error = '' } else { $estat[$id] = @{ Versio = [string]$r.Versio; Baixat = $ara; Error = ''; Mida = 0 } }
                }
                default  {
                    if ($r.Motiu -eq 'nova') { $n.Noves++; & $log ('Nova   ' + $id) } else { $n.Act++; & $log ('Actualitzada  ' + $id) }
                    $estat[$id] = @{ Versio = [string]$r.Versio; Baixat = $ara; Error = ''; Mida = [long]$r.Mida; Via = [string]$r.Via }
                }
            }
            if ($null -ne $pas) { & $pas }
        }
    } finally {
        try { _NormativaDesaEstat $dir $estat } catch { & $log ("No s'ha pogut desar l'estat: " + $_.Exception.Message) }
        $errIdx = _NormativaEscriuIndex $dir $normes $estat
        if ($errIdx) { & $log ("No s'ha pogut escriure l'índex (és obert a l'Excel?): " + $errIdx) }
    }
    return $n
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
        $baix = 0; $man = 0; $err = 0; $col = 0
        $docsCol = @(@($estat.Keys) | Where-Object { $estat[$_].Pare -and -not $estat[$_].Error }).Count
        foreach ($e in $normes) {
            if ($e.Colleccio) { $col++; continue }
            if (Test-Path -LiteralPath (Join-Path $dir (_NormativaNomFitxer $e))) { $baix++ }
            elseif ((_NormativaFont ([string]$e.Url)) -eq 'manual') { $man++ }
            elseif ($estat.ContainsKey([string]$e.Id) -and $estat[[string]$e.Id].Error) { $err++ }
        }
        $nn = $normes.Count - $col
        $lblResum.Text = ([string]$nn + ' normes i guies al catàleg: ' + $baix + ' baixades, ' + ($nn - $baix - $man) +
                          ' per baixar' + $(if ($err) { ' (' + $err + " amb error l'últim cop)" } else { '' }) + ', ' + $man + ' per desar a mà. Col·leccions (ITC, TINSCI): ' + $col + ', amb ' + $docsCol + ' documents baixats.' +
                          "`r`n" + $dir)
        $lnkIdx.Enabled = (Test-Path -LiteralPath (Join-Path $dir $Script:NormativaIndexNom))
    }.GetNewClosure()
    $fn.Log = { param($t) $log.AppendText($t + "`r`n"); [System.Windows.Forms.Application]::DoEvents() }.GetNewClosure()
    $fn.Pas = { $bar.Value = [Math]::Min($bar.Maximum, $bar.Value + 1); [System.Windows.Forms.Application]::DoEvents() }.GetNewClosure()
    $fn.Cancel = { [bool]$ui.Cancel }.GetNewClosure()

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
        $bar.Value = 0
        $n = @{ Noves = 0; Act = 0; Igual = 0; Err = 0; Man = 0 }
        try {
            $n = Invoke-NormativaBaixada $normes $dir ([bool]$chkTot.Checked) $fn.Log $fn.Pas $fn.Cancel
        } finally {
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
