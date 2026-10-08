#requires -Version 5.1
<#
.SYNOPSIS
  Diagnostic de la VIGENCIA de la normativa (eina "Revisar requeriments").

.DESCRIPTION
  Octubre 2026: a la revisio, quasi tot el Portal Juridic surt "no s'ha pogut
  saber si es vigent". Les primeres normes, perque la pagina que torna l'Edge
  no porta l'etiqueta VIGENT on el programa la busca; la resta, perque l'Edge
  es penja (45 s) i en aquella passada ja no s'hi torna a provar. Des d'on es
  programa el Portal Juridic i el BOE no s'hi pot arribar, o sigui que aquest
  script DESA el que responen al PC de l'usuari:

    - la pagina tal com la dona el servidor i les metadades ELI;
    - els scripts de la pagina (el Portal Juridic es una aplicacio JavaScript:
      la crida que porta la vigencia hi ha de sortir);
    - la pagina dibuixada per l'Edge, amb mes temps que a la revisio;
    - del BOE, la pagina i les dades obertes (metadatos) de quatre normes.

  NO toca res del programa ni dels catalegs. Ho deixa tot en una carpeta i un
  .zip a local\revisions\, amb un resum.txt. Nomes son pagines publiques.

  El llanca Provar-Vigencia.bat (doble clic).
#>

$ErrorActionPreference = 'Continue'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

$pvNom = 'diagnostic-vigencia-' + (Get-Date).ToString('yyyyMMdd-HHmm')
$pvDir = Join-Path (Get-LocalSubdir $RepoRoot 'Revisions') $pvNom
New-Item -ItemType Directory -Path $pvDir -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $pvDir 'scripts') -Force | Out-Null
$pvResum = New-Object System.Collections.ArrayList
$pvLog = {
    param($t)
    Write-Host $t
    [void]$pvResum.Add([string]$t)
}
$pvDesa = {
    param([string]$nom, [string]$text)
    [System.IO.File]::WriteAllText((Join-Path $pvDir $nom), [string]$text, (New-Object System.Text.UTF8Encoding($false)))
}

_NormativaPreparaXarxa
& $pvLog ('Diagnostic de vigencia ' + (Get-Date).ToString('dd/MM/yyyy HH:mm') + '  -  PowerShell ' + $PSVersionTable.PSVersion)
try { & $pvLog ('Edge: ' + (_NormativaEdgeExe)) } catch { & $pvLog ('Edge: no trobat (' + $_.Exception.Message + ')') }

# ---- PORTAL JURIDIC -------------------------------------------------------
$pvPjur = @(
    @{ Nom = 'pjur-llei-20-2009-eli';        Url = 'https://portaljuridic.gencat.cat/eli/es-ct/l/2009/12/04/20' },
    @{ Nom = 'pjur-decret-130-2003-docid';   Url = 'https://portaljuridic.gencat.cat/ca/document-del-pjur/?documentId=322238' },
    @{ Nom = 'pjur-llei-13-2017-no-vigent';  Url = 'https://portaljuridic.gencat.cat/ca/document-del-pjur/?documentId=792564' })
$pvScripts = @{}
foreach ($p in $pvPjur) {
    & $pvLog ''
    & $pvLog ('== ' + $p.Nom + '  ' + $p.Url)
    $html = ''
    try {
        $r = _NormativaGet $p.Url
        $html = [string]$r.Content
        & $pvDesa ($p.Nom + '-servidor.html') $html
        & $pvLog ('  servidor: ' + [int]$r.StatusCode + ', ' + $html.Length + ' caracters')
    } catch { & $pvLog ('  servidor: ERROR ' + $_.Exception.Message) }

    try {
        $Script:NormativaPjurCache = @{}
        $pj = _NormativaFontsPjur $p.Url
        $i = 0
        foreach ($t in @($pj.Textos)) {
            $i++
            & $pvDesa ($p.Nom + '-font-' + $i + '.txt') ([string]$t)
            & $pvLog ('  font ' + $i + ': ' + ([string]$t).Length + ' caracters; ELI=' + (_RevEstatEli ([string]$t)).Estat + '; etiqueta=' + (_RevEstatPjur ([string]$t)).Estat)
        }
        & $pvLog ('  ELI: ' + [string]$pj.Eli + '   PDF: ' + [string]$pj.Pdf)
        foreach ($u in @(_NormativaUrlsMetaPjur $html $p.Url ([string]$pj.Eli))) { & $pvLog ('  metadades provades: ' + $u) }
    } catch { & $pvLog ('  fonts: ERROR ' + $_.Exception.Message) }

    # Els scripts de la pagina (una sola vegada cada un).
    foreach ($m in [regex]::Matches($html, '(?i)<script[^>]+src\s*=\s*["'']([^"'']+)["'']')) {
        $src = $m.Groups[1].Value
        try { $abs = (New-Object System.Uri((New-Object System.Uri($p.Url)), $src)).AbsoluteUri } catch { continue }
        if ($pvScripts.ContainsKey($abs)) { continue }
        $pvScripts[$abs] = $true
        try {
            $js = [string](_NormativaGet $abs).Content
            $fn = 'scripts\' + ($pvScripts.Count) + '-' + ([System.IO.Path]::GetFileName((New-Object System.Uri($abs)).AbsolutePath) -replace '[^\w\.\-]', '_')
            & $pvDesa $fn $js
            $n = ([regex]::Matches($js, '(?i)vigen')).Count
            & $pvLog ('  script ' + $abs + ': ' + $js.Length + ' caracters, "vigen" ' + $n + ' vegades')
        } catch { & $pvLog ('  script ' + $abs + ': ERROR ' + $_.Exception.Message) }
    }

    # L'Edge, amb mes temps que a la revisio (30 s de pagina, 90 s en total).
    $Script:NormativaEdgeKO = @{}
    $sortida = Join-Path $pvDir ($p.Nom + '-edge.html')
    $t0 = Get-Date
    try {
        _NormativaEdge @('--virtual-time-budget=30000', '--dump-dom', ('"' + $p.Url + '"')) 90 $p.Url $sortida
        $dom = if (Test-Path -LiteralPath $sortida) { [System.IO.File]::ReadAllText($sortida, [System.Text.Encoding]::UTF8) } else { '' }
        $txt = _RevTextDeHtml $dom
        $k = $txt.IndexOf('Copia la URI ELI')
        & $pvLog ('  Edge: ' + [int]((Get-Date) - $t0).TotalSeconds + ' s, ' + $dom.Length + ' caracters; etiqueta=' + (_RevEstatPjur $dom).Estat + '; "Copia la URI ELI" a ' + $k)
        foreach ($mm in [regex]::Matches($txt, '(?i).{0,60}\bvigen\w*.{0,60}') | Select-Object -First 6) { & $pvLog ('    ...' + $mm.Value + '...') }
    } catch { & $pvLog ('  Edge: ERROR despres de ' + [int]((Get-Date) - $t0).TotalSeconds + ' s: ' + $_.Exception.Message) }
}

# ---- BOE ------------------------------------------------------------------
$pvBoe = @(
    @{ Nom = 'boe-rd-1562-1998';               Id = 'BOE-A-1998-19183' },
    @{ Nom = 'boe-rd-1002-2002';               Id = 'BOE-A-2002-19574' },
    @{ Nom = 'boe-rd-842-2002-vigent';         Id = 'BOE-A-2002-18099' },
    @{ Nom = 'boe-rd-1836-1999-derogat';       Id = 'BOE-A-1999-24924' })
foreach ($b in $pvBoe) {
    & $pvLog ''
    & $pvLog ('== ' + $b.Nom + '  ' + $b.Id)
    foreach ($v in @(
        @{ Suf = 'act.html';  Url = ('https://www.boe.es/buscar/act.php?id=' + $b.Id) },
        @{ Suf = 'doc.html';  Url = ('https://www.boe.es/buscar/doc.php?id=' + $b.Id) },
        @{ Suf = 'metadatos.xml'; Url = ('https://www.boe.es/datosabiertos/api/legislacion-consolidada/id/' + $b.Id + '/metadatos'); Accept = 'application/xml' })) {
        try {
            # Per _NormativaGet, com la resta del fitxer: el diagnostic ha de
            # demanar les pagines EXACTAMENT com les demana el programa. L'Accept
            # nomes el porta l'XML de dades obertes; a les dues .html, cap -que
            # es el que les feia tornar una cosa diferent de la que veu el
            # programa, i tot seguit es passaven per _RevEstatBoe-.
            $r = _NormativaGet $v.Url ([string]$v.Accept)
            $c = [string]$r.Content
            & $pvDesa ($b.Nom + '-' + $v.Suf) $c
            $e = if ($v.Suf -like '*.html') { (_RevEstatBoe $c).Estat } else { ([regex]::Match($c, '(?is)<(estatus_derogacion|vigencia_agotada|estado_consolidacion)[^>]*>[^<]*')).Value }
            & $pvLog ('  ' + $v.Suf + ': ' + [int]$r.StatusCode + ', ' + $c.Length + ' caracters -> ' + $e)
        } catch { & $pvLog ('  ' + $v.Suf + ': ERROR ' + $_.Exception.Message) }
    }
}

& $pvDesa 'resum.txt' ($pvResum -join "`r`n")
$pvZip = $pvDir + '.zip'
try {
    if (Test-Path -LiteralPath $pvZip) { Remove-Item -LiteralPath $pvZip -Force }
    Compress-Archive -Path (Join-Path $pvDir '*') -DestinationPath $pvZip -Force
    Write-Host ''
    Write-Host ('Fet. Passa a Claude aquest fitxer: ' + $pvZip)
    try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $pvZip + '"') | Out-Null } catch { }
} catch {
    Write-Host ('Fet, pero no s''ha pogut fer el .zip: ' + $_.Exception.Message)
    Write-Host ('Passa a Claude la carpeta: ' + $pvDir)
}
