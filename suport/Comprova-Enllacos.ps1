#requires -Version 5.1
<#
.SYNOPSIS
  Comprova els enllacos (URLs) d'un cataleg .docx i avisa dels que estan CAIGUTS.

.DESCRIPTION
  Llegeix un .docx (per defecte ESTRUCTURALS\REQ1.docx), n'extreu tots els
  enllaces (estil Cita, text amb http i hipervincles incrustats) i fa una
  peticio a cadascun per saber si responen. Marca en VERD els que van be i en
  VERMELL els que estan CAIGUTS (codi 4xx/5xx o sense resposta), i al final fa
  un resum dels caiguts.

  NO necessita Word: llegeix el .docx com un ZIP (XML intern). Si nomes hi ha
  Windows PowerShell 5.1, funciona igual.

.PARAMETER Docx
  Ruta del .docx a comprovar. Si no s'indica, s'usa ESTRUCTURALS\REQ1.docx.
  Es pot passar 'all' per comprovar TOTS els catalegs REQ*.docx d'ESTRUCTURALS.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1 all
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1 "C:\ruta\REQ2.docx"
#>

param([string]$Docx)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot   = Split-Path -Parent $ScriptRoot
$EstructDir = Join-Path $RepoRoot 'ESTRUCTURALS'

# Decideix quins fitxers comprovar.
$targets = @()
if ([string]::IsNullOrWhiteSpace($Docx)) {
    $targets = @(Join-Path $EstructDir 'REQ1.docx')
} elseif ($Docx -eq 'all') {
    $targets = @(Get-ChildItem -LiteralPath $EstructDir -Filter '*.docx' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '0 *' -and -not $_.Name.StartsWith('~$') } |
        Sort-Object Name | ForEach-Object { $_.FullName })
} else {
    $targets = @($Docx)
}
if ($targets.Count -eq 0) { Write-Host "No s'ha trobat cap cataleg a comprovar." -ForegroundColor Red; exit 1 }

# Extreu els URLs d'un .docx (sense Word). Agafa nomes els enllacos REALS:
#   1) els hipervincles (relacions amb Type '...hyperlink'),
#   2) els URLs escrits com a TEXT visible (REQ1 posa l'URL en estil Cita).
# NO agafa els URLs dels espais de noms XML (schemas.openxmlformats.org...),
# que no son enllacos de contingut.
function _ReadZipEntryText($zip, [string]$name) {
    $entry = $zip.GetEntry($name); if (-not $entry) { return $null }
    $sr = New-Object System.IO.StreamReader($entry.Open())
    $txt = $sr.ReadToEnd(); $sr.Close(); return $txt
}
function Get-DocxUrls([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $urls = New-Object System.Collections.Generic.List[string]
    $add = {
        param($u)
        if ([string]::IsNullOrWhiteSpace($u)) { return }
        $u = ([string]$u).Trim().TrimEnd('.',',',';',')')
        if ($u -match '^https?://' -and -not $urls.Contains($u)) { [void]$urls.Add($u) }
    }
    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $path))
    try {
        # 1) Hipervincles (word/_rels/document.xml.rels)
        $relsTxt = _ReadZipEntryText $zip 'word/_rels/document.xml.rels'
        if ($relsTxt) {
            try {
                [xml]$rels = $relsTxt
                foreach ($rel in $rels.Relationships.Relationship) {
                    if (([string]$rel.Type) -like '*hyperlink*') { & $add $rel.Target }
                }
            } catch { }
        }
        # 2) URLs escrits com a text visible (contingut de <w:t>)
        $docTxt = _ReadZipEntryText $zip 'word/document.xml'
        if ($docTxt) {
            $sb = New-Object System.Text.StringBuilder
            foreach ($m in [regex]::Matches($docTxt, '(?s)<w:t[^>]*>(.*?)</w:t>')) {
                [void]$sb.Append($m.Groups[1].Value).Append(' ')
            }
            $visible = [System.Net.WebUtility]::HtmlDecode($sb.ToString())
            foreach ($m in [regex]::Matches($visible, 'https?://[^\s"<>\]\)]+')) { & $add $m.Value }
        }
    } finally { $zip.Dispose() }
    return $urls
}

# Forcem TLS 1.2 (Windows PowerShell 5.1 per defecte pot no negociar-lo).
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol } catch { }

$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
$totalCaiguts = @()

foreach ($t in $targets) {
    if (-not (Test-Path -LiteralPath $t)) { Write-Host "No trobo: $t" -ForegroundColor Red; continue }
    $urls = Get-DocxUrls $t
    Write-Host ("`n===== {0}  ({1} enllacos) =====" -f (Split-Path -Leaf $t), $urls.Count) -ForegroundColor Cyan
    foreach ($u in $urls) {
        $code = $null; $ok = $false
        # Provem HEAD i, si el servidor no l'accepta (405) o falla, GET.
        foreach ($method in 'Head','Get') {
            try {
                $r = Invoke-WebRequest -Uri $u -Method $method -TimeoutSec 25 -UserAgent $ua -UseBasicParsing -MaximumRedirection 5 -ErrorAction Stop
                $code = [int]$r.StatusCode; $ok = ($code -lt 400); break
            } catch {
                $resp = $null; try { $resp = $_.Exception.Response } catch { }
                if ($resp -and $resp.StatusCode) {
                    $code = [int]$resp.StatusCode
                    # Un 405 a HEAD no vol dir caigut: ho reintentem amb GET.
                    if ($code -eq 405 -and $method -eq 'Head') { continue }
                    $ok = ($code -lt 400); break
                } else {
                    $code = 'sense resposta'
                    if ($method -eq 'Get') { break }
                }
            }
        }
        if ($ok) {
            Write-Host ("  OK     [{0}] {1}" -f $code, $u) -ForegroundColor Green
        } else {
            Write-Host ("  CAIGUT [{0}] {1}" -f $code, $u) -ForegroundColor Red
            $totalCaiguts += [pscustomobject]@{ Fitxer=(Split-Path -Leaf $t); Codi=$code; Url=$u }
        }
    }
}

Write-Host ("`n========================================") -ForegroundColor Cyan
if ($totalCaiguts.Count -eq 0) {
    Write-Host "Tots els enllacos responen correctament." -ForegroundColor Green
} else {
    Write-Host ("ENLLACOS CAIGUTS: {0}" -f $totalCaiguts.Count) -ForegroundColor Red
    $totalCaiguts | ForEach-Object { Write-Host ("  [{0}] {1}  ({2})" -f $_.Codi, $_.Url, $_.Fitxer) -ForegroundColor Yellow }
}
Write-Host ("========================================")
