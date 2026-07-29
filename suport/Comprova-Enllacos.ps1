#requires -Version 5.1
<#
.SYNOPSIS
  Comprova els enllacos (URLs) dels catalegs i avisa dels que estan CAIGUTS.

.DESCRIPTION
  Llegeix els catalegs .json d'ESTRUCTURALS, n'extreu tots els enllacos i fa una
  peticio a cadascun per saber si responen. Marca en VERD els que van be i en
  VERMELL els que estan CAIGUTS (codi 4xx/5xx o sense resposta), i al final en
  fa un resum.

  Llegeix el JSON, que es la FONT DE VERITAT dels catalegs. Abans obria el .docx
  com un ZIP i en treia els hipervincles i el text visible; els .docx ja no son
  catalegs (son vistes generades), i al JSON els enllacos ja venen marcats amb
  "url": true, o sigui que no cal endevinar res.

.PARAMETER Cataleg
  Nom o ruta del cataleg a comprovar. Sense valor, es comproven TOTS els
  catalegs d'ESTRUCTURALS (els .json que no comencen per "0 ").

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1
  powershell -ExecutionPolicy Bypass -File suport\Comprova-Enllacos.ps1 REQ1
#>

param([string]$Cataleg)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot   = Split-Path -Parent $ScriptRoot
$EstructDir = Join-Path $RepoRoot 'ESTRUCTURALS'

# Decideix quins catalegs comprovar.
$targets = @()
if ([string]::IsNullOrWhiteSpace($Cataleg)) {
    $targets = @(Get-ChildItem -LiteralPath $EstructDir -Filter '*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '0 *' -and -not $_.Name.StartsWith('~$') } |
        Sort-Object Name | ForEach-Object { $_.FullName })
} elseif (Test-Path -LiteralPath $Cataleg) {
    $targets = @((Resolve-Path -LiteralPath $Cataleg).Path)
} else {
    # Nom curt ('REQ1'): el busquem a ESTRUCTURALS.
    $n = if ([System.IO.Path]::GetExtension($Cataleg) -ieq '.json') { $Cataleg } else { "$Cataleg.json" }
    $targets = @(Join-Path $EstructDir $n)
}
if ($targets.Count -eq 0) { Write-Host "No s'ha trobat cap cataleg a comprovar." -ForegroundColor Red; exit 1 }

# Tots els URLs d'un cataleg .json, en ordre i sense repetits. Els paragrafs
# d'enllac venen marcats amb "url": true; per si de cas, tambe s'accepta un
# http(s) escrit dins del text d'un paragraf normal.
function Get-CatalegUrls([string]$path) {
    $urls = New-Object System.Collections.Generic.List[string]
    $add = {
        param($u)
        if ([string]::IsNullOrWhiteSpace($u)) { return }
        $u = ([string]$u).Trim().TrimEnd('.', ',', ';', ')')
        if ($u -match '^https?://' -and -not $urls.Contains($u)) { [void]$urls.Add($u) }
    }
    $o = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    # Recorregut en profunditat: cada node pot tenir 'cos' (paragrafs) i 'fills'.
    $visita = {
        param($nodes)
        foreach ($n in @($nodes)) {
            foreach ($par in @($n.cos)) {
                $txt = -join (@($par.runs) | ForEach-Object { [string]$_.t })
                foreach ($m in [regex]::Matches($txt, 'https?://[^\s"<>\]\)]+')) { & $add $m.Value }
            }
            if ($n.fills) { & $visita $n.fills }
        }
    }
    & $visita $o.nodes
    foreach ($par in @($o.intro)) {
        $txt = -join (@($par.runs) | ForEach-Object { [string]$_.t })
        foreach ($m in [regex]::Matches($txt, 'https?://[^\s"<>\]\)]+')) { & $add $m.Value }
    }
    return $urls
}

# Forcem TLS 1.2 (Windows PowerShell 5.1 per defecte pot no negociar-lo).
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol } catch { }

$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
$totalCaiguts = @()

foreach ($t in $targets) {
    if (-not (Test-Path -LiteralPath $t)) { Write-Host "No trobo: $t" -ForegroundColor Red; continue }
    $urls = Get-CatalegUrls $t
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
