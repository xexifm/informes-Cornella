#requires -Version 5.1
<#
  Els ENLLACOS dels catalegs: quins hi ha (pur) i si responen (xarxa).

  NOMES DEFINEIX FUNCIONS: el fan servir Comprova-Enllacos.ps1 (per consola,
  ComprovarEnllacos.bat) i l'eina "Revisar requeriments" (Revisio.ps1). Abans el
  recorregut era a dins de Comprova-Enllacos.ps1 i, a sobre, aquell script
  cridava Read-JsonFile sense carregar Json.ps1: petava a cada cataleg.

  Els enllacos de la FITXA D'AJUDA (ajuda.enllac) tambe hi son: son els que obre
  el boto "Obre la norma" i abans no els comprovava ningu.
#>

# Tots els enllacos d'un cataleg ja parsejat (ConvertFrom-Json), amb el punt on
# son. PURA. Torna [pscustomobject]@{ Url; Punt; Camp } ('text' o 'fitxa'), en
# ordre i sense repetir el mateix URL al mateix punt.
function _EnllacosDeCataleg($o) {
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $o) { return $out.ToArray() }
    $vist = @{}
    $afegeix = {
        param($u, $punt, $camp)
        if ([string]::IsNullOrWhiteSpace($u)) { return }
        $u = ([string]$u).Trim().TrimEnd('.', ',', ';', ')')
        if ($u -notmatch '^https?://') { return }
        $k = $punt + '|' + $u
        if ($vist.ContainsKey($k)) { return }
        $vist[$k] = $true
        [void]$out.Add([pscustomobject]@{ Url = $u; Punt = [string]$punt; Camp = $camp })
    }
    $deParagrafs = {
        param($pars, $punt)
        foreach ($par in @($pars)) {
            if ($null -eq $par) { continue }
            $txt = -join (@($par.runs) | ForEach-Object { [string]$_.t })
            foreach ($m in [regex]::Matches($txt, 'https?://[^\s"<>\]\)]+')) { & $afegeix $m.Value $punt 'text' }
        }
    }
    # Un sub-punt SENSE titol (a Llicencia, el "No es disposa..." d'un punt) porta
    # el nom del punt de sobre: abans l'informe de la revisio el deixava en blanc.
    $visita = $null
    $visita = {
        param($nodes, [string]$pare)
        foreach ($n in @($nodes)) {
            if ($null -eq $n) { continue }
            $punt = [string]$n.titol
            if ([string]::IsNullOrWhiteSpace($punt)) { $punt = $pare }
            & $deParagrafs $n.cos $punt
            if ($null -ne $n.ajuda -and $n.ajuda.enllac) { & $afegeix ([string]$n.ajuda.enllac) $punt 'fitxa' }
            if ($n.fills) { & $visita $n.fills $punt }
        }
    }
    & $deParagrafs $o.intro '(introducció)'
    & $visita $o.nodes ''
    return $out.ToArray()
}

# Els URLs d'un cataleg (ruta del .json), sense repetits.
function Get-CatalegUrls([string]$path) {
    $urls = New-Object System.Collections.Generic.List[string]
    foreach ($e in @(_EnllacosDeCataleg (Read-JsonFile $path))) {
        if (-not $urls.Contains([string]$e.Url)) { [void]$urls.Add([string]$e.Url) }
    }
    return $urls
}

# RESPON L'ENLLAC? @{ Ok; Codi }. HEAD i, si el servidor no l'accepta (405) o
# falla sense dir res, GET: hi ha servidors que no contesten el HEAD.
function Test-EnllacViu([string]$u) {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol } catch { }
    $codi = $null; $ok = $false
    foreach ($method in 'Head', 'Get') {
        try {
            $r = Invoke-WebRequest -Uri $u -Method $method -TimeoutSec 25 -UserAgent $Script:WebUA -UseBasicParsing -MaximumRedirection 5 -UseDefaultCredentials -ErrorAction Stop
            $codi = [int]$r.StatusCode; $ok = ($codi -lt 400); break
        } catch {
            $resp = $null; try { $resp = $_.Exception.Response } catch { }
            if ($resp -and $resp.StatusCode) {
                $codi = [int]$resp.StatusCode
                if ($codi -eq 405 -and $method -eq 'Head') { continue }
                $ok = ($codi -lt 400); break
            } else {
                $codi = 'sense resposta'
                if ($method -eq 'Get') { break }
            }
        }
    }
    return @{ Ok = $ok; Codi = $codi }
}
