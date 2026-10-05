<#
  genera-planol-prova.ps1 - Genera el planol que fa servir prova-planol.mjs, amb
  les funcions de debo de PlanolDades.ps1 i les dades de prova de
  run-tests-planol.ps1 (fetes a ma: cap dada real).
#>
param([Parameter(Mandatory = $true)][string]$Dir)

$env:PLANOL_TEST = '1'
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }
. (Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path '..' (Join-Path 'rutes' 'Planol.ps1')))
$dades = Join-Path (Split-Path -Parent $PSScriptRoot) 'dades'

function _E($id, $rc, $act, $x, $y, $buit, $local, $adr) {
    return [pscustomobject]@{ IdEst = $id; Rc = $rc; IdActivitat = $act; Local = $local; Bloc = ''; Escala = ''; Pis = ''; Porta = ''
                              Buit = $buit; UtmX = $x; UtmY = $y; Adreca = $adr }
}
function _A($id, $nom, $act, $rc, $prec, $x, $y, $adr) {
    return [pscustomobject]@{ Id = $id; Nom = $nom; Activitat = $act; Rc = $rc; Precinte = $prec; Adreca = $adr; UtmX = $x; UtmY = $y }
}
$ests = @(
    (_E '1' '2295827DF2729E0011RQ' '1447' 421975.0 4579505.0 $false '' 'C CADIS 19')
    (_E '2' '2295827DF2729E0008RQ' '1403' 421975.0 4579505.0 $false '5' 'C CADIS 19')
    (_E '3' '2295827DF2729E0003XL' '' 421975.0 4579505.0 $true '' 'C CADIS 19')
    (_E '4' '4091106DF2749A0006XJ' '9' 423912.16 4578928.25 $true '' 'CTRA HOSPITALET 147')
    (_E '5' '1111111DF1111A0001AA' '' 422300.0 4579300.0 $true '2' 'C BUIT 1')
    (_E '6' '3085213DF2738E0001AB' '20' 422800.0 4579200.0 $false '' 'PG FERROCARRILS 177')
)
$acts = @{
    '1447' = (_A '1447' 'EL RACO' 'BAR' '2295827DF2729E0011RQ' $true 421975.0 4579505.0 'C CADIS 19')
    '1403' = (_A '1403' '' 'TALLER' '2295827DF2729E0008RQ' $false 421975.0 4579505.0 'C CADIS 19')
    '9'    = (_A '9' 'ACME' 'OFICINES' '4091106DF2749A0006XJ' $false 423912.16 4578928.25 'CTRA HOSPITALET 147')
    '20'   = (_A '20' '' 'MAGATZEM' '3085213DF2738E0001AB' $false 422800.0 4579200.0 'PG FERROCARRILS 177')
}
$estats = @{
    '1403' = [pscustomobject]@{ Estat = 'Requeriment'; NInformes = 2 }
    '9'    = [pscustomobject]@{ Estat = 'Favorable'; NInformes = 1 }
}
$unitats = @{ '2295827DF2729E0011RQ' = (ConvertFrom-CatastroDnprcXml ([System.IO.File]::ReadAllText((Join-Path $dades 'dnprc-exemple.xml')))) }
$geos = @{ '2295827DF2729E' = @(ConvertFrom-CatastroParcelXml ([System.IO.File]::ReadAllText((Join-Path $dades 'wfsCP-exemple.xml')))) }

$model = Build-PlanolModel $ests $acts $estats $unitats
$mapa = ConvertTo-PlanolDadesMapa $model $geos
$meta = [pscustomobject]@{ BaseActivitats = '2026-10-05 ACTIVITATS.xls'; BaseEstabliments = '2026-10-05 ESTABLIMENTS.xls'
                           BaseInformes = 'Base d''informes de prova'; Avisos = @('Avis de prova </script> amb <b>') }
New-Item -ItemType Directory -Path $Dir -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path $Dir 'Planol.html'), (Build-PlanolHtml $mapa $meta), (New-Object System.Text.UTF8Encoding($false)))
