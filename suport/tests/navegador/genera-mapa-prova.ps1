<#
  genera-mapa-prova.ps1 - Genera els mapes de Coordenades que fa servir la prova
  del navegador (prova-mapa-coordenades.mjs). NOMES funcions pures: ni Excel ni
  Cadastre ni finestres.

  Escriu tres mapes:
    a/Coordenades_A.html  base '2026-08-18 ACTIVITATS.xls', sis activitats
    b/Coordenades_B.html  la MATEIXA base, amb un nom de fitxer i una carpeta
                          diferents: es el que passa cada cop que es torna a
                          generar el mapa, i el repas s'hi ha de veure.
    c/Coordenades_C.html  una ALTRA base: el repas de l'A no s'hi ha de veure.
#>
param([Parameter(Mandatory = $true)][string]$Dir)

$env:COORDENADES_TEST = '1'
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }
. (Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path '..' (Join-Path 'rutes' 'Coordenades.ps1')))

function _Item($id, [double]$x, [double]$y, [double]$xf, [double]$yf, $prec, $adreca, $act, $tit = '') {
    $a = Convert-UtmToLatLon $x $y 31 $true
    $b = Convert-UtmToLatLon $xf $yf 31 $true
    return [pscustomobject]@{
        Id = [string]$id; Zona = (Get-ZonaDeCoord $x $y); Rc = '4091106DF2749A0001XX'
        Adreca = $adreca; Activitat = $act; Titular = $tit
        XExcel = $x; YExcel = $y; LatExcel = $a.Lat; LonExcel = $a.Lon
        XFacana = $xf; YFacana = $yf; LatFacana = $b.Lat; LonFacana = $b.Lon
        Precisio = $prec
    }
}

# Quatre apilades al mateix punt (una de cada color) i dues en un altre lloc.
$x0 = 421968.09; $y0 = 4578100.5
$items = @(
    (_Item 101 $x0 $y0 ($x0 + 40) ($y0 + 25) 'facana'         'C/ Cadis 19'                 'BAR' 'EL RACO DE CADIS SL')
    (_Item 102 $x0 $y0 ($x0 - 35) ($y0 + 10) 'facana-dubtosa' 'C/ Cadis 1'                  ('CAF' + [char]0x00C8))
    (_Item 103 $x0 $y0 ($x0 + 15) ($y0 - 45) 'facana-aprox'   'C/ Huelva 3'                 'FORN')
    (_Item 104 $x0 $y0 $x0 $y0                'cadastre'       'C/ Falsa 1 </script><b>'     'BOTIGA')
    (_Item 105 ($x0 + 300) ($y0 + 300) ($x0 + 320) ($y0 + 310) 'facana' ('Pla' + [char]0x00E7 + 'a Catalunya 2') 'PERRUQUERIA')
    (_Item 106 ($x0 + 300) ($y0 + 300) ($x0 + 280) ($y0 + 290) 'facana' 'Pla Catalunya 4'   'FARMACIA')
)
$portals = @(
    [pscustomobject]@{ Numero = '19'; Via = 'CL CADIS'; Lat = $items[0].LatFacana; Lon = $items[0].LonFacana }
    [pscustomobject]@{ Numero = '1';  Via = 'CL CADIS'; Lat = $items[1].LatFacana; Lon = $items[1].LonFacana }
)

$base = '2026-08-18 ACTIVITATS.xls'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$escriu = {
    param($sub, $nom, $its, $font)
    $d = Join-Path $Dir $sub
    New-Item -ItemType Directory -Path $d -Force | Out-Null
    $h = Build-CoordenadesHtml $its "Base de dades: $font" 'F2 (apilades)' $font $portals
    [System.IO.File]::WriteAllText((Join-Path $d $nom), $h, $utf8)
}
& $escriu 'a' 'Coordenades_A.html' $items $base
& $escriu 'b' 'Coordenades_B.html' @($items[0], $items[4]) $base
& $escriu 'c' 'Coordenades_C.html' $items '2026-10-01 ACTIVITATS.xls'
