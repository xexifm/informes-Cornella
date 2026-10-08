<#
  MapaHtml.ps1 - El que fan igual els mapes HTML de 'rutes/' que surten d'una
  plantilla (CoordenadesMapa.html, PlanolMapa.html): omplir les marcas {{nom}}
  i posar-hi les dades en JSON dins d'un <script>.

  Vivia a Coordenades.ps1; amb el Planol activitats ja son dos mapes, i
  copiar-ho era tornar a obrir les dues trampes que aixo tanca (la doble
  substitucio i el '</script>' a les dades).

  NOMES DEFINEIX FUNCIONS. ASCII pur.
#>

# Omple les marques {{nom}} d'una plantilla en UNA sola passada. PURA.
#
# Una sola passada a posta: si es fes un .Replace() per marca, un valor que
# portes el text '{{dbEnc}}' (ve de l'Excel) quedaria substituit pel seguent.
# I es fa a ma, sense [regex]::Replace amb un scriptblock: els scriptblocks
# convertits a delegat no veuen les variables locals igual a totes les versions
# del PowerShell (la trampa de les closures del CLAUDE.md).
#
# Una marca que no te valor es un error de programacio, no de dades: llanca,
# perque una pagina amb '{{nTot}}' a la vista no ha de sortir mai.
function Expand-PlantillaHtml([string]$plantilla, $valors) {
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    foreach ($m in [regex]::Matches($plantilla, '\{\{([A-Za-z]+)\}\}')) {
        $nom = $m.Groups[1].Value
        if (-not $valors.ContainsKey($nom)) { throw "Plantilla del mapa: falta el valor de {{$nom}}" }
        [void]$sb.Append($plantilla, $pos, $m.Index - $pos)
        [void]$sb.Append([string]$valors[$nom])
        $pos = $m.Index + $m.Length
    }
    [void]$sb.Append($plantilla, $pos, $plantilla.Length - $pos)
    return $sb.ToString()
}

# EL FONS DELS MAPES (MapaFons.js), per posar-lo TAL QUAL dins d'un <script>.
# Un sol fitxer per als tres mapes: quan OpenStreetMap va deixar de servir les
# rajoles a les pagines obertes des del disc, cada mapa tenia la seva copia de
# l'adreca i van caure tots tres alhora (octubre 2026).
function Get-MapaFonsJs {
    $js = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'MapaFons.js'), [System.Text.Encoding]::UTF8)
    return $js.Replace('</', '<\/')
}

# ----------------------------------------------------------------------------
# LA BIBLIOTECA DEL MAPA (Leaflet): LA VERSIO I EL SEU SRI, EN UN SOL LLOC
# ----------------------------------------------------------------------------
# Estaven fixats a QUATRE llocs independents -els dos mapes d'aqui, l'HTML que
# es fa Ruta.ps1 i docs/precintades.html-, cada un amb la seva versio i el seu
# hash. I dos dels quatre no tenien el respatller de jsDelivr.
#
# Que passa si divergeixen: RES VISIBLE. El navegador es troba un hash que no
# quadra, no carrega el fitxer i NO diu res; la pagina surt amb "No s'ha pogut
# carregar el mapa" com si fos un problema de connexio. Es el defecte tipic
# d'aquest projecte: no falla, empitjora en silenci.
#
# El hash es del paquet de npm, que unpkg i jsDelivr serveixen byte a byte
# igual: per aixo el respatller pot dur EL MATEIX SRI sense afluixar-lo.
$Script:LeafletVersio = '1.9.4'
$Script:LeafletSriCss = 'sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY='
$Script:LeafletSriJs  = 'sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo='

# El full d'estil, amb el segon intent a jsDelivr si unpkg no respon.
function Get-MapaLeafletCss {
    $v = $Script:LeafletVersio
    return @"
<link rel="stylesheet" href="https://unpkg.com/leaflet@$v/dist/leaflet.css"
      integrity="$($Script:LeafletSriCss)" crossorigin=""
      onerror="this.onerror=null;this.href='https://cdn.jsdelivr.net/npm/leaflet@$v/dist/leaflet.css';"/>
"@
}

# La biblioteca, amb el segon intent a jsDelivr. document.write i no un <script>
# creat a ma: aixi es carrega ABANS que el codi del mapa, que ve just despres i
# el necessita.
function Get-MapaLeafletJs {
    $v = $Script:LeafletVersio
    return @"
<script src="https://unpkg.com/leaflet@$v/dist/leaflet.js"
        integrity="$($Script:LeafletSriJs)" crossorigin=""></script>
<script>
if (typeof L === 'undefined') {
  document.write('<script src="https://cdn.jsdelivr.net/npm/leaflet@$v/dist/leaflet.js" ' +
    'integrity="$($Script:LeafletSriJs)" crossorigin=""><\/script>');
}
</script>
"@
}

# Llegeix una plantilla en UTF-8 EXPLICIT (el Windows PowerShell 5.1, sense dir-li
# res, la llegiria com a ANSI i els accents sortirien com 'Ã§') i l'omple.
# {{fonsJs}} (el fons del mapa) i el Leaflet els posa aqui, per a totes: no son
# cosa de cap eina.
function Get-PlantillaHtml([string]$ruta, $valors) {
    $plantilla = [System.IO.File]::ReadAllText($ruta, [System.Text.Encoding]::UTF8)
    $tots = @{}
    foreach ($k in @($valors.Keys)) { $tots[$k] = $valors[$k] }
    if (-not $tots.ContainsKey('fonsJs'))     { $tots['fonsJs']     = Get-MapaFonsJs }
    if (-not $tots.ContainsKey('leafletCss')) { $tots['leafletCss'] = Get-MapaLeafletCss }
    if (-not $tots.ContainsKey('leafletJs'))  { $tots['leafletJs']  = Get-MapaLeafletJs }
    return (Expand-PlantillaHtml $plantilla.TrimEnd() $tots)
}

# Un JSON per posar dins d'un <script>.
#   - Una LLISTA sempre (amb -Llista): el guard mira la SORTIDA i no el
#     nombre d'elements, perque segons la versio del PowerShell ConvertTo-Json
#     desembolcalla o no un array d'un sol element (al PC de l'usuari, no ho
#     feia i el mapa d'una activitat sortia [[{...}]]).
#   - '</' -> '<\/': un '</script>' a les dades tancaria l'etiqueta i la pagina
#     no arrencaria. Per al JSON es el mateix caracter.
function ConvertTo-JsonScript($valor, [switch]$Llista, [int]$Fondaria = 6) {
    $json = ConvertTo-Json $valor -Depth $Fondaria -Compress
    if ($Llista) {
        if ([string]::IsNullOrWhiteSpace($json) -or $json -eq 'null') { $json = '[]' }
        elseif (-not $json.TrimStart().StartsWith('[')) { $json = "[$json]" }
    }
    return $json.Replace('</', '<\/')
}
