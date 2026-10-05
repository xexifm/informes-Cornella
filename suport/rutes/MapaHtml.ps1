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

# Llegeix una plantilla en UTF-8 EXPLICIT (el Windows PowerShell 5.1, sense dir-li
# res, la llegiria com a ANSI i els accents sortirien com 'Ã§') i l'omple.
function Get-PlantillaHtml([string]$ruta, $valors) {
    $plantilla = [System.IO.File]::ReadAllText($ruta, [System.Text.Encoding]::UTF8)
    return (Expand-PlantillaHtml $plantilla.TrimEnd() $valors)
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
