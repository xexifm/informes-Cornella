#requires -Version 5.1
<#
.SYNOPSIS
  El comu de l'ARRENCADA de les eines de 'rutes/' (Coordenades, Planol
  activitats, Precintades).

.DESCRIPTION
  Les tres corren en un PROCES PROPI i totes tres comencen igual: carregar
  Ruta.ps1 nomes per tenir-ne les funcions (cerca de l'Excel, Find-HeaderColumn,
  la conversio UTM, el format d'adreca) i perque arrossegui config.ps1,
  Excel.ps1, Json.ps1 i UiFinestra.ps1.

  Per fer-ho cal enganyar-lo amb $env:RUTA_TEST = '1', i DESPRES tornar la
  variable com estava: si es queda posada, la resta del proces -i qualsevol
  cosa que s'hi llanci- es pensa que corre en mode de proves i no obre cap
  finestra. Es un defecte que no peta: l'eina simplement no fa res.

  Aquell ball de desar/posar/restaurar estava escrit TRES vegades, identic
  menys el nom de la variable ($_prevRutaTest, $_prevRutaTestPlanol,
  $_prevRutaTestCoord). Ara la part delicada -el cas del $null, que vol dir
  "abans no hi era i s'ha d'esborrar", no "posar-hi cadena buida"- es a un sol
  lloc.

  PER QUE EL DOT-SOURCE NO ENTRA AQUI DINS, i esta MESURAT: fer
  ". fitxer.ps1" DINS D'UNA FUNCIO carrega el fitxer a l'ambit de la FUNCIO, no
  al de l'script, i les funcions desapareixen en tornar. Comprovat amb un modul
  de prova: carregat des de dins d'una funcio, Get-Command no el troba;
  carregat a l'ambit de l'script, si. Per tant el ". Ruta.ps1" i el try/finally
  s'han de quedar al cos de cada eina, i aqui nomes hi ha el que es pot
  compartir.

  NO ES CARREGA RES EN AQUEST FITXER: nomes defineix funcions (regla 5). Les
  tres eines el carreguen abans que res.
#>

# Posa el mode headless de Ruta.ps1 i TORNA el valor que hi havia, que s'ha de
# passar tal qual a Exit-RutaHeadless. S'usa sempre aixi:
#
#     $prevRuta = Enter-RutaHeadless
#     try { . (Join-Path $ScriptRoot 'Ruta.ps1') } finally { Exit-RutaHeadless $prevRuta }
#
# El try/finally no es decoracio: si Ruta.ps1 peta carregant-se, sense ell la
# variable es quedaria posada i la resta del proces correria en silenci.
function Enter-RutaHeadless {
    $prev = $env:RUTA_TEST
    $env:RUTA_TEST = '1'
    return $prev
}

# Torna $env:RUTA_TEST com estava. El cas que importa es el $null: vol dir que
# la variable NO existia, i llavors s'ha d'ESBORRAR. Assignar-hi '' deixaria una
# variable d'entorn buida, que no es el mateix i que [bool] llegeix com a fals
# per casualitat, no per disseny.
function Exit-RutaHeadless($prev) {
    if ($null -eq $prev) {
        Remove-Item Env:\RUTA_TEST -ErrorAction SilentlyContinue
    } else {
        $env:RUTA_TEST = $prev
    }
}

# WinForms per a les eines que obren finestra (Coordenades i Planol; Precintades
# no en te). L'ordre importa: l'EnableVisualStyles va DESPRES dels dos Add-Type.
# El criden dins del seu "if (-not $headless)": en headless no s'ha de tocar.
function Initialize-EinaWinForms {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}
