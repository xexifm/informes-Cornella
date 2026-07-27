#requires -Version 5.1
<#
.SYNOPSIS
  Punt d'entrada del generador d'informes de l'Ajuntament de Cornella.

.DESCRIPTION
  Aquest fitxer NOMES ARRENCA el programa. Tota la logica (funcions, rutes,
  configuracio) viu a Motor.ps1, que aqui es carrega amb dot-source.

  La separacio existeix perque el motor es COMPARTIT: el reutilitzen
  mobil/Vigilant.ps1, mobil/ExportaDades.ps1 i les proves de suport/tests/.
  Abans tot era un sol fitxer, i per carregar-lo sense que s'executes calia
  fer-lo passar per proves ($env:GENINFORME_TEST = '1'); ara qui vol la
  biblioteca carrega Motor.ps1 i prou.

  El llanca GenerarInforme.bat (a l'arrel del clone) a traves de
  GenerarInforme.vbs, que l'obre sense cap finestra de consola.

.PARAMETER DesDePaquet
  Mode no interactiu: genera l'informe directament des d'un paquet JSON (el
  mateix model que lastreport.json) en lloc d'obrir l'assistent de passos. El
  paquet l'omple el formulari web del mobil i el porta fins aqui el vigilant
  (Vigilant.ps1). Necessita Word (com el flux normal), pero NO obre cap
  finestra. Si no s'indica, el programa funciona com sempre.
#>

param(
    [string]$DesDePaquet
)

$ErrorActionPreference = 'Stop'

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Motor.ps1')

# En mode proves nomes es volien les definicions del motor: no arrenquem res.
if ($Script:HeadlessTest) { return }

if ($DesDePaquet) {
    # Mode no interactiu (vigilant del mobil): NO apliquem el candau d'una sola
    # instancia (poden processar-se diversos paquets alhora i no obre cap
    # finestra).
    Invoke-GenerateFromPaquet $DesDePaquet
}
else {
    # Nomes una instancia: si ja n'hi ha una d'oberta, l'enfoquem i sortim.
    if (Enter-SingleInstance) { Main }
}
