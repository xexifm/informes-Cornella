#requires -Version 5.1
<#
.SYNOPSIS
  Fa UNA passada d'"Actualitzar base" sense interficie (mode automatic).

.DESCRIPTION
  El llanca el MENU del programa (l'interruptor A/M de sota la rajola
  "Actualitzar base") quan toca: cada dia a les 14:00 amb el programa obert i,
  si aquell venciment no s'ha arribat a servir, en obrir el programa.

  Fa NOMES una passada i surt. Mateix patro que CopiaInformesAuto.ps1: un
  proces a part perque recorrer la carpeta d'informes (unitat de xarxa) no
  deixi el menu congelat.

  No ensenya res i no pregunta res. Tot el que passa queda al registre
  (%LOCALAPPDATA%\InformesCornella\informes-db-log.txt). Si el boto
  "Actualitzar base" esta escanejant alhora, no en fa un altre al damunt
  (Invoke-InformesDbAuto ho mira amb el mateix mutex).
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# El motor NOMES com a biblioteca: som un proces de consola, sense finestres.
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

try {
    [void](Invoke-InformesDbAuto)
    exit 0
} catch {
    _BaseAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                  $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
}
