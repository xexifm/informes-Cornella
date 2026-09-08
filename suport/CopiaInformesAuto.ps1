#requires -Version 5.1
<#
.SYNOPSIS
  Fa UNA passada de "Copiar informes" sense interficie (mode automatic).

.DESCRIPTION
  El llanca el MENU del programa (l'interruptor A/M de sota la rajola "Copiar
  informes") quan toca: cada dia a les 14:30 amb el programa obert i, si aquell
  venciment no s'ha arribat a servir, en obrir el programa.

  Fa NOMES una passada i surt: no es queda en segon pla. Mateix patro que
  GeneraVistes.ps1 i mobil/Vigilant.ps1.

  PER QUE UN PROCES A PART: recorrer la carpeta d'informes (recursiva, en una
  unitat de xarxa) pot trigar. Fet dins del menu, la finestra es quedaria
  congelada -que es justament el contrari de "no es veura res"- i, amb els
  DoEvents de WinForms, l'usuari podria obrir una eina a mig copiar.

  No ensenya res i no pregunta res. Tot el que passa queda al registre
  (%LOCALAPPDATA%\InformesCornella\copia-informes-log.txt).
#>

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem el motor NOMES com a biblioteca: som un proces de consola, no volem
# WinForms ni obrir cap finestra.
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

# NOMES UN A LA VEGADA. El menu ja mira si el proces anterior encara corre, pero
# el seu handle es d'aquella instancia del programa: si se n'obre una altra (o
# el mateix menu es torna a obrir despres d'un tancament brusc), dos processos
# copiant a la mateixa carpeta es trepitjarien l'estat i el segon reescriuria
# 'copiat_el' amb una passada a mitges. Aqui NO s'espera: si ja n'hi ha un fent
# la feina, aquest no hi te res a fer.
$mutex = $null
$tinc = $false
try {
    $mutex = New-Object System.Threading.Mutex($false, 'Global\InformesCornella.CopiaInformesAuto')
    try { $tinc = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $tinc = $true }
} catch { $mutex = $null; $tinc = $true }
if (-not $tinc) {
    _CopiaAutoLog 'Passada automatica: ja n hi ha una en marxa, no es fa res.'
    exit 0
}

try {
    [void](Invoke-CopiarInformesAuto)
    exit 0
} catch {
    _CopiaAutoLog ("ERROR no controlat: " + $_.Exception.Message + ' @ ' +
                   $_.InvocationInfo.ScriptName + ':' + $_.InvocationInfo.ScriptLineNumber)
    exit 1
} finally {
    if ($null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { }; try { $mutex.Dispose() } catch { } }
}
