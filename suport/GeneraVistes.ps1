#requires -Version 5.1
<#
.SYNOPSIS
  Regenera les VISTES en Word (.docx) de tots els catalegs, des dels JSON.

.DESCRIPTION
  Els .docx d'ESTRUCTURALS ja no serveixen per generar informes: la font de
  veritat son els .json. Aquest script els torna a escriure perque l'usuari
  pugui consultar el contingut sencer de cada cataleg (tots els requeriments,
  totes les conclusions...) sense obrir el programa, amb els titols de Word.

  NO toca '0 CAPCALERA.docx', que si que es una plantilla de veritat.

  El crida Actualitzar.bat. Necessita Word instal·lat; si no hi es, avisa i
  no fa res (mai atura l'actualitzacio).
#>

$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem el MOTOR nomes com a biblioteca (cap finestra, cap WinForms).
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

# NOMES UN A LA VEGADA. Ara aquest script el llanca tambe l'editor de catalegs
# en segon pla, en tancar-lo, i podria coincidir amb el pas 4b de l'Actualitzar.bat.
# Dos processos conduint el Word alhora (Documents.Add + SaveAs) es la manera de
# treure'n una vista a mitges. S'ESPERA el que hi hagi -aixo corre en segon pla i
# no bloqueja ningu-, i si no arriba a entrar en dos minuts es deixa corrar: qui
# te el pany ja esta fent la mateixa feina.
$mutex = $null
$tinc = $false
try {
    $mutex = New-Object System.Threading.Mutex($false, 'Global\InformesCornella.GeneraVistes')
    try { $tinc = $mutex.WaitOne(120000) } catch [System.Threading.AbandonedMutexException] { $tinc = $true }
} catch { $mutex = $null; $tinc = $true }
if (-not $tinc) {
    Write-Host "  (les vistes ja les esta generant un altre proces)"
    exit 0
}

try {
    $n = Invoke-ExportarVistesWord
    if ($n -gt 0) { Write-Host ("  {0} vistes de cataleg actualitzades." -f $n) }
} catch {
    Write-Host ("  Avis: no s'han pogut generar les vistes (" + $_.Exception.Message + ").")
} finally {
    if ($null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { }; try { $mutex.Dispose() } catch { } }
}
exit 0
