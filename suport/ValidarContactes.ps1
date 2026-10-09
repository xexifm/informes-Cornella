#requires -Version 5.1
<#
.SYNOPSIS
  Compara el repas de contactes d'"Actualitzar base" amb la referencia feta a
  ma (contactes-referencia_*.json). Nomes per al PC de l'usuari (no es de la
  suite): germa de ValidarClassificacio.ps1.

.DESCRIPTION
  local\base-dades-activitats\contactes-referencia_2026-10-09.json porta 336
  activitats revisades a ma llegint els mateixos documents contra l'Excel
  d'activitats. Porta DADES PERSONALS i el repositori es public: no pot anar a
  la suite. Aquest script es la manera de mesurar el repas contra la realitat.

  Fa el repas DE DEBO (Invoke-ContactesEscaneig, el mateix codi que "Actualitzar
  base") SENSE les correccions a ma i SENSE escriure la base (-NoDesis): el que
  surt es nomes el que diuen les regles. Aprofita els documents ja llegits de
  contactes-db.json (nomes llegeix els nous). L'atribucio al GIA es la de la
  base d'informes (informes-db.json): cal haver fet "Actualitzar base" abans.

  Compara, activitat a activitat (_CtComparaActivitat): e-mail i telefons del
  titular, nom del representant legal, els e-mails de les persones autoritzades
  i els avisos es_el_tecnic. Les dades que a la referencia diuen
  confianca "probable" van en una llista a part: l'objectiu es 0 discrepancies
  en les "segur".

  Escriu la llista a la consola i a local\base-dades-activitats\
  validacio-contactes_<data>.txt (dins de local\: mai es puja).

  Us: local\ValidarContactes.bat, o des de l'arrel del repositori:
    powershell -NoProfile -ExecutionPolicy Bypass -File suport\ValidarContactes.ps1
    ... -Referencia <fitxer.json> -Informes <carpeta d'informes>
#>
param(
    [string]$Referencia = '',
    [string]$Informes = ''
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

if ([string]::IsNullOrWhiteSpace($Referencia)) {
    $vtCands = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'contactes-referencia_*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    if ($vtCands.Count -eq 0) {
        throw ("No hi ha cap referencia (contactes-referencia_*.json) a:`n  " + $LocalActivitatsDir +
               "`nLa carpeta local\ no es puja al GitHub: si es a l'altre PC, copia-la aqui, o digues on es amb -Referencia <fitxer.json>.")
    }
    $Referencia = $vtCands[0].FullName
}
if (-not [string]::IsNullOrWhiteSpace($Informes)) { $InformesDir = $Informes }
if (-not (Test-Path -LiteralPath $Referencia)) { throw "No trobo la referencia: $Referencia" }
if (-not (_InformesDirAccessible $InformesDir)) { throw "No trobo la carpeta d'informes: $InformesDir" }
$vtBase = Read-JsonFile (Get-InformesDbPath)
if ($null -eq $vtBase) { throw "No hi ha base d'informes (informes-db.json): fes 'Actualitzar base' primer (l'atribucio al GIA es la seva)." }

$vtSortida = Join-Path $LocalActivitatsDir ('validacio-contactes_' + (Get-Date).ToString('yyyyMMdd-HHmm') + '.txt')
$vtLinies = New-Object System.Collections.ArrayList
$vtDiu = { param($t) Write-Host $t; [void]$vtLinies.Add([string]$t) }

# ---- 1. El repas de debo, sense correccions i sense desar -----------------
# Els informes de la base, amb la ruta a la carpeta d'ara (la base pot ser
# d'una altra unitat: I: a la feina, F: fora).
$vtArrelBase = [string](_PropInf $vtBase 'carpeta_arrel')
$vtInformes = @((_FlattenInformesDb $vtBase).Values | ForEach-Object {
    [pscustomobject]@{ Ruta = (Join-Path $InformesDir (_ClauInforme ([string]$_.Ruta) $vtArrelBase)); Gia = [string]$_.Gia }
})
$vtCache = $null; $vtExp = @{}
$vtExcel = Find-LatestActivitatsExcel
if ($null -ne $vtExcel) {
    Write-Host ('Llegint l''Excel d''activitats: ' + $vtExcel.File.Name)
    $vtCache = Initialize-ActivitatsCache $vtExcel.File
    $vtExp = Build-ExpedientToGiaMap $vtCache
}
$vtRes = Invoke-ContactesEscaneig $InformesDir $vtInformes $vtCache $vtExp {
    param($t, $i, $n)
    if ($n -gt 0 -and ($i % 50) -eq 0) { Write-Host ("  ... $i de $n") }
} -SenseCorreccions -NoDesis

# ---- 2. La referencia --------------------------------------------------------
$vtRef = Read-JsonFile $Referencia
if ($null -eq $vtRef -or $null -eq $vtRef.PSObject.Properties['activitats']) { throw "La referencia no porta 'activitats': $Referencia" }
$vtRefActs = ConvertTo-Mapa $vtRef.activitats

# ---- 3. Activitat a activitat ------------------------------------------------
$vtSegur = New-Object System.Collections.ArrayList
$vtProb = New-Object System.Collections.ArrayList
$vtSenseDocs = New-Object System.Collections.ArrayList
$nComp = 0
foreach ($g in @($vtRefActs.Keys | Sort-Object { _GiaNumeric $_ })) {
    if (-not $vtRes.Activitats.Contains([string]$g)) { [void]$vtSenseDocs.Add([string]$g); continue }
    $nComp++
    foreach ($d in @(_CtComparaActivitat $vtRefActs[$g] $vtRes.Activitats[[string]$g])) {
        $t = ('  GIA ' + $g + '  ' + $d.Que + "`n      referencia: " + $d.Ref + "`n      ara:        " + $d.Ara)
        if ($d.Confianca -eq 'probable') { [void]$vtProb.Add($t) } else { [void]$vtSegur.Add($t) }
    }
}

# ---- 4. El resum, i la LLISTA ------------------------------------------------
& $vtDiu ('Validacio del repas de contactes ' + $Script:ContactesVersio + '  -  ' + (Get-Date).ToString('dd/MM/yyyy HH:mm'))
& $vtDiu ('Carpeta d''informes: ' + $InformesDir)
& $vtDiu ('Referencia:          ' + $Referencia)
& $vtDiu ('Tecnics coneguts:    ' + $(if ([string]$vtRes.TecnicsFitxer) { [string]$vtRes.TecnicsFitxer } else { '(cap fitxer tecnics-coneguts_*.json)' }))
if ([string]$vtRes.TecnicsError) { & $vtDiu ('AVIS: ' + [string]$vtRes.TecnicsError) }
& $vtDiu ''
& $vtDiu ('Documents: ' + $vtRes.NDocs + '   llegits ara: ' + $vtRes.Llegits + '   no llegibles: ' + $vtRes.NNoLlegibles + '   sense GIA: ' + @($vtRes.SenseGia).Count)
& $vtDiu ('Activitats a la referencia: ' + $vtRefActs.Count + '   comparades: ' + $nComp + '   sense cap document al repas: ' + $vtSenseDocs.Count)
& $vtDiu ('Discrepancies en dades "segur": ' + $vtSegur.Count + '   (en dades "probable", a part: ' + $vtProb.Count + ')')
& $vtDiu ''
& $vtDiu '== DISCREPANCIES (segur) =='
foreach ($t in $vtSegur) { & $vtDiu $t }
& $vtDiu ''
& $vtDiu '== Discrepancies en dades "probable" (no compten per a l''objectiu) =='
foreach ($t in $vtProb) { & $vtDiu $t }
if ($vtSenseDocs.Count -gt 0) {
    & $vtDiu ''
    & $vtDiu '== A la referencia pero sense cap document al repas (no s''han trobat o no tenen GIA) =='
    & $vtDiu ('  ' + ($vtSenseDocs -join ', '))
}
& $vtDiu ''
& $vtDiu '== Documents no llegibles (escanejats o sense text: DiagnosticPdf.bat per mirar-ne un) =='
foreach ($k in @($vtRes.NoLlegibles)) { & $vtDiu ('  ' + $k) }
& $vtDiu ''
& $vtDiu '== Documents dels quals no se sap el GIA (carpeta de diverses activitats) =='
foreach ($k in @($vtRes.SenseGia)) { & $vtDiu ('  ' + $k) }
[void](New-Item -ItemType Directory -Path (Split-Path -Parent $vtSortida) -Force)
[System.IO.File]::WriteAllText($vtSortida, ($vtLinies -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
Write-Host ''
Write-Host ('Desat a: ' + $vtSortida)
