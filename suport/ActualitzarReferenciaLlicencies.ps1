#requires -Version 5.1
<#
.SYNOPSIS
  Posa a la classificacio de referencia els dos estats nous dels favorables de
  llicencia: "Favorable pre-llicència (Requeriment)" i "Favorable
  post-llicència (Favorable)". Nomes per al PC de l'usuari (no es de la suite).

.DESCRIPTION
  La classificacio feta a ma (local\base-dades-activitats\
  classificacio-informes_*.json) es va fer quan aquests informes eren
  'Favorable' i prou. Des del classificador 2026-10-09.1 en surten pre o post,
  i ValidarClassificacio.ps1 els donaria tots per discrepancia. L'usuari:
  "s'ha de tocar la referencia i posar-los on toca".

  Per a cada entrada de la referencia que diu 'Favorable' llegeix l'informe de
  la carpeta d'informes i, si es un favorable de LLICENCIA (el que el
  classificador diu 'llicfav', o el tipus que ja portava la referencia), el
  canvia a pre (la conclusio diu "a l'espera de rebre la documentacio") o a
  post. La resta d'entrades no es toca.

  NO SOBREESCRIU RES: desa una referencia NOVA al costat
  (classificacio-informes_<data>_llic.json), que es la que agafara
  ValidarClassificacio perque surt la primera per ordre de nom, i la llista de
  canvis a referencia-llicencies_<data-hora>.txt. Tot dins de local\ (dades
  personals: mai es puja).

  Us (des de l'arrel del repositori):
    powershell -NoProfile -ExecutionPolicy Bypass -File suport\ActualitzarReferenciaLlicencies.ps1
    ... -Classificacio <fitxer.json> -Informes <carpeta d'informes>
#>
param(
    [string]$Classificacio = '',
    [string]$Informes = ''
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

if ([string]::IsNullOrWhiteSpace($Classificacio)) {
    $arCands = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'classificacio-informes_*.json' -File -ErrorAction SilentlyContinue |
                 Sort-Object Name -Descending)
    if ($arCands.Count -eq 0) { throw ("No hi ha cap classificacio (classificacio-informes_*.json) a:`n  " + $LocalActivitatsDir) }
    $Classificacio = $arCands[0].FullName
}
if (-not [string]::IsNullOrWhiteSpace($Informes)) { $InformesDir = $Informes }
if (-not (Test-Path -LiteralPath $Classificacio)) { throw "No trobo la classificacio: $Classificacio" }
if (-not (_InformesDirAccessible $InformesDir)) { throw "No trobo la carpeta d'informes: $InformesDir" }

$arObj = Read-JsonFile $Classificacio
if ($null -eq $arObj) { throw "La classificacio no es un JSON valid: $Classificacio" }
$arEntrades = @()
if ($arObj -is [System.Collections.IList]) { $arEntrades = @($arObj) }
else {
    foreach ($p in $arObj.PSObject.Properties) {
        $v = @($p.Value)
        if ($v.Count -gt 0 -and $null -ne $v[0] -and $null -ne $v[0].PSObject.Properties['ruta_relativa']) { $arEntrades = $v; break }
    }
}
if ($arEntrades.Count -eq 0) { throw "La classificacio no porta cap entrada amb 'ruta_relativa'." }

$arCanvis = New-Object System.Collections.ArrayList
$arNoTrobats = New-Object System.Collections.ArrayList
$arWord = $null
$nFav = 0
try {
    foreach ($e in $arEntrades) {
        if ([string](_PropInf $e 'conclusio_breu') -ne $Script:EstatFavorable) { continue }
        $nFav++
        $rel = ([string]$e.ruta_relativa).Trim('\', '/')
        $ruta = Join-Path $InformesDir ($rel -replace '[\\/]', [string][System.IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $ruta)) { [void]$arNoTrobats.Add($rel); continue }
        $f = Get-Item -LiteralPath $ruta
        if ($f.Extension -ieq '.doc' -and $null -eq $arWord) { $arWord = New-WordApp -Opcional }
        $lines = @()
        try { $lines = @(_ReadInformeParagraphs $f $arWord) } catch { $lines = @() }
        $cl = _ClassificaInforme $lines $f.Name (_ExtractExpedient $lines)
        $esLlic = ([string]$cl.Tipus -eq 'llicfav') -or ([string](_PropInf $e 'tipus') -eq 'llicfav')
        if (-not $esLlic) { continue }
        # El mateix criteri que _ClassificaInforme.
        $nou = if ((_ConclNorm $cl.Conclusio) -match 'a lespera de rebre') { $Script:EstatFavorablePre } else { $Script:EstatFavorablePost }
        $e.conclusio_breu = $nou
        if ($null -eq $e.PSObject.Properties['tipus']) { Add-Member -InputObject $e -NotePropertyName tipus -NotePropertyValue 'llicfav' }
        elseif ([string]$e.tipus -eq '') { $e.tipus = 'llicfav' }
        [void]$arCanvis.Add(('  ' + $rel + "`n      Favorable -> " + $nou))
    }
} finally {
    if ($null -ne $arWord) { try { $arWord.Quit() } catch { } }
}

$arData = (Get-Date).ToString('yyyy-MM-dd')
# Al costat de la de partida (normalment local\base-dades-activitats).
$arDir = Split-Path -Parent (Resolve-Path -LiteralPath $Classificacio).Path
$arSortida = Join-Path $arDir ('classificacio-informes_' + $arData + '_llic.json')
$arLog = Join-Path $arDir ('referencia-llicencies_' + (Get-Date).ToString('yyyyMMdd-HHmm') + '.txt')
$arLinies = New-Object System.Collections.ArrayList
$arDiu = { param($t) Write-Host $t; [void]$arLinies.Add([string]$t) }
& $arDiu ('Referencia de partida: ' + $Classificacio)
& $arDiu ('Carpeta d''informes:   ' + $InformesDir)
& $arDiu ('Entrades: ' + $arEntrades.Count + '   que deien Favorable: ' + $nFav + '   canviades: ' + $arCanvis.Count + '   no trobades a la carpeta: ' + $arNoTrobats.Count)
& $arDiu ''
& $arDiu '== Canviades =='
foreach ($t in $arCanvis) { & $arDiu $t }
if ($arNoTrobats.Count -gt 0) {
    & $arDiu ''
    & $arDiu '== Deien Favorable pero l''informe no es a la carpeta (no s''han tocat) =='
    foreach ($t in $arNoTrobats) { & $arDiu ('  ' + $t) }
}
if ($arCanvis.Count -gt 0) {
    Write-JsonFile $arSortida $arObj 8
    & $arDiu ''
    & $arDiu ('Referencia nova: ' + $arSortida)
    & $arDiu '(la de partida no s''ha tocat; ValidarClassificacio agafara la nova)'
} else {
    & $arDiu ''
    & $arDiu 'Cap canvi: no s''ha desat cap referencia nova.'
}
[System.IO.File]::WriteAllText($arLog, ($arLinies -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
Write-Host ('Desat a: ' + $arLog)
