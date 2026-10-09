#requires -Version 5.1
<#
.SYNOPSIS
  Compara el que treu "Actualitzar base" amb la classificacio feta a ma de la
  carpeta REAL d'informes. Nomes per al PC de l'usuari (no es de la suite).

.DESCRIPTION
  El 7 d'octubre de 2026 es van llegir un per un els 802 informes de la carpeta
  real (427 activitats) i se'n va desar la classificacio correcta a
  local\base-dades-activitats\classificacio-informes_2026-10-07.json (cada
  entrada: ruta_relativa, conclusio_breu, ignorat, tipus, nota...). Aquell
  fitxer porta DADES PERSONALS i el repositori es public: no pot anar a la
  suite. Aquest script es la manera de mesurar el classificador contra la
  realitat (regla 7 de suport/CLAUDE.md: mesura-ho, no ho dedueixis).

  Fa l'escaneig DE DEBO (Invoke-InformesDbEscaneig, el mateix codi que el boto)
  pero en una carpeta temporal i SENSE la base d'ara, o sigui sense cap
  correccio a ma: el que surt es nomes el que diu el classificador. La base de
  l'usuari no es toca.

  Despres compara, informe a informe, la conclusio breu i l'ignorat, i activitat
  a activitat, l'ESTAT: el que surt ara contra el que sortiria amb la
  classificacio correcta (la mateixa regla, _EstatActualActivitat, amb els
  valors bons). Es deixen fora les entrades amb una 'nota' que comenca per
  "DUBTE"; les que diuen que son judici de l'usuari ("JUDICI: ...", decisions
  de lectura) es llisten a part.

  Als informes de tipus 'mns' i 'actextr' NO es compara l'ignorat: l'estat el
  decideix el tipus (_InformeQueDeterminaEstat) i el programa ja no desa cap
  ignorat automatic, o sigui que l'ignorat de la classificacio no te res amb
  que comparar-se. Nomes es compara la conclusio breu, i l'estat al bloc
  d'activitats (on tampoc no s'hi aplica).

  Una diferencia d'ESTAT que nomes ve d'un informe JUDICI (amb la resta de
  valors bons, pero el del JUDICI com el diu el programa, l'estat surt igual
  que ara) tampoc no compta: es llista a part.

  Escriu la LLISTA de discrepancies (no nomes el recompte) a la consola i a
  local\base-dades-activitats\validacio-classificacio_<data>.txt (dins de
  local\: mai es puja).

  Sense -Classificacio fa servir la classificacio-informes_*.json MES RECENT
  de local\base-dades-activitats.

  L'agrupament per activitat es el de l'escaneig, o sigui que fa servir tambe
  l'ID GIA ASSIGNAT A MA: el gia-assignats_*.json MES RECENT de la mateixa
  carpeta (o el de -GiaAssignats). La classificacio de referencia no canvia (es
  per informe). Les assignacions que ja no troben l'informe surten com a AVIS.

  Us (des de l'arrel del repositori):
    powershell -NoProfile -ExecutionPolicy Bypass -File suport\ValidarClassificacio.ps1
    ... -Classificacio <fitxer.json> -Informes <carpeta d'informes> -GiaAssignats <fitxer.json>
#>
param(
    [string]$Classificacio = '',
    [string]$Informes = '',
    [string]$GiaAssignats = ''
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

# Sense -Classificacio: la MES RECENT de local\base-dades-activitats
# (classificacio-informes_*.json), no un nom fix. local\ no es puja mai: si la
# classificacio es va fer a l'altre PC, nomes es alla i s'hi ha de copiar a ma.
if ([string]::IsNullOrWhiteSpace($Classificacio)) {
    $vcCands = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'classificacio-informes_*.json' -File -ErrorAction SilentlyContinue |
                 Sort-Object Name -Descending)
    if ($vcCands.Count -eq 0) {
        throw ("No hi ha cap classificacio (classificacio-informes_*.json) a:`n  " + $LocalActivitatsDir +
               "`nLa carpeta local\ no es puja al GitHub: si la vas fer a l'altre PC, copia-la aqui, o digues on es amb -Classificacio <fitxer.json>.")
    }
    $Classificacio = $vcCands[0].FullName
}
if (-not [string]::IsNullOrWhiteSpace($Informes)) { $InformesDir = $Informes }
# L'ID GIA assignat a ma: el MES RECENT de local\base-dades-activitats. Es busca
# ARA, abans que l'escaneig passi a la carpeta temporal (on no n'hi ha cap).
if ([string]::IsNullOrWhiteSpace($GiaAssignats)) {
    $vcGa = Find-GiaAssignats
    if ($null -ne $vcGa) { $GiaAssignats = $vcGa.FullName }
}
if (-not (Test-Path -LiteralPath $Classificacio)) { throw "No trobo la classificacio: $Classificacio" }
if (-not (_InformesDirAccessible $InformesDir)) { throw "No trobo la carpeta d'informes: $InformesDir" }

$vcSortida = Join-Path $LocalActivitatsDir ('validacio-classificacio_' + (Get-Date).ToString('yyyyMMdd-HHmm') + '.txt')
$vcLinies = New-Object System.Collections.ArrayList
$vcDiu = { param($t) Write-Host $t; [void]$vcLinies.Add([string]$t) }

# ---- 1. L'escaneig de debo, en una carpeta temporal i sense correccions ----
# L'Excel d'activitats, el mateix que faria servir el boto (pot ser el de
# local\, que la carpeta temporal no te): se li dona com a carpeta "principal".
$vcExcel = Find-LatestActivitatsExcel
$vcTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('validar-classificacio-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $vcTmp -Force)
$vcVells = @{ Loc = $LocalActivitatsDir; Act = $ActivitatsDir }
try {
    if ($null -ne $vcExcel) { $ActivitatsDir = Split-Path -Parent $vcExcel.File.FullName }
    $LocalActivitatsDir = $vcTmp
    $vcRes = Invoke-InformesDbEscaneig {
        param($t, $i, $n)
        if ($n -gt 0 -and ($i % 50) -eq 0) { Write-Host ("  ... $i de $n") }
    } $null $GiaAssignats $false
    if (-not [bool]$vcRes.Ok) { throw ("L'escaneig no s'ha pogut fer: " + [string]$vcRes.Error) }
    # Per Get-InformesDbPath, que llegeix $LocalActivitatsDir: aqui dins encara
    # apunta a $vcTmp, o sigui que es EXACTAMENT el fitxer que acaba d'escriure
    # l'escaneig. Escrivint el nom a ma, el dia que canviés deixariem de llegir
    # el que hem generat i la comparacio quedaria muda.
    $vcDb = Read-JsonFile (Get-InformesDbPath)
} finally {
    $LocalActivitatsDir = $vcVells.Loc; $ActivitatsDir = $vcVells.Act
    Remove-Item -LiteralPath $vcTmp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---- 2. La classificacio correcta, per ruta relativa ------------------------
$vcCls = Read-JsonFile $Classificacio
$vcEntrades = @()
if ($vcCls -is [System.Collections.IList]) { $vcEntrades = @($vcCls) }
else {
    foreach ($p in $vcCls.PSObject.Properties) {
        $v = @($p.Value)
        if ($v.Count -gt 0 -and $null -ne $v[0] -and $null -ne $v[0].PSObject.Properties['ruta_relativa']) { $vcEntrades = $v; break }
    }
}
if ($vcEntrades.Count -eq 0) { throw "La classificacio no porta cap entrada amb 'ruta_relativa'." }
$vcBo = @{}
foreach ($e in $vcEntrades) { $vcBo[((_ClauInforme ([string]$e.ruta_relativa) '').TrimStart('\'))] = $e }

$vcArrel = [string]$vcDb.carpeta_arrel
$vcEsDubte  = { param($e) ([string](_PropInf $e 'nota')).Trim() -match '^(?i)dubte' }
$vcEsJudici = { param($e) ([string](_PropInf $e 'nota')) -match '(?i)judici' }
# 'mns' i 'actextr': l'estat el decideix el tipus i l'ignorat no es compara (ni
# s'aplica a l'estat). El tipus BO, i si la classificacio no en porta, el d'ara.
$vcSenseIgnorat = { param($e, $inf)
    $t = [string](_PropInf $e 'tipus')
    if ([string]::IsNullOrWhiteSpace($t)) { $t = [string](_PropInf $inf 'tipus') }
    return ($t -eq 'mns' -or $t -eq 'actextr')
}

# ---- 3. Informe a informe ---------------------------------------------------
$vcDifInf = New-Object System.Collections.ArrayList
$vcJudici = New-Object System.Collections.ArrayList
$vcVistos = @{}
$nDubte = 0; $nComparats = 0
foreach ($act in @($vcDb.activitats)) {
    foreach ($inf in @($act.informes)) {
        $clau = _ClauInforme ([string]$inf.ruta) $vcArrel
        if (-not $vcBo.ContainsKey($clau)) { continue }
        $vcVistos[$clau] = $true
        $e = $vcBo[$clau]
        if (& $vcEsDubte $e) { $nDubte++; continue }
        $nComparats++
        $okBreu = ([string]$e.conclusio_breu -eq [string]$inf.conclusio_breu)
        $okIgn = (& $vcSenseIgnorat $e $inf) -or ([bool]$e.ignorat -eq [bool]$inf.ignorat)
        if ($okBreu -and $okIgn) { continue }
        $vcIgnTxt = -not (& $vcSenseIgnorat $e $inf)
        $t = ('  ' + $clau + "`n" +
              '      correcte: ' + [string]$e.conclusio_breu + $(if ($vcIgnTxt -and [bool]$e.ignorat) { ' (ignorat)' } else { '' }) + '  tipus=' + [string](_PropInf $e 'tipus') + "`n" +
              '      ara:      ' + [string]$inf.conclusio_breu + $(if ($vcIgnTxt -and [bool]$inf.ignorat) { ' (ignorat)' } else { '' }) + '  tipus=' + [string]$inf.tipus +
              $(if ([string]$inf.motiu) { '  motiu=' + [string]$inf.motiu } else { '' }))
        if ([string](_PropInf $e 'nota')) { $t += "`n      nota:     " + [string]$e.nota }
        if (& $vcEsJudici $e) { [void]$vcJudici.Add($t) } else { [void]$vcDifInf.Add($t) }
    }
}
$vcNoTrobats = @($vcBo.Keys | Where-Object { -not $vcVistos.ContainsKey($_) -and -not (& $vcEsDubte $vcBo[$_]) } | Sort-Object)

# ---- 4. Activitat a activitat: l'ESTAT --------------------------------------
# El de debo amb els valors bons (tipus inclos) i la mateixa regla. Els DUBTE es
# queden amb el que diu el programa (no se sap quin es el bo). Als 'mns' i
# 'actextr' no s'hi aplica l'ignorat bo (vegeu $vcSenseIgnorat).
# $ambJudici = $false: els JUDICI tambe es queden amb el que diu el programa.
# Si aixi l'estat surt com ara, la diferencia nomes ve d'un JUDICI i no compta.
$vcActivitatBona = { param($act, [bool]$ambJudici)
    $bons = foreach ($inf in @($act.informes)) {
        $o = $inf.PSObject.Copy()
        $clau = _ClauInforme ([string]$inf.ruta) $vcArrel
        if ($vcBo.ContainsKey($clau) -and -not (& $vcEsDubte $vcBo[$clau]) -and ($ambJudici -or -not (& $vcEsJudici $vcBo[$clau]))) {
            $e = $vcBo[$clau]
            $o.conclusio_breu = [string]$e.conclusio_breu
            if (-not (& $vcSenseIgnorat $e $inf)) { $o.ignorat = [bool]$e.ignorat }
            $o.tipus = [string](_PropInf $e 'tipus')
        }
        $o
    }
    [pscustomobject]@{ id_gia = $act.id_gia; informes = @($bons) }
}
$vcDifEstat = New-Object System.Collections.ArrayList
$vcDifEstatJudici = New-Object System.Collections.ArrayList
foreach ($act in @($vcDb.activitats)) {
    $actBo = & $vcActivitatBona $act $true
    $estatBo = _EstatActualActivitat $actBo
    if ($estatBo -eq [string]$act.estat_actual) { continue }
    $nom = if ([string]$act.id_gia) { 'GIA ' + [string]$act.id_gia } else { 'carpeta ' + [string]$act.carpeta }
    $infBo = _InformeQueDeterminaEstat $actBo
    $infAra = _InformeQueDeterminaEstat $act
    $t = ('  ' + $nom + ':  correcte=' + $estatBo + $(if ($infBo) { ' (' + [string]$infBo.fitxer + ')' } else { '' }) +
          '  ara=' + [string]$act.estat_actual + $(if ($infAra) { ' (' + [string]$infAra.fitxer + ')' } else { '' }))
    $estatSenseJudici = _EstatActualActivitat (& $vcActivitatBona $act $false)
    if ($estatSenseJudici -eq [string]$act.estat_actual) { [void]$vcDifEstatJudici.Add($t) } else { [void]$vcDifEstat.Add($t) }
}

# ---- 5. El resum, i la LLISTA ----------------------------------------------
& $vcDiu ('Validacio del classificador ' + $Script:ClassificadorVersio + '  -  ' + (Get-Date).ToString('dd/MM/yyyy HH:mm'))
& $vcDiu ('Carpeta d''informes: ' + $vcArrel)
& $vcDiu ('Classificacio:       ' + $Classificacio)
& $vcDiu ('GIA assignats:       ' + $(if ([string]$vcRes.GiaAssignatsFitxer) { [string]$vcRes.GiaAssignatsFitxer } else { '(cap fitxer gia-assignats_*.json)' }))
& $vcDiu ''
& $vcDiu ('Informes escanejats: ' + $vcRes.NInformes + '   activitats: ' + $vcRes.NActivitats + '   a revisar: ' + $vcRes.NRevisar)
& $vcDiu ('Entrades de la classificacio: ' + $vcEntrades.Count + '   comparades: ' + $nComparats + '   DUBTE (fora): ' + $nDubte + '   no trobades a la carpeta: ' + $vcNoTrobats.Count)
& $vcDiu ('Discrepancies d''informe: ' + $vcDifInf.Count + '   (judici de l''usuari, a part: ' + $vcJudici.Count + ')')
& $vcDiu ('Discrepancies d''ESTAT d''activitat: ' + $vcDifEstat.Count + '   (nomes per un informe JUDICI, a part: ' + $vcDifEstatJudici.Count + ')')
& $vcDiu ''
$vcGaNo = @($vcRes.GiaAssignatsNoTrobats)
if ([string]$vcRes.GiaAssignatsError -or $vcGaNo.Count -gt 0) {
    & $vcDiu '== AVIS: assignacions de GIA =='
    if ([string]$vcRes.GiaAssignatsError) { & $vcDiu ('  ' + [string]$vcRes.GiaAssignatsError) }
    if ($vcGaNo.Count -gt 0) {
        & $vcDiu ('  ' + $vcGaNo.Count + ' entrades ja no troben l''informe a la carpeta:')
        foreach ($k in $vcGaNo) { & $vcDiu ('    ' + $k) }
    }
    & $vcDiu ''
}
& $vcDiu '== ESTAT D''ACTIVITAT diferent =='
foreach ($t in $vcDifEstat) { & $vcDiu $t }
& $vcDiu ''
& $vcDiu '== INFORMES amb conclusio breu o ignorat diferents (mns i actextr: nomes la conclusio breu) =='
foreach ($t in $vcDifInf) { & $vcDiu $t }
& $vcDiu ''
& $vcDiu '== Judici de l''usuari (no compten) =='
foreach ($t in $vcJudici) { & $vcDiu $t }
& $vcDiu ''
& $vcDiu '== ESTAT diferent NOMES per un informe JUDICI (no compten) =='
foreach ($t in $vcDifEstatJudici) { & $vcDiu $t }
if ($vcNoTrobats.Count -gt 0) {
    & $vcDiu ''
    & $vcDiu '== A la classificacio pero no a la carpeta =='
    foreach ($k in $vcNoTrobats) { & $vcDiu ('  ' + $k) }
}
[void](New-Item -ItemType Directory -Path (Split-Path -Parent $vcSortida) -Force)
[System.IO.File]::WriteAllText($vcSortida, ($vcLinies -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
Write-Host ''
Write-Host ('Desat a: ' + $vcSortida)
