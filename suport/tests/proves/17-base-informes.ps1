# "Actualitzar base": els defectes de l'auditoria d'octubre de 2026. Cada un es
# va MESURAR abans d'arreglar-lo (una passada de debo amb .docx inventats) i
# cada prova es va validar tornant a posar el defecte i veient-la en vermell.
#   1. Una carpeta que no surt en un llistat (tall de xarxa) treia de la base
#      els seus informes i les correccions a ma, per sempre; una carpeta
#      d'informes buida deixava la base amb 0 informes.
#   2. "ID GIA:" sense valor donava el GIA ":" (i ajuntava activitats).
#   3. Un .docx que no s'obria es classificava en blanc i ja no es tornava a llegir.
#   4. Un informe desat DURANT la passada no es tornava a llegir mai.
#   5. Sense l'Excel d'activitats no es deia res (i els contactes, 0 avisos).
#   6. Dues regles per al GIA del nom de la carpeta ("ESTRATEGIA 2025" = GIA 2025).
#   7. El GIA per l'expedient i el titular d'un informe reaprofitat no s'actualitzaven.
#   8. L'expedient i el titular de l'activitat eren els del primer informe del disc.
#   9. El text desat dels .doc esborrats no es treia mai.
# Fa servir $nouDocx (04-correu.ps1) i $xmlCt (11-contactes.ps1): les arees
# comparteixen ambit (run-tests.ps1).

Write-Host "`n--- Base d'informes: l'ID GIA (punts 2 i 6) ---"
AssertEq (_ExtractIdGia @('ID GIA:', '361')) '' '_ExtractIdGia: "ID GIA:" sense valor -> buit (abans ":")'
AssertEq (_ExtractIdGia @('ID GIA :')) '' '_ExtractIdGia: "ID GIA :" -> buit'
AssertEq (_ExtractIdGia @('ID GIA: 361.')) '361' '_ExtractIdGia: el punt del final no es de l''ID'
AssertEq (_ExtractIdGia @('ID GIA: 0361')) '361' '_ExtractIdGia: sense zeros al davant'
AssertEq (_ExtractIdGia @('ID GIA: 361 - BAR')) '361' '_ExtractIdGia: nomes el primer token'
AssertEq (_ExtractIdGia @('ID GIA:', 'ID GIA: 77')) '77' '_ExtractIdGia: una linia sense valor no amaga la seguent'
AssertEq (_NormalitzaGia ' 0042 ') '42' '_NormalitzaGia: espais i zeros'
AssertEq (_NormalitzaGia '0') '0' '_NormalitzaGia: el zero sol es queda'
AssertEq (_NormalitzaGia 'A12/') 'A12' '_NormalitzaGia: un ID no numeric nomes perd la puntuacio del final'
AssertEq (_GiaDeNomCarpeta '2025-1-2563 GIA 361 - BAR') '361' '_GiaDeNomCarpeta: "GIA 361"'
AssertEq (_GiaDeNomCarpeta 'GIA361 BAR') '361' '_GiaDeNomCarpeta: enganxat'
AssertEq (_GiaDeNomCarpeta 'GIA 0361 BAR') '361' '_GiaDeNomCarpeta: sense zeros'
AssertEq (_GiaDeNomCarpeta '2025-1-2563 ESTRATEGIA 2025 SL') '' '_GiaDeNomCarpeta: "ESTRATEGIA 2025" no es el GIA 2025'
AssertEq (_GiaDeNomCarpeta ('ESTRAT' + [char]0x00C8 + 'GIA 2025')) '' '_GiaDeNomCarpeta: ni amb accent ("ESTRATÈGIA")'
AssertEq (_GiaDeNomCarpeta 'ENERGIA 3 SL') '' '_GiaDeNomCarpeta: "ENERGIA 3" no es el GIA 3'
AssertEq (_GiaFromFolderName 'I:\Inf\ENERGIA 3 SL\2025-03-01 REQ1.docx') '' '_GiaFromFolderName: fa servir la mateixa regla'
AssertEq (_GiaFromFolderName 'I:\Inf\GIA 0361 BAR\2025-03-01 REQ1.docx') '361' '_GiaFromFolderName: ...normalitzada'
# Una sola regla: cap altre fitxer no busca el GIA al nom d'una carpeta pel seu compte.
$srcCtGia = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'ContactesDb.ps1')
Assert ($srcCtGia.Contains('_GiaDeNomCarpeta') -and -not ($srcCtGia -match 'gia\\s\*\(\\d')) 'contactes: el GIA del nom de la carpeta, amb la funcio de la base (cap regex propia)'
$srcInfGia = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Informes.ps1')
Assert (-not $srcInfGia.Contains("-match 'GIA\s*(\d+)'")) 'base: _GiaFromFolderName ja no te la regex vella'

Write-Host "`n--- Base d'informes: la pregunta dels informes que desapareixen (punt 1) ---"
AssertEq (_AvisInformesDesapareguts 0 0 0) '' 'sense base, res a preguntar'
AssertEq (_AvisInformesDesapareguts 100 2 0) '' 'en falten 2 de 100 i cap correccio: no es pregunta'
AssertEq (_AvisInformesDesapareguts 30 4 0) '' 'en falten 4 (menys de 5): no es pregunta'
Assert ((_AvisInformesDesapareguts 100 10 0) -ne '') 'en falten 10 de 100 (el 10%): es pregunta'
Assert ((_AvisInformesDesapareguts 800 6 0) -eq '') 'en falten 6 de 800 (menys del 10%): no es pregunta'
$avD = _AvisInformesDesapareguts 800 1 1 @('GIA 9\2026-01-01_Req.docx')
Assert ($avD.Contains("perdrien 1 correccions") -and $avD.Contains('GIA 9\2026-01-01_Req.docx') -and $avD.EndsWith("Vols continuar igualment?")) 'una sola correccio a ma que es perdria: es pregunta, i es diu quin informe'
$avM = _AvisInformesDesapareguts 100 12 0 @(1..12 | ForEach-Object { "C\$_.docx" })
Assert ($avM.Contains('... i 7 m') -and -not $avM.Contains('C\6.docx')) 'en llista 5 i diu quants mes'

$bi = Join-Path ([System.IO.Path]::GetTempPath()) ('base-audit-' + [guid]::NewGuid().ToString('N'))
$biI = Join-Path $bi 'Informes'; $biLoc = Join-Path $bi 'local'
$vellsBI = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
$origFindXl = ${function:Find-LatestActivitatsExcel}; $origInitXl = ${function:Initialize-ActivitatsCache}
$falsXl = @{ Cache = $null; Error = '' }
try {
    foreach ($d in @('GIA 901', 'GIA 902', 'SENSE GIA C', 'SENSE GIA D', 'ESTRATEGIA 2025 SL', 'SENSE GIA E')) { [void](New-Item -ItemType Directory -Path (Join-Path $biI $d) -Force) }
    foreach ($d in @($biLoc, (Join-Path $bi 'act'), (Join-Path $bi 'app'))) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $pBI = { param($c, $f) Join-Path (Join-Path $biI $c) $f }
    $InformesDir = $biI; $LocalActivitatsDir = $biLoc; $ActivitatsDir = Join-Path $bi 'act'; $env:LOCALAPPDATA = Join-Path $bi 'app'
    $reqB = "Vist l'anterior, s'inicia d'ofici el procediment d'esmena, disposant d'un termini d'un mes."
    $fiB  = "Vist l'anterior s'informa que es pot donar per finalitzat el procediment d'esmena."
    $tanc = 'Ho poso al seu coneixement als efectes oportuns,'
    & $nouDocx (& $pBI 'GIA 901' '2026-01-10_Req.docx') @('ID GIA: 901', 'Exp. Num: 2026/1/2563', $reqB, $tanc)
    & $nouDocx (& $pBI 'GIA 901' '2026-02-10_Seg.docx') @('ID GIA: 901', 'Exp. Num: 2026/2/2579', $fiB, $tanc)
    & $nouDocx (& $pBI 'GIA 902' '2026-03-01_Req.docx') @('ID GIA: 902', $reqB, $tanc)
    & $nouDocx (& $pBI 'SENSE GIA C' '2026-03-02_Req.docx') @('ID GIA:', 'Exp. Num: 2026/9/2563', $reqB, $tanc)
    & $nouDocx (& $pBI 'SENSE GIA D' '2026-03-03_Req.docx') @('ID GIA:', 'Exp. Num: 2026/8/2563', $reqB, $tanc)
    & $nouDocx (& $pBI 'ESTRATEGIA 2025 SL' '2026-03-04_Req.docx') @('Exp. Num: 2026/7/2563', $reqB, $tanc)
    & $nouDocx (& $pBI 'SENSE GIA E' '2026-03-05_Req.docx') @('ID GIA:', 'Exp. Num: 2026/5/2563', $reqB, $tanc)
    $vellBI = (Get-Date).ToUniversalTime().AddHours(-2)
    Get-ChildItem -LiteralPath $biI -Recurse -File | ForEach-Object { $_.LastWriteTimeUtc = $vellBI }
    $dbBI = Join-Path $biLoc 'informes-db.json'
    $actBI = { param($gia) @((Read-JsonFile $dbBI).activitats | Where-Object { [string]$_.id_gia -eq $gia }) }
    $infBI = { param($fitxer) foreach ($a in @((Read-JsonFile $dbBI).activitats)) { foreach ($x in @($a.informes)) { if ([string]$x.fitxer -eq $fitxer) { return $x } } } }

    # L'Excel d'activitats, de mentida (Find + Initialize): $falsXl.Cache, o llanca $falsXl.Error.
    # Assignades i NO amb 'function': una definicio 'function' dins d'un script
    # existeix des del principi del fitxer (mesurat: el Get-ChildItem de mentida
    # de mes avall ja responia a la primera passada), i el $orig* de dalt hauria
    # desat la de mentida.
    ${function:Find-LatestActivitatsExcel} = { if ($null -eq $falsXl.Cache -and $falsXl.Error -eq '') { return $null }; return [pscustomobject]@{ File = 'ExportacioActivitats_20261027.xls' } }
    ${function:Initialize-ActivitatsCache} = { param($f) if ($falsXl.Error -ne '') { throw $falsXl.Error }; return $falsXl.Cache }

    Write-Host "`n--- Base d'informes: una passada de debo (punts 2, 5, 6 i 8) ---"
    $r1 = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($r1.Ok)|$($r1.NInformes)" 'True|7' 'passada: llegeix els 7 informes'
    Assert ((@((Read-JsonFile $dbBI).activitats | ForEach-Object { [string]$_.id_gia }) -notcontains ':')) 'punt 2: cap activitat amb el GIA ":"'
    AssertEq "$(@(& $actBI '').Count)" '4' 'punt 2: els informes sense GIA (C, D, E i ESTRATEGIA) son quatre activitats, una per carpeta'
    AssertEq "$(@(& $actBI '2025').Count)" '0' 'punt 6: "ESTRATEGIA 2025 SL" no fa el GIA 2025'
    AssertEq ([string]$r1.ExcelAvis) "no s'ha trobat l'Excel d'activitats" 'punt 5: sense l''Excel, la passada ho diu'
    AssertEq ([string]@(& $actBI '901')[0].expedient) '2026/2/2579' 'punt 8: sense Excel, l''expedient de l''activitat es el de l''informe MES RECENT'

    Write-Host "`n--- Base d'informes: un informe desat durant la passada (punt 4) ---"
    # Canvia el text i la data queda ENTRE la de la lectura i la de la base
    # (mes vella que actualitzat_el): abans no es tornava a llegir mai.
    $f902 = & $pBI 'GIA 902' '2026-03-01_Req.docx'
    & $nouDocx $f902 @('ID GIA: 902', $fiB, $tanc)
    (Get-Item -LiteralPath $f902).LastWriteTimeUtc = $vellBI.AddMinutes(30)
    $r2 = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($r2.Reprocessats)|$((& $infBI '2026-03-01_Req.docx').conclusio_breu)" '1|FI Requeriment' 'punt 4: l''informe amb una data que no es la desada es torna a llegir (encara que sigui mes vell que la base)'
    $r2b = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq ([int]$r2b.Reprocessats) 0 'punt 4: ...i la passada seguent ja no (la data desada es la seva)'

    Write-Host "`n--- Base d'informes: un .docx que no s'obre (punt 3) ---"
    [System.IO.File]::WriteAllText($f902, 'no soc un zip')
    (Get-Item -LiteralPath $f902).LastWriteTimeUtc = $vellBI.AddMinutes(40)
    $r3 = Invoke-InformesDbEscaneig $null $null '' $false
    $i902 = & $infBI '2026-03-01_Req.docx'
    AssertEq "$($r3.NoLlegits)|$($i902.no_llegit)|$($i902.conclusio_breu)|$(@(& $actBI '902').Count)" '1|True|FI Requeriment|1' 'punt 3: es queda el que se''n sabia (conclusio i GIA) i queda pendent'
    & $nouDocx $f902 @('ID GIA: 902', $reqB, $tanc)
    (Get-Item -LiteralPath $f902).LastWriteTimeUtc = $vellBI.AddMinutes(40)
    $r3b = Invoke-InformesDbEscaneig $null $null '' $false
    $i902b = & $infBI '2026-03-01_Req.docx'
    AssertEq "$($r3b.NoLlegits)|$([bool]$i902b.no_llegit)|$($i902b.conclusio_breu)" '0|False|Requeriment' 'punt 3: la passada seguent el torna a provar (encara que la data no hagi canviat) i el llegeix'
    $fNou = & $pBI 'GIA 902' '2026-04-01_Nou.docx'
    [System.IO.File]::WriteAllText($fNou, 'a mig copiar')
    $r3c = Invoke-InformesDbEscaneig $null $null '' $false
    $iNou = & $infBI '2026-04-01_Nou.docx'
    Assert ([bool]$iNou.no_llegit -and ([string]$iNou.motiu).Contains("no s'ha pogut obrir")) 'punt 3: un .docx NOU que no s''obre surt a revisar dient per que, i pendent'
    Remove-Item -LiteralPath $fNou -Force

    Write-Host "`n--- Base d'informes: l'Excel i els informes reaprofitats (punts 5, 7 i 8) ---"
    $falsXl.Cache = [pscustomobject]@{ ById = @{ '905' = @{ TITULAR = 'TITULAR NOU SL'; EXP_NUM = '2026/5/2563' }; '901' = @{ TITULAR = 'BAR 901 SL'; EXP_NUM = '2026/1/2563' } }; Warnings = @() }
    $r5 = Invoke-InformesDbEscaneig $null $null '' $false
    $iE = & $infBI '2026-03-05_Req.docx'
    AssertEq "$($r5.Reprocessats)|$($iE.gia_font)|$(@(& $actBI '905').Count)" '0|excel|1' 'punt 7: un informe reaprofitat sense GIA el troba a l''Excel per l''expedient (sense tornar-lo a llegir)'
    Assert (-not ([string]$iE.motiu).Contains('sense ID GIA')) 'punt 7: ...i ja no surt "sense ID GIA"'
    AssertEq "$(@(& $actBI '905')[0].titular)|$(@(& $actBI '901')[0].titular)|$(@(& $actBI '901')[0].expedient)" 'TITULAR NOU SL|BAR 901 SL|2026/1/2563' 'punt 8: el titular i l''expedient de l''activitat, els de l''Excel'
    AssertEq ([string]$r5.ExcelAvis) '' 'punt 5: amb l''Excel, cap avis'
    $falsXl.Cache.ById['905'].TITULAR = 'TITULAR CANVIAT SL'
    $r5b = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($r5b.Reprocessats)|$(@(& $actBI '905')[0].titular)" '0|TITULAR CANVIAT SL' 'punt 7: un canvi de titular a l''Excel arriba a la base sense tornar a llegir l''informe'
    $falsXl.Error = 'Excel no respon'
    $r5c = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($r5c.Ok)|$($r5c.ExcelAvis)" "True|no s'ha pogut llegir l'Excel d'activitats (Excel no respon)" 'punt 5: si l''Excel peta, la passada es fa i ho diu'
    AssertEq ([string](& $infBI '2026-03-05_Req.docx').gia_font) 'excel' 'punt 7: sense l''Excel, el GIA que en venia es queda (no es buida)'
    $falsXl.Error = ''; $falsXl.Cache = $null

    Write-Host "`n--- Base d'informes: el text desat dels .doc que ja no hi son (punt 9) ---"
    Write-JsonFile (_DocTextCachePath) ([ordered]@{ versio = 1; documents = @{ 'GIA 901\2019-01-01_Esborrat.doc' = [ordered]@{ modificat = '2019-01-01T00:00:00.0000000Z'; linies = @('x') } } }) 5
    [void](Invoke-InformesDbEscaneig $null $null '' $false)
    AssertEq (@((_LlegeixDocTextCache).Map.Keys).Count) 0 'punt 9: el text d''un .doc que ja no hi es se''n treu'

    Write-Host "`n--- Base d'informes: els informes que desapareixen (punt 1) ---"
    $dbAra = Read-JsonFile $dbBI
    $i901 = @(@($dbAra.activitats | Where-Object { [string]$_.id_gia -eq '901' })[0].informes)[0]
    _MarcaEditatAMa $i901; $i901.conclusio_breu = 'Precinte / Cessament'
    Write-JsonFile $dbBI $dbAra 8
    $nAbans = [int]$dbAra.n_informes
    $amagada = Join-Path $bi 'amagada'
    Move-Item -LiteralPath (Join-Path $biI 'GIA 901') -Destination $amagada
    $r1a = Invoke-InformesDbEscaneig
    AssertEq "$($r1a.Ok)|$($r1a.Cancelat)|$([int](Read-JsonFile $dbBI).n_informes)" "False|True|$nAbans" 'punt 1: una carpeta que no surt (amb una correccio a ma): l''automatic no escriu'
    $pregBI = @{ Text = '' }
    $r1b = Invoke-InformesDbEscaneig $null { param($p) $pregBI.Text = $p; $false } '' $false
    Assert ($pregBI.Text.Contains('perdrien 1 correccions') -and $pregBI.Text.Contains('GIA 901\2026-01-10_Req.docx')) 'punt 1: el boto pregunta, i diu quin informe i quantes correccions'
    Move-Item -LiteralPath $amagada -Destination (Join-Path $biI 'GIA 901')
    $r1c = Invoke-InformesDbEscaneig $null $null '' $false
    $i901b = & $infBI '2026-01-10_Req.docx'
    AssertEq "$($r1c.Ok)|$($i901b.conclusio_breu)|$($i901b.editat_a_ma)" 'True|Precinte / Cessament|True' 'punt 1: quan la carpeta torna, la correccio a ma hi es'
    # La carpeta d'informes buida: ni el boto ni l'automatic escriuen.
    Move-Item -LiteralPath $biI -Destination (Join-Path $bi 'Inf-bak'); [void](New-Item -ItemType Directory -Path $biI)
    $r1d = Invoke-InformesDbEscaneig $null { param($p) $true } '' $false
    AssertEq "$($r1d.Ok)|$([int](Read-JsonFile $dbBI).n_informes)" "False|$nAbans" 'punt 1: la carpeta d''informes surt buida -> no s''escriu (ni dient que si)'
    Assert ([string]$r1d.Error).Contains("No s'ha tocat la base") 'punt 1: ...i es diu'
    Remove-Item -LiteralPath $biI -Force; Move-Item -LiteralPath (Join-Path $bi 'Inf-bak') -Destination $biI
    # Una carpeta que la xarxa no deixa llistar (un Get-ChildItem de mentida
    # que hi falla): ni preguntar, s'atura.
    $fallaLl = @{ Ruta = (Join-Path $biI 'GIA 902'); Parcial = $false }
    ${function:Get-ChildItem} = {
        [CmdletBinding()] param([string]$LiteralPath, [switch]$Directory, [switch]$File, [switch]$Recurse, [string]$Filter)
        if ($LiteralPath -eq $fallaLl.Ruta -and -not $fallaLl.Parcial) { Write-Error ("El nombre de red ya no est" + [char]0x00E1 + " disponible: " + $LiteralPath); return }
        if ($LiteralPath -eq $fallaLl.Ruta) { Write-Error ("The specified path, file name, or both are too long: " + $LiteralPath) }
        $pp = @{ LiteralPath = $LiteralPath }
        if ($Directory) { $pp.Directory = $true }; if ($File) { $pp.File = $true }; if ($Recurse) { $pp.Recurse = $true }; if ($Filter) { $pp.Filter = $Filter }
        Microsoft.PowerShell.Management\Get-ChildItem @pp -ErrorAction SilentlyContinue
    }
    try {
        $errLl = New-Object System.Collections.ArrayList
        $llPN = Get-FitxersPrimerNivell $biI $errLl
        AssertEq "$($errLl.Count)|$(@($errLl)[0].Carpeta)" '1|GIA 902' 'Get-FitxersPrimerNivell: apunta la carpeta que no s''ha pogut llistar'
        $r1e = Invoke-InformesDbEscaneig $null { param($p) $true } '' $false
        AssertEq "$($r1e.Ok)|$([int](Read-JsonFile $dbBI).n_informes)" "False|$nAbans" 'punt 1: una carpeta de la base que no es pot llistar -> no s''escriu'
        Assert ([string]$r1e.Error).Contains('GIA 902') 'punt 1: ...i diu quina'
        # Un error d'UN fitxer (la ruta massa llarga del PowerShell 5.1) i la
        # resta de la carpeta llistada: no falta cap informe, no ho atura.
        $fallaLl.Parcial = $true
        $r1p = Invoke-InformesDbEscaneig $null $null '' $false
        AssertEq "$($r1p.Ok)|$($r1p.NInformes)" "True|$nAbans" 'punt 1: un error que no fa perdre cap informe (un fitxer amb la ruta massa llarga) no ho atura'
        $fallaLl.Parcial = $false
        $fallaLl.Ruta = Join-Path $biI 'MAI A LA BASE'
        [void](New-Item -ItemType Directory -Path $fallaLl.Ruta -Force)
        $r1f = Invoke-InformesDbEscaneig $null $null '' $false
        AssertEq ([string]$r1f.Ok) 'True' 'punt 1: una carpeta que no ha tingut mai informes a la base no ho atura'
    } finally {
        Remove-Item -LiteralPath function:\Get-ChildItem -Force -ErrorAction SilentlyContinue
    }

    Write-Host "`n--- Contactes sense l'Excel: la fila de la passada anterior (punt 5) ---"
    _CtOblidaMemo
    [System.IO.File]::WriteAllText((& $pBI 'GIA 901' 'tramit_etram-tramit.xml'), $xmlCt, (New-Object System.Text.UTF8Encoding($false)))
    $xlCt901 = @{ TITULAR = 'EXEMPLE INVENTAT SL'; NIF = 'B12345674'; EMAIL = 'pere@enginyeria-inventada.test'; MOBIL = ''; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    $infCt = @([pscustomobject]@{ Ruta = (& $pBI 'GIA 901' '2026-01-10_Req.docx'); Gia = '901' })
    $rCa = Invoke-ContactesEscaneig $biI $infCt ([pscustomobject]@{ ById = @{}; Contactes = @{ '901' = $xlCt901 } }) @{}
    $nAvisos = @((Read-ContactesDb).activitats.'901'.avisos).Count
    Assert ($nAvisos -gt 0 -and -not [bool]$rCa.ExcelAnterior) 'contactes amb l''Excel: avisos'
    $rCb = Invoke-ContactesEscaneig $biI $infCt $null @{}
    AssertEq "$([bool]$rCb.ExcelAnterior)|$(@((Read-ContactesDb).activitats.'901'.avisos).Count)" "True|$nAvisos" 'punt 5: sense l''Excel, els avisos es fan amb la fila de la passada anterior (abans 0)'
    Assert ([string]$rCb.Text).Contains("sense l'Excel d'activitats") 'punt 5: ...i el resum ho diu'
    AssertEq ([string](Read-ContactesDb).activitats.'901'.excel.NIF) 'B12345674' 'punt 5: ...i la fila de l''Excel es conserva a la base'
} catch {
    Assert $false ('bloc auditoria de la base: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    ${function:Find-LatestActivitatsExcel} = $origFindXl; ${function:Initialize-ActivitatsCache} = $origInitXl
    Remove-Item -LiteralPath function:\Get-ChildItem -Force -ErrorAction SilentlyContinue
    $InformesDir = $vellsBI.Inf; $LocalActivitatsDir = $vellsBI.Loc; $ActivitatsDir = $vellsBI.Act; $env:LOCALAPPDATA = $vellsBI.App
    _CtOblidaMemo
    Remove-Item -LiteralPath $bi -Recurse -Force -ErrorAction SilentlyContinue
}
