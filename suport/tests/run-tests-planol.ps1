# Proves automatiques de les funcions PURES del "Planol activitats"
# (rutes/Planol.ps1 + rutes/PlanolDades.ps1).
#
# NO prova l'Excel (COM), el Cadastre de veritat (xarxa) ni les finestres. Les
# respostes del Cadastre (dades/wfsCP-exemple.xml, dades/dnprc-*.xml) estan
# MUNTADES A MA seguint els esquemes: el host estava bloquejat des d'on es va
# escriure. Per provar-ho de veritat: suport\rutes\Provar-Planol.bat
#
# Execucio: pwsh -File tests/run-tests-planol.ps1

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }
$env:PLANOL_TEST = '1'
. (Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'rutes' 'Planol.ps1'))
. (Join-Path $PSScriptRoot 'TestLib.ps1')
$dades = Join-Path $PSScriptRoot 'dades'

Write-Host "`n--- Referencies cadastrals ---"
AssertEq (Get-PlanolParcela '2295827DF2729E0011RQ') '2295827DF2729E' 'parcel.la: els 14 primers'
AssertEq (Get-PlanolParcela ' 2295827df2729e ') '2295827DF2729E' 'parcel.la: retallada i en majuscules'
AssertEq (Get-PlanolParcela '2295827DF') '' 'massa curta: cap parcel.la'
AssertEq (Get-PlanolParcela '') '' 'buida: cap parcel.la'
AssertEq (Get-PlanolUnitat '2295827DF2729E0011RQ') '0011' 'unitat: caracters 15-18'
AssertEq (Get-PlanolUnitat '2295827DF2729E') '' 'nomes la parcel.la: cap unitat'

Write-Host "`n--- Get-EstatPlanol: els quatre colors ---"
AssertEq (Get-EstatPlanol $true 'Favorable') 'vermell' 'precintada a l Excel mana sobre l informe'
AssertEq (Get-EstatPlanol $false 'Precinte / Cessament') 'vermell' 'darrer informe de precinte -> vermell'
AssertEq (Get-EstatPlanol $false 'Requeriment') 'groc' 'requeriment -> groc'
AssertEq (Get-EstatPlanol $false ('Ampliaci' + [char]0x00F3 + ' termini')) 'groc' 'ampliacio de termini (amb accent) -> groc'
AssertEq (Get-EstatPlanol $false 'Favorable') 'verd' 'favorable -> verd'
AssertEq (Get-EstatPlanol $false 'FI Requeriment') 'verd' 'FI requeriment -> verd'
AssertEq (Get-EstatPlanol $false 'FI Precinte / Cessament') 'verd' 'FI precinte -> verd'
AssertEq (Get-EstatPlanol $false $Script:EstatFavorablePre) 'groc' 'favorable pre-llicencia (es tracta com Requeriment) -> groc'
AssertEq (Get-EstatPlanol $false $Script:EstatFavorablePost) 'verd' 'favorable post-llicencia (es tracta com Favorable) -> verd'
AssertEq (Get-EstatPlanol $false '') 'blau' 'sense informes -> blau'
AssertEq (Get-EstatPlanol $false 'Revisar') 'blau' 'Revisar -> blau'
AssertEq (Get-EstatPlanol $false 'Sense efecte') 'blau' 'Sense efecte -> blau'
AssertEq (Get-PitjorEstatPlanol @('verd', 'groc', 'blau')) 'groc' 'el pitjor: groc > blau > verd'
AssertEq (Get-PitjorEstatPlanol @('verd', 'vermell', 'groc')) 'vermell' 'el pitjor: vermell'
AssertEq (Get-PitjorEstatPlanol @()) '' 'cap estat -> buit'

Write-Host "`n--- Get-SubEstabliment: local, planta, porta ---"
$eBase = @{ Rc = '2295827DF2729E0011RQ'; Local = ''; Bloc = ''; Escala = ''; Pis = ''; Porta = '' }
$e1 = [pscustomobject]($eBase.Clone()); $e1.Local = '5'
AssertEq (Get-SubEstabliment $e1 $null).Text 'Local 5' 'un local numeric porta l etiqueta'
$e2 = [pscustomobject]($eBase.Clone()); $e2.Local = 'NAU 6'
AssertEq (Get-SubEstabliment $e2 $null).Text 'NAU 6' 'un local amb nom, tal qual'
$e2b = [pscustomobject]($eBase.Clone()); $e2b.Bloc = 'C'; $e2b.Pis = 'BXS'
AssertEq (Get-SubEstabliment $e2b $null).Text 'Bl. C - Pl. BXS' 'bloc i planta amb text: amb etiqueta (un C sol no diu res)'
$e3 = [pscustomobject]($eBase.Clone()); $e3.Pis = '2'; $e3.Porta = '1A'
AssertEq (Get-SubEstabliment $e3 $null).Text 'Pl. 2 - Pt. 1A' 'planta i porta'
AssertEq (Get-SubEstabliment $e3 $null).Font 'gia' 'i diu que surt de l Excel'
$e4 = [pscustomobject]($eBase.Clone())
$uCad = ConvertFrom-CatastroDnprcXml ([System.IO.File]::ReadAllText((Join-Path $dades 'dnprc-exemple.xml')))
$s4 = Get-SubEstabliment $e4 $uCad
AssertEq "$($s4.Text)|$($s4.Font)" 'Esc. 1 - Pl. 2 - Pt. 16|cadastre' 'sense res a l Excel: la planta/porta del Cadastre'
$s5 = Get-SubEstabliment $e4 $null
AssertEq "$($s5.Text)|$($s5.Font)" 'unitat 0011|unitat' 'sense res enlloc: el numero d unitat'
$e6 = [pscustomobject]($eBase.Clone()); $e6.Rc = '2295827DF2729E'
AssertEq (Get-SubEstabliment $e6 $null).Text '' 'sense unitat a la refcat: res'
$uBaixa = [pscustomobject]@{ Escala = ''; Planta = '00'; Porta = '01'; Bloc = '' }
AssertEq (Get-SubEstabliment $e4 $uBaixa).Text 'Pl. baixa - Pt. 1' 'planta 00 del Cadastre = baixa'

Write-Host "`n--- Cadastre: geometria de la parcel.la (wfsCP) ---"
$xmlCp = [System.IO.File]::ReadAllText((Join-Path $dades 'wfsCP-exemple.xml'))
$polys = @(ConvertFrom-CatastroParcelXml $xmlCp)
AssertEq $polys.Count 2 'dos poligons (el punt de referencia no compta)'
AssertEq @($polys[0].Anells).Count 2 'el primer: exterior + un forat'
AssertEq @($polys[0].Anells[0]).Count 10 'l exterior: 5 vertexs = 10 numeros (array pla)'
AssertNear ([double]@($polys[0].Anells[0])[0]) 421950.0 0.001 'primer x'
AssertNear ([double]@($polys[0].Anells[0])[1]) 4579480.0 0.001 'primer y'
$xmlGirat = $xmlCp -replace '421950.00 4579480.00 422000.00 4579480.00 422000.00 4579530.00 421950.00 4579530.00 421950.00 4579480.00', '4579480.00 421950.00 4579480.00 422000.00 4579530.00 422000.00 4579530.00 421950.00 4579480.00 421950.00'
$polysG = @(ConvertFrom-CatastroParcelXml $xmlGirat)
AssertEq "$(@($polysG[0].Anells[0])[0])|$(@($polysG[0].Anells[0])[1])" '421950|4579480' 'eixos a l inreves: es giren'
AssertEq @(ConvertFrom-CatastroParcelXml '').Count 0 'resposta buida: cap poligon'
AssertEq @(ConvertFrom-CatastroParcelXml '<no es xml').Count 0 'resposta que no s entén: cap poligon, i no peta'
AssertEq @(ConvertFrom-CatastroParcelXml '<a><b/></a>').Count 0 'XML sense poligons: cap'
# El punt de la parcel.la al Cadastre (referencePoint), per a la linia de punts.
AssertEq (@(Get-PuntReferenciaParcela $xmlCp) -join '|') '421975|4579505' 'el punt de la parcel.la (referencePoint)'
AssertEq (@(Get-PuntReferenciaParcela ($xmlCp -replace '421975.00 4579505.00', '4579505.00 421975.00')) -join '|') '421975|4579505' 'eixos a l inreves: es giren'
AssertEq (Get-PuntReferenciaParcela '<a/>') $null 'sense referencePoint: cap'
AssertEq (Get-PuntReferenciaParcela '<no es xml') $null 'resposta que no s entén: cap, i no peta'
$parcCp = ConvertFrom-CatastroParcela $xmlCp
AssertEq "$(@($parcCp.Poligons).Count)|$(@($parcCp.Punt) -join ',')" '2|421975,4579505' 'la parcel.la sencera: poligons i punt'
AssertEq (ConvertFrom-CatastroParcela '<a/>') $null 'sense res: $null (la memoria cau la tornara a demanar)'
$centre = Get-CentreAnell @(0.0, 0.0, 10.0, 0.0, 10.0, 10.0, 0.0, 10.0, 0.0, 0.0)
AssertEq "$($centre[0])|$($centre[1])" '5|5' 'el centre d un quadrat'
AssertNear (Get-AreaAnell @(0.0, 0.0, 10.0, 0.0, 10.0, 10.0, 0.0, 10.0, 0.0, 0.0)) 100.0 0.0001 'l area d un quadrat de 10'

Write-Host "`n--- Cadastre: la unitat (Consulta_DNPRC) ---"
AssertEq "$($uCad.Escala)|$($uCad.Planta)|$($uCad.Porta)|$($uCad.Us)|$($uCad.Superficie)" '1|02|16|Industrial|250' 'escala, planta, porta, us i superficie'
Assert ($uCad.Text.Contains('CADIS 19')) 'i la descripcio sencera'
AssertEq (ConvertFrom-CatastroDnprcXml ([System.IO.File]::ReadAllText((Join-Path $dades 'dnprc-error.xml')))) $null 'error del Cadastre -> null'
AssertEq (ConvertFrom-CatastroDnprcXml 'res') $null 'resposta illegible -> null'

Write-Host "`n--- Les fulles de l'Excel (matriu feta a ma) ---"
function _Mat($files) {
    $nf = @($files).Count; $nc = @($files[0]).Count
    $m = [Array]::CreateInstance([object], @($nf, $nc), @(1, 1))
    for ($i = 0; $i -lt $nf; $i++) { for ($c = 0; $c -lt $nc; $c++) { $m[($i + 1), ($c + 1)] = $files[$i][$c] } }
    return ,$m
}
# Les capcaleres de debo de l'Excel d'establiments, amb les seves rareses.
$capE = @('ID Establiment GIA', 'Ref. cadastral', 'UTM X', 'UTM Y', 'Emp. Tipus via', 'Emp. Carrer', ('Emp._N' + [char]0x00FA + 'mero_'),
          'Emp. Lletra', 'Emp. Bloc', ('Emp. N' + [char]0x00BA + ' Local'), 'Emp. Escala', 'Emp. Pis', 'Emp. Porta', 'Local buit', 'ID Activitat')
$si = 'S' + [char]0x00ED
$filesE = @(
    $capE,
    @([double]1, '2295827DF2729E0011RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', '', '', '', '', '', '', 'No', [double]1447),
    @([double]2, '2295827DF2729E0008RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', '', '', '5', '', '', '', 'No', [double]1403),
    @([double]3, '2295827DF2729E0003XL', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', '', '', '', '', '', '', $si, $null),
    @([double]4, '4091106DF2749A0006XJ', [double]423912.16, [double]4578928.25, 'CTRA', 'HOSPITALET', '147', '', '', '', '', '', '', $si, [double]9),
    @([double]5, '', [double]422500.0, [double]4579000.0, 'C', 'ENLLOC', '1', '', '', '', '', '', '', 'No', [double]10),
    @([double]6, '', $null, $null, 'C', 'ENLLOC', '2', '', '', '', '', '', '', 'No', [double]11),
    @([double]7, '2295827DF2729E0011RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', '', '', '', '', '', '', 'No', [double]777)
)
$ests = @(ConvertFrom-FullaEstabliments (_Mat $filesE) $filesE.Count $capE)
AssertEq $ests.Count 7 'set establiments'
AssertEq "$($ests[0].IdEst)|$($ests[0].IdActivitat)|$($ests[0].Rc)" '1|1447|2295827DF2729E0011RQ' 'IDs com a text i la refcat sencera'
AssertEq $ests[1].Local '5' "el local, de la columna 'Emp. N<ordinal> Local'"
AssertEq $ests[0].Adreca 'C CADIS 19' "l'adreca, amb el numero de 'Emp._Numero_'"
AssertEq "$($ests[2].Buit)|$($ests[2].IdActivitat)" 'True|' 'local buit sense activitat'
AssertEq $ests[3].Buit $true "'Si' amb accent -> buit"

$capA = @('ID Activitat', 'Ref. cadastral', 'UTM X', 'UTM Y', 'Emp. Tipus via', 'Emp. Carrer', 'Emp. Numero', 'Activitat principal', 'Nom comercial activitat', 'Rao social', 'Camp Info 1 - Nom', 'Camp Info 1 - Valor')
$filesA = @(
    $capA,
    @([double]1447, '2295827DF2729E0011RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', 'BAR', 'EL RACO', 'RACO SL', 'PRECINTE ACTIVITAT?', 'SI, PRECINTAT 01/10/2026'),
    @([double]1403, '2295827DF2729E0008RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', 'TALLER', '', '', '', ''),
    @([double]9, '4091106DF2749A0006XJ', [double]423912.16, [double]4578928.25, 'CTRA', 'HOSPITALET', '147', 'OFICINES', 'ACME', 'ACME INVERSIONS SL', '', ''),
    @([double]10, '', [double]422500.0, [double]4579000.0, 'C', 'ENLLOC', '1', 'BOTIGA', '', '', '', ''),
    @([double]11, '', $null, $null, 'C', 'ENLLOC', '2', 'BOTIGA', '', '', '', ''),
    @([double]2000, '3085213DF2738E0001AB', [double]422800.0, [double]4579200.0, 'PG', 'FERROCARRILS', '177', 'MAGATZEM', '', '', '', '')
)
$acts = ConvertFrom-FullaActivitatsPlanol (_Mat $filesA) $filesA.Count $capA
AssertEq $acts.Count 6 'sis activitats'
AssertEq $acts['1447'].Precinte $true 'precintada pel camp lliure'
AssertEq $acts['1403'].Precinte $false 'la resta, no'
AssertEq $acts['9'].Titular 'ACME INVERSIONS SL' 'el TITULAR (rao social), no el nom comercial (ACME)'

Write-Host "`n--- La base d'informes ---"
$db = [pscustomobject]@{ activitats = @(
    [pscustomobject]@{ id_gia = '1403'; estat_actual = 'Requeriment'; informes = @(1, 2) }
    [pscustomobject]@{ id_gia = '9'; estat_actual = 'Favorable'; informes = @(1) }
    [pscustomobject]@{ id_gia = '-'; estat_actual = 'Requeriment'; informes = @() }
) }
$estats = ConvertFrom-InformesDbPlanol $db
AssertEq "$($estats.Count)|$($estats['1403'].Estat)|$($estats['1403'].NInformes)" '2|Requeriment|2' 'per ID GIA, amb estat i nombre d informes (el "-" no compta)'
AssertEq (ConvertFrom-InformesDbPlanol $null).Count 0 'sense base: res'

Write-Host "`n--- Quines unitats es pregunten al Cadastre ---"
$aConsultar = @(Get-UnitatsAConsultar $ests)
AssertEq ($aConsultar -join ',') '2295827DF2729E0003XL,2295827DF2729E0011RQ' 'les de parcel.les compartides sense local/planta/porta (i cap de repetida)'

Write-Host "`n--- Build-PlanolModel ---"
$unitats = @{ '2295827DF2729E0011RQ' = $uCad }
$model = Build-PlanolModel $ests $acts $estats $unitats
$perClau = @{}; foreach ($p in $model.Parceles) { $perClau[$p.Clau] = $p }
$cadis = $perClau['2295827DF2729E']
AssertEq (@($cadis.Entrades | ForEach-Object { "$($_.Tipus):$($_.Gia)" }) -join ' ') 'activitat:777 activitat:1447 activitat:1403 buit:' 'Cadis 19: per local/planta (Esc. 1... abans que Local 5), despres per ID NUMERIC (777 abans que 1447), i el local buit al final'
$en1447 = @($cadis.Entrades | Where-Object { $_.Gia -eq '1447' })[0]
AssertEq "$($en1447.Estat)|$($en1447.Sub)|$($en1447.SubFont)" 'vermell|Esc. 1 - Pl. 2 - Pt. 16|cadastre' '1447: precintada, amb la planta/porta del Cadastre'
$en1403 = @($cadis.Entrades | Where-Object { $_.Gia -eq '1403' })[0]
AssertEq "$($en1403.Estat)|$($en1403.EstatText)|$($en1403.Sub)" 'groc|Requeriment|Local 5' '1403: requeriment, Local 5'
$en777 = @($cadis.Entrades | Where-Object { $_.Gia -eq '777' })[0]
AssertEq "$($en777.NoBase)|$($en777.Estat)" 'True|blau' 'una activitat que no es a la base d activitats: es marca, en blau'
$hosp = $perClau['4091106DF2749A']
AssertEq "$(@($hosp.Entrades)[0].Gia)|$(@($hosp.Entrades)[0].MarcatBuit)|$(@($hosp.Entrades)[0].Estat)" '9|True|verd' 'activitat en un local marcat com a buit: es pinta i es marca'
AssertEq @($model.Parceles | Where-Object { $_.Clau -like 'xy:*' }).Count 1 'sense refcat pero amb coordenades: un punt'
$ferro = $perClau['3085213DF2738E']
AssertEq "$(@($ferro.Entrades)[0].Gia)|$(@($ferro.Entrades)[0].SenseEstabliment)" '2000|True' 'activitat sense establiment: amb la refcat de l Excel d activitats'
$r = $model.Resum
AssertEq "$($r.Establiments)|$($r.Buits)|$($r.SenseEstabliment)|$($r.NoBase)|$($r.MarcatsBuit)|$($r.SensePosicio)" '7|1|1|1|1|1' 'el resum (l activitat 11 te establiment, sense posicio: 1)'
# LA LLISTA EMBOLCALLADA DUES VEGADES (el defecte d'octubre de 2026): la crida
# feia @(Read-EstablimentsExcel ...), que ja la torna amb coma, i el model rebia
# UN element que les contenia totes. Totes les activitats sortien "sense
# establiment" i cap local buit. Ara el model la desplega.
# BUITS DUPLICATS (el 1365 de l'usuari): el GIA te el mateix local buit i
# ocupat. El buit no compta com a buit; l'ocupat diu quin establiment es.
$eDup = { param($id, $act, $local = '') [pscustomobject]@{ IdEst = $id; Rc = '2782719DF2728B0001LL'; IdActivitat = $act; Local = $local; Bloc = ''; Escala = ''; Pis = ''; Porta = ''
          Buit = ($act -eq ''); UtmX = 422000.0; UtmY = 4579000.0; Adreca = 'C FRANCESC LAYRET 78'; Carrer = 'FRANCESC LAYRET'; Numero = '78' } }
$mDup = Build-PlanolModel @((& $eDup '1427' ''), (& $eDup '1428' '1365'), (& $eDup '1429' '' '2')) @{} @{} @{}
$enDup = @($mDup.Parceles[0].Entrades)
AssertEq (@($enDup | ForEach-Object { "$($_.Tipus):$($_.IdEst):$($_.BuitDuplicat):$($_.BuitsIguals)" }) -join ' ') 'activitat:1428:False:1427 buit:1429:False: buit:1427:True:' `
         'el buit identic al local ocupat es duplicat (i l ocupat ho diu); el del local 2, un buit de debo'
AssertEq "$($mDup.Resum.Buits)|$($mDup.Resum.BuitsDuplicats)" '1|1' 'el resum els compta a part'
$jDup = @(ConvertTo-PlanolDadesMapa $mDup @{})[0]
AssertEq (@($jDup.e | ForEach-Object { "$($_.t):$($_['ie']):$($_['du']):$($_['bd'])" }) -join ' ') 'a:1428::1427 b:1429:: b:1427:True:' 'al mapa: ie, du i bd nomes quan hi son'
$modelE = Build-PlanolModel (,@($ests)) $acts $estats $unitats
$rE = $modelE.Resum
AssertEq "$($rE.Establiments)|$($rE.Buits)|$($rE.SenseEstabliment)" "$($r.Establiments)|$($r.Buits)|$($r.SenseEstabliment)" 'Build-PlanolModel: amb la llista embolcallada, el mateix resultat (no "tot sense establiment")'
# I el lector de debo, amb un Read-FullaEstesa fals, tal com el crida l'eina.
$capsE = & {
    function Read-FullaEstesa($f, [scriptblock]$cos, [switch]$Desa, [string]$Fulla = '') { & $cos ([pscustomobject]@{ Data = (_Mat $filesE); Rows = $filesE.Count; Headers = $capE }) }
    $e1 = Read-EstablimentsExcel 'x.xls'
    @($e1).Count; [string](@($e1)[0].GetType().Name)
}
AssertEq "$($capsE[0])|$($capsE[1])" "$(@($ests).Count)|PSCustomObject" 'Read-EstablimentsExcel: una llista d establiments, cada un un registre'

# L'ordre per ID GIA es NUMERIC: 9 abans que 10, i no '10' abans que '9'.
$ord = Build-PlanolModel @(
    [pscustomobject]@{ IdEst = 'a'; Rc = '1111111DF1111A0001AA'; IdActivitat = '10'; Local = ''; Bloc = ''; Escala = ''; Pis = ''; Porta = ''; Buit = $false; UtmX = 422000.0; UtmY = 4579000.0; Adreca = '' }
    [pscustomobject]@{ IdEst = 'b'; Rc = '1111111DF1111A0001AA'; IdActivitat = '9'; Local = ''; Bloc = ''; Escala = ''; Pis = ''; Porta = ''; Buit = $false; UtmX = 422000.0; UtmY = 4579000.0; Adreca = '' }
) @{} @{} @{}
AssertEq (@($ord.Parceles[0].Entrades | ForEach-Object { $_.Gia }) -join ',') '9,10' 'ID GIA en ordre numeric (9 abans que 10)'

Write-Host "`n--- ConvertTo-PlanolDadesMapa ---"
$geos = @{ '2295827DF2729E' = $polys }
$mapa = @(ConvertTo-PlanolDadesMapa $model $geos)
$mCadis = @($mapa | Where-Object { $_.k -eq '2295827DF2729E' })[0]
AssertEq @($mCadis.p).Count 2 'la parcel.la amb geometria porta els seus dos poligons'
AssertEq @(@($mCadis.p)[0]).Count 2 'el primer amb el forat'
AssertNear ([double]@(@(@($mCadis.p)[0])[0])[0]) 41.36 0.02 'en graus (latitud)'
# El centre va DINS de la parcel.la. El centroide de l'exterior mes gran,
# (421975, 4579505), cau just al PATI (el forat 421970-421980): fins a l'octubre
# de 2026 l'etiqueta hi anava. Ara, el mig del tram interior mes ample de
# l'horitzontal del centroide: (421960, 4579505).
$llC = Convert-UtmToLatLon 421960.0 4579505.0 31 $true
AssertNear ([double]$mCadis.c[0]) $llC.Lat 0.000002 'el centre, dins de la parcel.la i fora del pati (lat)'
AssertNear ([double]$mCadis.c[1]) $llC.Lon 0.000002 'el centre, dins de la parcel.la i fora del pati (lon)'
$mHosp = @($mapa | Where-Object { $_.k -eq '4091106DF2749A' })[0]
AssertEq @($mHosp.p).Count 0 'sense geometria: cap poligon (surt com un punt)'
Assert ($null -ne $mHosp.c) 'i l etiqueta, a la coordenada de l Excel'
$json = ConvertTo-JsonScript $mapa -Llista -Fondaria 10
$torna = $json | ConvertFrom-Json
AssertEq @($torna).Count @($mapa).Count 'el JSON del mapa es valid i hi son totes'

Write-Host "`n--- Les capcaleres REALS de l'Excel d'establiments ---"
# Tal com venen al fitxer de l'usuari (2026-10-05): "Emp._Numero_" amb accent i
# guions baixos, "Emp. N? Local". Sense el numero no es pot buscar l'entrada.
$capReal = @('ID Establiment GIA', 'ID Establiment Gencat', 'Ref. cadastral', 'UTM X', 'UTM Y', 'Emp. Nucli o barri', 'Emp. Municipi', 'Emp. CP', 'Emp. Tipus via', 'Emp. Carrer',
             ('Emp._N' + [char]0x00FA + 'mero_'), 'Emp. Lletra', 'Emp. Bloc', 'Emp. Km', ('Emp. N' + [char]0x00BA + ' Local'), 'Emp. Escala', 'Emp. Pis', 'Emp. Porta', 'Local buit', 'ID Activitat')
AssertEq "$(Find-HeaderColumn $capReal 'Emp. Carrer')|$(Find-HeaderColumn $capReal 'Emp. Numero')|$(Find-HeaderColumn $capReal 'Emp. N Local')|$(Find-HeaderColumn $capReal 'ID Activitat')" '10|11|15|20' 'carrer, numero, local i ID Activitat es troben a les capcaleres reals'

Write-Host "`n--- Allotjaments turistics (CCAE 552/5520) ---"
AssertEq (@(Get-ColumnesCcae @('ID Activitat', 'CCAE Descripcio', 'CCAE Codi')) -join ',') '3' 'la columna es "CCAE Codi" (la que diu l usuari), encara que n hi hagi d altres amb CCAE'
AssertEq (@(Get-ColumnesCcae @('ID Activitat', 'CCAE secundari', 'CCAE principal')) -join ',') '3' 'sense "CCAE Codi": la principal'
AssertEq @(Get-ColumnesCcae @('ID Activitat', 'Nom')).Count 0 'sense cap CCAE: cap columna'
foreach ($cas in @(@('5520', $true), @('55.20', $true), @('552', $true), @(5520.0, $true), @('5520 - Allotjaments turistics', $true), @('5510', $false), @('4520', $false), @('', $false), @('I5520', $true))) {
    AssertEq (Test-CcaeAllotjamentTuristic @($cas[0])) $cas[1] ("CCAE '" + $cas[0] + "' -> " + $cas[1])
}

Write-Host "`n--- PlanolGeometria: dins, vora, entrada ---"
$quad = { param($x0, $y0, $x1, $y1) [pscustomobject]@{ Anells = @(,([double[]]@($x0, $y0, $x1, $y0, $x1, $y1, $x0, $y1, $x0, $y0))) } }
$q20 = & $quad 0 0 20 20
Assert (Test-PuntDinsPoligons 5 5 @($q20)) 'un punt dins del quadrat'
Assert (-not (Test-PuntDinsPoligons 25 5 @($q20))) 'i un de fora'
$ambForat = [pscustomobject]@{ Anells = @(([double[]]@(0, 0, 20, 0, 20, 20, 0, 20)), ([double[]]@(8, 8, 12, 8, 12, 12, 8, 12))) }
Assert (-not (Test-PuntDinsPoligons 10 10 @($ambForat))) 'dins del forat (el pati) no es dins de la parcel.la'
$aF = Resolve-AncoraEntrada 10 -1 @($q20)
AssertEq "$([math]::Round($aF.X,2))|$([math]::Round($aF.Y,2))|$(Get-DireccioEtiqueta $aF.Dx $aF.Dy)" "10|$($Script:PlanolEntradaMargeM)|t" 'portal a 1 m fora (al carrer): l etiqueta, a dins i creixent cap amunt (cap a dins)'
$aV = Resolve-AncoraEntrada 10 0 @($q20)
Assert (Test-PuntDinsPoligons $aV.X $aV.Y @($q20)) 'portal just sobre la linia: l etiqueta, a dins'
AssertEq (Resolve-AncoraEntrada 10 -20 @($q20)) $null 'portal a 20 m: no es d aquesta parcel.la'
$aD = Resolve-AncoraEntrada 10 10 @($q20)
AssertEq "$($aD.X)|$($aD.Y)" '10|10' 'portal ben dins: es queda on es'
$aE = Resolve-AncoraEntrada 21 10 @($q20)
AssertEq (Get-DireccioEtiqueta $aE.Dx $aE.Dy) 'l' 'portal a la facana de la dreta: l etiqueta creix cap a l esquerra'
$ele = [pscustomobject]@{ Anells = @(,([double[]]@(0, 0, 30, 0, 30, 10, 10, 10, 10, 30, 0, 30))) }
$pi = Get-PuntInterior @($ele)
Assert (Test-PuntDinsPoligons $pi.X $pi.Y @($ele)) 'Get-PuntInterior: dins, tambe en una L'

Write-Host "`n--- PlanolGeometria: juntar parcel.les ---"
$jA = Join-Poligons @((& $quad 0 0 10 10), (& $quad 10 0 20 10))
AssertEq "$($jA.Toquen)|$($jA.Unit)|$(@($jA.Polys).Count)" 'True|True|1' 'dues parcel.les amb un costat comu: una sola forma'
AssertNear (Get-AreaAnell @($jA.Polys)[0].Anells[0]) 200 0.01 'i l area es la de les dues'
$jP = Join-Poligons @((& $quad 0 0 10 10), (& $quad 10 2 20 6))
AssertEq "$($jP.Toquen)|$($jP.Unit)|$(@($jP.Polys).Count)" 'True|True|1' 'una toca nomes un tros del costat de l altra: tambe s ajunten'
AssertNear (Get-AreaAnell @($jP.Polys)[0].Anells[0]) 140 0.01 'i l area es la suma'
$jN = Join-Poligons @((& $quad 0 0 10 10), (& $quad 15 0 25 10))
AssertEq "$($jN.Toquen)|$(@($jN.Polys).Count)" 'False|2' 'separades: no es toquen i queden les dues'
$jH = Join-Poligons @($ambForat, (& $quad 20 0 30 20))
AssertEq "$($jH.Unit)|$(@(@($jH.Polys)[0].Anells).Count)" 'True|2' 'el pati d una es conserva en juntar-la'

Write-Host "`n--- L'ID GIA a la COORDENADA UTM de l'Excel d'activitats ---"
# L'usuari (octubre 2026): "dibuixa les etiquetes amb el ID GIA segons les
# coordenades UTM de la base de dades d'activitats". Una parcel.la de 40 x 20
# al carrer Progres (la facana, a baix).
$eP = { param($g, $num, $estat = 'groc', $ax = $null, $ay = $null) [pscustomobject]@{ Tipus = 'activitat'; Gia = $g; Titular = ''; Activitat = ''; Sub = ''; SubFont = ''; Estat = $estat; EstatText = ''; Precinte = $false; MarcatBuit = $false; SenseEstabliment = $false; NoBase = $false; NInformes = 0; Adreca = ''; Rc = ''; Carrer = 'Progres'; Numero = $num; Turistic = $false; Classificacio = 'III'; ActX = $ax; ActY = $ay } }
$pcP = [pscustomobject]@{ Clau = '3678311DF2737H'; Rc = '3678311DF2737H'; X = 422020.0; Y = 4579010.0; Entrades = @(
    (& $eP '1340' '73' 'groc' 422008.0 4579005.0),       # dins
    (& $eP '288' '75' 'blau' 422030.0 4578999.0),        # 1 m al carrer: a dins
    (& $eP '999' '81' 'groc' 423000.0 4580000.0),        # lluny: vermell
    (& $eP '777' '73' 'groc')) }                         # sense coordenada: vermell
$geoP = @{ '3678311DF2737H' = @((& $quad 422000 4579000 422040 4579020)) }
$portP = @{ '3678311DF2737H' = @(
    [pscustomobject]@{ Numero = '73'; Via = 'CL PROGRES'; X = 422008.0; Y = 4578999.5 },
    [pscustomobject]@{ Numero = '75'; Via = 'CL PROGRES'; X = 422030.0; Y = 4578999.5 }) }
$mP = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcP) }) $geoP $portP)[0]
$etq = @($mP.l)
$e1340 = @($etq | Where-Object { @($_.g) -contains '1340' })[0]; $e288 = @($etq | Where-Object { @($_.g) -contains '288' })[0]
$ll1340 = Convert-UtmToLatLon 422008.0 4579005.0 31 $true
AssertNear ([double]$e1340.c[0]) $ll1340.Lat 0.0000005 'el 1340, a la seva coordenada (lat)'
AssertNear ([double]$e1340.c[1]) $ll1340.Lon 0.0000005 'el 1340, a la seva coordenada (lon)'
AssertEq "$($e1340.v)|$(@($e1340.g) -join ',')" '0|1340' 'una etiqueta per al 1340, no vermella'
$ll288 = Convert-UtmToLatLon 422030.0 (4579000.0 + $Script:PlanolEntradaMargeM) 31 $true
AssertNear ([double]$e288.c[0]) $ll288.Lat 0.0000005 'el 288 (1 m al carrer) es posa DINS de la parcel.la'
AssertEq $e288.d 't' '...creixent cap a dins'
$eV = @($etq | Where-Object { $_.v -eq 1 })
AssertEq "$($eV.Count)|$(@($eV[0].g) -join ',')" '1|999,777' 'la coordenada lluny i la que no en te: en vermell'
AssertNear ([double]$eV[0].c[0]) ([double]$mP.c[0]) 0.000001 '...al centre de la parcel.la'
AssertEq (@($mP.e | ForEach-Object { "$($_.g):$($_.x)" }) -join ' ') '1340:1 288:1 999:0 777:3' 'cada activitat diu on ha anat (1 coordenada, 0 fora, 3 sense)'
AssertEq (@($mP.e)[0].cl) 'III' 'la classificacio (annex) arriba al mapa'
# Les DUES adreces (l'usuari: "no te per que ser exactament la mateixa"): la del
# Cadastre (el portal amb el numero de l'establiment) i les de la parcel.la.
AssertEq (@($mP.e | ForEach-Object { "$($_.g)=$($_.ca)" }) -join ' | ') '1340=CL PROGRES 73 | 288=CL PROGRES 75 | 999= | 777=CL PROGRES 73' 'cada entrada porta l adreca del Cadastre amb el seu numero (cap si no hi es)'
AssertEq (@($mP.pa) -join ' / ') 'CL PROGRES 73 / CL PROGRES 75' 'la parcel.la porta les seves adreces del Cadastre'
AssertEq (@(Get-AdrecesCadastre @{ 'R' = @([pscustomobject]@{ Via = 'CL A'; Numero = '11' }, [pscustomobject]@{ Via = 'CL A'; Numero = '9' }, [pscustomobject]@{ Via = 'CL A'; Numero = '9' }) } @('R')) -join ' / ') 'CL A 9 / CL A 11' 'les adreces del Cadastre, sense repetir i el 9 abans que l 11'
AssertEq @(Get-AdrecesCadastre $null @('R')).Count 0 'sense portals, cap adreca del Cadastre'
# La LINIA DE PUNTS: cada etiqueta a la seva coordenada porta r, el punt de la
# parcel.la al Cadastre; les vermelles del centre, no; a menys d'1 m, tampoc.
$puntsP = @{ '3678311DF2737H' = @(422020.0, 4579010.0) }
$mR = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcP) }) $geoP $portP $puntsP)[0]
$llRef = Convert-UtmToLatLon 422020.0 4579010.0 31 $true
$eR1340 = @($mR.l | Where-Object { @($_.g) -contains '1340' })[0]
AssertNear ([double]$eR1340.r[0]) $llRef.Lat 0.0000005 'l etiqueta del 1340 porta el punt de la parcel.la (lat)'
AssertNear ([double]$eR1340.r[1]) $llRef.Lon 0.0000005 '...i (lon)'
AssertEq (@($mR.l | Where-Object { $_.v -eq 1 } | Where-Object { $null -ne $_['r'] }).Count) 0 'la vermella del centre no en porta'
$pcIgual = [pscustomobject]@{ Clau = 'I'; Rc = '3678311DF2737H'; X = 0.0; Y = 0.0; Entrades = @((& $eP '42' '73' 'groc' 422020.3 4579010.3)) }
$mI = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcIgual) }) $geoP $portP $puntsP)[0]
AssertEq (@($mI.l)[0].Contains('r')) $false 'a menys d 1 m del punt del Cadastre (encara no corregida): sense linia'
AssertEq (@(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcP) }) $geoP $portP)[0].l | Where-Object { $_.Contains('r') }).Count 0 'sense punts del Cadastre: cap linia'
# El MATEIX punt (abans de repassar-les, les d'una parcel.la solen coincidir): una etiqueta.
$pcM = [pscustomobject]@{ Clau = 'M'; Rc = '3678311DF2737H'; X = 0.0; Y = 0.0; Entrades = @((& $eP '1340' '73' 'groc' 422008.0 4579005.0), (& $eP '50' '73' 'verd' 422009.0 4579006.0), (& $eP '60' '73' 'verd' 422030.0 4579010.0)) }
$mM = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcM) }) $geoP $portP)[0]
AssertEq (@($mM.l | ForEach-Object { @($_.g) -join ',' }) -join ' | ') '1340,50 | 60' 'dues al mateix punt, una etiqueta; la de 20 m enlla, la seva'
# Una activitat amb DOS establiments (dues parcel.les) i UNA coordenada: l'ID a
# la parcel.la on cau i, a l'altra, al centre SENSE vermell.
$pcD1 = [pscustomobject]@{ Clau = '4444444DF4444A'; Rc = '4444444DF4444A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '700' '1' 'groc' 422105.0 4579005.0), (& $eP '701' '1' 'blau' 422105.0 4579005.0)) }
$pcD2 = [pscustomobject]@{ Clau = '5555555DF5555A'; Rc = '5555555DF5555A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '700' '9' 'groc' 422105.0 4579005.0)) }
$geoD = @{ '4444444DF4444A' = @((& $quad 422100 4579000 422110 4579010)); '5555555DF5555A' = @((& $quad 422200 4579000 422210 4579010)) }
$mD = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcD1, $pcD2) }) $geoD @{})
$mD2 = @($mD | Where-Object { $_.rc -eq '5555555DF5555A' })[0]
AssertEq (@($mD | ForEach-Object { $r = $_.rc; @($_.e) | ForEach-Object { "$($r.Substring(0,1))$($_.g):$($_.x)" } }) -join ' ') '4700:1 4701:1 5700:2' 'la coordenada mana: a la 4444 l ID, a la 5555 "es a l altre establiment"'
AssertEq "$(@($mD2.l).Count)|$(@($mD2.l)[0].v)|$(@($mD2.l)[0].d)" '1|0|c' '...i alli al centre, sense vermell'
# Una coordenada DINS de la parcel.la d'una ALTRA activitat (encara que sigui a
# 2 m de la seva): esta malament, en vermell.
$pcF1 = [pscustomobject]@{ Clau = '6666666DF6666A'; Rc = '6666666DF6666A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '800' '1' 'groc' 422312.0 4579005.0)) }
$pcF2 = [pscustomobject]@{ Clau = '7777777DF7777A'; Rc = '7777777DF7777A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '801' '3' 'groc' 422315.0 4579005.0)) }
$geoF = @{ '6666666DF6666A' = @((& $quad 422300 4579000 422310 4579010)); '7777777DF7777A' = @((& $quad 422310 4579000 422320 4579010)) }
$mF = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcF1, $pcF2) }) $geoF @{})
AssertEq (@($mF | ForEach-Object { @($_.e) | ForEach-Object { "$($_.g):$($_.x)" } }) -join ' ') '800:0 801:1' 'dins de la parcel.la del 801, el 800 va en vermell (no s arrossega a la seva)'
# Sense dibuix (un punt): no se sap, al punt i sense vermell.
$mS = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcP) }) @{} $portP)[0]
AssertEq "$(@($mS.l).Count)|$(@($mS.l)[0].v)|$(@($mS.e)[0].x)" '1|0|-1' 'sense dibuix de la parcel.la: al punt, sense vermell (no se sap)'
AssertEq (@(Get-CapsaPoligons @((& $quad 1 2 3 4))) -join ',') '1,2,3,4' 'la capsa d uns poligons'
AssertEq (Get-CapsaPoligons @()) $null 'sense poligons, cap capsa'
# Juntar: dues parcel.les que es toquen amb les MATEIXES activitats.
$pcJ1 = [pscustomobject]@{ Clau = '1111111DF1111A'; Rc = '1111111DF1111A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '500' '1')) }
$pcJ2 = [pscustomobject]@{ Clau = '2222222DF2222A'; Rc = '2222222DF2222A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '500' '3')) }
$pcJ3 = [pscustomobject]@{ Clau = '3333333DF3333A'; Rc = '3333333DF3333A'; X = 0.0; Y = 0.0; Entrades = @((& $eP '500' '5'), (& $eP '501' '5')) }
$geoJ = @{ '1111111DF1111A' = @((& $quad 422000 4579000 422010 4579010)); '2222222DF2222A' = @((& $quad 422010 4579000 422020 4579010)); '3333333DF3333A' = @((& $quad 422020 4579000 422030 4579010)) }
$mJ = @(ConvertTo-PlanolDadesMapa ([pscustomobject]@{ Parceles = @($pcJ1, $pcJ2, $pcJ3) }) $geoJ @{})
AssertEq $mJ.Count 2 'les dues de nomes el 500 es junten; la que te el 500 i el 501, no'
$mJ1 = @($mJ | Where-Object { @($_.rcs).Count -eq 2 })[0]
AssertEq "$(@($mJ1.p).Count)|$(@($mJ1.e).Count)" '1|2' 'la forma juntada: un sol poligon, amb els dos establiments'

Write-Host "`n--- Build-PlanolHtml: la plantilla (PlanolMapa.html) ---"
$metaT = [pscustomobject]@{ BaseActivitats = 'A.xls'; BaseEstabliments = 'E.xls'; BaseInformes = 'Base'; Avisos = @('un avis </script>') }
$htmlT = Build-PlanolHtml $mapa $metaT
Assert (-not $htmlT.Contains('{{')) 'cap marca {{...}} sense omplir'
Assert ($htmlT.Contains('Pl' + [char]0x00E0 + 'nol activitats')) 'la plantilla es llegeix en UTF-8 (accents intactes)'
Assert ($htmlT -match 'var PARCELES = (\[.*?\]);\s*</script>') 'hi ha les dades'
$parsT = $null; try { $parsT = $Matches[1] | ConvertFrom-Json } catch { }
AssertEq @($parsT).Count @($mapa).Count 'el JSON de les dades es valid'
Assert ($htmlT.Contains('un avis &lt;/script&gt;')) 'els avisos van escapats'
$htmlBuit = Build-PlanolHtml @() ([pscustomobject]@{ BaseActivitats = ''; BaseEstabliments = ''; BaseInformes = ''; Avisos = @() })
Assert ($htmlBuit.Contains('var PARCELES = [];')) 'sense cap parcel.la: llista buida (i la pagina arrenca)'
Assert (-not $htmlBuit.Contains('id="avisos"')) 'sense avisos: cap franja d avisos'

Write-Host "`n--- El fons dels mapes: MapaFons.js, res d'OpenStreetMap ---"
# OpenStreetMap rebutja les pagines obertes des del disc ("Access blocked"): cap
# mapa de rutes/ hi pot demanar rajoles, i el fons ve d'un sol lloc.
Assert ($htmlT.Contains('function afegeixFonsMapa') -and $htmlT.Contains('var FONS = afegeixFonsMapa(map)')) 'el Planol porta el fons comu (MapaFons.js) i el fa servir'
Assert ($htmlT.Contains('geoserveis.icgc.cat')) 'el primer fons, l''ICGC'
$dirRutesF = Split-Path -Parent $PSScriptRoot | Join-Path -ChildPath 'rutes'
$ambOsm = @(Get-ChildItem -LiteralPath $dirRutesF -File | Where-Object { $_.Extension -in '.html', '.ps1', '.js' } |
            Where-Object { [System.IO.File]::ReadAllText($_.FullName) -match 'tile\.openstreetmap\.org' } | ForEach-Object { $_.Name })
AssertEq ($ambOsm -join ', ') '' 'cap fitxer de rutes/ demana rajoles a OpenStreetMap'
# El planol del Cadastre (WMS) tambe es del fons comu: Planol i Coordenades
# l'encenen amb la mateixa capa. Cap altre fitxer no el pot tornar a definir.
$ambWms = @(Get-ChildItem -LiteralPath $dirRutesF -File | Where-Object { $_.Extension -in '.html', '.ps1', '.js' } |
            Where-Object { [System.IO.File]::ReadAllText($_.FullName) -match 'Cartografia/WMS/ServidorWMS' } | ForEach-Object { $_.Name })
AssertEq ($ambWms -join ', ') 'MapaFons.js' 'el WMS del Cadastre, nomes a MapaFons.js'
Assert ($htmlT.Contains("lligaCasellaCadastre(map, FONS, 'f-cadastre')")) 'la casella del Planol encen la capa comuna'
$senseFons = @(Get-ChildItem -LiteralPath $dirRutesF -File | Where-Object { $_.Extension -in '.html', '.ps1' } |
               Where-Object { $t = [System.IO.File]::ReadAllText($_.FullName); $t.Contains("L.map('map'") -and -not $t.Contains('afegeixFonsMapa(map)') } | ForEach-Object { $_.Name })
AssertEq ($senseFons -join ', ') '' 'tots els mapes de rutes/ fan servir el fons comu'
Assert (-not (Get-MapaFonsJs).Contains('</')) 'el fons va dins d''un <script> sense trencar-lo (cap "</")'

Write-Host "`n--- Les consultes al Cadastre (servei fals) ---\"
$tmpP = Join-Path ([System.IO.Path]::GetTempPath()) ('planol-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpP -Force | Out-Null
$Script:Crides = New-Object System.Collections.ArrayList
function Get-CacheCadastrePath([string]$fitxer) { return (Join-Path $tmpP $fitxer) }
function Invoke-CadastreGet([string]$url) {
    [void]$Script:Crides.Add($url)
    if ($url.Contains('wfsCP') -and $url.Contains('2295827DF2729E')) { return $xmlCp }
    if ($url.Contains('Consulta_DNPRC') -and $url.Contains('2295827DF2729E0011RQ')) { return [System.IO.File]::ReadAllText((Join-Path $dades 'dnprc-exemple.xml')) }
    if ($url.Contains('Consulta_DNPRC')) { return [System.IO.File]::ReadAllText((Join-Path $dades 'dnprc-error.xml')) }
    return $null
}
try {
    $pq = Get-ParcelesCadastre @('2295827DF2729E', '9999999DF9999Z')
    $g = $pq.Geometries
    AssertEq @($g['2295827DF2729E']).Count 2 'geometria: dos poligons'
    AssertEq @($g['9999999DF9999Z']).Count 0 'servei caigut: cap (llista buida)'
    AssertEq "$(@($pq.Punts['2295827DF2729E']) -join ',')|$($pq.Punts.ContainsKey('9999999DF9999Z'))" '421975,4579505|False' 'i el punt de la parcel.la (la caiguda, sense)'
    $Script:Crides.Clear()
    $pq2 = Get-ParcelesCadastre @('2295827DF2729E')
    $g2 = $pq2.Geometries
    AssertEq "$($Script:Crides.Count)|$(@($g2['2295827DF2729E']).Count)|$(@(@($g2['2295827DF2729E'])[0].Anells).Count)" '0|2|2' 'la segona vegada, de la memoria cau i sencera (poligons i forats)'
    AssertEq (@($pq2.Punts['2295827DF2729E']) -join ',') '421975,4579505' '...amb el punt (rellegit del JSON)'
    AssertEq (_ConsultaParcelesCadastre).Fitxer 'parceles2.json' 'un fitxer nou: les entrades velles (sense punt) no es donen per bones'
    $u = Get-UnitatsCadastre @('2295827DF2729E0011RQ', '2295827DF2729E0003XL')
    AssertEq "$($u['2295827DF2729E0011RQ'].Planta)|$($null -eq $u['2295827DF2729E0003XL'])" '02|True' 'unitats: la que existeix i la que no'
    $Script:Crides.Clear()
    $u2 = Get-UnitatsCadastre @('2295827DF2729E0011RQ')
    AssertEq "$($Script:Crides.Count)|$($u2['2295827DF2729E0011RQ'].Porta)" '0|16' 'unitats: la segona vegada, de la memoria cau'
} catch {
    Assert $false ("Consultes: una excepcio s'ha escapat del bloc de proves -> " + $_.Exception.Message)
} finally {
    Remove-Item -LiteralPath $tmpP -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Invoke-PlanolGenera en SILENCI (el mode automatic setmanal) ---"
# Sense cap Excel d'activitats: torna l'error (no obre cap finestra).
$vellsG = @{ A = $ActivitatsDir; L = $LocalActivitatsDir }
$buitG = Join-Path ([System.IO.Path]::GetTempPath()) ('planol-gen-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $buitG -Force | Out-Null
try {
    $ActivitatsDir = $buitG; $LocalActivitatsDir = $buitG
    $rG = Invoke-PlanolGenera $true
    AssertEq "$($rG.Ok)|$($rG.Error.Contains('No s''ha trobat cap base de dades'))" 'False|True' 'sense Excel: no fa res i ho diu (sense finestres)'
    # La barra de progres, en silenci, no s'obre: la feina es fa igual.
    $Script:PlanolSilenci = $true
    $fG = _PlanolAmbProgres @('a', 'b') 'x' { param($l, $p) @{ N = @($l).Count; P = ($null -eq $p) } }
    AssertEq "$($fG.Resultat.N)|$($fG.Resultat.P)|$($fG.Cancelat)" '2|True|False' 'en silenci, la feina sense barra (i sense cancel.lar)'
    $Script:PlanolSilenci = $false
} catch {
    Assert $false ("Silenci: una excepcio s'ha escapat -> " + $_.Exception.Message)
} finally {
    $ActivitatsDir = $vellsG.A; $LocalActivitatsDir = $vellsG.L
    Remove-Item -LiteralPath $buitG -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Consultar l'ultim o fer-ne un de nou ---"
$dirU = Join-Path ([System.IO.Path]::GetTempPath()) ('planol-ultim-' + [guid]::NewGuid().ToString('N'))
try {
    AssertEq "$($null -eq (Get-PlanolUltim $dirU))" 'True' 'sense carpeta: cap planol (es fa directament, sense preguntar)'
    New-Item -ItemType Directory -Path $dirU -Force | Out-Null
    AssertEq "$($null -eq (Get-PlanolUltim $dirU))" 'True' 'carpeta buida: cap planol'
    foreach ($n in @('Planol_2026-10-01_090000.html', 'Planol_2026-10-05_130000.html', 'altre.html')) {
        [System.IO.File]::WriteAllText((Join-Path $dirU $n), 'x')
    }
    (Get-Item -LiteralPath (Join-Path $dirU 'Planol_2026-10-01_090000.html')).LastWriteTime = [datetime]'2026-10-01 09:00'
    (Get-Item -LiteralPath (Join-Path $dirU 'Planol_2026-10-05_130000.html')).LastWriteTime = [datetime]'2026-10-05 13:00'
    (Get-Item -LiteralPath (Join-Path $dirU 'altre.html')).LastWriteTime = [datetime]'2026-10-06 10:00'
    $u = Get-PlanolUltim $dirU
    AssertEq $u.Name 'Planol_2026-10-05_130000.html' 'l ultim planol es el mes nou (i nomes els Planol_*.html)'
    $tU = Get-PlanolPreguntaText $u ([datetime]'2026-10-07 11:00')
    Assert ($tU.Contains('05/10/2026 a les 13:00') -and $tU.Contains('fa 2 dies') -and $tU.Contains('fer-ne un de nou')) 'la pregunta diu de quan es l ultim'
    Assert ((Get-PlanolPreguntaText $u ([datetime]'2026-10-05 18:00')).Contains('(avui)')) 'fet avui'
    Assert ((Get-PlanolPreguntaText $u ([datetime]'2026-10-06 08:00')).Contains('(ahir)')) 'fet ahir'
} catch {
    Assert $false ("Ultim planol: una excepcio s'ha escapat -> " + $_.Exception.Message)
} finally {
    Remove-Item -LiteralPath $dirU -Recurse -Force -ErrorAction SilentlyContinue
}
# La finestra de la pregunta fa servir el peu comu (UiFinestra.ps1, via Ruta.ps1).
AssertEq ((@(Get-Command _AddPeuBotons, _PeuPosicions, _PeuAmple, Show-EinaTria -ErrorAction SilentlyContinue)).Count) 4 'la pregunta te el peu de botons carregat dins de l eina'
# La tria va abans de generar: Invoke-PlanolMain no comenca a fer-ne un de nou
# sense haver mirat si ja n'hi ha.
$srcMain = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'rutes' 'Planol.ps1')))
$cosMain = $srcMain.Substring($srcMain.IndexOf('function Invoke-PlanolMain'))
Assert ($cosMain.IndexOf('Get-PlanolUltim') -ge 0 -and $cosMain.IndexOf('Get-PlanolUltim') -lt $cosMain.IndexOf('Invoke-PlanolGenera')) 'el boto pregunta abans de generar'

exit (Write-TestSummary 'RESULTAT')
