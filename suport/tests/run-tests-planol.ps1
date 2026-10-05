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

$capA = @('ID Activitat', 'Ref. cadastral', 'UTM X', 'UTM Y', 'Emp. Tipus via', 'Emp. Carrer', 'Emp. Numero', 'Activitat principal', 'Nom comercial activitat', 'Camp Info 1 - Nom', 'Camp Info 1 - Valor')
$filesA = @(
    $capA,
    @([double]1447, '2295827DF2729E0011RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', 'BAR', 'EL RACO', 'PRECINTE ACTIVITAT?', 'SI, PRECINTAT 01/10/2026'),
    @([double]1403, '2295827DF2729E0008RQ', [double]421968.09, [double]4579505.55, 'C', 'CADIS', '19', 'TALLER', '', '', ''),
    @([double]9, '4091106DF2749A0006XJ', [double]423912.16, [double]4578928.25, 'CTRA', 'HOSPITALET', '147', 'OFICINES', 'ACME', '', ''),
    @([double]10, '', [double]422500.0, [double]4579000.0, 'C', 'ENLLOC', '1', 'BOTIGA', '', '', ''),
    @([double]11, '', $null, $null, 'C', 'ENLLOC', '2', 'BOTIGA', '', '', ''),
    @([double]2000, '3085213DF2738E0001AB', [double]422800.0, [double]4579200.0, 'PG', 'FERROCARRILS', '177', 'MAGATZEM', '', '', '')
)
$acts = ConvertFrom-FullaActivitatsPlanol (_Mat $filesA) $filesA.Count $capA
AssertEq $acts.Count 6 'sis activitats'
AssertEq $acts['1447'].Precinte $true 'precintada pel camp lliure'
AssertEq $acts['1403'].Precinte $false 'la resta, no'
AssertEq $acts['9'].Nom 'ACME' 'el nom comercial'

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
# L'etiqueta va al centre de l'exterior mes gran: (421975, 4579505).
$llC = Convert-UtmToLatLon 421975.0 4579505.0 31 $true
AssertNear ([double]$mCadis.c[0]) $llC.Lat 0.000002 'l etiqueta, al centre de la parcel.la (lat)'
AssertNear ([double]$mCadis.c[1]) $llC.Lon 0.000002 'l etiqueta, al centre de la parcel.la (lon)'
$mHosp = @($mapa | Where-Object { $_.k -eq '4091106DF2749A' })[0]
AssertEq @($mHosp.p).Count 0 'sense geometria: cap poligon (surt com un punt)'
Assert ($null -ne $mHosp.c) 'i l etiqueta, a la coordenada de l Excel'
$json = ConvertTo-JsonScript $mapa -Llista -Fondaria 10
$torna = $json | ConvertFrom-Json
AssertEq @($torna).Count @($mapa).Count 'el JSON del mapa es valid i hi son totes'

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
    $g = Get-GeometriesParceles @('2295827DF2729E', '9999999DF9999Z')
    AssertEq @($g['2295827DF2729E']).Count 2 'geometria: dos poligons'
    AssertEq @($g['9999999DF9999Z']).Count 0 'servei caigut: cap (llista buida)'
    $Script:Crides.Clear()
    $g2 = Get-GeometriesParceles @('2295827DF2729E')
    AssertEq "$($Script:Crides.Count)|$(@($g2['2295827DF2729E']).Count)|$(@(@($g2['2295827DF2729E'])[0].Anells).Count)" '0|2|2' 'la segona vegada, de la memoria cau i sencera (poligons i forats)'
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

exit (Write-TestSummary 'RESULTAT')
