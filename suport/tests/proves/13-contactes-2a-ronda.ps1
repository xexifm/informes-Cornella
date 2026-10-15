# El repas de contactes, 2a ronda (versio 2026-10-15.2): el representant legal
# (la signatura "(R: CIF)" amb espai i sense DNI, l'interessat persona fisica
# d'una entitat, l'autoritzacio sense "Jo", els noms brossa, les etiquetes que
# quedaven), els falsos "es el tecnic", els documents sense GIA i les bases
# (informes-db.json, contactes-db.json) que no es poden tornar a fer de zero en
# silenci.
#
# TOTES LES DADES SON INVENTADES (domini .test, NIF i telefons de mostra).
# Es DOT-SOURCE des de run-tests.ps1 (mateix ambit i comptador d'asserts).

Write-Host "`n--- A1. La signatura (R: CIF) amb espai, i sense el DNI ---"
$sgA1 = @(_CtSignaturesRep "Pagina 1`nMARIA EXEMPLE PROVA (R: B00000000) Firmado digitalmente por MARIA EXEMPLE PROVA (R: B00000000)`nFecha: 2025.01.01")
AssertEq (@($sgA1 | ForEach-Object { $_.nom + '/' + $_.nif + '/' + $_.cif }) -join ',') 'MARIA EXEMPLE PROVA//B00000000' 'A1: "(R: B...)" amb espai i el nom sense DNI (una sola vegada)'
$sgA1b = @(_CtSignaturesRep "Firmado digitalmente por 00000000T MARIA EXEMPLE (R: B000 00000) Motivo: x. Firmado digitalmente por: 22222222J JOAN ALTRE (R:B33333333)")
AssertEq (@($sgA1b | ForEach-Object { $_.nom + '/' + $_.nif + '/' + $_.cif }) -join ',') 'MARIA EXEMPLE/00000000T/B00000000,JOAN ALTRE/22222222J/B33333333' 'A1: el CIF amb un espai pel mig i dues signatures seguides'
$sgA1c = @(_CtSignaturesRep "MARIA EXEMPLE (R: B00000000) Firmado digitalmente por 00000000T MARIA EXEMPLE (R: B00000000)")
AssertEq (@($sgA1c | ForEach-Object { $_.nom + '/' + $_.nif }) -join ',') 'MARIA EXEMPLE/00000000T' 'A1: la mateixa persona amb DNI i sense, una sola (amb el DNI)'
$ctx2 = _CtContext @{} @{} @{ Emails = @{}; Noms = @(); Dominis = @{} } @()
$nouDoc = { param([string]$ruta, [string]$tipus, $int, $rep, [string]$data, $sigs = @(), $aut = $null)
    $e = _CtDocNou $tipus
    $e.interessat = $int; $e.representant = $rep; $e.data = $data; $e.signatures = @($sigs); $e.autoritzacio = $aut
    return @{ ruta = $ruta; modificat = ($data + 'T00:00:00Z'); extret = $e }
}
$aA1 = Get-ContactesActivitat '9601' @((& $nouDoc 'GIA 9601\Instancia_generica.pdf' 'instancia' (_CtPersona 'EXEMPLE SL' 'B00000000' 'info@exemple-sl.test') $null '2025-02-01' $sgA1)) @{ TITULAR = 'EXEMPLE SL'; NIF = 'B00000000' } $ctx2
AssertEq "$($aA1.representant_legal.nom)|$($aA1.representant_legal.via)|$($aA1.representant_legal.confianca)" 'MARIA EXEMPLE PROVA|signatura|segur' 'A1: la signatura sense DNI tambe fa el representant legal'

Write-Host "`n--- A2. L'interessat persona fisica d'una entitat ---"
$aA2 = Get-ContactesActivitat '9602' @((& $nouDoc 'GIA 9602\etram-tramit.xml' 'etram' (_CtPersona 'Marta Gerent Prova' '33333333P' 'marta.gerent@correu.test' '' '677777777') $null '2025-05-01')) @{ TITULAR = 'EMPRESA PROVA SL'; NIF = 'B88888888' } $ctx2
AssertEq "$($aA2.representant_legal.nom)|$($aA2.representant_legal.via)|$($aA2.representant_legal.confianca)" 'Marta Gerent Prova|interessat|probable' 'A2: candidata a representant legal encara que el correu no sigui el de l''empresa'

Write-Host "`n--- A3. L'autoritzacio sense 'Jo' (GIA 5) ---"
$tA3 = "AUTORITZACI$([char]0x00D3) PER A LA TRAMITACI$([char]0x00D3)`nMARIA EXEMPLE PROVA amb DN(I): 00000000T, en nom i representaci$([char]0x00F3) de l'empresa EXEMPLE DISTRIBUCIONS SL, amb NIF B00000000, i domicili a Cornell$([char]0x00E0), designa com a T$([char]0x00E8)cnica Responsable del projecte de l'activitat i de la tramitaci$([char]0x00F3) de la llic$([char]0x00E8)ncia davant l'Ajuntament de Cornell$([char]0x00E0) de Llobregat a LAIA TECNICA INVENTADA, Enginyera Industrial, col$([char]0x00B7)legiada 12345."
$rA3 = Read-AutoritzacioText $tA3
AssertEq "$($rA3.autoritzacio.signant.nom)|$($rA3.autoritzacio.signant.nif)|$(@($rA3.autoritzacio.autoritzats)[0].nom)" 'MARIA EXEMPLE PROVA|00000000T|LAIA TECNICA INVENTADA' 'A3: el que signa al principi de la seva linia (amb "DN(I)") i l''autoritzada despres de l''ultima " a "'
$rA3b = Read-AutoritzacioText "La sotasignada MARIA EXEMPLE PROVA amb DNI 00000000T, en nom i representaci$([char]0x00F3) de PROVES SA, amb NIF A87654321, autoritza a Pere Tecnic Inventat, amb DNI 12345678Z."
AssertEq ([string]$rA3b.autoritzacio.signant.nom) 'MARIA EXEMPLE PROVA' 'A3: davant del nom hi ha text en minuscules: nomes les MAJUSCULES'

Write-Host "`n--- A4. Noms brossa ---"
AssertEq ([string](_CtPersona 'N.I. 478981881<').nom) '' 'A4: "N.I. ..." no es cap nom'
AssertEq ([string](_CtPersona ('N.I. y N. I. F. n' + [char]0x00FA + 'mero 123 INTERVIENE en su calidad de APODERADO')).nom) '' 'A4: "... INTERVIENE en su calidad de ..." tampoc'
AssertEq ([string](_CtPersona ('/D' + [char]0x00BB + ' NIKOLOZ EXEMPLE')).nom) 'NIKOLOZ EXEMPLE' 'A4: el "/D>>" del principi fora'
AssertEq ([string](_CtPersona ('D./D' + [char]0x00F1 + 'a. MARIA EXEMPLE')).nom) 'MARIA EXEMPLE' 'A4: "D./Dna." fora'
AssertEq ([string](_CtPersona 'JOAN 2 PROVA').nom) '' 'A4: una persona amb xifres al nom no'
AssertEq ([string](_CtPersona 'BAR 2000 SL').nom) 'BAR 2000 SL' 'A4: ...una empresa si ("BAR 2000 SL")'

Write-Host "`n--- A5. Les etiquetes que quedaven ---"
$etA5 = 'Adre' + [char]0x00E7 + 'a postal (carrer, pla' + [char]0x00E7 + 'a, avinguda, n' + [char]0x00FA + 'mero, escala, pis, porta) Direcci' + [char]0x00F3 + 'n postal (calle, plaza, avenida, n' + [char]0x00FA + 'mero, escalera, piso, puerta)'
AssertEq (_CtEsLiniaEtiqueta $etA5) $true 'A5: "Adreca postal (... avinguda, escala ...) Direccion postal (...)" es etiqueta'
AssertEq ([string](_CtPersona $etA5).nom) '' 'A5: ...i no pot ser mai un nom'
AssertEq (_CtEsLiniaEtiqueta ('Mitj' + [char]0x00E0 + ' de notificaci' + [char]0x00F3 + ' Medio de notificaci' + [char]0x00F3 + 'n')) $true 'A5: el mitja de notificacio tambe'
AssertEq (_CtEsLiniaEtiqueta 'CARRER INVENTAT 1 BAIXOS') $false 'A5: una adreca de debo no es una etiqueta'
AssertEq (_CtEsLiniaEtiqueta ('Adre' + [char]0x00E7 + 'a postal (carrer, pla' + [char]0x00E7 + 'a, n' + [char]0x00FA + 'm., xamfra, pis, porta) Direcci' + [char]0x00F3 + 'n postal (calle, plaza, chaflan, piso)')) $true 'A5: amb alguna paraula que no es a la llista ("xamfra"), si gairebe tot es etiqueta, tambe'

Write-Host "`n--- B1. Els noms casen per un cognom, no per 'MARIA' ---"
AssertEq (_CtMateixNom 'MARIA DEL MAR' 'MARIA DEL MAR GARCIA LOPEZ') $false 'B1: noms de pila sols ("Maria del Mar") no fan la mateixa persona'
AssertEq (_CtMateixNom 'Maria Exemple' 'MARIA EXEMPLE PROVA') $true 'B1: amb un cognom, si'
AssertEq (_CtMateixNom 'SIEMENS SA' 'SIEMENS, S.A.') $true 'B1: el mateix nom d''empresa (una paraula) segueix casant'
$tecB1 = _CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'maria@tecniques.test'; noms = @('Maria del Mar'); empresa = @(); rol = 'tecnic' }) })
AssertEq (_CtMotiuTecnic (_CtPersona 'MARIA DEL MAR GARCIA LOPEZ') (_CtContext @{} @{} $tecB1 @())) '' 'B1: la representant legal no es la tecnica de la llista per compartir el nom de pila'

Write-Host "`n--- B2. El correu de la llista, una altra persona ---"
$tecB2 = _CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'tramits@gestoria-proves.test'; noms = @('Lluis Gestor Prova'); empresa = @('Gestoria Proves'); rol = 'gestoria' }) })
$ctxB2 = _CtContext @{} @{} $tecB2 @()
$aB2 = Get-ContactesActivitat '9701' @((& $nouDoc 'GIA 9701\etram-tramit.xml' 'etram' (_CtPersona 'PROVES B2 SL' 'B11111111') (_CtPersona 'Joan Admin Exemple' '11111111H' 'tramits@gestoria-proves.test') '2025-03-01')) @{ TITULAR = 'PROVES B2 SL'; NIF = 'B11111111' } $ctxB2
AssertEq "$($aB2.representant_legal.nom)|$($aB2.representant_legal.email)" 'Joan Admin Exemple|' 'B2: qui porta el correu de la gestoria conserva el seu paper (representant legal), sense aquell correu'
AssertEq (@($aB2.persones_autoritzades | ForEach-Object { $_.nom + '/' + $_.email + '/' + $_.rol }) -join ',') 'Lluis Gestor Prova/tramits@gestoria-proves.test/gestoria' 'B2: el correu es de la gestoria (persona autoritzada)'
AssertEq (_CtMotiuTecnic (_CtPersona 'Lluis Gestor Prova' '' 'tramits@gestoria-proves.test') $ctxB2) "es a la llista de t$([char]0x00E8)cnics coneguts" 'B2: el de la llista, amb el seu nom, si que es tecnic'
AssertEq (_CtMotiuTecnic (_CtPersona 'Joan Admin Exemple' '' 'tramits@gestoria-proves.test') $ctxB2) '' 'B2: una altra persona amb aquell correu no es tecnic pel correu'

Write-Host "`n--- B3. El domini del titular no es del tecnic ---"
$tecB3 = _CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'pere@enginyeria-inventada.test'; noms = @('Pere Tecnic Inventat'); empresa = @('Enginyeria Inventada'); rol = 'tecnic' }) })
$ctxB3 = _CtContext @{} @{} $tecB3 @()
$xlB3 = @{ TITULAR = 'INDUSTRIES EXEMPLARS SA'; NIF = 'A11111111'; EMAIL = 'laura@exemplars.test'; MOBIL = '688888888'; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
$aB3 = Get-ContactesActivitat '9801' @((& $nouDoc 'GIA 9801\etram-tramit.xml' 'etram' (_CtPersona 'INDUSTRIES EXEMPLARS SA' 'A11111111') (_CtPersona 'Pere Tecnic Inventat' '12345678Z' 'laura@exemplars.test' '' '688888888') '2025-04-01')) $xlB3 $ctxB3
$pB3 = @($aB3.persones_autoritzades)[0]
AssertEq "$($pB3.nom)|$($pB3.email)|$($pB3.mobil)" 'Pere Tecnic Inventat||' 'B3: el tecnic, sense el correu del domini de l''empresa ni el mobil que portava amb ell'
AssertEq (@($aB3.avisos | Where-Object { $_.tipus -eq 'es_el_tecnic' }).Count) 0 'B3: cap avis "es el tecnic" sobre el correu i el mobil de la treballadora de l''empresa'
AssertEq "$($aB3.titular.email)|$((@($aB3.dominis_titular)) -join ',')" 'laura@exemplars.test|exemplars.test' 'B3: el correu va al titular; el domini, pel nom de l''empresa'

Write-Host "`n--- B4. Titulars diferents: 3, o 2 amb un altre senyal ---"
$dB4 = { param([string]$g, [string]$nif, $rep, $int = $null) if ($null -eq $int) { $int = (_CtPersona ('TITULAR ' + $g + ' SL') $nif) }; @(@{ ruta = ('GIA ' + $g + '\a.pdf'); modificat = '2025-01-01T00:00:00Z'; extret = @{ interessat = $int; representant = $rep } }) }
$ctxB4 = _CtContext @{ '1' = (& $dB4 '1' 'B10000001' $null (_CtPersona 'U SL' 'B10000001' 'compartit@correu.test')); '2' = (& $dB4 '2' 'B20000002' $null (_CtPersona 'DOS SL' 'B20000002' 'compartit@correu.test')) } @{} @{ Emails = @{}; Noms = @(); Dominis = @{} } @()
AssertEq (_CtMotiuTecnic @{ email = 'compartit@correu.test' } $ctxB4) '' 'B4: un correu a 2 titulars, sense cap altre senyal, no fa tecnic'
$ctxB4b = _CtContext @{ '1' = (& $dB4 '1' 'B10000001' (_CtPersona 'X' '' 'pres@correu.test')); '2' = (& $dB4 '2' 'B20000002' (_CtPersona 'X' '' 'pres@correu.test')) } @{} @{ Emails = @{}; Noms = @(); Dominis = @{} } @()
Assert ((_CtMotiuTecnic @{ email = 'pres@correu.test' } $ctxB4b) -ne '') 'B4: 2 titulars presentant com a representant als dos, si'
$xlB4 = @{ '1' = @{ NIF = 'B10000001'; REP_NIF = '44444444A' }; '2' = @{ NIF = 'B20000002'; REP_NIF = '44444444A' }; '3' = @{ NIF = 'B30000003' } }
$docsB4 = @{ '1' = (& $dB4 '1' 'B10000001' (_CtPersona 'Y' '' 'germanes@correu.test')); '2' = (& $dB4 '2' 'B20000002' (_CtPersona 'Y' '' 'germanes@correu.test')); '3' = (& $dB4 '3' 'B30000003' (_CtPersona 'Y' '' 'germanes@correu.test')) }
$ctxB4c = _CtContext $docsB4 $xlB4 @{ Emails = @{}; Noms = @(); Dominis = @{} } @()
AssertEq $ctxB4c.EmailTitulars['germanes@correu.test'].Count 2 'B4: dues empreses amb el mateix representant legal a l''Excel compten com UN titular'

Write-Host "`n--- B5. Correus de gestoria pel nom, i el domini mal escrit ---"
AssertEq (_CtCorreuGestoriaPelNom 'asesoriarovira2000@hotmail.test' 'BAR EXEMPLE SL') $true 'B5: "asesoria..." no es el correu del titular'
AssertEq (_CtCorreuGestoriaPelNom 'info@asesoria-exemple.test' 'ASESORIA EXEMPLE SL') $false 'B5: ...tret que el titular sigui l''assessoria'
$aB5 = Get-ContactesActivitat '9901' @((& $nouDoc 'GIA 9901\Instancia_generica.pdf' 'instancia' (_CtPersona 'BAR EXEMPLE SL' 'B99999999' 'asesoriarovira2000@hotmail.test') $null '2025-01-01')) @{ TITULAR = 'BAR EXEMPLE SL'; NIF = 'B99999999' } $ctx2
AssertEq ([string]$aB5.titular.email) '' 'B5: el titular no es queda el correu de l''assessoria'
$tecB5 = _CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'info@inargesa.test'; noms = @('Gestor Inventat Prova'); empresa = @(); rol = 'gestoria' }) })
Assert ((_CtMotiuTecnic @{ email = 'pep@inaargesa.test' } (_CtContext @{} @{} $tecB5 @())).Contains('assembla')) 'B5: un domini a 2 lletres del d''un tecnic de la llista es seu'
AssertEq (_CtMotiuTecnic @{ email = 'pep@gmail.com' } (_CtContext @{} @{} (_CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'tecnic@gmail.com'; noms = @(); empresa = @(); rol = '' }) })) @())) '' 'B5: el domini d''un proveidor corrent (gmail) no compta'

Write-Host "`n--- C. Els documents sense GIA ---"
AssertEq (_CtGiaMajoritari @{ '685' = 3; '999' = 1 }) '685' 'C1: el GIA de mes de la meitat dels informes de la carpeta'
AssertEq (_CtGiaMajoritari @{ '1' = 1; '2' = 1 }) '' 'C1: meitat i meitat, cap'
AssertEq (_CtGiaDocument 'Exp RCOBR - EXEMPLE\Firmadoc\Autoritzacio.pdf' @{ gia = '' } @{ 'Exp RCOBR - EXEMPLE' = @{ '685' = 3; '999' = 1 } } { param($g) $true } @{}) '685' 'C1: un document de Firmadoc\ amb el GIA dels informes de la carpeta d''activitat'
Assert ((_CtMotiuSenseGia 'Exp\Firmadoc\a.pdf' @{ 'Exp' = @{ '1' = 1; '2' = 1 } }).Contains('GIA 1 (1), GIA 2 (1)')) 'C1: per a la validacio es diu per que no en te'

$ctT2 = Join-Path ([System.IO.Path]::GetTempPath()) ('contactes2-' + [guid]::NewGuid().ToString('N'))
$ctI2 = Join-Path $ctT2 'Informes'
$vells2 = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
try {
    foreach ($d in @((Join-Path $ctI2 'Carpeta sense res'), (Join-Path $ctT2 'local'), (Join-Path $ctT2 'act'), (Join-Path $ctT2 'app'))) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $LocalActivitatsDir = Join-Path $ctT2 'local'; $ActivitatsDir = Join-Path $ctT2 'act'; $env:LOCALAPPDATA = Join-Path $ctT2 'app'; $InformesDir = $ctI2
    _CtOblidaMemo
    _CtProvaPdf (Join-Path (Join-Path $ctI2 'Carpeta sense res') 'Instancia_generica.pdf') @('DADES DE LA PERSONA O ENTITAT INTERESSADA', 'Nom o rao social: PER NIF SL', 'DNI/NIF/NIE: B70000007', 'Correu: pernif@proves.test')
    $cacheC2 = [pscustomobject]@{ ById = @{}; Contactes = @{ '9902' = @{ TITULAR = 'PER NIF SL'; NIF = 'B70000007' } } }
    $rC2 = Invoke-ContactesEscaneig $ctI2 @() $cacheC2 @{}
    $dC2 = (Read-ContactesDb).documents.'Carpeta sense res\Instancia_generica.pdf'
    AssertEq "$($dC2.gia)|$($dC2.gia_via)|$($rC2.Estadistiques.PerNif)|$(@($rC2.SenseGia).Count)" '9902|nif|1|0' 'C2: sense carpeta ni informes, el NIF de l''interessat (d''una sola activitat de l''Excel)'
    AssertEq "$($rC2.Activitats['9902'].titular.email)|$($rC2.Activitats['9902'].titular.confianca)" 'pernif@proves.test|probable' 'C2: ...i les seves dades, "probable"'

    Write-Host "`n--- E1. Les bases no es tornen a fer de zero en silenci ---"
    $pathC = Get-ContactesDbPath
    $bakC = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'contactes-db_*.bak' -File)
    AssertEq $bakC.Count 0 'E1: la primera vegada no hi ha res a copiar'
    [void](Invoke-ContactesEscaneig $ctI2 @() $cacheC2 @{})
    AssertEq @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'contactes-db_*.bak' -File).Count 1 'E1: abans de sobreescriure contactes-db.json, una copia contactes-db_AAAAMMDD-HHMM.bak'
    [System.IO.File]::WriteAllText($pathC, '{ "versio": "trencat", ')
    $errC = ''
    try { [void](Invoke-ContactesEscaneig $ctI2 @() $cacheC2 @{}) } catch { $errC = [string]$_.Exception.Message }
    Assert ($errC -like "*No s'ha pogut llegir*") ('E1: contactes-db.json que no es pot llegir: s''atura (' + $errC + ')')
    AssertEq ([System.IO.File]::ReadAllText($pathC)) '{ "versio": "trencat", ' 'E1: ...i no l''ha tocat'
    $pathI = Get-InformesDbPath
    [System.IO.File]::WriteAllText($pathI, '{ "activitats": [ ')
    $rI = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($rI.Ok)|$([System.IO.File]::ReadAllText($pathI))" 'False|{ "activitats": [ ' 'E1: informes-db.json que no es pot llegir: no se''n fa una de nova, s''avisa i no es toca'
    Assert ([string]$rI.Error).Contains('.bak') 'E1: l''avis diu on son les copies'
    Remove-Item -LiteralPath $pathI -Force
    $rI2 = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($rI2.Ok)|$(@(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'informes-db_*.bak' -File).Count)" 'True|0' 'E1: sense base, se''n fa una (i no hi ha res a copiar)'
    $rI3 = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq "$($rI3.Ok)|$(@(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'informes-db_*.bak' -File).Count)" 'True|1' 'E1: la segona, abans de sobreescriure-la, la copia'
    $fx = Join-Path $ctT2 'prova.json'
    [System.IO.File]::WriteAllText($fx, '{}')
    foreach ($k in 1..7) { [System.IO.File]::WriteAllText((Join-Path $ctT2 ('prova_2026010' + $k + '-1200.bak')), 'x') }
    [void](Backup-JsonFile $fx 5)
    AssertEq @(Get-ChildItem -LiteralPath $ctT2 -Filter 'prova_*.bak' -File).Count 5 'E1: nomes es guarden les 5 copies mes recents'
    AssertEq (Read-JsonFileStrict (Join-Path $ctT2 'no-hi-es.json')) $null 'Read-JsonFileStrict: un fitxer que no hi es -> $null (no es cap error)'
} catch {
    Assert $false ('bloc C2/E1: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    $InformesDir = $vells2.Inf; $LocalActivitatsDir = $vells2.Loc; $ActivitatsDir = $vells2.Act; $env:LOCALAPPDATA = $vells2.App
    _CtOblidaMemo
    Remove-Item -LiteralPath $ctT2 -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- E2. ValidarClassificacio: el pre/post de llicencia contra 'Favorable' ---"
AssertEq (_BreuEquivalentRef 'Favorable' $Script:EstatFavorablePre) 'llicencia' 'E2: Favorable de la referencia i pre-llicencia d''ara: equivalents (a part)'
AssertEq (_BreuEquivalentRef 'Favorable' $Script:EstatFavorablePost) 'llicencia' 'E2: ...i post-llicencia'
AssertEq (_BreuEquivalentRef 'Favorable' 'Requeriment') 'diferent' 'E2: un Requeriment segueix sent diferent'
AssertEq (_BreuEquivalentRef 'Requeriment' 'Requeriment') 'igual' 'E2: igual'
$srcVc = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'ValidarClassificacio.ps1')
Assert ($srcVc.Contains('$eqBreu = _BreuEquivalentRef') -and $srcVc.Contains("-ne 'llicencia') { `$o.conclusio_breu")) 'E2: ValidarClassificacio compara amb _BreuEquivalentRef (informe i estat)'
$srcIe = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'InformesEscaneig.ps1')
Assert ($srcIe.Contains('$prevDb = Read-JsonFileStrict $outPath') -and $srcIe.Contains('[void](Backup-JsonFile $outPath)')) 'E1: Actualitzar base llegeix la base anterior amb Read-JsonFileStrict i en fa copia abans de desar'

AssertEq $Script:ContactesVersio '2026-10-15.2' 'la versio puja (es tornen a llegir tots els documents)'
