# Proves automatiques de les funcions PURES de Precintades.ps1 (mapa public
# d'activitats precintades).
#
# NO prova la part d'Excel (COM) ni l'escriptura del JSON al disc: aixo nomes
# funciona a Windows amb Excel. Aqui es valida la logica: deteccio del camp
# "PRECINTE ACTIVITAT?", regla del valor que comenca per "SI", deteccio
# dinamica de columnes/parells Camp Info i construccio de l'objecte del JSON.
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File suport/tests/run-tests-precintades.ps1

$ErrorActionPreference = 'Stop'
$env:PRECINTADES_TEST = '1'   # mode headless: nomes defineix funcions

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'rutes' 'Precintades.ps1')
. $scriptPath

$script:pass = 0
$script:fail = 0
function Assert($cond, $name) {
    if ($cond) { $script:pass++; Write-Host "  OK   $name" -ForegroundColor Green }
    else       { $script:fail++; Write-Host "  FAIL $name" -ForegroundColor Red }
}
function AssertEq($actual, $expected, $name) {
    Assert ([string]$actual -eq [string]$expected) "$name (esperat '$expected', obtingut '$actual')"
}

Write-Host "`n--- Test-IsPrecintada (camp + regla del valor 'SI') ---"
Assert  (Test-IsPrecintada 'PRECINTE ACTIVITAT?' 'SI, PRECINTAT 24/03/2026')       'valor comenca per SI, ...'
Assert  (Test-IsPrecintada 'PRECINTE ACTIVITAT?' 'SI')                             'valor es nomes SI'
Assert  (Test-IsPrecintada 'PRECINTE ACTIVITAT?' 'si, cessada activitat')          'insensible a majuscules'
Assert  (Test-IsPrecintada 'precinte activitat?' 'SI, decret dos mesos')           'nom insensible a majuscules'
Assert  (-not (Test-IsPrecintada 'PRECINTE ACTIVITAT?' '28/06/2026 ES PRECINTE'))  'valor que NO comenca per SI -> fals'
Assert  (-not (Test-IsPrecintada 'PRECINTE ACTIVITAT?' '2026/3213 PROPOSTA'))      'valor amb numero al davant -> fals'
Assert  (-not (Test-IsPrecintada 'PRECINTE ACTIVITAT?' 'SITUACIO IRREGULAR'))      'SI ha de ser paraula, no prefix (SITUACIO)'
Assert  (-not (Test-IsPrecintada 'PRECINTE ACTIVITAT?' ''))                        'valor buit -> fals'
Assert  (-not (Test-IsPrecintada 'DENUNCIA?' 'SI, la que sigui'))                  'un altre camp amb SI -> fals'
Assert  (-not (Test-IsPrecintada 'REQUERIT PER DECRET?' 'SI'))                     'REQUERIT PER DECRET amb SI -> fals'

Write-Host "`n--- Find-HeaderColumn (localitza columnes per nom, insensible a accents) ---"
# Construim la capcalera accentuada amb codi de caracter (U+00FA = 'ú') per NO
# dependre de la codificacio del fitxer: el Windows PowerShell 5.1 llegeix els
# .ps1 sense BOM com a ANSI i corromp un literal 'Número' escrit directament.
# Aixi el test comprova de veritat que un nom ASCII ('Emp. Numero') encaixa amb
# una capcalera accentuada ('Emp. Número'), com passa amb les dades reals d'Excel.
$numHeader = "Emp. N$([char]0x00FA)mero"   # "Emp. Número"
$hdr = @('ID Activitat','UTM X','UTM Y','Emp. Tipus via','Emp. Carrer',$numHeader,'Activitat principal')
AssertEq (Find-HeaderColumn $hdr 'ID Activitat')        1 'ID Activitat -> col 1'
AssertEq (Find-HeaderColumn $hdr 'UTM Y')               3 'UTM Y -> col 3'
AssertEq (Find-HeaderColumn $hdr 'Emp. Numero')         6 'accent ignorat (Numero ASCII == capcalera accentuada) -> col 6'
AssertEq (Find-HeaderColumn $hdr 'Activitat principal') 7 'Activitat principal -> col 7'
AssertEq (Find-HeaderColumn $hdr 'No existeix')         0 'columna inexistent -> 0'

Write-Host "`n--- Get-CampInfoPairs (deteccio dinamica dels parells Nom/Valor) ---"
# Capcalera amb 3 parells Camp Info intercalats amb altres columnes (com l'Excel real).
$hdr2 = @(
    'Dada tecnica - Nom','Dada tecnica - Valor','',                  # 1..3 (soroll)
    'Camp Info 1 - Nom','Camp Info 1 - Valor','Camp Info 1 - Unitat', # 4..6
    'Camp Info 2 - Nom','Camp Info 2 - Valor','Camp Info 2 - Unitat', # 7..9
    'Camp Info 3 - Nom','Camp Info 3 - Valor','Camp Info 3 - Unitat'  # 10..12
)
$pairs = Get-CampInfoPairs $hdr2
AssertEq @($pairs).Count 3 'detecta 3 parells Camp Info'
AssertEq $pairs[0].NomCol   4 'parell 1: Nom a col 4'
AssertEq $pairs[0].ValorCol 5 'parell 1: Valor a col 5'
AssertEq $pairs[2].NomCol   10 'parell 3: Nom a col 10'
AssertEq $pairs[2].ValorCol 11 'parell 3: Valor a col 11'
# Robustesa: n'hi pot haver mes de 3.
$hdr3 = @('Camp Info 1 - Nom','Camp Info 1 - Valor','Camp Info 7 - Nom','Camp Info 7 - Valor')
$pairs3 = Get-CampInfoPairs $hdr3
AssertEq @($pairs3).Count 2 'detecta parells encara que la numeracio no sigui consecutiva'
AssertEq $pairs3[1].NomCol 3 'Camp Info 7 - Nom a col 3'

Write-Host "`n--- Build-PrecintadesObject (estructura del JSON) ---"
$recs = @(
    [pscustomobject]@{ Id='1369'; ActivitatPrincipal='MAGATZEM I GARATGE'; Adreca='PG FERROCARRILS CATALANS 24'; Lat=41.3515651; Lon=2.0689571 }
    [pscustomobject]@{ Id='776';  ActivitatPrincipal='BAR RESTAURANT';     Adreca='C RAMONEDA 34';               Lat=41.3612199; Lon=2.0775688 }
)
$obj = Build-PrecintadesObject $recs '2026-07-16 ACTIVITATS.xls'
AssertEq $obj.Comptador 2 'Comptador = nombre d activitats'
AssertEq $obj.Font '2026-07-16 ACTIVITATS.xls' 'Font conserva el nom del fitxer'
AssertEq @($obj.Activitats).Count 2 'Activitats amb tots els registres'
AssertEq $obj.Activitats[0].id '1369' 'primer registre: id'
AssertEq $obj.Activitats[0].activitat 'MAGATZEM I GARATGE' 'primer registre: activitat'
AssertEq $obj.Activitats[0].adreca 'PG FERROCARRILS CATALANS 24' 'primer registre: adreca'
Assert  ([math]::Abs([double]$obj.Activitats[1].lat - 41.3612199) -lt 1e-6) 'segon registre: lat numeric'
# El JSON serialitzat ha de contenir els camps esperats i cap nota interna.
$json = $obj | ConvertTo-Json -Depth 6
Assert ($json -match '"Activitats"') 'JSON conte Activitats'
Assert ($json -match 'MAGATZEM I GARATGE') 'JSON conte l activitat principal'

Write-Host "`n--- SourceDate / _ParsePrecintadesDate / Test-ShouldUpdatePrecintades ---"
AssertEq $obj.SourceDate '2026-07-16' 'Build-PrecintadesObject afegeix SourceDate del nom'
$objSD = [pscustomobject]@{ SourceDate = '2026-07-16'; Font = '2026-07-16 ACTIVITATS.xls' }
AssertEq ((_ParsePrecintadesDate $objSD).ToString('yyyy-MM-dd')) '2026-07-16' '_ParsePrecintadesDate llegeix SourceDate'
$objFont = [pscustomobject]@{ Font = '2026-05-29 ACTIVITATS.xlsx' }
AssertEq ((_ParsePrecintadesDate $objFont).ToString('yyyy-MM-dd')) '2026-05-29' '_ParsePrecintadesDate fallback al nom Font'
Assert ((_ParsePrecintadesDate ([pscustomobject]@{ x=1 })) -eq [datetime]::MinValue) '_ParsePrecintadesDate sense data -> MinValue'
Assert ((_ParsePrecintadesDate $null) -eq [datetime]::MinValue) '_ParsePrecintadesDate null -> MinValue'

$dPub = [datetime]'2026-07-16'
Assert (-not (Test-ShouldUpdatePrecintades ([datetime]'2026-07-16') $dPub)) 'mateixa data -> NO actualitza'
Assert (-not (Test-ShouldUpdatePrecintades ([datetime]'2026-07-01') $dPub)) 'base trobada mes antiga -> NO actualitza'
Assert (Test-ShouldUpdatePrecintades ([datetime]'2026-07-20') $dPub)         'base trobada mes nova -> actualitza'
Assert (Test-ShouldUpdatePrecintades ([datetime]'2026-07-16') ([datetime]::MinValue)) 'sense mapa previ -> actualitza'
Assert (Test-ShouldUpdatePrecintades ([datetime]::MinValue) $dPub)           'sense data local fiable -> actualitza (per seguretat)'

$summaryColor = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "`n========================================"
Write-Host ("RESULTAT: {0} OK, {1} FAIL" -f $script:pass, $script:fail) -ForegroundColor $summaryColor
Write-Host "========================================"
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
