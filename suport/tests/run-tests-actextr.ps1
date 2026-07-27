# Proves automatiques de les funcions PURES del mode ACT_EXTR (ActExtr.ps1).
#
# NO prova la part de Word (COM) ni les finestres (WinForms): aixo nomes es pot
# provar a Windows amb Office. Aqui es valida la LOGICA del Decret 112/2010
# (aplicabilitat, calcul de valors), el parseig de plantilles (blocs keyed),
# la inclusio de blocs (requeriment/favorable) i el registre local.
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File suport/tests/run-tests-actextr.ps1
#
# Carrega Motor.ps1 en mode "headless" (GENINFORME_TEST=1), que al seu torn fa
# dot-source d'ActExtr.ps1.

$ErrorActionPreference = 'Stop'
$env:GENINFORME_TEST = '1'
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }

# Registre en una carpeta temporal aillada (NO el del repo).
$script:ActExtrRegistryDir = Join-Path ([System.IO.Path]::GetTempPath()) ("actextr-test-" + [guid]::NewGuid().ToString('N'))

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Motor.ps1'
. $scriptPath   # dot-source: defineix les funcions del motor

. (Join-Path $PSScriptRoot 'TestLib.ps1')   # Assert / AssertEq / Write-TestSummary

Write-Host "`n--- Lookups del Decret (valors calculats) ---"
# Mostra de l'Excel: aforament 34719 -> 35 vigilants, 37 controladors,
# 276 lavabos / 828 cabines, polissa RC 3.800.000 EUR.
AssertEq (Get-ActExtrVigilants 34719)    35  'vigilants(34719)=35'
AssertEq (Get-ActExtrVigilants 500)       0  'vigilants(500)=0'
AssertEq (Get-ActExtrVigilants 501)       1  'vigilants(501)=1'
AssertEq (Get-ActExtrVigilants 1001)      2  'vigilants(1001)=2'
AssertEq (Get-ActExtrControladors 34719) 37  'controladors(34719)=37'
AssertEq (Get-ActExtrControladors 149)    0  'controladors(149)=0'
AssertEq (Get-ActExtrControladors 150)    2  'controladors(150)=2'
AssertEq (Get-ActExtrControladors 501)    3  'controladors(501)=3'
$h = Get-ActExtrHigiene 34719
AssertEq $h.Lavabos 276 'lavabos(34719)=276'
AssertEq $h.Cabines 828 'cabines(34719)=828'
$h0 = Get-ActExtrHigiene 0
AssertEq $h0.Lavabos 1 'lavabos(0)=1'
AssertEq $h0.Cabines 2 'cabines(0)=2'
AssertEq (Get-ActExtrPolissaRC 34719 $false $false) 3800000 'RC(34719)=3.800.000'
AssertEq (Get-ActExtrPolissaRC 500 $false $false)    750000 'RC(500)=750.000'
AssertEq (Get-ActExtrPolissaRC 100 $false $false)    300000 'RC(100)=300.000'
AssertEq (Get-ActExtrPolissaRC 500 $true $false)     937500 'RC(500) parcial *1,25'
AssertEq (Get-ActExtrPolissaRC 500 $false $true)     975000 'RC(500) total *1,30'
AssertEq (Get-ActExtrPolissaRC 500 $true $true)      975000 'RC(500) agafa el factor mes alt'

Write-Host "`n--- _ActExtrThousands ---"
AssertEq (_ActExtrThousands 3800000) '3.800.000' 'milers amb punt'
AssertEq (_ActExtrThousands 900)     '900'       'sense separador'
AssertEq (_ActExtrThousands 0)       '0'         'zero'

Write-Host "`n--- Build-ActExtrDecret + Get-ActExtrComputed ---"
$ans = @{ Aforament='34719'; Incendis='Si'; Mobilitat='Si'; ControlAccessos='Si'; PauCatalunya='No'; PauLocal='Si'; EstablimentDotat='No'; ParcialSotaRasant='No'; TotalSotaRasant='No' }
$decret = Build-ActExtrDecret $ans
AssertEq $decret.Aforament 34719 'decret aforament parsejat'
AssertEq $decret.Incendis 'Si'   'decret incendis Si'
$comp = Get-ActExtrComputed $decret
AssertEq $comp.VIGILANTS 35 'computed vigilants'
AssertEq $comp.CONTROLADORS 37 'computed controladors'
AssertEq $comp.RC_IMPORT '3.800.000' 'computed RC formatat'
AssertEq $comp.PAU_ORGAN_KEY 'LOCAL' 'PAU organ local (CatB=Si)'
Assert  ($comp.PAU_OBLIGAT) 'PAU obligat'

# Prioritat Catalunya si esta a tots dos catalegs.
$d2 = Build-ActExtrDecret @{ Aforament='100'; PauCatalunya='Si'; PauLocal='Si' }
AssertEq (Get-ActExtrComputed $d2).PAU_ORGAN_KEY 'CAT' 'PAU prioritza Catalunya'
# Cap dels dos -> no obligat.
$d3 = Build-ActExtrDecret @{ Aforament='100'; PauCatalunya='No'; PauLocal='No' }
Assert (-not (Get-ActExtrComputed $d3).PAU_OBLIGAT) 'PAU no obligat si cap cataleg'

Write-Host "`n--- Aplicabilitat + motiu ('aplica i per que') ---"
$ap = Get-ActExtrPointApplicability 'INCENDIS' $decret $comp
Assert ($ap.Applies) 'INCENDIS aplica (Incendis=Si)'
Assert ($ap.Reason -match 'Art\. 23') 'INCENDIS motiu cita Art. 23'
$apN = Get-ActExtrPointApplicability 'VIGILANTS' (Build-ActExtrDecret @{ Aforament='200' }) (Get-ActExtrComputed (Build-ActExtrDecret @{ Aforament='200' }))
Assert (-not $apN.Applies) 'VIGILANTS no aplica si aforament<501'
$apRC = Get-ActExtrPointApplicability 'RC' $decret $comp
Assert ($apRC.Applies) 'RC sempre aplica'
$apCtrlNo = Get-ActExtrPointApplicability 'CONTROLADORS' (Build-ActExtrDecret @{ Aforament='34719'; ControlAccessos='No' }) $comp
Assert (-not $apCtrlNo.Applies) 'CONTROLADORS no aplica si ControlAccessos=No'

Write-Host "`n--- Status + deficiencies ---"
$delivered = @{ INCENDIS=$true }   # nomes incendis lliurat
$status = Get-ActExtrStatus $decret $comp $delivered
$incStatus = $status | Where-Object { $_.Key -eq 'INCENDIS' }
Assert ($incStatus.Delivered) 'status: INCENDIS lliurat'
$defs = Get-ActExtrDeficiencies $decret $comp $delivered
Assert ($defs -notcontains 'INCENDIS') 'deficiencies: INCENDIS no hi es (lliurat)'
Assert ($defs -contains 'RC') 'deficiencies: RC hi es (no lliurat)'
$estat = Get-ActExtrActivityEstat $decret $comp $delivered
AssertEq $estat 'pendent' 'estat pendent si hi ha deficiencies'

# Tot lliurat -> tancat.
$allKeys = $script:ActExtrPoints | ForEach-Object { $_.Key }
$allDelivered = @{}; foreach ($k in $allKeys) { $allDelivered[$k] = $true }
AssertEq (Get-ActExtrActivityEstat $decret $comp $allDelivered) 'tancat' 'estat tancat si tot lliurat'

Write-Host "`n--- Resolve-ActExtrTokens ---"
AssertEq (Resolve-ActExtrTokens 'cal {{VIGILANTS}} vigilants i {{CONTROLADORS}} controladors' $comp) 'cal 35 vigilants i 37 controladors' 'substitueix tokens numerics'
AssertEq (Resolve-ActExtrTokens '{{LAVABOS}} lavabos i {{CABINES}} cabines' (Get-ActExtrComputed (Build-ActExtrDecret @{ Aforament='34719' }))) '276 lavabos i 828 cabines' 'tokens higiene'

Write-Host "`n--- _ParseActExtrMarker (Titol 2 -> clau + tipus) ---"
$mk0 = _ParseActExtrMarker '[[INCENDIS]] Incendis'
AssertEq $mk0.Key 'INCENDIS' 'marcador: clau extreta'
AssertEq $mk0.Kind 'item' 'marcador sense :: -> item'
AssertEq (_ParseActExtrMarker '[[REQ_INTRO]] ::TEXT:: Introduccio').Kind 'text' '::TEXT:: -> text'
AssertEq (_ParseActExtrMarker '[[MEMORIA_A]] ::CHILD:: a) ...').Kind 'child' '::CHILD:: -> child'
AssertEq (_ParseActExtrMarker '[[N]] ::NOTE:: nota').Kind 'note' '::NOTE:: -> note'
AssertEq (_ParseActExtrMarker '[[L]] ::LABEL:: rot').Kind 'label' '::LABEL:: -> label'
AssertEq (_ParseActExtrMarker '[[H]] ::HEADER:: cap').Kind 'header' '::HEADER:: -> header'
AssertEq (_ParseActExtrMarker '[[C]] ::CONC:: conc').Kind 'conc' '::CONC:: -> conc'
Assert ($null -eq (_ParseActExtrMarker 'Titol visual sense clau')) 'titol sense [[KEY]] -> $null'

Write-Host "`n--- Build-ActExtrBlocks (mateix format que REQ1) ---"
$recs = @(
    @{ Text='DEFICIENCIES'; Style='h1' }                                  # seccio: ignorada
    @{ Text='[[REQ_INTRO]] ::TEXT:: Introduccio'; Style='h2' }
    @{ Text='Intro del requeriment:'; Style='normal' }
    @{ Text='[[INCENDIS]] Incendis'; Style='h2' }
    @{ Text='Incendis. Text.'; Style='normal' }
    @{ Text='https://exemple.cat/incendis'; Style='url' }
    @{ Text='[[MEMORIA_A]] ::CHILD:: a)'; Style='h2' }
    @{ Text='Identificacio.'; Style='normal' }
)
$blocks = Build-ActExtrBlocks $recs
AssertEq $blocks.Count 3 'blocs: REQ_INTRO, INCENDIS, MEMORIA_A (la seccio H1 no obre bloc)'
AssertEq $blocks[0].Key 'REQ_INTRO' 'primer bloc REQ_INTRO'
AssertEq $blocks[0].Kind 'text' 'REQ_INTRO es de tipus text'
AssertEq $blocks[0].Contents[0].Text 'Intro del requeriment:' 'contingut del bloc (Normal)'
AssertEq $blocks[1].Key 'INCENDIS' 'segon bloc INCENDIS'
AssertEq $blocks[1].Kind 'item' 'INCENDIS es item'
AssertEq $blocks[1].Contents.Count 2 'INCENDIS te text + url'
Assert ($blocks[1].Contents[1].IsUrl) 'segona linia es url'
AssertEq $blocks[2].Kind 'child' 'MEMORIA_A es sub-apartat (child)'
# Seccions: tots aquests blocs son sota la mateixa H1 -> mateixa Section.
AssertEq $blocks[0].Section 1 'primer bloc a la seccio 1 (primera H1)'
AssertEq $blocks[1].Section 1 'INCENDIS a la mateixa seccio'
# Nova H1 -> nova seccio.
$recs2 = @(
    @{ Text='SEC A'; Style='h1' }
    @{ Text='[[A]] item'; Style='h2' }
    @{ Text='text a'; Style='normal' }
    @{ Text='SEC B'; Style='h1' }
    @{ Text='[[B]] item'; Style='h2' }
    @{ Text='text b'; Style='normal' }
)
$blk2 = Build-ActExtrBlocks $recs2
AssertEq $blk2[0].Section 1 'bloc A a seccio 1'
AssertEq $blk2[1].Section 2 'bloc B a seccio 2 (nova H1)'

Write-Host "`n--- Test-ActExtrIncludeBlock (requeriment) ---"
$statusByKey = @{}; foreach ($s in $status) { $statusByKey[$s.Key] = $s }
$defKeys = Get-ActExtrDeficiencies $decret $comp $delivered
$ctxReq = @{ Decret=$decret; Computed=$comp; Delivered=$delivered; StatusByKey=$statusByKey; DefKeys=$defKeys }
Assert (Test-ActExtrIncludeBlock 'REQ_INTRO' 'req' $ctxReq) 'req: REQ_INTRO inclos (hi ha deficiencies)'
Assert (-not (Test-ActExtrIncludeBlock 'INCENDIS' 'req' $ctxReq)) 'req: INCENDIS NO inclos (lliurat)'
Assert (Test-ActExtrIncludeBlock 'RC' 'req' $ctxReq) 'req: RC inclos (pendent)'
Assert (Test-ActExtrIncludeBlock 'PAU_LOCAL' 'req' $ctxReq) 'req: PAU_LOCAL inclos (CatB, pendent)'
Assert (-not (Test-ActExtrIncludeBlock 'PAU_CAT' 'req' $ctxReq)) 'req: PAU_CAT NO inclos (organ local)'
Assert (-not (Test-ActExtrIncludeBlock 'MOBILITAT' 'req' $ctxReq)) 'req: MOBILITAT mai al requeriment'
Assert (Test-ActExtrIncludeBlock 'MEMORIA_HEADER' 'req' $ctxReq) 'req: capcalera memoria (hi ha subpunts pendents)'
Assert (-not (Test-ActExtrIncludeBlock 'FAV_NORMATIVA' 'req' $ctxReq)) 'req: blocs FAV_* nomes al favorable'
Assert (Test-ActExtrIncludeBlock 'REQ_CLOSING' 'req' $ctxReq) 'req: bloc de tancament (Ho poso.../Cornella) inclos'

Write-Host "`n--- Test-ActExtrIncludeBlock (favorable) ---"
$ctxFav = @{ Decret=$decret; Computed=$comp; Delivered=$allDelivered; StatusByKey=$statusByKey; DefKeys=@() }
Assert (Test-ActExtrIncludeBlock 'FAV_INTRO' 'fav' $ctxFav) 'fav: FAV_INTRO inclos'
Assert (Test-ActExtrIncludeBlock 'FAV_NORMATIVA' 'fav' $ctxFav) 'fav: normativa fixa inclosa'
Assert (Test-ActExtrIncludeBlock 'FAV_CLOSING' 'fav' $ctxFav) 'fav: bloc final fix inclos'
Assert (Test-ActExtrIncludeBlock 'MOBILITAT' 'fav' $ctxFav) 'fav: MOBILITAT inclosa (Mobilitat=Si)'
Assert (Test-ActExtrIncludeBlock 'PAU_LOCAL' 'fav' $ctxFav) 'fav: PAU_LOCAL inclos'
Assert (-not (Test-ActExtrIncludeBlock 'PAU_CAT' 'fav' $ctxFav)) 'fav: PAU_CAT NO inclos (organ local)'
# Assistencia sanitaria: bloc unic ASSIST (el text el resol {{ASSISTENCIA}}).
Assert (Test-ActExtrIncludeBlock 'ASSIST' 'fav' $ctxFav) 'fav: bloc ASSIST inclos (sempre)'
Assert (-not (Test-ActExtrIncludeBlock 'ASSIST' 'req' $ctxFav)) 'fav: ASSIST nomes al favorable'
Assert (-not (Test-ActExtrIncludeBlock 'MEMORIA_A' 'fav' $ctxFav)) 'fav: memoria no apareix al favorable'
Assert (-not (Test-ActExtrIncludeBlock 'REQ_CLOSING' 'fav' $ctxFav)) 'fav: REQ_CLOSING nomes al requeriment'

Write-Host "`n--- Lasers (marcatge + variants req/fav) ---"
# Sense lasers ($ans no te HiHaLasers -> 'No'): al requeriment el punt LASERS
# no aplica; al favorable surt la PROHIBICIO (NO_LASERS), no l'autoritzacio.
Assert (-not (Get-ActExtrPointApplicability 'LASERS' $decret $comp).Applies) 'LASERS no aplica sense lasers'
Assert (-not (Test-ActExtrIncludeBlock 'LASERS' 'req' $ctxReq)) 'req: LASERS no inclos (sense lasers)'
Assert (-not (Test-ActExtrIncludeBlock 'FAV_LASERS' 'fav' $ctxFav)) 'fav: FAV_LASERS no inclos (sense lasers)'
Assert (Test-ActExtrIncludeBlock 'NO_LASERS' 'fav' $ctxFav) 'fav: NO_LASERS inclos (prohibicio)'
# Amb lasers: aplica al requeriment (pendent) i surt l'AUTORITZACIO al favorable.
$ansL = @{ Aforament='34719'; PauLocal='Si'; HiHaLasers='Si' }
$decL = Build-ActExtrDecret $ansL
$compL = Get-ActExtrComputed $decL
Assert ($compL.HAS_LASERS) 'computed HAS_LASERS quan HiHaLasers=Si'
Assert (Get-ActExtrPointApplicability 'LASERS' $decL $compL).Applies 'LASERS aplica amb lasers'
$stL = Get-ActExtrStatus $decL $compL @{}
$sbkL = @{}; foreach ($s in $stL) { $sbkL[$s.Key] = $s }
$ctxL = @{ Decret=$decL; Computed=$compL; Delivered=@{}; StatusByKey=$sbkL; DefKeys=(Get-ActExtrDeficiencies $decL $compL @{}) }
Assert (Test-ActExtrIncludeBlock 'LASERS' 'req' $ctxL) 'req: LASERS inclos (amb lasers, pendent)'
Assert (Test-ActExtrIncludeBlock 'FAV_LASERS' 'fav' $ctxL) 'fav: FAV_LASERS inclos (autoritzacio)'
Assert (-not (Test-ActExtrIncludeBlock 'NO_LASERS' 'fav' $ctxL)) 'fav: NO_LASERS no inclos (amb lasers)'

Write-Host "`n--- Assistencia sanitaria ({{ASSISTENCIA}}) ---"
# Amb PAU obligat (PauLocal=Si) -> Annex III del Decret 30/2015.
Assert ($comp.PAU_OBLIGAT) 'PAU obligat (context assistencia)'
Assert ($comp.ASSISTENCIA -match 'Annex III') 'assistencia amb PAU -> Annex III Decret 30/2015'
# Sense PAU i aforament < 1000 -> farmaciola (Art. 48).
$decFarm = Build-ActExtrDecret @{ Aforament='300' }
$compFarm = Get-ActExtrComputed $decFarm
Assert (-not $compFarm.PAU_OBLIGAT) 'sense PAU (aforament 300)'
Assert ($compFarm.ASSISTENCIA -match 'farmaciola') 'assistencia sense PAU <1000 -> farmaciola'
# Sense PAU i aforament >= 1000 -> infermeria (Art. 48).
$decInf = Build-ActExtrDecret @{ Aforament='1500' }
Assert ((Get-ActExtrComputed $decInf).ASSISTENCIA -match 'infermeria') 'assistencia sense PAU >=1000 -> infermeria'
AssertEq (Resolve-ActExtrTokens 'En aquest cas {{ASSISTENCIA}}.' $compFarm) ('En aquest cas ' + $compFarm.ASSISTENCIA + '.') 'token {{ASSISTENCIA}} substituit'

Write-Host "`n--- Registre local (round-trip) ---"
$reg = Load-ActExtrRegistry
AssertEq (@($reg.Activitats).Count) 0 'registre buit inicialment'
$activity = [pscustomobject]@{
    IdGia='1429'; Header=@{ ID_GIA='1429'; TITULAR='BIG TOURS'; AFORAMENT='34719' }
    Decret=$ans; Punts=$delivered; Estat='pendent'
    CreatAt=(Get-Date).ToString('o'); ModificatAt=(Get-Date).ToString('o'); Historial=@()
}
$reg = Set-ActExtrActivity $reg $activity
Save-ActExtrRegistry $reg
$reg2 = Load-ActExtrRegistry
AssertEq (@($reg2.Activitats).Count) 1 'registre te 1 activitat desada'
$got = Get-ActExtrActivity $reg2 '1429'
Assert ($null -ne $got) 'es recupera l activitat per ID'
AssertEq $got.Header.TITULAR 'BIG TOURS' 'dades de l activitat persistides'
# Upsert (mateix ID) no duplica.
$activity2 = [pscustomobject]@{ IdGia='1429'; Header=@{ ID_GIA='1429'; TITULAR='BIG TOURS 2' }; Decret=$ans; Punts=@{}; Estat='tancat'; CreatAt=$got.CreatAt; ModificatAt=(Get-Date).ToString('o'); Historial=@() }
$reg2 = Set-ActExtrActivity $reg2 $activity2
AssertEq (@($reg2.Activitats).Count) 1 'upsert no duplica'
AssertEq ((Get-ActExtrActivity $reg2 '1429').Header.TITULAR) 'BIG TOURS 2' 'upsert actualitza'

# Neteja de la carpeta temporal del registre.
try { Remove-Item -LiteralPath $script:ActExtrRegistryDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }

exit (Write-TestSummary 'RESULTAT ACT_EXTR')
