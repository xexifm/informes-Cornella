# Llicencia: cataleg, fases, generacio i sincronitzacio
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

Write-Host "`n--- LLIC.json: la capa de Llicencia sobre REQ1 ---"
# LLIC no es un cataleg de deficiencies: per cada requeriment de REQ1 hi desa
# nomes el que es propi de Llicencia, i el text surt de REQ1 EN VIU. Per aixo:
#  - no pot sortir al menu "Requeriment - Nou" (donaria un informe buit);
#  - cada clau ha d'existir a REQ1, si no el lligam s'ha trencat en silenci.
$llicPath = Join-Path $EstructuralsDir 'LLIC.json'
if (Test-Path -LiteralPath $llicPath) {
    Assert (-not (@(Get-Catalegs) | Where-Object { $_.Name -eq 'LLIC.json' })) 'Get-Catalegs: LLIC NO surt com a cataleg triable'
    $llic = Get-Content -LiteralPath $llicPath -Raw -Encoding UTF8 | ConvertFrom-Json
    AssertEq ([string]$llic.familia) 'llicencia' 'LLIC.json: familia llicencia'
    $llicSecs = @($llic.nodes | ForEach-Object { [string]$_.titol })
    AssertEq (($llicSecs | Select-Object -First 3) -join ',') 'ABANS,DESPRES,PROPIS' 'LLIC.json: les tres seccions de punts, en ordre'
    Assert ([bool](@($llicSecs | Where-Object { $_ -like 'ANNEX 1*' }).Count -eq 1)) 'LLIC.json: hi ha la seccio de l''ANNEX 1'
    # L'ANNEX 1 nomes va al REQUERIMENT d'una llicencia PROVISIONAL.
    $annexSec = @($llic.nodes | Where-Object { [string]$_.titol -like 'ANNEX 1*' })[0]
    Assert ([bool](@($annexSec.fills).Count -ge 15)) 'LLIC.json: l''ANNEX 1 porta tot el text (no s''ha quedat a mitges)'
    # Totes les claus han d'existir a REQ1.
    $req1 = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json')
    # Una clau pot apuntar a un ITEM de REQ1 o -els dos punts d'instal-lacions
    # del bloc DESPRES- a una SUBSECCIO sencera, i llavors els seus items son
    # els SUB-PUNTS (vegeu _LlicItemsDeSubseccio).
    $clausReq1 = @{}
    foreach ($sec in $req1.Sections) {
        foreach ($el in $sec.Items) {
            if ($el.Kind -in 'item', 'subsection' -and -not [string]::IsNullOrWhiteSpace([string]$el.Short)) {
                $clausReq1[(_ItemKey $sec.Title $el.Short)] = $true
            }
        }
        # ...i la SECCIO sencera: el bloc DESPRES es porta seccions enteres
        # (Instal-lacions, Controls inicials...) amb una sola entrada.
        $clausReq1[[string]$sec.Title] = $true
    }
    $llicOrfes = New-Object System.Collections.ArrayList
    $llicLligats = 0
    foreach ($sec in $llic.nodes) {
        foreach ($it in @($sec.fills)) {
            $k = [string]$it.clau
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            $llicLligats++
            if (-not $clausReq1.ContainsKey($k)) { [void]$llicOrfes.Add($k) }
        }
    }
    AssertEq ($llicOrfes -join ' | ') '' 'LLIC.json: cap clau orfe (totes existeixen a REQ1)'
    Assert ([bool]($llicLligats -ge 20)) 'LLIC.json: la majoria de punts van lligats a REQ1, no copiats'
    # Els punts lligats NO poden portar text propi: si en portessin, el de REQ1
    # deixaria de manar i tornariem a tenir dos textos que mantenir.
    # Els que apunten a un ITEM no poden portar text propi (el de REQ1 mana). Els
    # que apunten a una SUBSECCIO si: la clau nomes els dona els SUB-PUNTS, i la
    # frase que els encapcala ("...que acrediti que s'han legalitzat les
    # seguents instal-lacions:") es d'ells.
    $subseccionsReq1 = @{}
    foreach ($sec in $req1.Sections) {
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'subsection' -and -not [string]::IsNullOrWhiteSpace([string]$el.Short)) {
                $subseccionsReq1[(_ItemKey $sec.Title $el.Short)] = $true
            }
        }
        $subseccionsReq1[[string]$sec.Title] = $true
    }
    $ambText = New-Object System.Collections.ArrayList
    foreach ($sec in $llic.nodes) {
        foreach ($it in @($sec.fills)) {
            $k = [string]$it.clau
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            if ($subseccionsReq1.ContainsKey($k)) { continue }
            if (@($it.cos).Count -gt 0) { [void]$ambText.Add([string]$it.titol) }
        }
    }
    AssertEq ($ambText -join ' | ') '' 'LLIC.json: cap punt lligat a un ITEM porta text propi (el text mana a REQ1)'

    # --- Resolucio dels punts contra REQ1 (el cor de l'eina) ------------------
    $idxR1 = _LlicIndexReq1 $req1
    Assert ([bool]($idxR1.Count -gt 100)) '_LlicIndexReq1: indexa els items de REQ1'
    foreach ($b in @('ABANS', 'DESPRES', 'PROPIS')) {
        $r = _LlicPuntsPerBloc $llic $idxR1 $b $req1
        AssertEq (@($r.Orfes) -join ' | ') '' ("_LlicPuntsPerBloc " + $b + ": cap clau orfe")
        Assert ([bool](@($r.Punts).Count -gt 0)) ("_LlicPuntsPerBloc " + $b + ": hi ha punts")
    }
    # El text ha de venir de REQ1, no de LLIC.
    $pAbans = @((_LlicPuntsPerBloc $llic $idxR1 'ABANS').Punts)
    $pSan = @($pAbans | Where-Object { [string]$_.Titol -eq 'Sanitat' })[0]
    Assert ([bool](@($pSan.Cos).Count -gt 0)) '_LlicPuntsPerBloc: el punt agafa el cos de REQ1'
    Assert ([bool]([string]@($pSan.Cos)[0] -like 'Sanitat.*')) '_LlicPuntsPerBloc: i es el text de REQ1 de debo'
    Assert ([bool](@($pSan.NoDisposa).Count -gt 0)) '_LlicPuntsPerBloc: i el "No es disposa" de LLIC'
    Assert ([bool](@($pSan.SiDisposa).Count -gt 0)) '_LlicPuntsPerBloc: i el "Es disposa" de LLIC'
    # Una clau que ja no existeix a REQ1 s'ha de DENUNCIAR, no ignorar: si no,
    # el punt desapareixeria de l'informe sense que ningu se n'assabentes.
    $rOrfe = _LlicPuntsPerBloc $llic @{} 'ABANS'
    Assert ([bool](@($rOrfe.Orfes).Count -ge 15)) '_LlicPuntsPerBloc: si REQ1 no te les claus, TOTES surten com a orfes'
    AssertEq @($rOrfe.Punts).Count 0 '_LlicPuntsPerBloc: i cap punt no es dona per bo'
    # El bloc DESPRES ha de portar el "Quan:".
    $pDesp = @((_LlicPuntsPerBloc $llic $idxR1 'DESPRES' $req1).Punts)
    Assert ([bool](@($pDesp | Where-Object { @($_.Quan).Count -gt 0 }).Count -ge 40)) '_LlicPuntsPerBloc DESPRES: els punts porten el "Quan:" (tambe els de les seccions expandides)'
}

Write-Host "`n--- Llicencia.ps1: fases, condicionals i textos ---"
$llFases = @(_LlicFases)
AssertEq $llFases.Count 3 '_LlicFases: els tres informes de la llicencia'
AssertEq ([string]$llFases[0].Clau) 'requeriment' '_LlicFases: el primer es el requeriment'
# Els dos punts CONDICIONALS: un nomes si es provisional i l'altre nomes si no.
Assert (_LlicCondicioEntra 'annexii' $false)        '_LlicCondicioEntra: el d''Annex II entra si NO es provisional'
Assert (-not (_LlicCondicioEntra 'annexii' $true))  '_LlicCondicioEntra: ...i no si ho es'
Assert (_LlicCondicioEntra 'provisional' $true)     '_LlicCondicioEntra: el de l''AMB entra si ES provisional'
Assert (-not (_LlicCondicioEntra 'provisional' $false)) '_LlicCondicioEntra: ...i no si no ho es'
Assert (_LlicCondicioEntra '' $true)                '_LlicCondicioEntra: sense condicio, entra sempre'
Assert (_LlicCondicioEntra '' $false)               '_LlicCondicioEntra: sense condicio, tambe sense provisional'
# La conclusio de cada fase. Al favorable PRE, la coda de les condicions es
# OPCIONAL: sense condicions la frase ha d'acabar amb un punt.
$cReq = _LlicConclusioText 'requeriment' $false
Assert ([bool]($cReq -like '*esmena de les defici*')) '_LlicConclusioText: requeriment'
$cPreSense = _LlicConclusioText 'favorable-pre' $false
$cPreAmb   = _LlicConclusioText 'favorable-pre' $true
Assert ([bool]($cPreSense -like '*tancat l*expedient.*')) '_LlicConclusioText: pre SENSE condicions acaba amb punt'
Assert (-not ($cPreSense -like '*sota les seg*'))        '_LlicConclusioText: pre sense condicions NO promet condicions'
Assert ([bool]($cPreAmb -like '*i sota les seg*ents condicions.*')) '_LlicConclusioText: pre AMB condicions hi afegeix la coda'
$cPost = _LlicConclusioText 'favorable-post' $false
Assert ([bool]($cPost -like '*per tancat l*expedient.*')) '_LlicConclusioText: post tanca l''expedient'
# EL TEXT VE DEL CATALEG, no del codi: hi ha de portar la negreta del **...**
# (com les de REQ1) i no pot quedar cap frase de conclusio escrita al programa.
Assert ([bool]($cPost -like '`*`**')) '_LlicConclusioText: la negreta ve del cataleg (**...**)'
$srcLlicC = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Llicencia.ps1') -Raw
Assert (-not ($srcLlicC -match 'Conclusio\s*=')) 'Llicencia: cap text de conclusio escrit al codi'
Assert (-not ($srcLlicC.Contains('Ho poso al seu coneixement'))) 'Llicencia: ni el tancament'
$srcMnsC = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'MnsTraspas.ps1') -Raw
Assert (-not ($srcMnsC.Contains('Ho poso al seu coneixement'))) 'MnsTraspas: ni el tancament'
# ...i el cataleg els te tots quatre, un per fase (i el pre, amb i sense condicions).
$grLlic = Read-Conclusions $Global:ConclusionsPath 'LLIC'
AssertEq (@($grLlic.Selectable).Count) 4 'cataleg: el grup LLIC porta les quatre conclusions'
foreach ($fLl in @('requeriment', 'favorable-pre', 'favorable-pre-condicions', 'favorable-post')) {
    Assert ([bool](@($grLlic.Selectable) | Where-Object { [string]$_.Title -eq $fLl })) ('cataleg: hi ha la conclusio "' + $fLl + '"')
}
AssertEq (_LlicConclusioText 'no-existeix' $false) '' '_LlicConclusioText: fase desconeguda -> buit'
# El paragraf del tecnic redactor.
$td = _LlicTextDocumentacio 'Simon Aledo Vives' '1.780' 'COITI d''Alacant' '20 de febrer de 2024'
Assert ([bool]($td -like '*Simon Aledo Vives*'))  '_LlicTextDocumentacio: hi surt el tecnic'
Assert ([bool]($td -like '*1.780*'))              '_LlicTextDocumentacio: i el numero de col·legiat'
Assert ([bool]($td -like '*20 de febrer de 2024.')) '_LlicTextDocumentacio: i la data, acabant amb punt'
AssertEq (_LlicTextDocumentacio '' '1' 'X' 'Y') '' '_LlicTextDocumentacio: sense tecnic no hi ha paragraf'

# ELS TRES INFORMES SON EL MATEIX DOCUMENT; el que canvia es que es diu de cada
# punt del bloc DESPRES. Abans el post era un informe curt que LLEGIA el
# pre-llicencia; l'usuari va ensenyar que a ma el fa sencer.
$efReq  = _LlicEstatDespres 'requeriment'
$efPre  = _LlicEstatDespres 'favorable-pre'
$efPost = _LlicEstatDespres 'favorable-post'
AssertEq ([string]$efReq.Estat) '' '_LlicEstatDespres: al requeriment no es diu si es te o no'
AssertEq ([bool]$efReq.AmbEstat) $false '_LlicEstatDespres: ...i per tant no es demana'
AssertEq ([string]$efPre.Estat) 'no' '_LlicEstatDespres: al favorable pre, per defecte NO es disposa'
AssertEq ([string]$efPost.Estat) 'si' '_LlicEstatDespres: al favorable post, per defecte SI'
Assert ([bool]((@($efPre.NoDisposa) -join ' ') -like 'No es disposa de la documentaci*.')) '_LlicEstatDespres: el text del pre'
Assert ([bool]((@($efPost.SiDisposa) -join ' ') -like '*`[CAMP: Id Firmadoc`]*')) '_LlicEstatDespres: el post demana l''Id Firmadoc'
Assert ([bool]$efPost.AmbDades) '_LlicEstatDespres: ...i per aixo la pantalla demana dades'
AssertEq ([string](_LlicEstatDespres 'no-existeix').Estat) '' '_LlicEstatDespres: fase desconeguda -> res'

# Els punts del bloc DESPRES agafen els textos de la fase, i els que ja en tenen
# de propis al cataleg se'ls queden.
$pFase = @(
    [pscustomobject]@{ Clau='A'; Subseccio=''; Titol='Sense text'; Condicio=''; Cos=@('x'); NoDisposa=@(); SiDisposa=@(); Quan=@('q'); Subs=@() },
    [pscustomobject]@{ Clau='B'; Subseccio=''; Titol='Amb text';   Condicio=''; Cos=@('y'); NoDisposa=@('El seu propi'); SiDisposa=@('El seu propi si'); Quan=@(); Subs=@() })
$rFase = @(_LlicPuntsAmbEstatFase $pFase 'favorable-pre')
AssertEq $rFase.Count 2 '_LlicPuntsAmbEstatFase: no perd cap punt'
Assert ([bool]((@($rFase[0].NoDisposa) -join ' ') -like 'No es disposa de la documentaci*')) '_LlicPuntsAmbEstatFase: el que no en te, agafa el de la fase'
AssertEq (@($rFase[1].NoDisposa) -join ' ') 'El seu propi' '_LlicPuntsAmbEstatFase: el que en te, se''l queda'
AssertEq (@($rFase[0].Quan) -join ' ') 'q' '_LlicPuntsAmbEstatFase: el "Quan:" no es toca'
# I NO toca els punts originals (funcio pura).
AssertEq (@($pFase[0].NoDisposa).Count) 0 '_LlicPuntsAmbEstatFase: no modifica el que li arriba'
AssertEq (@(_LlicPuntsAmbEstatFase @() 'requeriment').Count) 0 '_LlicPuntsAmbEstatFase: sense punts, cap'

# Nom del fitxer: data al principi, com la resta d'informes (aixi "Actualitzar
# base d'informes" el reconeix).
$nf = _LlicNomFitxer ([datetime]'2026-08-03') 'requeriment' '1433'
AssertEq $nf '2026-08-03_LlicReq_GIA 1433.docx' '_LlicNomFitxer: requeriment'
Assert ([bool]((_LlicNomFitxer ([datetime]'2026-08-03') 'favorable-pre' '1') -like '*LlicFavPre*'))  '_LlicNomFitxer: favorable pre'
Assert ([bool]((_LlicNomFitxer ([datetime]'2026-08-03') 'favorable-post' '1') -like '*LlicFavPost*')) '_LlicNomFitxer: favorable post'
Assert (-not ((_LlicNomFitxer ([datetime]'2026-08-03') 'requeriment' 'A/B:C') -match '[\\/:*?"<>|]')) '_LlicNomFitxer: fora els caracters que Windows no admet'
AssertEq (_VistaActExtrTitol '[[INCENDIS]] Incendis') 'Incendis  [INCENDIS]' '_VistaActExtrTitol: etiqueta + clau'
AssertEq (_VistaActExtrTitol '[[MEMORIA_A]] ::CHILD:: a) Identificacio') 'a) Identificacio  [MEMORIA_A]' '_VistaActExtrTitol: treu el token'
AssertEq (_VistaActExtrTitol '[[REQ_INTRO]]') 'REQ_INTRO' '_VistaActExtrTitol: sense etiqueta -> la clau'
AssertEq ([bool](_VistaCalRegenerar $false ([datetime]'2026-01-01') ([datetime]::MinValue) $false)) $true '_VistaCalRegenerar: no hi ha vista -> si'
AssertEq ([bool](_VistaCalRegenerar $true ([datetime]'2026-02-01') ([datetime]'2026-01-01') $false)) $true '_VistaCalRegenerar: JSON mes nou -> si'
AssertEq ([bool](_VistaCalRegenerar $true ([datetime]'2026-01-01') ([datetime]'2026-02-01') $false)) $false '_VistaCalRegenerar: vista al dia -> no (evita commits inutils)'
AssertEq ([bool](_VistaCalRegenerar $true ([datetime]'2026-01-01') ([datetime]'2026-02-01') $true)) $true '_VistaCalRegenerar: -Force -> sempre'
# En canviar el FORMAT de les vistes cal regenerar-les encara que el .docx sigui
# mes nou que el JSON (si no, es quedarien amb el format antic per sempre).
AssertEq ([bool]($Script:VistaWordVersio -ge 3)) $true 'VistaWordVersio: versio de format definida'
# La tipografia base viu a Format.ps1 (no a la vista): un document nou de Word
# sortiria en Calibri alineat a l'esquerra i no s'assemblaria a l'informe.
AssertEq ($Script:ReportFormatConfig.BodyFontName) 'Bookman Old Style' 'Format: el tipus de lletra base es Bookman Old Style'
AssertEq ($Script:ReportFormatConfig.BodyAlignment) 3 'Format: justificat (3 = wdAlignParagraphJustify)'
AssertEq ($Script:ReportFormatConfig.BaseLineSpacing) 1.15 'Format: interlineat 1,15 com la plantilla'
# ELS MARGES HAN DE QUADRAR AMB LA PLANTILLA, i per aixo la prova els llegeix
# d'ella en lloc de repetir-ne els numeros: l'informe COPIA '0 CAPCALERA.docx' i
# les vistes surten d'un document nou amb $ReportFormatConfig. Si divergeixen,
# la vista deixa de semblar-se a l'informe i no ho diria ningu.
$capMargePath = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
if (Test-Path -LiteralPath $capMargePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zipM = [System.IO.Compression.ZipFile]::OpenRead($capMargePath)
    try {
        $entM = $zipM.GetEntry('word/document.xml')
        $srM = New-Object System.IO.StreamReader($entM.Open())
        $xmlM = $srM.ReadToEnd(); $srM.Close()
    } finally { $zipM.Dispose() }
    $mM = [regex]::Match($xmlM, '<w:pgMar[^/]*/>')
    Assert $mM.Success 'Format: la plantilla porta el seu pgMar'
    $twM = @{}
    foreach ($mm in [regex]::Matches($mM.Value, 'w:(\w+)="(-?\d+)"')) { $twM[$mm.Groups[1].Value] = [int]$mm.Groups[2].Value }
    # 20 twips = 1 pt. Mig punt de tolerancia: el Word arrodoneix.
    AssertNear $Script:ReportFormatConfig.PageMarginLeftPt  ($twM['left']   / 20.0) 0.5 'Format: el marge esquerre quadra amb la plantilla'
    AssertNear $Script:ReportFormatConfig.PageMarginRightPt ($twM['right']  / 20.0) 0.5 'Format: ...i el dret'
    AssertNear $Script:ReportFormatConfig.PageMarginTopPt   ($twM['top']    / 20.0) 0.5 'Format: ...i el de dalt'
    AssertNear $Script:ReportFormatConfig.PageMarginBottomPt ($twM['bottom'] / 20.0) 0.5 'Format: ...i el de baix'
    # I que siguin els 2,5 cm que va demanar l'usuari (1 cm = 566,93 twips).
    AssertNear ($twM['left'] / 566.93) 2.5 0.02 'Format: 2,5 cm de marge esquerre'
    AssertNear ($twM['right'] / 566.93) 2.5 0.02 'Format: 2,5 cm de marge dret'
}
AssertEq ([bool](Get-Command Format-ApplyBaseStyle -ErrorAction SilentlyContinue)) $true 'Format-ApplyBaseStyle existeix (l''apliquen les vistes)'

Write-Host "`n--- Migracio.ps1: carpeta 'local' ---"
# Rutes: totes pengen de local\ i el nom de cada subcarpeta viu NOMES aqui
# (abans Ruta.ps1 repetia 'BASE DE DADES ACTIVITATS' pel seu compte).
$tstLocal = $tstClone + $tstSep + 'local' + $tstSep
AssertEq (Get-LocalDir $tstClone) ($tstClone + $tstSep + 'local') 'Get-LocalDir: local a l''arrel del clone'
AssertEq (Get-LocalSubdir $tstClone 'Informes')   ($tstLocal + 'informes-generats')      'Get-LocalSubdir: informes generats'
AssertEq (Get-LocalSubdir $tstClone 'Rutes')      ($tstLocal + 'rutes-generades')        'Get-LocalSubdir: mapes de ruta'
AssertEq (Get-LocalSubdir $tstClone 'Activitats') ($tstLocal + 'base-dades-activitats')  'Get-LocalSubdir: base de dades d''activitats'
AssertEq (Get-LocalSubdir $tstClone 'ActExtr')    ($tstLocal + 'base-dades-actextr')     'Get-LocalSubdir: registre ACT_EXTR'
AssertEq (Get-LocalSubdir $tstClone 'Vistes')     ($tstLocal + 'vistes-catalegs')        'Get-LocalSubdir: vistes en Word'
AssertEq (Get-LocalSubdir $tstClone 'Seguiment')  ($tstLocal + 'seguiment-gia')          'Get-LocalSubdir: llistats de seguiment del GIA'
AssertEq (Get-LocalSubdir $tstClone 'Geocodificacio') ($tstLocal + 'geocodificacio')     'Get-LocalSubdir: portals del Cadastre i mapes de coordenades'
$errClau = $false
try { [void](Get-LocalSubdir $tstClone 'NoExisteix') } catch { $errClau = $true }
AssertEq $errClau $true 'Get-LocalSubdir: una clau desconeguda peta (no retorna una ruta inventada)'
# Migracions: les 4 carpetes velles de l'arrel, cadascuna al seu lloc nou.
$migs = @(Get-MigracionsLocal $tstClone)
AssertEq $migs.Count 4 'Get-MigracionsLocal: 4 carpetes a moure'
AssertEq ([bool]($migs[0] -is [pscustomobject])) $true 'Get-MigracionsLocal: retorna un array PLA (no la llista sencera)'
$vells = @($migs | ForEach-Object { Split-Path -Leaf $_.Origen })
AssertEq ($vells -join '|') 'Informes generats|Rutes generades|BASE DE DADES ACTIVITATS|BASE DE DADES ACT_EXTR' 'Get-MigracionsLocal: origens = les carpetes velles de l''arrel'
AssertEq $migs[2].Desti ($tstLocal + 'base-dades-activitats') 'Get-MigracionsLocal: desti dins de local'
AssertEq ([bool](@($migs | Where-Object { $_.Origen -like '*ESTRUCTURALS*' }).Count -eq 0)) $true 'Get-MigracionsLocal: ESTRUCTURALS no es mou (les vistes van a part)'

# La CLASSIFICACIO surt SOLA de l'Excel; ja no es pregunta. La llei la diu la
# columna "Classificacio general annex".
AssertEq (_ClassificacioText 'L18 Cert' '12.25')        'Llei 18/2020; Epígraf 12.25'            '_ClassificacioText: L18 Cert -> Llei 18/2020, sense annex'
AssertEq (_ClassificacioText 'L18 Proj i Cert' '3.1')   'Llei 18/2020; Epígraf 3.1'              '_ClassificacioText: L18 Proj i Cert'
AssertEq (_ClassificacioText 'L18' '')                  'Llei 18/2020'                          '_ClassificacioText: L18 sense apartat'
AssertEq (_ClassificacioText 'II' '12.25')              'Llei 20/2009; Annex II; Epígraf 12.25'  '_ClassificacioText: Annex II'
AssertEq (_ClassificacioText 'III' '4')                 'Llei 20/2009; Annex III; Epígraf 4'     '_ClassificacioText: Annex III'
AssertEq (_ClassificacioText '' '')                     ''                                      '_ClassificacioText: buit'

# El bloc ABANS surt de REQ1 (4 seccions), no de la llista de LLIC: aixi un
# requeriment nou d'aquelles seccions hi surt sol.
$secAb = @(_LlicSeccionsAbans)
AssertEq $secAb.Count 2 '_LlicSeccionsAbans: les DUES seccions de documentacio (el PAU i els controls inicials han passat a DESPRES)'
Assert ([bool](_LlicEsSeccioAbans 'Autoritzacions / Informes preceptius')) '_LlicEsSeccioAbans: Autoritzacions'
Assert ([bool](_LlicEsSeccioAbans 'Registres')) '_LlicEsSeccioAbans: Registres (d''aqui surt el RASIC)'
Assert (-not (_LlicEsSeccioAbans ('Pla d' + [char]0x2019 + 'Autoprotecci' + [char]0x00F3))) '_LlicEsSeccioAbans: el PAU ja NO (va a DESPRES)'
Assert (-not (_LlicEsSeccioAbans 'Controls inicials')) '_LlicEsSeccioAbans: ni els controls inicials'
# La comparacio segueix sent sense accents ni apostrof tipografic.
Assert ([bool](_LlicEsSeccioAbans 'AUTORITZACIONS / INFORMES PRECEPTIUS')) '_LlicEsSeccioAbans: sense distingir majuscules'
Assert (-not (_LlicEsSeccioAbans 'Projecte')) '_LlicEsSeccioAbans: Projecte NO'
Assert (-not (_LlicEsSeccioAbans ('Controls peri' + [char]0x00F2 + 'dics'))) '_LlicEsSeccioAbans: els periodics tampoc (son una altra cosa)'
if ((Test-Path -LiteralPath $llicPathX) -and (Test-Path -LiteralPath (Join-Path $EstructuralsDir 'REQ1.json'))) {
    $req1Ab = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json')
    $llicAb = Read-LlicCataleg $llicPathX
    $bAb = _LlicPuntsPerBloc $llicAb (_LlicIndexReq1 $req1Ab) 'ABANS' $req1Ab
    $esperats = 0
    foreach ($sc in @($req1Ab.Sections)) {
        if (-not (_LlicEsSeccioAbans ([string]$sc.Title))) { continue }
        $esperats += @($sc.Items | Where-Object { $_.Kind -eq 'item' -and $_.Short }).Count
    }
    AssertEq (@($bAb.Punts).Count) $esperats 'ABANS: hi son TOTS els items de les 4 seccions de REQ1'
    Assert ((@($bAb.Punts).Count) -gt 25) 'ABANS: inclou els items de dins de les subseccions'
    # I les 4 seccions NO poden quedar al pas Projecte.
    $secProj = @(@($req1Ab.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
    Assert (-not (@($secProj) | Where-Object { _LlicEsSeccioAbans ([string]$_.Title) })) 'Projecte: cap seccio de documentacio (no es pot demanar dues vegades)'
    Assert ((@($secProj).Count) -lt (@($req1Ab.Sections).Count)) 'Projecte: se n''han tret seccions'
}

# ---------------------------------------------------------------------------
# EL BLOC DESPRES ES PORTA SECCIONS SENCERES DE REQ1
# ---------------------------------------------------------------------------
# Una entrada de LLIC.json amb una clau que NO es un item (una SECCIO, o una
# "Seccio::Subseccio") s'EXPANDEIX: un punt per cada item d'aquella part, amb el
# text LITERAL de REQ1 i el mateix "Quan:" per a tots. Abans hi havia dos punts
# amb text propi que mantenien a ma la llista d'instal-lacions.
$Global:_secInst = ('Instal' + [char]0x00B7 + 'lacions')
$Global:_secCtrlI = 'Controls inicials'
$Global:_subITC = ('Incendis::Documentaci' + [char]0x00F3 + ' (ITC SP)')

if ((Test-Path -LiteralPath $llicPathX) -and (Test-Path -LiteralPath (Join-Path $EstructuralsDir 'REQ1.json'))) {
    $req1In = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json')
    $llicIn = Read-LlicCataleg $llicPathX
    $idxIn  = _LlicIndexReq1 $req1In

    # _LlicItemsDeSubseccio: tota la seccio, o nomes una subseccio.
    $itInst = @(_LlicItemsDeSubseccio $req1In $Global:_secInst)
    $itITC  = @(_LlicItemsDeSubseccio $req1In $Global:_subITC)
    Assert ($itInst.Count -ge 30) ('_LlicItemsDeSubseccio: la seccio SENCERA d''instal-lacions (' + $itInst.Count + ')')
    AssertEq $itITC.Count 4 '_LlicItemsDeSubseccio: nomes la subseccio Documentacio (ITC SP)'
    Assert (-not (@($itITC | ForEach-Object { [string]$_.Short }) | Where-Object { $_ -like 'RIPCI*' })) '_LlicItemsDeSubseccio: no s''endu la subseccio seguent'
    AssertEq (@(_LlicItemsDeSubseccio $req1In 'No existeix').Count) 0 '_LlicItemsDeSubseccio: seccio desconeguda -> cap item'
    AssertEq (@(_LlicItemsDeSubseccio $req1In '').Count) 0 '_LlicItemsDeSubseccio: clau buida -> cap item'

    # _LlicSeccionsExpandides: surt del CATALEG, no d'una llista al codi.
    $exp = @(_LlicSeccionsExpandides $llicIn $idxIn)
    Assert ([bool]($exp -contains $Global:_secInst)) '_LlicSeccionsExpandides: Instal-lacions'
    Assert ([bool]($exp -contains $Global:_secCtrlI)) '_LlicSeccionsExpandides: Controls inicials'
    Assert ([bool]($exp -contains $Global:_subITC)) '_LlicSeccionsExpandides: Incendis / Documentacio (ITC SP)'
    Assert (-not (@($exp) | Where-Object { $idxIn.ContainsKey($_) })) '_LlicSeccionsExpandides: cap clau d''item'

    # El bloc DESPRES sencer.
    $rDe = _LlicPuntsPerBloc $llicIn $idxIn 'DESPRES' $req1In
    AssertEq (@($rDe.Orfes) -join ' | ') '' 'DESPRES: cap clau orfe'
    $pDe = @($rDe.Punts)
    Assert ($pDe.Count -ge 50) ('DESPRES: hi caben totes les seccions expandides (' + $pDe.Count + ' punts)')
    # ...i el text es LITERAL, no retallat.
    $unInst = @($pDe | Where-Object { [string]$_.Clau -like ($Global:_secInst + '::*') })[0]
    $elInst = $idxIn[[string]$unInst.Clau]
    AssertEq ((@($unInst.Cos) -join '|')) ((@($elInst.BodyLines) -join '|')) 'DESPRES: el text de les seccions expandides es LITERAL de REQ1'
    Assert ([bool](@($unInst.Quan).Count -gt 0)) 'DESPRES: ...i cada punt porta el "Quan:" de l''entrada'
    # El PAU porta el seu "Quan:" propi, diferent de la resta.
    $pPau = @($pDe | Where-Object { [string]$_.Clau -like ('Pla d*Autoprotecci*::*') })[0]
    Assert ($null -ne $pPau) 'DESPRES: hi ha el Pla d''Autoproteccio'
    Assert ([bool]((@($pPau.Quan) -join ' ') -match 'sis mesos')) 'DESPRES: el PAU porta el seu "Quan:" de sis mesos'

    # LES SECCIONS, en l'ordre del cataleg i amb Instal-lacions al final.
    $grDe = @(_LlicAgrupaPunts $pDe)
    AssertEq ([string]$grDe[0].Titol) '' 'DESPRES: primer els punts propis (sense seccio)'
    AssertEq ([string]$grDe[$grDe.Count - 1].Titol) $Global:_secInst 'DESPRES: i Instal-lacions al final'
    # L'ARBRE ES DE DOS NIVELLS (seccio > subseccio), com el Pas 3 de REQ1: una
    # mateixa seccio hi surt tants cops com subseccions te, i el que no s'ha de
    # repetir es el PARELL.
    $parellsDe = @(@($grDe) | ForEach-Object { [string]$_.Titol + '||' + [string]$_.Sub })
    AssertEq (@($parellsDe | Select-Object -Unique).Count) $parellsDe.Count 'DESPRES: cap seccio+subseccio repetida a l''arbre'
    $instDe = @(@($grDe) | Where-Object { [string]$_.Titol -eq $Global:_secInst })
    Assert ($instDe.Count -ge 2) ('DESPRES: Instal-lacions es parteix per subseccions (' + $instDe.Count + ')')
    Assert (-not (@($instDe) | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.Sub) })) 'DESPRES: ...i totes tenen subseccio'

    # EL PAU I ELS CONTROLS INICIALS JA NO SON A ABANS.
    $pAb = @((_LlicPuntsPerBloc $llicIn $idxIn 'ABANS' $req1In).Punts)
    Assert (-not (@($pAb) | Where-Object { (_LlicSeccioDePunt $_) -like 'Pla d*' })) 'ABANS: el PAU ja no hi es'
    Assert (-not (@($pAb) | Where-Object { (_LlicSeccioDePunt $_) -eq $Global:_secCtrlI })) 'ABANS: ni els controls inicials'
    AssertEq (@($pAb | ForEach-Object { _LlicSeccioDePunt $_ } | Select-Object -Unique).Count) 2 'ABANS: nomes dues seccions'

    # CAP PUNT DE REQ1 A DOS BLOCS ALHORA.
    $clausAb = @($pAb | ForEach-Object { [string]$_.Clau })
    $clausDe = @($pDe | ForEach-Object { [string]$_.Clau } | Where-Object { $_ })
    $tots2 = @($clausAb | Where-Object { $clausDe -contains $_ })
    AssertEq ($tots2 -join ' | ') '' 'Cap punt de REQ1 surt a ABANS i a DESPRES alhora'

    # ...i al pas PROJECTE no hi ha res del que ja es demana.
    $senseAb = @(@($req1In.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
    $secPr = @(_LlicSeccionsSenseSubseccions $senseAb $exp)
    AssertEq (@(_LlicItemsDeSubseccio ([pscustomobject]@{ Sections = $secPr }) $Global:_secInst).Count) 0 'Projecte: Instal-lacions ja no hi es'
    AssertEq (@(_LlicItemsDeSubseccio ([pscustomobject]@{ Sections = $secPr }) $Global:_subITC).Count) 0 'Projecte: ni la Documentacio (ITC SP)'
    # ...pero la resta d'Incendis SI (nomes en marxa una subseccio).
    $inc = @(_LlicItemsDeSubseccio ([pscustomobject]@{ Sections = $secPr }) 'Incendis')
    Assert ($inc.Count -ge 20) ('Projecte: la resta d''Incendis s''hi queda (' + $inc.Count + ')')
    # I CADA punt de REQ1 es EN ALGUN LLOC, un sol cop.
    $totsReq1 = (@($req1In.Sections) | ForEach-Object { @($_.Items | Where-Object { $_.Kind -eq 'item' }).Count } | Measure-Object -Sum).Sum
    $nPr = (@($secPr) | ForEach-Object { @($_.Items | Where-Object { $_.Kind -eq 'item' }).Count } | Measure-Object -Sum).Sum
    AssertEq ($clausAb.Count + $clausDe.Count + $nPr) $totsReq1 'Tots els punts de REQ1 son en algun bloc, i nomes en un'
}

# ---------------------------------------------------------------------------
# CONTROLS QUE ES TREPITGEN (_TrobaSolapaments, UiComuns.ps1)
# ---------------------------------------------------------------------------
# El defecte recurrent del programa. La geometria es pura i es pot provar aqui;
# el que no es pot es dibuixar una finestra, i per aixo la comprovacio de debo
# la fa el PROGRAMA en obrir cada pantalla.
AssertEq (@(_TrobaSolapaments @()).Count) 0 '_TrobaSolapaments: sense controls, cap solapament'
$rcA = @(@{ Nom='A'; X=0; Y=0; W=100; H=20 }, @{ Nom='B'; X=200; Y=0; W=100; H=20 })
AssertEq (@(_TrobaSolapaments $rcA).Count) 0 '_TrobaSolapaments: separats, cap'
$rcB = @(@{ Nom='A'; X=0; Y=0; W=100; H=20 }, @{ Nom='B'; X=100; Y=0; W=100; H=20 })
AssertEq (@(_TrobaSolapaments $rcB).Count) 0 '_TrobaSolapaments: tocant-se per la vora, cap'
$rcC = @(@{ Nom='Titol'; X=0; Y=0; W=300; H=20 }, @{ Nom='Xip'; X=250; Y=0; W=100; H=20 })
$solC = @(_TrobaSolapaments $rcC)
AssertEq $solC.Count 1 '_TrobaSolapaments: el titol per sota del xip, ENXAMPAT'
Assert ([bool]($solC[0] -like '*Titol*' -and $solC[0] -like '*Xip*')) '_TrobaSolapaments: i diu quins son'
Assert ([bool]($solC[0] -like '*250,0*')) '_TrobaSolapaments: ...i on'
$rcD = @(@{ Nom='Fons'; X=0; Y=0; W=400; H=60 }, @{ Nom='Boto'; X=10; Y=10; W=80; H=24 })
AssertEq (@(_TrobaSolapaments $rcD).Count) 0 '_TrobaSolapaments: un DINS de l''altre es un fons, no un error'
$rcE = @(@{ Nom='A'; X=0; Y=0; W=100; H=20 }, @{ Nom='B'; X=0; Y=10; W=100; H=20 })
AssertEq (@(_TrobaSolapaments $rcE).Count) 1 '_TrobaSolapaments: tambe en vertical'

# LA TOLERANCIA. Sense ella el programa avisava de pantalles que es veuen
# perfectament (una etiqueta que passa un pixel per sota d'un radio) i l'avis es
# tornava soroll que ningu llegia. Ha de callar quan es freguen i cridar quan
# es tapen de debo. Els rectangles son ELS DE VERITAT del programa.
$rcTitol = @(@{ Nom='Titol'; X=76; Y=7; W=300; H=31 }, @{ Nom='Subtitol'; X=76; Y=33; W=300; H=18 })
AssertEq (@(_TrobaSolapaments $rcTitol).Count) 0 '_TrobaSolapaments: capcalera titol/subtitol, no molesta'
$rcRadio = @(@{ Nom='Radio'; X=30; Y=98; W=460; H=22 }, @{ Nom='Descripcio'; X=50; Y=119; W=440; H=18 })
AssertEq (@(_TrobaSolapaments $rcRadio).Count) 0 '_TrobaSolapaments: radio i la seva descripcio, no molesta'
$rcPoc = @(@{ Nom='A'; X=0; Y=0; W=200; H=40 }, @{ Nom='B'; X=195; Y=0; W=200; H=40 })
AssertEq (@(_TrobaSolapaments $rcPoc).Count) 0 '_TrobaSolapaments: 5 px de frec, no molesta'
$rcMolt = @(@{ Nom='Etiqueta'; X=50; Y=264; W=450; H=32 }, @{ Nom='Continuar'; X=370; Y=286; W=130; H=32 })
AssertEq (@(_TrobaSolapaments $rcMolt).Count) 1 '_TrobaSolapaments: etiqueta sota el boto, AVISA'

# ...i la pantalla que ho patia: els botons del Pas 1 de Llicencia ja no van a
# una Y clavada al codi, surten del peu de l'etiqueta.
$srcLlic = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Llicencia.ps1') -Raw
$iFase = $srcLlic.IndexOf('function Select-LlicFase')
$blocFase = $srcLlic.Substring($iFase, 4000)
Assert ($blocFase.Contains('$lbl2.Bottom')) 'Select-LlicFase: els botons surten del peu de l''etiqueta'
Assert (-not ($blocFase.Contains('Point(370, 286)'))) 'Select-LlicFase: ...i ja no d''una Y clavada'

# EL TITOL DE LA RAJOLA DEL MENU s'ha de dibuixar ACOTAT pels xips. Prova de
# FONT (el menu nomes es pinta a Windows): el xip "Dades" tapava el "LL Prov" de
# Llicencia perque el titol es dibuixava en un PUNT, sense limit d'amplada.
$srcMenu = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1') -Raw
$iPaint = $srcMenu.IndexOf('$paintHandler = {')
$iFiPaint = $srcMenu.IndexOf('$result = @{ Choice = $null }')
$paint = $srcMenu.Substring($iPaint, $iFiPaint - $iPaint)
Assert ($paint.Contains('EndEllipsis')) 'menu: el titol de la rajola es retalla amb punts suspensius'
Assert ($paint.Contains('$limit = $rect.Width - 14')) 'menu: hi ha un limit dret per al text'
Assert ($paint.Contains('$entry.ExtraChipRect.Left')) 'menu: ...i el marca el xip mes a l''esquerra'
Assert ($paint.IndexOf('$entry.DocChipRect = New-Object') -lt $paint.IndexOf('Titol + subtitol')) 'menu: els xips es calculen ABANS del titol'
Assert (-not ($paint -match 'DrawText\(\$g, \$main, \$fMain, \(New-Object System\.Drawing\.Point')) 'menu: el titol ja no es dibuixa en un punt sense limit'

# ---------------------------------------------------------------------------
# L'ARBRE de la pantalla de documentacio (agrupacio per seccions)
# ---------------------------------------------------------------------------
# La seccio d'un punt surt de la seva clau ("Seccio::Item", _ItemKey): els
# punts de REQ1 en tenen i els PROPIS no, i aquests van al primer nivell.
AssertEq (_LlicSeccioDePunt ([pscustomobject]@{ Clau = 'Registres::RASIC' })) 'Registres' '_LlicSeccioDePunt: la seccio de la clau'
AssertEq (_LlicSeccioDePunt ([pscustomobject]@{ Clau = 'Autoritzacions / Informes preceptius::Sanitat' })) 'Autoritzacions / Informes preceptius' '_LlicSeccioDePunt: amb barres i espais'
AssertEq (_LlicSeccioDePunt ([pscustomobject]@{ Clau = '' })) '' '_LlicSeccioDePunt: sense clau, primer nivell'
AssertEq (_LlicSeccioDePunt ([pscustomobject]@{ Clau = 'sense separador' })) '' '_LlicSeccioDePunt: sense ::, primer nivell'
AssertEq (_LlicSeccioDePunt ([pscustomobject]@{ Titol = 'x' })) '' '_LlicSeccioDePunt: un punt sense propietat Clau no peta'

# L'etiqueta del node: el titol si en te, si no la primera linia del cos.
AssertEq (_LlicEtiquetaPunt ([pscustomobject]@{ Titol = 'Sanitat'; Cos = @('Sanitat. Molt de text...') })) 'Sanitat' '_LlicEtiquetaPunt: el titol mana'
AssertEq (_LlicEtiquetaPunt ([pscustomobject]@{ Titol = ''; Cos = @('', 'La primera de debo') })) 'La primera de debo' '_LlicEtiquetaPunt: sense titol, la primera linia amb text'
AssertEq (_LlicEtiquetaPunt ([pscustomobject]@{ Titol = "  dos   espais  " })) 'dos espais' '_LlicEtiquetaPunt: espais collapsats'
$etLlarga = _LlicEtiquetaPunt ([pscustomobject]@{ Titol = ('x' * 200) }) 20
AssertEq $etLlarga.Length 21 '_LlicEtiquetaPunt: tallada a la mida demanada (+ els punts suspensius)'
AssertEq $etLlarga[20] ([char]0x2026) '_LlicEtiquetaPunt: i acaba amb punts suspensius'
AssertEq (_LlicEtiquetaPunt ([pscustomobject]@{ Titol = ('x' * 200) }) 0).Length 200 '_LlicEtiquetaPunt: amb max 0 no talla'

# L'agrupacio, sobre els punts REALS d'ABANS (els 2 propis condicionals al
# davant, com els munta el pas 3 del wizard).
if ((Test-Path -LiteralPath $llicPathX) -and (Test-Path -LiteralPath (Join-Path $EstructuralsDir 'REQ1.json'))) {
    $req1Gr = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json')
    $llicGr = Read-LlicCataleg $llicPathX
    $idxGr  = _LlicIndexReq1 $req1Gr
    $propisGr = @((_LlicPuntsPerBloc $llicGr $idxGr 'PROPIS').Punts)
    $abansGr  = @((_LlicPuntsPerBloc $llicGr $idxGr 'ABANS' $req1Gr).Punts)
    $totsGr   = @($propisGr) + @($abansGr)
    $grups = @(_LlicAgrupaPunts $totsGr)

    Assert ([bool](@($propisGr).Count -ge 2)) '_LlicAgrupaPunts: els punts propis hi son'
    AssertEq ([string]$grups[0].Titol) '' '_LlicAgrupaPunts: el primer grup es el del primer nivell (els propis)'
    AssertEq (@($grups[0].Idx).Count) (@($propisGr).Count) '_LlicAgrupaPunts: i hi son tots els propis, nomes ells'
    # Despres, les 4 seccions de documentacio de REQ1, en ordre de cataleg.
    $titolsGr = @(@($grups) | Select-Object -Skip 1 | ForEach-Object { [string]$_.Titol })
    AssertEq (@($titolsGr | Select-Object -Unique).Count) 2 '_LlicAgrupaPunts: les 2 seccions de documentacio de REQ1'
    # Dos nivells: dins d'una seccio, un grup per subseccio i en ordre.
    $parellsGr = @(@($grups) | ForEach-Object { [string]$_.Titol + '||' + [string]$_.Sub })
    AssertEq (@($parellsGr | Select-Object -Unique).Count) $parellsGr.Count '_LlicAgrupaPunts: cap parell seccio+subseccio repetit'
    Assert (-not (@($titolsGr) | Where-Object { -not (_LlicEsSeccioAbans $_) })) '_LlicAgrupaPunts: i totes son de documentacio'
    # CAP punt perdut ni duplicat: aplanar els grups ha de donar 0..N-1.
    $plans = @(@($grups) | ForEach-Object { @($_.Idx) })
    AssertEq $plans.Count $totsGr.Count '_LlicAgrupaPunts: cap punt perdut ni duplicat'
    AssertEq ((@($plans) | Sort-Object) -join ',') ((0..($totsGr.Count - 1)) -join ',') '_LlicAgrupaPunts: hi son tots els indexs, un sol cop'
    # I DINS de cada grup, l'ordre del cataleg es respecta.
    foreach ($g in $grups) {
        $ix = @($g.Idx)
        AssertEq ($ix -join ',') ((@($ix) | Sort-Object) -join ',') ('_LlicAgrupaPunts: ordre de cataleg dins de "' + [string]$g.Titol + '"')
    }
}

# El titol que obre el full de signatures de l'ANNEX 1.
Assert ([bool](_LlicEsTitolAcceptacio ('Document d' + [char]0x2019 + 'acceptaci' + [char]0x00F3 + ' del cessament dels usos...'))) '_LlicEsTitolAcceptacio: el titol de la plantilla'
Assert ([bool](_LlicEsTitolAcceptacio "document d'acceptacio del cessament")) '_LlicEsTitolAcceptacio: sense accents ni majuscules'
Assert (-not (_LlicEsTitolAcceptacio 'Jo.........., amb DNI......')) '_LlicEsTitolAcceptacio: una linia del full, no'
Assert (-not (_LlicEsTitolAcceptacio '')) '_LlicEsTitolAcceptacio: buit'

Write-Host "`n--- Llicencia: la GENERACIO sencera (amb el Word simulat) ---"
# Es genera un informe de debo amb els dobles de Format.ps1 i un Word de
# mentida. Aixo hauria enxampat, de cop, gairebe tot el que va fallar a la
# primera prova real: el [[URL]] a la vista, els enllacos repetits, la manca de
# CONCLUSIONS i de negreta, i el titular al nom del fitxer.
if ((Test-Path -LiteralPath $llicPathX) -and (Test-Path -LiteralPath (Join-Path $EstructuralsDir 'REQ1.json'))) {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $llicG = Read-LlicCataleg $llicPathX
    $idxG  = _LlicIndexReq1 (Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json'))
    $bAG = (_LlicPuntsPerBloc $llicG $idxG 'ABANS').Punts
    $bDG = (_LlicPuntsPerBloc $llicG $idxG 'DESPRES').Punts
    # Word simulat: nomes el que Build-LlicenciaDocument li demana.
    $selG = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 0 } }
    $selG | Add-Member ScriptMethod EndKey { param($u) } -Force
    $selG | Add-Member ScriptMethod InsertBreak { param($b) } -Force
    $docG = [pscustomobject]@{}
    $docG | Add-Member ScriptMethod Activate {} -Force
    $docG | Add-Member ScriptMethod Save {} -Force
    $docG | Add-Member ScriptMethod Close { param($x) } -Force
    $wordG = [pscustomobject]@{ Selection = $selG }
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_docGlobalProva }
    function Select-CapcaleraBlock($d, $w) { }
    function Apply-HeaderReplacements { param($doc, $header) }
    $script:_docGlobalProva = $docG
    # Build-LlicenciaDocument escriu a $env:TEMP (a Windows sempre hi es; en
    # aquest Linux de proves, no).
    $tempAbans = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }
    $unPunt = @(@($bAG)[0] | ForEach-Object { $_ | Add-Member NoteProperty Estat 'no' -PassThru -Force })
    $unDesp = @(@($bDG)[0] | ForEach-Object { $_ | Add-Member NoteProperty Estat 'no' -PassThru -Force })
    $modelG = @{
        Fase = 'requeriment'; EsProvisional = $false
        Header = @{ ID_GIA = '357'; TITULAR = 'PROVA SL'; CLASSIFICACIO = 'Llei 20/2009; Annex II' }
        Fields = [ordered]@{}
        Abans = $unPunt; Projecte = @(); Despres = $unDesp
        Doc = @{ Text = ''; Items = @() }; Condicions = ''; Cataleg = $llicG
    }
    $global:emitCalls.Clear()
    $petaG = $false
    try { [void](Build-LlicenciaDocument $wordG $modelG) } catch { $petaG = $true; Write-Host ("    EXCEPCIO: " + $_.Exception.Message) -ForegroundColor Red }
    AssertEq $petaG $false 'Build-LlicenciaDocument: genera sense petar'
    $emG = @($global:emitCalls)
    Assert ($emG.Count -gt 5) 'Build-LlicenciaDocument: escriu el document'
    # Cap marcador [[URL]] a la VISTA: els enllacos van per Format-Url.
    Assert (-not (@($emG) | Where-Object { $_ -like '*`[`[URL`]`]*' })) 'Llicencia: cap [[URL]] al text de l''informe'
    Assert ([bool](@($emG) | Where-Object { $_ -like 'URL|*' })) 'Llicencia: els enllacos surten com a Format-Url'
    # Els items van numerats "N." (amb punt), com a REQ1.
    $itemsG = @($emG | Where-Object { $_ -like 'ITEM|*' })
    Assert ($itemsG.Count -ge 2) 'Llicencia: hi ha items numerats'
    Assert ([bool](($itemsG[0] -split '\|')[1] -match '^\d+\.$')) 'Llicencia: el numero de l''item porta punt (1., 2....)'
    # CONCLUSIONS centrat i en negreta, i la conclusio en negreta.
    $iCapG = [Array]::FindIndex([string[]]$emG, [Predicate[string]]{ param($x) $x -like 'CONCLCAP|*' })
    $iConG = [Array]::FindIndex([string[]]$emG, [Predicate[string]]{ param($x) $x -like 'CONCL|*' })
    Assert ($iCapG -ge 0) 'Llicencia: hi ha la capcalera CONCLUSIONS (centrada i en negreta)'
    Assert ($iConG -gt $iCapG) 'Llicencia: ...i va ABANS de la conclusio'
    Assert ([bool]($emG[$iConG] -like '*`*`**')) 'Llicencia: la conclusio va en negreta'
    # Cap enllac repetit DINS del mateix punt (el text de REQ1 i el comentari
    # solen portar el mateix, i sortia dues vegades seguides).
    $urlsG = @($emG | Where-Object { $_ -like 'URL*|*' } | ForEach-Object { ($_ -split '\|', 2)[1] })
    AssertEq (@($urlsG).Count) (@($urlsG | Select-Object -Unique).Count) 'Llicencia: cap enllac repetit'
    # Cap marcador de camp pot arribar al document (els camps es resolen per
    # BLOC, no linia a linia: un [OPCIO:] pot ocupar dos paragrafs del cataleg).
    Assert (-not (@($emG) | Where-Object { $_ -match '\[OPCIO:|\[CAMP:' })) 'Llicencia: cap [OPCIO:]/[CAMP:] literal al document'
    # LA NUMERACIO CONTINUA: el bloc DESPRES no torna a comencar per 1.
    $nums = @($emG | Where-Object { $_ -like 'ITEM|*' } | ForEach-Object { [int](($_ -split '\|')[1] -replace '\.', '') })
    Assert ($nums.Count -ge 2) 'Llicencia: hi ha prou items per comprovar la numeracio'
    $trencats = @()
    for ($i = 1; $i -lt $nums.Count; $i++) { if ($nums[$i] -ne ($nums[$i - 1] + 1)) { $trencats += ("$($nums[$i-1])->$($nums[$i])") } }
    AssertEq ($trencats -join ',') '' 'Llicencia: la numeracio va seguida de cap a peus (DESPRES no reinicia)'
    # ...i el mateix informe com a LLICENCIA PROVISIONAL, que hi afegeix
    # l'ANNEX 1. Ha de sortir en TEXT PLA: cap pic, cap item numerat, negreta
    # nomes als dos titols, i el full de signatures a part i a cos 9.
    # ---- LES TRES FASES, EL MATEIX DOCUMENT ------------------------------
    # Es el que va ensenyar l'usuari comparant el generat amb el fet a ma: el
    # favorable POST no es un informe curt, es el mateix informe sencer amb
    # "Es disposa..." a cada punt del bloc DESPRES.
    $modelF = @{
        Fase = 'requeriment'; EsProvisional = $false
        Header = @{ ID_GIA = '357'; TITULAR = 'PROVA SL' }
        Fields = [ordered]@{}
        Abans = $unPunt; Projecte = @()
        Doc = @{ Text = 'Documentacio signada pel tecnic X.'; Items = @('Projecte (Id Firmadoc: 1)') }
        Condicions = ''; Cataleg = $llicG
    }
    foreach ($fs in @('requeriment', 'favorable-pre', 'favorable-post')) {
        $modelF.Fase = $fs
        $ef = _LlicEstatDespres $fs
        $dsp = @(_LlicPuntsAmbEstatFase (@($bDG)[0]) $fs)
        $modelF.Despres = @($dsp | ForEach-Object { $_ | Add-Member NoteProperty Estat ([string]$ef.Estat) -PassThru -Force })
        $global:emitCalls.Clear()
        try { [void](Build-LlicenciaDocument $wordG $modelF) } catch { Write-Host ("    EXCEPCIO ($fs): " + $_.Exception.Message) -ForegroundColor Red }
        $emF = @($global:emitCalls)
        # La documentacio del projecte va DALT DE TOT i fora de la numeracio.
        # BLOC| = Format-BlockTitle: el nivell de mes amunt de l'informe (majuscules
        # i subratllat). Les seccions de REQ1 que hi van a dins son SECT|.
        $iProj = [Array]::FindIndex([string[]]$emF, [Predicate[string]]{ param($x) $x -like 'BLOC|DOCUMENTACI* PROJECTE' })
        $iAb   = [Array]::FindIndex([string[]]$emF, [Predicate[string]]{ param($x) $x -like 'BLOC|*ABANS*' })
        $iDe   = [Array]::FindIndex([string[]]$emF, [Predicate[string]]{ param($x) $x -like 'BLOC|*DESPR*' })
        # LA DOCUMENTACIO DEL PROJECTE NOMES VA ALS FAVORABLES, i alli dalt de tot.
        # Al REQUERIMENT no hi va: s'hi demanen modificacions al projecte i als
        # planols, o sigui que aquella documentacio encara no es definitiva.
        if (_LlicPortaDocProjecte $fs) {
            Assert ($iProj -ge 0 -and $iProj -lt $iAb) ($fs + ': la documentacio del projecte va la primera')
        } else {
            AssertEq $iProj -1 ($fs + ': al requeriment NO hi va la documentacio del projecte')
        }
        Assert ($iAb -ge 0 -and $iAb -lt $iDe) ($fs + ': ...despres ABANS i despres DESPRES')
        Assert (-not (@($emF) | Where-Object { $_ -like 'SUB|Documentaci*' })) ($fs + ': ja no hi ha el subtitol "Documentacio"')
        # El bloc DESPRES: primer el "Quan:", despres si es disposa o no.
        $bloc = @($emF[$iDe..($emF.Count - 1)])
        $iQuan = [Array]::FindIndex([string[]]$bloc, [Predicate[string]]{ param($x) $x -like 'BODY|Quan: *' })
        Assert ($iQuan -ge 0) ($fs + ': el bloc DESPRES porta el "Quan:"')
        $iEstat = [Array]::FindIndex([string[]]$bloc, [Predicate[string]]{ param($x) $x -like '*es disposa*' })
        if ([string]$ef.Estat -eq '') {
            AssertEq $iEstat -1 ($fs + ': al requeriment no es diu si es disposa o no')
        } else {
            Assert ($iEstat -gt $iQuan) ($fs + ': "es disposa" va DESPRES del "Quan:"')
        }
        # La numeracio va seguida de cap a peus, tambe al post.
        $numsF = @($emF | Where-Object { $_ -like 'ITEM|*' } | ForEach-Object { [int](($_ -split '\|')[1] -replace '\.', '') })
        AssertEq ($numsF -join ',') ((1..$numsF.Count) -join ',') ($fs + ': la numeracio va seguida')
    }
    # El pre-llicencia diu que FALTA (negreta) i el post que ja hi es (normal).
    $modelF.Fase = 'favorable-pre'
    $modelF.Despres = @(@(_LlicPuntsAmbEstatFase (@($bDG)[0]) 'favorable-pre') | ForEach-Object { $_ | Add-Member NoteProperty Estat 'no' -PassThru -Force })
    $global:emitCalls.Clear()
    [void](Build-LlicenciaDocument $wordG $modelF)
    Assert ([bool](@($global:emitCalls) | Where-Object { $_ -like 'BODY/N/SEP|No es disposa de la documentaci*' })) 'favorable-pre: "No es disposa de la documentacio.", en negreta i separada'
    $modelF.Fase = 'favorable-post'
    $modelF.Despres = @(@(_LlicPuntsAmbEstatFase (@($bDG)[0]) 'favorable-post') | ForEach-Object { $_ | Add-Member NoteProperty Estat 'si' -PassThru -Force })
    $global:emitCalls.Clear()
    [void](Build-LlicenciaDocument $wordG $modelF)
    $emPost = @($global:emitCalls)
    Assert ([bool]($emPost | Where-Object { $_ -like 'BODY/SEP|Es disposa del document*' })) 'favorable-post: "Es disposa del document", separada i SENSE negreta'
    Assert (-not ($emPost | Where-Object { $_ -match '\[CAMP:' })) 'favorable-post: cap marcador de camp literal'
    Assert (-not ($emPost | Where-Object { $_ -like '*haver comprovat la seg*ent documentaci*' })) 'favorable-post: ja no es un informe a part'

    $modelG.EsProvisional = $true
    $global:emitCalls.Clear()
    try { [void](Build-LlicenciaDocument $wordG $modelG) } catch { Write-Host ("    EXCEPCIO (annex): " + $_.Exception.Message) -ForegroundColor Red }
    $emA = @($global:emitCalls)
    $iAnnex = [Array]::FindIndex([string[]]$emA, [Predicate[string]]{ param($x) $x -like 'PLA*|ANNEX 1*' })
    Assert ($iAnnex -ge 0) 'ANNEX 1: hi surt (nomes a la llicencia provisional)'
    $annex = @($emA[$iAnnex..($emA.Count - 1)])
    Assert (-not (@($annex) | Where-Object { $_ -like 'BULLET*' })) 'ANNEX 1: cap pic (text pla)'
    Assert (-not (@($annex) | Where-Object { $_ -like 'ITEM|*' })) 'ANNEX 1: cap item numerat (text pla)'
    Assert (-not (@($annex) | Where-Object { $_ -like 'BODY*' })) 'ANNEX 1: tot passa per Format-Plain'
    $negretes = @($annex | Where-Object { $_ -like 'PLA/N*' })
    AssertEq $negretes.Count 2 'ANNEX 1: negreta NOMES als dos titols'
    Assert ([bool]($negretes[0] -like '*ANNEX 1*')) 'ANNEX 1: el primer titol en negreta'
    Assert ([bool]($negretes[1] -like '*acceptaci*')) 'ANNEX 1: i el del full de signatures'
    $cos9 = @($annex | Where-Object { $_ -like '*/sz9|*' })
    Assert ($cos9.Count -ge 5) 'ANNEX 1: el full de signatures va a cos 9'
    # ELS NUMEROS I ELS GUIONS, com a TEXT (l'original els porta amb numeracio
    # automatica del Word; aqui van escrits al davant i sense sagnia).
    $numerats = @($annex | Where-Object { $_ -match '^PLA[^|]*\|\d+\. ' })
    Assert ($numerats.Count -ge 4) ('ANNEX 1: els punts van numerats "1. ", "2. "... (n''hi ha ' + $numerats.Count + ')')
    $primerNum = [int]((($numerats[0] -split '\|', 2)[1] -split '\.')[0])
    AssertEq $primerNum 1 'ANNEX 1: la numeracio comenca per 1'
    $guionats = @($annex | Where-Object { $_ -match '^PLA[^|]*\|- ' })
    Assert ($guionats.Count -ge 7) ('ANNEX 1: els sub-punts van amb guio (n''hi ha ' + $guionats.Count + ')')
    # El full de signatures NO porta ni numero ni guio.
    Assert (-not (@($annex | Where-Object { $_ -like '*/sz9|*' }) | Where-Object { $_ -match '\|(\d+\.|-) ' })) 'ANNEX 1: el full de signatures va sense marques'
    Assert (-not (@($annex[0..($annex.Count - $cos9.Count - 1)]) | Where-Object { $_ -like '*sz9*' })) 'ANNEX 1: ...i NOMES el full de signatures'
    $env:TEMP = $tempAbans
}
# L'ENLLAC VA DESPRES DE LA FRASE QUE L'ANUNCIA. El comentari acaba amb
# "...en el seguent enllac:" i el cos de l'item (de REQ1) sol portar EL MATEIX
# enllac: sortia abans, amb la frase penjada sense res al darrere.
if (Test-Path -LiteralPath $llicPathX) {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $selU = [pscustomobject]@{}
    $L1 = 'https://exemple.cat/tramit'
    $puntU = [pscustomobject]@{
        Clau = ''; Titol = 'Incendis'; Condicio = ''
        Cos = @('Incendis. S''ha d''obtenir l''informe.', ('[[URL]] ' + $L1))
        NoDisposa = @('No es disposa de l''informe. S''ha de sol·licitar en el seguent enllac:', ('[[URL]] ' + $L1))
        SiDisposa = @(); Quan = @(); Subs = @()
    }
    $global:emitCalls.Clear()
    _LlicEscriuPunt $selU $puntU '2.' ([ordered]@{}) 'no' $false
    $emU = @($global:emitCalls)
    $iCom = [Array]::FindIndex([string[]]$emU, [Predicate[string]]{ param($x) $x -like 'BODY*|No es disposa*' })
    $iUrl = [Array]::FindIndex([string[]]$emU, [Predicate[string]]{ param($x) $x -like 'URL*' })
    Assert ($iCom -ge 0) 'enllac: hi ha el comentari'
    Assert ($iUrl -gt $iCom) 'enllac: va DESPRES de la frase que l''anuncia (no abans)'
    AssertEq (@($emU | Where-Object { $_ -like 'URL*' }).Count) 1 'enllac: nomes una vegada, encara que sigui als dos textos'
    # Si el comentari NO porta enllac, el de l'item ha de sortir amb l'item.
    $puntU2 = [pscustomobject]@{
        Clau = ''; Titol = 'X'; Condicio = ''
        Cos = @('Text de l''item.', ('[[URL]] ' + $L1))
        NoDisposa = @('No es disposa.'); SiDisposa = @(); Quan = @(); Subs = @()
    }
    $global:emitCalls.Clear()
    _LlicEscriuPunt $selU $puntU2 '1.' ([ordered]@{}) 'no' $false
    $emU2 = @($global:emitCalls)
    $iUrl2 = [Array]::FindIndex([string[]]$emU2, [Predicate[string]]{ param($x) $x -like 'URL*' })
    $iCom2 = [Array]::FindIndex([string[]]$emU2, [Predicate[string]]{ param($x) $x -like 'BODY*|No es disposa*' })
    Assert ($iUrl2 -ge 0 -and $iUrl2 -lt $iCom2) 'enllac: si el comentari no en porta, el de l''item surt amb l''item'
}

# El nom del fitxer NO porta el titular (l'usuari no el vol; ja surt a dins).
$nfG = _LlicNomFitxer ([datetime]'2026-08-18') 'requeriment' '1457'
AssertEq $nfG '2026-08-18_LlicReq_GIA 1457.docx' '_LlicNomFitxer: el nom no porta el titular'

Write-Host "`n--- PdfCms.ps1: refer la signatura de dins del PDF ---"
# Els informes signats amb l'AutoFirma nomes es validaven a l'ordinador que
# signa. Comparant amb un de signat a ma amb l'Adobe (mateix certificat, mateix
# ordinador) l'unica diferencia que quedava era l'atribut ESS
# 'signingCertificateV2', que porta un camp 'policies' i, pel RFC 5035, obliga a
# validar la cadena RESTRINGIDA a aquelles politiques. Aqui es refa el CMS amb
# la mateixa estructura que l'Adobe, sense tocar ni un byte del document.
. (Join-Path (Split-Path -Parent $TestsDir) 'PdfCms.ps1')

# Un PDF de mentida, pero amb un /ByteRange i un forat de /Contents de veritat.
$pdfCap  = [System.Text.Encoding]::ASCII.GetBytes('%PDF-1.7' + "`n" + '/ByteRange[0 40 100 20]' + "`n")
$pdfFals = New-Object byte[] 120
for ($i = 0; $i -lt $pdfCap.Length; $i++) { $pdfFals[$i] = $pdfCap[$i] }
$pdfFals[40] = [byte][char]'<'
for ($i = 41; $i -lt 99; $i++) { $pdfFals[$i] = [byte][char]'0' }
$pdfFals[99] = [byte][char]'>'
$fCap = _PdfTrobaFirma $pdfFals
AssertEq ([bool]$fCap.Ok) $true '_PdfTrobaFirma: troba la signatura'
AssertEq (@($fCap.Ranges) -join ',') '0,40,100,20' '_PdfTrobaFirma: el /ByteRange'
AssertEq $fCap.HexStart 41 '_PdfTrobaFirma: el forat comenca despres del <'
AssertEq $fCap.HexLen   58 '_PdfTrobaFirma: i acaba abans del >'
# El contingut signat son els DOS trossos, seguits.
$cSig = _PdfContingutSignat $pdfFals $fCap.Ranges
AssertEq $cSig.Length 60 '_PdfContingutSignat: els dos trossos del /ByteRange'
# Un PDF sense signar no s'ha de confondre amb un de signat.
$senseFirma = [System.Text.Encoding]::ASCII.GetBytes('%PDF-1.7 res a veure aqui dins, nomes text i mes text per fer bulto')
AssertEq ([bool](_PdfTrobaFirma $senseFirma).Ok) $false '_PdfTrobaFirma: un PDF sense signar -> no'
AssertEq ([bool](_PdfTrobaFirma ([byte[]]@())).Ok) $false '_PdfTrobaFirma: fitxer buit -> no'
# Escriure el CMS: la MIDA DEL FITXER NO POT CANVIAR (el /ByteRange ja esta
# escrit i firmat; si el fitxer creix o minva, deixa de quadrar).
$cmsFals = [byte[]](0xDE, 0xAD, 0xBE, 0xEF)
$posat = _PdfPosaCms $pdfFals $fCap.HexStart $fCap.HexLen $cmsFals
AssertEq $posat.Length $pdfFals.Length '_PdfPosaCms: la mida del fitxer no canvia'
$hexPosat = [System.Text.Encoding]::ASCII.GetString($posat, $fCap.HexStart, 8)
AssertEq $hexPosat 'deadbeef' '_PdfPosaCms: el CMS hi va en hexadecimal'
AssertEq ([System.Text.Encoding]::ASCII.GetString($posat, $fCap.HexStart + 8, 4)) '0000' '_PdfPosaCms: la resta del forat, farcida de zeros'
AssertEq ([char]$posat[40]) '<' '_PdfPosaCms: no es toca el < d''obrir'
AssertEq ([char]$posat[99]) '>' '_PdfPosaCms: ni el > de tancar'
# ...i el DOCUMENT (el que hi ha fora del forat) ha de quedar intacte.
$igualFora = $true
for ($i = 0; $i -lt $pdfFals.Length; $i++) {
    if ($i -ge $fCap.HexStart -and $i -lt ($fCap.HexStart + $fCap.HexLen)) { continue }
    if ($pdfFals[$i] -ne $posat[$i]) { $igualFora = $false; break }
}
AssertEq $igualFora $true '_PdfPosaCms: fora del forat no es toca ni un byte'
# Un CMS que no hi cap ha de PETAR, no escriure a mitges.
$petat = $false
try { [void](_PdfPosaCms $pdfFals $fCap.HexStart $fCap.HexLen (New-Object byte[] 500)) } catch { $petat = $true }
AssertEq $petat $true '_PdfPosaCms: si el CMS no hi cap, peta (no escriu a mitges)'

# GEOMETRIA del dialeg "Convertir informes a PDF". Aixo es una prova de FONT
# (la finestra nomes es pot dibuixar a Windows), pero enxampa el que va passar
# de debo: en afegir-hi els dos radios del mode de signatura, els botons
# 'Comenca'/'Tanca' -clavats a y=438- van quedar FORA de la finestra, i el boto
# 'Document' de la fila de dalt (que acaba a x=536) ja sortia tallat amb els
# 510 px d'amplada que hi havia.
$srcPdf = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'PdfSignar.ps1') -Raw
# NOMES el tros del dialeg d'OPCIONS: la finestra de PROGRES que ve despres es
# de mida fixa i alli no hi ha res a dir.
$iDlg = $srcPdf.IndexOf('$row = _AddConfigRow $form 70')
$iFi  = $srcPdf.IndexOf("_AddBrandHeader `$form 'Convertir informes a PDF'")
$dlgPdf = $srcPdf.Substring($iDlg, $iFi - $iDlg)
Assert ($Script:PdfDlgAmple -ge 550) 'dialeg PDF: l''amplada dona per al boto Document (x fins 536)'
Assert (-not ($dlgPdf -match '\$btnGo\.Location = New-Object System\.Drawing\.Point\(\d+, \d+\)')) 'dialeg PDF: el boto Comenca NO va a una alcada fixa'
Assert (-not ($dlgPdf -match '\$btnCancel\.Location = New-Object System\.Drawing\.Point\(\d+, \d+\)')) 'dialeg PDF: ni el boto Tanca'
Assert ($dlgPdf.Contains('$yBotons = $y')) 'dialeg PDF: els botons es col·loquen a partir de l''ultim control'
Assert ($dlgPdf.Contains('$form.ClientSize = New-Object System.Drawing.Size($Script:PdfDlgAmple, ($yBotons + 44))')) 'dialeg PDF: i la finestra creix segons el contingut'

# On es busca l'Adobe per al mode de signatura a ma. Rutes en text pla (amb
# Join-Path petarien fora de Windows per la unitat C:).
$adC = @(_AdobeExeCandidats 'C:\PF' 'C:\PF86')
AssertEq $adC.Count 6 '_AdobeExeCandidats: 3 rutes per cada Program Files'
Assert ([bool]($adC[0] -like 'C:\PF\Adobe\*Acrobat.exe')) '_AdobeExeCandidats: primer l''Acrobat complet'
Assert ([bool]($adC -like '*AcroRd32.exe').Count -eq 4) '_AdobeExeCandidats: i els Readers'
AssertEq (@(_AdobeExeCandidats 'C:\PF' '').Count) 3 '_AdobeExeCandidats: sense Program Files (x86), nomes 3'
AssertEq (@(_AdobeExeCandidats '' '').Count) 0 '_AdobeExeCandidats: sense res, cap ruta (i cap petada)'

# EL CMS SENCER, amb un certificat EFIMER creat en memoria: es prova el cicle
# complet (crear -> igualar l'OID com l'Adobe -> comprovar -> posar-lo dins d'un
# PDF sintetic i rellegir-lo). Aixi el cami que corre a Windows es exactament el
# que s'ha provat aqui, no una replica.
$certT = $null
try {
    $rsaT = [System.Security.Cryptography.RSA]::Create(2048)
    $reqT = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest ('CN=PROVA CMS', $rsaT, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $certT = $reqT.CreateSelfSigned([System.DateTimeOffset]::Now.AddDays(-1), [System.DateTimeOffset]::Now.AddDays(1))
} catch { $certT = $null }
if ($null -ne $certT) {
    $contT = [System.Text.Encoding]::ASCII.GetBytes('contingut signat de prova, com si fos el PDF')
    $derT = _CmsComAdobe $contT $certT
    Assert ($derT.Length -gt 500) '_CmsComAdobe: genera un CMS'
    # L'OID s'iguala amb el de l'Adobe: UN byte, la mida no es mou, i el CMS
    # SEGUEIX VERIFICANT (el primer intent d'aixo corrompia el certificat).
    $derP = _CmsOidComAdobe $derT
    AssertEq $derP.Length $derT.Length '_CmsOidComAdobe: la mida no canvia'
    $difT = 0
    for ($i = 0; $i -lt $derT.Length; $i++) { if ($derT[$i] -ne $derP[$i]) { $difT++ } }
    Assert ($difT -le 1) '_CmsOidComAdobe: canvia UN byte com a molt'
    $provaT = _CmsComprova $derP $contT
    AssertEq ([bool]$provaT.Ok) $true ('_CmsComprova: el CMS igualat segueix verificant (' + $provaT.Motiu + ')')
    Assert ([bool]($provaT.Atributs -contains '1.2.840.113549.1.9.4')) '_CmsComprova: hi ha messageDigest'
    Assert ([bool]($provaT.Atributs -contains '1.2.840.113583.1.1.8')) '_CmsComprova: hi ha revocationInfoArchival (com l''Adobe)'
    Assert (-not ($provaT.Atributs -contains '1.2.840.113549.1.9.16.2.47')) '_CmsComprova: CAP atribut ESS (la causa de tot)'
    # La guarda posicional: en un blob que porta l'OID pero cap messageDigest
    # (com dins d'un certificat), NO s'ha de tocar res.
    $blobT = New-Object byte[] 64
    $oidT = [byte[]](0x06,0x09,0x2A,0x86,0x48,0x86,0xF7,0x0D,0x01,0x01,0x01)
    for ($i = 0; $i -lt $oidT.Length; $i++) { $blobT[20 + $i] = $oidT[$i] }
    $blobP = _CmsOidComAdobe $blobT
    AssertEq $blobP[30] $blobT[30] '_CmsOidComAdobe: sense messageDigest al darrere, no toca res (la guarda del certificat)'
    # Si algu toca la SIGNATURA, la comprovacio ho ha de veure.
    $derMal = New-Object byte[] $derP.Length
    [Array]::Copy($derP, $derMal, $derP.Length)
    $derMal[$derMal.Length - 10] = $derMal[$derMal.Length - 10] -bxor 0xFF
    AssertEq ([bool](_CmsComprova $derMal $contT).Ok) $false '_CmsComprova: una signatura tocada NO passa'
    # ...i el CICLE SENCER sobre un PDF sintetic amb un forat de debo.
    $tplT = '%PDF-1.7' + "`n" + '/ByteRange[0 100 4102 30]'
    $pdfT = New-Object byte[] 4132
    $tplB = [System.Text.Encoding]::ASCII.GetBytes($tplT)
    for ($i = 0; $i -lt $tplB.Length; $i++) { $pdfT[$i] = $tplB[$i] }
    for ($i = $tplB.Length; $i -lt 100; $i++) { $pdfT[$i] = 0x20 }
    $pdfT[100] = [byte][char]'<'
    for ($i = 101; $i -lt 4101; $i++) { $pdfT[$i] = [byte][char]'0' }
    $pdfT[4101] = [byte][char]'>'
    for ($i = 4102; $i -lt 4132; $i++) { $pdfT[$i] = 0x20 }
    $tmpPdfT = Join-Path ([System.IO.Path]::GetTempPath()) ('prova-cms-' + [guid]::NewGuid().ToString('N') + '.pdf')
    [System.IO.File]::WriteAllBytes($tmpPdfT, $pdfT)
    $rT = Repack-PdfFirmaComAdobe $tmpPdfT $certT
    AssertEq ([bool]$rT.Ok) $true ('Repack-PdfFirmaComAdobe: el cicle sencer acaba be (' + $rT.Motiu + ')')
    Assert ([bool]($rT.Motiu -like '*COMPROVAT*')) 'Repack-PdfFirmaComAdobe: el motiu diu que s''ha comprovat'
    $b2T = [System.IO.File]::ReadAllBytes($tmpPdfT)
    AssertEq $b2T.Length $pdfT.Length 'Repack-PdfFirmaComAdobe: la mida del PDF no canvia'
    # Es rellegeix el CMS de dins del PDF (la longitud real es treu del DER, no
    # es retallen zeros a cegues: una signatura pot acabar en 00 de debo).
    $hex2T = [System.Text.Encoding]::ASCII.GetString($b2T, 101, 4000)
    $raw2T = New-Object byte[] 2000
    for ($i = 0; $i -lt 2000; $i++) { $raw2T[$i] = [Convert]::ToByte($hex2T.Substring($i * 2, 2), 16) }
    # 48 = 0x30 (SEQUENCE) i 130 = 0x82. EN DECIMAL A POSTA: un literal hex com
    # a argument (AssertEq $x 0x30) arriba com a 48 pero CONSERVA el text del
    # token, i el [string] de dins d'AssertEq el converteix en "0x30" -> la
    # comparacio falla mentre el missatge diu '48' contra '48'. Mateixa familia
    # de trampes que el PSObject del Join-Path (vegeu CLAUDE.md).
    AssertEq $raw2T[0] 48 'Repack-PdfFirmaComAdobe: el CMS de dins comenca amb SEQUENCE'
    AssertEq $raw2T[1] 130 'Repack-PdfFirmaComAdobe: ...amb longitud de dos bytes'
    $lenT = ($raw2T[2] * 256) + $raw2T[3] + 4
    $cms2T = New-Object byte[] $lenT
    [Array]::Copy($raw2T, $cms2T, $lenT)
    $c2T = _PdfContingutSignat $b2T @(0, 100, 4102, 30)
    AssertEq ([bool](_CmsComprova $cms2T $c2T).Ok) $true 'Repack-PdfFirmaComAdobe: el CMS de dins del PDF verifica contra el contingut signat'
    Remove-Item $tmpPdfT -Force -ErrorAction SilentlyContinue
} else {
    Write-Host '  (omes: no es poden crear certificats en memoria en aquesta plataforma)'
}

Write-Host "`n--- SincronitzaCatalegs.ps1: protegir els catalegs en actualitzar ---"
# No el carrega Motor.ps1 (l'executa Actualitzar.bat pel seu compte); el
# dot-sourcegem aqui. El $env:GENINFORME_TEST fa que nomes en surtin definicions.
. (Join-Path (Split-Path -Parent $TestsDir) 'SincronitzaCatalegs.ps1')
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/REQ1.json')) $true '_CatalegEsProtegible: cataleg json'
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/0 CAPCALERA.docx')) $true '_CatalegEsProtegible: la plantilla de la capcalera'
AssertEq ([bool](_CatalegEsProtegible 'docs/dades/email-textos.json')) $true '_CatalegEsProtegible: dades del mobil'
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/REQ1.json.bak')) $false '_CatalegEsProtegible: els .bak de l''editor NO'
AssertEq ([bool](_CatalegEsProtegible 'suport/Motor.ps1')) $false '_CatalegEsProtegible: el codi NO'
AssertEq ([bool](_CatalegEsProtegible '')) $false '_CatalegEsProtegible: buit -> no'

# COL·LISIO en un fitxer BINARI. Historia real: l'usuari tenia '0 CAPCALERA.docx'
# retocat, el repositori hi acabava d'afegir el bloc [[CAP:LLIC]], el rebase va
# petar (binari: no es pot fusionar), la seva copia es va tornar a aplicar a
# sobre i la versio SENSE el bloc es va pujar a main. Els .json no tenen aquest
# problema: alli l'usuari mana i com a molt es torna a escriure un text.
AssertEq ([bool](_CatalegEsBinari 'ESTRUCTURALS/0 CAPCALERA.docx')) $true  '_CatalegEsBinari: la plantilla de la capcalera'
AssertEq ([bool](_CatalegEsBinari 'ESTRUCTURALS/REQ1.json'))        $false '_CatalegEsBinari: un cataleg json no'
AssertEq ([bool](_CatalegEsBinari 'docs/dades/email-textos.json'))  $false '_CatalegEsBinari: dades del mobil no'
AssertEq ([bool](_CatalegHiHaColisio 'aaa' 'bbb' $true))  $true  '_CatalegHiHaColisio: binari que ha canviat a les dues bandes'
AssertEq ([bool](_CatalegHiHaColisio 'aaa' 'aaa' $true))  $false '_CatalegHiHaColisio: binari que el repositori NO ha tocat'
AssertEq ([bool](_CatalegHiHaColisio 'aaa' 'bbb' $false)) $false '_CatalegHiHaColisio: als .json l''usuari mana sempre'
# Sense sha de base no se sap: val mes tornar a aplicar el de l'usuari (el
# comportament de sempre) que descartar-li la feina per un dubte.
AssertEq ([bool](_CatalegHiHaColisio '' 'bbb' $true))     $false '_CatalegHiHaColisio: sense base, no es decideix en contra de l''usuari'
AssertEq ([bool](_CatalegHiHaColisio 'aaa' '' $true))     $false '_CatalegHiHaColisio: sense el sha d''ara, tampoc'
# El Backup ha d'apuntar el commit de base: sense ell, el Restore no pot saber
# si el repositori ha tocat el mateix fitxer.
$syncSrc = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'SincronitzaCatalegs.ps1') -Raw
Assert ($syncSrc.Contains('base.txt')) 'SincronitzaCatalegs: la copia apunta el commit de base'
Assert ($syncSrc.Contains('git rev-parse HEAD')) 'SincronitzaCatalegs: el commit de base es llegeix al Backup (abans del pull)'
# I l'Actualitzar.bat ha de mirar el codi 2 per tornar a avisar al final.
$batCol = Get-Content -LiteralPath (Join-Path $RepoRoot 'Actualitzar.bat') -Raw
Assert ($batCol.Contains('if errorlevel 2 set "COLISIO_CATALEGS=1"')) 'Actualitzar.bat: recull la col·lisio del Restore'
Assert ($batCol.Contains('if "%COLISIO_CATALEGS%"=="1"')) 'Actualitzar.bat: torna a avisar de la col·lisio al final'

# La copia de seguretat ja protegia '0 CAPCALERA.docx' (proves de dalt), pero
# Actualitzar.bat nomes ESTADIAVA els *.json, o sigui que un canvi a la capcalera
# no es commitejava MAI: el 'pull --rebase' es negava a comencar, s'anava al cami
# d'error i el 'reset --hard' se l'enduia. Nomes es salvava per la copia. El
# 'git add -u -- ESTRUCTURALS' estadia qualsevol fitxer JA SEGUIT d'alli (i no
# 'ESTRUCTURALS/*.docx' a piu, que en un clone antic pujaria les vistes en Word
# que encara no s'han migrat a 'local\').
$batAct = Get-Content -LiteralPath (Join-Path $RepoRoot 'Actualitzar.bat') -Raw
$batAdds = @([regex]::Matches($batAct, [regex]::Escape('git add -u -- ESTRUCTURALS')))
AssertEq $batAdds.Count 2 'Actualitzar.bat: estadia els fitxers seguits d''ESTRUCTURALS als DOS commits (abans i despres del pull)'
Assert (-not ($batAct -like '*git add "ESTRUCTURALS/*.docx"*')) 'Actualitzar.bat: NO estadia els .docx a piu (les vistes velles no s''han de pujar)'
# El reset --hard esborra el que no s'ha commitejat: no hi pot arribar res sense
# haver-ho desat al stash abans.
$iStash = $batAct.IndexOf('Actualitzar.bat: abans de posar-me al dia')
$iReset = $batAct.IndexOf('git reset --hard origin/main')
Assert ([bool]($iStash -gt 0 -and $iReset -gt $iStash)) 'Actualitzar.bat: el stash de seguretat va ABANS del reset --hard'
# Sortida real de 'git status --porcelain' (git enquota els noms amb espais).
$gs = @(
    ' M "ESTRUCTURALS/0 CONCLUSIONS.json"',
    ' M ESTRUCTURALS/REQ1.json',
    ' M docs/dades/email-textos.json',
    '?? ESTRUCTURALS/NOU.json',
    '?? ESTRUCTURALS/REQ1.json.bak',
    ' M suport/Motor.ps1'
)
$paths = @(_ParseGitStatusPaths $gs)
AssertEq $paths.Count 4 '_ParseGitStatusPaths: 4 protegits (fora .bak i codi)'
AssertEq ([bool]($paths -contains 'ESTRUCTURALS/0 CONCLUSIONS.json')) $true '_ParseGitStatusPaths: treu les cometes dels noms amb espais'
AssertEq ([bool]($paths -contains 'ESTRUCTURALS/NOU.json')) $true '_ParseGitStatusPaths: inclou els fitxers nous (??)'
AssertEq ([bool]($paths -contains 'suport/Motor.ps1')) $false '_ParseGitStatusPaths: exclou el codi'
AssertEq (@(_ParseGitStatusPaths @(' M ESTRUCTURALS/REQ1.json -> ESTRUCTURALS/NOU.json')).Count) 1 '_ParseGitStatusPaths: canvi de nom -> es queda amb el desti'
AssertEq (_CatalegsBackupName ([datetime]'2026-07-28 11:39:41')) '20260728-113941' '_CatalegsBackupName: nom de carpeta per data'
