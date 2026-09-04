# Guards de font: el que no pot tornar a passar
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

Write-Host "`n--- El menu: ordre dels informes ---"
$segSrc = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1') -Raw
$q = [char]39
$rxMenu = '\$menu\.Add\(@\{\s*Action\s*=\s*' + $q + '([a-z]+)' + $q
$accions = @([regex]::Matches($segSrc, $rxMenu) | ForEach-Object { $_.Groups[1].Value })
# Les SIS entrades fixes, en ordre. La setena 'nou' que hi ha al codi es el
# bucle del final ("qualsevol altre cataleg no llistat s'afegeix al final") i
# no forma part de l'ordre.
$ordreEsperat = @('nou', 'seguiment', 'llicencia', 'actextr', 'mnstraspas', 'nou')
AssertEq (@($accions | Select-Object -First 6) -join ',') ($ordreEsperat -join ',') 'Menu: Nou, Seguiment, Llicencia, Act. extraordinaries, MNS/Traspas, Ampliacio termini'
# MNS/Traspas ha de tenir entrada propia i el seu despatx.
Assert ([bool]($accions -contains 'mnstraspas')) 'Menu: MNS/Traspas te entrada propia (ja no s''hi arriba des de dins de Llicencia)'
$wizSrc = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1') -Raw
Assert ([bool]($wizSrc -match ($q + 'mnstraspas' + $q + '\s*\{'))) 'Menu: ...i el despatxador la coneix'
# I cada entrada ofereix NOMES les seves fases.
AssertEq (@(_LlicFases).Count) 3 'Llicencia: tres fases (requeriment i els dos favorables)'
AssertEq (@(_MnsFases).Count) 2  'MNS/Traspas: dues fases'
# QUINA FASE SURT MARCADA. Aqui hi havia un 'requeriment' escrit al codi com a
# respatller i, des que MNS/Traspas te entrada propia -i la seva llista no en te
# cap-, la pantalla petava amb "La propiedad 'Checked' no se encuentra en este
# objeto". Cap llista de fases pot donar per fet quines fases porta.
AssertEq (_LlicFasePerDefecte (_LlicFases) 'requeriment') 'requeriment' 'Fase per defecte: si la preferida hi es, aquella'
AssertEq (_LlicFasePerDefecte (_MnsFases) 'requeriment') 'mns' 'Fase per defecte: si NO hi es, la primera de la llista'
AssertEq (_LlicFasePerDefecte (_MnsFases) 'traspas') 'traspas' 'Fase per defecte: respecta el que ja s''havia triat'
AssertEq (_LlicFasePerDefecte @() 'requeriment') '' 'Fase per defecte: llista buida, cadena buida (i no peta)'
AssertEq (_LlicFasePerDefecte (_LlicFases) '') 'requeriment' 'Fase per defecte: sense preferencia, la primera'
# I que no torni a apareixer cap clau de fase escrita a pel com a respatller.
$rxRad = '\$radios\[' + $q + '[a-z-]+' + $q + '\]'
$radLit = @()
foreach ($ln in ((Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Llicencia.ps1') -Raw) -split "`r?`n")) {
    if ($ln.TrimStart().StartsWith('#')) { continue }   # els comentaris no compten
    if ([regex]::IsMatch($ln, $rxRad)) { $radLit += $ln.Trim() }
}
AssertEq $radLit.Count 0 ('Select-LlicFase: cap radio buscat per una clau escrita al codi' + $(if ($radLit.Count) { ' -> ' + ($radLit -join ' | ') } else { '' }))
AssertEq (@(_LlicTotesLesFases).Count) 5 '_LlicTotesLesFases: segueix ajuntant-les (la fan servir la base de dades i la vista)'
foreach ($f in @(_MnsFases)) {
    Assert (-not (@(_LlicFases) | Where-Object { [string]$_.Clau -eq [string]$f.Clau })) ('Llicencia no ofereix la fase "' + $f.Clau + '"')
}

# ============================================================================
# GUARDS AFEGITS DESPRES DE L'AUDITORIA D'ARQUITECTURA
# ============================================================================

Write-Host "`n--- Tots els .ps1 PARSEGEN (guard) ---"
# PER QUE. RecordatorisAuto.ps1 va estar-se amb un ParserError: hi havia
# "$clau: ..." dins d'una cadena, i alli els dos punts els menja el parser com a
# prefix d'ambit ($global:, $env:...). El fitxer NO es podia carregar de cap
# manera, o sigui que el mode automatic dels recordatoris -el que corre sol des
# d'una tasca del Windows- no havia funcionat mai. Cap prova el tocava perque
# cap prova el CARREGAVA. Aquest guard els carrega tots, sense executar-ne res.
$rootRepo = Split-Path -Parent (Split-Path -Parent $TestsDir)
$ps1Tots  = @(Get-ChildItem -Path (Join-Path $rootRepo 'suport') -Recurse -Filter *.ps1 -File)
Assert ($ps1Tots.Count -gt 40) "el guard de parseig veu tots els .ps1 ($($ps1Tots.Count))"
$ambErrors = @()
foreach ($f in $ps1Tots) {
    $errsP = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errsP)
    if ($errsP -and $errsP.Count -gt 0) {
        $ambErrors += ('{0}:{1} {2}' -f $f.Name, $errsP[0].Extent.StartLineNumber, $errsP[0].Message)
    }
}
AssertEq $ambErrors.Count 0 ('cap .ps1 amb errors de sintaxi' + $(if ($ambErrors.Count) { ' -> ' + ($ambErrors -join ' | ') } else { '' }))

Write-Host "`n--- Tots els .ps1 tenen BOM (guard) ---"
# PER QUE. El Windows PowerShell 5.1 llegeix un .ps1 SENSE BOM com a ANSI i
# corromp els literals accentuats. Precintades.ps1 documentava aquesta trampa i
# la incomplia ell mateix: el missatge d'error de la fulla "Estes"/"Estes"
# sortia amb mojibake a l'usuari. Amb BOM a tots, no pot tornar a passar.
$senseBom = @()
foreach ($f in $ps1Tots) {
    $b3 = [System.IO.File]::ReadAllBytes($f.FullName)
    if ($b3.Length -lt 3 -or $b3[0] -ne 0xEF -or $b3[1] -ne 0xBB -or $b3[2] -ne 0xBF) { $senseBom += $f.Name }
}
AssertEq $senseBom.Count 0 ('cap .ps1 sense BOM' + $(if ($senseBom.Count) { ' -> ' + ($senseBom -join ', ') } else { '' }))

Write-Host "`n--- docs/app.js: paritat amb el motor del PC (guards) ---"
# PER QUE. app.js diu que "replica EXACTAMENT" la logica del PC, pero no hi ha
# cap prova que ho comprovi i havia divergit en quatre punts, tots ells defectes
# que el PC ja havia patit, diagnosticat i corregit. Aquests guards no proven el
# comportament del JavaScript (no hi ha arnes JS al projecte); vigilen que els
# quatre arranjaments no es desfacin sense adonar-se'n.
$appJs = Get-Content -LiteralPath (Join-Path $rootRepo (Join-Path 'docs' 'app.js')) -Raw

# 1) Una opcio BUIDA es una opcio (Camps.ps1:45-53). Filtrar-la treia la manera
#    de dir "res" i el text sortia SEMPRE.
Assert (-not ($appJs.Contains('if (o !== "") opts.push(o)'))) 'app.js: parseOpcio no filtra les opcions buides'
Assert ($appJs.Contains('OPCIO_ETIQUETA_BUIDA')) 'app.js: hi ha etiqueta per a l''opcio buida'
Assert ($appJs.Contains('op.textContent = opcioEtiqueta(o)')) 'app.js: el desplegable pinta l''etiqueta i desa el valor'

# 2) Els camps es resolen PER BLOC (Apply-FieldsToLines, Camps.ps1:132-149). Un
#    [OPCIO:] partit per un Enter no es resol linia a linia i surt CRU al correu.
Assert ($appJs.Contains('function applyFieldsToLines')) 'app.js: existeix applyFieldsToLines'
Assert (-not ($appJs.Contains('applyFields(l, values)'))) 'app.js: cap camp resolt linia a linia'

# 3) richTextOf uneix amb SALT DE LINIA (_RichTextOfBodyLines, Camps.ps1:327-337):
#    el que es detecta i el que es resol han de veure el mateix text.
Assert ($appJs.Contains('return parts.join("\n");')) 'app.js: richTextOf uneix amb salt de linia'

# 4) Un text fix es de la SECCIO o de la SUBSECCIO segons on estigui
#    (Build-CatalegBlocs, MotorInforme.ps1:373-412). Hi ha DUES copies del
#    recorregut del cataleg dins d'app.js i totes dues ho han de fer igual.
AssertEq ([regex]::Matches($appJs, 'introSeccio').Count -gt 0) $true 'app.js: distingeix l''intro de SECCIO'
AssertEq ([regex]::Matches($appJs, 'dinsSub = true').Count) 2 'app.js: les DUES copies del recorregut marquen dinsSub'
AssertEq ([regex]::Matches($appJs, 'introSeccio = el').Count) 2 'app.js: les DUES copies recullen l''intro de seccio'

# 4. CODI MORT I DEFECTES SILENCIOSOS D'app.js (bloc P1 de l'auditoria).
#    Els quatre venien de la mateixa familia: no peten, nomes fan una cosa
#    diferent de la que sembla, i cap prova de PowerShell mira el comportament
#    del JavaScript. Per aixo hi arriba el text del fitxer.
#
# 4a. 'prev-requeriments' es un <div>. Assignar-li '.value' NO fa res: la vista
#     previa del correu es quedava a la pantalla en fer "Fet" i l'informe nou
#     ensenyava el correu de l'anterior. Els tres altres punts del fitxer que hi
#     escriuen ja feien servir .innerHTML; nomes el reinici anava amb .value.
Assert (-not ($appJs.Contains('$("prev-requeriments").value'))) 'app.js: prev-requeriments es un div, mai .value'
Assert ($appJs.Contains('$("prev-requeriments").innerHTML = "";')) 'app.js: el reinici buida la vista previa'

# 4b. El comptador de passos ha de comptar els ABASTABLES. PASSOS en te 5 i el
#     de cataleg se salta quan nomes n'hi ha un: deia "Pas 1 / 5" i no s'hi
#     arribava mai. Amb passosAbastables() la navegacio i el comptador surten
#     del MATEIX lloc i no es poden desincronitzar.
Assert ($appJs.Contains('function passosAbastables')) 'app.js: el comptador compta els passos abastables'
Assert ($appJs.Contains('"Pas " + (pos + 1) + " / " + abast.length')) 'app.js: el comptador fa servir els abastables'
Assert (-not ($appJs.Contains('" / " + PASSOS.length'))) 'app.js: el comptador no fa servir el total cru'

# 4c. Un ternari '? true : true' es una condicio que no decideix res -els dos
#     costats son iguals-. Aqui amagava que #navegacio no s'amaga mai: es mostra
#     un sol cop en carregar. La condicio feia pensar el contrari.
Assert (-not ($appJs.Contains('? true : true'))) 'app.js: cap ternari amb les dues branques iguals'

# 4d. Codi mort: collectFields/passId no els crida ningu, i estat.fieldOrder /
#     fieldDefs no els llegeix ningu (els camps s'omplen INLINE, ja no tenen un
#     pas propi). Els deixava aqui la versio del pas 5 que es va treure.
foreach ($mort in @('function collectFields', 'function passId', 'fieldOrder', 'fieldDefs')) {
    Assert (-not ($appJs.Contains($mort))) "app.js: no torna el codi mort del pas 5 ($mort)"
}

# 4e. ORIGEN, DATES i CLASSIFICACIO son <<placeholders>> de la capcalera que
#     l'usuari NO omple: el mobil els pintava com a quadres de text buits. Es
#     comprova contra el capcalera.json de debo, que es d'on surten.
$capMobil = Read-JsonFile (Join-Path $rootRepo (Join-Path 'docs' (Join-Path 'dades' 'capcalera.json')))
Assert ($null -ne $capMobil) 'app.js: hi ha docs/dades/capcalera.json'
foreach ($ph in @('ORIGEN', 'DATES', 'CLASSIFICACIO')) {
    Assert (@($capMobil.Placeholders) -contains $ph) "capcalera.json: encara publica $ph"
    Assert ($appJs -match ("(?m)^\s*ORIGEN: 1, DATES: 1, CLASSIFICACIO: 1")) "app.js: HEADER_SKIP_GENERIC salta $ph"
}

Write-Host "`n--- docs/dades: el cataleg derivat no pot quedar-se enrere (guard) ---"
# PER QUE. docs/dades/cataleg-REQ1.json es una copia DERIVADA d'ESTRUCTURALS que
# viu al git, i es va quedar 6 commits enrere: al mobil hi faltaven 53
# requeriments (la seccio de VMP sencera) i encara citava normativa d'incendis
# DEROGADA, mentre el PC generava be perque llegeix ESTRUCTURALS directament.
# La regeneracio no es disparava mai quan el cataleg arribava per 'git pull'.
# Aquesta prova compara els titols dels dos costats: si algu torna a publicar
# una copia endarrerida, surt aqui i no a mans d'un titular.
function _TitolsOrigen($node, $acc) {
    foreach ($n in @($node)) {
        if ($null -eq $n) { continue }
        $t = [string]$n.titol
        if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$acc.Add($t.Trim()) }
        if ($n.fills) { _TitolsOrigen $n.fills $acc }
    }
}
function _TitolsDerivat($cat, $acc) {
    foreach ($sec in @($cat.Sections)) {
        if ($null -eq $sec) { continue }
        $t = [string]$sec.Title
        if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$acc.Add($t.Trim()) }
        foreach ($it in @($sec.Items)) {
            if ($null -eq $it) { continue }
            $s = [string]$it.Short
            if (-not [string]::IsNullOrWhiteSpace($s)) { [void]$acc.Add($s.Trim()) }
            foreach ($ch in @($it.Children)) {
                if ($null -eq $ch) { continue }
                $cs = [string]$ch.Short
                if (-not [string]::IsNullOrWhiteSpace($cs)) { [void]$acc.Add($cs.Trim()) }
            }
        }
    }
}
$dirDades = Join-Path $rootRepo (Join-Path 'docs' 'dades')
$derivats = @(Get-ChildItem -Path $dirDades -Filter 'cataleg-*.json' -File -ErrorAction SilentlyContinue)
Assert ($derivats.Count -ge 1) "hi ha catalegs derivats a docs/dades ($($derivats.Count))"
foreach ($d in $derivats) {
    $base = $d.BaseName -replace '^cataleg-', ''
    $orig = Join-Path $rootRepo (Join-Path 'ESTRUCTURALS' ($base + '.json'))
    Assert (Test-Path -LiteralPath $orig) "$base : el derivat te el seu origen a ESTRUCTURALS"
    if (-not (Test-Path -LiteralPath $orig)) { continue }
    $oJson = Get-Content -LiteralPath $orig -Raw -Encoding UTF8 | ConvertFrom-Json
    $dJson = Get-Content -LiteralPath $d.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $tO = New-Object System.Collections.ArrayList
    $tD = New-Object System.Collections.ArrayList
    _TitolsOrigen $oJson.nodes $tO
    _TitolsDerivat $dJson $tD
    $falten = @(@($tO) | Where-Object { $tD -notcontains $_ })
    $sobren = @(@($tD) | Where-Object { $tO -notcontains $_ })
    AssertEq $falten.Count 0 ("$base : cap titol de l'origen falta al derivat" + $(if ($falten.Count) { " (n'hi falten $($falten.Count), p.ex. '" + $falten[0] + "')" } else { '' }))
    AssertEq $sobren.Count 0 ("$base : el derivat no te titols que ja no son a l'origen" + $(if ($sobren.Count) { " (n'hi sobren $($sobren.Count), p.ex. '" + $sobren[0] + "')" } else { '' }))
}



Write-Host "`n--- Json.ps1: Read-JsonFile / Write-JsonFile ---"
# PER QUE. Abans cada modul llegia i escrivia JSON pel seu compte: 17 llocs amb
# BOM i 7 sense, cap escriptura atomica, i el mateix esquelet "carrega o valor
# per defecte" copiat quatre cops. El perill de debo no era l'estetica: un desat
# interromput deixa el fitxer TRUNCAT, i com que tots els lectors tracten un JSON
# corrupte igual que un que no hi es, la base de llicencies es podia perdre
# SENSE CAP AVIS.
$jsDir = Join-Path ([System.IO.Path]::GetTempPath()) ('json-test-' + [Guid]::NewGuid().ToString('N'))
try {
    # --- Read-JsonFile: els quatre casos en que no hi ha res utilitzable ---
    Assert ($null -eq (Read-JsonFile (Join-Path $jsDir 'no-hi-es.json'))) 'Read-JsonFile: fitxer inexistent -> $null'
    Assert ($null -eq (Read-JsonFile '')) 'Read-JsonFile: ruta buida -> $null'
    [void](New-Item -ItemType Directory -Path $jsDir -Force)
    $jsBuit = Join-Path $jsDir 'buit.json'
    [System.IO.File]::WriteAllText($jsBuit, '')
    Assert ($null -eq (Read-JsonFile $jsBuit)) 'Read-JsonFile: fitxer buit -> $null'
    [System.IO.File]::WriteAllText($jsBuit, "   `n  ")
    Assert ($null -eq (Read-JsonFile $jsBuit)) 'Read-JsonFile: nomes espais -> $null'
    $jsMal = Join-Path $jsDir 'corrupte.json'
    [System.IO.File]::WriteAllText($jsMal, '{"a": 1, tallat')
    Assert ($null -eq (Read-JsonFile $jsMal)) 'Read-JsonFile: JSON truncat -> $null (no peta)'

    # --- Write-JsonFile: anada i tornada, sense BOM, i crea la carpeta ---
    $jsSub = Join-Path (Join-Path $jsDir 'a') 'b'
    $jsOut = Join-Path $jsSub 'dades.json'
    $jsObj = [pscustomobject]@{
        Version = 1
        Nom     = "Accents: cal que hi siguin"
        Llista  = @('u', 'dos')
        Nested  = [pscustomobject]@{ Fons = [pscustomobject]@{ Valor = 42 } }
    }
    Write-JsonFile $jsOut $jsObj 8
    Assert (Test-Path -LiteralPath $jsOut) 'Write-JsonFile: crea la carpeta de desti si cal'

    $jsBytes = [System.IO.File]::ReadAllBytes($jsOut)
    $jsTeBom = ($jsBytes.Length -ge 3 -and $jsBytes[0] -eq 0xEF -and $jsBytes[1] -eq 0xBB -and $jsBytes[2] -eq 0xBF)
    AssertEq $jsTeBom $false 'Write-JsonFile: escriu UTF-8 SENSE BOM'

    $jsLlegit = Read-JsonFile $jsOut
    Assert ($null -ne $jsLlegit) 'Write-JsonFile + Read-JsonFile: anada i tornada'
    AssertEq ([int]$jsLlegit.Version) 1 'anada i tornada: valor simple'
    AssertEq ([string]$jsLlegit.Nom) 'Accents: cal que hi siguin' 'anada i tornada: accents intactes'
    AssertEq (@($jsLlegit.Llista).Count) 2 'anada i tornada: la llista es conserva'
    AssertEq ([int]$jsLlegit.Nested.Fons.Valor) 42 'anada i tornada: la -Depth arriba al fons'

    # --- No queda cap temporal, i sobreescriure funciona ---
    AssertEq (@(Get-ChildItem -Path $jsSub -Filter '*.tmp' -File).Count) 0 'Write-JsonFile: no deixa cap .tmp enrere'
    Write-JsonFile $jsOut ([pscustomobject]@{ Version = 2 }) 4
    AssertEq ([int](Read-JsonFile $jsOut).Version) 2 'Write-JsonFile: sobreescriu el fitxer existent'
    AssertEq (@(Get-ChildItem -Path $jsSub -Filter '*.tmp' -File).Count) 0 'Write-JsonFile: tampoc en deixa en sobreescriure'

    # --- La -Depth es del crider: massa poca TRUNCA (per aixo no te valor per defecte) ---
    $jsFons = Join-Path $jsSub 'fons.json'
    Write-JsonFile $jsFons $jsObj 2
    Assert (([string](Read-JsonFile $jsFons).Nested.Fons) -notmatch '^\s*$') 'Write-JsonFile: amb -Depth curta el fons es trunca (contracte del crider)'
} finally {
    if (Test-Path -LiteralPath $jsDir) { Remove-Item -LiteralPath $jsDir -Recurse -Force -ErrorAction SilentlyContinue }
}



Write-Host "`n--- ConfigJs.ps1: PowerShell llegint docs/config.js ---"
# PER QUE. _CorreuConfig portava QUATRE expressions regulars escampades, una per
# clau, i cadascuna exigia cometes DOBLES i el nom literal. Amb cometes simples,
# amb la clau comentada o amb una plantilla, la clau es quedava BUIDA en silenci
# i l'enviament de correu deixava de funcionar sense dir per que. Zero proves.
#
# Aquesta prova fa correr el lector contra el docs/config.js DE DEBO: si algu el
# reformata, ho diu la suite abans que es trenqui res.
$cjRoot = Split-Path -Parent (Split-Path -Parent $TestsDir)
$cjReal = Read-ConfigJs
foreach ($cjK in @('EMAILJS_PUBLIC_KEY','EMAILJS_SERVICE_ID','EMAILJS_TEMPLATE_ID','EMAIL_FROM_NAME',
                   'DRIVE_ENTRADA_FOLDER_ID','DRIVE_PROCESSATS_FOLDER_ID','DRIVE_DADES_FOLDER_ID',
                   'GOOGLE_CLIENT_ID')) {
    Assert (-not [string]::IsNullOrWhiteSpace((Get-ConfigJsValue $cjReal $cjK))) ("config.js real: hi ha " + $cjK)
}

# Els tres IDs de Drive han de venir d'AQUI i de cap altre lloc: abans estaven
# duplicats a suport/config.ps1 i el manual demanava escriure'ls dues vegades.
AssertEq (Get-ConfigJsValue $cjReal 'DRIVE_ENTRADA_FOLDER_ID') ([string]$DriveEntradaId) 'DriveEntradaId surt de config.js'
AssertEq (Get-ConfigJsValue $cjReal 'DRIVE_PROCESSATS_FOLDER_ID') ([string]$DriveProcessatsId) 'DriveProcessatsId surt de config.js'
AssertEq (Get-ConfigJsValue $cjReal 'DRIVE_DADES_FOLDER_ID') ([string]$DriveDadesId) 'DriveDadesId surt de config.js'

# I que cap dels tres no torni a apareixer escrit a suport/config.ps1.
$cjCfgPs1 = Get-Content -LiteralPath (Join-Path (Join-Path $cjRoot 'suport') 'config.ps1') -Raw -Encoding UTF8
$cjDup = @()
foreach ($cjV in @([string]$DriveEntradaId, [string]$DriveProcessatsId, [string]$DriveDadesId)) {
    if (-not [string]::IsNullOrWhiteSpace($cjV) -and $cjCfgPs1.Contains($cjV)) { $cjDup += $cjV }
}
AssertEq $cjDup.Count 0 ('cap ID de Drive duplicat a config.ps1' + $(if ($cjDup.Count) { ' -> ' + ($cjDup -join ', ') } else { '' }))

# --- Formats que el lector ha d'aguantar ---
$cjText = @'
// Un comentari qualsevol.
window.CONFIG = {
  AMB_DOBLES: "valor-doble",
  AMB_SIMPLES: 'valor-simple',
  // COMENTADA: "no-hauria-de-sortir",
  BUIDA: "",
  AMB_ESPAIS   :   "amb-espais"
};
'@
$cjP = Read-ConfigJs $cjText
AssertEq (Get-ConfigJsValue $cjP 'AMB_DOBLES') 'valor-doble' 'Read-ConfigJs: cometes dobles'
AssertEq (Get-ConfigJsValue $cjP 'AMB_SIMPLES') 'valor-simple' 'Read-ConfigJs: cometes SIMPLES (abans es perdia)'
AssertEq (Get-ConfigJsValue $cjP 'AMB_ESPAIS') 'amb-espais' 'Read-ConfigJs: espais al voltant dels dos punts'
AssertEq (Get-ConfigJsValue $cjP 'COMENTADA' '(cap)') '(cap)' 'Read-ConfigJs: una linia comentada NO compta'
AssertEq (Get-ConfigJsValue $cjP 'BUIDA' '(defecte)') '(defecte)' 'Get-ConfigJsValue: clau buida -> valor per defecte'
AssertEq (Get-ConfigJsValue $cjP 'NO_HI_ES' '(defecte)') '(defecte)' 'Get-ConfigJsValue: clau inexistent -> valor per defecte'
AssertEq ((Read-ConfigJs "// nomes un comentari`nres a veure").Count) 0 'Read-ConfigJs: text sense claus -> res, i no peta'


Write-Host "`n--- Excel.ps1: la seqüencia compartida de la fulla Estesa ---"
# PER QUE. Set funcions obrien l'Excel, buscaven la fulla "Estes", en llegien la
# matriu i tancaven, cada una amb la seva copia de vint linies. Ara la seqüencia
# es UNA, i aquestes proves la fan correr amb un DOBLE de COM -pscustomobject +
# Add-Member, el mateix patro que la prova de Write-InformeDocx-, que es l'unica
# manera de provar-la en un Linux sense Excel.
#
# TOT EL BLOC VA DINS D'UN try. Sense aixo, una excepcio que s'escapi -per
# exemple si el tancament de Read-FullaEstesa tornes a petar dins del finally-
# MATA el bloc sencer i les asercions de despres no s'executen mai... i la suite
# segueix dient "0 FAIL", perque nomes compta el que s'ha arribat a comprovar.
# Es la mateixa lliçó del throw que va matar mitja suite (vegeu CLAUDE.md).
{
  try {
    # L'Excel torna una matriu 2D amb els indexs comencant per 1. PowerShell no
    # tria l'overload de tres arguments d'Array::CreateInstance (aplana els dos
    # int[] en un de sol), aixi que la demanem per reflexio. I la coma del
    # 'return' es obligatoria: una matriu 2D retornada d'una funcio l'aplana el
    # pipeline igual que una llista.
    $ci = [Array].GetMethod('CreateInstance', [Type[]]@([Type], [int[]], [int[]]))
    function _XlMatriu([int]$files, [int]$cols) {
        return ,($ci.Invoke($null, @([object], [int[]]@($files, $cols), [int[]]@(1, 1))))
    }
    function _XlDoble([int]$nFiles, [int]$nCols, [string]$nomFulla = 'Estes') {
        $m = $null
        if ($nFiles -ge 0) {
            $m = _XlMatriu ($nFiles + 1) $nCols
            for ($c = 1; $c -le $nCols; $c++) { $m[1, $c] = "H$c" }
            # PARENTESIS a ($i+1): la coma lliga mes fort que el +, i "$m[$i+1, $c]"
            # es llegiria com "$m[$i + (1,$c)]". Trampa ja documentada al CLAUDE.md.
            for ($i = 1; $i -le $nFiles; $i++) {
                for ($c = 1; $c -le $nCols; $c++) { $m[($i + 1), $c] = "r${i}c$c" }
            }
        }
        $sh = [pscustomobject]@{ Name = $nomFulla; UsedRange = [pscustomobject]@{ Value2 = $m } }
        $wb = [pscustomobject]@{ Sheets = @($sh) }
        $wb | Add-Member ScriptMethod Close { param($q) $script:_xlTancat++ } -Force
        $wbs = [pscustomobject]@{}
        $wbs | Add-Member ScriptMethod Open { param($p, $a, $b) $wb }.GetNewClosure() -Force
        $ex = [pscustomobject]@{ Visible = $true; DisplayAlerts = $true; Workbooks = $wbs }
        $ex | Add-Member ScriptMethod Quit { $script:_xlQuit++ } -Force
        $script:_xlFals = $ex; $script:_xlTancat = 0; $script:_xlQuit = 0
    }
    # Substitueix NOMES la creacio de l'Excel; tota la resta de Read-FullaEstesa
    # es el codi de produccio tal qual.
    $srcXl = Get-Content -LiteralPath (Join-Path $rootRepo (Join-Path 'suport' 'Excel.ps1')) -Raw
    $srcXl = $srcXl.Replace(
        'try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }',
        '$excel = $script:_xlFals')
    Invoke-Expression $srcXl
    $fx = [pscustomobject]@{ FullName = 'prova.xlsx' }

    # 1. LA FORMA DE CONSUMIR-LA. El cos torna ,@(...) i el crider ASSIGNA i
    #    torna ,@($out). El cas d'UN SOL registre es l'unic que distingeix la
    #    forma bona de la dolenta: amb 'return (Read-FullaEstesa ...)' directe
    #    hi arribava l'objecte PELAT en lloc d'un array.
    function _XlLlegeix($f) {
        $out = Read-FullaEstesa $f {
            param($x)
            $recs = @()
            for ($r = 2; $r -le $x.Rows; $r++) { $recs += [pscustomobject]@{ Id = $x.Data[$r, 1] } }
            return ,@($recs)
        }
        return ,@($out)
    }
    # S'ASSIGNA primer, que es com ho fan els tres criders de debo
    # (Precintades:249, Ruta:977, Coordenades:1344). Amb '@(_XlLlegeix $fx).Count'
    # a pel el compte dona SEMPRE 1: el pipeline desenrotlla el resultat i l'@()
    # el torna a embolcallar. La regla val tambe aqui.
    foreach ($n in 0, 1, 2, 5) {
        _XlDoble $n 2
        $recs = _XlLlegeix $fx
        AssertEq (@($recs).Count) $n "Read-FullaEstesa: $n registres arriben sencers al crider"
    }
    _XlDoble 1 2
    $un = _XlLlegeix $fx
    AssertEq ([string](@($un)[0].Id)) 'r1c1' 'Read-FullaEstesa: amb UN registre no arriba l''objecte pelat'

    # 1b. EL LECTOR DE CEL.LA ve amb el context. Estava copiat als CINC cossos
    #     -Activitats, ControlsPeriodics, Ruta, Precintades i Coordenades-,
    #     identic fins a l'espaiat. Es comprova que retalla, que tolera els
    #     nulls i el fora de rang, i -aixo es el que calia mesurar- que el
    #     $data i el $cols que el COS es declara pel seu compte NO el despisten:
    #     el bloc porta .GetNewClosure() i es queda els de Read-FullaEstesa.
    _XlDoble 2 2
    $celR = Read-FullaEstesa $fx {
        param($x)
        $get = $x.Cel
        $data = 'JO SOC UN ALTRE'; $cols = 1
        return "[$(& $get 2 1)][$(& $get 2 9)][$(& $get 3 2)]"
    }
    AssertEq $celR '[r1c1][][r2c2]' '$x.Cel: llegeix, i el $data del cos no el despista'

    # 2. Les capcaleres de la fila 1, consumides dins del cos.
    foreach ($nc in 1, 2, 5) {
        _XlDoble 2 $nc
        $h = Read-FullaEstesa $fx { param($x) return "$(@($x.Headers).Count)|$($x.Headers[0])" }
        AssertEq $h "$nc|H1" "_HeadersDeFila1: $nc columnes (el cas d'1 es el que enxampa la cadena pelada)"
    }

    # 3. Un Excel BUIT no es un error: el cos es crida igual i decideix ell.
    _XlDoble -1 0
    $buit = Read-FullaEstesa $fx { param($x) "$($x.Rows)|$($x.Cols)|$(@($x.Headers).Count)" }
    AssertEq $buit '0|0|0' 'Read-FullaEstesa: Excel buit -> el cos rep 0 files i decideix'

    # 4. EL DEFECTE DE L'ORFE. Les tres copies de rutes/ feien el Close i el Quit
    #    SENSE try: si el Close petava, el Quit no s'executava i quedava un
    #    EXCEL.EXE corrent amb el fitxer agafat. Aqui es comprova que es tanca
    #    tot i que el cos LLANCI.
    _XlDoble 3 2
    $peto = ''
    try { Read-FullaEstesa $fx { param($x) throw 'el cos peta' } } catch { $peto = $_.Exception.Message }
    AssertEq $peto 'el cos peta' 'Read-FullaEstesa: l''error del cos es propaga'
    AssertEq $script:_xlTancat 1 'Read-FullaEstesa: tanca el llibre encara que el cos peti'
    AssertEq $script:_xlQuit   1 'Read-FullaEstesa: ...i surt de l''Excel (res d''EXCEL.EXE orfe)'

    # 5. Si la fulla no hi es, el missatge diu com es diuen les que hi ha.
    _XlDoble 2 2 'Resum'
    $errF = ''
    try { Read-FullaEstesa $fx { param($x) 1 } } catch { $errF = $_.Exception.Message }
    Assert ($errF.Contains('Resum')) 'Read-FullaEstesa: si no troba la fulla, diu quines hi ha'

    # 6. _NormalitzaText: un sol normalitzador per als DOS processos.
    AssertEq (_NormalitzaText ('  Est' + [char]0x00E8 + 's  ')) 'estes' '_NormalitzaText: sense accents, sense espais, en minuscules'
    AssertEq (_NormalitzaText $null) '' '_NormalitzaText: $null -> cadena buida'
  } catch {
    Assert $false ("Excel.ps1: una excepcio s'ha escapat del bloc de proves -> " + $_.Exception.Message)
  }
}.Invoke() | Out-Null

Write-Host "`n--- Nomes Excel.ps1 obre l'Excel per COM (guard) ---"
# PER QUE. Set funcions creaven la seva propia instancia d'Excel i tres d'elles
# ni comprovaven el $null ni protegien el tancament. Amb un sol lloc, arreglar-ho
# alla val per a totes; aquest guard es el que impedeix que en torni a apareixer
# una de nova sense adonar-se'n.
#
# SeguimentGia.ps1 es l'UNICA excepcio, i esta raonada al seu codi: no nomes
# llegeix, tambe copia la fulla dins d'un llibre NOU amb la mateixa instancia i
# l'exporta a PDF, o sigui que necessita l'Excel obert mes enlla del cos.
# Es mira NOMES el codi de produccio: les proves parlen del literal per forca
# (l'han de substituir pel doble), i comptar-les seria un fals positiu.
$excelPermes = @('Excel.ps1', 'SeguimentGia.ps1')
$obrenExcel = @()
foreach ($f in $ps1Tots) {
    if ($f.FullName -like ('*' + [System.IO.Path]::DirectorySeparatorChar + 'tests' + [System.IO.Path]::DirectorySeparatorChar + '*')) { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    if ($t.Contains('New-Object -ComObject Excel.Application') -and ($excelPermes -notcontains $f.Name)) {
        $obrenExcel += $f.Name
    }
}
AssertEq $obrenExcel.Count 0 ('nomes Excel.ps1 (i SeguimentGia) obren l''Excel' + $(if ($obrenExcel.Count) { ' -> ' + ($obrenExcel -join ', ') } else { '' }))

Write-Host "`n--- Nomes Motor.ps1 obre el Word per COM (guard) ---"
# PER QUE. New-WordApp (Motor.ps1) ja feia tres coses que calen sempre: la guarda
# del $null amb un missatge que diu que passa, Visible/DisplayAlerts, i
# AutomationSecurity = 1, que es el que impedeix que el Word obri en VISTA
# PROTEGIDA els fitxers d'una unitat de xarxa. QUATRE llocs se la saltaven amb un
# New-Object a pel:
#
#   EnviarCorreu.ps1       sense CAP guarda del $null, i sense AutomationSecurity
#   VistaWord.ps1          guarda propia, sense AutomationSecurity
#   Informes.ps1           guarda propia, sense AutomationSecurity  (.doc antics)
#   ControlsPeriodics.ps1  copia sencera, amb un missatge pitjor
#
# I els informes viuen a "I:\...", que es exactament el cas per al qual
# New-WordApp porta aquella linia. Aquest guard es el que impedeix que en torni a
# apareixer un de nou sense adonar-se'n.
$wordPermes = @('Motor.ps1')
$obrenWord = @()
foreach ($f in $ps1Tots) {
    if ($f.FullName -like ('*' + [System.IO.Path]::DirectorySeparatorChar + 'tests' + [System.IO.Path]::DirectorySeparatorChar + '*')) { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    if ($t.Contains('New-Object -ComObject Word.Application') -and ($wordPermes -notcontains $f.Name)) {
        $obrenWord += $f.Name
    }
}
AssertEq $obrenWord.Count 0 ('nomes Motor.ps1 obre el Word' + $(if ($obrenWord.Count) { ' -> ' + ($obrenWord -join ', ') } else { '' }))

# I que New-WordApp segueixi fent les tres coses: si algu li tragues
# l'AutomationSecurity, els quatre criders ho perdrien tots alhora i en silenci.
$srcMotor = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' 'Motor.ps1')))
Assert ($srcMotor.Contains('$w.AutomationSecurity = 1')) 'New-WordApp posa AutomationSecurity (res de Vista protegida)'
Assert ($srcMotor.Contains('param([switch]$Opcional)')) 'New-WordApp te -Opcional (qui pot continuar sense Word)'


Write-Host "`n--- El correu: una manera de fer HTML i una d'omplir variables ---"
# PER QUE. _RecCosHtml (Recordatoris) i _ControlsCpEmailHtml (Controls
# periodics) eren IDENTIQUES linia a linia, estil inline inclos; i hi havia TRES
# bucles iguals per substituir les variables {X}. Ara l'HTML el fa _CosAHtml i el
# bucle _OmpleVariables; el MAPA de variables el segueix posant cada eina, que es
# l'unica cosa que difereix de debo.

# ELS DOS URLs A LA MATEIXA LINIA. Aquest era un DEFECTE REAL, no una millora:
# la cursiva es "//...//" i un "https://" en porta un "//" a dins, o sigui que
# amb dues adreces seguides l'expressio es menjava tot el tros d'una a l'altra i
# les destrossava totes dues. _TextToHtml ja la feien servir els recordatoris,
# amb un text que l'usuari pot editar. Ara els URLs s'aparten abans de mirar la
# negreta i la cursiva, i es tornen a posar al final.
$dosUrls = _TextToHtml 'Mira https://a.cat i tambe https://b.cat aqui'
Assert ($dosUrls.Contains('<a href="https://a.cat">')) 'dos URLs a una linia: el primer segueix sent un enllac'
Assert ($dosUrls.Contains('<a href="https://b.cat">')) 'dos URLs a una linia: i el segon tambe'
Assert (-not $dosUrls.Contains('<i>')) 'dos URLs a una linia: el "//" d''un http ja no obre cap cursiva'
AssertEq (_TextToHtml 'amb //cursiva// de debo') 'amb <i>cursiva</i> de debo' 'la cursiva de debo segueix funcionant'
AssertEq (_TextToHtml 'i **negreta** amb https://x.cat/a?b=1&c=2') 'i <b>negreta</b> amb <a href="https://x.cat/a?b=1&amp;c=2">https://x.cat/a?b=1&amp;c=2</a>' 'negreta i URL amb & escapat, tot alhora'
AssertEq (_TextToHtml '') '' '_TextToHtml: text buit -> buit'

# _OmpleVariables: nomes el bucle; el mapa el posa el crider.
AssertEq (_OmpleVariables 'GIA {A} i {B}' ([ordered]@{ '{A}' = '1'; '{B}' = 'dos' })) 'GIA 1 i dos' '_OmpleVariables: substitueix les claus del mapa'
AssertEq (_OmpleVariables 'sense res' ([ordered]@{})) 'sense res' '_OmpleVariables: mapa buit -> text igual'
AssertEq (_OmpleVariables 'queda {NOCONEC}' ([ordered]@{ '{A}' = '1' })) 'queda {NOCONEC}' '_OmpleVariables: una clau que no hi es es queda tal qual'
AssertEq (_OmpleVariables 'text' $null) 'text' '_OmpleVariables: sense mapa no peta'

# Guard: cap tercera copia de l'estil inline del cos del correu.
$copiesEstil = 0
foreach ($f in $ps1Tots) {
    if ($f.FullName -like ('*' + [System.IO.Path]::DirectorySeparatorChar + 'tests' + [System.IO.Path]::DirectorySeparatorChar + '*')) { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $copiesEstil += ([regex]::Matches($t, 'font-family:Segoe UI,Arial,sans-serif;font-size:11pt')).Count
}
AssertEq $copiesEstil 1 'l''estil del cos del correu viu en UN sol lloc (_CosAHtml)'


Write-Host "`n--- L'editor d'assumpte + cos: una sola pantalla ---"
# PER QUE. Hi havia TRES pantalles gairebe iguals, amb les MATEIXES coordenades
# (textos del mobil, avis de control periodic, text d'una campanya de
# recordatoris): ~100 linies cadascuna per pintar dues etiquetes, dos quadres,
# una linia d'ajuda i tres botons. Ara la pantalla es Show-EditorAssumpteCos
# (UiComuns.ps1) i cada eina hi posa NOMES el que difereix de debo: que vol dir
# DESAR i que vol dir RESTAURAR, tots dos com a scriptblock.
#
# La finestra es WinForms i no es pot obrir aqui; el que si que es pot provar es
# el mecanisme del qual depen tot: que un scriptblock passat com a PARAMETRE
# segueixi veient els LOCALS de la funcio que el va escriure quan qui el crida es
# un handler d'una altra funcio. Si no fos aixi, el bloc -Desa dels textos del
# mobil rebria un $textos buit i DESAR S'ENDURIA LA LLISTA DE CCO en silenci.
# Esta MESURAT que funciona sense .GetNewClosure() -per aixo no se n'hi ha posat
# cap-, i aquesta prova ho deixa clavat: validada substituint el bloc per un de
# fet amb [scriptblock]::Create, que no te scope de definicio, i comprovant que
# el valor arriba BUIT.
{
  try {
    function _EacMotor([scriptblock]$Desa, [scriptblock]$Restaurar) {
        # El handler SI que porta closure, com els de la pantalla de debo.
        $h = { "$(& $Desa 'v') | $(& $Restaurar)" }.GetNewClosure()
        return (& $h)
    }
    function _EacCrider {
        $textos = @{ bcc = 'CCO' }
        $clau   = 'requeriments'
        return (_EacMotor -Desa { param($v) "desa:$($textos.bcc)/$v" } -Restaurar { "rest:$clau" })
    }
    AssertEq (_EacCrider) 'desa:CCO/v | rest:requeriments' 'un scriptblock parametre veu els locals de qui el va escriure'
  } catch {
    Assert $false ("editor assumpte+cos: excepcio -> " + $_.Exception.Message)
  }
}.Invoke() | Out-Null

# Les tres eines han de passar per la pantalla compartida, i cap d'elles pot
# tornar a muntar-se la seva.
$eacFitxers = @{
    'EmailTextos.ps1'     = 'Invoke-EmailTextos'
    'ControlsCpEmail.ps1' = 'Invoke-ControlsCpEmailTextos'
    'Recordatoris.ps1'    = 'Invoke-RecordatorisTextos'
}
foreach ($nom in $eacFitxers.Keys) {
    $t = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' $nom)))
    Assert ($t.Contains('Show-EditorAssumpteCos')) "$nom : $($eacFitxers[$nom]) passa per la pantalla compartida"
    Assert (-not ($t -match "(?m)^\s*\`$lblH\.Font\s*=")) "$nom : ja no es torna a muntar la pantalla pel seu compte"
}
# ...i la pantalla viu en UN sol lloc.
$copiesEac = 0
foreach ($f in $ps1Tots) {
    if ($f.FullName -like ('*' + [System.IO.Path]::DirectorySeparatorChar + 'tests' + [System.IO.Path]::DirectorySeparatorChar + '*')) { continue }
    $copiesEac += ([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), 'function Show-EditorAssumpteCos')).Count
}
AssertEq $copiesEac 1 'Show-EditorAssumpteCos esta definida UNA sola vegada'

# El TextBox multilinia de WinForms nomes ensenya CRLF i els cossos es desen amb
# LF. Ara la conversio es fa en UN sol lloc (abans cada pantalla ho havia de
# recordar, i la de recordatoris NO ho feia).
$srcUi = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' 'UiComuns.ps1')))
Assert ($srcUi.Contains('$Cos -replace "`r?`n", "`r`n"')) 'l''editor normalitza a CRLF en pintar'
Assert ($srcUi.Contains('$tbC.Text -replace "`r`n", "`n"')) 'l''editor torna a LF en desar'


Write-Host "`n--- Cap .json s'escriu fora de Json.ps1 (guard) ---"
# PER QUE. Quan tot el JSON va passar per Json.ps1 en va quedar UN de fora:
# Activitats.ps1 escrivia activitats.json amb 'Set-Content -Encoding UTF8'. Hi
# perdia dues coses i les dues es noten al disc:
#
#   1. BOM. Set-Content -Encoding UTF8 al PowerShell 5.1 n'hi posa, i el MATEIX
#      contingut que puja a Drive (Save-DriveJson) no en porta: dues copies del
#      mateix fitxer amb codificacio diferent.
#   2. No era ATOMIC. I activitats.json es tota la base d'activitats del mobil:
#      una escriptura interrompuda el deixa truncat, i tots els lectors tracten
#      un JSON corrupte igual que un que no hi es. El mobil es quedaria sense
#      base sense dir res.
#
# Es mira el codi de produccio: una linia amb Set-Content o WriteAllText que
# tingui '.json' a la vista (la mateixa linia o les tres anteriors, que es on hi
# sol anar el Join-Path del nom).
$jsonAPel = @()
foreach ($f in $ps1Tots) {
    if ($f.FullName -like ('*' + [System.IO.Path]::DirectorySeparatorChar + 'tests' + [System.IO.Path]::DirectorySeparatorChar + '*')) { continue }
    if ($f.Name -eq 'Json.ps1') { continue }
    $ln = [System.IO.File]::ReadAllLines($f.FullName)
    # Els comentaris de BLOC compten: la capcalera de Settings.ps1 explicava
    # com es desava ABANS i el guard la va enxampar. (El comentari estava
    # desfasat de debo -deia Set-Content i el codi ja feia Write-JsonFile-, o
    # sigui que el guard va servir d'alguna cosa; pero no ha de mirar text.)
    $dinsBloc = $false
    for ($i = 0; $i -lt $ln.Count; $i++) {
        $l = $ln[$i]
        if ($l -match '<#') { $dinsBloc = $true }
        if ($dinsBloc) { if ($l -match '#>') { $dinsBloc = $false }; continue }
        if ($l -match '^\s*#') { continue }
        if ($l -notmatch 'Set-Content|WriteAllText') { continue }
        $desde = [Math]::Max(0, $i - 3)
        $ctx = ($ln[$desde..$i] -join "`n")
        if ($ctx -match '\.json') { $jsonAPel += ($f.Name + ':' + ($i + 1)) }
    }
}
AssertEq $jsonAPel.Count 0 ('cap .json escrit fora de Json.ps1' + $(if ($jsonAPel.Count) { ' -> ' + ($jsonAPel -join ', ') } else { '' }))

# I que activitats.json hi passi de debo.
$srcAct = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' 'Activitats.ps1')))
Assert ($srcAct.Contains('Write-JsonText $outFile $json')) 'activitats.json s''escriu amb Write-JsonText (UTF-8 sense BOM i atomic)'


Write-Host "`n--- El nom del fitxer de sortida: un sol sanejat ---"
# PER QUE. Hi havia TRES funcions que feien el nom i NO deien el mateix:
# _GetOutputFileName (Document.ps1) posava '_' i nomes sanejava el GIA;
# _LlicNomFitxer i _MnsNomFitxer posaven '-' i sanejaven el nom sencer. El
# MATEIX caracter dolent donava un nom diferent segons quin informe fessis.
# Ara el sanejat es un (_SanejaNomFitxer) i les dues bessones -que nomes es
# diferenciaven en com trien el $curt- comparteixen _NomInformeFitxer.
AssertEq (_SanejaNomFitxer 'a/b:c*d?e"f<g>h|i\i') 'a-b-c-d-e-f-g-h-i-i' '_SanejaNomFitxer: tots els caracters que Windows no admet'
AssertEq (_SanejaNomFitxer 'sense res a treure') 'sense res a treure' '_SanejaNomFitxer: no toca el que ja va be'
AssertEq (_NomInformeFitxer ([datetime]'2026-09-04') 'LlicReq' '1433') '2026-09-04_LlicReq_GIA 1433.docx' '_NomInformeFitxer: data_curt_GIA'
AssertEq (_NomInformeFitxer ([datetime]'2026-09-04') 'LlicReq' '')     '2026-09-04_LlicReq.docx'          '_NomInformeFitxer: sense GIA, sense el tros del GIA'

# LES TRES han de sanejar IGUAL: es el que no passava.
$sanejats = @(
    (_GetOutputFileName 'REQ1' 'A/B'),
    (_LlicNomFitxer ([datetime]'2026-09-04') 'requeriment' 'A/B'),
    (_MnsNomFitxer  ([datetime]'2026-09-04') 'mns' 'A/B')
)
foreach ($n in $sanejats) { Assert ($n.Contains('A-B')) "les tres families sanegen igual ('-'): $n" }
Assert (-not (($sanejats -join '') -match 'A_B')) 'cap familia sanejant amb "_" (era la discrepancia)'

# El parametre $titular de _LlicNomFitxer era MORT: el crider i les proves el
# passaven i el cos no el feia servir mai. Fora.
$srcLlic = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' 'Llicencia.ps1')))
Assert (-not ($srcLlic -match 'function _LlicNomFitxer\([^)]*titular')) '_LlicNomFitxer ja no te el parametre mort $titular'


Write-Host "`n--- Dos blocs mes que eren el mateix ---"
# 1. LA CARCASSA DE LA FINESTRA DE PROGRES AMB CANCEL.LAR era identica -23
#    linies, fins a les coordenades- a "Enviar correu (esborranys)" i a "Generar
#    informes": nomes canviava el titol. Ara es Show-ProgresCancel (UiComuns.ps1).
#    El que en fa una peca de debo es el detall delicat: el hashtable $cancel i
#    el FormClosing que converteix el tancament en cancel.lacio mentre la tanda
#    corre. Fer-ho malament no dona error, deixa el programa penjat.
#    Les altres dues finestres amb Cancel.lar NO hi entren, i es a posta:
#    "Copiar informes" (Informes.ps1) comenca en marquee i passa a continua a
#    mig cami despres d'una confirmacio, i la de "Word a PDF" (PdfSignar.ps1) es
#    d'una altra mida i porta la seva comptabilitat. Son parents, no la mateixa
#    finestra. Per aixo el guard no compta copies: diu QUI ha de passar-hi i qui
#    no pot tornar a muntar-se-la.
$srcProg = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' 'UiComuns.ps1')))
Assert ($srcProg.Contains('function Show-ProgresCancel')) 'la finestra de progres amb Cancel.lar viu a UiComuns.ps1'
foreach ($nom in @('ControlsCpEmail.ps1', 'ControlsPeriodics.ps1')) {
    $t = [System.IO.File]::ReadAllText((Join-Path $rootRepo (Join-Path 'suport' $nom)))
    Assert ($t.Contains('Show-ProgresCancel')) "$nom passa per Show-ProgresCancel"
    Assert (-not ($t.Contains('$cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true }'))) "$nom ja no es munta la carcassa pel seu compte"
}

# 2. EL MODEL DE LA TRIA. _SeccionsTriades converteix l'estat de les caselles en
#    la llista de seccions, i estava copiada linia a linia (42) a Paquet.ps1: la
#    pantalla del Pas 3 i el cami "des d'un paquet del mobil" muntaven el mateix
#    model pel seu compte. Si divergissin, el mobil i el PC generarien informes
#    DIFERENTS amb la mateixa tria. Es pura, o sigui que es pot provar aqui.
$secProva = @(
    [pscustomobject]@{ Title = 'Incendis'; Items = @(
        [pscustomobject]@{ Kind='intro'; Short='intro'; BodyLines=@('text'); Children=@() },
        [pscustomobject]@{ Kind='item'; Short='BIE'; BodyLines=@('cos'); Children=@(
            [pscustomobject]@{ Kind='child'; Short='fill1'; BodyLines=@('c1') },
            [pscustomobject]@{ Kind='child'; Short='fill2'; BodyLines=@('c2') }) },
        [pscustomobject]@{ Kind='item'; Short='Extintors'; BodyLines=@('cos'); Children=@() }
    )}
)
$cs = @{}
$cs[(_ItemKey 'Incendis' 'BIE')] = $true
$cs[(_ItemKey 'Incendis' 'BIE' 'fill2')] = $true
$tri = @(_SeccionsTriades $secProva $cs)
AssertEq $tri.Count 1 '_SeccionsTriades: nomes la seccio que te alguna cosa triada'
$itemsTri = @($tri[0].Items | Where-Object { $_.Kind -eq 'item' })
AssertEq $itemsTri.Count 1 '_SeccionsTriades: nomes l''item marcat'
AssertEq ([string]$itemsTri[0].Short) 'BIE' '_SeccionsTriades: i es el que toca'
AssertEq (@($itemsTri[0].Children).Count) 1 '_SeccionsTriades: nomes el fill marcat'
AssertEq ([string]@($itemsTri[0].Children)[0].Short) 'fill2' '_SeccionsTriades: i es el fill que toca'
Assert ([bool](@($tri[0].Items | Where-Object { $_.Kind -eq 'intro' }).Count -ge 1)) '_SeccionsTriades: l''intro de la seccio es conserva'
AssertEq (@(_SeccionsTriades $secProva @{}).Count) 0 '_SeccionsTriades: res marcat -> cap seccio'
