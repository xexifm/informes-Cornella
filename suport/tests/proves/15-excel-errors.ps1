# L'eina "Excel Errors" (ExcelErrors.ps1) i el mode AUTOMATIC dels "Controls
# periodics" (ControlsPeriodicsAutomatic.ps1 + l'historial de ControlsCpEmail.ps1),
# octubre 2026. Funcions PURES i el que es desa en disc, sense Excel ni Outlook.
# Es DOT-SOURCE des de run-tests.ps1 (mateix ambit i comptador d'asserts).

try {
Write-Host "`n--- Excel Errors: les columnes per la capcalera ---"
$exH = @('ID Activitat', 'UTM X', 'UTM Y', ('N' + [char]0x00FA + 'm. expedient'), ('Ra' + [char]0x00F3 + ' social'), ('Ra' + [char]0x00F3 + ' soc. NIF'),
         ('Ra' + [char]0x00F3 + ' soc. E-mail'), ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil'), 'Representant legal', 'Rep. Leg. NIF', 'Rep. Leg. E-mail',
         'Emp. Tipus via', 'Emp. Carrer', ('Emp. N' + [char]0x00FA + 'mero'), 'Ref. cadastral', ('Classificaci' + [char]0x00F3 + ' general annex'),
         ('Classificaci' + [char]0x00F3 + ' general Apartat'), 'Activitat principal', ('Data llic' + [char]0x00E8 + 'ncia/comunicaci' + [char]0x00F3),
         'Periodicitat CP', 'Proper CP previst')
$exCol = _ExColumnes $exH
AssertEq "$($exCol.ID)|$($exCol.UTMX)|$($exCol.UTMY)|$($exCol.EXP)|$($exCol.RAO)|$($exCol.NIF)|$($exCol.EMAIL)|$($exCol.MOBIL)|$($exCol.REP_NOM)|$($exCol.REP_NIF)|$($exCol.REP_EMAIL)" '1|2|3|4|5|6|7|8|9|10|11' 'Excel Errors: les columnes de les dades i del contacte, pel nom'
AssertEq "$($exCol.VIA)|$($exCol.CARRER)|$($exCol.NUMERO)|$($exCol.RC)|$($exCol.ANNEX)|$($exCol.APARTAT)|$($exCol.ACT)|$($exCol.LLIC)|$($exCol.PERIOD)|$($exCol.PROPER)" '12|13|14|15|16|17|18|19|20|21' 'Excel Errors: l''emplacament, el cadastre, la classificacio i el control periodic'
AssertEq "$($exCol.TELEFON)|$($exCol.REP_MOBIL)" '0|0' 'Excel Errors: una columna que no hi es val 0 (no es mira)'
# El respatller de Get-ColumnesContacte (el mobil a la 23, el correu a la 25)
# no pot fer mirar una columna que no es la del mobil.
$exCol2 = _ExColumnes @('ID Activitat', ('Ra' + [char]0x00F3 + ' social'))
AssertEq "$($exCol2.MOBIL)|$($exCol2.EMAIL)" '0|0' 'Excel Errors: sense les capcaleres, el respatller 23/25 no es fa servir'
Assert ((@(_ExColumnesQueFalten $exCol) -join ',') -notmatch 'UTM') 'Excel Errors: les columnes trobades no surten com a perdudes'
Assert ((@(_ExColumnesQueFalten $exCol2) -join ',').Contains('Ref. cadastral')) 'Excel Errors: les que falten si (el peu les diu)'

Write-Host "`n--- Excel Errors: les comprovacions d'una fila ---"
$exBona = @{ RAO = 'BAR EXEMPLE SL'; NIF = 'B12345674'; EMAIL = 'bar@exemple.cat'; MOBIL = '612345678'; REP_NOM = 'Pere Prova Exemple'; REP_NIF = '12345678Z'
             REP_EMAIL = ''; VIA = 'C/'; CARRER = 'Major'; NUMERO = '1'; RC = '1234567DF2813C0001AB'; UTMX = 424456.0; UTMY = '4578205,5'
             ANNEX = 'III'; APARTAT = '12.25'; ACT = 'Bar'; EXP = '2020/1'; LLIC = '01/01/2020'; PERIOD = ''; PROPER = '' }
AssertEq (@(Get-ExcelErrorsFila $exBona $exCol | ForEach-Object { $_.Tipus }) -join ',') 'control_periodic,control_periodic' 'Excel Errors: una fila bona nomes avisa del control periodic (annex III sense data prevista ni periodicitat)'
$exBona.PERIOD = '4 anys'; $exBona.PROPER = '01/01/2028'
AssertEq @(Get-ExcelErrorsFila $exBona $exCol).Count 0 'Excel Errors: i amb el control periodic, cap error'
$exFila = { param($canvis) $f = @{}; foreach ($k in $exBona.Keys) { $f[$k] = $exBona[$k] }; foreach ($k in $canvis.Keys) { $f[$k] = $canvis[$k] }; return @(Get-ExcelErrorsFila $f $exCol) }
$exT = { param($canvis) (@(& $exFila $canvis) | ForEach-Object { $_.Tipus + ':' + $_.Detall }) -join ' | ' }
AssertEq (& $exT @{ RAO = '' }) 'rao_social:buida' 'Excel Errors: falta la rao social'
AssertEq (& $exT @{ NIF = '' }) 'nif_titular:buit' 'Excel Errors: falta el NIF del titular'
Assert ((& $exT @{ NIF = '12345678A' }).StartsWith('nif_titular:la lletra del DNI')) 'Excel Errors: la lletra del DNI no quadra (_CtErrada)'
Assert ((& $exT @{ REP_NOM = ''; REP_NIF = '' }).StartsWith('rep_legal:buit, i el titular')) 'Excel Errors: una entitat sense representant legal'
AssertEq (& $exT @{ NIF = '12345678Z'; REP_NOM = ''; REP_NIF = '' }) '' 'Excel Errors: una persona fisica sense representant, no es cap error'
Assert ((& $exT @{ REP_NIF = '' }).StartsWith('rep_legal:buit (representant')) 'Excel Errors: representant sense NIF'
Assert ((& $exT @{ REP_NIF = 'B12345674' }).Contains('NIF del titular')) 'Excel Errors: el NIF de l''empresa al del representant'
Assert ((& $exT @{ REP_NIF = 'A58818501' }).Contains('entitat')) 'Excel Errors: un NIF d''entitat al del representant'
AssertEq (& $exT @{ NUMERO = ''; VIA = '' }) ('adreca:buit: Emp. Tipus via, Emp. N' + [char]0x00FA + 'mero') 'Excel Errors: l''adreca incompleta diu que falta'
AssertEq (& $exT @{ RC = '' }) 'cadastre:falta' 'Excel Errors: falta la referencia cadastral'
Assert ((& $exT @{ RC = '1234567DF2813' }).StartsWith('cadastre:t')) 'Excel Errors: una referencia cadastral massa curta'
AssertEq (& $exT @{ RC = '1234567DF2813C' }) '' 'Excel Errors: els 14 de la parcel.la valen'
AssertEq (& $exT @{ UTMX = $null; UTMY = '' }) 'utm:falten' 'Excel Errors: falten les UTM'
AssertEq (& $exT @{ UTMX = '4578205'; UTMY = '424456' }) 'utm:la X i la Y estan girades' 'Excel Errors: les UTM girades'
Assert ((& $exT @{ UTMX = '2.09'; UTMY = '41.35' }).StartsWith('utm:fora de Cornell')) 'Excel Errors: graus en lloc d''UTM -> fora de Cornella'
AssertEq (& $exT @{ UTMX = '424.456,2'; UTMY = '4.578.205' }) '' 'Excel Errors: les UTM amb punts de milers i coma decimal s''entenen'
AssertEq (& $exT @{ ANNEX = '' ; PERIOD = ''; PROPER = '' }) 'classificacio:buida' 'Excel Errors: falta la classificacio (i sense annex no hi ha control periodic)'
Assert ((& $exT @{ APARTAT = '' }).StartsWith('classificacio:buit (annex III)')) 'Excel Errors: annex sense apartat'
AssertEq (& $exT @{ ANNEX = 'L18 Cert'; APARTAT = '' }) '' 'Excel Errors: una Llei 18/2020 sense apartat no es cap error'
AssertEq (& $exT @{ ACT = '' }) 'activitat:buida' 'Excel Errors: falta l''activitat principal'
AssertEq (& $exT @{ EXP = '' }) 'expedient:buit' 'Excel Errors: falta l''expedient'
AssertEq (& $exT @{ LLIC = '' }) 'data_llicencia:buida' 'Excel Errors: falta la data de llicencia'
Assert ((& $exT @{ EMAIL = '' }).StartsWith("correu:no n'hi ha cap")) 'Excel Errors: sense cap correu'
Assert ((& $exT @{ EMAIL = 'bar@gmailo.com' }).StartsWith('correu:domini mal escrit')) 'Excel Errors: domini mal escrit (_CtErrada)'
AssertEq (@(& $exFila @{ EMAIL = 'a@b.cat; c@gmailo.com' }) | ForEach-Object { $_.Valor }) 'c@gmailo.com' 'Excel Errors: dos correus a la cel.la es miren un per un'
Assert ((& $exT @{ MOBIL = '934567890' }).Contains('un fix a la columna del m')) 'Excel Errors: un fix al mobil'
Assert ((& $exT @{ MOBIL = 'bar@exemple.cat' }).StartsWith('telefon:hi ha un correu')) 'Excel Errors: un correu al telefon'
AssertEq (& $exT @{ APARTAT = '5610'; ANNEX = 'I'; PERIOD = ''; PROPER = '' }) ('control_periodic:buit, i en t' + [char]0x00E9 + ' (apartat 561) | control_periodic:buida, i en t' + [char]0x00E9 + ' (apartat 561)') 'Excel Errors: el control periodic amb el criteri de Controls periodics (561)'
# Sense la columna, no es mira (i no surt cap error fals).
$exColSense = @{}; foreach ($k in $exCol.Keys) { $exColSense[$k] = $exCol[$k] }; $exColSense.RC = 0; $exColSense.UTMX = 0
AssertEq (@(Get-ExcelErrorsFila @{ RAO = 'X'; RC = ''; UTMX = $null; UTMY = $null } $exColSense | Where-Object { $_.Tipus -eq 'cadastre' -or $_.Tipus -eq 'utm' })).Count 0 'Excel Errors: una columna que no hi es no fa saltar cap error'

Write-Host "`n--- Excel Errors: totes les activitats, els duplicats i la base de contactes ---"
$exFiles = @(
    @{ Id = '10'; Valors = $exBona }
    @{ Id = '2'; Valors = @{ RAO = ''; NIF = 'B12345674'; EMAIL = 'x@y.cat' } }
    @{ Id = '10'; Valors = $exBona }
)
$exCt = @{ '2' = [pscustomobject]@{ avisos = @(
    [pscustomobject]@{ tipus = 'es_el_tecnic'; camp = 'Representant legal'; valor_excel = 'Joan Tecnic'; problema = 'es del tecnic'; proposta = '' },
    [pscustomobject]@{ tipus = 'error'; camp = 'E-mail'; valor_excel = 'x'; problema = 'repetit'; proposta = '' },
    [pscustomobject]@{ tipus = 'diferent'; camp = 'E-mail'; valor_excel = 'x@y.cat'; problema = 'els documents diuen'; proposta = 'z@y.cat' }) } }
$exErr = @(Get-ExcelErrors $exFiles $exCol $exCt)
AssertEq @($exErr | Where-Object { $_.Tipus -eq 'id_duplicat' }).Count 1 'Excel Errors: un ID repetit surt UNA vegada'
AssertEq (@($exErr | Where-Object { $_.Id -eq '2' -and @('tecnic', 'documents', 'correu') -contains $_.Tipus } | ForEach-Object { $_.Tipus }) -join ',') 'tecnic,documents' 'Excel Errors: els avisos de contactes hi entren (tecnic, documents) i els ''error'' no (sortirien dues vegades)'
Assert ((@($exErr | Where-Object { $_.Tipus -eq 'documents' })[0].Detall).Contains('proposta: z@y.cat')) 'Excel Errors: la proposta dels documents va al detall'
AssertEq (@(Get-ExcelErrors $exFiles $exCol $null) | Where-Object { $_.Tipus -eq 'tecnic' }).Count 0 'Excel Errors: sense base de contactes, cap avis del tecnic (el peu ho diu)'
Assert ((_ExPeu $exCol $false 3).Contains("Actualitzar base")) 'Excel Errors: el peu diu que sense la base de contactes no s''ha pogut mirar el tecnic'
$exN = _ExComptes $exErr
AssertEq "$($exN['id_duplicat'])|$($exN['tecnic'])|$($exN['telefon'])" '1|1|0' 'Excel Errors: els comptes per tipus (tambe els de 0)'
AssertEq @($exN.Keys).Count @($Script:ExTipus.Keys).Count 'Excel Errors: un compte per cada tipus'

Write-Host "`n--- Excel Errors: l'exportacio (nomes els tipus marcats) ---"
$exExp = @(_ExFilesExport $exErr @('rao_social', 'id_duplicat'))
AssertEq (@($exExp | ForEach-Object { $_[0] + '/' + $_[3] }) -join ' ; ') ('2/Falta la ra' + [char]0x00F3 + ' social ; 10/ID Activitat repetit') 'Excel Errors: nomes els tipus marcats, per ID (numeric) i amb el nom del tipus'
AssertEq @(_ExFilesExport $exErr @()).Count 0 'Excel Errors: sense cap tipus marcat, res'
AssertEq @(_ExFilesExport @($exErr | Select-Object -First 1) @('id_duplicat')).Count 1 'Excel Errors: una sola fila tambe es una fila'
AssertEq $Script:ExExportCapcalera[0] 'ID Activitat' 'Excel Errors: la primera columna es l''ID Activitat'
AssertEq ($Script:ExExportCapcalera -join '|') ('ID Activitat|Classificaci' + [char]0x00F3 + ' general annex|Classificaci' + [char]0x00F3 + " general Apartat|Tipus d'error|Camp|Valor a l'Excel|Detall") 'Excel Errors: despres de l''ID, la classificacio general annex i l''apartat; sense la rao social'
AssertEq @($Script:ExExportAmples).Count @($Script:ExExportCapcalera).Count 'Excel Errors: una amplada per columna'
AssertEq (@(_ExFilesExport $exErr @('tecnic'))[0] -join '|') ('2|||' + $Script:ExTipus['tecnic'] + '|Representant legal|Joan Tecnic|es del tecnic') 'Excel Errors: la fila exportada (activitat sense annex: buit)'
AssertEq @(_ExFilesExport $exErr @('id_duplicat'))[0][1] 'III' 'Excel Errors: la classificacio general annex de la base, a la segona columna'
AssertEq @(_ExFilesExport $exErr @('id_duplicat'))[0][2] '12.25' 'Excel Errors: i la classificacio general Apartat, a la tercera'
$exDir = Join-Path ([System.IO.Path]::GetTempPath()) ('exerr-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $exDir -Force)
try {
    $exXlsx = Export-ExcelErrors (Join-Path $exDir 'e.xlsx') $exErr @('rao_social')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $exZip = [System.IO.Compression.ZipFile]::OpenRead($exXlsx)
    try {
        $exSr = New-Object System.IO.StreamReader($exZip.GetEntry('xl/worksheets/sheet1.xml').Open())
        $exXml = $exSr.ReadToEnd(); $exSr.Dispose()
    } finally { $exZip.Dispose() }
    Assert ($exXml.Contains('Falta la ra') -and -not $exXml.Contains('ID Activitat repetit')) 'Excel Errors: l''Excel exportat porta nomes el que s''ha marcat'
    # LA TRIA DESADA: s'hi apunten els desmarcats; un tipus nou surt marcat.
    $exTp = Join-Path $exDir 'tria.json'
    AssertEq @(Read-ExTriats $exTp).Count @($Script:ExTipus.Keys).Count 'Excel Errors: la primera vegada, tots marcats'
    Save-ExTriats @('rao_social', 'utm') $exTp
    AssertEq (@(Read-ExTriats $exTp) -join ',') 'rao_social,utm' 'Excel Errors: la tria es recorda'
    Assert ([System.IO.File]::ReadAllText($exTp).Contains('desmarcats')) 'Excel Errors: es desen els desmarcats (un tipus nou surt marcat)'
} finally { Remove-Item -LiteralPath $exDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`n--- Excel Errors: la rajola i el despatxador ---"
$exMenu = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1')
Assert ($exMenu.Contains("Label = 'Excel Errors';    Kind = 'action'; Action = 'excelerrors'")) 'menu: la rajola Excel Errors (fila GIA)'
Assert ($exMenu.Contains("Label = 'Excel Comprovar'; Kind = 'action'; Action = 'comprovarexcel'")) 'menu: "Comprovar Excel" es diu ara "Excel Comprovar" (la mateixa accio: no perd el segell)'
Assert (-not $exMenu.Contains("'Comprovar Excel'")) 'menu: el nom vell ja no hi es'
Assert ((_SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1')).Contains("'excelerrors'    { Invoke-ExcelErrors }")) 'despatxador: excelerrors obre l''eina'
Assert ((_SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'ExcelErrors.ps1')) -notmatch 'Workbooks\.Open|\.Save\(\)') 'Excel Errors: no obre l''Excel per escriure-hi (l''Excel del GIA no es toca mai)'

Write-Host "`n--- Controls periodics AUTOMATIC: quines activitats toquen ---"
$cpAvui = New-Object datetime(2026, 10, 19)
$cpRow = { param($id, $data) [pscustomobject]@{ Id = $id; RaoSocial = ('T' + $id); ProperCP = $(if ($data) { $data.ToString('dd/MM/yyyy') } else { '' }); KProperCP = $data; Avisat = '' } }
$cpRows = @((& $cpRow '1' (New-Object datetime(2020, 1, 1))), (& $cpRow '2' (New-Object datetime(2026, 10, 1))), (& $cpRow '3' $cpAvui),
            (& $cpRow '4' (New-Object datetime(2027, 1, 1))), (& $cpRow '5' $null), (& $cpRow '6' (New-Object datetime(2026, 9, 1))))
$cpT = _CpAutoTriades $cpRows @{} $cpAvui 20
AssertEq (@($cpT.Triades | ForEach-Object { $_.Id }) -join ',') '2,6,1' 'CP auto: nomes els VENÇUTS (abans d''avui; ni el d''avui, ni els futurs, ni els sense data), primer els que han vençut fa menys'
AssertEq "$($cpT.Vencuts)|$($cpT.Pendents)" '3|0' 'CP auto: els comptes'
$cpT2 = _CpAutoTriades $cpRows @{} $cpAvui 2
AssertEq "$((@($cpT2.Triades | ForEach-Object { $_.Id })) -join ',')|$($cpT2.Pendents)" '2,6|1' 'CP auto: com a molt el maxim per passada; la resta, per a la propera'
$cpAv = @{ '2' = @{ ProperCp = '01/10/2026'; El = '2026-10-12T13:00:00'; Mode = 'auto' }; '6' = @{ ProperCp = '01/09/2025'; El = ''; Mode = 'manual' } }
AssertEq (@((_CpAutoTriades $cpRows $cpAv $cpAvui 20).Triades | ForEach-Object { $_.Id }) -join ',') '6,1' 'CP auto: un avisat del MATEIX Proper CP no es torna a preparar; amb una data nova, si'
AssertEq @((_CpAutoTriades @() $null $cpAvui 20).Triades).Count 0 'CP auto: sense files, res'

Write-Host "`n--- Controls periodics: l'historial dels avisos (anada i tornada pel JSON) ---"
$cpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('cpav-' + [guid]::NewGuid().ToString('N'))
try {
    $cpP = Join-Path $cpDir 'avisats.json'
    AssertEq @((Read-CpAvisats $cpP).Keys).Count 0 'CP historial: sense fitxer, buit'
    Assert (Add-CpAvisats @($cpRows[1], $cpRows[0]) 'auto' $cpP) 'CP historial: es desa (i crea la carpeta)'
    [void](Add-CpAvisats @($cpRows[5]) 'manual' $cpP)
    $cpL = Read-CpAvisats $cpP
    AssertEq "$($cpL['2'].ProperCp)|$($cpL['2'].Mode)|$($cpL['6'].Mode)|$(@($cpL.Keys).Count)" '01/10/2026|auto|manual|3' 'CP historial: el Proper CP i qui el va fer, sense perdre els d''abans'
    Assert (_CpJaAvisat $cpRows[1] $cpL) 'CP historial: _CpJaAvisat amb la mateixa data'
    AssertEq @((_CpAutoTriades $cpRows $cpL $cpAvui 20).Triades).Count 0 'CP historial: un cop avisats, l''automatic no hi torna'
    _CpMarcaAvisats $cpRows $cpL
    Assert ([string]$cpRows[1].Avisat).EndsWith('(A)') 'CP finestra: la columna "Correu preparat" diu la data i (A) si el va fer l''automatic'
    AssertEq "$($cpRows[3].Avisat)|$($cpRows[5].Avisat -ne '')" '|True' 'CP finestra: buida si no s''ha avisat; plena si a ma'
    AssertEq (Add-CpAvisats @() 'auto' $cpP) $false 'CP historial: sense files no es toca el fitxer'
} finally { Remove-Item -LiteralPath $cpDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`n--- Controls periodics AUTOMATIC: el registre i la programacio ---"
Assert ($Script:ModesAuto.Contains('controlsperiodics')) 'CP auto: te interruptor A/M (registre)'
AssertEq ([string](Get-ProgramacioAuto 'controlsperiodics' ([pscustomobject]@{})).Freq) 'setmana' 'CP auto: per defecte, cada setmana'
Assert ((Get-ProgramacioText (Get-ProgramacioAuto 'controlsperiodics' ([pscustomobject]@{}))).Contains('dilluns')) 'CP auto: el dilluns'
Assert ($Script:ScriptsSegonPla -contains 'ControlsPeriodicsAuto.ps1') 'CP auto: el boto Aturar el coneix'
$cpRgDir = Join-Path ([System.IO.Path]::GetTempPath()) ('cpreg-' + [guid]::NewGuid().ToString('N'))
$cpVell = $LocalActivitatsDir
$LocalActivitatsDir = $cpRgDir
try {
    $mCp = $Script:ModesAuto['controlsperiodics']
    AssertEq ([bool](& $mCp.Actiu)) $false 'CP auto: l''interruptor arrenca apagat'
    [void](& $mCp.DesaActiu $true)
    AssertEq ([bool](& $mCp.Actiu)) $true 'CP auto: DesaActiu / Actiu'
    Assert ((_CpAutoStatePath).EndsWith('controls-periodics-auto.json')) 'CP auto: l''estat, a local\base-dades-activitats'
    AssertEq ([string](& $mCp.Requisit)) '' 'CP auto: amb la carpeta de l''Excel es pot engegar'
    Assert (([string](& $mCp.TipA)).Contains('Esborranys')) 'CP auto: l''ajuda diu que nomes son esborranys'
} finally {
    $LocalActivitatsDir = $cpVell
    Remove-Item -LiteralPath $cpRgDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Controls periodics AUTOMATIC: nomes esborranys, i el boto i l'automatic pel mateix lloc (guard) ---"
$cpSup = Split-Path -Parent $TestsDir
$cpAutoSrc = _SenseComentaris (Join-Path $cpSup 'ControlsPeriodicsAutomatic.ps1')
$cpMailSrc = _SenseComentaris (Join-Path $cpSup 'ControlsCpEmail.ps1')
Assert ($cpAutoSrc.Contains('_CpPreparaEsborranys @($t.Triades) $cfgE $bcc ''auto''')) 'CP auto: prepara els esborranys amb la MATEIXA funcio que el boto'
Assert ($cpMailSrc.Contains('$res = _CpPreparaEsborranys $rows $cfgE $bcc ''manual''')) 'CP boto: tambe passa per _CpPreparaEsborranys'
Assert (-not ($cpAutoSrc -match 'Open-CorreuSessio|\.Send\(|Send-CorreuSessio')) 'CP auto: no obre cap sessio pel seu compte ni envia res'
Assert ($cpAutoSrc.Contains('Add-CorreuEsborranysPendents')) 'CP auto: el menu avisa dels esborranys que ha deixat'
AssertEq (@([regex]::Matches($cpMailSrc, 'Send-CorreuSessio')).Count) 1 'CP: un sol lloc fa els correus (no hi ha una copia per a l''automatic)'
} catch {
    Assert $false ("15-excel-errors: excepcio no controlada: " + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}

try {
Write-Host "`n--- Excel Errors: els TRES LLISTATS per correu (Innocues, Bars i Restaurants, Annexos II i III) ---"
# Les classificacions de debo dels Excels que va passar l'usuari (octubre 2026).
AssertEq (_ExLlistatDe 'L18-Proj. i Cert.' '521') 'innocues' 'Llistats: L18-Proj. i Cert. sense 561 -> Innocues'
AssertEq (_ExLlistatDe 'L18-Cert.' '855') 'innocues' 'Llistats: L18-Cert. -> Innocues'
AssertEq (_ExLlistatDe 'Sectorial' 'Centre de culte') 'innocues' 'Llistats: Sectorial -> Innocues'
AssertEq (_ExLlistatDe 'L18-Proj. i Cert.' '561') 'bars' 'Llistats: L18-Proj. i Cert. amb 561 -> Bars i Restaurants'
AssertEq (_ExLlistatDe 'L18-Proj. i Cert.' '5610') 'bars' 'Llistats: el 561 amb el criteri de Controls periodics (5610 tambe)'
AssertEq (_ExLlistatDe ('Recreativa cat' + [char]0x00E0 + 'leg') '') 'bars' 'Llistats: Recreativa cataleg -> Bars i Restaurants'
AssertEq (_ExLlistatDe 'II' '12.46') 'annexos' 'Llistats: annex II -> Annexos II i III'
AssertEq (_ExLlistatDe 'III' '561') 'annexos' 'Llistats: annex III (tambe amb 561) -> Annexos II i III'
AssertEq (_ExLlistatDe '' '') '' 'Llistats: sense annex -> cap'
AssertEq (_ExLlistatDe 'I.1' '1.1') '' 'Llistats: annex I -> cap'
AssertEq (_ExLlistatDe 'L18-Cert.' '561') '' 'Llistats: L18-Cert. amb 561 -> cap (no es cap dels tres)'
AssertEq (_ExLlistatDe 'Sectorial' '561') '' 'Llistats: Sectorial amb 561 -> cap'
$llE = { param($id, $annex, $apartat, $tipus) [pscustomobject]@{ Id = $id; Titular = ('T' + $id); Annex = $annex; Apartat = $apartat; Tipus = $tipus; TipusNom = $Script:ExTipus[$tipus]; Camp = 'c'; Valor = 'v'; Detall = 'd' } }
$llErr = @((& $llE '1' 'L18-Cert.' '855' 'cadastre'), (& $llE '2' 'L18-Proj. i Cert.' '561' 'cadastre'), (& $llE '3' 'II' '12.1' 'utm'),
           (& $llE '4' '' '' 'classificacio'), (& $llE '5' 'III' '3.2' 'cadastre'), (& $llE '1' 'L18-Cert.' '855' 'utm'))
$llPer = _ExPerLlistat $llErr @('cadastre', 'utm', 'classificacio') @()
AssertEq "$(@($llPer.innocues).Count)|$(@($llPer.bars).Count)|$(@($llPer.annexos).Count)|$((@($llPer.SenseLlistat) | ForEach-Object { $_.Id }) -join ',')" '2|1|2|4' 'Llistats: el repartiment, i el que no va a cap'
$llPer2 = _ExPerLlistat $llErr @('cadastre') @(@{ Id = '5' })
AssertEq "$(@($llPer2.innocues).Count)|$(@($llPer2.annexos).Count)" '1|0' 'Llistats: nomes els tipus marcats, i sense les activitats tretes'
AssertEq (_ExResumLlistat @($llPer.innocues)) '2 errors (1 activitats)' 'Llistats: el resum del correu'
AssertEq (_ExNomFitxerLlistat 'bars' (New-Object datetime(2026, 10, 15))) 'Excel_errors_2026-10-15_Bars_i_Restaurants.xlsx' 'Llistats: el nom del fitxer, com els de l''usuari'
AssertEq (_ExNomFitxerLlistat 'annexos' (New-Object datetime(2026, 10, 15))) 'Excel_errors_2026-10-15_Annexos_II_i_III.xlsx' 'Llistats: Annexos II i III'

Write-Host "`n--- Excel Errors: els tres Excels i les tretes dels llistats (en disc) ---"
$llDir = Join-Path ([System.IO.Path]::GetTempPath()) ('exll-' + [guid]::NewGuid().ToString('N'))
try {
    $llG = New-ExcelErrorsLlistats $llErr @('cadastre', 'utm') @() $llDir (New-Object datetime(2026, 10, 15))
    Assert ([bool]$llG.Ok) ('Llistats: es fan (' + [string]$llG.Error + ')')
    AssertEq (@($llG.Fitxers) | ForEach-Object { Split-Path -Leaf $_ }) 'Excel_errors_2026-10-15_Innocues.xlsx Excel_errors_2026-10-15_Bars_i_Restaurants.xlsx Excel_errors_2026-10-15_Annexos_II_i_III.xlsx' 'Llistats: els tres Excels, en ordre'
    AssertEq ((@($llG.Resum.Keys)) -join '|') (('Inn' + [char]0x00F2 + 'cues') + '|Bars i Restaurants|Annexos II i III') 'Llistats: el resum, per llistat'
    $llZip = [System.IO.Compression.ZipFile]::OpenRead(@($llG.Fitxers)[1])
    try { $llSr = New-Object System.IO.StreamReader($llZip.GetEntry('xl/worksheets/sheet1.xml').Open()); $llXml = $llSr.ReadToEnd(); $llSr.Dispose() } finally { $llZip.Dispose() }
    Assert ($llXml.Contains('>2<') -and -not $llXml.Contains('>3<') -and $llXml.Contains('ID Activitat')) 'Llistats: el de Bars i Restaurants porta nomes les seves activitats (i la capcalera)'
    $llP = Join-Path $llDir 'exclosos.json'
    AssertEq @(Read-ExExclosos $llP).Count 0 'Tretes: sense fitxer, cap'
    $llX = @(_ExExclososAmb @() @(@{ Id = '7'; Text = 'BAR' }, @{ Id = '3'; Text = 'X' }, @{ Id = '7'; Text = 'BAR' }))
    Save-ExExclosos $llX $llP
    AssertEq ((@(Read-ExExclosos $llP) | ForEach-Object { $_.Id }) -join ',') '3,7' 'Tretes: es desen (sense repetides, per GIA)'
    AssertEq ((@(_ExExclososAmb (Read-ExExclosos $llP) @(@{ Id = '3' }) $true) | ForEach-Object { $_.Id }) -join ',') '7' 'Tretes: es poden tornar a posar'
} finally { Remove-Item -LiteralPath $llDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`n--- Excel Errors AUTOMATIC: el registre, el correu i els guards ---"
Assert ($Script:ModesAuto.Contains('excelerrors')) 'Excel Errors auto: te interruptor A/M'
AssertEq ([string](Get-ProgramacioAuto 'excelerrors' ([pscustomobject]@{})).Freq) 'quinzena' 'Excel Errors auto: per defecte cada dues setmanes'
Assert ($Script:ScriptsSegonPla -contains 'ExcelErrorsAuto.ps1') 'Excel Errors auto: el boto Aturar el coneix'
Assert ($Script:CorreuEines.Contains('excelerrors')) 'Excel Errors: es una eina de Correus de cada eina'
AssertEq (@($Script:CorreuEines['excelerrors'].Vies) -join ',') 'outlook-esborrany,outlook' 'Excel Errors: nomes Outlook (porta adjunts)'
Assert ($Script:CorreuTextosEditors.Contains('excelerrors') -and $Script:CorreuProves.Contains('excelerrors')) 'Excel Errors: te el text editable i el correu de prova'
AssertEq ([string](_LoadExCorreu)['assumpte']) 'Dades incorrectes GIA' 'Excel Errors: l''assumpte que va proposar l''usuari'
Assert ([string](_ExCorreuDefecte)['cos']).StartsWith('Podeu mirar aquests errors del GIA si us plau??') 'Excel Errors: i el cos'
$exSup = Split-Path -Parent $TestsDir
$exAutoSrc = _SenseComentaris (Join-Path $exSup 'ExcelErrorsAutomatic.ps1')
Assert ($exAutoSrc.Contains('$caixa.Gen = Invoke-ExcelErrorsLlistats') -and $exAutoSrc.Contains('Send-ExcelErrorsCorreu $caixa.Gen $true')) 'Excel Errors auto: la MATEIXA feina que el boto (Invoke-ExcelErrorsLlistats + Send-ExcelErrorsCorreu)'
Assert ((_SenseComentaris (Join-Path $exSup 'ExcelErrorsLlistats.ps1')).Contains('if (-not $silenci -and @($ccoTriades).Count -gt 0) { $cco = _CorreuCcoAmb $cco $ccoTriades }')) 'Excel Errors: l''automatic no hi posa mai la CCO a triar'
Assert ((_SenseComentaris (Join-Path $exSup 'ExcelErrorsPantalla.ps1')).Contains("Select-CorreuCcoTria 'excelerrors'")) 'Excel Errors: a ma, la CCO a triar'
} catch {
    Assert $false ("15-excel-errors (llistats): excepcio no controlada: " + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}
