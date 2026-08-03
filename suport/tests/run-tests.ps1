# Proves automatiques de les funcions PURES del motor (Motor.ps1).
#
# NO prova la part de Word/Excel (COM) ni les finestres (WinForms): aixo
# nomes es pot provar a Windows amb Office. Aqui es validen les funcions de
# logica (parseig de camps, normalitzacio, cerca de columnes, dates, claus...).
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File tests/run-tests.ps1
#
# Carrega Motor.ps1 (nomes definicions) en mode "headless" (GENINFORME_TEST=1)
# perque no carregui WinForms. El motor no arrenca res per si sol: qui executa
# el programa es GenerarInforme.ps1, que aqui no toquem.

$ErrorActionPreference = 'Stop'
$env:GENINFORME_TEST = '1'
# A Linux no existeix LOCALAPPDATA; el donem perque el dot-source no falli.
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Motor.ps1'
. $scriptPath   # dot-source: defineix les funcions del motor

. (Join-Path $PSScriptRoot 'TestLib.ps1')   # Assert / AssertEq / Write-TestSummary

Write-Host "`n--- _NormalizeText ---"
AssertEq (_NormalizeText 'Estès')            'estes'           '_NormalizeText treu accents i minuscula'
AssertEq (_NormalizeText '  ÀCTIVITAT  ')    'activitat'       '_NormalizeText trim + accents'
AssertEq (_NormalizeText $null)              ''                '_NormalizeText null -> buit'

Write-Host "`n--- _CellToString ---"
AssertEq (_CellToString ([double]1000))      '1000'            '_CellToString enter sense decimals'
AssertEq (_CellToString ([double]12.5))      '12.5'            '_CellToString decimal'
AssertEq (_CellToString '  hola ')           'hola'            '_CellToString string trim'
AssertEq (_CellToString $null)               ''                '_CellToString null -> buit'

Write-Host "`n--- _FormatDateOnly ---"
# 45000 (serial OLE) = 2023-03-15. Comprovem nomes que NO porti hora i tingui forma de data.
$d = _FormatDateOnly ([double]45000)
Assert ($d -match '^\d{2}/\d{2}/\d{4}$') "_FormatDateOnly serial OLE -> dd/MM/yyyy (obtingut '$d')"
$d2 = _FormatDateOnly '15/03/2023 9:30:00'
Assert ($d2 -match '^\d{2}/\d{2}/\d{4}$') "_FormatDateOnly text amb hora -> nomes data (obtingut '$d2')"
AssertEq (_FormatDateOnly $null)             ''                '_FormatDateOnly null -> buit'

Write-Host "`n--- _ParseActivitatsDate (data de l'activitats.json del Drive) ---"
$pd1 = _ParseActivitatsDate ([pscustomobject]@{ SourceDate = '2026-06-01' })
AssertEq ($pd1.ToString('yyyy-MM-dd')) '2026-06-01' '_ParseActivitatsDate llegeix SourceDate'
$pd2 = _ParseActivitatsDate ([pscustomobject]@{ Source = '2026-05-29 ACTIVITATS.xlsx' })
AssertEq ($pd2.ToString('yyyy-MM-dd')) '2026-05-29' '_ParseActivitatsDate fallback al nom Source'
Assert ((_ParseActivitatsDate ([pscustomobject]@{ x = 1 })) -eq [datetime]::MinValue) '_ParseActivitatsDate sense data -> MinValue'
Assert ((_ParseActivitatsDate $null) -eq [datetime]::MinValue) '_ParseActivitatsDate null -> MinValue'

Write-Host "`n--- Test-ShouldExportActivitats (skip si el Drive ja esta al dia) ---"
$dLocal = [datetime]'2026-06-01'
Assert (-not (Test-ShouldExportActivitats $dLocal ([datetime]'2026-06-01'))) 'mateixa data -> NO exporta (skip)'
Assert (-not (Test-ShouldExportActivitats $dLocal ([datetime]'2026-06-05'))) 'Drive mes nou -> NO exporta'
Assert (Test-ShouldExportActivitats $dLocal ([datetime]'2026-05-20'))        'local mes nova -> exporta'
Assert (Test-ShouldExportActivitats $dLocal ([datetime]::MinValue))          'Drive sense base -> exporta'
Assert (Test-ShouldExportActivitats ([datetime]::MinValue) ([datetime]'2026-06-01')) 'sense data local fiable -> exporta (per seguretat)'

Write-Host "`n--- _ItemKey ---"
AssertEq (_ItemKey 'Sec' 'Item')             'Sec::Item'       '_ItemKey 2 parts'
AssertEq (_ItemKey 'Sec' 'Item' 'Fill')      'Sec::Item::Fill' '_ItemKey 3 parts'

Write-Host "`n--- _TextMatches ---"
Assert (_TextMatches 'Autoritzacions' 'auto')   '_TextMatches conte (case-insensitive)'
Assert (_TextMatches 'Sanitat' '')              '_TextMatches needle buit -> sempre cert'
Assert (-not (_TextMatches 'Sanitat' 'zzz'))    '_TextMatches no conte -> fals'

Write-Host "`n--- _FindColIndex (cerca columna per text de capcalera) ---"
# Construim una fila de capcalera 1-based com la que retorna Excel (.Value2).
$headers = @('ID Activitat','Num. expedient','Rao social','Num. registre entrada','Data registre entrada')
$cols = $headers.Count
$data = [Array]::CreateInstance([object], [int[]]@(1,$cols), [int[]]@(1,1))
for ($c = 1; $c -le $cols; $c++) { $data.SetValue($headers[$c-1], 1, $c) }
AssertEq (_FindColIndex $data $cols @('expedient') $null)               2 '_FindColIndex troba "expedient"'
AssertEq (_FindColIndex $data $cols @('registre','entrada') @('data'))  4 '_FindColIndex "registre entrada" exclou "data"'
AssertEq (_FindColIndex $data $cols @('data','registre','entrada') $null) 5 '_FindColIndex "data registre entrada"'
AssertEq (_FindColIndex $data $cols @('inexistent') $null)              0 '_FindColIndex no trobat -> 0'

Write-Host "`n--- Regressio: avisos de validacio NO han de petar amb Excel correcte ---"
# Reprodueix el patro que abans donava "method on NULL": construir l'ArrayList
# d'avisos a partir del validador i cridar .ToArray() encara que sigui buit.
$cols2 = 94
$data2 = [Array]::CreateInstance([object], [int[]]@(2,$cols2), [int[]]@(1,1))
foreach ($col in $Script:ActivitatsColumns) { $data2.SetValue($col.HeaderHint, 1, $col.Col) }
$data2.SetValue('1000', 2, 1)
$warnings = New-Object System.Collections.ArrayList
foreach ($w in (_ValidateActivitatsHeaders $data2 2 $cols2)) { if ($null -ne $w) { [void]$warnings.Add($w) } }
$arr = $null
try { $arr = $warnings.ToArray(); $ok = $true } catch { $ok = $false }
Assert $ok '_ValidateActivitatsHeaders + ArrayList: .ToArray() no peta amb capcaleres correctes'
AssertEq $warnings.Count 0 'Sense avisos quan totes les capcaleres coincideixen'

Write-Host "`n--- Get-FieldsFromSelection + Apply-Fields ([CAMP: ...]) ---"
$sec = [pscustomobject]@{
    Title = 'Sec'
    Items = @(
        [pscustomobject]@{ Kind='item'; Short='it1'; Selected=$true;
            BodyLines=@('Text amb [CAMP: Grup CAPCA] i [CAMP: Organ (ajuda)]');
            Children=@() }
    )
}
$fields = Get-FieldsFromSelection @($sec)
Assert ($fields.Contains('Grup CAPCA')) 'Get-FieldsFromSelection detecta "Grup CAPCA"'
Assert ($fields.Contains('Organ'))      'Get-FieldsFromSelection separa el hint entre parentesis'
AssertEq $fields['Organ'].Hint 'ajuda'  'Get-FieldsFromSelection captura el hint'
$fields['Grup CAPCA'].Value = 'ABC'
$fields['Organ'].Value = 'XYZ'
AssertEq (Apply-Fields 'a [CAMP: Grup CAPCA] b [CAMP: Organ (ajuda)] c' $fields) 'a ABC b XYZ c' 'Apply-Fields substitueix pels valors'

Write-Host "`n--- _SplitTextAndUrls (separar enllac del text de l'item) ---"
$r1 = _SplitTextAndUrls 'Instal·lacio de baixa tensio https://canalempresa.gencat.cat/baixa/'
AssertEq $r1.Text 'Instal·lacio de baixa tensio' '_SplitTextAndUrls separa el text'
AssertEq $r1.Urls.Count 1                         '_SplitTextAndUrls detecta 1 URL'
AssertEq $r1.Urls[0] 'https://canalempresa.gencat.cat/baixa/' '_SplitTextAndUrls captura l URL'
$r2 = _SplitTextAndUrls 'Text sense cap enllac'
AssertEq $r2.Text 'Text sense cap enllac'          '_SplitTextAndUrls sense URL -> tot text'
AssertEq $r2.Urls.Count 0                          '_SplitTextAndUrls sense URL -> 0 urls'
$r3 = _SplitTextAndUrls 'https://nomes.url/aqui'
AssertEq $r3.Text ''                               '_SplitTextAndUrls nomes URL -> text buit'
AssertEq $r3.Urls.Count 1                          '_SplitTextAndUrls nomes URL -> 1 url'
$r4 = _SplitTextAndUrls 'Veure http://a.cat/x i http://b.cat/y'
AssertEq $r4.Text 'Veure'                          '_SplitTextAndUrls text abans de diversos URLs'
AssertEq $r4.Urls.Count 2                          '_SplitTextAndUrls detecta 2 URLs'

Write-Host "`n--- Regressio: text de l'item i URL en paragrafs separats (build emit) ---"
# Dobles de les funcions Format (de Word) per capturar que reben, sense COM.
. (Join-Path $PSScriptRoot 'FormatDoubles.ps1')
$sel = [pscustomobject]@{}
$cfg = $Script:ReportFormatConfig
# Item amb el text i l'URL en linies SEPARADES (estructura real del REQ1).
$secs = @(
  [pscustomobject]@{ Title='Instal·lacions'; Items=@(
     [pscustomobject]@{ Kind='item'; Short='bt'; Selected=$true; Children=@();
        BodyLines=@('Instal·lació de baixa tensió','https://exemple.cat/baixa') }
  )}
)
$fbuild = Get-FieldsFromSelection $secs
$threw = $false
try { _WriteCatalegBody $sel $cfg $secs $fbuild '' } catch { $threw = $true; Write-Host "    EXCEPCIO: $($_.Exception.Message)" -ForegroundColor Red }
Assert (-not $threw) '_WriteCatalegBody no llanca excepcio (Substring) amb text+URL separats'
$itemCall = $global:emitCalls | Where-Object { $_ -like 'ITEM|*' } | Select-Object -First 1
$urlCall  = $global:emitCalls | Where-Object { $_ -like 'URL|*' }  | Select-Object -First 1
AssertEq $itemCall 'ITEM|1.|Instal·lació de baixa tensió' "Format-Item rep nomes el text de l'item (sense URL)"
AssertEq $urlCall  'URL|https://exemple.cat/baixa'        "Format-Url rep l'URL en paragraf propi"

Write-Host "`n--- [OPCIO: ...] desplegables ---"
$op = _ParseOpcio "Destinatari | a l'ajuntament (Annex II) | a l'ACA"
AssertEq $op.Name 'Destinatari'                       '_ParseOpcio extreu el nom'
AssertEq $op.Options.Count 2                          '_ParseOpcio extreu 2 opcions'
AssertEq $op.Options[0] "a l'ajuntament (Annex II)"   '_ParseOpcio opcio 1 (amb parentesis)'
AssertEq $op.Options[1] "a l'ACA"                      '_ParseOpcio opcio 2'

$secO = [pscustomobject]@{
    Title = 'Sec'
    Items = @(
        [pscustomobject]@{ Kind='item'; Short='it'; Selected=$true; Children=@();
            BodyLines=@('Presentar projecte [OPCIO: Destinatari | ajuntament | ACA] amb el contingut.') }
    )
}
$fo = Get-FieldsFromSelection @($secO)
Assert ($fo.Contains('Destinatari'))                 'Get-FieldsFromSelection detecta el desplegable'
AssertEq $fo['Destinatari'].Type 'choice'            'El camp es de tipus choice'
AssertEq $fo['Destinatari'].Options.Count 2          'El desplegable te 2 opcions'
AssertEq $fo['Destinatari'].Value 'ajuntament'       'Per defecte tria la primera opcio'
$fo['Destinatari'].Value = 'ACA'
AssertEq (Apply-Fields 'Presentar projecte [OPCIO: Destinatari | ajuntament | ACA] amb el contingut.' $fo) 'Presentar projecte ACA amb el contingut.' 'Apply-Fields substitueix el desplegable per l opcio triada'

Write-Host "`n--- _SplitTextAndUrls amb marcador [[URL]] (estil Cita) ---"
$cita1 = _SplitTextAndUrls '[[URL]] https://example.com/path'
AssertEq $cita1.Text ''                                '[[URL]] tota la linia es URL: Text buit'
AssertEq $cita1.Urls.Count 1                           '[[URL]] genera 1 url'
AssertEq $cita1.Urls[0] 'https://example.com/path'     '[[URL]] captura l URL'
# Sense http (per quan l'usuari posi un text amb estil Cita): tambe ha de funcionar
$cita2 = _SplitTextAndUrls '[[URL]] www.exemple.cat'
AssertEq $cita2.Urls.Count 1                           '[[URL]] tambe accepta www sense http'
AssertEq $cita2.Urls[0] 'www.exemple.cat'              '[[URL]] captura www correctament'
# Cas retrocompatible: sense marcador, http al mig (continua funcionant)
$noflag = _SplitTextAndUrls 'Veure https://x.cat/y al web'
AssertEq $noflag.Text 'Veure'                          'Sense [[URL]]: text abans de http (retrocompatibilitat)'
AssertEq $noflag.Urls.Count 1                          'Sense [[URL]]: 1 url detectada per contingut'

Write-Host "`n--- Fills sense numeracio (build emit) ---"
. (Join-Path $PSScriptRoot 'FormatDoubles.ps1')   # reinicia $global:emitCalls

$sec3 = [pscustomobject]@{ Title='X'; Items=@(
  [pscustomobject]@{ Kind='item'; Short='pare'; Selected=$true;
    BodyLines=@('Text pare.');
    Children=@(
      [pscustomobject]@{ Kind='child'; Short='f1'; BodyLines=@('Fill A.'); Children=@() }
      [pscustomobject]@{ Kind='child'; Short='f2'; BodyLines=@('Fill B.'); Children=@() }
    )
  }
)}
$f3 = Get-FieldsFromSelection @($sec3)
$global:emitCalls.Clear()
_WriteCatalegBody ([pscustomobject]@{}) $Script:ReportFormatConfig @($sec3) $f3 ''
$itemCalls = @($global:emitCalls | Where-Object { $_ -like 'ITEM|*' })
$bulletChildCalls = @($global:emitCalls | Where-Object { $_ -like 'BULLET/CH*' })
AssertEq $itemCalls.Count 1                "Nomes 1 ITEM (el pare); cap ITEM per als fills"
AssertEq $bulletChildCalls.Count 2         "Els 2 fills surten com a BULLET (punt de llista) amb sagnia de fill"
AssertEq $bulletChildCalls[0] 'BULLET/CH/1r|Fill A.' "Primer fill: punt de llista, marcat -First (se separa de l'item)"
AssertEq $bulletChildCalls[1] 'BULLET/CH|Fill B.'    "Segon fill: punt de llista, separacio normal entre punts"
# Cap fill ha de tenir patro de numeracio "X.Y."
$childHasNum = @($global:emitCalls | Where-Object { $_ -match 'ITEM\|\d+\.\d+\.' }).Count
AssertEq $childHasNum 0                    "Cap fill amb numeracio jerarquica (X.Y.)"

Write-Host "`n--- Sangries de la vinyeta d'un fill (twips del Word) ---"
# Els valors del cataleg es guarden en cm, pero el que es veu al .docx son
# twips. Aquests son els que ha de produir el Word (mateixos que el document de
# referencia corregit a ma: w:ind left="567" hanging="283").
$cfgB = $Script:ReportFormatConfig
AssertEq (_CmToTwips $cfgB.BulletChildIndentCm) 567 'Fill: sangria esquerra 1 cm = 567 twips'
AssertEq (_CmToTwips $cfgB.BulletChildHangCm)   283 'Fill: sangria francesa 0,5 cm = 283 twips'
AssertEq (_CmToTwips $cfgB.ChildIndentCm)       567 'Fill: la vinyeta s alinea amb les sub-linies i els enllacos del fill'
AssertEq (_PtToTwips $cfgB.ItemSpaceAfterPt)    240 'Fill: separacio amb l item = 12 pt = 240 twips'
AssertEq (_PtToTwips $cfgB.BulletSpaceBeforePt) 120 'Fill: separacio entre punts = 6 pt = 120 twips'

Write-Host "`n--- OPCIO se substitueix dins emit (regressio) ---"
$global:emitCalls.Clear()
$secO = [pscustomobject]@{ Title='Y'; Items=@(
  [pscustomobject]@{ Kind='item'; Short='it'; Selected=$true; Children=@();
    BodyLines=@('Presentar a [OPCIO: Lloc | ajuntament | ACA] amb el contingut.')
  }
)}
$fO = Get-FieldsFromSelection @($secO)
$fO['Lloc'].Value = 'ACA'
_WriteCatalegBody ([pscustomobject]@{}) $Script:ReportFormatConfig @($secO) $fO ''
$itemCalls = @($global:emitCalls | Where-Object { $_ -like 'ITEM|*' })
$itemTxt = [string]$itemCalls[0]
Assert ($itemTxt -notmatch '\[OPCIO:')        "L'OPCIO NO ha de quedar literal al text de l'item"
Assert ($itemTxt.Contains('ACA'))             "El valor triat (ACA) apareix al text de l'item"
Assert ($itemTxt.Contains('Presentar a ACA')) "L'OPCIO substituit forma una frase coherent"

Write-Host "`n--- Subseccions buides NO s'emeten (regressio) ---"
# Cas real: secció "Instal·lacions" amb 3 ::SUB:: (Legalitzacions,
# Inspeccions inicials, Inspeccions periòdiques). L'usuari només
# selecciona 1 item sota "Inspeccions periòdiques". Han de sortir:
#   SECT "Instal·lacions"
#   SUB  "Inspeccions periòdiques"
#   ITEM "PCI..."
# NO han de sortir "Legalitzacions" ni "Inspeccions inicials".
$global:emitCalls = New-Object System.Collections.ArrayList
$secSub = [pscustomobject]@{ Title='Instal·lacions'; Items=@(
  [pscustomobject]@{ Kind='subsection'; Short='Legalitzacions';        BodyLines=@(); Children=@(); Selected=$false }
  [pscustomobject]@{ Kind='subsection'; Short='Inspeccions inicials';  BodyLines=@(); Children=@(); Selected=$false }
  [pscustomobject]@{ Kind='subsection'; Short='Inspeccions periòdiques'; BodyLines=@(); Children=@(); Selected=$false }
  [pscustomobject]@{ Kind='item'; Short='PCI'; Selected=$true; Children=@();
    BodyLines=@('PCI. Real Decreto 513/2017.')
  }
)}
_WriteCatalegBody ([pscustomobject]@{}) $Script:ReportFormatConfig @($secSub) ([ordered]@{}) ''
$subCalls = @($global:emitCalls | Where-Object { $_ -like 'SUB|*' })
AssertEq $subCalls.Count 1                              'Nomes 1 subseccio emesa (la de l item triat)'
AssertEq $subCalls[0]    'SUB|Inspeccions periòdiques' 'La subseccio emesa es la correcta'
Assert (-not ($subCalls -contains 'SUB|Legalitzacions'))       'Legalitzacions NO surt (no te items triats)'
Assert (-not ($subCalls -contains 'SUB|Inspeccions inicials')) 'Inspeccions inicials NO surt'

# Segon cas: 2 subseccions diferents, cada una amb un item triat.
# Han de sortir TOTES DUES (i en l'ordre correcte).
$global:emitCalls.Clear()
$secMix = [pscustomobject]@{ Title='Instal·lacions'; Items=@(
  [pscustomobject]@{ Kind='subsection'; Short='Legalitzacions';        BodyLines=@(); Children=@(); Selected=$false }
  [pscustomobject]@{ Kind='item'; Short='A'; Selected=$true; Children=@(); BodyLines=@('Text A.') }
  [pscustomobject]@{ Kind='subsection'; Short='Inspeccions periòdiques'; BodyLines=@(); Children=@(); Selected=$false }
  [pscustomobject]@{ Kind='item'; Short='B'; Selected=$true; Children=@(); BodyLines=@('Text B.') }
)}
_WriteCatalegBody ([pscustomobject]@{}) $Script:ReportFormatConfig @($secMix) ([ordered]@{}) ''
$subCalls2 = @($global:emitCalls | Where-Object { $_ -like 'SUB|*' })
AssertEq $subCalls2.Count 2                          'Les 2 subseccions amb items triats SI surten'
AssertEq $subCalls2[0] 'SUB|Legalitzacions'          'Primera subseccio en l ordre del doc'
AssertEq $subCalls2[1] 'SUB|Inspeccions periòdiques' 'Segona subseccio en l ordre del doc'

Write-Host "`n--- _GetUniqueOutputPath (sufixos _2, _3 quan ja existeix) ---"
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("uniquetest_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir | Out-Null
try {
    $p1 = _GetUniqueOutputPath $tmpDir '2026-05-29_Req1_GIA 1379.docx'
    AssertEq ([System.IO.Path]::GetFileName($p1)) '2026-05-29_Req1_GIA 1379.docx' 'Primer informe: nom base sense sufix'
    # "Creem" el primer perque el seguent l'hagi d'evitar
    Set-Content -LiteralPath $p1 -Value 'x'
    $p2 = _GetUniqueOutputPath $tmpDir '2026-05-29_Req1_GIA 1379.docx'
    AssertEq ([System.IO.Path]::GetFileName($p2)) '2026-05-29_Req1_GIA 1379_2.docx' 'Segon informe del mateix dia: sufix _2'
    Set-Content -LiteralPath $p2 -Value 'x'
    $p3 = _GetUniqueOutputPath $tmpDir '2026-05-29_Req1_GIA 1379.docx'
    AssertEq ([System.IO.Path]::GetFileName($p3)) '2026-05-29_Req1_GIA 1379_3.docx' 'Tercer informe del mateix dia: sufix _3'
    # GIA diferent ha de quedar net (no afectat)
    $p4 = _GetUniqueOutputPath $tmpDir '2026-05-29_Req1_GIA 9999.docx'
    AssertEq ([System.IO.Path]::GetFileName($p4)) '2026-05-29_Req1_GIA 9999.docx' 'GIA diferent: sense sufix'
} finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- _BuildOrigenText (linia Objecte: segons origen triat al Pas 2) ---"
$uG = [char]0x00FA; $oG = [char]0x00F3; $apG = [char]0x2019
AssertEq (_BuildOrigenText @{ ORIGEN_TIPUS='doc'; NUM_ANOTACIO='A-123'; DATA_ANOTACIO='10/06/2025' }) `
    ("Doc. aportada amb N${uG}m. d${apG}anotaci${oG} A-123 del 10/06/2025") '_BuildOrigenText doc -> Doc. aportada amb Num. anotacio'
AssertEq (_BuildOrigenText @{ ORIGEN_TIPUS='insp'; DATA_INSPECCIO='18/07/2026' }) `
    ("Visita inspecci${oG} 18/07/2026") '_BuildOrigenText insp -> Visita inspeccio DATA'
AssertEq (_BuildOrigenText @{ NUM_ANOTACIO='A-9'; DATA_ANOTACIO='01/01/2026' }) `
    ("Doc. aportada amb N${uG}m. d${apG}anotaci${oG} A-9 del 01/01/2026") '_BuildOrigenText sense tipus -> doc per defecte'
AssertEq (_BuildOrigenText @{ ORIGEN_TIPUS='insp'; DATA_INSPECCIO='' }) `
    ("Visita inspecci${oG} ") '_BuildOrigenText insp sense data -> nomes el prefix'
AssertEq (_BuildOrigenText @{ ORIGEN_TIPUS='cap' }) '' '_BuildOrigenText cap -> buit (sense Objecte)'

Write-Host "`n--- Camps detectats a conclusions (Add-FieldsFromConclusions) ---"
$fc = [ordered]@{}
$selectedConcl = @(
    [pscustomobject]@{ Title='C1'; Body='Termini [OPCIO: Mesos | un | dos] segons cas.' }
)
$alwaysConcl = @('Ho poso al seu coneixement. [CAMP: lloc]')
Add-FieldsFromConclusions $fc $selectedConcl $alwaysConcl
AssertEq $fc.Count 2                       'Detectats 2 camps a les conclusions (OPCIO + CAMP)'
Assert ($fc.Contains('Mesos'))             "Detectat l'OPCIO 'Mesos'"
AssertEq $fc['Mesos'].Type 'choice'        "'Mesos' es desplegable"
AssertEq $fc['Mesos'].Options.Count 2      "'Mesos' te 2 opcions"
Assert ($fc.Contains('lloc'))              "Detectat el CAMP 'lloc'"
AssertEq $fc['lloc'].Type 'text'           "'lloc' es camp de text"

# Tambe ha de fer servir el cos ($c.Body) i no el titol
$fc2 = [ordered]@{}
Add-FieldsFromConclusions $fc2 @([pscustomobject]@{ Title='[CAMP: NoVull]'; Body='Cap camp aqui.' }) @()
AssertEq $fc2.Count 0                      "Add-FieldsFromConclusions IGNORA el Title (nomes mira Body)"

Write-Host "`n--- Type-RichText (negreta **text** i cursiva //text//) ---"
# Mock minim de $sel per capturar el que es teclejaria
$global:typed = New-Object System.Collections.ArrayList
$selMock = New-Object PSObject -Property @{
    Font = New-Object PSObject -Property @{ Bold = 0; Italic = 0 }
}
Add-Member -InputObject $selMock -MemberType ScriptMethod -Name TypeText -Value {
    param($t)
    $b = if ($this.Font.Bold -eq 1) { 'B' } else { '-' }
    $i = if ($this.Font.Italic -eq 1) { 'I' } else { '-' }
    [void]$global:typed.Add(("{0}{1}:{2}" -f $b, $i, $t))
}

# Cas 1: tot text normal
$global:typed.Clear()
Type-RichText $selMock 'text normal'
AssertEq ($global:typed -join '|') '--:text normal' 'Type-RichText: text normal'

# Cas 2: negreta entremig
$global:typed.Clear()
Type-RichText $selMock 'abans **negreta** despres'
AssertEq ($global:typed -join '|') '--:abans |B-:negreta|--: despres' 'Type-RichText: negreta entremig'

# Cas 3: cursiva al final
$global:typed.Clear()
Type-RichText $selMock 'abans //cursiva//'
AssertEq ($global:typed -join '|') '--:abans |-I:cursiva' 'Type-RichText: cursiva al final'

# Cas 4: barreja
$global:typed.Clear()
Type-RichText $selMock '**neg** mig //cur//'
AssertEq ($global:typed -join '|') 'B-:neg|--: mig |-I:cur' 'Type-RichText: barreja de negreta i cursiva'

# Cas 5: cap marcador, text buit
$global:typed.Clear()
Type-RichText $selMock ''
AssertEq $global:typed.Count 0 'Type-RichText: text buit no fa res'

Write-Host "`n--- Find-LatestActivitatsExcel (primary -> fallback local) ---"
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("xlsdb_" + [Guid]::NewGuid().ToString('N'))
$tmpPrimary  = Join-Path $tmpRoot 'primary'
$tmpFallback = Join-Path $tmpRoot 'fallback'
New-Item -ItemType Directory -Path $tmpPrimary, $tmpFallback | Out-Null
try {
    # Sobreescrivim les variables que llegeix Find-LatestActivitatsExcel
    $savedAD = $ActivitatsDir
    $savedLD = $LocalActivitatsDir
    Set-Variable -Name ActivitatsDir      -Value $tmpPrimary  -Scope Script
    Set-Variable -Name LocalActivitatsDir -Value $tmpFallback -Scope Script

    # Cas A: cap fitxer enlloc
    Assert ($null -eq (Find-LatestActivitatsExcel)) 'A: cap fitxer enlloc -> null'

    # Cas B: nomes al fallback -> Source=fallback
    Set-Content -LiteralPath (Join-Path $tmpFallback '2023-01-15 ACTIVITATS.xlsx') -Value 'x'
    $rB = Find-LatestActivitatsExcel
    Assert ($null -ne $rB)               'B: trobat (fallback)'
    AssertEq $rB.Source 'fallback'       'B: Source = fallback'
    AssertEq $rB.File.Name '2023-01-15 ACTIVITATS.xlsx' 'B: fitxer correcte'

    # Cas C: hi ha al primary (encara que el fallback tambe en tingui) -> primary guanya
    Set-Content -LiteralPath (Join-Path $tmpPrimary '2024-03-20 ACTIVITATS.xls') -Value 'x'
    Set-Content -LiteralPath (Join-Path $tmpPrimary '2024-08-01 ACTIVITATS.xlsx') -Value 'x'
    $rC = Find-LatestActivitatsExcel
    AssertEq $rC.Source 'primary'                       'C: primary guanya sobre fallback'
    AssertEq $rC.File.Name '2024-08-01 ACTIVITATS.xlsx' 'C: agafa el mes recent del primary'

    # Cas D: primary inaccessible -> cau al fallback
    Remove-Item -LiteralPath $tmpPrimary -Recurse -Force
    $rD = Find-LatestActivitatsExcel
    AssertEq $rD.Source 'fallback'  'D: primary no existeix -> fallback'

    Set-Variable -Name ActivitatsDir      -Value $savedAD -Scope Script
    Set-Variable -Name LocalActivitatsDir -Value $savedLD -Scope Script
} finally {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Seguiment: _ClassifyParagraph ---"
$cReq = _ClassifyParagraph '1. Baixa tensio. Cal legalitzar.' ''
AssertEq $cReq.Kind 'requirement'      '_ClassifyParagraph requeriment literal -> requirement'
AssertEq $cReq.Number 1                '_ClassifyParagraph captura el numero del requeriment'
Assert (-not $cReq.ViaList)            '_ClassifyParagraph requeriment literal: ViaList=false'
$cAnn = _ClassifyParagraph "01/06/2026: No s'entrega." ''
AssertEq $cAnn.Kind 'annotation'       '_ClassifyParagraph anotacio datada -> annotation'
AssertEq $cAnn.Date '01/06/2026'       '_ClassifyParagraph normalitza la data'
$cAnn2 = _ClassifyParagraph '3/6/2026: ok' ''
AssertEq $cAnn2.Date '03/06/2026'      '_ClassifyParagraph normalitza data d/M -> dd/MM'
$cOther = _ClassifyParagraph 'Text qualsevol sense numero' ''
AssertEq $cOther.Kind 'other'          '_ClassifyParagraph text normal -> other'
$cList = _ClassifyParagraph 'Baixa tensio sense numero literal' '1.'
AssertEq $cList.Kind 'requirement'     '_ClassifyParagraph auto-numerat per ListString -> requirement'
AssertEq $cList.Number 1               '_ClassifyParagraph ListString aporta el numero'
Assert ($cList.ViaList)                '_ClassifyParagraph auto-numerat: ViaList=true'

Write-Host "`n--- Seguiment: _InferResolvedFromBold / _ShouldBeBold ---"
Assert (_InferResolvedFromBold 0)          '_InferResolvedFromBold 0 (res negreta) -> resolt'
Assert (-not (_InferResolvedFromBold -1))  '_InferResolvedFromBold -1 (tot negreta) -> pendent'
Assert (-not (_InferResolvedFromBold 9999999)) '_InferResolvedFromBold 9999999 (mixt) -> pendent'
Assert (_ShouldBeBold $false)              '_ShouldBeBold: no resolt -> negreta'
Assert (-not (_ShouldBeBold $true))        '_ShouldBeBold: resolt -> sense negreta'

Write-Host "`n--- Seguiment: _ShortenText / _SeguimentParentTopic ---"
AssertEq (_ShortenText 'abc' 10) 'abc' '_ShortenText curt -> igual'
AssertEq (_ShortenText 'abcdefghij' 5) ('abcde' + [char]0x2026) '_ShortenText llarg -> talla amb ...'
AssertEq (_SeguimentParentTopic '1. Elements de coccio. Segons article 11.') 'Elements de coccio' '_SeguimentParentTopic treu el num i agafa la 1a frase'

Write-Host "`n--- Seguiment: _BuildSeguimentModel (fills / sub-punts) ---"
$recs = @(
    [pscustomobject]@{ Index=1; Text='1. Elements de coccio. Segons article.'; ListString=''; Bold=-1; IsBulletChild=$false }
    [pscustomobject]@{ Index=2; Text='Retirar la fregidora.';                   ListString=''; Bold=-1; IsBulletChild=$true }
    [pscustomobject]@{ Index=3; Text='Presentar certificat.';                   ListString=''; Bold=-1; IsBulletChild=$true }
    [pscustomobject]@{ Index=4; Text='2. Calor. Incompleix article.';           ListString=''; Bold=-1; IsBulletChild=$false }
)
$m = _BuildSeguimentModel $recs
$u = @($m.Requirements)
AssertEq $u.Count 3 '_BuildSeguimentModel: req amb 2 fills + req sense fills -> 3 unitats'
AssertEq ([bool]$u[0].IsChild) $true  '_BuildSeguimentModel: la 1a unitat es un fill'
AssertEq $u[0].ParaIndex 2            '_BuildSeguimentModel: el fill apunta al seu paragraf (2)'
AssertEq ([bool]($u[0].Label -like 'Req. 1 (*): Retirar*')) $true '_BuildSeguimentModel: etiqueta del fill amb num i tema del pare'
AssertEq ([bool]$u[2].IsChild) $false '_BuildSeguimentModel: el req sense fills NO es child'
AssertEq $u[2].ParaIndex 4            '_BuildSeguimentModel: req sense fills apunta al seu paragraf (4)'
AssertEq $m.LastReqParaIndex 4        '_BuildSeguimentModel: LastReqParaIndex = ultim paragraf de unitat'
# Les anotacions van al fill correcte i en determinen l'estat.
$recs2 = @(
    [pscustomobject]@{ Index=1; Text='1. Punt amb fills.';    ListString=''; Bold=-1; IsBulletChild=$false }
    [pscustomobject]@{ Index=2; Text='Fill A.';               ListString=''; Bold=-1; IsBulletChild=$true }
    [pscustomobject]@{ Index=3; Text="01/06/2026: S'aporta."; ListString=''; Bold=0;  IsBulletChild=$false }
    [pscustomobject]@{ Index=4; Text='Fill B.';               ListString=''; Bold=-1; IsBulletChild=$true }
)
$u2 = @((_BuildSeguimentModel $recs2).Requirements)
AssertEq $u2.Count 2                      '_BuildSeguimentModel: 1 req amb 2 fills -> 2 unitats'
AssertEq (@($u2[0].Annotations).Count) 1  '_BuildSeguimentModel: l''anotacio va al fill A'
AssertEq ([bool]$u2[0].WasResolved) $true '_BuildSeguimentModel: fill A amb anotacio no-negreta -> resolt'
AssertEq ([bool]$u2[1].WasResolved) $false '_BuildSeguimentModel: fill B sense anotacio -> pendent'

Write-Host "`n--- Seguiment: _FindConclusionStartIndex ---"
$pt = @(
    '1. Baixa tensio. Vist l anterior cal aportar.',  # frase DINS un requeriment (no ha de disparar)
    '01/06/2026: No s entrega.',
    "Vist l'anterior, cal requerir l'esmena.",          # inici real de conclusions (paragraf 3)
    'Ho poso al seu coneixement als efectes oportuns,',
    'Cornella de Llobregat,'
)
AssertEq (_FindConclusionStartIndex $pt 1 $SeguimentConclusionPhrases) 3 '_FindConclusionStartIndex ancora despres de l ultim requeriment'
$ptAcc = @('1. req', "CORNELLÀ DE LLOBREGAT,")
AssertEq (_FindConclusionStartIndex $ptAcc 1 @('Cornella de Llobregat,')) 2 '_FindConclusionStartIndex insensible a accents/majuscules'
$ptNone = @('1. req', 'res a veure aqui')
AssertEq (_FindConclusionStartIndex $ptNone 1 $SeguimentConclusionPhrases) -1 '_FindConclusionStartIndex sense frase -> -1'

Write-Host "`n--- Seguiment: _ValidateRoundDate ---"
$vd1 = _ValidateRoundDate '01/06/2026'
Assert $vd1.Ok                              '_ValidateRoundDate data valida -> Ok'
AssertEq $vd1.Normalized '01/06/2026'       '_ValidateRoundDate conserva la data valida'
$vd2 = _ValidateRoundDate '3/6/2026'
AssertEq $vd2.Normalized '03/06/2026'       '_ValidateRoundDate normalitza d/M/yyyy'
$vd3 = _ValidateRoundDate '32/13/2026'
Assert (-not $vd3.Ok)                       '_ValidateRoundDate data impossible -> no Ok'
$vd4 = _ValidateRoundDate ''
Assert $vd4.Ok                              '_ValidateRoundDate buit -> Ok (avui)'
Assert ($vd4.Normalized -match '^\d{2}/\d{2}/\d{4}$') '_ValidateRoundDate buit -> forma dd/MM/yyyy'

Write-Host "`n--- Seguiment: _FormatAnnotationLine ---"
AssertEq (_FormatAnnotationLine '01/06/2026' "No s'entrega.") "01/06/2026: No s'entrega." '_FormatAnnotationLine compon "data: comentari"'
AssertEq (_FormatAnnotationLine '01/06/2026' '  espais  ') '01/06/2026: espais' '_FormatAnnotationLine fa trim del comentari'

Write-Host "`n--- Seguiment: _SeguimentOutputName ---"
$d = [datetime]'2026-06-05'
AssertEq (_SeguimentOutputName '2026-05-29_Req1_GIA 1379' $d) '2026-06-05_Req2_GIA 1379.docx' '_SeguimentOutputName Req1 -> Req2, data avui'
AssertEq (_SeguimentOutputName '2026-06-05_Req2_GIA 1379' $d) '2026-06-05_Req3_GIA 1379.docx' '_SeguimentOutputName Req2 -> Req3'
AssertEq (_SeguimentOutputName '2026-05-15_Req_SELECTIUM CORNELLA' $d) '2026-06-05_Req2_SELECTIUM CORNELLA.docx' '_SeguimentOutputName Req sense numero -> Req2 (requeriment antic)'
AssertEq (_SeguimentOutputName 'Informe antic fet a ma' $d) '2026-06-05_Seguiment_Informe antic fet a ma.docx' '_SeguimentOutputName sense Req -> prefix Seguiment'
AssertEq (_SeguimentOutputName 'a/b:c' $d) '2026-06-05_Seguiment_a_b_c.docx' '_SeguimentOutputName saneja caracters il-legals'

Write-Host "`n--- Seguiment: _BuildSeguimentModel ---"
$recs = @(
    [pscustomobject]@{ Index=1; Text='Capcalera intro';                 ListString=''; Bold=0 }
    [pscustomobject]@{ Index=2; Text='1. Baixa tensio. Cal legalitzar.'; ListString=''; Bold=9999999 }
    [pscustomobject]@{ Index=3; Text='01/06/2026: No s entrega.';        ListString=''; Bold=-1 }
    [pscustomobject]@{ Index=4; Text='03/06/2026: Falten dades.';        ListString=''; Bold=-1 }
    [pscustomobject]@{ Index=5; Text='2. Alta tensio. Cal projecte.';    ListString=''; Bold=0 }
    [pscustomobject]@{ Index=6; Text='05/06/2026: S aporta.';            ListString=''; Bold=0 }
    [pscustomobject]@{ Index=7; Text="Vist l'anterior, cal requerir.";   ListString=''; Bold=0 }
)
$model = _BuildSeguimentModel $recs
AssertEq $model.Requirements.Count 2          '_BuildSeguimentModel detecta 2 requeriments'
AssertEq $model.LastReqParaIndex 5            '_BuildSeguimentModel ultim requeriment a l index 5'
AssertEq $model.Requirements[0].ParaIndex 2   '_BuildSeguimentModel req1 a l index 2'
AssertEq $model.Requirements[0].Annotations.Count 2 '_BuildSeguimentModel req1 amb 2 anotacions'
AssertEq $model.Requirements[0].Annotations[1].ParaIndex 4 '_BuildSeguimentModel anotacio 2 del req1 a l index 4'
Assert (-not $model.Requirements[0].WasResolved) '_BuildSeguimentModel req1 (bold mixt, amb anotacions) -> pendent'
Assert ($model.Requirements[1].WasResolved)      '_BuildSeguimentModel req2 (bold 0, amb anotacio) -> resolt'
$startC = _FindConclusionStartIndex (@($recs | ForEach-Object { $_.Text })) $model.LastReqParaIndex $SeguimentConclusionPhrases
AssertEq $startC 7                            '_BuildSeguimentModel + deteccio: conclusions a l index 7'

# Requeriment FRESC (sense anotacions) -> pendent encara que bold=0
$recsFresh = @(
    [pscustomobject]@{ Index=1; Text='1. Cal aportar X.'; ListString=''; Bold=0 }
)
$mFresh = _BuildSeguimentModel $recsFresh
Assert (-not $mFresh.Requirements[0].WasResolved) '_BuildSeguimentModel requeriment fresc (sense anotacions) -> pendent'

Write-Host "`n--- Seguiment XML: helpers de lectura (text, negreta, body) ---"
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function New-XmlInfoFromString($xmlText, $numFmt=@{}, $styleNum=@{}) {
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($xmlText)
    $ns = _NewWordNsMgr $xml
    return [pscustomobject]@{ Path=''; Xml=$xml; Ns=$ns; Body=$xml.SelectSingleNode('//w:body',$ns); NumFmt=$numFmt; StyleNum=$styleNum }
}
# Validacio ESTRICTA (com fa el Word, no com el XmlDocument tolerant): parseig
# complet amb XmlReader + cap prefix il-legal per al namespace XML reservat.
function Test-StrictXml([string]$xmlText) {
    if ([regex]::Matches($xmlText, 'xmlns:\w+="http://www\.w3\.org/XML/1998/namespace"').Count -gt 0) { return $false }
    try {
        $rd = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xmlText)))
        while ($rd.Read()) { }
        $rd.Close()
        return $true
    } catch { return $false }
}
$docStr = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W"><w:body>
<w:p><w:r><w:t>Capcalera</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">1.</w:t></w:r><w:r><w:tab/><w:t>Baixa tensio.</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">2. </w:t></w:r><w:r><w:t>Alta tensio.</w:t></w:r></w:p>
<w:p><w:r><w:t>Vist l anterior, cal requerir.</w:t></w:r></w:p>
<w:p><w:r><w:t>Ho poso al seu coneixement.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body></w:document>
"@
$xi = New-XmlInfoFromString $docStr
$bp = @(_BodyParagraphsXml $xi)
AssertEq $bp.Count 5                            'Seguiment XML: 5 paragrafs directes al body (taules/sectPr fora)'
AssertEq (_ParagraphTextXml $bp[1] $xi.Ns) "1.`tBaixa tensio." 'Seguiment XML: _ParagraphTextXml inclou el tab entre numero i text'
AssertEq (_ParagraphBoldStateXml $bp[0] $xi.Ns) 0      'Seguiment XML: paragraf sense negreta -> 0'
AssertEq (_ParagraphBoldStateXml $bp[1] $xi.Ns) 9999999 'Seguiment XML: numero negreta + text no -> mixt (9999999)'

Write-Host "`n--- Seguiment XML: model + deteccio de conclusions ---"
$records = _CollectParaRecordsXml $xi $bp
$model   = _BuildSeguimentModel $records
AssertEq $model.Requirements.Count 2     'Seguiment XML: 2 requeriments detectats (amb tab i amb espai)'
AssertEq $model.LastReqParaIndex 3       'Seguiment XML: ultim requeriment a l index 3'
$paraTexts = @($records | ForEach-Object { $_.Text })
AssertEq (_FindConclusionStartIndex $paraTexts $model.LastReqParaIndex @("Vist l anterior","Ho poso al seu coneixement")) 4 'Seguiment XML: conclusions detectades a l index 4'

Write-Host "`n--- Seguiment XML: _SetParagraphBoldXml on/off ---"
_SetParagraphBoldXml $xi $bp[0] $true
Assert ($null -ne $bp[0].SelectSingleNode('w:r/w:rPr/w:b', $xi.Ns)) '_SetParagraphBoldXml on: afegeix <w:b>'
AssertEq (_ParagraphBoldStateXml $bp[0] $xi.Ns) (-1) '_SetParagraphBoldXml on: tot negreta (-1)'
_SetParagraphBoldXml $xi $bp[0] $false
$bOff = $bp[0].SelectSingleNode('w:r/w:rPr/w:b', $xi.Ns)
AssertEq ($bOff.GetAttribute('val',$W)) 'false' '_SetParagraphBoldXml off: <w:b w:val="false">'
AssertEq (_ParagraphBoldStateXml $bp[0] $xi.Ns) 0 '_SetParagraphBoldXml off: sense negreta (0)'

Write-Host "`n--- Seguiment XML: transformacio completa (blocs, seccions i subseccions) ---"
# Doc realista: req1 + sub-linia + enllac, ESPAIADOR, SECCIO (negreta), ESPAIADOR,
# SUBSECCIO (subratllat), ESPAIADOR, req2, i bloc de conclusions.
$docStr2 = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W"><w:body>
<w:p><w:r><w:t>Capcalera</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">1. </w:t></w:r><w:r><w:t>Baixa tensio.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>- Subdocument A</w:t></w:r></w:p>
<w:p><w:r><w:t>https://exemple.cat/x</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Instal·lacions</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Inspeccions inicials</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">2. </w:t></w:r><w:r><w:t>Alta tensio.</w:t></w:r></w:p>
<w:p><w:r><w:t>Vist l anterior, cal requerir.</w:t></w:r></w:p>
<w:p><w:r><w:t>Ho poso al seu coneixement.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body></w:document>
"@
$xi2 = New-XmlInfoFromString $docStr2
$bp2 = @(_BodyParagraphsXml $xi2)
$model2 = _BuildSeguimentModel (_CollectParaRecordsXml $xi2 $bp2)
AssertEq $model2.Requirements.Count 2 'Transform: 2 requeriments detectats'
Assert (_IsUrlParagraphXml $bp2[3] $xi2.Ns)        '_IsUrlParagraphXml: el paragraf de l enllac -> true'
Assert (-not (_IsUrlParagraphXml $bp2[1] $xi2.Ns)) '_IsUrlParagraphXml: el requeriment -> false'
Assert (_IsSubsectionXml $bp2[7] $xi2.Ns)          '_IsSubsectionXml: subseccio subratllada -> true'
Assert (-not (_IsSubsectionXml $bp2[5] $xi2.Ns))   '_IsSubsectionXml: seccio (negreta) -> false'

$dec = @(
    [pscustomobject]@{ Resolved=$false; NewComment="No s'aporta." },
    [pscustomobject]@{ Resolved=$true;  NewComment="S'aporta." }
)
$flds = [ordered]@{}; $flds['lloc'] = [pscustomobject]@{ Type='text'; Value='Cornella'; Hint=$null; Options=$null }
$concl = @([pscustomobject]@{ Title='C'; Body='Cos **negreta** aqui.' })
$always = @('Ho poso, [CAMP: lloc].')
_ApplySeguimentTransform -xmlInfo $xi2 -bodyParas $bp2 -model $model2 -conclusionStartIndex 11 `
    -decisions $dec -dateStr '04/06/2026' -conclHeaderText 'CONCLUSIONS' `
    -selectedConclusions $concl -alwaysConclusions $always -fields $flds
$afterN = @(_BodyParagraphsXml $xi2)
$after = @($afterN | ForEach-Object { [pscustomobject]@{ T=(_ParagraphTextXml $_ $xi2.Ns); B=(_ParagraphBoldStateXml $_ $xi2.Ns) } })
$texts = @($after | ForEach-Object { $_.T })

Assert ($texts -contains "04/06/2026: No s'aporta.")          'Transform: anotacio del req1 inserida'
Assert ($texts -contains "04/06/2026: S'aporta.")             'Transform: anotacio del req2 inserida'
Assert (-not ($texts -contains 'Vist l anterior, cal requerir.')) 'Transform: conclusio antiga esborrada'
Assert ($texts -contains 'CONCLUSIONS')                       'Transform: titol de conclusions afegit'
Assert ($texts -contains 'Ho poso, Cornella.')               'Transform: [CAMP: lloc] resolt a "Cornella"'

# PUNT 1: l'anotacio del req1 va despres de l enllac i ABANS de la seccio
$idxUrl    = [array]::IndexOf($texts, 'https://exemple.cat/x')
$idxAnnot1 = [array]::IndexOf($texts, "04/06/2026: No s'aporta.")
$idxSecc   = [array]::IndexOf($texts, 'Instal·lacions')
$idxSubs   = [array]::IndexOf($texts, 'Inspeccions inicials')
Assert ($idxAnnot1 -gt $idxUrl)  'Transform (punt 1): l anotacio va despres de l enllac'
Assert ($idxAnnot1 -lt $idxSecc) 'Transform (punt 1): l anotacio va ABANS de la seccio (no la traspassa)'

# La SECCIO i la SUBSECCIO NO s han de tocar (segueixen negreta/subratllat)
AssertEq $after[$idxSecc].B (-1)                   'Transform: la seccio segueix en negreta (no es modifica)'
Assert (_IsSubsectionXml $afterN[$idxSubs] $xi2.Ns) 'Transform: la subseccio segueix subratllada (no es modifica)'

# NOMES el comentari de l'anotacio va en negreta (si pendent). NI requeriment,
# NI sub-linia, NI enllac, NI la data.
$idxSub  = [array]::IndexOf($texts, '- Subdocument A')
AssertEq $after[$idxSub].B 0 'Transform: la sub-linia NO es posa en negreta'
AssertEq $after[$idxUrl].B 0 'Transform: l enllac MAI en negreta'
# Anotacio PENDENT (req1): data normal + comentari negreta -> paragraf mixt
AssertEq $after[$idxAnnot1].B 9999999 'Transform: anotacio pendent -> nomes el comentari en negreta (mixt)'
$annRuns = @($afterN[$idxAnnot1].SelectNodes('w:r', $xi2.Ns))
AssertEq $annRuns.Count 2 'Transform: anotacio amb 2 runs (data + comentari)'
Assert ($null -eq $annRuns[0].SelectSingleNode('w:rPr/w:b', $xi2.Ns)) 'Transform: el run de la DATA NO porta negreta'
Assert ($null -ne $annRuns[1].SelectSingleNode('w:rPr/w:b', $xi2.Ns)) 'Transform: el run del COMENTARI porta negreta'

# req2 resolt -> l'anotacio NO porta cap negreta
$idxAnnot2 = [array]::IndexOf($texts, "04/06/2026: S'aporta.")
AssertEq $after[$idxAnnot2].B 0  'Transform: anotacio resolta -> cap negreta'

# PUNT 2: l anotacio NO s ha d enumerar -> numPr amb numId=0
$annotNode = $afterN[$idxAnnot1]
$nid = $annotNode.SelectSingleNode('w:pPr/w:numPr/w:numId', $xi2.Ns)
Assert ($null -ne $nid -and $nid.GetAttribute('val',$W) -eq '0') 'Transform (punt 2): l anotacio te numId=0 (sense numeracio)'

# Espai a sobre de la primera anotacio (separacio amb el cos)
$sp = $annotNode.SelectSingleNode('w:pPr/w:spacing', $xi2.Ns)
Assert ($null -ne $sp -and [int]$sp.GetAttribute('before',$W) -gt 0) 'Transform: la primera anotacio te espai a sobre'

# PUNT 6: font Bookman Old Style a l anotacio i a la conclusio
$annFont = $annotNode.SelectSingleNode('w:r/w:rPr/w:rFonts', $xi2.Ns)
Assert ($null -ne $annFont -and $annFont.GetAttribute('ascii',$W) -eq 'Bookman Old Style') 'Transform (punt 6): anotacio en Bookman Old Style'
$idxConcl = [array]::IndexOf($texts, 'Cos negreta aqui.')
Assert ($idxConcl -ge 0) 'Transform: conclusio nova present (markdown processat)'
$concFont = $afterN[$idxConcl].SelectSingleNode('w:r/w:rPr/w:rFonts', $xi2.Ns)
Assert ($null -ne $concFont -and $concFont.GetAttribute('ascii',$W) -eq 'Bookman Old Style') 'Transform (punt 6): conclusio en Bookman Old Style'
$concJc = $afterN[$idxConcl].SelectSingleNode('w:pPr/w:jc', $xi2.Ns)
Assert ($null -ne $concJc -and $concJc.GetAttribute('val',$W) -eq 'both') 'Transform (punt 6): conclusio justificada'

# sectPr preservat + XML estrictament valid (com l obre el Word)
Assert ($null -ne $xi2.Body.SelectSingleNode('w:sectPr', $xi2.Ns)) 'Transform: <w:sectPr> preservat'
$swT = New-Object System.IO.StringWriter; $xi2.Xml.Save($swT)
Assert (Test-StrictXml $swT.ToString()) 'Transform: el document.xml resultant es estrictament valid (Word)'

Write-Host "`n--- Seguiment XML: format de l'anotacio d'un SUB-PUNT ---"
# Requeriment numerat amb 2 sub-punts en una llista REAL del Word (numId=33, la
# sagnia la posa la numeracio i NO hi ha cap <w:ind>). En forcar numId=0 perque
# l'anotacio no s'enumeri es perdia aquella sagnia i l'anotacio quedava
# desalineada; a mes el sub-punt seguent li quedava enganxat.
$docChild = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W"><w:body>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">1. </w:t></w:r><w:r><w:t>Elements de coccio.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="33"/></w:numPr></w:pPr><w:r><w:t>Retirar la fregidora.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="34"/></w:numPr><w:ind w:left="1134"/></w:pPr><w:r><w:t>Presentar certificat.</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:t>Vist l anterior, cal requerir.</w:t></w:r></w:p>
<w:sectPr/>
</w:body></w:document>
"@
$xiC = New-XmlInfoFromString $docChild @{ '33' = 'bullet'; '34' = 'bullet' }
$bpC = @(_BodyParagraphsXml $xiC)
$modelC = _BuildSeguimentModel (_CollectParaRecordsXml $xiC $bpC)
$uC = @($modelC.Requirements)
AssertEq $uC.Count 2                  'Sub-punt: el req amb 2 sub-punts dona 2 unitats'
AssertEq ([bool]$uC[0].IsChild) $true 'Sub-punt: la 1a unitat es un fill'
_ApplySeguimentTransform -xmlInfo $xiC -bodyParas $bpC -model $modelC -conclusionStartIndex 5 `
    -decisions @(
        [pscustomobject]@{ Resolved=$true; NewComment="S'aporta." },
        [pscustomobject]@{ Resolved=$true; NewComment="S'aporta certificat." }
    ) -dateStr '27/07/2026' -conclHeaderText '' -selectedConclusions @() -alwaysConclusions @() -fields ([ordered]@{})
$afterC = @(_BodyParagraphsXml $xiC)
$textsC = @($afterC | ForEach-Object { _ParagraphTextXml $_ $xiC.Ns })
$iA1 = [array]::IndexOf($textsC, "27/07/2026: S'aporta.")
$iA2 = [array]::IndexOf($textsC, "27/07/2026: S'aporta certificat.")
Assert ($iA1 -ge 0 -and $iA2 -ge 0) 'Sub-punt: les dues anotacions inserides'
# 1r sub-punt: sense <w:ind> propi -> cal posar la sagnia EXPLICITA (1,25 cm)
$indC1 = $afterC[$iA1].SelectSingleNode('w:pPr/w:ind', $xiC.Ns)
Assert ($null -ne $indC1) "Sub-punt: l'anotacio d'un sub-punt sense sagnia propia en rep una d'explicita"
AssertEq ([int]$indC1.GetAttribute('left',$W)) (_CmToTwips $Script:ReportFormatConfig.AnnotationIndentCm) 'Sub-punt: la sagnia es AnnotationIndentCm (1,25 cm = 709 twips)'
$spC1 = $afterC[$iA1].SelectSingleNode('w:pPr/w:spacing', $xiC.Ns)
AssertEq ([int]$spC1.GetAttribute('after',$W)) (_PtToTwips $Script:ReportFormatConfig.AnnotationSpaceAfterPt) 'Sub-punt: espai a sota = AnnotationSpaceAfterPt (12 pt = 240)'
AssertEq ([int]$spC1.GetAttribute('before',$W)) (_PtToTwips $Script:ReportFormatConfig.AnnotationSpaceBeforePt) 'Sub-punt: espai a sobre = AnnotationSpaceBeforePt (10 pt = 200)'
# 2n sub-punt: JA porta <w:ind w:left="1134"> -> es respecta (no es trepitja)
$indC2 = $afterC[$iA2].SelectSingleNode('w:pPr/w:ind', $xiC.Ns)
AssertEq ([int]$indC2.GetAttribute('left',$W)) 1134 "Sub-punt: si el sub-punt ja porta sagnia propia, es conserva"
# Ordre OOXML dins de <w:pPr>: numPr, spacing, ind (si no, el Word es queixa)
$namesC = @($afterC[$iA1].SelectSingleNode('w:pPr', $xiC.Ns).ChildNodes | ForEach-Object { $_.LocalName })
AssertEq ($namesC -join ',') 'pStyle,numPr,spacing,ind' 'Sub-punt: ordre correcte dels elements dins de <w:pPr>'
$swC = New-Object System.IO.StringWriter; $xiC.Xml.Save($swC)
Assert (Test-StrictXml $swC.ToString()) 'Sub-punt: document.xml estrictament valid'

Write-Host "`n--- Seguiment XML: estat resolt/pendent guardat a la negreta del comentari ---"
# req1 ja RESOLT (anotacio previa amb comentari SENSE negreta).
# req2 PENDENT (anotacio previa amb comentari EN negreta).
$rb = '<w:rPr><w:rFonts w:ascii="Bookman Old Style"/></w:rPr>'
$docState = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W"><w:body>
<w:p><w:r><w:t>1. Cal aportar X.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r>$rb<w:t xml:space="preserve">10/06/2026: </w:t></w:r><w:r>$rb<w:t>S'aporta.</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:t>2. Cal aportar Y.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r>$rb<w:t xml:space="preserve">10/06/2026: </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Bookman Old Style"/><w:b/><w:bCs/></w:rPr><w:t>No s'aporta.</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:t>Vist l anterior, cal requerir.</w:t></w:r></w:p>
<w:sectPr/>
</w:body></w:document>
"@
$xiS = New-XmlInfoFromString $docState
$bpS = @(_BodyParagraphsXml $xiS)
$modelS = _BuildSeguimentModel (_CollectParaRecordsXml $xiS $bpS)
AssertEq $modelS.Requirements.Count 2 'Estat: 2 requeriments'
Assert ($modelS.Requirements[0].WasResolved)        'Estat: req1 amb comentari no-negreta -> RESOLT'
Assert (-not $modelS.Requirements[1].WasResolved)   'Estat: req2 amb comentari en negreta -> PENDENT'

# Transform: tots dos es marquen Resolt. req1 ja ho estava -> NO s afegeix linia.
# req2 passa de pendent a resolt -> s afegeix "S'aporta." i es des-negreta l antic.
$decS = @(
    [pscustomobject]@{ Resolved=$true; NewComment="S'aporta." },
    [pscustomobject]@{ Resolved=$true; NewComment="S'aporta." }
)
_ApplySeguimentTransform -xmlInfo $xiS -bodyParas $bpS -model $modelS -conclusionStartIndex 7 `
    -decisions $decS -dateStr '20/06/2026' -conclHeaderText '' -selectedConclusions @() -alwaysConclusions @() -fields ([ordered]@{})
$afterS = @(_BodyParagraphsXml $xiS | ForEach-Object { [pscustomobject]@{ T=(_ParagraphTextXml $_ $xiS.Ns); B=(_ParagraphBoldStateXml $_ $xiS.Ns) } })
$textsS = @($afterS | ForEach-Object { $_.T })
$nAnnot = @($textsS | Where-Object { $_ -match "^\s*\d{1,2}/\d{1,2}/\d{4}\s*:" }).Count
AssertEq $nAnnot 3 'D: req1 (ja resolt) NO afegeix linia; req2 si -> 2+1 = 3 anotacions'
$nNew20 = @($textsS | Where-Object { $_ -eq "20/06/2026: S'aporta." }).Count
AssertEq $nNew20 1 'D: nomes req2 afegeix una linia nova del 20/06'
# L antic "No s'aporta." de req2 ara queda SENSE negreta (historic)
$idxOld = [array]::IndexOf($textsS, "10/06/2026: No s'aporta.")
AssertEq $afterS[$idxOld].B 0 'C: el comentari pendent anterior es des-negreta en resoldre-s'

Write-Host "`n--- Seguiment XML: numeracio AUTOMATICA del Word (numId via estil) ---"
# Requeriments sense numero al text: la numeracio ve de l'estil (numId=5 decimal),
# com als informes fets amb numeracio automatica del Word.
$numFmt   = @{ '5' = 'decimal'; '3' = 'bullet' }
$styleNum = @{ 'Prrafodelista' = @{ NumId='5'; Ilvl=0 } }
$docAuto = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W"><w:body>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/></w:pPr><w:r><w:t>Vector Aigua. Cal aportar X.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>https://exemple.cat/aigua</w:t></w:r></w:p>
<w:p/>
<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/></w:pPr><w:r><w:t>Vector Residus. Cal aportar Y.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Un bullet, no requeriment.</w:t></w:r></w:p>
<w:sectPr/>
</w:body></w:document>
"@
$xiA = New-XmlInfoFromString $docAuto $numFmt $styleNum
$bpA = @(_BodyParagraphsXml $xiA)
$recA = @(_CollectParaRecordsXml $xiA $bpA)
AssertEq $recA[0].ListString '1.' 'Auto-num: requeriment 1 detectat via estil (ListString=1.)'
AssertEq $recA[1].ListString ''   'Auto-num: URL amb numId=0 -> no numerat'
AssertEq $recA[3].ListString '2.' 'Auto-num: requeriment 2 detectat (ListString=2.)'
AssertEq $recA[4].ListString ''   'Auto-num: bullet (numId=3) -> NO es requeriment'
$modelA = _BuildSeguimentModel $recA
AssertEq $modelA.Requirements.Count 2 'Auto-num: 2 requeriments detectats per numeracio d estil'
AssertEq $modelA.Requirements[0].Text 'Vector Aigua. Cal aportar X.' 'Auto-num: text del requeriment 1'
Assert (-not $modelA.Requirements[0].WasResolved) 'Auto-num: requeriment fresc (sense anotacions) -> pendent per defecte'

Write-Host "`n--- _StripMarkers (treu ** i // per a la previsualitzacio) ---"
AssertEq (_StripMarkers 'text **negreta** i //cursiva//') 'text negreta i cursiva' '_StripMarkers: parells nets'
AssertEq (_StripMarkers '**ampliar el termini ')          'ampliar el termini '   '_StripMarkers: marcador partit (orfe) tambe es treu'
AssertEq (_StripMarkers $null)                            ''                      '_StripMarkers: null -> buit'

Write-Host "`n--- _SegmentRichText (camps inline al Pas 3 i conclusions) ---"
$segA = @(_SegmentRichText 'es valora **ampliar el termini [OPCIO: Mesos ampliacio | un mes | dos mesos]**.')
AssertEq $segA.Count 3                       '_SegmentRichText: 3 trossos (text, opcio, text)'
AssertEq $segA[0].Kind 'text'                '_SegmentRichText: 1r tros es text'
AssertEq $segA[0].Text 'es valora ampliar el termini ' '_SegmentRichText: text sense marcadors **'
AssertEq $segA[1].Kind 'opcio'               '_SegmentRichText: 2n tros es opcio'
AssertEq $segA[1].Name 'Mesos ampliacio'     '_SegmentRichText: nom de l opcio'
AssertEq $segA[1].Options.Count 2            '_SegmentRichText: 2 opcions'
AssertEq $segA[1].Options[0] 'un mes'        '_SegmentRichText: 1a opcio'
AssertEq $segA[2].Kind 'text'                '_SegmentRichText: 3r tros es text (.)'

$segB = @(_SegmentRichText 'precintar [CAMP: que es precinta] fins esmenar')
AssertEq $segB.Count 3                       '_SegmentRichText: camp lliure -> 3 trossos'
AssertEq $segB[1].Kind 'camp'                '_SegmentRichText: tros central es camp'
AssertEq $segB[1].Name 'que es precinta'     '_SegmentRichText: nom del camp'

$segH = @(_SegmentRichText '[CAMP: data (dd/mm/aaaa)]')
AssertEq $segH.Count 1                       '_SegmentRichText: nomes el camp'
AssertEq $segH[0].Name 'data'                '_SegmentRichText: nom sense el hint'
AssertEq $segH[0].Hint 'dd/mm/aaaa'          '_SegmentRichText: hint separat'

AssertEq (@(_SegmentRichText 'sense cap camp').Count) 1 '_SegmentRichText: text pla -> 1 tros'
AssertEq (@(_SegmentRichText '').Count)               0 '_SegmentRichText: buit -> cap tros'

Write-Host "`n--- Format-Section: MAJUSCULES sense negreta ---"
# Re-carreguem Format.ps1 perque tests anteriors han estubat Format-Section.
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'Format.ps1')
# Estubem $sel amb totes les propietats que _Reset-Char i _Apply-Indent toquen.
$global:secOps = New-Object System.Collections.ArrayList
$selSec = New-Object PSObject -Property @{
    Font = New-Object PSObject -Property @{ Bold = 0; Italic = 0; Underline = 0; Size = 11 }
    ParagraphFormat = New-Object PSObject -Property @{
        LeftIndent = 0; FirstLineIndent = 0; Alignment = 3; SpaceAfter = 0
    }
}
Add-Member -InputObject $selSec -MemberType ScriptMethod -Name TypeParagraph -Value {}
Add-Member -InputObject $selSec -MemberType ScriptMethod -Name TypeText -Value {
    param($t)
    $b = if ($this.Font.Bold -eq 1) { 'B' } else { '-' }
    [void]$global:secOps.Add(("{0}:{1}" -f $b, $t))
}
Format-Section $selSec 'Instal·lacions'
AssertEq $global:secOps.Count 1                          'Format-Section: 1 sola crida a TypeText'
AssertEq $global:secOps[0]    '-:INSTAL·LACIONS'         'Format-Section: text en MAJUSCULES, SENSE negreta'

Write-Host "`n--- _GetItemTooltip (text per al tooltip del Pas 3 sense enllacos) ---"
$el1 = [pscustomobject]@{ BodyLines = @(
    'Instal·lacio de baixa tensio. Veure document.',
    '[[URL]] https://canalempresa.gencat.cat/...'
)}
AssertEq (_GetItemTooltip $el1) 'Instal·lacio de baixa tensio. Veure document.' '_GetItemTooltip: descarta linies [[URL]]'

$el2 = [pscustomobject]@{ BodyLines = @(
    'Text principal.',
    'https://example.com/url',
    'Comentari extra.'
)}
$tip2 = _GetItemTooltip $el2
Assert ($tip2.Contains('Text principal.'))   '_GetItemTooltip: conserva el text principal'
Assert ($tip2.Contains('Comentari extra.'))  '_GetItemTooltip: conserva text extra (no-URL)'
Assert (-not $tip2.Contains('http'))         '_GetItemTooltip: descarta linies URL-only'

$el3 = [pscustomobject]@{ BodyLines = @('Mirar https://x.cat/y al final.') }
$tip3 = _GetItemTooltip $el3
AssertEq $tip3 'Mirar'                       '_GetItemTooltip: linies mixtes -> nomes la part de text'

$el4 = [pscustomobject]@{ BodyLines = @() }
AssertEq (_GetItemTooltip $el4) ''           '_GetItemTooltip: cap linia -> buit'

AssertEq (_GetItemTooltip $null) ''          '_GetItemTooltip: null -> buit'

Write-Host "`n--- Mode paquet: Build-SelectionFromKeys (reconstruir Pas 3 sense UI) ---"
# Cataleg de prova amb 2 seccions, items, un fill i una subseccio.
$secA = [pscustomobject]@{ Title='Instal·lacions'; Items=(New-Object System.Collections.ArrayList) }
[void]$secA.Items.Add([pscustomobject]@{ Kind='subsection'; Short='Legalitzacions'; BodyLines=@(); Children=@() })
$itBT = [pscustomobject]@{ Kind='item'; Short='Baixa tensio'; BodyLines=@('Cal legalitzar.'); Children=(New-Object System.Collections.ArrayList) }
[void]$itBT.Children.Add([pscustomobject]@{ Kind='child'; Short='Memoria'; BodyLines=@('Memoria tecnica.'); Children=@() })
[void]$secA.Items.Add($itBT)
[void]$secA.Items.Add([pscustomobject]@{ Kind='item'; Short='Gas'; BodyLines=@('Revisar gas.'); Children=(New-Object System.Collections.ArrayList) })
$secB = [pscustomobject]@{ Title='Sanitat'; Items=(New-Object System.Collections.ArrayList) }
[void]$secB.Items.Add([pscustomobject]@{ Kind='item'; Short='Aigua'; BodyLines=@('Potabilitat.'); Children=(New-Object System.Collections.ArrayList) })
$sectionsTest = @($secA, $secB)

# Selecciono nomes "Baixa tensio" i "Aigua" (no "Gas").
$sel = Build-SelectionFromKeys $sectionsTest @('Instal·lacions::Baixa tensio','Sanitat::Aigua')
AssertEq $sel.Count 2                                    'Build-SelectionFromKeys: 2 seccions amb items reals'
AssertEq $sel[0].Title 'Instal·lacions'                 'Build-SelectionFromKeys: 1a seccio'
# La subseccio es conserva sempre; despres l'item seleccionat.
$itemsA = @($sel[0].Items | Where-Object { $_.Kind -eq 'item' })
AssertEq $itemsA.Count 1                                 'Build-SelectionFromKeys: nomes 1 item triat a la seccio A'
AssertEq $itemsA[0].Short 'Baixa tensio'                'Build-SelectionFromKeys: item correcte'
$hasSub = @($sel[0].Items | Where-Object { $_.Kind -eq 'subsection' }).Count
AssertEq $hasSub 1                                       'Build-SelectionFromKeys: la subseccio es conserva'

# Round-trip: les claus del resultat coincideixen amb les demanades (mes el fill, que no vam triar -> no hi es).
$keys = Get-SelectedKeysFromResult $sel
Assert ($keys -contains 'Instal·lacions::Baixa tensio') 'Build-SelectionFromKeys: round-trip conté l''item'
Assert ($keys -contains 'Sanitat::Aigua')               'Build-SelectionFromKeys: round-trip conté l''altre item'
Assert (-not ($keys -contains 'Instal·lacions::Gas'))   'Build-SelectionFromKeys: NO conté l''item no triat'

# Si nomes triem el fill, l'item pare s'inclou (Selected=$false) per portar-lo.
$selChild = Build-SelectionFromKeys $sectionsTest @('Instal·lacions::Baixa tensio::Memoria')
$itemsChild = @($selChild[0].Items | Where-Object { $_.Kind -eq 'item' })
AssertEq $itemsChild.Count 1                             'Build-SelectionFromKeys: nomes fill -> pare inclos'
AssertEq ([bool]$itemsChild[0].Selected) $false         'Build-SelectionFromKeys: pare amb Selected=false'
AssertEq $itemsChild[0].Children.Count 1                'Build-SelectionFromKeys: el fill triat hi es'

# Cap clau -> cap seccio.
$selNone = Build-SelectionFromKeys $sectionsTest @()
AssertEq $selNone.Count 0                                'Build-SelectionFromKeys: cap clau -> resultat buit'

Write-Host "`n--- Mode paquet: Build-ConclusionsFromTitles ---"
$selectable = @(
    [pscustomobject]@{ Title='Terrassa projecte'; Body='La terrassa...' },
    [pscustomobject]@{ Title='Requeriment';        Body='Vist l''anterior...' },
    [pscustomobject]@{ Title='Arxiu';              Body='S''arxiva...' }
)
$concl = Build-ConclusionsFromTitles $selectable @('Requeriment','Terrassa projecte')
AssertEq $concl.Count 2                                  'Build-ConclusionsFromTitles: 2 triades'
# Es preserva l'ordre del FITXER (selectable), no l'ordre dels titols demanats.
AssertEq $concl[0].Title 'Terrassa projecte'            'Build-ConclusionsFromTitles: ordre del fitxer (1)'
AssertEq $concl[1].Title 'Requeriment'                  'Build-ConclusionsFromTitles: ordre del fitxer (2)'
AssertEq $concl[0].Body 'La terrassa...'                'Build-ConclusionsFromTitles: conserva el Body'
AssertEq (Build-ConclusionsFromTitles $selectable @()).Count 0 'Build-ConclusionsFromTitles: cap titol -> buit'

Write-Host "`n--- Mode paquet: Build-FieldsFromPaquet (camps + valors) ---"
# Item amb un [CAMP:] i un [OPCIO:]; conclusio amb un altre [CAMP:].
$secF = [pscustomobject]@{ Title='S'; Items=(New-Object System.Collections.ArrayList) }
[void]$secF.Items.Add([pscustomobject]@{ Kind='item'; Short='X';
    BodyLines=@('Cal aportar el [CAMP: certificat] [OPCIO: Termini | 1 mes | 3 mesos].'); Children=@() })
$selF = @($secF)
$conclF = @([pscustomobject]@{ Title='C'; Body='Signat per [CAMP: tecnic].' })
# Valors com a hashtable (cas Vigilant) i deixant un camp sense valor (->default).
$fv = @{ 'certificat'='ISO-9001'; 'tecnic'='Joan' }
$fields = Build-FieldsFromPaquet $selF $conclF @() $fv
Assert ($fields.Contains('certificat'))                 'Build-FieldsFromPaquet: detecta CAMP del REQ'
Assert ($fields.Contains('Termini'))                    'Build-FieldsFromPaquet: detecta OPCIO del REQ'
Assert ($fields.Contains('tecnic'))                     'Build-FieldsFromPaquet: detecta CAMP de la conclusio'
AssertEq $fields['certificat'].Value 'ISO-9001'         'Build-FieldsFromPaquet: aplica valor de text'
AssertEq $fields['tecnic'].Value 'Joan'                 'Build-FieldsFromPaquet: aplica valor de conclusio'
AssertEq $fields['Termini'].Value '1 mes'               'Build-FieldsFromPaquet: OPCIO sense valor -> 1a opcio'

# Tambe ha d'acceptar un PSCustomObject (cas ConvertFrom-Json del paquet real).
$fvObj = [pscustomobject]@{ certificat='X'; tecnic='Y'; Termini='3 mesos' }
$fields2 = Build-FieldsFromPaquet $selF $conclF @() $fvObj
AssertEq $fields2['certificat'].Value 'X'               'Build-FieldsFromPaquet: valors des de PSCustomObject'
AssertEq $fields2['Termini'].Value '3 mesos'            'Build-FieldsFromPaquet: OPCIO amb valor del paquet'

Write-Host "`n--- Informes.ps1: _ParseDataInformeFromName ---"
AssertEq (_ParseDataInformeFromName '20260710_Req4_KRICHI.docx') '2026-07-10' 'data AAAAMMDD sense separador'
AssertEq (_ParseDataInformeFromName '2026-07-10_Req.docx')       '2026-07-10' 'data AAAA-MM-DD'
AssertEq (_ParseDataInformeFromName '20250114_Req_X.docx')       '2025-01-14' 'data AAAAMMDD (gener)'
AssertEq (_ParseDataInformeFromName '2026.05.21 informe.docx')   '2026-05-21' 'data AAAA.MM.DD'
AssertEq (_ParseDataInformeFromName '26-07-10 antic.docx')       '2026-07-10' 'data AA-MM-DD (any 2 xifres)'
AssertEq (_ParseDataInformeFromName '20261332_x.docx')          ''           'data invalida (mes/dia) -> buit'
AssertEq (_ParseDataInformeFromName 'Informe final.docx')       ''           'sense data al principi -> buit'
AssertEq (_ParseDataInformeFromName '')                         ''           'nom buit -> buit'

Write-Host "`n--- Informes.ps1: _ExtractIdGia / _ExtractExpedient ---"
$linesA = @('Deficiencies', 'ID GIA:361', 'Exp. Num: 2025/1/2563', 'Sol licitat: KRICHI')
AssertEq (_ExtractIdGia $linesA)      '361'         '_ExtractIdGia enganxat "ID GIA:361"'
AssertEq (_ExtractIdGia @('ID GIA: 1379')) '1379'   '_ExtractIdGia amb espai'
AssertEq (_ExtractIdGia @('Sol licitat: X')) ''      '_ExtractIdGia sense GIA -> buit'
AssertEq (_ExtractExpedient $linesA)  '2025/1/2563' '_ExtractExpedient de "Exp. Num:"'
AssertEq (_ExtractExpedient @('res')) ''            '_ExtractExpedient sense exp -> buit'

Write-Host "`n--- Informes.ps1: _ExtractIdGia (placeholders 'encara sense GIA') ---"
AssertEq (_ExtractIdGia @("ID GIA:`t-"))   '' '_ExtractIdGia "-" -> placeholder, tractat com a buit'
AssertEq (_ExtractIdGia @('ID GIA: XXX'))  '' '_ExtractIdGia "XXX" -> placeholder, tractat com a buit'
AssertEq (_ExtractIdGia @('ID GIA: N/A'))  '' '_ExtractIdGia "N/A" -> placeholder, tractat com a buit'
AssertEq (_ExtractIdGia @('ID GIA: --'))   '' '_ExtractIdGia "--" -> placeholder, tractat com a buit'

Write-Host "`n--- Informes.ps1: _ExtractConclusio (apostrof tipografic) ---"
$ap = [char]0x2019   # apostrof tipografic (com als informes reals)
$linesC = @(
    "INFORME:",
    "Gas. L${ap}activitat...",
    "Vist l${ap}anterior s${ap}informa que NO es pot donar per finalitzat el procediment d${ap}esmena.",
    "Ho poso al seu coneixement als efectes oportuns,",
    "Cornella de Llobregat,"
)
$cC = _ExtractConclusio $linesC
AssertEq $cC.Text "Vist l${ap}anterior s${ap}informa que NO es pot donar per finalitzat el procediment d${ap}esmena." '_ExtractConclusio captura la frase "Vist l''anterior"'
AssertEq $cC.Font 'vist_anterior' '_ExtractConclusio "Vist l''anterior" -> Font vist_anterior (fiable)'
Assert ($cC.Text -notmatch 'Ho poso') '_ExtractConclusio exclou "Ho poso al seu coneixement"'
$cNone = _ExtractConclusio @('res', 'de res')
AssertEq $cNone.Text '' '_ExtractConclusio sense cap frase coneguda -> Text buit'
AssertEq $cNone.Font '' '_ExtractConclusio sense cap frase coneguda -> Font buit'

Write-Host "`n--- Informes.ps1: _ExtractConclusio (variants reals: risc, MNS, acte extraordinari) ---"
$linesRisc = @(
    'INFORME:',
    "Tenint en consideració el risc greu o imminent de seguretat, és pertinent precintar l${ap}activitat.",
    'Ho poso al seu coneixement als efectes oportuns,',
    'Cornella de Llobregat,'
)
$cRisc = _ExtractConclusio $linesRisc
AssertEq $cRisc.Font 'risc' '_ExtractConclusio "Tenint en consideració el risc" -> Font risc (fiable)'
Assert ($cRisc.Text -match 'precintar') '_ExtractConclusio "risc" captura el text de la decisio'

$linesMns = @(
    'INFORME:',
    "S${ap}informa FAVORABLEMENT de la Modificació NO substancial presentada.",
    'Ho poso al seu coneixement als efectes oportuns,',
    'Cornella de Llobregat,'
)
$cMns = _ExtractConclusio $linesMns
AssertEq $cMns.Font 'mns' '_ExtractConclusio "S''informa FAVORABLEMENT" -> Font mns (insensible a majuscules)'

$linesAct = @(
    'INFORME:',
    "El titular és responsable d${ap}executar i mantenir les mesures de seguretat.",
    'Ho poso al seu coneixement als efectes oportuns,',
    'Cornella de Llobregat,'
)
$cAct = _ExtractConclusio $linesAct
AssertEq $cAct.Font 'act_extr' '_ExtractConclusio "El titular és responsable d''executar" -> Font act_extr'

Write-Host "`n--- Informes.ps1: _ExtractConclusio (variants de tancament) ---"
$linesDoc = @(
    'INFORME:',
    "Vist l${ap}anterior, s${ap}informa favorablement.",
    "S${ap}informa als efectes oportuns,",
    'Cornella de Llobregat,'
)
$cDoc = _ExtractConclusio $linesDoc
AssertEq $cDoc.Text "Vist l${ap}anterior, s${ap}informa favorablement." '_ExtractConclusio tanca tambe amb "S''informa als efectes oportuns,"'

$linesSign = @(
    'INFORME:',
    "Vist l${ap}anterior, s${ap}informa favorablement.",
    "A Cornella de Llobregat, en la data i amb les signatures electròniques que figuren en aquest document."
)
$cSign = _ExtractConclusio $linesSign
AssertEq $cSign.Text "Vist l${ap}anterior, s${ap}informa favorablement." '_ExtractConclusio tanca tambe amb "A Cornella de Llobregat, en la data..." (signatura electronica)'

Write-Host "`n--- Informes.ps1: _ConclusioMotiu (motiu de revisio segons Font) ---"
AssertEq (_ConclusioMotiu ([pscustomobject]@{ Text=''; Font='' }))              'sense conclusio' '_ConclusioMotiu sense conclusio -> motiu'
AssertEq (_ConclusioMotiu ([pscustomobject]@{ Text='x'; Font='vist_anterior' })) ''                '_ConclusioMotiu vist_anterior -> sense motiu'
AssertEq (_ConclusioMotiu ([pscustomobject]@{ Text='x'; Font='risc' }))          ''                '_ConclusioMotiu risc -> sense motiu'
AssertEq (_ConclusioMotiu ([pscustomobject]@{ Text='x'; Font='mns' }))           ''                '_ConclusioMotiu mns -> sense motiu (es marca ignorat, no motiu)'
AssertEq (_ConclusioMotiu ([pscustomobject]@{ Text='x'; Font='act_extr' }))      ''                '_ConclusioMotiu act_extr -> sense motiu (es marca ignorat, no motiu)'

Write-Host "`n--- Informes.ps1: _ConclusioIgnorarPerDefecte (mns/act_extr -> ignorat per defecte) ---"
Assert (-not (_ConclusioIgnorarPerDefecte ([pscustomobject]@{ Text='x'; Font='vist_anterior' }))) '_ConclusioIgnorarPerDefecte vist_anterior -> no ignorat'
Assert (-not (_ConclusioIgnorarPerDefecte ([pscustomobject]@{ Text='x'; Font='risc' })))          '_ConclusioIgnorarPerDefecte risc -> no ignorat'
Assert (_ConclusioIgnorarPerDefecte ([pscustomobject]@{ Text='x'; Font='mns' }))                  '_ConclusioIgnorarPerDefecte mns -> ignorat per defecte'
Assert (_ConclusioIgnorarPerDefecte ([pscustomobject]@{ Text='x'; Font='act_extr' }))             '_ConclusioIgnorarPerDefecte act_extr -> ignorat per defecte'
Assert (-not (_ConclusioIgnorarPerDefecte ([pscustomobject]@{ Text=''; Font='' })))               '_ConclusioIgnorarPerDefecte sense conclusio -> no ignorat'

Write-Host "`n--- Informes.ps1: _ConclusioBreu (classificacio de la conclusio en categories curtes) ---"
AssertEq (_ConclusioBreu '')     'Revisar' '_ConclusioBreu buit -> Revisar'
AssertEq (_ConclusioBreu $null)  'Revisar' '_ConclusioBreu null -> Revisar'
AssertEq (_ConclusioBreu "No s${ap}han esmenat les deficiencies indicades al requeriment anterior.") 'Requeriment' '_ConclusioBreu no esmenat -> Requeriment (pendent)'
AssertEq (_ConclusioBreu 'Per tot lo anterior, no es pot donar per finalitzat el present expedient.') 'Requeriment' '_ConclusioBreu "no es pot donar...finalitzat" -> Requeriment'
AssertEq (_ConclusioBreu 'Per tot lo anterior, es pot donar per finalitzat el present tramit.') 'FI Requeriment' '_ConclusioBreu "es pot donar...finalitzat" -> FI Requeriment'
AssertEq (_ConclusioBreu 'Es pot donar per tancada la denuncia.') 'FI Requeriment' '_ConclusioBreu denuncia tancada -> FI Requeriment (fusionat)'
AssertEq (_ConclusioBreu "Vist l${ap}anterior, no es pot donar per tancat l${ap}expedient sancionador.") 'Requeriment' '_ConclusioBreu "no es pot donar per tancat" -> Requeriment (pendent)'
AssertEq (_ConclusioBreu 'Per tot lo exposat, no es pot donar per tancada la denuncia presentada.') 'Requeriment' '_ConclusioBreu "no es pot donar per tancada la denuncia" -> Requeriment (no FI)'
AssertEq (_ConclusioBreu "Es valora aixecar el precinte de l${ap}activitat.") 'FI Precinte / Cessament' '_ConclusioBreu aixecar precinte -> FI Precinte / Cessament'
AssertEq (_ConclusioBreu "Es considera pertinent desprecintar l${ap}establiment.") 'FI Precinte / Cessament' '_ConclusioBreu pertinent desprecintar -> FI Precinte / Cessament'
AssertEq (_ConclusioBreu 'Es deixa sense efecte la comunicacio previa presentada.') 'Sense efecte' '_ConclusioBreu deixa sense efecte -> Sense efecte'
AssertEq (_ConclusioBreu "Es considera pertinent precintar l${ap}activitat.") 'Precinte / Cessament' '_ConclusioBreu pertinent precintar -> Precinte / Cessament'
AssertEq (_ConclusioBreu 'Tenint en consideracio el risc greu i imminent per a les persones.') 'Precinte / Cessament' '_ConclusioBreu risc greu -> Precinte / Cessament'
AssertEq (_ConclusioBreu "S${ap}ordeni el cessament de l${ap}activitat de manera immediata.") 'Precinte / Cessament' '_ConclusioBreu ordeni el cessament -> Precinte / Cessament'
AssertEq (_ConclusioBreu "S${ap}informa desfavorablement la sol.licitud presentada.") 'Revisar' '_ConclusioBreu desfavorable -> Revisar (mai Favorable)'
AssertEq (_ConclusioBreu "S${ap}informa favorablement la sol.licitud presentada.") 'Favorable' '_ConclusioBreu favorable -> Favorable'
AssertEq (_ConclusioBreu "Es proposa ampliar el termini per a l${ap}esmena de les deficiencies.") 'Ampliació termini' '_ConclusioBreu ampliar el termini -> Ampliacio termini'
AssertEq (_ConclusioBreu "Un cop feta la recepcio del requeriment, caldra esmenar les deficiencies indicades.") 'Requeriment' '_ConclusioBreu recepcio del requeriment -> Requeriment (nou)'
AssertEq (_ConclusioBreu "Vist l${ap}anterior, s${ap}inicia d${ap}ofici el procediment d${ap}esmena, disposant d${ap}un termini d${ap}un mes per a esmenar els defectes constatats.") 'Requeriment' '_ConclusioBreu procediment d''esmena -> Requeriment (nou)'
AssertEq (_ConclusioBreu 'Aquest text no conte cap de les formules reconegudes.') 'Revisar' '_ConclusioBreu text no reconegut -> Revisar'

Write-Host "`n--- Informes.ps1: _ExcelActivitatActualitzada (Camp Info REQUERIT PER DECRET? / PRECINTE ACTIVITAT? amb SI) ---"
# El camp de l'Excel es 'PRECINTE ACTIVITAT?'. Aqui hi havia una prova que
# assegurava justament el contrari ('PRECINTE ACTIVITAT?' -> false) i per aixo el
# criteri equivocat va sobreviure: activitats ben marcades sortien com a
# desactualitzades. Les proves han de dir el que ha de passar, no el que passa.
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='SI' })) $true '_ExcelActivitatActualitzada PRECINTE ACTIVITAT? + SI -> true'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='precinte activitat?'; Valor='si, des del 2022' })) $true '_ExcelActivitatActualitzada: el nom i el valor no distingeixen majuscules'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='REQUERIT PER DECRET?'; Valor='Si' })) $true '_ExcelActivitatActualitzada REQUERIT PER DECRET? + Si -> true'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='NO' })) $false '_ExcelActivitatActualitzada PRECINTE ACTIVITAT? + NO -> false'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='' })) $false '_ExcelActivitatActualitzada valor buit -> false'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='ALTRA COSA?'; Valor='SI' })) $false '_ExcelActivitatActualitzada nom no objectiu -> false'
# La comparacio es EXACTA a proposit: un "comenca per precinte" agafaria
# 'PRECINTE AIXECAT?', que voldria dir exactament el contrari.
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE AIXECAT?'; Valor='SI' })) $false '_ExcelActivitatActualitzada: PRECINTE AIXECAT? NO val (voldria dir el contrari)'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE?'; Valor='SI' })) $false '_ExcelActivitatActualitzada: PRECINTE? sol ja no existeix a l''Excel'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='X?'; Valor='' }, @{ Nom='PRECINTE ACTIVITAT?'; Valor='SI' })) $true '_ExcelActivitatActualitzada algun parell valid entre varis -> true'
AssertEq (_ExcelActivitatActualitzada @()) $false '_ExcelActivitatActualitzada sense parells -> false'
AssertEq (_ExcelActivitatActualitzada $null) $false '_ExcelActivitatActualitzada null -> false'
# Diagnostic: quins camps de la familia hi ha a l'Excel (per veure un canvi de nom).
$mapDiag = @{
    '10' = @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='SI' }, @{ Nom='AFORAMENT'; Valor='120' })
    '20' = @(@{ Nom='REQUERIT PER DECRET?'; Valor='NO' }, @{ Nom='PRECINTE AIXECAT?'; Valor='SI' })
    '30' = @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='NO' })
}
$campsDiag = @(_ExcelCampsPrecinteTrobats $mapDiag)
AssertEq ($campsDiag -join ' | ') 'PRECINTE ACTIVITAT? | PRECINTE AIXECAT? | REQUERIT PER DECRET?' '_ExcelCampsPrecinteTrobats: els de precinte/decret, sense repetits i ordenats'
AssertEq (@(_ExcelCampsPrecinteTrobats $null).Count) 0 '_ExcelCampsPrecinteTrobats null -> buit'
AssertEq (@(_ExcelCampsPrecinteTrobats @{ '1' = @(@{ Nom='AFORAMENT'; Valor='10' }) }).Count) 0 '_ExcelCampsPrecinteTrobats: ignora els camps que no parlen de precinte ni decret'
AssertEq (_ExcelPrecinteCampsText) "'REQUERIT PER DECRET?' o 'PRECINTE ACTIVITAT?'" '_ExcelPrecinteCampsText: etiqueta dels missatges, treta de la mateixa llista'

Write-Host "`n--- Informes.ps1: _FindCampInfoPairs (localitza parells Nom/Valor dels Camp Info) ---"
$hdrCI = @('ID Activitat','Camp Info 1 - Nom','Camp Info 1 - Valor','Camp Info 2 - Nom','Camp Info 2 - Unitat','Camp Info 2 - Valor')
$pairsCI = _FindCampInfoPairs $hdrCI
AssertEq $pairsCI.Count 2 '_FindCampInfoPairs troba 2 parells (amb Unitat entremig)'
AssertEq $pairsCI[0].NomCol 2 '_FindCampInfoPairs parell 1 NomCol=2 (1-based)'
AssertEq $pairsCI[0].ValorCol 3 '_FindCampInfoPairs parell 1 ValorCol=3'
AssertEq $pairsCI[1].NomCol 4 '_FindCampInfoPairs parell 2 NomCol=4'
AssertEq $pairsCI[1].ValorCol 6 '_FindCampInfoPairs parell 2 ValorCol=6 (salta la Unitat)'
AssertEq ((_FindCampInfoPairs @('ID','Titular','Adreça')).Count) 0 '_FindCampInfoPairs sense Camp Info -> 0 parells'

Write-Host "`n--- ControlsPeriodics.ps1: _ControlPeriodicClassify (annex II/III o apartat 561) ---"
AssertEq (_ControlPeriodicClassify 'II' '').Qualifies      $true  '_ControlPeriodicClassify annex II -> qualifica'
AssertEq (_ControlPeriodicClassify 'II' '').IsII           $true  '_ControlPeriodicClassify annex II -> IsII'
AssertEq (_ControlPeriodicClassify 'III' '').IsIII         $true  '_ControlPeriodicClassify annex III -> IsIII'
AssertEq (_ControlPeriodicClassify 'III' '').IsII          $false '_ControlPeriodicClassify III NO es II (limit de paraula)'
AssertEq (_ControlPeriodicClassify 'Annex II' '').IsII     $true  '_ControlPeriodicClassify "Annex II" -> IsII'
AssertEq (_ControlPeriodicClassify 'I' '').Qualifies       $false '_ControlPeriodicClassify annex I (sense apartat) -> no qualifica'
AssertEq (_ControlPeriodicClassify '' '561').Is561         $true  '_ControlPeriodicClassify apartat 561 -> Is561'
AssertEq (_ControlPeriodicClassify '' '5610').Is561        $true  '_ControlPeriodicClassify apartat 5610 -> Is561'
AssertEq (_ControlPeriodicClassify '' 'Apartat 5610').Is561 $true '_ControlPeriodicClassify "Apartat 5610" -> Is561'
AssertEq (_ControlPeriodicClassify '' '1561').Is561        $false '_ControlPeriodicClassify apartat 1561 -> NO comenca per 561'
AssertEq (_ControlPeriodicClassify '' '').Qualifies        $false '_ControlPeriodicClassify buit -> no qualifica'
AssertEq (_ControlPeriodicClassify 'II' '1561').Qualifies  $true  '_ControlPeriodicClassify II tot i apartat 1561 -> qualifica per annex'

Write-Host "`n--- ControlsPeriodics.ps1: _ParseCellDate (clau d'ordenacio de dates) ---"
AssertEq ((_ParseCellDate '2024-03-01').ToString('yyyy-MM-dd')) '2024-03-01' '_ParseCellDate ISO -> data'
AssertEq ((_ParseCellDate '25/12/2024').ToString('yyyy-MM-dd')) '2024-12-25' '_ParseCellDate dd/MM/yyyy (dia 25 no ambigu) -> data'
AssertEq ([bool]($null -eq (_ParseCellDate ''))) $true '_ParseCellDate buit -> null'
AssertEq ([bool]($null -eq (_ParseCellDate $null))) $true '_ParseCellDate null -> null'
AssertEq ([bool]((_ParseCellDate ([double]45000)) -is [datetime])) $true '_ParseCellDate double OLE -> datetime'

Write-Host "`n--- ControlsPeriodics.ps1: _ControlCatalegKind (precedencia 561 > III > II) ---"
AssertEq (_ControlCatalegKind ([pscustomobject]@{ IsII=$true;  IsIII=$false; Is561=$false })) 'II'  '_ControlCatalegKind nomes II -> II'
AssertEq (_ControlCatalegKind ([pscustomobject]@{ IsII=$false; IsIII=$true;  Is561=$false })) 'III' '_ControlCatalegKind nomes III -> III'
AssertEq (_ControlCatalegKind ([pscustomobject]@{ IsII=$false; IsIII=$false; Is561=$true  })) '112' '_ControlCatalegKind nomes 561 -> 112'
AssertEq (_ControlCatalegKind ([pscustomobject]@{ IsII=$true;  IsIII=$true;  Is561=$true  })) '112' '_ControlCatalegKind II+III+561 -> 112 (561 mana)'
AssertEq (_ControlCatalegKind ([pscustomobject]@{ IsII=$true;  IsIII=$true;  Is561=$false })) 'III' '_ControlCatalegKind II+III -> III'
AssertEq ([bool]($null -eq (_ControlCatalegKind ([pscustomobject]@{ IsII=$false; IsIII=$false; Is561=$false })))) $true '_ControlCatalegKind cap -> null'

Write-Host "`n--- ControlsPeriodics.ps1: _ControlSectionTitle (titol del Titol 2 de REQ1 per regim) ---"
AssertEq (_ControlSectionTitle '112') 'Decret 112/2010 - control periòdic'        '_ControlSectionTitle 112'
AssertEq (_ControlSectionTitle 'III') 'Annex III Llei 20/2009 - control periòdic' '_ControlSectionTitle III'
AssertEq (_ControlSectionTitle 'II')  'Annex II Llei 20/2009 - control periòdic'  '_ControlSectionTitle II'
AssertEq ([bool]($null -eq (_ControlSectionTitle $null))) $true '_ControlSectionTitle desconegut -> null'

Write-Host "`n--- ControlsPeriodics.ps1: _FindItemKeysByTitle (item Titol 2 dins de REQ1) ---"
$parsedCP = [pscustomobject]@{ Sections = @(
    [pscustomobject]@{ Title = 'Controls periòdics'; Items = @(
        [pscustomobject]@{ Kind = 'item'; Short = 'Decret 112/2010 - control periòdic'; Children = @() },
        [pscustomobject]@{ Kind = 'item'; Short = 'Annex II Llei 20/2009 - control periòdic'; Children = @() }
    ) }
) }
$keysCP = @(_FindItemKeysByTitle $parsedCP 'Decret 112/2010 - control periòdic')
AssertEq $keysCP.Count 1 '_FindItemKeysByTitle troba l''item (1 clau)'
AssertEq $keysCP[0] (_ItemKey 'Controls periòdics' 'Decret 112/2010 - control periòdic') '_FindItemKeysByTitle clau correcta'
AssertEq (@(_FindItemKeysByTitle $parsedCP 'Inexistent').Count) 0 '_FindItemKeysByTitle no trobat -> 0 claus'

Write-Host "`n--- CatalegJson.ps1: _RunsToMarkup / _JsonParaToBodyLine ---"
AssertEq (_RunsToMarkup @(@{t='a'}, @{t='b'; b=$true}, @{t='c'; i=$true})) 'a**b**//c//' '_RunsToMarkup negreta + cursiva'
AssertEq (_RunsToMarkup @(@{t='sol text'})) 'sol text' '_RunsToMarkup text sense format'
AssertEq (_RunsToMarkup @()) '' '_RunsToMarkup buit -> cadena buida'
AssertEq (_JsonParaToBodyLine @{ runs=@(@{t='http://x'}); url=$true }) '[[URL]] http://x' '_JsonParaToBodyLine url -> prefix [[URL]]'
AssertEq (_JsonParaToBodyLine @{ runs=@(@{t='normal'}) }) 'normal' '_JsonParaToBodyLine normal'

Write-Host "`n--- CatalegJson.ps1: lectura dels JSON reals (REQ1.json / 0 CONCLUSIONS.json) ---"
$req1Json = Join-Path $EstructuralsDir 'REQ1.json'
if (Test-Path -LiteralPath $req1Json) {
    $reqCat = Read-CatalegJson $req1Json
    AssertEq ([bool]($reqCat.Sections.Count -gt 0)) $true 'Read-CatalegJson REQ1: te seccions'
    AssertEq $reqCat.IsFixedBody $false 'Read-CatalegJson REQ1: no es cos fix'
    $secCP = @($reqCat.Sections | Where-Object { $_.Title -eq 'Controls periòdics' })
    AssertEq ([bool]($secCP.Count -eq 1)) $true 'Read-CatalegJson REQ1: hi ha la seccio "Controls periòdics"'
    $it112 = @($secCP[0].Items | Where-Object { $_.Short -eq 'Decret 112/2010 - control periòdic' })
    AssertEq ([bool]($it112.Count -eq 1 -and $it112[0].Kind -eq 'item')) $true 'Read-CatalegJson REQ1: item "Decret 112/2010 - control periòdic"'
    AssertEq ([bool]($it112[0].BodyLines.Count -gt 0)) $true 'Read-CatalegJson REQ1: l''item te cos'
    # Fills imbricats (subitems): algun item ha de tenir Children (Kind 'child').
    $ambFills = @()
    foreach ($s in $reqCat.Sections) { foreach ($i in $s.Items) { if (@($i.Children).Count -gt 0) { $ambFills += $i } } }
    AssertEq ([bool]($ambFills.Count -gt 0)) $true 'Read-CatalegJson REQ1: hi ha items amb fills imbricats'
    AssertEq ([bool](@($ambFills[0].Children)[0].Kind -eq 'child')) $true 'Read-CatalegJson REQ1: el fill imbricat te Kind child'
} else {
    Write-Host "  (omes: no hi ha REQ1.json)" -ForegroundColor Yellow
}
$conJson = Join-Path $EstructuralsDir '0 CONCLUSIONS.json'
if (Test-Path -LiteralPath $conJson) {
    $conReq = Read-ConclusionsJson $conJson 'REQ1'
    $reqConcl = @($conReq.Selectable | Where-Object { $_.Title -eq 'Requeriment' })
    AssertEq ([bool]($reqConcl.Count -eq 1)) $true 'Read-ConclusionsJson REQ1: hi ha la conclusio "Requeriment"'
    AssertEq ([bool]($reqConcl[0].Body.Length -gt 0)) $true 'Read-ConclusionsJson REQ1: la conclusio te cos'
    AssertEq $conReq.Always.Count 2 'Read-ConclusionsJson: 2 frases ::SEMPRE::'
    $conAll = Read-ConclusionsJson $conJson ''
    AssertEq ([bool]($conAll.Selectable.Count -ge $conReq.Selectable.Count)) $true 'Read-ConclusionsJson buit -> totes les conclusions'
} else {
    Write-Host "  (omes: no hi ha 0 CONCLUSIONS.json)" -ForegroundColor Yellow
}
$terJson = Join-Path $EstructuralsDir 'TERMINI.json'
if (Test-Path -LiteralPath $terJson) {
    $terCat = Read-CatalegJson $terJson
    AssertEq $terCat.IsFixedBody $true 'Read-CatalegJson TERMINI: es cos fix'
    AssertEq ([bool]($terCat.FixedBodyLines.Count -gt 0)) $true 'Read-CatalegJson TERMINI: te cos fix'
}
$aeJson = Join-Path $EstructuralsDir 'ACT_EXTR_REQ.json'
if (Test-Path -LiteralPath $aeJson) {
    # Un sol argument: Parse-ActExtrTemplate ja no rep el Word (llegia el .docx
    # i el respatller es va treure). Se li pot passar el .json o el .docx.
    $aeBlocks = Parse-ActExtrTemplate $aeJson
    AssertEq ([bool](@($aeBlocks).Count -gt 0)) $true 'Parse-ActExtrTemplate (des del JSON): retorna blocs keyed'
    # Els records reconstruits de l'arbre comencen amb un h1 (seccio) i duen h2 keyed.
    $aeRecs = Read-ActExtrRecordsJson $aeJson
    AssertEq ([bool](@($aeRecs).Count -gt 0)) $true 'Read-ActExtrRecordsJson: retorna records'
    AssertEq ([string](@($aeRecs)[0].Style)) 'h1' 'Read-ActExtrRecordsJson: el primer record es una seccio (h1)'
    $h2recs = @($aeRecs | Where-Object { $_.Style -eq 'h2' -and $_.Text -match '^\s*\[\[[A-Z0-9_]+\]\]' })
    AssertEq ([bool]($h2recs.Count -gt 0)) $true 'Read-ActExtrRecordsJson: hi ha blocs h2 amb [[KEY]]'
}

Write-Host "`n--- EditorCatalegs.ps1: funcions pures (segments<->runs) ---"
$edR1 = _Ed_SegmentsToRuns @(@{Text='a';Bold=$false;Italic=$false}, @{Text='b';Bold=$true;Italic=$false}, @{Text='c';Bold=$true;Italic=$false})
AssertEq (_RunsToMarkup $edR1) 'a**bc**' '_Ed_SegmentsToRuns: fusiona negreta consecutiva'
$edR2 = _Ed_SegmentsToRuns @(@{Text='x';Bold=$true;Italic=$true})
AssertEq (_RunsToMarkup $edR2) '**x**' '_Ed_SegmentsToRuns: negreta+cursiva -> nomes negreta (sense solapament)'
$edR3 = _Ed_SegmentsToRuns @(@{Text='p';Bold=$false;Italic=$true})
AssertEq (_RunsToMarkup $edR3) '//p//' '_Ed_SegmentsToRuns: cursiva'
AssertEq (@(_Ed_SegmentsToRuns @()).Count) 1 '_Ed_SegmentsToRuns: buit -> un run buit'
# CosToRich/RichToCos round-trip (aplana igual).
$cosOrig = @(@{ runs=@(@{t='Text '},@{t='fort';b=$true},@{t=' i '},@{t='inclinat';i=$true}); url=$false })
$rich = _Ed_CosToRich $cosOrig
$cosBack = _Ed_RichToCos $rich
AssertEq (_JsonParaToBodyLine @($cosBack)[0]) 'Text **fort** i //inclinat//' '_Ed_CosToRich/_Ed_RichToCos: round-trip de negreta/cursiva'
# Tipus per familia i pare (vocabulari unificat).
AssertEq (@(_Ed_TipusOptions 'cataleg' '') -join ',') 'seccio' '_Ed_TipusOptions cataleg arrel -> seccio'
AssertEq (@(_Ed_TipusOptions 'cataleg' 'seccio') -join ',') 'item,subseccio,text' '_Ed_TipusOptions cataleg sota seccio'
AssertEq (@(_Ed_TipusOptions 'cataleg' 'subseccio') -join ',') 'item' '_Ed_TipusOptions cataleg sota subseccio -> item'
AssertEq (@(_Ed_TipusOptions 'cataleg' 'item') -join ',') 'subitem' '_Ed_TipusOptions cataleg sota item -> subitem'
AssertEq (@(_Ed_TipusOptions 'conclusions' '') -join ',') 'seccio,sempre' '_Ed_TipusOptions conclusions arrel'
AssertEq (@(_Ed_TipusOptions 'conclusions' 'seccio') -join ',') 'item' '_Ed_TipusOptions conclusions sota seccio -> item'
AssertEq (@(_Ed_TipusOptions 'actextr' '') -join ',') 'seccio' '_Ed_TipusOptions actextr arrel -> seccio'
AssertEq (@(_Ed_TipusOptions 'actextr' 'seccio') -join ',') 'item,subitem,text,nota,etiqueta,capcalera,paragraf' '_Ed_TipusOptions actextr sota seccio'
AssertEq (_Ed_ChildTipus 'cataleg' 'seccio') 'item' '_Ed_ChildTipus cataleg seccio -> item'
AssertEq (_Ed_ChildTipus 'cataleg' 'item') 'subitem' '_Ed_ChildTipus cataleg item -> subitem'
AssertEq (_Ed_ChildTipus 'actextr' 'seccio') 'item' '_Ed_ChildTipus actextr seccio -> item'
AssertEq ([bool](_Ed_CanAddChild 'conclusions' @{tipus='seccio'})) $true '_Ed_CanAddChild conclusions seccio'
AssertEq ([bool](_Ed_CanAddChild 'conclusions' @{tipus='item'})) $false '_Ed_CanAddChild conclusions item -> no'
AssertEq ([bool](_Ed_CanAddChild 'cataleg' @{tipus='subseccio'})) $true '_Ed_CanAddChild cataleg subseccio'
AssertEq ([bool](_Ed_CanAddChild 'actextr' @{tipus='seccio'})) $true '_Ed_CanAddChild actextr seccio'
# Etiqueta d'arbre: [tipus] titol (vocabulari nou, sense [[KEY]] ni ::CHILD::).
AssertEq (_Ed_NodeLabel @{tipus='item';titol='Incendis';clau='INCENDIS';cos=@()}) '[item] Incendis' '_Ed_NodeLabel [tipus] titol'

Write-Host "`n--- EditorCatalegs.ps1: model<->JSON sense perdues (els 5 ESTRUCTURALS) ---"
foreach ($docKey in @('REQ1', 'TERMINI', '0 CONCLUSIONS', 'ACT_EXTR_REQ', 'ACT_EXTR_FAV')) {
    $src = Join-Path $EstructuralsDir ($docKey + '.json')
    if (-not (Test-Path -LiteralPath $src)) { continue }
    $o = Get-Content -LiteralPath $src -Raw -Encoding UTF8 | ConvertFrom-Json
    $model = _Ed_JsonToModel $o
    $back = (_Ed_ModelToJson $model) | ConvertTo-Json -Depth 40
    $tmp = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($tmp, $back, (New-Object System.Text.UTF8Encoding($false)))
    $fam = [string]$model.familia
    if ($fam -eq 'cataleg') {
        $da = (Read-CatalegJson $src) | ConvertTo-Json -Depth 40
        $db = (Read-CatalegJson $tmp) | ConvertTo-Json -Depth 40
    } elseif ($fam -eq 'conclusions') {
        $da = (Read-ConclusionsJson $src '') | ConvertTo-Json -Depth 40
        $db = (Read-ConclusionsJson $tmp '') | ConvertTo-Json -Depth 40
    } else {
        $da = (Read-ActExtrRecordsJson $src) | ConvertTo-Json -Depth 40
        $db = (Read-ActExtrRecordsJson $tmp) | ConvertTo-Json -Depth 40
    }
    Assert ($da -eq $db) "round-trip editor ${docKey}: el model del lector es identic (sense perdues)"
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- PdfSignar.ps1: funcions pures (rutes, decisió, arguments AutoFirma) ---"
AssertEq (_PdfPathForDoc 'C:\Inf\2026-01-02 informe.docx') 'C:\Inf\2026-01-02 informe.pdf' '_PdfPathForDoc canvia extensio a .pdf'
AssertEq (_PdfShouldConvert $false ([datetime]'2026-01-01') ([datetime]::MinValue) $false) $true '_PdfShouldConvert: PDF no existeix -> si'
AssertEq (_PdfShouldConvert $true ([datetime]'2026-02-01') ([datetime]'2026-01-01') $false) $true '_PdfShouldConvert: Word mes nou -> si'
AssertEq (_PdfShouldConvert $true ([datetime]'2026-01-01') ([datetime]'2026-02-01') $false) $false '_PdfShouldConvert: PDF al dia -> no'
AssertEq (_PdfShouldConvert $true ([datetime]'2026-01-01') ([datetime]'2026-02-01') $true) $true '_PdfShouldConvert: overwrite -> sempre si'
AssertEq (_CertFilterValue '') '' '_CertFilterValue buit -> sense filtre'
AssertEq (_CertFilterValue '  12345678Z ') 'subject.contains:12345678Z' '_CertFilterValue -> subject.contains:'
# L'ULTIM INFORME GENERAT (valor per defecte del quadre de l'eina).
$tmpGen = Join-Path ([System.IO.Path]::GetTempPath()) ('cornella-ultim-' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $tmpGen -Force)
try {
    AssertEq (_UltimInformeGenerat $tmpGen) '' '_UltimInformeGenerat: carpeta buida -> cadena buida'
    $base = Get-Date '2026-07-01 09:00:00'
    $fitxers = @(
        @{ N = 'vell.docx';             T = $base }
        @{ N = 'ultim.docx';            T = $base.AddHours(3) }   # el mes nou dels Word
        @{ N = 'mitja.doc';             T = $base.AddHours(1) }
        @{ N = '~$ultim.docx';          T = $base.AddHours(9) }   # temporal del Word
        @{ N = 'mes-nou-de-tots.pdf';   T = $base.AddHours(9) }   # no es Word
        @{ N = 'notes.txt';             T = $base.AddHours(9) }
    )
    foreach ($x in $fitxers) {
        $p = Join-Path $tmpGen $x.N
        Set-Content -LiteralPath $p -Value 'x' -Encoding UTF8
        (Get-Item -LiteralPath $p).LastWriteTime = [datetime]$x.T
    }
    AssertEq ([System.IO.Path]::GetFileName((_UltimInformeGenerat $tmpGen))) 'ultim.docx' '_UltimInformeGenerat: el Word mes nou (ignora ~$, .pdf i .txt)'
    AssertEq (_UltimInformeGenerat (Join-Path $tmpGen 'no-hi-es')) '' '_UltimInformeGenerat: carpeta inexistent -> cadena buida'
    AssertEq (_UltimInformeGenerat $null) (_UltimInformeGenerat) '_UltimInformeGenerat: $null es tracta com a "sense carpeta"'
} finally {
    Remove-Item -LiteralPath $tmpGen -Recurse -Force -ErrorAction SilentlyContinue
}
$afArgs = @(_BuildAutoFirmaSignArgv 'C:\a b\in.pdf' 'C:\a b\out.pdf' 'subject.contains:X' '')
AssertEq ($afArgs[0]) 'sign' '_BuildAutoFirmaSignArgv: comenca per sign'
AssertEq ([bool]($afArgs -contains 'windows')) $true '_BuildAutoFirmaSignArgv: magatzem windows'
AssertEq ([bool]($afArgs -contains 'pades')) $true '_BuildAutoFirmaSignArgv: format pades'
AssertEq ($afArgs[[Array]::IndexOf($afArgs, '-i') + 1]) 'C:\a b\in.pdf' '_BuildAutoFirmaSignArgv: entrada com a element propi'
AssertEq ($afArgs[[Array]::IndexOf($afArgs, '-o') + 1]) 'C:\a b\out.pdf' '_BuildAutoFirmaSignArgv: sortida com a element propi'
AssertEq ($afArgs[[Array]::IndexOf($afArgs, '-filter') + 1]) 'subject.contains:X' '_BuildAutoFirmaSignArgv: filtre inclos'
AssertEq ([bool]($afArgs -contains 'SHA256withRSA')) $true '_BuildAutoFirmaSignArgv: algoritme per defecte'
$afNoFilter = @(_BuildAutoFirmaSignArgv 'in.pdf' 'out.pdf' '' 'SHA512withRSA')
AssertEq ([bool]($afNoFilter -contains '-filter')) $false '_BuildAutoFirmaSignArgv: sense filtre si no s''indica'
AssertEq ([bool]($afNoFilter -contains 'SHA512withRSA')) $true '_BuildAutoFirmaSignArgv: algoritme indicat'
AssertEq ([bool](@(_AutoFirmaCandidatePaths).Count -gt 0)) $true '_AutoFirmaCandidatePaths: retorna candidats'
AssertEq ([bool](@(_AutoFirmaCandidatePaths) -like '*AutoFirma.exe')) $true '_AutoFirmaCandidatePaths: apunten a AutoFirma.exe'
AssertEq ([bool]((@(_AutoFirmaCandidatePaths))[0] -is [string])) $true '_AutoFirmaCandidatePaths: elements son cadenes (no la llista sencera)'
AssertEq (_CertCommonName 'CN=NOM COGNOM - 12345678Z, OU=X, O=Y, C=ES') 'NOM COGNOM - 12345678Z' '_CertCommonName extreu el CN'
AssertEq (_CertCommonName 'sense cn') 'sense cn' '_CertCommonName sense CN -> retorna el subjecte'
AssertEq (_CertCommonName '') '' '_CertCommonName buit -> buit'
# Signatura VISIBLE (caixetí).
AssertEq (_AutoFirmaVisibleExtraParams '') '' '_AutoFirmaVisibleExtraParams buit -> buit'
$cxDef = _DefaultCaixeti
AssertEq ([bool]($cxDef -like "*Sergi Fadurdo Modesto*")) $true '_DefaultCaixeti: conté el nom'
AssertEq ([bool]($cxDef -like '*$$SIGNDATE=*')) $true '_DefaultCaixeti: conté el marcador de data'
$ep = _AutoFirmaVisibleExtraParams $cxDef
AssertEq ([bool]($ep -like '*signaturePage=1*')) $true '_AutoFirmaVisibleExtraParams: signaturePage=1'
$cxP = $Script:AutoFirmaCaixetiPos
AssertEq ([bool]($ep -like ('*signaturePositionOnPageLowerLeftX=' + [int]$cxP.LLX + '*'))) $true '_AutoFirmaVisibleExtraParams: hi va la X de la posicio configurada'
AssertEq ([bool]($ep -like ('*signaturePositionOnPageUpperRightY=' + [int]$cxP.URY + '*'))) $true '_AutoFirmaVisibleExtraParams: hi va la Y de la posicio configurada'
# ALINEACIO amb la capçalera de l'informe. Els numeros de referencia son MESURATS
# d'un informe ja generat (descomprimint el flux de contingut de la pagina 1):
#   · la imatge del logo es col·loca amb el seu dalt a y=806,52, pero porta 18 px
#     de blanc a dalt (de 199) = 6,51 pt, o sigui que la punta de l'escut es a
#     y=800,01;
#   · el requadre de la "Nota:" de l'informe va de x=85,58 a x=552,45, que es el
#     marge dret del text.
# Si algu mou el caixeti, aquestes dues proves l'obliguen a saber respecte de que
# l'esta movent.
$refEscutDalt = 800.01
$refTextDreta = 552.45
Assert ([bool]([Math]::Abs([double]$cxP.URY - $refEscutDalt) -le 1.0)) 'AutoFirmaCaixetiPos: el dalt del caixeti va alineat amb la punta de l''escut de la capcalera'
Assert ([bool]([Math]::Abs([double]$cxP.URX - $refTextDreta) -le 1.0)) 'AutoFirmaCaixetiPos: la dreta del caixeti va alineada amb el marge dret del text'
AssertEq ([int]$cxP.URX - [int]$cxP.LLX) 200 'AutoFirmaCaixetiPos: l''amplada del caixeti no ha canviat (nomes s''ha mogut)'
AssertEq ([int]$cxP.URY - [int]$cxP.LLY) 48  'AutoFirmaCaixetiPos: l''alcada del caixeti no ha canviat (nomes s''ha mogut)'
# SEPARADOR: el \n LITERAL (2 caracters), MAI un salt de linia real. Ho fa
# CommandLineLauncher.buildProperties() d'AutoFirma amb indexOf("\\n"), que en
# Java busca els caracters \ + n. Amb salts REALS no en troba cap i es queda amb
# UNA propietat (signaturePage = tota la resta) -> signa pero SENSE caixeti.
AssertEq ([bool]($ep -like '*layer2Text=*')) $true '_AutoFirmaVisibleExtraParams: hi ha layer2Text'
AssertEq ([bool]($ep -match "`n")) $false '_AutoFirmaVisibleExtraParams: cap salt de linia REAL (AutoFirma no els parteix)'
AssertEq (@($ep -split '\\n').Count) 8 '_AutoFirmaVisibleExtraParams: 8 parells separats pel \n LITERAL'
# El layer2Text ha d'anar en UNA SOLA LINIA: si hi posessim el separador, el tros
# seguent no tindria cap '=' i AutoFirma petaria amb "begin 0, end -1, length 21"
# (21 = els caracters de "Enginyer d'Activitats"). Comprovat al registre real.
$valLayer2 = (@($ep -split '\\n') | Where-Object { $_ -like 'layer2Text=*' })
AssertEq ([bool]($valLayer2 -like ('*Sergi Fadurdo Modesto ' + [char]0x00B7 + ' Enginyer*'))) $true '_AutoFirmaVisibleExtraParams: el text va en una linia amb punt volat'
AssertEq ([bool]([string]$valLayer2 -like '*\*')) $false '_AutoFirmaVisibleExtraParams: cap barra invertida dins del layer2Text'
# Mode IMATGE: l'unica manera de tenir el caixeti de diverses linies.
$epImg = _AutoFirmaVisibleExtraParams $cxDef $ara 'imatge'
if ([string]::IsNullOrWhiteSpace($epImg)) {
    Write-Host "  (omes: mode imatge no disponible sense System.Drawing)" -ForegroundColor Yellow
} else {
    AssertEq ([bool]($epImg -like '*signatureRubricImage=*')) $true '_AutoFirmaVisibleExtraParams imatge: hi ha signatureRubricImage'
    AssertEq ([bool]($epImg -like '*layer2Text=*')) $false '_AutoFirmaVisibleExtraParams imatge: sense layer2Text (la imatge ja ho porta tot)'
    AssertEq ([bool]($epImg -like '*signaturePage=1*')) $true '_AutoFirmaVisibleExtraParams imatge: manté la posició'
    AssertEq ([bool]((_AutoFirmaArgvToText @('-config', $epImg)) -like '*<imatge base64, *')) $true '_AutoFirmaArgvToText: al registre la imatge surt resumida, no sencera'
}
# ARGV (array) : AutoFirma agafa el -config TAL QUAL (sense Base64) i el parteix
# pel \n LITERAL; per aixo va com un element propi de l'array, no enganxat.
$argvSense = @(_BuildAutoFirmaSignArgv 'C:\a b\in.pdf' 'C:\a b\out.pdf' 'subject.contains:X' '')
AssertEq ($argvSense[0]) 'sign' '_BuildAutoFirmaSignArgv: primer element sign'
AssertEq ([bool]($argvSense -contains 'C:\a b\in.pdf')) $true '_BuildAutoFirmaSignArgv: la ruta va SENSE cometes (element propi)'
AssertEq ([bool]($argvSense -contains '-config')) $false '_BuildAutoFirmaSignArgv: sense caixetí -> cap -config'
AssertEq ([bool]($argvSense -contains 'subject.contains:X')) $true '_BuildAutoFirmaSignArgv: filtre com a element'
$argvCx = @(_BuildAutoFirmaSignArgv 'i.pdf' 'o.pdf' '' '' $cxDef)
$iCfg = [Array]::IndexOf($argvCx, '-config')
AssertEq ([bool]($iCfg -ge 0)) $true '_BuildAutoFirmaSignArgv: amb caixetí -> hi ha -config'
$valCfg = [string]$argvCx[$iCfg + 1]
AssertEq ([bool]($valCfg -like '*signaturePage=1*' -and $valCfg -like '*layer2Text=*')) $true '_BuildAutoFirmaSignArgv: el -config es el TEXT PLA dels extraParams'
AssertEq (@($valCfg -split '\\n').Count) 8 '_BuildAutoFirmaSignArgv: el -config porta 8 propietats separades pel \n LITERAL'
AssertEq ([bool]($valCfg -match "`n")) $false '_BuildAutoFirmaSignArgv: el -config no porta cap salt de linia REAL'
AssertEq ([bool]($valCfg -match '^[A-Za-z0-9+/=]+$')) $false '_BuildAutoFirmaSignArgv: el -config NO va en Base64'
AssertEq ([bool]((_AutoFirmaArgvToText $argvCx) -like '*<LF>*')) $false '_AutoFirmaArgvToText: ja no hi ha salts REALS a marcar amb <LF>'
# Caixeti INVISIBLE: codi 0 no vol dir que es vegi res.
AssertEq (_PdfTextCaixetiInvisible '/Type/XObject/Subtype/Form/BBox[0 0 0 0]/FormType 1') $true '_PdfTextCaixetiInvisible: /BBox[0 0 0 0] -> signatura invisible'
AssertEq (_PdfTextCaixetiInvisible '/Type/XObject/Subtype/Form/BBox[0 0 200 75]/FormType 1') $false '_PdfTextCaixetiInvisible: BBox amb mida -> caixeti visible'
AssertEq (_PdfTextCaixetiInvisible '') $false '_PdfTextCaixetiInvisible: text buit -> no es pot dir que sigui invisible'
# Barres invertides FINALS: dins de cometes escaparien la cometa de tancament.
AssertEq (_ArgvToCommandLine @('C:\a b\')) '"C:\a b\\"' '_ArgvToCommandLine: barra final duplicada dins de cometes'
# La data la resolem NOSALTRES: AutoFirma no ha de veure cap marcador $$...$$
# (amb ells petava amb "begin 0, end -1, length 21" i no signava).
$ara = [datetime]'2026-07-28 11:39:41'
AssertEq (_ResolveCaixetiDate 'Data: $$SIGNDATE=yyyy.MM.dd HH:mm:ss$$' $ara) 'Data: 2026.07.28 11:39:41' '_ResolveCaixetiDate: substitueix el marcador per la data'
AssertEq (_ResolveCaixetiDate 'sense marcador' $ara) 'sense marcador' '_ResolveCaixetiDate: text sense marcador no es toca'
AssertEq (_ResolveCaixetiDate 'x $$SUBJECTCN$$ y' $ara) 'x  y' '_ResolveCaixetiDate: treu els altres marcadors (AutoFirma no els paeix)'
AssertEq (_ResolveCaixetiDate '' $ara) '' '_ResolveCaixetiDate: buit -> buit'
AssertEq ([bool]((_AutoFirmaVisibleExtraParams $cxDef $ara) -like '*$$*')) $false '_AutoFirmaVisibleExtraParams: cap marcador $$ arriba a AutoFirma'
AssertEq ([bool]((_AutoFirmaVisibleExtraParams $cxDef $ara) -like '*2026.07.28 11:39:41*')) $true '_AutoFirmaVisibleExtraParams: hi surt la data ja resolta'
AssertEq (_CaixetiUnaLinia "a`nb`nc") ('a ' + [char]0x00B7 + ' b ' + [char]0x00B7 + ' c') '_CaixetiUnaLinia: ajunta les linies amb un punt volat'
AssertEq (_CaixetiUnaLinia 'nomes una') 'nomes una' '_CaixetiUnaLinia: una sola linia no canvia'
# LINIA D'ORDRES: PS 5.1 no enquota els elements de -ArgumentList; ho fem nosaltres.
# Sense aixo, "5.- Sergi Fadurdo" arribava tallat i AutoFirma deia "El fichero de
# entrada no existe: I:\...\5.-".
$cmdEspais = _ArgvToCommandLine @('sign', '-i', 'I:\5.- Sergi Fadurdo\a.pdf', '-store', 'windows')
AssertEq ([bool]($cmdEspais -like '*"I:\5.- Sergi Fadurdo\a.pdf"*')) $true '_ArgvToCommandLine: enquota les rutes amb espais'
AssertEq ([bool]($cmdEspais -like 'sign -i *')) $true '_ArgvToCommandLine: els que no tenen espais van sense cometes'
AssertEq ([bool]($cmdEspais -like '*-store windows')) $true '_ArgvToCommandLine: manté l''ordre dels arguments'
$cmdCx = _ArgvToCommandLine @('-config', (_AutoFirmaVisibleExtraParams $cxDef))
AssertEq ([bool]($cmdCx -like '-config "*')) $true '_ArgvToCommandLine: el -config multilínia va entre cometes'
AssertEq ([bool]($cmdCx -like "*signaturePage=1*")) $true '_ArgvToCommandLine: el -config conserva les propietats'
# Dins de les cometes s'hi han de conservar els 8 parells, separats pel \n
# LITERAL (dos caracters). Salts de linia REALS no n'hi ha d'haver cap: AutoFirma
# no els parteix (ho comprova tambe l'asserció de mes amunt).
AssertEq (@($cmdCx -split '\\n').Count) 8 '_ArgvToCommandLine: el -config conserva els 8 parells separats pel \n LITERAL'
AssertEq (@($cmdCx -split "`n").Count) 1 '_ArgvToCommandLine: cap salt de línia REAL a la línia d''ordres'
# LIMIT de la linia d'ordres de Windows (32767). La imatge del caixeti viatja en
# base64 dins de l'ordre; si no hi cap, AutoFirma ni s'arrenca ("El nombre del
# archivo o la extension es demasiado largo"). Per aixo hi ha topalls.
AssertEq ([bool]($Script:MaxCommandLine -le 32767 -and $Script:MaxCommandLine -ge 10000)) $true 'MaxCommandLine: topall sota el limit de Windows amb marge'
AssertEq ([bool]($Script:MaxCaixetiBase64 -lt $Script:MaxCommandLine)) $true 'MaxCaixetiBase64: la imatge ha de deixar lloc per a les rutes'
# Mesurat en una ordre REAL del registre de l'usuari (rutes a la unitat de
# xarxa, filtre amb el CN sencer i les 6 propietats de posicio): tot el que no es
# la imatge ocupa 628 caracters. El marge ha de ser prou gran per a aixo amb
# folgança, i tot plegat ha de quedar sota el limit dur de Windows.
AssertEq ([bool](($Script:MaxCommandLine - $Script:MaxCaixetiBase64) -ge 1200)) $true 'MaxCaixetiBase64: hi ha d''haver marge per a la resta de l''ordre'
AssertEq ([bool]($Script:MaxCommandLine -le 32767)) $true 'MaxCommandLine: per sota del limit dur de Windows'

Write-Host "`n--- PdfSignar.ps1: aspecte del caixeti-imatge ---"
# L'ESCUT ha d'existir de debo: la signatura el dibuixa de fons i, si no hi es,
# el caixeti surt sense escut i ningu no se n'assabenta.
$cxIco = _CaixetiEscutPath
Assert ([bool]($cxIco -like '*cornella.ico')) '_CaixetiEscutPath: apunta a l''escut'
Assert (Test-Path -LiteralPath $cxIco)        '_CaixetiEscutPath: l''escut hi es de debo'
# Els intents van de MES a MENYS qualitat: el primer que hi capiga es el que
# s'usa, o sigui que si l'ordre s'inverteix sempre sortiria el pitjor.
$cxInt = @($Script:CaixetiIntents)
Assert ([bool]($cxInt.Count -ge 2)) 'CaixetiIntents: hi ha respatller si el millor no hi cap'
# NOMES JPEG. Es va provar el PNG (ocupa molt menys amb text pla) i AutoFirma el
# va rebutjar: "Se ha proporcionado una imagen de rubrica que no esta codificada
# en JPEG". El caixeti va caure al respatller de text i no va sortir la imatge.
$cxNoJpeg = @($cxInt | Where-Object { [string]$_.Format -ne 'jpeg' })
AssertEq $cxNoJpeg.Count 0 'CaixetiIntents: TOTS en JPEG (AutoFirma nomes accepta JPEG per a la rubrica)'
$cxSenseQ = @($cxInt | Where-Object { $null -eq $_.Qualitat })
AssertEq $cxSenseQ.Count 0 'CaixetiIntents: cada intent JPEG diu la seva qualitat'
$cxOk = $true
for ($i = 1; $i -lt $cxInt.Count; $i++) {
    # Dins del mateix format, l'escala ha d'anar baixant.
    if ([string]$cxInt[$i].Format -eq [string]$cxInt[$i - 1].Format -and
        [int]$cxInt[$i].Escala -gt [int]$cxInt[$i - 1].Escala) { $cxOk = $false }
}
Assert $cxOk 'CaixetiIntents: ordenats de mes a menys resolucio'

# --- Lectura del .ico de l'escut -----------------------------------------------
# El .ico es un CONTENIDOR amb diverses mides a dins i el de l'Ajuntament les
# porta TOTES en PNG. El .NET, amb icones aixi, no les descomprimeix be amb
# Icon.ToBitmap() i l'escut sortia BUIT al caixeti, sense dir-ho ningu. Per aixo
# ens llegim la taula del .ico i n'agafem el PNG que toca.
$icoRaw = [System.IO.File]::ReadAllBytes((_CaixetiEscutPath))
$fr128 = _IcoTriaFrame $icoRaw 128
Assert ($null -ne $fr128)        '_IcoTriaFrame: llegeix el .ico de l''escut'
AssertEq ([int]$fr128.Amplada) 128 '_IcoTriaFrame: demanant-ne 128 dona la de 128'
Assert ([bool]$fr128.EsPng)      '_IcoTriaFrame: les imatges del nostre .ico son PNG (per aixo peta Icon.ToBitmap)'
# Els bytes triats han de ser un PNG de debo (signatura de 8 bytes).
$sig = @(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
$okSig = $true
for ($i = 0; $i -lt 8; $i++) { if ($icoRaw[[int]$fr128.Offset + $i] -ne $sig[$i]) { $okSig = $false } }
Assert $okSig '_IcoTriaFrame: l''offset apunta al principi d''un PNG'
Assert ([bool](([int]$fr128.Offset + [int]$fr128.Mida) -le $icoRaw.Length)) '_IcoTriaFrame: el tros cau dins del fitxer'
# La mes petita que ja sigui prou gran (no la mes gran de totes: no cal
# arrossegar 256x256 per a un escut de 130 px).
AssertEq ([int](_IcoTriaFrame $icoRaw 40).Amplada)  48  '_IcoTriaFrame: agafa la mes petita que hi arribi'
AssertEq ([int](_IcoTriaFrame $icoRaw 64).Amplada)  64  '_IcoTriaFrame: mida exacta'
AssertEq ([int](_IcoTriaFrame $icoRaw 200).Amplada) 256 '_IcoTriaFrame: si en cal una de gran, la gran'
AssertEq ([int](_IcoTriaFrame $icoRaw 9999).Amplada) 256 '_IcoTriaFrame: si cap no hi arriba, la mes gran de totes'
Assert ($null -eq (_IcoTriaFrame $null 32))                   '_IcoTriaFrame: $null -> $null'
Assert ($null -eq (_IcoTriaFrame ([byte[]]@(1, 2, 3)) 32))     '_IcoTriaFrame: massa curt -> $null'
Assert ($null -eq (_IcoTriaFrame ([byte[]](@(0) * 40)) 32))    '_IcoTriaFrame: no es un .ico -> $null'
# El requadre ha de ser prou alt per a les linies del caixeti (si no, la lletra
# sortiria minuscula) i prou baix per no menjar-se la capcalera de l'informe.
$cxPos = $Script:AutoFirmaCaixetiPos
$cxAlt = [int]$cxPos.URY - [int]$cxPos.LLY
$cxNLin = @((_DefaultCaixeti) -split "`n").Count
Assert ([bool](($cxAlt / $cxNLin) -ge 9))  'AutoFirmaCaixetiPos: hi caben les linies amb una mida llegible'
Assert ([bool](($cxAlt / $cxNLin) -le 16)) 'AutoFirmaCaixetiPos: sense espai buit de sobres entre linies'
Assert ([bool]([int]$cxPos.URY -le 842))   'AutoFirmaCaixetiPos: no se surt de la pagina A4'
Assert ([bool]([int]$cxPos.URX -le 595))   'AutoFirmaCaixetiPos: no se surt per la dreta'
AssertEq ([bool]($Script:CaixetiEstil.FactorLletra -gt 0.6)) $true 'CaixetiEstil: la lletra omple la linia (menys espai entre files)'
AssertEq ([bool]($Script:CaixetiEstil.EscutOpacitat -gt 0 -and $Script:CaixetiEstil.EscutOpacitat -lt 1)) $true 'CaixetiEstil: l''escut es de FONS (ni invisible ni opac)'
$argvLlarg = @('sign', '-i', 'a.pdf', '-config', ('x=' + ('A' * 40000)))
AssertEq ([bool]((_ArgvToCommandLine $argvLlarg).Length -gt $Script:MaxCommandLine)) $true '_ArgvToCommandLine: es pot mesurar si una ordre no hi cabria'

# ARREL DE LES PROVES DE RUTES. El programa nomes corre a Windows i munta rutes
# de Windows amb Join-Path / [System.IO.Path], pero totes dues coses depenen de
# la plataforma: Join-Path RESOL LA UNITAT (fora de Windows, 'C:\...' peta amb
# "A drive with the name 'C' does not exist") i [System.IO.Path] no reconeix la
# '\' com a separador en un Linux. Per poder executar la suite a totes dues
# bandes, aquestes proves fan servir una arrel valida a la plataforma on corren i
# comparen amb el separador d'aquella plataforma: el que es comprova es la
# LOGICA (quin nom de carpeta, quin nivell), que es el que ha de ser correcte.
$tstSep   = [string][System.IO.Path]::DirectorySeparatorChar
$tstEstr  = if ($env:OS -eq 'Windows_NT') { 'C:\E' }      else { '/tmp/E' }
$tstVist  = if ($env:OS -eq 'Windows_NT') { 'C:\L\vistes' } else { '/tmp/L/vistes' }
$tstClone = if ($env:OS -eq 'Windows_NT') { 'C:\clone' }  else { '/tmp/clone' }

Write-Host "`n--- VistaWord.ps1: vistes en Word dels catalegs (des dels JSON) ---"
AssertEq (_VistaWordPathFor ($tstEstr + $tstSep + 'REQ1.json') $tstVist) ($tstVist + $tstSep + 'REQ1.docx') "_VistaWordPathFor: mateix nom, .docx, a la carpeta de vistes"
AssertEq (_VistaWordPathFor ($tstEstr + $tstSep + '0 CONCLUSIONS.json') $tstVist) ($tstVist + $tstSep + '0 CONCLUSIONS.docx') '_VistaWordPathFor: respecta els espais del nom'
AssertEq ([bool](_VistaEsProtegit ($tstEstr + $tstSep + '0 CAPCALERA.json'))) $true '_VistaEsProtegit: 0 CAPCALERA no es toca mai'
AssertEq ([bool](_VistaEsProtegit ($tstEstr + $tstSep + 'REQ1.json'))) $false '_VistaEsProtegit: la resta si'
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
AssertEq ([math]::Round($Script:ReportFormatConfig.PageMarginLeftPt, 2)) 85.05 'Format: marge esquerre = 1701 twips de la plantilla'
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

Write-Host "`n--- SincronitzaCatalegs.ps1: protegir els catalegs en actualitzar ---"
# No el carrega Motor.ps1 (l'executa Actualitzar.bat pel seu compte); el
# dot-sourcegem aqui. El $env:GENINFORME_TEST fa que nomes en surtin definicions.
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'SincronitzaCatalegs.ps1')
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/REQ1.json')) $true '_CatalegEsProtegible: cataleg json'
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/0 CAPCALERA.docx')) $true '_CatalegEsProtegible: la plantilla de la capcalera'
AssertEq ([bool](_CatalegEsProtegible 'docs/dades/email-textos.json')) $true '_CatalegEsProtegible: dades del mobil'
AssertEq ([bool](_CatalegEsProtegible 'ESTRUCTURALS/REQ1.json.bak')) $false '_CatalegEsProtegible: els .bak de l''editor NO'
AssertEq ([bool](_CatalegEsProtegible 'suport/Motor.ps1')) $false '_CatalegEsProtegible: el codi NO'
AssertEq ([bool](_CatalegEsProtegible '')) $false '_CatalegEsProtegible: buit -> no'
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

Write-Host "`n--- EmailTextos.ps1: funcions pures (textos del correu del mobil) ---"
$edefs = _DefaultEmailTextos
AssertEq (@($edefs.Keys).Count) 2 '_DefaultEmailTextos: 2 claus (assumpte, cos)'
AssertEq ([bool]($edefs.Contains('assumpte') -and $edefs.Contains('cos'))) $true '_DefaultEmailTextos: te assumpte i cos'
AssertEq ([bool]([string]$edefs['cos'] -like '*{REQUERIMENTS}*')) $true '_DefaultEmailTextos: el cos te la variable {REQUERIMENTS}'
$efields = @(_EmailTextosFields)
AssertEq ($efields.Count) 2 '_EmailTextosFields: 2 camps'
$fkeys = ($efields | ForEach-Object { $_.Key }) -join ','
$dkeys = (@($edefs.Keys)) -join ','
AssertEq $fkeys $dkeys '_EmailTextosFields: claus i ordre coincideixen amb els defaults'
$eload = _LoadEmailTextos
AssertEq ([bool]($eload.Contains('cos') -and ([string]$eload['cos']).Contains('seuelectronica'))) $true '_LoadEmailTextos: el cos conte l''enllac de la seu'
AssertEq ([bool]([string]$eload['cos'] -like '*{REQUERIMENTS}*')) $true '_LoadEmailTextos: el cos conserva {REQUERIMENTS}'
AssertEq ([bool]([string]$eload['assumpte'] -like '*{ID_GIA}*')) $true '_LoadEmailTextos: assumpte conserva {ID_GIA}'

Write-Host "`n--- ControlsCpEmail.ps1: avisos de control periodic per correu ---"
$ccdef = _DefaultControlsCpEmail
AssertEq (@($ccdef.Keys).Count) 2 '_DefaultControlsCpEmail: 2 claus (assumpte, cos)'
AssertEq ([bool]($ccdef.Contains('assumpte') -and $ccdef.Contains('cos'))) $true '_DefaultControlsCpEmail: te assumpte i cos'
AssertEq ([bool]([string]$ccdef['cos'] -like '*{ACTIVITAT}*' -and [string]$ccdef['cos'] -like '*{ADRECA}*' -and [string]$ccdef['cos'] -like '*{PROPER_CP}*')) $true '_DefaultControlsCpEmail: el cos te les variables clau'
AssertEq ([bool]([string]$ccdef['assumpte'] -like '*{ID_GIA}*')) $true '_DefaultControlsCpEmail: assumpte te {ID_GIA}'
# Destinataris: titular a To, representant a CC (i marxa enrere si en falta un).
$rc1 = _ControlsCpRecipients 'titular@x.cat' 'rep@x.cat'
AssertEq $rc1.To 'titular@x.cat' '_ControlsCpRecipients: titular a To'
AssertEq $rc1.Cc 'rep@x.cat'     '_ControlsCpRecipients: representant a CC'
AssertEq $rc1.Ok $true           '_ControlsCpRecipients: dos correus -> Ok'
$rc2 = _ControlsCpRecipients '' 'rep@x.cat'
AssertEq $rc2.To 'rep@x.cat' '_ControlsCpRecipients: sense titular -> representant a To'
AssertEq $rc2.Cc ''          '_ControlsCpRecipients: sense titular -> CC buit'
$rc3 = _ControlsCpRecipients 'titular@x.cat' ''
AssertEq $rc3.To 'titular@x.cat' '_ControlsCpRecipients: sense representant -> titular a To'
$rc4 = _ControlsCpRecipients '  ' 'no-es-un-correu'
AssertEq $rc4.Ok $false '_ControlsCpRecipients: cap correu valid -> Ok=false'
# Substitucio de variables amb una fila d'activitat.
$fila = [pscustomobject]@{ ActPrincipal='BAR'; Adreca='C/ Major 1'; Id='361'; RaoSocial='ACME SL'; ProperCP='10/01/2026'; DataControlPer='10/01/2024' }
$sub = _FillControlsCpPh 'Activitat {ACTIVITAT} a {ADRECA} (GIA {ID_GIA}), data {PROPER_CP}' $fila
AssertEq $sub 'Activitat BAR a C/ Major 1 (GIA 361), data 10/01/2026' '_FillControlsCpPh: substitueix les variables'
# HTML: escapa, negreta i enllacos.
AssertEq (_ControlsCpLineHtml 'a & b < c') 'a &amp; b &lt; c' '_ControlsCpLineHtml: escapa &, <'
AssertEq (_ControlsCpLineHtml '**negreta**') '<b>negreta</b>' '_ControlsCpLineHtml: **negreta** -> <b>'
AssertEq (_ControlsCpLineHtml 'veure http://x.cat/a ok') 'veure <a href="http://x.cat/a">http://x.cat/a</a> ok' '_ControlsCpLineHtml: enllac http -> <a>'
$html = _ControlsCpEmailHtml "linia1`n`nlinia2"
AssertEq ([bool]($html -like '*<div>linia1</div>*' -and $html -like '*<div>linia2</div>*')) $true '_ControlsCpEmailHtml: una linia = un <div>'

Write-Host "`n--- Informes.ps1: _EstatActualActivitat (estat = conclusio breu del darrer informe fiable, per data) ---"
AssertEq (_EstatActualActivitat $null) '' '_EstatActualActivitat null -> buit'
AssertEq (_EstatActualActivitat @()) '' '_EstatActualActivitat llista buida -> buit'
$infsTotIgnorats = @(
    [pscustomobject]@{ data = '2026-01-01'; ignorat = $true; conclusio_breu = 'Requeriment' },
    [pscustomobject]@{ data = '2026-02-01'; ignorat = $true; conclusio_breu = 'Favorable' }
)
AssertEq (_EstatActualActivitat $infsTotIgnorats) '' '_EstatActualActivitat tots ignorats -> buit'
$infsNormal = @(
    [pscustomobject]@{ data = '2026-01-01'; ignorat = $false; conclusio_breu = 'Requeriment' },
    [pscustomobject]@{ data = '2026-02-01'; ignorat = $true;  conclusio_breu = 'Favorable' },
    [pscustomobject]@{ data = '2026-03-01'; ignorat = $false; conclusio_breu = 'FI Requeriment' }
)
AssertEq (_EstatActualActivitat $infsNormal) 'FI Requeriment' '_EstatActualActivitat es queda amb el darrer NO ignorat (per data)'
# QUIN informe decideix l'estat: el mateix que fa servir _EstatActualActivitat.
# Ho necessita "Comprovar Excel" per dir-ne la data (INFORME ENGINYER dd/MM/aaaa).
AssertEq ($null -eq (_InformeQueDeterminaEstat $null)) $true '_InformeQueDeterminaEstat null -> $null'
AssertEq ($null -eq (_InformeQueDeterminaEstat @())) $true '_InformeQueDeterminaEstat llista buida -> $null'
AssertEq ($null -eq (_InformeQueDeterminaEstat $infsTotIgnorats)) $true '_InformeQueDeterminaEstat tots ignorats -> $null'
AssertEq (_InformeQueDeterminaEstat $infsNormal).data '2026-03-01' '_InformeQueDeterminaEstat: el darrer NO ignorat (SALTA el del 02, ignorat)'
AssertEq (_InformeQueDeterminaEstat $infsNormal).conclusio_breu (_EstatActualActivitat $infsNormal) '_InformeQueDeterminaEstat i _EstatActualActivitat parlen del MATEIX informe'

Write-Host "`n--- Informes.ps1: _DataInformeDdMmAaaa ---"
AssertEq (_DataInformeDdMmAaaa '2022-11-11') '11/11/2022' '_DataInformeDdMmAaaa: yyyy-MM-dd -> dd/MM/yyyy'
AssertEq (_DataInformeDdMmAaaa '2026-01-05') '05/01/2026' '_DataInformeDdMmAaaa: conserva els zeros'
AssertEq (_DataInformeDdMmAaaa '2026-03-01T00:00:00') '01/03/2026' '_DataInformeDdMmAaaa: ignora el que hi hagi despres de la data'
AssertEq (_DataInformeDdMmAaaa '') '' '_DataInformeDdMmAaaa: buit -> buit'
AssertEq (_DataInformeDdMmAaaa $null) '' '_DataInformeDdMmAaaa: null -> buit'
AssertEq (_DataInformeDdMmAaaa '11/11/2022') '' '_DataInformeDdMmAaaa: format desconegut -> buit (no inventa res)'

Write-Host "`n--- Informes.ps1: _GiaFromFolderName / _CarpetaActivitat ---"
$p = 'I:\Activitats\Informes\2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL\20260710_Req4.docx'
AssertEq (_GiaFromFolderName $p) '361' '_GiaFromFolderName treu "GIA 361" de la carpeta'
AssertEq (_CarpetaActivitat $p) '2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL' '_CarpetaActivitat = carpeta pare'
AssertEq (_GiaFromFolderName 'I:\Informes\sense marca\x.docx') '' '_GiaFromFolderName sense GIA -> buit'

Write-Host "`n--- Seguiment.ps1: segell d'ultima execucio de les eines ---"
# La marca es desa amb (Get-Date).ToString('o'), o sigui hora LOCAL amb el seu
# desplacament, i es torna a llegir en hora local: el viatge d'anada i tornada ha
# de donar la mateixa hora. Es comprova aixi i no amb una cadena fixa perque
# aquesta maquina pot anar en una zona horaria diferent de la del PC de l'usuari.
$isoLocal = ([datetime]'2026-07-30T14:22:05').ToString('o')
AssertEq (_FormatRunStamp $isoLocal) '30/07/26 14:22' '_FormatRunStamp: ISO -> dd/MM/aa HH:mm (anada i tornada en hora local)'
AssertEq (_FormatRunStamp '') '(mai)' '_FormatRunStamp: buit -> (mai)'
AssertEq (_FormatRunStamp '   ') '(mai)' '_FormatRunStamp: nomes espais -> (mai)'
AssertEq (_FormatRunStamp 'aixo no es una data') '(mai)' '_FormatRunStamp: text que no es data -> (mai)'
AssertEq (_FormatRunStamp $null) '(mai)' '_FormatRunStamp: $null -> (mai)'
# Els tipus d'informe i les pantalles de sistema NO porten segell; qualsevol
# altra accio del menu es una eina i si (aixi una rajola nova hi entra sola).
Assert ('nou' -in $Script:AccionsSenseSegell)          'AccionsSenseSegell: "nou" no porta segell'
Assert ('seguiment' -in $Script:AccionsSenseSegell)    'AccionsSenseSegell: el seguiment d''un informe tampoc'
Assert ('config' -in $Script:AccionsSenseSegell)       'AccionsSenseSegell: la configuracio tampoc'
Assert ('seguimentgia' -notin $Script:AccionsSenseSegell) 'AccionsSenseSegell: l''eina Seguiment del GIA SI que en porta'
Assert ('comprovarexcel' -notin $Script:AccionsSenseSegell) 'AccionsSenseSegell: Comprovar Excel SI que en porta'
Assert ('precintades' -notin $Script:AccionsSenseSegell)   'AccionsSenseSegell: la rajola d''enllac SI que en porta'
# Registre: escriure i tornar a llegir, i que conservi les altres eines.
$segellDir = Join-Path ([System.IO.Path]::GetTempPath()) ('segells-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $segellDir -Force)
$LocalActivitatsDirVell = $LocalActivitatsDir
$LocalActivitatsDir = $segellDir
try {
    AssertEq (_LastRunEina 'seguimentgia') '(mai)' '_LastRunEina: sense registre -> (mai)'
    _MarcaEinaUsada 'seguimentgia'
    AssertEq ([bool]((_LastRunEina 'seguimentgia') -ne '(mai)')) $true '_LastRunEina: despres de marcar-la, hi ha data'
    _MarcaEinaUsada 'comprovarexcel'
    AssertEq ([bool]((_LastRunEina 'seguimentgia') -ne '(mai)')) $true '_MarcaEinaUsada: marcar-ne una altra NO esborra la primera'
    AssertEq @(((Get-Content -LiteralPath (_EinesStatePath) -Raw | ConvertFrom-Json).PSObject.Properties)).Count 2 '_MarcaEinaUsada: un sol fitxer amb una clau per eina'
    # Les dues que tenen marca propia la fan servir si hi es.
    ([pscustomobject]@{ actualitzat_el = '2026-07-28T09:05:00Z' } | ConvertTo-Json) |
        Set-Content -LiteralPath (Join-Path $segellDir 'informes-db.json') -Encoding UTF8
    AssertEq (_LastRunEina 'informesdb') (_FormatRunStamp '2026-07-28T09:05:00Z') '_LastRunEina: informesdb fa servir actualitzat_el (mes precis)'
    AssertEq (_LastRunEina 'einaqueno') '(mai)' '_LastRunEina: una accio sense registre -> (mai)'
} finally {
    $LocalActivitatsDir = $LocalActivitatsDirVell
    Remove-Item -LiteralPath $segellDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Informes.ps1: _NormalitzaExpedient / Build-ExpedientToGiaMap ---"
AssertEq (_NormalitzaExpedient '2025/1/2563')  '2025-1-2563' '_NormalitzaExpedient barres -> guions'
AssertEq (_NormalitzaExpedient '2025-01-2563') '2025-1-2563' '_NormalitzaExpedient treu zeros inicials'
AssertEq (_NormalitzaExpedient '  2025 / 1 / 2563 ') '2025-1-2563' '_NormalitzaExpedient espais'
$fakeCache = [pscustomobject]@{ ById = @{ '361' = @{ EXP_NUM = '2025/1/2563'; TITULAR = 'KRICHI' }; '99' = @{ EXP_NUM = '2024/2/10' } } }
$e2g = Build-ExpedientToGiaMap $fakeCache
AssertEq $e2g['2025-1-2563'] '361' 'Build-ExpedientToGiaMap mapa expedient->GIA'
Assert ($e2g.ContainsKey('2024-2-10')) 'Build-ExpedientToGiaMap inclou totes les files'

Write-Host "`n--- Informes.ps1: _HaDeReprocessar (escaneig incremental) ---"
$prev = [datetime]'2026-07-01T10:00:00Z'
Assert (_HaDeReprocessar ([datetime]'2026-07-05') $prev $false)            'sense entrada previa -> reprocessar (nou)'
Assert (_HaDeReprocessar ([datetime]'2026-07-05') $prev $true)             'modificat despres de l''ultima actualitzacio -> reprocessar'
Assert (-not (_HaDeReprocessar ([datetime]'2026-06-20') $prev $true))      'no tocat des de l''ultima actualitzacio -> reutilitzar'
Assert (-not (_HaDeReprocessar $prev $prev $true))                          'mateixa data exacta -> reutilitzar'

Write-Host "`n--- Informes.ps1: _FlattenInformesDb (aplanat + preservar ignorat) ---"
$fakeDb = [pscustomobject]@{
    actualitzat_el = '2026-07-01T10:00:00Z'
    activitats = @(
        [pscustomobject]@{
            id_gia='361'; expedient='2025/1/2563'; titular='KRICHI'; carpeta='2025-1-2563 GIA 361'
            informes = @(
                [pscustomobject]@{ data='2026-07-10'; fitxer='a.docx'; ruta='I:\x\a.docx'; conclusio='Vist...'; modificat='2026-07-10T00:00:00Z'; ignorat=$true;  motiu='' },
                [pscustomobject]@{ data='2025-01-14'; fitxer='b.docx'; ruta='I:\x\b.docx'; conclusio='';        modificat='2025-01-14T00:00:00Z'; ignorat=$false; motiu='sense conclusio' }
            )
        }
    )
}
$flat = _FlattenInformesDb $fakeDb
AssertEq $flat.Count 2                               '_FlattenInformesDb: 2 registres'
AssertEq $flat['I:\x\a.docx'].Gia '361'              '_FlattenInformesDb: GIA de l''activitat'
AssertEq $flat['I:\x\a.docx'].Titular 'KRICHI'       '_FlattenInformesDb: titular de l''activitat'
Assert ([bool]$flat['I:\x\a.docx'].Ignorat)          '_FlattenInformesDb: conserva ignorat=true'
Assert (-not [bool]$flat['I:\x\b.docx'].Ignorat)     '_FlattenInformesDb: conserva ignorat=false'
AssertEq $flat['I:\x\b.docx'].Motius.Count 1         '_FlattenInformesDb: motiu -> Motius'

Write-Host "`n--- Settings.ps1: _ResolveEffectiveValue (override d'aquest PC vs valor per defecte) ---"
AssertEq (_ResolveEffectiveValue 'F:\Informes' 'I:\Informes') 'F:\Informes' '_ResolveEffectiveValue amb override -> guanya l''override'
AssertEq (_ResolveEffectiveValue '' 'I:\Informes')            'I:\Informes' '_ResolveEffectiveValue buit -> per defecte'
AssertEq (_ResolveEffectiveValue '   ' 'I:\Informes')         'I:\Informes' '_ResolveEffectiveValue nomes espais -> per defecte'
AssertEq (_ResolveEffectiveValue $null 'I:\Informes')         'I:\Informes' '_ResolveEffectiveValue null -> per defecte'

Write-Host "`n--- Settings.ps1: _BuildSettingsOverrides (que es desa a settings.json) ---"
$defaults = @{ InformesDir = 'I:\Informes'; ActivitatsDir = 'I:\Activitats'; OutputDir = 'C:\Repo\Sortida' }
$valuesCanviats = @{ InformesDir = 'F:\Informes'; ActivitatsDir = 'I:\Activitats'; OutputDir = 'C:\Repo\Sortida' }
$ov1 = _BuildSettingsOverrides $valuesCanviats $defaults
AssertEq $ov1.Count 1                    '_BuildSettingsOverrides: nomes el camp canviat'
AssertEq $ov1['InformesDir'] 'F:\Informes' '_BuildSettingsOverrides: valor correcte'
Assert (-not $ov1.Contains('ActivitatsDir')) '_BuildSettingsOverrides: igual al per defecte -> no es desa'

$valuesBuits = @{ InformesDir = ''; ActivitatsDir = '   '; OutputDir = 'C:\Repo\Sortida' }
$ov2 = _BuildSettingsOverrides $valuesBuits $defaults
AssertEq $ov2.Count 0 '_BuildSettingsOverrides: camps buits -> cap override (es fa servir el per defecte)'

$valuesTots = @{ InformesDir = 'F:\Informes'; ActivitatsDir = 'F:\Activitats'; OutputDir = 'F:\Sortida' }
$ov3 = _BuildSettingsOverrides $valuesTots $defaults
AssertEq $ov3.Count 3 '_BuildSettingsOverrides: tots els camps diferents -> tots es desen'

Write-Host "`n--- _TextMatches com a filtre de les graelles (Informes / Controls periodics) ---"
# Les graelles li passen la concatenacio dels camps cercables de la fila i el
# text de cerca JA net (.Trim().ToLower()), que es el contracte de la funcio.
$filaHay = 'GIA 361 BAR PEPE Cornella de Llobregat Requeriment'
Assert (_TextMatches $filaHay 'bar')             'filtre graella: troba sense distingir majuscules'
Assert (_TextMatches $filaHay 'pepe')            'filtre graella: troba el titular'
Assert (_TextMatches $filaHay '361')             'filtre graella: troba un numero (GIA)'
Assert (_TextMatches $filaHay 'lla de llo')      'filtre graella: subcadena amb espais'
Assert (_TextMatches $filaHay '')                'filtre graella: cerca buida -> passa tot'
Assert (-not (_TextMatches $filaHay 'restaurant')) 'filtre graella: sense coincidencia -> no passa'
Assert (-not (_TextMatches '' 'x'))              'filtre graella: fila sense text i cerca amb contingut -> no passa'

Write-Host "`n--- SeguimentGia.ps1: els cinc llistats del GIA (funcions pures) ---"
# Taula d'exemple amb la MATEIXA forma que la fulla Estes de la base de dades:
# capceleres amb 'Classificacio general annex' REPETIDA (a la base hi surt dues
# vegades: la 1a es la que es MOSTRA, la 2a la que fa de CRITERI a ANNEX II) i
# dues parelles de Camp Info.
$sgH = @(
    'ID Activitat', ('N' + [char]0x00FA + 'm. expedient '), ('Ra' + [char]0x00F3 + ' social'),
    ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil'), 'Representant legal',
    ('Rep. Leg. M' + [char]0x00F2 + 'bil'), 'Nom comercial activitat ',
    'Emp. Tipus via', 'Emp. Carrer', ('Emp. N' + [char]0x00FA + 'mero'),
    ('Classificaci' + [char]0x00F3 + ' general annex'), ('Classificaci' + [char]0x00F3 + ' general Apartat'),
    ('Data control inicial/verificaci' + [char]0x00F3), ('Data control peri' + [char]0x00F2 + 'dic'),
    'Activitat principal', ('Classificaci' + [char]0x00F3 + ' general annex'), ('Descripci' + [char]0x00F3 + ' lliure'),
    'Camp Info 1 - Nom', 'Camp Info 1 - Valor', 'Camp Info 2 - Nom', 'Camp Info 2 - Valor'
)
# SENSE @(): _FindCampInfoPairs protegeix el seu retorn amb una coma (per al cas
# d'UNA sola parella) i un @() al voltant hi torna a posar la capa. Aixo es
# exactament el que feia petar l'eina Seguiment amb "No se puede convertir el
# valor System.Object[] ... al tipo System.Int32" (vegeu _SgAplanaPairs).
$sgPairs = _FindCampInfoPairs $sgH
AssertEq @($sgPairs).Count 2 'SeguimentGia: es reaprofita _FindCampInfoPairs (2 parelles)'
AssertEq @(_FindCampInfoPairs $sgH | Select-Object -First 1).Count 1 '_FindCampInfoPairs: el retorn es una llista de parelles, no una llista d''una llista'
# La forma EMBOLCALLADA (com si algu hi tornes a posar un @()) i la neta han de
# donar el mateix: _SgAplanaPairs ho aplana al punt d'entrada.
$sgPairsEmbolcallat = @(_FindCampInfoPairs $sgH)
AssertEq @($sgPairsEmbolcallat).Count 1 '_FindCampInfoPairs amb @(): queda embolcallat (per aixo no s''hi posa)'
AssertEq @(_SgAplanaPairs $sgPairs).Count 2 '_SgAplanaPairs: la forma neta es queda igual'
AssertEq @(_SgAplanaPairs $sgPairsEmbolcallat).Count 2 '_SgAplanaPairs: la forma embolcallada s''aplana'
AssertEq @(_SgAplanaPairs $null).Count 0 '_SgAplanaPairs: $null -> llista buida'
AssertEq @(_SgAplanaPairs @()).Count 0 '_SgAplanaPairs: llista buida -> llista buida'
AssertEq ([int](@(_SgAplanaPairs $sgPairsEmbolcallat)[0].NomCol)) 18 '_SgAplanaPairs: els elements son parelles de debo (NomCol es un enter)'
AssertEq (_SgColIndex $sgH 'ID Activitat') 1 '_SgColIndex: 1-based'
AssertEq (_SgColIndex $sgH ('classificacio general annex')) 11 '_SgColIndex: la PRIMERA que coincideix (com el MATCH de la plantilla)'
AssertEq (_SgColIndex $sgH 'no existeix') 0 '_SgColIndex: si no hi es -> 0'
#            1      2         3        4       5      6      7     8     9    10    11   12       13           14           15      16   17        18                      19                20              21
# ATENCIO als PARENTESIS de ('DEN' + [char]0x00DA + 'NCIA?'): dins d'un literal
# @(...) la COMA lliga mes fort que el '+', o sigui que sense parentesis
# 'DEN' + [char]0x00DA + 'NCIA?' es converteix en TRES elements ('DEN', 'U',
# 'NCIA?') i desalinea tota la fila. Va passar: la prova de DENUNCIES no trobava
# res perque la fila tenia 23 columnes en lloc de 21.
$sgFiles = @(
  @('100','E-1','TITULAR U','600','REP U','601','BAR U','C','MAJOR','1','II','5.17.b','2019-03-15 00:00:00.0','2025-07-08 00:00:00.0','BAR','II','Text llarg de descripcio','PRECINTE ACTIVITAT?','SI, precintada',('DEN' + [char]0x00DA + 'NCIA?'),'SI, soroll'),
  @('101','E-2','TITULAR D','602','','','BAR D','AV','PICASSO','6','III','','','','REST','III','','SONOMETRIA?','SI, queixa','',''),
  @('102','E-3','TITULAR T','603','','','BAR T','PG','FERRO','24','II','1.1','','','IND','II','','REQUERIT PER DECRET?','PROCEDIMENT ESMENA','',''),
  @('','E-4','SENSE ID','','','','','','','','II','','','','','II','hi ha descripcio','PRECINTE ACTIVITAT?','SI',''),
  @('103','E-5','TITULAR C','604','','','BAR C','C','LLUNA','9','II','2.2','','','COM','II','','','','','')
)
$sgDefs = @(_SgFullesDef)
AssertEq $sgDefs.Count 5 '_SgFullesDef: 5 pestanyes'

# --- Tria de pestanyes (les caselles de la finestra) ---------------------------
# La llista de caselles i el que construeix el llibre surten de la MATEIXA font,
# _SgOpcionsExport, per no poder-se desincronitzar.
$sgOps = @(_SgOpcionsExport)
AssertEq $sgOps.Count 6 '_SgOpcionsExport: la copia de la base + els 5 llistats'
AssertEq $sgOps[0] (_SgNomEstes) '_SgOpcionsExport: la copia de la base va primer'
AssertEq $sgOps[1] 'PRECINTES'   '_SgOpcionsExport: despres els llistats, en ordre'
AssertEq $sgOps[5] 'ANNEX II'    '_SgOpcionsExport: l ultim es ANNEX II'
Assert (_SgSeleccioTeEstes @((_SgNomEstes), 'PRECINTES')) '_SgSeleccioTeEstes: la troba'
Assert (_SgSeleccioTeEstes @('estes'))                    '_SgSeleccioTeEstes: sense accents i en minuscules tambe'
Assert (-not (_SgSeleccioTeEstes @('PRECINTES')))         '_SgSeleccioTeEstes: si no hi es, no'
Assert (-not (_SgSeleccioTeEstes @()))                    '_SgSeleccioTeEstes: seleccio buida'
AssertEq @(_SgFullesTriades @('PRECINTES', 'ANNEX II')).Count 2 '_SgFullesTriades: nomes les triades'
AssertEq @(_SgFullesTriades @('PRECINTES', 'ANNEX II'))[0].Nom 'PRECINTES' '_SgFullesTriades: en l ordre de _SgFullesDef, no en el de la tria'
AssertEq @(_SgFullesTriades @('ANNEX II', 'PRECINTES'))[0].Nom 'PRECINTES' '_SgFullesTriades: l ordre de la tria no compta'
AssertEq @(_SgFullesTriades @((_SgNomEstes))).Count 0 '_SgFullesTriades: la copia de la base no es cap llistat'
AssertEq @(_SgFullesTriades @()).Count 0              '_SgFullesTriades: res triat -> cap llistat'
AssertEq @(_SgFullesTriades $null).Count 0            '_SgFullesTriades: $null -> cap llistat'
AssertEq @(_SgFullesTriades $sgOps).Count 5           '_SgFullesTriades: totes marcades -> els 5 llistats'
AssertEq @(_SgFullesTriades @('PRECINTES')).Count 1   '_SgFullesTriades: una sola tria no es desenrotlla'
AssertEq ($sgDefs[0].Nom) 'PRECINTES' '_SgFullesDef: la primera es PRECINTES'
$sgPre = @(_SgFilesPerFulla $sgDefs[0] $sgH $sgFiles $sgPairs)
AssertEq $sgPre.Count 1 'PRECINTES: nomes l activitat 100 (la fila sense ID Activitat NO compta)'
AssertEq ($sgPre[0][0]) 1 'PRECINTES: la columna N numera 1..N (no es el numero de fila de l Excel)'
AssertEq ($sgPre[0][1]) '100' 'PRECINTES: ID Activitat'
AssertEq ($sgPre[0][11]) 'PRECINTE ACTIVITAT?' 'PRECINTES: penultima columna = el Camp Info que ha coincidit'
AssertEq ($sgPre[0][12]) 'SI, precintada' 'PRECINTES: ultima columna = el seu valor'
# El criteri es NOMES tenir el camp: el valor pot dir qualsevol cosa.
$sgReq = @(_SgFilesPerFulla $sgDefs[2] $sgH $sgFiles $sgPairs)
AssertEq $sgReq.Count 1 'REQUERIT DECRET: hi entra tot i que el valor no comenci per SI'
AssertEq ($sgReq[0][12]) 'PROCEDIMENT ESMENA' 'REQUERIT DECRET: el valor surt tal qual (aqui NO es demana SI)'
# La 2a parella de Camp Info tambe val (a la plantilla el MATCH mirava tota la fila).
$sgDen = @(_SgFilesPerFulla $sgDefs[1] $sgH $sgFiles $sgPairs)
AssertEq $sgDen.Count 1 ('DEN' + [char]0x00DA + 'NCIES: troba el criteri a la 2a parella de Camp Info')
AssertEq ($sgDen[0][12]) 'SI, soroll' ('DEN' + [char]0x00DA + 'NCIES: valor de la 2a parella')
$sgSon = @(_SgFilesPerFulla $sgDefs[3] $sgH $sgFiles $sgPairs)
AssertEq $sgSon.Count 1 'SONOMETRIA: nomes l activitat 101'
# ANNEX II: classificacio II (2a columna repetida) I descripcio lliure amb text.
$sgAnx = @(_SgFilesPerFulla $sgDefs[4] $sgH $sgFiles $sgPairs)
AssertEq $sgAnx.Count 1 'ANNEX II: nomes la 100 (la 102 i la 103 son II pero sense descripcio; la 101 es III)'
AssertEq ($sgAnx[0][11]) 'II' 'ANNEX II: mostra la PRIMERA columna de classificacio'
AssertEq ($sgAnx[0][13]) '15/03/2019' 'ANNEX II: la data surt com a dd/MM/aaaa, no en cru'
AssertEq ($sgAnx[0][14]) '08/07/2025' 'ANNEX II: la 2a data, igual'
AssertEq ($sgAnx[0][16]) 'Text llarg de descripcio' ('ANNEX II: ultima columna = Descripci' + [char]0x00F3 + ' lliure')
AssertEq ([bool](_SgAnnexCoincideix 'II' 'text')) $true '_SgAnnexCoincideix: II + descripcio -> si'
AssertEq ([bool](_SgAnnexCoincideix 'II' '')) $false '_SgAnnexCoincideix: sense descripcio -> no'
AssertEq ([bool](_SgAnnexCoincideix 'III' 'text')) $false '_SgAnnexCoincideix: annex III -> no'
AssertEq ([bool](_SgAnnexCoincideix 'ii' '  x ')) $true '_SgAnnexCoincideix: no distingeix majuscules'
# Titols i nom del fitxer
$sgData = Get-Date '2026-07-30'
AssertEq (_SgTitolFulla $sgDefs[0] $sgData) 'ACTIVITAT PRECINTADA? 30/07/2026' '_SgTitolFulla: titol + data de la fila 1'
AssertEq (_SgTitolFulla $sgDefs[4] $sgData) 'ANNEXOS II 30/07/2026' '_SgTitolFulla: ANNEX II te el seu propi titol'
AssertEq (_SgNomFitxer $sgData) '2026-07-30 Seguiment GIA.xlsx' '_SgNomFitxer: xlsx per defecte'
AssertEq (_SgNomFitxer $sgData 'pdf') '2026-07-30 Seguiment GIA.pdf' '_SgNomFitxer: pdf'
# Missatge d'error: ha de dir el PAS i la LINIA, no nomes el missatge pelat. Amb
# COM, un "Unable to cast object of type X to type Y" no diu on ha estat, i
# aquesta eina nomes es pot provar amb l'Excel de debo.
$sgErr = $null
try { throw 'petada de prova' } catch { $sgErr = $_ }
$sgTxt = _SgTextError $sgErr "bolcant les dades a 'PRECINTES'"
Assert ($sgTxt -like '*petada de prova*')                 '_SgTextError: hi surt el missatge original'
Assert ($sgTxt -like "*bolcant les dades a 'PRECINTES'*") '_SgTextError: hi surt el pas on estavem'
Assert ($sgTxt -like '*SeguimentGia.ps1, linia*')         '_SgTextError: hi surt la linia del codi'
Assert ((_SgTextError $sgErr '') -like '*petada de prova*') '_SgTextError: sense pas, encara dona el missatge'

# _SgEscriuMatriu: si l'escriptura EN BLOC falla, ha d'escriure cel·la a cel·la i
# deixar les dades igualment. Amb l'Excel de l'usuari l'assignacio en bloc peta
# ("Unable to cast object of type 'System.Object[,]' to type 'System.String'"),
# aixi que aquest cami alternatiu es el que ha de salvar el fitxer.
# La fulla falsa no te InvokeMember de COM, o sigui que el bloc falla sol i es
# prova exactament el que interessa: el respatller.
$sgEscrites = @{}
$sgFullaFalsa = [pscustomobject]@{}
$sgFullaFalsa | Add-Member -MemberType ScriptProperty -Name Cells -Value {
    [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Item -Value {
        param($r, $c)
        $cel = [pscustomobject]@{ R = $r; C = $c }
        $cel | Add-Member -MemberType ScriptProperty -Name Value2 -Value { '' } `
                          -SecondValue { param($v) $script:sgEscrites["$($this.R),$($this.C)"] = $v } -PassThru
    } -PassThru
}
$sgFullaFalsa | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a, $b) throw 'sense COM' }
$sgMat = New-Object 'object[,]' 2, 3
$sgMat[0, 0] = 'a'; $sgMat[0, 1] = 'b'; $sgMat[0, 2] = 'c'
$sgMat[1, 0] = 'd'; $sgMat[1, 1] = 'e'; $sgMat[1, 2] = 'f'
$sgVia = _SgEscriuMatriu $sgFullaFalsa 3 $sgMat
AssertEq $sgVia ('cel' + [char]0x00B7 + 'la') '_SgEscriuMatriu: si el bloc falla, passa a cel·la a cel·la'
AssertEq $sgEscrites.Count 6      '_SgEscriuMatriu: escriu totes les cel·les de la matriu'
AssertEq $sgEscrites['3,1'] 'a'   '_SgEscriuMatriu: la matriu comenca a la fila indicada, columna 1'
AssertEq $sgEscrites['3,3'] 'c'   '_SgEscriuMatriu: ultima columna de la 1a fila'
AssertEq $sgEscrites['4,1'] 'd'   '_SgEscriuMatriu: 2a fila'
AssertEq $sgEscrites['4,3'] 'f'   '_SgEscriuMatriu: ultima cel·la'
# A l'Excel de l'usuari, Value2 nomes accepta CADENES: un Int32 peta igual que
# la matriu. Per aixo tot el que s'escriu passa per _SgValorCella.
AssertEq (_SgValorCella 1) '1'            '_SgValorCella: un enter surt com a text (Value2 nomes accepta cadenes)'
AssertEq (_SgValorCella $null) ''         '_SgValorCella: $null -> cadena buida'
AssertEq (_SgValorCella '') ''            '_SgValorCella: buit -> buit'
AssertEq (_SgValorCella 'SI, precintat') 'SI, precintat' '_SgValorCella: el text no es toca'
Assert ((_SgValorCella 1) -is [string])   '_SgValorCella: retorna SEMPRE un String'
Assert ((_SgValorCella $null) -is [string]) '_SgValorCella: fins i tot amb $null'
# La columna N de veritat es un enter: es el cas que va fer petar l'eina.
$sgMatN = New-Object 'object[,]' 1, 2
$sgMatN[0, 0] = 1; $sgMatN[0, 1] = 'TEXT'
$sgEscrites = @{}
[void](_SgEscriuMatriu $sgFullaFalsa 3 $sgMatN)
AssertEq $sgEscrites['3,1'] '1' '_SgEscriuMatriu: la columna N (enter) s''escriu com a text'
Assert ($sgEscrites['3,1'] -is [string]) '_SgEscriuMatriu: i arriba a la cel·la com a String'
# Si CAP de les tres maneres funciona, ha de petar dient que ha passat a cada una.
$sgFullaMorta = [pscustomobject]@{}
$sgFullaMorta | Add-Member -MemberType ScriptProperty -Name Cells -Value { throw 'sense Cells' }
$sgFullaMorta | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a, $b) throw 'sense Range' }
$sgErrTot = ''
try { [void](_SgEscriuMatriu $sgFullaMorta 3 $sgMat) } catch { $sgErrTot = [string]$_.Exception.Message }
Assert ($sgErrTot -like '*de cap de les tres maneres*') '_SgEscriuMatriu: si tot falla, peta (no es queda mut)'
Assert ($sgErrTot -like '*bloc-invoke:*')               '_SgEscriuMatriu: l''error diu que ha passat a cada intent'
# Amplades: tantes com columnes (N + les de dades).
foreach ($d in $sgDefs) {
    AssertEq (@($d.Amplades).Count) ((@($d.Cols).Count) + 1) ("Amplades de " + $d.Nom + ": una per columna, incloent-hi la 'N'")
}

exit (Write-TestSummary 'RESULTAT')
