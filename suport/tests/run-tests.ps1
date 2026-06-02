# Proves automatiques de les funcions PURES de GenerarInforme.ps1.
#
# NO prova la part de Word/Excel (COM) ni les finestres (WinForms): aixo
# nomes es pot provar a Windows amb Office. Aqui es validen les funcions de
# logica (parseig de camps, normalitzacio, cerca de columnes, dates, claus...).
#
# Execucio (Windows o Linux amb pwsh):
#   pwsh -File tests/run-tests.ps1
#
# Carrega GenerarInforme.ps1 en mode "headless" (GENINFORME_TEST=1) perque
# no obri finestres ni executi el programa.

$ErrorActionPreference = 'Stop'
$env:GENINFORME_TEST = '1'
# A Linux no existeix LOCALAPPDATA; el donem perque el dot-source no falli.
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GenerarInforme.ps1'
. $scriptPath   # dot-source: defineix les funcions, no executa Main

$script:pass = 0
$script:fail = 0
function Assert($cond, $name) {
    if ($cond) { $script:pass++; Write-Host "  OK   $name" -ForegroundColor Green }
    else       { $script:fail++; Write-Host "  FAIL $name" -ForegroundColor Red }
}
function AssertEq($actual, $expected, $name) {
    Assert ([string]$actual -eq [string]$expected) "$name (esperat '$expected', obtingut '$actual')"
}

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
# Estubem les funcions Format (de Word) per capturar que reben, sense COM.
$global:emitCalls = New-Object System.Collections.ArrayList
function Format-Section    { param($s,$t) [void]$global:emitCalls.Add("SECT|$t") }
function Format-Subsection { param($s,$t) [void]$global:emitCalls.Add("SUB|$t") }
function Format-Item       { param($s,$n,$t,[switch]$IsChild) [void]$global:emitCalls.Add("ITEM|$n|$t") }
function Format-Body       { param($s,$t,[switch]$IsChild) [void]$global:emitCalls.Add("BODY|$t") }
function Format-Url        { param($s,$u,[switch]$IsChild) [void]$global:emitCalls.Add("URL|$u") }
function Format-Spacer     { param($s) }
function Format-Conclusion { param($s,$t) }
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

Write-Host "`n--- Test-StyleMatch (variants d'estil Word EN/CA/ES) ---"
foreach ($s in 'Heading 1','Titulo 1','Título 1','Titol 1','Títol 1','Ttulo1','Ttulo 1','TÍTULO 1') {
    Assert (Test-StyleMatch $s 1) ("Test-StyleMatch nivell 1: '$s' -> true")
}
foreach ($s in 'Heading 2','Titulo 2','Título 2','Ttulo2','Ttulo 2') {
    Assert (Test-StyleMatch $s 2) ("Test-StyleMatch nivell 2: '$s' -> true")
}
Assert (-not (Test-StyleMatch 'Normal' 1))    "Test-StyleMatch 'Normal' nivell 1 -> false"
Assert (-not (Test-StyleMatch 'Cita' 1))      "Test-StyleMatch 'Cita' nivell 1 -> false"
Assert (-not (Test-StyleMatch 'Titulo 1' 2))  "Test-StyleMatch 'Titulo 1' nivell 2 -> false"
Assert (-not (Test-StyleMatch $null 1))       "Test-StyleMatch null -> false"
Assert (-not (Test-StyleMatch '' 1))          "Test-StyleMatch buit -> false"

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
$global:emitCalls = New-Object System.Collections.ArrayList
function Format-Section    { param($s,$t) [void]$global:emitCalls.Add("SECT|$t") }
function Format-Subsection { param($s,$t) [void]$global:emitCalls.Add("SUB|$t") }
function Format-Item       { param($s,$n,$t,[switch]$IsChild) [void]$global:emitCalls.Add("ITEM|$n|$t" + $(if($IsChild){' (fill)'}else{''})) }
function Format-Body       { param($s,$t,[switch]$IsChild) [void]$global:emitCalls.Add('BODY' + $(if($IsChild){'/CH'}else{''}) + "|$t") }
function Format-Url        { param($s,$u,[switch]$IsChild) [void]$global:emitCalls.Add('URL'  + $(if($IsChild){'/CH'}else{''}) + "|$u") }
function Format-Spacer     { param($s) }
function Format-Conclusion { param($s,$t) [void]$global:emitCalls.Add("CONCL|$t") }

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
$bodyChildCalls = @($global:emitCalls | Where-Object { $_ -like 'BODY/CH|*' })
AssertEq $itemCalls.Count 1                "Nomes 1 ITEM (el pare); cap ITEM per als fills"
AssertEq $bodyChildCalls.Count 2           "Els 2 fills surten com a BODY amb sagnia de fill"
AssertEq $bodyChildCalls[0] 'BODY/CH|Fill A.' "Primer fill: text sense numero"
AssertEq $bodyChildCalls[1] 'BODY/CH|Fill B.' "Segon fill: text sense numero"
# Cap fill ha de tenir patro de numeracio "X.Y."
$childHasNum = @($global:emitCalls | Where-Object { $_ -match 'ITEM\|\d+\.\d+\.' }).Count
AssertEq $childHasNum 0                    "Cap fill amb numeracio jerarquica (X.Y.)"

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
AssertEq (_SeguimentOutputName '2026-05-29_Req1_GIA 1379' $d) '2026-06-05_Req1_GIA 1379_SEG.docx' '_SeguimentOutputName esquema programa -> _SEG amb GIA preservat'
AssertEq (_SeguimentOutputName '2026-05-29_Req1_GIA 1379_SEG' $d) '2026-06-05_Req1_GIA 1379_SEG.docx' '_SeguimentOutputName re-seguiment no duplica _SEG'
AssertEq (_SeguimentOutputName 'Informe antic fet a ma' $d) '2026-06-05_Seguiment_Informe antic fet a ma.docx' '_SeguimentOutputName font feta a ma -> prefix Seguiment'
AssertEq (_SeguimentOutputName 'a/b:c' $d) '2026-06-05_Seguiment_a_b_c.docx' '_SeguimentOutputName saneja caracters il-legals'

Write-Host "`n--- Seguiment: _BuildSeguimentModel ---"
$recs = @(
    [pscustomobject]@{ Index=1; Text='Capcalera intro';                 ListString=''; Bold=0 }
    [pscustomobject]@{ Index=2; Text='1. Baixa tensio. Cal legalitzar.'; ListString=''; Bold=9999999 }
    [pscustomobject]@{ Index=3; Text='01/06/2026: No s entrega.';        ListString=''; Bold=-1 }
    [pscustomobject]@{ Index=4; Text='03/06/2026: Falten dades.';        ListString=''; Bold=-1 }
    [pscustomobject]@{ Index=5; Text='2. Alta tensio. Cal projecte.';    ListString=''; Bold=0 }
    [pscustomobject]@{ Index=6; Text="Vist l'anterior, cal requerir.";   ListString=''; Bold=0 }
)
$model = _BuildSeguimentModel $recs
AssertEq $model.Requirements.Count 2          '_BuildSeguimentModel detecta 2 requeriments'
AssertEq $model.LastReqParaIndex 5            '_BuildSeguimentModel ultim requeriment a l index 5'
AssertEq $model.Requirements[0].ParaIndex 2   '_BuildSeguimentModel req1 a l index 2'
AssertEq $model.Requirements[0].Annotations.Count 2 '_BuildSeguimentModel req1 amb 2 anotacions'
AssertEq $model.Requirements[0].Annotations[1].ParaIndex 4 '_BuildSeguimentModel anotacio 2 del req1 a l index 4'
Assert (-not $model.Requirements[0].WasResolved) '_BuildSeguimentModel req1 (bold mixt) -> pendent'
Assert ($model.Requirements[1].WasResolved)      '_BuildSeguimentModel req2 (bold 0) -> resolt'
$startC = _FindConclusionStartIndex (@($recs | ForEach-Object { $_.Text })) $model.LastReqParaIndex $SeguimentConclusionPhrases
AssertEq $startC 6                            '_BuildSeguimentModel + deteccio: conclusions a l index 6'

Write-Host "`n--- Seguiment XML: helpers de lectura (text, negreta, body) ---"
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function New-XmlInfoFromString($xmlText) {
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($xmlText)
    $ns = _NewWordNsMgr $xml
    return [pscustomobject]@{ Path=''; Xml=$xml; Ns=$ns; Body=$xml.SelectSingleNode('//w:body',$ns) }
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
$records = _CollectParaRecordsXml $bp $xi.Ns
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

Write-Host "`n--- Seguiment XML: transformacio completa ---"
$xi2 = New-XmlInfoFromString $docStr
$bp2 = @(_BodyParagraphsXml $xi2)
$model2 = _BuildSeguimentModel (_CollectParaRecordsXml $bp2 $xi2.Ns)
$dec = @(
    [pscustomobject]@{ Resolved=$false; NewComment='No s entrega.' },
    [pscustomobject]@{ Resolved=$true;  NewComment='S entrega correctament.' }
)
$flds = [ordered]@{}; $flds['lloc'] = [pscustomobject]@{ Type='text'; Value='Cornella'; Hint=$null; Options=$null }
$concl = @([pscustomobject]@{ Title='C'; Body='Cos **negreta** aqui.' })
$always = @('Ho poso, [CAMP: lloc].')
_ApplySeguimentTransform -xmlInfo $xi2 -bodyParas $bp2 -model $model2 -conclusionStartIndex 4 `
    -decisions $dec -dateStr '01/06/2026' -conclHeaderText 'CONCLUSIONS' `
    -selectedConclusions $concl -alwaysConclusions $always -fields $flds
$after = @(_BodyParagraphsXml $xi2 | ForEach-Object { [pscustomobject]@{ T=(_ParagraphTextXml $_ $xi2.Ns); B=(_ParagraphBoldStateXml $_ $xi2.Ns) } })
$texts = @($after | ForEach-Object { $_.T })
Assert ($texts -contains '01/06/2026: No s entrega.')        'Transform: anotacio del req1 inserida'
Assert ($texts -contains '01/06/2026: S entrega correctament.') 'Transform: anotacio del req2 inserida'
Assert (-not ($texts -contains 'Vist l anterior, cal requerir.')) 'Transform: conclusio antiga esborrada'
Assert ($texts -contains 'CONCLUSIONS')                       'Transform: titol de conclusions afegit'
Assert ($texts -contains 'Ho poso, Cornella.')               'Transform: [CAMP: lloc] resolt a "Cornella"'
# req1 (index del text "1.\tBaixa tensio.") i la seva anotacio: en negreta
$idxReq1 = [array]::IndexOf($texts, "1.`tBaixa tensio.")
AssertEq $after[$idxReq1].B (-1)     'Transform: req1 pendent -> tot negreta'
AssertEq $after[$idxReq1 + 1].B (-1) 'Transform: anotacio de req1 -> negreta'
$idxReq2 = [array]::IndexOf($texts, '2. Alta tensio.')
AssertEq $after[$idxReq2].B 0        'Transform: req2 resolt -> sense negreta'
AssertEq $after[$idxReq2 + 1].B 0    'Transform: anotacio de req2 -> sense negreta'
# **negreta** inline a la conclusio -> el paragraf "Cos negreta aqui." es mixt
$idxConcl = [array]::IndexOf($texts, 'Cos negreta aqui.')
Assert ($idxConcl -ge 0)             'Transform: conclusio nova present (markdown processat)'
AssertEq $after[$idxConcl].B 9999999 'Transform: **negreta** inline -> run en negreta (mixt)'
# sectPr preservat
Assert ($null -ne $xi2.Body.SelectSingleNode('w:sectPr', $xi2.Ns)) 'Transform: <w:sectPr> preservat'
# El XML resultant ha de ser ESTRICTAMENT valid (com l'obre el Word)
$swT = New-Object System.IO.StringWriter; $xi2.Xml.Save($swT)
Assert (Test-StrictXml $swT.ToString()) 'Transform: el document.xml resultant es estrictament valid (Word)'

Write-Host "`n--- Seguiment XML: round-trip sobre .docx real + Read-ConclusionsXml ---"
$estr = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'ESTRUCTURALS'
$cap  = Join-Path $estr '0 CAPCALERA.docx'
$conc = Join-Path $estr '0 CONCLUSIONS.docx'
if (Test-Path $cap) {
    $xiR = _LoadDocxXml $cap
    $sect = $xiR.Body.SelectSingleNode('w:sectPr', $xiR.Ns)
    [void]$xiR.Body.InsertBefore((_MakeConclusionParagraphXml $xiR 'PROVA_RT' $false), $sect)
    $outRt = Join-Path ([System.IO.Path]::GetTempPath()) ('rt_' + [Guid]::NewGuid().ToString('N') + '.docx')
    _SaveDocxXml $xiR $cap $outRt
    Assert (Test-Path $outRt) 'Round-trip: el .docx s ha desat'
    $xiR2 = _LoadDocxXml $outRt
    $rtTexts = @(_BodyParagraphsXml $xiR2 | ForEach-Object { _ParagraphTextXml $_ $xiR2.Ns })
    Assert ([bool]($rtTexts -match 'PROVA_RT')) 'Round-trip: el paragraf inserit hi es despres de reobrir'
    AssertEq $xiR2.Body.LastChild.LocalName 'sectPr' 'Round-trip: <w:sectPr> segueix sent l ultim fill'
    # El document.xml desat ha de passar el parseig ESTRICTE (com el Word).
    $savedXmlText = _ReadDocxPartText $outRt 'word/document.xml'
    Assert (Test-StrictXml $savedXmlText) 'Round-trip: el document.xml desat es estrictament valid (Word)'
    Remove-Item -LiteralPath $outRt -Force -ErrorAction SilentlyContinue
} else {
    Write-Host '  (omes: no s ha trobat 0 CAPCALERA.docx)'
}
if (Test-Path $conc) {
    $cx = Read-ConclusionsXml $conc
    AssertEq $cx.HeaderText 'CONCLUSIONS' 'Read-ConclusionsXml: HeaderText = CONCLUSIONS'
    Assert ($cx.Selectable.Count -ge 1)   'Read-ConclusionsXml: hi ha conclusions triables'
    Assert ($cx.Always.Count -ge 1)       'Read-ConclusionsXml: hi ha frases ::SEMPRE::'
} else {
    Write-Host '  (omes: no s ha trobat 0 CONCLUSIONS.docx)'
}
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

$summaryColor = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "`n========================================"
Write-Host ("RESULTAT: {0} OK, {1} FAIL" -f $script:pass, $script:fail) -ForegroundColor $summaryColor
Write-Host "========================================"
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
