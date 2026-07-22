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
# Estubem les funcions Format (de Word) per capturar que reben, sense COM.
$global:emitCalls = New-Object System.Collections.ArrayList
function Format-Section    { param($s,$t) [void]$global:emitCalls.Add("SECT|$t") }
function Format-Subsection { param($s,$t) [void]$global:emitCalls.Add("SUB|$t") }
function Format-Item       { param($s,$n,$t,[switch]$IsChild) [void]$global:emitCalls.Add("ITEM|$n|$t") }
function Format-Body       { param($s,$t,[switch]$IsChild) [void]$global:emitCalls.Add("BODY|$t") }
function Format-Bullet     { param($s,$t,[switch]$IsChild) [void]$global:emitCalls.Add("BULLET|$t") }
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
function Format-Bullet     { param($s,$t,[switch]$IsChild) [void]$global:emitCalls.Add('BULLET' + $(if($IsChild){'/CH'}else{''}) + "|$t") }
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
$bulletChildCalls = @($global:emitCalls | Where-Object { $_ -like 'BULLET/CH|*' })
AssertEq $itemCalls.Count 1                "Nomes 1 ITEM (el pare); cap ITEM per als fills"
AssertEq $bulletChildCalls.Count 2         "Els 2 fills surten com a BULLET (punt de llista) amb sagnia de fill"
AssertEq $bulletChildCalls[0] 'BULLET/CH|Fill A.' "Primer fill: punt de llista, sense numero"
AssertEq $bulletChildCalls[1] 'BULLET/CH|Fill B.' "Segon fill: punt de llista, sense numero"
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
    Assert ($cx.Selectable.Count -ge 1)   'Read-ConclusionsXml: hi ha conclusions triables'
    Assert ($cx.Always.Count -ge 1)       'Read-ConclusionsXml: hi ha frases ::SEMPRE::'
    # El titol del bloc s'ha de detectar SEMPRE (surt a tots els informes amb
    # conclusions, incloent els de seguiment), estigui centrat o no.
    Assert (-not [string]::IsNullOrWhiteSpace($cx.HeaderText)) 'Read-ConclusionsXml: detecta el titol del bloc (HeaderText)'

    # Conclusions per TIPUS D'INFORME (grups Ttulo1: REQ1 / TERMINI).
    $titlesOf = { param($r) @($r.Selectable | ForEach-Object { $_.Title }) }
    $req = Read-ConclusionsXml $conc 'REQ1'
    $ter = Read-ConclusionsXml $conc 'TERMINI'
    $reqTitles = & $titlesOf $req
    $terTitles = & $titlesOf $ter
    Assert ($req.Selectable.Count -ge 1)                'Read-ConclusionsXml REQ1: te conclusions'
    Assert ($ter.Selectable.Count -ge 1)               'Read-ConclusionsXml TERMINI: te conclusions'
    Assert ([bool]($reqTitles -contains 'Requeriment')) 'REQ1: inclou la conclusio Requeriment'
    Assert (-not ($reqTitles -contains 'Ampliar'))      'REQ1: NO inclou conclusions de TERMINI (Ampliar)'
    Assert ([bool]($terTitles -contains 'Ampliar'))     'TERMINI: inclou la conclusio Ampliar'
    Assert (-not ($terTitles -contains 'Requeriment'))  'TERMINI: NO inclou conclusions de REQ1 (Requeriment)'
    # El seguiment ha d'oferir NOMES el grup "SEGUIMENT".
    $seg = Read-ConclusionsXml $conc 'SEGUIMENT'
    $segTitles = & $titlesOf $seg
    Assert ($seg.Selectable.Count -ge 1)                'Read-ConclusionsXml SEGUIMENT: te conclusions'
    Assert ([bool]($segTitles -contains 'Finalitzat'))  'SEGUIMENT: inclou la conclusio Finalitzat'
    Assert (-not ($segTitles -contains 'Requeriment'))  'SEGUIMENT: NO inclou conclusions de REQ1'
    Assert (-not ($segTitles -contains 'Ampliar'))      'SEGUIMENT: NO inclou conclusions de TERMINI'
    # El total filtrat (REQ1 + TERMINI) no supera el total sense filtre.
    Assert (($req.Selectable.Count + $ter.Selectable.Count) -le $cx.Selectable.Count) 'Filtrat per tipus <= total sense filtre'
    # Les frases ::SEMPRE:: son globals: surten per a qualsevol tipus.
    Assert ($req.Always.Count -ge 1)                    'REQ1: les frases ::SEMPRE:: segueixen sent globals'
    Assert ($ter.Always.Count -ge 1)                    'TERMINI: les frases ::SEMPRE:: segueixen sent globals'
    # Un tipus inexistent no retorna conclusions (pero si les ::SEMPRE::).
    $none = Read-ConclusionsXml $conc 'NO_EXISTEIX'
    AssertEq $none.Selectable.Count 0                   'Tipus inexistent: cap conclusio triable'
    Assert ($none.Always.Count -ge 1)                   'Tipus inexistent: ::SEMPRE:: encara globals'
} else {
    Write-Host '  (omes: no s ha trobat 0 CONCLUSIONS.docx)'
}

Write-Host "`n--- _ReportTypeFromFileName (dedueix el tipus del nom de l informe) ---"
AssertEq (_ReportTypeFromFileName '2026-06-23_Termini_GIA 1379.docx') 'Termini' '_ReportTypeFromFileName: 2n segment'
AssertEq (_ReportTypeFromFileName '2026-06-23_Req1_GIA 10.docx')      'Req1'    '_ReportTypeFromFileName: REQ1'
AssertEq (_ReportTypeFromFileName 'sense_separadors')                'separadors' '_ReportTypeFromFileName: 2 segments'
AssertEq (_ReportTypeFromFileName 'nomesun')                          ''        '_ReportTypeFromFileName: sense 2n segment -> buit'
AssertEq (_ReportTypeFromFileName $null)                              ''        '_ReportTypeFromFileName: null -> buit'

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

Write-Host "`n--- Informes.ps1: _ExcelActivitatActualitzada (Camp Info REQUERIT PER DECRET? / PRECINTE? amb SI) ---"
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE?'; Valor='SI, requerit' })) $true '_ExcelActivitatActualitzada PRECINTE? + SI -> true'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='REQUERIT PER DECRET?'; Valor='Si' })) $true '_ExcelActivitatActualitzada REQUERIT PER DECRET? + Si -> true'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE?'; Valor='NO' })) $false '_ExcelActivitatActualitzada PRECINTE? + NO -> false'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='ALTRA COSA?'; Valor='SI' })) $false '_ExcelActivitatActualitzada nom no objectiu -> false'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='PRECINTE ACTIVITAT?'; Valor='SI' })) $false '_ExcelActivitatActualitzada "PRECINTE ACTIVITAT?" (nom diferent) -> false'
AssertEq (_ExcelActivitatActualitzada @(@{ Nom='X?'; Valor='' }, @{ Nom='PRECINTE?'; Valor='SI' })) $true '_ExcelActivitatActualitzada algun parell valid entre varis -> true'
AssertEq (_ExcelActivitatActualitzada @()) $false '_ExcelActivitatActualitzada sense parells -> false'
AssertEq (_ExcelActivitatActualitzada $null) $false '_ExcelActivitatActualitzada null -> false'

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

Write-Host "`n--- Informes.ps1: _GiaFromFolderName / _CarpetaActivitat ---"
$p = 'I:\Activitats\Informes\2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL\20260710_Req4.docx'
AssertEq (_GiaFromFolderName $p) '361' '_GiaFromFolderName treu "GIA 361" de la carpeta'
AssertEq (_CarpetaActivitat $p) '2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL' '_CarpetaActivitat = carpeta pare'
AssertEq (_GiaFromFolderName 'I:\Informes\sense marca\x.docx') '' '_GiaFromFolderName sense GIA -> buit'

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

$summaryColor = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "`n========================================"
Write-Host ("RESULTAT: {0} OK, {1} FAIL" -f $script:pass, $script:fail) -ForegroundColor $summaryColor
Write-Host "========================================"
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
