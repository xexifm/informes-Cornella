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

$summaryColor = if ($script:fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "`n========================================"
Write-Host ("RESULTAT: {0} OK, {1} FAIL" -f $script:pass, $script:fail) -ForegroundColor $summaryColor
Write-Host "========================================"
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
