# L'eina "Pla verificacio" (PlaVerificacio.ps1, octubre 2026): els indicadors
# de la Memoria anual del Pla de verificacio, a partir de l'Excel d'activitats
# i la base d'informes. Funcions PURES i el que es desa, sense Excel.
# Es DOT-SOURCE des de run-tests.ps1 (mateix ambit i comptador d'asserts).

try {
Write-Host "`n--- Les dates de l'Excel (ExportacioActivitats: text ISO) ---"
# L'Excel d'octubre de 2026 porta les dates com a TEXT "2025-07-25 00:00:00.0".
AssertEq ((_DataDeCella '2025-07-25 00:00:00.0').ToString('yyyy-MM-dd')) '2025-07-25' '_DataDeCella: el text ISO amb hora de l''ExportacioActivitats'
AssertEq ((_DataDeCella '2025-7-5').ToString('yyyy-MM-dd')) '2025-07-05' '_DataDeCella: ISO sense zeros'
AssertEq ((_DataDeCella '25/07/2025').ToString('yyyy-MM-dd')) '2025-07-25' '_DataDeCella: dd/mm/aaaa'
AssertEq ((_DataDeCella 45863.0).ToString('yyyy-MM-dd')) '2025-07-25' '_DataDeCella: data d''Excel (numero)'
AssertEq ((_DataDeCella ([datetime]'2025-07-25 13:30')).ToString('yyyy-MM-dd HH:mm')) '2025-07-25 00:00' '_DataDeCella: un [datetime], sense hora'
Assert ($null -eq (_DataDeCella '') -and $null -eq (_DataDeCella $null) -and $null -eq (_DataDeCella 'pendent') -and $null -eq (_DataDeCella '2025-13-40')) '_DataDeCella: buit, text o data impossible -> $null'
# Excel Comprovar mirava la visita contra aquesta columna i no hi veia cap data.
AssertEq (@(_DatesDeCella '2025-07-25 00:00:00.0') -join ',') '25/07/2025' 'Excel Comprovar: la data de control en text ISO es llegeix (abans no en veia cap)'
AssertEq (@(_DatesDeCella '01/02/2024 i 03/04/2025') -join ',') '01/02/2024,03/04/2025' 'Excel Comprovar: les dates dd/mm/aaaa d''un text, com sempre'

Write-Host "`n--- Pla verificacio: peces ---"
AssertEq "$(_PvAnnexIIoIII 'II')|$(_PvAnnexIIoIII ' iii ')|$(_PvAnnexIIoIII 'L18-Proj. i Cert.')|$(_PvAnnexIIoIII 'I.1')|$(_PvAnnexIIoIII '')" 'II|III|||' '_PvAnnexIIoIII: nomes II i III'
AssertEq "$(_PvPercent 10 30)|$(_PvPercent 0 10)|$(_PvPercent 45 45)|$(_PvPercent 1 3)|$(_PvPercent 5 0)" '33,3%|0%|100%|33,3%|-' '_PvPercent: com a la Memoria (coma, un decimal, sense ,0) i - sense denominador'
AssertEq "$(_PvPrevistes 2025 @{})|$(_PvPrevistes 2026 @{})|$(_PvPrevistes 2027 @{})|$(_PvPrevistes 2027 @{ '2027' = 50 })|$(_PvPrevistes 2026 @{ '2026' = 40 })" '30|45|0|50|40' '_PvPrevistes: les del programa signat (30 i 45), la de la finestra si n''hi ha, 0 si no se''n sap res'
AssertEq ((_PvAnys ([datetime]'2026-10-16') @()) -join ',') '2025,2026' '_PvAnys: del 2025 (primer programa) a l''any d''avui'
AssertEq ((_PvAnys ([datetime]'2026-10-16') @(@{ Control = [datetime]'2027-01-10'; Inspeccio = $null })) -join ',') '2025,2026,2027' '_PvAnys: fins al darrer any amb dades'
AssertEq (_PvAnys ([datetime]'2025-03-01') @()).Count 1 '_PvAnys: un sol any torna una llista (no desenrotllada)'

# Les columnes, amb la capcalera de debo de la fulla Estes (n'hi ha dues de
# 'Classificacio general annex' i diverses amb 'inspeccio' a dins).
$pvH = @('ID Activitat', 'Emp. Tipus via', 'Emp. Carrer', ('Emp. N' + [char]0x00FA + 'mero'), 'Activitat principal', 'Data control inicial/funcionament/verificaci', ('Data control inicial/verificaci' + [char]0x00F3), ('Data inspecci' + [char]0x00F3 + ' sanitat'), ('Inspecci' + [char]0x00F3), ('Classificaci' + [char]0x00F3 + ' general annex'), ('Classificaci' + [char]0x00F3 + ' general annex'))
$pvCol = _PvColumnes $pvH
AssertEq "$($pvCol.ID)|$($pvCol.VIA)|$($pvCol.CARRER)|$($pvCol.NUMERO)|$($pvCol.ACT)|$($pvCol.INSPECCIO)|$($pvCol.ANNEX)" '1|2|3|4|5|9|10' '_PvColumnes: per nom; Inspeccio EXACTA, la primera Classificacio general annex'
Assert ([int]$pvCol.CONTROL -in @(6, 7)) '_PvColumnes: la data de control inicial'
AssertEq (@(_PvColumnesQueFalten $pvCol).Count) 0 '_PvColumnesQueFalten: hi son totes'
Assert ((@(_PvColumnesQueFalten (_PvColumnes @('ID Activitat'))) -join ',').Contains('Inspecci')) '_PvColumnesQueFalten: diu les que no hi son'

Write-Host "`n--- Pla verificacio: els indicadors d'un any ---"
$pvF = {
    param($id, $annex, $ctl, $ins)
    @{ Id = $id; Annex = $annex; Control = $(if ($ctl) { [datetime]$ctl } else { $null }); Inspeccio = $(if ($ins) { [datetime]$ins } else { $null }); Activitat = 'BAR'; Via = 'C'; Carrer = 'MAJOR'; Numero = '1' }
}
$pvFiles = @(
    (& $pvF '161' 'L18-Proj. i Cert.' '2025-03-24' '2025-02-10'),   # programa I denuncia
    (& $pvF '284' 'L18-Proj. i Cert.' '2025-04-08' $null),           # programa, precinte despres
    (& $pvF '287' 'L18-Cert.' '2025-07-25' $null),                    # programa, precinte ABANS del control: no compta
    (& $pvF '546' 'I.1' '2025-04-14' $null),                          # programa (no es II ni III)
    (& $pvF '88' 'III' '2025-12-01' $null),                           # EC-PCA
    (& $pvF '8' 'II' '2025-01-27' $null),                             # annex II: no compta enlloc
    (& $pvF '900' 'L18-Cert.' '2026-01-14' '2025-11-03'),             # control d'un altre any, denuncia d'aquest
    (& $pvF '901' 'L18-Cert.' $null $null)                            # res
)
$pvDb = ('{"activitats":[' +
    '{"id_gia":"284","informes":[{"data":"2025-05-02","conclusio_breu":"Precinte / Cessament","fitxer":"a.docx"},{"data":"2025-04-08","conclusio_breu":"Requeriment","fitxer":"b.docx"}]},' +
    '{"id_gia":"287","informes":[{"data":"2025-06-01","conclusio_breu":"Precinte / Cessament","fitxer":"c.docx"}]},' +
    '{"id_gia":"546","informes":[{"data":"2025-05-01","conclusio_breu":"Precinte / Cessament","ignorat":true,"fitxer":"d.docx"},{"data":"sense data","conclusio_breu":"Precinte / Cessament"}]},' +
    '{"id_gia":"","informes":[{"data":"2025-05-01","conclusio_breu":"Precinte / Cessament"}]}]}') | ConvertFrom-Json
$pvInf = _PvInformesPerGia $pvDb
AssertEq "$(@($pvInf['284']).Count)|$(@($pvInf['287']).Count)|$($pvInf.ContainsKey('546'))|$($pvInf.Count)" '2|1|False|2' '_PvInformesPerGia: per GIA, sense els ignorats, els sense data ni els sense GIA'
$pvI = Get-PlaVerificacioIndicadors $pvFiles 2025 30 $pvInf
AssertEq "$($pvI.Previstes)|$($pvI.Programa)|$($pvI.Grau)|$($pvI.Denuncies)|$($pvI.EcPca)|$($pvI.Greus)|$($pvI.PctGreus)" '30|4|13,3%|2|1|1|25%' 'Get-PlaVerificacioIndicadors: previstes, programa (no II/III), grau, denuncies, EC-PCA (III) i greus'
AssertEq ((@($pvI.Llista) | ForEach-Object { [string]$_.Grup + ':' + [string]$_.Fila.Id }) -join ',') 'programa:161,programa:284,programa:546,programa:287,ecpca:88,denuncies:161,denuncies:900' 'Get-PlaVerificacioIndicadors: la llista per grup i per data'
AssertEq ((@($pvI.Llista) | Where-Object { $null -ne $_.Precinte } | ForEach-Object { [string]$_.Fila.Id + ' ' + [string]$_.Precinte.Fitxer }) -join ',') '284 a.docx' 'la greu: el precinte del dia del control o despres (el d''abans, no; l''ignorat, no)'
$pvI26 = Get-PlaVerificacioIndicadors $pvFiles 2026 45 $pvInf
AssertEq "$($pvI26.Programa)|$($pvI26.Grau)|$($pvI26.Denuncies)|$($pvI26.EcPca)|$($pvI26.Greus)|$($pvI26.PctGreus)" '1|2,2%|0|0|0|0%' 'Get-PlaVerificacioIndicadors: un altre any, nomes el seu'
$pvSense = Get-PlaVerificacioIndicadors $pvFiles 2025 0 $null
AssertEq "$($pvSense.Grau)|$($null -eq $pvSense.Greus)|$($pvSense.PctGreus)|$($pvSense.SenseBase)" '-|True|-|True' 'sense previstes ni base d''informes: - (no un 0 que enganyi)'
$pvBuit = Get-PlaVerificacioIndicadors @() 2025 30 @{}
AssertEq "$($pvBuit.Programa)|$($pvBuit.Grau)|$($pvBuit.PctGreus)|$(@($pvBuit.Llista).Count)" '0|0%|-|0' 'Get-PlaVerificacioIndicadors: sense cap activitat, a zero'

Write-Host "`n--- Pla verificacio: les taules (format de la Memoria) ---"
$pvFe = @(_PvFilesExecucio $pvI)
AssertEq $pvFe.Count 5 'taula de l''execucio: cinc indicadors'
AssertEq (($pvFe | ForEach-Object { [string]$_[1] }) -join '|') '30|4|13,3%|2|1' 'taula de l''execucio: els valors, en ordre'
Assert ([string]$pvFe[0][0] -eq 'Nombre de verificacions previstes al Programa anual' -and ([string]$pvFe[4][0]).Contains('EC-PCA (III)')) 'taula de l''execucio: els textos de la Memoria'
AssertEq ((@(_PvFilesActivitats $pvI) | ForEach-Object { [string]$_[1] }) -join '|') '1|25%' 'taula de les activitats verificades'
AssertEq ((@(_PvFilesActivitats $pvSense) | ForEach-Object { [string]$_[1] }) -join '|') '-|-' 'taula de les activitats verificades: sense base, -'
$pvHtml = _PvTaulesHtml $pvI
Assert ($pvHtml -cmatch '^[\x20-\x7E]*$') 'les taules en HTML son ASCII (els accents com a entitat): el CF_HTML compta bytes'
Assert ($pvHtml.Contains('AVALUACI&#211; DE L&#8217;EXECUCI&#211; DEL PLA') -and $pvHtml.Contains('Seguiment') -and $pvHtml.Contains('Resultat') -and $pvHtml.Contains('font-style:italic') -and $pvHtml.Contains('13,3%')) 'les taules en HTML: titols, capcaleres, cursiva i valors'
$pvCf = _PvCfHtml '<table>x</table>'
$pvNum = { param($k) [int]([regex]::Match($pvCf, $k + ':(\d{10})').Groups[1].Value) }
AssertEq $pvCf.Substring((& $pvNum 'StartFragment'), (& $pvNum 'EndFragment') - (& $pvNum 'StartFragment')) '<table>x</table>' '_PvCfHtml: StartFragment/EndFragment assenyalen el fragment'
AssertEq $pvCf.Substring((& $pvNum 'StartHTML'), (& $pvNum 'EndHTML') - (& $pvNum 'StartHTML')) '<html><body><!--StartFragment--><table>x</table><!--EndFragment--></body></html>' '_PvCfHtml: StartHTML/EndHTML, el document sencer'
Assert ((_PvTaulesText $pvI).Contains("Grau d" + [char]0x2019 + "execuci" + [char]0x00F3 + " del Programa anual`t13,3%")) 'les taules en text (tabulades), per si no s''enganxa l''HTML'
$pvFl = @(_PvFilaLlista @($pvI.Llista)[1])
AssertEq $pvFl.Count @($Script:PvLlistaCapcalera).Count 'la fila de la llista te tantes columnes com la capcalera'
AssertEq "$($pvFl[0])|$($pvFl[1])|$($pvFl[2])|$($pvFl[8])" ('Programa de verificaci' + [char]0x00F3 + '|284|08/04/2025|02/05/2025 a.docx') 'la fila de la llista: grup, ID, data i l''informe de precinte'

Write-Host "`n--- Pla verificacio: les previstes d'un any nou (aquest PC) ---"
$pvTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('pv-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    AssertEq @((Read-PvPrevistes $pvTmp).Keys).Count 0 'Read-PvPrevistes: sense fitxer, cap'
    Save-PvPrevistes @{ '2027' = 50; '2026' = 44 } $pvTmp
    $pvR = Read-PvPrevistes $pvTmp
    AssertEq "$($pvR['2027'])|$($pvR['2026'])|$(_PvPrevistes 2027 $pvR)" '50|44|50' 'Save/Read-PvPrevistes: es desen i es llegeixen'
    Set-Content -LiteralPath $pvTmp -Value 'no es json' -Encoding UTF8
    AssertEq @((Read-PvPrevistes $pvTmp).Keys).Count 0 'Read-PvPrevistes: un fitxer espatllat no peta'
} finally { Remove-Item -LiteralPath $pvTmp -Force -ErrorAction SilentlyContinue }

Write-Host "`n--- Pla verificacio: al menu (guard) ---"
$pvMenu = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1'))
$pvIxEd = $pvMenu.IndexOf("Action = 'informesdbedit'"); $pvIxPv = $pvMenu.IndexOf("Action = 'plaverificacio'"); $pvIxCp = $pvMenu.IndexOf("Action = 'copiarinformes'")
Assert ($pvIxEd -gt 0 -and $pvIxEd -lt $pvIxPv -and $pvIxPv -lt $pvIxCp) 'menu: Pla verificacio entre Editar base i Copiar informes'
Assert ((_SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1')).Contains("'plaverificacio' { Invoke-PlaVerificacio }")) 'menu: la rajola obre l''eina'
} catch {
    Assert $false ("16-pla-verificacio: excepcio no controlada: " + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}
