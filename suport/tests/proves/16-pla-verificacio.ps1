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
# Els programes signats i les fotos (suport\pla-verificacio.json, el de debo).
$pvD = Read-PvDades
AssertEq "$($pvD.Programes['2025'].Previstes)|$($pvD.Programes['2026'].Previstes)|$(@($pvD.Programes['2025'].Annex1).Count)|$(@($pvD.Programes['2026'].Annex1).Count)|$(@($pvD.Programes['2026'].Annex2).Count)" '30|45|30|30|15' 'pla-verificacio.json: 30 previstes el 2025 (30 de l''annex I); 45 el 2026 (30 de l''annex I i 15 aleatories)'
AssertEq "$(@($pvD.Programes['2025'].Annex1)[0])|$(@($pvD.Programes['2025'].Annex1)[29])|$(@($pvD.Programes['2026'].Annex1)[0])|$(@($pvD.Programes['2026'].Annex1)[29])|$(@($pvD.Programes['2026'].Annex2)[14])" '287|33|632|99|138' 'pla-verificacio.json: les activitats dels programes, en l''ordre del PDF'
AssertEq "$(@($pvD.Fotos['2025'].Verificades).Count)|$(@($pvD.Fotos['2025'].Denuncies).Count)|$(@($pvD.Fotos['2025'].Greus).Count)" '17|21|0' 'pla-verificacio.json: la foto del 2025 son les dades de la Memoria (17 verificades, 21 denuncies, cap greu)'
AssertEq @((Read-PvDades (Join-Path ([System.IO.Path]::GetTempPath()) ('no-hi-es-' + [guid]::NewGuid().ToString('N') + '.json'))).Programes.Keys).Count 0 'Read-PvDades: sense fitxer, buit (i no peta)'
AssertEq "$(_PvPrevistes 2025 @{} $pvD)|$(_PvPrevistes 2026 @{} $pvD)|$(_PvPrevistes 2027 @{} $pvD)|$(_PvPrevistes 2027 @{ '2027' = 50 } $pvD)|$(_PvPrevistes 2026 @{ '2026' = 40 } $pvD)|$(_PvPrevistes 2025 @{} $null)" '30|45|0|50|40|0' '_PvPrevistes: les del programa signat (30 i 45), la de la finestra si n''hi ha, 0 si no se''n sap res'
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
    (& $pvF '284' 'L18-Proj. i Cert.' '2025-04-08' $null),           # programa, precinte d'una visita de l'any
    (& $pvF '287' 'L18-Cert.' '2025-07-25' $null),                    # programa, precinte sense visita o d'una visita d'un altre any: no compta
    (& $pvF '546' 'I.1' '2025-04-14' $null),                          # programa (no es II ni III)
    (& $pvF '88' 'III' '2025-12-01' $null),                           # EC-PCA
    (& $pvF '8' 'II' '2025-01-27' $null),                             # annex II: no compta enlloc
    (& $pvF '900' 'L18-Cert.' '2026-01-14' '2025-11-03'),             # control d'un altre any, denuncia d'aquest
    (& $pvF '901' 'L18-Cert.' $null $null)                            # res
)
$pvDb = ('{"activitats":[' +
    '{"id_gia":"284","informes":[{"data":"2025-05-04","conclusio_breu":"Precinte / Cessament","fitxer":"a.docx","visita":"02/05/2025"},{"data":"2025-04-08","conclusio_breu":"Requeriment","fitxer":"b.docx","visita":"08/04/2025"}]},' +
    '{"id_gia":"287","informes":[{"data":"2025-09-01","conclusio_breu":"Precinte / Cessament","fitxer":"c.docx","visita":""},{"data":"2025-02-01","conclusio_breu":"Precinte / Cessament","fitxer":"e.docx","visita":"10/12/2024"},{"data":"2025-10-01","conclusio_breu":"Precinte / Cessament","fitxer":"f.docx"}]},' +
    '{"id_gia":"546","informes":[{"data":"2025-05-01","conclusio_breu":"Precinte / Cessament","ignorat":true,"fitxer":"d.docx"},{"data":"sense data","conclusio_breu":"Precinte / Cessament"}]},' +
    '{"id_gia":"","informes":[{"data":"2025-05-01","conclusio_breu":"Precinte / Cessament"}]}]}') | ConvertFrom-Json
$pvInf = _PvInformesPerGia $pvDb
AssertEq "$(@($pvInf['284']).Count)|$(@($pvInf['287']).Count)|$($pvInf.ContainsKey('546'))|$($pvInf.Count)" '2|3|False|2' '_PvInformesPerGia: per GIA, sense els ignorats, els sense data ni els sense GIA'
$pvI = Get-PlaVerificacioIndicadors $pvFiles 2025 30 $pvInf
AssertEq "$($pvI.Previstes)|$($pvI.Programa)|$($pvI.Grau)|$($pvI.Denuncies)|$($pvI.EcPca)|$($pvI.Greus)|$($pvI.PctGreus)" '30|4|13,3%|2|1|1|25%' 'Get-PlaVerificacioIndicadors: previstes, programa (no II/III), grau, denuncies, EC-PCA (III) i greus'
AssertEq ((@($pvI.Llista) | ForEach-Object { [string]$_.Grup + ':' + [string]$_.Fila.Id }) -join ',') 'programa:161,programa:284,programa:546,programa:287,ecpca:88,denuncies:161,denuncies:900' 'Get-PlaVerificacioIndicadors: la llista per grup i per data'
AssertEq ((@($pvI.Llista) | Where-Object { $null -ne $_.Precinte } | ForEach-Object { [string]$_.Fila.Id + ' ' + [string]$_.Precinte.Fitxer }) -join ',') '284 a.docx' 'la greu: un informe de VISITA D''INSPECCIO de l''any amb Precinte (el de documentacio aportada, no; el d''una visita d''un altre any, no; l''ignorat, no)'
AssertEq ((@($pvInf['284']) | ForEach-Object { if ($_.Visita -is [datetime]) { $_.Visita.ToString('dd/MM/yyyy') } else { '-' } }) -join ',') '02/05/2025,08/04/2025' '_PvInformesPerGia: la data de la visita de l''Objecte'
AssertEq ((@($pvInf['287']) | ForEach-Object { if ($_.Visita -is [datetime]) { $_.Visita.ToString('dd/MM/yyyy') } else { '-' } }) -join ',') '-,10/12/2024,-' '_PvInformesPerGia: sense visita (buida o d''una base d''abans), $null'
# EL MOTIU DE LA VISITA (octubre 2026): una Queixa sola no es del Pla; Pla i
# Queixa + Pla, si; un informe d'abans sense motiu, com abans (si).
$pvMot = { param($m) @{ Data = [datetime]'2025-05-04'; Breu = $Script:EstatPrecinte; Fitxer = "$m.docx"; Visita = [datetime]'2025-05-02'; Motiu = $m } }
AssertEq ((@('queixa', 'pla', 'queixa+pla', '') | ForEach-Object { if ($null -ne (_PvInformePrecinte @(& $pvMot $_) ([datetime]'2025-04-08'))) { 'si' } else { 'no' } }) -join ',') 'no,si,si,si' '_PvInformePrecinte: Queixa sola no compta; Pla, Queixa + Pla i els d''abans (sense motiu), si'
AssertEq ([string](_PvInformePrecinte @((& $pvMot 'queixa'), (& $pvMot 'pla')) ([datetime]'2025-04-08')).Fitxer) 'pla.docx' '_PvInformePrecinte: amb una Queixa i un Pla, compta el del Pla'
$pvDbM = ('{"activitats":[{"id_gia":"284","informes":[{"data":"2025-05-04","conclusio_breu":"Precinte / Cessament","fitxer":"q.docx","visita":"02/05/2025","visita_motiu":"queixa"}]}]}') | ConvertFrom-Json
$pvInfM = _PvInformesPerGia $pvDbM
AssertEq "$(@($pvInfM['284'])[0].Motiu)|$((Get-PlaVerificacioIndicadors $pvFiles 2025 30 $pvInfM).Greus)" 'queixa|0' '_PvInformesPerGia llegeix el motiu, i un precinte d''una Queixa sola no es greu del programa'
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

Write-Host "`n--- Pla verificacio: la FOTO d'un any tancat (la Memoria 2025) ---"
$pvFo = Get-PvFotoIndicadors $pvD.Fotos['2025'] 2025 (_PvPrevistes 2025 @{} $pvD)
AssertEq "$($pvFo.Previstes)|$($pvFo.Programa)|$($pvFo.Grau)|$($pvFo.Denuncies)|$($pvFo.EcPca)|$($pvFo.Greus)|$($pvFo.PctGreus)|$($pvFo.Foto)" '30|10|33,3%|21|7|0|0%|True' 'la foto del 2025 dona les xifres de la Memoria: 30, 10, 33,3%, 21, 7, 0 i 0%'
AssertEq ((@($pvFo.Llista) | Where-Object { $_.Grup -eq 'programa' } | ForEach-Object { [string]$_.Fila.Id }) -join ',') '760,425,397,317,316,313,311,287,284,161' 'la foto: les del programa, en l''ordre de la Memoria'
$pvFoF = @(_PvFilaLlista @($pvFo.Llista)[0])
AssertEq "$($pvFoF[1])|$($pvFoF[2])|$($pvFoF[4])|$($pvFoF[7])" '760||DOCTOR ARUS|L18-Proj. i Cert.' 'la fila d''una foto: sense data (la Memoria no la dona) i amb l''adreca i l''annex'
$pvFoG = Get-PvFotoIndicadors @{ Document = 'x'; Verificades = @(@{ Id = '1'; Annex = 'L18-Cert.' }, @{ Id = '2'; Annex = 'II' }, @{ Id = '3'; Annex = 'L18-Cert.' }); Denuncies = @(); Greus = @('3') } 2030 4
AssertEq "$($pvFoG.Programa)|$($pvFoG.Greus)|$($pvFoG.PctGreus)|$($pvFoG.Grau)" '2|1|50%|50%' 'la foto: les greus de la llista i l''annex II fora'

Write-Host "`n--- Pla verificacio: el seguiment del programa ---"
$pvProg = @{ Document = 'Programa X'; Previstes = 4; Annex1 = @('161', '287', '999', '900'); Annex2 = @('88') }
$pvSg = Get-PvSeguimentPrograma $pvProg $pvFiles 2025
AssertEq ((@($pvSg.Files) | ForEach-Object { [string]$_.Annex + [string]$_.Ordre + ':' + [string]$_.Id + '=' + [string]$_.Estat }) -join ',') ("I1:161=Verificada,I2:287=Verificada,I3:999=No " + [char]0x00E9 + "s a l'Excel,I4:900=Verificada el 2026,II1:88=Verificada") 'Get-PvSeguimentPrograma: verificada, no trobada, verificada un altre any, i l''annex II'
AssertEq "$($pvSg.FetesI)|$($pvSg.TotalI)|$($pvSg.FetesII)|$($pvSg.TotalII)" '2|4|1|1' 'Get-PvSeguimentPrograma: el recompte per annex'
$pvSgF = @(_PvFilaSeguiment @($pvSg.Files)[0])
AssertEq "$($pvSgF.Count)|$($pvSgF[2])|$($pvSgF[4])|$($pvSgF[5])|$($pvSgF[6])" "$(@($Script:PvSegCapcalera).Count)|161|C MAJOR 1|24/03/2025|10/02/2025" 'la fila del seguiment: ID, adreca i les dates'
$pvSg25 = Get-PvSeguimentPrograma $pvD.Programes['2025'] @() 2025 $pvD.Fotos['2025']
AssertEq "$($pvSg25.FetesI)|$($pvSg25.TotalI)" '6|30' 'el seguiment del 2025 amb la foto: 6 de les 30 previstes verificades segons la Memoria'
AssertEq ((@($pvSg25.Files) | Where-Object { $_.Feta } | ForEach-Object { $_.Id }) -join ',') '287,284,311,313,316,317' 'el seguiment del 2025: les de la Memoria que eren al programa'
$pvSg26 = Get-PvSeguimentPrograma $pvD.Programes['2026'] @() 2026
AssertEq "$($pvSg26.TotalI)|$($pvSg26.TotalII)|$($pvSg26.FetesI)" '30|15|0' 'el seguiment del 2026 sense Excel: 45 previstes, cap de feta'

Write-Host "`n--- Pla verificacio: les previstes d'un any nou (aquest PC) ---"
$pvTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('pv-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    AssertEq @((Read-PvPrevistes $pvTmp).Keys).Count 0 'Read-PvPrevistes: sense fitxer, cap'
    Save-PvPrevistes @{ '2027' = 50; '2026' = 44 } $pvTmp
    $pvR = Read-PvPrevistes $pvTmp
    AssertEq "$($pvR['2027'])|$($pvR['2026'])|$(_PvPrevistes 2027 $pvR $pvD)" '50|44|50' 'Save/Read-PvPrevistes: es desen i es llegeixen'
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
