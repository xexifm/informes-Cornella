# Les eines: informes, controls periodics, catalegs, editor, PDF
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

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
# Dins d'una subseccio hi pot anar un TEXT FIX (n'hi ha a REQ1: l'intro de
# "Documentacio (ITC SP)"). Abans nomes s'oferia 'item' i, com que el combo del
# Tipus es bloqueja quan nomes hi ha una opcio, des de l'editor no es podia ni
# crear ni desfer una cosa que el lector si que llegeix.
AssertEq (@(_Ed_TipusOptions 'cataleg' 'subseccio') -join ',') 'item,text' '_Ed_TipusOptions cataleg sota subseccio -> item i text'
AssertEq (@(_Ed_TipusOptions 'cataleg' 'item') -join ',') 'subitem' '_Ed_TipusOptions cataleg sota item -> subitem'

# --- Scroll vertical i ajust a la pantalla (_MidaFinestraDinsPantalla) --------
# El cas real: al PC de casa la pantalla es mes baixa i pantalles d'aquest
# programa (l'editor de catalegs en fa 700 de minim) surten mes altes que l'area
# de treball; els botons de baix quedaven fora i no s'hi podia arribar.

# Si hi cap, no s'hi toca res.
$scA = _MidaFinestraDinsPantalla 900 700 800 600 100 50 0 0 1920 1040
Assert (-not [bool]$scA.Cal) 'Finestra: si hi cap, no cal tocar res'

# Mes alta que l'area: s'encongeix a l'area.
$scB = _MidaFinestraDinsPantalla 900 760 800 700 0 0 0 0 1366 728
AssertEq ([string]$scB.H) '728' 'Finestra: l alcada es retalla a l area de treball'
Assert ([bool]$scB.Cal) 'Finestra: i diu que cal aplicar-ho'

# EL MinimumSize TAMBE: si no es baixa, el Windows no deixa encongir la finestra
# i el retall de l alcada no serveix de res. Es la meitat que fa que funcioni.
AssertEq ([string]$scB.MinH) '700' 'Finestra: un minim que hi cap no es toca'
AssertEq ([string]$scB.MinW) '800' 'Finestra: l amplada minima, si hi cap, no es toca'
$scB2 = _MidaFinestraDinsPantalla 900 800 800 760 0 0 0 0 1366 728
AssertEq ([string]$scB2.MinH) '728' 'Finestra: un minim MES ALT que l area baixa fins a l area'
AssertEq ([string]$scB2.H) '728' 'Finestra: i l alcada tambe'

# La finestra ha de quedar SENCERA a dins: una centrada que sobresurt per baix
# tambe sobresurt per dalt, i llavors ni la barra de titol es pot agafar.
$scC = _MidaFinestraDinsPantalla 600 500 0 0 900 600 0 0 1366 728
AssertEq ([string]$scC.Y) '228' 'Finestra: es puja perque no sobresurti per baix'
AssertEq ([string]$scC.X) '766' 'Finestra: i es corre perque no sobresurti per la dreta'
$scD = _MidaFinestraDinsPantalla 600 500 0 0 -80 -40 0 0 1366 728
AssertEq ([string]$scD.X) '0' 'Finestra: mai per sobre de la vora esquerra'
AssertEq ([string]$scD.Y) '0' 'Finestra: ni per sobre de la vora de dalt'

# Una area mes petita que el minim: la mida final no pot ser mes gran que l area.
$scE = _MidaFinestraDinsPantalla 970 700 836 700 0 0 0 0 800 600
AssertEq ([string]$scE.W) '800' 'Finestra: amplada retallada encara que el minim fos mes gran'
AssertEq ([string]$scE.H) '600' 'Finestra: alcada retallada encara que el minim fos mes gran'

# PROVA DE FONT: cap pantalla es pot quedar sense l'ajust. Les del programa
# passen totes per _NewForm; les que es fan a ma han de cridar-lo elles.
$scFitxers = @(Get-ChildItem -Recurse -Filter '*.ps1' (Split-Path -Parent $TestsDir) |
               Where-Object { $_.FullName -notmatch '[\\/]tests[\\/]' })
$scOrfes = New-Object System.Collections.ArrayList
foreach ($scF in $scFitxers) {
    $scT = Get-Content -LiteralPath $scF.FullName -Raw
    if ($scF.Name -eq 'UiComuns.ps1' -or $scF.Name -eq 'UiFinestra.ps1') { continue }
    $scN = ([regex]::Matches($scT, 'New-Object\s+System\.Windows\.Forms\.Form\b')).Count
    if ($scN -eq 0) { continue }
    $scA2 = ([regex]::Matches($scT, '_AjustaFinestraAPantalla')).Count
    if ($scA2 -lt $scN) {
        [void]$scOrfes.Add(($scF.Name + ": " + $scN + " finestres a ma, " + $scA2 + " ajustades"))
    }
}
Assert ($scOrfes.Count -eq 0) ("Finestres a ma sense ajust a la pantalla: " + ($scOrfes -join '; '))

# --- Canviar de NIVELL un node (Treure / Ficar) ------------------------------
$mvModel = @{
    familia = 'cataleg'
    intro = (New-Object System.Collections.ArrayList)
    nodes = (New-Object System.Collections.ArrayList)
}
$mvSec = @{ tipus='seccio'; titol='S'; clau=''; cos=(New-Object System.Collections.ArrayList); fills=(New-Object System.Collections.ArrayList) }
$mvSub = @{ tipus='subseccio'; titol='Sub'; clau=''; cos=(New-Object System.Collections.ArrayList); fills=(New-Object System.Collections.ArrayList) }
$mvTxt = @{ tipus='text'; titol=''; clau=''; cos=(New-Object System.Collections.ArrayList); fills=(New-Object System.Collections.ArrayList) }
$mvIt  = @{ tipus='item'; titol='I'; clau=''; cos=(New-Object System.Collections.ArrayList); fills=(New-Object System.Collections.ArrayList) }
[void]$mvModel.nodes.Add($mvSec)
[void]$mvSec.fills.Add($mvSub)
[void]$mvSub.fills.Add($mvTxt)
[void]$mvSub.fills.Add($mvIt)

$mvP = _Ed_TrobaPare $mvModel.nodes $mvTxt
Assert ($null -ne $mvP) '_Ed_TrobaPare troba un node imbricat'
AssertEq ([string]$mvP.Pare.titol) 'Sub' '_Ed_TrobaPare: el pare es la subseccio'
AssertEq ([string]$mvP.Index) '0' '_Ed_TrobaPare: la posicio dins dels germans'
AssertEq ([string](_Ed_TrobaPare $mvModel.nodes $mvSec).Pare) '' '_Ed_TrobaPare: un node d arrel no te pare'
Assert ($null -eq (_Ed_TrobaPare $mvModel.nodes @{ tipus='item' })) '_Ed_TrobaPare: un node que no hi es -> null'

AssertEq (_Ed_TipusEnMoure 'cataleg' 'seccio' 'text') 'text' '_Ed_TipusEnMoure: si el tipus val al pare nou, es queda'
AssertEq (_Ed_TipusEnMoure 'cataleg' 'item' 'text') 'subitem' '_Ed_TipusEnMoure: si no val, el primer del pare nou'

# TREURE: el text fix surt de la subseccio i queda a la seccio, just despres.
# Es EXACTAMENT el moviment que calia per posar l'article 4 de l'Ordenanca a la
# seccio Instal-lacions i que no es podia fer des de l'editor.
$mvR = _Ed_MouNivell $mvModel $mvTxt -1
Assert ([bool]$mvR.Ok) '_Ed_MouNivell treure: es pot'
AssertEq ([string]$mvSub.fills.Count) '1' '_Ed_MouNivell treure: surt de la subseccio'
AssertEq ([string]$mvSec.fills.Count) '2' '_Ed_MouNivell treure: entra a la seccio'
Assert ([object]::ReferenceEquals($mvSec.fills[1], $mvTxt)) '_Ed_MouNivell treure: queda DESPRES del pare'
AssertEq ([string]$mvTxt.tipus) 'text' '_Ed_MouNivell treure: el tipus segueix valent a la seccio'

# I torna a entrar: FICAR el posa dins del germa de sobre (la subseccio).
$mvR2 = _Ed_MouNivell $mvModel $mvTxt 1
Assert ([bool]$mvR2.Ok) '_Ed_MouNivell ficar: es pot'
AssertEq ([string]$mvSec.fills.Count) '1' '_Ed_MouNivell ficar: surt de la seccio'
AssertEq ([string]$mvSub.fills.Count) '2' '_Ed_MouNivell ficar: torna a la subseccio'

# Un node d'arrel no es pot treure mes, i el primer germa no te on ficar-se.
$mvR3 = _Ed_MouNivell $mvModel $mvSec -1
Assert (-not [bool]$mvR3.Ok) '_Ed_MouNivell: un node d arrel no es pot treure'
Assert ([string]$mvR3.Motiu -ne '') '_Ed_MouNivell: i diu per que'
$mvR4 = _Ed_MouNivell $mvModel $mvSub 1
Assert (-not [bool]$mvR4.Ok) '_Ed_MouNivell: el primer germa no te on ficar-se'
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
AssertEq (@(_Ed_TipusOptions 'mnstraspas' '') -join ',') 'seccio' '_Ed_TipusOptions mnstraspas arrel -> seccio'
AssertEq (@(_Ed_TipusOptions 'mnstraspas' 'seccio') -join ',') 'text,item' '_Ed_TipusOptions mnstraspas sota seccio'
AssertEq ([bool](_Ed_CanAddChild 'mnstraspas' @{tipus='seccio'})) $true '_Ed_CanAddChild mnstraspas seccio'
AssertEq ([bool](_Ed_CanAddChild 'mnstraspas' @{tipus='text'})) $false '_Ed_CanAddChild mnstraspas text -> no'
AssertEq (_Ed_ChildTipus 'mnstraspas' 'seccio') 'text' '_Ed_ChildTipus mnstraspas -> text'
# LLIC tambe hi era: sense aixo l'editor no deixava afegir fills a cap punt.
AssertEq ([bool](_Ed_CanAddChild 'llicencia' @{tipus='item'})) $true '_Ed_CanAddChild llicencia item'
AssertEq (_Ed_ChildTipus 'llicencia' 'item') 'nodisposa' '_Ed_ChildTipus llicencia item -> nodisposa'
# Etiqueta d'arbre: [tipus] titol (vocabulari nou, sense [[KEY]] ni ::CHILD::).
AssertEq (_Ed_NodeLabel @{tipus='item';titol='Incendis';clau='INCENDIS';cos=@()}) '[item] Incendis' '_Ed_NodeLabel [tipus] titol'

Write-Host "`n--- EditorCatalegs.ps1: model<->JSON sense perdues (els 5 ESTRUCTURALS) ---"
foreach ($docKey in @('REQ1', 'TERMINI', '0 CONCLUSIONS', 'ACT_EXTR_REQ', 'ACT_EXTR_FAV', 'MNSTRAS')) {
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
    } elseif ($fam -eq 'mnstraspas') {
        # El que compta es el que en surt: els paragrafs de cada informe, amb
        # observacions i sense.
        $da = @(); $db = @()
        foreach ($fx in @('mns', 'traspas')) {
            foreach ($ax in @($true, $false)) {
                $da += @(_MnsParagrafs (Read-MnsCataleg $src) $fx $ax) | ForEach-Object { ($_.Tipus + '|' + (@($_.Linies) -join "`n")) }
                $db += @(_MnsParagrafs (Read-MnsCataleg $tmp) $fx $ax) | ForEach-Object { ($_.Tipus + '|' + (@($_.Linies) -join "`n")) }
            }
        }
        $da = ($da -join '###'); $db = ($db -join '###')
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

# COMPATIBILITAT amb els ordinadors dels altres. Comparant el CMS d'un informe
# nostre amb el d'un de signat a ma (mateix certificat!), les uniques
# diferencies eren el SubFilter i que nosaltres encastavem TOTA la cadena
# (arrel + subCA + signant) i el que funciona, nomes el signant.
# ASPECTE del caixeti. Per defecte, el que dibuixa l'AutoFirma (el mateix que
# l'eina "Utilizar un certificado" de l'Adobe): nomes se li diu ON va. El
# caixeti propi (lletra, contorn, escut) NO s'ha esborrat: es recupera posant
# $Script:CaixetiAspecte a 'propi'.
AssertEq ([string]$Script:CaixetiAspecte) 'defecte' 'CaixetiAspecte: per defecte, el de l''AutoFirma'
$casDef = @(_CaixetiCascada "Nom`nCarrec")
AssertEq $casDef.Count 1 '_CaixetiCascada: aspecte per defecte -> un sol intent'
AssertEq ([string]$casDef[0].Mode) 'defecte' '_CaixetiCascada: i es el mode defecte'
AssertEq (@(_CaixetiCascada '').Count) 0 '_CaixetiCascada: sense caixeti no hi ha cap intent'
$epDef = _AutoFirmaVisibleExtraParams "Nom`nCarrec" $null 'defecte'
Assert ($epDef.Contains('signaturePage=1')) '_AutoFirmaVisibleExtraParams: el mode defecte diu ON va'
Assert (-not $epDef.Contains('layer2Text')) '_AutoFirmaVisibleExtraParams: ...i cap text nostre'
Assert (-not $epDef.Contains('signatureRubricImage')) '_AutoFirmaVisibleExtraParams: ...ni cap imatge'
AssertEq (@($epDef -split '\\n').Count) 5 '_AutoFirmaVisibleExtraParams: el mode defecte son NOMES les 5 linies de posicio'
AssertEq (_AutoFirmaVisibleExtraParams '' $null 'defecte') '' '_AutoFirmaVisibleExtraParams: sense caixeti, signatura invisible (com sempre)'
# El caixeti propi segueix sencer al fitxer: si algun dia es vol tornar, hi es.
$capPropi = $Script:CaixetiAspecte
$Script:CaixetiAspecte = 'propi'
$casPropi = @(_CaixetiCascada "Nom`nCarrec")
AssertEq $casPropi.Count 2 '_CaixetiCascada: amb ''propi'' tornen els dos intents (imatge i text)'
AssertEq ([string]$casPropi[0].Mode) 'imatge' '_CaixetiCascada: ...i el primer segueix sent la imatge'
$Script:CaixetiAspecte = $capPropi

$afCompat = @(_AutoFirmaCompatLines)
Assert ([bool]($afCompat -contains 'signatureSubFilter=adbe.pkcs7.detached')) '_AutoFirmaCompatLines: SubFilter classic, com el que valida arreu'
# El nom del parametre porta DUES ENES ("Signning") al Client @firma. Si algu
# el "corregeix", AutoFirma l'ignora en silenci i tornem a encastar la cadena.
Assert ([bool]($afCompat -contains 'includeOnlySignningCertificate=true')) '_AutoFirmaCompatLines: nomes el certificat del signant (ull: Signning, amb dues enes)'
# L'atribut ESS en la versio ANTIGA (RFC 2634): la V2 hi porta un camp 'policies'
# amb les politiques del certificat i el RFC 5035 obliga a validar la cadena
# RESTRINGIDA a aquelles politiques. Es l'unica diferencia que quedava amb la
# signatura de l'Adobe, que valida a tot arreu.
Assert ([bool]($afCompat -contains 'signingCertificateV2=false')) '_AutoFirmaCompatLines: l''atribut ESS sense el camp de politiques'
# Hi han de ser SEMPRE, tambe quan no hi ha caixeti: abans el -config nomes
# existia si hi havia caixeti i una signatura invisible no se n'hauria
# beneficiat.
$afSenseCaix = @(_BuildAutoFirmaSignArgv 'in.pdf' 'out.pdf' '' '' '')
Assert ([bool]($afSenseCaix -contains '-config')) '_BuildAutoFirmaSignArgv: hi ha -config fins i tot sense caixeti'
$cfgSense = $afSenseCaix[[Array]::IndexOf($afSenseCaix, '-config') + 1]
Assert ($cfgSense.Contains('adbe.pkcs7.detached')) '_BuildAutoFirmaSignArgv: la compatibilitat hi es sense caixeti'
Assert (-not $cfgSense.Contains('signaturePage')) '_BuildAutoFirmaSignArgv: sense caixeti no s''hi cola cap posicio'
# ...i amb caixeti, les dues coses al mateix -config, amb el separador LITERAL.
$afAmbCaix = @(_BuildAutoFirmaSignArgv 'in.pdf' 'out.pdf' '' '' "Nom`nCarrec" 'text')
$cfgAmb = $afAmbCaix[[Array]::IndexOf($afAmbCaix, '-config') + 1]
Assert ($cfgAmb.Contains('adbe.pkcs7.detached')) '_BuildAutoFirmaSignArgv: compatibilitat + caixeti al mateix -config'
Assert ($cfgAmb.Contains('signaturePage=1')) '_BuildAutoFirmaSignArgv: i la posicio del caixeti tambe'
AssertEq ([bool]($cfgAmb -like "*`n*")) $false '_BuildAutoFirmaSignArgv: cap salt de linia REAL al -config (el separador es el \n LITERAL)'
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
# El marge dret del text NO es un numero congelat: surt de Format.ps1, que es qui
# mana en el format del document. Amb un 552,45 escrit aqui, canviar el marge de
# la plantilla hauria deixat el caixeti despenjat del text sense que ho digues
# ningu -que es exactament el que va passar en passar a 2,5 cm.
$refTextDreta = $Script:A4AmplePt - $Script:ReportFormatConfig.PageMarginRightPt
Assert ([bool]([Math]::Abs([double]$cxP.URY - $refEscutDalt) -le 1.0)) 'AutoFirmaCaixetiPos: el dalt del caixeti va alineat amb la punta de l''escut de la capcalera'
Assert ([bool]([Math]::Abs([double]$cxP.URX - $refTextDreta) -le 1.0)) 'AutoFirmaCaixetiPos: la dreta del caixeti va alineada amb el marge dret del text'

# ELS QUATRE VALORS HAN DE SER ENTERS, i aixo NO es cosmetica: AutoFirma llegeix
# signaturePositionOnPage* com a nombres ENTERS. Amb un "324.48" es queda sense
# la configuracio del caixeti i la signatura surt INVISIBLE, sense dir res ni
# donar cap codi d'error. Va passar de debo en fer que la dreta sortis del marge
# del text: el calcul donava 524,476 i el caixeti va desapareixer.
foreach ($k in @('Page', 'LLX', 'LLY', 'URX', 'URY')) {
    $v = $cxP[$k]
    Assert ([bool]([double]$v -eq [Math]::Floor([double]$v))) ('AutoFirmaCaixetiPos: ' + $k + ' es un ENTER (AutoFirma no accepta decimals)')
}
# I el que se li passa de debo tampoc pot portar cap punt decimal.
foreach ($ln in @(_AutoFirmaPosLines)) {
    $val = ([string]$ln -split '=', 2)[1]
    Assert ([bool]($val -match '^\d+$')) ('AutoFirmaPosLines: "' + $ln + '" ha de ser un enter pelat')
}
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
# Abans aqui s'hi comprovava que sense caixeti NO hi hagues -config. Ja no val:
# les linies de COMPATIBILITAT (SubFilter classic + nomes el certificat del
# signant) hi van sempre, perque son el que fa que la firma es validi als
# ordinadors dels altres, i amb caixeti o sense fa igual.
AssertEq ([bool]($argvSense -contains '-config')) $true '_BuildAutoFirmaSignArgv: sense caixetí, el -config hi es igualment (compatibilitat)'
$valSense = [string]$argvSense[[Array]::IndexOf($argvSense, '-config') + 1]
AssertEq (@($valSense -split '\\n').Count) 3 '_BuildAutoFirmaSignArgv: sense caixetí, nomes les linies de compatibilitat'
AssertEq ([bool]($argvSense -contains 'subject.contains:X')) $true '_BuildAutoFirmaSignArgv: filtre com a element'
$argvCx = @(_BuildAutoFirmaSignArgv 'i.pdf' 'o.pdf' '' '' $cxDef)
$iCfg = [Array]::IndexOf($argvCx, '-config')
AssertEq ([bool]($iCfg -ge 0)) $true '_BuildAutoFirmaSignArgv: amb caixetí -> hi ha -config'
$valCfg = [string]$argvCx[$iCfg + 1]
AssertEq ([bool]($valCfg -like '*signaturePage=1*' -and $valCfg -like '*layer2Text=*')) $true '_BuildAutoFirmaSignArgv: el -config es el TEXT PLA dels extraParams'
AssertEq (@($valCfg -split '\\n').Count) 11 '_BuildAutoFirmaSignArgv: el -config porta 11 propietats (3 de compatibilitat + 8 del caixetí) separades pel \n LITERAL'
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
Assert (-not (_VistaEsProtegit 'C:\x\LLIC.json')) '_VistaEsProtegit: LLIC SI que te vista (_VistaLlicencia: els punts resolts contra REQ1 i el que hi afegeix cada bloc)'

Write-Host "`n--- Activitats.ps1: _ClassificacioText (linia de la capcalera de Llicencia) ---"
AssertEq (_ClassificacioText 'II' '12.25')  ('Llei 20/2009; Annex II; Ep' + [char]0x00ED + 'graf 12.25') '_ClassificacioText: annex + apartat'
AssertEq (_ClassificacioText 'III' '')      'Llei 20/2009; Annex III' '_ClassificacioText: nomes annex'
AssertEq (_ClassificacioText '' '')         '' '_ClassificacioText: sense dades -> buit (no s''inventa res)'
AssertEq (_ClassificacioText '  ' '  ')     '' '_ClassificacioText: nomes espais -> buit'
# Si a l'Excel ja hi diu "Annex II" o "Epigraf x", no s'ha de repetir la paraula.
AssertEq (_ClassificacioText 'Annex II' '12.25') ('Llei 20/2009; Annex II; Ep' + [char]0x00ED + 'graf 12.25') '_ClassificacioText: no repeteix "Annex"'
AssertEq (_ClassificacioText 'II' ('Ep' + [char]0x00ED + 'graf 3.1')) ('Llei 20/2009; Annex II; Ep' + [char]0x00ED + 'graf 3.1') '_ClassificacioText: no repeteix "Epigraf"'

Write-Host "`n--- Document.ps1: blocs de 0 CAPCALERA.docx ---"
AssertEq (_CapMarcador '[[CAP:ACT_EXTR]]') 'ACT_EXTR' '_CapMarcador: reconeix el marcador'
AssertEq (_CapMarcador '  [[CAP:LLIC]]  ') 'LLIC'     '_CapMarcador: amb espais al voltant'
AssertEq (_CapMarcador 'ID GIA: 1234')     ''         '_CapMarcador: un paragraf normal no ho es'
AssertEq (_CapMarcador '')                 ''         '_CapMarcador: buit'
AssertEq (_CapMarcador '[[CAP:]]')         ''         '_CapMarcador: sense nom no val'
# La plantilla ha de portar els tres blocs: generic (REQ1/TERMINI), ACT_EXTR i
# LLIC. Es mira al .docx de debo, que es l'unica plantilla de Word que queda.
$capPath = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
if (Test-Path -LiteralPath $capPath) {
    $capTxt = _ReadDocxPartText $capPath 'word/document.xml'
    # INTEGRITAT del .docx, ABANS de mirar-ne el contingut. Aixo va passar de
    # debo: una cirurgia feta amb un serialitzador d'XML (ElementTree) va
    # reescriure document.xml deixant 3 espais de noms dels 19, inventant-se
    # prefixos 'ns0:' i perdent el mc:Ignorable. Word deia "El archivo parece
    # estar corrompido" i NO ES PODIA GENERAR CAP INFORME -de cap tipus, no
    # nomes Llicencia-, perque tots parteixen d'aquesta capcalera.
    # Regla: aquesta plantilla es toca amb edicions de TEXT sobre el XML, mai
    # reserialitzant-lo.
    Assert ($capTxt.Contains('mc:Ignorable')) '0 CAPCALERA.docx: conserva el mc:Ignorable (si no, Word el dona per corrupte)'
    Assert (-not ($capTxt -match 'ns\d+:')) '0 CAPCALERA.docx: cap prefix d''espai de noms inventat (ns0:, ns1:...)'
    Assert ($capTxt.Contains('standalone="yes"')) '0 CAPCALERA.docx: conserva la declaracio XML original'
    $capNs = ([regex]::Matches($capTxt.Substring(0, [Math]::Min(2000, $capTxt.Length)), 'xmlns:')).Count
    Assert ($capNs -ge 15) ('0 CAPCALERA.docx: hi ha tots els espais de noms (' + $capNs + ', n''hi ha d''haver 19)')
    # Amb .Contains() i NO amb -like: en un patro de -like, '[[CAP:LLIC]' es una
    # CLASSE DE CARACTERS, o sigui que '*[[CAP:LLIC]]*' li dona per bo qualsevol
    # text que porti un dels caracters '[ C A P : L I' seguit d'un ']'. Amb
    # aquest .docx encertava de casualitat (no hi ha cap ']'), pero una prova que
    # nomes funciona per casualitat no protegeix res.
    Assert ($capTxt.Contains('[[CAP:ACT_EXTR]]')) '0 CAPCALERA.docx: hi ha el bloc ACT_EXTR'
    Assert ($capTxt.Contains('[[CAP:LLIC]]'))     '0 CAPCALERA.docx: hi ha el bloc LLIC'
    # ...i NOMES una vegada: si estigues tambe al bloc generic, sortiria als
    # REQ1, on no hi ha de ser.
    AssertEq ([regex]::Matches($capTxt, 'CLASSIFICACIO').Count) 1 '0 CAPCALERA.docx: la classificacio nomes al bloc de Llicencia'

    # A partir d'aqui, sobre el text DESCODIFICAT dels paragrafs: al XML cru el
    # marcador surt escapat ('&lt;&lt;CLASSIFICACIO&gt;&gt;') i buscar-hi
    # '<<CLASSIFICACIO>>' no trobaria mai res.
    $capDoc = _LoadDocxXml $capPath
    $capBloc = ''
    $capLinies = @{ '' = (New-Object System.Collections.ArrayList) }
    foreach ($p in $capDoc.Body.SelectNodes('.//w:p', $capDoc.Ns)) {
        $t = ([string](_ParagraphTextXml $p $capDoc.Ns)).Trim()
        $marca = _CapMarcador $t
        if (-not [string]::IsNullOrWhiteSpace($marca)) {
            $capBloc = $marca
            if (-not $capLinies.ContainsKey($marca)) { $capLinies[$marca] = New-Object System.Collections.ArrayList }
            continue
        }
        if ($t.Length -gt 0) { [void]$capLinies[$capBloc].Add($t) }
    }
    Assert ([bool](@($capLinies['LLIC']) | Where-Object { $_.Contains('<<CLASSIFICACIO>>') })) '0 CAPCALERA.docx: el bloc LLIC porta el marcador de la classificacio'
    Assert (-not (@($capLinies['']) | Where-Object { $_ -like 'Classificaci*' })) '0 CAPCALERA.docx: el bloc generic NO porta la classificacio'

    # Cap ETIQUETA SENSE MARCADOR. Una linia com "Classificacio:" sense cap
    # <<...>> al darrere nomes pot sortir BUIDA a l'informe, i es exactament el
    # que hi havia al bloc generic: sortia una "Classificacio:" en blanc a tots
    # els REQ1 i TERMINI. Aixo ho enxampa vingui d'on vingui.
    $capBuides = New-Object System.Collections.ArrayList
    foreach ($bloc in $capLinies.Keys) {
        foreach ($t in @($capLinies[$bloc])) {
            if (([string]$t).Length -ge 2 -and ([string]$t).EndsWith(':')) { [void]$capBuides.Add($t) }
        }
    }
    AssertEq $capBuides.Count 0 ('0 CAPCALERA.docx: cap etiqueta sense marcador (' + ($capBuides -join ' | ') + ')')
    # LA CLASSIFICACIO NO VA EN NEGRETA. A la resta de linies de la capcalera
    # l'etiqueta es en negreta i el VALOR no; el bloc [[CAP:LLIC]] es va escriure
    # amb tot en un sol run bold, i "Llei 20/2009; Annex III..." sortia en
    # negreta a l'informe. El valor ha de tenir el seu propi run, sense <w:b/>.
    $paraCl = @([regex]::Matches($capTxt, '<w:p[ >].*?</w:p>', 'Singleline') |
                Where-Object { $_.Value -match 'Classificaci' })
    Assert ($paraCl.Count -ge 1) '0 CAPCALERA.docx: hi ha el paragraf de la classificacio'
    $runCl = @([regex]::Matches($paraCl[0].Value, '<w:r[ >].*?</w:r>', 'Singleline') |
               Where-Object { $_.Value -match 'CLASSIFICACIO' })
    AssertEq $runCl.Count 1 '0 CAPCALERA.docx: el marcador de la classificacio va en un run propi'
    Assert (-not ($runCl[0].Value -match '<w:b/>')) '0 CAPCALERA.docx: ...i aquell run NO va en negreta'
    Assert ([bool]($paraCl[0].Value -match 'Classificaci[^<]*</w:t>')) '0 CAPCALERA.docx: l''etiqueta "Classificacio:" va a part'
}

# "Planols" sortia partit en TRES caselles (Pl / a / nols): dins d'un @(...) la
# coma lliga MES FORT que el '+'. Ja esta avisat a CLAUDE.md i hi vam caure
# igualment; per aixo la llista es una funcio PURA i es pot COMPTAR.
# LA MATEIXA TRAMPA, EN ARGUMENTS D'UNA CRIDA, i a TOT el codi de suport/.
#
#   _AddBrandHeader $form 'X' 'refer' + [char]0x00E8 + 'ncia' 56
#
# PowerShell NO concatena: passa 'refer', '+', 'e', '+', 'ncia' i 56 com a
# arguments SOLTS, o sigui que el 56 (l'alcada) acaba en un altre parametre i
# la crida peta amb "no se puede convertir el valor '+' al tipo System.Int32".
# Va passar de debo al boto "Omplir..." de Llicencia.
#
# La deteccio la fa EL PROPI PARSER i no te falsos positius: un argument que
# sigui LITERALMENT '+' nomes pot venir d'una concatenacio sense parentesis.
$plusSolts = New-Object System.Collections.ArrayList
Get-ChildItem -Recurse -Filter *.ps1 (Split-Path -Parent $TestsDir) | ForEach-Object {
    $fitxer = $_
    $arbre = [System.Management.Automation.Language.Parser]::ParseFile($fitxer.FullName, [ref]$null, [ref]$null)
    foreach ($cmd in $arbre.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        foreach ($el in $cmd.CommandElements) {
            if ($el -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $el.Value -eq '+') {
                [void]$plusSolts.Add(($fitxer.Name + ':' + $el.Extent.StartLineNumber))
                break
            }
        }
    }
}
AssertEq $plusSolts.Count 0 ("cap concatenacio sense parentesis en arguments d'una crida (" + ($plusSolts -join ', ') + ')')

# UNA CLOSURE QUE ES CRIDA A SI MATEIXA, i a TOT el codi de suport/.
#
#   $pinta = { ... & $pinta $idx ... }.GetNewClosure()
#
# .GetNewClosure() copia el VALOR de les variables EN EL MOMENT de crear el
# scriptblock, i en aquell moment $pinta encara val $null. En cridar-la peta
# amb "l'expressio que segueix a & ... no es un nom d'ordre ni un scriptblock"
# (un dialeg gris de .NET, sense fitxer ni linia). Va passar de debo a la
# pantalla de documentacio de Llicencia, en triar "Es disposa del document".
#
# La variant SILENCIOSA es encara pitjor: $cerca = $null; ... { $cerca.Text
# }.GetNewClosure() ... $cerca = _AddSearchBox ... -> la closure es queda amb
# $null, $null.Text no peta, i el cercador simplement NO FILTRA MAI.
#
# SOLUCIO: un hashtable de funcions ($fn = @{}; $fn.Pinta = { ... & $fn.Pinta
# ... }.GetNewClosure()), que es captura per REFERENCIA i es resol en cridar-lo.
# Un scriptblock SENSE .GetNewClosure() no te el problema (resol en temps
# d'execucio), o sigui que nomes es miren els que la fan servir.
$closuresRecursives = New-Object System.Collections.ArrayList
Get-ChildItem -Recurse -Filter *.ps1 (Split-Path -Parent $TestsDir) | ForEach-Object {
    $fitxer = $_
    $arbre = [System.Management.Automation.Language.Parser]::ParseFile($fitxer.FullName, [ref]$null, [ref]$null)
    foreach ($asg in $arbre.FindAll({ param($x) $x -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
        $esq = $asg.Left -as [System.Management.Automation.Language.VariableExpressionAst]
        if ($null -eq $esq) { continue }
        $nom = $esq.VariablePath.UserPath
        $tancades = $asg.Right.FindAll({ param($x)
            $x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            ([string]$x.Member.Value -eq 'GetNewClosure') }, $true)
        foreach ($c in $tancades) {
            $trobat = $false
            foreach ($v in $c.Expression.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                if ($v.VariablePath.UserPath -eq $nom) { $trobat = $true; break }
            }
            if ($trobat) {
                [void]$closuresRecursives.Add(($fitxer.Name + ':' + $asg.Extent.StartLineNumber + ' ($' + $nom + ')'))
                break
            }
        }
    }
}
AssertEq $closuresRecursives.Count 0 ("cap .GetNewClosure() es refereix a la variable que s'hi assigna (" + ($closuresRecursives -join ', ') + ')')

# LA SEGONA CARA DE LA MATEIXA TRAMPA: UNA CLOSURE DINS D'UNA ALTRA.
#
# .GetNewClosure() copia NOMES els LOCALS del context que la crida. Una closure
# creada A DINS d'una altra es queda, doncs, amb els locals d'aquella invocacio
# -parametres i variables que s'hi assignen- i PERD tot el que venia del modul
# de la closure de fora:
#
#   $fn = @{}; $fora = 'x'
#   $fn.Pinta = { param($idx)
#       $local = 'l'
#       $inner = { "$local $idx $fora $fn" }.GetNewClosure()   # <- $fora i $fn: NULL
#   }.GetNewClosure()
#
# Va passar al handler dels radios de la pantalla de documentacio: $fn hi
# arribava buit i "& $fn.Pinta" petava en triar "Es disposa del document".
# El detector de mes amunt (una closure que es refereix a SI MATEIXA) no el veu.
#
# S'exclouen les automatiques i les de $Script:/$Global:/$env:, que si que es
# resolen en temps d'execucio.
$Global:_autoTancades = @('_','null','true','false','args','this','PSItem','input','error',
                          'matches','host','PSScriptRoot','PSCommandPath','MyInvocation',
                          'LASTEXITCODE','sender','e','ev','s','o')
function _LocalsDeScriptBlock($sb) {
    $set = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    if ($sb.ParamBlock) {
        foreach ($p in $sb.ParamBlock.Parameters) { [void]$set.Add($p.Name.VariablePath.UserPath) }
    }
    foreach ($a in $sb.FindAll({ param($x) $x -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
        $l = $a.Left -as [System.Management.Automation.Language.VariableExpressionAst]
        if ($l) { [void]$set.Add($l.VariablePath.UserPath) }
    }
    foreach ($f in $sb.FindAll({ param($x) $x -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)) {
        [void]$set.Add($f.Variable.VariablePath.UserPath)
    }
    # Amb la coma: si no, PowerShell ENUMERA el HashSet i el retorn es una
    # cadena (o $null si es buit), i despres .Contains() peta.
    return ,$set
}
$closuresImbricades = New-Object System.Collections.ArrayList
Get-ChildItem -Recurse -Filter *.ps1 (Split-Path -Parent $TestsDir) | ForEach-Object {
    $fitxer = $_
    $arbre = [System.Management.Automation.Language.Parser]::ParseFile($fitxer.FullName, [ref]$null, [ref]$null)
    $esTancada = { param($x)
        $x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
        ([string]$x.Member.Value -eq 'GetNewClosure') }
    foreach ($fora in $arbre.FindAll($esTancada, $true)) {
        $sbFora = $fora.Expression -as [System.Management.Automation.Language.ScriptBlockExpressionAst]
        if ($null -eq $sbFora) { continue }
        $localsFora = _LocalsDeScriptBlock $sbFora.ScriptBlock
        foreach ($dins in $sbFora.ScriptBlock.FindAll($esTancada, $true)) {
            $sbDins = $dins.Expression -as [System.Management.Automation.Language.ScriptBlockExpressionAst]
            if ($null -eq $sbDins) { continue }
            $localsDins = _LocalsDeScriptBlock $sbDins.ScriptBlock
            $perdudes = @()
            foreach ($v in $sbDins.ScriptBlock.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                $n = $v.VariablePath.UserPath
                if ($v.VariablePath.IsGlobal -or $v.VariablePath.IsScript -or $n -like 'env:*') { continue }
                if ($Global:_autoTancades -contains $n) { continue }
                if ($localsDins.Contains($n) -or $localsFora.Contains($n)) { continue }
                if ($perdudes -notcontains $n) { $perdudes += $n }
            }
            if ($perdudes.Count -gt 0) {
                [void]$closuresImbricades.Add(($fitxer.Name + ':' + $dins.Extent.StartLineNumber +
                                               ' ($' + ($perdudes -join ', $') + ')'))
            }
        }
    }
}
AssertEq $closuresImbricades.Count 0 ('cap closure imbricada depen del que ve de fora (' + ($closuresImbricades -join ', ') + ')')

# UN EMOJI ASTRAL AMB [char]: EL PROGRAMA NO S'OBRE.
#
#   Icon = [string][char]0x1F5C2 + [char]0xFE0F
#
# [char] es de 16 bits (max 0xFFFF) i 0x1F5C2 (128450) NO hi cap: PowerShell
# peta amb "no se puede convertir el valor 128450 al tipo System.Char". Va
# passar al xip "Dades" del menu, i com que aquell codi construeix la PRIMERA
# pantalla, el programa no arrencava gens.
#
# La manera bona ja era al fitxer dues linies mes amunt:
#   [System.Char]::ConvertFromUtf32(0x1F5C2)
#
# PER QUE CAL LA PROVA: Select-Mode es WinForms i les proves no el criden mai,
# o sigui que aixo no ho enxampa cap prova de comportament. El parser, en canvi,
# ho veu sense executar res i sense falsos positius: un [char] amb una constant
# mes gran que 0xFFFF no pot ser correcte MAI.
$charsForaDeRang = New-Object System.Collections.ArrayList
Get-ChildItem -Recurse -Filter *.ps1 (Split-Path -Parent $TestsDir) | ForEach-Object {
    $fitxer = $_
    $arbre = [System.Management.Automation.Language.Parser]::ParseFile($fitxer.FullName, [ref]$null, [ref]$null)
    foreach ($cast in $arbre.FindAll({ param($x) $x -is [System.Management.Automation.Language.ConvertExpressionAst] }, $true)) {
        $tipus = [string]$cast.Type.TypeName.FullName
        if ($tipus -notin @('char', 'System.Char')) { continue }
        $const = $cast.Child -as [System.Management.Automation.Language.ConstantExpressionAst]
        if ($null -eq $const) { continue }
        $n = 0
        if (-not [int]::TryParse([string]$const.Value, [ref]$n)) { continue }
        if ($n -gt 0xFFFF) {
            [void]$charsForaDeRang.Add(($fitxer.Name + ':' + $cast.Extent.StartLineNumber + ' ' + $cast.Extent.Text))
        }
    }
}
AssertEq $charsForaDeRang.Count 0 ("cap [char] amb un codi mes gran que 0xFFFF -fes servir [System.Char]::ConvertFromUtf32- (" + ($charsForaDeRang -join ', ') + ')')

# EL DESPATXADOR DEL MENU HA D'APUNTAR A FUNCIONS QUE EXISTEIXIN.
#
# El 'switch' de Main (Wizard.ps1) reparteix cada rajola del menu cap a la seva
# funcio. Si algu la reanomena -o l'afegeix al switch abans d'escriure-la-, el
# programa arrenca igual i nomes peta EN CLICAR aquella rajola. Cap prova de
# comportament ho veu: totes aquestes funcions obren finestres.
#
# Aqui es treuen del PROPI switch (AST) els noms d'ordre que invoca i es
# comprova que tots es resolen amb el motor ja carregat.
$accionsSenseFuncio = New-Object System.Collections.ArrayList
$wizardPath = Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1'
if (Test-Path -LiteralPath $wizardPath) {
    $astW = [System.Management.Automation.Language.Parser]::ParseFile($wizardPath, [ref]$null, [ref]$null)
    $mainFn = $astW.Find({ param($x)
        $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'Main' }, $true)
    Assert ($null -ne $mainFn) 'Wizard.ps1: hi ha la funcio Main'
    $sw = $mainFn.Find({ param($x) $x -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
    Assert ($null -ne $sw) 'Main: hi ha el switch del despatxador'
    foreach ($cmd in $sw.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $nom = [string]$cmd.GetCommandName()
        if ([string]::IsNullOrWhiteSpace($nom)) { continue }
        # Nomes ens interessen les funcions del programa (Verb-Nom).
        if ($nom -notmatch '^[A-Za-z]+-[A-Za-z]') { continue }
        if (-not (Get-Command $nom -ErrorAction SilentlyContinue)) {
            [void]$accionsSenseFuncio.Add($nom)
        }
    }
}
AssertEq $accionsSenseFuncio.Count 0 ('el despatxador del menu apunta a funcions que no existeixen (' + ($accionsSenseFuncio -join ', ') + ')')

$docsLl = @(_LlicDocsSignats)
AssertEq $docsLl.Count 3 '_LlicDocsSignats: TRES documents (no cinc: la coma dins d''un @() parteix el text)'
AssertEq $docsLl[1] ('Pl' + [char]0x00E0 + 'nols') '_LlicDocsSignats: Planols, sencer'
# I cap element pot ser un tros solt d'una paraula.
Assert (-not (@($docsLl) | Where-Object { $_.Length -le 2 })) '_LlicDocsSignats: cap element trencat'

# Les DADES de cada document ("Es disposa (Id Firmadoc: ...; Expedient: ...)").
# Els XXX del Word de l'usuari son forats: al cataleg han de ser [CAMP: ...].
$txtSi = @('Es disposa de l''informe (Id Firmadoc: [CAMP: Id Firmadoc]; Expedient: [CAMP: Expedient])')
$campsSi = @(_LlicCampsDelText $txtSi)
AssertEq $campsSi.Count 2 '_LlicCampsDelText: troba els dos camps'
AssertEq $campsSi[0] 'Id Firmadoc' '_LlicCampsDelText: el primer'
AssertEq $campsSi[1] 'Expedient'   '_LlicCampsDelText: el segon'
AssertEq (@(_LlicCampsDelText @('sense cap camp')).Count) 0 '_LlicCampsDelText: sense camps, cap'
AssertEq (@(_LlicCampsDelText @()).Count) 0 '_LlicCampsDelText: sense text, cap'
# El mateix camp repetit compta un cop (surt una sola casella al dialeg).
AssertEq (@(_LlicCampsDelText @('[CAMP: A] i [CAMP: A]')).Count) 1 '_LlicCampsDelText: el mateix camp, una vegada'
$resolt = @(_LlicAplicaCamps $txtSi @{ 'Id Firmadoc' = '2024/123'; 'Expedient' = 'E-9' })
Assert ($resolt[0].Contains('Id Firmadoc: 2024/123')) '_LlicAplicaCamps: hi posa el valor'
Assert ($resolt[0].Contains('Expedient: E-9')) '_LlicAplicaCamps: i el segon'
Assert (-not ($resolt[0] -match '\[CAMP:')) '_LlicAplicaCamps: no queda cap marcador al text'
# Un camp sense valor no deixa el marcador a la vista: queda buit.
$buit = @(_LlicAplicaCamps $txtSi @{})
Assert (-not ($buit[0] -match '\[CAMP:')) '_LlicAplicaCamps: sense valor, el marcador desapareix igualment'
# La clau amb que es recorda cada punt en tornar ENRERE.
AssertEq (_LlicClauPunt ([pscustomobject]@{ Clau='S::T'; Titol='X' })) 'S::T' '_LlicClauPunt: la clau de REQ1 si en te'
AssertEq (_LlicClauPunt ([pscustomobject]@{ Clau=''; Titol='Propi' })) '#Propi' '_LlicClauPunt: i el titol si es un punt propi'

# Al CATALEG real: cap "Es disposa" pot quedar amb XXX (seria un forat que
# l'usuari no pot omplir des del programa).
$llicPathX = Join-Path $EstructuralsDir 'LLIC.json'
if (Test-Path -LiteralPath $llicPathX) {
    $llicTxt = Get-Content -LiteralPath $llicPathX -Raw -Encoding UTF8
    Assert (-not $llicTxt.Contains('XXX')) 'LLIC.json: cap XXX (els forats son camps [CAMP: ...])'
    # ...i cada "Es disposa" ha de demanar com a minim l'Id Firmadoc.
    $llicJ = $llicTxt | ConvertFrom-Json
    $senseId = New-Object System.Collections.ArrayList
    foreach ($sx in @($llicJ.nodes)) {
        foreach ($itx in @($sx.fills)) {
            foreach ($fx in @($itx.fills)) {
                if ([string]$fx.tipus -ne 'sidisposa') { continue }
                $tx = -join (@($fx.cos) | ForEach-Object { @($_.runs) | ForEach-Object { [string]$_.t } })
                if ($tx -notmatch '\[CAMP:') { [void]$senseId.Add([string]$itx.titol) }
            }
        }
    }
    AssertEq $senseId.Count 0 ("LLIC.json: tot 'Es disposa' demana dades (" + ($senseId -join ', ') + ')')
}

# La pantalla de tria de documentacio de LLICENCIA va quedar MORTA: cap boto
# responia. La causa: _StyleListGrid fa Dock='Fill' (pensat per a graelles dins
# d'un panell) i la graella, posada directament al formulari, ocupava TOTA la
# finestra i tapava els botons. Prova de FONT: despres de l'estil, el Dock s'ha
# de desfer i la posicio s'ha de fixar DESPRES (abans, l'estil la trepitjava).
$srcLlic = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Llicencia.ps1') -Raw
# La pantalla de documentacio ja NO es una graella: es llista + detall, com
# Select-Items, per poder omplir els camps INLINE (i per aixo tampoc pot
# tornar a passar que el Dock='Fill' de _StyleListGrid tapi els botons).
Assert (-not $srcLlic.Contains('_StyleListGrid')) 'Llicencia: la pantalla ja no es una graella (llista + detall, com REQ1)'
Assert ($srcLlic.Contains('_RenderRichInto')) 'Llicencia: els camps es pinten amb la MATEIXA funcio que REQ1'
Assert (-not $srcLlic.Contains('DataGridViewButtonColumn')) 'Llicencia: fora el boto "Omplir..."'
# I el Word, DIFERIT: al pas 3 nomes cal el JSON; si s'arrenca alla, la tria de
# punts triga i un usuari que tira enrere deixa un Word obert per res.
# Ara hi ha DOS camins que generen (l'informe llarg i els dos curts), o sigui
# que la prova mira que CADA arrencada del Word vagi seguida d'un Build-...
$iPas3 = $srcLlic.IndexOf('# Aqui nomes cal REQ1')
Assert ($iPas3 -ge 0) 'Llicencia: el pas 3 nomes llegeix el JSON'
$nWord = 0
$posW = $srcLlic.IndexOf('$word = New-WordApp')
while ($posW -ge 0) {
    $nWord++
    $tros = $srcLlic.Substring($posW, [Math]::Min(1200, $srcLlic.Length - $posW))
    Assert ($tros.Contains('Build-')) 'Llicencia: el Word s''arrenca DIFERIT (al pas de generar, no abans)'
    $posW = $srcLlic.IndexOf('$word = New-WordApp', $posW + 1)
}
Assert ($nWord -ge 2) 'Llicencia: els dos camins de generacio arrenquen el Word'
