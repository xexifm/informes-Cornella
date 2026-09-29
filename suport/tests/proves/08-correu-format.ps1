# El FORMAT del correu (CorreuFormat.ps1 i docs\correu.js)
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.
#
# Peticio de l'usuari (setembre 2026): el correu del PC i el del mobil han de
# tenir el format de l'informe de REQ1 i del seguiment, i han de coincidir. El
# que ho garanteix, i el que es prova aqui:
#   - les mides surten de $ReportFormatConfig i el mobil llegeix les MATEIXES
#     (docs\dades\correu-format.json, que ha d'estar al dia);
#   - la capcalera son les linies de '0 CAPCALERA', les mateixes als dos;
#   - el mobil (JavaScript) dona el MATEIX HTML que el PC per a la mateixa
#     seleccio de REQ1 (prova creuada amb Node, mes avall);
#   - el PC llegeix el .docx sense Word i NO hi repeteix la capcalera.

$cfRoot = Split-Path -Parent (Split-Path -Parent $TestsDir)
$cfDades = Join-Path $cfRoot (Join-Path 'docs' 'dades')
$cfFmt = _CorreuFormat
$cfCfg = $Script:ReportFormatConfig

Write-Host "`n--- CorreuFormat.ps1: les mides surten de Format.ps1 ---"
AssertEq $cfFmt.Blocs.pic.Esq (_CmAPx $cfCfg.BulletIndentCm) '_CorreuFormat: el pic, a la sagnia del Word'
AssertEq $cfFmt.Blocs.pic.Penjat (_CmAPx $cfCfg.BulletHangCm) '_CorreuFormat: el pic, amb la sagnia francesa del Word'
AssertEq ([double]$cfFmt.Blocs.picPrimer.Abans) ([double]$cfCfg.PrimerSubpuntSpaceBeforePt) '_CorreuFormat: el PRIMER sub-punt, separat com al Word'
AssertEq ([double]$cfFmt.Blocs.pic.Abans) ([double]$cfCfg.BulletSpaceBeforePt) '_CorreuFormat: els altres sub-punts, amb l''espai curt'
AssertEq $cfFmt.Blocs.cosFill.Esq (_CmAPx $cfCfg.ChildIndentCm) '_CorreuFormat: les linies d''un fill, sagnades com al Word'
AssertEq $cfFmt.Blocs.enllac.Esq (_CmAPx $cfCfg.ItemIndentCm) '_CorreuFormat: l''enllac d''un punt NO va sagnat (com al Word)'
AssertEq ([double]$cfFmt.Blocs.enllac.MidaPt) ([double]$cfCfg.UrlFontSize) '_CorreuFormat: l''enllac, a la mida del Word'
AssertEq $cfFmt.Blocs.conclusiocap.Alinea 'center' '_CorreuFormat: CONCLUSIONS centrat'
AssertEq ([double]$cfFmt.Blocs.conclusio.Despres) ([double]$cfCfg.ConclusionSpaceAfterPt) '_CorreuFormat: l''espai de sota d''una conclusio'
AssertEq ([bool]$cfFmt.Aire['item']) ([bool]$cfCfg.SpacerAfterItem) '_CorreuFormat: l''aire despres de cada punt, de la bandera'
AssertEq $cfFmt.Lletra 'Calibri, Arial, sans-serif' '_CorreuFormat: la lletra del correu es Calibri (decisio de l''usuari)'
AssertEq (_CorreuNum 6.5) '6.5' '_CorreuNum: punt decimal sempre (en catala seria "6,5" i el CSS el descartaria)'

# El fitxer que llegeix el mobil ha de ser el que surt d'aqui. Si algu canvia
# una sagnia a Format.ps1 i no regenera les dades, el mobil divergiria.
$cfPub = Read-JsonFile (Join-Path $cfDades 'correu-format.json')
Assert ($null -ne $cfPub) 'docs/dades/correu-format.json existeix'
# Es compara el CONTINGUT, no com s'escriu el numero: el PowerShell 5.1 de
# l'usuari desa un 11 on el 7 d'aqui escriu 11.0 (Actualitzar.bat el regenera).
$cfNum = { param($j) [regex]::Replace($j, '(?<=[:\[,])(-?\d+)\.0(?=[,}\]])', '$1') }
$cfA = & $cfNum ($cfPub | ConvertTo-Json -Depth 6 -Compress)
$cfB = & $cfNum (($cfFmt | ConvertTo-Json -Depth 6) | ConvertFrom-Json | ConvertTo-Json -Depth 6 -Compress)
AssertEq $cfA $cfB 'docs/dades/correu-format.json esta al dia amb _CorreuFormat (si no: ExportaDades -Plantilles)'

Write-Host "`n--- CorreuFormat.ps1: la capcalera del correu es la de '0 CAPCALERA' ---"
$cfLin = @(_CorreuCapcaleraLinies (Read-JsonFile (Get-CapcaleraJsonPath)))
AssertEq (($cfLin | ForEach-Object { $_.Plantilla }) -join ',') '<<ID_GIA>>,<<EXP_NUM>>,<<ADRECA>>,<<ACTIVITAT>>,<<TITULAR>>,<<ORIGEN>>' '_CorreuCapcaleraLinies: les sis linies, en l''ordre de la capcalera'
AssertEq $cfLin[0].Etiqueta 'ID GIA:' '_CorreuCapcaleraLinies: l''etiqueta tal com es a la capcalera'
$cfCapPub = Read-JsonFile (Join-Path $cfDades 'capcalera.json')
AssertEq ((@($cfCapPub.Correu) | ForEach-Object { $_.Etiqueta + $_.Plantilla }) -join '|') (($cfLin | ForEach-Object { $_.Etiqueta + $_.Plantilla }) -join '|') 'capcalera.json: el mobil te les MATEIXES linies que el PC'
AssertEq ([string]$cfCapPub.Origen.doc) ([string]$Script:OrigenPlantilles['doc']) 'capcalera.json: la plantilla de l''Objecte (doc) es la del PC'
AssertEq ([string]$cfCapPub.Origen.insp) ([string]$Script:OrigenPlantilles['insp']) 'capcalera.json: la plantilla de l''Objecte (visita) es la del PC'
# L'amplada de l'etiqueta es la sagnia francesa de la plantilla.
$cfCapXml = _ReadDocxPartText (Get-CapcaleraDocxPath) 'word/document.xml'
$cfMid = [regex]::Match($cfCapXml, '<w:p[ >](?:(?!</w:p>).)*?&lt;&lt;ID_GIA&gt;&gt;', 'Singleline')
AssertEq ([regex]::Match($cfMid.Value, 'w:hanging="(\d+)"').Groups[1].Value) ([string]$Script:CorreuCapcaleraTwips) 'la columna de l''etiqueta del correu = la sagnia francesa de 0 CAPCALERA.docx'
AssertEq (_CorreuCapcaleraTitol (Read-JsonFile (Get-CapcaleraJsonPath))) 'INFORME' '_CorreuCapcaleraTitol: el titol de sota la capcalera, de 0 CAPCALERA'
AssertEq ([string]$cfCapPub.Titol) 'INFORME' 'capcalera.json: el mobil te el mateix titol'
$cfTitHtml = _CorreuCapcaleraHtml @([pscustomobject]@{ Etiqueta = 'ID GIA:'; Valor = '1398' }) $cfFmt 'INFORME'
Assert ($cfTitHtml.Contains('text-align:center"><b>INFORME</b></p>')) '_CorreuCapcaleraHtml: INFORME centrat i en negreta (com a l''informe)'
Assert ($cfTitHtml.IndexOf('1398') -lt $cfTitHtml.IndexOf('INFORME')) '_CorreuCapcaleraHtml: INFORME va DESPRES de les linies'
$cfCapHtml = _CorreuCapcaleraHtml @([pscustomobject]@{ Etiqueta = 'ID GIA:'; Valor = '1398' }, [pscustomobject]@{ Etiqueta = 'Objecte:'; Valor = '' }) $cfFmt
Assert ($cfCapHtml.Contains('<b>ID GIA:</b>') -and $cfCapHtml.Contains('1398')) '_CorreuCapcaleraHtml: etiqueta en negreta i valor'
Assert (-not $cfCapHtml.Contains('Objecte')) '_CorreuCapcaleraHtml: una linia sense valor no surt'

Write-Host "`n--- CorreuFormat.ps1: l'Objecte al reves i la frase d'introduccio ---"
$cfDoc = _OrigenDesDeText (_BuildOrigenText @{ ORIGEN_TIPUS = 'doc'; NUM_ANOTACIO = '21327'; DATA_ANOTACIO = '30/04/2026' })
AssertEq "$($cfDoc.ORIGEN_TIPUS)|$($cfDoc.NUM_ANOTACIO)|$($cfDoc.DATA_ANOTACIO)" 'doc|21327|30/04/2026' '_OrigenDesDeText: llegeix el que escriu _BuildOrigenText (doc)'
$cfIns = _OrigenDesDeText (_BuildOrigenText @{ ORIGEN_TIPUS = 'insp'; DATA_INSPECCIO = '29/09/2026' })
AssertEq "$($cfIns.ORIGEN_TIPUS)|$($cfIns.DATA_INSPECCIO)" 'insp|29/09/2026' '_OrigenDesDeText: llegeix el que escriu _BuildOrigenText (visita)'
$cfApo = _OrigenDesDeText ("Doc. aportada amb N" + [char]0x00FA + "m. d'anotaci" + [char]0x00F3 + " 43294 del 07/09/2026")
AssertEq $cfApo.NUM_ANOTACIO '43294' '_OrigenDesDeText: tambe amb l''apostrof recte (el Word el pot canviar)'
AssertEq (_OrigenDesDeText 'Qualsevol altra cosa').ORIGEN_TIPUS '' '_OrigenDesDeText: un text que no encaixa no s''inventa l''origen'
$cfTx = @{ introDoc = 'D {NUM_ANOTACIO} {DATA_ANOTACIO}'; introDocSenseAnotacio = 'DS'; introInsp = 'V {DATA_INSPECCIO}' }
AssertEq (_CorreuIntro $cfTx $cfDoc '01/01/2027') 'D 21327 30/04/2026' '_CorreuIntro: documentacio aportada, amb num. i data d''anotacio'
AssertEq (_CorreuIntro $cfTx @{ ORIGEN_TIPUS = 'doc'; NUM_ANOTACIO = '' } '01/01/2027') 'DS' '_CorreuIntro: documentacio aportada sense anotacio'
AssertEq (_CorreuIntro $cfTx @{ ORIGEN_TIPUS = '' } '01/01/2027') 'DS' '_CorreuIntro: sense Objecte -> documentacio aportada (MAI "la visita")'
AssertEq (_CorreuIntro $cfTx $cfIns '01/01/2027') 'V 29/09/2026' '_CorreuIntro: visita amb la seva data'
AssertEq (_CorreuIntro $cfTx @{ ORIGEN_TIPUS = 'insp'; DATA_INSPECCIO = '' } '01/01/2027') 'V 01/01/2027' '_CorreuIntro: visita sense data -> la d''avui (peticio de l''usuari)'
AssertEq (_TextToHtml 'a !!vermell!! b') 'a <span style="color:#C00000">vermell</span> b' '_TextToHtml: !!vermell!!'
$cfTxReal = _LoadEmailTextos
foreach ($k in @('introDoc', 'introDocSenseAnotacio', 'introInsp')) { Assert (-not [string]::IsNullOrWhiteSpace([string]$cfTxReal[$k])) "email-textos.json porta $k" }
Assert (-not ([string]$cfTxReal['introDoc']).Contains('visita')) 'email-textos.json: la frase de documentacio aportada no parla de la visita'
foreach ($v in @('{CAPCALERA}', '{INTRO}', '{REQUERIMENTS}')) { Assert (([string]$cfTxReal['cos']).Contains($v)) "email-textos.json: el cos porta $v" }
# El peu (peticio de l'usuari, setembre 2026): "Com presentar la documentacio /
# Como presentar la documentacion" en UNA linia i en vermell; a sota, el bloc
# CATALA i el bloc ESPANOL, cadascun amb tota la seva informacio (abans anaven
# barrejats), i res en negreta.
$cfPeu = [string]$cfTxReal['cos']
$cfPeu = $cfPeu.Substring($cfPeu.IndexOf('{REQUERIMENTS}'))
Assert ($cfPeu.Contains('!!Com presentar la documentaci' + [char]0x00F3 + ' / C' + [char]0x00F3 + 'mo presentar la documentaci' + [char]0x00F3 + 'n!!')) 'email-textos.json: "Com presentar / Como presentar" en una linia i en vermell'
$cfIxCa = $cfPeu.IndexOf("`nCATAL"); $cfIxEs = $cfPeu.IndexOf("`nESPA")
Assert ($cfIxCa -gt 0 -and $cfIxEs -gt $cfIxCa) 'email-textos.json: el bloc CATALA i despres el ESPANOL'
$cfCa = $cfPeu.Substring($cfIxCa, $cfIxEs - $cfIxCa); $cfEs = $cfPeu.Substring($cfIxEs)
Assert ($cfCa.Contains('idioma=2') -and $cfCa.Contains('IMPORTANT:') -and -not $cfCa.Contains('IMPORTANTE')) 'email-textos.json: el bloc catala ho porta tot en catala'
Assert ($cfEs.Contains('idioma=1') -and $cfEs.Contains('IMPORTANTE:') -and -not $cfEs.Contains('Heu de presentar')) 'email-textos.json: el bloc espanyol ho porta tot en castella'
AssertEq ([regex]::Matches($cfPeu, '\*\*').Count) 0 'email-textos.json: al peu, cap negreta'

# Dins d'un try: una excepcio aqui mataria la resta de la suite i el resum
# seguiria dient "0 FAIL" (ja ha passat dues vegades en aquest projecte).
try {
Write-Host "`n--- CorreuFormat.ps1: blocs -> HTML (el mateix bucle que Write-Informe) ---"
$cfBl = @(
    @{ T = 'seccio'; Text = 'Instal·lacions' }, @{ T = 'aire'; Clau = 'seccio' },
    @{ T = 'unitat'; Blocs = @(
        @{ T = 'item'; Num = '1.'; Text = 'Baixa **tensio**' },
        @{ T = 'enllac'; Url = 'https://a.cat/x' },
        @{ T = 'pic'; Text = 'Primer'; Fill = $true },
        @{ T = 'pic'; Text = 'Segon'; Fill = $true }) },
    @{ T = 'conclusiocap'; Text = 'CONCLUSIONS' })
$cfH = _CorreuBlocsAHtml $cfBl $cfFmt
Assert ($cfH.Contains(('INSTAL' + [char]0x00B7 + 'LACIONS'))) 'blocs->HTML: la seccio en MAJUSCULES, com Format-Section'
Assert (-not ($cfH -match '<b>INSTAL')) 'blocs->HTML: la seccio SENSE negreta, com a l''informe'
Assert ($cfH.Contains('<b>1.</b> Baixa <b>tensio</b>')) 'blocs->HTML: numero en negreta i **negreta** del cataleg'
Assert ($cfH.Contains('style="font-size:10pt;word-break:break-all">https://a.cat/x</a>')) 'blocs->HTML: l''enllac, a 10 pt'
# El pic va en una fila de taula (l'Outlook no fa cas de la sagnia francesa):
# el marge es la sagnia menys la columna del pic (38 - 19 px).
Assert ($cfH.Contains('margin:12pt 0 0pt 19px') -and $cfH.Contains('margin:6pt 0 0pt 19px')) 'blocs->HTML: el PRIMER sub-punt a 12 pt i el segon a 6 pt'
AssertEq ([regex]::Matches($cfH, '&nbsp;</p>').Count) 2 'blocs->HTML: una linia en blanc despres de la seccio i una despres del punt SENCER'
$cfPeta = $false; try { [void](_CorreuBlocsAHtml @(@{ T = 'inventat' }) $cfFmt) } catch { $cfPeta = $true }
Assert $cfPeta 'blocs->HTML: un tipus de bloc desconegut PETA (com Write-Informe)'

Write-Host "`n--- CorreuFormat.ps1: el .docx de l'informe, sense Word ---"
# Un document.xml com el que escriu el Word: la capcalera amb tabulador,
# INFORME, la nota, la frase del cataleg, una seccio, un punt amb el numero en
# negreta, l'enllac a 10 pt, un sub-punt amb pic i sagnia francesa, una
# anotacio de seguiment amb el comentari en negreta, les conclusions i el
# tancament.
$cfW = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
$cfP = { param($ppr, $runs) "<w:p><w:pPr>$ppr</w:pPr>$runs</w:p>" }
$cfR = { param($t, $rpr) "<w:r><w:rPr>$rpr</w:rPr><w:t xml:space=`"preserve`">$t</w:t></w:r>" }
$cfTab = '<w:r><w:tab/></w:r>'
$cfCap = '<w:ind w:left="1560" w:hanging="1560"/>'
$cfXml = "<w:document $cfW><w:body>" +
    (& $cfP $cfCap ((& $cfR 'ID GIA:' '<w:b/>') + $cfTab + (& $cfR '1398' ''))) +
    (& $cfP $cfCap ((& $cfR 'Exp. N&#250;m: ' '<w:b/>') + $cfTab + (& $cfR '2026/28/3165' ''))) +
    (& $cfP $cfCap ((& $cfR 'Objecte: ' '<w:b/>') + $cfTab + (& $cfR ('Doc. aportada amb N&#250;m. d&#8217;anotaci&#243; 21327 del 30/04/2026') ''))) +
    (& $cfP '' '') + (& $cfP '' (& $cfR 'INFORME' '<w:b/>')) +
    (& $cfP '' (& $cfR 'Nota: S&#8217;ha publicat l&#8217;Ordenan&#231;a' '')) +
    (& $cfP '' (& $cfR 'S''han observat les seg&#252;ents defici&#232;ncies que cal esmenar per poder exercir l''activitat:' '')) +
    (& $cfP '' '') +
    (& $cfP '' (& $cfR 'INSTAL&#183;LACIONS' '')) + (& $cfP '' '') +
    (& $cfP '' ((& $cfR '1. ' '<w:b/>') + (& $cfR 'Ascensors' ''))) +
    (& $cfP '' ((& $cfR '2. Punt pen' '<w:b/>') + (& $cfR 'dent del seguiment' '<w:b/>'))) +
    (& $cfP '' ((& $cfR '3. Amb ' '') + (& $cfR 'negreta del cataleg' '<w:b/>'))) +
    (& $cfP '' ((& $cfR 'https://canal' '<w:sz w:val="20"/>') + (& $cfR 'empresa.cat/a' '<w:sz w:val="20"/>'))) +
    (& $cfP '<w:spacing w:before="240"/><w:ind w:left="567" w:hanging="283"/>' ((& $cfR ([string][char]0x2022) '') + $cfTab + (& $cfR 'Sub-punt' ''))) +
    (& $cfP '<w:spacing w:before="200" w:after="240"/>' ((& $cfR '07/09/2026: ' '') + (& $cfR 'No s''aporta.' '<w:b/>'))) +
    (& $cfP '' '') +
    (& $cfP '<w:jc w:val="center"/><w:spacing w:after="240"/>' (& $cfR 'CONCLUSIONS' '<w:b/>')) +
    (& $cfP '<w:spacing w:after="240"/>' (& $cfR 'Vist l&#8217;anterior, cal requerir l&#8217;esmena.' '')) +
    (& $cfP '' '') +
    (& $cfP '' (& $cfR 'Ho poso al seu coneixement als efectes oportuns,' '')) +
    (& $cfP '' (& $cfR 'Cornell&#224; de Llobregat, 29 de setembre de 2026' '')) +
    '</w:body></w:document>'
$cfLl = _CorreuDocxLlegeix $cfXml $cfFmt
AssertEq ([string]$cfLl.Capcalera['id gia']) '1398' '_CorreuDocxLlegeix: la capcalera de l''informe (etiqueta<tab>valor)'
AssertEq ([string]$cfLl.Capcalera['exp. num']) '2026/28/3165' '_CorreuDocxLlegeix: l''etiqueta es compara normalitzada ("Exp. Num: ")'
Assert (-not $cfLl.Html.Contains('1398')) '_CorreuDocxLlegeix: el cos NO repeteix la capcalera (defecte d''abans: "ID GIA" en majuscules)'
Assert (-not $cfLl.Html.Contains('INFORME')) '_CorreuDocxLlegeix: ni "INFORME"'
Assert (-not $cfLl.Html.Contains('Nota:')) '_CorreuDocxLlegeix: ni la nota de l''Ordenanca'
Assert (-not $cfLl.Html.Contains('observat')) '_CorreuDocxLlegeix: ni la frase del cataleg (el correu porta la seva, {INTRO})'
Assert ($cfLl.Html.StartsWith('<p') -and $cfLl.Html.Contains('INSTAL')) '_CorreuDocxLlegeix: comenca a la primera seccio'
Assert ($cfLl.Html.Contains('<b>1. </b>Ascensors')) '_CorreuDocxLlegeix: la negreta del numero, del .docx'
Assert ($cfLl.Html.Contains('<a href="https://canalempresa.cat/a" style="font-size:10pt;word-break:break-all">')) '_CorreuDocxLlegeix: un URL partit en dos runs surt sencer, a 10 pt'
Assert ($cfLl.Html.Contains('<table') -and $cfLl.Html.Contains('Sub-punt')) '_CorreuDocxLlegeix: el pic amb sagnia francesa, com els blocs'
Assert ($cfLl.Html.Contains('<b>2. </b>Punt pendent del seguiment')) '_CorreuDocxLlegeix: un punt TOT en negreta (seguiment antic) -> negreta NOMES al numero'
Assert ($cfLl.Html.Contains('3. Amb <b>negreta del cataleg</b>')) '_CorreuDocxLlegeix: la negreta d''una part del punt (cataleg) es queda'
Assert ($cfLl.Html.Contains('07/09/2026: <b>No s''aporta.</b>')) '_CorreuDocxLlegeix: l''anotacio del seguiment, amb el comentari en negreta'
Assert ($cfLl.Html.Contains('margin:10pt 0 12pt 0px')) '_CorreuDocxLlegeix: els espais de l''anotacio, del .docx'
Assert ($cfLl.Html.Contains('text-align:center"><b>CONCLUSIONS</b>')) '_CorreuDocxLlegeix: CONCLUSIONS centrat i en negreta'
Assert ($cfLl.Html.Contains('cal requerir')) '_CorreuDocxLlegeix: les conclusions HI SON'
Assert (-not $cfLl.Html.Contains('Ho poso') -and -not $cfLl.Html.Contains('Llobregat, 29')) '_CorreuDocxLlegeix: el tancament i la signatura, NO'
Assert (-not $cfLl.Html.EndsWith('&nbsp;</p>')) '_CorreuDocxLlegeix: sense linies en blanc al final'
# Un "CONCLUSIONS" que es queda sol (cap conclusio triada) no surt.
$cfSol = _CorreuDocxLlegeix ($cfXml.Replace((& $cfP '<w:spacing w:after="240"/>' (& $cfR 'Vist l&#8217;anterior, cal requerir l&#8217;esmena.' '')), '')) $cfFmt
Assert (-not $cfSol.Html.Contains('CONCLUSIONS')) '_CorreuDocxLlegeix: un CONCLUSIONS sense cap conclusio no surt'

Write-Host "`n--- EnviarCorreu.ps1: el correu sencer del PC ---"
$cfC = _BuildCorreu $cfXml @{ ID_GIA = '1398'; ADRECA = 'C VISTALEGRE 24-26'; ACTIVITAT = 'APARCAMENT'; TITULAR = 'MEGADOCAR SL' } '29/09/2026'
Assert ($cfC.Html.Contains('<b>Exp. N' + [char]0x00FA + 'm:</b>') -and $cfC.Html.Contains('2026/28/3165')) '_BuildCorreu: la capcalera porta l''Exp. Num (de l''informe)'
Assert ($cfC.Html.Contains('MEGADOCAR SL')) '_BuildCorreu: el que l''informe no porta, de l''Excel'
Assert ($cfC.Intro.Contains('21327') -and $cfC.Intro.Contains('30/04/2026')) '_BuildCorreu: documentacio aportada -> la frase amb l''anotacio'
Assert (-not $cfC.Intro.Contains('visita')) '_BuildCorreu: documentacio aportada -> MAI "la visita" (defecte que va veure l''usuari)'
# UN sol cop a la part de dalt (l'altre es al peu: "feu-hi constar: ID GIA ...").
$cfDalt = $cfC.Html.Substring(0, $cfC.Html.IndexOf('INSTAL'))
AssertEq ([regex]::Matches($cfDalt, '1398').Count) 1 '_BuildCorreu: l''ID GIA surt UN sol cop abans dels requeriments'
$cfXmlV = $cfXml.Replace('Doc. aportada amb N&#250;m. d&#8217;anotaci&#243; 21327 del 30/04/2026', ('Visita inspecci' + [char]0x00F3 + ' '))
$cfCV = _BuildCorreu $cfXmlV @{ ID_GIA = '1398' } '29/09/2026'
# Sense la dada a l'Excel, el peu ("feu-hi constar: ... Titular X") la pren de
# la capcalera de l'informe.
$cfLinTit = (& $cfP $cfCap ((& $cfR 'Titular: ' '<w:b/>') + $cfTab + (& $cfR 'MEGADOCAR DOCX' '')))
$cfCN = _BuildCorreu ($cfXml.Replace('<w:body>', '<w:body>' + $cfLinTit)) @{ ID_GIA = '1398' } '29/09/2026'
Assert ($cfCN.Html.Contains('Titular MEGADOCAR DOCX')) '_BuildCorreu: sense l''Excel, el peu agafa el titular de l''informe'
Assert ($cfCV.Intro.Contains('visita') -and $cfCV.Intro.Contains('29/09/2026')) '_BuildCorreu: visita sense data -> la d''avui a la frase'
Assert ($cfCV.Html.Contains(('Visita inspecci' + [char]0x00F3 + ' 29/09/2026'))) '_BuildCorreu: ...i tambe a l''Objecte'

} catch {
    Assert $false ('CorreuFormat: la prova ha petat: ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
}

Write-Host "`n--- El MOBIL dona el MATEIX HTML que el PC (prova creuada amb Node) ---"
# La seleccio de debo de REQ1 (tots els punts sense [CAMP:]/[OPCIO:], que al
# mobil es resolen amb els valors del formulari) passa pels blocs i l'HTML del
# PC i pels de docs\correu.js, i ha de sortir EXACTAMENT el mateix.
$cfNode = Get-Command node -ErrorAction SilentlyContinue
if ($null -eq $cfNode) {
    Write-Host '  OMES  sense Node en aquesta maquina: la prova creuada no s''ha fet' -ForegroundColor Yellow
} else {
    try {
        $cfSenseCamps = { param($ls) -not ((@($ls) -join ' ') -match '\[(CAMP|OPCIO):') }
        $cfParsed = Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')
        $cfSel = @()
        foreach ($sec in @($cfParsed.Sections)) {
            $its = @()
            foreach ($el in @($sec.Items)) {
                if ($el.Kind -eq 'subsection' -or $el.Kind -eq 'intro') {
                    $its += [pscustomobject]@{ Kind = [string]$el.Kind; Short = [string]$el.Short; BodyLines = @($el.BodyLines); Children = @(); Selected = $false }
                    continue
                }
                if (-not (& $cfSenseCamps $el.BodyLines)) { continue }
                $fills = @(@($el.Children) | Where-Object { & $cfSenseCamps $_.BodyLines } | ForEach-Object { [pscustomobject]@{ Kind = [string]$_.Kind; Short = [string]$_.Short; BodyLines = @($_.BodyLines) } })
                $its += [pscustomobject]@{ Kind = 'item'; Short = [string]$el.Short; BodyLines = @($el.BodyLines); Children = $fills; Selected = $true }
            }
            $cfSel += [pscustomobject]@{ Title = [string]$sec.Title; Items = $its }
        }
        $cfConcl = @(('Vist l' + [char]0x2019 + 'anterior, cal requerir l' + [char]0x2019 + 'esmena.'), 'La terrassa **no** forma part d''aquest informe.')
        $cfBlocsPc = @(Build-CatalegBlocs $cfSel @{} '') + @(_BlocsConclusions 'CONCLUSIONS' $cfConcl @() @{})
        $cfHtmlPc = _CorreuBlocsAHtml $cfBlocsPc $cfFmt
        $cfLinCap = @(@{ Etiqueta = 'ID GIA:'; Valor = '118' }, @{ Etiqueta = 'Objecte:'; Valor = ('Visita inspecci' + [char]0x00F3 + ' 29/09/2026') })
        $cfCapPc = _CorreuCapcaleraHtml $cfLinCap $cfFmt 'INFORME'
        $cfCosPc = _CorreuCosAHtml "{CAPCALERA}`n`nHola {X} !!vermell!!`n{REQUERIMENTS}`nhttps://seu.cat/a?b=1&c=2" { param($s) $s.Replace('{X}', 'Y') } 'CAP' 'REQ' $cfFmt

        $cfTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('correu-creuat-' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $cfTmp -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $cfTmp 'in.json'), (([ordered]@{ Sel = $cfSel; Concl = $cfConcl; Cap = $cfLinCap } | ConvertTo-Json -Depth 12)), (New-Object System.Text.UTF8Encoding($false)))
        $cfJs = @"
var C = require(process.argv[2]), fs = require('fs');
var inp = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
var fmt = JSON.parse(fs.readFileSync(process.argv[4], 'utf8'));
function arr(x) { return x == null ? [] : (Array.isArray(x) ? x : [x]); }
var sel = arr(inp.Sel).map(function (s) { return { Title: s.Title, Items: arr(s.Items).map(function (i) {
  return { Kind: i.Kind, Short: i.Short, BodyLines: arr(i.BodyLines), Selected: i.Selected,
           Children: arr(i.Children).map(function (c) { return { Kind: c.Kind, Short: c.Short, BodyLines: arr(c.BodyLines) }; }) }; }) }; });
var blocs = C.blocsDeSeleccio(sel, function (n) { return arr(n.BodyLines); }).concat(C.blocsConclusions('CONCLUSIONS', arr(inp.Concl)));
fs.writeFileSync(process.argv[5], JSON.stringify({
  req: C.blocsAHtml(blocs, fmt),
  cap: C.capcaleraHtml(arr(inp.Cap), fmt, 'INFORME'),
  cos: C.cosAHtml('{CAPCALERA}\n\nHola {X} !!vermell!!\n{REQUERIMENTS}\nhttps://seu.cat/a?b=1&c=2', function (s) { return s.replace('{X}', 'Y'); }, 'CAP', 'REQ', fmt)
}));
"@
        [System.IO.File]::WriteAllText((Join-Path $cfTmp 't.js'), $cfJs, (New-Object System.Text.UTF8Encoding($false)))
        & $cfNode.Source (Join-Path $cfTmp 't.js') (Join-Path $cfRoot (Join-Path 'docs' 'correu.js')) (Join-Path $cfTmp 'in.json') (Join-Path $cfDades 'correu-format.json') (Join-Path $cfTmp 'out.json')
        $cfOut = [System.IO.File]::ReadAllText((Join-Path $cfTmp 'out.json')) | ConvertFrom-Json
        Assert ($cfHtmlPc.Length -gt 5000) ('creuat: la seleccio de REQ1 es gran (' + $cfHtmlPc.Length + ' caracters)')
        $cfDif = -1
        for ($i = 0; $i -lt [Math]::Min($cfHtmlPc.Length, ([string]$cfOut.req).Length); $i++) { if ($cfHtmlPc[$i] -ne ([string]$cfOut.req)[$i]) { $cfDif = $i; break } }
        if ($cfDif -ge 0) { Write-Host ('    PC:    ' + $cfHtmlPc.Substring([Math]::Max(0, $cfDif - 80), 200)); Write-Host ('    mobil: ' + ([string]$cfOut.req).Substring([Math]::Max(0, $cfDif - 80), 200)) }
        AssertEq ([string]$cfOut.req) $cfHtmlPc 'creuat: els requeriments i les conclusions, IGUALS al mobil i al PC'
        AssertEq ([string]$cfOut.cap) $cfCapPc 'creuat: la capcalera, IGUAL al mobil i al PC'
        AssertEq ([string]$cfOut.cos) $cfCosPc 'creuat: el cos del correu, IGUAL al mobil i al PC'
        Remove-Item -LiteralPath $cfTmp -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Assert $false ('creuat: la prova ha petat: ' + $_.Exception.Message)
    }
}
