# L'eina NORMATIVA (Normativa.ps1 i suport\normativa.json)
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.
#
# Peticio de l'usuari (setembre 2026): tota la normativa de REQ1 i dels seus
# marcadors en una sola carpeta (local\normativa), classificada pel NOM del
# fitxer ("Vector ambiental_Residus_2016_Decret 197-2016 ..."), amb el text
# consolidat, un index en Excel, que s'actualitzi sola i que la fitxa d'ajuda
# de cada punt obri el PDF desat. Les baixades nomes corren a Windows; aqui es
# prova tot el que decideix.

try {

$nmNormes = @(Get-NormativaCataleg)

Write-Host "`n--- normativa.json: el cataleg ---"
Assert ($nmNormes.Count -ge 150) ('cataleg: hi ha les normes de REQ1 i dels marcadors (' + $nmNormes.Count + ')')
$nmIds = @($nmNormes | ForEach-Object { [string]$_.Id })
AssertEq (@($nmIds | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name }) -join ', ') '' 'cataleg: cap Id repetit'
$nmNoms = @($nmNormes | ForEach-Object { _NormativaNomFitxer $_ })
AssertEq (@($nmNoms | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name }) -join ', ') '' 'cataleg: cap nom de fitxer repetit (una norma no en pot trepitjar una altra)'
AssertEq (@($nmNoms | Where-Object { $_ -match '[\\/:*?"<>|]' }) -join ', ') '' 'cataleg: cap nom de fitxer amb caracters prohibits al Windows'
AssertEq (@($nmNoms | Where-Object { $_.Length -gt 150 }) -join ', ') '' 'cataleg: noms prou curts perque la ruta sencera hi capiga (260)'
$nmSenseAmbit = @($nmNormes | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.Ambit) -or [string]::IsNullOrWhiteSpace([string]$_.Titol) } | ForEach-Object { [string]$_.Id })
AssertEq ($nmSenseAmbit -join ', ') '' 'cataleg: totes tenen ambit i titol'
$nmUrlDolent = @($nmNormes | Where-Object { [string]$_.Url -and -not ([string]$_.Url).StartsWith('https://') } | ForEach-Object { [string]$_.Id })
AssertEq ($nmUrlDolent -join ', ') '' 'cataleg: els URL son https (o buits: es desa a ma)'
# Res personal al repositori public: ni el OneDrive de la persona que va
# exportar els marcadors, ni cap adreca de correu.
$nmRaw = [System.IO.File]::ReadAllText($Script:NormativaCatalegPath)
Assert (-not ($nmRaw -match '(?i)sharepoint\.com|onedrive|@[a-z0-9-]+\.(cat|com|es)\b')) 'cataleg: cap enllac personal (OneDrive) ni adreca de correu'

Write-Host "`n--- el NOM del fitxer ---"
$nmE = [pscustomobject]@{ Id = 'Decret 197/2016'; Ambit = 'Vector ambiental'; Tema = 'Residus'; Any = '2016'; Tipus = 'Decret'; Num = '197/2016'; Titol = 'Registres productors' }
AssertEq (_NormativaNomFitxer $nmE) 'Vector ambiental_Residus_2016_Decret 197-2016 Registres productors.pdf' 'nom: Ambit_Tema_Any_Norma Titol (l''exemple de l''usuari, amb l''any)'
$nmE2 = [pscustomobject]@{ Id = 'Llei 3/2010'; Url = 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3'; Ambit = 'Incendis'; Tema = ''; Any = '2010'; Tipus = 'Llei'; Num = '3/2010'; Titol = 'Prevenció d' + [char]0x2019 + 'incendis' }
AssertEq (_NormativaNomFitxer $nmE2) ('Incendis_2010_Llei 3-2010 Prevenció d''incendis.pdf') 'nom: sense tema no hi ha dos guions baixos, i l''apostrof tipografic passa a recte'
$nmE3 = [pscustomobject]@{ Ambit = 'A'; Tema = ''; Any = ''; Tipus = 'Llei'; Num = '1/2000'; Titol = ('x' * 90) }
Assert ((_NormativaNomFitxer $nmE3).Length -le (20 + $Script:NormativaTitolMax)) 'nom: el titol es talla'

Write-Host "`n--- reconeixer una norma dins d'un text ---"
AssertEq (_NormativaNormText 'Real Decreto 842/2002') 'rd 842/2002' 'normalitza: Real Decreto -> rd'
AssertEq (_NormativaNormText 'Reial Decret 842/2002') 'rd 842/2002' 'normalitza: Reial Decret -> rd'
AssertEq (_NormativaNormText 'Decreto Legislativo 1/2010') 'dleg 1/2010' 'normalitza: Decreto Legislativo -> dleg'
AssertEq (_NormativaNormText 'Real Decreto Legislativo 2/2004') 'rdleg 2/2004' 'normalitza: el legislatiu no es confon amb el RD'
AssertEq (_NormativaNormText 'Ley 34/2007') (_NormativaNormText 'Llei 34/2007') 'normalitza: Ley = Llei'
AssertEq (_NormativaNormText 'Reglamento (CE) nº 852/2004') 'reglament 852/2004' 'normalitza: Reglamento (CE) nº'
$nmTrobada = _NormativaBuscaEnText $nmNormes 'Llei 16/2002, del 28 de juny, de protecció contra la contaminació acústica. Desplegament: Decret 176/2009'
AssertEq ([string]$nmTrobada.Id) 'Llei 16/2002' 'busca: la PRIMERA norma citada (la principal)'
AssertEq ([string](_NormativaBuscaEnText $nmNormes 'Real Decreto 164/2025, de 4 de marzo').Id) 'RD 164/2025' 'busca: un Real Decreto el troba com a RD'
AssertEq ([string](_NormativaBuscaEnText $nmNormes 'RD 1021/2022').Id) 'RD 1021/2022' 'busca: tambe abreujat'
Assert ($null -eq (_NormativaBuscaEnText @([pscustomobject]@{ Id = 'RD 9/2005'; Tipus = 'RD'; Num = '9/2005' }) 'Real Decreto 19/2005')) 'busca: "RD 9/2005" no es troba dins "RD 19/2005"'
Assert ($null -eq (_NormativaBuscaEnText @([pscustomobject]@{ Id = 'Decret 30/2015'; Tipus = 'Decret'; Num = '30/2015' }) 'Decret 130/2015')) 'busca: ni "Decret 30/2015" dins "Decret 130/2015"'
AssertEq ([string](_NormativaBuscaEnText $nmNormes 'Ordenança relativa a la intervenció administrativa municipal en les activitats de Cornellà de Llobregat, art. 9').Id) 'Ordenança Cornellà activitats' 'busca: l''ordenanca municipal, per les seves claus'
Assert ($null -eq (_NormativaBuscaEnText $nmNormes 'Normes UNE (no tenen text consolidat)')) 'busca: un text sense cap norma -> $null'

Write-Host "`n--- TOTA la normativa que cita REQ1 hi es ---"
# Cada fitxa d'ajuda de REQ1 porta la seva norma: si la fitxa en cita una que
# no es al cataleg, la carpeta queda coixa sense que ningu no ho vegi.
$nmReq1 = Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')
$nmSenseCat = New-Object System.Collections.ArrayList
foreach ($s in @($nmReq1.Sections)) {
    foreach ($el in @(@($s.Items) + @(@($s.Items) | ForEach-Object { @($_.Children) }))) {
        if ($null -eq $el -or $null -eq $el.Ajuda) { continue }
        $n = [string]$el.Ajuda.Norma
        if (-not $n -or $n -match '(?i)^Normes UNE|^Ordenances metropolitanes') { continue }
        if ($null -eq (_NormativaBuscaEnText $nmNormes $n)) { [void]$nmSenseCat.Add(([string]$el.Short + ': ' + $n.Substring(0, [Math]::Min(60, $n.Length)))) }
    }
}
AssertEq (@($nmSenseCat) -join ' | ') '' 'REQ1: la norma de cada fitxa d''ajuda es al cataleg de normativa'
$nmPunts = _NormativaPuntsReq1 $nmNormes $nmReq1
Assert ($nmPunts.ContainsKey('RD 164/2025') -and @($nmPunts['RD 164/2025']).Count -ge 10) 'REQ1: l''index diu a quins punts surt cada norma (RSCIEI, a molts)'

Write-Host "`n--- d'on i com es baixa ---"
AssertEq (_NormativaFont 'https://www.boe.es/buscar/act.php?id=BOE-A-2013-12913') 'boe' 'font: pagina del BOE'
AssertEq (_NormativaFont 'https://www.boe.es/eli/es/rd/2025/03/04/164/con') 'boe' 'font: ELI del BOE'
AssertEq (_NormativaFont 'https://www.codigotecnico.org/pdf/Documentos/SI/DccSI.pdf') 'pdf' 'font: un PDF'
AssertEq (_NormativaFont 'https://eur-lex.europa.eu/legal-content/ES/TXT/PDF/?uri=CELEX:32004R0852') 'pdf' 'font: EUR-Lex en PDF'
AssertEq (_NormativaFont 'https://portaljuridic.gencat.cat/eli/es-ct/l/2009/12/04/20') 'web' 'font: el Portal Juridic s''imprimeix a PDF'
AssertEq (_NormativaFont '') 'manual' 'font: sense URL, a ma'
$nmHtml = '<a href="/buscar/pdf/2025/BOE-A-2025-7036-consolidado.pdf">PDF</a> <p>Última actualización publicada el 10/04/2025</p> <a href="/boe/dias/2025/04/10/pdfs/BOE-A-2025-7036.pdf">orig</a>'
$nmInfo = _NormativaBoeInfo $nmHtml
AssertEq $nmInfo.Id 'BOE-A-2025-7036' 'BOE: l''identificador, de l''enllac del consolidat'
AssertEq $nmInfo.Versio '10/04/2025' 'BOE: la data de l''ultima actualitzacio (la versio)'
AssertEq $nmInfo.Original 'https://www.boe.es/boe/dias/2025/04/10/pdfs/BOE-A-2025-7036.pdf' 'BOE: l''original, amb el domini'
AssertEq (_NormativaBoePdfConsolidat 'BOE-A-2025-7036') 'https://www.boe.es/buscar/pdf/2025/BOE-A-2025-7036-consolidado.pdf' 'BOE: el PDF consolidat'
AssertEq (_NormativaBoeInfo '<p>sense res</p>').Versio '' 'BOE: una pagina sense consolidat no te versio'
Assert (_NormativaEsPdf ([byte[]](0x25, 0x50, 0x44, 0x46, 0x2D))) 'PDF: %PDF- es un PDF'
Assert (-not (_NormativaEsPdf ([System.Text.Encoding]::ASCII.GetBytes('<html>')))) 'PDF: una pagina d''error no ho es'

Write-Host "`n--- el boto PDF d'una pagina (Portal Juridic, BOPB) ---"
# Com el de la captura de l'usuari: "Descarrega  PDF  RDF  TTL  XML", i al
# costat el resum fet amb IA, que NO es la norma.
$nmPj = @'
<div class="resum"><a href="/documents/resum-ia-ca.pdf">Descarrega (CA)</a> <a href="/documents/resum-ia-es.pdf">Descarga (ES)</a></div>
<span>Descarrega</span>
<a class="pdf" href="/ca/document-del-pjur/?documentId=547998&amp;format=pdf"><img alt=""> PDF</a>
<a href="/eli/es-ct/l/2010/02/18/3/rdf">RDF</a> <a href="/eli/es-ct/l/2010/02/18/3/ttl">TTL (Turtle)</a> <a href="/eli/es-ct/l/2010/02/18/3/xml">XML</a>
<a href="https://altra.cat/guia.pdf">guia relacionada</a>
'@
$nmCands = @(_NormativaPdfsDeHtml $nmPj 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3')
AssertEq ([string]$nmCands[0]) 'https://portaljuridic.gencat.cat/ca/document-del-pjur/?documentId=547998&format=pdf' 'PDF de la pagina: primer el boto "PDF", amb l''adreca completa'
Assert (-not (@($nmCands) -match 'resum')) 'PDF de la pagina: mai el resum fet amb IA'
Assert (-not (@($nmCands) -match '/(rdf|ttl|xml)$')) 'PDF de la pagina: ni RDF, TTL o XML'
AssertEq @(_NormativaPdfsDeHtml '' 'https://x.cat/').Count 0 'PDF de la pagina: sense HTML, cap'
AssertEq @(_NormativaPdfsDeHtml '<a href="javascript:void(0)">PDF</a>' 'https://x.cat/').Count 0 'PDF de la pagina: un boto de JavaScript no serveix'

Write-Host "`n--- les col·leccions (ITC de Bombers, TINSCI) ---"
$nmCols = @($nmNormes | Where-Object { $_.Colleccio })
AssertEq (@($nmCols | ForEach-Object { $_.Id }) -join ', ') 'Col·lecció ITC Bombers, Col·lecció TINSCI' 'col·leccions: ITC i TINSCI'
# Les "ITC antigues" eren les TINSCI (pagina "Documentacio normativa: TINSCI",
# 36 documents DT-x a l'index de l'usuari, octubre 2026): ara son una font mes
# de la col·leccio TINSCI, i els fitxers ja baixats s'hi reanomenen (Abans).
$nmTin = @($nmCols | Where-Object { $_.Id -eq 'Col·lecció TINSCI' })[0]
Assert (@($nmTin.AltresUrls) -contains 'https://interior.gencat.cat/ca/detalls/Article/Documentacio_normativa_TINSCI') 'TINSCI: tambe de la pagina "Documentacio normativa: TINSCI"'
Assert (@($nmTin.Abans) -contains 'Col·lecció ITC Bombers antigues') 'TINSCI: hereta els fitxers de les "ITC antigues"'
$nmTinE = [pscustomobject]@{ Ambit = 'Incendis'; Tema = 'TINSCI' }
AssertEq (_NormativaNomDocColleccio $nmTinE 'Document actualitzat DT-10' 'https://x.cat/dt10.pdf') 'Incendis_TINSCI_Document actualitzat DT-10.pdf' 'TINSCI: la versio vigent, al tema TINSCI'
AssertEq (_NormativaNomDocColleccio $nmTinE 'Document anterior DT-10 (Obre en una nova finestra)' 'https://x.cat/dt10a.pdf') 'Incendis_Antic_TINSCI Document anterior DT-10.pdf' 'TINSCI: les versions anteriors, a Antic'
AssertEq (_NormativaNetejaTextEnllac 'SP 112 (Obre en una nova finestra)') 'SP 112' 'col·leccio: el titol tambe sense el "(Obre en una nova finestra)"'
# Una web que demana iniciar sessio: l'enllac hi es, pero es desa a ma.
$nmApa = @($nmNormes | Where-Object { $_.Id -like '*APABCN 093*' })[0]
AssertEq (_NormativaFontDe $nmApa) 'manual' 'APABCN 093: area privada, es desa a ma'
Assert ([string]$nmApa.Url) 'APABCN 093: pero conserva l''enllac per obrir-la'
$nmApaF = _NormativaFilesIndex @($nmApa) @{} @{} { param($n) $false }
Assert (([string]@($nmApaF)[0][9]) -like 'Web amb accés restringit*') 'APABCN 093: l''index diu per que s''ha de desar a ma'
$nmSrcN = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Normativa.ps1'))
Assert ($nmSrcN -match "catch \{ if \(-not \(\(\[string\]\`$e\.Url\) -match '/con/\?\`$'\)\) \{ throw \} \}") 'BOE: el 404 de la pagina /con (RD 1002/2002) no talla el respatller sense /con'
$nmRd = @($nmNormes | Where-Object { $_.Num -eq '1002/2002' })[0]
AssertEq ([string]$nmRd.Url) 'https://www.boe.es/buscar/doc.php?id=BOE-A-2002-19574' 'RD 1002/2002: sense text consolidat, la publicacio del BOE (l''ELI tornava 404 tambe sense /con)'
AssertEq (_NormativaFont ([string]$nmRd.Url)) 'boe' 'RD 1002/2002: es baixa pel cami del BOE'
foreach ($nmUe in @('2016/679', '2017/745')) {
    $nmUeE = @($nmNormes | Where-Object { $_.Num -eq $nmUe })[0]
    Assert ([string]$nmUeE.Url -like 'https://www.boe.es/buscar/doc.php?id=DOUE-L-*') "Reglament (UE) ${nmUe}: des del BOE (EUR-Lex tornava un PDF buit)"
}

$nmColHtml = @'
<ul>
<li><a href="/web/.content/sp/SP-120_entorn.pdf">SP 120. Condicions d&#39;entorn i d&#39;aproximació als edificis (PDF, 1,2 MB)</a></li>
<li><a href="https://interior.gencat.cat/web/.content/sp/SP-122.pdf">SP 122. Aparcaments</a></li>
<li><a href="/web/.content/sp/SP-122.pdf">SP 122 (duplicat)</a></li>
<li><a href="/web/.content/resum-ia.pdf">Resum IA</a></li>
<li><a href="/ca/arees/fitxa-sp-130/">SP 130</a></li>
</ul>
'@
$nmColBase = 'https://interior.gencat.cat/ca/arees_dactuacio/bombers/instruccions_tecniques_complementaries/'
$nmDocs = @(_NormativaDocsDeColleccio $nmColHtml $nmColBase)
AssertEq $nmDocs.Count 2 'col·leccio: els PDF de la pagina, sense repetits ni el resum'
AssertEq ([string]$nmDocs[0].Url) 'https://interior.gencat.cat/web/.content/sp/SP-120_entorn.pdf' 'col·leccio: l''adreca completa'
AssertEq ([string]$nmDocs[0].Text) ('SP 120. Condicions d' + "'" + 'entorn i d' + "'" + 'aproximació als edificis (PDF, 1,2 MB)') 'col·leccio: el text de l''enllac'
$nmColE = [pscustomobject]@{ Ambit = 'Incendis'; Tema = 'ITC Bombers' }
AssertEq (_NormativaNomDocColleccio $nmColE ([string]$nmDocs[0].Text) ([string]$nmDocs[0].Url)) ('Incendis_ITC Bombers_SP 120. Condicions d' + "'" + 'entorn i d' + "'" + 'aproximació als edificis.pdf') 'col·leccio: el nom (sense el "(PDF, 1,2 MB)")'
AssertEq (_NormativaNomDocColleccio $nmColE 'PDF' 'https://x.cat/docs/TINSCI_03_evacuacio.pdf') 'Incendis_ITC Bombers_TINSCI 03 evacuacio.pdf' 'col·leccio: un enllac que nomes diu "PDF" pren el nom del fitxer'
$nmSubs = @(_NormativaSubpagines '<a href="/ca/arees_dactuacio/bombers/instruccions_tecniques_complementaries/sp-130/">SP 130</a><a href="/ca/altres/">No</a><a href="./">Aqui</a>' $nmColBase)
AssertEq $nmSubs.Count 1 'col·leccio: les pagines filles, nomes les que pengen de la pagina'
AssertEq ([string]$nmSubs[0].Text) 'SP 130' 'col·leccio: amb el seu text'
# Les TINSCI: la pagina enllaca la fitxa de cada document al repositori
# d'Interior (DSpace, un altre servidor), i el PDF es a dins de la fitxa.
$nmTiBase = 'https://interior.gencat.cat/ca/arees_dactuacio/bombers/prevencio_d_incendis/instruccions_guies_i_recomanacions/interpretacio_normativa_tinsci/documents-tinsci/'
$nmTi = @(_NormativaSubpagines ('<a href="https://dsp.interior.gencat.cat/handle/20.500.14007/6228">DT-04 Reducció de l' + "'" + 'amplada d' + "'" + 'escala (Obre en una nova finestra)</a><a href="http://hdl.handle.net/20.500.14007/7001">AT-09</a><a href="https://dsp.interior.gencat.cat/items/0b6c4a4e-1f0e-4c33-9d7c-6a1b2c3d4e5f">DT-12</a><a href="https://dsp.interior.gencat.cat/discover">Cerca</a><a href="https://www.gencat.cat/">Gencat</a>') $nmTiBase)
AssertEq $nmTi.Count 3 'TINSCI: les fitxes del repositori (DSpace, hdl.handle.net i la versio nova), encara que siguin d''un altre servidor'
$nmTiItem = '<a href="/bitstream/handle/20.500.14007/6228/DT-04-Reduccio_amplada_escala_ascensor_2015_07.pdf.jpg?sequence=12&amp;isAllowed=y"><img src="x"/></a><a href="/bitstream/handle/20.500.14007/6228/DT-04-Reduccio_amplada_escala_ascensor_2015_07.pdf?sequence=10&amp;isAllowed=y">Visualitza/Obre</a><a href="/bitstream/handle/20.500.14007/6228/license.txt?sequence=2">Llicència</a>'
$nmTiDocs = @(_NormativaDocsDeColleccio $nmTiItem 'https://dsp.interior.gencat.cat/handle/20.500.14007/6228')
AssertEq $nmTiDocs.Count 1 'TINSCI: a la fitxa, el PDF (ni la miniatura ni la llicencia)'
AssertEq ([string]$nmTiDocs[0].Url) 'https://dsp.interior.gencat.cat/bitstream/handle/20.500.14007/6228/DT-04-Reduccio_amplada_escala_ascensor_2015_07.pdf?sequence=10&isAllowed=y' 'TINSCI: l''adreca que va passar l''usuari'
AssertEq @(_NormativaDocsDeColleccio '<a href="/bitstreams/0b6c4a4e-1f0e-4c33-9d7c-6a1b2c3d4e5f/download">DT-12.pdf</a>' 'https://dsp.interior.gencat.cat/items/x').Count 1 'TINSCI: i el fitxer de la versio nova del DSpace (sense .pdf a l''adreca)'
$nmTiE = [pscustomobject]@{ Ambit = 'Incendis'; Tema = 'TINSCI' }
AssertEq (_NormativaNomDocColleccio $nmTiE 'Visualitza/Obre' ([string]$nmTiDocs[0].Url)) 'Incendis_TINSCI_DT-04-Reduccio amplada escala ascensor 2015 07.pdf' 'TINSCI: "Visualitza/Obre" no fa de nom: el del fitxer'
# La fitxa d'una ITC obre el PDF de la ITC (no el de la Llei 3/2010 que cita).
AssertEq (_NormativaSpDeText 'Instrucció tècnica complementària SP 144:2023, Condicions... Desplega la Llei 3/2010') '144' 'ITC: el numero de la fitxa'
AssertEq (_NormativaSpDeText 'Real Decreto 164/2025, annex II') '' 'ITC: una fitxa que no en cita cap'
$nmSpNoms = @('Incendis_ITC Bombers_SP 136 A.pdf', 'Incendis_ITC Bombers_SP 136.pdf', 'Incendis_ITC Bombers_SP 1360 x.pdf', 'Incendis_ITC Bombers_Nota 1 - SP 144.pdf', 'Incendis_ITC Bombers_SP 144 (Obre en una nova finestra).pdf')
AssertEq (_NormativaFitxerSp $nmSpNoms '136') 'Incendis_ITC Bombers_SP 136.pdf' 'ITC: el document principal, no el model A'
AssertEq (_NormativaFitxerSp $nmSpNoms '144') 'Incendis_ITC Bombers_SP 144 (Obre en una nova finestra).pdf' 'ITC: la SP 144 i no la nota (tambe amb el nom antic)'
AssertEq (_NormativaFitxerSp $nmSpNoms '109') '' 'ITC: si no hi es, res'
AssertEq (_NormativaDtDeText 'Document TINSCI DT-9, Control de fums en els aparcaments') '9' 'TINSCI: el numero de la fitxa'
$nmDtNoms = @('Incendis_TINSCI_Document DT-18.pdf', 'Incendis_TINSCI_Document actualitzat DT-18.pdf', 'Incendis_TINSCI_Document actualitzat DT-1.pdf', 'Incendis_TINSCI_Document actualizat DT-8.pdf', 'Incendis_Antic_TINSCI Document anterior DT-9.pdf', 'Incendis_TINSCI_Document actualitzat DT-9.pdf', 'Incendis_TINSCI_Annex 1.1 de DT-13.pdf', 'Incendis_TINSCI_Document DT-13.pdf')
AssertEq (_NormativaFitxerDt $nmDtNoms '18') 'Incendis_TINSCI_Document actualitzat DT-18.pdf' 'TINSCI: la versio actualitzada si n''hi ha dues'
AssertEq (_NormativaFitxerDt $nmDtNoms '8') 'Incendis_TINSCI_Document actualizat DT-8.pdf' 'TINSCI: tambe amb la falta del web ("actualizat")'
AssertEq (_NormativaFitxerDt $nmDtNoms '9') 'Incendis_TINSCI_Document actualitzat DT-9.pdf' 'TINSCI: mai l''anterior (Antic)'
AssertEq (_NormativaFitxerDt $nmDtNoms '13') 'Incendis_TINSCI_Document DT-13.pdf' 'TINSCI: el document, no l''annex'
AssertEq (_NormativaFitxerDt $nmDtNoms '1') 'Incendis_TINSCI_Document actualitzat DT-1.pdf' 'TINSCI: DT-1 no agafa la DT-18'
AssertEq (_NormativaNomDocColleccio $nmColE 'SP 132 (Obre en una nova finestra)' 'https://x.cat/sp132.pdf') 'Incendis_ITC Bombers_SP 132.pdf' 'col·leccio: fora el "(Obre en una nova finestra)" (sortia al nom de les ITC)'
$nmFc = _NormativaFilesIndex @([pscustomobject]@{ Id = 'Col·lecció ITC Bombers'; Ambit = 'Incendis'; Tema = 'ITC Bombers'; Tipus = 'ITC Bombers'; Titol = 'ITC'; Url = 'https://x.cat/'; Colleccio = $true }) @{
    'Col·lecció ITC Bombers' = @{ Error = ''; Baixat = '2026-09-30T10:00:00'; Versio = ''; Mida = 0 }
    'Col·lecció ITC Bombers | a.pdf' = @{ Pare = 'Col·lecció ITC Bombers'; Titol = 'SP 120'; Url = 'https://x.cat/a.pdf'; Nom = 'a.pdf'; Baixat = '2026-09-30T10:00:00'; Error = '' } } @{} { param($n) $true }
AssertEq @($nmFc).Count 2 'index: la col·leccio i un document'
AssertEq ([string]@($nmFc)[0][9]) '1 documents' 'index: la col·leccio diu quants documents'
AssertEq ([string]@($nmFc)[1][3]) 'SP 120' 'index: cada document amb el seu titol'

Write-Host "`n--- el Portal Juridic sense navegador (el PDF del DOGC) ---"
# L'adreca real del boto PDF, copiada per l'usuari (Llei 3/2010).
$nmPdfPj = 'https://portaldogc.gencat.cat/utilsEADOP/AppJava/PdfProviderServlet?versionId=2164170&type=01'
AssertEq (_NormativaPdfPjurDeText '<a href="https://portaldogc.gencat.cat/utilsEADOP/AppJava/PdfProviderServlet?versionId=2164170&amp;type=01">PDF</a>') $nmPdfPj 'Portal Juridic: l''enllac del boto PDF (amb &amp;)'
AssertEq (_NormativaPdfPjurDeText '{"documentId":547998,"versionId":"2164170"}') $nmPdfPj 'Portal Juridic: el numero de versio dins d''unes dades JSON'
AssertEq (_NormativaPdfPjurDeText '<eli:is_embodied_by rdf:resource="https://portaldogc.gencat.cat/utilsEADOP/AppJava/PdfProviderServlet?versionId=2164170&amp;type=02"/>') ($nmPdfPj -replace '01$', '02') 'Portal Juridic: el tipus es respecta si hi es'
AssertEq (_NormativaPdfPjurDeText '<p>res</p>') '' 'Portal Juridic: sense numero, res'
AssertEq (_NormativaEliDeText 'URI ELI: <a href="https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3">x</a>') 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3' 'Portal Juridic: l''ELI d''una pagina per documentId'
$nmMeta = @(_NormativaUrlsMetaPjur '<a href="/eli/es-ct/l/2010/02/18/3/rdf">RDF</a> <a href="/x/pdf">PDF</a>' 'https://portaljuridic.gencat.cat/ca/document-del-pjur/?documentId=547998' 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3')
AssertEq ([string]$nmMeta[0]) 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3/rdf' 'Portal Juridic: primer les metadades que enllaca la pagina'
Assert ($nmMeta -contains 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3/ttl') 'Portal Juridic: i les representacions de l''ELI'
Assert (-not ($nmMeta -contains 'https://portaljuridic.gencat.cat/x/pdf')) 'Portal Juridic: el PDF no es una metadada'
$nmSrcN = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Normativa.ps1'))
Assert ($nmSrcN -match "informes-normativa-edge-' \+ \[guid\]") 'Edge: un perfil NOU a cada crida (un de compartit el bloquejava el que es penjava)'
Assert ($nmSrcN.Contains("'/T', '/F', '/PID'")) 'Edge: si es penja, es mata tot l''arbre de processos'
Assert ($nmSrcN.Contains('if ($Script:NormativaEdgeKO.ContainsKey($host1)) { throw')) 'Edge: penjat amb una web, no es torna a fer servir AMB AQUELLA WEB en la passada'
Assert ($nmSrcN.Contains('$Script:NormativaEdgeKO[$host1] = $true') -and -not $nmSrcN.Contains('$Script:NormativaEdgeKO = $true')) 'Edge: el penjament del CIDO ja no deixa sense Edge el Portal Juridic (abans 34 normes van fallar per aixo)'
AssertEq (_NormativaHostDe 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3') 'portaljuridic.gencat.cat' 'Edge: el servidor d''un URL'
AssertEq (_NormativaFont 'https://www.boe.es/buscar/doc.php?id=DOUE-L-2004-81035') 'web' 'reglaments europeus: la pagina del BOE (EUR-Lex no deixa baixar fora d''un navegador) i el seu boto PDF'
$nmIdxBw = $nmSrcN.IndexOf('function _NormativaBaixaWeb')
Assert ($nmSrcN.IndexOf('_NormativaFontsPjur', $nmIdxBw) -lt $nmSrcN.IndexOf('_NormativaDomEdge', $nmIdxBw)) 'Portal Juridic: primer sense navegador, l''Edge nomes al final'

Write-Host "`n--- guies i manuals ---"
$nmGuies = @($nmNormes | Where-Object { $_.Guia })
Assert ($nmGuies.Count -ge 20) ('guies: hi son les dels marcadors (' + $nmGuies.Count + ')')
AssertEq (@($nmGuies | Where-Object { [string]$_.Tema -ne 'Guies' } | ForEach-Object { $_.Id }) -join ', ') '' 'guies: totes al tema Guies de cada ambit'
$nmFg = _NormativaFilesIndex @($nmGuies[0]) @{} @{} { param($n) $false }
AssertEq ([string]@($nmFg)[0][4]) $(if ($nmGuies[0].Derogada) { 'Guia (antiga)' } else { 'Guia' }) 'guies: a l''index surten com a Guia, no com a norma vigent'
AssertEq (@($nmNormes | Where-Object { (_NormativaFont ([string]$_.Url)) -eq 'manual' } | ForEach-Object { $_.Id }) -join ', ') '' 'cataleg: ja no n''hi ha cap sense enllac (les dues ordenances ja en tenen)'
AssertEq ([string](_NormativaBuscaEnText $nmNormes 'Llei 31/1991, del 13 de desembre. Desplegament: Decret 40/1992').Id) 'Llei 31/1991' 'REQ1: farmacies, amb el Decret 40/1992 (la fitxa deia 40/2006)'
Assert (-not ([System.IO.File]::ReadAllText((Join-Path $EstructuralsDir 'REQ1.json')).Contains('Decret 40/2006'))) 'REQ1: el Decret 40/2006 (que no existeix per a farmacies) ja no hi es'

Write-Host "`n--- quan s'ha de tornar a baixar (actualitzar sola) ---"
$nmAra = [datetime]'2026-09-30T10:00:00'
AssertEq (_NormativaCalBaixar $null $false '' $nmAra $false) 'nova' 'actualitzar: si no hi es, es baixa'
AssertEq (_NormativaCalBaixar @{ Versio = '10/04/2025'; Baixat = '2026-09-01T00:00:00' } $true '10/04/2025' $nmAra $false) '' 'actualitzar: BOE amb la mateixa versio -> res'
AssertEq (_NormativaCalBaixar @{ Versio = '10/04/2025'; Baixat = '2026-09-01T00:00:00' } $true '02/09/2026' $nmAra $false) 'versio' 'actualitzar: BOE amb versio nova -> es torna a baixar'
AssertEq (_NormativaCalBaixar @{ Versio = ''; Baixat = '2026-09-01T00:00:00' } $true '' $nmAra $false) '' 'actualitzar: una pagina baixada fa poc -> res'
AssertEq (_NormativaCalBaixar @{ Versio = ''; Baixat = '2025-09-01T00:00:00' } $true '' $nmAra $false) 'antiga' ('actualitzar: una pagina de fa mes de ' + $Script:NormativaDiesRefresc + ' dies -> es torna a baixar')
AssertEq (_NormativaCalBaixar @{ Versio = 'x'; Baixat = '2026-09-29T00:00:00' } $true 'x' $nmAra $true) 'forcat' 'actualitzar: "tornar-les a baixar totes"'
AssertEq (_NormativaNomAnterior 'Incendis_2010_Llei 3-2010 X.pdf' $nmAra) 'Incendis_2010_Llei 3-2010 X (fins 30-09-2026).pdf' 'actualitzar: l''anterior es guarda amb la data'

Write-Host "`n--- l'INDEX en Excel (sense Excel) ---"
Add-Type -AssemblyName System.IO.Compression
$nmFiles = _NormativaFilesIndex @($nmE, $nmE2) @{ ($nmE2.Tipus + ' ' + $nmE2.Num) = @{ Versio = ''; Baixat = '2026-09-01T00:00:00'; Error = 'Timeout'; Mida = 0 } } @{} { param($n) $n -like 'Vector*' }
AssertEq @($nmFiles).Count 2 'index: una fila per norma'
AssertEq ([string]@($nmFiles)[0][9]) 'Baixada' 'index: la que hi es, "Baixada"'
Assert (([string]@($nmFiles)[1][9]).StartsWith('Error: Timeout')) 'index: la que va fallar diu per que'
$nmBytes = _NormativaXlsxBytes $Script:NormativaCapcaleraIndex $nmFiles $Script:NormativaAmplesIndex
$nmZip = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $nmBytes)), [System.IO.Compression.ZipArchiveMode]::Read)
$nmParts = @($nmZip.Entries | ForEach-Object { $_.FullName })
foreach ($p in @('[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels')) {
    Assert ($nmParts -contains $p) ('xlsx: hi ha ' + $p)
}
$nmSr = New-Object System.IO.StreamReader($nmZip.GetEntry('xl/worksheets/sheet1.xml').Open())
$nmSheet = $nmSr.ReadToEnd(); $nmSr.Dispose(); $nmZip.Dispose()
$nmXml = $null
try { $nmXml = [xml]$nmSheet } catch { }
Assert ($null -ne $nmXml) 'xlsx: el full es XML valid'
Assert ($nmSheet.Contains('HYPERLINK(&quot;Vector ambiental_Residus_2016_Decret 197-2016 Registres productors.pdf&quot;')) 'xlsx: el fitxer baixat es un enllac relatiu (obre el PDF del costat)'
Assert ($nmSheet.Contains('<autoFilter ref="A1:K3"/>') -and $nmSheet.Contains('state="frozen"')) 'xlsx: amb filtres i la capcalera fixa'
Assert ($nmSheet.Contains([string][char]0x00C0 + 'mbit')) 'xlsx: els accents hi arriben (UTF-8)'
AssertEq (_NormativaXlsxCol 0) 'A' 'xlsx: columna A'
AssertEq (_NormativaXlsxCol 26) 'AA' 'xlsx: columna AA'

Write-Host "`n--- la carpeta i el menu ---"
AssertEq ([string]$Script:LocalSubdirs['Normativa']) 'normativa' 'carpeta: local\normativa (Get-LocalSubdir)'
$nmMenu = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1'))
$nmWiz = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1'))
Assert ($nmMenu -match "Action = 'normativa'" -and $nmWiz -match "'normativa'\s*\{\s*Invoke-Normativa\s*\}") 'menu: la rajola Normativa obre l''eina'
$nmAj = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'UiComuns.ps1'))
Assert ($nmAj.Contains('Get-NormativaPdfDeText') -and $nmAj.Contains("'Obre el PDF desat'")) 'fitxa d''ajuda: "Obre el PDF desat" si la norma ja es a la carpeta'
AssertEq (Get-NormativaPdfDeText 'Llei 3/2010') '' 'fitxa d''ajuda: sense la carpeta, cap PDF (i no peta)'

Write-Host "`n--- Les peticions HTTP viuen on toca (guard) ---"
# PER QUE. Les banderes d'una peticio son el que la fa funcionar darrere del
# proxy de l'Ajuntament (-UseDefaultCredentials) i amb els servidors que
# redirigeixen mes de 5 vegades (-MaximumRedirection 10). Escrites a ma a cada
# lloc, un lloc se'n deixa alguna: ProvarVigencia.ps1 es muntava la seva nomes
# per afegir-hi un Accept, i pel cami perdia les dues i enviava
# 'application/xml' a dues pagines .html que tot seguit passava per
# _RevEstatBoe. Era l'eina feta NOMES per explicar per que falla la vigencia,
# diagnosticant una resposta que el programa no rep mai.
#
# El guard no compta crides: diu en QUINS FITXERS hi poden ser. Cada un es la
# porta d'un servei (normativa, Cadastre, OSRM, Drive, EmailJS, enllacos), i qui
# vulgui parlar amb aquell servei hi ha de passar.
$nmHttpPermesos = @(
    'Normativa.ps1',              # _NormativaGet + _NormativaGetBytes
    'Enllacos.ps1',               # Test-EnllacViu (HEAD, i GET si el 405)
    'DriveApi.ps1',               # l'API del Drive
    'CorreuVia.ps1',              # EmailJS (Send-EmailJs)
    'mobil/Authorize-Drive.ps1',  # el token del Drive, un sol cop
    'rutes/Cadastre.ps1',         # el comu de totes les consultes al Cadastre
    'rutes/Ruta.ps1'              # OSRM (el planificador de rutes)
)
$nmArrelSup = Split-Path -Parent $TestsDir
$nmHttpTrobats = New-Object System.Collections.ArrayList
$nmHttpPerFitxer = @{}
foreach ($f in @(Get-ChildItem -LiteralPath $nmArrelSup -Recurse -Filter '*.ps1' -File)) {
    $rel = $f.FullName.Substring($nmArrelSup.Length + 1).Replace('\', '/')
    if ($rel.StartsWith('tests/')) { continue }
    $dinsComentari = $false
    $n = 0
    foreach ($l in [System.IO.File]::ReadAllLines($f.FullName)) {
        $n++
        # ELS COMENTARIS EN PARLEN, i han de poder parlar-ne: aquesta mateixa
        # seccio de Normativa.ps1 explica per que hi son les banderes i hi
        # escriu el nom del cmdlet. Si el guard comptes el text cru, el propi
        # comentari el faria fallar (hi va caure la primera versio: 4 en lloc
        # de 2). Un sol lector, que salta comentaris, per a les dues preguntes.
        if ($dinsComentari) { if ($l.Contains('#>')) { $dinsComentari = $false }; continue }
        if ($l.TrimStart().StartsWith('<#')) { if (-not $l.Contains('#>')) { $dinsComentari = $true }; continue }
        if ($l.TrimStart().StartsWith('#')) { continue }
        if ($l -match '\b(Invoke-WebRequest|Invoke-RestMethod|System\.Net\.WebClient|System\.Net\.Http\.HttpClient|Start-BitsTransfer)\b') {
            if ($nmHttpPermesos -notcontains $rel) { [void]$nmHttpTrobats.Add($rel + ':' + $n) }
            if (-not $nmHttpPerFitxer.ContainsKey($rel)) { $nmHttpPerFitxer[$rel] = 0 }
            $nmHttpPerFitxer[$rel]++
        }
    }
}
AssertEq ($nmHttpTrobats -join ', ') '' 'cap peticio HTTP fora dels fitxers que son la porta d''un servei'

# I la part que importa de debo: el diagnostic ha de passar per _NormativaGet,
# que es l'unica manera que demani les pagines com les demana el programa.
$nmPv = [System.IO.File]::ReadAllText((Join-Path $nmArrelSup 'ProvarVigencia.ps1'))
Assert ($nmPv.Contains('_NormativaGet $v.Url')) 'ProvarVigencia: el BOE es demana amb _NormativaGet'
# Compte amb la FORMA: una capcalera es passa com a "-Headers @{ ... }", SENSE
# "=" entre el nom i el valor. La primera versio d'aquest assert buscava
# "Headers = @{" i per tant no hauria disparat mai; ho va destapar la injeccio
# del defecte, no la lectura. Ara busca el parametre.
Assert (-not ($nmPv -match '(?m)^[^#]*-Headers\b')) 'ProvarVigencia: cap capcalera passada a ma a una peticio'
# L'Accept es un PARAMETRE de _NormativaGet i nomes el porta l'XML de dades
# obertes. Si algun dia el porten tambe les .html, tornem al defecte.
$nmPvXml = @([regex]::Matches($nmPv, "Accept\s*=\s*'application/xml'")).Count
AssertEq $nmPvXml 1 'ProvarVigencia: l''Accept nomes el porta l''XML de dades obertes (una sola vegada)'
$nmNorm = [System.IO.File]::ReadAllText((Join-Path $nmArrelSup 'Normativa.ps1'))
Assert ($nmNorm.Contains('function _NormativaGet([string]$url, [string]$accept')) '_NormativaGet accepta l''Accept opcional'
# Normativa.ps1 ha de tenir DUES peticions i prou (el GET i el GET a fitxer).
# Abans n'hi havia una TERCERA escrita a ma dins de _NormativaBaixaWeb, que
# nomes es diferenciava pel timeout i per un -PassThru que no es llegia enlloc;
# ara aquell cas passa per _NormativaGetBytes amb el timeout com a parametre.
AssertEq ([int]$nmHttpPerFitxer['Normativa.ps1']) 2 'Normativa.ps1: nomes DUES peticions (el GET i el GET a fitxer), no una copia per cas'
# ...i totes dues amb les banderes. Aqui si que es pot comptar sobre el text
# cru: els comentaris parlen de les banderes pel seu nom, no en la forma
# exacta en que van escrites a la crida.
AssertEq (@([regex]::Matches($nmNorm, '-MaximumRedirection 10 -UseDefaultCredentials')).Count) 2 'Normativa.ps1: les dues peticions porten les dues banderes'

} catch {
    Assert $false ('Normativa: la prova ha petat: ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
}
