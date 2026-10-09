# El repas de contactes d'"Actualitzar base": el lector de text dels PDF, els
# lectors de cada document, les regles (de qui es cada dada), els avisos contra
# l'Excel, les correccions a ma i la base contactes-db.json.
#
# TOTES LES DADES SON INVENTADES (l'usuari: "res de dades reals ni a la suite
# ni a la documentacio"). Domini .test, NIF i telefons de mostra.
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

Write-Host "`n--- PdfText.ps1: el text d'un PDF generat ---"
$dirPdfCt = Join-Path $TestsDir (Join-Path 'dades' 'pdf')
try {
    $txW = Get-PdfText (Join-Path $dirPdfCt 'text-winansi.pdf')
    $linW = @($txW -split "`n" | Where-Object { $_ -ne '' })
    AssertEq $linW[0] 'DADES DE LA PERSONA O ENTITAT INTERESSADA' 'PdfText: la primera linia'
    AssertEq $linW[1] ('Nom o ra' + [char]0x00F3 + ' social: EXEMPLE INVENTAT SL') 'PdfText: etiqueta i valor a la mateixa alcada surten en una linia, per X (el valor anava abans al contingut)'
    AssertEq $linW[2] ('Tel' + [char]0x00E8 + 'fon m' + [char]0x00F2 + 'bil:') 'PdfText: TJ amb ajust petit (kerning) enganxa, i el gran (-300) es un espai; WinAnsi amb accents'
    AssertEq $linW[3] '600000001' "PdfText: l'operador ' passa de linia"
    AssertEq $linW[4] 'Correu: prova@exemple.test' 'PdfText: T* passa de linia; \100 (octal) es @'
    AssertEq $linW[5] ([string][char]0x00C8 + 's') 'PdfText: /Differences (Egrave) de la font'
    AssertEq $linW[6] ('Despla' + [char]0x00E7 + 'at') 'PdfText: la matriu cm (q/Q) mou el text'
    $txU = Get-PdfText (Join-Path $dirPdfCt 'text-tounicode.pdf')
    $linU = @($txU -split "`n" | Where-Object { $_ -ne '' })
    AssertEq ($linU -join '|') ('A' + [char]0x00E0 + '12|x yz|A') 'PdfText: Type0 Identity-H amb ToUnicode (bfchar, bfrange amb inici i amb llista), W i la forma XObject'
    AssertEq (Get-PdfText (Join-Path $dirPdfCt 'escanejat.pdf')) '' 'PdfText: un escanejat (nomes una imatge, i una imatge en linia BI/ID/EI) no dona text'
    AssertEq (_PdfTeText '') $false '_PdfTeText: buit -> no'
    AssertEq (_PdfTeText 'Pagina 1') $false '_PdfTeText: quatre lletres (un numero de pagina) -> no'
    AssertEq (_PdfTeText ($txW)) $true '_PdfTeText: una instancia -> si'
    AssertEq (Get-PdfText (Join-Path $dirPdfCt 'informe.pdf')).Contains('INFORME AJUNTAMENT - pagina 2') $true 'PdfText: dues pagines i el filtre ASCII85 (el de les proves de la unio)'
    $errXif = ''
    try { [void](Get-PdfText (Join-Path $dirPdfCt 'xifrat.pdf')) } catch { $errXif = $_.Exception.Message }
    Assert ($errXif -ne '') 'PdfText: un PDF xifrat llanca (qui crida prova el Word)'
} catch {
    Assert $false ('bloc PdfText: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}

Write-Host "`n--- ContactesExtraccio.ps1: peces ---"
AssertEq (_CtTipusDocument 'tramit_1_etram-tramit.xml') 'etram' '_CtTipusDocument: XML de l''e-TRAM'
AssertEq (_CtTipusDocument ('Inst' + [char]0x00E0 + 'ncia gen' + [char]0x00E8 + 'rica.pdf')) 'instancia' '_CtTipusDocument: instancia generica (amb accents)'
AssertEq (_CtTipusDocument 'Instancia_generica_2.PDF') 'instancia' '_CtTipusDocument: Instancia_generica (majuscules a l''extensio)'
AssertEq (_CtTipusDocument ('Esmena sol' + [char]0x00B7 + 'licitud.pdf')) 'instancia' '_CtTipusDocument: esmena'
AssertEq (_CtTipusDocument 'doc SEU  OAC.pdf') 'instancia' '_CtTipusDocument: SEU  OAC (dos espais)'
AssertEq (_CtTipusDocument ('Autoritzaci' + [char]0x00F3 + ' representaci' + [char]0x00F3 + '.pdf')) 'autoritzacio' '_CtTipusDocument: autoritzacio'
AssertEq (_CtTipusDocument 'autorizacion.pdf') 'autoritzacio' '_CtTipusDocument: autorizacion'
AssertEq (_CtTipusDocument ([char]0x00CD + 'ndex electr' + [char]0x00F2 + 'nic autoritzaci' + [char]0x00F3 + '.pdf')) '' '_CtTipusDocument: l''index electronic no'
AssertEq (_CtTipusDocument 'Memoria tecnica.pdf') '' '_CtTipusDocument: una memoria no es llegeix'
AssertEq (_CtTipusDocument 'XML_TRAMIT_1.xml') 'xmlantic' '_CtTipusDocument: XML antic'
AssertEq (_CtTipusDocument 'instancia generica.docx') '' '_CtTipusDocument: nomes PDF'
AssertEq (_CtTipusNif '12345678Z') 'fisica' '_CtTipusNif: DNI'
AssertEq (_CtTipusNif 'x-1234567-l') 'fisica' '_CtTipusNif: NIE (amb guions i minuscules)'
AssertEq (_CtTipusNif 'B12345674') 'juridica' '_CtTipusNif: NIF d''entitat'
AssertEq (_CtTipusNif '1234') '' '_CtTipusNif: res'
AssertEq (_CtTipusTelefon '+34 600 00 00 01') 'mobil' '_CtTipusTelefon: mobil amb prefix'
AssertEq (_CtTipusTelefon '930000001') 'fix' '_CtTipusTelefon: fix'
AssertEq (_CtTipusTelefon '9300000011') '' '_CtTipusTelefon: 10 xifres no es un telefon'
AssertEq (_CtDataDeText ('Cornell' + [char]0x00E0 + ' de Llobregat, 3 d' + [char]0x2019 + 'abril de 2025')) '2025-04-03' '_CtDataDeText: "3 d''abril de 2025" (apostrof tipografic)'
AssertEq (_CtDataDeText ('Cornell' + [char]0x00E0 + ' de Llobregat, 12 de mar' + [char]0x00E7 + ' de 2024')) '2024-03-12' '_CtDataDeText: marc'
AssertEq (_CtDataDeText 'Data: 05/06/2024') '2024-06-05' '_CtDataDeText: dd/mm/aaaa'
AssertEq (_CtDataDeText '5 de mayo de 2023') '2023-05-05' '_CtDataDeText: en castella'
AssertEq (_CtMateixNom 'Maria Exemple' 'MARIA EXEMPLE PROVA') $true '_CtMateixNom: les paraules del curt son al llarg'
AssertEq (_CtMateixNom 'Maria' 'Maria Exemple') $false '_CtMateixNom: una sola paraula no n''hi ha prou'
AssertEq (_CtMateixNom 'EXEMPLE INVENTAT, S.L.' 'Exemple Inventat SL') $true '_CtMateixNom: sense la forma juridica ni els signes'
AssertEq (_CtNetejaNom "Sr. Pere Tecnic Inventat, major d'edat") 'Pere Tecnic Inventat' '_CtNetejaNom: sense tractament ni "major d''edat"'
AssertEq (_CtEsQueixa '2579/2025/12') $true '_CtEsQueixa: serie 2579'
AssertEq (_CtEsQueixa '2025/1/2563') $false '_CtEsQueixa: una altra serie'
$pTel = _CtPersona 'X' '' '' '600000001' '930000002'
AssertEq "$($pTel.telefon)|$($pTel.mobil)" '930000002|600000001' '_CtPersona: els telefons van al camp que diu el prefix, no al que diu el formulari'

Write-Host "`n--- ContactesExtraccio.ps1: els tres documents ---"
$xmlCt = @"
<?xml version="1.0" encoding="UTF-8"?>
<tramit xmlns="http://www.aoc.cat/etram"><dataCreacio>2025-03-14T10:11:12</dataCreacio><nomTramit>Comunicacio previa</nomTramit>
<solicitant><tipusDocument>NIF</tipusDocument><numeroDocument>B12345674</numeroDocument><raoSocial>EXEMPLE INVENTAT SL</raoSocial><telefon>930000001</telefon><correuElectronic>Info@Exemple-Inventat.test</correuElectronic></solicitant>
<representant><nom>Pere</nom><cognom1>Tecnic</cognom1><cognom2>Inventat</cognom2><numeroDocument>12345678Z</numeroDocument><telefonMobil>600 000 002</telefonMobil><correuElectronic>pere@enginyeria-inventada.test</correuElectronic></representant>
<canalSms><telefon>+34600000002</telefon></canalSms>
<altresDades><dada clau="Tel$([char]0x00E8)fon establiment">930000009</dada><dada><clau>email</clau><valor>botiga@exemple-inventat.test</valor></dada><dada clau="comercialName" valor="La Botiga Inventada"/></altresDades>
</tramit>
"@
$et = Read-EtramXml $xmlCt
AssertEq "$($et.data)|$($et.interessat.nom)|$($et.interessat.nif)|$($et.interessat.email)|$($et.interessat.telefon)" '2025-03-14|EXEMPLE INVENTAT SL|B12345674|info@exemple-inventat.test|930000001' 'e-TRAM: el sol·licitant (raoSocial, numeroDocument, correu en minuscules) i la dataCreacio'
AssertEq "$($et.representant.nom)|$($et.representant.mobil)" 'Pere Tecnic Inventat|600000002' 'e-TRAM: el representant (nom + cognoms, el mobil net)'
AssertEq "$($et.establiment.nom_comercial)|$($et.establiment.telefon)|$($et.establiment.email)" 'La Botiga Inventada|930000009|botiga@exemple-inventat.test' 'e-TRAM: altresDades (atribut, fills clau/valor i atributs clau/valor)'
AssertEq (Read-EtramXml '<no es xml') $null 'e-TRAM: un XML trencat torna $null (no peta)'
$xa = Read-XmlAntic '<r><Nombre_Interesado>JOAN</Nombre_Interesado><Apellido1_Interesado>PROVA</Apellido1_Interesado><Documento_Identificacion_Interesado>11111111H</Documento_Identificacion_Interesado><Telefono_Contacto_Interesado>611111111</Telefono_Contacto_Interesado><Nombre_Representante>ANNA TRAMITS</Nombre_Representante><Fecha_Registro>02/01/2020</Fecha_Registro></r>'
AssertEq "$($xa.interessat.nom)|$($xa.interessat.nif)|$($xa.interessat.mobil)|$($xa.representant.nom)|$($xa.data)" 'JOAN PROVA|11111111H|611111111|ANNA TRAMITS|2020-01-02' 'XML antic: _Interesado i _Representante'

$instCt = @"
Inst$([char]0x00E0)ncia gen$([char]0x00E8)rica - Exp. 2025/1/999 - GIA 9999
DADES DE LA PERSONA O ENTITAT INTERESSADA
Nom o ra$([char]0x00F3) social: EXEMPLE INVENTAT SL
DNI/NIF/NIE: B12345674
Adre$([char]0x00E7)a: Carrer Inventat 1
Tel$([char]0x00E8)fon fix: 930000001 M$([char]0x00F2)bil: 600000003
Correu electr$([char]0x00F2)nic
info@exemple-inventat.test
DADES DE LA PERSONA REPRESENTANT
Nom i cognoms
Pere Tecnic Inventat
DNI/NIF/NIE 12345678Z
Correu electr$([char]0x00F2)nic: pere@enginyeria-inventada.test
EXPOSO
Que demano...
Cornell$([char]0x00E0) de Llobregat, 3 d'abril de 2025
"@
$in = Read-InstanciaText $instCt
AssertEq "$($in.interessat.nom)|$($in.interessat.nif)|$($in.interessat.telefon)|$($in.interessat.mobil)|$($in.interessat.email)" 'EXEMPLE INVENTAT SL|B12345674|930000001|600000003|info@exemple-inventat.test' 'instancia: l''interessat (dos telefons a la mateixa linia; el correu a la linia de sota)'
AssertEq "$($in.representant.nom)|$($in.representant.nif)|$($in.representant.email)" 'Pere Tecnic Inventat|12345678Z|pere@enginyeria-inventada.test' 'instancia: el representant (nom a sota de l''etiqueta, NIF sense dos punts)'
AssertEq "$($in.data)|$($in.gia)|$($in.expedient)" '2025-04-03|9999|2025/1/999' 'instancia: la data, el GIA i l''expedient del titol'

$autCt = "AUTORITZACI$([char]0x00D3) DE REPRESENTACI$([char]0x00D3)`nJo, Maria Exemple Prova, major d'edat, amb DNI 87654321X, en nom i representaci$([char]0x00F3) de l'empresa EXEMPLE INVENTAT SL, amb NIF B12345674, AUTORITZO a Pere Tecnic Inventat, amb DNI 12345678Z i correu pere@enginyeria-inventada.test, perqu$([char]0x00E8) presenti la documentaci$([char]0x00F3).`nCornell$([char]0x00E0) de Llobregat, 2 de febrer de 2025"
$au = Read-AutoritzacioText $autCt
AssertEq "$($au.autoritzacio.signant.nom)|$($au.autoritzacio.signant.nif)|$($au.autoritzacio.empresa.nom)|$($au.autoritzacio.empresa.nif)" 'Maria Exemple Prova|87654321X|EXEMPLE INVENTAT SL|B12345674' 'autoritzacio (ca): qui signa i en nom de quina empresa'
AssertEq "$(@($au.autoritzacio.autoritzats)[0].nom)|$(@($au.autoritzacio.autoritzats)[0].nif)|$(@($au.autoritzacio.autoritzats)[0].email)|$($au.data)" 'Pere Tecnic Inventat|12345678Z|pere@enginyeria-inventada.test|2025-02-02' 'autoritzacio (ca): a qui autoritza, i la data'
$au2 = Read-AutoritzacioText "D./D$([char]0x00F1)a. Joan Exemple Segon, con DNI 11111111H, como administrador $([char]0x00FA)nico de la sociedad PROVES INVENTADES SA, con CIF A87654321, autoriza a D. Pere Tecnic Inventat, con DNI 12345678Z, para tramitar."
AssertEq "$($au2.autoritzacio.signant.nom)|$($au2.autoritzacio.empresa.nom)|$($au2.autoritzacio.empresa.nif)|$(@($au2.autoritzacio.autoritzats)[0].nom)" 'Joan Exemple Segon|PROVES INVENTADES SA|A87654321|Pere Tecnic Inventat' 'autoritzacio (es): "como administrador unico de ... autoriza a"'
$au3 = Read-AutoritzacioText 'Jo, Anna Fisica Prova, amb DNI 22222222J, AUTORITZO a Pere Tecnic Inventat perque tramiti.'
AssertEq "$($au3.autoritzacio.en_nom_propi)|$($au3.interessat.nom)|$($au3.autoritzacio.empresa)" 'True|Anna Fisica Prova|' 'autoritzacio en nom propi: qui signa es el titular'
AssertEq (Read-AutoritzacioText 'Un text que no autoritza res.') $null 'autoritzacio: sense "autoritzo" -> $null'

Write-Host "`n--- ContactesRegles.ps1: les regles ---"
# Un document com el desa la base: @{ ruta; modificat; extret }.
$ctDoc = {
    param([string]$ruta, [string]$tipus, $int, $rep, [string]$data, $aut = $null, [string]$exp = '', $est = $null)
    $e = _CtDocNou $tipus
    $e.interessat = $int; $e.representant = $rep; $e.data = $data; $e.autoritzacio = $aut; $e.expedient = $exp; $e.establiment = $est
    return @{ ruta = $ruta; modificat = ($data + 'T00:00:00Z'); extret = $e }
}
$pereT = _CtPersona 'Pere Tecnic Inventat' '12345678Z' 'pere.prova@correu.test' '' '600000002'
$docsCt = @{
    # 9001 i 9002: dos titulars diferents, el MATEIX representant (correu sense
    # cap paraula professional): el senyal es que surt a titulars diferents.
    '9001' = @((& $ctDoc 'GIA 9001\Instancia_generica.pdf' 'instancia' (_CtPersona 'EXEMPLE U SL' 'B11111118' 'info@exemple-u.test' '930000011') $pereT '2025-01-10'))
    '9002' = @((& $ctDoc 'GIA 9002\Instancia_generica.pdf' 'instancia' (_CtPersona 'EXEMPLE DOS SL' 'B22222226' 'pere.prova@correu.test' '' '600000002') $pereT '2025-02-10'))
    # 9003: persona fisica amb el tecnic de representant.
    '9003' = @((& $ctDoc 'GIA 9003\etram-tramit.xml' 'etram' (_CtPersona 'Anna Fisica Prova' '22222222J' 'anna@correu.test' '' '611111112') $pereT '2025-03-10'))
    # 9004: juridica amb autoritzacio de l'administradora i el tecnic al tramit.
    '9004' = @(
        (& $ctDoc 'GIA 9004\Autoritzacio.pdf' 'autoritzacio' (_CtPersona 'PROVES QUATRE SL' 'B44444442') $null '2025-01-05' ([ordered]@{ signant = (_CtPersona 'Maria Admin Prova' '33333333P'); empresa = (_CtPersona 'PROVES QUATRE SL' 'B44444442'); autoritzats = @((_CtPersona 'Jordi Enginyer Inventat' '44444444A' 'jordi@enginyeria-inventada.test')); en_nom_propi = $false })),
        (& $ctDoc 'GIA 9004\etram-tramit.xml' 'etram' (_CtPersona 'PROVES QUATRE SL' 'B44444442' 'jordi@enginyeria-inventada.test' '930000044') (_CtPersona 'Jordi Enginyer Inventat' '44444444A' 'jordi@enginyeria-inventada.test' '' '622222222') '2025-06-01' $null '' ([ordered]@{ nom_comercial = 'Bar Proves'; telefon = '930000045'; email = 'jordi@enginyeria-inventada.test' })))
    # 9005: el titular i una QUEIXA (serie 2579) d'un vei.
    '9005' = @(
        (& $ctDoc 'GIA 9005\Instancia_generica.pdf' 'instancia' (_CtPersona 'PROVES CINC SL' 'B55555553' 'cinc@proves.test') $null '2024-05-01'),
        (& $ctDoc 'GIA 9005\Instancia_generica 2.pdf' 'instancia' (_CtPersona 'Vei Queixos Prova' '55555555K' 'vei@correu.test' '' '633333333') $null '2025-07-01' $null '2579/2025/7'))
    # 28: un recinte (organitzadors d'actes).
    '28'   = @((& $ctDoc 'GIA 28\Instancia_generica.pdf' 'instancia' (_CtPersona 'ORGANITZADORA PROVA SL' 'B66666664' 'org@proves.test') $null '2025-08-01'))
}
$xlCt = @{
    '9001' = @{ TITULAR = 'EXEMPLE U SL'; NIF = 'B11111118'; EMAIL = 'exempleu@gmailo.com'; MOBIL = '930000011'; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    '9002' = @{ TITULAR = 'EXEMPLE DOS SL'; NIF = 'B22222226'; EMAIL = 'pere.prova@correu.test'; MOBIL = ''; TELEFON = ''; REP_NOM = 'Pere Tecnic Inventat'; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    '9003' = @{ TITULAR = 'Anna Fisica Prova'; NIF = '22222222J'; EMAIL = ''; MOBIL = '611111112'; TELEFON = ''; REP_NOM = 'Pere Tecnic Inventat'; REP_NIF = '12345678Z'; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    '9004' = @{ TITULAR = 'PROVES QUATRE SL'; NIF = 'B44444442'; EMAIL = ''; MOBIL = ''; TELEFON = '930000044'; REP_NOM = 'Jordi Enginyer Inventat'; REP_NIF = 'B44444442'; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    '9005' = @{ TITULAR = 'PROVES CINC SL'; NIF = 'B55555553'; EMAIL = 'cinc@proves.test'; MOBIL = ''; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
    '28'   = @{ TITULAR = 'RECINTE PROVA'; NIF = 'P0800000B'; EMAIL = ''; MOBIL = ''; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
}
try {
    $ctxCt = _CtContext $docsCt $xlCt @{ Emails = @{}; Noms = @() } @('28', '1324')
    AssertEq $ctxCt.EmailTitulars['pere.prova@correu.test'].Count 3 'context: el correu del tecnic surt a 3 activitats de titulars diferents (comptat sobre tota la base, Excel inclos)'
    $a1 = Get-ContactesActivitat '9001' $docsCt['9001'] $xlCt['9001'] $ctxCt
    $a2 = Get-ContactesActivitat '9002' $docsCt['9002'] $xlCt['9002'] $ctxCt
    AssertEq "$(@($a1.persones_autoritzades)[0].nom)|$($a1.representant_legal)" 'Pere Tecnic Inventat|' 'tecnic repetit: es persona autoritzada, no el representant legal'
    Assert (([string]@($a1.persones_autoritzades)[0].motiu).Contains('titulars diferents')) 'tecnic repetit: el motiu diu que surt a titulars diferents'
    AssertEq "$($a2.titular.nom)|$($a2.titular.email)|$($a2.titular.mobil)" 'EXEMPLE DOS SL||' 'regla 3: el correu i el mobil del tecnic posats a l''interessat no son del titular'
    $av2 = @($a2.avisos | Where-Object { $_.tipus -eq 'es_el_tecnic' } | ForEach-Object { $_.camp }) -join '|'
    AssertEq $av2 'E-mail|Representant legal' 'avisos: l''e-mail i el representant legal de l''Excel son del tecnic (es_el_tecnic)'
    $av1 = @($a1.avisos | ForEach-Object { $_.tipus + ':' + $_.camp })
    Assert ($av1 -contains 'error:E-mail') 'avisos: domini mal escrit (gmailo.com) -> error'
    Assert ($av1 -contains ('error:M' + [char]0x00F2 + 'bil')) 'avisos: un fix a la columna del mobil -> error'
    Assert ($av1 -contains ('falta:Tel' + [char]0x00E8 + 'fon')) 'avisos: l''Excel no te el telefon i els documents si -> falta'

    $a3 = Get-ContactesActivitat '9003' $docsCt['9003'] $xlCt['9003'] $ctxCt
    AssertEq "$($a3.titular_tipus)|$($a3.representant_legal)|$($a3.titular.email)" 'fisica||anna@correu.test' 'persona fisica: el representant legal en blanc'
    $av3 = @($a3.avisos | Where-Object { $_.tipus -eq 'es_el_tecnic' })
    AssertEq (@($av3 | ForEach-Object { $_.camp }) -join '|') 'Representant legal|Rep. Leg. NIF' 'persona fisica amb el tecnic de representant a l''Excel: es_el_tecnic al nom i al NIF'
    AssertEq ([string]$av3[0].proposta) '' '...i la proposta es deixar-lo en blanc'

    $a4 = Get-ContactesActivitat '9004' $docsCt['9004'] $xlCt['9004'] $ctxCt
    AssertEq "$($a4.representant_legal.nom)|$($a4.representant_legal.confianca)" 'Maria Admin Prova|segur' 'juridica: el representant legal es qui signa l''autoritzacio (segur)'
    AssertEq "$($a4.titular.email)|$($a4.titular.telefon)" '|930000044' 'juridica: el correu del tecnic a l''interessat no passa al titular'
    AssertEq "$($a4.establiment.nom_comercial)|$($a4.establiment.telefon)|$($a4.establiment.email)" 'Bar Proves|930000045|' 'establiment: el de l''e-TRAM, sense el correu del tecnic'
    $av4 = @($a4.avisos | ForEach-Object { $_.tipus + ':' + $_.camp })
    Assert ($av4 -contains 'es_el_tecnic:Representant legal') 'juridica: el tecnic posat de representant legal a l''Excel -> es_el_tecnic'
    Assert ($av4 -contains 'error:Rep. Leg. NIF') 'avisos: el CIF de l''empresa al NIF del representant -> error'

    $a5 = Get-ContactesActivitat '9005' $docsCt['9005'] $xlCt['9005'] $ctxCt
    AssertEq "$($a5.titular.nom)|$($a5.titular.mobil)|$(@($a5.avisos).Count)" 'PROVES CINC SL||0' 'queixa (2579): el vei no es el titular (ni canvi de titular ni avisos)'
    $a28 = Get-ContactesActivitat '28' $docsCt['28'] $xlCt['28'] $ctxCt
    AssertEq "$($a28.titular)|$(@($a28.avisos).Count)" '|0' 'recinte (GIA 28): no es proposa cap titular ni cap avis'
    Assert ([string]$a28.notes).Contains('organitzadors') 'recinte: la nota diu per que'

    # Canvi de titular: el document mes recent es d'un altre NIF.
    $docsCanvi = @((& $ctDoc 'GIA 9006\a.pdf' 'instancia' (_CtPersona 'VELL SL' 'B77777775') $null '2023-01-01'), (& $ctDoc 'GIA 9006\b.pdf' 'instancia' (_CtPersona 'NOU SL' 'B88888883') $null '2025-01-01'))
    $a6 = Get-ContactesActivitat '9006' $docsCanvi @{ TITULAR = 'VELL SL'; NIF = 'B77777775' } $ctxCt
    AssertEq (@($a6.avisos | ForEach-Object { $_.tipus }) -join '|') 'canvi_titular' 'canvi_titular: nomes avis (l''Excel pot ser mes nou)'
    AssertEq "$($a6.titular.nom)|$($a6.canvi_titular.nif)" 'VELL SL|B88888883' 'canvi_titular: el titular segueix sent el de l''Excel, i es diu qui diuen els documents'

    # Tecnics coneguts (la llista local).
    $ctxTec = _CtContext @{} @{} (_CtTecnicsDe ([pscustomobject]@{ tecnics = @([pscustomobject]@{ email = 'Gestoria@Proves.test'; noms = @('Lluis Gestor Prova'); empresa = @('Gestoria Proves'); rol = 'gestoria' }) })) @()
    AssertEq (_CtMotiuTecnic @{ email = 'gestoria@proves.test' } $ctxTec) "es a la llista de t$([char]0x00E8)cnics coneguts" 'tecnics coneguts: pel correu'
    AssertEq (_CtMotiuTecnic @{ nom = 'LLUIS GESTOR PROVA' } $ctxTec) "es a la llista de t$([char]0x00E8)cnics coneguts" 'tecnics coneguts: pel nom'
    AssertEq (_CtMotiuTecnic @{ email = 'info@estudi-dansa.test' } $ctxTec -Fort) '' '-Fort: un correu "professional" no treu el correu al titular (un estudi de dansa)'
    AssertEq (_CtMotiuTecnic @{ email = 'info@estudi-dansa.test' } $ctxTec) 'correu de professional' 'sense -Fort: el patro professional si que compta (per al representant d''un tramit)'
    AssertEq (_CtEmailExclos 'inspeccio@aj-cornella.cat') $true 'regla 5: un correu de l''Ajuntament no es mai del titular'
    AssertEq (_CtEmailExclos 'oca@tuvsud.test') $true 'regla 5: una entitat de control tampoc'
    AssertEq (_CtErrada 'email' 'joan@hotmal.com' '') 'domini mal escrit (hotmal.com)' '_CtErrada: hotmal'
    AssertEq (_CtErrada 'email' 'joan@gmial.com' '') 'domini mal escrit (gmial.com)' '_CtErrada: gmial'
    AssertEq (_CtErrada 'email' 'joan@prova.tesst' 'joan@prova.test') 'domini mal escrit (prova.tesst en lloc de prova.test)' '_CtErrada: a una lletra del dels documents'
    AssertEq (_CtErrada 'telefon' '9300000011' '') "el tel$([char]0x00E8)fon t$([char]0x00E9) 10 xifres" '_CtErrada: 10 xifres'
    AssertEq (_CtErrada 'telefon' 'a@b.test' '') "hi ha un correu a la columna del tel$([char]0x00E8)fon" '_CtErrada: un correu a la columna del telefon'
    AssertEq (_CtErrada 'nif' '12345678A' '') 'la lletra del DNI no quadra' '_CtErrada: la lletra del DNI'
    AssertEq (_CtErrada 'nif' '12345678Z' '') '' '_CtErrada: un DNI bo'

    Write-Host "`n--- ContactesRegles.ps1: les correccions a ma manen ---"
    # Marcar l'administradora (de l'autoritzacio) com a tecnic: passa a
    # autoritzada i el representant queda en blanc.
    $manT = _CtManDeCorreccions @([ordered]@{ tipus = 'es_tecnic'; claus = @(_CtClausPersona (_CtPersona 'Maria Admin Prova' '33333333P')); persona = (_CtPersona 'Maria Admin Prova' '33333333P') })
    $a4t = Get-ContactesActivitat '9004' $docsCt['9004'] $xlCt['9004'] $ctxCt $manT
    AssertEq "$($a4t.representant_legal)|$((@($a4t.persones_autoritzades | ForEach-Object { $_.nom }) -join ','))|$($a4t.editat_a_ma)" '|Jordi Enginyer Inventat,Maria Admin Prova|True' 'es_tecnic: la persona marcada passa a autoritzada i deixa de ser representant legal'
    # Marcar el tecnic com a representant legal (l'usuari ho sap millor).
    $manR = _CtManDeCorreccions @([ordered]@{ tipus = 'es_rep_legal'; claus = @(_CtClausPersona $pereT); persona = $pereT })
    $a3r = Get-ContactesActivitat '9003' $docsCt['9003'] $xlCt['9003'] $ctxCt $manR
    AssertEq "$($a3r.representant_legal.nom)|$($a3r.representant_legal.confianca)|$(@($a3r.persones_autoritzades).Count)" ('Pere Tecnic Inventat|a m' + [char]0x00E0 + '|0') 'es_rep_legal: mana sobre la regla (fins i tot amb persona fisica) i surt dels autoritzats'
    AssertEq (@($a3r.avisos | Where-Object { $_.tipus -eq 'es_el_tecnic' }).Count) 0 '...i ja no hi ha avis es_el_tecnic del representant'
    # Descartar un avis i editar una dada.
    $idDesc = [string](@($a1.avisos | Where-Object { $_.tipus -eq 'error' -and $_.camp -eq 'E-mail' })[0].id)
    $manE = _CtManDeCorreccions @([ordered]@{ tipus = 'descarta'; avis = $idDesc }, [ordered]@{ tipus = 'edita'; camp = 'titular.telefon'; valor = '930000099' })
    $a1e = Get-ContactesActivitat '9001' $docsCt['9001'] $xlCt['9001'] $ctxCt $manE
    AssertEq (@($a1e.avisos | Where-Object { $_.id -eq $idDesc }).Count) 0 'descarta: l''avis descartat no torna a sortir'
    AssertEq "$($a1e.avisos_descartats)|$($a1e.titular.telefon)|$(@($a1e.correccions | Where-Object { $_.tipus -eq 'edita' })[0].auto)" '1|930000099|930000011' 'edita: la dada corregida mana i es guarda el valor automatic per desfer-ho'
    AssertEq ([string]@($a1e.avisos | Where-Object { $_.tipus -eq 'falta' -and $_.camp -eq ('Tel' + [char]0x00E8 + 'fon') })[0].proposta) '930000099' 'edita: els avisos es comparen amb el valor corregit'
} catch {
    Assert $false ('bloc regles de contactes: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}

Write-Host "`n--- ContactesDb.ps1: el repas sencer (carpeta temporal) ---"
# Un PDF de text, fet aqui (Helvetica, WinAnsi, sense comprimir).
function _CtProvaPdf([string]$path, [string[]]$linies) {
    $enc = [System.Text.Encoding]::GetEncoding(28591)
    $c = New-Object System.Text.StringBuilder
    $y = 800
    foreach ($l in $linies) {
        $esc = ($l -replace '\\', '\\\\' -replace '\(', '\(' -replace '\)', '\)')
        [void]$c.Append("BT /F1 10 Tf 50 $y Td ($esc) Tj ET`n"); $y -= 14
    }
    $cont = $enc.GetBytes($c.ToString())
    $objs = @('<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
              '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
              $null, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>')
    $ms = New-Object System.IO.MemoryStream
    $w = { param($t) $b = $enc.GetBytes($t); $ms.Write($b, 0, $b.Length) }
    & $w "%PDF-1.4`n"
    $offs = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offs += $ms.Position
        & $w ("" + ($i + 1) + " 0 obj`n")
        if ($null -eq $objs[$i]) { & $w ("<< /Length " + $cont.Length + " >>`nstream`n"); $ms.Write($cont, 0, $cont.Length); & $w "`nendstream" } else { & $w $objs[$i] }
        & $w "`nendobj`n"
    }
    $x = $ms.Position
    & $w ("xref`n0 " + ($objs.Count + 1) + "`n0000000000 65535 f `n")
    foreach ($o in $offs) { & $w (('{0:D10}' -f $o) + " 00000 n `n") }
    & $w ("trailer`n<< /Size " + ($objs.Count + 1) + " /Root 1 0 R >>`nstartxref`n$x`n%%EOF`n")
    [System.IO.File]::WriteAllBytes($path, $ms.ToArray())
}
$ctT = Join-Path ([System.IO.Path]::GetTempPath()) ('contactes-' + [guid]::NewGuid().ToString('N'))
$ctI = Join-Path $ctT 'Informes'
$vellsCt = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
try {
    foreach ($d in @((Join-Path $ctI 'GIA 9101'), (Join-Path $ctI 'Exp 2025-1-77'), (Join-Path (Join-Path $ctI 'GIA 9101') 'Expedient'), (Join-Path $ctT 'local'), (Join-Path $ctT 'act'), (Join-Path $ctT 'app'))) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $LocalActivitatsDir = Join-Path $ctT 'local'; $ActivitatsDir = Join-Path $ctT 'act'; $env:LOCALAPPDATA = Join-Path $ctT 'app'; $InformesDir = $ctI
    _CtOblidaMemo
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $ctI 'GIA 9101') 'tramit_etram-tramit.xml'), $xmlCt, (New-Object System.Text.UTF8Encoding($false)))
    _CtProvaPdf (Join-Path (Join-Path $ctI 'GIA 9101') 'Instancia_generica.pdf') @($instCt -split "`r?`n")
    Copy-Item -LiteralPath (Join-Path $dirPdfCt 'escanejat.pdf') -Destination (Join-Path (Join-Path $ctI 'GIA 9101') 'Autoritzacio escanejada.pdf')
    # Una subcarpeta i l'arrel: no s'hi llegeix res (primer nivell).
    _CtProvaPdf (Join-Path (Join-Path (Join-Path $ctI 'GIA 9101') 'Expedient') 'Instancia_generica.pdf') @('DADES DE LA PERSONA O ENTITAT INTERESSADA', 'Nom o rao social: NO HI HA DE SER SL', 'DNI/NIF/NIE: B99999990')
    _CtProvaPdf (Join-Path $ctI 'Instancia_generica.pdf') @('DADES DE LA PERSONA O ENTITAT INTERESSADA', 'Nom o rao social: TAMPOC SL')
    # La carpeta sense "GIA" al nom: el GIA es el dels informes que hi ha.
    _CtProvaPdf (Join-Path (Join-Path $ctI 'Exp 2025-1-77') 'Instancia_generica.pdf') @('Instancia generica', 'DADES DE LA PERSONA O ENTITAT INTERESSADA', 'Nom o rao social: EXPEDIENT SET SL', 'DNI/NIF/NIE: B70000007', 'Correu: set@proves.test', 'Cornella de Llobregat, 1 de juny de 2025')
    $informesCt = @([pscustomobject]@{ Ruta = (Join-Path (Join-Path $ctI 'Exp 2025-1-77') '2025-06-01_Req.docx'); Gia = '9102' })
    $cacheCt = [pscustomobject]@{ ById = @{}; Contactes = @{ '9101' = @{ TITULAR = 'EXEMPLE INVENTAT SL'; NIF = 'B12345674'; EMAIL = 'pere@enginyeria-inventada.test'; MOBIL = ''; TELEFON = ''; REP_NOM = ''; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' } } }

    $rC = Invoke-ContactesEscaneig $ctI $informesCt $cacheCt @{}
    AssertEq "$($rC.Ok)|$($rC.NDocs)|$($rC.Llegits)|$($rC.NNoLlegibles)" 'True|4|4|1' 'repas: nomes els 4 documents del primer nivell; l''escanejat, a "no llegibles"'
    $dbC = Read-ContactesDb
    AssertEq ((@($dbC.activitats.PSObject.Properties.Name) | Sort-Object) -join ',') '9101,9102' 'repas: el GIA de la carpeta "GIA n" i el dels informes de la carpeta sense GIA al nom'
    $c9101 = $dbC.activitats.'9101'
    AssertEq "$($c9101.titular.nom)|$($c9101.titular.email)|$(@($c9101.persones_autoritzades)[0].nom)" 'EXEMPLE INVENTAT SL|info@exemple-inventat.test|Pere Tecnic Inventat' 'repas: el titular (sense el correu del tecnic) i la persona autoritzada'
    Assert (@($c9101.avisos | ForEach-Object { $_.tipus + ':' + $_.camp }) -contains 'es_el_tecnic:E-mail') 'repas: el correu de l''Excel es del tecnic -> avis'
    AssertEq ([string]$c9101.excel.TITULAR) 'EXEMPLE INVENTAT SL' 'repas: la fila de l''Excel es desa al costat (la finestra la ensenya)'
    Assert ($rC.Text.StartsWith('Contactes: 2 activitats amb dades noves')) ('repas: el resum de la linia de "Actualitzar base" (' + $rC.Text + ')')
    Assert ($rC.Text.Contains('1 documents no llegibles')) 'repas: el resum diu els no llegibles'

    # INCREMENTAL: la segona passada no torna a llegir res; si un canvia, nomes aquell.
    $rC2 = Invoke-ContactesEscaneig $ctI $informesCt $cacheCt @{}
    AssertEq "$($rC2.Llegits)|$($rC2.NNoves)" '0|0' 'repas incremental: res de nou -> no es llegeix cap document'
    (Get-Item -LiteralPath (Join-Path (Join-Path $ctI 'GIA 9101') 'tramit_etram-tramit.xml')).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddMinutes(5)
    $rC3 = Invoke-ContactesEscaneig $ctI $informesCt $cacheCt @{}
    AssertEq $rC3.Llegits 1 'repas incremental: nomes el document modificat'

    # LES CORRECCIONS A MA MANEN despres d'un altre "Actualitzar base".
    Assert (Set-ContactesCorreccio '9101' ([ordered]@{ tipus = 'edita'; camp = 'titular.email'; valor = 'correcte@exemple-inventat.test' })) 'Set-ContactesCorreccio: desa'
    $rC4 = Invoke-ContactesEscaneig $ctI $informesCt $cacheCt @{}
    $dbC4 = Read-ContactesDb
    AssertEq "$($dbC4.activitats.'9101'.titular.email)|$($dbC4.activitats.'9101'.editat_a_ma)" 'correcte@exemple-inventat.test|True' 'correccions: el proxim repas no desfa la correccio a ma'
    AssertEq ([string]@($dbC4.activitats.'9101'.correccions)[0].auto) 'info@exemple-inventat.test' 'correccions: es guarda el valor automatic al costat (per "Desfer canvi a ma")'
    $rSense = Invoke-ContactesEscaneig $ctI $informesCt $cacheCt @{} -SenseCorreccions -NoDesis
    AssertEq ([string]$rSense.Activitats['9101'].titular.email) 'info@exemple-inventat.test' '-SenseCorreccions (ValidarContactes): el que diu el repas sol'
    AssertEq ([string](Read-ContactesDb).activitats.'9101'.titular.email) 'correcte@exemple-inventat.test' '-NoDesis: la base no es toca'
    Assert (Set-ContactesCorreccio '9101' $null 0) 'Set-ContactesCorreccio: treu (Desfer)'
    AssertEq ([string](Read-ContactesDb).activitats.'9101'.titular.email) 'info@exemple-inventat.test' 'Desfer canvi a ma: torna l''automatic sense tornar a llegir res'

    # Fer-la servir als correus.
    _CtOblidaMemo
    AssertEq (@(Get-ContactesAutoritzatsEmails '9101') -join ',') 'pere@enginyeria-inventada.test' 'Get-ContactesAutoritzatsEmails: els correus de les persones autoritzades'
    $cmp = Get-ContactesEmailsCompletats '9101' @{ titular = ''; representant = '' }
    AssertEq "$($cmp.Emails.titular)|$(@($cmp.Notes).Count)" 'info@exemple-inventat.test|1' 'correus: si l''Excel es buit, el dels documents, i es diu que surt dels documents'
    $cmpT = Get-ContactesEmailsCompletats '9101' @{ titular = 'pere@enginyeria-inventada.test'; representant = '' }
    Assert ([string]$cmpT.AvisTecnic -ne '') 'correus: el correu de l''Excel es del tecnic -> avis abans d''enviar'
    AssertEq ([string](Get-ContactesEmailsCompletats '1' @{ titular = 'a@b.test'; representant = '' }).Emails.titular) 'a@b.test' 'correus: una activitat que no es a la base es queda com estava'

    # Exportar les correccions.
    $xlsC = Join-Path $ctT 'correccions.xlsx'
    [void](Export-ContactesCorreccions $xlsC (_CtMapaDe (Read-ContactesDb).activitats))
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zC = [System.IO.Compression.ZipFile]::OpenRead($xlsC)
    try { $shC = (New-Object System.IO.StreamReader($zC.GetEntry('xl/worksheets/sheet1.xml').Open())).ReadToEnd() } finally { $zC.Dispose() }
    Assert ($shC.Contains('es_el_tecnic') -and $shC.Contains('9101') -and $shC.Contains("Valor a l'Excel")) 'Exportar: un .xlsx amb la capcalera i un avis per fila'

    # Dins d'"Actualitzar base": el repas surt al resultat.
    $resIC = Invoke-InformesDbEscaneig
    Assert ($null -ne $resIC.Contactes -and [string]$resIC.Contactes.Text -like 'Contactes:*') 'Actualitzar base: porta el repas de contactes al resultat'
    $resIC2 = Invoke-InformesDbEscaneig $null $null '' $false
    AssertEq $resIC2.Contactes $null 'Actualitzar base sense contactes (ValidarClassificacio): no el fa'
} catch {
    Assert $false ('bloc repas de contactes: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    $InformesDir = $vellsCt.Inf; $LocalActivitatsDir = $vellsCt.Loc; $ActivitatsDir = $vellsCt.Act; $env:LOCALAPPDATA = $vellsCt.App
    _CtOblidaMemo
    Remove-Item -LiteralPath $ctT -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- ContactesPantalla.ps1: el que decideixen els botons (pur) ---"
$actP = [ordered]@{
    titular = [ordered]@{ nom = 'EXEMPLE SL'; nif = 'B12345674'; email = 'info@exemple.test'; telefon = ''; mobil = ''; font = 'GIA 1\a.pdf'; data = '2025-01-01'; confianca = 'segur' }
    representant_legal = $null; establiment = $null
    persones_autoritzades = @([ordered]@{ nom = 'Pere Tecnic Inventat'; rol = 'tecnic'; empresa = ''; nif = ''; email = 'pere@x.test'; telefon = ''; mobil = ''; font = 'GIA 1\a.pdf'; data = ''; confianca = 'segur'; motiu = 'x' })
    avisos = @([ordered]@{ id = 'es_el_tecnic|e-mail|pere@x.test'; tipus = 'es_el_tecnic'; camp = 'E-mail'; valor_excel = 'pere@x.test' })
    editat_a_ma = $false; correccions = @()
    excel = @{ TITULAR = 'EXEMPLE SL'; EMAIL = 'pere@x.test'; REP_NOM = 'Pere Tecnic Inventat'; REP_NIF = ''; REP_EMAIL = ''; REP_TELEFON = ''; REP_MOBIL = '' }
}
$filesP = @(_CtFilesDetall $actP)
AssertEq $filesP.Count 14 'detall: 5 del titular, 5 del representant, 3 de l''establiment i 1 persona autoritzada'
$fEmail = @($filesP | Where-Object { $_.Qui -eq 'titular' -and $_.Dada -eq 'email' })[0]
AssertEq "$($fEmail.Excel)|$($fEmail.Documents)" 'pere@x.test|info@exemple.test' 'detall: l''Excel i els documents de costat'
AssertEq ([string](_CtPersonaDeFila $fEmail $actP).email) 'pere@x.test' 'Es el tecnic (correu del titular): el de l''Excel'
$fRep = @($filesP | Where-Object { $_.Qui -eq 'representant_legal' })[0]
AssertEq ([string](_CtPersonaDeFila $fRep $actP).nom) 'Pere Tecnic Inventat' 'Es el tecnic (representant legal sense documents): el que diu l''Excel'
$cT = _CtCorreccioPersona 'es_tecnic' (_CtPersonaDeFila $fRep $actP)
AssertEq "$($cT.tipus)|$(@($cT.claus) -join ',')" 'es_tecnic|m:pere tecnic inventat' 'la correccio porta les claus de la persona'
AssertEq (_CtCorreccioEdita @{ Qui = 'autoritzat'; Dada = '' } 'x') $null 'Edita: una persona autoritzada no s''edita (es marca)'
AssertEq ([string](_CtCorreccioEdita $fEmail ' a@b.test ').camp) 'titular.email' 'Edita: el camp es qui.dada'
AssertEq ([string](_CtCorreccioDescarta $actP.avisos[0]).avis) 'es_el_tecnic|e-mail|pere@x.test' 'Descarta: per l''identificador de l''avis'
$llP = @(_CtFilesLlista @{ '1' = $actP; '2' = [ordered]@{ avisos = @(); editat_a_ma = $false; excel = @{ TITULAR = 'ALTRE SL' } } } 'avisos' '')
AssertEq "$($llP.Count)|$($llP[0].Gia)|$($llP[0].Color)" '1|1|tecnic' 'llista: el filtre "amb avisos" i el color del tecnic'
AssertEq @(_CtFilesLlista @{ '1' = $actP; '2' = [ordered]@{ avisos = @(); editat_a_ma = $false; excel = @{ TITULAR = 'ALTRE SL' } } } 'totes' 'altre').Count 1 'llista: la cerca pel titular'

Write-Host "`n--- Contactes: guards de font ---"
$srcCt = (@(Get-ChildItem -LiteralPath (Split-Path -Parent $TestsDir) -Filter 'Contactes*.ps1' -File | Sort-Object Name) | ForEach-Object { _SenseComentaris $_.FullName }) -join "`n"
# L'Excel d'activitats NO ES MODIFICA MAI (l'usuari): ni s'obre per escriure ni
# s'hi desa res.
Assert (-not ($srcCt -match 'Read-FullaEstesa|Excel\.Application|SaveAs|\.Save\(')) 'Contactes: no obre l''Excel d''activitats (ni per llegir: la fila arriba de la cache) i no desa res a l''Excel'
Assert ($srcCt.Contains('Get-FitxersPrimerNivell $dir')) 'Contactes: nomes el primer nivell de cada carpeta (Get-FitxersPrimerNivell)'

Write-Host "`n--- Activitats.ps1: les columnes de contacte de l'Excel ---"
# Amb un doble de Read-FullaEstesa (sense Excel): una fulla amb les capcaleres
# de debo de la base de dades a les seves columnes.
$origRFE = ${function:Read-FullaEstesa}
try {
    $nColsX = 100
    $matX = New-Object 'object[,]' 3, ($nColsX + 1)
    $capX = @{ 1 = 'ID Activitat'; 2 = ('N' + [char]0x00FA + 'm. expedient '); 10 = ('Ra' + [char]0x00F3 + ' social'); 11 = ('Ra' + [char]0x00F3 + ' soc. NIF')
               12 = ('Ra' + [char]0x00F3 + ' soc. Tel' + [char]0x00E8 + 'fon'); 23 = ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil'); 25 = ('Ra' + [char]0x00F3 + ' soc. E-mail')
               30 = 'Representant legal'; 31 = 'Rep. Leg. NIF'; 32 = ('Rep. Leg. Tel' + [char]0x00E8 + 'fon'); 33 = ('Rep. Leg. M' + [char]0x00F2 + 'bil'); 34 = 'Rep. Leg. E-mail' }
    $valX = @{ 1 = 9001.0; 2 = '2025/1/1'; 10 = 'EXEMPLE U SL'; 11 = 'B11111118'; 12 = '930000011'; 23 = '600000011'; 25 = 'u@exemple.test'
               30 = 'Maria Admin Prova'; 31 = '33333333P'; 32 = '930000012'; 33 = '600000012'; 34 = 'maria@exemple.test' }
    foreach ($k in $capX.Keys) { $matX[1, $k] = $capX[$k]; $matX[2, $k] = $valX[$k] }
    function Read-FullaEstesa($f, [scriptblock]$cos, [switch]$Desa, [string]$Fulla = '') {
        $d = $matX
        $cel = { param($r, $c) $v = $d[$r, $c]; if ($null -eq $v) { '' } else { ([string]$v).Trim() } }.GetNewClosure()
        return (& $cos @{ Data = $d; Rows = 2; Cols = $nColsX; Headers = @(); Cel = $cel })
    }
    $cacheX = Initialize-ActivitatsCache ([pscustomobject]@{ FullName = 'x.xls' })
    $ctX = $cacheX.Contactes['9001']
    AssertEq "$($ctX.NIF)|$($ctX.TELEFON)|$($ctX.MOBIL)|$($ctX.EMAIL)" 'B11111118|930000011|600000011|u@exemple.test' 'Activitats: les columnes del titular (NIF, telefon, mobil, e-mail) per la capcalera'
    AssertEq "$($ctX.REP_NOM)|$($ctX.REP_NIF)|$($ctX.REP_TELEFON)|$($ctX.REP_MOBIL)|$($ctX.REP_EMAIL)" 'Maria Admin Prova|33333333P|930000012|600000012|maria@exemple.test' 'Activitats: les del representant legal'
    AssertEq ((@($cacheX.ById['9001'].Keys) | Where-Object { @('NIF', 'TELEFON', 'REP_NOM', 'REP_NIF', 'REP_TELEFON', 'REP_MOBIL') -contains $_ }).Count) 0 'Activitats: les dades de contacte NO van a ById (que es puja al Drive del mobil)'
} catch {
    Assert $false ('bloc columnes de contacte: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    Set-Item -Path function:Read-FullaEstesa -Value $origRFE
}

Write-Host "`n--- ValidarContactes: la comparacio amb la referencia (pura) ---"
AssertEq (_CtCampCanonic 'Rep. Leg. E-mail') 'rep_email' '_CtCampCanonic: Rep. Leg. E-mail'
AssertEq (_CtCampCanonic 'Representant legal') 'rep_nom' '_CtCampCanonic: Representant legal'
AssertEq (_CtCampCanonic ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil')) 'mobil' '_CtCampCanonic: Rao soc. Mobil'
AssertEq (_CtCampCanonic 'E-mail') 'email' '_CtCampCanonic: E-mail'
AssertEq (_CtCampCanonic 'Rep. Leg. NIF') 'rep_nif' '_CtCampCanonic: Rep. Leg. NIF'
$refV = [pscustomobject]@{
    titular = [pscustomobject]@{ nom = 'EXEMPLE SL'; email = 'Info@Exemple.test'; telefon = '600000001'; mobil = ''; confianca = 'segur' }
    representant_legal = [pscustomobject]@{ nom = 'Maria Exemple'; confianca = 'probable' }
    persones_autoritzades = @([pscustomobject]@{ nom = 'Pere'; email = 'pere@x.test'; confianca = 'segur' })
    verificacio_excel = @([pscustomobject]@{ camp = 'Representant legal'; valor_excel = 'Pere'; problema = ([char]0x00C9 + 's el t' + [char]0x00E8 + 'cnic'); proposta = '' })
}
$araV = [ordered]@{
    titular = [ordered]@{ nom = 'EXEMPLE SL'; email = 'info@exemple.test'; telefon = ''; mobil = '600000001' }
    representant_legal = [ordered]@{ nom = 'MARIA EXEMPLE' }
    persones_autoritzades = @([ordered]@{ email = 'pere@x.test' })
    avisos = @([ordered]@{ tipus = 'es_el_tecnic'; camp = 'Representant legal' })
}
AssertEq @(_CtComparaActivitat $refV $araV).Count 0 'comparacio: el mateix (correu en majuscules, telefon a un altre camp, nom sense majuscules) -> cap discrepancia'
$araV.representant_legal = $null
$araV.avisos = @()
$difV = @(_CtComparaActivitat $refV $araV)
AssertEq (@($difV | ForEach-Object { $_.Que + ':' + $_.Confianca }) -join '|') 'representant legal:probable|avisos es_el_tecnic (camps de l''Excel):segur' 'comparacio: el representant (probable a la referencia, a part) i l''avis es_el_tecnic que falta'

Write-Host "`n--- Migracio: els .bat de local\ ---"
$batsL = _BatsLocal
AssertEq ((@($batsL.Keys) | Sort-Object) -join ',') 'DiagnosticPdf.bat,ValidarContactes.bat' 'local\: el .bat de la validacio i el del diagnostic'
foreach ($k in $batsL.Keys) {
    $b = [string]$batsL[$k]
    Assert (-not ($b -match '[^\x00-\x7F]')) "$k`: ASCII pur"
    Assert (-not ($b -match '"[^"]*\^[^"]*"')) "$k`: cap '^' entre cometes"
    $scr = [regex]::Match($b, 'suport\\(\w+\.ps1)').Groups[1].Value
    Assert (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) $scr)) "$k`: el script que crida existeix ($scr)"
}
$tmpBats = Join-Path ([System.IO.Path]::GetTempPath()) ('bats-' + [guid]::NewGuid().ToString('N'))
try {
    _PosaBatsLocal $tmpBats
    AssertEq ([System.IO.File]::ReadAllText((Join-Path (Join-Path $tmpBats 'local') 'ValidarContactes.bat'))) ([string]$batsL['ValidarContactes.bat']) '_PosaBatsLocal: els escriu a local\'
} finally { Remove-Item -LiteralPath $tmpBats -Recurse -Force -ErrorAction SilentlyContinue }
