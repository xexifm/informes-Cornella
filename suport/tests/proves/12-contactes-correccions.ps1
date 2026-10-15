# Les correccions del repas de contactes (octubre 2026, la versio 2026-10-15.1):
# l'usuari va comparar la base amb la seva referencia feta a ma i va trobar
# errors de programa que s'encadenaven (subcarpetes, el GIA del document, les
# etiquetes en castella, la signatura (R:CIF), les ECA, la queixa per carpeta,
# el titular, els tecnics pel DNI). Les proves A-G son les del seu encarrec.
#
# TOTES LES DADES SON INVENTADES (domini .test, NIF i telefons de mostra).
# Es DOT-SOURCE des de run-tests.ps1 (mateix ambit i comptador d'asserts).

Write-Host "`n--- Contactes: les etiquetes no son noms (ca/es) ---"
AssertEq (_CtEsLiniaEtiqueta ('Nom i cognoms o ra' + [char]0x00F3 + ' social Nombre y apellidos o raz' + [char]0x00F3 + 'n social')) $true '_CtEsLiniaEtiqueta: l''etiqueta en catala i en castella'
AssertEq (_CtEsLiniaEtiqueta 'DNI / NIF / NIE / Pas.') $true '_CtEsLiniaEtiqueta: DNI / NIF / NIE / Pas.'
AssertEq (_CtEsLiniaEtiqueta ('Adre' + [char]0x00E7 + 'a postal (carrer, pla' + [char]0x00E7 + 'a, av., n' + [char]0x00FA + 'm., esc., pis, porta)')) $true '_CtEsLiniaEtiqueta: l''adreca amb el parentesi'
AssertEq (_CtEsLiniaEtiqueta 'CARRER INVENTAT Numero 1') $false '_CtEsLiniaEtiqueta: un valor amb paraules d''etiqueta no ho es'
AssertEq (_CtEsLiniaEtiqueta '08940') $false '_CtEsLiniaEtiqueta: un CP no es cap etiqueta'
AssertEq (_CtEsLiniaEtiqueta 'DE LA FONT') $false '_CtEsLiniaEtiqueta: sense cap paraula forta, no'
AssertEq (_CtNomEsEtiqueta ('Nombre y apellidos o raz' + [char]0x00F3 + 'n social DNI / NIF / NIE / Pas')) $true '_CtNomEsEtiqueta: el nom que sortia (132 titulars)'
AssertEq (_CtNomEsEtiqueta ') Representado por (nombre o razon social) DNI / NIF / NIE / Pas') $true '_CtNomEsEtiqueta: "Representado por..."'
AssertEq (_CtNomEsEtiqueta 'REPRESENTACIONS EXEMPLE SL') $false '_CtNomEsEtiqueta: una empresa de "representacions" si que es un nom'
AssertEq ([string](_CtPersona ('Nom i cognoms o ra' + [char]0x00F3 + ' social') 'B00000000').nom) '' '_CtPersona: una etiqueta no es mai el nom'
AssertEq (_CtSenseEtiquetes 'Nom i cognoms o rao social Nombre y apellidos DNI / NIF EXEMPLE SL B00000000') 'EXEMPLE SL B00000000' '_CtSenseEtiquetes: el valor de darrere de les etiquetes'
AssertEq (_CtSenseEtiquetes 'DE LA FONT GARCIA, JOAN') 'DE LA FONT GARCIA, JOAN' '_CtSenseEtiquetes: un nom que comenca amb "de la" no es toca'

Write-Host "`n--- Contactes: l'expedient no es mai una data ---"
AssertEq (_CtExpedientDeText 'Registre 25/03/2025, exp. 2024/3165/33') '2024/3165/33' '_CtExpedientDeText: la data no, l''expedient si'
AssertEq (_CtExpedientDeText 'Data: 23/08/2024') '' '_CtExpedientDeText: nomes una data -> res'
AssertEq (_CtExpedientDeText 'el 2025/03/14') '' '_CtExpedientDeText: una data AAAA/MM/DD tampoc'
AssertEq (_CtExpedientDeText 'Exp. 2579/2025/12') '2579/2025/12' '_CtExpedientDeText: la serie de queixes (antiga) segueix'

Write-Host "`n--- A. La instancia de la seu: etiquetes en dues llengues, valors a sota ---"
$instA = @"
INST$([char]0x00C0)NCIA GEN$([char]0x00C8)RICA
INSTANCIA GEN$([char]0x00C9)RICA
DADES DE LA PERSONA O ENTITAT INTERESSADA DATOS DE LA PERSONA O ENTIDAD INTERESADA
Nom i cognoms o ra$([char]0x00F3) social Nombre y apellidos o raz$([char]0x00F3)n social
DNI / NIF / NIE / Pas.
EXEMPLE DISTRIBUCIONS SL
B00000000
Adre$([char]0x00E7)a postal (carrer, pla$([char]0x00E7)a, av., n$([char]0x00FA)m., esc., pis, porta) Direcci$([char]0x00F3)n postal (calle, plaza, avda., n$([char]0x00FA)m., esc., piso, puerta)
CARRER INVENTAT N$([char]0x00FA)mero 1
Municipi Municipio
CP
Prov$([char]0x00ED)ncia Provincia
 CORNELLA DE LLOBREGAT
08940
BARCELONA
Tel$([char]0x00E8)fon fix Tel$([char]0x00E9)fono fijo
Tel$([char]0x00E8)fon m$([char]0x00F2)bil Tel$([char]0x00E9)fono m$([char]0x00F3)vil
Correu electr$([char]0x00F2)nic Correo electr$([char]0x00F3)nico
600000001
info@exemple-distribucions.test
DADES DE LA PERSONA REPRESENTANT DATOS DE LA PERSONA REPRESENTANTE
Representat per (nom o ra$([char]0x00F3) social) Representado por  (nombre o raz$([char]0x00F3)n social)
DNI / NIF / NIE / Pas.
PERE INVENTAT T$([char]0x00C8)CNIC
00000000T
Correu electr$([char]0x00F2)nic Correo electr$([char]0x00F3)nico
699999999
pere@enginyeria-inventada.test
DADES A EFECTES DE NOTIFICACI$([char]0x00D3)
Data de registre: 25/03/2025
"@
$inA = Read-InstanciaText $instA
AssertEq "$($inA.interessat.nom)|$($inA.interessat.nif)|$($inA.interessat.mobil)|$($inA.interessat.email)" 'EXEMPLE DISTRIBUCIONS SL|B00000000|600000001|info@exemple-distribucions.test' 'A: l''interessat (el nom es la primera linia que no es etiqueta, el NIF la seguent)'
AssertEq "$($inA.representant.nom)|$($inA.representant.nif)|$($inA.representant.mobil)|$($inA.representant.email)" ('PERE INVENTAT T' + [char]0x00C8 + 'CNIC|00000000T|699999999|pere@enginyeria-inventada.test') 'A: el representant'
AssertEq ([string]$inA.expedient) '' 'A: la data de registre no es l''expedient'
$instA2 = @"
Inst$([char]0x00E0)ncia Gen$([char]0x00E8)rica Telem$([char]0x00E0)tica
DADES DE LA PERSONA O ENTITAT INTERESSADA DATOS DE LA PERSONA O ENTIDAD INTERESADA
Nom i cognoms o ra$([char]0x00F3) social Nombre y apellidos o raz$([char]0x00F3)n social DNI / NIF / NIE / Pas. EXEMPLE ANTIC SL B00000000
Tel$([char]0x00E8)fon fix Tel$([char]0x00E9)fono fijo Tel$([char]0x00E8)fon m$([char]0x00F2)bil Tel$([char]0x00E9)fono m$([char]0x00F3)vil Correu electr$([char]0x00F2)nic Correo electr$([char]0x00F3)nico 930000000 info@exemple-antic.test
DADES A EFECTES DE NOTIFICACI$([char]0x00D3)
"@
$inA2 = Read-InstanciaText $instA2
AssertEq "$($inA2.interessat.nom)|$($inA2.interessat.nif)|$($inA2.interessat.telefon)|$($inA2.interessat.email)" 'EXEMPLE ANTIC SL|B00000000|930000000|info@exemple-antic.test' 'A (variant antiga): les etiquetes en una linia i els valors al final'
$ctxA = _CtContext @{} @{} @{ Emails = @{}; Noms = @() } @()
$docA = @{ ruta = 'GIA 9201\Instancia_generica.pdf'; modificat = '2025-03-25T00:00:00Z'; extret = $inA }
$aA = Get-ContactesActivitat '9201' @($docA) @{ TITULAR = 'EXEMPLE DISTRIBUCIONS SL'; NIF = 'B00000000' } $ctxA
AssertEq "$($aA.titular.email)|$(@($aA.persones_autoritzades)[0].nom)|$($aA.representant_legal)" ('info@exemple-distribucions.test|PERE INVENTAT T' + [char]0x00C8 + 'CNIC|') 'A: el representant amb el correu de l''enginyeria es persona autoritzada'

Write-Host "`n--- B. Autoritzacio amb 'designa' i la signatura (R:CIF) ---"
$autB = "AUTORITZACI$([char]0x00D3) DE REPRESENTACI$([char]0x00D3)`nMaria Exemple Prova, amb DNI 00000000T, en nom i representaci$([char]0x00F3) de l'empresa EXEMPLE DISTRIBUCIONS SL, amb NIF B00000000, designa com a t$([char]0x00E8)cnica responsable del projecte a Laia Tecnica Inventada, amb DNI 11111111H, perqu$([char]0x00E8) presenti la documentaci$([char]0x00F3).`nCornell$([char]0x00E0) de Llobregat, 2 de febrer de 2025`nFirmado digitalmente por`n00000000T`nMARIA`nEXEMPLE PROVA`n(R:B000000`n00)"
$rB = Read-AutoritzacioText $autB
AssertEq "$($rB.autoritzacio.signant.nom)|$($rB.autoritzacio.signant.nif)|$($rB.autoritzacio.empresa.nif)|$(@($rB.autoritzacio.autoritzats)[0].nom)" 'Maria Exemple Prova|00000000T|B00000000|Laia Tecnica Inventada' 'B: "X, amb DNI ..., en nom de Y, designa com a tecnica ... a Z"'
AssertEq (@($rB.signatures | ForEach-Object { $_.nom + '/' + $_.nif + '/' + $_.cif }) -join ',') 'MARIA EXEMPLE PROVA/00000000T/B00000000' 'B: la signatura partida en linies (i el CIF tallat) es llegeix'
AssertEq (@(_CtSignaturesRep 'Firmado digitalmente por 00000000T MARIA EXEMPLE (R:B00000000) Fecha: 2025.02.02').Count) 1 '_CtSignaturesRep: en una linia'
AssertEq (@(_CtSignaturesRep 'Firmado digitalmente por 00000000T MARIA EXEMPLE').Count) 0 '_CtSignaturesRep: sense (R:CIF) no es de representant'
$docB = @{ ruta = 'GIA 9202\Autoritzacio.pdf'; modificat = '2025-02-02T00:00:00Z'; extret = $rB }
$aB = Get-ContactesActivitat '9202' @($docB) @{ TITULAR = 'EXEMPLE DISTRIBUCIONS SL'; NIF = 'B00000000' } $ctxA
AssertEq "$($aB.representant_legal.nom)|$($aB.representant_legal.confianca)|$($aB.representant_legal.via)|$((@($aB.persones_autoritzades) | ForEach-Object { $_.nom }) -join ',')" 'MARIA EXEMPLE PROVA|segur|signatura|Laia Tecnica Inventada' 'B: el representant legal es qui signa amb (R:CIF del titular), segur; l''autoritzada, la designada'
$manB = _CtManDeCorreccions @([ordered]@{ tipus = 'es_tecnic'; claus = @(_CtClausPersona (_CtPersona 'MARIA EXEMPLE PROVA' '00000000T')); persona = (_CtPersona 'MARIA EXEMPLE PROVA' '00000000T') })
$aBm = Get-ContactesActivitat '9202' @($docB) @{ TITULAR = 'EXEMPLE DISTRIBUCIONS SL'; NIF = 'B00000000' } $ctxA $manB
AssertEq "$($aBm.representant_legal)|$(@($aBm.persones_autoritzades).Count)" '|2' 'B: marcada a ma com a tecnic, la signatura ja no la fa representant (passa a autoritzada)'
$docB2 = @{ ruta = 'GIA 9203\Instancia_generica.pdf'; modificat = '2025-02-02T00:00:00Z'; extret = (Read-InstanciaText "INSTANCIA`nFirmado digitalmente por 22222222J JOAN ENGINYER PROVA (R:B33333333)`nFirmado digitalmente por 33333333P ANNA ALCALDIA PROVA (R:P0800000B)") }
$aB2 = Get-ContactesActivitat '9203' @($docB2) @{ TITULAR = 'EXEMPLE DISTRIBUCIONS SL'; NIF = 'B00000000' } $ctxA
AssertEq "$((@($aB2.persones_autoritzades) | ForEach-Object { $_.nom }) -join ',')|$($aB2.representant_legal)" 'JOAN ENGINYER PROVA|' 'B: (R:CIF) d''una altra entitat -> autoritzada; el d''una administracio (P...) no es de ningu'

Write-Host "`n--- Contactes: el nom del que signa s'atura on toca ---"
AssertEq (_CtNetejaNom 'N.I. 12345678Z en representacio de l''empresa X') '' '_CtNetejaNom: mai un DNI dins d''un nom'
AssertEq (_CtNetejaNom 'Joan Exemple amb DNI 12345678Z') 'Joan Exemple' '_CtNetejaNom: el nom s''acaba al DNI'
AssertEq (_CtTallaNom 'Joan Exemple actuando como representante legal de la empresa X') 'Joan Exemple' '_CtTallaNom: "actuando como..."'
$au5 = Read-AutoritzacioText "D. Joan Exemple Segon amb N.I. 11111111H en representaci$([char]0x00F3) de l'empresa PROVES INVENTADES SA, amb CIF A87654321, autoritza a Pere Tecnic Inventat, amb DNI 12345678Z, per tramitar."
AssertEq "$($au5.autoritzacio.signant.nom)|$($au5.autoritzacio.signant.nif)|$($au5.autoritzacio.empresa.nom)" 'Joan Exemple Segon|11111111H|PROVES INVENTADES SA' 'autoritzacio: el nom del que signa s''atura a "amb N.I."'
$au6 = Read-AutoritzacioText "D. Joan Exemple Segon actuando como representante legal de la empresa PROVES INVENTADES SA, con CIF A87654321, otorga su representaci$([char]0x00F3)n a D. Pere Tecnic Inventat, con DNI 12345678Z."
AssertEq "$($au6.autoritzacio.signant.nom)|$(@($au6.autoritzacio.autoritzats)[0].nom)" 'Joan Exemple Segon|Pere Tecnic Inventat' 'autoritzacio: "actuando como" i "otorga su representacion a"'
foreach ($vb in @('nomena a', 'apodera a', 'faculta a', ('atorga la seva representaci' + [char]0x00F3 + ' a'))) {
    $au7 = Read-AutoritzacioText ("Jo, Anna Fisica Prova, amb DNI 22222222J, " + $vb + " Pere Tecnic Inventat, amb DNI 12345678Z.")
    AssertEq ([string]@($au7.autoritzacio.autoritzats)[0].nom) 'Pere Tecnic Inventat' ('autoritzacio: el verb "' + $vb + '"')
}

Write-Host "`n--- C. El mateix DNI de tecnic per a titulars diferents ---"
$jordiDni = '44444444A'
$docsC = @{
    '9301' = @(@{ ruta = 'GIA 9301\etram-tramit.xml'; modificat = '2025-04-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'etram'; data = '2025-04-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'CLIENT U SL' 'B11111111'); representant = (_CtPersona 'Jordi Tramitador Inventat' $jordiDni 'botiga@client-u.test' '' '644444444'); establiment = $null; autoritzacio = $null; signatures = @() }) })
    '9302' = @(@{ ruta = 'GIA 9302\etram-tramit.xml'; modificat = '2025-05-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'etram'; data = '2025-05-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'CLIENT DOS SL' 'B22222222' 'info@client-dos.test'); representant = (_CtPersona 'JORDI TRAMITADOR INVENTAT' $jordiDni 'jordi.t@correu.test'); establiment = $null; autoritzacio = $null; signatures = @() }) })
    '9303' = @(@{ ruta = 'GIA 9303\etram-tramit.xml'; modificat = '2025-06-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'etram'; data = '2025-06-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'CLIENT TRES SL' 'B33333333' 'info@client-tres.test'); representant = (_CtPersona 'Jordi Tramitador Inventat' $jordiDni 'jordi.t@correu.test'); establiment = $null; autoritzacio = $null; signatures = @() }) })
}
$xlC = @{ '9301' = @{ TITULAR = 'CLIENT U SL'; NIF = 'B11111111' }; '9302' = @{ TITULAR = 'CLIENT DOS SL'; NIF = 'B22222222' }; '9303' = @{ TITULAR = 'CLIENT TRES SL'; NIF = 'B33333333' } }
$ctxC = _CtContext $docsC $xlC @{ Emails = @{}; Noms = @() } @()
AssertEq $ctxC.NifTitulars[$jordiDni].Count 3 'C: el context compta el DNI de qui presenta per a cada titular'
$aC = Get-ContactesActivitat '9301' $docsC['9301'] $xlC['9301'] $ctxC
$pC = @($aC.persones_autoritzades)[0]
AssertEq "$($pC.nom)|$($pC.email)|$($pC.mobil)|$($aC.representant_legal)" 'Jordi Tramitador Inventat||644444444|' 'C: tecnic pel DNI (no representant legal), i el correu del client no es seu'
Assert ([string]$pC.motiu).Contains('DNI') 'C: el motiu diu que es pel DNI'
AssertEq "$($aC.titular.email)|$($aC.titular.confianca)" 'botiga@client-u.test|probable' 'C: el correu que havia posat el tecnic es del titular'
$aC2 = Get-ContactesActivitat '9302' $docsC['9302'] $xlC['9302'] $ctxC
AssertEq "$(@($aC2.persones_autoritzades)[0].email)|$($aC2.titular.email)" 'jordi.t@correu.test|info@client-dos.test' 'C: el seu correu (el que surt a dos titulars) si que es seu'
$ctxN = _CtContext @{ '9301' = @(@{ ruta = 'a'; modificat = ''; extret = @{ representant = (_CtPersona 'Jordi Tramitador Inventat') } }); '9302' = @(@{ ruta = 'b'; modificat = ''; extret = @{ representant = (_CtPersona 'JORDI TRAMITADOR INVENTAT') } }) } $xlC @{ Emails = @{}; Noms = @() } @()
Assert ((_CtMotiuTecnic (_CtPersona 'Jordi Tramitador Inventat') $ctxN).Contains('nom')) 'C: sense DNI, el nom normalitzat repetit tambe compta'

Write-Host "`n--- D. El GIA de la carpeta mana sobre el del document ---"
$conegutD = { param($g) $true }
AssertEq (_CtGiaDocument 'X GIA 595 - BAR\Esmena solicitud.pdf' @{ gia = '5' } @{} $conegutD @{}) '595' 'D: carpeta GIA 595, esmena "GIA 5" -> 595'
AssertEq (_CtGiaDocument 'X GIA 595 - BAR\Firmadoc\Esmena solicitud.pdf' @{ gia = '5' } @{} $conegutD @{}) '595' 'D: tambe des d''una subcarpeta (la carpeta mes propera que ho diu)'
AssertEq (_CtGiaDocument 'X GIA 36 - QUEIXA\2021-2-2563 GIA 361 - RC\Instancia.pdf' @{ gia = '' } @{} $conegutD @{}) '361' 'D: la subcarpeta d''expedient amb el seu GIA mana sobre la de sobre'
AssertEq (_CtGiaDocument 'Exp 2025-16-3165 BAR\Esmena solicitud.pdf' @{ gia = '5'; expedient = '2025/3165/16' } @{} $conegutD @{ '2025-3165-16' = '595' }) '595' 'D: la carpeta sense GIA pero amb l''expedient a l''Excel (en l''ordre del document) mana'
AssertEq (_CtGiaDocument 'Exp sense res\Esmena solicitud.pdf' @{ gia = '5' } @{} $conegutD @{}) '5' 'D: si cap carpeta no en diu cap, el del document'
AssertEq (_CtGiaDocument 'Exp sense res\Sub\Esmena solicitud.pdf' @{ gia = '' } @{ 'Exp sense res' = @{ '777' = $true } } $conegutD @{}) '777' 'D: els informes de la carpeta de sobre'

Write-Host "`n--- E. La queixa per la carpeta, i la carpeta del titular a dins ---"
$cq = 'GIA 36 BAR\2021-2-2579 GIA 36 - QUEIXA OLORS'
AssertEq (_CtEsQueixa '' ($cq + '\Instancia_generica.pdf')) $true 'E: un document a la carpeta d''una queixa (serie 2579 al nom)'
AssertEq (_CtEsQueixa '' ($cq + '\Firmadoc\Instancia_generica.pdf')) $true 'E: ...i a una subcarpeta seva que no es d''expedient'
AssertEq (_CtEsQueixa '' ($cq + '\2021-2-2563 GIA 36 - RC\Instancia_generica.pdf')) $false 'E: la carpeta d''expedient del titular a dins de la queixa NO es queixa'
AssertEq (_CtEsQueixa '' 'GIA 36 BAR\DENUNCIA VEINS\a.pdf') $true 'E: les paraules (QUEIXA, QUEXIA, DENUNCIA) tambe'
AssertEq (_CtEsQueixa '2021/2579/5' 'GIA 36 BAR\a.pdf') $true 'E: l''expedient del document AAAA/2579/N'
$docsE = @(
    @{ ruta = ($cq + '\Instancia_generica.pdf'); modificat = '2025-07-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'instancia'; data = '2025-07-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'COMUNITAT DE PROPIETARIS PROVA' 'H00000000' 'comunitat@correu.test'); representant = $null; establiment = $null; autoritzacio = $null; signatures = @() }) },
    @{ ruta = ($cq + '\2021-2-2563 GIA 36 - RC\Instancia_generica.pdf'); modificat = '2025-06-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'instancia'; data = '2025-06-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'BAR PROVA SL' 'B44444444' 'bar@prova.test' '' '655555555'); representant = $null; establiment = $null; autoritzacio = $null; signatures = @() }) })
$aE = Get-ContactesActivitat '36' $docsE @{ TITULAR = 'BAR PROVA SL'; NIF = 'B44444444' } $ctxA
AssertEq "$($aE.titular.email)|$($aE.titular.mobil)|$($aE.canvi_titular)" 'bar@prova.test|655555555|' 'E: l''interessat de la queixa s''exclou i el de la subcarpeta compta (ni canvi de titular)'

Write-Host "`n--- F. L'e-TRAM antic amb raoSocial 'rao' ---"
$xmlF = '<tramit xmlns="http://www.aoc.cat/etram"><solicitant><numeroDocument>22222222J</numeroDocument><raoSocial>rao</raoSocial><nom>Anna</nom><cognom1>Fisica</cognom1><cognom2>Prova</cognom2><telefon>0</telefon><telefonMobil>611111112</telefonMobil></solicitant><representant><raoSocial>-</raoSocial><nom>Pere</nom><cognom1>Tecnic</cognom1></representant></tramit>'
$eF = Read-EtramXml $xmlF
AssertEq "$($eF.interessat.nom)|$($eF.interessat.telefon)|$($eF.interessat.mobil)|$($eF.representant.nom)" 'Anna Fisica Prova||611111112|Pere Tecnic' 'F: "rao" i "-" no son cap nom (el nom + cognoms), i el telefon 0 no compta'

Write-Host "`n--- G. Les entitats de control no son persones autoritzades ---"
AssertEq (_CtEmailExclos 'oca@es.tuv.com') $true 'G: es.tuv.com'
AssertEq (_CtEmailExclos 'x@aucatel.com') $true 'G: aucatel'
AssertEq (_CtEmailExclos 'x@eca-proves.test') $true 'G: "eca-" al principi del domini'
AssertEq (_CtEmailExclos 'x@biblioteca-proves.test') $false 'G: ...pero no a mig d''una paraula ("biblioteca-")'
AssertEq (_CtEmailExclos 'x@eicon.test') $true 'G: eicon'
$docsG = @(
    @{ ruta = 'GIA 9401\Autoritzacio.pdf'; modificat = '2025-02-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'autoritzacio'; data = '2025-02-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'PROVES G SL' 'B55555555'); representant = $null; establiment = $null; signatures = @()
        autoritzacio = [ordered]@{ signant = (_CtPersona 'Maria Admin Prova' '33333333P'); empresa = (_CtPersona 'PROVES G SL' 'B55555555'); autoritzats = @((_CtPersona 'Inspector Oca Prova' '' 'inspector@aucatel.com')); en_nom_propi = $false } }) },
    @{ ruta = 'GIA 9401\etram-tramit.xml'; modificat = '2025-03-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'etram'; data = '2025-03-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'PROVES G SL' 'B55555555' 'info@proves-g.test'); representant = (_CtPersona 'Laia Oca Prova' '' 'laia@es.tuv.com'); establiment = $null; autoritzacio = $null; signatures = @() }) })
$aG = Get-ContactesActivitat '9401' $docsG @{ TITULAR = 'PROVES G SL'; NIF = 'B55555555' } $ctxA
AssertEq "$(@($aG.persones_autoritzades).Count)|$($aG.representant_legal.nom)" '0|Maria Admin Prova' 'G: ni la de l''autoritzacio (aucatel) ni la del tramit (es.tuv.com) son autoritzades'

Write-Host "`n--- Contactes: el titular no perd les dades de l'interessat ---"
$docs8 = @(@{ ruta = 'GIA 9501\Instancia_generica.pdf'; modificat = '2025-05-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'instancia'; data = '2025-05-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'NOU TITULAR SL' 'B66666666' 'info@nou-titular.test' '' '666666666'); representant = $null; establiment = $null; autoritzacio = $null; signatures = @() }) })
$a8 = Get-ContactesActivitat '9501' $docs8 @{ TITULAR = 'VELL TITULAR SL'; NIF = 'B77777777' } $ctxA
AssertEq "$($a8.titular.nom)|$($a8.titular.email)|$($a8.titular.mobil)|$($a8.titular.confianca)|$($a8.canvi_titular.nif)" 'VELL TITULAR SL|info@nou-titular.test|666666666|probable|B66666666' '8a: un altre NIF: el correu i el mobil al titular ("probable"), el nom de l''Excel, i l''avis de canvi de titular'
$a8n = Get-ContactesActivitat '9501' $docs8 @{ TITULAR = 'NOU TITULAR SL'; NIF = '12345678Z' } $ctxA
AssertEq "$($a8n.titular.email)|$($a8n.canvi_titular)" 'info@nou-titular.test|' '8a: el mateix nom amb un altre NIF (l''Excel te el NIF malament) no es canvi de titular'
$docs8b = @(@{ ruta = 'GIA 9502\Instancia_generica.pdf'; modificat = '2025-05-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'instancia'; data = '2025-05-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona 'Marta Gerent Prova' '33333333P' 'marta@empresa-prova.test' '' '677777777'); representant = $null; establiment = $null; autoritzacio = $null; signatures = @() }) })
$a8b = Get-ContactesActivitat '9502' $docs8b @{ TITULAR = 'EMPRESA PROVA SL'; NIF = 'B88888888' } $ctxA
AssertEq "$($a8b.representant_legal.nom)|$($a8b.representant_legal.confianca)|$($a8b.titular.email)|$($a8b.canvi_titular)" 'Marta Gerent Prova|probable|marta@empresa-prova.test|' '8b: la persona fisica que presenta per l''empresa amb el correu de l''empresa: candidata a representant, no canvi de titular'
$docs10 = @(@{ ruta = 'GIA 9503\Instancia_generica.pdf'; modificat = '2025-05-01T00:00:00Z'; extret = ([ordered]@{ tipus = 'instancia'; data = '2025-05-01'; titol = ''; gia = ''; expedient = ''; interessat = (_CtPersona ('Nom i cognoms o ra' + [char]0x00F3 + ' social') 'B99999999' 'info@nou.test'); representant = $null; establiment = $null; autoritzacio = $null; signatures = @() }) })
AssertEq ([string](Get-ContactesActivitat '9503' $docs10 @{ TITULAR = 'RAO DE L''EXCEL SL'; NIF = 'B99999999' } $ctxA).titular.nom) 'RAO DE L''EXCEL SL' '10: el nom que era una etiqueta (o buit) es la Rao social de l''Excel'
AssertEq $Script:ContactesVersio '2026-10-15.1' '10: la versio puja (es tornen a llegir tots els documents)'
