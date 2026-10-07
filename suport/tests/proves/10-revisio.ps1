# L'eina REVISAR REQUERIMENTS (RevisioDades.ps1, Revisio.ps1) i els ENLLACOS
# dels catalegs (Enllacos.ps1).
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

try {

Write-Host "`n--- Enllacos.ps1: els enllacos d'un cataleg ---"
$rvCat = [pscustomobject]@{
    intro = @([pscustomobject]@{ runs = @([pscustomobject]@{ t = 'Veure https://a.cat/intro.' }) })
    nodes = @([pscustomobject]@{ tipus = 'seccio'; titol = 'Incendis'; cos = @(); fills = @(
        [pscustomobject]@{ tipus = 'item'; titol = 'Llei 3/2010'; cos = @([pscustomobject]@{ runs = @([pscustomobject]@{ t = 'https://b.cat/x' }); url = $true });
                           ajuda = [pscustomobject]@{ norma = 'Llei 3/2010'; criteri = 'c'; enllac = 'https://portaljuridic.gencat.cat/eli/es-ct/l/2010/02/18/3' }; fills = @() },
        [pscustomobject]@{ tipus = 'item'; titol = 'Sense res'; cos = @(); fills = @() }) })
}
$rvEn = @(_EnllacosDeCataleg $rvCat)
AssertEq $rvEn.Count 3 'enllacos: el de la introduccio, el del text i el de la fitxa'
AssertEq ([string]$rvEn[0].Url) 'https://a.cat/intro' 'enllacos: sense el punt final de la frase'
AssertEq ([string]$rvEn[2].Camp) 'fitxa' 'enllacos: tambe el de la fitxa d''ajuda (abans no es comprovava)'
AssertEq ([string]$rvEn[2].Punt) 'Llei 3/2010' 'enllacos: amb el punt on es'
$rvLl = [pscustomobject]@{ nodes = @([pscustomobject]@{ tipus = 'item'; titol = 'Sanitat'; cos = @(); fills = @(
    [pscustomobject]@{ tipus = 'nodisposa'; titol = ''; cos = @([pscustomobject]@{ runs = @([pscustomobject]@{ t = 'https://c.cat/y' }); url = $true }); fills = @() }) }) }
AssertEq ([string]@(_EnllacosDeCataleg $rvLl)[0].Punt) 'Sanitat' 'enllacos: un sub-punt sense titol porta el nom del punt de sobre (abans sortia en blanc)'
Assert (-not ((Get-Content -Raw -LiteralPath (Join-Path $EstructuralsDir 'LLIC.json')).Contains('salutweb.gencat.cat'))) 'LLIC: sense l''adreca de salutweb, que ja no existeix'
$rvReq1Json = Read-JsonFile (Join-Path $EstructuralsDir 'REQ1.json')
$rvEnReq1 = @(_EnllacosDeCataleg $rvReq1Json)
Assert ((@($rvEnReq1 | Where-Object { $_.Camp -eq 'fitxa' })).Count -ge 150) 'enllacos: REQ1 porta els de les fitxes'
$rvCe = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Comprova-Enllacos.ps1'))
Assert ($rvCe.Contains("'Json.ps1'") -and $rvCe.Contains("'Enllacos.ps1'")) 'Comprova-Enllacos.ps1: carrega Json.ps1 (abans cridava Read-JsonFile sense) i Enllacos.ps1'
Assert (-not ($rvCe -match 'function Get-CatalegUrls')) 'Comprova-Enllacos.ps1: el recorregut NO es copia alli'

Write-Host "`n--- punts sense fitxa d'informacio ---"
$rvSense = @(_RevPuntsSenseFitxa $rvCat)
AssertEq $rvSense.Count 1 'fitxes: el punt sense fitxa'
AssertEq ([string]$rvSense[0].Punt) 'Sense res' 'fitxes: quin'
AssertEq ([string]$rvSense[0].Seccio) 'Incendis' 'fitxes: amb la seccio'
$rvReq1Sense = @(_RevPuntsSenseFitxa $rvReq1Json | ForEach-Object { $_.Punt })
# Els dos que va trobar la primera revisio de debo (octubre 2026) ja la tenen.
Assert (-not ($rvReq1Sense -contains ('Tatuatge, p' + [char]0x00ED + 'rcing i micropigmentaci' + [char]0x00F3))) 'fitxes: el tatuatge ja te la fitxa'
Assert (-not ($rvReq1Sense -contains 'Incendis - acte')) 'fitxes: l''acte d''incendis tambe'

Write-Host "`n--- vigencia: el BOE ---"
$rvDer = '<div><p>Norma derogada, con efectos de 10/05/2025, por el Real Decreto 164/2025, de 4 de marzo (Ref. BOE-A-2025-7036).</p></div>'
$rvV = _RevEstatBoe $rvDer
AssertEq $rvV.Estat 'derogada' 'BOE: "Norma derogada ... por" -> derogada'
AssertEq $rvV.SubstitutaId 'BOE-A-2025-7036' 'BOE: l''identificador de la que la substitueix'
Assert ($rvV.Substituta -like 'Real Decreto 164/2025*') 'BOE: i el nom'
AssertEq (_RevEstatBoe '<p>Texto consolidado. Última actualización publicada el 10/04/2025</p><p>Se deroga el art. 5 por la Ley 2/2020</p>').Estat 'vigent' 'BOE: una derogacio PARCIAL no fa derogada la norma'
AssertEq (_RevEstatBoe '<p>Pàgina no trobada</p>').Estat '?' 'BOE: una pagina que no diu res -> ? (no s''endevina)'
# Falses alarmes de debo (octubre 2026): el preambul parla d'ALTRES normes.
AssertEq (_RevEstatBoe '<p>Texto consolidado.</p><p>...se consideran rectificados de acuerdo con la versión de la norma anulada.</p>').Estat 'vigent' 'BOE: "la norma anulada" dins del text no fa derogada la norma (REBT, gas, alta tensio)'
AssertEq (_RevEstatBoe '<p>Texto consolidado.</p><p>La Directiva 95/16/CE fue derogada por la Directiva 2014/33/UE.</p>').Estat 'vigent' 'BOE: "fue derogada por" una directiva no fa derogat el RD d''ascensors'
AssertEq (_RevEstatBoe '<p>Texto consolidado.</p><p>Téngase en cuenta que esta disposición ya fue derogada por el Real Decreto-ley 8/2014.</p>').Estat 'vigent' 'BOE: la nota d''un article derogat no fa derogada la llei'
AssertEq (_RevEstatBoe '<h3>Real Decreto 1836/1999 ... radiactivas. [Disposición derogada]</h3><p>Publicado en: BOE</p>').Estat 'derogada' 'BOE: l''etiqueta [Disposicion derogada] del costat del titol -> derogada'
Assert ([string](_RevEstatBoe '<p>Res</p>').Detall).Contains('no diu l') 'BOE: si no se sap, diu per que'
AssertEq (_RevEstatBoe '').Estat '?' 'BOE: sense pagina -> ?'

Write-Host "`n--- vigencia: el Portal Juridic ---"
$rvPj = '<div>Descarrega PDF RDF TTL (Turtle) XML Desa Copia la URI ELI</div><span class="badge">VIGENT</span><h1>LLEI 3/2010, del 18 de febrer</h1><p>Article 5 (DEROGAT)</p>'
AssertEq (_RevEstatPjur $rvPj).Estat 'vigent' 'Portal Juridic: l''etiqueta VIGENT del costat del titol'
AssertEq (_RevEstatPjur ($rvPj.Replace('>VIGENT<', '>DEROGAT<'))).Estat 'derogada' 'Portal Juridic: DEROGAT'
AssertEq (_RevEstatPjur '<p>Article 5 (Derogat). Text de la llei vigent.</p>').Estat '?' 'Portal Juridic: un "derogat" dins del text no compta'
AssertEq (_RevEstatPjur '<p>DISPOSICIÓ DEROGATÒRIA</p> Copia la URI ELI VIGENT').Estat 'vigent' 'Portal Juridic: "DEROGATÒRIA" no es "DEROGAT"'

AssertEq (_RevMotiuPjur '' 'L''Edge no ha acabat en 45 segons' 2) 'L''Edge no ha pogut obrir la pàgina: L''Edge no ha acabat en 45 segons' 'Portal Juridic: el motiu, si l''Edge falla'
Assert ((_RevMotiuPjur '<p>x</p>' '' 0).Contains('el servidor no ha respost')) 'Portal Juridic: el motiu, si no respon ningu'

Write-Host "`n--- les retirades: no citades o derogades (a 'derogades') ---"
$rvNs = @(
    [pscustomobject]@{ Id = 'RD 842/2002'; Tipus = 'RD'; Num = '842/2002' },
    [pscustomobject]@{ Id = 'Ley 34/1998'; Tipus = 'Ley'; Num = '34/1998' },
    [pscustomobject]@{ Id = 'RD 1836/1999'; Tipus = 'RD'; Num = '1836/1999'; Derogada = $true },
    [pscustomobject]@{ Id = 'Guia X'; Tipus = 'Guia'; Num = ''; Guia = $true },
    [pscustomobject]@{ Id = 'Col·lecció TINSCI'; Tipus = 'TINSCI'; Num = ''; Colleccio = $true },
    [pscustomobject]@{ Id = 'Ordenança tipus'; Tipus = 'Ordenança'; Num = '' })
$rvSep = _NormativaSepara $rvNs (_NormativaNormText 'Segons el Real Decreto 842/2002, de 2 de agosto...')
AssertEq (@($rvSep.Actives | ForEach-Object { $_.Id }) -join ',') 'RD 842/2002,Guia X,Col·lecció TINSCI,Ordenança tipus' 'separa: es queden la citada, les guies, les col.leccions i la que no es pot reconeixer'
AssertEq (@($rvSep.Retirades | ForEach-Object { $_.Norma.Id + '=' + $_.Motiu }) -join ',') 'Ley 34/1998=no citada,RD 1836/1999=derogada' 'separa: fora la no citada i la derogada'
Assert (-not (_NormativaEsCitada $rvNs[0] (_NormativaNormText 'RD 1842/2002'))) 'citada: el numero sencer (1842 no es 842)'
$rvReal = Get-NormativaActives
$rvAct = @($rvReal.Actives | ForEach-Object { [string]$_.Id })
$rvRet = @($rvReal.Retirades | ForEach-Object { [string]$_.Norma.Id })
Assert (($rvAct -contains 'RD 1217/2024') -and ($rvRet -contains 'RD 1836/1999')) 'real: radioactives, el RD 1217/2024 nou es queda i el 1836/1999 va a derogades'
Assert (($rvAct -contains 'RD 919/2006') -and ($rvAct -contains 'RD 842/2002')) 'real: el gas i el REBT (falses alarmes) es queden'
Assert (($rvRet -contains 'Ley 34/1998') -and ($rvRet -contains 'Llei 13/2017')) 'real: les que no cita ningu, fora'
Assert ($rvAct -contains 'Decret 64/2014') 'real: la que nomes cita Llicencia es queda (es miren tots els catalegs)'
AssertEq (@($rvReal.Actives | Where-Object { $_.Derogada }).Count) 0 'real: cap derogada a la llista (tot el tema Antic va a derogades)'
$rvGuies = @(Get-NormativaCataleg | Where-Object { $_.Guia -and -not $_.Derogada }).Count
AssertEq (@($rvReal.Actives | Where-Object { $_.Guia }).Count) $rvGuies 'real: totes les guies (no derogades) es queden'

Write-Host "`n--- vigencia: les metadades ELI ---"
AssertEq (_RevEstatEli '<eli:in_force rdf:resource="http://data.europa.eu/eli/ontology#InForce-inForce"/>').Estat 'vigent' 'ELI: InForce-inForce -> vigent'
AssertEq (_RevEstatEli 'eli:in_force <http://data.europa.eu/eli/ontology#InForce-notInForce> .').Estat 'derogada' 'ELI: InForce-notInForce -> derogada'
AssertEq (_RevEstatEli '#InForce-partiallyInForce').Estat 'vigent' 'ELI: parcialment vigent -> vigent'
AssertEq (_RevEstatEli '<p>res</p>').Estat '?' 'ELI: sense la dada -> ?'

Write-Host "`n--- l'informe i la rajola ---"
$rvFila = _RevFila 'Enllaç trencat' 'REQ1' 'Incendis' 'No respon' 'https://x.cat' 'Canvia-la'
AssertEq @($rvFila).Count 6 'informe: sis columnes'
AssertEq ([string]$rvFila[4].Link) 'https://x.cat' 'informe: l''enllac es clicable'
$rvBytes = _NormativaXlsxBytes $Script:RevCapcalera @($rvFila, $rvFila) $Script:RevAmples 'Revisio'
Add-Type -AssemblyName System.IO.Compression
$rvZip = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $rvBytes)), [System.IO.Compression.ZipArchiveMode]::Read)
$rvSr = New-Object System.IO.StreamReader($rvZip.GetEntry('xl/workbook.xml').Open())
$rvWb = $rvSr.ReadToEnd(); $rvSr.Dispose(); $rvZip.Dispose()
Assert ($rvWb.Contains('<sheet name="Revisio"') -and $rvWb.Contains('Revisio!$A$1:$F$3')) 'informe: el full es diu Revisio i el filtre hi apunta'
AssertEq ([string]$Script:LocalSubdirs['Revisions']) 'revisions' 'informe: a local\revisions'
# El que desa la pantalla (des d'una closure): _RevInformeXlsxBytes, amb la
# capcalera. Abans s'hi passaven $Script:RevCapcalera/$Script:RevAmples des de
# la closure, on valien buit, i l'informe sortia sense capcalera.
$rvB2 = _RevInformeXlsxBytes @($rvFila)
$rvZ2 = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $rvB2)), [System.IO.Compression.ZipArchiveMode]::Read)
$rvS2 = New-Object System.IO.StreamReader($rvZ2.GetEntry('xl/worksheets/sheet1.xml').Open())
$rvX2 = $rvS2.ReadToEnd(); $rvS2.Dispose(); $rvZ2.Dispose()
Assert ($rvX2.Contains('>Qu' + [char]0x00E8 + ' cal fer<') -and $rvX2.Contains('<col ')) 'informe de la pantalla: amb capcalera i amples'
# I la memoria d'una passada es reinicia de debo (abans, des de la closure, no).
$Script:NormativaEdgeKO['servidor.penjat'] = $true
$Script:NormativaPjurCache['u'] = 'x'
Reset-NormativaCaches
AssertEq "$($Script:NormativaEdgeKO.Count)|$($Script:NormativaPjurCache.Count)" '0|0' 'Reset-NormativaCaches: cada revisio comenca de zero'
$rvMenu = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1'))
$rvWiz = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1'))
Assert ($rvMenu -match "Action = 'revisio'" -and $rvWiz -match "'revisio'\s*\{\s*Invoke-RevisioRequeriments\s*\}") 'menu: la rajola Revisar requeriments obre l''eina'
Assert ($rvMenu.Contains("'NORMATIVA'")) 'menu: les dues eines de normativa tenen la seva fila'

} catch {
    Assert $false ('Revisio: la prova ha petat: ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
}
