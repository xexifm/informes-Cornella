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
Assert ($rvReq1Sense -contains ('Tatuatge, p' + [char]0x00ED + 'rcing i micropigmentaci' + [char]0x00F3)) 'fitxes: a REQ1, el punt nou de tatuatge (que l''usuari va afegir sense fitxa)'

Write-Host "`n--- vigencia: el BOE ---"
$rvDer = '<div><p>Norma derogada, con efectos de 10/05/2025, por el Real Decreto 164/2025, de 4 de marzo (Ref. BOE-A-2025-7036).</p></div>'
$rvV = _RevEstatBoe $rvDer
AssertEq $rvV.Estat 'derogada' 'BOE: "Norma derogada ... por" -> derogada'
AssertEq $rvV.SubstitutaId 'BOE-A-2025-7036' 'BOE: l''identificador de la que la substitueix'
Assert ($rvV.Substituta -like 'Real Decreto 164/2025*') 'BOE: i el nom'
AssertEq (_RevEstatBoe '<p>Texto consolidado. Última actualización publicada el 10/04/2025</p><p>Se deroga el art. 5 por la Ley 2/2020</p>').Estat 'vigent' 'BOE: una derogacio PARCIAL no fa derogada la norma'
AssertEq (_RevEstatBoe '<p>Pàgina no trobada</p>').Estat '?' 'BOE: una pagina que no diu res -> ? (no s''endevina)'
AssertEq (_RevEstatBoe '').Estat '?' 'BOE: sense pagina -> ?'

Write-Host "`n--- vigencia: el Portal Juridic ---"
$rvPj = '<div>Descarrega PDF RDF TTL (Turtle) XML Desa Copia la URI ELI</div><span class="badge">VIGENT</span><h1>LLEI 3/2010, del 18 de febrer</h1><p>Article 5 (DEROGAT)</p>'
AssertEq (_RevEstatPjur $rvPj).Estat 'vigent' 'Portal Juridic: l''etiqueta VIGENT del costat del titol'
AssertEq (_RevEstatPjur ($rvPj.Replace('>VIGENT<', '>DEROGAT<'))).Estat 'derogada' 'Portal Juridic: DEROGAT'
AssertEq (_RevEstatPjur '<p>Article 5 (Derogat). Text de la llei vigent.</p>').Estat '?' 'Portal Juridic: un "derogat" dins del text no compta'
AssertEq (_RevEstatPjur '<p>DISPOSICIÓ DEROGATÒRIA</p> Copia la URI ELI VIGENT').Estat 'vigent' 'Portal Juridic: "DEROGATÒRIA" no es "DEROGAT"'

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
$rvMenu = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1'))
$rvWiz = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1'))
Assert ($rvMenu -match "Action = 'revisio'" -and $rvWiz -match "'revisio'\s*\{\s*Invoke-RevisioRequeriments\s*\}") 'menu: la rajola Revisar requeriments obre l''eina'
Assert ($rvMenu.Contains("'NORMATIVA'")) 'menu: les dues eines de normativa tenen la seva fila'

} catch {
    Assert $false ('Revisio: la prova ha petat: ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
}
