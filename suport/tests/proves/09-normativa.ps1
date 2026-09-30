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

} catch {
    Assert $false ('Normativa: la prova ha petat: ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
}
