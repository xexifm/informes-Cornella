# Ajuntar l'informe de llicencia amb els PDF dels organismes (PdfUnio.ps1) i
# tot el que ho envolta: la copia local dels adjunts, l'historial que diu quins
# van darrere de cada informe, i el pas de "Word a PDF" que els hi ajunta.
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.
#
# ELS PDF DE PROVA (dades\pdf\) estan fets a posta, amb certificats INVENTATS
# ("Prova OGAU", "Prova ARC") i cap dada real. Cada un cobreix un cas que un
# lector de PDF senzill fa malbe:
#   informe.pdf               el nostre: dues pagines, sense signatura
#   signat-objstm.pdf         signat en REVISIO INCREMENTAL damunt d'un flux
#                             d'objectes comprimit + taula xref en flux. Es el
#                             cas on PDFsharp 1.50/1.51 llegia la pagina VELLA i
#                             la signatura desapareixia
#   dues-signatures.pdf       dues revisions: una signatura visible (pag. 2) i
#                             una d'invisible
#   heretat-girat-enllac.pdf  Resources i MediaBox HERETATS del node Pages, pagina
#                             girada 90 graus amb la signatura i un enllac intern
#   linearitzat.pdf           linearitzat, amb fluxos d'objectes, signat
#   brossa-davant.pdf         bytes abans del "%PDF-": els desplacaments no quadren
#   xref-malmesa.pdf          el "startxref" apunta a qualsevol lloc: cal refer la taula
#   xifrat.pdf                protegit: no es pot ajuntar i s'ha de DIR

Write-Host "`n--- Ajuntar PDF: l'informe + els informes dels organismes ---"
$dirPdfProva = Join-Path $TestsDir (Join-Path 'dades' 'pdf')
$tmpUnio = Join-Path ([System.IO.Path]::GetTempPath()) ('pdfunio-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $tmpUnio -Force)
function _PU([string]$nom) { return (Join-Path $dirPdfProva $nom) }

try {
    # ---- El resum de C# a hashtable (pur) ----------------------------------
    $rs = _PdfUnioResum 'pagines=5;aplanades=2;invisibles=1'
    AssertEq $rs.Pagines 5 '_PdfUnioResum: pagines'
    AssertEq $rs.Aplanades 2 '_PdfUnioResum: aplanades'
    AssertEq $rs.Invisibles 1 '_PdfUnioResum: invisibles'

    # ---- Els originals: el que hi ha abans d'ajuntar -----------------------
    $dOg = Get-PdfDescripcio (_PU 'signat-objstm.pdf')
    AssertEq $dOg.Pagines 1 'lector: signat en revisio incremental sobre flux d''objectes, 1 pagina'
    AssertEq $dOg.Signatures 1 'lector: ...i hi TROBA la signatura (PDFsharp 1.50 no la veia)'
    AssertEq (Get-PdfDescripcio (_PU 'dues-signatures.pdf')).Signatures 2 'lector: dues revisions, dues signatures'
    AssertEq (Get-PdfDescripcio (_PU 'brossa-davant.pdf')).Signatures 1 'lector: amb brossa davant del %PDF- tambe la troba'
    AssertEq (Get-PdfDescripcio (_PU 'xref-malmesa.pdf')).Pagines 2 'lector: amb la taula malmesa, la refa'
    AssertEq (Get-PdfDescripcio (_PU 'xref-malmesa.pdf')).Signatures 2 'lector: ...i hi segueixen les dues signatures'

    # ---- La unio ------------------------------------------------------------
    $fontsU = @('signat-objstm.pdf', 'dues-signatures.pdf', 'heretat-girat-enllac.pdf', 'linearitzat.pdf', 'brossa-davant.pdf', 'xref-malmesa.pdf')
    $esperades = 2
    foreach ($f in $fontsU) { $esperades += (Get-PdfDescripcio (_PU $f)).Pagines }
    $sortU = Join-Path $tmpUnio 'unit.pdf'
    $infoU = Join-PdfAmbAdjunts (_PU 'informe.pdf') @($fontsU | ForEach-Object { _PU $_ }) $sortU
    AssertEq $infoU.Pagines $esperades 'unio: cap pagina perduda ni repetida'
    # Visibles: 1 + 1 + 1 + 1 + 1 + 1 (la visible de xref-malmesa); invisibles: 1 + 1.
    AssertEq $infoU.Aplanades 6 'unio: les signatures VISIBLES passen al contingut de la pagina'
    AssertEq $infoU.Invisibles 2 'unio: les INVISIBLES simplement es treuen'
    $dU = Get-PdfDescripcio $sortU
    AssertEq $dU.Pagines $esperades 'unio: el PDF que surt es torna a llegir sencer'
    AssertEq $dU.Signatures 0 'unio: CAP camp de signatura (seria "signatura no valida")'
    AssertEq $dU.AcroForm 0 'unio: cap formulari (AcroForm) arrossegat dels adjunts'
    $txtU = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($sortU))
    Assert ($txtU.Contains('/SigAplanada1 Do')) 'unio: l''aparenca de la signatura es DIBUIXA a la pagina (es segueix veient)'
    Assert ($txtU.StartsWith('%PDF-1.7')) 'unio: surt amb capcalera de PDF'
    AssertEq ([regex]::Matches($txtU, 'startxref').Count) 1 'unio: una sola revisio (fitxer net per signar-lo despres)'
    # El nostre informe va PRIMER: la seva pagina 1 continua sent la 1.
    $infoNomes = Join-PdfAmbAdjunts (_PU 'informe.pdf') @() (Join-Path $tmpUnio 'sol.pdf')
    AssertEq $infoNomes.Pagines 2 'unio: sense adjunts, l''informe tal qual'

    # ---- Els que NO es poden ajuntar ho han de dir ---------------------------
    $errX = ''
    try { [void](Join-PdfAmbAdjunts (_PU 'informe.pdf') @((_PU 'xifrat.pdf')) (Join-Path $tmpUnio 'x.pdf')) } catch { $errX = $_.Exception.Message }
    Assert ($errX -like '*xifrat*') ('unio: un PDF protegit peta DIENT-HO (' + $errX + ')')
    $noPdf = Join-Path $tmpUnio 'no-es-pdf.pdf'
    [System.IO.File]::WriteAllText($noPdf, 'hola')
    $errN = ''
    try { [void](Join-PdfAmbAdjunts (_PU 'informe.pdf') @($noPdf) (Join-Path $tmpUnio 'n.pdf')) } catch { $errN = $_.Exception.Message }
    Assert ($errN -like '*no es un PDF*') ('unio: un fitxer que no es PDF peta dient-ho (' + $errN + ')')

    # ---- Noms i carpeta de les copies locals (purs) ------------------------
    AssertEq (_LlicNomAdjunt 1 'OGAU') 'a.OGAU.pdf' '_LlicNomAdjunt: a.OGAU.pdf (com va demanar l''usuari)'
    AssertEq (_LlicNomAdjunt 2 ('Ag' + [char]0x00E8 + 'ncia de Residus de Catalunya')) ('b.Ag' + [char]0x00E8 + 'ncia de Residus de Catalunya.pdf') '_LlicNomAdjunt: la lletra segueix l''ordre de l''informe'
    AssertEq (_LlicNomAdjunt 3 'A/B: C?') 'c.A-B- C-.pdf' '_LlicNomAdjunt: fora els caracters que el Windows no admet'
    $vellDir = $Script:LlicDbDir
    $Script:LlicDbDir = Join-Path $tmpUnio 'base-dades-llicencies'
    Assert ((_LlicCarpetaAdjunts '924') -like '*base-dades-llicencies*GIA 924') '_LlicCarpetaAdjunts: local\base-dades-llicencies\GIA 924'

    # ---- La copia local ----------------------------------------------------
    $fontsC = @{ 'OGAU' = (_PU 'signat-objstm.pdf'); 'ARC' = (_PU 'dues-signatures.pdf'); 'SENSE' = '' }
    $cp = Copy-LlicAdjunts '924' @('OGAU', 'SENSE', 'ARC') $fontsC
    AssertEq @($cp.Errors).Count 0 'Copy-LlicAdjunts: sense errors'
    AssertEq @($cp.Llista).Count 2 'Copy-LlicAdjunts: nomes els que tenen PDF'
    AssertEq (Split-Path -Leaf @($cp.Llista)[0]) 'a.OGAU.pdf' 'Copy-LlicAdjunts: a.OGAU.pdf'
    AssertEq (Split-Path -Leaf @($cp.Llista)[1]) 'c.ARC.pdf' 'Copy-LlicAdjunts: la lletra es la de l''INFORME (la b es d''un sense PDF)'
    Assert (Test-Path -LiteralPath @($cp.Llista)[0]) 'Copy-LlicAdjunts: la copia hi es'
    AssertEq ((Get-FileHash -LiteralPath @($cp.Llista)[0]).Hash) ((Get-FileHash -LiteralPath (_PU 'signat-objstm.pdf')).Hash) 'Copy-LlicAdjunts: copia EXACTA (l''original signat queda valid)'
    AssertEq ([string]$cp.Pdfs['OGAU']) ([string]@($cp.Llista)[0]) 'Copy-LlicAdjunts: es recorda la copia, no l''original'
    # Tornar-hi amb la copia com a origen: no peta ni la duplica.
    $cp2 = Copy-LlicAdjunts '924' @('OGAU') $cp.Pdfs
    AssertEq @($cp2.Errors).Count 0 'Copy-LlicAdjunts: la copia com a origen (memoria) no es un error'
    AssertEq @(Get-ChildItem -LiteralPath (_LlicCarpetaAdjunts '924')).Count 2 'Copy-LlicAdjunts: ...ni en fa una altra'
    $cpE = Copy-LlicAdjunts '924' @('X', 'Y') @{ 'X' = (Join-Path $tmpUnio 'no-hi-es.pdf'); 'Y' = $noPdf }
    AssertEq @($cpE.Errors).Count 2 'Copy-LlicAdjunts: un que no hi es i un que no es PDF son errors'

    # ---- L'historial diu quins adjunts van darrere de cada informe --------
    $h1 = New-LlicenciaHistorial 'favorable-pre' 'I:\Informes\2026-09-24_LlicFavPre_GIA 924.docx' @($cp.Llista)
    $h0 = New-LlicenciaHistorial 'requeriment' 'I:\Informes\2026-09-01_LlicReq_GIA 924.docx'
    $h0.Data = '2026-09-01T10:00:00.0000000+02:00'
    $dbA = [pscustomobject]@{ Version = 1; Llicencies = @([ordered]@{ IdGia = '924'; Historial = @($h0, $h1) }) }
    # Amb el JSON pel mig, que es com viura: una llista d'UN adjunt no es pot
    # tornar un text pelat.
    $h2 = New-LlicenciaHistorial 'favorable-post' 'C:\x\2026-10-01_LlicFavPost_GIA 924.docx' @(@($cp.Llista)[0])
    $dbA.Llicencies[0].Historial += $h2
    $dbJ = ($dbA | ConvertTo-Json -Depth 20) | ConvertFrom-Json
    $adjPre = @(Get-LlicenciaAdjuntsDeInforme $dbJ ('/on/sigui/2026-09-24_LlicFavPre_GIA 924.docx'))
    AssertEq $adjPre.Count 2 'Get-LlicenciaAdjuntsDeInforme: troba l''informe pel NOM (encara que s''hagi mogut)'
    AssertEq (Split-Path -Leaf $adjPre[0]) 'a.OGAU.pdf' 'Get-LlicenciaAdjuntsDeInforme: en ordre'
    AssertEq @(Get-LlicenciaAdjuntsDeInforme $dbJ 'C:\x\2026-10-01_LlicFavPost_GIA 924.docx').Count 1 'Get-LlicenciaAdjuntsDeInforme: UN adjunt segueix sent una llista'
    AssertEq @(Get-LlicenciaAdjuntsDeInforme $dbJ 'C:\x\2026-09-01_LlicReq_GIA 924.docx').Count 0 'Get-LlicenciaAdjuntsDeInforme: un informe sense adjunts, cap'
    AssertEq @(Get-LlicenciaAdjuntsDeInforme $dbJ 'C:\x\un-altre.docx').Count 0 'Get-LlicenciaAdjuntsDeInforme: un informe que no es de llicencia, cap'
    AssertEq @(Get-LlicenciaAdjuntsDeInforme $null 'C:\x\un-altre.docx').Count 0 'Get-LlicenciaAdjuntsDeInforme: sense base, cap (i no peta)'

    # ---- El pas de "Word a PDF": ajuntar abans de signar -------------------
    $docPre = Join-Path $tmpUnio '2026-09-24_LlicFavPre_GIA 924.docx'
    $pdfPre = Join-Path $tmpUnio '2026-09-24_LlicFavPre_GIA 924.pdf'
    Copy-Item -LiteralPath (_PU 'informe.pdf') -Destination $pdfPre
    $aj = _PdfAdjuntaLlicencia $docPre $pdfPre $dbJ
    Assert ([bool]$aj.Fet) ('Word a PDF: ajunta els adjunts de l''informe (' + $aj.Error + ')')
    AssertEq (Get-PdfDescripcio $pdfPre).Pagines (2 + 1 + 2) 'Word a PDF: el PDF de l''informe ara porta els adjunts darrere'
    AssertEq (Get-PdfDescripcio $pdfPre).Signatures 0 'Word a PDF: ...i cap signatura que surti "no valida"'
    Assert (-not (Test-Path -LiteralPath ($pdfPre + '.ajuntant.pdf'))) 'Word a PDF: no queda cap temporal'
    $pdfAltre = Join-Path $tmpUnio 'un-altre.pdf'
    Copy-Item -LiteralPath (_PU 'informe.pdf') -Destination $pdfAltre
    $ajNo = _PdfAdjuntaLlicencia (Join-Path $tmpUnio 'un-altre.docx') $pdfAltre $dbJ
    Assert (-not $ajNo.Fet -and -not $ajNo.Error) 'Word a PDF: un informe sense adjunts, tal qual i sense error'
    # Si en falta un, NO s'ajunta res (i qui crida no el signa): un informe que
    # anuncia uns adjunts que no hi son no pot sortir com si fos complet.
    Remove-Item -LiteralPath @($cp.Llista)[1] -Force
    $pdfFalta = Join-Path $tmpUnio '2026-09-24_LlicFavPre_GIA 924-b.pdf'
    Copy-Item -LiteralPath (_PU 'informe.pdf') -Destination $pdfFalta
    $abans = (Get-FileHash -LiteralPath $pdfFalta).Hash
    $ajF = _PdfAdjuntaLlicencia $docPre $pdfFalta $dbJ
    Assert (-not $ajF.Fet -and ([string]$ajF.Error -like '*c.ARC.pdf*')) ('Word a PDF: si falta un adjunt, error que diu QUIN (' + $ajF.Error + ')')
    AssertEq (Get-FileHash -LiteralPath $pdfFalta).Hash $abans 'Word a PDF: ...i el PDF de l''informe no es toca'
    $srcSig = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'PdfSignar.ps1'))
    Assert ($srcSig -match '(?s)_PdfAdjuntaLlicencia.{0,600}NO s''ha signat.{0,300}continue') 'Word a PDF: si els adjunts fallen, aquell PDF NO es signa (continue)'
    $iAj = $srcSig.IndexOf('$aj = _PdfAdjuntaLlicencia')
    $iSig = $srcSig.IndexOf('# 2a. Signatura amb l')
    Assert ($iAj -gt 0 -and $iAj -lt $iSig) 'Word a PDF: s''ajunta ABANS de signar (la firma ha de cobrir-ho tot)'
    $Script:LlicDbDir = $vellDir
} catch {
    Assert $false ('07-pdfunio: excepcio -> ' + $_.Exception.Message + ' (linia ' + $_.InvocationInfo.ScriptLineNumber + ')')
} finally {
    if ($vellDir) { $Script:LlicDbDir = $vellDir }
    try { Remove-Item -LiteralPath $tmpUnio -Recurse -Force } catch { }
}

Write-Host "`n--- PdfUnio.ps1: el C# ha de compilar al PowerShell 5.1 (C# 5) ---"
# Al PC de l'usuari el compila el csc del .NET Framework, que es C# 5. Aqui la
# suite corre amb pwsh 7 (C# modern) i compilaria coses que alla peten. Es va
# comprovar amb 'mcs -langversion:5'; aquest guard atura el que es mes facil
# d'escriure sense adonar-se'n.
$srcUnio = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $TestsDir) 'PdfUnio.ps1'))
$csUnio = $srcUnio.Substring($srcUnio.IndexOf("@'"), $srcUnio.LastIndexOf("'@") - $srcUnio.IndexOf("@'"))
$prohibit = @()
foreach ($pat in @(@{ R = '\$"'; N = 'cadena interpolada $"..."' }, @{ R = '\?\.'; N = 'operador ?.' },
                   @{ R = '\)\s*=>'; N = 'membre amb =>' }, @{ R = 'nameof\('; N = 'nameof' },
                   @{ R = 'out var '; N = 'out var' }, @{ R = '\bis\s+\w+\s+\w+\s*\)'; N = 'patro "is T x"' })) {
    if ([regex]::IsMatch($csUnio, $pat.R)) { $prohibit += $pat.N }
}
AssertEq ($prohibit -join ', ') '' 'PdfUnio.ps1: cap construccio posterior a C# 5'
Assert ($srcUnio -notmatch '(?m)^\s*Add-Type\b(?!.*PdfUnioCs)') 'PdfUnio.ps1: nomes compila en demanar-ho (carregar el fitxer no compila res)'
