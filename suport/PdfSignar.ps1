#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Convertir informes a PDF (i signar)": passa tots els Word d'una carpeta a
  PDF i, opcionalment, els signa amb AutoFirma (certificat del magatzem de
  Windows).

.DESCRIPTION
  - Converteix cada .doc/.docx d'una carpeta (i subcarpetes) a PDF, al MATEIX
    lloc i amb el mateix nom (informe.docx -> informe.pdf), fent servir Word (COM,
    ExportAsFixedFormat). Salta els que ja tenen un PDF al dia (si no es marca
    "sobreescriure").
  - Si es demana signar, crida AutoFirma per linia de comandes (operacio 'sign',
    format PAdES, magatzem 'windows') per a cada PDF. Un filtre de certificat
    (text del titular, p. ex. el NOM o el NIF) permet que AutoFirma triï el
    certificat sol, sense diàleg. El PDF signat substitueix el sense signar.

  Per què AutoFirma amb el magatzem de Windows? Reutilitza EXACTAMENT el mateix
  certificat que ja fas servir a Windows, és l'eina oficial (la signatura és
  vàlida per a l'administració) i sap fer PAdES correctament. Només cal tenir
  AutoFirma instal·lat (habitual a l'administració).

  Les funcions PURES (rutes, decisió de conversió, arguments d'AutoFirma) son
  testejables en headless; el Word (COM) i AutoFirma només s'executen a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (testejables)
# ----------------------------------------------------------------------------

# Ruta del PDF corresponent a un document (mateixa carpeta, extensió .pdf).
function _PdfPathForDoc([string]$docPath) {
    return [System.IO.Path]::ChangeExtension($docPath, '.pdf')
}

# Decideix si cal (re)generar el PDF. Funció PURA.
#   $overwrite  -> sempre.
#   PDF no existeix -> sí.
#   El Word és més nou que el PDF -> sí (regenerar).
#   Altrament -> no (ja està al dia).
function _PdfShouldConvert([bool]$pdfExists, [datetime]$docTimeUtc, [datetime]$pdfTimeUtc, [bool]$overwrite) {
    if ($overwrite) { return $true }
    if (-not $pdfExists) { return $true }
    return ($docTimeUtc -gt $pdfTimeUtc)
}

# Valor del filtre de certificat d'AutoFirma a partir d'un text del titular.
# Buit -> '' (sense filtre). Altrament 'subject.contains:TEXT' (AutoFirma triarà
# el certificat el subjecte del qual contingui aquest text, p. ex. el NIF/nom).
function _CertFilterValue([string]$text) {
    $t = ([string]$text).Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return '' }
    return 'subject.contains:' + $t
}

# Text per defecte del CAIXETÍ de la signatura visible (reprodueix l'aspecte
# "CERTIFICAT SENSE DNI" de l'usuari: nom / càrrec / organisme / data, SENSE DNI).
# La data la posa AutoFirma amb el marcador $$SIGNDATE=...$$.
function _DefaultCaixeti {
    return (@(
        'Sergi Fadurdo Modesto'
        "Enginyer d'Activitats"
        'Aj.Cornellà de Llobregat'
        '$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$'
    ) -join "`n")
}

# Posició del caixetí a la pàgina (A4 595x842 pt): a DALT A LA DRETA. Tunejable.
$Script:AutoFirmaCaixetiPos = @{ Page = 1; LLX = 360; LLY = 740; URX = 560; URY = 815 }

# Límit REAL de Windows per a una línia d'ordres: 32767 caràcters. Deixem marge
# per a les rutes (que poden ser llargues, en xarxa) i per al filtre del
# certificat. Si un intent no hi cap, se salta (mai es prova i peta).
$Script:MaxCommandLine = 30000
# Mida màxima del base64 de la imatge del caixetí. Amb el marge de dalt, la
# imatge no pot passar d'aquí o l'ordre no hi cabria.
$Script:MaxCaixetiBase64 = 20000

# Resol els marcadors de data del caixetí NOSALTRES, en lloc de deixar-los a
# AutoFirma. Funció PURA (per això la data entra com a paràmetre).
#
# PER QUE: amb '$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$' dins de layer2Text, AutoFirma
# petava amb "Error no reconocido: begin 0, end -1, length 21" (un substring amb
# un índex no trobat) i no signava. Resolent la data aquí, AutoFirma no veu mai
# cap marcador '$$...$$' i ens estalviem tot aquest camí de codi seu. La data
# surt igual de bé: és el moment en què llancem la signatura.
function _ResolveCaixetiDate([string]$caixeti, [datetime]$ara) {
    $s = [string]$caixeti
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $rx = [regex]'\$\$SIGNDATE=([^$]*)\$\$'
    for ($n = 0; $n -lt 10; $n++) {
        $m = $rx.Match($s)
        if (-not $m.Success) { break }
        $fmt = $m.Groups[1].Value
        $val = ''
        try { $val = $ara.ToString($fmt) } catch { $val = $ara.ToString('yyyy.MM.dd HH:mm:ss') }
        $s = $s.Substring(0, $m.Index) + $val + $s.Substring($m.Index + $m.Length)
    }
    # Qualsevol altre marcador que quedi es treu: AutoFirma no els paeix.
    return [regex]::Replace($s, '\$\$[A-Z]+(=[^$]*)?\$\$', '')
}

# Dibuixa el caixetí com a IMATGE i la retorna en JPEG/base64, que és el que
# demana 'signatureRubricImage' d'AutoFirma. És l'ÚNICA manera de tenir el
# caixetí de diverses línies (el layer2Text de text només admet una línia).
# Reprodueix l'aspecte "CERTIFICAT SENSE DNI": una línia per fila, alineades a
# l'esquerra, sobre fons blanc. Només Windows (System.Drawing); si peta, retorna
# '' i el crider passa al mode de text.
function _BuildCaixetiImageBase64([string]$caixeti) {
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $p = $Script:AutoFirmaCaixetiPos
        # Mateixa proporcio que el requadre de la signatura, x2 (144 ppp: prou
        # nitid). NO es pot pujar gaire: la imatge viatja en BASE64 dins de la
        # linia d'ordres, i Windows no admet mes de 32767 caracters. Amb x4 el
        # base64 se n'anava i AutoFirma ni s'arrencava ("El nombre del archivo o
        # la extension es demasiado largo").
        $escala = 2
        $w = [int](([int]$p.URX - [int]$p.LLX) * $escala)
        $h = [int](([int]$p.URY - [int]$p.LLY) * $escala)
        if ($w -le 0 -or $h -le 0) { return '' }

        $lines = @((([string]$caixeti -replace "`r`n", "`n") -split "`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -eq 0) { return '' }

        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.Clear([System.Drawing.Color]::White)
            $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
            $alcadaLinia = [double]$h / [double]$lines.Count
            # Mida de lletra que hi cap amb un marge raonable.
            $mida = [int][Math]::Max(8, [Math]::Floor($alcadaLinia * 0.58))
            $font = New-Object System.Drawing.Font('Arial', $mida, [System.Drawing.GraphicsUnit]::Pixel)
            try {
                $y = 0.0
                foreach ($l in $lines) {
                    $g.DrawString([string]$l, $font, [System.Drawing.Brushes]::Black,
                                  [single]2, [single]($y + ($alcadaLinia - $mida) / 2.0))
                    $y += $alcadaLinia
                }
            } finally { $font.Dispose() }
        } finally { $g.Dispose() }

        $ms = New-Object System.IO.MemoryStream
        try {
            # JPEG amb qualitat moderada: el que compta es que el base64 sigui
            # curt (ha de cabre a la linia d'ordres), i el caixeti es text negre
            # sobre blanc, que comprimeix molt be.
            $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
                     Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
            if ($null -ne $codec) {
                $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
                $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
                                    [System.Drawing.Imaging.Encoder]::Quality, [long]70)
                $bmp.Save($ms, $codec, $ep)
                $ep.Dispose()
            } else {
                $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            }
            $b64 = [System.Convert]::ToBase64String($ms.ToArray())
            # Ultim tall de seguretat: si tot i aixi no hi cabria, millor no
            # intentar-ho (el crider passara al caixeti de text).
            if ($b64.Length -gt $Script:MaxCaixetiBase64) { return '' }
            return $b64
        } finally { $ms.Dispose(); $bmp.Dispose() }
    } catch {
        return ''
    }
}

# El caixetí en UNA SOLA línia, per al mode de text.
# Trosseja tant pels salts de línia REALS com pel \n LITERAL: el \n literal és el
# separador de propietats d'AutoFirma i no en pot quedar cap dins del valor.
# Per si de cas, al final es treu qualsevol barra invertida que hi quedi.
function _CaixetiUnaLinia([string]$caixeti) {
    $s = ([string]$caixeti -replace "`r`n", "`n") -replace '\\n', "`n"
    $parts = @(($s -split "`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $linia = ($parts -join (' ' + [char]0x00B7 + ' '))
    return ($linia -replace '\\', '')
}

# SEPARADOR dels extraParams: els DOS CARACTERS barra-invertida + ena, NO un salt
# de linia real.
#
# Aixo NO es una suposicio: es el codi d'AutoFirma. L'ordre 'sign' munta els
# extraParams amb CommandLineLauncher.buildProperties(), que fa
#
#     while ((endIndex = params.indexOf("\\n", beginIndex)) != -1) { ... }
#
# i a Java el literal "\\n" son els dos caracters \ + n; indexOf busca una cadena
# LITERAL (no una expressio regular). El mateix codi ho diu al comentari:
# "La division no funciona correctamente con split porque el caracter salto de
# linea se protege al insertarse por consola, asi que lo hacemos manualmente."
#
# Amb salts de linia REALS el bucle no troba mai res i AutoFirma es queda amb UNA
# sola propietat: clau 'signaturePage' i com a valor TOTA la resta de la cadena.
# Resultat: signa be (codi 0) pero SENSE signatura visible. Es exactament el que
# passava: al PDF signat hi havia /Rect[0 0 0 0] i un /AP amb BBox [0 0 0 0].
$Script:AutoFirmaConfigSep = '\n'

# Construeix la cadena d'extraParams (TEXT PLA, determinista) per a una signatura
# VISIBLE PAdES amb el caixetí donat.
# Funció PURA. Caixetí buit -> '' (signatura invisible, com abans).
# Les linies de posicio, comunes a tots els modes. Funcio PURA.
function _AutoFirmaPosLines {
    $p = $Script:AutoFirmaCaixetiPos
    return @(
        "signaturePage=$($p.Page)"
        "signaturePositionOnPageLowerLeftX=$($p.LLX)"
        "signaturePositionOnPageLowerLeftY=$($p.LLY)"
        "signaturePositionOnPageUpperRightX=$($p.URX)"
        "signaturePositionOnPageUpperRightY=$($p.URY)"
    )
}

# $mode:
#   'text'   -> caixeti de TEXT (layer2Text). ATENCIO: ha d'anar en UNA SOLA
#               LINIA, perque el \n LITERAL es el separador de PROPIETATS: si
#               n'hi posessim un dins del layer2Text, AutoFirma tallaria per
#               alli i el tros seguent (sense cap '=') faria petar
#               keyValue.substring(0, keyValue.indexOf('=')) amb indexOf = -1.
#               Es exactament l'error del registre: "Error no reconocido:
#               begin 0, end -1, length 21", i 21 son els caracters de
#               "Enginyer d'Activitats", la 2a linia del caixeti.
#   'imatge' -> caixeti dibuixat com a IMATGE (signatureRubricImage, JPEG en
#               base64). Es l'unica manera de tenir-lo de VARIES LINIES.
#               El base64 (A-Z a-z 0-9 + / =) no pot contenir mai cap '\',
#               o sigui que no trenca el separador.
function _AutoFirmaVisibleExtraParams([string]$caixeti, $ara = $null, [string]$mode = 'text') {
    if ([string]::IsNullOrWhiteSpace($caixeti)) { return '' }
    if ($null -eq $ara) { $ara = Get-Date }
    $resolt = _ResolveCaixetiDate $caixeti ([datetime]$ara)
    $lines = @(_AutoFirmaPosLines)
    if ($mode -eq 'imatge') {
        $b64 = _BuildCaixetiImageBase64 $resolt
        if ([string]::IsNullOrWhiteSpace($b64)) { return '' }
        $lines += "signatureRubricImage=$b64"
        return ($lines -join $Script:AutoFirmaConfigSep)
    }
    # Text: SEMPRE en una sola linia (vegeu el comentari de dalt).
    $unaLinia = _CaixetiUnaLinia $resolt
    $lines += 'layer2FontFamily=1'
    $lines += 'layer2FontSize=8'
    $lines += "layer2Text=$unaLinia"
    return ($lines -join $Script:AutoFirmaConfigSep)
}

# Construeix la LLISTA d'arguments (ARRAY PLA) per signar un PDF amb AutoFirma.
# Funció PURA. Operació 'sign', PAdES, magatzem 'windows'.
#
# PER QUE UN ARRAY: cada argument ha d'arribar SENCER a AutoFirma (les rutes
# porten espais). L'enquotat el fa _ArgvToCommandLine, no PowerShell.
# El valor de -config son els extraParams en TEXT PLA (CommandLineParameters.java
# el guarda tal qual, NO el descodifica de Base64) amb les propietats separades
# pel \n LITERAL (vegeu $Script:AutoFirmaConfigSep).
# IMPORTANT: retorna un array PLA (no ,$ArrayList): aixi @() l'enumera bé.
function _BuildAutoFirmaSignArgv([string]$inPdf, [string]$outPdf, [string]$filter, [string]$algorithm, [string]$caixeti = '', [string]$mode = 'text') {
    if ([string]::IsNullOrWhiteSpace($algorithm)) { $algorithm = 'SHA256withRSA' }
    $argv = @('sign', '-i', $inPdf, '-o', $outPdf, '-store', 'windows', '-format', 'pades', '-algorithm', $algorithm)
    if (-not [string]::IsNullOrWhiteSpace($filter)) { $argv += @('-filter', $filter) }
    $ep = _AutoFirmaVisibleExtraParams $caixeti $null $mode
    if (-not [string]::IsNullOrWhiteSpace($ep)) { $argv += @('-config', $ep) }
    return $argv
}

# Representació llegible de l'argv per al registre de diagnòstic (els salts de
# línia del -config es mostren com a \n perquè el log quedi en una sola línia).
# IMPORTANT: els salts de línia REALS es mostren com a <LF> i els \n LITERALS es
# deixen tal qual. Abans tots dos sortien com a '\n' i el registre no permetia
# distingir-los, que és justament el que calia saber per depurar el caixetí.
function _AutoFirmaArgvToText($argv) {
    $parts = @()
    foreach ($a in @($argv)) {
        $s = ([string]$a) -replace "`r`n", '<LF>' -replace "`n", '<LF>'
        # La imatge del caixeti son milers de caracters en base64: al registre
        # nomes n'hi posem la mida, si no el fitxer es fa inservible.
        $s = [regex]::Replace($s, 'signatureRubricImage=([A-Za-z0-9+/=]+)', {
            param($m) 'signatureRubricImage=<imatge base64, ' + $m.Groups[1].Value.Length + ' car.>'
        })
        if ($s -match '\s') { $parts += ('"' + $s + '"') } else { $parts += $s }
    }
    return ($parts -join ' ')
}

# Converteix l'argv en la LINIA D'ORDRES real, amb les cometes posades per
# NOSALTRES. Funció PURA.
#
# PER QUE: a PowerShell 5.1, 'Start-Process -ArgumentList @(...)' NO enquota res:
# ajunta els elements amb espais. Amb rutes com "5.- Sergi Fadurdo" AutoFirma
# rebia la ruta TALLADA al primer espai i responia "El fichero de entrada no
# existe: I:\...\5.-". Per això construïm la línia i l'enquotem aquí.
# Els salts de línia del -config es conserven DINS de les cometes (Windows només
# separa arguments per espais i tabuladors, no per salts de línia).
function _ArgvToCommandLine($argv) {
    $parts = @()
    foreach ($a in @($argv)) {
        $s = [string]$a
        # Cal enquotar si porta espais, tabuladors, salts de línia o ja va buit.
        if ($s -eq '' -or $s -match '[\s"]') {
            $s = $s -replace '"', '\"'
            # Barres invertides FINALS: dins de cometes, la barra escaparia la
            # cometa de tancament (p. ex. "C:\carpeta\" -> el Windows llegiria
            # una cometa literal). Es dupliquen, que és el que espera
            # CommandLineToArgvW.
            $m = [regex]::Match($s, '\\+$')
            if ($m.Success) { $s = $s.Substring(0, $m.Index) + ($m.Value + $m.Value) }
            $parts += ('"' + $s + '"')
        } else {
            $parts += $s
        }
    }
    return ($parts -join ' ')
}

# AutoFirma pot signar PERFECTAMENT (codi de sortida 0) i deixar la signatura
# INVISIBLE: si els extraParams no li arriben bé, es limita a posar un widget amb
# /Rect[0 0 0 0] i una aparença /Form amb /BBox[0 0 0 0]. Sense aquesta
# comprovació el programa donava el fitxer per bo i l'usuari no veia el caixetí
# enlloc, sense cap avís.
#
# Funció PURA sobre el TEXT del tros afegit pel signador (la revisió
# incremental), per poder-la provar sense Word ni AutoFirma.
function _PdfTextCaixetiInvisible([string]$txt) {
    if ([string]::IsNullOrEmpty($txt)) { return $false }
    return ([regex]::IsMatch($txt, '/BBox\s*\[\s*0\s+0\s+0\s+0\s*\]'))
}

# Mira NOMÉS el que el signador ha afegit al final ($lenOriginal en endavant):
# així un /BBox[0 0 0 0] que ja vingués del document original no ens enganya.
function _PdfCaixetiEsInvisible([string]$pathSignat, [long]$lenOriginal) {
    try {
        $bytes = [System.IO.File]::ReadAllBytes($pathSignat)
        $desde = [int]([Math]::Max(0, $lenOriginal))
        if ($desde -ge $bytes.Length) { $desde = 0 }
        # ISO-8859-1: 1 byte = 1 caràcter, no falla mai amb dades binàries.
        $txt = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes, $desde, $bytes.Length - $desde)
        return (_PdfTextCaixetiInvisible $txt)
    } catch { return $false }
}

# Rutes candidates on sol estar instal·lat AutoFirma.exe a Windows. Si les
# variables d'entorn no hi son (p. ex. proves a Linux), s'usen els camins
# literals habituals de Windows perque la funció sempre retorni candidats.
# Retorna un array pla de cadenes (NO un ArrayList amb ,$out: aixi @() sempre
# l'enumera bé i el bucle rep cadenes, no la llista sencera).
function _AutoFirmaCandidatePaths {
    $pf  = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles))       { 'C:\Program Files' }       else { $env:ProgramFiles }
    $px  = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }
    $paths = @(
        (Join-Path $pf 'AutoFirma\AutoFirma\AutoFirma.exe')
        (Join-Path $pf 'AutoFirma\AutoFirma.exe')
        (Join-Path $px 'AutoFirma\AutoFirma\AutoFirma.exe')
        (Join-Path $px 'AutoFirma\AutoFirma.exe')
    )
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $paths += (Join-Path $env:LOCALAPPDATA 'AutoFirma\AutoFirma\AutoFirma.exe')
    }
    return $paths
}

# Localitza AutoFirma.exe. Retorna SEMPRE una cadena (la ruta o ''). $preferit
# té prioritat (ruta desada per l'usuari).
function _FindAutoFirmaExe([string]$preferit) {
    if (-not [string]::IsNullOrWhiteSpace($preferit) -and (Test-Path -LiteralPath $preferit)) { return [string]$preferit }
    foreach ($c in @(_AutoFirmaCandidatePaths)) {
        $cs = [string]$c
        if (Test-Path -LiteralPath $cs) { return $cs }
    }
    return ''
}

# Extreu el "nom comú" (CN) d'un subjecte de certificat (DN). Funció PURA.
# "CN=NOM COGNOM - 12345678Z, O=..., C=ES" -> "NOM COGNOM - 12345678Z".
function _CertCommonName([string]$subject) {
    if ([string]::IsNullOrWhiteSpace($subject)) { return '' }
    $m = [regex]::Match($subject, 'CN=([^,]+)')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ([string]$subject).Trim()
}

# ----------------------------------------------------------------------------
# Estat desat (carpeta, opcions de signatura). Sidecar JSON, ignorat per git.
# ----------------------------------------------------------------------------
function _PdfSignarStatePath {
    $dir = if ($LocalActivitatsDir) { $LocalActivitatsDir } else { $env:TEMP }
    return (Join-Path $dir 'pdf-signar-state.json')
}

# Carpeta on el programa DESA els informes que genera, SENSE crear-la.
# Es la mateixa cascada que _ResolveOutputDir (Motor.ps1) pero en mode NOMES
# LECTURA: aqui nomes volem mirar-hi, i obrir el dialeg no ha de crear carpetes.
# Retorna '' si no n'hi ha cap d'accessible (p. ex. sense la unitat de xarxa).
function _CarpetaInformesGenerats {
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$OutputDir) -and (Test-Path -LiteralPath $OutputDir)) {
            return [string]$OutputDir
        }
    } catch { }
    try {
        $local = Join-Path $RepoRoot 'Informes generats'
        if (Test-Path -LiteralPath $local) { return [string]$local }
    } catch { }
    return ''
}

# L'ULTIM INFORME GENERAT: el .docx (o .doc) mes NOU de la carpeta de sortida.
# Es el valor per defecte del quadre de l'eina "Word a PDF": el cas real es
# sempre el mateix (acabes de generar un informe i el vols passar a PDF i
# signar-lo), i abans tocava anar-hi a buscar cada vegada.
#
# SENSE -Recurse: els informes es desen PLANS a l'arrel de la carpeta.
# Se salten els temporals de Word (~$...) i tot el que no sigui Word.
# $dir buit -> _CarpetaInformesGenerats (el parametre hi es per poder provar-la).
# Retorna '' si no hi ha cap informe o si la carpeta no es accessible.
function _UltimInformeGenerat([string]$dir = '') {
    try {
        $d = if ([string]::IsNullOrWhiteSpace($dir)) { _CarpetaInformesGenerats } else { $dir }
        if ([string]::IsNullOrWhiteSpace($d) -or -not (Test-Path -LiteralPath $d -PathType Container)) { return '' }
        $ultim = Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' -and ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -eq $ultim) { return '' }
        return [string]$ultim.FullName
    } catch { return '' }
}

# Registre de diagnòstic de les crides a AutoFirma (argv exacte + codi de sortida
# + sortida del programa). Serveix per saber QUE ha passat quan la signatura
# visible no surt, sense haver d'endevinar.
function _PdfSignarLogPath {
    $dir = if ($LocalActivitatsDir) { $LocalActivitatsDir } else { $env:TEMP }
    return (Join-Path $dir 'pdf-signar-log.txt')
}

# Executa AutoFirma amb la línia d'ordres que construïm nosaltres (amb cometes)
# i en recull el codi de sortida i el que hagi escrit. Retorna @{ExitCode;Output}.
# Fem servir ProcessStartInfo (i no Start-Process) perquè volem controlar
# EXACTAMENT la línia d'ordres: vegeu _ArgvToCommandLine.
function _RunAutoFirma([string]$afExe, $argv) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $afExe
    $psi.Arguments = (_ArgvToCommandLine $argv)
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $so = $p.StandardOutput.ReadToEnd()
    $se = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    $out = (([string]$so) + "`n" + ([string]$se)).Trim()
    return @{ ExitCode = $p.ExitCode; Output = $out }
}

function _PdfSignarLog([string]$text) {
    try {
        $line = ('[' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '] ' + $text)
        Add-Content -LiteralPath (_PdfSignarLogPath) -Value $line -Encoding UTF8
    } catch { }
}

function _LoadPdfSignarState {
    $p = _PdfSignarStatePath
    $def = @{ folder = ''; sign = $false; certFilter = ''; autofirma = ''; overwrite = $false; visibleSign = $true; caixeti = (_DefaultCaixeti) }
    if (-not (Test-Path -LiteralPath $p)) { return $def }
    try {
        $o = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @('folder','certFilter','autofirma','caixeti')) { if ($o.PSObject.Properties[$k]) { $def[$k] = [string]$o.$k } }
        if ($o.PSObject.Properties['sign'])        { $def['sign'] = [bool]$o.sign }
        if ($o.PSObject.Properties['overwrite'])   { $def['overwrite'] = [bool]$o.overwrite }
        if ($o.PSObject.Properties['visibleSign']) { $def['visibleSign'] = [bool]$o.visibleSign }
        if ([string]::IsNullOrWhiteSpace([string]$def['caixeti'])) { $def['caixeti'] = (_DefaultCaixeti) }
    } catch { }
    return $def
}

function _SavePdfSignarState($state) {
    try {
        $dir = Split-Path -Parent (_PdfSignarStatePath)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        ($state | ConvertTo-Json) | Set-Content -LiteralPath (_PdfSignarStatePath) -Encoding UTF8
    } catch { }
}

# ----------------------------------------------------------------------------
# Diàleg d'opcions (carpeta + signatura). Retorna @{Folder;Sign;CertText;
# Overwrite;AutoFirma} o $null si es cancel·la.
# ----------------------------------------------------------------------------
function _ShowConvertPdfOptions {
    $st = _LoadPdfSignarState
    # Per defecte, L'ULTIM INFORME GENERAT: es el que gairebe sempre es vol
    # signar. Si no n'hi ha cap (o la carpeta no es accessible), es recupera el
    # comportament d'abans: l'ultima ruta que s'hi va fer servir i, si tampoc,
    # la carpeta d'informes.
    $ultim = _UltimInformeGenerat
    $folder = if (-not [string]::IsNullOrWhiteSpace($ultim)) { [string]$ultim }
              elseif (-not [string]::IsNullOrWhiteSpace($st.folder) -and (Test-Path -LiteralPath $st.folder)) { [string]$st.folder }
              elseif ($InformesDir -and (Test-Path -LiteralPath $InformesDir)) { [string]$InformesDir }
              else { '' }
    $autofirma = _FindAutoFirmaExe $st.autofirma

    # Certificats disponibles al magatzem de Windows (per triar-lo d'una llista,
    # en lloc d'escriure text). Primer element: deixar-ho a AutoFirma.
    $certOpts = New-Object System.Collections.ArrayList
    [void]$certOpts.Add(@{ Display = ('(triar-lo a AutoFirma en signar)'); Filter = '' })
    try {
        Get-ChildItem 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue |
            Where-Object { $_.HasPrivateKey } | Sort-Object NotAfter -Descending | ForEach-Object {
                $cn = _CertCommonName ([string]$_.Subject)
                $exp = ''
                try { $exp = $_.NotAfter.ToString('dd/MM/yyyy') } catch { }
                [void]$certOpts.Add(@{ Display = ("{0}  (fins {1})" -f $cn, $exp); Filter = (_CertFilterValue $cn) })
            }
    } catch { }

    $form = _NewForm
    $form.Text = 'Convertir informes a PDF'
    $form.ClientSize = New-Object System.Drawing.Size(510, 476)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # Carpeta: mateix format que la Configuració (quadre editable + "..." +
    # indicador ✓/⚠ en viu). Helper comú _AddConfigRow (Configuracio.ps1).
    # Es pot triar una CARPETA (tots els Word de dins i subcarpetes) o UN SOL
    # DOCUMENT Word: el segon boto obre el dialeg de fitxers. El quadre ve ple
    # amb l'ultim informe generat, pero es pot canviar (o escriure-hi a ma).
    $row = _AddConfigRow $form 70 ('Carpeta o document (Word) ' + [char]0x2014 + " per defecte, l'ultim informe generat:") $folder 'Documents Word|*.docx;*.doc|Tots els fitxers|*.*'
    $tbF = $row.TextBox
    $y = [int]$row.NextY

    $cbOver = New-Object System.Windows.Forms.CheckBox
    $cbOver.Text = 'Tornar a generar els PDF que ja existeixen'
    $cbOver.Location = New-Object System.Drawing.Point(14, $y)
    $cbOver.AutoSize = $true
    $cbOver.Checked = [bool]$st.overwrite
    [void]$form.Controls.Add($cbOver)
    $y += 30

    $cbSign = New-Object System.Windows.Forms.CheckBox
    $cbSign.Text = 'Signar els PDF amb AutoFirma (certificat de Windows)'
    $cbSign.Location = New-Object System.Drawing.Point(14, $y)
    $cbSign.AutoSize = $true
    $cbSign.Checked = [bool]$st.sign
    [void]$form.Controls.Add($cbSign)
    $y += 26

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Certificat amb què signar:'
    $lblC.Location = New-Object System.Drawing.Point(34, $y)
    $lblC.AutoSize = $true
    [void]$form.Controls.Add($lblC)
    $y += 20

    $cbCert = New-Object System.Windows.Forms.ComboBox
    $cbCert.DropDownStyle = 'DropDownList'
    $cbCert.Location = New-Object System.Drawing.Point(34, $y)
    $cbCert.Size = New-Object System.Drawing.Size(454, 24)
    foreach ($o in $certOpts) { [void]$cbCert.Items.Add([string]$o.Display) }
    $selIdx = 0
    for ($i = 0; $i -lt $certOpts.Count; $i++) { if ([string]$certOpts[$i].Filter -eq [string]$st.certFilter -and [string]$st.certFilter -ne '') { $selIdx = $i } }
    if ($cbCert.Items.Count -gt 0) { $cbCert.SelectedIndex = $selIdx }
    [void]$form.Controls.Add($cbCert)
    $y += 30

    $lblAF = New-Object System.Windows.Forms.Label
    $lblAF.Location = New-Object System.Drawing.Point(34, $y)
    $lblAF.MaximumSize = New-Object System.Drawing.Size(454, 0)
    $lblAF.AutoSize = $true
    $lblAF.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
    if ([string]::IsNullOrWhiteSpace($autofirma)) {
        $lblAF.Text = ([char]0x26A0 + ' AutoFirma no trobat. Instal·la''l o desmarca la signatura.')
        $lblAF.ForeColor = [System.Drawing.Color]::Firebrick
    } else {
        $lblAF.Text = ([char]0x2713 + ' AutoFirma: ' + $autofirma)
        $lblAF.ForeColor = [System.Drawing.Color]::SeaGreen
    }
    [void]$form.Controls.Add($lblAF)
    $y += 40

    # Signatura VISIBLE (caixetí a dalt a la dreta) + text editable del caixetí.
    $cbVis = New-Object System.Windows.Forms.CheckBox
    $cbVis.Text = 'Signatura visible (caixetí a dalt a la dreta)'
    $cbVis.Location = New-Object System.Drawing.Point(34, $y)
    $cbVis.AutoSize = $true
    $cbVis.Checked = [bool]$st.visibleSign
    [void]$form.Controls.Add($cbVis)
    $y += 26

    $lblCx = New-Object System.Windows.Forms.Label
    $lblCx.Text = 'Text del caixetí (una línia per fila; $$SIGNDATE=...$$ = data):'
    $lblCx.Location = New-Object System.Drawing.Point(34, $y)
    $lblCx.AutoSize = $true
    $lblCx.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
    [void]$form.Controls.Add($lblCx)
    $y += 20

    $tbCx = New-Object System.Windows.Forms.TextBox
    $tbCx.Location = New-Object System.Drawing.Point(34, $y)
    $tbCx.Size = New-Object System.Drawing.Size(454, 76)
    $tbCx.Multiline = $true
    $tbCx.ScrollBars = 'Vertical'
    $tbCx.AcceptsReturn = $true
    $tbCx.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    # El TextBox multilínia només mostra CRLF; el caixetí es guarda amb LF.
    $tbCx.Text = ([string]$st.caixeti -replace "`r?`n", "`r`n")
    [void]$form.Controls.Add($tbCx)

    $syncSign = {
        $on = $cbSign.Checked
        $lblC.Enabled = $on; $cbCert.Enabled = $on; $lblAF.Enabled = $on
        $cbVis.Enabled = $on
        $onVis = ($on -and $cbVis.Checked)
        $lblCx.Enabled = $onVis; $tbCx.Enabled = $onVis
    }.GetNewClosure()
    $cbSign.add_CheckedChanged($syncSign)
    $cbVis.add_CheckedChanged($syncSign)
    & $syncSign

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Comença'
    $btnGo.Location = New-Object System.Drawing.Point(280, 438)
    $btnGo.Size = New-Object System.Drawing.Size(120, 30)
    $btnGo.Anchor = 'Bottom, Right'
    _StylePrimaryButton $btnGo
    [void]$form.Controls.Add($btnGo)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Tanca'
    $btnCancel.Location = New-Object System.Drawing.Point(406, 438)
    $btnCancel.Size = New-Object System.Drawing.Size(88, 30)
    $btnCancel.Anchor = 'Bottom, Right'
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $form.DialogResult = 'Cancel'; $form.Close() }.GetNewClosure())
    [void]$form.Controls.Add($btnCancel)

    $result = @{ Value = $null }
    $btnGo.add_Click({
        $f = [string]$tbF.Text
        if ([string]::IsNullOrWhiteSpace($f) -or -not (Test-Path -LiteralPath $f)) {
            [System.Windows.Forms.MessageBox]::Show('Tria una carpeta o un document vàlids.', 'Convertir informes a PDF', 'OK', 'Warning') | Out-Null
            return
        }
        if ($cbSign.Checked -and [string]::IsNullOrWhiteSpace($autofirma)) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha trobat AutoFirma. Desmarca la signatura (es faran només els PDF) o instal·la AutoFirma.", 'Convertir informes a PDF', 'OK', 'Warning') | Out-Null
            return
        }
        $certFilter = ''
        if ($cbCert.SelectedIndex -ge 0 -and $cbCert.SelectedIndex -lt $certOpts.Count) { $certFilter = [string]$certOpts[$cbCert.SelectedIndex].Filter }
        $caixeti = ([string]$tbCx.Text -replace "`r`n", "`n")
        $result.Value = @{
            Folder = $f; Sign = [bool]$cbSign.Checked; CertFilter = $certFilter
            Overwrite = [bool]$cbOver.Checked; AutoFirma = [string]$autofirma
            VisibleSign = [bool]$cbVis.Checked; Caixeti = $caixeti
        }
        _SavePdfSignarState @{ folder = $f; sign = [bool]$cbSign.Checked; certFilter = $certFilter; autofirma = [string]$autofirma; overwrite = [bool]$cbOver.Checked; visibleSign = [bool]$cbVis.Checked; caixeti = $caixeti }
        $form.DialogResult = 'OK'; $form.Close()
    }.GetNewClosure())

    [void](_AddBrandHeader $form 'Convertir informes a PDF' ('PDF a la mateixa carpeta ' + [char]0x00B7 + ' signatura opcional amb AutoFirma'))
    [void]$form.ShowDialog()
    $v = $result.Value
    $form.Dispose()
    return $v
}

# ----------------------------------------------------------------------------
# Execució: converteix (i signa) amb barra de progrés i cancel·lació.
# ----------------------------------------------------------------------------
function _RunConvertPdf($opts) {
    # Defensa: si $opts arribés embolcallat en un array, agafem l'element real.
    if ($opts -is [System.Array]) { $opts = @($opts)[-1] }
    $folderPath = [string]$opts.Folder
    $afExe      = [string]$opts.AutoFirma
    # Pot ser UN SOL DOCUMENT o una CARPETA (llavors, tots els Word de dins i de
    # les subcarpetes, saltant els temporals ~$).
    $esFitxer = $false
    try { $esFitxer = Test-Path -LiteralPath $folderPath -PathType Leaf } catch { }
    $files = New-Object System.Collections.ArrayList
    if ($esFitxer) {
        $fi = Get-Item -LiteralPath $folderPath -ErrorAction SilentlyContinue
        if ($null -ne $fi -and ($fi.Extension -ieq '.docx' -or $fi.Extension -ieq '.doc')) { [void]$files.Add($fi) }
    } else {
        Get-ChildItem -LiteralPath $folderPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.Name -notlike '~$*' -and ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc')) {
                [void]$files.Add($_)
            }
        }
    }
    if ($files.Count -eq 0) {
        $qui = if ($esFitxer) { "El fitxer triat no es un Word (.doc/.docx)." } else { "No s'ha trobat cap Word (.doc/.docx) a la carpeta." }
        [System.Windows.Forms.MessageBox]::Show($qui, 'Convertir informes a PDF', 'OK', 'Information') | Out-Null
        return
    }

    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("S'han trobat {0} documents Word.`n`nEs generaran els PDF a la mateixa carpeta{1}.`n`nVols continuar?" -f $files.Count, $(if ($opts.Sign) { ' i es signaran amb AutoFirma' } else { '' })),
        'Convertir informes a PDF', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # ---- Finestra de progrés amb cancel·lació ----
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Convertir informes a PDF'
    $form.Size = New-Object System.Drawing.Size(580, 200)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(530, 60)
    $lbl.Text = 'Preparant...'
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 92)
    $bar.Size = New-Object System.Drawing.Size(530, 22)
    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $files.Count); $bar.Value = 0
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(430, 124)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({
        param($s, $e)
        if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true }
    }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $filter = [string]$opts.CertFilter
    # Caixetí de signatura visible (buit = signatura invisible, com abans).
    $caixeti = if ([bool]$opts.VisibleSign) { [string]$opts.Caixeti } else { '' }
    $converted = 0; $skipped = 0; $signed = 0; $errors = 0; $done = 0
    # Quants s'han hagut de signar SENSE caixeti perque el visible ha fallat.
    $senseCaixeti = 0
    $errDetalls = New-Object System.Collections.ArrayList
    $word = $null
    try {
        $word = New-WordApp
        if ($null -eq $word) { return }   # New-WordApp ja avisa

        foreach ($f in $files) {
            if ($cancel.Flag) { break }
            $done++
            $bar.Value = [Math]::Min($bar.Maximum, $done)
            $lbl.Text = ("Processant {0} de {1}...`n{2}" -f $done, $files.Count, $f.Name)
            [System.Windows.Forms.Application]::DoEvents()

            $pdf = _PdfPathForDoc $f.FullName
            $pdfExists = Test-Path -LiteralPath $pdf
            $pdfTime = if ($pdfExists) { (Get-Item -LiteralPath $pdf).LastWriteTimeUtc } else { [datetime]::MinValue }

            # 1. Conversió a PDF (si cal).
            if (_PdfShouldConvert $pdfExists $f.LastWriteTimeUtc $pdfTime ([bool]$opts.Overwrite)) {
                try {
                    $doc = $word.Documents.Open($f.FullName, $false, $true)   # ReadOnly
                    try {
                        $doc.ExportAsFixedFormat($pdf, 17)   # 17 = wdExportFormatPDF
                    } finally {
                        $doc.Close($false)
                    }
                    $converted++
                    $pdfExists = $true
                } catch {
                    $errors++
                    [void]$errDetalls.Add(("PDF: {0} -> {1}" -f $f.Name, $_.Exception.Message))
                    continue
                }
            } else {
                $skipped++
            }

            # 2. Signatura (si es demana i el PDF existeix).
            if ($opts.Sign -and $pdfExists) {
                if ($cancel.Flag) { break }
                $lbl.Text = ("Signant {0} de {1}...`n{2}" -f $done, $files.Count, $f.Name)
                [System.Windows.Forms.Application]::DoEvents()
                $tmpSigned = $pdf + '.signat.pdf'
                try { if (Test-Path -LiteralPath $tmpSigned) { Remove-Item -LiteralPath $tmpSigned -Force } } catch { }
                try {
                    # Intents en ordre de preferència. Si el caixetí multilínia
                    # falla, provem el d'UNA LÍNIA (així sabem si el problema son
                    # els salts de línia) i, si tampoc, SENSE caixetí: val més un
                    # PDF signat sense caixetí que cap PDF signat.
                    $intents = New-Object System.Collections.ArrayList
                    if ($caixeti) {
                        # 1r: IMATGE (l'unica manera de tenir-lo de diverses linies).
                        # 2n: TEXT d'una linia (comprovat que AutoFirma l'accepta).
                        [void]$intents.Add(@{ Nom = 'caixeti (imatge)'; Text = $caixeti; Mode = 'imatge' })
                        [void]$intents.Add(@{ Nom = 'caixeti (text d''una linia)'; Text = $caixeti; Mode = 'text' })
                    }
                    [void]$intents.Add(@{ Nom = 'sense caixeti'; Text = ''; Mode = 'text' })

                    $res = $null; $argv = @(); $usat = ''
                    foreach ($intent in $intents) {
                        try { if (Test-Path -LiteralPath $tmpSigned) { Remove-Item -LiteralPath $tmpSigned -Force } } catch { }
                        $usat = [string]$intent.Nom
                        # IMPORTANT: el try/catch va DINS del bucle. Si una
                        # excepcio d'un intent sortia fora, s'enduia TOTS els
                        # intents i el fitxer es quedava sense signar (va passar
                        # amb la imatge: "El nombre del archivo o la extension es
                        # demasiado largo").
                        try {
                            $argv = @(_BuildAutoFirmaSignArgv $pdf $tmpSigned $filter '' ([string]$intent.Text) ([string]$intent.Mode))
                            # Windows no admet linies d'ordres de mes de 32767
                            # caracters. Si l'intent no hi cabria, ni el provem.
                            $llarg = (_ArgvToCommandLine $argv).Length
                            if ($llarg -gt $Script:MaxCommandLine) {
                                _PdfSignarLog ("AVIS: l'intent '" + $usat + "' no hi cap a la linia d'ordres (" + $llarg + " car.); el salto.")
                                $res = @{ ExitCode = -1; Output = 'ordre massa llarga' }
                                continue
                            }
                            $res = _RunAutoFirma $afExe $argv
                        } catch {
                            $res = @{ ExitCode = -1; Output = $_.Exception.Message }
                        }
                        if ($res.ExitCode -eq 0 -and (Test-Path -LiteralPath $tmpSigned)) {
                            # Codi 0 NO vol dir caixeti visible: si els
                            # extraParams no li arriben be, AutoFirma signa
                            # igualment pero amb /BBox[0 0 0 0] (invisible).
                            $volCaixeti = -not [string]::IsNullOrWhiteSpace([string]$intent.Text)
                            $lenOrig = 0
                            try { $lenOrig = (Get-Item -LiteralPath $pdf).Length } catch { }
                            if ($volCaixeti -and (_PdfCaixetiEsInvisible $tmpSigned $lenOrig)) {
                                _PdfSignarLog ("AVIS: l'intent '" + $usat + "' ha signat pero el caixeti ha quedat INVISIBLE (/BBox[0 0 0 0]) a " + $f.Name)
                                _PdfSignarLog ("   ordre: " + (_AutoFirmaArgvToText $argv))
                                continue
                            }
                            break
                        }
                        _PdfSignarLog ("AVIS: ha fallat l'intent '" + $usat + "' (codi " + $res.ExitCode + ") a " + $f.Name)
                        _PdfSignarLog ("   ordre: " + (_AutoFirmaArgvToText $argv))
                        if ($res.Output) { _PdfSignarLog ("   sortida: " + $res.Output) }
                    }
                    if ($usat -eq 'sense caixeti' -and $res.ExitCode -eq 0) { $senseCaixeti++ }
                    if ($res.ExitCode -eq 0 -and (Test-Path -LiteralPath $tmpSigned)) {
                        Move-Item -LiteralPath $tmpSigned -Destination $pdf -Force
                        $signed++
                        # Nomes la primera: serveix per comprovar l'ordre exacta.
                        if ($signed -eq 1) { _PdfSignarLog ("OK (" + $usat + ")  " + $f.Name + "  ::  " + (_AutoFirmaArgvToText $argv)) }
                    } else {
                        $errors++
                        [void]$errDetalls.Add(("Signatura: {0} (codi {1})" -f $f.Name, $res.ExitCode))
                        _PdfSignarLog ("ERROR codi " + $res.ExitCode + "  " + $f.Name + "  ::  " + (_AutoFirmaArgvToText $argv))
                        if ($res.Output) { _PdfSignarLog ("   sortida: " + $res.Output) }
                        try { if (Test-Path -LiteralPath $tmpSigned) { Remove-Item -LiteralPath $tmpSigned -Force } } catch { }
                    }
                } catch {
                    $errors++
                    [void]$errDetalls.Add(("Signatura: {0} -> {1}" -f $f.Name, $_.Exception.Message))
                    _PdfSignarLog ("EXCEPCIO " + $f.Name + " -> " + $_.Exception.Message)
                }
            }
        }
    } finally {
        if ($null -ne $word) {
            try { $word.Quit() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
        }
        $cancel.Running = $false
        $form.Close(); $form.Dispose()
    }

    # ---- Resum ----
    $msg = New-Object System.Text.StringBuilder
    if ($cancel.Flag) { [void]$msg.AppendLine('Cancel·lat.') ; [void]$msg.AppendLine('') }
    [void]$msg.AppendLine(("PDF generats: {0}" -f $converted))
    [void]$msg.AppendLine(("Ja estaven al dia (saltats): {0}" -f $skipped))
    if ($opts.Sign) {
        [void]$msg.AppendLine(("PDF signats: {0}" -f $signed))
        if ($senseCaixeti -gt 0) {
            [void]$msg.AppendLine(("  (dels quals {0} SENSE caixeti: el caixeti ha fallat i s'han signat igualment)" -f $senseCaixeti))
        }
    }
    if ($errors -gt 0) {
        [void]$msg.AppendLine(("Errors: {0}" -f $errors))
        [void]$msg.AppendLine('')
        $mostra = @($errDetalls) | Select-Object -First 8
        foreach ($e in $mostra) { [void]$msg.AppendLine('  - ' + $e) }
        if ($errDetalls.Count -gt 8) { [void]$msg.AppendLine(('  ... i {0} més.' -f ($errDetalls.Count - 8))) }
    }
    # Si s'ha signat, oferim el registre: hi ha l'ordre EXACTA que s'ha passat a
    # AutoFirma i serveix per veure per que no surt el caixeti, si es el cas.
    if ($opts.Sign -and ($signed -gt 0 -or $errors -gt 0)) {
        [void]$msg.AppendLine('')
        [void]$msg.AppendLine("Vols obrir el registre de la signatura? (hi ha l'ordre exacta")
        [void]$msg.AppendLine('enviada a AutoFirma; util si el caixeti no apareix al PDF)')
        $r = [System.Windows.Forms.MessageBox]::Show($msg.ToString(), 'Convertir informes a PDF', 'YesNo',
                                                     $(if ($errors -gt 0) { 'Warning' } else { 'Information' }))
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            try { Start-Process -FilePath (_PdfSignarLogPath) | Out-Null } catch { }
        }
        return
    }
    $icon = if ($errors -gt 0) { 'Warning' } else { 'Information' }
    [System.Windows.Forms.MessageBox]::Show($msg.ToString(), 'Convertir informes a PDF', 'OK', $icon) | Out-Null
}

# Punt d'entrada de l'eina (des del menú principal).
function Invoke-ConvertirPdf {
    $opts = _ShowConvertPdfOptions
    if ($null -eq $opts) { return }
    _RunConvertPdf $opts
}
