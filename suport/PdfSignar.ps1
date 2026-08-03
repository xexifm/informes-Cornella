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

# Posició del caixetí a la pàgina (A4 595x842 pt): a DALT A LA DRETA, ALINEAT amb
# la capçalera de l'informe. L'alçada són 48 pt per a 4 línies (abans 75): amb 75
# les línies quedaven molt separades i el caixetí, innecessàriament alt.
#
# Els dos números d'alineació NO són a ull: surten de mesurar un informe ja
# generat (es descomprimeix el flux de contingut de la pàgina 1 del PDF).
#   · Dalt (URY = 800) = la punta de l'escut de la capçalera. Compte, que NO és
#     el 806,52 on està col·locada la imatge del logo: aquella imatge porta 18
#     píxels de blanc a dalt (de 199), que a la pàgina són 6,5 pt. Alineant amb
#     el 806,52 el caixetí hauria quedat mig dit massa amunt.
#   · Dreta (URX = 552) = el marge dret del text, tret del requadre de la "Nota:"
#     de l'informe, que va de x=85,58 a x=552,45.
$Script:AutoFirmaCaixetiPos = @{ Page = 1; LLX = 352; LLY = 752; URX = 552; URY = 800 }

# Aspecte del caixetí-imatge. Tot el que es pot tocar sense entrar al codi.
$Script:CaixetiEstil = @{
    # Proporció de la mida de lletra respecte de l'alçada de línia. Com més
    # alta, més omplen les lletres i menys espai buit queda entremig (era 0,58).
    FactorLletra   = 0.72
    # Contorn del requadre: gris fosc, gruix en píxels per unitat d'escala.
    ContornRGB     = @(64, 64, 64)
    ContornGruix   = 1
    # Escut de fons, a la dreta. Opacitat baixa perquè el text hi pugui passar
    # per sobre i seguir llegint-se.
    EscutOpacitat  = 0.35
    EscutFitxer    = 'cornella.ico'
    MargePt        = 2
}

# Intents de generació de la imatge, EN ORDRE de preferència. Es va provant fins
# que el base64 hi cap (vegeu $Script:MaxCaixetiBase64).
#
# NOMÉS JPEG. Es va provar de fer-ho en PNG (no perd qualitat i, amb text pla,
# comprimeix molt millor: la mateixa imatge feia 24.352 caràcters en PNG i
# 35.060 en JPEG) i **AutoFirma no el va acceptar**: el caixetí no va sortir i la
# signatura va caure al respatller de text d'una línia. No ho diu cap missatge
# d'error — simplement l'intent amb imatge falla. Per tant: la rúbrica es passa
# SEMPRE en JPEG, que és l'únic format que s'ha vist funcionar de debò.
#
# La nitidesa, doncs, ha de sortir de la RESOLUCIÓ i de la QUALITAT, no del
# format: abans anava a escala x2 i qualitat 70, i el JPEG a qualitat baixa
# deixa halos al voltant del text negre sobre blanc — això era la borrositat.
# Es va de mes a menys i es fa servir el PRIMER que hi capiga; hi ha forca
# graons perque no se sap exactament quant ocupara cada JPEG fins que el Windows
# no el genera (el seu codificador no dona la mateixa mida que cap altre).
# Mesurat de debo: al Windows de l'usuari, l'escala x3 amb qualitat 92 ha ocupat
# 26.416 caracters (el JPEG del Windows surt ~20% mes petit del que estimaven les
# meves proves). Amb el topall a 30.500, doncs, l'escala x4 hi cap fins a
# qualitat 85, i mes resolucio val mes que mes qualitat quan el que hi ha es
# TEXT: 800x192 son 288 ppp, davant dels 216 de la x3.
$Script:CaixetiIntents = @(
    @{ Format = 'jpeg'; Escala = 4; Qualitat = 88 }
    @{ Format = 'jpeg'; Escala = 4; Qualitat = 85 }
    @{ Format = 'jpeg'; Escala = 4; Qualitat = 80 }
    @{ Format = 'jpeg'; Escala = 3; Qualitat = 92 }
    @{ Format = 'jpeg'; Escala = 3; Qualitat = 86 }
    @{ Format = 'jpeg'; Escala = 2; Qualitat = 95 }
)

# Quin intent ha entrat, per al registre. Serveix per saber, sense endevinar, si
# el caixetí ha sortit com a imatge i amb quina resolució.
$Script:CaixetiUltimIntent = ''
# Si l'escut no s'ha pogut dibuixar, per què. Va al registre: la primera vegada
# l'escut no sortia i no ho deia enlloc, perquè el dibuix va dins d'un try/catch.
$Script:CaixetiAvisEscut = ''
# Cert si l'escut s'ha arribat a pintar. Igual que l'avis: sense aixo, que no hi
# fos no es notava fins que algu mirava el PDF amb lupa.
$Script:CaixetiEscutDibuixat = $false

# Límit REAL de Windows per a una línia d'ordres: 32767 caràcters. Deixem marge
# per a les rutes (que poden ser llargues, en xarxa) i per al filtre del
# certificat. Si un intent no hi cap, se salta (mai es prova i peta).
$Script:MaxCommandLine = 32000
# Mida màxima del base64 de la imatge del caixetí. Amb el marge de dalt, la
# imatge no pot passar d'aquí o l'ordre no hi cabria.
#
# Eren 20000 "per si de cas", i era MOLT curt: amb l'escut de fons només hi
# cabia l'escala x2, que és justament la borrosa. Amb una ordre real del registre
# de l'usuari (rutes a la unitat de xarxa `I:\...\5.- Sergi Fadurdo\...`, filtre
# amb el CN sencer i les 6 propietats de posició), tot el que NO és la imatge
# ocupa **628 caràcters**. O sigui que del límit dur de Windows (32767) en sobren
# més de 30.000 per a la imatge, i el topall es pot pujar sense por.
# Si algun dia una ruta fos molt més llarga, la comprovació de
# $Script:MaxCommandLine salta l'intent abans d'executar-lo: aquest topall no és
# l'única xarxa de seguretat.
$Script:MaxCaixetiBase64 = 30500

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
#
# NO LA CRIDIS FORA DE WINDOWS: el guard ha d'anar al CRIDADOR (vegeu
# _AutoFirmaVisibleExtraParams). Motiu: PowerShell compila el cos sencer de la
# funcio la primera vegada que s'invoca, i en compilar-lo resol els literals de
# tipus [System.Drawing.*]; aixo dispara l'inicialitzador estatic de GDI+, que
# fora de Windows llanca PlatformNotSupportedException embolcallada en una
# TypeInitializationException. Aquesta excepcio surt ABANS de la primera linia
# del cos, o sigui que ni un 'if' de guard aqui dins ni el try/catch la poden
# aturar: s'enduia tota la suite de proves quan s'executa en un Linux.
function _BuildCaixetiImageBase64([string]$caixeti) {
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $lines = @((([string]$caixeti -replace "`r`n", "`n") -split "`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -eq 0) { return '' }
        # Es prova de mes nitid a menys, i ens quedem amb el PRIMER que hi cap:
        # la imatge viatja en base64 dins de la linia d'ordres, i Windows no
        # admet mes de 32767 caracters (amb l'escala x4 en JPEG, AutoFirma ni
        # s'arrencava: "El nombre del archivo o la extension es demasiado
        # largo"). Aixi es fa servir sempre la millor qualitat possible en lloc
        # d'anar a la fixa amb la pitjor.
        $Script:CaixetiUltimIntent = ''
        $Script:CaixetiAvisEscut = ''
        $Script:CaixetiEscutDibuixat = $false
        $provats = New-Object System.Collections.ArrayList
        foreach ($intent in $Script:CaixetiIntents) {
            $b64 = _CaixetiImatgeIntent $lines $intent
            $etiq = ([string]$intent.Format + ' x' + [string]$intent.Escala +
                     $(if ($null -ne $intent.Qualitat) { ' q' + [string]$intent.Qualitat } else { '' }))
            if ([string]::IsNullOrWhiteSpace($b64)) { [void]$provats.Add($etiq + ': no s''ha pogut dibuixar'); continue }
            [void]$provats.Add($etiq + ': ' + $b64.Length + ' car.')
            if ($b64.Length -le $Script:MaxCaixetiBase64) {
                $Script:CaixetiUltimIntent = $etiq + ' (' + $b64.Length + ' car.)'
                if (-not [string]::IsNullOrWhiteSpace($Script:CaixetiAvisEscut)) {
                    $Script:CaixetiUltimIntent += ' [' + $Script:CaixetiAvisEscut + ']'
                } elseif (-not $Script:CaixetiEscutDibuixat) {
                    $Script:CaixetiUltimIntent += ' [SENSE escut]'
                }
                return $b64
            }
        }
        # Cap no hi cap: es deixa dit al registre amb les mides de tots, que es
        # l'unica manera de saber per que el caixeti ha sortit en text.
        $Script:CaixetiUltimIntent = 'CAP (topall ' + $Script:MaxCaixetiBase64 + '): ' + ($provats -join ' | ')
        return ''
    } catch {
        return ''
    }
}

# Ruta de l'escut de l'Ajuntament. Es al costat del codi, a suport\.
function _CaixetiEscutPath {
    $d = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($d)) { $d = Join-Path $RepoRoot 'suport' }
    return [string](Join-Path $d $Script:CaixetiEstil.EscutFitxer)
}

# Tria la millor imatge de dins d'un .ico i en retorna @{ Offset; Mida; Amplada;
# EsPng } (o $null si el fitxer no es un .ico valid).
#
# Per que ens ho fem nosaltres: un .ico es un CONTENIDOR amb diverses mides a
# dins, i el de l'Ajuntament les porta TOTES set comprimides en PNG (16, 24, 32,
# 48, 64, 128 i 256 px). El .NET, amb icones aixi, va maldestre:
# Icon.ToBitmap() no les descomprimeix be i el resultat surt buit — que es
# exactament el que passava, l'escut no apareixia al caixeti i no ho deia
# ningu, perque el dibuix va dins d'un try/catch. Llegint nosaltres la taula del
# .ico podem agafar el PNG que ens convé i passar-lo a Image.FromStream, que si
# que el sap llegir.
#
# Format del .ico: capcalera de 6 bytes (reservat, tipus, nombre d'imatges) i
# despres una entrada de 16 bytes per imatge; l'amplada i l'alcada hi van en UN
# sol byte, i el 0 vol dir 256. Funcio PURA (rep els bytes).
function _IcoTriaFrame($bytes, [int]$midaVolguda) {
    if ($null -eq $bytes -or $bytes.Length -lt 22) { return $null }
    if ($bytes[0] -ne 0 -or $bytes[1] -ne 0 -or $bytes[2] -ne 1 -or $bytes[3] -ne 0) { return $null }
    $n = [int]$bytes[4] + ([int]$bytes[5] * 256)
    if ($n -le 0) { return $null }
    $millor = $null
    for ($i = 0; $i -lt $n; $i++) {
        $o = 6 + ($i * 16)
        if (($o + 16) -gt $bytes.Length) { break }
        $ampl = [int]$bytes[$o]
        if ($ampl -eq 0) { $ampl = 256 }
        $mida = [BitConverter]::ToInt32($bytes, $o + 8)
        $desp = [BitConverter]::ToInt32($bytes, $o + 12)
        if ($mida -le 0 -or $desp -lt 0 -or ($desp + $mida) -gt $bytes.Length) { continue }
        $esPng = ($mida -gt 8 -and $bytes[$desp] -eq 0x89 -and $bytes[$desp + 1] -eq 0x50 -and
                  $bytes[$desp + 2] -eq 0x4E -and $bytes[$desp + 3] -eq 0x47)
        $cand = @{ Offset = $desp; Mida = $mida; Amplada = $ampl; EsPng = $esPng }
        if ($null -eq $millor) { $millor = $cand; continue }
        # La mes petita que ja sigui prou gran; si cap no hi arriba, la mes gran
        # (val mes reduir una imatge gran que no pas estirar-ne una de petita).
        $mA = [int]$millor.Amplada
        if ($mA -lt $midaVolguda) {
            if ($ampl -gt $mA) { $millor = $cand }
        } elseif ($ampl -ge $midaVolguda -and $ampl -lt $mA) {
            $millor = $cand
        }
    }
    return $millor
}

# Dibuixa el caixeti amb UN intent concret i el retorna en base64 ('' si falla).
# Separat de _BuildCaixetiImageBase64 perque aquella nomes decideix quin intent
# es queda; aqui hi ha tot el dibuix.
function _CaixetiImatgeIntent($lines, $intent) {
    $p    = $Script:AutoFirmaCaixetiPos
    $est  = $Script:CaixetiEstil
    $esc  = [int]$intent.Escala
    $w    = [int](([int]$p.URX - [int]$p.LLX) * $esc)
    $h    = [int](([int]$p.URY - [int]$p.LLY) * $esc)
    if ($w -le 0 -or $h -le 0) { return '' }

    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    try {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.Clear([System.Drawing.Color]::White)
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            # AntiAlias (no AntiAliasGridFit): la imatge es reescala dins del PDF,
            # i l'ajust a la graella de pixels que fa el GridFit hi queda pitjor.
            $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

            $marge = [int]([double]$est.MargePt * $esc)

            # --- ESCUT de fons, a la dreta. Va PRIMER perque el text hi passi per
            #     sobre (l'usuari ja ho ha donat per bo). Amb poca opacitat, si no
            #     es menja les lletres.
            try {
                $ico = _CaixetiEscutPath
                if (Test-Path -LiteralPath $ico) {
                    $sz = $h - (2 * $marge)
                    if ($sz -gt 0) {
                        # Llegim el .ico a ma i n'agafem el PNG que toca (vegeu
                        # _IcoTriaFrame): amb Icon.ToBitmap() l'escut sortia buit.
                        $raw = [System.IO.File]::ReadAllBytes($ico)
                        $fr = _IcoTriaFrame $raw $sz
                        $imgE = $null
                        if ($null -ne $fr -and $fr.EsPng) {
                            $tros = New-Object byte[] ([int]$fr.Mida)
                            [Array]::Copy($raw, [int]$fr.Offset, $tros, 0, [int]$fr.Mida)
                            $msE = New-Object System.IO.MemoryStream(, $tros)
                            $imgE = [System.Drawing.Image]::FromStream($msE)
                        }
                        # Respatller: si no era PNG (un .ico amb imatges DIB de
                        # tota la vida), que ho provi el .NET a la seva manera.
                        if ($null -eq $imgE) {
                            $icona = New-Object System.Drawing.Icon($ico, (New-Object System.Drawing.Size($sz, $sz)))
                            try { $imgE = $icona.ToBitmap() } finally { $icona.Dispose() }
                        }
                        if ($null -ne $imgE) {
                            try {
                                # Matrix33 es el canal ALFA: amb 1 es veuria opac.
                                $cmx = New-Object System.Drawing.Imaging.ColorMatrix
                                $cmx.Matrix33 = [single]$est.EscutOpacitat
                                $ia = New-Object System.Drawing.Imaging.ImageAttributes
                                $ia.SetColorMatrix($cmx)
                                $rc = New-Object System.Drawing.Rectangle(($w - $marge - $sz), $marge, $sz, $sz)
                                $g.DrawImage($imgE, $rc, 0, 0, $imgE.Width, $imgE.Height,
                                             [System.Drawing.GraphicsUnit]::Pixel, $ia)
                                $ia.Dispose()
                                $Script:CaixetiEscutDibuixat = $true
                            } finally { $imgE.Dispose() }
                        }
                    }
                }
            } catch {
                # No es prou greu per deixar el caixeti sense fer, pero SI que
                # s'ha de saber: sense aixo, l'escut no sortia i no ho deia ningu.
                $Script:CaixetiAvisEscut = "no s'ha pogut dibuixar l'escut: " + $_.Exception.Message
            }

            # --- CONTORN gris fosc. El requadre es dibuixa cap endins mig gruix:
            #     si no, la meitat del trac cauria fora de la imatge i es veuria
            #     mes prim per dos costats.
            try {
                $gruix = [single][Math]::Max(1, ([int]$est.ContornGruix * $esc))
                $col = [System.Drawing.Color]::FromArgb([int]$est.ContornRGB[0], [int]$est.ContornRGB[1], [int]$est.ContornRGB[2])
                $pen = New-Object System.Drawing.Pen($col, $gruix)
                try {
                    $o = $gruix / 2.0
                    $g.DrawRectangle($pen, [single]$o, [single]$o, [single]($w - $gruix), [single]($h - $gruix))
                } finally { $pen.Dispose() }
            } catch { }

            # --- TEXT. L'alcada de linia es reparteix l'espai util i la lletra
            #     n'ocupa FactorLletra: com mes alt el factor, menys espai buit
            #     queda entre linia i linia.
            $util = $h - (2 * $marge)
            $alcadaLinia = [double]$util / [double]@($lines).Count
            $mida = [int][Math]::Max(6, [Math]::Floor($alcadaLinia * [double]$est.FactorLletra))
            $font = New-Object System.Drawing.Font('Arial', $mida, [System.Drawing.GraphicsUnit]::Pixel)
            try {
                # GenericTypographic: sense el farciment extra que hi posa el
                # format de defecte, que separava mes les linies.
                $sf = [System.Drawing.StringFormat]::GenericTypographic
                $y = [double]$marge
                foreach ($l in $lines) {
                    $g.DrawString([string]$l, $font, [System.Drawing.Brushes]::Black,
                                  [single]($marge + $esc), [single]($y + ($alcadaLinia - $mida) / 2.0), $sf)
                    $y += $alcadaLinia
                }
            } finally { $font.Dispose() }
        } finally { $g.Dispose() }

        $ms = New-Object System.IO.MemoryStream
        try {
            if ([string]$intent.Format -eq 'png') {
                $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            } else {
                $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
                         Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
                if ($null -ne $codec) {
                    $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
                    $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
                                        [System.Drawing.Imaging.Encoder]::Quality, [long]$intent.Qualitat)
                    $bmp.Save($ms, $codec, $ep)
                    $ep.Dispose()
                } else {
                    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                }
            }
            return [System.Convert]::ToBase64String($ms.ToArray())
        } finally { $ms.Dispose() }
    } catch {
        return ''
    } finally { $bmp.Dispose() }
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
        # El caixeti-imatge necessita System.Drawing, que nomes hi es a Windows.
        # El guard va AQUI i no dins de _BuildCaixetiImageBase64: alli no serviria
        # (l'excepcio del carregador de tipus salta en compilar-ne el cos, abans
        # de la primera linia). Fora de Windows -> '' i el crider prova el text.
        if ($env:OS -ne 'Windows_NT') { return '' }
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
#
# Els camins es munten amb [System.IO.Path]::Combine, NO amb Join-Path: Join-Path
# resol la UNITAT i, en provar-ho fora de Windows, peta amb "Cannot find drive. A
# drive with the name 'C' does not exist" (i, a sobre, torna la ruta dins d'un
# PSObject). Combine és un mètode .NET: no toca el sistema de fitxers.
function _AutoFirmaCandidatePaths {
    $pf  = if ([string]::IsNullOrWhiteSpace($env:ProgramFiles))       { 'C:\Program Files' }       else { $env:ProgramFiles }
    $px  = if ([string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) { 'C:\Program Files (x86)' } else { ${env:ProgramFiles(x86)} }
    $paths = @(
        [System.IO.Path]::Combine($pf, 'AutoFirma\AutoFirma\AutoFirma.exe')
        [System.IO.Path]::Combine($pf, 'AutoFirma\AutoFirma.exe')
        [System.IO.Path]::Combine($px, 'AutoFirma\AutoFirma\AutoFirma.exe')
        [System.IO.Path]::Combine($px, 'AutoFirma\AutoFirma.exe')
    )
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $paths += [System.IO.Path]::Combine($env:LOCALAPPDATA, 'AutoFirma\AutoFirma\AutoFirma.exe')
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
        $local = Get-LocalSubdir $RepoRoot 'Informes'
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
    $def = @{ folder = ''; sign = $false; certFilter = ''; autofirma = ''; overwrite = $false; visibleSign = $true; caixeti = (_DefaultCaixeti); obrirRegistre = $false }
    if (-not (Test-Path -LiteralPath $p)) { return $def }
    try {
        $o = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @('folder','certFilter','autofirma','caixeti')) { if ($o.PSObject.Properties[$k]) { $def[$k] = [string]$o.$k } }
        if ($o.PSObject.Properties['sign'])        { $def['sign'] = [bool]$o.sign }
        if ($o.PSObject.Properties['overwrite'])   { $def['overwrite'] = [bool]$o.overwrite }
        if ($o.PSObject.Properties['visibleSign']) { $def['visibleSign'] = [bool]$o.visibleSign }
        if ($o.PSObject.Properties['obrirRegistre']) { $def['obrirRegistre'] = [bool]$o.obrirRegistre }
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

    # El registre de la signatura es DIAGNOSTIC: abans, en acabar, sortia una
    # pregunta de si es volia obrir, i preguntar-ho cada vegada fa nosa. Ara es
    # una casella d'aqui, que es recorda: qui el vol, el marca i prou.
    $cbLog = New-Object System.Windows.Forms.CheckBox
    $cbLog.Text = 'Obrir el registre de la signatura en acabar'
    $cbLog.Location = New-Object System.Drawing.Point(34, $y)
    $cbLog.AutoSize = $true
    $cbLog.Checked = [bool]$st.obrirRegistre
    [void]$form.Controls.Add($cbLog)
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
            ObrirRegistre = [bool]$cbLog.Checked
        }
        _SavePdfSignarState @{ folder = $f; sign = [bool]$cbSign.Checked; certFilter = $certFilter; autofirma = [string]$autofirma; overwrite = [bool]$cbOver.Checked; visibleSign = [bool]$cbVis.Checked; caixeti = $caixeti; obrirRegistre = [bool]$cbLog.Checked }
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
                        # Quina variant d'imatge s'ha fet servir (o per que cap):
                        # sense aixo, quan el caixeti sortia en text calia
                        # endevinar si era la mida o el format.
                        if ([string]$intent.Mode -eq 'imatge' -and -not [string]::IsNullOrWhiteSpace($Script:CaixetiUltimIntent)) {
                            _PdfSignarLog ("   imatge: " + $Script:CaixetiUltimIntent)
                        }
                    }
                    if ($usat -eq 'sense caixeti' -and $res.ExitCode -eq 0) { $senseCaixeti++ }
                    if ($res.ExitCode -eq 0 -and (Test-Path -LiteralPath $tmpSigned)) {
                        Move-Item -LiteralPath $tmpSigned -Destination $pdf -Force
                        $signed++
                        # Nomes la primera: serveix per comprovar l'ordre exacta.
                        if ($signed -eq 1) {
                            _PdfSignarLog ("OK (" + $usat + ")  " + $f.Name + "  ::  " + (_AutoFirmaArgvToText $argv))
                            if ($usat -like 'caixeti (imatge)*' -and -not [string]::IsNullOrWhiteSpace($Script:CaixetiUltimIntent)) {
                                _PdfSignarLog ("   imatge: " + $Script:CaixetiUltimIntent)
                            }
                        }
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
    $icon = if ($errors -gt 0) { 'Warning' } else { 'Information' }
    [System.Windows.Forms.MessageBox]::Show($msg.ToString(), 'Convertir informes a PDF', 'OK', $icon) | Out-Null

    # El registre porta l'ordre EXACTA que s'ha passat a AutoFirma i serveix per
    # veure per que no surt el caixeti, si es el cas. Abans es preguntava en
    # acabar CADA VEGADA, i preguntar-ho sempre fa nosa; ara hi ha la casella
    # 'Obrir el registre de la signatura en acabar' al diàleg d'opcions, que es
    # recorda: qui el vol, el marca i prou.
    if ($opts.ObrirRegistre -and $opts.Sign -and ($signed -gt 0 -or $errors -gt 0)) {
        try { Start-Process -FilePath (_PdfSignarLogPath) | Out-Null } catch { }
    }
}

# Punt d'entrada de l'eina (des del menú principal).
function Invoke-ConvertirPdf {
    $opts = _ShowConvertPdfOptions
    if ($null -eq $opts) { return }
    _RunConvertPdf $opts
}
