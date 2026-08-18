#requires -Version 5.1
<#
.SYNOPSIS
  Refa la SIGNATURA de dins d'un PDF ja signat, exactament com la fa l'Adobe.

.DESCRIPTION
  PER QUE EXISTEIX AIXO. Els informes signats amb l'AutoFirma es validaven al PC
  de l'usuari pero a qualsevol altre hi sortia "La validez de la firma es
  DESCONOCIDA". Un informe signat a ma amb l'Adobe, AMB EL MATEIX CERTIFICAT i
  al MATEIX ordinador, si que hi sortia valid. Comparant els dos fitxers es van
  anar descartant, un a un i amb mesures:

    - el certificat            -> identic als dos (mateix numero de serie)
    - el magatzem de confianca -> mateix ordinador i mateix Adobe: no era aixo
    - la cadena encastada      -> es va deixar nomes el signant: igual
    - el SubFilter             -> es va posar adbe.pkcs7.detached: igual
    - l'OID de l'algorisme     -> provat canviant NOMES aquell byte del PDF ja
                                  signat (no esta cobert per la signatura;
                                  l'OpenSSL seguia validant): igual
    - la versio de l'atribut ESS (signingCertificateV2=false): igual

  L'unica cosa que quedava era que l'AutoFirma hi posa un atribut ESS
  ('signingCertificateV2') que porta a dins un camp 'policies' amb les
  politiques del certificat. El RFC 5035 diu que aixo NO es informatiu: obliga
  el validador a validar la cadena RESTRINGIDA a aquelles politiques. Al PC de
  l'usuari, amb tota la cadena de l'AOC instal·lada, passa; en un altre, no.

  QUE ES FA AQUI, I PER QUE ES SEGUR. En lloc de perseguir mes diferencies, es
  deixa de negociar amb l'AutoFirma: ell munta el PDF (i aixo ho fa be -el
  document mai surt alterat, el camp de signatura i el caixeti son correctes-) i
  NOSALTRES li reemplacem NOMES el CMS de dins del /Contents per un de fet aqui,
  amb el mateix certificat i amb l'estructura EXACTA de l'Adobe:

      atributs signats : contentType + messageDigest (+ revocationInfoArchival
                         buit, que es el que hi posa l'Adobe)
      certificats      : nomes el del signant
      cap atribut ESS  : ni signingCertificate ni signingCertificateV2

  L'OID de l'algorisme del SignerInfo (el .NET hi posa 'rsaEncryption' i l'Adobe
  'sha256WithRSAEncryption') es deixa TAL COM SURT: es va comprovar canviant
  nomes aquell byte a un PDF ja signat que no canvia gens la validesa. Es va
  arribar a escriure la funcio que ho corregia i la prova la va enxampar
  corrompent el CERTIFICAT (el patro que buscava tambe apareix a la clau publica
  del certificat, i el del SignerInfo no porta el NULL). Codi que no cal, fora.

  Aixo no toca ni un byte del document: el /ByteRange no canvia, el forat del
  /Contents ja hi es (l'AutoFirma el deixa de 27.000 bytes llargs i el nostre
  CMS n'ocupa uns 2.600) i el que es firma son EXACTAMENT els mateixos bytes.
  Per tant el "no ha habido modificaciones" segueix sortint igual.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES (es proven en headless, sense Windows ni certificats)
# ----------------------------------------------------------------------------

# Localitza la signatura dins d'un PDF: el /ByteRange i el forat hexadecimal del
# /Contents. Retorna @{ Ok; Ranges; HexStart; HexLen } (HexStart/HexLen NO
# inclouen els '<' '>').  Funcio PURA.
function _PdfTrobaFirma([byte[]]$bytes) {
    $buit = @{ Ok = $false; Ranges = @(); HexStart = 0; HexLen = 0; Motiu = '' }
    if ($null -eq $bytes -or $bytes.Length -lt 32) { $buit.Motiu = 'fitxer buit'; return $buit }
    # Latin-1: cada byte -> un caracter. Aixi les posicions de la cadena son les
    # del fitxer i no cal anar amb compte amb cap codificacio.
    $txt = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
    $mBR = [regex]::Match($txt, '/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]')
    if (-not $mBR.Success) { $buit.Motiu = 'no hi ha /ByteRange (el PDF no esta signat)'; return $buit }
    $r = @([int]$mBR.Groups[1].Value, [int]$mBR.Groups[2].Value, [int]$mBR.Groups[3].Value, [int]$mBR.Groups[4].Value)
    # El forat del /Contents es, per definicio, entre el final del 1r tros i el
    # principi del 2n. Aixi no cal endevinar on es el '<': el diu el ByteRange.
    $ini = $r[0] + $r[1]
    $fi  = $r[2]
    if ($fi -le $ini -or $fi -gt $bytes.Length) { $buit.Motiu = 'el /ByteRange no quadra amb la mida del fitxer'; return $buit }
    if ($bytes[$ini] -ne 0x3C) { $buit.Motiu = "al forat del /Contents no hi ha cap '<'"; return $buit }
    if ($bytes[$fi - 1] -ne 0x3E) { $buit.Motiu = "al forat del /Contents no hi ha cap '>'"; return $buit }
    return @{ Ok = $true; Ranges = $r; HexStart = ($ini + 1); HexLen = ($fi - $ini - 2); Motiu = '' }
}

# Els bytes que la signatura cobreix: els dos trossos del /ByteRange, seguits.
# Funcio PURA.
function _PdfContingutSignat([byte[]]$bytes, $ranges) {
    $r = @($ranges)
    $out = New-Object byte[] ($r[1] + $r[3])
    [Array]::Copy($bytes, $r[0], $out, 0, $r[1])
    [Array]::Copy($bytes, $r[2], $out, $r[1], $r[3])
    return ,$out
}

# Escriu el CMS al forat del /Contents, en hexadecimal i farcint amb zeros fins
# a omplir-lo. La MIDA DEL FITXER NO CANVIA: si canvies, el /ByteRange (que ja
# esta escrit i firmat) deixaria de quadrar. Funcio PURA.
function _PdfPosaCms([byte[]]$bytes, [int]$hexStart, [int]$hexLen, [byte[]]$cms) {
    if ($cms.Length * 2 -gt $hexLen) {
        throw ("El CMS ocupa " + ($cms.Length * 2) + " caracters i al forat del /Contents nomes n'hi caben " + $hexLen + ".")
    }
    $out = New-Object byte[] $bytes.Length
    [Array]::Copy($bytes, $out, $bytes.Length)
    $hex = [System.BitConverter]::ToString($cms).Replace('-', '').ToLower()
    for ($i = 0; $i -lt $hexLen; $i++) {
        $c = if ($i -lt $hex.Length) { [byte][char]$hex[$i] } else { [byte][char]'0' }
        $out[$hexStart + $i] = $c
    }
    return ,$out
}

# ----------------------------------------------------------------------------
# EL CMS (necessita el certificat; es pot provar amb un de fals)
# ----------------------------------------------------------------------------

# Un CMS DESLLIGAT (detached) amb la mateixa forma que el que fa l'Adobe.
#
# ELS ATRIBUTS SIGNATS: si no se n'hi posa CAP, el .NET firma el contingut
# directament i el PDF queda sense el 'messageDigest' que l'Adobe espera. Per
# aixo s'hi posa l'unic que hi posa l'Adobe -'adbe-revocationInfoArchival', amb
# el valor BUIT- i llavors el .NET hi afegeix sol el contentType i el
# messageDigest. Resultat: els mateixos tres atributs que l'Adobe, i cap atribut
# ESS (que es el que fa que la signatura no es validi enlloc mes).
# Carrega l'assemblatge dels tipus PKCS. Al PowerShell 5.1 (Windows) viuen a
# 'System.Security'; al PowerShell 7 (que es el que fa anar les proves), a
# 'System.Security.Cryptography.Pkcs'. Es proven els dos i s'ignora el que falli.
function _CmsCarregaTipus {
    foreach ($nom in @('System.Security', 'System.Security.Cryptography.Pkcs')) {
        try { Add-Type -AssemblyName $nom -ErrorAction SilentlyContinue } catch { }
    }
}

function _CmsComAdobe([byte[]]$contingut, $cert) {
    _CmsCarregaTipus
    $ci = New-Object System.Security.Cryptography.Pkcs.ContentInfo (,$contingut)
    $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms ($ci, $true)
    $signer = New-Object System.Security.Cryptography.Pkcs.CmsSigner ($cert)
    # Per defecte el .NET fa SHA-1: s'ha de dir explicitament.
    $signer.DigestAlgorithm = New-Object System.Security.Cryptography.Oid '2.16.840.1.101.3.4.2.1'
    $signer.IncludeOption = [System.Security.Cryptography.X509Certificates.X509IncludeOption]::EndCertOnly
    # adbe-revocationInfoArchival amb valor buit (SEQUENCE {}), com l'Adobe.
    $oidRev = New-Object System.Security.Cryptography.Oid '1.2.840.113583.1.1.8'
    $val = New-Object System.Security.Cryptography.AsnEncodedData ($oidRev, [byte[]](0x30, 0x00))
    [void]$signer.SignedAttributes.Add((New-Object System.Security.Cryptography.CryptographicAttributeObject ($oidRev, (New-Object System.Security.Cryptography.AsnEncodedDataCollection ($val)))))
    $cms.ComputeSignature($signer, $false)
    return ,($cms.Encode())
}

# ----------------------------------------------------------------------------
# L'OPERACIO SENCERA
# ----------------------------------------------------------------------------

# Refa la signatura d'un PDF ja signat. Retorna @{ Ok; Motiu }.
# Si algun pas falla NO es toca el fitxer: val mes deixar-hi la signatura de
# l'AutoFirma (que al PC de l'usuari es valida) que espatllar-lo.
function Repack-PdfFirmaComAdobe([string]$path, $cert) {
    if ($null -eq $cert) { return @{ Ok = $false; Motiu = 'no s''ha triat cap certificat' } }
    if (-not (Test-Path -LiteralPath $path)) { return @{ Ok = $false; Motiu = 'el PDF no hi es' } }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $inf = _PdfTrobaFirma $bytes
        if (-not $inf.Ok) { return @{ Ok = $false; Motiu = $inf.Motiu } }
        $contingut = _PdfContingutSignat $bytes $inf.Ranges
        $cms = _CmsComAdobe $contingut $cert
        $nou = _PdfPosaCms $bytes $inf.HexStart $inf.HexLen $cms
        [System.IO.File]::WriteAllBytes($path, $nou)
        return @{ Ok = $true; Motiu = ('CMS refet, ' + $cms.Length + ' bytes') }
    } catch {
        return @{ Ok = $false; Motiu = $_.Exception.Message }
    }
}
