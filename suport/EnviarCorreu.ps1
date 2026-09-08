#requires -Version 5.1
<#
  EnviarCorreu.ps1 - Eina "Enviar correu" (secció MÒBIL).

  Envia el correu de requeriments des del PC, amb el MATEIX format i remitent que
  el mòbil (mateixa plantilla d'EmailJS). El cos de requeriments es construeix
  llegint el .docx GENERAT (estructura REQ1: seccions, subseccions, numeració i
  qualsevol edició manual), i s'embolcalla amb les condicions del correu (text
  introductori + text final, SENSE conclusions) definides a
  docs/dades/email-textos.json.

  Enviament: API REST d'EmailJS des de PowerShell. La Public key / Service ID /
  Template ID es llegeixen de docs/config.js (no secretes). La Private key es
  desa a la carpeta local/ del repositori (gitignored): local/emailjs.json ->
  { "private_key": "..." }.
#>

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

function _CorreuRepoRoot {
    if ($RepoRoot) { return [string]$RepoRoot }
    return (Split-Path -Parent $PSScriptRoot)
}

# --- Configuració (claus d'EmailJS) ------------------------------------------
function _CorreuConfig {
    $repo = _CorreuRepoRoot
    # Un sol lector (ConfigJs.ps1), no quatre regex escampats: cadascun exigia
    # cometes dobles i el nom literal, i amb qualsevol reformat del .js la clau
    # es quedava buida EN SILENCI. Ara hi ha prova contra el fitxer de debo.
    $cjs = Read-ConfigJs
    $pub = Get-ConfigJsValue $cjs 'EMAILJS_PUBLIC_KEY'
    $svc = Get-ConfigJsValue $cjs 'EMAILJS_SERVICE_ID'
    $tpl = Get-ConfigJsValue $cjs 'EMAILJS_TEMPLATE_ID'
    $from = Get-ConfigJsValue $cjs 'EMAIL_FROM_NAME' 'Ajuntament de Cornellà de Llobregat - Activitats'
    # Private key: carpeta local/ (fora del repositori public).
    $priv = ''
    $pkPath = Join-Path $repo (Join-Path 'local' 'emailjs.json')
    $j = Read-JsonFile $pkPath
    if ($null -ne $j -and $j.private_key) { $priv = [string]$j.private_key }
    return [pscustomobject]@{
        PublicKey = $pub; ServiceId = $svc; TemplateId = $tpl; FromName = $from; PrivateKey = $priv
        PrivatePath = $pkPath
    }
}

# --- Utils de text -> HTML ----------------------------------------------------
function _EscHtml($s) {
    return ([string]$s).Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}
# Escapa, **negreta** -> <b>, //cursiva// -> <i>, i enllaça els URLs.
#
# ELS URLs S'APARTEN ABANS DE MIRAR LA CURSIVA, i no es un detall d'estil: la
# cursiva es "//...//" i un "https://" en porta un de "//" a dins. Amb DUES
# adreces a la mateixa linia, l'expressio de la cursiva es menjava tot el tros
# d'una a l'altra:
#
#   Mira https://a.cat i tambe https://b.cat
#   -> Mira https:<i>a.cat i tambe https:</i>b.cat        (les dues destrossades)
#
# No era hipotetic: aquesta funcio ja la fan servir els recordatoris, i el text
# el pot editar l'usuari. Ara cada URL es substitueix per una marca amb caracters
# de control -que no poden sortir en un text escrit a ma i que cap de les dues
# expressions toca- i es torna a posar, ja com a enllac, al final.
function _TextToHtml($s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $h = _EscHtml $s

    $urls = New-Object System.Collections.ArrayList
    $h = [regex]::Replace($h, '(https?://[^\s<]+)', {
        param($m)
        $i = $urls.Add($m.Groups[1].Value)
        return ([char]1 + [string]$i + [char]1)
    })

    $h = [regex]::Replace($h, '\*\*(.+?)\*\*', '<b>$1</b>')
    $h = [regex]::Replace($h, '//(.+?)//', '<i>$1</i>')

    for ($i = 0; $i -lt $urls.Count; $i++) {
        $u = [string]$urls[$i]
        $h = $h.Replace(([char]1 + [string]$i + [char]1), ('<a href="' + $u + '">' + $u + '</a>'))
    }
    return $h.Replace("`r`n","`n").Replace("`n",'<br>')
}

# El COS sencer a HTML: una linia = un <div>, i una linia buida un espaiador.
#
# Aixo estava DUPLICAT linia a linia -_RecCosHtml (Recordatoris) i
# _ControlsCpEmailHtml (Controls periodics), amb el mateix estil inline inclos-.
# L'unica diferencia era la funcio de linia: els recordatoris ja passaven per
# _TextToHtml i els controls periodics per un _ControlsCpLineHtml propi que feia
# el mateix pero SENSE cursiva. Ara passen tots dos per _TextToHtml, o sigui que
# els controls periodics guanyen la cursiva (i la proteccio dels URLs de dalt).
function _CosAHtml([string]$cos) {
    $lines = (([string]$cos) -replace "`r`n", "`n") -split "`n"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div style="font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#1d2733;line-height:1.4">')
    foreach ($ln in $lines) {
        if ([string]::IsNullOrWhiteSpace($ln)) {
            [void]$sb.Append('<div style="height:8px;line-height:8px">&nbsp;</div>')
        } else {
            [void]$sb.Append('<div>' + (_TextToHtml $ln) + '</div>')
        }
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

# Substitueix les variables {X} d'un text. El MAPA el posa cada crider: les
# claus i d'on surten els valors son diferents de debo a cada eina (una fila de
# controls periodics, una de recordatoris, la capcalera d'un informe) i no s'han
# d'unificar. El que estava copiat tres vegades era NOMES aquest bucle.
function _OmpleVariables([string]$text, $mapa) {
    $t = [string]$text
    if ($null -eq $mapa) { return $t }
    foreach ($k in @($mapa.Keys)) { $t = $t.Replace([string]$k, [string]$mapa[$k]) }
    return $t
}

# --- Llegir el cos de requeriments del .docx generat --------------------------
# Recorre els paragrafs: comença despres de la intro de deficiencies (o del
# primer titol/numero) i acaba abans de les conclusions. Torna HTML.
function _DocxRequerimentsHtml($docxPath) {
    $phrases = if ($SeguimentConclusionPhrases) { @($SeguimentConclusionPhrases) } else {
        @("Vist l'anterior", 'Ho poso al seu coneixement', 'Cornella de Llobregat,', 'CONCLUSIONS')
    }
    $phrN = $phrases | ForEach-Object { _NormalitzaText $_ }

    # New-WordApp (Motor.ps1) i no un New-Object a pel: aqui no hi havia CAP
    # guarda del $null -New-Object -ComObject pot tornar $null sense llancar- i
    # llavors el '$word.Visible' de sota petava amb un "metode sobre NULL" que
    # arribava tal qual al quadre d'error del crider. A mes, New-WordApp posa
    # AutomationSecurity = 1, que es el que evita que el Word obri en VISTA
    # PROTEGIDA els fitxers d'una unitat de xarxa -i els informes hi son-.
    $word = New-WordApp -Opcional
    if ($null -eq $word) { throw "No s'ha pogut iniciar Microsoft Word per llegir l'informe." }
    $out = New-Object System.Collections.ArrayList
    $started = $false
    try {
        $doc = $word.Documents.Open($docxPath, $false, $true)   # ReadOnly
        try {
            foreach ($p in $doc.Paragraphs) {
                $txt = ([string]$p.Range.Text).TrimEnd("`r","`n","`a"," ")
                if ([string]::IsNullOrWhiteSpace($txt)) { continue }
                $n = _NormalitzaText $txt

                # Fi: primera frase de conclusions.
                $isConcl = $false
                foreach ($pn in $phrN) { if ($pn -and $n.StartsWith($pn)) { $isConcl = $true; break } }
                if ($isConcl) { break }

                $isCaps = ($txt -cmatch '\p{Lu}') -and -not ($txt -cmatch '\p{Ll}')
                $isNum  = $txt -match '^\s*\d+\.\s'
                $isUrl  = $txt -match '^\s*https?://'

                if (-not $started) {
                    # Comença despres de la intro ("...deficiencies... esmenar...")
                    if ($n -match 'defici' -and $n -match 'esmenar') { $started = $true; continue }
                    if ($isCaps -or $isNum) { $started = $true }  # o al primer titol/numero
                    else { continue }
                }

                # Subratllat -> subseccio.
                $isUnder = $false
                try { $u = $p.Range.Font.Underline; if ($u -ne 0 -and $u -ne 9999999) { $isUnder = $true } } catch { }

                if ($isCaps) {
                    [void]$out.Add('<div style="font-weight:bold;margin-top:12px">' + (_EscHtml $txt) + '</div>')
                } elseif ($isUrl) {
                    [void]$out.Add('<div><a href="' + (_EscHtml $txt) + '">' + (_EscHtml $txt) + '</a></div>')
                } elseif ($isUnder -and -not $isNum) {
                    [void]$out.Add('<div style="text-decoration:underline;margin-top:4px">' + (_EscHtml $txt) + '</div>')
                } else {
                    $mt = if ($isNum) { 'margin-top:6px' } else { 'margin-left:18px' }
                    [void]$out.Add('<div style="' + $mt + '">' + (_TextToHtml $txt) + '</div>')
                }
            }
        } finally { $doc.Close($false) }
    } finally {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    return ($out -join "`n")
}

# --- Embolcall del correu (plantilla email-textos.json) -----------------------
# Els textos del correu. NOMES delega: qui els llegeix (i qui peta si no hi son)
# es _LoadEmailTextos, a EmailTextos.ps1.
#
# Aqui hi havia un fallback propi que, si el JSON no s'hi trobava, tornava
# 'cos = {REQUERIMENTS}': un correu al titular sense capcalera, sense les
# instruccions de la seu i sense l'avis legal. Era la QUARTA copia dels textos i,
# a mes, codi mort: Motor.ps1 carrega EmailTextos.ps1 (linia 457) abans que
# aquest fitxer (465), o sigui que el Get-Command sempre encertava i aquella
# branca no s'executava mai.
function _CorreuTextos {
    return (_LoadEmailTextos)
}
function _FillVars([string]$s, $h) {
    $g = { param($k) if ($h -and $h.ContainsKey($k)) { [string]$h[$k] } else { '' } }
    return (_OmpleVariables $s ([ordered]@{
        '{ID_GIA}'    = (& $g 'ID_GIA')
        '{ADRECA}'    = (& $g 'ADRECA')
        '{ACTIVITAT}' = (& $g 'ACTIVITAT')
        '{TITULAR}'   = (& $g 'TITULAR')
        '{DATA}'      = (Get-Date).ToString('dd/MM/yyyy')
    }))
}
function _BuildCorreu($requerimentsHtml, $header) {
    $tx = _CorreuTextos
    $subject = _FillVars ([string]$tx.assumpte) $header
    $cos = _FillVars ([string]$tx.cos) $header
    $parts = $cos -split '\{REQUERIMENTS\}', 2
    $html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;line-height:1.4">'
    $html += _TextToHtml $parts[0]
    $html += $requerimentsHtml
    if ($parts.Count -gt 1) { $html += _TextToHtml $parts[1] }
    $html += '</div>'
    return [pscustomobject]@{ Subject = $subject; Html = $html }
}

# Colors dels botons d'accio del dialeg: blau mari per ENVIAR i vermell per NO
# ENVIAR. Son dues accions oposades i irreversibles (un correu no es pot
# desenviar): el color ha de dir quina es quina d'un cop d'ull.
$Script:CorreuBlauMari      = [System.Drawing.Color]::FromArgb(16, 42, 87)
$Script:CorreuBlauMariHover = [System.Drawing.Color]::FromArgb(28, 62, 120)
$Script:CorreuVermell       = [System.Drawing.Color]::FromArgb(176, 0, 32)
$Script:CorreuVermellHover  = [System.Drawing.Color]::FromArgb(208, 26, 58)

# --- Enviament EmailJS --------------------------------------------------------
# Adreces disponibles per a Copia Oculta (CCO). Surten de la clau 'bcc' de
# docs\dades\email-textos.json, que es el mateix fitxer que porta l'assumpte i
# el cos. Abans les quatre adreces estaven escrites aqui I a docs\app.js (i la
# primera, una tercera vegada a Recordatoris.ps1): quatre adreces nominals de
# treballadors, repetides a ma, en un repositori public.
function _CorreuBccOpcions {
    # Array PLA: el crider hi posa @() (vegeu _EmailBccDeJson).
    try { return @((_LoadEmailTextos)['bcc']) } catch { return @() }
}

function Send-EmailJs($cfg, $toEmail, $bcc, $subject, $htmlMessage) {
    $payload = @{
        service_id  = $cfg.ServiceId
        template_id = $cfg.TemplateId
        user_id     = $cfg.PublicKey
        accessToken = $cfg.PrivateKey
        template_params = @{ to_email = $toEmail; bcc = $bcc; subject = $subject; message = $htmlMessage; name = $cfg.FromName }
    }
    $json  = $payload | ConvertTo-Json -Depth 6
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    Invoke-RestMethod -Method Post -Uri 'https://api.emailjs.com/api/v1.0/email/send' -ContentType 'application/json' -Body $bytes | Out-Null
    # Un correu que ha SORTIT compta per a la quota mensual d'EmailJS (200 al
    # pla gratuit). Es compta aqui, i no a cada eina, perque aixi hi entren
    # TOTS els enviaments del PC: aquesta eina i els recordatoris.
    _QuotaApunta 1
}

# Compon el missatge d'error d'un enviament fallit (PURA, testejable). EmailJS
# torna el MOTIU real al cos de la resposta ("API calls are disabled for
# non-browser applications", "The Public Key is invalid"...); sense això
# l'usuari només veu el "(403) Prohibido" genèric de .NET, que no diu res.
# El 403 típic d'aquest programa: EmailJS rebutja la crida perquè NO ve d'un
# navegador (el PC envia des de PowerShell; el mòbil, des del navegador, sí que
# passa). Es resol al panell d'EmailJS, no al codi.
function _EmailJsErrorText([int]$status, [string]$body, [string]$fallback) {
    $b = ([string]$body).Trim()
    $lines = New-Object System.Collections.ArrayList
    if ($status) { [void]$lines.Add("No s'ha pogut enviar (EmailJS, HTTP $status).") }
    elseif ($fallback) { [void]$lines.Add("No s'ha pogut enviar: " + [string]$fallback) }
    else { [void]$lines.Add("No s'ha pogut enviar el correu.") }
    if ($b) { [void]$lines.Add("Resposta del servei: $b") }
    if ($status -eq 403) {
        [void]$lines.Add('')
        [void]$lines.Add("El 403 (Prohibit) vol dir que EmailJS rebutja la crida des del PC. Comprova, al teu compte d'EmailJS (https://dashboard.emailjs.com):")
        [void]$lines.Add(" 1) Account -> Security: activa 'Allow EmailJS API for non-browser applications'. Aquesta eina envia des del PC (PowerShell), no des del navegador, i per defecte EmailJS ho bloqueja.")
        [void]$lines.Add(" 2) Que la Private key desada a local\emailjs.json sigui la correcta (Account -> General -> Private Key).")
        [void]$lines.Add("El mòbil segueix enviant perquè ho fa des del navegador; el PC necessita aquest permís.")
    }
    return ($lines -join "`n")
}

# Extreu l'estat HTTP i el cos de la resposta d'un error d'Invoke-RestMethod i
# en compon el missatge amb _EmailJsErrorText. (Toca .NET: no és pura.)
function _EmailJsRespError($err) {
    $status = 0
    $body = ''
    # A partir de PS 5.1, el cos de la resposta d'error sol venir a ErrorDetails.
    try { if ($err.ErrorDetails -and $err.ErrorDetails.Message) { $body = [string]$err.ErrorDetails.Message } } catch { }
    $resp = $null
    try { $resp = $err.Exception.Response } catch { }
    if ($resp) {
        try { $status = [int]$resp.StatusCode } catch { }
        if ([string]::IsNullOrEmpty($body)) {
            try {
                $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
                $body = $reader.ReadToEnd(); $reader.Close()
            } catch { }
        }
    }
    if (-not $status) {
        $m = [regex]::Match([string]$err.Exception.Message, '\((\d{3})\)')
        if ($m.Success) { $status = [int]$m.Groups[1].Value }
    }
    return (_EmailJsErrorText $status $body ([string]$err.Exception.Message))
}

# --- Destinatari per defecte (Rao social + Rep. legal de l'Excel) ------------
# Combina les dues adreces de correu del titular. Dedupe SENSE distingir
# majuscules i avisa (Duplicat) si son la mateixa: llavors nomes surt un cop.
# PURA i testejable.
function _CorreuDestinatarisPerDefecte([string]$raoEmail, [string]$repEmail) {
    $llista = New-Object System.Collections.ArrayList
    $vistes = @{}
    $dup = $false
    foreach ($e in @(([string]$raoEmail).Trim(), ([string]$repEmail).Trim())) {
        if ([string]::IsNullOrWhiteSpace($e)) { continue }
        $k = $e.ToLowerInvariant()
        if ($vistes.ContainsKey($k)) { $dup = $true; continue }
        $vistes[$k] = $true
        [void]$llista.Add($e)
    }
    return @{ Text = ($llista -join '; '); Duplicat = $dup; Compte = $llista.Count }
}

# --- DE QUINA ACTIVITAT és aquest informe? -----------------------------------
# El GIA ha de sortir del DOCUMENT QUE S'ENVIA, mai de l'últim informe generat:
# un informe de Seguiment no passa per l'assistent, i Load-LastReport encara
# tenia la capçalera d'un "Requeriment - Nou" anterior. Resultat real: es va
# obrir el correu del GIA 1466 amb l'assumpte i els destinataris del GIA 1000.

# Treu l'ID GIA del NOM del fitxer ("2026-09-08_Req2_GIA 1466.docx" -> "1466").
# Els noms els fa _GetOutputFileName ("..._GIA <id>.docx") i el Seguiment els
# conserva. El sufix d'unicitat ("_2", "_3"...) NO forma part de l'id: per això
# només s'agafen els dígits que van just darrere de "GIA". PURA.
function _GiaDelNomFitxer([string]$nom) {
    if ([string]::IsNullOrWhiteSpace($nom)) { return '' }
    $m = [regex]::Match([string]$nom, '(?i)GIA[\s_-]*([0-9]+)')
    if (-not $m.Success) { return '' }
    return $m.Groups[1].Value
}

# Decideix quin ID GIA val, comparant el del nom del fitxer i el de la
# capçalera del document. PURA i testejable: és tota la regla de decisió.
#   · cap dels dos           -> preguntar (informes antics sense ID GIA)
#   · només un               -> aquell
#   · tots dos i coincideixen-> aquell
#   · tots dos i difereixen  -> preguntar (no ho tenim clar; mai endevinar)
function _CorreuGiaDecideix([string]$giaNom, [string]$giaDoc) {
    $n = ([string]$giaNom).Trim()
    $d = ([string]$giaDoc).Trim()
    if ($n -eq '' -and $d -eq '') {
        return @{ Gia = ''; CalPreguntar = $true; Motiu = "no s'ha trobat cap ID GIA ni al nom del fitxer ni a la capçalera." }
    }
    if ($n -eq '') { return @{ Gia = $d; CalPreguntar = $false; Motiu = 'de la capçalera del document' } }
    if ($d -eq '') { return @{ Gia = $n; CalPreguntar = $false; Motiu = 'del nom del fitxer' } }
    if ($n -eq $d) { return @{ Gia = $d; CalPreguntar = $false; Motiu = 'del nom del fitxer i de la capçalera' } }
    return @{ Gia = $d; CalPreguntar = $true
              Motiu = "el nom del fitxer diu GIA $n i la capçalera del document diu GIA $d." }
}

# Llegeix l'ID GIA del document (nom + capçalera). La capçalera es llegeix del
# .docx com a ZIP (_ReadDocxParagraphs), SENSE obrir el Word: així es pot
# preguntar abans de posar-se a llegir el cos, que sí que costa.
function _CorreuGiaDelDocx($docxPath) {
    $giaNom = ''
    try { $giaNom = _GiaDelNomFitxer ([System.IO.Path]::GetFileNameWithoutExtension([string]$docxPath)) } catch { }
    $giaDoc = ''
    try {
        $lines = @(_ReadDocxParagraphs $docxPath)
        $giaDoc = _ExtractIdGia $lines
    } catch { $giaDoc = '' }
    return (_CorreuGiaDecideix $giaNom $giaDoc)
}

# Demana l'ID GIA quan no es pot deduir amb prou seguretat. Torna '' si es
# cancel·la (i llavors no s'envia res: val més no enviar que enviar a qui no és).
function _DemanaIdGia($docxPath, $info) {
    $form = _NewForm
    $form.Text = 'Enviar correu - de quina activitat és?'
    $form.FormBorderStyle = 'FixedDialog'
    $form.ClientSize = New-Object System.Drawing.Size(560, 230)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16, 14)
    $lbl.Size = New-Object System.Drawing.Size(528, 96)
    $lbl.Text = ("No es pot saber del cert de quina activitat és aquest informe:`n" +
                 [System.IO.Path]::GetFileName([string]$docxPath) + "`n`n" +
                 [string]$info.Motiu + "`n`n" +
                 "Escriu l'ID GIA de l'activitat perquè el correu vagi al titular correcte.")
    $form.Controls.Add($lbl)

    $lblG = New-Object System.Windows.Forms.Label
    $lblG.Text = 'ID GIA:'
    $lblG.Location = New-Object System.Drawing.Point(16, 120)
    $lblG.Size = New-Object System.Drawing.Size(60, 22)
    $form.Controls.Add($lblG)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(80, 118)
    $tb.Size = New-Object System.Drawing.Size(140, 24)
    $tb.Text = [string]$info.Gia
    $form.Controls.Add($tb)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Continuar'; $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(344, 180)
    $ok.Size = New-Object System.Drawing.Size(100, 32)
    _StyleAccentButton $ok $Script:CorreuBlauMari $Script:CorreuBlauMariHover
    $form.AcceptButton = $ok; $form.Controls.Add($ok)

    $no = New-Object System.Windows.Forms.Button
    $no.Text = 'Cancel·lar'; $no.DialogResult = 'Cancel'
    $no.Location = New-Object System.Drawing.Point(452, 180)
    $no.Size = New-Object System.Drawing.Size(92, 32)
    _StyleAccentButton $no $Script:CorreuVermell $Script:CorreuVermellHover
    $form.CancelButton = $no; $form.Controls.Add($no)

    if ($form.ShowDialog() -ne 'OK') { return '' }
    return ([string]$tb.Text).Trim()
}

# Fitxa de l'activitat a l'Excel per ID GIA (UNA sola obertura de l'Excel).
function _CorreuActivitatPerGia([string]$gia) {
    if ([string]::IsNullOrWhiteSpace($gia)) { return $null }
    try {
        $xls = Find-LatestActivitatsExcel
        if ($null -eq $xls) { return $null }
        $cache = Initialize-ActivitatsCache $xls.File
        return (Get-ActivitatFromCache $cache $gia)
    } catch { return $null }
}

# Munta la capçalera del correu per a un GIA concret. PURA.
#  · l'ID GIA mana: és el del document que s'envia;
#  · les dades surten de l'EXCEL (titular, adreça, activitat...);
#  · la capçalera de l'últim informe només s'aprofita si és del MATEIX GIA
#    (porta coses que l'Excel no té, com el núm. d'anotació), i mai per
#    trepitjar el que ja ha dit l'Excel.
function _CorreuHeaderMerge([string]$gia, $act, $repHeader) {
    $h = @{ ID_GIA = [string]$gia }
    $camps = @('TITULAR', 'ADRECA', 'ACTIVITAT', 'EXP_NUM', 'EMAIL', 'EMAIL_REP',
               'NUM_ANOTACIO', 'DATA_ANOTACIO', 'CLASSIFICACIO')
    if ($null -ne $act) {
        foreach ($k in $camps) {
            try { if ($act.ContainsKey($k)) { $h[$k] = [string]$act[$k] } } catch { }
        }
    }
    # L'últim informe NOMÉS si parla de la mateixa activitat.
    $repGia = ''
    if ($null -ne $repHeader) {
        try { $repGia = [string]$repHeader['ID_GIA'] } catch { }
    }
    if ($repGia -ne '' -and $repGia -eq [string]$gia) {
        foreach ($k in $camps) {
            $actual = ''
            if ($h.ContainsKey($k)) { $actual = [string]$h[$k] }
            if (-not [string]::IsNullOrWhiteSpace($actual)) { continue }
            try { if ($repHeader.ContainsKey($k)) { $h[$k] = [string]$repHeader[$k] } } catch { }
        }
    }
    return $h
}

# --- Localitzar el .docx mes recent ------------------------------------------
function _LatestDocx {
    $dir = if ($OutputDir) { [string]$OutputDir } else { Join-Path (_CorreuRepoRoot) 'Informes generats' }
    if (Get-Command _ResolveOutputDir -ErrorAction SilentlyContinue) {
        try { $dir = _ResolveOutputDir } catch { }
    }
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    $f = Get-ChildItem -LiteralPath $dir -Filter '*.docx' -File -ErrorAction SilentlyContinue |
         Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    return $null
}

# --- Diàleg d'enviament -------------------------------------------------------
# Diàleg únic: fusiona el missatge "Informe generat" amb l'enviament del correu.
# Torna @{ To = @(...); Bcc = @(...) } o $null si es cancel·la.
function _DialegEnviar($build, $destinatariDefault, $docxPath) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Informe generat - Enviar correu'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.Size = New-Object System.Drawing.Size(600, 420)

    $lblInf = New-Object System.Windows.Forms.Label
    $lblInf.Text = "Informe generat:`n$docxPath"
    $lblInf.Location = New-Object System.Drawing.Point(15, 12)
    $lblInf.Size = New-Object System.Drawing.Size(560, 42)
    $form.Controls.Add($lblInf)

    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = "Assumpte: $($build.Subject)"
    $lblA.Location = New-Object System.Drawing.Point(15, 58)
    $lblA.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblA)

    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Text = "Destinataris (separa'ls amb ; si n'hi ha més d'un):"
    $lblD.Location = New-Object System.Drawing.Point(15, 88)
    $lblD.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblD)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(15, 110)
    $tb.Size = New-Object System.Drawing.Size(560, 24)
    $tb.Text = [string]$destinatariDefault
    $form.Controls.Add($tb)

    $lblB = New-Object System.Windows.Forms.Label
    $lblB.Text = 'Còpia oculta (CCO):'
    $lblB.Location = New-Object System.Drawing.Point(15, 146)
    $lblB.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($lblB)

    $checks = @()
    $y = 170
    foreach ($opt in (_CorreuBccOpcions)) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $opt.Addr
        $cb.Checked = [bool]$opt.Default
        $cb.Location = New-Object System.Drawing.Point(25, $y)
        $cb.Size = New-Object System.Drawing.Size(540, 22)
        $form.Controls.Add($cb)
        $checks += $cb
        $y += 26
    }

    # Blau mari = enviar, vermell = no enviar. Un correu no es pot desenviar:
    # les dues accions han de ser distingibles d'un cop d'ull.
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Enviar'; $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(390, 335); $ok.Size = New-Object System.Drawing.Size(90, 32)
    _StyleAccentButton $ok $Script:CorreuBlauMari $Script:CorreuBlauMariHover
    $form.AcceptButton = $ok; $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'No enviar'; $cancel.DialogResult = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(485, 335); $cancel.Size = New-Object System.Drawing.Size(90, 32)
    _StyleAccentButton $cancel $Script:CorreuVermell $Script:CorreuVermellHover
    $form.CancelButton = $cancel; $form.Controls.Add($cancel)

    # Scroll vertical i ajust a la pantalla (vegeu suport/UiFinestra.ps1).
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })
    if ($form.ShowDialog() -ne 'OK') { return $null }

    $tos = @($tb.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $bccs = @()
    foreach ($cb in $checks) { if ($cb.Checked) { $bccs += $cb.Text } }
    return @{ To = $tos; Bcc = $bccs }
}

# Envia el correu per a un .docx concret. Reutilitzable (eina i final de generacio).
function Send-CorreuPerDocx($docxPath) {
    $cfg = _CorreuConfig
    if (-not $cfg.PublicKey -or -not $cfg.ServiceId -or -not $cfg.TemplateId) {
        [System.Windows.Forms.MessageBox]::Show("Falten les claus d'EmailJS a docs\config.js.",'Enviar correu','OK','Warning') | Out-Null
        return
    }
    if (-not $cfg.PrivateKey) {
        [System.Windows.Forms.MessageBox]::Show(
            "Falta la Private key d'EmailJS. Crea el fitxer:`n$($cfg.PrivatePath)`n`namb el contingut:`n{ ""private_key"": ""GOCSPX...la teva private key..."" }",
            'Enviar correu', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $docxPath -or -not (Test-Path -LiteralPath $docxPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat cap informe (.docx) per enviar.",'Enviar correu','OK','Information') | Out-Null
        return
    }

    # De QUINA ACTIVITAT es aquest informe? Surt del document que s'envia (nom
    # del fitxer + capcalera), NO de l'ultim informe generat: un Seguiment no
    # passa per l'assistent i Load-LastReport encara duia una altra activitat.
    # Es mira ABANS de llegir el cos (que obre el Word i costa): si s'ha de
    # preguntar o es cancel.la, no s'ha fet feina de franc.
    $giaInfo = _CorreuGiaDelDocx $docxPath
    $gia = [string]$giaInfo.Gia
    if ($giaInfo.CalPreguntar) { $gia = _DemanaIdGia $docxPath $giaInfo }
    if ([string]::IsNullOrWhiteSpace($gia)) { return }

    # Les dades del correu surten de l'Excel per aquest GIA. La capcalera de
    # l'ultim informe nomes s'aprofita si parla de la MATEIXA activitat.
    $act = _CorreuActivitatPerGia $gia
    $repHeader = $null
    $rep = if (Get-Command Load-LastReport -ErrorAction SilentlyContinue) { Load-LastReport } else { $null }
    if ($rep -and $rep.Header) {
        $repHeader = @{}
        foreach ($p in $rep.Header.PSObject.Properties) { $repHeader[$p.Name] = $p.Value }
    }
    $header = _CorreuHeaderMerge $gia $act $repHeader

    if ($null -eq $act) {
        [System.Windows.Forms.MessageBox]::Show(
            ("L'ID GIA $gia no s'ha trobat a l'Excel d'activitats.`n`n" +
             "El correu s'enviara igual, pero hauras d'escriure el destinatari a ma."),
            'Enviar correu', 'OK', 'Warning') | Out-Null
    }

    try {
        $reqHtml = _DocxRequerimentsHtml $docxPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error llegint l'informe:`n$($_.Exception.Message)",'Enviar correu','OK','Error') | Out-Null
        return
    }
    $build = _BuildCorreu $reqHtml $header

    # Destinatari per defecte: Rao social + Rep. legal (columnes de l'Excel).
    # Si son la mateixa adreca, nomes s'hi posa un cop i s'avisa.
    $def = _CorreuDestinatarisPerDefecte ([string]$header['EMAIL']) ([string]$header['EMAIL_REP'])
    $destinatariDefault = [string]$def.Text
    if ($def.Duplicat) {
        [System.Windows.Forms.MessageBox]::Show(
            "L'adreca de Rao social i la del Representant legal son la mateixa; s'ha posat una sola vegada.",
            'Enviar correu', 'OK', 'Information') | Out-Null
    }

    $res = _DialegEnviar $build $destinatariDefault $docxPath
    if ($null -eq $res) { return }
    if (-not $res.To -or @($res.To).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Indica almenys un destinatari.','Enviar correu','OK','Warning') | Out-Null
        return
    }
    $toStr  = ($res.To -join ',')
    $bccStr = ($res.Bcc -join ',')
    try {
        Send-EmailJs $cfg $toStr $bccStr $build.Subject $build.Html
        $resum = "Correu enviat a: $toStr"
        if ($bccStr) { $resum += "`nCCO: $bccStr" }
        [System.Windows.Forms.MessageBox]::Show($resum,'Enviar correu','OK','Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show((_EmailJsRespError $_),'Enviar correu','OK','Error') | Out-Null
    }
}

# Eina de menu: agafa el .docx mes recent (l'ultim generat, potser editat a ma).
function Invoke-EnviarCorreu {
    $docx = _LatestDocx
    Send-CorreuPerDocx $docx
}

# Ofereix enviar el correu al final de generar un informe (normal o seguiment).
# El diàleg d'enviament és la confirmació (ja fusionat amb "Informe generat"):
# no es fa cap pregunta Sí/No prèvia. Si l'usuari no vol enviar, tanca el diàleg.
function Offer-EnviarCorreu($docxPath) {
    if ([string]::IsNullOrWhiteSpace([string]$docxPath) -or -not (Test-Path -LiteralPath $docxPath)) {
        $docxPath = _LatestDocx
    }
    Send-CorreuPerDocx $docxPath
}
