#requires -Version 5.1
<#
.SYNOPSIS
  Per on surten els correus del PC: EmailJS (com sempre) o l'Outlook de
  l'ordinador (octubre 2026).

.DESCRIPTION
  L'usuari: "es poden tenir les dues opcions? Poder triar entre EmailJS per no
  desfer-ho i Outlook per fer proves". Les dues conviuen, i es tria a
  Configuracio o a la finestra d'"Enviar correu" (la mateixa preferencia, a
  settings.json: 'CorreuVia'). Tres vies:

    emailjs             com sempre (la plantilla d'EmailJS, la quota de 200/mes)
    outlook-esborrany   l'Outlook de l'ordinador, i el correu es queda a
                        Esborranys: no s'envia res, es per veure'l
    outlook             l'Outlook de l'ordinador, i s'envia

  Per que l'Outlook: surt de la bustia de l'usuari (a Elements enviats, les
  respostes li tornen), sense quota ni Private key. Compte: l'enviament
  directe pot fer saltar l'avis de seguretat de l'Outlook ("un programa intenta
  enviar correu") o el pot bloquejar la politica d'informatica; desar
  esborranys no el fa saltar (per aixo Controls periodics nomes en desa, i hi
  ha guard).

  ELS RECORDATORIS, TAMBE ELS AUTOMATICS, segueixen la via triada (l'usuari,
  octubre 2026: "vull que Recordatoris tambe funcioni amb Outlook, que salti
  un avis al programa que els correus estan a Esborranys"). Els que la tasca
  del Windows deixa a Esborranys s'apunten com a PENDENTS
  (Add-CorreuEsborranysPendents) i el menu ho avisa (Show-AvisEsborranysSiCal).

  EL REMITENT (l'usuari: "a la feina puc enviar correus des d'adreces
  diferents; com ho puc seleccionar per no haver de canviar-la a tots els
  correus?"): una adreca per defecte a Configuracio ('CorreuRemitent'). Si es
  un compte de l'Outlook, SendUsingAccount; si no (una bustia compartida),
  SentOnBehalfOfName. Nomes l'Outlook: amb EmailJS el remitent es el del
  servei configurat a EmailJS.

  Aqui hi ha TOT el que fa sortir un correu: les claus d'EmailJS i
  Send-EmailJs (vivien a EnviarCorreu.ps1; amb la tria, EnviarCorreu i
  CorreuVia haurien depes l'un de l'altre) i l'Outlook, que s'obre NOMES aqui
  (New-OutlookApp), com el Word a Motor.ps1 i l'Excel a Excel.ps1: hi ha guard
  (06-guards.ps1).

  NOMES DEFINEIX FUNCIONS.
#>

# ============================================================================
# EMAILJS: les claus i l'enviament (vivien a EnviarCorreu.ps1)
# ============================================================================
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

# ES POT ENVIAR CORREU? Torna '' si si, i si no el MOTIU, per ensenyar-lo tal
# qual. Viu aqui, al costat de qui llegeix les claus, i no a cada eina: ho
# necessiten l'enviament dels recordatoris (per aturar una tanda) i
# l'interruptor automatic de la seva rajola (per no deixar ences un automatic
# que no pot fer res). Escrit dues vegades, el dia que es canvies una clau
# n'hi hauria una que no se n'assabentaria.
function Test-CorreuLlest {
    $c = _CorreuConfig
    if (-not $c.PublicKey -or -not $c.ServiceId -or -not $c.TemplateId) {
        return "falten les claus d'EmailJS a docs\config.js"
    }
    if (-not $c.PrivateKey) { return "falta la Private key d'EmailJS a $($c.PrivatePath)" }
    return ''
}

# --- Enviament EmailJS --------------------------------------------------------
function Send-EmailJs($cfg, $toEmail, $bcc, $subject, $htmlMessage) {
    # TLS 1.2: el Windows PowerShell 5.1 no el fa servir per defecte i EmailJS
    # no n'accepta d'altre. Abans es posava en carregar EnviarCorreu.ps1.
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
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

# ============================================================================
# LA VIA: EmailJS o l'Outlook
# ============================================================================
$Script:CorreuVies = [ordered]@{
    'emailjs'           = 'EmailJS (com sempre)'
    'outlook-esborrany' = ("Outlook: el deixa a Esborranys (no s'envia)")
    'outlook'           = 'Outlook: envia el correu'
}
$Script:CorreuViaDefecte = 'emailjs'

# Una via valida, o la per defecte (un settings.json escrit a ma o d'una
# versio que en tingues una altra no pot deixar el PC sense correu). PURA.
function _CorreuViaValida($v) {
    $s = ([string]$v).Trim().ToLowerInvariant()
    if ($Script:CorreuVies.Contains($s)) { return $s }
    return $Script:CorreuViaDefecte
}

function _CorreuViaEsOutlook([string]$via) { return ($via -like 'outlook*') }

# La via triada en AQUEST PC.
function Get-CorreuVia {
    $s = Load-AppSettings
    return (_CorreuViaValida (_PropInf $s 'CorreuVia'))
}

# El settings.json amb la via nova, i la resta tal com era. La per defecte no
# s'hi escriu (mateix criteri que _BuildSettingsOverrides: nomes el que
# difereix). PURA.
function _SettingsAmbCorreuVia($settings, [string]$via) {
    $v = _CorreuViaValida $via
    return (_SettingsAmbClau $settings 'CorreuVia' $(if ($v -eq $Script:CorreuViaDefecte) { '' } else { $v }))
}

# Una clau de settings.json canviada i la resta tal com era; buida, s'esborra
# (mateix criteri que _BuildSettingsOverrides: nomes el que difereix). PURA.
function _SettingsAmbClau($settings, [string]$clau, [string]$valor) {
    $h = ConvertTo-Mapa $settings
    if ($null -eq $h) { $h = @{} }
    $v = ([string]$valor).Trim()
    if ($v -eq '') { [void]$h.Remove($clau) } else { $h[$clau] = $v }
    return $h
}

# Desa la via SENSE TOCAR la resta de settings.json (les carpetes, els
# automatismes). Torna $true si ha pogut.
function Set-CorreuVia([string]$via) {
    return (Save-AppSettings (_SettingsAmbCorreuVia (Load-AppSettings) $via))
}

# EL REMITENT per defecte de l'Outlook ('' = el compte per defecte de l'Outlook).
function Get-CorreuRemitent { return ([string](_PropInf (Load-AppSettings) 'CorreuRemitent')).Trim() }

# Una adreca de correu amb cara de ser-ho ('' tambe val: vol dir "la per
# defecte"). PURA.
function _CorreuRemitentValid([string]$r) {
    $r = ([string]$r).Trim()
    return ($r -eq '' -or $r -match '^[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+$')
}

# Com es diu, al costat de la via (dialeg d'enviar, confirmacions). PURA.
function _CorreuRemitentText([string]$via, [string]$remitent) {
    if (-not (_CorreuViaEsOutlook $via)) { return '' }
    $r = ([string]$remitent).Trim()
    if ($r -eq '') { return "des del compte per defecte de l'Outlook" }
    return ('des de ' + $r)
}

# El desplegable de la via, per a les dues pantalles que el porten
# (Configuracio i "Enviar correu"). Viu aqui i no a UiComuns.ps1 perque es
# d'aquest modul: UiComuns no sap res del motor. Torna el ComboBox; la via
# triada surt de _CorreuViaDelCombo.
function Add-CorreuViaCombo($parent, [int]$x, [int]$y, [int]$ample, [string]$via) {
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.DropDownStyle = 'DropDownList'
    $cb.Location = New-Object System.Drawing.Point($x, $y)
    $cb.Size = New-Object System.Drawing.Size($ample, 24)
    $claus = @($Script:CorreuVies.Keys)
    foreach ($k in $claus) { [void]$cb.Items.Add([string]$Script:CorreuVies[$k]) }
    $cb.SelectedIndex = [math]::Max(0, [array]::IndexOf($claus, (_CorreuViaValida $via)))
    $cb.Tag = $claus
    [void]$parent.Controls.Add($cb)
    return $cb
}

function _CorreuViaDelCombo($cb) {
    $claus = @($cb.Tag)
    if ($cb.SelectedIndex -lt 0 -or $cb.SelectedIndex -ge $claus.Count) { return $Script:CorreuViaDefecte }
    return [string]$claus[$cb.SelectedIndex]
}

# L'Outlook de l'ordinador, en UN sol lloc. $null si no hi es: l'Outlook
# classic (el "nou Outlook" de Windows no es pot controlar per COM).
function New-OutlookApp {
    try { return (New-Object -ComObject Outlook.Application) } catch { return $null }
}

# Els comptes de l'Outlook d'aquest PC (les adreces), per triar el remitent a
# Configuracio. Una bustia compartida NO hi surt (no es un compte): s'escriu a
# ma i va per SentOnBehalfOfName.
function Get-OutlookComptes {
    $ol = New-OutlookApp
    if ($null -eq $ol) { throw "No s'ha pogut obrir l'Outlook d'aquest ordinador (cal l'Outlook cl" + [char]0x00E0 + "ssic)." }
    $out = New-Object System.Collections.ArrayList
    try {
        foreach ($a in $ol.Session.Accounts) {
            $adr = ''
            try { $adr = ([string]$a.SmtpAddress).Trim() } catch { }
            if ($adr -and -not ($out -contains $adr)) { [void]$out.Add($adr) }
        }
    } finally {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) } catch { }
    }
    return ,@($out)
}

# El compte de l'Outlook que te aquesta adreca, o $null (llavors es una
# bustia compartida o una adreca en nom de la qual s'envia).
function _OutlookTriaCompte($outlook, [string]$remitent) {
    $r = ([string]$remitent).Trim()
    if ($r -eq '' -or $null -eq $outlook) { return $null }
    try {
        foreach ($a in $outlook.Session.Accounts) {
            if (([string]$a.SmtpAddress).Trim() -ieq $r) { return $a }
        }
    } catch { }
    return $null
}

# Posa el remitent a un correu de l'Outlook. Assignar SendUsingAccount
# directament falla en algunes versions del PowerShell (un objecte COM dins
# d'una propietat COM); llavors, per InvokeMember.
function _OutlookPosaRemitent($m, $compte, [string]$remitent) {
    if ($null -ne $compte) {
        try { $m.SendUsingAccount = $compte }
        catch { [void]$m.GetType().InvokeMember('SendUsingAccount', [System.Reflection.BindingFlags]::SetProperty, $null, $m, @($compte)) }
        return
    }
    $r = ([string]$remitent).Trim()
    if ($r -ne '') { $m.SentOnBehalfOfName = $r }
}

# L'Outlook vol les adreces separades per ';' (EmailJS les accepta amb ',').
# PURA.
function _OutlookAdreces([string]$s) {
    return ((@(([string]$s) -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) -join '; ')
}

# ES POT ENVIAR PER AQUESTA VIA? '' si si, o el motiu. EmailJS ho mira a les
# claus (Test-CorreuLlest); l'Outlook, en obrir la sessio (cal crear-lo).
function Test-CorreuViaLlest([string]$via) {
    if (_CorreuViaEsOutlook $via) { return '' }
    return (Test-CorreuLlest)
}

# Una SESSIO d'enviament: el que cal per enviar N correus per una via (les
# claus d'EmailJS, o l'Outlook obert un sol cop per a tota una tanda). Llanca
# amb un missatge clar si l'Outlook no hi es.
function Open-CorreuSessio([string]$via) {
    $via = _CorreuViaValida $via
    $s = @{ Via = $via; Cfg = $null; Outlook = $null; JaObert = $false; Desats = 0; Enviats = 0; Remitent = ''; Compte = $null }
    if (_CorreuViaEsOutlook $via) {
        # Si l'Outlook no era obert, el que l'obrim som nosaltres i, en
        # deixar-lo anar, podria tancar-se amb els correus a la Safata de
        # sortida: al final se li demana que els envii (Close-CorreuSessio).
        try { $s.JaObert = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) } catch { }
        $s.Outlook = New-OutlookApp
        if ($null -eq $s.Outlook) {
            throw "No s'ha pogut obrir l'Outlook d'aquest ordinador. Cal l'Outlook clàssic (el «nou Outlook» de Windows no es pot fer servir des d'un programa). Torna a EmailJS a Configuració."
        }
        # El compte es busca UN cop per sessio, no a cada correu.
        $s.Remitent = Get-CorreuRemitent
        $s.Compte = _OutlookTriaCompte $s.Outlook $s.Remitent
    } else {
        $s.Cfg = _CorreuConfig
    }
    return $s
}

# Envia (o desa a Esborranys) UN correu. Llanca si falla: el cridador en fa el
# missatge amb _CorreuSessioError.
# $cc: la plantilla d'EmailJS no en te, i alla va amb el destinatari.
# $adjunts: fitxers que van adjunts (el Seguiment, octubre 2026). Nomes per
# l'Outlook: la plantilla d'EmailJS no en sap, i un correu que hauria de portar
# adjunts no pot sortir sense dir res.
function Send-CorreuSessio($s, [string]$to, [string]$bcc, [string]$subject, [string]$html, [string]$cc = '', $adjunts = @()) {
    $adj = @(@($adjunts) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if (-not (_CorreuViaEsOutlook $s.Via)) {
        if ($adj.Count -gt 0) { throw "EmailJS no pot enviar adjunts: tria l'Outlook per a aquesta eina a Configuraci" + [char]0x00F3 + '.' }
        if (-not [string]::IsNullOrWhiteSpace($cc)) { $to = (@($to, $cc) | Where-Object { $_ }) -join ',' }
        # Send-EmailJs ja apunta la quota.
        Send-EmailJs $s.Cfg $to $bcc $subject $html
        $s.Enviats++
        return
    }
    $m = $null
    try {
        $m = $s.Outlook.CreateItem(0)   # olMailItem
        _OutlookPosaRemitent $m $s.Compte $s.Remitent
        $m.To = (_OutlookAdreces $to)
        if (-not [string]::IsNullOrWhiteSpace($cc)) { $m.CC = (_OutlookAdreces $cc) }
        if (-not [string]::IsNullOrWhiteSpace($bcc)) { $m.BCC = (_OutlookAdreces $bcc) }
        $m.Subject = $subject
        $m.HTMLBody = $html
        foreach ($f in $adj) {
            if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { throw "No trobo l'adjunt: $f" }
            [void]$m.Attachments.Add([string]$f)
        }
        if ($s.Via -eq 'outlook-esborrany') { $m.Save(); $s.Desats++ } else { $m.Send(); $s.Enviats++ }
    } finally {
        if ($null -ne $m) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($m) } catch { } }
    }
}

# Tanca la sessio. NO es fa Quit() de l'Outlook (podria tancar el de
# l'usuari, com ja diu Controls periodics); nomes s'allibera.
# Si l'Outlook l'hem obert nosaltres (la tasca automatica, amb l'Outlook
# tancat), s'espera fins a $esperaSeg que la Safata de sortida es buidi: si no,
# en deixar-lo anar es tancaria amb els correus a dins.
function Close-CorreuSessio($s, [int]$esperaSeg = 60) {
    if ($null -eq $s -or $null -eq $s.Outlook) { return }
    if ($s.Enviats -gt 0 -and -not $s.JaObert) {
        try { $s.Outlook.Session.SendAndReceive($false) } catch { }
        try {
            $sortida = $s.Outlook.Session.GetDefaultFolder(4)   # olFolderOutbox
            for ($i = 0; $i -lt $esperaSeg -and [int]$sortida.Items.Count -gt 0; $i++) { Start-Sleep -Seconds 1 }
        } catch { }
    }
    if ($null -ne $s.Compte) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s.Compte) } catch { } }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s.Outlook) } catch { }
    $s.Outlook = $null
}

# El missatge d'un enviament fallit, segons la via. Els d'EmailJS porten el
# motiu del servidor (_EmailJsRespError); els de l'Outlook, el de l'Outlook.
function _CorreuSessioError($s, $err) {
    if ($null -ne $s -and (_CorreuViaEsOutlook $s.Via)) {
        return ("L'Outlook no ha pogut " + $(if ($s.Via -eq 'outlook-esborrany') { 'desar' } else { 'enviar' }) +
                " el correu: " + [string]$err.Exception.Message +
                "`n`nSi ha sortit un avís de seguretat de l'Outlook i l'has rebutjat, o la política d'informàtica no ho deixa, torna a EmailJS a Configuració.")
    }
    return (_EmailJsRespError $err)
}

# Com queda dit en una frase, per a les confirmacions i els resums.
function _CorreuViaText([string]$via) { return [string]$Script:CorreuVies[(_CorreuViaValida $via)] }

# ============================================================================
# ELS ESBORRANYS PENDENTS: el que la tasca automatica deixa a l'Outlook
# ============================================================================
# Amb la via 'outlook-esborrany', els recordatoris automatics es queden a
# Esborranys i ningu no els veu. S'apunten aqui (%LOCALAPPDATA%: porten noms i
# adreces, MAI al repositori) i el menu ho avisa fins que l'usuari diu que ja
# els ha enviat.
function _CorreuEsborranysPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'correus-esborranys.json'))
}

# La llista amb una tanda nova al final. PURA.
function _EsborranysAmbNous($pendents, [string]$origen, $detalls, [datetime]$quan) {
    $l = New-Object System.Collections.ArrayList
    foreach ($p in @($pendents)) { if ($null -ne $p) { [void]$l.Add($p) } }
    $d = @(@($detalls) | Where-Object { $_ } | ForEach-Object { [string]$_ })
    if ($d.Count -gt 0) {
        [void]$l.Add([pscustomobject]@{ data = $quan.ToString('yyyy-MM-dd HH:mm'); origen = $origen; correus = $d })
    }
    return ,@($l)
}

# El text de l'avis. PURA.
function _EsborranysAvisText($pendents) {
    $ps = @(@($pendents) | Where-Object { $null -ne $_ })
    $n = 0
    foreach ($p in $ps) { $n += @($p.correus).Count }
    $linies = New-Object System.Collections.ArrayList
    [void]$linies.Add("Hi ha $n correus a la carpeta Esborranys de l'Outlook que encara no s'han enviat:")
    [void]$linies.Add('')
    foreach ($p in $ps) {
        [void]$linies.Add(' ' + [char]0x00B7 + ' ' + [string]$p.data + ' - ' + [string]$p.origen + ': ' + @($p.correus).Count)
        foreach ($c in @(@($p.correus) | Select-Object -First 5)) { [void]$linies.Add('      ' + [string]$c) }
        if (@($p.correus).Count -gt 5) { [void]$linies.Add('      ...') }
    }
    [void]$linies.Add('')
    [void]$linies.Add("Revisa'ls i envia'ls des de l'Outlook. Ja consten com a recordatoris fets: si n'esborres algun sense enviar-lo, aquell titular no el rebr" + [char]0x00E0 + " fins al proper per" + [char]0x00ED + "ode.")
    [void]$linies.Add('')
    [void]$linies.Add("Ja els has enviat? S" + [char]0x00ED + " = no m'ho tornis a dir. No = recorda-m'ho m" + [char]0x00E9 + "s tard.")
    return ($linies -join "`n")
}

function Get-CorreuEsborranysPendents {
    $j = Read-JsonFile (_CorreuEsborranysPath)
    if ($null -eq $j) { return ,@() }
    return ,@(@($j.pendents) | Where-Object { $null -ne $_ })
}

function Add-CorreuEsborranysPendents([string]$origen, $detalls) {
    $l = _EsborranysAmbNous (Get-CorreuEsborranysPendents) $origen $detalls (Get-Date)
    try { Write-JsonFile (_CorreuEsborranysPath) ([pscustomobject]@{ pendents = @($l) }) 6 } catch { }
}

function Clear-CorreuEsborranysPendents {
    try { Write-JsonFile (_CorreuEsborranysPath) ([pscustomobject]@{ pendents = @() }) 6 } catch { }
}

# L'AVIS. El criden el menu en obrir-se i el seu rellotge (cada minut: la
# tasca automatica pot deixar-ne mentre el programa es obert). "No" el calla
# fins que n'arribin de nous; "Si" buida la llista. Viu en una funcio i no a
# la closure del menu: dins d'una closure, $Script: no es la de l'script.
$Script:EsborranysAvisCallat = ''
function Show-AvisEsborranysSiCal {
    $ps = Get-CorreuEsborranysPendents
    if ($ps.Count -eq 0) { return }
    $marca = [string]$ps.Count + '|' + [string]$ps[$ps.Count - 1].data
    if ($marca -eq $Script:EsborranysAvisCallat) { return }
    $Script:EsborranysAvisCallat = $marca   # abans del MessageBox: el rellotge no el pot repetir
    $r = [System.Windows.Forms.MessageBox]::Show((_EsborranysAvisText $ps), "Correus a Esborranys de l'Outlook", 'YesNo', 'Warning')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { Clear-CorreuEsborranysPendents; $Script:EsborranysAvisCallat = '' }
}
