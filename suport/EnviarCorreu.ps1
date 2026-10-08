#requires -Version 5.1
<#
  EnviarCorreu.ps1 - Eina "Enviar correu" (secció MÒBIL).

  Envia el correu de requeriments des del PC, amb el MATEIX format i remitent que
  el mòbil (mateixa plantilla d'EmailJS). El cos de requeriments es construeix
  llegint el .docx GENERAT (estructura REQ1: seccions, subseccions, numeració i
  qualsevol edició manual), i s'embolcalla amb les condicions del correu (text
  introductori + text final, SENSE conclusions) definides a
  docs/dades/email-textos.json.

  Enviament: per EmailJS (API REST) o per l'Outlook de l'ordinador, segons el
  que es triï (CorreuVia.ps1, que també té les claus d'EmailJS). La Public key /
  Service ID / Template ID es llegeixen de docs/config.js (no secretes). La
  Private key es desa a la carpeta local/ del repositori (gitignored):
  local/emailjs.json -> { "private_key": "..." }.
#>

# _CorreuRepoRoot, les claus d'EmailJS (_CorreuConfig, Test-CorreuLlest) i
# l'enviament (Send-EmailJs, _EmailJsRespError) viuen a CorreuVia.ps1, amb
# l'Outlook: es "per on surten els correus", i ho fan servir aquesta eina i
# els Recordatoris (octubre 2026; si es quedaven aqui, els dos fitxers
# dependrien l'un de l'altre).

# --- Utils de text -> HTML ---------------------------------------------------
# _EscHtml i _TextToHtml viuen a CorreuFormat.ps1: les fa servir tambe el
# format del correu de requeriments, i CorreuFormat es la capa de sota (si
# fossin aqui, els dos fitxers dependrien l'un de l'altre).

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

# --- Llegir l'informe (.docx) SENSE Word -----------------------------------
# Abans s'obria el Word per llegir-lo paragraf a paragraf i es comencava al
# primer paragraf en MAJUSCULES: "ID GIA: 1398" ho es, i el correu repetia la
# capcalera sencera, "INFORME" i la nota de l'Ordenanca. Ara es llegeix el XML
# (_CorreuDocxLlegeix, CorreuFormat.ps1): mes rapid, sense Word, provat a Linux,
# i amb la sagnia, els espais i la negreta de debo de cada paragraf.
#
# FileShare.ReadWrite: l'informe pot estar obert al Word (l'usuari l'ha retocat
# abans d'enviar-lo) i un ZipFile.OpenRead a pel no el podria obrir.
function _CorreuDocumentXml([string]$docxPath) {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue | Out-Null
    $fs = New-Object System.IO.FileStream($docxPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $entry = $zip.GetEntry('word/document.xml')
            if ($null -eq $entry) { throw "El fitxer no sembla un .docx valid (falta word/document.xml)." }
            $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { return $sr.ReadToEnd() } finally { $sr.Close() }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
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
        '{EXP_NUM}'   = (& $g 'EXP_NUM')
        '{ADRECA}'    = (& $g 'ADRECA')
        '{ACTIVITAT}' = (& $g 'ACTIVITAT')
        '{TITULAR}'   = (& $g 'TITULAR')
        '{DATA}'      = (Get-Date).ToString('dd/MM/yyyy')
    }))
}

# Els valors de la capcalera del correu: els de l'INFORME que s'envia (el que
# diu el document) i, si alla no hi son, els de l'Excel. PURA.
function _CorreuValorsCapcalera($linies, $capDocx, $header) {
    $out = @()
    foreach ($l in @($linies)) {
        $k = _CorreuClauEtiqueta ([string]$l.Etiqueta)
        $v = ''
        if ($null -ne $capDocx -and $capDocx.Contains($k)) { $v = [string]$capDocx[$k] }
        if ([string]::IsNullOrWhiteSpace($v)) { $v = _CorreuOmplePlantilla ([string]$l.Plantilla) $header }
        $out += [pscustomobject]@{ Etiqueta = [string]$l.Etiqueta; Plantilla = [string]$l.Plantilla; Valor = [string]$v }
    }
    return $out
}

# EL CORREU SENCER a partir del document.xml de l'informe. $avui arriba de fora
# (dd/MM/yyyy) perque es pugui provar. Torna @{ Subject; Html; Intro }.
#   - Capcalera: les linies de '0 CAPCALERA' (ID GIA ... Objecte).
#   - {INTRO}: "documentacio aportada" o "visita", segons l'Objecte de l'informe.
#     Una visita sense data agafa la d'avui (tambe a l'Objecte del correu).
#   - {REQUERIMENTS}: el cos de l'informe, SENSE les conclusions.
function _BuildCorreu([string]$documentXml, $header, [string]$avui) {
    $tx  = _CorreuTextos
    $fmt = _CorreuFormat
    $doc = _CorreuDocxLlegeix $documentXml $fmt
    $capJson = Read-JsonFile (Get-CapcaleraJsonPath)
    $linies = @(_CorreuValorsCapcalera (_CorreuCapcaleraLinies $capJson) $doc.Capcalera $header)
    $lo = @($linies | Where-Object { $_.Plantilla -match '<<\s*ORIGEN\s*>>' }) | Select-Object -First 1
    $origen = _OrigenDesDeText $(if ($null -ne $lo) { $lo.Valor } else { '' })
    if ($origen.ORIGEN_TIPUS -eq 'insp' -and [string]::IsNullOrWhiteSpace($origen.DATA_INSPECCIO)) {
        $origen.DATA_INSPECCIO = $avui
        $lo.Valor = _BuildOrigenText $origen
    }
    $intro = _CorreuIntro $tx $origen $avui
    # Les variables del cos ({ADRECA}, {TITULAR}... del peu): les de l'Excel i,
    # si alla no hi son, les de la capcalera de l'informe. Si no, el peu deia
    # "feu-hi constar: ID GIA 1398, Adreca , Titular ." amb l'informe al davant.
    $vars = @{}
    if ($null -ne $header) { foreach ($k in @($header.Keys)) { $vars[[string]$k] = $header[$k] } }
    foreach ($l in $linies) {
        $mt = [regex]::Match([string]$l.Plantilla, '^<<\s*([A-Za-z0-9_]+)\s*>>$')
        if ($mt.Success -and [string]::IsNullOrWhiteSpace([string]$vars[$mt.Groups[1].Value])) { $vars[$mt.Groups[1].Value] = [string]$l.Valor }
    }
    $omple = { param($s) (_FillVars ([string]$s) $vars).Replace('{INTRO}', $intro) }
    $html = _CorreuCosAHtml ([string]$tx.cos) $omple (_CorreuCapcaleraHtml $linies $fmt (_CorreuCapcaleraTitol $capJson)) ([string]$doc.Html) $fmt
    return [pscustomobject]@{ Subject = (_FillVars ([string]$tx.assumpte) $vars); Html = $html; Intro = $intro }
}

# Colors dels botons d'accio del dialeg: blau mari per ENVIAR i vermell per NO
# ENVIAR. Son dues accions oposades i irreversibles (un correu no es pot
# desenviar): el color ha de dir quina es quina d'un cop d'ull.
#
# NOMES en interactiu, com $Script:BrandMaroon (UiComuns.ps1): en headless
# (Actualitzar.bat, les proves, RecordatorisAuto) System.Drawing NO esta
# carregat i [System.Drawing.Color] peta EN CARREGAR el fitxer, o sigui que
# s'endu el motor sencer. Va passar: "No se encuentra el tipo
# [System.Drawing.Color]" a les vistes en Word, les dades del mobil i el Drive.
$Script:CorreuBlauMari      = $null
$Script:CorreuBlauMariHover = $null
$Script:CorreuVermell       = $null
$Script:CorreuVermellHover  = $null
if (-not $Script:HeadlessTest) {
    $Script:CorreuBlauMari      = [System.Drawing.Color]::FromArgb(16, 42, 87)
    $Script:CorreuBlauMariHover = [System.Drawing.Color]::FromArgb(28, 62, 120)
    $Script:CorreuVermell       = [System.Drawing.Color]::FromArgb(176, 0, 32)
    $Script:CorreuVermellHover  = [System.Drawing.Color]::FromArgb(208, 26, 58)
}

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

# Destinatari BUIT = correu de prova per a un mateix. L'usuari esborrava el
# destinatari per rebre'l nomes ell en CCO i el dialeg ho aturava ("Indica
# almenys un destinatari"). No es pot enviar nomes en CCO: la plantilla
# d'EmailJS necessita un To_Email i, buit, el servei torna "The recipients
# address is empty". Per aixo l'adreca propia passa a ser el destinatari.
# L'adreca propia es la CCO marcada per defecte a email-textos.json: cap
# adreca escrita al codi (el repositori es public). Es treu de la CCO perque
# no arribi dues vegades. Sense cap CCO per defecte, To queda buit i el
# cridador ho atura com sempre. PURA.
function _CorreuDestinatariBuit($tos, $bccs, $opcions) {
    $tos  = @(@($tos) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $bccs = @(@($bccs) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($tos.Count -gt 0) { return @{ To = $tos; Bcc = $bccs; Prova = $false } }
    $jo = @(@($opcions) | Where-Object { $_.Default } | ForEach-Object { ([string]$_.Addr).Trim() } | Where-Object { $_ }) | Select-Object -First 1
    if (-not $jo) { return @{ To = @(); Bcc = $bccs; Prova = $false } }
    $resta = @($bccs | Where-Object { ([string]$_).Trim().ToLowerInvariant() -ne $jo.ToLowerInvariant() })
    return @{ To = @($jo); Bcc = $resta; Prova = $true }
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

    [void](_AddPeuBotons $form @(
        @{ Nom = 'No'; Text = 'Cancel·lar'; Resultat = 'Cancel'; Esc = $true; Estil = 'accent'; Fons = $Script:CorreuVermell; FonsHover = $Script:CorreuVermellHover }) @(
        @{ Nom = 'Ok'; Text = 'Continuar'; Resultat = 'OK'; Intro = $true; Estil = 'accent'; Fons = $Script:CorreuBlauMari; FonsHover = $Script:CorreuBlauMariHover }) 180)

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
    $lblD.Text = "Destinataris (separa'ls amb ;). Si ho deixes buit, t'arriba només a tu:"
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

    # PER ON SURT (CorreuVia.ps1): EmailJS o l'Outlook de l'ordinador. Es la
    # mateixa preferencia que a Configuracio, i el que triis aqui s'hi desa.
    $lblVia = New-Object System.Windows.Forms.Label
    $lblVia.Text = 'Enviar amb:'
    $lblVia.Location = New-Object System.Drawing.Point(15, ($y + 12))
    $lblVia.Size = New-Object System.Drawing.Size(90, 20)
    $form.Controls.Add($lblVia)
    $cbVia = Add-CorreuViaCombo $form 105 ($y + 9) 300 (Get-CorreuVia)
    # Des de quina adreca (Configuracio): que es vegi abans d'enviar.
    $remitent = Get-CorreuRemitent
    $lblDes = New-Object System.Windows.Forms.Label
    $lblDes.Location = New-Object System.Drawing.Point(412, ($y + 12))
    $lblDes.Size = New-Object System.Drawing.Size(160, 20)
    $lblDes.AutoEllipsis = $true
    $lblDes.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblDes.Text = _CorreuRemitentText (_CorreuViaDelCombo $cbVia) $remitent
    $form.Controls.Add($lblDes)
    $cbVia.add_SelectedIndexChanged({ $lblDes.Text = _CorreuRemitentText (_CorreuViaDelCombo $cbVia) $remitent }.GetNewClosure())
    $yPeu = [math]::Max(335, $y + 48)

    # Blau mari = enviar, vermell = no enviar. Un correu no es pot desenviar:
    # les dues accions han de ser distingibles d'un cop d'ull, i per aixo
    # tambe van cadascuna a una punta del peu.
    [void](_AddPeuBotons $form @(
        @{ Nom = 'No'; Text = 'No enviar'; Resultat = 'Cancel'; Esc = $true; Estil = 'accent'; Fons = $Script:CorreuVermell; FonsHover = $Script:CorreuVermellHover }) @(
        @{ Nom = 'Ok'; Text = 'Enviar'; Resultat = 'OK'; Intro = $true; Estil = 'accent'; Fons = $Script:CorreuBlauMari; FonsHover = $Script:CorreuBlauMariHover }) $yPeu)
    # El peu baixa si hi ha moltes adreces de CCO: la finestra creix amb ell
    # (abans era fixa a 420 i el peu a 335).
    $form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, ($yPeu + 46))

    # Scroll vertical i ajust a la pantalla (vegeu suport/UiFinestra.ps1).
    $form.add_Shown({ param($s, $e) _AjustaFinestraAPantalla $s })
    if ($form.ShowDialog() -ne 'OK') { return $null }

    $tos = @($tb.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $bccs = @()
    foreach ($cb in $checks) { if ($cb.Checked) { $bccs += $cb.Text } }
    $via = _CorreuViaDelCombo $cbVia
    if ($via -ne (Get-CorreuVia)) { [void](Set-CorreuVia $via) }
    return @{ To = $tos; Bcc = $bccs; Via = $via }
}

# Envia el correu per a un .docx concret. Reutilitzable (eina i final de generacio).
function Send-CorreuPerDocx($docxPath) {
    if (-not $docxPath -or -not (Test-Path -LiteralPath $docxPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat cap informe (.docx) per enviar.",'Enviar correu','OK','Information') | Out-Null
        return
    }

    # De QUINA ACTIVITAT es aquest informe? Surt del document que s'envia (nom
    # del fitxer + capcalera), NO de l'ultim informe generat: un Seguiment no
    # passa per l'assistent i Load-LastReport encara duia una altra activitat.
    # Es mira ABANS de llegir el cos: si s'ha de preguntar o es cancel.la, no
    # s'ha fet feina de franc (l'Excel de l'activitat).
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
        $build = _BuildCorreu (_CorreuDocumentXml $docxPath) $header ((Get-Date).ToString('dd/MM/yyyy'))
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error llegint l'informe:`n$($_.Exception.Message)",'Enviar correu','OK','Error') | Out-Null
        return
    }

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
    $via = [string]$res.Via
    # Les claus d'EmailJS nomes calen si s'envia per EmailJS (abans es miraven
    # d'entrada, i sense elles no es podia ni provar l'Outlook).
    if (-not (_CorreuViaEsOutlook $via)) {
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
    }
    $res = _CorreuDestinatariBuit $res.To $res.Bcc @(_CorreuBccOpcions)
    if (-not $res.To -or @($res.To).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Indica almenys un destinatari.','Enviar correu','OK','Warning') | Out-Null
        return
    }
    $toStr  = ($res.To -join ',')
    $bccStr = ($res.Bcc -join ',')
    $ses = $null
    try {
        $ses = Open-CorreuSessio $via
        Send-CorreuSessio $ses $toStr $bccStr $build.Subject $build.Html
        $resum = if ($via -eq 'outlook-esborrany') { "Correu DESAT a Esborranys de l'Outlook (no s'ha enviat), per a: $toStr" }
                 elseif ($res.Prova) { "Correu de prova enviat només a: $toStr" } else { "Correu enviat a: $toStr" }
        if ($bccStr) { $resum += "`nCCO: $bccStr" }
        $resum += "`n`n(" + (_CorreuViaText $via) + $(if ($ses.Remitent) { ', des de ' + $ses.Remitent } else { '' }) + ')'
        [System.Windows.Forms.MessageBox]::Show($resum,'Enviar correu','OK','Information') | Out-Null
    } catch {
        $txt = if ($null -eq $ses) { [string]$_.Exception.Message } else { _CorreuSessioError $ses $_ }
        [System.Windows.Forms.MessageBox]::Show($txt,'Enviar correu','OK','Error') | Out-Null
    } finally {
        Close-CorreuSessio $ses
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
