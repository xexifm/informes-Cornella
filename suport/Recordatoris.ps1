#requires -Version 5.1
<#
  Recordatoris.ps1 - Eina "Recordatoris" (secció EINES).

  Envia recordatoris PERIÒDICS als titulars de les activitats que tenen un
  tràmit pendent, a partir de l'estat de la BASE D'INFORMES (informes-db.json).

  DUES CAMPANYES INDEPENDENTS dins d'una sola eina (decisió de l'usuari), cada
  una amb la seva encesa/apagada, periodicitat, mode (manual/automàtic), topall
  per tanda i TEXT propi:
    · requeriments -> activitats amb estat_actual = 'Requeriment'
    · precintes    -> activitats amb estat_actual = 'Precinte / Cessament'

  QUOTA: EmailJS només deixa 200 correus/mes. El comptador (EmailQuota.ps1)
  compta TOTS els enviaments del PC i el límit de seguretat és 150, de manera
  que sempre en queden 50 de reserva.

  ON ES DESA: %LOCALAPPDATA%\InformesCornella\recordatoris.json (configuració +
  historial). Va a %LOCALAPPDATA% i NO al repositori perquè porta ID GIA i dates
  d'enviament; el repositori és PÚBLIC.
#>

# Pausa entre correus d'una mateixa tanda (EmailJS limita els cops seguits).
$Script:RecPausaMs = 1500
# Si la base d'informes és més vella que això, el mode AUTOMÀTIC no envia res.
$Script:RecMaxAntiguitatDbDies = 45
# A partir d'aquests dies, la finestra avisa que la base està desfasada.
$Script:RecAvisAntiguitatDbDies = 30

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# Les DUES campanyes, definides en UN SOL LLOC: la finestra i l'execució
# automàtica en beuen, així no es poden desincronitzar. Array pla (es consumeix
# amb @() al lloc de la crida).
function _RecCampanyes {
    return @(
        [pscustomobject]@{ Clau = 'requeriments'; Nom = 'Requeriments'; Estats = @('Requeriment') }
        [pscustomobject]@{ Clau = 'precintes';    Nom = 'Precintes';    Estats = @('Precinte / Cessament') }
    )
}

function _RecCampanyaPerClau([string]$clau) {
    foreach ($c in @(_RecCampanyes)) { if ($c.Clau -eq $clau) { return $c } }
    return $null
}

function _RecPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'recordatoris.json'))
}

function _RecLogPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'recordatoris-log.txt'))
}

# L'article 5 de l'Ordenança, LITERAL (versió vigent des del 19/06/2025). Viu en
# una funció pròpia perquè les dues campanyes el comparteixen i perquè una prova
# el pugui vigilar: si algú l'esborra del text per defecte, la suite ho diu.
function _RecArticle5 {
    return @(
        "**Article 5. Condicionant per a la transmissió de l'activitat**"
        "//No es poden transmetre les comunicacions prèvies, ni les llicències, ni les autoritzacions, quan llur titular, explotador o organitzador sigui objecte d'un expedient sancionador, d'un procediment de mesures provisionals o de qualsevol altre procediment d'exigència de responsabilitats administratives, mentre no s'hagi complert la sanció imposada, no s'hagi aixecat la mesura provisional, no s'hagi resolt l'arxiu de l'expedient per manca de responsabilitats o no s'hagi acreditat suficientment que la responsabilitat en la comissió de la infracció no afecta el propietari de l'establiment o el titular de la llicència o comunicació prèvia. Tampoc no es poden transmetre les comunicacions, ni les llicències, ni les autoritzacions subjectes a un expedient de revocació o caducitat, fins que no hi hagi una resolució ferma que confirmi la comunicació, la llicència o l'autorització.//"
    ) -join "`n"
}

# L'avís de "si ja ho has presentat, no en facis cas", bilingüe. També en funció
# pròpia i vigilat per una prova: és el que evita ensurts quan el titular ja ha
# complert i la base encara no ho sap.
function _RecAvisJaPresentat {
    return @(
        "**Si ja heu presentat la documentació, no cal que tingueu en compte aquest correu.** Les nostres dades es revisen periòdicament i pot ser que la vostra presentació encara no hi consti."
        "**Si ya ha presentado la documentación, no tenga en cuenta este correo.** Nuestros datos se revisan periódicamente y es posible que su presentación aún no conste."
    ) -join "`n"
}

function _RecPeu {
    return @(
        '________________________________________'
        ''
        "Departament d'Activitats · Ajuntament de Cornellà de Llobregat · Carrer de l'Energia, 97 · Tel. 93 377 02 12 (ext. 1227)"
        "Aquest és un recordatori informatiu de caràcter automàtic i NO té la consideració de notificació administrativa. / Este es un recordatorio informativo de carácter automático y NO tiene la consideración de notificación administrativa."
    ) -join "`n"
}

# Textos PER DEFECTE de cada campanya (assumpte + cos). Editables des de l'eina.
function _RecDefaultTextos([string]$clau) {
    $capcalera = @(
        'Activitat: {ACTIVITAT}'
        'Adreça: {ADRECA}'
        'ID GIA: {ID_GIA}'
        'Titular: {TITULAR}'
        ''
    ) -join "`n"

    if ($clau -eq 'precintes') {
        $cos = @(
            $capcalera
            (_RecAvisJaPresentat)
            ''
            '**Català**'
            'Benvolgut/da,'
            "Segons les nostres dades, l'activitat situada a {ADRECA} consta amb una mesura de **precinte / cessament** en vigor des de l'informe de data **{DATA_INFORME}**, i a hores d'ara no ens consta que s'hagi esmenat la situació ni que se n'hagi sol·licitat l'aixecament."
            "Us recordem que, mentre la mesura estigui vigent, l'activitat no es pot exercir. Per demanar-ne l'aixecament cal presentar la documentació que acrediti que s'han esmenat els incompliments."
            "Podeu presentar la documentació mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat, a l'atenció del **Departament d'Activitats**:"
            'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
            ''
            '**Castellano**'
            'Estimado/a,'
            "Según nuestros datos, la actividad situada en {ADRECA} consta con una medida de **precinto / cese** en vigor desde el informe de fecha **{DATA_INFORME}**, y a día de hoy no nos consta que se haya subsanado la situación ni que se haya solicitado su levantamiento."
            "Le recordamos que, mientras la medida esté vigente, la actividad no se puede ejercer. Para solicitar su levantamiento debe presentar la documentación que acredite que se han subsanado los incumplimientos."
            "Puede presentar la documentación mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat, a la atención del **Departamento de Actividades**:"
            'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
            ''
            (_RecArticle5)
            ''
            (_RecPeu)
        ) -join "`n"
        return ,([ordered]@{ assumpte = 'Precinte / cessament en vigor · GIA {ID_GIA}'; cos = $cos })
    }

    $cos = @(
        $capcalera
        (_RecAvisJaPresentat)
        ''
        '**Català**'
        'Benvolgut/da,'
        "Segons les nostres dades, l'activitat situada a {ADRECA} té pendent el compliment del **requeriment** notificat amb l'informe de data **{DATA_INFORME}**, i a hores d'ara no ens consta que s'hagi presentat la documentació requerida."
        "Us demanem que hi doneu compliment tan aviat com sigui possible. El fet de no atendre un requeriment pot donar lloc a la incoació d'un expedient sancionador i a l'adopció de mesures provisionals."
        "Podeu presentar la documentació mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat, a l'atenció del **Departament d'Activitats**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        '**Castellano**'
        'Estimado/a,'
        "Según nuestros datos, la actividad situada en {ADRECA} tiene pendiente el cumplimiento del **requerimiento** notificado con el informe de fecha **{DATA_INFORME}**, y a día de hoy no nos consta que se haya presentado la documentación requerida."
        "Le pedimos que le dé cumplimiento lo antes posible. No atender un requerimiento puede dar lugar a la incoación de un expediente sancionador y a la adopción de medidas provisionales."
        "Puede presentar la documentación mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat, a la atención del **Departamento de Actividades**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        (_RecArticle5)
        ''
        (_RecPeu)
    ) -join "`n"
    return ,([ordered]@{ assumpte = 'Requeriment pendent · GIA {ID_GIA}'; cos = $cos })
}

# Text d'ajuda amb les variables disponibles (editor de textos).
function _RecAjuda {
    return ('Variables: {ID_GIA} {TITULAR} {ADRECA} {ACTIVITAT} {DATA_INFORME} {DATA}   ' +
            [char]0x00B7 + '   **negreta**   ' + [char]0x00B7 + '   //cursiva//   ' +
            [char]0x00B7 + '   els enllaços http es fan clicables')
}

# Configuració PER DEFECTE d'una campanya.
function _RecDefaultConfig([string]$clau) {
    $tx = _RecDefaultTextos $clau
    return @{
        actiu             = $false          # apagada fins que l'usuari l'encengui
        mode              = 'manual'
        periodicitatDies  = 60
        esperaInicialDies = 30
        maxPerTanda       = 15
        # Les adreces de CCO surten del mateix lloc que les de l'eina "Enviar
        # correu": la clau 'bcc' de docs\dades\email-textos.json. Aqui hi havia
        # la primera d'aquelles quatre escrita a ma, que era la TERCERA copia de
        # la mateixa llista al projecte. Es queden les marcades per defecte.
        bcc               = @(@(_CorreuBccOpcions) | Where-Object { $_.Default } | ForEach-Object { [string]$_.Addr })
        assumpte          = [string]$tx['assumpte']
        cos               = [string]$tx['cos']
    }
}

# Normalitza/valida una configuració llegida del JSON contra els valors per
# defecte. Un valor absurd (0 dies, text buit) cau al de defecte. PURA.
function _RecNormalitzaConfig($cfg, [string]$clau) {
    $def = _RecDefaultConfig $clau
    if ($null -eq $cfg) { return $def }
    $out = @{}
    foreach ($k in @($def.Keys)) { $out[$k] = $def[$k] }
    $get = {
        param($o, $n)
        try { if ($o.PSObject.Properties[$n]) { return $o.$n } } catch { }
        try { if ($o -is [hashtable] -and $o.ContainsKey($n)) { return $o[$n] } } catch { }
        return $null
    }
    $v = & $get $cfg 'actiu';  if ($null -ne $v) { $out['actiu'] = [bool]$v }
    $v = & $get $cfg 'mode';   if ("$v" -eq 'auto' -or "$v" -eq 'manual') { $out['mode'] = [string]$v }
    foreach ($n in @('periodicitatDies', 'esperaInicialDies', 'maxPerTanda')) {
        $v = & $get $cfg $n
        if ($null -ne $v) { try { $i = [int]$v; if ($i -gt 0) { $out[$n] = $i } } catch { } }
    }
    # L'espera inicial SÍ que pot ser 0 (avisar de seguida).
    $v = & $get $cfg 'esperaInicialDies'
    if ($null -ne $v) { try { $i = [int]$v; if ($i -ge 0) { $out['esperaInicialDies'] = $i } } catch { } }
    foreach ($n in @('assumpte', 'cos')) {
        $v = & $get $cfg $n
        if (-not [string]::IsNullOrWhiteSpace([string]$v)) { $out[$n] = [string]$v }
    }
    $v = & $get $cfg 'bcc'
    if ($null -ne $v) { $out['bcc'] = @($v | ForEach-Object { [string]$_ } | Where-Object { $_ -like '*@*' }) }
    return $out
}

# Dies transcorreguts des d'una data 'yyyy-MM-dd'. Retorna -1 si no es pot
# llegir (data buida o mal formada). PURA.
function _RecDiesDes([string]$dataIso, [datetime]$avui) {
    if ([string]::IsNullOrWhiteSpace($dataIso)) { return -1 }
    if ($null -eq $avui -or $avui -eq [datetime]::MinValue) { $avui = Get-Date }
    $d = [datetime]::MinValue
    $ok = [datetime]::TryParseExact([string]$dataIso, 'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$d)
    if (-not $ok) { return -1 }
    return [int]($avui.Date - $d.Date).TotalDays
}

# Data de l'informe que DETERMINA l'estat de l'activitat (l'últim no ignorat).
# Reutilitza _InformeQueDeterminaEstat (Informes.ps1): l'estat i la data han de
# sortir del MATEIX informe, si no el recordatori diria una data que no lliga.
function _RecDataInforme($act) {
    if ($null -eq $act) { return '' }
    $infs = @()
    try { $infs = @($act.informes) } catch { $infs = @() }
    if ($infs.Count -eq 0) { return '' }
    $inf = _InformeQueDeterminaEstat $infs
    if ($null -eq $inf) { return '' }
    return [string]$inf.data
}

# Entrada d'historial d'un GIA, normalitzada. PURA.
function _RecHistEntrada($historialCampanya, [string]$gia) {
    $buida = @{ ultim = ''; compte = 0; excloure = $false; enviaments = @() }
    if ($null -eq $historialCampanya -or [string]::IsNullOrWhiteSpace($gia)) { return $buida }
    $e = $null
    try { if ($historialCampanya.ContainsKey($gia)) { $e = $historialCampanya[$gia] } } catch { }
    if ($null -eq $e) { return $buida }
    $o = @{ ultim = ''; compte = 0; excloure = $false; enviaments = @() }
    try { if ($e.ContainsKey('ultim'))      { $o.ultim      = [string]$e['ultim'] } } catch { }
    try { if ($e.ContainsKey('compte'))     { $o.compte     = [int]$e['compte'] } } catch { }
    try { if ($e.ContainsKey('excloure'))   { $o.excloure   = [bool]$e['excloure'] } } catch { }
    try { if ($e.ContainsKey('enviaments')) { $o.enviaments = @($e['enviaments']) } } catch { }
    return $o
}

# DECIDEIX si a una activitat li toca recordatori avui. PURA i testejable: és el
# cor de l'eina. Retorna @{ Toca; Motiu; DataInforme; Dies }.
#
# Ordre de les regles (i el perquè de cada una):
#  1. Sense ID GIA -> fora: sense GIA no es pot creuar amb l'Excel i no hi ha
#     adreça de correu. Es compta a part, mai s'ignora en silenci.
#  2. Exclosa a mà -> fora.
#  3. Espera inicial: el termini del requeriment encara corre; avisar-ne abans
#     d'hora és empipar el titular.
#  4. Periodicitat: no es repeteix fins que han passat els dies configurats.
function _RecToca($act, $cfg, $hist, [datetime]$avui) {
    $gia = ''
    try { $gia = [string]$act.id_gia } catch { }
    $dataInf = _RecDataInforme $act
    $res = @{ Toca = $false; Motiu = ''; DataInforme = $dataInf; Dies = -1 }

    if ([string]::IsNullOrWhiteSpace($gia)) { $res.Motiu = 'sense ID GIA'; return $res }
    if ([bool]$hist.excloure)               { $res.Motiu = 'exclosa'; return $res }

    $espera = [int]$cfg['esperaInicialDies']
    $diesInf = _RecDiesDes $dataInf $avui
    $res.Dies = $diesInf
    if ($diesInf -lt 0) { $res.Motiu = "sense data d'informe"; return $res }
    if ($diesInf -lt $espera) {
        $res.Motiu = "espera inicial ($diesInf de $espera dies)"
        return $res
    }

    $ultim = [string]$hist.ultim
    if (-not [string]::IsNullOrWhiteSpace($ultim)) {
        $per = [int]$cfg['periodicitatDies']
        $diesUlt = _RecDiesDes $ultim $avui
        if ($diesUlt -lt 0) { $res.Motiu = 'últim recordatori il·legible'; return $res }
        if ($diesUlt -lt $per) {
            $res.Motiu = "enviat fa $diesUlt dies (cada $per)"
            return $res
        }
    }
    $res.Toca = $true
    $res.Motiu = 'toca'
    return $res
}

# Recorre la base i retorna TOTES les activitats de la campanya (toquin o no,
# perquè la finestra les pugui ensenyar amb el seu motiu), ja ordenades per
# prioritat: primer les que no han rebut mai cap recordatori, després per
# recordatori més antic i finalment per informe més antic.
# Retorna @{ Files = @(...); SenseGia = n }.
function _RecDueActivitats($db, $campanya, $cfg, $historialCampanya, [datetime]$avui) {
    $files = New-Object System.Collections.ArrayList
    $senseGia = 0
    if ($null -eq $db) { return @{ Files = @(); SenseGia = 0 } }
    $acts = @()
    try { $acts = @($db.activitats) } catch { $acts = @() }
    $estats = @($campanya.Estats)

    foreach ($act in $acts) {
        if ($null -eq $act) { continue }
        $estat = ''
        try { $estat = [string]$act.estat_actual } catch { }
        if ($estats -notcontains $estat) { continue }
        $gia = ''
        try { $gia = [string]$act.id_gia } catch { }
        $hist = _RecHistEntrada $historialCampanya $gia
        $d = _RecToca $act $cfg $hist $avui
        if ([string]::IsNullOrWhiteSpace($gia)) { $senseGia++ }
        [void]$files.Add([pscustomobject]@{
            Id          = $gia
            Titular     = [string]$act.titular
            Expedient   = [string]$act.expedient
            Estat       = $estat
            DataInforme = [string]$d.DataInforme
            Ultim       = [string]$hist.ultim
            Compte      = [int]$hist.compte
            Excloure    = [bool]$hist.excloure
            Toca        = [bool]$d.Toca
            Motiu       = [string]$d.Motiu
            Adreca      = ''
            Correus     = ''
            Sel         = [bool]$d.Toca
        })
    }
    # Prioritat: mai avisats primer ('' ordena abans que qualsevol data ISO),
    # després pel recordatori més antic i finalment per l'informe més antic.
    $ord = @($files | Sort-Object @{ Expression = { [string]$_.Ultim } },
                                  @{ Expression = { [string]$_.DataInforme } })
    return @{ Files = $ord; SenseGia = $senseGia }
}

# Substitueix les variables del text amb les dades d'una fila. PURA.
function _RecFillPh([string]$text, $row) {
    $today = (Get-Date).ToString('dd/MM/yyyy')
    $map = [ordered]@{
        '{ID_GIA}'       = [string]$row.Id
        '{TITULAR}'      = [string]$row.Titular
        '{ADRECA}'       = [string]$row.Adreca
        '{ACTIVITAT}'    = [string]$row.Activitat
        '{DATA_INFORME}' = [string]$row.DataInforme
        '{DATA}'         = $today
    }
    $t = [string]$text
    foreach ($k in @($map.Keys)) { $t = $t.Replace($k, [string]$map[$k]) }
    return $t
}

# Apunta un enviament a l'historial (retorna el mapa NOU). PURA.
function _RecHistorialActualitza($historialCampanya, [string]$gia, [string]$dataIso) {
    $h = @{}
    if ($null -ne $historialCampanya) {
        foreach ($k in @($historialCampanya.Keys)) { $h[$k] = $historialCampanya[$k] }
    }
    if ([string]::IsNullOrWhiteSpace($gia)) { return $h }
    $e = _RecHistEntrada $h $gia
    $env = @($e.enviaments)
    $env += [string]$dataIso
    $h[$gia] = @{
        ultim      = [string]$dataIso
        compte     = ([int]$e.compte + 1)
        excloure   = [bool]$e.excloure
        enviaments = $env
    }
    return $h
}

# Marca (o desmarca) una activitat com a EXCLOSA. PURA.
function _RecHistorialExclou($historialCampanya, [string]$gia, [bool]$excloure) {
    $h = @{}
    if ($null -ne $historialCampanya) {
        foreach ($k in @($historialCampanya.Keys)) { $h[$k] = $historialCampanya[$k] }
    }
    if ([string]::IsNullOrWhiteSpace($gia)) { return $h }
    $e = _RecHistEntrada $h $gia
    $h[$gia] = @{
        ultim      = [string]$e.ultim
        compte     = [int]$e.compte
        excloure   = $excloure
        enviaments = @($e.enviaments)
    }
    return $h
}

# PSCustomObject (del ConvertFrom-Json) -> hashtable, recursiu on cal. El JSON
# torna PSCustomObjects i l'historial s'indexa per GIA: sense això, .ContainsKey
# no existeix i tot l'historial es perdria en silenci.
function _RecObjAMapa($o) {
    $m = @{}
    if ($null -eq $o) { return $m }
    if ($o -is [hashtable]) {
        foreach ($k in @($o.Keys)) { $m[[string]$k] = $o[$k] }
        return $m
    }
    try {
        foreach ($p in $o.PSObject.Properties) { $m[[string]$p.Name] = $p.Value }
    } catch { }
    return $m
}

# Historial d'una campanya: mapa GIA -> mapa d'entrada. PURA.
function _RecHistorialAMapa($o) {
    $out = @{}
    $top = _RecObjAMapa $o
    foreach ($k in @($top.Keys)) { $out[[string]$k] = _RecObjAMapa $top[$k] }
    return $out
}

# Llegeix configuració + historial del disc, ja normalitzats. Mai llança.
function _RecLlegeix {
    $out = @{ campanyes = @{}; historial = @{} }
    $raw = Read-JsonFile (_RecPath)
    $camps = @{}
    $hist  = @{}
    if ($null -ne $raw) {
        try { $camps = _RecObjAMapa $raw.campanyes } catch { }
        try { $hist  = _RecObjAMapa $raw.historial } catch { }
    }
    foreach ($c in @(_RecCampanyes)) {
        $cfgRaw = $null
        if ($camps.ContainsKey($c.Clau)) { $cfgRaw = $camps[$c.Clau] }
        $out.campanyes[$c.Clau] = _RecNormalitzaConfig $cfgRaw $c.Clau
        $hRaw = $null
        if ($hist.ContainsKey($c.Clau)) { $hRaw = $hist[$c.Clau] }
        $out.historial[$c.Clau] = _RecHistorialAMapa $hRaw
    }
    return $out
}

# Escriu configuració + historial (UTF-8 sense BOM).
function _RecDesa($estat) {
    if ($null -eq $estat) { return }
    $path = _RecPath
    $dir = Split-Path -Parent $path
    try {
        Write-JsonFile $path ([pscustomobject]@{
            campanyes = [pscustomobject]$estat.campanyes
            historial = [pscustomobject]$estat.historial
        }) 12
    } catch { }
}

# Cos (multilínia) -> HTML. Una línia = un <div>; línia buida = espaiador.
# Cada línia passa per _TextToHtml (EnviarCorreu.ps1), que ja escapa, aplica
# **negreta** / //cursiva// i fa clicables els enllaços: no se'n fa cap còpia.
function _RecCosHtml([string]$cos) {
    $norm = ([string]$cos) -replace "`r`n", "`n"
    $lines = $norm -split "`n"
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

# Nom de la tasca programada del Windows (mode automàtic).
$Script:RecTascaNom = 'InformesCornella-Recordatoris'

# Línia d'ordres de la tasca programada. PURA: es prova sense Windows.
# Les rutes van ENTRE COMETES perquè el clone de l'usuari té espais (trampa ja
# coneguda: Start-Process -ArgumentList no enquota res).
function _RecSchtasksTr([string]$psExe, [string]$script) {
    return ('"' + [string]$psExe + '" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + [string]$script + '"')
}

# Arguments de schtasks per CREAR la tasca (diària a l'hora indicada). PURA.
function _RecSchtasksArgv([string]$nom, [string]$psExe, [string]$script, [string]$hora) {
    return @(
        '/Create', '/TN', [string]$nom,
        '/TR', (_RecSchtasksTr $psExe $script),
        '/SC', 'DAILY', '/ST', [string]$hora, '/F'
    )
}

# Antiguitat (en dies) de la base d'informes a partir del seu actualitzat_el.
# -1 si no es pot llegir. PURA.
function _RecAntiguitatDb($db, [datetime]$avui) {
    if ($null -eq $avui -or $avui -eq [datetime]::MinValue) { $avui = Get-Date }
    $s = ''
    try { $s = [string]$db.actualitzat_el } catch { }
    if ([string]::IsNullOrWhiteSpace($s)) { return -1 }
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParse($s, [ref]$d)) { return -1 }
    return [int]($avui.Date - $d.Date).TotalDays
}

# Carrega la base d'informes. Retorna $null si encara no s'ha generat mai.
function _RecCarregaDb {
    $p = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return (Read-JsonFile $p)
}

# Escriu una línia al registre dels recordatoris (diagnòstic del mode automàtic).
function _RecLog([string]$msg) {
    try {
        $p = _RecLogPath
        $dir = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $ln = '[' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '] ' + [string]$msg
        Add-Content -LiteralPath $p -Value $ln -Encoding UTF8
    } catch { }
}

# Completa una fila amb l'adreça, l'activitat i els correus, des de la cache de
# l'Excel (que es carrega UNA sola vegada per tanda: obrir l'Excel per cada
# correu seria inviable).
function _RecOmpleDadesFila($row, $cache) {
    $row.Adreca = ''
    $row.Correus = ''
    if ($null -eq $cache -or [string]::IsNullOrWhiteSpace([string]$row.Id)) { return $row }
    $act = Get-ActivitatFromCache $cache ([string]$row.Id)
    if ($null -eq $act) { return $row }
    try { if ($act.ContainsKey('ADRECA'))    { $row.Adreca = [string]$act['ADRECA'] } } catch { }
    try { if ($act.ContainsKey('ACTIVITAT')) { Add-Member -InputObject $row -NotePropertyName Activitat -NotePropertyValue ([string]$act['ACTIVITAT']) -Force } } catch { }
    $rao = ''; $rep = ''
    try { if ($act.ContainsKey('EMAIL'))     { $rao = [string]$act['EMAIL'] } } catch { }
    try { if ($act.ContainsKey('EMAIL_REP')) { $rep = [string]$act['EMAIL_REP'] } } catch { }
    $d = _CorreuDestinatarisPerDefecte $rao $rep
    $row.Correus = [string]$d.Text
    return $row
}

# ----------------------------------------------------------------------------
# ENVIAMENT D'UNA TANDA (la comparteixen el mode manual i l'automatic)
# ----------------------------------------------------------------------------
# $rows: files ja triades. $silenci: mode automatic (cap finestra).
# Retorna @{ Enviats; Fallats; SenseCorreu; Aturat; Motiu }.
function Invoke-RecordatorisTanda([string]$clau, $rows, [bool]$silenci) {
    $res = @{ Enviats = 0; Fallats = 0; SenseCorreu = 0; Aturat = $false; Motiu = '' }
    $files = @($rows)
    if ($files.Count -eq 0) { $res.Motiu = 'cap activitat'; return $res }

    $cfgAll = _RecLlegeix
    $cfg = $cfgAll.campanyes[$clau]
    if ($null -eq $cfg) { $res.Motiu = 'campanya desconeguda'; return $res }

    # Claus d'EmailJS: si en falta cap, val mes dir-ho que provar-ho N vegades.
    $ecfg = _CorreuConfig
    if (-not $ecfg.PublicKey -or -not $ecfg.ServiceId -or -not $ecfg.TemplateId) {
        $res.Motiu = "falten les claus d'EmailJS a docs\config.js"; $res.Aturat = $true; return $res
    }
    if (-not $ecfg.PrivateKey) {
        $res.Motiu = "falta la Private key d'EmailJS a $($ecfg.PrivatePath)"; $res.Aturat = $true; return $res
    }

    # Quota: el topall mana per sobre del maxPerTanda.
    $quota = _QuotaLlegeix
    $restant = _QuotaRestant $quota
    if ($restant -le 0) {
        $res.Motiu = "quota mensual exhaurida ($($quota.enviats)/$($quota.limit))"; $res.Aturat = $true; return $res
    }
    $maxTanda = [int]$cfg['maxPerTanda']
    $limitAra = [Math]::Min($files.Count, [Math]::Min($maxTanda, $restant))

    # L'Excel es carrega UNA sola vegada per a tota la tanda.
    $cache = $null
    try {
        $xls = Find-LatestActivitatsExcel
        if ($null -ne $xls) { $cache = Initialize-ActivitatsCache $xls.File }
    } catch { $cache = $null }
    if ($null -eq $cache) {
        $res.Motiu = "no s'ha pogut llegir l'Excel d'activitats (cal per als correus)"
        $res.Aturat = $true; return $res
    }

    $bcc = ''
    try { $bcc = (@($cfg['bcc']) -join ',') } catch { }

    # Finestra de progres amb Cancel.lar (mai en mode automatic).
    $st = @{ Cancel = $false }
    $form = $null; $lbl = $null; $bar = $null
    if (-not $silenci) {
        $form = _NewForm
        $form.Text = 'Enviant recordatoris'
        $form.FormBorderStyle = 'FixedDialog'
        $form.ControlBox = $false
        $form.ClientSize = New-Object System.Drawing.Size(460, 130)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(16, 16)
        $lbl.Size = New-Object System.Drawing.Size(428, 40)
        $lbl.Text = 'Preparant...'
        $form.Controls.Add($lbl)
        $bar = New-Object System.Windows.Forms.ProgressBar
        $bar.Location = New-Object System.Drawing.Point(16, 62)
        $bar.Size = New-Object System.Drawing.Size(428, 18)
        $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $limitAra)
        $form.Controls.Add($bar)
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = 'Cancel·lar'
        $btn.Location = New-Object System.Drawing.Point(354, 90)
        $btn.Size = New-Object System.Drawing.Size(90, 28)
        $btn.add_Click({ $st.Cancel = $true }.GetNewClosure())
        $form.Controls.Add($btn)
        $form.Show(); [System.Windows.Forms.Application]::DoEvents()
    }

    try {
        $n = 0
        foreach ($row in $files) {
            if ($n -ge $limitAra) { break }
            if ($st.Cancel) { $res.Aturat = $true; $res.Motiu = "cancel·lat per l'usuari"; break }

            $row = _RecOmpleDadesFila $row $cache
            if ([string]::IsNullOrWhiteSpace([string]$row.Correus)) {
                $res.SenseCorreu++
                _RecLog "GIA $($row.Id): sense correu a l'Excel, omesa"
                continue
            }

            if ($null -ne $lbl) {
                $lbl.Text = "GIA $($row.Id) - $($row.Titular)`n$($n + 1) de $limitAra"
                $bar.Value = [Math]::Min($bar.Maximum, $n)
                [System.Windows.Forms.Application]::DoEvents()
            }

            $assumpte = _RecFillPh ([string]$cfg['assumpte']) $row
            $cos      = _RecFillPh ([string]$cfg['cos']) $row
            $html     = _RecCosHtml $cos
            $to       = (([string]$row.Correus) -split '\s*;\s*' | Where-Object { $_ }) -join ','

            # El try/catch va DINS del bucle: un error d'una activitat no pot
            # endur-se la resta de la tanda.
            try {
                Send-EmailJs $ecfg $to $bcc $assumpte $html
                $res.Enviats++
                $n++
                # Es desa DESPRES DE CADA enviament: si peta o es cancel.la, el
                # que ja ha sortit consta i no es tornara a enviar.
                $avuiIso = (Get-Date).ToString('yyyy-MM-dd')
                $cfgAll.historial[$clau] = _RecHistorialActualitza $cfgAll.historial[$clau] ([string]$row.Id) $avuiIso
                _RecDesa $cfgAll
                _QuotaApunta 1
                _RecLog "GIA $($row.Id): enviat a $to"
                if ($n -lt $limitAra) { Start-Sleep -Milliseconds $Script:RecPausaMs }
            } catch {
                $res.Fallats++
                $txt = _EmailJsRespError $_
                _RecLog "GIA $($row.Id): ERROR -> $txt"
                # Un 401/403 vol dir que TOTS els seguents fallaran igual: no te
                # cap sentit cremar la tanda sencera provant-ho.
                if ($txt -match 'HTTP (401|403)') {
                    $res.Aturat = $true
                    $res.Motiu = $txt
                    break
                }
            }
        }
    } finally {
        if ($null -ne $form) { try { $form.Close() } catch { } }
    }
    return $res
}

# ----------------------------------------------------------------------------
# EDITOR DELS TEXTOS D'UNA CAMPANYA (mateix patró que "Textos del correu")
# ----------------------------------------------------------------------------
function Invoke-RecordatorisTextos([string]$clau) {
    $camp = _RecCampanyaPerClau $clau
    if ($null -eq $camp) { return $false }
    $estat = _RecLlegeix
    $cfg = $estat.campanyes[$clau]

    $form = _NewForm
    $form.Text = 'Text del recordatori - ' + $camp.Nom
    $form.ClientSize = New-Object System.Drawing.Size(860, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 480)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)

    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = 'Assumpte:'
    $lblA.Location = New-Object System.Drawing.Point(16, 12)
    $lblA.Size = New-Object System.Drawing.Size(120, 20)
    $panel.Controls.Add($lblA)

    $txtA = New-Object System.Windows.Forms.TextBox
    $txtA.Location = New-Object System.Drawing.Point(16, 34)
    $txtA.Size = New-Object System.Drawing.Size(812, 24)
    $txtA.Anchor = 'Top,Left,Right'
    $txtA.Text = [string]$cfg['assumpte']
    $panel.Controls.Add($txtA)

    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Cos del correu:'
    $lblC.Location = New-Object System.Drawing.Point(16, 66)
    $lblC.Size = New-Object System.Drawing.Size(200, 20)
    $panel.Controls.Add($lblC)

    $txtC = New-Object System.Windows.Forms.TextBox
    $txtC.Location = New-Object System.Drawing.Point(16, 88)
    $txtC.Size = New-Object System.Drawing.Size(812, 430)
    $txtC.Anchor = 'Top,Left,Right,Bottom'
    $txtC.Multiline = $true
    $txtC.ScrollBars = 'Vertical'
    $txtC.AcceptsReturn = $true
    $txtC.WordWrap = $true
    $txtC.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtC.Text = [string]$cfg['cos']
    $panel.Controls.Add($txtC)

    $lblH = New-Object System.Windows.Forms.Label
    $lblH.Text = _RecAjuda
    $lblH.Location = New-Object System.Drawing.Point(16, 524)
    $lblH.Size = New-Object System.Drawing.Size(812, 34)
    $lblH.Anchor = 'Left,Right,Bottom'
    $lblH.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    $panel.Controls.Add($lblH)

    $btnDef = New-Object System.Windows.Forms.Button
    $btnDef.Text = 'Restaurar el text per defecte'
    $btnDef.Location = New-Object System.Drawing.Point(16, 562)
    $btnDef.Size = New-Object System.Drawing.Size(210, 30)
    $btnDef.Anchor = 'Left,Bottom'
    $panel.Controls.Add($btnDef)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Desar'
    $btnOk.Location = New-Object System.Drawing.Point(638, 562)
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.Anchor = 'Right,Bottom'
    $btnOk.DialogResult = 'OK'
    $panel.Controls.Add($btnOk)

    $btnNo = New-Object System.Windows.Forms.Button
    $btnNo.Text = 'Cancel·lar'
    $btnNo.Location = New-Object System.Drawing.Point(738, 562)
    $btnNo.Size = New-Object System.Drawing.Size(90, 30)
    $btnNo.Anchor = 'Right,Bottom'
    $btnNo.DialogResult = 'Cancel'
    $panel.Controls.Add($btnNo)

    $form.Controls.Add($panel)
    [void](_AddBrandHeader $form ('Recordatoris - ' + $camp.Nom) 'Text que rebrà el titular' 56)
    $form.AcceptButton = $null
    $form.CancelButton = $btnNo

    $btnDef.add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show(
            'Vols recuperar el text per defecte? Es perdrà el que hi ha ara.',
            'Recordatoris', 'YesNo', 'Question')
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            $d = _RecDefaultTextos $clau
            $txtA.Text = [string]$d['assumpte']
            $txtC.Text = [string]$d['cos']
        }
    }.GetNewClosure())

    if ($form.ShowDialog() -ne 'OK') { return $false }

    $estat.campanyes[$clau]['assumpte'] = [string]$txtA.Text
    $estat.campanyes[$clau]['cos']      = [string]$txtC.Text
    _RecDesa $estat
    return $true
}

# ----------------------------------------------------------------------------
# TASCA PROGRAMADA DEL WINDOWS (mode automàtic)
# ----------------------------------------------------------------------------
function _RecScriptAuto {
    return [string](Join-Path $PSScriptRoot 'RecordatorisAuto.ps1')
}

function _RecExecutaSchtasks($argv) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'schtasks.exe'
    # Les cometes les posem NOSALTRES (_ArgvToCommandLine, PdfSignar.ps1): el
    # clone té espais i -ArgumentList no enquota res.
    $psi.Arguments = _ArgvToCommandLine $argv
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $out = $p.StandardOutput.ReadToEnd() + $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    return @{ Codi = $p.ExitCode; Sortida = $out }
}

function Invoke-RecordatorisTasca {
    $script = _RecScriptAuto
    if (-not (Test-Path -LiteralPath $script)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat:`n$script", 'Recordatoris', 'OK', 'Error') | Out-Null
        return
    }
    $msg = "Vols programar l'enviament AUTOMÀTIC dels recordatoris?`n`n" +
           "Es crearà una tasca del Windows que cada dia a les 09:00 enviarà els`n" +
           "recordatoris de les campanyes que tinguis en mode Automàtic.`n`n" +
           "Tingues en compte que:`n" +
           " · Només s'executa amb el PC engegat i la sessió iniciada.`n" +
           " · Els correus surten SENSE que ningú els revisi.`n" +
           " · Si la base d'informes té més de $($Script:RecMaxAntiguitatDbDies) dies, no enviarà res.`n`n" +
           "Sí = crear-la / No = esborrar-la / Cancel·lar = deixar-ho com està."
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Recordatoris automàtics', 'YesNoCancel', 'Question')
    if ($r -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        $psExe = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
        $argv = _RecSchtasksArgv $Script:RecTascaNom $psExe $script '09:00'
        $res = _RecExecutaSchtasks $argv
        if ($res.Codi -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Tasca creada.`n`nNom: $($Script:RecTascaNom)`nCada dia a les 09:00.", 'Recordatoris', 'OK', 'Information') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut crear la tasca:`n`n$($res.Sortida)", 'Recordatoris', 'OK', 'Error') | Out-Null
        }
    } else {
        $res = _RecExecutaSchtasks @('/Delete', '/TN', $Script:RecTascaNom, '/F')
        if ($res.Codi -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Tasca esborrada.', 'Recordatoris', 'OK', 'Information') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut esborrar (potser no existia):`n`n$($res.Sortida)", 'Recordatoris', 'OK', 'Warning') | Out-Null
        }
    }
}

# ----------------------------------------------------------------------------
# FINESTRA PRINCIPAL
# ----------------------------------------------------------------------------
function Invoke-Recordatoris {
    $db = _RecCarregaDb
    if ($null -eq $db) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base' (secció INFORMES).",
            'Recordatoris', 'OK', 'Information') | Out-Null
        return
    }
    $antig = _RecAntiguitatDb $db (Get-Date)
    $estat = _RecLlegeix
    $quota = _QuotaLlegeix

    $form = _NewForm
    $form.Text = 'Recordatoris'
    $form.ClientSize = New-Object System.Drawing.Size(1080, 660)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $tabs.Padding = New-Object System.Drawing.Point(14, 5)

    # Panell d'informació: quota + frescor de la base. La data de la base és la
    # protecció principal: amb una base vella s'escriuria a qui ja ha complert.
    $info = New-Object System.Windows.Forms.Panel
    $info.Dock = 'Top'; $info.Height = 50
    $lblQ = New-Object System.Windows.Forms.Label
    $lblQ.Location = New-Object System.Drawing.Point(16, 6)
    $lblQ.Size = New-Object System.Drawing.Size(1040, 18)
    $lblQ.Text = ("EmailJS aquest mes: $($quota.enviats) / $($quota.limit) correus " +
                  "(reserva de $($Script:QuotaLimitCompte - $quota.limit) sobre els $($Script:QuotaLimitCompte) del compte)")
    $info.Controls.Add($lblQ)
    $lblD = New-Object System.Windows.Forms.Label
    $lblD.Location = New-Object System.Drawing.Point(16, 26)
    $lblD.Size = New-Object System.Drawing.Size(1040, 18)
    if ($antig -lt 0) {
        $lblD.Text = "Base d'informes: data desconeguda. Actualitza-la abans d'enviar res."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(176, 0, 32)
    } elseif ($antig -ge $Script:RecAvisAntiguitatDbDies) {
        $lblD.Text = "ATENCIÓ: la base d'informes té $antig dies. Actualitza-la abans d'enviar: podries escriure a qui ja ha complert."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(176, 0, 32)
    } else {
        $lblD.Text = "Base d'informes: actualitzada fa $antig dies · $(@($db.activitats).Count) activitats."
        $lblD.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    }
    $info.Controls.Add($lblD)

    foreach ($camp in @(_RecCampanyes)) {
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $camp.Nom
        $tab.BackColor = [System.Drawing.Color]::White
        [void]$tabs.TabPages.Add($tab)
        _RecMuntaTab $tab $camp $estat $db
    }

    $form.Controls.Add($tabs)
    $form.Controls.Add($info)
    [void](_AddBrandHeader $form 'Recordatoris' 'Avisos periòdics als titulars amb tràmits pendents' 56)
    [void]$form.ShowDialog()
}

# Munta el contingut d'UNA pestanya (una campanya). Les dues pestanyes són
# independents: cada una té la seva configuració, el seu text i el seu historial.
function _RecMuntaTab($tab, $camp, $estat, $db) {
    $clau = [string]$camp.Clau
    $cfg  = $estat.campanyes[$clau]
    # Hashtable de funcions: es captura per REFERÈNCIA, que és l'única manera
    # que els handlers vegin les funcions que es defineixen més avall.
    $fn = @{}
    $ui = @{ Rows = @(); SenseGia = 0 }

    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    $grid.AllowUserToResizeRows = $false
    $cSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cSel.HeaderText = 'Enviar'; $cSel.Width = 55
    [void]$grid.Columns.Add($cSel)
    foreach ($c in @(
        @{ H = 'GIA';           W = 70  }
        @{ H = 'Titular';       W = 260 }
        @{ H = 'Estat';         W = 130 }
        @{ H = 'Data informe';  W = 95  }
        @{ H = 'Últim avís';    W = 95  }
        @{ H = 'Avisos';        W = 60  }
        @{ H = 'Situació';      W = 230 }
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c.H; $col.Width = $c.W; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }

    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 84

    $chkActiu = New-Object System.Windows.Forms.CheckBox
    $chkActiu.Text = 'Campanya activa'
    $chkActiu.Location = New-Object System.Drawing.Point(16, 10)
    $chkActiu.Size = New-Object System.Drawing.Size(130, 22)
    $chkActiu.Checked = [bool]$cfg['actiu']
    $top.Controls.Add($chkActiu)

    $mkNum = {
        param($etiqueta, $x, $valor, $min, $max)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $etiqueta
        $l.Location = New-Object System.Drawing.Point($x, 13)
        $l.Size = New-Object System.Drawing.Size(96, 18)
        $top.Controls.Add($l)
        $n = New-Object System.Windows.Forms.NumericUpDown
        $n.Location = New-Object System.Drawing.Point(($x + 98), 10)
        $n.Size = New-Object System.Drawing.Size(58, 22)
        $n.Minimum = $min; $n.Maximum = $max; $n.Value = $valor
        $top.Controls.Add($n)
        return $n
    }
    $numPer  = & $mkNum 'Cada (dies):'      160 ([int]$cfg['periodicitatDies'])  1 365
    $numEsp  = & $mkNum 'Espera (dies):'    330 ([int]$cfg['esperaInicialDies']) 0 365
    $numMax  = & $mkNum 'Màx. per tanda:'   500 ([int]$cfg['maxPerTanda'])       1 150

    $rbMan = New-Object System.Windows.Forms.RadioButton
    $rbMan.Text = 'Manual'
    $rbMan.Location = New-Object System.Drawing.Point(690, 10)
    $rbMan.Size = New-Object System.Drawing.Size(80, 22)
    $rbAut = New-Object System.Windows.Forms.RadioButton
    $rbAut.Text = 'Automàtic'
    $rbAut.Location = New-Object System.Drawing.Point(772, 10)
    $rbAut.Size = New-Object System.Drawing.Size(90, 22)
    if ([string]$cfg['mode'] -eq 'auto') { $rbAut.Checked = $true } else { $rbMan.Checked = $true }
    $top.Controls.Add($rbMan); $top.Controls.Add($rbAut)

    $chkNomes = New-Object System.Windows.Forms.CheckBox
    $chkNomes.Text = 'Només els que toquen avui'
    $chkNomes.Location = New-Object System.Drawing.Point(16, 48)
    $chkNomes.Size = New-Object System.Drawing.Size(200, 22)
    $chkNomes.Checked = $true
    $top.Controls.Add($chkNomes)

    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Location = New-Object System.Drawing.Point(560, 51)
    $lblN.Size = New-Object System.Drawing.Size(400, 18)
    $lblN.ForeColor = [System.Drawing.Color]::FromArgb(107, 116, 128)
    $top.Controls.Add($lblN)

    $bot = New-Object System.Windows.Forms.Panel
    $bot.Dock = 'Bottom'; $bot.Height = 48
    $mkBtn = {
        param($text, $x, $w)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        $b.Location = New-Object System.Drawing.Point($x, 9)
        $b.Size = New-Object System.Drawing.Size($w, 30)
        $bot.Controls.Add($b)
        return $b
    }
    $btnText = & $mkBtn 'Editar text...'      16  120
    $btnCsv  = & $mkBtn 'Exportar CSV'        144 110
    $btnExc  = & $mkBtn 'Excloure / incloure' 262 150
    $btnAuto = & $mkBtn 'Automàtic...'        420 110
    $btnSend = & $mkBtn 'Enviar tanda'        870 140
    $btnSend.Anchor = 'Top,Right'

    # --- Funcions de la pestanya --------------------------------------------
    $txtCerca = _AddSearchBox $top 240 48 300 'Cerca:' { & $fn.Pinta }

    $fn.Desa = {
        $cfg['actiu']             = [bool]$chkActiu.Checked
        $cfg['periodicitatDies']  = [int]$numPer.Value
        $cfg['esperaInicialDies'] = [int]$numEsp.Value
        $cfg['maxPerTanda']       = [int]$numMax.Value
        $cfg['mode']              = if ($rbAut.Checked) { 'auto' } else { 'manual' }
        $estat.campanyes[$clau]   = $cfg
        _RecDesa $estat
    }.GetNewClosure()

    $fn.Calcula = {
        $r = _RecDueActivitats $db $camp $cfg $estat.historial[$clau] (Get-Date)
        $ui.Rows = @($r.Files)
        $ui.SenseGia = [int]$r.SenseGia
    }.GetNewClosure()

    $fn.Pinta = {
        $cerca = ''
        try { $cerca = ([string]$txtCerca.Text).Trim().ToLower() } catch { }
        $nomes = [bool]$chkNomes.Checked
        $grid.Rows.Clear()
        $nToca = 0
        foreach ($row in @($ui.Rows)) {
            if ($row.Toca) { $nToca++ }
            if ($nomes -and -not $row.Toca) { continue }
            if ($cerca -ne '') {
                $hay = ([string]$row.Id + ' ' + [string]$row.Titular + ' ' + [string]$row.Estat).ToLower()
                if (-not $hay.Contains($cerca)) { continue }
            }
            $i = $grid.Rows.Add(@(
                [bool]$row.Sel, [string]$row.Id, [string]$row.Titular, [string]$row.Estat,
                [string]$row.DataInforme, [string]$row.Ultim, [string]$row.Compte, [string]$row.Motiu
            ))
            $grid.Rows[$i].Tag = $row
            if ($row.Excloure) { $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150) }
        }
        $lblN.Text = "$nToca activitats toquen avui, de $(@($ui.Rows).Count) en aquest estat" +
                     $(if ($ui.SenseGia -gt 0) { " · $($ui.SenseGia) sense GIA (no es poden avisar)" } else { '' })
    }.GetNewClosure()

    $fn.Refresca = { & $fn.Calcula; & $fn.Pinta }.GetNewClosure()

    # La casella "Enviar" es desa a l'objecte fila: així sobreviu a filtres i
    # a repintats (mateix patró que Controls periòdics).
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($e.ColumnIndex -ne 0 -or $e.RowIndex -lt 0) { return }
        $r = $s.Rows[$e.RowIndex].Tag
        if ($null -ne $r) { $r.Sel = [bool]$s.Rows[$e.RowIndex].Cells[0].Value }
    })
    $grid.add_CurrentCellDirtyStateChanged({
        param($s, $e)
        if ($s.IsCurrentCellDirty) { $s.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    })

    foreach ($c in @($chkActiu, $chkNomes)) { $c.add_CheckedChanged({ & $fn.Desa; & $fn.Pinta }.GetNewClosure()) }
    foreach ($n in @($numPer, $numEsp, $numMax)) { $n.add_ValueChanged({ & $fn.Desa; & $fn.Refresca }.GetNewClosure()) }
    # Un clic en un radio dispara DOS esdeveniments (el que es marca i el germà
    # que es desmarca): només reaccionem al que queda marcat.
    foreach ($rb in @($rbMan, $rbAut)) {
        $rb.add_CheckedChanged({ param($s, $e) if ($s.Checked) { & $fn.Desa } }.GetNewClosure())
    }

    $btnText.add_Click({ if (Invoke-RecordatorisTextos $clau) { $estat.campanyes[$clau] = (_RecLlegeix).campanyes[$clau] } }.GetNewClosure())
    $btnAuto.add_Click({ Invoke-RecordatorisTasca }.GetNewClosure())

    $btnExc.add_Click({
        if ($null -eq $grid.CurrentRow -or $null -eq $grid.CurrentRow.Tag) { return }
        $r = $grid.CurrentRow.Tag
        $nou = -not [bool]$r.Excloure
        $estat.historial[$clau] = _RecHistorialExclou $estat.historial[$clau] ([string]$r.Id) $nou
        _RecDesa $estat
        & $fn.Refresca
    }.GetNewClosure())

    $btnCsv.add_Click({
        $sel = @($ui.Rows)
        if ($sel.Count -eq 0) { return }
        $cache = $null
        try {
            $xls = Find-LatestActivitatsExcel
            if ($null -ne $xls) { $cache = Initialize-ActivitatsCache $xls.File }
        } catch { }
        $out = New-Object System.Collections.ArrayList
        foreach ($r in $sel) {
            $r2 = _RecOmpleDadesFila $r $cache
            [void]$out.Add([pscustomobject]@{
                'GIA' = $r2.Id; 'Titular' = $r2.Titular; 'Adreça' = $r2.Adreca
                'Estat' = $r2.Estat; 'Data informe' = $r2.DataInforme
                'Últim avís' = $r2.Ultim; 'Avisos' = $r2.Compte
                'Toca avui' = $(if ($r2.Toca) { 'SI' } else { 'NO' })
                'Situació' = $r2.Motiu; 'Correus' = $r2.Correus
            })
        }
        $dir = _ResolveOutputDir
        $path = _GetUniqueOutputPath $dir ('Recordatoris ' + $clau + ' ' + (Get-Date).ToString('yyyy-MM-dd') + '.csv')
        try {
            $out | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            $r = [System.Windows.Forms.MessageBox]::Show("CSV generat:`n$path`n`nVols obrir-lo?", 'Recordatoris', 'YesNo', 'Information')
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { try { Start-Process -FilePath $path | Out-Null } catch { } }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut escriure el CSV:`n$($_.Exception.Message)", 'Recordatoris', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $btnSend.add_Click({
        & $fn.Desa
        $tria = @(@($ui.Rows) | Where-Object { $_.Sel -and $_.Toca -and -not $_.Excloure })
        if ($tria.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No hi ha cap activitat marcada que toqui avui.', 'Recordatoris', 'OK', 'Information') | Out-Null
            return
        }
        $q = _QuotaLlegeix
        $rest = _QuotaRestant $q
        $prev = [Math]::Min($tria.Count, [Math]::Min([int]$cfg['maxPerTanda'], $rest))
        $msg = "S'enviaran fins a $prev correus (de $($tria.Count) marcats).`n`n" +
               "Topall per tanda: $($cfg['maxPerTanda'])`n" +
               "Quota d'aquest mes: $($q.enviats) / $($q.limit) (en queden $rest)`n`n" +
               'Vols continuar?'
        if ([System.Windows.Forms.MessageBox]::Show($msg, 'Enviar recordatoris', 'YesNo', 'Question') -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $res = Invoke-RecordatorisTanda $clau $tria $false
        $resum = "Enviats: $($res.Enviats)`nFallats: $($res.Fallats)`nSense correu: $($res.SenseCorreu)"
        if ($res.Aturat) { $resum += "`n`nATURAT: $($res.Motiu)" }
        $q2 = _QuotaLlegeix
        $resum += "`n`nQuota: $($q2.enviats) / $($q2.limit) aquest mes."
        [System.Windows.Forms.MessageBox]::Show($resum, 'Recordatoris', 'OK', 'Information') | Out-Null
        $estat.historial[$clau] = (_RecLlegeix).historial[$clau]
        & $fn.Refresca
    }.GetNewClosure())

    # Ordre de docking calcat de Controls periòdics: primer la graella (Fill),
    # després els panells de dalt i de baix.
    $tab.Controls.Add($grid)
    $tab.Controls.Add($top)
    $tab.Controls.Add($bot)
    & $fn.Refresca
}
