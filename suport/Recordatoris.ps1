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
        [pscustomobject]@{ Clau = 'requeriments'; Nom = 'Requeriments'; Estats = @($Script:EstatRequeriment) }
        [pscustomobject]@{ Clau = 'precintes';    Nom = 'Precintes';    Estats = @($Script:EstatPrecinte) }
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
# El que vol dir cada valor de la campanya (la (i) del costat). PURA.
function _RecAjudaCamp([string]$que) {
    $a = [char]0x00E0; $e = [char]0x00E8; $ee = [char]0x00E9; $o = [char]0x00F2; $oo = [char]0x00F3; $u = [char]0x00FA
    switch ($que) {
        'cada'   { return ("Cada quants dies es torna a enviar el recordatori a una mateixa activitat, mentre segueixi en aquest estat. Es compta des de l'" + $u + "ltim recordatori que se li va enviar.") }
        'espera' { return ("Quants dies han de passar des de la data de l'informe que ha deixat l'activitat en aquest estat abans d'enviar-li el PRIMER recordatori. Serveix perqu" + $e + " el titular tingui temps de complir el termini del requeriment.") }
        'max'    { return ("Quants correus s'envien com a m" + $a + "xim cada vegada (amb el bot" + $oo + " " + [char]0x00AB + "Enviar tanda" + [char]0x00BB + " o a la passada autom" + $a + "tica). La finestra ja en marca com a molt aquests, per ordre: primer les que no han rebut mai cap recordatori; les que no hi caben surten a la tanda seg" + [char]0x00FC + "ent.") }
        'mode'   { return ("Manual: nom" + $ee + "s s'envia quan cliques " + [char]0x00AB + "Enviar tanda" + [char]0x00BB + ".`nAutom" + $a + "tic: aquesta campanya entra a la passada autom" + $a + "tica, que fa la tasca del Windows. La tasca s'engega i s'atura amb l'interruptor A/M de la rajola Recordatoris del men" + $u + "; sense, cap campanya no s'envia sola.") }
    }
    return ''
}

function _RecDefaultConfig([string]$clau) {
    $tx = _RecDefaultTextos $clau
    # Sense 'actiu': la casella "Campanya activa" deia el mateix que el
    # Manual/Automatic i l'usuari no sabia quina manava (octubre 2026). Ara
    # nomes hi ha el mode; neix en MANUAL (no envia res sol).
    return @{
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
    $v = & $get $cfg 'mode';   if ("$v" -eq 'auto' -or "$v" -eq 'manual') { $out['mode'] = [string]$v }
    # Una configuracio d'abans amb la campanya APAGADA ('actiu' = false) i el
    # mode en Automatic no enviava res; ara, sense la casella, l'Automatic
    # enviaria. Es passa a Manual: cap campanya no pot comencar a escriure a
    # titulars per una actualitzacio del programa.
    $v = & $get $cfg 'actiu';  if ($null -ne $v -and -not [bool]$v) { $out['mode'] = 'manual' }
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
# Reutilitza _InformeQueDeterminaEstat (InformesClassificacio.ps1): l'estat i la data han de
# sortir del MATEIX informe, si no el recordatori diria una data que no lliga.
function _RecDataInforme($act) {
    if ($null -eq $act -or $null -eq $act.PSObject.Properties['informes']) { return '' }
    $inf = _InformeQueDeterminaEstat $act
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
        # Per com es TRACTA (el favorable pre-llicencia entra a Requeriments).
        if ($estats -notcontains (_EstatEquivalent $estat)) { continue }
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
    _RecPreselecciona $ord ([int]$cfg['maxPerTanda'])
    return @{ Files = $ord; SenseGia = $senseGia }
}

# Marca per enviar ('Sel') NOMES les 'max' primeres que toquen, en l'ordre de
# prioritat que ja porten. Abans es marcaven totes les que tocaven (155 amb un
# maxim de 15) i el topall nomes s'aplicava en enviar: la finestra deia una
# cosa i la tanda en feia una altra. PURA.
function _RecPreselecciona($files, [int]$max) {
    $n = 0
    foreach ($f in @($files)) {
        if ($null -eq $f) { continue }
        $f.Sel = ([bool]$f.Toca -and -not [bool]$f.Excloure -and $n -lt $max)
        if ($f.Sel) { $n++ }
    }
}

# Substitueix les variables del text amb les dades d'una fila. PURA.
# El MAPA es d'aqui; el bucle el fa _OmpleVariables (EnviarCorreu.ps1).
function _RecFillPh([string]$text, $row) {
    return (_OmpleVariables $text ([ordered]@{
        '{ID_GIA}'       = [string]$row.Id
        '{TITULAR}'      = [string]$row.Titular
        '{ADRECA}'       = [string]$row.Adreca
        '{ACTIVITAT}'    = [string]$row.Activitat
        '{DATA_INFORME}' = [string]$row.DataInforme
        '{DATA}'         = (Get-Date).ToString('dd/MM/yyyy')
    }))
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

# Historial d'una campanya: mapa GIA -> mapa d'entrada. PURA.
function _RecHistorialAMapa($o) {
    $out = @{}
    $top = ConvertTo-Mapa $o
    foreach ($k in @($top.Keys)) { $out[[string]$k] = ConvertTo-Mapa $top[$k] }
    return $out
}

# Llegeix configuració + historial del disc, ja normalitzats. Mai llança.
function _RecLlegeix {
    $out = @{ campanyes = @{}; historial = @{} }
    $raw = Read-JsonFile (_RecPath)
    $camps = @{}
    $hist  = @{}
    if ($null -ne $raw) {
        try { $camps = ConvertTo-Mapa $raw.campanyes } catch { }
        try { $hist  = ConvertTo-Mapa $raw.historial } catch { }
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

# L'HTML del cos el fa _CosAHtml (EnviarCorreu.ps1). Aqui hi havia _RecCosHtml,
# identica linia a linia a la dels controls periodics.

# Nom de la tasca programada del Windows (mode automàtic).
$Script:RecTascaNom = 'InformesCornella-Recordatoris'


# LA TASCA, EN XML. Abans es creava amb '/SC DAILY /ST 09:00', i amb aquesta
# línia d'ordres el schtasks no deixa dir "si s'ha saltat, fes-la quan puguis":
# amb el PC apagat a l'hora, aquell dia no s'enviava res. L'usuari (octubre
# 2026): a les 13:00, com tots els modes automàtics (ModeAutomatic.ps1), i si no
# s'ha arribat a fer l'última vegada que tocava, que es faci. Això és
# <StartWhenAvailable>, que només es pot posar amb /XML. PURA.
#   $hora 'HH:mm'; $avui: el dia de la primera vegada (StartBoundary).
# Els dies de la setmana (1 dilluns ... 7 diumenge) com els vol el Programador.
$Script:RecDiesXml = @('', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')

# $diaSetmana: 0 cada dia; 1..7 un cop a la setmana (la programacio de
# 'recordatoris', que es canvia a Configuracio; octubre 2026).
# $setmanes: 2 si la programacio es "cada dues setmanes" (la tasca del Windows
# ho sap fer sola: WeeksInterval).
function _RecTascaXml([string]$psExe, [string]$script, [string]$hora, [datetime]$avui, [int]$diaSetmana = 0, [int]$setmanes = 1) {
    $esc = { param($t) [System.Security.SecurityElement]::Escape([string]$t) }
    $args1 = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + [string]$script + '"'
    $quan = if ($diaSetmana -ge 1 -and $diaSetmana -le 7) {
        '<ScheduleByWeek><WeeksInterval>' + [math]::Max(1, $setmanes) + '</WeeksInterval><DaysOfWeek><' + $Script:RecDiesXml[$diaSetmana] + ' /></DaysOfWeek></ScheduleByWeek>'
    } else { '<ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>' }
    return ('<?xml version="1.0" encoding="UTF-16"?>' + "`r`n" +
        '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">' +
        '<RegistrationInfo><Description>Informes Cornella: recordatoris automatics</Description></RegistrationInfo>' +
        '<Triggers><CalendarTrigger><StartBoundary>' + $avui.ToString('yyyy-MM-dd') + 'T' + [string]$hora + ':00</StartBoundary>' +
        '<Enabled>true</Enabled>' + $quan + '</CalendarTrigger></Triggers>' +
        '<Principals><Principal id="Author"><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>' +
        '<Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>' +
        '<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>' +
        '<StartWhenAvailable>true</StartWhenAvailable><Enabled>true</Enabled><ExecutionTimeLimit>PT2H</ExecutionTimeLimit></Settings>' +
        '<Actions Context="Author"><Exec><Command>' + (& $esc $psExe) + '</Command><Arguments>' + (& $esc $args1) + '</Arguments></Exec></Actions>' +
        '</Task>')
}

# Arguments de schtasks per CREAR la tasca a partir de l'XML. PURA.
function _RecSchtasksArgv([string]$nom, [string]$xmlPath) {
    return @('/Create', '/TN', [string]$nom, '/XML', [string]$xmlPath, '/F')
}

# La tasca que hi ha al Windows ja és la d'ara (l'hora dels modes automàtics i
# StartWhenAvailable)? PURA: rep l'XML de 'schtasks /Query /XML'.
function _RecTascaAlDia([string]$xml, [string]$hora, [int]$diaSetmana = 0, [int]$setmanes = 1) {
    if ([string]::IsNullOrWhiteSpace($xml)) { return $false }
    # La sortida del schtasks pot arribar en UTF-16 llegida com a 8 bits: els
    # zeros de cada caracter fora (si no, mai quadraria i es refaria a cada
    # obertura).
    $xml = $xml.Replace([string][char]0, '')
    if ($xml -notmatch '<StartWhenAvailable>\s*true\s*</StartWhenAvailable>') { return $false }
    if ($xml -notmatch ('<StartBoundary>[^<]*T' + [regex]::Escape([string]$hora) + ':00')) { return $false }
    # I la frequencia: cada dia, o el dia de la setmana de la programacio.
    if ($diaSetmana -ge 1 -and $diaSetmana -le 7) {
        return ($xml -match '<ScheduleByWeek>' -and $xml -match ('<' + $Script:RecDiesXml[$diaSetmana] + '\s*/>') -and
                $xml -match ('<WeeksInterval>\s*' + [math]::Max(1, $setmanes) + '\s*</WeeksInterval>'))
    }
    return ($xml -match '<ScheduleByDay>')
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
    $p = Get-InformesDbPath
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
# $cfgE: la configuracio del correu de la campanya (Get-CorreuEina); a qui va
# es tria a Configuracio -> Correus de cada eina. Sense, el titular i el
# representant de l'Excel (el d'abans).
function _RecOmpleDadesFila($row, $cache, $cfgE = $null) {
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
    if ($null -ne $cfgE) {
        $row.Correus = (@(_CorreuDestinataris $cfgE @{ titular = $rao; representant = $rep } (Get-CorreuAutoritzats ([string]$row.Id)))) -join '; '
        return $row
    }
    $d = _CorreuDestinatarisPerDefecte $rao $rep
    $row.Correus = [string]$d.Text
    return $row
}

# ----------------------------------------------------------------------------
# ENVIAMENT D'UNA TANDA (la comparteixen el mode manual i l'automatic)
# ----------------------------------------------------------------------------
# $rows: files ja triades. $silenci: mode automatic (cap finestra).
# Retorna @{ Enviats; Esborranys; Fallats; SenseCorreu; Aturat; Motiu; Via }.
function Invoke-RecordatorisTanda([string]$clau, $rows, [bool]$silenci) {
    # PER ON i A QUI: la de la campanya a Configuracio -> Correus de cada eina
    # (CorreuEines.ps1), tambe en automatic (la tasca del Windows corre amb la
    # sessio iniciada, que es el que l'Outlook necessita).
    $cfgE = Get-CorreuEina ('rec-' + $clau)
    $via = [string]$cfgE.Via
    $res = @{ Enviats = 0; Esborranys = 0; Fallats = 0; SenseCorreu = 0; Aturat = $false; Motiu = ''; Via = $via }
    $desats = New-Object System.Collections.ArrayList
    $files = @($rows)
    if ($files.Count -eq 0) { $res.Motiu = 'cap activitat'; return $res }

    $cfgAll = _RecLlegeix
    $cfg = $cfgAll.campanyes[$clau]
    if ($null -eq $cfg) { $res.Motiu = 'campanya desconeguda'; return $res }

    # Claus d'EmailJS: si en falta cap, val mes dir-ho que provar-ho N vegades.
    # La comprovacio es la MATEIXA que mira l'interruptor de la rajola
    # (Test-CorreuLlest, CorreuVia.ps1): si fossin dues, un automatic es
    # podria encendre amb unes claus que l'enviament no accepta.
    $motiuCorreu = Test-CorreuViaLlest $via
    if ($motiuCorreu -ne '') { $res.Motiu = $motiuCorreu; $res.Aturat = $true; return $res }

    # Quota (nomes la d'EmailJS): el topall mana per sobre del maxPerTanda.
    $quota = _QuotaLlegeix
    $restant = if (_CorreuViaEsOutlook $via) { [int]::MaxValue } else { _QuotaRestant $quota }
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
    if ($null -ne $cfgE.Cco) { $bcc = [string]$cfgE.Cco }
    else { try { $bcc = (@($cfg['bcc']) -join ',') } catch { } }

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
        [void](_AddPeuBotons $form @(@{ Nom = 'Cancel'; Text = 'Cancel·lar'; Clic = { $st.Cancel = $true }.GetNewClosure() }) @() 90)
        $form.Show(); [System.Windows.Forms.Application]::DoEvents()
    }

    $ses = $null
    try {
        try { $ses = Open-CorreuSessio $via } catch { $res.Motiu = [string]$_.Exception.Message; $res.Aturat = $true; return $res }
        $n = 0
        foreach ($row in $files) {
            if ($n -ge $limitAra) { break }
            if ($st.Cancel) { $res.Aturat = $true; $res.Motiu = "cancel·lat per l'usuari"; break }

            $row = _RecOmpleDadesFila $row $cache $cfgE
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
            $html     = _CosAHtml $cos
            $to       = (([string]$row.Correus) -split '\s*;\s*' | Where-Object { $_ }) -join ','

            # El try/catch va DINS del bucle: un error d'una activitat no pot
            # endur-se la resta de la tanda.
            try {
                Send-CorreuSessio $ses $to $bcc $assumpte $html
                $n++
                # A Esborranys TAMBE va a l'historial: si no, la tasca de l'endema
                # en tornaria a desar un altre per al mateix titular. Que s'enviin
                # ho vigila l'avis dels esborranys pendents (CorreuVia.ps1).
                $esb = ($via -eq 'outlook-esborrany')
                if ($esb) { $res.Esborranys++; [void]$desats.Add("GIA $($row.Id) - $($row.Titular) ($to)") } else { $res.Enviats++ }
                # Es desa DESPRES DE CADA enviament: si peta o es cancel.la, el
                # que ja ha sortit consta i no es tornara a enviar. (La quota
                # d'EmailJS ja l'apunta Send-EmailJs: abans s'apuntava tambe
                # aqui i cada recordatori comptava doble.)
                $avuiIso = (Get-Date).ToString('yyyy-MM-dd')
                $cfgAll.historial[$clau] = _RecHistorialActualitza $cfgAll.historial[$clau] ([string]$row.Id) $avuiIso
                _RecDesa $cfgAll
                _RecLog ("GIA $($row.Id): " + $(if ($esb) { "desat a Esborranys de l'Outlook" } else { 'enviat' }) + " per a $to ($via)")
                if ($n -lt $limitAra -and -not (_CorreuViaEsOutlook $via)) { Start-Sleep -Milliseconds $Script:RecPausaMs }
            } catch {
                $res.Fallats++
                $txt = _CorreuSessioError $ses $_
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
        Close-CorreuSessio $ses
        if ($null -ne $form) { try { $form.Close() } catch { } }
        # Els de la tasca automatica no els veu ningu: el menu ho avisara.
        if ($silenci -and $desats.Count -gt 0) { Add-CorreuEsborranysPendents ('Recordatoris (' + [string](_RecCampanyaPerClau $clau).Nom + ')') $desats }
    }
    # QUI HA FET AQUESTA TANDA, per a la data de sota la rajola: verda si la va
    # fer l'automatic, grisa si la vas fer tu. Mateix criteri que Copiar
    # informes i Actualitzar base. Nomes s'apunta si s'ha enviat alguna cosa:
    # una tanda que no tenia res a enviar no canvia qui va ser l'ultim.
    if ([int]$res.Enviats + [int]$res.Esborranys -gt 0) {
        [void](_RecAutoDesaEstat @{ mode = $(if ($silenci) { 'auto' } else { 'manual' }); enviat_el = (Get-Date).ToString('o') })
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

    return (Show-EditorAssumpteCos `
        -TextFinestra ('Text del recordatori - ' + $camp.Nom) `
        -Titol ('Recordatoris - ' + $camp.Nom) `
        -Subtitol ('Text que rebr' + [char]0x00E0 + ' el titular') `
        -Ajuda (_RecAjuda) `
        -Assumpte ([string]$cfg['assumpte']) `
        -Cos ([string]$cfg['cos']) `
        -EtiquetaRestaurar 'Restaurar el text per defecte' `
        -Restaurar {
            $r = [System.Windows.Forms.MessageBox]::Show(
                'Vols recuperar el text per defecte? Es perdr' + [char]0x00E0 + ' el que hi ha ara.',
                'Recordatoris', 'YesNo', 'Question')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $null }
            return (_RecDefaultTextos $clau)
        } `
        -Desa {
            param($v)
            # Es torna a LLEGIR l'estat abans de desar: la finestra ha estat
            # oberta i el que hi ha al disc pot no ser el que es va llegir en
            # obrir-la (l'enviament automatic hi escriu l'historial).
            $ara = _RecLlegeix
            $ara.campanyes[$clau]['assumpte'] = [string]$v['assumpte']
            $ara.campanyes[$clau]['cos']      = [string]$v['cos']
            try {
                _RecDesa $ara
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Recordatoris', 'OK', 'Error') | Out-Null
                return $false
            }
        })
}

# ----------------------------------------------------------------------------
# TASCA PROGRAMADA DEL WINDOWS (mode automàtic)
# ----------------------------------------------------------------------------
function _RecScriptAuto {
    return [string](Join-Path $PSScriptRoot 'RecordatorisAuto.ps1')
}

# Crea (o reescriu, /F) la tasca a l'hora dels modes automàtics. Retorna el
# resultat de schtasks (@{ Codi; Sortida }). L'XML va en UTF-16, que és el que
# el schtasks espera quan la capçalera ho diu.
function _RecCreaTasca {
    $psExe = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
    $pr = Get-ProgramacioAuto 'recordatoris'
    $x = _ProgramacioParts $pr
    $xml = _RecTascaXml $psExe (_RecScriptAuto) ([string]$pr.Hora) (Get-Date) $x[2] $x[3]
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('recordatoris-tasca-' + [guid]::NewGuid().ToString('N') + '.xml')
    try {
        [System.IO.File]::WriteAllText($tmp, $xml, [System.Text.Encoding]::Unicode)
        return (_RecExecutaSchtasks (_RecSchtasksArgv $Script:RecTascaNom $tmp))
    } finally { try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch { } }
}

# EN OBRIR EL PROGRAMA: si la tasca existeix pero és d'abans (les 09:00, sense
# recuperar la passada perduda), es reescriu sola. Si no existeix no es crea:
# això ho decideix l'usuari (botó "Automàtic..."). Mai llança, i una sola vegada
# per execució del programa.
$Script:RecTascaRevisada = $false

# La programacio dels recordatoris: la mateixa llista que la resta
# d'automatismes, i es canvia a Configuracio.
Register-ProgramacioAuto 'recordatoris' 'Recordatoris'

# ----------------------------------------------------------------------------
# L'INTERRUPTOR A/M DE LA RAJOLA (el mateix criteri que les altres tres eines)
# ----------------------------------------------------------------------------
# L'usuari (octubre 2026): "l'eina Recordatoris, que te un estat manual i un
# automatic, aplica el mateix criteri que la resta (Planol activitats,
# Actualitzar base i Copiar informes)".
#
# Fins ara l'automatic s'encenia des d'un boto "Automatic..." DINS de l'eina,
# amb un quadre de Si/No/Cancel.lar. Les altres tres fa temps que es commuten
# des de la rajola del menu, i aquesta era l'unica que no: la mateixa cosa amb
# dues interficies.
#
# EL QUE NO CANVIA, I ES PER QUE AQUESTA EINA ES DIFERENT: les altres tres
# corren EN OBRIR EL PROGRAMA (el SiToca del rellotge del menu). Els
# recordatoris han de sortir encara que no l'obris en setmanes, o sigui que
# l'automatic es una TASCA DEL WINDOWS. Per aixo el seu SiToca no envia res
# -nomes manté la tasca al dia si la programacio ha canviat- i encendre
# l'interruptor vol dir CREAR la tasca.
$Script:RecAutoPlantilla = [ordered]@{ generat_el = ''; mode = ''; enviat_el = '' }

function _RecAutoStatePath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'recordatoris-auto.json'))
}
function _RecAutoEstat { return (Read-EstatAuto (_RecAutoStatePath) $Script:RecAutoPlantilla) }
function _RecAutoDesaEstat($canvis) { return (Save-EstatAuto (_RecAutoStatePath) $canvis $Script:RecAutoPlantilla) }

# 'auto' | 'manual' | '': qui va fer l'ultim enviament. El menu hi pinta la data
# en verd o en gris, com a les altres tres.
function _RecUltimMode { return [string](_RecAutoEstat)['mode'] }

# HI ES, LA TASCA? Es consulta UNA vegada i es recorda: Actiu el crida el
# rellotge del menu cada minut, i engegar un schtasks per minut nomes per pintar
# una pastilla no te cap sentit. Qui la crea o l'esborra refresca la memoria.
$Script:RecTascaHiEs = $null
function _RecTascaActiva {
    if ($null -eq $Script:RecTascaHiEs) {
        if ([string]::IsNullOrWhiteSpace($env:SystemRoot)) { return $false }   # no es un Windows
        $q = _RecExecutaSchtasks @('/Query', '/TN', $Script:RecTascaNom)
        $Script:RecTascaHiEs = ($q.Codi -eq 0)
    }
    return [bool]$Script:RecTascaHiEs
}

# Encendre = crear la tasca; apagar = esborrar-la. Torna $true si ha anat be
# (el menu nomes mou l'interruptor si el canvi s'ha pogut desar).
function _RecTascaDesaActiu([bool]$on) {
    if ($on) {
        $script = _RecScriptAuto
        if (-not (Test-Path -LiteralPath $script)) { return $false }
        # L'AVIS ES QUEDA, i es l'unica cosa que aquest interruptor fa diferent
        # dels altres tres: encendre'l vol dir que sortiran correus a titulars
        # SENSE que ningu els miri. Les altres eines automatiques copien fitxers
        # o refan una base; aquesta escriu a gent. Si l'usuari es fa enrere no
        # es toca res, i el menu, que torna a llegir Actiu, deixa l'interruptor
        # on era.
        $msg = "Vols engegar l'enviament AUTOM" + [char]0x00C0 + "TIC dels recordatoris?`n`n" +
               "Es crear" + [char]0x00E0 + " una tasca del Windows que " + (Get-ProgramacioText 'recordatoris') + " enviar" + [char]0x00E0 + " els`n" +
               "recordatoris de les campanyes que tinguis en mode Autom" + [char]0x00E0 + "tic. Si a`n" +
               "aquella hora el PC estava apagat, s'enviaran en engegar-lo.`n`n" +
               "Tingues en compte que:`n" +
               " " + [char]0x00B7 + " Nom" + [char]0x00E9 + "s s'executa amb el PC engegat i la sessi" + [char]0x00F3 + " iniciada.`n" +
               " " + [char]0x00B7 + " Cada campanya surt per la via que tingui a Configuraci" + [char]0x00F3 + " (Correus de cada eina): " + ((@(_RecCampanyes) | ForEach-Object { [string]$_.Nom + ' = ' + (_CorreuViaText (Get-CorreuEina ('rec-' + $_.Clau)).Via) }) -join '; ') + ". Els d'Esborranys, el programa t'avisa que els has d'enviar; els altres surten SENSE que ning" + [char]0x00FA + " els revisi.`n" +
               " " + [char]0x00B7 + " Si la base d'informes t" + [char]0x00E9 + " m" + [char]0x00E9 + "s de " + [string]$Script:RecMaxAntiguitatDbDies + " dies, no enviar" + [char]0x00E0 + " res."
        $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Recordatoris autom' + [char]0x00E0 + 'tics', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
        $res = _RecCreaTasca
    } else {
        $res = _RecExecutaSchtasks @('/Delete', '/TN', $Script:RecTascaNom, '/F')
    }
    if ($res.Codi -ne 0) { $Script:RecTascaHiEs = $null; return $false }   # torna-ho a mirar
    $Script:RecTascaHiEs = $on
    return $true
}

$Script:ModesAuto['recordatoris'] = @{
    Titol     = 'Recordatoris'
    Actiu     = { _RecTascaActiva }
    DesaActiu = { param($on) _RecTascaDesaActiu $on }
    UltimMode = { _RecUltimMode }
    # NO ENVIA RES: els recordatoris els envia la tasca del Windows, tambe amb
    # el programa tancat. Aqui nomes es manté la tasca al dia si la programacio
    # ha canviat a Configuracio.
    SiToca    = { Update-RecordatorisTascaSiCal }
    # Sense les claus d'EmailJS no es pot enviar cap correu, i deixar-ho ences
    # seria un automatic que no fa res i no ho diu (mateix criteri que la
    # carpeta de Copiar informes).
    Requisit  = {
        # La via de CADA campanya (Configuracio -> Correus de cada eina).
        $m = @(@(_RecCampanyes) | ForEach-Object { Test-CorreuViaLlest ([string](Get-CorreuEina ('rec-' + $_.Clau)).Via) } | Where-Object { $_ }) | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace([string]$m)) { return '' }
        return ("Per enviar els recordatoris sols, " + $m + ".`n`nVegeu suport\documentacio\DESPLEGAMENT-MOBIL.md.")
    }
    TipA      = { Get-AutoTipText 's envien sols' 'recordatoris' }
    TipM      = "Mode MANUAL: nomes s'envien quan obres l'eina i ho demanes. Clica per posar-ho en automatic."
}

# -Forca: torna-ho a mirar encara que ja s'hagi fet (Configuracio, en desar
# una programacio nova).
function Update-RecordatorisTascaSiCal([switch]$Forca) {
    if ($Script:RecTascaRevisada -and -not $Forca) { return }
    $Script:RecTascaRevisada = $true
    try {
        if ([string]::IsNullOrWhiteSpace($env:SystemRoot)) { return }   # no és un Windows
        $q = _RecExecutaSchtasks @('/Query', '/TN', $Script:RecTascaNom, '/XML')
        if ($q.Codi -ne 0) { return }                                   # no hi és: res a fer
        $pr = Get-ProgramacioAuto 'recordatoris'
        $x = _ProgramacioParts $pr
        if (_RecTascaAlDia ([string]$q.Sortida) ([string]$pr.Hora) $x[2] $x[3]) { return }
        $c = _RecCreaTasca
        if ($c.Codi -eq 0) { _RecLog ('Tasca programada actualitzada: ' + (Get-ProgramacioText $pr) + ' i, si es salta, en quant es pugui.') }
        else { _RecLog ("No s'ha pogut actualitzar la tasca programada: " + ([string]$c.Sortida).Trim()) }
    } catch { }
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

# ----------------------------------------------------------------------------
# EL CORREU DE PROVA (Configuracio -> Correus de cada eina)
# ----------------------------------------------------------------------------
# El recordatori de cada campanya amb les dades de l'activitat de prova: el
# text i la CCO de la campanya, i la data de l'informe de la base si hi es.
foreach ($recCampP in @(_RecCampanyes)) {
    $recClauP = [string]$recCampP.Clau
    $Script:CorreuProves['rec-' + $recClauP] = {
        param($gia, $cache, $cfgE)
        $act = Get-ActivitatFromCache $cache $gia
        if ($null -eq $act) { throw "L'ID GIA de prova ($gia) no es a l'Excel d'activitats." }
        $cfg = (_RecLlegeix).campanyes[$recClauP]
        $dataInf = (Get-Date).ToString('yyyy-MM-dd')
        $db = _RecCarregaDb
        if ($null -ne $db) {
            $a = @(@($db.activitats) | Where-Object { [string]$_.id_gia -eq $gia }) | Select-Object -First 1
            $d = if ($null -ne $a) { _RecDataInforme $a } else { '' }
            if ($d) { $dataInf = $d }
        }
        $tit = ''; try { if ($act.ContainsKey('TITULAR')) { $tit = [string]$act['TITULAR'] } } catch { }
        $row = [pscustomobject]@{ Id = $gia; Titular = $tit; DataInforme = $dataInf; Adreca = ''; Correus = '' }
        $row = _RecOmpleDadesFila $row $cache $cfgE
        return @{
            Assumpte = (_RecFillPh ([string]$cfg['assumpte']) $row)
            Html = (_CosAHtml (_RecFillPh ([string]$cfg['cos']) $row))
            Destinataris = @(([string]$row.Correus) -split '\s*;\s*' | Where-Object { $_ })
            CcoAbans = (@($cfg['bcc']) -join '; ')
        }
    }.GetNewClosure()
    $Script:CorreuCcoAbans['rec-' + $recClauP] = { (@((_RecLlegeix).campanyes[$recClauP]['bcc']) -join '; ') }.GetNewClosure()
}
