#requires -Version 5.1
<#
.SYNOPSIS
  Avisos de CONTROL PERIODIC per correu (esborranys a Outlook) des de l'eina
  "Controls periodics".

.DESCRIPTION
  Per a les activitats SELECCIONADES a la graella de Controls periodics, crea un
  correu per titular avisant que constava un control periodic a passar (en data
  X) per la seva activitat X situada a X. Els correus i les dades surten de la
  base d'activitats (Excel) que ja llegeix _ReadControlsPeriodics; el programa
  els deixa com a ESBORRANYS a Outlook (mai els envia sol) perque l'usuari els
  revisi i envii.

  El text es EDITABLE (assumpte + cos amb variables), com els "Textos del correu"
  del mobil, pero es guarda LOCALMENT (%LOCALAPPDATA%\InformesCornella\
  controls-cp-email.json), fora del repositori: no conte cap dada personal i
  sobreviu a Actualitzar.bat (git pull) sense conflictes.

  Variables del text:
    {ACTIVITAT} {ADRECA} {ID_GIA} {TITULAR}   dades de l'activitat
    {PROPER_CP}      Proper CP previst (la data en que tocava el control)
    {DATA_CONTROL}   Data de l'ultim control periodic registrat
    {DATA}           data d'avui (dd/MM/yyyy)
  El cos admet **negreta** i els enllacos http(s) es fan clicables.

  Funcions PURES (plantilla, destinataris, substitucio de variables, HTML)
  testejables en headless; Outlook (COM) i les finestres (WinForms) nomes a
  Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# Ruta del fitxer d'overrides local (mai al repositori; sense dades personals).
function _ControlsCpEmailPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path $base (Join-Path 'InformesCornella' 'controls-cp-email.json'))
}

# Valors PER DEFECTE (assumpte + cos, bilingue CA/ES, cordial).
function _DefaultControlsCpEmail {
    $cos = @(
        'Activitat: {ACTIVITAT}'
        'Adreça: {ADRECA}'
        'ID GIA: {ID_GIA}'
        'Titular: {TITULAR}'
        ''
        '**Català**'
        "Benvolgut/da,"
        "Segons les nostres dades, la vostra activitat situada a {ADRECA} havia de passar un **control periòdic** amb data prevista **{PROPER_CP}**, i a hores d'ara no ens consta que s'hagi dut a terme."
        "Us recordem que aquest control és una obligació periòdica de l'activitat. Us demanem que, com abans millor, encarregueu el control a una entitat/organisme de control habilitat i que ens en presenteu l'acta favorable."
        "Podeu presentar la documentació mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat, a l'atenció del **Departament d'Activitats**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        '**Castellano**'
        "Estimado/a,"
        "Según nuestros datos, su actividad situada en {ADRECA} debía pasar un **control periódico** con fecha prevista **{PROPER_CP}**, y a día de hoy no nos consta que se haya realizado."
        "Le recordamos que este control es una obligación periódica de la actividad. Le pedimos que, cuanto antes, encargue el control a una entidad/organismo de control habilitado y nos presente el acta favorable."
        "Puede presentar la documentación mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat, a la atención del **Departamento de Actividades**:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        ''
        '________________________________________'
        ''
        "Departament d'Activitats · Ajuntament de Cornellà de Llobregat · Carrer de l'Energia, 97 · Tel. 93 377 02 12 (ext. 1227)"
        "IMPORTANT: aquest és un correu automàtic. Per a qualsevol consulta, adreceu-vos al Departament d'Activitats. / IMPORTANTE: este es un correo automático. Para cualquier consulta, diríjase al Departamento de Actividades."
    ) -join "`n"
    $d = [ordered]@{
        assumpte = 'Control periòdic pendent · GIA {ID_GIA}'
        cos      = $cos
    }
    return ,$d
}

# Text d'ajuda amb les variables disponibles.
function _ControlsCpEmailAjuda {
    return ('Variables: {ACTIVITAT} {ADRECA} {ID_GIA} {TITULAR} {PROPER_CP} {DATA_CONTROL} {DATA}   ' + [char]0x00B7 + '   **negreta**   ' + [char]0x00B7 + '   //cursiva//   ' + [char]0x00B7 + '   els enllaços http es fan clicables')
}

# Llegeix el JSON local (si hi es) fusionat sobre els valors per defecte.
function _LoadControlsCpEmail {
    $def = _DefaultControlsCpEmail
    $path = _ControlsCpEmailPath
    $o = Read-JsonFile $path
    if ($null -ne $o) {
        try {
            foreach ($k in @($def.Keys)) {
                if ($o.PSObject.Properties[$k] -and -not [string]::IsNullOrEmpty([string]$o.$k)) {
                    $def[$k] = [string]$o.$k
                }
            }
        } catch { }
    }
    return ,$def
}

# Escriu el JSON local (UTF-8 sense BOM), creant la carpeta si cal.
function _SaveControlsCpEmail($obj) {
    $path = _ControlsCpEmailPath
    Write-JsonFile $path $obj 5
}

# (Els destinataris els munta ara _CorreuDestinataris + _CorreuParteixToCc,
# CorreuEines.ps1: a qui va es tria a Configuracio -> Correus de cada eina. El
# primer a "Per a" i la resta a "CC", com feia _ControlsCpRecipients amb el
# titular i el representant.)


# Substitueix les variables del text amb les dades d'una fila d'activitat.
# El MAPA es d'aqui (les claus i d'on surten els valors son d'aquesta eina); el
# bucle el fa _OmpleVariables (EnviarCorreu.ps1), que estava copiat tres cops.
function _FillControlsCpPh([string]$text, $row) {
    return (_OmpleVariables $text ([ordered]@{
        '{ACTIVITAT}'    = [string]$row.ActPrincipal
        '{ADRECA}'       = [string]$row.Adreca
        '{ID_GIA}'       = [string]$row.Id
        '{TITULAR}'      = [string]$row.RaoSocial
        '{PROPER_CP}'    = [string]$row.ProperCP
        '{DATA_CONTROL}' = [string]$row.DataControlPer
        '{DATA}'         = (Get-Date).ToString('dd/MM/yyyy')
    }))
}

# L'HTML del cos el fa _CosAHtml (EnviarCorreu.ps1). Aqui hi havia
# _ControlsCpEmailHtml, IDENTICA linia a linia a _RecCosHtml dels recordatoris
# -mateix estil inline inclos-, i un _ControlsCpLineHtml propi que feia el mateix
# que _TextToHtml pero SENSE cursiva. Ara aquest correu tambe accepta //cursiva//.

# ----------------------------------------------------------------------------
# L'HISTORIAL DELS AVISOS (octubre 2026, l'automatic: "aplica l'automatitzacio
# de les eines a Controls periodics"). Per GIA, el "Proper CP previst" del qual
# ja s'ha preparat l'esborrany: l'automatic no el torna a preparar cada setmana,
# i si l'Excel porta una data nova (s'ha passat el control i n'hi ha un altre de
# previst), torna a tocar. Tambe hi apunta el boto "Enviar correu (esborranys)":
# el que has fet a ma, l'automatic ja no ho repeteix. Porta ID GIA i dates:
# local\base-dades-activitats (aquest PC), mai al GitHub.
#   { avisats: { "<GIA>": { proper_cp: "dd/MM/yyyy"; el: iso; mode: auto|manual } } }
# ----------------------------------------------------------------------------
function _CpAvisatsPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'controls-periodics-avisats.json')
}

# GIA -> @{ ProperCp; El; Mode }. Mai peta (sense fitxer, buit).
function Read-CpAvisats([string]$path = '') {
    if (-not $path) { $path = _CpAvisatsPath }
    $out = @{}
    if (-not $path) { return $out }
    $o = $null
    try { $o = Read-JsonFile $path } catch { $o = $null }
    $a = _PropInf $o 'avisats'
    if ($null -eq $a) { return $out }
    foreach ($p in @($a.PSObject.Properties)) {
        $v = $p.Value
        $out[[string]$p.Name] = @{ ProperCp = [string](_PropInf $v 'proper_cp'); El = [string](Read-JsonIso (_PropInf $v 'el')); Mode = [string](_PropInf $v 'mode') }
    }
    return $out
}

# Hi afegeix (o hi posa al dia) les files avisades i ho desa. Torna $true si
# s'ha pogut desar (una unitat de xarxa fora de servei no pot fer petar res).
function Add-CpAvisats($rows, [string]$mode, [string]$path = '') {
    if (-not $path) { $path = _CpAvisatsPath }
    if (-not $path -or @($rows).Count -eq 0) { return $false }
    try {
        $tots = Read-CpAvisats $path
        $ara = (Get-Date).ToString('o')
        foreach ($r in @($rows)) { $tots[[string]$r.Id] = @{ ProperCp = [string]$r.ProperCP; El = $ara; Mode = $mode } }
        $obj = [ordered]@{}
        foreach ($k in @($tots.Keys | Sort-Object { _GiaNumeric $_ })) {
            $obj[$k] = [ordered]@{ proper_cp = [string]$tots[$k].ProperCp; el = [string]$tots[$k].El; mode = [string]$tots[$k].Mode }
        }
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile $path ([ordered]@{ avisats = $obj }) 5
        return $true
    } catch { return $false }
}

# L'avis d'aquesta fila ja es fet? Nomes si es del MATEIX "Proper CP previst":
# amb una data nova, torna a tocar. PURA.
function _CpJaAvisat($row, $avisats) {
    if ($null -eq $avisats) { return $false }
    $a = $avisats[[string]$row.Id]
    if ($null -eq $a) { return $false }
    return ([string]$a.ProperCp -ne '' -and [string]$a.ProperCp -eq [string]$row.ProperCP)
}

# ----------------------------------------------------------------------------
# ELS ESBORRANYS, SENSE FINESTRES: el boto (amb la barra) i l'automatic (sense
# ningu davant) passen pel mateix lloc. Abans tot vivia dins del boto, i fer-ne
# una copia per a l'automatic era la manera segura que divergissin.
#   $cfgE   Get-CorreuEina 'controls' (a qui va)
#   $bcc    la CCO ja muntada (l'automatic, NOMES la de "sempre")
#   $mode   'manual' | 'auto' (a l'historial)
#   $cada   { param($i, $row) } abans de cada correu; si torna $true, s'atura
# Torna @{ Fets; SenseCorreu; Errors; AvisTecnic; DelsDocs; Cancelat }. Si
# l'Outlook no s'obre, LLANCA (el missatge ja diu que passa).
# ----------------------------------------------------------------------------
function _CpPreparaEsborranys($rows, $cfgE, [string]$bcc, [string]$mode, [scriptblock]$cada = $null) {
    $res = @{
        Fets = (New-Object System.Collections.ArrayList); SenseCorreu = (New-Object System.Collections.ArrayList)
        Errors = (New-Object System.Collections.ArrayList); AvisTecnic = (New-Object System.Collections.ArrayList)
        DelsDocs = (New-Object System.Collections.ArrayList); Cancelat = $false
    }
    $textos = _LoadControlsCpEmail
    $assTpl = [string]$textos['assumpte']
    $cosTpl = [string]$textos['cos']
    # La mateixa sessio que Enviar correu i Recordatoris (CorreuVia.ps1): un sol
    # lloc obre l'Outlook i hi posa el remitent de Configuracio. Aqui SEMPRE a
    # Esborranys, triis la via que triis: l'usuari els revisa.
    $ses = Open-CorreuSessio 'outlook-esborrany'
    try {
        $i = 0
        foreach ($r in @($rows)) {
            $i++
            if ($null -ne $cada -and [bool](& $cada $i $r)) { $res.Cancelat = $true; break }
            $etiq = "GIA $($r.Id) - $($r.RaoSocial)"
            # El repas de contactes: el correu dels documents si l'Excel no en
            # te, i la llista dels que son del tecnic (es diu al final: son
            # esborranys, l'usuari els revisa abans d'enviar-los).
            $ctEm = Get-ContactesEmailsCompletats ([string]$r.Id) @{ titular = [string]$r.RaoEmail; representant = [string]$r.RepEmail }
            if ([string]$ctEm.AvisTecnic -ne '') { [void]$res.AvisTecnic.Add($etiq) }
            if (@($ctEm.Notes).Count -gt 0) { [void]$res.DelsDocs.Add($etiq) }
            $rec = _CorreuParteixToCc (_CorreuDestinataris $cfgE @{ titular = [string]$ctEm.Emails.titular; representant = [string]$ctEm.Emails.representant } (Get-CorreuAutoritzats ([string]$r.Id)))
            if (-not $rec.Ok) { [void]$res.SenseCorreu.Add($etiq); continue }
            try {
                Send-CorreuSessio $ses $rec.To $bcc (_FillControlsCpPh $assTpl $r) (_CosAHtml (_FillControlsCpPh $cosTpl $r)) $rec.Cc
                [void]$res.Fets.Add($r)
            } catch {
                [void]$res.Errors.Add("GIA $($r.Id): $($_.Exception.Message)")
            }
        }
    } finally {
        # NO es fa Quit() de l'Outlook (podria tancar el de l'usuari); nomes s'allibera.
        Close-CorreuSessio $ses
        # A l'historial ENCARA QUE S'HAGI ATURAT A MIG CAMI: el que ja es a
        # Esborranys no s'ha de tornar a preparar.
        [void](Add-CpAvisats @($res.Fets) $mode)
    }
    return $res
}

# ----------------------------------------------------------------------------
# OUTLOOK (COM) - nomes a Windows. Crea ESBORRANYS; MAI envia.
# ----------------------------------------------------------------------------
function Invoke-ControlsCpEmailDrafts($rows) {
    $rows = @($rows)
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Marca almenys una activitat (columna 'Generar') per preparar-ne el correu.", 'Enviar correu', 'OK', 'Information') | Out-Null
        return
    }

    # A QUI VA i la CCO: Configuracio -> Correus de cada eina (CorreuEines.ps1).
    $cfgE = Get-CorreuEina 'controls'
    $bcc = if ($null -ne $cfgE.Cco) { [string]$cfgE.Cco } else { '' }
    # La CCO "a triar" (Correus de cada eina): abans de confirmar, que es vegi.
    $ccoT = Select-CorreuCcoTria 'controls' $cfgE
    if (-not $ccoT.Ok) { return }
    $bcc = _CorreuCcoAmb $bcc @($ccoT.Adreces)
    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("Es prepararan $($rows.Count) correus (un per activitat triada) com a ESBORRANYS a Outlook.`n`n" +
         "A qui: " + (_CorreuDestText $cfgE) + " (el primer a 'Per a' i la resta a 'CC')" + $(if ($bcc) { "`nCCO: $bcc" } else { '' }) +
         "`n`nNO s'envia res: els revisaràs i enviaràs tu des d'Outlook.`n`nVols continuar?"),
        'Enviar correu', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # La carcassa (finestra + etiqueta + barra + Cancel.lar) es a UiComuns.ps1:
    # era identica a la de "Generar informes".
    $pg     = Show-ProgresCancel 'Enviar correu' $rows.Count
    $cancel = $pg.Cancel
    $form   = $pg.Form
    $lbl    = $pg.Label
    $bar    = $pg.Bar
    $res = $null
    try {
        $res = _CpPreparaEsborranys $rows $cfgE $bcc 'manual' {
            param($i, $r)
            $lbl.Text = "Preparant esborranys...  $i de $($rows.Count)`nGIA $($r.Id) - $($r.RaoSocial)"
            if ($bar.Value -lt $bar.Maximum) { $bar.Value = $i }
            [System.Windows.Forms.Application]::DoEvents()
            return [bool]$cancel.Flag
        }.GetNewClosure()
    } catch {
        $cancel.Running = $false
        try { $form.Close() } catch { }
        [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Enviar correu', 'OK', 'Error') | Out-Null
        return
    } finally {
        $cancel.Running = $false
        try { $form.Close() } catch { }
    }
    $ok = $res.Fets.Count; $cancelled = [bool]$res.Cancelat
    $senseCorreu = $res.SenseCorreu; $err = $res.Errors.Count; $errDetalls = $res.Errors
    $avisTec = $res.AvisTecnic; $delsDocs = $res.DelsDocs
    $avui = (Get-Date).ToString('dd/MM/yyyy')
    foreach ($f in @($res.Fets)) { if ($f.PSObject.Properties['Avisat']) { $f.Avisat = $avui } }

    $titol = if ($cancelled) { 'Preparació cancel·lada' } else { 'Esborranys preparats' }
    $msg = "$titol`n`nEsborranys creats a Outlook: $ok"
    if ($senseCorreu.Count -gt 0) {
        $msg += "`n`nActivitats SENSE correu (omeses): $($senseCorreu.Count)`n - " + (($senseCorreu | Select-Object -First 15) -join "`n - ")
    }
    if ($err -gt 0) { $msg += "`n`nErrors: $err`n - " + (($errDetalls | Select-Object -First 10) -join "`n - ") }
    if ($avisTec.Count -gt 0) { $msg += "`n`nATENCI" + [char]0x00D3 + ": el correu de l'Excel " + [char]0x00E9 + "s del t" + [char]0x00E8 + "cnic segons els documents (mira'ls abans d'enviar-los): $($avisTec.Count)`n - " + (($avisTec | Select-Object -First 15) -join "`n - ") }
    if ($delsDocs.Count -gt 0) { $msg += "`n`nAmb el correu tret dels documents (l'Excel no en tenia): $($delsDocs.Count)`n - " + (($delsDocs | Select-Object -First 15) -join "`n - ") }
    $msg += "`n`nRevisa'ls a Outlook (carpeta Esborranys) abans d'enviar-los."
    [System.Windows.Forms.MessageBox]::Show($msg, 'Enviar correu', 'OK', 'Information') | Out-Null
}

# ----------------------------------------------------------------------------
# EDITOR del text (WinForms) - nomes a Windows
# ----------------------------------------------------------------------------
function Invoke-ControlsCpEmailTextos {
    $textos = _LoadControlsCpEmail

    [void](Show-EditorAssumpteCos `
        -TextFinestra 'Text del correu (controls periodics)' `
        -Titol 'Text del correu' `
        -Subtitol ('Av' + [char]0x00ED + 's de control peri' + [char]0x00F2 + 'dic ' + [char]0x00B7 + ' es desa en aquest ordinador') `
        -Ajuda (_ControlsCpEmailAjuda) `
        -Assumpte ([string]$textos['assumpte']) `
        -Cos ([string]$textos['cos']) `
        -EtiquetaRestaurar 'Restaurar original' `
        -Restaurar {
            $r = [System.Windows.Forms.MessageBox]::Show('Vols recuperar el text original? (No es desa fins que premis Desar.)', 'Text del correu', 'YesNo', 'Question')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $null }
            return (_DefaultControlsCpEmail)
        } `
        -Desa {
            param($v)
            try {
                _SaveControlsCpEmail ([ordered]@{ assumpte = [string]$v['assumpte']; cos = [string]$v['cos'] })
                [System.Windows.Forms.MessageBox]::Show('Text desat en aquest ordinador.', 'Text del correu', 'OK', 'Information') | Out-Null
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Text del correu', 'OK', 'Error') | Out-Null
                return $false
            }
        })
}

# ----------------------------------------------------------------------------
# EL CORREU DE PROVA (Configuracio -> Correus de cada eina)
# ----------------------------------------------------------------------------
# Amb les dades de l'Excel de l'activitat de prova. Les dates del control
# periodic surten de l'Excel en aquesta eina; aqui es diu que son de mostra.
$Script:CorreuProves['controls'] = {
    param($gia, $cache, $cfgE)
    $act = Get-ActivitatFromCache $cache $gia
    if ($null -eq $act) { throw "L'ID GIA de prova ($gia) no es a l'Excel d'activitats." }
    $v = { param($k) try { if ($act.ContainsKey($k)) { return [string]$act[$k] } } catch { }; return '' }
    $row = [pscustomobject]@{ Id = $gia; ActPrincipal = (& $v 'ACTIVITAT'); Adreca = (& $v 'ADRECA'); RaoSocial = (& $v 'TITULAR')
                              ProperCP = '(data prevista de mostra)'; DataControlPer = '(data de mostra)' }
    $t = _LoadControlsCpEmail
    return @{
        Assumpte = (_FillControlsCpPh ([string]$t['assumpte']) $row)
        Html = (_CosAHtml (_FillControlsCpPh ([string]$t['cos']) $row))
        Destinataris = @(_CorreuDestinataris $cfgE (_CorreuEmailsDeAct $act) (Get-CorreuAutoritzats $gia))
        CcoAbans = ''
    }
}
$Script:CorreuCcoAbans['controls'] = { '' }
