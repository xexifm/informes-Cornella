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
    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("Es prepararan $($rows.Count) correus (un per activitat triada) com a ESBORRANYS a Outlook.`n`n" +
         "A qui: " + (_CorreuDestText $cfgE) + " (el primer a 'Per a' i la resta a 'CC')" + $(if ($bcc) { "`nCCO: $bcc" } else { '' }) +
         "`n`nNO s'envia res: els revisaràs i enviaràs tu des d'Outlook.`n`nVols continuar?"),
        'Enviar correu', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $textos = _LoadControlsCpEmail
    $assTpl = [string]$textos['assumpte']
    $cosTpl = [string]$textos['cos']

    # La carcassa (finestra + etiqueta + barra + Cancel.lar) es a UiComuns.ps1:
    # era identica a la de "Generar informes".
    $pg     = Show-ProgresCancel 'Enviar correu' $rows.Count
    $cancel = $pg.Cancel
    $form   = $pg.Form
    $lbl    = $pg.Label
    $bar    = $pg.Bar

    $ok = 0; $senseCorreu = New-Object System.Collections.ArrayList
    $err = 0; $errDetalls = New-Object System.Collections.ArrayList; $cancelled = $false
    $avisTec = New-Object System.Collections.ArrayList; $delsDocs = New-Object System.Collections.ArrayList
    $ses = $null
    try {
        # La mateixa sessio que Enviar correu i Recordatoris (CorreuVia.ps1): un
        # sol lloc obre l'Outlook i hi posa el remitent de Configuracio. Aqui
        # SEMPRE a Esborranys, triis la via que triis: l'usuari els revisa.
        try { $ses = Open-CorreuSessio 'outlook-esborrany' } catch {
            try { $form.Close() } catch { }
            [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Enviar correu', 'OK', 'Error') | Out-Null
            return
        }

        $done = 0
        foreach ($r in $rows) {
            if ($cancel.Flag) { $cancelled = $true; break }
            $done++
            $lbl.Text = "Preparant esborranys...  $done de $($rows.Count)`nGIA $($r.Id) - $($r.RaoSocial)"
            if ($bar.Value -lt $bar.Maximum) { $bar.Value = $done }
            [System.Windows.Forms.Application]::DoEvents()

            # El repas de contactes: el correu dels documents si l'Excel no en
            # te, i la llista dels que son del tecnic (es diu al final: son
            # esborranys, l'usuari els revisa abans d'enviar-los).
            $ctEm = Get-ContactesEmailsCompletats ([string]$r.Id) @{ titular = [string]$r.RaoEmail; representant = [string]$r.RepEmail }
            if ([string]$ctEm.AvisTecnic -ne '') { [void]$avisTec.Add("GIA $($r.Id) - $($r.RaoSocial)") }
            if (@($ctEm.Notes).Count -gt 0) { [void]$delsDocs.Add("GIA $($r.Id) - $($r.RaoSocial)") }
            $rec = _CorreuParteixToCc (_CorreuDestinataris $cfgE @{ titular = [string]$ctEm.Emails.titular; representant = [string]$ctEm.Emails.representant } (Get-CorreuAutoritzats ([string]$r.Id)))
            if (-not $rec.Ok) { [void]$senseCorreu.Add("GIA $($r.Id) - $($r.RaoSocial)"); continue }

            try {
                Send-CorreuSessio $ses $rec.To $bcc (_FillControlsCpPh $assTpl $r) (_CosAHtml (_FillControlsCpPh $cosTpl $r)) $rec.Cc
                $ok++
            } catch {
                $err++; [void]$errDetalls.Add("GIA $($r.Id): $($_.Exception.Message)")
            }
        }
    } finally {
        $cancel.Running = $false
        try { $form.Close() } catch { }
        # NO es fa Quit() de l'Outlook (podria tancar el de l'usuari); nomes s'allibera.
        Close-CorreuSessio $ses
    }

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
