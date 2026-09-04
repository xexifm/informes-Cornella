#requires -Version 5.1
<#
.SYNOPSIS
  Editor dels TEXTOS del correu que l'app mòbil envia al titular (EmailJS).

.DESCRIPTION
  Només hi ha dos camps: ASSUMPTE i COS. Al cos hi surt TOT; els requeriments
  seleccionats s'insereixen allà on posis la variable {REQUERIMENTS}. Viuen a
  docs\dades\email-textos.json, que l'app mòbil (docs\app.js) llegeix. En desar
  s'escriu el JSON i es publica la propera vegada que es faci Actualitzar.bat
  (que ja puja docs\dades\email-textos.json, pas 2b).

  Variables disponibles al cos i a l'assumpte:
    {REQUERIMENTS}  la llista de requeriments (deficiències) seleccionats
    {ID_GIA} {ADRECA} {ACTIVITAT} {TITULAR}   dades de l'activitat
    {DATA}          data d'avui (dd/MM/yyyy)
  El cos admet **negreta** (dos asteriscs) i els enllaços http(s) es tornen
  clicables sols.

  Funcions PURES (rutes, valors per defecte, càrrega/desat) testejables en
  headless; la finestra (WinForms) només a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# Ruta del fitxer de textos (docs\dades\email-textos.json a l'arrel del clone).
function _EmailTextosPath {
    $root = if ($RepoRoot) { $RepoRoot } else { (Get-Location).Path }
    return (Join-Path $root (Join-Path 'docs' (Join-Path 'dades' 'email-textos.json')))
}

# Valors PER DEFECTE (han de coincidir amb EMAIL_TEXTOS_DEFAULT de docs\app.js i
# amb el email-textos.json que es distribueix).
function _DefaultEmailTextos {
    $cos = @(
        'ID GIA: {ID_GIA}'
        'Adreça: {ADRECA}'
        'Activitat: {ACTIVITAT}'
        'Titular: {TITULAR}'
        ''
        "Aquestes són les deficiències que s'han detectat a la visita de l'activitat per part de l'Ajuntament el dia {DATA} i que s'han d'esmenar:"
        ''
        '{REQUERIMENTS}'
        ''
        '**Com presentar la documentació / Cómo presentar la documentación**'
        "Heu de presentar **tota la documentació alhora** (important: no la presenteu per parts), mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat:"
        'https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2'
        "Debe presentar **toda la documentación a la vez** (importante: no la presente por partes), mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat."
        "Indiqueu que la instància va **a l'atenció del Departament d'Activitats** / Indique que la instancia va dirigida **a la atención del Departamento de Actividades**, i feu-hi constar: ID GIA {ID_GIA}, Adreça {ADRECA}, Titular {TITULAR}."
        ''
        '________________________________________'
        ''
        "IMPORTANT: aquest és un correu automàtic i no s'admeten respostes. Aquest llistat NO és definitiu ni oficial i pot variar respecte del requeriment oficial que rebreu properament. Per a qualsevol consulta podeu adreçar-vos al Departament d'Activitats de l'Ajuntament de Cornellà de Llobregat (Carrer de l'Energia, 97) o trucar al 93 377 02 12, extensió 1227."
        ''
        "IMPORTANTE: este es un correo automático y no se admiten respuestas. Este listado NO es definitivo ni oficial y puede variar respecto del requerimiento oficial que recibirá próximamente. Para cualquier consulta puede dirigirse al Departamento de Actividades del Ayuntamiento de Cornellà de Llobregat (Calle de l'Energia, 97) o llamar al 93 377 02 12, extensión 1227."
    ) -join "`n"
    $d = [ordered]@{
        assumpte = 'GIA {ID_GIA} Requeriments'
        cos      = $cos
    }
    return ,$d
}

# Metadades dels camps (ordre + clau). Nomes assumpte i cos.
function _EmailTextosFields {
    return @(
        @{ Key = 'assumpte'; Label = 'Assumpte del correu' }
        @{ Key = 'cos';      Label = 'Cos del correu' }
    )
}

# Text d'ajuda amb les variables disponibles.
function _EmailTextosAjuda {
    return ('Variables: {REQUERIMENTS} = els requeriments  ' + [char]0x00B7 + '  {ID_GIA} {ADRECA} {ACTIVITAT} {TITULAR} {DATA}   ' + [char]0x00B7 + '   **negreta**   ' + [char]0x00B7 + '   els enllacos http es fan clicables')
}

# Llegeix el JSON (si hi es) fusionat sobre els valors per defecte. Retorna un
# ordered hashtable amb TOTES les claus.
function _LoadEmailTextos {
    $def = _DefaultEmailTextos
    $path = _EmailTextosPath
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

# Escriu el JSON (UTF-8 sense BOM). $obj = ordered hashtable amb les claus.
function _SaveEmailTextos($obj) {
    $path = _EmailTextosPath
    Write-JsonFile $path $obj 5
}

# ----------------------------------------------------------------------------
# INTERFICIE (WinForms) - nomes a Windows
# ----------------------------------------------------------------------------
function Invoke-EmailTextos {
    $textos = _LoadEmailTextos

    $form = _NewForm
    $form.Text = 'Textos del correu (mobil)'
    $form.ClientSize = New-Object System.Drawing.Size(760, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 470)

    # Assumpte.
    $lblA = New-Object System.Windows.Forms.Label
    $lblA.Text = 'Assumpte del correu:'
    $lblA.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblA.Location = New-Object System.Drawing.Point(16, 70)
    $lblA.AutoSize = $true
    [void]$form.Controls.Add($lblA)

    $tbA = New-Object System.Windows.Forms.TextBox
    $tbA.Location = New-Object System.Drawing.Point(16, 92)
    $tbA.Size = New-Object System.Drawing.Size(728, 24)
    $tbA.Anchor = 'Top, Left, Right'
    $tbA.Text = [string]$textos['assumpte']
    [void]$form.Controls.Add($tbA)

    # Cos.
    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Cos del correu:'
    $lblC.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblC.Location = New-Object System.Drawing.Point(16, 126)
    $lblC.AutoSize = $true
    [void]$form.Controls.Add($lblC)

    $lblH = New-Object System.Windows.Forms.Label
    $lblH.Text = _EmailTextosAjuda
    $lblH.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
    $lblH.ForeColor = [System.Drawing.Color]::FromArgb(120, 128, 138)
    $lblH.Location = New-Object System.Drawing.Point(16, 146)
    $lblH.AutoSize = $false
    $lblH.Size = New-Object System.Drawing.Size(728, 16)
    $lblH.Anchor = 'Top, Left, Right'
    [void]$form.Controls.Add($lblH)

    $tbC = New-Object System.Windows.Forms.TextBox
    $tbC.Location = New-Object System.Drawing.Point(16, 166)
    $tbC.Size = New-Object System.Drawing.Size(728, 396)
    $tbC.Anchor = 'Top, Bottom, Left, Right'
    $tbC.Multiline = $true
    $tbC.ScrollBars = 'Vertical'
    $tbC.WordWrap = $true
    $tbC.AcceptsReturn = $true
    $tbC.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    # El TextBox multilinia de WinForms NOMES mostra els salts com a CRLF; el cos
    # es guarda amb LF (\n), aixi que el normalitzem a CRLF per veure'l bé.
    $tbC.Text = [string]$textos['cos'] -replace "`r?`n", "`r`n"
    [void]$form.Controls.Add($tbC)

    # Barra inferior.
    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Enrere'
    $btnBack.Location = New-Object System.Drawing.Point(16, 578)
    $btnBack.Size = New-Object System.Drawing.Size(110, 30)
    $btnBack.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnBack
    [void]$form.Controls.Add($btnBack)

    $btnDefault = New-Object System.Windows.Forms.Button
    $btnDefault.Text = 'Restaurar original'
    $btnDefault.Location = New-Object System.Drawing.Point(136, 578)
    $btnDefault.Size = New-Object System.Drawing.Size(150, 30)
    $btnDefault.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $btnDefault
    [void]$form.Controls.Add($btnDefault)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Desar'
    $btnSave.Location = New-Object System.Drawing.Point(624, 578)
    $btnSave.Size = New-Object System.Drawing.Size(120, 30)
    $btnSave.Anchor = 'Bottom, Right'
    _StylePrimaryButton $btnSave
    [void]$form.Controls.Add($btnSave)

    [void](_AddBrandHeader $form 'Textos del correu' ('El que envia l''app m' + [char]0x00F2 + 'bil al titular ' + [char]0x00B7 + ' es publica amb Actualitzar'))

    $btnBack.add_Click({ $form.Close() }.GetNewClosure())

    $btnDefault.add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show('Vols recuperar els textos originals? (No es desa fins que premis Desar.)', 'Textos del correu', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $def = _DefaultEmailTextos
        $tbA.Text = [string]$def['assumpte']
        $tbC.Text = [string]$def['cos'] -replace "`r?`n", "`r`n"
    }.GetNewClosure())

    $btnSave.add_Click({
        if ([string]$tbC.Text -notmatch '\{REQUERIMENTS\}') {
            $r = [System.Windows.Forms.MessageBox]::Show("El cos no conte la variable {REQUERIMENTS}: els requeriments NO sortiran al correu.`n`nVols desar igualment?", 'Textos del correu', 'YesNo', 'Warning')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        # Tornem a LF (\n) per al JSON (app.js parteix el cos per \n).
        $out = [ordered]@{ assumpte = [string]$tbA.Text; cos = ([string]$tbC.Text -replace "`r`n", "`n") }
        try {
            _SaveEmailTextos $out
            [System.Windows.Forms.MessageBox]::Show("Textos desats.`n`nEs publicaran al mobil la propera vegada que facis Actualitzar.", 'Textos del correu', 'OK', 'Information') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'han pogut desar:`n$($_.Exception.Message)", 'Textos del correu', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    [void]$form.ShowDialog()
    $form.Dispose()
}
